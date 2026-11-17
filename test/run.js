// Tests for startup cost, sign-in, reading the calendar and inserting events.
// Run with: node test/run.js   (TZ is pinned so day and DST checks are stable)
process.env.TZ = "Europe/Copenhagen";

const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");

const stubPath = path.join(__dirname, "obsidian-stub.js");
const codemirrorStubPath = path.join(__dirname, "codemirror-stub.js");
const realResolve = Module._resolveFilename;
Module._resolveFilename = function (request, ...rest) {
  if (request === "obsidian") return stubPath;
  if (request === "@codemirror/view" || request === "@codemirror/state") return codemirrorStubPath;
  return realResolve.call(this, request, ...rest);
};
const mainPath = path.join(__dirname, "..", "src", "plugin.js");
const source =
  fs.readFileSync(mainPath, "utf8") +
  "\nmodule.exports.__internals = { MicrosoftAuth, fetchDayEvents, eventView, dayEvents, dayBounds, addDays, placeOnOwnLine, insertOnOwnLine, describeAuthError, localToday, lineKeys, eventKeys, noteLineKeys, buildEventIndex, eventKeyForLine };\n";
const mod = new Module(mainPath, module);
mod.filename = mainPath;
mod.paths = Module._nodeModulePaths(path.dirname(mainPath));
mod._compile(source, mainPath);
const CalendarPlugin = mod.exports;
const {
  fetchDayEvents,
  eventView,
  dayEvents,
  dayBounds,
  addDays,
  placeOnOwnLine,
  insertOnOwnLine,
  describeAuthError,
  localToday,
  lineKeys,
  eventKeys,
  noteLineKeys,
  buildEventIndex,
  eventKeyForLine,
} = CalendarPlugin.__internals;

global.window = { setTimeout: (fn) => setImmediate(fn), clearTimeout: (id) => id && clearImmediate(id), setInterval, clearInterval };
global.atob = global.atob || ((s) => Buffer.from(s, "base64").toString("binary"));

const idToken = `h.${Buffer.from(JSON.stringify({ preferred_username: "user@example.com" })).toString("base64url")}.s`;

const tests = [];
const test = (name, fn) => tests.push([name, fn]);

function reset() {
  global.__requests = [];
  global.__secretReads = 0;
  global.__leaves = [];
  global.__activeFile = null;
  global.__layoutReady = null;
  global.__notices = [];
}

async function signedInPlugin() {
  reset();
  global.__data = { clientId: "client", tenant: "tenant" };
  const plugin = new CalendarPlugin();
  await plugin.onload();
  global.__handler = () => ({ body: { access_token: "AT1", refresh_token: "RT1", expires_in: 3600, id_token: idToken } });
  await plugin.auth.completeDeviceFlow({ device_code: "d", interval: 1 }, { cancelled: false });
  global.__requests = [];
  return plugin;
}

// A minimal DOM element: enough for the panel to build itself, and to read
// back what it shows.
function el(tag = "div", opts = {}) {
  const node = {
    tag,
    text: opts.text || "",
    cls: new Set(String(opts.cls || "").split(" ").filter(Boolean)),
    children: [],
    listeners: {},
    isConnected: true,
    isShown: () => node.shown !== false,
    offsetWidth: 0,
    scrollIntoView: (opts) => (node.scrolledTo = opts.block),
    empty() {
      node.children = [];
    },
    addClass: (c) => node.cls.add(c),
    removeClass: (c) => node.cls.delete(c),
    toggleClass: (c, on) => (on ? node.cls.add(c) : node.cls.delete(c)),
    setAttr() {},
    setText: (t) => (node.text = t),
    addEventListener: (type, fn) => (node.listeners[type] = fn),
    createDiv: (o) => node.createEl("div", o),
    createSpan: (o) => node.createEl("span", o),
    createEl(t, o) {
      const child = el(t, o || {});
      node.children.push(child);
      return child;
    },
    find(cls) {
      const out = [];
      const walk = (n) => {
        if (n.cls.has(cls)) out.push(n);
        n.children.forEach(walk);
      };
      walk(node);
      return out;
    },
  };
  return node;
}

async function openView(plugin) {
  const view = global.__viewFactory({});
  view.app = plugin.app;
  view.contentEl = el();
  global.__leaves = [{ view }];
  await view.onOpen();
  return view;
}

// Graph's shape of an event, in UTC.
function graphEvent(id, subject, start, end, extra = {}) {
  return {
    id,
    subject,
    start: { dateTime: `${start}.0000000`, timeZone: "UTC" },
    end: { dateTime: `${end}.0000000`, timeZone: "UTC" },
    isAllDay: false,
    responseStatus: { response: "accepted" },
    ...extra,
  };
}

// ── startup cost ──

test("loading the plugin makes no request and reads no secret", async () => {
  reset();
  global.__data = { clientId: "client", tenant: "tenant" };
  global.__handler = () => assert.fail("no request expected");
  const plugin = new CalendarPlugin();
  await plugin.onload();
  assert.strictEqual(global.__requests.length, 0);
  assert.strictEqual(global.__secretReads, 0, "secret storage is read on first use, not on load");
  plugin.tick();
  assert.strictEqual(global.__secretReads, 0, "a tick with no panel does nothing");
});

test("opening the panel waits for neither the network nor the layout", async () => {
  const plugin = await signedInPlugin();
  let finish;
  global.__handler = () => new Promise((resolve) => (finish = resolve));
  const view = global.__viewFactory({});
  view.app = plugin.app;
  view.contentEl = el();
  const opened = await Promise.race([view.onOpen().then(() => "opened"), new Promise((r) => setTimeout(() => r("blocked"), 50))]);
  assert.strictEqual(opened, "opened");
  assert.strictEqual(global.__requests.length, 0, "nothing fetched before the layout is ready");
  assert.ok(global.__layoutReady, "fetch deferred until the layout is ready");
  global.__layoutReady();
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(global.__requests.length, 1, "then fetched once, in the background");
  assert.ok(view.loading);
  finish({ body: { value: [] } });
  await new Promise((r) => setTimeout(r, 10));
  assert.ok(!view.loading);
});

test("a panel that is not on screen fetches nothing", async () => {
  const plugin = await signedInPlugin();
  global.__handler = () => ({ body: { value: [] } });
  const view = await openView(plugin);
  view.contentEl.shown = false;
  global.__layoutReady();
  await new Promise((r) => setTimeout(r, 10));
  plugin.tick();
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(global.__requests.length, 0);
  view.contentEl.shown = true;
  plugin.tick();
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(global.__requests.length, 1, "it catches up once shown");
  plugin.tick();
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(global.__requests.length, 1, "a fresh day is not fetched again");
  // Opening another daily note while hidden fetches nothing until shown.
  view.contentEl.shown = false;
  view.followNote("2026-12-24");
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(view.day, "2026-12-24");
  assert.strictEqual(global.__requests.length, 1);
  view.contentEl.shown = true;
  plugin.tick();
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(global.__requests.length, 2);
});

test("only a forced refresh is queued behind a running one", async () => {
  const plugin = await signedInPlugin();
  const finishes = [];
  global.__handler = () => new Promise((resolve) => finishes.push(resolve));
  const view = await openView(plugin);
  global.__layoutReady();
  await new Promise((r) => setTimeout(r, 10));
  view.refresh(false);
  finishes.shift()({ body: { value: [] } });
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(global.__requests.length, 1, "a non-forced call mid-load adds no fetch");
  const again = view.refresh(true);
  view.refresh(true);
  await new Promise((r) => setTimeout(r, 10));
  finishes.shift()({ body: { value: [] } });
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(global.__requests.length, 3, "a forced call mid-load runs once more");
  finishes.shift()({ body: { value: [] } });
  await again;
});

// ── sign-in ──

test("device code flow asks for Calendars.Read and waits through pending and slow_down", async () => {
  reset();
  global.__data = { clientId: "client", tenant: "tenant" };
  const plugin = new CalendarPlugin();
  await plugin.onload();
  let polls = 0;
  global.__handler = (req) => {
    if (req.url.endsWith("/devicecode")) return { body: { user_code: "ABC", device_code: "dev", verification_uri: "https://microsoft.com/devicelogin", interval: 1 } };
    polls++;
    if (polls === 1) return { status: 400, body: { error: "authorization_pending" } };
    if (polls === 2) return { status: 400, body: { error: "slow_down" } };
    return { body: { access_token: "AT", refresh_token: "RT", expires_in: 3600, id_token: idToken } };
  };
  const flow = await plugin.auth.startDeviceFlow();
  assert.ok(global.__requests[0].url.startsWith("https://login.microsoftonline.com/tenant/oauth2/v2.0/devicecode"));
  assert.ok(global.__requests[0].body.includes("scope=Calendars.Read+offline_access+openid+profile"));
  await plugin.auth.completeDeviceFlow(flow, { cancelled: false });
  assert.strictEqual(polls, 3);
  assert.ok(plugin.auth.signedIn);
  assert.strictEqual(plugin.auth.account, "user@example.com");
});

test("tokens live in secret storage under this plugin's own ID, never in plugin data", async () => {
  const plugin = await signedInPlugin();
  await plugin.saveSettings();
  assert.ok(!JSON.stringify(global.__saved).includes("RT1"));
  assert.ok(plugin.auth.secretId.startsWith("outlook-calendar-panel-session-"));
  assert.ok(plugin.app.secrets[plugin.auth.secretId].includes("RT1"));
});

test("concurrent callers share one refresh and keep the rotated token", async () => {
  const plugin = await signedInPlugin();
  plugin.auth.session.expires_at = 0;
  let refreshes = 0;
  global.__handler = () => {
    refreshes++;
    return { body: { access_token: "AT2", refresh_token: "RT2", expires_in: 3600 } };
  };
  const [a, b] = await Promise.all([plugin.auth.accessToken(), plugin.auth.accessToken()]);
  assert.strictEqual(refreshes, 1);
  assert.strictEqual(a, "AT2");
  assert.strictEqual(b, "AT2");
  assert.strictEqual(plugin.auth.session.refresh_token, "RT2");
});

test("an expired refresh token signs out and the panel says so", async () => {
  const plugin = await signedInPlugin();
  global.__handler = () => ({ body: { value: [] } });
  await plugin.loadDay("2026-11-17");
  assert.ok(plugin.app.local["outlook-calendar-panel-days"]);
  plugin.auth.session.expires_at = 0;
  global.__handler = () => ({ status: 400, body: { error: "invalid_grant", error_description: "AADSTS70008: expired.\r\nTrace ID: x" } });
  const view = await openView(plugin);
  await view.refresh(true);
  assert.strictEqual(plugin.auth.signedIn, false);
  assert.match(view.error, /Sign in again\. \(AADSTS70008: expired\.\)/);
  assert.strictEqual(plugin.app.local["outlook-calendar-panel-days"], undefined, "cached events go too");
});

test("auth errors are trimmed and get a setup hint", () => {
  assert.match(describeAuthError({ error_description: "AADSTS7000218: body must contain client_secret.\r\nTrace ID: 1" }), /Allow public client flows/);
  assert.strictEqual(describeAuthError(null), "No response from Microsoft.");
});

// ── reading the calendar ──

test("reads a local day: window, UTC header, paging, 401 retry and mapping", async () => {
  const plugin = await signedInPlugin();
  let calls = 0;
  global.__handler = (req) => {
    if (req.url.endsWith("/token")) return { body: { access_token: "AT3", refresh_token: "RT3", expires_in: 3600 } };
    calls++;
    if (calls === 1) return { status: 401, body: { error: { message: "expired" } } };
    if (!req.url.includes("page=2")) {
      return {
        body: {
          value: [
            graphEvent("b", "Lunch", "2026-11-17T11:00:00", "2026-11-17T12:00:00", {
              location: { displayName: "Canteen" },
              onlineMeeting: { joinUrl: "https://teams.microsoft.com/l/meetup-join/x" },
              webLink: "https://outlook.office365.com/owa/?itemid=b",
            }),
            graphEvent("x", "Gone", "2026-11-17T08:00:00", "2026-11-17T09:00:00", { isCancelled: true }),
          ],
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/calendarView?page=2",
        },
      };
    }
    return { body: { value: [graphEvent("a", "Standup", "2026-11-17T08:00:00", "2026-11-17T08:15:00", { webLink: "javascript:alert(1)" })] } };
  };
  const { events, failures } = await fetchDayEvents(plugin.auth, "2026-11-17");
  assert.deepStrictEqual(failures, []);
  const first = global.__requests.find((r) => r.url.includes("calendarView"));
  assert.ok(first.url.includes("startDateTime=2026-11-16T23%3A00%3A00.000Z"));
  assert.ok(first.url.includes("endDateTime=2026-11-17T23%3A00%3A00.000Z"));
  assert.ok(first.url.includes("$select=") && first.url.includes("$top=100"));
  assert.strictEqual(first.headers.Prefer, 'outlook.timezone="UTC"');
  assert.deepStrictEqual(events.map((e) => e.id), ["x", "a", "b"], "sorted by start");
  const lunch = events.find((e) => e.id === "b");
  assert.strictEqual(lunch.location, "Canteen");
  assert.strictEqual(lunch.joinUrl, "https://teams.microsoft.com/l/meetup-join/x");
  assert.strictEqual(events.find((e) => e.id === "a").webLink, "", "only https links are kept");
  assert.deepStrictEqual(plugin.visibleEvents(events).map((e) => e.id), ["a", "b"], "cancelled events are hidden");
  assert.strictEqual(plugin.timeLabel(lunch, "2026-11-17"), "12:00–13:00", "shown in local time");
});

test("chosen calendars are read one by one, and one that fails does not hide the others", async () => {
  const plugin = await signedInPlugin();
  global.__handler = (req) => {
    if (req.url.includes("/calendars/work/")) return { body: { value: [graphEvent("w", "Standup", "2026-11-17T08:00:00", "2026-11-17T08:15:00")] } };
    if (req.url.includes("/calendars/home/")) return { body: { value: [graphEvent("h", "Dentist", "2026-11-17T07:00:00", "2026-11-17T07:30:00")] } };
    return { status: 403, body: { error: { message: "Access is denied." } } };
  };
  const calendars = [
    { id: "work", name: "Work", color: "#4f9fe6" },
    { id: "home", name: "Home", color: "" },
    { id: "shared", name: "Ann's calendar", color: "" },
  ];
  await plugin.setCalendars(calendars);
  const events = await plugin.loadDay("2026-11-17");
  assert.deepStrictEqual(events.map((e) => [e.id, e.calendar]), [["h", "Home"], ["w", "Work"]], "merged and sorted by start");
  assert.strictEqual(events[1].color, "#4f9fe6");
  const entry = plugin.dayCache().get("2026-11-17");
  assert.deepStrictEqual(entry.failures.map((f) => f.name), ["Ann's calendar"]);
  assert.match(entry.failures[0].message, /Access is denied/);
  plugin.settings.template = "- {{time}} {{title}} ({{calendar}})";
  assert.strictEqual(plugin.renderEvent(events[0], "2026-11-17"), "- 08:00–08:30 Dentist (Home)");
  // All of them failing is an error.
  global.__handler = () => ({ status: 403, body: { error: { message: "Access is denied." } } });
  await assert.rejects(plugin.loadDay("2026-11-18"), /Access is denied/);
});

test("choosing other calendars drops the days fetched for the old choice", async () => {
  const plugin = await signedInPlugin();
  global.__handler = () => ({ body: { value: [] } });
  await plugin.loadDay("2026-11-17");
  assert.strictEqual(plugin.dayCache().size, 1);
  await plugin.setCalendars([{ id: "work", name: "Work" }]);
  assert.strictEqual(plugin.dayCache().size, 0);
  await plugin.loadDay("2026-11-17");
  // A start with another choice saved ignores the stored days.
  plugin.days = null;
  plugin.settings.calendars = [];
  assert.strictEqual(plugin.dayCache().size, 0);
  // Calendars chosen with another account are not used.
  plugin.settings.calendars = [{ id: "work", name: "Work" }];
  plugin.settings.calendarsAccount = "someone@else.com";
  assert.deepStrictEqual(plugin.calendars(), []);
});

test("the calendar list comes back default first, with owners and colours", async () => {
  const plugin = await signedInPlugin();
  global.__handler = () => ({
    body: {
      value: [
        { id: "b", name: "Ann's calendar", color: "lightGreen", hexColor: "", owner: { name: "Ann", address: "ann@example.com" } },
        { id: "a", name: "Calendar", color: "auto", hexColor: "#123abc", isDefaultCalendar: true, owner: { name: "Me", address: "USER@example.com" } },
      ],
    },
  });
  const list = await plugin.listCalendars();
  assert.deepStrictEqual(list.map((c) => [c.id, c.color, c.owner, c.isDefault]), [
    ["a", "#123abc", "", true],
    ["b", "#5cb85c", "Shared by Ann", false],
  ]);
  assert.ok(global.__requests[0].url.includes("/me/calendars?"));
});

test("a day that leaves summer time is 25 hours long", async () => {
  const { start, end } = dayBounds("2026-10-25");
  assert.strictEqual(new Date(start).toISOString(), "2026-10-24T22:00:00.000Z");
  assert.strictEqual(new Date(end).toISOString(), "2026-10-25T23:00:00.000Z");
  assert.strictEqual(addDays("2026-03-28", 1), "2026-03-29");
  assert.strictEqual(addDays("2026-03-29", 1), "2026-03-30");
});

test("all-day events land on their own date, whichever way Graph sends them", () => {
  const bare = eventView({ id: "1", subject: "Holiday", isAllDay: true, start: { dateTime: "2026-11-17T00:00:00.0000000", timeZone: "UTC" }, end: { dateTime: "2026-11-18T00:00:00.0000000", timeZone: "UTC" } });
  const shifted = eventView({ id: "2", subject: "Offsite", isAllDay: true, start: { dateTime: "2026-11-16T23:00:00.0000000", timeZone: "UTC" }, end: { dateTime: "2026-11-18T23:00:00.0000000", timeZone: "UTC" } });
  assert.strictEqual(bare.startDate, "2026-11-17");
  assert.strictEqual(bare.endDate, "2026-11-18");
  assert.strictEqual(shifted.startDate, "2026-11-17");
  assert.strictEqual(shifted.endDate, "2026-11-19");
  // A neighbouring day's all-day event caught by the UTC window is dropped.
  const before = eventView({ id: "3", subject: "Yesterday", isAllDay: true, start: { dateTime: "2026-11-16T00:00:00.0000000" }, end: { dateTime: "2026-11-17T00:00:00.0000000" } });
  assert.deepStrictEqual(dayEvents([bare, shifted, before], "2026-11-17").map((e) => e.id), ["1", "2"]);
  assert.deepStrictEqual(dayEvents([bare, shifted, before], "2026-11-18").map((e) => e.id), ["2"]);
});

test("an event across midnight is clipped to the day shown", async () => {
  const plugin = await signedInPlugin();
  const night = eventView(graphEvent("n", "Night shift", "2026-11-17T21:00:00", "2026-11-18T05:00:00"));
  assert.strictEqual(plugin.timeLabel(night, "2026-11-17"), "From 22:00");
  assert.strictEqual(plugin.timeLabel(night, "2026-11-18"), "Until 06:00");
  const trip = eventView(graphEvent("t", "Trip", "2026-11-16T10:00:00", "2026-11-19T10:00:00"));
  assert.strictEqual(plugin.timeLabel(trip, "2026-11-17"), "All day");
  assert.deepStrictEqual(dayEvents([night], "2026-11-19"), []);
});

test("events fetched for a session that ended are thrown away", async () => {
  const plugin = await signedInPlugin();
  let finish;
  global.__handler = () => new Promise((resolve) => (finish = resolve));
  const loading = plugin.loadDay("2026-11-17").catch((e) => e);
  await new Promise((r) => setTimeout(r, 10));
  plugin.signOut();
  finish({ body: { value: [graphEvent("a", "Secret", "2026-11-17T08:00:00", "2026-11-17T09:00:00")] } });
  assert.ok((await loading) instanceof Error);
  assert.strictEqual(plugin.dayCache().size, 0);
});

test("days are cached per account, capped, and cleared on sign-out", async () => {
  const plugin = await signedInPlugin();
  global.__handler = () => ({ body: { value: [graphEvent("a", "Standup", "2026-11-17T08:00:00", "2026-11-17T08:15:00")] } });
  for (let i = 0; i < 10; i++) await plugin.loadDay(addDays("2026-11-10", i));
  const stored = plugin.app.local["outlook-calendar-panel-days"];
  assert.strictEqual(stored.account, "user@example.com");
  assert.strictEqual(Object.keys(stored.days).length, 7);
  // A new start reads them back at once, for the same account only.
  plugin.days = null;
  assert.strictEqual(plugin.dayCache().size, 7);
  plugin.days = null;
  plugin.auth.session.account = "someone@else.com";
  assert.strictEqual(plugin.dayCache().size, 0);
  plugin.auth.session.account = "user@example.com";
  plugin.signOut();
  assert.strictEqual(plugin.app.local["outlook-calendar-panel-days"], undefined);
});

// ── the panel ──

test("the panel shows the day's events, follows the daily note and goes back to today", async () => {
  const plugin = await signedInPlugin();
  global.__handler = (req) => {
    const day = decodeURIComponent(req.url.match(/endDateTime=([^&]+)/)[1]);
    return day.startsWith("2026-11-20")
      ? { body: { value: [graphEvent("d", "Review", "2026-11-20T09:00:00", "2026-11-20T10:00:00", { responseStatus: { response: "declined" } }), graphEvent("e", "Retro", "2026-11-20T13:00:00", "2026-11-20T14:00:00")] } }
      : { body: { value: [] } };
  };
  global.__activeFile = { basename: "2026-11-20" };
  const view = await openView(plugin);
  assert.strictEqual(view.day, "2026-11-20", "starts on the open daily note's day");
  global.__layoutReady();
  await new Promise((r) => setTimeout(r, 10));
  const subjects = () => view.listEl.find("oc-subject").map((n) => n.text);
  assert.deepStrictEqual(subjects(), ["Retro"], "declined events are hidden");
  plugin.settings.showDeclined = true;
  view.render();
  assert.deepStrictEqual(subjects(), ["Review", "Retro"]);

  view.followNote(null);
  assert.strictEqual(view.day, localToday(), "leaving the daily note goes back to today");
  view.followNote("2026-11-20");
  view.setDay("2026-12-01", false);
  view.followNote("2026-11-20");
  assert.strictEqual(view.day, "2026-12-01", "clicking in the panel keeps a day the user picked");
  view.followNote(null);
  assert.strictEqual(view.day, "2026-12-01", "and so does leaving the daily note");
  view.followNote("2026-11-20");
  assert.strictEqual(view.day, "2026-11-20", "opening a daily note shows its day");
});

test("an error for a day left behind is not shown on the day shown", async () => {
  const plugin = await signedInPlugin();
  let fail;
  global.__handler = () => new Promise((resolve) => (fail = resolve));
  const view = await openView(plugin);
  global.__layoutReady();
  await new Promise((r) => setTimeout(r, 10));
  plugin.dayCache().set("2026-12-01", { events: [], fetchedAt: Date.now() });
  view.setDay("2026-12-01", false);
  fail({ status: 500, body: { error: { message: "boom" } } });
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(view.error, null);
});

test("midnight moves a panel showing today on, but not one following a daily note", async () => {
  const plugin = await signedInPlugin();
  global.__handler = () => ({ body: { value: [] } });
  const view = await openView(plugin);
  global.__layoutReady();
  await new Promise((r) => setTimeout(r, 10));
  const today = localToday();
  view.day = view.lastToday = addDays(today, -1);
  view.tick();
  assert.strictEqual(view.day, today);
  view.followNote(addDays(today, -1));
  view.lastToday = addDays(today, -1);
  view.tick();
  assert.strictEqual(view.day, addDays(today, -1), "stays on the open daily note's day");
});

test("an empty today is not rebuilt every minute", async () => {
  const plugin = await signedInPlugin();
  global.__handler = () => ({ body: { value: [] } });
  const view = await openView(plugin);
  global.__layoutReady();
  await new Promise((r) => setTimeout(r, 10));
  let renders = 0;
  const render = view.render.bind(view);
  view.render = () => (renders++, render());
  view.tick();
  view.tick();
  assert.strictEqual(renders, 0);
});

test("a nested daily note format matches the path, not a bare day number", async () => {
  const plugin = await signedInPlugin();
  plugin.settings.dailyNoteFormat = "YYYY/MM/DD";
  assert.strictEqual(plugin.dailyNoteDate({ basename: "17", path: "journals/2025/03/17.md" }), "2025-03-17");
  assert.strictEqual(plugin.dailyNoteDate({ basename: "17", path: "Projects/17.md" }), null);
});

test("+ and the agenda command insert events with the template", async () => {
  const plugin = await signedInPlugin();
  let text = ["# 2026-11-17", "- Notes"];
  const editor = {
    getCursor: () => ({ line: 1, ch: 0 }),
    setCursor() {},
    getLine: (n) => text[n],
    transaction({ changes }) {
      const c = changes[0];
      const lines = text.slice();
      lines[c.from.line] = lines[c.from.line].slice(0, c.from.ch) + c.text + lines[c.from.line].slice(c.to.ch);
      text = lines.join("\n").split("\n");
    },
    focus() {},
  };
  const { MarkdownView } = require("./obsidian-stub.js");
  const note = Object.assign(new MarkdownView(), { editor, leaf: {}, file: { basename: "2026-11-17" }, getMode: () => "source" });
  plugin.app.workspace.getActiveViewOfType = () => note;

  const lunch = eventView(graphEvent("b", "Lunch", "2026-11-17T11:00:00", "2026-11-17T12:00:00", { location: { displayName: "Canteen" } }));
  await plugin.addToNote(lunch, "2026-11-17");
  assert.deepStrictEqual(text, ["# 2026-11-17", "- Notes", "- 12:00–13:00 Lunch"]);
  plugin.settings.template = "- [ ] {{start}} {{title}} @ {{location}} {{join}}";
  assert.strictEqual(plugin.renderEvent(lunch, "2026-11-17"), "- [ ] 12:00 Lunch @ Canteen", "empty placeholders leave no spaces");

  plugin.settings.template = "- {{time}} {{title}}";
  global.__handler = () => ({ body: { value: [graphEvent("a", "Standup", "2026-11-17T08:00:00", "2026-11-17T08:15:00"), graphEvent("b", "Lunch", "2026-11-17T11:00:00", "2026-11-17T12:00:00")] } });
  text = ["# 2026-11-17", ""];
  editor.getCursor = () => ({ line: 1, ch: 0 });
  await plugin.insertAgenda();
  assert.deepStrictEqual(text, ["# 2026-11-17", "- 09:00–09:15 Standup", "- 12:00–13:00 Lunch"], "the daily note's day, on the blank line");
});

test("an event dropped inside a line moves to a line of its own", () => {
  const line = "- 09:00–10:00 Standup";
  let text = `Call Ann ${line}about it`;
  let selection = { anchor: { line: 0, ch: 9 }, head: { line: 0, ch: 9 + line.length } };
  const offset = (pos) => text.split("\n").slice(0, pos.line).reduce((n, l) => n + l.length + 1, 0) + pos.ch;
  const editor = {
    listSelections: () => [selection],
    getLine: (n) => text.split("\n")[n],
    getRange: (a, b) => text.slice(offset(a), offset(b)),
    transaction({ changes }) {
      for (const c of changes) text = text.slice(0, offset(c.from)) + c.text + text.slice(offset(c.to));
    },
    setSelection: (anchor, head) => (selection = { anchor, head }),
  };
  placeOnOwnLine(editor, line);
  assert.deepStrictEqual(text.split("\n"), ["Call Ann about it", line]);
});

test("several lines insert below an indented line with its indentation", () => {
  let text = ["  - Today"];
  let cursor;
  const editor = {
    getCursor: () => ({ line: 0, ch: 0 }),
    setCursor: (c) => (cursor = c),
    getLine: (n) => text[n],
    transaction({ changes }) {
      const c = changes[0];
      text = (text[0].slice(0, c.from.ch) + c.text + text[0].slice(c.to.ch)).split("\n");
    },
  };
  insertOnOwnLine(editor, "- a\n- b");
  assert.deepStrictEqual(text, ["  - Today", "  - a", "  - b"]);
  assert.deepStrictEqual(cursor, { line: 2, ch: 5 });
});

// ── links from note lines, and the events a note holds ──

const tick = () => new Promise((resolve) => setImmediate(resolve));

test("a note line matches an event by its inserted line, or by its time and title", async () => {
  const plugin = await signedInPlugin();
  const standup = eventView(graphEvent("s", "Standup", "2026-11-17T08:00:00", "2026-11-17T08:15:00"));
  const keys = plugin.eventKeys(standup, "2026-11-17");
  const days = new Map([["2026-11-17", { events: [standup], fetchedAt: 1 }]]);
  const index = buildEventIndex(days, plugin);
  const match = (line) => eventKeyForLine(index, line);
  assert.strictEqual(match("  - 09:00–09:15 Standup  "), "=- 09:00–09:15 Standup", "the inserted line, trimmed");
  assert.strictEqual(match("* [ ] 09:00-09:15   STANDUP #meeting [Join](https://teams/x)"), keys[1], "case, dash, spaces and a tail of noise");
  assert.strictEqual(match("1. 09:00 standup ✅ 2026-11-17"), keys[2], "the start time alone");
  assert.strictEqual(match("- Standup"), null, "a title alone is not enough");
  assert.strictEqual(match("09:00 Standup"), null, "not a list item");
  assert.strictEqual(match(`- 09:00 Standup ${"x".repeat(2000)}`), null, "long lines are skipped");
  // Runs of digits that could split into dates several ways stay fast.
  for (const unit of ["12-12-12", "12.12.12", "1/1/1111/", "#a1/1/11"]) {
    const started = Date.now();
    match(`- 09:00 Standup ${unit.repeat(Math.floor(900 / unit.length))}!`);
    assert.ok(Date.now() - started < 200, `no runaway backtracking on ${unit}`);
  }
  assert.deepStrictEqual(lineKeys("   "), []);
  assert.deepStrictEqual(eventKeys("", "T", "", ""), []);
  assert.strictEqual(eventKeys("- 09:00 T\n  Room 1", "T", "", "")[0], "=- 09:00 T", "a template of several lines matches by its first");
  // The panel's "in the note" marks use the same keys.
  const note = noteLineKeys("# Day\n- 9:00 nothing\n* [x] 09:00 Standup");
  assert.ok(keys.some((key) => note.has(key)));
  // A recurring event: one key, one candidate per day.
  const next = eventView(graphEvent("s2", "Standup", "2026-11-18T08:00:00", "2026-11-18T08:15:00"));
  days.set("2026-11-18", { events: [next], fetchedAt: 1 });
  assert.deepStrictEqual(buildEventIndex(days, plugin).get(keys[1]), [{ id: "s", day: "2026-11-17" }, { id: "s2", day: "2026-11-18" }]);
  assert.strictEqual(buildEventIndex(new Map(), plugin), null, "no events, no index");
});

// A fake CodeMirror view over lines of text, showing the given line ranges.
function fakeEditorView(lines, visible, file = null) {
  const starts = [];
  let pos = 0;
  for (const text of lines) {
    starts.push(pos);
    pos += text.length + 1;
  }
  const line = (n) => ({ number: n, from: starts[n - 1], to: starts[n - 1] + lines[n - 1].length, text: lines[n - 1] });
  const doc = { lines: lines.length, line, lineAt: (p) => line(starts.filter((s) => s <= p).length) };
  return {
    state: { doc, field: () => ({ file }) },
    visibleRanges: visible.map(([a, b]) => ({ from: line(a).from, to: line(b).to })),
    dispatched: [],
    dispatch(spec) {
      this.dispatched.push(spec);
    },
  };
}

test("an editor opened at startup reads neither the sign-in nor the cache", async () => {
  reset();
  global.__data = { clientId: "client", tenant: "tenant" };
  global.__handler = () => assert.fail("no request expected");
  const plugin = new CalendarPlugin();
  await plugin.onload();
  const { cls } = global.__editorExtension;
  const links = new cls(fakeEditorView(["- 09:00 Standup"], [[1, 1]]));
  assert.deepStrictEqual(links.decorations, []);
  assert.strictEqual(global.__secretReads, 0);
  assert.strictEqual(global.__requests.length, 0);
});

test("the editor links cover visible event lines only, and redraw when events change", async () => {
  const plugin = await signedInPlugin();
  const { cls } = global.__editorExtension;
  global.__handler = () => ({ body: { value: [graphEvent("a", "Standup", "2026-11-17T08:00:00", "2026-11-17T08:15:00"), graphEvent("b", "Lunch", "2026-11-17T11:00:00", "2026-11-17T12:00:00")] } });
  await plugin.loadDay("2026-11-17");
  const lines = ["- 09:00–09:15 Standup", "text", "- 12:00 Lunch", "- 09:00 standup", "- 12:00 Lunch"];
  // Two visible ranges that share line 3; line 5 is off screen.
  const view = fakeEditorView(lines, [[1, 3], [3, 4]]);
  const links = new cls(view);
  assert.deepStrictEqual(
    links.decorations.map((d) => d.from),
    [21, 40, 56]
  );
  const built = links.decorations;
  links.update({ docChanged: false, viewportChanged: false, transactions: [{ effects: [] }] });
  assert.strictEqual(links.decorations, built, "moving the cursor does not rebuild");
  // New events: every open editor is told to redraw, once.
  await tick();
  view.dispatched = [];
  global.__handler = () => ({ body: { value: [graphEvent("c", "Lunch", "2026-11-17T11:00:00", "2026-11-17T12:00:00")] } });
  await plugin.loadDay("2026-11-17");
  plugin.eventsChanged();
  await tick();
  assert.strictEqual(view.dispatched.length, 1);
  links.update({ docChanged: false, viewportChanged: false, transactions: [{ effects: [view.dispatched[0].effects] }] });
  assert.deepStrictEqual(links.decorations.map((d) => d.from), [40]);
  links.destroy();
  assert.strictEqual(plugin.eventLinkViews.size, 0);
  // Signed out: no events, no links.
  plugin.signOut();
  assert.deepStrictEqual(new cls(view).decorations, []);
});

test("a link shows its event in a panel that is just opening, after the refresh", async () => {
  const plugin = await signedInPlugin();
  const event = (id, day) => graphEvent(id, "Standup", `${day}T08:00:00`, `${day}T08:15:00`);
  for (const day of ["2026-11-17", "2026-11-18"]) {
    global.__handler = () => ({ body: { value: [event(`s-${day}`, day)] } });
    await plugin.loadDay(day);
  }
  // Both days are old now: showing one fetches it again.
  for (const entry of plugin.dayCache().values()) entry.fetchedAt = 0;
  const { cls } = global.__editorExtension;
  const links = new cls(fakeEditorView(["- 09:00 Standup"], [[1, 1]], { basename: "2026-11-18" }));
  const widget = links.decorations[0].value.spec.widget;

  // The panel is closed: the link opens it.
  let view;
  plugin.app.workspace.getRightLeaf = () => ({
    async setViewState() {
      view = global.__viewFactory(this);
      view.app = plugin.app;
      view.contentEl = el();
      this.view = view;
      global.__leaves = [this];
      await view.onOpen();
    },
  });
  plugin.app.workspace.revealLeaf = async () => {};
  let finish;
  global.__handler = () => new Promise((resolve) => (finish = resolve));
  await plugin.revealEvent(widget.key, "2026-11-18");
  assert.strictEqual(view.day, "2026-11-18", "the note's own day wins");
  assert.ok(view.pendingReveal, "nothing to flash before the panel has started");
  global.__layoutReady();
  await new Promise((r) => setTimeout(r, 10));
  const row = () => view.listEl.find("oc-event")[0];
  assert.ok(row().cls.has("is-revealed"), "the cached row flashes at once");
  assert.ok(view.pendingReveal, "and again after the fetch");
  finish({ body: { value: [event("s-2026-11-18", "2026-11-18")] } });
  await new Promise((r) => setTimeout(r, 10));
  assert.ok(row().cls.has("is-revealed"), "the fetched row flashes");
  assert.strictEqual(row().scrolledTo, "center");
  assert.strictEqual(view.pendingReveal, null);
  view.render();
  assert.ok(!row().cls.has("is-revealed"), "a later render does not flash it again");

  // An open panel moves to another day; without a note day, the panel's wins.
  global.__handler = () => ({ body: { value: [event("s-2026-11-17", "2026-11-17")] } });
  await plugin.revealEvent(widget.key, null);
  assert.strictEqual(view.day, "2026-11-18");
  view.setDay("2026-11-20", false);
  await plugin.revealEvent(widget.key, "2026-11-17");
  await new Promise((r) => setTimeout(r, 10));
  assert.strictEqual(view.day, "2026-11-17");
  assert.ok(row().cls.has("is-revealed"));
  global.__notices = [];
  view.reveal("gone", "2026-11-17");
  assert.deepStrictEqual(global.__notices, ["That event is no longer in your calendar."]);
});

test("the + button hides for events the active note holds, checked only while on screen", async () => {
  const plugin = await signedInPlugin();
  global.__handler = () => ({ body: { value: [graphEvent("a", "Standup", "2026-11-17T08:00:00", "2026-11-17T08:15:00"), graphEvent("b", "Lunch", "2026-11-17T11:00:00", "2026-11-17T12:00:00")] } });
  let text = "# Today\n- 09:00–09:15 Standup";
  let reads = 0;
  const { MarkdownView } = require("./obsidian-stub.js");
  const note = Object.assign(new MarkdownView(), { file: { path: "Note.md", basename: "Note" }, getViewData: () => (reads++, text) });
  plugin.app.workspace.getActiveViewOfType = () => note;
  plugin.settings.followDailyNote = false;
  const view = await openView(plugin);
  view.day = "2026-11-17";
  global.__layoutReady();
  await new Promise((r) => setTimeout(r, 10));
  const inNote = () => view.listEl.find("oc-event").filter((r) => r.cls.has("is-in-note")).map((r) => r.ocEvent.id);
  assert.deepStrictEqual(inNote(), ["a"]);
  // Typing: one check per burst, and none when the note did not change.
  text += "\n* 12:00 lunch";
  plugin.scheduleNoteCheck(300);
  plugin.scheduleNoteCheck(300);
  await tick();
  assert.deepStrictEqual(inNote(), ["a", "b"]);
  const rows = view.listEl.find("oc-event");
  plugin.scheduleNoteCheck(300);
  await tick();
  assert.strictEqual(view.listEl.find("oc-event")[0], rows[0], "unchanged note: nothing redrawn");
  // Another day's daily note does not hold this day's events.
  note.file = { path: "2026-11-10.md", basename: "2026-11-10" };
  view.updateNoteMatches();
  assert.deepStrictEqual(inNote(), []);
  note.file = { path: "2026-11-17.md", basename: "2026-11-17" };
  view.updateNoteMatches();
  assert.deepStrictEqual(inNote(), ["a", "b"]);
  // A panel that is not on screen does no work while typing.
  view.contentEl.shown = false;
  reads = 0;
  plugin.scheduleNoteCheck(300);
  await tick();
  assert.strictEqual(reads, 0);
});

(async () => {
  let failed = 0;
  for (const [name, fn] of tests) {
    try {
      await fn();
      console.log(`ok   ${name}`);
    } catch (error) {
      failed++;
      console.log(`FAIL ${name}\n     ${error.stack.split("\n").slice(0, 4).join("\n     ")}`);
    }
  }
  console.log(`\n${tests.length - failed} passed, ${failed} failed`);
  process.exit(failed ? 1 : 0);
})();
