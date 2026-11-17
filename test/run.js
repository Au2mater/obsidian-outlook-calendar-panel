// Tests for startup cost, sign-in, reading the calendar and inserting events.
// Run with: node test/run.js   (TZ is pinned so day and DST checks are stable)
process.env.TZ = "Europe/Copenhagen";

const assert = require("assert");
const fs = require("fs");
const Module = require("module");
const path = require("path");

const stubPath = path.join(__dirname, "obsidian-stub.js");
const realResolve = Module._resolveFilename;
Module._resolveFilename = function (request, ...rest) {
  return request === "obsidian" ? stubPath : realResolve.call(this, request, ...rest);
};
const mainPath = path.join(__dirname, "..", "src", "plugin.js");
const source =
  fs.readFileSync(mainPath, "utf8") +
  "\nmodule.exports.__internals = { MicrosoftAuth, fetchDayEvents, eventView, dayEvents, dayBounds, addDays, placeOnOwnLine, insertOnOwnLine, describeAuthError, localToday };\n";
const mod = new Module(mainPath, module);
mod.filename = mainPath;
mod.paths = Module._nodeModulePaths(path.dirname(mainPath));
mod._compile(source, mainPath);
const CalendarPlugin = mod.exports;
const { fetchDayEvents, eventView, dayEvents, dayBounds, addDays, placeOnOwnLine, insertOnOwnLine, describeAuthError, localToday } =
  CalendarPlugin.__internals;

global.window = { setTimeout: (fn) => setImmediate(fn), clearTimeout() {}, setInterval, clearInterval };
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
  const events = await fetchDayEvents(plugin.auth, "2026-11-17");
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
