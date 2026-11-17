// Outlook Calendar Panel: a sidebar showing one day of your Outlook calendar.
// The plugin signs in to Microsoft itself (OAuth device code flow, read-only
// Calendars permission) and reads events from Microsoft Graph. Dragging an
// event into a note inserts it as a Markdown line, and so does its + button,
// below the cursor.
//
// Startup cost is kept near zero on purpose: onload only registers things and
// reads neither secret storage nor the network. A panel restored with the
// workspace shows its cached events, and fetches only once the layout is
// ready, the app is idle and the panel is on screen. The only work while
// typing is one shared timer that, once typing pauses for 300 ms and only
// while a panel is on screen, checks which events the active note holds; and
// the editor's icons beside event lines, which look only at the lines on
// screen.
//
// Plain JavaScript with no dependencies: `npm run build` checks this file and
// writes it unchanged to main.js, the file Obsidian loads (see README.md).

const {
  ItemView,
  MarkdownView,
  Modal,
  Notice,
  Plugin,
  PluginSettingTab,
  Setting,
  moment,
  requestUrl,
  editorInfoField,
  setIcon,
} = require("obsidian");
// Obsidian's own editor modules, for the links beside event lines in notes.
const { Decoration, ViewPlugin, WidgetType } = require("@codemirror/view");
const { RangeSetBuilder, StateEffect } = require("@codemirror/state");

const VIEW_TYPE = "outlook-calendar-panel";
const DRAG_TYPE = "application/x-outlook-calendar-event";
const GRAPH = "https://graph.microsoft.com/v1.0";
// Where the sign-in session lives: Obsidian's secret storage, on this device
// and outside the vault. The ID carries a hash of the vault name, so vaults
// signed in to different apps do not overwrite each other.
const SESSION_SECRET_PREFIX = "outlook-calendar-panel-session";
// The last fetched days, in this vault's local storage on this device, so the
// panel can show them at once on startup while it refreshes.
const EVENT_CACHE_KEY = "outlook-calendar-panel-days";
const CACHED_DAYS = 7;
// One tick a minute moves the "now" marker, rolls over at midnight and
// refreshes a stale day, and only for a panel that is on screen.
const TICK_MS = 60 * 1000;

const DEFAULT_SETTINGS = {
  // Microsoft sign-in. There is no shared app registration: each user
  // registers their own (see README.md) and pastes its IDs here.
  clientId: "",
  tenant: "common",
  scopes: "Calendars.Read",
  template: "- {{time}} {{title}}",
  timeFormat: "HH:mm",
  // How daily note file names spell their date (Moment.js format).
  dailyNoteFormat: "YYYY-MM-DD",
  followDailyNote: true,
  showDeclined: false,
  // The calendars to show, as { id, name, color }; empty shows the default
  // calendar only.
  calendars: [],
  refreshMinutes: 15,
};

const EVENT_FIELDS = [
  "id",
  "subject",
  "start",
  "end",
  "isAllDay",
  "isCancelled",
  "showAs",
  "location",
  "onlineMeeting",
  "onlineMeetingUrl",
  "webLink",
  "responseStatus",
  "organizer",
  "bodyPreview",
].join(",");

module.exports = class OutlookCalendarPlugin extends Plugin {
  async onload() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    this.registerView(VIEW_TYPE, (leaf) => new CalendarPanelView(leaf, this));
    this.addRibbonIcon("calendar-days", "Open Outlook calendar", () => this.openPanel());
    this.addCommand({ id: "open-panel", name: "Open panel", callback: () => this.openPanel() });
    this.addCommand({
      id: "refresh",
      name: "Refresh events",
      callback: () => this.panels().forEach((view) => view.refresh(true)),
    });
    this.addCommand({
      id: "insert-agenda",
      name: "Insert the day's agenda",
      callback: () => this.insertAgenda().catch((e) => new Notice(`Could not insert the agenda: ${e.message}`)),
    });
    this.settingTab = new CalendarSettingTab(this.app, this);
    this.addSettingTab(this.settingTab);

    // Follow the daily note being worked on, and mark the events the active
    // note holds. Opening a note fires several events at once, and typing
    // fires one per keystroke, so they share one short timer and the note is
    // checked once per burst. With no panel on screen the handlers return at
    // once.
    const follow = () => {
      if (this.panels().length) this.scheduleFollow();
    };
    this.registerEvent(this.app.workspace.on("file-open", follow));
    this.registerEvent(this.app.workspace.on("active-leaf-change", follow));
    this.registerEvent(this.app.workspace.on("editor-change", () => this.scheduleNoteCheck(300)));
    this.registerEvent(
      this.app.vault.on("modify", (file) => {
        if (file === this.app.workspace.getActiveFile()) this.scheduleNoteCheck(300);
      })
    );

    // An event dropped into a note lands on a line of its own.
    this.registerEvent(
      this.app.workspace.on("editor-drop", (evt, editor) => {
        const dt = evt.dataTransfer;
        if (!dt || !Array.from(dt.types || []).includes(DRAG_TYPE)) return;
        const text = dt.getData("text/plain");
        window.setTimeout(() => placeOnOwnLine(editor, text), 0);
      })
    );

    // A link beside each note line that holds a panel event, which shows the
    // event in the panel. A visual layer only: the note's text is untouched.
    this.eventIndex = null;
    this.eventIndexDirty = true;
    this.eventShape = this.eventLineShape();
    this.eventLinkViews = new Set();
    this.registerEditorExtension(eventLinkExtension(this));

    this.registerInterval(window.setInterval(() => this.tick(), TICK_MS));
  }

  onunload() {
    window.clearTimeout(this.followTimer);
    window.clearTimeout(this.noteCheckTimer);
    window.clearTimeout(this.eventLinkTimer);
    this.eventLinkViews.clear();
    if (this.signInModal) this.signInModal.close();
  }

  // Created on first use, so loading the plugin never touches secret storage.
  get auth() {
    if (!this._auth) this._auth = new MicrosoftAuth(this);
    return this._auth;
  }

  tick() {
    for (const view of this.panels()) if (view.isVisible()) view.tick();
  }

  scheduleFollow() {
    window.clearTimeout(this.followTimer);
    this.followTimer = window.setTimeout(() => {
      const date = this.dailyNoteDate(this.app.workspace.getActiveFile());
      this.panels().forEach((view) => view.followNote(date));
      this.panels().forEach((view) => view.updateNoteMatches());
    }, 50);
  }

  // Typing and saving: the active note is checked once the burst is over,
  // and only by a panel on screen.
  scheduleNoteCheck(delay) {
    if (!this.panels().some((view) => view.isVisible())) return;
    window.clearTimeout(this.noteCheckTimer);
    this.noteCheckTimer = window.setTimeout(() => {
      this.panels().forEach((view) => view.updateNoteMatches());
    }, delay);
  }

  // The active note's text, or null.
  activeNote() {
    const view = this.noteView();
    return view && view.file ? { file: view.file, text: view.getViewData() } : null;
  }

  panels() {
    return this.app.workspace
      .getLeavesOfType(VIEW_TYPE)
      .map((leaf) => leaf.view)
      .filter((view) => view instanceof CalendarPanelView);
  }

  // Opens the panel, or shows it, and returns it (null if it cannot be).
  async openPanel() {
    let leaf = this.app.workspace.getLeavesOfType(VIEW_TYPE)[0];
    if (!leaf) {
      leaf = this.app.workspace.getRightLeaf(false);
      if (!leaf) return null;
      await leaf.setViewState({ type: VIEW_TYPE, active: true });
    }
    await this.app.workspace.revealLeaf(leaf);
    // A panel restored in a background tab loads only when first shown.
    if (typeof leaf.loadIfDeferred === "function") await leaf.loadIfDeferred();
    return leaf.view instanceof CalendarPanelView ? leaf.view : null;
  }

  async saveSettings() {
    const shape = this.eventLineShape();
    if (shape !== this.eventShape) this.eventsChanged();
    this.eventShape = shape;
    await this.saveData(this.settings);
  }

  // The settings that decide what an event line looks like, and which
  // events are shown.
  eventLineShape() {
    const s = this.settings;
    return `${s.template}\u0000${s.timeFormat}\u0000${s.showDeclined}`;
  }

  // ── links from note lines to the panel ──

  // The fetched events changed: the links in open notes are redrawn, once,
  // after whatever changed them has finished.
  eventsChanged() {
    this.eventIndexDirty = true;
    window.clearTimeout(this.eventLinkTimer);
    this.eventLinkTimer = window.setTimeout(() => {
      for (const links of this.eventLinkViews) links.redraw();
    }, 0);
  }

  // What note lines are matched against, built on first use after a change:
  // the events of every day fetched or read from the cache this session, as
  // the panel shows them. Null when there are none. It never reads the
  // sign-in or the cache itself: until a panel has done so, there are no
  // links, so an editor opened at startup costs nothing.
  eventIndexNow() {
    if (this.eventIndexDirty) {
      this.eventIndexDirty = false;
      this.eventIndex = this.days instanceof Map ? buildEventIndex(this.days, this) : null;
    }
    return this.eventIndex;
  }

  // The keys an event's line can be found by, on a given day (see eventKeys).
  eventKeys(event, day) {
    return eventKeys(this.renderEvent(event, day), event.title, this.timeLabel(event, day), event.allDay ? "" : this.formatTime(event.start));
  }

  // Clicking a link in a note: shows that event in the panel. When the line
  // matches several events (a recurring one, on each cached day), the one on
  // the note's own day wins, then the panel's day, then the nearest to today.
  async revealEvent(key, noteDay) {
    const index = this.eventIndexNow();
    const candidates = (index && index.get(key)) || [];
    if (!candidates.length) {
      new Notice("That event is no longer in the panel's calendar.");
      return;
    }
    const shown = this.panels()[0];
    const today = localToday();
    const distance = (day) => Math.abs(parseDay(day) - parseDay(today));
    const pick =
      candidates.find((c) => c.day === noteDay) ||
      candidates.find((c) => shown && c.day === shown.day) ||
      candidates.slice().sort((a, b) => distance(a.day) - distance(b.day) || (a.day < b.day ? 1 : -1))[0];
    const view = await this.openPanel();
    if (view) view.reveal(pick.id, pick.day);
  }

  signIn() {
    if (this.signInModal) return;
    this.signInModal = new SignInModal(this.app, this);
    this.signInModal.open();
  }

  signOut() {
    this.auth.signOut();
    this.calendarList = null;
    this.resetDays();
    this.saveDayCache();
    this.panels().forEach((view) => view.clear());
    this.refreshSettingTab();
    new Notice("Signed out of Microsoft.");
  }

  // The sign-in expired: forget its events, as signing out does.
  sessionEnded() {
    this.calendarList = null;
    this.resetDays();
    this.saveDayCache();
    this.refreshSettingTab();
  }

  signedIn() {
    const who = this.auth.account ? ` as ${this.auth.account}` : "";
    new Notice(`Signed in to Microsoft${who}.`);
    this.calendarList = null;
    this.resetDays();
    this.panels().forEach((view) => view.refresh(true));
    this.refreshSettingTab();
  }

  refreshSettingTab() {
    if (this.settingTab && this.settingTab.containerEl && this.settingTab.containerEl.isConnected) this.settingTab.display();
  }

  // ── fetched days ──
  //
  // One shared map from day (YYYY-MM-DD) to { events, fetchedAt }, filled
  // from local storage the first time a panel needs it.

  dayCache() {
    if (!this.days) {
      this.days = this.loadDayCache();
      this.eventsChanged();
    }
    return this.days;
  }

  // Forgets every fetched day: a new sign-in or choice of calendars.
  resetDays() {
    this.days = new Map();
    this.eventsChanged();
  }

  loadDayCache() {
    try {
      const cached = this.app.loadLocalStorage(EVENT_CACHE_KEY);
      if (!cached || !cached.days || !this.auth.signedIn || cached.account !== this.auth.account) return new Map();
      // Days fetched for another choice of calendars do not count.
      if (cached.calendars !== this.calendarKey()) return new Map();
      return new Map(Object.entries(cached.days));
    } catch (_) {
      return new Map();
    }
  }

  saveDayCache() {
    try {
      const days = [...(this.days || new Map()).entries()]
        .sort((a, b) => b[1].fetchedAt - a[1].fetchedAt)
        .slice(0, CACHED_DAYS);
      this.app.saveLocalStorage(
        EVENT_CACHE_KEY,
        days.length ? { account: this.auth.account, calendars: this.calendarKey(), days: Object.fromEntries(days) } : null
      );
    } catch (_) {
      // A full or unavailable local storage only costs the instant start.
    }
  }

  // The chosen calendars, for the account they were chosen with only.
  calendars() {
    const s = this.settings;
    if (!Array.isArray(s.calendars) || s.calendarsAccount !== this.auth.account) return [];
    return s.calendars.filter((c) => c && c.id);
  }

  calendarKey() {
    return this.calendars().map((c) => c.id).join(",");
  }

  // A new choice of calendars: what was fetched for the old one is dropped.
  async setCalendars(calendars) {
    this.settings.calendars = calendars;
    this.settings.calendarsAccount = this.auth.account;
    await this.saveSettings();
    this.resetDays();
    this.saveDayCache();
    this.panels().forEach((view) => view.refresh(true));
  }

  // The user's calendars, for the settings tab, fetched once per session.
  listCalendars(force = false) {
    if (!this.calendarList || force) {
      this.calendarList = fetchCalendars(this.auth).catch((e) => {
        this.calendarList = null;
        throw e;
      });
    }
    return this.calendarList;
  }

  // Fetches a day's events from Microsoft Graph and keeps them, with the
  // calendars that could not be read.
  async loadDay(day) {
    const auth = this.auth;
    const epoch = auth.epoch;
    const key = this.calendarKey();
    const { events, failures } = await fetchDayEvents(auth, day, this.calendars());
    // Signed out, or into another account, while this was loading.
    if (auth.epoch !== epoch) throw new NotSignedInError("Not signed in to Microsoft.");
    // Other calendars were chosen meanwhile; these are for the old choice.
    if (key !== this.calendarKey()) return events;
    this.dayCache().set(day, { events, failures, fetchedAt: Date.now() });
    this.eventsChanged();
    this.saveDayCache();
    return events;
  }

  // The events to show for a day: cancelled ones never, declined ones only
  // when asked for.
  visibleEvents(events) {
    return (events || []).filter((e) => !e.cancelled && (this.settings.showDeclined || e.response !== "declined"));
  }

  // ── Markdown ──

  formatTime(ms) {
    const format = (this.settings.timeFormat || "").trim() || "HH:mm";
    return moment(ms).format(format);
  }

  // The time an event takes on a given day, clipped to that day: "09:00–10:00",
  // "All day", "Until 10:00" or "From 22:00".
  timeLabel(event, day) {
    if (event.allDay) return "All day";
    const { start, end } = dayBounds(day);
    const startsBefore = event.start < start;
    const endsAfter = event.end > end;
    if (startsBefore && endsAfter) return "All day";
    if (startsBefore) return `Until ${this.formatTime(event.end)}`;
    if (endsAfter) return `From ${this.formatTime(event.start)}`;
    if (event.end === event.start) return this.formatTime(event.start);
    return `${this.formatTime(event.start)}–${this.formatTime(event.end)}`;
  }

  renderEvent(event, day) {
    const values = {
      title: event.title,
      time: this.timeLabel(event, day),
      start: event.allDay ? "" : this.formatTime(event.start),
      end: event.allDay ? "" : this.formatTime(event.end),
      date: day,
      location: event.location || "",
      organizer: event.organizer || "",
      link: event.webLink || "",
      join: event.joinUrl || "",
      calendar: event.calendar || "",
    };
    return fill(this.settings.template, values)
      .replace(/[ \t]{2,}/g, " ")
      .replace(/[ \t]+$/gm, "")
      .trim();
  }

  dailyNoteDate(file) {
    if (!file || !file.basename) return null;
    const format = (this.settings.dailyNoteFormat || "").trim() || "YYYY-MM-DD";
    // A format with folders ("YYYY/MM/DD") is matched against as many of
    // the path's last parts, so a note named "17" alone is no daily note.
    const parts = format.split("/").length;
    const name = parts > 1 && file.path ? file.path.replace(/\.md$/i, "").split("/").slice(-parts).join("/") : file.basename;
    const parsed = moment(name, format, true);
    return parsed.isValid() ? parsed.format("YYYY-MM-DD") : null;
  }

  // ── inserting into notes ──

  noteView() {
    const active = this.app.workspace.getActiveViewOfType(MarkdownView);
    if (active) return active;
    const recent = this.app.workspace.getMostRecentLeaf();
    return recent && recent.view instanceof MarkdownView ? recent.view : null;
  }

  async noteEditor() {
    const view = this.noteView();
    if (!view) {
      new Notice("Open a note first.");
      return null;
    }
    if (view.getMode() === "preview") await view.setState({ ...view.getState(), mode: "source" }, { history: false });
    this.app.workspace.setActiveLeaf(view.leaf, { focus: true });
    return { view, editor: view.editor };
  }

  // Inserts text below the cursor's line, never inside the note's properties.
  async insertIntoNote(text) {
    const target = await this.noteEditor();
    if (!target) return;
    const { view, editor } = target;
    const cache = view.file ? this.app.metadataCache.getFileCache(view.file) : null;
    const properties = cache && cache.frontmatterPosition;
    let line = editor.getCursor().line;
    if (properties && line <= properties.end.line) line = properties.end.line;
    insertOnOwnLine(editor, text, line);
    editor.focus();
  }

  addToNote(event, day) {
    return this.insertIntoNote(this.renderEvent(event, day));
  }

  // Every event of a day as lines: the panel's day, else the active daily
  // note's, else today.
  async insertAgenda() {
    const view = this.panels()[0];
    const note = this.noteView();
    const day = (view && view.day) || this.dailyNoteDate(note && note.file) || localToday();
    if (!this.auth.signedIn) {
      new Notice("Sign in to Microsoft first.");
      return;
    }
    const cached = this.dayCache().get(day);
    const events = this.visibleEvents(cached ? cached.events : await this.loadDay(day));
    if (!events.length) {
      new Notice(`No events on ${longDate(day)}.`);
      return;
    }
    await this.insertIntoNote(events.map((e) => this.renderEvent(e, day)).join("\n"));
  }
};

class CalendarPanelView extends ItemView {
  constructor(leaf, plugin) {
    super(leaf);
    this.plugin = plugin;
    this.day = localToday();
    // True while the day shown came from the active daily note or is today
    // by default; false once the user picked a day themselves.
    this.following = true;
    this.loading = false;
    this.queued = false;
    this.error = null;
    this.ready = false;
    this.openDetails = new Set();
    this.lastToday = localToday();
    // The event shown from a note, until the panel has settled on its day
    // (see applyReveal), and the keys of the active note's lines.
    this.pendingReveal = null;
    this.noteKeys = null;
    this.checkedNote = null;
    this.rowsById = new Map();
  }

  getViewType() {
    return VIEW_TYPE;
  }

  getDisplayText() {
    return "Outlook calendar";
  }

  getIcon() {
    return "calendar-days";
  }

  async onOpen() {
    const root = this.contentEl;
    root.empty();
    root.addClass("oc-panel");

    const nav = root.createDiv({ cls: "oc-nav" });
    const prev = nav.createEl("button", { attr: { "aria-label": "Previous day" } });
    setIcon(prev, "chevron-left");
    prev.addEventListener("click", () => this.setDay(addDays(this.day, -1), false));
    this.titleEl = nav.createDiv({ cls: "oc-title" });
    const next = nav.createEl("button", { attr: { "aria-label": "Next day" } });
    setIcon(next, "chevron-right");
    next.addEventListener("click", () => this.setDay(addDays(this.day, 1), false));
    this.todayEl = nav.createEl("button", { cls: "oc-today", text: "Today" });
    this.todayEl.addEventListener("click", () => this.setDay(localToday(), true));
    const refresh = nav.createEl("button", { attr: { "aria-label": "Refresh" } });
    setIcon(refresh, "refresh-cw");
    refresh.addEventListener("click", () => this.refresh(true));

    this.statusEl = root.createDiv({ cls: "oc-status" });
    this.listEl = root.createDiv({ cls: "oc-events" });

    // Start on the open daily note's day, if any (cheap: a file name parse).
    const noteDay = this.plugin.settings.followDailyNote
      ? this.plugin.dailyNoteDate(this.app.workspace.getActiveFile())
      : null;
    if (noteDay) this.day = noteDay;
    this.noteDate = noteDay;
    // Only the title for now: the events need the sign-in and the cache,
    // which are read once the app is idle, not while Obsidian waits for this.
    this.render();

    // Obsidian waits for onOpen while it restores the workspace, so the fetch
    // starts only once the layout is ready and the app is idle, and is never
    // awaited here.
    this.app.workspace.onLayoutReady(() => {
      const start = () => {
        this.ready = true;
        // The active note may not have been known yet while restoring.
        if (this.plugin.settings.followDailyNote) this.followNote(this.plugin.dailyNoteDate(this.app.workspace.getActiveFile()));
        // A stale day starts loading before the first render, so an event
        // shown from a note stays pending until the fetched rows are drawn.
        if (this.isVisible()) this.refresh(false);
        this.render();
        if (this.settled()) this.settleReveal();
        this.updateNoteMatches();
      };
      if (window.requestIdleCallback) window.requestIdleCallback(start, { timeout: 3000 });
      else window.setTimeout(start, 500);
    });

    // A panel in a collapsed sidebar or a background tab does no work; when
    // it comes on screen, it catches up.
    if (typeof IntersectionObserver === "function") {
      this.observer = new IntersectionObserver((entries) => {
        if (this.ready && entries.some((e) => e.isIntersecting)) this.tick();
      });
      this.observer.observe(root);
    }
  }

  async onClose() {
    if (this.observer) this.observer.disconnect();
  }

  isVisible() {
    const el = this.contentEl;
    if (!el || el.isConnected === false) return false;
    if (typeof el.isShown === "function") return el.isShown();
    return el.offsetParent !== undefined ? el.offsetParent !== null : true;
  }

  // Once a minute while on screen, and when the panel comes into view.
  tick() {
    this.updateNoteMatches();
    const today = localToday();
    if (this.lastToday !== today) {
      const yesterday = this.lastToday;
      this.lastToday = today;
      // Past midnight, a panel showing today by default moves on; one that
      // follows an open daily note, or shows a day the user picked, stays.
      if (this.following && !this.noteDate && this.day === yesterday) {
        this.setDay(today, true);
        return;
      }
      this.render(); // the "Today" title and button
    }
    // Signed out, the sign-in box stays as it is (and keeps its focus).
    if (!this.plugin.auth.signedIn) return;
    if (this.isStale()) this.refresh(false);
    // Moves the "now" highlight, rebuilding only when an event's state changed.
    else if (this.day === today && this.nowKey !== this.timeStates()) this.render();
  }

  // Which of the shown events are past and which happening now, as a string.
  timeStates() {
    const entry = this.plugin.dayCache().get(this.day);
    const now = Date.now();
    return entry ? entry.events.map((e) => (e.allDay ? "a" : e.end <= now ? "p" : e.start <= now ? "n" : "f")).join("") : "";
  }

  isStale() {
    const entry = this.plugin.dayCache().get(this.day);
    const minutes = Math.max(1, Number(this.plugin.settings.refreshMinutes) || 15);
    return !entry || Date.now() - entry.fetchedAt > minutes * 60 * 1000;
  }

  // The active note may have changed: show its day if it is a daily note.
  // Leaving a daily note goes back to today, unless the user picked a day.
  // Only a change of note counts, so clicking in the panel (which also
  // changes the active leaf) never undoes a day the user picked.
  followNote(date) {
    if (!this.plugin.settings.followDailyNote || date === this.noteDate) return;
    const left = this.noteDate;
    this.noteDate = date;
    if (date) this.setDay(date, true);
    else if (left && this.following && this.day !== localToday()) this.setDay(localToday(), true);
  }

  setDay(day, following) {
    this.day = day;
    this.following = following;
    this.error = null;
    this.render();
    // A panel that is not on screen fetches when it is shown.
    if (this.ready && this.isVisible() && this.isStale()) this.refresh(false);
  }

  // Fetches the day shown. `force` refetches even a fresh day.
  async refresh(force) {
    if (this.loading) {
      // A forced refresh runs again once this one ends; any other is covered
      // by it (or by the day check below).
      if (force) this.queued = true;
      return;
    }
    if (!this.plugin.auth.signedIn) {
      this.clear();
      return;
    }
    if (!force && !this.isStale()) return;
    const day = this.day;
    this.loading = true;
    this.error = null;
    this.renderStatus();
    try {
      await this.plugin.loadDay(day);
    } catch (e) {
      // Shown only on the day it belongs to; an expired sign-in shows anyway.
      if ((this.plugin.auth.signedIn && this.day === day) || e.expired) this.error = e.message;
    } finally {
      this.loading = false;
      this.render();
      // The fetched rows are drawn: an event shown from a note has had its
      // flash, or is gone. One that failed to fetch is not retried each
      // minute either.
      if (this.day === day || this.settled()) this.settleReveal();
    }
    // Asked again while loading, or the day changed meanwhile.
    if (this.queued || this.day !== day) {
      const again = this.queued;
      this.queued = false;
      this.refresh(again);
    }
  }

  // Signed out: no events, and a way to sign in (with the reason, if any).
  clear() {
    this.render();
  }

  renderStatus() {
    if (!this.statusEl) return;
    this.statusEl.empty();
    const entry = this.plugin.auth.signedIn ? this.plugin.dayCache().get(this.day) : null;
    if (!this.plugin.auth.signedIn && !this.loading) {
      const box = this.statusEl.createDiv({ cls: "oc-signin" });
      if (this.error) box.createDiv({ cls: "oc-error", text: this.error });
      box.createDiv({ text: "Sign in to Microsoft to show your Outlook calendar." });
      const button = box.createEl("button", { cls: "mod-cta", text: "Sign in" });
      button.addEventListener("click", () => this.plugin.signIn());
    } else if (this.loading) {
      this.statusEl.setText(entry ? "Refreshing…" : "Loading events…");
    } else if (this.error) {
      this.statusEl.createDiv({ cls: "oc-error", text: this.error });
    } else if (entry) {
      const time = new Date(entry.fetchedAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
      this.statusEl.createDiv({ text: `Updated ${time} · drag an event into a note, or use +` });
      for (const f of entry.failures || []) {
        this.statusEl.createDiv({ cls: "oc-error", text: `Could not read ${f.name}: ${f.message}` });
      }
    }
  }

  render() {
    if (this.titleEl) {
      const today = localToday();
      this.titleEl.setText(this.day === today ? `Today · ${longDate(this.day)}` : longDate(this.day));
      this.todayEl.toggleClass("is-hidden", this.day === today);
    }
    if (!this.ready) return;
    this.renderStatus();
    if (!this.listEl) return;
    this.listEl.empty();
    this.rowsById = new Map();
    if (!this.plugin.auth.signedIn) return;
    const entry = this.plugin.dayCache().get(this.day);
    this.nowKey = this.timeStates();
    if (!entry) return;
    const events = this.plugin.visibleEvents(dayEvents(entry.events, this.day));
    if (!events.length) {
      this.listEl.createDiv({ cls: "oc-empty", text: "No events" });
      return;
    }
    const now = Date.now();
    for (const event of events) this.renderRow(event, now);
    if (this.pendingReveal) this.applyReveal();
  }

  // Shows one event from a note: moves to its day, scrolls to it and flashes
  // it. A panel that is just opening, or a day that is still being fetched,
  // rebuilds its rows once more; the event stays pending and is shown again
  // after that, so the flash is not cut short.
  reveal(id, day) {
    this.pendingReveal = { id, day };
    // The open daily note's day, or today, still counts as followed.
    if (this.day !== day) this.setDay(day, day === this.noteDate || (!this.noteDate && day === localToday()));
    else this.render();
    if (this.settled()) this.settleReveal();
  }

  // Whether the rows drawn are the ones the panel will keep: it has started
  // and is not fetching, nor about to fetch, the day shown.
  settled() {
    return this.ready && !this.loading && !(this.plugin.auth.signedIn && this.isVisible() && this.isStale());
  }

  applyReveal() {
    const row = this.rowsById.get(this.pendingReveal.id);
    if (!row) return false;
    row.scrollIntoView({ block: "center" });
    row.removeClass("is-revealed");
    // Restarts the flash when the same event is shown twice in a row.
    void row.offsetWidth;
    row.addClass("is-revealed");
    return true;
  }

  // Ends a pending reveal once the panel has settled, with a notice when the
  // event is not there any more.
  settleReveal() {
    const pending = this.pendingReveal;
    if (!pending) return;
    this.pendingReveal = null;
    // Another day was picked meanwhile, or the fetch failed and says so.
    if (pending.day !== this.day || this.error) return;
    if (!this.rowsById.has(pending.id)) new Notice("That event is no longer in your calendar.");
  }

  // Re-reads the active note and marks the events it already holds, without
  // rebuilding the panel. Typing, saving and switching tabs fire often, so
  // the work is skipped when the note and the settings that shape matching
  // are what they were, and while the panel is not on screen.
  updateNoteMatches() {
    if (!this.ready || !this.isVisible()) return;
    const note = this.plugin.activeNote();
    const checked = { path: note && note.file.path, text: note && note.text, shape: this.plugin.eventShape };
    const last = this.checkedNote;
    if (last && last.path === checked.path && last.text === checked.text && last.shape === checked.shape) return;
    this.checkedNote = checked;
    this.noteKeys = note ? noteLineKeys(note.text) : null;
    // A daily note holds the events of its own day only: last Monday's
    // "- 09:00 Standup" is not today's.
    this.noteKeysDay = note ? this.plugin.dailyNoteDate(note.file) : null;
    for (const row of this.rowsById.values()) this.markNoteMatch(row);
  }

  // Whether the active note holds this event on the day shown: the same
  // matching as the links beside note lines (see lineKeys and eventKeys).
  inNote(event) {
    if (!this.noteKeys || (this.noteKeysDay && this.noteKeysDay !== this.day)) return false;
    return this.plugin.eventKeys(event, this.day).some((key) => this.noteKeys.has(key));
  }

  markNoteMatch(row) {
    if (row.ocEvent) row.toggleClass("is-in-note", this.inNote(row.ocEvent));
  }

  renderRow(event, now) {
    const day = this.day;
    const row = this.listEl.createDiv({ cls: "oc-event", attr: { draggable: "true" } });
    if (event.allDay) row.addClass("is-all-day");
    else if (event.end === event.start ? event.start < now : event.end <= now) row.addClass("is-past");
    else if (event.start <= now && now < event.end) row.addClass("is-now");
    if (event.showAs) row.addClass(`is-${event.showAs}`);
    if (event.color) setCssProps(row, { "--oc-calendar-color": event.color });
    if (event.response === "declined") row.addClass("is-declined");
    row.ocEvent = event;
    this.rowsById.set(event.id, row);
    this.markNoteMatch(row);

    row.createDiv({ cls: "oc-time", text: this.plugin.timeLabel(event, day) });
    const body = row.createDiv({ cls: "oc-body" });
    body.createDiv({ cls: "oc-subject", text: event.title });
    const meta = body.createDiv({ cls: "oc-meta" });
    if (event.location) meta.createSpan({ cls: "oc-location", text: event.location });
    if (event.calendar && this.plugin.calendars().length > 1) meta.createSpan({ cls: "oc-calendar", text: event.calendar });
    if (event.showAs === "tentative" || event.response === "tentativelyAccepted") meta.createSpan({ text: "Tentative" });
    if (event.response === "notResponded") meta.createSpan({ cls: "oc-unanswered", text: "Not answered" });
    if (event.joinUrl) {
      const join = meta.createEl("a", { cls: "oc-join", text: "Join", href: event.joinUrl, attr: { target: "_blank", rel: "noopener" } });
      join.addEventListener("click", (evt) => evt.stopPropagation());
    }

    const details = body.createDiv({ cls: "oc-details" });
    if (event.organizer) details.createDiv({ text: `Organiser: ${event.organizer}` });
    if (event.preview) details.createDiv({ cls: "oc-preview", text: event.preview });
    if (event.webLink) {
      const open = details.createEl("a", { text: "Open in Outlook", href: event.webLink, attr: { target: "_blank", rel: "noopener" } });
      open.addEventListener("click", (evt) => evt.stopPropagation());
    }
    const show = (open) => details.toggleClass("is-hidden", !open);
    show(this.openDetails.has(event.id));
    details.addEventListener("click", (evt) => evt.stopPropagation());
    row.addEventListener("click", () => {
      const open = !this.openDetails.has(event.id);
      if (open) this.openDetails.add(event.id);
      else this.openDetails.delete(event.id);
      show(open);
    });

    const add = row.createEl("button", { cls: "oc-add", attr: { "aria-label": "Add to note" } });
    setIcon(add, "plus");
    add.addEventListener("click", (evt) => {
      evt.stopPropagation();
      this.plugin.addToNote(event, day).catch((e) => new Notice(`Could not add the event: ${e.message}`));
    });

    row.addEventListener("dragstart", (evt) => {
      evt.dataTransfer.setData("text/plain", this.plugin.renderEvent(event, day));
      evt.dataTransfer.setData(DRAG_TYPE, event.id);
      row.addClass("is-dragging");
    });
    row.addEventListener("dragend", () => row.removeClass("is-dragging"));
  }
}

class CalendarSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    const s = this.plugin.settings;
    const auth = this.plugin.auth;
    containerEl.empty();

    const account = new Setting(containerEl);
    if (auth.signedIn) {
      account
        .setName(auth.account ? `Signed in as ${auth.account}` : "Signed in to Microsoft")
        .setDesc("The sign-in is kept in Obsidian's secret storage on this device, never in the vault.")
        .addButton((button) => button.setButtonText("Sign out").onClick(() => this.plugin.signOut()));
    } else {
      account
        .setName("Not signed in")
        .setDesc("Register an app in Microsoft Entra (see the README), enter its IDs below, then sign in.")
        .addButton((button) => button.setButtonText("Sign in").setCta().onClick(() => this.plugin.signIn()));
      this.appSettings(containerEl);
    }

    this.text(
      containerEl,
      "Event template",
      "What an added event becomes. Placeholders: {{time}}, {{start}}, {{end}}, {{title}}, {{date}}, {{location}}, {{organizer}}, {{calendar}}, {{link}}, {{join}}.",
      "template",
      () => this.plugin.panels().forEach((view) => view.render())
    );

    if (auth.signedIn) this.calendarSettings(containerEl);

    const advanced = containerEl.createEl("details", { cls: "oc-advanced" });
    advanced.createEl("summary", { text: "Advanced" });
    if (auth.signedIn) this.appSettings(advanced);
    this.text(advanced, "Time format", "How times are written, in Moment.js format, such as HH:mm or h:mm A.", "timeFormat", () =>
      this.plugin.panels().forEach((view) => view.render())
    );
    this.text(advanced, "Daily note date format", "How daily note file names spell their date, in Moment.js format.", "dailyNoteFormat");
    new Setting(advanced)
      .setName("Follow the daily note")
      .setDesc("Show the day of the daily note you open.")
      .addToggle((t) =>
        t.setValue(s.followDailyNote).onChange(async (v) => {
          s.followDailyNote = v;
          await this.plugin.saveSettings();
        })
      );
    new Setting(advanced)
      .setName("Show declined events")
      .addToggle((t) =>
        t.setValue(s.showDeclined).onChange(async (v) => {
          s.showDeclined = v;
          await this.plugin.saveSettings();
          this.plugin.panels().forEach((view) => view.render());
        })
      );
    new Setting(advanced)
      .setName("Refresh after (minutes)")
      .setDesc("How old a shown day may get before the panel fetches it again, while the panel is on screen.")
      .addText((text) => {
        text.setValue(String(s.refreshMinutes));
        text.inputEl.addEventListener("change", async () => {
          const minutes = parseInt(text.getValue(), 10);
          if (!Number.isFinite(minutes) || minutes < 1) return;
          s.refreshMinutes = minutes;
          await this.plugin.saveSettings();
        });
      });
    this.authText(
      advanced,
      "Permissions",
      "Microsoft Graph permissions to ask for. Calendars.Read covers your own calendars; add Calendars.Read.Shared if a calendar shared with you cannot be read.",
      "scopes",
      "Calendars.Read"
    );
  }

  // A toggle per calendar, filled in once Microsoft has listed them. Only
  // fetched while the settings tab is open.
  calendarSettings(el) {
    new Setting(el)
      .setName("Calendars")
      .setDesc("Which calendars the panel shows. With none chosen, it shows your default calendar.")
      .setHeading();
    const list = el.createDiv({ cls: "oc-calendar-list" });
    const status = list.createDiv({ cls: "setting-item-description", text: "Loading your calendars…" });
    const tab = this;
    this.plugin
      .listCalendars()
      .then((calendars) => {
        if (list.isConnected === false) return;
        status.remove();
        const chosen = () => {
          const ids = tab.plugin.calendars().map((c) => c.id);
          return ids.length ? ids : calendars.filter((c) => c.isDefault).map((c) => c.id);
        };
        for (const calendar of calendars) {
          const setting = new Setting(list).setName(calendar.name);
          if (calendar.owner) setting.setDesc(calendar.owner);
          if (calendar.color) setCssProps(setting.nameEl, { "--oc-calendar-color": calendar.color });
          setting.nameEl.addClass("oc-calendar-name");
          setting.addToggle((toggle) =>
            toggle.setValue(chosen().includes(calendar.id)).onChange(async (on) => {
              const ids = new Set(chosen());
              if (on) ids.add(calendar.id);
              else ids.delete(calendar.id);
              const next = calendars
                .filter((c) => ids.has(c.id))
                .map((c) => ({ id: c.id, name: c.name, color: c.color }));
              // Only the default calendar is the same as none chosen.
              const onlyDefault = next.length === 1 && calendars.find((c) => c.id === next[0].id).isDefault;
              await tab.plugin.setCalendars(next.length === 0 || onlyDefault ? [] : next);
              if (!next.length) toggle.setValue(chosen().includes(calendar.id));
            })
          );
        }
      })
      .catch((e) => {
        status.setText(`Could not list your calendars: ${e.message}`);
        status.addClass("oc-error");
      });
  }

  appSettings(el) {
    this.authText(el, "Client ID", "The Application (client) ID of the app you registered in Microsoft Entra.", "clientId", "00000000-0000-0000-0000-000000000000");
    this.authText(el, "Tenant", "Your directory (tenant) ID if the app accepts only accounts in your organisation; otherwise common or consumers.", "tenant", "common");
  }

  // Applies when the field is left, and ends a sign-in made for the old value.
  authText(el, name, desc, key, placeholder) {
    const s = this.plugin.settings;
    const auth = this.plugin.auth;
    new Setting(el)
      .setName(name)
      .setDesc(desc)
      .addText((text) => {
        text.setPlaceholder(placeholder).setValue(s[key]);
        text.inputEl.addEventListener("change", async () => {
          const value = text.getValue().trim();
          if (value === s[key]) return;
          const wasSignedIn = auth.signedIn;
          s[key] = value;
          // The sign-in dialog's remembered IDs would otherwise undo this.
          if (this.plugin.signInDraft && key in this.plugin.signInDraft) this.plugin.signInDraft[key] = value;
          await this.plugin.saveSettings();
          if (wasSignedIn && !auth.matchesSettings()) this.plugin.signOut();
        });
      });
  }

  text(el, name, desc, key, after) {
    new Setting(el)
      .setName(name)
      .setDesc(desc)
      .addText((text) =>
        text.setValue(this.plugin.settings[key]).onChange(async (value) => {
          this.plugin.settings[key] = value;
          await this.plugin.saveSettings();
          if (after) after();
        })
      );
  }
}

// ── days and times ───────────────────────────────────────────────────────────

const pad = (n) => String(n).padStart(2, "0");

function isoDay(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function localToday() {
  return isoDay(new Date());
}

function parseDay(day) {
  const [year, month, date] = day.split("-").map(Number);
  return new Date(year, month - 1, date);
}

function addDays(day, n) {
  const date = parseDay(day);
  date.setDate(date.getDate() + n);
  return isoDay(date);
}

// A local day as [start, end) in milliseconds. Built from calendar dates, not
// start + 24 h, so days that change to or from summer time come out right.
function dayBounds(day) {
  return { start: parseDay(day).getTime(), end: parseDay(addDays(day, 1)).getTime() };
}

function longDate(day) {
  const date = parseDay(day);
  const options = { weekday: "short", day: "numeric", month: "short" };
  if (date.getFullYear() !== new Date().getFullYear()) options.year = "numeric";
  return date.toLocaleDateString([], options);
}

// The events that fall on a day, all-day ones first, then by start.
function dayEvents(events, day) {
  const { start, end } = dayBounds(day);
  return (events || [])
    .filter((e) =>
      e.allDay
        ? e.startDate <= day && day < e.endDate
        : e.end === e.start
          ? e.start >= start && e.start < end
          : e.start < end && e.end > start
    )
    .sort((a, b) => b.allDay - a.allDay || (a.allDay ? 0 : a.start - b.start) || a.title.localeCompare(b.title));
}

// A Graph date-time requested in UTC, as milliseconds.
function utcMillis(value) {
  if (!value || !value.dateTime) return NaN;
  const raw = value.dateTime.replace(/(\.\d{3})\d*$/, "$1");
  const zone = (value.timeZone || "UTC").toUpperCase();
  return new Date(zone === "UTC" || zone === "Z" ? `${raw}Z` : raw).getTime();
}

// The calendar date an all-day event's start or end stands for. Graph sends
// all-day events as midnight with no real time zone, so a bare midnight is
// taken as that date. Any other time is an instant whose local date counts,
// in case a time zone shifted it.
function allDayDate(value) {
  if (!value || !value.dateTime) return null;
  if (/T00:00:00(?:\.0+)?$/.test(value.dateTime)) return value.dateTime.slice(0, 10);
  const ms = utcMillis(value);
  return Number.isNaN(ms) ? value.dateTime.slice(0, 10) : isoDay(new Date(ms));
}

function safeUrl(url) {
  return typeof url === "string" && /^https:\/\//i.test(url) ? url : "";
}

function eventView(event, calendar) {
  const allDay = Boolean(event.isAllDay);
  const view = {
    id: event.id,
    title: (event.subject || "").trim() || "(No title)",
    allDay,
    cancelled: Boolean(event.isCancelled),
    showAs: event.showAs || null,
    response: (event.responseStatus && event.responseStatus.response) || null,
    location: (event.location && event.location.displayName) || "",
    organizer: (event.organizer && event.organizer.emailAddress && event.organizer.emailAddress.name) || "",
    preview: (event.bodyPreview || "").trim(),
    webLink: safeUrl(event.webLink),
    calendar: (calendar && calendar.name) || "",
    color: (calendar && calendar.color) || "",
    joinUrl: safeUrl((event.onlineMeeting && event.onlineMeeting.joinUrl) || event.onlineMeetingUrl),
  };
  if (allDay) {
    view.startDate = allDayDate(event.start);
    view.endDate = allDayDate(event.end) || addDays(view.startDate, 1);
    if (view.endDate <= view.startDate) view.endDate = addDays(view.startDate, 1);
  } else {
    view.start = utcMillis(event.start);
    view.end = utcMillis(event.end);
    if (Number.isNaN(view.end)) view.end = view.start;
  }
  return view;
}

// The events of a local day from the chosen calendars (the default one when
// none are chosen), as { events, failures }. A calendar that cannot be read
// is listed in failures and the others still show; when none can be read,
// the first error is thrown.
async function fetchDayEvents(auth, day, calendars = []) {
  const { start, end } = dayBounds(day);
  const params = [
    `startDateTime=${encodeURIComponent(new Date(start).toISOString())}`,
    `endDateTime=${encodeURIComponent(new Date(end).toISOString())}`,
    `$select=${EVENT_FIELDS}`,
    "$orderby=start/dateTime",
    "$top=100",
  ].join("&");
  // The UTC window can catch a neighbouring day's all-day events; keep the
  // ones on this day.
  const onDay = (items, calendar) =>
    dayEvents(items.map((item) => eventView(item, calendar)).filter((e) => (e.allDay ? e.startDate : !Number.isNaN(e.start))), day);
  if (!calendars.length) {
    return { events: onDay(await graphGetAll(auth, `${GRAPH}/me/calendarView?${params}`), null), failures: [] };
  }
  // One token up front, so parallel requests do not each try to refresh.
  await auth.accessToken();
  const results = await mapLimit(calendars, GRAPH_MAX_CONCURRENT, (calendar) =>
    graphGetAll(auth, `${GRAPH}/me/calendars/${encodeURIComponent(calendar.id)}/calendarView?${params}`).then(
      (items) => ({ events: onDay(items, calendar) }),
      (error) => ({ error })
    )
  );
  const failed = results.filter((r) => r.error);
  if (failed.length === results.length) throw failed[0].error;
  return {
    events: dayEvents(results.flatMap((r) => r.events || []), day),
    failures: results.flatMap((r, i) => (r.error ? [{ name: calendars[i].name, message: r.error.message }] : [])),
  };
}

// The user's calendars: their own and the ones shared with them, default
// first, as { id, name, color, owner, isDefault }.
async function fetchCalendars(auth) {
  const items = await graphGetAll(auth, `${GRAPH}/me/calendars?$select=id,name,color,hexColor,isDefaultCalendar,owner,canEdit&$top=100`);
  const me = (auth.account || "").toLowerCase();
  return items
    .map((c) => {
      const owner = c.owner && (c.owner.address || "").toLowerCase() !== me ? c.owner.name || c.owner.address || "" : "";
      return {
        id: c.id,
        name: c.name || "(No name)",
        color: calendarColor(c),
        owner: owner ? `Shared by ${owner}` : "",
        isDefault: Boolean(c.isDefaultCalendar),
      };
    })
    .sort((a, b) => b.isDefault - a.isDefault);
}

// Outlook's colour for a calendar: its hex colour, or one close to its
// named colour, or none (the theme's accent).
const CALENDAR_COLORS = {
  lightBlue: "#4f9fe6",
  lightGreen: "#5cb85c",
  lightOrange: "#f0a35e",
  lightGray: "#a0a0a0",
  lightYellow: "#e6c84f",
  lightTeal: "#4cc2c4",
  lightPink: "#e57eb5",
  lightBrown: "#b08b6e",
  lightRed: "#e06666",
};

function calendarColor(calendar) {
  if (/^#[0-9a-f]{6}$/i.test(calendar.hexColor || "")) return calendar.hexColor;
  return CALENDAR_COLORS[calendar.color] || "";
}

// Sets CSS custom properties, through Obsidian's helper where it exists.
function setCssProps(el, props) {
  if (typeof el.setCssProps === "function") el.setCssProps(props);
  else if (el.style && el.style.setProperty) for (const [key, value] of Object.entries(props)) el.style.setProperty(key, value);
}

// Runs fn over items with at most `limit` in flight, keeping results in order.
async function mapLimit(items, limit, fn) {
  const results = new Array(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i], i);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

// ── small helpers ────────────────────────────────────────────────────────────

function fill(template, values) {
  return (template || "").replace(/\{\{\s*(\w+)\s*\}\}/g, (whole, key) =>
    Object.prototype.hasOwnProperty.call(values, key) ? values[key] : whole
  );
}

// A short, stable lowercase hex hash (FNV-1a), for IDs derived from names.
function hashString(text) {
  let hash = 0x811c9dc5;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

// Adds text (one line or several) as lines of their own at a line: on that
// line when it is blank, otherwise on a new line below it, with the same
// indentation.
function insertOnOwnLine(editor, text, line = editor.getCursor().line) {
  const current = editor.getLine(line);
  const indent = current.match(/^[ \t]*/)[0];
  const blank = !current.trim() || isBareListMarker(current);
  const lines = text.split("\n");
  const indented = lines.map((l) => indent + l).join("\n");
  const target = (blank ? line : line + 1) + lines.length - 1;
  editor.transaction({
    changes: [
      blank
        ? { from: { line, ch: 0 }, to: { line, ch: current.length }, text: indented }
        : { from: { line, ch: current.length }, to: { line, ch: current.length }, text: `\n${indented}` },
    ],
  });
  editor.setCursor({ line: target, ch: indent.length + lines[lines.length - 1].length });
}

function isBareListMarker(line) {
  return /^\s*(?:[-*+]|\d+[.)])\s*(?:\[.\]\s*)?$/.test(line);
}

// After a drop, the editor selects the text it inserted. If that text landed
// inside a line with other text, move it to a line of its own below that line
// (above it when dropped at the line's start), with the same indentation.
function placeOnOwnLine(editor, text) {
  if (!text || text.includes("\n")) return;
  const selection = editor.listSelections()[0];
  if (!selection) return;
  const [from, to] = [selection.anchor, selection.head].sort((a, b) => a.line - b.line || a.ch - b.ch);
  if (from.line !== to.line || editor.getRange(from, to) !== text) return;
  const line = editor.getLine(from.line);
  const before = line.slice(0, from.ch);
  const after = line.slice(to.ch);
  if (!before.trim() && !after.trim()) return;
  const remaining = before + after;
  const indent = remaining.match(/^[ \t]*/)[0];
  if (isBareListMarker(remaining)) {
    editor.transaction({ changes: [{ from: { line: from.line, ch: 0 }, to: { line: from.line, ch: line.length }, text: indent + text }] });
    editor.setSelection({ line: from.line, ch: indent.length }, { line: from.line, ch: indent.length + text.length });
    return;
  }
  const atStart = !before.trim();
  const target = atStart ? from.line : from.line + 1;
  editor.transaction({
    changes: [
      {
        from: { line: from.line, ch: 0 },
        to: { line: from.line, ch: line.length },
        text: atStart ? `${indent}${text}\n${remaining}` : `${remaining}\n${indent}${text}`,
      },
    ],
  });
  editor.setSelection({ line: target, ch: indent.length }, { line: target, ch: indent.length + text.length });
}

// ── matching note lines to events ────────────────────────────────────────────
//
// One rule, shared by the links beside note lines and the panel's "in the
// note" marks, so the two always agree. A line holds an event when, trimmed,
// it is exactly what adding the event would insert; or when it is a list item
// whose text, normalized (see normalizeText), is the event's time and title,
// as "09:00–10:00 Standup" or "09:00 Standup". A title alone is not enough:
// "- Lunch" is too common to be taken for an event.

// A list item, with its checkbox mark (if any) and its text.
const LIST_ITEM = /^(?:[-*+]|\d+[.)])\s+(?:\[(.)\]\s+)?(.*)$/;

// What makes two event texts "the same": case, runs of whitespace and the
// kind of dash do not count, and neither does a tail of emojis, dates, tags
// and links, such as "📅 2026-11-20", "#meeting", "[Join](https://…)" or a
// bare https:// link.
const TRAILING_NOISE = new RegExp(
  "(?:\\s*(?:" +
    // Each token ends where its kind of text ends, so a run such as
    // "12.12.1212.12.12" splits one way only; with overlapping splits, a
    // line of a few hundred digits would backtrack for minutes.
    "\\[[^\\]]*\\]\\([^)]*\\)" + // Markdown link
    "|https?:\\/\\/\\S+(?!\\S)" + // bare link
    "|#[\\p{L}\\p{N}_/-]+(?![\\p{L}\\p{N}_/-])" + // tag
    "|\\d{4}-\\d{2}-\\d{2}(?:[ T]\\d{1,2}:\\d{2})?(?!\\d)" + // 2026-11-20, optional time
    "|\\d{1,2}[./-]\\d{1,2}[./-]\\d{2,4}(?!\\d)" + // 20/11/2026, 20.11.26
    // One symbol per repetition: a "+" here as well would nest two
    // quantifiers and make long emoji runs backtrack for seconds.
    "|[\\p{Extended_Pictographic}\\u2190-\\u21FF\\u2B00-\\u2BFF\\uFE0F\\u200D]" + // emojis and arrows
    "))+\\s*$",
  "u"
);

function normalizeText(text) {
  return (text || "")
    .replace(TRAILING_NOISE, "")
    .replace(/[‒-―−]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

// Longer lines are never event lines, and are not worth normalizing.
const MAX_EVENT_LINE = 1000;

// The keys a note line can match an event by: "=" and the trimmed line, and
// for a list item "~" and its normalized text. Empty for blank or long lines.
function lineKeys(text) {
  if (!text || text.length > MAX_EVENT_LINE) return [];
  const trimmed = text.trim();
  if (!trimmed) return [];
  const keys = [`=${trimmed}`];
  const item = trimmed.match(LIST_ITEM);
  const normalized = item && normalizeText(item[2]);
  if (normalized) keys.push(`~${normalized}`);
  return keys;
}

// The keys an event can be found by: its line as inserted (the first line,
// for a template of several), and its time and title, with the time as the
// panel shows it or as the start alone.
function eventKeys(rendered, title, time, start) {
  const first = (rendered || "").split("\n").map((l) => l.trim()).find(Boolean);
  const keys = first ? [`=${first}`] : [];
  for (const prefix of [time, start]) {
    const normalized = prefix && normalizeText(`${prefix} ${title}`);
    if (normalized && !keys.includes(`~${normalized}`)) keys.push(`~${normalized}`);
  }
  return keys;
}

// Every key of every line of a note, for the panel's "in the note" marks.
function noteLineKeys(text) {
  const keys = new Set();
  for (const line of (text || "").split("\n")) for (const key of lineKeys(line)) keys.add(key);
  return keys;
}

// The fetched days' events keyed like note lines: each key maps to the
// { id, day } of every shown event it stands for, in day order.
function buildEventIndex(days, plugin) {
  const index = new Map();
  for (const day of [...days.keys()].sort()) {
    for (const event of plugin.visibleEvents(dayEvents(days.get(day).events, day))) {
      for (const key of plugin.eventKeys(event, day)) {
        const list = index.get(key) || [];
        if (!list.some((c) => c.id === event.id && c.day === day)) list.push({ id: event.id, day });
        index.set(key, list);
      }
    }
  }
  return index.size ? index : null;
}

// The index key a note line matches, or null: its exact line first, then
// its normalized text.
function eventKeyForLine(index, text) {
  for (const key of lineKeys(text)) if (index.has(key)) return key;
  return null;
}

// The editor side of the links: a small button at the end of each line that
// holds a panel event. Only the lines on screen are looked at, and only when
// the text, the visible part or the events change; moving the cursor costs
// nothing. Each distinct line text is looked up once, so scrolling back and
// typing elsewhere reuse earlier answers.
function eventLinkExtension(plugin) {
  const redrawEffect = StateEffect.define();
  const decorations = new Map(); // line text → decoration or null
  let cachedIndex = null;

  // Holds only the matched key: which event it shows is decided on click,
  // against the events then, and the note's own day.
  class EventLinkWidget extends WidgetType {
    constructor(key) {
      super();
      this.key = key;
    }
    eq(other) {
      return other.key === this.key;
    }
    toDOM(view) {
      const el = document.createElement("span");
      el.className = "oc-reveal";
      el.setAttribute("role", "button");
      el.setAttribute("aria-label", "Show in Outlook calendar panel");
      // Keeps the cursor and focus where they are.
      el.addEventListener("mousedown", (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
      });
      el.addEventListener("click", (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
        const info = editorInfoField && view.state.field ? view.state.field(editorInfoField, false) : null;
        const noteDay = plugin.dailyNoteDate(info && info.file);
        plugin.revealEvent(this.key, noteDay).catch((e) => new Notice(`Could not show the event: ${e.message}`));
      });
      return el;
    }
    // The editor leaves clicks on the button alone.
    ignoreEvent() {
      return true;
    }
  }

  const decorationFor = (index, text) => {
    if (index !== cachedIndex || decorations.size > 2000) {
      decorations.clear();
      cachedIndex = index;
    }
    let deco = decorations.get(text);
    if (deco === undefined) {
      const key = eventKeyForLine(index, text);
      deco = key ? Decoration.widget({ widget: new EventLinkWidget(key), side: 1 }) : null;
      decorations.set(text, deco);
    }
    return deco;
  };

  return ViewPlugin.fromClass(
    class {
      constructor(view) {
        this.view = view;
        this.decorations = this.build();
        plugin.eventLinkViews.add(this);
      }
      update(update) {
        if (update.docChanged || update.viewportChanged || update.transactions.some((tr) => tr.effects.some((e) => e.is(redrawEffect)))) {
          this.decorations = this.build();
        }
      }
      redraw() {
        this.view.dispatch({ effects: redrawEffect.of(null) });
      }
      destroy() {
        plugin.eventLinkViews.delete(this);
      }
      build() {
        const index = plugin.eventIndexNow();
        if (!index) return Decoration.none;
        const doc = this.view.state.doc;
        const builder = new RangeSetBuilder();
        let last = 0;
        for (const { from, to } of this.view.visibleRanges) {
          let line = doc.lineAt(from);
          // Ranges can share a line; each line is decorated once, in order.
          if (line.number <= last) {
            if (last >= doc.lines) break;
            line = doc.line(last + 1);
          }
          while (line.from <= to) {
            const deco = decorationFor(index, line.text);
            if (deco) builder.add(line.to, line.to, deco);
            last = line.number;
            if (line.number === doc.lines) break;
            line = doc.line(line.number + 1);
          }
        }
        return builder.finish();
      }
    },
    { decorations: (links) => links.decorations }
  );
}

// ── Microsoft sign-in ────────────────────────────────────────────────────────
//
// OAuth 2.0 device code flow against the Microsoft identity platform, with no
// library. Requests go through Obsidian's requestUrl, which is not subject to
// browser CORS rules.

class NotSignedInError extends Error {}

class MicrosoftAuth {
  constructor(plugin) {
    this.plugin = plugin;
    this.refreshing = null;
    // Bumped whenever a session ends or begins, so responses that belong to
    // an older session are thrown away.
    this.epoch = 0;
    this.secretId = `${SESSION_SECRET_PREFIX}-${hashString(plugin.app.vault.getName())}`;
    this._session = undefined;
  }

  // Read from secret storage on first use.
  get session() {
    if (this._session === undefined) this._session = this.loadSession();
    return this._session;
  }

  get settings() {
    return this.plugin.settings;
  }

  get signedIn() {
    return Boolean(this.session && this.session.refresh_token && this.matchesSettings());
  }

  get account() {
    return this.session ? this.session.account || "" : "";
  }

  matchesSettings() {
    return (
      Boolean(this.session) &&
      this.session.client_id === this.settings.clientId.trim() &&
      this.session.tenant === this.tenant()
    );
  }

  tenant() {
    return (this.settings.tenant || "").trim() || "common";
  }

  context() {
    return { clientId: this.settings.clientId.trim(), tenant: this.tenant(), epoch: this.epoch };
  }

  endpoint(path, tenant = this.tenant()) {
    return `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/${path}`;
  }

  scope() {
    const asked = (this.settings.scopes || "").split(/[\s,]+/).filter(Boolean);
    return [...new Set([...asked, "offline_access", "openid", "profile"])].join(" ");
  }

  loadSession() {
    try {
      const raw = this.plugin.app.secretStorage.getSecret(this.secretId);
      return raw ? JSON.parse(raw) : null;
    } catch (_) {
      return null;
    }
  }

  saveSession(session) {
    this._session = session;
    // Secret storage has no delete; an empty value reads back as signed out.
    this.plugin.app.secretStorage.setSecret(this.secretId, session ? JSON.stringify(session) : "");
  }

  signOut() {
    this.epoch++;
    this.saveSession(null);
  }

  async startDeviceFlow() {
    const ctx = this.context();
    const res = await postForm(this.endpoint("devicecode", ctx.tenant), { client_id: ctx.clientId, scope: this.scope() });
    if (res.status !== 200) throw new Error(describeAuthError(res.json));
    return { ...res.json, ctx };
  }

  // Polls until the user finishes signing in, or `cancel.cancelled` is set.
  async completeDeviceFlow(flow, cancel) {
    const ctx = flow.ctx || this.context();
    let interval = Math.max(1, Number(flow.interval) || 5);
    const deadline = Date.now() + (Number(flow.expires_in) || 900) * 1000;
    while (!cancel.cancelled) {
      await sleep(interval * 1000);
      if (cancel.cancelled) break;
      if (Date.now() > deadline) throw new Error("The sign-in code expired. Start again.");
      const res = await postForm(this.endpoint("token", ctx.tenant), {
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: ctx.clientId,
        device_code: flow.device_code,
      });
      if (cancel.cancelled) break;
      if (res.status === 200) {
        this.epoch++;
        this.storeTokens(res.json, ctx);
        return;
      }
      const error = res.json && res.json.error;
      if (error === "authorization_pending") continue;
      if (error === "slow_down") {
        interval += 5;
        continue;
      }
      if (error === "authorization_declined") throw new Error("Sign-in was declined.");
      if (error === "expired_token") throw new Error("The sign-in code expired. Start again.");
      throw new Error(describeAuthError(res.json));
    }
    throw new Error("Sign-in cancelled.");
  }

  storeTokens(tokens, ctx) {
    const previous = this.session || {};
    const claims = jwtClaims(tokens.id_token);
    this.saveSession({
      client_id: ctx.clientId,
      tenant: ctx.tenant,
      access_token: tokens.access_token,
      // Microsoft rotates refresh tokens; keep the newest one.
      refresh_token: tokens.refresh_token || previous.refresh_token,
      expires_at: Date.now() + (Number(tokens.expires_in) || 3600) * 1000,
      account: (claims && (claims.preferred_username || claims.email || claims.name)) || previous.account || "",
    });
  }

  async accessToken(force = false) {
    if (!this.signedIn) throw new NotSignedInError("Not signed in to Microsoft.");
    const fresh = this.session.access_token && this.session.expires_at - 120000 > Date.now();
    if (fresh && !force) return this.session.access_token;
    // Concurrent callers share one refresh, so a rotated refresh token is
    // never spent twice.
    if (!this.refreshing) {
      this.refreshing = this.refresh().finally(() => {
        this.refreshing = null;
      });
    }
    await this.refreshing;
    if (!this.signedIn) throw new NotSignedInError("Not signed in to Microsoft.");
    return this.session.access_token;
  }

  async refresh() {
    const ctx = { clientId: this.session.client_id, tenant: this.session.tenant, epoch: this.epoch };
    const res = await postForm(this.endpoint("token", ctx.tenant), {
      grant_type: "refresh_token",
      client_id: ctx.clientId,
      refresh_token: this.session.refresh_token,
      scope: this.scope(),
    });
    if (ctx.epoch !== this.epoch) throw new NotSignedInError("Not signed in to Microsoft.");
    if (res.status === 200) {
      this.storeTokens(res.json, ctx);
      return;
    }
    const error = res.json && res.json.error;
    if (error === "invalid_grant" || error === "interaction_required" || error === "invalid_client") {
      this.signOut();
      this.plugin.sessionEnded();
      const expired = new NotSignedInError(`Your Microsoft sign-in has expired. Sign in again. (${describeAuthError(res.json)})`);
      expired.expired = true;
      throw expired;
    }
    throw new Error(describeAuthError(res.json));
  }
}

class SignInModal extends Modal {
  constructor(app, plugin) {
    super(app);
    this.plugin = plugin;
    this.cancel = { cancelled: false };
  }

  async onOpen() {
    this.setTitle("Sign in to Microsoft");
    // Signing in means leaving for a browser and coming back. A click on
    // the dimmed background on return would close the dialog and lose what
    // was typed, so only Escape and the close button close it.
    const bg = this.containerEl.querySelector(".modal-bg");
    if (bg) bg.addEventListener("click", (evt) => evt.stopImmediatePropagation(), true);
    this.askForApp();
  }

  // The app registration's IDs, filled in with what was entered last, so
  // the order they are typed in never matters and a closed dialog loses
  // nothing.
  askForApp() {
    const { contentEl } = this;
    const s = this.plugin.settings;
    const draft = this.plugin.signInDraft || (this.plugin.signInDraft = { clientId: s.clientId.trim(), tenant: this.plugin.auth.tenant() });
    let started = false;
    contentEl.createDiv({
      cls: "oc-signin-status",
      text: "Enter the IDs of the app you registered in Microsoft Entra. The plugin's README explains how to register one.",
    });
    new Setting(contentEl)
      .setName("Client ID")
      .setDesc("Application (client) ID")
      .addText((text) =>
        text
          .setPlaceholder("00000000-0000-0000-0000-000000000000")
          .setValue(draft.clientId)
          .onChange((v) => (draft.clientId = v.trim()))
      );
    new Setting(contentEl)
      .setName("Tenant")
      .setDesc("Directory (tenant) ID for an app limited to your organisation; otherwise common or consumers.")
      .addText((text) => text.setPlaceholder("common").setValue(draft.tenant).onChange((v) => (draft.tenant = v.trim())));
    new Setting(contentEl).addButton((button) =>
      button
        .setButtonText("Continue")
        .setCta()
        .onClick(async () => {
          if (started) return;
          if (!draft.clientId) {
            new Notice("Enter the client ID.");
            return;
          }
          started = true;
          s.clientId = draft.clientId;
          s.tenant = draft.tenant || "common";
          try {
            await this.plugin.saveSettings();
          } catch (e) {
            started = false;
            new Notice(`Could not save the IDs: ${e.message}`);
            return;
          }
          this.plugin.refreshSettingTab();
          contentEl.empty();
          this.startSignIn();
        })
    );
  }

  async startSignIn() {
    const { contentEl } = this;
    const status = contentEl.createDiv({ cls: "oc-signin-status", text: "Asking Microsoft for a sign-in code…" });
    let flow;
    try {
      flow = await this.plugin.auth.startDeviceFlow();
    } catch (e) {
      if (!this.cancel.cancelled) this.failed(status, `Could not start sign-in: ${e.message}`);
      return;
    }
    if (this.cancel.cancelled) return;

    status.setText("Open the Microsoft sign-in page, enter this code and sign in. This window closes by itself when you are done.");
    const codeRow = contentEl.createDiv({ cls: "oc-signin-code" });
    codeRow.createEl("code", { text: flow.user_code });
    const actions = contentEl.createDiv({ cls: "oc-signin-actions" });
    const copy = actions.createEl("button", { text: "Copy code" });
    copy.addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(flow.user_code);
        copy.setText("Copied");
      } catch (_) {
        new Notice("Could not copy; select the code instead.");
      }
    });
    const open = actions.createEl("button", { cls: "mod-cta", text: "Open sign-in page" });
    open.addEventListener("click", () => window.open(flow.verification_uri, "_blank"));
    contentEl.createDiv({ cls: "oc-signin-hint", text: flow.verification_uri });

    try {
      await this.plugin.auth.completeDeviceFlow(flow, this.cancel);
      if (this.cancel.cancelled) return;
      this.plugin.signInDraft = null;
      this.close();
      this.plugin.signedIn();
    } catch (e) {
      if (this.cancel.cancelled) return;
      codeRow.remove();
      actions.remove();
      this.failed(status, e.message);
    }
  }

  // A failed sign-in can go back to the IDs, which are often what was wrong.
  failed(status, message) {
    status.setText(message);
    status.addClass("oc-error");
    new Setting(this.contentEl).addButton((button) =>
      button.setButtonText("Back").onClick(() => {
        this.contentEl.empty();
        this.askForApp();
      })
    );
  }

  onClose() {
    this.cancel.cancelled = true;
    this.contentEl.empty();
    if (this.plugin.signInModal === this) this.plugin.signInModal = null;
  }
}

// ── Microsoft Graph ──────────────────────────────────────────────────────────

const GRAPH_MAX_RETRIES = 5;
const GRAPH_MAX_CONCURRENT = 3;

async function graphGet(auth, url) {
  let token = await auth.accessToken();
  let retriedAuth = false;
  for (let attempt = 0; ; attempt++) {
    const res = await requestUrl({
      url,
      // Times come back in UTC; the plugin converts them to local time.
      headers: { Authorization: `Bearer ${token}`, Prefer: 'outlook.timezone="UTC"' },
      throw: false,
    });
    if (res.status === 401 && !retriedAuth) {
      retriedAuth = true;
      token = await auth.accessToken(true);
      continue;
    }
    if ([429, 503, 504].includes(res.status) && attempt < GRAPH_MAX_RETRIES) {
      await sleep(retryDelay(res.headers, attempt) * 1000);
      continue;
    }
    if (res.status >= 400) {
      const body = safeJson(res);
      const message = (body && body.error && body.error.message) || `HTTP ${res.status}`;
      if (res.status === 403) throw new Error(`Microsoft refused access to your calendar: ${message}`);
      throw new Error(`Microsoft Graph error: ${message}`);
    }
    return safeJson(res) || {};
  }
}

async function graphGetAll(auth, url) {
  const items = [];
  while (url) {
    const page = await graphGet(auth, url);
    items.push(...(page.value || []));
    url = page["@odata.nextLink"];
  }
  return items;
}

async function postForm(url, fields) {
  const res = await requestUrl({
    url,
    method: "POST",
    contentType: "application/x-www-form-urlencoded",
    body: new URLSearchParams(fields).toString(),
    throw: false,
  });
  return { status: res.status, headers: res.headers, json: safeJson(res) };
}

function safeJson(res) {
  try {
    return res.json;
  } catch (_) {
    return null;
  }
}

const AUTH_HINTS = {
  AADSTS50059: "Set Tenant to your directory (tenant) ID: this app accepts only accounts in its own organisation.",
  AADSTS700016: "Check the client ID and the tenant: Microsoft cannot find that app there.",
  AADSTS7000218: "In the app registration, turn on Authentication → Allow public client flows.",
  AADSTS65001: "The permission has not been granted yet. An administrator may need to approve the app.",
  AADSTS90002: "Check the tenant setting: Microsoft does not know that directory.",
};

function describeAuthError(body) {
  if (!body) return "No response from Microsoft.";
  const text = String(body.error_description || body.error || "Unknown error.");
  const message = text.split(/\r?\n/)[0].replace(/\s*Trace ID:.*$/, "");
  const code = (message.match(/AADSTS\d+/) || [])[0];
  return code && AUTH_HINTS[code] ? `${message} ${AUTH_HINTS[code]}` : message;
}

function jwtClaims(token) {
  try {
    const part = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    const json = decodeURIComponent(
      atob(part)
        .split("")
        .map((c) => "%" + c.charCodeAt(0).toString(16).padStart(2, "0"))
        .join("")
    );
    return JSON.parse(json);
  } catch (_) {
    return null;
  }
}

function retryDelay(headers, attempt) {
  const header = Object.entries(headers || {}).find(([key]) => key.toLowerCase() === "retry-after");
  const seconds = header ? parseInt(header[1], 10) : NaN;
  return Number.isFinite(seconds) && seconds > 0 ? Math.min(seconds, 60) : Math.min(60, 2 ** (attempt + 1));
}

function sleep(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}
