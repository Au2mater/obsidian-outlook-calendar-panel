// A stand-in for the "obsidian" module, enough to load the plugin in Node.
// requestUrl answers through global.__handler and records every request;
// secret storage counts its reads, so tests can check startup touches none.
class Plugin {
  constructor() {
    const store = {};
    this.manifest = { id: "outlook-calendar-panel" };
    this.app = {
      secretStorage: {
        setSecret(id, value) {
          if (!/^[a-z0-9-]+$/.test(id)) throw new Error("Invalid secret ID");
          store[id] = value;
        },
        getSecret: (id) => {
          global.__secretReads = (global.__secretReads || 0) + 1;
          return store[id] ?? null;
        },
      },
      workspace: {
        on: () => ({}),
        getLeavesOfType: () => global.__leaves || [],
        getActiveFile: () => global.__activeFile || null,
        getActiveViewOfType: () => null,
        getMostRecentLeaf: () => null,
        setActiveLeaf() {},
        // Every waiting callback runs when a test calls global.__layoutReady().
        onLayoutReady: (fn) => {
          const earlier = global.__layoutReady;
          global.__layoutReady = () => {
            if (earlier) earlier();
            fn();
          };
        },
      },
      vault: { on: () => ({}), getName: () => "Test vault" },
      metadataCache: { getFileCache: () => global.__fileCache || null },
      secrets: store,
      local: {},
      loadLocalStorage(key) {
        return this.local[key] ?? null;
      },
      saveLocalStorage(key, value) {
        if (value === null) delete this.local[key];
        else this.local[key] = JSON.parse(JSON.stringify(value));
      },
    };
  }
  registerEvent() {}
  registerEditorExtension(extension) {
    global.__editorExtension = extension;
  }
  registerInterval(id) {
    clearInterval(id);
    return id;
  }
  registerView(type, factory) {
    global.__viewFactory = factory;
  }
  addRibbonIcon() {}
  addCommand() {}
  addSettingTab() {}
  async loadData() {
    return global.__data || {};
  }
  async saveData(data) {
    global.__saved = data;
  }
}

class Stub {}

const pad = (n) => String(n).padStart(2, "0");
// Enough of moment: format a time as HH:mm or a date as YYYY-MM-DD, and parse
// a YYYY-MM-DD file name strictly.
function moment(value, format) {
  if (typeof value === "string") {
    const pattern = { "YYYY-MM-DD": /^\d{4}-\d{2}-\d{2}$/, "YYYY/MM/DD": /^\d{4}\/\d{2}\/\d{2}$/ }[format];
    const ok = Boolean(pattern && pattern.test(value));
    return { isValid: () => ok, format: () => value.replace(/\//g, "-") };
  }
  const d = new Date(value);
  return {
    isValid: () => !Number.isNaN(d.getTime()),
    format: (f) =>
      f === "YYYY-MM-DD"
        ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`
        : `${pad(d.getHours())}:${pad(d.getMinutes())}`,
  };
}

module.exports = {
  ItemView: class {
    constructor(leaf) {
      this.leaf = leaf;
    }
  },
  MarkdownView: Stub,
  Modal: Stub,
  Notice: class {
    constructor(message) {
      (global.__notices = global.__notices || []).push(message);
    }
  },
  Plugin,
  PluginSettingTab: Stub,
  Setting: Stub,
  // What the editor knows about its note; the fake views below carry the file.
  editorInfoField: { name: "editorInfoField" },
  setIcon() {},
  moment,
  async requestUrl(req) {
    global.__requests.push(req);
    const res = await global.__handler(req);
    return {
      status: res.status || 200,
      headers: res.headers || {},
      get json() {
        if (res.body === undefined) throw new Error("Not JSON");
        return res.body;
      },
    };
  },
};
