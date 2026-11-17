// A stand-in for @codemirror/view and @codemirror/state: just enough to run
// the plugin's editor extension against a fake view in Node.
class WidgetType {}

const Decoration = {
  none: [],
  widget: (spec) => ({ spec }),
};

class RangeSetBuilder {
  constructor() {
    this.ranges = [];
  }
  add(from, to, value) {
    const last = this.ranges[this.ranges.length - 1];
    if (last && from < last.from) throw new Error("Ranges must be added sorted");
    this.ranges.push({ from, to, value });
  }
  finish() {
    return this.ranges;
  }
}

const StateEffect = {
  define() {
    const type = {
      of: (value) => ({ value, is: (t) => t === type }),
    };
    return type;
  },
};

const ViewPlugin = {
  fromClass: (cls, spec) => ({ cls, spec }),
};

module.exports = { WidgetType, Decoration, RangeSetBuilder, StateEffect, ViewPlugin };
