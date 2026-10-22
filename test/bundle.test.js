import { test } from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";

import { bundle } from "../tools/bundle.js";

// Excel for Mac runs the custom-functions script as a classic script, so the
// bundle must evaluate with no module syntax and register every TB.* function.
test("bundle: functions.js flattens to a classic script that registers every function", async () => {
  const code = bundle("src/functions/functions.js");
  assert.doesNotMatch(code, /^\s*(import|export)\s/m);

  const registered = {};
  class CfError extends Error {
    constructor(code, message) {
      super(message);
      this.code = code;
    }
  }
  const context = vm.createContext({
    CustomFunctions: { associate: (n, f) => (registered[n] = f), Error: CfError, ErrorCode: { invalidValue: "#VALUE!" } },
    console,
    URL,
    URLSearchParams,
    AbortController,
    setTimeout,
    clearTimeout,
  });
  vm.runInContext(code, context); // classic-script evaluation: a stray `import` would throw here

  assert.deepEqual(Object.keys(registered).sort(), ["DATASETS", "HISTORY", "SAMPLE", "TAGS", "VALUE"]);
  await assert.rejects(() => registered.DATASETS(), (e) => e.code === "#VALUE!" && /address is not set/.test(e.message));
});
