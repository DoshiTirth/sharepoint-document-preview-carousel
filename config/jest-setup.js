// jsdom (this project's Jest test environment) doesn't expose
// `structuredClone` on its global scope, even though every real browser has
// supported it natively since 2022. fake-indexeddb (used in cache tests)
// relies on it being globally available, matching real browser behavior -
// so we polyfill it here using Node's v8 serialize/deserialize, which is
// the standard workaround for environments missing the native global.
if (typeof globalThis.structuredClone === 'undefined') {
  const v8 = require('node:v8');
  globalThis.structuredClone = (value) => v8.deserialize(v8.serialize(value));
}
