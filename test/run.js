/**
 * Tiny test runner — no dependencies, so `node test/run.js` is all it takes.
 * Each test file exports a function that receives { test, assert }.
 */

const results = { passed: 0, failed: 0 };
const failures = [];
let currentSuite = '';

function test(name, fn) {
  try {
    fn();
    results.passed++;
    console.log(`  ✓ ${name}`);
  } catch (err) {
    results.failed++;
    failures.push({ suite: currentSuite, name, err });
    console.log(`  ✗ ${name}`);
    console.log(`      ${err && err.message ? err.message : err}`);
  }
}

const assert = {
  ok(value, message) {
    if (!value) throw new Error(message || `Expected a truthy value, got ${JSON.stringify(value)}`);
  },
  equal(actual, expected, message) {
    if (actual !== expected) {
      throw new Error(message || `Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
    }
  },
  deepEqual(actual, expected, message) {
    const a = JSON.stringify(actual);
    const b = JSON.stringify(expected);
    if (a !== b) throw new Error(message || `Expected ${b}, got ${a}`);
  },
  /** The call must have failed, with a message matching `pattern`. */
  rejected(result, pattern, message) {
    if (result.ok) {
      throw new Error(message || `Expected the call to be refused, but it returned ${JSON.stringify(result.value)}`);
    }
    if (pattern && !pattern.test(result.error.message)) {
      throw new Error(message || `Expected an error matching ${pattern}, got "${result.error.message}"`);
    }
  },
  /** The call must have succeeded. */
  allowed(result, message) {
    if (!result.ok) {
      throw new Error(message || `Expected the call to succeed, but it threw "${result.error.message}"`);
    }
    return result.value;
  }
};

const suites = ['./authorization.test.js', './assignments.test.js', './ui_flows.test.js', './legacy_globals.test.js'];

suites.forEach(file => {
  const suite = require(file);
  currentSuite = suite.name || file;
  console.log(`\n${currentSuite}`);
  suite.run({ test, assert });
});

console.log(`\n${results.passed} passed, ${results.failed} failed`);
if (results.failed > 0) {
  console.log('\nFailures:');
  failures.forEach(f => console.log(`  ${f.suite} › ${f.name}\n    ${f.err.stack || f.err}`));
  process.exit(1);
}
