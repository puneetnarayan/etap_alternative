import { readFileSync, writeFileSync } from "node:fs";
import { expect, it } from "vitest";
import { EXAMPLE_LIBRARY } from "../library";
import { SAMPLE_PROJECT } from "../sample";

// Keeps solver/sample.json (used by the Python solver tests) in sync with the TS sample.
// Regenerate with: UPDATE_FIXTURE=1 npm test
const path = new URL("../../../solver/sample.json", import.meta.url);
const expected = JSON.stringify({ project: SAMPLE_PROJECT, library: EXAMPLE_LIBRARY }, null, 2) + "\n";

it("solver/sample.json matches the TS sample", () => {
  if (process.env.UPDATE_FIXTURE) writeFileSync(path, expected);
  expect(readFileSync(path, "utf8")).toBe(expected);
});
