import { describe, expect, it } from "vitest";
import {
  EXAMPLE_LIBRARY, emptyProject, findIslands, hasErrors, parseProject,
  serializeProject, validateProject, type Project,
} from "..";
import { SAMPLE_PROJECT } from "../sample";

const clone = (p: Project): Project => structuredClone(p);
const codes = (p: Project) => validateProject(p, EXAMPLE_LIBRARY).map((i) => i.code);

describe("io", () => {
  it("round-trips a project", () => {
    const result = parseProject(serializeProject(SAMPLE_PROJECT));
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.project).toEqual(SAMPLE_PROJECT);
  });
  it("applies defaults to an empty project", () => {
    expect(emptyProject().frequencyHz).toBe(50);
  });
  it("reports bad JSON and schema errors", () => {
    expect(parseProject("{").ok).toBe(false);
    const bad = parseProject(JSON.stringify({ schemaVersion: 1, buses: [{ id: "x", nominalKv: -1 }] }));
    expect(bad.ok).toBe(false);
    if (!bad.ok) expect(bad.errors[0]).toContain("buses.0.nominalKv");
  });
  it("rejects unknown schema versions", () => {
    expect(parseProject(JSON.stringify({ schemaVersion: 99 })).ok).toBe(false);
  });
});

describe("topology", () => {
  it("sample plant is one energised island", () => {
    const islands = findIslands(SAMPLE_PROJECT);
    expect(islands).toHaveLength(1);
    expect(islands[0].energised).toBe(true);
  });
  it("an open element switch de-energises the LV side", () => {
    const p = clone(SAMPLE_PROJECT);
    const sw = p.switches[0];
    if (sw.kind === "element") sw.closed = false;
    const islands = findIslands(p);
    expect(islands).toHaveLength(2);
    expect(islands.filter((i) => !i.energised)[0].buses.sort()).toEqual(["B2", "B3"]);
  });
  it("a closed coupler joins buses, an open one does not", () => {
    const p = clone(SAMPLE_PROJECT);
    p.lines = [];
    expect(findIslands(p)).toHaveLength(2);
    p.switches.push({ kind: "coupler", id: "BC1", name: "", busA: "B2", busB: "B3", closed: true });
    expect(findIslands(p)).toHaveLength(1);
  });
});

describe("validation", () => {
  it("sample plant has no errors", () => {
    expect(hasErrors(validateProject(SAMPLE_PROJECT, EXAMPLE_LIBRARY))).toBe(false);
  });
  it("flags duplicate ids and missing bus refs", () => {
    const p = clone(SAMPLE_PROJECT);
    p.loads[0].id = "B1";
    p.loads[0].bus = "NOPE";
    expect(codes(p)).toEqual(expect.arrayContaining(["DUPLICATE_ID", "MISSING_BUS_REF"]));
  });
  it("flags a line across voltage levels", () => {
    const p = clone(SAMPLE_PROJECT);
    p.lines[0].toBus = "B1";
    expect(codes(p)).toContain("LINE_VOLTAGE_MISMATCH");
  });
  it("flags transformer rating vs bus voltage", () => {
    const p = clone(SAMPLE_PROJECT);
    p.transformers[0].ratedLvKv = 0.69;
    expect(codes(p)).toContain("TRANSFORMER_VOLTAGE_MISMATCH");
  });
  it("flags unknown cable type and missing impedance", () => {
    const p = clone(SAMPLE_PROJECT);
    p.lines[0].cableTypeId = "missing";
    expect(codes(p)).toContain("UNKNOWN_CABLE_TYPE");
    p.lines[0].cableTypeId = undefined;
    expect(codes(p)).toContain("LINE_NO_IMPEDANCE");
  });
  it("flags no source, isolated bus and unpowered island", () => {
    const p = clone(SAMPLE_PROJECT);
    p.sources = [];
    expect(codes(p)).toContain("NO_SOURCE");
    const q = clone(SAMPLE_PROJECT);
    q.buses.push({ id: "B9", name: "", nominalKv: 0.4 });
    expect(codes(q)).toContain("ISOLATED_BUS");
    const r = clone(SAMPLE_PROJECT);
    const sw = r.switches[0];
    if (sw.kind === "element") sw.closed = false;
    expect(codes(r)).toContain("ISLAND_NO_SOURCE");
  });
});
