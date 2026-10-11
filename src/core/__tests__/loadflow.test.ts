import { describe, expect, it } from "vitest";
import { EXAMPLE_LIBRARY, prepareForSolve, runLoadFlow } from "..";
import { SAMPLE_PROJECT } from "../sample";

const reply = (status: number, body: string) =>
  (async () => new Response(body, { status })) as unknown as typeof fetch;

describe("runLoadFlow", () => {
  it("returns parsed results on success", async () => {
    const ok = JSON.stringify({ converged: true, buses: {}, lines: {}, transformers: {} });
    const r = await runLoadFlow(SAMPLE_PROJECT, EXAMPLE_LIBRARY, reply(200, ok));
    expect(r.converged).toBe(true);
  });
  it("surfaces the server's error message", async () => {
    await expect(
      runLoadFlow(SAMPLE_PROJECT, EXAMPLE_LIBRARY, reply(400, JSON.stringify({ error: "ValueError: boom" }))),
    ).rejects.toThrow("ValueError: boom");
  });
  it("explains non-JSON responses (e.g. a 404 page)", async () => {
    await expect(
      runLoadFlow(SAMPLE_PROJECT, EXAMPLE_LIBRARY, reply(404, "<html>Not found</html>")),
    ).rejects.toThrow(/HTTP 404/);
  });
});

describe("prepareForSolve", () => {
  it("accepts the sample plant", () => {
    expect(prepareForSolve(SAMPLE_PROJECT, EXAMPLE_LIBRARY).ok).toBe(true);
  });
  it("names the element and field for schema problems", () => {
    const bad = structuredClone(SAMPLE_PROJECT);
    bad.buses[1].nominalKv = -1;
    const r = prepareForSolve(bad, EXAMPLE_LIBRARY);
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.problems[0]).toContain("B2 › nominalKv");
  });
  it("blocks on validation errors", () => {
    const bad = structuredClone(SAMPLE_PROJECT);
    bad.sources = [];
    const r = prepareForSolve(bad, EXAMPLE_LIBRARY);
    expect(r.ok).toBe(false);
  });
});
