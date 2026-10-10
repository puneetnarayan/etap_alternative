import { ProjectSchema, SCHEMA_VERSION, type Project } from "./schema";

/** Small plant: 11 kV utility -> 1 MVA transformer -> 0.4 kV bus with a cable-fed motor and load. */
export const SAMPLE_PROJECT: Project = ProjectSchema.parse({
  schemaVersion: SCHEMA_VERSION,
  name: "Sample plant",
  frequencyHz: 50,
  buses: [
    { id: "B1", name: "11 kV incomer", nominalKv: 11 },
    { id: "B2", name: "LV main", nominalKv: 0.4 },
    { id: "B3", name: "Motor control centre", nominalKv: 0.4 },
  ],
  sources: [{ id: "S1", name: "Utility", bus: "B1", shortCircuitMva: 250 }],
  transformers: [
    {
      id: "T1", name: "TX-1", hvBus: "B1", lvBus: "B2",
      ratedMva: 1, ratedHvKv: 11, ratedLvKv: 0.4,
      vscPercent: 6, vscrPercent: 1, ironLossKw: 2, noLoadCurrentPercent: 0.5,
    },
  ],
  lines: [
    { id: "L1", name: "MCC feeder", fromBus: "B2", toBus: "B3", lengthKm: 0.1, cableTypeId: "example-cu-240" },
  ],
  loads: [{ id: "LD1", name: "Lighting & small power", bus: "B2", pKw: 120, qKvar: 40, demandFactor: 0.8 }],
  motors: [
    { id: "M1", name: "Pump motor", bus: "B3", ratedKw: 110, ratedKv: 0.4, efficiency: 0.94, powerFactor: 0.88, startingCurrentRatio: 6.5 },
  ],
  switches: [{ kind: "element", id: "CB1", name: "TX-1 LV breaker", bus: "B2", elementId: "T1", closed: true }],
});
