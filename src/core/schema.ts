import { z } from "zod";

/**
 * Network data model (schema v1).
 *
 * Units are fixed and documented per field so every module (library,
 * validation, future solvers) agrees without conversion:
 *   voltage kV, power kW / kVAr / MVA, length km, impedance ohm/km, current A.
 */

export const SCHEMA_VERSION = 1;

const id = z.string().min(1);
const name = z.string().default("");
const positive = z.number().positive();

export const BusSchema = z.object({
  id,
  name,
  nominalKv: positive,
});

/** Cable / overhead line between two buses. Parameters come from the library or are overridden inline. */
export const LineSchema = z.object({
  id,
  name,
  fromBus: id,
  toBus: id,
  lengthKm: positive,
  parallel: z.number().int().min(1).default(1),
  cableTypeId: id.optional(),
  rOhmPerKm: z.number().nonnegative().optional(),
  xOhmPerKm: z.number().nonnegative().optional(),
  maxCurrentA: positive.optional(),
});

/** Two-winding transformer. */
export const TransformerSchema = z.object({
  id,
  name,
  hvBus: id,
  lvBus: id,
  ratedMva: positive,
  ratedHvKv: positive,
  ratedLvKv: positive,
  vscPercent: positive,
  vscrPercent: z.number().nonnegative(),
  ironLossKw: z.number().nonnegative().default(0),
  noLoadCurrentPercent: z.number().nonnegative().default(0),
  vectorGroup: z.string().optional(),
});

/** Static load (constant power). */
export const LoadSchema = z.object({
  id,
  name,
  bus: id,
  pKw: z.number(),
  qKvar: z.number().default(0),
  demandFactor: z.number().min(0).max(1).default(1),
});

/** Induction motor (data needed for later starting studies). */
export const MotorSchema = z.object({
  id,
  name,
  bus: id,
  ratedKw: positive,
  ratedKv: positive.optional(),
  efficiency: z.number().gt(0).max(1),
  powerFactor: z.number().gt(0).max(1),
  startingCurrentRatio: z.number().gt(1).optional(),
});

/** Utility / external grid connection; acts as the voltage-reference source. */
export const SourceSchema = z.object({
  id,
  name,
  bus: id,
  voltagePu: positive.default(1),
  shortCircuitMva: positive.optional(),
});

export const SwitchSchema = z.discriminatedUnion("kind", [
  /** Bus coupler: connects/disconnects two buses. */
  z.object({
    kind: z.literal("coupler"),
    id,
    name,
    busA: id,
    busB: id,
    closed: z.boolean().default(true),
  }),
  /** Breaker at one end of a line or transformer: open isolates that element. */
  z.object({
    kind: z.literal("element"),
    id,
    name,
    bus: id,
    elementId: id,
    closed: z.boolean().default(true),
  }),
]);

export const ProjectSchema = z.object({
  schemaVersion: z.literal(SCHEMA_VERSION),
  name: z.string().default("Untitled project"),
  frequencyHz: z.union([z.literal(50), z.literal(60)]).default(50),
  buses: z.array(BusSchema).default([]),
  lines: z.array(LineSchema).default([]),
  transformers: z.array(TransformerSchema).default([]),
  loads: z.array(LoadSchema).default([]),
  motors: z.array(MotorSchema).default([]),
  sources: z.array(SourceSchema).default([]),
  switches: z.array(SwitchSchema).default([]),
});

export type Bus = z.infer<typeof BusSchema>;
export type Line = z.infer<typeof LineSchema>;
export type Transformer = z.infer<typeof TransformerSchema>;
export type Load = z.infer<typeof LoadSchema>;
export type Motor = z.infer<typeof MotorSchema>;
export type Source = z.infer<typeof SourceSchema>;
export type Switch = z.infer<typeof SwitchSchema>;
export type Project = z.infer<typeof ProjectSchema>;
