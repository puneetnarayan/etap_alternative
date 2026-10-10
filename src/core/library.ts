import { z } from "zod";

/**
 * Equipment library. Entries carry a `source` so users can tell
 * authoritative data (manufacturer datasheet, IEC table) from placeholders.
 */

export const LibrarySourceSchema = z.enum(["example", "manufacturer", "standard", "user"]);

export const CableTypeSchema = z.object({
  id: z.string().min(1),
  description: z.string(),
  ratedKv: z.number().positive(),
  conductor: z.enum(["Cu", "Al"]),
  crossSectionMm2: z.number().positive(),
  rOhmPerKm: z.number().nonnegative(),
  xOhmPerKm: z.number().nonnegative(),
  /** Continuous current in a stated installation method; reference lives in `note`. */
  ampacityA: z.number().positive().optional(),
  source: LibrarySourceSchema,
  note: z.string().optional(),
});

export const TransformerTypeSchema = z.object({
  id: z.string().min(1),
  description: z.string(),
  ratedMva: z.number().positive(),
  ratedHvKv: z.number().positive(),
  ratedLvKv: z.number().positive(),
  vscPercent: z.number().positive(),
  vscrPercent: z.number().nonnegative(),
  ironLossKw: z.number().nonnegative(),
  noLoadCurrentPercent: z.number().nonnegative(),
  vectorGroup: z.string().optional(),
  source: LibrarySourceSchema,
  note: z.string().optional(),
});

export const LibrarySchema = z.object({
  cables: z.array(CableTypeSchema).default([]),
  transformers: z.array(TransformerTypeSchema).default([]),
});

export type CableType = z.infer<typeof CableTypeSchema>;
export type TransformerType = z.infer<typeof TransformerTypeSchema>;
export type Library = z.infer<typeof LibrarySchema>;

/**
 * Seed data for UI development only. Values are round, illustrative numbers
 * and are NOT taken from a standard or datasheet; replace before real use.
 */
export const EXAMPLE_LIBRARY: Library = {
  cables: [
    {
      id: "example-cu-95",
      description: "Example Cu cable 95 mm2, 0.6/1 kV",
      ratedKv: 1,
      conductor: "Cu",
      crossSectionMm2: 95,
      rOhmPerKm: 0.2,
      xOhmPerKm: 0.08,
      ampacityA: 250,
      source: "example",
      note: "Placeholder values; verify against IEC 60287 / IEC 60364-5-52 or datasheet.",
    },
    {
      id: "example-cu-240",
      description: "Example Cu cable 240 mm2, 0.6/1 kV",
      ratedKv: 1,
      conductor: "Cu",
      crossSectionMm2: 240,
      rOhmPerKm: 0.08,
      xOhmPerKm: 0.07,
      ampacityA: 450,
      source: "example",
      note: "Placeholder values; verify against IEC 60287 / IEC 60364-5-52 or datasheet.",
    },
  ],
  transformers: [
    {
      id: "example-1mva-11-0.4",
      description: "Example 1 MVA 11/0.4 kV",
      ratedMva: 1,
      ratedHvKv: 11,
      ratedLvKv: 0.4,
      vscPercent: 6,
      vscrPercent: 1,
      ironLossKw: 2,
      noLoadCurrentPercent: 0.5,
      vectorGroup: "Dyn11",
      source: "example",
      note: "Placeholder values; replace with datasheet or IEC 60076 data.",
    },
  ],
};

export function findCableType(lib: Library, id: string): CableType | undefined {
  return lib.cables.find((c) => c.id === id);
}

export function findTransformerType(lib: Library, id: string): TransformerType | undefined {
  return lib.transformers.find((t) => t.id === id);
}
