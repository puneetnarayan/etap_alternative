import { findIslands } from "./topology";
import { findCableType, type Library } from "./library";
import type { Project } from "./schema";

export type Severity = "error" | "warning" | "info";

export interface Issue {
  code: string;
  severity: Severity;
  message: string;
  elementId?: string;
}

/** Relative tolerance for comparing rated vs. bus nominal voltage. */
const VOLTAGE_TOLERANCE = 0.1;

const mismatch = (a: number, b: number) => Math.abs(a - b) / b > VOLTAGE_TOLERANCE;

export function validateProject(project: Project, library?: Library): Issue[] {
  const issues: Issue[] = [];
  const add = (i: Issue) => issues.push(i);

  // Unique ids across all element kinds.
  const seen = new Map<string, string>();
  const collections: [string, { id: string }[]][] = [
    ["bus", project.buses],
    ["line", project.lines],
    ["transformer", project.transformers],
    ["load", project.loads],
    ["motor", project.motors],
    ["source", project.sources],
    ["switch", project.switches],
  ];
  for (const [kind, items] of collections) {
    for (const item of items) {
      const prev = seen.get(item.id);
      if (prev) {
        add({
          code: "DUPLICATE_ID",
          severity: "error",
          elementId: item.id,
          message: `Id "${item.id}" used by both a ${prev} and a ${kind}.`,
        });
      } else {
        seen.set(item.id, kind);
      }
    }
  }

  const buses = new Map(project.buses.map((b) => [b.id, b]));
  const needBus = (busId: string, owner: string) => {
    if (buses.has(busId)) return true;
    add({
      code: "MISSING_BUS_REF",
      severity: "error",
      elementId: owner,
      message: `"${owner}" references unknown bus "${busId}".`,
    });
    return false;
  };

  for (const l of project.lines) {
    const okA = needBus(l.fromBus, l.id);
    const okB = needBus(l.toBus, l.id);
    if (l.fromBus === l.toBus) {
      add({ code: "SELF_LOOP", severity: "error", elementId: l.id, message: `Line "${l.id}" connects bus "${l.fromBus}" to itself.` });
    }
    if (okA && okB && buses.get(l.fromBus)!.nominalKv !== buses.get(l.toBus)!.nominalKv) {
      add({
        code: "LINE_VOLTAGE_MISMATCH",
        severity: "error",
        elementId: l.id,
        message: `Line "${l.id}" joins buses at different nominal voltages; use a transformer.`,
      });
    }
    if (library && l.cableTypeId !== undefined && !findCableType(library, l.cableTypeId)) {
      add({ code: "UNKNOWN_CABLE_TYPE", severity: "error", elementId: l.id, message: `Line "${l.id}" uses cable type "${l.cableTypeId}" which is not in the library.` });
    }
    const hasImpedance = l.cableTypeId !== undefined || (l.rOhmPerKm !== undefined && l.xOhmPerKm !== undefined);
    if (!hasImpedance) {
      add({
        code: "LINE_NO_IMPEDANCE",
        severity: "error",
        elementId: l.id,
        message: `Line "${l.id}" needs a cableTypeId or both rOhmPerKm and xOhmPerKm.`,
      });
    }
  }

  for (const t of project.transformers) {
    const okH = needBus(t.hvBus, t.id);
    const okL = needBus(t.lvBus, t.id);
    if (t.hvBus === t.lvBus) {
      add({ code: "SELF_LOOP", severity: "error", elementId: t.id, message: `Transformer "${t.id}" has the same HV and LV bus.` });
    }
    if (t.ratedHvKv <= t.ratedLvKv) {
      add({ code: "TRANSFORMER_RATIO", severity: "warning", elementId: t.id, message: `Transformer "${t.id}": HV rating is not above LV rating.` });
    }
    if (t.vscrPercent > t.vscPercent) {
      add({ code: "TRANSFORMER_IMPEDANCE", severity: "error", elementId: t.id, message: `Transformer "${t.id}": resistive part of Vsc exceeds total Vsc.` });
    }
    if (okH && mismatch(t.ratedHvKv, buses.get(t.hvBus)!.nominalKv)) {
      add({ code: "TRANSFORMER_VOLTAGE_MISMATCH", severity: "warning", elementId: t.id, message: `Transformer "${t.id}" HV rating differs from bus "${t.hvBus}" nominal voltage by more than 10%.` });
    }
    if (okL && mismatch(t.ratedLvKv, buses.get(t.lvBus)!.nominalKv)) {
      add({ code: "TRANSFORMER_VOLTAGE_MISMATCH", severity: "warning", elementId: t.id, message: `Transformer "${t.id}" LV rating differs from bus "${t.lvBus}" nominal voltage by more than 10%.` });
    }
  }

  for (const x of project.loads) needBus(x.bus, x.id);
  for (const m of project.motors) {
    if (needBus(m.bus, m.id) && m.ratedKv !== undefined && mismatch(m.ratedKv, buses.get(m.bus)!.nominalKv)) {
      add({ code: "MOTOR_VOLTAGE_MISMATCH", severity: "warning", elementId: m.id, message: `Motor "${m.id}" rated voltage differs from bus "${m.bus}" nominal voltage by more than 10%.` });
    }
  }
  for (const s of project.sources) needBus(s.bus, s.id);

  const elementIds = new Set([...project.lines, ...project.transformers].map((e) => e.id));
  for (const sw of project.switches) {
    if (sw.kind === "coupler") {
      needBus(sw.busA, sw.id);
      needBus(sw.busB, sw.id);
    } else {
      needBus(sw.bus, sw.id);
      if (!elementIds.has(sw.elementId)) {
        add({ code: "SWITCH_UNKNOWN_ELEMENT", severity: "error", elementId: sw.id, message: `Switch "${sw.id}" references unknown line/transformer "${sw.elementId}".` });
      }
    }
  }

  // Topology checks.
  if (project.sources.length === 0 && project.buses.length > 0) {
    add({ code: "NO_SOURCE", severity: "error", message: "Project has no source (utility connection)." });
  }
  for (const island of findIslands(project)) {
    if (island.buses.length === 1 && !island.energised) {
      add({ code: "ISOLATED_BUS", severity: "warning", elementId: island.buses[0], message: `Bus "${island.buses[0]}" is not connected to anything.` });
    } else if (!island.energised && project.sources.length > 0) {
      add({ code: "ISLAND_NO_SOURCE", severity: "warning", message: `Buses [${island.buses.join(", ")}] form an island with no source (de-energised).` });
    } else if (island.sourceIds.length > 1) {
      add({ code: "MULTIPLE_SOURCES", severity: "info", message: `Island with buses [${island.buses.join(", ")}] has ${island.sourceIds.length} sources; the solver will need a slack-bus choice.` });
    }
  }

  return issues;
}

export const hasErrors = (issues: Issue[]) => issues.some((i) => i.severity === "error");
