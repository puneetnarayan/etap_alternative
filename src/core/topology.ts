import type { Project } from "./schema";

export interface Island {
  buses: string[];
  sourceIds: string[];
  energised: boolean;
}

class UnionFind {
  private parent = new Map<string, string>();
  add(x: string) {
    if (!this.parent.has(x)) this.parent.set(x, x);
  }
  find(x: string): string {
    let p = this.parent.get(x)!;
    while (p !== this.parent.get(p)!) {
      this.parent.set(p, this.parent.get(this.parent.get(p)!)!);
      p = this.parent.get(p)!;
    }
    return p;
  }
  union(a: string, b: string) {
    this.add(a);
    this.add(b);
    this.parent.set(this.find(a), this.find(b));
  }
}

/** Element ids isolated by at least one open element-switch. */
export function isolatedElements(project: Project): Set<string> {
  const out = new Set<string>();
  for (const s of project.switches) if (s.kind === "element" && !s.closed) out.add(s.elementId);
  return out;
}

/**
 * Group buses into electrically connected islands given current switch
 * states. Branches referencing unknown buses are skipped (validation
 * reports them separately).
 */
export function findIslands(project: Project): Island[] {
  const busIds = new Set(project.buses.map((b) => b.id));
  const uf = new UnionFind();
  for (const b of project.buses) uf.add(b.id);

  const isolated = isolatedElements(project);
  const connect = (a: string, b: string) => {
    if (busIds.has(a) && busIds.has(b)) uf.union(a, b);
  };

  for (const l of project.lines) if (!isolated.has(l.id)) connect(l.fromBus, l.toBus);
  for (const t of project.transformers) if (!isolated.has(t.id)) connect(t.hvBus, t.lvBus);
  for (const s of project.switches) if (s.kind === "coupler" && s.closed) connect(s.busA, s.busB);

  const groups = new Map<string, string[]>();
  for (const b of project.buses) {
    const root = uf.find(b.id);
    groups.set(root, [...(groups.get(root) ?? []), b.id]);
  }

  return [...groups.values()].map((buses) => {
    const set = new Set(buses);
    const sourceIds = project.sources.filter((s) => set.has(s.bus)).map((s) => s.id);
    return { buses, sourceIds, energised: sourceIds.length > 0 };
  });
}
