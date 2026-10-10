# Open Power Design

Open-source electrical network design for **industrial plants**, **IEC** conventions, hosted on Vercel.
Early-stage: this is an independent project, not affiliated with ETAP.

## Status: Phase 1 – Foundation

- `src/core/schema.ts` – network data model (zod). Units: kV, kW/kVAr/MVA, km, ohm/km, A.
- `src/core/library.ts` – equipment library with a `source` tag per entry. **The bundled entries are illustrative placeholders, not standard or datasheet data.**
- `src/core/topology.ts` – switch-aware island detection.
- `src/core/validate.ts` – model validation (ids, references, voltage levels, sources, islands).
- `src/core/io.ts` – project JSON parse/serialize.
- `src/app` – minimal Next.js page (deployable to Vercel as-is).

```
npm install
npm test          # unit tests
npm run typecheck
npm run dev
```

## Roadmap

1. **Foundation** (this phase)
2. Solver: load flow via an existing permissively-licensed solver (candidate: pandapower; verify licence and Vercel bundle limits) plus load schedule
3. Single-line diagram editor
4. Cable sizing / voltage drop (IEC), loading checks, motor starting
5. Reports, scenarios, import/export

Deferred: fault level, power generation, protection grading (and arc flash, which depends on fault levels).

## Licence

MIT. Dependencies must be under free/open-source licences.
