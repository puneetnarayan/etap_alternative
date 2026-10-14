import { DEFAULT_MCC_ASSUMPTIONS, DEFAULT_OPTIONS } from "@/core";

const pct = (x: number) => `${Math.round(x * 100)} %`;

const STEPS: [string, string][] = [
  ["Collect the load list", "Every load with its description, type (DOL motor, VFD motor or feeder), kW or amps, voltage level and whether it is working or standby."],
  ["Convert loads to kVA and current", "Motor kW becomes kVA through power factor and efficiency; loads given in amps become kVA at the line voltage. Standby loads are counted at 100 %."],
  ["Size and arrange the transformers", "Choose the number and rating of transformers, the sets of 2 or 3 with bus couplers, and spread the loads so that each section and each set stays within its limit, also with one transformer out."],
  ["Select the switchgear", "Choose feeder and incomer breakers (MCCB or EDO ACB), overload settings and CT ratios for each MCC or PMCC / IMCC."],
  ["Size the cables", "Cable current capacity, derating and voltage drop for each feeder."],
  ["Check faults and protection", "Prospective fault levels, breaking capacity, motor starting, earthing and discrimination between devices."],
  ["Review and approve", "Every figure from a tool like this one is preliminary until an engineer has checked it against the project's design basis and standards."],
];

const STANDARDS: [string, string][] = [
  ["IEC 60076", "Power transformers"],
  ["IEC 60947-2", "Low-voltage circuit-breakers (MCCB, ACB)"],
  ["IEC 61439", "Low-voltage switchgear and controlgear assemblies (MCC, PMCC / IMCC)"],
  ["IEC 60364", "Low-voltage electrical installations"],
  ["IEC 60287", "Calculation of the current rating of cables"],
  ["IEC 60909", "Short-circuit current calculation"],
];

export default function Home() {
  const o = DEFAULT_OPTIONS;
  const a = DEFAULT_MCC_ASSUMPTIONS;
  return (
    <main>
      <h1>Open Power Design</h1>
      <p>
        Open-source electrical design for industrial plants, built around IEC practice. The calculation tools are in the sidebar: pick one,
        drop your load list (Excel) and get the results on screen and as a formatted Excel file.
      </p>

      <h2>How an industrial power system is designed</h2>
      <ol>
        {STEPS.map(([t, d]) => <li key={t} style={{ marginBottom: 6 }}><b>{t}.</b> {d}</li>)}
      </ol>
      <p>
        Today the sidebar covers transformer sizing and the MCC load list with breaker selection (steps 2 to 4, from a load list you provide), and a load-flow preview on a built-in sample plant. Cable sizing is next;
        fault level, power generation and protection grading come later.
      </p>

      <h2>Design basis used in the tools</h2>
      <ul>
        <li>Motor kVA = kW ÷ (power factor × efficiency), with power factor {a.pfDol} for DOL motors, {a.pfVfd} for VFD motors and {a.pfGeneral} for other feeders, and {pct(o.efficiency)} efficiency in the transformer tool.</li>
        <li>Amps to kVA: √3 × V × A at the voltage level (433 V or 690 V; the 433 V calculation voltage is editable).</li>
        <li>Each transformer is loaded to at most {pct(o.normalLimit)} in normal running and {pct(o.contingencyLimit)} with one transformer of its set out. These limits are fixed.</li>
        <li>Transformers come in sets of 3 (3 incomers, 2 bus couplers), with a set of 2 (2 incomers, 1 bus coupler) only for the remainder; the number grows when the loads cannot be balanced.</li>
        <li>Breakers: MCCB 6 to 400 A from the BOM list (6, 10, 16, 25, 32, 50, 63, 100, 125, 160, 200, 250, 400 A), a 630 A feeder as MCCB or ACB, EDO ACB up to 4000 A, engineering review between 400 and 500 A unless 630 A is taken as an MCCB. CT secondary 1 A.</li>
      </ul>

      <h2>Standards worth knowing</h2>
      <p>These are the main IEC documents behind the topics above. Use the editions your project specifies; the list is not exhaustive.</p>
      <ul>
        {STANDARDS.map(([n, t]) => <li key={n}><b>{n}</b>: {t}</li>)}
      </ul>

      <p style={{ marginTop: 24, fontSize: 14 }}>
        The tools are preliminary engineering aids and not a substitute for approved electrical design.
      </p>
    </main>
  );
}
