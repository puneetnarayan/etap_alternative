import MccLoadList from "./MccLoadList";

export const metadata = { title: "MCC Load List / Breaker Selection – Open Power Design" };

export default function Page() {
  return (
    <main>
      <h1>MCC Load List / Breaker Selection</h1>
      <p>
        Prepares an MCC load schedule with <b>preliminary</b> breaker selections as an Excel workbook (a Load list sheet with working
        formulas, and a Notes sheet with the assumptions). Currents come from the rated amps in the load list, or from kW at power factor
        0.7 (general feeder, DOL) and 0.9 (VFD). MCCB ratings 32, 63, 100, 160, 250 and 400 A; from 500 A an EDO ACB (never a 630 A MCCB);
        400 to 500 A needs engineering review. Selections need engineering approval.{" "}
        <a href="/PMCC_sample_load_list.xlsx" download>Download a sample PMCC load list (.xlsx)</a>.
      </p>
      <MccLoadList />
    </main>
  );
}
