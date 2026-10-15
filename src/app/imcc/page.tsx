import ImccDesign from "./ImccDesign";

export const metadata = { title: "IMCC Design – Open Power Design" };

export default function Page() {
  return (
    <main>
      <h1>IMCC Design</h1>
      <p>
        Each IMCC has two incomers, fed from outgoing feeders of PMCCs, and a bus coupler. Exactly one incomer is ON at a time; the other incomer and the
        coupler stay OFF. Saving an IMCC creates and updates its incomer feeders in the upstream PMCC lists automatically. Rules and data model: <code>docs/imcc.md</code>;
        unresolved engineering items: <code>OPEN_ITEMS.md</code>.
      </p>
      <ImccDesign />
    </main>
  );
}
