import ImccDesign from "./ImccDesign";

export const metadata = { title: "IMCC Design – Open Power Design" };

export default function Page() {
  return (
    <main>
      <h1>IMCC Design</h1>
      <p>
        A separate tool: it works from the IMCC load list (file name starting with IMCC) and is not connected to the other tabs, except through the <b>Automatic PMCC feeders</b> switch in the sidebar (ON by default).
        Each IMCC has two incomers and a bus coupler; exactly one incomer is ON at a time. The number of IMCCs follows from the working loads (incomer limit 1000 A).
        Rules and data flow: <code>docs/imcc.md</code>; unresolved items: <code>OPEN_ITEMS.md</code>.
      </p>
      <ImccDesign />
    </main>
  );
}
