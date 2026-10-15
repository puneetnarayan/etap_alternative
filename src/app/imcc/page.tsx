import ImccDesign from "./ImccDesign";

export const metadata = { title: "IMCC Design – Open Power Design" };

export default function Page() {
  return (
    <main>
      <h1>IMCC Design</h1>
      <p>
        A separate tool: it works from the IMCC load list (file name starting with IMCC) and is not connected to the other tabs, except through the <b>Automatic PMCC feeders</b> switch in the sidebar (ON by default).
        Incomer 1 is at the beginning of the panel; Incomer 2 and the bus coupler are side by side at the other end. Normally Incomer 1 is ON and the other two are OFF; if Incomer 1 is unhealthy,
        Incomer 2 and the bus coupler are both switched ON (Incomer 1 is switched OFF first by the interlock). The incomer ACB (630 / 800 / 1000 / 1250 A) and its maximum loading are set below; the number of IMCCs follows from the working loads.
        Rules and data flow: <code>docs/imcc.md</code>; unresolved items: <code>OPEN_ITEMS.md</code>.
      </p>
      <ImccDesign />
    </main>
  );
}
