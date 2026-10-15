import Link from "next/link";
import TransformerSizing from "./TransformerSizing";

export const metadata = { title: "Transformer sizing – Open Power Design" };

export default function Page() {
  return (
    <main>
      <p><Link href="/">← Home</Link></p>
      <h1>Transformer sizing</h1>
      <p>
        Add your load list (Excel, top of the sidebar) to find how many transformers you need and what rating, with the loads shared
        across them. <a href="/PMCC_sample_load_list.xlsx" download>Download a sample PMCC load list (.xlsx)</a>.
      </p>
      <TransformerSizing />
    </main>
  );
}
