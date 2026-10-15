import BomPage from "./BomPage";

export const metadata = { title: "BOM – Open Power Design" };

export default function Page() {
  return (
    <main>
      <h1>BOM</h1>
      <p>
        One unit BOM for each unique item (type + rating) in your load list, built from the unit-BOM template with its sheet layouts unchanged.
        The ratings come from the preliminary breaker selection of the MCC Load List tab, so its assumptions apply. Quantities are counted from the load list.{" "}
        <a href="/PMCC_sample_load_list.xlsx" download>Download a sample PMCC load list (.xlsx)</a>.
      </p>
      <BomPage />
    </main>
  );
}
