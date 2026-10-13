import type { ReactNode } from "react";
import AppShell from "./AppShell";

export const metadata = {
  title: "Open Power Design",
  description: "Open-source electrical network design for industrial plants (IEC).",
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en">
      <body style={{ fontFamily: "system-ui, sans-serif", margin: 0 }}>
        <AppShell>{children}</AppShell>
      </body>
    </html>
  );
}
