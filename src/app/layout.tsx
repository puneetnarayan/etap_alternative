import type { ReactNode } from "react";

export const metadata = {
  title: "Open Power Design",
  description: "Open-source electrical network design for industrial plants (IEC).",
};

export default function RootLayout({ children }: { children: ReactNode }) {
  return (
    <html lang="en">
      <body style={{ fontFamily: "system-ui, sans-serif", margin: "2rem auto", maxWidth: 1100, padding: "0 1rem" }}>
        {children}
      </body>
    </html>
  );
}
