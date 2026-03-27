import "./globals.css";

export const metadata = {
  title: "Excel Template Portal",
  description: "Generate enrollment templates for schools."
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
