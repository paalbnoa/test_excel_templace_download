import "./globals.css";

export const metadata = {
  title: "Semester fee template portal",
  description: "Generate semester fee templates for institutions."
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
