const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer");
const { marked } = require("marked");

async function main() {
  const md = fs.readFileSync(path.join(__dirname, "QAReturns_Quotation.md"), "utf-8");
  const html = marked(md);

  const fullHtml = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @page { margin: 20mm 18mm; }
  body {
    font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
    font-size: 13px;
    color: #1e293b;
    line-height: 1.6;
    max-width: 100%;
  }
  h1 {
    font-size: 28px;
    color: #1f4e79;
    border-bottom: 3px solid #1f4e79;
    padding-bottom: 8px;
    margin-top: 0;
    margin-bottom: 4px;
  }
  /* Business info block right after h1 */
  h1 + p {
    margin-top: 0;
    font-size: 13px;
    color: #334155;
    line-height: 1.8;
  }
  h2 {
    font-size: 18px;
    color: #1f4e79;
    margin-top: 28px;
    border-bottom: 1px solid #cbd5e1;
    padding-bottom: 4px;
  }
  h3 {
    font-size: 15px;
    color: #334155;
    margin-top: 20px;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    margin: 12px 0;
    font-size: 12px;
  }
  th {
    background: #1f4e79;
    color: white;
    padding: 8px 10px;
    text-align: left;
    font-weight: 600;
  }
  td {
    padding: 7px 10px;
    border-bottom: 1px solid #e2e8f0;
  }
  tr:nth-child(even) td {
    background: #f8fafc;
  }
  strong {
    color: #0f172a;
  }
  em {
    color: #64748b;
    font-size: 12px;
  }
  hr {
    border: none;
    border-top: 1px solid #e2e8f0;
    margin: 24px 0;
  }
  ul, ol {
    padding-left: 20px;
  }
  li {
    margin: 4px 0;
  }
  /* Prevent sections from breaking across pages */
  h2, h3 {
    page-break-after: avoid;
  }
  table {
    page-break-inside: avoid;
  }
  /* Force page break before Pricing Options */
  h2:nth-of-type(3) {
    page-break-before: always;
  }
</style>
</head>
<body>
${html}
</body>
</html>`;

  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.setContent(fullHtml, { waitUntil: "networkidle0" });

  const outputPath = path.join(__dirname, "QAReturns_Quotation.pdf");
  await page.pdf({
    path: outputPath,
    format: "A4",
    printBackground: true,
  });

  await browser.close();
  console.log("PDF generated:", outputPath);
}

main().catch(console.error);
