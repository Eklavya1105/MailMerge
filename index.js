const fs = require("fs");
const path = require("path");
const carbone = require("carbone");
const { parse } = require("csv-parse/sync");

function loadCSV(filepath) {
  const content = fs.readFileSync(filepath, "utf-8");
  return parse(content, { columns: true, skip_empty_lines: true });
}

function mailMerge(templatePath, dataPath, outputDir = "output") {
  fs.mkdirSync(outputDir, { recursive: true });

  const rows = loadCSV(dataPath);

  rows.forEach((data, i) => {
    // Generate .docx
    carbone.render(templatePath, data, (err, result) => {
      if (err) {
        console.error(`Error generating document ${i + 1}:`, err);
        return;
      }
      const docxFile = path.join(outputDir, `output_${i + 1}.docx`);
      fs.writeFileSync(docxFile, result);
      console.log(`Generated: ${docxFile}`);
    });

    // Generate .pdf
    carbone.render(templatePath, data, { convertTo: "pdf" }, (err, result) => {
      if (err) {
        console.error(`Error generating PDF ${i + 1}:`, err);
        return;
      }
      const pdfFile = path.join(outputDir, `output_${i + 1}.pdf`);
      fs.writeFileSync(pdfFile, result);
      console.log(`Generated: ${pdfFile}`);
    });
  });
}

const [templatePath, dataPath] = process.argv.slice(2);

if (!templatePath || !dataPath) {
  console.log("Usage: node index.js <template.docx> <data.csv>");
  process.exit(1);
}

mailMerge(templatePath, dataPath);
