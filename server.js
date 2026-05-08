const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const { execSync } = require("child_process");
const carbone = require("carbone");
const { parse } = require("csv-parse/sync");
const archiver = require("archiver");
const XLSX = require("xlsx");

const app = express();

const storage = multer.diskStorage({
  destination: "uploads/",
  filename: (req, file, cb) => cb(null, Date.now() + "-" + file.originalname)
});
const upload = multer({ storage });

app.use(express.static("public"));
app.use("/output", express.static("output"));

function convertToPDF(docxPath, outputDir) {
  execSync(`libreoffice --headless --convert-to pdf --outdir "${outputDir}" "${docxPath}"`);
}

app.post("/merge", upload.fields([{ name: "template" }, { name: "data" }]), (req, res) => {
  try {
    if (!req.files["template"] || !req.files["data"]) {
      return res.status(400).json({ success: false, error: "Please upload both a template and a data file." });
    }
    const templateFile = req.files["template"][0];
    const dataFile = req.files["data"][0];

    if (!templateFile.originalname.endsWith(".docx")) {
      return res.status(400).json({ success: false, error: "Template must be a .docx file" });
    }
    const ext = path.extname(dataFile.originalname).toLowerCase();
    if (![".csv", ".xlsx", ".xls"].includes(ext)) {
      return res.status(400).json({ success: false, error: "Data must be a .csv or .xlsx file" });
    }

    let rows;
    if (ext === ".csv") {
      const csvContent = fs.readFileSync(dataFile.path, "utf-8");
      rows = parse(csvContent, { columns: true, skip_empty_lines: true });
    } else {
      const workbook = XLSX.readFile(dataFile.path);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet);
    }

    const outputDir = path.join(__dirname, "output");
    fs.mkdirSync(outputDir, { recursive: true });
    fs.readdirSync(outputDir).forEach(f => fs.unlinkSync(path.join(outputDir, f)));

    const results = [];

    function processRow(i) {
      if (i >= rows.length) {
        fs.unlinkSync(templateFile.path);
        fs.unlinkSync(dataFile.path);
        results.sort((a, b) => a.index - b.index);
        return res.json({ success: true, files: results });
      }

      carbone.render(templateFile.path, rows[i], (err, result) => {
        if (err) {
          console.error(`Error generating doc ${i + 1}:`, err);
          return processRow(i + 1);
        }
        const docxFile = path.join(outputDir, `output_${i + 1}.docx`);
        fs.writeFileSync(docxFile, result);

        try {
          convertToPDF(docxFile, outputDir);
          fs.unlinkSync(docxFile);
          results.push({ index: i + 1, filename: `output_${i + 1}.pdf`, data: rows[i] });
        } catch (e) {
          console.error(`Error converting to PDF ${i + 1}:`, e.message);
        }

        processRow(i + 1);
      });
    }

    processRow(0);
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get("/download-all", (req, res) => {
  const outputDir = path.join(__dirname, "output");
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", "attachment; filename=mail_merge_output.zip");

  const archive = archiver("zip");
  archive.pipe(res);
  archive.directory(outputDir, false);
  archive.finalize();
});

const PORT = 3001;
app.listen(PORT, () => {
  console.log(`Mail Merge app running at http://localhost:${PORT}`);
});
