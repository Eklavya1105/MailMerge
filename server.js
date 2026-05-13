const express = require("express");
const multer = require("multer");
const path = require("node:path");
const fs = require("node:fs");
const { execSync } = require("node:child_process");
const carbone = require("carbone");
const { parse } = require("csv-parse/sync");
const archiver = require("archiver");
const XLSX = require("xlsx");
const mammoth = require("mammoth");
const { PDFDocument } = require("pdf-lib");

const app = express();

const storage = multer.diskStorage({
  destination: "uploads/",
  filename: (req, file, cb) => cb(null, Date.now() + "-" + file.originalname)
});
const upload = multer({ storage });

app.use(express.json());
app.use(express.static("public"));
app.use("/output", express.static("output"));
app.use("/template-preview", express.static("preview"));

function convertToPDF(docxPath, outputDir) {
  execSync(`libreoffice --headless -env:UserInstallation=file:///tmp/libreoffice_headless --convert-to pdf --outdir "${outputDir}" "${docxPath}"`);
}

app.post("/preview", upload.single("template"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: "No file uploaded" });
    }
    const previewDir = path.join(__dirname, "preview");
    fs.mkdirSync(previewDir, { recursive: true });
    // Clean previous preview
    fs.readdirSync(previewDir).forEach(f => fs.unlinkSync(path.join(previewDir, f)));

    // Convert docx to PDF for preview
    const output = execSync(`libreoffice --headless -env:UserInstallation=file:///tmp/libreoffice_headless --convert-to pdf --outdir "${previewDir}" "${req.file.path}"`, { encoding: "utf-8" });
    console.log("LibreOffice output:", output);

    // Find the generated PDF
    const pdfs = fs.readdirSync(previewDir).filter(f => f.endsWith(".pdf"));
    console.log("PDFs found:", pdfs);

    if (pdfs.length > 0) {
      const previewPath = path.join(previewDir, "preview.pdf");
      if (pdfs[0] !== "preview.pdf") {
        fs.renameSync(path.join(previewDir, pdfs[0]), previewPath);
      }
      fs.unlinkSync(req.file.path);
      res.json({ success: true, pdf: "/template-preview/preview.pdf?t=" + Date.now() });
    } else {
      fs.unlinkSync(req.file.path);
      res.status(500).json({ success: false, error: "PDF conversion failed - no output generated" });
    }
  } catch (err) {
    console.error("Preview error:", err.message);
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post("/merge", upload.fields([{ name: "template" }, { name: "data" }]), (req, res) => {
  try {
    if (!req.files["template"] || !req.files["data"]) {
      return res.status(400).json({ success: false, error: "Please upload both a template and a data file." });
    }
    const templateFile = req.files["template"][0];
    const dataFile = req.files["data"][0];

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
          // Save a backup of the original single-page PDF
          const originalDir = path.join(__dirname, "originals");
          fs.mkdirSync(originalDir, { recursive: true });
          fs.copyFileSync(path.join(outputDir, `output_${i + 1}.pdf`), path.join(originalDir, `output_${i + 1}.pdf`));
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

app.post("/combine", upload.array("pages", 3), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ success: false, error: "No files uploaded" });
    }

    const outputDir = path.join(__dirname, "output");
    fs.mkdirSync(outputDir, { recursive: true });

    // Convert each file to PDF
    const pdfPaths = [];
    for (const file of req.files) {
      if (file.originalname.endsWith(".pdf")) {
        pdfPaths.push(file.path);
      } else {
        execSync(`libreoffice --headless -env:UserInstallation=file:///tmp/libreoffice_headless --convert-to pdf --outdir "${outputDir}" "${file.path}"`);
        const baseName = path.basename(file.path, path.extname(file.path));
        pdfPaths.push(path.join(outputDir, baseName + ".pdf"));
        fs.unlinkSync(file.path);
      }
    }

    // Merge all PDFs into one
    const combinedPath = path.join(outputDir, "combined.pdf");
    const mergedPdf = await PDFDocument.create();
    for (const pdfPath of pdfPaths) {
      if (!fs.existsSync(pdfPath)) continue;
      const pdfBytes = fs.readFileSync(pdfPath);
      const pdf = await PDFDocument.load(pdfBytes);
      const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
      pages.forEach(page => mergedPdf.addPage(page));
    }
    const mergedBytes = await mergedPdf.save();
    fs.writeFileSync(combinedPath, mergedBytes);

    // Cleanup intermediate PDFs
    pdfPaths.forEach(p => {
      if (fs.existsSync(p) && p !== combinedPath) fs.unlinkSync(p);
    });

    res.json({ success: true, filename: "combined.pdf" });
  } catch (err) {
    console.error("Combine error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post("/apply-to-all", async (req, res) => {
  try {
    const { source, targets, personalizedPageIndex } = req.body;
    const outputDir = path.join(__dirname, "output");
    const originalDir = path.join(__dirname, "originals");
    const sourcePath = path.join(outputDir, source);

    if (!fs.existsSync(sourcePath)) {
      return res.status(404).json({ success: false, error: "Source file not found" });
    }

    const sourcePdf = await PDFDocument.load(fs.readFileSync(sourcePath));
    const sourcePageCount = sourcePdf.getPageCount();

    // Get indices of extra (static) pages from source
    const extraIndices = [];
    for (let i = 0; i < sourcePageCount; i++) {
      if (i !== personalizedPageIndex) extraIndices.push(i);
    }

    for (const target of targets) {
      const originalPath = path.join(originalDir, target);
      if (!fs.existsSync(originalPath)) continue;

      const targetOrigPdf = await PDFDocument.load(fs.readFileSync(originalPath));
      const newPdf = await PDFDocument.create();

      // Build the new PDF following source's page order
      for (let i = 0; i < sourcePageCount; i++) {
        if (i === personalizedPageIndex) {
          // Insert target's personalized page here
          const [page] = await newPdf.copyPages(targetOrigPdf, [0]);
          newPdf.addPage(page);
        } else {
          // Copy the static page from source
          const [page] = await newPdf.copyPages(sourcePdf, [i]);
          newPdf.addPage(page);
        }
      }

      const finalBytes = await newPdf.save();
      fs.writeFileSync(path.join(outputDir, target), finalBytes);
    }

    res.json({ success: true });
  } catch (err) {
    console.error("Apply to all error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post("/reset-pages", (req, res) => {
  try {
    const { target } = req.body;
    const outputDir = path.join(__dirname, "output");
    const originalDir = path.join(__dirname, "originals");
    const originalPath = path.join(originalDir, target);
    const targetPath = path.join(outputDir, target);

    if (!fs.existsSync(originalPath)) {
      return res.status(404).json({ success: false, error: "Original file not found" });
    }

    // Restore original
    fs.copyFileSync(originalPath, targetPath);
    res.json({ success: true });
  } catch (err) {
    console.error("Reset error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post("/add-page", upload.single("page"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: "No file uploaded" });
    }
    const { target, position, maxPages } = req.body;
    const outputDir = path.join(__dirname, "output");
    const targetPath = path.join(outputDir, target);

    if (!fs.existsSync(targetPath)) {
      return res.status(404).json({ success: false, error: "Target PDF not found" });
    }

    // Convert uploaded file to PDF if it's not already
    let pagePdfPath = req.file.path;
    if (!req.file.originalname.endsWith(".pdf")) {
      execSync(`libreoffice --headless -env:UserInstallation=file:///tmp/libreoffice_headless --convert-to pdf --outdir "${outputDir}" "${req.file.path}"`);
      const baseName = path.basename(req.file.path, path.extname(req.file.path));
      pagePdfPath = path.join(outputDir, baseName + ".pdf");
      fs.unlinkSync(req.file.path);
    }

    // Check page limits
    const targetPdf = await PDFDocument.load(fs.readFileSync(targetPath));
    const pagePdf = await PDFDocument.load(fs.readFileSync(pagePdfPath));

    // Each uploaded file must be exactly 1 page
    if (pagePdf.getPageCount() > 1) {
      if (fs.existsSync(pagePdfPath) && pagePdfPath !== req.file.path) fs.unlinkSync(pagePdfPath);
      return res.status(400).json({ success: false, error: `Uploaded file has ${pagePdf.getPageCount()} pages. Only single-page documents are allowed.` });
    }

    const totalPages = targetPdf.getPageCount() + pagePdf.getPageCount();
    const limit = parseInt(maxPages) || 999;

    if (totalPages > limit) {
      if (fs.existsSync(pagePdfPath) && pagePdfPath !== req.file.path) fs.unlinkSync(pagePdfPath);
      return res.status(400).json({ success: false, error: `Page limit exceeded. Adding this file would make ${totalPages} pages (max ${limit}).` });
    }

    // Merge PDFs
    const copiedPages = await targetPdf.copyPages(pagePdf, pagePdf.getPageIndices());

    if (position === "before") {
      copiedPages.reverse().forEach(page => targetPdf.insertPage(0, page));
    } else {
      copiedPages.forEach(page => targetPdf.addPage(page));
    }

    const mergedBytes = await targetPdf.save();
    fs.writeFileSync(targetPath, mergedBytes);

    // Cleanup
    if (fs.existsSync(pagePdfPath) && pagePdfPath !== targetPath) fs.unlinkSync(pagePdfPath);

    res.json({ success: true });
  } catch (err) {
    console.error("Add page error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post("/rearrange", async (req, res) => {
  try {
    const { target, pageOrder } = req.body; // pageOrder = [2, 0, 1] means page 3 first, then page 1, then page 2
    const outputDir = path.join(__dirname, "output");
    const targetPath = path.join(outputDir, target);

    if (!fs.existsSync(targetPath)) {
      return res.status(404).json({ success: false, error: "Target PDF not found" });
    }

    const srcPdf = await PDFDocument.load(fs.readFileSync(targetPath));
    const newPdf = await PDFDocument.create();

    for (const pageIndex of pageOrder) {
      const [copiedPage] = await newPdf.copyPages(srcPdf, [pageIndex]);
      newPdf.addPage(copiedPage);
    }

    const newBytes = await newPdf.save();
    fs.writeFileSync(targetPath, newBytes);

    res.json({ success: true, pageCount: newPdf.getPageCount() });
  } catch (err) {
    console.error("Rearrange error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get("/page-count/:filename", async (req, res) => {
  try {
    const outputDir = path.join(__dirname, "output");
    const filePath = path.join(outputDir, req.params.filename);
    const pdfDoc = await PDFDocument.load(fs.readFileSync(filePath));
    res.json({ success: true, pageCount: pdfDoc.getPageCount() });
  } catch (err) {
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