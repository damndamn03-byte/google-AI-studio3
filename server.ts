import express from "express";
import { createServer as createViteServer } from "vite";
import multer from "multer";
import cors from "cors";
import path from "path";
import fs from "fs-extra";
import createReport from "docx-templates";
import { fileURLToPath } from "url";
import JSZip from "jszip";

import sizeOf from "image-size";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Ensure upload directory exists
const UPLOADS_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR);
}

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(cors());
  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ extended: true, limit: '50mb' }));

  const upload = multer({ 
    dest: UPLOADS_DIR,
    limits: { fileSize: 10 * 1024 * 1024 } // 10MB per file
  });

  // Add a health check
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  // API Endpoints
  app.post("/api/generate", upload.fields([
    { name: "template", maxCount: 1 },
    { name: "images" }
  ]), async (req: any, res: any) => {
    console.log("Received generate request");
    try {
      const templateFile = req.files?.["template"]?.[0];
      const imageFiles = req.files?.["images"];

      if (!templateFile || !imageFiles || imageFiles.length === 0) {
        return res.status(400).json({ error: "缺少範本檔案或圖片" });
      }

      const templatePath = templateFile.path;
      const buffer = await fs.readFile(templatePath);

      // Inspect XML to find [圖片] placeholders and wrap the table in a loop
      const zip = await JSZip.loadAsync(buffer);
      
      let documentXml = await zip.file("word/document.xml")?.async("string") || "";
      
      // Function to replace text that might be split by XML tags
      const replaceSplitText = (xml: string, search: string, replacement: string) => {
        const pattern = search.split('').map(c => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('(?:<[^>]+>)*');
        return xml.replace(new RegExp(pattern, 'g'), replacement);
      };

      // 1. Normalize the [圖片] placeholder in the XML to ensure it's not split
      documentXml = replaceSplitText(documentXml, "[圖片]", "[圖片]");

      // 2. Find the table containing [圖片]
      const tableRegex = /<w:tbl>(?:(?!<w:tbl>).)*?\[圖片\].*?<\/w:tbl>/gs;
      const tableMatch = documentXml.match(tableRegex);

      if (tableMatch) {
         const targetTable = tableMatch[0];
         const tableIndex = documentXml.indexOf(targetTable);

         // Better header detection: find all paragraphs between the previous table and this one
         const precedingXml = documentXml.substring(0, tableIndex);
         const blocksBefore = precedingXml.split('</w:tbl>');
         const immediatePrecedingXml = blocksBefore[blocksBefore.length - 1];
         
         const pMatches = Array.from(immediatePrecedingXml.matchAll(/<w:p(?:(?!<w:p>).)*?<\/w:p>/gs));
         let headerXml = "";
         pMatches.forEach(m => headerXml += m[0]);

         console.log(`Found ${pMatches.length} header paragraphs to repeat.`);

         const fullBlockToRepeat = headerXml + targetTable;
         const placeholderRegex = /\[圖片\]/g;
         const k = (targetTable.match(placeholderRegex) || []).length || 1;

         // Extract column widths from tblGrid
         const colWidths: number[] = [];
         const gridMatch = targetTable.match(/<w:tblGrid>(.*?)<\/w:tblGrid>/s);
         if (gridMatch) {
            const cols = Array.from(gridMatch[1].matchAll(/<w:gridCol w:w="(\d+)"/g));
            cols.forEach(m => colWidths.push(parseInt(m[1])));
         }
         
         // Helper to find width of a specific placeholder's cell
         const getCellWidthForPlaceholder = (tableXml: string, placeholderIndex: number) => {
            const cells = tableXml.split(/<w:tc[ >]/);
            // First part is anything before first cell
            let count = -1;
            for (let i = 1; i < cells.length; i++) {
               if (cells[i].includes('[圖片]')) {
                  count++;
                  if (count === placeholderIndex) {
                     // Found the cell. Try to find tcW inside it.
                     const widthMatch = cells[i].match(/<w:tcW[^>]+w:w="(\d+)"/);
                     if (widthMatch) return parseInt(widthMatch[1]);
                     
                     // If no local tcW, we'd need to track column index...
                     // Let's fallback to the max grid width for now as it's most reliable for A4
                     return 0;
                  }
               }
            }
            return 0;
         };

         // Determine global max width if localized fails
         let globalMaxWidth = 9000;
         if (colWidths.length > 0) globalMaxWidth = Math.max(...colWidths);
         
         console.log("Detected Grid Widths:", colWidths, "Global Max:", globalMaxWidth);

         const numTablesNeeded = Math.ceil(imageFiles.length / k);
         let allGeneratedXml = "";
         const data: any = {};
         const pageBreakXml = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

         for (let i = 0; i < numTablesNeeded; i++) {
            let currentBlockXml = fullBlockToRepeat;
            
            if (i > 0) {
               currentBlockXml = pageBreakXml + currentBlockXml;
            }

            for (let j = 0; j < k; j++) {
               const imgIdx = i * k + j;
               const placeholderId = `img_${imgIdx}`;
               
               if (imgIdx < imageFiles.length) {
                  const img = imageFiles[imgIdx];
                  const imgData = await fs.readFile(img.path);
                  let dimensions;
                  try {
                     dimensions = sizeOf(imgData);
                  } catch (e) {
                     dimensions = { width: 100, height: 100 };
                  }
                  
                  const originalWidth = dimensions.width || 100;
                  const originalHeight = dimensions.height || 100;
                  
                  // Detect cell width specifically for THIS placeholder in the table
                  let cellWidthTwips = getCellWidthForPlaceholder(targetTable, j) || globalMaxWidth;
                  
                  // Ensure it's not too small (A4 is ~10000-11000 twips wide total)
                  // If it's very small (<2000), it's likely a label column accidentally matched
                  if (cellWidthTwips < 2000 && globalMaxWidth > 5000) {
                     cellWidthTwips = globalMaxWidth;
                  }

                  const cellWidthCm = (cellWidthTwips / 1440) * 2.54;
                  const targetWidth = cellWidthCm * 0.95;
                  const targetHeight = (originalHeight / originalWidth) * targetWidth;

                  data[placeholderId] = {
                     data: new Uint8Array(imgData),
                     extension: path.extname(img.originalname),
                     width: Math.round(targetWidth * 100) / 100,
                     height: Math.round(targetHeight * 100) / 100
                  };
                  currentBlockXml = currentBlockXml.replace(/\[圖片\]/, `{IMAGE ${placeholderId}}`);
               } else {
                  currentBlockXml = currentBlockXml.replace(/\[圖片\]/, "");
               }
            }
            allGeneratedXml += currentBlockXml;
         }

         // Use function to substitute replacements to avoid issues with special symbols in the generated XML
         documentXml = documentXml.replace(fullBlockToRepeat, () => allGeneratedXml);
         
         zip.file("word/document.xml", documentXml);
         const modifiedBuffer = await zip.generateAsync({ type: "uint8array" });

         console.log(`Processing with docx-templates, data keys: ${Object.keys(data).length}`);

         // Render
         const creator: any = typeof createReport === 'function' ? createReport : (createReport as any).default;
         const report = await creator({
            template: modifiedBuffer,
            data,
            cmdDelimiter: ['{', '}'],
            fixSpans: true,
         });

         const outputFileName = `generated_${Date.now()}.docx`;
         const outputPath = path.join(UPLOADS_DIR, outputFileName);
         await fs.writeFile(outputPath, Buffer.from(report));

         console.log(`Docx generated successfully: ${outputFileName}`);

         // Clean up uploads
         await fs.remove(templatePath).catch(() => {});
         if (imageFiles) {
           for (const f of imageFiles) await fs.remove(f.path).catch(() => {});
         }

         res.json({ downloadUrl: `/api/download/${outputFileName}` });
      } else {
         console.error("No [圖片] placeholder found in XML");
         res.status(400).json({ error: "找不到 [圖片] 標記，或標記格式不正確。請確保表格內有 [圖片] 字樣。" });
      }
    } catch (error: any) {
      console.error("Error generating docx:", error);
      // Ensure we send JSON even if it's an error
      res.status(500).json({ 
        error: error.message || "伺服器內部錯誤",
        stack: process.env.NODE_ENV !== 'production' ? error.stack : undefined 
      });
    }
  });

  app.get("/api/download/:filename", async (req, res) => {
    const filename = req.params.filename;
    // Basic security: prevent path traversal
    if (filename.includes("..") || filename.includes("/") || filename.includes("\\")) {
      return res.status(403).send("Forbidden");
    }
    const filePath = path.join(UPLOADS_DIR, filename);
    if (await fs.pathExists(filePath)) {
      res.download(filePath);
    } else {
      res.status(404).send("File not found");
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
