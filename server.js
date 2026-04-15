/**
 * docx2pptx — local web server
 * Run: node server.js
 * Then open: http://localhost:4242
 */

"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");
const os = require("os");
const { execSync, spawn } = require("child_process");

const PORT = Number(process.env.PORT) || 4242;
const UI = path.join(__dirname, "ui.html");

// ── Tiny multipart parser (no dependencies) ──────────────────────────────────
function parseMultipart(body, boundary) {
  const parts = {};
  const bnd = Buffer.from("--" + boundary);
  let pos = 0;

  while (pos < body.length) {
    const start = indexOf(body, bnd, pos);
    if (start === -1) break;
    pos = start + bnd.length + 2; // skip \r\n

    // headers
    const headerEnd = indexOf(body, Buffer.from("\r\n\r\n"), pos);
    if (headerEnd === -1) break;
    const headerStr = body.slice(pos, headerEnd).toString();
    pos = headerEnd + 4;

    // find next boundary
    const nextBnd = indexOf(body, bnd, pos);
    if (nextBnd === -1) break;
    const data = body.slice(pos, nextBnd - 2); // trim trailing \r\n
    pos = nextBnd;

    // parse content-disposition
    const nameMatch = headerStr.match(/name="([^"]+)"/);
    const fileMatch = headerStr.match(/filename="([^"]+)"/);
    if (!nameMatch) continue;

    parts[nameMatch[1]] = {
      data,
      filename: fileMatch ? fileMatch[1] : null,
    };
  }
  return parts;
}

/** RFC 5987 + ASCII fallback — Node rejects header bytes > 255 in plain filename="..." */
function contentDispositionAttachment(filename) {
  const safe = filename.replace(/[\r\n"]/g, "").replace(/[^\x20-\x7E]/g, "_");
  const star = encodeURIComponent(filename);
  return `attachment; filename="${safe}"; filename*=UTF-8''${star}`;
}

function indexOf(buf, search, start = 0) {
  for (let i = start; i <= buf.length - search.length; i++) {
    let found = true;
    for (let j = 0; j < search.length; j++) {
      if (buf[i + j] !== search[j]) {
        found = false;
        break;
      }
    }
    if (found) return i;
  }
  return -1;
}

// ── Read body ─────────────────────────────────────────────────────────────────
function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

// ── HTTP server ───────────────────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  const url = req.url.split("?")[0];

  // ── Serve UI ────────────────────────────────────────────────────────────────
  if (req.method === "GET" && url === "/") {
    res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
    return res.end(fs.readFileSync(UI));
  }

  // ── Convert endpoint ────────────────────────────────────────────────────────
  if (req.method === "POST" && url === "/convert") {
    const ct = req.headers["content-type"] || "";
    const bnd = ct.match(/boundary=(.+)/)?.[1];

    if (!bnd) {
      res.writeHead(400);
      return res.end("Missing boundary");
    }

    let body;
    try {
      body = await readBody(req);
    } catch (e) {
      res.writeHead(500);
      return res.end("Read error");
    }

    const parts = parseMultipart(body, bnd);
    const file = parts["file"];

    if (!file || !file.filename) {
      res.writeHead(400);
      return res.end("No file uploaded");
    }

    const ext = path.extname(file.filename).toLowerCase();
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "docx2pptx-"));
    const inPath = path.join(tmpDir, "input" + ext);
    const outPath = path.join(tmpDir, "output.pptx");

    fs.writeFileSync(inPath, file.data);

    const scriptPath = path.join(__dirname, "docx_to_pptx.js");

    let log = "";
    try {
      log = execSync(`node "${scriptPath}" "${inPath}" "${outPath}"`, {
        cwd: __dirname,
        encoding: "utf8",
        timeout: 120_000,
      });
    } catch (e) {
      log = (e.stdout || "") + (e.stderr || "") + e.message;
      // cleanup
      fs.rmSync(tmpDir, { recursive: true, force: true });
      res.writeHead(500, { "Content-Type": "application/json" });
      return res.end(JSON.stringify({ error: log }));
    }

    if (!fs.existsSync(outPath)) {
      fs.rmSync(tmpDir, { recursive: true, force: true });
      res.writeHead(500, { "Content-Type": "application/json" });
      return res.end(
        JSON.stringify({ error: "Output file not created.\n" + log }),
      );
    }

    const pptxData = fs.readFileSync(outPath);
    const outName = file.filename.replace(/\.(docx|txt)$/i, ".pptx");

    fs.rmSync(tmpDir, { recursive: true, force: true });

    res.writeHead(200, {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-Disposition": contentDispositionAttachment(outName),
      "Content-Length": pptxData.length,
      "X-Log": encodeURIComponent(log.slice(0, 500)),
    });
    return res.end(pptxData);
  }

  // ── 404 ─────────────────────────────────────────────────────────────────────
  res.writeHead(404);
  res.end("Not found");
});

server.listen(PORT, "0.0.0.0", () => {
  console.log(`\n✅  docx2pptx running on port ${PORT}\n`);

  // Only auto-open browser for local development sessions.
  if (!process.env.RENDER) {
    const url = `http://localhost:${PORT}`;
    try {
      const platform = process.platform;
      if (platform === "darwin") execSync(`open "${url}"`);
      else if (platform === "win32") execSync(`start "" "${url}"`);
      else execSync(`xdg-open "${url}" 2>/dev/null || true`);
    } catch (_) {
      /* user can open manually */
    }
  }
});
