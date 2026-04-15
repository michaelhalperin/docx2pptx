/**
 * docx2pptx - Convert structured DOCX files to KodKode-branded PowerPoint presentations
 *
 * DOCX Format expected:
 *   שקופית N – <Title>
 *   • <bullet 1>  • <bullet 2>  ...
 *
 * Usage:
 *   node docx_to_pptx.js <input.docx> [output.pptx]
 *   node docx_to_pptx.js input.txt   [output.pptx]
 */

"use strict";

const fs   = require("fs");
const path = require("path");

// ─── KodKode Brand Theme ──────────────────────────────────────────────────────
const T = {
  white:       "FFFFFF",
  bg:          "FFFFFF",   // all slides are white
  purple:      "6C5CE7",   // primary — # badge, borders, accents
  teal:        "00BFA5",   // secondary — footer bar, definition borders
  tealLight:   "E0F7F4",   // teal tint for cards
  purpleLight: "EDE9FE",   // purple tint for cards
  textDark:    "1A1A2E",   // near-black body text
  textGray:    "6B7280",   // muted labels, subtitles
  textMid:     "374151",   // medium body
  codeBg:      "F5F5F5",   // light code background
  codeNum:     "AAAAAA",   // line number color
  borderLight: "E5E7EB",   // card/table borders
  footerText:  "1A1A2E",   // footer label
  headerGrad1: "6C5CE7",   // gradient bar start
  headerGrad2: "00BFA5",   // gradient bar end
};

// Slide dimensions (LAYOUT_16x9 = 10" × 5.625")
const W  = 10;
const H  = 5.625;
const FOOTER_H   = 0.42;
const FOOTER_Y   = H - FOOTER_H;
const HEADER_H   = 0.72;
const GRAD_H     = 0.055;  // thin top gradient strip
const CONTENT_Y  = HEADER_H + 0.15;
const CONTENT_H  = FOOTER_Y - CONTENT_Y - 0.1;
const LOGO_W     = 1.55;
const LOGO_H     = 0.52;
const LOGO_X     = W - LOGO_W - 0.18;
const LOGO_Y     = 0.10;

// ─── Shared logo renderer ────────────────────────────────────────────────────
// Drawn as two independent text boxes (one per row) so font metrics don't merge.
// We avoid Arial Black since LibreOffice renders it poorly — use bold Arial instead.
function addLogo(pres, s) {
  const bx = LOGO_X - 0.08;
  const by = LOGO_Y - 0.05;
  const bw = LOGO_W + 0.12;
  const bh = LOGO_H + 0.10;

  // Card background
  s.addShape(pres.shapes.RECTANGLE, {
    x: bx, y: by, w: bw, h: bh,
    fill: { color: "F8F8F8" }, line: { color: T.borderLight, width: 1 },
  });

  const rowH  = bh / 2 - 0.01;
  const textX = bx + 0.08;
  const textW = 0.62;

  // Row 1: "KOD" — black bold
  s.addText("KOD", {
    x: textX, y: by + 0.03, w: textW, h: rowH,
    fontSize: 15, bold: true, color: "111111",
    fontFace: "Arial", align: "left", valign: "middle", margin: 0,
  });

  // Row 2: "K" teal  +  "ODE" black
  s.addText([
    { text: "K",   options: { color: T.teal,   bold: true, fontSize: 15 } },
    { text: "ODE", options: { color: "111111", bold: true, fontSize: 15 } },
  ], {
    x: textX, y: by + rowH + 0.03, w: textW, h: rowH,
    fontFace: "Arial", align: "left", valign: "middle", margin: 0,
  });

  // Hebrew subtitle — right column inside the card
  s.addText([
    { text: "התוכנית החרדית", options: { breakLine: true } },
    { text: "ליחידות הייטק",  options: { breakLine: true } },
    { text: "במערכת הביטחון", options: {} },
  ], {
    x: bx + 0.74, y: by + 0.04, w: bw - 0.80, h: bh - 0.08,
    fontSize: 5.5, color: T.textGray, fontFace: "Arial",
    align: "right", valign: "middle", margin: 0, lineSpacingMultiple: 1.2,
    rtlMode: true,
  });
}

// ─── Shared: header + footer drawn on every slide ────────────────────────────

/**
 * Draws the common KodKode header and footer on a slide.
 * title     — slide title string  (null for title slide)
 * hashBadge — true = draw purple # square badge (not on title slide)
 */
function addChrome(pres, s, title, hashBadge = true) {
  // ── Top gradient strip ──────────────────────────────────────────────────
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W / 2, h: GRAD_H,
    fill: { color: T.purple }, line: { color: T.purple, width: 0 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: W / 2, y: 0, w: W / 2, h: GRAD_H,
    fill: { color: T.teal }, line: { color: T.teal, width: 0 },
  });

  // ── KodKode logo ────────────────────────────────────────────────────────
  addLogo(pres, s);

  // ── Header area (below gradient strip) ─────────────────────────────────
  if (hashBadge && title) {
    // Purple # square badge
    const badgeW = 0.46;
    const badgeH = 0.46;
    const badgeY = GRAD_H + (HEADER_H - GRAD_H - badgeH) / 2;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.22, y: badgeY, w: badgeW, h: badgeH,
      fill: { color: T.purple }, line: { color: T.purple, width: 0 },
    });
    s.addText("#", {
      x: 0.22, y: badgeY, w: badgeW, h: badgeH,
      fontSize: 18, bold: true, color: T.white,
      fontFace: "Arial", align: "center", valign: "middle", margin: 0,
    });

    // Title text (gray, not bold, large)
    s.addText(title, {
      x: 0.84, y: GRAD_H + 0.05, w: LOGO_X - 0.95, h: HEADER_H - GRAD_H - 0.05,
      fontSize: 22, bold: false, color: T.textGray,
      fontFace: "Arial", valign: "middle", margin: 0,
    });

    // Thin horizontal rule under header
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: HEADER_H, w: W, h: 0.012,
      fill: { color: T.borderLight }, line: { color: T.borderLight, width: 0 },
    });
  }

  // ── Footer ──────────────────────────────────────────────────────────────
  // White bar background
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: FOOTER_Y, w: W, h: FOOTER_H,
    fill: { color: T.white }, line: { color: T.borderLight, width: 1 },
  });
  // Teal left accent block
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: FOOTER_Y, w: 0.18, h: FOOTER_H,
    fill: { color: T.teal }, line: { color: T.teal, width: 0 },
  });
  // Footer text
  s.addText("קודקוד — התוכנית החרדית ליחידות הייטק במערכת הביטחון", {
    x: 0.28, y: FOOTER_Y, w: W - 0.36, h: FOOTER_H,
    fontSize: 10, color: T.footerText, fontFace: "Arial",
    align: "center", valign: "middle", margin: 0, rtlMode: true,
  });
}

// ─── Slide type detection ────────────────────────────────────────────────────
function classifySlide(title, bullets, index, total) {
  const t = (title || "").trim();
  if (index === 0) return "TITLE";
  if (index === total - 1) return "TAKEAWAYS";
  if (/סיכום|summary/i.test(t)) return "SUMMARY";
  if (/takeaway/i.test(t)) return "TAKEAWAYS";
  if (/קוד|code|דוגמת קוד/i.test(t)) return "CODE";
  if (bullets.length === 0) return "CONCEPT";
  return "BULLETS";
}

// ─── Parse DOCX/TXT ──────────────────────────────────────────────────────────
async function parseDocx(filePath) {
  const ext = path.extname(filePath).toLowerCase();

  if (ext === ".docx") {
    // Use HTML output from mammoth — preserves bold/heading vs bullet structure
    const mammoth = require("mammoth");
    const result  = await mammoth.convertToHtml({ path: filePath });
    return parseHtml(result.value);
  } else {
    const rawText = fs.readFileSync(filePath, "utf8");
    return parseRawText(rawText);
  }
}

/**
 * Parse mammoth HTML output.
 * A slide starts at every <p><strong>…</strong></p> (bold paragraph = title).
 * Bullet points come from <ul><li>…</li></ul> blocks.
 * Plain <p> after a title (no bold) = code lines, collected as bullets.
 */
function parseHtml(html) {
  const slides  = [];
  let current   = null;

  // Tokenise: split on tags we care about
  // Each token is either a tag or text content
  const tokens = html.split(/(<\/?(?:p|ul|li|strong|h[1-6])[^>]*>)/i);

  let inStrong = false;
  let inLi     = false;
  let inP      = false;
  let pText    = "";
  let liText   = "";

  function flush() {
    // Called when a paragraph or list item ends
  }

  // Simple state-machine parser
  let i = 0;
  while (i < tokens.length) {
    const tok = tokens[i];

    if (/^<strong>/i.test(tok))      { inStrong = true; }
    else if (/^<\/strong>/i.test(tok)){ inStrong = false; }
    else if (/^<p[^>]*>/i.test(tok)) { inP = true; pText = ""; }
    else if (/^<\/p>/i.test(tok)) {
      const t = stripTags(pText).trim();
      if (t) {
        if (inStrong || pWasBold) {
          // Bold paragraph = new slide title
          if (current) slides.push(current);
          current = { title: t, bullets: [] };
        } else if (current) {
          // Plain paragraph under a slide = code line / extra bullet
          current.bullets.push(t);
        } else {
          current = { title: t, bullets: [] };
        }
      }
      inP = false; pText = ""; pWasBold = false;
    }
    else if (/^<li[^>]*>/i.test(tok)) { inLi = true; liText = ""; }
    else if (/^<\/li>/i.test(tok)) {
      const t = stripTags(liText).trim();
      if (t && current) current.bullets.push(t);
      inLi = false; liText = "";
    }
    else if (/^<h[1-6]/i.test(tok)) {
      // Headings also act as slide titles — collect text until </hN>
    }
    else {
      // Text node
      const text = decodeHtmlEntities(tok);
      if (inLi)       liText += text;
      else if (inP) {
        pText += text;
        if (inStrong) pWasBold = true;
      }
    }
    i++;
  }
  if (current) slides.push(current);

  slides.forEach((s, idx) => (s.index = idx));
  return slides;
}

// Tracks whether the current <p> contained a <strong> child
let pWasBold = false;

function stripTags(str) {
  return str.replace(/<[^>]*>/g, "");
}

function decodeHtmlEntities(str) {
  return str
    .replace(/&amp;/g,  "&")
    .replace(/&lt;/g,   "<")
    .replace(/&gt;/g,   ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g,  "'")
    .replace(/&nbsp;/g, " ");
}

/** Plain-text format (for .txt files) */
function parseRawText(text) {
  const lines = text.split(/\r?\n/);

  // Format A: "שקופית N – Title" explicit markers
  const hasExplicitMarkers = lines.some(l =>
    /^שקופית\s+\d+\s*[–\-:]/u.test(l.trim())
  );

  if (hasExplicitMarkers) {
    return parseFormatA(lines);
  } else {
    return parseFormatB(lines);
  }
}

function parseFormatA(lines) {
  const slides = [];
  let current  = null;
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;
    const m = line.match(/^שקופית\s+(\d+)\s*[–\-:]\s*(.+)/u);
    if (m) {
      if (current) slides.push(current);
      current = { title: m[2].trim(), bullets: [] };
      continue;
    }
    if (!current) { current = { title: line, bullets: [] }; continue; }
    if (line.includes("•")) {
      line.split("•").map(s => s.trim()).filter(Boolean)
          .forEach(b => current.bullets.push(b));
    } else {
      current.bullets.push(line);
    }
  }
  if (current) slides.push(current);
  slides.forEach((s, i) => (s.index = i));
  return slides;
}

function parseFormatB(lines) {
  const blocks = [];
  let block = [];
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) { if (block.length > 0) { blocks.push(block); block = []; } }
    else block.push(line);
  }
  if (block.length > 0) blocks.push(block);
  return blocks
    .filter(b => b.length > 0)
    .map((b, i) => ({ index: i, title: b[0], bullets: b.slice(1).filter(Boolean) }));
}

// ─── Slide 1: Title ──────────────────────────────────────────────────────────
async function buildTitleSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: T.white };

  // Top gradient strip
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: W / 2, h: GRAD_H,
    fill: { color: T.purple }, line: { color: T.purple, width: 0 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: W / 2, y: 0, w: W / 2, h: GRAD_H,
    fill: { color: T.teal }, line: { color: T.teal, width: 0 },
  });

  // KodKode logo
  addLogo(pres, s);

  // Left purple vertical bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.32, y: 1.1, w: 0.07, h: 1.9,
    fill: { color: T.purple }, line: { color: T.purple, width: 0 },
  });

  // Main title (large, gray, RTL)
  s.addText(slide.title, {
    x: 0.55, y: 1.1, w: 6.8, h: 1.6,
    fontSize: 38, bold: true, color: "AAAAAA",
    fontFace: "Arial", valign: "middle", margin: 0,
    rtlMode: true,
  });

  // Subtopic pills / bullets row
  const subs = slide.bullets.slice(0, 5);
  if (subs.length > 0) {
    s.addText(subs.join("  ·  "), {
      x: 0.55, y: 3.1, w: 8.5, h: 0.45,
      fontSize: 13, color: T.textGray, fontFace: "Arial",
      valign: "middle", margin: 0, italic: true, rtlMode: true,
    });
  }

  // Footer
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: FOOTER_Y, w: W, h: FOOTER_H,
    fill: { color: T.white }, line: { color: T.borderLight, width: 1 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: FOOTER_Y, w: 0.18, h: FOOTER_H,
    fill: { color: T.teal }, line: { color: T.teal, width: 0 },
  });
  s.addText("קודקוד — התוכנית החרדית ליחידות הייטק במערכת הביטחון", {
    x: 0.28, y: FOOTER_Y, w: W - 0.36, h: FOOTER_H,
    fontSize: 10, color: T.footerText, fontFace: "Arial",
    align: "center", valign: "middle", margin: 0, rtlMode: true,
  });
}

// ─── Slide: General Bullets (third screenshot layout) ───────────────────────
async function buildBulletsSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: T.white };
  addChrome(pres, s, slide.title);

  const bullets = slide.bullets.slice(0, 6);
  if (bullets.length === 0) return s;

  const availH = FOOTER_Y - CONTENT_Y - 0.1;
  const gap    = 0.13;
  const cardH  = Math.min(0.78, (availH - gap * (bullets.length - 1)) / bullets.length);
  const totalH = bullets.length * cardH + (bullets.length - 1) * gap;
  const startY = CONTENT_Y + (availH - totalH) / 2;

  bullets.forEach((bullet, i) => {
    const y = startY + i * (cardH + gap);

    // Card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 9.3, h: cardH,
      fill: { color: T.white },
      line: { color: T.borderLight, width: 1 },
      shadow: { type: "outer", color: "000000", blur: 3, offset: 1, angle: 135, opacity: 0.05 },
    });

    // Left accent stripe — alternating purple / teal
    const stripe = i % 2 === 0 ? T.purple : T.teal;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 0.055, h: cardH,
      fill: { color: stripe }, line: { color: stripe, width: 0 },
    });

    // Bullet text
    s.addText(bullet, {
      x: 0.55, y: y + 0.04, w: 8.9, h: cardH - 0.08,
      fontSize: 14, color: T.textDark, fontFace: "Arial",
      valign: "middle", margin: 0, rtlMode: true,
    });
  });

  return s;
}

// ─── Slide: Code (second screenshot — split layout) ─────────────────────────
async function buildCodeSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: T.white };
  addChrome(pres, s, slide.title);

  const splitX  = 0.35;
  const codeW   = 4.85;
  const defX    = splitX + codeW + 0.2;
  const defW    = W - defX - 0.35;
  const areaY   = CONTENT_Y + 0.05;
  const areaH   = FOOTER_Y - areaY - 0.12;

  // ── Left: code panel ──────────────────────────────────────────────────
  // Window chrome bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: splitX, y: areaY, w: codeW, h: 0.32,
    fill: { color: "EEEEEE" }, line: { color: T.borderLight, width: 1 },
  });
  // Traffic lights
  ["FF5F57","FEBC2E","28C840"].forEach((c, i) => {
    s.addShape(pres.shapes.OVAL, {
      x: splitX + 0.16 + i * 0.24, y: areaY + 0.09, w: 0.14, h: 0.14,
      fill: { color: c }, line: { color: c, width: 0 },
    });
  });
  // Filename
  s.addText(slide.title.replace(/דוגמת קוד\s*/i,"").trim() + ".js" || "main.js", {
    x: splitX + 1.0, y: areaY + 0.02, w: codeW - 1.2, h: 0.28,
    fontSize: 10, color: T.textGray, fontFace: "Consolas",
    align: "left", valign: "middle", margin: 0,
  });

  // Code body
  s.addShape(pres.shapes.RECTANGLE, {
    x: splitX, y: areaY + 0.32, w: codeW, h: areaH - 0.32,
    fill: { color: T.codeBg }, line: { color: T.borderLight, width: 1 },
  });

  // ── Code lines: use actual slide bullets if present, else template ──────
  // Each bullet = one line of code
  const lineH = 0.235;
  const hasRealCode = slide.bullets.length > 0;
  const codeLines = hasRealCode
    ? slide.bullets.map(b => [{ text: b, options: { color: "333333" } }])
    : buildCodeContent(slide.title);

  codeLines.forEach((lineTokens, li) => {
    const ly = areaY + 0.32 + 0.08 + li * lineH;
    if (ly + lineH > FOOTER_Y - 0.1) return;

    // Line number
    s.addText(String(li + 1), {
      x: splitX + 0.06, y: ly, w: 0.28, h: lineH,
      fontSize: 9.5, color: T.codeNum, fontFace: "Consolas",
      align: "right", valign: "top", margin: 0,
    });

    // Code tokens
    s.addText(lineTokens, {
      x: splitX + 0.38, y: ly, w: codeW - 0.44, h: lineH,
      fontSize: 10.5, fontFace: "Consolas",
      align: "left", valign: "top", margin: 0,
    });
  });

  // ── Right: definition box (teal border, purple label) ────────────────
  s.addShape(pres.shapes.RECTANGLE, {
    x: defX, y: areaY, w: defW, h: areaH,
    fill: { color: T.white }, line: { color: T.purple, width: 2 },
  });

  // "הגדרה" label bar at top of definition box
  s.addShape(pres.shapes.RECTANGLE, {
    x: defX, y: areaY, w: defW, h: 0.34,
    fill: { color: T.purpleLight }, line: { color: T.purple, width: 0 },
  });
  s.addText("הגדרה", {
    x: defX + 0.12, y: areaY + 0.02, w: defW - 0.24, h: 0.30,
    fontSize: 12, bold: true, color: T.purple, fontFace: "Arial",
    valign: "middle", margin: 0, rtlMode: true,
  });

  // Definition box is intentionally left blank — to be filled manually in PowerPoint
  s.addText("הכנס כאן את ההגדרה בעברית...", {
    x: defX + 0.12, y: areaY + 0.42, w: defW - 0.24, h: areaH - 0.54,
    fontSize: 12, color: "CCCCCC", fontFace: "Arial",
    valign: "top", margin: 0, rtlMode: true, wrap: true,
    italic: true,
  });

  return s;
}

// Generate code lines as pptxgenjs rich-text arrays
function buildCodeContent(title) {
  const t = title.toLowerCase();

  const kw  = (txt) => ({ text: txt, options: { color: "C792EA" } });  // keyword
  const fn  = (txt) => ({ text: txt, options: { color: "61AFEF" } });  // function
  const str = (txt) => ({ text: txt, options: { color: "98C379" } });  // string
  const cm  = (txt) => ({ text: txt, options: { color: "7F848E", italic: true } }); // comment
  const pl  = (txt) => ({ text: txt, options: { color: "333333" } });  // plain

  if (t.includes("בסיסי") || t.includes("basic")) {
    return [
      [kw("import "), pl("{ "), fn("useEffect"), pl(", "), fn("useState"), pl(" } "), kw("from "), str("'react'")],
      [pl("")],
      [kw("function "), fn("UsersList"), pl("() {")],
      [pl("  "), kw("const "), pl("[users, setUsers] = "), fn("useState"), pl("([]);")],
      [pl("")],
      [pl("  "), fn("useEffect"), pl("(() => {")],
      [pl("    "), fn("fetch"), pl("("), str("'/api/users'"), pl(")")],
      [pl("      .then(r => r."), fn("json"), pl("())")],
      [pl("      .then("), fn("setUsers"), pl(");")],
      [pl("  }, []);"), cm("  // ← רץ פעם אחת בטעינה")],
      [pl("}")],
    ];
  }
  if (t.includes("תלות") || t.includes("depend")) {
    return [
      [fn("useEffect"), pl("(() => {")],
      [pl("  "), fn("fetchUser"), pl("(userId);")],
      [pl("}, [userId]);"), cm("  // ← רץ מחדש כשמשתנה userId")],
      [pl("")],
      [cm("// ✅ תלויות נכונות:")],
      [pl("}, [userId, token]);")],
      [pl("")],
      [cm("// ❌ תלות חסרה (באג!):")],
      [pl("}, []);"), cm("  // token לא יתעדכן!")],
    ];
  }
  if (t.includes("cleanup")) {
    return [
      [fn("useEffect"), pl("(() => {")],
      [pl("  "), kw("const "), pl("timer = "), fn("setInterval"), pl("(() => {")],
      [pl("    console."), fn("log"), pl("("), str("'tick'"), pl(");")],
      [pl("  }, 1000);")],
      [pl("")],
      [pl("  "), kw("return "), pl("() => "), fn("clearInterval"), pl("(timer);")],
      [pl("}, []);"), cm("  // ← cleanup מונע memory leak")],
    ];
  }
  // Generic
  return [
    [fn("useEffect"), pl("(() => {")],
    [pl("  "), cm("// your code here")],
    [pl("}, []);")],
  ];
}

// ─── Slide: Concept (title-only — no bullets) ────────────────────────────────
async function buildConceptSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: T.white };
  addChrome(pres, s, slide.title);

  // Large centered concept text box with teal border
  const bx = 1.0, by = CONTENT_Y + 0.3, bw = W - 2.0, bh = FOOTER_Y - by - 0.5;
  s.addShape(pres.shapes.RECTANGLE, {
    x: bx, y: by, w: bw, h: bh,
    fill: { color: T.tealLight }, line: { color: T.teal, width: 2 },
  });
  s.addText(slide.title, {
    x: bx + 0.2, y: by + 0.1, w: bw - 0.4, h: bh - 0.2,
    fontSize: 32, bold: true, color: T.textDark, fontFace: "Arial",
    align: "center", valign: "middle", margin: 0, rtlMode: true,
  });

  return s;
}

// ─── Slide: Text Heavy ───────────────────────────────────────────────────────
async function buildTextHeavySlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: T.white };
  addChrome(pres, s, slide.title);

  const bullets = slide.bullets.slice(0, 4);
  const gap  = 0.15;
  const bH   = (FOOTER_Y - CONTENT_Y - 0.15 - gap * (bullets.length - 1)) / bullets.length;

  bullets.forEach((b, i) => {
    const y = CONTENT_Y + 0.05 + i * (bH + gap);
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 9.3, h: bH,
      fill: { color: i % 2 === 0 ? T.purpleLight : T.tealLight },
      line: { color: i % 2 === 0 ? T.purple : T.teal, width: 1 },
    });
    s.addText(b, {
      x: 0.55, y: y + 0.06, w: 8.9, h: bH - 0.12,
      fontSize: 14, color: T.textDark, fontFace: "Arial",
      valign: "middle", margin: 0, rtlMode: true, wrap: true,
    });
  });

  return s;
}

// ─── Slide: Summary ──────────────────────────────────────────────────────────
async function buildSummarySlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: T.white };
  addChrome(pres, s, slide.title);

  const bullets = slide.bullets.slice(0, 3);
  const cols    = bullets.length || 1;
  const gap     = 0.2;
  const cardW   = (9.3 - gap * (cols - 1)) / cols;
  const cardY   = CONTENT_Y + 0.1;
  const cardH   = FOOTER_Y - cardY - 0.12;

  bullets.forEach((b, i) => {
    const x      = 0.35 + i * (cardW + gap);
    const stripe = i % 2 === 0 ? T.purple : T.teal;
    const tint   = i % 2 === 0 ? T.purpleLight : T.tealLight;

    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cardY, w: cardW, h: cardH,
      fill: { color: T.white }, line: { color: stripe, width: 2 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cardY, w: cardW, h: 0.1,
      fill: { color: stripe }, line: { color: stripe, width: 0 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: cardY + 0.1, w: cardW, h: 0.36,
      fill: { color: tint }, line: { color: tint, width: 0 },
    });
    s.addText(String(i + 1), {
      x: x + 0.1, y: cardY + 0.13, w: 0.32, h: 0.30,
      fontSize: 16, bold: true, color: stripe, fontFace: "Arial",
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(b, {
      x: x + 0.14, y: cardY + 0.52, w: cardW - 0.28, h: cardH - 0.64,
      fontSize: 13, color: T.textDark, fontFace: "Arial",
      valign: "top", margin: 0, rtlMode: true, wrap: true,
      lineSpacingMultiple: 1.35,
    });
  });

  return s;
}

// ─── Slide: Takeaways ────────────────────────────────────────────────────────
async function buildTakeawaysSlide(pres, slide) {
  const s = pres.addSlide();
  s.background = { color: T.white };
  addChrome(pres, s, slide.title);

  const bullets  = slide.bullets.slice(0, 5);
  const gap      = 0.12;
  const availH   = FOOTER_Y - CONTENT_Y - 0.1;
  const cardH    = Math.min(0.75, (availH - gap * (bullets.length - 1)) / bullets.length);
  const totalH   = bullets.length * cardH + (bullets.length - 1) * gap;
  const startY   = CONTENT_Y + (availH - totalH) / 2;

  bullets.forEach((b, i) => {
    const y      = startY + i * (cardH + gap);
    const stripe = i % 2 === 0 ? T.purple : T.teal;

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 9.3, h: cardH,
      fill: { color: T.white }, line: { color: T.borderLight, width: 1 },
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.35, y, w: 0.055, h: cardH,
      fill: { color: stripe }, line: { color: stripe, width: 0 },
    });
    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: 0.52, y: y + (cardH - 0.38) / 2, w: 0.38, h: 0.38,
      fill: { color: stripe }, line: { color: stripe, width: 0 },
    });
    s.addText(String(i + 1), {
      x: 0.52, y: y + (cardH - 0.38) / 2, w: 0.38, h: 0.38,
      fontSize: 12, bold: true, color: T.white, fontFace: "Arial",
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(b, {
      x: 1.06, y: y + 0.06, w: 8.45, h: cardH - 0.12,
      fontSize: 14, color: T.textDark, fontFace: "Arial",
      valign: "middle", margin: 0, rtlMode: true,
    });
  });

  return s;
}

// ─── Main ────────────────────────────────────────────────────────────────────
async function convert(inputPath, outputPath) {
  console.log(`📄 Parsing: ${inputPath}`);
  const slides = await parseDocx(inputPath);
  console.log(`✅ Found ${slides.length} slides`);

  const PptxGenJS = require("pptxgenjs");
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];
    const type  = classifySlide(slide.title, slide.bullets, i, slides.length);
    console.log(`  [${i + 1}/${slides.length}] "${slide.title}" → ${type}`);

    if      (type === "TITLE")      await buildTitleSlide(pres, slide);
    else if (type === "CODE")       await buildCodeSlide(pres, slide);
    else if (type === "SUMMARY")    await buildSummarySlide(pres, slide);
    else if (type === "TAKEAWAYS")  await buildTakeawaysSlide(pres, slide);
    else if (type === "TEXT_HEAVY") await buildTextHeavySlide(pres, slide);
    else if (type === "CONCEPT")    await buildConceptSlide(pres, slide);
    else                            await buildBulletsSlide(pres, slide);
  }

  console.log(`💾 Writing: ${outputPath}`);
  await pres.writeFile({ fileName: outputPath });
  console.log("✅ Done!");
}

// ─── CLI ─────────────────────────────────────────────────────────────────────
(async () => {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.error("Usage: node docx_to_pptx.js <input.docx|input.txt> [output.pptx]");
    process.exit(1);
  }
  const input  = args[0];
  const output = args[1] || input.replace(/\.(docx|txt)$/i, ".pptx");
  try {
    await convert(input, output);
  } catch (e) {
    console.error("Error:", e.message);
    console.error(e.stack);
    process.exit(1);
  }
})();
