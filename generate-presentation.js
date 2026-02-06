const pptxgen = require("pptxgenjs");

// ── Color Palette (Midnight Executive inspired, tech-forward) ──
const C = {
  darkNavy:   "0F1B2D",
  navy:       "1B2A4A",
  steelBlue:  "2D6A8F",
  teal:       "00B4D8",
  mint:       "00E5A0",
  lightGray:  "EDF2F7",
  offWhite:   "F7FAFC",
  white:      "FFFFFF",
  charcoal:   "2D3748",
  medGray:    "718096",
  subtleGray: "A0AEC0",
  errorRed:   "FC8181",
  warmYellow: "F6E05E",
};

// ── Helper Factories (fresh objects each call — avoids PptxGenJS mutation bug) ──
const makeShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.18 });
const makeCardShadow = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.12 });

// ── Create Presentation ──
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "GitHub Copilot";
pres.title = "Node.js Calculator Application";

// ═══════════════════════════════════════════
// SLIDE 1 — Title Slide
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.darkNavy };

  // Decorative top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }
  });

  // Large geometric accent shape (right side)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 7.5, y: 1.0, w: 3.5, h: 3.5,
    fill: { color: C.steelBlue, transparency: 15 },
    rotate: 45
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 8.2, y: 1.5, w: 2.5, h: 2.5,
    fill: { color: C.teal, transparency: 30 },
    rotate: 45
  });

  // Title
  slide.addText("Node.js Calculator", {
    x: 0.8, y: 1.4, w: 7, h: 1.2,
    fontSize: 44, fontFace: "Trebuchet MS", color: C.white,
    bold: true, charSpacing: 2
  });

  // Subtitle
  slide.addText("A Modern Web-Based Arithmetic Application", {
    x: 0.8, y: 2.6, w: 6.5, h: 0.6,
    fontSize: 20, fontFace: "Calibri", color: C.teal,
    italic: true
  });

  // Divider line
  slide.addShape(pres.shapes.LINE, {
    x: 0.8, y: 3.4, w: 3.5, h: 0,
    line: { color: C.mint, width: 2.5 }
  });

  // Metadata
  slide.addText([
    { text: "Built with Express.js  |  RESTful API  |  State Machine UI", options: { breakLine: true } },
    { text: "Client-Server Architecture  |  Full Test Coverage" }
  ], {
    x: 0.8, y: 3.7, w: 6, h: 0.8,
    fontSize: 13, fontFace: "Calibri", color: C.subtleGray
  });

  // Bottom bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.35, w: 10, h: 0.28, fill: { color: C.navy }
  });
  slide.addText("February 2026", {
    x: 0.8, y: 5.35, w: 3, h: 0.28,
    fontSize: 10, fontFace: "Calibri", color: C.medGray
  });
}

// ═══════════════════════════════════════════
// SLIDE 2 — Project Overview
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title area with colored background
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.darkNavy }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.04, fill: { color: C.teal }
  });
  slide.addText("Project Overview", {
    x: 0.7, y: 0.2, w: 8, h: 0.75,
    fontSize: 30, fontFace: "Trebuchet MS", color: C.white, bold: true
  });

  // Description text
  slide.addText(
    "The Node.js Calculator is a web-based application providing basic arithmetic operations through a clean, browser-based interface. It follows a client-server architecture where the client manages user input via a state machine, and the server performs calculations through a RESTful API.",
    {
      x: 0.7, y: 1.4, w: 8.6, h: 0.9,
      fontSize: 13, fontFace: "Calibri", color: C.charcoal,
      lineSpacingMultiple: 1.3
    }
  );

  // Feature cards — 2x2 grid
  const features = [
    { title: "Arithmetic Operations", desc: "Addition, subtraction,\nmultiplication, and division", icon: "+" },
    { title: "Responsive Web UI", desc: "Browser-based calculator\ninterface with keyboard support", icon: "UI" },
    { title: "RESTful API", desc: "Clean API endpoint for\ncalculation operations", icon: "{ }" },
    { title: "Input Validation", desc: "Server-side validation\nfor all operands & operations", icon: "!" },
  ];

  const cardW = 4.05, cardH = 1.2, startX = 0.7, startY = 2.55, gapX = 0.2, gapY = 0.2;

  features.forEach((f, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const cx = startX + col * (cardW + gapX);
    const cy = startY + row * (cardH + gapY);

    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW, h: cardH,
      fill: { color: C.white },
      shadow: makeCardShadow()
    });

    // Left accent bar
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: 0.06, h: cardH,
      fill: { color: C.teal }
    });

    // Icon circle
    slide.addShape(pres.shapes.OVAL, {
      x: cx + 0.2, y: cy + 0.25, w: 0.7, h: 0.7,
      fill: { color: C.teal, transparency: 15 }
    });
    slide.addText(f.icon, {
      x: cx + 0.2, y: cy + 0.25, w: 0.7, h: 0.7,
      fontSize: 16, fontFace: "Consolas", color: C.steelBlue,
      align: "center", valign: "middle", bold: true
    });

    // Title
    slide.addText(f.title, {
      x: cx + 1.05, y: cy + 0.12, w: 2.8, h: 0.35,
      fontSize: 14, fontFace: "Trebuchet MS", color: C.darkNavy, bold: true,
      margin: 0
    });

    // Description
    slide.addText(f.desc, {
      x: cx + 1.05, y: cy + 0.5, w: 2.8, h: 0.6,
      fontSize: 11, fontFace: "Calibri", color: C.medGray,
      margin: 0
    });
  });
}

// ═══════════════════════════════════════════
// SLIDE 3 — Technology Stack
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.darkNavy }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.04, fill: { color: C.teal }
  });
  slide.addText("Technology Stack", {
    x: 0.7, y: 0.2, w: 8, h: 0.75,
    fontSize: 30, fontFace: "Trebuchet MS", color: C.white, bold: true
  });

  // Left column — Runtime & Framework
  const leftItems = [
    { cat: "Runtime", name: "Node.js", ver: "" },
    { cat: "Framework", name: "Express.js", ver: "^4.16.4" },
    { cat: "Dev Server", name: "Nodemon", ver: "^2.0.20" },
    { cat: "Build Tools", name: "Gulp", ver: "^4.0.0" },
  ];

  const rightItems = [
    { cat: "Testing", name: "Mocha + Chai", ver: "^5.2.0 / ^4.2.0" },
    { cat: "HTTP Testing", name: "Supertest", ver: "^3.4.2" },
    { cat: "Coverage", name: "NYC (Istanbul)", ver: "^13.3.0" },
    { cat: "Linting", name: "ESLint", ver: "^8.29.0" },
  ];

  function drawStackCards(items, startX, startY, slide) {
    items.forEach((item, i) => {
      const cy = startY + i * 0.95;

      // Card
      slide.addShape(pres.shapes.RECTANGLE, {
        x: startX, y: cy, w: 4.1, h: 0.78,
        fill: { color: C.white },
        shadow: makeCardShadow()
      });

      // Category badge
      slide.addShape(pres.shapes.RECTANGLE, {
        x: startX + 0.15, y: cy + 0.18, w: 1.1, h: 0.42,
        fill: { color: C.steelBlue, transparency: 10 },
      });
      slide.addText(item.cat, {
        x: startX + 0.15, y: cy + 0.18, w: 1.1, h: 0.42,
        fontSize: 9, fontFace: "Calibri", color: C.white,
        align: "center", valign: "middle", bold: true
      });

      // Name
      slide.addText(item.name, {
        x: startX + 1.4, y: cy + 0.1, w: 2.0, h: 0.35,
        fontSize: 14, fontFace: "Trebuchet MS", color: C.darkNavy, bold: true,
        margin: 0
      });

      // Version
      if (item.ver) {
        slide.addText(item.ver, {
          x: startX + 1.4, y: cy + 0.42, w: 2.5, h: 0.28,
          fontSize: 10, fontFace: "Consolas", color: C.subtleGray,
          margin: 0
        });
      }
    });
  }

  // Section labels
  slide.addText("CORE", {
    x: 0.7, y: 1.3, w: 2, h: 0.35,
    fontSize: 12, fontFace: "Trebuchet MS", color: C.teal, bold: true, charSpacing: 3
  });
  slide.addText("DEVELOPMENT & TESTING", {
    x: 5.2, y: 1.3, w: 4, h: 0.35,
    fontSize: 12, fontFace: "Trebuchet MS", color: C.teal, bold: true, charSpacing: 3
  });

  drawStackCards(leftItems, 0.7, 1.7, slide);
  drawStackCards(rightItems, 5.2, 1.7, slide);
}

// ═══════════════════════════════════════════
// SLIDE 4 — Architecture Diagram
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.darkNavy }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.04, fill: { color: C.teal }
  });
  slide.addText("Application Architecture", {
    x: 0.7, y: 0.2, w: 8, h: 0.75,
    fontSize: 30, fontFace: "Trebuchet MS", color: C.white, bold: true
  });

  // ── Client Layer Box ──
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.45, w: 2.8, h: 3.7,
    fill: { color: C.teal, transparency: 92 },
    line: { color: C.teal, width: 1.5, dashType: "dash" }
  });
  slide.addText("CLIENT LAYER", {
    x: 0.5, y: 1.5, w: 2.8, h: 0.35,
    fontSize: 10, fontFace: "Trebuchet MS", color: C.teal,
    align: "center", bold: true, charSpacing: 2
  });
  slide.addText("Browser", {
    x: 0.5, y: 1.78, w: 2.8, h: 0.22,
    fontSize: 9, fontFace: "Calibri", color: C.medGray,
    align: "center"
  });

  // Client components
  const clientComponents = [
    { name: "index.html", sub: "Calculator UI", y: 2.15 },
    { name: "client.js", sub: "State Machine", y: 3.0 },
    { name: "default.css", sub: "Styles", y: 3.85 },
  ];
  clientComponents.forEach(c => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.85, y: c.y, w: 2.1, h: 0.7,
      fill: { color: C.white },
      shadow: makeCardShadow()
    });
    slide.addText(c.name, {
      x: 0.85, y: c.y + 0.08, w: 2.1, h: 0.32,
      fontSize: 12, fontFace: "Consolas", color: C.darkNavy,
      align: "center", bold: true, margin: 0
    });
    slide.addText(c.sub, {
      x: 0.85, y: c.y + 0.38, w: 2.1, h: 0.25,
      fontSize: 9, fontFace: "Calibri", color: C.medGray,
      align: "center", margin: 0
    });
  });

  // ── Arrow: Client → Server ──
  slide.addShape(pres.shapes.LINE, {
    x: 3.3, y: 3.1, w: 0.8, h: 0,
    line: { color: C.teal, width: 2.5 }
  });
  slide.addText("HTTP GET\n/arithmetic", {
    x: 3.25, y: 2.45, w: 0.9, h: 0.55,
    fontSize: 7, fontFace: "Consolas", color: C.steelBlue,
    align: "center", margin: 0
  });

  // ── Server Layer Box ──
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 4.1, y: 1.45, w: 5.4, h: 3.7,
    fill: { color: C.steelBlue, transparency: 92 },
    line: { color: C.steelBlue, width: 1.5, dashType: "dash" }
  });
  slide.addText("SERVER LAYER", {
    x: 4.1, y: 1.5, w: 5.4, h: 0.35,
    fontSize: 10, fontFace: "Trebuchet MS", color: C.steelBlue,
    align: "center", bold: true, charSpacing: 2
  });
  slide.addText("Node.js + Express", {
    x: 4.1, y: 1.78, w: 5.4, h: 0.22,
    fontSize: 9, fontFace: "Calibri", color: C.medGray,
    align: "center"
  });

  // Server: Express Server box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 4.45, y: 2.15, w: 2.2, h: 0.7,
    fill: { color: C.white },
    shadow: makeCardShadow()
  });
  slide.addText("server.js", {
    x: 4.45, y: 2.2, w: 2.2, h: 0.32,
    fontSize: 12, fontFace: "Consolas", color: C.darkNavy,
    align: "center", bold: true, margin: 0
  });
  slide.addText("Express Server", {
    x: 4.45, y: 2.52, w: 2.2, h: 0.25,
    fontSize: 9, fontFace: "Calibri", color: C.medGray,
    align: "center", margin: 0
  });

  // Static Middleware
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 4.45, y: 3.1, w: 2.2, h: 0.55,
    fill: { color: C.white },
    shadow: makeCardShadow()
  });
  slide.addText("Static File Middleware", {
    x: 4.45, y: 3.15, w: 2.2, h: 0.45,
    fontSize: 10, fontFace: "Calibri", color: C.darkNavy,
    align: "center", margin: 0
  });

  // API Layer sub-box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 7.0, y: 2.15, w: 2.2, h: 2.7,
    fill: { color: C.darkNavy, transparency: 94 },
    line: { color: C.darkNavy, width: 1, dashType: "dash" }
  });
  slide.addText("API LAYER", {
    x: 7.0, y: 2.2, w: 2.2, h: 0.3,
    fontSize: 9, fontFace: "Trebuchet MS", color: C.darkNavy,
    align: "center", bold: true, charSpacing: 2
  });

  // Routes
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 7.2, y: 2.6, w: 1.8, h: 0.65,
    fill: { color: C.white },
    shadow: makeCardShadow()
  });
  slide.addText("routes.js", {
    x: 7.2, y: 2.65, w: 1.8, h: 0.3,
    fontSize: 11, fontFace: "Consolas", color: C.darkNavy,
    align: "center", bold: true, margin: 0
  });
  slide.addText("Route Definitions", {
    x: 7.2, y: 2.93, w: 1.8, h: 0.22,
    fontSize: 8, fontFace: "Calibri", color: C.medGray,
    align: "center", margin: 0
  });

  // Controller
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 7.2, y: 3.5, w: 1.8, h: 0.65,
    fill: { color: C.white },
    shadow: makeCardShadow()
  });
  slide.addText("controller.js", {
    x: 7.2, y: 3.55, w: 1.8, h: 0.3,
    fontSize: 11, fontFace: "Consolas", color: C.darkNavy,
    align: "center", bold: true, margin: 0
  });
  slide.addText("Business Logic", {
    x: 7.2, y: 3.83, w: 1.8, h: 0.22,
    fontSize: 8, fontFace: "Calibri", color: C.medGray,
    align: "center", margin: 0
  });

  // Connector arrows within server
  slide.addShape(pres.shapes.LINE, {
    x: 6.65, y: 2.5, w: 0.35, h: 0,
    line: { color: C.medGray, width: 1.5 }
  });
  slide.addShape(pres.shapes.LINE, {
    x: 5.55, y: 2.85, w: 0, h: 0.25,
    line: { color: C.medGray, width: 1.5 }
  });
  slide.addShape(pres.shapes.LINE, {
    x: 8.1, y: 3.25, w: 0, h: 0.25,
    line: { color: C.medGray, width: 1.5 }
  });

  // JSON Response label
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 7.2, y: 4.35, w: 1.8, h: 0.35,
    fill: { color: C.mint, transparency: 70 }
  });
  slide.addText("{ result: number }", {
    x: 7.2, y: 4.35, w: 1.8, h: 0.35,
    fontSize: 9, fontFace: "Consolas", color: C.steelBlue,
    align: "center", margin: 0
  });
}

// ═══════════════════════════════════════════
// SLIDE 5 — API Endpoint Details
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.darkNavy }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.04, fill: { color: C.teal }
  });
  slide.addText("API Design", {
    x: 0.7, y: 0.2, w: 8, h: 0.75,
    fontSize: 30, fontFace: "Trebuchet MS", color: C.white, bold: true
  });

  // Endpoint badge
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 1.35, w: 1.0, h: 0.38,
    fill: { color: C.mint, transparency: 50 }
  });
  slide.addText("GET", {
    x: 0.7, y: 1.35, w: 1.0, h: 0.38,
    fontSize: 14, fontFace: "Consolas", color: C.darkNavy,
    align: "center", valign: "middle", bold: true, margin: 0
  });
  slide.addText("/arithmetic", {
    x: 1.8, y: 1.35, w: 3, h: 0.38,
    fontSize: 16, fontFace: "Consolas", color: C.darkNavy,
    valign: "middle", margin: 0
  });

  // Query Parameters table
  slide.addText("QUERY PARAMETERS", {
    x: 0.7, y: 1.95, w: 4, h: 0.3,
    fontSize: 10, fontFace: "Trebuchet MS", color: C.teal, bold: true, charSpacing: 2
  });

  const tableHeader = [
    [
      { text: "Parameter", options: { fill: { color: C.darkNavy }, color: C.white, bold: true, fontSize: 11, fontFace: "Calibri" } },
      { text: "Type", options: { fill: { color: C.darkNavy }, color: C.white, bold: true, fontSize: 11, fontFace: "Calibri" } },
      { text: "Required", options: { fill: { color: C.darkNavy }, color: C.white, bold: true, fontSize: 11, fontFace: "Calibri" } },
      { text: "Valid Values", options: { fill: { color: C.darkNavy }, color: C.white, bold: true, fontSize: 11, fontFace: "Calibri" } },
    ],
    [
      { text: "operation", options: { fontFace: "Consolas", fontSize: 10, color: C.charcoal } },
      { text: "string", options: { fontSize: 10, color: C.medGray } },
      { text: "Yes", options: { fontSize: 10, color: C.charcoal, bold: true } },
      { text: "add, subtract, multiply, divide", options: { fontFace: "Consolas", fontSize: 9, color: C.steelBlue } },
    ],
    [
      { text: "operand1", options: { fontFace: "Consolas", fontSize: 10, color: C.charcoal } },
      { text: "number", options: { fontSize: 10, color: C.medGray } },
      { text: "Yes", options: { fontSize: 10, color: C.charcoal, bold: true } },
      { text: "Any valid number (decimals, scientific)", options: { fontSize: 9, color: C.steelBlue } },
    ],
    [
      { text: "operand2", options: { fontFace: "Consolas", fontSize: 10, color: C.charcoal } },
      { text: "number", options: { fontSize: 10, color: C.medGray } },
      { text: "Yes", options: { fontSize: 10, color: C.charcoal, bold: true } },
      { text: "Any valid number (decimals, scientific)", options: { fontSize: 9, color: C.steelBlue } },
    ],
  ];

  slide.addTable(tableHeader, {
    x: 0.7, y: 2.3, w: 8.6,
    colW: [1.6, 1.1, 1.1, 4.8],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.4, 0.35, 0.35, 0.35],
    autoPage: false
  });

  // Response examples — side by side
  // Success
  slide.addText("SUCCESS RESPONSE", {
    x: 0.7, y: 3.85, w: 4, h: 0.3,
    fontSize: 10, fontFace: "Trebuchet MS", color: C.mint, bold: true, charSpacing: 2
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 4.2, w: 4.0, h: 1.1,
    fill: { color: C.darkNavy }
  });
  slide.addText([
    { text: "200 OK", options: { fontSize: 9, color: C.mint, bold: true, breakLine: true } },
    { text: '{ "result": 8 }', options: { fontSize: 14, fontFace: "Consolas", color: C.white } },
  ], {
    x: 0.9, y: 4.3, w: 3.6, h: 0.9,
    valign: "middle"
  });

  // Error
  slide.addText("ERROR RESPONSE", {
    x: 5.3, y: 3.85, w: 4, h: 0.3,
    fontSize: 10, fontFace: "Trebuchet MS", color: C.errorRed, bold: true, charSpacing: 2
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 5.3, y: 4.2, w: 4.0, h: 1.1,
    fill: { color: C.darkNavy }
  });
  slide.addText([
    { text: "400 Bad Request", options: { fontSize: 9, color: C.errorRed, bold: true, breakLine: true } },
    { text: '{ "error": "Invalid operation" }', options: { fontSize: 12, fontFace: "Consolas", color: C.white } },
  ], {
    x: 5.5, y: 4.3, w: 3.6, h: 0.9,
    valign: "middle"
  });
}

// ═══════════════════════════════════════════
// SLIDE 6 — Client State Machine
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.darkNavy }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.04, fill: { color: C.teal }
  });
  slide.addText("Client State Machine", {
    x: 0.7, y: 0.2, w: 8, h: 0.75,
    fontSize: 30, fontFace: "Trebuchet MS", color: C.white, bold: true
  });

  slide.addText("The calculator UI uses a finite state machine to manage input flow and transitions", {
    x: 0.7, y: 1.3, w: 8.6, h: 0.35,
    fontSize: 12, fontFace: "Calibri", color: C.medGray, italic: true
  });

  // State nodes — horizontal flow
  const states = [
    { name: "START", val: "0", x: 0.5, color: C.steelBlue },
    { name: "OPERAND1", val: "1", x: 2.4, color: C.teal },
    { name: "OPERATOR", val: "2", x: 4.3, color: C.mint },
    { name: "OPERAND2", val: "3", x: 6.2, color: C.teal },
    { name: "COMPLETE", val: "4", x: 8.1, color: C.steelBlue },
  ];

  const stateY = 2.0, stateW = 1.5, stateH = 0.9;

  states.forEach(s => {
    // State circle/rounded rect
    slide.addShape(pres.shapes.OVAL, {
      x: s.x, y: stateY, w: stateW, h: stateH,
      fill: { color: s.color },
      shadow: makeCardShadow()
    });
    slide.addText(s.name, {
      x: s.x, y: stateY + 0.12, w: stateW, h: 0.38,
      fontSize: 11, fontFace: "Trebuchet MS", color: C.white,
      align: "center", bold: true, margin: 0
    });
    slide.addText(s.val, {
      x: s.x, y: stateY + 0.48, w: stateW, h: 0.28,
      fontSize: 9, fontFace: "Consolas", color: C.white,
      align: "center", margin: 0
    });
  });

  // Arrows between states
  const arrowY = stateY + stateH / 2;
  for (let i = 0; i < states.length - 1; i++) {
    const x1 = states[i].x + stateW;
    const x2 = states[i + 1].x;
    slide.addShape(pres.shapes.LINE, {
      x: x1, y: arrowY, w: x2 - x1, h: 0,
      line: { color: C.charcoal, width: 2 }
    });
  }

  // Transition labels
  const transitions = [
    { text: "Press\n1-9", x: 1.7, y: 3.1 },
    { text: "Press\n+,-,*,/", x: 3.6, y: 3.1 },
    { text: "Press\ndigit", x: 5.5, y: 3.1 },
    { text: "Press\n=", x: 7.4, y: 3.1 },
  ];
  transitions.forEach(t => {
    slide.addText(t.text, {
      x: t.x, y: t.y, w: 1.0, h: 0.5,
      fontSize: 8, fontFace: "Calibri", color: C.medGray,
      align: "center", margin: 0
    });
  });

  // State description table below
  const stateTableData = [
    [
      { text: "State", options: { fill: { color: C.darkNavy }, color: C.white, bold: true, fontSize: 10 } },
      { text: "Value", options: { fill: { color: C.darkNavy }, color: C.white, bold: true, fontSize: 10 } },
      { text: "Description", options: { fill: { color: C.darkNavy }, color: C.white, bold: true, fontSize: 10 } },
    ],
    [
      { text: "start", options: { fontFace: "Consolas", fontSize: 10, color: C.charcoal } },
      { text: "0", options: { fontSize: 10, color: C.medGray, align: "center" } },
      { text: "Initial state, display shows 0, waiting for first input", options: { fontSize: 10, color: C.charcoal } },
    ],
    [
      { text: "operand1", options: { fontFace: "Consolas", fontSize: 10, color: C.charcoal } },
      { text: "1", options: { fontSize: 10, color: C.medGray, align: "center" } },
      { text: "User is entering the first operand", options: { fontSize: 10, color: C.charcoal } },
    ],
    [
      { text: "operator", options: { fontFace: "Consolas", fontSize: 10, color: C.charcoal } },
      { text: "2", options: { fontSize: 10, color: C.medGray, align: "center" } },
      { text: "Operator selected, waiting for second operand", options: { fontSize: 10, color: C.charcoal } },
    ],
    [
      { text: "operand2", options: { fontFace: "Consolas", fontSize: 10, color: C.charcoal } },
      { text: "3", options: { fontSize: 10, color: C.medGray, align: "center" } },
      { text: "User is entering the second operand", options: { fontSize: 10, color: C.charcoal } },
    ],
    [
      { text: "complete", options: { fontFace: "Consolas", fontSize: 10, color: C.charcoal } },
      { text: "4", options: { fontSize: 10, color: C.medGray, align: "center" } },
      { text: "Calculation complete, result displayed", options: { fontSize: 10, color: C.charcoal } },
    ],
  ];

  slide.addTable(stateTableData, {
    x: 0.7, y: 3.75, w: 8.6,
    colW: [1.3, 0.7, 6.6],
    border: { pt: 0.5, color: C.lightGray },
    rowH: [0.3, 0.28, 0.28, 0.28, 0.28, 0.28],
    autoPage: false
  });
}

// ═══════════════════════════════════════════
// SLIDE 7 — Request Flow (Sequence Diagram)
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.darkNavy }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.04, fill: { color: C.teal }
  });
  slide.addText("Request Flow", {
    x: 0.7, y: 0.2, w: 8, h: 0.75,
    fontSize: 30, fontFace: "Trebuchet MS", color: C.white, bold: true
  });

  slide.addText("End-to-end flow of a calculation: 5 + 3 = 8", {
    x: 0.7, y: 1.3, w: 8.6, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: C.medGray, italic: true
  });

  // Participant headers
  const participants = [
    { name: "User", x: 0.6 },
    { name: "Calculator UI", x: 2.3 },
    { name: "client.js", x: 4.0 },
    { name: "Express", x: 5.7 },
    { name: "routes.js", x: 7.0 },
    { name: "controller.js", x: 8.5 },
  ];
  const pW = 1.3, pH = 0.4, pY = 1.75;

  participants.forEach(p => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: p.x, y: pY, w: pW, h: pH,
      fill: { color: C.steelBlue },
      shadow: makeCardShadow()
    });
    slide.addText(p.name, {
      x: p.x, y: pY, w: pW, h: pH,
      fontSize: 9, fontFace: "Trebuchet MS", color: C.white,
      align: "center", valign: "middle", bold: true, margin: 0
    });

    // Lifeline
    slide.addShape(pres.shapes.LINE, {
      x: p.x + pW / 2, y: pY + pH, w: 0, h: 3.12,
      line: { color: C.subtleGray, width: 1, dashType: "dash" }
    });
  });

  // Flow steps — horizontal arrows with labels
  const steps = [
    { from: 0, to: 1, y: 2.4, label: "Enter '5'" },
    { from: 1, to: 2, y: 2.7, label: "numberPressed('5')" },
    { from: 0, to: 1, y: 3.0, label: "Press '+'" },
    { from: 1, to: 2, y: 3.3, label: "operationPressed('+')" },
    { from: 0, to: 1, y: 3.6, label: "Press '='" },
    { from: 2, to: 3, y: 3.9, label: "GET /arithmetic?..." },
    { from: 3, to: 4, y: 4.15, label: "route()" },
    { from: 4, to: 5, y: 4.4, label: "calculate()" },
    { from: 5, to: 2, y: 4.7, label: "{ result: 8 }", isReturn: true },
  ];

  steps.forEach(s => {
    const fromX = participants[s.from].x + pW / 2;
    const toX = participants[s.to].x + pW / 2;
    const lineW = toX - fromX;

    slide.addShape(pres.shapes.LINE, {
      x: Math.min(fromX, toX), y: s.y, w: Math.abs(lineW), h: 0,
      line: { color: s.isReturn ? C.mint : C.steelBlue, width: 1.5, dashType: s.isReturn ? "dash" : "solid" }
    });

    // Label
    const labelX = Math.min(fromX, toX) + Math.abs(lineW) * 0.1;
    slide.addText(s.label, {
      x: labelX, y: s.y - 0.2, w: Math.abs(lineW) * 0.9, h: 0.18,
      fontSize: 7, fontFace: "Consolas", color: s.isReturn ? C.steelBlue : C.charcoal,
      margin: 0
    });
  });
}

// ═══════════════════════════════════════════
// SLIDE 8 — Testing & Quality
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.darkNavy }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.04, fill: { color: C.teal }
  });
  slide.addText("Testing & Quality", {
    x: 0.7, y: 0.2, w: 8, h: 0.75,
    fontSize: 30, fontFace: "Trebuchet MS", color: C.white, bold: true
  });

  // Testing framework info — left side
  slide.addText("TESTING FRAMEWORK", {
    x: 0.7, y: 1.35, w: 4, h: 0.3,
    fontSize: 10, fontFace: "Trebuchet MS", color: C.teal, bold: true, charSpacing: 2
  });

  const testTools = [
    { name: "Mocha", desc: "Test runner with BDD-style describe/it blocks" },
    { name: "Chai", desc: "Assertion library with expect/should syntax" },
    { name: "Supertest", desc: "HTTP endpoint testing via request simulation" },
    { name: "NYC (Istanbul)", desc: "Code coverage with Cobertura & HTML reports" },
  ];

  testTools.forEach((t, i) => {
    const cy = 1.75 + i * 0.7;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y: cy, w: 4.2, h: 0.58,
      fill: { color: C.white },
      shadow: makeCardShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.7, y: cy, w: 0.06, h: 0.58,
      fill: { color: C.teal }
    });
    slide.addText(t.name, {
      x: 0.95, y: cy + 0.05, w: 3.8, h: 0.25,
      fontSize: 13, fontFace: "Trebuchet MS", color: C.darkNavy, bold: true, margin: 0
    });
    slide.addText(t.desc, {
      x: 0.95, y: cy + 0.3, w: 3.8, h: 0.23,
      fontSize: 10, fontFace: "Calibri", color: C.medGray, margin: 0
    });
  });

  // Test categories — right side
  slide.addText("TEST COVERAGE", {
    x: 5.3, y: 1.35, w: 4, h: 0.3,
    fontSize: 10, fontFace: "Trebuchet MS", color: C.teal, bold: true, charSpacing: 2
  });

  const categories = [
    { name: "Validation", count: "5 tests", desc: "Missing/invalid operations & operands" },
    { name: "Addition", count: "6 tests", desc: "Integers, negatives, floats, scientific notation" },
    { name: "Multiplication", count: "tests", desc: "Positive, negative, float multiplication" },
    { name: "Division", count: "tests", desc: "Standard and edge case division" },
  ];

  categories.forEach((c, i) => {
    const cy = 1.75 + i * 0.7;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.3, y: cy, w: 4.0, h: 0.58,
      fill: { color: C.white },
      shadow: makeCardShadow()
    });

    // Count badge
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.45, y: cy + 0.12, w: 0.85, h: 0.34,
      fill: { color: C.steelBlue }
    });
    slide.addText(c.count, {
      x: 5.45, y: cy + 0.12, w: 0.85, h: 0.34,
      fontSize: 8, fontFace: "Calibri", color: C.white,
      align: "center", valign: "middle", bold: true, margin: 0
    });

    slide.addText(c.name, {
      x: 6.45, y: cy + 0.05, w: 2.7, h: 0.25,
      fontSize: 13, fontFace: "Trebuchet MS", color: C.darkNavy, bold: true, margin: 0
    });
    slide.addText(c.desc, {
      x: 6.45, y: cy + 0.3, w: 2.7, h: 0.23,
      fontSize: 9, fontFace: "Calibri", color: C.medGray, margin: 0
    });
  });

  // npm test command
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 4.7, w: 8.6, h: 0.55,
    fill: { color: C.darkNavy }
  });
  slide.addText([
    { text: "$  ", options: { color: C.mint, fontFace: "Consolas", fontSize: 13 } },
    { text: "npm test", options: { color: C.white, fontFace: "Consolas", fontSize: 13, bold: true } },
    { text: "   # Runs Mocha + NYC coverage + generates HTML & Cobertura reports", options: { color: C.subtleGray, fontFace: "Consolas", fontSize: 10 } },
  ], {
    x: 0.9, y: 4.7, w: 8.2, h: 0.55,
    valign: "middle"
  });
}

// ═══════════════════════════════════════════
// SLIDE 9 — Project Structure
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.darkNavy }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.04, fill: { color: C.teal }
  });
  slide.addText("Project Structure", {
    x: 0.7, y: 0.2, w: 8, h: 0.75,
    fontSize: 30, fontFace: "Trebuchet MS", color: C.white, bold: true
  });

  // File tree — left side (dark code block style)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.35, w: 4.3, h: 3.9,
    fill: { color: C.darkNavy }
  });

  const treeLines = [
    { text: "copilot-node-calculator/", indent: 0, isDir: true },
    { text: "  server.js", indent: 1 },
    { text: "  package.json", indent: 1 },
    { text: "  gulpfile.js", indent: 1 },
    { text: "  documentation.md", indent: 1 },
    { text: "", indent: 0 },
    { text: "  api/", indent: 1, isDir: true },
    { text: "    controller.js", indent: 2 },
    { text: "    routes.js", indent: 2 },
    { text: "", indent: 0 },
    { text: "  public/", indent: 1, isDir: true },
    { text: "    index.html", indent: 2 },
    { text: "    client.js", indent: 2 },
    { text: "    default.css", indent: 2 },
    { text: "", indent: 0 },
    { text: "  test/", indent: 1, isDir: true },
    { text: "    arithmetic.test.js", indent: 2 },
    { text: "    helpers.js", indent: 2 },
  ];

  const treeTextParts = treeLines.map((line, i) => ({
    text: line.text,
    options: {
      fontFace: "Consolas",
      fontSize: 10,
      color: line.isDir ? C.teal : C.white,
      bold: !!line.isDir,
      breakLine: i < treeLines.length - 1,
    }
  }));

  slide.addText(treeTextParts, {
    x: 0.7, y: 1.5, w: 3.9, h: 3.6,
    valign: "top",
    lineSpacingMultiple: 1.15
  });

  // Key files — right side cards
  slide.addText("KEY FILES", {
    x: 5.3, y: 1.35, w: 4, h: 0.3,
    fontSize: 10, fontFace: "Trebuchet MS", color: C.teal, bold: true, charSpacing: 2
  });

  const keyFiles = [
    { file: "server.js", desc: "Express server entry point, serves\nstatic files and mounts API routes" },
    { file: "api/controller.js", desc: "Core calculation logic and\ninput validation" },
    { file: "api/routes.js", desc: "Defines /arithmetic endpoint\nrouting" },
    { file: "public/client.js", desc: "Client-side state machine and\nAPI communication" },
    { file: "public/index.html", desc: "Calculator user interface\nwith button grid" },
  ];

  keyFiles.forEach((f, i) => {
    const cy = 1.75 + i * 0.72;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.3, y: cy, w: 4.2, h: 0.62,
      fill: { color: C.white },
      shadow: makeCardShadow()
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.3, y: cy, w: 0.06, h: 0.62,
      fill: { color: C.steelBlue }
    });
    slide.addText(f.file, {
      x: 5.55, y: cy + 0.05, w: 3.8, h: 0.22,
      fontSize: 11, fontFace: "Consolas", color: C.darkNavy, bold: true, margin: 0
    });
    slide.addText(f.desc, {
      x: 5.55, y: cy + 0.28, w: 3.8, h: 0.3,
      fontSize: 9, fontFace: "Calibri", color: C.medGray, margin: 0
    });
  });
}

// ═══════════════════════════════════════════
// SLIDE 10 — Getting Started
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.offWhite };

  // Title bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 1.1, fill: { color: C.darkNavy }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 1.1, w: 10, h: 0.04, fill: { color: C.teal }
  });
  slide.addText("Getting Started", {
    x: 0.7, y: 0.2, w: 8, h: 0.75,
    fontSize: 30, fontFace: "Trebuchet MS", color: C.white, bold: true
  });

  // Step cards — numbered
  const steps = [
    { num: "1", title: "Clone & Install", cmd: "git clone <repo-url>\ncd node-calculator\nnpm install" },
    { num: "2", title: "Run Development Server", cmd: "npm start\n# Starts Nodemon on port 3000" },
    { num: "3", title: "Run Tests", cmd: "npm test\n# Mocha + NYC coverage reports" },
    { num: "4", title: "Lint Code", cmd: "npm run lint\n# ESLint check" },
  ];

  steps.forEach((s, i) => {
    const cy = 1.4 + i * 1.0;

    // Number circle
    slide.addShape(pres.shapes.OVAL, {
      x: 0.7, y: cy + 0.05, w: 0.55, h: 0.55,
      fill: { color: C.teal }
    });
    slide.addText(s.num, {
      x: 0.7, y: cy + 0.05, w: 0.55, h: 0.55,
      fontSize: 18, fontFace: "Trebuchet MS", color: C.white,
      align: "center", valign: "middle", bold: true, margin: 0
    });

    // Step title
    slide.addText(s.title, {
      x: 1.4, y: cy, w: 3, h: 0.35,
      fontSize: 16, fontFace: "Trebuchet MS", color: C.darkNavy, bold: true, margin: 0
    });

    // Command box
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 1.4, y: cy + 0.38, w: 7.9, h: 0.52,
      fill: { color: C.darkNavy }
    });
    slide.addText(s.cmd, {
      x: 1.6, y: cy + 0.38, w: 7.5, h: 0.52,
      fontSize: 10, fontFace: "Consolas", color: C.mint,
      valign: "middle", margin: 0
    });
  });

  // Environment variable note
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 5.0, w: 8.6, h: 0.35,
    fill: { color: C.teal, transparency: 85 }
  });
  slide.addText("Environment: Set PORT to customize server port (default: 3000)", {
    x: 0.9, y: 5.0, w: 8.2, h: 0.35,
    fontSize: 11, fontFace: "Calibri", color: C.steelBlue, margin: 0, valign: "middle"
  });
}

// ═══════════════════════════════════════════
// SLIDE 11 — Thank You / Closing
// ═══════════════════════════════════════════
{
  const slide = pres.addSlide();
  slide.background = { color: C.darkNavy };

  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }
  });

  // Decorative shapes (mirroring title slide)
  slide.addShape(pres.shapes.RECTANGLE, {
    x: -0.5, y: 2.5, w: 3.0, h: 3.0,
    fill: { color: C.steelBlue, transparency: 20 },
    rotate: 45
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 8.0, y: 0.8, w: 2.5, h: 2.5,
    fill: { color: C.teal, transparency: 30 },
    rotate: 45
  });

  // Main text
  slide.addText("Thank You", {
    x: 1.5, y: 1.5, w: 7, h: 1.2,
    fontSize: 48, fontFace: "Trebuchet MS", color: C.white,
    align: "center", bold: true, charSpacing: 3
  });

  // Divider
  slide.addShape(pres.shapes.LINE, {
    x: 3.5, y: 2.8, w: 3, h: 0,
    line: { color: C.mint, width: 2.5 }
  });

  // Subtitle
  slide.addText("Node.js Calculator Application", {
    x: 1.5, y: 3.1, w: 7, h: 0.5,
    fontSize: 18, fontFace: "Calibri", color: C.teal,
    align: "center", italic: true
  });

  // Tech summary
  slide.addText("Express.js  |  RESTful API  |  State Machine  |  Full Test Suite", {
    x: 1.5, y: 3.8, w: 7, h: 0.4,
    fontSize: 12, fontFace: "Calibri", color: C.subtleGray,
    align: "center"
  });

  // Bottom bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.35, w: 10, h: 0.28, fill: { color: C.navy }
  });
}

// ── Save ──
const outputPath = "NodeJS_Calculator_Presentation.pptx";
pres.writeFile({ fileName: outputPath })
  .then(() => console.log(`Presentation saved: ${outputPath}`))
  .catch(err => console.error("Error:", err));
