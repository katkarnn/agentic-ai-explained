// ============================================================================
//  How Agentic AI Actually Works  —  LinkedIn carousel (square 1:1)
//  Theme: HYBRID — light, high-contrast slides with embedded dark terminal
//         panels for the developer feel. Built programmatically with PptxGenJS.
//  Build + QA:  powershell -ExecutionPolicy Bypass -File render.ps1
// ============================================================================
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

// ---------- palette (hex, no '#') — tuned for sunlight + dark-room contrast --
const DARK   = "082A30";  // dark teal ink — divider / hero / keystone bg
const DARK2  = "0C3942";  // decorative circles / cards on dark
const BG     = "FFFFFF";  // light slide bg (most readable in sunlight)
const BGSOFT = "F3F9FA";  // alt light bg
const CARD   = "F1F8F9";  // subtle card tint on light
const INK    = "0B2329";  // body text on light (near-black, high contrast)
const HEAD   = "082A30";  // headers on light
const TEAL   = "027A8A";  // primary accent (fills, badges)
const TEALTX = "045E6B";  // darker teal for accent TEXT on light (contrast)
const MINT   = "06D6A0";  // vibrant accent — fills/badges or text on dark only
const MINTTX = "0A7A5C";  // dark mint for text on light
const AMBER  = "E8893A";  // warm accent for warnings / "gotcha" fills
const AMBERTX= "B45F16";  // amber text on light
const MUTED  = "4B666E";  // captions / secondary text on light
const LINE   = "D7E6E9";  // hairline borders

// terminal panel palette (dark)
const TBG    = "0A1E22";  // terminal background
const TBORDER= "16434C";  // terminal border / separators
const TGREEN = "5EEAA8";  // prompt / success
const TCYAN  = "6FD3E6";  // keywords / tool names
const TYEL   = "F2C14E";  // strings / values
const TCMT   = "7FA0A8";  // comments / filenames
const TWHITE = "E8F4F6";  // default terminal text
const TRED   = "E76F51";  // window dot / errors

const HF   = "Cambria";   // headers (serif, QA-safe)
const BF   = "Calibri";   // body (sans, QA-safe)
const MONO = "Consolas";  // terminal / code (monospace)

// ---------- icon rasterizer (react-icons -> base64 PNG via sharp) -----------
async function icon(Comp, color = "#FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}
const mkShadow = () => ({ type: "outer", color: "0A2A30", blur: 7, offset: 3, angle: 90, opacity: 0.16 });

(async () => {
  // ---- icons ----
  const need = {
    robot: FA.FaRobot, infinity: FA.FaInfinity, brain: FA.FaBrain, bolt: FA.FaBolt,
    eye: FA.FaEye, sync: FA.FaSyncAlt, redo: FA.FaRedo, inbox: FA.FaInbox,
    check2: FA.FaCheckDouble, shield: FA.FaShieldAlt, tools: FA.FaToolbox,
    mem: FA.FaMemory, db: FA.FaDatabase, mark: FA.FaBookmark, layers: FA.FaLayerGroup,
    sitemap: FA.FaSitemap, users: FA.FaUsersCog, route: FA.FaRoute, compass: FA.FaCompass,
    link: FA.FaLink, branch: FA.FaCodeBranch, stream: FA.FaStream, diagram: FA.FaProjectDiagram,
    bulb: FA.FaLightbulb, cap: FA.FaGraduationCap, term: FA.FaTerminal, plug: FA.FaPlug,
    search: FA.FaSearch, comments: FA.FaComments, hand: FA.FaHandPaper, warn: FA.FaExclamationTriangle,
    coins: FA.FaCoins, scope: FA.FaMicroscope, tasks: FA.FaTasks, signs: FA.FaMapSigns,
    book: FA.FaBook, key: FA.FaKey, filter: FA.FaFilter, gears: FA.FaCogs, user: FA.FaUserCheck,
    fork: FA.FaUtensils, rocket: FA.FaRocket, magic: FA.FaMagic, lock: FA.FaLock, flag: FA.FaFlagCheckered,
  };
  const I = {};            // white icons (for dark fills)
  const Id = {};           // dark icons (for light fills, if ever needed)
  for (const k in need) { I[k] = await icon(need[k], "#FFFFFF", 256); }
  const arrowTeal = await icon(FA.FaArrowRight, "#" + TEAL, 256);

  // ---- presentation ----
  const pres = new pptxgen();
  pres.defineLayout({ name: "SQ", width: 10, height: 10 });
  pres.layout = "SQ";
  pres.author = "Nikhil K";
  pres.title = "How Agentic AI Actually Works";
  const SH = pres.shapes;

  // track slides so page numbers get the right denominator automatically
  const slides = [];
  function newSlide(bg, onDark) {
    const s = pres.addSlide();
    s.background = { color: bg };
    slides.push({ s, onDark: !!onDark });
    return s;
  }

  // ---------------- helpers ----------------
  function badge(slide, x, y, d, fill, ic) {
    slide.addShape(SH.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, shadow: mkShadow() });
    const ins = d * 0.27;
    slide.addImage({ data: ic, x: x + ins, y: y + ins, w: d - 2 * ins, h: d - 2 * ins });
  }
  function card(slide, x, y, w, h, fill) {
    slide.addShape(SH.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12,
      fill: { color: fill || BG }, line: { color: LINE, width: 1 }, shadow: mkShadow() });
  }
  function kicker(slide, txt, color, x = 0.6, y = 0.52) {
    slide.addText(txt, { x, y, w: 8.8, h: 0.4, fontFace: BF, fontSize: 12.5, bold: true,
      color, charSpacing: 3, align: "left" });
  }
  function title(slide, txt, color, y = 0.95, size = 31, x = 0.6, w = 8.85, h = 1.15) {
    slide.addText(txt, { x, y, w, h, fontFace: HF, fontSize: size, bold: true, color, align: "left", lineSpacingMultiple: 1.0 });
  }
  function lead(slide, txt, y = 2.0, color = INK, x = 0.6, w = 8.85, size = 16.5) {
    slide.addText(txt, { x, y, w, h: 1.0, fontFace: BF, fontSize: size, color, lineSpacingMultiple: 1.12 });
  }
  function note(slide, txt, y, color = TEALTX, x = 0.6, w = 8.85, size = 15) {
    slide.addText(txt, { x, y, w, h: 0.8, fontFace: BF, fontSize: size, italic: true, color, lineSpacingMultiple: 1.08 });
  }
  function chip(slide, x, y, w, h, txt, fill, txtColor, fontSize = 14) {
    slide.addShape(SH.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { type: "none" }, shadow: mkShadow() });
    slide.addText(txt, { x, y, w, h, align: "center", valign: "middle", fontFace: BF, fontSize, bold: true, color: txtColor });
  }
  function arrow(slide, x1, y1, x2, y2, color, width = 2.75) {
    slide.addShape(SH.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1,
      line: { color, width, endArrowType: "triangle" } });
  }
  // dark terminal/code panel. lines = array of rows; each row = array of {t,c,b} runs.
  function term(slide, x, y, w, h, fname, lines, opt) {
    opt = opt || {};
    const fs = opt.fs || 13.5, bar = 0.46;
    slide.addShape(SH.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.09,
      fill: { color: TBG }, line: { color: TBORDER, width: 1 }, shadow: mkShadow() });
    [TRED, TYEL, TGREEN].forEach((c, i) =>
      slide.addShape(SH.OVAL, { x: x + 0.24 + i * 0.26, y: y + bar / 2 - 0.075, w: 0.15, h: 0.15, fill: { color: c }, line: { type: "none" } }));
    if (fname) slide.addText(fname, { x: x + 1.25, y, w: w - 1.5, h: bar, fontFace: MONO, fontSize: 10.5, color: TCMT, align: "left", valign: "middle" });
    slide.addShape(SH.LINE, { x, y: y + bar, w, h: 0, line: { color: TBORDER, width: 1 } });
    const runs = [];
    lines.forEach(line => {
      if (!line || line.length === 0) { runs.push({ text: " ", options: { breakLine: true } }); return; }
      line.forEach((r, j) => runs.push({ text: r.t, options: { color: r.c || TWHITE, bold: !!r.b, breakLine: j === line.length - 1 } }));
    });
    slide.addText(runs, { x: x + 0.32, y: y + bar + 0.13, w: w - 0.6, h: h - bar - 0.26,
      fontFace: MONO, fontSize: fs, align: "left", valign: "top", lineSpacingMultiple: opt.ls || 1.3, color: TWHITE });
  }
  function pageno(slide, n, total, onDark) {
    slide.addText(`${n} / ${total}`, { x: 8.55, y: 9.5, w: 1.05, h: 0.32, fontFace: MONO, fontSize: 10.5,
      color: onDark ? "5E8C94" : MUTED, align: "right" });
  }
  // dark section divider
  function divider(part, ttl, sub, ic) {
    const s = newSlide(DARK, true);
    s.addShape(SH.OVAL, { x: 7.0, y: -1.7, w: 5.0, h: 5.0, fill: { color: DARK2 }, line: { type: "none" } });
    s.addShape(SH.OVAL, { x: -1.7, y: 7.2, w: 4.4, h: 4.4, fill: { color: DARK2 }, line: { type: "none" } });
    badge(s, 0.6, 2.55, 1.4, MINT, ic);
    s.addText(part, { x: 0.62, y: 4.15, w: 8.8, h: 0.4, fontFace: MONO, fontSize: 14, bold: true, color: MINT, charSpacing: 2 });
    s.addText(ttl, { x: 0.58, y: 4.55, w: 8.9, h: 1.6, fontFace: HF, fontSize: 44, bold: true, color: BG, lineSpacingMultiple: 1.0 });
    s.addText(sub, { x: 0.6, y: 6.5, w: 8.6, h: 1.0, fontFace: BF, fontSize: 18, color: "CDEDF0", lineSpacingMultiple: 1.12 });
    return s;
  }

  // ============================================================ SLIDE 1 — TITLE
  {
    const s = newSlide(DARK, true);
    s.addShape(SH.OVAL, { x: 6.9, y: -1.6, w: 5.2, h: 5.2, fill: { color: DARK2 }, line: { type: "none" } });
    s.addShape(SH.OVAL, { x: -1.8, y: 7.1, w: 4.6, h: 4.6, fill: { color: DARK2 }, line: { type: "none" } });
    badge(s, 0.6, 0.85, 1.45, MINT, I.robot);
    s.addText("AGENTIC AI  ·  EXPLAINED LIKE YOU'RE NEW", { x: 0.62, y: 2.6, w: 8.8, h: 0.4,
      fontFace: MONO, fontSize: 13.5, bold: true, color: MINT, charSpacing: 2 });
    s.addText("How Agentic AI\nActually Works", { x: 0.55, y: 3.05, w: 9, h: 2.0,
      fontFace: HF, fontSize: 50, bold: true, color: BG, lineSpacingMultiple: 1.0 });
    term(s, 0.6, 5.25, 6.7, 1.95, "intro.sh", [
      [{ t: "$ ", c: TGREEN, b: true }, { t: "ask-ai ", c: TWHITE }, { t: '"teach me agentic AI, from zero"', c: TYEL }],
      [{ t: "> no buzzwords. real examples. one step at a time.", c: TCMT }],
      [{ t: "> if you can read code, you can get this.", c: TCYAN }],
    ], { fs: 12.5 });
    s.addText([{ text: "Swipe to learn  ", options: { color: BG, bold: true } },
               { text: "→", options: { color: MINT, bold: true } }],
      { x: 0.62, y: 8.75, w: 5, h: 0.5, fontFace: BF, fontSize: 17 });
  }

  // ============================================ SLIDE 2 — PROMISE / WHAT YOU GET
  {
    const s = newSlide(BG);
    kicker(s, "WHAT YOU'LL WALK AWAY WITH", TEALTX);
    title(s, "By the last slide, you can explain all of this", HEAD, 0.95, 27);
    lead(s, "No math. No PhD. Just clear ideas, pictures, and tiny code examples — built up one piece at a time.", 1.9, MUTED);
    const wins = [
      { ic: I.brain,  t: "What an LLM really is", d: "and the one thing it actually does" },
      { ic: I.sync,   t: "Why agents run in a loop", d: "instead of answering in one shot" },
      { ic: I.tools,  t: "Tools, memory & guardrails", d: "the parts of a real system" },
      { ic: I.route,  t: "Workflow vs Agent", d: "the decision that ends the confusion" },
      { ic: I.diagram,t: "The 5 patterns", d: "you'll spot in every product" },
      { ic: I.rocket, t: "When to build an agent", d: "and when NOT to" },
    ];
    const cw = 4.28, cx = [0.6, 5.12], ch = 1.45, gy = 0.18; let row = 0;
    wins.forEach((w, i) => {
      const x = cx[i % 2]; const y = 2.75 + row * (ch + gy);
      card(s, x, y, cw, ch, CARD);
      badge(s, x + 0.28, y + ch / 2 - 0.4, 0.8, i % 2 ? TEAL : DARK, w.ic);
      s.addText(w.t, { x: x + 1.25, y: y + 0.22, w: cw - 1.45, h: 0.5, fontFace: BF, fontSize: 15.5, bold: true, color: INK });
      s.addText(w.d, { x: x + 1.25, y: y + 0.72, w: cw - 1.45, h: 0.55, fontFace: BF, fontSize: 12.5, color: MUTED, lineSpacingMultiple: 1.0 });
      if (i % 2) row++;
    });
    note(s, "Save this post — it doubles as a cheat-sheet you can come back to.", 8.65, TEALTX);
  }

  // ================================================= SLIDE 3 — BIG PICTURE (zoom-out)
  {
    const s = newSlide(BG);
    kicker(s, "THE 10,000-FOOT VIEW", TEALTX);
    title(s, "The whole idea in one picture", HEAD);
    lead(s, "Everything in this deck is just zooming into the middle box.", 1.9, MUTED);

    // input
    chip(s, 0.6, 4.35, 2.0, 1.3, "", TEAL, BG); // base
    s.addText([{ text: "GOAL\n", options: { bold: true, fontSize: 16, color: BG } },
               { text: "a task or trigger", options: { fontSize: 11.5, color: "EAF6F7" } }],
      { x: 0.6, y: 4.35, w: 2.0, h: 1.3, align: "center", valign: "middle", fontFace: BF, lineSpacingMultiple: 1.0 });
    // agent card (center)
    card(s, 3.15, 3.3, 3.7, 3.4, DARK);
    s.addText("THE AGENT", { x: 3.15, y: 3.55, w: 3.7, h: 0.4, align: "center", fontFace: BF, fontSize: 15, bold: true, color: MINT });
    const loop = [["REASON", TCYAN], ["ACT", MINT], ["OBSERVE", TCYAN]];
    loop.forEach((L, i) => {
      const ly = 4.15 + i * 0.72;
      s.addShape(SH.ROUNDED_RECTANGLE, { x: 3.55, y: ly, w: 2.9, h: 0.56, rectRadius: 0.08, fill: { color: DARK2 }, line: { color: "11515C", width: 1 } });
      s.addText(L[0], { x: 3.55, y: ly, w: 2.9, h: 0.56, align: "center", valign: "middle", fontFace: MONO, fontSize: 14, bold: true, color: L[1] });
    });
    s.addText("↻ loops until done", { x: 3.15, y: 6.3, w: 3.7, h: 0.35, align: "center", fontFace: MONO, fontSize: 12, italic: true, color: "9FD8C6" });
    // output
    chip(s, 7.4, 4.35, 2.0, 1.3, "", MINT, DARK);
    s.addText([{ text: "ANSWER\n", options: { bold: true, fontSize: 16, color: DARK } },
               { text: "+ real actions", options: { fontSize: 11.5, color: "06403A" } }],
      { x: 7.4, y: 4.35, w: 2.0, h: 1.3, align: "center", valign: "middle", fontFace: BF, lineSpacingMultiple: 1.0 });
    // arrows in
    arrow(s, 2.65, 5.0, 3.1, 5.0, TEAL, 3);
    arrow(s, 6.9, 5.0, 7.35, 5.0, TEAL, 3);
    // feeders
    chip(s, 3.15, 7.45, 1.75, 0.75, "TOOLS", CARD, TEALTX, 13);
    chip(s, 5.1, 7.45, 1.75, 0.75, "MEMORY", CARD, TEALTX, 13);
    s.addShape(SH.RECTANGLE, { x: 3.15, y: 7.45, w: 1.75, h: 0.75, fill: { type: "none" }, line: { color: LINE, width: 1 } });
    s.addShape(SH.RECTANGLE, { x: 5.1, y: 7.45, w: 1.75, h: 0.75, fill: { type: "none" }, line: { color: LINE, width: 1 } });
    arrow(s, 4.02, 7.4, 4.4, 6.75, TEAL, 2.25);
    arrow(s, 5.98, 7.4, 5.6, 6.75, TEAL, 2.25);
    note(s, "A goal goes in. The agent thinks, uses tools, checks the result, repeats — then the answer comes out.", 8.75, TEALTX);
  }

  // ============================================================ PART 1 DIVIDER
  divider("PART 1", "The Brain", "What an LLM is — and what it absolutely cannot do on its own.", I.brain);

  // ================================================= SLIDE 5 — WHAT IS AN LLM
  {
    const s = newSlide(BG);
    kicker(s, "START WITH THE BRAIN", TEALTX);
    title(s, "What is an LLM, really?", HEAD);
    lead(s, "Think of a friend who has read the entire internet — but can only talk. That's a Large Language Model.", 1.9, INK);
    card(s, 0.6, 3.0, 8.8, 1.5, CARD);
    badge(s, 0.95, 3.35, 0.8, TEAL, I.comments);
    s.addText([{ text: "It does ONE thing: ", options: { bold: true, color: INK } },
               { text: "guess the next word", options: { bold: true, color: TEALTX } },
               { text: " — over and over, really well.", options: { color: INK } }],
      { x: 2.05, y: 3.2, w: 7.0, h: 0.6, fontFace: BF, fontSize: 18, valign: "middle" });
    s.addText("That's the whole trick. No magic. It's a very, very good autocomplete.", { x: 2.05, y: 3.8, w: 7.0, h: 0.5, fontFace: BF, fontSize: 13.5, color: MUTED });
    term(s, 0.6, 4.85, 8.8, 3.1, "llm.txt", [
      [{ t: "prompt  ", c: TCMT }, { t: '"The capital of Japan is"', c: TYEL }],
      [],
      [{ t: "the model asks itself, again and again:", c: TCMT }],
      [{ t: "  what word most likely comes next?", c: TWHITE }],
      [],
      [{ t: "  → ", c: TGREEN }, { t: '"Tokyo"', c: TCYAN }, { t: "   (then \".\", then stop)", c: TCMT }],
      [],
      [{ t: "output  ", c: TCMT }, { t: '"Tokyo."', c: TGREEN, b: true }],
    ]);
    note(s, "Predict a word, add it, predict the next. Do that fast enough and it looks like thinking.", 8.15, TEALTX);
  }

  // ============================================= SLIDE 6 — WHAT IT CAN'T DO ALONE
  {
    const s = newSlide(BGSOFT);
    kicker(s, "THE CATCH", AMBERTX);
    title(s, "On its own, the brain is stuck in a room", HEAD, 0.95, 28);
    lead(s, "Brilliant, but locked in a box: no phone, no internet, no notepad. By itself an LLM can only do this:", 1.9, MUTED);
    const lims = [
      { ic: I.bolt,  t: "Can't DO anything", d: "It only produces text. It can't send an email, query a database, or open a webpage." },
      { ic: I.mem,   t: "Forgets instantly", d: "Once the chat ends, it remembers nothing. Every conversation starts from a blank slate." },
      { ic: I.warn,  t: "Can be confidently wrong", d: "It predicts likely words — not verified facts. It will state a wrong answer with a straight face." },
    ];
    let y = 2.95; const h = 1.55, g = 0.22;
    lims.forEach((L, i) => {
      card(s, 0.6, y, 8.8, h, BG);
      badge(s, 0.95, y + h / 2 - 0.42, 0.84, AMBER, L.ic);
      s.addText(L.t, { x: 2.1, y: y + 0.24, w: 7.0, h: 0.45, fontFace: BF, fontSize: 17, bold: true, color: INK });
      s.addText(L.d, { x: 2.1, y: y + 0.72, w: 7.05, h: 0.7, fontFace: BF, fontSize: 13.5, color: MUTED, lineSpacingMultiple: 1.06 });
      y += h + g;
    });
    note(s, "So how do we turn this talking brain into something that gets real work done?  →", 8.6, AMBERTX);
  }

  // ============================================= SLIDE 7 — THE LEAP: TOOLS -> AGENT
  {
    const s = newSlide(BG);
    kicker(s, "THE LEAP", TEALTX);
    title(s, "Give the brain hands: tools", HEAD);
    lead(s, "A tool is just a function the model is allowed to call. The model asks; your code runs it and hands back the result.", 1.9, INK);
    term(s, 0.6, 3.0, 8.8, 4.35, "tool-call.json", [
      [{ t: "user  ", c: TCMT }, { t: '"What’s the weather in Pune?"', c: TYEL }],
      [],
      [{ t: "// the LLM can't check weather — so it asks for a tool", c: TCMT }],
      [{ t: "model → ", c: TCYAN }, { t: "{", c: TWHITE }],
      [{ t: '  "tool": ', c: TWHITE }, { t: '"get_weather"', c: TGREEN }, { t: ",", c: TWHITE }],
      [{ t: '  "args": ', c: TWHITE }, { t: '{ "city": "Pune" }', c: TYEL }],
      [{ t: "}", c: TWHITE }],
      [],
      [{ t: "// your code runs it, returns real data", c: TCMT }],
      [{ t: "tool  → ", c: TGREEN }, { t: '{ "tempC": 31, "sky": "clear" }', c: TYEL }],
      [],
      [{ t: "model → ", c: TCYAN }, { t: '"It’s 31°C and clear in Pune."', c: TGREEN, b: true }],
    ], { fs: 13 });
    note(s, "A brain that can call tools and react to the results = an AGENT. That's the whole leap.", 7.65, TEALTX);
  }

  // ============================================================ PART 2 DIVIDER
  divider("PART 2", "The Loop", "The single trick that powers every agent on the planet.", I.sync);

  // ================================================= SLIDE 9 — IT'S ONE LOOP
  {
    const s = newSlide(BG);
    kicker(s, "THE MENTAL MODEL", TEALTX);
    title(s, "Forget the jargon — it's one loop", HEAD);
    lead(s, "Every “agentic” system has the same shape: a goal goes in, the model decides → acts → looks, again and again, until done.", 1.9, INK);
    const fy = 3.5, fh = 1.8;
    const blocks = [
      { x: 0.6, w: 2.6, fill: TEAL, ic: I.inbox, t: "INPUT", d: "a goal or trigger" },
      { x: 3.7, w: 2.6, fill: DARK, ic: I.sync, t: "THE LOOP", d: "reason · act · observe" },
      { x: 6.8, w: 2.6, fill: MINT, ic: I.check2, t: "OUTPUT", d: "answer + actions" },
    ];
    blocks.forEach(b => {
      s.addShape(SH.ROUNDED_RECTANGLE, { x: b.x, y: fy, w: b.w, h: fh, rectRadius: 0.1, fill: { color: b.fill }, line: { type: "none" }, shadow: mkShadow() });
      badge(s, b.x + b.w / 2 - 0.43, fy + 0.24, 0.86, b.fill === MINT ? TEAL : MINT, b.ic);
      s.addText(b.t, { x: b.x, y: fy + 1.05, w: b.w, h: 0.35, align: "center", fontFace: BF, fontSize: 15.5, bold: true, color: b.fill === MINT ? DARK : BG });
      s.addText(b.d, { x: b.x, y: fy + 1.4, w: b.w, h: 0.3, align: "center", fontFace: BF, fontSize: 11.5, color: b.fill === MINT ? "06403A" : "EAF6F7" });
    });
    s.addImage({ data: arrowTeal, x: 3.26, y: fy + fh / 2 - 0.22, w: 0.42, h: 0.42 });
    s.addImage({ data: arrowTeal, x: 6.36, y: fy + fh / 2 - 0.22, w: 0.42, h: 0.42 });
    term(s, 0.6, 5.75, 8.8, 2.25, "loop.py", [
      [{ t: "while ", c: TCYAN }, { t: "not ", c: TCYAN }, { t: "done:", c: TWHITE }],
      [{ t: "    step  = ", c: TWHITE }, { t: "model.reason(state)", c: TGREEN }, { t: "   # think", c: TCMT }],
      [{ t: "    result= ", c: TWHITE }, { t: "run_tool(step)", c: TGREEN }, { t: "        # act", c: TCMT }],
      [{ t: "    state = ", c: TWHITE }, { t: "state + result", c: TGREEN }, { t: "        # observe", c: TCMT }],
    ], { fs: 13 });
    note(s, "The trick: it never answers in one shot. It takes a step, reads what came back, then picks the next move.", 8.2, TEALTX);
  }

  // ================================================= SLIDE 10 — ReAct ZOOM IN
  {
    const s = newSlide(DARK, true);
    kicker(s, "ZOOM INTO THE LOOP", MINT);
    title(s, "Reason → Act → Observe", BG, 0.95, 33);
    s.addText("This is called ReAct. One turn at a time — here's what each word means.", { x: 0.6, y: 2.0, w: 8.8, h: 0.5, fontFace: BF, fontSize: 16, color: "CDEDF0" });
    const steps = [
      { ic: I.brain, t: "REASON", d: "The model thinks out loud: “what should I do next to get closer to the goal?”", c: TEAL },
      { ic: I.bolt,  t: "ACT", d: "It picks one tool and calls it — search, query a DB, run code, send a request.", c: MINT },
      { ic: I.eye,   t: "OBSERVE", d: "It reads the result of that action — then feeds it back in and loops.", c: TEAL },
    ];
    let cy = 2.85; const ch = 1.55, cg = 0.28;
    steps.forEach((st, i) => {
      s.addShape(SH.ROUNDED_RECTANGLE, { x: 0.9, y: cy, w: 8.2, h: ch, rectRadius: 0.12, fill: { color: DARK2 }, line: { color: "11515C", width: 1 } });
      badge(s, 1.25, cy + ch / 2 - 0.46, 0.92, st.c, st.ic);
      s.addText(st.t, { x: 2.6, y: cy + 0.26, w: 6.2, h: 0.45, fontFace: MONO, fontSize: 19, bold: true, color: BG });
      s.addText(st.d, { x: 2.6, y: cy + 0.74, w: 6.25, h: 0.7, fontFace: BF, fontSize: 14, color: "CDEDF0", lineSpacingMultiple: 1.05 });
      if (i < 2) s.addText("↓", { x: 1.55, y: cy + ch - 0.02, w: 0.5, h: cg, align: "center", valign: "middle", fontFace: BF, fontSize: 20, bold: true, color: MINT });
      cy += ch + cg;
    });
    badge(s, 0.9, 8.55, 0.62, MINT, I.redo);
    s.addText("Repeat until the goal is met — each observation shapes the next decision.", { x: 1.7, y: 8.55, w: 7.4, h: 0.62, fontFace: BF, fontSize: 14.5, italic: true, color: MINT, valign: "middle" });
  }

  // ================================================= SLIDE 11 — WORKED EXAMPLE
  {
    const s = newSlide(BG);
    kicker(s, "WATCH IT RUN — A REAL EXAMPLE", TEALTX);
    title(s, "“Book a table for 2, Friday 7pm”", HEAD, 0.95, 28);
    lead(s, "Follow the loop turn by turn. Notice it stops to ask you before doing something it can't undo.", 1.85, MUTED, 0.6, 8.85, 14.5);
    term(s, 0.6, 2.75, 8.8, 5.45, "agent.log", [
      [{ t: "GOAL  ", c: TYEL, b: true }, { t: "book a table for 2, Friday 7pm", c: TWHITE }],
      [{ t: "── turn 1 ────────────────────────", c: TBORDER }],
      [{ t: "reason  ", c: TCYAN }, { t: "need a place; recall: home=downtown, hates loud", c: TWHITE }],
      [{ t: "act     ", c: TGREEN }, { t: "maps_search(\"dinner, downtown\")", c: TYEL }],
      [{ t: "observe ", c: TCMT }, { t: "3 options found", c: TWHITE }],
      [{ t: "── turn 2 ────────────────────────", c: TBORDER }],
      [{ t: "reason  ", c: TCYAN }, { t: "pick the quiet, top-rated one", c: TWHITE }],
      [{ t: "act     ", c: TGREEN }, { t: "check_availability(Fri 19:00)", c: TYEL }],
      [{ t: "observe ", c: TCMT }, { t: "7:00 full · 7:30 open", c: TWHITE }],
      [{ t: "── turn 3  ", c: TBORDER }, { t: "⚠ ask the human", c: TYEL, b: true }, { t: " ────────", c: TBORDER }],
      [{ t: "reason  ", c: TCYAN }, { t: "time changed + booking is irreversible", c: TWHITE }],
      [{ t: "act     ", c: TGREEN }, { t: "ask_user(\"7:30 ok?\")", c: TYEL }, { t: "  → \"yes\"", c: TGREEN }],
      [{ t: "── turn 4 ────────────────────────", c: TBORDER }],
      [{ t: "act     ", c: TGREEN }, { t: "book_table(Fri 19:30)", c: TYEL }, { t: " → #A12", c: TGREEN }],
      [{ t: "done ✓ ", c: TGREEN, b: true }, { t: "added to calendar · loop exits", c: TWHITE }],
    ], { fs: 12.5, ls: 1.24 });
  }

  // ================================================= SLIDE 12 — WHY LOOP > ONE-SHOT
  {
    const s = newSlide(BGSOFT);
    kicker(s, "WHY A LOOP AT ALL?", TEALTX);
    title(s, "Because the real world talks back", HEAD, 0.95, 28);
    lead(s, "One-shot answers can't adapt. A loop reads reality after every step and corrects course — like a GPS, not a printed map.", 1.95, INK);
    const lx = 0.6, rx = 5.15, w = 4.25, y = 3.3, h = 4.3;
    card(s, lx, y, w, h, BG);
    badge(s, lx + 0.35, y + 0.4, 0.85, AMBER, I.signs);
    s.addText("One-shot", { x: lx + 0.35, y: y + 1.4, w: w - 0.7, h: 0.5, fontFace: HF, fontSize: 20, bold: true, color: INK });
    s.addText("“Printed directions”", { x: lx + 0.35, y: y + 1.9, w: w - 0.7, h: 0.4, fontFace: BF, fontSize: 13.5, italic: true, color: AMBERTX });
    s.addText([
      "Guesses everything up front",
      "No way to react if a step fails",
      "One wrong turn ruins the trip",
    ].map(t => ({ text: t, options: { bullet: { indent: 14 }, breakLine: true, paraSpaceAfter: 9 } })),
      { x: lx + 0.35, y: y + 2.5, w: w - 0.6, h: 1.6, fontFace: BF, fontSize: 14, color: INK, lineSpacingMultiple: 1.05 });
    card(s, rx, y, w, h, BG);
    badge(s, rx + 0.35, y + 0.4, 0.85, TEAL, I.compass);
    s.addText("The loop", { x: rx + 0.35, y: y + 1.4, w: w - 0.7, h: 0.5, fontFace: HF, fontSize: 20, bold: true, color: INK });
    s.addText("“Live GPS that reroutes”", { x: rx + 0.35, y: y + 1.9, w: w - 0.7, h: 0.4, fontFace: BF, fontSize: 13.5, italic: true, color: TEALTX });
    s.addText([
      "Takes one step, then looks",
      "Sees the result, adjusts the plan",
      "Recovers from surprises + errors",
    ].map(t => ({ text: t, options: { bullet: { indent: 14 }, breakLine: true, paraSpaceAfter: 9 } })),
      { x: rx + 0.35, y: y + 2.5, w: w - 0.6, h: 1.6, fontFace: BF, fontSize: 14, color: INK, lineSpacingMultiple: 1.05 });
    note(s, "Loops cost more (more model calls) — but they handle the messy, unpredictable real world.", 8.0, TEALTX);
  }

  // ============================================================ PART 3 DIVIDER
  divider("PART 3", "Inside the Box", "The real parts of a production system — tools, memory, and the safety gates.", I.layers);

  // ================================================= SLIDE 14 — 6 LAYERS
  {
    const s = newSlide(BG);
    kicker(s, "THE PRODUCTION ANATOMY", TEALTX);
    title(s, "What's really in the box: 6 layers", HEAD);
    const layers = [
      { ic: I.shield, c: TEAL, t: "1 · Input & guardrails", d: "Auth, validation, PII scrub, prompt-injection checks — before anything runs." },
      { ic: I.brain,  c: DARK, t: "2 · Agentic core", d: "The planner + the reason–act–observe loop. The decision-maker." },
      { ic: I.tools,  c: TEAL, t: "3 · Tools / actions", d: "APIs, DB queries, code, search — how the agent touches the real world." },
      { ic: I.mem,    c: DARK, t: "4 · Memory & state", d: "The context window, a vector store (RAG), and checkpoints to pause/resume." },
      { ic: I.check2, c: TEAL, t: "5 · Output gates", d: "Evaluator (is it correct?) + guardrails (is it safe?) + optional human OK." },
      { ic: I.scope,  c: DARK, t: "6 · Observability", d: "Tracing, logging, evals, cost/token tracking, retries — wraps everything." },
    ];
    let y = 1.8; const h = 1.05, g = 0.1;
    layers.forEach(L => {
      card(s, 0.6, y, 8.8, h, CARD);
      badge(s, 0.85, y + h / 2 - 0.34, 0.68, L.c, L.ic);
      s.addText(L.t, { x: 1.78, y: y + 0.15, w: 7.45, h: 0.4, fontFace: BF, fontSize: 15.5, bold: true, color: INK });
      s.addText(L.d, { x: 1.78, y: y + 0.53, w: 7.47, h: 0.48, fontFace: BF, fontSize: 12.5, color: MUTED, lineSpacingMultiple: 1.0 });
      y += h + g;
    });
    note(s, "The next few slides zoom into the parts people get confused about.", 8.92, TEALTX);
  }

  // ================================================= SLIDE 15 — TOOLS DEEP DIVE
  {
    const s = newSlide(BG);
    kicker(s, "ZOOM IN · TOOLS", TEALTX);
    title(s, "A tool is just a function + a label", HEAD, 0.95, 28);
    lead(s, "You describe what each tool does in plain English. The model reads that and decides when to call it.", 1.9, INK);
    term(s, 0.6, 2.95, 8.8, 3.05, "tools.py", [
      [{ t: "@tool", c: TYEL }],
      [{ t: "def ", c: TCYAN }, { t: "get_weather", c: TGREEN }, { t: "(city: str):", c: TWHITE }],
      [{ t: '    """Get the current weather for a city."""', c: TCMT }],
      [{ t: "    return ", c: TCYAN }, { t: "weather_api.fetch(city)", c: TWHITE }],
      [],
      [{ t: "# the description is the part the model actually reads", c: TCMT }],
    ], { fs: 13 });
    const facts = [
      { ic: I.search, t: "Read tools", d: "look things up: search, fetch a page, query a DB" },
      { ic: I.gears,  t: "Write tools", d: "change things: send email, book, run code" },
    ];
    let x = 0.6; const w = 4.3, y = 6.3, h = 1.5;
    facts.forEach((f, i) => {
      card(s, x, y, w, h, CARD);
      badge(s, x + 0.3, y + h / 2 - 0.38, 0.76, i ? AMBER : TEAL, f.ic);
      s.addText(f.t, { x: x + 1.2, y: y + 0.22, w: w - 1.4, h: 0.4, fontFace: BF, fontSize: 16, bold: true, color: INK });
      s.addText(f.d, { x: x + 1.2, y: y + 0.66, w: w - 1.35, h: 0.7, fontFace: BF, fontSize: 12.5, color: MUTED, lineSpacingMultiple: 1.04 });
      x += w + 0.2;
    });
    note(s, "Good tool descriptions = a smarter agent. Bad ones = it calls the wrong tool.", 8.05, AMBERTX);
  }

  // ================================================= SLIDE 16 — MEMORY 3 KINDS
  {
    const s = newSlide(BG);
    kicker(s, "ZOOM IN · MEMORY", TEALTX);
    title(s, "Three kinds of “memory”", HEAD);
    lead(s, "People lump these together and get confused. They're three different things.", 1.9, MUTED);
    const mems = [
      { c: TEAL, ic: I.brain, t: "Short-term", a: "a sticky note", q: "“What did we just say?”", d: "The context window. This run only — gone when it ends." },
      { c: DARK, ic: I.db,    t: "Long-term", a: "a filing cabinet", q: "“What do I know?”", d: "A vector store (RAG). Facts that survive across runs." },
      { c: MINT, ic: I.mark,  t: "Workflow state", a: "a bookmark", q: "“Where was I?”", d: "Checkpoints. Lets a long job pause and resume later." },
    ];
    const mw = 2.87, mxs = [0.6, 3.57, 6.54], y = 2.55, h = 5.5;
    mems.forEach((M, i) => {
      card(s, mxs[i], y, mw, h, CARD);
      badge(s, mxs[i] + mw / 2 - 0.5, y + 0.45, 1.0, M.c, M.ic);
      s.addText(M.t, { x: mxs[i], y: y + 1.65, w: mw, h: 0.45, align: "center", fontFace: HF, fontSize: 18, bold: true, color: INK });
      s.addText("like " + M.a, { x: mxs[i], y: y + 2.1, w: mw, h: 0.35, align: "center", fontFace: BF, fontSize: 12.5, italic: true, color: M.c === MINT ? MINTTX : TEALTX });
      s.addText(M.q, { x: mxs[i] + 0.2, y: y + 2.65, w: mw - 0.4, h: 0.7, align: "center", fontFace: BF, fontSize: 14, italic: true, bold: true, color: INK, lineSpacingMultiple: 1.05 });
      s.addText(M.d, { x: mxs[i] + 0.22, y: y + 3.5, w: mw - 0.44, h: 1.6, align: "center", fontFace: BF, fontSize: 12.5, color: MUTED, lineSpacingMultiple: 1.12 });
    });
    note(s, "Sticky note = now.  Filing cabinet = forever.  Bookmark = where you paused.", 8.4, TEALTX);
  }

  // ================================================= SLIDE 17 — RAG ZOOM IN
  {
    const s = newSlide(BG);
    kicker(s, "ZOOM IN · LONG-TERM MEMORY (RAG)", TEALTX);
    title(s, "How an AI “looks things up”: RAG", HEAD, 0.95, 27);
    lead(s, "RAG = Retrieval-Augmented Generation. Fancy name, simple idea: find the relevant notes first, then answer using them.", 1.9, INK);
    const steps = [
      { n: "1", t: "Store", d: "Chop your docs into chunks, save them in a vector store." },
      { n: "2", t: "Search", d: "Turn the question into the same form; find the closest chunks." },
      { n: "3", t: "Stuff", d: "Paste those chunks into the prompt as context." },
      { n: "4", t: "Answer", d: "The model replies using real, retrieved facts." },
    ];
    let x = 0.6; const w = 2.1, y = 3.1, h = 2.5, g = 0.13;
    steps.forEach((st, i) => {
      card(s, x, y, w, h, CARD);
      s.addShape(SH.OVAL, { x: x + w / 2 - 0.33, y: y + 0.3, w: 0.66, h: 0.66, fill: { color: i === 3 ? MINT : TEAL }, line: { type: "none" } });
      s.addText(st.n, { x: x + w / 2 - 0.33, y: y + 0.3, w: 0.66, h: 0.66, align: "center", valign: "middle", fontFace: HF, fontSize: 22, bold: true, color: i === 3 ? DARK : BG });
      s.addText(st.t, { x: x, y: y + 1.1, w: w, h: 0.4, align: "center", fontFace: BF, fontSize: 16, bold: true, color: INK });
      s.addText(st.d, { x: x + 0.15, y: y + 1.55, w: w - 0.3, h: 0.85, align: "center", fontFace: BF, fontSize: 11.5, color: MUTED, lineSpacingMultiple: 1.05 });
      if (i < 3) arrow(s, x + w + 0.005, y + h / 2, x + w + g - 0.005, y + h / 2, TEAL, 2.5);
      x += w + g;
    });
    term(s, 0.6, 6.05, 8.8, 1.95, "why-rag.txt", [
      [{ t: "// without RAG", c: TCMT }, { t: "  → the model guesses from old training data", c: TWHITE }],
      [{ t: "// with RAG", c: TCMT }, { t: "    → it answers from ", c: TWHITE }, { t: "your", c: TGREEN, b: true }, { t: " up-to-date documents", c: TWHITE }],
      [{ t: "// bonus", c: TCMT }, { t: "       → fewer made-up answers (“hallucinations”)", c: TWHITE }],
    ], { fs: 12.5 });
  }

  // ================================================= SLIDE 18 — GUARDRAILS vs EVALUATOR
  {
    const s = newSlide(BGSOFT);
    kicker(s, "TWO GATES, NOT ONE", TEALTX);
    title(s, "“Is it good?” vs “Is it safe?”", HEAD, 0.95, 28);
    lead(s, "A result can pass one and fail the other. Keep them separate — they answer different questions.", 1.95, MUTED);
    const lx = 0.6, rx = 5.15, w = 4.25, y = 3.2, h = 4.5;
    card(s, lx, y, w, h, BG);
    badge(s, lx + 0.35, y + 0.4, 0.85, TEAL, I.check2);
    s.addText("Evaluator", { x: lx + 0.35, y: y + 1.4, w: w - 0.7, h: 0.5, fontFace: HF, fontSize: 21, bold: true, color: INK });
    s.addText("“Is it correct / good enough?”", { x: lx + 0.35, y: y + 1.92, w: w - 0.7, h: 0.45, fontFace: BF, fontSize: 13.5, italic: true, color: TEALTX });
    s.addText(["Judges the quality of the answer", "Can trigger a retry / redo", "Often the model grading itself", "Pattern: evaluator-optimizer"]
      .map(t => ({ text: t, options: { bullet: { indent: 14 }, breakLine: true, paraSpaceAfter: 9 } })),
      { x: lx + 0.35, y: y + 2.5, w: w - 0.6, h: 1.8, fontFace: BF, fontSize: 13.5, color: INK, lineSpacingMultiple: 1.05 });
    card(s, rx, y, w, h, BG);
    badge(s, rx + 0.35, y + 0.4, 0.85, DARK, I.shield);
    s.addText("Guardrails", { x: rx + 0.35, y: y + 1.4, w: w - 0.7, h: 0.5, fontFace: HF, fontSize: 21, bold: true, color: INK });
    s.addText("“Is it safe, allowed, well-formed?”", { x: rx + 0.35, y: y + 1.92, w: w - 0.7, h: 0.45, fontFace: BF, fontSize: 13.5, italic: true, color: TEALTX });
    s.addText(["Checks rules + policy + safety", "Blocks leaks (e.g. an SSN)", "A fixed rule, not an opinion", "Runs on BOTH input and output"]
      .map(t => ({ text: t, options: { bullet: { indent: 14 }, breakLine: true, paraSpaceAfter: 9 } })),
      { x: rx + 0.35, y: y + 2.5, w: w - 0.6, h: 1.8, fontFace: BF, fontSize: 13.5, color: INK, lineSpacingMultiple: 1.05 });
    note(s, "Evaluator = a picky teacher grading the work.  Guardrails = a security guard at the door.", 8.1, TEALTX);
  }

  // ================================================= SLIDE 19 — HUMAN IN THE LOOP
  {
    const s = newSlide(BG);
    kicker(s, "THE SAFETY VALVE", TEALTX);
    title(s, "When in doubt, ask a human", HEAD);
    lead(s, "For anything risky or irreversible, the agent should stop and get a yes/no before acting. This is the human-in-the-loop gate.", 1.9, INK);
    term(s, 0.6, 3.0, 8.8, 2.7, "approval.log", [
      [{ t: "agent → ", c: TCYAN }, { t: "about to: ", c: TWHITE }, { t: "transfer $5,000 to vendor #88", c: TYEL }],
      [{ t: "guard → ", c: TYEL }, { t: "⚠ irreversible + high value → needs human OK", c: TWHITE }],
      [{ t: "human → ", c: TGREEN }, { t: "approve", c: TGREEN, b: true }, { t: "  ✓", c: TGREEN }],
      [{ t: "agent → ", c: TCYAN }, { t: "done. logged for audit.", c: TWHITE }],
    ], { fs: 13 });
    const when = [
      { ic: I.coins, t: "Costs money or is irreversible" },
      { ic: I.lock,  t: "Touches sensitive / private data" },
      { ic: I.warn,  t: "Low confidence or unclear request" },
    ];
    let x = 0.6; const w = 2.87, y = 6.1, h = 1.55, g = 0.1;
    when.forEach((c, i) => {
      card(s, x, y, w, h, CARD);
      badge(s, x + w / 2 - 0.38, y + 0.25, 0.76, AMBER, c.ic);
      s.addText(c.t, { x: x + 0.15, y: y + 1.02, w: w - 0.3, h: 0.45, align: "center", fontFace: BF, fontSize: 12.5, bold: true, color: INK, lineSpacingMultiple: 1.0 });
      x += w + g;
    });
    note(s, "Rule of thumb: easy to undo → let it run.  Hard to undo → ask first.", 8.0, TEALTX);
  }

  // ================================================= SLIDE 20 — OBSERVABILITY
  {
    const s = newSlide(BG);
    kicker(s, "YOU CAN'T FIX WHAT YOU CAN'T SEE", TEALTX);
    title(s, "Watch everything: observability", HEAD, 0.95, 28);
    lead(s, "Agents are loops of model calls. Without tracing, a wrong answer is a black box. So you log every step.", 1.9, INK);
    term(s, 0.6, 3.0, 8.8, 3.3, "trace.log", [
      [{ t: "trace ", c: TCMT }, { t: "id=run_42", c: TWHITE }],
      [{ t: "  → reason", c: TCYAN }, { t: "   120 tokens   $0.002   0.4s", c: TCMT }],
      [{ t: "  → tool   ", c: TGREEN }, { t: "  maps_search           0.8s", c: TCMT }],
      [{ t: "  → reason", c: TCYAN }, { t: "    95 tokens   $0.001   0.3s", c: TCMT }],
      [{ t: "  → tool   ", c: TGREEN }, { t: "  book_table            0.6s", c: TCMT }],
      [{ t: "  total ", c: TYEL, b: true }, { t: " 4 steps   215 tokens   $0.003   2.1s", c: TWHITE }],
    ], { fs: 12.5 });
    const items = ["Tracing — every step, in order", "Tokens & cost — per run", "Evals — quality over time", "Retries — when a step fails"];
    let x = 0.6; const w = 4.3, y = 6.55, h = 0.62, g = 0.18;
    items.forEach((t, i) => {
      const xx = 0.6 + (i % 2) * (w + g); const yy = y + Math.floor(i / 2) * (h + 0.15);
      s.addShape(SH.OVAL, { x: xx, y: yy + h / 2 - 0.09, w: 0.18, h: 0.18, fill: { color: TEAL }, line: { type: "none" } });
      s.addText(t, { x: xx + 0.32, y: yy, w: w - 0.32, h: h, valign: "middle", fontFace: BF, fontSize: 13.5, color: INK });
    });
    note(s, "If you log it, you can debug it, price it, and improve it. If you don't, you're guessing.", 8.25, TEALTX);
  }

  // ============================================================ PART 4 DIVIDER
  divider("PART 4", "Workflow vs Agent", "The one distinction that ends the confusion — plus the 5 patterns built on it.", I.route);

  // ================================================= SLIDE 22 — WORKFLOW vs AGENT (keystone)
  {
    const s = newSlide(DARK, true);
    kicker(s, "THE DISTINCTION THAT ENDS THE CONFUSION", MINT);
    title(s, "Workflow vs Agent", BG, 0.95, 36);
    const wx = 0.6, ax = 5.15, ww = 4.25, wy = 2.4, wh = 5.0;
    [
      { x: wx, c: TEAL, ic: I.route, t: "Workflow", s: "YOU fix the path", b: ["Steps decided ahead of time", "The model fills in the blanks", "Predictable, cheap, easy to debug", "Most “AI products” live here"] },
      { x: ax, c: MINT, ic: I.compass, t: "Agent", s: "the MODEL picks the path", b: ["The model decides its own steps", "Reacts + self-corrects at runtime", "Flexible, but harder to control", "Use only when it can't be scripted"] },
    ].forEach(o => {
      s.addShape(SH.ROUNDED_RECTANGLE, { x: o.x, y: wy, w: ww, h: wh, rectRadius: 0.12, fill: { color: DARK2 }, line: { color: "11515C", width: 1 }, shadow: mkShadow() });
      badge(s, o.x + 0.35, wy + 0.35, 0.82, o.c, o.ic);
      s.addText(o.t, { x: o.x + 0.35, y: wy + 1.3, w: ww - 0.7, h: 0.5, fontFace: HF, fontSize: 22, bold: true, color: BG });
      s.addText(o.s, { x: o.x + 0.35, y: wy + 1.85, w: ww - 0.7, h: 0.4, fontFace: MONO, fontSize: 13.5, italic: true, color: o.c });
      s.addText(o.b.map(line => ({ text: line, options: { bullet: { indent: 14 }, breakLine: true, paraSpaceAfter: 8 } })),
        { x: o.x + 0.35, y: wy + 2.45, w: ww - 0.6, h: wh - 2.6, fontFace: BF, fontSize: 13.5, color: "CDEDF0", lineSpacingMultiple: 1.05 });
    });
    s.addText("The one question that settles every case:\nwho decided the steps — you (ahead of time) or the model (at runtime)?",
      { x: 0.6, y: 7.7, w: 8.8, h: 1.1, fontFace: BF, fontSize: 16, italic: true, color: MINT, align: "center", lineSpacingMultiple: 1.12 });
  }

  // ================================================= SLIDE 23 — 5 PATTERNS OVERVIEW
  {
    const s = newSlide(BG);
    kicker(s, "THE BUILDING BLOCKS", TEALTX);
    title(s, "5 patterns you'll see everywhere", HEAD, 0.95, 28);
    lead(s, "Almost every “agentic” product is one of these five — or a mix. We'll zoom into each next.", 1.9, MUTED);
    const pats = [
      { ic: I.link,    t: "Prompt chaining", d: "Output of step 1 feeds step 2, in a fixed order.", w: true },
      { ic: I.branch,  t: "Routing", d: "Sort the input, send it down one of several fixed paths.", w: true },
      { ic: I.stream,  t: "Parallelization", d: "Run several calls at once, then combine the results.", w: true },
      { ic: I.sitemap, t: "Orchestrator–workers", d: "A lead splits a task into subtasks chosen at runtime.", w: false },
      { ic: I.redo,    t: "Evaluator–optimizer", d: "Generate → grade → retry until it's good enough.", w: false },
    ];
    let y = 2.95; const h = 0.97, g = 0.13;
    pats.forEach(P => {
      card(s, 0.6, y, 8.8, h, CARD);
      badge(s, 0.85, y + h / 2 - 0.32, 0.64, P.w ? TEAL : MINT, P.ic);
      s.addText(P.t, { x: 1.72, y: y + 0.12, w: 5.2, h: 0.4, fontFace: BF, fontSize: 15.5, bold: true, color: INK });
      s.addText(P.d, { x: 1.72, y: y + 0.5, w: 6.0, h: 0.4, fontFace: BF, fontSize: 12, color: MUTED });
      chip(s, 8.05, y + h / 2 - 0.25, 1.1, 0.5, P.w ? "WORKFLOW" : "AGENT", P.w ? CARD : MINT, P.w ? TEALTX : DARK, 9.5);
      if (P.w) s.addShape(SH.RECTANGLE, { x: 8.05, y: y + h / 2 - 0.25, w: 1.1, h: 0.5, fill: { type: "none" }, line: { color: LINE, width: 1 } });
      y += h + g;
    });
    note(s, "First three are workflows. Only the last two hand real decisions to the model.", 8.5, TEALTX);
  }

  // ---- shared layout for the 5 pattern zoom-in slides ----
  function patternSlide(cfg) {
    const s = newSlide(BG);
    kicker(s, cfg.kick, TEALTX);
    title(s, cfg.title, HEAD, 0.95, 29);
    lead(s, cfg.lead, 1.95, INK);
    // diagram band
    cfg.diagram(s);
    // terminal example
    term(s, 0.6, cfg.termY || 5.7, 8.8, cfg.termH || 2.3, cfg.fname, cfg.lines, { fs: cfg.fs || 12.5 });
    note(s, cfg.when, (cfg.termY || 5.7) + (cfg.termH || 2.3) + 0.18, cfg.warn ? AMBERTX : TEALTX);
    return s;
  }
  // small labelled node for diagrams
  function node(s, x, y, w, h, txt, fill, txtColor, fs = 13) {
    s.addShape(SH.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { type: "none" }, shadow: mkShadow() });
    s.addText(txt, { x, y, w, h, align: "center", valign: "middle", fontFace: BF, fontSize: fs, bold: true, color: txtColor, lineSpacingMultiple: 0.95 });
  }

  // ---- SLIDE 24 — prompt chaining ----
  patternSlide({
    kick: "PATTERN 1 · WORKFLOW", title: "Prompt chaining",
    lead: "A fixed assembly line. Each step's output becomes the next step's input. You set the order; the model does each step.",
    diagram(s) {
      const y = 3.15, h = 0.95;
      node(s, 0.6, y, 2.0, h, "Draft", TEAL, BG);
      node(s, 3.5, y, 2.0, h, "Improve", TEAL, BG);
      node(s, 6.4, y, 2.0, h, "Translate", MINT, DARK);
      arrow(s, 2.6, y + h / 2, 3.5, y + h / 2, TEALTX, 3);
      arrow(s, 5.5, y + h / 2, 6.4, y + h / 2, TEALTX, 3);
      s.addText("fixed order · you decide the steps", { x: 0.6, y: y + h + 0.15, w: 8.0, h: 0.4, fontFace: MONO, fontSize: 11.5, italic: true, color: MUTED });
    },
    fname: "chain.py",
    lines: [
      [{ t: "draft  = ", c: TWHITE }, { t: "llm(", c: TGREEN }, { t: '"write a tagline"', c: TYEL }, { t: ")", c: TGREEN }],
      [{ t: "better = ", c: TWHITE }, { t: "llm(", c: TGREEN }, { t: '"make it punchier: " + draft', c: TYEL }, { t: ")", c: TGREEN }],
      [{ t: "final  = ", c: TWHITE }, { t: "llm(", c: TGREEN }, { t: '"translate to Hindi: " + better', c: TYEL }, { t: ")", c: TGREEN }],
    ],
    when: "Use when: the task splits into clear, ordered steps. Simple + reliable.",
  });

  // ---- SLIDE 25 — routing ----
  patternSlide({
    kick: "PATTERN 2 · WORKFLOW", title: "Routing",
    lead: "First sort the request, then send it down the right pre-built path — like a receptionist directing visitors.",
    diagram(s) {
      const y = 3.1;
      node(s, 0.6, y + 0.6, 2.1, 0.9, "Classify", DARK, BG);
      node(s, 4.0, y, 4.8, 0.7, "→ billing flow", CARD, TEALTX, 12.5);
      node(s, 4.0, y + 0.85, 4.8, 0.7, "→ tech-support flow", CARD, TEALTX, 12.5);
      node(s, 4.0, y + 1.7, 4.8, 0.7, "→ refund flow", CARD, TEALTX, 12.5);
      [0, 0.85, 1.7].forEach(dy => arrow(s, 2.7, y + 1.05, 4.0, y + 0.35 + dy, TEAL, 2.25));
      [0, 0.85, 1.7].forEach(dy => s.addShape(SH.RECTANGLE, { x: 4.0, y: y + dy, w: 4.8, h: 0.7, fill: { type: "none" }, line: { color: LINE, width: 1 } }));
    },
    termY: 5.85, termH: 2.15, fname: "router.py",
    lines: [
      [{ t: "kind = ", c: TWHITE }, { t: "classify(message)", c: TGREEN }, { t: "   # billing? tech? refund?", c: TCMT }],
      [{ t: "route = ", c: TWHITE }, { t: "HANDLERS[kind]", c: TCYAN }],
      [{ t: "return ", c: TCYAN }, { t: "route(message)", c: TGREEN }],
    ],
    when: "Use when: different request types need different, specialized handling.",
  });

  // ---- SLIDE 26 — parallelization ----
  patternSlide({
    kick: "PATTERN 3 · WORKFLOW", title: "Parallelization",
    lead: "Fire several calls at the same time, then merge the answers. Faster, and great for cross-checking.",
    diagram(s) {
      const y = 3.1;
      node(s, 0.6, y + 0.6, 1.9, 0.9, "Task", DARK, BG);
      ["Check A", "Check B", "Check C"].forEach((t, i) => node(s, 3.7, y + i * 0.78, 2.4, 0.62, t, TEAL, BG, 12.5));
      node(s, 7.3, y + 0.6, 2.1, 0.9, "Merge", MINT, DARK);
      [0, 0.78, 1.56].forEach(dy => arrow(s, 2.5, y + 1.05, 3.7, y + 0.31 + dy, TEALTX, 2.25));
      [0, 0.78, 1.56].forEach(dy => arrow(s, 6.1, y + 0.31 + dy, 7.3, y + 1.05, TEALTX, 2.25));
    },
    termY: 5.85, termH: 2.15, fname: "parallel.py",
    lines: [
      [{ t: "results = ", c: TWHITE }, { t: "await ", c: TCYAN }, { t: "gather(", c: TGREEN }],
      [{ t: "    check_grammar(x), check_facts(x), check_tone(x)", c: TYEL }],
      [{ t: ")", c: TGREEN }, { t: "                  # all at once, then combine", c: TCMT }],
    ],
    when: "Use when: independent subtasks can run together, or you want multiple opinions.",
  });

  // ---- SLIDE 27 — orchestrator-workers ----
  patternSlide({
    kick: "PATTERN 4 · AGENT", title: "Orchestrator–workers",
    lead: "A lead agent breaks a big goal into subtasks AT RUNTIME, hands each to a worker, then assembles the result.",
    diagram(s) {
      const y = 3.05;
      node(s, 3.5, y, 3.0, 0.85, "Orchestrator", MINT, DARK);
      ["Worker 1", "Worker 2", "Worker 3"].forEach((t, i) => node(s, 0.7 + i * 3.0, y + 1.5, 2.6, 0.75, t, TEAL, BG, 13));
      [0, 1, 2].forEach(i => arrow(s, 5.0, y + 0.85, 2.0 + i * 3.0, y + 1.5, MINTTX, 2.25));
    },
    termY: 5.7, termH: 2.3, fname: "orchestrator.log",
    lines: [
      [{ t: "goal  ", c: TCMT }, { t: '"build a landing page"', c: TYEL }],
      [{ t: "lead → ", c: TCYAN }, { t: "splits it (decided now, not pre-written):", c: TWHITE }],
      [{ t: "   worker_1 ", c: TGREEN }, { t: "write copy", c: TWHITE }],
      [{ t: "   worker_2 ", c: TGREEN }, { t: "design layout", c: TWHITE }],
      [{ t: "   worker_3 ", c: TGREEN }, { t: "generate images", c: TWHITE }],
      [{ t: "lead → ", c: TCYAN }, { t: "assembles + reviews → done ✓", c: TGREEN }],
    ],
    when: "Use when: you can't know the subtasks in advance — the model must decide them.",
  });

  // ---- SLIDE 28 — evaluator-optimizer ----
  patternSlide({
    kick: "PATTERN 5 · AGENT", title: "Evaluator–optimizer",
    lead: "One model makes an attempt; another grades it and says what's wrong. Loop until it passes — like a writer + editor.",
    diagram(s) {
      const y = 3.15, h = 0.95;
      node(s, 1.0, y, 2.6, h, "Generate", TEAL, BG);
      node(s, 5.4, y, 2.6, h, "Evaluate", MINT, DARK);
      arrow(s, 3.6, y + h / 2, 5.4, y + h / 2, TEALTX, 3);
      s.addText("retry with feedback", { x: 1.0, y: y + h + 0.12, w: 7.0, h: 0.35, align: "center", fontFace: MONO, fontSize: 11.5, italic: true, color: MUTED });
      arrow(s, 6.7, y + h + 0.05, 2.3, y + h + 0.05, AMBER, 2.25);
    },
    termY: 5.75, termH: 2.25, fname: "loop.py",
    lines: [
      [{ t: "while ", c: TCYAN }, { t: "True", c: TYEL }, { t: ":", c: TWHITE }],
      [{ t: "    draft  = ", c: TWHITE }, { t: "generate(task, feedback)", c: TGREEN }],
      [{ t: "    score, feedback = ", c: TWHITE }, { t: "evaluate(draft)", c: TGREEN }],
      [{ t: "    if ", c: TCYAN }, { t: "score >= ", c: TWHITE }, { t: "0.9", c: TYEL }, { t: ": ", c: TWHITE }, { t: "break", c: TCYAN }, { t: "   # good enough", c: TCMT }],
    ],
    when: "Use when: quality matters and you can describe what “good” looks like.",
  });

  // ================================================= SLIDE 29 — ORCHESTRATOR vs WORKER
  {
    const s = newSlide(BGSOFT);
    kicker(s, "GOING MULTI-AGENT", TEALTX);
    title(s, "Who's the boss? Orchestrator vs Worker", HEAD, 0.95, 25);
    lead(s, "When one agent isn't enough, you split the roles: a coordinator and the doers it manages.", 1.95, MUTED);
    const lx = 0.6, rx = 5.15, w = 4.25, y = 3.3, h = 4.4;
    card(s, lx, y, w, h, BG);
    badge(s, lx + 0.35, y + 0.4, 0.85, DARK, I.sitemap);
    s.addText("Orchestrator", { x: lx + 0.35, y: y + 1.4, w: w - 0.7, h: 0.5, fontFace: HF, fontSize: 20, bold: true, color: INK });
    s.addText("the coordinator", { x: lx + 0.35, y: y + 1.9, w: w - 0.7, h: 0.4, fontFace: BF, fontSize: 13.5, italic: true, color: TEALTX });
    s.addText(["Reads the goal, makes a plan", "Splits it into subtasks", "Assigns each to a worker", "Judges + assembles the result"]
      .map(t => ({ text: t, options: { bullet: { indent: 14 }, breakLine: true, paraSpaceAfter: 9 } })),
      { x: lx + 0.35, y: y + 2.45, w: w - 0.6, h: 1.8, fontFace: BF, fontSize: 13.5, color: INK, lineSpacingMultiple: 1.05 });
    card(s, rx, y, w, h, BG);
    badge(s, rx + 0.35, y + 0.4, 0.85, TEAL, I.users);
    s.addText("Worker", { x: rx + 0.35, y: y + 1.4, w: w - 0.7, h: 0.5, fontFace: HF, fontSize: 20, bold: true, color: INK });
    s.addText("the doer", { x: rx + 0.35, y: y + 1.9, w: w - 0.7, h: 0.4, fontFace: BF, fontSize: 13.5, italic: true, color: TEALTX });
    s.addText(["Gets one assigned subtask", "Uses tools to finish it", "Returns the result up", "Revises if asked"]
      .map(t => ({ text: t, options: { bullet: { indent: 14 }, breakLine: true, paraSpaceAfter: 9 } })),
      { x: rx + 0.35, y: y + 2.45, w: w - 0.6, h: 1.8, fontFace: BF, fontSize: 13.5, color: INK, lineSpacingMultiple: 1.05 });
    note(s, "Like a team lead and their team: one plans and reviews, the others execute.", 8.05, TEALTX);
  }

  // ============================================================ PART 5 DIVIDER
  divider("PART 5", "Building It Well", "How to choose, what it costs, the traps to avoid — and your first checklist.", I.rocket);

  // ================================================= SLIDE 31 — DECISION FLOWCHART
  {
    const s = newSlide(BG);
    kicker(s, "THE GOLDEN RULE", TEALTX);
    title(s, "Workflow first. Agent only if you must.", HEAD, 0.95, 26);
    lead(s, "Start with the simplest thing that works. Add freedom only when the task truly needs it.", 1.9, MUTED);
    const qx = 1.4, qw = 7.2;
    const steps = [
      { t: "Can you write the steps in advance?", yes: "→ use a WORKFLOW. Done.", c: TEAL },
      { t: "Are the steps the same every time?", yes: "→ chaining or routing.", c: TEAL },
      { t: "Do the steps depend on what's found at runtime?", yes: "→ NOW reach for an AGENT.", c: MINT },
    ];
    let y = 3.0; const h = 1.5, g = 0.3;
    steps.forEach((st, i) => {
      s.addShape(SH.ROUNDED_RECTANGLE, { x: qx, y, w: qw, h, rectRadius: 0.1, fill: { color: i === 2 ? DARK : CARD }, line: { color: i === 2 ? DARK : LINE, width: 1 }, shadow: mkShadow() });
      s.addText(st.t, { x: qx + 0.4, y: y + 0.22, w: qw - 0.8, h: 0.55, fontFace: BF, fontSize: 16, bold: true, color: i === 2 ? BG : INK });
      s.addText(st.yes, { x: qx + 0.4, y: y + 0.78, w: qw - 0.8, h: 0.5, fontFace: MONO, fontSize: 13.5, color: i === 2 ? MINT : TEALTX });
      if (i < 2) s.addText("↓  if no", { x: qx, y: y + h + 0.01, w: qw, h: g, align: "center", valign: "middle", fontFace: MONO, fontSize: 12, bold: true, color: MUTED });
      y += h + g;
    });
    note(s, "Agents are powerful but pricier and harder to predict. Don't reach for one out of habit.", 8.35, TEALTX);
  }

  // ================================================= SLIDE 32 — COST & TOKENS
  {
    const s = newSlide(BG);
    kicker(s, "THE BILL", TEALTX);
    title(s, "Every loop costs tokens — and money", HEAD, 0.95, 27);
    lead(s, "Tokens are the chunks of text the model reads and writes. You pay per token. Loops repeat → cost adds up fast.", 1.9, INK);
    term(s, 0.6, 3.0, 8.8, 2.55, "cost.txt", [
      [{ t: '"agentic"', c: TGREEN }, { t: " ≈ 1 token        ", c: TCMT }, { t: '"un-be-liev-able"', c: TGREEN }, { t: " ≈ 4 tokens", c: TCMT }],
      [],
      [{ t: "1 loop  = ", c: TWHITE }, { t: "a few hundred tokens", c: TYEL }],
      [{ t: "20 loops= ", c: TWHITE }, { t: "thousands of tokens", c: TYEL }, { t: "  → watch the bill", c: TCMT }],
    ], { fs: 13 });
    const tips = [
      { ic: I.filter, t: "Cap the loops", d: "set a max number of turns so it can't run forever" },
      { ic: I.mark,   t: "Trim the context", d: "don't resend the whole history every single call" },
      { ic: I.gears,  t: "Right-size the model", d: "use a small model for easy steps, big for hard ones" },
    ];
    let y = 5.85; const h = 0.78, g = 0.13;
    tips.forEach(t => {
      card(s, 0.6, y, 8.8, h, CARD);
      badge(s, 0.85, y + h / 2 - 0.27, 0.54, TEAL, t.ic);
      s.addText([{ text: t.t + "  ", options: { bold: true, color: INK } }, { text: "— " + t.d, options: { color: MUTED } }],
        { x: 1.6, y, w: 7.6, h, valign: "middle", fontFace: BF, fontSize: 13.5 });
      y += h + g;
    });
  }

  // ================================================= SLIDE 33 — PITFALLS
  {
    const s = newSlide(BGSOFT);
    kicker(s, "LEARN FROM OTHERS' BRUISES", AMBERTX);
    title(s, "5 traps beginners fall into", HEAD, 0.95, 28);
    const traps = [
      { t: "Using an agent when a workflow would do", d: "more cost, more bugs, no benefit" },
      { t: "No loop limit", d: "it spins forever, burning tokens" },
      { t: "Trusting the output blindly", d: "no evaluator, no guardrails = surprises" },
      { t: "Giving it too many tools", d: "it gets confused and picks the wrong one" },
      { t: "No tracing", d: "when it breaks, you can't see why" },
    ];
    let y = 2.1; const h = 1.18, g = 0.13;
    traps.forEach((t, i) => {
      card(s, 0.6, y, 8.8, h, BG);
      s.addShape(SH.OVAL, { x: 0.9, y: y + h / 2 - 0.33, w: 0.66, h: 0.66, fill: { color: AMBER }, line: { type: "none" } });
      s.addImage({ data: I.warn, x: 1.06, y: y + h / 2 - 0.17, w: 0.34, h: 0.34 });
      s.addText(t.t, { x: 1.8, y: y + 0.2, w: 7.4, h: 0.45, fontFace: BF, fontSize: 16, bold: true, color: INK });
      s.addText("→ " + t.d, { x: 1.8, y: y + 0.64, w: 7.4, h: 0.4, fontFace: BF, fontSize: 13, color: MUTED });
      y += h + g;
    });
    note(s, "Every one of these comes down to the same fix: start simple, watch it, add limits.", 8.7, AMBERTX);
  }

  // ================================================= SLIDE 34 — FIRST AGENT CHECKLIST
  {
    const s = newSlide(BG);
    kicker(s, "YOUR TURN", TEALTX);
    title(s, "Build your first agent: a checklist", HEAD, 0.95, 27);
    lead(s, "Six small steps. Do them in order and you'll have a working, safe little agent.", 1.9, MUTED);
    term(s, 0.6, 2.85, 8.8, 5.15, "first-agent.md", [
      [{ t: "[1] ", c: TGREEN, b: true }, { t: "Pick ONE clear goal", c: TWHITE }],
      [{ t: "    ", c: TCMT }, { t: "“summarize my unread emails”", c: TYEL }],
      [],
      [{ t: "[2] ", c: TGREEN, b: true }, { t: "Give it 1–3 tools (not 20)", c: TWHITE }],
      [{ t: "    ", c: TCMT }, { t: "list_emails(), read_email()", c: TYEL }],
      [],
      [{ t: "[3] ", c: TGREEN, b: true }, { t: "Write the reason–act–observe loop", c: TWHITE }],
      [{ t: "    ", c: TCMT }, { t: "with a max-turns cap", c: TYEL }],
      [],
      [{ t: "[4] ", c: TGREEN, b: true }, { t: "Add a guardrail + a human OK", c: TWHITE }],
      [{ t: "    ", c: TCMT }, { t: "for anything risky", c: TYEL }],
      [],
      [{ t: "[5] ", c: TGREEN, b: true }, { t: "Log every step (tracing)", c: TWHITE }],
      [{ t: "[6] ", c: TGREEN, b: true }, { t: "Run, read the trace, improve — repeat", c: TWHITE }],
    ], { fs: 12.5, ls: 1.22 });
  }

  // ================================================= SLIDE 35 — GLOSSARY
  {
    const s = newSlide(BGSOFT);
    kicker(s, "POCKET CHEAT-SHEET", TEALTX);
    title(s, "Every term, in one line", HEAD, 0.95, 28);
    lead(s, "Screenshot this one. It's the whole deck, compressed.", 1.9, MUTED);
    const terms = [
      ["LLM", "a brain that predicts the next word"],
      ["Token", "a chunk of text; you pay per token"],
      ["Tool", "a function the model can call"],
      ["Agent", "an LLM that loops with tools"],
      ["ReAct", "reason → act → observe, repeat"],
      ["Context", "the short-term memory of one run"],
      ["RAG", "look up real docs, then answer"],
      ["Guardrail", "a safety rule on in/out"],
      ["Evaluator", "checks if the answer is good"],
      ["Workflow", "steps you fixed in advance"],
      ["Orchestrator", "the agent that splits + assigns"],
      ["Observability", "tracing, tokens, cost, evals"],
    ];
    let y = 2.95; const rh = 0.85, g = 0.1; const cw = 4.3;
    terms.forEach((t, i) => {
      const x = 0.6 + (i % 2) * (cw + 0.2); const row = Math.floor(i / 2);
      const yy = y + row * (rh + g);
      s.addShape(SH.ROUNDED_RECTANGLE, { x, y: yy, w: cw, h: rh, rectRadius: 0.07, fill: { color: BG }, line: { color: LINE, width: 1 } });
      s.addText(t[0], { x: x + 0.18, y: yy + 0.1, w: cw - 0.36, h: 0.32, fontFace: MONO, fontSize: 13.5, bold: true, color: TEALTX });
      s.addText(t[1], { x: x + 0.18, y: yy + 0.42, w: cw - 0.36, h: 0.36, fontFace: BF, fontSize: 12, color: INK });
    });
  }

  // ================================================= SLIDE 36 — CLOSING / CTA
  {
    const s = newSlide(DARK, true);
    s.addShape(SH.OVAL, { x: 7.0, y: 7.0, w: 5.0, h: 5.0, fill: { color: DARK2 }, line: { type: "none" } });
    s.addShape(SH.OVAL, { x: -1.6, y: -1.6, w: 4.4, h: 4.4, fill: { color: DARK2 }, line: { type: "none" } });
    badge(s, 0.6, 1.0, 1.35, MINT, I.flag);
    s.addText("YOU MADE IT", { x: 0.62, y: 2.6, w: 8.8, h: 0.4, fontFace: MONO, fontSize: 14, bold: true, color: MINT, charSpacing: 2 });
    s.addText("You can now explain\nagentic AI.", { x: 0.55, y: 3.0, w: 9, h: 1.8, fontFace: HF, fontSize: 42, bold: true, color: BG, lineSpacingMultiple: 1.0 });
    s.addText("The takeaway: agents are more capable but less predictable. So reach for the simplest thing that works — a workflow first, an agent only when the task genuinely can't be pre-scripted.",
      { x: 0.6, y: 5.05, w: 8.6, h: 1.5, fontFace: BF, fontSize: 17, color: "CDEDF0", lineSpacingMultiple: 1.16 });
    s.addShape(SH.LINE, { x: 0.6, y: 6.95, w: 8.8, h: 0, line: { color: "11515C", width: 1.5 } });
    s.addText([{ text: "Found this useful?\n", options: { bold: true, color: BG, fontSize: 19 } },
               { text: "I'm a developer learning AI engineering in public. Follow + repost so the next person can find it — and tell me what to break down next.", options: { color: "CDEDF0", fontSize: 15 } }],
      { x: 0.6, y: 7.25, w: 8.8, h: 1.7, fontFace: BF, lineSpacingMultiple: 1.18 });
  }

  // ---- page numbers (auto total) ----
  const TOTAL = slides.length;
  slides.forEach((o, i) => pageno(o.s, i + 1, TOTAL, o.onDark));

  await pres.writeFile({ fileName: path.join(__dirname, "Agentic-AI-Explained.pptx") });
  console.log("done — " + TOTAL + " slides");
})();
