// make-short.js — derive build-short.js (20-slide LinkedIn cut) from the
// verified master build.js. Keeps a whitelist of slide sections, drops the
// rest (dividers, deep-dives, the 5 pattern zoom-ins). Master is untouched.
const fs = require("fs");
const path = require("path");

const src = fs.readFileSync(path.join(__dirname, "build.js"), "utf8").split(/\r?\n/);

// slide numbers to KEEP (20 total) — the tightest complete arc
const KEEP = new Set([1, 3, 5, 6, 7, 9, 10, 11, 14, 16, 18, 22, 23, 29, 31, 32, 33, 34, 35, 36]);

const isBanner = (l) => /^  \/\/ (={6,}|---- )/.test(l);

// preamble = everything before the SLIDE 1 banner (helpers, palette, icons…)
const firstSlide = src.findIndex((l) => isBanner(l) && /SLIDE 1\b/.test(l));
const preamble = src.slice(0, firstSlide);

// split the rest into regions, each starting at a banner line
const regions = [];
for (let i = firstSlide; i < src.length; i++) {
  if (isBanner(src[i])) regions.push({ banner: src[i], lines: [src[i]] });
  else if (regions.length) regions[regions.length - 1].lines.push(src[i]);
}

const kept = regions.filter((r) => {
  if (/page numbers/.test(r.banner)) return true;          // footer (writeFile)
  const m = r.banner.match(/SLIDE (\d+)/);
  return m && KEEP.has(Number(m[1]));                      // chosen slides only
});

let out = preamble.concat(kept.flatMap((r) => r.lines)).join("\n");

// retarget output file + console label so the master file is never overwritten
out = out.replace(/Agentic-AI-Explained\.pptx/g, "Agentic-AI-Explained-Short.pptx");
out = out.replace(/done — " \+ TOTAL \+ " slides/, 'done (SHORT) — " + TOTAL + " slides');

// slide 23 promised "we'll zoom into each next" — those zoom-ins are cut here
out = out.replace("— or a mix. We'll zoom into each next.", "— or a mix of them.");

// header comment
out = out.replace(
  "//  How Agentic AI Actually Works  —  LinkedIn carousel (square 1:1)",
  "//  How Agentic AI Actually Works — SHORT 20-slide LinkedIn cut (square 1:1)\n//  GENERATED from build.js by make-short.js — do not hand-edit; edit build.js."
);

fs.writeFileSync(path.join(__dirname, "build-short.js"), out + "\n");
const n = kept.filter((r) => /SLIDE/.test(r.banner)).length;
console.log(`build-short.js written — ${n} slides kept`);
