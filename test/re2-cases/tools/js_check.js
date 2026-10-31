// Runs the given checks with JavaScript's RegExp, in Re2Cases.bas's output format.
//   node js_check.js IN OUT
// IN has one check per line: pattern and subject in hex (UTF-16 code units, 4 digits each) and
// L (leftmost match) or A (match of ^(?:pattern)$), separated by tabs.
const fs = require("fs");
const [, , inFile, outFile] = process.argv;
const dec = h => (h.match(/.{4}/g) || []).map(x => String.fromCharCode(parseInt(x, 16))).join("");
const enc = s => Array.from({ length: s.length }, (_, i) => s.charCodeAt(i).toString(16).padStart(4, "0")).join("");
const out = [];
for (const line of fs.readFileSync(inFile, "utf8").split("\n")) {
  if (!line) continue;
  const [hp, hs, kind] = line.split("\t");
  const p = dec(hp), s = dec(hs);
  let re;
  try { re = new RegExp(kind === "A" ? "^(?:" + p + ")$" : p); }
  catch (e) { out.push("E"); continue; }
  const m = re.exec(s);
  out.push(m ? m.index + "," + (m.index + m[0].length) + m.slice(1).map(g => "|" + enc(g ?? "")).join("") : "-");
}
fs.writeFileSync(outFile, out.join("\n") + "\n");
