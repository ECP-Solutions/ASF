// Run each case with JavaScript RegExp; emit Exec captures, matchAll full matches, replace "<$&>".
const fs = require("fs");
const [, , inFile, outFile] = process.argv;
const dec = h => h.match(/.{4}/g)?.map(x => String.fromCharCode(parseInt(x, 16))).join("") ?? "";
const enc = s => [...s].map(c => c.charCodeAt(0).toString(16).padStart(4, "0")).join("");
const out = [];
for (const line of fs.readFileSync(inFile, "utf8").split("\n")) {
  if (!line) continue;
  const [hp, hs, flags] = line.split("\t");
  const p = dec(hp), s = dec(hs);
  let re, reg;
  try { re = new RegExp(p, flags); reg = new RegExp(p, flags + "g"); }
  catch (e) { out.push("ERR"); continue; }
  const m = re.exec(s);
  const ex = m ? "M " + m.map(x => enc(x ?? "")).join(",") : "NOMATCH";
  const all = [...s.matchAll(reg)].map(x => enc(x[0]));
  const rep = s.replace(reg, "<$&>");
  const rep2 = s.replace(reg, "<$&|$1>");
  out.push(ex + "\t" + (all.length ? all.join(",") : "NONE") + "\t" + enc(rep) + "\t" + enc(rep2));
}
fs.writeFileSync(outFile, out.join("\n") + "\n");
