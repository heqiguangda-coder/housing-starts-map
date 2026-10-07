// e-Stat（建築着工統計調査）の月次Excelを取得し、アプリが読める月別データ（data/）に変換する。
// 出典：国土交通省「建築着工統計調査」（e-Stat）
import fs from "node:fs";
import path from "node:path";
import { JSDOM } from "jsdom";
import * as XLSX from "xlsx";

const ROOT = process.cwd();
const DATA = path.join(ROOT, "data");
const FROM = process.env.FROM || "2023-01";
const RECHECK = +(process.env.RECHECK || 3); // 直近何か月分は修正の有無を確認し直すか
const SETS = [
  { tstat: "000001016966", name: "住宅着工統計", tables: { "15": "h15", "16": "h16", "18": "h18" } },
  { tstat: "000001016965", name: "建築物着工統計", tables: { "4-1": "b4-1", "5": "b5", "6-1": "b6-1", "7-1": "b7-1" } },
];
const META = {
  h15: { key: "表15", label: "表15 都道府県別・利用関係別" },
  h16: { key: "表16", label: "表16 都道府県別・建て方別" },
  h18: { key: "表18", label: "表18 都道府県別・構造別・建て方別（新設）" },
  "b4-1": { key: "表4-1", label: "表4-1 建築物 都道府県別・用途別・階数別・構造別（新築）" },
  b5: { key: "表5", label: "表5 建築物 都道府県別・建築主別" },
  "b6-1": { key: "表6-1", label: "表6-1 建築物 都道府県別・構造別" },
  "b7-1": { key: "表7-1", label: "表7-1 建築物 都道府県別・用途別" },
};
const BASE = "https://www.e-stat.go.jp";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const z2h = (s) => String(s).replace(/[０-９－]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));

async function get(url, buf = false) {
  for (let i = 0; i < 4; i++) {
    try {
      const r = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0 (housing-starts-map data updater)" } });
      if (!r.ok) throw new Error("HTTP " + r.status);
      return buf ? Buffer.from(await r.arrayBuffer()) : await r.text();
    } catch (e) {
      if (i === 3) throw new Error(url + " : " + e.message);
      await sleep(4000 * (i + 1));
    }
  }
}

// アプリ本体の読み取り処理（rowsToRecords）をそのまま使う
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8").replace(/<link[^>]+>/g, "");
const dom = new JSDOM(html, { runScripts: "dangerously", pretendToBeVisual: true });
await sleep(500);
const rowsToRecords = dom.window.eval("rowsToRecords");

function parseExcel(buf, ym) {
  const wb = XLSX.read(buf, { type: "buffer" });
  let recs = [];
  for (const n of wb.SheetNames) {
    try {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: "" });
      recs = recs.concat(Array.from(rowsToRecords(rows, ym)));
    } catch (e) { /* 市区町村だけのシートなどは読み飛ばす */ }
  }
  return recs.filter((r) => !r._empty && r.t === ym);
}

function packMonth(recs, ym) {
  const names = [], vi = [], vals = [], units = [""], ui = new Map([["", 0]]);
  for (const r of recs) for (const k of Object.keys(r.dims)) if (!names.includes(k)) { names.push(k); vals.push([]); vi.push(new Map()); }
  const rows = [];
  for (const r of recs) {
    names.forEach((k, j) => {
      const v = r.dims[k];
      if (v === undefined) { rows.push(-1); return; }
      if (!vi[j].has(v)) { vi[j].set(v, vals[j].length); vals[j].push(v); }
      rows.push(vi[j].get(v));
    });
    const u = r.unit || "";
    if (!ui.has(u)) { ui.set(u, units.length); units.push(u); }
    rows.push(r.pref, ui.get(u), r.v);
  }
  return { t: ym, names, vals, units, w: names.length + 3, rows };
}

const idxPath = path.join(DATA, "index.json");
const index = fs.existsSync(idxPath) ? JSON.parse(fs.readFileSync(idxPath, "utf8")) : { tables: {} };
for (const [k, m] of Object.entries(META)) index.tables[k] = { ...(index.tables[k] || { months: {} }), ...m };
let changed = 0;

for (const set of SETS) {
  const list = await get(`${BASE}/stat-search/files?toukei=00600120&tstat=${set.tstat}&cycle=1&layout=datalist`);
  const months = new Map();
  for (const m of list.matchAll(/href="([^"]*month=\d{8}[^"]*)"/g)) {
    const q = new URLSearchParams(m[1].replace(/&amp;/g, "&").split("?")[1]);
    const y = (q.get("year") || "").slice(0, 4), code = q.get("month") || "";
    const mm = +code.slice(-2);
    if (!/^\d{4}$/.test(y) || !(mm >= 1 && mm <= 12)) continue;
    const ym = `${y}-${String(mm).padStart(2, "0")}`;
    if (ym >= FROM) months.set(ym, { y, code });
  }
  const yms = [...months.keys()].sort();
  const recent = new Set(yms.slice(-RECHECK));
  console.log(`${set.name}: ${yms.length}か月（${yms[0]}〜${yms[yms.length - 1]}）`);
  for (const ym of yms) {
    const keys = Object.values(set.tables);
    const done = keys.every((k) => index.tables[k].months[ym]);
    if (done && !recent.has(ym)) continue;
    const { y, code } = months.get(ym);
    await sleep(800);
    const page = await get(`${BASE}/stat-search/files?layout=datalist&cycle=1&toukei=00600120&tstat=${set.tstat}&tclass1val=0&year=${y}0&month=${code}&result_back=1`);
    const found = {};
    for (const b of page.split("stat-dataset_list-item").slice(1)) {
      const text = z2h(b.replace(/<[^>]+>/g, " ").replace(/\s+/g, " "));
      const no = (text.match(/表番号\s*(\d+(?:\s*-\s*\d+)?)/) || [])[1];
      const id = (b.match(/statInfId=(\d+)&(?:amp;)?fileKind=0/) || [])[1];
      if (no && id) found[no.replace(/\s+/g, "")] = id;
    }
    for (const [no, k] of Object.entries(set.tables)) {
      const id = found[no];
      if (!id) { console.log(`  ${ym} 表${no}: ファイルが見つかりません`); continue; }
      const prev = index.tables[k].months[ym];
      if (prev && prev.id === id) continue;
      await sleep(800);
      try {
        const buf = await get(`${BASE}/stat-search/file-download?statInfId=${id}&fileKind=0`, true);
        const recs = parseExcel(buf, ym);
        if (!recs.length) { console.log(`  ${ym} 表${no}: 数値を読み取れませんでした`); continue; }
        const file = `data/${k}/${ym}.json`;
        fs.mkdirSync(path.join(DATA, k), { recursive: true });
        fs.writeFileSync(path.join(ROOT, file), JSON.stringify(packMonth(recs, ym)));
        index.tables[k].months[ym] = { id, n: recs.length, file };
        changed++;
        console.log(`  ${ym} 表${no}: ${recs.length}件`);
      } catch (e) {
        console.log(`  ${ym} 表${no}: 取得に失敗 ${e.message}`);
      }
    }
  }
}
if (changed) {
  index.updated = new Date().toISOString();
  index.source = "国土交通省「建築着工統計調査」（e-Stat）を加工して作成";
  fs.writeFileSync(idxPath, JSON.stringify(index, null, 1));
}
console.log(`更新：${changed}ファイル`);
