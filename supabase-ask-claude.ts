// ask-claude: real-time "Ask Claude" for the DPD RFP Desk.
// The desk inserts a row in claude_requests and calls this function with the request id.
// Claude works the request against the group's SharePoint folder (Microsoft Graph) and the desk (Supabase),
// then writes its reply back to the row. The desk shows the reply live.
//
// Secrets (Supabase > Edge Functions > Secrets): ANTHROPIC_API_KEY, MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET.
// SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY are provided by Supabase.
import { createClient } from "https://esm.sh/@supabase/supabase-js@2.45.4";
import JSZip from "https://esm.sh/jszip@3.10.1";

const SP_HOSTNAME = "directpaydesigns432.sharepoint.com";
const SP_SITE = "/sites/DirectPayDesignsDocuments";
const SP_LIBRARY = "03_Clients  INTERNAL";
const SP_ROOT = "2026 DPD RFPs/Level-Funded";
const MODEL = "claude-sonnet-5-5";
const cors = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, content-type, apikey, x-client-info" };

const env = (k: string) => Deno.env.get(k) ?? "";
const sb = createClient(env("SUPABASE_URL"), env("SUPABASE_SERVICE_ROLE_KEY"));

// ---------------------------------------------------------------- Microsoft Graph
let token = "", driveId = "";
async function graphToken() {
  if (token) return token;
  const r = await fetch(`https://login.microsoftonline.com/${env("MS_TENANT_ID")}/oauth2/v2.0/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: env("MS_CLIENT_ID"), client_secret: env("MS_CLIENT_SECRET"), scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }),
  });
  const j = await r.json(); if (!j.access_token) throw new Error("Microsoft sign-in failed: " + (j.error_description || j.error));
  return (token = j.access_token);
}
async function graph(path: string, init: RequestInit = {}, raw = false) {
  const r = await fetch("https://graph.microsoft.com/v1.0" + path, { ...init, headers: { Authorization: `Bearer ${await graphToken()}`, "Content-Type": "application/json", ...(init.headers || {}) } });
  if (!r.ok) throw new Error(`SharePoint ${r.status}: ${(await r.text()).slice(0, 300)}`);
  return raw ? r : r.status === 204 ? null : r.json();
}
async function drive() {
  if (driveId) return driveId;
  const site = await graph(`/sites/${SP_HOSTNAME}:${SP_SITE}`);
  const ds = await graph(`/sites/${site.id}/drives`);
  const d = ds.value.find((x: any) => x.name === SP_LIBRARY || decodeURIComponent(x.webUrl).endsWith("/" + SP_LIBRARY));
  if (!d) throw new Error("Library not found: " + SP_LIBRARY);
  return (driveId = d.id);
}
const enc = (p: string) => p.split("/").map(encodeURIComponent).join("/");
const itemPath = (folder: string, rel = "") => `/drives/${driveId}/root:/${enc([SP_ROOT, folder, rel].filter(Boolean).join("/"))}`;
// a group may only touch its own folder
function safeRel(rel: string) { if (!rel || rel.includes("..") || rel.startsWith("/")) throw new Error("Path must be inside the group folder"); return rel; }

// ---------------------------------------------------------------- tools
const TOOLS = [
  { name: "list_folder", description: "List files in the group's SharePoint folder or one of its subfolders (e.g. '02 Current & Renewal'). Empty path = the group folder.",
    input_schema: { type: "object", properties: { path: { type: "string" } } } },
  { name: "open_file", description: "Open any file from the group folder so you can read it: PDF (all pages, or a page range like '15-20'), Excel .xlsx/.xlsm (every sheet as rows of values), Word .docx (text, tables, and any pasted images such as a screenshot of a benefits grid), or an image. Use list_folder first for the exact name.",
    input_schema: { type: "object", properties: { path: { type: "string" }, pages: { type: "string" }, sheet: { type: "string", description: "xlsx only: one sheet name" } }, required: ["path"] } },
  { name: "read_calculator", description: "Read a range from the group's calculator (06 Decision/DPD Calculator - <group>.xlsx). Sheets: 'Start Here', 'Current Plans', 'Decision', 'Decided Quote'.",
    input_schema: { type: "object", properties: { sheet: { type: "string" }, range: { type: "string" } }, required: ["sheet", "range"] } },
  { name: "write_calculator", description: "Write values into YELLOW input cells of the group's calculator. values is a 2-D array matching the range. Never write over formulas.",
    input_schema: { type: "object", properties: { sheet: { type: "string" }, range: { type: "string" }, values: { type: "array", items: { type: "array" } }, why: { type: "string" } }, required: ["sheet", "range", "values", "why"] } },
  { name: "update_desk", description: "Update the group on the desk. checklist: {item: 'yes'|'partial'|'no'} for items census,sbc,rates,renewal,cap,claims,lca,fonzo,workbook,calculator,proposal. fields: any of effective, quote_no, broker, broker_comp_pepm, broker_comp_note, kerix_tier (1-5), notes, missing. resolve_waiting: text to match an open Waiting On item to close.",
    input_schema: { type: "object", properties: { checklist: { type: "object" }, fields: { type: "object" }, resolve_waiting: { type: "string" } } } },
];

async function calcPath(g: any) {
  const folder = "06 Decision (chosen quote + group calculator)";
  const kids = await graph(`${itemPath(g.folder_path, folder)}:/children?$select=name`);
  const f = kids.value.find((x: any) => /^DPD Calculator.*\.xlsx$/i.test(x.name));
  if (!f) throw new Error("No calculator in 06 Decision");
  return `${folder}/${f.name}`;
}
// ---- direct .xlsx editing (Excel Online's workbook API does not accept app-only sign-in) ----
const attr = (tag: string, n: string) => (tag.match(new RegExp(`\\b${n}="([^"]*)"`)) || [])[1];
const unxml = (t: string) => t.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const xml = (t: string) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const colNum = (c: string) => c.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colStr = (n: number) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
function parseRange(rg: string) {
  const [a, b] = rg.replace(/\$/g, "").toUpperCase().split(":");
  const p = (x: string) => { const m = x.match(/^([A-Z]+)(\d+)$/); if (!m) throw new Error("Bad range " + rg); return [colNum(m[1]), +m[2]]; };
  const s0 = p(a), e0 = b ? p(b) : s0; return { c1: s0[0], r1: s0[1], c2: e0[0], r2: e0[1] };
}
async function loadCalc(g: any) {
  const rel = await calcPath(g);
  const res = await graph(`${itemPath(g.folder_path, rel)}:/content`, {}, true) as Response;
  const zip = await JSZip.loadAsync(new Uint8Array(await res.arrayBuffer()));
  return { rel, zip };
}
async function sheetFile(zip: any, sheet: string) {
  const wb = await zip.file("xl/workbook.xml").async("string");
  const tag = [...wb.matchAll(/<sheet\b[^>]*>/g)].map((x: any) => x[0]).find((t: string) => unxml(attr(t, "name") || "") === sheet);
  if (!tag) throw new Error(`No sheet named '${sheet}'`);
  const rid = attr(tag, "r:id");
  const rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const rel = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((x: any) => x[0]).find((t: string) => attr(t, "Id") === rid);
  const tgt = attr(rel, "Target"); return tgt.startsWith("/") ? tgt.slice(1) : "xl/" + tgt.replace(/^\.\//, "");
}
const cellRe = (addr: string) => new RegExp(`<c r="${addr}"(?:\\s[^>]*?)?(?:/>|>[\\s\\S]*?</c>)`);
async function readRange(zip: any, sheet: string, range: string) {
  const sx = await zip.file(await sheetFile(zip, sheet)).async("string");
  const ssx = zip.file("xl/sharedStrings.xml") ? await zip.file("xl/sharedStrings.xml").async("string") : "";
  const ss = [...ssx.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m: any) => unxml([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t: any) => t[1]).join("")));
  const { c1, r1, c2, r2 } = parseRange(range); const values: any[][] = [], formulas: any[][] = [];
  for (let r = r1; r <= r2; r++) {
    const vr: any[] = [], fr: any[] = [];
    for (let c = c1; c <= c2; c++) {
      const m = sx.match(cellRe(colStr(c) + r)); let v: any = null, f: any = null;
      if (m) {
        const t = attr(m[0].split(">")[0], "t"); const fm = m[0].match(/<f[^>]*>([\s\S]*?)<\/f>/) || m[0].match(/<f[^>]*\/>/);
        if (fm) f = "=" + unxml(fm[1] || "");
        const vm = m[0].match(/<v>([\s\S]*?)<\/v>/);
        if (t === "s" && vm) v = ss[+vm[1]]; else if (t === "inlineStr") v = unxml((m[0].match(/<t[^>]*>([\s\S]*?)<\/t>/) || [])[1] || "");
        else if (vm) v = t === "str" || t === "e" ? unxml(vm[1]) : (t === "b" ? vm[1] === "1" : Number(vm[1]));
      }
      vr.push(v); fr.push(f);
    }
    values.push(vr); formulas.push(fr);
  }
  return { values, formulas };
}
function cellXml(addr: string, s: string | undefined, v: any) {
  const st = s ? ` s="${s}"` : "";
  if (v === null || v === undefined || v === "") return `<c r="${addr}"${st}/>`;
  if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}$/.test(v)) v = (Date.UTC(+v.slice(0, 4), +v.slice(5, 7) - 1, +v.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 86400000;
  if (typeof v === "string" && /^-?\$?[\d,]+(\.\d+)?$/.test(v.trim())) v = Number(v.replace(/[$,]/g, ""));
  if (typeof v === "number") return `<c r="${addr}"${st}><v>${v}</v></c>`;
  if (typeof v === "boolean") return `<c r="${addr}"${st} t="b"><v>${v ? 1 : 0}</v></c>`;
  return `<c r="${addr}"${st} t="inlineStr"><is><t xml:space="preserve">${xml(String(v))}</t></is></c>`;
}
async function writeRange(g: any, sheet: string, range: string, values: any[][]) {
  const { rel, zip } = await loadCalc(g);
  const path = await sheetFile(zip, sheet); let sx = await zip.file(path).async("string");
  const { c1, r1, c2, r2 } = parseRange(range);
  if (values.length !== r2 - r1 + 1 || values.some((row) => row.length !== c2 - c1 + 1)) throw new Error("values do not match the range size");
  const before = (await readRange(zip, sheet, range));
  if (before.formulas.flat().some((f: any) => f)) return { refused: true, before };
  for (let r = r1; r <= r2; r++) {
    let rowM = sx.match(new RegExp(`<row r="${r}"(?:\\s[^>]*?)?(?:/>|>[\\s\\S]*?</row>)`));
    if (!rowM) {   // insert a new row in order
      const rows = [...sx.matchAll(/<row r="(\d+)"/g)]; const nxt = rows.find((m: any) => +m[1] > r);
      const nr = `<row r="${r}"></row>`;
      sx = nxt ? sx.slice(0, nxt.index) + nr + sx.slice(nxt.index) : sx.replace("</sheetData>", nr + "</sheetData>").replace("<sheetData/>", `<sheetData>${nr}</sheetData>`);
      rowM = sx.match(new RegExp(`<row r="${r}"(?:\\s[^>]*?)?(?:/>|>[\\s\\S]*?</row>)`));
    }
    let row = rowM![0]; if (row.endsWith("/>")) row = row.slice(0, -2) + "></row>";
    for (let c = c1; c <= c2; c++) {
      const addr = colStr(c) + r; const v = values[r - r1][c - c1]; const m = row.match(cellRe(addr));
      if (m) row = row.replace(m[0], cellXml(addr, attr(m[0].split(">")[0], "s"), v));
      else {
        const cells = [...row.matchAll(/<c r="([A-Z]+)\d+"/g)]; const nxt = cells.find((x: any) => colNum(x[1]) > c);
        const nc = cellXml(addr, undefined, v);
        row = nxt ? row.slice(0, nxt.index) + nc + row.slice(nxt.index) : row.replace("</row>", nc + "</row>");
      }
    }
    sx = sx.replace(rowM![0], row);
  }
  zip.file(path, sx);
  let wbx = await zip.file("xl/workbook.xml").async("string");
  if (/<calcPr\b/.test(wbx)) { if (!/fullCalcOnLoad=/.test(wbx)) wbx = wbx.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"'); }
  else wbx = wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>');
  zip.file("xl/workbook.xml", wbx);
  const out = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  const r = await fetch("https://graph.microsoft.com/v1.0" + `${itemPath(g.folder_path, rel)}:/content`, { method: "PUT", headers: { Authorization: `Bearer ${await graphToken()}`, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }, body: out });
  if (r.status === 423) throw new Error("The calculator is open and locked by someone in Excel. Close it and ask again.");
  if (!r.ok) throw new Error(`SharePoint ${r.status} saving the calculator: ${(await r.text()).slice(0, 200)}`);
  return { refused: false, before };
}

// ---- reading any file type for the model ----
const b64of = (buf: Uint8Array) => { let s = ""; for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode(...buf.subarray(i, i + 0x8000)); return btoa(s); };
const IMG: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp" };
async function openAny(a: any, buf: Uint8Array): Promise<any> {
  const ext = (a.path.split(".").pop() || "").toLowerCase();
  const head = String.fromCharCode(...buf.subarray(0, 5));
  if (head.startsWith("%PDF")) return [{ type: "document", source: { type: "base64", media_type: "application/pdf", data: b64of(buf) }, title: a.path, context: a.pages ? `Focus on pages ${a.pages}.` : undefined }];
  if (IMG[ext]) return [{ type: "image", source: { type: "base64", media_type: IMG[ext], data: b64of(buf) } }];
  if (!head.startsWith("PK")) return `Can't read ${a.path}: it is not a PDF, Excel, Word or image file (.${ext}).`;
  const zip = await JSZip.loadAsync(buf);
  if (zip.file("xl/workbook.xml")) {
    const wb = await zip.file("xl/workbook.xml").async("string");
    const names = [...wb.matchAll(/<sheet\b[^>]*>/g)].map((x: any) => unxml(attr(x[0], "name") || ""));
    let out = `Workbook ${a.path}. Sheets: ${names.join(", ")}\n`;
    for (const nm of names) {
      if (a.sheet && nm !== a.sheet) continue;
      const sx = await zip.file(await sheetFile(zip, nm)).async("string");
      const dim = (sx.match(/<dimension ref="([^"]+)"/) || [])[1];
      const last = [...sx.matchAll(/<c r="([A-Z]+)(\d+)"/g)];
      if (!last.length) { out += `\n== ${nm}: (empty)\n`; continue; }
      const maxC = Math.min(40, Math.max(...last.map((m: any) => colNum(m[1])))), maxR = Math.min(400, Math.max(...last.map((m: any) => +m[2])));
      const { values } = await readRange(zip, nm, `A1:${colStr(maxC)}${maxR}`);
      out += `\n== ${nm} (${dim || ""})\n` + values.map((r, i) => r.some((v) => v !== null && v !== "") ? `${i + 1}\t` + r.map((v) => v ?? "").join("\t") : "").filter(Boolean).join("\n") + "\n";
      if (out.length > 150_000) { out += "\n(truncated: ask for one sheet with the sheet parameter)"; break; }
    }
    return out.slice(0, 160_000);
  }
  if (zip.file("word/document.xml")) {
    const dx = await zip.file("word/document.xml").async("string");
    const text = dx.replace(/<\/w:tc>/g, " | ").replace(/<\/w:p>/g, "\n").replace(/<w:tab\/>/g, "\t").replace(/<[^>]+>/g, "");
    const blocks: any[] = [{ type: "text", text: `Word document ${a.path}:\n` + (unxml(text).trim() || "(no typed text; content is in the images below)") }];
    const imgs = Object.keys(zip.files).filter((f) => /^word\/media\/.+\.(png|jpe?g|gif)$/i.test(f)).slice(0, 6);
    for (const f of imgs) {
      const ib = await zip.file(f).async("uint8array"); if (ib.length > 4_500_000) continue;
      blocks.push({ type: "image", source: { type: "base64", media_type: IMG[f.split(".").pop()!.toLowerCase()], data: b64of(ib) } });
    }
    return blocks;
  }
  return `Can't read ${a.path}: unsupported Office file.`;
}
const DESK_FIELDS = new Set(["effective", "quote_no", "broker", "broker_comp_pepm", "broker_comp_note", "kerix_tier", "notes", "missing"]);

async function runTool(name: string, a: any, g: any, who: string, log: any[]): Promise<any> {
  if (name === "list_folder") {
    const r = await graph(`${itemPath(g.folder_path, a.path ? safeRel(a.path) : "")}:/children?$select=name,size,folder,lastModifiedDateTime&$top=200`);
    return r.value.map((x: any) => `${x.folder ? "[folder] " : ""}${x.name}${x.folder ? "" : ` (${Math.round(x.size / 1024)} KB, ${x.lastModifiedDateTime.slice(0, 10)})`}`).join("\n") || "(empty)";
  }
  if (name === "open_file" || name === "open_pdf") {
    const res = await graph(`${itemPath(g.folder_path, safeRel(a.path))}:/content`, {}, true) as Response;
    const buf = new Uint8Array(await res.arrayBuffer());
    if (buf.length > 30_000_000) return "File too large to open (over 30 MB).";
    return await openAny(a, buf);
  }
  if (name === "read_calculator") {
    const { zip } = await loadCalc(g);
    return JSON.stringify({ ...(await readRange(zip, a.sheet, a.range)), note: "Formula results are as of the file's last save in Excel." });
  }
  if (name === "write_calculator") {
    const w = await writeRange(g, a.sheet, a.range, a.values);
    if (w.refused) return "Refused: that range has formulas. Write only to yellow input cells.";
    log.push({ calc: `${a.sheet}!${a.range}`, before: w.before.values, after: a.values, why: a.why });
    await sb.from("activity").insert({ who: `Claude (asked by ${who})`, group_id: g.id, change: `Calculator ${a.sheet}!${a.range}: ${a.why}` });
    return "Written. Totals recalculate when the file is next opened in Excel.";
  }
  if (name === "update_desk") {
    const out: string[] = [];
    for (const [item, state] of Object.entries(a.checklist || {})) {
      await sb.from("docs").upsert({ group_id: g.id, item, state }, { onConflict: "group_id,item" });
      out.push(`${item} → ${state}`);
    }
    const f = Object.fromEntries(Object.entries(a.fields || {}).filter(([k]) => DESK_FIELDS.has(k)));
    if (Object.keys(f).length) { await sb.from("groups").update({ ...f, updated_by: `Claude (asked by ${who})`, updated_at: new Date().toISOString() }).eq("id", g.id); out.push(JSON.stringify(f)); }
    if (a.resolve_waiting) {
      const { data } = await sb.from("waiting_on").update({ done: true, notes: `Resolved by Claude (asked by ${who})` }).eq("group_id", g.id).eq("done", false).ilike("what", `%${a.resolve_waiting}%`).select("what");
      out.push(`closed waiting: ${(data || []).map((x: any) => x.what).join("; ") || "none matched"}`);
    }
    log.push({ desk: out });
    await sb.from("activity").insert({ who: `Claude (asked by ${who})`, group_id: g.id, change: "Desk: " + out.join(", ") });
    return out.join("\n") || "Nothing changed.";
  }
  return "Unknown tool";
}

const SYSTEM = (g: any) => `You are Claude working the DPD RFP Desk for the group "${g.name}" (SharePoint folder "${g.folder_path}").
A DPD team member typed a request. Do it now with the tools, then answer in 1-3 short sentences: what you found, what you changed, and anything that still needs a person.

Folder layout: 01 Intake, 02 Current & Renewal, 03 Claims, 04 DPD Workbook & FONZO, 05 Quotes, 06 Decision (calculator), 07 Proposal, 08 Sold - Implementation.
Calculator 'Current Plans' tab, one row per incumbent plan, rows 5-10. Inputs: A plan name, B type, C deductible (ind), D out-of-pocket max (ind), E plan pays % (e.g. 0.8),
F-I current monthly rates EE/ES/EC/FA, J-M renewal monthly rates EE/ES/EC/FA, N-Q enrollment EE/ES/EC/FA, S why-this-match note, T PCP copay, U generic Rx copay. Column R (decided plan) is a formula: the decided plan is typed on the Decision tab C28:C33, row 28 = Current Plans row 5, and so on (team picks; leave it alone unless asked).
Current rates are often a per-employee roster (one line per employee with plan and tier): count employees per plan and tier for enrollment and take each plan/tier's monthly total as the rate. Plan designs may be a Word file whose benefits grid is a pasted picture: open_file returns the picture, read it.
If there is no carrier renewal yet, put the current rates in J-M too and set Start Here B21 to Yes; never invent a renewal increase.
Start Here: B11 effective date, D14 Kerix pricing tier 1-5 (Tier 1 $70/$140/$140/$210, 2 $85, 3 $100, 4 $115, 5 $130 per member; ES/EC x2, FA x3; keep it equal to the desk kerix_tier), B21 "No renewal yet: compare to current rates?" (Yes/No). Decision tab: Broker commission row, $ PEPM per tier.

Rules (accuracy is non-negotiable):
- Copy numbers exactly as printed. Read before you write. If the document is ambiguous (which plan is which, tier order, monthly vs annual, composite vs age-banded), do not guess: ask in your reply and change nothing.
- Map tiers carefully: Employee = EE, Employee+Spouse = ES, Employee+Child(ren) = EC, Family = FA. If the carrier uses other tiers, say so and stop.
- Only write to input cells. Never touch formulas. Never move, rename or delete files. Never send anything to a broker or client. Never change a group's status to RELEASED.
- When you load rates or a renewal, also tick the matching checklist items (rates / renewal) and close the matching Waiting On item.
- Quote page numbers and file names in your reply so the team can check you.`;

async function work(id: number) {
  const { data: req } = await sb.from("claude_requests").select("*").eq("id", id).single();
  if (!req || req.status !== "queued") return;
  await sb.from("claude_requests").update({ status: "working", updated_at: new Date().toISOString() }).eq("id", id);
  const { data: g } = await sb.from("groups").select("*").eq("id", req.group_id).single();
  const log: any[] = [];
  try {
    if (!g?.folder_path) throw new Error("This group has no SharePoint folder set on the desk.");
    await drive();
    const msgs: any[] = [{ role: "user", content: req.ask }];
    for (let turn = 0; turn < 20; turn++) {
      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "x-api-key": env("ANTHROPIC_API_KEY"), "anthropic-version": "2023-06-01", "content-type": "application/json" },
        body: JSON.stringify({ model: MODEL, max_tokens: 4000, system: SYSTEM(g), tools: TOOLS, messages: msgs }),
      });
      const j = await r.json(); if (!r.ok) throw new Error("Claude API: " + (j.error?.message || r.status));
      msgs.push({ role: "assistant", content: j.content });
      if (j.stop_reason !== "tool_use") {
        const text = j.content.filter((c: any) => c.type === "text").map((c: any) => c.text).join("\n").trim();
        const needs = /\?\s*$/.test(text) && !log.length;
        await sb.from("claude_requests").update({ status: needs ? "needs_info" : "done", reply: text, actions: log, updated_at: new Date().toISOString() }).eq("id", id);
        return;
      }
      const results = [];
      for (const c of j.content.filter((c: any) => c.type === "tool_use")) {
        let out: any; try { out = await runTool(c.name, c.input, g, req.created_by || "team", log); } catch (e) { out = "Error: " + (e as Error).message; }
        results.push({ type: "tool_result", tool_use_id: c.id, content: out });
      }
      msgs.push({ role: "user", content: results });
    }
    throw new Error("Took too many steps; stopped.");
  } catch (e) {
    await sb.from("claude_requests").update({ status: "error", reply: (e as Error).message, actions: log, updated_at: new Date().toISOString() }).eq("id", id);
  }
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors });
  // only signed-in desk users
  const jwt = (req.headers.get("authorization") || "").replace("Bearer ", "");
  const { data: u } = await sb.auth.getUser(jwt);
  if (!u?.user) return new Response("Sign in first", { status: 401, headers: cors });
  const { id } = await req.json();
  // keep working after we answer the page
  // @ts-ignore EdgeRuntime is provided by Supabase
  EdgeRuntime.waitUntil(work(Number(id)));
  return new Response(JSON.stringify({ ok: true }), { headers: { ...cors, "Content-Type": "application/json" } });
});
