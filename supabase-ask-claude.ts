// ask-claude: real-time "Ask Claude" for the DPD RFP Desk.
// The desk inserts a row in claude_requests and calls this function with the request id.
// Claude works the request against the group's SharePoint folder (Microsoft Graph) and the desk (Supabase),
// then writes its reply back to the row. The desk shows the reply live.
//
// Secrets (Supabase > Edge Functions > Secrets): ANTHROPIC_API_KEY, MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET.
// SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY are provided by Supabase.
import { createClient } from "https://esm.sh/@supabase/supabase-js@2.45.4";

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
  { name: "open_pdf", description: "Open a PDF from the group folder so you can read it (all pages, or a page range like '15-20' to keep it small).",
    input_schema: { type: "object", properties: { path: { type: "string" }, pages: { type: "string" } }, required: ["path"] } },
  { name: "read_calculator", description: "Read a range from the group's calculator (06 Decision/DPD Calculator - <group>.xlsx). Sheets: 'Start Here', 'Current Plans', 'Decision', 'Decided Quote'.",
    input_schema: { type: "object", properties: { sheet: { type: "string" }, range: { type: "string" } }, required: ["sheet", "range"] } },
  { name: "write_calculator", description: "Write values into YELLOW input cells of the group's calculator. values is a 2-D array matching the range. Never write over formulas.",
    input_schema: { type: "object", properties: { sheet: { type: "string" }, range: { type: "string" }, values: { type: "array", items: { type: "array" } }, why: { type: "string" } }, required: ["sheet", "range", "values", "why"] } },
  { name: "update_desk", description: "Update the group on the desk. checklist: {item: 'yes'|'partial'|'no'} for items census,sbc,rates,renewal,cap,claims,lca,fonzo,workbook,calculator,proposal. fields: any of effective, quote_no, broker, broker_comp_pepm, broker_comp_note, notes, missing. resolve_waiting: text to match an open Waiting On item to close.",
    input_schema: { type: "object", properties: { checklist: { type: "object" }, fields: { type: "object" }, resolve_waiting: { type: "string" } } } },
];

async function calcPath(g: any) {
  const folder = "06 Decision (chosen quote + group calculator)";
  const kids = await graph(`${itemPath(g.folder_path, folder)}:/children?$select=name`);
  const f = kids.value.find((x: any) => /^DPD Calculator.*\.xlsx$/i.test(x.name));
  if (!f) throw new Error("No calculator in 06 Decision");
  return `${folder}/${f.name}`;
}
const wbUrl = async (g: any, sheet: string, range: string) => `${itemPath(g.folder_path, await calcPath(g))}:/workbook/worksheets('${encodeURIComponent(sheet)}')/range(address='${encodeURIComponent(range)}')`;
const DESK_FIELDS = new Set(["effective", "quote_no", "broker", "broker_comp_pepm", "broker_comp_note", "notes", "missing"]);

async function runTool(name: string, a: any, g: any, who: string, log: any[]): Promise<any> {
  if (name === "list_folder") {
    const r = await graph(`${itemPath(g.folder_path, a.path ? safeRel(a.path) : "")}:/children?$select=name,size,folder,lastModifiedDateTime&$top=200`);
    return r.value.map((x: any) => `${x.folder ? "[folder] " : ""}${x.name}${x.folder ? "" : ` (${Math.round(x.size / 1024)} KB, ${x.lastModifiedDateTime.slice(0, 10)})`}`).join("\n") || "(empty)";
  }
  if (name === "open_pdf") {
    const res = await graph(`${itemPath(g.folder_path, safeRel(a.path))}:/content`, {}, true) as Response;
    const buf = new Uint8Array(await res.arrayBuffer());
    if (buf.length > 30_000_000) return "File too large to open (over 30 MB).";
    let b64 = ""; for (let i = 0; i < buf.length; i += 0x8000) b64 += String.fromCharCode(...buf.subarray(i, i + 0x8000));
    return [{ type: "document", source: { type: "base64", media_type: "application/pdf", data: btoa(b64) }, title: a.path, context: a.pages ? `Focus on pages ${a.pages}.` : undefined }];
  }
  if (name === "read_calculator") {
    const r = await graph(await wbUrl(g, a.sheet, a.range));
    return JSON.stringify({ values: r.values, formulas: r.formulas });
  }
  if (name === "write_calculator") {
    const url = await wbUrl(g, a.sheet, a.range);
    const cur = await graph(url);
    if (cur.formulas.flat().some((f: any) => typeof f === "string" && f.startsWith("="))) return "Refused: that range has formulas. Write only to yellow input cells.";
    await graph(url, { method: "PATCH", body: JSON.stringify({ values: a.values }) });
    log.push({ calc: `${a.sheet}!${a.range}`, before: cur.values, after: a.values, why: a.why });
    await sb.from("activity").insert({ who: `Claude (asked by ${who})`, group_id: g.id, change: `Calculator ${a.sheet}!${a.range}: ${a.why}` });
    return "Written.";
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
F-I current monthly rates EE/ES/EC/FA, J-M renewal monthly rates EE/ES/EC/FA, N-Q enrollment EE/ES/EC/FA, R decided plan (team picks; leave it alone unless asked).
Start Here: B11 effective date, B21 "No renewal yet: compare to current rates?" (Yes/No). Decision tab: Broker commission row, $ PEPM per tier.

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
