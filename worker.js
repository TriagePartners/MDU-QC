// Triage Partners — MDU QC upload API (Cloudflare Worker, free plan)
// Receives tech submissions and writes them into SharePoint via Microsoft Graph.
// Supervisors never touch this: the review page talks to SharePoint directly with their own Microsoft login.

const MAX_PHOTO_BYTES = 8 * 1024 * 1024;
const TICKET_HOURS = 3;
const UUID_RE = /^[0-9a-f-]{36}$/i;
const STATUSES = ["Pending", "Needs Fix", "Approved", "Rejected"];

export default {
  async fetch(request, env) {
    const cors = corsHeaders(request, env);
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });
    const url = new URL(request.url);
    try {
      let res;
      if (url.pathname === "/start" && request.method === "POST") res = await start(request, env);
      else if (url.pathname === "/photo" && request.method === "POST") res = await photo(request, env, url);
      else if (url.pathname === "/submit" && request.method === "POST") res = await submit(request, env);
      else if (url.pathname === "/admin/setup") res = await setup(env, url);
      else res = json({ error: "Not found" }, 404);
      Object.entries(cors).forEach(([k, v]) => res.headers.set(k, v));
      return res;
    } catch (err) {
      console.error(err);
      return json({ error: err.publicMessage || "Server error — please try again." }, err.status || 500, cors);
    }
  },
};

/* ---------- helpers ---------- */
const json = (d, s = 200, h = {}) => new Response(JSON.stringify(d), { status: s, headers: { "content-type": "application/json", ...h } });
const fail = (msg, status = 400) => Object.assign(new Error(msg), { publicMessage: msg, status });
const clean = (v, max = 255) => (v == null ? "" : String(v).replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max));
const cleanML = (v, max = 4000) => (v == null ? "" : String(v).trim().slice(0, max));
const safeSeg = (s) => clean(s, 40).replace(/[^\w\- ]/g, "").replace(/\s+/g, "-") || "x";

function corsHeaders(request, env) {
  const origin = request.headers.get("origin") || "";
  const allowed = (env.ALLOWED_ORIGINS || "").split(",").map((s) => s.trim()).filter(Boolean);
  const ok = allowed.includes(origin) || allowed.includes("*");
  return {
    "access-control-allow-origin": ok ? origin : allowed[0] || "",
    "access-control-allow-methods": "POST, OPTIONS",
    "access-control-allow-headers": "content-type",
    "access-control-max-age": "86400",
    vary: "origin",
  };
}

async function hmac(secret, data) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  const sig = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(data));
  return btoa(String.fromCharCode(...new Uint8Array(sig))).replace(/[+/=]/g, (c) => ({ "+": "-", "/": "_", "=": "" }[c]));
}
// Ticket = base64url(JSON payload) + "." + signature. Issued after the bot check; required for photos + submit.
async function makeTicket(env, payload) {
  const body = btoa(JSON.stringify(payload)).replace(/=+$/, "");
  return body + "." + (await hmac(env.TICKET_SECRET, body));
}
async function readTicket(env, ticket) {
  const [body, sig] = String(ticket || "").split(".");
  if (!body || !sig || sig !== (await hmac(env.TICKET_SECRET, body))) throw fail("Session expired — tap Submit again.", 401);
  const p = JSON.parse(atob(body));
  if (p.exp < Date.now()) throw fail("Session expired — tap Submit again.", 401);
  return p;
}

/* ---------- tech endpoints ---------- */
async function start(request, env) {
  const b = await request.json();
  if (env.TURNSTILE_SECRET) {
    const form = new FormData();
    form.append("secret", env.TURNSTILE_SECRET);
    form.append("response", b.turnstileToken || "");
    const r = await fetch("https://challenges.cloudflare.com/turnstile/v0/siteverify", { method: "POST", body: form });
    if (!(await r.json()).success) throw fail("Bot check failed — refresh the page and try again.", 403);
  }
  const project = clean(b.projectNumber, 40);
  const state = clean(b.state, 2).toUpperCase();
  if (!project || !/^[A-Z]{2}$/.test(state)) throw fail("Project # and state are required.");
  const d = new Date();
  const jobId = crypto.randomUUID();
  const month = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
  const folder = `${env.PHOTO_ROOT}/${month}/${safeSeg(project)}_${state}_${jobId.slice(0, 8)}`;
  const ticket = await makeTicket(env, { jobId, folder, exp: Date.now() + TICKET_HOURS * 3600e3 });
  return json({ jobId, ticket });
}

async function photo(request, env, url) {
  const t = await readTicket(env, url.searchParams.get("t"));
  const n = url.searchParams.get("n") || "";
  if (!/^\d{1,3}$/.test(n)) throw fail("Bad photo number");
  if (!(request.headers.get("content-type") || "").startsWith("image/")) throw fail("Not an image");
  const body = await request.arrayBuffer();
  if (!body.byteLength || body.byteLength > MAX_PHOTO_BYTES) throw fail("Photo too large");
  const name = `photo-${String(+n + 1).padStart(2, "0")}.jpg`;
  await graph(env, `/sites/${await siteId(env)}/drive/root:/${encPath(t.folder + "/" + name)}:/content`, {
    method: "PUT", headers: { "content-type": "image/jpeg" }, body,
  });
  return json({ ok: true, name });
}

async function submit(request, env) {
  const b = await request.json();
  const t = await readTicket(env, b.ticket);
  const tech = clean(b.techName, 80), company = clean(b.company, 120), project = clean(b.projectNumber, 40);
  const state = clean(b.state, 2).toUpperCase(), address = clean(b.address, 200);
  if (!tech || !company || !project || !state || !address) throw fail("Name, company, state, project # and address are required.");
  const lines = (Array.isArray(b.lines) ? b.lines : []).slice(0, 20).map((l) => ({
    code: clean(l.code, 8).toUpperCase(), desc: clean(l.desc, 120), um: clean(l.um, 4).toUpperCase(), qty: Number(l.qty),
  })).filter((l) => /^[A-Z0-9]{3,6}$/.test(l.code) && Number.isFinite(l.qty) && l.qty > 0 && l.qty < 1e6);
  if (!lines.length) throw fail("Add at least one unit code with a quantity.");
  const photoCount = Math.max(0, Math.min(200, parseInt(b.photoCount, 10) || 0));
  if (!photoCount) throw fail("At least one photo is required.");

  const site = await siteId(env);
  const list = encodeURIComponent(env.LIST_NAME);
  // Idempotent: a retry after a dropped connection must not create a second job.
  const existing = await graph(env, `/sites/${site}/lists/${list}/items?$expand=fields($select=JobId)&$filter=fields/JobId eq '${t.jobId}'`, {
    headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" },
  });
  if (existing.value && existing.value.length) return json({ ok: true, num: +existing.value[0].id, duplicate: true });

  let folderUrl = "";
  try { folderUrl = (await graph(env, `/sites/${site}/drive/root:/${encPath(t.folder)}`)).webUrl || ""; } catch {}
  const now = new Date().toISOString();
  const item = await graph(env, `/sites/${site}/lists/${list}/items`, {
    method: "POST",
    body: JSON.stringify({ fields: {
      Title: project, JobId: t.jobId, State: state, Tech: tech, Company: company, TechEmail: clean(b.email, 120),
      Address: address, Apartment: clean(b.apartment, 60),
      Lines: lines.map((l) => `${l.code} × ${l.qty} ${l.um}`).join("; "), LinesJson: JSON.stringify(lines),
      Comments: cleanML(b.comments, 2000), PhotoFolder: t.folder, PhotoFolderUrl: folderUrl, PhotoCount: photoCount,
      Status: "Pending", History: `${now}|${tech}|Submitted|${photoCount} photos`,
    } }),
  });
  return json({ ok: true, num: +item.id });
}

/* ---------- one-time setup: creates the SharePoint list + columns ---------- */
async function setup(env, url) {
  if (!env.SETUP_KEY || url.searchParams.get("key") !== env.SETUP_KEY) throw fail("Not authorized", 401);
  const site = await siteId(env);
  const text = (name, extra = {}) => ({ name, text: { maxLength: 255 }, ...extra });
  const multi = (name) => ({ name, text: { allowMultipleLines: true, linesForEditing: 6, textType: "plain" } });
  const columns = [
    text("JobId", { indexed: true }), text("State"), text("Tech"), text("Company"), text("TechEmail"),
    text("Address"), text("Apartment"), multi("Lines"), multi("LinesJson"), multi("Comments"),
    text("PhotoFolder"), text("PhotoFolderUrl"), { name: "PhotoCount", number: {} },
    { name: "Status", indexed: true, choice: { choices: STATUSES, displayAs: "dropDownMenu" }, defaultValue: { value: "Pending" } },
    text("Reviewer"), { name: "ReviewedAt", dateTime: { format: "dateTime" } }, multi("ReviewNote"), multi("History"),
  ];
  let list;
  try { list = await graph(env, `/sites/${site}/lists/${encodeURIComponent(env.LIST_NAME)}?$expand=columns($select=name)`); } catch (e) { if (e.status !== 404) throw e; }
  const log = [];
  if (!list) {
    list = await graph(env, `/sites/${site}/lists`, { method: "POST", body: JSON.stringify({ displayName: env.LIST_NAME, columns, list: { template: "genericList" } }) });
    log.push(`Created list "${env.LIST_NAME}" with ${columns.length} columns`);
  } else {
    const have = new Set((list.columns || []).map((c) => c.name));
    for (const c of columns.filter((c) => !have.has(c.name))) {
      await graph(env, `/sites/${site}/lists/${list.id}/columns`, { method: "POST", body: JSON.stringify(c) });
      log.push("Added column " + c.name);
    }
    if (!log.length) log.push("List already set up — nothing to do");
  }
  return json({ ok: true, log });
}

/* ---------- Microsoft Graph (app-only) ---------- */
let tokenCache = null, siteCache = null;
async function token(env) {
  if (tokenCache && tokenCache.exp > Date.now() + 60e3) return tokenCache.value;
  const loginBase = env.LOGIN_BASE || "https://login.microsoftonline.com";
  const r = await fetch(`${loginBase}/${env.MS_TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: env.MS_CLIENT_ID, client_secret: env.MS_CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }),
  });
  const d = await r.json();
  if (!r.ok) throw new Error("Token error: " + JSON.stringify(d));
  tokenCache = { value: d.access_token, exp: Date.now() + d.expires_in * 1000 };
  return tokenCache.value;
}
async function graph(env, path, opts = {}) {
  const base = env.GRAPH_BASE || "https://graph.microsoft.com/v1.0";
  const r = await fetch(base + path, { ...opts, headers: { authorization: "Bearer " + (await token(env)), "content-type": "application/json", ...(opts.headers || {}) } });
  if (r.status === 204) return {};
  const d = await r.json().catch(() => ({}));
  if (!r.ok) throw Object.assign(new Error(`Graph ${r.status} ${path}: ${JSON.stringify(d).slice(0, 300)}`), { status: r.status });
  return d;
}
async function siteId(env) {
  if (!siteCache) siteCache = (await graph(env, `/sites/${env.SP_HOSTNAME}:${env.SP_SITE_PATH}`)).id;
  return siteCache;
}
const encPath = (p) => p.split("/").map(encodeURIComponent).join("/");
