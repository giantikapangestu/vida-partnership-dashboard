require("dotenv").config();
const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const { Client } = require("@hubspot/api-client");
const { google } = require("googleapis");

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

const SHEET_RANGE = "'LIST STRATEGIC PARTNER'!A1:P300";
const CACHE_TTL = parseInt(process.env.CACHE_TTL || "300") * 1000;

const USERS = [
  { username: process.env.ADMIN_USERNAME || "vida_partnership", password: process.env.ADMIN_PASSWORD || "VidaPartner2026", name: "VIDA Partnership Admin", role: "admin" },
  { username: process.env.VIEWER_USERNAME || "vida_viewer", password: process.env.VIEWER_PASSWORD || "VidaPartner2026", name: "VIDA Partnership Viewer", role: "viewer" }
];

const sessions = new Map();
function generateToken() { return crypto.randomBytes(32).toString("hex"); }
function createSession(user) {
  const token = generateToken();
  sessions.set(token, { username: user.username, name: user.name, role: user.role, expires: Date.now() + (8 * 60 * 60 * 1000) });
  return token;
}
function validateSession(token) {
  if (!token) return null;
  const session = sessions.get(token);
  if (!session) return null;
  if (Date.now() > session.expires) { sessions.delete(token); return null; }
  return session;
}
function requireAuth(req, res, next) {
  const token = (req.headers["authorization"] || "").replace("Bearer ", "") || req.headers["x-auth-token"];
  const session = validateSession(token);
  if (!session) return res.status(401).json({ error: "Unauthorized" });
  req.user = session;
  next();
}

app.post("/api/login", (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) return res.status(400).json({ error: "Username and password required" });
  const user = USERS.find(u => u.username === username && u.password === password);
  if (!user) return res.status(401).json({ error: "Invalid username or password" });
  const token = createSession(user);
  res.json({ success: true, token, user: { username: user.username, name: user.name, role: user.role } });
});

app.post("/api/logout", (req, res) => {
  const token = (req.headers["authorization"] || "").replace("Bearer ", "");
  if (token) sessions.delete(token);
  res.json({ success: true });
});

app.get("/api/me", requireAuth, (req, res) => res.json({ user: req.user }));

const COL = { name:1, type:2, subType:3, tier:4, useCase:5, owner:6, industry:7, status:10, totalClient:12, existingClient:13, targetClient:14, abpStatus:15 };
let cache = { data: null, ts: 0 };
const hubspot = process.env.HUBSPOT_TOKEN ? new Client({ accessToken: process.env.HUBSPOT_TOKEN }) : null;

function getSheetsClient() {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) return null;
  const auth = new google.auth.JWT({ email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL, key: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"), scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"] });
  return google.sheets({ version: "v4", auth });
}
function parseNum(v) { if (v == null || v === "") return null; const n = Number(String(v).replace(/,/g, "")); return isNaN(n) ? null : n; }
function str(v) { return (v == null || v === "") ? "" : String(v).trim(); }
function rowToPartner(row) {
  return { name: str(row[COL.name]), type: str(row[COL.type]), subType: str(row[COL.subType]), tier: str(row[COL.tier]), useCase: str(row[COL.useCase]), owner: str(row[COL.owner]), industry: str(row[COL.industry]), status: str(row[COL.status]), totalClient: parseNum(row[COL.totalClient]), existingClient: parseNum(row[COL.existingClient]), targetClient: parseNum(row[COL.targetClient]), abpStatus: str(row[COL.abpStatus]) };
}
async function loadFromSheets() {
  const sheets = getSheetsClient();
  if (!sheets || !process.env.GOOGLE_SHEET_ID) return null;
  try {
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: process.env.GOOGLE_SHEET_ID, range: SHEET_RANGE });
    const rows = res.data.values || [];
    if (rows.length < 2) return null;
    return rows.slice(1).map(rowToPartner).filter(p => p.name && p.name !== "Partner Name");
  } catch (err) { console.error("[Sheets]", err.message); return null; }
}
function loadSeedData() {
  try { return JSON.parse(fs.readFileSync(path.join(__dirname, "public", "data", "partners.json"), "utf8")); }
  catch (err) { return []; }
}
async function loadHubSpotFunnel(partners) {
  if (!hubspot) return buildFallbackFunnel(partners);
  try {
    const pipelinesRes = await hubspot.crm.pipelines.pipelinesApi.getAll("deals");
    const pipelines = pipelinesRes.results || [];
    const partnerPipeline = pipelines.find(p => p.label.toLowerCase().includes("partner")) || pipelines[0];
    if (!partnerPipeline) return buildFallbackFunnel(partners);
    const stages = (partnerPipeline.stages || []).sort((a, b) => (a.displayOrder||0)-(b.displayOrder||0)).map(s => ({ id: s.id, label: s.label }));
    const dealsRes = await hubspot.crm.deals.searchApi.doSearch({ filterGroups: [], properties: ["dealstage"], limit: 200, after: 0 });
    const stageCounts = {};
    stages.forEach(s => { stageCounts[s.id] = 0; });
    (dealsRes.results || []).forEach(d => { const sid = d.properties.dealstage; if (sid && stageCounts[sid] !== undefined) stageCounts[sid]++; });
    return { source: "hubspot", stages, stageCounts };
  } catch (err) { return buildFallbackFunnel(partners); }
}
function buildFallbackFunnel(partners) {
  const e = partners.filter(p=>p.status==="Existing").length, n = partners.filter(p=>p.status==="New").length, pi = partners.filter(p=>p.status==="Pipeline").length, t = partners.length||108;
  const stages = [
    {id:"awareness",label:"Awareness / Identified",count:t},{id:"engaged",label:"Pipeline / Engaged",count:pi+e+n},
    {id:"proposed",label:"Proposed (MoU / Deal)",count:e+n},{id:"contracted",label:"Contracted",count:e+n},
    {id:"new",label:"Onboarded / New",count:n},{id:"existing",label:"Existing / Active",count:e}
  ];
  const stageCounts = {};
  stages.forEach(s => { stageCounts[s.id] = s.count; });
  return { source: "sheet", stages, stageCounts };
}
function buildStats(partners) {
  return {
    total: partners.length,
    existing: partners.filter(p=>p.status==="Existing").length,
    newPartners: partners.filter(p=>p.status==="New").length,
    pipeline: partners.filter(p=>p.status==="Pipeline").length,
    solutionCount: partners.filter(p=>p.type==="Solution Partner").length,
    accessCount: partners.filter(p=>p.type==="Access Partner").length,
    abpCount: partners.filter(p=>p.abpStatus==="Support to ABP").length,
    nonAbpCount: partners.filter(p=>p.abpStatus==="Non Support to ABP").length,
    priorityCount: partners.filter(p=>p.tier==="Priority").length,
    regularCount: partners.filter(p=>p.tier==="Regular").length,
    platformCount: partners.filter(p=>p.subType==="Platform Partner (ISV)").length,
    siCount: partners.filter(p=>p.subType==="System Integrator Partner").length,
    resellerCount: partners.filter(p=>p.subType==="Reseller Partner").length,
    developerCount: partners.filter(p=>p.subType==="Developer Partner").length,
  };
}

app.get("/api/dashboard", requireAuth, async (req, res) => {
  if (cache.data && Date.now() - cache.ts < CACHE_TTL) return res.json({ ...cache.data, cached: true });
  try {
    const sheetPartners = await loadFromSheets();
    const partners = sheetPartners || loadSeedData();
    const dataSource = sheetPartners ? "google_sheets" : "seed_json";
    const funnel = await loadHubSpotFunnel(partners);
    const stats = buildStats(partners);
    const payload = { stats, funnel, existing: partners.filter(p=>p.status==="Existing"), newPartners: partners.filter(p=>p.status==="New"), pipeline: partners.filter(p=>p.status==="Pipeline"), abpList: partners.filter(p=>p.abpStatus==="Support to ABP"||p.abpStatus==="Non Support to ABP"), dataSource, lastUpdated: new Date().toISOString(), cached: false };
    cache = { data: payload, ts: Date.now() };
    res.json(payload);
  } catch (err) { res.status(500).json({ error: "Failed to load dashboard data" }); }
});

app.post("/api/refresh", requireAuth, (req, res) => { cache = { data: null, ts: 0 }; res.json({ ok: true }); });
app.get("/health", (req, res) => res.json({ status: "ok" }));
app.get("/login", (req, res) => res.sendFile(path.join(__dirname, "public", "login.html")));
app.get("*", (req, res) => res.sendFile(path.join(__dirname, "public", "index.html")));

const PORT = process.env.PORT || 3000;
app.listen(PORT, "0.0.0.0", () => {
  console.log(`VIDA Dashboard running on http://0.0.0.0:${PORT}`);
  console.log(`  HubSpot:      ${process.env.HUBSPOT_TOKEN ? "✓" : "✗"}`);
  console.log(`  Google Sheets:${process.env.GOOGLE_SHEET_ID ? " ✓" : " ✗"}`);
});
