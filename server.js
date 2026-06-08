/**
 * VIDA Partnership Dashboard — server.js
 *
 * Data flow:
 *   1. Google Sheets  → partner list (live sync of Partner_list.xlsx)
 *   2. HubSpot API    → deal pipeline / funnel stage counts
 *   3. Fallback       → public/data/partners.json (bundled seed data)
 *
 * Endpoints:
 *   GET /api/dashboard   → merged JSON consumed by the frontend
 *   GET /health          → Railway health check
 *   GET *                → serves public/index.html (SPA)
 */

require("dotenv").config();

const express  = require("express");
const cors     = require("cors");
const path     = require("path");
const fs       = require("fs");
const { Client } = require("@hubspot/api-client");
const { google }  = require("googleapis");

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// ─── Constants ────────────────────────────────────────────────────────────────
const SHEET_NAME  = "LIST STRATEGIC PARTNER";
const SHEET_RANGE = `'${SHEET_NAME}'!A1:P300`;
const CACHE_TTL   = parseInt(process.env.CACHE_TTL || "300") * 1000; // ms

// Column index map (0-based) — matches Partner_list.xlsx exactly
const COL = {
  parentId:       0,
  name:           1,
  type:           2,
  subType:        3,
  tier:           4,
  useCase:        5,
  owner:          6,
  industry:       7,
  childCompanies: 8,
  is177:          9,
  status:        10, // "VIDA Partner Management" → Existing / New / Pipeline
  dealIds:       11,
  totalClient:   12,
  existingClient:13,
  targetClient:  14,
  abpStatus:     15, // "ABP / Non ABP"
};

// ─── In-memory cache ──────────────────────────────────────────────────────────
let cache = { data: null, ts: 0 };

// ─── HubSpot client ───────────────────────────────────────────────────────────
const hubspot = process.env.HUBSPOT_TOKEN
  ? new Client({ accessToken: process.env.HUBSPOT_TOKEN })
  : null;

// ─── Google Sheets auth ───────────────────────────────────────────────────────
function getSheetsClient() {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) return null;
  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key:   (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
  });
  return google.sheets({ version: "v4", auth });
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function parseNum(v) {
  if (v == null || v === "") return null;
  const n = Number(String(v).replace(/,/g, ""));
  return isNaN(n) ? null : n;
}
function str(v) {
  return (v == null || v === "") ? "" : String(v).trim();
}

// ─── Parse a sheet row array → partner object ─────────────────────────────────
function rowToPartner(row) {
  return {
    parentId:       str(row[COL.parentId]),
    name:           str(row[COL.name]),
    type:           str(row[COL.type]),
    subType:        str(row[COL.subType]),
    tier:           str(row[COL.tier]),
    useCase:        str(row[COL.useCase]),
    owner:          str(row[COL.owner]),
    industry:       str(row[COL.industry]),
    childCompanies: str(row[COL.childCompanies]),
    is177:          str(row[COL.is177]),
    status:         str(row[COL.status]),      // Existing | New | Pipeline
    dealIds:        str(row[COL.dealIds]),
    totalClient:    parseNum(row[COL.totalClient]),
    existingClient: parseNum(row[COL.existingClient]),
    targetClient:   parseNum(row[COL.targetClient]),
    abpStatus:      str(row[COL.abpStatus]),   // Support to ABP | Non Support to ABP
  };
}

// ─── 1. Load from Google Sheets ───────────────────────────────────────────────
async function loadFromSheets() {
  const sheets = getSheetsClient();
  if (!sheets || !process.env.GOOGLE_SHEET_ID) return null;

  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: SHEET_RANGE,
    });
    const rows = res.data.values || [];
    if (rows.length < 2) return null;

    // Row 0 = header, rows 1+ = data
    const partners = rows.slice(1)
      .map(rowToPartner)
      .filter(p => p.name && p.name !== "Partner Name");

    console.log(`[Sheets] Loaded ${partners.length} partners`);
    return partners;
  } catch (err) {
    console.error("[Sheets] Error:", err.message);
    return null;
  }
}

// ─── 2. Load bundled seed data (fallback) ─────────────────────────────────────
function loadSeedData() {
  try {
    const raw = fs.readFileSync(
      path.join(__dirname, "public", "data", "partners.json"),
      "utf8"
    );
    const data = JSON.parse(raw);
    console.log(`[Seed] Loaded ${data.length} partners from bundled JSON`);
    return data;
  } catch (err) {
    console.error("[Seed] Error:", err.message);
    return [];
  }
}

// ─── 3. HubSpot funnel ────────────────────────────────────────────────────────
async function loadHubSpotFunnel() {
  if (!hubspot) return buildFallbackFunnel([]);

  try {
    // Get all deal pipelines
    const pipelinesRes = await hubspot.crm.pipelines.pipelinesApi.getAll("deals");
    const pipelines = pipelinesRes.results || [];

    // Find the partner acquisition pipeline
    const partnerPipeline =
      pipelines.find(p =>
        p.label.toLowerCase().includes("partner") ||
        p.label.toLowerCase().includes("acquisition")
      ) || pipelines[0];

    if (!partnerPipeline) return buildFallbackFunnel([]);

    const stages = (partnerPipeline.stages || [])
      .sort((a, b) => (a.displayOrder || 0) - (b.displayOrder || 0))
      .map(s => ({ id: s.id, label: s.label, order: s.displayOrder }));

    // Count deals per stage
    const dealsRes = await hubspot.crm.deals.searchApi.doSearch({
      filterGroups: [],
      properties: ["dealstage", "pipeline"],
      limit: 200,
      after: 0,
    });
    const deals = dealsRes.results || [];

    const stageCounts = {};
    stages.forEach(s => { stageCounts[s.id] = 0; });
    deals.forEach(d => {
      const sid = d.properties.dealstage;
      if (sid && stageCounts[sid] !== undefined) stageCounts[sid]++;
    });

    return { source: "hubspot", pipelineId: partnerPipeline.id, stages, stageCounts };
  } catch (err) {
    console.error("[HubSpot] Funnel error:", err.message);
    return buildFallbackFunnel([]);
  }
}

// ─── Fallback funnel built from partner sheet statuses ───────────────────────
function buildFallbackFunnel(partners) {
  const existing = partners.filter(p => p.status === "Existing").length;
  const newP     = partners.filter(p => p.status === "New").length;
  const pipeline = partners.filter(p => p.status === "Pipeline").length;
  const total    = partners.length || 108;

  const stages = [
    { id: "awareness",  label: "Awareness / Identified",  count: total },
    { id: "engaged",    label: "Pipeline / Engaged",       count: pipeline + existing + newP },
    { id: "proposed",   label: "Proposed (MoU / Deal)",    count: existing + newP },
    { id: "contracted", label: "Contracted",               count: existing + newP },
    { id: "new",        label: "Onboarded / New",          count: newP },
    { id: "existing",   label: "Existing / Active",        count: existing },
  ];

  const stageCounts = {};
  stages.forEach(s => { stageCounts[s.id] = s.count; });
  return { source: "sheet", stages, stageCounts };
}

// ─── Build summary stats ──────────────────────────────────────────────────────
function buildStats(partners) {
  const byStatus   = (s)  => partners.filter(p => p.status === s).length;
  const byType     = (t)  => partners.filter(p => p.type === t).length;
  const byAbp      = (a)  => partners.filter(p => p.abpStatus === a).length;
  const byTier     = (t)  => partners.filter(p => p.tier === t).length;
  const bySubType  = (st) => partners.filter(p => p.subType === st).length;

  return {
    total:         partners.length,
    existing:      byStatus("Existing"),
    newPartners:   byStatus("New"),
    pipeline:      byStatus("Pipeline"),
    solutionCount: byType("Solution Partner"),
    accessCount:   byType("Access Partner"),
    abpCount:      byAbp("Support to ABP"),
    nonAbpCount:   byAbp("Non Support to ABP"),
    priorityCount: byTier("Priority"),
    regularCount:  byTier("Regular"),
    platformCount: bySubType("Platform Partner (ISV)"),
    siCount:       bySubType("System Integrator Partner"),
    resellerCount: bySubType("Reseller Partner"),
    developerCount:bySubType("Developer Partner"),
  };
}

// ─── Main API endpoint ────────────────────────────────────────────────────────
app.get("/api/dashboard", async (req, res) => {
  // Serve from cache if fresh
  if (cache.data && Date.now() - cache.ts < CACHE_TTL) {
    return res.json({ ...cache.data, cached: true });
  }

  try {
    // Load partner data: try Sheets first, fall back to seed JSON
    const sheetPartners = await loadFromSheets();
    const partners = sheetPartners || loadSeedData();
    const dataSource = sheetPartners ? "google_sheets" : "seed_json";

    // Load funnel — try HubSpot, fall back to sheet-derived counts
    let funnel;
    try {
      funnel = await loadHubSpotFunnel();
      // If HubSpot returned no meaningful stages, use fallback
      if (!funnel.stages || funnel.stages.length === 0) {
        funnel = buildFallbackFunnel(partners);
      }
    } catch (_) {
      funnel = buildFallbackFunnel(partners);
    }

    const stats    = buildStats(partners);
    const existing = partners.filter(p => p.status === "Existing");
    const newP     = partners.filter(p => p.status === "New");
    const pipeline = partners.filter(p => p.status === "Pipeline");
    const abpList  = partners.filter(p =>
      p.abpStatus === "Support to ABP" || p.abpStatus === "Non Support to ABP"
    );

    const payload = {
      stats,
      funnel,
      existing,
      newPartners: newP,
      pipeline,
      abpList,
      dataSource,
      lastUpdated: new Date().toISOString(),
      cached: false,
    };

    // Store in cache
    cache = { data: payload, ts: Date.now() };

    res.json(payload);
  } catch (err) {
    console.error("[API] Dashboard error:", err);
    res.status(500).json({ error: "Failed to load dashboard data", detail: err.message });
  }
});

// Force-refresh endpoint (clears cache)
app.post("/api/refresh", (req, res) => {
  cache = { data: null, ts: 0 };
  res.json({ ok: true, message: "Cache cleared. Next request will fetch fresh data." });
});

// Health check (Railway uses GET /health)
app.get("/health", (req, res) => res.json({ status: "ok", uptime: process.uptime() }));

// SPA fallback — everything else serves index.html
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, "0.0.0.0", () => {
  console.log(`VIDA Dashboard running on http://0.0.0.0:${PORT}`);
  console.log(`  HubSpot:      ${process.env.HUBSPOT_TOKEN ? "✓ configured" : "✗ not set (funnel uses sheet data)"}`);
  console.log(`  Google Sheets:${process.env.GOOGLE_SHEET_ID ? " ✓ configured" : " ✗ not set (using seed JSON)"}`);
});
