// index.js
// ─────────────────────────────────────────────────────────────────
// Backend for the NIST 800-53 Rev4 -> Rev5 AI migration pipeline.
//
//   POST /api/parse-document  -> parse DOCX / PDF / TXT files, return raw text
//   POST /api/chat            -> proxies to configured LLM provider
//   GET  /api/storage?prefix= -> list keys
//   GET  /api/storage/:key    -> get value
//   POST /api/storage         -> set { key, value }
//   DELETE /api/storage/:key  -> delete key
//   POST /api/storage/clear   -> { prefix } delete all keys with prefix
//   GET  /api/health          -> provider config sanity check (public, no auth)
//
// Serves the built frontend from ../client/dist if present.
// ─────────────────────────────────────────────────────────────────

import "dotenv/config";
import express from "express";
import cors from "cors";
import rateLimit from "express-rate-limit";
import multer from "multer";
import mammoth from "mammoth";
import pdfParse from "pdf-parse";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { callLLM, config as llmConfig } from "./llm.js";
import * as store from "./storage.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3001;

// ── CORS ─────────────────────────────────────────────────────────
// Default: allow all origins (local dev). Set CORS_ORIGIN in .env
// to a comma-separated list to restrict (e.g. http://localhost:5173).
const corsOrigins = process.env.CORS_ORIGIN
  ? process.env.CORS_ORIGIN.split(",").map(o => o.trim())
  : "*";
app.use(cors({ origin: corsOrigins }));

app.use(express.json({ limit: "25mb" }));

// ── Auth middleware ───────────────────────────────────────────────
// Opt-in: set AUTH_TOKEN in .env to require X-Api-Key or
// Authorization: Bearer <token> on all /api routes except /api/health.
// If AUTH_TOKEN is not set, auth is disabled (default, local use).
const AUTH_TOKEN = process.env.AUTH_TOKEN || "";

function requireAuth(req, res, next) {
  if (!AUTH_TOKEN) return next();
  const provided =
    req.headers["x-api-key"] ||
    (req.headers["authorization"] || "").replace(/^Bearer\s+/i, "");
  if (provided !== AUTH_TOKEN) {
    return res.status(401).json({ error: "Unauthorized. Set X-Api-Key header." });
  }
  next();
}

// ── Rate limiting on /api/chat ────────────────────────────────────
// Prevents accidental (or intentional) API key abuse.
// Adjust limits via RATE_LIMIT_MAX / RATE_LIMIT_WINDOW_MS in .env.
const chatLimiter = rateLimit({
  windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS || "60000", 10),
  max: parseInt(process.env.RATE_LIMIT_MAX || "120", 10),
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: "Too many LLM requests — please wait before retrying." },
});

// ── NIST SP 800-53 Rev 5 control catalog ─────────────────────────
let CATALOG = {};
try {
  const catalogPath = path.join(__dirname, "data", "nist-catalog.json");
  CATALOG = JSON.parse(fs.readFileSync(catalogPath, "utf-8"));
  console.log(`  Loaded NIST control catalog: ${Object.keys(CATALOG).length} controls/enhancements`);
} catch (e) {
  console.warn("  WARNING: could not load nist-catalog.json — guidance/grounding will be unavailable:", e.message);
}

// ── CCI mapping ───────────────────────────────────────────────────
let CCI_MAP = {};
try {
  const cciPath = path.join(__dirname, "data", "cci-mapping.json");
  CCI_MAP = JSON.parse(fs.readFileSync(cciPath, "utf-8"));
  console.log(`  Loaded CCI mapping: ${Object.keys(CCI_MAP).length} controls`);
} catch (e) {
  console.warn("  No cci-mapping.json found — CCI column in eMASS export will be blank.");
}

// ── Health (public — no auth so Docker health checks work) ────────
app.get("/api/health", (req, res) => {
  res.json({
    ok: true,
    provider: llmConfig.PROVIDER,
    baseUrl: llmConfig.BASE_URL,
    model: llmConfig.MODEL,
    apiKeyConfigured: llmConfig.hasKey,
    authEnabled: !!AUTH_TOKEN,
  });
});

// Apply auth to all remaining /api routes
app.use("/api", requireAuth);

// ── Document parsing endpoint ─────────────────────────────────────
// Accepts multipart uploads of .docx / .pdf / .txt files.
// Returns { text: "..." } with the combined extracted plain text.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }, // 50 MB per file
  fileFilter: (_req, file, cb) => {
    const ext = file.originalname.toLowerCase();
    if (ext.endsWith(".docx") || ext.endsWith(".pdf") || ext.endsWith(".txt")) {
      cb(null, true);
    } else {
      cb(new Error(`Unsupported file type: ${file.originalname}. Accepted: .docx, .pdf, .txt`));
    }
  },
});

app.post("/api/parse-document", upload.array("files", 20), async (req, res) => {
  if (!req.files || req.files.length === 0) {
    return res.status(400).json({ error: "No files uploaded." });
  }
  let combined = "";
  for (const file of req.files) {
    const name = file.originalname;
    try {
      let text = "";
      if (name.toLowerCase().endsWith(".docx")) {
        const result = await mammoth.extractRawText({ buffer: file.buffer });
        text = result.value;
      } else if (name.toLowerCase().endsWith(".pdf")) {
        const result = await pdfParse(file.buffer);
        text = result.text;
      } else {
        text = file.buffer.toString("utf-8");
      }
      combined += `\n\n=== SSP SECTION: ${name} ===\n\n${text}`;
    } catch (e) {
      return res.status(422).json({ error: `Failed to parse ${name}: ${e.message}` });
    }
  }
  res.json({ text: combined });
});

// Handle multer errors (wrong file type, size limit, etc.)
app.use((err, _req, res, _next) => {
  if (err.name === "MulterError" || err.message?.includes("Unsupported file type")) {
    return res.status(400).json({ error: err.message });
  }
  console.error("Unhandled error:", err.message);
  res.status(500).json({ error: "Internal server error." });
});

// ── NIST catalog lookup ───────────────────────────────────────────
app.get("/api/catalog/:id", (req, res) => {
  const id = req.params.id.toUpperCase();
  const entry = CATALOG[id];
  if (!entry) return res.status(404).json({ error: "Control not found in catalog", id });
  res.json({ id, ...entry, ccis: CCI_MAP[id] || [] });
});

// ── LLM proxy ─────────────────────────────────────────────────────
app.post("/api/chat", chatLimiter, async (req, res) => {
  const { system, user, maxTokens } = req.body || {};
  if (!user || typeof user !== "string") {
    return res.status(400).json({ error: "Missing or invalid 'user' field." });
  }
  try {
    const text = await callLLM(system || "", user, maxTokens || 900);
    res.json({ text });
  } catch (err) {
    console.error("LLM call failed:", err.message);
    res.status(502).json({ error: err.message });
  }
});

// ── Storage: list keys by prefix ─────────────────────────────────
app.get("/api/storage", (req, res) => {
  const prefix = req.query.prefix || "";
  res.json({ keys: store.listKeys(prefix) });
});

// ── Storage: get single key ──────────────────────────────────────
app.get("/api/storage/:key", (req, res) => {
  const value = store.get(req.params.key);
  if (value === null) return res.status(404).json({ error: "Not found" });
  res.json({ key: req.params.key, value });
});

// ── Storage: set key ──────────────────────────────────────────────
app.post("/api/storage", (req, res) => {
  const { key, value } = req.body || {};
  if (!key || typeof key !== "string") {
    return res.status(400).json({ error: "Missing or invalid 'key'." });
  }
  if (key.length > 500) {
    return res.status(400).json({ error: "'key' must be 500 characters or fewer." });
  }
  store.set(key, value);
  res.json({ key, value, ok: true });
});

// ── Storage: delete single key ────────────────────────────────────
app.delete("/api/storage/:key", (req, res) => {
  const ok = store.del(req.params.key);
  res.json({ key: req.params.key, deleted: ok });
});

// ── Storage: bulk clear by prefix ─────────────────────────────────
app.post("/api/storage/clear", (req, res) => {
  const { prefix } = req.body || {};
  if (!prefix || typeof prefix !== "string") {
    return res.status(400).json({ error: "Missing or invalid 'prefix'." });
  }
  const count = store.clearPrefix(prefix);
  res.json({ prefix, deleted: count });
});

// ── Serve built frontend ──────────────────────────────────────────
const clientDist = path.join(__dirname, "..", "client", "dist");
app.use(express.static(clientDist));
app.get("*", (req, res, next) => {
  if (req.path.startsWith("/api/")) return next();
  res.sendFile(path.join(clientDist, "index.html"), (err) => {
    if (err) res.status(200).send(
      "Frontend not built yet. Run `npm run build` in /client, " +
      "or run the client dev server separately (see README)."
    );
  });
});

app.listen(PORT, () => {
  console.log(`\n  NIST Migration Pipeline backend running on http://localhost:${PORT}`);
  console.log(`  LLM provider: ${llmConfig.PROVIDER}  |  model: ${llmConfig.MODEL}`);
  console.log(`  Base URL:     ${llmConfig.BASE_URL}`);
  console.log(`  API key set:  ${llmConfig.hasKey ? "yes" : "NO -- set LLM_API_KEY in .env"}`);
  console.log(`  Auth:         ${AUTH_TOKEN ? "enabled (AUTH_TOKEN set)" : "disabled (set AUTH_TOKEN to enable)"}`);
  console.log(`  CORS:         ${corsOrigins === "*" ? "open (all origins)" : corsOrigins.join(", ")}\n`);
});
