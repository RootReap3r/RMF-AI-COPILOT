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
import os from "os";
import { timingSafeEqual } from "crypto";
import { parseDocument } from "./parse-document.js";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { callLLM, config as llmConfig } from "./llm.js";
import * as store from "./storage.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3001;

// ── CORS ─────────────────────────────────────────────────────────
// Default: local development origins only. Set CORS_ORIGIN in .env
// to a comma-separated list to restrict (e.g. http://localhost:5173).
const corsOrigins = (process.env.CORS_ORIGIN || "http://localhost:5173,http://localhost:3001")
  .split(",").map(o => o.trim()).filter(Boolean);
if (corsOrigins.includes("*")) throw new Error("CORS_ORIGIN must list explicit origins");
app.use(cors({ origin: corsOrigins }));



// ── Auth middleware ───────────────────────────────────────────────
// Required: set AUTH_TOKEN in .env to require X-Api-Key or
// Authorization: Bearer <token> on all /api routes except /api/health.
// Missing or weak configuration stops startup.
const AUTH_TOKEN = process.env.AUTH_TOKEN || "";
if (AUTH_TOKEN.length < 32 || AUTH_TOKEN.startsWith("change-me")) {
  throw new Error("Set AUTH_TOKEN to a strong random token of at least 32 characters before starting");
}

function requireAuth(req, res, next) {
  const provided =
    req.headers["x-api-key"] ||
    (req.headers["authorization"] || "").replace(/^Bearer\s+/i, "");
  if (typeof provided !== "string" || Buffer.byteLength(provided) !== Buffer.byteLength(AUTH_TOKEN) ||
      !timingSafeEqual(Buffer.from(provided), Buffer.from(AUTH_TOKEN))) {
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
app.use(express.json({ limit: "2mb" }));

// ── Document parsing endpoint ─────────────────────────────────────
// Accepts multipart uploads of .docx / .pdf / .txt files.
// Returns { text: "..." } with the combined extracted plain text.
const uploadLimiter = rateLimit({ windowMs: 60000, max: 10 });
let activeParsers = 0;
function acquireParser(req, res, next) {
  if (activeParsers >= 2) return res.status(503).json({ error: "Document parser busy. Please retry." });
  activeParsers++;
  let released = false;
  const release = () => { if (!released) { released = true; activeParsers--; } };
  // Release only after parsing/cleanup finishes, even if the client disconnects.
  req.releaseParser = release;
  next();
}
const upload = multer({
  dest: os.tmpdir(),
  limits: { fileSize: 5 * 1024 * 1024, files: 4, fields: 0, parts: 4 },
  fileFilter: (_req, file, cb) => {
    cb(null, /\.(docx|pdf|txt)$/i.test(file.originalname));
  },
});
app.post("/api/parse-document", uploadLimiter, acquireParser, (req, res, next) => {
  upload.array("files", 4)(req, res, async (err) => {
    try {
      if (err) throw err;
      if (!req.files?.length) return res.status(400).json({ error: "Upload 1–4 DOCX, PDF or TXT files (5 MiB each)." });
      let combined = "";
      for (const file of req.files) {
        if (req.aborted || res.destroyed) return;
        const text = await parseDocument(file.path, path.extname(file.originalname).toLowerCase());
        combined += `\n\n=== SSP SECTION: ${file.originalname} ===\n\n${text}`;
        if (Buffer.byteLength(combined) > 2 * 1024 * 1024) throw new Error("Extracted text exceeds 2 MiB");
      }
      res.json({ text: combined });
    } catch (error) {
      if (error.name === "MulterError") next(error);
      else if (!res.destroyed) res.status(422).json({ error: "Document could not be parsed within resource limits." });
    } finally {
      await Promise.all((req.files || []).map(f => fs.promises.unlink(f.path).catch(() => {})));
      req.releaseParser();
    }
  });
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

app.listen(PORT, process.env.HOST || "127.0.0.1", () => {
  console.log(`\n  NIST Migration Pipeline backend running on http://localhost:${PORT}`);
  console.log(`  LLM provider: ${llmConfig.PROVIDER}  |  model: ${llmConfig.MODEL}`);
  console.log(`  Base URL:     ${llmConfig.BASE_URL}`);
  console.log(`  API key set:  ${llmConfig.hasKey ? "yes" : "NO -- set LLM_API_KEY in .env"}`);
  console.log("  Auth:         required");
  console.log(`  CORS:         ${corsOrigins.join(", ")}\n`);
});

