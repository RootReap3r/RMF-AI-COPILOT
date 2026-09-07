# NIST 800-53 Rev 4 → Rev 5 Migration Pipeline

AI-driven pipeline: upload SSP content → AI generates a 20-tab SCTM, an
eMASS-ready export, POA&M items, and a 20-family migration tracker.
Progress accumulates incrementally and persists across sessions.

This package is **provider-agnostic** by design. The frontend never
talks to an LLM directly — it calls your own local backend, which
talks to whichever model API you configure. Swapping providers (e.g.
to an Anthropic-compatible endpoint or AWS Bedrock) is a `.env`
change, not a code change.

```
┌──────────────┐      ┌──────────────────┐      ┌─────────────────────┐
│  Browser UI   │─────▶│  Local backend    │─────▶│  Your LLM provider   │
│  (React)      │      │  (Node/Express)   │      │  (configured in .env)│
│               │◀─────│  + storage.json   │◀─────│                      │
└──────────────┘      └──────────────────┘      └─────────────────────┘
```

---

## Quick Start (local machine)

**Requirements:** Node.js 18+ (or Docker)

### Option A — Docker (recommended, one command)

```bash
cd nist-pipeline-app
cp server/.env.example server/.env
# edit server/.env -> set LLM_API_KEY, LLM_BASE_URL, LLM_MODEL for your provider
docker compose up --build
```

Open **http://localhost:3001**

### Option B — without Docker

```bash
cd nist-pipeline-app

# 1. Backend
cd server
cp .env.example .env
# edit .env -> set LLM_API_KEY, LLM_BASE_URL, LLM_MODEL for your provider
npm install
npm start
# Backend now running on http://localhost:3001

# 2. Frontend (separate terminal)
cd ../client
npm install
npm run dev
# Open http://localhost:5173 (Vite dev server, proxies /api to :3001)
```

For a single deployable build (no separate dev server):

```bash
cd client && npm install && npm run build
cd ../server && npm install && npm start
# Open http://localhost:3001 -- backend now serves the built frontend too
```

---

## Switching LLM Providers

Everything is controlled by `server/.env`. No code changes needed.

### Anthropic API directly
```ini
LLM_PROVIDER=anthropic
LLM_BASE_URL=https://api.anthropic.com
LLM_MODEL=claude-sonnet-4-20250514
LLM_API_KEY=sk-ant-your-key
```

### OpenAI-compatible endpoint
```ini
LLM_PROVIDER=openai
LLM_BASE_URL=https://api.openai.com/v1
LLM_MODEL=gpt-4o
LLM_API_KEY=sk-your-key
```

### AWS Bedrock (NIPR / GovCloud)
Bedrock's native API shape differs slightly from both adapters above.
Two ways to use it:

1. **Recommended:** put a small Bedrock-compatible proxy in front
   (many teams already run one) that exposes either an OpenAI-style
   `/chat/completions` or Anthropic-style `/v1/messages` endpoint —
   then point `LLM_BASE_URL` at that proxy and pick the matching
   `LLM_PROVIDER`.
2. **Direct integration:** add a `callBedrock()` function to
   `server/llm.js` using `@aws-sdk/client-bedrock-runtime` with the
   `InvokeModel` API for an Anthropic Claude model on Bedrock. AWS
   credentials would come from the instance role / IAM, not
   `LLM_API_KEY`. This is a small, contained change — only
   `server/llm.js` needs to be touched.

If your endpoint uses a different request/response shape entirely,
add a new case in `server/llm.js` following the pattern of the
existing two adapters — each is ~20 lines. This is the **only** file
that ever needs to change to support a new provider.

---

## Output Formats (Fixed, Not Configurable)

To keep this simple across teams, output formats are standardized
rather than user-configurable:

- **SCTM** — always a 20-tab workbook (one tab per NIST control
  family), fixed 10-column header row including a **Guidance** column
  populated with the *official NIST SP 800-53 Rev 5 Discussion text*
  for each control (sourced from `server/data/nist-catalog.json`, not
  AI-generated).
- **eMASS export** — one row per **base control** (e.g. `AC-2`), with
  any documented enhancements (`AC-2(1)`, `AC-2(2)`, ...) rolled up
  into the `Implementation_Narrative`, hard-capped at 2,000 characters
  per the eMASS field limit. Includes a `CCI` column populated from
  `server/data/cci-mapping.json`.

### Updating the CCI Mapping with Real DISA Data

`server/data/cci-mapping.json` ships with a small starter mapping
(placeholder-format CCI numbers for the demo control families) so the
CCI column has something to show out of the box. For real use:

1. Get `U_CCI_List.xml` from `public.cyber.mil` -> STIGs -> CCI (CAC
   login required)
2. Run:
   ```bash
   python3 server/scripts/parse_cci_xml.py /path/to/U_CCI_List.xml > server/data/cci-mapping.json
   ```
3. Restart the server — no code changes needed.

### Grounding: Why Gap Analysis Uses the Official Catalog

Every per-control AI call includes the official NIST control text and
discussion from `nist-catalog.json` as context, so the AI evaluates
your SSP statement against the *actual* control requirements (e.g.,
all sub-parts a through l of AC-2) rather than a paraphrase from
training data. Expect gap analysis to be **stricter and more
accurate** than a model working from memory alone.

---

## Where Progress Is Stored

`server/data/storage.json` — one entry per documented control
(`ctrl:AC-2`, `ctrl:AC-3(1)`, etc.), holding the AI-generated SCTM
status, implementation narrative, gap analysis, NIST guidance, CCIs,
and eMASS fields.

- Survives restarts (mounted as a Docker volume in `docker-compose.yml`)
- Back it up like any file — copy it, version it in git (it contains
  no credentials, only control data), or move it between environments
- The dashboard's **"Reset all progress"** button clears this file via
  the API — useful for starting fresh on a new system
- For team/shared deployments, replace `server/storage.js` with a real
  database (Postgres, DynamoDB, etc.) — it's a 4-function module
  (`get`, `set`, `del`, `listKeys`) so the swap is contained

---

## Deploying to AWS (NIPR / GovCloud)

The Docker image is environment-agnostic. Typical paths:

- **ECS Fargate / EC2:** push the image to ECR, run as a task/service,
  mount an EFS volume at `/app/server/data` for persistence, set
  `LLM_*` env vars via task definition or Secrets Manager
- **App Runner:** point at the ECR image, set env vars in the console
- **EC2 + Docker Compose:** simplest — `docker compose up -d` on an
  instance, EBS volume for `./server/data`

None of this requires internet egress beyond whatever URL you put in
`LLM_BASE_URL` — if that's an internal VPC endpoint, the container
never needs to leave your network boundary.

---

## What's Inside

```
nist-pipeline-app/
├── server/
│   ├── index.js              # Express API: /api/chat, /api/storage/*, /api/catalog/:id
│   ├── llm.js                 # Provider adapters (THE file to extend for new providers)
│   ├── storage.js              # File-based KV store (swap for a DB if needed)
│   ├── .env.example            # Copy to .env and configure (see Quick Start)
│   ├── data/
│   │   ├── storage.json         # Documented controls — gitignored, persists progress
│   │   ├── nist-catalog.json     # Official SP 800-53 Rev 5 control catalog (1,189 entries)
│   │   └── cci-mapping.json      # Control -> CCI mapping (starter/demo data)
│   └── scripts/
│       └── parse_cci_xml.py      # Converts official DISA U_CCI_List.xml -> cci-mapping.json
├── client/
│   └── src/App.jsx              # The full pipeline UI
├── Dockerfile
├── docker-compose.yml
└── README.md
```

---

## Notes on Classified Environments

This package is just code — it makes no assumptions about
classification level. What determines whether it's appropriate for a
given network is:

1. Whatever LLM endpoint you point `LLM_BASE_URL` at, and whether
   that endpoint is approved for that environment's data
2. Standard ATO/accreditation review for any new tool running on that
   network, regardless of what it does

Nothing in this codebase calls out to the public internet except the
single configurable LLM endpoint — there's no telemetry, analytics, or
external dependencies fetched at runtime once built.

## Secure deployment

Set `AUTH_TOKEN` before starting the backend; it must contain at least 32 characters. Generate a fresh value with `openssl rand -hex 32` and put it in the untracked `server/.env`. The UI prompts for this token on the first protected request and retains it in memory only. This remains a single shared workspace, not a multi-tenant service.

The Node process binds to `127.0.0.1` by default. Compose binds the host port to loopback and sets the container's `HOST=0.0.0.0`. Put a TLS reverse proxy in front for remote access and configure explicit `CORS_ORIGIN` values. Do not publish the container port directly.

Document uploads accept at most four files of 5 MiB each, with two active upload/parse jobs and ten uploads per minute per IP. Files are spooled to temporary disk and deleted after processing. Each parser runs in a child process with a 128 MiB V8 heap, ten-second deadline, and 2 MiB output cap; total returned text is also capped at 2 MiB. Compose limits the container to 512 MiB. For non-Compose deployments, apply an equivalent OS/container memory limit: a V8 heap limit does not cap all native allocations. The frontend still supports locally parsed documents; this hardening applies to the server upload endpoint.

Validation: `cd server && npm ci && node --test tests/security.test.js`. Tests exercise a real local HTTP server and text-file parser, missing/wrong credentials, oversized files, and excess file count. Real PDF/DOCX fixtures and the output cap are tested with `python -m unittest discover -s server/tests -p test_documents.py -v` from the repository root.

Runtime: Node.js 22 or newer. Docker uses Node 22 and lockfile installs. The client manifest now declares its existing mammoth import; both package lockfiles were regenerated to match the manifests.
