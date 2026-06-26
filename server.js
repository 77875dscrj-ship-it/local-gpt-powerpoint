const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const { spawn } = require("child_process");

const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, "public");
const BRIDGE = path.join(ROOT, "scripts", "ppt-bridge.ps1");
const SOURCE_EXTRACTOR = path.join(ROOT, "scripts", "extract-source.ps1");
const CLIPBOARD_IMAGE_READER = path.join(ROOT, "scripts", "read-clipboard-image.ps1");
const PFX_PATH = path.join(ROOT, "certs", "localhost.pfx");
const PFX_PASSPHRASE = "local-gpt-powerpoint";
const CONFIG_DIR = path.join(ROOT, "config");
const POLICIES_PATH = path.join(CONFIG_DIR, "policies.json");
const SCHEMA_DIR = path.join(ROOT, "schemas");
const PRESENTATION_PLAN_SCHEMA = path.join(SCHEMA_DIR, "presentation-plan.schema.json");
const EXECUTION_PLAN_SCHEMA = path.join(SCHEMA_DIR, "execution-plan.schema.json");
const RUNTIME_DIR = path.join(ROOT, "runtime");
const REQUEST_DIR = path.join(RUNTIME_DIR, "requests");
const PLAN_DIR = path.join(RUNTIME_DIR, "plans");
const TRANSACTION_DIR = path.join(RUNTIME_DIR, "transactions");
const PREVIEW_DIR = path.join(RUNTIME_DIR, "previews");
const UPLOAD_DIR = path.join(RUNTIME_DIR, "uploads");
const COPIED_REAL_DIR = path.join(RUNTIME_DIR, "copied-real");
const LOG_DIR = path.join(RUNTIME_DIR, "logs");
const LOCALAPPDATA = process.env.LOCALAPPDATA || path.join(os.homedir(), "AppData", "Local");
const BACKUP_ROOT = path.join(LOCALAPPDATA, "LocalGptPowerPoint", "backups");
const PORT = Number(process.env.LOCAL_GPT_POWERPOINT_PORT || 8765);
const MAX_JSON_BODY_CHARS = 32 * 1024 * 1024;
const MAX_DECK_CONTEXT_CHARS = 52000;
const MAX_SOURCE_CONTEXT_CHARS = 36000;
const MAX_HISTORY_CHARS = 9000;
const MAX_CODEX_STDERR_CHARS = 8000;
const MAX_SHAPE_TEXT_CHARS = 700;
const MAX_PROMPT_SHAPES = 80;
const MAX_IMAGE_ATTACHMENTS = 8;
const MAX_IMAGE_BYTES = 12 * 1024 * 1024;

const LEGACY_ACTION_TYPES = {
  replace_deck: "slide.replace_deck",
  add_slides: "slide.add_many",
  set_title: "text.set_title",
  set_body: "text.set_body",
  set_notes: "notes.set",
  replace_text: "text.replace",
  format_selection: "selection.format",
  add_table_slide: "table.create_slide",
  add_bar_chart_slide: "chart.create_shape_bar_slide",
};

const DEFAULT_POLICIES = {
  schemaVersion: "1.0",
  defaultMode: "review",
  operationRisk: {
    set_title: "low",
    set_body: "low",
    set_notes: "low",
    replace_text: "low",
    add_slides: "medium",
    add_table_slide: "medium",
    add_bar_chart_slide: "medium",
    format_selection: "medium",
    replace_deck: "high",
  },
  preview: {
    requiredBeforeCommit: false,
    openMode: "hidden_first_visible_fallback",
  },
  undo: {
    startNewEntryBeforeApply: true,
  },
  backup: {
    beforeCommit: false,
    forRiskAtLeast: "high",
  },
  rollback: {
    primary: "backup_copy",
    inverseRollback: "safe_primitives_only",
    automaticOverwrite: false,
  },
  selectionEdit: {
    requiresFrozenTarget: true,
    requiresStableShapeTag: true,
    requiresFingerprint: true,
    requiresPreview: true,
    commitEnabled: true,
  },
  replaceDeck: {
    liveCommitEnabled: false,
    fallback: "create_rebuilt_copy",
  },
  liveDeck: {
    allowLiveBusinessDeckCommit: true,
    createEditableCopyForLiveDeck: false,
    commitAllowPathIncludes: [
      "\\runtime\\test-decks\\",
      "\\runtime\\copied-real\\",
      "\\runtime\\canary-copies\\",
    ],
  },
  backupRetention: {
    normalDays: 30,
    failedOrRecoveryDays: 90,
    maxBackupsPerDeck: 20,
    maxTotalBytes: 5368709120,
  },
};

function mergePolicy(base, override) {
  const out = clone(base);
  Object.keys(override || {}).forEach((key) => {
    const value = override[key];
    if (value && typeof value === "object" && !Array.isArray(value) && out[key] && typeof out[key] === "object" && !Array.isArray(out[key])) {
      out[key] = mergePolicy(out[key], value);
    } else {
      out[key] = value;
    }
  });
  return out;
}

function loadPolicies() {
  try {
    if (!fs.existsSync(POLICIES_PATH)) return clone(DEFAULT_POLICIES);
    const parsed = JSON.parse(fs.readFileSync(POLICIES_PATH, "utf8"));
    return mergePolicy(DEFAULT_POLICIES, parsed);
  } catch (err) {
    const fallback = clone(DEFAULT_POLICIES);
    fallback.loadWarning = err.message;
    return fallback;
  }
}

const POLICIES = loadPolicies();
const RISK_ORDER = { low: 1, medium: 2, high: 3 };

let commitLock = null;

function ensureDirs() {
  [CONFIG_DIR, RUNTIME_DIR, REQUEST_DIR, PLAN_DIR, TRANSACTION_DIR, PREVIEW_DIR, UPLOAD_DIR, COPIED_REAL_DIR, LOG_DIR, BACKUP_ROOT].forEach((dir) => {
    fs.mkdirSync(dir, { recursive: true });
  });
}

ensureDirs();

function allowedOrigin(req) {
  const origin = req.headers.origin || "";
  if (!origin) return "";
  if (/^https?:\/\/localhost(?::\d+)?$/i.test(origin)) return origin;
  if (/^https?:\/\/127\.0\.0\.1(?::\d+)?$/i.test(origin)) return origin;
  if (/^https?:\/\/\[::1\](?::\d+)?$/i.test(origin)) return origin;
  return "";
}

function writeJsonHeaders(req, res, status) {
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Access-Control-Allow-Headers": "Content-Type, X-Local-GPT-Session",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  };
  const origin = allowedOrigin(req);
  if (origin) headers["Access-Control-Allow-Origin"] = origin;
  res.writeHead(status, headers);
}

function json(res, status, value, req) {
  const body = JSON.stringify(value);
  writeJsonHeaders(req || { headers: {} }, res, status);
  res.end(body);
}

function readBody(req, cb) {
  let body = "";
  req.setEncoding("utf8");
  req.on("data", (chunk) => {
    body += chunk;
    if (body.length > MAX_JSON_BODY_CHARS) req.destroy();
  });
  req.on("end", () => {
    if (!body) return cb(null, {});
    try {
      cb(null, JSON.parse(body));
    } catch (err) {
      cb(err);
    }
  });
}

function limitText(text, maxChars, label) {
  const value = String(text || "").trim();
  if (value.length <= maxChars) return value;
  return value.slice(0, maxChars) + `\n\n[${label} truncated after ${maxChars} chars. Ask for a narrower scope if exact edits are needed.]`;
}

function limitInlineText(text, maxChars) {
  const value = String(text || "").replace(/\s+/g, " ").trim();
  if (value.length <= maxChars) return value;
  return `${value.slice(0, maxChars)}...`;
}

function makeId(prefix) {
  const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  const suffix = crypto.randomBytes(4).toString("hex");
  return `${prefix}-${stamp}-${suffix}`;
}

function sha256(value) {
  return "sha256:" + crypto.createHash("sha256").update(String(value || ""), "utf8").digest("hex");
}

function clone(value) {
  return JSON.parse(JSON.stringify(value || {}));
}

function safePathForPlan(planId) {
  if (!/^plan-[A-Za-z0-9_-]+$/.test(planId || "")) throw new Error("Invalid plan id.");
  return path.join(PLAN_DIR, `${planId}.json`);
}

function saveJson(filePath, value) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, JSON.stringify(value, null, 2), "utf8");
}

function saveJsonAtomic(filePath, value) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(tempPath, JSON.stringify(value, null, 2), "utf8");
  fs.renameSync(tempPath, filePath);
}

function loadPlanRecord(planId) {
  const filePath = safePathForPlan(planId);
  if (!fs.existsSync(filePath)) throw new Error("Plan not found.");
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function safePathForTransaction(transactionId) {
  if (!/^tx-[A-Za-z0-9_-]+$/.test(transactionId || "")) throw new Error("Invalid transaction id.");
  return path.join(TRANSACTION_DIR, `${transactionId}.json`);
}

function loadTransaction(transactionId) {
  const filePath = safePathForTransaction(transactionId);
  if (!fs.existsSync(filePath)) throw new Error("Transaction not found.");
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function saveTransaction(record) {
  record.updatedAt = new Date().toISOString();
  saveJsonAtomic(safePathForTransaction(record.transactionId), record);
}

function safeFileName(value) {
  return String(value || "deck").replace(/[<>:"/\\|?*\x00-\x1f]/g, "_").replace(/\s+/g, "_").slice(0, 80) || "deck";
}

function safeAssetPath(transactionId, name) {
  if (!/^tx-[A-Za-z0-9_-]+$/.test(transactionId || "")) throw new Error("Invalid transaction id.");
  if (!/^[A-Za-z0-9_.-]+$/.test(name || "")) throw new Error("Invalid asset name.");
  const dir = path.join(PREVIEW_DIR, transactionId);
  const assetPath = path.normalize(path.join(dir, name));
  if (!assetPath.startsWith(dir)) throw new Error("Invalid asset path.");
  return assetPath;
}

function safeUploadPath(uploadId, extension) {
  if (!/^img-[A-Za-z0-9_-]+$/.test(uploadId || "")) throw new Error("Invalid image upload id.");
  const ext = String(extension || ".png").toLowerCase();
  if (!/^\.(png|jpg|jpeg|webp|gif|bmp)$/.test(ext)) throw new Error("Unsupported image extension.");
  const filePath = path.normalize(path.join(UPLOAD_DIR, `${uploadId}${ext}`));
  if (!filePath.startsWith(UPLOAD_DIR)) throw new Error("Invalid image upload path.");
  return filePath;
}

function assetUrl(transactionId, filePath) {
  return `/api/transactions/${transactionId}/assets/${encodeURIComponent(path.basename(filePath))}`;
}

function uploadUrl(filePath) {
  return `/api/images/${encodeURIComponent(path.basename(filePath))}`;
}

function safeUploadedImagePath(assetName) {
  if (!/^[A-Za-z0-9_.-]+$/.test(assetName || "")) throw new Error("Invalid image asset name.");
  const imagePath = path.normalize(path.join(UPLOAD_DIR, assetName));
  if (!imagePath.startsWith(UPLOAD_DIR)) throw new Error("Invalid image asset path.");
  return imagePath;
}

function copyIfExists(src, dest) {
  if (!src || !fs.existsSync(src)) return false;
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.copyFileSync(src, dest);
  return true;
}

function imageExtensionForMime(mimeType, fileName) {
  const mime = String(mimeType || "").toLowerCase();
  if (mime === "image/png") return ".png";
  if (mime === "image/jpeg" || mime === "image/jpg") return ".jpg";
  if (mime === "image/webp") return ".webp";
  if (mime === "image/gif") return ".gif";
  if (mime === "image/bmp") return ".bmp";
  const ext = path.extname(String(fileName || "")).toLowerCase();
  if (/^\.(png|jpg|jpeg|webp|gif|bmp)$/.test(ext)) return ext;
  return ".png";
}

function normalizeImageAttachments(images) {
  if (!Array.isArray(images)) return [];
  const out = [];
  for (let i = 0; i < images.length && out.length < MAX_IMAGE_ATTACHMENTS; i++) {
    const item = images[i] && typeof images[i] === "object" ? images[i] : {};
    const imagePath = path.normalize(String(item.path || ""));
    if (!imagePath || !imagePath.startsWith(UPLOAD_DIR) || !fs.existsSync(imagePath)) continue;
    let stats = null;
    try { stats = fs.statSync(imagePath); } catch (_) {}
    if (!stats || !stats.isFile() || stats.size <= 0 || stats.size > MAX_IMAGE_BYTES) continue;
    out.push({
      name: safeFileName(item.name || path.basename(imagePath)),
      path: imagePath,
      mimeType: String(item.mimeType || "image/png"),
      sizeBytes: stats.size,
    });
  }
  return out;
}

function imageAttachmentsText(images) {
  const normalized = normalizeImageAttachments(images);
  if (!normalized.length) return "";
  return normalized.map((image, index) => [
    `[Image ${index + 1}: ${image.name}]`,
    `Path: ${image.path}`,
    `MIME: ${image.mimeType}`,
    `Size: ${image.sizeBytes} bytes`,
  ].join("\n")).join("\n\n");
}

function serveFile(req, res) {
  const urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  if (urlPath === "/icon-32.png") {
    const png = Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAAJElEQVR4nO3NMQEAAAgDILV/5zgBQsE0N2kAAAAAAAAAAADwGgYgAAHL3C1oAAAAAElFTkSuQmCC",
      "base64"
    );
    res.writeHead(200, { "Content-Type": "image/png", "Cache-Control": "public, max-age=3600" });
    res.end(png);
    return;
  }

  const rel = urlPath === "/" ? "taskpane.html" : urlPath.replace(/^\/+/, "");
  const filePath = path.normalize(path.join(PUBLIC_DIR, rel));
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    const type = ext === ".html" ? "text/html; charset=utf-8"
      : ext === ".css" ? "text/css; charset=utf-8"
      : ext === ".js" ? "application/javascript; charset=utf-8"
      : "application/octet-stream";
    res.writeHead(200, { "Content-Type": type, "Cache-Control": "no-store" });
    res.end(data);
  });
}

function runBridge(action, text, cb) {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "local-gpt-ppt-"));
  const inputPath = path.join(tempDir, "input.txt");
  if (typeof text === "string") fs.writeFileSync(inputPath, text, "utf8");
  const args = [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    BRIDGE,
    "-Action",
    action,
    "-InputPath",
    inputPath,
  ];
  const child = spawn("powershell.exe", args, { windowsHide: true });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => stdout += chunk);
  child.stderr.on("data", (chunk) => stderr += chunk);
  child.on("close", (code) => {
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch (_) {}
    let data = null;
    try { data = stdout.trim() ? JSON.parse(stdout.trim()) : null; } catch (_) {}
    if (code !== 0) {
      cb(new Error((data && data.error) || stderr || stdout || `PowerShell bridge exited ${code}`));
      return;
    }
    cb(null, data || { ok: true });
  });
}

function callCodexOAuth(instructions, prompt, schemaPath, requestId, imageAttachments, cb) {
  if (typeof imageAttachments === "function") {
    cb = imageAttachments;
    imageAttachments = [];
  }
  const images = normalizeImageAttachments(imageAttachments);
  const requestDir = path.join(REQUEST_DIR, requestId);
  const outPath = path.join(PLAN_DIR, `${requestId}-raw.json`);
  fs.mkdirSync(requestDir, { recursive: true });
  const fullPrompt = [
    instructions,
    "",
    "Important local runtime:",
    "- You are being invoked by Codex CLI with the user's existing ChatGPT OAuth session.",
    "- The caller intentionally does not pass a model flag. Use the current Codex configured model/options.",
    "- Return the requested final JSON object only.",
    "",
    prompt,
  ].join("\n");

  const args = [
    "exec",
    "--ephemeral",
    "--skip-git-repo-check",
    "--sandbox",
    "read-only",
    "--cd",
    requestDir,
  ];
  images.forEach((image) => {
    args.push("--image", image.path);
  });
  args.push(
    "--output-schema",
    schemaPath,
    "--output-last-message",
    outPath,
    "-"
  );
  const child = spawn("codex", args, {
    cwd: ROOT,
    windowsHide: true,
    env: Object.assign({}, process.env, { NO_COLOR: "1" }),
  });
  let stderr = "";
  child.stderr.setEncoding("utf8");
  child.stderr.on("data", (chunk) => {
    stderr += chunk;
    if (stderr.length > MAX_CODEX_STDERR_CHARS) stderr = stderr.slice(-MAX_CODEX_STDERR_CHARS);
  });
  child.on("error", (err) => cb(err));
  child.on("close", (code) => {
    let text = "";
    try {
      text = fs.existsSync(outPath) ? fs.readFileSync(outPath, "utf8").trim() : "";
    } catch (_) {}
    if (code !== 0) return cb(new Error(stderr.trim() || `Codex OAuth call failed. exit=${code}`));
    if (!text) return cb(new Error("Codex OAuth returned an empty response."));
    cb(null, text, outPath);
  });
  child.stdin.setDefaultEncoding("utf8");
  child.stdin.end(fullPrompt);
}

function extractJsonObject(text) {
  const trimmed = String(text || "").trim();
  if (!trimmed) throw new Error("Planner returned an empty response.");
  try { return JSON.parse(trimmed); } catch (_) {}
  const fence = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) {
    try { return JSON.parse(fence[1].trim()); } catch (_) {}
  }
  const start = trimmed.indexOf("{");
  const end = trimmed.lastIndexOf("}");
  if (start >= 0 && end > start) return JSON.parse(trimmed.slice(start, end + 1));
  throw new Error("Could not parse planner JSON.");
}

function summarizeHistory(history) {
  if (!Array.isArray(history)) return "";
  const lines = [];
  for (let i = Math.max(0, history.length - 10); i < history.length; i++) {
    const item = history[i] || {};
    const role = item.role === "user" ? "User" : "Assistant";
    const content = String(item.content || "").replace(/\s+/g, " ").trim();
    if (content) lines.push(`${role}: ${content.slice(0, 800)}`);
  }
  return limitText(lines.join("\n"), MAX_HISTORY_CHARS, "conversation history");
}

function isReadOnlyRequest(message) {
  const text = String(message || "").toLowerCase();
  const readOnlyHints = [
    "수정하지",
    "편집하지",
    "변경하지",
    "건드리지",
    "적용하지",
    "실행하지",
    "리뷰",
    "검토",
    "요약",
    "설명",
    "분석",
    "위험 요소",
    "문제점",
    "weak",
    "review",
    "summarize",
    "explain",
    "analyze"
  ];
  const editHints = [
    "고쳐",
    "수정해",
    "변경해",
    "추가해",
    "만들어",
    "생성해",
    "적용해",
    "바꿔",
    "줄여",
    "다듬어",
    "insert",
    "create",
    "change",
    "edit",
    "apply"
  ];
  let readOnlyScore = 0;
  let editScore = 0;
  for (let i = 0; i < readOnlyHints.length; i++) {
    if (text.indexOf(readOnlyHints[i]) >= 0) readOnlyScore += 1;
  }
  for (let j = 0; j < editHints.length; j++) {
    if (text.indexOf(editHints[j]) >= 0) editScore += 1;
  }
  if (text.indexOf("수정하지") >= 0 || text.indexOf("편집하지") >= 0 || text.indexOf("변경하지") >= 0) return true;
  return readOnlyScore > 0 && editScore === 0;
}

function normalizeShape(shape) {
  const source = shape && typeof shape === "object" ? shape : {};
  const text = String(source.text || source.textPreview || "");
  const normalized = {
    shapeIndex: Number(source.shapeIndex) || null,
    id: source.id !== undefined && source.id !== null ? Number(source.id) : null,
    name: String(source.name || ""),
    type: source.type !== undefined && source.type !== null ? Number(source.type) : null,
    autoShapeType: source.autoShapeType !== undefined && source.autoShapeType !== null ? Number(source.autoShapeType) : null,
    placeholderType: source.placeholderType !== undefined && source.placeholderType !== null ? Number(source.placeholderType) : null,
    isPlaceholder: !!source.isPlaceholder,
    zOrderPosition: source.zOrderPosition !== undefined && source.zOrderPosition !== null ? Number(source.zOrderPosition) : null,
    left: Number(source.left) || 0,
    top: Number(source.top) || 0,
    width: Number(source.width) || 0,
    height: Number(source.height) || 0,
    rotation: source.rotation !== undefined && source.rotation !== null ? Number(source.rotation) : 0,
    textPreview: limitInlineText(text, MAX_SHAPE_TEXT_CHARS),
    textLength: Number(source.textLength) || text.length,
    fontSize: Number(source.fontSize) || 0,
    hasTextFrame: !!source.hasTextFrame,
    hasTable: !!source.hasTable,
    hasChart: !!source.hasChart,
    fillRgb: source.fillRgb !== undefined && source.fillRgb !== null ? Number(source.fillRgb) : null,
    lineRgb: source.lineRgb !== undefined && source.lineRgb !== null ? Number(source.lineRgb) : null,
    altText: limitInlineText(source.altText || "", 240),
    tags: source.tags && typeof source.tags === "object" ? source.tags : {},
  };
  normalized.shapeFingerprint = sha256(JSON.stringify({
    id: normalized.id,
    name: normalized.name,
    type: normalized.type,
    placeholderType: normalized.placeholderType,
    left: normalized.left,
    top: normalized.top,
    width: normalized.width,
    height: normalized.height,
    textPreview: normalized.textPreview,
    textLength: normalized.textLength,
    zOrderPosition: normalized.zOrderPosition,
  }));
  return normalized;
}

function augmentShapeMap(map) {
  const out = map && typeof map === "object" ? clone(map) : {};
  const slides = Array.isArray(out.slides) ? out.slides : [];
  out.slides = slides.map((slide) => {
    const shapes = Array.isArray(slide.shapes) ? slide.shapes.map(normalizeShape) : [];
    return Object.assign({}, slide, {
      slideIndex: Number(slide.slideIndex) || null,
      slideId: slide.slideId !== undefined && slide.slideId !== null ? Number(slide.slideId) : null,
      shapeCount: Number(slide.shapeCount) || shapes.length,
      shapes,
      shapeMapFingerprint: sha256(JSON.stringify(shapes.map((shape) => shape.shapeFingerprint))),
    });
  });
  out.shapeMapFingerprint = sha256(JSON.stringify(out.slides.map((slide) => ({
    slideId: slide.slideId,
    shapeMapFingerprint: slide.shapeMapFingerprint,
  }))));
  return out;
}

function shapeMapForPrompt(shapeMap) {
  const slides = Array.isArray(shapeMap && shapeMap.slides) ? shapeMap.slides : [];
  return slides.map((slide) => ({
    slideIndex: slide.slideIndex,
    slideId: slide.slideId,
    title: slide.title || "",
    shapeCount: slide.shapeCount || 0,
    shapeMapFingerprint: slide.shapeMapFingerprint || "",
    shapes: (slide.shapes || []).slice(0, MAX_PROMPT_SHAPES).map((shape) => ({
      shapeIndex: shape.shapeIndex,
      id: shape.id,
      name: shape.name,
      type: shape.type,
      placeholderType: shape.placeholderType,
      isPlaceholder: shape.isPlaceholder,
      zOrderPosition: shape.zOrderPosition,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      textPreview: shape.textPreview,
      textLength: shape.textLength,
      fontSize: shape.fontSize,
      hasTable: shape.hasTable,
      hasChart: shape.hasChart,
      tags: shape.tags,
      shapeFingerprint: shape.shapeFingerprint,
    })),
  }));
}

function augmentSelection(selection) {
  const out = selection && typeof selection === "object" ? clone(selection) : {};
  if (Array.isArray(out.shapes)) {
    out.shapes = out.shapes.map(normalizeShape);
    out.shapeFingerprints = out.shapes.map((shape) => shape.shapeFingerprint);
    out.selectionFingerprint = sha256(JSON.stringify({
      slideId: out.slideId || null,
      slideIndex: out.slideIndex || null,
      shapes: out.shapeFingerprints,
    }));
  }
  return out;
}

function freezeSelectionTarget(context) {
  const selection = context && context.selection ? context.selection : {};
  const shapes = Array.isArray(selection.shapes) ? selection.shapes : [];
  if (!shapes.length) {
    throw policyViolation("format_selection requires one or more selected PowerPoint shapes. Select a text box or shape and try again.", 409);
  }
  const slideIndex = Number(selection.slideIndex || context.slideIndex || 1);
  const slideId = selection.slideId !== undefined && selection.slideId !== null ? Number(selection.slideId) : Number(context.slideId || 0);
  return {
    type: "shape_range",
    capturedAt: new Date().toISOString(),
    slideIndex,
    slideId: slideId || null,
    selectionFingerprint: selection.selectionFingerprint || null,
    shapes: shapes.map((shape) => ({
      shapeIndex: shape.shapeIndex,
      id: shape.id,
      name: shape.name,
      left: shape.left,
      top: shape.top,
      width: shape.width,
      height: shape.height,
      textPreview: shape.textPreview,
      textLength: shape.textLength,
      fontSize: shape.fontSize,
      shapeFingerprint: shape.shapeFingerprint,
    })),
  };
}

function actionHasFrozenSelection(action) {
  return !!(action && action.frozenSelection && Array.isArray(action.frozenSelection.shapes) && action.frozenSelection.shapes.length);
}

function formatSelectionHasMutableArgs(action) {
  const fields = ["fontSize", "width", "height", "left", "top", "autofit", "bold", "fillRgb", "lineSpacing", "spaceBefore", "spaceAfter"];
  return fields.some((field) => action && action[field] !== undefined && action[field] !== null);
}

function assertFrozenSelectionActions(executionPlan, statusCode) {
  (executionPlan.legacyActions || []).forEach((action) => {
    if (action.type === "format_selection" && !actionHasFrozenSelection(action)) {
      throw policyViolation("format_selection needs a frozen shape target. Select the shape again and create a new preview.", statusCode || 409);
    }
    if (action.type === "format_selection" && !formatSelectionHasMutableArgs(action)) {
      throw policyViolation("format_selection has no actual formatting fields to apply.", statusCode || 409);
    }
  });
}

function augmentContext(context) {
  const out = context && typeof context === "object" ? context : {};
  const slides = Array.isArray(out.slides) ? out.slides : [];
  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i] || {};
    const textHash = sha256(slide.text || "");
    slide.textHash = textHash;
    slide.slideFingerprint = sha256(JSON.stringify({
      slideId: slide.slideId || 0,
      title: slide.title || "",
      textHash,
      shapeCount: slide.shapeCount || 0,
    }));
    delete slide.text;
  }
  out.slides = slides;
  out.selection = augmentSelection(out.selection);
  if (out.activeSlideShapeMap) {
    out.activeSlideShapeMap = augmentShapeMap({
      ok: true,
      scope: "active",
      presentationName: out.presentationName || "",
      presentationFullName: out.presentationFullName || "",
      slideCount: out.slideCount || 0,
      slideWidth: out.slideWidth || 0,
      slideHeight: out.slideHeight || 0,
      slides: [out.activeSlideShapeMap],
    });
  }
  const fingerprintBasis = {
    presentationFullName: out.presentationFullName || "",
    slideCount: out.slideCount || 0,
    slideWidth: out.slideWidth || 0,
    slideHeight: out.slideHeight || 0,
    slides: slides.map((slide) => ({
      slideIndex: slide.slideIndex || 0,
      slideId: slide.slideId || 0,
      title: slide.title || "",
      textHash: slide.textHash || "",
      shapeCount: slide.shapeCount || 0,
    })),
  };
  out.deckFingerprint = sha256(JSON.stringify(fingerprintBasis));
  return out;
}

function getContextSnapshot(cb) {
  runBridge("context", "", (err, context) => {
    if (err) return cb(err);
    cb(null, augmentContext(context));
  });
}

function getSelectionSnapshot(cb) {
  runBridge("selection-context", "", (err, result) => {
    if (err) return cb(err);
    const selection = augmentSelection(result && result.selection);
    cb(null, {
      ok: true,
      presentationName: result && result.presentationName || "",
      presentationFullName: result && result.presentationFullName || "",
      slideCount: Number(result && result.slideCount) || 0,
      slideIndex: Number(result && result.slideIndex) || Number(selection.slideIndex) || 0,
      slideId: Number(result && result.slideId) || Number(selection.slideId) || 0,
      selection,
    });
  });
}

function getShapeMapSnapshot(payload, cb) {
  runBridge("shape-map", JSON.stringify(payload || {}), (err, shapeMap) => {
    if (err) return cb(err);
    cb(null, augmentShapeMap(shapeMap));
  });
}

function slideByIndex(context, index) {
  const slides = Array.isArray(context && context.slides) ? context.slides : [];
  for (let i = 0; i < slides.length; i++) {
    if (Number(slides[i].slideIndex) === Number(index)) return slides[i];
  }
  return null;
}

function slideById(context, slideId) {
  const slides = Array.isArray(context && context.slides) ? context.slides : [];
  for (let i = 0; i < slides.length; i++) {
    if (Number(slides[i].slideId) === Number(slideId)) return slides[i];
  }
  return null;
}

function defaultSlideRef(context, action) {
  let slideIndex = Number(action && action.slide);
  if ((!slideIndex || slideIndex < 1) && action && action.after !== undefined && action.after !== null) {
    slideIndex = Number(action.after);
  }
  if (!slideIndex || slideIndex < 1) slideIndex = Number(context && context.slideIndex) || 1;
  const slide = slideByIndex(context, slideIndex);
  return {
    slideIndex,
    slideId: slide ? slide.slideId : undefined,
  };
}

function affectedSlideSnapshots(context, executionPlan) {
  const out = [];
  const seen = {};
  const ids = Array.isArray(executionPlan && executionPlan.affectedSlideIds) ? executionPlan.affectedSlideIds : [];
  for (let i = 0; i < ids.length; i++) {
    const slide = slideById(context, ids[i]);
    if (slide && !seen[slide.slideId]) {
      out.push({
        slideId: slide.slideId,
        slideIndex: slide.slideIndex,
        slideFingerprint: slide.slideFingerprint,
        textHash: slide.textHash,
        title: slide.title || "",
      });
      seen[slide.slideId] = true;
    }
  }
  return out;
}

function assertAffectedSlidesUnchanged(before, after) {
  const affected = Array.isArray(before) ? before : [];
  for (let i = 0; i < affected.length; i++) {
    const current = slideById(after, affected[i].slideId);
    if (!current) {
      const err = new Error(`대상 슬라이드가 사라졌습니다: ${affected[i].slideId}`);
      err.statusCode = 409;
      throw err;
    }
    if (current.slideFingerprint !== affected[i].slideFingerprint) {
      const err = new Error(`계획 생성 이후 대상 슬라이드가 바뀌었습니다: ${affected[i].slideIndex || affected[i].slideId}`);
      err.statusCode = 409;
      throw err;
    }
  }
}

function safeText(value, max) {
  return String(value || "").replace(/\s+/g, " ").trim().slice(0, max || 400);
}

function synthesizeOutlineFromAction(action, index, context) {
  const changeId = action.changeId || `chg-${String(index + 1).padStart(3, "0")}`;
  const type = String(action.type || "");
  const slideRef = defaultSlideRef(context, action);
  const title = type === "replace_deck" ? "전체 덱 재구성"
    : type === "add_slides" ? "새 슬라이드 추가"
    : type === "format_selection" ? "선택 영역 서식 수정"
    : type === "add_table_slide" ? "표 슬라이드 추가"
    : type === "add_bar_chart_slide" ? "막대그래프 슬라이드 추가"
    : `슬라이드 ${slideRef.slideIndex} 수정`;
  return {
    changeId,
    operation: type === "replace_deck" || type === "add_slides" || type.indexOf("add_") === 0 ? "insert" : "revise",
    slideRef,
    title,
    keyMessage: safeText(action.title || action.text || action.message || "", 280),
    visualRole: type,
    changes: [{ type, summary: readableActionSummary(action, slideRef) }],
    rationale: "기존 action 호환 레이어에서 생성된 outline입니다.",
    sourceRefs: [],
    risk: riskForAction(action),
    selected: true,
  };
}

function readableActionSummary(action, slideRef) {
  const type = String(action && action.type || "");
  if (type === "replace_deck") return "프레젠테이션 전체를 새 슬라이드 구성으로 교체";
  if (type === "add_slides") return "현재 덱 뒤쪽에 편집 가능한 새 슬라이드 추가";
  if (type === "set_title") return `슬라이드 ${slideRef.slideIndex} 제목 변경`;
  if (type === "set_body") return `슬라이드 ${slideRef.slideIndex} 본문 변경`;
  if (type === "set_notes") return `슬라이드 ${slideRef.slideIndex} 발표자 노트 변경`;
  if (type === "replace_text") return `슬라이드 ${slideRef.slideIndex} 텍스트 찾아 바꾸기`;
  if (type === "format_selection") return "현재 선택된 텍스트/도형 서식 변경";
  if (type === "add_table_slide") return "편집 가능한 표 슬라이드 추가";
  if (type === "add_bar_chart_slide") return "도형 기반 막대그래프 슬라이드 추가";
  return "지원되는 PowerPoint 편집 작업";
}

function sanitizeLegacyAction(action, fallbackChangeId) {
  const value = clone(action);
  const type = String(value.type || "");
  if (!Object.prototype.hasOwnProperty.call(LEGACY_ACTION_TYPES, type)) {
    throw new Error(`Unsupported action type: ${type}`);
  }
  value.changeId = String(value.changeId || fallbackChangeId || makeId("chg")).replace(/[^A-Za-z0-9_-]/g, "-");
  if (!/^chg-/.test(value.changeId)) value.changeId = `chg-${value.changeId}`;
  if (value.slide !== undefined) {
    value.slide = Math.max(1, Math.floor(Number(value.slide) || 1));
  }
  if (value.after !== undefined) {
    value.after = Math.max(0, Math.floor(Number(value.after) || 0));
  }
  return value;
}

function normalizePresentationPlan(raw, context) {
  const input = raw && typeof raw === "object" ? clone(raw) : {};
  const legacyActions = Array.isArray(input.legacyActions) ? input.legacyActions
    : Array.isArray(input.actions) ? input.actions
      : [];
  const sanitizedActions = [];
  for (let i = 0; i < legacyActions.length; i++) {
    sanitizedActions.push(sanitizeLegacyAction(legacyActions[i], `chg-${String(i + 1).padStart(3, "0")}`));
  }

  const plan = {
    schemaVersion: input.schemaVersion || "2.0",
    kind: input.kind || (sanitizedActions.length ? "edit_plan" : "review"),
    assistantMessage: String(input.assistantMessage || input.summary || "작업 계획을 만들었습니다."),
    intent: input.intent && typeof input.intent === "object" ? input.intent : {},
    steps: Array.isArray(input.steps) ? input.steps : [],
    outline: Array.isArray(input.outline) ? input.outline : [],
    legacyActions: sanitizedActions,
    warnings: Array.isArray(input.warnings) ? input.warnings : [],
    requiresApproval: input.requiresApproval !== false && sanitizedActions.length > 0,
  };

  if (!plan.outline.length && sanitizedActions.length) {
    plan.outline = sanitizedActions.map((action, index) => synthesizeOutlineFromAction(action, index, context));
  }
  for (let i = 0; i < plan.outline.length; i++) {
    const item = plan.outline[i] || {};
    item.changeId = String(item.changeId || `chg-${String(i + 1).padStart(3, "0")}`);
    if (!/^chg-/.test(item.changeId)) item.changeId = `chg-${item.changeId}`;
    item.selected = item.selected !== false;
    if (!Array.isArray(item.changes) || !item.changes.length) {
      item.changes = [{ type: "update", summary: safeText(item.title || "슬라이드 변경", 300) }];
    }
    if (!item.risk) item.risk = "low";
    if (!item.operation) item.operation = "revise";
  }
  if (!plan.steps.length && plan.outline.length) {
    plan.steps = plan.outline.slice(0, 7).map((item) => ({
      title: item.title || "슬라이드 변경",
      detail: item.keyMessage || (item.changes && item.changes[0] && item.changes[0].summary) || "",
    }));
  }
  return plan;
}

function validatePresentationPlan(plan) {
  const errors = [];
  if (!plan || typeof plan !== "object") errors.push("Plan must be an object.");
  if (plan.schemaVersion !== "2.0") errors.push("schemaVersion must be 2.0.");
  if (["edit_plan", "review"].indexOf(plan.kind) < 0) errors.push("kind is invalid.");
  if (!String(plan.assistantMessage || "").trim()) errors.push("assistantMessage is required.");
  if (!Array.isArray(plan.outline)) errors.push("outline must be an array.");
  if (!Array.isArray(plan.legacyActions)) errors.push("legacyActions must be an array.");
  const seen = {};
  const risks = { low: true, medium: true, high: true };
  (plan.outline || []).forEach((item, index) => {
    if (!item.changeId || !/^chg-[A-Za-z0-9_-]+$/.test(item.changeId)) errors.push(`outline[${index}].changeId is invalid.`);
    if (seen[item.changeId]) errors.push(`Duplicate changeId: ${item.changeId}`);
    seen[item.changeId] = true;
    if (!risks[item.risk]) errors.push(`outline[${index}].risk is invalid.`);
  });
  (plan.legacyActions || []).forEach((action, index) => {
    if (!Object.prototype.hasOwnProperty.call(LEGACY_ACTION_TYPES, action.type)) errors.push(`legacyActions[${index}].type is unsupported.`);
  });
  const raw = JSON.stringify(plan);
  ["Invoke-Expression", "powershell.exe", "cmd.exe", "Start-Process", "WScript.Shell", "Add-Type"].forEach((term) => {
    if (raw.toLowerCase().indexOf(term.toLowerCase()) >= 0) errors.push(`Forbidden command-like term found: ${term}`);
  });
  if (errors.length) {
    const err = new Error(errors.join(" "));
    err.validationErrors = errors;
    throw err;
  }
}

function operationTypeForAction(action) {
  return LEGACY_ACTION_TYPES[action.type] || "unsupported";
}

function riskForAction(action) {
  const key = String(action && action.type || "");
  const risk = POLICIES.operationRisk && POLICIES.operationRisk[key];
  return RISK_ORDER[risk] ? risk : "high";
}

function maxRiskLevel(actions) {
  let max = "low";
  (actions || []).forEach((action) => {
    const risk = action && action.risk ? action.risk : riskForAction(action);
    if ((RISK_ORDER[risk] || 0) > (RISK_ORDER[max] || 0)) max = risk;
  });
  return max;
}

function executionHasAction(executionPlan, actionType) {
  return (executionPlan.legacyActions || []).some((action) => action.type === actionType);
}

function policyViolation(message, statusCode) {
  const err = new Error(message);
  err.statusCode = statusCode || 409;
  return err;
}

function assertPreviewPolicy(executionPlan) {
  if (executionHasAction(executionPlan, "replace_deck") && !(POLICIES.replaceDeck && POLICIES.replaceDeck.liveCommitEnabled)) {
    throw policyViolation("replace_deck is blocked in this build. Policy requires a rebuilt-copy workflow before any whole-deck replacement can be committed.", 409);
  }
  if (executionHasAction(executionPlan, "format_selection")) {
    if (POLICIES.selectionEdit && POLICIES.selectionEdit.commitEnabled === false) {
      throw policyViolation("format_selection is blocked until frozen selection tags, shape fingerprint checks, and preview-safe commit are implemented.", 409);
    }
    assertFrozenSelectionActions(executionPlan, 409);
  }
}

function normalizeFsPath(value) {
  return String(value || "").replace(/\//g, "\\").toLowerCase();
}

function isCommitAllowlistedPath(fullName) {
  const normalized = normalizeFsPath(fullName);
  const markers = (POLICIES.liveDeck && POLICIES.liveDeck.commitAllowPathIncludes) || [];
  return markers.some((marker) => normalized.indexOf(normalizeFsPath(marker)) >= 0);
}

function shouldCommitToEditableCopy(currentContext) {
  if (POLICIES.liveDeck && POLICIES.liveDeck.createEditableCopyForLiveDeck === false) return false;
  if (!(POLICIES.liveDeck && POLICIES.liveDeck.allowLiveBusinessDeckCommit === false)) return false;
  return !isCommitAllowlistedPath(currentContext && currentContext.presentationFullName);
}

function riskAtLeast(actual, threshold) {
  const actualRank = RISK_ORDER[String(actual || "low")] || 1;
  const thresholdRank = RISK_ORDER[String(threshold || "high")] || 3;
  return actualRank >= thresholdRank;
}

function shouldCreateCommitBackup(transaction) {
  const backupPolicy = POLICIES.backup || {};
  if (backupPolicy.beforeCommit === false) {
    return riskAtLeast(transaction && transaction.maxRisk, backupPolicy.forRiskAtLeast || "high");
  }
  return true;
}

function assertCommitPolicy(transaction, currentContext, options) {
  const opts = options || {};
  if (POLICIES.preview && POLICIES.preview.requiredBeforeCommit && transaction.status !== "previewed") {
    throw policyViolation(`Transaction is not previewed: ${transaction.status}`, 409);
  }
  if (executionHasAction(transaction.executionPlan, "replace_deck") && !(POLICIES.replaceDeck && POLICIES.replaceDeck.liveCommitEnabled)) {
    throw policyViolation("replace_deck live commit is disabled by policy. Create a rebuilt copy instead of overwriting the active presentation.", 403);
  }
  if (executionHasAction(transaction.executionPlan, "format_selection")) {
    if (POLICIES.selectionEdit && POLICIES.selectionEdit.commitEnabled === false) {
      throw policyViolation("format_selection commit is disabled by policy until selection freeze/tag/fingerprint support is complete.", 403);
    }
    assertFrozenSelectionActions(transaction.executionPlan, 403);
  }
  if (POLICIES.liveDeck && POLICIES.liveDeck.allowLiveBusinessDeckCommit === false) {
    const fullName = currentContext.presentationFullName || transaction.deckIdentity.fullName || "";
    if (!isCommitAllowlistedPath(fullName) && !opts.allowEditableCopyFallback) {
      throw policyViolation("Live business deck commit is disabled by policy. Open a disposable or copied test deck under runtime\\test-decks before applying changes.", 403);
    }
  }
}

function verifyBackupFile(backupPath) {
  let stats = null;
  try {
    stats = fs.statSync(backupPath);
  } catch (err) {
    throw new Error(`Backup file was not created: ${backupPath}`);
  }
  if (!stats.isFile() || stats.size <= 0) {
    throw new Error(`Backup file is empty or invalid: ${backupPath}`);
  }
  return stats;
}

function compileExecutionPlan(plan, context, planId, selectedChangeIds) {
  const selected = {};
  let hasExplicitSelection = false;
  if (Array.isArray(selectedChangeIds)) {
    for (let i = 0; i < selectedChangeIds.length; i++) {
      selected[String(selectedChangeIds[i])] = true;
      hasExplicitSelection = true;
    }
  }

  if (hasExplicitSelection) {
    const known = {};
    for (let i = 0; i < plan.legacyActions.length; i++) known[plan.legacyActions[i].changeId] = true;
    Object.keys(selected).forEach((changeId) => {
      if (!known[changeId]) throw new Error(`Unknown changeId selected: ${changeId}`);
    });
    const replaceDeck = plan.legacyActions.filter((action) => action.type === "replace_deck");
    if (replaceDeck.length && Object.keys(selected).length !== plan.legacyActions.length) {
      throw new Error("replace_deck 작업은 atomic operation이라 일부 변경만 선택할 수 없습니다.");
    }
  }

  const legacyActions = [];
  const operations = [];
  const affected = {};
  for (let i = 0; i < plan.legacyActions.length; i++) {
    const action = clone(plan.legacyActions[i]);
    const changeId = action.changeId || (plan.outline[i] && plan.outline[i].changeId) || `chg-${String(i + 1).padStart(3, "0")}`;
    action.changeId = changeId;
    if (hasExplicitSelection && !selected[changeId]) continue;
    if (action.type === "format_selection") {
      delete action.text;
      action.frozenSelection = freezeSelectionTarget(context);
      action.slide = action.frozenSelection.slideIndex;
    }
    const slideRef = defaultSlideRef(context, action);
    if (slideRef.slideId) affected[String(slideRef.slideId)] = slideRef.slideId;
    const opId = `op-${String(operations.length + 1).padStart(3, "0")}`;
    const risk = riskForAction(action);
    operations.push({
      opId,
      changeId,
      type: operationTypeForAction(action),
      risk,
      target: {
        slideIndex: slideRef.slideIndex,
        slideId: slideRef.slideId || null,
        mode: action.type === "format_selection" ? "frozen_shape_range" : "slide",
        frozenShapeCount: action.frozenSelection && Array.isArray(action.frozenSelection.shapes) ? action.frozenSelection.shapes.length : 0,
      },
      expected: {
        deckFingerprint: context.deckFingerprint,
        selectionFingerprint: action.frozenSelection ? action.frozenSelection.selectionFingerprint || null : null,
        shapeFingerprints: action.frozenSelection ? action.frozenSelection.shapes.map((shape) => shape.shapeFingerprint) : [],
      },
      args: clone(action),
      idempotencyKey: `${planId}-${opId}`,
    });
    legacyActions.push(action);
  }

  return {
    schemaVersion: "2.0",
    planId,
    transactionId: makeId("tx"),
    deckFingerprint: context.deckFingerprint,
    affectedSlideIds: Object.keys(affected).map((key) => affected[key]),
    maxRisk: maxRiskLevel(legacyActions),
    operations,
    legacyActions,
    validation: [
      { type: "schema_validated" },
      { type: "deck_fingerprint_precondition" },
      { type: "legacy_action_allowlist" },
      { type: "policy_risk_classified", maxRisk: maxRiskLevel(legacyActions) },
      { type: "frozen_selection_required_for_format_selection" },
    ],
  };
}

function publicPlan(plan, record) {
  const out = clone(plan);
  out.planId = record.planId;
  out.deckFingerprint = record.deckFingerprint;
  out.generatedAt = record.generatedAt;
  out.requestedMode = record.requestedMode || "review";
  out.executionPlan = {
    transactionId: record.executionPlan.transactionId,
    operationCount: record.executionPlan.operations.length,
    affectedSlideIds: record.executionPlan.affectedSlideIds,
    maxRisk: record.executionPlan.maxRisk || "low",
  };
  out.actions = record.executionPlan.legacyActions;
  out.needsPermission = out.actions.length > 0;
  return out;
}

function transactionPublic(record) {
  const out = clone(record);
  if (out.preview && Array.isArray(out.preview.slides)) {
    out.preview.slides = out.preview.slides.map((slide) => Object.assign({}, slide, {
      beforeImageUrl: slide.beforeImage ? assetUrl(record.transactionId, slide.beforeImage) : null,
      afterImageUrl: slide.afterImage ? assetUrl(record.transactionId, slide.afterImage) : null,
    }));
  }
  return out;
}

function newTransaction(planRecord, executionPlan, selectedChangeIds, status) {
  const transactionId = executionPlan.transactionId;
  const now = new Date().toISOString();
  return {
    schemaVersion: "1.0",
    transactionId,
    planId: planRecord.planId,
    status,
    createdAt: now,
    updatedAt: now,
    deckIdentity: {
      fullName: planRecord.context.presentationFullName || "",
      presentationName: planRecord.context.presentationName || "",
      deckFingerprintAtPlan: planRecord.deckFingerprint,
      deckFingerprintAtCommit: null,
    },
    selectedChangeIds: selectedChangeIds || null,
    affectedSlideIds: executionPlan.affectedSlideIds,
    affectedSlidesAtPlan: affectedSlideSnapshots(planRecord.context, executionPlan),
    maxRisk: executionPlan.maxRisk || "low",
    policySnapshot: {
      defaultMode: POLICIES.defaultMode || "review",
      previewRequiredBeforeCommit: !!(POLICIES.preview && POLICIES.preview.requiredBeforeCommit),
      liveBusinessDeckCommit: !!(POLICIES.liveDeck && POLICIES.liveDeck.allowLiveBusinessDeckCommit),
      replaceDeckLiveCommit: !!(POLICIES.replaceDeck && POLICIES.replaceDeck.liveCommitEnabled),
      selectionEditCommit: !!(POLICIES.selectionEdit && POLICIES.selectionEdit.commitEnabled),
    },
    backup: {
      created: false,
      path: null,
      format: null,
      createdAt: null,
      verified: false,
      sizeBytes: 0,
    },
    preview: {
      created: false,
      shadowDeckPath: null,
      slides: [],
      liveDeckModified: null,
    },
    executionPlan,
    journal: [],
    result: null,
    error: null,
  };
}

function planBridgePayload(planRecord, executionPlan) {
  return {
    assistantMessage: planRecord.presentationPlan.assistantMessage,
    actions: executionPlan.legacyActions,
  };
}

function previewPlan(planId, selectedChangeIds, cb) {
  let planRecord;
  try {
    planRecord = loadPlanRecord(planId);
  } catch (err) {
    return cb(err);
  }
  if ((planRecord.requestedMode || "review") === "review") {
    const err = new Error("review 모드 계획은 preview/commit할 수 없습니다.");
    err.statusCode = 409;
    return cb(err);
  }

  getContextSnapshot((contextErr, currentContext) => {
    if (contextErr) return cb(contextErr);
    let executionPlan;
    let transaction;
    try {
      executionPlan = compileExecutionPlan(planRecord.presentationPlan, currentContext, planRecord.planId, selectedChangeIds);
      if (!executionPlan.legacyActions.length) {
        const err = new Error("선택된 편집 operation이 없습니다.");
        err.statusCode = 409;
        throw err;
      }
      assertPreviewPolicy(executionPlan);
      assertAffectedSlidesUnchanged(affectedSlideSnapshots(planRecord.context, executionPlan), currentContext);
      transaction = newTransaction(planRecord, executionPlan, selectedChangeIds || null, "previewing");
      saveTransaction(transaction);
    } catch (err) {
      return cb(err);
    }

    const previewDir = path.join(PREVIEW_DIR, transaction.transactionId);
    fs.mkdirSync(previewDir, { recursive: true });
    const ext = path.extname(currentContext.presentationFullName || "") || ".pptx";
    const shadowPath = path.join(previewDir, `shadow${ext}`);
    const bridgePayload = {
      transactionId: transaction.transactionId,
      shadowPath,
      assetDir: previewDir,
      affectedSlideIds: executionPlan.affectedSlideIds,
      actions: executionPlan.legacyActions,
      assistantMessage: planRecord.presentationPlan.assistantMessage,
    };

    runBridge("preview-json", JSON.stringify(bridgePayload), (bridgeErr, previewResult) => {
      if (bridgeErr) {
        transaction.status = "failed";
        transaction.error = bridgeErr.message;
        saveTransaction(transaction);
        return cb(bridgeErr);
      }
      getContextSnapshot((afterErr, afterContext) => {
        if (afterErr) {
          transaction.status = "failed";
          transaction.error = afterErr.message;
          saveTransaction(transaction);
          return cb(afterErr);
        }
        let liveDeckModified = false;
        try {
          assertAffectedSlidesUnchanged(transaction.affectedSlidesAtPlan, afterContext);
        } catch (_) {
          liveDeckModified = true;
        }
        transaction.status = liveDeckModified ? "failed" : "previewed";
        transaction.preview = {
          created: !liveDeckModified,
          shadowDeckPath: previewResult.shadowDeckPath || shadowPath,
          openMode: previewResult.openMode || null,
          slides: (previewResult.slides || []).map((slide) => ({
            slideId: slide.slideId || null,
            slideIndex: slide.slideIndex || null,
            beforeImage: slide.beforeImage || null,
            afterImage: slide.afterImage || null,
            textDiff: slide.textDiff || [],
          })),
          liveDeckModified,
        };
        transaction.error = liveDeckModified ? "Preview unexpectedly changed the live deck target fingerprint." : null;
        saveTransaction(transaction);
        if (liveDeckModified) {
          const err = new Error("미리보기 중 live deck 변경이 감지되어 중단했습니다.");
          err.statusCode = 500;
          return cb(err);
        }
        cb(null, {
          ok: true,
          transactionId: transaction.transactionId,
          status: transaction.status,
          liveDeckModified: false,
          slides: transactionPublic(transaction).preview.slides,
        });
      });
    });
  });
}

function backupPathForTransaction(transaction, currentContext) {
  const deckName = safeFileName(currentContext.presentationName || transaction.deckIdentity.presentationName || "unsaved-deck");
  const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  const dir = path.join(BACKUP_ROOT, deckName, `${stamp}-${transaction.transactionId}`);
  const ext = path.extname(currentContext.presentationFullName || "") || ".pptx";
  return path.join(dir, `original-copy${ext}`);
}

function editableCopyPathForTransaction(transaction, currentContext) {
  const deckName = safeFileName(currentContext.presentationName || transaction.deckIdentity.presentationName || "deck");
  const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  const ext = path.extname(currentContext.presentationFullName || "") || ".pptx";
  const dir = path.join(COPIED_REAL_DIR, deckName, `${stamp}-${transaction.transactionId}`);
  return path.join(dir, `editable-copy${ext}`);
}

function prepareCommitTarget(transaction, currentContext, cb) {
  if (!shouldCommitToEditableCopy(currentContext)) {
    return cb(null, {
      mode: "active_presentation",
      context: currentContext,
      sourceFullName: currentContext.presentationFullName || "",
      editableCopyPath: null,
      editableCopySizeBytes: 0,
      openResult: null,
    });
  }

  const editableCopyPath = editableCopyPathForTransaction(transaction, currentContext);
  runBridge("save-copy", JSON.stringify({ backupPath: editableCopyPath }), (copyErr, copyResult) => {
    if (copyErr) return cb(copyErr);
    const actualCopyPath = copyResult.path || editableCopyPath;
    let copyStats = null;
    try {
      copyStats = verifyBackupFile(actualCopyPath);
    } catch (verifyErr) {
      return cb(verifyErr);
    }
    runBridge("open-presentation", JSON.stringify({ path: actualCopyPath }), (openErr, openResult) => {
      if (openErr) return cb(openErr);
      getContextSnapshot((copyContextErr, copyContext) => {
        if (copyContextErr) return cb(copyContextErr);
        cb(null, {
          mode: "editable_copy",
          context: copyContext,
          sourceFullName: currentContext.presentationFullName || "",
          editableCopyPath: actualCopyPath,
          editableCopySizeBytes: copyStats.size,
          openResult,
        });
      });
    });
  });
}

function commitTransaction(transactionId, approval, cb) {
  let transaction;
  try {
    transaction = loadTransaction(transactionId);
  } catch (err) {
    return cb(err);
  }
  if (transaction.status === "committed") {
    return cb(null, { ok: true, applied: true, idempotent: true, transaction: transactionPublic(transaction), result: transaction.result });
  }
  const retryableEditableCopyFallback = transaction.status === "rejected"
    && /Live business deck commit is disabled by policy/i.test(String(transaction.error || ""));
  if (retryableEditableCopyFallback) {
    transaction.status = "previewed";
    transaction.error = null;
    saveTransaction(transaction);
  }
  if (transaction.status !== "previewed") {
    const err = new Error(`Transaction is not previewed: ${transaction.status}`);
    err.statusCode = 409;
    return cb(err);
  }
  if (!approval || approval.approved !== true) {
    const err = new Error("사용자 승인 flag가 없어 commit을 중단했습니다.");
    err.statusCode = 403;
    return cb(err);
  }
  if (commitLock && commitLock !== transactionId) {
    const err = new Error("다른 PowerPoint 편집 transaction이 진행 중입니다.");
    err.statusCode = 409;
    return cb(err);
  }
  commitLock = transactionId;

  getContextSnapshot((contextErr, currentContext) => {
    if (contextErr) {
      commitLock = null;
      return cb(contextErr);
    }
    try {
      assertCommitPolicy(transaction, currentContext, { allowEditableCopyFallback: true });
      assertAffectedSlidesUnchanged(transaction.affectedSlidesAtPlan, currentContext);
    } catch (err) {
      transaction.status = err.statusCode === 403 ? "rejected" : "conflicted";
      transaction.error = err.message;
      saveTransaction(transaction);
      commitLock = null;
      return cb(err);
    }

    prepareCommitTarget(transaction, currentContext, (targetErr, target) => {
      if (targetErr) {
        transaction.status = "failed";
        transaction.error = `Editable copy preparation failed: ${targetErr.message}`;
        saveTransaction(transaction);
        commitLock = null;
        return cb(targetErr);
      }
      const targetContext = target.context;
      try {
        assertCommitPolicy(transaction, targetContext);
        assertAffectedSlidesUnchanged(transaction.affectedSlidesAtPlan, targetContext);
      } catch (err) {
        transaction.status = err.statusCode === 403 ? "rejected" : "conflicted";
        transaction.error = err.message;
        saveTransaction(transaction);
        commitLock = null;
        return cb(err);
      }

      transaction.commitTarget = {
        mode: target.mode,
        sourceFullName: target.sourceFullName,
        editableCopyPath: target.editableCopyPath,
        editableCopySizeBytes: target.editableCopySizeBytes,
        openedAt: target.mode === "editable_copy" ? new Date().toISOString() : null,
        openResult: target.openResult,
      };
      saveTransaction(transaction);

      const applyAfterBackupDecision = () => {
        transaction.status = "applying";
        if (!transaction.backup || transaction.backup.created !== true) {
          transaction.backup = {
            created: false,
            path: null,
            format: null,
            createdAt: null,
            verified: false,
            sizeBytes: 0,
            skippedReason: "disabled_by_policy_use_powerpoint_undo",
          };
        }
        saveTransaction(transaction);

        const bridgePlan = {
          assistantMessage: "",
          actions: transaction.executionPlan.legacyActions,
        };
        runBridge("apply-json", JSON.stringify(bridgePlan), (applyErr, result) => {
        if (applyErr) {
          transaction.status = "recovery_required";
          transaction.error = applyErr.message;
          saveTransaction(transaction);
          commitLock = null;
          return cb(applyErr);
        }
        const finishAfterApply = () => getContextSnapshot((afterErr, afterContext) => {
          transaction.deckIdentity.deckFingerprintAtCommit = afterContext && afterContext.deckFingerprint ? afterContext.deckFingerprint : null;
          if (transaction.commitTarget) {
            transaction.commitTarget.committedFullName = afterContext && afterContext.presentationFullName ? afterContext.presentationFullName : null;
            transaction.commitTarget.committedPresentationName = afterContext && afterContext.presentationName ? afterContext.presentationName : null;
          }
          transaction.status = afterErr ? "recovery_required" : "committed";
          transaction.result = result;
          transaction.journal.push({
            type: "bridge_apply",
            status: afterErr ? "postflight_failed" : "applied",
            startedAt: transaction.backup.createdAt || transaction.updatedAt,
            finishedAt: new Date().toISOString(),
            operationCount: transaction.executionPlan.operations.length,
            result,
            backupPath: transaction.backup.path || null,
          });
          if (afterErr) transaction.error = afterErr.message;
          saveTransaction(transaction);
          commitLock = null;
          if (afterErr) return cb(afterErr);
          cb(null, {
            ok: true,
            applied: true,
            transactionId: transaction.transactionId,
            transaction: transactionPublic(transaction),
            result,
          });
        });
        if (transaction.commitTarget && transaction.commitTarget.mode === "editable_copy") {
          runBridge("save-active", "", (saveErr, saveResult) => {
            if (transaction.commitTarget) transaction.commitTarget.saveResult = saveResult || null;
            if (saveErr) {
              transaction.status = "recovery_required";
              transaction.error = `Editable copy save failed: ${saveErr.message}`;
              saveTransaction(transaction);
              commitLock = null;
              return cb(saveErr);
            }
            finishAfterApply();
          });
        } else {
          finishAfterApply();
        }
      });
      };

      if (!shouldCreateCommitBackup(transaction)) {
        applyAfterBackupDecision();
        return;
      }

      const backupPath = backupPathForTransaction(transaction, targetContext);
      runBridge("save-copy", JSON.stringify({ backupPath }), (backupErr, backupResult) => {
        if (backupErr) {
          transaction.status = "failed";
          transaction.error = `Backup failed: ${backupErr.message}`;
          saveTransaction(transaction);
          commitLock = null;
          return cb(backupErr);
        }
        const actualBackupPath = backupResult.path || backupPath;
        let backupStats = null;
        try {
          backupStats = verifyBackupFile(actualBackupPath);
        } catch (verifyErr) {
          transaction.status = "failed";
          transaction.error = `Backup verification failed: ${verifyErr.message}`;
          saveTransaction(transaction);
          commitLock = null;
          return cb(verifyErr);
        }
        transaction.backup = {
          created: true,
          path: actualBackupPath,
          format: path.extname(backupPath).replace(".", "").toLowerCase(),
          createdAt: new Date().toISOString(),
          verified: true,
          sizeBytes: backupStats.size,
        };
        applyAfterBackupDecision();
      });
    });
  });
}

function rollbackTransaction(transactionId, cb) {
  let transaction;
  try {
    transaction = loadTransaction(transactionId);
  } catch (err) {
    return cb(err);
  }
  if (!transaction.backup || !transaction.backup.path) {
    const err = new Error("사용 가능한 backup이 없습니다.");
    err.statusCode = 409;
    return cb(err);
  }
  runBridge("open-presentation", JSON.stringify({ path: transaction.backup.path }), (err, result) => {
    if (err) {
      transaction.status = "recovery_required";
      transaction.error = err.message;
      saveTransaction(transaction);
      return cb(err);
    }
    transaction.status = "recovery_required";
    transaction.journal.push({
      type: "backup_opened",
      status: "opened",
      finishedAt: new Date().toISOString(),
      path: transaction.backup.path,
    });
    saveTransaction(transaction);
    cb(null, {
      ok: true,
      status: transaction.status,
      message: "백업본을 새 프레젠테이션으로 열었습니다. 원본 파일은 자동 덮어쓰지 않았습니다.",
      backupPath: transaction.backup.path,
      result,
    });
  });
}

function buildPlannerPrompt(message, sourceText, history, context, readOnlyIntent, requestedMode, imageAttachments) {
  const slidesForPrompt = (context.slides || []).map((slide) => ({
    slideIndex: slide.slideIndex,
    slideId: slide.slideId,
    title: slide.title,
    textHash: slide.textHash,
    shapeCount: slide.shapeCount,
  }));
  const deckText = limitText(context.deckText || "", MAX_DECK_CONTEXT_CHARS, "deck text");
  const limitedSource = limitText(sourceText, MAX_SOURCE_CONTEXT_CHARS, "source material");
  const selection = context.selection || {};
  const activeSlideShapeMap = shapeMapForPrompt(context.activeSlideShapeMap);
  const imageContext = imageAttachmentsText(imageAttachments);

  const instructions = [
    "You are the planning brain for a PowerPoint add-in.",
    "Return a semantic PresentationPlan JSON object that conforms to the provided schema.",
    "Respond in Korean.",
    "Do not generate PowerShell code, COM method names, shell commands, rollback commands, or raw automation code.",
    "Use legacyActions only as high-level edit recipes from the allowed action type list.",
    "If the user only asks to review, summarize, or explain, return kind=review, outline=[], legacyActions=[], requiresApproval=false.",
  ].join("\n");

  const prompt = [
    "Create a Local GPT PresentationPlan.",
    "",
    "Allowed legacyActions recipe types:",
    "- replace_deck: {type, changeId, slides:[{kind,title,subtitle,bullets,notes}]}",
    "- add_slides: {type, changeId, after, slides:[{kind,title,subtitle,bullets,notes}]}",
    "- set_title: {type, changeId, slide, text}",
    "- set_body: {type, changeId, slide, text}",
    "- set_notes: {type, changeId, slide, text}",
    "- replace_text: {type, changeId, slide, find, replace}",
    "- format_selection: {type, changeId, fontSize, width, height, left, top, autofit, bold, fillRgb, lineSpacing, spaceBefore, spaceAfter}. Do not include text; selection.format never replaces text.",
    "- add_table_slide: {type, changeId, after, title, columns, rows, notes}",
    "- add_bar_chart_slide: {type, changeId, after, title, message, items:[{label,value}], notes}",
    "",
    "Planning rules:",
    `- Requested mode: ${requestedMode}.`,
    `- Default UI mode is ${POLICIES.defaultMode || "review"}. In review mode, legacyActions must stay empty.`,
    "- Whole-deck replacement must be treated as a rebuilt-copy workflow, not a direct live-deck overwrite.",
    "- Selection formatting is allowed only for the currently selected shape range; the server freezes the target shape IDs and fingerprints before preview.",
    "- If the user asks to improve spacing, readability, size, position, color, or layout of a selected block, use format_selection without any text replacement. For line spacing, set lineSpacing such as 1.15 or 1.25. For paragraph gaps, set spaceBefore/spaceAfter such as 0.1 to 0.3.",
    "- If the user asks to rewrite selected text, use replace_text or another explicit text action; do not put rewritten prose into format_selection.text.",
    `- Server read-only classification: ${readOnlyIntent ? "true" : "false"}.`,
    "- If Server read-only classification is true, legacyActions must be [] and requiresApproval must be false. Put the requested analysis directly in assistantMessage.",
    "- The schema is strict. Include all fields shown by the schema. Use null for unused scalar legacyAction fields and [] for unused arrays.",
    "- For outline.slideRef, use null for fields that do not apply.",
    "- Every actionable legacyAction must have a matching outline item with the same changeId.",
    "- Use slideId in outline.slideRef when editing an existing slide.",
    "- Use slideIndex only as a display fallback.",
    "- Use activeSlideShapeMap to understand text boxes, tables, charts, and layout density on the current slide.",
    "- Do not invent shape IDs. If a shape-level edit is needed but no safe action exists, explain the needed shape target in the outline and wait for preview-safe tooling.",
    "- For new slides, use a tempSlideKey in outline.slideRef.",
    "- Keep numeric claims from source or current deck. Do not invent numbers.",
    "- If pasted images are attached, inspect them as visual reference material. Use them to infer requested layout, screenshot context, or content the user wants reflected in the deck.",
    "- Keep slide content concise and editable.",
    "- Prefer selected/current slide only when the user's request says current slide or selection.",
    "- Never include raw code or command strings.",
    "",
    `User message:\n${message}`,
    "",
    history ? `Recent chat:\n${history}` : "Recent chat: (none)",
    "",
    limitedSource ? `Attached or pasted source material:\n${limitedSource}` : "Attached or pasted source material: (none)",
    "",
    imageContext ? `Attached pasted image files:\n${imageContext}` : "Attached pasted image files: (none)",
    "",
    `PowerPoint snapshot:\n${JSON.stringify({
      presentationName: context.presentationName,
      slideIndex: context.slideIndex,
      slideId: context.slideId,
      slideCount: context.slideCount,
      slideWidth: context.slideWidth,
      slideHeight: context.slideHeight,
      deckFingerprint: context.deckFingerprint,
      slides: slidesForPrompt,
      activeSlideShapeMap,
      selection,
    })}`,
    "",
    deckText ? `Current deck text:\n${deckText}` : "Current deck text: (empty)",
  ].join("\n");

  return { instructions, prompt };
}

function createPlan(payload, cb) {
  const message = String(payload.message || "").trim();
  const source = String(payload.source || "");
  const imageAttachments = normalizeImageAttachments(payload.images || payload.imageAttachments || []);
  const history = summarizeHistory(payload.history || []);
  const defaultMode = POLICIES.defaultMode === "edit" ? "edit" : "review";
  const requestedModeRaw = String(payload.requestedMode || payload.mode || defaultMode).toLowerCase();
  const requestedMode = requestedModeRaw === "edit" ? "edit" : "review";
  const readOnlyIntent = requestedMode === "review" || isReadOnlyRequest(message);
  if (!message) return cb(new Error("메시지를 입력하세요."));

  getContextSnapshot((contextErr, context) => {
    if (contextErr) return cb(contextErr);
    const planId = makeId("plan");
    const requestId = planId;
    const built = buildPlannerPrompt(message, source, history, context, readOnlyIntent, requestedMode, imageAttachments);
    callCodexOAuth(built.instructions, built.prompt, PRESENTATION_PLAN_SCHEMA, requestId, imageAttachments, (aiErr, text, rawPath) => {
      if (aiErr) return cb(aiErr);
      let raw;
      let plan;
      try {
        raw = extractJsonObject(text);
        plan = normalizePresentationPlan(raw, context);
        if (readOnlyIntent) {
          plan.kind = "review";
          plan.legacyActions = [];
          plan.requiresApproval = false;
        }
        validatePresentationPlan(plan);
      } catch (err) {
        return cb(err);
      }
      const executionPlan = compileExecutionPlan(plan, context, planId, null);
      const record = {
        ok: true,
        planId,
        status: "compiled",
        generatedAt: new Date().toISOString(),
        rawResponsePath: rawPath,
        schema: {
          presentationPlan: path.relative(ROOT, PRESENTATION_PLAN_SCHEMA),
          executionPlan: path.relative(ROOT, EXECUTION_PLAN_SCHEMA),
        },
        deckFingerprint: context.deckFingerprint,
        requestedMode,
        imageAttachments,
        context,
        presentationPlan: plan,
        executionPlan,
      };
      saveJson(safePathForPlan(planId), record);
      cb(null, {
        ok: true,
        plan: publicPlan(plan, record),
        executionPlan: {
          transactionId: executionPlan.transactionId,
          operationCount: executionPlan.operations.length,
          affectedSlideIds: executionPlan.affectedSlideIds,
          maxRisk: executionPlan.maxRisk || "low",
        },
        context: {
          slideIndex: context.slideIndex,
          slideId: context.slideId,
          slideCount: context.slideCount,
          deckFingerprint: context.deckFingerprint,
          selection: context.selection || {},
        },
      });
    });
  });
}

function directCommitPlan(planId, selectedChangeIds, cb) {
  let planRecord;
  try {
    planRecord = loadPlanRecord(planId);
  } catch (err) {
    return cb(err);
  }
  if ((planRecord.requestedMode || "review") === "review") {
    const err = new Error("review 모드 계획은 바로 적용할 수 없습니다.");
    err.statusCode = 409;
    return cb(err);
  }

  getContextSnapshot((contextErr, currentContext) => {
    if (contextErr) return cb(contextErr);
    let executionPlan;
    let transaction;
    try {
      executionPlan = compileExecutionPlan(planRecord.presentationPlan, currentContext, planRecord.planId, selectedChangeIds);
      if (!executionPlan.legacyActions.length) {
        const err = new Error("선택된 편집 operation이 없습니다.");
        err.statusCode = 409;
        throw err;
      }
      assertAffectedSlidesUnchanged(affectedSlideSnapshots(planRecord.context, executionPlan), currentContext);
      transaction = newTransaction(planRecord, executionPlan, selectedChangeIds || null, "previewed");
      transaction.preview = {
        created: false,
        shadowDeckPath: null,
        slides: [],
        liveDeckModified: null,
        skippedReason: "direct_apply_policy",
      };
      saveTransaction(transaction);
    } catch (err) {
      return cb(err);
    }
    commitTransaction(transaction.transactionId, { approved: true }, cb);
  });
}

function commitPlan(planId, selectedChangeIds, cb) {
  if (POLICIES.preview && POLICIES.preview.requiredBeforeCommit === false) {
    return directCommitPlan(planId, selectedChangeIds, cb);
  }
  previewPlan(planId, selectedChangeIds, (previewErr, preview) => {
    if (previewErr) return cb(previewErr);
    commitTransaction(preview.transactionId, { approved: true }, cb);
  });
}

function applyLegacyPlan(payload, cb) {
  const rawPlan = payload && payload.plan ? payload.plan : payload;
  getContextSnapshot((contextErr, context) => {
    if (contextErr) return cb(contextErr);
    let plan;
    let planId;
    let executionPlan;
    try {
      plan = normalizePresentationPlan(rawPlan, context);
      validatePresentationPlan(plan);
      if (!plan.legacyActions.length) return cb(null, { ok: true, applied: false, plan, result: { ok: true, results: [] } });
      planId = makeId("plan");
      executionPlan = compileExecutionPlan(plan, context, planId, payload && payload.selectedChangeIds);
    } catch (err) {
      return cb(err);
    }
    const record = {
      ok: true,
      planId,
      status: "compiled",
      generatedAt: new Date().toISOString(),
      rawResponsePath: null,
      schema: {
        presentationPlan: path.relative(ROOT, PRESENTATION_PLAN_SCHEMA),
        executionPlan: path.relative(ROOT, EXECUTION_PLAN_SCHEMA),
      },
      deckFingerprint: context.deckFingerprint,
      requestedMode: "edit",
      context,
      presentationPlan: plan,
      executionPlan,
    };
    saveJson(safePathForPlan(planId), record);
    commitPlan(planId, payload && payload.selectedChangeIds, cb);
  });
}

function getCodexStatus(cb) {
  const child = spawn("codex", ["doctor", "--json"], { cwd: ROOT, windowsHide: true });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => stdout += chunk);
  child.stderr.on("data", (chunk) => stderr += chunk);
  child.on("close", (code) => {
    if (code !== 0) {
      cb(null, {
        cli: false,
        modelMode: "codex-config-default",
        modelOverride: false,
        effectiveModel: null,
        effectiveModelVerified: false,
        modelSource: null,
        authMode: "codex-chatgpt-oauth",
        warning: stderr.trim(),
      });
      return;
    }
    try {
      const doctor = JSON.parse(stdout);
      const config = doctor.checks && doctor.checks["config.load"] && doctor.checks["config.load"].details;
      const auth = doctor.checks && doctor.checks["auth.credentials"] && doctor.checks["auth.credentials"].details;
      const effectiveModel = config && config.model ? config.model : null;
      cb(null, {
        cli: true,
        modelMode: "codex-config-default",
        modelOverride: false,
        effectiveModel,
        effectiveModelVerified: !!effectiveModel,
        modelSource: effectiveModel ? "codex doctor --json config.load" : null,
        model: effectiveModel || "Codex 현재 설정",
        provider: config && config["model provider"] ? config["model provider"] : "openai",
        authMode: auth && auth["stored auth mode"] ? auth["stored auth mode"] : "chatgpt",
        hasApiKey: auth && auth["stored API key"] === "true",
        hasChatGptTokens: auth && auth["stored ChatGPT tokens"] === "true",
      });
    } catch (err) {
      cb(null, {
        cli: true,
        modelMode: "codex-config-default",
        modelOverride: false,
        effectiveModel: null,
        effectiveModelVerified: false,
        modelSource: null,
        authMode: "codex-chatgpt-oauth",
        warning: err.message,
      });
    }
  });
}

function extractSource(payload, cb) {
  const fileName = payload && typeof payload.name === "string" ? payload.name : "source";
  const contentBase64 = payload && typeof payload.contentBase64 === "string" ? payload.contentBase64 : "";
  if (!contentBase64) return cb(new Error("No file content was provided."));

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "local-gpt-source-"));
  const inputPath = path.join(tempDir, "source.bin");
  try {
    fs.writeFileSync(inputPath, Buffer.from(contentBase64, "base64"));
  } catch (err) {
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch (_) {}
    cb(err);
    return;
  }

  const args = ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", SOURCE_EXTRACTOR, "-InputPath", inputPath, "-FileName", fileName];
  const child = spawn("powershell.exe", args, { windowsHide: true });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => stdout += chunk);
  child.stderr.on("data", (chunk) => stderr += chunk);
  child.on("close", (code) => {
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch (_) {}
    let data = null;
    try { data = stdout.trim() ? JSON.parse(stdout.trim()) : null; } catch (_) {}
    if (code !== 0 || !data || data.ok === false) {
      cb(new Error((data && data.error) || stderr || stdout || "Source extraction failed."));
      return;
    }
    cb(null, data);
  });
}

function saveImageAttachment(payload, cb) {
  const fileName = payload && typeof payload.name === "string" ? payload.name : "pasted-image.png";
  const mimeType = payload && typeof payload.mimeType === "string" ? payload.mimeType : "image/png";
  const contentBase64 = payload && typeof payload.contentBase64 === "string" ? payload.contentBase64 : "";
  if (!/^image\/(png|jpeg|jpg|webp|gif|bmp)$/i.test(mimeType)) {
    const err = new Error("Only image clipboard/file attachments are supported.");
    err.statusCode = 415;
    return cb(err);
  }
  if (!contentBase64) return cb(new Error("No image content was provided."));
  let buffer = null;
  try {
    buffer = Buffer.from(contentBase64, "base64");
  } catch (err) {
    return cb(err);
  }
  if (!buffer || !buffer.length) return cb(new Error("Image content was empty."));
  if (buffer.length > MAX_IMAGE_BYTES) {
    const err = new Error(`Image is too large. Maximum is ${Math.floor(MAX_IMAGE_BYTES / 1024 / 1024)}MB.`);
    err.statusCode = 413;
    return cb(err);
  }
  try {
    const uploadId = makeId("img");
    const ext = imageExtensionForMime(mimeType, fileName);
    const imagePath = safeUploadPath(uploadId, ext);
    fs.mkdirSync(path.dirname(imagePath), { recursive: true });
    fs.writeFileSync(imagePath, buffer);
    cb(null, {
      ok: true,
      id: uploadId,
      name: safeFileName(fileName),
      mimeType,
      sizeBytes: buffer.length,
      path: imagePath,
      url: uploadUrl(imagePath),
    });
  } catch (err) {
    cb(err);
  }
}

function readClipboardImage(cb) {
  let imagePath = null;
  try {
    const uploadId = makeId("img");
    imagePath = safeUploadPath(uploadId, ".png");
  } catch (err) {
    cb(err);
    return;
  }

  const args = [
    "-Sta",
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    CLIPBOARD_IMAGE_READER,
    "-OutputPath",
    imagePath,
  ];
  const child = spawn("powershell.exe", args, { windowsHide: true });
  let stdout = "";
  let stderr = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk) => stdout += chunk);
  child.stderr.on("data", (chunk) => stderr += chunk);
  child.on("close", (code) => {
    let data = null;
    try { data = stdout.trim() ? JSON.parse(stdout.trim()) : null; } catch (_) {}
    if (code !== 0 || !data || data.ok === false) {
      cb(new Error((data && data.error) || stderr || stdout || "Clipboard image read failed."));
      return;
    }
    if (!data.hasImage) {
      cb(null, { ok: true, hasImage: false });
      return;
    }
    let stats = null;
    try { stats = fs.statSync(imagePath); } catch (_) {}
    if (!stats || !stats.isFile() || stats.size <= 0) {
      cb(new Error("Clipboard image was not saved."));
      return;
    }
    if (stats.size > MAX_IMAGE_BYTES) {
      try { fs.unlinkSync(imagePath); } catch (_) {}
      const err = new Error(`Image is too large. Maximum is ${Math.floor(MAX_IMAGE_BYTES / 1024 / 1024)}MB.`);
      err.statusCode = 413;
      cb(err);
      return;
    }
    cb(null, {
      ok: true,
      hasImage: true,
      id: path.basename(imagePath, path.extname(imagePath)),
      name: "clipboard-image.png",
      mimeType: "image/png",
      sizeBytes: stats.size,
      path: imagePath,
      url: uploadUrl(imagePath),
    });
  });
}

function getCapabilities(cb) {
  getCodexStatus((_, codex) => {
    getContextSnapshot((contextErr, context) => {
      const powerPointCom = !contextErr;
      cb(null, {
        ok: true,
        officeVersion: powerPointCom ? context.officeVersion || null : null,
        officeBitness: null,
        mode: "office2016-legacy",
        officeJsLoaded: false,
        powerPointCom,
        codexCli: !!codex.cli,
        codexAuthenticated: !!codex.hasChatGptTokens,
        authMode: "codex-chatgpt-oauth",
        modelMode: codex.modelMode || "codex-config-default",
        modelOverride: false,
        effectiveModel: codex.effectiveModel || null,
        effectiveModelVerified: !!codex.effectiveModelVerified,
        modelSource: codex.modelSource || null,
        presentation: powerPointCom ? {
          name: context.presentationName,
          slideCount: context.slideCount,
          deckFingerprint: context.deckFingerprint,
        } : null,
        errors: powerPointCom ? [] : [contextErr.message],
        features: {
          selectionRead: powerPointCom,
          activeSlideShapeMap: powerPointCom,
          slideIdSnapshot: powerPointCom,
          shapeTags: powerPointCom,
          slideExport: powerPointCom,
          saveCopyAs: powerPointCom,
          startNewUndoEntry: !!(POLICIES.undo && POLICIES.undo.startNewEntryBeforeApply),
          addChart2: false,
          visualPreview: powerPointCom,
          imagePaste: true,
          codexImageInput: true,
          schemaValidatedPlanning: fs.existsSync(PRESENTATION_PLAN_SCHEMA),
          transactionPreview: true,
          backupBeforeCommit: !!(POLICIES.backup && POLICIES.backup.beforeCommit),
          liveBusinessDeckCommit: !!(POLICIES.liveDeck && POLICIES.liveDeck.allowLiveBusinessDeckCommit),
          replaceDeckLiveCommit: !!(POLICIES.replaceDeck && POLICIES.replaceDeck.liveCommitEnabled),
          selectionEditCommit: !!(POLICIES.selectionEdit && POLICIES.selectionEdit.commitEnabled),
        },
        policies: {
          defaultMode: POLICIES.defaultMode || "review",
          previewRequiredBeforeCommit: !!(POLICIES.preview && POLICIES.preview.requiredBeforeCommit),
          previewOpenMode: POLICIES.preview && POLICIES.preview.openMode,
          backupRetention: POLICIES.backupRetention,
          backup: POLICIES.backup,
          undo: POLICIES.undo,
          liveDeck: POLICIES.liveDeck,
          replaceDeck: POLICIES.replaceDeck,
          selectionEdit: POLICIES.selectionEdit,
        },
      });
    });
  });
}

function sendError(res, err) {
  json(res, err.statusCode || 500, { ok: false, error: err.message, validationErrors: err.validationErrors || undefined });
}

function handlePlanRoute(req, res, routePath) {
  if (routePath === "/api/plans" && req.method === "POST") {
    return readBody(req, (err, body) => {
      if (err) return json(res, 400, { ok: false, error: "Could not read JSON." });
      createPlan(body || {}, (planErr, result) => {
        if (planErr) return sendError(res, planErr);
        json(res, 200, result);
      });
    });
  }

  const match = routePath.match(/^\/api\/plans\/([^/]+)(?:\/(compile|preview|commit))?$/);
  if (!match) return false;
  const planId = match[1];
  const action = match[2] || "get";

  if (action === "get" && req.method === "GET") {
    try {
      const record = loadPlanRecord(planId);
      return json(res, 200, {
        ok: true,
        plan: publicPlan(record.presentationPlan, record),
        executionPlan: record.executionPlan,
        status: record.status,
      });
    } catch (err) {
      return sendError(res, err);
    }
  }

  if (action === "compile" && req.method === "POST") {
    return readBody(req, (err, body) => {
      if (err) return json(res, 400, { ok: false, error: "Could not read JSON." });
      try {
        const record = loadPlanRecord(planId);
        const executionPlan = compileExecutionPlan(record.presentationPlan, record.context, record.planId, body && body.selectedChangeIds);
        return json(res, 200, { ok: true, executionPlan });
      } catch (compileErr) {
        return sendError(res, compileErr);
      }
    });
  }

  if (action === "preview" && req.method === "POST") {
    return readBody(req, (err, body) => {
      if (err) return json(res, 400, { ok: false, error: "Could not read JSON." });
      previewPlan(planId, body && body.selectedChangeIds, (previewErr, result) => {
        if (previewErr) return sendError(res, previewErr);
        json(res, 200, result);
      });
    });
  }

  if (action === "commit" && req.method === "POST") {
    return readBody(req, (err, body) => {
      if (err) return json(res, 400, { ok: false, error: "Could not read JSON." });
      commitPlan(planId, body && body.selectedChangeIds, (commitErr, result) => {
        if (commitErr) return sendError(res, commitErr);
        json(res, 200, result);
      });
    });
  }

  return json(res, 405, { ok: false, error: "Unsupported plan route method." });
}

function serveTransactionAsset(req, res, transactionId, assetName) {
  let assetPath;
  try {
    assetPath = safeAssetPath(transactionId, assetName);
  } catch (err) {
    return sendError(res, err);
  }
  if (!fs.existsSync(assetPath)) return json(res, 404, { ok: false, error: "Asset not found." });
  const ext = path.extname(assetPath).toLowerCase();
  const type = ext === ".png" ? "image/png"
    : ext === ".jpg" || ext === ".jpeg" ? "image/jpeg"
    : "application/octet-stream";
  res.writeHead(200, { "Content-Type": type, "Cache-Control": "no-store" });
  fs.createReadStream(assetPath).pipe(res);
}

function serveUploadedImage(req, res, assetName) {
  let imagePath;
  try {
    imagePath = safeUploadedImagePath(assetName);
  } catch (err) {
    return sendError(res, err);
  }
  if (!fs.existsSync(imagePath)) return json(res, 404, { ok: false, error: "Image not found." });
  const ext = path.extname(imagePath).toLowerCase();
  const type = ext === ".png" ? "image/png"
    : ext === ".jpg" || ext === ".jpeg" ? "image/jpeg"
    : ext === ".webp" ? "image/webp"
    : ext === ".gif" ? "image/gif"
    : ext === ".bmp" ? "image/bmp"
    : "application/octet-stream";
  res.writeHead(200, { "Content-Type": type, "Cache-Control": "no-store" });
  fs.createReadStream(imagePath).pipe(res);
}

function handleTransactionRoute(req, res, routePath) {
  const asset = routePath.match(/^\/api\/transactions\/([^/]+)\/assets\/([^/]+)$/);
  if (asset && req.method === "GET") {
    return serveTransactionAsset(req, res, asset[1], decodeURIComponent(asset[2]));
  }

  const match = routePath.match(/^\/api\/transactions\/([^/]+)(?:\/(commit|rollback))?$/);
  if (!match) return false;
  const transactionId = match[1];
  const action = match[2] || "get";

  if (action === "get" && req.method === "GET") {
    try {
      return json(res, 200, { ok: true, transaction: transactionPublic(loadTransaction(transactionId)) });
    } catch (err) {
      return sendError(res, err);
    }
  }

  if (action === "commit" && req.method === "POST") {
    return readBody(req, (err, body) => {
      if (err) return json(res, 400, { ok: false, error: "Could not read JSON." });
      commitTransaction(transactionId, body || {}, (commitErr, result) => {
        if (commitErr) return sendError(res, commitErr);
        json(res, 200, result);
      });
    });
  }

  if (action === "rollback" && req.method === "POST") {
    return rollbackTransaction(transactionId, (rollbackErr, result) => {
      if (rollbackErr) return sendError(res, rollbackErr);
      json(res, 200, result);
    });
  }

  return json(res, 405, { ok: false, error: "Unsupported transaction route method." });
}

function handleApi(req, res) {
  if (req.method === "OPTIONS") return json(res, 200, { ok: true });

  const routePath = (req.url || "").split("?")[0];

  if (routePath === "/api/health" && req.method === "GET") {
    return getCodexStatus((_, status) => {
      json(res, 200, {
        ok: true,
        port: PORT,
        authMode: "codex-chatgpt-oauth",
        codex: status,
        modelMode: status.modelMode || "codex-config-default",
        modelOverride: false,
        effectiveModel: status.effectiveModel || null,
        effectiveModelVerified: !!status.effectiveModelVerified,
        modelSource: status.modelSource || null,
      });
    });
  }

  if (routePath === "/api/capabilities" && req.method === "GET") {
    return getCapabilities((_, result) => json(res, 200, result));
  }

  if (routePath === "/api/ppt/context" && req.method === "GET") {
    return getContextSnapshot((err, result) => {
      if (err) return json(res, 500, { ok: false, error: err.message });
      json(res, 200, result);
    });
  }

  if (routePath === "/api/ppt/selection" && req.method === "GET") {
    return getSelectionSnapshot((err, result) => {
      if (err) return json(res, 500, { ok: false, error: err.message });
      json(res, 200, result);
    });
  }

  if (routePath === "/api/ppt/shape-map" && req.method === "GET") {
    const url = new URL(req.url || "/", "https://localhost");
    const maxSlidesRaw = Number(url.searchParams.get("maxSlides") || 12);
    const slideIndexRaw = Number(url.searchParams.get("slideIndex") || 0);
    const scopeRaw = String(url.searchParams.get("scope") || "active").toLowerCase();
    const payload = {
      scope: scopeRaw === "deck" ? "deck" : "active",
      maxSlides: Number.isFinite(maxSlidesRaw) ? maxSlidesRaw : 12,
      slideIndex: Number.isFinite(slideIndexRaw) ? slideIndexRaw : 0,
    };
    return getShapeMapSnapshot(payload, (err, result) => {
      if (err) return json(res, 500, { ok: false, error: err.message });
      json(res, 200, result);
    });
  }

  if (routePath === "/api/source/extract" && req.method === "POST") {
    return readBody(req, (err, body) => {
      if (err) return json(res, 400, { ok: false, error: "Could not read JSON." });
      extractSource(body || {}, (extractErr, result) => {
        if (extractErr) return sendError(res, extractErr);
        json(res, 200, result);
      });
    });
  }

  if (routePath === "/api/images" && req.method === "POST") {
    return readBody(req, (err, body) => {
      if (err) return json(res, 400, { ok: false, error: "Could not read JSON." });
      saveImageAttachment(body || {}, (imageErr, result) => {
        if (imageErr) return sendError(res, imageErr);
        json(res, 200, result);
      });
    });
  }

  const imageAsset = routePath.match(/^\/api\/images\/([^/]+)$/);
  if (imageAsset && req.method === "GET") {
    return serveUploadedImage(req, res, decodeURIComponent(imageAsset[1]));
  }

  if (routePath === "/api/clipboard/image" && req.method === "POST") {
    return readClipboardImage((imageErr, result) => {
      if (imageErr) return sendError(res, imageErr);
      json(res, 200, result);
    });
  }

  const transactionHandled = handleTransactionRoute(req, res, routePath);
  if (transactionHandled !== false) return;

  const planHandled = handlePlanRoute(req, res, routePath);
  if (planHandled !== false) return;

  if (routePath === "/api/chat/plan" && req.method === "POST") {
    return readBody(req, (err, body) => {
      if (err) return json(res, 400, { ok: false, error: "Could not read JSON." });
      createPlan(body || {}, (planErr, result) => {
        if (planErr) return sendError(res, planErr);
        json(res, 200, result);
      });
    });
  }

  if (routePath === "/api/chat/apply" && req.method === "POST") {
    return readBody(req, (err, body) => {
      if (err) return json(res, 400, { ok: false, error: "Could not read JSON." });
      const planId = body && (body.planId || (body.plan && body.plan.planId));
      if (planId) {
        return commitPlan(planId, body.selectedChangeIds, (commitErr, result) => {
          if (commitErr) return sendError(res, commitErr);
          json(res, 200, result);
        });
      }
      applyLegacyPlan(body || {}, (applyErr, result) => {
        if (applyErr) return sendError(res, applyErr);
        json(res, 200, result);
      });
    });
  }

  json(res, 404, { ok: false, error: "Unknown API route" });
}

function isAllowedHost(req) {
  const host = String(req.headers.host || "").toLowerCase();
  return /^localhost(?::\d+)?$/.test(host)
    || /^127\.0\.0\.1(?::\d+)?$/.test(host)
    || /^\[::1\](?::\d+)?$/.test(host);
}

const requestHandler = (req, res) => {
  if (!isAllowedHost(req)) {
    return json(res, 403, { ok: false, error: "Forbidden host." }, req);
  }
  if ((req.url || "").startsWith("/api/")) return handleApi(req, res);
  serveFile(req, res);
};

const useHttps = fs.existsSync(PFX_PATH);
const server = useHttps
  ? https.createServer({ pfx: fs.readFileSync(PFX_PATH), passphrase: PFX_PASSPHRASE }, requestHandler)
  : http.createServer(requestHandler);

server.listen(PORT, useHttps ? "localhost" : "127.0.0.1", () => {
  console.log(`Local GPT for PowerPoint listening on ${useHttps ? "https://localhost" : "http://127.0.0.1"}:${PORT}`);
  console.log("Auth mode: Codex ChatGPT OAuth");
  console.log("Model mode: Codex config default (no --model override)");
});
