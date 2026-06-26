require("dotenv").config();

const crypto = require("crypto");
const fs = require("fs/promises");
const path = require("path");

const express = require("express");
const mammoth = require("mammoth");
const nodemailer = require("nodemailer");
const xlsx = require("xlsx");

const app = express();

const PORT = Number(process.env.PORT || 3000);
const SEND_INTERVAL_MS = Number(process.env.SEND_INTERVAL_MS || 2 * 60 * 1000);
const SEND_FIRST_IMMEDIATELY = parseBoolean(process.env.SEND_FIRST_IMMEDIATELY, false);

const jobs = new Map();
let activeJobId = null;
let lastJobId = null;

app.use(express.json({ limit: "1mb" }));
app.use(express.static(path.join(__dirname, "public")));

app.get("/api/files", async (_req, res) => {
  try {
    const wordDir = path.join(process.cwd(), "word");
    const excelDir = path.join(process.cwd(), "excel");
    
    await fs.mkdir(wordDir, { recursive: true }).catch(() => {});
    await fs.mkdir(excelDir, { recursive: true }).catch(() => {});

    const wordFiles = (await fs.readdir(wordDir)).filter(f => f.endsWith('.docx'));
    const excelFiles = (await fs.readdir(excelDir)).filter(f => f.endsWith('.xlsx'));

    res.json({ wordFiles, excelFiles });
  } catch (error) {
    res.status(500).json({ error: "Failed to read directories." });
  }
});

app.get("/api/config", async (_req, res) => {
  const missing = getMissingConfig();

  res.json({
    smtpReady: missing.length === 0,
    missing,
    intervalMs: SEND_INTERVAL_MS,
    intervalText: formatDuration(SEND_INTERVAL_MS),
    firstSendImmediately: SEND_FIRST_IMMEDIATELY,
    defaultSubject: process.env.MAIL_SUBJECT || "EU RATE BAO GIA",
    hasActiveJob: Boolean(activeJobId)
  });
});

app.get("/api/jobs/current", (_req, res) => {
  const visibleJobId = activeJobId || lastJobId;

  if (!visibleJobId || !jobs.has(visibleJobId)) {
    res.json({ job: null });
    return;
  }

  res.json({ job: toPublicJob(jobs.get(visibleJobId)) });
});

app.get("/api/jobs/:id", (req, res) => {
  const job = jobs.get(req.params.id);

  if (!job) {
    res.status(404).json({ error: "Job not found." });
    return;
  }

  res.json({ job: toPublicJob(job) });
});

app.post("/api/send-campaign", async (req, res) => {
  if (activeJobId) {
    res.status(409).json({ error: "A sending job is already running. Cancel it before starting a new one." });
    return;
  }

  const configErrors = getMissingConfig();
  if (configErrors.length > 0) {
    res.status(400).json({ error: "Missing Gmail SMTP configuration.", missing: configErrors });
    return;
  }

  const docxName = req.body.docxFile;
  const excelName = req.body.excelFile;

  if (!docxName || !excelName) {
    res.status(400).json({ error: "Vui lòng chọn file Word và file Excel." });
    return;
  }

  const docxPath = path.resolve(process.cwd(), "word", docxName);
  const excelPath = path.resolve(process.cwd(), "excel", excelName);

  if (!(await fileExists(docxPath))) {
    res.status(400).json({ error: `DOCX file not found: ${docxPath}` });
    return;
  }

  if (!(await fileExists(excelPath))) {
    res.status(400).json({ error: `Excel file not found: ${excelPath}` });
    return;
  }

  const parsed = readRecipientsFromExcel(excelPath);
  if (parsed.valid.length === 0) {
    res.status(400).json({ error: "No valid email addresses found in Excel.", invalid: parsed.invalid });
    return;
  }

  if (parsed.invalid.length > 0) {
    res.status(400).json({ error: "Some Excel rows are invalid.", invalid: parsed.invalid });
    return;
  }

  try {
    const template = await renderDocxTemplate(docxPath);
    const subject = normalizeSubject(req.body.subject);
    const job = createJob(parsed.valid, subject, template, docxName, excelName);

    res.status(201).json({ job: toPublicJob(job) });
  } catch (error) {
    res.status(500).json({ error: error.message || "Cannot prepare DOCX email content." });
  }
});

app.delete("/api/jobs/:id", (req, res) => {
  const job = jobs.get(req.params.id);

  if (!job) {
    res.status(404).json({ error: "Job not found." });
    return;
  }

  cancelJob(job, "Canceled from web UI.");
  res.json({ job: toPublicJob(job) });
});

app.listen(PORT, () => {
  console.log(`Gmail SMTP DOCX mailer is running at http://localhost:${PORT}`);
});

function createJob(recipients, subject, template, docxName, excelName) {
  const now = new Date();
  const job = {
    id: crypto.randomUUID(),
    status: "queued",
    recipients,
    subject,
    template,
    docxName,
    excelName,
    currentIndex: 0,
    sent: [],
    failed: [],
    logs: [],
    createdAt: now.toISOString(),
    startedAt: null,
    finishedAt: null,
    nextRunAt: null,
    timer: null
  };

  jobs.set(job.id, job);
  activeJobId = job.id;
  lastJobId = job.id;
  addLog(job, "info", `Prepared DOCX content with ${template.inlineAttachments.length} inline image(s).`);
  addLog(job, "info", `Loaded ${recipients.length} recipient(s) from ${excelName}.`);

  const firstDelay = SEND_FIRST_IMMEDIATELY ? 0 : SEND_INTERVAL_MS;
  scheduleNext(job, firstDelay);

  return job;
}

function scheduleNext(job, delayMs) {
  if (job.status === "canceled") {
    return;
  }

  job.status = job.currentIndex === 0 ? "queued" : "waiting";
  job.nextRunAt = new Date(Date.now() + delayMs).toISOString();
  addLog(job, "info", `Next send scheduled at ${job.nextRunAt}.`);

  job.timer = setTimeout(() => {
    sendNextEmail(job).catch((error) => {
      addLog(job, "error", error.message || "Unexpected job error.");
      finishJob(job, "failed");
    });
  }, delayMs);
}

async function sendNextEmail(job) {
  if (job.status === "canceled") {
    return;
  }

  if (job.currentIndex >= job.recipients.length) {
    finishJob(job, job.failed.length > 0 ? "completed_with_errors" : "completed");
    return;
  }

  if (!job.startedAt) {
    job.startedAt = new Date().toISOString();
  }

  job.status = "sending";
  job.nextRunAt = null;

  const recipient = job.recipients[job.currentIndex];
  addLog(job, "info", `Sending to ${formatRecipient(recipient)}.`);

  try {
    const info = await sendEmail(recipient, job.subject, job.template);
    job.sent.push({
      name: recipient.name,
      email: recipient.email,
      rowNumber: recipient.rowNumber,
      at: new Date().toISOString(),
      messageId: info.messageId || null
    });
    addLog(job, "success", `Sent to ${formatRecipient(recipient)}.`);
  } catch (error) {
    job.failed.push({
      name: recipient.name,
      email: recipient.email,
      rowNumber: recipient.rowNumber,
      at: new Date().toISOString(),
      error: error.message || "Send failed."
    });
    addLog(job, "error", `Failed to send to ${formatRecipient(recipient)}: ${error.message || "Unknown error."}`);
  }

  job.currentIndex += 1;

  if (job.currentIndex >= job.recipients.length) {
    finishJob(job, job.failed.length > 0 ? "completed_with_errors" : "completed");
    return;
  }

  scheduleNext(job, SEND_INTERVAL_MS);
}

async function sendEmail(recipient, subject, template) {
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    secure: parseBoolean(process.env.SMTP_SECURE, false),
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    }
  });

  const html = personalizeTemplate(template.html, recipient.name);

  return transporter.sendMail({
    from: process.env.MAIL_FROM || process.env.SMTP_USER,
    to: recipient.email,
    subject,
    html,
    attachments: template.inlineAttachments.map((attachment) => ({ ...attachment }))
  });
}

async function renderDocxTemplate(docxPath) {
  let imageIndex = 0;
  const inlineAttachments = [];

  const result = await mammoth.convertToHtml(
    { path: docxPath },
    {
      convertImage: mammoth.images.imgElement(async (image) => {
        imageIndex += 1;

        const base64 = await image.read("base64");
        const contentType = image.contentType || "application/octet-stream";
        const extension = contentTypeToExtension(contentType);
        const cid = `docx-image-${Date.now()}-${imageIndex}@gmail-docx-mailer`;

        inlineAttachments.push({
          filename: `docx-image-${imageIndex}.${extension}`,
          content: base64,
          encoding: "base64",
          cid,
          contentType,
          contentDisposition: "inline"
        });

        return { src: `cid:${cid}` };
      })
    }
  );

  const html = markLastImageTableAsBorderless(result.value);

  return {
    html: wrapEmailHtml(html),
    inlineAttachments,
    warnings: result.messages || []
  };
}

function wrapEmailHtml(body) {
  return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <style>
      body {
        margin: 0;
        padding: 0;
        background: #ffffff;
        color: #111827;
        font-family: Arial, Helvetica, sans-serif;
        font-size: 14px;
        line-height: 1.5;
      }

      .docx-mail {
        max-width: 920px;
        margin: 0 auto;
        padding: 16px;
      }

      p {
        margin: 0 0 10px;
      }

      table {
        border-collapse: collapse;
        margin: 12px 0;
        width: auto;
        max-width: 100%;
      }

      th,
      td {
        border: 1px solid #4b5563;
        padding: 6px 8px;
        vertical-align: top;
      }

      table.signature-table,
      table.signature-table th,
      table.signature-table td {
        border: 0 !important;
      }

      img {
        max-width: 100%;
        height: auto;
      }
    </style>
  </head>
  <body>
    <div class="docx-mail">
      ${body}
    </div>
  </body>
</html>`;
}

function finishJob(job, status) {
  if (job.timer) {
    clearTimeout(job.timer);
    job.timer = null;
  }

  job.status = status;
  job.finishedAt = new Date().toISOString();
  job.nextRunAt = null;
  addLog(job, status === "completed" ? "success" : "info", `Job ${status}.`);

  if (activeJobId === job.id) {
    activeJobId = null;
  }
}

function cancelJob(job, message) {
  if (job.timer) {
    clearTimeout(job.timer);
    job.timer = null;
  }

  job.status = "canceled";
  job.finishedAt = new Date().toISOString();
  job.nextRunAt = null;
  addLog(job, "info", message);

  if (activeJobId === job.id) {
    activeJobId = null;
  }
}

function toPublicJob(job) {
  return {
    id: job.id,
    status: job.status,
    subject: job.subject,
    docxName: job.docxName,
    excelName: job.excelName,
    total: job.recipients.length,
    currentIndex: job.currentIndex,
    sentCount: job.sent.length,
    failedCount: job.failed.length,
    remainingCount: Math.max(job.recipients.length - job.currentIndex, 0),
    nextRecipient: job.recipients[job.currentIndex] ? formatRecipient(job.recipients[job.currentIndex]) : null,
    createdAt: job.createdAt,
    startedAt: job.startedAt,
    finishedAt: job.finishedAt,
    nextRunAt: job.nextRunAt,
    sent: job.sent,
    failed: job.failed,
    logs: job.logs.slice(-50)
  };
}

function addLog(job, level, message) {
  const entry = {
    level,
    message,
    at: new Date().toISOString()
  };

  job.logs.push(entry);
  console.log(`[${entry.at}] [${job.id}] [${level}] ${message}`);
}

function readRecipientsFromExcel(excelPath) {
  const workbook = xlsx.readFile(excelPath, {
    cellDates: false
  });

  if (workbook.SheetNames.length === 0) {
    return {
      valid: [],
      invalid: [{ sheetName: "", rowNumber: 0, reason: "No worksheet found." }]
    };
  }

  const seen = new Set();
  const valid = [];
  const invalid = [];

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      invalid.push({ sheetName, rowNumber: 0, reason: "Worksheet is empty or unavailable." });
      return;
    }

    const rows = xlsx.utils.sheet_to_json(sheet, {
      header: 1,
      blankrows: false,
      defval: "",
      raw: false
    });

    rows.forEach((row, index) => {
      const rowNumber = index + 1;
      const name = String(row[0] || "").trim();
      const email = String(row[4] || "").trim();

      if (isHeaderRow(name, email) || (!name && !email)) {
        return;
      }

      if (!name) {
        invalid.push({ sheetName, rowNumber, email, reason: "Missing FULL NAME in column 1." });
        return;
      }

      if (!email) {
        invalid.push({ sheetName, rowNumber, name, reason: "Missing EMAIL in column 5." });
        return;
      }

      if (!isValidEmail(email)) {
        invalid.push({ sheetName, rowNumber, name, email, reason: "Invalid email format." });
        return;
      }

      const normalized = email.toLowerCase();
      if (seen.has(normalized)) {
        invalid.push({ sheetName, rowNumber, name, email, reason: "Duplicate email across workbook." });
        return;
      }

      seen.add(normalized);
      valid.push({ name, email, sheetName, rowNumber });
    });
  });

  return { valid, invalid };
}

function personalizeTemplate(html, name) {
  const safeName = escapeHtml(name);

  return html
    .replace(/\{\{\s*Name\s*\}\}/gi, safeName)
    .replace(/\{\s*Name\s*\}/gi, safeName);
}

function markLastImageTableAsBorderless(html) {
  const tableRegex = /<table\b[\s\S]*?<\/table>/gi;
  let match;
  let lastImageTable = null;

  while ((match = tableRegex.exec(html)) !== null) {
    if (/<img\b/i.test(match[0])) {
      lastImageTable = {
        start: match.index,
        end: tableRegex.lastIndex,
        html: match[0]
      };
    }
  }

  if (!lastImageTable) {
    return html;
  }

  const updated = addClassToOpeningTag(lastImageTable.html, "table", "signature-table");
  return `${html.slice(0, lastImageTable.start)}${updated}${html.slice(lastImageTable.end)}`;
}

function addClassToOpeningTag(html, tagName, className) {
  const openingTagRegex = new RegExp(`<${tagName}\\b([^>]*)>`, "i");
  const match = html.match(openingTagRegex);

  if (!match) {
    return html;
  }

  const openingTag = match[0];
  const updatedOpeningTag = /\sclass=["'][^"']*["']/i.test(openingTag)
    ? openingTag.replace(/class=(["'])([^"']*)(["'])/i, `class=$1$2 ${className}$3`)
    : openingTag.replace(new RegExp(`^<${tagName}\\b`, "i"), `<${tagName} class="${className}"`);

  return html.replace(openingTag, updatedOpeningTag);
}

function isHeaderRow(name, email) {
  return normalizeHeader(name) === "full name" || normalizeHeader(email) === "email";
}

function normalizeHeader(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function formatRecipient(recipient) {
  const source = recipient.sheetName ? ` (${recipient.sheetName} row ${recipient.rowNumber})` : "";
  return `${recipient.name} <${recipient.email}>${source}`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeSubject(value) {
  const subject = String(value || "").trim();
  return subject || process.env.MAIL_SUBJECT || "EU RATE BAO GIA";
}

function getMissingConfig() {
  return ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASS"]
    .filter((key) => !String(process.env[key] || "").trim());
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch (_error) {
    return false;
  }
}

function parseBoolean(value, fallback) {
  if (value === undefined || value === null || value === "") {
    return fallback;
  }

  return ["1", "true", "yes", "on"].includes(String(value).trim().toLowerCase());
}

function contentTypeToExtension(contentType) {
  const map = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/webp": "webp",
    "image/svg+xml": "svg"
  };

  return map[contentType] || "bin";
}

function formatDuration(ms) {
  const seconds = Math.round(ms / 1000);

  if (seconds < 60) {
    return `${seconds} second(s)`;
  }

  const minutes = Math.round(seconds / 60);
  return `${minutes} minute(s)`;
}
