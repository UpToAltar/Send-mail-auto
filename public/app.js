const sendForm = document.querySelector("#sendForm");
const subjectInput = document.querySelector("#subject");
const docxFileSelect = document.querySelector("#docxFileSelect");
const excelFileSelect = document.querySelector("#excelFileSelect");
const startButton = document.querySelector("#startButton");
const cancelButton = document.querySelector("#cancelButton");
const refreshButton = document.querySelector("#refreshButton");
const notice = document.querySelector("#notice");
const configBadge = document.querySelector("#configBadge");
const docxFile = document.querySelector("#docxFile");
const excelFile = document.querySelector("#excelFile");
const intervalText = document.querySelector("#intervalText");
const jobStatus = document.querySelector("#jobStatus");
const progressText = document.querySelector("#progressText");
const nextRecipient = document.querySelector("#nextRecipient");
const nextRunAt = document.querySelector("#nextRunAt");
const logs = document.querySelector("#logs");

let currentJobId = null;

sendForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  setNotice("", "");
  startButton.disabled = true;

  try {
    const response = await fetch("/api/send-campaign", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        subject: subjectInput.value,
        docxFile: docxFileSelect.value,
        excelFile: excelFileSelect.value
      })
    });

    const data = await response.json();
    if (!response.ok) {
      throw new Error(formatApiError(data));
    }

    currentJobId = data.job.id;
    renderJob(data.job);
    setNotice("Đã tạo job gửi mail.", "success");
  } catch (error) {
    setNotice(error.message, "error");
  } finally {
    startButton.disabled = false;
  }
});

cancelButton.addEventListener("click", async () => {
  if (!currentJobId) {
    return;
  }

  cancelButton.disabled = true;

  try {
    const response = await fetch(`/api/jobs/${currentJobId}`, {
      method: "DELETE"
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(formatApiError(data));
    }

    renderJob(data.job);
    setNotice("Đã hủy job.", "success");
  } catch (error) {
    setNotice(error.message, "error");
  }
});

refreshButton.addEventListener("click", refresh);

refresh();
setInterval(refresh, 5000);

async function refresh() {
  await Promise.all([loadConfig(), loadFiles(), loadCurrentJob()]);
}

async function loadFiles() {
  try {
    const response = await fetch("/api/files");
    const data = await response.json();
    
    populateSelect(docxFileSelect, data.wordFiles);
    populateSelect(excelFileSelect, data.excelFiles);
  } catch (_error) {
    console.error("Lỗi lấy danh sách file");
  }
}

function populateSelect(selectEl, files) {
  const currentVal = selectEl.value;
  selectEl.innerHTML = "";
  if (!files || files.length === 0) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "-- Không có file --";
    selectEl.appendChild(opt);
  } else {
    files.forEach(f => {
      const opt = document.createElement("option");
      opt.value = f;
      opt.textContent = f;
      selectEl.appendChild(opt);
    });
    if (files.includes(currentVal)) {
      selectEl.value = currentVal;
    }
  }
}

async function loadConfig() {
  try {
    const response = await fetch("/api/config");
    const config = await response.json();

    intervalText.textContent = config.firstSendImmediately
      ? `${config.intervalText}, gửi mail đầu tiên ngay`
      : `${config.intervalText}, gửi mail đầu tiên sau chu kỳ này`;

    if (!subjectInput.value && config.defaultSubject) {
      subjectInput.value = config.defaultSubject;
    }

    if (!config.smtpReady) {
      setBadge(`Thiếu ENV SMTP: ${config.missing.join(", ")}`, "error");
      return;
    }

    setBadge("Gmail SMTP sẵn sàng", "success");
  } catch (_error) {
    setBadge("Không đọc được config", "error");
  }
}

async function loadCurrentJob() {
  try {
    const response = await fetch("/api/jobs/current");
    const data = await response.json();
    renderJob(data.job);
  } catch (_error) {
    renderJob(null);
  }
}

function renderJob(job) {
  if (!job) {
    currentJobId = null;
    jobStatus.textContent = "Chưa chạy";
    progressText.textContent = "0 / 0";
    nextRecipient.textContent = "-";
    nextRunAt.textContent = "-";
    docxFile.textContent = "-";
    excelFile.textContent = "-";
    logs.innerHTML = "";
    cancelButton.disabled = true;
    return;
  }

  currentJobId = job.id;
  jobStatus.textContent = statusLabel(job.status);
  progressText.textContent = `${job.sentCount + job.failedCount} / ${job.total} (${job.sentCount} thành công, ${job.failedCount} lỗi)`;
  nextRecipient.textContent = job.nextRecipient || "-";
  nextRunAt.textContent = job.nextRunAt ? formatDateTime(job.nextRunAt) : "-";
  docxFile.textContent = job.docxName || "-";
  excelFile.textContent = job.excelName || "-";
  cancelButton.disabled = !["queued", "waiting", "sending"].includes(job.status);

  logs.innerHTML = job.logs
    .slice()
    .reverse()
    .map((entry) => {
      return `<li class="${entry.level}">
        <time>${formatDateTime(entry.at)}</time>
        <span>${escapeHtml(entry.message)}</span>
      </li>`;
    })
    .join("");
}

function setBadge(text, type) {
  configBadge.textContent = text;
  configBadge.className = `badge ${type || ""}`.trim();
}

function setNotice(message, type) {
  notice.textContent = message;
  notice.className = message ? `notice show ${type || ""}` : "notice";
}

function formatApiError(data) {
  if (!data) {
    return "Có lỗi xảy ra.";
  }

  if (data.invalid && data.invalid.length > 0) {
    return `${data.error} ${formatInvalidRows(data.invalid)}`;
  }

  if (data.missing && data.missing.length > 0) {
    return `${data.error} Thiếu: ${data.missing.join(", ")}`;
  }

  return data.error || "Có lỗi xảy ra.";
}

function formatInvalidRows(invalidRows) {
  return invalidRows
    .map((row) => {
      if (typeof row === "string") {
        return row;
      }

      const parts = [];
      if (row.sheetName) {
        parts.push(row.sheetName);
      }
      parts.push(`dòng ${row.rowNumber || "?"}`);
      if (row.name) {
        parts.push(row.name);
      }
      if (row.email) {
        parts.push(row.email);
      }
      if (row.reason) {
        parts.push(row.reason);
      }

      return parts.join(" - ");
    })
    .join("; ");
}

function statusLabel(status) {
  const labels = {
    queued: "Đang chờ",
    waiting: "Chờ lần gửi kế tiếp",
    sending: "Đang gửi",
    completed: "Hoàn thành",
    completed_with_errors: "Hoàn thành có lỗi",
    failed: "Thất bại",
    canceled: "Đã hủy"
  };

  return labels[status] || status || "-";
}

function formatDateTime(value) {
  return new Intl.DateTimeFormat("vi-VN", {
    dateStyle: "short",
    timeStyle: "medium"
  }).format(new Date(value));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
