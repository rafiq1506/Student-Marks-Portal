(() => {
const sourceSelect = document.querySelector("#sourceSelect");
const statusBar = document.querySelector("#statusBar");
const statusIcon = document.querySelector("#statusIcon");
const statusText = document.querySelector("#statusText");
const excelFields = document.querySelector("#excelFields");
const googleFields = document.querySelector("#googleFields");
const paperNameExcel = document.querySelector("#paperNameExcel");
const paperNameGoogle = document.querySelector("#paperNameGoogle");
const sheetUrl = document.querySelector("#sheetUrl");
const excelZone = document.querySelector("#excelZone");
const fileInput = document.querySelector("#fileInput");
const activateBtn = document.querySelector("#activateBtn");
const btnText = document.querySelector("#btnText");

if (!sourceSelect || !statusBar || !activateBtn) return;

function adminPassword() {
  const dashboardPassword = document.querySelector("#admin-password")?.value.trim();
  return (dashboardPassword || sessionStorage.getItem("adminPassword") || "").trim();
}

function setStatus(text, type = "success", icon = "ti-circle-check") {
  statusText.textContent = text;
  statusBar.className = `status-bar ${type === "success" ? "" : type}`.trim();
  statusIcon.className = `ti ${icon}`;
}

function ensurePassword() {
  if (adminPassword()) return true;
  setStatus("Please open this page from the admin console before uploading data.", "error", "ti-alert-circle");
  return false;
}

async function adminFetch(url, options = {}) {
  const headers = new Headers(options.headers || {});
  headers.set("X-Admin-Password", adminPassword());
  const response = await fetch(url, { ...options, headers });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "Request failed.");
  return data;
}

function updateButtonForSource(source) {
  if (source === "excel") {
    btnText.textContent = "Upload Excel File";
    activateBtn.querySelector("i").className = "ti ti-upload";
    return;
  }

  if (source === "google") {
    btnText.textContent = "Activate Google Sheet";
    activateBtn.querySelector("i").className = "ti ti-check";
    return;
  }

  btnText.textContent = "Activate";
  activateBtn.querySelector("i").className = "ti ti-check";
}

function handleSourceChange(source) {
  excelFields.classList.toggle("hidden", source !== "excel");
  googleFields.classList.toggle("hidden", source !== "google");
  updateButtonForSource(source);

  if (source === "excel") {
    setStatus("Excel file upload selected.", "success", "ti-file-spreadsheet");
    return;
  }

  if (source === "google") {
    setStatus("Google Sheet upload selected.", "google", "ti-brand-google-drive");
    return;
  }

  setStatus("Select a data source to get started", "success", "ti-circle-check");
}

function handleFile(input) {
  if (!input.files || !input.files[0]) return;
  const file = input.files[0];
  excelZone.querySelector(".ez-title").textContent = file.name;
  excelZone.querySelector(".ez-sub").textContent = `${(file.size / 1024).toFixed(1)} KB - ready to upload`;
  excelZone.classList.add("has-file");
}

function setButtonLoading(loading, loadingText = "Processing...") {
  const icon = activateBtn.querySelector("i");
  if (loading) {
    activateBtn.disabled = true;
    icon.className = "ti ti-loader-2 spinner";
    btnText.textContent = loadingText;
    return;
  }

  activateBtn.disabled = false;
  updateButtonForSource(sourceSelect.value);
}

async function uploadExcel() {
  const paperName = paperNameExcel.value.trim();
  if (!paperName) {
    setStatus("Please enter the name of the paper.", "error", "ti-alert-circle");
    paperNameExcel.focus();
    return;
  }

  if (!fileInput.files.length) {
    setStatus("Please choose an Excel file first.", "error", "ti-alert-circle");
    return;
  }

  const formData = new FormData();
  formData.append("paperName", paperName);
  formData.append("file", fileInput.files[0]);

  setStatus("Uploading Excel data...", "success", "ti-loader-2");
  setButtonLoading(true, "Uploading...");

  try {
    await adminFetch("/api/admin/upload", { method: "POST", body: formData });
    setStatus("Excel data uploaded successfully. You can return to the admin panel.", "success", "ti-circle-check");
    paperNameExcel.value = "";
    fileInput.value = "";
    excelZone.querySelector(".ez-title").textContent = "Click to upload your Excel file";
    excelZone.querySelector(".ez-sub").textContent = "Supports .xlsx and .xls formats";
    excelZone.classList.remove("has-file");
    window.dispatchEvent(new CustomEvent("upload:data-success"));
  } catch (error) {
    setStatus(error.message, "error", "ti-alert-circle");
  } finally {
    setButtonLoading(false);
  }
}

async function uploadGoogleSheet() {
  const paperName = paperNameGoogle.value.trim();
  const googleSheetUrl = sheetUrl.value.trim();

  if (!paperName) {
    setStatus("Please enter the name of the paper.", "error", "ti-alert-circle");
    paperNameGoogle.focus();
    return;
  }

  if (!googleSheetUrl) {
    setStatus("Please enter the Google Sheet URL.", "error", "ti-alert-circle");
    sheetUrl.focus();
    return;
  }

  setStatus("Connecting to Google Sheet...", "google", "ti-loader-2");
  setButtonLoading(true, "Processing...");

  try {
    await adminFetch("/api/admin/upload-google", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ paperName, googleSheetUrl }),
    });
    setStatus("Data source activated successfully.", "success", "ti-circle-check");
    paperNameGoogle.value = "";
    sheetUrl.value = "";
    window.dispatchEvent(new CustomEvent("upload:data-success"));
  } catch (error) {
    setStatus(error.message, "error", "ti-alert-circle");
  } finally {
    setButtonLoading(false);
  }
}

async function handleActivate() {
  const source = sourceSelect.value;

  if (!source) {
    setStatus("Please select a data source first.", "error", "ti-alert-circle");
    sourceSelect.focus();
    return;
  }

  if (!ensurePassword()) return;

  if (source === "excel") {
    await uploadExcel();
    return;
  }

  await uploadGoogleSheet();
}

sourceSelect.addEventListener("change", () => handleSourceChange(sourceSelect.value));
fileInput.addEventListener("change", () => handleFile(fileInput));
excelZone.addEventListener("click", () => fileInput.click());
activateBtn.addEventListener("click", handleActivate);

["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
  excelZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    event.stopPropagation();
  });
});

["dragenter", "dragover"].forEach((eventName) => {
  excelZone.addEventListener(eventName, () => excelZone.classList.add("drag-over"));
});

["dragleave", "drop"].forEach((eventName) => {
  excelZone.addEventListener(eventName, () => excelZone.classList.remove("drag-over"));
});

excelZone.addEventListener("drop", (event) => {
  const files = event.dataTransfer.files;
  if (files.length && files[0].name.match(/\.(xlsx|xls)$/i)) {
    fileInput.files = files;
    handleFile(fileInput);
    setStatus("File ready for upload.", "success", "ti-circle-check");
    return;
  }

  setStatus("Please drop a valid Excel file.", "error", "ti-alert-circle");
});
})();
