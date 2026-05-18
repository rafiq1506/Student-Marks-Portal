const passwordInput = document.querySelector("#admin-password");
const message = document.querySelector("#admin-message");
const dashboardMessage = document.querySelector("#dashboard-message");
const loginPanel = document.querySelector("#admin-login-panel");
const adminStats = document.querySelector("#admin-stats");
const adminDashboard = document.querySelector("#admin-dashboard");
const paperList = document.querySelector("#paper-list");
const showUploadDataButton = document.querySelector("#show-upload-data");
const uploadPanel = document.querySelector("#admin-upload-panel");
const hideUploadDataButton = document.querySelector("#hide-upload-data");

let isAuthenticated = false;

function adminPassword() {
  return passwordInput.value.trim();
}

function setAdminMessage(text, type = "") {
  const target = isAuthenticated ? dashboardMessage : message;
  target.textContent = text;
  target.style.opacity = "1";
  target.className = `message-luxury ${type}`.trim();
  if (text) {
    setTimeout(() => {
      if (target.textContent === text) target.style.opacity = "0.7";
    }, 3000);
  }
}

function setButtonLoading(button, loading, loadingText) {
  if (!button) return;
  const textElement = button.querySelector(".btn-text") || button;
  if (loading) {
    button.dataset.originalText = textElement.textContent;
    if (loadingText) textElement.textContent = loadingText;
    button.classList.add("is-loading");
    button.disabled = true;
    return;
  }
  if (button.dataset.originalText) textElement.textContent = button.dataset.originalText;
  button.classList.remove("is-loading");
  button.disabled = false;
}

function sourceTypeLabel(sourceType) {
  return sourceType === "google_sheet" ? "Google Sheet" : "Excel Sheet";
}

function showDashboard() {
  isAuthenticated = true;
  sessionStorage.setItem("adminPassword", adminPassword());
  document.body.classList.remove("admin-login-mode");
  loginPanel.classList.add("hidden");
  adminStats.classList.remove("hidden");
  adminDashboard.classList.remove("hidden");
}

async function adminFetch(url, options = {}) {
  const headers = new Headers(options.headers || {});
  headers.set("X-Admin-Password", adminPassword());
  const response = await fetch(url, { ...options, headers });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "Request failed.");
  }
  return data;
}

function renderStatus(data) {
  paperList.innerHTML = "";
  const papers = data.papers || [];
  if (!papers.length) {
    paperList.innerHTML = '<div class="chip-skeleton">No paper uploads detected</div>';
    return;
  }

  const header = document.createElement("div");
  header.className = "paper-list-row paper-list-header";
  header.innerHTML = `
    <span>Paper</span>
    <span>Source File</span>
    <span></span>
  `;
  paperList.appendChild(header);

  papers.forEach((paper) => {
    const item = document.createElement("div");
    item.className = "paper-list-row";
    item.innerHTML = `
      <span class="paper-title">${paper.name || "Untitled Paper"}</span>
      <span class="paper-source">${sourceTypeLabel(paper.sourceType)}</span>
      <button class="paper-action delete-paper" data-paper-id="${paper.id}">Delete</button>
    `;
    paperList.appendChild(item);
  });
}

async function refreshStatus() {
  const button = document.querySelector("#refresh-status");
  if (!adminPassword()) {
    setAdminMessage("Please enter admin password first.", "error");
    return;
  }
  setAdminMessage("Authenticating and fetching uploaded papers...", "info");
  setButtonLoading(button, true, "Authenticating...");
  try {
    const data = await adminFetch("/api/admin/status");
    renderStatus(data);
    showDashboard();
    dashboardMessage.textContent = "";
    dashboardMessage.className = "message-luxury";
  } catch (error) {
    setAdminMessage(error.message, "error");
  } finally {
    setButtonLoading(button, false);
  }
}

document.querySelector("#refresh-status").addEventListener("click", refreshStatus);

showUploadDataButton.addEventListener("click", () => {
  sessionStorage.setItem("adminPassword", adminPassword());
  uploadPanel.classList.remove("hidden");
  uploadPanel.scrollIntoView({ behavior: "smooth", block: "start" });
});

hideUploadDataButton.addEventListener("click", () => {
  uploadPanel.classList.add("hidden");
});

window.addEventListener("upload:data-success", async () => {
  uploadPanel.classList.add("hidden");
  await refreshStatus();
  setAdminMessage("Data uploaded successfully.", "success");
});

paperList.addEventListener("click", async (event) => {
  const button = event.target.closest("button[data-paper-id]");
  if (!button) return;
  const paperId = button.dataset.paperId;
  if (!paperId) return;

  if (!window.confirm("Delete this paper data file?")) return;
  setAdminMessage("Deleting paper...", "info");
  setButtonLoading(button, true, "Deleting...");
  try {
    const data = await adminFetch("/api/admin/paper/delete", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ paperId }),
    });
    renderStatus(data);
    setAdminMessage("Paper deleted.", "success");
  } catch (error) {
    setAdminMessage(error.message, "error");
    setButtonLoading(button, false);
  }
});
