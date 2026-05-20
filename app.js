const loginForm = document.querySelector("#login-form");
const otpForm = document.querySelector("#otp-form");
const paperSelect = document.querySelector("#paper-select");
const rollInput = document.querySelector("#roll-number");
const otpInput = document.querySelector("#otp");
const sendButton = document.querySelector("#send-otp");
const verifyButton = document.querySelector("#verify-otp");
const changeRollButton = document.querySelector("#change-roll");
const startOverButton = document.querySelector("#start-over");
const messageDiv = document.querySelector("#message");
const otpStatusSpan = document.querySelector("#otp-status");

const studentLoginScreen = document.querySelector("#student-login-screen");
const studentDashboard = document.querySelector("#student-dashboard");
const emptyStateDiv = document.querySelector("#empty-state");
const studentMarksheetDiv = document.querySelector("#student-marksheet");
const verifiedSessionDiv = document.querySelector("#verified-session");

const studentNameSpan = document.querySelector("#student-name");
const courseNameSpan = document.querySelector("#course-name");
const paperNameSpan = document.querySelector("#paper-name");
const collegeRollSpan = document.querySelector("#college-roll");
const examRollSpan = document.querySelector("#exam-roll");
const assignmentSpan = document.querySelector("#assignment-marks");
const testSpan = document.querySelector("#test-marks");
const attendanceMarksSpan = document.querySelector("#attendance-marks");
const attendancePercentSpan = document.querySelector("#attendance");
const totalAttendanceSpan = document.querySelector("#total-attendance");
const attendanceMonthsBody = document.querySelector("#attendance-months");
const internalsSpan = document.querySelector("#internals");
const totalIaMarksSpan = document.querySelector("#total-ia-marks");
const iaBarEl = document.querySelector("#ia-bar");
const attBarEl = document.querySelector("#att-bar");
const iaPercentSubEl = document.querySelector("#ia-percent-sub");

let activeRollNumber = "";
let activePaperId = "";
const PAPER_CACHE_KEY = "studentPortalPapers";
const PAPER_OPTION_INDENT = "\u00a0\u00a0";

function fallback(value) {
  return value && value !== "" ? value : "-";
}

function twoDigit(value) {
  if (!value && value !== 0) return "-";
  const text = String(value).trim();
  if (!text) return "-";
  const number = Number(text);
  if (!Number.isFinite(number)) return text;
  if (!Number.isInteger(number)) return number.toFixed(1).replace(/\.0$/, "");
  return String(number).padStart(2, "0");
}

function setMessage(text, type = "") {
  messageDiv.textContent = text;
  messageDiv.className = `message ${type}`.trim();
}

function setButtonLoading(button, loading, loadingText) {
  if (!button) return;
  const textElement = button.querySelector(".btn-text") || button.querySelector("span") || button;
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

async function postJson(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "Something went wrong.");
  }
  return data;
}

function renderPaperOptions(papers) {
  paperSelect.innerHTML = "";
  if (!papers.length) {
    paperSelect.appendChild(new Option(`${PAPER_OPTION_INDENT}No papers available`, ""));
    sendButton.disabled = true;
    setMessage("No paper data has been uploaded yet.", "error");
    return;
  }
  paperSelect.appendChild(new Option(`${PAPER_OPTION_INDENT}Select paper`, ""));
  papers.forEach((paper) => {
    paperSelect.appendChild(new Option(`${PAPER_OPTION_INDENT}${paper.name}${PAPER_OPTION_INDENT}`, paper.id));
  });
  sendButton.disabled = false;
}

function loadCachedPapers() {
  try {
    const cached = JSON.parse(sessionStorage.getItem(PAPER_CACHE_KEY) || "[]");
    if (Array.isArray(cached) && cached.length) {
      renderPaperOptions(cached);
    }
  } catch {
    sessionStorage.removeItem(PAPER_CACHE_KEY);
  }
}

async function loadPapers() {
  loadCachedPapers();
  try {
    const response = await fetch("/api/papers");
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Could not load papers.");
    const papers = data.papers || [];
    sessionStorage.setItem(PAPER_CACHE_KEY, JSON.stringify(papers));
    renderPaperOptions(papers);
  } catch (error) {
    if (!paperSelect.options.length || paperSelect.options[0].textContent.trim() === "Loading papers...") {
      paperSelect.innerHTML = "";
      paperSelect.appendChild(new Option(`${PAPER_OPTION_INDENT}Could not load papers`, ""));
      sendButton.disabled = true;
    }
    setMessage(error.message, "error");
  }
}

function formatPercent(value) {
  if (!value && value !== 0) return "-";
  const number = parseFloat(value);
  if (Number.isNaN(number)) return value;
  return `${number.toFixed(1).replace(/\.0$/, "")}%`;
}

function showStudent(student) {
  // Header info
  studentNameSpan.textContent = fallback(student.name);
  courseNameSpan.textContent = fallback(student.course_name);
  paperNameSpan.textContent = fallback(
    student.paper_name || paperSelect.options[paperSelect.selectedIndex]?.text,
  );
  collegeRollSpan.textContent = fallback(student.college_roll_number);
  examRollSpan.textContent = fallback(student.exam_roll_number);

  // Component breakdown table
  assignmentSpan.textContent = twoDigit(student.assignment_marks);
  testSpan.textContent = twoDigit(student.test_marks);
  attendanceMarksSpan.textContent = twoDigit(student.attendance_marks);

  // Total IA marks + progress bar
  const totalIa = twoDigit(student.internal_marks);
  internalsSpan.textContent = totalIa;
  totalIaMarksSpan.textContent = totalIa;
  const iaNum = parseFloat(student.internal_marks);
  if (iaBarEl) iaBarEl.style.width = (Number.isFinite(iaNum) ? Math.min(100, (iaNum / 30) * 100) : 0).toFixed(1) + "%";
  if (iaPercentSubEl) iaPercentSubEl.textContent = Number.isFinite(iaNum) ? Math.round((iaNum / 30) * 100) + "% score" : "—";

  // Attendance % + progress bar
  attendancePercentSpan.textContent = formatPercent(student.attendance_percentage);
  const attNum = parseFloat(student.attendance_percentage);
  const attPct = Number.isFinite(attNum) ? (attNum <= 1 ? attNum * 100 : attNum) : 0;
  if (attBarEl) attBarEl.style.width = Math.min(100, attPct).toFixed(1) + "%";

  // Attendance sub-label
  const attended = twoDigit(student.total_attendance);
  const taken = student.lectures_taken ? ` of ${twoDigit(student.lectures_taken)}` : "";
  totalAttendanceSpan.textContent = attended !== "-" ? `${attended}${taken} lectures attended` : "—";

  // Month-wise attendance table
  attendanceMonthsBody.innerHTML = "";
  const months = student.attendance_months || [];
  if (!months.length) {
    attendanceMonthsBody.innerHTML = '<tr><td colspan="3" class="ms-empty-row">No month-wise attendance available</td></tr>';
  } else {
    months.forEach((item) => {
      const row = document.createElement("tr");
      row.innerHTML = `
        <td><span class="ms-month-pill">${fallback(item.month || item.label)}</span></td>
        <td class="ms-tv">${twoDigit(item.total)}</td>
        <td class="ms-tv">${twoDigit(item.attended)}</td>
      `;
      attendanceMonthsBody.appendChild(row);
    });
  }

  emptyStateDiv.classList.add("hidden");
  studentMarksheetDiv.classList.remove("hidden");
  studentLoginScreen.classList.add("hidden");
  studentDashboard.classList.remove("hidden");
  document.body.classList.remove("student-login-mode");
}

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  setMessage("");
  activeRollNumber = rollInput.value.trim();
  activePaperId = paperSelect.value;
  if (!activePaperId) {
    setMessage("Please choose a paper.", "error");
    return;
  }
  if (!activeRollNumber) {
    setMessage("Please enter exam roll number.", "error");
    return;
  }
  setButtonLoading(sendButton, true, "Sending...");

  try {
    const data = await postJson("/api/request-otp", {
      paperId: activePaperId,
      rollNumber: activeRollNumber,
    });
    loginForm.classList.add("hidden");
    otpForm.classList.remove("hidden");
    otpInput.focus();
    otpStatusSpan.textContent = data.demoOtp
      ? `Email was not sent to ${data.emailMasked}.`
      : `OTP sent to ${data.emailMasked}.`;
    if (data.demoOtp) {
      const reason = data.emailStatus ? ` Reason: ${data.emailStatus}` : "";
      setMessage(`Email was not sent. Demo OTP: ${data.demoOtp}.${reason}`, "error");
    } else {
      setMessage("Please check your spam or junk folder if you do not receive it.", "success");
    }
  } catch (error) {
    setMessage(error.message, "error");
  } finally {
    setButtonLoading(sendButton, false);
  }
});

otpForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  setMessage("");
  setButtonLoading(verifyButton, true, "Verifying...");

  try {
    const data = await postJson("/api/verify-otp", {
      paperId: activePaperId,
      rollNumber: activeRollNumber,
      otp: otpInput.value.trim(),
    });
    showStudent(data.student);
    setMessage("Verified successfully. Marksheet unlocked.", "success");
    otpForm.classList.add("hidden");
    verifiedSessionDiv.classList.remove("hidden");
  } catch (error) {
    setMessage(error.message, "error");
  } finally {
    setButtonLoading(verifyButton, false);
  }
});

function resetToLogin() {
  activeRollNumber = "";
  activePaperId = "";
  otpInput.value = "";
  rollInput.value = "";
  paperSelect.value = "";
  otpForm.classList.add("hidden");
  verifiedSessionDiv.classList.add("hidden");
  loginForm.classList.remove("hidden");
  studentDashboard.classList.add("hidden");
  studentLoginScreen.classList.remove("hidden");
  document.body.classList.add("student-login-mode");
  setMessage("");
  emptyStateDiv.classList.remove("hidden");
  studentMarksheetDiv.classList.add("hidden");
  // Reset progress bars
  if (iaBarEl) iaBarEl.style.width = "0%";
  if (attBarEl) attBarEl.style.width = "0%";
  paperSelect.focus();
}

changeRollButton.addEventListener("click", resetToLogin);
startOverButton.addEventListener("click", resetToLogin);

loadPapers().then(() => paperSelect.focus());
