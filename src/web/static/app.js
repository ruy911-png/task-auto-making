// 최소 UI 원칙: 진행률 바 없이, 완료 후 성공/실패 상태만 보여준다.

document.querySelectorAll(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll(".panel").forEach((p) => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(`panel-${btn.dataset.tab}`).classList.add("active");
    if (btn.dataset.tab === "run") loadRunPanel();
  });
});

document.querySelectorAll('input[name="matching_method"]').forEach((radio) => {
  radio.addEventListener("change", updateMatchingDetailVisibility);
});

function updateMatchingDetailVisibility() {
  const method = document.querySelector('input[name="matching_method"]:checked').value;
  document.querySelectorAll(".matching-detail").forEach((el) => {
    el.style.display = el.id === `matching-${method}` ? "block" : "none";
  });
}

function addRow(containerId, placeholders) {
  const container = document.getElementById(containerId);
  const row = document.createElement("div");
  row.className = "repeat-row";
  placeholders.forEach((ph) => {
    const input = document.createElement("input");
    input.type = "text";
    input.placeholder = ph;
    row.appendChild(input);
  });
  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "small";
  removeBtn.textContent = "삭제";
  removeBtn.onclick = () => row.remove();
  row.appendChild(removeBtn);
  container.appendChild(row);
}

function rowsToPairs(containerId) {
  return Array.from(document.getElementById(containerId).children).map((row) =>
    Array.from(row.querySelectorAll("input")).map((i) => i.value.trim())
  );
}

function splitCsv(value) {
  return value
    .split(",")
    .map((v) => v.trim())
    .filter((v) => v.length > 0);
}

function buildMatchingPayload() {
  const method = document.querySelector('input[name="matching_method"]:checked').value;
  if (method === "hr_data") {
    return { method };
  }
  if (method === "reference_upload") {
    return {
      method,
      match_column: document.getElementById("match_column").value.trim() || null,
      reference_key_column: document.getElementById("reference_key_column").value.trim() || null,
      reference_name_column: document.getElementById("reference_name_column").value.trim() || "담당자",
      reference_department_column:
        document.getElementById("reference_department_column").value.trim() || null,
    };
  }
  return { method: "sap_lookup" };
}

async function registerControl() {
  const fields = {};
  rowsToPairs("field-mapping-rows").forEach(([col, sapId]) => {
    if (col && sapId) fields[col] = sapId;
  });

  const additionalFields = {};
  rowsToPairs("additional-field-rows").forEach(([col, sapId]) => {
    if (col && sapId) additionalFields[col] = sapId;
  });

  const rename_columns = {};
  rowsToPairs("rename-rows").forEach(([from, to]) => {
    if (from && to) rename_columns[from] = to;
  });

  const filters = rowsToPairs("filter-rows")
    .filter(([col, op, val]) => col && op && val !== "")
    .map(([col, op, val]) => ({ column: col, op, value: isNaN(Number(val)) ? val : Number(val) }));

  const calculated_columns = rowsToPairs("calc-rows")
    .filter(([name, op]) => name && op)
    .map(([name, op, cols, sep]) => ({
      name,
      op,
      columns: splitCsv(cols || ""),
      separator: sep || "",
    }));

  const method = document.querySelector('input[name="matching_method"]:checked').value;

  const payload = {
    control_id: document.getElementById("control_id").value.trim(),
    description: document.getElementById("description").value.trim(),
    transaction: document.getElementById("transaction").value.trim(),
    execute_action: "f8",
    fields,
    additional_screen: {
      open_button_id: document.getElementById("additional_open_button_id").value.trim() || null,
      fields: additionalFields,
      confirm_button_id: document.getElementById("additional_confirm_button_id").value.trim() || null,
    },
    layout: {
      select_button_id: document.getElementById("layout_select_button_id").value.trim() || null,
      layout_name: document.getElementById("layout_name").value.trim() || null,
      layout_name_field_id: document.getElementById("layout_name_field_id").value.trim() || null,
      confirm_button_id: document.getElementById("layout_confirm_button_id").value.trim() || null,
    },
    download: {
      menu_path: splitCsv(document.getElementById("menu_path").value),
      file_path_field_id: document.getElementById("file_path_field_id").value.trim() || null,
      confirm_button_id: document.getElementById("confirm_button_id").value.trim() || null,
    },
    edit_rules: {
      rename_columns,
      drop_columns: splitCsv(document.getElementById("drop_columns").value),
      filters,
      calculated_columns,
    },
    validation: {
      key_columns: splitCsv(document.getElementById("key_columns").value),
      responsible_column:
        method === "hr_data" ? document.getElementById("responsible_column").value.trim() || null : null,
      amount_column: document.getElementById("amount_column").value.trim() || null,
    },
    matching: buildMatchingPayload(),
  };

  const statusBox = document.getElementById("register-status");
  statusBox.style.display = "block";
  statusBox.className = "status-box";
  statusBox.textContent = "등록 중...";

  try {
    const res = await fetch("/api/controls", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "등록 실패");
    statusBox.className = "status-box success";
    statusBox.textContent = `등록 완료: ${data.control_id} (${data.path})`;
  } catch (err) {
    statusBox.className = "status-box failed";
    statusBox.textContent = `실패: ${err.message}`;
  }
}

let loadedControls = [];

async function loadRunPanel() {
  const controlSelect = document.getElementById("run_control_id");
  controlSelect.innerHTML = "";
  loadedControls = await (await fetch("/api/controls")).json();
  loadedControls.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c.control_id;
    opt.textContent = `${c.control_id} - ${c.description || ""}`;
    controlSelect.appendChild(opt);
  });
  controlSelect.onchange = updateRunFileInputs;
  updateRunFileInputs();

  const sessionSelect = document.getElementById("run_session_id");
  sessionSelect.innerHTML = "";
  const hint = document.getElementById("session-hint");
  const sessionData = await (await fetch("/api/sap/sessions")).json();

  if (sessionData.available && sessionData.sessions.length > 0) {
    sessionData.sessions.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s.session_id;
      opt.textContent = `${s.connection_name} (${s.transaction})`;
      sessionSelect.appendChild(opt);
    });
    hint.textContent = "";
  } else {
    hint.textContent = sessionData.message || "SAP 세션을 찾을 수 없습니다.";
  }

  const mockOpt = document.createElement("option");
  mockOpt.value = "__mock__";
  mockOpt.textContent = "__mock__ (SAP 없이 테스트)";
  sessionSelect.appendChild(mockOpt);
}

function updateRunFileInputs() {
  const controlId = document.getElementById("run_control_id").value;
  const control = loadedControls.find((c) => c.control_id === controlId);
  const method = control && control.matching ? control.matching.method : "hr_data";

  document.getElementById("run-hr-file-row").style.display = method === "hr_data" ? "block" : "none";
  document.getElementById("run-reference-file-row").style.display =
    method === "reference_upload" ? "block" : "none";
  document.getElementById("run-sap-lookup-notice").style.display =
    method === "sap_lookup" ? "block" : "none";
  document.getElementById("run-btn").disabled = method === "sap_lookup";
}

let currentJobId = null;

async function startRun() {
  const controlId = document.getElementById("run_control_id").value;
  const control = loadedControls.find((c) => c.control_id === controlId);
  const method = control && control.matching ? control.matching.method : "hr_data";

  const conditionFile = document.getElementById("condition_file").files[0];
  const hrFile = document.getElementById("hr_file").files[0];
  const referenceFile = document.getElementById("reference_file").files[0];
  const statusBox = document.getElementById("run-status");
  const actions = document.getElementById("run-actions");
  actions.style.display = "none";

  if (!conditionFile) {
    statusBox.style.display = "block";
    statusBox.className = "status-box failed";
    statusBox.textContent = "조건 엑셀을 업로드하세요.";
    return;
  }
  if (method === "hr_data" && !hrFile) {
    statusBox.style.display = "block";
    statusBox.className = "status-box failed";
    statusBox.textContent = "인사데이터 엑셀을 업로드하세요.";
    return;
  }
  if (method === "reference_upload" && !referenceFile) {
    statusBox.style.display = "block";
    statusBox.className = "status-box failed";
    statusBox.textContent = "참조 엑셀을 업로드하세요.";
    return;
  }

  const form = new FormData();
  form.append("control_id", controlId);
  form.append("session_id", document.getElementById("run_session_id").value);
  form.append("condition_file", conditionFile);
  if (method === "hr_data") form.append("hr_file", hrFile);
  if (method === "reference_upload") form.append("reference_file", referenceFile);

  statusBox.style.display = "block";
  statusBox.className = "status-box";
  statusBox.textContent = "실행 중...";
  document.getElementById("run-btn").disabled = true;

  try {
    const res = await fetch("/api/runs", { method: "POST", body: form });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || "실행 실패");
    currentJobId = data.job_id;
    renderSummary(data.summary);
    actions.style.display = "flex";
  } catch (err) {
    statusBox.className = "status-box failed";
    statusBox.textContent = `실패: ${err.message}`;
  } finally {
    document.getElementById("run-btn").disabled = false;
  }
}

function renderSummary(summary) {
  const statusBox = document.getElementById("run-status");
  statusBox.className = summary.failed > 0 ? "status-box failed" : "status-box success";
  statusBox.textContent = `총 ${summary.total}건 중 성공 ${summary.success}건, 실패 ${summary.failed}건`;
}

async function retryFailed() {
  if (!currentJobId) return;
  const statusBox = document.getElementById("run-status");
  statusBox.textContent = "실패 건 재처리 중...";
  const res = await fetch(`/api/runs/${currentJobId}/retry`, { method: "POST" });
  const data = await res.json();
  if (res.ok) renderSummary(data.summary);
}

function downloadResult() {
  if (!currentJobId) return;
  window.location.href = `/api/runs/${currentJobId}/download`;
}
