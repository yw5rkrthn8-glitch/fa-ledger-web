const form = document.getElementById("record-form");
const recordsBody = document.getElementById("records-body");
const incomeTotalEl = document.getElementById("income-total");
const expenseTotalEl = document.getElementById("expense-total");
const balanceTotalEl = document.getElementById("balance-total");
const selectExcelBtn = document.getElementById("select-excel-btn");
const excelFileInput = document.getElementById("excel-file");
const loadExcelBtn = document.getElementById("load-excel-btn");
const saveExcelBtn = document.getElementById("save-excel-btn");
const excelFileNameEl = document.getElementById("excel-file-name");
const saveHintEl = document.getElementById("save-hint");
const monthlyStatsEl = document.getElementById("monthly-stats");
const categoryStatsEl = document.getElementById("category-stats");
let selectedExcelHandle = null;
const API_BASE_URL = String(window.APP_CONFIG?.API_BASE_URL || "").replace(/\/$/, "");

function apiUrl(path) {
  if (!API_BASE_URL) return path;
  return `${API_BASE_URL}${path}`;
}

async function fetchRecords() {
  const res = await fetch(apiUrl("/api/records"));
  if (!res.ok) throw new Error("获取记录失败");
  return res.json();
}

async function createRecord(payload) {
  const res = await fetch(apiUrl("/api/records"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.message || "保存失败");
  }
  return res.json();
}

async function deleteRecord(id) {
  const res = await fetch(apiUrl(`/api/records/${id}`), { method: "DELETE" });
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.message || "删除失败");
  }
}

async function loadExcel(base64) {
  const res = await fetch(apiUrl("/api/load-excel"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ base64 })
  });
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.message || "Excel 导入失败");
  }
  return res.json();
}

async function saveExcel() {
  const res = await fetch(apiUrl("/api/save-excel"));
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.message || "Excel 导出失败");
  }
  return res.blob();
}

function toBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      const commaIndex = result.indexOf(",");
      if (commaIndex < 0) {
        reject(new Error("文件读取失败"));
        return;
      }
      resolve(result.slice(commaIndex + 1));
    };
    reader.onerror = () => reject(new Error("文件读取失败"));
    reader.readAsDataURL(file);
  });
}

function supportsFileSystemAccess() {
  return typeof window.showOpenFilePicker === "function";
}

function setSaveHint() {
  const backendLabel = API_BASE_URL || "当前站点同域 API";
  if (supportsFileSystemAccess()) {
    saveHintEl.textContent = `提示：选中文件后可直接写回该 Excel。后端：${backendLabel}`;
  } else {
    saveHintEl.textContent =
      `提示：当前浏览器不支持直接写回文件，Save Excel 将下载文件。后端：${backendLabel}`;
  }
}

function downloadBlob(blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `ledger-${Date.now()}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function writeBlobToSelectedFile(blob) {
  if (!selectedExcelHandle) return false;
  const writable = await selectedExcelHandle.createWritable();
  await writable.write(blob);
  await writable.close();
  return true;
}

async function syncSelectedExcelFile() {
  if (!selectedExcelHandle) return;
  const blob = await saveExcel();
  await writeBlobToSelectedFile(blob);
}

function updateSummary(records) {
  const income = records
    .filter((r) => r.type === "收入")
    .reduce((sum, r) => sum + Number(r.amount || 0), 0);
  const expense = records
    .filter((r) => r.type === "支出")
    .reduce((sum, r) => sum + Number(r.amount || 0), 0);
  const balance = income - expense;

  incomeTotalEl.textContent = `收入：${income.toFixed(2)}`;
  expenseTotalEl.textContent = `支出：${expense.toFixed(2)}`;
  balanceTotalEl.textContent = `结余：${balance.toFixed(2)}`;
}

function renderMonthlyStats(records) {
  const monthMap = new Map();
  for (const item of records) {
    const monthKey = String(item.date || "").slice(0, 7) || "未填写日期";
    if (!monthMap.has(monthKey)) {
      monthMap.set(monthKey, { income: 0, expense: 0 });
    }
    const bucket = monthMap.get(monthKey);
    if (item.type === "收入") {
      bucket.income += Number(item.amount || 0);
    } else {
      bucket.expense += Number(item.amount || 0);
    }
  }

  const rows = Array.from(monthMap.entries()).sort((a, b) => b[0].localeCompare(a[0]));
  monthlyStatsEl.innerHTML = "";

  if (!rows.length) {
    monthlyStatsEl.textContent = "暂无数据";
    return;
  }

  for (const [month, data] of rows) {
    const total = data.income + data.expense;
    const incomeWidth = total ? (data.income / total) * 100 : 0;
    const expenseWidth = total ? (data.expense / total) * 100 : 0;
    const item = document.createElement("div");
    item.className = "stat-item";
    item.innerHTML = `
      <div class="stat-title">${month}</div>
      <div class="bar-row"><div class="bar-income" style="width:${incomeWidth.toFixed(2)}%"></div></div>
      <div class="bar-row"><div class="bar-expense" style="width:${expenseWidth.toFixed(2)}%"></div></div>
      <div class="stat-meta">收入 ${data.income.toFixed(2)} / 支出 ${data.expense.toFixed(2)} / 结余 ${(data.income - data.expense).toFixed(2)}</div>
    `;
    monthlyStatsEl.appendChild(item);
  }
}

function renderCategoryStats(records) {
  const categoryMap = new Map();
  for (const item of records) {
    if (item.type !== "支出") continue;
    const category = item.category || "未分类";
    categoryMap.set(category, (categoryMap.get(category) || 0) + Number(item.amount || 0));
  }

  const rows = Array.from(categoryMap.entries()).sort((a, b) => b[1] - a[1]);
  categoryStatsEl.innerHTML = "";

  if (!rows.length) {
    categoryStatsEl.textContent = "暂无支出数据";
    return;
  }

  const max = rows[0][1] || 1;
  for (const [category, amount] of rows) {
    const width = (amount / max) * 100;
    const item = document.createElement("div");
    item.className = "stat-item";
    item.innerHTML = `
      <div class="stat-title">${category}</div>
      <div class="bar-row"><div class="bar-expense" style="width:${width.toFixed(2)}%"></div></div>
      <div class="stat-meta">支出 ${amount.toFixed(2)}</div>
    `;
    categoryStatsEl.appendChild(item);
  }
}

function renderRecords(records) {
  recordsBody.innerHTML = "";

  for (const item of records) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${item.date || ""}</td>
      <td>${item.type || ""}</td>
      <td>${item.category || ""}</td>
      <td>${Number(item.amount || 0).toFixed(2)}</td>
      <td>${item.note || ""}</td>
      <td><button class="delete-btn" data-id="${item.id}">删除</button></td>
    `;
    recordsBody.appendChild(tr);
  }

  updateSummary(records);
  renderMonthlyStats(records);
  renderCategoryStats(records);
}

async function loadAndRender() {
  try {
    const records = await fetchRecords();
    records.sort((a, b) => new Date(b.date) - new Date(a.date));
    renderRecords(records);
  } catch (error) {
    alert(error.message);
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const payload = {
    date: document.getElementById("date").value,
    type: document.getElementById("type").value,
    category: document.getElementById("category").value.trim(),
    amount: document.getElementById("amount").value,
    note: document.getElementById("note").value.trim()
  };

  try {
    await createRecord(payload);
    form.reset();
    await loadAndRender();
    await syncSelectedExcelFile();
  } catch (error) {
    alert(error.message);
  }
});

recordsBody.addEventListener("click", async (event) => {
  if (!event.target.matches(".delete-btn")) return;
  const id = event.target.getAttribute("data-id");
  if (!id) return;

  if (!window.confirm("确定删除这条记录吗？")) return;
  try {
    await deleteRecord(id);
    await loadAndRender();
  } catch (error) {
    alert(error.message);
  }
});

selectExcelBtn.addEventListener("click", async () => {
  if (!supportsFileSystemAccess()) {
    alert("当前浏览器不支持直接写回文件，请使用下方文件选择框。");
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [
        {
          description: "Excel Files",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
            "application/vnd.ms-excel": [".xls"]
          }
        }
      ]
    });

    if (!handle) return;
    selectedExcelHandle = handle;
    excelFileNameEl.textContent = `已选择：${handle.name}`;
  } catch (error) {
    if (error && error.name === "AbortError") return;
    alert("选择文件失败");
  }
});

loadExcelBtn.addEventListener("click", async () => {
  try {
    let file = null;
    if (selectedExcelHandle) {
      file = await selectedExcelHandle.getFile();
    } else {
      file = excelFileInput.files && excelFileInput.files[0];
    }

    if (!file) {
      alert("请先选择 Excel 文件");
      return;
    }

    const base64 = await toBase64(file);
    const result = await loadExcel(base64);
    await loadAndRender();
    alert(`${result.message}，共 ${result.count} 条`);
  } catch (error) {
    alert(error.message);
  }
});

saveExcelBtn.addEventListener("click", async () => {
  try {
    const blob = await saveExcel();
    const wrote = await writeBlobToSelectedFile(blob);
    if (wrote) {
      alert("已保存到已选择的 Excel 文件");
      return;
    }
    downloadBlob(blob);
  } catch (error) {
    alert(error.message);
  }
});

loadAndRender();
setSaveHint();
