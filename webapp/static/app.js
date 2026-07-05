const fileInput = document.getElementById("fileInput");
const fileList = document.getElementById("fileList");
const generateBtn = document.getElementById("generateBtn");
const clearBtn = document.getElementById("clearBtn");
const dropzone = document.getElementById("dropzone");
const statusText = document.getElementById("statusText");
const asinCount = document.getElementById("asinCount");
const moduleCount = document.getElementById("moduleCount");
const downloadLink = document.getElementById("downloadLink");
const directions = document.getElementById("directions");
const moduleTags = document.getElementById("moduleTags");
const modulePreview = document.getElementById("modulePreview");
const asinRows = document.getElementById("asinRows");
const tableHint = document.getElementById("tableHint");
const historyList = document.getElementById("historyList");
const importerInput = document.getElementById("importerInput");
const categoryInput = document.getElementById("categoryInput");
const styleInput = document.getElementById("styleInput");
const reportModeInput = document.getElementById("reportModeInput");
const reportTypeInput = document.getElementById("reportTypeInput");
const llmEnabledInput = document.getElementById("llmEnabledInput");
const llmBaseUrlInput = document.getElementById("llmBaseUrlInput");
const llmModelInput = document.getElementById("llmModelInput");
const llmApiKeyInput = document.getElementById("llmApiKeyInput");
const llmProtocolInput = document.getElementById("llmProtocolInput");
const historyDateFilter = document.getElementById("historyDateFilter");
const historyCategoryFilter = document.getElementById("historyCategoryFilter");

let selectedFiles = [];
let historyItems = [];

const moduleOrder = [
  "01_竞对好评差评总览",
  "02_VOC全量分析_ASIN好评差评",
  "03_评论证据样例",
  "04_未归类观点检查",
  "05_各ASIN需求层级摘要",
  "06_ASIN码数订单评论",
  "07_尺码订单统计表",
  "08_ASIN颜色订单占比",
  "09_颜色统计表",
  "10_需求汇总统计",
  "11_使用场景汇总",
  "12_各ASIN使用场景明细表",
  "13_用户画像",
  "14_竞对横向数据对比表",
  "TOP5好评",
  "TOP5差评",
];

function setStatus(text) {
  statusText.textContent = text;
}

function renderFiles() {
  fileList.innerHTML = "";
  if (!selectedFiles.length) {
    fileList.innerHTML = '<li class="empty">还没有选择文件</li>';
    generateBtn.disabled = true;
    return;
  }
  selectedFiles.forEach((file) => {
    const li = document.createElement("li");
    li.textContent = `${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
    fileList.appendChild(li);
  });
  generateBtn.disabled = false;
}

function addFiles(files) {
  const existing = new Set(selectedFiles.map((file) => `${file.name}-${file.size}`));
  Array.from(files).forEach((file) => {
    const key = `${file.name}-${file.size}`;
    if (!existing.has(key)) selectedFiles.push(file);
  });
  renderFiles();
}

fileInput.addEventListener("change", (event) => addFiles(event.target.files));
clearBtn.addEventListener("click", () => {
  selectedFiles = [];
  fileInput.value = "";
  renderFiles();
});

["dragenter", "dragover"].forEach((name) => {
  dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    dropzone.classList.add("drag");
  });
});

["dragleave", "drop"].forEach((name) => {
  dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    dropzone.classList.remove("drag");
  });
});

dropzone.addEventListener("drop", (event) => addFiles(event.dataTransfer.files));

function renderResult(result) {
  asinCount.textContent = result.asinCount || 0;
  moduleCount.textContent = (result.modules || []).length;
  downloadLink.href = result.downloadUrl;
  downloadLink.classList.remove("disabled");

  directions.innerHTML = "";
  (result.directions || []).forEach((item) => {
    const li = document.createElement("li");
    li.textContent = item;
    directions.appendChild(li);
  });

  moduleTags.innerHTML = "";
  (result.modules || []).forEach((item) => {
    const link = document.createElement("a");
    link.textContent = item;
    link.href = (result.moduleDownloads || {})[item] || "#";
    moduleTags.appendChild(link);
  });

  renderModules(result.moduleData || {}, result.moduleDownloads || {});
  renderAsinRows(result.records || []);
  if (result.llm && result.llm.message) {
    setStatus(`${result.llm.message}${result.llm.model ? `｜${result.llm.model}` : ""}`);
  }
  loadHistory();
}

function renderAsinRows(records) {
  asinRows.innerHTML = "";
  records.forEach((record) => {
    const tr = document.createElement("tr");
    const cells = [
      record["竞对 ASIN"] || record["ASIN"],
      record["品牌/角色"] || record["品牌"],
      record["日均销量"] ? Number(record["日均销量"]).toFixed(1) : (record["订单尺码Top"] || ""),
      record["结论2痛点定位"] || record["VOC分析"] || record["用户画像"] || record["核心关注的点"] || "",
      record["优先级"] || record["RICE/优先级"] || "",
    ];
    cells.forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value || "";
      tr.appendChild(td);
    });
    asinRows.appendChild(tr);
  });
  if (!records.length) {
    asinRows.innerHTML = '<tr><td colspan="5">暂无数据</td></tr>';
  }
  tableHint.textContent = `已生成 ${records.length} 个 ASIN`;
}

function renderModules(moduleData, downloads) {
  modulePreview.innerHTML = "";
  const names = [...moduleOrder, ...Object.keys(moduleData).filter((name) => !moduleOrder.includes(name))];
  names.forEach((name) => {
    const rows = moduleData[name] || [];
    const card = document.createElement("div");
    card.className = "module-card";

    const header = document.createElement("header");
    const title = document.createElement("strong");
    title.textContent = `${name} (${rows.length})`;
    const link = document.createElement("a");
    link.href = downloads[name] || "#";
    link.textContent = "下载";
    header.appendChild(title);
    header.appendChild(link);
    card.appendChild(header);

    const table = document.createElement("table");
    table.className = "mini-table";
    if (!rows.length) {
      table.innerHTML = "<tbody><tr><td>暂无数据</td></tr></tbody>";
    } else {
      const headers = Object.keys(rows[0]).slice(0, 5);
      const thead = document.createElement("thead");
      const trh = document.createElement("tr");
      headers.forEach((h) => {
        const th = document.createElement("th");
        th.textContent = h;
        trh.appendChild(th);
      });
      thead.appendChild(trh);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      rows.slice(0, 5).forEach((row) => {
        const tr = document.createElement("tr");
        headers.forEach((h) => {
          const td = document.createElement("td");
          const value = row[h];
          td.textContent = typeof value === "number" ? value.toFixed(value < 1 ? 2 : 0) : value || "";
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
    }
    card.appendChild(table);
    modulePreview.appendChild(card);
  });
}

function renderHistory(items) {
  historyList.innerHTML = "";
  const dateValue = historyDateFilter.value;
  const categoryValue = historyCategoryFilter.value.trim().toLowerCase();
  const filtered = items.filter((item) => {
    const dateOk = !dateValue || String(item.createdAt || "").startsWith(dateValue);
    const categoryOk = !categoryValue || String(item.productCategory || "").toLowerCase().includes(categoryValue);
    return dateOk && categoryOk;
  });

  if (!filtered.length) {
    historyList.innerHTML = '<div class="empty-state">暂无导入记录</div>';
    return;
  }

  filtered.slice(0, 20).forEach((item) => {
    const card = document.createElement("div");
    card.className = "history-item";

    const main = document.createElement("div");
    const title = document.createElement("strong");
    title.textContent = `${item.createdAt || ""} · ${item.asinCount || 0} 个 ASIN · ${item.importer || "未记录导入人"}`;
    const meta = document.createElement("span");
    const fileNames = (item.files || []).map((file) => file.name).join("、");
    const asins = (item.asins || []).join("、");
    meta.textContent = `分类：${item.productCategory || "未填"}｜款号：${item.styleNumber || "未填"}｜${fileNames || "未记录文件"}${asins ? "｜" + asins : ""}`;
    main.appendChild(title);
    main.appendChild(meta);

    const link = document.createElement("a");
    link.href = item.downloadUrl || "#";
    link.textContent = "下载报告";
    card.appendChild(main);
    card.appendChild(link);
    historyList.appendChild(card);
  });
}

async function loadHistory() {
  try {
    const response = await fetch("/api/history");
    const payload = await response.json();
    historyItems = payload.items || [];
    renderHistory(historyItems);
  } catch (error) {
    renderHistory([]);
  }
}

[historyDateFilter, historyCategoryFilter].forEach((input) => {
  input.addEventListener("input", () => renderHistory(historyItems));
});

generateBtn.addEventListener("click", async () => {
  if (!selectedFiles.length) return;
  if (!importerInput.value.trim()) {
    setStatus("请先填写导入人");
    importerInput.focus();
    return;
  }

  generateBtn.disabled = true;
  downloadLink.classList.add("disabled");
  setStatus("正在上传和生成");

  const form = new FormData();
  selectedFiles.forEach((file) => form.append("files", file));
  form.append("importer", importerInput.value.trim());
  form.append("productCategory", categoryInput.value.trim());
  form.append("styleNumber", styleInput.value.trim());
  form.append("reportMode", reportModeInput.value);
  form.append("reportType", reportTypeInput.value);
  form.append("llmEnabled", llmEnabledInput.checked ? "true" : "false");
  form.append("llmBaseUrl", llmBaseUrlInput.value.trim());
  form.append("llmModel", llmModelInput.value.trim());
  form.append("llmApiKey", llmApiKeyInput.value.trim());
  form.append("llmProtocol", llmProtocolInput.value);

  try {
    const response = await fetch("/api/generate", { method: "POST", body: form });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "生成失败");
    renderResult(payload);
    setStatus(payload.llm?.message || "生成完成");
  } catch (error) {
    setStatus(error.message);
  } finally {
    generateBtn.disabled = false;
  }
});

renderFiles();
loadHistory();
