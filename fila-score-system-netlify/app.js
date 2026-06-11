const headers = [
  "stat_date",
  "style_code",
  "store_id",
  "store_name",
  "product_id",
  "image_url",
  "product_name",
  "category",
  "uv_7d",
  "bounce_rate",
  "ctr",
  "add_cart_rate",
  "conversion_rate",
  "sales_qty_7d",
  "gmv_7d",
  "discount_rate",
  "promo_burst_ratio",
  "unit_price",
  "available_stock",
  "sell_through_rate_30d",
  "positive_rate",
  "return_rate",
];

const fieldLabels = {
  stat_date: "统计日期",
  style_code: "统一款号",
  store_id: "店铺ID",
  store_name: "店铺名称",
  product_id: "商品ID",
  image_url: "商品图片",
  product_name: "商品名称",
  category: "商品类目",
  uv_7d: "7日自然UV",
  bounce_rate: "商详页跳失率",
  ctr: "商品点击率CTR",
  add_cart_rate: "加购率",
  conversion_rate: "成交转化率",
  sales_qty_7d: "7日成交销量",
  gmv_7d: "7日GMV",
  discount_rate: "折扣力度",
  promo_burst_ratio: "大促爆发系数",
  unit_price: "单品客单价",
  available_stock: "当前可售库存",
  sell_through_rate_30d: "30日整体售罄率",
  positive_rate: "好评率",
  return_rate: "整体退货率",
};

const labelToHeader = Object.fromEntries(Object.entries(fieldLabels).map(([key, label]) => [label, key]));

const sampleCsv = `统计日期,统一款号,店铺ID,店铺名称,商品ID,商品图片,商品名称,商品类目,7日自然UV,商详页跳失率,商品点击率CTR,加购率,成交转化率,7日成交销量,7日GMV,折扣力度,大促爆发系数,单品客单价,当前可售库存,30日整体售罄率,好评率,整体退货率
2026-06-10,FILA-001,STORE-A,FILA旗舰店,710001,https://images.unsplash.com/photo-1542291026-7eec264c27ff?w=220,FILA爆款跑鞋,运动鞋,5200,28.00%,5.60%,4.50%,4.00%,700,210000,92.00%,3.20,820,4600,66.00%,96.00%,38.00%
2026-06-10,FILA-001,STORE-B,FILA专卖店,710002,https://images.unsplash.com/photo-1542291026-7eec264c27ff?w=220,FILA爆款跑鞋,运动鞋,3600,29.00%,5.40%,4.30%,3.80%,500,150000,92.00%,3.20,820,3800,66.00%,96.00%,38.00%
2026-06-10,FILA-002,STORE-A,FILA旗舰店,720001,https://images.unsplash.com/photo-1523398002811-999ca8dec234?w=220,FILA潜力卫衣,运动服,1500,42.00%,2.80%,3.20%,2.10%,210,45000,82.00%,,520,800,42.00%,92.00%,45.00%
2026-06-10,FILA-002,STORE-C,FILA奥莱店,720002,https://images.unsplash.com/photo-1523398002811-999ca8dec234?w=220,FILA潜力卫衣,运动服,1100,41.00%,2.70%,3.10%,2.00%,150,31000,82.00%,,520,1000,42.00%,92.00%,45.00%
2026-06-10,FILA-003,STORE-B,FILA专卖店,730001,https://images.unsplash.com/photo-1473966968600-fa801b869a1a?w=220,FILA平销长裤,运动服,1300,49.00%,1.60%,1.80%,1.10%,180,36000,75.00%,1.20,320,700,22.00%,88.00%,58.00%
2026-06-10,FILA-004,STORE-C,FILA奥莱店,740001,https://images.unsplash.com/photo-1521572163474-6864f9cf17ab?w=220,FILA滞销T恤,运动服,420,68.00%,0.70%,0.80%,0.40%,0,0,58.00%,,180,1600,6.00%,82.00%,86.00%`;

let latestResults = [];
let uploadedImages = new Map();
let importMode = "csv";
const expandedStyles = new Set();
let selectedImageStyle = "";

const csvInput = document.getElementById("csvInput");
const fileInput = document.getElementById("fileInput");
const scoreBtn = document.getElementById("scoreBtn");
const imageInput = document.getElementById("imageInput");
const csvModeBtn = document.getElementById("csvModeBtn");
const formModeBtn = document.getElementById("formModeBtn");
const csvPanel = document.getElementById("csvPanel");
const formPanel = document.getElementById("formPanel");
const addRowBtn = document.getElementById("addRowBtn");
const clearRowsBtn = document.getElementById("clearRowsBtn");
const manualHead = document.getElementById("manualHead");
const manualBody = document.getElementById("manualBody");
const loadSampleBtn = document.getElementById("loadSampleBtn");
const downloadTemplateBtn = document.getElementById("downloadTemplateBtn");
const exportBtn = document.getElementById("exportBtn");
const resultBody = document.getElementById("resultBody");
const issues = document.getElementById("issues");
const statusText = document.getElementById("statusText");
const levelFilter = document.getElementById("levelFilter");
const searchInput = document.getElementById("searchInput");
const fieldChips = document.getElementById("fieldChips");

fieldChips.innerHTML = headers.map((field) => `<span>${fieldLabels[field]}</span>`).join("");
renderManualTable();

loadSampleBtn.addEventListener("click", () => {
  csvInput.value = sampleCsv;
  fillManualRows(parseCsv(sampleCsv).rows);
  statusText.textContent = "已载入示例";
});

downloadTemplateBtn.addEventListener("click", () => {
  const link = document.createElement("a");
  link.href = `./fila_score_template.csv?t=${Date.now()}`;
  link.download = "fila_score_template.csv";
  link.target = "_blank";
  document.body.appendChild(link);
  link.click();
  link.remove();
  navigator.clipboard?.writeText(sampleCsv).catch(() => {});
  statusText.textContent = "已打开模板文件；如未自动下载，中文模板和案例已复制";
});

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    if (/\.(xlsx|xls)$/i.test(file.name)) {
      csvInput.value = await readSpreadsheetAsCsv(file);
      fillManualRows(parseCsv(csvInput.value).rows);
      statusText.textContent = `已读取Excel：${file.name}`;
    } else {
      csvInput.value = await file.text();
      fillManualRows(parseCsv(csvInput.value).rows);
      statusText.textContent = `已读取CSV：${file.name}`;
    }
  } catch (error) {
    renderIssues([error.message], true);
    statusText.textContent = "读取失败";
  }
});

imageInput.addEventListener("change", (event) => {
  const files = [...event.target.files].filter((file) => file.type.startsWith("image/"));
  if (!files.length) return;

  uploadedImages.forEach((item) => URL.revokeObjectURL(item.url));
  uploadedImages = new Map();

  files.forEach((file) => {
    const key = normalizeImageKey(file.name.replace(/\.[^.]+$/, ""));
    uploadedImages.set(key, {
      url: URL.createObjectURL(file),
      name: file.name,
    });
  });

  if (latestResults.length) {
    latestResults = latestResults.map((item) => ({
      ...item,
      image_url: resolveImageUrl(item.source_rows, item.image_url),
    }));
    renderResults();
  }

  statusText.textContent = `已上传 ${files.length} 张图片`;
  renderIssues([`图片已上传。文件名需与统一款号或商品ID一致，例如 FILA-001.jpg、710001.png。`], false);
});

scoreBtn.addEventListener("click", () => {
  const parsed = importMode === "csv" ? parseCsv(csvInput.value.trim()) : { rows: collectManualRows(), errors: [] };
  if (parsed.errors.length) {
    renderIssues(parsed.errors, true);
    return;
  }
  const validation = validateRows(parsed.rows);
  if (validation.errors.length) {
    renderIssues(validation.errors, true);
    return;
  }
  latestResults = scoreRows(validation.rows);
  renderIssues(validation.warnings, false);
  renderResults();
  renderSummary();
  exportBtn.disabled = latestResults.length === 0;
  statusText.textContent = `完成 ${latestResults.length} 个款号评分`;
});

csvModeBtn.addEventListener("click", () => setImportMode("csv"));
formModeBtn.addEventListener("click", () => setImportMode("form"));
addRowBtn.addEventListener("click", () => addManualRow({}));
clearRowsBtn.addEventListener("click", () => {
  manualBody.innerHTML = "";
  addManualRow({});
});

exportBtn.addEventListener("click", () => {
  if (!latestResults.length) return;
  const rows = [
    [
      "统一款号",
      "商品图片",
      "商品ID数量",
      "店铺数量",
      "综合得分",
      "商品层级",
      "流量得分",
      "转化得分",
      "销售得分",
      "库存得分",
      "口碑售后得分",
      "运营动作",
    ],
    ...latestResults.map((item) => [
      item.style_code,
      item.image_url,
      item.product_count,
      item.store_count,
      item.final_score.toFixed(2),
      item.product_level,
      item.traffic_score.toFixed(2),
      item.conversion_score.toFixed(2),
      item.sales_score.toFixed(2),
      item.inventory_score.toFixed(2),
      item.reputation_score.toFixed(2),
      item.action_suggestion,
    ]),
  ];
  downloadFile("fila_score_results.csv", rows.map(toCsvLine).join("\n"));
});

levelFilter.addEventListener("change", renderResults);
searchInput.addEventListener("input", renderResults);
resultBody.addEventListener("click", (event) => {
  const imageTarget = event.target.closest("[data-image-target]");
  if (imageTarget) {
    selectedImageStyle = imageTarget.getAttribute("data-image-target");
    statusText.textContent = `已选择 ${selectedImageStyle}，可直接粘贴图片`;
    renderResults();
    return;
  }

  const button = event.target.closest("[data-style-toggle]");
  if (!button) return;
  const styleCode = button.getAttribute("data-style-toggle");
  if (expandedStyles.has(styleCode)) {
    expandedStyles.delete(styleCode);
  } else {
    expandedStyles.add(styleCode);
  }
  renderResults();
});

document.addEventListener("paste", (event) => {
  const imageItem = [...(event.clipboardData?.items || [])].find((item) => item.type.startsWith("image/"));
  if (!imageItem) return;

  const targetStyle = selectedImageStyle || (latestResults.length === 1 ? latestResults[0].style_code : "");
  if (!targetStyle) {
    renderIssues(["请先在评分结果中点击某个款的图片区域，再粘贴图片。"], true);
    return;
  }

  const file = imageItem.getAsFile();
  if (!file) return;

  const key = normalizeImageKey(targetStyle);
  const existing = uploadedImages.get(key);
  if (existing?.url?.startsWith("blob:")) URL.revokeObjectURL(existing.url);
  uploadedImages.set(key, {
    url: URL.createObjectURL(file),
    name: `${targetStyle}-paste.png`,
  });

  latestResults = latestResults.map((item) => ({
    ...item,
    image_url: item.style_code === targetStyle ? uploadedImages.get(key).url : item.image_url,
  }));
  renderResults();
  renderIssues([`${targetStyle} 图片已通过粘贴上传。`], false);
  statusText.textContent = "图片粘贴完成";
});

function parseCsv(text) {
  if (!text) return { rows: [], errors: ["请先粘贴或选择CSV数据。"] };
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  const parsedLines = lines.map(parseCsvLine);
  const inputHeaders = parsedLines[0].map((field) => field.trim());
  const normalizedHeaders = inputHeaders.map((header) => labelToHeader[header] || header);
  const missing = headers.filter((header) => !normalizedHeaders.includes(header));
  if (missing.length) return { rows: [], errors: [`缺少字段：${missing.map((field) => fieldLabels[field]).join(", ")}`] };

  const rows = parsedLines.slice(1).map((values, index) => {
    const row = { _line: index + 2 };
    normalizedHeaders.forEach((header, i) => {
      row[header] = values[i] ?? "";
    });
    return row;
  });
  return { rows, errors: [] };
}

function readSpreadsheetAsCsv(file) {
  return new Promise((resolve, reject) => {
    if (!window.XLSX) {
      reject(new Error("XLSX解析库未加载，请确认网络可访问CDN后刷新页面。"));
      return;
    }

    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Excel文件读取失败。"));
    reader.onload = (event) => {
      try {
        const workbook = window.XLSX.read(event.target.result, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) {
          reject(new Error("Excel文件没有可读取的工作表。"));
          return;
        }
        const sheet = workbook.Sheets[firstSheetName];
        const csv = window.XLSX.utils.sheet_to_csv(sheet, { FS: "," });
        resolve(csv);
      } catch (error) {
        reject(new Error(`Excel解析失败：${error.message}`));
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

function setImportMode(mode) {
  importMode = mode;
  csvModeBtn.classList.toggle("active", mode === "csv");
  formModeBtn.classList.toggle("active", mode === "form");
  csvPanel.classList.toggle("hidden", mode !== "csv");
  formPanel.classList.toggle("hidden", mode !== "form");
  statusText.textContent = mode === "csv" ? "已切换导表模式" : "已切换页面填写模式";
}

function renderManualTable() {
  manualHead.innerHTML = `<tr><th>操作</th>${headers.map((header) => `<th>${fieldLabels[header]}</th>`).join("")}</tr>`;
  manualBody.innerHTML = "";
  addManualRow({});
}

function addManualRow(data = {}) {
  const tr = document.createElement("tr");
  const removeCell = document.createElement("td");
  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "remove-row";
  removeBtn.textContent = "×";
  removeBtn.addEventListener("click", () => {
    tr.remove();
    if (!manualBody.children.length) addManualRow({});
  });
  removeCell.appendChild(removeBtn);
  tr.appendChild(removeCell);

  headers.forEach((header) => {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.name = header;
    input.value = data[header] ?? "";
    if (["product_name", "image_url", "store_name"].includes(header)) input.className = "wide-input";
    td.appendChild(input);
    tr.appendChild(td);
  });

  manualBody.appendChild(tr);
}

function fillManualRows(rows) {
  manualBody.innerHTML = "";
  rows.forEach((row) => addManualRow(row));
}

function collectManualRows() {
  return [...manualBody.querySelectorAll("tr")]
    .map((tr, index) => {
      const row = { _line: index + 1 };
      headers.forEach((header) => {
        row[header] = tr.querySelector(`[name="${header}"]`)?.value ?? "";
      });
      return row;
    })
    .filter((row) => headers.some((header) => String(row[header] || "").trim() !== ""));
}

function parseCsvLine(line) {
  const result = [];
  let current = "";
  let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const next = line[i + 1];
    if (char === '"' && quoted && next === '"') {
      current += '"';
      i += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === "," && !quoted) {
      result.push(current);
      current = "";
    } else {
      current += char;
    }
  }
  result.push(current);
  return result;
}

function validateRows(rows) {
  const errors = [];
  const warnings = [];
  const numericFields = [
    "uv_7d",
    "bounce_rate",
    "ctr",
    "add_cart_rate",
    "conversion_rate",
    "sales_qty_7d",
    "gmv_7d",
    "discount_rate",
    "unit_price",
    "available_stock",
    "sell_through_rate_30d",
    "positive_rate",
    "return_rate",
  ];

  const cleanRows = rows.map((row) => ({ ...row }));
  cleanRows.forEach((row) => {
    ["style_code", "store_id", "store_name", "product_id"].forEach((field) => {
      if (!String(row[field] || "").trim()) {
        errors.push(`第${row._line}行缺少必填字段 ${field}`);
      }
    });

    numericFields.forEach((field) => {
      row[field] = toNumber(row[field]);
      if (Number.isNaN(row[field])) {
        errors.push(`第${row._line}行 ${field} 不是有效数字`);
      }
    });

    row.promo_burst_ratio = String(row.promo_burst_ratio || "").trim() === "" ? null : toNumber(row.promo_burst_ratio);
    if (row.promo_burst_ratio !== null && Number.isNaN(row.promo_burst_ratio)) {
      errors.push(`第${row._line}行 promo_burst_ratio 不是有效数字`);
    }
  });

  const dedupe = new Map();
  cleanRows.forEach((row) => {
    const key = `${row.stat_date}|${row.product_id}`;
    const existing = dedupe.get(key);
    if (existing && existing.style_code !== row.style_code) {
      errors.push(`商品ID ${row.product_id} 同日绑定多个款号`);
    }
    dedupe.set(key, row);
  });

  const excluded = cleanRows.filter((row) => row.uv_7d > 0 && row.add_cart_rate === 0).length;
  if (excluded) {
    warnings.push(`已在加购率汇总中剔除 ${excluded} 条 UV>0 且加购率=0% 的记录。`);
  }

  return { rows: cleanRows, errors, warnings };
}

function scoreRows(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    const key = row.style_code.trim();
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });

  return [...groups.entries()].map(([styleCode, groupRows]) => {
    const productCount = new Set(groupRows.map((row) => row.product_id)).size;
    const storeCount = new Set(groupRows.map((row) => row.store_id)).size;
    const uv = sum(groupRows, "uv_7d");
    const salesQty = sum(groupRows, "sales_qty_7d");
    const gmv = sum(groupRows, "gmv_7d");
    const stockAvg = average(groupRows, "available_stock");
    const dailySales = salesQty / 7;
    const turnoverDays = salesQty === 0 ? null : stockAvg / dailySales;

    const addCartRows = groupRows.filter((row) => !(row.uv_7d > 0 && row.add_cart_rate === 0));
    const promoValues = groupRows.map((row) => row.promo_burst_ratio).filter((value) => value !== null);

    const metrics = {
      uv_7d: uv,
      bounce_rate: weightedAverage(groupRows, "bounce_rate", "uv_7d"),
      ctr: weightedAverage(groupRows, "ctr", "uv_7d"),
      add_cart_rate: addCartRows.length ? weightedAverage(addCartRows, "add_cart_rate", "uv_7d") : 0,
      conversion_rate: weightedAverage(groupRows, "conversion_rate", "uv_7d"),
      sales_qty_7d: salesQty,
      gmv_7d: gmv,
      discount_rate: weightedAverage(groupRows, "discount_rate", "sales_qty_7d"),
      promo_burst_ratio: promoValues.length ? averageValue(promoValues) : null,
      unit_price: weightedAverage(groupRows, "unit_price", "sales_qty_7d"),
      available_stock_avg: stockAvg,
      turnover_days: turnoverDays,
      sell_through_rate_30d: weightedAverage(groupRows, "sell_through_rate_30d", "sales_qty_7d"),
      positive_rate: weightedAverage(groupRows, "positive_rate", "sales_qty_7d"),
      return_rate: weightedAverage(groupRows, "return_rate", "sales_qty_7d"),
    };

    const score = calculateScore(metrics);
    const level = getLevel(score.final);
    return {
      style_code: styleCode,
      product_name: groupRows[0].product_name,
      image_url: resolveImageUrl(groupRows),
      source_rows: groupRows,
      product_count: productCount,
      store_count: storeCount,
      ...metrics,
      traffic_score: score.traffic,
      conversion_score: score.conversion,
      sales_score: score.sales,
      inventory_score: score.inventory,
      reputation_score: score.reputation,
      raw_score: score.raw,
      participation_full_score: score.full,
      final_score: score.final,
      product_level: level.name,
      level_class: level.className,
      action_suggestion: getAction(level.name),
    };
  }).sort((a, b) => b.final_score - a.final_score);
}

function calculateScore(metric) {
  const traffic =
    positive(metric.uv_7d, 8, [5000, 3000, 1000, 500]) +
    negative(metric.bounce_rate, 7, [30, 40, 50, 60]);

  const conversion =
    positive(metric.ctr, 5, [5, 3, 2, 1]) +
    positive(metric.add_cart_rate, 7, [4, 3, 2, 1]) +
    positive(metric.conversion_rate, 8, [3.5, 2.5, 1.5, 0.8]);

  const promoBlank = metric.promo_burst_ratio === null;
  const sales =
    positive(metric.sales_qty_7d, 10, [1000, 500, 300, 100]) +
    positive(metric.gmv_7d, 10, [100000, 50000, 20000, 10000]) +
    discountScore(metric.discount_rate) +
    (promoBlank ? 0 : positive(metric.promo_burst_ratio, 5, [3, 2, 1.5, 1])) +
    positive(metric.unit_price, 5, [800, 600, 400, 200]);

  const inventory =
    metric.sales_qty_7d === 0
      ? 0
      : stockScore(metric.available_stock_avg) +
        negative(metric.turnover_days, 7, [10, 30, 50, 75]) +
        positive(metric.sell_through_rate_30d, 6, [60, 40, 25, 10]);

  const reputation =
    positive(metric.positive_rate, 5, [95, 90, 85, 80]) +
    negative(metric.return_rate, 5, [40, 50, 60, 80]);

  const full = promoBlank ? 95 : 100;
  const raw = traffic + conversion + sales + inventory + reputation;
  return {
    traffic,
    conversion,
    sales,
    inventory,
    reputation,
    raw,
    full,
    final: (raw / full) * 100,
  };
}

function positive(value, max, limits) {
  if (value >= limits[0]) return max;
  if (value >= limits[1]) return max * 0.7;
  if (value >= limits[2]) return max * 0.4;
  if (value >= limits[3]) return max * 0.2;
  return 0;
}

function negative(value, max, limits) {
  if (value <= limits[0]) return max;
  if (value <= limits[1]) return max * 0.7;
  if (value <= limits[2]) return max * 0.4;
  if (value <= limits[3]) return max * 0.2;
  return 0;
}

function discountScore(value) {
  if (value > 90) return 5;
  if (value > 80) return 3.5;
  if (value > 70) return 2;
  if (value > 60) return 1;
  return 0;
}

function stockScore(value) {
  if (value < 300) return 7;
  if (value <= 1000) return 4.9;
  if (value <= 3000) return 2.8;
  if (value <= 5000) return 1.4;
  return 0;
}

function getLevel(score) {
  if (score >= 85) return { name: "顶级爆品", className: "top" };
  if (score >= 70) return { name: "主力爆款", className: "main" };
  if (score >= 55) return { name: "潜力款", className: "potential" };
  if (score >= 40) return { name: "常规平销款", className: "normal" };
  if (score >= 25) return { name: "待优化款", className: "optimize" };
  return { name: "滞销清仓款", className: "clearance" };
}

function getAction(level) {
  const actions = {
    顶级爆品: "保障库存，优先补货，稳定自然流量承接。",
    主力爆款: "加强内容、搜索和活动资源，监控库存波动。",
    潜力款: "优化主图、标题、详情页和短视频承接。",
    常规平销款: "保持基础运营，控制补货，观察转化短板。",
    待优化款: "排查流量、转化、价格、库存和售后问题。",
    滞销清仓款: "降低资源投入，清仓、组合销售或下架处理。",
  };
  return actions[level];
}

function renderResults() {
  const filter = levelFilter.value;
  const keyword = searchInput.value.trim().toLowerCase();
  const rows = latestResults.filter((item) => {
    const levelMatch = !filter || item.product_level === filter;
    const keywordMatch = !keyword || item.style_code.toLowerCase().includes(keyword) || String(item.product_name || "").toLowerCase().includes(keyword);
    return levelMatch && keywordMatch;
  });

  if (!rows.length) {
    resultBody.innerHTML = `<tr><td colspan="12" class="empty">暂无结果</td></tr>`;
    return;
  }

  resultBody.innerHTML = rows.map((item) => `
    <tr>
      <td><button class="image-target" type="button" data-image-target="${escapeHtml(item.style_code)}">${renderThumb(item.image_url, item.style_code)}</button></td>
      <td><button class="style-toggle" type="button" data-style-toggle="${escapeHtml(item.style_code)}">${expandedStyles.has(item.style_code) ? "收起" : "展开"} ${escapeHtml(item.style_code)}</button></td>
      <td>${item.product_count}</td>
      <td>${item.store_count}</td>
      <td class="score">${item.final_score.toFixed(2)}</td>
      <td><span class="level ${item.level_class}">${item.product_level}</span></td>
      <td>${item.traffic_score.toFixed(2)}</td>
      <td>${item.conversion_score.toFixed(2)}</td>
      <td>${item.sales_score.toFixed(2)}</td>
      <td>${item.inventory_score.toFixed(2)}</td>
      <td>${item.reputation_score.toFixed(2)}</td>
      <td>${escapeHtml(item.action_suggestion)}</td>
    </tr>
    ${expandedStyles.has(item.style_code) ? renderDetailRow(item) : ""}
  `).join("");
}

function renderDetailRow(item) {
  const rows = item.source_rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.store_name)}</td>
      <td>${escapeHtml(row.store_id)}</td>
      <td>${escapeHtml(row.product_id)}</td>
      <td>${escapeHtml(row.product_name)}</td>
      <td>${formatNumber(row.uv_7d)}</td>
      <td>${formatNumber(row.sales_qty_7d)}</td>
      <td>${formatNumber(row.gmv_7d)}</td>
      <td>${formatNumber(row.available_stock)}</td>
      <td>${formatNumber(row.ctr)}%</td>
      <td>${formatNumber(row.add_cart_rate)}%</td>
      <td>${formatNumber(row.conversion_rate)}%</td>
      <td>${formatNumber(row.return_rate)}%</td>
    </tr>
  `).join("");

  return `
    <tr class="detail-row">
      <td colspan="12">
        <div class="detail-box">
          <div class="detail-title">店铺商品ID明细</div>
          <table class="detail-table">
            <thead>
              <tr>
                <th>店铺</th>
                <th>店铺ID</th>
                <th>商品ID</th>
                <th>商品名</th>
                <th>UV</th>
                <th>销量</th>
                <th>GMV</th>
                <th>库存</th>
                <th>CTR</th>
                <th>加购率</th>
                <th>转化率</th>
                <th>退货率</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </td>
    </tr>
  `;
}

function formatNumber(value) {
  if (value === null || value === undefined || value === "") return "";
  const number = Number(value);
  if (Number.isNaN(number)) return escapeHtml(value);
  return Number.isInteger(number) ? String(number) : number.toFixed(2);
}

function renderThumb(url, styleCode) {
  const cleaned = String(url || "").trim();
  const selectedClass = selectedImageStyle === styleCode ? " selected-image-target" : "";
  if (!cleaned) return `<span class="thumb-empty${selectedClass}">无图</span>`;
  return `<img class="thumb${selectedClass}" src="${escapeHtml(cleaned)}" alt="商品图" loading="lazy" onerror="this.replaceWith(Object.assign(document.createElement('span'),{className:'thumb-empty',textContent:'失效'}))">`;
}

function resolveImageUrl(groupRows, fallback = "") {
  const styleCode = normalizeImageKey(groupRows[0]?.style_code || "");
  if (uploadedImages.has(styleCode)) return uploadedImages.get(styleCode).url;

  for (const row of groupRows) {
    const productId = normalizeImageKey(row.product_id || "");
    if (uploadedImages.has(productId)) return uploadedImages.get(productId).url;
  }

  const csvUrl = groupRows.find((row) => String(row.image_url || "").trim())?.image_url || "";
  return csvUrl || fallback || "";
}

function normalizeImageKey(value) {
  return String(value || "").trim().toLowerCase();
}

function renderSummary() {
  const total = latestResults.length;
  const avg = total ? latestResults.reduce((acc, item) => acc + item.final_score, 0) / total : 0;
  document.getElementById("styleCount").textContent = total;
  document.getElementById("avgScore").textContent = avg.toFixed(2);
  document.getElementById("levelTop").textContent = countLevel("顶级爆品");
  document.getElementById("levelMain").textContent = countLevel("主力爆款");
  document.getElementById("levelPotential").textContent = countLevel("潜力款");
  document.getElementById("levelRisk").textContent = latestResults.filter((item) => ["待优化款", "滞销清仓款"].includes(item.product_level)).length;
}

function renderIssues(messages, isError) {
  if (!messages.length) {
    issues.innerHTML = "校验通过。";
    return;
  }
  issues.innerHTML = messages.map((message) => `<div class="${isError ? "error" : ""}">${escapeHtml(message)}</div>`).join("");
}

function countLevel(level) {
  return latestResults.filter((item) => item.product_level === level).length;
}

function sum(rows, field) {
  return rows.reduce((acc, row) => acc + row[field], 0);
}

function average(rows, field) {
  return rows.length ? sum(rows, field) / rows.length : 0;
}

function averageValue(values) {
  return values.reduce((acc, value) => acc + value, 0) / values.length;
}

function weightedAverage(rows, field, weightField) {
  if (!rows.length) return 0;
  const weightSum = rows.reduce((acc, row) => acc + Math.max(0, row[weightField]), 0);
  if (weightSum === 0) return average(rows, field);
  return rows.reduce((acc, row) => acc + row[field] * Math.max(0, row[weightField]), 0) / weightSum;
}

function toNumber(value) {
  const cleaned = String(value ?? "").replace(/%/g, "").replace(/,/g, "").trim();
  if (cleaned === "") return NaN;
  return Number(cleaned);
}

function toCsvLine(values) {
  return values.map((value) => {
    const text = String(value ?? "");
    if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
    return text;
  }).join(",");
}

function downloadFile(filename, content) {
  const url = `data:text/csv;charset=utf-8,${encodeURIComponent(`\ufeff${content}`)}`;
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.target = "_self";
  document.body.appendChild(link);
  link.click();
  link.remove();
  return true;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
