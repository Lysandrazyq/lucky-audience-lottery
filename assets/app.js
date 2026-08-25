"use strict";

const GROUPS = [
  { key: "junior1", label: "初一", shortLabel: "初一" },
  { key: "junior2", label: "初二", shortLabel: "初二" },
  { key: "junior3", label: "初三＋四高一", shortLabel: "初三＋四高一" },
  { key: "senior1", label: "高一", shortLabel: "高一" },
];
const WINNERS_PER_GROUP = 3;

const fileInput = document.querySelector("#file-input");
const dropZone = document.querySelector("#drop-zone");
const fileSummary = document.querySelector("#file-summary");
const fileSummaryText = document.querySelector("#file-summary-text");
const groupSummary = document.querySelector("#group-summary");
const participantCount = document.querySelector("#participant-count");
const excludeWinners = document.querySelector("#exclude-winners");
const resetHistoryButton = document.querySelector("#reset-history");
const copyResultButton = document.querySelector("#copy-result");
const drawButton = document.querySelector("#draw-button");
const results = document.querySelector("#results");
const status = document.querySelector("#status");

const importedFiles = new Map();
const pools = Object.fromEntries(GROUPS.map((group) => [group.key, []]));
let winnerKeys = new Set();
let lastWinners = [];

function setStatus(message) {
  status.textContent = message;
}

function normalize(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function classifyFileName(fileName) {
  const name = normalize(fileName);
  if (/四高一|初三/.test(name)) return "junior3";
  if (/初二/.test(name)) return "junior2";
  if (/初一/.test(name)) return "junior1";
  if (/高一/.test(name)) return "senior1";
  return "";
}

function normalizedHeader(value) {
  return normalize(value).toLowerCase().replace(/[\s_\-()（）【】\[\]：:]/g, "");
}

function findColumn(headers, terms) {
  return headers.findIndex((header) => terms.some((term) => header === term || header.includes(term)));
}

function parseSheetRows(rows) {
  const nonEmptyRows = rows.filter((row) => row.some((cell) => normalize(cell)));
  if (!nonEmptyRows.length) return { people: [], skipped: 0, hasHeader: false, hasIdHeader: false };

  const headers = nonEmptyRows[0].map(normalizedHeader);
  const nameColumn = findColumn(headers, ["姓名", "名字", "学生名单", "学生姓名", "名单", "name", "观众", "参会人", "参与者"]);
  const idColumn = findColumn(headers, ["学号", "工号", "编号", "人员编号", "学员编号", "id"]);
  const hasHeader = nameColumn !== -1;
  const hasIdHeader = idColumn !== -1;
  const dataRows = hasHeader ? nonEmptyRows.slice(1) : nonEmptyRows;
  const nameIndex = hasHeader ? nameColumn : 0;
  const idIndex = hasHeader && hasIdHeader ? idColumn : 1;

  const people = [];
  let skipped = 0;
  dataRows.forEach((row) => {
    const name = normalize(row[nameIndex]);
    const id = normalize(row[idIndex]);
    if (name) people.push({ name, id });
    else skipped += 1;
  });
  return { people, skipped, hasHeader, hasIdHeader };
}

function decodeCsvBuffer(buffer) {
  const bytes = new Uint8Array(buffer);
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("gb18030").decode(bytes);
  }
}

function selectRosterSheet(workbook) {
  let best = null;
  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
    const parsed = parseSheetRows(rows);
    const idCount = parsed.people.filter((person) => person.id).length;
    const score = (parsed.hasHeader ? 10000 : 0) + (parsed.hasIdHeader ? 5000 : 0) + idCount * 10 + parsed.people.length;
    if (!best || score > best.score) best = { ...parsed, sheetName, score };
  });
  return best || { people: [], skipped: 0, sheetName: "" };
}

function dedupePeople(items, groupKey) {
  const seen = new Set();
  return items.reduce((all, item, index) => {
    const name = normalize(item.name);
    const id = normalize(item.id);
    const dedupeKey = id ? `id:${id}` : `name:${name}`;
    if (!name || seen.has(dedupeKey)) return all;
    seen.add(dedupeKey);
    all.push({ name, id, key: `${groupKey}:${dedupeKey}:${index}` });
    return all;
  }, []);
}

function rebuildPools() {
  GROUPS.forEach((group) => {
    const records = [...importedFiles.values()]
      .filter((file) => file.groupKey === group.key)
      .flatMap((file) => file.people);
    pools[group.key] = dedupePeople(records, group.key);
  });
  winnerKeys = new Set();
  lastWinners = [];
  copyResultButton.disabled = true;
  drawButton.textContent = "开始抽奖";
  updateInterface();
  renderPlaceholders();
}

function updateGroupSummary() {
  groupSummary.replaceChildren();
  GROUPS.forEach((group) => {
    const item = document.createElement("div");
    item.className = `group-count${pools[group.key].length ? " loaded" : ""}`;
    const label = document.createElement("span");
    label.textContent = group.label;
    const count = document.createElement("strong");
    count.textContent = pools[group.key].length ? `${pools[group.key].length} 人` : "未导入";
    item.append(label, count);
    groupSummary.append(item);
  });
}

function updateInterface() {
  const total = GROUPS.reduce((sum, group) => sum + pools[group.key].length, 0);
  const loadedGroups = GROUPS.filter((group) => pools[group.key].length).length;
  const remaining = GROUPS.reduce((sum, group) => (
    sum + pools[group.key].filter((person) => !winnerKeys.has(person.key)).length
  ), 0);
  participantCount.textContent = total
    ? `${loadedGroups} 组 · 共 ${total} 位观众${winnerKeys.size ? ` · 剩余 ${remaining} 位未中奖` : ""}`
    : "请先一次选择 5 份名单";

  const everyGroupReady = GROUPS.every((group) => {
    const groupPool = excludeWinners.checked
      ? pools[group.key].filter((person) => !winnerKeys.has(person.key))
      : pools[group.key];
    return groupPool.length >= WINNERS_PER_GROUP;
  });
  drawButton.disabled = !everyGroupReady;
  resetHistoryButton.disabled = winnerKeys.size === 0;
  updateGroupSummary();
}

function renderPlaceholders(message = "每组抽取 3 人") {
  results.replaceChildren();
  GROUPS.forEach((group) => {
    const card = document.createElement("article");
    card.className = "winner-card placeholder";
    const label = document.createElement("span");
    label.className = "winner-rank";
    label.textContent = group.label;
    const name = document.createElement("strong");
    name.className = "winner-name";
    name.textContent = message;
    const count = document.createElement("span");
    count.className = "winner-id";
    count.textContent = pools[group.key].length ? `${pools[group.key].length} 位观众` : "请导入对应名单";
    card.append(label, name, count);
    results.append(card);
  });
}

function renderWinners(winners) {
  results.replaceChildren();
  winners.forEach((winner, index) => {
    const card = document.createElement("article");
    card.className = "winner-card";
    card.style.animationDelay = `${index * 110}ms`;

    const rank = document.createElement("span");
    rank.className = "winner-rank";
    rank.textContent = winner.groupLabel;

    const list = document.createElement("div");
    list.className = "winner-list";
    winner.members.forEach((person, personIndex) => {
      const row = document.createElement("div");
      row.className = "winner-person";
      const number = document.createElement("span");
      number.className = "winner-number";
      number.textContent = String(personIndex + 1);
      const personText = document.createElement("div");
      const name = document.createElement("strong");
      name.className = "winner-person-name";
      name.textContent = person.name;
      const id = document.createElement("span");
      id.className = "winner-person-id";
      id.textContent = person.id ? `学号：${person.id}` : "";
      personText.append(name, id);
      row.append(number, personText);
      list.append(row);
    });

    card.append(rank, list);
    results.append(card);
  });
}

async function importFiles(fileList) {
  const files = [...fileList];
  if (!files.length) return;
  if (!window.XLSX) {
    setStatus("Excel 解析组件未正确加载，请刷新页面后重试。");
    return;
  }

  drawButton.disabled = true;
  setStatus(`正在读取 ${files.length} 个文件…`);
  const unsupported = [];
  const failed = [];
  let loaded = 0;

  for (const file of files) {
    const groupKey = classifyFileName(file.name);
    if (!groupKey) {
      unsupported.push(file.name);
      continue;
    }
    try {
      const buffer = await file.arrayBuffer();
      const isCsv = /\.csv$/i.test(file.name);
      const workbook = isCsv
        ? XLSX.read(decodeCsvBuffer(buffer), { type: "string", raw: false })
        : XLSX.read(buffer, { type: "array", raw: false });
      const roster = selectRosterSheet(workbook);
      if (!roster.people.length) throw new Error("未找到有效名单");
      importedFiles.set(file.name, {
        groupKey,
        people: roster.people,
        skipped: roster.skipped,
        sheetName: roster.sheetName,
      });
      loaded += 1;
    } catch (error) {
      console.error(error);
      failed.push(file.name);
    }
  }

  fileInput.value = "";
  if (loaded) rebuildPools();
  const total = GROUPS.reduce((sum, group) => sum + pools[group.key].length, 0);
  const missing = GROUPS.filter((group) => !pools[group.key].length).map((group) => group.label);
  fileSummary.hidden = total === 0;
  fileSummaryText.textContent = `已读取 ${importedFiles.size} 份名单 · 共 ${total} 位观众`;

  if (failed.length || unsupported.length) {
    setStatus(`部分文件未导入：${[...failed, ...unsupported].join("、")}`);
  } else if (missing.length) {
    setStatus(`还需导入：${missing.join("、")}`);
  } else {
    setStatus("四组名单已就绪，初三与四高一已合并。");
  }
}

function shuffle(array) {
  const copy = [...array];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function drawWinners() {
  const winners = [];
  for (const group of GROUPS) {
    const pool = excludeWinners.checked
      ? pools[group.key].filter((person) => !winnerKeys.has(person.key))
      : pools[group.key];
    if (pool.length < WINNERS_PER_GROUP) {
      setStatus(`${group.label} 剩余人数不足 ${WINNERS_PER_GROUP} 人，请重置中奖记录或重新导入。`);
      return;
    }
    winners.push({
      groupKey: group.key,
      groupLabel: group.label,
      members: shuffle(pool).slice(0, WINNERS_PER_GROUP),
    });
  }

  drawButton.disabled = true;
  drawButton.textContent = "抽取中…";
  setStatus("正在从四组名单中随机抽取…");

  window.setTimeout(() => {
    if (excludeWinners.checked) {
      winners.forEach((winner) => winner.members.forEach((person) => winnerKeys.add(person.key)));
    }
    lastWinners = winners;
    renderWinners(winners);
    copyResultButton.disabled = false;
    updateInterface();
    drawButton.textContent = "开始下一轮抽奖";
    setStatus("抽奖完成！每组各抽取 3 位，共 12 位幸运观众。");
  }, 420);
}

function roundedRect(context, x, y, width, height, radius) {
  context.beginPath();
  context.roundRect(x, y, width, height, radius);
  context.fill();
  context.stroke();
}

function drawFittedText(context, text, centerX, baselineY, maxWidth, startSize, minSize) {
  let size = startSize;
  do {
    context.font = `700 ${size}px "Microsoft YaHei", "PingFang SC", sans-serif`;
    if (context.measureText(text).width <= maxWidth) break;
    size -= 2;
  } while (size > minSize);
  context.fillText(text, centerX, baselineY);
}

function createResultCanvas() {
  const canvas = document.createElement("canvas");
  canvas.width = 1600;
  canvas.height = 900;
  const context = canvas.getContext("2d");
  const background = context.createLinearGradient(0, 0, 1600, 900);
  background.addColorStop(0, "#f8faff");
  background.addColorStop(0.55, "#eef4ff");
  background.addColorStop(1, "#faf7ff");
  context.fillStyle = background;
  context.fillRect(0, 0, 1600, 900);

  context.textAlign = "center";
  context.fillStyle = "#172033";
  context.font = '700 64px "Microsoft YaHei", "PingFang SC", sans-serif';
  context.fillText("幸运观众抽奖", 800, 145);
  context.fillStyle = "#6d7890";
  context.font = '400 25px "Microsoft YaHei", "PingFang SC", sans-serif';
  context.fillText("本轮抽取结果", 800, 192);

  const cardWidth = 340;
  const cardHeight = 440;
  const gap = 28;
  const startX = (1600 - cardWidth * 4 - gap * 3) / 2;
  const cardY = 270;

  lastWinners.forEach((winner, index) => {
    const x = startX + index * (cardWidth + gap);
    context.fillStyle = "#ffffff";
    context.strokeStyle = "#aac0ff";
    context.lineWidth = 2;
    context.shadowColor = "rgba(49, 95, 211, 0.12)";
    context.shadowBlur = 22;
    context.shadowOffsetY = 10;
    roundedRect(context, x, cardY, cardWidth, cardHeight, 30);
    context.shadowColor = "transparent";

    context.fillStyle = "#315fd3";
    context.font = '700 28px "Microsoft YaHei", "PingFang SC", sans-serif';
    context.fillText(winner.groupLabel, x + cardWidth / 2, cardY + 62);

    winner.members.forEach((person, personIndex) => {
      const nameY = cardY + 145 + personIndex * 105;
      context.fillStyle = "#1e377b";
      drawFittedText(context, person.name, x + cardWidth / 2, nameY, cardWidth - 50, 34, 24);
      context.fillStyle = "#72809a";
      context.font = '400 17px "Microsoft YaHei", "PingFang SC", sans-serif';
      context.fillText(person.id ? `学号：${person.id}` : "", x + cardWidth / 2, nameY + 31);
    });
  });

  context.fillStyle = "#8b96aa";
  context.font = '400 20px "Microsoft YaHei", "PingFang SC", sans-serif';
  context.fillText("初一 · 初二 · 初三＋四高一 · 高一", 800, 795);
  return canvas;
}

function downloadCanvas(canvas) {
  const link = document.createElement("a");
  link.download = `幸运观众抽奖结果-${new Date().toISOString().slice(0, 10)}.png`;
  link.href = canvas.toDataURL("image/png");
  link.click();
}

async function copyResultImage() {
  if (!lastWinners.length) return;
  const canvas = createResultCanvas();
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  try {
    if (!navigator.clipboard?.write || !window.ClipboardItem) throw new Error("Clipboard image is unavailable");
    await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
    setStatus("结果图片已复制，可直接粘贴到 WPS/PPT。");
  } catch {
    downloadCanvas(canvas);
    setStatus("浏览器不支持复制图片，已自动下载 PNG，可插入 WPS/PPT。");
  }
}

fileInput.addEventListener("change", () => importFiles(fileInput.files));
["dragenter", "dragover"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
});
["dragleave", "drop"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
});
dropZone.addEventListener("drop", (event) => importFiles(event.dataTransfer.files));
excludeWinners.addEventListener("change", updateInterface);
resetHistoryButton.addEventListener("click", () => {
  winnerKeys = new Set();
  lastWinners = [];
  copyResultButton.disabled = true;
  updateInterface();
  renderPlaceholders("等待抽取");
  setStatus("中奖记录已重置");
  drawButton.textContent = "开始抽奖";
});
copyResultButton.addEventListener("click", copyResultImage);
drawButton.addEventListener("click", drawWinners);

updateInterface();
renderPlaceholders();
