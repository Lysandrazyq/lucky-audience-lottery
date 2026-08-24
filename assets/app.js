"use strict";

const fileInput = document.querySelector("#file-input");
const dropZone = document.querySelector("#drop-zone");
const manualList = document.querySelector("#manual-list");
const applyManualButton = document.querySelector("#apply-manual");
const loadSampleButton = document.querySelector("#load-sample");
const clearListButton = document.querySelector("#clear-list");
const fileSummary = document.querySelector("#file-summary");
const fileSummaryText = document.querySelector("#file-summary-text");
const participantCount = document.querySelector("#participant-count");
const winnerCount = document.querySelector("#winner-count");
const excludeWinners = document.querySelector("#exclude-winners");
const resetHistoryButton = document.querySelector("#reset-history");
const drawButton = document.querySelector("#draw-button");
const results = document.querySelector("#results");
const status = document.querySelector("#status");

let participants = [];
let winnerKeys = new Set();

const sample = `张三,20260001
李四,20260002
王五,20260003
赵六,20260004
陈同学,20260005
刘同学,20260006
周同学,20260007
吴同学,20260008
郑同学,20260009
黄同学,20260010`;

function setStatus(message) {
  status.textContent = message;
}

function normalize(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function makeKey(person, index) {
  return person.id ? `id:${person.id}` : `name:${person.name}:${index}`;
}

function setParticipants(items, sourceLabel, skipped = 0) {
  const seenIds = new Set();
  participants = items.reduce((all, item) => {
    const name = normalize(item.name);
    const id = normalize(item.id);
    if (!name || (id && seenIds.has(id))) return all;
    if (id) seenIds.add(id);
    all.push({ name, id });
    return all;
  }, []).map((person, index) => ({ ...person, key: makeKey(person, index) }));

  winnerKeys = new Set();
  updateInterface();
  renderEmptyResults();

  const ignored = skipped + items.length - participants.length;
  const ignoredText = ignored > 0 ? `，已忽略 ${ignored} 条空白或重复记录` : "";
  const message = participants.length ? `已导入 ${participants.length} 位观众${ignoredText}` : "未找到有效的姓名，请检查名单格式";
  setStatus(message);

  if (sourceLabel) {
    fileSummary.hidden = false;
    fileSummaryText.textContent = `${sourceLabel} · ${participants.length} 位有效观众`;
  }
}

function updateInterface() {
  const remaining = participants.filter((person) => !winnerKeys.has(person.key)).length;
  participantCount.textContent = participants.length
    ? `共 ${participants.length} 位观众${winnerKeys.size ? `，剩余 ${remaining} 位未中奖` : ""}`
    : "请先导入名单";

  winnerCount.max = String(participants.length || 10);
  if (participants.length && Number(winnerCount.value) > Number(winnerCount.max)) {
    winnerCount.value = winnerCount.max;
  }

  const eligibleCount = excludeWinners.checked ? remaining : participants.length;
  drawButton.disabled = eligibleCount < 1;
  resetHistoryButton.disabled = winnerKeys.size === 0;
}

function renderEmptyResults(message = "导入名单后，点击开始抽奖") {
  results.replaceChildren();
  const empty = document.createElement("div");
  empty.className = "empty-results";
  empty.innerHTML = '<span class="empty-icon" aria-hidden="true">✦</span>';
  const text = document.createElement("p");
  text.textContent = message;
  empty.append(text);
  results.append(empty);
}

function renderWinners(winners) {
  results.replaceChildren();
  winners.forEach((winner, index) => {
    const card = document.createElement("article");
    card.className = "winner-card";
    card.style.animationDelay = `${index * 130}ms`;

    const rank = document.createElement("span");
    rank.className = "winner-rank";
    rank.textContent = `幸运观众 ${index + 1}`;

    const name = document.createElement("strong");
    name.className = "winner-name";
    name.textContent = winner.name;

    const id = document.createElement("span");
    id.className = "winner-id";
    id.textContent = winner.id ? `学号：${winner.id}` : "";

    card.append(rank, name, id);
    results.append(card);
  });
}

function parseManualList(value) {
  return value.split(/\r?\n/).map((line) => {
    const [name, id = ""] = line.split(/[,，\t]/, 2);
    return { name, id };
  }).filter((item) => normalize(item.name));
}

function normalizedHeader(value) {
  return normalize(value).toLowerCase().replace(/[\s_\-()（）【】\[\]：:]/g, "");
}

function findColumn(headers, terms) {
  return headers.findIndex((header) => terms.some((term) => header === term || header.includes(term)));
}

function parseSheetRows(rows) {
  const nonEmptyRows = rows.filter((row) => row.some((cell) => normalize(cell)));
  if (!nonEmptyRows.length) return { people: [], skipped: 0 };

  const firstRow = nonEmptyRows[0];
  const headers = firstRow.map(normalizedHeader);
  const nameColumn = findColumn(headers, ["姓名", "名字", "name", "观众", "参会人", "参与者"]);
  const idColumn = findColumn(headers, ["学号", "工号", "编号", "人员编号", "学员编号", "id"]);
  const hasHeader = nameColumn !== -1;
  const dataRows = hasHeader ? nonEmptyRows.slice(1) : nonEmptyRows;
  const nameIndex = hasHeader ? nameColumn : 0;
  const idIndex = hasHeader && idColumn !== -1 ? idColumn : 1;

  const people = [];
  let skipped = 0;
  dataRows.forEach((row) => {
    const name = normalize(row[nameIndex]);
    const id = normalize(row[idIndex]);
    if (name) people.push({ name, id });
    else skipped += 1;
  });
  return { people, skipped };
}

async function importFile(file) {
  if (!file) return;
  if (!window.XLSX) {
    setStatus("Excel 解析组件未加载，请确认网络连接后刷新页面再试。");
    return;
  }

  try {
    setStatus(`正在读取 ${file.name}…`);
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", raw: false });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) throw new Error("文件中没有工作表");

    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
    const { people, skipped } = parseSheetRows(rows);
    manualList.value = people.map((person) => `${person.name}${person.id ? `,${person.id}` : ""}`).join("\n");
    setParticipants(people, `${file.name}（工作表：${sheetName}）`, skipped);
  } catch (error) {
    console.error(error);
    setStatus(`无法导入文件：${error.message || "请确认文件格式正确"}`);
  } finally {
    fileInput.value = "";
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
  const pool = excludeWinners.checked
    ? participants.filter((person) => !winnerKeys.has(person.key))
    : participants;
  const amount = Math.min(Math.max(1, Number(winnerCount.value) || 1), pool.length);

  if (!pool.length) {
    setStatus("没有可参与抽奖的观众。可重置中奖记录后继续。 ");
    return;
  }

  drawButton.disabled = true;
  drawButton.textContent = "抽取中…";
  setStatus("正在随机抽取幸运观众…");

  window.setTimeout(() => {
    const winners = shuffle(pool).slice(0, amount);
    if (excludeWinners.checked) winners.forEach((winner) => winnerKeys.add(winner.key));
    renderWinners(winners);
    updateInterface();
    drawButton.textContent = "开始下一轮抽奖";
    setStatus(`抽奖完成！本轮共抽取 ${winners.length} 位幸运观众。`);
  }, 420);
}

fileInput.addEventListener("change", () => importFile(fileInput.files[0]));
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
dropZone.addEventListener("drop", (event) => importFile(event.dataTransfer.files[0]));

applyManualButton.addEventListener("click", () => {
  fileSummary.hidden = true;
  setParticipants(parseManualList(manualList.value), "手动粘贴的名单");
});
loadSampleButton.addEventListener("click", () => {
  manualList.value = sample;
  fileSummary.hidden = true;
  setParticipants(parseManualList(sample), "示例名单");
});
clearListButton.addEventListener("click", () => {
  manualList.value = "";
  participants = [];
  winnerKeys = new Set();
  fileSummary.hidden = true;
  updateInterface();
  renderEmptyResults();
  setStatus("名单已清空");
  drawButton.textContent = "开始抽奖";
});
excludeWinners.addEventListener("change", updateInterface);
winnerCount.addEventListener("change", () => {
  if (Number(winnerCount.value) < 1) winnerCount.value = "1";
  updateInterface();
});
resetHistoryButton.addEventListener("click", () => {
  winnerKeys = new Set();
  updateInterface();
  renderEmptyResults("中奖记录已重置，可以重新开始抽奖");
  setStatus("中奖记录已重置");
  drawButton.textContent = "开始抽奖";
});
drawButton.addEventListener("click", drawWinners);

updateInterface();
