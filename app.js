// Rank Score Calculator (Multi-file / Multi-sheet)
// - Import multiple .txt files (each file = one group)
// - Calculate each group independently
// - Export ONE .xlsx with multiple sheets (one sheet per group)
// - Supports opponent tokens like: 6 OR W6/D6/L6 (W/D/L shown in Excel, ignored in calculation)
//
// IMPORTANT: Excel styling requires xlsx-js-style in index.html
// <script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>

let groups = []; // [{ fileName: string, sheetBase: string, players: string[][] }]

let roundCount = 4;
let fullMark = 2;
let outputFileName = "output";

// === UI reset hook (index.html calls this) ===
window.clearImportedGroups = function () {
  groups = [];
};

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("Failed to read file"));
    reader.readAsText(file);
  });
}

function sanitizeSheetName(name) {
  // Excel sheet name rules: max 31 chars, cannot contain: : \ / ? * [ ]
  const cleaned = String(name)
    .replace(/\.[^/.]+$/, "") // remove extension
    .replace(/[:\\/?*\[\]]/g, "_")
    .trim();
  return (cleaned || "Group").slice(0, 31);
}

function makeUniqueSheetName(baseName, existing) {
  let name = baseName.slice(0, 31);
  if (!existing.has(name)) {
    existing.add(name);
    return name;
  }

  // Add suffix _2, _3 ... and keep within 31 chars
  let i = 2;
  while (true) {
    const suffix = `_${i}`;
    const trimmedBase = baseName.slice(0, Math.max(0, 31 - suffix.length));
    const candidate = `${trimmedBase}${suffix}`;
    if (!existing.has(candidate)) {
      existing.add(candidate);
      return candidate;
    }
    i++;
  }
}

function isValidPlayerRow(row) {
  // Expect at least: name, rank, k, ...
  if (!row || row.length < 4) return false;
  const rank = parseInt(row[1], 10);
  const k = parseInt(row[2], 10);
  return Number.isFinite(rank) && Number.isFinite(k);
}

function parsePlayersFromText(text) {
  const lines = String(text)
    .replace(/\r/g, "")
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);

  return lines
    .map((line) => line.split(",").map((x) => x.trim()))
    .filter(isValidPlayerRow);
}

function parseOpponentToken(token) {
  if (token === undefined || token === null) {
    return { opponentId: null, display: "" };
  }

  const raw = String(token).trim();
  if (!raw) {
    return { opponentId: null, display: "" };
  }

  const first = raw[0].toUpperCase();
  const maybeNumber =
    first === "W" || first === "D" || first === "L" ? raw.slice(1) : raw;

  const opponentId = parseInt(maybeNumber, 10);
  return {
    opponentId: Number.isFinite(opponentId) ? opponentId : null,
    display: raw,
  };
}

// ===== Import Handler (multi-file) =====
document.getElementById("importFile").addEventListener("change", async function (e) {
  const files = Array.from(e.target.files || []);
  groups = [];

  if (files.length === 0) return;

  try {
    const texts = await Promise.all(files.map(readFileAsText));
    groups = texts.map((txt, idx) => {
      const file = files[idx];
      const players = parsePlayersFromText(txt);
      return {
        fileName: file.name,
        sheetBase: sanitizeSheetName(file.name),
        players,
      };
    });
  } catch (err) {
    console.error(err);
    alert("Failed to read one of the imported files. Please try again.");
    groups = [];
  }
});

// ===== Calculation Entry =====
function calculate() {
  roundCount = parseInt(document.getElementById("rounds").value, 10);
  fullMark = parseInt(document.getElementById("markPerRound").value, 10) * roundCount;
  outputFileName = document.getElementById("fileName").value || "output";

  if (!groups.length) {
    alert("Please import at least one valid .txt file");
    return;
  }

  // Build one workbook with multiple sheets
  const workbook = XLSX.utils.book_new();
  const usedSheetNames = new Set();

  // Also build one update text file (compatible with your Update Excel tool)
  const updateLines = [];

  for (const group of groups) {
    if (!group.players || group.players.length === 0) continue;

    const { wsData, meta, groupUpdateLines } = calculateGroup(group.players, group.sheetBase);
    updateLines.push(...groupUpdateLines);

    const worksheet = XLSX.utils.aoa_to_sheet(wsData);

    applySheetFormatting(worksheet, {
      sheetTitle: meta.sheetTitle,
      groupName: meta.groupName,
      roundCount: meta.roundCount,
      totalCols: meta.totalCols,
      headerRowIndex: meta.headerRowIndex,
      dataStartRowIndex: meta.dataStartRowIndex,
      dataRowCount: meta.dataRowCount,
      avgOppCol: meta.avgOppCol,
      expectedCol: meta.expectedCol,
      roundsStartCol: meta.roundsStartCol,
      changeCol: meta.changeCol,
      legendRowIndex: meta.legendRowIndex,
      groupRowIndex: meta.groupRowIndex,
      kTableStartRow: meta.kTableStartRow,
    });

    const sheetName = makeUniqueSheetName(group.sheetBase, usedSheetNames);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  }

  if (workbook.SheetNames.length === 0) {
    alert("No valid player rows found in the imported files.");
    return;
  }

  // Export workbook
  XLSX.writeFile(workbook, `${outputFileName}.xlsx`);

  // Export update text file (optional but keeps old workflow working)
  if (updateLines.length > 0) {
    downloadText(`${outputFileName}.txt`, updateLines.join("\n"));
  }

  document.getElementById("resultMsg").innerText = "✅ Calculation completed. Files ready for download.";

  // Reset file list so user can start a new operation (keep result message)
  if (typeof window.resetRankImportUI === "function") {
    window.resetRankImportUI({ clearResult: false });
  }
}

// ===== Group Calculation & Sheet Layout =====
function calculateGroup(players, groupName) {
  const updatedPlayers = players.map((player, index) => {
    const name = player[0];
    const rank = parseInt(player[1], 10);
    const k = parseInt(player[2], 10);
    const score = parseFloat(player[3 + roundCount]);

    // Calculate opponent average (ignore W/D/L, use only opponent id)
    let total = 0;
    let empty = 0;

    for (let i = 0; i < roundCount; i++) {
      const { opponentId } = parseOpponentToken(player[3 + i]);
      const opponentIndex = opponentId ? opponentId - 1 : -1;

      if (opponentIndex >= 0 && players[opponentIndex]) {
        total += parseInt(players[opponentIndex][1], 10);
      } else {
        empty++;
      }
    }

    const divisor = roundCount - empty;
    const avgOpponent = divisor > 0 ? Math.ceil(total / divisor) : 0;

    const expected = Number(getExpectedScore(rank, avgOpponent)).toFixed(1);
    const change = (score - parseFloat(expected)) * k;
    const finalRank = Math.round(rank + change);

    return [
      index + 1,
      name,
      rank,
      k,
      ...player.slice(3, 3 + roundCount), // keep W/D/L token for display
      score,
      avgOpponent,
      expected,
      change.toFixed(1),
      finalRank,
    ];
  });

  // Layout indices (0-based)
  // Requirement:
  // 1) Legend row is 2 rows before header
  // 2) Group name row is between legend and header
  // 3) K-table sits above 平均对手等级分 & 期望分, and its last row aligns with Legend row
  const totalCols = 9 + roundCount;
  const headerRowIndex = 8;
  const legendRowIndex = headerRowIndex - 2;
  const groupRowIndex = headerRowIndex - 1;

  const dataStartRowIndex = headerRowIndex + 1;
  const dataRowCount = updatedPlayers.length;

  const roundsStartCol = 4;
  const avgOppCol = 5 + roundCount;     // 平均对手等级分
  const expectedCol = 6 + roundCount;   // 期望分
  const changeCol = 7 + roundCount;     // 变化

  // K-table has 5 rows, last row aligns with legendRowIndex
  const kTableStartRow = legendRowIndex - 4;

  // Build sheet array with empty rows up to the end of data
  const wsData = [];
  const totalRows = dataStartRowIndex + dataRowCount;
  for (let r = 0; r < totalRows; r++) {
    wsData.push(Array.from({ length: totalCols }, () => ""));
  }

  // Title row
  wsData[0][0] = `${groupName} 等级分比赛`;

  // Legend row (2 rows before header)
  wsData[legendRowIndex][roundsStartCol] = "W=WIN";
  wsData[legendRowIndex][roundsStartCol + 1] = "D=DRAW";
  wsData[legendRowIndex][roundsStartCol + 2] = "L=LOSE";

  // Group name row (between legend and header)
  wsData[groupRowIndex][roundsStartCol + 1] = groupName;

  // K-table above 平均对手等级分 & 期望分 (two columns: avgOppCol and expectedCol)
  // last row aligns with legendRowIndex
  wsData[kTableStartRow][avgOppCol] = "等级分";
  wsData[kTableStartRow][expectedCol] = "K值";

  wsData[kTableStartRow + 1][avgOppCol] = "2000或以上";
  wsData[kTableStartRow + 1][expectedCol] = 10;

  wsData[kTableStartRow + 2][avgOppCol] = "1700-1999";
  wsData[kTableStartRow + 2][expectedCol] = 15;

  wsData[kTableStartRow + 3][avgOppCol] = "1550-1699";
  wsData[kTableStartRow + 3][expectedCol] = 20;

  wsData[kTableStartRow + 4][avgOppCol] = "1549或以下";
  wsData[kTableStartRow + 4][expectedCol] = 30;

  // Header row (ONLY 等级分 column colored; others no fill)
  wsData[headerRowIndex] = [
    "编号",
    "棋手",
    "等级分",
    "K值",
    ...Array.from({ length: roundCount }, (_, i) => `第${i + 1}轮`),
    "总得分",
    "平均对手等级分",
    "期望分",
    "变化",
    "最终等级分",
  ];

  // Data rows
  for (let i = 0; i < updatedPlayers.length; i++) {
    wsData[dataStartRowIndex + i] = updatedPlayers[i];
  }

  // Update text lines: name, oldRank, newRank
  const groupUpdateLines = updatedPlayers.map((row) => {
    const name = row[1];
    const oldRank = row[2];
    const newRank = row[row.length - 1];
    return `${name},${oldRank},${newRank}`;
  });

  return {
    wsData,
    meta: {
      sheetTitle: `${groupName} 等级分比赛`,
      groupName,
      roundCount,
      totalCols,
      headerRowIndex,
      legendRowIndex,
      groupRowIndex,
      dataStartRowIndex,
      dataRowCount,
      avgOppCol,
      expectedCol,
      roundsStartCol,
      changeCol,
      kTableStartRow,
    },
    groupUpdateLines,
  };
}

function downloadText(filename, content) {
  const blob = new Blob([content], { type: "text/plain" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}

// ===== Expected Score (existing logic) =====
function getExpectedScore(playerMark, oppMark) {
  const diff = playerMark - oppMark;
  const absDiff = Math.abs(diff);
  const breakpoints = [
    [0, 3, 0.5],
    [4, 10, diff >= 0 ? 0.51 : 0.49],
    [11, 17, diff >= 0 ? 0.52 : 0.48],
    [18, 25, diff >= 0 ? 0.53 : 0.47],
    [26, 32, diff >= 0 ? 0.54 : 0.46],
    [33, 39, diff >= 0 ? 0.55 : 0.45],
    [40, 46, diff >= 0 ? 0.56 : 0.44],
    [47, 53, diff >= 0 ? 0.57 : 0.43],
    [54, 61, diff >= 0 ? 0.58 : 0.42],
    [62, 68, diff >= 0 ? 0.59 : 0.41],
    [69, 76, diff >= 0 ? 0.6 : 0.4],
    [77, 83, diff >= 0 ? 0.61 : 0.39],
    [84, 91, diff >= 0 ? 0.62 : 0.38],
    [92, 98, diff >= 0 ? 0.63 : 0.37],
    [99, 106, diff >= 0 ? 0.64 : 0.36],
    [107, 113, diff >= 0 ? 0.65 : 0.35],
    [114, 121, diff >= 0 ? 0.66 : 0.34],
    [122, 129, diff >= 0 ? 0.67 : 0.33],
    [130, 137, diff >= 0 ? 0.68 : 0.32],
    [138, 145, diff >= 0 ? 0.69 : 0.31],
    [146, 153, diff >= 0 ? 0.7 : 0.3],
    [154, 162, diff >= 0 ? 0.71 : 0.29],
    [163, 170, diff >= 0 ? 0.72 : 0.28],
    [171, 179, diff >= 0 ? 0.73 : 0.27],
    [180, 188, diff >= 0 ? 0.74 : 0.26],
    [189, 197, diff >= 0 ? 0.75 : 0.25],
    [198, 206, diff >= 0 ? 0.76 : 0.24],
    [207, 215, diff >= 0 ? 0.77 : 0.23],
    [216, 225, diff >= 0 ? 0.78 : 0.22],
    [226, 235, diff >= 0 ? 0.79 : 0.21],
    [236, 245, diff >= 0 ? 0.8 : 0.2],
    [246, 256, diff >= 0 ? 0.81 : 0.19],
    [257, 267, diff >= 0 ? 0.82 : 0.18],
    [268, 278, diff >= 0 ? 0.83 : 0.17],
    [279, 290, diff >= 0 ? 0.84 : 0.16],
    [291, 302, diff >= 0 ? 0.85 : 0.15],
    [303, 315, diff >= 0 ? 0.86 : 0.14],
    [316, 328, diff >= 0 ? 0.87 : 0.13],
    [329, 344, diff >= 0 ? 0.88 : 0.12],
    [345, 357, diff >= 0 ? 0.89 : 0.11],
    [358, 374, diff >= 0 ? 0.9 : 0.1],
    [375, 391, diff >= 0 ? 0.91 : 0.09],
    [392, 411, diff >= 0 ? 0.92 : 0.08],
    [412, 432, diff >= 0 ? 0.93 : 0.07],
    [433, 456, diff >= 0 ? 0.94 : 0.06],
    [457, 484, diff >= 0 ? 0.95 : 0.05],
    [485, 517, diff >= 0 ? 0.96 : 0.04],
    [518, 559, diff >= 0 ? 0.97 : 0.03],
    [560, 619, diff >= 0 ? 0.98 : 0.02],
    [620, 734, diff >= 0 ? 0.99 : 0.01],
  ];

  for (const [min, max, val] of breakpoints) {
    if (absDiff >= min && absDiff <= max) return fullMark * val;
  }

  return diff >= 0 ? Number(fullMark).toFixed(1) : 0;
}

// ===== Excel Styling Helpers (xlsx-js-style) =====
function applySheetFormatting(ws, opts) {
  const {
    sheetTitle,
    groupName,
    roundCount,
    totalCols,
    headerRowIndex,
    dataStartRowIndex,
    dataRowCount,
    avgOppCol,
    expectedCol,
    roundsStartCol,
    changeCol,
    legendRowIndex,
    groupRowIndex,
    kTableStartRow,
  } = opts;

  // Row heights (make header row same as data rows)
  ws["!rows"] = ws["!rows"] || [];
  ws["!rows"][0] = { hpt: 24 }; // title row (optional, nice)
  ws["!rows"][headerRowIndex] = { hpt: 15 }; // header row = normal height

  // Make data rows consistent too (optional but recommended)
  for (let r = dataStartRowIndex; r < dataStartRowIndex + dataRowCount; r++) {
    ws["!rows"][r] = { hpt: 15 };
  }

  ws["!merges"] = ws["!merges"] || [];

  // Column widths
  const widths = [];
  widths.push({ wch: 6 });   // 编号
  widths.push({ wch: 12 });  // 棋手
  widths.push({ wch: 8 });   // 等级分
  widths.push({ wch: 6 });   // K值
  for (let i = 0; i < roundCount; i++) widths.push({ wch: 8 });
  widths.push({ wch: 8 });   // 总得分
  widths.push({ wch: 16 });  // 平均对手等级分
  widths.push({ wch: 8 });   // 期望分
  widths.push({ wch: 8 });   // 变化
  widths.push({ wch: 10 });  // 最终等级分
  ws["!cols"] = widths;

  const thinBorder = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } },
  };

  const baseCell = {
    font: { name: "Calibri", sz: 11 },
    alignment: { vertical: "center", horizontal: "center", wrapText: true },
    border: thinBorder,
  };

  const titleStyle = {
    font: { name: "Calibri", sz: 18, bold: true },
    alignment: { vertical: "center", horizontal: "left" },
  };

  const headerStyle = {
    font: { name: "Calibri", sz: 11}, //, bold: true 
    alignment: {
      vertical: "center",
      horizontal: "center",
      wrapText: false,
      shrinkToFit: true,   // keeps row height small even for long headers
    },
    border: thinBorder,
  };


  // Only 等级分 header filled
  const ratingHeaderFill = { patternType: "solid", fgColor: { rgb: "F4B183" } };

  // Legend fills
  const legendWFill = { patternType: "solid", fgColor: { rgb: "FFD966" } };
  const legendDFill = { patternType: "solid", fgColor: { rgb: "C6E0B4" } };
  const legendLFill = { patternType: "solid", fgColor: { rgb: "9DC3E6" } };

  // K table header fill
  const kHeaderFill = { patternType: "solid", fgColor: { rgb: "D9D9D9" } };

  function addr(r, c) {
    return XLSX.utils.encode_cell({ r, c });
  }

  function setCell(r, c, style) {
    const a = addr(r, c);
    if (!ws[a]) ws[a] = { t: "s", v: "" };
    ws[a].s = { ...(ws[a].s || {}), ...style };
  }

  // Title row merge across full table width
  ws["!merges"].push({
    s: { r: 0, c: 0 },
    e: { r: 0, c: totalCols - 1 },
  });
  setCell(0, 0, titleStyle);

  // Legend row colors (2 rows before header)
  setCell(legendRowIndex, roundsStartCol,     { ...baseCell, font: { bold: true }, fill: legendWFill });
  setCell(legendRowIndex, roundsStartCol + 1, { ...baseCell, font: { bold: true }, fill: legendDFill });
  setCell(legendRowIndex, roundsStartCol + 2, { ...baseCell, font: { bold: true }, fill: legendLFill });

  // Group name between legend and header (merge across 3 cells)
  ws["!merges"].push({
    s: { r: groupRowIndex, c: roundsStartCol },
    e: { r: groupRowIndex, c: roundsStartCol + 2 },
  });

  // Keep group row clean (no need borders), just centered bold text
  setCell(groupRowIndex, roundsStartCol, {
    font: { name: "Calibri", sz: 12, bold: true },
    alignment: { vertical: "center", horizontal: "center" },
  });
  for (let c = roundsStartCol; c <= roundsStartCol + 2; c++) {
    setCell(groupRowIndex, c, { alignment: { vertical: "center", horizontal: "center" } });
  }

  // K-table formatting (with borders), last row aligns with legend row
  for (let r = kTableStartRow; r <= legendRowIndex; r++) {
    for (let c = avgOppCol; c <= expectedCol; c++) {
      const isHeader = r === kTableStartRow;
      setCell(r, c, {
        ...baseCell,
        font: { name: "Calibri", sz: 11, bold: isHeader },
        fill: isHeader ? kHeaderFill : undefined,
      });
    }
  }

  // Header row: ONLY 等级分 filled, all header cells bold + thin border
  for (let c = 0; c < totalCols; c++) {
    const style = { ...headerStyle };
    if (c === 2) style.fill = ratingHeaderFill;
    setCell(headerRowIndex, c, style);
  }

  // Data rows: thin borders, round colors, change font colors
  for (let r = dataStartRowIndex; r < dataStartRowIndex + dataRowCount; r++) {
    for (let c = 0; c < totalCols; c++) {
      setCell(r, c, baseCell);

      // Round coloring by W/D/L
      if (c >= roundsStartCol && c < roundsStartCol + roundCount) {
        const a = addr(r, c);
        const v = ws[a] ? String(ws[a].v || "").trim() : "";
        const first = v ? v[0].toUpperCase() : "";
        if (first === "W") setCell(r, c, { ...baseCell, fill: legendWFill });
        else if (first === "D") setCell(r, c, { ...baseCell, fill: legendDFill });
        else if (first === "L") setCell(r, c, { ...baseCell, fill: legendLFill });
      }

      // 变化 column: green positive, red negative
      if (c === changeCol) {
        const a = addr(r, c);
        const num = ws[a] ? parseFloat(ws[a].v) : NaN;
        if (!Number.isNaN(num)) {
          if (num > 0) setCell(r, c, { ...baseCell, font: { color: { rgb: "008000" }, bold: true } });
          else if (num < 0) setCell(r, c, { ...baseCell, font: { color: { rgb: "C00000" }, bold: true } });
        }
      }
    }
  }

  // No bold outside border (Requirement #5) -> we only use thin borders everywhere.

  // Ensure title text
  const titleAddr = addr(0, 0);
  if (ws[titleAddr]) ws[titleAddr].v = sheetTitle;

  // Ensure group text
  ws[addr(groupRowIndex, roundsStartCol)].v = groupName;
}

function clearForm() {
  document.getElementById("fileName").value = "";
  document.getElementById("rounds").value = 4;
  document.getElementById("markPerRound").value = 2;
  document.getElementById("resultMsg").innerText = "";

  // Reset file list + imported groups
  if (typeof window.resetRankImportUI === "function") {
    window.resetRankImportUI({ clearResult: true });
  } else {
    // Fallback (in case index.html doesn't expose resetRankImportUI)
    const input = document.getElementById("importFile");
    if (input) input.value = "";
    groups = [];
  }
}
