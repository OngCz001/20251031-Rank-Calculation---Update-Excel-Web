// Rank Score Calculator (Multi-file / Multi-sheet)
// - Import multiple .txt files (each file = one group)
// - Calculate each group independently
// - Export ONE .xlsx with multiple sheets (one sheet per group)
// - Supports opponent tokens like: 6 OR W6/D6/L6 (W/D/L shown in Excel, ignored in calculation)

let groups = []; // [{ fileName: string, sheetBase: string, players: string[][] }]

let roundCount = 4;
let fullMark = 2;
let outputFileName = "output";

// Expose for index.html reset function
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

// ===== Calculation =====
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
    if (!group.players || group.players.length === 0) {
      continue;
    }

    const { wsData, meta, groupUpdateLines } = calculateGroup(group.players, group.sheetBase);
    updateLines.push(...groupUpdateLines);

    const worksheet = XLSX.utils.aoa_to_sheet(wsData);
    // Apply Excel-like styling (requires xlsx-js-style in index.html)
    applySheetFormatting(worksheet, {
      sheetTitle: group.sheetBase,
      roundCount,
      headerRowIndex: meta.headerRowIndex,
      dataStartRowIndex: meta.dataStartRowIndex,
      mainColCount: meta.mainColCount,
      totalCols: meta.totalCols,
      dataRowCount: meta.dataRowCount,
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

function calculateGroup(players, sheetTitle) {
  // Build rows for Excel
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
      // Keep the raw tokens for display (e.g., W6)
      ...player.slice(3, 3 + roundCount),
      score,
      avgOpponent,
      expected,
      change.toFixed(1),
      finalRank,
    ];
  });

  // ---- Sheet layout (to mimic the screenshot) ----
  // We create some top rows for: title, legend, and K-table (right side)
  // Main table begins at `headerRowIndex`.
  const headerRow = [
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

  const mainColCount = headerRow.length;
  const totalCols = Math.max(mainColCount, 16); // extra cols for K-table at the right

  const padRow = (arr) => {
    const row = Array.from(arr);
    while (row.length < totalCols) row.push("");
    return row;
  };

  const titleText = `${sheetTitle || ""} 等级分比赛`.trim();

  // Rows (1-indexed in Excel):
  // 1: Title (merged)
  // 2-6: K-table on the right
  // 3: Legend W/D/L
  // 8: Header row
  const wsData = [];
  wsData.push(padRow([titleText])); // row 1

  // row 2 (K table header on the right)
  const row2 = padRow([]);
  row2[14] = "等级分"; // O
  row2[15] = "K值";   // P
  wsData.push(row2);

  // row 3 (legend + first K row)
  const row3 = padRow([]);
  row3[4] = "W=WIN";
  row3[5] = "D=DRAW";
  row3[6] = "L=LOSE";
  row3[14] = "2000或以上";
  row3[15] = 10;
  wsData.push(row3);

  const row4 = padRow([]);
  row4[14] = "1700-1999";
  row4[15] = 15;
  wsData.push(row4);

  const row5 = padRow([]);
  row5[14] = "1550-1699";
  row5[15] = 20;
  wsData.push(row5);

  const row6 = padRow([]);
  row6[14] = "1549或以下";
  row6[15] = 30;
  wsData.push(row6);

  wsData.push(padRow([])); // row 7 blank

  const headerRowIndex = wsData.length; // 0-based index, header will be added next
  wsData.push(padRow(headerRow));       // row 8 header

  const dataStartRowIndex = wsData.length; // first data row (0-based)
  updatedPlayers.forEach((r) => wsData.push(padRow(r)));

  // Build update text lines: name,oldRank,newRank
  const groupUpdateLines = updatedPlayers.map((row) => {
    const name = row[1];
    const oldRank = row[2];
    const newRank = row[row.length - 1];
    return `${name},${oldRank},${newRank}`;
  });

  const meta = {
    headerRowIndex,
    dataStartRowIndex,
    mainColCount,
    totalCols,
    dataRowCount: updatedPlayers.length,
  };

  return { wsData, meta, groupUpdateLines };
}

// ===== Excel Formatting (xlsx-js-style) =====
// Note: To make styles work in the browser, index.html must load `xlsx-js-style`.
function applySheetFormatting(ws, opts) {
  if (!ws || !opts) return;

  const {
    sheetTitle,
    roundCount,
    headerRowIndex,
    dataStartRowIndex,
    mainColCount,
    totalCols,
    dataRowCount,
  } = opts;

  const BLACK = "FF000000";

  const borderThin = {
    top: { style: "thin", color: { rgb: BLACK } },
    bottom: { style: "thin", color: { rgb: BLACK } },
    left: { style: "thin", color: { rgb: BLACK } },
    right: { style: "thin", color: { rgb: BLACK } },
  };

  const borderMedium = {
    top: { style: "medium", color: { rgb: BLACK } },
    bottom: { style: "medium", color: { rgb: BLACK } },
    left: { style: "medium", color: { rgb: BLACK } },
    right: { style: "medium", color: { rgb: BLACK } },
  };

  const styleBase = {
    font: { name: "Calibri", sz: 11, color: { rgb: "FF000000" } },
    alignment: { vertical: "center", horizontal: "center", wrapText: true },
    border: borderThin,
  };

  const fill = (argb) => ({ fill: { patternType: "solid", fgColor: { rgb: argb } } });

  // Screenshot-like colors
  const FILL_W = "FFFFE699"; // yellow
  const FILL_D = "FFC6E0B4"; // green
  const FILL_L = "FF9DC3E6"; // blue
  const FILL_HEADER_GRAY = "FFE7E6E6";
  const FILL_HEADER_ORANGE = "FFF4B183";
  const FILL_HEADER_BLUE = "FFBDD7EE";
  const FILL_HEADER_YELLOW = "FFFFF2CC";
  const FILL_WHITE = "FFFFFFFF";

  function ensureCell(r, c) {
    const addr = XLSX.utils.encode_cell({ r, c });
    if (!ws[addr]) ws[addr] = { t: "s", v: "" };
    return ws[addr];
  }

  function setStyle(r, c, s) {
    const cell = ensureCell(r, c);
    cell.s = s;
  }

  // ----- Column widths (feel free to adjust) -----
  const cols = [];
  // Main table
  cols[0] = { wch: 6 };  // 编号
  cols[1] = { wch: 12 }; // 棋手
  cols[2] = { wch: 8 };  // 等级分
  cols[3] = { wch: 6 };  // K值
  for (let i = 0; i < roundCount; i++) cols[4 + i] = { wch: 6 }; // rounds
  cols[4 + roundCount] = { wch: 8 };  // 总得分
  cols[5 + roundCount] = { wch: 14 }; // 平均对手等级分
  cols[6 + roundCount] = { wch: 8 };  // 期望分
  cols[7 + roundCount] = { wch: 8 };  // 变化
  cols[8 + roundCount] = { wch: 10 }; // 最终等级分

  // Extra cols for K-table (O,P) if present
  for (let c = cols.length; c < totalCols; c++) cols[c] = { wch: 12 };
  cols[14] = { wch: 12 }; // O
  cols[15] = { wch: 6 };  // P
  ws["!cols"] = cols;

  // ----- Merges -----
  ws["!merges"] = ws["!merges"] || [];
  // Title merge across main table only
  ws["!merges"].push({
    s: { r: 0, c: 0 },
    e: { r: 0, c: Math.max(0, mainColCount - 1) },
  });

  // ----- Title styling -----
  // Put title value in A1
  ensureCell(0, 0).v = `${sheetTitle || ""} 等级分比赛`.trim();
  setStyle(0, 0, {
    font: { name: "Calibri", sz: 16, bold: true, color: { rgb: "FF000000" } },
    alignment: { vertical: "center", horizontal: "left" },
  });

  // ----- Legend (row 3, cols E-G) -----
  const legendRow = 2; // row 3 in Excel
  const legend = [
    { c: 4, txt: "W=WIN", fill: FILL_W },
    { c: 5, txt: "D=DRAW", fill: FILL_D },
    { c: 6, txt: "L=LOSE", fill: FILL_L },
  ];
  for (const item of legend) {
    ensureCell(legendRow, item.c).v = item.txt;
    setStyle(legendRow, item.c, {
      ...styleBase,
      ...fill(item.fill),
      font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FF000000" } },
    });
  }

  // ----- K-table (O2:P6) -----
  const kTop = 1;
  const kLeft = 14;
  const kRight = 15;
  const kBottom = 5;
  for (let r = kTop; r <= kBottom; r++) {
    for (let c = kLeft; c <= kRight; c++) {
      ensureCell(r, c);
      const isHeader = r === kTop;
      const st = {
        ...styleBase,
        ...fill(isHeader ? FILL_HEADER_GRAY : FILL_WHITE),
        font: { name: "Calibri", sz: 11, bold: isHeader, color: { rgb: "FF000000" } },
      };

      // outer border medium
      if (r === kTop || r === kBottom || c === kLeft || c === kRight) {
        st.border = {
          top: { style: r === kTop ? "medium" : "thin", color: { rgb: BLACK } },
          bottom: { style: r === kBottom ? "medium" : "thin", color: { rgb: BLACK } },
          left: { style: c === kLeft ? "medium" : "thin", color: { rgb: BLACK } },
          right: { style: c === kRight ? "medium" : "thin", color: { rgb: BLACK } },
        };
      }

      setStyle(r, c, st);
    }
  }

  // ----- Main table styling -----
  const headerR = headerRowIndex;
  const firstDataR = dataStartRowIndex;
  const lastDataR = dataStartRowIndex + dataRowCount - 1;
  const lastMainC = mainColCount - 1;

  // Ensure all main cells exist, then style
  for (let r = headerR; r <= Math.max(headerR, lastDataR); r++) {
    for (let c = 0; c <= lastMainC; c++) {
      ensureCell(r, c);
    }
  }

  // Header row fill colors
  for (let c = 0; c <= lastMainC; c++) {
    let headerFill = FILL_HEADER_GRAY;

    if (c === 2) headerFill = FILL_HEADER_ORANGE; // 等级分
    else if (c >= 4 && c < 4 + roundCount) headerFill = FILL_HEADER_YELLOW; // rounds
    else if (c === lastMainC) headerFill = FILL_HEADER_BLUE; // 最终等级分

    setStyle(headerR, c, {
      ...styleBase,
      ...fill(headerFill),
      font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FF000000" } },
    });
  }

  // Data rows base style + conditional colors
  const changeCol = 7 + roundCount;
  const roundStartCol = 4;

  for (let r = firstDataR; r <= lastDataR; r++) {
    for (let c = 0; c <= lastMainC; c++) {
      // base
      const base = { ...styleBase, ...fill(FILL_WHITE) };

      // Round result coloring by prefix
      if (c >= roundStartCol && c < roundStartCol + roundCount) {
        const v = String((ws[XLSX.utils.encode_cell({ r, c })] || {}).v || "").trim();
        const p = (v[0] || "").toUpperCase();
        if (p === "W") Object.assign(base, fill(FILL_W));
        else if (p === "D") Object.assign(base, fill(FILL_D));
        else if (p === "L") Object.assign(base, fill(FILL_L));
      }

      // Change column: green for +, red for -
      if (c === changeCol) {
        const v = parseFloat(String((ws[XLSX.utils.encode_cell({ r, c })] || {}).v || "0"));
        if (Number.isFinite(v) && v > 0) {
          base.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FF1F7A1F" } };
        } else if (Number.isFinite(v) && v < 0) {
          base.font = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFB00020" } };
        }
      }

      setStyle(r, c, base);
    }
  }

  // Outer border medium around main table
  for (let r = headerR; r <= Math.max(headerR, lastDataR); r++) {
    for (let c = 0; c <= lastMainC; c++) {
      const cell = ensureCell(r, c);
      const b = { ...borderThin };
      if (r === headerR) b.top = { style: "medium", color: { rgb: BLACK } };
      if (r === Math.max(headerR, lastDataR)) b.bottom = { style: "medium", color: { rgb: BLACK } };
      if (c === 0) b.left = { style: "medium", color: { rgb: BLACK } };
      if (c === lastMainC) b.right = { style: "medium", color: { rgb: BLACK } };
      if (!cell.s) cell.s = { ...styleBase };
      cell.s.border = b;
    }
  }
}

function downloadText(filename, content) {
  const blob = new Blob([content], { type: "text/plain" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}

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
    [146, 153, diff >= 0 ? 0.70 : 0.30],
    [154, 162, diff >= 0 ? 0.71 : 0.29],
    [163, 170, diff >= 0 ? 0.72 : 0.28],
    [171, 179, diff >= 0 ? 0.73 : 0.27],
    [180, 188, diff >= 0 ? 0.74 : 0.26],
    [189, 197, diff >= 0 ? 0.75 : 0.25],
    [198, 206, diff >= 0 ? 0.76 : 0.24],
    [207, 215, diff >= 0 ? 0.77 : 0.23],
    [216, 225, diff >= 0 ? 0.78 : 0.22],
    [226, 235, diff >= 0 ? 0.79 : 0.21],
    [236, 245, diff >= 0 ? 0.80 : 0.20],
    [246, 256, diff >= 0 ? 0.81 : 0.19],
    [257, 267, diff >= 0 ? 0.82 : 0.18],
    [268, 278, diff >= 0 ? 0.83 : 0.17],
    [279, 290, diff >= 0 ? 0.84 : 0.16],
    [291, 302, diff >= 0 ? 0.85 : 0.15],
    [303, 315, diff >= 0 ? 0.86 : 0.14],
    [316, 328, diff >= 0 ? 0.87 : 0.13],
    [329, 344, diff >= 0 ? 0.88 : 0.12],
    [345, 357, diff >= 0 ? 0.89 : 0.11],
    [358, 374, diff >= 0 ? 0.90 : 0.10],
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
