let players = [];
let roundCount = 4;
let fullMark = 2;
let fileName = "output";

document.getElementById("importFile").addEventListener("change", function (e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function (e) {
    const lines = e.target.result.trim().split("\n");
    players = lines.map(line => line.split(","));
  };
  reader.readAsText(file);
});

function calculate() {
  roundCount = parseInt(document.getElementById("rounds").value);
  fullMark = parseInt(document.getElementById("markPerRound").value) * roundCount;
  fileName = document.getElementById("fileName").value || "output";

  if (!players.length) {
    alert("Please import a valid .txt file");
    return;
  }

  const updatedPlayers = players.map((player, index) => {
    const name = player[0];
    const rank = parseInt(player[1]);
    const k = parseInt(player[2]);
    const score = parseFloat(player[3 + roundCount]);

    // Calculate opponent average
    let total = 0;
    let empty = 0;
    for (let i = 0; i < roundCount; i++) {
      const opponentIndex = parseInt(player[3 + i]) - 1;
      if (opponentIndex >= 0 && players[opponentIndex]) {
        total += parseInt(players[opponentIndex][1]);
      } else {
        empty++;
      }
    }

    const avgOpponent = Math.ceil(total / (roundCount - empty));
    const expected = (getExpectedScore(rank, avgOpponent)).toFixed(1);
    const change = (score - expected) * k;
    const final = Math.round(rank + change);

    return [
      index + 1,
      name,
      rank,
      k,
      ...player.slice(3, 3 + roundCount),
      score,
      avgOpponent,
      expected,
      change.toFixed(1),
      final
    ];
  });

  generateExcel(updatedPlayers);
  generateText(updatedPlayers);

  document.getElementById("resultMsg").innerText = `✅ Calculation completed. Files ready for download.`;
}

function getExpectedScore(playerMark, oppMark) {
  const diff = playerMark - oppMark;
  const absDiff = Math.abs(diff);
  const breakpoints = [
    [0, 3, 0.5], [4, 10, diff >= 0 ? 0.51 : 0.49], [11, 17, diff >= 0 ? 0.52 : 0.48],
    [18, 25, diff >= 0 ? 0.53 : 0.47], [26, 32, diff >= 0 ? 0.54 : 0.46],
    [33, 39, diff >= 0 ? 0.55 : 0.45], [40, 46, diff >= 0 ? 0.56 : 0.44],
    [47, 53, diff >= 0 ? 0.57 : 0.43], [54, 61, diff >= 0 ? 0.58 : 0.42],
    [62, 68, diff >= 0 ? 0.59 : 0.41], [69, 76, diff >= 0 ? 0.6 : 0.4],
    [77, 83, diff >= 0 ? 0.61 : 0.39], [84, 91, diff >= 0 ? 0.62 : 0.38],
    [92, 98, diff >= 0 ? 0.63 : 0.37], [99, 106, diff >= 0 ? 0.64 : 0.36],
    [107, 113, diff >= 0 ? 0.65 : 0.35], [114, 121, diff >= 0 ? 0.66 : 0.34],
    [122, 129, diff >= 0 ? 0.67 : 0.33], [130, 137, diff >= 0 ? 0.68 : 0.32],
    [138, 145, diff >= 0 ? 0.69 : 0.31], [146, 153, diff >= 0 ? 0.70 : 0.30],
    [154, 162, diff >= 0 ? 0.71 : 0.29], [163, 170, diff >= 0 ? 0.72 : 0.28],
    [171, 179, diff >= 0 ? 0.73 : 0.27], [180, 188, diff >= 0 ? 0.74 : 0.26],
    [189, 197, diff >= 0 ? 0.75 : 0.25], [198, 206, diff >= 0 ? 0.76 : 0.24],
    [207, 215, diff >= 0 ? 0.77 : 0.23], [216, 225, diff >= 0 ? 0.78 : 0.22],
    [226, 235, diff >= 0 ? 0.79 : 0.21], [236, 245, diff >= 0 ? 0.80 : 0.20],
    [246, 256, diff >= 0 ? 0.81 : 0.19], [257, 267, diff >= 0 ? 0.82 : 0.18],
    [268, 278, diff >= 0 ? 0.83 : 0.17], [279, 290, diff >= 0 ? 0.84 : 0.16],
    [291, 302, diff >= 0 ? 0.85 : 0.15], [303, 315, diff >= 0 ? 0.86 : 0.14],
    [316, 328, diff >= 0 ? 0.87 : 0.13], [329, 344, diff >= 0 ? 0.88 : 0.12],
    [345, 357, diff >= 0 ? 0.89 : 0.11], [358, 374, diff >= 0 ? 0.90 : 0.10],
    [375, 391, diff >= 0 ? 0.91 : 0.09], [392, 411, diff >= 0 ? 0.92 : 0.08],
    [412, 432, diff >= 0 ? 0.93 : 0.07], [433, 456, diff >= 0 ? 0.94 : 0.06],
    [457, 484, diff >= 0 ? 0.95 : 0.05], [485, 517, diff >= 0 ? 0.96 : 0.04],
    [518, 559, diff >= 0 ? 0.97 : 0.03], [560, 619, diff >= 0 ? 0.98 : 0.02],
    [620, 734, diff >= 0 ? 0.99 : 0.01]
  ];

  for (const [min, max, val] of breakpoints) {
    if (absDiff >= min && absDiff <= max) return fullMark * val;
  }

  return diff >= 0 ? fullMark.toFixed(1) : 0;
}

function generateExcel(data) {
  const wsData = [
    ["编号", "棋手", "等级分", "K值", ...Array.from({ length: roundCount }, (_, i) => `第${i + 1}轮`), "总得分", "平均对手等级分", "期望分", "变化", "最终等级分"],
    ...data
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(wsData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Rank Calculation");

  XLSX.writeFile(workbook, `${fileName}.xlsx`);
}

function generateText(data) {
  const lines = data.map(row => {
    const name = row[1];          // Player name
    const oldRank = row[2];       // Original rank
    const newRank = row[row.length - 1]; // Final rank (last column)
    return `${name},${oldRank},${newRank}`;
  }).join("\n");

  const blob = new Blob([lines], { type: "text/plain" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${fileName}.txt`;
  a.click();
}


function clearForm() {
  document.getElementById("importFile").value = "";
  document.getElementById("fileName").value = "";
  document.getElementById("rounds").value = 4;
  document.getElementById("markPerRound").value = 2;
  document.getElementById("resultMsg").innerText = "";
  players = [];
}
