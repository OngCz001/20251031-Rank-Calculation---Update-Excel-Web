function processRankUpdate() {
  const excelInput = document.getElementById("excelFile").files[0];
  const textInput = document.getElementById("rankUpdateFile").files[0];

  if (!excelInput || !textInput) {
    alert("Please upload both Excel and Rank Update text files.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheetRaw = workbook.Sheets[sheetName];
    const sheet = XLSX.utils.sheet_to_json(sheetRaw, { header: 1 });

    if (!sheet || sheet.length < 2) {
      alert("The Excel file doesn't contain enough data.");
      return;
    }

    // Use the 2nd row as actual header
    const headerRow = sheet[1];
    const header = headerRow.map(cell => (cell || "").toString().replace(/\s/g, "").trim());

    console.log("Detected Headers:", header);

    const nameIndex = header.findIndex(h => h === "姓名");
    const rankIndex = header.findIndex(h => h === "等级分");

    if (nameIndex === -1 || rankIndex === -1) {
      alert("⚠️ '姓名' or '等级分' columns not found in the second row.\nDetected Headers: " + header.join(" | "));
      return;
    }

    const textReader = new FileReader();
    textReader.onload = function (e) {
      const lines = e.target.result.split("\n").filter(Boolean);

      const updates = {};
      lines.forEach(line => {
        const [name, , newRank] = line.trim().split(",");
        updates[name.trim()] = newRank.trim();
      });

      const successList = [];
      const failList = [];

      for (const [name, newRank] of Object.entries(updates)) {
        let matched = false;

        for (let i = 2; i < sheet.length; i++) {
          const row = sheet[i];
          const excelName = (row[nameIndex] || "").toString().trim();

          if (excelName === name) {
            row[rankIndex] = newRank;
            matched = true;
            successList.push(name);
            break;
          }
        }

        if (!matched) failList.push(name);
      }

      document.getElementById("successList").value = successList.join("\n");
      document.getElementById("failList").value = failList.join("\n");

      // Rebuild full sheet with sorted data
      const titleRow = sheet[0];      // First row (title or blank)
      const headerRow = sheet[1];     // Actual headers
      const dataRows = sheet.slice(2); // Data starts from row 3

      // Find the index of 等级分 again (safe sort)
      const rankIndexInHeader = headerRow.findIndex(h => h && h.toString().replace(/\s/g, "").trim() === "等级分");

      // Sort data descending by 等级分
      const sortedRows = [...dataRows].sort((a, b) => {
        const rankA = parseFloat(a[rankIndexInHeader]) || 0;
        const rankB = parseFloat(b[rankIndexInHeader]) || 0;
        return rankB - rankA;
      });

      // Reassemble final sorted sheet
      const finalSheetData = [titleRow, headerRow, ...sortedRows];
      const updatedSheet = XLSX.utils.aoa_to_sheet(finalSheetData);
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, updatedSheet, sheetName);
      XLSX.writeFile(newWorkbook, "Updated_Rankings.xlsx");

      XLSX.utils.book_append_sheet(newWorkbook, updatedSheet, sheetName);
      XLSX.writeFile(newWorkbook, "Updated_Rankings.xlsx");
    };

    textReader.readAsText(textInput, "utf-8");
  };

  reader.readAsArrayBuffer(excelInput);
}
