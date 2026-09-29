(function (root) {
  "use strict";
  const text = value => String(value ?? "").trim();
  const key = value => text(value).replace(/\s+/g, "").toLowerCase();
  const modelNames = ["型號", "品項", "產品", "商品", "商品名稱", "model", "name"];
  function item(model, description, note) {
    model = text(model);
    const heading = model.startsWith('"分類"');
    return { model, description: text(description), note: text(note), heading, title: heading ? model.slice(4).trim() || "分類" : model };
  }
  function parseWorkbook(workbook, XLSX) {
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!sheet) throw new Error("料表沒有工作表。");
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    if (!matrix.length) throw new Error("料表是空白的。");
    const header = matrix[0].map(text);
    const find = names => header.findIndex(h => names.includes(key(h)));
    const modelCol = find(modelNames);
    const categories = [];
    if (modelCol >= 0) {
      const categoryCol = find(["分類", "類別", "品名", "類型", "category"]);
      const descriptionCol = find(["說明", "規格", "描述", "內容", "description", "desc"]);
      const noteCol = find(["備註", "note", "memo", "remarks"]);
      for (const row of matrix.slice(1)) {
        if (!text(row[modelCol])) continue;
        const name = text(row[categoryCol]) || "匯入清單";
        let category = categories.find(c => c.name === name);
        if (!category) { category = { name, items: [] }; categories.push(category); }
        category.items.push(item(row[modelCol], row[descriptionCol], row[noteCol]));
      }
    } else {
      for (let col = 0; col + 1 < header.length; col++) {
        if (!header[col] || ["說明", "備註"].includes(header[col]) || header[col + 1] !== "說明") continue;
        const noteCol = header[col + 2] === "備註" ? col + 2 : -1;
        const category = { name: header[col], items: [] };
        for (const row of matrix.slice(1)) {
          if (text(row[col])) category.items.push(item(row[col], row[col + 1], row[noteCol]));
        }
        categories.push(category);
        col = noteCol >= 0 ? noteCol : col + 1;
      }
    }
    if (!categories.some(c => c.items.length)) throw new Error("沒有可用資料。請使用「分類名稱、說明、備註」分組欄位，或「分類、型號、說明、備註」清單格式。");
    return categories;
  }
  function fileName(name, now = new Date()) {
    if (!text(name)) throw new Error("請輸入案件名稱。");
    const date = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
    const safe = text(name).replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").replace(/[. ]+$/g, "_");
    return `${date}_${safe}_料表.xlsx`;
  }
  function outputRows(categories, selections) {
    const rows = [];
    categories.forEach((category, i) => (selections[i] || []).forEach(slot => {
      const part = slot.index === null ? null : category.items[slot.index];
      if (!part || part.heading) return;
      rows.push({ no: rows.length + 1, model: part.model, description: part.description, quantity: slot.quantity, note: text(slot.note) });
    }));
    return rows;
  }
  function exportWorkbook(rows, XLSX) {
    const output = rows.length ? rows : Array.from({ length: 15 }, (_, i) => ({ no: i + 1, model: "", description: "", quantity: "", note: "" }));
    const sheet = XLSX.utils.aoa_to_sheet([["項次", "型號", "說明", "數量", "備註"], ...output.map(r => [r.no, r.model, r.description, r.quantity, r.note])]);
    sheet["!cols"] = [7, 28, 42, 8, 28].map(wch => ({ wch }));
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "輸出清單");
    return book;
  }
  const api = { parseWorkbook, fileName, outputRows, exportWorkbook };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.Bom = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
