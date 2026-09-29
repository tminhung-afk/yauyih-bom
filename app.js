"use strict";
const $ = id => document.getElementById(id);
const sourceFileName = "料表.xlsx";
let categories = [], selections = [], source = null, busy = false;
const blankSlot = () => ({ index: null, quantity: 1, note: "" });
const fmt = new Intl.DateTimeFormat("zh-TW", { dateStyle: "short", timeStyle: "medium" });
$("workbookName").value = "";

function notify(message, error = false) {
  $("notice").textContent = message;
  $("notice").hidden = !message;
  $("notice").classList.toggle("error", error);
}
function setBusy(value) {
  busy = value;
  document.querySelectorAll(".actions button").forEach(button => { button.disabled = value; });
  for (const id of ["exportButton", "clearButton", "downloadSourceButton"]) $(id).disabled = value || !source;
  $("partsBody").querySelectorAll("input, select, button").forEach(control => { control.disabled = value; });
  $("sourcePanel").setAttribute("aria-busy", String(value));
}
function allowReset() {
  return !selections.some(slots => slots.some(slot => slot.index !== null || slot.note)) || confirm("更新料表會清空已選品項與備註，案件名稱會保留。是否繼續？");
}
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url; link.download = filename;
  document.body.appendChild(link); link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
async function applyWorkbook(buffer, info) {
  if (!window.XLSX) throw new Error("Excel 元件載入失敗，請重新整理網頁。");
  const encrypted = WorkbookSecurity.isEncrypted(buffer);
  if (!info.local && !encrypted) throw new Error("網站料表尚未加密，已停止載入。請管理者改用設定『開啟密碼』的 Excel；未加密檔案不可放在公開網站。");
  const book = encrypted ? await WorkbookSecurity.unlock(buffer) : XLSX.read(buffer, { type: "array" });
  const loaded = Bom.parseWorkbook(book, XLSX);
  // Commit only after parsing succeeds; failures leave current selections intact.
  categories = loaded;
  selections = categories.map(() => [blankSlot()]);
  source = { ...info, buffer };
  const products = categories.flatMap(c => c.items).filter(p => !p.heading);
  $("sourceBadge").textContent = info.local ? "本機料表" : "網站料表";
  $("sourceStatus").textContent = `${info.name} · ${products.length.toLocaleString()} 筆品項 · ${categories.length} 個分類 · ${products.filter(p => p.note).length} 筆備註`;
  $("sourceMeta").textContent = `讀取時間：${fmt.format(new Date())}${info.modified ? `　檔案更新：${fmt.format(info.modified)}` : ""}${info.local ? "　僅在本機讀取，未上傳網站。" : ""}`;
  $("sourcePanel").classList.remove("load-error");
  $("emptyState").hidden = true;
  renderParts(); renderPreview(); notify("");
}
async function refreshWebsite(initial = false) {
  if (busy || (!initial && !allowReset())) return;
  setBusy(true);
  notify("正在讀取網站上的最新料表…");
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 20000);
  try {
    if (location.protocol === "file:") throw new Error("直接開啟 HTML 無法自動讀取資料夾中的 Excel。請從網站網址開啟，或按「指定料表」選擇檔案。");
    const response = await fetch(`./${encodeURIComponent(sourceFileName)}?t=${Date.now()}`, { cache: "no-store", signal: controller.signal });
    if (!response.ok) throw new Error(`網站料表讀取失敗（HTTP ${response.status}）。請確認網站已放置「料表.xlsx」，或按「指定料表」選擇檔案。`);
    const modified = new Date(response.headers.get("Last-Modified") || "");
    await applyWorkbook(await response.arrayBuffer(), { name: sourceFileName, local: false, modified: Number.isNaN(modified.valueOf()) ? null : modified });
  } catch (error) {
    notify(`${error.name === "AbortError" ? "讀取逾時，請確認連線後重試。" : error.message}${source ? " 已保留目前料表與選擇。" : ""}`, true);
    if (!source) {
      $("sourceBadge").textContent = "尚未載入";
      $("sourceStatus").textContent = "料表未能載入";
      $("sourceMeta").textContent = "請重試刷新，或指定本機 Excel。";
      $("emptyState").textContent = "載入成功後，可選品項會顯示在這裡。";
      $("sourcePanel").classList.add("load-error");
    }
  } finally { clearTimeout(timeout); setBusy(false); }
}

function renderParts(focus = null) {
  const body = $("partsBody");
  body.replaceChildren();
  categories.forEach((category, ci) => selections[ci].forEach((slot, si) => {
    const row = document.createElement("tr");
    row.className = si ? "addon-row" : "";
    const cell = (className, content) => {
      const td = document.createElement("td"); td.className = className;
      if (typeof content === "string") td.textContent = content; else if (content) td.appendChild(content);
      row.appendChild(td); return td;
    };
    cell("idx", si ? "+" : String(ci + 1));
    cell("category", si ? `${category.name} · 加購` : category.name);
    const choice = document.createElement("div"); choice.className = "choice-controls";
    const select = document.createElement("select"); select.className = "item-select";
    select.setAttribute("aria-label", `${category.name}${si ? "加購" : ""}品項`);
    select.appendChild(new Option(si ? "請選擇加購品項" : `共有商品 ${category.items.filter(p => !p.heading).length} 樣，請選擇`, ""));
    let group = select;
    category.items.forEach((part, index) => {
      if (part.heading) {
        group = document.createElement("optgroup"); group.label = part.title; select.appendChild(group);
      } else group.appendChild(new Option(part.description ? `${part.model}｜${part.description}` : part.model, String(index)));
    });
    select.value = slot.index === null ? "" : String(slot.index);
    const inspect = document.createElement("button"); inspect.type = "button"; inspect.className = "inspect-button"; inspect.textContent = "說明";
    inspect.setAttribute("aria-label", `查看${category.name}品項說明`);
    choice.append(select, inspect); cell("choice-cell", choice);
    const quantity = document.createElement("select"); quantity.className = "qty-select"; quantity.setAttribute("aria-label", `${category.name}數量`);
    for (let n = 1; n <= 20; n++) quantity.appendChild(new Option(String(n), String(n)));
    quantity.value = String(slot.quantity);
    const add = document.createElement("button"); add.type = "button"; add.className = "add-row-button"; add.textContent = "＋"; add.setAttribute("aria-label", `加購${category.name}`);
    const remove = document.createElement("button"); remove.type = "button"; remove.className = "remove-row-button"; remove.textContent = "×"; remove.setAttribute("aria-label", `移除${category.name}加購列`); remove.hidden = !si;
    const controls = document.createElement("div"); controls.className = "qty-controls"; controls.append(quantity, add, remove); cell("qty-cell", controls);
    const note = document.createElement("input"); note.type = "text"; note.className = "note-input"; note.value = slot.note; note.placeholder = "備註（選填）"; note.setAttribute("aria-label", `${category.name}備註`);
    cell("note-cell", note);
    select.addEventListener("change", () => {
      const index = select.value === "" ? null : Number(select.value);
      if (index !== null && (!category.items[index] || category.items[index].heading)) { select.value = slot.index === null ? "" : String(slot.index); return; }
      slot.index = index; slot.note = index === null ? "" : category.items[index].note; note.value = slot.note; renderPreview();
    });
    quantity.addEventListener("change", () => { slot.quantity = Number(quantity.value); renderPreview(); });
    note.addEventListener("input", () => { slot.note = note.value; renderPreview(); });
    add.addEventListener("click", () => { selections[ci].splice(si + 1, 0, blankSlot()); renderParts({ ci, si: si + 1 }); });
    remove.addEventListener("click", () => { selections[ci].splice(si, 1); renderParts({ ci, si: si - 1 }); renderPreview(); });
    inspect.addEventListener("click", () => {
      const part = slot.index === null ? null : category.items[slot.index];
      $("detailTitle").textContent = part?.model || category.name;
      $("detailText").textContent = part ? part.description || "沒有說明內容。" : "請先選擇一個品項。";
      $("detailDialog").showModal();
    });
    body.appendChild(row);
    if (focus?.ci === ci && focus.si === si) select.focus({ preventScroll: true });
  }));
}
function renderPreview() {
  const rows = Bom.outputRows(categories, selections);
  $("selectedCount").textContent = `${rows.length} 項`;
  $("previewEmpty").hidden = rows.length > 0;
  $("previewWrap").hidden = !rows.length;
  $("previewBody").replaceChildren();
  rows.forEach(item => {
    const tr = document.createElement("tr");
    [item.no, item.model, item.description, item.quantity, item.note].forEach((value, i) => {
      const td = document.createElement("td"); td.textContent = value;
      td.dataset.label = ["項次", "型號", "說明", "數量", "備註"][i]; tr.appendChild(td);
    });
    $("previewBody").appendChild(tr);
  });
}
$("lockButton").addEventListener("click", () => { if (allowReset()) location.reload(); });
$("refreshButton").addEventListener("click", () => refreshWebsite());
$("selectWorkbookButton").addEventListener("click", () => { if (!busy) { $("workbookFile").value = ""; $("workbookFile").click(); } });
$("workbookFile").addEventListener("change", async event => {
  const file = event.target.files[0];
  if (!file || busy || !allowReset()) return;
  setBusy(true);
  try { await applyWorkbook(await file.arrayBuffer(), { name: file.name, local: true, modified: file.lastModified ? new Date(file.lastModified) : null }); }
  catch (error) { notify(`無法讀取此檔案：${error.message} 已保留原有資料。`, true); }
  finally { setBusy(false); }
});
$("downloadSourceButton").addEventListener("click", () => { if (source) downloadBlob(new Blob([source.buffer], { type: source.name.toLowerCase().endsWith(".xls") ? "application/vnd.ms-excel" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), source.name); });
$("clearButton").addEventListener("click", () => {
  if (!allowReset()) return;
  selections = categories.map(() => [blankSlot()]); renderParts(); renderPreview(); notify("已清空選擇，案件名稱保留。");
});
$("workbookName").addEventListener("input", () => {
  $("filenamePreview").textContent = $("workbookName").value.trim() ? Bom.fileName($("workbookName").value) : "日期_案件名稱_料表.xlsx";
});
$("exportButton").addEventListener("click", () => {
  if (!$("workbookName").value.trim()) { notify("請先輸入案件名稱。", true); $("workbookName").focus(); return; }
  try {
    const filename = Bom.fileName($("workbookName").value);
    const book = Bom.exportWorkbook(Bom.outputRows(categories, selections), XLSX);
    const buffer = XLSX.write(book, { bookType: "xlsx", type: "array" });
    downloadBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename);
    notify(`已產生 ${filename}。請在瀏覽器下載項目中查看；iPhone／iPad 可儲存至「檔案」。`);
  } catch (error) { notify(`匯出失敗：${error.message}`, true); }
});
renderPreview();
refreshWebsite(true);
