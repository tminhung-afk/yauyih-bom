"use strict";
// Only encrypted Office containers may be used as the shared website source.
window.WorkbookSecurity = {
  isEncrypted(buffer) {
    const bytes = new Uint8Array(buffer);
    const magic = [208, 207, 17, 224, 161, 177, 26, 225];
    if (!magic.every((value, index) => bytes[index] === value)) return false;
    try {
      const container = XLSX.CFB.read(bytes, { type: "array" });
      return !!(XLSX.CFB.find(container, "EncryptionInfo") && XLSX.CFB.find(container, "EncryptedPackage"));
    } catch { return false; }
  },
  async unlock(buffer) {
    const dialog = document.getElementById("passwordDialog");
    const input = document.getElementById("sharedPassword");
    const form = document.getElementById("passwordForm");
    const error = document.getElementById("passwordError");
    const submit = document.getElementById("unlockButton");
    const cancel = document.getElementById("cancelUnlockButton");
    input.value = ""; error.textContent = "";
    dialog.showModal(); input.focus();
    return new Promise((resolve, reject) => {
      let working = false;
      const finish = (book, failure) => {
        input.value = "";
        form.removeEventListener("submit", onSubmit);
        dialog.removeEventListener("cancel", onCancel);
        cancel.removeEventListener("click", onCancel);
        dialog.close();
        if (failure) reject(failure); else resolve(book);
      };
      const onCancel = event => {
        event.preventDefault();
        if (!working) finish(null, new Error("已取消解鎖料表。"));
      };
      const onSubmit = async event => {
        event.preventDefault();
        if (working || !input.value) return;
        working = true; submit.disabled = cancel.disabled = input.disabled = true;
        error.textContent = "正在解鎖，請稍候…";
        let password = input.value; input.value = "";
        try {
          // Allow the progress message to paint before the CPU-intensive decrypt.
          await new Promise(done => setTimeout(done, 30));
          const workbook = await XlsxPopulate.fromDataAsync(buffer, { password });
          const plain = await workbook.outputAsync({ type: "uint8array" });
          finish(XLSX.read(plain, { type: "array" }));
        } catch {
          error.textContent = "密碼不正確，或檔案不是支援的 Excel 加密格式。請重試或取消。";
        } finally {
          password = ""; working = false;
          submit.disabled = cancel.disabled = input.disabled = false;
          if (dialog.open) input.focus();
        }
      };
      form.addEventListener("submit", onSubmit);
      dialog.addEventListener("cancel", onCancel);
      cancel.addEventListener("click", onCancel);
    });
  }
};
