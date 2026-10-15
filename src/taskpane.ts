/**
 * taskpane.ts
 * Logic giao diện điều khiển Taskpane Word Add-in
 */

import {
  applyPageSetup,
  formatBodyContent,
  insertHeaderTable,
  insertSignatureTable,
  lintDocument,
  standardizeDocument,
} from "./index";
import { LintReport, LintSeverity } from "./types";

/* global Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("btn-standardize-all")?.addEventListener("click", handleStandardizeAll);
    document.getElementById("btn-page-setup")?.addEventListener("click", handlePageSetup);
    document.getElementById("btn-format-body")?.addEventListener("click", handleFormatBody);
    document.getElementById("btn-insert-header")?.addEventListener("click", handleInsertHeader);
    document.getElementById("btn-insert-signature")?.addEventListener("click", handleInsertSignature);
    document.getElementById("btn-lint")?.addEventListener("click", handleLint);
  }
});

function showToast(message: string, type: "success" | "error" = "success") {
  const toast = document.getElementById("status-toast");
  if (!toast) return;
  toast.textContent = message;
  toast.className = `toast ${type}`;
  setTimeout(() => {
    toast.className = "toast hidden";
  }, 3500);
}

async function handleStandardizeAll() {
  try {
    await Word.run(async (context) => {
      showToast("Đang chuẩn hóa toàn bộ văn bản...");
      const report = await standardizeDocument(context, {
        fontSize: 14,
      });
      showToast("✅ Đã chuẩn hóa toàn bộ văn bản thành công!", "success");
      renderLintResults(report);
    });
  } catch (err: any) {
    console.error(err);
    showToast(`❌ Lỗi: ${err.message || err}`, "error");
  }
}

async function handlePageSetup() {
  try {
    await Word.run(async (context) => {
      showToast("Đang thiết lập khổ A4 và căn lề...");
      await applyPageSetup(context);
      showToast("✅ Đã thiết lập khổ A4, lề 20-20-30-15mm và đánh số trang!", "success");
    });
  } catch (err: any) {
    console.error(err);
    showToast(`❌ Lỗi: ${err.message || err}`, "error");
  }
}

async function handleFormatBody() {
  try {
    await Word.run(async (context) => {
      showToast("Đang chuẩn hóa đoạn văn nội dung...");
      await formatBodyContent(context, 14);
      showToast("✅ Đã chuẩn hóa phông Times New Roman, Justified, thụt lề 10mm!", "success");
    });
  } catch (err: any) {
    console.error(err);
    showToast(`❌ Lỗi: ${err.message || err}`, "error");
  }
}

async function handleInsertHeader() {
  try {
    const superiorOrgan = (document.getElementById("inp-superior") as HTMLInputElement)?.value;
    const issuingOrgan = (document.getElementById("inp-issuing") as HTMLInputElement)?.value || "VĂN PHÒNG";
    const docNumber = (document.getElementById("inp-doc-num") as HTMLInputElement)?.value || "05";
    const docCode = (document.getElementById("inp-doc-code") as HTMLInputElement)?.value || "HD/VPTW";
    const location = (document.getElementById("inp-location") as HTMLInputElement)?.value || "Hà Nội";

    await Word.run(async (context) => {
      showToast("Đang chèn Header...");
      await insertHeaderTable(context, {
        superiorOrgan: superiorOrgan.trim() || undefined,
        issuingOrgan: issuingOrgan.trim(),
        docNumber: docNumber,
        docCode: docCode,
        location: location,
        date: new Date(),
      });
      showToast("✅ Đã chèn Header Ô 1 & Ô 2 chuẩn Hướng dẫn 05-HD/VPTW!", "success");
    });
  } catch (err: any) {
    console.error(err);
    showToast(`❌ Lỗi: ${err.message || err}`, "error");
  }
}

async function handleInsertSignature() {
  try {
    const recipientsRaw = (document.getElementById("inp-recipients") as HTMLTextAreaElement)?.value || "";
    const signTitle = (document.getElementById("inp-sign-title") as HTMLInputElement)?.value || "T/M BAN THƯỜNG VỤ";
    const signerRole = (document.getElementById("inp-sign-role") as HTMLInputElement)?.value || "";
    const signerName = (document.getElementById("inp-sign-name") as HTMLInputElement)?.value || "Nguyễn Văn A";

    const recipients = recipientsRaw
      .split("\n")
      .map((line) => line.trim())
      .filter((line) => line.length > 0);

    await Word.run(async (context) => {
      showToast("Đang chèn bảng Chữ ký & Nơi nhận...");
      await insertSignatureTable(context, {
        recipients,
        signTitle,
        signerRole: signerRole.trim() || undefined,
        signerName,
        signatureSpaceMm: 38,
      });
      showToast("✅ Đã chèn Chân trang (Nơi nhận & Chữ ký) thành công!", "success");
    });
  } catch (err: any) {
    console.error(err);
    showToast(`❌ Lỗi: ${err.message || err}`, "error");
  }
}

async function handleLint() {
  try {
    await Word.run(async (context) => {
      showToast("Đang quét lỗi thể thức...");
      const report = await lintDocument(context);
      renderLintResults(report);
      showToast(`Đã quét xong: Phát hiện ${report.totalIssues} vấn đề`, report.isValid ? "success" : "error");
    });
  } catch (err: any) {
    console.error(err);
    showToast(`❌ Lỗi: ${err.message || err}`, "error");
  }
}

function renderLintResults(report: LintReport) {
  const container = document.getElementById("lint-result-container");
  if (!container) return;

  if (report.issues.length === 0) {
    container.innerHTML = `
      <div class="lint-badge-success">
        ✓ Tài liệu chuẩn 100% theo Hướng dẫn 05-HD/VPTW!
      </div>
    `;
    return;
  }

  let html = `<div style="margin-bottom: 6px; font-weight: 600;">Phát hiện ${report.totalIssues} điểm cần lưu ý:</div>`;
  report.issues.forEach((issue) => {
    const sevClass = issue.severity === LintSeverity.ERROR ? "error" : issue.severity === LintSeverity.WARNING ? "warning" : "info";
    html += `
      <div class="lint-item ${sevClass}">
        <div class="lint-title">[${issue.locationInfo || "Văn bản"}] ${issue.message}</div>
        <div class="lint-suggest">👉 <em>${issue.suggestion}</em></div>
      </div>
    `;
  });

  container.innerHTML = html;
}
