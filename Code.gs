/**
 * Bentala Carbon Personalities - backend sederhana (Google Apps Script)
 * Fungsi:
 *  1. Menyimpan data peserta (register & hasil) ke Google Sheet.
 *  2. Mengirim email hasil tes ke peserta, lengkap dengan kartu hasil (PNG).
 *
 * Cara pakai ada di README.md.
 */

// ====== KONFIGURASI ======
const SHEET_ID   = "";   // ID Google Sheet (dari URL: docs.google.com/spreadsheets/d/<ID>/edit)
const SITE_URL   = "";   // URL GitHub Pages, contoh: https://username.github.io/bentala-carbon/
const SENDER_NAME = "Bentala";
const SUBJECT    = "Hasil Tes Kepribadian Karbon Bentala kamu";
// =========================

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === "register") {
      sheet_("Peserta", ["Waktu", "Session", "Nama", "Email", "Event"])
        .appendRow([new Date(), d.sid, d.name, d.email, d.event]);
      return json_({ ok: true });
    }
    if (d.action === "result") {
      sheet_("Hasil", ["Waktu", "Session", "Nama", "Email", "Event", "Karakter", "Total tCO2e",
                      "Transportasi", "Energi rumah", "Makanan", "Belanja", "Jawaban"])
        .appendRow([new Date(), d.sid, d.name, d.email, d.event, d.character, d.total,
                    d.transport, d.energi, d.makanan, d.belanja, d.answers]);
      return json_({ ok: true });
    }
    if (d.action === "email") {
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(d.email)) return json_({ ok: false, error: "email" });
      sendResult_(d);
      sheet_("Email Terkirim", ["Waktu", "Session", "Nama", "Email", "Karakter"])
        .appendRow([new Date(), d.sid, d.name, d.email, d.character]);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: "unknown action" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function sendResult_(d) {
  const color = { tunas: "#3F7D3A", penyeimbang: "#2A5B9E", mager: "#B8741A", josjis: "#C2432E" }[d.key] || "#0F4A2C";
  const poster = SITE_URL ? SITE_URL.replace(/\/?$/, "/") + "assets/poster-" + d.key + ".jpg" : "";
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

  const html = `
  <div style="background:#FAF6EE;padding:28px 0;font-family:Arial,Helvetica,sans-serif;color:#2B2622">
   <div style="max-width:560px;margin:0 auto;background:#FFFDF8;border:1px solid #E4DAC8;border-radius:18px;overflow:hidden">
    ${poster ? `<img src="${poster}" alt="${esc(d.character)}" style="width:100%;display:block">` : ""}
    <div style="padding:26px 28px">
     <p style="margin:0 0 6px;font-size:16px">Selamat <b>${esc(d.name)}</b>, karakter kamu adalah</p>
     <h1 style="margin:0 0 6px;font-size:34px;color:${color}">${esc(d.character)}</h1>
     <p style="margin:0 0 16px;font-style:italic">${esc(d.tag)}</p>
     <p style="margin:0 0 18px;font-size:15px"><b style="font-size:24px">± ${esc(d.total)} tCO₂e/tahun</b><br>
       <span style="color:${color}">${esc(d.level)} · ${esc(d.range)}</span></p>
     <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:18px">
      ${d.breakdown.map(b => `<tr><td style="padding:6px 0;border-bottom:1px solid #EFE7D8">${esc(b[0])}</td><td style="padding:6px 0;border-bottom:1px solid #EFE7D8;text-align:right"><b>${esc(b[1])} t</b></td></tr>`).join("")}
     </table>
     <p style="margin:0 0 6px"><b>Kamu mungkin orang yang...</b></p>
     <ul style="margin:0 0 14px;padding-left:20px">${d.traits.map(t => `<li>${esc(t)}</li>`).join("")}</ul>
     <p style="margin:0 0 10px"><b style="color:${color}">Kekuatan super:</b> ${esc(d.sup)}</p>
     <p style="margin:0 0 10px"><b style="color:${color}">Kriptonit karbon:</b> ${esc(d.kry)}</p>
     <p style="margin:0 0 16px"><b style="color:${color}">Langkah berikutnya:</b> ${esc(d.next)}</p>
     <p style="margin:0 0 18px;font-size:18px;color:${color}"><b>${esc(d.quote)}</b></p>
     <p style="margin:0;font-size:12px;color:#8A8077">Angka ini perkiraan kasar dari 10 kebiasaan, bukan perhitungan jejak karbon lengkap.
       Kartu hasilmu terlampir di email ini.</p>
    </div>
   </div>
  </div>`;

  const opts = { name: SENDER_NAME, htmlBody: html };
  if (d.card && d.card.indexOf("data:image/png;base64,") === 0) {
    const bytes = Utilities.base64Decode(d.card.split(",")[1]);
    opts.attachments = [Utilities.newBlob(bytes, "image/png", "Hasil-Bentala-" + d.character.replace(/\s/g, "-") + ".png")];
  }
  MailApp.sendEmail(d.email, SUBJECT, "Selamat " + d.name + ", karakter kamu adalah " + d.character + ".", opts);
}

function sheet_(name, header) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Jalankan sekali dari editor untuk memberi izin akses Sheet & Gmail. */
function authorize() {
  sheet_("Peserta", ["Waktu", "Session", "Nama", "Email", "Event"]);
  Logger.log("Sisa kuota email hari ini: " + MailApp.getRemainingDailyQuota());
}
