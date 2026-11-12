/* Forwardlog Progress Reports: PDF and Excel builders, loaded by index.html only when an export is made.
   index.html calls window.FL_PROGRESS_EXPORTS(helpers) and gets { timetablePdf, timetableXlsx, buildReportPdf }.
   helpers carries the page's own functions plus the letterhead images as data URLs (WATERMARK, LOGO). */
window.FL_PROGRESS_EXPORTS = function (h) {
"use strict";
const { WATERMARK, LOGO, LOGO_RATIO, COMPANY, DOW, MON, D, addDays, validDate, isWork, greyed, projDates, plannedFinish,
  fmtDMY, pct, repNo, activitySummary, drawTimetable, actualDays, slots, isAsset, photoSrc, loadScript, loadPdfLibs } = h;

async function imgData(id, max = 1400, aspect = 0) {
  const src = await photoSrc(id);
  const bmp = await new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = () => rej(new Error("img")); im.src = src; setTimeout(() => rej(new Error("img timeout")), 20000); });
  // With an aspect (width / height) the whole photo is fitted into a box of that shape on a light grey
  // background: every box is the same size and nothing is cut off.
  let bw = bmp.width, bh = bmp.height;
  if (aspect) { if (bw / bh > aspect) bh = bw / aspect; else bw = bh * aspect; }
  const k = Math.min(1, max / Math.max(bw, bh));
  const c = document.createElement("canvas"); c.width = Math.round(bw * k); c.height = Math.round(bh * k);
  const g = c.getContext("2d"); g.fillStyle = aspect ? "#F2F2F2" : "#fff"; g.fillRect(0, 0, c.width, c.height);
  const dw = bmp.width * k, dh = bmp.height * k; g.drawImage(bmp, (c.width - dw) / 2, (c.height - dh) / 2, dw, dh);
  return { data: c.toDataURL("image/jpeg", .82), w: c.width, h: c.height };
}
// Letterhead shared by the report and the timetable PDF.
function pdfLetterhead(doc, W, title, info, right) {
  const M = 15;
  if (W > 250) { const ph = 210, pw = ph / 1.4142; doc.addImage(WATERMARK, "PNG", (W - pw) / 2, 0, pw, ph); }
  else doc.addImage(WATERMARK, "PNG", 0, 0, W, W * 1.4142);
  // Company block: each line shrunk to fit a 58 mm column, like the Excel letterhead.
  const BW = 58, cx = M + BW / 2;
  const fit = (s, size, bold) => { doc.setFont("helvetica", bold ? "bold" : "normal"); let z = size; doc.setFontSize(z); while (doc.getTextWidth(s) > BW && z > 3.5) { z -= 0.1; doc.setFontSize(z); } doc.text(s, cx, fit.y, { align: "center" }); };
  doc.setTextColor(0, 0, 0);
  fit.y = 13.5; fit(COMPANY[0], 6.6, true);
  [["A-06-28, Kompleks Perindustrian Emhub, Persiaran Surian, Seksyen 3,", 16.2], ["Taman Sains Selangor, Kota Damansara,", 18.4], ["47810 Petaling Jaya, Selangor, Malaysia", 20.6]].forEach(([l, yy]) => { fit.y = yy; fit(l, 5, false); });
  fit.y = 23.4; fit("Tel:  +603 2740 1988     Email: sales@forwardlogauto.com", 6.2, false);
  fit.y = 26; fit("Website: www.forwardlogauto.com", 6.2, false);
  const logoW = 42, tx = (M + BW + (W - M - logoW)) / 2;
  doc.setFont("helvetica", "bold"); doc.setFontSize(15); doc.text(title, tx, 20.5, { align: "center" });
  doc.addImage(LOGO, "PNG", W - M - logoW, 15, logoW, logoW / LOGO_RATIO);
  doc.setDrawColor(0, 0, 0); doc.setLineWidth(0.3); doc.line(M, 31, W - M, 31);
  doc.setFont("helvetica", "normal"); doc.setFontSize(9.5);
  info.forEach(([k, v], i) => { doc.text(k, M + 1, 37 + i * 4.6); doc.text(String(v || ""), M + 36, 37 + i * 4.6); });
  if (right) { doc.text(right[0], W - M - 34, 37); doc.text(String(right[1]), W - M - 34, 41.6); }
  const yEnd = 37 + info.length * 4.6 - 1;
  doc.line(M, yEnd, W - M, yEnd);
  return yEnd + 9;
}
function pdfFooters(doc, W, H) {
  const n = doc.getNumberOfPages();
  for (let i = 1; i <= n; i++) { doc.setPage(i); doc.setFont("helvetica", "normal"); doc.setFontSize(9); doc.setTextColor(0, 0, 0); doc.text("Page " + i, W / 2, H - 8, { align: "center" }); }
}
function pdfImage(doc, canvas, x, y, maxW, maxH) {
  const r = canvas.height / canvas.width; let w = maxW, h = w * r;
  if (h > maxH) { h = maxH; w = h / r; }
  doc.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", x, y, w, h);
  return h;
}
// Timetable PDF (A4 landscape) in the Forwardlog letterhead.
async function timetablePdf(p, ctx, kind) {
  const jsPDF = await loadPdfLibs();
  const doc = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });
  const W = 297, H = 210, M = 15;
  const { start } = projDates(p); const fin = plannedFinish(p);
  let y = pdfLetterhead(doc, W, "Project Timetable", [["Project Name:", p.name], ["Project No:", p.code], ["Client:", p.client], ["Location:", p.location], ["Planned:", fmtDMY(start) + " - " + fmtDMY(fin)]], kind === "plan" ? ["Status:", "Planning"] : ["Schedule:", "Rev. " + (p.rev || 0)]);
  pdfImage(doc, drawTimetable(p, { ...ctx, scale: 3 }), M, y, W - 2 * M, H - y - 14);
  pdfFooters(doc, W, H);
  return doc;
}
// Built from pieces: a literal "<" followed by "?" breaks Google Apps Script's HTML service.
const XML_HEAD = "<" + "?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?" + ">";
// Excel in the timetable-conversion layout: Phase | Activity | STATUS | Scheduled/Actual + daily grid.
async function timetableXlsx(p, ctx) {
  await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");
  const list = (p.phases || []).map((ph) => ({ ph, items: (ph.activities || []).filter((a) => validDate(a.start) && validDate(a.end)) })).filter((x) => x.items.length);
  const { start, end } = projDates(p);
  const days = []; for (let d = start; d <= end && days.length < 400; d = addDays(d, 1)) days.push(d);
  const colL = (i) => { let s = "", n = i + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const x = (s) => String(s ?? "").replace(/[&<>"]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch]));
  const rows = new Map(); const merges = [];
  const put = (r, c, v, s) => { if (!rows.has(r)) rows.set(r, []); rows.get(r).push({ c, v, s }); };
  const FIRST = 4;
  // styles: 1 header, 2 month, 3 weekday, 4 date, 5 activity, 6 centre, 7 label, 8 cell, 9 grey, 10 red, 11 green, 12 phase
  [["Phase", 0], ["Activity", 1], ["STATUS", 2], ["", 3]].forEach(([t, c]) => { for (let r = 1; r <= 3; r++) put(r, c, r === 1 ? t : "", 1); merges.push(colL(c) + "1:" + colL(c) + "3"); });
  let mStart = 0;
  days.forEach((d, i) => {
    const c = FIRST + i; const off = greyed(p, d);
    put(1, c, "", 2); put(2, c, DOW[D(d).getUTCDay()], off ? 13 : 3); put(3, c, D(d).getUTCDate(), off ? 14 : 4);
    const next = days[i + 1];
    if (!next || next.slice(0, 7) !== d.slice(0, 7)) {
      const dt = D(d); rows.get(1).find((q) => q.c === FIRST + mStart).v = MON[dt.getUTCMonth()] + "-" + String(dt.getUTCFullYear()).slice(2);
      if (i > mStart) merges.push(colL(FIRST + mStart) + "1:" + colL(c) + "1");
      mStart = i + 1;
    }
  });
  let r = 4;
  list.forEach(({ ph, items }) => {
    const r0 = r;
    items.forEach((a) => {
      const f = Number(ctx.prog?.[a.id]) || 0; const set = ctx.actual?.get(a.id) || new Set();
      put(r, 0, r === r0 ? ph.name : "", 12); put(r + 1, 0, "", 12);
      put(r, 1, a.name, 5); put(r + 1, 1, "", 5); merges.push("B" + r + ":B" + (r + 1));
      put(r, 2, ctx.prog ? (f >= 1 ? "Completed" : f > 0 || set.size ? "Ongoing" : "") : "", 6); put(r + 1, 2, "", 6); merges.push("C" + r + ":C" + (r + 1));
      put(r, 3, "Scheduled", 7); put(r + 1, 3, "Actual", 7);
      days.forEach((d, i) => {
        const off = greyed(p, d);
        put(r, FIRST + i, "", d >= a.start && d <= a.end && isWork(p, d) ? 10 : off ? 9 : 8);
        put(r + 1, FIRST + i, "", set.has(d) ? 11 : off ? 9 : 8);
      });
      r += 2;
    });
    merges.push("A" + r0 + ":A" + (r - 1));
  });
  const cell = (rr, q) => {
    const ref = colL(q.c) + rr;
    if (q.v === "" || q.v === null || q.v === undefined) return `<c r="${ref}" s="${q.s}"/>`;
    if (typeof q.v === "number") return `<c r="${ref}" s="${q.s}"><v>${q.v}</v></c>`;
    return `<c r="${ref}" s="${q.s}" t="inlineStr"><is><t xml:space="preserve">${x(q.v)}</t></is></c>`;
  };
  const sheetRows = [...rows.keys()].sort((a, b) => a - b).map((rr) => `<row r="${rr}"${rr <= 3 ? ` ht="${rr === 1 ? 24 : 18}" customHeight="1"` : ` ht="16" customHeight="1"`}>${rows.get(rr).sort((a, b) => a.c - b.c).map((q) => cell(rr, q)).join("")}</row>`).join("");
  const cols = `<cols><col min="1" max="1" width="26" customWidth="1"/><col min="2" max="2" width="52" customWidth="1"/><col min="3" max="3" width="13" customWidth="1"/><col min="4" max="4" width="11" customWidth="1"/>${days.length ? `<col min="5" max="${4 + days.length}" width="5.8" customWidth="1"/>` : ""}</cols>`;
  const sheet = `${XML_HEAD}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="4" ySplit="3" topLeftCell="E4" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>${cols}<sheetData>${sheetRows}</sheetData><mergeCells count="${merges.length}">${merges.map((m) => `<mergeCell ref="${m}"/>`).join("")}</mergeCells><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`;
  const fill = (rgb) => `<fill><patternFill patternType="solid"><fgColor rgb="FF${rgb}"/><bgColor indexed="64"/></patternFill></fill>`;
  const xf = (font, fillId, h, wrap) => `<xf numFmtId="0" fontId="${font}" fillId="${fillId}" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="${h}" vertical="center"${wrap ? ' wrapText="1"' : ""}/></xf>`;
  const styles = `${XML_HEAD}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><name val="Calibri"/></font><font><sz val="9"/><name val="Calibri"/></font></fonts>
<fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fill("F8CBAD")}${fill("FF0000")}${fill("00B050")}${fill("D9D9D9")}${fill("E2EFDA")}</fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="15"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>${xf(2, 0, "center", true)}${xf(2, 2, "center")}${xf(1, 0, "center")}${xf(0, 0, "center")}${xf(0, 6, "left", true)}${xf(0, 0, "center", true)}${xf(3, 0, "center")}${xf(0, 0, "center")}${xf(0, 5, "center")}${xf(0, 3, "center")}${xf(0, 4, "center")}${xf(1, 0, "center", true)}${xf(1, 5, "center")}${xf(0, 5, "center")}</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  const zip = new window.JSZip();
  zip.file("[Content_Types].xml", `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
  zip.file("_rels/.rels", `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file("xl/workbook.xml", `${XML_HEAD}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Schedule" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file("xl/styles.xml", styles);
  zip.file("xl/worksheets/sheet1.xml", sheet);
  return zip.generateAsync({ type: "blob" });
}
async function buildReportPdf(r, p) {
  {
    const jsPDF = await loadPdfLibs();
    const W = 210, H = 297, M = 15, CW = W - 2 * M, BOTTOM = H - 15;
    const info = [["Project Name:", p.name], ["Project No:", p.code], ["Client:", p.client], ["Location:", p.location], ["Date:", fmtDMY(r.date)]];
    const sum = activitySummary(p, r);
    const tt = drawTimetable(p, { prog: r.progress, actual: actualDays(p, r.date), scale: 3 });
    const ratio = tt.height / tt.width;
    let doc, y, TOP;
    const header = () => pdfLetterhead(doc, W, "Project Progress Report", info, ["Report No:", repNo(r.reportNo)]);
    const heading = (t) => { doc.setFont("helvetica", "bold"); doc.setFontSize(11.5); doc.setTextColor(0, 0, 0); doc.text(t, M + 1, y); y += 5; };
    const remarkLines = () => { doc.setFont("helvetica", "normal"); doc.setFontSize(10.5); return doc.splitTextToSize(r.remarks || "-", CW - 22); };
    const remarks = (label) => { const lines = remarkLines(); doc.text(label, M + 1, y); doc.text(lines, M + 22, y); y += lines.length * 5 + 4; };

    // The client report is always 2 pages. Page 1 shrinks the summary table step by step until the timetable has room.
    const LEVELS = [[10, 2.6, 11], [9, 2, 10], [8, 1.5, 9], [7, 1.1, 8], [6, 0.7, 7]];
    let room = 0;
    for (let li = 0; li < LEVELS.length; li++) {
      const [fs, pad, hfs] = LEVELS[li];
      doc = new jsPDF({ unit: "mm", format: "a4" });
      TOP = header(); y = TOP;
      heading("1.0 ACTIVITY SUMMARY:"); y += 1;
      doc.autoTable({
        startY: y, margin: { left: M, right: M, top: TOP, bottom: 15 }, theme: "grid",
        headStyles: { fillColor: [95, 122, 97], textColor: 255, fontStyle: "bold", fontSize: hfs, halign: "center", lineColor: 0, lineWidth: 0.25, cellPadding: pad },
        styles: { font: "helvetica", fontSize: fs, cellPadding: pad, lineColor: 0, lineWidth: 0.25, textColor: 0, valign: "middle" },
        columnStyles: { 0: { cellWidth: 11, halign: "center" }, 1: { cellWidth: 52 }, 3: { cellWidth: 30, halign: "center" } },
        head: [["NO", "TASK", "DESCRIPTION", "STATUS"]],
        body: sum.length ? sum.map((s, i) => [String(i + 1), s.a.name, s.note, s.f >= 1 ? "Completed" : "Ongoing"]) : [["-", "-", "No activity notes in this report.", "-"]]
      });
      y = doc.lastAutoTable.finalY + 8;
      room = BOTTOM - (y + 18) - (remarkLines().length * 5 + 6);
      const fits = doc.getNumberOfPages() === 1 && room >= Math.min(CW * ratio, 75);
      if (fits || li === LEVELS.length - 1) break;
    }
    while (doc.getNumberOfPages() > 1) doc.deletePage(doc.getNumberOfPages());
    heading("2.0 PROGRESS SUMMARY:"); y += 3;
    doc.setFont("helvetica", "normal"); doc.setFontSize(10.5); doc.text("Scheduled Completion:", M + 1, y); doc.text("Actual Completion:", M + 1, y + 5.5);
    doc.setFont("helvetica", "bold"); doc.setFontSize(11.5); doc.text(pct(r.scheduledPct), M + 52, y);
    doc.setTextColor(0, 176, 80); doc.text(pct(r.actualPct), M + 52, y + 5.5); doc.setTextColor(0, 0, 0);
    y += 10;
    const ttH = Math.max(20, Math.min(CW * ratio, room)), ttW = ttH / ratio;
    y += pdfImage(doc, tt, M + (CW - ttW) / 2, y, ttW, ttH) + 6;
    remarks("Remarks:");

    doc.addPage(); y = header();
    heading("3.0 PROJECT PHOTOS:"); y += 3;
    const photos = slots(r).filter((ph) => ph && isAsset(ph.id));
    const cw = (CW - 14) / 2;
    doc.setFont("helvetica", "normal"); doc.setFontSize(9.5);
    const capRows = [];
    for (let i = 0; i < photos.length; i += 2) capRows.push(photos.slice(i, i + 2).map((ph) => doc.splitTextToSize(ph.caption || "", cw).slice(0, 3)));
    // Photos shrink so both rows, the remarks and End of Report stay on page 2.
    const capH = capRows.reduce((s, row) => s + 10 + Math.max(...row.map((c) => c.length)) * 4, 0);
    const ch = Math.min(58, (BOTTOM - y - capH - (remarkLines().length * 5 + 4) - 16) / Math.max(1, capRows.length));
    for (let i = 0; i < photos.length; i += 2) {
      const pair = photos.slice(i, i + 2), caps = capRows[i / 2];
      const rowh = ch + 5 + Math.max(...caps.map((c) => c.length)) * 4 + 5;
      for (let k = 0; k < pair.length; k++) {
        const x0 = M + 3 + k * (cw + 8);
        try {
          // Every photo sits whole in the same 4:3 box, so the page looks even and nothing is cut off.
          const im = await imgData(pair[k].id, 1600, 4 / 3);
          const iw = Math.min(cw, ch * 4 / 3), ih = iw * 3 / 4, ix = x0 + (cw - iw) / 2, iy = y + (ch - ih) / 2;
          doc.addImage(im.data, "JPEG", ix, iy, iw, ih);
        } catch { doc.setDrawColor(0, 0, 0); doc.setLineWidth(0.4); doc.rect(x0, y, cw, ch); doc.setFontSize(9); doc.text("Photo unavailable", x0 + cw / 2, y + ch / 2, { align: "center" }); }
        doc.setFont("helvetica", "normal"); doc.setFontSize(9.5); doc.setTextColor(0, 0, 0);
        doc.text(caps[k], x0 + cw / 2, y + ch + 5, { align: "center" });
      }
      y += rowh;
    }
    y += 4;
    remarks("Remarks:");
    doc.setFont("helvetica", "bold"); doc.setFontSize(14); doc.text("~ End of Report ~", W / 2, y + 7, { align: "center" });
    pdfFooters(doc, W, H);
    return doc.output("blob");
  }
}

return { timetablePdf, timetableXlsx, buildReportPdf };
};
