"use client";

import { formatChange, formatPercent, periodLabel, reportNarrative, requiredPeriods, type Entry, type ProgramResult, type ReportScope } from "./report";
import type { ReportSettings } from "./report-settings";

const COLORS = { dark: "114C50", teal: "0C938E", pale: "E6F4F2", ink: "173333", gray: "6D8484", gold: "C69A3C", red: "B04A4A", white: "FFFFFF" };

function download(data: Blob, name: string) {
  const url = URL.createObjectURL(data);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function chartData(item: ProgramResult, kind: "group" | "trend") {
  return kind === "group"
    ? item.groups.map((group) => ({ label: group.label, value: group.mean }))
    : item.trend.map((point) => ({ label: point.label, value: point.mean }));
}

function barChartPng(title: string, rows: { label: string; value: number | null }[]) {
  const available = rows.filter((row) => row.value !== null);
  const canvas = document.createElement("canvas");
  canvas.width = 1200;
  canvas.height = Math.max(430, available.length * 66 + 130);
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("تعذر إنشاء الرسم.");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#114c50";
  ctx.font = "bold 28px Tahoma, Arial";
  ctx.textAlign = "right";
  ctx.fillText(title, 1145, 48);
  if (!available.length) {
    ctx.font = "24px Tahoma, Arial";
    ctx.fillStyle = "#6d8484";
    ctx.fillText("البيانات غير مكتملة لعرض الرسم", 1145, 155);
    return canvas.toDataURL("image/png");
  }
  const barLeft = 115;
  const barWidth = 690;
  const labelRight = 1140;
  const startY = 96;
  const rowHeight = (canvas.height - 160) / available.length;
  available.forEach((row, index) => {
    const y = startY + index * rowHeight;
    ctx.fillStyle = "#eef5f4";
    ctx.fillRect(barLeft, y, barWidth, 27);
    ctx.fillStyle = "#0c938e";
    ctx.fillRect(barLeft, y, barWidth * Math.max(0, Math.min(100, row.value ?? 0)) / 100, 27);
    ctx.fillStyle = "#173333";
    ctx.font = "20px Tahoma, Arial";
    ctx.textAlign = "right";
    const label = row.label.length > 31 ? `${row.label.slice(0, 30)}…` : row.label;
    ctx.fillText(label, labelRight, y + 22);
    ctx.textAlign = "left";
    ctx.fillText(formatPercent(row.value), barLeft + barWidth + 14, y + 22);
  });
  return canvas.toDataURL("image/png");
}

function chartPng(item: ProgramResult, kind: "group" | "trend") {
  return barChartPng(kind === "group" ? "متوسطات تصنيفات المنشآت" : "مقارنة الفترات", chartData(item, kind));
}

function selectedEntries(scope: ReportScope, reports: ProgramResult[], entries: Entry[]) {
  const programs = new Set(reports.map((item) => item.program));
  return entries.filter((entry) => programs.has(entry.program) && (requiredPeriods(entry.program, scope)?.includes(entry.period) ?? false));
}

function safeScope(scope: ReportScope) {
  return scope.replace(/[^A-Za-z0-9,-]/g, "-");
}

export async function exportExcel(year: number, scope: ReportScope, reports: ProgramResult[], entries: Entry[], settings: ReportSettings) {
  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "تقارير برامج وقاية";
  workbook.created = new Date();
  const overview = workbook.addWorksheet("ملخص التقرير", { views: [{ rightToLeft: true }] });
  overview.columns = [{ width: 37 }, { width: 22 }, { width: 19 }, { width: 17 }, { width: 22 }];
  overview.mergeCells("A1:E1");
  overview.getCell("A1").value = `${settings.title} · ${periodLabel(scope)} · ${year}`;
  overview.getCell("A1").font = { name: "Arial", size: 17, bold: true, color: { argb: "FF114C50" } };
  overview.getRow(1).height = 34;
  overview.getCell("A3").value = "الملخص التنفيذي";
  overview.getCell("B3").value = settings.executiveSummary;
  overview.mergeCells("B3:E3");
  overview.getRow(3).height = 42;
  overview.addRow(["البرنامج", "المتوسط %", "منشآت مكتملة", "منشآت جزئية", "إجمالي المنشآت"]);
  for (const item of reports) overview.addRow([item.label, item.mean, item.complete, item.partial, item.total]);
  overview.getRow(4).eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF114C50" } };
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { horizontal: "right" };
  });
  overview.getColumn(2).numFmt = '0.0"%"';
  let summaryRow = 7 + reports.length;
  for (const item of reports) {
    overview.mergeCells(`A${summaryRow}:E${summaryRow}`);
    overview.getCell(`A${summaryRow}`).value = settings.programNotes[item.program] ?? reportNarrative(item, scope);
    overview.getCell(`A${summaryRow}`).alignment = { wrapText: true, vertical: "middle" };
    overview.getRow(summaryRow).height = 54;
    summaryRow += 1;
  }
  if (settings.includeMethodology) {
    overview.mergeCells(`A${summaryRow}:E${summaryRow}`);
    overview.getCell(`A${summaryRow}`).value = `المنهج: ${settings.methodology}`;
    overview.getCell(`A${summaryRow}`).alignment = { wrapText: true };
    overview.getRow(summaryRow).height = 54;
    summaryRow += 1;
  }
  overview.mergeCells(`A${summaryRow}:E${summaryRow}`);
  overview.getCell(`A${summaryRow}`).value = `الخلاصة: ${settings.conclusion}`;
  overview.getCell(`A${summaryRow}`).alignment = { wrapText: true };
  overview.getRow(summaryRow).height = 54;

  for (const item of reports) {
    const sheet = workbook.addWorksheet(item.program.toUpperCase().slice(0, 31), { views: [{ rightToLeft: true, state: "frozen", ySplit: 4 }] });
    sheet.columns = [{ width: 38 }, { width: 44 }, { width: 17 }, { width: 18 }, { width: 19 }, { width: 19 }, { width: 44 }, { width: 44 }];
    sheet.mergeCells("A1:H1");
    sheet.getCell("A1").value = `${item.label} · ${periodLabel(scope)} ${year}`;
    sheet.getCell("A1").font = { name: "Arial", size: 16, bold: true, color: { argb: "FF114C50" } };
    sheet.getRow(1).height = 32;
    sheet.mergeCells("A2:H2");
    sheet.getCell("A2").value = settings.programNotes[item.program] ?? reportNarrative(item, scope);
    sheet.getCell("A2").alignment = { wrapText: true };
    sheet.getRow(2).height = 54;
    sheet.addRow(["التصنيف", "المنشأة", "النتيجة %", "فترات مدخلة", "الفترات المطلوبة", "الحالة", "الملاحظة", "الإجراء"]);
    const head = sheet.getRow(3);
    head.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF114C50" } };
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.alignment = { horizontal: "right" };
    });
    head.height = 29;
    if (settings.includeFacilityDetails) {
      for (const group of item.groups) for (const row of group.facilities) {
        const added = sheet.addRow([group.label, row.facility.name, row.score, row.entered, row.required, row.score !== null ? "مكتملة" : row.entered ? "جزئية" : "لم تُدخل", settings.includeNotes ? row.notes.map((note) => note.observation).filter(Boolean).join("؛ ") : "", settings.includeNotes ? row.notes.map((note) => note.action).filter(Boolean).join("؛ ") : ""]);
        added.getCell(3).numFmt = '0.0"%"';
        if (row.score === null) added.getCell(6).font = { color: { argb: "FF9B6B15" } };
      }
    } else {
      for (const group of item.groups) {
        const added = sheet.addRow([group.label, "ملخص التصنيف", group.mean, group.complete, group.total, group.mean !== null ? "متاح" : "ينتظر البيانات", "", ""]);
        added.getCell(3).numFmt = '0.0"%"';
      }
    }
    const comparisonRow = sheet.rowCount + 3;
    sheet.getCell(`A${comparisonRow}`).value = "مقارنة الفترات";
    sheet.getCell(`A${comparisonRow}`).font = { bold: true, color: { argb: "FF114C50" } };
    sheet.addRow(["الفترة", "المتوسط %", "المكتملة", "الجزئية", "التغير عن الفترة السابقة"]);
    const comparisonHead = sheet.getRow(comparisonRow + 1);
    comparisonHead.eachCell((cell) => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0C938E" } }; cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; });
    for (const point of item.trend) {
      const row = sheet.addRow([point.label, point.mean, point.complete, point.partial, point.change]);
      row.getCell(2).numFmt = '0.0"%"';
      row.getCell(5).numFmt = '0.0" نقطة"';
    }
    const image = workbook.addImage({ base64: chartPng(item, "group"), extension: "png" });
    const chartRow = sheet.rowCount + 3;
    sheet.addImage(image, { tl: { col: 0, row: chartRow }, ext: { width: 980, height: Math.min(620, Math.max(370, item.groups.length * 58 + 110)) } });
    sheet.autoFilter = { from: "A3", to: `H${Math.max(4, comparisonRow - 2)}` };
  }

  const raw = workbook.addWorksheet("المدخلات الأصلية", { views: [{ rightToLeft: true, state: "frozen", ySplit: 1 }] });
  raw.columns = [{ width: 19 }, { width: 18 }, { width: 29 }, { width: 18 }, { width: 17 }, { width: 17 }, { width: 22 }, { width: 48 }, { width: 48 }, { width: 23 }, { width: 17 }];
  raw.addRow(["البرنامج", "معرف المنشأة", "الفترة", "النسبة %", "البسط", "المقام", "وقت التحديث", "الملاحظة", "الإجراء", "المسؤول", "موعد المتابعة"]);
  raw.getRow(1).eachCell((cell) => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF114C50" } }; cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; });
  for (const entry of selectedEntries(scope, reports, entries)) raw.addRow([entry.program, entry.facilityId, entry.period, entry.score, entry.numerator, entry.denominator, entry.updatedAt, entry.observation, entry.action, entry.owner, entry.dueDate]);
  const buffer = await workbook.xlsx.writeBuffer();
  download(new Blob([buffer as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `وقاية-تقرير-${year}-${safeScope(scope)}.xlsx`);
}

async function logoData() {
  const response = await fetch("/wiqaya-logo.png");
  if (!response.ok) throw new Error("تعذر تحميل الشعار.");
  const blob = await response.blob();
  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

export async function exportPowerPoint(year: number, scope: ReportScope, reports: ProgramResult[], settings: ReportSettings) {
  const PptxGenJS = (await import("pptxgenjs")).default;
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "تقارير برامج وقاية";
  pptx.subject = `${periodLabel(scope)} ${year}`;
  pptx.title = settings.title;
  const logo = await logoData();
  let slideNumber = 0;
  const slideBase = (title: string, section?: string) => {
    const slide = pptx.addSlide();
    slideNumber += 1;
    slide.background = { color: "FFFFFF" };
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: .16, line: { color: COLORS.teal }, fill: { color: COLORS.teal } });
    slide.addImage({ data: logo, x: .45, y: .26, w: 1.65, h: .92 });
    if (section) slide.addText(section, { x: 2.3, y: .28, w: 2.7, h: .25, fontFace: "Arial", fontSize: 9, color: COLORS.teal, align: "right", rtlMode: true, margin: 0 });
    slide.addText(title, { x: 3.0, y: .50, w: 9.75, h: .5, fontFace: "Arial", fontSize: 23, bold: true, color: COLORS.dark, align: "right", rtlMode: true, margin: 0 });
    slide.addShape(pptx.ShapeType.line, { x: .55, y: 6.98, w: 12.2, h: 0, line: { color: "D8E6E3", pt: .8 } });
    slide.addText(`${settings.organization}  ·  ${periodLabel(scope)} ${year}`, { x: 2.0, y: 7.08, w: 10.5, h: .18, fontFace: "Arial", fontSize: 8.5, color: COLORS.gray, align: "right", rtlMode: true, margin: 0 });
    slide.addText(String(slideNumber), { x: .55, y: 7.08, w: .5, h: .18, fontFace: "Arial", fontSize: 8.5, color: COLORS.gray, align: "left", margin: 0 });
    return slide;
  };
  const metric = (slide: ReturnType<typeof slideBase>, x: number, y: number, w: number, label: string, value: string, color = COLORS.teal) => {
    slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h: .88, rectRadius: .08, line: { color: "D8E6E3", pt: .7 }, fill: { color: "F7FBFA" } });
    slide.addShape(pptx.ShapeType.rect, { x: x + w - .08, y, w: .08, h: .88, line: { color }, fill: { color } });
    slide.addText(label, { x: x + .12, y: y + .12, w: w - .3, h: .22, fontFace: "Arial", fontSize: 10, color: COLORS.gray, align: "right", rtlMode: true, margin: 0 });
    slide.addText(value, { x: x + .12, y: y + .39, w: w - .3, h: .31, fontFace: "Arial", fontSize: 19, bold: true, color: COLORS.dark, align: "right", rtlMode: true, margin: 0 });
  };

  const totalComplete = reports.reduce((sum, item) => sum + item.complete, 0);
  const totalFacilities = reports.reduce((sum, item) => sum + item.total, 0);
  const totalPartial = reports.reduce((sum, item) => sum + item.partial, 0);
  const title = slideBase(settings.title);
  title.addShape(pptx.ShapeType.roundRect, { x: 1.05, y: 1.62, w: 11.25, h: 3.75, rectRadius: .12, line: { color: "D8E6E3", pt: 1 }, fill: { color: "F4FAF8" } });
  title.addText(periodLabel(scope), { x: 1.5, y: 2.03, w: 10.3, h: .7, fontFace: "Arial", fontSize: 32, bold: true, color: COLORS.dark, align: "center", rtlMode: true, margin: 0 });
  title.addText(String(year), { x: 1.5, y: 2.82, w: 10.3, h: .52, fontFace: "Arial", fontSize: 25, color: COLORS.teal, align: "center", margin: 0 });
  title.addText(`يشمل ${reports.length} برامج و${totalFacilities} منشأة`, { x: 1.5, y: 3.62, w: 10.3, h: .42, fontFace: "Arial", fontSize: 17, color: COLORS.ink, align: "center", rtlMode: true, margin: 0 });
  title.addText(settings.organization, { x: 1.5, y: 4.35, w: 10.3, h: .38, fontFace: "Arial", fontSize: 15, color: COLORS.gray, align: "center", rtlMode: true, margin: 0 });

  const summary = slideBase("الملخص التنفيذي", "نظرة عامة");
  metric(summary, .65, 1.40, 2.85, "البرامج", String(reports.length));
  metric(summary, 3.73, 1.40, 2.85, "المنشآت المكتملة", `${totalComplete}/${totalFacilities}`);
  metric(summary, 6.81, 1.40, 2.85, "البيانات الجزئية", String(totalPartial), COLORS.gold);
  const availableMeans = reports.map((item) => item.mean).filter((value): value is number => value !== null);
  metric(summary, 9.89, 1.40, 2.85, "متوسط البرامج", formatPercent(availableMeans.length ? availableMeans.reduce((a, b) => a + b, 0) / availableMeans.length : null));
  summary.addShape(pptx.ShapeType.roundRect, { x: .75, y: 2.62, w: 11.85, h: 1.7, rectRadius: .08, line: { color: "D8E6E3", pt: .7 }, fill: { color: "FFFFFF" } });
  summary.addText(settings.executiveSummary, { x: 1.0, y: 2.92, w: 11.35, h: 1.1, fontFace: "Arial", fontSize: 17, color: COLORS.ink, align: "right", rtlMode: true, breakLine: false, margin: .04, valign: "middle" });
  if (settings.includeMethodology) summary.addText(`المنهج والتغطية\n${settings.methodology}`, { x: .85, y: 4.65, w: 11.6, h: 1.42, fontFace: "Arial", fontSize: 13.5, color: COLORS.dark, align: "right", rtlMode: true, breakLine: false, margin: .06, fill: { color: COLORS.pale }, valign: "middle" });

  const portfolio = slideBase("مقارنة البرامج", "نظرة عامة");
  portfolio.addImage({ data: barChartPng("متوسط البرامج", reports.map((item) => ({ label: item.label, value: item.mean }))), x: .55, y: 1.30, w: 6.55, h: 4.85 });
  const portfolioRows = [[{ text: "البرنامج", options: { bold: true, color: COLORS.white } }, { text: "المتوسط", options: { bold: true, color: COLORS.white } }, { text: "التغطية", options: { bold: true, color: COLORS.white } }], ...reports.map((item) => [{ text: item.label }, { text: formatPercent(item.mean) }, { text: `${item.complete}/${item.total}` }])];
  portfolio.addTable(portfolioRows, { x: 7.28, y: 1.52, w: 5.42, colW: [3.15, 1.08, 1.19], rowH: .48, fontFace: "Arial", fontSize: 11, color: COLORS.ink, fill: { color: "FFFFFF" }, border: { type: "solid", pt: .5, color: "D5E3E1" }, margin: .06, align: "right", bold: false });

  for (const item of reports) {
    const overview = slideBase(item.label, "تحليل البرنامج");
    metric(overview, .65, 1.30, 2.75, "المتوسط", formatPercent(item.mean));
    metric(overview, 3.62, 1.30, 2.75, "المكتملة", `${item.complete}/${item.total}`);
    metric(overview, 6.59, 1.30, 2.75, "الجزئية", String(item.partial), COLORS.gold);
    const validTrend = item.trend.filter((point) => point.mean !== null);
    const change = validTrend.length > 1 ? (validTrend.at(-1)?.mean ?? 0) - (validTrend[0].mean ?? 0) : null;
    metric(overview, 9.56, 1.30, 2.75, "التغير", formatChange(change), change !== null && change < 0 ? COLORS.red : COLORS.teal);
    overview.addImage({ data: chartPng(item, "trend"), x: .55, y: 2.48, w: 6.1, h: 3.10 });
    overview.addImage({ data: chartPng(item, "group"), x: 6.80, y: 2.48, w: 5.95, h: 3.10 });
    overview.addText(settings.programNotes[item.program] ?? reportNarrative(item, scope), { x: .75, y: 5.77, w: 11.8, h: .82, fontFace: "Arial", fontSize: 12.5, color: COLORS.ink, align: "right", rtlMode: true, margin: .05, breakLine: false, fill: { color: COLORS.pale }, valign: "middle" });

    const comparison = slideBase(`${item.label} · مقارنة الفترات`, "تحليل البرنامج");
    comparison.addImage({ data: chartPng(item, "trend"), x: .55, y: 1.34, w: 6.75, h: 4.55 });
    const trendRows = [[{ text: "الفترة", options: { bold: true, color: COLORS.white } }, { text: "المتوسط", options: { bold: true, color: COLORS.white } }, { text: "المكتملة", options: { bold: true, color: COLORS.white } }, { text: "التغير", options: { bold: true, color: COLORS.white } }], ...item.trend.map((point) => [{ text: point.label }, { text: formatPercent(point.mean) }, { text: `${point.complete}/${item.total}` }, { text: formatChange(point.change) }])];
    comparison.addTable(trendRows, { x: 7.52, y: 1.55, w: 5.18, colW: [1.82, 1.08, 1.18, 1.1], rowH: .48, fontFace: "Arial", fontSize: 10.5, color: COLORS.ink, fill: { color: "FFFFFF" }, border: { type: "solid", pt: .5, color: "D5E3E1" }, margin: .05, align: "right" });
    comparison.addText("تُحسب المقارنة من المنشآت المكتملة لكل فترة. قد يتغير عدد المنشآت بين الفترات بحسب اكتمال الإدخال.", { x: 7.60, y: 4.75, w: 4.95, h: .82, fontFace: "Arial", fontSize: 11.5, color: COLORS.gray, align: "right", rtlMode: true, margin: .04, fill: { color: "F7FBFA" } });

    const completeFacilities = item.groups.flatMap((group) => group.facilities.map((row) => ({ ...row, group: group.label }))).filter((row) => row.score !== null).sort((a, b) => (b.score ?? 0) - (a.score ?? 0));
    if (completeFacilities.length) {
      const highlights = slideBase(`${item.label} · أبرز النتائج`, "تحليل البرنامج");
      const top = completeFacilities.slice(0, 5);
      const low = [...completeFacilities].reverse().slice(0, 5);
      const topRows = [[{ text: "أعلى المنشآت", options: { bold: true, color: COLORS.white } }, { text: "النتيجة", options: { bold: true, color: COLORS.white } }], ...top.map((row) => [{ text: row.facility.name }, { text: formatPercent(row.score) }])];
      const lowRows = [[{ text: "المنشآت ذات الأولوية", options: { bold: true, color: COLORS.white } }, { text: "النتيجة", options: { bold: true, color: COLORS.white } }], ...low.map((row) => [{ text: row.facility.name }, { text: formatPercent(row.score) }])];
      highlights.addTable(topRows, { x: .72, y: 1.52, w: 5.72, colW: [4.35, 1.37], rowH: .53, fontFace: "Arial", fontSize: 12, color: COLORS.ink, fill: { color: "FFFFFF" }, border: { type: "solid", pt: .5, color: "D5E3E1" }, margin: .06, align: "right" });
      highlights.addTable(lowRows, { x: 6.88, y: 1.52, w: 5.72, colW: [4.35, 1.37], rowH: .53, fontFace: "Arial", fontSize: 12, color: COLORS.ink, fill: { color: "FFFFFF" }, border: { type: "solid", pt: .5, color: "D5E3E1" }, margin: .06, align: "right" });
      highlights.addText(`التغطية المكتملة: ${item.complete}/${item.total}. البيانات الجزئية: ${item.partial}. تُستخدم قائمة الأولوية لبدء المراجعة، مع الرجوع إلى الملاحظات والإجراءات المسجلة قبل اتخاذ القرار.`, { x: .85, y: 5.25, w: 11.6, h: .85, fontFace: "Arial", fontSize: 13.5, color: COLORS.dark, align: "right", rtlMode: true, margin: .05, fill: { color: COLORS.pale }, valign: "middle" });
    }

    if (settings.includeFacilityDetails) {
      const all = item.groups.flatMap((group) => group.facilities.map((row) => ({ ...row, groupLabel: group.label })));
      for (let start = 0; start < all.length; start += 13) {
        const detail = slideBase(`${item.label} · تفاصيل المنشآت ${start + 1}–${Math.min(all.length, start + 13)}`, "ملحق البيانات");
        const rows = [["المنشأة", "التصنيف", "النتيجة", "الاكتمال"].map((text) => ({ text, options: { bold: true, color: COLORS.white } })), ...all.slice(start, start + 13).map((row) => [row.facility.name, row.groupLabel, formatPercent(row.score), `${row.entered}/${row.required}`].map((text) => ({ text })))];
        detail.addTable(rows, { x: .66, y: 1.38, w: 12.0, colW: [4.35, 4.4, 1.35, 1.9], rowH: .39, fontFace: "Arial", fontSize: 9.7, color: COLORS.ink, fill: { color: "FFFFFF" }, border: { type: "solid", pt: .5, color: "D5E3E1" }, margin: .04, align: "right" });
      }
    }
  }

  const final = slideBase("الخلاصة والتوصيات", "الإغلاق");
  final.addShape(pptx.ShapeType.roundRect, { x: .85, y: 1.55, w: 11.65, h: 2.0, rectRadius: .08, line: { color: "D8E6E3", pt: .8 }, fill: { color: "F7FBFA" } });
  final.addText(settings.conclusion, { x: 1.12, y: 1.89, w: 11.1, h: 1.32, fontFace: "Arial", fontSize: 18, color: COLORS.ink, align: "right", rtlMode: true, margin: .05, valign: "middle" });
  final.addText("مراجعة", { x: 7.05, y: 4.62, w: 2.1, h: .35, fontFace: "Arial", fontSize: 14, bold: true, color: COLORS.dark, align: "right", rtlMode: true, margin: 0 });
  final.addText(settings.reviewedBy || "____________________", { x: 7.05, y: 5.12, w: 2.7, h: .35, fontFace: "Arial", fontSize: 14, color: COLORS.ink, align: "right", rtlMode: true, margin: 0 });
  final.addText("اعتماد", { x: 10.05, y: 4.62, w: 2.1, h: .35, fontFace: "Arial", fontSize: 14, bold: true, color: COLORS.dark, align: "right", rtlMode: true, margin: 0 });
  final.addText(settings.approvedBy || "____________________", { x: 10.05, y: 5.12, w: 2.55, h: .35, fontFace: "Arial", fontSize: 14, color: COLORS.ink, align: "right", rtlMode: true, margin: 0 });
  const output = await pptx.write({ outputType: "blob" });
  download(output as Blob, `وقاية-عرض-${year}-${safeScope(scope)}.pptx`);
}
