"use client";

import { FACILITIES, type ProgramId } from "./catalog";

export type ImportedEntry = {
  program: ProgramId;
  facilityId: string;
  facilityName: string;
  period: string;
  score: number;
  sourceSheet: string;
  sourceRow: number;
};

export type ImportPreview = {
  entries: ImportedEntry[];
  unmatched: string[];
  skippedZeros: number;
  duplicateValues: number;
  sheets: string[];
};

const monthMap: Record<string, string> = {
  jan: "M01", january: "M01", يناير: "M01",
  feb: "M02", february: "M02", فبراير: "M02",
  mar: "M03", march: "M03", مارس: "M03",
  apr: "M04", april: "M04", ابريل: "M04", أبريل: "M04",
  may: "M05", مايو: "M05",
  jun: "M06", june: "M06", يونيو: "M06",
  jul: "M07", july: "M07", يوليو: "M07",
  aug: "M08", august: "M08", اغسطس: "M08", أغسطس: "M08",
  sep: "M09", sept: "M09", september: "M09", سبتمبر: "M09",
  oct: "M10", october: "M10", اكتوبر: "M10", أكتوبر: "M10",
  nov: "M11", november: "M11", نوفمبر: "M11",
  dec: "M12", december: "M12", ديسمبر: "M12",
};

function westernDigits(value: string) {
  return value.replace(/[٠-٩]/g, (digit) => String("٠١٢٣٤٥٦٧٨٩".indexOf(digit)))
    .replace(/[۰-۹]/g, (digit) => String("۰۱۲۳۴۵۶۷۸۹".indexOf(digit)));
}

function normalize(value: unknown) {
  return westernDigits(String(value ?? ""))
    .trim()
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9\u0600-\u06ff]+/g, " ")
    .replace(/\bcentre\b/g, "center")
    .replace(/\s+/g, " ")
    .trim();
}

const facilityLookup = new Map<string, (typeof FACILITIES)[number]>();
for (const facility of FACILITIES) {
  for (const candidate of [facility.name, facility.arabicName]) {
    const key = normalize(candidate);
    if (key && !facilityLookup.has(`${facility.program}|${key}`)) facilityLookup.set(`${facility.program}|${key}`, facility);
  }
}

const sourceFacilityAliases = new Map<string, (typeof FACILITIES)[number]>();
for (const alias of [
  { program: "hh", sourceName: "King Abdul-Aziz Hospital", sourceGroup: "Dental Centers", facilityId: "hh-046" },
  { program: "hh", sourceName: "Al-Noor Specialist Hospital", sourceGroup: "Dental Centers", facilityId: "hh-047" },
] as const) {
  const facility = FACILITIES.find((item) => item.id === alias.facilityId);
  if (facility) sourceFacilityAliases.set(`${alias.program}|${normalize(alias.sourceName)}|${normalize(alias.sourceGroup)}`, facility);
}

function unwrap(value: unknown): unknown {
  if (value && typeof value === "object" && "result" in value) return unwrap((value as { result?: unknown }).result);
  if (value && typeof value === "object" && "text" in value) return (value as { text?: unknown }).text;
  return value;
}

function scoreValue(value: unknown): number | null {
  const raw = unwrap(value);
  if (raw === null || raw === undefined || raw === "") return null;
  if (typeof raw === "number") {
    if (!Number.isFinite(raw)) return null;
    return raw <= 1 ? raw * 100 : raw;
  }
  const text = westernDigits(String(raw)).replace(/[٪%]/g, "").replace(/[٫,]/g, ".").replace(/\s/g, "");
  const parsed = Number(text);
  if (!Number.isFinite(parsed)) return null;
  return parsed <= 1 && !String(raw).includes("%") && !String(raw).includes("٪") ? parsed * 100 : parsed;
}

function programFromSheet(name: string): ProgramId | null {
  const key = normalize(name);
  if (key.includes("ipccc")) return "ipccc";
  if (/\brpp\b/.test(key)) return "rpp";
  if (key.includes("phc") && (key.includes("ica") || key.includes("self assess"))) return "phc_ica";
  if (key.includes("h h") || key.includes("hand hygiene")) return "hh";
  return null;
}

function periodFromHeader(value: unknown, program: ProgramId): string | null {
  const key = normalize(unwrap(value));
  if (!key) return null;
  if (program === "hh") {
    const direct = monthMap[key];
    if (direct) return direct;
    return null;
  }
  if (program === "phc_ica") {
    if (/(^| )1(st)? half( |$)/.test(key) || key.includes("النصف الاول") || key.includes("النصف الأول")) return "H1";
    if (/(^| )2(nd)? half( |$)/.test(key) || key.includes("النصف الثاني")) return "H2";
    return null;
  }
  if (/(^| )(1st|quarter 1|quarterly 1|q1)( |$)/.test(key) || key.includes("الربع الاول") || key.includes("الربع الأول")) return "Q1";
  if (/(^| )(2nd|quarter 2|quarterly 2|q2)( |$)/.test(key) || key.includes("الربع الثاني")) return "Q2";
  if (/(^| )(3rd|quarter 3|quarterly 3|q3)( |$)/.test(key) || key.includes("الربع الثالث")) return "Q3";
  if (/(^| )(4th|4rd|quarter 4|quarterly 4|q4)( |$)/.test(key) || key.includes("الربع الرابع")) return "Q4";
  return null;
}

function isFacilityHeader(value: unknown) {
  const key = normalize(unwrap(value));
  return key.includes("name of the hospital") || key.includes("name of the phc") || key === "facility" || key === "facility name" || key.includes("اسم المنشأة") || key.includes("اسم المركز");
}

export async function parseExcelFile(file: File): Promise<ImportPreview> {
  const ExcelJS = (await import("exceljs")).default;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const entries: ImportedEntry[] = [];
  const unmatched = new Set<string>();
  const sheets: string[] = [];
  const seen = new Set<string>();
  let skippedZeros = 0;
  let duplicateValues = 0;

  for (const worksheet of workbook.worksheets) {
    const program = programFromSheet(worksheet.name);
    if (!program) continue;
    let headerRow = 0;
    let facilityColumn = 0;
    const periodColumns = new Map<number, string>();
    for (let rowNumber = 1; rowNumber <= Math.min(12, worksheet.rowCount); rowNumber += 1) {
      const row = worksheet.getRow(rowNumber);
      const candidatePeriods = new Map<number, string>();
      let candidateFacility = 0;
      for (let column = 1; column <= Math.min(worksheet.columnCount, 40); column += 1) {
        const value = row.getCell(column).value;
        if (isFacilityHeader(value)) candidateFacility = column;
        const period = periodFromHeader(value, program);
        if (period) candidatePeriods.set(column, period);
      }
      if (candidateFacility && candidatePeriods.size) {
        headerRow = rowNumber;
        facilityColumn = candidateFacility;
        for (const item of candidatePeriods) periodColumns.set(...item);
        break;
      }
    }
    if (!headerRow || !facilityColumn || !periodColumns.size) continue;
    sheets.push(worksheet.name.trim());
    for (let rowNumber = headerRow + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      const row = worksheet.getRow(rowNumber);
      const name = String(unwrap(row.getCell(facilityColumn).value) ?? "").trim();
      if (!name) continue;
      const sourceGroup = normalize(row.getCell(1).value);
      const facility = sourceFacilityAliases.get(`${program}|${normalize(name)}|${sourceGroup}`) ?? facilityLookup.get(`${program}|${normalize(name)}`);
      const hasScore = [...periodColumns.keys()].some((column) => scoreValue(row.getCell(column).value) !== null);
      if (!facility) {
        if (hasScore) unmatched.add(`${worksheet.name.trim()}: ${name}`);
        continue;
      }
      for (const [column, period] of periodColumns) {
        const score = scoreValue(row.getCell(column).value);
        if (score === null || score < 0 || score > 100) continue;
        if (score === 0) {
          skippedZeros += 1;
          continue;
        }
        const key = `${program}|${facility.id}|${period}`;
        if (seen.has(key)) {
          duplicateValues += 1;
          continue;
        }
        seen.add(key);
        entries.push({ program, facilityId: facility.id, facilityName: facility.name, period, score: Math.round(score * 100) / 100, sourceSheet: worksheet.name.trim(), sourceRow: rowNumber });
      }
    }
  }
  return { entries, unmatched: [...unmatched].sort(), skippedZeros, duplicateValues, sheets };
}
