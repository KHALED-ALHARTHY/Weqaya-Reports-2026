import rawFacilities from "./facilities.json";

export const PROGRAMS = [
  { id: "ipccc", name: "برنامج مكافحة العدوى IPCCC", cadence: "quarter" },
  { id: "hh", name: "تقييم نظافة اليدين", cadence: "month" },
  { id: "rpp", name: "برنامج الوقاية التنفسية RPP", cadence: "quarter" },
  { id: "phc_ica", name: "التقييم الذاتي لمراكز الرعاية PHC ICA", cadence: "half" },
] as const;

export type ProgramId = (typeof PROGRAMS)[number]["id"];
export type Facility = { id: string; program: ProgramId; group: string; name: string; arabicName?: string };
export const FACILITIES = rawFacilities as Facility[];
export const FACILITY_BY_ID = new Map(FACILITIES.map((item) => [item.id, item]));
export const FACILITY_ID_ALIASES: Record<string, string> = {
  "hh-090": "hh-064",
  "hh-127": "hh-065",
  "hh-178": "hh-066",
};

export const MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
export const PERIODS: Record<"quarter" | "month" | "half", { key: string; label: string }[]> = {
  quarter: [1, 2, 3, 4].map((n) => ({ key: `Q${n}`, label: `الربع ${["الأول", "الثاني", "الثالث", "الرابع"][n - 1]}` })),
  month: MONTHS.map((label, index) => ({ key: `M${String(index + 1).padStart(2, "0")}`, label })),
  half: [{ key: "H1", label: "النصف الأول" }, { key: "H2", label: "النصف الثاني" }],
};
export const GROUP_LABELS: Record<string, string> = {
  "Government MOH Hospitals": "المستشفيات الحكومية التابعة لوزارة الصحة",
  "Private Hospitals": "المستشفيات الخاصة",
  "Non-MOH Government Hospital": "مستشفى حكومي خارج وزارة الصحة",
  "Dental Centers": "مراكز الأسنان",
  "HD Private Centers": "مراكز الغسيل الكلوي الخاصة",
  "HD Charity Centers": "مراكز الغسيل الكلوي الخيرية",
  PHCs: "مراكز الرعاية الصحية الأولية",
};
export function programName(id: ProgramId) { return PROGRAMS.find((item) => item.id === id)?.name ?? id; }
