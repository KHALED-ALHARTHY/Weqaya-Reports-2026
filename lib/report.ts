import { FACILITIES, GROUP_LABELS, MONTHS, PROGRAMS, programName, type Facility, type ProgramId } from "./catalog";

export type Entry = {
  id: number;
  year: number;
  program: ProgramId;
  facilityId: string;
  period: string;
  score: number;
  numerator: number | null;
  denominator: number | null;
  observation: string;
  action: string;
  owner: string;
  dueDate: string;
  updatedAt: string;
};

export type Scope = "year" | `Q${1 | 2 | 3 | 4}` | "H1" | "H2" | `M${string}`;
export type ReportScope = Scope | `compare:${string}`;
export type FacilityResult = {
  facility: Facility;
  score: number | null;
  entered: number;
  required: number;
  values: (number | null)[];
  notes: Entry[];
};
export type GroupResult = {
  group: string;
  label: string;
  mean: number | null;
  complete: number;
  partial: number;
  total: number;
  facilities: FacilityResult[];
};
export type TrendPoint = {
  scope: Scope;
  label: string;
  mean: number | null;
  complete: number;
  partial: number;
  change: number | null;
};
export type ProgramResult = {
  program: ProgramId;
  label: string;
  applicable: boolean;
  mean: number | null;
  complete: number;
  partial: number;
  total: number;
  groups: GroupResult[];
  periodLabels: string[];
  trend: TrendPoint[];
  observations: string[];
  actions: string[];
};

export const SCOPE_OPTIONS: { value: Scope; label: string }[] = [
  { value: "year", label: "التقرير السنوي" },
  ...[1, 2, 3, 4].map((n) => ({
    value: `Q${n}` as Scope,
    label: `الربع ${["الأول", "الثاني", "الثالث", "الرابع"][n - 1]}`,
  })),
  { value: "H1", label: "النصف الأول" },
  { value: "H2", label: "النصف الثاني" },
  ...MONTHS.map((label, i) => ({ value: `M${String(i + 1).padStart(2, "0")}` as Scope, label })),
];

const allMonths = MONTHS.map((_, i) => `M${String(i + 1).padStart(2, "0")}` as Scope);
const quarterScopes = ["Q1", "Q2", "Q3", "Q4"] as Scope[];

export function reportScopes(scope: ReportScope): Scope[] {
  if (!scope.startsWith("compare:")) return [scope as Scope];
  const values = scope.slice(8).split(",").filter((value): value is Scope => /^Q[1-4]$/.test(value));
  return [...new Set(values)].sort() as Scope[];
}

export function createComparisonScope(scopes: Scope[]): ReportScope {
  const values = [...new Set(scopes.filter((value) => /^Q[1-4]$/.test(value)))].sort();
  if (values.length === 1) return values[0];
  return `compare:${values.join(",")}` as ReportScope;
}

export function isValidReportScope(value: string): value is ReportScope {
  if (SCOPE_OPTIONS.some((item) => item.value === value)) return true;
  const scopes = reportScopes(value as ReportScope);
  return value.startsWith("compare:") && scopes.length >= 2 && scopes.length <= 4;
}

function requiredForSingle(program: ProgramId, scope: Scope): string[] | null {
  if (scope === "year") {
    return program === "hh" ? allMonths : program === "phc_ica" ? ["H1", "H2"] : quarterScopes;
  }
  if (/^Q[1-4]$/.test(scope)) {
    if (program === "phc_ica") return null;
    if (program !== "hh") return [scope];
    const quarter = Number(scope[1]);
    return allMonths.slice((quarter - 1) * 3, quarter * 3);
  }
  if (scope === "H1" || scope === "H2") {
    if (program === "phc_ica") return [scope];
    const first = scope === "H1" ? 0 : 6;
    return program === "hh" ? allMonths.slice(first, first + 6) : scope === "H1" ? ["Q1", "Q2"] : ["Q3", "Q4"];
  }
  return program === "hh" && allMonths.includes(scope) ? [scope] : null;
}

export function requiredPeriods(program: ProgramId, scope: ReportScope): string[] | null {
  const selections = reportScopes(scope);
  const groups = selections.map((selected) => requiredForSingle(program, selected));
  if (!groups.length || groups.some((group) => group === null)) return null;
  return [...new Set(groups.flatMap((group) => group ?? []))];
}

const average = (values: number[]) => values.length ? values.reduce((a, b) => a + b, 0) / values.length : null;

function compute(program: ProgramId, scope: ReportScope, entries: Entry[], includeTrend: boolean): ProgramResult {
  const required = requiredPeriods(program, scope);
  const facilities = FACILITIES.filter((item) => item.program === program);
  const lookup = new Map(entries.filter((item) => item.program === program).map((item) => [`${item.facilityId}|${item.period}`, item]));
  const groupNames = [...new Set(facilities.map((item) => item.group))];
  const groups = groupNames.map((group): GroupResult => {
    const rows = facilities.filter((item) => item.group === group).map((facility): FacilityResult => {
      const found = (required ?? []).map((key) => lookup.get(`${facility.id}|${key}`));
      const scores = found.filter((item): item is Entry => !!item).map((item) => item.score);
      return {
        facility,
        score: required && scores.length === required.length ? average(scores) : null,
        entered: scores.length,
        required: required?.length ?? 0,
        values: found.map((item) => item?.score ?? null),
        notes: found.filter((item): item is Entry => !!item),
      };
    });
    const completeValues = rows.map((item) => item.score).filter((n): n is number => n !== null);
    return {
      group,
      label: GROUP_LABELS[group] ?? group,
      mean: average(completeValues),
      complete: completeValues.length,
      partial: rows.filter((item) => item.entered > 0 && item.score === null).length,
      total: rows.length,
      facilities: rows,
    };
  });
  const allResults = groups.flatMap((group) => group.facilities);
  const completeValues = allResults.map((item) => item.score).filter((n): n is number => n !== null);
  const relevant = allResults.flatMap((item) => item.notes);
  const uniqueText = (values: string[]) => [...new Set(values.map((text) => text.trim()).filter(Boolean))].slice(0, 8);
  const selected = reportScopes(scope);
  let trendScopes: Scope[];
  if (scope.startsWith("compare:")) trendScopes = selected;
  else if (scope === "year") trendScopes = program === "phc_ica" ? ["H1", "H2"] : quarterScopes;
  else if (scope === "H1") trendScopes = program === "phc_ica" ? ["H1"] : ["Q1", "Q2"];
  else if (scope === "H2") trendScopes = program === "phc_ica" ? ["H2"] : ["Q3", "Q4"];
  else if (program === "hh" && scope.startsWith("M")) trendScopes = allMonths;
  else trendScopes = quarterScopes;
  let previousMean: number | null = null;
  const trend = includeTrend ? trendScopes.map((selectedScope): TrendPoint => {
    const next = compute(program, selectedScope, entries, false);
    const change = previousMean !== null && next.mean !== null ? next.mean - previousMean : null;
    if (next.mean !== null) previousMean = next.mean;
    return {
      scope: selectedScope,
      label: periodLabel(selectedScope),
      mean: next.mean,
      complete: next.complete,
      partial: next.partial,
      change,
    };
  }) : [];
  return {
    program,
    label: programName(program),
    applicable: required !== null,
    mean: average(completeValues),
    complete: completeValues.length,
    partial: allResults.filter((item) => item.entered > 0 && item.score === null).length,
    total: facilities.length,
    groups,
    periodLabels: required ?? [],
    trend,
    observations: uniqueText(relevant.map((item) => item.observation)),
    actions: uniqueText(relevant.map((item) => item.action)),
  };
}

export function buildReport(scope: ReportScope, entries: Entry[]) {
  return PROGRAMS.map((program) => compute(program.id, scope, entries, true)).filter((item) => item.applicable);
}

export function periodLabel(scope: ReportScope): string {
  const selected = reportScopes(scope);
  if (scope.startsWith("compare:")) return selected.map((item) => periodLabel(item)).join(" و");
  return SCOPE_OPTIONS.find((item) => item.value === scope)?.label ?? scope;
}

export function formatPercent(value: number | null, digits = 1) {
  return value === null ? "—" : `${value.toFixed(digits)}%`;
}

export function formatChange(value: number | null) {
  if (value === null) return "—";
  const sign = value > 0 ? "+" : "";
  return `${sign}${value.toFixed(1)} نقطة`;
}

export function reportNarrative(item: ProgramResult, scope: ReportScope) {
  if (item.complete === 0) {
    return `لا توجد نتائج مكتملة لبرنامج ${item.label} خلال ${periodLabel(scope)}. توجد بيانات جزئية لدى ${item.partial} منشأة، ويجب استكمال الفترات المطلوبة قبل اعتماد المقارنة.`;
  }
  const rankedGroups = item.groups.filter((group) => group.mean !== null).sort((a, b) => (b.mean ?? 0) - (a.mean ?? 0));
  const best = rankedGroups[0];
  const lowest = rankedGroups.at(-1);
  const validTrend = item.trend.filter((point) => point.mean !== null);
  const first = validTrend[0];
  const last = validTrend[validTrend.length - 1];
  let direction = "";
  if (first && last && validTrend.length > 1) {
    const change = (last.mean ?? 0) - (first.mean ?? 0);
    direction = change > 0
      ? `ارتفع الأداء بين ${first.label} و${last.label} بمقدار ${formatChange(change)}. `
      : change < 0
        ? `انخفض الأداء بين ${first.label} و${last.label} بمقدار ${formatChange(Math.abs(change))}. `
        : `استقر الأداء بين ${first.label} و${last.label}. `;
  }
  return `بلغ متوسط ${item.label} ${formatPercent(item.mean)} لدى ${item.complete} من أصل ${item.total} منشأة مكتملة البيانات. ${direction}${item.partial ? `توجد بيانات جزئية لدى ${item.partial} منشأة لا تدخل في المتوسط. ` : ""}${best ? `أعلى تصنيف هو ${best.label} بمتوسط ${formatPercent(best.mean)}. ` : ""}${lowest && lowest !== best ? `أقل تصنيف هو ${lowest.label} بمتوسط ${formatPercent(lowest.mean)}. ` : ""}المتوسط غير موزون بعدد فرص الرصد.`;
}
