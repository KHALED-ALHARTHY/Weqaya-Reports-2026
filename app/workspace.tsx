"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Activity, AlertTriangle, CheckCircle2, FileDown, FileSpreadsheet, FileText, PencilLine, Presentation, RotateCcw, Save, Search, Trash2, Upload } from "lucide-react";
import { Bar, BarChart, CartesianGrid, Line, LineChart, XAxis, YAxis } from "recharts";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import { AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger } from "@/components/ui/alert-dialog";
import { ChartContainer, ChartTooltip, ChartTooltipContent } from "@/components/ui/chart";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { FACILITIES, GROUP_LABELS, MONTHS, PERIODS, PROGRAMS, type ProgramId } from "@/lib/catalog";
import { buildReport, createComparisonScope, formatChange, formatPercent, periodLabel, reportNarrative, type Entry, type ProgramResult, type ReportScope, type Scope } from "@/lib/report";
import { defaultReportSettings, type ReportSettings } from "@/lib/report-settings";
import type { ImportPreview } from "@/lib/excel-import";

type Mode = "entry" | "dashboard";
type ReportMode = "year" | "quarters" | "H1" | "H2" | "month";
type Form = { score: string; numerator: string; denominator: string; observation: string; action: string; owner: string; dueDate: string };
const emptyForm: Form = { score: "", numerator: "", denominator: "", observation: "", action: "", owner: "", dueDate: "" };
const quarterOptions = [
  { value: "Q1" as Scope, label: "الربع الأول" },
  { value: "Q2" as Scope, label: "الربع الثاني" },
  { value: "Q3" as Scope, label: "الربع الثالث" },
  { value: "Q4" as Scope, label: "الربع الرابع" },
];

function SelectField({ label, value, onChange, options }: { label: string; value: string; onChange: (value: string) => void; options: { value: string; label: string }[] }) {
  return <div className="field"><label>{label}</label><Select value={value} onValueChange={onChange}><SelectTrigger aria-label={label}><SelectValue /></SelectTrigger><SelectContent>{options.map((option) => <SelectItem key={option.value} value={option.value}>{option.label}</SelectItem>)}</SelectContent></Select></div>;
}

function Header({ mode }: { mode: Mode }) {
  return <header className="site-header"><div className="site-header-inner"><div className="brand"><img src="/wiqaya-logo.png" alt="شعار وقاية" /><div className="brand-title">تقارير البرامج<small>هيئة الصحة العامة · مكتب مكة المكرمة</small></div></div><nav className="site-nav" aria-label="التنقل الرئيسي"><a href="/" className={mode === "entry" ? "active" : ""}>إدخال النتائج</a><a href="/dashboard" className={mode === "dashboard" ? "active" : ""}>لوحة المتابعة والتقارير</a></nav></div></header>;
}

function ErrorNotice({ message }: { message: string }) {
  const staleModule = /Failed to fetch dynamically imported module|Importing a module script failed|Loading chunk .* failed|error loading dynamically imported module/i.test(message);
  return <div className="notice error" role="alert">{staleModule ? <div>
    <p>تعذر تحميل أداة الاستيراد أو التصدير. قد تكون الصفحة مفتوحة من قبل تحديث الموقع، أو انقطع الاتصال.</p>
    <p>تأكد من الاتصال، ثم حدّث الصفحة وأعد اختيار الملف أو التصدير. احفظ أي تعديلات غير محفوظة قبل التحديث؛ بياناتك المحفوظة ستبقى.</p>
    <Button variant="outline" onClick={() => window.location.reload()}><RotateCcw size={16} />تحديث الصفحة</Button>
  </div> : message}</div>;
}

function ReportChart({ item, type }: { item: ProgramResult; type: "trend" | "groups" }) {
  const data = type === "trend" ? item.trend.map((point) => ({ name: point.label, value: point.mean, coverage: point.complete })) : item.groups.map((group) => ({ name: group.label, value: group.mean, coverage: group.complete }));
  if (!data.some((point) => point.value !== null)) return <div className="empty"><strong>الرسم ينتظر إدخال النتائج</strong>سيظهر تلقائيًا بعد اكتمال بيانات الفترة.</div>;
  return <ChartContainer config={{ value: { label: "المتوسط", color: "#0c938e" } }} className="h-[250px] w-full">
    {type === "trend" ? <LineChart accessibilityLayer data={data} margin={{ top: 15, right: 20, left: 12, bottom: 12 }}><CartesianGrid vertical={false} stroke="#dce9e6" /><XAxis dataKey="name" tick={{ fontSize: 12 }} /><YAxis domain={[0, 100]} tick={{ fontSize: 12 }} tickFormatter={(value) => `${value}%`} /><ChartTooltip content={<ChartTooltipContent formatter={(value) => `${Number(value).toFixed(1)}%`} />} /><Line dataKey="value" type="monotone" stroke="#0c938e" strokeWidth={3} dot={{ r: 4, fill: "#0c938e" }} connectNulls={false} /></LineChart>
      : <BarChart accessibilityLayer data={data} margin={{ top: 15, right: 18, left: 12, bottom: 45 }}><CartesianGrid vertical={false} stroke="#dce9e6" /><XAxis dataKey="name" tick={{ fontSize: 11 }} interval={0} angle={-18} textAnchor="end" height={75} /><YAxis domain={[0, 100]} tick={{ fontSize: 12 }} tickFormatter={(value) => `${value}%`} /><ChartTooltip content={<ChartTooltipContent formatter={(value) => `${Number(value).toFixed(1)}%`} />} /><Bar dataKey="value" fill="#0c938e" radius={[5, 5, 0, 0]} /></BarChart>}
  </ChartContainer>;
}

function ImportPanel({ year, onImported }: { year: number; onImported: () => Promise<void> }) {
  const [preview, setPreview] = useState<ImportPreview | null>(null);
  const [fileName, setFileName] = useState("");
  const [reading, setReading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [reportScope, setReportScope] = useState("H1");
  const [reportProgram, setReportProgram] = useState("all");
  const [quarters, setQuarters] = useState(["Q1", "Q2"]);
  const [skipZeros, setSkipZeros] = useState(true);
  const [progress, setProgress] = useState("");
  const [download, setDownload] = useState<{ url: string; name: string } | null>(null);
  const worker = useRef<Worker | null>(null);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const objectUrl = useRef<string | null>(null);
  const jobYear = useRef(year);
  function stop() {
    worker.current?.terminate(); worker.current = null;
    if (timer.current) clearTimeout(timer.current);
    setReading(false);
  }
  useEffect(() => () => { worker.current?.terminate(); if (timer.current) clearTimeout(timer.current); if (objectUrl.current) URL.revokeObjectURL(objectUrl.current); }, []);
  useEffect(() => { stop(); setPreview(null); setDownload(null); setFileName(""); }, [year]);

  async function readFile(file: File) {
    stop();
    setReading(true);
    setPreview(null);
    setMessage("");
    setError("");
    setFileName(file.name);
    setDownload(null);
    if (objectUrl.current) URL.revokeObjectURL(objectUrl.current);
    jobYear.current = year;
    try {
      if (!file.name.toLowerCase().endsWith(".xlsx")) throw new Error("اختر ملف Excel بصيغة XLSX.");
      if (file.size > 12 * 1024 * 1024) throw new Error("الحد الأقصى للملف 12 MB.");
      if (reportScope === "custom" && !quarters.length) throw new Error("اختر ربعًا واحدًا على الأقل.");
      const selectedScope = reportScope === "custom" ? `compare:${[...quarters].sort().join(",")}` : reportScope;
      const engine = new Worker("/report-worker.js");
      worker.current = engine;
      engine.onmessage = ({ data }) => {
        if (worker.current !== engine) return;
        if (data.type === "progress") setProgress(data.message);
        if (data.type === "error") { setError(data.message); stop(); }
        if (data.type === "done") {
          const url = URL.createObjectURL(new Blob([data.buffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
          objectUrl.current = url;
          setDownload({ url, name: `Weqaya-${jobYear.current}-${selectedScope.replace(/[:,]/g, "-")}.pptx` });
          setPreview(data.preview); stop();
        }
      };
      engine.onerror = () => { setError("تعذر بدء محرك التقرير. حدّث الصفحة وأعد المحاولة."); stop(); };
      timer.current = setTimeout(() => { setError("استغرق التحميل وقتًا طويلًا. تحقق من الاتصال وأعد اختيار الملف."); stop(); }, 600000);
      setProgress("تجهيز الملف للتحليل");
      const buffer = await file.arrayBuffer();
      if (worker.current === engine) engine.postMessage({ type: "generate", buffer, year, scope: selectedScope, program: reportProgram, skipZeros }, [buffer]);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "تعذر قراءة ملف Excel.");
      stop();
    }
  }

  async function saveImport() {
    if (!preview?.entries.length) return;
    setSaving(true);
    setError("");
    setMessage("");
    try {
      const response = await fetch("/api/entries/import", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ year: jobYear.current, entries: preview.entries }) });
      const body = await response.json() as { saved?: number; error?: string };
      if (!response.ok) throw new Error(body.error || "تعذر حفظ نتائج الملف.");
      await onImported();
      setMessage(`تم استيراد ${body.saved ?? preview.entries.length} نتيجة إلى سنة ${year}.`);
      setPreview(null);
      setFileName("");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "تعذر حفظ نتائج الملف.");
    } finally {
      setSaving(false);
    }
  }

  return <section className="panel panel-padding import-panel">
    <div className="section-title"><div><h2><FileSpreadsheet size={21} /> استيراد النتائج من Excel</h2><p>من ملف المتابعة إلى عرض رسمي قابل للتحرير · سنة {year}</p></div><span className="engine-badge">مركز إعداد التقارير</span></div>
    <ol className="report-steps"><li>1 · حدد نطاق التقرير</li><li>2 · اختر ملف Tracker</li><li>3 · راجع النتائج وحمّل العرض</li></ol>
    <fieldset className="engine-controls" disabled={reading || saving}><legend>إعداد العرض</legend><SelectField label="فترة التقرير" value={reportScope} onChange={setReportScope} options={[...quarterOptions, {value:"H1",label:"النصف الأول"},{value:"H2",label:"النصف الثاني"},{value:"year",label:"السنة كاملة"},{value:"custom",label:"مقارنة أرباع مختارة"}]} /><SelectField label="البرامج في العرض" value={reportProgram} onChange={setReportProgram} options={[{value:"all",label:"جميع البرامج"},...PROGRAMS.map(p=>({value:p.id,label:p.name}))]} />
      {reportScope === "custom" && <div className="engine-quarters">{quarterOptions.map(q=><label key={q.value}><input type="checkbox" checked={quarters.includes(q.value)} onChange={e=>setQuarters(v=>e.target.checked?[...v,q.value]:v.filter(x=>x!==q.value))} />{q.label}</label>)}</div>}
      <label className="engine-zero"><input type="checkbox" checked={skipZeros} onChange={e=>setSkipZeros(e.target.checked)} />اعتبار أصفار النموذج بيانات غير مدخلة (ألغِ الاختيار إذا كانت نتائج فعلية)</label>
    </fieldset>
    <p className="muted">يبدأ إنشاء العرض عند اختيار الملف. في التشغيل الأول تُحمّل مكتبات التحليل وقد يستغرق ذلك عدة دقائق. يُحلّل الملف على جهازك؛ الحفظ في الموقع يتم بعد مراجعتك. إعدادات الفترة والبرنامج تخص العرض، والاستيراد يشمل جميع النتائج المطابقة في الملف. لتغيير السنة استخدم حقل السنة في صفحة الإدخال.</p>
    <div className="import-drop"><Upload size={28} /><div><strong>{fileName || "اختر ملف Excel"}</strong><span>تُقرأ البيانات داخل المتصفح، ولا يُحفظ الملف نفسه.</span></div><label className="file-button">{reading ? "جارٍ التحليل" : "اختيار الملف"}<input type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" disabled={reading || saving} onChange={(event) => { const file = event.target.files?.[0]; if (file) void readFile(file); event.currentTarget.value = ""; }} /></label></div>
    {error && <ErrorNotice message={error} />}
    {reading && <div className="engine-progress" role="status"><span className="engine-pulse" /><div><strong>{progress}</strong><p>يمكنك متابعة العمل بعد اكتمال التحليل؛ لا تغلق الصفحة.</p></div><Button variant="outline" onClick={() => { stop(); setMessage("تم إلغاء إنشاء التقرير. لم تُحفظ أي نتائج."); }}>إلغاء التحليل</Button></div>}
    {download && <div className="engine-result"><div><strong><Presentation size={20} /> العرض جاهز للمراجعة</strong><p>جداول تفصيلية، رسوم قابلة للتحرير، مقارنة الفترات وشروح الحساب. يمكنك تعديل النصوص في PowerPoint.</p></div><a className="file-button" href={download.url} download={download.name}><FileDown size={18} />تحميل PowerPoint</a></div>}
    {message && <div className="notice" role="status"><CheckCircle2 size={18} />{message}</div>}
    {preview && <div className="import-preview">
      <div className="metric-grid import-metrics"><div className="metric"><small>نتائج جاهزة</small><strong>{preview.entries.length}</strong></div><div className="metric"><small>أوراق تمت قراءتها</small><strong>{preview.sheets.length}</strong></div><div className="metric"><small>قيم صفرية متروكة</small><strong>{preview.skippedZeros}</strong></div><div className="metric"><small>قيم مكررة متروكة</small><strong>{preview.duplicateValues}</strong></div></div>
      <p className="muted">الأوراق: {preview.sheets.join("، ")}. تمت معالجة الأصفار حسب الخيار المحدد قبل التحليل. الحفظ يحدّث نتائج المنشآت والفترات الموجودة في الملف لسنة {jobYear.current}.</p>
      {preview.unmatched.length > 0 && <details><summary><AlertTriangle size={17} /> منشآت لم تتطابق ({preview.unmatched.length})</summary><ul>{preview.unmatched.slice(0, 30).map((item) => <li key={item}>{item}</li>)}</ul>{preview.unmatched.length > 30 && <p>وتوجد {preview.unmatched.length - 30} منشأة أخرى.</p>}</details>}
      <div className="table-wrap preview-table"><table><thead><tr><th>البرنامج</th><th>المنشأة</th><th>الفترة</th><th>النسبة</th><th>المصدر</th></tr></thead><tbody>{preview.entries.slice(0, 12).map((entry) => <tr key={`${entry.program}-${entry.facilityId}-${entry.period}`}><td>{PROGRAMS.find((item) => item.id === entry.program)?.name}</td><td>{entry.facilityName}</td><td>{entry.period}</td><td>{formatPercent(entry.score, 2)}</td><td>{entry.sourceSheet} · صف {entry.sourceRow}</td></tr>)}</tbody></table></div>
      <div className="actions"><Button disabled={saving} onClick={() => void saveImport()}><Save size={16} />{saving ? "جارٍ الحفظ" : `حفظ ${preview.entries.length} نتيجة`}</Button><Button variant="outline" onClick={() => { setPreview(null); setFileName(""); }}>إلغاء</Button></div>
    </div>}
  </section>;
}

function PeriodComparison({ item }: { item: ProgramResult }) {
  return <><div className="subheading">تفاصيل مقارنة الفترات</div><div className="table-wrap"><table><thead><tr><th>الفترة</th><th>المتوسط</th><th>المنشآت المكتملة</th><th>الإدخال الجزئي</th><th>التغير</th></tr></thead><tbody>{item.trend.map((point) => <tr key={point.scope}><td>{point.label}</td><td>{formatPercent(point.mean)}</td><td>{point.complete} / {item.total}</td><td>{point.partial}</td><td>{formatChange(point.change)}</td></tr>)}</tbody></table></div></>;
}

export default function Workspace({ mode }: { mode: Mode }) {
  const [year, setYear] = useState(2026);
  const [entries, setEntries] = useState<Entry[]>([]);
  const [loading, setLoading] = useState(true);
  const [notice, setNotice] = useState("");
  const [error, setError] = useState("");
  const [program, setProgram] = useState<ProgramId>("ipccc");
  const [period, setPeriod] = useState("Q1");
  const [group, setGroup] = useState("all");
  const [search, setSearch] = useState("");
  const [draftScores, setDraftScores] = useState<Record<string, string>>({});
  const [detailsId, setDetailsId] = useState("");
  const [detailForm, setDetailForm] = useState<Form>(emptyForm);
  const [saving, setSaving] = useState("");
  const [reportMode, setReportMode] = useState<ReportMode>("year");
  const [selectedQuarters, setSelectedQuarters] = useState<Scope[]>(["Q1"]);
  const [selectedMonth, setSelectedMonth] = useState<Scope>("M01");
  const [reportProgram, setReportProgram] = useState("all");
  const [exporting, setExporting] = useState("");
  const [reportSettings, setReportSettings] = useState<ReportSettings | null>(null);
  const [editorOpen, setEditorOpen] = useState(true);
  const [savingSettings, setSavingSettings] = useState(false);

  const scope = useMemo<ReportScope>(() => {
    if (reportMode === "quarters") return createComparisonScope(selectedQuarters);
    if (reportMode === "month") return selectedMonth;
    return reportMode;
  }, [reportMode, selectedQuarters, selectedMonth]);

  const load = useCallback(async (currentYear: number) => {
    setLoading(true);
    setError("");
    try {
      const response = await fetch(`/api/entries?year=${currentYear}`, { cache: "no-store" });
      const body = await response.json() as { error?: string; entries?: Entry[] };
      if (!response.ok) throw new Error(body.error || "تعذر تحميل البيانات.");
      setEntries(body.entries ?? []);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "تعذر تحميل البيانات.");
    } finally {
      setLoading(false);
    }
  }, []);
  useEffect(() => { void load(year); }, [year, load]);

  const definition = PROGRAMS.find((item) => item.id === program)!;
  const periodOptions = PERIODS[definition.cadence].map((item) => ({ value: item.key, label: item.label }));
  const ownFacilities = FACILITIES.filter((item) => item.program === program);
  const groupOptions = [{ value: "all", label: "كل التصنيفات" }, ...[...new Set(ownFacilities.map((item) => item.group))].map((value) => ({ value, label: GROUP_LABELS[value] ?? value }))];
  const visibleFacilities = ownFacilities.filter((item) => (group === "all" || item.group === group) && (!search || `${item.name} ${item.arabicName ?? ""}`.toLowerCase().includes(search.toLowerCase())));
  const existingByFacility = useMemo(() => new Map(entries.filter((item) => item.program === program && item.period === period).map((item) => [item.facilityId, item])), [entries, program, period]);
  useEffect(() => { const scores: Record<string, string> = {}; for (const [id, item] of existingByFacility) scores[id] = String(item.score); setDraftScores(scores); }, [existingByFacility]);
  useEffect(() => { const item = existingByFacility.get(detailsId); setDetailForm(item ? { score: String(item.score), numerator: item.numerator?.toString() ?? "", denominator: item.denominator?.toString() ?? "", observation: item.observation, action: item.action, owner: item.owner, dueDate: item.dueDate } : emptyForm); }, [detailsId, existingByFacility]);

  function changeProgram(value: string) {
    const next = value as ProgramId;
    setProgram(next);
    setPeriod(PERIODS[PROGRAMS.find((item) => item.id === next)!.cadence][0].key);
    setGroup("all");
    setSearch("");
    setDetailsId("");
  }
  async function save(facilityId: string, form: Form) {
    setSaving(facilityId);
    setError("");
    setNotice("");
    try {
      const response = await fetch("/api/entries", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ year, program, facilityId, period, ...form }) });
      const body = await response.json() as { error?: string };
      if (!response.ok) throw new Error(body.error || "تعذر الحفظ.");
      await load(year);
      setNotice("تم حفظ النتيجة. ستظهر في لوحة المتابعة عند اكتمال بيانات الفترة.");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "تعذر الحفظ.");
    } finally {
      setSaving("");
    }
  }
  function quickSave(facilityId: string) {
    const score = draftScores[facilityId];
    if (score === undefined || score === "") { setError("أدخل النسبة قبل الحفظ."); return; }
    const previous = existingByFacility.get(facilityId);
    void save(facilityId, { score, numerator: previous?.numerator?.toString() ?? "", denominator: previous?.denominator?.toString() ?? "", observation: previous?.observation ?? "", action: previous?.action ?? "", owner: previous?.owner ?? "", dueDate: previous?.dueDate ?? "" });
  }
  async function deleteEntry(facilityId: string) {
    setError("");
    setNotice("");
    try {
      const query = new URLSearchParams({ year: String(year), program, facilityId, period });
      const response = await fetch(`/api/entries?${query}`, { method: "DELETE" });
      const body = await response.json() as { error?: string };
      if (!response.ok) throw new Error(body.error || "تعذر الحذف.");
      await load(year);
      setDetailsId("");
      setNotice("حُذفت النتيجة من الفترة المحددة.");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "تعذر الحذف.");
    }
  }

  const allReports = useMemo(() => buildReport(scope, entries), [scope, entries]);
  const reports = useMemo(() => allReports.filter((item) => reportProgram === "all" || item.program === reportProgram), [allReports, reportProgram]);
  const complete = reports.reduce((sum, item) => sum + item.complete, 0);
  const total = reports.reduce((sum, item) => sum + item.total, 0);
  const partial = reports.reduce((sum, item) => sum + item.partial, 0);
  const automaticSettings = useMemo(() => defaultReportSettings(year, scope, reports), [year, scope, reports]);
  const effectiveSettings = reportSettings ?? automaticSettings;

  useEffect(() => {
    if (mode !== "dashboard" || loading) return;
    const controller = new AbortController();
    setReportSettings(null);
    async function loadSettings() {
      try {
        const query = new URLSearchParams({ year: String(year), scope, program: reportProgram });
        const response = await fetch(`/api/report-settings?${query}`, { cache: "no-store", signal: controller.signal });
        const body = await response.json() as { settings?: ReportSettings | null; error?: string };
        if (!response.ok) throw new Error(body.error || "تعذر تحميل إعدادات التقرير.");
        setReportSettings(body.settings ?? defaultReportSettings(year, scope, reports));
      } catch (caught) {
        if ((caught as Error).name !== "AbortError") setError(caught instanceof Error ? caught.message : "تعذر تحميل إعدادات التقرير.");
      }
    }
    void loadSettings();
    return () => controller.abort();
  }, [mode, loading, year, scope, reportProgram, reports]);

  async function saveSettings() {
    setSavingSettings(true);
    setError("");
    setNotice("");
    try {
      const response = await fetch("/api/report-settings", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ year, scope, programFilter: reportProgram, settings: effectiveSettings }) });
      const body = await response.json() as { error?: string };
      if (!response.ok) throw new Error(body.error || "تعذر حفظ إعدادات التقرير.");
      setReportSettings(effectiveSettings);
      setNotice("حُفظت صياغة التقرير وستُستخدم في الطباعة وملفات التصدير.");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "تعذر حفظ إعدادات التقرير.");
    } finally {
      setSavingSettings(false);
    }
  }

  async function exportFile(kind: "excel" | "pptx") {
    setExporting(kind);
    setError("");
    try {
      const module = await import("@/lib/export-client");
      if (kind === "excel") await module.exportExcel(year, scope, reports, entries, effectiveSettings);
      else await module.exportPowerPoint(year, scope, reports, effectiveSettings);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "تعذر تجهيز الملف.");
    } finally {
      setExporting("");
    }
  }

  function toggleQuarter(value: Scope) {
    setSelectedQuarters((current) => current.includes(value) ? (current.length === 1 ? current : current.filter((item) => item !== value)) : [...current, value].sort());
  }

  useEffect(() => {
    type BrowserTool = { name: string; title: string; description: string; inputSchema: object; annotations: { readOnlyHint: boolean; untrustedContentHint: boolean }; execute: (input: unknown) => Promise<unknown> };
    const context = (document as Document & { modelContext?: { registerTool: (tool: BrowserTool, options: { signal: AbortSignal }) => void | Promise<void> } }).modelContext;
    if (!context?.registerTool) return;
    const lifecycle = new AbortController();
    const tool: BrowserTool = mode === "entry" ? {
      name: "save_selected_program_score", title: "حفظ نتيجة البرنامج", description: "احفظ نسبة منشأة للبرنامج والسنة والفترة المحددة حاليًا في صفحة الإدخال.",
      inputSchema: { type: "object", properties: { facilityId: { type: "string" }, score: { type: "number", minimum: 0, maximum: 100 } }, required: ["facilityId", "score"], additionalProperties: false },
      annotations: { readOnlyHint: false, untrustedContentHint: false },
      async execute(input) {
        const value = input as { facilityId?: string; score?: number };
        if (typeof value?.facilityId !== "string" || !FACILITIES.some((item) => item.id === value.facilityId && item.program === program) || typeof value.score !== "number" || !Number.isFinite(value.score) || value.score < 0 || value.score > 100) throw new Error("المنشأة أو النسبة غير صحيحة.");
        const previous = entries.find((item) => item.facilityId === value.facilityId && item.period === period && item.program === program);
        const response = await fetch("/api/entries", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ year, program, facilityId: value.facilityId, period, score: value.score, numerator: previous?.numerator ?? null, denominator: previous?.denominator ?? null, observation: previous?.observation ?? "", action: previous?.action ?? "", owner: previous?.owner ?? "", dueDate: previous?.dueDate ?? "" }) });
        if (!response.ok) throw new Error("تعذر حفظ النتيجة.");
        await load(year);
        setNotice("تم حفظ النتيجة.");
        return { saved: true, facilityId: value.facilityId, score: value.score, year, period };
      },
    } : {
      name: "read_selected_report_summary", title: "قراءة ملخص التقرير", description: "اقرأ ملخص البرامج للفترة والسنة المحددتين في لوحة المتابعة.",
      inputSchema: { type: "object", properties: {}, additionalProperties: false }, annotations: { readOnlyHint: true, untrustedContentHint: false },
      async execute() { return { year, period: scope, programs: reports.map((item) => ({ program: item.program, mean: item.mean, complete: item.complete, total: item.total, partial: item.partial })) }; },
    };
    try { void Promise.resolve(context.registerTool(tool, { signal: lifecycle.signal })).catch(() => {}); } catch {}
    return () => lifecycle.abort();
  }, [mode, year, program, period, scope, entries, reports, load]);

  return <><Header mode={mode} /><main className="page">
    {mode === "entry" ? <>
      <div className="page-heading"><div><h1>إدخال نتائج البرامج</h1><p>يمكنك استيراد ملف Excel كامل أو إدخال النتائج يدويًا حسب المنشأة والفترة.</p></div><span className="pill pending"><Activity size={16} />البيانات المدخلة هنا فقط</span></div>
      <ImportPanel year={year} onImported={() => load(year)} />
      <section className="panel panel-padding controls-panel"><div className="section-title compact-title"><div><h2>الإدخال اليدوي</h2><p>استخدم هذا القسم لتعديل نتيجة واحدة أو إضافة الملاحظات والإجراءات.</p></div></div><div className="control-grid"><div className="field"><label htmlFor="entry-year">سنة التقرير</label><input id="entry-year" type="number" min="2020" max="2100" value={year} onChange={(event) => setYear(Number(event.target.value))} /></div><SelectField label="البرنامج" value={program} onChange={changeProgram} options={PROGRAMS.map((item) => ({ value: item.id, label: item.name }))} /><SelectField label="فترة الإدخال" value={period} onChange={(value) => { setPeriod(value); setDetailsId(""); }} options={periodOptions} /><SelectField label="تصنيف المنشأة" value={group} onChange={setGroup} options={groupOptions} /><div className="field"><label htmlFor="facility-search">بحث عن منشأة</label><input id="facility-search" value={search} onChange={(event) => setSearch(event.target.value)} placeholder="اكتب الاسم" /></div></div><p className="muted" style={{ margin: "14px 0 0", fontSize: 14 }}>نظافة اليدين تُدخل شهريًا، وبرنامجا IPCCC وRPP ربعيًا، وتقييم PHC ICA نصف سنوي. تعني 0% نتيجة فعلية إذا أدخلتها أنت.</p></section>
      {error && <ErrorNotice message={error} />}{notice && <div className="notice" role="status">{notice}</div>}
      <section className="panel data-panel"><div className="data-toolbar"><div className="section-title" style={{ margin: 0 }}><h2>قائمة المنشآت</h2><p>{visibleFacilities.length} منشأة في العرض</p></div><Search size={19} color="#5e8584" /></div><div className="table-wrap"><table><thead><tr><th>المنشأة</th><th>التصنيف</th><th>النسبة %</th><th>الحالة</th><th>الإجراء</th></tr></thead><tbody>{visibleFacilities.map((facility) => { const current = existingByFacility.get(facility.id); return <tr key={facility.id}><td><strong>{facility.name}</strong>{facility.arabicName && <div className="muted">{facility.arabicName}</div>}</td><td>{GROUP_LABELS[facility.group] ?? facility.group}</td><td><input className="score-input" aria-label={`نسبة ${facility.name}`} type="number" min="0" max="100" step="0.01" value={draftScores[facility.id] ?? ""} onChange={(event) => setDraftScores((previous) => ({ ...previous, [facility.id]: event.target.value }))} /></td><td>{current ? <span className="pill success">محفوظة</span> : <span className="pill pending">لم تُدخل</span>}</td><td><div className="actions"><Button size="sm" disabled={saving === facility.id || loading} onClick={() => quickSave(facility.id)}><Save size={15} />{saving === facility.id ? "جارٍ الحفظ" : "حفظ"}</Button><Button size="sm" variant="outline" onClick={() => setDetailsId(facility.id)}>تفاصيل</Button></div></td></tr>; })}</tbody></table>{loading && <div className="empty">جارٍ تحميل النتائج...</div>}{!loading && visibleFacilities.length === 0 && <div className="empty">لا توجد منشأة تطابق البحث.</div>}</div></section>
      {detailsId && <section className="panel panel-padding detail-panel"><div className="section-title"><div><h2>تفاصيل النتيجة والمتابعة</h2><p>{FACILITIES.find((item) => item.id === detailsId)?.name} · {periodOptions.find((item) => item.value === period)?.label}</p></div><Button variant="ghost" onClick={() => setDetailsId("")}>إغلاق</Button></div><div className="detail-grid"><div className="field"><label htmlFor="detail-score">النسبة %</label><input id="detail-score" type="number" min="0" max="100" step="0.01" value={detailForm.score} onChange={(event) => setDetailForm({ ...detailForm, score: event.target.value })} /></div><div className="field"><label htmlFor="detail-num">البسط (اختياري)</label><input id="detail-num" type="number" min="0" value={detailForm.numerator} onChange={(event) => setDetailForm({ ...detailForm, numerator: event.target.value })} /></div><div className="field"><label htmlFor="detail-den">المقام (اختياري)</label><input id="detail-den" type="number" min="1" value={detailForm.denominator} onChange={(event) => setDetailForm({ ...detailForm, denominator: event.target.value })} /></div><div className="field wide"><label htmlFor="detail-observation">تفسير النتيجة أو الملاحظة</label><textarea id="detail-observation" value={detailForm.observation} onChange={(event) => setDetailForm({ ...detailForm, observation: event.target.value })} /></div><div className="field wide"><label htmlFor="detail-action">الإجراء التصحيحي أو الإنجاز</label><textarea id="detail-action" value={detailForm.action} onChange={(event) => setDetailForm({ ...detailForm, action: event.target.value })} /></div><div className="field"><label htmlFor="detail-owner">المسؤول</label><input id="detail-owner" value={detailForm.owner} onChange={(event) => setDetailForm({ ...detailForm, owner: event.target.value })} /></div><div className="field"><label htmlFor="detail-due">موعد المتابعة</label><input id="detail-due" type="date" value={detailForm.dueDate} onChange={(event) => setDetailForm({ ...detailForm, dueDate: event.target.value })} /></div></div><div className="actions" style={{ marginTop: 20 }}><Button disabled={!detailForm.score || saving === detailsId} onClick={() => void save(detailsId, detailForm)}><Save size={16} />حفظ التفاصيل</Button>{existingByFacility.has(detailsId) && <AlertDialog><AlertDialogTrigger asChild><Button variant="destructive"><Trash2 size={16} />حذف الإدخال</Button></AlertDialogTrigger><AlertDialogContent><AlertDialogHeader><AlertDialogTitle>حذف نتيجة هذه الفترة؟</AlertDialogTitle><AlertDialogDescription>سيُزال الإدخال لهذه المنشأة والفترة من لوحة المتابعة والتقارير.</AlertDialogDescription></AlertDialogHeader><AlertDialogFooter><AlertDialogCancel>إلغاء</AlertDialogCancel><AlertDialogAction onClick={() => void deleteEntry(detailsId)}>حذف النتيجة</AlertDialogAction></AlertDialogFooter></AlertDialogContent></AlertDialog>}</div></section>}
    </> : <>
      <div className="page-heading"><div><h1>لوحة المتابعة والتقارير</h1><p>أنشئ تقريرًا لفترة واحدة، أو ادمج ربعين أو ثلاثة أو أربعة أرباع في تقرير مقارنة واحد.</p></div></div>
      <section className="panel panel-padding controls-panel"><div className="control-grid dashboard-controls"><div className="field"><label htmlFor="dashboard-year">السنة</label><input id="dashboard-year" type="number" min="2020" max="2100" value={year} onChange={(event) => setYear(Number(event.target.value))} /></div><SelectField label="نوع التقرير" value={reportMode} onChange={(value) => setReportMode(value as ReportMode)} options={[{ value: "year", label: "التقرير السنوي" }, { value: "quarters", label: "اختيار ربع أو عدة أرباع" }, { value: "H1", label: "النصف الأول" }, { value: "H2", label: "النصف الثاني" }, { value: "month", label: "شهر محدد لنظافة اليدين" }]} />{reportMode === "month" && <SelectField label="الشهر" value={selectedMonth} onChange={(value) => setSelectedMonth(value as Scope)} options={MONTHS.map((label, index) => ({ value: `M${String(index + 1).padStart(2, "0")}`, label }))} />}<SelectField label="البرنامج" value={reportProgram} onChange={setReportProgram} options={[{ value: "all", label: "جميع البرامج المتاحة للفترة" }, ...PROGRAMS.map((item) => ({ value: item.id, label: item.name }))]} /><div className="export-bar"><Button variant="outline" onClick={() => window.print()}><FileText size={16} />PDF</Button><Button variant="outline" disabled={!!exporting} onClick={() => void exportFile("excel")}><FileDown size={16} />Excel</Button><Button disabled={!!exporting} onClick={() => void exportFile("pptx")}><Presentation size={16} />PowerPoint</Button></div></div>{reportMode === "quarters" && <div className="period-selector"><strong>الأرباع المطلوبة في التقرير</strong><div>{quarterOptions.map((option) => <label key={option.value}><Checkbox checked={selectedQuarters.includes(option.value)} onCheckedChange={() => toggleQuarter(option.value)} />{option.label}</label>)}</div><p>يمكن اختيار ربع واحد أو ربعين أو ثلاثة أو الأرباع الأربعة. يجمع التقرير النتائج ويعرض المقارنة بين الفترات المختارة.</p></div>}</section>
      {error && <ErrorNotice message={error} />}{exporting && <div className="notice" role="status">جارٍ تجهيز الملف...</div>}{notice && <div className="notice" role="status">{notice}</div>}
      <section className="panel panel-padding report-editor no-print"><div className="section-title"><div><h2><PencilLine size={20} /> تحرير محتوى التقرير</h2><p>عدّل الصياغة والعناصر التي ستظهر في PDF وExcel وPowerPoint قبل التصدير.</p></div><Button variant="outline" onClick={() => setEditorOpen((value) => !value)}>{editorOpen ? "إخفاء اللوحة" : "فتح اللوحة"}</Button></div>{editorOpen && <><div className="detail-grid"><div className="field wide"><label htmlFor="report-title">عنوان التقرير</label><input id="report-title" value={effectiveSettings.title} onChange={(event) => setReportSettings({ ...effectiveSettings, title: event.target.value })} /></div><div className="field wide"><label htmlFor="report-org">الجهة</label><input id="report-org" value={effectiveSettings.organization} onChange={(event) => setReportSettings({ ...effectiveSettings, organization: event.target.value })} /></div><div className="field wide"><label htmlFor="report-summary">الملخص التنفيذي</label><textarea id="report-summary" value={effectiveSettings.executiveSummary} onChange={(event) => setReportSettings({ ...effectiveSettings, executiveSummary: event.target.value })} /></div><div className="field wide"><label htmlFor="report-method">المنهج والتغطية</label><textarea id="report-method" value={effectiveSettings.methodology} onChange={(event) => setReportSettings({ ...effectiveSettings, methodology: event.target.value })} /></div><div className="field wide"><label htmlFor="report-conclusion">الخلاصة والتوصيات</label><textarea id="report-conclusion" value={effectiveSettings.conclusion} onChange={(event) => setReportSettings({ ...effectiveSettings, conclusion: event.target.value })} /></div>{reports.map((item) => <div className="field wide" key={item.program}><label htmlFor={`program-note-${item.program}`}>شرح نتائج {item.label}</label><textarea id={`program-note-${item.program}`} value={effectiveSettings.programNotes[item.program] ?? ""} onChange={(event) => setReportSettings({ ...effectiveSettings, programNotes: { ...effectiveSettings.programNotes, [item.program]: event.target.value } })} /></div>)}<div className="field"><label htmlFor="reviewed-by">مراجعة</label><input id="reviewed-by" value={effectiveSettings.reviewedBy} onChange={(event) => setReportSettings({ ...effectiveSettings, reviewedBy: event.target.value })} placeholder="الاسم أو المسمى" /></div><div className="field"><label htmlFor="approved-by">اعتماد</label><input id="approved-by" value={effectiveSettings.approvedBy} onChange={(event) => setReportSettings({ ...effectiveSettings, approvedBy: event.target.value })} placeholder="الاسم أو المسمى" /></div></div><div className="report-options"><label><Checkbox checked={effectiveSettings.includeMethodology} onCheckedChange={(checked) => setReportSettings({ ...effectiveSettings, includeMethodology: checked === true })} />إظهار المنهج والتغطية</label><label><Checkbox checked={effectiveSettings.includeFacilityDetails} onCheckedChange={(checked) => setReportSettings({ ...effectiveSettings, includeFacilityDetails: checked === true })} />إظهار تفاصيل المنشآت</label><label><Checkbox checked={effectiveSettings.includeNotes} onCheckedChange={(checked) => setReportSettings({ ...effectiveSettings, includeNotes: checked === true })} />إظهار الملاحظات والإجراءات</label></div><div className="actions" style={{ marginTop: 18 }}><Button disabled={savingSettings} onClick={() => void saveSettings()}><Save size={16} />{savingSettings ? "جارٍ الحفظ" : "حفظ صياغة التقرير"}</Button><Button variant="outline" onClick={() => setReportSettings(automaticSettings)}><RotateCcw size={16} />استعادة النص التلقائي</Button></div></>}</section>
      <div className="metric-grid no-print"><div className="panel metric"><small>البرامج في التقرير</small><strong>{reports.length}</strong></div><div className="panel metric"><small>المنشآت المكتملة</small><strong dir="ltr" style={{ textAlign: "right" }}>{complete} / {total}</strong></div><div className="panel metric"><small>إدخالات جزئية</small><strong>{partial}</strong></div><div className="panel metric"><small>الفترة المختارة</small><strong style={{ fontSize: 19 }}>{periodLabel(scope)}</strong></div></div>
      <article className="report-sheet" id="report-content"><header className="report-header"><div><h2>{effectiveSettings.title}</h2><p>{periodLabel(scope)} · {year} · {effectiveSettings.organization}</p></div><img src="/wiqaya-logo.png" alt="شعار وقاية" /></header><div className="explanation"><strong>الملخص التنفيذي:</strong> {effectiveSettings.executiveSummary}</div>{effectiveSettings.includeMethodology && <><div className="subheading">المنهج والتغطية</div><p className="muted">{effectiveSettings.methodology}</p></>}
        {loading ? <div className="empty">جارٍ تحميل التقرير...</div> : reports.length === 0 ? <div className="empty"><strong>لا يوجد برنامج لهذه الفترة</strong>{reportProgram === "phc_ica" && reportMode === "quarters" ? "اختر النصف الأول أو النصف الثاني لبرنامج PHC ICA." : "اختر فترة تتوافق مع البرنامج."}</div> : reports.map((item) => <section className="report-program" key={item.program}><h3>{item.label}</h3><div className="program-lead"><span className={item.complete ? "pill success" : "pill pending"}>المتوسط: {formatPercent(item.mean)}</span><span className="pill partial">مكتملة: {item.complete} / {item.total}</span>{item.partial > 0 && <span className="pill pending">جزئية: {item.partial}</span>}</div><div className="chart-layout"><div className="chart-box"><h4>مقارنة الفترات</h4><ReportChart item={item} type="trend" /></div><div className="chart-box"><h4>متوسطات تصنيفات المنشآت</h4><ReportChart item={item} type="groups" /></div></div><PeriodComparison item={item} /><p className="explanation">{effectiveSettings.programNotes[item.program] ?? reportNarrative(item, scope)}</p><div className="subheading">جدول التصنيفات</div><div className="table-wrap"><table><thead><tr><th>التصنيف</th><th>المتوسط</th><th>المكتملة</th><th>إدخال جزئي</th></tr></thead><tbody>{item.groups.map((groupItem) => <tr key={groupItem.group}><td>{groupItem.label}</td><td>{formatPercent(groupItem.mean)}</td><td>{groupItem.complete} / {groupItem.total}</td><td>{groupItem.partial}</td></tr>)}</tbody></table></div><p className="muted">شرح الجدول: استُبعدت المنشآت التي لم تكتمل فتراتها المطلوبة من المتوسط، وظلت محسوبة ضمن التغطية.</p>{effectiveSettings.includeFacilityDetails && <><div className="subheading">تفاصيل المنشآت</div>{item.groups.map((groupItem) => <details key={groupItem.group} open><summary>{groupItem.label} · {groupItem.complete} مكتملة من {groupItem.total}</summary><div className="table-wrap"><table><thead><tr><th>المنشأة</th><th>النتيجة</th><th>الفترات المدخلة</th><th>الحالة</th></tr></thead><tbody>{groupItem.facilities.map((row) => <tr key={row.facility.id}><td>{row.facility.name}</td><td>{formatPercent(row.score)}</td><td>{row.entered} / {row.required}</td><td>{row.score !== null ? "مكتملة" : row.entered ? "جزئية" : "لم تُدخل"}</td></tr>)}</tbody></table></div></details>)}<p className="muted">شرح التفاصيل: تظهر النتيجة فقط إذا كانت جميع الفترات المطلوبة للمنشأة مدخلة.</p></>}{effectiveSettings.includeNotes && <><div className="subheading">الملاحظات والإجراءات</div>{item.observations.length || item.actions.length ? <div className="chart-layout"><div><strong>ملاحظات موثقة</strong><ul>{item.observations.map((text, index) => <li key={index}>{text}</li>)}</ul></div><div><strong>إجراءات مسجلة</strong><ul>{item.actions.map((text, index) => <li key={index}>{text}</li>)}</ul></div></div> : <p className="muted">لم تُسجل ملاحظات أو إجراءات لهذه الفترة بعد.</p>}</>}</section>)}<div className="subheading">الخلاصة والمتابعة</div><p className="muted">{effectiveSettings.conclusion}</p><div className="report-footer"><div className="signature">مراجعة{effectiveSettings.reviewedBy && <strong>{effectiveSettings.reviewedBy}</strong>}<span>الاسم والتوقيع</span></div><div className="signature">اعتماد{effectiveSettings.approvedBy && <strong>{effectiveSettings.approvedBy}</strong>}<span>الاسم والتوقيع</span></div></div></article>
    </>}
  </main></>;
}
