/* Isolated Python execution. Workbook bytes never leave this worker. */
const PYODIDE_URL = "https://cdn.jsdelivr.net/pyodide/v0.29.3/full/";
let runtime;
let busy = false;
const progress = (message) => self.postMessage({ type: "progress", message });

async function initialize() {
  progress("تهيئة بيئة التحليل · قد يستغرق التشغيل الأول عدة دقائق");
  importScripts(PYODIDE_URL + "pyodide.js");
  const py = await loadPyodide({ indexURL: PYODIDE_URL });
  progress("تحميل مكتبات التحليل والعروض");
  await py.loadPackage(["micropip", "pandas", "lxml", "pillow"]);
  await py.runPythonAsync('import micropip\nawait micropip.install(["openpyxl==3.1.5", "python-pptx==1.0.2", "XlsxWriter==3.2.5"])');
  const response = await fetch("/api/report-engine", { cache: "no-store" });
  if (!response.ok) throw new Error("تعذر تحميل محرك التقارير. حدّث الصفحة ثم أعد المحاولة.");
  const { source, base, catalog } = await response.json();
  const logoResponse = await fetch("/wiqaya-logo.png");
  if (!logoResponse.ok) throw new Error("تعذر تحميل شعار التقرير.");
  py.FS.mkdirTree("/home/pyodide/python_engine");
  py.FS.writeFile("/home/pyodide/python_engine/__init__.py", "");
  py.FS.writeFile("/home/pyodide/python_engine/base_pipeline.py", base);
  py.FS.writeFile("/home/pyodide/report_generator.py", source);
  py.FS.writeFile("/home/pyodide/catalog.json", JSON.stringify(catalog));
  py.FS.writeFile("/home/pyodide/logo.png", new Uint8Array(await logoResponse.arrayBuffer()));
  py.globals.set("report_progress", progress);
  return py;
}

self.onmessage = async ({ data }) => {
  if (busy || data.type !== "generate") return;
  busy = true;
  let py;
  try {
    runtime ??= initialize().catch((error) => { runtime = undefined; throw error; });
    py = await runtime;
    py.FS.writeFile("/home/pyodide/tracker.xlsx", new Uint8Array(data.buffer));
    py.globals.set("report_year", data.year);
    py.globals.set("skip_zeros", data.skipZeros);
    py.globals.set("report_scope", data.scope);
    py.globals.set("report_program", data.program);
    await py.runPythonAsync(`
import json
from pathlib import Path
from report_generator import InfectionControlSupervisor
report_result = InfectionControlSupervisor().run(
    Path('tracker.xlsx').read_bytes(), json.loads(Path('catalog.json').read_text()),
    year=report_year, logo=Path('logo.png').read_bytes(), skip_zeros=skip_zeros,
    progress=report_progress, scope=report_scope, program=report_program)
Path('report.pptx').write_bytes(report_result.pptx_bytes)
report_preview = json.dumps(report_result.preview, ensure_ascii=False)
del report_result
`);
    const preview = JSON.parse(py.globals.get("report_preview"));
    const output = py.FS.readFile("/home/pyodide/report.pptx");
    self.postMessage({ type: "done", buffer: output.buffer, preview }, [output.buffer]);
  } catch (error) {
    const raw = String(error?.message || error);
    const validation = raw.match(/ValueError: ([^\n]+)/);
    self.postMessage({ type: "error", message: validation?.[1] || "تعذر تشغيل محرك التقارير. تأكد من الاتصال لتنزيل المكتبات، ثم أعد المحاولة." });
  } finally {
    if (py) {
      for (const name of ["tracker.xlsx", "report.pptx"]) {
        try { py.FS.unlink("/home/pyodide/" + name); } catch {}
      }
      py.globals.delete("report_preview");
    }
    busy = false;
  }
};
