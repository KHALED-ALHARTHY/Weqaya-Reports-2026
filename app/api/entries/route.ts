import { FACILITY_BY_ID, PERIODS, PROGRAMS, type ProgramId } from "@/lib/catalog";
import { readEntries, removeEntry, upsertEntry } from "@/db/records";

export const dynamic = "force-dynamic";

function validYear(value: unknown): value is number {
  return Number.isInteger(value) && Number(value) >= 2020 && Number(value) <= 2100;
}

function validate(payload: Record<string, unknown>) {
  const year = Number(payload.year);
  const program = String(payload.program ?? "") as ProgramId;
  const facilityId = String(payload.facilityId ?? "");
  const period = String(payload.period ?? "");
  if (!validYear(year)) throw new Error("سنة التقرير غير صحيحة.");
  const definition = PROGRAMS.find((item) => item.id === program);
  if (!definition) throw new Error("البرنامج غير معروف.");
  if (FACILITY_BY_ID.get(facilityId)?.program !== program) throw new Error("المنشأة لا تتبع البرنامج المحدد.");
  if (!PERIODS[definition.cadence].some((item) => item.key === period)) throw new Error("الفترة لا تتبع البرنامج المحدد.");
  return { year, program, facilityId, period };
}

function safeText(value: unknown, max = 1200) {
  return String(value ?? "").trim().slice(0, max);
}

export async function GET(request: Request) {
  const year = Number(new URL(request.url).searchParams.get("year") ?? 2026);
  if (!validYear(year)) return Response.json({ error: "سنة التقرير غير صحيحة." }, { status: 400 });
  try { return Response.json({ entries: await readEntries(year) }, { headers: { "Cache-Control": "no-store" } }); }
  catch { return Response.json({ error: "تعذر تحميل النتائج. حاول مرة أخرى." }, { status: 503 }); }
}

export async function POST(request: Request) {
  try {
    const body = await request.json() as Record<string, unknown>;
    const key = validate(body);
    const score = Number(body.score);
    if (!Number.isFinite(score) || score < 0 || score > 100) throw new Error("أدخل نسبة من 0 إلى 100.");
    const numerator = body.numerator === "" || body.numerator == null ? null : Number(body.numerator);
    const denominator = body.denominator === "" || body.denominator == null ? null : Number(body.denominator);
    if ((numerator !== null && (!Number.isInteger(numerator) || numerator < 0)) ||
        (denominator !== null && (!Number.isInteger(denominator) || denominator <= 0)) ||
        (numerator !== null && denominator !== null && numerator > denominator)) throw new Error("راجع البسط والمقام.");
    const dueDate = safeText(body.dueDate, 10);
    if (dueDate && !/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) throw new Error("تاريخ المتابعة غير صحيح.");
    await upsertEntry({ ...key, score, numerator, denominator,
      observation: safeText(body.observation), action: safeText(body.action),
      owner: safeText(body.owner, 120), dueDate });
    return Response.json({ ok: true });
  } catch (error) {
    const message = error instanceof Error ? error.message : "تعذر الحفظ.";
    const status = message.startsWith("تعذر") ? 503 : 400;
    return Response.json({ error: message }, { status });
  }
}

export async function DELETE(request: Request) {
  try {
    const query = Object.fromEntries(new URL(request.url).searchParams.entries());
    const key = validate(query);
    await removeEntry(key.year, key.program, key.facilityId, key.period);
    return Response.json({ ok: true });
  } catch (error) {
    return Response.json({ error: error instanceof Error ? error.message : "تعذر الحذف." }, { status: 400 });
  }
}
