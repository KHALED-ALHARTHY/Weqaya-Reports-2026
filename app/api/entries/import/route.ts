import { FACILITY_BY_ID, PERIODS, PROGRAMS, type ProgramId } from "@/lib/catalog";
import { upsertEntries } from "@/db/records";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    const body = await request.json() as { year?: unknown; entries?: unknown };
    const year = Number(body.year);
    if (!Number.isInteger(year) || year < 2020 || year > 2100) throw new Error("سنة التقرير غير صحيحة.");
    if (!Array.isArray(body.entries) || !body.entries.length || body.entries.length > 2500) throw new Error("ملف الاستيراد لا يحتوي على نتائج صالحة.");
    const prepared = body.entries.map((raw) => {
      const value = raw as Record<string, unknown>;
      const program = String(value.program ?? "") as ProgramId;
      const facilityId = String(value.facilityId ?? "");
      const period = String(value.period ?? "");
      const score = Number(value.score);
      const definition = PROGRAMS.find((item) => item.id === program);
      if (!definition || FACILITY_BY_ID.get(facilityId)?.program !== program) throw new Error("يوجد مركز غير مطابق في ملف الاستيراد.");
      if (!PERIODS[definition.cadence].some((item) => item.key === period)) throw new Error("يوجد عمود فترة غير مطابق في ملف الاستيراد.");
      if (!Number.isFinite(score) || score < 0 || score > 100) throw new Error("يوجد رقم خارج النطاق من 0 إلى 100.");
      return { year, program, facilityId, period, score, numerator: null, denominator: null, observation: "", action: "", owner: "", dueDate: "" };
    });
    const unique = new Map(prepared.map((entry) => [`${entry.program}|${entry.facilityId}|${entry.period}`, entry]));
    await upsertEntries([...unique.values()]);
    return Response.json({ ok: true, saved: unique.size });
  } catch (error) {
    const message = error instanceof Error ? error.message : "تعذر استيراد الملف.";
    return Response.json({ error: message }, { status: message.startsWith("تعذر") ? 503 : 400 });
  }
}
