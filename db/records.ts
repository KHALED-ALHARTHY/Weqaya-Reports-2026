import { env } from "cloudflare:workers";
import { FACILITY_ID_ALIASES } from "@/lib/catalog";

export type Entry = {
  id: number;
  year: number;
  program: string;
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

function database() {
  if (!env.DB) throw new Error("تعذر الاتصال بقاعدة البيانات. حاول مرة أخرى لاحقًا.");
  return env.DB;
}

export async function readEntries(year: number): Promise<Entry[]> {
  const result = await database().prepare(
    "SELECT id, year, program, facility_id AS facilityId, period, score, numerator, denominator, observation, action, owner, due_date AS dueDate, updated_at AS updatedAt FROM entries WHERE year = ? ORDER BY program, facility_id, period"
  ).bind(year).all<Entry>();
  const normalized = new Map<string, Entry>();
  for (const row of result.results ?? []) {
    const facilityId = FACILITY_ID_ALIASES[row.facilityId] ?? row.facilityId;
    const entry = { ...row, facilityId };
    const key = `${entry.program}|${facilityId}|${entry.period}`;
    const previous = normalized.get(key);
    if (!previous || previous.updatedAt <= entry.updatedAt) normalized.set(key, entry);
  }
  return [...normalized.values()];
}

export async function upsertEntry(entry: Omit<Entry, "id" | "updatedAt">): Promise<void> {
  await database().prepare(
    `INSERT INTO entries (year, program, facility_id, period, score, numerator, denominator, observation, action, owner, due_date, updated_at)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
     ON CONFLICT(year, program, facility_id, period) DO UPDATE SET
       score=excluded.score, numerator=excluded.numerator, denominator=excluded.denominator,
       observation=excluded.observation, action=excluded.action, owner=excluded.owner,
       due_date=excluded.due_date, updated_at=excluded.updated_at`
  ).bind(entry.year, entry.program, entry.facilityId, entry.period, entry.score,
    entry.numerator, entry.denominator, entry.observation, entry.action, entry.owner,
    entry.dueDate, new Date().toISOString()).run();
}

export async function upsertEntries(entries: Omit<Entry, "id" | "updatedAt">[]): Promise<void> {
  const db = database();
  const now = new Date().toISOString();
  for (let start = 0; start < entries.length; start += 50) {
    const statements = entries.slice(start, start + 50).map((entry) => db.prepare(
      `INSERT INTO entries (year, program, facility_id, period, score, numerator, denominator, observation, action, owner, due_date, updated_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
       ON CONFLICT(year, program, facility_id, period) DO UPDATE SET
         score=excluded.score, numerator=excluded.numerator, denominator=excluded.denominator,
         observation=excluded.observation, action=excluded.action, owner=excluded.owner,
         due_date=excluded.due_date, updated_at=excluded.updated_at`
    ).bind(entry.year, entry.program, entry.facilityId, entry.period, entry.score,
      entry.numerator, entry.denominator, entry.observation, entry.action, entry.owner,
      entry.dueDate, now));
    await db.batch(statements);
  }
}

export async function removeEntry(year: number, program: string, facilityId: string, period: string): Promise<void> {
  await database().prepare(
    "DELETE FROM entries WHERE year = ? AND program = ? AND facility_id = ? AND period = ?"
  ).bind(year, program, facilityId, period).run();
}
