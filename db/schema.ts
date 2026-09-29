import { integer, real, sqliteTable, text, uniqueIndex, index } from "drizzle-orm/sqlite-core";

export const entries = sqliteTable("entries", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  year: integer("year").notNull(),
  program: text("program").notNull(),
  facilityId: text("facility_id").notNull(),
  period: text("period").notNull(),
  score: real("score").notNull(),
  numerator: integer("numerator"),
  denominator: integer("denominator"),
  observation: text("observation").notNull().default(""),
  action: text("action").notNull().default(""),
  owner: text("owner").notNull().default(""),
  dueDate: text("due_date").notNull().default(""),
  updatedAt: text("updated_at").notNull(),
}, (table) => [
  uniqueIndex("entries_year_program_facility_period_unique").on(table.year, table.program, table.facilityId, table.period),
  index("entries_year_program_idx").on(table.year, table.program),
]);

export const reportSettings = sqliteTable("report_settings", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  year: integer("year").notNull(),
  scope: text("scope").notNull(),
  programFilter: text("program_filter").notNull(),
  title: text("title").notNull(),
  organization: text("organization").notNull(),
  executiveSummary: text("executive_summary").notNull(),
  methodology: text("methodology").notNull(),
  conclusion: text("conclusion").notNull(),
  preparedBy: text("prepared_by").notNull().default(""),
  reviewedBy: text("reviewed_by").notNull().default(""),
  approvedBy: text("approved_by").notNull().default(""),
  programNotesJson: text("program_notes_json").notNull().default("{}"),
  includeMethodology: integer("include_methodology", { mode: "boolean" }).notNull().default(true),
  includeFacilityDetails: integer("include_facility_details", { mode: "boolean" }).notNull().default(true),
  includeNotes: integer("include_notes", { mode: "boolean" }).notNull().default(true),
  updatedAt: text("updated_at").notNull(),
}, (table) => [
  uniqueIndex("report_settings_year_scope_program_unique").on(table.year, table.scope, table.programFilter),
]);
