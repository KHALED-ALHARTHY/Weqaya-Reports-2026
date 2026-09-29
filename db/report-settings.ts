import { env } from "cloudflare:workers";
import type { ReportSettings } from "@/lib/report-settings";

function database(){if(!env.DB)throw new Error("تعذر الاتصال بقاعدة البيانات.");return env.DB;}

export async function readReportSettings(year:number,scope:string,programFilter:string):Promise<ReportSettings|null>{
  const row=await database().prepare(`SELECT title, organization, executive_summary AS executiveSummary,
    methodology, conclusion, reviewed_by AS reviewedBy, approved_by AS approvedBy,
    program_notes_json AS programNotesJson, include_methodology AS includeMethodology,
    include_facility_details AS includeFacilityDetails, include_notes AS includeNotes
    FROM report_settings WHERE year=? AND scope=? AND program_filter=?`).bind(year,scope,programFilter).first<Record<string,unknown>>();
  if(!row)return null;
  let programNotes:Record<string,string>={};
  try{programNotes=JSON.parse(String(row.programNotesJson??"{}"));}catch{}
  return {title:String(row.title??""),organization:String(row.organization??""),executiveSummary:String(row.executiveSummary??""),
    methodology:String(row.methodology??""),conclusion:String(row.conclusion??""),
    reviewedBy:String(row.reviewedBy??""),approvedBy:String(row.approvedBy??""),programNotes,
    includeMethodology:Boolean(row.includeMethodology),includeFacilityDetails:Boolean(row.includeFacilityDetails),includeNotes:Boolean(row.includeNotes)};
}

export async function saveReportSettings(year:number,scope:string,programFilter:string,settings:ReportSettings){
  await database().prepare(`INSERT INTO report_settings
    (year,scope,program_filter,title,organization,executive_summary,methodology,conclusion,prepared_by,reviewed_by,approved_by,program_notes_json,include_methodology,include_facility_details,include_notes,updated_at)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ON CONFLICT(year,scope,program_filter) DO UPDATE SET title=excluded.title,organization=excluded.organization,
    executive_summary=excluded.executive_summary,methodology=excluded.methodology,conclusion=excluded.conclusion,
    prepared_by='',reviewed_by=excluded.reviewed_by,approved_by=excluded.approved_by,
    program_notes_json=excluded.program_notes_json,include_methodology=excluded.include_methodology,
    include_facility_details=excluded.include_facility_details,include_notes=excluded.include_notes,updated_at=excluded.updated_at`)
    .bind(year,scope,programFilter,settings.title,settings.organization,settings.executiveSummary,settings.methodology,
      settings.conclusion,"",settings.reviewedBy,settings.approvedBy,JSON.stringify(settings.programNotes),
      settings.includeMethodology?1:0,settings.includeFacilityDetails?1:0,settings.includeNotes?1:0,new Date().toISOString()).run();
}
