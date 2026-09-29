import { readReportSettings, saveReportSettings } from "@/db/report-settings";
import { PROGRAMS } from "@/lib/catalog";
import type { ReportSettings } from "@/lib/report-settings";
import { isValidReportScope } from "@/lib/report";

export const dynamic="force-dynamic";
const validScope=(value:string)=>isValidReportScope(value);
const validProgram=(value:string)=>value==="all"||PROGRAMS.some((item)=>item.id===value);
const text=(value:unknown,max:number)=>String(value??"").trim().slice(0,max);

export async function GET(request:Request){
  const query=new URL(request.url).searchParams;const year=Number(query.get("year"));const scope=query.get("scope")??"";const programFilter=query.get("program")??"all";
  if(!Number.isInteger(year)||year<2020||year>2100||!validScope(scope)||!validProgram(programFilter))return Response.json({error:"معايير التقرير غير صحيحة."},{status:400});
  try{return Response.json({settings:await readReportSettings(year,scope,programFilter)},{headers:{"Cache-Control":"no-store"}});}
  catch{return Response.json({error:"تعذر تحميل إعدادات التقرير."},{status:503});}
}

export async function POST(request:Request){
  try{
    const body=await request.json() as Record<string,unknown>;const year=Number(body.year);const scope=String(body.scope??"");const programFilter=String(body.programFilter??"all");
    if(!Number.isInteger(year)||year<2020||year>2100||!validScope(scope)||!validProgram(programFilter))throw new Error("معايير التقرير غير صحيحة.");
    const source=body.settings as Record<string,unknown>|undefined;if(!source)throw new Error("إعدادات التقرير مفقودة.");
    const rawNotes=source.programNotes&&typeof source.programNotes==="object"?source.programNotes as Record<string,unknown>:{};
    const programNotes=Object.fromEntries(Object.entries(rawNotes).filter(([key])=>PROGRAMS.some((item)=>item.id===key)).map(([key,value])=>[key,text(value,1800)]));
    const settings:ReportSettings={title:text(source.title,180),organization:text(source.organization,220),executiveSummary:text(source.executiveSummary,2600),methodology:text(source.methodology,2600),conclusion:text(source.conclusion,2600),reviewedBy:text(source.reviewedBy,150),approvedBy:text(source.approvedBy,150),programNotes,includeMethodology:Boolean(source.includeMethodology),includeFacilityDetails:Boolean(source.includeFacilityDetails),includeNotes:Boolean(source.includeNotes)};
    if(!settings.title||!settings.organization)throw new Error("عنوان التقرير والجهة مطلوبان.");
    await saveReportSettings(year,scope,programFilter,settings);return Response.json({ok:true});
  }catch(error){return Response.json({error:error instanceof Error?error.message:"تعذر حفظ إعدادات التقرير."},{status:400});}
}
