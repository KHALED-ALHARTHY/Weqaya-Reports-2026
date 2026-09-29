import { periodLabel, reportNarrative, type ProgramResult, type ReportScope } from "./report";

export type ReportSettings = {
  title: string;
  organization: string;
  executiveSummary: string;
  methodology: string;
  conclusion: string;
  reviewedBy: string;
  approvedBy: string;
  programNotes: Record<string, string>;
  includeMethodology: boolean;
  includeFacilityDetails: boolean;
  includeNotes: boolean;
};

export function defaultReportSettings(year:number, scope:ReportScope, reports:ProgramResult[]):ReportSettings {
  const complete=reports.reduce((sum,item)=>sum+item.complete,0);
  const total=reports.reduce((sum,item)=>sum+item.total,0);
  const partial=reports.reduce((sum,item)=>sum+item.partial,0);
  return {
    title:"تقرير برامج مكافحة العدوى",
    organization:"هيئة الصحة العامة، مكتب مكة المكرمة",
    executiveSummary:`يتناول التقرير ${reports.length} برامج و${total} منشأة، منها ${complete} منشأة مكتملة البيانات خلال ${periodLabel(scope)} لعام ${year}، و${partial} منشأة لديها إدخالات جزئية. النتائج مأخوذة من إدخالات الموقع فقط ولا تتضمن النسب المرجعية السابقة.`,
    methodology:"تحسب نتيجة المنشأة كمتوسط بسيط لفتراتها المطلوبة، ثم يحسب متوسط البرنامج من المنشآت المكتملة. لا تُحسب الفترة الربع سنوية لنظافة اليدين حتى تكتمل أشهرها الثلاثة، ويُعرض PHC ICA حسب النصف السنوي. لا يمثل المتوسط معدلًا موزونًا بفرص الرصد.",
    conclusion:"تراجع أسباب التغير وخطط التصحيح من الملاحظات المسجلة مع كل منشأة. قبل اعتماد التقرير السنوي، تستكمل الفترات الناقصة وتراجع التغطية وصحة البيانات.",
    reviewedBy:"",approvedBy:"",
    programNotes:Object.fromEntries(reports.map((item)=>[item.program,reportNarrative(item,scope)])),
    includeMethodology:true,includeFacilityDetails:true,includeNotes:true,
  };
}
