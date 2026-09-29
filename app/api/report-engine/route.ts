import source from "@/report_generator.py?raw";
import base from "@/python_engine/base_pipeline.py?raw";
import catalog from "@/lib/facilities.json";

export function GET() {
  return Response.json({ source, base, catalog }, { headers: { "Cache-Control": "no-store" } });
}
