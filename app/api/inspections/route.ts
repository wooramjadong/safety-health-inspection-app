import { NextResponse } from "next/server";
import { createInspection } from "@/lib/sheets";

export async function POST(request: Request) {
  const form = await request.formData();
  const result = await createInspection({
    type: String(form.get("type") ?? ""),
    siteName: String(form.get("siteName") ?? ""),
    inspectors: String(form.get("inspectors") ?? "")
  });
  return NextResponse.json(result);
}
