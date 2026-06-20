import { NextResponse } from "next/server";
import { suggestFinding } from "@/lib/ai";

export async function POST(request: Request) {
  const form = await request.formData();
  const memo = String(form.get("memo") ?? "");
  const result = await suggestFinding({ memo });
  return NextResponse.json(result);
}
