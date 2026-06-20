import { NextResponse } from "next/server";
import { saveActionSubmit } from "@/lib/sheets";

export async function POST(request: Request) {
  const form = await request.formData();
  const result = await saveActionSubmit({
    token: String(form.get("token") ?? ""),
    actionText: String(form.get("actionText") ?? "")
  });
  return NextResponse.json(result);
}
