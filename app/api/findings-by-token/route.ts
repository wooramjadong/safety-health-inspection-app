import { NextRequest, NextResponse } from "next/server";
import { getInspectionByToken, getFindingsByInspection } from "@/lib/sheets";

export async function GET(req: NextRequest) {
  try {
    const token = req.nextUrl.searchParams.get("token");
    if (!token) return NextResponse.json({ error: "token required" }, { status: 400 });

    const insp = await getInspectionByToken(token);
    if (!insp) return NextResponse.json({ error: "invalid token" }, { status: 404 });

    const findings = await getFindingsByInspection(insp.id);
    return NextResponse.json(findings);
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
