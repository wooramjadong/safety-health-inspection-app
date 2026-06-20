import { NextRequest, NextResponse } from "next/server";
import { createFinding, getFindingsByInspection, FindingInput } from "@/lib/sheets";

export async function GET(req: NextRequest) {
  try {
    const inspectionId = req.nextUrl.searchParams.get("inspectionId");
    if (!inspectionId) return NextResponse.json({ error: "inspectionId required" }, { status: 400 });
    const data = await getFindingsByInspection(inspectionId);
    return NextResponse.json(data);
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}

export async function POST(req: NextRequest) {
  try {
    const body: FindingInput = await req.json();
    const id = await createFinding(body);
    return NextResponse.json({ id }, { status: 201 });
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
