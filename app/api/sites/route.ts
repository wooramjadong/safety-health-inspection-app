import { NextRequest, NextResponse } from "next/server";
import { createSite, getSites, SiteInput } from "@/lib/sheets";

export async function GET() {
  try {
    const data = await getSites();
    return NextResponse.json(data);
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}

export async function POST(req: NextRequest) {
  try {
    const body: SiteInput = await req.json();
    const id = await createSite(body);
    return NextResponse.json({ id }, { status: 201 });
  } catch (e: any) {
    return NextResponse.json({ error: e.message }, { status: 500 });
  }
}
