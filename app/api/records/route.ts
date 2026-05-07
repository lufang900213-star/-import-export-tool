import { NextResponse } from "next/server";
import { listRecords } from "../../../lib/db";

export async function GET() {
  try {
    const data = await listRecords();
    return NextResponse.json({ data });
  } catch (error) {
    return NextResponse.json(
      { data: [], error: error instanceof Error ? error.message : "读取失败" },
      { status: 500 },
    );
  }
}
