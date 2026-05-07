import type { PersistedRecord } from "./types";

interface RecordsResponse {
  data: PersistedRecord[];
}

interface SubmitResponse extends RecordsResponse {
  successCount: number;
  failedCount: number;
}

export async function fetchRecords(): Promise<PersistedRecord[]> {
  const response = await fetch("/api/records", { cache: "no-store" });
  if (!response.ok) {
    throw new Error("读取历史运单失败");
  }
  const payload = (await response.json()) as RecordsResponse;
  return payload.data || [];
}

export async function submitRecords(rows: PersistedRecord[]): Promise<SubmitResponse> {
  const response = await fetch("/api/submit", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ rows }),
  });
  if (!response.ok) {
    throw new Error("提交运单失败");
  }
  return (await response.json()) as SubmitResponse;
}
