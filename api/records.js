const TABLE = process.env.SUPABASE_TABLE || "waybill_records";

async function loadRecords() {
  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_ANON_KEY;
  if (!url || !key) return { mode: "local_fallback", data: [] };

  const endpoint = `${url}/rest/v1/${TABLE}?select=*&order=submittedAt.desc`;
  const response = await fetch(endpoint, {
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
    },
  });
  if (!response.ok) {
    return { mode: "local_fallback", data: [], error: `Supabase read failed: ${response.status}` };
  }
  const data = await response.json();
  return { mode: "supabase", data };
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }
  try {
    const result = await loadRecords();
    res.status(200).json(result);
  } catch (error) {
    res.status(200).json({ mode: "local_fallback", data: [], error: error.message });
  }
};
