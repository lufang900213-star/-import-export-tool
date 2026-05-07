const TABLE = process.env.SUPABASE_TABLE || "waybill_records";

async function insertRows(rows) {
  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_ANON_KEY;
  if (!url || !key) return { mode: "local_fallback", data: rows, successCount: rows.length, failedCount: 0 };

  const endpoint = `${url}/rest/v1/${TABLE}`;
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
      apikey: key,
      Authorization: `Bearer ${key}`,
    },
    body: JSON.stringify(rows),
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(message || `Supabase write failed: ${response.status}`);
  }

  const inserted = await response.json();
  const listResponse = await fetch(`${url}/rest/v1/${TABLE}?select=*&order=submittedAt.desc`, {
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
    },
  });
  const allData = listResponse.ok ? await listResponse.json() : inserted;
  return {
    mode: "supabase",
    data: allData,
    successCount: inserted.length,
    failedCount: 0,
  };
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];
    const result = await insertRows(rows);
    res.status(200).json(result);
  } catch (error) {
    res.status(200).json({
      mode: "local_fallback",
      data: Array.isArray(req.body?.rows) ? req.body.rows : [],
      successCount: Array.isArray(req.body?.rows) ? req.body.rows.length : 0,
      failedCount: 0,
      error: error.message,
    });
  }
};
