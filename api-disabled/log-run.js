module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const SUPABASE_URL = process.env.SUPABASE_URL;
    const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;

    if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
      return res.status(500).json({ error: "Missing Supabase environment variables" });
    }

    const body = typeof req.body === "string" ? JSON.parse(req.body) : (req.body || {});

    const client = String(body.client || "").trim();
    const certificateCount = Number(body.certificate_count);
    const deliveryTypes = Array.isArray(body.delivery_types)
      ? body.delivery_types.map(function(item) { return String(item).trim(); }).filter(Boolean)
      : [];
    const source = String(body.source || "web").trim();
    const notes = body.notes ? String(body.notes).trim() : null;

    const allowedClients = ["liberty", "stanlib", "pm", "alexforbes", "sanlam"];

    if (!allowedClients.includes(client)) {
      return res.status(400).json({ error: "Invalid client" });
    }

    if (!Number.isInteger(certificateCount) || certificateCount < 0) {
      return res.status(400).json({ error: "Invalid certificate_count" });
    }

    const payload = {
      client: client,
      certificate_count: certificateCount,
      delivery_types: deliveryTypes,
      source: source,
      notes: notes
    };

    const response = await fetch(SUPABASE_URL + "/rest/v1/certificate_runs", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": "Bearer " + SUPABASE_SERVICE_ROLE_KEY,
        "Prefer": "return=minimal"
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorText = await response.text();
      return res.status(500).json({
        error: "Supabase insert failed",
        details: errorText
      });
    }

    return res.status(200).json({ ok: true });
  } catch (error) {
    return res.status(500).json({
      error: "Unexpected server error",
      details: error && error.message ? error.message : String(error)
    });
  }
};
