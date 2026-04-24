module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const SUPABASE_URL = process.env.SUPABASE_URL;
    const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
    const authHeader = req.headers.authorization || "";
    const body = req.body || {};
    const targetId = String(body.id || "").trim();
    const isActive = body.is_active;

    if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
      return res.status(500).json({ error: "Missing Supabase environment variables" });
    }

    if (!authHeader.startsWith("Bearer ")) {
      return res.status(401).json({ error: "Missing bearer token" });
    }

    if (!targetId || typeof isActive !== "boolean") {
      return res.status(400).json({ error: "Missing id or is_active" });
    }

    const accessToken = authHeader.slice("Bearer ".length).trim();

    const userResponse = await fetch(SUPABASE_URL + "/auth/v1/user", {
      method: "GET",
      headers: {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": "Bearer " + accessToken
      }
    });

    if (!userResponse.ok) {
      const userText = await userResponse.text();
      return res.status(401).json({ error: "Invalid session", details: userText });
    }

    const user = await userResponse.json();

    const adminResponse = await fetch(
      SUPABASE_URL + "/rest/v1/admin_users?id=eq." + encodeURIComponent(user.id) + "&select=id,email,role,is_active",
      {
        method: "GET",
        headers: {
          "apikey": SUPABASE_SERVICE_ROLE_KEY,
          "Authorization": "Bearer " + SUPABASE_SERVICE_ROLE_KEY,
          "Accept": "application/json"
        }
      }
    );

    if (!adminResponse.ok) {
      const adminText = await adminResponse.text();
      return res.status(500).json({ error: "Failed to verify admin user", details: adminText });
    }

    const adminRows = await adminResponse.json();
    const adminUser = Array.isArray(adminRows) ? adminRows[0] : null;

    if (!adminUser || !adminUser.is_active || adminUser.role !== "super_admin") {
      return res.status(403).json({ error: "Not allowed" });
    }

    if (adminUser.id === targetId && isActive === false) {
      return res.status(400).json({ error: "You cannot deactivate your own super admin account." });
    }

    const updateResponse = await fetch(
      SUPABASE_URL + "/rest/v1/admin_users?id=eq." + encodeURIComponent(targetId),
      {
        method: "PATCH",
        headers: {
          "apikey": SUPABASE_SERVICE_ROLE_KEY,
          "Authorization": "Bearer " + SUPABASE_SERVICE_ROLE_KEY,
          "Content-Type": "application/json",
          "Prefer": "return=representation"
        },
        body: JSON.stringify({ is_active: isActive })
      }
    );

    if (!updateResponse.ok) {
      const updateText = await updateResponse.text();
      return res.status(500).json({ error: "Failed to update user", details: updateText });
    }

    const rows = await updateResponse.json();

    return res.status(200).json({
      ok: true,
      user: Array.isArray(rows) ? rows[0] : null
    });
  } catch (error) {
    return res.status(500).json({
      error: "Unexpected server error",
      details: error && error.message ? error.message : String(error)
    });
  }
};
