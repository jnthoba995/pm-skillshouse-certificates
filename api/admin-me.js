module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "GET") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const SUPABASE_URL = process.env.SUPABASE_URL;
    const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
    const authHeader = req.headers.authorization || "";

    if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
      return res.status(500).json({ error: "Missing Supabase environment variables" });
    }

    if (!authHeader.startsWith("Bearer ")) {
      return res.status(401).json({ error: "Missing bearer token" });
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
      SUPABASE_URL + "/rest/v1/admin_users?id=eq." + encodeURIComponent(user.id) + "&select=id,email,full_name,role,is_active",
      {
        method: "GET",
        headers: {
          "apikey": SUPABASE_SERVICE_ROLE_KEY,
          "Authorization": "Bearer " + SUPABASE_SERVICE_ROLE_KEY
        }
      }
    );

    if (!adminResponse.ok) {
      const adminText = await adminResponse.text();
      return res.status(500).json({ error: "Failed to load admin record", details: adminText });
    }

    const rows = await adminResponse.json();
    const adminUser = Array.isArray(rows) ? rows[0] : null;

    if (!adminUser || !adminUser.is_active) {
      return res.status(403).json({ error: "Not authorized" });
    }

    return res.status(200).json({
      ok: true,
      user: {
        id: user.id,
        email: user.email
      },
      admin: adminUser
    });
  } catch (error) {
    return res.status(500).json({
      error: "Unexpected server error",
      details: error && error.message ? error.message : String(error)
    });
  }
};
