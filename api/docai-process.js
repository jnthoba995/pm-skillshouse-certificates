import fs from "fs";
import path from "path";
import { GoogleAuth } from "google-auth-library";
import fetch from "node-fetch";

export default async function handler(req, res) {
  try {
    const { base64, mimeType } = req.body;

    if (!base64) {
      return res.status(400).json({ error: "No document provided" });
    }

    const keyPath = path.join(process.cwd(), "secrets", "docai-key.json");
    const key = JSON.parse(fs.readFileSync(keyPath, "utf8"));

    const auth = new GoogleAuth({
      credentials: key,
      scopes: ["https://www.googleapis.com/auth/cloud-platform"],
    });

    const client = await auth.getClient();
    const token = await client.getAccessToken();

    const endpoint =
      "https://us-documentai.googleapis.com/v1/projects/63403376518/locations/us/processors/f068d45b33bc417d:process";

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token.token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        rawDocument: {
          content: base64,
          mimeType: mimeType || "application/pdf",
        },
      }),
    });

    const data = await response.json();

    return res.status(200).json(data);
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: "Processing failed" });
  }
}
