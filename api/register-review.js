import { GoogleAuth } from "google-auth-library";

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    const { fileBase64, mimeType } = req.body || {};

    if (!fileBase64) {
      return res.status(400).json({ error: "No document provided" });
    }

    const key = JSON.parse(process.env.GOOGLE_DOC_AI_KEY);

    const auth = new GoogleAuth({
      credentials: key,
      scopes: ["https://www.googleapis.com/auth/cloud-platform"],
    });

    const client = await auth.getClient();
    const token = await client.getAccessToken();

    const endpoint = "https://us-documentai.googleapis.com/v1/projects/63403376518/locations/us/processors/f068d45b33bc417d:process";

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token.token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        rawDocument: {
          content: fileBase64,
          mimeType: mimeType || "application/pdf",
        },
      }),
    });

    const docai = await response.json();

    if (!response.ok) {
      return res.status(response.status).json({
        error: "Document AI request failed",
        details: docai,
      });
    }

    const rows = extractRowsFromDocumentAI(docai);

    return res.status(200).json({
      status: "success",
      message: "Register processed. Please review highlighted fields before exporting.",
      rows,
      raw: docai,
    });
  } catch (err) {
    console.error("REGISTER REVIEW ERROR:", err);
    return res.status(500).json({ error: "Register review failed" });
  }
}

function extractRowsFromDocumentAI(docai) {
  const pages = docai && docai.document && docai.document.pages ? docai.document.pages : [];
  const text = docai && docai.document && docai.document.text ? docai.document.text : "";

  const rows = [];

  pages.forEach((page) => {
    const tables = page.tables || [];

    tables.forEach((table) => {
      const bodyRows = table.bodyRows || [];

      bodyRows.forEach((row) => {
        const cells = row.cells || [];
        const values = cells.map((cell) => getTextFromAnchor(text, cell.layout && cell.layout.textAnchor).trim());

        if (values.length < 3) return;

        const name = cleanCell(values[0]);
        const surname = cleanCell(values[1]);
        const idNumber = cleanCell(values[2]);
        const contact = cleanCell(values[8] || values[7] || "");
        const gender = detectChoice(values.join(" "), ["F", "M"]);
        const race = detectChoice(values.join(" "), ["B", "C", "I", "W"]);

        if (!name && !surname && !idNumber) return;

        const joinedValue = values.join(" ").toLowerCase();
        if (
          joinedValue.includes("assupol collects") ||
          joinedValue.includes("personal information") ||
          joinedValue.includes("privacy") ||
          joinedValue.includes("reporting and auditing") ||
          joinedValue.includes("consent to assupol")
        ) return;

        rows.push({
          name,
          surname,
          idNumber,
          contact,
          email: "",
          gender,
          race,
          status: needsReview(name, surname, idNumber) ? "Review" : "OK",
        });
      });
    });
  });

  return rows;
}

function getTextFromAnchor(fullText, textAnchor) {
  if (!textAnchor || !textAnchor.textSegments) return "";

  return textAnchor.textSegments.map((segment) => {
    const start = Number(segment.startIndex || 0);
    const end = Number(segment.endIndex || 0);
    return fullText.slice(start, end);
  }).join("");
}

function cleanCell(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .replace(/[|[\]{}]/g, "")
    .trim();
}

function detectChoice(value, choices) {
  const upper = String(value || "").toUpperCase();
  return choices.find((choice) => upper.includes(choice)) || "";
}

function needsReview(name, surname, idNumber) {
  const digits = String(idNumber || "").replace(/\D/g, "");
  return !name || !surname || digits.length < 10 || /\d/.test(name) || /\d/.test(surname);
}
