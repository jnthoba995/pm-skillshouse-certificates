export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    const { email, name, pdfBase64, subject, message, senderName } = req.body;

    if (!email || !name || !pdfBase64) {
      return res.status(400).json({ error: "Missing required fields" });
    }

    const cleanBase64 = pdfBase64.includes(",")
      ? pdfBase64.split(",")[1]
      : pdfBase64;

    const resendResponse = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${process.env.RESEND_API_KEY}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        from: `${senderName || "PM SkillsHouse"} <onboarding@resend.dev>`,
        to: [email],
        subject: subject: subject || `Your Certificate of Attendance - ${name}`,
        html: `
          <div style="font-family: Arial, sans-serif; line-height: 1.6;">
            <h2>Congratulations ${name}</h2>
            <p>${(message || "Please find your certificate attached.")
              .replace(/\n/g, "<br>")}</p>
          </div>
        `,
        attachments: [
          {
            filename: `${name} - Certificate.pdf`,
            content: cleanBase64
          }
        ]
      })
    });

    const data = await resendResponse.json();

    if (!resendResponse.ok) {
      return res.status(500).json({
        error: "Email sending failed",
        details: data
      });
    }

    return res.status(200).json({
      success: true,
      response: data
    });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Email sending failed" });
  }
}
