import { Resend } from "resend";

const resend = new Resend(process.env.RESEND_API_KEY);

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" });
    }

    const { email, name, pdfBase64 } = req.body;

    if (!email || !name || !pdfBase64) {
      return res.status(400).json({ error: "Missing required fields" });
    }

    // Convert base64 to buffer
    const pdfBuffer = Buffer.from(pdfBase64.split(",")[1], "base64");

    const response = await resend.emails.send({
      from: "PM SkillsHouse <onboarding@resend.dev>",
      to: email,
      subject: "Your Certificate 🎓",
      html: `
        <div style="font-family: Arial, sans-serif;">
          <h2>Congratulations ${name} 🎉</h2>
          <p>Your certificate is attached.</p>
          <p>Keep up the great work.</p>
        </div>
      `,
      attachments: [
        {
          filename: "certificate.pdf",
          content: pdfBuffer,
        },
      ],
    });

    return res.status(200).json({ success: true, response });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: "Email sending failed" });
  }
}
