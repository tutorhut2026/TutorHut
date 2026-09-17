const stripe = require("stripe")(process.env.STRIPE_SECRET_KEY);

function cors(res) {
  res.setHeader("Access-Control-Allow-Origin",  process.env.SITE_URL || "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
}

module.exports = async (req, res) => {
  cors(res);
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST")   return res.status(405).json({ error: "Method not allowed" });

  const { tutorId, tutorName, studentUid, studentName } = req.body || {};
  if (!tutorId)     return res.status(400).json({ error: "tutorId is required" });
  if (!studentUid)  return res.status(400).json({ error: "studentUid is required" });

  const feePence   = parseInt(process.env.INTRO_FEE_PENCE || "1500", 10);
  const siteUrl    = (process.env.SITE_URL || "").replace(/\/$/, "");

  try {
    const session = await stripe.checkout.sessions.create({
      mode: "payment",
      line_items: [{
        price_data: {
          currency:     "gbp",
          product_data: { name: `TutorHut – Connect with ${tutorName || "Tutor"}` },
          unit_amount:  feePence
        },
        quantity: 1
      }],
      metadata:    { tutorId, tutorName: tutorName || "", studentUid, studentName: studentName || "" },
      success_url: `${siteUrl}/student-dashboard.html?payment=success`,
      cancel_url:  `${siteUrl}/student-dashboard.html?payment=cancelled`
    });

    return res.status(200).json({ url: session.url });
  } catch (err) {
    console.error("[create-checkout]", err.message);
    return res.status(500).json({ error: err.message });
  }
};
