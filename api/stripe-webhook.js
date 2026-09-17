const stripe = require("stripe")(process.env.STRIPE_SECRET_KEY);

// Vercel: disable body parser so we get raw bytes for Stripe signature verification.
module.exports.config = { api: { bodyParser: false } };

function getRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", c => chunks.push(c));
    req.on("end",  () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

module.exports = async (req, res) => {
  if (req.method !== "POST") return res.status(405).end();

  const sig     = req.headers["stripe-signature"];
  const secret  = process.env.STRIPE_WEBHOOK_SECRET;
  const rawBody = await getRawBody(req);

  let event;
  try {
    event = stripe.webhooks.constructEvent(rawBody, sig, secret);
  } catch (err) {
    console.error("[stripe-webhook] Signature verification failed:", err.message);
    return res.status(400).send(`Webhook error: ${err.message}`);
  }

  if (event.type === "checkout.session.completed") {
    const session     = event.data.object;
    const { tutorId, tutorName, studentUid, studentName } = session.metadata || {};

    if (!tutorId || !studentUid) {
      console.warn("[stripe-webhook] Missing tutorId or studentUid in metadata");
      return res.status(200).json({ received: true });
    }

    const appsScriptUrl = process.env.APPS_SCRIPT_URL;
    const params = new URLSearchParams({
      action:          "create_connection",
      tutorId,
      tutorName:       tutorName       || "",
      studentUid,
      studentName:     studentName     || "",
      paymentIntentId: session.payment_intent || "",
      paymentAmount:   String(session.amount_total || 0)
    });

    try {
      const r = await fetch(`${appsScriptUrl}?${params}`);
      const d = await r.json();
      console.log("[stripe-webhook] create_connection →", d);
    } catch (err) {
      console.error("[stripe-webhook] Apps Script call failed:", err.message);
    }
  }

  return res.status(200).json({ received: true });
};
