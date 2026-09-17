const stripe = require("stripe")(process.env.STRIPE_SECRET_KEY);

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
  const secret  = process.env.STRIPE_CONNECT_WEBHOOK_SECRET;
  const rawBody = await getRawBody(req);

  let event;
  try {
    event = stripe.webhooks.constructEvent(rawBody, sig, secret);
  } catch (err) {
    console.error("[stripe-connect-webhook] Signature failed:", err.message);
    return res.status(400).send(`Webhook error: ${err.message}`);
  }

  if (event.type === "account.updated") {
    const account       = event.data.object;
    const applicationId = account.metadata?.applicationId;

    if (account.charges_enabled && applicationId) {
      const appsScriptUrl = process.env.APPS_SCRIPT_URL;
      const params = new URLSearchParams({
        action:          "update_stripe_account",
        id:              applicationId,
        stripeAccountId: account.id
      });
      try {
        const r = await fetch(`${appsScriptUrl}?${params}`);
        const d = await r.json();
        console.log("[stripe-connect-webhook] update_stripe_account →", d);
      } catch (err) {
        console.error("[stripe-connect-webhook] Apps Script call failed:", err.message);
      }
    }
  }

  return res.status(200).json({ received: true });
};
