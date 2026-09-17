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

  const { applicationId } = req.body || {};
  if (!applicationId) return res.status(400).json({ error: "applicationId is required" });

  const appsScriptUrl = process.env.APPS_SCRIPT_URL;
  const siteUrl       = (process.env.SITE_URL || "").replace(/\/$/, "");

  try {
    // Fetch the tutor's application row to check for an existing stripeAccountId.
    const r   = await fetch(`${appsScriptUrl}?${new URLSearchParams({ action: "get_tutor_by_id", id: applicationId })}`);
    const app = await r.json();

    if (!app || !app.id) {
      return res.status(404).json({ error: "Tutor application not found or not approved" });
    }

    let stripeAccountId = app.stripeAccountId || null;

    if (!stripeAccountId) {
      // Create a new Stripe Connect Express account for this tutor.
      const account = await stripe.accounts.create({
        type:    "express",
        country: "GB",
        email:   app.email || undefined,
        metadata: { applicationId }
      });
      stripeAccountId = account.id;

      // Persist immediately so we never create duplicate accounts.
      await fetch(`${appsScriptUrl}?${new URLSearchParams({
        action:          "update_stripe_account",
        id:              applicationId,
        stripeAccountId: stripeAccountId
      })}`);
    }

    // Create (or re-create) the onboarding link.
    const link = await stripe.accountLinks.create({
      account:     stripeAccountId,
      refresh_url: `${siteUrl}/tutor-dashboard.html?stripe=refresh`,
      return_url:  `${siteUrl}/tutor-dashboard.html?stripe=complete`,
      type:        "account_onboarding"
    });

    return res.status(200).json({ url: link.url });
  } catch (err) {
    console.error("[create-connect-onboarding]", err.message);
    return res.status(500).json({ error: err.message });
  }
};
