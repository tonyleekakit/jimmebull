export async function onRequestPost(context) {
  const { request, env } = context;

  // 1. Get origin for success/cancel URLs
  const url = new URL(request.url);
  const origin = url.origin; // e.g., https://example.com or http://localhost:8788

  // 2. Validate Env
  const stripeKey = env.STRIPE_SECRET_KEY;
  if (!stripeKey) {
    return new Response(JSON.stringify({ error: "Missing STRIPE_SECRET_KEY" }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }

  try {
    // 3. Create Stripe Checkout Session via Fetch API (no npm package needed)
    // We use URLSearchParams for x-www-form-urlencoded body
    const body = new URLSearchParams();
    
    // Product Details
    body.append("payment_method_types[]", "card");
    body.append("line_items[0][price_data][currency]", "hkd");
    body.append("line_items[0][price_data][product_data][name]", "完整訓練計畫 (52週)");
    body.append("line_items[0][price_data][unit_amount]", "8800"); // $88.00 HKD
    body.append("line_items[0][quantity]", "1");
    
    // Mode
    body.append("mode", "payment");
    
    // Redirect URLs
    body.append("success_url", `${origin}/?payment=success`);
    body.append("cancel_url", `${origin}/?payment=cancelled`);

    const stripeResponse = await fetch("https://api.stripe.com/v1/checkout/sessions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${stripeKey}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: body,
    });

    const session = await stripeResponse.json();

    if (session.error) {
      throw new Error(session.error.message);
    }

    // 4. Return the checkout URL
    return new Response(JSON.stringify({ url: session.url }), {
      headers: { "Content-Type": "application/json" },
    });

  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }
}
