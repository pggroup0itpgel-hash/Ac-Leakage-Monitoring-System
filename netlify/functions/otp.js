const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzkSCykxKGF_HDNvNiEjCMoW25C3HD7DkGHFTYwfEeBpYPgPhWq6RcQTsISiDgsN_5KgA/exec";

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") {
    return {
      statusCode: 204,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST,OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
      },
      body: "",
    };
  }

  try {
    let payload = {};
    if (event.httpMethod === "GET") {
      payload = event.queryStringParameters || {};
      if (!payload.action) {
        return {
          statusCode: 200,
          headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" },
          body: JSON.stringify({
            ok: true,
            message: "OTP proxy is healthy",
            usage: "Use POST /api/otp with JSON { action, ...payload }",
          }),
        };
      }
    } else if (event.httpMethod === "POST") {
      payload = JSON.parse(event.body || "{}");
    } else {
      return {
        statusCode: 200,
        headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" },
        body: JSON.stringify({ ok: false, error: "Unsupported method" }),
      };
    }

    const response = await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    const text = await response.text();
    let data = {};
    try {
      data = JSON.parse(text || "{}");
    } catch {
      data = { ok: false, error: "Invalid response from Apps Script", raw: text };
    }

    return {
      statusCode: response.ok ? 200 : 500,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*",
      },
      body: JSON.stringify(data),
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*",
      },
      body: JSON.stringify({ ok: false, error: err.message || "Proxy failed" }),
    };
  }
};
