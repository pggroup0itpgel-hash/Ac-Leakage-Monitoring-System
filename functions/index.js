const functions = require("firebase-functions");
const admin = require("firebase-admin");
const nodemailer = require("nodemailer");
const crypto = require("crypto");

admin.initializeApp();
const db = admin.database();

const SMTP_HOST = process.env.SMTP_HOST || "";
const SMTP_PORT = Number(process.env.SMTP_PORT || 587);
const SMTP_USER = process.env.SMTP_USER || "";
const SMTP_PASS = process.env.SMTP_PASS || "";
const SMTP_FROM = process.env.SMTP_FROM || SMTP_USER || "no-reply@example.com";
const OTP_TTL_MS = 10 * 60 * 1000;

const transporter = nodemailer.createTransport({
  host: SMTP_HOST,
  port: SMTP_PORT,
  secure: SMTP_PORT === 465,
  auth: SMTP_USER && SMTP_PASS ? { user: SMTP_USER, pass: SMTP_PASS } : undefined,
});

function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

function emailKey(email) {
  return normalizeEmail(email).replace(/\./g, ",");
}

function sha256(value) {
  return crypto.createHash("sha256").update(String(value)).digest("hex");
}

function randomOtp() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function randomToken() {
  return crypto.randomBytes(32).toString("hex");
}

exports.sendOtp = functions.https.onRequest(async (req, res) => {
  res.set("Access-Control-Allow-Origin", "*");
  res.set("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.set("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(204).send("");
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const email = normalizeEmail(req.body && req.body.email);
    if (!email) return res.status(400).json({ error: "Email required" });

    const userSnap = await db.ref(`appConfig/users/${emailKey(email)}`).once("value");
    if (!userSnap.exists()) return res.status(403).json({ error: "Email not authorized" });

    const otp = randomOtp();
    const otpHash = sha256(otp);
    const expiresAt = Date.now() + OTP_TTL_MS;
    const otpRef = db.ref("authOtp").push();
    await otpRef.set({
      email,
      otpHash,
      expiresAt,
      attempts: 0,
      createdAt: Date.now(),
      used: false,
    });

    if (!SMTP_HOST || !SMTP_USER || !SMTP_PASS) {
      return res.status(500).json({ error: "SMTP not configured on server" });
    }

    await transporter.sendMail({
      from: SMTP_FROM,
      to: email,
      subject: "AC Leakage OTP Login",
      text: `Your OTP is ${otp}. It is valid for 10 minutes.`,
      html: `<p>Your OTP is <b>${otp}</b>.</p><p>It is valid for 10 minutes.</p>`,
    });

    return res.json({ ok: true, message: "OTP sent" });
  } catch (err) {
    console.error("sendOtp error", err);
    return res.status(500).json({ error: err.message || "Failed to send OTP" });
  }
});

exports.verifyOtp = functions.https.onRequest(async (req, res) => {
  res.set("Access-Control-Allow-Origin", "*");
  res.set("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.set("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(204).send("");
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const email = normalizeEmail(req.body && req.body.email);
    const otp = String((req.body && req.body.otp) || "").trim();
    if (!email || !otp) return res.status(400).json({ error: "Email and OTP required" });

    const snap = await db.ref("authOtp").orderByChild("email").equalTo(email).once("value");
    let best = null;
    snap.forEach((child) => {
      const v = child.val() || {};
      if (!best || (v.createdAt || 0) > (best.value.createdAt || 0)) {
        best = { key: child.key, value: v };
      }
    });
    if (!best) return res.status(400).json({ error: "OTP not found" });

    const record = best.value;
    if (record.used) return res.status(400).json({ error: "OTP already used" });
    if (Date.now() > Number(record.expiresAt || 0)) return res.status(400).json({ error: "OTP expired" });
    if (Number(record.attempts || 0) >= 5) return res.status(429).json({ error: "Too many attempts" });

    const otpRef = db.ref(`authOtp/${best.key}`);
    const valid = sha256(otp) === record.otpHash;
    if (!valid) {
      await otpRef.update({ attempts: Number(record.attempts || 0) + 1 });
      return res.status(401).json({ error: "Invalid OTP" });
    }

    await otpRef.update({ used: true, usedAt: Date.now() });
    const userSnap = await db.ref(`appConfig/users/${emailKey(email)}`).once("value");
    if (!userSnap.exists()) return res.status(403).json({ error: "Email not authorized" });

    const token = randomToken();
    const tokenHash = sha256(token);
    const tokenExpiresAt = Date.now() + 7 * 24 * 60 * 60 * 1000;
    await db.ref(`authSessions/${tokenHash}`).set({
      email,
      createdAt: Date.now(),
      expiresAt: tokenExpiresAt,
    });

    return res.json({
      ok: true,
      token,
      tokenExpiresAt,
      profile: userSnap.val(),
      user: { email },
    });
  } catch (err) {
    console.error("verifyOtp error", err);
    return res.status(500).json({ error: err.message || "OTP verification failed" });
  }
});
