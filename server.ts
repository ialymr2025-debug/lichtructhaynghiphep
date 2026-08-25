import express from "express";
// import { createServer as createViteServer } from "vite"; // Move to dynamic import
import path from "path";
import { google } from "googleapis";
import cookieParser from "cookie-parser";
import dotenv from "dotenv";
import https from "https";
import http from "http";

import fs from "fs";
import crypto from "crypto";
import nodemailer from "nodemailer";
import { pool as sqlPool, initLeaveTables, calcTravelDays, calcEntitledDays, calcSeniority } from "./db.js";
import { listWorkingShifts, allocateLeaveDays, type ShiftConfig } from "./leaveCalc.js";
dotenv.config();

initLeaveTables().catch((e: any) => console.error("Failed to initialize SQL leave tables:", e.message));

type LeaveNotifyData = {
  name: string;
  chucDanh: string;
  kip: string;
  startDate: string;
  endDate: string;
  reason: string;
  phone: string;
  location: string;
  dateStr: string;
  leaveBalance?: { entitled: string; used: string; remaining: string } | null;
};

// `to` records where the message was actually addressed, so a wrong-recipient report can
// be answered from the response itself instead of guessing at the stored config.
type NotifyResult = { ok: boolean; skipped?: boolean; error?: string; to?: string };

function buildNotifyText(data: LeaveNotifyData): string {
  let text = `CÓ ĐƠN NGHỈ PHÉP MỚI
• Họ và tên: ${data.name}
• Chức danh: ${data.chucDanh}
• Kíp: Kíp ${data.kip}
• Thời gian: Từ ${data.startDate} đến ${data.endDate}
• Lý do: ${data.reason || "Giải quyết việc riêng gia đình"}`;

  if (data.leaveBalance) {
    text += `
• Phép được hưởng: ${data.leaveBalance.entitled} ngày
• Phép đã nghỉ: ${data.leaveBalance.used} ngày
• Phép còn lại: ${data.leaveBalance.remaining} ngày`;
  }
  return text;
}

// Sends to the webhook configured on THIS workshop's own row (workshops.config.zaloWebhookUrl),
// not the old global app_settings document — previously every workshop silently shared one
// webhook regardless of what each admin configured. Returns an honest result instead of firing
// and forgetting, so the caller can report what actually happened rather than assuming success.
async function sendZaloNotification(workshopId: string | undefined, data: LeaveNotifyData): Promise<NotifyResult> {
  let webhookUrl = process.env.ZALO_WEBHOOK_URL || "";

  if (!webhookUrl && workshopId) {
    try {
      const result = await sqlPool.query(`SELECT config FROM workshops WHERE id = $1`, [workshopId]);
      webhookUrl = result.rows[0]?.config?.zaloWebhookUrl || "";
    } catch (dbErr: any) {
      console.log("Unable to load workshop zaloWebhookUrl:", dbErr.message);
    }
  }

  if (!webhookUrl) {
    return { ok: false, skipped: true, error: "Chưa cấu hình Zalo Webhook URL cho phân xưởng này." };
  }

  const textMessage = buildNotifyText(data);
  const payload = JSON.stringify({
    text: textMessage,
    message: textMessage,
    content: textMessage,
    data: {
      name: data.name,
      chucDanh: data.chucDanh,
      kip: data.kip,
      startDate: data.startDate,
      endDate: data.endDate,
      reason: data.reason,
      phone: data.phone,
      location: data.location,
      createdAt: data.dateStr,
      leaveBalance: data.leaveBalance || null
    }
  });

  return new Promise((resolve) => {
    let parsedUrl: URL;
    try {
      parsedUrl = new URL(webhookUrl);
    } catch (e: any) {
      resolve({ ok: false, error: "Webhook URL không hợp lệ: " + e.message });
      return;
    }

    // http:// was silently swallowed before (only https.request was ever used); pick the
    // matching module so a plain-http webhook actually gets called instead of failing silently.
    const transport = parsedUrl.protocol === "http:" ? http : https;
    const options = {
      hostname: parsedUrl.hostname,
      port: parsedUrl.port || undefined,
      path: parsedUrl.pathname + parsedUrl.search,
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Content-Length": Buffer.byteLength(payload)
      },
      timeout: 8000
    };

    const req = transport.request(options, (res) => {
      let body = "";
      res.on("data", (chunk) => { body += chunk; });
      res.on("end", () => {
        const ok = (res.statusCode || 0) >= 200 && (res.statusCode || 0) < 300;
        console.log(`Zalo notification response (status ${res.statusCode}): ${body}`);
        resolve(ok ? { ok: true } : { ok: false, error: `Webhook trả về mã lỗi ${res.statusCode}` });
      });
    });

    req.on("error", (e: any) => {
      console.log("Zalo Webhook unreachable (e.g. offline tunnel or invalid URL):", e.message);
      resolve({ ok: false, error: "Không kết nối được webhook: " + e.message });
    });

    req.on("timeout", () => {
      req.destroy();
      resolve({ ok: false, error: "Webhook không phản hồi (hết thời gian chờ)." });
    });

    req.write(payload);
    req.end();
  });
}

// Lazily created once and reused across requests; nodemailer keeps its own connection pool.
let mailTransporter: nodemailer.Transporter | null | undefined;
function getMailTransporter(): nodemailer.Transporter | null {
  if (mailTransporter !== undefined) return mailTransporter;
  const user = process.env.GMAIL_USER;
  const pass = process.env.GMAIL_APP_PASSWORD;
  if (!user || !pass) {
    console.log("GMAIL_USER/GMAIL_APP_PASSWORD not set — email notifications disabled.");
    mailTransporter = null;
    return mailTransporter;
  }
  mailTransporter = nodemailer.createTransport({ service: "gmail", auth: { user, pass } });
  return mailTransporter;
}

// Reads the recipient address(es) from THIS workshop's own config.notifyEmail (comma-separated),
// mirroring how the Zalo webhook is scoped per workshop rather than sent from one shared address.
async function sendEmailNotification(workshopId: string | undefined, data: LeaveNotifyData): Promise<NotifyResult> {
  const transporter = getMailTransporter();
  if (!transporter) return { ok: false, skipped: true, error: "Chưa cấu hình Gmail gửi thư trên máy chủ." };
  if (!workshopId) return { ok: false, skipped: true, error: "Thiếu phân xưởng." };

  let recipients = "";
  try {
    const result = await sqlPool.query(`SELECT config FROM workshops WHERE id = $1`, [workshopId]);
    recipients = String(result.rows[0]?.config?.notifyEmail || "").trim();
  } catch (dbErr: any) {
    return { ok: false, error: "Không đọc được cấu hình email: " + dbErr.message };
  }
  if (!recipients) {
    return { ok: false, skipped: true, error: "Chưa cấu hình Email nhận thông báo cho phân xưởng này." };
  }

  const textMessage = buildNotifyText(data);
  const htmlMessage = textMessage
    .split("\n")
    .map((line) => `<p style="margin:0 0 4px">${line.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</p>`)
    .join("");

  try {
    await transporter.sendMail({
      from: `"Hệ thống Lịch trực thay ca" <${process.env.GMAIL_USER}>`,
      to: recipients,
      subject: `Đơn nghỉ phép mới: ${data.name} (Kíp ${data.kip}, ${data.startDate} - ${data.endDate})`,
      text: textMessage,
      html: htmlMessage
    });
    console.log(`Email notification sent to: ${recipients} (workshop ${workshopId})`);
    return { ok: true, to: recipients };
  } catch (e: any) {
    console.log("Failed to send email notification:", e.message);
    return { ok: false, error: "Gửi email thất bại: " + e.message, to: recipients };
  }
}

let lastDbError: string | null = null;

// app_config replaces the old Firestore "config" collection: one row per document key.
async function readConfig(key: string): Promise<any | null> {
  try {
    const result = await sqlPool.query(`SELECT value FROM app_config WHERE key = $1`, [key]);
    lastDbError = null;
    return result.rows[0]?.value ?? null;
  } catch (e: any) {
    lastDbError = `Read Error (${key}): ${e.message}`;
    console.error(`Error reading app_config/${key}:`, e.message);
    return null;
  }
}

async function writeConfig(key: string, value: any): Promise<boolean> {
  try {
    await sqlPool.query(
      `INSERT INTO app_config (key, value) VALUES ($1, $2)
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = now()`,
      [key, JSON.stringify(value)]
    );
    lastDbError = null;
    return true;
  } catch (e: any) {
    lastDbError = `Write Error (${key}): ${e.message}`;
    console.error(`Error writing app_config/${key}:`, e.message);
    return false;
  }
}

const saveTokens = async (newTokens: any) => {
  const existing = await readConfig("google_auth");
  const finalTokens = { ...newTokens };

  // Google only returns refresh_token on first consent, so never lose the stored one.
  if (!finalTokens.refresh_token && existing?.tokens?.refresh_token) {
    finalTokens.refresh_token = existing.tokens.refresh_token;
  }

  await writeConfig("google_auth", { tokens: finalTokens, updatedAt: new Date().toISOString() });
};

const loadTokens = async () => {
  const doc = await readConfig("google_auth");
  return doc?.tokens || null;
};

const clearTokens = async () => {
  try {
    await sqlPool.query(`DELETE FROM app_config WHERE key = 'google_auth'`);
    lastDbError = null;
  } catch (e: any) {
    lastDbError = `Clear Error: ${e.message}`;
    console.error("Error clearing tokens:", e.message);
  }
};

const handleAuthErrorIfAny = async (error: any, res: any) => {
  const errorMsg = (error.message || "").toLowerCase();
  const isAuthError = 
    errorMsg.includes("invalid_grant") || 
    errorMsg.includes("insufficient authentication scopes") ||
    errorMsg.includes("scope") ||
    error.code === 400 || 
    error.code === 401 ||
    error.code === 403;

  if (isAuthError) {
    console.warn("Auth error detected in sheets API:", error.message, ". Clearing tokens and resetting status.");
    res.clearCookie("google_tokens");
    await clearTokens();
    res.status(401).json({ 
      error: "Liên kết tài khoản Google đã hết hạn hoặc bị hủy. Vui lòng kết nối lại tài khoản của bạn.",
      details: error.message,
      auth_expired: true
    });
    return true;
  }
  return false;
};

console.log("GOOGLE_CLIENT_ID exists:", !!process.env.GOOGLE_CLIENT_ID);
console.log("GOOGLE_CLIENT_SECRET exists:", !!process.env.GOOGLE_CLIENT_SECRET);
console.log("APP_URL:", process.env.APP_URL);

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(cookieParser());

// Registered here, before any route, because Express middleware only guards routes
// declared after it. requireAuth itself is a hoisted function declaration.
app.use(requireAuth);
app.use(scopeWorkshopId);
app.use(auditLog);

// Debug routes early
app.get("/api/debug/auth", async (req, res) => {
  const storedTokens = await loadTokens();
  res.json({
    database: "supabase",
    database_url_set: !!process.env.DATABASE_URL,
    stored_tokens_exist: !!storedTokens,
    cookie_tokens_exist: !!req.cookies.google_tokens,
    last_error: lastDbError,
    env_vars: {
      GOOGLE_CLIENT_ID: !!process.env.GOOGLE_CLIENT_ID,
      GOOGLE_CLIENT_SECRET: !!process.env.GOOGLE_CLIENT_SECRET,
      APP_URL: process.env.APP_URL
    }
  });
});

app.get("/api/debug/test-db", async (req, res) => {
  try {
    const ok = await writeConfig("test_connection", { time: new Date().toISOString(), status: "ok" });
    if (!ok) throw new Error(lastDbError || "unknown");
    const counts = await sqlPool.query(
      `SELECT (SELECT count(*) FROM signatures)::int AS signatures,
              (SELECT count(*) FROM employees)::int AS employees,
              (SELECT count(*) FROM leave_requests)::int AS leave_requests,
              (SELECT count(*) FROM app_config)::int AS app_config`
    );
    res.json({ success: true, message: "Ghi dữ liệu thành công vào Supabase!", counts: counts.rows[0] });
  } catch (e: any) {
    res.status(500).json({ success: false, error: e.message });
  }
});

const getOAuth2Client = () => {
  const appUrl = (process.env.APP_URL || 'http://localhost:3000').trim().replace(/\/$/, "");
  const clientId = process.env.GOOGLE_CLIENT_ID?.trim();
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET?.trim();
  
  if (!clientId || !clientSecret) {
    throw new Error("Missing GOOGLE_CLIENT_ID or GOOGLE_CLIENT_SECRET environment variables");
  }

  const client = new google.auth.OAuth2(
    clientId,
    clientSecret,
    `${appUrl}/api/auth/google/callback`
  );

  client.on('tokens', (tokens) => {
    saveTokens(tokens).catch(e => console.error("Error auto-saving refreshed tokens:", e));
  });

  return client;
};

function removeVietnameseAccents(str: string): string {
  return str
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

async function fetchLeaveBalanceHelper(name: string, workshopId: string, year?: string): Promise<{ entitled: string; used: string; remaining: string } | null> {
  try {
    const quotas = await getQuotaByYear(name, workshopId);
    const targetYear = year || String(new Date().getFullYear());
    const q = quotas[targetYear];
    if (!q) return null;

    return {
      entitled: String(q.entitled + q.travelDays),
      used: String(q.used),
      remaining: String(q.remaining)
    };
  } catch (err: any) {
    console.error("fetchLeaveBalanceHelper (SQL) error:", err.message);
  }
  return null;
}

// The rota (base date + shift matrix) is configured per workshop, stored on the workshop's
// own row; falls back to the legacy global app_settings for requests with no workshop yet.
async function loadShiftConfig(workshopId?: string): Promise<ShiftConfig | undefined> {
  try {
    if (workshopId) {
      const ws = await sqlPool.query(`SELECT config FROM workshops WHERE id = $1`, [workshopId]);
      const schedule = ws.rows[0]?.config?.shiftSchedule;
      if (schedule) return { baseDate: schedule.baseDate, shiftsMatrix: schedule.shiftsMatrix };
    }
    const doc = await readConfig("app_settings");
    const schedule = doc?.config?.shiftSchedule;
    if (!schedule) return undefined;
    return { baseDate: schedule.baseDate, shiftsMatrix: schedule.shiftsMatrix };
  } catch (e: any) {
    console.log("Could not load shift config, using defaults:", e.message);
    return undefined;
  }
}

const toNum = (v: any): number => {
  const n = parseFloat(String(v ?? "").replace(",", "."));
  return isNaN(n) ? 0 : n;
};

export interface YearQuota {
  year: string;
  entitled: number;
  travelDays: number;
  used: number;
  remaining: number;
  seniority?: number;
  source?: "seniority" | "manual";
}

// Build the per-year leave account for one person: allowance + travel days earned, minus days taken.
// Scoped to a single workshop so two people who happen to share a name in different
// workshops never share a leave balance.
async function getQuotaByYear(name: string, workshopId: string, excludeLeaveId?: string, extraYears: string[] = []): Promise<Record<string, YearQuota>> {
  // Entitlement normally comes from the hire year, recomputed for whichever year is asked about.
  const employee = await sqlPool.query(
    `SELECT hire_year, base_days FROM employees WHERE workshop_id = $1 AND lower(trim(name)) = lower(trim($2))`,
    [workshopId, name]
  );
  const emp = employee.rows[0];

  const quotas: Record<string, YearQuota> = {};
  const ensure = (year: string): YearQuota => {
    if (!quotas[year]) {
      const q: YearQuota = { year, entitled: 0, travelDays: 0, used: 0, remaining: 0 };
      if (emp) {
        q.entitled = calcEntitledDays(Number(emp.base_days), Number(emp.hire_year), Number(year));
        q.seniority = calcSeniority(Number(emp.hire_year), Number(year));
        q.source = "seniority";
      }
      quotas[year] = q;
    }
    return quotas[year];
  };

  if (emp) {
    const thisYear = new Date().getFullYear();
    for (let y = Math.max(Number(emp.hire_year), thisYear - 2); y <= thisYear + 1; y++) ensure(String(y));
  }
  for (const y of extraYears) if (y) ensure(String(y));

  // A row in leave_balances overrides the computed figures for that year. The two
  // overrides are separate: a row written to record days taken outside the system
  // leaves entitled NULL, and must not be read as "entitlement is 0, set by hand".
  const balances = await sqlPool.query(
    `SELECT year, entitled, used_adjust FROM leave_balances
     WHERE workshop_id = $1 AND lower(trim(name)) = lower(trim($2))`,
    [workshopId, name]
  );
  const usedAdjustByYear = new Map<string, number>();
  for (const row of balances.rows) {
    const q = ensure(String(row.year));
    if (row.entitled !== null && row.entitled !== undefined) {
      q.entitled = toNum(row.entitled);
      q.source = "manual";
    }
    if (row.used_adjust !== null && row.used_adjust !== undefined) {
      usedAdjustByYear.set(String(row.year), toNum(row.used_adjust));
    }
  }

  const travel = await sqlPool.query(
    `SELECT leave_year, SUM(travel_days) AS days FROM leave_requests
     WHERE workshop_id = $1 AND lower(trim(name)) = lower(trim($2)) AND travel_days > 0
       AND ($3::text IS NULL OR id <> $3)
     GROUP BY leave_year`,
    [workshopId, name, excludeLeaveId || null]
  );
  for (const row of travel.rows) ensure(String(row.leave_year)).travelDays = toNum(row.days);

  const used = await sqlPool.query(
    `SELECT a.leave_year, SUM(a.days) AS days FROM leave_allocations a
     JOIN leave_requests r ON r.id = a.leave_id
     WHERE r.workshop_id = $1 AND lower(trim(r.name)) = lower(trim($2))
       AND ($3::text IS NULL OR r.id <> $3)
     GROUP BY a.leave_year`,
    [workshopId, name, excludeLeaveId || null]
  );
  for (const row of used.rows) ensure(String(row.leave_year)).used = toNum(row.days);

  // Added on top of the summed requests, not substituted for them. Years whose leave
  // was all taken before this system existed have an adjustment and no requests at
  // all, so this runs over the adjustments rather than over the rows above.
  for (const [year, adjust] of usedAdjustByYear) ensure(year).used += adjust;

  for (const q of Object.values(quotas)) q.remaining = q.entitled + q.travelDays - q.used;
  return quotas;
}

// Look up the road distance from Pleiku to a location, matching with or without accents.
async function lookupDistanceKm(location: string): Promise<number | null> {
  const target = String(location || "").trim();
  if (!target) return null;

  try {
    const result = await sqlPool.query(`SELECT name, distance_km FROM location_distances`);
    const targetNormalized = removeVietnameseAccents(target);

    let matched = result.rows.find((r: any) => String(r.name || '').trim().toLowerCase() === target.toLowerCase());
    if (!matched) {
      matched = result.rows.find((r: any) => removeVietnameseAccents(String(r.name || '').trim()) === targetNormalized);
    }
    // Fall back to a contains match so "Huyện X, Nghệ An" still resolves to "Nghệ An"
    if (!matched) {
      matched = result.rows.find((r: any) => targetNormalized.includes(removeVietnameseAccents(String(r.name || '').trim())));
    }

    return matched ? Number(matched.distance_km) : null;
  } catch (err: any) {
    console.error("lookupDistanceKm error:", err.message);
    return null;
  }
}

// The regulation grants travel days once per leave year, so an earlier grant blocks a new one.
async function findExistingTravelDays(name: string, workshopId: string, leaveYear: string, excludeId?: string) {
  const result = await sqlPool.query(
    `SELECT id, travel_days, start_date FROM leave_requests
     WHERE workshop_id = $1 AND lower(trim(name)) = lower(trim($2)) AND leave_year = $3 AND travel_days > 0
       AND ($4::text IS NULL OR id <> $4)
     ORDER BY inserted_at ASC LIMIT 1`,
    [workshopId, name, leaveYear, excludeId || null]
  );
  return result.rows[0] || null;
}

// Resolve how many travel days a request qualifies for, and explain the outcome.
async function resolveTravelDays(opts: {
  name: string;
  workshopId: string;
  location: string;
  leaveYear: string;
  hasLeavePermit: boolean;
  excludeId?: string;
}) {
  const { name, workshopId, location, leaveYear, hasLeavePermit, excludeId } = opts;

  if (!hasLeavePermit) {
    return { travelDays: 0, distanceKm: null, note: "Đơn không lấy giấy phép nên không tính ngày đi đường." };
  }

  const distanceKm = await lookupDistanceKm(location);
  if (distanceKm === null) {
    return { travelDays: 0, distanceKm: null, note: `Chưa có số km từ Pleiku đến "${location}" trong bảng khoảng cách.` };
  }

  const eligibleDays = calcTravelDays(distanceKm);
  if (eligibleDays === 0) {
    return { travelDays: 0, distanceKm, note: `Quãng đường ${distanceKm} km (dưới 200 km) nên không được tính ngày đi đường.` };
  }

  const existing = await findExistingTravelDays(name, workshopId, leaveYear, excludeId);
  if (existing) {
    return {
      travelDays: 0,
      distanceKm,
      note: `Đồng chí ${name} đã được tính ${existing.travel_days} ngày đi đường trong năm ${leaveYear} (đơn nghỉ từ ${existing.start_date}). Ngày đi đường chỉ tính một lần trong năm.`
    };
  }

  return { travelDays: eligibleDays, distanceKm, note: `Quãng đường ${distanceKm} km, được cộng thêm ${eligibleDays} ngày đi đường.` };
}

// Work out everything a leave request implies: how many shifts it costs, which leave
// year(s) those shifts are charged to, and whether travel days are granted.
async function planLeaveRequest(opts: {
  name: string;
  workshopId: string;
  kip: string | number;
  startDate: string;
  endDate: string;
  location: string;
  hasLeavePermit: boolean;
  excludeLeaveId?: string;
}) {
  const { name, workshopId, kip, startDate, endDate, location, hasLeavePermit, excludeLeaveId } = opts;

  const shiftConfig = await loadShiftConfig(workshopId);
  const shifts = listWorkingShifts(startDate, endDate, Number(kip), shiftConfig);

  // The leave may fall outside the default window, and each shift can also reach back a year.
  const touchedYears = new Set<string>();
  for (const s of shifts) {
    touchedYears.add(String(s.date.getFullYear()));
    touchedYears.add(String(s.date.getFullYear() - 1));
  }
  const quotas = await getQuotaByYear(name, workshopId, excludeLeaveId, [...touchedYears]);

  const remainingByYear: Record<string, number> = {};
  for (const [year, q] of Object.entries(quotas)) remainingByYear[year] = q.remaining;

  if (shifts.length === 0) {
    return {
      leaveDays: 0,
      leaveYear: String(new Date(startDate + "T00:00:00").getFullYear() || new Date().getFullYear()),
      allocations: {} as Record<string, number>,
      detail: [] as { date: string; shift: string; year: string }[],
      travel: { travelDays: 0, distanceKm: null as number | null, note: "Đơn không trùng ca trực nào nên không trừ ngày phép." },
      quotas
    };
  }

  // The primary leave year is where the first shift lands; travel days attach to it.
  const firstProbe = allocateLeaveDays([shifts[0]], remainingByYear);
  const leaveYear = firstProbe.detail[0].year;

  const travel = await resolveTravelDays({ name, workshopId, location, leaveYear, hasLeavePermit, excludeId: excludeLeaveId });

  // Travel days enlarge that year's allowance before the days are charged against it.
  if (travel.travelDays > 0) {
    remainingByYear[leaveYear] = (remainingByYear[leaveYear] ?? 0) + travel.travelDays;
  }

  const { allocations, detail } = allocateLeaveDays(shifts, remainingByYear);

  return { leaveDays: shifts.length, leaveYear, allocations, detail, travel, quotas };
}

app.get("/api/leave/plan", async (req, res) => {
  const { name, kip, startDate, endDate, location, hasLeavePermit, workshopId } = req.query;
  if (!startDate || !endDate || !kip) {
    return res.status(400).json({ error: "Thiếu ngày bắt đầu, ngày kết thúc hoặc kíp." });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }

  try {
    const plan = await planLeaveRequest({
      name: String(name || ""),
      workshopId: String(workshopId),
      kip: String(kip),
      startDate: String(startDate),
      endDate: String(endDate),
      location: String(location || ""),
      hasLeavePermit: String(hasLeavePermit) === "true"
    });
    res.json({ success: true, ...plan });
  } catch (error: any) {
    console.error("Error planning leave request:", error);
    res.status(500).json({ error: "Không thể tính ngày phép", details: error.message });
  }
});

app.get("/api/leave/employees", async (req, res) => {
  const year = Number(req.query.year) || new Date().getFullYear();
  const workshopId = req.query.workshopId;
  if (!workshopId) return res.status(400).json({ error: "Thiếu phân xưởng." });
  try {
    const result = await sqlPool.query(
      `SELECT name, hire_year, base_days FROM employees WHERE workshop_id = $1 ORDER BY name ASC`,
      [workshopId]
    );
    res.json(result.rows.map((r: any) => ({
      name: r.name,
      hireYear: Number(r.hire_year),
      baseDays: Number(r.base_days),
      seniority: calcSeniority(Number(r.hire_year), year),
      entitled: calcEntitledDays(Number(r.base_days), Number(r.hire_year), year),
      year: String(year)
    })));
  } catch (error: any) {
    console.error("Error fetching employees:", error);
    res.status(500).json({ error: "Không thể tải danh sách nhân viên", details: error.message });
  }
});

app.post("/api/leave/employees/import", express.json({ limit: "5mb" }), async (req, res) => {
  const { rows, replaceAll, workshopId, year } = req.body;
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }
  if (!Array.isArray(rows) || rows.length === 0) {
    return res.status(400).json({ error: "Không có dòng dữ liệu nào để nhập." });
  }
  const importYear = /^\d{4}$/.test(String(year)) ? String(year) : String(new Date().getFullYear());

  const valid: { name: string; hireYear: number; baseDays: number; used?: number }[] = [];
  const skipped: string[] = [];

  for (const row of rows) {
    const name = String(row?.name ?? "").trim();
    const hireYear = Number(row?.hireYear);
    const baseDays = Number(row?.baseDays);
    if (!name || !Number.isInteger(hireYear) || hireYear < 1950 || hireYear > 2100) {
      skipped.push(`${name || "(không tên)"}: năm vào làm không hợp lệ`);
      continue;
    }
    if (!Number.isFinite(baseDays) || baseDays <= 0 || baseDays > 40) {
      skipped.push(`${name}: mức phép cơ bản không hợp lệ`);
      continue;
    }

    // Absent means "leave the days-used figure alone"; a 0 in the sheet is a real
    // instruction to set it to zero. The two must not collapse into each other, or a
    // file exported from the older template would silently wipe recorded days.
    let used: number | undefined;
    if (row?.used !== undefined && row?.used !== null && String(row.used).trim() !== "") {
      const n = Number(row.used);
      if (!Number.isFinite(n) || n < 0 || n > 366) {
        skipped.push(`${name}: số ngày đã nghỉ không hợp lệ`);
        continue;
      }
      used = n;
    }

    valid.push({ name, hireYear, baseDays: Math.round(baseDays), used });
  }

  if (valid.length === 0) {
    return res.status(400).json({ error: "Không có dòng hợp lệ nào.", skipped });
  }

  const client = await sqlPool.connect();
  try {
    await client.query("BEGIN");
    if (replaceAll === true) await client.query(`DELETE FROM employees WHERE workshop_id = $1`, [workshopId]);

    for (const emp of valid) {
      await client.query(
        `INSERT INTO employees (name, hire_year, base_days, workshop_id) VALUES ($1,$2,$3,$4)
         ON CONFLICT (workshop_id, name) DO UPDATE SET hire_year = EXCLUDED.hire_year, base_days = EXCLUDED.base_days, updated_at = now()`,
        [emp.name, emp.hireYear, emp.baseDays, workshopId]
      );
    }

    // Same rule as editing the cell by hand: the sheet gives the total to display, and
    // what gets stored is the part not already covered by saved requests. One grouped
    // query rather than a lookup per person, so a 78-row sheet stays a single round trip.
    const withUsed = valid.filter(v => v.used !== undefined);
    if (withUsed.length > 0) {
      const agg = await client.query(
        `SELECT lower(trim(r.name)) AS key, SUM(a.days) AS days FROM leave_allocations a
         JOIN leave_requests r ON r.id = a.leave_id
         WHERE r.workshop_id = $1 AND a.leave_year = $2 GROUP BY lower(trim(r.name))`,
        [workshopId, importYear]
      );
      const computedByName = new Map(agg.rows.map((r: any) => [r.key, toNum(r.days)]));

      for (const emp of withUsed) {
        const computed = computedByName.get(emp.name.trim().toLowerCase()) || 0;
        await client.query(
          `INSERT INTO leave_balances (name, year, used_adjust, workshop_id) VALUES ($1,$2,$3,$4)
           ON CONFLICT (workshop_id, name, year) DO UPDATE SET used_adjust = EXCLUDED.used_adjust, updated_at = now()`,
          [emp.name, importYear, String((emp.used as number) - computed), workshopId]
        );
      }
    }
    await client.query("COMMIT");

    const usedNote = withUsed.length > 0 ? ` Đã ghi số ngày đã nghỉ năm ${importYear} cho ${withUsed.length} người.` : "";
    res.json({
      success: true,
      imported: valid.length,
      usedImported: withUsed.length,
      skipped,
      message: `✅ Đã nhập ${valid.length} nhân viên từ Excel.${usedNote}${skipped.length > 0 ? ` Bỏ qua ${skipped.length} dòng lỗi.` : ""}`
    });
  } catch (error: any) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Error importing employees:", error);
    res.status(500).json({ error: "Nhập dữ liệu thất bại", details: error.message });
  } finally {
    client.release();
  }
});

app.post("/api/leave/employees/delete", async (req, res) => {
  const { name, workshopId } = req.body;
  if (!name) return res.status(400).json({ error: "Thiếu tên nhân viên." });
  if (!workshopId) return res.status(400).json({ error: "Thiếu phân xưởng." });
  try {
    const result = await sqlPool.query(
      `DELETE FROM employees WHERE workshop_id = $1 AND lower(trim(name)) = lower(trim($2))`,
      [workshopId, name]
    );
    if (result.rowCount === 0) return res.status(404).json({ error: "Không tìm thấy nhân viên." });
    res.json({ success: true, deletedCount: result.rowCount });
  } catch (error: any) {
    console.error("Error deleting employee:", error);
    res.status(500).json({ error: "Xóa thất bại", details: error.message });
  }
});

app.get("/api/leave/balances", async (req, res) => {
  const year = String(req.query.year || new Date().getFullYear());
  const yearNum = Number(year);
  const workshopId = req.query.workshopId;
  if (!workshopId) return res.status(400).json({ error: "Thiếu phân xưởng." });

  try {
    // Four aggregate queries instead of a per-person round trip, so 78 staff stay fast.
    const [employees, overrides, travel, used] = await Promise.all([
      sqlPool.query(`SELECT name, hire_year, base_days FROM employees WHERE workshop_id = $1`, [workshopId]),
      sqlPool.query(`SELECT name, entitled, used_adjust FROM leave_balances WHERE workshop_id = $1 AND year = $2`, [workshopId, year]),
      sqlPool.query(
        `SELECT lower(trim(name)) AS key, SUM(travel_days) AS days FROM leave_requests
         WHERE workshop_id = $1 AND leave_year = $2 AND travel_days > 0 GROUP BY lower(trim(name))`,
        [workshopId, year]
      ),
      sqlPool.query(
        `SELECT lower(trim(r.name)) AS key, SUM(a.days) AS days FROM leave_allocations a
         JOIN leave_requests r ON r.id = a.leave_id
         WHERE r.workshop_id = $1 AND a.leave_year = $2 GROUP BY lower(trim(r.name))`,
        [workshopId, year]
      )
    ]);

    const key = (n: string) => String(n || "").trim().toLowerCase();
    const travelByName = new Map(travel.rows.map((r: any) => [r.key, toNum(r.days)]));
    const usedByName = new Map(used.rows.map((r: any) => [r.key, toNum(r.days)]));

    const rows = new Map<string, any>();

    for (const emp of employees.rows) {
      rows.set(key(emp.name), {
        name: emp.name,
        year,
        entitled: calcEntitledDays(Number(emp.base_days), Number(emp.hire_year), yearNum),
        seniority: calcSeniority(Number(emp.hire_year), yearNum),
        source: "seniority" as const
      });
    }

    // A row here can carry either override independently. One written only to record
    // days taken outside the system leaves entitled NULL, which must not be read as
    // an entitlement of 0 set by hand — that would wipe out the seniority figure.
    const usedAdjustByName = new Map<string, number>();
    for (const ov of overrides.rows) {
      const k = key(ov.name);
      const existing = rows.get(k);
      if (ov.used_adjust !== null && ov.used_adjust !== undefined) {
        usedAdjustByName.set(k, toNum(ov.used_adjust));
      }
      if (ov.entitled === null || ov.entitled === undefined) {
        // Keep the person visible even if they have no employees row to derive from.
        if (!existing) {
          rows.set(k, { name: ov.name, year, entitled: 0, seniority: null, source: "seniority" as const });
        }
        continue;
      }
      rows.set(k, {
        name: existing?.name || ov.name,
        year,
        entitled: toNum(ov.entitled),
        seniority: existing?.seniority ?? null,
        source: "manual" as const
      });
    }

    const out = [...rows.entries()].map(([k, row]) => {
      const travelDays = travelByName.get(k) || 0;
      const usedAdjust = usedAdjustByName.get(k) || 0;
      const usedDays = (usedByName.get(k) || 0) + usedAdjust;
      return { ...row, travelDays, used: usedDays, usedAdjust, remaining: row.entitled + travelDays - usedDays };
    });
    out.sort((a, b) => a.name.localeCompare(b.name, "vi"));

    res.json(out);
  } catch (error: any) {
    console.error("Error fetching leave balances:", error);
    res.status(500).json({ error: "Không thể tải bảng phép năm", details: error.message });
  }
});

app.post("/api/leave/balances", async (req, res) => {
  const { name, year, entitled, workshopId } = req.body;
  if (!name || !year) {
    return res.status(400).json({ error: "Thiếu tên nhân viên hoặc năm." });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }

  try {
    await sqlPool.query(
      `INSERT INTO leave_balances (name, year, entitled, workshop_id) VALUES ($1,$2,$3,$4)
       ON CONFLICT (workshop_id, name, year) DO UPDATE SET entitled = EXCLUDED.entitled, updated_at = now()`,
      [String(name).trim(), String(year), String(toNum(entitled)), workshopId]
    );
    res.json({ success: true, message: `✅ Đã lưu ${toNum(entitled)} ngày phép năm ${year} cho đồng chí ${name}.` });
  } catch (error: any) {
    console.error("Error saving leave balance:", error);
    res.status(500).json({ error: "Lưu phép năm thất bại", details: error.message });
  }
});

// Correct the days-used figure. The caller sends the total it wants to see; what gets
// stored is the difference against the days already summed from saved requests, so the
// figure keeps tracking requests entered later instead of freezing at the typed number.
// Storing the total verbatim would look right today and drift silently from then on.
app.post("/api/leave/balances/used", async (req, res) => {
  const { name, year, used, workshopId } = req.body;
  if (!name || !year) {
    return res.status(400).json({ error: "Thiếu tên nhân viên hoặc năm." });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }
  const desired = toNum(used);
  if (desired < 0) {
    return res.status(400).json({ error: "Số ngày đã nghỉ không được là số âm." });
  }

  try {
    const fromRequests = await sqlPool.query(
      `SELECT COALESCE(SUM(a.days), 0) AS days FROM leave_allocations a
       JOIN leave_requests r ON r.id = a.leave_id
       WHERE r.workshop_id = $1 AND lower(trim(r.name)) = lower(trim($2)) AND a.leave_year = $3`,
      [workshopId, name, String(year)]
    );
    const computed = toNum(fromRequests.rows[0]?.days);
    const adjust = desired - computed;

    // Only used_adjust is written. Touching entitled here would flip the row to a
    // manual entitlement override and stop it growing with seniority.
    await sqlPool.query(
      `INSERT INTO leave_balances (name, year, used_adjust, workshop_id) VALUES ($1,$2,$3,$4)
       ON CONFLICT (workshop_id, name, year) DO UPDATE SET used_adjust = EXCLUDED.used_adjust, updated_at = now()`,
      [String(name).trim(), String(year), String(adjust), workshopId]
    );

    const note = computed > 0
      ? ` (${computed} ngày từ đơn đã lưu + ${adjust} ngày ghi nhận thủ công)`
      : adjust > 0 ? ` (ghi nhận thủ công, chưa có đơn nào trong hệ thống)` : "";
    res.json({
      success: true,
      computed,
      adjust,
      message: `✅ Đã đặt số ngày đã nghỉ năm ${year} của đồng chí ${name} thành ${desired}${note}.`
    });
  } catch (error: any) {
    console.error("Error saving used days:", error);
    res.status(500).json({ error: "Lưu số ngày đã nghỉ thất bại", details: error.message });
  }
});

app.post("/api/leave/balances/delete", async (req, res) => {
  const { name, year, workshopId } = req.body;
  if (!name || !year) {
    return res.status(400).json({ error: "Thiếu tên nhân viên hoặc năm." });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }

  try {
    const result = await sqlPool.query(
      `DELETE FROM leave_balances WHERE workshop_id = $1 AND lower(trim(name)) = lower(trim($2)) AND year = $3`,
      [workshopId, name, String(year)]
    );
    if (result.rowCount === 0) return res.status(404).json({ error: "Không tìm thấy dòng cần xóa." });
    res.json({ success: true, deletedCount: result.rowCount });
  } catch (error: any) {
    console.error("Error deleting leave balance:", error);
    res.status(500).json({ error: "Xóa thất bại", details: error.message });
  }
});

app.get("/api/leave/locations", async (req, res) => {
  try {
    const result = await sqlPool.query(`SELECT name, distance_km FROM location_distances ORDER BY distance_km ASC`);
    res.json(result.rows.map((r: any) => ({ name: r.name, distanceKm: Number(r.distance_km) })));
  } catch (error: any) {
    console.error("Error fetching location distances:", error);
    res.status(500).json({ error: "Không thể tải bảng khoảng cách", details: error.message });
  }
});

app.get("/api/leave/travel-days", async (req, res) => {
  const { name, location, leaveYear, hasLeavePermit, workshopId } = req.query;
  if (!workshopId) return res.status(400).json({ error: "Thiếu phân xưởng." });
  try {
    const result = await resolveTravelDays({
      name: String(name || ""),
      workshopId: String(workshopId),
      location: String(location || ""),
      leaveYear: String(leaveYear || new Date().getFullYear()),
      hasLeavePermit: String(hasLeavePermit) === "true"
    });
    res.json({ success: true, ...result });
  } catch (error: any) {
    console.error("Error resolving travel days:", error);
    res.status(500).json({ error: "Không thể tính ngày đi đường", details: error.message });
  }
});

// Signature Management Endpoints
app.get("/api/signatures", async (req, res) => {
  const workshopId = req.query.workshopId;
  if (!workshopId) return res.status(400).json({ error: "Thiếu phân xưởng." });
  try {
    const result = await sqlPool.query(`SELECT name, data FROM signatures WHERE workshop_id = $1`, [workshopId]);
    const signatures: Record<string, string> = {};
    for (const row of result.rows) signatures[row.name] = row.data;
    res.json(signatures);
  } catch (e: any) {
    console.error("Error fetching signatures:", e);
    res.status(500).json({ error: e.message });
  }
});

app.post("/api/signatures", express.json({ limit: '10mb' }), async (req, res) => {
  const { name, data, workshopId } = req.body;
  if (!name) {
    return res.status(400).json({ error: "Name is required" });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }
  try {
    // An empty data payload is how the client asks for the signature to be removed.
    if (!data) {
      await sqlPool.query(`DELETE FROM signatures WHERE workshop_id = $1 AND name = $2`, [workshopId, name]);
      return res.json({ success: true, deleted: true });
    }
    await sqlPool.query(
      `INSERT INTO signatures (name, data, workshop_id) VALUES ($1,$2,$3)
       ON CONFLICT (workshop_id, name) DO UPDATE SET data = EXCLUDED.data, updated_at = now()`,
      [name, data, workshopId]
    );
    res.json({ success: true });
  } catch (e: any) {
    console.error("Error saving signature:", e);
    res.status(500).json({ error: e.message });
  }
});

// Batch upload used by the "match files to staff names" flow in SignatureManager.
app.post("/api/signatures/batch", express.json({ limit: '50mb' }), async (req, res) => {
  const { items, workshopId } = req.body;
  if (!Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: "items is required" });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }

  const client = await sqlPool.connect();
  try {
    await client.query("BEGIN");
    let saved = 0;
    for (const item of items) {
      if (!item?.name || !item?.data) continue;
      await client.query(
        `INSERT INTO signatures (name, data, workshop_id) VALUES ($1,$2,$3)
         ON CONFLICT (workshop_id, name) DO UPDATE SET data = EXCLUDED.data, updated_at = now()`,
        [item.name, item.data, workshopId]
      );
      saved++;
    }
    await client.query("COMMIT");
    res.json({ success: true, saved });
  } catch (e: any) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Error saving signatures batch:", e);
    res.status(500).json({ error: e.message });
  } finally {
    client.release();
  }
});

// ============================================================================
// Multi-workshop login/admin system (LoginForm, UserHeaderBar, WorkshopManagerModal).
// Every /api route is guarded by requireAuth below. The caller's identity comes from
// an httpOnly session cookie that the browser cannot read or forge from JavaScript —
// never from the request body, and never from what the client claims its role is.
// ============================================================================

const SESSION_COOKIE = "sid";
const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days

function hashSessionToken(token: string): string {
  return crypto.createHash("sha256").update(token).digest("hex");
}

// Returns the raw token to hand to the browser; only its hash reaches the database,
// so read access to user_sessions cannot be replayed as a login.
async function createSession(userId: string): Promise<string> {
  const token = crypto.randomBytes(32).toString("hex");
  const expiresAt = new Date(Date.now() + SESSION_TTL_MS);
  await sqlPool.query(
    `INSERT INTO user_sessions (token_hash, user_id, expires_at) VALUES ($1, $2, $3)`,
    [hashSessionToken(token), userId, expiresAt]
  );
  return token;
}

async function loadSession(token: string | undefined) {
  if (!token) return null;
  const result = await sqlPool.query(
    `SELECT u.* FROM user_sessions s
       JOIN user_accounts u ON u.id = s.user_id
      WHERE s.token_hash = $1 AND s.expires_at > now()`,
    [hashSessionToken(token)]
  );
  return result.rows[0] || null;
}

async function destroySession(token: string | undefined) {
  if (!token) return;
  await sqlPool.query(`DELETE FROM user_sessions WHERE token_hash = $1`, [hashSessionToken(token)]);
}

function setSessionCookie(req: any, res: any, token: string) {
  // Deriving this from the request rather than NODE_ENV means a deploy that forgets to
  // set the env var still gets a Secure cookie whenever the connection is HTTPS.
  const isHttps = req.secure
    || String(req.headers["x-forwarded-proto"] || "").split(",")[0].trim() === "https"
    || process.env.NODE_ENV === "production";
  res.cookie(SESSION_COOKIE, token, {
    httpOnly: true,
    sameSite: "lax",
    secure: isHttps,
    maxAge: SESSION_TTL_MS,
    path: "/"
  });
}

// True only while the system has no accounts at all. It lets the very first admin be
// created through the login screen and then closes permanently — without it there is
// no way to bootstrap, and with it left open anyone could mint themselves an account.
async function needsBootstrap(): Promise<boolean> {
  const r = await sqlPool.query(`SELECT 1 FROM user_accounts LIMIT 1`);
  return r.rowCount === 0;
}

// Reachable without a session. Everything else under /api requires one.
const PUBLIC_API_PATHS = new Set([
  "/api/auth/login",
  "/api/auth/logout",
  "/api/auth/me",
  "/api/auth/register",
  "/api/setup/status"
]);

// Only reachable without a session while needsBootstrap() is true — the escape hatch
// for creating the very first super admin on an empty database.
const BOOTSTRAP_API_PATHS: Array<{ method: string; path: string }> = [
  { method: "GET", path: "/api/workshops" },
  { method: "POST", path: "/api/workshops" },
  { method: "POST", path: "/api/accounts" }
];

async function requireAuth(req: any, res: any, next: any) {
  const path = req.path;
  // Mounted globally rather than on "/api" because Express strips the mount prefix
  // from req.path, which would make every check below silently miss.
  if (!path.startsWith("/api/")) return next();
  if (PUBLIC_API_PATHS.has(path)) return next();

  try {
    const user = await loadSession(req.cookies?.[SESSION_COOKIE]);
    if (user) {
      req.user = toUserAccount(user);
      return next();
    }

    if (BOOTSTRAP_API_PATHS.some(r => r.method === req.method && r.path === path) && await needsBootstrap()) {
      return next();
    }

    // The Google OAuth popup lands here after a cross-site redirect. Answering with
    // JSON would leave a blank window, so send readable HTML instead.
    if (path.startsWith("/api/auth/google")) {
      return res.status(401).send("<html><body style=\"font-family:sans-serif;padding:24px\">"
        + "<h3>Phiên đăng nhập đã hết hạn</h3>"
        + "<p>Vui lòng đóng cửa sổ này, đăng nhập lại rồi thử lại.</p></body></html>");
    }
    return res.status(401).json({ error: "Chưa đăng nhập hoặc phiên đã hết hạn." });
  } catch (e: any) {
    console.error("Auth check failed:", e);
    return res.status(500).json({ error: "Lỗi kiểm tra phiên đăng nhập." });
  }
}

// Records who did what. Reads are skipped — logging every GET would bury the writes
// that actually matter. The actor comes from the session, so it cannot be spoofed, and
// the row is written after the response so a failed request is logged as failed.
const AUDIT_SKIP_PATHS = new Set(["/api/auth/login", "/api/auth/logout", "/api/auth/me"]);

function auditSummary(req: any): string {
  const b = req.body || {};
  const bits: string[] = [];
  if (b.name) bits.push(`name=${b.name}`);
  if (b.username) bits.push(`username=${b.username}`);
  if (b.role) bits.push(`role=${b.role}`);
  if (req.params?.id) bits.push(`id=${req.params.id}`);
  if (b.id) bits.push(`id=${b.id}`);
  if (Array.isArray(b.rows)) bits.push(`rows=${b.rows.length}`);
  if (Array.isArray(b.ids)) bits.push(`ids=${b.ids.length}`);
  if (Array.isArray(b.staffData)) bits.push(`staffRows=${b.staffData.length}`);
  if (b.replaceAll === true) bits.push("replaceAll=true");
  return bits.join(" ").slice(0, 500);
}

function auditLog(req: any, res: any, next: any) {
  if (!req.path.startsWith("/api/")) return next();
  if (req.method === "GET" || req.method === "HEAD") return next();
  if (AUDIT_SKIP_PATHS.has(req.path)) return next();

  const summary = auditSummary(req);
  res.on("finish", () => {
    sqlPool.query(
      `INSERT INTO audit_log (user_id, username, role, workshop_id, method, path, status, summary, ip)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)`,
      [
        req.user?.id || null,
        req.user?.username || null,
        req.user?.role || null,
        req.user?.workshopId || req.body?.workshopId || null,
        req.method,
        req.path,
        res.statusCode,
        summary,
        req.socket?.remoteAddress || null
      ]
    ).catch((e: any) => console.error("Audit log write failed:", e.message));
  });
  next();
}

// Every workshop-scoped query filters on a workshopId that arrives in the query string
// or body. Left alone, any logged-in user could simply name another workshop and read
// or delete its data. Rather than patch ~15 handlers individually (and miss one), pin
// the value here, once, for everybody who is not a super admin.
function scopeWorkshopId(req: any, res: any, next: any) {
  if (!req.path.startsWith("/api/")) return next();
  if (!req.user || isSuperAdmin(req)) return next(); // super admins legitimately switch

  // Past this point req.user is guaranteed NOT a super admin. A non-super-admin should
  // always carry a real workshop id — "all" is the value toUserAccount() assigns when
  // workshop_id is NULL, which is otherwise only meant for super_admin rows. If a
  // workshop_admin ever ends up with an empty or "all" workshopId (e.g. leftover data
  // from before workshop deletion cascaded to accounts), deny rather than treat it as
  // unrestricted access — the previous `return next()` here let such an account read
  // and modify every other workshop's data.
  const own = String(req.user.workshopId || "");
  if (!own || own === "all") {
    return res.status(403).json({ error: "Tài khoản chưa được gán phân xưởng hợp lệ." });
  }

  const asked = req.query?.workshopId ?? req.body?.workshopId;
  if (asked !== undefined && asked !== null && String(asked) !== "" && String(asked) !== own) {
    return res.status(403).json({ error: "Bạn chỉ truy cập được dữ liệu phân xưởng của mình." });
  }

  // Overwrite rather than merely validate, so a handler reading either location can
  // never end up with a value the caller chose.
  if (req.query && typeof req.query === "object") req.query.workshopId = own;
  if (req.body && typeof req.body === "object" && !Array.isArray(req.body)) req.body.workshopId = own;
  next();
}

app.get("/api/audit-log", async (req: any, res) => {
  try {
    const limit = Math.min(Number(req.query.limit) || 200, 1000);
    if (isSuperAdmin(req)) {
      const r = await sqlPool.query(`SELECT * FROM audit_log ORDER BY at DESC LIMIT $1`, [limit]);
      return res.json(r.rows);
    }
    const wsId = req.user?.workshopId;
    if (!wsId || wsId === "all") return res.json([]);
    const r = await sqlPool.query(
      `SELECT * FROM audit_log WHERE workshop_id = $1 ORDER BY at DESC LIMIT $2`, [wsId, limit]
    );
    res.json(r.rows);
  } catch (e: any) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/setup/status", async (_req, res) => {
  try {
    res.json({ needsBootstrap: await needsBootstrap() });
  } catch (e: any) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/auth/me", async (req: any, res) => {
  try {
    const user = await loadSession(req.cookies?.[SESSION_COOKIE]);
    if (!user) return res.status(401).json({ error: "Chưa đăng nhập." });
    res.json({ user: toUserAccount(user) });
  } catch (e: any) {
    res.status(500).json({ error: e.message });
  }
});

// Open self-service signup: anyone may create their own account and their own new
// workshop. The role is hardcoded here and any role/workshopId in the request body is
// ignored, so this endpoint can never mint a super admin or attach the caller to
// someone else's workshop — that is what keeps it safe to leave open.
const SIGNUP_ROLE = "workshop_admin";

app.post("/api/auth/register", async (req, res) => {
  const { username, password, fullName, workshopName, workshopCode, workshopDesc, companyName, notifyEmail } = req.body || {};
  if (!String(username || "").trim() || !String(password || "").trim()
      || !String(workshopName || "").trim()) {
    return res.status(400).json({ error: "Vui lòng nhập đủ tên đăng nhập, mật khẩu và tên phân xưởng." });
  }
  if (String(password).length < 6) {
    return res.status(400).json({ error: "Mật khẩu phải có ít nhất 6 ký tự." });
  }

  const uname = String(username).trim();
  const wsName = String(workshopName).trim();
  const wsCode = String(workshopCode || "").trim() || wsName.substring(0, 5).toUpperCase();
  const client = await sqlPool.connect();
  try {
    const taken = await client.query(
      `SELECT 1 FROM user_accounts WHERE lower(username) = lower($1)`, [uname]
    );
    if (taken.rowCount) {
      return res.status(409).json({ error: "Tên đăng nhập này đã được sử dụng." });
    }

    await client.query("BEGIN");
    const wsId = "ws_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
    await client.query(
      `INSERT INTO workshops (id, name, code, description, staff_data, config, features)
       VALUES ($1, $2, $3, $4, $5, $6, $7)`,
      [wsId, wsName, wsCode, String(workshopDesc || "").trim() || `Phân xưởng ${wsName}`,
       JSON.stringify([]),
       JSON.stringify({
         companyName: String(companyName || "").trim() || "CÔNG TY THỦY ĐIỆN IALY",
         headerWorkshopName: wsName.toUpperCase(),
         documentCodeSuffix: `/${wsCode}`,
         recipientWorkshopName: wsName,
         shortWorkshopName: wsName,
         locationName: "Gia Lai",
         soVanBan: "",
         nguoiKy: "Lãnh đạo Phân xưởng",
         chucVuNguoiKy: "Quản đốc Phân xưởng",
         notifyEmail: String(notifyEmail || "").trim()
       }),
       JSON.stringify({})]
    );

    const userId = "user_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
    const inserted = await client.query(
      `INSERT INTO user_accounts (id, username, password_hash, full_name, role, workshop_id)
       VALUES ($1, $2, $3, $4, $5, $6) RETURNING *`,
      [userId, uname, hashPassword(String(password)), String(fullName || "").trim() || uname, SIGNUP_ROLE, wsId]
    );
    await client.query("COMMIT");

    setSessionCookie(req, res, await createSession(userId));
    res.json({ success: true, user: toUserAccount(inserted.rows[0]) });
  } catch (e: any) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Error registering:", e);
    res.status(500).json({ error: e.message });
  } finally {
    client.release();
  }
});

app.post("/api/auth/logout", async (req: any, res) => {
  try {
    await destroySession(req.cookies?.[SESSION_COOKIE]);
  } catch (e) {
    console.error("Logout cleanup failed:", e);
  }
  res.clearCookie(SESSION_COOKIE, { path: "/" });
  res.json({ success: true });
});

function hashPassword(password: string): string {
  const salt = crypto.randomBytes(16).toString("hex");
  const hash = crypto.scryptSync(password, salt, 64).toString("hex");
  return `${salt}:${hash}`;
}

function verifyPassword(password: string, stored: string): boolean {
  const [salt, hash] = String(stored || "").split(":");
  if (!salt || !hash) return false;
  const check = crypto.scryptSync(password, salt, 64).toString("hex");
  const a = Buffer.from(hash, "hex");
  const b = Buffer.from(check, "hex");
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function toUserAccount(row: any) {
  return {
    id: row.id,
    username: row.username,
    fullName: row.full_name,
    role: row.role,
    workshopId: row.workshop_id || "all",
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function toWorkshop(row: any) {
  return {
    id: row.id,
    name: row.name,
    code: row.code,
    description: row.description || "",
    staffData: row.staff_data || [],
    config: row.config || {},
    features: row.features || {},
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

// Throttling failed logins. Without it a weak password is only as strong as how fast
// an attacker can send requests. Kept in memory on purpose: this resets on restart,
// which is fine for a lockout, and avoids a database write on every wrong guess.
const LOGIN_MAX_ATTEMPTS = 8;
const LOGIN_WINDOW_MS = 15 * 60 * 1000;
const loginAttempts = new Map<string, { count: number; firstAt: number; until: number }>();

function loginKey(req: any, username: string): string {
  // Deliberately the socket address, not X-Forwarded-For: that header is set by the
  // caller and an attacker could rotate it to sidestep the lockout entirely. Behind a
  // reverse proxy every request shares one address, which makes this effectively a
  // per-username lockout — still enough to stop guessing at speed.
  const ip = req.socket?.remoteAddress || "unknown";
  return `${ip}|${String(username).trim().toLowerCase()}`;
}

function loginLockRemaining(key: string): number {
  const rec = loginAttempts.get(key);
  if (!rec) return 0;
  if (rec.until > Date.now()) return Math.ceil((rec.until - Date.now()) / 1000);
  if (Date.now() - rec.firstAt > LOGIN_WINDOW_MS) loginAttempts.delete(key);
  return 0;
}

function recordLoginFailure(key: string) {
  const now = Date.now();
  const rec = loginAttempts.get(key);
  if (!rec || now - rec.firstAt > LOGIN_WINDOW_MS) {
    loginAttempts.set(key, { count: 1, firstAt: now, until: 0 });
    return;
  }
  rec.count += 1;
  if (rec.count >= LOGIN_MAX_ATTEMPTS) rec.until = now + LOGIN_WINDOW_MS;
}

// The map would otherwise grow forever on a long-running server.
setInterval(() => {
  const now = Date.now();
  for (const [k, v] of loginAttempts) {
    if (v.until < now && now - v.firstAt > LOGIN_WINDOW_MS) loginAttempts.delete(k);
  }
}, LOGIN_WINDOW_MS).unref();

app.post("/api/auth/login", async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: "Thiếu tên đăng nhập hoặc mật khẩu." });
  }

  const key = loginKey(req, username);
  const locked = loginLockRemaining(key);
  if (locked > 0) {
    return res.status(429).json({
      error: `Sai quá nhiều lần. Vui lòng thử lại sau ${Math.ceil(locked / 60)} phút.`
    });
  }

  try {
    const result = await sqlPool.query(
      `SELECT * FROM user_accounts WHERE lower(username) = lower(trim($1))`,
      [username]
    );
    const acc = result.rows[0];
    if (!acc || !verifyPassword(password, acc.password_hash)) {
      recordLoginFailure(key);
      return res.status(401).json({ error: "Tên đăng nhập hoặc mật khẩu không chính xác." });
    }
    loginAttempts.delete(key);
    // Drop any expired rows now that we are writing one anyway.
    await sqlPool.query(`DELETE FROM user_sessions WHERE expires_at <= now()`);
    setSessionCookie(req, res, await createSession(acc.id));
    res.json({ success: true, user: toUserAccount(acc) });
  } catch (e: any) {
    console.error("Error logging in:", e);
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/workshops", async (req: any, res) => {
  try {
    // Signup is open to anyone, so this list must not double as a directory of every
    // other workshop's staff roster and config. Only a super admin sees them all.
    if (req.user && !isSuperAdmin(req)) {
      const wsId = req.user.workshopId;
      if (!wsId || wsId === "all") return res.json([]);
      const own = await sqlPool.query(`SELECT * FROM workshops WHERE id = $1`, [wsId]);
      return res.json(own.rows.map(toWorkshop));
    }
    const result = await sqlPool.query(`SELECT * FROM workshops ORDER BY created_at ASC`);
    res.json(result.rows.map(toWorkshop));
  } catch (e: any) {
    console.error("Error fetching workshops:", e);
    res.status(500).json({ error: e.message });
  }
});

// Upsert: creates a new workshop when `id` is absent, updates in place when present
// (matches WorkshopManagerModal, which reuses this one endpoint for both).
app.post("/api/workshops", express.json({ limit: "5mb" }), async (req: any, res) => {
  const { id, name, code, description, staffData, config, features } = req.body;
  if (!name || !String(name).trim()) {
    return res.status(400).json({ error: "Thiếu tên phân xưởng." });
  }

  // This route both creates and updates. A workshop admin may save their own workshop
  // (the app autosaves config here) but not spin up extra ones or edit anyone else's;
  // public signup creates its workshop through /api/auth/register instead.
  if (req.user && !isSuperAdmin(req)) {
    if (!id) {
      return res.status(403).json({ error: "Chỉ Super Admin mới tạo được phân xưởng mới." });
    }
    if (String(id) !== String(req.user.workshopId || "")) {
      return res.status(403).json({ error: "Bạn chỉ sửa được phân xưởng của mình." });
    }
  }

  try {
    const wsId = id || "ws_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
    const result = await sqlPool.query(
      `INSERT INTO workshops (id, name, code, description, staff_data, config, features)
       VALUES ($1,$2,$3,$4,$5,$6,$7)
       ON CONFLICT (id) DO UPDATE SET
         name = EXCLUDED.name, code = EXCLUDED.code, description = EXCLUDED.description,
         staff_data = EXCLUDED.staff_data, config = EXCLUDED.config, features = EXCLUDED.features,
         updated_at = now()
       RETURNING *`,
      [
        wsId,
        String(name).trim(),
        String(code || "").trim(),
        String(description || "").trim(),
        JSON.stringify(staffData || []),
        JSON.stringify(config || {}),
        JSON.stringify(features || {})
      ]
    );
    res.json({ success: true, workshop: toWorkshop(result.rows[0]) });
  } catch (e: any) {
    console.error("Error saving workshop:", e);
    res.status(500).json({ error: e.message });
  }
});

app.delete("/api/workshops/:id", async (req: any, res) => {
  try {
    if (!isSuperAdmin(req)) {
      return res.status(403).json({ error: "Chỉ Super Admin mới xóa được phân xưởng." });
    }
    const result = await sqlPool.query(`DELETE FROM workshops WHERE id = $1`, [req.params.id]);
    if (result.rowCount === 0) return res.status(404).json({ error: "Không tìm thấy phân xưởng." });
    res.json({ success: true });
  } catch (e: any) {
    console.error("Error deleting workshop:", e);
    res.status(500).json({ error: e.message });
  }
});

// Roles a request is ever allowed to name. "super_admin" is deliberately absent from
// anything a non-super-admin can reach, so the only way to become one is for an
// existing super admin to grant it.
const ASSIGNABLE_ROLES = new Set(["super_admin", "workshop_admin", "user"]);

function isSuperAdmin(req: any): boolean {
  return req.user?.role === "super_admin";
}

app.get("/api/accounts", async (req: any, res) => {
  try {
    // A workshop admin manages only their own workshop's accounts; the full directory
    // (including who holds super_admin) is for super admins.
    if (!isSuperAdmin(req)) {
      const wsId = req.user?.workshopId;
      if (!wsId || wsId === "all") return res.json([]);
      const own = await sqlPool.query(
        `SELECT * FROM user_accounts WHERE workshop_id = $1 ORDER BY created_at ASC`, [wsId]
      );
      return res.json(own.rows.map(toUserAccount));
    }
    const result = await sqlPool.query(`SELECT * FROM user_accounts ORDER BY created_at ASC`);
    res.json(result.rows.map(toUserAccount));
  } catch (e: any) {
    console.error("Error fetching accounts:", e);
    res.status(500).json({ error: e.message });
  }
});

// Upsert: creates when `id` is absent, updates when present. Password is only
// changed if a new one is supplied (matches the "leave blank to keep" edit form).
app.post("/api/accounts", async (req: any, res) => {
  const { id, username, password, fullName, role, workshopId } = req.body;
  if (!username || !String(username).trim() || !fullName || !String(fullName).trim() || !role) {
    return res.status(400).json({ error: "Thiếu tên đăng nhập, họ tên hoặc quyền hạn." });
  }

  if (!ASSIGNABLE_ROLES.has(String(role))) {
    return res.status(400).json({ error: "Quyền hạn không hợp lệ." });
  }

  // req.user is absent only on the empty-database bootstrap path allowed by requireAuth.
  const bootstrapping = !req.user;
  if (!bootstrapping && !isSuperAdmin(req)) {
    // A workshop admin may manage staff logins inside their own workshop and nothing
    // more: they cannot grant super_admin, and cannot touch another workshop.
    if (String(role) === "super_admin") {
      return res.status(403).json({ error: "Chỉ Super Admin mới cấp được quyền Super Admin." });
    }
    if (String(workshopId || "") !== String(req.user.workshopId || "")) {
      return res.status(403).json({ error: "Bạn chỉ quản lý được tài khoản trong phân xưởng của mình." });
    }
    if (id) {
      const target = await sqlPool.query(`SELECT role, workshop_id FROM user_accounts WHERE id = $1`, [id]);
      const t = target.rows[0];
      if (!t) return res.status(404).json({ error: "Không tìm thấy tài khoản." });
      if (t.role === "super_admin" || String(t.workshop_id || "") !== String(req.user.workshopId || "")) {
        return res.status(403).json({ error: "Không có quyền sửa tài khoản này." });
      }
    }
  }

  // "all" (super admin, not tied to one workshop) isn't a real workshop id, so the
  // workshop_id FK would reject it — store NULL instead; toUserAccount() maps NULL back to "all".
  const normalizedWorkshopId = workshopId && workshopId !== "all" ? workshopId : null;

  try {
    if (id) {
      const existing = await sqlPool.query(`SELECT password_hash FROM user_accounts WHERE id = $1`, [id]);
      if (existing.rows.length === 0) return res.status(404).json({ error: "Không tìm thấy tài khoản." });

      const dup = await sqlPool.query(
        `SELECT id FROM user_accounts WHERE lower(username) = lower(trim($1)) AND id <> $2`,
        [username, id]
      );
      if (dup.rows.length > 0) return res.status(400).json({ error: "Tên đăng nhập đã được dùng bởi tài khoản khác." });

      const passwordHash = password ? hashPassword(password) : existing.rows[0].password_hash;
      // A password change is how you lock out someone who has your credentials or a
      // stolen session cookie, so it has to end the existing sessions too.
      if (password) {
        await sqlPool.query(`DELETE FROM user_sessions WHERE user_id = $1`, [id]);
      }
      const result = await sqlPool.query(
        `UPDATE user_accounts SET username=$1, password_hash=$2, full_name=$3, role=$4, workshop_id=$5, updated_at=now()
         WHERE id=$6 RETURNING *`,
        [String(username).trim(), passwordHash, String(fullName).trim(), role, normalizedWorkshopId, id]
      );
      return res.json({ success: true, account: toUserAccount(result.rows[0]) });
    }

    const dup = await sqlPool.query(`SELECT id FROM user_accounts WHERE lower(username) = lower(trim($1))`, [username]);
    if (dup.rows.length > 0) return res.status(400).json({ error: "Tên đăng nhập đã tồn tại." });

    const newId = "user_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
    const passwordHash = hashPassword(password || "123456");
    const result = await sqlPool.query(
      `INSERT INTO user_accounts (id, username, password_hash, full_name, role, workshop_id)
       VALUES ($1,$2,$3,$4,$5,$6) RETURNING *`,
      [newId, String(username).trim(), passwordHash, String(fullName).trim(), role, normalizedWorkshopId]
    );
    res.json({ success: true, account: toUserAccount(result.rows[0]) });
  } catch (e: any) {
    console.error("Error saving account:", e);
    res.status(500).json({ error: e.message });
  }
});

app.delete("/api/accounts/:id", async (req: any, res) => {
  try {
    if (!isSuperAdmin(req)) {
      const target = await sqlPool.query(`SELECT role, workshop_id FROM user_accounts WHERE id = $1`, [req.params.id]);
      const t = target.rows[0];
      if (!t) return res.status(404).json({ error: "Không tìm thấy tài khoản." });
      if (t.role === "super_admin" || String(t.workshop_id || "") !== String(req.user?.workshopId || "")) {
        return res.status(403).json({ error: "Không có quyền xóa tài khoản này." });
      }
    }
    const result = await sqlPool.query(`DELETE FROM user_accounts WHERE id = $1`, [req.params.id]);
    if (result.rowCount === 0) return res.status(404).json({ error: "Không tìm thấy tài khoản." });
    res.json({ success: true });
  } catch (e: any) {
    console.error("Error deleting account:", e);
    res.status(500).json({ error: e.message });
  }
});

// App Settings Endpoints (Staff List & Config)
app.get("/api/app-settings", async (req, res) => {
  try {
    const data = await readConfig("app_settings");
    if (!data) {
      return res.json({ staffData: null, config: null });
    }

    if (typeof data.staffData === 'string') {
      try {
        data.staffData = JSON.parse(data.staffData);
      } catch (e) {
        console.error("Failed to parse staffData JSON", e);
        data.staffData = []; // Fallback to empty array on parse error
      }
    }

    let needsSave = false;

    // Ensure it's an array
    if (!Array.isArray(data.staffData)) {
      data.staffData = [];
    } else {
      // Perform server-side migration for 'Trực phụ cơ MR' -> 'Trực phụ máy MR'
      let hasMigration = false;
      const migrated = data.staffData.map((row: any) => {
        if (Array.isArray(row) && row[0] === 'Trực phụ cơ MR') {
          hasMigration = true;
          return ['Trực phụ máy MR', ...row.slice(1)];
        }
        return row;
      });
      if (hasMigration) {
        data.staffData = migrated;
        needsSave = true;
        console.log("Migrated 'Trực phụ cơ MR' to 'Trực phụ máy MR'.");
      }
    }

    // Migrate Zalo Webhook URL
    if (
      data.config &&
      data.config.zaloWebhookUrl &&
      (
        data.config.zaloWebhookUrl.includes("cookies-blue-pen-bikini.trycloudflare.com") ||
        data.config.zaloWebhookUrl.includes("specialists-intro-exterior-advocacy.trycloudflare.com") ||
        data.config.zaloWebhookUrl.includes("committed-intellectual-lunch-clone.trycloudflare.com")
      )
    ) {
      data.config.zaloWebhookUrl = "https://vhialy.dpdns.org/webhook/notify";
      needsSave = true;
      console.log("Migrated 'zaloWebhookUrl' to 'vhialy.dpdns.org'.");
    }

    if (needsSave) {
      // staffData is stored as a JSON string, matching how the POST handler writes it
      await writeConfig("app_settings", { ...data, staffData: JSON.stringify(data.staffData) });
    }

    res.json(data);
  } catch (e: any) {
    console.error("Error fetching app settings:", e);
    res.status(500).json({ error: e.message });
  }
});

app.post("/api/app-settings", express.json({ limit: '5mb' }), async (req, res) => {
  const { staffData, config } = req.body;
  try {
    // Merge into the stored document so unrelated keys (e.g. staffBackups) survive
    const existing = (await readConfig("app_settings")) || {};
    const ok = await writeConfig("app_settings", {
      ...existing,
      staffData: JSON.stringify(staffData),
      config,
      updatedAt: new Date().toISOString()
    });
    if (!ok) throw new Error(lastDbError || "unknown");
    res.json({ success: true });
  } catch (e: any) {
    console.error("Error saving app settings:", e);
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/debug/env", (req, res) => {
  const id = process.env.GOOGLE_CLIENT_ID || "";
  const secret = process.env.GOOGLE_CLIENT_SECRET || "";
  const url = process.env.APP_URL || "";
  res.json({
    id_exists: !!id,
    id_length: id.length,
    id_start: id.substring(0, 5),
    id_end: id.substring(id.length - 5),
    secret_exists: !!secret,
    secret_length: secret.length,
    url_exists: !!url,
    url: url
  });
});

// Auth Routes
app.get("/api/auth/google", (req, res) => {
  try {
    const oauth2Client = getOAuth2Client();
    const url = oauth2Client.generateAuthUrl({
      access_type: "offline",
      scope: [
        "https://www.googleapis.com/auth/spreadsheets", 
        "https://www.googleapis.com/auth/userinfo.email",
        "https://www.googleapis.com/auth/drive"
      ],
      prompt: "consent" // Force consent to ensure we get a refresh_token
    });
    res.redirect(url);
  } catch (error) {
    console.error("Error generating Auth URL:", error);
    res.status(500).json({ 
      error: "Failed to generate Google Auth URL", 
      details: error instanceof Error ? error.message : String(error) 
    });
  }
});

app.get("/api/auth/google/url", (req, res) => {
  const oauth2Client = getOAuth2Client();
  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: [
      "https://www.googleapis.com/auth/spreadsheets", 
      "https://www.googleapis.com/auth/userinfo.email",
      "https://www.googleapis.com/auth/drive"
    ],
    prompt: "consent"
  });
  res.json({ url });
});

app.get("/api/auth/google/callback", async (req, res) => {
  const oauth2Client = getOAuth2Client();
  const { code } = req.query;
  try {
    const { tokens } = await oauth2Client.getToken(code as string);
    
    // Save tokens to server for "permanent" access for everyone
    await saveTokens(tokens);

    res.cookie("google_tokens", JSON.stringify(tokens), {
      httpOnly: true,
      secure: true,
      sameSite: "none",
    });
    res.send(`
      <html>
        <body>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS' }, '*');
              setTimeout(() => window.close(), 1000);
            } else {
              window.location.href = '/';
            }
          </script>
          <p>Authentication successful. This window should close automatically.</p>
        </body>
      </html>
    `);
  } catch (error) {
    console.error("Error getting tokens:", error);
    res.status(500).send("Authentication failed");
  }
});

app.get("/api/auth/status", async (req, res) => {
  const cookieTokensStr = req.cookies.google_tokens;
  let storedTokens = await loadTokens();

  // Auto-sync: If cookie has tokens but the database doesn't, try to save them
  if (cookieTokensStr && !storedTokens) {
    console.log("Cookie exists but database is empty. Attempting auto-sync...");
    try {
      const tokens = JSON.parse(cookieTokensStr);
      await saveTokens(tokens);
      storedTokens = await loadTokens();
    } catch (e) {
      console.error("Auto-sync failed:", e);
    }
  }

  res.json({
    authenticated: !!cookieTokensStr || !!storedTokens,
    source: cookieTokensStr ? "cookie" : (storedTokens ? "database" : "none"),
    sync_attempted: cookieTokensStr && !storedTokens
  });
});

app.get("/api/sheets/leave-requests", async (req, res) => {
  const workshopId = req.query.workshopId;
  if (!workshopId) return res.status(400).json({ error: "Thiếu phân xưởng." });
  try {
    const result = await sqlPool.query(
      `SELECT r.id, r.name, r.birth_year, r.chuc_danh, r.kip, r.start_date, r.end_date, r.reason, r.phone, r.location,
              r.status, r.created_at, r.leave_year, r.has_leave_permit, r.distance_km, r.travel_days, r.leave_days,
              COALESCE(
                (SELECT json_agg(json_build_object('year', a.leave_year, 'days', a.days) ORDER BY a.leave_year)
                 FROM leave_allocations a WHERE a.leave_id = r.id),
                '[]'::json
              ) AS allocations
       FROM leave_requests r WHERE r.workshop_id = $1 ORDER BY r.inserted_at ASC`,
      [workshopId]
    );

    const leaveRequests = result.rows.map((row: any) => ({
      id: row.id || "",
      name: row.name || "",
      birthYear: row.birth_year || "",
      chucDanh: row.chuc_danh || "",
      kip: row.kip || "",
      startDate: row.start_date || "",
      endDate: row.end_date || "",
      reason: row.reason || "",
      phone: row.phone || "",
      location: row.location || "",
      status: row.status || "",
      createdAt: row.created_at || "",
      leaveYear: row.leave_year || "",
      hasLeavePermit: row.has_leave_permit === true,
      distanceKm: row.distance_km,
      travelDays: Number(row.travel_days || 0),
      leaveDays: Number(row.leave_days || 0),
      allocations: row.allocations || []
    }));

    res.json(leaveRequests);
  } catch (error: any) {
    console.error("Error fetching leave requests from SQL:", error);
    res.status(500).json({ error: "Failed to fetch leave requests", details: error.message });
  }
});

app.post("/api/sheets/leave-requests", async (req, res) => {
  const { name, birthYear, chucDanh, kip, startDate, endDate, reason, phone, location, leaveBalance, leaveYear, hasLeavePermit, workshopId } = req.body;

  if (!name || !chucDanh || !kip || !startDate || !endDate) {
    return res.status(400).json({ error: "Thiếu thông tin bắt buộc" });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }

  const normalizedChucDanh = chucDanh === "Trực phụ cơ MR" ? "Trực phụ máy MR" : chucDanh;
  const id = "LEAVE_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
  const dateStr = new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
  const status = "Chờ phân ca";
  const finalHasLeavePermit = hasLeavePermit === true || hasLeavePermit === "true";

  // The leave year is derived from the rota and the carry-over rule, not taken from the client.
  const plan = await planLeaveRequest({
    name,
    workshopId,
    kip,
    startDate,
    endDate,
    location: location || "Gia Lai",
    hasLeavePermit: finalHasLeavePermit
  });
  const travel = plan.travel;
  const finalLeaveYear = plan.leaveYear;

  // Try to get leave balance if not supplied in the request body
  let finalLeaveBalance = leaveBalance;
  if (!finalLeaveBalance) {
    finalLeaveBalance = await fetchLeaveBalanceHelper(name, workshopId, finalLeaveYear);
  }

  // Kick off both notification channels now, in parallel with the DB write below, so the
  // save isn't delayed by network calls — but keep the promises (instead of the old
  // fire-and-forget) so the response can report what actually happened, awaited right
  // before the response is built rather than here.
  const notifyPayload = {
    name,
    chucDanh: normalizedChucDanh,
    kip: String(kip),
    startDate,
    endDate,
    reason,
    phone,
    location,
    dateStr,
    leaveBalance: finalLeaveBalance
  };
  const zaloPromise = sendZaloNotification(workshopId, notifyPayload);
  const emailPromise = sendEmailNotification(workshopId, notifyPayload);

  const client = await sqlPool.connect();
  try {
    await client.query("BEGIN");
    await client.query(
      `INSERT INTO leave_requests (id, name, birth_year, chuc_danh, kip, start_date, end_date, reason, phone, location, status, created_at, leave_year, has_leave_permit, distance_km, travel_days, leave_days, workshop_id)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18)`,
      [
        id,
        name,
        birthYear || "",
        normalizedChucDanh,
        String(kip),
        startDate,
        endDate,
        reason || "Giải quyết việc riêng gia đình",
        phone || "",
        location || "Gia Lai",
        status,
        dateStr,
        finalLeaveYear,
        finalHasLeavePermit,
        travel.distanceKm,
        travel.travelDays,
        plan.leaveDays,
        workshopId
      ]
    );

    for (const [year, days] of Object.entries(plan.allocations)) {
      await client.query(
        `INSERT INTO leave_allocations (leave_id, leave_year, days) VALUES ($1,$2,$3)`,
        [id, year, days]
      );
    }
    await client.query("COMMIT");

    const [zaloResult, emailResult] = await Promise.all([zaloPromise, emailPromise]);
    // Only the email outcome is surfaced, per the wording asked for. The Zalo result and
    // the leave-day breakdown still travel in the JSON below (notify / allocations), so
    // nothing is lost for callers that need them — they are just kept out of the banner.
    let notifyMsg: string;
    if (emailResult.ok) notifyMsg = " Đã gửi đơn nghỉ phép tới gmail TKPX.";
    else if (emailResult.skipped) notifyMsg = " Chưa cấu hình gmail TKPX nên chưa gửi được đơn.";
    else notifyMsg = " Gửi đơn nghỉ phép tới gmail TKPX thất bại.";

    res.json({
      success: true,
      message: `✅ Đã lưu đơn của đồng chí ${name}!${notifyMsg}`,
      id,
      status,
      leaveDays: plan.leaveDays,
      leaveYear: finalLeaveYear,
      allocations: plan.allocations,
      travelDays: travel.travelDays,
      distanceKm: travel.distanceKm,
      travelNote: travel.note,
      notify: {
        zalo: zaloResult,
        email: emailResult
      }
    });
  } catch (error: any) {
    await client.query("ROLLBACK").catch(() => {});
    console.error("Error saving leave request to SQL:", error);
    res.status(500).json({ error: "Lưu đơn thất bại", details: error.message });
  } finally {
    client.release();
  }
});

app.get("/api/sheets/leave-balance", async (req, res) => {
  const { name, year, workshopId } = req.query;
  if (!name) {
    return res.status(400).json({ error: "Thiếu tên nhân viên" });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }

  const targetYear = String(year || new Date().getFullYear());
  try {
    const quotas = await getQuotaByYear(String(name), String(workshopId), undefined, [targetYear]);
    const years = Object.keys(quotas).sort();
    const current = quotas[targetYear];
    if (!current || (current.entitled === 0 && current.source !== "manual")) {
      return res.status(404).json({ error: `Chưa khai báo số ngày phép năm của đồng chí ${name}` });
    }

    return res.json({
      success: true,
      name,
      year: current.year,
      entitled: String(current.entitled + current.travelDays),
      used: String(current.used),
      remaining: String(current.remaining),
      byYear: years.map(y => quotas[y])
    });
  } catch (error: any) {
    console.error("Error fetching leave balance from SQL:", error);
    return res.status(500).json({ error: "Lỗi truy vấn dữ liệu", details: error.message });
  }
});

app.post("/api/sheets/leave-requests/update-status", async (req, res) => {
  const { ids, status, workshopId } = req.body;

  if (!ids || !Array.isArray(ids) || ids.length === 0 || !status) {
    return res.status(400).json({ error: "Tham số không hợp lệ." });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }

  try {
    const result = await sqlPool.query(
      `UPDATE leave_requests SET status = $1 WHERE workshop_id = $2 AND id = ANY($3::text[]) RETURNING id`,
      [status, workshopId, ids]
    );

    res.json({ success: true, updatedCount: result.rowCount, updatedRows: result.rows });
  } catch (error: any) {
    console.error("Error updating leave status in SQL:", error);
    res.status(500).json({ error: "Failed to update leave status", details: error.message });
  }
});

app.post("/api/sheets/leave-requests/delete", async (req, res) => {
  const { ids, workshopId } = req.body;
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ error: "Tham số không hợp lệ." });
  }
  if (!workshopId) {
    return res.status(400).json({ error: "Thiếu phân xưởng." });
  }

  try {
    const normalizedTargetIds = ids.map((id: any) => String(id).trim());
    const result = await sqlPool.query(
      `DELETE FROM leave_requests WHERE workshop_id = $1 AND id = ANY($2::text[])`,
      [workshopId, normalizedTargetIds]
    );

    if (result.rowCount === 0) {
      return res.status(404).json({ error: "Không tìm thấy mã đơn cần xóa trong cơ sở dữ liệu." });
    }

    res.json({ success: true, deletedCount: result.rowCount });
  } catch (error: any) {
    console.error("Error deleting leave request from SQL:", error);
    res.status(500).json({ error: "Failed to delete leave request", details: error.message });
  }
});

app.post("/api/sheets/update", async (req, res) => {
  const oauth2Client = getOAuth2Client();
  let tokens = null;
  
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try {
      tokens = JSON.parse(tokensStr);
    } catch (e) {
      console.error("Error parsing cookie tokens:", e);
    }
  } 
  
  if (!tokens) {
    tokens = await loadTokens();
  }

  if (!tokens) return res.status(401).json({ error: "Not authenticated" });

  oauth2Client.setCredentials(tokens);

  // Listen for token refreshes and save them
  oauth2Client.on('tokens', (newTokens) => {
    console.log("Tokens refreshed, saving to database...");
    saveTokens(newTokens);
  });

  const { spreadsheetId, updates } = req.body;
  // updates: Array<{ date: string, person: string, shift: string }>

  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  try {
    // Get spreadsheet metadata to find sheet names
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];

    const sheetsCache: { [range: string]: any[][] } = {};
    const batchData: any[] = [];

    for (const update of updates) {
      const date = new Date(update.date);
      const month = date.getMonth() + 1;
      const monthStr = month < 10 ? `0${month}` : `${month}`;
      const year = date.getFullYear();
      const day = date.getDate();
      const dayStr = day < 10 ? `0${day}` : `${day}`;
      
      // Match LICH_MM_YYYY or other common formats
      const sheetName = sheetNames.find(n => 
        n === `LICH_${monthStr}_${year}` ||
        n === `LICH_${month}_${year}` ||
        n?.includes(`Tháng ${month}`) || 
        n?.includes(`Tháng ${monthStr}`) || 
        n?.includes(`T${month}`) ||
        n?.includes(`T${monthStr}`) ||
        n === `${month}` || 
        n === monthStr
      );

      if (!sheetName) {
        console.warn(`Sheet for month ${month} not found in:`, sheetNames);
        continue;
      }

      // Read the sheet to find the person and the date (with cache)
      const range = `${sheetName}!A1:CZ500`; // Increased range to cover more columns (up to CZ) and rows
      let rows = sheetsCache[range];
      if (!rows) {
        console.log(`Cache miss: Fetching sheet range ${range}...`);
        const response = await sheets.spreadsheets.values.get({ spreadsheetId, range });
        rows = response.data.values || [];
        sheetsCache[range] = rows;
      } else {
        console.log(`Cache hit: Using cached rows for range ${range}`);
      }

      if (rows.length === 0) {
        console.warn(`No data found in sheet: ${sheetName}`);
        continue;
      }

      // Find person column (assuming names are in the first row)
      const header = rows[0];
      let personColIdx = -1;
      const searchName = update.name.trim().normalize('NFC').toLowerCase();
      
      for (let i = 1; i < header.length; i++) {
        const cell = header[i]?.toString().trim().normalize('NFC').toLowerCase();
        if (cell && cell.includes(searchName)) {
          personColIdx = i;
          break;
        }
      }

      if (personColIdx === -1) {
        console.warn(`Person "${update.name}" not found in sheet: ${sheetName}. Normalized search: "${searchName}". Header columns:`, header.map(h => h?.toString().normalize('NFC')));
        continue;
      }

      // Find date row (assuming dates are in the first column)
      let dateRowIdx = -1;
      const dateParts = update.date.split('-'); // YYYY-MM-DD
      const dayNum = parseInt(dateParts[2]);
      const currentDayStr = dayNum < 10 ? `0${dayNum}` : `${dayNum}`;
      const dd_mm_yyyy = `${currentDayStr}/${dateParts[1]}/${dateParts[0]}`;
      const d_m_yyyy = `${dayNum}/${parseInt(dateParts[1])}/${dateParts[0]}`;

      for (let i = 1; i < rows.length; i++) {
        const cell = rows[i][0]?.toString().trim();
        if (!cell) continue;
        
        // Match various date formats: full date, dd/mm/yyyy, or just the day number
        if (
          cell === update.date || 
          cell === dd_mm_yyyy || 
          cell === d_m_yyyy || 
          cell.includes(dd_mm_yyyy) ||
          cell === dayNum.toString() ||
          cell === currentDayStr
        ) {
          dateRowIdx = i;
          break;
        }
      }

      if (dateRowIdx === -1) {
        console.warn(`Date "${update.date}" not found in sheet: ${sheetName}. First column sample:`, rows.slice(0, 10).map(r => r[0]));
        continue;
      }

      // Update the cell in LICH sheet
      const colLetter = getColumnLetter(personColIdx);
      const cellRange = `${sheetName}!${colLetter}${dateRowIdx + 1}`;
      console.log(`Adding to batch: ${update.name} on ${update.date} to ${update.shift} at ${cellRange}`);
      batchData.push({
        range: cellRange,
        values: [[update.shift]]
      });

      // Also try to update TONG_HOP sheet if it exists
      const summarySheetName = `TONG_HOP_${monthStr}_${year}`;
      if (sheetNames.includes(summarySheetName)) {
        const summaryRange = `${summarySheetName}!A1:CZ500`; // Increased range
        let summaryRows = sheetsCache[summaryRange];
        if (!summaryRows) {
          console.log(`Cache miss: Fetching summary sheet range ${summaryRange}...`);
          const summaryResponse = await sheets.spreadsheets.values.get({ spreadsheetId, range: summaryRange });
          summaryRows = summaryResponse.data.values || [];
          sheetsCache[summaryRange] = summaryRows;
        } else {
          console.log(`Cache hit: Using cached rows for range ${summaryRange}`);
        }
        
        if (summaryRows.length > 0) {
          // Find person row in TONG_HOP (Column A, starting from row 5)
          let summaryRowIdx = -1;
          for (let i = 4; i < summaryRows.length; i++) {
            const cell = summaryRows[i][0]?.toString().trim();
            if (cell && cell.toLowerCase() === update.name.toLowerCase().trim()) {
              summaryRowIdx = i;
              break;
            }
          }

          // Find day column in TONG_HOP (Row 1, starting from column B)
          const dayNum = parseInt(dateParts[2]);
          const summaryColIdx = dayNum; // Day 1 is in Column B (index 1)

          if (summaryRowIdx !== -1) {
            const summaryColLetter = getColumnLetter(summaryColIdx);
            const summaryCellRange = `${summarySheetName}!${summaryColLetter}${summaryRowIdx + 1}`;
            console.log(`Adding to batch (summary): ${update.name} on ${update.date} to ${update.shift} at ${summaryCellRange}`);
            batchData.push({
              range: summaryCellRange,
              values: [[update.shift]]
            });
          }
        }
      }
    }

    if (batchData.length > 0) {
      console.log(`Executing batchUpdate of ${batchData.length} cells...`);
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: "RAW",
          data: batchData
        }
      });
      console.log("batchUpdate executed successfully!");
    }

    // 4. Append notification to "Bảng báo cơm ca" sheet if it exists
    const mealSheetName = sheetNames.find(n => 
      n?.toLowerCase().includes("báo cơm") || 
      n?.toLowerCase().includes("cơm ca") ||
      n?.toLowerCase().includes("meal")
    );
    
    if (mealSheetName) {
      const now = new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
      const mealUpdates = updates.map(u => [
        now, 
        u.name, 
        u.date, 
        u.shift, 
        "Đã thay đổi"
      ]);
      
      try {
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${mealSheetName}!A:E`,
          valueInputOption: "RAW",
          requestBody: { values: mealUpdates }
        });
        console.log(`Appended ${mealUpdates.length} rows to ${mealSheetName}`);
      } catch (appendError) {
        console.error("Error appending to meal sheet:", appendError);
      }
    }

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error updating sheet details:", {
      message: error.message,
      code: error.code,
      errors: error.response?.data?.error?.errors,
      status: error.response?.status
    });
    
    // Handle invalid_grant or insufficient scopes
    const errorMsg = (error.message || "").toLowerCase();
    const isAuthError = 
      errorMsg.includes("invalid_grant") || 
      errorMsg.includes("insufficient authentication scopes") ||
      errorMsg.includes("scope") ||
      error.code === 400 || 
      error.code === 401 ||
      error.code === 403;

    if (isAuthError) {
      console.warn("Auth error detected. Clearing tokens and requesting re-auth.");
      res.clearCookie("google_tokens");
      await clearTokens();
      return res.status(401).json({ 
        error: "Google Authentication expired or has insufficient permissions. Please log in again.",
        details: error.message 
      });
    }

    res.status(500).json({ 
      error: "Failed to update sheet", 
      details: error.message,
      code: error.code
    });
  }
});

app.post("/api/sheets/update-annual-leaves", async (req, res) => {
  const oauth2Client = getOAuth2Client();
  let tokens = null;
  
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try {
      tokens = JSON.parse(tokensStr);
    } catch (e) {
      console.error("Error parsing cookie tokens:", e);
    }
  } 
  
  if (!tokens) {
    tokens = await loadTokens();
  }

  if (!tokens) return res.status(401).json({ error: "Not authenticated" });

  oauth2Client.setCredentials(tokens);

  oauth2Client.on('tokens', (newTokens) => {
    console.log("Tokens refreshed, saving to database...");
    saveTokens(newTokens);
  });

  const { spreadsheetId, updates } = req.body;

  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  try {
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];
    const targetSheetName = sheetNames.find(n => n?.toLowerCase().trim() === "số ngày phép") || "Số ngày phép";

    const range = `'${targetSheetName}'!A1:G500`;
    const response = await sheets.spreadsheets.values.get({ spreadsheetId, range });
    const rows = response.data.values;

    if (!rows || rows.length === 0) {
      return res.status(400).json({ error: `Không tìm thấy dữ liệu trong trang tính ${targetSheetName}` });
    }

    const skippedNames: string[] = [];
    const updatedNames: string[] = [];
    const batchData: any[] = [];

    for (const update of updates) {
      if (!update.name) continue;
      const searchName = update.name.trim().normalize('NFC').toLowerCase();
      
      let rowIdx = -1;
      for (let i = 0; i < rows.length; i++) {
        const cell = rows[i][1]?.toString().trim().normalize('NFC').toLowerCase();
        if (cell === searchName) {
          rowIdx = i;
          break;
        }
      }

      if (rowIdx === -1 && searchName.length > 2) {
        for (let i = 0; i < rows.length; i++) {
          const cell = rows[i][1]?.toString().trim().normalize('NFC').toLowerCase();
          if (cell && (cell.includes(searchName) || searchName.includes(cell))) {
            rowIdx = i;
            break;
          }
        }
      }

      if (rowIdx === -1) {
        console.warn(`Không tìm thấy nhân viên "${update.name}" trong cột B của sheet Số ngày phép.`);
        skippedNames.push(update.name);
        continue;
      }

      const cellRange = `'${targetSheetName}'!F${rowIdx + 1}`;
      console.log(`Adding to batch (annual leaves): ${update.tongcatrucphaithay} ca phép cho ${update.name} tại ${cellRange}`);
      batchData.push({
        range: cellRange,
        values: [[update.tongcatrucphaithay]]
      });
      updatedNames.push(update.name);
    }

    if (batchData.length > 0) {
      console.log(`Executing batchUpdate of ${batchData.length} annual leaves cells...`);
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: batchData
        }
      });
      console.log("Annual leaves batchUpdate executed successfully!");
    }

    res.json({ success: true, updatedNames, skippedNames });
  } catch (error: any) {
    console.error("Lỗi cập nhật bảng theo dõi phép năm:", error);
    res.status(500).json({ 
      error: "Không thể cập nhật bảng theo dõi phép năm", 
      details: error.message,
      code: error.code
    });
  }
});

// Catch-all for API routes that don't exist
app.all("/api/*", (req, res) => {
  res.status(404).json({ error: `API route ${req.path} not found` });
});

function getColumnLetter(columnIdx: number): string {
  let temp, letter = "";
  while (columnIdx >= 0) {
    temp = columnIdx % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnIdx = (columnIdx - temp) / 26 - 1;
  }
  return letter;
}

async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const viteModule = "vite";
    const { createServer: createViteServer } = await import(viteModule);
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  }

  // Listen on port in BOTH dev and production
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error("Failed to start server:", err);
  process.exit(1);
});

export default app;
