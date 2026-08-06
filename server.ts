import express from "express";
// import { createServer as createViteServer } from "vite"; // Move to dynamic import
import path from "path";
import { google } from "googleapis";
import cookieParser from "cookie-parser";
import dotenv from "dotenv";
import https from "https";

import fs from "fs";
import admin from "firebase-admin";
import { getFirestore } from "firebase-admin/firestore";
dotenv.config();

// Helper to send notification to Zalo personal webhook
async function sendZaloNotification(data: {
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
}) {
  let webhookUrl = process.env.ZALO_WEBHOOK_URL || "https://vhialy.dpdns.org/webhook/notify";
  
  if (!process.env.ZALO_WEBHOOK_URL) {
    try {
      const db = getFirestoreInstance();
      const doc = await db.collection("config").doc("app_settings").get();
      if (doc.exists) {
        const docData = doc.data();
        if (docData && docData.config && docData.config.zaloWebhookUrl) {
          let url = docData.config.zaloWebhookUrl;
          if (
            url.includes("cookies-blue-pen-bikini.trycloudflare.com") || 
            url.includes("specialists-intro-exterior-advocacy.trycloudflare.com") ||
            url.includes("committed-intellectual-lunch-clone.trycloudflare.com")
          ) {
            url = "https://vhialy.dpdns.org/webhook/notify";
          }
          webhookUrl = url;
        }
      }
    } catch (dbErr: any) {
      console.log("Unable to load zaloWebhookUrl from Firestore, using default:", dbErr.message);
    }
  }
  
  let textMessage = `CÓ ĐƠN NGHỈ PHÉP MỚI 
• Họ và tên: ${data.name}
• Chức danh: ${data.chucDanh}
• Kíp: Kíp ${data.kip}
• Thời gian: Từ ${data.startDate} đến ${data.endDate}
• Lý do: ${data.reason || "Giải quyết việc riêng gia đình"}`;

  if (data.leaveBalance) {
    textMessage += `
• Phép được hưởng: ${data.leaveBalance.entitled} ngày
• Phép đã nghỉ: ${data.leaveBalance.used} ngày
• Phép còn lại: ${data.leaveBalance.remaining} ngày`;
  }

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

  try {
    const parsedUrl = new URL(webhookUrl);
    const options = {
      hostname: parsedUrl.hostname,
      path: parsedUrl.pathname + parsedUrl.search,
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Content-Length": Buffer.byteLength(payload)
      },
      timeout: 5000
    };

    const req = https.request(options, (res) => {
      let body = "";
      res.on("data", (chunk) => { body += chunk; });
      res.on("end", () => {
        console.log(`Zalo notification response (status ${res.statusCode}): ${body}`);
      });
    });

    req.on("error", (e: any) => {
      console.log("Zalo Webhook unreachable (e.g. offline tunnel or invalid URL):", e.message);
    });

    req.on("timeout", () => {
      req.destroy();
      console.log("Zalo notification request timed out");
    });

    req.write(payload);
    req.end();
  } catch (e: any) {
    console.log("Failed to parse webhook URL or send Zalo notification:", e.message);
  }
}

let firebaseConfig: any = {};
try {
  const configPath = path.join(process.cwd(), "firebase-applet-config.json");
  if (fs.existsSync(configPath)) {
    firebaseConfig = JSON.parse(fs.readFileSync(configPath, "utf-8"));
  } else {
    console.warn("firebase-applet-config.json not found, using environment variables");
    firebaseConfig = {
      projectId: process.env.FIREBASE_PROJECT_ID || process.env.GOOGLE_CLOUD_PROJECT || "default-project",
    };
  }
} catch (e) {
  console.error("Error loading firebase-applet-config.json:", e);
  firebaseConfig = { projectId: process.env.FIREBASE_PROJECT_ID || "default-project" };
}

let loadedServiceAccount: any = null;
try {
  const saPath = path.join(process.cwd(), "service-account.json");
  if (process.env.FIREBASE_SERVICE_ACCOUNT) {
    loadedServiceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
  } else if (fs.existsSync(saPath)) {
    loadedServiceAccount = JSON.parse(fs.readFileSync(saPath, "utf-8"));
  }
} catch (saErr: any) {
  console.warn("Notice reading service-account.json:", saErr.message || saErr);
}

if (loadedServiceAccount) {
  firebaseConfig.projectId = loadedServiceAccount.project_id || firebaseConfig.projectId;
  firebaseConfig.firestoreDatabaseId = "(default)";
}

// Initialize Firebase Admin
if (admin.apps.length === 0) {
  try {
    if (loadedServiceAccount) {
      admin.initializeApp({
        credential: admin.credential.cert(loadedServiceAccount),
        projectId: loadedServiceAccount.project_id || firebaseConfig.projectId,
      });
      console.log(`Firebase Admin initialized with Service Account for project: ${loadedServiceAccount.project_id}`);
    } else {
      admin.initializeApp({
        projectId: firebaseConfig.projectId,
      });
      console.log("Firebase Admin initialized with default credentials");
    }
  } catch (e) {
    console.error("Firebase Admin initialization error:", e);
  }
}
// const firestore = admin.firestore(); // Initialize lazily
let lastFirestoreError: string | null = null;

const getFirestoreInstance = (useDefault = false) => {
  try {
    const app = admin.app();
    if (useDefault) {
      return getFirestore(app);
    }
    
    const dbId = process.env.FIREBASE_DATABASE_ID || firebaseConfig.firestoreDatabaseId;
    
    if (dbId && dbId !== "(default)") {
      // In firebase-admin, the signature is getFirestore(app, databaseId)
      return getFirestore(app, dbId);
    }
    return getFirestore(app);
  } catch (e: any) {
    lastFirestoreError = `GetFirestore Error: ${e.message}`;
    return getFirestore(admin.app());
  }
};

const saveTokens = async (newTokens: any) => {
  const trySave = async (db: any) => {
    const docRef = db.collection("config").doc("google_auth");
    const existingDoc = await docRef.get();
    let finalTokens = { ...newTokens };
    
    if (existingDoc.exists) {
      const existingTokens = existingDoc.data()?.tokens;
      if (!finalTokens.refresh_token && existingTokens?.refresh_token) {
        finalTokens.refresh_token = existingTokens.refresh_token;
      }
    }

    await docRef.set({
      tokens: finalTokens,
      updatedAt: new Date().toISOString()
    });
  };

  try {
    try {
      await trySave(getFirestoreInstance());
      lastFirestoreError = null;
    } catch (e: any) {
      console.warn("Primary DB save failed, trying fallback:", e.message);
      if (e.code === 5 || e.message?.includes("NOT_FOUND") || e.message?.includes("not found")) {
        await trySave(getFirestoreInstance(true));
        lastFirestoreError = null;
      } else {
        throw e;
      }
    }
  } catch (e: any) {
    lastFirestoreError = `Save Error: ${e.message}`;
    console.warn("Notice: Save tokens to Firestore notice:", e.message);
  }
};

const loadTokens = async () => {
  const tryLoad = async (db: any) => {
    const doc = await db.collection("config").doc("google_auth").get();
    if (doc.exists) {
      return doc.data()?.tokens || null;
    }
    return null;
  };

  try {
    try {
      const tokens = await tryLoad(getFirestoreInstance());
      if (tokens) return tokens;
    } catch (e: any) {
      if (e.code === 5 || e.message?.includes("NOT_FOUND") || e.message?.includes("not found")) {
        const tokens = await tryLoad(getFirestoreInstance(true));
        if (tokens) return tokens;
      }
    }
  } catch (e: any) {
    lastFirestoreError = `Load Error: ${e.message}`;
    console.warn("Notice: Load tokens from Firestore notice:", e.message);
  }
  return null;
};

const clearTokens = async () => {
  const tryClear = async (db: any) => {
    const docRef = db.collection("config").doc("google_auth");
    await docRef.delete();
  };

  try {
    try {
      await tryClear(getFirestoreInstance());
      lastFirestoreError = null;
    } catch (e: any) {
      if (e.code === 5 || e.message?.includes("NOT_FOUND") || e.message?.includes("not found")) {
        await tryClear(getFirestoreInstance(true));
        lastFirestoreError = null;
      }
    }
  } catch (e: any) {
    lastFirestoreError = `Clear Error: ${e.message}`;
    console.warn("Notice: Clear tokens from Firestore notice:", e.message);
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

// Debug routes early
app.get("/api/debug/auth", async (req, res) => {
  const firestoreTokens = await loadTokens();
  let sa_project_id = "none";
  try {
    if (process.env.FIREBASE_SERVICE_ACCOUNT) {
      sa_project_id = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT).project_id;
    }
  } catch (e) {}

  res.json({
    has_service_account: !!process.env.FIREBASE_SERVICE_ACCOUNT,
    sa_project_id: sa_project_id,
    config_project_id: firebaseConfig.projectId,
    firestore_database_id: firebaseConfig.firestoreDatabaseId,
    firestore_tokens_exist: !!firestoreTokens,
    cookie_tokens_exist: !!req.cookies.google_tokens,
    last_error: lastFirestoreError,
    env_vars: {
      GOOGLE_CLIENT_ID: !!process.env.GOOGLE_CLIENT_ID,
      GOOGLE_CLIENT_SECRET: !!process.env.GOOGLE_CLIENT_SECRET,
      APP_URL: process.env.APP_URL,
      FIREBASE_DATABASE_ID: !!process.env.FIREBASE_DATABASE_ID
    }
  });
});

app.get("/api/debug/test-db", async (req, res) => {
  try {
    const db = getFirestoreInstance();
    await db.collection("config").doc("test_connection").set({
      time: new Date().toISOString(),
      status: "ok"
    });
    res.json({ success: true, message: "Ghi dữ liệu thành công vào Database chính!" });
  } catch (e: any) {
    try {
      const dbFallback = getFirestoreInstance(true);
      await dbFallback.collection("config").doc("test_connection").set({
        time: new Date().toISOString(),
        status: "ok_fallback"
      });
      res.json({ success: true, message: "Ghi dữ liệu thành công vào Database mặc định (Fallback)!" });
    } catch (fallbackErr: any) {
      res.status(500).json({ 
        success: false, 
        primary_error: e.message,
        fallback_error: fallbackErr.message,
        code: fallbackErr.code
      });
    }
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

function serializeForFirestore(obj: any): any {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj !== 'object') return obj;

  if (Array.isArray(obj)) {
    if (obj.some(item => Array.isArray(item))) {
      return JSON.stringify(obj);
    }
    return obj.map(serializeForFirestore);
  }

  const result: Record<string, any> = {};
  for (const key of Object.keys(obj)) {
    result[key] = serializeForFirestore(obj[key]);
  }
  return result;
}

function deserializeFromFirestore(obj: any): any {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj === 'string') {
    const trimmed = obj.trim();
    if ((trimmed.startsWith('[') && trimmed.endsWith(']')) || (trimmed.startsWith('{') && trimmed.endsWith('}'))) {
      try {
        const parsed = JSON.parse(trimmed);
        return deserializeFromFirestore(parsed);
      } catch (e) {
        return obj;
      }
    }
    return obj;
  }
  if (typeof obj !== 'object') return obj;

  if (Array.isArray(obj)) {
    return obj.map(deserializeFromFirestore);
  }

  const result: Record<string, any> = {};
  for (const key of Object.keys(obj)) {
    result[key] = deserializeFromFirestore(obj[key]);
  }
  return result;
}

const SHIFTS_CYCLE = [
  ["N","C","O","K","O"],
  ["K","O","N","C","O"],
  ["C","O","K","O","N"],
  ["O","N","C","O","K"],
  ["O","K","O","N","C"]
];
const BASE_DATE_REF = new Date(2025, 9, 1); // 01/10/2025

function xacDinhCaHelper(ngay: Date, kip: number): string {
  if (!ngay || isNaN(ngay.getTime()) || kip < 1 || kip > 5) return 'O';
  const d1 = new Date(ngay.getFullYear(), ngay.getMonth(), ngay.getDate());
  const d2 = new Date(BASE_DATE_REF.getFullYear(), BASE_DATE_REF.getMonth(), BASE_DATE_REF.getDate());
  const diffDays = Math.round((d1.getTime() - d2.getTime()) / 86400000);
  const kipIdx = ((kip - 1) % 5 + 5) % 5;
  const cycleIdx = ((diffDays % 5) + 5) % 5;
  return SHIFTS_CYCLE[kipIdx][cycleIdx] || 'O';
}

function calculateLeaveDays(startDateStr?: string, endDateStr?: string, kip?: number | string): number {
  if (!startDateStr) return 1;
  
  const parseVN = (s: string) => {
    if (!s) return null;
    const str = String(s).trim();
    // Handle Excel serial date numbers (e.g. 45500)
    if (/^\d{5}(\.\d+)?$/.test(str)) {
      const serial = parseFloat(str);
      const utc_days = Math.floor(serial - 25569);
      const utc_value = utc_days * 86400;
      return new Date(utc_value * 1000);
    }

    const parts = str.split(/[\/\-\.]/);
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        // YYYY-MM-DD
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10) - 1;
        const d = parseInt(parts[2], 10);
        if (!isNaN(d) && !isNaN(m) && !isNaN(y)) return new Date(y, m, d);
      } else {
        // DD/MM/YYYY
        const d = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10) - 1;
        const y = parseInt(parts[2], 10);
        if (!isNaN(d) && !isNaN(m) && !isNaN(y)) return new Date(y, m, d);
      }
    }
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  };

  try {
    const start = parseVN(startDateStr);
    const end = endDateStr ? parseVN(endDateStr) : start;
    if (!start || !end || isNaN(start.getTime()) || isNaN(end.getTime())) return 1;

    // Sanity check: Leave requests year must be realistic (>= 2020)
    const currentYear = new Date().getFullYear();
    if (start.getFullYear() < 2020 || start.getFullYear() > currentYear + 2) {
      return 1;
    }

    let numKip = 0;
    if (typeof kip === 'number') numKip = kip;
    else if (typeof kip === 'string') {
      const match = kip.match(/\d+/);
      if (match) numKip = parseInt(match[0], 10);
    }

    if (numKip >= 1 && numKip <= 5) {
      let workShifts = 0;
      let cur = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const endClean = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      let maxDays = 60;
      while (cur <= endClean && maxDays > 0) {
        maxDays--;
        const shiftType = xacDinhCaHelper(cur, numKip);
        if (shiftType !== 'O') {
          workShifts++;
        }
        cur.setDate(cur.getDate() + 1);
      }
      return workShifts;
    }

    const diffTime = Math.abs(end.getTime() - start.getTime());
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    if (diffDays > 30) return 1;
    return diffDays > 0 ? diffDays : 1;
  } catch (e) {
    return 1;
  }
}

function calculateLeaveBalanceLogic(
  oldLeavesInput: number | string,
  newLeavesInput: number | string,
  usedLeavesInput: number | string,
  currentDateObj: Date = new Date()
) {
  const oldLeaves = Math.max(0, parseFloat(String(oldLeavesInput)) || 0);
  const newLeaves = Math.max(0, parseFloat(String(newLeavesInput)) || 12);
  const usedLeaves = Math.max(0, parseFloat(String(usedLeavesInput)) || 0);

  const month = currentDateObj.getMonth() + 1; // 1..12
  const isBeforeMarch31 = month <= 3;

  let remaining = 0;
  let note = "";

  if (isBeforeMarch31) {
    const totalEntitled = oldLeaves + newLeaves;
    remaining = Math.max(0, totalEntitled - usedLeaves);
    
    if (oldLeaves > 0) {
      if (usedLeaves < oldLeaves) {
        note = `Đang ưu tiên trừ phép năm cũ (còn ${oldLeaves - usedLeaves}/${oldLeaves} ngày năm cũ, hạn đến 31/03).`;
      } else {
        note = `Đã dùng hết ${oldLeaves} ngày phép năm cũ trong Q1. Số ngày nghỉ còn lại trừ vào phép năm mới.`;
      }
    } else {
      note = `Không có phép năm cũ. Trừ trực tiếp vào phép năm mới (${newLeaves} ngày).`;
    }
  } else {
    const usedDeductedFromNew = Math.max(0, usedLeaves - oldLeaves);
    remaining = Math.max(0, newLeaves - usedDeductedFromNew);

    if (oldLeaves > 0) {
      const expiredOldLeaves = Math.max(0, oldLeaves - usedLeaves);
      if (expiredOldLeaves > 0) {
        note = `Phép năm cũ (${expiredOldLeaves} ngày chưa dùng) đã tự động hết hạn từ 31/03. Đang tính phép năm mới (${newLeaves} ngày).`;
      } else {
        note = `Phép năm cũ đã dùng hết từ Q1. Chỉ còn tính phép năm mới.`;
      }
    } else {
      note = `Tính theo tiêu chuẩn phép năm hiện tại (${newLeaves} ngày).`;
    }
  }

  return {
    oldLeaves,
    newLeaves,
    usedLeaves,
    entitled: isBeforeMarch31 ? (oldLeaves + newLeaves) : newLeaves,
    remaining,
    isBeforeMarch31,
    note
  };
}

async function fetchLeaveBalanceHelper(name: string, customSheetId?: string): Promise<{ entitled: string; used: string; remaining: string; note?: string; oldLeaves?: string; newLeaves?: string } | null> {
  try {
    const oauth2Client = getOAuth2Client();
    let tokens = await loadTokens();
    if (!tokens) return null;

    oauth2Client.setCredentials(tokens);
    const sheets = google.sheets({ version: "v4", auth: oauth2Client });
    const spreadsheetId = customSheetId ? extractSpreadsheetId(customSheetId) : resolveSpreadsheetId();

    let targetSheetName = "Số ngày phép";
    try {
      const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
      const sheetNames = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];
      targetSheetName = sheetNames.find(n => n?.toLowerCase().trim() === "số ngày phép") || sheetNames[0] || "Số ngày phép";
    } catch (e) {}

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${targetSheetName}'!A1:I1000`
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return null;

    const targetName = String(name).trim();
    const targetNormalized = removeVietnameseAccents(targetName);

    let matchedRow: any[] | null = null;
    
    // First pass: exact match
    for (const row of rows) {
      if (!row || row.length === 0) continue;
      for (let colIdx = 0; colIdx < Math.min(row.length, 4); colIdx++) {
        const cellVal = String(row[colIdx] || '').trim().toLowerCase();
        if (cellVal === targetName.toLowerCase()) {
          matchedRow = row;
          break;
        }
      }
      if (matchedRow) break;
    }

    // Second pass: normalized match
    if (!matchedRow) {
      for (const row of rows) {
        if (!row || row.length === 0) continue;
        for (let colIdx = 0; colIdx < Math.min(row.length, 4); colIdx++) {
          const cellVal = String(row[colIdx] || '').trim();
          if (removeVietnameseAccents(cellVal) === targetNormalized) {
            matchedRow = row;
            break;
          }
        }
        if (matchedRow) break;
      }
    }

    // Dynamic header column mapping
    const headerRow = rows[0] || [];
    let oldLeavesCol = 4; // Column E
    let newLeavesCol = 5; // Column F (Phép năm mới được hưởng)
    let usedCol = 6;      // Column G (Phép đã nghỉ năm)
    let remainingCol = 7; // Column H (Phép còn lại hiện tại)

    for (let c = 0; c < headerRow.length; c++) {
      const hText = removeVietnameseAccents(String(headerRow[c] || '')).toLowerCase();
      if (hText.includes("nam cu") || hText.includes("phep nam cu")) {
        oldLeavesCol = c;
      } else if (hText.includes("nam moi") || hText.includes("duoc huong") || hText.includes("dinh muc")) {
        newLeavesCol = c;
      } else if (hText.includes("da nghi")) {
        usedCol = c;
      } else if (hText.includes("con lai")) {
        remainingCol = c;
      }
    }

    if (matchedRow) {
      const oldLeaves = oldLeavesCol >= 0 ? (parseFloat(String(matchedRow[oldLeavesCol] || '0').trim()) || 0) : 0;
      const newLeaves = parseFloat(String(matchedRow[newLeavesCol] || '12').trim()) || 12;
      const used = parseFloat(String(matchedRow[usedCol] || '0').trim()) || 0;
      const calc = calculateLeaveBalanceLogic(oldLeaves, newLeaves, used);
      const remaining = (matchedRow[remainingCol] !== undefined && String(matchedRow[remainingCol]).trim() !== '')
        ? String(matchedRow[remainingCol]).trim()
        : String(calc.remaining);

      return { 
        entitled: String(calc.entitled), 
        used: String(calc.usedLeaves), 
        remaining,
        note: calc.note,
        oldLeaves: String(calc.oldLeaves),
        newLeaves: String(calc.newLeaves)
      };
    }
  } catch (err: any) {
    console.error("fetchLeaveBalanceHelper error:", err.message);
  }
  return null;
}

// Signature Management Endpoints
app.get("/api/signatures", async (req, res) => {
  try {
    const db = getFirestoreInstance();
    const snapshot = await db.collection("signatures").get();
    const signatures: Record<string, string> = {};
    snapshot.forEach(doc => {
      signatures[doc.id] = doc.data().data;
    });
    res.json(signatures);
  } catch (e: any) {
    console.error("Error fetching signatures:", e);
    res.status(500).json({ error: e.message });
  }
});

// ==========================================
// WORKSHOP & AUTHENTICATION MANAGEMENT
// ==========================================

const DEFAULT_WORKSHOPS = [
  {
    id: "px_vanhanh",
    name: "Phân xưởng Vận hành Thủy điện",
    code: "PX-VH",
    description: "Phân xưởng trực ban vận hành và điều khiển thiết bị",
    staffData: [
      ["Trưởng ca",        "Nguyễn Tiến Danh",  "Tạ Văn Hà",          "Nguyễn Văn Trường", "Bùi Chí Thanh",      "Hoàng Tấn Hùng"],
      ["Trực TTĐK",        "Vũ Đức Cường",       "Lê Trí Dũng",        "Nguyễn Văn Toàn",   "Mai Xuân Sơn",       "Trần Trung Liêm"],
      ["Trực chính điện",  "Phan Văn Hùng",      "Ngô Xuân Đoàn",      "Nguyễn Yên Nam",    "Nguyễn Trung Chính", "Lê Hướng"],
      ["Trực phụ điện",    "Hoàng Ngọc Ân",      "Nguyễn Ngọc Hùng",   "Nguyễn Phỉ Được",   "Kiều Cao Khởi1",       "Trương Đình Thắng"],
      ["Trực chính máy",   "Nguyễn Tấn Phước",   "Lê Văn Dân",         "Phạm Văn Toàn",     "Phan Ngọc Hùng",     "Nguyễn Khắc Học"],
      ["Trực phụ máy",     "Phạm Văn Von",       "Thái Trần Hoàng Vũ", "Lê Văn Tích",       "Kiều Cao Khởi",          "Puih Thăn"],
      ["TC TBA 500 kV",    "Võ Quang Minh",      "Trần Hữu Thuận",     "Lê Thành Cao",      "Bùi Ngọc Thuận",     "Phạm Hồng Thắng"],
      ["Trực OPY",         "Ngô Xuân Vỹ",        "Trần Văn Thiên",     "Hoàng Văn Thăng",   "Trần Tiễn Quân",     "A Ran"],
      ["Trực CNN",         "Phạm Văn Mạnh",      "Hà Văn Chăn",        "Lê Thế Đàm",        "Cù Minh Trung",      "Nguyễn Vinh Quang"],
      ["Trưởng kíp",       "Nguyễn Lâm Tiến",    "Đỗ Văn Anh",         "Trịnh Xuân An",     "Tào Trọng Thi",      "Vũ Huy Hùng"],
      ["Trực chính GM",    "Nguyễn Hồng Quang",  "Trần Nhật Huy",      "Võ Thành Trung",    "Phùng Ngọc Tú",      "Nguyễn Thành Nguyên"],
      ["Trực phụ điện MR", "Nguyễn Khánh Toàn", "Lê Hoài Bảo",        "Lê Vũ Minh Trung",  "Phạm Đình Đức",      "Lê Trọng Toàn"],
      ["Trực phụ máy MR",  "Nguyễn Quang Minh",  "Rmah Thắng",         "Phạm Thanh Tùng",   "Nguyễn Văn Trung",   "Lê Trọng Toàn1"]
    ],
    config: {
      soVanBan: '123/PX-VH',
      ngayKy: '',
      nguoiKy: 'Nguyễn Văn Nghị',
      chucVuNguoiKy: 'Quản đốc Phân xưởng',
      zaloWebhookUrl: 'https://vhialy.dpdns.org/webhook/notify'
    },
    features: {
      enablePhanCa: true,
      enableDonNghiPhep: true,
      enableDoiCa: true,
      enableChuKySo: true,
      enableGoogleSheets: true,
      enableBaoComCa: true
    },
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  },
  {
    id: "px_suachua",
    name: "Phân xưởng Sửa chữa Cơ điện",
    code: "PX-SC",
    description: "Phân xưởng chịu trách nhiệm bảo trì, bảo dưỡng và sửa chữa thiết bị",
    staffData: [
      ["Trưởng ca Sửa chữa", "Lê Văn Tùng", "Trần Đình Bảo", "Phạm Văn Nam"],
      ["Kỹ thuật viên Điện", "Nguyễn Hoàng Việt", "Đỗ Văn Toàn", "Vũ Minh Tân"],
      ["Kỹ thuật viên Cơ", "Nguyễn Khắc Tuấn", "Phan Thanh Hải", "Lê Hữu Nghĩa"]
    ],
    config: {
      soVanBan: '45/PX-SC',
      ngayKy: '',
      nguoiKy: 'Trần Văn Tuyên',
      chucVuNguoiKy: 'Quản đốc PX Sửa chữa',
      zaloWebhookUrl: 'https://vhialy.dpdns.org/webhook/notify'
    },
    features: {
      enablePhanCa: true,
      enableDonNghiPhep: true,
      enableDoiCa: true,
      enableChuKySo: true,
      enableGoogleSheets: false,
      enableBaoComCa: true
    },
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  }
];

const DEFAULT_ACCOUNTS = [
  {
    id: "acc_admin",
    username: "admin",
    password: "admin123",
    fullName: "Quản trị viên Hệ thống (Super Admin)",
    role: "super_admin",
    workshopId: "all",
    createdAt: new Date().toISOString()
  },
  {
    id: "acc_admin_vh",
    username: "admin_vanhanh",
    password: "123456",
    fullName: "Admin Phân xưởng Vận hành",
    role: "workshop_admin",
    workshopId: "px_vanhanh",
    createdAt: new Date().toISOString()
  },
  {
    id: "acc_user_vh",
    username: "user_vanhanh",
    password: "123456",
    fullName: "Tài khoản PX Vận hành",
    role: "workshop_user",
    workshopId: "px_vanhanh",
    createdAt: new Date().toISOString()
  },
  {
    id: "acc_admin_sc",
    username: "admin_suachua",
    password: "123456",
    fullName: "Admin Phân xưởng Sửa chữa",
    role: "workshop_admin",
    workshopId: "px_suachua",
    createdAt: new Date().toISOString()
  },
  {
    id: "acc_user_sc",
    username: "user_suachua",
    password: "123456",
    fullName: "Tài khoản PX Sửa chữa",
    role: "workshop_user",
    workshopId: "px_suachua",
    createdAt: new Date().toISOString()
  }
];

const LOCAL_STORE_PATH = path.join(process.cwd(), "local_store.json");

let localStore: {
  workshops: any[];
  accounts: any[];
  appSettings: any;
  pendingLeaves: any[];
  signatures: Record<string, any>;
} = {
  workshops: [...DEFAULT_WORKSHOPS],
  accounts: [...DEFAULT_ACCOUNTS],
  appSettings: null,
  pendingLeaves: [],
  signatures: {}
};

try {
  if (fs.existsSync(LOCAL_STORE_PATH)) {
    const fileData = JSON.parse(fs.readFileSync(LOCAL_STORE_PATH, "utf-8"));
    if (fileData.workshops && Array.isArray(fileData.workshops)) localStore.workshops = fileData.workshops;
    if (fileData.accounts && Array.isArray(fileData.accounts)) localStore.accounts = fileData.accounts;
    if (fileData.appSettings) localStore.appSettings = fileData.appSettings;
    if (fileData.pendingLeaves && Array.isArray(fileData.pendingLeaves)) localStore.pendingLeaves = fileData.pendingLeaves;
    if (fileData.signatures) localStore.signatures = fileData.signatures;
  }
} catch (err) {
  console.warn("Notice: Reading local_store.json fallback notice:", err);
}

function persistLocalStore() {
  try {
    fs.writeFileSync(LOCAL_STORE_PATH, JSON.stringify(localStore, null, 2), "utf-8");
  } catch (err) {
    // Ignore if environment is read-only
  }
}

let hasSeededWorkshops = false;
let hasSeededAccounts = false;

async function ensureSeedData() {
  try {
    const db = getFirestoreInstance();
    const seedDoc = await db.collection("system").doc("seed_status").get().catch(() => null);
    if (seedDoc && seedDoc.exists) {
      hasSeededWorkshops = true;
      hasSeededAccounts = true;
      return;
    }
    
    // Seed default workshops only if workshops collection is empty
    if (!hasSeededWorkshops) {
      const wsSnap = await db.collection("workshops").limit(1).get();
      if (wsSnap.empty) {
        for (const ws of DEFAULT_WORKSHOPS) {
          await db.collection("workshops").doc(ws.id).set(serializeForFirestore(ws));
        }
      }
      hasSeededWorkshops = true;
    }

    // Seed default accounts only if accounts collection is empty
    if (!hasSeededAccounts) {
      const accSnap = await db.collection("accounts").limit(1).get();
      if (accSnap.empty) {
        for (const acc of DEFAULT_ACCOUNTS) {
          await db.collection("accounts").doc(acc.id).set(acc);
        }
      }
      hasSeededAccounts = true;
    }

    await db.collection("system").doc("seed_status").set({ seededAt: new Date().toISOString() }).catch(() => null);
  } catch (e: any) {
    console.warn("Firestore seed data check notice (using local fallback store):", e.message || e);
  }
}

ensureSeedData().catch((err) => console.warn("Seed data notice:", err?.message || err));

// Track failed login attempts and temporary account locks (5 fails -> 15 min lock)
const failedLoginAttempts: Record<string, { count: number; lockUntil: number }> = {};

app.post("/api/auth/login", express.json(), async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: "Tên đăng nhập và mật khẩu là bắt buộc." });
  }

  const cleanUsername = String(username).trim().toLowerCase();
  const cleanPassword = String(password).trim();
  const now = Date.now();

  // Check if account is locked out
  const userLock = failedLoginAttempts[cleanUsername];
  if (userLock && userLock.lockUntil > now) {
    const remainingMs = userLock.lockUntil - now;
    const remainingMins = Math.ceil(remainingMs / 60000);
    return res.status(429).json({
      error: `Tài khoản tạm thời bị khóa do nhập sai mật khẩu 5 lần liên tiếp. Vui lòng thử lại sau ${remainingMins} phút.`
    });
  }

  await ensureSeedData();

  let userDoc: any = null;

  try {
    const db = getFirestoreInstance();
    const snapshot = await db.collection("accounts").where("username", "==", cleanUsername).get();
    if (!snapshot.empty) {
      snapshot.forEach(doc => {
        const data = doc.data();
        if (data.password === cleanPassword) {
          userDoc = { id: doc.id, ...data };
        }
      });
    }
  } catch (dbErr: any) {
    console.warn("Firestore accounts search notice:", dbErr.message);
  }

  // Fallback to localStore / DEFAULT_ACCOUNTS
  if (!userDoc) {
    const localAcc = localStore.accounts.find(
      acc => acc.username.toLowerCase() === cleanUsername && acc.password === cleanPassword
    );
    if (localAcc) {
      userDoc = { ...localAcc };
    } else {
      const defaultAcc = DEFAULT_ACCOUNTS.find(
        acc => acc.username.toLowerCase() === cleanUsername && acc.password === cleanPassword
      );
      if (defaultAcc) {
        userDoc = { ...defaultAcc };
      }
    }
  }

  if (!userDoc) {
    const currentAttempt = failedLoginAttempts[cleanUsername] || { count: 0, lockUntil: 0 };
    if (currentAttempt.lockUntil && currentAttempt.lockUntil <= now) {
      currentAttempt.count = 0;
      currentAttempt.lockUntil = 0;
    }

    currentAttempt.count += 1;

    if (currentAttempt.count >= 5) {
      currentAttempt.lockUntil = now + 15 * 60 * 1000; // 15 minutes lockout
      failedLoginAttempts[cleanUsername] = currentAttempt;
      return res.status(429).json({
        error: "Bạn đã nhập sai mật khẩu 5 lần liên tiếp. Tài khoản tạm thời bị khóa trong 15 phút."
      });
    }

    failedLoginAttempts[cleanUsername] = currentAttempt;
    const remaining = 5 - currentAttempt.count;
    return res.status(401).json({
      error: `Tên đăng nhập hoặc mật khẩu không chính xác (còn ${remaining} lần thử trước khi khóa 15 phút).`
    });
  }

  // Successful login -> clear failed attempts
  delete failedLoginAttempts[cleanUsername];

  const { password: _, ...userWithoutPassword } = userDoc;

  const isHttps = req.secure || req.headers["x-forwarded-proto"] === "https";
  res.cookie("auth_user", JSON.stringify(userWithoutPassword), {
    httpOnly: true,
    secure: isHttps,
    sameSite: isHttps ? "none" : "lax",
    maxAge: 7 * 24 * 60 * 60 * 1000
  });

  res.json({
    success: true,
    user: userWithoutPassword
  });
});

app.get("/api/auth/me", async (req, res) => {
  let cookieUser = req.cookies.auth_user;

  if (!cookieUser && req.headers["x-auth-user"]) {
    try {
      const rawHeader = req.headers["x-auth-user"] as string;
      cookieUser = rawHeader.startsWith("%") ? decodeURIComponent(rawHeader) : rawHeader;
    } catch (e) {
      cookieUser = req.headers["x-auth-user"];
    }
  }

  if (!cookieUser) {
    return res.json({ authenticated: false, user: null });
  }
  try {
    const user = typeof cookieUser === 'string' ? JSON.parse(cookieUser) : cookieUser;
    res.json({ authenticated: true, user });
  } catch (e) {
    res.json({ authenticated: false, user: null });
  }
});

app.post("/api/auth/logout", (req, res) => {
  res.clearCookie("auth_user");
  res.json({ success: true });
});

app.post("/api/auth/change-password", express.json(), async (req, res) => {
  let cookieUser = req.cookies.auth_user;
  if (!cookieUser && req.headers["x-auth-user"]) {
    try {
      const rawHeader = req.headers["x-auth-user"] as string;
      cookieUser = rawHeader.startsWith("%") ? decodeURIComponent(rawHeader) : rawHeader;
    } catch (e) {
      cookieUser = req.headers["x-auth-user"];
    }
  }

  if (!cookieUser) {
    return res.status(401).json({ error: "Bạn chưa đăng nhập." });
  }

  let currentUser: any = null;
  try {
    currentUser = typeof cookieUser === 'string' ? JSON.parse(cookieUser) : cookieUser;
  } catch (e) {
    return res.status(401).json({ error: "Phiên đăng nhập không hợp lệ." });
  }

  const { oldPassword, newPassword } = req.body;
  if (!oldPassword || !newPassword) {
    return res.status(400).json({ error: "Vui lòng nhập đầy đủ mật khẩu cũ và mật khẩu mới." });
  }

  const cleanOld = String(oldPassword).trim();
  const cleanNew = String(newPassword).trim();

  if (cleanNew.length < 4) {
    return res.status(400).json({ error: "Mật khẩu mới phải có ít nhất 4 ký tự." });
  }

  let userAccountDoc: any = null;
  const usernameClean = currentUser.username.toLowerCase();

  try {
    const db = getFirestoreInstance();
    const snapshot = await db.collection("accounts").where("username", "==", usernameClean).get();
    if (!snapshot.empty) {
      snapshot.forEach(doc => {
        const data = doc.data();
        if (data.password === cleanOld) {
          userAccountDoc = { id: doc.id, ...data };
        }
      });
    }
  } catch (e: any) {
    console.warn("Firestore find account notice:", e?.message);
  }

  if (!userAccountDoc) {
    const localAcc = localStore.accounts.find(
      acc => acc.username.toLowerCase() === usernameClean && acc.password === cleanOld
    );
    if (localAcc) {
      userAccountDoc = { ...localAcc };
    } else {
      const defaultAcc = DEFAULT_ACCOUNTS.find(
        acc => acc.username.toLowerCase() === usernameClean && acc.password === cleanOld
      );
      if (defaultAcc) {
        userAccountDoc = { ...defaultAcc };
      }
    }
  }

  if (!userAccountDoc) {
    return res.status(400).json({ error: "Mật khẩu cũ không chính xác." });
  }

  userAccountDoc.password = cleanNew;
  userAccountDoc.updatedAt = new Date().toISOString();

  const accIdx = localStore.accounts.findIndex(a => a.username.toLowerCase() === usernameClean || a.id === userAccountDoc.id);
  if (accIdx >= 0) {
    localStore.accounts[accIdx].password = cleanNew;
  } else {
    localStore.accounts.push(userAccountDoc);
  }
  persistLocalStore();

  try {
    const db = getFirestoreInstance();
    await db.collection("accounts").doc(userAccountDoc.id).set({
      password: cleanNew,
      updatedAt: userAccountDoc.updatedAt
    }, { merge: true });
  } catch (e: any) {
    console.warn("Firestore password update notice:", e.message);
  }

  res.json({ success: true, message: "Đổi mật khẩu thành công!" });
});

app.get("/api/workshops", async (req, res) => {
  try {
    const db = getFirestoreInstance();
    await ensureSeedData();
    
    const snapshot = await db.collection("workshops").get();
    const workshops: any[] = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      const deserialized = deserializeFromFirestore(data);
      workshops.push({
        id: doc.id,
        ...deserialized
      });
    });
    localStore.workshops = workshops;
    persistLocalStore();
    return res.json(workshops);
  } catch (e: any) {
    console.warn("Firestore get workshops notice (using localStore fallback):", e.message);
  }
  res.json(localStore.workshops || []);
});

app.post("/api/workshops", express.json({ limit: '10mb' }), async (req, res) => {
  const { id, name, code, description, staffData, config, features } = req.body;
  if (!name) {
    return res.status(400).json({ error: "Tên phân xưởng là bắt buộc" });
  }

  const wsId = id || ("px_" + Date.now());
  const now = new Date().toISOString();

  const workshopDoc = {
    id: wsId,
    name: name.trim(),
    code: (code || name.substring(0, 5).toUpperCase()).trim(),
    description: description || "",
    staffData: staffData || [],
    config: config || {
      soVanBan: '123/PX',
      ngayKy: '',
      nguoiKy: 'Lãnh đạo Phân xưởng',
      chucVuNguoiKy: 'Quản đốc Phân xưởng',
      zaloWebhookUrl: 'https://vhialy.dpdns.org/webhook/notify'
    },
    features: features || {
      enablePhanCa: true,
      enableDonNghiPhep: true,
      enableDoiCa: true,
      enableChuKySo: true,
      enableGoogleSheets: true,
      enableBaoComCa: true
    },
    updatedAt: now,
    createdAt: req.body.createdAt || now
  };

  const wsIdx = localStore.workshops.findIndex(w => w.id === wsId);
  if (wsIdx >= 0) {
    localStore.workshops[wsIdx] = { ...localStore.workshops[wsIdx], ...workshopDoc };
  } else {
    localStore.workshops.push(workshopDoc);
  }
  persistLocalStore();

  try {
    const db = getFirestoreInstance();
    const docToSave = serializeForFirestore(workshopDoc);
    await db.collection("workshops").doc(wsId).set(docToSave, { merge: true });
  } catch (e: any) {
    console.warn("Firestore save workshop notice (saved in localStore fallback):", e.message);
  }

  res.json({
    success: true,
    workshop: workshopDoc
  });
});

app.delete("/api/workshops/:id", async (req, res) => {
  const { id } = req.params;
  localStore.workshops = localStore.workshops.filter(w => w.id !== id);
  localStore.accounts = localStore.accounts.filter(a => a.workshopId !== id);
  persistLocalStore();

  try {
    const db = getFirestoreInstance();
    await db.collection("workshops").doc(id).delete();

    const accSnap = await db.collection("accounts").where("workshopId", "==", id).get();
    if (!accSnap.empty) {
      const batch = db.batch();
      accSnap.forEach(doc => {
        batch.delete(doc.ref);
      });
      await batch.commit();
    }
  } catch (e: any) {
    console.warn("Firestore delete workshop notice:", e.message);
  }
  res.json({ success: true });
});

app.get("/api/accounts", async (req, res) => {
  try {
    const db = getFirestoreInstance();
    await ensureSeedData();
    
    const snapshot = await db.collection("accounts").get();
    const accounts: any[] = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      accounts.push({
        id: doc.id,
        username: data.username,
        fullName: data.fullName,
        role: data.role,
        workshopId: data.workshopId,
        createdAt: data.createdAt
      });
    });
    localStore.accounts = accounts;
    persistLocalStore();
    return res.json(accounts);
  } catch (e: any) {
    console.warn("Firestore get accounts notice (using localStore fallback):", e.message);
  }
  res.json(localStore.accounts || []);
});

app.post("/api/accounts", express.json(), async (req, res) => {
  const { id, username, password, fullName, role, workshopId } = req.body;
  if (!username || !fullName || !role || !workshopId) {
    return res.status(400).json({ error: "Vui lòng nhập đủ các thông tin bắt buộc." });
  }

  const cleanUsername = username.trim().toLowerCase();
  const accId = id || ("acc_" + Date.now());
  const now = new Date().toISOString();

  const isDuplicateLocal = localStore.accounts.some(a => a.username.toLowerCase() === cleanUsername && a.id !== accId);
  if (isDuplicateLocal) {
    return res.status(400).json({ error: "Tên đăng nhập đã được sử dụng." });
  }

  const existingLocal = localStore.accounts.find(a => a.id === accId);
  const accountData: any = {
    id: accId,
    username: cleanUsername,
    fullName: fullName.trim(),
    role,
    workshopId,
    updatedAt: now,
    createdAt: existingLocal?.createdAt || now
  };

  if (password && password.trim()) {
    accountData.password = password.trim();
  } else if (existingLocal?.password) {
    accountData.password = existingLocal.password;
  } else {
    accountData.password = "123456";
  }

  const accIdx = localStore.accounts.findIndex(a => a.id === accId);
  if (accIdx >= 0) {
    localStore.accounts[accIdx] = accountData;
  } else {
    localStore.accounts.push(accountData);
  }
  persistLocalStore();

  try {
    const db = getFirestoreInstance();
    const docRef = db.collection("accounts").doc(accId);
    await docRef.set(accountData, { merge: true });
  } catch (e: any) {
    console.warn("Firestore save account notice (saved in localStore fallback):", e.message);
  }

  const { password: _, ...accountPublic } = accountData;
  res.json({
    success: true,
    account: accountPublic
  });
});

app.delete("/api/accounts/:id", async (req, res) => {
  const { id } = req.params;
  localStore.accounts = localStore.accounts.filter(a => a.id !== id && a.username.toLowerCase() !== id.toLowerCase());
  persistLocalStore();

  try {
    const db = getFirestoreInstance();
    await db.collection("accounts").doc(id).delete();

    const querySnap = await db.collection("accounts").where("username", "==", id.toLowerCase()).get();
    if (!querySnap.empty) {
      const batch = db.batch();
      querySnap.forEach(doc => {
        batch.delete(doc.ref);
      });
      await batch.commit();
    }
  } catch (e: any) {
    console.warn("Firestore delete account notice:", e.message);
  }
  res.json({ success: true });
});

app.get("/api/signatures", async (req, res) => {
  try {
    const db = getFirestoreInstance();
    const snapshot = await db.collection("signatures").get();
    const result: Record<string, string> = {};
    snapshot.forEach(doc => {
      const data = doc.data();
      if (data.name && data.data) {
        result[data.name] = data.data;
      }
    });
    if (Object.keys(result).length > 0) {
      localStore.signatures = result;
      persistLocalStore();
      return res.json(result);
    }
  } catch (e: any) {
    console.warn("Firestore signatures read notice:", e.message);
  }
  res.json(localStore.signatures || {});
});

app.post("/api/signatures", express.json({ limit: '10mb' }), async (req, res) => {
  const { name, data } = req.body;
  if (!name || !data) {
    return res.status(400).json({ error: "Name and data are required" });
  }
  localStore.signatures[name] = data;
  persistLocalStore();

  try {
    const db = getFirestoreInstance();
    await db.collection("signatures").doc(name).set({
      name,
      data,
      updatedAt: new Date().toISOString()
    });
  } catch (e: any) {
    console.warn("Firestore signature save notice:", e.message);
  }
  res.json({ success: true });
});

app.post("/api/signatures/batch", express.json({ limit: '50mb' }), async (req, res) => {
  const { items } = req.body;
  if (!Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: "Items array is required" });
  }

  for (const item of items) {
    if (item.name && item.data) {
      localStore.signatures[item.name] = item.data;
      try {
        const db = getFirestoreInstance();
        await db.collection("signatures").doc(item.name).set({
          name: item.name,
          data: item.data,
          updatedAt: new Date().toISOString()
        });
      } catch (e: any) {
        console.warn("Firestore batch signature save notice:", e.message);
      }
    }
  }
  persistLocalStore();
  res.json({ success: true, count: items.length });
});

// App Settings Endpoints (Staff List & Config)
app.get("/api/app-settings", async (req, res) => {
  try {
    const db = getFirestoreInstance();
    const doc = await db.collection("config").doc("app_settings").get();
    if (doc.exists) {
      const data = doc.data();
      if (data) {
        const deserialized = deserializeFromFirestore(data);
        if (!Array.isArray(deserialized.staffData)) {
          deserialized.staffData = [];
        } else {
          // Perform server-side migration for 'Trực phụ cơ MR' -> 'Trực phụ máy MR'
          let hasMigration = false;
          const migrated = deserialized.staffData.map((row: any) => {
            if (Array.isArray(row) && row[0] === 'Trực phụ cơ MR') {
              hasMigration = true;
              return ['Trực phụ máy MR', ...row.slice(1)];
            }
            return row;
          });
          if (hasMigration) {
            deserialized.staffData = migrated;
            try {
              await db.collection("config").doc("app_settings").set(serializeForFirestore({
                staffData: migrated
              }), { merge: true });
            } catch (saveErr) {}
          }
        }

        // Migrate Zalo Webhook URL
        if (
          deserialized.config && 
          deserialized.config.zaloWebhookUrl && 
          (
            deserialized.config.zaloWebhookUrl.includes("cookies-blue-pen-bikini.trycloudflare.com") || 
            deserialized.config.zaloWebhookUrl.includes("specialists-intro-exterior-advocacy.trycloudflare.com") ||
            deserialized.config.zaloWebhookUrl.includes("committed-intellectual-lunch-clone.trycloudflare.com")
          )
        ) {
          deserialized.config.zaloWebhookUrl = "https://vhialy.dpdns.org/webhook/notify";
          try {
            await db.collection("config").doc("app_settings").set(serializeForFirestore({
              config: deserialized.config
            }), { merge: true });
          } catch (saveErr) {}
        }

        localStore.appSettings = deserialized;
        persistLocalStore();
        return res.json(deserialized);
      }
    }
  } catch (e: any) {
    console.warn("Firestore get app-settings notice (using localStore fallback):", e.message);
  }

  if (localStore.appSettings) {
    return res.json(localStore.appSettings);
  }

  const defaultSettings = {
    staffData: DEFAULT_WORKSHOPS[0]?.staffData || [],
    config: DEFAULT_WORKSHOPS[0]?.config || {}
  };
  res.json(defaultSettings);
});

app.post("/api/app-settings", express.json({ limit: '5mb' }), async (req, res) => {
  const { staffData, config } = req.body;

  localStore.appSettings = { staffData, config };
  persistLocalStore();

  try {
    const db = getFirestoreInstance();
    const docToSave = serializeForFirestore({
      staffData,
      config,
      updatedAt: new Date().toISOString()
    });
    
    await db.collection("config").doc("app_settings").set(docToSave, { merge: true });
  } catch (e: any) {
    console.warn("Firestore save app-settings notice (saved in localStore fallback):", e.message);
  }

  res.json({ success: true });
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
  let firestoreTokens = await loadTokens();
  
  // Auto-sync: If cookie has tokens but Firestore doesn't, try to save them
  if (cookieTokensStr && !firestoreTokens) {
    console.log("Cookie exists but Firestore is empty. Attempting auto-sync...");
    try {
      const tokens = JSON.parse(cookieTokensStr);
      await saveTokens(tokens);
      firestoreTokens = await loadTokens();
    } catch (e) {
      console.error("Auto-sync failed:", e);
    }
  }
  
  res.json({ 
    authenticated: !!cookieTokensStr || !!firestoreTokens,
    source: cookieTokensStr ? "cookie" : (firestoreTokens ? "firestore" : "none"),
    sync_attempted: cookieTokensStr && !firestoreTokens
  });
});

const DEFAULT_LEAVE_SPREADSHEET_ID = '1m8B-CVJEyzt5KhJ-I_1hFTvWczz1jl7PPm_7PNScYYQ';
const LEAVE_SPREADSHEET_ID = DEFAULT_LEAVE_SPREADSHEET_ID;

function extractSpreadsheetId(input?: string): string {
  if (!input || typeof input !== 'string' || !input.trim()) {
    return DEFAULT_LEAVE_SPREADSHEET_ID;
  }
  const trimmed = input.trim();
  const match = trimmed.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    return match[1];
  }
  if (/^[a-zA-Z0-9-_]{20,}$/.test(trimmed)) {
    return trimmed;
  }
  return DEFAULT_LEAVE_SPREADSHEET_ID;
}

function resolveSpreadsheetId(req?: any): string {
  if (!req) return DEFAULT_LEAVE_SPREADSHEET_ID;
  const candidate = 
    req.body?.spreadsheetId ||
    req.body?.googleSheetUrl ||
    req.body?.leaveSpreadsheetUrl ||
    req.query?.spreadsheetId ||
    req.query?.googleSheetUrl ||
    req.query?.leaveSpreadsheetUrl;

  if (candidate && typeof candidate === 'string' && candidate.trim()) {
    return extractSpreadsheetId(candidate);
  }

  if (localStore?.appSettings?.config?.googleSheetUrl) {
    return extractSpreadsheetId(localStore.appSettings.config.googleSheetUrl);
  }
  if (localStore?.appSettings?.config?.leaveSpreadsheetUrl) {
    return extractSpreadsheetId(localStore.appSettings.config.leaveSpreadsheetUrl);
  }

  return DEFAULT_LEAVE_SPREADSHEET_ID;
}

async function updateAnnualLeaveUsedDays(sheets: any, spreadsheetId: string) {
  try {
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetObjects = spreadsheet.data.sheets || [];
    const sheetTitles = sheetObjects.map((s: any) => s.properties?.title || "");

    const annualLeaveSheetName = sheetTitles.find(t => t.toLowerCase().trim() === "số ngày phép") || "Số ngày phép";

    // Find all workshop leave sheets
    const leaveSheets = sheetTitles.filter(t => t.toLowerCase().startsWith("danh_sach_nghi"));

    const usedDaysMap = new Map<string, number>();

    for (const sheetName of leaveSheets) {
      try {
        const res = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `'${sheetName}'!A1:M1000`
        });
        const allRows = res.data.values || [];
        if (allRows.length <= 1) continue;

        const headerRow = allRows[0] || [];
        let nameCol = 1;
        let kipCol = 4;
        let startCol = 5;
        let endCol = 6;
        let numDaysCol = -1;
        let statusCol = 10;

        for (let c = 0; c < headerRow.length; c++) {
          const hText = removeVietnameseAccents(String(headerRow[c] || '')).toLowerCase().trim();
          if (hText.includes("ho va ten") || hText.includes("ho ten") || hText === "ten" || hText.includes("nhan su")) {
            nameCol = c;
          } else if (hText.includes("kip") || hText === "kip") {
            kipCol = c;
          } else if (hText.includes("so ca") || hText.includes("so ngay nghi") || hText.includes("so ca nghi") || hText === "so ngay") {
            numDaysCol = c;
          } else if (hText.includes("tu ngay") || hText.includes("ngay bat dau") || hText === "tu") {
            startCol = c;
          } else if (hText.includes("den ngay") || hText.includes("ngay ket thuc") || hText === "den") {
            endCol = c;
          } else if (hText.includes("trang thai") || hText.includes("tinh trang")) {
            statusCol = c;
          }
        }

        const dataRows = allRows.slice(1);
        for (const r of dataRows) {
          if (!r || r.length <= nameCol) continue;
          const personName = (r[nameCol] || '').trim();
          if (!personName) continue;

          const status = statusCol >= 0 && statusCol < r.length ? (r[statusCol] || '').trim().toLowerCase() : '';
          if (status === 'đã hủy' || status === 'từ chối' || status === 'da huy' || status === 'tu choi') continue;

          let days = 1;
          if (numDaysCol >= 0 && numDaysCol < r.length && r[numDaysCol] && !isNaN(parseFloat(r[numDaysCol]))) {
            days = Math.max(0, parseFloat(r[numDaysCol]));
          } else {
            const startDate = startCol >= 0 && startCol < r.length ? r[startCol] : '';
            const endDate = endCol >= 0 && endCol < r.length ? r[endCol] : startDate;
            const kipVal = kipCol >= 0 && kipCol < r.length ? r[kipCol] : '';
            days = calculateLeaveDays(startDate, endDate, kipVal);
          }

          const currentTotal = usedDaysMap.get(personName.toLowerCase()) || 0;
          usedDaysMap.set(personName.toLowerCase(), currentTotal + days);
        }
      } catch (e) {}
    }

    // Now read sheet "Số ngày phép"
    const annualRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${annualLeaveSheetName}'!A1:I500`
    });
    const annualRows = annualRes.data.values || [];
    if (annualRows.length <= 1) return;

    const updates: any[] = [];

    for (let i = 1; i < annualRows.length; i++) {
      const row = annualRows[i];
      if (!row || !row[1]) continue;
      const personName = String(row[1]).trim();
      const normName = personName.toLowerCase();
      const normAccents = removeVietnameseAccents(normName);

      // Find matching used days
      let usedDays = usedDaysMap.get(normName);
      if (usedDays === undefined) {
        for (const [key, val] of usedDaysMap.entries()) {
          if (removeVietnameseAccents(key) === normAccents) {
            usedDays = val;
            break;
          }
        }
      }
      if (usedDays === undefined) usedDays = 0;

      const rowIndex = i + 1; // 1-indexed row in Sheets
      const oldLeaves = parseFloat(String(row[4] || '0')) || 0;
      const newLeaves = parseFloat(String(row[5] || '12')) || 12;

      const calc = calculateLeaveBalanceLogic(oldLeaves, newLeaves, usedDays);

      const formulaStr = `=IF(MONTH(TODAY())<=3; MAX(0; E${rowIndex} + F${rowIndex} - G${rowIndex}); MAX(0; F${rowIndex} - MAX(0; G${rowIndex} - E${rowIndex})))`;
      
      updates.push({
        range: `'${annualLeaveSheetName}'!G${rowIndex}:I${rowIndex}`,
        values: [[
          usedDays,
          formulaStr,
          calc.note
        ]]
      });
    }

    if (updates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data: updates
        }
      });
      console.log(`Updated used leave days for ${updates.length} staff members in '${annualLeaveSheetName}'.`);
    }
  } catch (err: any) {
    console.error("updateAnnualLeaveUsedDays error:", err.message);
  }
}

async function ensureAnnualLeaveSheetExists(sheets: any, spreadsheetId: string, staffData?: string[][]): Promise<string> {
  const targetSheetName = "Số ngày phép";
  const headerValues = [
    "STT", 
    "Họ và tên", 
    "Chức danh", 
    "Kíp", 
    "Phép năm cũ còn lại", 
    "Phép năm mới được hưởng", 
    "Phép đã nghỉ năm nay", 
    "Phép còn lại hiện tại", 
    "Ghi chú"
  ];

  try {
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetObjects = spreadsheet.data.sheets || [];
    const sheetNames = sheetObjects.map((s: any) => s.properties?.title || "");

    const existingName = sheetNames.find(n => n.toLowerCase().trim() === "số ngày phép");

    if (!existingName) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: targetSheetName
                }
              }
            }
          ]
        }
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${targetSheetName}'!A1:I1`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [headerValues]
        }
      });
      console.log(`Created sheet '${targetSheetName}' and written 9-column headers.`);
    } else {
      // Ensure header has all 9 columns
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${existingName}'!A1:I1`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: [headerValues]
        }
      });
    }

    const actualSheetName = existingName || targetSheetName;

    if (staffData && Array.isArray(staffData) && staffData.length > 0) {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `'${actualSheetName}'!A1:I500`
      });
      const rows = response.data.values || [];
      const existingNames = new Set<string>();
      
      for (let i = 1; i < rows.length; i++) {
        if (rows[i] && rows[i][1]) {
          existingNames.add(rows[i][1].toString().trim().toLowerCase());
        }
      }

      const rowsToAppend: any[][] = [];
      let currentStt = rows.length > 1 ? rows.length : 1;

      for (const row of staffData) {
        if (!Array.isArray(row) || row.length === 0) continue;
        const title = (row[0] || '').trim();
        for (let k = 1; k <= 5; k++) {
          const personName = row[k] ? row[k].trim() : '';
          if (personName && !existingNames.has(personName.toLowerCase())) {
            existingNames.add(personName.toLowerCase());
            const rowIndex = currentStt + 1; // Row index in Google Sheets
            const formulaStr = `=IF(MONTH(TODAY())<=3; MAX(0; E${rowIndex} + F${rowIndex} - G${rowIndex}); MAX(0; F${rowIndex} - MAX(0; G${rowIndex} - E${rowIndex})))`;
            rowsToAppend.push([
              currentStt++,
              personName,
              title,
              `Kíp ${k}`,
              0,  // Phép năm cũ còn lại
              12, // Phép năm mới được hưởng
              0,  // Phép đã nghỉ năm nay
              formulaStr, // Phép còn lại hiện tại
              "Hạn phép năm cũ: hết ngày 31/03"
            ]);
          }
        }
      }

      if (rowsToAppend.length > 0) {
        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `'${actualSheetName}'!A:I`,
          valueInputOption: "USER_ENTERED",
          requestBody: {
            values: rowsToAppend
          }
        });
        console.log(`Appended ${rowsToAppend.length} staff members to '${actualSheetName}'.`);
      }
    }

    // Auto update used leave days from all leave requests
    updateAnnualLeaveUsedDays(sheets, spreadsheetId).catch(() => {});

    return actualSheetName;
  } catch (e: any) {
    console.error(`Error in ensureAnnualLeaveSheetExists:`, e.message);
    return targetSheetName;
  }
}

function formatWorkshopSheetTitle(workshopId?: string): string {
  if (!workshopId || workshopId === 'all') {
    return "DANH_SACH_NGHI";
  }
  const cleanId = workshopId.trim().replace(/[^a-zA-Z0-9_]/g, '_');
  return `DANH_SACH_NGHI_${cleanId}`;
}

async function getOrEnsureLeaveSheetTitle(sheets: any, spreadsheetId: string, workshopId?: string): Promise<string> {
  const desiredTitle = formatWorkshopSheetTitle(workshopId);
  try {
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetObjects = spreadsheet.data.sheets || [];
    const sheetNames = sheetObjects.map((s: any) => s.properties?.title || "");

    // 1. Exact match
    if (sheetNames.includes(desiredTitle)) {
      return desiredTitle;
    }

    // 2. Case-insensitive, substring, or alias match
    if (workshopId && workshopId !== 'all') {
      const targetLower = workshopId.trim().toLowerCase();
      const existingMatch = sheetNames.find(name => {
        const nLower = name.toLowerCase();
        if (!nLower.startsWith("danh_sach_nghi")) return false;
        const sub = nLower.replace("danh_sach_nghi_", "").replace("danh_sach_nghi", "");
        if (!sub) return false;
        return nLower.includes(targetLower) || 
               targetLower.includes(sub) || 
               (targetLower.includes("vanhanh") && sub.includes("vanhanh")) ||
               (targetLower.includes("vh") && sub.includes("vanhanh")) ||
               (targetLower.includes("suachua") && sub.includes("suachua")) ||
               (targetLower.includes("sc") && sub.includes("suachua"));
      });
      if (existingMatch) return existingMatch;
    } else {
      const defaultMatch = sheetNames.find(name => name.toLowerCase().trim() === "danh_sach_nghi");
      if (defaultMatch) return defaultMatch;
    }

    // 3. If not found, create new sheet tab with desiredTitle
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: desiredTitle
              }
            }
          }
        ]
      }
    });

    // Append headers
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${desiredTitle}'!A1:M1`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[
          "Mã đơn", 
          "Họ và tên", 
          "Năm sinh", 
          "Chức danh", 
          "Kíp", 
          "Từ ngày", 
          "Đến ngày", 
          "Lý do", 
          "Điện thoại", 
          "Địa điểm", 
          "Trạng thái", 
          "Ngày tạo",
          "Phân xưởng"
        ]]
      }
    });
    console.log(`Created leave sheet '${desiredTitle}' and headers.`);
    return desiredTitle;
  } catch (e: any) {
    console.error(`Error ensuring leave sheet for ${workshopId}:`, e.message);
    return desiredTitle;
  }
}

async function ensureLeaveSheetExists(sheets: any, spreadsheetId: string, workshopId?: string) {
  return await getOrEnsureLeaveSheetTitle(sheets, spreadsheetId, workshopId);
}

async function syncPendingLeavesToSheets(sheets: any, spreadsheetId: string) {
  try {
    const pendingMap = new Map<string, any>();

    // 1. Collect from localStore
    if (localStore.pendingLeaves && Array.isArray(localStore.pendingLeaves)) {
      localStore.pendingLeaves.forEach((item: any) => {
        if (item && item.id) {
          pendingMap.set(item.id, item);
        }
      });
    }

    // 2. Collect from Firestore
    try {
      const db = getFirestoreInstance();
      const snapshot = await db.collection("pending_leaves").get();
      snapshot.forEach(doc => {
        const item = doc.data();
        pendingMap.set(doc.id, { id: doc.id, ...item });
      });
    } catch (dbErr: any) {
      console.warn("Notice reading Firestore pending leaves for sync:", dbErr.message);
    }

    if (pendingMap.size === 0) return;

    console.log(`Found ${pendingMap.size} pending leaves to sync to Google Sheets...`);
    
    const pendingByWs: Record<string, { docId: string; data: any }[]> = {};
    pendingMap.forEach((item, docId) => {
      const wsId = item.workshopId || 'default';
      if (!pendingByWs[wsId]) pendingByWs[wsId] = [];
      pendingByWs[wsId].push({ docId, data: item });
    });

    const syncedDocIds: string[] = [];

    for (const [wsId, items] of Object.entries(pendingByWs)) {
      const targetTitle = await getOrEnsureLeaveSheetTitle(sheets, spreadsheetId, wsId === 'default' ? undefined : wsId);
      const rowsToAppend = items.map(item => [
        item.docId,
        item.data.name,
        item.data.birthYear || "",
        item.data.chucDanh,
        String(item.data.kip),
        item.data.startDate,
        item.data.endDate,
        item.data.reason || "Giải quyết việc riêng gia đình",
        item.data.phone || "",
        item.data.location || "Gia Lai",
        item.data.status || "Chờ phân ca",
        item.data.dateStr || item.data.createdAt || "",
        item.data.workshopId || (wsId === 'default' ? "" : wsId)
      ]);

      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: `'${targetTitle}'!A:M`,
        valueInputOption: "RAW",
        requestBody: {
          values: rowsToAppend
        }
      });

      items.forEach(i => syncedDocIds.push(i.docId));
    }

    if (syncedDocIds.length > 0) {
      // Clear from localStore
      if (localStore.pendingLeaves && Array.isArray(localStore.pendingLeaves)) {
        localStore.pendingLeaves = localStore.pendingLeaves.filter((p: any) => !syncedDocIds.includes(p.id));
        persistLocalStore();
      }

      // Clear from Firestore
      try {
        const db = getFirestoreInstance();
        const batch = db.batch();
        syncedDocIds.forEach(id => {
          batch.delete(db.collection("pending_leaves").doc(id));
        });
        await batch.commit();
      } catch (e: any) {}

      console.log(`Successfully synced ${syncedDocIds.length} pending leaves to Google Sheets.`);
    }
  } catch (err: any) {
    console.error("Failed to sync pending leaves to Google Sheets:", err.message);
  }
}

app.get("/api/sheets/leave-requests", async (req, res) => {
  const { workshopId } = req.query;
  const targetWsId = typeof workshopId === 'string' ? workshopId : undefined;

  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();

  // FALLBACK: If not connected to Google Sheets, load from Firestore or localStore pending_leaves
  if (!tokens) {
    const pendingLeavesMap = new Map<string, any>();
    
    // 1. Add from localStore
    if (localStore.pendingLeaves && Array.isArray(localStore.pendingLeaves)) {
      localStore.pendingLeaves.forEach((item: any) => {
        pendingLeavesMap.set(item.id, {
          id: item.id,
          name: item.name || "",
          birthYear: item.birthYear || "",
          chucDanh: item.chucDanh || "",
          kip: String(item.kip || ""),
          startDate: item.startDate || "",
          endDate: item.endDate || "",
          reason: item.reason || "",
          phone: item.phone || "",
          location: item.location || "",
          status: item.status || "Chờ phân ca",
          createdAt: item.createdAt || item.dateStr || "",
          workshopId: item.workshopId || "",
          isPendingSync: true
        });
      });
    }

    // 2. Add from Firestore
    try {
      const db = getFirestoreInstance();
      const snapshot = await db.collection("pending_leaves").get();
      snapshot.forEach(doc => {
        const item = doc.data();
        pendingLeavesMap.set(doc.id, {
          id: doc.id,
          name: item.name || "",
          birthYear: item.birthYear || "",
          chucDanh: item.chucDanh || "",
          kip: String(item.kip || ""),
          startDate: item.startDate || "",
          endDate: item.endDate || "",
          reason: item.reason || "",
          phone: item.phone || "",
          location: item.location || "",
          status: item.status || "Chờ phân ca",
          createdAt: item.dateStr || item.createdAt || "",
          workshopId: item.workshopId || "",
          isPendingSync: true
        });
      });
    } catch (dbErr: any) {
      console.warn("Firestore pending leaves read notice (using localStore fallback):", dbErr.message);
    }

    const result = Array.from(pendingLeavesMap.values());
    return res.json(result);
  }

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const spreadsheetId = resolveSpreadsheetId(req);

  try {
    await syncPendingLeavesToSheets(sheets, spreadsheetId);

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const allSheetObjects = spreadsheet.data.sheets || [];
    const allSheetTitles = allSheetObjects.map((s: any) => s.properties?.title || "");

    let sheetsToRead: string[] = [];
    if (targetWsId && targetWsId !== 'all') {
      const sheetTitle = await getOrEnsureLeaveSheetTitle(sheets, spreadsheetId, targetWsId);
      sheetsToRead = [sheetTitle];
    } else {
      sheetsToRead = allSheetTitles.filter(t => t.toLowerCase().startsWith("danh_sach_nghi"));
      if (sheetsToRead.length === 0) {
        const defaultTitle = await getOrEnsureLeaveSheetTitle(sheets, spreadsheetId);
        sheetsToRead = [defaultTitle];
      }
    }

    const leaveRequests: any[] = [];

    // Also include any leftover pending leaves from localStore and Firestore cache
    const pendingMap = new Map<string, any>();
    if (localStore.pendingLeaves && Array.isArray(localStore.pendingLeaves)) {
      localStore.pendingLeaves.forEach(item => {
        pendingMap.set(item.id, {
          id: item.id,
          name: item.name || "",
          birthYear: item.birthYear || "",
          chucDanh: item.chucDanh || "",
          kip: String(item.kip || ""),
          startDate: item.startDate || "",
          endDate: item.endDate || "",
          reason: item.reason || "",
          phone: item.phone || "",
          location: item.location || "",
          status: item.status || "Chờ phân ca",
          createdAt: item.createdAt || item.dateStr || "",
          workshopId: item.workshopId || "",
          isPendingSync: true
        });
      });
    }

    try {
      const db = getFirestoreInstance();
      const snapshot = await db.collection("pending_leaves").get();
      snapshot.forEach(doc => {
        const item = doc.data();
        pendingMap.set(doc.id, {
          id: doc.id,
          name: item.name || "",
          birthYear: item.birthYear || "",
          chucDanh: item.chucDanh || "",
          kip: String(item.kip || ""),
          startDate: item.startDate || "",
          endDate: item.endDate || "",
          reason: item.reason || "",
          phone: item.phone || "",
          location: item.location || "",
          status: item.status || "Chờ phân ca",
          createdAt: item.dateStr || item.createdAt || "",
          workshopId: item.workshopId || "",
          isPendingSync: true
        });
      });
    } catch (dbErr) {}

    pendingMap.forEach(item => leaveRequests.push(item));

    for (const sheetTitle of sheetsToRead) {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `'${sheetTitle}'!A1:M2000`
        });
        
        const rows = response.data.values;
        if (!rows || rows.length <= 1) continue;

        const headers = rows[0];
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0 || !row[0]) continue;
          
          const item: any = { rowIndex: i + 1, sheetTitle };
          headers.forEach((header: string, index: number) => {
            item[header] = row[index] || "";
          });
          
          const leaveChucDanh = (item["Chức danh"] || "") === "Trực phụ cơ MR" ? "Trực phụ máy MR" : (item["Chức danh"] || "");
          
          if (leaveRequests.some(l => l.id === item["Mã đơn"])) continue;

          const leave = {
            id: item["Mã đơn"] || "",
            name: item["Họ và tên"] || "",
            birthYear: item["Năm sinh"] || "",
            chucDanh: leaveChucDanh,
            kip: item["Kíp"] || "",
            startDate: item["Từ ngày"] || "",
            endDate: item["Đến ngày"] || "",
            reason: item["Lý do"] || "",
            phone: item["Điện thoại"] || "",
            location: item["Địa điểm"] || "",
            status: item["Trạng thái"] || "",
            createdAt: item["Ngày tạo"] || "",
            workshopId: item["Phân xưởng"] || item.workshopId || "",
            rowIndex: item.rowIndex,
            sheetTitle
          };
          leaveRequests.push(leave);
        }
      } catch (sheetErr: any) {
        console.error(`Error reading sheet ${sheetTitle}:`, sheetErr.message);
      }
    }

    res.json(leaveRequests);
  } catch (error: any) {
    console.error("Error fetching leave requests:", error);
    if (await handleAuthErrorIfAny(error, res)) return;
    res.status(500).json({ error: "Failed to fetch leave requests", details: error.message });
  }
});

app.post("/api/sheets/leave-requests", async (req, res) => {
  const { name, birthYear, chucDanh, kip, startDate, endDate, reason, phone, location, leaveBalance, workshopId } = req.body;
  
  if (!name || !chucDanh || !kip || !startDate || !endDate) {
    return res.status(400).json({ error: "Thiếu thông tin bắt buộc" });
  }

  const normalizedChucDanh = chucDanh === "Trực phụ cơ MR" ? "Trực phụ máy MR" : chucDanh;
  const id = "LEAVE_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
  const dateStr = new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
  const status = req.body.status || "Chờ phân ca";

  // Try to get leave balance if not supplied in the request body
  let finalLeaveBalance = leaveBalance;
  if (!finalLeaveBalance) {
    finalLeaveBalance = await fetchLeaveBalanceHelper(name, req.body.googleSheetUrl || req.body.spreadsheetId);
  }

  // ALWAYS send Zalo Notification first (unblocked by Google Sheets state)
  sendZaloNotification({
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
  });

  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();

  // If Google Sheets is not authorized, save to localStore and Firestore pending_leaves
  if (!tokens) {
    const leaveItem = {
      id,
      name,
      birthYear: birthYear || "",
      chucDanh: normalizedChucDanh,
      kip,
      startDate,
      endDate,
      reason: reason || "Giải quyết việc riêng gia đình",
      phone: phone || "",
      location: location || "Gia Lai",
      status,
      dateStr,
      workshopId: workshopId || "",
      createdAt: new Date().toISOString()
    };

    const pIdx = localStore.pendingLeaves.findIndex(p => p.id === id);
    if (pIdx >= 0) localStore.pendingLeaves[pIdx] = leaveItem;
    else localStore.pendingLeaves.push(leaveItem);
    persistLocalStore();

    try {
      const db = getFirestoreInstance();
      await db.collection("pending_leaves").doc(id).set(leaveItem);
    } catch (dbErr: any) {
      console.warn("Firestore save pending leave notice (saved in localStore fallback):", dbErr.message);
    }

    return res.json({ 
      success: true, 
      message: `✅ Đã lưu đơn của đồng chí ${name} lên hệ thống và gửi thông báo Zalo thành công! (Sẽ tự động đồng bộ lên Google Sheets sau)`,
      id,
      status
    });
  }

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const spreadsheetId = resolveSpreadsheetId(req);
  ensureAnnualLeaveSheetExists(sheets, spreadsheetId, localStore.appSettings?.staffData).catch(() => {});

  try {
    // Attempt auto-sync of any previous pending leaves first
    await syncPendingLeavesToSheets(sheets, spreadsheetId);

    const sheetTitle = await getOrEnsureLeaveSheetTitle(sheets, spreadsheetId, workshopId);
    
    const rowValue = [
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
      workshopId || ""
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `'${sheetTitle}'!A:M`,
      valueInputOption: "RAW",
      requestBody: {
        values: [rowValue]
      }
    });

    res.json({ 
      success: true, 
      message: `✅ Đã lưu đơn của đồng chí ${name} lên Google Sheets và gửi thông báo Zalo thành công!`,
      id, 
      status 
    });
  } catch (error: any) {
    console.error("Error saving leave request to Sheets, saving to Firestore as fallback:", error);
    
    // Save to Firestore so it is not lost, and we can retry later
    try {
      const db = getFirestoreInstance();
      await db.collection("pending_leaves").doc(id).set({
        name,
        birthYear: birthYear || "",
        chucDanh: normalizedChucDanh,
        kip,
        startDate,
        endDate,
        reason: reason || "Giải quyết việc riêng gia đình",
        phone: phone || "",
        location: location || "Gia Lai",
        status,
        dateStr,
        workshopId: workshopId || "",
        createdAt: new Date().toISOString()
      });
      
      res.json({ 
        success: true, 
        message: `✅ Đã gửi thông báo Zalo & lưu đơn của đồng chí ${name} thành công! (Do Google Sheets đang gián đoạn, đơn đã được lưu tạm trên hệ thống).`,
        id,
        status
      });
    } catch (dbErr: any) {
      res.status(500).json({ error: "Lưu đơn thất bại", details: error.message });
    }
  }
});

app.get("/api/sheets/leave-balance", async (req, res) => {
  const { name } = req.query;
  if (!name) {
    return res.status(400).json({ error: "Thiếu tên nhân viên" });
  }

  const targetName = String(name).trim();
  const targetNormalized = removeVietnameseAccents(targetName);
  const spreadsheetId = resolveSpreadsheetId(req);

  const fallbackToLocalStaffData = () => {
    let staffData: string[][] = localStore.appSettings?.staffData || [];
    if (staffData.length === 0) {
      staffData = DEFAULT_WORKSHOPS[0]?.staffData || [];
    }
    
    for (const row of staffData) {
      if (!Array.isArray(row)) continue;
      for (let colIdx = 1; colIdx < row.length; colIdx++) {
        const pName = (row[colIdx] || '').trim();
        if (pName && (pName.toLowerCase() === targetName.toLowerCase() || removeVietnameseAccents(pName) === targetNormalized)) {
          return {
            success: true,
            entitled: "12",
            used: "0",
            remaining: "12",
            source: "default"
          };
        }
      }
    }
    return null;
  };

  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();

  if (!tokens) {
    const localMatch = fallbackToLocalStaffData();
    if (localMatch) return res.json(localMatch);
    return res.status(401).json({ error: "Chưa kết nối Google Sheets. Vui lòng kết nối trước." });
  }

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  try {
    let targetSheetName = "Số ngày phép";
    try {
      const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
      const sheetNames = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];
      targetSheetName = sheetNames.find(n => n?.toLowerCase().trim() === "số ngày phép") || sheetNames[0] || "Số ngày phép";
    } catch (e) {}

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `'${targetSheetName}'!A1:H1000`
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      const localMatch = fallbackToLocalStaffData();
      if (localMatch) return res.json(localMatch);
      return res.status(404).json({ error: "Không tìm thấy dữ liệu trong bảng tính" });
    }

    let matchedRow: any[] | null = null;
    
    for (const row of rows) {
      if (!row || row.length === 0) continue;
      for (let colIdx = 0; colIdx < Math.min(row.length, 4); colIdx++) {
        const cellVal = String(row[colIdx] || '').trim().toLowerCase();
        if (cellVal === targetName.toLowerCase()) {
          matchedRow = row;
          break;
        }
      }
      if (matchedRow) break;
    }

    if (!matchedRow) {
      for (const row of rows) {
        if (!row || row.length === 0) continue;
        for (let colIdx = 0; colIdx < Math.min(row.length, 4); colIdx++) {
          const cellVal = String(row[colIdx] || '').trim();
          if (removeVietnameseAccents(cellVal) === targetNormalized) {
            matchedRow = row;
            break;
          }
        }
        if (matchedRow) break;
      }
    }

    if (!matchedRow) {
      const localMatch = fallbackToLocalStaffData();
      if (localMatch) return res.json(localMatch);
      return res.status(404).json({ error: `Không tìm thấy thông tin phép năm của đồng chí ${targetName}` });
    }

    const headerRow = rows[0] || [];
    let oldLeavesCol = 4; // Column E
    let newLeavesCol = 5; // Column F (Phép năm mới được hưởng)
    let usedCol = 6;      // Column G (Phép đã nghỉ năm)
    let remainingCol = 7; // Column H (Phép còn lại hiện tại)

    for (let c = 0; c < headerRow.length; c++) {
      const hText = removeVietnameseAccents(String(headerRow[c] || '')).toLowerCase();
      if (hText.includes("nam cu") || hText.includes("phep nam cu")) {
        oldLeavesCol = c;
      } else if (hText.includes("nam moi") || hText.includes("duoc huong") || hText.includes("dinh muc")) {
        newLeavesCol = c;
      } else if (hText.includes("da nghi")) {
        usedCol = c;
      } else if (hText.includes("con lai")) {
        remainingCol = c;
      }
    }

    const oldLeaves = parseFloat(String(matchedRow[oldLeavesCol] || '0').trim()) || 0;
    const newLeaves = parseFloat(String(matchedRow[newLeavesCol] || '12').trim()) || 12;
    const usedVal = parseFloat(String(matchedRow[usedCol] || '0').trim()) || 0;

    const calc = calculateLeaveBalanceLogic(oldLeaves, newLeaves, usedVal);

    const entitled = String(calc.entitled);
    const used = String(calc.usedLeaves);
    const remaining = (matchedRow[remainingCol] !== undefined && String(matchedRow[remainingCol]).trim() !== '')
      ? String(matchedRow[remainingCol]).trim()
      : String(calc.remaining);

    return res.json({
      success: true,
      name: targetName,
      entitled,
      used,
      remaining,
      note: calc.note,
      oldLeaves: String(calc.oldLeaves),
      newLeaves: String(calc.newLeaves)
    });

  } catch (error: any) {
    console.error("Error fetching leave balance:", error);
    const localMatch = fallbackToLocalStaffData();
    if (localMatch) return res.json(localMatch);
    return res.status(500).json({ error: "Lỗi kết nối hoặc không thể lấy dữ liệu", details: error.message });
  }
});

app.post("/api/sheets/sync-staff-leaves", async (req, res) => {
  const { staffData } = req.body;
  const spreadsheetId = resolveSpreadsheetId(req);

  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();

  if (!tokens) {
    return res.status(401).json({ error: "Chưa kết nối Google Sheets. Vui lòng kết nối tài khoản Google trước." });
  }

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  try {
    const actualData = staffData || localStore.appSettings?.staffData || [];
    const sheetName = await ensureAnnualLeaveSheetExists(sheets, spreadsheetId, actualData);
    res.json({
      success: true,
      message: `✅ Đã khởi tạo và đồng bộ thành công danh sách nhân sự sang sheet '${sheetName}' trên Google Sheets!`,
      spreadsheetId
    });
  } catch (error: any) {
    console.error("Error in sync-staff-leaves:", error);
    if (await handleAuthErrorIfAny(error, res)) return;
    res.status(500).json({ error: "Không thể đồng bộ danh sách nhân sự sang Google Sheets", details: error.message });
  }
});

app.post("/api/sheets/leave-requests/update-status", async (req, res) => {
  const { ids, status, workshopId } = req.body;
  
  if (!ids || !Array.isArray(ids) || ids.length === 0 || !status) {
    return res.status(400).json({ error: "Tham số không hợp lệ." });
  }

  // Always update in localStore.pendingLeaves
  if (localStore.pendingLeaves && Array.isArray(localStore.pendingLeaves)) {
    localStore.pendingLeaves.forEach((p: any) => {
      if (ids.includes(p.id)) p.status = status;
    });
    persistLocalStore();
  }

  // Always update in Firestore
  try {
    const db = getFirestoreInstance();
    const batch = db.batch();
    ids.forEach((id: string) => {
      batch.set(db.collection("pending_leaves").doc(id), { status }, { merge: true });
    });
    await batch.commit();
  } catch (e) {}

  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();
  
  if (!tokens) {
    return res.json({ success: true, message: "Đã cập nhật trạng thái đơn trên hệ thống nội bộ!" });
  }

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const spreadsheetId = resolveSpreadsheetId(req);

  try {
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const allSheetObjects = spreadsheet.data.sheets || [];
    const allSheetTitles = allSheetObjects.map((s: any) => s.properties?.title || "");

    let sheetsToSearch = allSheetTitles.filter(t => t.toLowerCase().startsWith("danh_sach_nghi"));
    if (sheetsToSearch.length === 0) {
      const defaultTitle = await getOrEnsureLeaveSheetTitle(sheets, spreadsheetId, workshopId);
      sheetsToSearch = [defaultTitle];
    }

    const updatedRows = [];
    for (const sheetTitle of sheetsToSearch) {
      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `'${sheetTitle}'!A1:M2000`
        });
        
        const rows = response.data.values;
        if (!rows || rows.length <= 1) continue;

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || !row[0]) continue;
          
          const itemId = row[0];
          if (ids.includes(itemId)) {
            const rowNum = i + 1;
            await sheets.spreadsheets.values.update({
              spreadsheetId,
              range: `'${sheetTitle}'!K${rowNum}`,
              valueInputOption: "RAW",
              requestBody: {
                values: [[status]]
              }
            });
            updatedRows.push({ id: itemId, rowNum, sheetTitle });
          }
        }
      } catch (sheetErr: any) {
        console.error(`Error updating status on sheet ${sheetTitle}:`, sheetErr.message);
      }
    }

    res.json({ success: true, updatedCount: updatedRows.length, updatedRows });
  } catch (error: any) {
    console.error("Error updating status:", error);
    if (await handleAuthErrorIfAny(error, res)) return;
    res.status(500).json({ error: "Failed to update leave status", details: error.message });
  }
});

app.post("/api/sheets/leave-requests/delete", async (req, res) => {
  const { ids } = req.body;
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ error: "Tham số không hợp lệ." });
  }

  // Always delete from localStore.pendingLeaves
  if (localStore.pendingLeaves && Array.isArray(localStore.pendingLeaves)) {
    localStore.pendingLeaves = localStore.pendingLeaves.filter((p: any) => !ids.includes(p.id));
    persistLocalStore();
  }

  // Always delete from Firestore
  try {
    const db = getFirestoreInstance();
    const batch = db.batch();
    ids.forEach((id: string) => {
      batch.delete(db.collection("pending_leaves").doc(id));
    });
    await batch.commit();
  } catch (e) {}

  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();
  
  if (!tokens) {
    return res.json({ success: true, message: "Đã xóa đơn khỏi hệ thống nội bộ!" });
  }

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const spreadsheetId = resolveSpreadsheetId(req);

  try {
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const allSheetObjects = spreadsheet.data.sheets || [];
    const normalizedTargetIds = ids.map(id => String(id).trim().toLowerCase());

    let deletedCount = 0;

    for (const sheetObj of allSheetObjects) {
      const sheetTitle = sheetObj.properties?.title || "";
      if (!sheetTitle.toLowerCase().startsWith("danh_sach_nghi")) continue;

      const sheetId = sheetObj.properties?.sheetId;
      if (sheetId === undefined || sheetId === null) continue;

      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `'${sheetTitle}'!A1:M2000`
        });
        
        const rows = response.data.values;
        if (!rows || rows.length <= 1) continue;

        const rowsToDelete: number[] = [];

        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || !row[0]) continue;
          
          const itemId = String(row[0]).trim().toLowerCase();
          if (normalizedTargetIds.includes(itemId)) {
            rowsToDelete.push(i + 1);
          }
        }

        if (rowsToDelete.length > 0) {
          rowsToDelete.sort((a, b) => b - a);
          const requests = rowsToDelete.map(rowNum => ({
            deleteDimension: {
              range: {
                sheetId: Number(sheetId),
                dimension: "ROWS",
                startIndex: Number(rowNum - 1),
                endIndex: Number(rowNum)
              }
            }
          }));

          await sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests }
          });

          deletedCount += rowsToDelete.length;
        }
      } catch (sheetErr: any) {
        console.error(`Error deleting rows from sheet ${sheetTitle}:`, sheetErr.message);
      }
    }

    res.json({ success: true, deletedCount });
  } catch (error: any) {
    console.error("Error deleting leave request:", error);
    if (await handleAuthErrorIfAny(error, res)) return;
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
    console.log("Tokens refreshed, saving to Firestore...");
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
      try {
        const db = getFirestoreInstance();
        await db.collection("config").doc("google_auth").delete();
      } catch (e) {}
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
    console.log("Tokens refreshed, saving to Firestore...");
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
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();

export default app;
