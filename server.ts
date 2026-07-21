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

// Initialize Firebase Admin
if (admin.apps.length === 0) {
  try {
    if (process.env.FIREBASE_SERVICE_ACCOUNT) {
      const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
      admin.initializeApp({
        credential: admin.credential.cert(serviceAccount),
        projectId: serviceAccount.project_id || firebaseConfig.projectId,
      });
      console.log("Firebase Admin initialized with Service Account from ENV");
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
    console.error("Error saving tokens to Firestore:", e);
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
      console.warn("Primary DB load failed, trying fallback:", e.message);
      if (e.code === 5 || e.message?.includes("NOT_FOUND") || e.message?.includes("not found")) {
        const tokens = await tryLoad(getFirestoreInstance(true));
        if (tokens) return tokens;
      } else {
        throw e;
      }
    }
  } catch (e: any) {
    lastFirestoreError = `Load Error: ${e.message}`;
    console.error("Error loading tokens from Firestore:", e);
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
      console.warn("Primary DB clear failed, trying fallback:", e.message);
      if (e.code === 5 || e.message?.includes("NOT_FOUND") || e.message?.includes("not found")) {
        await tryClear(getFirestoreInstance(true));
        lastFirestoreError = null;
      } else {
        throw e;
      }
    }
  } catch (e: any) {
    lastFirestoreError = `Clear Error: ${e.message}`;
    console.error("Error clearing tokens from Firestore:", e);
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

async function fetchLeaveBalanceHelper(name: string): Promise<{ entitled: string; used: string; remaining: string } | null> {
  try {
    const oauth2Client = getOAuth2Client();
    let tokens = await loadTokens();
    if (!tokens) return null;

    oauth2Client.setCredentials(tokens);
    const sheets = google.sheets({ version: "v4", auth: oauth2Client });
    const spreadsheetId = '1pH1-Nj4B1nauoEfO5cZG13Wlk_UrUrFDq_eucf5a-IY';

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "A1:H1000"
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) return null;

    const targetName = String(name).trim();
    const targetNormalized = removeVietnameseAccents(targetName);

    let matchedRow: any[] | null = null;
    
    // First pass: exact match (ignoring case and whitespace)
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

    // Second pass: normalized match (removing accents) if exact match not found
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

    if (matchedRow) {
      const entitled = String(matchedRow[4] || '0').trim();
      const used = String(matchedRow[5] || '0').trim();
      const remaining = String(matchedRow[6] || '0').trim();
      return { entitled, used, remaining };
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

app.post("/api/signatures", express.json({ limit: '10mb' }), async (req, res) => {
  const { name, data } = req.body;
  if (!name || !data) {
    return res.status(400).json({ error: "Name and data are required" });
  }
  try {
    const db = getFirestoreInstance();
    await db.collection("signatures").doc(name).set({
      name,
      data,
      updatedAt: new Date().toISOString()
    });
    res.json({ success: true });
  } catch (e: any) {
    console.error("Error saving signature:", e);
    res.status(500).json({ error: e.message });
  }
});

// App Settings Endpoints (Staff List & Config)
app.get("/api/app-settings", async (req, res) => {
  try {
    const db = getFirestoreInstance();
    const doc = await db.collection("config").doc("app_settings").get();
    if (doc.exists) {
      const data = doc.data();
      if (data) {
        if (typeof data.staffData === 'string') {
          try {
            data.staffData = JSON.parse(data.staffData);
          } catch (e) {
            console.error("Failed to parse staffData JSON", e);
            data.staffData = []; // Fallback to empty array on parse error
          }
        }
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
            try {
              await db.collection("config").doc("app_settings").set({
                staffData: JSON.stringify(migrated)
              }, { merge: true });
              console.log("Migrated 'Trực phụ cơ MR' to 'Trực phụ máy MR' in Firestore settings.");
            } catch (saveErr) {
              console.error("Failed to save migrated staffData in Firestore", saveErr);
            }
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
          try {
            await db.collection("config").doc("app_settings").set({
              config: data.config
            }, { merge: true });
            console.log("Migrated 'zaloWebhookUrl' to 'vhialy.dpdns.org' in Firestore settings.");
          } catch (saveErr) {
            console.error("Failed to save migrated zaloWebhookUrl in Firestore", saveErr);
          }
        }
      }
      res.json(data);
    } else {
      res.json({ staffData: null, config: null });
    }
  } catch (e: any) {
    console.error("Error fetching app settings:", e);
    res.status(500).json({ error: e.message });
  }
});

app.post("/api/app-settings", express.json({ limit: '5mb' }), async (req, res) => {
  const { staffData, config } = req.body;
  try {
    const db = getFirestoreInstance();
    // Firestore does not support nested arrays, so we stringify staffData
    const serializedStaffData = JSON.stringify(staffData);
    
    await db.collection("config").doc("app_settings").set({
      staffData: serializedStaffData,
      config,
      updatedAt: new Date().toISOString()
    }, { merge: true });
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

async function ensureLeaveSheetExists(sheets: any, spreadsheetId: string) {
  try {
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = spreadsheet.data.sheets?.map((s: any) => s.properties?.title) || [];
    
    if (!sheetNames.includes("DANH_SACH_NGHI")) {
      // Create sheet
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: "DANH_SACH_NGHI"
                }
              }
            }
          ]
        }
      });
      // Append header
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: "DANH_SACH_NGHI!A1:L1",
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
            "Ngày tạo"
          ]]
        }
      });
      console.log("Created 'DANH_SACH_NGHI' sheet and headers.");
    }
  } catch (e: any) {
    console.error("Error securing leave sheet:", e.message);
    throw e;
  }
}

async function syncPendingLeavesToSheets(sheets: any, spreadsheetId: string) {
  try {
    const db = getFirestoreInstance();
    const snapshot = await db.collection("pending_leaves").get();
    if (snapshot.empty) return;

    console.log(`Found ${snapshot.size} pending leaves in Firestore. Syncing to Google Sheets...`);
    
    await ensureLeaveSheetExists(sheets, spreadsheetId);
    
    const rowsToAppend: any[][] = [];
    const docIds: string[] = [];

    snapshot.forEach(doc => {
      const data = doc.data();
      const rowValue = [
        doc.id,
        data.name,
        data.birthYear || "",
        data.chucDanh,
        String(data.kip),
        data.startDate,
        data.endDate,
        data.reason || "Giải quyết việc riêng gia đình",
        data.phone || "",
        data.location || "Gia Lai",
        data.status || "Chờ phân ca",
        data.dateStr
      ];
      rowsToAppend.push(rowValue);
      docIds.push(doc.id);
    });

    if (rowsToAppend.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: "DANH_SACH_NGHI!A:L",
        valueInputOption: "RAW",
        requestBody: {
          values: rowsToAppend
        }
      });
      
      const batch = db.batch();
      docIds.forEach(id => {
        batch.delete(db.collection("pending_leaves").doc(id));
      });
      await batch.commit();
      console.log(`Successfully synced ${rowsToAppend.length} leaves to Google Sheets and cleaned up Firestore.`);
    }
  } catch (err: any) {
    console.error("Failed to sync pending leaves to Google Sheets:", err.message);
  }
}

app.get("/api/sheets/leave-requests", async (req, res) => {
  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();

  // FALLBACK: If not connected to Google Sheets, load from Firestore pending_leaves
  if (!tokens) {
    try {
      const db = getFirestoreInstance();
      const snapshot = await db.collection("pending_leaves").get();
      const pendingLeaves: any[] = [];
      snapshot.forEach(doc => {
        const item = doc.data();
        pendingLeaves.push({
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
          createdAt: item.dateStr || "",
          isPendingSync: true
        });
      });
      return res.json(pendingLeaves);
    } catch (dbErr: any) {
      console.error("Failed to load pending leaves from Firestore:", dbErr);
      return res.json([]);
    }
  }

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const spreadsheetId = '1HgGW-FvoGQXtj7V_JCMD-7Tuue0rTIM-bmohGmgqm6I';

  try {
    // Try to sync pending leaves from Firestore to Sheets first
    await syncPendingLeavesToSheets(sheets, spreadsheetId);

    await ensureLeaveSheetExists(sheets, spreadsheetId);
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "DANH_SACH_NGHI!A1:L2000"
    });
    
    const rows = response.data.values;
    const leaveRequests: any[] = [];
    
    // Also include any leftover pending leaves in the Firestore cache
    try {
      const db = getFirestoreInstance();
      const snapshot = await db.collection("pending_leaves").get();
      snapshot.forEach(doc => {
        const item = doc.data();
        leaveRequests.push({
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
          createdAt: item.dateStr || "",
          isPendingSync: true
        });
      });
    } catch (dbErr) {}

    if (rows && rows.length > 1) {
      const headers = rows[0];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0 || !row[0]) continue;
        
        const item: any = { rowIndex: i + 1 };
        headers.forEach((header: string, index: number) => {
          item[header] = row[index] || "";
        });
        
        const leaveChucDanh = (item["Chức danh"] || "") === "Trực phụ cơ MR" ? "Trực phụ máy MR" : (item["Chức danh"] || "");
        
        // Prevent duplicates if already loaded from Firestore pending cache
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
          rowIndex: item.rowIndex
        };
        leaveRequests.push(leave);
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
  const { name, birthYear, chucDanh, kip, startDate, endDate, reason, phone, location, leaveBalance } = req.body;
  
  if (!name || !chucDanh || !kip || !startDate || !endDate) {
    return res.status(400).json({ error: "Thiếu thông tin bắt buộc" });
  }

  const normalizedChucDanh = chucDanh === "Trực phụ cơ MR" ? "Trực phụ máy MR" : chucDanh;
  const id = "LEAVE_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
  const dateStr = new Date().toLocaleString('vi-VN', { timeZone: 'Asia/Ho_Chi_Minh' });
  const status = "Chờ phân ca";

  // Try to get leave balance if not supplied in the request body
  let finalLeaveBalance = leaveBalance;
  if (!finalLeaveBalance) {
    finalLeaveBalance = await fetchLeaveBalanceHelper(name);
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

  // If Google Sheets is not authorized, save to Firestore pending_leaves
  if (!tokens) {
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
        createdAt: new Date().toISOString()
      });
      
      return res.json({ 
        success: true, 
        message: `✅ Đã lưu đơn của đồng chí ${name} lên hệ thống và gửi thông báo Zalo thành công! (Sẽ tự động đồng bộ lên Google Sheets sau)`,
        id,
        status
      });
    } catch (dbErr: any) {
      console.error("Failed to save pending leave:", dbErr);
      return res.status(500).json({ error: "Lưu đơn thất bại. Không thể lưu vào hệ thống." });
    }
  }

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const spreadsheetId = '1HgGW-FvoGQXtj7V_JCMD-7Tuue0rTIM-bmohGmgqm6I';

  try {
    // Attempt auto-sync of any previous pending leaves first
    await syncPendingLeavesToSheets(sheets, spreadsheetId);

    await ensureLeaveSheetExists(sheets, spreadsheetId);
    
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
      dateStr
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "DANH_SACH_NGHI!A:L",
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

  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();

  if (!tokens) {
    return res.status(401).json({ error: "Chưa kết nối Google Sheets. Vui lòng kết nối trước." });
  }

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const spreadsheetId = '1pH1-Nj4B1nauoEfO5cZG13Wlk_UrUrFDq_eucf5a-IY';

  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "A1:H1000"
    });

    const rows = response.data.values;
    if (!rows || rows.length === 0) {
      return res.status(404).json({ error: "Không tìm thấy dữ liệu trong bảng tính" });
    }

    const targetName = String(name).trim();
    const targetNormalized = removeVietnameseAccents(targetName);

    let matchedRow: any[] | null = null;
    
    // First pass: exact match (ignoring case and whitespace)
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

    // Second pass: normalized match (removing accents) if exact match not found
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
      return res.status(404).json({ error: `Không tìm thấy thông tin phép năm của đồng chí ${targetName}` });
    }

    // Extract columns E (index 4), F (index 5), G (index 6)
    // E: Số ngày phép được hưởng
    // F: Số ngày phép đã nghỉ
    // G: Số phép còn lại
    const entitled = String(matchedRow[4] || '0').trim();
    const used = String(matchedRow[5] || '0').trim();
    const remaining = String(matchedRow[6] || '0').trim();

    return res.json({
      success: true,
      name: targetName,
      entitled,
      used,
      remaining
    });

  } catch (error: any) {
    console.error("Error fetching leave balance:", error);
    return res.status(500).json({ error: "Lỗi kết nối hoặc không thể lấy dữ liệu", details: error.message });
  }
});

app.post("/api/sheets/leave-requests/update-status", async (req, res) => {
  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();
  if (!tokens) return res.status(401).json({ error: "Chưa kết nối Google Sheets." });

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const spreadsheetId = '1HgGW-FvoGQXtj7V_JCMD-7Tuue0rTIM-bmohGmgqm6I';

  const { ids, status } = req.body;
  
  if (!ids || !Array.isArray(ids) || ids.length === 0 || !status) {
    return res.status(400).json({ error: "Tham số không hợp lệ." });
  }

  try {
    await ensureLeaveSheetExists(sheets, spreadsheetId);
    
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "DANH_SACH_NGHI!A1:L2000"
    });
    
    const rows = response.data.values;
    if (!rows || rows.length <= 1) {
      return res.status(404).json({ error: "Không tìm thấy dữ liệu đơn nghỉ" });
    }

    const updatedRows = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row[0]) continue;
      
      const itemId = row[0];
      if (ids.includes(itemId)) {
        const rowNum = i + 1;
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `DANH_SACH_NGHI!K${rowNum}`,
          valueInputOption: "RAW",
          requestBody: {
            values: [[status]]
          }
        });
        updatedRows.push({ id: itemId, rowNum });
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
  const oauth2Client = getOAuth2Client();
  let tokens = null;
  const tokensStr = req.cookies.google_tokens;
  if (tokensStr) {
    try { tokens = JSON.parse(tokensStr); } catch (e) {}
  }
  if (!tokens) tokens = await loadTokens();
  if (!tokens) return res.status(401).json({ error: "Chưa kết nối Google Sheets." });

  oauth2Client.setCredentials(tokens);
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });
  const spreadsheetId = '1HgGW-FvoGQXtj7V_JCMD-7Tuue0rTIM-bmohGmgqm6I';

  const { ids } = req.body;
  if (!ids || !Array.isArray(ids) || ids.length === 0) {
    return res.status(400).json({ error: "Tham số không hợp lệ." });
  }

  try {
    await ensureLeaveSheetExists(sheets, spreadsheetId);
    
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetObj = spreadsheet.data.sheets?.find((s: any) => s.properties?.title === "DANH_SACH_NGHI");
    if (!sheetObj) {
      return res.status(404).json({ error: "Không tìm thấy bảng DANH_SACH_NGHI" });
    }
    const sheetId = sheetObj.properties?.sheetId;
    if (sheetId === undefined || sheetId === null) {
      return res.status(500).json({ error: "Không tìm thấy Sheet ID" });
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "DANH_SACH_NGHI!A1:L2000"
    });
    
    const rows = response.data.values;
    if (!rows || rows.length <= 1) {
      return res.status(404).json({ error: "Không tìm thấy dữ liệu đơn nghỉ" });
    }

    const normalizedTargetIds = ids.map(id => String(id).trim().toLowerCase());
    const rowsToDelete: number[] = [];
    const loggedIds: string[] = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row[0]) continue;
      
      const itemId = String(row[0]).trim();
      loggedIds.push(itemId);
      if (normalizedTargetIds.includes(itemId.toLowerCase())) {
        rowsToDelete.push(i + 1);
      }
    }

    if (rowsToDelete.length === 0) {
      console.warn("Không tìm thấy đơn cần xóa. Cần xóa:", normalizedTargetIds, "Hiện có trong sheet:", loggedIds);
      return res.status(404).json({ 
        error: "Không tìm thấy mã đơn cần xóa trong trang tính.", 
        details: `Cần xóa: ${ids.join(', ')}. Hiện có: ${loggedIds.slice(0, 10).join(', ')}`
      });
    }

    // Sort row numbers in descending order to prevent index shift during sequential deletion
    rowsToDelete.sort((a, b) => b - a);

    // Build the requests array for a SINGLE batchUpdate call
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
      requestBody: {
        requests
      }
    });

    res.json({ success: true, deletedCount: rowsToDelete.length });
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

    app.listen(PORT, "0.0.0.0", () => {
      console.log(`Server running on http://localhost:${PORT}`);
    });
  }
}

startServer();

export default app;
