import "dotenv/config";
import { Pool } from "pg";
import fs from "fs";
import path from "path";

const connectionString = process.env.DATABASE_URL;

if (!connectionString) {
  console.warn("DATABASE_URL is not set. SQL leave storage will not work.");
}

// Supabase signs its Postgres endpoint with its own CA, which is not in Node's trust
// store, so plain `rejectUnauthorized: true` fails. Supply that CA and the connection
// becomes both encrypted and authenticated; without it TLS still encrypts, but any
// certificate is accepted, so anyone able to intercept the route to the database could
// read and rewrite every query. Download the CA from the Supabase dashboard
// (Project Settings > Database > SSL Configuration) and drop it in certs/ under any
// name, or point DATABASE_CA_FILE at it.
function findCaFile(): string | null {
  const dir = path.join(process.cwd(), "certs");
  if (!fs.existsSync(dir)) return null;
  const match = fs.readdirSync(dir).find(f => /\.(crt|pem|cer)$/i.test(f));
  return match ? path.join(dir, match) : null;
}

function buildSslConfig() {
  if (!connectionString) return undefined;

  // Accept whatever the CA file is called: Supabase's download keeps its own name
  // (prod-ca-2021.crt), so requiring one exact filename would silently fall back to
  // the unvalidated path even though the certificate was there all along.
  const caPath = process.env.DATABASE_CA_FILE || findCaFile();

  if (caPath && fs.existsSync(caPath)) {
    return { ca: fs.readFileSync(caPath, "utf8"), rejectUnauthorized: true };
  }

  console.warn([
    "",
    "[BAO MAT] Khong tim thay chung chi CA cua Supabase trong thu muc certs/",
    "          Ket noi CSDL van duoc ma hoa nhung KHONG xac thuc may chu.",
    "          Tai CA tai: Supabase Dashboard > Project Settings > Database > SSL Configuration",
    "          roi luu vao duong dan tren de bat xac thuc chung chi.",
    ""
  ].join(String.fromCharCode(10)));

  return { rejectUnauthorized: false };
}

export const pool = new Pool({
  connectionString,
  ssl: buildSslConfig(),
});

export async function initLeaveTables() {
  if (!connectionString) return;

  await pool.query(`
    CREATE TABLE IF NOT EXISTS leave_requests (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      birth_year TEXT,
      chuc_danh TEXT,
      kip TEXT,
      start_date TEXT,
      end_date TEXT,
      reason TEXT,
      phone TEXT,
      location TEXT,
      status TEXT,
      created_at TEXT,
      inserted_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);

  // Travel-allowance columns, added separately so existing tables get upgraded too
  await pool.query(`
    ALTER TABLE leave_requests
      ADD COLUMN IF NOT EXISTS leave_year TEXT,
      ADD COLUMN IF NOT EXISTS has_leave_permit BOOLEAN NOT NULL DEFAULT false,
      ADD COLUMN IF NOT EXISTS distance_km INTEGER,
      ADD COLUMN IF NOT EXISTS travel_days INTEGER NOT NULL DEFAULT 0;
  `);

  await pool.query(`
    ALTER TABLE leave_requests
      ADD COLUMN IF NOT EXISTS leave_days INTEGER NOT NULL DEFAULT 0;
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS leave_balances (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL UNIQUE,
      entitled TEXT,
      used TEXT,
      remaining TEXT,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);

  // Balances become per-person-per-year so leave carried over from the previous
  // year can coexist with the current year's allowance until 31/3.
  await pool.query(`ALTER TABLE leave_balances ADD COLUMN IF NOT EXISTS year TEXT;`);
  await pool.query(`UPDATE leave_balances SET year = $1 WHERE year IS NULL;`, [String(new Date().getFullYear())]);
  await pool.query(`ALTER TABLE leave_balances DROP CONSTRAINT IF EXISTS leave_balances_name_key;`);
  await pool.query(`
    DO $$ BEGIN
      IF NOT EXISTS (
        SELECT 1 FROM pg_constraint WHERE conname = 'leave_balances_name_year_key'
      ) THEN
        ALTER TABLE leave_balances ADD CONSTRAINT leave_balances_name_year_key UNIQUE (name, year);
      END IF;
    END $$;
  `);

  // Days taken outside this system — leave used before the company started keeping
  // requests here, or a correction an admin makes by hand. Kept apart from the days
  // summed out of leave_allocations rather than replacing them, so a request saved
  // after the correction still adds to the total instead of being swallowed by a
  // frozen figure. A row may carry this with entitled left NULL: the two overrides
  // are independent, and setting one must not turn the other on.
  await pool.query(`ALTER TABLE leave_balances ADD COLUMN IF NOT EXISTS used_adjust TEXT;`);

  // Entitlement is derived from hire year, so it grows on its own as seniority accrues.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS employees (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL UNIQUE,
      hire_year INTEGER NOT NULL,
      base_days INTEGER NOT NULL DEFAULT 16,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);

  // One leave request can draw days from two leave years, so allocations live in their own table.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS leave_allocations (
      id SERIAL PRIMARY KEY,
      leave_id TEXT NOT NULL REFERENCES leave_requests(id) ON DELETE CASCADE,
      leave_year TEXT NOT NULL,
      days INTEGER NOT NULL
    );
  `);
  await pool.query(`CREATE INDEX IF NOT EXISTS leave_allocations_leave_id_idx ON leave_allocations(leave_id);`);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS location_distances (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL UNIQUE,
      distance_km INTEGER NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);

  // Replaces the Firestore "config" collection: google_auth tokens and app_settings.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS app_config (
      key TEXT PRIMARY KEY,
      value JSONB NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);

  // Replaces the Firestore "signatures" collection; data is a base64 image data URI.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS signatures (
      name TEXT PRIMARY KEY,
      data TEXT NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);

  // Backs the multi-workshop login/admin system (LoginForm, UserHeaderBar, WorkshopManagerModal),
  // which existed as unwired frontend components before this. Each workshop owns its own
  // staff list + config, independent of the single global app_config used by the rest of the app.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS workshops (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      code TEXT NOT NULL,
      description TEXT NOT NULL DEFAULT '',
      staff_data JSONB NOT NULL DEFAULT '[]',
      config JSONB NOT NULL DEFAULT '{}',
      features JSONB NOT NULL DEFAULT '{}',
      created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS user_accounts (
      id TEXT PRIMARY KEY,
      username TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      full_name TEXT NOT NULL,
      role TEXT NOT NULL,
      workshop_id TEXT REFERENCES workshops(id) ON DELETE CASCADE,
      created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
      updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
    );
  `);

  // Answers "who changed this, and when". Written for every mutating API call; the
  // actor is taken from the session, never from the request, so it cannot be forged.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS audit_log (
      id BIGSERIAL PRIMARY KEY,
      at TIMESTAMPTZ NOT NULL DEFAULT now(),
      user_id TEXT,
      username TEXT,
      role TEXT,
      workshop_id TEXT,
      method TEXT NOT NULL,
      path TEXT NOT NULL,
      status INTEGER,
      summary TEXT,
      ip TEXT
    );
  `);
  await pool.query(`CREATE INDEX IF NOT EXISTS audit_log_at_idx ON audit_log(at DESC);`);
  await pool.query(`CREATE INDEX IF NOT EXISTS audit_log_workshop_idx ON audit_log(workshop_id);`);

  // Server-side sessions. Only the SHA-256 of the session token is stored, so a
  // database leak cannot be replayed as a valid login. Rows are removed on logout
  // and lazily pruned once expired.
  await pool.query(`
    CREATE TABLE IF NOT EXISTS user_sessions (
      token_hash TEXT PRIMARY KEY,
      user_id TEXT NOT NULL REFERENCES user_accounts(id) ON DELETE CASCADE,
      created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
      expires_at TIMESTAMPTZ NOT NULL
    );
  `);
  await pool.query(`CREATE INDEX IF NOT EXISTS user_sessions_user_id_idx ON user_sessions(user_id);`);

  // Migration: user_accounts.workshop_id used to be ON DELETE SET NULL. Deleting a
  // workshop then left its workshop_admin accounts alive with workshop_id = NULL, which
  // toUserAccount() reads back as workshopId "all" — the same value used for a real
  // super admin's unrestricted scope. scopeWorkshopId() let those orphaned accounts
  // through as if they were super admins, so a deleted workshop's admin could still
  // read every other workshop's staff and signatures. CASCADE removes the account
  // instead of orphaning it; super_admin rows (workshop_id already NULL by design,
  // never referencing a row) are untouched by this since there is nothing to cascade.
  await pool.query(`
    DO $$
    BEGIN
      IF EXISTS (
        SELECT 1 FROM pg_constraint
        WHERE conname = 'user_accounts_workshop_id_fkey' AND confdeltype <> 'c'
      ) THEN
        ALTER TABLE user_accounts DROP CONSTRAINT user_accounts_workshop_id_fkey;
        ALTER TABLE user_accounts
          ADD CONSTRAINT user_accounts_workshop_id_fkey
          FOREIGN KEY (workshop_id) REFERENCES workshops(id) ON DELETE CASCADE;
      END IF;
    END $$;
  `);

  await scopeDataPerWorkshop();
  await seedLocationDistances();

  console.log("SQL tables ready (leave_requests, leave_allocations, leave_balances, employees, location_distances, app_config, signatures, workshops, user_accounts).");
}

// Give employees/leave_requests/leave_balances/signatures a workshop_id so multiple
// workshops each get their own staff, leave history and signatures instead of sharing one
// global pool. Existing rows (from before workshops existed) are backfilled onto the
// oldest workshop so nothing already entered appears to vanish.
async function scopeDataPerWorkshop() {
  const fallbackWs = await pool.query(`SELECT id FROM workshops ORDER BY created_at ASC LIMIT 1`);
  const fallbackWorkshopId = fallbackWs.rows[0]?.id || null;

  // employees: name UNIQUE -> (workshop_id, name) UNIQUE
  await pool.query(`ALTER TABLE employees ADD COLUMN IF NOT EXISTS workshop_id TEXT;`);
  if (fallbackWorkshopId) {
    await pool.query(`UPDATE employees SET workshop_id = $1 WHERE workshop_id IS NULL`, [fallbackWorkshopId]);
  }
  await pool.query(`ALTER TABLE employees DROP CONSTRAINT IF EXISTS employees_name_key;`);
  await pool.query(`
    DO $$ BEGIN
      IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'employees_workshop_name_key') THEN
        ALTER TABLE employees ADD CONSTRAINT employees_workshop_name_key UNIQUE (workshop_id, name);
      END IF;
    END $$;
  `);
  await pool.query(`
    DO $$ BEGIN
      IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'employees_workshop_id_fkey') THEN
        ALTER TABLE employees ADD CONSTRAINT employees_workshop_id_fkey FOREIGN KEY (workshop_id) REFERENCES workshops(id) ON DELETE CASCADE;
      END IF;
    END $$;
  `);

  // leave_balances: (name, year) UNIQUE -> (workshop_id, name, year) UNIQUE
  await pool.query(`ALTER TABLE leave_balances ADD COLUMN IF NOT EXISTS workshop_id TEXT;`);
  if (fallbackWorkshopId) {
    await pool.query(`UPDATE leave_balances SET workshop_id = $1 WHERE workshop_id IS NULL`, [fallbackWorkshopId]);
  }
  await pool.query(`ALTER TABLE leave_balances DROP CONSTRAINT IF EXISTS leave_balances_name_year_key;`);
  await pool.query(`
    DO $$ BEGIN
      IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'leave_balances_workshop_name_year_key') THEN
        ALTER TABLE leave_balances ADD CONSTRAINT leave_balances_workshop_name_year_key UNIQUE (workshop_id, name, year);
      END IF;
    END $$;
  `);
  await pool.query(`
    DO $$ BEGIN
      IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'leave_balances_workshop_id_fkey') THEN
        ALTER TABLE leave_balances ADD CONSTRAINT leave_balances_workshop_id_fkey FOREIGN KEY (workshop_id) REFERENCES workshops(id) ON DELETE CASCADE;
      END IF;
    END $$;
  `);

  // leave_requests: id is already a globally unique PK, just tag ownership.
  // leave_allocations inherits scoping transitively via its leave_id FK.
  await pool.query(`ALTER TABLE leave_requests ADD COLUMN IF NOT EXISTS workshop_id TEXT;`);
  if (fallbackWorkshopId) {
    await pool.query(`UPDATE leave_requests SET workshop_id = $1 WHERE workshop_id IS NULL`, [fallbackWorkshopId]);
  }
  await pool.query(`
    DO $$ BEGIN
      IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'leave_requests_workshop_id_fkey') THEN
        ALTER TABLE leave_requests ADD CONSTRAINT leave_requests_workshop_id_fkey FOREIGN KEY (workshop_id) REFERENCES workshops(id) ON DELETE CASCADE;
      END IF;
    END $$;
  `);

  // signatures: name PK -> (workshop_id, name) UNIQUE, id SERIAL as the real PK
  await pool.query(`ALTER TABLE signatures ADD COLUMN IF NOT EXISTS workshop_id TEXT;`);
  if (fallbackWorkshopId) {
    await pool.query(`UPDATE signatures SET workshop_id = $1 WHERE workshop_id IS NULL`, [fallbackWorkshopId]);
  }
  const hasIdCol = await pool.query(`SELECT 1 FROM information_schema.columns WHERE table_name = 'signatures' AND column_name = 'id'`);
  if (hasIdCol.rows.length === 0) {
    await pool.query(`ALTER TABLE signatures ADD COLUMN id SERIAL;`);
    await pool.query(`ALTER TABLE signatures DROP CONSTRAINT IF EXISTS signatures_pkey;`);
    await pool.query(`ALTER TABLE signatures ADD PRIMARY KEY (id);`);
  }
  await pool.query(`
    DO $$ BEGIN
      IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'signatures_workshop_name_key') THEN
        ALTER TABLE signatures ADD CONSTRAINT signatures_workshop_name_key UNIQUE (workshop_id, name);
      END IF;
    END $$;
  `);
  await pool.query(`
    DO $$ BEGIN
      IF NOT EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'signatures_workshop_id_fkey') THEN
        ALTER TABLE signatures ADD CONSTRAINT signatures_workshop_id_fkey FOREIGN KEY (workshop_id) REFERENCES workshops(id) ON DELETE CASCADE;
      END IF;
    END $$;
  `);
}

// Approximate road distances from Pleiku (Gia Lai) in kilometres.
// Seeded once; edit the rows in Supabase Table Editor to correct them.
const PLEIKU_DISTANCES: [string, number][] = [
  ["Gia Lai", 0],
  ["Kon Tum", 50],
  ["Bình Định", 165],
  ["Đắk Lắk", 190],
  ["Phú Yên", 200],
  ["Quảng Ngãi", 275],
  ["Khánh Hòa", 320],
  ["Đắk Nông", 320],
  ["Quảng Nam", 340],
  ["Đà Nẵng", 400],
  ["Lâm Đồng", 400],
  ["Ninh Thuận", 420],
  ["Bình Phước", 450],
  ["Bình Thuận", 500],
  ["Bình Dương", 500],
  ["Huế", 500],
  ["Thừa Thiên Huế", 500],
  ["Đồng Nai", 520],
  ["TP. Hồ Chí Minh", 550],
  ["Quảng Trị", 570],
  ["Bà Rịa - Vũng Tàu", 590],
  ["Tây Ninh", 600],
  ["Long An", 600],
  ["Tiền Giang", 620],
  ["Bến Tre", 640],
  ["Quảng Bình", 670],
  ["Vĩnh Long", 670],
  ["Đồng Tháp", 690],
  ["Trà Vinh", 700],
  ["Cần Thơ", 720],
  ["An Giang", 780],
  ["Hậu Giang", 780],
  ["Sóc Trăng", 780],
  ["Hà Tĩnh", 800],
  ["Kiên Giang", 830],
  ["Nghệ An", 850],
  ["Bạc Liêu", 850],
  ["Cà Mau", 900],
  ["Thanh Hóa", 990],
  ["Hà Nội", 1080],
  ["Ninh Bình", 1080],
  ["Hòa Bình", 1080],
  ["Hà Nam", 1090],
  ["Nam Định", 1100],
  ["Hưng Yên", 1110],
  ["Bắc Ninh", 1120],
  ["Vĩnh Phúc", 1130],
  ["Thái Bình", 1140],
  ["Hải Dương", 1150],
  ["Phú Thọ", 1160],
  ["Bắc Giang", 1160],
  ["Thái Nguyên", 1160],
  ["Hải Phòng", 1180],
  ["Tuyên Quang", 1220],
  ["Yên Bái", 1230],
  ["Bắc Kạn", 1240],
  ["Lạng Sơn", 1240],
  ["Quảng Ninh", 1250],
  ["Sơn La", 1300],
  ["Lào Cai", 1350],
  ["Cao Bằng", 1370],
  ["Hà Giang", 1380],
  ["Điện Biên", 1470],
  ["Lai Châu", 1480],
];

async function seedLocationDistances() {
  // ON CONFLICT DO NOTHING so manual corrections in Supabase are never overwritten
  const names = PLEIKU_DISTANCES.map(d => d[0]);
  const kms = PLEIKU_DISTANCES.map(d => d[1]);
  await pool.query(
    `INSERT INTO location_distances (name, distance_km)
     SELECT * FROM UNNEST($1::text[], $2::int[])
     ON CONFLICT (name) DO NOTHING`,
    [names, kms]
  );
}

// Seniority earns one extra leave day per completed 5 years of service.
export const SENIORITY_STEP_YEARS = 5;

export function calcSeniority(hireYear: number, forYear: number): number {
  return Math.max(0, forYear - hireYear);
}

export function calcEntitledDays(baseDays: number, hireYear: number, forYear: number): number {
  return baseDays + Math.floor(calcSeniority(hireYear, forYear) / SENIORITY_STEP_YEARS);
}

// Travel days granted on top of annual leave, per the distance brackets in the regulation.
export function calcTravelDays(distanceKm: number | null | undefined): number {
  if (distanceKm === null || distanceKm === undefined || isNaN(distanceKm)) return 0;
  if (distanceKm >= 1500) return 4;
  if (distanceKm >= 1000) return 3;
  if (distanceKm >= 500) return 2;
  if (distanceKm >= 200) return 1;
  return 0;
}
