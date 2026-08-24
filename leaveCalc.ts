import { SHIFTS, BASE_DATE } from "./src/constants";

export interface ShiftConfig {
  baseDate?: string;
  shiftsMatrix?: string[][];
}

// Server-side mirror of xacDinhCa() in src/utils/shiftHelpers.ts.
export function xacDinhCaServer(date: Date, kip: number, cfg?: ShiftConfig): string {
  let baseDateObj = BASE_DATE;
  if (cfg?.baseDate) {
    const parsed = new Date(cfg.baseDate + (cfg.baseDate.includes("T") ? "" : "T00:00:00"));
    if (!isNaN(parsed.getTime())) baseDateObj = parsed;
  }

  const shifts = (cfg?.shiftsMatrix && cfg.shiftsMatrix.length > 0) ? cfg.shiftsMatrix : SHIFTS;
  const diff = Math.floor((date.getTime() - baseDateObj.getTime()) / 86400000);
  const cycleLen = shifts[0]?.length || 5;
  const kipIdx = ((kip - 1) % shifts.length + shifts.length) % shifts.length;
  return shifts[kipIdx][((diff % cycleLen) + cycleLen) % cycleLen] || "O";
}

export interface WorkingShift {
  date: Date;
  shift: string;
}

// A leave day is only consumed on a date the team actually stands a shift (N/C/K).
// Days off in the rota ("O") fall inside the leave period but cost nothing.
export function listWorkingShifts(startDate: string, endDate: string, kip: number, cfg?: ShiftConfig): WorkingShift[] {
  const out: WorkingShift[] = [];
  const start = new Date(startDate + "T00:00:00");
  const end = new Date(endDate + "T00:00:00");
  if (isNaN(start.getTime()) || isNaN(end.getTime()) || end < start) return out;

  for (let d = new Date(start); d <= end; d = new Date(d.getTime() + 86400000)) {
    const shift = xacDinhCaServer(d, kip, cfg);
    if (shift && shift !== "O") out.push({ date: new Date(d), shift });
  }
  return out;
}

// Leave from year Y stays usable until 31/3 of year Y+1; from 1/4 only the new year counts.
export const CARRY_OVER_LAST_MONTH = 3;

export interface AllocationResult {
  allocations: Record<string, number>;
  detail: { date: string; shift: string; year: string }[];
}

// Walk the shifts in order, charging each to the old leave year while it still has
// days left (only up to 31/3), otherwise to the year the shift falls in.
export function allocateLeaveDays(
  shifts: WorkingShift[],
  quotaRemaining: Record<string, number>
): AllocationResult {
  const remaining = { ...quotaRemaining };
  const allocations: Record<string, number> = {};
  const detail: AllocationResult["detail"] = [];

  for (const ws of shifts) {
    const year = ws.date.getFullYear();
    const month = ws.date.getMonth() + 1;
    const prevYear = String(year - 1);

    let target = String(year);
    if (month <= CARRY_OVER_LAST_MONTH && (remaining[prevYear] ?? 0) > 0) {
      target = prevYear;
    }

    allocations[target] = (allocations[target] || 0) + 1;
    remaining[target] = (remaining[target] ?? 0) - 1;
    detail.push({
      date: `${year}-${String(month).padStart(2, "0")}-${String(ws.date.getDate()).padStart(2, "0")}`,
      shift: ws.shift,
      year: target
    });
  }

  return { allocations, detail };
}
