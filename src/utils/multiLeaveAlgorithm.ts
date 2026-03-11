import { RULES, SHIFTS } from '../constants';
import { xacDinhCa, timThay, isForbidden, shiftPenalty, buildConflict, fmtIn } from './shiftHelpers';

export interface Leave {
  kip: number;
  start: Date;
  end: Date;
  ten: string;
  chucDanh: string;
}

export interface ResultItem {
  ngay: Date;
  ca: string;
  kipThay: number;
  nguoiThay: string;
  isConflict: boolean;
  conflictNote?: string;
  isOverlapDay?: boolean;
  isCKSwap?: boolean;
  swapAbsentTen?: string;
}

export function buildMultiLeaveResults(leaves: Leave[], chucDanh: string, staffData: string[][]) {
  const results = leaves.map(l => ({
    ten: l.ten,
    kip: l.kip,
    start: l.start,
    end: l.end,
    chucDanh: l.chucDanh || '',
    ketQua: [] as ResultItem[]
  }));

  const kipToIdx: Record<number, number> = {};
  leaves.forEach((l, i) => { kipToIdx[l.kip] = i; });

  const coverCount: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  const accumulatedCoverCount: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  const reliefTracker: Record<number, { C: number, K: number, N: number, workingShiftsMissed: number, reliefsDone: number, lastCycleShift?: 'N' | 'C' }> = {};
  const kShiftCounts: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  const dayShifts: Record<string, Record<number, string | undefined>> = {};
  const blockedNextK: Record<string, number[]> = {};
  const blockedNextKMeta: Record<string, number> = {};
  const extraRows: any[] = [];
  let hasConflict = false;

  const allDates: Record<string, Date> = {};
  leaves.forEach(l => {
    let d = new Date(l.start);
    while (d <= l.end) {
      allDates[fmtIn(d)] = new Date(d);
      d.setDate(d.getDate() + 1);
    }
  });

  const processedDates: Record<string, boolean> = {};

  // Pre-calculate total covers for each kip during the entire leave period
  // to determine eligibility for relief swaps (>= 2 covers = 1 relief)
  const totalPlannedCovers: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  Object.keys(allDates).forEach(dKey => {
    const d = allDates[dKey];
    const activeOnDay = leaves.filter(l => d >= l.start && d <= l.end);
    activeOnDay.forEach(l => {
      const absentKip = l.kip;
      const s = xacDinhCa(d, absentKip);
      if (s !== 'O' && RULES[absentKip] && RULES[absentKip][s]) {
        const coverer = RULES[absentKip][s].k;
        totalPlannedCovers[coverer]++;
      }
    });
  });

  function getNextUnprocessed() {
    const keys = Object.keys(allDates).sort();
    for (let i = 0; i < keys.length; i++) if (!processedDates[keys[i]]) return keys[i];
    return null;
  }

  let dateKey: string | null;
  while ((dateKey = getNextUnprocessed()) !== null) {
    processedDates[dateKey] = true;
    const ngay = allDates[dateKey];
    const tomorrow = new Date(ngay.getTime() + 86400000);
    const prevKey = fmtIn(new Date(ngay.getTime() - 86400000));

    const activeLeaves = leaves.filter(l => ngay >= l.start && ngay <= l.end);
    const absentSet: Record<number, boolean> = {};
    activeLeaves.forEach(l => { absentSet[l.kip] = true; });
    const availKips = [1, 2, 3, 4, 5].filter(k => !absentSet[k]);

    const sh: Record<number, string> = {};
    const shNext: Record<number, string> = {};
    for (let k = 1; k <= 5; k++) {
      sh[k] = xacDinhCa(ngay, k);
      shNext[k] = xacDinhCa(tomorrow, k);
    }

    const prevDay = new Date(ngay.getTime() - 86400000);
    const prevPrevDay = new Date(ngay.getTime() - 172800000);
    const prevShift: Record<number, string> = {};
    const prevPrevShift: Record<number, string> = {};
    for (let k = 1; k <= 5; k++) {
      prevShift[k] = (dayShifts[prevKey] && dayShifts[prevKey][k] != null)
        ? (dayShifts[prevKey][k] as string)
        : xacDinhCa(prevDay, k);
      
      const prevPrevKey = fmtIn(prevPrevDay);
      prevPrevShift[k] = (dayShifts[prevPrevKey] && dayShifts[prevPrevKey][k] != null)
        ? (dayShifts[prevPrevKey][k] as string)
        : xacDinhCa(prevPrevDay, k);
    }
    const tomorrowKey = fmtIn(tomorrow);
    
    if (!dayShifts[dateKey]) dayShifts[dateKey] = {};

    function getNextActual(kip: number, offset: number = 1) {
      const targetDate = new Date(ngay.getTime() + 86400000 * offset);
      const targetKey = fmtIn(targetDate);
      
      const isOnLeave = leaves.some(l => targetDate >= l.start && targetDate <= l.end && l.kip === kip);
      if (isOnLeave) return 'O';

      if (dayShifts[targetKey] && dayShifts[targetKey][kip])
        return dayShifts[targetKey][kip];
      const tmrBlocked = blockedNextK[targetKey] || [];
      if (tmrBlocked.indexOf(kip) !== -1) return 'O';
      return xacDinhCa(targetDate, kip);
    }

    // Map to track who is originally in which shift today
    const origAbsentKipMap: Record<string, number> = {};
    ['N', 'C', 'K', 'O'].forEach(s => {
      for (let k = 1; k <= 5; k++) {
        if (sh[k] === s) {
          origAbsentKipMap[s] = k;
        }
      }
    });

    // Track relief progress and determine if a swap is needed today
    const forcedAssignments: Record<number, string> = {};
    const forcedReliefs: Array<{ absentKip: number, shift: string, helperKip: number }> = [];

    activeLeaves.forEach(l => {
      const absentKip = l.kip;
      if (!reliefTracker[absentKip]) {
        reliefTracker[absentKip] = { C: 0, K: 0, N: 0, workingShiftsMissed: 0, reliefsDone: 0 };
      }
      
      const s = xacDinhCa(ngay, absentKip);
      const helperKip = [1, 2, 3, 4, 5].find(k => k !== absentKip && k !== RULES[absentKip].N.k && k !== RULES[absentKip].C.k && k !== RULES[absentKip].K.k)!;

      if (s !== 'O') {
        // Increment tracker
        reliefTracker[absentKip].workingShiftsMissed++;
        if (s === 'C') reliefTracker[absentKip].C++;
        if (s === 'K') {
          reliefTracker[absentKip].K++;
          kShiftCounts[absentKip]++;
        }
        if (s === 'N') reliefTracker[absentKip].N++;
      } else {
        // Relief logic: Only for 1-person leave
        // Requirement: Only start relief after at least 3 working shifts (N, C, K) have been missed
        if (activeLeaves.length === 1 && reliefTracker[absentKip].workingShiftsMissed >= 3 && sh[helperKip] === 'O' && !absentSet[helperKip]) {
          const cycleIdx = reliefTracker[absentKip].reliefsDone % 3;
          let targetShift: 'N' | 'C' | 'K' = 'K';
          
          if (cycleIdx === 0) {
            targetShift = 'K';
          } else if (cycleIdx === 1) {
            // Compare N and C missed counts to decide next relief
            if (reliefTracker[absentKip].N >= reliefTracker[absentKip].C) {
              targetShift = 'N';
            } else {
              targetShift = 'C';
            }
            reliefTracker[absentKip].lastCycleShift = targetShift;
          } else {
            // cycleIdx === 2: Pick the remaining shift from N and C
            targetShift = (reliefTracker[absentKip].lastCycleShift === 'N') ? 'C' : 'N';
          }

          const kipToRelieve = [1, 2, 3, 4, 5].find(k => sh[k] === targetShift);
          if (kipToRelieve && !forcedAssignments[helperKip] && !forcedAssignments[kipToRelieve]) {
            forcedAssignments[helperKip] = targetShift;
            forcedAssignments[kipToRelieve] = 'O';
            forcedReliefs.push({ absentKip, shift: targetShift, helperKip });
            reliefTracker[absentKip].reliefsDone++;
          }
        }
      }
    });

    let bestScore = Infinity;
    let bestConfig: Record<number, string> = {};

    function solve(shiftIdx: number, usedPeople: Set<number>, current: Record<number, string>, currentAvail: number[]) {
      if (shiftIdx === 3) {
        const fullConfig: Record<number, string> = { ...current };
        // Fill in 'O' for anyone not assigned and not already forced
        [1, 2, 3, 4, 5].forEach(k => {
          if (!fullConfig[k]) fullConfig[k] = 'O';
        });

        let score = 0;
        const assignedShifts = new Set(Object.values(fullConfig));
        ['N', 'C', 'K'].forEach(s => {
          if (!assignedShifts.has(s)) score += 1000000; 
        });

        for (let k = 1; k <= 5; k++) {
          const s = fullConfig[k];
          const naturalS = sh[k];
          
          if (s === 'O') {
            if (naturalS !== 'O' && !absentSet[k]) {
              // Naturally working but assigned Off - massive penalty unless forced by relief
              score += 1000000; 
            }
            continue;
          }
          
          const origKip = origAbsentKipMap[s];
          const isOrigAbsent = absentSet[origKip];

          if (k !== origKip) {
            if (isOrigAbsent) {
              // Covering a leave - Good, but prefer rule-based person
              score += 1000;
              let ruleKip = (origKip && RULES[origKip] && RULES[origKip][s]) ? RULES[origKip][s].k : null;
              
              // Apply K rotation rule
              if (s === 'K' && origKip) {
                const count = kShiftCounts[origKip];
                if (count % 3 === 2) ruleKip = RULES[origKip].N.k;
                else if (count % 3 === 0 && count > 0) ruleKip = RULES[origKip].C.k;
              }

              if (ruleKip === k) score -= 500;
            } else {
              // Stealing a shift from someone who is NOT on leave - Massive penalty
              score += 2000000;
            }
          }

          const fb = isForbidden(k, s, prevShift, getNextActual, 'O', prevPrevShift);
          if (fb.bad) score += 5000000; 
          
          if (s === 'K' && prevShift[k] === 'N') score -= 30;
          if (s === 'K' && prevShift[k] === 'K') score += 100;
        }

        if (score < bestScore) {
          bestScore = score;
          bestConfig = fullConfig;
        }
        return;
      }

      const s = ['N', 'C', 'K'][shiftIdx];
      
      // If this shift is already forced, skip to next shift
      const forcedKip = Object.keys(forcedAssignments).find(k => forcedAssignments[Number(k)] === s);
      if (forcedKip) {
        solve(shiftIdx + 1, usedPeople, current, currentAvail);
        return;
      }

      let assigned = false;
      for (const p of currentAvail) {
        if (!usedPeople.has(p)) {
          usedPeople.add(p);
          current[p] = s;
          solve(shiftIdx + 1, usedPeople, current, currentAvail);
          delete current[p];
          usedPeople.delete(p);
          assigned = true;
        }
      }
      if (!assigned || currentAvail.length < 3) {
         solve(shiftIdx + 1, usedPeople, current, currentAvail);
      }
    }

    // Initialize solver with forced assignments
    const initialUsed = new Set<number>();
    const initialCurrent: Record<number, string> = {};
    const effectiveAvailKips = availKips.filter(k => {
      if (forcedAssignments[k] === 'O') {
        initialCurrent[k] = 'O';
        return false;
      }
      if (forcedAssignments[k]) {
        initialUsed.add(k);
        initialCurrent[k] = forcedAssignments[k];
        return false;
      }
      return true;
    });

    solve(0, initialUsed, initialCurrent, effectiveAvailKips);

    if (bestScore === Infinity) {
      bestConfig = {};
      for (let k = 1; k <= 5; k++) bestConfig[k] = absentSet[k] ? 'O' : sh[k];
    }

    // Check for missing shifts and report them
    const finalAssignedShifts = new Set(Object.values(bestConfig));
    ['N', 'C', 'K'].forEach(s => {
      if (!finalAssignedShifts.has(s)) {
        const absentKip = origAbsentKipMap[s];
        const tenAbsent = timThay(absentKip, chucDanh, staffData);
        extraRows.push({
          ngay, ca: s, kipThay: 0,
          nguoiThay: '⚠️ CHƯA CÓ NGƯỜI TRỰC',
          absentKip: absentKip, absentTen: tenAbsent,
          isConflict: true, conflictNote: `Không tìm được người thay cho Ca ${s} của ${tenAbsent}`,
          isCKChain: false, isSwap: false, isOverlapDay: activeLeaves.length >= 2
        });
        hasConflict = true;
      }
    });

    for (let k = 1; k <= 5; k++) {
      const assignedShift = bestConfig[k];
      dayShifts[dateKey!][k] = assignedShift;
      
      if (assignedShift !== 'O') {
        const naturalShift = sh[k];
        const isReplacement = assignedShift !== naturalShift;
        const absentKip = origAbsentKipMap[assignedShift];

        if (isReplacement || absentSet[absentKip]) {
          const fb = isForbidden(k, assignedShift, prevShift, getNextActual, 'O', prevPrevShift);
          const isConf = fb.bad;
          const noteConf = isConf ? (fb.note || '⚠ Vi phạm ràng buộc ca') : '';
          
          const isCK = prevShift[k] === 'C' && assignedShift === 'O' && sh[k] === 'K';
          const ckNote = isCK ? `⥵ C→K: ${timThay(k, chucDanh, staffData)} vướng ca C hôm trước, không thể trực ca K hôm nay` : '';

          if (absentSet[absentKip]) {
            const idx = kipToIdx[absentKip];
            if (idx !== undefined) {
              const reliefInfo = forcedReliefs.find(r => r.helperKip === k && r.shift === assignedShift && r.absentKip === absentKip);

              results[idx].ketQua.push({
                ngay, ca: assignedShift, kipThay: k,
                nguoiThay: timThay(k, chucDanh, staffData),
                isConflict: isConf,
                conflictNote: reliefInfo ? `⇄ Đổi ca: Kíp ${k} trực thay ca ${assignedShift} cho Kíp ${absentKip} (Kíp hỗ trợ)` : (ckNote || noteConf),
                isOverlapDay: activeLeaves.length >= 2,
                isCKSwap: isCK
              });
              coverCount[k]++;
              accumulatedCoverCount[k]++;
            }
          } else if (isReplacement) {
            extraRows.push({
              ngay, ca: assignedShift, kipThay: k,
              nguoiThay: timThay(k, chucDanh, staffData),
              absentKip: absentKip, absentTen: timThay(absentKip, chucDanh, staffData),
              isConflict: isConf, 
              conflictNote: ckNote || (isConf ? noteConf : `△ Điều chỉnh hệ thống: ${timThay(k, chucDanh, staffData)} thay cho ${timThay(absentKip, chucDanh, staffData)}`),
              isCKChain: isCK, isSwap: true, isOverlapDay: activeLeaves.length >= 2
            });
            coverCount[k]++;
          }
        }

        if (assignedShift === 'C') {
          if (!blockedNextK[tomorrowKey]) blockedNextK[tomorrowKey] = [];
          if (blockedNextK[tomorrowKey].indexOf(k) === -1) blockedNextK[tomorrowKey].push(k);
          
          // Only add tomorrow to allDates if it's within the original leave range + 1 day
          // AND the person is actually blocked from their natural shift tomorrow.
          const maxLeaveEnd = leaves.length > 0 ? Math.max(...leaves.map(l => l.end.getTime())) : 0;
          const maxDate = new Date(maxLeaveEnd + 86400000);
          const naturalShiftTomorrow = xacDinhCa(tomorrow, k);
          
          if (!allDates[tomorrowKey] && tomorrow <= maxDate && naturalShiftTomorrow === 'K') {
            allDates[tomorrowKey] = new Date(tomorrow);
          }
        }
      }
    }
    
    if (bestScore >= 10000 && bestScore < 1000000) hasConflict = true;
  }

  return { results, extraRows, hasConflict, coverCount };
}
