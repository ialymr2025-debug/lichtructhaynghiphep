import { RULES, SHIFTS } from '../constants';
import { xacDinhCa, timThay, isForbidden, shiftPenalty, buildConflict, fmtIn, timNghi } from './shiftHelpers';

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
  relievedTen?: string;
  relievedKip?: number;
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
  const reliefTracker: Record<number, { C: number, K: number, N: number, workingShiftsMissed: number, reliefsDone: number, reliefsReceivedByKip: Record<number, number>, lastCycleShift?: 'N' | 'C' | 'K', firstSwapType?: 'N' | 'C' | 'K' }> = {};
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
  // to determine eligibility for relief swaps (>= 3 covers = 1 relief)
  const kipCoverStats: Record<number, { N: number, C: number, K: number, total: number }> = {
    1: { N: 0, C: 0, K: 0, total: 0 },
    2: { N: 0, C: 0, K: 0, total: 0 },
    3: { N: 0, C: 0, K: 0, total: 0 },
    4: { N: 0, C: 0, K: 0, total: 0 },
    5: { N: 0, C: 0, K: 0, total: 0 }
  };
  Object.keys(allDates).forEach(dKey => {
    const d = allDates[dKey];
    const activeOnDay = leaves.filter(l => d >= l.start && d <= l.end);
    activeOnDay.forEach(l => {
      const absentKip = l.kip;
      const s = xacDinhCa(d, absentKip);
      if (s !== 'O' && RULES[absentKip] && RULES[absentKip][s]) {
        const coverer = RULES[absentKip][s].k;
        kipCoverStats[coverer].total++;
        if (s === 'N') kipCoverStats[coverer].N++;
        if (s === 'C') kipCoverStats[coverer].C++;
        if (s === 'K') kipCoverStats[coverer].K++;
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
    const forcedReliefs: Array<{ absentKip: number, shift: string, helperKip: number, relievedKip: number, relievedTen: string }> = [];

    activeLeaves.forEach(l => {
      const absentKip = l.kip;
      if (!reliefTracker[absentKip]) {
        reliefTracker[absentKip] = { 
          C: 0, K: 0, N: 0, 
          workingShiftsMissed: 0, 
          reliefsDone: 0, 
          reliefsReceivedByKip: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 },
          firstSwapType: undefined
        };
      }
      
      const s = xacDinhCa(ngay, absentKip);
      const helperKip = [1, 2, 3, 4, 5].find(k => k !== absentKip && k !== RULES[absentKip].N.k && k !== RULES[absentKip].C.k && k !== RULES[absentKip].K.k)!;

      if (s !== 'O') {
        // Increment tracker
        reliefTracker[absentKip].workingShiftsMissed++;
        if (s === 'C') reliefTracker[absentKip].C++;
        if (s === 'K') {
          reliefTracker[absentKip].K++;
        }
        if (s === 'N') reliefTracker[absentKip].N++;
      }

      // Relief logic: Only for 1-person leave
      // Requirement: Only start relief after at least 3 working shifts (N, C, K) have been missed
      // And the helper kip must be naturally Off today AND it must be an "O tròn" (Off after K)
      const isOTron = prevShift[helperKip] === 'K' && sh[helperKip] === 'O';
      
      // 1-person leave relief logic (Đổi ca)
      if (leaves.length === 1 && isOTron && !absentSet[helperKip]) {
        const requiredShiftsToday = new Set<string>();
        for (let k = 1; k <= 5; k++) {
          const st = xacDinhCa(ngay, k);
          if (st !== 'O') requiredShiftsToday.add(st);
        }

        let missedShiftToCompensate: 'N' | 'C' | 'K' | null = null;
        let reliefShift: 'N' | 'C' | 'K' | null = null;
        let targetKip: number | null = null;
        let targetRelief: 'N' | 'C' | 'K' | null = null;
        
        // Chu kỳ đổi ca: Lần 1 (K), Lần 2 (N/C), Lần 3 (N/C), Lần 4 (K)...
        const cycleIdx = reliefTracker[absentKip].reliefsDone % 3;

        // Calculate cover stats for THIS specific leave to ensure "trong suốt kỳ nghỉ" condition
        const currentLeaveCoverStats: Record<number, { N: number, C: number, K: number }> = {
          1: { N: 0, C: 0, K: 0 }, 2: { N: 0, C: 0, K: 0 }, 3: { N: 0, C: 0, K: 0 }, 4: { N: 0, C: 0, K: 0 }, 5: { N: 0, C: 0, K: 0 }
        };
        Object.keys(allDates).forEach(dKey => {
          const d = allDates[dKey];
          const s = xacDinhCa(d, absentKip);
          if (s !== 'O' && RULES[absentKip] && RULES[absentKip][s]) {
            const coverer = RULES[absentKip][s].k;
            if (s === 'N') currentLeaveCoverStats[coverer].N++;
            if (s === 'C') currentLeaveCoverStats[coverer].C++;
            if (s === 'K') currentLeaveCoverStats[coverer].K++;
          }
        });

        const kipN = RULES[absentKip].N.k;
        const kipC = RULES[absentKip].C.k;
        const kipK = RULES[absentKip].K.k;
        
        const countN = currentLeaveCoverStats[kipN].N;
        const countC = currentLeaveCoverStats[kipC].C;
        const countK = currentLeaveCoverStats[kipK].K;

        if (cycleIdx === 0) {
          // Lần 1: Kích hoạt khi kíp hỗ trợ đạt mốc 3 ca trực thay cùng loại (3N, 3C hoặc 3K) tính trong suốt kỳ nghỉ
          // Logic ưu tiên: 
          // 1. Nếu bất kỳ loại ca nào (N, C, K) đạt mốc >=4 ca trực thay: Ưu tiên đổi ca K lần đầu.
          // 2. Nếu không đạt mốc 4, nhưng đạt mốc 3:
          //    - Nếu có cả 3 ca N, 3 ca C và 3 ca K: Ưu tiên đổi ca K.
          //    - Nếu có cả 3 ca N và 3 ca C cùng lúc: Ưu tiên đổi ca C.
          //    - Nếu đạt mốc 3 ca C: Đổi ca C.
          //    - Nếu đạt mốc 3 ca N: Đổi ca N.
          //    - Nếu đạt mốc 3 ca K: Đổi ca K.
          // 3. Nếu không đạt mốc 4 và 3, nhưng đạt mốc 2:
          //    - Nếu có cả 2 ca K, 2 ca N và 2 ca C cùng lúc: Ưu tiên đổi ca K.
          //    - Nếu có cả 2 ca N và 2 ca C cùng lúc: Ưu tiên đổi ca C.
          //    - Nếu có cả 2 ca K và 2 ca N cùng lúc: Ưu tiên đổi ca K.
          //    - Nếu có cả 2 ca K và 2 ca C cùng lúc: Ưu tiên đổi ca K.
          //    - Nếu đạt mốc 2 ca C: Đổi ca C.
          //    - Nếu đạt mốc 2 ca N: Đổi ca N.
          //    - Nếu đạt mốc 2 ca K: Đổi ca K.
          if (countC >= 4 || countN >= 4 || countK >= 4) {
            targetRelief = 'K';
          } else if (countC >= 3 && countN >= 3 && countK >= 3) {
            targetRelief = 'K';
          } else if (countC >= 3 && countN >= 3) {
            targetRelief = 'C';
          } else if (countC >= 3) {
            targetRelief = 'C';
          } else if (countN >= 3) {
            targetRelief = 'N';
          } else if (countK >= 3) {
            targetRelief = 'K';
          } else if (countK >= 2 && countN >= 2 && countC >= 2) {
            targetRelief = 'K';
          } else if (countC >= 2 && countN >= 2) {
            targetRelief = 'C';
          } else if (countK >= 2 && countN >= 2) {
            targetRelief = 'K';
          } else if (countK >= 2 && countC >= 2) {
            targetRelief = 'K';
          } else if (countC >= 2) {
            targetRelief = 'C';
          } else if (countN >= 2) {
            targetRelief = 'N';
          } else if (countK >= 2) {
            targetRelief = 'K';
          }

          if (targetRelief) {
            missedShiftToCompensate = targetRelief;
            // Kíp được thay (targetKip) là kíp đã đạt mốc 3 ca trực thay
            const tk = RULES[absentKip][targetRelief].k;
            // Kíp hỗ trợ (helperKip) sẽ trực thay ca hiện tại của kíp đó
            if (sh[tk] !== 'O') {
              reliefShift = sh[tk] as 'N' | 'C' | 'K';
            }
          }
        } else if (cycleIdx === 1) {
          // Lần 2, 5, 8...: 
          // Nếu Lần 1 đã là K (do mốc >=4 hoặc 3N-3C-3K)
          if (reliefTracker[absentKip].firstSwapType === 'K') {
            // Xoay vòng N/C
            if (reliefTracker[absentKip].N > reliefTracker[absentKip].C) {
              missedShiftToCompensate = 'N';
              reliefShift = 'C';
            } else {
              missedShiftToCompensate = 'C';
              reliefShift = 'N';
            }
          } 
          // Nếu Lần 1 KHÔNG phải là K, nhưng đạt mốc 3N-3C hoặc 3C-3K -> Lần 2 là K
          else if ((countN >= 3 && countC >= 3) || (countC >= 3 && countK >= 3)) {
            missedShiftToCompensate = 'K';
            reliefShift = 'K';
          } else {
            // Nếu lần 1 là N hoặc C: Đổi ca còn lại trong cặp N-C
            if (reliefTracker[absentKip].firstSwapType === 'N') {
              missedShiftToCompensate = 'C';
              reliefShift = 'N';
            } else if (reliefTracker[absentKip].firstSwapType === 'C') {
              missedShiftToCompensate = 'N';
              reliefShift = 'C';
            } else {
              // Fallback
              if (reliefTracker[absentKip].N > reliefTracker[absentKip].C) {
                missedShiftToCompensate = 'N';
                reliefShift = 'C';
              } else {
                missedShiftToCompensate = 'C';
                reliefShift = 'N';
              }
            }
          }
        } else {
          // Lần 3, 6, 9...: Ca còn lại
          // Nếu lần 1 là K, lần 2 là N/C thì lần 3 là C/N
          if (reliefTracker[absentKip].firstSwapType === 'K') {
            if (reliefTracker[absentKip].lastCycleShift === 'N') {
              missedShiftToCompensate = 'C';
              reliefShift = 'N';
            } else {
              missedShiftToCompensate = 'N';
              reliefShift = 'C';
            }
          } else {
            // Nếu lần 1 là N hoặc C, lần 2 là C hoặc N, thì lần 3 là K
            // Nếu lần 2 đã bị ghi đè thành K (do mốc 3N-3C hoặc 3C-3K), thì lần 3 là ca còn lại của cặp N-C
            if (reliefTracker[absentKip].lastCycleShift === 'K') {
              if (reliefTracker[absentKip].firstSwapType === 'N') {
                missedShiftToCompensate = 'C';
                reliefShift = 'N';
              } else {
                missedShiftToCompensate = 'N';
                reliefShift = 'C';
              }
            } else {
              missedShiftToCompensate = 'K';
              reliefShift = 'K';
            }
          }
        }

        if (missedShiftToCompensate) {
          targetKip = RULES[absentKip][missedShiftToCompensate].k;
        }

        // Điều kiện: Kíp được thay (targetKip) phải đang có ca trực tự nhiên trùng với reliefShift
        if (targetKip && reliefShift && sh[targetKip] === reliefShift) {
          if (!absentSet[targetKip] && !forcedAssignments[helperKip] && !forcedAssignments[targetKip]) {
            // Điều kiện tích lũy theo yêu cầu người dùng:
            // Lần 1: Covered >= 2, Missed >= 2
            // Lần 2: Covered >= 3, Missed >= 5
            // Lần 3: Covered >= 4, Missed >= 7
            // Lần 4: Covered >= 5, Missed >= 9
            // Lần 5: Covered >= 6, Missed >= 11
            // Lần 6: Covered >= 7, Missed >= 13
            let canRelieve = false;
            const totalReliefs = reliefTracker[absentKip].reliefsDone;
            const totalMissed = reliefTracker[absentKip].workingShiftsMissed;
            const activeShiftType = (cycleIdx === 0) ? targetRelief : missedShiftToCompensate;
            const currentShiftCount = activeShiftType ? currentLeaveCoverStats[targetKip][activeShiftType] : 0;

            if (totalReliefs === 0) { // Lần 1
              if (currentShiftCount >= 2 && totalMissed >= 2) canRelieve = true;
            } else if (totalReliefs === 1) { // Lần 2
              if (currentShiftCount >= 3 && totalMissed >= 5) canRelieve = true;
            } else if (totalReliefs === 2) { // Lần 3
              if (currentShiftCount >= 4 && totalMissed >= 7) canRelieve = true;
            } else if (totalReliefs === 3) { // Lần 4
              if (currentShiftCount >= 5 && totalMissed >= 9) canRelieve = true;
            } else if (totalReliefs === 4) { // Lần 5
              if (currentShiftCount >= 6 && totalMissed >= 11) canRelieve = true;
            } else if (totalReliefs === 5) { // Lần 6
              if (currentShiftCount >= 7 && totalMissed >= 13) canRelieve = true;
            } else {
              // Từ lần 7 trở đi: Tiếp tục tăng dần số ca tích lũy và số ca nghỉ yêu cầu
              if (currentShiftCount >= (totalReliefs + 2) && totalMissed >= (totalReliefs * 2 + 3)) canRelieve = true;
            }

              if (canRelieve) {
                forcedAssignments[helperKip] = reliefShift;
                forcedAssignments[targetKip] = 'O';
                forcedReliefs.push({ 
                  absentKip, 
                  shift: reliefShift, 
                  helperKip, 
                  relievedKip: targetKip,
                  relievedTen: timThay(targetKip, chucDanh, staffData)
                });
                
                // Lưu lại loại ca đã đổi ở lần 1
                if (cycleIdx === 0) {
                  reliefTracker[absentKip].firstSwapType = missedShiftToCompensate as 'N' | 'C' | 'K';
                }
                
                // Lưu lại ca đã bù đắp ở lần 2 để lần 3 chọn ca còn lại
                if (cycleIdx === 1) {
                  reliefTracker[absentKip].lastCycleShift = missedShiftToCompensate as 'N' | 'C' | 'K';
                }
                
                reliefTracker[absentKip].reliefsReceivedByKip[targetKip]++;
                reliefTracker[absentKip].reliefsDone++;
              }
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

        const requiredShifts = new Set<string>();
        for (let k = 1; k <= 5; k++) {
          const s = xacDinhCa(ngay, k);
          if (s !== 'O') requiredShifts.add(s);
        }

        let score = 0;
        const assignedShifts = new Set(Object.values(fullConfig));
        requiredShifts.forEach(s => {
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
    const requiredShiftsToday = new Set<string>();
    for (let k = 1; k <= 5; k++) {
      const s = xacDinhCa(ngay, k);
      if (s !== 'O') requiredShiftsToday.add(s);
    }

    requiredShiftsToday.forEach(s => {
      if (!finalAssignedShifts.has(s)) {
        const absentKip = origAbsentKipMap[s];
        const tenAbsent = timThay(absentKip, chucDanh, staffData);
        extraRows.push({
          ngay, ca: s, kipThay: 0,
          nguoiThay: '⚠️ CHƯA CÓ NGƯỜI TRỰC',
          absentKip: absentKip, absentTen: tenAbsent, chucDanh,
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

          const reliefInfo = forcedReliefs.find(fr => fr.helperKip === k && fr.shift === assignedShift);
          const actualAbsentKip = reliefInfo ? reliefInfo.absentKip : absentKip;
          const currentAbsentTen = timThay(actualAbsentKip, chucDanh, staffData);
          const relievedTen = reliefInfo ? reliefInfo.relievedTen : null;
          const relievedKip = reliefInfo ? reliefInfo.relievedKip : null;

          if (absentSet[absentKip]) {
            const idx = kipToIdx[absentKip];
            if (idx !== undefined) {
              results[idx].ketQua.push({
                ngay, ca: assignedShift, kipThay: k,
                nguoiThay: timThay(k, chucDanh, staffData),
                relievedTen: relievedTen || undefined,
                relievedKip: relievedKip || undefined,
                isConflict: isConf,
                conflictNote: reliefInfo ? `${timThay(k, chucDanh, staffData)} trực thay ${relievedTen} ca ${assignedShift}` : (ckNote || noteConf),
                isOverlapDay: activeLeaves.length >= 2,
                isCKSwap: isCK,
                swapAbsentTen: reliefInfo ? currentAbsentTen : undefined
              });
              coverCount[k]++;
              accumulatedCoverCount[k]++;
            }
          } else if (isReplacement) {
            const isManualSwap = !absentSet[actualAbsentKip];
            extraRows.push({
              ngay, ca: assignedShift, kipThay: k,
              nguoiThay: timThay(k, chucDanh, staffData),
              absentKip: actualAbsentKip, absentTen: currentAbsentTen, chucDanh,
              relievedTen,
              relievedKip,
              isConflict: isConf, 
              conflictNote: ckNote || (isConf ? noteConf : (isManualSwap || reliefInfo ? `${timThay(k, chucDanh, staffData)} trực thay ${relievedTen || currentAbsentTen} ca ${assignedShift}` : `△ Điều chỉnh hệ thống: ${timThay(k, chucDanh, staffData)} thay cho ${currentAbsentTen}`)),
              isCKChain: isCK, isSwap: isManualSwap || !!reliefInfo, isOverlapDay: activeLeaves.length >= 2
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
