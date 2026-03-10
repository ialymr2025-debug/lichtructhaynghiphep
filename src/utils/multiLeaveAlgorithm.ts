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
  const coverKCount: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  const coverCCount: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
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

  let totalC = 0;
  let totalK = 0;
  if (leaves.length === 1) {
    const l = leaves[0];
    const d = new Date(l.start);
    while (d <= l.end) {
      const s = xacDinhCa(d, l.kip);
      if (s === 'C') totalC++;
      if (s === 'K') totalK++;
      d.setDate(d.getDate() + 1);
    }
  }

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

    // Balancing logic: If only 1 person is on leave, check if their K-replacement needs relief
    if (activeLeaves.length === 1) {
      const absentKipId = activeLeaves[0].kip;
      const designatedKipId = RULES[absentKipId].K.k;
      const kCount = coverKCount[designatedKipId];
      const cCountCovered = coverCCount[designatedKipId];

      if (totalC >= 3) {
        // Case 2: >= 3 C shifts
        // Relief for N after 2nd C replacement
        if (sh[designatedKipId] === 'N' && cCountCovered >= 2) {
          origAbsentKipMap['N'] = designatedKipId;
        }
        // Relief for K before 3rd K replacement
        if (sh[designatedKipId] === 'K' && kCount >= 2) {
          origAbsentKipMap['K'] = designatedKipId;
        }
      } else if (totalK >= 2) {
        // Case 1: 2 K shifts -> Relief for K before 2nd K replacement
        if (sh[designatedKipId] === 'K' && kCount >= 1) {
          origAbsentKipMap['K'] = designatedKipId;
        }
      }
    }

    let bestScore = Infinity;
    let bestConfig: Record<number, string> = {};

    function solve(shiftIdx: number, usedPeople: Set<number>, current: Record<number, string>) {
      if (shiftIdx === 3) {
        const fullConfig: Record<number, string> = { ...current };
        availKips.forEach(k => {
          if (!usedPeople.has(k)) fullConfig[k] = 'O';
        });
        [1, 2, 3, 4, 5].forEach(k => { if (absentSet[k]) fullConfig[k] = 'O'; });

        let score = 0;
        const assignedShifts = new Set(Object.values(fullConfig));
        ['N', 'C', 'K'].forEach(s => {
          if (!assignedShifts.has(s)) score += 1000000; // Massive penalty for missing shift
        });

        for (let k = 1; k <= 5; k++) {
          const s = fullConfig[k];
          const naturalS = sh[k];
          if (s === 'O') continue;
          
          score += (coverCount[k] || 0) * 10;
          
          // Penalty for deviation from natural schedule
          if (s !== naturalS) {
            score += 1000;
          }

          const fb = isForbidden(k, s, prevShift, getNextActual, 'O', prevPrevShift);
          if (fb.bad) score += 10000;
          
          const origKip = origAbsentKipMap[s];
          const ruleKip = (origKip && RULES[origKip] && RULES[origKip][s]) ? RULES[origKip][s].k : null;

          // Special rule for single leave: Balance 'K' shifts
          if (activeLeaves.length === 1) {
            const absentKipId = activeLeaves[0].kip;
            const designatedKipId = RULES[absentKipId].K.k;
            const unsuitableKipId = [1, 2, 3, 4, 5].find(p =>
              p !== absentKipId &&
              p !== RULES[absentKipId].N.k &&
              p !== RULES[absentKipId].C.k &&
              p !== RULES[absentKipId].K.k
            );

            if (k === unsuitableKipId && (s === 'K' || s === 'N')) {
              const kCount = coverKCount[designatedKipId];
              const cCountCovered = coverCCount[designatedKipId];

              if (totalC >= 3) {
                if (s === 'N' && origKip === designatedKipId && cCountCovered >= 2) {
                  score -= 2000;
                }
                if (s === 'K' && origKip === designatedKipId && kCount >= 2) {
                  score -= 2000;
                }
              } else if (totalK >= 2) {
                if (s === 'K' && origKip === designatedKipId && kCount >= 1) {
                  score -= 2000;
                }
              }
            }
          }

          if (ruleKip === k) score -= 50;
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
      let assigned = false;
      for (const p of availKips) {
        if (!usedPeople.has(p)) {
          usedPeople.add(p);
          current[p] = s;
          solve(shiftIdx + 1, usedPeople, current);
          delete current[p];
          usedPeople.delete(p);
          assigned = true;
        }
      }
      if (!assigned || availKips.length < 3) {
         solve(shiftIdx + 1, usedPeople, current);
      }
    }

    solve(0, new Set(), {});

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
          
          if (absentSet[absentKip]) {
            const idx = kipToIdx[absentKip];
            if (idx !== undefined) {
              const isUnsuitable = activeLeaves.length === 1 && (assignedShift === 'K' || assignedShift === 'N') && (() => {
                const absentKipId = activeLeaves[0].kip;
                const designatedKipId = RULES[absentKipId].K.k;
                const unsuitableKipId = [1, 2, 3, 4, 5].find(p =>
                  p !== absentKipId &&
                  p !== RULES[absentKipId].N.k &&
                  p !== RULES[absentKipId].C.k &&
                  p !== RULES[absentKipId].K.k
                );
                
                if (k !== unsuitableKipId) return false;
                
                const kCount = coverKCount[designatedKipId];
                const cCountCovered = coverCCount[designatedKipId];
                const origKip = origAbsentKipMap[assignedShift];

                if (totalC >= 3) {
                  if (assignedShift === 'N' && origKip === designatedKipId && cCountCovered >= 2) return true;
                  if (assignedShift === 'K' && origKip === designatedKipId && kCount >= 2) return true;
                } else if (totalK >= 2) {
                  return assignedShift === 'K' && origKip === designatedKipId && kCount >= 1;
                }
                return false;
              })();

              results[idx].ketQua.push({
                ngay, ca: assignedShift, kipThay: k,
                nguoiThay: timThay(k, chucDanh, staffData),
                isConflict: isConf,
                conflictNote: isUnsuitable ? `Thay ca ${assignedShift} (điều chỉnh cân bằng tải)` : noteConf,
                isOverlapDay: activeLeaves.length >= 2
              });
              coverCount[k]++;
              if (assignedShift === 'K') coverKCount[k]++;
              if (assignedShift === 'C') coverCCount[k]++;
            }
          } else if (isReplacement) {
            const isCK = prevShift[k] === 'C' && assignedShift === 'O' && sh[k] === 'K';
            const isUnsuitable = activeLeaves.length === 1 && (assignedShift === 'K' || assignedShift === 'N') && (() => {
              const absentKipId = activeLeaves[0].kip;
              const designatedKipId = RULES[absentKipId].K.k;
              const unsuitableKipId = [1, 2, 3, 4, 5].find(p =>
                p !== absentKipId &&
                p !== RULES[absentKipId].N.k &&
                p !== RULES[absentKipId].C.k &&
                p !== RULES[absentKipId].K.k
              );
              
              if (k !== unsuitableKipId) return false;

              const kCount = coverKCount[designatedKipId];
              const cCountCovered = coverCCount[designatedKipId];
              const origKip = origAbsentKipMap[assignedShift];

              if (totalC >= 3) {
                if (assignedShift === 'N' && origKip === designatedKipId && cCountCovered >= 2) return true;
                if (assignedShift === 'K' && origKip === designatedKipId && kCount >= 2) return true;
              } else if (totalK >= 2) {
                return assignedShift === 'K' && origKip === designatedKipId && kCount >= 1;
              }
              return false;
            })();

            extraRows.push({
              ngay, ca: assignedShift, kipThay: k,
              nguoiThay: timThay(k, chucDanh, staffData),
              absentKip: absentKip, absentTen: timThay(absentKip, chucDanh, staffData),
              isConflict: isConf, 
              conflictNote: isUnsuitable ? `Thay ca ${assignedShift} (điều chỉnh cân bằng tải)` : (isConf ? noteConf : `${timThay(k, chucDanh, staffData)} trực thay cho ${timThay(absentKip, chucDanh, staffData)} (do điều chỉnh hệ thống)`),
              isCKChain: isCK, isSwap: true, isOverlapDay: activeLeaves.length >= 2
            });
            coverCount[k]++;
            if (assignedShift === 'K') coverKCount[k]++;
            if (assignedShift === 'C') coverCCount[k]++;
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
