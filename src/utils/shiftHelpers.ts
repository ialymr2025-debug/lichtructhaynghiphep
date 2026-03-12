import { BASE_DATE, SHIFTS, RULES } from '../constants';

export function fmtVN(d: Date) {
  return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
}

export function fmtIn(d: Date) {
  const y = d.getFullYear();
  const m = ('0' + (d.getMonth() + 1)).slice(-2);
  const day = ('0' + d.getDate()).slice(-2);
  return `${y}-${m}-${day}`;
}

export function dayN(d: Date) {
  return ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'][d.getDay()];
}

export function abbrev(n: string) {
  const p = n.trim().split(/\s+/);
  if (p.length <= 1) return n;
  return p.slice(0, -1).map(x => x[0] + '.').join('') + p[p.length - 1];
}

export function xacDinhCa(ngay: Date, kip: number) {
  const diff = Math.floor((ngay.getTime() - BASE_DATE.getTime()) / 86400000);
  const cycleLen = SHIFTS[0].length;
  return SHIFTS[kip - 1][((diff % cycleLen) + cycleLen) % cycleLen];
}

export function timNghi(cd: string, kip: number, staffData: string[][]) {
  for (let i = 0; i < staffData.length; i++) {
    const r = staffData[i];
    if (r[0].trim() === cd && r[kip] && r[kip].trim()) return r[kip].trim();
  }
  return null;
}

export function timThay(kipThay: number, cd: string, staffData: string[][]) {
  if (kipThay < 1 || kipThay > 5) return '';
  for (let i = 0; i < staffData.length; i++) {
    const r = staffData[i];
    if (r[0].trim() === cd && r[kipThay] && r[kipThay].trim()) return r[kipThay].trim();
  }
  for (let i = 0; i < staffData.length; i++) {
    if (staffData[i][kipThay] && staffData[i][kipThay].trim()) return staffData[i][kipThay].trim();
  }
  return 'N/A';
}

export function isForbidden(kip: number, shift: string, prevShift: Record<number, string>, getNext: any, todayActual?: string, prevPrevShift?: Record<number, string>) {
  const pre = prevShift[kip];
  const post = typeof getNext === 'function' ? getNext(kip, 1) : (getNext ? getNext[kip] : null);
  const postPost = typeof getNext === 'function' ? getNext(kip, 2) : null;
  const cur = todayActual || null;
  const prePre = prevPrevShift ? prevPrevShift[kip] : null;
  
  // 0. Trùng ca: Nếu kíp này đã có ca trực (tự nhiên hoặc đã phân thay) thì không thể đi thay ca khác
  if (shift !== 'O' && cur !== 'O' && cur !== null && shift !== cur) {
    return { bad: true, note: `Vi phạm: Kíp ${kip} đã có ca ${cur}` };
  }

  // 1. Ca C hôm trước (ngày n) -> Ca K hôm nay (ngày n+1): CẤM TUYỆT ĐỐI
  if (pre === 'C' && shift === 'K') return { bad: true, note: 'Vi phạm: Ca C hôm qua → Ca K hôm nay (0h nghỉ)' };
  
  // 2. Ca K hôm nay (ngày n) -> Ca N hôm nay (ngày n): CẤM TUYỆT ĐỐI
  if (shift === 'K' && cur === 'N') return { bad: true, note: 'Vi phạm: Ca K và Ca N cùng ngày (0h nghỉ)' };
  if (shift === 'N' && cur === 'K') return { bad: true, note: 'Vi phạm: Ca N và Ca K cùng ngày (0h nghỉ)' };

  // 3. Ca K hôm nay (ngày n) -> Ca C hôm nay (ngày n): CẤM (8h nghỉ - rất tight)
  if (shift === 'K' && cur === 'C') return { bad: true, note: 'Vi phạm: Ca K và Ca C cùng ngày (8h nghỉ)' };
  if (shift === 'C' && cur === 'K') return { bad: true, note: 'Vi phạm: Ca C và Ca K cùng ngày (8h nghỉ)' };

  // 4. Ca C hôm nay (ngày n) -> Ca K ngày mai (ngày n+1): CẤM
  // Nếu post là ca tự nhiên, ta có thể coi là "có thể điều chỉnh" nhưng vẫn rất xấu
  if (shift === 'C' && post === 'K') return { bad: true, note: 'Vi phạm: Ca C hôm nay → Ca K ngày mai' };
  
  // 5. Ca K liên tiếp 3 ngày: CẤM
  if (shift === 'K') {
    if (pre === 'K' && prePre === 'K') return { bad: true, note: 'Vi phạm: 3 ca K liên tiếp (K-K-K)' };
    if (pre === 'K' && post === 'K') return { bad: true, note: 'Vi phạm: 3 ca K liên tiếp (K-K-K ngày mai)' };
    if (post === 'K' && postPost === 'K') return { bad: true, note: 'Vi phạm: 3 ca K liên tiếp (K-K-K ngày kia)' };
  }
  
  return { bad: false, note: '' };
}

export function shiftPenalty(kip: number, shift: string, prevShift: Record<number, string>, getNext: any, coverCount: Record<number, number>, todayActual?: string, prevPrevShift?: Record<number, string>) {
  let p = (coverCount[kip] || 0) * 50;
  const fb = isForbidden(kip, shift, prevShift, getNext, todayActual, prevPrevShift);
  
  if (fb.bad) {
    // Nếu vi phạm là do ca tương lai (post/postPost), ta giảm nhẹ penalty để thuật toán có thể chọn nếu cực kỳ cần thiết
    // nhưng vẫn để ở mức rất cao (200.000) để ưu tiên các phương án khác.
    // Các vi phạm quá khứ hoặc cùng ngày thì phạt cực nặng (500.000).
    const isFutureViolation = fb.note.includes('ngày mai') || fb.note.includes('ngày kia');
    p += isFutureViolation ? 200000 : 500000;
  }
  
  if (shift === 'K') {
    const pre = prevShift[kip];
    if (pre === 'K') p += 2000; // Đã đi K hôm qua, ưu tiên người khác
  }
  
  return p;
}

export function pickBest(offPool: number[], shift: string, prevShift: Record<number, string>, getNext: any, coverCount: Record<number, number>, preferKip: number | null, sh?: Record<number, string>, prevPrevShift?: Record<number, string>) {
  const noCK = offPool.filter(k => {
    if (shift === 'K' && prevShift[k] === 'C') return false;
    return true;
  });
  const valid = noCK.filter(k => {
    return !isForbidden(k, shift, prevShift, getNext, sh ? sh[k] : undefined, prevPrevShift).bad;
  });
  const pool = valid.length > 0 ? valid : (noCK.length > 0 ? noCK : offPool);
  return pool.slice().sort((a, b) => {
    const bA = (a === preferKip) ? -15 : 0;
    const bB = (b === preferKip) ? -15 : 0;
    return (shiftPenalty(a, shift, prevShift, getNext, coverCount, sh ? sh[a] : undefined, prevPrevShift) + bA)
      - (shiftPenalty(b, shift, prevShift, getNext, coverCount, sh ? sh[b] : undefined, prevPrevShift) + bB);
  })[0] || null;
}

export function buildConflict(kip: number, shift: string, prevShift: Record<number, string>, getNext: any, ruleKip: number | null, prevPrevShift?: Record<number, string>) {
  const fb = isForbidden(kip, shift, prevShift, getNext, undefined, prevPrevShift);
  if (fb.bad) return { flag: true, note: '⚠ Không tránh được: ' + fb.note };
  if (ruleKip && kip !== ruleKip) return { flag: true, note: 'Điều chỉnh (kíp quy tắc cũng nghỉ / cân bằng)' };
  return { flag: false, note: '' };
}
