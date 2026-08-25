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
  kiptructhay: number;
  nguoitructhay: string;
  isConflict: boolean;
  conflictNote?: string;
  isOverlapDay?: boolean;
  isCKSwap?: boolean;
  swapAbsentTen?: string;
  relievedTen?: string;
  relievedKip?: number;
}

export interface CustomShiftConfig {
  baseDate?: Date | string;
  shifts?: string[][];
  rules?: Record<number, Record<string, { k: number }>>;
}

export function buildMultiLeaveResults(
  danhSachLichNghi: Leave[],
  chucdanh: string,
  dulieunhanvien: string[][],
  customShiftConfig?: CustomShiftConfig
) {
  const activeBaseDate = customShiftConfig?.baseDate;
  const activeShifts = (customShiftConfig?.shifts && customShiftConfig.shifts.length > 0) ? customShiftConfig.shifts : SHIFTS;
  const activeRules = customShiftConfig?.rules || RULES;

  const fnXacDinhCa = (ngay: any, kip: number) => xacDinhCa(ngay, kip, activeBaseDate, activeShifts);

  const danhsachketqua = danhSachLichNghi.map(dongnghi => ({
    ten: dongnghi.ten,
    kip: dongnghi.kip,
    start: dongnghi.start,
    end: dongnghi.end,
    chucDanh: dongnghi.chucDanh || chucdanh,
    ketQua: [] as ResultItem[]
  }));

  const bangtraKIpnghi: Record<number, number> = {};
  const tongsocatrucbinghi: Record<number, number> = {};
  danhSachLichNghi.forEach((dongnghict, indexnghi) => { 
    bangtraKIpnghi[dongnghict.kip] = indexnghi; 
    let dem = 0;
    let ngaychao = new Date(dongnghict.start);
    while (ngaychao <= dongnghict.end) {
      const cahientai = fnXacDinhCa(ngaychao, dongnghict.kip);
      if (cahientai !== 'O') dem++;
      ngaychao.setDate(ngaychao.getDate() + 1);
    }
    tongsocatrucbinghi[dongnghict.kip] = dem;
  });

  const solantructhay: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  const solantructhaytichluy: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  const theodoidoica: Record<number, { caC: number, caK: number, caN: number, tongsocatructhaybimat: number, socadoicathuchien: number, socadoicanhantheokip: Record<number, number>, lastCycleShift?: 'N' | 'C' | 'K', firstSwapType?: 'N' | 'C' | 'K' }> = {};
  const bangphancongngay: Record<string, Record<number, string | undefined>> = {};
  const chancakhuyahomsau: Record<string, number[]> = {};
  const thongtinchancakhuyahomsau: Record<string, number> = {};
  const dongphatsinh: any[] = [];
  let coxungdot = false;

  const tatcangaynghi: Record<string, Date> = {};
  danhSachLichNghi.forEach(dongnghi => {
    let ngaychao = new Date(dongnghi.start);
    while (ngaychao <= dongnghi.end) {
      tatcangaynghi[fmtIn(ngaychao)] = new Date(ngaychao);
      ngaychao.setDate(ngaychao.getDate() + 1);
    }
  });

  const daxulyngay: Record<string, boolean> = {};

  // Pre-calculate total covers for each kip during the entire leave period
  // to determine eligibility for relief swaps (>= 3 covers = 1 relief)
  const thongketructhaykip: Record<number, { N: number, C: number, K: number, total: number }> = {
    1: { N: 0, C: 0, K: 0, total: 0 },
    2: { N: 0, C: 0, K: 0, total: 0 },
    3: { N: 0, C: 0, K: 0, total: 0 },
    4: { N: 0, C: 0, K: 0, total: 0 },
    5: { N: 0, C: 0, K: 0, total: 0 }
  };
  Object.keys(tatcangaynghi).forEach(khoangay => {
    const ngayhientruc = tatcangaynghi[khoangay];
    const danhsachnghingayhomnay = danhSachLichNghi.filter(dongnghi => ngayhientruc >= dongnghi.start && ngayhientruc <= dongnghi.end);
    danhsachnghingayhomnay.forEach(dongnghi => {
      const kipnghi = dongnghi.kip;
      const cahientai = fnXacDinhCa(ngayhientruc, kipnghi);
      if (cahientai !== 'O' && activeRules[kipnghi] && activeRules[kipnghi][cahientai]) {
        const kiptructhay = activeRules[kipnghi][cahientai].k;
        thongketructhaykip[kiptructhay].total++;
        if (cahientai === 'N') thongketructhaykip[kiptructhay].N++;
        if (cahientai === 'C') thongketructhaykip[kiptructhay].C++;
        if (cahientai === 'K') thongketructhaykip[kiptructhay].K++;
      }
    });
  });

  function layngaychuaxulytiep() {
    const danhsachkhoa = Object.keys(tatcangaynghi).sort();
    for (let indexnghi = 0; indexnghi < danhsachkhoa.length; indexnghi++) {
      if (!daxulyngay[danhsachkhoa[indexnghi]]) return danhsachkhoa[indexnghi];
    }
    return null;
  }

  let khoangayhientai: string | null;
  while ((khoangayhientai = layngaychuaxulytiep()) !== null) {
    daxulyngay[khoangayhientai] = true;
    const ngayhientruc = tatcangaynghi[khoangayhientai];
    const ngaymai = new Date(ngayhientruc.getTime() + 86400000);
    const khoangayhomtruoc = fmtIn(new Date(ngayhientruc.getTime() - 86400000));

    const lichnghihientaingay = danhSachLichNghi.filter(dongnghi => ngayhientruc >= dongnghi.start && ngayhientruc <= dongnghi.end);
    const danhsachtatcakipnghi: Record<number, boolean> = {};
    lichnghihientaingay.forEach(dongnghi => { danhsachtatcakipnghi[dongnghi.kip] = true; });
    const danhsachkipkhadung = [1, 2, 3, 4, 5].filter(k => !danhsachtatcakipnghi[k]);

    const calamviectunhien: Record<number, string> = {};
    const calamviectunhienngaymai: Record<number, string> = {};
    for (let k = 1; k <= 5; k++) {
      calamviectunhien[k] = fnXacDinhCa(ngayhientruc, k);
      calamviectunhienngaymai[k] = fnXacDinhCa(ngaymai, k);
    }

    const ngayhomtruoc = new Date(ngayhientruc.getTime() - 86400000);
    const ngayhomtruochia = new Date(ngayhientruc.getTime() - 172800000);
    const caphanconghomtruoc: Record<number, string> = {};
    const caphanconghomtruochia: Record<number, string> = {};
    for (let k = 1; k <= 5; k++) {
      caphanconghomtruoc[k] = (bangphancongngay[khoangayhomtruoc] && bangphancongngay[khoangayhomtruoc][k] != null)
        ? (bangphancongngay[khoangayhomtruoc][k] as string)
        : fnXacDinhCa(ngayhomtruoc, k);
      
      const prevPrevKey = fmtIn(ngayhomtruochia);
      caphanconghomtruochia[k] = (bangphancongngay[prevPrevKey] && bangphancongngay[prevPrevKey][k] != null)
        ? (bangphancongngay[prevPrevKey][k] as string)
        : fnXacDinhCa(ngayhomtruochia, k);
    }
    const khoangaymai = fmtIn(ngaymai);
    
    if (!bangphancongngay[khoangayhientai]) bangphancongngay[khoangayhientai] = {};

    function laycathuctehomsau(kip: number, offset: number = 1) {
      const ngaymuontoi = new Date(ngayhientruc.getTime() + 86400000 * offset);
      const khoangaymuontoi = fmtIn(ngaymuontoi);
      
      const langhingaymuontoi = danhSachLichNghi.some(dongnghi => ngaymuontoi >= dongnghi.start && ngaymuontoi <= dongnghi.end && dongnghi.kip === kip);
      if (langhingaymuontoi) return 'O';

      if (bangphancongngay[khoangaymuontoi] && bangphancongngay[khoangaymuontoi][kip])
        return bangphancongngay[khoangaymuontoi][kip];
      const chancahomsau = chancakhuyahomsau[khoangaymuontoi] || [];
      if (chancahomsau.indexOf(kip) !== -1) return 'O';
      return fnXacDinhCa(ngaymuontoi, kip);
    }

    // Map to track who is originally in which shift today
    const anhxanhomcagoc: Record<string, number> = {};
    ['N', 'C', 'K', 'O'].forEach(s => {
      for (let k = 1; k <= 5; k++) {
        if (calamviectunhien[k] === s) {
          anhxanhomcagoc[s] = k;
        }
      }
    });

    // Track relief progress and determine if a swap is needed today
    const phancongdoica: Record<number, string> = {};
    const danhsachcaclichdoica: Array<{ absentKip: number, shift: string, helperKip: number, relievedKip: number, relievedTen: string }> = [];

    lichnghihientaingay.forEach(dongnghi => {
      const kipnghi = dongnghi.kip;
      if (!theodoidoica[kipnghi]) {
        theodoidoica[kipnghi] = { 
          caC: 0, caK: 0, caN: 0, 
          tongsocatructhaybimat: 0, 
          socadoicathuchien: 0, 
          socadoicanhantheokip: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 },
          firstSwapType: undefined
        };
      }
      
      const s = fnXacDinhCa(ngayhientruc, kipnghi);
      const helperKip = [1, 2, 3, 4, 5].find(k => k !== kipnghi && k !== activeRules[kipnghi]?.N?.k && k !== activeRules[kipnghi]?.C?.k && k !== activeRules[kipnghi]?.K?.k)!;

      if (s !== 'O') {
        // Increment tracker
        theodoidoica[kipnghi].tongsocatructhaybimat++;
        if (s === 'C') theodoidoica[kipnghi].caC++;
        if (s === 'K') {
          theodoidoica[kipnghi].caK++;
        }
        if (s === 'N') theodoidoica[kipnghi].caN++;
      }

      // Relief logic: Only for 1-person leave
      // Requirement: Only start relief after at least 3 working shifts (N, C, K) have been missed
      // And the helper kip must be naturally Off today AND it must be an "O tròn" (Off after K)
      const la_o_tron = caphanconghomtruoc[helperKip] === 'K' && calamviectunhien[helperKip] === 'O';
      
      // 1-person leave relief logic (Đổi ca)
      if (danhSachLichNghi.length === 1 && la_o_tron && !danhsachtatcakipnghi[helperKip]) {
        const cacayeucauhomnay = new Set<string>();
        for (let k = 1; k <= 5; k++) {
          const st = fnXacDinhCa(ngayhientruc, k);
          if (st !== 'O') cacayeucauhomnay.add(st);
        }

        let CaBiMatCanDuocBu: 'N' | 'C' | 'K' | null = null;
        let cadoica: 'N' | 'C' | 'K' | null = null;
        let kipduocdoica: number | null = null;
        let CaCanDuocBu: 'N' | 'C' | 'K' | null = null;
        
        // Chu kỳ đổi ca: Lần 1 (K), Lần 2 (N/C), Lần 3 (N/C), Lần 4 (K)...
        const chukyidx = theodoidoica[kipnghi].socadoicathuchien % 3;

        // Calculate cover stats for THIS specific leave to ensure "trong suốt kỳ nghỉ" condition
        const thongKeTrucThay: Record<number, { N: number, C: number, K: number }> = {
          1: { N: 0, C: 0, K: 0 }, 2: { N: 0, C: 0, K: 0 }, 3: { N: 0, C: 0, K: 0 }, 4: { N: 0, C: 0, K: 0 }, 5: { N: 0, C: 0, K: 0 }
        };
        Object.keys(tatcangaynghi).forEach(khoangay => {
          const ngayhientruc_d = tatcangaynghi[khoangay];
          // Chỉ tính trong phạm vi nghỉ của người này để tránh cộng dồn sai từ các kỳ nghỉ khác
          if (ngayhientruc_d < dongnghi.start || ngayhientruc_d > dongnghi.end) return;
          const s_d = fnXacDinhCa(ngayhientruc_d, kipnghi);
          if (s_d !== 'O' && activeRules[kipnghi] && activeRules[kipnghi][s_d]) {
            const kiptructhay = activeRules[kipnghi][s_d].k;
            if (s_d === 'N') thongKeTrucThay[kiptructhay].N++;
            if (s_d === 'C') thongKeTrucThay[kiptructhay].C++;
            if (s_d === 'K') thongKeTrucThay[kiptructhay].K++;
          }
        });

        const kipcaN = activeRules[kipnghi]?.N?.k || 1;
        const kipcaC = activeRules[kipnghi]?.C?.k || 2;
        const kipcaK = activeRules[kipnghi]?.K?.k || 3;
        
        const demcaN = thongKeTrucThay[kipcaN].N;
        const demcaC = thongKeTrucThay[kipcaC].C;
        const demcaK = thongKeTrucThay[kipcaK].K;

       
           const uuTien = { K: 3, N: 2, C: 1 };

const danhSachCa = [
  { ca: 'K' as const, dem: demcaK },
  { ca: 'N' as const, dem: demcaN },
  { ca: 'C' as const, dem: demcaC }
].sort((a, b) => {
  // Ưu tiên số lượng lớn hơn
  if (b.dem !== a.dem) {
    return b.dem - a.dem;
  }

  // Nếu bằng nhau: K > N > C
  return uuTien[b.ca] - uuTien[a.ca];
});

const caThuNhat = danhSachCa[0].ca;
const caThuHai = danhSachCa[1].ca;
const caThuBa = danhSachCa[2].ca;

// Chọn ca theo chu kỳ
if (chukyidx === 0) {
  CaCanDuocBu = caThuNhat;
  CaBiMatCanDuocBu = caThuNhat;
}
else if (chukyidx === 1) {
  CaBiMatCanDuocBu = caThuHai;
}
else {
  CaBiMatCanDuocBu = caThuBa;
}
        if (CaBiMatCanDuocBu) {
          kipduocdoica = activeRules[kipnghi][CaBiMatCanDuocBu].k;
          // Quan trọng: cadoica phải là ca mà kipduocdoica đang trực tự nhiên hôm nay
          if (kipduocdoica && calamviectunhien[kipduocdoica] !== 'O') {
            cadoica = calamviectunhien[kipduocdoica] as 'N' | 'C' | 'K';
          }
        }

        // Điều kiện: Kíp được thay (kipduocdoica) phải đang có ca trực tự nhiên (cadoica đã được gán ở trên)
        if (kipduocdoica && cadoica) {
          if (!danhsachtatcakipnghi[kipduocdoica] && !phancongdoica[helperKip] && !phancongdoica[kipduocdoica]) {
            // Điều kiện tích lũy theo yêu cầu người dùng:
            // Lần 1: Covered >= 2, Missed >= 2
            // Lần 2: Covered >= 3, Missed >= 5
            // Lần 3: Covered >= 4, Missed >= 7
            // Lần 4: Covered >= 5, Missed >= 9
            // Lần 5: Covered >= 6, Missed >= 11
            // Lần 6: Covered >= 7, Missed >= 13
            let cothedoica = false;
            const tongsocadoica = theodoidoica[kipnghi].socadoicathuchien;
            const tongsocanghi = theodoidoica[kipnghi].tongsocatructhaybimat;
            const loaicahandle = (chukyidx === 0) ? CaCanDuocBu : CaBiMatCanDuocBu;
            const demcahientai = loaicahandle ? thongKeTrucThay[kipduocdoica][loaicahandle] : 0;

            if (tongsocadoica === 0) { // ===== LẦN 1 =====

  const tongcatrucphaithay =
    tongsocatrucbinghi[kipnghi] || 0;

  // ==================================================
  // NGHỈ DÀI (>= 9 ca)
  // ==================================================
  if (tongcatrucphaithay >= 9) {

    // Chỉ mở đổi khi đã nghỉ khá lâu
    if (
      tongsocanghi >= 3 &&
      demcahientai >= 2
    ) {
      cothedoica = true;
    }

    // Ưu tiên đặc biệt nhưng vẫn phải nghỉ đủ lâu
    if (
      tongsocanghi >= 5 &&
      demcaK >= 3 &&
      demcaN >= 3 &&
      demcaC >= 2
    ) {
      cothedoica = true;
    }

  }

  // ==================================================
  // NGHỈ TRUNG BÌNH (5 -> 7 ca)
  // ==================================================
  else if (tongcatrucphaithay >= 5) {

    if (
      tongsocanghi >= 2 &&
      demcahientai >= 2
    ) {
      cothedoica = true;
    }

    if (
      tongsocanghi >= 2 &&
      demcaK >= 3 &&
      demcaN >= 3 &&
      demcaC >= 2
    ) {
      cothedoica = true;
    }

  }

  // ==================================================
  // NGHỈ NGẮN (< 5 ca)
  // ==================================================
  else {

    if (
      tongsocanghi >= 1 &&
      demcahientai >= 2
    ) {
      cothedoica = true;
    }

    if (
      demcaK >= 3 &&
      demcaN >= 3 &&
      demcaC >= 2
    ) {
      cothedoica = true;
    }
  }

}

// ======================================================
// LẦN 2
// ======================================================
else if (tongsocadoica === 1) {

  const tongcatrucphaithay =
    tongsocatrucbinghi[kipnghi] || 0;

  if (tongcatrucphaithay >= 8) {

    if (
      demcahientai >= 3 &&
      tongsocanghi >= 5
    ) {
      cothedoica = true;
    }

  }
}

// ======================================================
// CÁC LẦN TIẾP THEO
// ======================================================
else if (tongsocadoica === 2) { // Lần 3

  const tongcatrucphaithay =
    tongsocatrucbinghi[kipnghi] || 0;

  if (
    tongcatrucphaithay >= 12 &&
    demcahientai >= 4 &&
    tongsocanghi >= 7
  ) {
    cothedoica = true;
  }

}
else if (tongsocadoica === 3) { // Lần 4
const tongcatrucphaithay =
    tongsocatrucbinghi[kipnghi] || 0;

  if (
    tongcatrucphaithay >= 16 &&
    demcahientai >= 5 &&
    tongsocanghi >= 9
  ) {
    cothedoica = true;
  }}


else if (tongsocadoica === 4) { // Lần 5
const tongcatrucphaithay =
    tongsocatrucbinghi[kipnghi] || 0;
  if (
    tongcatrucphaithay >= 20 &&
    demcahientai >= 6 &&
    tongsocanghi >= 13
  ) {
    cothedoica = true;
  }

}
else if (tongsocadoica === 5) { // Lần 6
const tongcatrucphaithay =
    tongsocatrucbinghi[kipnghi] || 0;
  if (
     tongcatrucphaithay >= 24 &&
    demcahientai >= 7 &&
    tongsocanghi >= 15
  ) {
    cothedoica = true;
  }

}
else {

  // Từ lần 7 trở đi
  if (
    demcahientai >= (tongsocadoica + 2) &&
    tongsocanghi >= (tongsocadoica * 2 + 3)
  ) {
    cothedoica = true;
  }

}


              if (cothedoica) {
                phancongdoica[helperKip] = cadoica;
                phancongdoica[kipduocdoica] = 'O';
                danhsachcaclichdoica.push({ 
                  absentKip: kipnghi, 
                  shift: cadoica, 
                  helperKip, 
                  relievedKip: kipduocdoica,
                  relievedTen: timThay(kipduocdoica, chucdanh, dulieunhanvien)
                });
                
                // Lưu lại loại ca đã đổi ở lần 1
                if (chukyidx === 0) {
                  theodoidoica[kipnghi].firstSwapType = CaBiMatCanDuocBu as 'N' | 'C' | 'K';
                }
                
                // Lưu lại ca đã bù đắp ở lần 2 để lần 3 chọn ca còn lại
                if (chukyidx === 1) {
                  theodoidoica[kipnghi].lastCycleShift = CaBiMatCanDuocBu as 'N' | 'C' | 'K';
                }
                
                theodoidoica[kipnghi].socadoicanhantheokip[kipduocdoica]++;
                theodoidoica[kipnghi].socadoicathuchien++;
              }
          }
        }
      }
    });

    let diemtotnhat = Infinity;
    let cauhinhbest: Record<number, string> = {};

    function giaiphancong(caxidx: number, nguoidadung: Set<number>, cauhinhhientai: Record<number, string>, kipkhadunghientai: number[]) {
      if (caxidx === 3) {
        const cauhinhdaydu: Record<number, string> = { ...cauhinhhientai };
        // Fill in 'O' for anyone not assigned and not already forced
        [1, 2, 3, 4, 5].forEach(k => {
          if (!cauhinhdaydu[k]) cauhinhdaydu[k] = 'O';
        });

        const cacayeucautructep = new Set<string>();
        for (let k = 1; k <= 5; k++) {
          const s = fnXacDinhCa(ngayhientruc, k);
          if (s !== 'O') cacayeucautructep.add(s);
        }

        let diemso = 0;
        const cadaphancong = new Set(Object.values(cauhinhdaydu));
        cacayeucautructep.forEach(s => {
          if (!cadaphancong.has(s)) diemso += 1000000; 
        });

        for (let k = 1; k <= 5; k++) {
          const s = cauhinhdaydu[k];
          const catunhien = calamviectunhien[k];
          
          if (s === 'O') {
            if (catunhien !== 'O' && !danhsachtatcakipnghi[k]) {
              // Naturally working but assigned Off - massive penalty unless forced by relief
              diemso += 1000000; 
            }
            continue;
          }
          
          const kipgoc = anhxanhomcagoc[s];
          const lakipnghigoc = danhsachtatcakipnghi[kipgoc];

          if (k !== kipgoc) {
            if (lakipnghigoc) {
              // Covering a leave - Good, but prefer rule-based person
              diemso += 1000;
              let kiptheoquytac = (kipgoc && activeRules[kipgoc] && activeRules[kipgoc][s]) ? activeRules[kipgoc][s].k : null;
              
              if (kiptheoquytac === k) diemso -= 500;
            } else {
              // Stealing a shift from someone who is NOT on leave - Massive penalty
              diemso += 2000000;
            }
          }

          const rangbuoccaphan = isForbidden(k, s, caphanconghomtruoc, laycathuctehomsau, 'O', caphanconghomtruochia);
          if (rangbuoccaphan.bad) diemso += 5000000; 
          
          if (s === 'K' && caphanconghomtruoc[k] === 'N') diemso -= 30;
          if (s === 'K' && caphanconghomtruoc[k] === 'K') diemso += 100;
        }

        if (diemso < diemtotnhat) {
          diemtotnhat = diemso;
          cauhinhbest = cauhinhdaydu;
        }
        return;
      }

      const s = ['N', 'C', 'K'][caxidx];
      
      // If this shift is already forced, skip to next shift
      const forcedKip = Object.keys(phancongdoica).find(k => phancongdoica[Number(k)] === s);
      if (forcedKip) {
        giaiphancong(caxidx + 1, nguoidadung, cauhinhhientai, kipkhadunghientai);
        return;
      }

      let assigned = false;
      for (const p of kipkhadunghientai) {
        if (!nguoidadung.has(p)) {
          nguoidadung.add(p);
          cauhinhhientai[p] = s;
          giaiphancong(caxidx + 1, nguoidadung, cauhinhhientai, kipkhadunghientai);
          delete cauhinhhientai[p];
          nguoidadung.delete(p);
          assigned = true;
        }
      }
      if (!assigned || kipkhadunghientai.length < 3) {
         giaiphancong(caxidx + 1, nguoidadung, cauhinhhientai, kipkhadunghientai);
      }
    }

    // Initialize solver with forced assignments
    const dadungbandau = new Set<number>();
    const hientaibandau: Record<number, string> = {};
    const kipkhadunghieuluc = danhsachkipkhadung.filter(k => {
      if (phancongdoica[k] === 'O') {
        hientaibandau[k] = 'O';
        return false;
      }
      if (phancongdoica[k]) {
        dadungbandau.add(k);
        hientaibandau[k] = phancongdoica[k];
        return false;
      }
      return true;
    });

    giaiphancong(0, dadungbandau, hientaibandau, kipkhadunghieuluc);

    if (diemtotnhat === Infinity) {
      cauhinhbest = {};
      for (let k = 1; k <= 5; k++) cauhinhbest[k] = danhsachtatcakipnghi[k] ? 'O' : calamviectunhien[k];
    }

    // Check for missing shifts and report them
    const cacacuoicungduocphan = new Set(Object.values(cauhinhbest));
    const cacayeucauhomnay = new Set<string>();
    for (let k = 1; k <= 5; k++) {
      const s = fnXacDinhCa(ngayhientruc, k);
      if (s !== 'O') cacayeucauhomnay.add(s);
    }

    cacayeucauhomnay.forEach(s => {
      if (!cacacuoicungduocphan.has(s)) {
        const absentKip = anhxanhomcagoc[s];
        const tenAbsent = timThay(absentKip, chucdanh, dulieunhanvien);
        dongphatsinh.push({
          ngay: ngayhientruc, ca: s, kiptructhay: 0,
          nguoitructhay: '⚠️ CHƯA CÓ NGƯỜI TRỰC',
          absentKip: absentKip, absentTen: tenAbsent, chucDanh: chucdanh,
          isConflict: true, conflictNote: `Không tìm được người thay cho Ca ${s} của ${tenAbsent}`,
          isCKChain: false, isSwap: false, isOverlapDay: lichnghihientaingay.length >= 2
        });
        coxungdot = true;
      }
    });

    for (let k = 1; k <= 5; k++) {
      const caphanchinh = cauhinhbest[k];
      bangphancongngay[khoangayhientai!][k] = caphanchinh;
      
      if (caphanchinh !== 'O') {
        const catunhien = calamviectunhien[k];
        const lathaythe = caphanchinh !== catunhien;
        const absentKip = anhxanhomcagoc[caphanchinh];

        if (lathaythe || danhsachtatcakipnghi[absentKip]) {
          const rangbuoccaphan = isForbidden(k, caphanchinh, caphanconghomtruoc, laycathuctehomsau, 'O', caphanconghomtruochia);
          const bixungdot = rangbuoccaphan.bad;
          const ghichuxungdot = bixungdot ? (rangbuoccaphan.note || '⚠ Vi phạm ràng buộc ca') : '';
          
          const lavuongCK = caphanconghomtruoc[k] === 'C' && caphanchinh === 'O' && calamviectunhien[k] === 'K';
          const ghichuCK = lavuongCK ? `⥵ C→K: ${timThay(k, chucdanh, dulieunhanvien)} vướng ca C hôm trước, không thể trực ca K hôm nay` : '';

          const thongtindoica = danhsachcaclichdoica.find(fr => fr.helperKip === k && fr.shift === caphanchinh);
          const kipnghithucte = thongtindoica ? thongtindoica.absentKip : absentKip;
          const tenkipnghihientai = timThay(kipnghithucte, chucdanh, dulieunhanvien);
          const tennguoiduocbu = thongtindoica ? thongtindoica.relievedTen : null;
          const kipduocbu = thongtindoica ? thongtindoica.relievedKip : null;

          if (danhsachtatcakipnghi[absentKip]) {
            const idx = bangtraKIpnghi[absentKip];
            if (idx !== undefined) {
              danhsachketqua[idx].ketQua.push({
                ngay: ngayhientruc, ca: caphanchinh, kiptructhay: k,
                nguoitructhay: timThay(k, chucdanh, dulieunhanvien),
                relievedTen: tennguoiduocbu || undefined,
                relievedKip: kipduocbu || undefined,
                isConflict: bixungdot,
                conflictNote: thongtindoica ? `${timThay(k, chucdanh, dulieunhanvien)} trực thay ${tennguoiduocbu} ca ${caphanchinh}` : (ghichuCK || ghichuxungdot),
                isOverlapDay: lichnghihientaingay.length >= 2,
                isCKSwap: lavuongCK,
                swapAbsentTen: thongtindoica ? tenkipnghihientai : undefined
              });
              solantructhay[k]++;
              solantructhaytichluy[k]++;
            }
          } else if (lathaythe) {
            const ladoicathucong = !danhsachtatcakipnghi[kipnghithucte];
            dongphatsinh.push({
              ngay: ngayhientruc, ca: caphanchinh, kiptructhay: k,
              nguoitructhay: timThay(k, chucdanh, dulieunhanvien),
              absentKip: kipnghithucte, absentTen: tenkipnghihientai, chucDanh: chucdanh,
              relievedTen: tennguoiduocbu,
              relievedKip: kipduocbu,
              isConflict: bixungdot, 
              conflictNote: ghichuCK || (bixungdot ? ghichuxungdot : (ladoicathucong || thongtindoica ? `${timThay(k, chucdanh, dulieunhanvien)} trực thay ${tennguoiduocbu || tenkipnghihientai} ca ${caphanchinh}` : `△ Điều chỉnh hệ thống: ${timThay(k, chucdanh, dulieunhanvien)} thay cho ${tenkipnghihientai}`)),
              isCKChain: lavuongCK, isSwap: ladoicathucong || !!thongtindoica, isOverlapDay: lichnghihientaingay.length >= 2
            });
            solantructhay[k]++;
          }
        }

        if (caphanchinh === 'C') {
          if (!chancakhuyahomsau[khoangaymai]) chancakhuyahomsau[khoangaymai] = [];
          if (chancakhuyahomsau[khoangaymai].indexOf(k) === -1) chancakhuyahomsau[khoangaymai].push(k);
          
          // Only add tomorrow to tatcangaynghi if it's within the original leave range + 1 day
          // AND the person is actually blocked from their natural shift tomorrow.
          const maxLeaveEnd = danhSachLichNghi.length > 0 ? Math.max(...danhSachLichNghi.map(dongnghi => dongnghi.end.getTime())) : 0;
          const ngaynghimaxdate = new Date(maxLeaveEnd + 86400000);
          const canhientaihomsau = fnXacDinhCa(ngaymai, k);
          
          if (!tatcangaynghi[khoangaymai] && ngaymai <= ngaynghimaxdate && canhientaihomsau === 'K') {
            tatcangaynghi[khoangaymai] = new Date(ngaymai);
          }
        }
      }
    }
    
    if (diemtotnhat >= 10000 && diemtotnhat < 1000000) coxungdot = true;
  }

  return { results: danhsachketqua, extraRows: dongphatsinh, hasConflict: coxungdot, coverCount: solantructhay };
}
