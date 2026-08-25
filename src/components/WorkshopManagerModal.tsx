import React, { useState, useEffect, useRef } from 'react';
import { API_BASE } from '../utils/api';
import { UserAccount, Workshop, WorkshopFeatures, WorkshopConfig } from '../types/auth';
import { SHIFTS, RULES } from '../constants';
import StaffDataEditor from './StaffDataEditor';
import * as XLSX from 'xlsx';
import { 
  X, Factory, Plus, Edit2, Trash2, Check, Shield, Users, Settings, 
  ToggleLeft, ToggleRight, Save, UserPlus, KeyRound, Sparkles, Building2, AlertCircle, Search, RefreshCw, LayoutGrid, Code, ArrowRight,
  FileSpreadsheet, Download, Upload
} from 'lucide-react';

interface WorkshopManagerModalProps {
  user: UserAccount;
  workshops: Workshop[];
  activeWorkshop: Workshop | null;
  onClose: () => void;
  onRefreshWorkshops: () => void;
}

export default function WorkshopManagerModal({
  user,
  workshops,
  activeWorkshop,
  onClose,
  onRefreshWorkshops
}: WorkshopManagerModalProps) {
  const isSuperAdmin = user.role === 'super_admin';
  const userWorkshopId = user.workshopId;

  // Filter workshops & accounts based on user role
  const visibleWorkshops = isSuperAdmin 
    ? workshops 
    : workshops.filter(w => w.id === userWorkshopId);

  const [activeTab, setActiveTab] = useState<'workshops' | 'accounts' | 'create_admin' | 'edit_workshop' | 'audit'>('workshops');
  // Audit trail. Loaded on demand so opening the modal stays fast.
  const [auditRows, setAuditRows] = useState<any[]>([]);
  const [auditLoading, setAuditLoading] = useState(false);

  const loadAudit = async () => {
    setAuditLoading(true);
    try {
      const res = await fetch(API_BASE + '/api/audit-log?limit=300');
      const data = await res.json();
      setAuditRows(Array.isArray(data) ? data : []);
    } catch (e) {
      console.error('Loi tai nhat ky', e);
    } finally {
      setAuditLoading(false);
    }
  };
  const [selectedWs, setSelectedWs] = useState<Workshop | null>(
    isSuperAdmin ? (activeWorkshop || workshops[0] || null) : (workshops.find(w => w.id === userWorkshopId) || null)
  );

  // Create Admin Account & Initial Setup State
  const [createAdminUsername, setCreateAdminUsername] = useState('');
  const [createAdminPassword, setCreateAdminPassword] = useState('123456');
  const [createAdminFullName, setCreateAdminFullName] = useState('');
  const [createAdminWsOption, setCreateAdminWsOption] = useState<'new' | 'existing'>('new');
  const [createAdminWsName, setCreateAdminWsName] = useState('');
  const [createAdminWsCode, setCreateAdminWsCode] = useState('');
  const [createAdminWsDesc, setCreateAdminWsDesc] = useState('');
  const [createAdminSelectedWsId, setCreateAdminSelectedWsId] = useState('');
  const [isCreatingAdminAndSetup, setIsCreatingAdminAndSetup] = useState(false);

  // Accounts state
  const [accounts, setAccounts] = useState<UserAccount[]>([]);
  const [loadingAccounts, setLoadingAccounts] = useState(false);

  // Workshop Form State
  const [editingWsId, setEditingWsId] = useState<string | null>(isSuperAdmin ? null : userWorkshopId);
  const [wsName, setWsName] = useState('');
  const [wsCode, setWsCode] = useState('');
  const [wsDesc, setWsDesc] = useState('');
  const [wsNguoiKy, setWsNguoiKy] = useState('Nguyễn Văn Nghị');
  const [wsChucVu, setWsChucVu] = useState('Quản đốc Phân xưởng');
  const [wsSoVanBan, setWsSoVanBan] = useState('');
  const [wsWebhookUrl, setWsWebhookUrl] = useState('https://vhialy.dpdns.org/webhook/notify');
  const [wsNotifyEmail, setWsNotifyEmail] = useState('');

  // Header / Title Customization
  const [wsCompanyName, setWsCompanyName] = useState('CÔNG TY THỦY ĐIỆN IALY');
  const [wsHeaderWorkshopName, setWsHeaderWorkshopName] = useState('PHÂN XƯỞNG VẬN HÀNH IALY');
  const [wsDocumentCodeSuffix, setWsDocumentCodeSuffix] = useState('/VHIALY');
  const [wsRecipientWorkshopName, setWsRecipientWorkshopName] = useState('Phân xưởng vận hành Ialy');
  const [wsShortWorkshopName, setWsShortWorkshopName] = useState('PXVH Ialy');
  const [wsLocationName, setWsLocationName] = useState('Gia Lai');

  // Shift Schedule Settings
  const [shiftCa1Name, setShiftCa1Name] = useState('Ca 1 (Ca Ngày)');
  const [shiftCa1Time, setShiftCa1Time] = useState('08:00 - 16:00');
  const [shiftCa2Name, setShiftCa2Name] = useState('Ca 2 (Ca Chiều)');
  const [shiftCa2Time, setShiftCa2Time] = useState('16:00 - 22:20');
  const [shiftCa3Name, setShiftCa3Name] = useState('Ca 3 (Ca Đêm)');
  const [shiftCa3Time, setShiftCa3Time] = useState('22:20 - 08:00');
  const [teamsListText, setTeamsListText] = useState('Kíp 1, Kíp 2, Kíp 3, Kíp 4, Kíp 5');
  const [cycleLengthDays, setCycleLengthDays] = useState(5);
  const [shiftNote, setShiftNote] = useState('Lịch đi ca 3 ca 5 kíp xoay vòng liên tục.');

  // Custom Shift Schedule & Rules Matrices
  const [baseDate, setBaseDate] = useState('2025-10-01');
  const [shiftsMatrixText, setShiftsMatrixText] = useState(JSON.stringify(SHIFTS, null, 2));
  const [rulesMatrixText, setRulesMatrixText] = useState(JSON.stringify(RULES, null, 2));
  const [matrixMode, setMatrixMode] = useState<'visual' | 'json'>('visual');
  const [matrixUploadNotice, setMatrixUploadNotice] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const shiftMatrixFileInputRef = useRef<HTMLInputElement>(null);

  const getParsedShifts = (): string[][] => {
    try {
      const arr = JSON.parse(shiftsMatrixText);
      if (Array.isArray(arr) && arr.length > 0) return arr;
    } catch (e) {}
    return SHIFTS;
  };

  const validateShiftsMatrix = (matrix: string[][]): { isValid: boolean; invalidDaysDetails: { day: number; n: number; c: number; k: number; o: number }[]; message: string } => {
    if (!Array.isArray(matrix) || matrix.length === 0) {
      return { isValid: false, invalidDaysDetails: [], message: 'Ma trận ca làm việc trống hoặc không đúng định dạng.' };
    }

    const numTeams = matrix.length;
    const numDays = matrix[0]?.length || 0;
    if (numDays === 0) {
      return { isValid: false, invalidDaysDetails: [], message: 'Ma trận ca làm việc chưa có cột ngày nào.' };
    }

    const invalidDaysDetails: { day: number; n: number; c: number; k: number; o: number }[] = [];

    for (let d = 0; d < numDays; d++) {
      let n = 0, c = 0, k = 0, o = 0;
      for (let row = 0; row < numTeams; row++) {
        const val = (matrix[row]?.[d] || '').toUpperCase().trim();
        if (val === 'N') n++;
        else if (val === 'C') c++;
        else if (val === 'K') k++;
        else if (val === 'O') o++;
      }

      if (n !== 1 || c !== 1 || k !== 1) {
        invalidDaysDetails.push({ day: d + 1, n, c, k, o });
      }
    }

    if (invalidDaysDetails.length > 0) {
      const dayStr = invalidDaysDetails.map(item => `Ngày ${item.day} (Hiện có: ${item.n}N, ${item.c}C, ${item.k}K)`).join(', ');
      return {
        isValid: false,
        invalidDaysDetails,
        message: `Lỗi nhập dữ liệu ma trận ca trực: ${dayStr} bị thiếu hoặc thừa ca N, C, K! Mỗi ngày của ${numTeams} kíp phải có đúng 1 ca N, 1 ca C, 1 ca K (và ${numTeams - 3} ca O). Yêu cầu kiểm tra và nhập lại.`
      };
    }

    return { isValid: true, invalidDaysDetails: [], message: '' };
  };

  const validateRulesMatrix = (
    rulesObj: any,
    numTeams: number = 5
  ): { isValid: boolean; invalidRulesDetails: string[]; message: string } => {
    if (!rulesObj || typeof rulesObj !== 'object') {
      return {
        isValid: false,
        invalidRulesDetails: ['Toàn bộ quy luật trực thay bị trống hoặc không hợp lệ.'],
        message: 'Lỗi nhập quy luật trực thay tự động: Bảng quy luật bị trống hoặc không đúng định dạng!'
      };
    }

    const invalidRulesDetails: string[] = [];
    const shiftLabels: Record<string, string> = {
      N: 'Ca Ngày (N)',
      C: 'Ca Chiều (C)',
      K: 'Ca Đêm (K)'
    };

    const totalTeams = Math.max(numTeams || 5, 5);

    for (let kNum = 1; kNum <= totalTeams; kNum++) {
      const kipRule = rulesObj[kNum] || rulesObj[String(kNum)];
      if (!kipRule || typeof kipRule !== 'object') {
        invalidRulesDetails.push(`Kíp ${kNum} (Trống cả 3 ca N, C, K)`);
        continue;
      }

      ['N', 'C', 'K'].forEach((shiftType) => {
        const targetVal = kipRule[shiftType]?.k;
        if (!targetVal || typeof targetVal !== 'number' || targetVal <= 0) {
          invalidRulesDetails.push(`Kíp ${kNum} (${shiftLabels[shiftType] || shiftType})`);
        }
      });
    }

    if (invalidRulesDetails.length > 0) {
      const detailsStr = invalidRulesDetails.join(', ');
      return {
        isValid: false,
        invalidRulesDetails,
        message: `Lỗi nhập quy luật trực thay tự động: Chưa chọn kíp đi thay ở [${detailsStr}]! Tất cả các ca nghỉ (N, C, K) phải chọn kíp đi thay đầy đủ, không được để trống.`
      };
    }

    return { isValid: true, invalidRulesDetails: [], message: '' };
  };

  const getParsedRules = (): Record<number, Record<string, { k: number }>> => {
    try {
      const obj = JSON.parse(rulesMatrixText);
      if (obj && typeof obj === 'object') return obj;
    } catch (e) {}
    return RULES;
  };

  const handleShiftCellChange = (kipIdx: number, dayIdx: number, val: string) => {
    const current = getParsedShifts();
    const updated = current.map((row, rIdx) => {
      if (rIdx !== kipIdx) return [...row];
      const newRow = [...row];
      newRow[dayIdx] = val;
      return newRow;
    });
    setShiftsMatrixText(JSON.stringify(updated, null, 2));
  };

  const handleRuleChange = (kipNum: number, shiftType: string, targetKipNum: number) => {
    const current = getParsedRules();
    const updated = {
      ...current,
      [kipNum]: {
        ...(current[kipNum] || {}),
        [shiftType]: { k: targetKipNum }
      }
    };
    setRulesMatrixText(JSON.stringify(updated, null, 2));
  };

  // Download Excel template for Shift Schedule & Replacement Rules
  const handleDownloadShiftScheduleExcel = () => {
    const currentShifts = getParsedShifts();
    const currentRules = getParsedRules();

    // SHEET 1: SHIFTS Matrix
    const daysCount = currentShifts[0]?.length || 5;
    const shiftHeader = ['Kíp Trực', ...Array.from({ length: daysCount }, (_, i) => `Ngày ${i + 1}`)];
    const shiftRows = currentShifts.map((row, idx) => [
      `Kíp ${idx + 1}`,
      ...row
    ]);

    const shiftNotes = [
      [],
      ['Chú giải ký hiệu loại ca:'],
      ['N', 'Ca Ngày (08:00 - 16:00)'],
      ['C', 'Ca Chiều (16:00 - 22:20)'],
      ['K', 'Ca Đêm (22:20 - 08:00)'],
      ['O', 'Nghỉ Ca (Off)']
    ];

    const sheet1Data = [shiftHeader, ...shiftRows, ...shiftNotes];
    const sheet1 = XLSX.utils.aoa_to_sheet(sheet1Data);
    sheet1['!cols'] = [{ wch: 15 }, ...Array.from({ length: daysCount }, () => ({ wch: 12 }))];

    // SHEET 2: RULES Matrix
    const rulesHeader = ['Kíp Nghỉ', 'Nghỉ Ca Ngày (N) - Kíp Thay', 'Nghỉ Ca Chiều (C) - Kíp Thay', 'Nghỉ Ca Đêm (K) - Kíp Thay'];
    const rulesRows = [1, 2, 3, 4, 5].map((k) => {
      const r = currentRules[k] || { N: { k: 1 }, C: { k: 2 }, K: { k: 3 } };
      return [
        `Kíp ${k}`,
        `Kíp ${r.N?.k || 1}`,
        `Kíp ${r.C?.k || 1}`,
        `Kíp ${r.K?.k || 1}`
      ];
    });

    const rulesNotes = [
      [],
      ['Chú thích quy luật:'],
      ['Điền tên Kíp hoặc số Kíp thay thế tương ứng khi kíp chính nghỉ ở từng ca.']
    ];

    const sheet2Data = [rulesHeader, ...rulesRows, ...rulesNotes];
    const sheet2 = XLSX.utils.aoa_to_sheet(sheet2Data);
    sheet2['!cols'] = [{ wch: 15 }, { wch: 30 }, { wch: 30 }, { wch: 30 }];

    // Create Workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet1, 'Ma_Tran_Ca_Truc_SHIFTS');
    XLSX.utils.book_append_sheet(workbook, sheet2, 'Quy_Luat_Truc_Thay_RULES');

    XLSX.writeFile(workbook, 'Mau_Ma_Tran_Lich_Ca_Va_Quy_Luat_Truc_Thay.xlsx');
  };

  // Upload Excel file to update Shift Schedule & Replacement Rules
  const handleUploadShiftScheduleExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const workbook = XLSX.read(buffer, { type: 'array' });

        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          setMatrixUploadNotice({ type: 'error', text: 'File Excel không chứa dữ liệu.' });
          return;
        }

        let parsedShifts: string[][] | null = null;
        let parsedRules: Record<number, Record<string, { k: number }>> | null = null;

        // 1. Process Sheet for SHIFTS Matrix
        const shiftSheetName = workbook.SheetNames.find(s => 
          s.toLowerCase().includes('shift') || s.toLowerCase().includes('ca_truc') || s.toLowerCase().includes('ma_tran')
        ) || workbook.SheetNames[0];

        if (shiftSheetName) {
          const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[shiftSheetName], { header: 1 }) as any[][];
          if (sheetData && sheetData.length > 0) {
            const tempShiftsMap: Record<number, string[]> = {};
            let numDays = 0;

            sheetData.forEach((row) => {
              if (!row || row.length < 2) return;
              const firstCol = String(row[0] || '').trim().toLowerCase();

              // Detect header row (e.g. "Kíp Trực", "Kíp / Ngày", "Kíp") to count number of days
              if (numDays === 0 && (firstCol.includes('kíp') || firstCol.includes('kip'))) {
                const headerCells = row.slice(1).filter(c => String(c || '').trim().length > 0);
                if (headerCells.length > 0) {
                  numDays = headerCells.length;
                }
              }

              // Match strictly Kíp 1, Kíp 2, ... or Kip 1, Kip 2, ... or 1, 2, 3 ...
              const kipMatch = firstCol.match(/^(?:kíp|kip)\s*(\d+)$/i) || firstCol.match(/^(\d+)$/);
              if (kipMatch) {
                const kipNum = parseInt(kipMatch[1], 10);
                const dayCellsRaw = numDays > 0 ? row.slice(1, 1 + numDays) : row.slice(1);
                const dayCells = dayCellsRaw.map(cell => {
                  const val = String(cell || '').trim().toUpperCase();
                  if (val.startsWith('N')) return 'N';
                  if (val.startsWith('C')) return 'C';
                  if (val.startsWith('K')) return 'K';
                  if (val.startsWith('O') || val.includes('NGHĨ') || val.includes('OFF')) return 'O';
                  if (['N', 'C', 'K', 'O'].includes(val)) return val;
                  return 'O';
                });
                if (dayCells.length > 0) {
                  tempShiftsMap[kipNum] = dayCells;
                }
              }
            });

            const sortedKipKeys = Object.keys(tempShiftsMap).map(Number).sort((a, b) => a - b);
            if (sortedKipKeys.length >= 1) {
              const tempShifts = sortedKipKeys.map(k => tempShiftsMap[k]);
              const maxCols = Math.max(...tempShifts.map(r => r.length));
              const normalizedShifts = tempShifts.map(r => {
                const rowCopy = [...r];
                while (rowCopy.length < maxCols) rowCopy.push('O');
                return rowCopy;
              });
              parsedShifts = normalizedShifts;
            }
          }
        }

        // 2. Process Sheet for RULES Matrix
        const rulesSheetName = workbook.SheetNames.find(s => 
          s.toLowerCase().includes('rule') || s.toLowerCase().includes('quy_luat') || s.toLowerCase().includes('truc_thay')
        ) || (workbook.SheetNames.length > 1 ? workbook.SheetNames[1] : null);

        if (rulesSheetName && rulesSheetName !== shiftSheetName) {
          const rulesSheetData = XLSX.utils.sheet_to_json(workbook.Sheets[rulesSheetName], { header: 1 }) as any[][];
          if (rulesSheetData && rulesSheetData.length > 0) {
            const tempRules: Record<number, Record<string, { k: number }>> = {};

            rulesSheetData.forEach((row) => {
              if (!row || row.length < 4) return;
              const firstCol = String(row[0] || '').trim();
              const kipMatch = firstCol.match(/^(?:kíp|kip)\s*(\d+)$/i) || firstCol.match(/^(\d+)$/);
              if (kipMatch) {
                const kipNum = parseInt(kipMatch[1], 10);
                if (kipNum >= 1 && kipNum <= 10) {
                  const parseKipNumber = (val: any) => {
                    const numMatch = String(val || '').match(/\d+/);
                    return numMatch ? parseInt(numMatch[0], 10) : 1;
                  };

                  tempRules[kipNum] = {
                    N: { k: parseKipNumber(row[1]) },
                    C: { k: parseKipNumber(row[2]) },
                    K: { k: parseKipNumber(row[3]) }
                  };
                }
              }
            });

            if (Object.keys(tempRules).length > 0) {
              parsedRules = tempRules;
            }
          }
        }

        const updateMsgs: string[] = [];

        if (parsedShifts) {
          setShiftsMatrixText(JSON.stringify(parsedShifts, null, 2));
          updateMsgs.push('Ma trận Ca trực (SHIFTS)');
        }

        if (parsedRules) {
          setRulesMatrixText(JSON.stringify(parsedRules, null, 2));
          updateMsgs.push('Quy luật Trực thay (RULES)');
        }

        if (updateMsgs.length > 0) {
          let extraWarn = '';
          if (parsedShifts) {
            const shiftVal = validateShiftsMatrix(parsedShifts);
            if (!shiftVal.isValid) {
              extraWarn += ` ⚠️ CHÚ Ý CA TRỰC: ${shiftVal.message}`;
            }
          }
          if (parsedRules) {
            const ruleVal = validateRulesMatrix(parsedRules, parsedShifts ? parsedShifts.length : 5);
            if (!ruleVal.isValid) {
              extraWarn += ` ⚠️ CHÚ Ý QUY LUẬT: ${ruleVal.message}`;
            }
          }
          setMatrixUploadNotice({
            type: extraWarn ? 'error' : 'success',
            text: `✅ Đã tự động cập nhật ${updateMsgs.join(' và ')} từ file Excel!${extraWarn}`
          });
        } else {
          setMatrixUploadNotice({
            type: 'error',
            text: '⚠ Không thể nhận diện cấu trúc ma trận từ file Excel. Vui lòng tải file Mẫu Excel để làm theo chuẩn.'
          });
        }
      } catch (err: any) {
        setMatrixUploadNotice({
          type: 'error',
          text: 'Lỗi đọc file Excel: ' + (err?.message || 'File không đúng định dạng.')
        });
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Leave Rules Settings
  
  const [features, setFeatures] = useState<WorkshopFeatures>({
    enablePhanCa: true,
    enableDonNghiPhep: true,
    enableDoiCa: true,
    enableChuKySo: true,
    enableGoogleSheets: true,
    enableBaoComCa: true
  });

  const [staffDataText, setStaffDataText] = useState('');
  const [isSavingWs, setIsSavingWs] = useState(false);

  // New Account Form State
  const [newUsername, setNewUsername] = useState('');
  const [newPassword, setNewPassword] = useState('123456');
  const [newFullName, setNewFullName] = useState('');
  const [newRole, setNewRole] = useState<'super_admin' | 'workshop_admin' | 'workshop_user'>('workshop_user');
  const [newWsId, setNewWsId] = useState(isSuperAdmin ? (activeWorkshop?.id || workshops[0]?.id || '') : userWorkshopId);
  const [isCreatingAccount, setIsCreatingAccount] = useState(false);

  // Edit Account Form State
  const [editingAccount, setEditingAccount] = useState<UserAccount | null>(null);
  const [editUsername, setEditUsername] = useState('');
  const [editPassword, setEditPassword] = useState('');
  const [editFullName, setEditFullName] = useState('');
  const [editRole, setEditRole] = useState<'super_admin' | 'workshop_admin' | 'workshop_user'>('workshop_user');
  const [editWsId, setEditWsId] = useState('');
  const [isUpdatingAccount, setIsUpdatingAccount] = useState(false);

  // Account Filter & Search State
  const [accountSearch, setAccountSearch] = useState('');
  const [accountWsFilter, setAccountWsFilter] = useState('all');

  // Confirmation Modals State
  const [confirmDeleteAccount, setConfirmDeleteAccount] = useState<{ id: string; username: string } | null>(null);
  const [confirmDeleteWorkshop, setConfirmDeleteWorkshop] = useState<{ id: string; name: string } | null>(null);

  const [msg, setMsg] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  const visibleAccounts = isSuperAdmin 
    ? accounts 
    : accounts.filter(acc => acc.workshopId === userWorkshopId);

  const filteredAccounts = visibleAccounts.filter(acc => {
    const matchesQuery = 
      acc.username.toLowerCase().includes(accountSearch.toLowerCase()) ||
      acc.fullName.toLowerCase().includes(accountSearch.toLowerCase());
    const matchesWs = accountWsFilter === 'all' || acc.workshopId === accountWsFilter;
    return matchesQuery && matchesWs;
  });

  // Fetch accounts on load
  const fetchAccounts = async () => {
    setLoadingAccounts(true);
    try {
      const res = await fetch(API_BASE + '/api/accounts');
      if (res.ok) {
        const data = await res.json();
        setAccounts(data);
      }
    } catch (e) {
      console.error("Failed to load accounts", e);
    } finally {
      setLoadingAccounts(false);
    }
  };

  useEffect(() => {
    fetchAccounts();
  }, []);

  useEffect(() => {
    if (workshops.length > 0 && !createAdminSelectedWsId) {
      setCreateAdminSelectedWsId(workshops[0].id);
    }
  }, [workshops]);

  const handleCreateAdminAndSetup = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!createAdminUsername.trim() || !createAdminFullName.trim()) {
      setMsg({ type: 'error', text: 'Vui lòng nhập Tên đăng nhập và Họ tên Quản trị viên.' });
      return;
    }

    let targetWs: Workshop | null = null;
    setIsCreatingAdminAndSetup(true);
    setMsg(null);

    try {
      if (createAdminWsOption === 'new') {
        if (!createAdminWsName.trim()) {
          throw new Error('Vui lòng nhập Tên Phân xưởng mới.');
        }

        const wsRes = await fetch(API_BASE + '/api/workshops', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: createAdminWsName.trim(),
            code: createAdminWsCode.trim() || createAdminWsName.substring(0, 5).toUpperCase(),
            description: createAdminWsDesc.trim() || `Phân xưởng ${createAdminWsName.trim()}`,
            config: {
              companyName: 'CÔNG TY THỦY ĐIỆN IALY',
              headerWorkshopName: createAdminWsName.trim().toUpperCase(),
              documentCodeSuffix: createAdminWsCode ? `/${createAdminWsCode.trim()}` : '/PX',
              recipientWorkshopName: createAdminWsName.trim(),
              shortWorkshopName: createAdminWsName.trim(),
              locationName: 'Gia Lai',
              soVanBan: '',
              nguoiKy: 'Lãnh đạo Phân xưởng',
              chucVuNguoiKy: 'Quản đốc Phân xưởng',
              zaloWebhookUrl: 'https://vhialy.dpdns.org/webhook/notify',
              notifyEmail: ''
            }
          })
        });

        const wsData = await wsRes.json();
        if (!wsRes.ok || !wsData.success) {
          throw new Error(wsData.error || 'Lỗi khi tạo Phân xưởng mới.');
        }
        targetWs = wsData.workshop;
      } else {
        targetWs = workshops.find(w => w.id === (createAdminSelectedWsId || visibleWorkshops[0]?.id)) || null;
        if (!targetWs) {
          throw new Error('Vui lòng chọn một Phân xưởng hợp lệ.');
        }
      }

      // Step 2: Create Admin Account
      const accRes = await fetch(API_BASE + '/api/accounts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          username: createAdminUsername.trim(),
          password: createAdminPassword.trim() || '123456',
          fullName: createAdminFullName.trim(),
          role: 'workshop_admin',
          workshopId: targetWs.id
        })
      });

      const accData = await accRes.json();
      if (!accRes.ok || !accData.success) {
        throw new Error(accData.error || 'Lỗi khi tạo tài khoản Admin Phân xưởng.');
      }

      // Refresh list
      onRefreshWorkshops();
      await fetchAccounts();

      // Immediately open Initial Setup for targetWs
      openEditWorkshop(targetWs);

      setMsg({
        type: 'success',
        text: `🎉 Đã khởi tạo thành công tài khoản Admin "${createAdminUsername}" cho ${targetWs.name}! Mời bạn hoàn tất các thông tin thiết lập ban đầu bên dưới.`
      });

      // Clear fields
      setCreateAdminUsername('');
      setCreateAdminPassword('123456');
      setCreateAdminFullName('');
      setCreateAdminWsName('');
      setCreateAdminWsCode('');
      setCreateAdminWsDesc('');
    } catch (err: any) {
      setMsg({ type: 'error', text: err.message || 'Có lỗi xảy ra khi tạo Admin và Mở thiết lập.' });
    } finally {
      setIsCreatingAdminAndSetup(false);
    }
  };

  const openCreateWorkshop = () => {
    if (!isSuperAdmin) return;
    setEditingWsId(null);
    setWsName('');
    setWsCode('');
    setWsDesc('');
    setWsNguoiKy('Lãnh đạo Phân xưởng');
    setWsChucVu('Quản đốc Phân xưởng');
    setWsSoVanBan('');
    setWsWebhookUrl('https://vhialy.dpdns.org/webhook/notify');

    setWsCompanyName('CÔNG TY THỦY ĐIỆN IALY');
    setWsHeaderWorkshopName('PHÂN XƯỞNG VẬN HÀNH');
    setWsDocumentCodeSuffix('/PX');
    setWsRecipientWorkshopName('Phân xưởng Vận hành');
    setWsShortWorkshopName('PX Vận hành');

    setShiftCa1Name('Ca 1 (Ca Ngày)');
    setShiftCa1Time('08:00 - 16:00');
    setShiftCa2Name('Ca 2 (Ca Chiều)');
    setShiftCa2Time('16:00 - 22:20');
    setShiftCa3Name('Ca 3 (Ca Đêm)');
    setShiftCa3Time('22:20 - 08:00');
    setTeamsListText('Kíp 1, Kíp 2, Kíp 3, Kíp 4, Kíp 5');
    setCycleLengthDays(5);
    setShiftNote('Lịch đi ca 3 ca 5 kíp xoay vòng liên tục.');
    setBaseDate('2025-10-01');
    setShiftsMatrixText(JSON.stringify(SHIFTS, null, 2));
    setRulesMatrixText(JSON.stringify(RULES, null, 2));

    setFeatures({
      enablePhanCa: true,
      enableDonNghiPhep: true,
      enableDoiCa: true,
      enableChuKySo: true,
      enableGoogleSheets: true,
      enableBaoComCa: true
    });
    setStaffDataText('[]');
    setActiveTab('edit_workshop');
  };

  const openEditWorkshop = (ws: Workshop) => {
    if (!isSuperAdmin && ws.id !== userWorkshopId) {
      setMsg({ type: 'error', text: 'Bạn chỉ có quyền sửa đổi thông tin của phân xưởng bạn được phân công.' });
      return;
    }
    setEditingWsId(ws.id);
    setWsName(ws.name);
    setWsCode(ws.code);
    setWsDesc(ws.description || '');
    setWsNguoiKy(ws.config?.nguoiKy || 'Nguyễn Văn Nghị');
    setWsChucVu(ws.config?.chucVuNguoiKy || 'Quản đốc Phân xưởng');
    setWsSoVanBan(ws.config?.soVanBan !== undefined ? ws.config.soVanBan : '');
    setWsWebhookUrl(ws.config?.zaloWebhookUrl || 'https://vhialy.dpdns.org/webhook/notify');
    setWsNotifyEmail(ws.config?.notifyEmail || '');

    setWsCompanyName(ws.config?.companyName || 'CÔNG TY THỦY ĐIỆN IALY');
    setWsHeaderWorkshopName(ws.config?.headerWorkshopName || ws.name.toUpperCase());
    setWsDocumentCodeSuffix(ws.config?.documentCodeSuffix || (ws.code ? `/${ws.code}` : '/VH'));
    setWsRecipientWorkshopName(ws.config?.recipientWorkshopName || ws.name);
    setWsShortWorkshopName(ws.config?.shortWorkshopName || ws.code || ws.name);
    setWsLocationName(ws.config?.locationName || (ws.config as any)?.location || 'Gia Lai');

    setShiftCa1Name(ws.config?.shiftSchedule?.shiftCa1Name || 'Ca 1 (Ca Ngày)');
    setShiftCa1Time(ws.config?.shiftSchedule?.shiftCa1Time || '08:00 - 16:00');
    setShiftCa2Name(ws.config?.shiftSchedule?.shiftCa2Name || 'Ca 2 (Ca Chiều)');
    setShiftCa2Time(ws.config?.shiftSchedule?.shiftCa2Time || '16:00 - 22:20');
    setShiftCa3Name(ws.config?.shiftSchedule?.shiftCa3Name || 'Ca 3 (Ca Đêm)');
    setShiftCa3Time(ws.config?.shiftSchedule?.shiftCa3Time || '22:20 - 08:00');
    setTeamsListText(ws.config?.shiftSchedule?.teams?.join(', ') || 'Kíp 1, Kíp 2, Kíp 3, Kíp 4, Kíp 5');
    setCycleLengthDays(ws.config?.shiftSchedule?.cycleLengthDays || 5);
    setShiftNote(ws.config?.shiftSchedule?.shiftNote || 'Lịch đi ca 3 ca 5 kíp xoay vòng liên tục.');
    setBaseDate(ws.config?.shiftSchedule?.baseDate || '2025-10-01');
    setShiftsMatrixText(JSON.stringify(ws.config?.shiftSchedule?.shiftsMatrix || SHIFTS, null, 2));
    setRulesMatrixText(JSON.stringify(ws.config?.shiftSchedule?.rulesMatrix || RULES, null, 2));

    setFeatures(ws.features || {
      enablePhanCa: true,
      enableDonNghiPhep: true,
      enableDoiCa: true,
      enableChuKySo: true,
      enableGoogleSheets: true,
      enableBaoComCa: true
    });
    
    try {
      setStaffDataText(JSON.stringify(ws.staffData, null, 2));
    } catch (e) {
      setStaffDataText('[]');
    }

    setActiveTab('edit_workshop');
  };

  const handleSaveWorkshop = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!wsName.trim()) {
      setMsg({ type: 'error', text: 'Vui lòng nhập Tên Phân xưởng.' });
      return;
    }

    // Security enforce: non-super-admin can only edit their own workshop
    const targetWsId = isSuperAdmin ? editingWsId : userWorkshopId;

    let parsedStaff: string[][] = [];
    try {
      parsedStaff = JSON.parse(staffDataText);
      if (!Array.isArray(parsedStaff)) {
        throw new Error("Mảng nhân sự phải là mảng 2 chiều JSON valid.");
      }
    } catch (e: any) {
      setMsg({ type: 'error', text: 'Danh sách nhân sự không hợp lệ dạng JSON mảng 2 chiều. Vui lòng kiểm tra lại.' });
      return;
    }

    let parsedShifts: string[][] = SHIFTS;
    try {
      if (shiftsMatrixText.trim()) {
        parsedShifts = JSON.parse(shiftsMatrixText);
        if (!Array.isArray(parsedShifts)) {
          throw new Error("Ma trận ca làm việc phải là mảng JSON 2 chiều.");
        }
      }
    } catch (e: any) {
      setMsg({ type: 'error', text: 'Ma trận lịch ca (SHIFTS) không hợp lệ dạng JSON mảng 2 chiều.' });
      return;
    }

    const shiftValidation = validateShiftsMatrix(parsedShifts);
    if (!shiftValidation.isValid) {
      setMsg({ type: 'error', text: shiftValidation.message });
      return;
    }

    let parsedRules: Record<number, Record<string, { k: number }>> = RULES;
    try {
      if (rulesMatrixText.trim()) {
        parsedRules = JSON.parse(rulesMatrixText);
      }
    } catch (e: any) {
      setMsg({ type: 'error', text: 'Bảng quy luật kíp trực thay (RULES) không hợp lệ dạng JSON.' });
      return;
    }

    const ruleValidation = validateRulesMatrix(parsedRules, parsedShifts.length || 5);
    if (!ruleValidation.isValid) {
      setMsg({ type: 'error', text: ruleValidation.message });
      return;
    }

    setIsSavingWs(true);
    setMsg(null);

    try {
      const payload = {
        id: targetWsId,
        name: wsName.trim(),
        code: wsCode.trim() || wsName.substring(0, 5).toUpperCase(),
        description: wsDesc.trim(),
        staffData: parsedStaff,
        config: {
          companyName: wsCompanyName.trim(),
          headerWorkshopName: wsHeaderWorkshopName.trim(),
          documentCodeSuffix: wsDocumentCodeSuffix.trim(),
          recipientWorkshopName: wsRecipientWorkshopName.trim(),
          shortWorkshopName: wsShortWorkshopName.trim(),
          locationName: wsLocationName.trim() || 'Gia Lai',
          soVanBan: wsSoVanBan.trim(),
          ngayKy: '',
          nguoiKy: wsNguoiKy.trim(),
          chucVuNguoiKy: wsChucVu.trim(),
          zaloWebhookUrl: wsWebhookUrl.trim(),
          notifyEmail: wsNotifyEmail.trim(),
          shiftSchedule: {
            shiftCa1Name: shiftCa1Name.trim(),
            shiftCa1Time: shiftCa1Time.trim(),
            shiftCa2Name: shiftCa2Name.trim(),
            shiftCa2Time: shiftCa2Time.trim(),
            shiftCa3Name: shiftCa3Name.trim(),
            shiftCa3Time: shiftCa3Time.trim(),
            teams: teamsListText.split(',').map(t => t.trim()).filter(Boolean),
            cycleLengthDays: Number(cycleLengthDays) || 5,
            shiftNote: shiftNote.trim(),
            baseDate: baseDate.trim() || '2025-10-01',
            shiftsMatrix: parsedShifts,
            rulesMatrix: parsedRules
          }
        },
        features
      };

      const res = await fetch(API_BASE + '/api/workshops', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Không thể lưu phân xưởng.');
      }

      setMsg({ type: 'success', text: `✅ Đã lưu cấu hình phân xưởng "${wsName}" thành công!` });
      onRefreshWorkshops();
      setTimeout(() => {
        setActiveTab('workshops');
        setMsg(null);
      }, 1200);
    } catch (err: any) {
      setMsg({ type: 'error', text: err.message || 'Lỗi lưu dữ liệu.' });
    } finally {
      setIsSavingWs(false);
    }
  };

  const handleDeleteWorkshop = (id: string, name: string) => {
    if (!isSuperAdmin) return;
    setConfirmDeleteWorkshop({ id, name });
  };

  const executeDeleteWorkshop = async () => {
    if (!confirmDeleteWorkshop || !isSuperAdmin) return;
    const { id, name } = confirmDeleteWorkshop;
    setConfirmDeleteWorkshop(null);

    try {
      const res = await fetch(`/api/workshops/${id}`, { method: 'DELETE' });
      const data = await res.json().catch(() => ({}));
      if (res.ok) {
        setMsg({ type: 'success', text: `✅ Đã xóa phân xưởng "${name}" thành công.` });
        onRefreshWorkshops();
        fetchAccounts();
      } else {
        setMsg({ type: 'error', text: data.error || 'Lỗi khi xóa phân xưởng.' });
      }
    } catch (e) {
      setMsg({ type: 'error', text: 'Lỗi khi xóa phân xưởng.' });
    }
  };

  const handleCreateAccount = async (e: React.FormEvent) => {
    e.preventDefault();
    const targetWsId = isSuperAdmin ? newWsId : userWorkshopId;

    if (!newUsername.trim() || !newFullName.trim() || !targetWsId) {
      setMsg({ type: 'error', text: 'Vui lòng nhập tên đăng nhập, họ tên và chọn phân xưởng.' });
      return;
    }

    setIsCreatingAccount(true);
    setMsg(null);

    try {
      const res = await fetch(API_BASE + '/api/accounts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          username: newUsername.trim(),
          password: newPassword.trim() || '123456',
          fullName: newFullName.trim(),
          role: newRole,
          workshopId: targetWsId
        })
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Không thể tạo tài khoản.');
      }

      setMsg({ type: 'success', text: `✅ Đã tạo tài khoản "${newUsername}" cho phân xưởng thành công!` });
      setNewUsername('');
      setNewFullName('');
      setNewPassword('123456');
      fetchAccounts();
    } catch (err: any) {
      setMsg({ type: 'error', text: err.message || 'Lỗi tạo tài khoản.' });
    } finally {
      setIsCreatingAccount(false);
    }
  };

  const handleDeleteAccount = (accId: string, uname: string) => {
    if (uname === 'admin') {
      setMsg({ type: 'error', text: 'Không thể xóa tài khoản Quản trị hệ thống gốc (admin).' });
      return;
    }
    if (uname.toLowerCase() === user.username.toLowerCase()) {
      setMsg({ type: 'error', text: 'Bạn không thể tự xóa tài khoản đang đăng nhập của chính mình.' });
      return;
    }
    setConfirmDeleteAccount({ id: accId, username: uname });
  };

  const executeDeleteAccount = async () => {
    if (!confirmDeleteAccount) return;
    const { id, username: uname } = confirmDeleteAccount;
    setConfirmDeleteAccount(null);

    try {
      const res = await fetch(`/api/accounts/${id}`, { method: 'DELETE' });
      const data = await res.json().catch(() => ({}));
      if (res.ok) {
        setMsg({ type: 'success', text: `✅ Đã xóa tài khoản "${uname}" thành công.` });
        fetchAccounts();
      } else {
        setMsg({ type: 'error', text: data.error || 'Lỗi khi xóa tài khoản.' });
      }
    } catch (e) {
      setMsg({ type: 'error', text: 'Lỗi khi xóa tài khoản.' });
    }
  };

  const openEditAccount = (acc: UserAccount) => {
    setEditingAccount(acc);
    setEditUsername(acc.username);
    setEditPassword('');
    setEditFullName(acc.fullName);
    setEditRole(acc.role);
    setEditWsId(acc.workshopId);
  };

  const handleUpdateAccount = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingAccount) return;
    if (!editUsername.trim() || !editFullName.trim() || !editWsId) {
      setMsg({ type: 'error', text: 'Vui lòng nhập đủ tên đăng nhập, họ tên và phân xưởng.' });
      return;
    }

    setIsUpdatingAccount(true);
    setMsg(null);

    try {
      const payload: any = {
        id: editingAccount.id,
        username: editUsername.trim(),
        fullName: editFullName.trim(),
        role: editRole,
        workshopId: editWsId
      };
      if (editPassword.trim()) {
        payload.password = editPassword.trim();
      }

      const res = await fetch(API_BASE + '/api/accounts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Không thể cập nhật tài khoản.');
      }

      setMsg({ type: 'success', text: `✅ Đã cập nhật tài khoản "${editUsername}" thành công!` });
      setEditingAccount(null);
      fetchAccounts();
    } catch (err: any) {
      setMsg({ type: 'error', text: err.message || 'Lỗi khi cập nhật tài khoản.' });
    } finally {
      setIsUpdatingAccount(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/80 backdrop-blur-md p-3 sm:p-6 overflow-y-auto animate-in fade-in duration-200">
      <div className="bg-white w-full max-w-4xl rounded-2xl shadow-2xl border border-slate-200 my-auto overflow-hidden flex flex-col max-h-[90vh]">
        
        {/* Modal Header */}
        <div className="bg-slate-900 text-white p-5 flex items-center justify-between border-b border-slate-800 shrink-0">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-sky-600/20 text-sky-500 rounded-xl border border-sky-600/30">
              <Factory size={22} />
            </div>
            <div>
              <h3 className="font-bold text-base sm:text-lg tracking-wide text-white flex items-center gap-2">
                QUẢN LÝ PHÂN XƯỞNG & QUẢN TRỊ TÀI KHOẢN
              </h3>
              <p className="text-slate-400 text-xs mt-0.5">
                Cấu hình phân xưởng, bật/tắt chức năng và phân quyền tài khoản
              </p>
            </div>
          </div>
          <button 
            onClick={onClose}
            className="p-2 text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 rounded-xl transition-all cursor-pointer"
          >
            <X size={18} />
          </button>
        </div>

        {/* Modal Navigation Tabs */}
        <div className="bg-slate-100 border-b border-slate-200 px-5 pt-3 flex gap-2 shrink-0 overflow-x-auto">
          <button
            onClick={() => { setActiveTab('workshops'); setMsg(null); }}
            className={`px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all cursor-pointer flex items-center gap-2 ${
              activeTab === 'workshops'
                ? 'bg-white text-slate-900 border-t-2 border-sky-700 shadow-xs'
                : 'text-slate-600 hover:bg-slate-200/70'
            }`}
          >
            <Building2 size={16} />
            <span>Danh sách Phân xưởng</span>
            <span className="bg-sky-100 text-sky-800 px-1.5 py-0.2 rounded-full text-[11px]">
              {visibleWorkshops.length}
            </span>
          </button>

          <button
            onClick={() => { setActiveTab('accounts'); setMsg(null); }}
            className={`px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all cursor-pointer flex items-center gap-2 ${
              activeTab === 'accounts'
                ? 'bg-white text-slate-900 border-t-2 border-sky-700 shadow-xs'
                : 'text-slate-600 hover:bg-slate-200/70'
            }`}
          >
            <Users size={16} />
            <span>Tài khoản Đăng nhập</span>
            <span className="bg-slate-200 text-slate-700 px-1.5 py-0.2 rounded-full text-[11px]">
              {visibleAccounts.length}
            </span>
          </button>

          <button
            onClick={() => { setActiveTab('audit'); setMsg(null); loadAudit(); }}
            className={`px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all cursor-pointer flex items-center gap-2 ${
              activeTab === 'audit'
                ? 'bg-white text-slate-900 border-t-2 border-sky-700 shadow-xs'
                : 'text-slate-600 hover:bg-slate-200/70'
            }`}
          >
            <span>📜</span>
            <span>Nhật ký thao tác</span>
          </button>

          {isSuperAdmin && (
            <button
              onClick={() => { setActiveTab('create_admin'); setMsg(null); }}
              className={`px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all cursor-pointer flex items-center gap-1.5 ${
                activeTab === 'create_admin'
                  ? 'bg-white text-sky-950 border-t-2 border-sky-700 shadow-xs'
                  : 'text-sky-800 bg-sky-50/80 hover:bg-sky-100 border border-sky-200/80'
              }`}
            >
              <Sparkles size={16} className="text-sky-700" />
              <span>Tạo Admin & Thiết lập Ban đầu</span>
            </button>
          )}

          {isSuperAdmin && (
            <button
              onClick={openCreateWorkshop}
              className={`px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all cursor-pointer flex items-center gap-1.5 ml-auto text-sky-700 bg-sky-50 hover:bg-sky-100 border border-sky-200/80 ${
                activeTab === 'edit_workshop' && !editingWsId ? 'bg-sky-700 text-white border-sky-700' : ''
              }`}
            >
              <Plus size={16} />
              <span>Thêm Phân xưởng Mới</span>
            </button>
          )}
        </div>

        {/* Modal Body Content */}
        <div className="p-5 overflow-y-auto flex-1 space-y-4">
          {/* Notification Alert */}
          {msg && (
            <div className={`p-3 rounded-xl text-xs font-semibold flex items-center gap-2 ${
              msg.type === 'success' ? 'bg-sky-50 text-sky-800 border border-sky-200' : 'bg-rose-50 text-rose-800 border border-rose-200'
            }`}>
              <AlertCircle size={16} className="shrink-0" />
              <span>{msg.text}</span>
            </div>
          )}

          {/* TAB 1: WORKSHOPS LIST */}
          {activeTab === 'workshops' && (
            <div className="space-y-4">
              <div className="flex items-center justify-between">
                <p className="text-xs text-slate-600">
                  {isSuperAdmin 
                    ? 'Dưới đây là danh sách phân xưởng toàn hệ thống. Chọn chỉnh sửa để thay đổi nhân sự, chữ ký số hoặc tính năng.'
                    : 'Dưới đây là thông tin phân xưởng được phân công quản lý của bạn.'}
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {visibleWorkshops.map(ws => (
                  <div 
                    key={ws.id} 
                    className="p-4 bg-slate-50 hover:bg-slate-100/80 border border-slate-200 rounded-2xl transition-all space-y-3 flex flex-col justify-between"
                  >
                    <div>
                      <div className="flex items-start justify-between gap-2 mb-2">
                        <div>
                          <span className="text-[11px] font-extrabold uppercase tracking-wider bg-sky-100 text-sky-800 px-2 py-0.5 rounded-md">
                            {ws.code}
                          </span>
                          <h4 className="font-bold text-slate-900 text-sm mt-1">{ws.name}</h4>
                        </div>
                        <span className="text-[11px] text-slate-500 font-medium bg-white px-2 py-0.5 rounded-lg border border-slate-200">
                          {ws.staffData?.length || 0} Chức danh
                        </span>
                      </div>

                      <p className="text-xs text-slate-600 line-clamp-2">{ws.description || 'Chưa có mô tả'}</p>

                      {/* Feature Pills */}
                      <div className="flex flex-wrap gap-1.5 mt-3">
                        {ws.features?.enablePhanCa && (
                          <span className="text-[11px] bg-white border border-slate-200 text-slate-700 px-2 py-0.5 rounded-md">⚡ Phân ca</span>
                        )}
                        {ws.features?.enableDonNghiPhep && (
                          <span className="text-[11px] bg-white border border-slate-200 text-slate-700 px-2 py-0.5 rounded-md">📄 Đơn nghỉ phép</span>
                        )}
                        {ws.features?.enableDoiCa && (
                          <span className="text-[11px] bg-white border border-slate-200 text-slate-700 px-2 py-0.5 rounded-md">🔄 Đổi ca</span>
                        )}
                        {ws.features?.enableChuKySo && (
                          <span className="text-[11px] bg-white border border-slate-200 text-slate-700 px-2 py-0.5 rounded-md">✍️ Chữ ký số</span>
                        )}
                      </div>
                    </div>

                    <div className="pt-3 border-t border-slate-200/80 flex items-center justify-between gap-2">
                      <div className="text-[11px] text-slate-500">
                        Người ký: <span className="font-semibold text-slate-800">{ws.config?.nguoiKy || 'N/A'}</span>
                      </div>

                      <div className="flex items-center gap-1.5">
                        <button
                          onClick={() => openEditWorkshop(ws)}
                          className="px-2.5 py-1.5 text-xs font-semibold bg-white hover:bg-sky-50 text-sky-700 border border-slate-200 hover:border-sky-300 rounded-xl transition-all cursor-pointer flex items-center gap-1 shadow-xs"
                        >
                          <Edit2 size={13} />
                          <span>Sửa Cấu hình</span>
                        </button>
                        
                        {isSuperAdmin && (
                          <button
                            onClick={() => handleDeleteWorkshop(ws.id, ws.name)}
                            className="p-1.5 text-rose-600 hover:bg-rose-50 rounded-xl border border-transparent hover:border-rose-200 transition-all cursor-pointer"
                            title="Xóa phân xưởng"
                          >
                            <Trash2 size={14} />
                          </button>
                        )}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* TAB 2: ACCOUNTS MANAGEMENT */}
          {activeTab === 'accounts' && (
            <div className="space-y-6">
              {/* Quick Admin & Setup Launcher Banner */}
              {isSuperAdmin && (
                <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 p-3.5 bg-gradient-to-r from-sky-50 via-sky-50 to-sky-50 border border-sky-200/80 rounded-2xl shadow-xs">
                  <div className="flex items-center gap-2.5">
                    <Sparkles size={20} className="text-sky-700 shrink-0" />
                    <div>
                      <div className="font-bold text-slate-900 text-xs">Tạo mới Tài khoản Admin Phân Xưởng + Thiết lập Ban đầu?</div>
                      <div className="text-[11px] text-slate-600">Khởi tạo tài khoản Quản trị phân xưởng và tự động mở ngay khung thiết lập ban đầu chỉ trong 1 thao tác.</div>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => { setActiveTab('create_admin'); setMsg(null); }}
                    className="px-3.5 py-2 bg-sky-700 hover:bg-sky-700 text-white font-bold text-xs rounded-xl shadow-xs transition-all cursor-pointer flex items-center gap-1.5 shrink-0 self-end sm:self-auto"
                  >
                    <span>Tạo Admin & Thiết lập ngay</span>
                    <ArrowRight size={14} />
                  </button>
                </div>
              )}

              {/* Create Account Form */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-4">
                <div className="flex items-center gap-2 text-slate-900 font-bold text-sm">
                  <UserPlus size={18} className="text-sky-700" />
                  <span>
                    Tạo Tài khoản Đăng nhập Mới {isSuperAdmin ? 'cho Phân xưởng' : `cho ${workshops.find(w => w.id === userWorkshopId)?.name}`}
                  </span>
                </div>

                <form onSubmit={handleCreateAccount} className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 text-xs">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Tên đăng nhập *</label>
                    <input
                      type="text"
                      value={newUsername}
                      onChange={(e) => setNewUsername(e.target.value)}
                      placeholder="ví dụ: user_px_dien"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                      required
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Mật khẩu</label>
                    <input
                      type="text"
                      value={newPassword}
                      onChange={(e) => setNewPassword(e.target.value)}
                      placeholder="Mặc định: 123456"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20 font-mono"
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Họ và tên người dùng *</label>
                    <input
                      type="text"
                      value={newFullName}
                      onChange={(e) => setNewFullName(e.target.value)}
                      placeholder="ví dụ: Nguyễn Văn A"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                      required
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Cấp độ phân quyền *</label>
                    <select
                      value={newRole}
                      onChange={(e: any) => setNewRole(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    >
                      {isSuperAdmin && (
                        <option value="super_admin">Super Admin (Quản trị toàn hệ thống)</option>
                      )}
                      <option value="workshop_admin">Admin Phân xưởng (Quản lý & Cấu hình PX)</option>
                      <option value="workshop_user">Tài khoản Phân xưởng (Xem & Sử dụng chức năng)</option>
                    </select>
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Thuộc Phân xưởng *</label>
                    <select
                      value={isSuperAdmin ? newWsId : userWorkshopId}
                      onChange={(e) => isSuperAdmin && setNewWsId(e.target.value)}
                      disabled={!isSuperAdmin}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20 disabled:bg-slate-100 disabled:text-slate-600"
                    >
                      {visibleWorkshops.map(ws => (
                        <option key={ws.id} value={ws.id}>
                          {ws.name} ({ws.code})
                        </option>
                      ))}
                    </select>
                  </div>

                  <div className="flex items-end">
                    <button
                      type="submit"
                      disabled={isCreatingAccount}
                      className="w-full py-2 px-3 bg-sky-700 hover:bg-sky-700 text-white font-semibold rounded-xl transition-all cursor-pointer flex items-center justify-center gap-1.5 shadow-sm"
                    >
                      {isCreatingAccount ? 'Đang tạo...' : 'Tạo Tài khoản'}
                    </button>
                  </div>
                </form>
              </div>

              {/* Existing Accounts Table & Search/Filter */}
              <div>
                <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 mb-3">
                  <div className="flex items-center gap-2">
                    <h4 className="font-bold text-slate-900 text-sm">Danh sách Tài khoản Hiện có:</h4>
                    <span className="px-2 py-0.5 bg-sky-100 text-sky-800 text-[11px] font-bold rounded-full">
                      {filteredAccounts.length} / {visibleAccounts.length}
                    </span>
                  </div>

                  {/* Search and Workshop Filter Controls */}
                  <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
                    <div className="relative flex-1 sm:w-48">
                      <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Tìm tên đăng nhập/họ tên..."
                        value={accountSearch}
                        onChange={(e) => setAccountSearch(e.target.value)}
                        className="w-full pl-8 pr-3 py-1.5 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                      />
                    </div>

                    {isSuperAdmin && (
                      <select
                        value={accountWsFilter}
                        onChange={(e) => setAccountWsFilter(e.target.value)}
                        className="py-1.5 px-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                      >
                        <option value="all">Tất cả Phân xưởng</option>
                        {workshops.map(ws => (
                          <option key={ws.id} value={ws.id}>{ws.name}</option>
                        ))}
                      </select>
                    )}
                  </div>
                </div>

                <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-bold uppercase tracking-wider text-[11px]">
                        <th className="p-3">Tên đăng nhập</th>
                        <th className="p-3">Họ và tên</th>
                        <th className="p-3">Cấp quyền</th>
                        <th className="p-3">Phân xưởng</th>
                        <th className="p-3 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredAccounts.length === 0 ? (
                        <tr>
                          <td colSpan={5} className="p-6 text-center text-slate-500 text-xs italic">
                            Không tìm thấy tài khoản phù hợp với điều kiện lọc.
                          </td>
                        </tr>
                      ) : (
                        filteredAccounts.map(acc => {
                          const ws = workshops.find(w => w.id === acc.workshopId);
                          return (
                            <tr key={acc.id} className="hover:bg-slate-50 transition-all">
                              <td className="p-3 font-mono font-bold text-sky-800">{acc.username}</td>
                              <td className="p-3 font-semibold text-slate-900">{acc.fullName}</td>
                              <td className="p-3">
                                <span className={`px-2 py-0.5 rounded-md font-bold text-[11px] ${
                                  acc.role === 'super_admin'
                                    ? 'bg-amber-100 text-amber-800 border border-amber-200'
                                    : acc.role === 'workshop_admin'
                                      ? 'bg-sky-100 text-sky-800 border border-sky-200'
                                      : 'bg-slate-100 text-slate-700 border border-slate-200'
                                }`}>
                                  {acc.role === 'super_admin' ? 'Super Admin' : acc.role === 'workshop_admin' ? 'Admin PX' : 'User PX'}
                                </span>
                              </td>
                              <td className="p-3 font-medium text-slate-700">
                                {acc.workshopId === 'all' ? 'Tất cả Phân xưởng' : ws ? ws.name : acc.workshopId}
                              </td>
                              <td className="p-3 text-right flex items-center justify-end gap-1">
                                <button
                                  onClick={() => openEditAccount(acc)}
                                  className="px-2 py-1 bg-slate-100 hover:bg-sky-50 text-slate-700 hover:text-sky-700 border border-slate-200 rounded-lg transition-all cursor-pointer font-semibold flex items-center gap-1 text-[11px]"
                                  title="Chỉnh sửa thông tin / Đổi mật khẩu"
                                >
                                  <Edit2 size={13} />
                                  <span>Sửa</span>
                                </button>

                                {acc.username !== 'admin' && acc.username.toLowerCase() !== user.username.toLowerCase() && (
                                  <button
                                    onClick={() => handleDeleteAccount(acc.id, acc.username)}
                                    className="p-1 text-rose-600 hover:bg-rose-50 rounded-lg transition-all cursor-pointer"
                                    title="Xóa tài khoản"
                                  >
                                    <Trash2 size={14} />
                                  </button>
                                )}
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* EDIT ACCOUNT MODAL DIALOG */}
              {editingAccount && (
                <div className="fixed inset-0 z-60 flex items-center justify-center bg-slate-900/70 backdrop-blur-xs p-4 animate-in fade-in duration-150">
                  <div className="bg-white w-full max-w-lg rounded-2xl shadow-2xl border border-slate-200 overflow-hidden space-y-4">
                    <div className="bg-slate-900 text-white p-4 flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <Edit2 size={18} className="text-sky-500" />
                        <h4 className="font-bold text-sm">Chỉnh sửa Tài khoản: {editingAccount.username}</h4>
                      </div>
                      <button
                        onClick={() => setEditingAccount(null)}
                        className="p-1 hover:bg-slate-800 rounded-lg transition-all cursor-pointer text-slate-400 hover:text-white"
                      >
                        <X size={16} />
                      </button>
                    </div>

                    <form onSubmit={handleUpdateAccount} className="p-5 space-y-4 text-xs">
                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Tên đăng nhập *</label>
                        <input
                          type="text"
                          value={editUsername}
                          onChange={(e) => setEditUsername(e.target.value)}
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                          required
                        />
                      </div>

                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Họ và tên người dùng *</label>
                        <input
                          type="text"
                          value={editFullName}
                          onChange={(e) => setEditFullName(e.target.value)}
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                          required
                        />
                      </div>

                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Đặt lại Mật khẩu (để trống nếu giữ nguyên)</label>
                        <input
                          type="text"
                          value={editPassword}
                          onChange={(e) => setEditPassword(e.target.value)}
                          placeholder="Nhập mật khẩu mới..."
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20 font-mono"
                        />
                      </div>

                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Cấp độ phân quyền *</label>
                        <select
                          value={editRole}
                          onChange={(e: any) => setEditRole(e.target.value)}
                          disabled={!isSuperAdmin && editingAccount.role === 'super_admin'}
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20 disabled:bg-slate-100"
                        >
                          {isSuperAdmin && (
                            <option value="super_admin">Super Admin (Quản trị toàn hệ thống)</option>
                          )}
                          <option value="workshop_admin">Admin Phân xưởng (Quản lý & Cấu hình PX)</option>
                          <option value="workshop_user">Tài khoản Phân xưởng (Xem & Sử dụng chức năng)</option>
                        </select>
                      </div>

                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Phân xưởng trực thuộc *</label>
                        <select
                          value={isSuperAdmin ? editWsId : userWorkshopId}
                          onChange={(e) => isSuperAdmin && setEditWsId(e.target.value)}
                          disabled={!isSuperAdmin}
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20 disabled:bg-slate-100"
                        >
                          {isSuperAdmin && (
                            <option value="all">Tất cả Phân xưởng (all)</option>
                          )}
                          {visibleWorkshops.map(ws => (
                            <option key={ws.id} value={ws.id}>
                              {ws.name} ({ws.code})
                            </option>
                          ))}
                        </select>
                      </div>

                      <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
                        <button
                          type="button"
                          onClick={() => setEditingAccount(null)}
                          className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold rounded-xl transition-all cursor-pointer"
                        >
                          Hủy bỏ
                        </button>
                        <button
                          type="submit"
                          disabled={isUpdatingAccount}
                          className="px-5 py-2 bg-sky-700 hover:bg-sky-700 text-white font-semibold rounded-xl transition-all cursor-pointer flex items-center gap-1.5 shadow-sm"
                        >
                          {isUpdatingAccount ? 'Đang lưu...' : 'Lưu cập nhật'}
                        </button>
                      </div>
                    </form>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB: CREATE ADMIN & INITIAL SETUP */}
          {activeTab === 'audit' && (
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="font-bold text-slate-900 text-sm">Nhật ký thao tác</h4>
                  <p className="text-[11px] text-slate-500">
                    Ghi lại mọi thao tác thay đổi dữ liệu: ai làm, lúc nào, kết quả ra sao.
                    {isSuperAdmin ? ' (Super Admin thấy toàn hệ thống)' : ' (Chỉ phân xưởng của bạn)'}
                  </p>
                </div>
                <button
                  type="button"
                  onClick={loadAudit}
                  className="px-3 py-2 rounded-xl border border-slate-200 text-xs font-bold text-slate-700 hover:bg-slate-50 cursor-pointer"
                >
                  ↻ Làm mới
                </button>
              </div>

              {auditLoading ? (
                <div className="py-8 text-center text-slate-400 text-xs font-medium">Đang tải nhật ký...</div>
              ) : auditRows.length === 0 ? (
                <div className="py-8 text-center text-slate-400 text-xs font-medium">Chưa có thao tác nào được ghi nhận.</div>
              ) : (
                <div className="table-scroll border border-slate-200 rounded-xl max-h-[460px] overflow-auto">
                  <table className="w-full min-w-[760px] text-[11px]">
                    <thead className="bg-slate-50 sticky top-0">
                      <tr className="text-left text-slate-600">
                        <th className="px-3 py-2 font-bold whitespace-nowrap">Thời gian</th>
                        <th className="px-3 py-2 font-bold whitespace-nowrap">Người thực hiện</th>
                        <th className="px-3 py-2 font-bold whitespace-nowrap">Thao tác</th>
                        <th className="px-3 py-2 font-bold">Chi tiết</th>
                        <th className="px-3 py-2 font-bold whitespace-nowrap">Kết quả</th>
                      </tr>
                    </thead>
                    <tbody>
                      {auditRows.map((r) => (
                        <tr key={r.id} className="border-t border-slate-100 hover:bg-slate-50/70">
                          <td className="px-3 py-1.5 whitespace-nowrap text-slate-500">
                            {new Date(r.at).toLocaleString('vi-VN')}
                          </td>
                          <td className="px-3 py-1.5 whitespace-nowrap">
                            <span className="font-semibold text-slate-800">{r.username || '—'}</span>
                            {r.role === 'super_admin' && (
                              <span className="ml-1 text-[11px] bg-amber-100 text-amber-800 px-1.5 rounded-full font-bold">SA</span>
                            )}
                          </td>
                          <td className="px-3 py-1.5 whitespace-nowrap font-mono text-slate-600">
                            {r.method} {r.path}
                          </td>
                          <td className="px-3 py-1.5 text-slate-600 break-words">{r.summary || ''}</td>
                          <td className="px-3 py-1.5 whitespace-nowrap">
                            <span className={`px-1.5 py-0.5 rounded-full font-bold ${
                              r.status >= 200 && r.status < 300
                                ? 'bg-sky-100 text-sky-800'
                                : 'bg-rose-100 text-rose-800'
                            }`}>
                              {r.status}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {activeTab === 'create_admin' && (
            <div className="space-y-5 text-xs">
              {/* Header Banner */}
              <div className="p-4 bg-gradient-to-r from-sky-700 via-sky-700 to-sky-700 text-white rounded-2xl shadow-md space-y-1">
                <div className="flex items-center gap-2">
                  <Sparkles size={20} className="text-amber-300 animate-pulse" />
                  <h4 className="font-extrabold text-base">Tạo mới Tài khoản Admin Phân xưởng & Mở Thiết lập Ban đầu</h4>
                </div>
                <p className="text-sky-100 text-xs">
                  Điền thông tin tài khoản quản trị bên dưới. Ngay khi bấm "Tạo & Mở thiết lập", hệ thống sẽ khởi tạo tài khoản và tự động mở hộp thoại cấu hình thiết lập ban đầu cho phân xưởng đó.
                </p>
              </div>

              <form onSubmit={handleCreateAdminAndSetup} className="space-y-4">
                {/* Section 1: Thông tin tài khoản Admin */}
                <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3">
                  <div className="flex items-center gap-2 font-bold text-slate-800 text-sm border-b border-slate-200 pb-2">
                    <UserPlus size={16} className="text-sky-700" />
                    <span>1. Thông tin Tài khoản Quản trị (Admin Phân xưởng)</span>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        Tên đăng nhập * <span className="text-slate-400 font-normal">(dùng đăng nhập)</span>
                      </label>
                      <input
                        type="text"
                        value={createAdminUsername}
                        onChange={(e) => setCreateAdminUsername(e.target.value)}
                        placeholder="ví dụ: admin_vh_pleikrong"
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                        required
                      />
                    </div>

                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        Mật khẩu * <span className="text-slate-400 font-normal">(mặc định: 123456)</span>
                      </label>
                      <input
                        type="text"
                        value={createAdminPassword}
                        onChange={(e) => setCreateAdminPassword(e.target.value)}
                        placeholder="123456"
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-mono font-medium focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                        required
                      />
                    </div>

                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        Họ và tên Quản trị viên *
                      </label>
                      <input
                        type="text"
                        value={createAdminFullName}
                        onChange={(e) => setCreateAdminFullName(e.target.value)}
                        placeholder="ví dụ: Nguyễn Văn Quản Lý"
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                        required
                      />
                    </div>
                  </div>
                </div>

                {/* Section 2: Chọn / Tạo Phân xưởng */}
                <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3">
                  <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                    <div className="flex items-center gap-2 font-bold text-slate-800 text-sm">
                      <Building2 size={16} className="text-sky-700" />
                      <span>2. Phân xưởng Áp dụng cho Tài khoản Admin này</span>
                    </div>

                    {/* Mode Selector */}
                    <div className="flex items-center gap-1 bg-white p-1 rounded-xl border border-slate-200">
                      <button
                        type="button"
                        onClick={() => setCreateAdminWsOption('new')}
                        className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                          createAdminWsOption === 'new'
                            ? 'bg-sky-700 text-white shadow-xs'
                            : 'text-slate-600 hover:bg-slate-100'
                        }`}
                      >
                        ➕ Tạo Phân xưởng Mới
                      </button>
                      <button
                        type="button"
                        onClick={() => setCreateAdminWsOption('existing')}
                        className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                          createAdminWsOption === 'existing'
                            ? 'bg-sky-700 text-white shadow-xs'
                            : 'text-slate-600 hover:bg-slate-100'
                        }`}
                      >
                        🏢 Chọn Phân xưởng Đã có
                      </button>
                    </div>
                  </div>

                  {createAdminWsOption === 'new' ? (
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-3 pt-1">
                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Tên Phân xưởng mới *</label>
                        <input
                          type="text"
                          value={createAdminWsName}
                          onChange={(e) => setCreateAdminWsName(e.target.value)}
                          placeholder="ví dụ: Phân xưởng Vận hành Pleikrông"
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                          required={createAdminWsOption === 'new'}
                        />
                      </div>

                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Mã Phân xưởng</label>
                        <input
                          type="text"
                          value={createAdminWsCode}
                          onChange={(e) => setCreateAdminWsCode(e.target.value)}
                          placeholder="ví dụ: VHPK"
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold uppercase focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                        />
                      </div>

                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Mô tả / Ghi chú</label>
                        <input
                          type="text"
                          value={createAdminWsDesc}
                          onChange={(e) => setCreateAdminWsDesc(e.target.value)}
                          placeholder="Mô tả chức năng nhiệm vụ..."
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                        />
                      </div>
                    </div>
                  ) : (
                    <div className="pt-1">
                      <label className="block font-semibold text-slate-700 mb-1">Chọn Phân xưởng từ danh sách *</label>
                      <select
                        value={createAdminSelectedWsId}
                        onChange={(e) => setCreateAdminSelectedWsId(e.target.value)}
                        className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                      >
                        {workshops.map(ws => (
                          <option key={ws.id} value={ws.id}>
                            {ws.name} ({ws.code})
                          </option>
                        ))}
                      </select>
                    </div>
                  )}
                </div>

                {/* Action Submit */}
                <div className="pt-2 flex justify-end">
                  <button
                    type="submit"
                    disabled={isCreatingAdminAndSetup}
                    className="w-full sm:w-auto px-6 py-3 bg-sky-700 hover:bg-sky-700 disabled:opacity-50 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all cursor-pointer flex items-center justify-center gap-2 active:scale-98"
                  >
                    {isCreatingAdminAndSetup ? (
                      <span>Đang khởi tạo tài khoản & mở thiết lập...</span>
                    ) : (
                      <>
                        <Sparkles size={16} />
                        <span>🚀 Tạo Tài Khoản Admin & Mở Khung Thiết Lập Ban Đầu</span>
                      </>
                    )}
                  </button>
                </div>
              </form>
            </div>
          )}

          {/* TAB 3: CREATE / EDIT WORKSHOP FORM */}
          {activeTab === 'edit_workshop' && (
            <form onSubmit={handleSaveWorkshop} className="space-y-5 text-xs">
              <div className="p-4 bg-sky-50/60 border border-sky-200/80 rounded-2xl flex items-center justify-between">
                <div>
                  <h4 className="font-bold text-sky-950 text-sm">
                    {editingWsId ? 'Cấu hình Lịch Ca, Quy luật Nghỉ phép & Văn bản Phân xưởng' : 'Tạo Phân xưởng Mới & Thiết lập Chức năng'}
                  </h4>
                  <p className="text-sky-700 text-xs">
                    Tùy chỉnh lịch đi ca riêng, quy luật nghỉ phép, tiêu đề xuất Word/PDF và danh sách nhân sự của phân xưởng.
                  </p>
                </div>
              </div>

              {/* Basic Info */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Tên Phân xưởng *</label>
                  <input
                    type="text"
                    value={wsName}
                    onChange={(e) => setWsName(e.target.value)}
                    placeholder="ví dụ: Phân xưởng Sửa chữa"
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    required
                  />
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Mã Phân xưởng *</label>
                  <input
                    type="text"
                    value={wsCode}
                    onChange={(e) => setWsCode(e.target.value)}
                    placeholder="ví dụ: PX-SC"
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-sky-600/20 uppercase"
                  />
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Mô tả / Ghi chú</label>
                  <input
                    type="text"
                    value={wsDesc}
                    onChange={(e) => setWsDesc(e.target.value)}
                    placeholder="Mô tả chức năng nhiệm vụ..."
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                  />
                </div>
              </div>

              {/* 1. Document Header Customization */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3">
                <h5 className="font-bold text-slate-900 text-xs uppercase tracking-wider text-sky-800">
                  1. Tùy chỉnh Tiêu đề & Tên Phân xưởng trên Văn bản (Word/PDF)
                </h5>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Tên Công ty </label>
                    <input
                      type="text"
                      value={wsCompanyName}
                      onChange={(e) => setWsCompanyName(e.target.value)}
                      placeholder="CÔNG TY THỦY ĐIỆN IALY"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    />
                  </div>

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Tên Phân xưởng In Tiêu Đề Quốc Hiệu </label>
                    <input
                      type="text"
                      value={wsHeaderWorkshopName}
                      onChange={(e) => setWsHeaderWorkshopName(e.target.value)}
                      placeholder="PHÂN XƯỞNG SỬA CHỮA"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20 uppercase"
                    />
                  </div>

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Hậu tố Số Văn Bản (VD: /PXSC)</label>
                    <input
                      type="text"
                      value={wsDocumentCodeSuffix}
                      onChange={(e) => setWsDocumentCodeSuffix(e.target.value)}
                      placeholder="/PXSC"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    />
                  </div>

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Tên Kính Gửi trong Đơn</label>
                    <input
                      type="text"
                      value={wsRecipientWorkshopName}
                      onChange={(e) => setWsRecipientWorkshopName(e.target.value)}
                      placeholder="Phân xưởng Sửa chữa"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    />
                  </div>

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Tên Viết Tắt Nơi Công Tác</label>
                    <input
                      type="text"
                      value={wsShortWorkshopName}
                      onChange={(e) => setWsShortWorkshopName(e.target.value)}
                      placeholder="PX Sửa chữa"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    />
                  </div>

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Địa điểm / Tỉnh thành (Xuất văn bản)</label>
                    <input
                      type="text"
                      value={wsLocationName}
                      onChange={(e) => setWsLocationName(e.target.value)}
                      placeholder="Ví dụ: Gia Lai, Kon Tum, Đắk Lắk..."
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    />
                  </div>
                </div>
              </div>

              {/* 2. Shift Schedule Config */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3">
                <h5 className="font-bold text-slate-900 text-xs uppercase tracking-wider text-sky-800">
                  2. Cấu hình Lịch Đi Ca & Kíp Trực Riêng của Phân xưởng
                </h5>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Tên Ca 1 & Giờ làm việc</label>
                    <div className="grid grid-cols-2 gap-1.5">
                      <input
                        type="text"
                        value={shiftCa1Name}
                        onChange={(e) => setShiftCa1Name(e.target.value)}
                        placeholder="Ca 1 (Ngày)"
                        className="px-2 py-1.5 bg-white border border-slate-200 rounded-xl text-xs"
                      />
                      <input
                        type="text"
                        value={shiftCa1Time}
                        onChange={(e) => setShiftCa1Time(e.target.value)}
                        placeholder="08:00 - 16:00"
                        className="px-2 py-1.5 bg-white border border-slate-200 rounded-xl text-xs"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Tên Ca 2 & Giờ làm việc</label>
                    <div className="grid grid-cols-2 gap-1.5">
                      <input
                        type="text"
                        value={shiftCa2Name}
                        onChange={(e) => setShiftCa2Name(e.target.value)}
                        placeholder="Ca 2 (Chiều)"
                        className="px-2 py-1.5 bg-white border border-slate-200 rounded-xl text-xs"
                      />
                      <input
                        type="text"
                        value={shiftCa2Time}
                        onChange={(e) => setShiftCa2Time(e.target.value)}
                        placeholder="16:00 - 22:20"
                        className="px-2 py-1.5 bg-white border border-slate-200 rounded-xl text-xs"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Tên Ca 3 & Giờ làm việc</label>
                    <div className="grid grid-cols-2 gap-1.5">
                      <input
                        type="text"
                        value={shiftCa3Name}
                        onChange={(e) => setShiftCa3Name(e.target.value)}
                        placeholder="Ca 3 (Đêm)"
                        className="px-2 py-1.5 bg-white border border-slate-200 rounded-xl text-xs"
                      />
                      <input
                        type="text"
                        value={shiftCa3Time}
                        onChange={(e) => setShiftCa3Time(e.target.value)}
                        placeholder="22:20 - 08:00"
                        className="px-2 py-1.5 bg-white border border-slate-200 rounded-xl text-xs"
                      />
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pt-2">
                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Danh sách Kíp Trực </label>
                    <input
                      type="text"
                      value={teamsListText}
                      onChange={(e) => setTeamsListText(e.target.value)}
                      placeholder="Kíp 1, Kíp 2, Kíp 3, Kíp 4, Kíp 5"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    />
                  </div>

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Chu kỳ xoay ca (ngày)</label>
                    <input
                      type="number"
                      min={1}
                      max={30}
                      value={cycleLengthDays}
                      onChange={(e) => setCycleLengthDays(Number(e.target.value))}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    />
                  </div>
                </div>

                <div>
                  <label className="block font-medium text-slate-700 mb-1">Ghi chú / Quy chế đi ca</label>
                  <input
                    type="text"
                    value={shiftNote}
                    onChange={(e) => setShiftNote(e.target.value)}
                    placeholder="Lịch đi ca 3 ca 5 kíp xoay vòng..."
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                  />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pt-3 border-t border-slate-200/80 mt-2">
                  <div>
                    <label className="block font-semibold text-slate-800 mb-1">
                      Ngày Gốc Bắt Đầu Chu Kỳ Ca Trực *
                    </label>
                    <input
                      type="date"
                      value={baseDate}
                      onChange={(e) => setBaseDate(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-mono font-bold text-sky-800 focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                      required
                    />
                    <p className="text-[11px] text-slate-500 mt-0.5">
                      Ngày chuẩn mốc 0 để tính toán ca trực "N", "C", "K", "O" cho các kíp.
                    </p>
                  </div>
                </div>

                <div className="pt-3 border-t border-slate-200/80 mt-2 space-y-3">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-sky-50/80 p-2.5 rounded-xl border border-sky-100">
                    <div>
                      <h6 className="font-bold text-sky-950 text-xs uppercase tracking-wider flex items-center gap-1.5">
                        <Sparkles size={14} className="text-sky-700" /> Bảng Bấm Chọn Lịch Ca & Kíp Trực Thay
                      </h6>
                      <p className="text-[11px] text-sky-800">
                        Thao tác dạng Bảng Bấm trực quan dễ sử dụng cho người dùng không biết lập trình.
                      </p>
                    </div>

                    <div className="flex bg-slate-200/80 p-0.5 rounded-xl text-xs font-medium self-start sm:self-auto">
                      <button
                        type="button"
                        onClick={() => setMatrixMode('visual')}
                        className={`px-3 py-1 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer ${
                          matrixMode === 'visual'
                            ? 'bg-sky-700 text-white font-semibold shadow-sm'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        <LayoutGrid size={13} /> Bảng Chọn Trực Quan
                      </button>
                      <button
                        type="button"
                        onClick={() => setMatrixMode('json')}
                        className={`px-3 py-1 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer ${
                          matrixMode === 'json'
                            ? 'bg-slate-800 text-white font-semibold shadow-sm'
                            : 'text-slate-600 hover:text-slate-900'
                        }`}
                      >
                        <Code size={13} /> Mã JSON (Nâng cao)
                      </button>
                    </div>
                  </div>

                  {/* Excel Download & Upload Action Toolbar for Shift Matrix */}
                  <div className="flex flex-wrap items-center justify-between gap-2 p-2.5 bg-slate-50 rounded-xl border border-slate-200 text-xs">
                    <div className="flex items-center gap-2">
                      <FileSpreadsheet size={16} className="text-sky-700" />
                      <span className="font-semibold text-slate-700">Thao tác File Excel Ma Trận & Quy Luật:</span>
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={handleDownloadShiftScheduleExcel}
                        className="px-3 py-1.5 bg-sky-700 hover:bg-sky-700 text-white rounded-lg font-medium shadow-sm transition-all flex items-center gap-1.5 cursor-pointer"
                        title="Tải xuống file Mẫu Excel gồm Ma trận ca trực & Quy luật trực thay"
                      >
                        <Download size={13} /> Tải Mẫu Excel Ma Trận
                      </button>

                      <button
                        type="button"
                        onClick={() => shiftMatrixFileInputRef.current?.click()}
                        className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium shadow-sm transition-all flex items-center gap-1.5 cursor-pointer"
                        title="Upload file Excel đã điền để tự động cập nhật"
                      >
                        <Upload size={13} /> Upload Excel Tự Cập Nhật
                      </button>

                      <input
                        ref={shiftMatrixFileInputRef}
                        type="file"
                        accept=".xlsx, .xls, .csv"
                        className="hidden"
                        onChange={handleUploadShiftScheduleExcel}
                      />
                    </div>
                  </div>

                  {matrixUploadNotice && (
                    <div className={`p-3 rounded-xl border text-xs flex items-center justify-between ${
                      matrixUploadNotice.type === 'success' 
                        ? 'bg-sky-50 border-sky-200 text-sky-800 font-medium' 
                        : 'bg-rose-50 border-rose-200 text-rose-800 font-medium'
                    }`}>
                      <span>{matrixUploadNotice.text}</span>
                      <button
                        type="button"
                        onClick={() => setMatrixUploadNotice(null)}
                        className="text-slate-400 hover:text-slate-600 ml-2 text-sm font-bold cursor-pointer"
                      >
                        ✕
                      </button>
                    </div>
                  )}

                  {matrixMode === 'visual' ? (
                    <div className="space-y-4 pt-1">
                      {/* 1. Visual SHIFTS Matrix */}
                      <div className="bg-white p-3.5 rounded-2xl border border-slate-200/80 shadow-sm">
                        <div className="flex items-center justify-between mb-2">
                          <div>
                            <h6 className="font-bold text-slate-800 text-xs flex items-center gap-1.5">
                              1. Ma Trận Chu Kỳ Ca Trực 
                            </h6>
                            <p className="text-[11px] text-slate-500">
                              Chọn loại ca trực cho từng Kíp theo các ngày trong chu kỳ.
                            </p>
                          </div>
                          <button
                            type="button"
                            onClick={() => setShiftsMatrixText(JSON.stringify(SHIFTS, null, 2))}
                            className="text-[11px] font-semibold text-sky-700 hover:text-sky-900 bg-sky-50 hover:bg-sky-100 px-2.5 py-1 rounded-lg transition-all flex items-center gap-1 cursor-pointer"
                          >
                            <RefreshCw size={11} /> Khôi phục Mặc định
                          </button>
                        </div>

                        <div className="overflow-x-auto border border-slate-200 rounded-xl">
                          <table className="w-full text-xs text-left border-collapse">
                            <thead>
                              <tr className="bg-slate-100/90 text-slate-700 font-bold border-b border-slate-200">
                                <th className="p-2.5 w-24">Kíp Trực</th>
                                {Array.from({ length: getParsedShifts()[0]?.length || 5 }).map((_, dIdx) => (
                                  <th key={dIdx} className="p-2.5 text-center">
                                    Ngày {dIdx + 1}
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                              {getParsedShifts().map((row, kIdx) => (
                                <tr key={kIdx} className="hover:bg-slate-50/50">
                                  <td className="p-2.5 font-bold text-slate-800 bg-slate-50/80">
                                    Kíp {kIdx + 1}
                                  </td>
                                  {row.map((cellVal, dIdx) => (
                                    <td key={dIdx} className="p-2 text-center">
                                      <select
                                        value={cellVal}
                                        onChange={(e) => handleShiftCellChange(kIdx, dIdx, e.target.value)}
                                        className={`px-2 py-1.5 rounded-lg font-bold text-xs border text-center cursor-pointer transition-all focus:outline-none focus:ring-2 ${
                                          cellVal === 'N'
                                            ? 'bg-sky-100 text-sky-800 border-sky-300 focus:ring-sky-500'
                                            : cellVal === 'C'
                                            ? 'bg-amber-100 text-amber-800 border-amber-300 focus:ring-amber-400'
                                            : cellVal === 'K'
                                            ? 'bg-indigo-100 text-indigo-800 border-indigo-300 focus:ring-indigo-400'
                                            : 'bg-slate-100 text-slate-600 border-slate-300 focus:ring-slate-400'
                                        }`}
                                      >
                                        <option value="N">N - Ca Ngày (08:00 - 16:00)</option>
                                        <option value="C">C - Ca Chiều (16:00 - 22:20)</option>
                                        <option value="K">K - Ca Đêm (22:20 - 08:00)</option>
                                        <option value="O">O - Nghỉ Ca (Off)</option>
                                      </select>
                                    </td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                            <tfoot>
                              <tr className="bg-slate-50 border-t-2 border-slate-200 text-[11px] font-bold">
                                <td className="p-2.5 text-slate-700 bg-slate-100/80">Kiểm tra ca / ngày</td>
                                {Array.from({ length: getParsedShifts()[0]?.length || 5 }).map((_, dIdx) => {
                                  const currentShifts = getParsedShifts();
                                  let n = 0, c = 0, k = 0, o = 0;
                                  currentShifts.forEach(row => {
                                    const val = (row[dIdx] || '').toUpperCase().trim();
                                    if (val === 'N') n++;
                                    else if (val === 'C') c++;
                                    else if (val === 'K') k++;
                                    else if (val === 'O') o++;
                                  });

                                  const isOk = (n === 1 && c === 1 && k === 1);
                                  return (
                                    <td key={dIdx} className="p-2 text-center">
                                      {isOk ? (
                                        <span className="inline-flex items-center justify-center gap-1 px-1.5 py-1 rounded-md bg-sky-100 text-sky-800 border border-sky-300 font-bold text-[11px]" title="Đủ 1N, 1C, 1K">
                                          <Check size={12} /> 1N-1C-1K
                                        </span>
                                      ) : (
                                        <span className="inline-flex items-center justify-center gap-1 px-1.5 py-1 rounded-md bg-rose-100 text-rose-800 border border-rose-300 font-bold text-[11px]" title={`Thừa/thiếu ca! Hiện có: ${n}N, ${c}C, ${k}K, ${o}O`}>
                                          <AlertCircle size={12} /> Lỗi (N:{n} C:{c} K:{k})
                                        </span>
                                      )}
                                    </td>
                                  );
                                })}
                              </tr>
                            </tfoot>
                          </table>
                        </div>

                        {/* Validation Banner */}
                        {(() => {
                          const currentVal = validateShiftsMatrix(getParsedShifts());
                          if (currentVal.isValid) {
                            return (
                              <div className="mt-2.5 p-2.5 bg-sky-50 border border-sky-200 text-sky-800 rounded-xl text-xs font-semibold flex items-center gap-2">
                                <Check size={16} className="text-sky-700 shrink-0" />
                                <span>Ma trận ca trực hợp lệ: Mỗi ngày đều có đúng <b>1 ca N, 1 ca C, 1 ca K</b> (và {getParsedShifts().length - 3} ca O).</span>
                              </div>
                            );
                          } else {
                            return (
                              <div className="mt-2.5 p-3 bg-rose-50 border border-rose-200 text-rose-800 rounded-xl text-xs font-medium space-y-1">
                                <div className="flex items-center gap-1.5 font-bold text-rose-900 text-xs">
                                  <AlertCircle size={16} className="text-rose-600 shrink-0" />
                                  <span>Cảnh báo lỗi nhập ma trận ca trực:</span>
                                </div>
                                <p className="text-[11px] leading-relaxed font-semibold">{currentVal.message}</p>
                              </div>
                            );
                          }
                        })()}

                        <div className="flex flex-wrap items-center gap-2 mt-2 text-[11px] text-slate-600">
                          <span className="font-semibold">Chú giải ca:</span>
                          <span className="px-2 py-0.5 rounded bg-sky-100 text-sky-800 font-bold">N: Ca Ngày</span>
                          <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 font-bold">C: Ca Chiều</span>
                          <span className="px-2 py-0.5 rounded bg-indigo-100 text-indigo-800 font-bold">K: Ca Đêm</span>
                          <span className="px-2 py-0.5 rounded bg-slate-100 text-slate-700 font-bold">O: Nghỉ Ca</span>
                        </div>
                      </div>

                      {/* 2. Visual RULES Matrix */}
                      <div className="bg-white p-3.5 rounded-2xl border border-slate-200/80 shadow-sm">
                        <div className="flex items-center justify-between mb-2">
                          <div>
                            <h6 className="font-bold text-slate-800 text-xs flex items-center gap-1.5">
                              2. Quy Luật Phân Công Trực Thay Tự Động 
                            </h6>
                            <p className="text-[11px] text-slate-500">
                              Chọn Kíp trực thay tương ứng khi một Kíp bất kỳ xin nghỉ ở từng ca (N, C, K).
                            </p>
                          </div>
                          <button
                            type="button"
                            onClick={() => setRulesMatrixText(JSON.stringify(RULES, null, 2))}
                            className="text-[11px] font-semibold text-sky-700 hover:text-sky-900 bg-sky-50 hover:bg-sky-100 px-2.5 py-1 rounded-lg transition-all flex items-center gap-1 cursor-pointer"
                          >
                            <RefreshCw size={11} /> Khôi phục Mặc định
                          </button>
                        </div>

                        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-2.5">
                          {Array.from({ length: Math.max(getParsedShifts().length || 5, 5) }, (_, i) => i + 1).map((kipNum) => {
                            const rawRules = getParsedRules();
                            const kipRule = rawRules[kipNum] || rawRules[String(kipNum)] || {};
                            const availableTeams = Array.from({ length: Math.max(getParsedShifts().length || 5, 5) }, (_, i) => i + 1);

                            return (
                              <div key={kipNum} className="p-2.5 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
                                <div className="font-bold text-slate-900 text-xs pb-1 border-b border-slate-200 flex items-center justify-between">
                                  <span>Khi Kíp {kipNum} nghỉ:</span>
                                </div>

                                {/* Ca N */}
                                <div>
                                  <label className="block text-[11px] text-slate-600 mb-0.5">
                                    Nghỉ Ca <span className="font-bold text-sky-700">Ngày (N)</span>:
                                  </label>
                                  <select
                                    value={kipRule.N?.k || ''}
                                    onChange={(e) => handleRuleChange(kipNum, 'N', e.target.value ? Number(e.target.value) : 0)}
                                    className={`w-full px-2 py-1 bg-white border rounded-lg text-xs font-semibold focus:ring-1 focus:ring-sky-600 cursor-pointer ${
                                      !kipRule.N?.k ? 'border-rose-400 bg-rose-50 text-rose-700 font-bold' : 'border-slate-300 text-slate-800'
                                    }`}
                                  >
                                    <option value="">-- Chưa chọn kíp thay --</option>
                                    {availableTeams.map((k) => (
                                      <option key={k} value={k}>
                                        Kíp {k} thay
                                      </option>
                                    ))}
                                  </select>
                                </div>

                                {/* Ca C */}
                                <div>
                                  <label className="block text-[11px] text-slate-600 mb-0.5">
                                    Nghỉ Ca <span className="font-bold text-amber-700">Chiều (C)</span>:
                                  </label>
                                  <select
                                    value={kipRule.C?.k || ''}
                                    onChange={(e) => handleRuleChange(kipNum, 'C', e.target.value ? Number(e.target.value) : 0)}
                                    className={`w-full px-2 py-1 bg-white border rounded-lg text-xs font-semibold focus:ring-1 focus:ring-sky-600 cursor-pointer ${
                                      !kipRule.C?.k ? 'border-rose-400 bg-rose-50 text-rose-700 font-bold' : 'border-slate-300 text-slate-800'
                                    }`}
                                  >
                                    <option value="">-- Chưa chọn kíp thay --</option>
                                    {availableTeams.map((k) => (
                                      <option key={k} value={k}>
                                        Kíp {k} thay
                                      </option>
                                    ))}
                                  </select>
                                </div>

                                {/* Ca K */}
                                <div>
                                  <label className="block text-[11px] text-slate-600 mb-0.5">
                                    Nghỉ Ca <span className="font-bold text-indigo-700">Đêm (K)</span>:
                                  </label>
                                  <select
                                    value={kipRule.K?.k || ''}
                                    onChange={(e) => handleRuleChange(kipNum, 'K', e.target.value ? Number(e.target.value) : 0)}
                                    className={`w-full px-2 py-1 bg-white border rounded-lg text-xs font-semibold focus:ring-1 focus:ring-sky-600 cursor-pointer ${
                                      !kipRule.K?.k ? 'border-rose-400 bg-rose-50 text-rose-700 font-bold' : 'border-slate-300 text-slate-800'
                                    }`}
                                  >
                                    <option value="">-- Chưa chọn kíp thay --</option>
                                    {availableTeams.map((k) => (
                                      <option key={k} value={k}>
                                        Kíp {k} thay
                                      </option>
                                    ))}
                                  </select>
                                </div>
                              </div>
                            );
                          })}
                        </div>

                        {/* Rules Validation Banner */}
                        {(() => {
                          const currentRulesVal = validateRulesMatrix(getParsedRules(), getParsedShifts().length);
                          if (currentRulesVal.isValid) {
                            return (
                              <div className="mt-2.5 p-2.5 bg-sky-50 border border-sky-200 text-sky-800 rounded-xl text-xs font-semibold flex items-center gap-2">
                                <Check size={16} className="text-sky-700 shrink-0" />
                                <span>Quy luật phân công trực thay hợp lệ: Tất cả các ca nghỉ (N, C, K) của các Kíp đều đã chọn kíp đi thay đầy đủ.</span>
                              </div>
                            );
                          } else {
                            return (
                              <div className="mt-2.5 p-3 bg-rose-50 border border-rose-200 text-rose-800 rounded-xl text-xs font-medium space-y-1">
                                <div className="flex items-center gap-1.5 font-bold text-rose-900 text-xs">
                                  <AlertCircle size={16} className="text-rose-600 shrink-0" />
                                  <span>Cảnh báo lỗi nhập quy luật trực thay tự động:</span>
                                </div>
                                <p className="text-[11px] leading-relaxed font-semibold">{currentRulesVal.message}</p>
                              </div>
                            );
                          }
                        })()}
                      </div>
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pt-1">
                      <div>
                        <div className="flex items-center justify-between mb-1">
                          <label className="block font-semibold text-slate-800">
                            Ma trận Ca Trực (SHIFTS) - JSON Mảng 2 Chiều
                          </label>
                          <button
                            type="button"
                            onClick={() => setShiftsMatrixText(JSON.stringify(SHIFTS, null, 2))}
                            className="text-[11px] font-semibold text-sky-700 hover:text-sky-900 bg-sky-50 hover:bg-sky-100 px-2 py-0.5 rounded-lg transition-all flex items-center gap-1 cursor-pointer"
                          >
                            <RefreshCw size={11} /> Mặc định
                          </button>
                        </div>
                        <textarea
                          value={shiftsMatrixText}
                          onChange={(e) => setShiftsMatrixText(e.target.value)}
                          rows={6}
                          className="w-full font-mono text-xs p-2.5 bg-slate-900 text-amber-300 rounded-xl border border-slate-700 focus:outline-none focus:ring-2 focus:ring-sky-600"
                        />
                        <p className="text-[11px] text-slate-500 mt-0.5">
                          Ma trận 5 hàng (tương ứng Kíp 1 đến Kíp 5) x 5 ngày chu kỳ. Giá trị: "N" (Ngày), "C" (Chiều), "K" (Đêm), "O" (Nghỉ).
                        </p>
                      </div>

                      <div>
                        <div className="flex items-center justify-between mb-1">
                          <label className="block font-semibold text-slate-800">
                            Bảng Quy Luật Kíp Trực Thay (RULES) - JSON Object
                          </label>
                          <button
                            type="button"
                            onClick={() => setRulesMatrixText(JSON.stringify(RULES, null, 2))}
                            className="text-[11px] font-semibold text-sky-700 hover:text-sky-900 bg-sky-50 hover:bg-sky-100 px-2 py-0.5 rounded-lg transition-all flex items-center gap-1 cursor-pointer"
                          >
                            <RefreshCw size={11} /> Mặc định
                          </button>
                        </div>
                        <textarea
                          value={rulesMatrixText}
                          onChange={(e) => setRulesMatrixText(e.target.value)}
                          rows={6}
                          className="w-full font-mono text-xs p-2.5 bg-slate-900 text-sky-300 rounded-xl border border-slate-700 focus:outline-none focus:ring-2 focus:ring-sky-600"
                        />
                        <p className="text-[11px] text-slate-500 mt-0.5">
                          Xác định kíp nào sẽ ưu tiên trực thay khi Kíp nghỉ xin phép ở từng ca N, C, K.
                        </p>
                      </div>
                    </div>
                  )}
                </div>
              </div>

            

              {/* 4. Leader Signature & Webhook Config */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3">
                <h5 className="font-bold text-slate-900 text-xs uppercase tracking-wider text-sky-800">
                  3. Người Ký Văn bản & Zalo Webhook
                </h5>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Họ và tên Người ký văn bản</label>
                    <input
                      type="text"
                      value={wsNguoiKy}
                      onChange={(e) => setWsNguoiKy(e.target.value)}
                      placeholder="ví dụ: Nguyễn Văn Nghị"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    />
                  </div>

                  <div>
                    <label className="block font-medium text-slate-700 mb-1">Chức vụ Người ký</label>
                    <input
                      type="text"
                      value={wsChucVu}
                      onChange={(e) => setWsChucVu(e.target.value)}
                      placeholder="Quản đốc Phân xưởng"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                    />
                  </div>

                </div>

                <div>
                  <label className="block font-medium text-slate-700 mb-1">Zalo Webhook URL Nhận Thông Báo cho PX này</label>
                  <input
                    type="text"
                    value={wsWebhookUrl}
                    onChange={(e) => setWsWebhookUrl(e.target.value)}
                    placeholder="https://vhialy.dpdns.org/webhook/notify"
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-mono focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                  />
                </div>

                <div>
                  <label className="block font-medium text-slate-700 mb-1">Email Nhận Thông Báo cho PX này</label>
                  <input
                    type="text"
                    value={wsNotifyEmail}
                    onChange={(e) => setWsNotifyEmail(e.target.value)}
                    placeholder="vd: quandoc@gmail.com (nhiều địa chỉ cách nhau bằng dấu phẩy)"
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-mono focus:outline-none focus:ring-2 focus:ring-sky-600/20"
                  />
                  <span className="text-[11px] text-slate-500 italic mt-1 block">
                    * Mỗi khi có đơn nghỉ phép mới, hệ thống gửi thông báo qua cả Zalo webhook và email này (nếu có điền).
                  </span>
                </div>

              </div>



              {/* 6. Staff Data Visual & Interactive Editor */}
              <StaffDataEditor
                staffDataText={staffDataText}
                onChangeStaffDataText={setStaffDataText}
                teamsListText={teamsListText}
              />

              {/* Form Buttons */}
              <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-200">
                <button
                  type="button"
                  onClick={() => setActiveTab('workshops')}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold rounded-xl transition-all cursor-pointer"
                >
                  Hủy bỏ
                </button>
                <button
                  type="submit"
                  disabled={isSavingWs}
                  className="px-5 py-2 bg-sky-700 hover:bg-sky-700 text-white font-semibold rounded-xl shadow-md transition-all cursor-pointer flex items-center gap-1.5"
                >
                  <Save size={16} />
                  <span>{isSavingWs ? 'Đang lưu...' : 'Lưu Cấu hình Phân xưởng'}</span>
                </button>
              </div>
            </form>
          )}
        </div>
      </div>

      {/* CONFIRM DELETE WORKSHOP DIALOG */}
      {confirmDeleteWorkshop && (
        <div className="fixed inset-0 z-70 flex items-center justify-center bg-slate-900/70 backdrop-blur-xs p-4 animate-in fade-in duration-150">
          <div className="bg-white w-full max-w-md rounded-2xl shadow-2xl border border-rose-100 p-6 space-y-4 text-center">
            <div className="w-12 h-12 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center mx-auto">
              <Trash2 size={24} />
            </div>
            <div>
              <h3 className="font-bold text-slate-900 text-lg">Xác nhận xóa phân xưởng</h3>
              <p className="text-sm text-slate-600 mt-2 leading-relaxed">
                Bạn có chắc chắn muốn xóa phân xưởng <span className="font-semibold text-rose-600">"{confirmDeleteWorkshop.name}"</span>?
                Tất cả tài khoản thuộc phân xưởng này cũng sẽ bị xóa vĩnh viễn khỏi hệ thống.
              </p>
            </div>
            <div className="flex items-center justify-center gap-3 pt-2">
              <button
                type="button"
                onClick={() => setConfirmDeleteWorkshop(null)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl transition-all cursor-pointer"
              >
                Hủy bỏ
              </button>
              <button
                type="button"
                onClick={executeDeleteWorkshop}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold rounded-xl transition-all cursor-pointer shadow-md"
              >
                Đồng ý Xóa
              </button>
            </div>
          </div>
        </div>
      )}

      {/* CONFIRM DELETE ACCOUNT DIALOG */}
      {confirmDeleteAccount && (
        <div className="fixed inset-0 z-70 flex items-center justify-center bg-slate-900/70 backdrop-blur-xs p-4 animate-in fade-in duration-150">
          <div className="bg-white w-full max-w-md rounded-2xl shadow-2xl border border-rose-100 p-6 space-y-4 text-center">
            <div className="w-12 h-12 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center mx-auto">
              <Trash2 size={24} />
            </div>
            <div>
              <h3 className="font-bold text-slate-900 text-lg">Xác nhận xóa tài khoản</h3>
              <p className="text-sm text-slate-600 mt-2 leading-relaxed">
                Bạn có chắc chắn muốn xóa tài khoản <span className="font-semibold text-rose-600">"{confirmDeleteAccount.username}"</span>?
                Hành động này không thể hoàn tác.
              </p>
            </div>
            <div className="flex items-center justify-center gap-3 pt-2">
              <button
                type="button"
                onClick={() => setConfirmDeleteAccount(null)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl transition-all cursor-pointer"
              >
                Hủy bỏ
              </button>
              <button
                type="button"
                onClick={executeDeleteAccount}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold rounded-xl transition-all cursor-pointer shadow-md"
              >
                Đồng ý Xóa
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}