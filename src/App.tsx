import React, { useState, useEffect, useCallback, useRef } from 'react';
import mammoth from 'mammoth';
import { DEFAULT_STAFF, SHIFTS } from './constants';
import { fmtVN, fmtIn, dayN, timNghi, timThay, abbrev, xacDinhCa } from './utils/shiftHelpers';
import { buildMultiLeaveResults, Leave, ResultItem } from './utils/Quytacxacdinhcatructhay';
import { exportWord, generateWordBlob, exportSwapDoc, generateSwapBlob, exportLeaveRequestDoc, generateLeaveRequestBlob, exportAllDocsZip } from './utils/wordExport';
import { renderAsync } from 'docx-preview';
import SignatureManager from './components/SignatureManager';
import { Trash2 } from 'lucide-react';

export default function App() {
  const [staffData, setStaffData] = useState<string[][]>(() => {
    const saved = localStorage.getItem('sd');
    const ver = localStorage.getItem('sv');
    if (ver !== 'v4') {
      localStorage.removeItem('sd');
      localStorage.setItem('sv', 'v4');
      return DEFAULT_STAFF;
    }
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          let hasMigration = false;
          const migrated = parsed.map(row => {
            if (Array.isArray(row) && row[0] === 'Trực phụ cơ MR') {
              hasMigration = true;
              return ['Trực phụ máy MR', ...row.slice(1)];
            }
            return row;
          });
          if (hasMigration) {
            localStorage.setItem('sd', JSON.stringify(migrated));
          }
          return migrated;
        }
      } catch (e) {
        console.error("Failed to parse local staffData", e);
      }
    }
    return DEFAULT_STAFF;
  });

  const [ngayBatDau, setNgayBatDau] = useState(() => fmtIn(new Date()));
  const [ngayKetThuc, setNgayKetThuc] = useState(() => {
    const next = new Date();
    next.setDate(next.getDate() + 7);
    return fmtIn(next);
  });
  const [chucDanh, setChucDanh] = useState('');
  const [kipNghi, setKipNghi] = useState('');
  const [additionalLeaves, setAdditionalLeaves] = useState<{ kip: string, start: string, end: string, chucDanh: string }[]>([]);
  const [showStaff, setShowStaff] = useState(true);
  const [alert, setAlert] = useState<string | null>(null);
  const [currentResult, setCurrentResult] = useState<any>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [previewBlob, setPreviewBlob] = useState<Blob | null>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [config, setConfig] = useState({
    soVanBan: '',
    ngayKy: '',
    nguoiKy: 'Nguyễn Văn Nghị',
    zaloWebhookUrl: 'https://vhialy.dpdns.org/webhook/notify'
  });

  const [isGoogleAuth, setIsGoogleAuth] = useState(false);
  const [loadingAuth, setLoadingAuth] = useState(true);
  const [isUpdatingSheets, setIsUpdatingSheets] = useState(false);
  const [signatures, setSignatures] = useState<Record<string, string>>({});
  const [showSignatureManager, setShowSignatureManager] = useState(false);
  const [isSettingsLoaded, setIsSettingsLoaded] = useState(false);

  // Manual Swap State
  const [swapData, setSwapData] = useState({
    date1: fmtIn(new Date()),
    date2: fmtIn(new Date()),
    person1: '',
    person2: '',
    shift1: 'N',
    shift2: 'K'
  });
  const [isPreviewingSwap, setIsPreviewingSwap] = useState(false);
  const [swapChucDanh, setSwapChucDanh] = useState(staffData[0]?.[0] || '');
  const [leaveData, setLeaveData] = useState({
    name: '',
    birthYear: '1980',
    chucDanh: staffData[0]?.[0] || '',
    kip: '1',
    startDate: fmtIn(new Date()),
    endDate: fmtIn(new Date()),
    reason: 'Giải quyết việc riêng gia đình',
    phone: '',
    leaveYear: String(new Date().getFullYear()),
    location: 'Gia Lai'
  });
  const [isPreviewingLeave, setIsPreviewingLeave] = useState(false);

  // States for Staff Leave Balance from Google Sheets
  const [leaveBalance, setLeaveBalance] = useState<{ entitled: string; used: string; remaining: string } | null>(null);
  const [isLoadingLeaveBalance, setIsLoadingLeaveBalance] = useState(false);
  const [leaveBalanceError, setLeaveBalanceError] = useState<string | null>(null);

  const fetchLeaveBalance = useCallback(async (name: string) => {
    if (!name || name.trim() === '') {
      setLeaveBalance(null);
      setLeaveBalanceError(null);
      return;
    }
    setIsLoadingLeaveBalance(true);
    setLeaveBalanceError(null);
    try {
      const res = await fetch(`/api/sheets/leave-balance?name=${encodeURIComponent(name.trim())}`);
      const data = await res.json();
      if (res.ok && data.success) {
        setLeaveBalance({
          entitled: data.entitled,
          used: data.used,
          remaining: data.remaining
        });
      } else {
        setLeaveBalance(null);
        setLeaveBalanceError(data.error || "Không tìm thấy thông tin phép năm");
      }
    } catch (err: any) {
      setLeaveBalance(null);
      setLeaveBalanceError("Không thể kết nối máy chủ để lấy thông tin phép năm");
    } finally {
      setIsLoadingLeaveBalance(false);
    }
  }, []);

  useEffect(() => {
    const nameStr = (leaveData.name || '').trim();
    if (!nameStr) {
      setLeaveBalance(null);
      setLeaveBalanceError(null);
      return;
    }
    const timer = setTimeout(() => {
      fetchLeaveBalance(nameStr);
    }, 600);
    return () => clearTimeout(timer);
  }, [leaveData.name, fetchLeaveBalance]);

  // Google Sheets Leave Queues States
  const [waitingLeaves, setWaitingLeaves] = useState<any[]>([]);
  const [isLoadingWaitingLeaves, setIsLoadingWaitingLeaves] = useState(false);
  const [selectedWaitingLeaveIds, setSelectedWaitingLeaveIds] = useState<string[]>([]);
  const [isSavingLeaveToSheets, setIsSavingLeaveToSheets] = useState(false);
  const [confirmModal, setConfirmModal] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    onConfirm: () => void;
    onCancel?: () => void;
    confirmText?: string;
    cancelText?: string;
    isDanger?: boolean;
    requirePassword?: boolean;
  } | null>(null);
  const [confirmPassword, setConfirmPassword] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [activeTab, setActiveTab] = useState<'schedule' | 'leave' | 'swap' | 'staff'>('schedule');

  const closeConfirmModal = () => {
    setConfirmModal(null);
    setConfirmPassword("");
    setPasswordError("");
  };

  const handleConfirmSubmit = () => {
    if (!confirmModal) return;
    if (confirmModal.requirePassword) {
      if (confirmPassword !== '1234') {
        setPasswordError("Mật khẩu không chính xác! .");
        return;
      }
    }
    confirmModal.onConfirm();
    closeConfirmModal();
  };

  useEffect(() => {
    const checkAuth = async () => {
      setLoadingAuth(true);
      try {
        const res = await fetch('/api/auth/status');
        const data = await res.json();
        setIsGoogleAuth(data.authenticated);
      } catch (e) {
        console.error("Auth check failed", e);
      } finally {
        setLoadingAuth(false);
      }
    };
    checkAuth();

    const handleMessage = (e: MessageEvent) => {
      if (e.data?.type === 'OAUTH_AUTH_SUCCESS') {
        setIsGoogleAuth(true);
        setAlert('✅ Kết nối Google thành công!');
      }
    };
    window.addEventListener('message', handleMessage);

    // Fetch cloud settings on mount
    const fetchSettings = async () => {
      try {
        // Fetch app settings
        const res = await fetch('/api/app-settings');
        const data = await res.json();
        if (data.staffData && Array.isArray(data.staffData)) {
          const migrated = data.staffData.map((row: any) => {
            if (Array.isArray(row) && row[0] === 'Trực phụ cơ MR') {
              return ['Trực phụ máy MR', ...row.slice(1)];
            }
            return row;
          });
          setStaffData(migrated);
          localStorage.setItem('sd', JSON.stringify(migrated));
        }
        if (data.config) {
          setConfig({
            soVanBan: data.config.soVanBan || '',
            ngayKy: data.config.ngayKy || '',
            nguoiKy: data.config.nguoiKy || 'Nguyễn Văn Nghị',
            zaloWebhookUrl: data.config.zaloWebhookUrl || 'https://vhialy.dpdns.org/webhook/notify'
          });
        }

        // Fetch signatures
        const sigRes = await fetch('/api/signatures');
        const sigData = await sigRes.json();
        setSignatures(sigData);

        setIsSettingsLoaded(true);
      } catch (e) {
        console.error("Failed to fetch app settings", e);
        setIsSettingsLoaded(true);
      }
    };
    fetchSettings();

    return () => window.removeEventListener('message', handleMessage);
  }, []);

  // Save settings to cloud whenever they change (debounced)
  useEffect(() => {
    if (!isSettingsLoaded) return;
    
    const saveSettings = async () => {
      try {
        await fetch('/api/app-settings', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ staffData, config })
        });
      } catch (e) {
        console.error("Failed to save app settings", e);
      }
    };
    
    const timer = setTimeout(saveSettings, 3000);
    return () => clearTimeout(timer);
  }, [staffData, config, isSettingsLoaded]);

  const fetchWaitingLeaves = useCallback(async () => {
    setIsLoadingWaitingLeaves(true);
    try {
      const res = await fetch('/api/sheets/leave-requests');
      if (res.ok) {
        const data = await res.json();
        const mappedData = data.map((item: any) => ({
          ...item,
          chucDanh: item.chucDanh === 'Trực phụ cơ MR' ? 'Trực phụ máy MR' : item.chucDanh
        }));
        const waiting = mappedData.filter((item: any) => item.status === 'Chờ phân ca');
        setWaitingLeaves(waiting);
        // Clear selected if they no longer exist in waiting
        setSelectedWaitingLeaveIds(prev => prev.filter(id => waiting.some((w: any) => w.id === id)));
      } else {
        if (res.status === 401) {
          setIsGoogleAuth(false);
        }
        console.error("Failed to fetch leave requests");
      }
    } catch (e) {
      console.error("Error fetching live leaves request", e);
    } finally {
      setIsLoadingWaitingLeaves(false);
    }
  }, []);

  useEffect(() => {
    fetchWaitingLeaves();
  }, [fetchWaitingLeaves]);

  const handleConnectGoogle = () => {
    const width = 500, height = 600;
    const left = (window.innerWidth - width) / 2;
    const top = (window.innerHeight - height) / 2;
    window.open('/api/auth/google', 'google-auth', `width=${width},height=${height},left=${left},top=${top}`);
  };

  const getSheetsUpdates = (resObj: any) => {
    if (!resObj) return [];
    const updateMap: Record<string, string> = {}; // key: "name|date"

    // 1. Process all leaves (Absent people)
    resObj.allResults.forEach((res: any) => {
      const start = new Date(res.start);
      const end = new Date(res.end);
      let d = new Date(start);
      while (d <= end) {
        const dateStr = fmtIn(d);
        const originalShift = xacDinhCa(d, res.kip);
        // Only change N, C, K to F. Keep O as O.
        const finalShift = (originalShift === 'N' || originalShift === 'C' || originalShift === 'K') ? 'F' : originalShift;
        
        if (!res.ten.includes('THIẾU NHÂN SỰ')) {
          const name = res.ten.trim().normalize('NFC');
          updateMap[`${name}|${dateStr}`] = finalShift;
        }
        d.setDate(d.getDate() + 1);
      }

      // 2. Process specific assignments for each leave
      res.ketQua.forEach((item: any) => {
        const dateStr = fmtIn(new Date(item.ngay));
        
        // Replacement person works the shift
        if (item.nguoitructhay && item.nguoitructhay !== 'N/A' && !item.nguoitructhay.includes('THIẾU NHÂN SỰ')) {
          const name = item.nguoitructhay.trim().normalize('NFC');
          updateMap[`${name}|${dateStr}`] = item.ca;
        }

        // If it's a CK Swap, the absent person works the other shift
        if (item.isCKSwap && item.swapAbsentTen && !item.swapAbsentTen.includes('THIẾU NHÂN SỰ')) {
          const name = item.swapAbsentTen.trim().normalize('NFC');
          const absentShift = xacDinhCa(new Date(item.ngay), item.kiptructhay);
          updateMap[`${name}|${dateStr}`] = absentShift;
        }

        // If someone is relieved
        if (item.relievedTen && !item.relievedTen.includes('THIẾU NHÂN SỰ')) {
          const name = item.relievedTen.trim().normalize('NFC');
          updateMap[`${name}|${dateStr}`] = 'O';
        }
      });
    });

    // 3. Process extraRows (Additional adjustments)
    resObj.extraRows.forEach((row: any) => {
      const dateStr = fmtIn(new Date(row.ngay));
      
      if (row.nguoitructhay && row.nguoitructhay !== 'N/A' && !row.nguoitructhay.includes('THIẾU NHÂN SỰ')) {
        const name = row.nguoitructhay.trim().normalize('NFC');
        updateMap[`${name}|${dateStr}`] = row.ca;
      }
      
      if (row.relievedTen && !row.relievedTen.includes('THIẾU NHÂN SỰ')) {
        const name = row.relievedTen.trim().normalize('NFC');
        updateMap[`${name}|${dateStr}`] = 'O';
      }

      if (row.absentTen && !row.absentTen.includes('THIẾU NHÂN SỰ')) {
        const name = row.absentTen.trim().normalize('NFC');
        // Only set to F if not already assigned a working shift (like in a swap)
        const currentVal = updateMap[`${name}|${dateStr}`];
        if (!currentVal || currentVal === 'F') {
          const originalShift = xacDinhCa(new Date(row.ngay), row.absentKip);
          const finalShift = (originalShift === 'N' || originalShift === 'C' || originalShift === 'K') ? 'F' : originalShift;
          updateMap[`${name}|${dateStr}`] = finalShift;
        }
      }
    });

    return Object.entries(updateMap).map(([key, shift]) => {
      const [name, date] = key.split('|');
      return { name, date, shift };
    });
  };

  const updateGoogleSheets = async () => {
    if (!isGoogleAuth || !currentResult) return;
    setIsUpdatingSheets(true);
    try {
      const updates = getSheetsUpdates(currentResult);

      console.log("Sending updates to Google Sheets:", updates);

      const res = await fetch('/api/sheets/update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          spreadsheetId: '1HgGW-FvoGQXtj7V_JCMD-7Tuue0rTIM-bmohGmgqm6I',
          updates
        })
      });
      const data = await res.json();
      if (data.success) {
        setAlert('✅ Đã cập nhật vào bảng theo dõi cơm ca');
      } else {
        if (res.status === 401) {
          setIsGoogleAuth(false);
        }
        setAlert('⚠ Lỗi cập nhật Google Sheets: ' + (data.error || 'Unknown error'));
      }
    } catch (e) {
      console.error(e);
      setAlert('⚠ Lỗi kết nối Google Sheets');
    } finally {
      setIsUpdatingSheets(false);
    }
  };

  const handleUpdateStaff = (r: number, c: number, val: string) => {
    const newData = [...staffData];
    newData[r][c] = val.trim();
    setStaffData(newData);
    localStorage.setItem('sd', JSON.stringify(newData));
  };

  const addConcurrentLeave = () => {
    if (!chucDanh) {
      setAlert('⚠ Vui lòng chọn chức danh trước!');
      return;
    }
    setAdditionalLeaves([...additionalLeaves, { kip: '', start: '', end: '', chucDanh: chucDanh || (staffData[0] ? staffData[0][0] : '') }]);
  };

  const removeConcurrentLeave = (idx: number) => {
    const newList = [...additionalLeaves];
    newList.splice(idx, 1);
    setAdditionalLeaves(newList);
  };

  const updateConcurrentLeave = (idx: number, field: string, val: string) => {
    const newList = [...additionalLeaves];
    (newList[idx] as any)[field] = val;
    setAdditionalLeaves(newList);
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setIsParsing(true);
    setAlert(null);
    
    let mainSet = !!(chucDanh && kipNghi);
    const newAdditionalLeaves = [...additionalLeaves];
    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      try {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        const text = result.value;

        // Extraction logic
        const nameMatch = text.match(/Tên tôi là:\s*(.*)/i);
        const positionMatch = text.match(/Chức vụ:\s*(.*)/i);
        const dateMatch = text.match(/Thời gian:\s*Từ ngày\s*(\d{2}\/\d{2}\/\d{4})\s*đến hết ngày\s*(\d{2}\/\d{2}\/\d{4})/i);

        if (!nameMatch || !positionMatch || !dateMatch) {
          errorCount++;
          continue;
        }

        const extractedName = nameMatch[1].trim();
        const extractedPosition = positionMatch[1].trim();
        const startDateStr = dateMatch[1].trim();
        const endDateStr = dateMatch[2].trim();

        // Parse dates (DD/MM/YYYY to YYYY-MM-DD)
        const parseDate = (dStr: string) => {
          const [d, m, y] = dStr.split('/');
          return `${y}-${m}-${d}`;
        };

        const startISO = parseDate(startDateStr);
        const endISO = parseDate(endDateStr);

        // Try to find the person in staffData to get the correct title and kip
        let foundTitle = '';
        let foundKip = '';

        for (let r = 0; r < staffData.length; r++) {
          for (let c = 1; c <= 5; c++) {
            if (staffData[r][c]?.toLowerCase() === extractedName.toLowerCase()) {
              foundTitle = staffData[r][0];
              foundKip = String(c);
              break;
            }
          }
          if (foundTitle) break;
        }

        // If name not found, try to extract kip from position string (e.g., "Kíp2")
        if (!foundKip) {
          const kipMatch = extractedPosition.match(/Kíp\s*(\d)/i);
          if (kipMatch) foundKip = kipMatch[1];
        }

        // If title not found, try to match extractedPosition with staffData titles
        if (!foundTitle) {
          const sortedTitles = [...staffData].sort((a, b) => b[0].length - a[0].length);
          const matchedTitle = sortedTitles.find(r => extractedPosition.toLowerCase().includes(r[0].toLowerCase()));
          if (matchedTitle) foundTitle = matchedTitle[0];
        }

        if (!mainSet) {
          setNgayBatDau(startISO);
          setNgayKetThuc(endISO);
          setChucDanh(foundTitle);
          setKipNghi(foundKip);
          mainSet = true;
          successCount++;
        } else {
          newAdditionalLeaves.push({
            kip: foundKip,
            start: startISO,
            end: endISO,
            chucDanh: foundTitle
          });
          successCount++;
        }
      } catch (err) {
        console.error(err);
        errorCount++;
      }
    }

    setAdditionalLeaves(newAdditionalLeaves);

    if (successCount > 0) {
      setAlert(`✅ Đã trích xuất thành công ${successCount} đơn nghỉ phép.${errorCount > 0 ? ` (Thất bại ${errorCount} file)` : ''}`);
    } else if (errorCount > 0) {
      setAlert(`⚠ Không thể trích xuất thông tin từ ${errorCount} file. Vui lòng kiểm tra định dạng.`);
    }

    setIsParsing(false);
    e.target.value = '';
  };

  const taoLich = () => {
    setAlert(null);
    if (!ngayBatDau || !ngayKetThuc || !chucDanh || !kipNghi) {
      setAlert('⚠ Vui lòng nhập đầy đủ thông tin!');
      return;
    }

    const start = new Date(ngayBatDau + 'T00:00:00');
    const end = new Date(ngayKetThuc + 'T00:00:00');
    const kip = +kipNghi;

    if (start > end) {
      setAlert('⚠ Ngày bắt đầu phải trước ngày kết thúc!');
      return;
    }

    const ten = timNghi(chucDanh, kip, staffData);
    if (!ten) {
      setAlert(`⚠ Không tìm thấy chức danh "${chucDanh}" trong Kíp ${kip}!`);
      return;
    }

    const allLeaves: Leave[] = [{ kip, start, end, ten, chucDanh }];
    const addErr: string[] = [];
    additionalLeaves.forEach((al, idx) => {
      if (!al.kip && !al.start && !al.end) return;
      if (!al.kip || !al.start || !al.end || !al.chucDanh) {
        addErr.push(`Người nghỉ #${idx + 2} thiếu thông tin`);
        return;
      }
      const alKip = +al.kip;
      const alChucDanh = al.chucDanh;

      // Check for duplicate kip within the same chucDanh
      if (allLeaves.some(l => l.kip === alKip && l.chucDanh === alChucDanh)) {
        addErr.push(`Kíp ${alKip} của chức danh ${alChucDanh} đã được thêm`);
        return;
      }

      const alStart = new Date(al.start + 'T00:00:00');
      const alEnd = new Date(al.end + 'T00:00:00');
      if (alStart > alEnd) {
        addErr.push(`Người nghỉ #${idx + 2} ngày bắt đầu sau ngày kết thúc`);
        return;
      }
      const alTen = timNghi(alChucDanh, alKip, staffData);
      if (!alTen) {
        addErr.push(`Không tìm thấy "${alChucDanh}" trong Kíp ${alKip}`);
        return;
      }
      allLeaves.push({ kip: alKip, start: alStart, end: alEnd, ten: alTen, chucDanh: alChucDanh });
    });

    if (addErr.length) {
      setAlert('⚠ ' + addErr.join(' | '));
      return;
    }

    setIsProcessing(true);
    setTimeout(() => {
      // Tự động phát hiện và bổ sung các vị trí thiếu nhân sự (ô trống trong danh sách)
      // Nếu một chức danh có người nghỉ phép, các kíp đang thiếu người ở chức danh đó cũng sẽ được xếp lịch thay
      const minStart = new Date(Math.min(...allLeaves.map(l => l.start.getTime())));
      const maxEnd = new Date(Math.max(...allLeaves.map(l => l.end.getTime())));
      const involvedTitles = Array.from(new Set(allLeaves.map(l => l.chucDanh)));
      
      involvedTitles.forEach(title => {
        const row = staffData.find(r => r[0] === title);
        if (row) {
          for (let k = 1; k <= 5; k++) {
            const staffName = row[k] ? row[k].trim() : '';
            if (staffName === '') {
              if (!allLeaves.some(l => l.kip === k && l.chucDanh === title)) {
                allLeaves.push({
                  kip: k,
                  start: minStart,
                  end: maxEnd,
                  ten: `THIẾU NHÂN SỰ (Kíp ${k})`,
                  chucDanh: title
                });
              }
            }
          }
        }
      });

      // Group leaves by job title
      const groups = allLeaves.reduce((acc, l) => {
        if (!acc[l.chucDanh]) acc[l.chucDanh] = [];
        acc[l.chucDanh].push(l);
        return acc;
      }, {} as Record<string, Leave[]>);

      let mergedResults: any[] = [];
      let mergedExtraRows: any[] = [];
      let mergedHasConflict = false;
      const mergedCoverCount: Record<number, number> = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };

      Object.keys(groups).forEach(cd => {
        const buildResult = buildMultiLeaveResults(groups[cd], cd, staffData);
        mergedResults = [...mergedResults, ...buildResult.results];
        mergedExtraRows = [...mergedExtraRows, ...buildResult.extraRows];
        if (buildResult.hasConflict) mergedHasConflict = true;
        for (let k = 1; k <= 5; k++) {
          mergedCoverCount[k] += (buildResult.coverCount[k] || 0);
        }
      });

      setCurrentResult({
        ten, chucDanh, kip, start, end,
        ketQua: mergedResults[0].ketQua,
        allResults: mergedResults,
        extraRows: mergedExtraRows,
        hasConflict: mergedHasConflict,
        coverCount: mergedCoverCount,
        isMulti: allLeaves.length > 1
      });
      setIsProcessing(false);
    }, 250);
  };

  const handleExportWord = async () => {
    setIsProcessing(true);
    await exportWord(currentResult, config);
    if (isGoogleAuth) {
      await updateGoogleSheets();
    }
    setIsProcessing(false);
  };

  useEffect(() => {
    if (showPreview) {
      document.body.style.overflow = 'hidden';
    } else {
      document.body.style.overflow = 'unset';
    }
    return () => {
      document.body.style.overflow = 'unset';
    };
  }, [showPreview]);

  const handlePreviewWord = async () => {
    setIsProcessing(true);
    const blob = await generateWordBlob(currentResult, config);
    if (blob) {
      setPreviewBlob(blob);
      setIsPreviewingSwap(false);
      setIsPreviewingLeave(false);
      setShowPreview(true);
    }
    setIsProcessing(false);
  };

  const handlePreviewSwap = async () => {
    setIsProcessing(true);
    const blob = await generateSwapBlob(swapData, config, signatures);
    if (blob) {
      setPreviewBlob(blob);
      setIsPreviewingSwap(true);
      setIsPreviewingLeave(false);
      setShowPreview(true);
    }
    setIsProcessing(false);
  };

  const handlePreviewLeave = async () => {
    setIsProcessing(true);
    const blob = await generateLeaveRequestBlob(leaveData, config, signatures);
    if (blob) {
      setPreviewBlob(blob);
      setIsPreviewingSwap(false);
      setIsPreviewingLeave(true);
      setShowPreview(true);
    }
    setIsProcessing(false);
  };

  const handleExportLeave = async () => {
    setIsProcessing(true);
    await exportLeaveRequestDoc(leaveData, config, signatures);
    setIsProcessing(false);
  };

  const saveLeaveToGoogleSheets = async () => {
    if (!leaveData.name) {
      setAlert("⚠ Vui lòng nhập Họ và tên người xin nghỉ.");
      return;
    }
    setIsSavingLeaveToSheets(true);
    try {
      const payload = {
        ...leaveData,
        leaveBalance: leaveBalance ? {
          entitled: leaveBalance.entitled,
          used: leaveBalance.used,
          remaining: leaveBalance.remaining
        } : null
      };
      const res = await fetch('/api/sheets/leave-requests', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const resData = await res.json();
      if (res.ok) {
        setAlert(resData.message || `✅ Đã lưu đơn của đồng chí ${leaveData.name} lên hệ thống ở trạng thái Chờ phân ca!`);
        fetchWaitingLeaves();
        setShowPreview(false);
      } else {
        if (res.status === 401 || resData.auth_expired) {
          setIsGoogleAuth(false);
        }
        setAlert(`❌ Lưu đơn thất bại: ${resData.error || "Lỗi máy chủ"}`);
      }
    } catch (e: any) {
      setAlert(`❌ Đã xảy ra lỗi kết nối: ${e.message}`);
    } finally {
      setIsSavingLeaveToSheets(false);
    }
  };

  const applyWaitingLeavesToForm = () => {
    if (selectedWaitingLeaveIds.length === 0) {
      setAlert("⚠ Vui lòng chọn ít nhất một đơn xin nghỉ phép trong danh sách chờ.");
      return;
    }
    
    const selectedItems = waitingLeaves.filter(item => selectedWaitingLeaveIds.includes(item.id));
    if (selectedItems.length === 0) return;

    // Đơn đầu tiên cho vào Người nghỉ chính
    const first = selectedItems[0];
    setNgayBatDau(first.startDate);
    setNgayKetThuc(first.endDate);
    setChucDanh(first.chucDanh);
    setKipNghi(String(first.kip));

    // Các đơn còn lại cho vào concurrent / additionalLeaves
    const rem = selectedItems.slice(1);
    const newAdditional = rem.map(item => ({
      chucDanh: item.chucDanh,
      kip: String(item.kip),
      start: item.startDate,
      end: item.endDate
    }));
    setAdditionalLeaves(newAdditional);

    setAlert(`✅ Đã áp dụng ${selectedItems.length} đơn nghỉ phép vào form lập lịch. Nhấn "Tạo Lịch Thay Ca" để tính toán.`);
  };

  const handleDeleteWaitingLeaves = (ids: string[], isSingleName?: string) => {
    if (!isGoogleAuth) return;
    if (ids.length === 0) return;
    
    const message = isSingleName 
      ? `Bạn có chắc chắn muốn xóa đơn nghỉ phép của anh/chị "${isSingleName}" khỏi Google Sheets không?`
      : `Bạn có chắc chắn muốn danh sách ${ids.length} đơn nghỉ phép đã chọn bị xóa vĩnh viễn khỏi Google Sheets?`;

    setConfirmModal({
      isOpen: true,
      title: "Xác nhận xóa đơn nghỉ phép",
      message: message,
      confirmText: "Đồng ý Xóa",
      cancelText: "Hủy bỏ",
      isDanger: true,
      requirePassword: true,
      onConfirm: async () => {
        setIsLoadingWaitingLeaves(true);
        try {
          const res = await fetch('/api/sheets/leave-requests/delete', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({ ids })
          });

          if (res.ok) {
            setAlert(isSingleName 
              ? `✅ Đã xóa đơn của "${isSingleName}" thành công!`
              : `✅ Đã xóa thành công ${ids.length} đơn nghỉ phép!`
            );
            setSelectedWaitingLeaveIds(prev => prev.filter(id => !ids.includes(id)));
            fetchWaitingLeaves();
          } else if (res.status === 401) {
            setIsGoogleAuth(false);
            setAlert("⚠ Phiên kết nối Google Sheets đã hết hạn. Vui lòng kết nối lại tài khoản.");
          } else {
            const data = await res.json();
            setAlert(`❌ Xóa đơn thất bại: ${data.error || "Lỗi máy chủ"}`);
          }
        } catch (err: any) {
          setAlert(`❌ Lỗi kết nối xóa đơn nghỉ: ${err.message}`);
        } finally {
          setIsLoadingWaitingLeaves(false);
        }
      }
    });
  };

  const proceedExportAllZipAndUpdateStatus = async (shouldUpdateAnnualLeave: boolean) => {
    if (!currentResult) return;
    setIsProcessing(true);
    try {
      const selectedLeavesData = waitingLeaves.filter(item => selectedWaitingLeaveIds.includes(item.id));
      
      // 1. Download the ZIP file
      await exportAllDocsZip(currentResult, selectedLeavesData, config, signatures);
      
      let isSheetSynced = false;
      let sheetSyncError = "";

      // 2. Clear/Synchronize schedule & meal reports if authenticated
      if (isGoogleAuth) {
        try {
          const updates = getSheetsUpdates(currentResult);
          if (updates.length > 0) {
            const syncRes = await fetch('/api/sheets/update', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                spreadsheetId: '1HgGW-FvoGQXtj7V_JCMD-7Tuue0rTIM-bmohGmgqm6I',
                updates
              })
            });
            const syncData = await syncRes.json();
            if (syncRes.ok && syncData.success) {
              isSheetSynced = true;
            } else {
              sheetSyncError = syncData.error || "Lỗi máy chủ";
            }
          } else {
            isSheetSynced = true; // No updates needed
          }
        } catch (syncErr: any) {
          console.error("Auto-sync error:", syncErr);
          sheetSyncError = syncErr.message;
        }

        // 2b. Write to "Số ngày phép" if requested
        if (shouldUpdateAnnualLeave) {
          try {
            const getTongCaTrucPhaiThay = (resItem: any) => {
              let dem = 0;
              let ngaychao = new Date(resItem.start);
              const ngayEnd = new Date(resItem.end);
              while (ngaychao <= ngayEnd) {
                const cahientai = xacDinhCa(ngaychao, resItem.kip);
                if (cahientai !== 'O') dem++;
                ngaychao.setDate(ngaychao.getDate() + 1);
              }
              return dem;
            };

            const updatesAnnualLeaves = currentResult.allResults
              .filter((r: any) => r.ten && !r.ten.includes("THIẾU NHÂN SỰ"))
              .map((r: any) => ({
                name: r.ten,
                tongcatrucphaithay: getTongCaTrucPhaiThay(r)
              }));

            if (updatesAnnualLeaves.length > 0) {
              const annualRes = await fetch('/api/sheets/update-annual-leaves', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  spreadsheetId: '1pH1-Nj4B1nauoEfO5cZG13Wlk_UrUrFDq_eucf5a-IY',
                  updates: updatesAnnualLeaves
                })
              });
              const annualData = await annualRes.json();
              if (annualRes.ok && annualData.success) {
                let msg = "Ghi bảng theo dõi phép năm thành công!";
                if (annualData.skippedNames && annualData.skippedNames.length > 0) {
                  msg += ` (Không tìm thấy tên: ${annualData.skippedNames.join(', ')})`;
                }
                console.log(msg);
              } else {
                console.error("Lỗi ghi phép năm:", annualData.error);
                sheetSyncError = (sheetSyncError ? sheetSyncError + "; " : "") + `Lỗi phép năm: ${annualData.error || "Không rõ"}`;
              }
            }
          } catch (annualErr: any) {
            console.error("Annual leave update error:", annualErr);
            sheetSyncError = (sheetSyncError ? sheetSyncError + "; " : "") + `Lỗi phép năm: ${annualErr.message}`;
          }
        }
      }

      // 3. Update waiting leaves status to 'Đã xử lý'
      if (selectedWaitingLeaveIds.length > 0) {
        const updateRes = await fetch('/api/sheets/leave-requests/update-status', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ids: selectedWaitingLeaveIds,
            status: "Đã xử lý"
          })
        });
        
        if (updateRes.ok) {
          if (isGoogleAuth) {
            let detailMsg = "Hệ thống đã đồng bộ lịch mới và tự động ghi vào Bảng báo cơm ca.";
            if (shouldUpdateAnnualLeave) {
              detailMsg += " Đồng thời đã tự động ghi số ngày phép vào sheet Số ngày phép.";
            }
            if (sheetSyncError) {
              setAlert(`🎁 Đã xuất ZIP & đổi trạng thái đơn thành 'Đã xử lý', nhưng gặp lỗi đồng bộ bảng: ${sheetSyncError}`);
            } else {
              setAlert(`🎁 Xuất trọn bộ file ZIP thành công! ${detailMsg} Trạng thái đơn chuyển thành 'Đã xử lý'.`);
            }
          } else {
            setAlert("✅ Đã xuất trọn bộ hồ sơ dạng ZIP và cập nhật trạng thái 'Đã xử lý' lên Google Sheets!");
          }
          fetchWaitingLeaves();
        } else {
          const errData = await updateRes.json();
          if (updateRes.status === 401 || errData.auth_expired) {
            setIsGoogleAuth(false);
          }
          setAlert(`✅ Xuất hồ sơ thành công nhưng lỗi cập nhật trạng thái Google Sheets: ${errData.error || ""}`);
        }
      } else {
        if (isGoogleAuth) {
          let detailMsg = "đồng bộ lịch trực, tự động ghi vào Bảng báo cơm ca";
          if (shouldUpdateAnnualLeave) {
            detailMsg += " & cập nhật số ngày phép vào sheet Số ngày phép";
          }
          if (sheetSyncError) {
            setAlert(`🎁 Đã xuất hồ sơ ZIP, nhưng gặp lỗi: ${sheetSyncError}`);
          } else {
            setAlert(`🎁 Đã xuất trọn bộ hồ sơ dạng ZIP, ${detailMsg} thành công!`);
          }
          setAlert(`🎁 Đã xuất trọn bộ hồ sơ dạng ZIP & ${detailMsg} thành công!`);
        } else {
          setAlert("✅ Đã tải xuống hồ sơ dạng ZIP thành công!");
        }
      }
    } catch (e: any) {
      console.error(e);
      setAlert(`❌ Lỗi trong quá trình xuất trọn bộ hồ sơ: ${e.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleExportAllZipAndUpdateStatus = async () => {
    if (!currentResult) return;
    
    if (isGoogleAuth) {
      setConfirmModal({
        isOpen: true,
        title: "Ghi vào Bảng theo dõi phép năm?",
        message: "Bạn có muốn ghi số ngày nghỉ phép vào bảng theo dõi phép năm không",
        confirmText: "Có, ghi bảng",
        cancelText: "Không, bỏ qua",
        onConfirm: () => {
          proceedExportAllZipAndUpdateStatus(true);
        },
        onCancel: () => {
          proceedExportAllZipAndUpdateStatus(false);
        }
      });
    } else {
      await proceedExportAllZipAndUpdateStatus(false);
    }
  };

  const updateSwapGoogleSheets = async () => {
    if (!isGoogleAuth || !swapData.person1 || !swapData.person2) return;
    setIsUpdatingSheets(true);
    try {
      const updateMap: Record<string, string> = {}; // key: "name|date"

      const p1 = swapData.person1.trim().normalize('NFC');
      const p2 = swapData.person2.trim().normalize('NFC');

      if (swapData.shift1 !== 'None' && swapData.shift2 !== 'None') {
        if (swapData.date1 === swapData.date2) {
          // Same day swap: P1 takes P2's shift, P2 takes P1's shift
          updateMap[`${p1}|${swapData.date1}`] = swapData.shift2;
          updateMap[`${p2}|${swapData.date1}`] = swapData.shift1;
        } else {
          // Different day swap
          updateMap[`${p1}|${swapData.date1}`] = 'O';
          updateMap[`${p1}|${swapData.date2}`] = swapData.shift2;
          updateMap[`${p2}|${swapData.date1}`] = swapData.shift1;
          updateMap[`${p2}|${swapData.date2}`] = 'O';
        }
      } else if (swapData.shift1 !== 'None' && swapData.shift2 === 'None') {
        // P1 absent, P2 covers P1. P2 has no shift to give back.
        updateMap[`${p1}|${swapData.date1}`] = 'O';
        updateMap[`${p2}|${swapData.date1}`] = swapData.shift1;
      } else if (swapData.shift1 === 'None' && swapData.shift2 !== 'None') {
        // P2 absent, P1 covers P2. P1 has no shift to give back.
        updateMap[`${p2}|${swapData.date2}`] = 'O';
        updateMap[`${p1}|${swapData.date2}`] = swapData.shift2;
      }

      const updates = Object.entries(updateMap).map(([key, shift]) => {
        const [name, date] = key.split('|');
        return { name, date, shift };
      });

      console.log("Sending manual swap updates to Google Sheets:", updates);

      const res = await fetch('/api/sheets/update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          spreadsheetId: '1HgGW-FvoGQXtj7V_JCMD-7Tuue0rTIM-bmohGmgqm6I',
          updates
        })
      });
      const data = await res.json();
      if (data.success) {
        setAlert('✅ Đã cập nhật Google Sheets thành công cho Lịch Đổi Ca');
      } else {
        if (res.status === 401 || data.auth_expired) {
          setIsGoogleAuth(false);
        }
        setAlert('⚠ Lỗi cập nhật Google Sheets: ' + (data.error || 'Unknown error'));
      }
    } catch (err) {
      console.error(err);
      setAlert('⚠ Lỗi kết nối Google Sheets');
    } finally {
      setIsUpdatingSheets(false);
    }
  };

  const handleExportSwap = async () => {
    setIsProcessing(true);
    await exportSwapDoc(swapData, config, signatures);
    if (isGoogleAuth) {
      await updateSwapGoogleSheets();
    }
    setIsProcessing(false);
  };

  useEffect(() => {
    if (showPreview && previewBlob && previewRef.current) {
      renderAsync(previewBlob, previewRef.current)
        .catch(err => console.error("Preview error:", err));
    }
  }, [showPreview, previewBlob]);

  const handleExportCSV = () => {
    if (!currentResult) return;
    const d = currentResult;
    let csv = '\uFEFF';
    csv += 'CÔNG TY THỦY ĐIỆN IALY - BẢNG PHÂN CÔNG TRỰC THAY NGHỈ PHÉP\n\n';
    csv += `Người nghỉ:,${d.ten}\nChức danh:,${d.chucDanh}\nKíp:,Kíp ${d.kip}\n`;
    csv += `Thời gian:,${fmtVN(d.start)} đến ${fmtVN(d.end)}\n\n`;
    csv += 'STT,Ngày,Ca nghỉ,Kíp thay,Người đi thay\n';
    d.ketQua.forEach((it: any, i: number) => {
      csv += `${i + 1},${fmtVN(it.ngay)},${it.ca},Kíp ${it.kiptructhay},${it.nguoitructhay}\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Lich_nghi_${d.ten.replace(/\s+/g, '_')}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Flatten rows for display
  const shiftOrd: Record<string, number> = { N: 0, C: 1, K: 2 };
  const allRows: any[] = [];
  if (currentResult) {
    currentResult.allResults.forEach((res: any) => {
      res.ketQua.forEach((it: any) => {
        allRows.push({
          ngay: it.ngay, ca: it.ca,
          absentKip: res.kip,
          absentTen: res.ten,
          chucDanh: res.chucDanh,
          kipThay: it.kiptructhay, nguoiThay: it.nguoitructhay,
          relievedKip: it.relievedKip, relievedTen: it.relievedTen,
          isConflict: it.isConflict, conflictNote: it.conflictNote || '',
          isOverlapDay: it.isOverlapDay,
          isCKChain: false,
          isSwap: false
        });
      });
    });
    (currentResult.extraRows || []).forEach((it: any) => {
      if (it.isSwapCRow) return;
      allRows.push({
        ngay: it.ngay, ca: it.ca,
        absentKip: it.absentKip, absentTen: it.absentTen,
        chucDanh: it.chucDanh,
        kipThay: it.kiptructhay, nguoiThay: it.nguoitructhay,
        relievedKip: it.relievedKip, relievedTen: it.relievedTen,
        isConflict: it.isConflict, conflictNote: it.conflictNote || '',
        isOverlapDay: it.isOverlapDay, 
        isCKChain: it.isCKChain,
        isSwap: it.isSwap || false, 
        isSwapCRow: false
      });
    });
    allRows.sort((a, b) => {
      if (a.ngay < b.ngay) return -1;
      if (a.ngay > b.ngay) return 1;
      return (shiftOrd[a.ca] || 0) - (shiftOrd[b.ca] || 0);
    });
  }

  const coverStats: Record<number, any> = {};
  allRows.forEach(row => {
    const k = row.kipThay;
    const cd = row.chucDanh;
    if (!coverStats[k]) coverStats[k] = { total: 0, N: 0, C: 0, K: 0, byCD: {} };
    coverStats[k].total++;
    coverStats[k][row.ca]++;
    if (!coverStats[k].byCD[cd]) coverStats[k].byCD[cd] = 0;
    coverStats[k].byCD[cd]++;
  });

  const isEvenDistribution = () => {
    const counts = Object.values(coverStats).map(s => s.total);
    if (counts.length === 0) return true;
    return (Math.max(...counts) - Math.min(...counts)) <= 1;
  };

  return (
    <div className="min-h-screen flex flex-col bg-[var(--bg)]">
      {/* Top Navbar */}
      <nav id="top-navbar" className="w-full bg-white border-b border-slate-200 py-3.5 px-4 md:px-8 shadow-sm flex flex-col lg:flex-row items-center justify-between gap-4 relative z-40">
        <div className="flex flex-col sm:flex-row items-center gap-4">
          {/* Logo container exactly like screenshot */}
          <div className="flex items-center select-none">
            <img 
              src="https://i.ibb.co/fYS7WtdB/LOGO.png" 
              alt="EVN Công Ty Thủy Điện Ialy Logo" 
              className="h-[48px] md:h-[54px] w-auto object-contain"
              referrerPolicy="no-referrer"
            />
          </div>
          {/* App Title with separator */}
          <div className="text-[#053d6c] font-black text-lg md:text-xl font-sans tracking-tight sm:border-l sm:border-slate-200 sm:pl-4 py-1 flex items-center h-full">
            Phân xưởng vận hành laly
          </div>
        </div>

        {/* Navigation Tabs on the Right */}
        <div className="flex items-center gap-4 md:gap-7 flex-wrap justify-center">
          <button
            id="nav-schedule-btn"
            className={`text-[12px] md:text-[13px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
              activeTab === 'schedule'
                ? 'text-[#00529c] border-[#00529c] font-extrabold'
                : 'text-slate-500 hover:text-[#00529c] border-transparent hover:border-[#00529c]/50'
            }`}
            onClick={() => setActiveTab('schedule')}
          >
            Lịch trực thay ca vận hành
          </button>
          <button
            id="nav-leave-btn"
            className={`text-[12px] md:text-[13px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
              activeTab === 'leave'
                ? 'text-[#00529c] border-[#00529c] font-extrabold'
                : 'text-slate-500 hover:text-[#00529c] border-transparent hover:border-[#00529c]/50'
            }`}
            onClick={() => setActiveTab('leave')}
          >
            Đơn xin nghỉ phép
          </button>
          <button
            id="nav-swap-btn"
            className={`text-[12px] md:text-[13px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
              activeTab === 'swap'
                ? 'text-[#00529c] border-[#00529c] font-extrabold'
                : 'text-slate-500 hover:text-[#00529c] border-transparent hover:border-[#00529c]/50'
            }`}
            onClick={() => setActiveTab('swap')}
          >
            Đơn đổi ca
          </button>
          <button
            id="nav-staff-btn"
            className={`text-[12px] md:text-[13px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
              activeTab === 'staff'
                ? 'text-[#00529c] border-[#00529c] font-extrabold'
                : 'text-slate-500 hover:text-[#00529c] border-transparent hover:border-[#00529c]/50'
            }`}
            onClick={() => setActiveTab('staff')}
          >
            Nhân sự
          </button>
        </div>
      </nav>

      {/* Hero Banner with beautiful high resolution background */}
      <div 
        id="hero-banner" 
        className="w-full h-[220px] md:h-[280px] lg:h-[320px] relative bg-cover bg-center flex items-center justify-center overflow-hidden" 
        style={{ 
          backgroundImage: "url('https://i.ibb.co/jZ6dDJzT/z7116558150434-802a4bd8dff3b332930235031b93fc49.jpg')",
          backgroundPosition: "center 42%"
        }}
      >
        <div className="absolute inset-0 bg-slate-950/20"></div>
        
      </div>

      {/* Container for Main Content */}
      <div className="wrap !pt-8 !pb-16 flex-1 w-full max-w-[1200px] mx-auto px-4 md:px-6">

      {activeTab === 'leave' && (
        <>
          <div className="card" id="leave-creator-card">
        <div className="ctitle">Tạo đơn xin nghỉ phép</div>
        <p className="text-[13px] text-var(--txt2) mb-4">Sử dụng chức năng này để tạo nhanh văn bản "Đơn xin nghỉ phép".</p>
        <div className="g2">
          <div className="field">
            <label>Chức danh</label>
            <select value={leaveData.chucDanh} onChange={e => {
              const cd = e.target.value;
              setLeaveData({...leaveData, chucDanh: cd, name: ''});
            }}>
              {staffData.map(r => <option key={r[0]} value={r[0]}>{r[0]}</option>)}
            </select>
          </div>
          <div className="field">
            <label>Kíp</label>
            <select value={leaveData.kip} onChange={e => setLeaveData({...leaveData, kip: e.target.value, name: ''})}>
              {[1, 2, 3, 4, 5].map(k => <option key={k} value={String(k)}>Kíp {k}</option>)}
            </select>
          </div>
          <div className="field">
            <label>Họ và tên</label>
            <input 
              type="text" 
              list="leave-staff-list"
              placeholder="Chọn hoặc nhập tên"
              value={leaveData.name} 
              onChange={e => setLeaveData({...leaveData, name: e.target.value})} 
            />
            <datalist id="leave-staff-list">
              {staffData.find(r => r[0] === leaveData.chucDanh)?.slice(1).filter(Boolean).map(name => (
                <option key={name} value={name} />
              ))}
            </datalist>
          </div>
          <div className="field">
            <label>Sinh năm</label>
            <input 
              type="number" 
              value={leaveData.birthYear} 
              onChange={e => setLeaveData({...leaveData, birthYear: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Ngày bắt đầu nghỉ</label>
            <input 
              type="date" 
              value={leaveData.startDate} 
              onChange={e => setLeaveData({...leaveData, startDate: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Ngày kết thúc nghỉ</label>
            <input 
              type="date" 
              value={leaveData.endDate} 
              onChange={e => setLeaveData({...leaveData, endDate: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Chế độ phép năm</label>
            <input 
              type="number" 
              value={leaveData.leaveYear} 
              onChange={e => setLeaveData({...leaveData, leaveYear: e.target.value})} 
            />
          </div>
          <div className="field col-span-2">
            <label>Lý do nghỉ</label>
            <input 
              type="text" 
              value={leaveData.reason} 
              onChange={e => setLeaveData({...leaveData, reason: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Điện thoại liên hệ</label>
            <input 
              type="text" 
              value={leaveData.phone} 
              onChange={e => setLeaveData({...leaveData, phone: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Địa điểm</label>
            <input 
              type="text" 
              value={leaveData.location} 
              onChange={e => setLeaveData({...leaveData, location: e.target.value})} 
            />
          </div>

          {/* Leave Balance Box from Google Sheets */}
          <div className="col-span-2 mt-2 p-3 bg-slate-50 border border-slate-100 rounded-xl">
            <div className="flex justify-between items-center mb-2">
              <span className="text-[12px] font-bold text-slate-700 uppercase tracking-wider">Thông tin phép năm </span>
              {leaveData.name && (
                <button 
                  type="button" 
                  onClick={() => fetchLeaveBalance(leaveData.name)}
                  className="text-xs text-[#00529c] hover:underline flex items-center gap-1 font-semibold cursor-pointer"
                  disabled={isLoadingLeaveBalance}
                >
                  {isLoadingLeaveBalance ? <span className="spin mr-1"></span> : '🔄'} Cập nhật
                </button>
              )}
            </div>
            {isLoadingLeaveBalance ? (
              <div className="text-xs text-slate-500 py-2 flex items-center gap-2 justify-center">
                <span className="spin"></span> Đang tải thông tin phép năm ...
              </div>
            ) : leaveBalance ? (
              <div className="grid grid-cols-3 gap-2">
                <div className="bg-blue-50/60 p-2.5 rounded-lg border border-blue-100 text-center">
                  <div className="text-[10px] text-blue-700 font-medium uppercase">Phép được hưởng</div>
                  <div className="text-lg font-extrabold text-blue-900 mt-0.5">{leaveBalance.entitled} <span className="text-xs font-normal">ngày</span></div>
                </div>
                <div className="bg-amber-50/60 p-2.5 rounded-lg border border-amber-100 text-center">
                  <div className="text-[10px] text-amber-700 font-medium uppercase">Phép đã nghỉ</div>
                  <div className="text-lg font-extrabold text-amber-900 mt-0.5">{leaveBalance.used} <span className="text-xs font-normal">ngày</span></div>
                </div>
                <div className="bg-emerald-50/60 p-2.5 rounded-lg border border-emerald-100 text-center">
                  <div className="text-[10px] text-emerald-700 font-medium uppercase">Phép còn lại</div>
                  <div className="text-lg font-extrabold text-emerald-900 mt-0.5">{leaveBalance.remaining} <span className="text-xs font-normal">ngày</span></div>
                </div>
              </div>
            ) : leaveBalanceError ? (
              <div className="text-xs text-amber-600 font-medium p-2 bg-amber-50 rounded-lg border border-amber-100 text-center">
                ⚠ {leaveBalanceError}
              </div>
            ) : null}
          </div>
        </div>
        <div className="flex gap-3 mt-4 flex-wrap">
          <button className="btn btn-primary flex-1 min-w-[200px]" onClick={handlePreviewLeave} disabled={isProcessing || !leaveData.name}>
            {isProcessing ? <span className="spin mr-2"></span> : '📝'} Xem trước Đơn Nghỉ Phép
          </button>
          <button className="btn btn-secondary flex-1 min-w-[150px]" onClick={handleExportLeave} disabled={!leaveData.name}>
            📥 Xuất File Word
          </button>
        </div>
      </div>
        </>
      )}
      {/* End Leave Tab */}

      {activeTab === 'schedule' && (
        <>
          <div className="card" id="upload-leave-card">
            <div className="ctitle">Tải lên đơn xin nghỉ phép</div>
            <div className="p-4 border-2 border-dashed border-var(--acc-light) rounded-xl text-center hover:bg-var(--acc-light) transition-colors cursor-pointer relative">
              <input 
                type="file" 
                accept=".docx" 
                multiple
                onChange={handleFileUpload} 
                className="absolute inset-0 opacity-0 cursor-pointer"
                disabled={isParsing}
              />
              <div className="flex flex-col items-center gap-2">
                <span className="text-2xl">{isParsing ? '⌛' : '📄'}</span>
                <span className="text-[13px] font-medium">
                  {isParsing ? 'Đang xử lý...' : 'Nhấn để chọn hoặc kéo thả các file Word (.docx)'}
                </span>
                <span className="text-[11px] text-var(--txt2)">Hệ thống sẽ lấy thông tin từ các đơn nghỉ phép </span>
              </div>
            </div>
          </div>

          <div className="card" id="schedule-planner-card">
        <div className="ctitle">Thông tin nghỉ phép
        <button
  className="btn-ex btn-word"
  onClick={() => window.open("https://script.google.com/macros/s/AKfycbwo8YVh0YbMLg3KSMoULVRG3moktSEodmY-H3ppk1ZJ0iia6hKxC-xkCKi6-WtKlBpG/exec", "_blank")}
>
  Bảng báo cơm ca
</button>
</div>
        {alert && <div className={`alert ${alert.startsWith('✅') ? 'asuc' : 'aerr'}`}>{alert}</div>}
        
        {isGoogleAuth && (
          <div className="mb-6 p-4 border border-var(--acc-light) rounded-xl bg-var(--acc-light-5) hover:border-var(--acc) transition-all">
            <div className="flex items-center justify-between mb-3 border-b pb-2 border-var(--acc-light) flex-wrap gap-2">
              <div className="flex items-center gap-2">
                <span className="text-sm font-bold text-var(--acc) uppercase tracking-wider flex items-center gap-1.5">
                  📥 Đơn nghỉ chờ xếp lịch đi ca ({waitingLeaves.length})
                </span>
                {isLoadingWaitingLeaves && <div className="spin w-3.5 h-3.5 border-2 border-t-transparent border-var(--acc) rounded-full animate-spin"></div>}
              </div>
              <div className="flex gap-2 select-none">
                <button 
                  className="px-2 py-1 text-xs bg-white text-var(--acc) border border-var(--acc-light) rounded hover:bg-var(--acc-light-10) font-medium cursor-pointer transition-colors"
                  onClick={fetchWaitingLeaves}
                  disabled={isLoadingWaitingLeaves}
                >
                  🔄 Cập nhật
                </button>
                {selectedWaitingLeaveIds.length > 0 && (
                  <>
                    <button 
                      className="px-2.5 py-1 text-xs bg-[#4f46e5] text-white rounded hover:bg-[#4338ca] font-bold cursor-pointer transition-colors"
                      onClick={applyWaitingLeavesToForm}
                    >
                      ⚡ Áp dụng {selectedWaitingLeaveIds.length} người
                    </button>
                    <button 
                      className="px-2.5 py-1 text-xs bg-red-600 text-white rounded hover:bg-red-700 font-bold cursor-pointer transition-colors flex items-center gap-1"
                      onClick={() => handleDeleteWaitingLeaves(selectedWaitingLeaveIds)}
                    >
                      <Trash2 size={12} /> Xoá {selectedWaitingLeaveIds.length} đơn
                    </button>
                  </>
                )}
              </div>
            </div>
            
            {waitingLeaves.length === 0 ? (
              <p className="text-[12px] text-var(--txt2) italic">Không có đơn xin nghỉ phép nào ở trạng thái Chờ phân ca trên Google Sheets.</p>
            ) : (
              <div className="max-h-[220px] overflow-y-auto pr-1">
                <div className="flex flex-col gap-2">
                  {waitingLeaves.map((leave) => {
                    const isSelected = selectedWaitingLeaveIds.includes(leave.id);
                    return (
                      <div 
                        key={leave.id} 
                        className={`p-2.5 rounded-lg border text-[12px] flex items-start gap-2.5 transition-all cursor-pointer ${
                          isSelected ? 'border-[var(--acc)] bg-cyan-50/90 shadow-sm ring-1 ring-[var(--acc)]' : 'border-slate-200 bg-white hover:border-slate-300 shadow-sm'
                        }`}
                        onClick={() => {
                          setSelectedWaitingLeaveIds(prev => 
                            prev.includes(leave.id) 
                              ? prev.filter(id => id !== leave.id) 
                              : [...prev, leave.id]
                          );
                        }}
                      >
                        <input 
                          type="checkbox" 
                          checked={isSelected}
                          onChange={() => {}} // handled by parent div click
                          className="mt-0.5 accent-[var(--acc)] cursor-pointer"
                        />
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center justify-between gap-2">
                            <span className="font-bold text-slate-950 text-[13px]">{leave.name}</span>
                            <div className="flex items-center gap-1.5" onClick={e => e.stopPropagation()}>
                              <span className="text-[10px] bg-yellow-100 text-yellow-800 border border-yellow-200 px-1.5 py-0.5 rounded-md font-mono font-medium">{leave.id.substring(0, 14)}...</span>
                              <button 
                                className="p-1 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded transition-colors cursor-pointer"
                                title="Xóa đơn này"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleDeleteWaitingLeaves([leave.id], leave.name);
                                }}
                              >
                                <Trash2 size={13} className="text-slate-400 hover:text-red-600 transition-colors" />
                              </button>
                            </div>
                          </div>
                          <div className="grid grid-cols-2 gap-x-3 gap-y-1 mt-1 text-slate-500">
                            <div>Kíp: <span className="font-semibold text-slate-800">Kíp {leave.kip}</span></div>
                            <div>Chức danh: <span className="font-semibold text-slate-800">{leave.chucDanh}</span></div>
                            <div className="col-span-2">Thời gian: <span className="font-semibold text-slate-800">{fmtVN(new Date(leave.startDate))} → {fmtVN(new Date(leave.endDate))}</span></div>
                            <div className="col-span-2">Lý do: <span className="font-semibold text-slate-800 truncate block" title={leave.reason}>{leave.reason}</span></div>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
            <p className="text-[10px] text-var(--txt2) mt-2 italic">* Chọn các đơn muốn xếp lịch và bấm "Áp dụng ... người" để tự động điền nhanh các thông tin.</p>
          </div>
        )}

        <div className="g2">
          <div className="field">
            <label>Ngày bắt đầu nghỉ</label>
            <input type="date" value={ngayBatDau} onChange={e => setNgayBatDau(e.target.value)} />
          </div>
          <div className="field">
            <label>Ngày kết thúc nghỉ</label>
            <input type="date" value={ngayKetThuc} onChange={e => setNgayKetThuc(e.target.value)} />
          </div>
          <div className="field">
            <label>Chức danh người nghỉ</label>
            <select value={chucDanh} onChange={e => setChucDanh(e.target.value)}>
              <option value="">-- Chọn chức danh --</option>
              {staffData.map(r => <option key={r[0]} value={r[0]}>{r[0]}</option>)}
            </select>
          </div>
          <div className="field">
            <label>Kíp nghỉ</label>
            <select value={kipNghi} onChange={e => setKipNghi(e.target.value)}>
              <option value="">-- Chọn kíp --</option>
              {[1, 2, 3, 4, 5].map(k => <option key={k} value={k}>Kíp {k}</option>)}
            </select>
          </div>
        </div>

        <div className="divider"></div>
        <div className="flex items-center justify-between mb-3">
          <span className="text-[12px] text-var(--txt2) font-semibold uppercase tracking-wider">
            Người nghỉ đồng thời <span className="font-normal">(có thể khác chức danh)</span>
          </span>
          <button className="staff-toggle font-bold text-[13px]" onClick={addConcurrentLeave}>+ Thêm người nghỉ</button>
        </div>

        <div id="concurrent-list">
          {additionalLeaves.map((leave, idx) => (
            <div key={idx} className="concurrent-item">
              <div className="absolute top-2 right-2.5">
                <button onClick={() => removeConcurrentLeave(idx)} className="text-var(--acc3) cursor-pointer text-lg leading-none" title="Xóa">✕</button>
              </div>
              <div className="text-[11px] font-bold text-var(--acc) uppercase tracking-wider mb-3">Người nghỉ #{idx + 2}</div>
              <div className="g2">
                <div className="field">
                  <label>Chức danh</label>
                  <select value={leave.chucDanh} onChange={e => updateConcurrentLeave(idx, 'chucDanh', e.target.value)}>
                    <option value="">-- Chức danh --</option>
                    {staffData.map(r => <option key={r[0]} value={r[0]}>{r[0]}</option>)}
                  </select>
                </div>
                <div className="field">
                  <label>Kíp nghỉ</label>
                  <select value={leave.kip} onChange={e => updateConcurrentLeave(idx, 'kip', e.target.value)}>
                    <option value="">-- Kíp --</option>
                    {[1, 2, 3, 4, 5].map(k => <option key={k} value={k}>Kíp {k}</option>)}
                  </select>
                </div>
                <div className="field">
                  <label>Ngày bắt đầu</label>
                  <input type="date" value={leave.start} onChange={e => updateConcurrentLeave(idx, 'start', e.target.value)} />
                </div>
                <div className="field">
                  <label>Ngày kết thúc</label>
                  <input type="date" value={leave.end} onChange={e => updateConcurrentLeave(idx, 'end', e.target.value)} />
                </div>
              </div>
            </div>
          ))}
        </div>

        <button className="btn btn-primary" onClick={taoLich} disabled={isProcessing}>
          {isProcessing ? <span className="spin mr-2"></span> : '⚡'} Tạo Lịch Thay Ca
        </button>
      </div>

      {/* End Schedule Tab Content */}
        </>
      )}

      {activeTab === 'swap' && (
        <div className="card" id="manual-swap-card">
          <div className="ctitle">Tạo lịch đổi ca thủ công</div>
        <p className="text-[13px] text-var(--txt2) mb-4">Sử dụng để tạo lịch đổi ca giữa 2 người.</p>
        <div className="g2">
          <div className="field">
            <label>Chức danh</label>
            <select value={swapChucDanh} onChange={e => {
              setSwapChucDanh(e.target.value);
              setSwapData({...swapData, person1: '', person2: ''});
            }}>
              {staffData.map(r => <option key={r[0]} value={r[0]}>{r[0]}</option>)}
            </select>
          </div>
          <div className="field"></div>
          <div className="field">
            <label>Ngày đổi ca 1</label>
            <input 
              type="date" 
              value={swapData.date1} 
              onChange={e => setSwapData({...swapData, date1: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Ngày đổi ca 2</label>
            <input 
              type="date" 
              value={swapData.date2} 
              onChange={e => setSwapData({...swapData, date2: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Người đổi ca (P1)</label>
            <input 
              type="text" 
              list="staff-list"
              placeholder="Chọn hoặc nhập tên"
              value={swapData.person1} 
              onChange={e => setSwapData({...swapData, person1: e.target.value})} 
            />
          </div>
          <div className="field">
            <label>Người đi ca thay (P2)</label>
            <input 
              type="text" 
              list="staff-list"
              placeholder="Chọn hoặc nhập tên"
              value={swapData.person2} 
              onChange={e => setSwapData({...swapData, person2: e.target.value})} 
            />
          </div>
          <datalist id="staff-list">
            {staffData.find(r => r[0] === swapChucDanh)?.slice(1).filter(Boolean).map(name => (
              <option key={name} value={name} />
            ))}
          </datalist>
          <div className="field">
            <label>Ca của P1 (nghỉ)</label>
            <select value={swapData.shift1} onChange={e => setSwapData({...swapData, shift1: e.target.value})}>
              <option value="N">Ca N</option>
              <option value="C">Ca C</option>
              <option value="K">Ca K</option>
              <option value="None">Không</option>
            </select>
          </div>
          <div className="field">
            <label>Ca của P2 (nghỉ)</label>
            <select value={swapData.shift2} onChange={e => setSwapData({...swapData, shift2: e.target.value})}>
              <option value="N">Ca N</option>
              <option value="C">Ca C</option>
              <option value="K">Ca K</option>
              <option value="None">Không</option>
            </select>
          </div>
        </div>
        <div className="flex gap-3 mt-4">
          <button className="btn btn-primary flex-1" onClick={handlePreviewSwap} disabled={isProcessing || !swapData.person1 || !swapData.person2}>
            {isProcessing ? <span className="spin mr-2"></span> : '📝'} Xem trước Lịch Đổi Ca
          </button>
          <button className="btn btn-secondary flex-1" onClick={handleExportSwap} disabled={!swapData.person1 || !swapData.person2}>
            📥 Xuất File Word
          </button>
        </div>
      </div>
      )}

      {activeTab === 'staff' && (
        <div className="flex flex-col gap-6 w-full">
          <div className="card" id="staff-roster-card">
            <div className="ctitle">
              Nhân sự của kíp 
              <div className="flex gap-2">
                <button className="staff-toggle" onClick={() => setShowStaff(!showStaff)}>
                  {showStaff ? 'Thu gọn ▲' : 'Chỉnh sửa ▼'}
                </button>
                <button className="staff-toggle" onClick={() => setShowSignatureManager(!showSignatureManager)}>
                  {showSignatureManager ? '✍️ Ẩn chữ ký' : '✍️ Quản lý chữ ký'}
                </button>
              </div>
            </div>
            <p className="text-[13px] text-var(--txt2)">Nhấn "Chỉnh sửa" để cập nhật tên nhân viên hoặc "Quản lý chữ ký" để tải lên ảnh chữ ký.</p>
            
            {showSignatureManager && (
              <div className="mb-6">
                <SignatureManager 
                  staffList={Array.from(new Set(staffData.flatMap(row => row.slice(1)).filter(Boolean)))} 
                  signatures={signatures}
                  onSignaturesChange={setSignatures} 
                />
              </div>
            )}

            {showStaff && (
              <div className="staff-wrap">
                <table className="st">
                  <thead>
                    <tr>
                      <th>Chức danh</th>
                      <th>Kíp 1</th>
                      <th>Kíp 2</th>
                      <th>Kíp 3</th>
                      <th>Kíp 4</th>
                      <th>Kíp 5</th>
                    </tr>
                  </thead>
                  <tbody>
                    {staffData.map((row, r) => (
                      <tr key={r}>
                        <td>{row[0]}</td>
                        {[1, 2, 3, 4, 5].map(c => (
                          <td key={c}>
                            <input 
                              value={row[c] || ''} 
                              onChange={e => handleUpdateStaff(r, c, e.target.value)}
                            />
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          <div className="card" id="system-config-card">
            <div className="ctitle">Cấu hình hệ thống</div>
            <p className="text-[13px] text-var(--txt2) mb-4">Điều chỉnh thông tin ký số văn bản và tích hợp thông báo Zalo cá nhân.</p>
            <div className="g2">
              <div className="field">
                <label>Người ký văn bản</label>
                <input 
                  type="text" 
                  value={config.nguoiKy || ''} 
                  onChange={e => setConfig({ ...config, nguoiKy: e.target.value })} 
                  placeholder="Họ và tên người ký"
                />
              </div>
              <div className="field">
                <label>Số văn bản mặc định</label>
                <input 
                  type="text" 
                  value={config.soVanBan || ''} 
                  onChange={e => setConfig({ ...config, soVanBan: e.target.value })} 
                  placeholder="Ví dụ: 123"
                />
              </div>
              <div className="field col-span-2">
                <label>Zalo Webhook URL nhận thông báo</label>
                <input 
                  type="text" 
                  value={config.zaloWebhookUrl || ''} 
                  onChange={e => setConfig({ ...config, zaloWebhookUrl: e.target.value })} 
                  placeholder="https://..."
                  className="font-mono text-[12.5px] w-full px-3 py-2 border rounded-md"
                />
                <span className="text-[11px] text-var(--txt2) italic mt-1.5 block">
                  * Hệ thống sẽ tự động gửi thông báo về Zalo cá nhân qua đường dẫn webhook này khi có đơn nghỉ phép mới được tạo hoặc lưu lên hệ thống chờ phân ca.
                </span>
              </div>
            </div>
          </div>
        </div>
      )}

      {activeTab === 'schedule' && currentResult && (
        <div id="results">
          <div className="card">
            <div className="res-hdr">
              <div className="ctitle mb-0">Kết quả phân công</div>
              <div className="ex-row flex-wrap gap-2">
                {isGoogleAuth && selectedWaitingLeaveIds.length > 0 && (
                  <button className="btn-ex btn-indigo font-medium flex items-center justify-center transition-colors" onClick={handleExportAllZipAndUpdateStatus} disabled={isProcessing}>
                    {isProcessing ? <span className="spin spinw mr-2"></span> : '📝'} Tạo lịch thay ca và các đơn nghỉ phép
                  </button>
                )}
                <button className="btn-ex btn-word" onClick={handlePreviewWord} disabled={isProcessing}>
                  {isProcessing ? <span className="spin spinw mr-2"></span> : '📝'} Xem trước Word
                </button>
                <button className="btn-ex btn-word" onClick={handleExportWord} disabled={isProcessing}>
                  {isProcessing ? <span className="spin spinw mr-2"></span> : '📝'} Xuất Word
                </button>
                {loadingAuth ? (
                  <button className="btn-ex btn-gray cursor-wait" disabled>
                    <span className="spin mr-2 border-gray-400"></span> Đang kiểm tra...
                  </button>
                ) : !isGoogleAuth ? (
                  <button className="btn-ex btn-google" onClick={handleConnectGoogle}>
                    🔗 Kết nối Google Sheets
                  </button>
                ) : (
                  <div className="flex gap-2">
                    <button 
                      className={`btn-ex ${isUpdatingSheets ? 'btn-gray' : 'btn-sheets'}`} 
                      onClick={updateGoogleSheets}
                      disabled={isUpdatingSheets}
                    >
                      {isUpdatingSheets ? <span className="spin spinw mr-2"></span> : '🔄'} 
                      {isUpdatingSheets ? 'Đang đồng bộ...' : 'Đồng bộ Google Sheets'}
                    </button>
                    <button 
                      className="btn-ex btn-danger"
                      onClick={() => {
                        setConfirmModal({
                          isOpen: true,
                          title: "Ngắt kết nối Google Sheets",
                          message: "Bạn có chắc chắn muốn ngắt kết nối Google Sheets? Bạn sẽ cần đăng nhập lại để đồng bộ và truy xuất dữ liệu.",
                          confirmText: "Ngắt kết nối",
                          cancelText: "Hủy",
                          isDanger: true,
                          onConfirm: () => {
                            setIsGoogleAuth(false);
                            setAlert('Đã ngắt kết nối. Vui lòng kết nối lại để cập nhật.');
                          }
                        });
                      }}
                      title="Ngắt kết nối và đăng nhập lại"
                    >
                      🚫
                    </button>
                  </div>
                )}
              </div>
            </div>

            <div className="igrid">
              {currentResult.isMulti ? (
                <>
                  <div className="iitem"><div className="ilbl">Chức danh</div><div className="ival">{currentResult.chucDanh}</div></div>
                  <div className="iitem"><div className="ilbl">Số người nghỉ</div><div className="ival">{currentResult.allResults.length} người</div></div>
                  <div className="iitem"><div className="ilbl">Người nghỉ</div><div className="ival">{currentResult.allResults.map((r: any) => `${r.ten} (K${r.kip})`).join(' | ')}</div></div>
                  <div className="iitem"><div className="ilbl">Tổng ca thay</div><div className="ival">{allRows.length} ca</div></div>
                </>
              ) : (
                <>
                  <div className="iitem"><div className="ilbl">Người nghỉ</div><div className="ival">{currentResult.ten}</div></div>
                  <div className="iitem"><div className="ilbl">Chức danh</div><div className="ival">{currentResult.chucDanh}</div></div>
                  <div className="iitem"><div className="ilbl">Kíp nghỉ</div><div className="ival">Kíp {currentResult.kip}</div></div>
                  <div className="iitem"><div className="ilbl">Thời gian</div><div className="ival">{fmtVN(currentResult.start)} → {fmtVN(currentResult.end)}</div></div>
                  <div className="iitem"><div className="ilbl">Số ngày</div><div className="ival">{Math.round((currentResult.end - currentResult.start) / 86400000) + 1} ngày</div></div>
                  <div className="iitem"><div className="ilbl">Ca cần thay</div><div className="ival">{allRows.length} ca</div></div>
                </>
              )}
            </div>

            {currentResult.isMulti && Object.keys(coverStats).length > 0 && (
              <div className="distrib-box">
                <div className="flex items-center justify-between">
                  <span className="text-[11px] font-bold text-var(--acc) uppercase tracking-wider">Phân bố ca thay</span>
                  {isEvenDistribution() ? (
                    <span className="text-[11px] color-[#22c55e] font-semibold">✓ Phân bố đều</span>
                  ) : (
                    <span className="text-[11px] color-var(--C) font-semibold">⚠ Chưa đều</span>
                  )}
                </div>
                <div className="distrib-grid">
                  {Object.keys(coverStats).sort().map(kip => (
                    <div key={kip} className="distrib-item">
                      <div className="text-[13px] font-bold text-var(--txt)">
                        Kíp {kip} 
                        <span className="text-[11px] font-normal text-var(--txt2) ml-1">
                          ({Object.entries(coverStats[+kip].byCD).map(([cd, count]) => `${cd}: ${count}`).join(', ')})
                        </span>
                      </div>
                      <div className="mt-1.5 flex gap-1.5 items-center flex-wrap">
                        {coverStats[+kip].N > 0 && <span className="badge bN">N ×{coverStats[+kip].N}</span>}
                        {coverStats[+kip].C > 0 && <span className="badge bC">C ×{coverStats[+kip].C}</span>}
                        {coverStats[+kip].K > 0 && <span className="badge bK">K ×{coverStats[+kip].K}</span>}
                        <span className="text-var(--acc) font-bold text-[12px]">= {coverStats[+kip].total} ca</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="rt-wrap">
              <table className="rt">
                <thead>
                  <tr>
                    <th>STT</th>
                    <th>Ngày</th>
                    <th>Ca</th>
                    <th>Người nghỉ (Kíp)</th>
                    <th>Kíp thay</th>
                    <th>Người đi thay</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {allRows.length === 0 ? (
                    <tr>
                      <td colSpan={7}>
                        <div className="no-res">📅 Không có ca làm việc nào trong khoảng thời gian này.</div>
                      </td>
                    </tr>
                  ) : (
                    allRows.map((row, i) => {
                      const ds = fmtVN(row.ngay);
                      const isNewDate = i === 0 || fmtVN(allRows[i - 1].ngay) !== ds;
                      return (
                        <tr key={i} className={`${isNewDate && i > 0 ? 'date-sep' : ''} ${row.isConflict ? 'conflict-row' : ''}`}>
                          <td className="text-var(--txt2) font-mono">{i + 1}</td>
                          <td>
                            {isNewDate && (
                              <>
                                <span className="font-semibold">{ds}</span>
                                <span className="text-var(--txt2) text-[11px] ml-1">{dayN(row.ngay)}</span>
                                {row.isOverlapDay && (
                                  <div className="text-[10px] text-var(--acc) font-bold">👥 Ngày trùng nghỉ</div>
                                )}
                              </>
                            )}
                          </td>
                          <td><span className={`badge b${row.ca}`}>{row.ca}</span></td>
                          <td className="text-[12px] text-var(--txt2)">
                            <div className="font-bold text-[10px] text-var(--acc) mb-0.5 uppercase">{row.chucDanh}</div>
                            {row.isSwap ? (
                              <>
                                <span className="text-[#22c55e] text-[11px]">đổi ca</span>
                                <br />
                                {row.relievedTen || row.absentTen}
                                <br />
                                <span className="text-[10px]">Kíp {row.relievedKip || row.absentKip}</span>
                              </>
                            ) : row.isCKChain ? (
                              <>
                                <span className="text-[#a855f7] text-[11px]">thay ca</span>
                                <br />
                                {row.absentTen}
                                <br />
                                <span className="text-[10px]">Kíp {row.absentKip}</span>
                              </>
                            ) : (
                              <>
                                {row.absentTen}
                                <br />
                                <span className="text-[10px]">Kíp {row.absentKip}</span>
                              </>
                            )}
                          </td>
                          <td className="text-var(--txt2)">Kíp {row.kipThay}</td>
                          <td className="font-semibold text-var(--acc2)">{row.nguoiThay}</td>
                          <td className="max-w-[140px]">
                            {row.isSwap ? (
                              <>
                                <span className="conflict-badge bg-[#22c55e1a] text-[#22c55e] border-[#22c55e4d]">⇄ Đổi ca</span>
                                {row.conflictNote && (
                                  <div className="text-[10px] text-var(--txt2)">{row.conflictNote}</div>
                                )}
                              </>
                            ) : row.isCKChain ? (
                              <>
                                <span className="conflict-badge bg-[#a855f71a] text-[#a855f7] border-[#a855f74d]">⥵ C→K</span>
                                <br />
                                <span className="text-[10px] text-var(--txt2)">Thay do ràng buộc Ca C→K</span>
                              </>
                            ) : row.isConflict ? (
                              <>
                                <span className="conflict-badge">△ Điều chỉnh</span>
                                {row.conflictNote && (
                                  <div className="text-[10px] text-var(--txt2)">{row.conflictNote}</div>
                                )}
                              </>
                            ) : null}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            <div className="legend">
              <div className="legend-item"><span className="badge bN">N</span>Ca Ngày (08:00–16:00)</div>
              <div className="legend-item"><span className="badge bC">C</span>Ca Chiều (16:00–22:20)</div>
              <div className="legend-item"><span className="badge bK">K</span>Ca Đêm (22:20–08:00)</div>
            </div>
          </div>
        </div>
      )}
      {showPreview && (
        <div className="modal-overlay" onClick={() => setShowPreview(false)}>
          <div className="modal-content" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h3>{isPreviewingSwap ? 'Xem trước Lịch Đổi Ca' : isPreviewingLeave ? 'Xem trước Đơn Nghỉ Phép' : 'Xem trước Lịch Trực Thay Ca'}</h3>
              <button className="close-btn" onClick={() => setShowPreview(false)}>✕</button>
            </div>
            <div className="modal-body">
              <div ref={previewRef} className="docx-container"></div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setShowPreview(false)}>Đóng</button>
              {isPreviewingSwap ? (
                <button className="btn btn-primary" onClick={handleExportSwap}>Tải xuống Word</button>
              ) : isPreviewingLeave ? (
                <>
                  <button 
                    className={`btn ${isSavingLeaveToSheets ? 'bg-gray-400 cursor-wait' : 'bg-[#34a853] text-white hover:bg-[#2c8e46] cursor-pointer'} font-medium rounded-lg px-4 py-2 text-sm flex items-center justify-center gap-1.5`}
                    onClick={saveLeaveToGoogleSheets}
                    disabled={isSavingLeaveToSheets || !leaveData.name}
                  >
                    {isSavingLeaveToSheets ? <span className="spin spinw mr-2"></span> : '☁️ '}
                    {isSavingLeaveToSheets ? 'Đang lưu...' : 'Lưu lên hệ thống'}
                  </button>
                  <button className="btn btn-primary" onClick={handleExportLeave}>Tải xuống Word</button>
                </>
              ) : (
                <button 
                  className="btn btn-primary flex items-center justify-center gap-1.5" 
                  onClick={() => {
                    handleExportAllZipAndUpdateStatus();
                    setShowPreview(false);
                  }}
                  disabled={isProcessing}
                >
                  {isProcessing ? <span className="spin spinw mr-2"></span> : '🎁'} Tạo lịch thay ca và các đơn nghỉ phép
                </button>
              )}
            </div>
          </div>
        </div>
      )}
      {confirmModal && confirmModal.isOpen && (
        <div className="modal-overlay" style={{ zIndex: 9999 }} onClick={closeConfirmModal}>
          <div className="modal-content !max-w-[420px] shadow-2xl border border-slate-100" onClick={e => e.stopPropagation()}>
            <div className="modal-header border-b pb-2.5 flex items-center justify-between">
              <h3 className="text-[15px] font-bold text-slate-900">{confirmModal.title}</h3>
              <button className="close-btn p-1 hover:bg-slate-100 rounded-lg text-slate-400 hover:text-slate-600 transition-colors" onClick={closeConfirmModal}>✕</button>
            </div>
            <div className="modal-body py-4">
              <p className="text-[13px] text-slate-600 font-medium leading-relaxed">{confirmModal.message}</p>
              
              {confirmModal.requirePassword && (
                <div className="mt-4">
                  <label className="block text-[12px] font-bold text-slate-700 mb-1.5 flex items-center gap-1">
                    🔒 Nhập mật khẩu xác nhận xóa đơn:
                  </label>
                  <input
                    type="password"
                    placeholder="Mật khẩu "
                    value={confirmPassword}
                    onChange={(e) => {
                      setConfirmPassword(e.target.value);
                      if (passwordError) setPasswordError("");
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        handleConfirmSubmit();
                      }
                    }}
                    className={`w-full px-3 py-2 border rounded-lg text-sm bg-slate-50 focus:bg-white focus:outline-none focus:ring-2 transition-all ${
                      passwordError 
                        ? 'border-red-500 focus:ring-red-200' 
                        : 'border-slate-200 focus:ring-indigo-100 focus:border-indigo-500'
                    }`}
                    autoFocus
                  />
                  {passwordError && (
                    <p className="mt-1.5 text-[11px] font-semibold text-red-600 flex items-center gap-1">
                      <span>⚠</span> {passwordError}
                    </p>
                  )}
                </div>
              )}
            </div>
            <div className="modal-footer border-t pt-2.5 flex justify-end gap-2.5">
              <button 
                className="px-3.5 py-1.5 text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold rounded-lg transition-colors cursor-pointer"
                onClick={() => {
                  if (confirmModal.onCancel) {
                    confirmModal.onCancel();
                  }
                  closeConfirmModal();
                }}
              >
                {confirmModal.cancelText || 'Hủy'}
              </button>
              <button 
                className={`px-4 py-1.5 text-xs font-semibold text-white rounded-lg transition-colors cursor-pointer ${
                  confirmModal.isDanger 
                    ? 'bg-red-600 hover:bg-red-700 active:bg-red-800' 
                    : 'bg-[#4f46e5] hover:bg-[#4338ca] active:bg-[#3730a3]'
                }`}
                onClick={handleConfirmSubmit}
              >
                {confirmModal.confirmText || 'Đồng ý'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>

    {/* Footer styled exactly like the provided screenshot/USAGov */}
    <footer id="app-footer" className="w-full bg-[#072540] text-slate-300 py-10 px-6 md:px-12 mt-auto border-t border-slate-900 select-none">
      <div className="max-w-[1200px] mx-auto flex flex-col md:flex-row items-center justify-between gap-6">
        <div className="flex flex-col gap-2.5 text-center md:text-left">
          <div className="text-[13px] md:text-[14px] text-white font-bold tracking-wide">
            Hệ thống tự dộng tạo lịch trực thay ca vận hành nghỉ phép VHIALY
          </div>
          <div className="text-[12.5px] text-slate-300 leading-relaxed font-medium">
            Trang thông tin hỗ trợ nội bộ trực thuộc <a href="https://ialyhpc.vn" target="_blank" rel="noopener noreferrer" className="underline hover:text-white transition-colors">Công ty Thủy Điện Ialy</a>
          </div>
          
          {/* Lower row of policy links */}
          <div className="flex flex-wrap items-center justify-center md:justify-start gap-x-6 gap-y-2 mt-2 text-[11.5px] font-semibold text-slate-400">
            <span className="underline cursor-pointer hover:text-white transition-colors">Hướng dẫn sử dụng</span>
            <span className="underline cursor-pointer hover:text-white transition-colors">Chính sách bảo mật nội bộ</span>
            <span className="underline cursor-pointer hover:text-white transition-colors">Hỗ trợ kỹ thuật </span>
          </div>
        </div>

        {/* Top scroll button exactly like the screenshot layout */}
        <div className="flex items-center">
          <button 
            onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })} 
            className="flex items-center bg-white text-slate-800 rounded-full pl-1.5 pr-4 py-1.5 transition-all outline-none border border-slate-200 cursor-pointer shadow-md hover:bg-slate-50 hover:shadow-lg active:scale-95 duration-150"
          >
            <span className="h-8 w-8 bg-[#00adef] text-white font-bold rounded-full flex items-center justify-center mr-2.5 shadow-sm leading-none text-base">
              ↑
            </span>
            <span className="text-[13px] font-extrabold tracking-wider uppercase text-[#05203c] underline decoration-[#00adef] decoration-2">
              Top
            </span>
          </button>
        </div>
      </div>
    </footer>
    </div>
  );
}

