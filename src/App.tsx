import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { DEFAULT_STAFF, SHIFTS } from './constants';
import { fmtVN, fmtIn, dayN, timNghi, timThay, abbrev, xacDinhCa } from './utils/shiftHelpers';
import { buildMultiLeaveResults, Leave, ResultItem } from './utils/Quytacxacdinhcatructhay';
import { exportWord, generateWordBlob, exportSwapDoc, generateSwapBlob, exportLeaveRequestDoc, generateLeaveRequestBlob, exportAllDocsZip } from './utils/wordExport';
import { renderAsync } from 'docx-preview';
import SignatureManager from './components/SignatureManager';
import LeaveBalanceManager from './components/LeaveBalanceManager';
import LoginForm from './components/LoginForm';
import WorkshopManagerModal from './components/WorkshopManagerModal';
import { UserAccount, Workshop } from './types/auth';
import { Trash2, Settings, LogOut, User } from 'lucide-react';
import { API_BASE } from './utils/api';

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
  // Collapsed by default: the table can run to dozens of rows and pushes the rest of
  // the staff tab off-screen, so it opens only when someone asks for it.
  const [showLeaveBalance, setShowLeaveBalance] = useState(false);
  const [alert, setAlert] = useState<string | null>(null);
  const [currentResult, setCurrentResult] = useState<any>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  // The footer already has a "back to top" control, but reaching it means scrolling to
  // the very bottom — the opposite of what someone deep in a long page wants. This one
  // follows the viewport instead.
  const [showScrollTop, setShowScrollTop] = useState(false);

  useEffect(() => {
    // Roughly one hero image down: far enough that the header is gone and the button
    // is actually useful, close enough that it appears as soon as scrolling starts.
    const onScroll = () => setShowScrollTop(window.scrollY > 240);
    onScroll();
    window.addEventListener('scroll', onScroll, { passive: true });
    return () => window.removeEventListener('scroll', onScroll);
  }, []);
  const [previewBlob, setPreviewBlob] = useState<Blob | null>(null);
  const previewRef = useRef<HTMLDivElement>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [config, setConfig] = useState({
    soVanBan: '',
    ngayKy: '',
    nguoiKy: 'Nguyễn Văn Nghị',
    zaloWebhookUrl: 'https://vhialy.dpdns.org/webhook/notify',
    notifyEmail: ''
  });

  // Login / workshop-admin system (LoginForm, UserHeaderBar, WorkshopManagerModal)
  const [currentUser, setCurrentUser] = useState<UserAccount | null>(() => {
    try {
      const saved = localStorage.getItem('auth_user');
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  });
  const [workshops, setWorkshops] = useState<Workshop[]>([]);
  const [activeWorkshop, setActiveWorkshop] = useState<Workshop | null>(null);
  const [showWorkshopManager, setShowWorkshopManager] = useState(false);
  // Which workshop the in-memory staffData/config were loaded from. staffData is
  // seeded from localStorage on first render, so before the active workshop has been
  // read into state it holds whatever the previous session left behind — autosaving
  // that would overwrite the real roster with stale (often empty) data.
  const [hydratedWsId, setHydratedWsId] = useState<string | null>(null);

  // localStorage only remembers who was logged in so the UI can render immediately;
  // it is not proof of anything. The httpOnly session cookie is the real credential,
  // so ask the server who we actually are and drop the stale user if it disagrees.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const res = await fetch(API_BASE + '/api/auth/me');
        if (cancelled) return;
        if (res.ok) {
          const data = await res.json();
          if (data?.user) {
            localStorage.setItem('auth_user', JSON.stringify(data.user));
            setCurrentUser(data.user);
            return;
          }
        }
        localStorage.removeItem('auth_user');
        setCurrentUser(null);
      } catch {
        // Network failure: leave the cached user in place rather than logging the
        // operator out mid-shift. Any real API call will still be rejected with 401.
      }
    })();
    return () => { cancelled = true; };
  }, []);

  // Once the cookie expires every request starts failing with 401. Catch that in one
  // place instead of at each of the ~30 call sites, and send the user back to login.
  useEffect(() => {
    const original = window.fetch;
    window.fetch = async (...args) => {
      const res = await original(...args);
      const url = typeof args[0] === 'string' ? args[0] : (args[0] as Request)?.url || '';
      if (res.status === 401 && url.includes('/api/') && !url.includes('/api/auth/')) {
        localStorage.removeItem('auth_user');
        setCurrentUser(null);
      }
      return res;
    };
    return () => { window.fetch = original; };
  }, []);

  const fetchWorkshops = useCallback(async () => {
    try {
      const res = await fetch(API_BASE + '/api/workshops');
      const data = await res.json();
      if (Array.isArray(data)) {
        setWorkshops(data);
        setActiveWorkshop(prev => {
          if (prev) {
            const stillExists = data.find((w: Workshop) => w.id === prev.id);
            if (stillExists) return stillExists;
          }
          if (!currentUser) return null;
          if (currentUser.role === 'super_admin') return data[0] || null;
          return data.find((w: Workshop) => w.id === currentUser.workshopId) || null;
        });
      }
    } catch (e) {
      console.error('Failed to load workshops', e);
    }
  }, [currentUser]);

  useEffect(() => {
    if (currentUser) fetchWorkshops();
  }, [currentUser, fetchWorkshops]);

  const handleLoginSuccess = (user: UserAccount) => {
    setCurrentUser(user);
  };

  const handleLogout = () => {
    fetch(API_BASE + '/api/auth/logout', { method: 'POST' }).catch(() => {});
    localStorage.removeItem('auth_user');
    setCurrentUser(null);
    setWorkshops([]);
    setActiveWorkshop(null);
    setShowWorkshopManager(false);
  };

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
    location: 'Gia Lai',
    hasLeavePermit: false
  });
  const [isPreviewingLeave, setIsPreviewingLeave] = useState(false);

  // Travel-day allowance (ngày đi đường) based on distance from Pleiku
  const [locationOptions, setLocationOptions] = useState<{ name: string; distanceKm: number }[]>([]);

  // Server-computed plan: shifts consumed, which leave year(s) they are charged to, travel days
  const [leavePlan, setLeavePlan] = useState<{
    leaveDays: number;
    leaveYear: string;
    allocations: Record<string, number>;
    detail: { date: string; shift: string; year: string }[];
    travel: { travelDays: number; distanceKm: number | null; note: string };
  } | null>(null);

  // States for Staff Leave Balance from Google Sheets
  const [leaveBalance, setLeaveBalance] = useState<{ entitled: string; used: string; remaining: string } | null>(null);
  const [isLoadingLeaveBalance, setIsLoadingLeaveBalance] = useState(false);
  const [leaveBalanceError, setLeaveBalanceError] = useState<string | null>(null);

  const fetchLeaveBalance = useCallback(async (name: string) => {
    if (!name || name.trim() === '' || !activeWorkshop) {
      setLeaveBalance(null);
      setLeaveBalanceError(null);
      return;
    }
    setIsLoadingLeaveBalance(true);
    setLeaveBalanceError(null);
    try {
      const res = await fetch(`/api/sheets/leave-balance?name=${encodeURIComponent(name.trim())}&workshopId=${encodeURIComponent(activeWorkshop.id)}`);
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
  }, [activeWorkshop]);

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

  useEffect(() => {
    fetch(API_BASE + '/api/leave/locations')
      .then(res => res.json())
      .then(data => { if (Array.isArray(data)) setLocationOptions(data); })
      .catch(() => {});
  }, []);

  // The server decides how many shifts the leave costs and which leave year pays for them
  useEffect(() => {
    if (!leaveData.startDate || !leaveData.endDate || !leaveData.kip || !activeWorkshop) {
      setLeavePlan(null);
      return;
    }
    const timer = setTimeout(() => {
      const params = new URLSearchParams({
        name: leaveData.name || '',
        kip: leaveData.kip,
        startDate: leaveData.startDate,
        endDate: leaveData.endDate,
        location: leaveData.location || '',
        hasLeavePermit: String(leaveData.hasLeavePermit),
        workshopId: activeWorkshop.id
      });
      fetch(`/api/leave/plan?${params.toString()}`)
        .then(res => res.json())
        .then(data => { if (data.success) setLeavePlan(data); })
        .catch(() => {});
    }, 500);
    return () => clearTimeout(timer);
  }, [leaveData.name, leaveData.kip, leaveData.startDate, leaveData.endDate, leaveData.location, leaveData.hasLeavePermit, activeWorkshop?.id]);

  // Keep the printed "Chế độ phép năm" in step with the year the server charged the leave to
  useEffect(() => {
    if (leavePlan?.leaveYear && leavePlan.leaveYear !== leaveData.leaveYear) {
      setLeaveData(prev => ({ ...prev, leaveYear: leavePlan.leaveYear }));
    }
  }, [leavePlan?.leaveYear]);

  // Leave request queues (backed by Supabase, not Google Sheets despite the name)
  const [waitingLeaves, setWaitingLeaves] = useState<any[]>([]);
  const [allLeaves, setAllLeaves] = useState<any[]>([]);
  const [isLoadingWaitingLeaves, setIsLoadingWaitingLeaves] = useState(false);
  const [selectedWaitingLeaveIds, setSelectedWaitingLeaveIds] = useState<string[]>([]);
  const [isSavingLeaveToSheets, setIsSavingLeaveToSheets] = useState(false);

  // "Bảng duyệt nghỉ phép" report: leaves already scheduled (status != Chờ phân ca)
  const [approvedSearch, setApprovedSearch] = useState('');
  const [approvedYearFilter, setApprovedYearFilter] = useState('all');
  const approvedLeaves = useMemo(() => {
    const processed = allLeaves.filter(l => l.status !== 'Chờ phân ca');
    const q = approvedSearch.trim().toLowerCase();
    return processed.filter(l => {
      const matchesSearch = !q || l.name.toLowerCase().includes(q) || (l.chucDanh || '').toLowerCase().includes(q);
      const matchesYear = approvedYearFilter === 'all' || l.leaveYear === approvedYearFilter;
      return matchesSearch && matchesYear;
    });
  }, [allLeaves, approvedSearch, approvedYearFilter]);
  const approvedYearOptions = useMemo(() => {
    const years = new Set(allLeaves.filter(l => l.status !== 'Chờ phân ca').map(l => l.leaveYear).filter(Boolean));
    return Array.from(years).sort();
  }, [allLeaves]);
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
  const [activeTab, setActiveTab] = useState<'schedule' | 'leave' | 'swap' | 'lookup' | 'staff' | 'auth'>('schedule');

  // A super admin's nav only offers "Tài khoản"; the schedule/leave/swap/staff tabs are
  // hidden for them. Without this they would sit on the default 'schedule' tab with no
  // button highlighted, seeing content the nav claims does not exist.
  useEffect(() => {
    if (currentUser?.role === 'super_admin' && activeTab !== 'auth') {
      setActiveTab('auth');
    }
  }, [currentUser?.role, activeTab]);

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


  // WorkshopManagerModal writes the same four fields this component keeps in `config`,
  // and it saves under the *same* workshop id. Keying the re-sync below on the id alone
  // would leave `config`/`staffData` holding pre-modal values, which the debounced save
  // would then write straight back over the modal's edits. Key on the content instead so
  // any change coming back from the server is picked up.
  const wsConfigKey = activeWorkshop
    ? JSON.stringify([
        activeWorkshop.config?.soVanBan,
        activeWorkshop.config?.ngayKy,
        activeWorkshop.config?.nguoiKy,
        activeWorkshop.config?.zaloWebhookUrl,
        activeWorkshop.config?.notifyEmail
      ])
    : '';
  const wsStaffKey = activeWorkshop ? JSON.stringify(activeWorkshop.staffData || []) : '';

  useEffect(() => {
    if (!activeWorkshop?.config) return;
    setConfig(prev => ({
      ...prev,
      soVanBan: activeWorkshop.config.soVanBan ?? '',
      ngayKy: activeWorkshop.config.ngayKy || '',
      nguoiKy: activeWorkshop.config.nguoiKy || 'Nguyễn Văn Nghị',
      zaloWebhookUrl: activeWorkshop.config.zaloWebhookUrl || 'https://vhialy.dpdns.org/webhook/notify',
      notifyEmail: activeWorkshop.config.notifyEmail || ''
    }));
  }, [activeWorkshop?.id, wsConfigKey]);

  // `config` only ever carries the 5 fields editable from the leave-request form
  // (soVanBan/ngayKy/nguoiKy/zaloWebhookUrl/notifyEmail) — it never gained
  // companyName/headerWorkshopName/documentCodeSuffix/recipientWorkshopName/
  // shortWorkshopName, the fields that actually make one workshop's exported Word
  // documents look different from another's. Every export call used to pass bare
  // `config`, so wordExport.ts's internal fallback defaults (hardcoded to Ialy) fired
  // for every workshop — Ialy happened to match them, so this only became visible once
  // a second workshop was created. Word exports should use this merged object instead.
  const docConfig = { ...activeWorkshop?.config, ...config };

  useEffect(() => {
    if (!activeWorkshop) return;
    // A freshly created workshop has no staff yet. Bailing out here would leave the
    // previous workshop's roster in state, which the debounced save would then write
    // into the new workshop — so clear it instead and let the admin fill it in.
    if (!Array.isArray(activeWorkshop.staffData) || activeWorkshop.staffData.length === 0) {
      setStaffData([]);
      localStorage.setItem('sd', JSON.stringify([]));
      setHydratedWsId(activeWorkshop.id);
      return;
    }
    const migrated = activeWorkshop.staffData.map((row: any) => {
      if (Array.isArray(row) && row[0] === 'Trực phụ cơ MR') {
        return ['Trực phụ máy MR', ...row.slice(1)];
      }
      return row;
    });
    setStaffData(migrated);
    localStorage.setItem('sd', JSON.stringify(migrated));
    setHydratedWsId(activeWorkshop.id);
  }, [activeWorkshop?.id, wsStaffKey]);

  // Signatures are per workshop, so they only need refetching when the workshop changes.
  useEffect(() => {
    if (!activeWorkshop) return;
    fetch(`/api/signatures?workshopId=${encodeURIComponent(activeWorkshop.id)}`)
      .then(res => res.json())
      .then(data => setSignatures(data))
      .catch(e => console.error("Failed to fetch signatures", e))
      .finally(() => setIsSettingsLoaded(true));
  }, [activeWorkshop?.id]);

  // Save staff list / config back onto the active workshop's row (debounced), instead
  // of the old single global app_config.
  useEffect(() => {
    if (!isSettingsLoaded || !activeWorkshop) return;
    // Never write state that belongs to a different workshop (or to no workshop yet).
    if (hydratedWsId !== activeWorkshop.id) return;

    const saveSettings = async () => {
      try {
        await fetch(API_BASE + '/api/workshops', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            id: activeWorkshop.id,
            name: activeWorkshop.name,
            code: activeWorkshop.code,
            description: activeWorkshop.description,
            features: activeWorkshop.features,
            staffData,
            config: { ...activeWorkshop.config, ...config }
          })
        });
      } catch (e) {
        console.error("Failed to save workshop settings", e);
      }
    };

    const timer = setTimeout(saveSettings, 3000);
    return () => clearTimeout(timer);
  }, [staffData, config, isSettingsLoaded, activeWorkshop, hydratedWsId]);

  const fetchWaitingLeaves = useCallback(async () => {
    if (!activeWorkshop) return;
    setIsLoadingWaitingLeaves(true);
    try {
      const res = await fetch(`/api/sheets/leave-requests?workshopId=${encodeURIComponent(activeWorkshop.id)}`);
      if (res.ok) {
        const data = await res.json();
        const mappedData = data.map((item: any) => ({
          ...item,
          chucDanh: item.chucDanh === 'Trực phụ cơ MR' ? 'Trực phụ máy MR' : item.chucDanh
        }));
        setAllLeaves(mappedData);
        const waiting = mappedData.filter((item: any) => item.status === 'Chờ phân ca');
        setWaitingLeaves(waiting);
        // Clear selected if they no longer exist in waiting
        setSelectedWaitingLeaveIds(prev => prev.filter(id => waiting.some((w: any) => w.id === id)));
      } else {
        console.error("Failed to fetch leave requests");
      }
    } catch (e) {
      console.error("Error fetching live leaves request", e);
    } finally {
      setIsLoadingWaitingLeaves(false);
    }
  }, [activeWorkshop]);

  useEffect(() => {
    fetchWaitingLeaves();
  }, [fetchWaitingLeaves]);

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
    await exportWord(currentResult, docConfig);
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
    const blob = await generateWordBlob(currentResult, docConfig);
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
    const blob = await generateSwapBlob(swapData, docConfig, signatures);
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
    const blob = await generateLeaveRequestBlob(leaveData, docConfig, signatures);
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
    // Saving to the waiting-schedule queue happens automatically before the download,
    // so there's no separate "Lưu lên hệ thống" step for the user to remember.
    await saveLeaveToGoogleSheets();
    await exportLeaveRequestDoc(leaveData, docConfig, signatures);
    setIsProcessing(false);
  };

  const saveLeaveToGoogleSheets = async () => {
    if (!leaveData.name) {
      setAlert("⚠ Vui lòng nhập Họ và tên người xin nghỉ.");
      return;
    }
    if (!activeWorkshop) {
      setAlert("⚠ Chưa chọn phân xưởng.");
      return;
    }
    setIsSavingLeaveToSheets(true);
    try {
      const payload = {
        ...leaveData,
        workshopId: activeWorkshop.id,
        leaveBalance: leaveBalance ? {
          entitled: leaveBalance.entitled,
          used: leaveBalance.used,
          remaining: leaveBalance.remaining
        } : null
      };
      const res = await fetch(API_BASE + '/api/sheets/leave-requests', {
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
        setAlert(`❌ Lưu đơn thất bại: ${resData.error || "Lỗi máy chủ"}`);
      }
    } catch (e: any) {
      setAlert(`❌ Đã xảy ra lỗi kết nối: ${e.message}`);
    } finally {
      setIsSavingLeaveToSheets(false);
    }
  };

  const handleExportApprovedLeavesExcel = () => {
    if (approvedLeaves.length === 0) {
      setAlert('⚠ Không có đơn nào trong danh sách để xuất Excel.');
      return;
    }

    const rows = approvedLeaves.map((l, idx) => ({
      'STT': idx + 1,
      'Họ và tên': l.name,
      'Chức danh': l.chucDanh,
      'Kíp': l.kip,
      'Từ ngày': l.startDate ? fmtVN(new Date(l.startDate)) : '',
      'Đến ngày': l.endDate ? fmtVN(new Date(l.endDate)) : '',
      'Số ngày phép trừ': l.leaveDays ?? '',
      'Phân bổ theo năm': Array.isArray(l.allocations) && l.allocations.length > 0
        ? l.allocations.map((a: any) => `${a.days} ngày (${a.year})`).join(', ')
        : '',
      'Có lấy giấy phép': l.hasLeavePermit ? 'Có' : 'Không',
      'Ngày đi đường': l.travelDays || 0,
      'Lý do': l.reason,
      'Địa điểm': l.location,
      'Điện thoại': l.phone,
      'Trạng thái': l.status,
      'Ngày tạo đơn': l.createdAt
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet['!cols'] = [
      { wch: 5 }, { wch: 22 }, { wch: 16 }, { wch: 6 }, { wch: 11 }, { wch: 11 },
      { wch: 15 }, { wch: 20 }, { wch: 14 }, { wch: 12 }, { wch: 28 }, { wch: 12 },
      { wch: 13 }, { wch: 12 }, { wch: 18 }
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Bang_duyet_nghi_phep');

    const stamp = fmtVN(new Date()).replace(/\//g, '-');
    XLSX.writeFile(workbook, `Bang_duyet_nghi_phep_${stamp}.xlsx`);
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
    if (ids.length === 0) return;

    const message = isSingleName
      ? `Bạn có chắc chắn muốn xóa đơn nghỉ phép của anh/chị "${isSingleName}" không?`
      : `Bạn có chắc chắn muốn danh sách ${ids.length} đơn nghỉ phép đã chọn bị xóa vĩnh viễn?`;

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
          const res = await fetch(API_BASE + '/api/sheets/leave-requests/delete', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({ ids, workshopId: activeWorkshop?.id })
          });

          if (res.ok) {
            setAlert(isSingleName 
              ? `✅ Đã xóa đơn của "${isSingleName}" thành công!`
              : `✅ Đã xóa thành công ${ids.length} đơn nghỉ phép!`
            );
            setSelectedWaitingLeaveIds(prev => prev.filter(id => !ids.includes(id)));
            fetchWaitingLeaves();
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

  const proceedExportAllZipAndUpdateStatus = async () => {
    if (!currentResult) return;
    setIsProcessing(true);
    try {
      const selectedLeavesData = waitingLeaves.filter(item => selectedWaitingLeaveIds.includes(item.id));

      // 1. Download the ZIP file
      await exportAllDocsZip(currentResult, selectedLeavesData, docConfig, signatures);

      // 2. Update waiting leaves status to 'Đã xử lý'
      if (selectedWaitingLeaveIds.length > 0) {
        const updateRes = await fetch(API_BASE + '/api/sheets/leave-requests/update-status', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ids: selectedWaitingLeaveIds,
            status: "Đã xử lý",
            workshopId: activeWorkshop?.id
          })
        });

        if (updateRes.ok) {
          setAlert("✅ Đã xuất trọn bộ hồ sơ dạng ZIP và cập nhật trạng thái 'Đã xử lý'!");
          fetchWaitingLeaves();
        } else {
          const errData = await updateRes.json();
          setAlert(`✅ Xuất hồ sơ thành công nhưng lỗi cập nhật trạng thái: ${errData.error || ""}`);
        }
      } else {
        setAlert("✅ Đã tải xuống hồ sơ dạng ZIP thành công!");
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
    await proceedExportAllZipAndUpdateStatus();
  };

  const handleExportSwap = async () => {
    setIsProcessing(true);
    await exportSwapDoc(swapData, docConfig, signatures);
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

  // Everything below requires a logged-in account.
  if (!currentUser) {
    return <LoginForm onLoginSuccess={handleLoginSuccess} />;
  }

  const isSuperAdmin = currentUser.role === 'super_admin';
  const isWorkshopAdmin = currentUser.role === 'workshop_admin';
  const isAdmin = isSuperAdmin || isWorkshopAdmin;

  return (
    <div className="min-h-screen flex flex-col bg-[var(--bg)]">
      {showWorkshopManager && (
        <WorkshopManagerModal
          user={currentUser}
          workshops={workshops}
          activeWorkshop={activeWorkshop}
          onClose={() => setShowWorkshopManager(false)}
          onRefreshWorkshops={fetchWorkshops}
        />
      )}
      {/* Top Navbar */}
      <nav id="top-navbar" className="w-full bg-white border-b border-slate-200 py-2 px-3 md:py-3.5 md:px-8 shadow-md flex flex-col lg:flex-row items-center justify-between gap-2 lg:gap-4 sticky top-0 z-40">
        <div className="flex flex-row items-center gap-2.5 sm:gap-4 w-full lg:w-auto">
          {/* Logo container */}
          <div className="flex items-center select-none">
            <img
              src="https://i.ibb.co/r2BTySkt/2343.png"
              alt="EVN Công Ty Thủy Điện Ialy Logo"
              className="h-[32px] sm:h-[48px] md:h-[54px] w-auto object-contain shrink-0"
              referrerPolicy="no-referrer"
            />
          </div>
          {/* App Title with separator */}
          <div className="border-l border-slate-200 pl-2.5 sm:pl-4 py-0.5 flex flex-col justify-center h-full text-left min-w-0">
            {/* Deliberately not `truncate`, which is overflow:hidden and so clips both axes.
                These lines pair a heavy weight with leading-tight, and the resulting line box
                is shorter than the ink: measured clearance above the text was 1px at 16px and
                -0.5px at the 12px mobile size. That shaves the top off tall diacritics — the
                hook on Ủ lost its loop and the leftover tail read as an acute, turning
                "THỦY" into "THUÝ". Clipping only the inline axis keeps the ellipsis
                behaviour while letting marks overflow upward. */}
            <div className="text-[#053d6c] font-black text-[12px] sm:text-base md:text-xl font-sans tracking-tight leading-tight mt-0.5 whitespace-nowrap text-ellipsis overflow-x-clip">
              {activeWorkshop?.config?.companyName || 'CÔNG TY THỦY ĐIỆN IALY'}
            </div>
            <div className="text-[#053d6c] font-black text-[12px] sm:text-base md:text-xl font-sans tracking-tight leading-tight mt-0.5 whitespace-nowrap text-ellipsis overflow-x-clip">
              {activeWorkshop ? activeWorkshop.name : 'Phân xưởng Vận hành Ialy'}
            </div>
          </div>
        </div>

        {/* Navigation Tabs on the Right */}
        <div className="nav-tabs flex items-center gap-4 md:gap-7 flex-nowrap overflow-x-auto w-full lg:w-auto lg:flex-wrap justify-start lg:justify-center">
          {!isSuperAdmin && (
            <>
              <button
                id="nav-schedule-btn"
                className={`text-[12px] md:text-[15px] whitespace-nowrap shrink-0 font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
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
                className={`text-[12px] md:text-[15px] whitespace-nowrap shrink-0 font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
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
                className={`text-[12px] md:text-[15px] whitespace-nowrap shrink-0 font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
                  activeTab === 'swap'
                    ? 'text-[#00529c] border-[#00529c] font-extrabold'
                    : 'text-slate-500 hover:text-[#00529c] border-transparent hover:border-[#00529c]/50'
                }`}
                onClick={() => setActiveTab('swap')}
              >
                Đơn đổi ca
              </button>
              {!isWorkshopAdmin && (
                <button
                  id="nav-lookup-btn"
                  className={`text-[12px] md:text-[15px] whitespace-nowrap shrink-0 font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
                    activeTab === 'lookup'
                      ? 'text-[#00529c] border-[#00529c] font-extrabold'
                      : 'text-slate-500 hover:text-[#00529c] border-transparent hover:border-[#00529c]/50'
                  }`}
                  onClick={() => setActiveTab('lookup')}
                >
                  Tra cứu thông tin
                </button>
              )}
            </>
          )}

          {isWorkshopAdmin && (
            <button
              id="nav-staff-btn"
              className={`text-[12px] md:text-[15px] whitespace-nowrap shrink-0 font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
                activeTab === 'staff'
                  ? 'text-[#00529c] border-[#00529c] font-extrabold'
                  : 'text-slate-500 hover:text-[#00529c] border-transparent hover:border-[#00529c]/50'
              }`}
              onClick={() => setActiveTab('staff')}
            >
              Nhân sự
            </button>
          )}

          <button
            id="nav-auth-btn"
            className={`text-[12px] md:text-[15px] whitespace-nowrap shrink-0 font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
              activeTab === 'auth'
                ? 'text-[#00529c] border-[#00529c] font-extrabold'
                : 'text-slate-500 hover:text-[#00529c] border-transparent hover:border-[#00529c]/50'
            }`}
            onClick={() => setActiveTab('auth')}
          >
            Tài khoản
          </button>
        </div>
      </nav>

      {/* Hero Banner with beautiful high resolution background */}
      <div 
        id="hero-banner" 
        className="w-full h-[120px] sm:h-[220px] md:h-[280px] lg:h-[320px] relative bg-cover bg-center flex items-center justify-center overflow-hidden" 
        style={{ 
          backgroundImage: "url('https://i.ibb.co/jZ6dDJzT/z7116558150434-802a4bd8dff3b332930235031b93fc49.jpg')",
          backgroundPosition: "center 42%"
        }}
      >
        <div className="absolute inset-0 bg-slate-950/20"></div>
        
      </div>

      {/* Container for Main Content */}
      <div className="wrap !pt-8 !pb-16 flex-1 w-full max-w-[1200px] mx-auto px-4 md:px-6">

      {/* Shared notification: lives outside the tab blocks so every tab can report back */}
      {alert && <div className={`alert ${alert.startsWith('✅') ? 'asuc' : 'aerr'}`}>{alert}</div>}

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
              list="leave-location-list"
              value={leaveData.location}
              onChange={e => setLeaveData({...leaveData, location: e.target.value})}
            />
            <datalist id="leave-location-list">
              {locationOptions.map(loc => (
                <option key={loc.name} value={loc.name}>{loc.distanceKm} km</option>
              ))}
            </datalist>
          </div>

          <div className="field col-span-2">
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={leaveData.hasLeavePermit}
                onChange={e => setLeaveData({...leaveData, hasLeavePermit: e.target.checked})}
                className="w-4 h-4 cursor-pointer"
              />
              Có lấy giấy phép (được tính ngày đi đường)
            </label>
          </div>

          {/* Travel-day allowance feedback (kept visible; the shift/deduction breakdown above it is hidden) */}
          {leavePlan && leaveData.hasLeavePermit && (
            <div className="col-span-2 p-3 bg-slate-50 border border-slate-100 rounded-xl text-xs">
              <div className={leavePlan.travel.travelDays > 0 ? 'text-emerald-800' : 'text-amber-700'}>
                {leavePlan.travel.travelDays > 0
                  ? <>🚗 Pleiku → {leaveData.location}: <b>{leavePlan.travel.distanceKm} km</b> — cộng thêm <b>{leavePlan.travel.travelDays} ngày đi đường</b> vào quỹ phép năm {leavePlan.leaveYear}.</>
                  : <>⚠ {leavePlan.travel.note}</>}
              </div>
            </div>
          )}

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
                  <div className="text-[11px] text-blue-700 font-medium uppercase">Phép được hưởng</div>
                  <div className="text-lg font-extrabold text-blue-900 mt-0.5">{leaveBalance.entitled} <span className="text-xs font-normal">ngày</span></div>
                </div>
                <div className="bg-amber-50/60 p-2.5 rounded-lg border border-amber-100 text-center">
                  <div className="text-[11px] text-amber-700 font-medium uppercase">Phép đã nghỉ</div>
                  <div className="text-lg font-extrabold text-amber-900 mt-0.5">{leaveBalance.used} <span className="text-xs font-normal">ngày</span></div>
                </div>
                <div className="bg-emerald-50/60 p-2.5 rounded-lg border border-emerald-100 text-center">
                  <div className="text-[11px] text-emerald-700 font-medium uppercase">Phép còn lại</div>
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
          <button
            className="btn btn-secondary flex-1 min-w-[150px]"
            onClick={handleExportLeave}
            disabled={isProcessing || isSavingLeaveToSheets || !leaveData.name}
          >
            {(isProcessing || isSavingLeaveToSheets) ? <span className="spin mr-2"></span> : '📥'}
            {isSavingLeaveToSheets ? ' Đang lưu đơn...' : ' Lưu đơn & Xuất File Word'}
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

        {(
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
                {waitingLeaves.length > 0 && (
                  <button
                    className="px-2.5 py-1 text-xs bg-white text-var(--acc) border border-var(--acc-light) rounded hover:bg-var(--acc-light-10) font-medium cursor-pointer transition-colors"
                    onClick={() => {
                      // Toggle: once everything is already selected the same button clears
                      // the selection, so there is no need for a second "bỏ chọn" control.
                      const allIds = waitingLeaves.map(l => l.id);
                      const allSelected = allIds.length > 0 && allIds.every(id => selectedWaitingLeaveIds.includes(id));
                      setSelectedWaitingLeaveIds(allSelected ? [] : allIds);
                    }}
                  >
                    {waitingLeaves.length > 0 && waitingLeaves.every(l => selectedWaitingLeaveIds.includes(l.id))
                      ? '☐ Bỏ chọn tất cả'
                      : `☑ Chọn tất cả (${waitingLeaves.length})`}
                  </button>
                )}
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
              <p className="text-[12px] text-var(--txt2) italic">Không có đơn xin nghỉ phép nào ở trạng thái Chờ phân ca.</p>
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
                              <span className="text-[11px] bg-yellow-100 text-yellow-800 border border-yellow-200 px-1.5 py-0.5 rounded-md font-mono font-medium">{leave.id.substring(0, 14)}...</span>
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
            <p className="text-[11px] text-var(--txt2) mt-2 italic">* Chọn các đơn muốn xếp lịch và bấm "Áp dụng ... người" để tự động điền nhanh các thông tin.</p>
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

      {activeTab === 'lookup' && (
        <div className="card w-full p-0 overflow-hidden min-h-[750px] shadow-sm border border-slate-200 rounded-2xl bg-white" id="info-lookup-card">
          <div className="p-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="text-xl">🔍</span>
              <h2 className="font-bold text-slate-800 text-base">Tra cứu thông tin & Thỏa ước lao động tập thể</h2>
            </div>
            <span className="text-xs text-slate-500 font-medium bg-slate-200/80 px-2.5 py-1 rounded-full">Trợ lý AI Trực tuyến</span>
          </div>
          <div className="w-full h-[720px]">
            <iframe
              src="https://dothanhphongpc1.tail1ac872.ts.net/chatbot/PMzknllbIHSxw2FT"
              className="w-full h-full border-0"
              title="Tra cứu thông tin"
              allow="microphone; clipboard-read; clipboard-write"
            />
          </div>
        </div>
      )}

      {activeTab === 'auth' && (
        <div className="w-full max-w-2xl mx-auto space-y-6 my-6 px-4">
          <div className="card w-full p-6 bg-white border border-slate-200 rounded-2xl shadow-md" id="account-info-card">
            <div className="flex items-center justify-between pb-4 border-b border-slate-100 flex-wrap gap-3">
              <div className="flex items-center gap-3">
                <div className="w-12 h-12 rounded-2xl bg-teal-600 text-white flex items-center justify-center font-bold text-lg shadow-md">
                  <User size={22} />
                </div>
                <div>
                  <h2 className="text-lg font-black text-slate-800">{currentUser.fullName}</h2>
                  <div className="flex items-center gap-2 mt-0.5">
                    <span className="text-xs text-slate-500 font-mono">@{currentUser.username}</span>
                    <span className={`text-[11px] font-extrabold px-2 py-0.5 rounded-md ${
                      isSuperAdmin
                        ? 'bg-amber-100 text-amber-800 border border-amber-300'
                        : isWorkshopAdmin
                          ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                          : 'bg-slate-100 text-slate-700 border border-slate-200'
                    }`}>
                      {isSuperAdmin ? 'Super Admin (Admin Tổng)' : isWorkshopAdmin ? 'Admin Phân Xưởng' : 'Tài Khoản Phân Xưởng'}
                    </span>
                  </div>
                </div>
              </div>

              <button
                onClick={handleLogout}
                className="flex items-center gap-2 px-4 py-2 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 rounded-xl text-xs font-bold transition-all cursor-pointer shadow-sm active:scale-95"
              >
                <LogOut size={14} /> Đăng xuất
              </button>
            </div>

            <div className="mt-6 space-y-4">
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-200/80">
                <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">
                  Phân xưởng hoạt động
                </label>
                {workshops.length > 1 && (isSuperAdmin || currentUser.workshopId === 'all') ? (
                  <select
                    value={activeWorkshop?.id || ''}
                    onChange={(e) => {
                      const found = workshops.find(w => w.id === e.target.value);
                      if (found) setActiveWorkshop(found);
                    }}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-sm font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-sky-600 cursor-pointer"
                  >
                    {workshops.map(ws => (
                      <option key={ws.id} value={ws.id}>{ws.name} ({ws.code})</option>
                    ))}
                  </select>
                ) : (
                  <div className="font-bold text-sky-800 text-sm bg-sky-50 px-3 py-2 rounded-xl border border-sky-200">
                    {activeWorkshop ? activeWorkshop.name : 'Chưa chọn phân xưởng'}
                  </div>
                )}
              </div>

              {isAdmin && (
                <button
                  onClick={() => setShowWorkshopManager(true)}
                  className="w-full py-3 px-4 bg-amber-500 hover:bg-amber-400 text-slate-950 text-sm font-black rounded-xl shadow-md transition-all cursor-pointer flex items-center justify-center gap-2 active:scale-98"
                >
                  <Settings size={18} />
                  <span>Quản lý Phân xưởng & Tài khoản</span>
                </button>
              )}
            </div>
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
            
            {showSignatureManager && activeWorkshop && (
              <div className="mb-6">
                <SignatureManager
                  staffList={Array.from(new Set(staffData.flatMap(row => row.slice(1)).filter(Boolean)))}
                  signatures={signatures}
                  onSignaturesChange={setSignatures}
                  workshopId={activeWorkshop.id}
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

          <div className="card" id="leave-balance-card">
            <div className="ctitle">
              Phép năm của nhân viên
              <button className="staff-toggle" onClick={() => setShowLeaveBalance(!showLeaveBalance)}>
                {showLeaveBalance ? 'Ẩn ▲' : 'Hiện ▼'}
              </button>
            </div>
            {showLeaveBalance && activeWorkshop && (
              <LeaveBalanceManager
                staffList={Array.from(new Set(staffData.flatMap(row => row.slice(1)).filter(Boolean)))}
                onAlert={setAlert}
                workshopId={activeWorkshop.id}
              />
            )}
          </div>

          <div className="card" id="approved-leaves-card">
            <div className="ctitle">
              📋 Bảng duyệt nghỉ phép — đã xếp lịch đi ca ({approvedLeaves.length})
              <div className="flex gap-2">
                <button
                  className="px-2 py-1 text-xs bg-white text-sky-700 border border-sky-300 rounded hover:bg-sky-100 font-medium cursor-pointer transition-colors"
                  onClick={fetchWaitingLeaves}
                  disabled={isLoadingWaitingLeaves}
                >
                  🔄 Cập nhật
                </button>
                <button
                  className="px-2.5 py-1 text-xs bg-sky-700 text-white rounded hover:bg-sky-800 font-bold cursor-pointer transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                  onClick={handleExportApprovedLeavesExcel}
                  disabled={approvedLeaves.length === 0}
                >
                  📊 Xuất Excel
                </button>
              </div>
            </div>

            <div className="flex gap-2 mb-3 mt-3 flex-wrap">
              <input
                type="text"
                placeholder="Tìm theo tên hoặc chức danh..."
                value={approvedSearch}
                onChange={e => setApprovedSearch(e.target.value)}
                className="flex-1 min-w-[180px] px-2.5 py-1.5 text-xs border border-slate-200 rounded-lg"
              />
              <select
                value={approvedYearFilter}
                onChange={e => setApprovedYearFilter(e.target.value)}
                className="px-2.5 py-1.5 text-xs border border-slate-200 rounded-lg"
              >
                <option value="all">Tất cả năm phép</option>
                {approvedYearOptions.map(y => <option key={y} value={y}>{y}</option>)}
              </select>
            </div>

            {approvedLeaves.length === 0 ? (
              <p className="text-[12px] text-var(--txt2) italic">Chưa có đơn nào được xếp lịch đi ca.</p>
            ) : (
              <div className="max-h-[280px] overflow-auto rounded-lg border border-sky-100">
                <table className="w-full text-[12px] border-collapse">
                  <thead className="sticky top-0 bg-sky-100 text-sky-900">
                    <tr>
                      <th className="text-left p-2 font-bold">Họ và tên</th>
                      <th className="text-left p-2 font-bold">Chức danh</th>
                      <th className="p-2 font-bold">Kíp</th>
                      <th className="text-left p-2 font-bold">Thời gian</th>
                      <th className="p-2 font-bold">Số ngày phép</th>
                      <th className="text-left p-2 font-bold">Năm phép</th>
                      <th className="text-left p-2 font-bold">Trạng thái</th>
                    </tr>
                  </thead>
                  <tbody>
                    {approvedLeaves.map(l => (
                      <tr key={l.id} className="border-t border-sky-100 bg-white hover:bg-sky-50/60">
                        <td className="p-2 font-semibold text-slate-800">{l.name}</td>
                        <td className="p-2 text-slate-600">{l.chucDanh}</td>
                        <td className="p-2 text-center text-slate-600">{l.kip}</td>
                        <td className="p-2 text-slate-600 whitespace-nowrap">
                          {l.startDate ? fmtVN(new Date(l.startDate)) : ''} → {l.endDate ? fmtVN(new Date(l.endDate)) : ''}
                        </td>
                        <td className="p-2 text-center text-slate-600">{l.leaveDays ?? '—'}</td>
                        <td className="p-2 text-slate-600">
                          {Array.isArray(l.allocations) && l.allocations.length > 0
                            ? l.allocations.map((a: any) => `${a.days} (${a.year})`).join(', ')
                            : '—'}
                        </td>
                        <td className="p-2">
                          <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded-md text-[11px] font-semibold">{l.status}</span>
                        </td>
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
                  className="font-mono text-[12px] w-full px-3 py-2 border rounded-md"
                />
                <span className="text-[11px] text-var(--txt2) italic mt-1.5 block">
                  * Hệ thống sẽ tự động gửi thông báo về Zalo cá nhân qua đường dẫn webhook này khi có đơn nghỉ phép mới được tạo hoặc lưu lên hệ thống chờ phân ca.
                </span>
              </div>
              <div className="field col-span-2">
                <label>Email nhận thông báo</label>
                <input
                  type="text"
                  value={config.notifyEmail || ''}
                  onChange={e => setConfig({ ...config, notifyEmail: e.target.value })}
                  placeholder="vd: quandoc@gmail.com (nhiều địa chỉ cách nhau bằng dấu phẩy)"
                  className="font-mono text-[12px] w-full px-3 py-2 border rounded-md"
                />
                <span className="text-[11px] text-var(--txt2) italic mt-1.5 block">
                  * Ngoài Zalo, hệ thống cũng gửi email thông báo tới (các) địa chỉ này khi có đơn nghỉ phép mới.
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
                {selectedWaitingLeaveIds.length > 0 && (
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
                                  <div className="text-[11px] text-var(--acc) font-bold">👥 Ngày trùng nghỉ</div>
                                )}
                              </>
                            )}
                          </td>
                          <td><span className={`badge b${row.ca}`}>{row.ca}</span></td>
                          <td className="text-[12px] text-var(--txt2)">
                            <div className="font-bold text-[11px] text-var(--acc) mb-0.5 uppercase">{row.chucDanh}</div>
                            {row.isSwap ? (
                              <>
                                <span className="text-[#22c55e] text-[11px]">đổi ca</span>
                                <br />
                                {row.relievedTen || row.absentTen}
                                <br />
                                <span className="text-[11px]">Kíp {row.relievedKip || row.absentKip}</span>
                              </>
                            ) : row.isCKChain ? (
                              <>
                                <span className="text-[#a855f7] text-[11px]">thay ca</span>
                                <br />
                                {row.absentTen}
                                <br />
                                <span className="text-[11px]">Kíp {row.absentKip}</span>
                              </>
                            ) : (
                              <>
                                {row.absentTen}
                                <br />
                                <span className="text-[11px]">Kíp {row.absentKip}</span>
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
                                  <div className="text-[11px] text-var(--txt2)">{row.conflictNote}</div>
                                )}
                              </>
                            ) : row.isCKChain ? (
                              <>
                                <span className="conflict-badge bg-[#a855f71a] text-[#a855f7] border-[#a855f74d]">⥵ C→K</span>
                                <br />
                                <span className="text-[11px] text-var(--txt2)">Thay do ràng buộc Ca C→K</span>
                              </>
                            ) : row.isConflict ? (
                              <>
                                <span className="conflict-badge">△ Điều chỉnh</span>
                                {row.conflictNote && (
                                  <div className="text-[11px] text-var(--txt2)">{row.conflictNote}</div>
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
                <button
                  className="btn btn-primary flex items-center justify-center gap-1.5"
                  onClick={handleExportLeave}
                  disabled={isProcessing || isSavingLeaveToSheets || !leaveData.name}
                >
                  {(isProcessing || isSavingLeaveToSheets) ? <span className="spin spinw mr-2"></span> : '📥 '}
                  {isSavingLeaveToSheets ? 'Đang lưu đơn...' : 'Lưu đơn & Tải xuống Word'}
                </button>
              ) : (
                // Downloads the shift schedule on its own. Bundling the leave request
                // documents alongside it lives on the dedicated "Tạo lịch thay ca và các
                // đơn nghỉ phép" button instead, because that one also flips the selected
                // requests to "Đã xử lý" — a side effect nobody expects from a preview.
                <button
                  className="btn btn-primary flex items-center justify-center gap-1.5"
                  onClick={async () => {
                    await handleExportWord();
                    setShowPreview(false);
                  }}
                  disabled={isProcessing}
                >
                  {isProcessing ? <span className="spin spinw mr-2"></span> : '📥'} Tải xuống Word
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

    {/* Floating back-to-top. Sits below the sticky nav (z-40) and well below the modals
        (z-50 / 9999) so it can never cover a dialog the user is working in. */}
    <button
      type="button"
      onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}
      aria-label="Về đầu trang"
      title="Về đầu trang"
      className={`fixed bottom-5 right-5 z-30 h-11 w-11 rounded-full bg-[#00529c] text-white shadow-lg border border-white/20
        flex items-center justify-center hover:bg-[#0a7bd4] active:scale-95 cursor-pointer
        transition-all duration-200 ${
          showScrollTop ? 'opacity-100 translate-y-0 pointer-events-auto' : 'opacity-0 translate-y-3 pointer-events-none'
        }`}
    >
      <span className="text-[19px] leading-none font-bold">↑</span>
    </button>

    {/* Three columns of content over a separated legal bar. The previous version put
        everything on one line, so the system name, the parent-company note and the
        policy links all carried the same weight and none of them led. */}
    <footer id="app-footer" className="w-full bg-[#062240] text-slate-300 mt-auto select-none">
      {/* A single hairline in the brand blue is the only ornament the footer needs. */}
      <div className="h-[3px] w-full bg-gradient-to-r from-[#00529c] via-[#0a7bd4] to-[#00529c]"></div>

      <div className="max-w-[1200px] mx-auto px-6 md:px-12 py-10">
        <div className="grid grid-cols-1 md:grid-cols-12 gap-8 md:gap-10">

          {/* Brand */}
          <div className="md:col-span-6">
            <div className="flex items-start gap-3.5">
              {/* The mark sits on a white chip rather than being colour-filtered: the
                  source PNG is not transparent, so inverting it produced a solid block. */}
              <div className="shrink-0 h-11 w-11 rounded-lg bg-white flex items-center justify-center shadow-sm">
                <img
                  src="https://i.ibb.co/r2BTySkt/2343.png"
                  alt="EVN"
                  className="h-8 w-8 object-contain"
                  referrerPolicy="no-referrer"
                />
              </div>
              <div className="min-w-0">
                <div className="text-[15px] text-white font-bold tracking-tight leading-snug">
                  Hệ thống Lịch trực thay ca vận hành
                </div>
                <p className="text-[12px] text-slate-400 leading-relaxed mt-1.5 max-w-md">
                  Trang thông tin hỗ trợ nội bộ trực thuộc{' '}
                  <a
                    href="https://ialyhpc.vn"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-slate-200 underline decoration-slate-600 underline-offset-2 hover:text-white hover:decoration-[#0a7bd4] transition-colors"
                  >
                    Công ty Thủy điện Ialy
                  </a>
                  . Tự động lập lịch trực thay ca, quản lý đơn nghỉ phép và phép năm.
                </p>
              </div>
            </div>
          </div>

          {/* Links */}
          <div className="md:col-span-3">
            <div className="text-[11px] font-bold text-slate-500 uppercase tracking-[0.12em] mb-3.5">
              Liên kết
            </div>
            <ul className="space-y-2.5">
              {['Hướng dẫn sử dụng', 'Chính sách bảo mật nội bộ', 'Hỗ trợ kỹ thuật'].map(label => (
                <li key={label}>
                  <span className="text-[12px] text-slate-400 hover:text-white transition-colors cursor-pointer inline-flex items-center gap-2 group">
                    <span className="w-1 h-1 rounded-full bg-slate-600 group-hover:bg-[#0a7bd4] transition-colors"></span>
                    {label}
                  </span>
                </li>
              ))}
            </ul>
          </div>

          {/* Current context: which workshop this session is actually working on. */}
          <div className="md:col-span-3">
            <div className="text-[11px] font-bold text-slate-500 uppercase tracking-[0.12em] mb-3.5">
              Đơn vị sử dụng
            </div>
            <div className="text-[12px] text-slate-300 font-semibold leading-snug">
              {activeWorkshop?.name || 'Phân xưởng Vận hành'}
            </div>
            {activeWorkshop?.code && (
              <div className="mt-2 inline-block text-[11px] font-mono text-slate-400 bg-white/5 border border-white/10 rounded px-2 py-1">
                {activeWorkshop.code}
              </div>
            )}
            <div className="text-[11px] text-slate-500 mt-3 leading-relaxed">
              {activeWorkshop?.config?.companyName || 'Công ty Thủy điện Ialy'}
            </div>
          </div>
        </div>

        {/* Legal bar. The back-to-top control lives in the floating button instead —
            here it only appeared once you had already scrolled to the bottom. */}
        <div className="mt-9 pt-5 border-t border-white/10">
          <div className="text-[11px] text-slate-500 text-center sm:text-left">
            © {new Date().getFullYear()} Công ty Thủy điện Ialy — Tập đoàn Điện lực Việt Nam
          </div>
        </div>
      </div>
    </footer>
    </div>
  );
}

