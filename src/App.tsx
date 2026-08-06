import React, { useState, useEffect, useCallback, useRef } from 'react';
import mammoth from 'mammoth';
import { DEFAULT_STAFF, SHIFTS } from './constants';
import { fmtVN, fmtIn, dayN, timNghi, timThay, abbrev, xacDinhCa } from './utils/shiftHelpers';
import { buildMultiLeaveResults, Leave, ResultItem } from './utils/Quytacxacdinhcatructhay';
import { exportWord, generateWordBlob, exportSwapDoc, generateSwapBlob, exportLeaveRequestDoc, generateLeaveRequestBlob, exportAllDocsZip } from './utils/wordExport';
import { renderAsync } from 'docx-preview';
import SignatureManager from './components/SignatureManager';
import LoginForm from './components/LoginForm';
import UserHeaderBar from './components/UserHeaderBar';
import WorkshopManagerModal from './components/WorkshopManagerModal';
import StaffDataEditor from './components/StaffDataEditor';
import { UserAccount, Workshop } from './types/auth';
import { Trash2, Settings, RefreshCw } from 'lucide-react';

export default function App() {
  // Authentication & Workshop RBAC State
  const [user, setUser] = useState<UserAccount | null>(null);
  const isAdmin = user?.role === 'super_admin' || user?.role === 'workshop_admin';
  const [workshops, setWorkshops] = useState<Workshop[]>([]);
  const [activeWorkshop, setActiveWorkshop] = useState<Workshop | null>(null);
  const [showWorkshopManager, setShowWorkshopManager] = useState(false);
  const [authChecking, setAuthChecking] = useState(true);

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
  const [showSystemConfig, setShowSystemConfig] = useState(false);
  const [config, setConfig] = useState<{
    soVanBan: string;
    ngayKy: string;
    nguoiKy: string;
    chucVuNguoiKy?: string;
    zaloWebhookUrl: string;
    googleSheetUrl?: string;
    companyName?: string;
    headerWorkshopName?: string;
    documentCodeSuffix?: string;
    shortWorkshopName?: string;
    locationName?: string;
  }>({
    soVanBan: '',
    ngayKy: '',
    nguoiKy: 'Nguyễn Văn Nghị',
    chucVuNguoiKy: 'Quản Đốc',
    zaloWebhookUrl: 'https://vhialy.dpdns.org/webhook/notify',
    googleSheetUrl: 'https://docs.google.com/spreadsheets/d/1m8B-CVJEyzt5KhJ-I_1hFTvWczz1jl7PPm_7PNScYYQ/edit?gid=0#gid=0',
    companyName: 'CÔNG TY THỦY ĐIỆN IALY',
    headerWorkshopName: 'PHÂN XƯỞNG VẬN HÀNH IALY',
    documentCodeSuffix: '/VHIALY',
    shortWorkshopName: 'PXVH Ialy',
    locationName: 'Gia Lai'
  });

  const [isGoogleAuth, setIsGoogleAuth] = useState(false);
  const [loadingAuth, setLoadingAuth] = useState(true);
  const [isUpdatingSheets, setIsUpdatingSheets] = useState(false);
  const [signatures, setSignatures] = useState<Record<string, string>>({});
  const [showSignatureManager, setShowSignatureManager] = useState(false);
  const [isSettingsLoaded, setIsSettingsLoaded] = useState(false);
  const [isSavingConfig, setIsSavingConfig] = useState(false);
  const [configSaveMsg, setConfigSaveMsg] = useState<string | null>(null);

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
  const [leaveBalance, setLeaveBalance] = useState<{ entitled: string; used: string; remaining: string; note?: string; oldLeaves?: string; newLeaves?: string } | null>(null);
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
          remaining: data.remaining,
          note: data.note,
          oldLeaves: data.oldLeaves,
          newLeaves: data.newLeaves
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
  const [activeTab, setActiveTab] = useState<'schedule' | 'leave' | 'swap' | 'lookup' | 'auth' | 'staff'>('schedule');
  const [showTopBtn, setShowTopBtn] = useState(false);

  // Change password states
  const [showChangePwForm, setShowChangePwForm] = useState(false);
  const [oldPasswordInput, setOldPasswordInput] = useState("");
  const [newPasswordInput, setNewPasswordInput] = useState("");
  const [confirmNewPasswordInput, setConfirmNewPasswordInput] = useState("");
  const [changePwLoading, setChangePwLoading] = useState(false);
  const [changePwMsg, setChangePwMsg] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  const handleChangePassword = async (e: React.FormEvent) => {
    e.preventDefault();
    setChangePwMsg(null);

    if (!oldPasswordInput || !newPasswordInput) {
      setChangePwMsg({ type: 'error', text: 'Vui lòng nhập đầy đủ mật khẩu cũ và mật khẩu mới.' });
      return;
    }

    if (newPasswordInput !== confirmNewPasswordInput) {
      setChangePwMsg({ type: 'error', text: 'Mật khẩu mới và xác nhận mật khẩu không trùng khớp.' });
      return;
    }

    if (newPasswordInput.length < 4) {
      setChangePwMsg({ type: 'error', text: 'Mật khẩu mới phải có ít nhất 4 ký tự.' });
      return;
    }

    setChangePwLoading(true);
    try {
      const savedUserStr = localStorage.getItem('auth_user');
      const headers: Record<string, string> = {
        'Content-Type': 'application/json'
      };
      if (savedUserStr) {
        headers['x-auth-user'] = encodeURIComponent(savedUserStr);
      }

      const res = await fetch('/api/auth/change-password', {
        method: 'POST',
        headers,
        body: JSON.stringify({
          oldPassword: oldPasswordInput,
          newPassword: newPasswordInput
        })
      });

      const data = await res.json();
      if (!res.ok || data.error) {
        setChangePwMsg({ type: 'error', text: data.error || 'Đổi mật khẩu thất bại!' });
      } else {
        setChangePwMsg({ type: 'success', text: data.message || 'Đổi mật khẩu thành công!' });
        setOldPasswordInput('');
        setNewPasswordInput('');
        setConfirmNewPasswordInput('');
      }
    } catch (err: any) {
      setChangePwMsg({ type: 'error', text: 'Lỗi hệ thống khi đổi mật khẩu.' });
    } finally {
      setChangePwLoading(false);
    }
  };

  useEffect(() => {
    const handleScroll = () => {
      if (window.scrollY > 200) {
        setShowTopBtn(true);
      } else {
        setShowTopBtn(false);
      }
    };
    window.addEventListener('scroll', handleScroll);
    return () => window.removeEventListener('scroll', handleScroll);
  }, []);

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

  const fetchAuthAndWorkshops = useCallback(async () => {
    setAuthChecking(true);
    try {
      const wsRes = await fetch('/api/workshops');
      const wsData = await wsRes.json();
      const wsList: Workshop[] = Array.isArray(wsData) ? wsData : [];
      setWorkshops(wsList);

      const savedUserStr = localStorage.getItem('auth_user');
      const headers: Record<string, string> = {};
      if (savedUserStr) {
        headers['x-auth-user'] = encodeURIComponent(savedUserStr);
      }

      const meRes = await fetch('/api/auth/me', { headers });
      const meData = await meRes.json();

      let activeUser: UserAccount | null = null;
      if (meData.authenticated && meData.user) {
        activeUser = meData.user;
        localStorage.setItem('auth_user', JSON.stringify(meData.user));
      } else if (savedUserStr) {
        try {
          activeUser = JSON.parse(savedUserStr);
        } catch (e) {
          localStorage.removeItem('auth_user');
        }
      }

      if (activeUser) {
        setUser(activeUser);
        if (activeUser.role === 'super_admin') {
          setActiveTab('staff');
        } else if (activeUser.role === 'workshop_admin') {
          setActiveTab(prev => prev === 'lookup' ? 'schedule' : prev);
        }
        
        let targetWs = null;
        if (activeUser.workshopId === 'all') {
          targetWs = wsList[0] || null;
        } else {
          targetWs = wsList.find(w => w.id === activeUser.workshopId) || wsList[0] || null;
        }

        if (targetWs) {
          setActiveWorkshop(targetWs);
          if (targetWs.staffData && Array.isArray(targetWs.staffData)) {
            setStaffData(targetWs.staffData);
            localStorage.setItem('sd', JSON.stringify(targetWs.staffData));
            if (targetWs.staffData[0]?.[0]) {
              setChucDanh(targetWs.staffData[0][0].trim());
              setSwapChucDanh(targetWs.staffData[0][0].trim());
            } else {
              setChucDanh('');
              setSwapChucDanh('');
            }
          }
          if (targetWs.config) {
            setConfig(prev => ({
              ...prev,
              ...targetWs.config,
              soVanBan: targetWs.config.soVanBan !== undefined ? targetWs.config.soVanBan : '',
              ngayKy: targetWs.config.ngayKy || '',
              nguoiKy: targetWs.config.nguoiKy || 'Nguyễn Văn Nghị',
              zaloWebhookUrl: targetWs.config.zaloWebhookUrl || 'https://vhialy.dpdns.org/webhook/notify',
              googleSheetUrl: (targetWs.config as any)?.googleSheetUrl || prev.googleSheetUrl || 'https://docs.google.com/spreadsheets/d/1m8B-CVJEyzt5KhJ-I_1hFTvWczz1jl7PPm_7PNScYYQ/edit?gid=0#gid=0',
              companyName: targetWs.config.companyName || 'CÔNG TY THỦY ĐIỆN IALY',
              headerWorkshopName: targetWs.config.headerWorkshopName || 'PHÂN XƯỞNG VẬN HÀNH IALY',
              documentCodeSuffix: targetWs.config.documentCodeSuffix || '/VHIALY',
              shortWorkshopName: targetWs.config.shortWorkshopName || 'PXVH Ialy',
              locationName: targetWs.config.locationName || (targetWs.config as any)?.location || 'Gia Lai'
            }));
            setLeaveData(prev => ({ ...prev, location: targetWs.config?.locationName || (targetWs.config as any)?.location || 'Gia Lai' }));
          }
        }
      } else {
        setUser(null);
        setActiveWorkshop(null);
      }
    } catch (e) {
      console.error("Auth & Workshop fetch error", e);
    } finally {
      setAuthChecking(false);
      setIsSettingsLoaded(true);
    }
  }, []);

  const handleSelectWorkshop = (ws: Workshop) => {
    setActiveWorkshop(ws);
    if (ws.staffData && Array.isArray(ws.staffData)) {
      setStaffData(ws.staffData);
      localStorage.setItem('sd', JSON.stringify(ws.staffData));
      if (ws.staffData[0]?.[0]) {
        setChucDanh(ws.staffData[0][0].trim());
        setSwapChucDanh(ws.staffData[0][0].trim());
      } else {
        setChucDanh('');
        setSwapChucDanh('');
      }
    }
    if (ws.config) {
      setConfig(prev => ({
        ...prev,
        ...ws.config,
        soVanBan: ws.config.soVanBan !== undefined ? ws.config.soVanBan : '',
        ngayKy: ws.config.ngayKy || '',
        nguoiKy: ws.config.nguoiKy || 'Nguyễn Văn Nghị',
        zaloWebhookUrl: ws.config.zaloWebhookUrl || 'https://vhialy.dpdns.org/webhook/notify',
        googleSheetUrl: (ws.config as any)?.googleSheetUrl || prev.googleSheetUrl || 'https://docs.google.com/spreadsheets/d/1m8B-CVJEyzt5KhJ-I_1hFTvWczz1jl7PPm_7PNScYYQ/edit?gid=0#gid=0',
        companyName: ws.config.companyName || 'CÔNG TY THỦY ĐIỆN IALY',
        headerWorkshopName: ws.config.headerWorkshopName || 'PHÂN XƯỞNG VẬN HÀNH IALY',
        documentCodeSuffix: ws.config.documentCodeSuffix || '/VHIALY',
        shortWorkshopName: ws.config.shortWorkshopName || 'PXVH Ialy',
        locationName: ws.config.locationName || (ws.config as any)?.location || 'Gia Lai'
      }));
      setLeaveData(prev => ({ ...prev, location: ws.config?.locationName || (ws.config as any)?.location || 'Gia Lai' }));
    }
  };

  const handleLogout = async () => {
    try {
      await fetch('/api/auth/logout', { method: 'POST' });
    } catch (e) {
      console.error("Logout error", e);
    } finally {
      localStorage.removeItem('auth_user');
      setUser(null);
      setActiveWorkshop(null);
    }
  };

  useEffect(() => {
    fetchAuthAndWorkshops();

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

    const fetchSignatures = async () => {
      try {
        const sigRes = await fetch('/api/signatures');
        const sigData = await sigRes.json();
        setSignatures(sigData);
      } catch (e) {
        console.error("Failed to fetch signatures", e);
      }
    };
    fetchSignatures();

    return () => window.removeEventListener('message', handleMessage);
  }, [fetchAuthAndWorkshops]);

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

        if (activeWorkshop && activeWorkshop.id) {
          const updatedWs = {
            ...activeWorkshop,
            staffData,
            config: {
              ...(activeWorkshop.config || {}),
              ...config
            }
          };
          await fetch('/api/workshops', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(updatedWs)
          });
        }
      } catch (e) {
        console.error("Failed to save app settings", e);
      }
    };
    
    const timer = setTimeout(saveSettings, 1500);
    return () => clearTimeout(timer);
  }, [staffData, config, isSettingsLoaded, activeWorkshop]);

  const handleSaveSystemConfig = async () => {
    setIsSavingConfig(true);
    setConfigSaveMsg(null);
    try {
      await fetch('/api/app-settings', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ staffData, config })
      });

      if (activeWorkshop && activeWorkshop.id) {
        const updatedWs = {
          ...activeWorkshop,
          staffData,
          config: {
            ...(activeWorkshop.config || {}),
            ...config
          }
        };
        const res = await fetch('/api/workshops', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(updatedWs)
        });
        if (res.ok) {
          fetchAuthAndWorkshops();
        }
      }
      setConfigSaveMsg('✅ Đã lưu cấu hình hệ thống thành công!');
      setTimeout(() => setConfigSaveMsg(null), 4000);
    } catch (e: any) {
      setConfigSaveMsg('❌ Lỗi khi lưu cấu hình: ' + (e?.message || 'Lỗi kết nối'));
    } finally {
      setIsSavingConfig(false);
    }
  };

  const [isSyncingLeaves, setIsSyncingLeaves] = useState(false);

  const handleSyncStaffLeavesToSheets = async () => {
    setIsSyncingLeaves(true);
    try {
      const res = await fetch('/api/sheets/sync-staff-leaves', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          staffData,
          googleSheetUrl: config.googleSheetUrl || ''
        })
      });
      const data = await res.json();
      if (res.ok && data.success) {
        setConfigSaveMsg(data.message || "✅ Đã tạo sẵn danh sách nhân viên và số ngày phép lên Google Sheet thành công!");
      } else {
        setConfigSaveMsg(`❌ Đồng bộ thất bại: ${data.error || "Chưa kết nối Google Sheets"}`);
      }
    } catch (e: any) {
      setConfigSaveMsg(`❌ Lỗi khi đồng bộ: ${e.message}`);
    } finally {
      setIsSyncingLeaves(false);
      setTimeout(() => setConfigSaveMsg(null), 6000);
    }
  };

  useEffect(() => {
    if (user && !isAdmin && activeTab === 'staff') {
      setActiveTab('schedule');
    }
  }, [user, isAdmin, activeTab]);

  const fetchWaitingLeaves = useCallback(async () => {
    setIsLoadingWaitingLeaves(true);
    try {
      const url = activeWorkshop?.id ? `/api/sheets/leave-requests?workshopId=${encodeURIComponent(activeWorkshop.id)}` : '/api/sheets/leave-requests';
      const res = await fetch(url);
      if (res.ok) {
        const data = await res.json();
        const mappedData = data.map((item: any) => ({
          ...item,
          chucDanh: ((item.chucDanh === 'Trực phụ cơ MR' ? 'Trực phụ máy MR' : item.chucDanh) || '').trim()
        }));
        
        const waiting = mappedData.filter((item: any) => {
          const statusClean = (item.status || '').trim().toLowerCase().normalize('NFC');
          if (statusClean !== 'chờ phân ca' && statusClean !== 'cho phan ca' && statusClean !== 'chờ duyệt' && statusClean !== 'cho duyat') return false;
          
          if (!activeWorkshop || activeWorkshop.id === 'all') return true;
          
          // Match by workshopId, code, or name
          if (item.workshopId) {
            const wsMatch = item.workshopId === activeWorkshop.id ||
                            item.workshopId.toLowerCase() === (activeWorkshop.code || '').toLowerCase() ||
                            item.workshopId.toLowerCase() === (activeWorkshop.name || '').toLowerCase();
            if (wsMatch) return true;
          }
          
          // Match by staff name in activeWorkshop staffData
          if (activeWorkshop.staffData && Array.isArray(activeWorkshop.staffData)) {
            const allStaffNames = new Set(
              activeWorkshop.staffData.flatMap((row: string[]) => row.slice(1)).filter(Boolean).map((n: string) => n.trim().normalize('NFC'))
            );
            if (allStaffNames.has((item.name || '').trim().normalize('NFC'))) {
              return true;
            }
          }
          
          // Default to true if workshopId is absent/empty so unassigned leaves are shown
          if (!item.workshopId) return true;
          
          return false;
        });

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
  }, [activeWorkshop]);

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
        const originalShift = xacDinhCa(d, res.kip, activeWorkshop?.config?.shiftSchedule?.baseDate, activeWorkshop?.config?.shiftSchedule?.shiftsMatrix);
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
          const absentShift = xacDinhCa(new Date(item.ngay), item.kiptructhay, activeWorkshop?.config?.shiftSchedule?.baseDate, activeWorkshop?.config?.shiftSchedule?.shiftsMatrix);
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
          const originalShift = xacDinhCa(new Date(row.ngay), row.absentKip, activeWorkshop?.config?.shiftSchedule?.baseDate, activeWorkshop?.config?.shiftSchedule?.shiftsMatrix);
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

    const cleanCD = (chucDanh || '').trim();
    const ten = timNghi(cleanCD, kip, staffData);
    if (!ten) {
      setAlert(`⚠ Không tìm thấy chức danh "${cleanCD}" trong Kíp ${kip}!`);
      return;
    }

    const allLeaves: Leave[] = [{ kip, start, end, ten, chucDanh: cleanCD }];
    const addErr: string[] = [];
    additionalLeaves.forEach((al, idx) => {
      if (!al.kip && !al.start && !al.end) return;
      if (!al.kip || !al.start || !al.end || !al.chucDanh) {
        addErr.push(`Người nghỉ #${idx + 2} thiếu thông tin`);
        return;
      }
      const alKip = +al.kip;
      const alChucDanh = (al.chucDanh || '').trim();

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
        const row = staffData.find(r => r[0] && r[0].trim().toLowerCase() === (title || '').trim().toLowerCase());
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

      const customShiftConfig = {
        baseDate: activeWorkshop?.config?.shiftSchedule?.baseDate,
        shifts: activeWorkshop?.config?.shiftSchedule?.shiftsMatrix,
        rules: activeWorkshop?.config?.shiftSchedule?.rulesMatrix,
      };

      Object.keys(groups).forEach(cd => {
        const buildResult = buildMultiLeaveResults(groups[cd], cd, staffData, customShiftConfig);
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

    // Coi như đã xác nhận đơn nghỉ phép, chuyển trạng thái 'Đã xử lý' để không hiển thị trên hàng chờ nữa
    let idsToConfirm: string[] = [...selectedWaitingLeaveIds];

    if (currentResult) {
      const namesInResult = new Set<string>();
      if (currentResult.allResults && Array.isArray(currentResult.allResults)) {
        currentResult.allResults.forEach((r: any) => {
          if (r.ten) namesInResult.add(r.ten.trim().toLowerCase());
        });
      }
      if (currentResult.ten) {
        namesInResult.add(currentResult.ten.trim().toLowerCase());
      }

      waitingLeaves.forEach(item => {
        if (item.name && namesInResult.has(item.name.trim().toLowerCase())) {
          if (!idsToConfirm.includes(item.id)) {
            idsToConfirm.push(item.id);
          }
        }
      });
    }

    if (idsToConfirm.length > 0) {
      try {
        await fetch('/api/sheets/leave-requests/update-status', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ids: idsToConfirm,
            status: "Đã xử lý",
            workshopId: activeWorkshop?.id || ''
          })
        });
        setSelectedWaitingLeaveIds([]);
        fetchWaitingLeaves();
      } catch (err) {
        console.error("Lỗi cập nhật trạng thái đơn nghỉ:", err);
      }
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
    try {
      await exportLeaveRequestDoc(leaveData, config, signatures);
      if (leaveData.name) {
        // Tự động lưu / cập nhật đơn xin nghỉ lên hệ thống với trạng thái 'Chờ phân ca'
        let idsToConfirm: string[] = [];
        const targetName = leaveData.name.trim().toLowerCase();
        waitingLeaves.forEach(item => {
          if (item.name && item.name.trim().toLowerCase() === targetName) {
            idsToConfirm.push(item.id);
          }
        });

        if (idsToConfirm.length > 0) {
          await fetch('/api/sheets/leave-requests/update-status', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              ids: idsToConfirm,
              status: "Chờ phân ca",
              workshopId: activeWorkshop?.id || ''
            })
          });
          fetchWaitingLeaves();
          setAlert(`✅ Đã tải file Word và lưu đơn của đồng chí ${leaveData.name} ở trạng thái 'Chờ phân ca' trên hệ thống!`);
        } else {
          await saveLeaveToGoogleSheets(false, "Chờ phân ca");
        }
      }
    } catch (err: any) {
      console.error("Lỗi khi xuất/lưu đơn:", err);
      setAlert(`❌ Có lỗi khi xuất file Word hoặc lưu lên hệ thống: ${err.message || err}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const saveLeaveToGoogleSheets = async (closeModal: boolean = true, statusOverride?: string) => {
    if (!leaveData.name) {
      setAlert("⚠ Vui lòng nhập Họ và tên người xin nghỉ.");
      return;
    }
    setIsSavingLeaveToSheets(true);
    try {
      const payload = {
        ...leaveData,
        status: statusOverride || "Chờ phân ca",
        workshopId: activeWorkshop?.id || '',
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
        setAlert(resData.message || `✅ Đã lưu đơn của đồng chí ${leaveData.name} lên hệ thống!`);
        fetchWaitingLeaves();
        if (closeModal) {
          setShowPreview(false);
        }
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
    setChucDanh((first.chucDanh || '').trim());
    setKipNghi(String(first.kip));

    // Các đơn còn lại cho vào concurrent / additionalLeaves
    const rem = selectedItems.slice(1);
    const newAdditional = rem.map(item => ({
      chucDanh: (item.chucDanh || '').trim(),
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

      // Write to "Số ngày phép" if requested
      if (isGoogleAuth && shouldUpdateAnnualLeave) {
        try {
          const getTongCaTrucPhaiThay = (resItem: any) => {
            let dem = 0;
            let ngaychao = new Date(resItem.start);
            const ngayEnd = new Date(resItem.end);
            while (ngaychao <= ngayEnd) {
              const cahientai = xacDinhCa(ngaychao, resItem.kip, activeWorkshop?.config?.shiftSchedule?.baseDate, activeWorkshop?.config?.shiftSchedule?.shiftsMatrix);
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
              sheetSyncError = `Lỗi phép năm: ${annualData.error || "Không rõ"}`;
            }
          }
        } catch (annualErr: any) {
          console.error("Annual leave update error:", annualErr);
          sheetSyncError = `Lỗi phép năm: ${annualErr.message}`;
        }
      }

      // Update waiting leaves status to 'Đã xử lý'
      if (selectedWaitingLeaveIds.length > 0) {
        const updateRes = await fetch('/api/sheets/leave-requests/update-status', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            ids: selectedWaitingLeaveIds,
            status: "Đã xử lý",
            workshopId: activeWorkshop?.id || ''
          })
        });
        
        if (updateRes.ok) {
          if (isGoogleAuth) {
            let detailMsg = "";
            if (shouldUpdateAnnualLeave) {
              detailMsg = "Đã tự động ghi số ngày phép vào sheet Số ngày phép.";
            }
            if (sheetSyncError) {
              setAlert(`🎁 Đã xuất ZIP & đổi trạng thái đơn thành 'Đã xử lý', nhưng gặp lỗi: ${sheetSyncError}`);
            } else {
              setAlert(`🎁 Xuất trọn bộ file ZIP thành công! ${detailMsg} Trạng thái đơn đã chuyển thành 'Đã xử lý'.`);
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
          let detailMsg = "";
          if (shouldUpdateAnnualLeave) {
            detailMsg = "Đã cập nhật số ngày phép vào sheet Số ngày phép.";
          }
          if (sheetSyncError) {
            setAlert(`🎁 Đã xuất hồ sơ ZIP, nhưng gặp lỗi: ${sheetSyncError}`);
          } else {
            setAlert(`🎁 Đã xuất trọn bộ hồ sơ dạng ZIP thành công! ${detailMsg}`);
          }
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

  if (authChecking) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center text-white text-sm font-semibold">
        <div className="flex flex-col items-center gap-3">
          <div className="w-8 h-8 border-4 border-emerald-500 border-t-transparent rounded-full animate-spin" />
          <span>Đang xác thực quyền truy cập hệ thống...</span>
        </div>
      </div>
    );
  }

  if (!user) {
    return <LoginForm onLoginSuccess={fetchAuthAndWorkshops} />;
  }

  const isSuperAdmin = user.role === 'super_admin';
  const isWorkshopAdmin = user.role === 'workshop_admin';

  return (
    <div className="min-h-screen flex flex-col bg-[var(--bg)]">
      {/* Top Navbar */}
      <nav id="top-navbar" className="w-full bg-white border-b border-slate-200 py-3.5 px-4 md:px-8 shadow-md flex flex-col lg:flex-row items-center justify-between gap-4 sticky top-0 z-40">
        <div className="flex flex-col sm:flex-row items-center gap-4">
          {/* Logo container */}
          <div className="flex items-center select-none">
            <img 
              src="https://i.ibb.co/r2BTySkt/2343.png" 
              alt="EVN Công Ty Thủy Điện Ialy Logo" 
              className="h-[48px] md:h-[54px] w-auto object-contain"
              referrerPolicy="no-referrer"
            />
          </div>
          {/* App Title with separator */}
          <div className="sm:border-l sm:border-slate-200 sm:pl-4 py-0.5 flex flex-col justify-center h-full text-center sm:text-left">
            <div className="text-[#053d6c] font-black text-base md:text-xl font-sans tracking-tight leading-tight mt-0.5">
              {config?.companyName || activeWorkshop?.config?.companyName || 'CÔNG TY THỦY ĐIỆN IALY'}
            </div>
            <div className="text-[#053d6c] font-black text-base md:text-xl font-sans tracking-tight leading-tight mt-0.5">
              {activeWorkshop ? activeWorkshop.name : 'Phân xưởng Vận hành Ialy'}
            </div>
          </div>
        </div>

        {/* Navigation Tabs on the Right */}
        <div className="flex items-center gap-4 md:gap-7 flex-wrap justify-center">
          {!isSuperAdmin && (
            <>
              <button
                id="nav-schedule-btn"
                className={`text-[14px] md:text-[15px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
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
                className={`text-[14px] md:text-[15px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
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
                className={`text-[14px] md:text-[15px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
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
                  className={`text-[14px] md:text-[15px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
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

          {isAdmin && (
            <button
              id="nav-staff-btn"
              className={`text-[14px] md:text-[15px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
                activeTab === 'staff'
                  ? 'text-[#00529c] border-[#00529c] font-extrabold'
                  : 'text-slate-500 hover:text-[#00529c] border-transparent hover:border-[#00529c]/50'
              }`}
              onClick={() => setActiveTab('staff')}
            >
              {isSuperAdmin ? 'Quản lý Admin' : 'Nhân sự'}
            </button>
          )}

          <button
            id="nav-auth-btn"
            className={`text-[14px] md:text-[15px] font-bold uppercase transition-all cursor-pointer pb-1.5 border-b-2 outline-none ${
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

      {/* Admin Quick Banner for Super Admin / Workshop Admin */}
      {(user?.role === 'super_admin' || user?.role === 'workshop_admin') && (
        <div className="mb-6 p-4 md:p-5 bg-gradient-to-r from-slate-900 via-emerald-950 to-slate-900 text-white rounded-2xl border border-emerald-500/40 shadow-xl flex flex-col sm:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-3.5">
            <div className="p-3 bg-emerald-500/20 text-emerald-400 rounded-xl border border-emerald-500/30 shrink-0">
              <Settings className="w-7 h-7" />
            </div>
            <div>
              <div className="font-extrabold text-base md:text-lg text-white flex items-center gap-2 flex-wrap">
                <span>Xin chào, {user.fullName}!</span>
                <span className={`text-[11px] font-black px-2 py-0.5 rounded-md uppercase tracking-wider ${
                  user.role === 'super_admin' 
                    ? 'bg-amber-400 text-slate-950' 
                    : 'bg-emerald-400 text-slate-950'
                }`}>
                  {user.role === 'super_admin' ? 'Super Admin' : 'Admin Phân Xưởng'}
                </span>
              </div>
              <p className="text-xs text-slate-300 mt-1">
                Tài khoản của bạn có quyền quản trị phân xưởng, phân quyền người dùng, cấu hình người ký văn bản và bật/tắt các tính năng hệ thống.
              </p>
            </div>
          </div>
          <button
            onClick={() => setShowWorkshopManager(true)}
            className="w-full sm:w-auto px-5 py-2.5 bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-black text-xs md:text-sm rounded-xl shadow-lg transition-all cursor-pointer flex items-center justify-center gap-2 whitespace-nowrap active:scale-95 shrink-0"
          >
            <Settings size={17} />
            <span>Mở Bảng Quản lý Phân xưởng & Tài khoản</span>
          </button>
        </div>
      )}

      {activeTab === 'leave' && (
        <>
          <div className="card" id="leave-creator-card">
        <div className="ctitle">Tạo đơn xin nghỉ phép</div>
        <p className="text-[13px] text-var(--txt2) mb-4">Sử dụng chức năng này để tạo nhanh văn bản "Đơn xin nghỉ phép".</p>
        <div className="g2">
          <div className="field">
            <label>Chức danh</label>
            <select value={leaveData.chucDanh} onChange={e => {
              const cd = e.target.value.trim();
              setLeaveData({...leaveData, chucDanh: cd, name: ''});
            }}>
              {staffData.map(r => {
                const title = (r[0] || '').trim();
                return <option key={title} value={title}>{title}</option>;
              })}
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
              {staffData.find(r => r[0] && r[0].trim().toLowerCase() === (leaveData.chucDanh || '').trim().toLowerCase())?.slice(1).filter(Boolean).map(name => (
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
          <div className="col-span-2 mt-2 p-3.5 bg-slate-50 border border-slate-100 rounded-xl">
            <div className="flex justify-between items-center mb-2">
              <span className="text-[13.5px] font-bold text-slate-800 uppercase tracking-wider">Thông tin phép năm</span>
              {leaveData.name && (
                <button 
                  type="button" 
                  onClick={() => fetchLeaveBalance(leaveData.name)}
                  className="text-sm text-[#00529c] hover:underline flex items-center gap-1 font-semibold cursor-pointer"
                  disabled={isLoadingLeaveBalance}
                >
                  {isLoadingLeaveBalance ? <span className="spin mr-1"></span> : '🔄'} Cập nhật
                </button>
              )}
            </div>
            {isLoadingLeaveBalance ? (
              <div className="text-sm text-slate-500 py-2 flex items-center gap-2 justify-center">
                <span className="spin"></span> Đang tải thông tin phép năm ...
              </div>
            ) : leaveBalance ? (
              <div>
                <div className="grid grid-cols-3 gap-2.5">
                  <div className="bg-blue-50/60 p-3 rounded-lg border border-blue-100 text-center">
                    <div className="text-[12px] text-blue-700 font-bold uppercase tracking-wide">Phép được hưởng</div>
                    <div className="text-xl font-extrabold text-blue-900 mt-0.5">{leaveBalance.entitled} <span className="text-sm font-normal">ngày</span></div>
                    {leaveBalance.oldLeaves && parseFloat(leaveBalance.oldLeaves) > 0 ? (
                      <div className="text-[10.5px] text-blue-600 font-medium mt-0.5">({leaveBalance.oldLeaves}d năm cũ + {leaveBalance.newLeaves || '12'}d năm mới)</div>
                    ) : null}
                  </div>
                  <div className="bg-amber-50/60 p-3 rounded-lg border border-amber-100 text-center">
                    <div className="text-[12px] text-amber-700 font-bold uppercase tracking-wide">Phép đã nghỉ</div>
                    <div className="text-xl font-extrabold text-amber-900 mt-0.5">{leaveBalance.used} <span className="text-sm font-normal">ngày</span></div>
                  </div>
                  <div className="bg-emerald-50/60 p-3 rounded-lg border border-emerald-100 text-center">
                    <div className="text-[12px] text-emerald-700 font-bold uppercase tracking-wide">Phép còn lại</div>
                    <div className="text-xl font-extrabold text-emerald-900 mt-0.5">{leaveBalance.remaining} <span className="text-sm font-normal">ngày</span></div>
                  </div>
                </div>
                {leaveBalance.note && (
                  <div className="mt-2 text-[11.5px] text-slate-700 bg-emerald-50/80 border border-emerald-100/80 px-3 py-1.5 rounded-lg flex items-center gap-1.5 font-medium leading-relaxed">
                    <span className="shrink-0">💡</span>
                    <span>{leaveBalance.note}</span>
                  </div>
                )}
              </div>
            ) : leaveBalanceError ? (
              <div className="text-sm text-amber-600 font-medium p-2.5 bg-amber-50 rounded-lg border border-amber-100 text-center">
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
     
</div>
        {alert && <div className={`alert ${alert.startsWith('✅') ? 'asuc' : 'aerr'}`}>{alert}</div>}
        
        <div className="mb-6 p-3.5 sm:p-4 border border-var(--acc-light) rounded-xl bg-var(--acc-light-5) hover:border-var(--acc) transition-all">
          <div className="flex items-center justify-between mb-3 border-b pb-2.5 border-var(--acc-light) flex-wrap gap-2">
            <div className="flex items-center gap-2">
              <span className="text-[13.5px] sm:text-sm md:text-base font-extrabold text-var(--acc) uppercase tracking-wider flex items-center gap-1.5">
                📥 Đơn nghỉ chờ xếp lịch đi ca ({waitingLeaves.length})
              </span>
              {isLoadingWaitingLeaves && <div className="spin w-4 h-4 border-2 border-t-transparent border-var(--acc) rounded-full animate-spin"></div>}
            </div>
            <div className="flex gap-2 flex-wrap select-none">
              <button 
                className="px-2.5 py-1.5 text-xs sm:text-[13px] bg-white text-var(--acc) border border-var(--acc-light) rounded-lg hover:bg-var(--acc-light-10) font-semibold cursor-pointer transition-colors shadow-xs"
                onClick={fetchWaitingLeaves}
                disabled={isLoadingWaitingLeaves}
              >
                🔄 Cập nhật
              </button>
              {selectedWaitingLeaveIds.length > 0 && (
                <>
                  <button 
                    className="px-3 py-1.5 text-xs sm:text-[13px] bg-[#4f46e5] text-white rounded-lg hover:bg-[#4338ca] font-bold cursor-pointer transition-colors shadow-xs"
                    onClick={applyWaitingLeavesToForm}
                  >
                    ⚡ Áp dụng {selectedWaitingLeaveIds.length} người
                  </button>
                  <button 
                    className="px-3 py-1.5 text-xs sm:text-[13px] bg-red-600 text-white rounded-lg hover:bg-red-700 font-bold cursor-pointer transition-colors flex items-center gap-1 shadow-xs"
                    onClick={() => handleDeleteWaitingLeaves(selectedWaitingLeaveIds)}
                  >
                    <Trash2 size={13} /> Xoá {selectedWaitingLeaveIds.length} đơn
                  </button>
                </>
              )}
            </div>
          </div>
          
          {waitingLeaves.length === 0 ? (
            <p className="text-[13px] text-var(--txt2) italic py-1">
              {isGoogleAuth ? "Không có đơn xin nghỉ phép nào ở trạng thái Chờ phân ca trên Google Sheets." : "Không có đơn xin nghỉ phép nào ở trạng thái Chờ phân ca trên hệ thống."}
            </p>
          ) : (
              <div className="max-h-[260px] overflow-y-auto pr-1">
                <div className="flex flex-col gap-2.5">
                  {waitingLeaves.map((leave) => {
                    const isSelected = selectedWaitingLeaveIds.includes(leave.id);
                    return (
                      <div 
                        key={leave.id} 
                        className={`p-3 rounded-lg border text-[13px] sm:text-[13.5px] flex items-start gap-3 transition-all cursor-pointer ${
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
                          className="mt-1 w-4 h-4 accent-[var(--acc)] cursor-pointer flex-shrink-0"
                        />
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center justify-between gap-2 flex-wrap">
                            <span className="font-extrabold text-slate-950 text-[14px] sm:text-[15px]">{leave.name}</span>
                            <div className="flex items-center gap-2" onClick={e => e.stopPropagation()}>
                              <span className="text-[11px] bg-yellow-100 text-yellow-900 border border-yellow-200 px-2 py-0.5 rounded-md font-mono font-semibold">{leave.id.substring(0, 16)}...</span>
                              <button 
                                className="p-1 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors cursor-pointer"
                                title="Xóa đơn này"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleDeleteWaitingLeaves([leave.id], leave.name);
                                }}
                              >
                                <Trash2 size={14} className="text-slate-400 hover:text-red-600 transition-colors" />
                              </button>
                            </div>
                          </div>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-1.5 mt-1.5 text-[13px] sm:text-[13.5px] text-slate-600">
                            <div>Kíp: <span className="font-bold text-slate-900">Kíp {leave.kip}</span></div>
                            <div>Chức danh: <span className="font-bold text-slate-900">{leave.chucDanh}</span></div>
                            <div className="sm:col-span-2">Thời gian: <span className="font-bold text-slate-900">{fmtVN(new Date(leave.startDate))} → {fmtVN(new Date(leave.endDate))}</span></div>
                            <div className="sm:col-span-2">Lý do: <span className="font-bold text-slate-900 truncate block" title={leave.reason}>{leave.reason}</span></div>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
            <p className="text-[12px] text-var(--txt2) mt-2.5 italic">* Chọn các đơn muốn xếp lịch và bấm "Áp dụng ... người" để tự động điền nhanh các thông tin.</p>
          </div>

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
            <select value={chucDanh} onChange={e => setChucDanh(e.target.value.trim())}>
              <option value="">-- Chọn chức danh --</option>
              {staffData.map(r => {
                const title = (r[0] || '').trim();
                return <option key={title} value={title}>{title}</option>;
              })}
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
                  <select value={leave.chucDanh} onChange={e => updateConcurrentLeave(idx, 'chucDanh', e.target.value.trim())}>
                    <option value="">-- Chức danh --</option>
                    {staffData.map(r => {
                      const title = (r[0] || '').trim();
                      return <option key={title} value={title}>{title}</option>;
                    })}
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
              const cd = e.target.value.trim();
              setSwapChucDanh(cd);
              setSwapData({...swapData, person1: '', person2: ''});
            }}>
              {staffData.map(r => {
                const title = (r[0] || '').trim();
                return <option key={title} value={title}>{title}</option>;
              })}
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
            {staffData.find(r => r[0] && r[0].trim().toLowerCase() === (swapChucDanh || '').trim().toLowerCase())?.slice(1).filter(Boolean).map(name => (
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
          {/* Card 1: Thông tin tài khoản */}
          <div className="card w-full p-6 bg-white border border-slate-200 rounded-2xl shadow-md" id="account-info-card">
            <div className="flex items-center justify-between pb-4 border-b border-slate-100">
              <div className="flex items-center gap-3">
                <div className="w-12 h-12 rounded-2xl bg-teal-600 text-white flex items-center justify-center font-bold text-lg shadow-md">
                  👤
                </div>
                <div>
                  <h2 className="text-lg font-black text-slate-800">{user.fullName}</h2>
                  <div className="flex items-center gap-2 mt-0.5">
                    <span className="text-xs text-slate-500 font-mono">@{user.username}</span>
                    <span className={`text-[10px] font-extrabold px-2 py-0.5 rounded-md ${
                      user.role === 'super_admin'
                        ? 'bg-amber-100 text-amber-800 border border-amber-300'
                        : user.role === 'workshop_admin'
                          ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                          : 'bg-slate-100 text-slate-700 border border-slate-200'
                    }`}>
                      {user.role === 'super_admin' ? 'Super Admin (Admin Tổng)' : user.role === 'workshop_admin' ? 'Admin Phân Xưởng' : 'Tài Khoản Phân Xưởng'}
                    </span>
                  </div>
                </div>
              </div>

              <button
                onClick={handleLogout}
                className="flex items-center gap-2 px-4 py-2 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 rounded-xl text-xs font-bold transition-all cursor-pointer shadow-sm active:scale-95"
              >
                <span>🚪 Đăng xuất</span>
              </button>
            </div>

            <div className="mt-6 space-y-4">
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-200/80">
                <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">
                  Phân xưởng hoạt động
                </label>
                {workshops.length > 1 && (user.role === 'super_admin' || user.workshopId === 'all') ? (
                  <select
                    value={activeWorkshop?.id || ''}
                    onChange={(e) => {
                      const found = workshops.find(w => w.id === e.target.value);
                      if (found) handleSelectWorkshop(found);
                    }}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-xl text-sm font-bold text-slate-800 focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                  >
                    {workshops.map(ws => (
                      <option key={ws.id} value={ws.id}>
                        {ws.name} ({ws.code})
                      </option>
                    ))}
                  </select>
                ) : (
                  <div className="font-bold text-emerald-800 text-sm bg-emerald-50 px-3 py-2 rounded-xl border border-emerald-200">
                    {activeWorkshop ? activeWorkshop.name : 'Phân xưởng Vận hành Ialy'}
                  </div>
                )}
              </div>

              {isAdmin && (
                <div className="pt-2">
                  <button
                    onClick={() => setShowWorkshopManager(true)}
                    className="w-full py-3 px-4 bg-amber-500 hover:bg-amber-400 text-slate-950 text-sm font-black rounded-xl shadow-md transition-all cursor-pointer flex items-center justify-center gap-2 active:scale-98"
                  >
                    <Settings size={18} />
                    <span>Quản lý Phân xưởng & Cấu hình Hệ thống</span>
                  </button>
                </div>
              )}
            </div>
          </div>

          {/* Card 2: Đổi mật khẩu */}
          <div className="card w-full bg-white border border-slate-200 rounded-2xl shadow-md overflow-hidden transition-all" id="change-password-card">
            <button
              type="button"
              onClick={() => setShowChangePwForm(!showChangePwForm)}
              className="w-full p-5 flex items-center justify-between bg-white hover:bg-slate-50/80 transition-colors text-left cursor-pointer"
            >
              <div className="flex items-center gap-3">
                <span className="text-xl p-2 bg-amber-50 text-amber-600 rounded-xl border border-amber-200/60">🔐</span>
                <div>
                  <h3 className="text-base font-black text-slate-800">Thay đổi mật khẩu tài khoản</h3>
                  <p className="text-xs text-slate-500 font-medium mt-0.5">
                    {showChangePwForm ? 'Nhấn để ẩn biểu mẫu' : 'Bấm vào đây để mở khung đổi mật khẩu tài khoản'}
                  </p>
                </div>
              </div>
              <span className={`px-3.5 py-1.5 rounded-xl text-xs font-extrabold transition-all ${
                showChangePwForm 
                  ? 'bg-slate-100 text-slate-700 border border-slate-200' 
                  : 'bg-teal-50 text-teal-700 border border-teal-200 hover:bg-teal-100'
              }`}>
                {showChangePwForm ? 'Thu gọn ▲' : 'Đổi mật khẩu ▼'}
              </span>
            </button>

            {showChangePwForm && (
              <form onSubmit={handleChangePassword} className="p-6 pt-4 space-y-4 border-t border-slate-100 bg-slate-50/50 animate-in fade-in duration-150">
                {changePwMsg && (
                  <div className={`p-3.5 rounded-xl text-xs font-bold border ${
                    changePwMsg.type === 'success' 
                      ? 'bg-emerald-50 text-emerald-800 border-emerald-200' 
                      : 'bg-rose-50 text-rose-800 border-rose-200'
                  }`}>
                    {changePwMsg.type === 'success' ? '✅ ' : '⚠️ '}{changePwMsg.text}
                  </div>
                )}

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Mật khẩu hiện tại <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="password"
                    value={oldPasswordInput}
                    onChange={(e) => setOldPasswordInput(e.target.value)}
                    placeholder="Nhập mật khẩu hiện tại"
                    required
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl text-sm font-medium focus:bg-white focus:outline-none focus:ring-2 focus:ring-teal-500 transition-all"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Mật khẩu mới <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="password"
                    value={newPasswordInput}
                    onChange={(e) => setNewPasswordInput(e.target.value)}
                    placeholder="Nhập mật khẩu mới (ít nhất 4 ký tự)"
                    required
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl text-sm font-medium focus:bg-white focus:outline-none focus:ring-2 focus:ring-teal-500 transition-all"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Xác nhận mật khẩu mới <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="password"
                    value={confirmNewPasswordInput}
                    onChange={(e) => setConfirmNewPasswordInput(e.target.value)}
                    placeholder="Xác nhận lại mật khẩu mới"
                    required
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-300 rounded-xl text-sm font-medium focus:bg-white focus:outline-none focus:ring-2 focus:ring-teal-500 transition-all"
                  />
                </div>

                <button
                  type="submit"
                  disabled={changePwLoading}
                  className="w-full py-3 px-4 bg-teal-600 hover:bg-teal-700 disabled:opacity-50 text-white text-sm font-bold rounded-xl shadow-md transition-all cursor-pointer flex items-center justify-center gap-2 active:scale-98"
                >
                  {changePwLoading ? (
                    <span>Đang cập nhật...</span>
                  ) : (
                    <span>💾 Cập nhật mật khẩu</span>
                  )}
                </button>
              </form>
            )}
          </div>
        </div>
      )}

      {activeTab === 'staff' && isAdmin && (
        <div className="flex flex-col gap-6 w-full">
          <div className="card" id="staff-roster-card">
            <div className="ctitle">
              Nhân sự của kíp 
              <div className="flex gap-2 flex-wrap items-center">
                <button 
                  className="px-3 py-1 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-xs rounded-lg flex items-center gap-1.5 transition-all shadow-xs cursor-pointer disabled:opacity-50"
                  onClick={handleSyncStaffLeavesToSheets} 
                  disabled={isSyncingLeaves}
                  title="Khởi tạo danh sách nhân sự và số ngày phép lên Google Sheet"
                >
                  <RefreshCw size={13} className={isSyncingLeaves ? 'animate-spin' : ''} />
                  <span>{isSyncingLeaves ? '⏳ Đang khởi tạo...' : ' Khởi tạo danh sách số ngày nghỉ phép'}</span>
                </button>
                <button className="staff-toggle" onClick={() => setShowStaff(!showStaff)}>
                  {showStaff ? 'Thu gọn ▲' : 'Chỉnh sửa ▼'}
                </button>
                <button className="staff-toggle" onClick={() => setShowSignatureManager(!showSignatureManager)}>
                  {showSignatureManager ? '✍️ Ẩn chữ ký' : '✍️ Quản lý chữ ký'}
                </button>
                <button className="staff-toggle" onClick={() => setShowWorkshopManager(true)}>
                  📊 Quản lý & Excel
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
              <div className="pt-2">
                <StaffDataEditor
                  staffDataText={JSON.stringify(staffData, null, 2)}
                  onChangeStaffDataText={(text) => {
                    try {
                      const parsed = JSON.parse(text);
                      if (Array.isArray(parsed)) {
                        setStaffData(parsed);
                        localStorage.setItem('sd', JSON.stringify(parsed));
                      }
                    } catch (e) {}
                  }}
                  teamsListText={activeWorkshop?.config?.shiftSchedule?.teams?.join(', ') || 'Kíp 1, Kíp 2, Kíp 3, Kíp 4, Kíp 5'}
                  onSyncLeaves={handleSyncStaffLeavesToSheets}
                  isSyncingLeaves={isSyncingLeaves}
                />
              </div>
            )}
          </div>

          <div className="mt-4 flex justify-end">
            <button 
              type="button"
              onClick={() => setShowSystemConfig(!showSystemConfig)}
              className="text-xs text-slate-500 hover:text-slate-800 flex items-center gap-1.5 font-medium px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 transition-all cursor-pointer shadow-xs"
            >
              ⚙️ {showSystemConfig ? 'Ẩn cấu hình hệ thống' : 'Cấu hình hệ thống'}
            </button>
          </div>

          {showSystemConfig && (
            <div className="card mt-3" id="system-config-card">
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
                    placeholder="Ví dụ: 123/PX"
                  />
                </div>
                <div className="field">
                  <label>Hậu tố ký hiệu văn bản</label>
                  <input 
                    type="text" 
                    value={config.documentCodeSuffix || ''} 
                    onChange={e => setConfig({ ...config, documentCodeSuffix: e.target.value })} 
                    placeholder="Ví dụ: /VHIALY hoặc /VHPLK"
                  />
                </div>
                <div className="field">
                  <label>Tên Phân xưởng trên văn bản</label>
                  <input 
                    type="text" 
                    value={config.headerWorkshopName || ''} 
                    onChange={e => setConfig({ ...config, headerWorkshopName: e.target.value })} 
                    placeholder="Ví dụ: PHÂN XƯỞNG VẬN HÀNH IALY"
                  />
                </div>
                <div className="field">
                  <label>Tên Công ty / Đơn vị</label>
                  <input 
                    type="text" 
                    value={config.companyName || ''} 
                    onChange={e => setConfig({ ...config, companyName: e.target.value })} 
                    placeholder="Ví dụ: CÔNG TY THỦY ĐIỆN IALY"
                  />
                </div>
                <div className="field">
                  <label>Địa điểm / Tỉnh thành (Xuất văn bản)</label>
                  <input 
                    type="text" 
                    value={config.locationName || ''} 
                    onChange={e => setConfig({ ...config, locationName: e.target.value })} 
                    placeholder="Ví dụ: Gia Lai, Kon Tum, Đắk Lắk..."
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

                <div className="field col-span-2 p-3.5 bg-emerald-50/50 border border-emerald-200/80 rounded-xl space-y-2">
                  <div className="flex items-center justify-between">
                    <label className="font-bold text-slate-800 text-xs uppercase tracking-wider text-emerald-800">
                      📊 Link Google Sheet (Lưu Lịch Sử & Tra Cứu Số Ngày Phép)
                    </label>
                    <a 
                      href={config.googleSheetUrl || 'https://docs.google.com/spreadsheets/d/1m8B-CVJEyzt5KhJ-I_1hFTvWczz1jl7PPm_7PNScYYQ/edit?gid=0#gid=0'} 
                      target="_blank" 
                      rel="noopener noreferrer" 
                      className="text-[11px] text-emerald-700 hover:underline flex items-center gap-1 font-semibold"
                    >
                      🔗 Mở Google Sheet
                    </a>
                  </div>
                  <div className="flex flex-col sm:flex-row gap-2">
                    <input 
                      type="text" 
                      value={config.googleSheetUrl || ''} 
                      onChange={e => setConfig({ ...config, googleSheetUrl: e.target.value })} 
                      placeholder="https://docs.google.com/spreadsheets/d/1m8B-CVJEyzt5KhJ-I_1hFTvWczz1jl7PPm_7PNScYYQ/edit"
                      className="font-mono text-[12px] flex-1 px-3 py-2 bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-emerald-500/20"
                    />
                    <button
                      type="button"
                      onClick={handleSyncStaffLeavesToSheets}
                      disabled={isSyncingLeaves}
                      className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg text-xs whitespace-nowrap flex items-center justify-center gap-1.5 cursor-pointer disabled:opacity-50 transition-all shadow-sm"
                      title="Khởi tạo sẵn danh sách nhân viên và 12 ngày phép/năm lên sheet 'Số ngày phép'"
                    >
                      {isSyncingLeaves ? '⏳ Đang khởi tạo...' : '🔄 Khởi tạo sheet Số ngày phép'}
                    </button>
                  </div>
                  <span className="text-[11px] text-slate-600 italic block leading-relaxed">
                    * Nếu để trống, hệ thống sẽ mặc định lưu và tra cứu tại: <code className="bg-white px-1 py-0.5 rounded border border-slate-200 text-emerald-800 font-mono">https://docs.google.com/spreadsheets/d/1m8B-CVJEyzt5KhJ-I_1hFTvWczz1jl7PPm_7PNScYYQ/edit</code>. Khi bạn dán link Google Sheet riêng của mình, lịch sử nghỉ phép và tra cứu số ngày phép sẽ được lưu trữ trực tiếp vào sheet đó.
                  </span>
                </div>

                <div className="col-span-2 flex flex-col sm:flex-row items-center justify-between gap-3 pt-3 border-t border-slate-200 mt-2">
                  <div className="text-xs font-semibold">
                    {configSaveMsg && (
                      <span className={configSaveMsg.startsWith('✅') ? 'text-emerald-700 bg-emerald-50 px-3 py-1.5 rounded-lg border border-emerald-200' : 'text-rose-700 bg-rose-50 px-3 py-1.5 rounded-lg border border-rose-200'}>
                        {configSaveMsg}
                      </span>
                    )}
                  </div>
                  <button
                    type="button"
                    onClick={handleSaveSystemConfig}
                    disabled={isSavingConfig}
                    className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-xs flex items-center gap-2 shadow-sm transition-all cursor-pointer disabled:opacity-50"
                  >
                    {isSavingConfig ? '⏳ Đang lưu...' : '💾 Lưu Cấu Hình Hệ Thống'}
                  </button>
                </div>
              </div>
            </div>
          )}
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
                    onClick={() => saveLeaveToGoogleSheets(true)}
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
            Hệ thống tự động tạo lịch trực thay ca vận hành nghỉ phép VHIALY
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
      </div>
    </footer>

    {/* Floating Top scroll button when scrolled down */}
    {showTopBtn && (
      <button 
        onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })} 
        className="fixed bottom-6 right-6 z-50 flex items-center bg-white text-slate-800 rounded-full pl-1.5 pr-4 py-1.5 transition-all duration-200 outline-none border border-slate-200 cursor-pointer shadow-xl hover:bg-slate-50 hover:shadow-2xl hover:scale-105 active:scale-95 animate-fade-in"
        title="Về đầu trang"
      >
        <span className="h-8 w-8 bg-[#00adef] text-white font-bold rounded-full flex items-center justify-center mr-2.5 shadow-sm leading-none text-base">
          ↑
        </span>
        <span className="text-[13px] font-extrabold tracking-wider uppercase text-[#05203c] underline decoration-[#00adef] decoration-2">
          Top
        </span>
      </button>
    )}

    {/* Workshop & Accounts Management Modal */}
    {showWorkshopManager && user && (
      <WorkshopManagerModal
        user={user}
        workshops={workshops}
        activeWorkshop={activeWorkshop}
        onClose={() => setShowWorkshopManager(false)}
        onRefreshWorkshops={fetchAuthAndWorkshops}
      />
    )}
    </div>
  );
}

