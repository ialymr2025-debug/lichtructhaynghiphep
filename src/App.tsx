
import React, { useState, useRef, useEffect, useMemo, useCallback } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import mammoth from 'mammoth';
import { initializeApp } from 'firebase/app';
import { 
  initializeFirestore, 
  collection, 
  addDoc, 
  onSnapshot, 
  query, 
  orderBy, 
  serverTimestamp,
  Timestamp,
  doc,
  getDocFromServer,
  updateDoc,
  arrayUnion,
  deleteDoc,
  where,
  getDocs,
  limit
} from 'firebase/firestore';
import { getAuth, signInWithPopup, GoogleAuthProvider, onAuthStateChanged, User, setPersistence, browserLocalPersistence } from 'firebase/auth';
import firebaseConfig from '../firebase-applet-config.json';
import { 
  PieChart, 
  Pie, 
  Cell, 
  ResponsiveContainer, 
  Tooltip 
} from 'recharts';
import { 
  ExternalLink, 
  Upload,
  FileText,
  CheckCircle2,
  Eye,
  Settings,
  FileSearch,
  ClipboardList,
  Activity as ActivityIcon,
  AlertTriangle,
  RotateCcw,
  Maximize2,
  Download,
  Info,
  LayoutDashboard,
  FolderArchive,
  ChevronDown,
  Calendar,
  BarChart3,
  Droplets,
  Zap,
  ShieldCheck,
  TrendingUp,
  RefreshCw,
  Clock,
  AlertCircle,
  ChevronRight,
  Trash2
} from 'lucide-react';

export enum DefectStatus {
  PENDING = 'PENDING',
  PROCESSING = 'PROCESSING',
  COMPLETED = 'COMPLETED',
  URGENT = 'URGENT'
}

export interface ChartData {
  name: string;
  detected: number;
  processed: number;
  processing: number;
  pending: number;
}
import StatCard from './components/StatCard';
import { DailyLoadChart, MonthlyProductionChart } from './components/ProductionChart';
import { DashboardCharts } from './components/DashboardCharts';
import { WastewaterChart } from './components/WastewaterChart';
import { ReportModal } from './components/ReportModal';
import { TargetModal } from './components/TargetModal';
import { BulkTargetModal } from './components/BulkTargetModal';
import { AuditCharts } from './components/AuditCharts';
import { AdminMonitor } from './components/AdminMonitor';
import { EvnLogoWidget } from './components/EvnLogoWidget';
import { generateOperationReport } from './services/exportService';
import { INITIAL_TARGETS, INITIAL_REPORTS, REPORTERS } from './constants';
import { Target, Report, TargetStatus } from './types';
import { cn } from './lib/utils';
import * as htmlToImage from 'html-to-image';

// Initialize Firebase
const app = initializeApp(firebaseConfig);
export const db = initializeFirestore(app, {
  experimentalForceLongPolling: true,
}, firebaseConfig.firestoreDatabaseId);
const auth = getAuth(app);
const googleProvider = new GoogleAuthProvider();
googleProvider.addScope('https://www.googleapis.com/auth/drive.file');

// Constants
const SHEET_ID = '1EVA37o8kSgi3Z86hwUQN5uyBtVwERDo3REO0xMtMqE0';
const CATEGORIES = ['QUẢN LÝ HÀNH CHÍNH', 'THIẾT BỊ CÔNG TRÌNH', 'AN TOÀN VỆ SINH LAO ĐỘNG', 'TPM, KAIZEN'];

const DRIVE_FOLDERS = {
  GENERAL: '1zEIi5d40X3jIYAHIAD0r8JSzZV69cILA',
  KPI: '1sPbWy_9n0OgtIEP_gHf7i07aTGe5bWno',
  LHC: '1SWAB8E6S1SDUcFhDk0XwLtcBpG-GS7Xy',
  SAFETY: '1kJpQJykLqoPflAemE6I-CEU3VcBntEqG',
  ECONOMY: '125pphQ43ycvbZOWqwXH0NkiYPIXfqrm9'
};

// Type definitions
interface UploadedFile {
  id: string;
  name: string;
  type: string;
  size: string;
  url: string;
  date: string;
  originalFile?: File;
  uploaderName?: string;
  uploaderId?: string;
  createdAt?: any;
  driveId?: string;
  category?: string;
  readBy?: string[];
}

interface ActivityLog {
  id: string;
  userId: string;
  userEmail: string;
  userName: string;
  action: string;
  details: string;
  timestamp: any;
}

const ADMIN_EMAILS = ['hoaibaole2k00@gmail.com', 'hoaibaole123@gmail.com', 'hungvu059@gmail.com'];

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [authReady, setAuthReady] = useState(false);
  const [mainView, setMainView] = useState<'portal' | 'dashboard' | 'repository' | 'admin'>('portal');
  const [activeTab, setActiveTab] = useState<'chutri' | 'phoihoap'>('phoihoap');
  const [hasLoggedAccess, setHasLoggedAccess] = useState(false);

  const logActivity = useCallback(async (action: string, details: string) => {
    if (!auth.currentUser) return;
    const path = 'activity_logs';
    try {
      await addDoc(collection(db, path), {
        userId: auth.currentUser.uid,
        userEmail: auth.currentUser.email,
        userName: auth.currentUser.displayName || auth.currentUser.email?.split('@')[0],
        action,
        details,
        timestamp: serverTimestamp()
      });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, path);
    }
  }, []);

  // Tự động ghi nhận hoạt động truy cập khi khôi phục phiên đăng nhập (lưu trạng thái)
  useEffect(() => {
    if (user && !hasLoggedAccess) {
      setHasLoggedAccess(true);
      logActivity('Truy cập', 'Truy cập ứng dụng (Khôi phục phiên đăng nhập)');
    }
  }, [user, hasLoggedAccess, logActivity]);

  const isAdmin = useMemo(() => user && ADMIN_EMAILS.includes(user.email || ''), [user]);
  const [portalTab, setPortalTab] = useState<'kt' | 'qt' | 'at'>('kt');
  const [openAcc, setOpenAcc] = useState<number | null>(null);
  const [openSub, setOpenSub] = useState<string | null>(null);
  const [dynamicPortalItems, setDynamicPortalItems] = useState<any[]>([]);
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [uploadCategory, setUploadCategory] = useState<'GENERAL' | 'KPI' | 'LHC' | 'SAFETY' | 'ECONOMY'>('GENERAL');
  const [selectedFile, setSelectedFile] = useState<UploadedFile | null>(null);
  const [notification, setNotification] = useState<{message: string, type: 'success' | 'error'} | null>(null);
  const [fileToDelete, setFileToDelete] = useState<{id: string, name: string} | null>(null);
  const [wordContent, setWordContent] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  const [stats, setStats] = useState({ total: 0, completed: 0, pending: 0, notStarted: 0 });
  
  // Dashboard Metrics State
  const [chartData, setChartData] = useState<ChartData[]>(
    CATEGORIES.map(cat => ({ name: cat, detected: 0, processed: 0, processing: 0, pending: 0 }))
  );
  const [recentActivities, setRecentActivities] = useState<any[]>([]);

  const fetchDashboardData = useCallback(async () => {
    setIsLoading(true);
    try {
      const response = await fetch('/api/dashboard-summary');
      if (!response.ok) throw new Error('Failed to fetch dashboard summary');
      
      const results = await response.json();
      
      let combinedActivities: any[] = [];
      let totalCount = 0;
      let completedCount = 0;
      let processingCount = 0;
      let notStartedCount = 0;

      results.forEach((res: any) => {
        if (res.error) return;
        
        totalCount += res.detected || 0;
        completedCount += res.processed || 0;
        processingCount += res.processing || 0;
        notStartedCount += res.pending || 0;
        
        if (res.activities) {
          combinedActivities = [...combinedActivities, ...res.activities];
        }
      });

      setChartData(results.map((r: any) => ({
        name: r.name,
        detected: r.detected || 0,
        processed: r.processed || 0,
        processing: r.processing || 0,
        pending: r.pending || 0
      })));

      const parseDateObj = (val: any) => {
        if (!val) return null;
        const str = String(val);
        if (str.startsWith('Date(')) {
          const p = str.match(/\d+/g);
          if (p) return new Date(Number(p[0]), Number(p[1]), Number(p[2]));
        }
        const ts = Date.parse(str);
        return isNaN(ts) ? null : new Date(ts);
      };

      setRecentActivities(combinedActivities.sort((a, b) => {
        const dateA = parseDateObj(a.rawTime)?.getTime() || 0;
        const dateB = parseDateObj(b.rawTime)?.getTime() || 0;
        return dateB - dateA;
      }).slice(0, 6));

      setStats({ 
        total: totalCount, 
        completed: completedCount, 
        pending: processingCount, 
        notStarted: notStartedCount 
      });
    } catch (error) {
      console.error('Failed to fetch dashboard data:', error);
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchDashboardData();
  }, [fetchDashboardData]);

  const onStatClick = (name: string, type: string) => {
    console.log(`Stat clicked: ${name} - ${type}`);
  };

  const onActivityClick = (category: string, row: number) => {
    console.log(`Activity clicked: ${category} - Row ${row}`);
  };

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Dashboard Module State
  const [manualTargets, setManualTargets] = useState<Target[]>(INITIAL_TARGETS);
  const [reports, setReports] = useState<Report[]>(INITIAL_REPORTS);
  const [googleTargets, setGoogleTargets] = useState<Target[]>([]);
  const [isLoadingTargets, setIsLoadingTargets] = useState(false);
  const [dashboardSubTab, setDashboardSubTab] = useState('dashboard');
  const [isReportModalOpen, setIsReportModalOpen] = useState(false);
  const [isTargetModalOpen, setIsTargetModalOpen] = useState(false);
  const [isBulkTargetModalOpen, setIsBulkTargetModalOpen] = useState(false);
  const [editingTarget, setEditingTarget] = useState<Target | null>(null);
  const [isExporting, setIsExporting] = useState(false);
  const [isPasswordModalOpen, setIsPasswordModalOpen] = useState(false);
  const [passwordValue, setPasswordValue] = useState('');
  const [pendingAction, setPendingAction] = useState<{ type: 'save' | 'delete', data: any } | null>(null);

  useEffect(() => {
    const fetchGoogleTargets = async () => {
      setIsLoadingTargets(true);
      try {
        const resp = await fetch('/api/google-targets');
        if (resp.ok) {
          const data = await resp.json();
          setGoogleTargets(data);
        }
      } catch (err) {
        console.error('Failed to fetch google targets:', err);
      } finally {
        setIsLoadingTargets(false);
      }
    };
    fetchGoogleTargets();
  }, []);

  const allTargets = useMemo(() => {
    const googleIds = new Set(googleTargets.map(t => t.id));
    const filteredInitial = manualTargets.filter(t => !googleIds.has(t.id));
    return [...filteredInitial, ...googleTargets];
  }, [manualTargets, googleTargets]);

  const handleDownloadReport = async () => {
    setIsExporting(true);
    try {
      let chartImages: { pie: string; bar: string } | undefined;
      const pieEl = document.getElementById(`chart-pie-yearly`);
      const barEl = document.getElementById(`chart-bar-plants`);
      if (pieEl && barEl) {
        const pieImg = await htmlToImage.toPng(pieEl, { quality: 1, backgroundColor: '#ffffff' });
        const barImg = await htmlToImage.toPng(barEl, { quality: 1, backgroundColor: '#ffffff' });
        chartImages = { pie: pieImg, bar: barImg };
      }
      await generateOperationReport(allTargets, reports, chartImages);
    } catch (error) {
      console.error('Export failed:', error);
    } finally {
      setIsExporting(false);
    }
  };

  const handleSaveReport = (newReportData: Omit<Report, 'id' | 'date'>) => {
    const today = new Date();
    const date = `${String(today.getDate()).padStart(2, '0')}/${String(today.getMonth() + 1).padStart(2, '0')}/${today.getFullYear()}`;
    const newReport: Report = { ...newReportData, id: `r-${Date.now()}`, date };
    setReports([newReport, ...reports]);
  };

  const handleConfirmAction = () => {
    if (passwordValue === '123456') {
      if (pendingAction?.type === 'save') {
        const targetData = pendingAction.data;
        if (editingTarget) {
          setManualTargets(manualTargets.map(t => t.id === editingTarget.id ? { ...targetData, id: t.id } : t));
        } else {
          setManualTargets([{ ...targetData, id: `t-${Date.now()}` }, ...manualTargets]);
        }
        setIsTargetModalOpen(false);
        setEditingTarget(null);
      } else if (pendingAction?.type === 'delete') {
        setManualTargets(manualTargets.filter(t => t.id !== pendingAction.data));
      }
      setIsPasswordModalOpen(false);
      setPasswordValue('');
      setPendingAction(null);
    } else { alert('Sai mật khẩu!'); }
  };

  const renderDashboardModule = () => {
    return (
      <div className="space-y-12">
        <div className="flex bg-white p-2 rounded-[28px] border border-slate-100 shadow-sm w-fit">
           {[
             { id: 'dashboard', label: 'Biểu đồ vận hành' }
           ].map(tab => (
             <button 
               key={tab.id}
               onClick={() => setDashboardSubTab(tab.id)}
               className={cn(
                 "px-6 py-3 rounded-[20px] text-xs font-black uppercase tracking-widest transition-all",
                 dashboardSubTab === tab.id ? "bg-blue-600 text-white shadow-lg shadow-blue-100" : "text-slate-400 hover:text-slate-600"
               )}
             >
               {tab.label}
             </button>
           ))}
        </div>

        <AnimatePresence mode="wait">
          {dashboardSubTab === 'dashboard' && (
            <motion.div key="db-charts" initial={{ opacity: 0 }} animate={{ opacity: 1 }} className="space-y-12">
               <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                <div>
                  <h2 className="text-3xl font-black text-gray-900 tracking-tight flex items-center gap-3">
                    <div className="p-2 bg-blue-600 rounded-2xl text-white shadow-lg shadow-blue-600/20">
                      <BarChart3 size={24} />
                    </div>
                    BIỂU ĐỒ DỮ LIỆU VẬN HÀNH
                  </h2>
                  <p className="text-sm text-gray-500 font-medium mt-2 flex items-center gap-2">
                    <RefreshCw size={14} className="text-blue-500" />
                    Cập nhật thời thực: {new Date().toLocaleTimeString('vi-VN')} · {new Date().toLocaleDateString('vi-VN')}
                  </p>
                </div>
                
                
              </div>
              <DashboardCharts />
              
              <div className="flex items-center gap-3 mb-6">
                <div className="p-2.5 bg-blue-100 dark:bg-blue-900/30 rounded-2xl text-blue-600"><ActivityIcon size={18} /></div>
                <h3 className="text-xl font-black text-slate-900 tracking-tight uppercase">TỒN TẠI, ĐIỂM KHÔNG PHÙ HỢP</h3>
              </div>
              
              <div className="grid grid-cols-2 lg:grid-cols-4 gap-6 mb-8">
                <div className="bg-white dark:bg-slate-800 p-6 rounded-[2rem] border border-slate-100 dark:border-slate-800 shadow-sm">
                  <p className="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2">Tổng phát hiện</p>
                  <p className="text-3xl font-black text-slate-900 dark:text-white leading-none">{stats.total}</p>
                </div>
                <div className="bg-white dark:bg-slate-800 p-6 rounded-[2rem] border border-slate-100 dark:border-slate-800 shadow-sm">
                  <p className="text-[10px] font-black uppercase tracking-widest text-emerald-500 mb-2">Đã hoàn thành</p>
                  <p className="text-3xl font-black text-emerald-600 leading-none">{stats.completed}</p>
                </div>
                <div className="bg-white dark:bg-slate-800 p-6 rounded-[2rem] border border-slate-100 dark:border-slate-800 shadow-sm">
                  <p className="text-[10px] font-black uppercase tracking-widest text-amber-500 mb-2">Đang xử lý</p>
                  <p className="text-3xl font-black text-amber-600 leading-none">{stats.pending}</p>
                </div>
                <div className="bg-white dark:bg-slate-800 p-6 rounded-[2rem] border border-slate-100 dark:border-slate-800 shadow-sm">
                  <p className="text-[10px] font-black uppercase tracking-widest text-rose-500 mb-2">Chưa xử lý</p>
                  <p className="text-3xl font-black text-rose-600 leading-none">{stats.notStarted}</p>
                </div>
              </div>

              <div className="bg-white dark:bg-slate-800 p-8 rounded-[2.5rem] shadow-2xl border border-slate-100 dark:border-slate-800 mb-6">
                <div className="flex items-center justify-between mb-8">
                  <div className="flex items-center gap-3">
                    <div className="p-2.5 bg-blue-100 dark:bg-blue-900/30 rounded-2xl text-blue-600"><ActivityIcon size={18} /></div>
                    
                    <h2 className="text-sm font-black uppercase tracking-widest text-slate-700 dark:text-slate-300">Biểu đồ hoạt động</h2>
                  </div>
                  <button onClick={fetchDashboardData} className="p-2 text-slate-400 hover:text-blue-500 transition-colors">
                    <RefreshCw size={16} className={isLoading ? 'animate-spin' : ''} />
                  </button>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-8">
                  {isLoading ? (
                    Array(4).fill(0).map((_, i) => (
                      <div key={i} className="h-48 bg-slate-50 dark:bg-slate-900/50 rounded-3xl animate-pulse" />
                    ))
                  ) : (
                    chartData.map((entry, idx) => {
                      const completed = entry.processed;
                      const totalStarted = entry.processing;
                      const pending = entry.pending;
                      const inProgressOnly = Math.max(0, entry.detected - completed - pending);
                      
                      const data = [
                        { name: 'Hoàn thành', value: completed, color: '#10b981', grad: `gradGreen-${idx}` },
                        { name: 'Đang xử lý', value: inProgressOnly, color: '#f59e0b', grad: `gradAmber-${idx}` },
                        { name: 'Chưa xử lý', value: pending, color: '#ef4444', grad: `gradRed-${idx}` }
                      ].filter(d => d.value > 0);

                      return (
                        <div key={idx} className="flex flex-col items-center group">
                          <div className="h-60 w-full relative bg-slate-900/5 dark:bg-slate-900/20 rounded-[2.5rem] overflow-hidden shadow-inner border border-slate-200/50 dark:border-slate-700/30 transition-all hover:shadow-2xl hover:scale-[1.02] duration-500" id={`activity-chart-${idx}`}>
                            <div className="absolute inset-0 bg-gradient-to-br from-white/40 to-transparent dark:from-white/5 pointer-events-none" />
                            
                            {/* Corner Buttons */}
                            <button 
                              onClick={() => onStatClick(entry.name, 'processed')}
                              className="absolute top-4 left-4 z-20 flex flex-col items-center p-3.5 bg-emerald-50/90 dark:bg-emerald-900/50 backdrop-blur-md rounded-[1.5rem] border border-emerald-100 dark:border-emerald-800/50 hover:scale-110 active:scale-95 transition-all shadow-xl min-w-[70px]"
                            >
                              <span className="text-[9px] font-black text-emerald-600 dark:text-emerald-400 uppercase tracking-widest mb-1">Hoàn thành</span>
                              <span className="text-lg font-black text-emerald-700 dark:text-emerald-300 leading-none">{completed}</span>
                            </button>

                            <button 
                              onClick={() => onStatClick(entry.name, 'nvvh')}
                              className="absolute top-4 right-4 z-20 flex flex-col items-center p-3.5 bg-amber-50/90 dark:bg-amber-900/50 backdrop-blur-md rounded-[1.5rem] border border-amber-100 dark:border-amber-800/50 hover:scale-110 active:scale-95 transition-all shadow-xl min-w-[70px]"
                            >
                              <span className="text-[9px] font-black text-amber-600 dark:text-amber-400 uppercase tracking-widest mb-1">Đang xử lý</span>
                              <span className="text-lg font-black text-amber-700 dark:text-amber-300 leading-none">{totalStarted}</span>
                            </button>

                            <button 
                              onClick={() => onStatClick(entry.name, 'pending')}
                              className="absolute bottom-4 left-4 z-20 flex flex-col items-center p-3.5 bg-rose-50/90 dark:bg-rose-900/50 backdrop-blur-md rounded-[1.5rem] border border-rose-100 dark:border-rose-800/50 hover:scale-110 active:scale-95 transition-all shadow-xl min-w-[70px]"
                            >
                              <span className="text-[9px] font-black text-rose-600 dark:text-rose-400 uppercase tracking-widest mb-1">Chưa xử lý</span>
                              <span className="text-lg font-black text-rose-700 dark:text-rose-300 leading-none">{pending}</span>
                            </button>

                            <ResponsiveContainer width="100%" height="100%" minWidth={100} minHeight={100}>
                              <PieChart>
                                <defs>
                                  <linearGradient id={`gradGreen-${idx}`} x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="0%" stopColor="#10b981" />
                                    <stop offset="100%" stopColor="#059669" />
                                  </linearGradient>
                                  <linearGradient id={`gradAmber-${idx}`} x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="0%" stopColor="#f59e0b" />
                                    <stop offset="100%" stopColor="#d97706" />
                                  </linearGradient>
                                  <linearGradient id={`gradRed-${idx}`} x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="0%" stopColor="#ef4444" />
                                    <stop offset="100%" stopColor="#b91c1c" />
                                  </linearGradient>
                                  <filter id="vividShadow" height="150%">
                                    <feGaussianBlur in="SourceAlpha" stdDeviation="3" />
                                    <feOffset dx="0" dy="10" result="offsetblur" />
                                    <feComponentTransfer>
                                      <feFuncA type="linear" slope="0.5" />
                                    </feComponentTransfer>
                                    <feMerge>
                                      <feMergeNode />
                                      <feMergeNode in="SourceGraphic" />
                                    </feMerge>
                                  </filter>
                                </defs>
                                <Pie
                                  data={data}
                                  cx="50%"
                                  cy="50%"
                                  startAngle={90}
                                  endAngle={-270}
                                  innerRadius={35}
                                  outerRadius={75}
                                  paddingAngle={4}
                                  dataKey="value"
                                  stroke="none"
                                  filter="url(#vividShadow)"
                                  animationDuration={1800}
                                  labelLine={false}
                                  label={({ cx, cy, midAngle, innerRadius, outerRadius, percent }) => {
                                    const radius = innerRadius + (outerRadius - innerRadius) * 0.5;
                                    const x = cx + radius * Math.cos(-midAngle * (Math.PI / 180));
                                    const y = cy + radius * Math.sin(-midAngle * (Math.PI / 180));
                                    if (percent < 0.1) return null;
                                    return (
                                      <text 
                                        x={x} 
                                        y={y} 
                                        fill="white" 
                                        textAnchor="middle" 
                                        dominantBaseline="central" 
                                        className="text-[11px] font-black drop-shadow-lg"
                                      >
                                        {(percent * 100).toFixed(0)}%
                                      </text>
                                    );
                                  }}
                                >
                                  {data.map((entry, index) => (
                                    <Cell 
                                      key={`cell-${index}`} 
                                      fill={`url(#${entry.grad})`}
                                      className="transition-all duration-500 hover:opacity-80 cursor-pointer"
                                    />
                                  ))}
                                </Pie>
                                <Tooltip 
                                  content={({ active, payload }) => {
                                    if (active && payload && payload.length) {
                                      const d = payload[0].payload;
                                      return (
                                        <div className="bg-white dark:bg-slate-900 p-3 rounded-2xl shadow-2xl border border-slate-100 dark:border-slate-800 animate-in fade-in zoom-in duration-200">
                                          <p className="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2">{d.name}</p>
                                          <p className="text-lg font-black text-slate-900 dark:text-white">{d.value}</p>
                                        </div>
                                      );
                                    }
                                    return null;
                                  }}
                                />
                              </PieChart>
                            </ResponsiveContainer>
                            <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-center pointer-events-none">
                              <p className="text-[10px] font-black text-slate-400 uppercase tracking-tighter leading-none">Tổng</p>
                              <p className="text-lg font-black text-slate-700 dark:text-slate-300">{entry.detected}</p>
                            </div>
                          </div>
                          <div className="mt-5 text-center w-full">
                            <p className="text-xs font-black uppercase tracking-[0.2em] text-slate-800 dark:text-slate-200 mb-1">{entry.name}</p>
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>
              </div>

              <div className="bg-white dark:bg-slate-800 p-6 rounded-[2.5rem] shadow-2xl border border-slate-100 dark:border-slate-800 mb-6">
                <div className="flex items-center gap-3 mb-6">
                  <div className="p-2.5 bg-blue-100 dark:bg-blue-900/30 rounded-2xl text-blue-600"><Clock size={18} /></div>
                  <h2 className="text-sm font-black uppercase tracking-widest text-slate-700 dark:text-slate-300">Hoạt động mới nhất</h2>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {isLoading ? (
                    Array(4).fill(0).map((_, i) => (
                      <div key={i} className="h-20 bg-slate-50 dark:bg-slate-900/50 rounded-3xl animate-pulse" />
                    ))
                  ) : recentActivities.length > 0 ? (
                    recentActivities.map((act, idx) => (
                      <div 
                        key={idx} 
                        onClick={() => onActivityClick(act.category, act.row)}
                        className="flex items-center gap-4 p-4 bg-slate-50 dark:bg-slate-900/40 rounded-3xl border border-slate-100 dark:border-slate-800/50 group transition-all hover:shadow-lg cursor-pointer active:scale-[0.98]"
                      >
                        <div className={`shrink-0 w-12 h-12 rounded-[1.2rem] flex items-center justify-center shadow-sm ${act.isDone ? 'bg-emerald-100 text-emerald-600' : 'bg-amber-100 text-amber-600'}`}>
                          {act.isDone ? <CheckCircle2 size={20} /> : <AlertCircle size={20} />}
                        </div>
                        <div className="flex-1 min-w-0">
                          <h3 className="text-[16px] font-black text-slate-800 dark:text-slate-200 truncate leading-tight mb-1 uppercase tracking-tighter">
                            {act.title}
                          </h3>
                          <p className="text-[15px] text-slate-500 dark:text-slate-400 truncate mb-1">
                            📍 {act.location}
                          </p>
                    <div className="flex items-center gap-2">
                      <span className="text-[15px] font-black text-blue-500 uppercase tracking-widest bg-blue-50 dark:bg-blue-900/30 px-2 py-0.5 rounded-md">
                        {act.category.split(',')[0]}
                      </span>
                      <span className="text-[15px] font-bold text-slate-400 flex items-center gap-1">
                        <Clock size={10} /> {act.time}
                      </span>
                    </div>
                        </div>
                        <ChevronRight size={14} className="text-slate-300 group-hover:translate-x-1 transition-transform" />
                      </div>
                    ))
                  ) : (
                    <div className="col-span-full py-12 text-center opacity-30 flex flex-col items-center">
                      <ActivityIcon size={40} strokeWidth={1} />
                      <p className="text-[10px] uppercase font-black tracking-widest mt-3">Hiện chưa có hoạt động</p>
                    </div>
                  )}
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                <AuditCharts />
                <WastewaterChart />
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        <ReportModal 
          isOpen={isReportModalOpen} 
          onClose={() => setIsReportModalOpen(false)} 
          targets={allTargets}
          onSave={handleSaveReport}
        />
        <TargetModal
          isOpen={isTargetModalOpen}
          onClose={() => { setIsTargetModalOpen(false); setEditingTarget(null); }}
          onSave={(data) => { setPendingAction({ type: 'save', data }); setIsPasswordModalOpen(true); }}
          editingTarget={editingTarget}
        />
        <BulkTargetModal
          isOpen={isBulkTargetModalOpen}
          onClose={() => setIsBulkTargetModalOpen(false)}
          onSave={(data) => {
            const newTargets = data.map(t => ({ ...t, id: `t-${Date.now()}-${Math.random().toString(36).substr(2, 9)}` }));
            setManualTargets([...newTargets, ...manualTargets]);
            setIsBulkTargetModalOpen(false);
          }}
        />
        {isPasswordModalOpen && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <div className="w-full max-w-sm bg-white rounded-3xl p-8 text-center shadow-2xl">
              <h3 className="text-lg font-black uppercase mb-6">Xác thực quyền hạn</h3>
              <input 
                type="password" autoFocus placeholder="Mật khẩu (123456)" 
                value={passwordValue} onChange={(e) => setPasswordValue(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleConfirmAction()}
                className="w-full h-12 px-4 rounded-xl border border-slate-200 bg-slate-50 text-center text-sm font-black mb-6"
              />
              <div className="grid grid-cols-2 gap-4">
                <button onClick={() => setIsPasswordModalOpen(false)} className="py-3 bg-slate-100 rounded-xl text-xs font-black uppercase text-slate-500">Hủy</button>
                <button onClick={handleConfirmAction} className="py-3 bg-blue-600 rounded-xl text-xs font-black uppercase text-white">Xác nhận</button>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  };

  // Connection Test
  useEffect(() => {
    async function testConnection() {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
      } catch (error) {
        if(error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration.");
        }
      }
    }
    testConnection();
  }, []);

  // Auth Listener
  useEffect(() => {
    try {
      setPersistence(auth, browserLocalPersistence);
    } catch (e) {
      console.error("Persistence error:", e);
    }
    return onAuthStateChanged(auth, (u) => {
      setUser(u);
      setAuthReady(true);
    });
  }, []);
  
  useEffect(() => {
    if (!user) return;
    const q = query(collection(db, 'portal_items'), orderBy('order', 'asc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      setDynamicPortalItems(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() })));
    }, (error) => {
      console.error("Error fetching portal items:", error);
    });
    return () => unsubscribe();
  }, [user]);

  // Real-time Firestore Listener
  useEffect(() => {
    if (!user) return;

    const q = query(collection(db, 'files'), orderBy('createdAt', 'desc'));
    
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const files: UploadedFile[] = [];
      snapshot.forEach((doc) => {
        const data = doc.data();
        files.push({
          id: doc.id,
          ...data,
          date: data.createdAt ? (data.createdAt as Timestamp).toDate().toLocaleDateString('vi-VN') : data.date
        } as UploadedFile);
      });

      setUploadedFiles(files);
      
      // Select first file if none selected or if it's the initial load
      if (files.length > 0 && isInitialLoad) {
        setSelectedFile(files[0]);
        setIsInitialLoad(false);
      }

      // Show notification for new files (not added by current session)
      snapshot.docChanges().forEach((change) => {
        if (change.type === 'added' && !isInitialLoad) {
          const newDoc = change.doc.data();
          if (newDoc.uploaderId !== user.uid) {
            setNotification({ 
              message: `Văn bản mới: ${newDoc.name} (Tải lên bởi ${newDoc.uploaderName || 'người dùng'})`, 
              type: 'success' 
            });
            setTimeout(() => setNotification(null), 5000);
          }
        }
      });
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, 'files');
    });

    return () => unsubscribe();
  }, [user, isInitialLoad]);

  const handleLogin = async () => {
    try {
      const result = await signInWithPopup(auth, googleProvider);
      if (result.user) {
        setHasLoggedAccess(true);
        setUser(result.user);
        const credential = GoogleAuthProvider.credentialFromResult(result);
        if (credential) {
          setAccessToken(credential.accessToken || null);
        }
        
        // Log login after successful sign in
        const u = result.user;
        try {
          await addDoc(collection(db, 'activity_logs'), {
            userId: u.uid,
            userEmail: u.email,
            userName: u.displayName || u.email?.split('@')[0],
            action: 'Đăng nhập',
            details: 'Đăng nhập thành công vào hệ thống (Chrome/Manual)',
            timestamp: serverTimestamp()
          });
        } catch (logError) {
          console.warn("Could not log activity, but user is signed in:", logError);
        }
      }
    } catch (error: any) {
      console.error("Login failed:", error);
      let msg = "Đăng nhập thất bại.";
      if (error.code === 'auth/popup-blocked') {
        msg = "Trình duyệt đã chặn cửa sổ bật lên. Vui lòng cho phép popup để đăng nhập.";
      } else if (error.code === 'auth/popup-closed-by-user') {
        msg = "Bạn đã đóng cửa sổ đăng nhập.";
      } else if (error.code === 'auth/network-request-failed') {
        msg = "Lỗi kết nối mạng hoặc bị trình duyệt chặn (Vui lòng kiểm tra Cookie bên thứ ba trong Chrome).";
      } else {
        msg = `Lỗi: ${error.message}`;
      }
      setNotification({ message: msg, type: 'error' });
      setTimeout(() => setNotification(null), 5000);
    }
  };

  const uploadToGoogleDrive = async (file: File, folderId: string) => {
    if (!accessToken) throw new Error("Chưa có quyền truy cập Drive. Vui lòng đăng nhập lại.");

    const metadata = {
      name: file.name,
      parents: [folderId],
    };

    const formData = new FormData();
    formData.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
    formData.append('file', file);

    const response = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink,webContentLink', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
      },
      body: formData,
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || "Lỗi tải lên Google Drive");
    }

    return await response.json();
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0 || !user) return;
    
    setIsLoading(true);
    const file = files[0];
    
    try {
      if (!accessToken) {
        setNotification({ message: 'Vui lòng đăng nhập lại để cấp quyền Drive!', type: 'error' });
        setTimeout(() => setNotification(null), 3000);
        return;
      }

      // 1. Upload to Google Drive
      const driveData = await uploadToGoogleDrive(file, DRIVE_FOLDERS[uploadCategory]);
      
      // 2. Save metadata to Firestore (this triggers real-time updates for everyone)
      const uploaderName = user.displayName || user.email?.split('@')[0] || 'Unknown';
      const fileData = {
        name: file.name,
        type: file.type,
        size: (file.size / 1024 / 1024).toFixed(1) + ' MB',
        url: driveData.webContentLink || driveData.webViewLink || '', 
        driveId: driveData.id,
        category: uploadCategory,
        date: new Date().toLocaleDateString('vi-VN'),
        uploaderId: user.uid,
        uploaderName: uploaderName,
        readBy: [],
        createdAt: serverTimestamp()
      };

      console.log("Saving file metadata to Firestore:", fileData);
      await addDoc(collection(db, 'files'), fileData);
      
      // Activity Log
      await addDoc(collection(db, 'activity_logs'), {
        userId: user.uid,
        userName: uploaderName,
        action: 'upload_file',
        details: `Tải lên văn bản: ${file.name}`,
        timestamp: serverTimestamp()
      });
      
      setNotification({ message: 'Đã tải lên Google Drive và thông báo thành công!', type: 'success' });
      setTimeout(() => setNotification(null), 3000);
    } catch (error: any) {
      console.error(error);
      let errorMsg = error.message || 'Lỗi hệ thống';
      const isApiError = errorMsg.includes('Google Drive API has not been used');
      
      if (isApiError) {
        setNotification({ 
          message: 'LỖI: Bạn chưa bật Google Drive API. Hãy nhấn vào link trong console hoặc hướng dẫn để kích hoạt.', 
          type: 'error' 
        });
      } else {
        setNotification({ message: errorMsg, type: 'error' });
      }
      
      setTimeout(() => setNotification(null), isApiError ? 15000 : 3000);
      
      if (!isApiError) {
        handleFirestoreError(error, OperationType.WRITE, 'files');
      }
    } finally {
      setIsLoading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleDeleteFile = async (e: React.MouseEvent, fileId: string, fileName: string) => {
    e.stopPropagation();
    console.log("Delete triggered for:", fileId, fileName);
    setFileToDelete({ id: fileId, name: fileName });
  };

  const confirmDeleteFile = async () => {
    console.log("Confirming delete for:", fileToDelete);
    if (!fileToDelete || !isAdmin) return;
    
    const { id: fileId, name: fileName } = fileToDelete;

    try {
      await deleteDoc(doc(db, 'files', fileId));
      
      // Log activity
      await addDoc(collection(db, 'activity_logs'), {
        userId: user?.uid,
        userName: user?.displayName || user?.email?.split('@')[0],
        action: 'delete_file',
        details: `Đã xoá thông báo: ${fileName}`,
        timestamp: serverTimestamp()
      });

      if (selectedFile?.id === fileId) {
        setSelectedFile(null);
      }
      
      setNotification({ message: 'Đã xoá thông báo thành công!', type: 'success' });
      setTimeout(() => setNotification(null), 3000);
      setFileToDelete(null);
    } catch (error) {
      console.error(error);
      handleFirestoreError(error, OperationType.DELETE, 'files');
      setNotification({ message: 'Lỗi khi xoá thông báo!', type: 'error' });
      setTimeout(() => setNotification(null), 3000);
    }
  };

  useEffect(() => {
    if (selectedFile) {
      renderFile(selectedFile);
    }
  }, [selectedFile]);

  const unreadFilesCount = useMemo(() => {
    if (!user) return 0;
    return uploadedFiles.filter(f => !f.readBy?.includes(user.uid)).length;
  }, [uploadedFiles, user]);

  const markFileAsRead = useCallback(async (fileId: string) => {
    if (!user) return;
    const path = `files/${fileId}`;
    try {
      await updateDoc(doc(db, 'files', fileId), {
        readBy: arrayUnion(user.uid)
      });
    } catch (error) {
      console.error("Error marking file as read:", error);
      handleFirestoreError(error, OperationType.UPDATE, path);
    }
  }, [user]);

  const handleMarkAllAsRead = async () => {
    if (!user || unreadFilesCount === 0) return;
    try {
      setIsLoading(true);
      const promises = uploadedFiles
        .filter(f => !f.readBy?.includes(user.uid))
        .map(async f => {
          try {
            await updateDoc(doc(db, 'files', f.id), {
              readBy: arrayUnion(user.uid)
            });
          } catch (err) {
            handleFirestoreError(err, OperationType.UPDATE, `files/${f.id}`);
          }
        });
      await Promise.all(promises);
      setNotification({ message: 'Đã đánh dấu tất cả là đã xem!', type: 'success' });
      setTimeout(() => setNotification(null), 3000);
    } catch (error) {
      console.error("Error marking all as read:", error);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    if (mainView === 'repository' && selectedFile && user && !selectedFile.readBy?.includes(user.uid)) {
      markFileAsRead(selectedFile.id);
    }
  }, [selectedFile, user, markFileAsRead, mainView]);

  // Tự động dọn dẹp dữ liệu (Văn bản 10 ngày, Nhật ký 3 tháng)
  const cleanupOldData = useCallback(async () => {
    if (!user) return;
    try {
      const tenDaysAgo = new Date();
      tenDaysAgo.setDate(tenDaysAgo.getDate() - 10);

      const threeMonthsAgo = new Date();
      threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
      
      // Dọn dẹp Files (10 ngày)
      const oldFilesQuery = query(
        collection(db, 'files'),
        where('createdAt', '<', tenDaysAgo),
        limit(50)
      );
      const fileSnap = await getDocs(oldFilesQuery);
      if (!fileSnap.empty) {
        console.log(`🧹 Đang dọn dẹp ${fileSnap.size} văn bản cũ hơn 10 ngày...`);
        const deletePromises = fileSnap.docs.map(d => deleteDoc(d.ref));
        await Promise.all(deletePromises);
      }

      // Dọn dẹp Logs (3 tháng)
      const oldLogsQuery = query(
        collection(db, 'activity_logs'),
        where('timestamp', '<', threeMonthsAgo),
        limit(50)
      );
      const logSnap = await getDocs(oldLogsQuery);
      if (!logSnap.empty) {
        console.log(`🧹 Đang dọn dẹp ${logSnap.size} nhật ký cũ hơn 3 tháng...`);
        const deletePromises = logSnap.docs.map(d => deleteDoc(d.ref));
        await Promise.all(deletePromises);
      }
    } catch (error) {
      console.error("Lỗi khi dọn dẹp dữ liệu cũ:", error);
    }
  }, [user]);

  useEffect(() => {
    if (user) {
      // Đợi một chút sau khi load xong rồi mới dọn dẹp để không ảnh hưởng hiệu năng ban đầu
      const timer = setTimeout(() => {
        cleanupOldData();
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [user, cleanupOldData]);

  const renderFile = async (file: UploadedFile) => {
    setWordContent(null);
    if (!file.url) return;

    // For Word files, we still try Mammoth for a "native" feel if possible, 
    // but fall back to the iframe viewer for everything else including Excel and PDF
    if (file.type.includes('word') || file.name.endsWith('.docx')) {
      setIsLoading(true);
      try {
        const response = await fetch(file.url);
        const blob = await response.blob();
        const arrayBuffer = await blob.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer });
        setWordContent(result.value);
      } catch (err) {
        console.error("Mammoth failed, falling back to iframe:", err);
        setWordContent(null); // Fallback to iframe
      } finally {
        setIsLoading(false);
      }
    } else {
      setWordContent(null);
    }
  };

  const getPreviewUrl = (file: UploadedFile) => {
    if (!file.url) return '';
    
    // If it's a Google Drive file, use the official preview endpoint
    if (file.driveId) {
      return `https://drive.google.com/file/d/${file.driveId}/preview`;
    }
    
    // Fallback for other URLs: Google Docs Viewer
    return `https://docs.google.com/viewer?url=${encodeURIComponent(file.url)}&embedded=true`;
  };

  const openInNewTab = () => {
    if (!selectedFile?.url) return;
    if (selectedFile.driveId) {
      window.open(`https://drive.google.com/file/d/${selectedFile.driveId}/view`, '_blank');
    } else {
      window.open(selectedFile.url, '_blank');
    }
  };

  if (!authReady) {
    return (
      <div className="fixed inset-0 bg-[#111820] flex items-center justify-center">
        <div className="w-8 h-8 border-4 border-blue-500 border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  if (!user) {
    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#111820]">
        <div className="absolute inset-0 grid-bg opacity-20" />
        <motion.div 
          initial={{ scale: 0.9, opacity: 0 }} 
          animate={{ scale: 1, opacity: 1 }} 
          className="relative z-10 w-[400px] p-10 bg-[#141c24] border border-[#243547] text-center shadow-2xl"
        >
          <div className="text-4xl mb-6">🏢</div>
          <h1 className="text-2xl font-bold text-white tracking-[4px] mb-2 font-sans">IALY PORTAL</h1>
          <p className="text-[11px] text-[#4a6278] mb-10 font-bold uppercase tracking-widest leading-relaxed">Hệ thống quản lý văn bản nội bộ<br/>Nhà máy Thuỷ điện Ialy</p>
          <button 
            onClick={handleLogin}
            className="w-full py-4 bg-[#00a8ff] text-white font-bold tracking-[3px] uppercase hover:brightness-110 transition-all shadow-[0_0_20px_rgba(0,168,255,0.2)] flex items-center justify-center gap-3"
          >
            <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" alt="Google" className="w-5 h-5" />
            ĐĂNG NHẬP VỚI GOOGLE
          </button>
          <div className="mt-6 text-[10px] text-[#4a6278] uppercase tracking-widest font-bold">Phiên bản 2026.5.3</div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen overflow-hidden bg-[#f8fafc] font-sans selection:bg-blue-100 selection:text-blue-700">
      {/* Main Top Header Navigation */}
      <header className="h-[60px] bg-white border-b border-gray-200 flex items-center justify-between px-4 sm:px-8 shrink-0 z-50 shadow-sm overflow-hidden">
        <div className="flex items-center gap-3 sm:gap-4 shrink-0">
          <div className="w-8 h-8 sm:w-9 sm:h-9 bg-gradient-to-br from-blue-600 to-sky-500 text-white font-black rounded-lg flex items-center justify-center text-base sm:text-lg italic shadow-lg shadow-blue-500/20">IA</div>
          <div className="flex flex-col hidden xs:flex">
            <span className="text-[13px] sm:text-[15px] font-black tracking-tighter text-gray-900 leading-none">VHIALY</span>
            <span className="text-[8px] sm:text-[9px] font-bold text-blue-600 uppercase tracking-widest mt-1">Ứng dụng dùng chung</span>
          </div>
        </div>

        <div className="flex items-center gap-2 sm:gap-8 h-full overflow-x-auto no-scrollbar scroll-smooth px-2">
          <button 
            onClick={() => { setMainView('portal'); }}
            className={`flex items-center gap-1.5 sm:gap-2 h-full border-b-2 font-bold text-xs sm:text-sm transition-all px-1.5 sm:px-2 flex-nowrap whitespace-nowrap ${mainView === 'portal' ? 'text-blue-600 border-blue-600 bg-blue-50/50' : 'text-gray-400 border-transparent hover:text-gray-600'}`}
          >
            <LayoutDashboard size={14} className="sm:w-4 sm:h-4" />
            <span className="hidden xs:inline">CỔNG THÔNG TIN</span>
            <span className="xs:hidden">CỔNG</span>
          </button>
          <button 
            onClick={() => { setMainView('dashboard'); }}
            className={`flex items-center gap-1.5 sm:gap-2 h-full border-b-2 font-bold text-xs sm:text-sm transition-all px-1.5 sm:px-2 flex-nowrap whitespace-nowrap ${mainView === 'dashboard' ? 'text-blue-600 border-blue-600 bg-blue-50/50' : 'text-gray-400 border-transparent hover:text-gray-600'}`}
          >
            <BarChart3 size={14} className="sm:w-4 sm:h-4" />
            <span className="hidden xs:inline">DASHBOARD</span>
            <span className="xs:hidden">DASH</span>
          </button>
          <button 
            onClick={() => { setMainView('repository'); }}
            className={`flex items-center gap-1.5 sm:gap-2 h-full border-b-2 font-bold text-xs sm:text-sm transition-all px-1.5 sm:px-2 flex-nowrap whitespace-nowrap ${mainView === 'repository' ? 'text-blue-600 border-blue-600 bg-blue-50/50' : 'text-gray-400 border-transparent hover:text-gray-600'}`}
          >
            <FolderArchive size={14} className="sm:w-4 sm:h-4" /> 
            VĂN BẢN
            {unreadFilesCount > 0 && (
              <span className="w-4 h-4 sm:w-5 sm:h-5 rounded-full bg-rose-600 text-white text-[9px] sm:text-[10px] font-black flex items-center justify-center animate-pulse shadow-sm px-1 min-w-[16px] sm:min-w-[20px]">
                {unreadFilesCount}
              </span>
            )}
          </button>
          {isAdmin && (
            <button 
              onClick={() => { setMainView('admin'); }}
              className={`flex items-center gap-1.5 sm:gap-2 h-full border-b-2 font-bold text-xs sm:text-sm transition-all px-1.5 sm:px-2 flex-nowrap whitespace-nowrap ${mainView === 'admin' ? 'text-purple-600 border-purple-600 bg-purple-50/50' : 'text-gray-400 border-transparent hover:text-purple-600'}`}
            >
              <ShieldCheck size={14} className="sm:w-4 sm:h-4" /> 
              ADMIN
            </button>
          )}
        </div>

        <div className="flex items-center gap-2 sm:gap-4 shrink-0">
          <div className="hidden sm:flex items-center gap-3 bg-gray-50 px-3 py-1.5 rounded-full border border-gray-100">
            <div className="w-6 h-6 rounded-full bg-gradient-to-br from-green-500 to-emerald-400 flex items-center justify-center text-[10px] text-white font-bold">
              {user.email?.[0].toUpperCase()}
            </div>
            <span className="text-[11px] font-semibold text-gray-600 hidden md:block">{user.email}</span>
          </div>
          <button onClick={() => auth.signOut()} className="text-[#e63946] font-bold text-[10px] sm:text-xs uppercase tracking-widest hover:underline px-1 sm:px-2 whitespace-nowrap">Đăng xuất</button>
        </div>
      </header>

      <AnimatePresence mode="wait">
        {mainView === 'repository' ? (
              <motion.div 
                key="repository"
                initial={{ opacity: 0, scale: 0.98 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 1.02 }}
                className="flex-1 flex flex-col lg:flex-row overflow-hidden bg-[#f0f4f8]"
              >
                <aside className="w-full lg:w-[320px] h-[35%] lg:h-full bg-white border-b lg:border-b-0 lg:border-r border-gray-200 flex flex-col shadow-xl z-20 shrink-0">
                  <div className="p-3 border-b border-gray-100 flex flex-col gap-2">
                    <div className="flex items-center p-1 bg-gray-50 rounded-lg w-full">
                      <div className={cn(
                        "flex-1 px-4 py-1.5 rounded-md text-[13px] font-black shadow-sm text-center uppercase tracking-wider transition-all",
                        unreadFilesCount > 0 ? "bg-rose-600 text-white animate-pulse" : "bg-[#e8f0fe] text-blue-700"
                      )}>
                        Thông báo ({unreadFilesCount})
                      </div>
                    </div>
                    {unreadFilesCount > 0 && (
                      <button 
                        onClick={handleMarkAllAsRead}
                        className="text-[10px] font-black text-blue-600 uppercase tracking-widest hover:text-blue-700 transition-colors flex items-center justify-center gap-1.5 py-1"
                      >
                        <CheckCircle2 size={12} />
                        Đã xem tất cả
                      </button>
                    )}
                  </div>
                  
              

                  <div className="flex-1 overflow-y-auto bg-white custom-scrollbar">
                    {uploadedFiles.map(file => (
                      <div 
                        key={file.id} 
                        onClick={() => setSelectedFile(file)}
                        className={`p-4 border-b border-gray-100 cursor-pointer transition-all relative ${selectedFile?.id === file.id ? 'bg-[#f0f7ff]' : 'hover:bg-gray-50'}`}
                      >
                        {selectedFile?.id === file.id && <div className="absolute left-0 top-0 bottom-0 w-[4px] bg-blue-600" />}
                        <div className="flex justify-between items-start mb-1.5 pr-1">
                          <div className="flex items-center gap-2">
                            {file.category && (
                              <span className={`text-[9px] font-black px-1.5 py-0.5 rounded w-fit uppercase ${
                                file.category === 'KPI' ? 'bg-indigo-100 text-indigo-700' :
                                file.category === 'LHC' ? 'bg-indigo-100 text-indigo-700' :
                                file.category === 'SAFETY' ? 'bg-orange-100 text-orange-700' : 
                                file.category === 'ECONOMY' ? 'bg-emerald-100 text-emerald-700' :
                                'bg-gray-100 text-gray-600'
                              }`}>
                                {file.category === 'KPI' ? 'KPI Phân xưởng' :
                                  file.category === 'LHC' ? 'Lịch hành chính' :
                                 file.category === 'SAFETY' ? 'Biên bản phổ biến TNGT' : 
                                 file.category === 'ECONOMY' ? 'Biên bản phổ biến TNLĐ' :
                                 'Thông báo chung'}
                              </span>
                            )}
                            {user && !file.readBy?.includes(user.uid) && (
                              <div className="w-2 h-2 bg-rose-500 rounded-full shadow-[0_0_8px_#f43f5e] animate-pulse" />
                            )}
                          </div>
                          <div className="flex items-center gap-3 relative z-10">
                            <span className="text-[11px] text-gray-500">{file.date}</span>
                            {isAdmin && (
                              <button 
                                onClick={(e) => {
                                  e.preventDefault();
                                  e.stopPropagation();
                                  handleDeleteFile(e, file.id!, file.name);
                                }}
                                className="w-8 h-8 flex items-center justify-center text-gray-400 hover:text-red-600 hover:bg-red-50 transition-all rounded-full shrink-0 group/del cursor-pointer pointer-events-auto"
                                style={{ pointerEvents: 'auto' }}
                                title="Xoá thông báo"
                              >
                                <Trash2 size={15} className="group-hover/del:scale-110 transition-transform" />
                              </button>
                            )}
                          </div>
                        </div>
                        <div className="text-[13px] text-gray-700 font-semibold leading-snug mb-2 line-clamp-2">
                          {file.name.includes('Lich_doi_ca') ? `Trình Xuân An-Đăng ký đổi ca (Tham gia họp phân xưởng ngày ${file.date})` : file.name}
                        </div>
                        <div className="text-[11px] leading-relaxed text-blue-600/80 font-medium">
                          {file.name.endsWith('.xlsx') || file.name.endsWith('.xls') ? 'Bảng tính Excel' : null}
                        </div>
                      </div>
                    ))}
                  </div>

                  <div className="p-4 border-t border-gray-100 bg-gray-50 space-y-3">
                    <div className="flex flex-col gap-1">
                      <label className="text-[10px] font-black text-gray-400 uppercase ml-1">Danh mục lưu trữ</label>
                      <select 
                        value={uploadCategory} 
                        onChange={(e) => setUploadCategory(e.target.value as any)}
                        className="w-full p-2 bg-white border border-gray-200 rounded text-[11px] font-bold text-gray-700 outline-none focus:border-blue-500 transition-colors"
                      >
                        <option value="GENERAL">THÔNG BÁO CHUNG</option>
                        <option value="KPI">KPI PHÂN XƯỞNG</option>
                        <option value="LHC">LỊCH HÀNH CHÍNH</option>
                        <option value="SAFETY">BIÊN BẢN PHỔ BIẾN TNGT</option>
                        <option value="ECONOMY">BIÊN BẢN PHỔ BIẾN TNLĐ</option>
                      </select>
                    </div>
                    <button 
                      onClick={() => fileInputRef.current?.click()}
                      className="w-full py-2.5 bg-blue-600 text-white text-[12px] font-bold rounded flex items-center justify-center gap-2 shadow-sm hover:bg-blue-700 transition-all uppercase tracking-wider"
                    >
                      <Upload size={14}/> TẢI LÊN VĂN BẢN MỚI
                    </button>
                    <input type="file" ref={fileInputRef} onChange={handleFileUpload} className="hidden" accept=".pdf,.doc,.docx,.xls,.xlsx" />
                  </div>
                </aside>

                <main className="flex-1 flex flex-col bg-[#525659] relative overflow-hidden h-[65%] lg:h-full">
                  {selectedFile ? (
                    <>
                      <div className="bg-white p-4 px-6 border-b border-gray-200 flex flex-col gap-2 shrink-0">
                        <div className="flex items-center gap-4 text-[13px]">
                          <span className="text-gray-400 font-medium min-w-[90px]">File văn bản:</span>
                          <div className="flex items-center gap-4">
                            {selectedFile.name.endsWith('.pdf') ? (
                              <span className="flex items-center gap-1.5 text-red-600 font-semibold cursor-pointer border-b border-transparent hover:border-red-600">
                                <div className="w-4 h-4 bg-red-500 flex items-center justify-center text-[10px] text-white font-bold rounded-sm">P</div>
                                {selectedFile.name}
                              </span>
                            ) : (selectedFile.name.endsWith('.xlsx') || selectedFile.name.endsWith('.xls')) ? (
                              <span className="flex items-center gap-1.5 text-green-600 font-semibold cursor-pointer border-b border-transparent hover:border-green-600">
                                <div className="w-4 h-4 bg-green-600 flex items-center justify-center text-[10px] text-white font-bold rounded-sm">X</div>
                                {selectedFile.name}
                              </span>
                            ) : (
                              <span className="flex items-center gap-1.5 text-blue-600 font-semibold cursor-pointer border-b border-transparent hover:border-blue-600">
                                <div className="w-4 h-4 bg-blue-500 flex items-center justify-center text-[10px] text-white font-bold rounded-sm">W</div>
                                {selectedFile.name}
                              </span>
                            )}
                          </div>
                        </div>
                        <div className="flex items-center gap-4 text-[13px]">
                          <span className="text-gray-400 font-medium min-w-[90px]">Người tải:</span>
                          <span className="text-blue-600 font-bold uppercase">{selectedFile.uploaderName || 'Hệ thống'}</span>
                        </div>
                      </div>

                      <div className="bg-[#cbd5e1] min-h-[48px] py-2 lg:py-0 flex flex-col sm:flex-row items-center justify-between px-4 sm:px-6 shrink-0 border-b border-gray-300 shadow-sm gap-2">
                        <div className="flex items-center gap-3 sm:gap-5 text-gray-700 w-full sm:w-auto overflow-hidden">
                          <button className="hover:bg-gray-400/20 p-2 rounded-md transition-colors shrink-0"><Settings size={18} className="text-gray-600"/></button>
                          <div className="flex items-center gap-2 text-[12px] bg-white px-3 py-1 border border-gray-400 shadow-inner rounded-sm shrink-0">
                            <input type="text" value="1" className="w-6 text-center outline-none text-blue-600 font-bold" readOnly />
                            <span className="text-gray-400">trên 1</span>
                          </div>
                          <span className="text-[12px] sm:text-[13px] font-bold text-gray-800 ml-1 tracking-wide truncate max-w-[200px]">{selectedFile.name}</span>
                        </div>
                        <div className="flex items-center gap-2 sm:gap-3 w-full sm:w-auto justify-end">
                          <button onClick={openInNewTab} className="flex items-center gap-1.5 px-3 sm:px-4 py-1.5 bg-[#0061f2] text-white text-[10px] sm:text-[11px] font-black rounded-md hover:bg-blue-700 transition-all shadow-sm active:scale-95 uppercase tracking-tighter sm:tracking-normal">
                            <Maximize2 size={13}/> <span className="hidden sm:inline">MỞ TRONG TAB MỚI</span><span className="sm:hidden">TAB MỚI</span>
                          </button>
                          <div className="w-[1px] h-6 bg-gray-400/30 mx-0.5" />
                          <button className="p-1.5 sm:p-2 hover:bg-gray-400/30 rounded-md transition-colors"><Eye size={20} className="text-gray-600"/></button>
                          <a href={selectedFile.url} download className="p-1.5 sm:p-2 hover:bg-gray-400/30 rounded-md transition-colors"><Download size={20} className="text-gray-600"/></a>
                        </div>
                      </div>

                      <div className="flex-1 overflow-auto p-2 sm:p-8 flex justify-center bg-[#525659] custom-scrollbar">
                        <div className="w-full max-w-[950px] min-h-[500px] lg:min-h-[1200px] h-fit bg-white shadow-[0_10px_30px_rgba(0,0,0,0.3)] relative">
                          {isLoading ? (
                            <div className="absolute inset-0 flex flex-col items-center justify-center p-12 text-center bg-white z-50">
                              <div className="w-10 h-10 border-4 border-blue-600/30 border-t-blue-600 rounded-full animate-spin mb-4" />
                              <p className="text-gray-400 text-xs font-bold uppercase tracking-widest">Đang xử lý tài liệu...</p>
                            </div>
                          ) : wordContent ? (
                            <div className="p-16 prose prose-slate max-w-none prose-sm" dangerouslySetInnerHTML={{ __html: wordContent }} />
                          ) : selectedFile.url ? (
                            <div className="absolute inset-0 z-10 flex flex-col h-full bg-white">
                              <iframe src={getPreviewUrl(selectedFile)} className="w-full h-full border-none" title="Document Viewer" />
                            </div>
                          ) : (
                            <div className="absolute inset-0 flex flex-col items-center justify-center p-12 text-center bg-gray-50">
                              <AlertTriangle size={64} className="text-gray-200 mb-6" />
                              <h3 className="text-lg font-bold text-gray-400 uppercase tracking-widest mb-2">Không thể hiển thị trực tiếp</h3>
                              <p className="text-gray-400 text-xs max-w-xs font-medium">Vui lòng sử dụng tài liệu của bạn tải lên để xem nội dung.</p>
                              <button onClick={() => fileInputRef.current?.click()} className="mt-6 px-6 py-2 bg-blue-600 text-white rounded text-xs font-bold uppercase shadow-lg shadow-blue-500/20">
                                Tải file lên
                              </button>
                            </div>
                          )}
                        </div>
                      </div>
                    </>
                  ) : (
                    <div className="flex-1 flex flex-col items-center justify-center text-gray-400 p-20">
                      <div className="w-20 h-20 bg-gray-200/50 rounded-full flex items-center justify-center mb-6">
                        <FileSearch size={32} className="opacity-40" />
                      </div>
                      <p className="text-sm font-bold uppercase tracking-[2px] text-gray-400">Chọn văn bản từ danh sách để xem chi tiết</p>
                    </div>
                  )}
                </main>
              </motion.div>
        ) : mainView === 'dashboard' ? (
          <motion.div 
            key="dashboard"
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className="flex-1 overflow-y-auto bg-[#f8fafc] p-8 custom-scrollbar"
          >
            <div className="max-w-7xl mx-auto space-y-12 pb-16">
              {renderDashboardModule()}
            </div>
          </motion.div>
        ) : mainView === 'admin' ? (
          <motion.div 
            key="admin"
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className="flex-1 overflow-y-auto bg-[#f8fafc] p-8 custom-scrollbar"
          >
            <div className="max-w-7xl mx-auto pb-16">
              <AdminMonitor />
            </div>
          </motion.div>
        ) : (
        <motion.div 
          key="portal"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="flex-1 flex flex-col overflow-y-auto custom-scrollbar"
          style={{
            backgroundColor: '#d0dbe6',
            backgroundImage: `
              radial-gradient(circle at 50% -10%, rgba(255, 255, 255, 0.95) 0%, rgba(255, 255, 255, 0.5) 40%, rgba(0, 0, 0, 0.05) 90%),
              url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='160' height='80' viewBox='0 0 160 80'%3E%3Crect width='160' height='80' fill='%23c9d6e4'/%3E%3Cline x1='0' y1='39.5' x2='160' y2='39.5' stroke='%23adbccb' stroke-width='1.5'/%3E%3Cline x1='159.5' y1='0' x2='159.5' y2='40' stroke='%23adbccb' stroke-width='1.5'/%3E%3Cline x1='0' y1='79.5' x2='160' y2='79.5' stroke='%23adbccb' stroke-width='1.5'/%3E%3Cline x1='79.5' y1='40' x2='79.5' y2='80' stroke='%23adbccb' stroke-width='1.5'/%3E%3Cline x1='0' y1='0.5' x2='160' y2='0.5' stroke='%23f4f7fa' stroke-width='1.5'/%3E%3Cline x1='0.5' y1='0' x2='0.5' y2='40' stroke='%23f4f7fa' stroke-width='1.5'/%3E%3Cline x1='0' y1='40.5' x2='160' y2='40.5' stroke='%23f4f7fa' stroke-width='1.5'/%3E%3Cline x1='80.5' y1='40' x2='80.5' y2='80' stroke='%23f4f7fa' stroke-width='1.5'/%3E%3C/svg%3E")
            `,
            backgroundAttachment: 'fixed'
          }}
        >
                {/* Hero section */}
                <div className="bg-[#0f172a] py-8 px-6 relative overflow-hidden shrink-0">
                  <img 
                    src="https://i.ibb.co/jZ6dDJzT/z7116558150434-802a4bd8dff3b332930235031b93fc49.jpg" 
                    className="absolute inset-0 w-full h-full object-cover opacity-70 brightness-[1.1] contrast-[1.25] saturate-[1.05]" 
                    alt="Hero Background"
                    referrerPolicy="no-referrer"
                  />
                  <div className="absolute inset-0 bg-gradient-to-r from-slate-950/90 via-slate-900/40 to-transparent" />
                  <div className="relative z-10 max-w-7xl mx-auto">
                    <h2 className="text-2xl md:text-3xl font-serif text-white leading-tight mb-2">
                      TỔNG HỢP DỮ LIỆU <br />
                      <span className="text-sky-300">PHÂN XƯỞNG VẬN HÀNH IALY</span>
                    </h2>

                    <p className="text-gray-400 text-xs md:text-sm leading-relaxed mb-6 max-w-2xl font-medium">
                      NHÀ MÁY THUỶ ĐIỆN IALY & IALY MỞ RỘNG 
                    </p>

                    <div className="flex flex-wrap gap-2.5">
                      {[
                        { id: 'kt', label: 'Kỹ Thuật', icon: '⚙️', color: 'red' },
                        { id: 'qt', label: 'Quản Trị', icon: '📊', color: 'grn' },
                        { id: 'at', label: 'An Toàn', icon: '🛡️', color: 'blu' }
                      ].map((tab) => (
                        <button
                          key={tab.id}
                          onClick={() => setPortalTab(tab.id as any)}
                          className={`flex items-center gap-2 px-4 py-2 rounded-full text-xs font-bold text-white transition-all border ${
                            portalTab === tab.id 
                              ? `bg-ialy-${tab.color}/40 border-ialy-${tab.color} shadow-md ring-2 ring-ialy-${tab.color}/20` 
                              : 'bg-white/5 border-white/10 hover:bg-white/20'
                          }`}
                        >
                          <span className="text-lg">{tab.icon}</span>
                          {tab.label}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Content section */}
                <div className="p-6 max-w-7xl mx-auto w-full">
                  <AnimatePresence mode="wait">
                    {portalTab === 'kt' && (
                      <motion.div key="kt" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="space-y-4">
                        <div className="flex items-center gap-3 mb-4">
                          <div className="w-8 h-8 rounded-lg bg-red-500/10 flex items-center justify-center text-base text-red-600 shadow-sm border border-red-200/60 backdrop-blur-sm">⚙️</div>
                          <div className="flex flex-col">
                            
                            <h3 className="text-base font-extrabold text-slate-800 tracking-tight">Kỹ Thuật</h3>
                          </div>
                        </div>

                        {dynamicPortalItems.filter(i => i.category === 'kt').map((item) => (
                          <AccordionCard 
                            key={item.id} 
                            id={item.id} 
                            num={item.num} 
                            title={item.title} 
                            badge={item.badge} 
                            color={item.color} 
                            href={item.href}
                            isOpen={true}
                            onToggle={() => {}}
                          >
                            {item.subItems?.map((sub: any, sIdx: number) => {
                              const subId = `${item.id}-${sIdx}`;
                              if (sub.childItems && sub.childItems.length > 0) {
                                return (
                                  <SubAccordion 
                                    key={sIdx}
                                    id={subId}
                                    icon={sub.icon}
                                    label={sub.label}
                                    isOpen={true}
                                    onToggle={() => {}}
                                    color={item.color}
                                  >
                                    {sub.childItems.map((child: any, cIdx: number) => (
                                      <LinkRow key={cIdx} icon={child.icon} label={child.label} href={child.href} color={item.color} level={3} />
                                    ))}
                                  </SubAccordion>
                                );
                              }
                              return (
                                <LinkRow key={sIdx} icon={sub.icon} label={sub.label} href={sub.href} color={item.color} level={2} />
                              );
                            })}
                          </AccordionCard>
                        ))}
                      </motion.div>
                    )}

                    {portalTab === 'qt' && (
                      <motion.div key="qt" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="space-y-4">
                         <div className="flex items-center gap-3 mb-4">
                            <div className="w-8 h-8 rounded-lg bg-green-500/10 flex items-center justify-center text-base text-emerald-600 shadow-sm border border-green-200/60 backdrop-blur-sm">📊</div>
                            <div className="flex flex-col">
                              <h3 className="text-base font-extrabold text-slate-800 tracking-tight">Quản Trị</h3>
                            </div>
                          </div>

                          {dynamicPortalItems.filter(i => i.category === 'qt').map((item) => (
                            <AccordionCard 
                              key={item.id} 
                              id={item.id} 
                              num={item.num} 
                              title={item.title} 
                              badge={item.badge} 
                              color={item.color} 
                              href={item.href}
                              isOpen={true}
                              onToggle={() => {}}
                            >
                              {item.subItems?.map((sub: any, sIdx: number) => {
                                const subId = `${item.id}-${sIdx}`;
                                if (sub.childItems && sub.childItems.length > 0) {
                                  return (
                                    <SubAccordion 
                                      key={sIdx}
                                      id={subId}
                                      icon={sub.icon}
                                      label={sub.label}
                                      isOpen={true}
                                      onToggle={() => {}}
                                      color={item.color}
                                    >
                                      {sub.childItems.map((child: any, cIdx: number) => (
                                        <LinkRow key={cIdx} icon={child.icon} label={child.label} href={child.href} color={item.color} level={3} />
                                      ))}
                                    </SubAccordion>
                                  );
                                }
                                return (
                                  <LinkRow key={sIdx} icon={sub.icon} label={sub.label} href={sub.href} color={item.color} level={2} />
                                );
                              })}
                            </AccordionCard>
                          ))}
                      </motion.div>
                    )}

                    {portalTab === 'at' && (
                      <motion.div key="at" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="space-y-4">
                        <div className="flex items-center gap-3 mb-4">
                            <div className="w-8 h-8 rounded-lg bg-blue-500/10 flex items-center justify-center text-base text-blue-600 shadow-sm border border-blue-200/60 backdrop-blur-sm">🛡️</div>
                            <div className="flex flex-col">
                              <h3 className="text-base font-extrabold text-slate-800 tracking-tight">An Toàn</h3>
                            </div>
                        </div>

                        {dynamicPortalItems.filter(i => i.category === 'at').map((item) => (
                          <AccordionCard 
                            key={item.id} 
                            id={item.id} 
                            num={item.num} 
                            title={item.title} 
                            badge={item.badge} 
                            color={item.color} 
                            href={item.href}
                            isOpen={true}
                            onToggle={() => {}}
                          >
                            {item.subItems?.map((sub: any, sIdx: number) => {
                              const subId = `${item.id}-${sIdx}`;
                              if (sub.childItems && sub.childItems.length > 0) {
                                  return (
                                    <SubAccordion 
                                      key={sIdx}
                                      id={subId}
                                      icon={sub.icon}
                                      label={sub.label}
                                      isOpen={true}
                                      onToggle={() => {}}
                                      color={item.color}
                                    >
                                      {sub.childItems.map((child: any, cIdx: number) => (
                                        <LinkRow key={cIdx} icon={child.icon} label={child.label} href={child.href} color={item.color} level={3} />
                                      ))}
                                    </SubAccordion>
                                  );
                                }
                                return (
                                  <LinkRow key={sIdx} icon={sub.icon} label={sub.label} href={sub.href} color={item.color} level={2} />
                                );
                              })}
                            </AccordionCard>
                          ))}
                        </motion.div>
                    )}
                  </AnimatePresence>
                </div>

                <EvnLogoWidget />

                <footer className="mt-auto border-t border-slate-300/30 py-6 px-8 bg-white/40 backdrop-blur-md flex items-center justify-between shrink-0">
                  <span className="text-[10px] uppercase font-black text-slate-600 tracking-widest leading-none">Phân Xưởng Vận Hành Ialy</span>
                </footer>
              </motion.div>
            )}
          </AnimatePresence>

          <AnimatePresence>
            {notification && (
              <motion.div 
                initial={{ y: -50, opacity: 0 }} 
                animate={{ y: 20, opacity: 1 }} 
                exit={{ y: -50, opacity: 0 }}
                className={`fixed top-4 left-1/2 -translate-x-1/2 z-[200] px-6 py-3 rounded shadow-2xl flex items-center gap-3 border bg-white ${notification.type === 'success' ? 'border-green-100' : 'border-red-100'}`}
              >
                <div className={`w-8 h-8 rounded-full flex items-center justify-center ${notification.type === 'success' ? 'bg-green-50 text-green-600' : 'bg-red-50 text-red-600'}`}>
                  {notification.type === 'success' ? <CheckCircle2 size={18} /> : <AlertTriangle size={18} />}
                </div>
                <span className={`text-[11px] font-black tracking-widest uppercase ${notification.type === 'success' ? 'text-green-700' : 'text-red-700'}`}>
                  {notification.message}
                </span>
              </motion.div>
            )}
          </AnimatePresence>

          <AnimatePresence>
            {fileToDelete && (
              <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
                <motion.div 
                  initial={{ scale: 0.9, opacity: 0 }} 
                  animate={{ scale: 1, opacity: 1 }}
                  exit={{ scale: 0.9, opacity: 0 }}
                  className="w-full max-w-sm bg-white rounded-[2rem] p-8 text-center shadow-2xl border border-slate-100"
                >
                  <div className="w-16 h-16 bg-red-50 text-red-500 rounded-2xl flex items-center justify-center mx-auto mb-6">
                    <Trash2 size={32} />
                  </div>
                  <h3 className="text-xl font-black text-slate-900 mb-2 uppercase tracking-tight">Xác nhận xoá</h3>
                  <p className="text-sm text-slate-500 font-medium mb-8 leading-relaxed">
                    Bạn có chắc chắn muốn xoá thông báo <br/>
                    <span className="text-slate-900 font-bold">"{fileToDelete.name}"</span>?
                    <br/>Hành động này không thể hoàn tác.
                  </p>
                  <div className="grid grid-cols-2 gap-4">
                    <button 
                      onClick={() => setFileToDelete(null)}
                      className="py-3.5 bg-slate-100 rounded-2xl text-[11px] font-black uppercase text-slate-500 hover:bg-slate-200 transition-colors"
                    >
                      Hủy bỏ
                    </button>
                    <button 
                      onClick={confirmDeleteFile}
                      className="py-3.5 bg-red-600 rounded-2xl text-[11px] font-black uppercase text-white hover:bg-red-700 transition-all shadow-lg shadow-red-200"
                    >
                      Xoá vĩnh viễn
                    </button>
                  </div>
                </motion.div>
              </div>
            )}
          </AnimatePresence>
    </div>
  );
}

// ═══════════════════════════════════════
// UI COMPONENTS FOR THE PORTAL
// ═══════════════════════════════════════

function AccordionCard({ id, num, title, badge, color, children, isOpen, onToggle, href }: any) {
  const content = (
    <div 
      className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 px-4.5 py-3 bg-slate-100/50 border-b border-slate-200/60"
    >
      <div className="flex items-center gap-4">
        <span className="text-[13px] font-extrabold text-slate-800 tracking-tight uppercase">{title}</span>
      </div>
    </div>
  );

  return (
    <div className="bg-white/95 backdrop-blur-md border border-slate-300 rounded-2xl overflow-hidden shadow-md group">
      {href ? (
        <a href={href} target="_blank" rel="noopener noreferrer" className="block hover:bg-slate-50/50 transition-colors">
          {content}
        </a>
      ) : (
        <>
          {content}
          <div className="p-3.5 flex flex-col gap-3 bg-slate-50/10">
            {children}
          </div>
        </>
      )}
    </div>
  );
}

function LinkRow({ icon, label, href, color, level = 2 }: any) {
  const borderColors = { red: 'border-l-red-500', grn: 'border-l-emerald-500', blu: 'border-l-blue-500' };
  
  if (level === 3) {
    return (
      <a 
        href={href} 
        target="_blank" 
        rel="noopener noreferrer"
        className="group relative flex items-center gap-2.5 bg-white hover:bg-slate-50 border border-slate-200/70 p-2.5 px-3 rounded-lg hover:translate-x-0.5 hover:shadow-xs transition-all relative z-10"
      >
        <span className="text-sm shrink-0">{icon}</span>
        <span className="flex-1 text-[11.5px] font-medium text-slate-600 group-hover:text-blue-600 transition-colors leading-snug">{label}</span>
        
        <ExternalLink size={11} className="text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity" />
      </a>
    );
  }

  return (
    <a 
      href={href} 
      target="_blank" 
      rel="noopener noreferrer"
      className={`flex items-center gap-3 bg-white border border-slate-200/80 border-l-[3px] ${borderColors[color as 'red'|'grn'|'blu']} p-3 rounded-lg hover:translate-x-0.5 hover:bg-slate-50 hover:shadow-xs transition-all group`}
    >
      <span className="text-base shrink-0">{icon}</span>
      <span className="flex-1 text-[12px] font-bold text-slate-800 leading-snug group-hover:text-blue-600 transition-colors">{label}</span>
      
      <ExternalLink size={12} className="text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity" />
    </a>
  );
}

function SubAccordion({ icon, label, children, isOpen, onToggle, color }: any) {
  const borderLeftColors = { red: 'border-l-red-500', grn: 'border-l-emerald-500', blu: 'border-l-blue-500' };
  
  return (
    <div className={`border border-slate-200 rounded-lg overflow-hidden bg-slate-50/40 shadow-xs border-l-[3px] ${borderLeftColors[color as 'red'|'grn'|'blu']}`}>
      <div 
        className="flex items-center justify-between px-3 py-2 bg-slate-100/70 border-b border-slate-200/50"
      >
        <div className="flex items-center gap-2.5">
          <span className="text-base shrink-0">{icon}</span>
          <span className="text-[12px] font-bold text-slate-800 leading-none">{label}</span>
        </div>
      </div>
      <div className="p-2.5 pr-2 flex flex-col gap-1.5 bg-white/40 border-t border-slate-200/30 relative">
        {/* Tree vertical timeline branch guide */}
        <div className="absolute left-[21px] top-4 bottom-6 w-[1px] bg-slate-200" />
        {children}
      </div>
    </div>
  );
}

function SubLink({ label, href, color }: any) {
  const dotColors = { red: 'bg-red-500', grn: 'bg-emerald-500', blu: 'bg-blue-500' };
  const textColors = { red: 'group-hover:text-red-700', grn: 'group-hover:text-emerald-700', blu: 'group-hover:text-blue-700' };
  return (
    <a 
      href={href} 
      target="_blank" 
      rel="noopener noreferrer"
      className="flex items-center gap-3 px-5 py-3 border-b border-slate-100 last:border-none hover:bg-slate-50 hover:pl-7 transition-all group"
    >
      <div className={`w-1.5 h-1.5 rounded-full ${dotColors[color as 'red'|'grn'|'blu']}`} />
      <span className={`flex-1 text-[12px] font-black text-slate-700 ${textColors[color as 'red'|'grn'|'blu']} transition-colors`}>{label}</span>
      <ExternalLink size={10} className="text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity" />
    </a>
  );
}

function DirectCard({ icon, title, sub, href, color }: any) {
  const bgColors = { red: 'bg-red-50 text-red-600', grn: 'bg-green-50 text-green-600', blu: 'bg-blue-50 text-blue-600' };
  return (
    <a 
      href={href} 
      target="_blank" 
      rel="noopener noreferrer"
      className="flex items-center gap-4 bg-white/85 border border-slate-200/60 p-5 rounded-2xl shadow-md hover:shadow-xl hover:-translate-y-0.5 hover:border-slate-300/80 transition-all group"
    >
      <div className={`w-12 h-12 rounded-2xl flex items-center justify-center text-2xl shrink-0 ${bgColors[color as 'red'|'grn'|'blu']}`}>
        {icon}
      </div>
      <div className="flex flex-col flex-1">
        <span className="text-[15px] font-black text-slate-800 leading-none mb-1.5">{title}</span>
        <span className="text-[11px] font-bold text-slate-400 uppercase tracking-tighter">{sub}</span>
      </div>
      <ExternalLink size={16} className="text-slate-400 opacity-0 group-hover:opacity-100 transition-opacity" />
    </a>
  );
}
