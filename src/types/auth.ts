export type UserRole = 'super_admin' | 'workshop_admin' | 'workshop_user';

export interface UserAccount {
  id: string;
  username: string;
  passwordHash?: string; // or plain password for demo
  password?: string;
  fullName: string;
  role: UserRole;
  workshopId: string; // 'all' for super_admin or specific workshop ID like 'px_vanhanh'
  createdAt: string;
  updatedAt?: string;
}

export interface WorkshopFeatures {
  enablePhanCa: boolean;
  enableDonNghiPhep: boolean;
  enableDoiCa: boolean;
  enableChuKySo: boolean;
  enableGoogleSheets: boolean;
  enableBaoComCa: boolean;
}

export interface ShiftOptionConfig {
  code: string;
  name: string;
  timeRange: string;
}

export interface ShiftScheduleConfig {
  shiftCa1Name?: string;
  shiftCa1Time?: string;
  shiftCa2Name?: string;
  shiftCa2Time?: string;
  shiftCa3Name?: string;
  shiftCa3Time?: string;
  teams?: string[];
  cycleLengthDays?: number;
  shiftNote?: string;
  baseDate?: string;
  shiftsMatrix?: string[][];
  rulesMatrix?: Record<number, Record<string, { k: number }>>;
}

export interface WorkshopConfig {
  companyName?: string;
  headerWorkshopName?: string;
  documentCodeSuffix?: string;
  recipientWorkshopName?: string;
  shortWorkshopName?: string;
  locationName?: string;
  soVanBan: string;
  ngayKy: string;
  nguoiKy: string;
  chucVuNguoiKy?: string;
  zaloWebhookUrl: string;
  notifyEmail?: string;
  shiftSchedule?: ShiftScheduleConfig;
}

export interface Workshop {
  id: string;
  name: string;
  code: string;
  description: string;
  staffData: string[][];
  config: WorkshopConfig;
  features: WorkshopFeatures;
  createdAt: string;
  updatedAt: string;
}

export interface AuthState {
  isAuthenticated: boolean;
  user: UserAccount | null;
  currentWorkshop: Workshop | null;
  allWorkshops: Workshop[];
}
