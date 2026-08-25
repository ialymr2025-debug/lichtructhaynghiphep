import React from 'react';
import { UserAccount, Workshop } from '../types/auth';
import { ShieldCheck, LogOut, Settings, Factory, ChevronDown, User, Sparkles, PlusCircle } from 'lucide-react';

interface UserHeaderBarProps {
  user: UserAccount;
  workshops: Workshop[];
  activeWorkshop: Workshop | null;
  onSelectWorkshop: (workshop: Workshop) => void;
  onOpenWorkshopManager: () => void;
  onLogout: () => void;
}

export default function UserHeaderBar({
  user,
  workshops,
  activeWorkshop,
  onSelectWorkshop,
  onOpenWorkshopManager,
  onLogout
}: UserHeaderBarProps) {
  const isSuperAdmin = user.role === 'super_admin';
  const isWorkshopAdmin = user.role === 'workshop_admin';
  const canManageWorkshops = isSuperAdmin || isWorkshopAdmin;

  // Filter available workshops for selector
  const availableWorkshops = isSuperAdmin
    ? workshops
    : workshops.filter(w => w.id === user.workshopId);

  return (
    <header className="bg-slate-900 text-white border-b border-slate-800 sticky top-0 z-40 shadow-lg">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 py-2.5 flex flex-wrap items-center justify-between gap-3">
        {/* Left: Brand & Workshop Badge */}
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2 bg-gradient-to-r from-teal-600 to-emerald-600 px-3 py-1.5 rounded-xl shadow-md">
            <Factory className="w-4 h-4 text-emerald-100" />
            <span className="font-bold text-sm tracking-wide text-white">HỆ THỐNG PHÂN XƯỞNG</span>
          </div>

          {/* Active Workshop Selector */}
          {activeWorkshop && (
            <div className="relative group">
              {availableWorkshops.length > 1 ? (
                <div className="flex items-center gap-2 bg-slate-800/80 hover:bg-slate-800 border border-slate-700/80 rounded-xl px-3 py-1.5 transition-all">
                  <span className="text-xs font-semibold text-emerald-400">Phân xưởng:</span>
                  <select
                    value={activeWorkshop.id}
                    onChange={(e) => {
                      const found = workshops.find(w => w.id === e.target.value);
                      if (found) onSelectWorkshop(found);
                    }}
                    className="bg-transparent text-xs font-bold text-white focus:outline-none cursor-pointer pr-2"
                  >
                    {availableWorkshops.map(ws => (
                      <option key={ws.id} value={ws.id} className="bg-slate-900 text-slate-100">
                        {ws.name} ({ws.code})
                      </option>
                    ))}
                  </select>
                  <ChevronDown size={14} className="text-slate-400" />
                </div>
              ) : (
                <div className="flex items-center gap-2 bg-slate-800/60 border border-slate-700/60 rounded-xl px-3 py-1.5">
                  <span className="text-xs font-semibold text-slate-400">Phân xưởng:</span>
                  <span className="text-xs font-bold text-emerald-300">{activeWorkshop.name}</span>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Right: User Profile & Actions */}
        <div className="flex items-center gap-2.5">
          {/* Admin Management Button */}
          {canManageWorkshops && (
            <button
              onClick={onOpenWorkshopManager}
              className="flex items-center gap-1.5 text-xs font-bold px-3 py-1.5 bg-emerald-500 text-slate-950 hover:bg-emerald-400 rounded-xl transition-all cursor-pointer shadow-md active:scale-95"
              title="Quản lý cấu hình chức năng & tài khoản phân xưởng"
            >
              <Settings size={15} className="animate-spin-slow" />
              <span className="inline">Quản lý Phân xưởng & Tài khoản</span>
            </button>
          )}

          {/* User Role Badge & Name */}
          <div className="flex items-center gap-2 bg-slate-800/80 border border-slate-700/80 px-3 py-1.5 rounded-xl text-xs">
            <div className="w-6 h-6 rounded-lg bg-teal-500/20 flex items-center justify-center text-teal-300 font-bold">
              <User size={14} />
            </div>
            <div>
              <div className="font-bold text-slate-100 flex items-center gap-1.5">
                <span>{user.fullName}</span>
                <span className={`text-[11px] font-extrabold px-1.5 py-0.2 rounded-md ${
                  isSuperAdmin 
                    ? 'bg-amber-500/20 text-amber-300 border border-amber-500/30' 
                    : isWorkshopAdmin 
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30' 
                      : 'bg-slate-700 text-slate-300'
                }`}>
                  {isSuperAdmin ? 'Admin Tổng' : isWorkshopAdmin ? 'Admin PX' : 'Tài khoản PX'}
                </span>
              </div>
            </div>
          </div>

          {/* Logout Button */}
          <button
            onClick={onLogout}
            className="flex items-center gap-1.5 text-xs font-semibold px-2.5 py-1.5 bg-rose-500/10 hover:bg-rose-500/20 text-rose-300 border border-rose-500/20 rounded-xl transition-all cursor-pointer active:scale-95"
            title="Đăng xuất"
          >
            <LogOut size={14} />
            <span className="hidden md:inline">Đăng xuất</span>
          </button>
        </div>
      </div>
    </header>
  );
}