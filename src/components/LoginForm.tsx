import React, { useState, useEffect } from 'react';
import { User, Lock, ArrowRight, Eye, EyeOff, Sparkles, UserPlus, Building2, X } from 'lucide-react';
import { UserAccount, Workshop } from '../types/auth';

interface LoginFormProps {
  onLoginSuccess: (user: UserAccount) => void;
  isEmbedded?: boolean;
}

export default function LoginForm({ onLoginSuccess, isEmbedded = false }: LoginFormProps) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [rememberMe, setRememberMe] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  // Modal State for Quick Admin Creation & Setup
  const [showCreateAdminModal, setShowCreateAdminModal] = useState(false);
  const [newAdminUsername, setNewAdminUsername] = useState('');
  const [newAdminPassword, setNewAdminPassword] = useState('123456');
  const [newAdminFullName, setNewAdminFullName] = useState('');
  const [newWsOption, setNewWsOption] = useState<'new' | 'existing'>('new');
  const [newWsName, setNewWsName] = useState('');
  const [newWsCode, setNewWsCode] = useState('');
  const [newWsDesc, setNewWsDesc] = useState('');
  const [selectedWsId, setSelectedWsId] = useState('');
  const [existingWorkshops, setExistingWorkshops] = useState<Workshop[]>([]);
  const [loadingWorkshops, setLoadingWorkshops] = useState(false);
  const [creatingAdmin, setCreatingAdmin] = useState(false);
  const [modalMsg, setModalMsg] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem('remembered_login');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.username) setUsername(parsed.username);
        if (parsed.password) setPassword(parsed.password);
        if (typeof parsed.rememberMe === 'boolean') setRememberMe(parsed.rememberMe);
      }
    } catch (e) {
      console.error('Lỗi đọc tài khoản đã lưu:', e);
    }
  }, []);

  const fetchExistingWorkshops = async () => {
    setLoadingWorkshops(true);
    try {
      const res = await fetch('/api/workshops');
      const data = await res.json();
      if (Array.isArray(data)) {
        setExistingWorkshops(data);
        if (data.length > 0 && !selectedWsId) {
          setSelectedWsId(data[0].id);
        }
      }
    } catch (err) {
      console.error('Lỗi tải danh sách phân xưởng:', err);
    } finally {
      setLoadingWorkshops(false);
    }
  };

  const handleCreateAdminAccount = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newAdminUsername.trim() || !newAdminFullName.trim()) {
      setModalMsg({ type: 'error', text: 'Vui lòng nhập Tên đăng nhập và Họ tên Quản trị viên.' });
      return;
    }

    setCreatingAdmin(true);
    setModalMsg(null);

    try {
      let targetWsId = selectedWsId;

      if (newWsOption === 'new') {
        if (!newWsName.trim()) {
          throw new Error('Vui lòng nhập Tên Phân xưởng mới.');
        }

        const wsRes = await fetch('/api/workshops', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: newWsName.trim(),
            code: newWsCode.trim() || newWsName.substring(0, 5).toUpperCase(),
            description: newWsDesc.trim() || `Phân xưởng ${newWsName.trim()}`,
            config: {
              companyName: 'CÔNG TY THỦY ĐIỆN IALY',
              headerWorkshopName: newWsName.trim().toUpperCase(),
              documentCodeSuffix: newWsCode ? `/${newWsCode.trim()}` : '/PX',
              recipientWorkshopName: newWsName.trim(),
              shortWorkshopName: newWsName.trim(),
              soVanBan: '01/PX',
              nguoiKy: 'Lãnh đạo Phân xưởng',
              chucVuNguoiKy: 'Quản đốc Phân xưởng',
              zaloWebhookUrl: 'https://vhialy.dpdns.org/webhook/notify'
            }
          })
        });

        const wsData = await wsRes.json();
        if (!wsRes.ok || !wsData.success) {
          throw new Error(wsData.error || 'Lỗi khi tạo Phân xưởng mới.');
        }
        targetWsId = wsData.workshop.id;
      } else {
        if (!targetWsId && existingWorkshops.length > 0) {
          targetWsId = existingWorkshops[0].id;
        }
        if (!targetWsId) {
          throw new Error('Vui lòng chọn một Phân xưởng hợp lệ.');
        }
      }

      // Step 2: Create Admin Account
      const accRes = await fetch('/api/accounts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          username: newAdminUsername.trim(),
          password: newAdminPassword.trim() || '123456',
          fullName: newAdminFullName.trim(),
          role: 'workshop_admin',
          workshopId: targetWsId
        })
      });

      const accData = await accRes.json();
      if (!accRes.ok || !accData.success) {
        throw new Error(accData.error || 'Lỗi khi tạo tài khoản Admin Phân xưởng.');
      }

      // Step 3: Auto login as newly created Admin account
      const loginRes = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          username: newAdminUsername.trim(),
          password: newAdminPassword.trim() || '123456'
        })
      });

      const loginData = await loginRes.json();
      if (!loginRes.ok || !loginData.success || !loginData.user) {
        throw new Error('Đã tạo tài khoản thành công! Vui lòng đăng nhập với mật khẩu vừa tạo.');
      }

      localStorage.setItem('auth_user', JSON.stringify(loginData.user));
      localStorage.setItem('remembered_login', JSON.stringify({
        username: newAdminUsername.trim(),
        password: newAdminPassword.trim() || '123456',
        rememberMe: true
      }));

      setShowCreateAdminModal(false);
      onLoginSuccess(loginData.user);
    } catch (err: any) {
      setModalMsg({ type: 'error', text: err.message || 'Có lỗi xảy ra khi tạo tài khoản Admin.' });
    } finally {
      setCreatingAdmin(false);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!username.trim() || !password.trim()) {
      setError('Vui lòng nhập tên đăng nhập và mật khẩu.');
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: username.trim(), password: password.trim() })
      });

      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Tên đăng nhập hoặc mật khẩu không chính xác.');
      }

      if (rememberMe) {
        localStorage.setItem(
          'remembered_login',
          JSON.stringify({ username: username.trim(), password: password.trim(), rememberMe: true })
        );
      } else {
        localStorage.removeItem('remembered_login');
      }

      if (data.user) {
        localStorage.setItem('auth_user', JSON.stringify(data.user));
      }
      onLoginSuccess(data.user);
    } catch (err: any) {
      setError(err.message || 'Đã có lỗi xảy ra. Vui lòng thử lại.');
    } finally {
      setLoading(false);
    }
  };

  const formCardContent = (
    <div className="w-full bg-white rounded-3xl shadow-2xl border border-slate-100 overflow-hidden">
      {/* Header with EVN Logo & Portal Title */}
      <div className="p-8 pb-6 text-center border-b border-slate-100 relative bg-gradient-to-b from-slate-50/80 to-white">
        <div className="flex justify-center mb-3">
          <img 
            src="https://i.ibb.co/b5p84s5f/ai-preview.png" 
            alt="EVN Ialy Logo" 
            className="h-32 md:h-36 w-auto object-contain drop-shadow-sm mb-1" 
          />
        </div>
        <h2 className="text-xl md:text-2xl font-black text-[#00529c] tracking-tight uppercase">
          ỨNG DỤNG TẠO LỊCH TRỰC THAY CA VẬN HÀNH
        </h2>
      </div>

      {/* Form Body */}
      <form onSubmit={handleSubmit} className="p-8 space-y-5">
        {error && (
          <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-2xl text-[15px] text-rose-700 font-semibold flex items-center gap-2 animate-in fade-in duration-200">
            <span className="w-2 h-2 rounded-full bg-rose-500 animate-pulse shrink-0"></span>
            <span>{error}</span>
          </div>
        )}

        <div>
          <label className="block text-[15px] font-bold text-slate-700 uppercase tracking-wider mb-2">
            Tên đăng nhập
          </label>
          <div className="relative">
            <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
              <User size={18} />
            </div>
            <input
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="Nhập tên đăng nhập tài khoản..."
              className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-[15px] font-semibold text-slate-800 placeholder:text-slate-400 focus:bg-white focus:outline-none focus:ring-2 focus:ring-[#00529c]/20 focus:border-[#00529c] transition-all"
              required
            />
          </div>
        </div>

        <div>
          <div className="flex items-center justify-between mb-2">
            <label className="block text-[15px] font-bold text-slate-700 uppercase tracking-wider">
              Mật khẩu
            </label>
          </div>
          <div className="relative">
            <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
              <Lock size={18} />
            </div>
            <input
              type={showPassword ? 'text' : 'password'}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Nhập mật khẩu truy cập..."
              className="w-full pl-10 pr-11 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-[15px] font-semibold text-slate-800 placeholder:text-slate-400 focus:bg-white focus:outline-none focus:ring-2 focus:ring-[#00529c]/20 focus:border-[#00529c] transition-all"
              required
            />
            <button
              type="button"
              onClick={() => setShowPassword(!showPassword)}
              className="absolute inset-y-0 right-0 pr-3.5 flex items-center text-slate-400 hover:text-slate-600 cursor-pointer"
            >
              {showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
            </button>
          </div>
        </div>

        <div className="flex items-center justify-between text-[15px] pt-1">
          <label 
            onClick={() => setRememberMe(!rememberMe)}
            className="inline-flex items-center gap-2.5 cursor-pointer text-slate-700 font-semibold select-none group"
          >
            <div className={`w-5 h-5 rounded-lg border-2 flex items-center justify-center transition-all ${
              rememberMe 
                ? 'bg-[#00529c] border-[#00529c] text-white shadow-sm' 
                : 'bg-slate-50 border-slate-300 group-hover:border-[#00529c]'
            }`}>
              {rememberMe && (
                <svg className="w-3.5 h-3.5 stroke-[3]" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                </svg>
              )}
            </div>
            <span className="group-hover:text-[#00529c] transition-colors">Ghi nhớ đăng nhập</span>
          </label>
        </div>

        <button
          type="submit"
          disabled={loading}
          className="w-full py-3.5 px-6 bg-[#00529c] hover:bg-[#003d75] active:scale-[0.98] text-white text-[15px] font-extrabold uppercase tracking-wider rounded-2xl shadow-lg shadow-[#00529c]/25 transition-all flex items-center justify-center gap-2 disabled:opacity-50 cursor-pointer"
        >
          {loading ? (
            <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
          ) : (
            <>
              <span>ĐĂNG NHẬP HỆ THỐNG</span>
              <ArrowRight size={18} />
            </>
          )}
        </button>

        {/* Option to create a new Workshop Admin Account */}
        <div className="pt-2 text-center border-t border-slate-100/80 mt-3">
          <button
            type="button"
            onClick={() => {
              setShowCreateAdminModal(true);
              setModalMsg(null);
              fetchExistingWorkshops();
            }}
            className="inline-flex items-center gap-2 text-[14px] font-bold text-[#00529c] hover:text-[#003d75] hover:underline cursor-pointer transition-all py-1.5 px-3 rounded-xl hover:bg-slate-50"
          >
            <Sparkles size={16} className="text-amber-500 animate-pulse shrink-0" />
            <span>Tạo tài khoản Admin Phân xưởng </span>
          </button>
        </div>
      </form>

      {/* Footer */}
      <div className="bg-slate-50 px-8 py-3.5 border-t border-slate-100 text-center text-[13px] text-slate-400 font-medium">
        © 2026 CÔNG TY THỦY ĐIỆN IALY - TẬP ĐOÀN ĐIỆN LỰC VIỆT NAM
      </div>

      {/* Quick Admin Creation Modal */}
      {showCreateAdminModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/70 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-hidden my-8 animate-in zoom-in-95 duration-150">
            {/* Modal Header */}
            <div className="p-5 bg-gradient-to-r from-[#00529c] to-teal-700 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <Sparkles size={20} className="text-amber-300 animate-pulse shrink-0" />
                <h3 className="font-extrabold text-base tracking-tight">Tạo Tài khoản Admin Phân Xưởng Mới</h3>
              </div>
              <button
                type="button"
                onClick={() => setShowCreateAdminModal(false)}
                className="p-1.5 rounded-full hover:bg-white/20 text-white transition-colors cursor-pointer"
              >
                <X size={18} />
              </button>
            </div>

            {/* Modal Content */}
            <form onSubmit={handleCreateAdminAccount} className="p-6 space-y-4 text-xs">
              {modalMsg && (
                <div className={`p-3.5 rounded-2xl text-xs font-bold border ${
                  modalMsg.type === 'success' ? 'bg-emerald-50 text-emerald-800 border-emerald-200' : 'bg-rose-50 text-rose-800 border-rose-200'
                }`}>
                  {modalMsg.type === 'success' ? '✅ ' : '⚠️ '}{modalMsg.text}
                </div>
              )}

              {/* Section 1: Admin Account Details */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3">
                <div className="flex items-center gap-2 font-bold text-slate-800 text-sm border-b border-slate-200 pb-2">
                  <UserPlus size={16} className="text-[#00529c]" />
                  <span>1. Thông tin Tài khoản Quản trị</span>
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Tên đăng nhập * <span className="text-slate-400 font-normal">(dùng để đăng nhập)</span>
                  </label>
                  <input
                    type="text"
                    value={newAdminUsername}
                    onChange={(e) => setNewAdminUsername(e.target.value)}
                    placeholder="ví dụ: admin_vh_pleikrong"
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#00529c]/20"
                    required
                  />
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Mật khẩu * <span className="text-slate-400 font-normal">(mặc định: 123456)</span>
                    </label>
                    <input
                      type="text"
                      value={newAdminPassword}
                      onChange={(e) => setNewAdminPassword(e.target.value)}
                      placeholder="123456"
                      className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-mono font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#00529c]/20"
                      required
                    />
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">
                      Họ và tên Quản trị viên *
                    </label>
                    <input
                      type="text"
                      value={newAdminFullName}
                      onChange={(e) => setNewAdminFullName(e.target.value)}
                      placeholder="ví dụ: Nguyễn Văn Quản Lý"
                      className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#00529c]/20"
                      required
                    />
                  </div>
                </div>
              </div>

              {/* Section 2: Workshop Option */}
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3">
                <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                  <div className="flex items-center gap-2 font-bold text-slate-800 text-sm">
                    <Building2 size={16} className="text-[#00529c]" />
                    <span>2. Chọn hoặc Tạo Phân xưởng</span>
                  </div>

                  <div className="flex items-center gap-1 bg-white p-1 rounded-xl border border-slate-200">
                    <button
                      type="button"
                      onClick={() => setNewWsOption('new')}
                      className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                        newWsOption === 'new' ? 'bg-[#00529c] text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
                      }`}
                    >
                      ➕ Phân xưởng mới
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setNewWsOption('existing');
                        fetchExistingWorkshops();
                      }}
                      className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                        newWsOption === 'existing' ? 'bg-[#00529c] text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
                      }`}
                    >
                      🏢 Chọn sẵn
                    </button>
                  </div>
                </div>

                {newWsOption === 'new' ? (
                  <div className="space-y-3">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Tên Phân xưởng mới *</label>
                      <input
                        type="text"
                        value={newWsName}
                        onChange={(e) => setNewWsName(e.target.value)}
                        placeholder="ví dụ: Phân xưởng Vận hành Pleikrông"
                        className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#00529c]/20"
                        required={newWsOption === 'new'}
                      />
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Mã Phân xưởng</label>
                        <input
                          type="text"
                          value={newWsCode}
                          onChange={(e) => setNewWsCode(e.target.value)}
                          placeholder="ví dụ: VHPK"
                          className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold uppercase text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#00529c]/20"
                        />
                      </div>

                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">Mô tả ngắn</label>
                        <input
                          type="text"
                          value={newWsDesc}
                          onChange={(e) => setNewWsDesc(e.target.value)}
                          placeholder="Nhà máy thủy điện..."
                          className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#00529c]/20"
                        />
                      </div>
                    </div>
                  </div>
                ) : (
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Danh sách Phân xưởng *</label>
                    {loadingWorkshops ? (
                      <div className="py-3 text-slate-400 font-medium text-center">Đang tải danh sách phân xưởng...</div>
                    ) : (
                      <select
                        value={selectedWsId}
                        onChange={(e) => setSelectedWsId(e.target.value)}
                        className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#00529c]/20"
                      >
                        {existingWorkshops.map(ws => (
                          <option key={ws.id} value={ws.id}>
                            {ws.name} ({ws.code})
                          </option>
                        ))}
                      </select>
                    )}
                  </div>
                )}
              </div>

              {/* Submit */}
              <div className="pt-2 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setShowCreateAdminModal(false)}
                  className="px-4 py-2.5 rounded-xl border border-slate-200 font-bold text-slate-600 hover:bg-slate-100 transition-all cursor-pointer"
                >
                  Hủy bỏ
                </button>
                <button
                  type="submit"
                  disabled={creatingAdmin}
                  className="px-5 py-2.5 bg-[#00529c] hover:bg-[#003d75] disabled:opacity-50 text-white font-extrabold rounded-xl shadow-md transition-all cursor-pointer flex items-center gap-2"
                >
                  {creatingAdmin ? (
                    <span>Đang tạo tài khoản & đăng nhập...</span>
                  ) : (
                    <>
                      <Sparkles size={16} />
                      <span>Tạo Admin & Đăng Nhập Thiết Lập</span>
                    </>
                  )}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );

  if (isEmbedded) {
    return <div className="w-full max-w-md mx-auto">{formCardContent}</div>;
  }

  return (
    <div className="min-h-screen w-full bg-slate-900 flex items-center justify-center p-4 relative overflow-hidden">
      <div className="w-full max-w-md mx-auto">
        {formCardContent}
      </div>
    </div>
  );
}


