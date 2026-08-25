import React, { useState, useRef } from 'react';
import { API_BASE } from '../utils/api';
import { 
  Upload, Trash2, CheckCircle2, Search
} from 'lucide-react';

interface SignatureManagerProps {
  staffList: string[];
  signatures: Record<string, string>;
  onSignaturesChange: (signatures: Record<string, string>) => void;
  workshopId: string;
}

// Utility function to normalize strings for comparison (strips accents, lowercases)
function normalizeName(str: string): string {
  return str
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

export default function SignatureManager({ staffList, signatures, onSignaturesChange, workshopId }: SignatureManagerProps) {
  const [saving, setSaving] = useState<string | null>(null);
  const [isBatchProcessing, setIsBatchProcessing] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [notice, setNotice] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);

  const bulkFileInputRef = useRef<HTMLInputElement>(null);

  const missingStaffList = staffList.filter(name => !signatures[name]);
  const hasSignaturesCount = staffList.filter(name => signatures[name]).length;

  const filteredStaffList = staffList.filter(name => 
    name.toLowerCase().includes(searchTerm.toLowerCase().trim())
  );

  // Single file upload
  const handleSingleUpload = async (name: string, file: File) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      const base64 = e.target?.result as string;
      setSaving(name);
      try {
        const res = await fetch(API_BASE + '/api/signatures', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name, data: base64, workshopId })
        });
        if (res.ok) {
          onSignaturesChange({ ...signatures, [name]: base64 });
          setNotice({ type: 'success', text: `Đã cập nhật chữ ký cho "${name}"!` });
        }
      } catch (err) {
        console.error("Failed to save signature", err);
        setNotice({ type: 'error', text: `Lỗi lưu chữ ký cho "${name}".` });
      } finally {
        setSaving(null);
      }
    };
    reader.readAsDataURL(file);
  };

  // Bulk File Upload matching filenames to staff names
  const handleBulkUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setIsBatchProcessing(true);
    setNotice({ type: 'info', text: `Đang xử lý ${files.length} file ảnh chữ ký...` });

    const matchedItems: { name: string; data: string }[] = [];
    const unmatchedFiles: string[] = [];

    // Map staff normalized names to real names
    const staffNormalizedMap: Record<string, string> = {};
    staffList.forEach(name => {
      staffNormalizedMap[normalizeName(name)] = name;
    });

    const fileList: File[] = Array.from(files);
    const filePromises = fileList.map((file: File) => {
      return new Promise<void>((resolve) => {
        const rawFileName = file.name.substring(0, file.name.lastIndexOf('.')) || file.name;
        const normFileName = normalizeName(rawFileName);

        let matchedName = staffNormalizedMap[normFileName];

        // Fuzzy fallback match if exact normalized match isn't found
        if (!matchedName) {
          matchedName = staffList.find(sName => {
            const sNorm = normalizeName(sName);
            return normFileName.includes(sNorm) || sNorm.includes(normFileName);
          }) || '';
        }

        if (matchedName) {
          const reader = new FileReader();
          reader.onload = (evt) => {
            const base64 = evt.target?.result as string;
            matchedItems.push({ name: matchedName, data: base64 });
            resolve();
          };
          reader.readAsDataURL(file);
        } else {
          unmatchedFiles.push(file.name);
          resolve();
        }
      });
    });

    await Promise.all(filePromises);

    if (matchedItems.length > 0) {
      try {
        const res = await fetch(API_BASE + '/api/signatures/batch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ items: matchedItems, workshopId })
        });

        if (res.ok) {
          const newSigs = { ...signatures };
          matchedItems.forEach(item => {
            newSigs[item.name] = item.data;
          });
          onSignaturesChange(newSigs);

          setNotice({
            type: 'success',
            text: `✅ Đã khớp và tải lên thành công ${matchedItems.length} chữ ký!` + 
                  (unmatchedFiles.length > 0 ? ` (${unmatchedFiles.length} file không khớp tên nhân sự)` : '')
          });
        }
      } catch (err) {
        setNotice({ type: 'error', text: 'Lỗi lưu danh sách chữ ký hàng loạt.' });
      }
    } else {
      setNotice({ 
        type: 'error', 
        text: '⚠ Không có file ảnh nào khớp tên với danh sách nhân sự. Vui lòng đặt tên file trùng tên nhân sự (Ví dụ: "Nguyen Van A.png").' 
      });
    }

    setIsBatchProcessing(false);
    e.target.value = '';
  };

  // Delete signature
  const handleDeleteSignature = async (name: string) => {
    if (!confirm(`Bạn có chắc muốn xóa chữ ký của "${name}"?`)) return;
    setSaving(name);
    try {
      const newSigs = { ...signatures };
      delete newSigs[name];
      onSignaturesChange(newSigs);
      
      await fetch(API_BASE + '/api/signatures', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, data: '', workshopId })
      });
      setNotice({ type: 'info', text: `Đã xóa chữ ký của "${name}".` });
    } catch (e) {
      console.error(e);
    } finally {
      setSaving(null);
    }
  };

  return (
    <div className="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 text-slate-800 space-y-4">
      {/* Header & Stats */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3.5">
        <div>
          <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
            <span className="p-1.5 bg-blue-50 text-blue-600 rounded-lg"><Upload size={18} /></span>
            Quản lý chữ ký nhân sự ({staffList.length} người)
          </h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Đã có chữ ký: <strong className="text-emerald-600">{hasSignaturesCount}</strong> / {staffList.length} người 
            {missingStaffList.length > 0 && (
              <span className="ml-2 text-amber-600 font-medium">(Còn thiếu {missingStaffList.length} người)</span>
            )}
          </p>
        </div>

        {/* Quick Action Toolbar */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Bulk File Upload Button */}
          <button
            type="button"
            disabled={isBatchProcessing}
            onClick={() => bulkFileInputRef.current?.click()}
            className="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-xl shadow-sm transition-all flex items-center gap-2 cursor-pointer disabled:opacity-50"
            title="Chọn nhiều file ảnh chữ ký cùng lúc (Tên file = Tên nhân viên)"
          >
            <Upload size={14} /> Tải ảnh chữ ký hàng loạt (Nhiều File)
          </button>
          <input
            ref={bulkFileInputRef}
            type="file"
            multiple
            accept="image/*"
            className="hidden"
            onChange={handleBulkUpload}
          />
        </div>
      </div>

      {/* Notice Alert */}
      {notice && (
        <div className={`p-3 rounded-xl text-xs flex items-center justify-between border ${
          notice.type === 'success' ? 'bg-emerald-50 border-emerald-200 text-emerald-800 font-medium' :
          notice.type === 'error' ? 'bg-rose-50 border-rose-200 text-rose-800 font-medium' :
          'bg-blue-50 border-blue-200 text-blue-800 font-medium'
        }`}>
          <span>{notice.text}</span>
          <button type="button" onClick={() => setNotice(null)} className="text-slate-400 hover:text-slate-600 font-bold ml-2 cursor-pointer">✕</button>
        </div>
      )}

      {/* Search Filter bar */}
      <div className="flex items-center gap-2 bg-slate-50 p-2 rounded-xl border border-slate-200">
        <Search size={15} className="text-slate-400 ml-1" />
        <input 
          type="text" 
          placeholder="Tìm tên nhân viên..." 
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          className="bg-transparent border-none text-xs text-slate-700 focus:outline-none w-full"
        />
        {searchTerm && (
          <button type="button" onClick={() => setSearchTerm('')} className="text-slate-400 text-xs hover:text-slate-600 mr-1 cursor-pointer">Xóa tìm kiếm</button>
        )}
      </div>

      {/* Staff Grid Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 max-h-[520px] overflow-y-auto pr-1">
        {filteredStaffList.map(name => {
          const hasSig = Boolean(signatures[name]);
          const isSavingThis = saving === name;

          return (
            <div 
              key={name} 
              className={`p-3 border rounded-xl flex flex-col justify-between gap-2.5 transition-all ${
                hasSig ? 'bg-white border-slate-200 shadow-2xs' : 'bg-slate-50/70 border-dashed border-slate-300'
              }`}
            >
              <div className="flex justify-between items-center">
                <span className="font-semibold text-xs text-slate-800 truncate max-w-[170px]" title={name}>
                  {name}
                </span>
                {hasSig ? (
                  <span className="inline-flex items-center gap-1 text-[11px] font-bold bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-full">
                    <CheckCircle2 size={11} /> Đã có
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1 text-[11px] font-semibold bg-amber-100 text-amber-800 px-2 py-0.5 rounded-full">
                    Chưa có
                  </span>
                )}
              </div>
              
              {/* Signature Display Box */}
              <div className="h-16 bg-slate-50 border border-slate-200 rounded-lg flex items-center justify-center p-1 overflow-hidden relative group">
                {hasSig ? (
                  <>
                    <img src={signatures[name]} alt={`Chữ ký ${name}`} className="max-h-full max-w-full object-contain" />
                    <button
                      type="button"
                      onClick={() => handleDeleteSignature(name)}
                      className="absolute top-1 right-1 p-1 bg-rose-500 hover:bg-rose-600 text-white rounded-md opacity-0 group-hover:opacity-100 transition-opacity cursor-pointer"
                      title="Xóa chữ ký này"
                    >
                      <Trash2 size={12} />
                    </button>
                  </>
                ) : (
                  <span className="text-[11px] text-slate-400 italic">Chưa cài đặt chữ ký</span>
                )}
              </div>

              {/* Upload Image Action */}
              <div>
                <label className={`
                  w-full cursor-pointer text-center py-1.5 px-2 rounded-lg text-[11px] font-semibold transition-all flex items-center justify-center gap-1.5 border
                  ${isSavingThis 
                    ? 'bg-slate-200 text-slate-500 border-slate-200 cursor-not-allowed' 
                    : 'bg-slate-50 hover:bg-blue-50 text-slate-700 hover:text-blue-700 border-slate-200 hover:border-blue-200'}
                `} title="Tải ảnh chữ ký từ máy tính">
                  <Upload size={12} /> {hasSig ? 'Thay đổi ảnh chữ ký' : 'Tải ảnh chữ ký'}
                  <input 
                    type="file" 
                    className="hidden" 
                    accept="image/*" 
                    onChange={(e) => e.target.files?.[0] && handleSingleUpload(name, e.target.files[0])}
                    disabled={isSavingThis}
                  />
                </label>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}