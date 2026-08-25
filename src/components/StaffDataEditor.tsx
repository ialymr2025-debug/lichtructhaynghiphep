import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Table, Plus, Trash2, FileSpreadsheet, Code, Check, AlertCircle, Copy, 
  Sparkles, Search, RefreshCw, Download, Upload, CheckCircle2, FileText 
} from 'lucide-react';

interface StaffDataEditorProps {
  staffDataText: string;
  onChangeStaffDataText: (text: string) => void;
  teamsListText?: string; // e.g. "Kíp 1, Kíp 2, Kíp 3, Kíp 4, Kíp 5"
  onSyncLeaves?: () => void;
  isSyncingLeaves?: boolean;
}

export default function StaffDataEditor({
  staffDataText,
  onChangeStaffDataText,
  teamsListText = "Kíp 1, Kíp 2, Kíp 3, Kíp 4, Kíp 5",
  onSyncLeaves,
  isSyncingLeaves = false
}: StaffDataEditorProps) {
  const [activeMode, setActiveMode] = useState<'table' | 'excel'>('table');
  const [searchFilter, setSearchFilter] = useState('');
  const [parseError, setParseError] = useState<string | null>(null);
  const [uploadNotice, setUploadNotice] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [importOption, setImportOption] = useState<'replace' | 'append'>('replace');

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Derive dynamic teams list (e.g. ['Kíp 1', 'Kíp 2', 'Kíp 3', 'Kíp 4', 'Kíp 5'])
  const teams = useMemo(() => {
    const list = teamsListText.split(',').map(t => t.trim()).filter(Boolean);
    return list.length > 0 ? list : ['Kíp 1', 'Kíp 2', 'Kíp 3', 'Kíp 4', 'Kíp 5'];
  }, [teamsListText]);

  // Parse current staffDataText safely into 2D array
  const rows: string[][] = useMemo(() => {
    try {
      if (!staffDataText || !staffDataText.trim()) return [];
      const parsed = JSON.parse(staffDataText);
      if (Array.isArray(parsed)) {
        setParseError(null);
        return parsed.map(row => Array.isArray(row) ? row.map(cell => String(cell || '')) : [String(row)]);
      }
      setParseError('Dữ liệu không phải là mảng.');
      return [];
    } catch (e: any) {
      setParseError('Mã JSON hiện tại chưa đúng định dạng. Bạn có thể tự chỉnh trong tab "Mã JSON" hoặc dán lại dữ liệu mới.');
      return [];
    }
  }, [staffDataText]);

  // Helper to update rows state and emit JSON text
  const updateRows = (newRows: string[][]) => {
    const jsonString = JSON.stringify(newRows, null, 2);
    onChangeStaffDataText(jsonString);
  };

  // Cell change handler
  const handleCellChange = (rowIndex: number, colIndex: number, value: string) => {
    const newRows = rows.map((r, rIdx) => {
      if (rIdx !== rowIndex) return [...r];
      const newRow = [...r];
      // Ensure row is long enough
      while (newRow.length <= colIndex) {
        newRow.push('');
      }
      newRow[colIndex] = value;
      return newRow;
    });
    updateRows(newRows);
  };

  // Add new position row
  const handleAddRow = () => {
    const defaultRow = ['Chức danh mới', ...teams.map(() => '')];
    updateRows([...rows, defaultRow]);
  };

  // Remove row
  const handleRemoveRow = (rowIndex: number) => {
    const newRows = rows.filter((_, idx) => idx !== rowIndex);
    updateRows(newRows);
  };

  // --- DOWNLOAD EXCEL TEMPLATE (.xlsx) ---
  const handleDownloadTemplate = () => {
    const headerRow = ['STT', 'Chức danh / Vị trí', ...teams];
    let templateRows: (string | number)[][] = [];

    if (rows && rows.length > 0) {
      templateRows = rows.map((r, idx) => [
        idx + 1,
        r[0] || '',
        ...teams.map((_, tIdx) => r[tIdx + 1] || '')
      ]);
    } else {
      templateRows = [
        [1, 'Trưởng ca', 'Nguyễn Tiến Danh', 'Tạ Văn Hà', 'Nguyễn Văn Trường', 'Lê Văn Dũng', 'Phạm Văn Nam'],
        [2, 'Trực TTĐK', 'Vũ Đức Cường', 'Lê Trí Dũng', 'Nguyễn Văn Toàn', 'Phan Văn Hùng', 'Trần Văn Bình'],
        [3, 'Trực chính điện', 'Phan Văn Hùng', 'Ngô Xuân Đoàn', 'Nguyễn Yên Nam', 'Lê Hữu Nghĩa', 'Hoàng Văn Tuấn'],
        [4, 'Kỹ thuật viên', 'Trần Đình Nam', 'Đỗ Hoàng Anh', 'Phạm Minh Đức', 'Bùi Văn Tiến', 'Đặng Quốc Huy']
      ];
    }

    const worksheetData = [headerRow, ...templateRows];
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

    // Set formatting column widths
    worksheet['!cols'] = [
      { wch: 6 },  // STT
      { wch: 24 }, // Chức danh
      ...teams.map(() => ({ wch: 20 }))
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Danh_Sach_Nhan_Su');

    XLSX.writeFile(workbook, 'Mau_Danh_Sach_Nhan_Su_Phan_Xuong.xlsx');
  };

  // --- EXPORT CURRENT LIST TO EXCEL (.xlsx) ---
  const handleExportCurrentList = () => {
    if (rows.length === 0) {
      alert('Chưa có danh sách nhân sự để xuất file Excel!');
      return;
    }

    const headerRow = ['STT', 'Chức danh / Vị trí', ...teams];
    const exportData = rows.map((r, idx) => [
      idx + 1,
      r[0] || '',
      ...teams.map((_, tIdx) => r[tIdx + 1] || '')
    ]);

    const worksheet = XLSX.utils.aoa_to_sheet([headerRow, ...exportData]);
    worksheet['!cols'] = [
      { wch: 6 },
      { wch: 22 },
      ...teams.map(() => ({ wch: 20 }))
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Danh_Sach_Nhan_Su');

    XLSX.writeFile(workbook, 'Danh_Sach_Nhan_Su_Hien_Tai.xlsx');
  };

  // --- UPLOAD & PARSE EXCEL FILE ---
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setUploadNotice(null);
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const workbook = XLSX.read(buffer, { type: 'array' });
        
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          throw new Error('File Excel không chứa Sheet dữ liệu nào.');
        }

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 }) as any[][];

        if (!rawRows || rawRows.length === 0) {
          setUploadNotice({ type: 'error', text: 'File Excel tải lên trống hoặc không có nội dung.' });
          return;
        }

        const parsedRows: string[][] = [];

        for (let i = 0; i < rawRows.length; i++) {
          const rawRow = rawRows[i];
          if (!Array.isArray(rawRow) || rawRow.length === 0) continue;

          const cleanRow = rawRow.map(cell => String(cell ?? '').trim());

          // Skip completely empty rows
          if (cleanRow.every(c => !c)) continue;

          const firstCell = (cleanRow[0] || '').toLowerCase();
          const secondCell = (cleanRow[1] || '').toLowerCase();

          // Skip Header row (e.g. STT, Chức danh, Kíp 1...)
          if (
            firstCell.includes('stt') || 
            firstCell.includes('chức danh') || 
            firstCell.includes('vị trí') ||
            secondCell.includes('chức danh') ||
            secondCell.includes('vị trí')
          ) {
            continue;
          }

          // Check if column 0 is STT (number like 1, 2, 3...)
          let titleIdx = 0;
          if (/^\d+$/.test(cleanRow[0]) && cleanRow[1]) {
            titleIdx = 1;
          }

          const jobTitle = cleanRow[titleIdx];
          if (!jobTitle) continue;

          // Personnel for teams start after jobTitle
          const personnel = cleanRow.slice(titleIdx + 1);
          parsedRows.push([jobTitle, ...personnel]);
        }

        if (parsedRows.length === 0) {
          setUploadNotice({ 
            type: 'error', 
            text: 'Không tìm thấy danh sách nhân sự hợp lệ. Hãy bấm "Tải File Excel Mẫu" để xem chuẩn định dạng.' 
          });
          return;
        }

        if (importOption === 'append') {
          const merged = [...rows, ...parsedRows];
          updateRows(merged);
          setUploadNotice({ 
            type: 'success', 
            text: `🎉 Đã nối thêm thành công ${parsedRows.length} chức danh mới từ file Excel vào danh sách hiện tại!` 
          });
        } else {
          updateRows(parsedRows);
          setUploadNotice({ 
            type: 'success', 
            text: `🎉 Đã tải thành công ${parsedRows.length} chức danh từ file Excel!` 
          });
        }

        setActiveMode('table');
      } catch (err: any) {
        console.error("Excel import error:", err);
        setUploadNotice({ 
          type: 'error', 
          text: 'Lỗi đọc file Excel: ' + (err?.message || 'File không đúng định dạng XLSX/CSV.') 
        });
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = ''; // Reset input
  };



  // Filtered rows for table search
  const filteredRowsWithIndices = useMemo(() => {
    return rows.map((row, originalIndex) => ({ row, originalIndex })).filter(({ row }) => {
      if (!searchFilter.trim()) return true;
      const query = searchFilter.toLowerCase();
      return row.some(cell => cell.toLowerCase().includes(query));
    });
  }, [rows, searchFilter]);

  return (
    <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 shadow-sm space-y-3">
      {/* Hidden File Input for Excel Import */}
      <input 
        type="file" 
        ref={fileInputRef} 
        onChange={handleFileUpload} 
        accept=".xlsx, .xls, .csv" 
        className="hidden" 
      />

      {/* Top Header & Mode Switcher */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 pb-3 border-b border-slate-200">
        <div>
          <label className="font-bold text-slate-800 text-sm flex items-center gap-2">
            <span>👥 6. Danh sách Nhân sự & Chức danh Phân xưởng</span>
            <span className="text-xs font-semibold px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded-full">
              {rows.length} Chức danh
            </span>
          </label>
          <p className="text-xs text-slate-500 mt-0.5">
            Quản lý tên chức danh và nhân sự từng kíp. Có thể tải file Excel mẫu hoặc tải lên file Excel.
          </p>
        </div>

        {/* Quick Excel Buttons */}
        <div className="flex flex-wrap items-center gap-2">
          {onSyncLeaves && (
            <button
              type="button"
              onClick={onSyncLeaves}
              disabled={isSyncingLeaves}
              className="px-3 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all cursor-pointer shadow-xs disabled:opacity-50"
              title="Khởi tạo danh sách nhân sự và cấp mặc định 12 ngày phép lên Google Sheet"
            >
              <RefreshCw size={14} className={isSyncingLeaves ? 'animate-spin' : ''} />
              <span>{isSyncingLeaves ? 'Đang khởi tạo...' : 'Khởi tạo sheet Số ngày phép'}</span>
            </button>
          )}

          <button
            type="button"
            onClick={handleDownloadTemplate}
            className="px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer shadow-2xs"
            title="Tải file Excel mẫu chuẩn (.xlsx) về máy"
          >
            <Download size={14} className="text-emerald-600" />
            <span>Tải Excel Mẫu</span>
          </button>

          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer shadow-xs"
            title="Tải lên file Excel (.xlsx, .csv) chứa danh sách nhân sự"
          >
            <Upload size={14} />
            <span>Upload File Excel</span>
          </button>
        </div>
      </div>

      {/* Mode Selector Tabs */}
      <div className="flex items-center gap-1 bg-slate-200/80 p-1 rounded-xl overflow-x-auto">
        <button
          type="button"
          onClick={() => setActiveMode('table')}
          className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer shrink-0 ${
            activeMode === 'table'
              ? 'bg-white text-emerald-700 shadow-sm'
              : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <Table size={14} />
          <span>Bảng nhập liệu ({rows.length})</span>
        </button>

        <button
          type="button"
          onClick={() => setActiveMode('excel')}
          className={`px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all cursor-pointer shrink-0 ${
            activeMode === 'excel'
              ? 'bg-white text-emerald-700 shadow-sm'
              : 'text-slate-600 hover:text-slate-900'
          }`}
        >
          <FileSpreadsheet size={14} />
          <span>Nhập / Tải File Excel</span>
        </button>
      </div>

      {/* Upload / Success / Error Banners */}
      {uploadNotice && (
        <div className={`p-3 rounded-xl text-xs flex items-center justify-between gap-2 border ${
          uploadNotice.type === 'success' 
            ? 'bg-emerald-50 border-emerald-200 text-emerald-900' 
            : 'bg-rose-50 border-rose-200 text-rose-900'
        }`}>
          <div className="flex items-center gap-2">
            {uploadNotice.type === 'success' ? (
              <CheckCircle2 size={16} className="text-emerald-600 shrink-0" />
            ) : (
              <AlertCircle size={16} className="text-rose-600 shrink-0" />
            )}
            <span className="font-medium">{uploadNotice.text}</span>
          </div>
          <button 
            type="button" 
            onClick={() => setUploadNotice(null)}
            className="text-slate-400 hover:text-slate-600 cursor-pointer text-xs font-bold px-1"
          >
            ✕
          </button>
        </div>
      )}

      {/* Parse Error Notice */}
      {parseError && activeMode !== 'json' && (
        <div className="p-3 bg-amber-50 border border-amber-200 text-amber-800 rounded-xl text-xs flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <AlertCircle size={16} className="text-amber-600 shrink-0" />
            <span>{parseError}</span>
          </div>
          <button
            type="button"
            onClick={() => setActiveMode('json')}
            className="px-2.5 py-1 bg-amber-200 hover:bg-amber-300 text-amber-900 rounded-lg text-[11px] font-bold cursor-pointer transition-all"
          >
            Sửa trong JSON
          </button>
        </div>
      )}

      {/* MODE 1: VISUAL TABLE EDITOR */}
      {activeMode === 'table' && (
        <div className="space-y-3">
          {/* Table Controls */}
          <div className="flex items-center justify-between gap-2">
            <div className="relative flex-1 max-w-xs">
              <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                placeholder="Tìm vị trí hoặc tên..."
                value={searchFilter}
                onChange={(e) => setSearchFilter(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 text-xs bg-white border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-emerald-500"
              />
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleExportCurrentList}
                className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-medium flex items-center gap-1 transition-all cursor-pointer"
                title="Xuất bảng danh sách hiện tại ra file Excel"
              >
                <Download size={13} />
                <span className="hidden sm:inline">Xuất Excel</span>
              </button>

              <button
                type="button"
                onClick={handleAddRow}
                className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-medium flex items-center gap-1.5 shadow-xs transition-all cursor-pointer"
              >
                <Plus size={14} />
                <span>Thêm Chức Danh / Vị Trí</span>
              </button>
            </div>
          </div>

          {/* Table Rendering */}
          <div className="overflow-x-auto border border-slate-200 rounded-xl bg-white shadow-xs max-h-[380px] overflow-y-auto">
            <table className="w-full text-xs text-left border-collapse min-w-[650px]">
              <thead className="sticky top-0 bg-slate-100 text-slate-700 font-bold border-b border-slate-200 z-10">
                <tr>
                  <th className="py-2.5 px-3 w-10 text-center text-slate-400">#</th>
                  <th className="py-2.5 px-3 min-w-[160px]">Chức danh / Vị trí</th>
                  {teams.map((teamName, tIdx) => (
                    <th key={tIdx} className="py-2.5 px-3 min-w-[130px] border-l border-slate-200">
                      {teamName}
                    </th>
                  ))}
                  <th className="py-2.5 px-2 w-12 text-center">Xóa</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredRowsWithIndices.length === 0 ? (
                  <tr>
                    <td colSpan={2 + teams.length} className="py-8 text-center text-slate-400 italic">
                      {searchFilter ? 'Không tìm thấy chức danh phù hợp.' : 'Chưa có danh sách nhân sự. Tải file Excel mẫu hoặc nhấn "Thêm Chức Danh / Vị Trí" để bắt đầu.'}
                    </td>
                  </tr>
                ) : (
                  filteredRowsWithIndices.map(({ row, originalIndex }) => (
                    <tr key={originalIndex} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-2 px-3 text-center text-slate-400 font-mono text-[11px]">
                        {originalIndex + 1}
                      </td>

                      {/* Title Cell (Index 0) */}
                      <td className="py-1.5 px-2">
                        <input
                          type="text"
                          value={row[0] || ''}
                          onChange={(e) => handleCellChange(originalIndex, 0, e.target.value)}
                          placeholder="Nhập tên chức danh (ví dụ: Trưởng ca)"
                          className="w-full px-2.5 py-1 text-xs font-semibold text-slate-800 bg-slate-50/50 border border-slate-200 focus:border-emerald-500 focus:bg-white rounded-md outline-none transition-all"
                        />
                      </td>

                      {/* Team Personnel Cells (Index 1..N) */}
                      {teams.map((_, teamIdx) => {
                        const cellColIndex = teamIdx + 1;
                        return (
                          <td key={teamIdx} className="py-1.5 px-2 border-l border-slate-100">
                            <input
                              type="text"
                              value={row[cellColIndex] || ''}
                              onChange={(e) => handleCellChange(originalIndex, cellColIndex, e.target.value)}
                              placeholder={`Họ tên ${teams[teamIdx]}`}
                              className="w-full px-2.5 py-1 text-xs text-slate-700 bg-white border border-slate-200 focus:border-emerald-500 rounded-md outline-none transition-all"
                            />
                          </td>
                        );
                      })}

                      {/* Delete Row Action */}
                      <td className="py-1.5 px-2 text-center">
                        <button
                          type="button"
                          onClick={() => handleRemoveRow(originalIndex)}
                          title="Xóa dòng này"
                          className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-all cursor-pointer"
                        >
                          <Trash2 size={14} />
                        </button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>

          <div className="flex items-center justify-between text-[11px] text-slate-500 italic px-1">
            <span>💡 Mẹo: Nhập chức danh ở cột 1 và họ tên nhân sự tương ứng từng kíp ở các cột tiếp theo.</span>
            <span>Tổng cộng: {rows.length} vị trí</span>
          </div>
        </div>
      )}

      {/* MODE 2: EXCEL FILE DOWNLOAD & UPLOAD TAB */}
      {activeMode === 'excel' && (
        <div className="space-y-4 py-1">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {/* Step 1: Download Template */}
            <div className="bg-white border border-slate-200 rounded-2xl p-4 flex flex-col justify-between shadow-2xs hover:border-emerald-300 transition-all">
              <div className="space-y-2">
                <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center font-bold">
                  <Download size={20} />
                </div>
                <div>
                  <h4 className="font-bold text-slate-900 text-sm">1. Tải File Excel Mẫu (.xlsx)</h4>
                  <p className="text-xs text-slate-500 mt-1 leading-relaxed">
                    Tải về file Excel chứa sẵn khung mẫu (Cột STT, Chức danh, Kíp 1, Kíp 2, Kíp 3...). Bạn chỉ cần mở file ra và điền họ tên nhân sự.
                  </p>
                </div>
              </div>

              <button
                type="button"
                onClick={handleDownloadTemplate}
                className="mt-4 w-full py-2.5 px-4 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 text-emerald-800 rounded-xl text-xs font-bold flex items-center justify-center gap-2 transition-all cursor-pointer shadow-xs"
              >
                <Download size={15} />
                <span>Tải File Excel Mẫu Trực Tiếp</span>
              </button>
            </div>

            {/* Step 2: Upload Excel File */}
            <div className="bg-white border border-slate-200 rounded-2xl p-4 flex flex-col justify-between shadow-2xs hover:border-emerald-300 transition-all">
              <div className="space-y-2">
                <div className="w-10 h-10 rounded-xl bg-teal-100 text-teal-700 flex items-center justify-center font-bold">
                  <Upload size={20} />
                </div>
                <div>
                  <h4 className="font-bold text-slate-900 text-sm">2. Tải Lên File Excel Đã Điền</h4>
                  <p className="text-xs text-slate-500 mt-1 leading-relaxed">
                    Chọn file Excel (.xlsx / .csv) chứa danh sách nhân sự từ máy tính của bạn để nhập tự động vào hệ thống.
                  </p>
                </div>

                {/* Mode Option: Replace vs Append */}
                <div className="pt-2 flex items-center gap-4 text-xs">
                  <label className="flex items-center gap-1.5 cursor-pointer font-medium text-slate-700">
                    <input 
                      type="radio" 
                      name="importOpt" 
                      value="replace" 
                      checked={importOption === 'replace'} 
                      onChange={() => setImportOption('replace')}
                      className="text-emerald-600 focus:ring-emerald-500"
                    />
                    <span>Thay thế danh sách cũ</span>
                  </label>

                  <label className="flex items-center gap-1.5 cursor-pointer font-medium text-slate-700">
                    <input 
                      type="radio" 
                      name="importOpt" 
                      value="append" 
                      checked={importOption === 'append'} 
                      onChange={() => setImportOption('append')}
                      className="text-emerald-600 focus:ring-emerald-500"
                    />
                    <span>Nối thêm vào danh sách</span>
                  </label>
                </div>
              </div>

              <button
                type="button"
                onClick={() => fileInputRef.current?.click()}
                className="mt-4 w-full py-2.5 px-4 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center justify-center gap-2 transition-all cursor-pointer shadow-sm"
              >
                <Upload size={15} />
                <span>Chọn File Excel & Upload</span>
              </button>
            </div>
          </div>
        </div>
      )}


    </div>
  );
}