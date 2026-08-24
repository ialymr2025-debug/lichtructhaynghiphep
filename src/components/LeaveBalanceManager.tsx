import { useCallback, useEffect, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

interface LeaveBalanceRow {
  name: string;
  year: string;
  entitled: number;
  seniority: number | null;
  source: 'seniority' | 'manual';
  travelDays: number;
  used: number;
  remaining: number;
}

interface Props {
  staffList: string[];
  onAlert: (msg: string) => void;
  workshopId: string;
}

const DEFAULT_BASE_DAYS = 16;

// Header labels in the workbook the company already keeps, matched without accents.
const stripAccents = (s: string) =>
  s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().trim();

export default function LeaveBalanceManager({ staffList, onAlert, workshopId }: Props) {
  const [rows, setRows] = useState<LeaveBalanceRow[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [status, setStatus] = useState<{ ok: boolean; text: string } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Report in place as well as through the shared banner, which sits far up the page.
  const report = useCallback((ok: boolean, text: string) => {
    setStatus({ ok, text });
    onAlert(text);
  }, [onAlert]);

  const thisYear = new Date().getFullYear();
  const [viewYear, setViewYear] = useState(String(thisYear));
  const [formName, setFormName] = useState('');
  const [formYear, setFormYear] = useState(String(thisYear));
  const [formEntitled, setFormEntitled] = useState('16');

  const loadRows = useCallback(async () => {
    setIsLoading(true);
    try {
      const res = await fetch(`/api/leave/balances?year=${viewYear}&workshopId=${encodeURIComponent(workshopId)}`);
      const data = await res.json();
      if (Array.isArray(data)) {
        setRows(data);
      } else {
        setRows([]);
        report(false, `❌ ${data?.error || 'Không tải được bảng phép năm.'}`);
      }
    } catch (e: any) {
      setRows([]);
      report(false, `❌ Không kết nối được máy chủ để tải bảng phép năm (${e.message}).`);
    } finally {
      setIsLoading(false);
    }
  }, [report, viewYear, workshopId]);

  useEffect(() => { loadRows(); }, [loadRows]);

  const handleImportFile = async (file: File) => {
    setIsImporting(true);
    setStatus({ ok: true, text: `⏳ Đang đọc file "${file.name}"...` });
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });

      // Prefer the sheet holding the entitlement list; fall back to the first sheet.
      const sheetName = wb.SheetNames.find(n => stripAccents(n).includes('so ngay phep')) || wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const table: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
      if (table.length < 2) throw new Error('Sheet không có dữ liệu.');

      const header = (table[0] || []).map((h: any) => stripAccents(String(h ?? '')));
      const findCol = (...keys: string[]) => header.findIndex(h => keys.some(k => h.includes(k)));
      const nameCol = findCol('ho va ten', 'ho ten');
      const hireCol = findCol('ngay vao lam', 'vao lam', 'nam vao');
      const entitledCol = findCol('so ngay phep', 'ngay phep duoc huong');
      const seniorityCol = findCol('tham nien');

      if (nameCol < 0 || hireCol < 0) {
        throw new Error('Không tìm thấy cột "Họ và tên" hoặc "Ngày vào làm việc".');
      }

      const parsed = [];
      for (let i = 1; i < table.length; i++) {
        const row = table[i] || [];
        const name = String(row[nameCol] ?? '').trim();
        if (!name) continue;

        const hireYear = parseInt(String(row[hireCol] ?? '').trim().slice(-4), 10);
        if (!hireYear) continue;

        // Recover each person's base from their sheet figures so exceptions survive the import.
        let baseDays = DEFAULT_BASE_DAYS;
        const entitled = parseFloat(String(row[entitledCol] ?? ''));
        if (entitledCol >= 0 && !isNaN(entitled)) {
          const seniority = seniorityCol >= 0 && !isNaN(parseFloat(String(row[seniorityCol] ?? '')))
            ? parseFloat(String(row[seniorityCol]))
            : thisYear - hireYear;
          baseDays = entitled - Math.floor(seniority / 5);
        }
        parsed.push({ name, hireYear, baseDays });
      }

      if (parsed.length === 0) throw new Error('Không đọc được dòng nhân viên nào.');

      const res = await fetch('/api/leave/employees/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rows: parsed, workshopId })
      });
      const data = await res.json();
      if (res.ok) {
        const skippedNote = data.skipped?.length ? ` Bỏ qua ${data.skipped.length} dòng: ${data.skipped.slice(0, 3).join('; ')}` : '';
        report(true, `✅ Đã nhập ${data.imported} nhân viên từ "${file.name}".${skippedNote}`);
        loadRows();
      } else {
        report(false, `❌ ${data.error || 'Nhập thất bại'}`);
      }
    } catch (e: any) {
      report(false, `❌ Không đọc được file Excel: ${e.message}`);
    } finally {
      setIsImporting(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  // Headers here must stay in step with the column detection in handleImportFile above
  // (stripAccents + substring match), otherwise the file this button hands out is one
  // the importer cannot read back. Sheet name contains "Số ngày phép" because the
  // importer prefers that sheet when a workbook has several.
  const handleDownloadTemplate = () => {
    const rows = [
      ['Họ và tên', 'Ngày vào làm việc', 'Thâm niên', 'Số ngày phép được hưởng'],
      ['Nguyễn Văn A', '01/08/2015', 11, 18],
      ['Trần Thị B', '2020', 6, 17],
      ['Lê Văn C', '15/03/1997', 29, 21],
    ];
    const sheet = XLSX.utils.aoa_to_sheet(rows);
    sheet['!cols'] = [{ wch: 28 }, { wch: 20 }, { wch: 12 }, { wch: 24 }];

    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, 'Số ngày phép năm');
    XLSX.writeFile(book, 'Mau_phep_nam_nhan_vien.xlsx');
    report(true, '✅ Đã tải file mẫu "Mau_phep_nam_nhan_vien.xlsx". Điền danh sách rồi bấm "Nhập từ Excel".');
  };

  const handleSave = async () => {
    const name = formName.trim();
    if (!name) {
      report(false, '⚠ Vui lòng chọn hoặc nhập họ tên nhân viên.');
      return;
    }
    if (!/^\d{4}$/.test(formYear)) {
      report(false, '⚠ Năm phải gồm 4 chữ số, ví dụ 2026.');
      return;
    }

    setIsSaving(true);
    try {
      const res = await fetch('/api/leave/balances', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, year: formYear, entitled: formEntitled, workshopId })
      });
      const data = await res.json();
      if (res.ok) {
        report(true, data.message || '✅ Đã lưu.');
        setFormName('');
        loadRows();
      } else {
        report(false, `❌ ${data.error || 'Lưu thất bại'}`);
      }
    } catch (e: any) {
      report(false, `❌ Lỗi kết nối: ${e.message}`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleDelete = async (row: LeaveBalanceRow) => {
    if (row.used > 0) {
      report(false, `⚠ Đồng chí ${row.name} đã nghỉ ${row.used} ngày phép năm ${row.year}, không nên xóa dòng này.`);
      return;
    }
    try {
      const res = await fetch('/api/leave/balances/delete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: row.name, year: row.year, workshopId })
      });
      const data = await res.json();
      if (res.ok) {
        report(true, `✅ Đã xóa phép năm ${row.year} của đồng chí ${row.name}.`);
        loadRows();
      } else {
        report(false, `❌ ${data.error || 'Xóa thất bại'}`);
      }
    } catch (e: any) {
      report(false, `❌ Lỗi kết nối: ${e.message}`);
    }
  };

  const startEdit = (row: LeaveBalanceRow) => {
    setFormName(row.name);
    setFormYear(row.year);
    setFormEntitled(String(row.entitled));
  };

  return (
    <div className="flex flex-col gap-4">
      <div className="p-3 bg-blue-50/50 border border-blue-100 rounded-xl text-[12px] text-slate-700">
        Số ngày phép <b>tự tính theo thâm niên</b>: mức cơ bản + 1 ngày cho mỗi 5 năm công tác.
        Sang năm mới thâm niên tự tăng, <b>không phải nhập lại</b>.
        Số ngày <b>đã nghỉ</b> và <b>còn lại</b> hệ thống tự tính từ các đơn đã lưu.
        Phép năm cũ vẫn dùng được đến hết <b>31/3</b> năm sau.
      </div>

      <div className="flex flex-wrap items-end gap-3">
        <div className="field">
          <label>Xem phép của năm</label>
          <select value={viewYear} onChange={e => setViewYear(e.target.value)}>
            {[thisYear - 2, thisYear - 1, thisYear, thisYear + 1].map(y => (
              <option key={y} value={String(y)}>{y}</option>
            ))}
          </select>
        </div>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          onChange={e => { const f = e.target.files?.[0]; if (f) handleImportFile(f); }}
        />
        <button className="btn btn-secondary" onClick={handleDownloadTemplate} disabled={isImporting}>
          ⬇️ Tải file mẫu
        </button>
        <button className="btn btn-secondary" onClick={() => fileInputRef.current?.click()} disabled={isImporting}>
          {isImporting ? <span className="spin mr-2"></span> : '📄'} Nhập từ Excel
        </button>
        <span className="text-[11px] text-slate-500 max-w-[340px]">
          Tải file mẫu, điền danh sách rồi bấm <b>Nhập từ Excel</b>. Nhập lại sẽ cập nhật người đã có, không xóa ai.
        </span>
      </div>

      {status && (
        <div className={`p-3 rounded-xl border text-[13px] font-medium ${
          status.ok ? 'bg-emerald-50 border-emerald-200 text-emerald-800' : 'bg-red-50 border-red-200 text-red-700'
        }`}>
          {status.text}
        </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-4 gap-3 items-end">
        <div className="field">
          <label>Họ và tên</label>
          <input
            type="text"
            list="balance-staff-list"
            placeholder="Chọn hoặc nhập tên"
            value={formName}
            onChange={e => setFormName(e.target.value)}
          />
          <datalist id="balance-staff-list">
            {staffList.map(n => <option key={n} value={n} />)}
          </datalist>
        </div>
        <div className="field">
          <label>Năm</label>
          <input type="number" value={formYear} onChange={e => setFormYear(e.target.value)} />
        </div>
        <div className="field">
          <label>Số ngày phép (ghi đè)</label>
          <input type="number" value={formEntitled} onChange={e => setFormEntitled(e.target.value)} />
        </div>
        <button className="btn btn-primary" onClick={handleSave} disabled={isSaving}>
          {isSaving ? <span className="spin mr-2"></span> : '💾'} Lưu
        </button>
      </div>
      <div className="text-[11px] text-slate-500 -mt-2">
        Ô trên chỉ dùng khi cần <b>ghi đè</b> số ngày phép của một người trong một năm cụ thể, khác với số tự tính theo thâm niên.
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-[13px] border-collapse">
          <thead>
            <tr className="bg-slate-50 text-slate-600 text-[11px] uppercase tracking-wider">
              <th className="text-left p-2 border-b border-slate-200">Họ và tên</th>
              <th className="p-2 border-b border-slate-200">Năm</th>
              <th className="p-2 border-b border-slate-200">Thâm niên</th>
              <th className="p-2 border-b border-slate-200">Được hưởng</th>
              <th className="p-2 border-b border-slate-200">Đi đường</th>
              <th className="p-2 border-b border-slate-200">Đã nghỉ</th>
              <th className="p-2 border-b border-slate-200">Còn lại</th>
              <th className="p-2 border-b border-slate-200"></th>
            </tr>
          </thead>
          <tbody>
            {isLoading ? (
              <tr><td colSpan={8} className="text-center p-4 text-slate-500"><span className="spin mr-2"></span>Đang tải...</td></tr>
            ) : rows.length === 0 ? (
              <tr><td colSpan={8} className="text-center p-4 text-slate-500">Chưa có nhân viên nào. Bấm "Nhập từ Excel" để bắt đầu.</td></tr>
            ) : rows.map(row => (
              <tr key={`${row.name}-${row.year}`} className="hover:bg-slate-50/60">
                <td className="p-2 border-b border-slate-100 font-medium">{row.name}</td>
                <td className="p-2 border-b border-slate-100 text-center">{row.year}</td>
                <td className="p-2 border-b border-slate-100 text-center text-slate-500">
                  {row.seniority !== null ? `${row.seniority} năm` : '—'}
                </td>
                <td className="p-2 border-b border-slate-100 text-center">
                  {row.entitled}
                  {row.source === 'manual' && <span className="ml-1 text-[11px] text-amber-600" title="Số ngày ghi đè thủ công">✎</span>}
                </td>
                <td className="p-2 border-b border-slate-100 text-center text-slate-500">{row.travelDays > 0 ? `+${row.travelDays}` : '—'}</td>
                <td className="p-2 border-b border-slate-100 text-center text-amber-700">{row.used}</td>
                <td className={`p-2 border-b border-slate-100 text-center font-bold ${row.remaining < 0 ? 'text-red-600' : 'text-emerald-700'}`}>
                  {row.remaining}
                </td>
                <td className="p-2 border-b border-slate-100 text-center whitespace-nowrap">
                  <button className="text-[#00529c] hover:underline mr-3 cursor-pointer" onClick={() => startEdit(row)}>Sửa</button>
                  <button className="text-red-600 hover:underline cursor-pointer" onClick={() => handleDelete(row)}>Xóa</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
