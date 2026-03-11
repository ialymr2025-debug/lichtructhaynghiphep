import { useState, useEffect, useCallback } from 'react';
import { DEFAULT_STAFF, SHIFTS } from './constants';
import { fmtVN, fmtIn, dayN, timNghi, timThay, abbrev } from './utils/shiftHelpers';
import { buildMultiLeaveResults, Leave, ResultItem } from './utils/multiLeaveAlgorithm';
import { exportWord } from './utils/wordExport';

export default function App() {
  const [staffData, setStaffData] = useState<string[][]>(() => {
    const saved = localStorage.getItem('sd');
    const ver = localStorage.getItem('sv');
    if (ver !== 'v4') {
      localStorage.removeItem('sd');
      localStorage.setItem('sv', 'v4');
      return DEFAULT_STAFF;
    }
    return saved ? JSON.parse(saved) : DEFAULT_STAFF;
  });

  const [ngayBatDau, setNgayBatDau] = useState(() => fmtIn(new Date()));
  const [ngayKetThuc, setNgayKetThuc] = useState(() => {
    const next = new Date();
    next.setDate(next.getDate() + 7);
    return fmtIn(next);
  });
  const [chucDanh, setChucDanh] = useState('');
  const [kipNghi, setKipNghi] = useState('');
  const [additionalLeaves, setAdditionalLeaves] = useState<{ kip: string, start: string, end: string }[]>([]);
  const [showStaff, setShowStaff] = useState(false);
  const [alert, setAlert] = useState<string | null>(null);
  const [currentResult, setCurrentResult] = useState<any>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [config, setConfig] = useState({
    soVanBan: '',
    ngayKy: '',
    nguoiKy: 'Nguyễn Văn Nghị'
  });

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
    setAdditionalLeaves([...additionalLeaves, { kip: '', start: '', end: '' }]);
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
      if (!al.kip || !al.start || !al.end) {
        addErr.push(`Người nghỉ #${idx + 2} thiếu thông tin`);
        return;
      }
      const alKip = +al.kip;
      if (alKip === kip) {
        addErr.push(`Người nghỉ #${idx + 2} trùng kíp với người nghỉ chính`);
        return;
      }
      if (allLeaves.some(l => l.kip === alKip)) {
        addErr.push(`Kíp ${alKip} đã được thêm`);
        return;
      }
      const alStart = new Date(al.start + 'T00:00:00');
      const alEnd = new Date(al.end + 'T00:00:00');
      if (alStart > alEnd) {
        addErr.push(`Người nghỉ #${idx + 2} ngày bắt đầu sau ngày kết thúc`);
        return;
      }
      const alTen = timNghi(chucDanh, alKip, staffData);
      if (!alTen) {
        addErr.push(`Không tìm thấy "${chucDanh}" trong Kíp ${alKip}`);
        return;
      }
      allLeaves.push({ kip: alKip, start: alStart, end: alEnd, ten: alTen, chucDanh });
    });

    if (addErr.length) {
      setAlert('⚠ ' + addErr.join(' | '));
      return;
    }

    setIsProcessing(true);
    setTimeout(() => {
      const buildResult = buildMultiLeaveResults(allLeaves, chucDanh, staffData);
      setCurrentResult({
        ten, chucDanh, kip, start, end,
        ketQua: buildResult.results[0].ketQua,
        allResults: buildResult.results,
        extraRows: buildResult.extraRows,
        hasConflict: buildResult.hasConflict,
        coverCount: buildResult.coverCount,
        isMulti: allLeaves.length > 1
      });
      setIsProcessing(false);
    }, 250);
  };

  const handleExportWord = async () => {
    setIsProcessing(true);
    await exportWord(currentResult, config);
    setIsProcessing(false);
  };

  const handleExportCSV = () => {
    if (!currentResult) return;
    const d = currentResult;
    let csv = '\uFEFF';
    csv += 'CÔNG TY THỦY ĐIỆN IALY - BẢNG PHÂN CÔNG TRỰC THAY NGHỈ PHÉP\n\n';
    csv += `Người nghỉ:,${d.ten}\nChức danh:,${d.chucDanh}\nKíp:,Kíp ${d.kip}\n`;
    csv += `Thời gian:,${fmtVN(d.start)} đến ${fmtVN(d.end)}\n\n`;
    csv += 'STT,Ngày,Ca nghỉ,Kíp thay,Người đi thay\n';
    d.ketQua.forEach((it: any, i: number) => {
      csv += `${i + 1},${fmtVN(it.ngay)},${it.ca},Kíp ${it.kipThay},${it.nguoiThay}\n`;
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
          kipThay: it.kipThay, nguoiThay: it.nguoiThay,
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
        kipThay: it.kipThay, nguoiThay: it.nguoiThay,
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
    if (!coverStats[k]) coverStats[k] = { total: 0, N: 0, C: 0, K: 0 };
    coverStats[k].total++;
    coverStats[k][row.ca]++;
  });

  const isEvenDistribution = () => {
    const counts = Object.values(coverStats).map(s => s.total);
    if (counts.length === 0) return true;
    return (Math.max(...counts) - Math.min(...counts)) <= 1;
  };

  return (
    <div className="wrap">
      <header>
        <div className="tag">Thủy Điện Ialy &middot; Phân Xưởng Vận Hành</div>
        <h1>Lịch Trực Thay Ca Nghỉ Phép</h1>
        <p className="sub">Hệ thống phân công tự động theo quy luật kíp vận hành</p>
      </header>

      <div className="card">
        <div className="ctitle">Thông tin nghỉ phép</div>
        {alert && <div className="alert aerr">{alert}</div>}
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
            Người nghỉ đồng thời <span className="font-normal">(cùng chức danh)</span>
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

      <div className="card">
        <div className="ctitle">
          Danh sách nhân sự 
          <button className="staff-toggle" onClick={() => setShowStaff(!showStaff)}>
            {showStaff ? 'Thu gọn ▲' : 'Chỉnh sửa ▼'}
          </button>
        </div>
        <p className="text-[13px] text-var(--txt2)">Nhấn "Chỉnh sửa" để cập nhật tên nhân viên trong từng kíp.</p>
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

      {currentResult && (
        <div id="results">
          <div className="card">
            <div className="res-hdr">
              <div className="ctitle mb-0">Kết quả phân công</div>
              <div className="ex-row">
                <button className="btn-ex btn-word" onClick={handleExportWord} disabled={isProcessing}>
                  {isProcessing ? <span className="spin spinw mr-2"></span> : '📝'} Xuất Word
                </button>
                <button className="btn-ex btn-csv" onClick={handleExportCSV}>📄 Xuất CSV</button>
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
                      <div className="text-[13px] font-bold text-var(--txt)">Kíp {kip} – {timThay(+kip, currentResult.chucDanh, staffData)}</div>
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
                            {row.isSwap ? (
                              <>
                                <span className="text-[#22c55e] text-[11px]">đổi ca</span>
                                <br />
                                {row.absentTen}
                                <br />
                                <span className="text-[10px]">Kíp {row.absentKip}</span>
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
              <div className="legend-item"><span className="badge bN">N</span>Ca Ngày (6:00–14:00)</div>
              <div className="legend-item"><span className="badge bC">C</span>Ca Chiều (14:00–22:00)</div>
              <div className="legend-item"><span className="badge bK">K</span>Ca Đêm (22:00–6:00)</div>
              <div className="legend-item"><span className="conflict-badge bg-[#22c55e1a] text-[#22c55e] border-[#22c55e4d]">⇄ O tròn</span>Ca nghỉ (O) ngay sau ca đêm (K) - Ưu tiên hỗ trợ ca N/C</div>
              <div className="legend-item"><span className="conflict-badge bg-[#eab3081a] text-[#eab308] border-[#eab3084d]">⚡ Pre-Relief</span>Hỗ trợ kíp thay chính nghỉ ngơi trước ca trực quan trọng</div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

