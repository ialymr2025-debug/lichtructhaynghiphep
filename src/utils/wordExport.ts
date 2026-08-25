import JSZip from 'jszip';
import { fmtVN, abbrev } from './shiftHelpers';
import { SIGNATURES } from '../constants/signatures';


function xe(s: string) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function rpr(o: any = {}) {
  const bold = o.bold || false, italic = o.italic || false, size = o.size || 24, underline = o.underline || false;
  return '<w:rPr>' + (bold ? '<w:b/>' : '')
    + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    + '<w:sz w:val="' + size + '"/><w:szCs w:val="' + size + '"/>'
    + (italic ? '<w:i/>' : '') + (underline ? '<w:u w:val="single"/>' : '')
    + '</w:rPr>';
}

function wrun(text: string, opts: any = {}) {
  return '<w:r>' + rpr(opts) + '<w:t xml:space="preserve">' + xe(text) + '</w:t></w:r>';
} 
function wpara(content: string, o: any = {}) {
  const align = o.align || 'left';
  const spB = o.spBefore || 0, spA = o.spAfter || 0;
  let indent = '';
  if (o.indent) {
    const left = o.indent.left !== undefined ? ' w:left="' + o.indent.left + '"' : '';
    const firstLine = o.indent.firstLine !== undefined ? ' w:firstLine="' + o.indent.firstLine + '"' : '';
    indent = '<w:ind' + left + firstLine + '/>';
  }
  return '<w:p><w:pPr><w:jc w:val="' + align + '"/><w:spacing w:before="' + spB + '" w:after="' + spA + '"/>' + indent + '</w:pPr>' + content + '</w:p>';
}

function wimage(rId: string, width: number = 1500000, height: number = 800000) {
  return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
    + '<wp:extent cx="' + width + '" cy="' + height + '"/>'
    + '<wp:docPr id="1" name="Picture 1"/>'
    + '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
    + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    + '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + '<pic:nvPicPr><pic:cNvPr id="0" name="Signature.png"/><pic:cNvPicPr/></pic:nvPicPr>'
    + '<pic:blipFill><a:blip r:embed="' + rId + '" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
    + '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + width + '" cy="' + height + '"/></a:xfrm>'
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}

function emptyP(spB: number = 0, spA: number = 0) {
  return '<w:p><w:pPr><w:spacing w:before="' + spB + '" w:after="' + spA + '"/></w:pPr></w:p>';
}

// A signature image is 800000 EMU tall (== 1260 twips, since 1 twip = 635 EMU). When someone
// has no signature on file yet, this stands in its place so their printed name lines up
// exactly with everyone else's — an approximate spacing split (e.g. emptyP(150,900)) drifts
// a bit short in practice and leaves names visibly uneven.
function sigPlaceholder() {
  return emptyP(0, 1260);
}

function makeBorders(solid: boolean) {
  if (solid) {
    return '<w:tcBorders>'
      + '<w:top w:val="single" w:sz="6" w:color="000000"/>'
      + '<w:left w:val="single" w:sz="6" w:color="000000"/>'
      + '<w:bottom w:val="single" w:sz="6" w:color="000000"/>'
      + '<w:right w:val="single" w:sz="6" w:color="000000"/>'
      + '</w:tcBorders>';
  }
  return '<w:tcBorders>'
    + '<w:top w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '<w:left w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '<w:bottom w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '<w:right w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '</w:tcBorders>';
}

function wtc(o: any) {
  const w = o.w, content = o.content, gridSpan = o.gridSpan || 1;
  const vMerge = o.vMerge || '', borders = o.borders !== false, shading = o.shading || null;
  const gs = gridSpan > 1 ? '<w:gridSpan w:val="' + gridSpan + '"/>' : '';
  const vm = vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : vMerge === 'cont' ? '<w:vMerge/>' : '';
  const sh = shading ? '<w:shd w:val="clear" w:color="auto" w:fill="' + shading + '"/>' : '';
  const noMargin = o.noMargin || false;
  return '<w:tc><w:tcPr>'
    + '<w:tcW w:w="' + w + '" w:type="dxa"/>' + gs + vm
    + makeBorders(borders) + sh
    + '<w:vAlign w:val="top"/>'
    + (noMargin ? '<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>' :
      '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
    + '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>')
    + '</w:tcPr>' + content + '</w:tc>';
}

function wtr(cells: string[]) {
  return '<w:tr><w:trPr/>' + cells.join('') + '</w:tr>';
}

function wline(width: number, cellWidth: number) {
  const indent = Math.max(0, Math.floor((cellWidth - width) / 2));
  return '<w:p><w:pPr>'
    + '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>'
    + '<w:ind w:left="' + indent + '" w:right="' + indent + '"/>'
    + '<w:jc w:val="center"/>'
    + '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
    + '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'
    + '</w:pPr></w:p>';
}

function wtable(rows: string[], colWidths: number[]) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  const grid = colWidths.map(w => '<w:gridCol w:w="' + w + '"/>').join('');
  const tblBrd = '<w:tblBorders>'
    + '<w:top w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '<w:left w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '<w:bottom w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '<w:right w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '<w:insideH w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '<w:insideV w:val="none" w:sz="0" w:color="FFFFFF"/>'
    + '</w:tblBorders>';
  return '<w:tbl><w:tblPr>'
    + '<w:tblW w:w="' + total + '" w:type="dxa"/>'
    + '<w:tblLayout w:type="fixed"/>'
    + tblBrd
    + '<w:tblCellMar>'
    + '<w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
    + '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
    + '</w:tblCellMar>'
    + '</w:tblPr><w:tblGrid>' + grid + '</w:tblGrid>' + rows.join('') + '</w:tbl>';
}

function twoCol(lc: string, rc: string, CW: number) {
  const HW = Math.floor(CW / 2);
  return wtable([wtr([
    wtc({ w: HW, borders: false, content: lc }),
    wtc({ w: CW - HW, borders: false, content: rc })
  ])], [HW, CW - HW]);
}

function buildDateStr(val: any, locationName: string = 'Gia Lai') {
  const loc = locationName || 'Gia Lai';
  if (val) {
    const nd = (typeof val === 'string') ? new Date(val + (val.includes('T') ? '' : 'T00:00:00')) : val;
    if (isNaN(nd.getTime())) return `${loc}, ngày      tháng      năm     `;
    return `${loc}, ngày ` + nd.getDate() + ' tháng ' + (nd.getMonth() + 1) + ' năm ' + nd.getFullYear();
  }
  return `${loc}, ngày      tháng      năm     `;
}
function isOverlap(aStart: string, aEnd: string, bStart: string, bEnd: string) {
  return !(new Date(aEnd) < new Date(bStart) || new Date(bEnd) < new Date(aStart));
}
export function buildDocXml(currentResult: any, config: any) {
  const d = currentResult;
  const soVanBan = config.soVanBan || '';
  const nguoiKy = config.nguoiKy || 'Quản Đốc';
  const ngayKyVal = config.ngayKy || '';
  const locationName = config?.locationName || config?.location || 'Gia Lai';

  const CW = 9184, c0 = 2200, c1 = 1500, HW = Math.floor(CW / 2);
  const GRAY = 'D9D9D9';
  const shiftOrd: Record<string, number> = { N: 0, C: 1, K: 2 };
  const MAX_PER_ROW = 4;

  const allResults = d.allResults;
  const extraRows = d.extraRows || [];

  const roleGroups: Record<string, any[]> = {};

  allResults.forEach((res: any) => {
    const role = res.chucDanh || d.chucDanh || '';
    if (!roleGroups[role]) roleGroups[role] = [];
    roleGroups[role].push(res);
  });

  const validRoles: string[] = [];

  Object.keys(roleGroups).forEach(role => {
    const list = roleGroups[role];
    let hasOverlap = false;
    for (let i = 0; i < list.length; i++) {
      for (let j = i + 1; j < list.length; j++) {
        if (isOverlap(list[i].start, list[i].end, list[j].start, list[j].end)) {
          hasOverlap = true;
          break;
        }
      }
      if (hasOverlap) break;
    }
    if (hasOverlap) validRoles.push(role);
  });

  const extraNotesByRole: Record<string, Set<string>> = {};
  validRoles.forEach(role => {
    extraNotesByRole[role] = new Set();
  });

  allResults.forEach((res: any) => {
    const role = res.chucDanh || d.chucDanh || '';
    if (!validRoles.includes(role)) return;
    res.ketQua.forEach((it: any) => {
      if (it.isSwap || it.relievedTen) {
        const dateStr = it.ca + '-' + fmtVN(it.ngay).slice(0, 5);
        const relieved = it.relievedTen || it.swapAbsentTen || res.ten;
        extraNotesByRole[role].add(`${abbrev(it.nguoitructhay)} trực thay ${abbrev(relieved)} ca ${dateStr}`);
      }
    });
  });

  extraRows.forEach((ex: any) => {
    if (!ex.isSwap) return;
    const role = ex.chucDanh;
    if (!validRoles.includes(role)) return;
    const dateStr = ex.ca + '-' + fmtVN(ex.ngay).slice(0, 5);
    const relieved = ex.relievedTen || ex.absentTen;
    extraNotesByRole[role].add(`${abbrev(ex.nguoitructhay)} trực thay ${abbrev(relieved)} ca ${dateStr}`);
  });

  const swapRows = extraRows.filter((r: any) => r.isSwap);
  const chainRows = extraRows.filter((r: any) => r.isCKChain && !r.isSwap);

  const companyName = (config?.companyName || 'CÔNG TY THỦY ĐIỆN IALY').toUpperCase();
  const headerWsName = (config?.headerWorkshopName || 'PHÂN XƯỞNG VẬN HÀNH IALY').toUpperCase();
  const docCodeSuffix = config?.documentCodeSuffix || '/VHIALY';
  const shortWsName = config?.shortWorkshopName || 'PXVH Ialy';
  const archiveCode = docCodeSuffix.replace('/', '').trim() || 'VHIALY';

  const hdrTbl = wtable([
    wtr([
      wtc({
        w: 4100, borders: false, noMargin: true,
        content: wpara(wrun(companyName, { size: 24 }), { align: 'center' })
      }),
      wtc({
        w: CW - 4100, borders: false, noMargin: true,
        content: wpara(wrun('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { bold: true, size: 24 }), { align: 'center' })
      })
    ]),
    wtr([
      wtc({
        w: 4100, borders: false, noMargin: true,
        content: wpara(wrun(headerWsName, { bold: true, size: 24 }), { align: 'center', spAfter: 0 })
          + wline(1400, 4200)
      }),
      wtc({
        w: CW - 4100, borders: false, noMargin: true,
        content: wpara(wrun('Độc lập - Tự do - Hạnh phúc', { bold: true, size: 26 }), { align: 'center', spAfter: 0 })
          + wline(3200, CW - 4200)
      })
    ]),
    wtr([
      wtc({
        w: 4100, borders: false, noMargin: true,
        content: wpara(wrun('Số: ' + (soVanBan || '      ') + docCodeSuffix, { size: 26 }), { align: 'center' })
      }),
      wtc({
        w: CW - 4100, borders: false, noMargin: true,
        content: wpara(wrun(buildDateStr(ngayKyVal, locationName), { italic: true, size: 26 }), { align: 'center' })
      })
    ])
  ], [4100, CW - 4100]);

  const title = wpara(wrun('LỊCH TRỰC THAY CA VẬN HÀNH', { bold: true, size: 28 }),
    { align: 'center', spBefore: 120, spAfter: 120 });

  function buildPersonRows(res: any) {
    const rows = res.ketQua.slice();
    const resCD = res.chucDanh || d.chucDanh || '';
    chainRows.forEach((ex: any) => {
      if (ex.absentKip === res.kip && ex.chucDanh === resCD) rows.push(ex);
    });
    rows.sort((a: any, b: any) => {
      if (a.ngay < b.ngay) return -1; if (a.ngay > b.ngay) return 1;
      return (shiftOrd[a.ca] || 0) - (shiftOrd[b.ca] || 0);
    });
    return rows;
  }

  let maxCa = 0;
  allResults.forEach((res: any) => {
    const rows = buildPersonRows(res);
    if (rows.length > maxCa) maxCa = rows.length;
  });
  const nCols = Math.min(MAX_PER_ROW, maxCa || 1);
  const caW = Math.floor((CW - c0 - c1) / nCols);
  const colW = [c0, c1]; for (let k = 0; k < nCols; k++) colW.push(caW);
  const tot = colW.reduce((a, b) => a + b, 0); colW[colW.length - 1] += (CW - tot);

  const tableRows = [];
  tableRows.push(wtr([
    wtc({ w: c0, shading: GRAY, content: wpara(wrun('Tên người cần được\ntrực thay', { bold: true, size: 24 }), { align: 'center' }) }),
    wtc({ w: c1, shading: GRAY, content: wpara(wrun('Chức danh\nhiện tại', { bold: true, size: 24 }), { align: 'center' }) }),
    wtc({ w: CW - c0 - c1, gridSpan: nCols, shading: GRAY, content: wpara(wrun('Tên người trực thay', { bold: true, size: 24 }), { align: 'center' }) })
  ]));

  allResults.forEach((res: any) => {
    // Collect notes for this person (Swaps/Conflicts) FIRST
    const personNotes: string[] = [];
    const resCD = res.chucDanh || d.chucDanh || '';

    res.ketQua.forEach((it: any) => {
      if (it.conflictNote && (it.conflictNote.includes('Đổi ca') || it.conflictNote.includes('Hoán đổi') || it.relievedTen)) {
        const dateStr = it.ca + '-' + fmtVN(it.ngay).slice(0, 5);
        const relieved = it.relievedTen || it.swapAbsentTen || res.ten;
        personNotes.push(`${abbrev(it.nguoitructhay)} trực thay ${abbrev(relieved)} ca ${dateStr}.`);
      }
    });
    extraRows.forEach((ex: any) => {
      if (ex.isSwap && ex.absentKip === res.kip && ex.chucDanh === resCD) {
        const dateStr = ex.ca + '-' + fmtVN(ex.ngay).slice(0, 5);
        const relieved = ex.relievedTen || ex.absentTen;
        personNotes.push(`${abbrev(ex.nguoitructhay)} trực thay ${abbrev(relieved)} ca ${dateStr}.`);
      }
    });

    const rows = buildPersonRows(res);
    let groups = [];
    for (let i = 0; i < rows.length; i += nCols) groups.push(rows.slice(i, i + nCols));
    if (!groups.length) groups = [[]];

    const nghiStr = '(nghỉ phép từ ' + fmtVN(res.start) + ' đến ' + fmtVN(res.end) + ')';

    groups.forEach((grp, gi) => {
      const isFirst = (gi === 0);
      const isLast = (gi === groups.length - 1);
      const caRow = [];
      if (isFirst) {
        caRow.push(wtc({
          w: c0, vMerge: 'restart',
          content: wpara(wrun(res.ten, { size: 24 }), { align: 'center', spAfter: 20 })
            + wpara(wrun(nghiStr, { italic: true, size: 20 }), { align: 'center' })
        }));
        caRow.push(wtc({
          w: c1, vMerge: 'restart',
          content: wpara(wrun(resCD, { size: 24 }), { align: 'center' })
        }));
      } else {
        caRow.push(wtc({ w: c0, vMerge: 'cont', content: emptyP() }));
        caRow.push(wtc({ w: c1, vMerge: 'cont', content: emptyP() }));
      }

      // Add shifts in this group
      for (let i = 0; i < grp.length; i++) {
        const it = grp[i];
        caRow.push(wtc({
          w: colW[i + 2], shading: it ? GRAY : '',
          content: wpara(wrun(it ? it.ca + '-' + fmtVN(it.ngay).slice(0, 5) : '', { bold: true, size: 24 }), { align: 'center' })
        }));
      }

      let noteIncludedInThisRow = false;
      const uniqueNotes = Array.from(new Set(personNotes)) as string[];

      if (isLast && grp.length < nCols) {
        const remainingCols = nCols - grp.length;
        const remainingW = colW.slice(grp.length + 2).reduce((a, b) => a + b, 0);
        
        // Gộp các ô trống: Nếu còn trống > 1 ô và có ghi chú, chèn ghi chú vào đó.
        // Nếu chỉ còn trống 1 ô, ta sẽ để trống ô đó và đẩy ghi chú xuống hàng riêng biệt bên dưới.
        if (uniqueNotes.length > 0 && remainingCols > 1) {
          let noteContent = '';
          uniqueNotes.forEach(line => {
            noteContent += wpara(wrun(line, { size: 24 }), { spBefore: 40, spAfter: 40 });
          });
          caRow.push(wtc({
            w: remainingW, gridSpan: remainingCols, vMerge: 'restart',
            content: noteContent
          }));
          noteIncludedInThisRow = true;
        } else {
          // Gộp các ô trống còn lại làm một (merge) để bảng sạch sẽ hơn
          caRow.push(wtc({ w: remainingW, gridSpan: remainingCols, content: emptyP() }));
        }
      }
      tableRows.push(wtr(caRow));

      const tenRow = [
        wtc({ w: c0, vMerge: 'cont', content: emptyP() }),
        wtc({ w: c1, vMerge: 'cont', content: emptyP() })
      ];
      for (let i = 0; i < grp.length; i++) {
        const it = grp[i];
        tenRow.push(wtc({
          w: colW[i + 2],
          content: it ? wpara(wrun(abbrev(it.nguoitructhay), { size: 22 }), { align: 'center' }) : emptyP()
        }));
      }

      if (noteIncludedInThisRow) {
        const remainingCols = nCols - grp.length;
        const remainingW = colW.slice(grp.length + 2).reduce((a, b) => a + b, 0);
        tenRow.push(wtc({ w: remainingW, gridSpan: remainingCols, vMerge: 'cont', content: emptyP() }));
      } else if (isLast && grp.length < nCols) {
        const remainingCols = nCols - grp.length;
        const remainingW = colW.slice(grp.length + 2).reduce((a, b) => a + b, 0);
        tenRow.push(wtc({ w: remainingW, gridSpan: remainingCols, content: emptyP() }));
      }
      tableRows.push(wtr(tenRow));

      // Nếu có ghi chú nhưng chưa được chèn (do hàng đầy hoặc chỉ còn 1 ô trống),
      // thì thêm một hàng riêng biệt bên dưới để hiển thị ghi chú.
      const shouldAddSeparateRow = (isLast && !noteIncludedInThisRow && uniqueNotes.length > 0);

      if (shouldAddSeparateRow) {
        let noteContent = '';
        uniqueNotes.forEach(line => {
          noteContent += wpara(wrun(line, { size: 24 }), { spBefore: 40, spAfter: 40 });
        });
        tableRows.push(wtr([
          wtc({ w: c0, vMerge: 'cont', content: emptyP() }),
          wtc({ w: c1, vMerge: 'cont', content: emptyP() }),
          wtc({ w: CW - c0 - c1, gridSpan: nCols, content: noteContent })
        ]));
      }
    });
  });

  const mainTbl = wtable(tableRows, colW);
let extraNoteBlock = '';

validRoles.forEach(role => {
  const notes = Array.from(extraNotesByRole[role] || []);

  extraNoteBlock += wpara(
    wrun(`Tại ${role} :`, { bold: true, size: 24 }),
    { spBefore: 100 }
  );

  if (notes.length > 0) {
    notes.forEach(n => {
      extraNoteBlock += wpara(
        wrun(`- ${n}`, { size: 24 }),
        { indent: { left: 400 } }
      );
    });
  } else {
    extraNoteBlock += wpara(
      wrun(`-`, { size: 24 }),
      { indent: { left: 400 } }
    );
  }
});
  const note = wpara(
    wrun('Ghi chú: Các chức danh kiểm tra lại lịch trực của mình, nếu có gì vướng mắc phải báo lại PX để kiểm tra và điều chỉnh kịp thời./.',
      { size: 24 }),
    { spBefore: 120, spAfter: 60, indent: { firstLine: 400 } }
  );

  const footTbl = wtable([
    wtr([
      wtc({
        w: HW, borders: false, content:
          wpara(wrun('Nơi nhận:', { bold: true, italic: true, size: 24 }), { spBefore: 80 })
          + wpara(wrun('- Các kíp (để t/hiện);', { italic: true, size: 22 }))
          + wpara(wrun(`- Lưu: ${archiveCode}.`, { italic: true, size: 22 }))
      }),
      wtc({
        w: CW - HW, borders: false, content:
          wpara(wrun('QUẢN ĐỐC', { bold: true, size: 24 }), { align: 'center', spBefore: 80 })
          + emptyP(500, 0) + emptyP(500, 0) 
          + wpara(wrun(nguoiKy, { bold: true, size: 24 }), { align: 'center' })
      })
    ])
  ], [HW, CW - HW]);

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:body>'
    + hdrTbl + title + mainTbl + extraNoteBlock + note + footTbl
    + '<w:sectPr>'
    + '<w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1134" w:right="1020" w:bottom="1134" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr>'
    + '</w:body></w:document>';
}

export async function generateWordBlob(currentResult: any, config: any) {
  if (!currentResult) return null;
  
  const docXml = buildDocXml(currentResult, config);

  const CT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '</Types>';

  const RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '</Relationships>';

  const WRELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    + '</Relationships>';

  const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:docDefaults><w:rPrDefault><w:rPr>'
    + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    + '<w:sz w:val="24"/><w:szCs w:val="24"/>'
    + '</w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    + '<w:name w:val="Normal"/>'
    + '</w:style>'
    + '</w:styles>';

  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT);
  zip.file('_rels/.rels', RELS);
  zip.file('word/_rels/document.xml.rels', WRELS);
  zip.file('word/styles.xml', STYLES);
  zip.file('word/document.xml', docXml);

  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE'
  });
}

export function buildLeaveRequestDocXml(data: any, config: any, rIds: any = {}) {
  const { name, birthYear, chucDanh, kip, startDate, endDate, reason, phone, leaveYear, location, hasLeavePermit } = data;
  const nguoiKy = config.nguoiKy || 'Nguyễn Văn Nghị';
  const chucVuNguoiKy = config.chucVuNguoiKy || 'QUẢN ĐỐC';
  const ngayKyVal = config.ngayKy || '';
  const locationName = config?.locationName || config?.location || 'Gia Lai';
  const CW = 9184, HW = Math.floor(CW / 2);

  const dStart = new Date(startDate + 'T00:00:00');
  const dEnd = new Date(endDate + 'T00:00:00');
  const dfStart = fmtVN(dStart);
  const dfEnd = fmtVN(dEnd);
  const currentYear = leaveYear || new Date().getFullYear();

  const companyName = (config?.companyName || 'CÔNG TY THỦY ĐIỆN IALY').toUpperCase();
  const headerWsName = (config?.headerWorkshopName || 'PHÂN XƯỞNG VẬN HÀNH IALY').toUpperCase();
  const docCodeSuffix = config?.documentCodeSuffix || '/VHIALY';
  const recipientWsName = config?.recipientWorkshopName || 'Phân xưởng vận hành Ialy';
  const shortWsName = config?.shortWorkshopName || 'PXVH Ialy';
  const archiveCode = docCodeSuffix.replace('/', '') || 'VHIALY';

  const hdrTbl = wtable([
    wtr([
      wtc({
        w: 4100, borders: false, noMargin: true,
        content: wpara(wrun(companyName, { size: 24 }), { align: 'center' })
      }),
      wtc({
        w: CW - 4100, borders: false, noMargin: true,
        content: wpara(wrun('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { bold: true, size: 24 }), { align: 'center' })
      })
    ]),
    wtr([
      wtc({
        w: 4100, borders: false, noMargin: true,
        content: wpara(wrun(headerWsName, { bold: true, size: 24 }), { align: 'center', spAfter: 0 })
          + wline(1400, 4100)
      }),
      wtc({
        w: CW - 4100, borders: false, noMargin: true,
        content: wpara(wrun('Độc lập - Tự do - Hạnh phúc', { bold: true, size: 26 }), { align: 'center', spAfter: 0 })
          + wline(3200, CW - 4100)
      })
    ]),
    wtr([
      wtc({
        w: 4100, borders: false, noMargin: true,
        content: wpara(wrun(`Số:        ${docCodeSuffix}`, { size: 26 }), { align: 'center' })
      }),
      wtc({
        w: CW - 4100, borders: false, noMargin: true,
        content: wpara(wrun(buildDateStr(ngayKyVal, locationName), { italic: true, size: 26 }), { align: 'center' })
      })
    ])
  ], [4100, CW - 4100]);

  const title = wpara(wrun('ĐƠN XIN NGHỈ PHÉP', { bold: true, size: 28 }),
    { align: 'center', spBefore: 240, spAfter: 240 });

  const kinhGui = wtable([
    wtr([
      wtc({ w: 3200, borders: false, content: wpara(wrun('Kính gửi:', {  size: 28 }), { align: 'right' }) }),
      wtc({ 
        w: CW - 3200, borders: false, 
        content: wpara(wrun(`- ${recipientWsName};`, { size: 28 }))
               + wpara(wrun('- Phòng Hành chính và Lao động.', { size: 28 }))
      })
    ])
  ], [3200, CW - 3200]);

  const info1 = wtable([
    wtr([
      wtc({
        w: 5000,
        borders: false,
        noMargin: true,
        content: wpara(wrun(`Tên tôi là: ${name}`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } })
      }),
      wtc({
        w: CW - 5000,
        borders: false,
        noMargin: true,
        content: wpara(wrun(`Sinh năm: ${birthYear}`, { size: 28 }), { spBefore: 120 })
      })
    ]),
    wtr([
      wtc({
        w: 5000,
        borders: false,
        noMargin: true,
        content: wpara(wrun(`Hiện đang công tác: ${shortWsName}`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } })
      }),
      wtc({
        w: CW - 5000,
        borders: false,
        noMargin: true,
        content: wpara(wrun(`Chức vụ: ${chucDanh} Kíp ${kip}`, { size: 28 }), { spBefore: 120 })
      })
    ])
  ], [5000, CW - 5000]);

  const body = [
    wpara(wrun(`Nay tôi làm đơn này kính đề nghị Lãnh đạo ${recipientWsName}, Phòng Hành chính và Lao động xem xét cho tôi được nghỉ phép với nội dung sau:`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } }),
    wpara(wrun(`Thời gian: Từ ngày ${dfStart} đến hết ngày ${dfEnd}`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } }),
    wpara(wrun(`Địa điểm: ${location || locationName}`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } }),
    wpara(wrun(`Lý do: ${reason}${hasLeavePermit ? ' (có lấy giấy phép)' : ''}`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } }),
    wpara(wrun(`Chế độ phép năm: ${currentYear}`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } }),
    wpara(wrun(`Điện thoại liên hệ: ${phone}`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } }),
    wpara(wrun(`Kính mong Lãnh đạo xem xét giải quyết.`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } }),
    wpara(wrun(`Tôi xin chân thành cảm ơn./.`, { size: 28 }), { spBefore: 120, indent: { firstLine: 720 } })
  ].join('');

  const sigWriter = rIds.writer ? wpara(wimage(rIds.writer), { align: 'center' }) : sigPlaceholder();
  const sigManager = rIds.manager ? wpara(wimage(rIds.manager), { align: 'center' }) : sigPlaceholder();

  // Row 1 holds "Nơi nhận" beside each signer's heading + signature space; row 2 holds
  // just the two printed names. Putting both names in their own shared row is what keeps
  // them level — a table row starts all its cells at the same baseline, so it no longer
  // matters that the manager's blank space and the writer's signature image differ in height.
  const footTbl = wtable([
    wtr([
      wtc({
        w: 3000, borders: false, content:
          wpara(wrun('Nơi nhận:', { bold: true, italic: true, size: 24 }), { spBefore: 80 })
          + wpara(wrun('- HCLĐ;', { size: 22 }))
          + wpara(wrun(`- Lưu: ${archiveCode}.`, { size: 22 }))
      }),
      wtc({
        w: 3180, borders: false, content:
          wpara(wrun(chucVuNguoiKy.toUpperCase(), { bold: true, size: 28 }), { align: 'center', spBefore: 80 })
          + sigManager
      }),
      wtc({
        w: 3180, borders: false, content:
          wpara(wrun('NGƯỜI VIẾT ĐƠN', { bold: true, size: 28 }), { align: 'center', spBefore: 80 })
          + sigWriter
      })
    ]),
    wtr([
      wtc({ w: 3000, borders: false, content: emptyP() }),
      wtc({
        w: 3180, borders: false, content:
          wpara(wrun(nguoiKy, { bold: true, size: 28 }), { align: 'center' })
      }),
      wtc({
        w: 3180, borders: false, content:
          wpara(wrun(name, { bold: true, size: 28 }), { align: 'center' })
      })
    ])
  ], [3000, 3180, 3180]);

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<w:body>'
    + hdrTbl + title + kinhGui + info1 + body + footTbl
    + '<w:sectPr>'
    + '<w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1134" w:right="1020" w:bottom="1134" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr>'
    + '</w:body></w:document>';
}

export async function generateLeaveRequestBlob(data: any, config: any, signaturesOverride?: Record<string, string>) {
  const { name } = data;
  const nguoiKy = config?.nguoiKy || 'Nguyễn Văn Nghị';
  
  const rIds: any = {};
  const imagesToAdd: any = [];

  const sigWriter = signaturesOverride?.[name] || SIGNATURES[name];
  const isValidBase64 = (s: string) => s && s.length > 50 && !s.includes("PLACEHOLDER");

  if (isValidBase64(sigWriter)) {
    rIds.writer = 'rIdImg1';
    imagesToAdd.push({ id: 'rIdImg1', data: sigWriter, name: 'sig_writer.png' });
  }

  // The manager/leader's slot is always left blank for a physical wet signature — no
  // image is embedded there at all, so nothing can render over the space they sign in.

  const docXml = buildLeaveRequestDocXml(data, config, rIds);

  const CT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Default Extension="png" ContentType="image/png"/>'
    + '<Default Extension="jpg" ContentType="image/jpeg"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '</Types>';

  const RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '</Relationships>';

  let WRELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
  
  imagesToAdd.forEach((img: any) => {
    WRELS += `<Relationship Id="${img.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.name}"/>`;
  });
  WRELS += '</Relationships>';

  const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:docDefaults><w:rPrDefault><w:rPr>'
    + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    + '<w:sz w:val="24"/><w:szCs w:val="24"/>'
    + '</w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    + '<w:name w:val="Normal"/>'
    + '</w:style>'
    + '</w:styles>';

  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT);
  zip.file('_rels/.rels', RELS);
  zip.file('word/_rels/document.xml.rels', WRELS);
  zip.file('word/styles.xml', STYLES);
  zip.file('word/document.xml', docXml);
  
  imagesToAdd.forEach((img: any) => {
    const base64Data = img.data.replace(/^data:image\/(png|jpeg);base64,/, "");
    zip.file(`word/media/${img.name}`, base64Data, { base64: true });
  });

  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE'
  });
}

export async function exportLeaveRequestDoc(data: any, config: any, signaturesOverride?: Record<string, string>) {
  const blob = await generateLeaveRequestBlob(data, config, signaturesOverride);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Don_xin_nghi_phep_${data.name.replace(/\s+/g, '_')}.docx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export function buildSwapDocXml(swapData: any, config: any, rIds: any = {}) {
  const { date1, date2, person1, person2, shift1, shift2 } = swapData;
  const nguoiKy = config.nguoiKy || 'Nguyễn Văn Nghị';
  const ngayKyVal = config.ngayKy || '';
  const locationName = config?.locationName || config?.location || 'Gia Lai';
  const CW = 9184, HW = Math.floor(CW / 2);

  const companyName = (config?.companyName || 'CÔNG TY THỦY ĐIỆN IALY').toUpperCase();
  const headerWsName = (config?.headerWorkshopName || 'PHÂN XƯỞNG VẬN HÀNH IALY').toUpperCase();
  const docCodeSuffix = config?.documentCodeSuffix || '/VHIALY';
  const archiveCode = docCodeSuffix.replace('/', '') || 'VHIALY';

  const hdrTbl = wtable([
    wtr([
      wtc({
        w: 4100, borders: false, noMargin: true,
        content: wpara(wrun(companyName, { size: 24 }), { align: 'center' })
      }),
      wtc({
        w: CW - 4100, borders: false, noMargin: true,
        content: wpara(wrun('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { bold: true, size: 24 }), { align: 'center' })
      })
    ]),
    wtr([
      wtc({
        w: 4100, borders: false, noMargin: true,
        content: wpara(wrun(headerWsName, { bold: true, size: 24 }), { align: 'center', spAfter: 0 })
          + wline(1400, 4100)
      }),
      wtc({
        w: CW - 4100, borders: false, noMargin: true,
        content: wpara(wrun('Độc lập - Tự do - Hạnh phúc', { bold: true, size: 26 }), { align: 'center', spAfter: 0 })
          + wline(3200, CW - 4100)
      })
    ]),
    wtr([
      wtc({
        w: 4100, borders: false, noMargin: true,
        content: wpara(wrun(`Số:        ${docCodeSuffix}`, { size: 26 }), { align: 'center' })
      }),
      wtc({
        w: CW - 4100, borders: false, noMargin: true,
        content: wpara(wrun(buildDateStr(ngayKyVal, locationName), { italic: true, size: 26 }), { align: 'center' })
      })
    ])
  ], [4100, CW - 4100]);

  const title = wpara(wrun('LỊCH ĐỔI CA', { bold: true, size: 28 }),
    { align: 'center', spBefore: 200, spAfter: 200 });

  const d1 = new Date(date1 + 'T00:00:00');
  const d2 = new Date(date2 + 'T00:00:00');
  const df1 = fmtVN(d1);
  const df2 = fmtVN(d2);

  let timeStr = '';
  let contentLines = [];

  if (shift1 !== 'None' && shift2 !== 'None') {
    if (date1 === date2) {
      timeStr = `Ngày ${df1}.`;
    } else {
      const sortedDates = [d1, d2].sort((a, b) => a.getTime() - b.getTime());
      timeStr = `Ngày ${fmtVN(sortedDates[0])} và ngày ${fmtVN(sortedDates[1])}.`;
    }
    contentLines.push(wpara(wrun(`+ ${person1} nghỉ ca ${shift1} ${df1}, đi ca ${shift2} ${df2}.`, { size: 26 }), { spBefore: 60, indent: { left: 1080 } }));
    contentLines.push(wpara(wrun(`+ ${person2} đi ca ${shift1} ${df1}, nghỉ ca ${shift2} ${df2}.`, { size: 26 }), { spBefore: 60, indent: { left: 1080 } }));
  } else if (shift1 !== 'None' && shift2 === 'None') {
    timeStr = `Ngày ${df1}.`;
    contentLines.push(wpara(wrun(`+ ${person1} nghỉ ca ${shift1} ${df1}.`, { size: 26 }), { spBefore: 60, indent: { left: 1080 } }));
    contentLines.push(wpara(wrun(`+ ${person2} đi ca ${shift1} ${df1}.`, { size: 26 }), { spBefore: 60, indent: { left: 1080 } }));
  } else if (shift1 === 'None' && shift2 !== 'None') {
    timeStr = `Ngày ${df2}.`;
    contentLines.push(wpara(wrun(`+ ${person2} nghỉ ca ${shift2} ${df2}.`, { size: 26 }), { spBefore: 60, indent: { left: 1080 } }));
    contentLines.push(wpara(wrun(`+ ${person1} đi ca ${shift2} ${df2}.`, { size: 26 }), { spBefore: 60, indent: { left: 1080 } }));
  }

  const content = [
    wpara(wrun(`- Thời gian: ${timeStr}`, { size: 26 }), { spBefore: 120, indent: { left: 720 } }),
    wpara(wrun(`- Lịch đổi ca như sau:`, { size: 26 }), { spBefore: 60, indent: { left: 720 } }),
    ...contentLines,
    wpara(wrun(`Các chức danh kiểm tra lại lịch trực của mình và tự chịu trách nhiệm trước Phân xưởng nếu không đi ca theo đúng lịch đã đổi./.`, { size: 26 }), { spBefore: 120, indent: { firstLine: 720 } })
  ].join('');

  const sig1 = rIds.person1 ? wpara(wimage(rIds.person1), { align: 'center' }) : sigPlaceholder();
  const sig2 = rIds.person2 ? wpara(wimage(rIds.person2), { align: 'center' }) : sigPlaceholder();
  const sigManager = rIds.manager ? wpara(wimage(rIds.manager), { align: 'center' }) : sigPlaceholder();

  // One single row of three columns, each column a self-contained stack. Splitting this
  // into several table rows would make the signature row wait for the tall "Nơi nhận"
  // cell to finish, pushing the signatures far below their headings.
  const footTbl = wtable([
    wtr([
      wtc({
        w: 3000, borders: false, content:
          wpara(wrun('Nơi nhận:', { bold: true, italic: true, size: 24 }), { spBefore: 80 })
          + wpara(wrun('- LĐPX;', { size: 22 }))
          + wpara(wrun('- Các cá nhân (để t/h);', { size: 22 }))
          + wpara(wrun(`- Lưu: ${archiveCode}.`, { size: 22 }))
      }),
      wtc({
        w: 3180, borders: false, content:
          wpara(wrun('NGƯỜI ĐỔI CA', { bold: true, size: 24 }), { align: 'center', spBefore: 80 })
          + sig1
          + wpara(wrun(person1, { bold: true, size: 24 }), { align: 'center' })
          + wpara(wrun('QUẢN ĐỐC', { bold: true, size: 24 }), { align: 'center', spBefore: 240 })
          + sigManager
          + wpara(wrun(nguoiKy, { bold: true, size: 24 }), { align: 'center' })
      }),
      wtc({
        w: 3180, borders: false, content:
          wpara(wrun('NGƯỜI ĐI CA THAY', { bold: true, size: 24 }), { align: 'center', spBefore: 80 })
          + sig2
          + wpara(wrun(person2, { bold: true, size: 24 }), { align: 'center' })
      })
    ])
  ], [3000, 3180, 3180]);

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<w:body>'
    + hdrTbl + title + content + footTbl
    + '<w:sectPr>'
    + '<w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1134" w:right="1020" w:bottom="1134" w:left="1701" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr>'
    + '</w:body></w:document>';
}

export async function generateSwapBlob(swapData: any, config: any, signaturesOverride?: Record<string, string>) {
  const { person1, person2 } = swapData;
  const nguoiKy = config.nguoiKy || 'Nguyễn Văn Nghị';
  
  const rIds: any = {};
  const imagesToAdd: any = [];

  const sig1 = signaturesOverride?.[person1] || SIGNATURES[person1];
  const sig2 = signaturesOverride?.[person2] || SIGNATURES[person2];

  const isValidBase64 = (s: string) => s && s.length > 50 && !s.includes("PLACEHOLDER");

  if (isValidBase64(sig1)) {
    rIds.person1 = 'rIdImg1';
    imagesToAdd.push({ id: 'rIdImg1', data: sig1, name: 'sig1.png' });
  }
  if (isValidBase64(sig2)) {
    rIds.person2 = 'rIdImg2';
    imagesToAdd.push({ id: 'rIdImg2', data: sig2, name: 'sig2.png' });
  }

  // The manager/leader's slot is always left blank for a physical wet signature — no
  // image is embedded there at all, so nothing can render over the space they sign in.

  const docXml = buildSwapDocXml(swapData, config, rIds);

  const CT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Default Extension="png" ContentType="image/png"/>'
    + '<Default Extension="jpg" ContentType="image/jpeg"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '</Types>';

  const RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '</Relationships>';

  let WRELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
  
  imagesToAdd.forEach((img: any) => {
    WRELS += `<Relationship Id="${img.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.name}"/>`;
  });
  WRELS += '</Relationships>';

  const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:docDefaults><w:rPrDefault><w:rPr>'
    + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
    + '<w:sz w:val="24"/><w:szCs w:val="24"/>'
    + '</w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    + '<w:name w:val="Normal"/>'
    + '</w:style>'
    + '</w:styles>';

  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT);
  zip.file('_rels/.rels', RELS);
  zip.file('word/_rels/document.xml.rels', WRELS);
  zip.file('word/styles.xml', STYLES);
  zip.file('word/document.xml', docXml);
  
  imagesToAdd.forEach((img: any) => {
    // Remove data:image/png;base64, prefix if exists
    const base64Data = img.data.replace(/^data:image\/(png|jpeg);base64,/, "");
    zip.file(`word/media/${img.name}`, base64Data, { base64: true });
  });

  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE'
  });
}

export async function exportSwapDoc(swapData: any, config: any, signaturesOverride?: Record<string, string>) {
  const blob = await generateSwapBlob(swapData, config, signaturesOverride);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Lich_doi_ca_${swapData.person1.replace(/\s+/g, '_')}_${swapData.person2.replace(/\s+/g, '_')}.docx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export async function exportWord(currentResult: any, config: any) {
  const blob = await generateWordBlob(currentResult, config);
  if (!blob) return;

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const fname = currentResult.allResults && currentResult.allResults.length > 1
    ? 'Lich_truc_thay_' + currentResult.allResults.map((r: any) => r.ten.split(' ').pop()).join('_') + '.docx'
    : 'Lich_truc_thay_' + currentResult.ten.replace(/\s+/g, '_') + '.docx';
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export async function exportAllDocsZip(currentResult: any, leavesList: any[], config: any, signatures: any) {
  const zip = new JSZip();
  
  // 1. Generate Lịch Trực Thay Ca Word Blob
  const scheduleBlob = await generateWordBlob(currentResult, config);
  if (scheduleBlob) {
    const fname = currentResult.allResults && currentResult.allResults.length > 1
      ? '01_Lich_truc_thay_ca_phoi_hop.docx'
      : `01_Lich_truc_thay_ca_${currentResult.ten.replace(/\s+/g, '_')}.docx`;
    zip.file(fname, scheduleBlob);
  }

  // 2. Generate Đơn Xin Nghỉ Phép Word Blobs
  for (let i = 0; i < leavesList.length; i++) {
    const leave = leavesList[i];
    const leaveData = {
      name: leave.name,
      birthYear: leave.birthYear || "1980",
      chucDanh: leave.chucDanh,
      kip: String(leave.kip),
      startDate: leave.startDate,
      endDate: leave.endDate,
      reason: leave.reason || "Giải quyết việc riêng gia đình",
      phone: leave.phone || "",
      leaveYear: String(new Date(leave.startDate).getFullYear() || new Date().getFullYear()),
      location: leave.location || config?.locationName || config?.location || "Gia Lai",
      // Carried over from the saved request, otherwise batch-exported documents would
      // silently drop the "(có lấy giấy phép)" note that single exports include.
      hasLeavePermit: leave.hasLeavePermit === true
    };
    
    const leaveBlob = await generateLeaveRequestBlob(leaveData, config, signatures);
    if (leaveBlob) {
      zip.file(`02_Don_xin_nghi_phep_${leave.name.replace(/\s+/g, '_')}.docx`, leaveBlob);
    }
  }

  // 3. Generate Zip Blob and trigger download
  const zipBlob = await zip.generateAsync({ type: 'blob' });
  const url = URL.createObjectURL(zipBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Lich_Truc_Thay_Ca_${Date.now()}.zip`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}