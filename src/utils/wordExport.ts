import JSZip from 'jszip';
import { fmtVN, abbrev } from './shiftHelpers';

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
  const indent = o.indent ? '<w:ind w:left="' + o.indent.left + '"/>' : '';
  return '<w:p><w:pPr><w:jc w:val="' + align + '"/><w:spacing w:before="' + spB + '" w:after="' + spA + '"/>' + indent + '</w:pPr>' + content + '</w:p>';
}

function emptyP(spB: number = 0, spA: number = 0) {
  return '<w:p><w:pPr><w:spacing w:before="' + spB + '" w:after="' + spA + '"/></w:pPr></w:p>';
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
  return '<w:tc><w:tcPr>'
    + '<w:tcW w:w="' + w + '" w:type="dxa"/>' + gs + vm
    + makeBorders(borders) + sh
    + '<w:vAlign w:val="top"/>'
    + '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
    + '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>'
    + '</w:tcPr>' + content + '</w:tc>';
}

function wtr(cells: string[]) {
  return '<w:tr><w:trPr/>' + cells.join('') + '</w:tr>';
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

function buildDateStr(val: string) {
  if (val) {
    const nd = new Date(val + 'T00:00:00');
    return 'Gia Lai, ngày ' + nd.getDate() + ' tháng ' + (nd.getMonth() + 1) + ' năm ' + nd.getFullYear();
  }
  return 'Gia Lai, ngày      tháng      năm     ';
}

export function buildDocXml(currentResult: any, config: any) {
  const d = currentResult;
  const soVanBan = config.soVanBan || '';
  const nguoiKy = config.nguoiKy || 'Quản Đốc';
  const ngayKyVal = config.ngayKy || '';

  const CW = 9360, c0 = 2200, c1 = 1500, HW = Math.floor(CW / 2);
  const GRAY = 'D9D9D9';
  const shiftOrd: Record<string, number> = { N: 0, C: 1, K: 2 };
  const MAX_PER_ROW = 4;

  const allResults = d.allResults;
  const extraRows = d.extraRows || [];

  const swapRows = extraRows.filter((r: any) => r.isSwap);
  const chainRows = extraRows.filter((r: any) => r.isCKChain && !r.isSwap);

  const hdrTbl = twoCol(
    wpara(wrun('CÔNG TY THỦY ĐIỆN IALY', { size: 22 }), { align: 'center', spAfter: 40 })
    + wpara(wrun('PHÂN XƯỞNG VẬN HÀNH IALY', { bold: true, size: 22, underline: true }), { align: 'center' }),
    wpara(wrun('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { bold: true, size: 20 }), { align: 'center', spAfter: 40 })
    + wpara(wrun('Độc lập - Tự do - Hạnh phúc', { italic: true, size: 22, underline: true }), { align: 'center' }),
    CW
  );
  const soNgayTbl = twoCol(
    wpara(wrun('Số: ' + (soVanBan || '      ') + '/VHIALY', { size: 22 }), { align: 'center', spBefore: 40, spAfter: 40 }),
    wpara(wrun(buildDateStr(ngayKyVal), { italic: true, size: 22 }), { align: 'center' }),
    CW
  );
  const title = wpara(wrun('LỊCH TRỰC THAY CA VẬN HÀNH', { bold: true, size: 28 }),
    { align: 'center', spBefore: 120, spAfter: 120 });

  function buildPersonRows(res: any) {
    const rows = res.ketQua.slice();
    chainRows.forEach((ex: any) => {
      if (ex.absentKip === res.kip) rows.push(ex);
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
    wtc({ w: c0, shading: GRAY, content: wpara(wrun('Tên người cần được\ntrực thay', { bold: true, size: 22 }), { align: 'center' }) }),
    wtc({ w: c1, shading: GRAY, content: wpara(wrun('Chức danh\nhiện tại', { bold: true, size: 22 }), { align: 'center' }) }),
    wtc({ w: CW - c0 - c1, gridSpan: nCols, shading: GRAY, content: wpara(wrun('Tên người trực thay', { bold: true, size: 22 }), { align: 'center' }) })
  ]));

  allResults.forEach((res: any) => {
    const rows = buildPersonRows(res);
    let groups = [];
    for (let i = 0; i < rows.length; i += nCols) groups.push(rows.slice(i, i + nCols));
    if (!groups.length) groups = [[]];

    const nghiStr = '(nghỉ phép từ\n' + fmtVN(res.start) + ' đến\n' + fmtVN(res.end) + ')';
    const chucDanh = res.chucDanh || d.chucDanh || '';

    groups.forEach((grp, gi) => {
      const isFirst = (gi === 0);
      const caRow = [];
      if (isFirst) {
        caRow.push(wtc({
          w: c0, vMerge: 'restart',
          content: wpara(wrun(res.ten, { size: 22 }), { align: 'center', spAfter: 20 })
            + wpara(wrun(nghiStr, { italic: true, size: 20 }), { align: 'center' })
        }));
        caRow.push(wtc({
          w: c1, vMerge: 'restart',
          content: wpara(wrun(chucDanh, { size: 22 }), { align: 'center' })
        }));
      } else {
        caRow.push(wtc({ w: c0, vMerge: 'cont', content: emptyP() }));
        caRow.push(wtc({ w: c1, vMerge: 'cont', content: emptyP() }));
      }
      for (let i = 0; i < nCols; i++) {
        const it = grp[i];
        caRow.push(wtc({
          w: colW[i + 2], shading: it ? GRAY : '',
          content: wpara(wrun(it ? it.ca + '-' + fmtVN(it.ngay).slice(0, 5) : '', { bold: true, size: 22 }), { align: 'center' })
        }));
      }
      tableRows.push(wtr(caRow));

      const tenRow = [
        wtc({ w: c0, vMerge: 'cont', content: emptyP() }),
        wtc({ w: c1, vMerge: 'cont', content: emptyP() })
      ];
      for (let i = 0; i < nCols; i++) {
        const it = grp[i];
        tenRow.push(wtc({
          w: colW[i + 2],
          content: it ? wpara(wrun(abbrev(it.nguoiThay), { size: 22 }), { align: 'center' }) : emptyP()
        }));
      }
      tableRows.push(wtr(tenRow));
    });
  });

  const swapNotes: string[] = [];
  allResults.forEach((res: any) => {
    res.ketQua.forEach((it: any) => {
      if (it.conflictNote && it.conflictNote.includes('Hoán đổi')) {
        const dateStr = it.ca + '/' + fmtVN(it.ngay).slice(0, 5);
        swapNotes.push(`${abbrev(it.nguoiThay)} trực thay ${abbrev(res.ten)} ca ${dateStr}`);
      }
    });
  });

  extraRows.forEach((ex: any) => {
    if (ex.isSwap) {
      const dateStr = ex.ca + '/' + fmtVN(ex.ngay).slice(0, 5);
      swapNotes.push(`${abbrev(ex.nguoiThay)} trực thay ${abbrev(ex.absentTen)} ca ${dateStr}`);
    }
  });

  if (swapNotes.length > 0) {
    const cdLabel = d.chucDanh || config.chucDanh || 'Trưởng ca';
    let noteContent = wpara(wrun('Tại ' + cdLabel + ' :', { bold: true, size: 22 }), { spBefore: 40, spAfter: 10 });
    swapNotes.forEach(line => {
      noteContent += wpara(wrun('- ' + line, { size: 22 }), { indent: { left: 360 }, spAfter: 10 });
    });
    tableRows.push(wtr([
      wtc({ w: CW, gridSpan: nCols + 2, borders: false, content: noteContent })
    ]));
  }

  const mainTbl = wtable(tableRows, colW);

  const note = wpara(
    wrun('Ghi chú: Các chức danh kiểm tra lại lịch trực của mình, nếu có gì vướng mắc phải báo lại PX để kiểm tra và điều chỉnh kịp thời./.',
      { italic: true, size: 20 }),
    { spBefore: 120, spAfter: 60 }
  );

  const footTbl = wtable([
    wtr([
      wtc({
        w: HW, borders: false, content:
          wpara(wrun('Nơi nhận:', { bold: true, italic: true, size: 22 }), { spBefore: 80 })
          + wpara(wrun('- Các kíp (để t/hiện)', { italic: true, size: 22 }))
          + wpara(wrun('- Lưu: VHIALY', { italic: true, size: 22 }))
      }),
      wtc({
        w: CW - HW, borders: false, content:
          wpara(wrun('QUẢN ĐỐC', { bold: true, size: 22 }), { align: 'center', spBefore: 80 })
          + emptyP(500, 0) + emptyP(500, 0) + emptyP(500, 0)
          + wpara(wrun(nguoiKy, { bold: true, size: 22 }), { align: 'center' })
      })
    ])
  ], [HW, CW - HW]);

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:body>'
    + hdrTbl + soNgayTbl + title + mainTbl + note + footTbl
    + '<w:sectPr>'
    + '<w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr>'
    + '</w:body></w:document>';
}

export async function exportWord(currentResult: any, config: any) {
  if (!currentResult) return;
  
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

  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE'
  });

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
