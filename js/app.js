/**
 * Exp Table Generator - Web Version
 * Genera documentos Word con tablas a partir de Excel
 */

// Constants - match Python implementation
const MONTH_MAP = {
  Enero: 'ene', Febrero: 'feb', Marzo: 'mar', Abril: 'abr',
  Mayo: 'may', Junio: 'jun', Julio: 'jul', Agosto: 'ago',
  Septiembre: 'sep', Octubre: 'oct', Noviembre: 'nov', Diciembre: 'dic',
};

const COUNTRY_KEYWORDS = {
  Argentina: 'Argentina', argentina: 'Argentina',
  Perú: 'Perú', Peru: 'Perú', Colombia: 'Colombia', Chile: 'Chile',
  Bolivia: 'Bolivia', Ecuador: 'Ecuador', Brasil: 'Brasil',
  Paraguay: 'Paraguay', Uruguay: 'Uruguay', México: 'México', Mexico: 'México',
  Panamá: 'Panamá', Panama: 'Panamá', 'Costa Rica': 'Costa Rica',
  Honduras: 'Honduras', 'El Salvador': 'El Salvador', Guatemala: 'Guatemala',
  Nicaragua: 'Nicaragua', Venezuela: 'Venezuela',
  'República Dominicana': 'República Dominicana',
};

const ENTITY_COUNTRY_HINTS = {
  Nación: 'Argentina', Nacion: 'Argentina', 'Buenos Aires': 'Argentina',
  CABA: 'Argentina', 'Provincia de': 'Argentina',
  'Ministerio de Seguridad': 'Argentina', 'Banco Hipotecario': 'Argentina',
};

const SPECIAL_SOURCES = ['(vacío)', '(auto-incremento)', '(extraer país)'];
const DATE_KEYWORDS_START = new Set(['inicio', 'desde', 'from', 'start']);
const DATE_KEYWORDS_END = new Set(['fin', 'hasta', 'end', 'until']);

// State
let templateInfo = null;
let excelWorkbook = null;
let excelHeaders = {};
let templateFile = null;
let excelFile = null;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function normalize(text) {
  if (!text) return '';
  return text.toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .trim();
}

function wordOverlap(a, b) {
  const wa = new Set(normalize(a).split(/\s+/).filter(Boolean));
  const wb = new Set(normalize(b).split(/\s+/).filter(Boolean));
  if (!wa.size || !wb.size) return 0;
  const intersection = [...wa].filter(x => wb.has(x)).length;
  return intersection / Math.max(wa.size, wb.size);
}

function convertDate(dateStr) {
  if (!dateStr) return '';
  const parts = String(dateStr).trim().split(/\s+/);
  if (parts.length === 2) {
    const abbr = MONTH_MAP[parts[0]] || parts[0].substring(0, 3).toLowerCase();
    return `${abbr}-${parts[1].slice(-2)}`;
  }
  return String(dateStr);
}

function extractCountry(text) {
  if (!text) return '';
  const t = String(text);
  for (const [kw, country] of Object.entries(COUNTRY_KEYWORDS)) {
    if (t.includes(kw)) return country;
  }
  for (const [kw, country] of Object.entries(ENTITY_COUNTRY_HINTS)) {
    if (t.includes(kw)) return country;
  }
  return '';
}

function colIndexToLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter || 'A';
}

function colLetterToIndex(letter) {
  let index = 0;
  const upper = String(letter).toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index;
}

function formatExcelDate(val) {
  if (val == null) return '';
  if (val instanceof Date) {
    const d = val;
    const pad = (n) => String(n).padStart(2, '0');
    const day = d.getDate();
    const month = d.getMonth() + 1;
    const year = d.getFullYear();
    if (day === 1) return `${pad(month)}/${year}`;
    return `${pad(day)}/${pad(month)}/${year}`;
  }
  return String(val);
}

// ---------------------------------------------------------------------------
// Parse Word Template (docx)
// ---------------------------------------------------------------------------
function getCellText(cell) {
  const ts = cell.getElementsByTagName ? cell.getElementsByTagName('*') : [];
  let text = '';
  for (const t of ts) {
    if ((t.localName || t.nodeName || '').split(':').pop() === 't') {
      text += (t.textContent || '');
    }
  }
  return text;
}

async function parseTemplate(file) {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docXml = await zip.file('word/document.xml')?.async('string');
  if (!docXml) throw new Error('El archivo Word no contiene document.xml');

  const parser = new DOMParser();
  const doc = parser.parseFromString(docXml, 'application/xml');
  let body = doc.documentElement;
  const findBody = (node) => {
    if (!node) return null;
    const tag = (node.localName || node.nodeName || '').split(':').pop();
    if (tag === 'body') return node;
    const chs = node.childNodes || [];
    for (let i = 0; i < chs.length; i++) {
      if (chs[i].nodeType === 1) {
        const found = findBody(chs[i]);
        if (found) return found;
      }
    }
    return null;
  };
  body = findBody(doc.documentElement) || body;

  let title = '';
  const paras = [];
  const collectParas = (node) => {
    const tag = (node.localName || node.nodeName || '').split(':').pop();
    if (tag === 'p') paras.push(node);
    const chs = node.childNodes || [];
    for (let i = 0; i < chs.length; i++) {
      if (chs[i].nodeType === 1) collectParas(chs[i]);
    }
  };
  collectParas(body);
  for (const p of paras) {
    const t = getCellText(p);
    if (t.trim()) {
      title = t.trim();
      break;
    }
  }

  const tables = [];
  const walk = (node) => {
    const tag = (node.localName || node.nodeName || '').split(':').pop();
    if (tag === 'tbl') {
      tables.push(node);
      return true;
    }
    const chs = node.childNodes || [];
    for (let i = 0; i < chs.length; i++) {
      if (chs[i].nodeType === 1 && walk(chs[i])) return true;
    }
    return false;
  };
  const kids = body.childNodes || body.children || [];
  for (let i = 0; i < kids.length; i++) {
    const ch = kids[i];
    if (ch.nodeType === 1 && walk(ch)) break;
  }
  if (!tables.length) throw new Error('El template no contiene tablas');

  const table = tables[0];
  const rows = [];
  const collect = (node, tag, out) => {
    const t = (node.localName || node.nodeName || '').split(':').pop();
    if (t === tag) out.push(node);
    const chs = node.childNodes || [];
    for (let i = 0; i < chs.length; i++) {
      if (chs[i].nodeType === 1) collect(chs[i], tag, out);
    }
  };
  collect(table, 'tr', rows);
  if (!rows.length) throw new Error('La tabla no tiene filas');

  const gridCols = [];
  collect(table, 'gridCol', gridCols);
  const widths = gridCols.map(g => parseInt(g.getAttribute('w:w') || g.getAttribute('w') || '1500', 10));

  const headerRow = rows[0];
  const headerCells = [];
  collect(headerRow, 'tc', headerCells);

  const columns = [];
  for (let i = 0; i < headerCells.length; i++) {
    const headerText = getCellText(headerCells[i]).trim();
    const width = widths[i] !== undefined ? widths[i] : 1500;
    let bold = false;
    if (rows.length > 1) {
      const dataCells = [];
      collect(rows[1], 'tc', dataCells);
      if (dataCells[i]) {
        const bEls = [];
        collect(dataCells[i], 'b', bEls);
        if (bEls.length) bold = true;
      }
    }
    columns.push({ header: headerText, width, bold, align: 'CENTER (1)' });
  }

  return {
    title,
    page: { orientation: 1, width: 11906, height: 16838, left_margin: 1440, right_margin: 1440, top_margin: 1440, bottom_margin: 1440 },
    columns,
  };
}

// ---------------------------------------------------------------------------
// Excel (SheetJS)
// sheet_to_json con header:1 devuelve array de filas: data[fila][columna]
// fila 0 = Excel fila 1, columna 0 = Excel columna A
// ---------------------------------------------------------------------------
function getSheetData(wb, sheetName) {
  const sheet = wb.Sheets[sheetName];
  if (!sheet) return null;
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    raw: false,
  });
}

function readExcelHeaders(wb, sheetName, headerRow) {
  const data = getSheetData(wb, sheetName);
  if (!data) throw new Error(`Hoja '${sheetName}' no encontrada. Disponibles: ${wb.SheetNames.join(', ')}`);

  const headers = {};
  const rowIdx1 = headerRow - 1;
  const rowIdx2 = headerRow;
  const row1 = data[rowIdx1] || [];
  const row2 = data[rowIdx2] || [];
  const maxCol = Math.min(Math.max(row1.length, row2.length), 700);

  for (let col = 0; col < maxCol; col++) {
    const colLetter = colIndexToLetter(col + 1);
    const v1 = row1[col];
    const v2 = row2[col];

    let combined = '';
    if (v1 != null && v2 != null && String(v1).trim() && String(v2).trim()) {
      combined = `${v1} / ${v2}`;
    } else if (v1 != null && String(v1).trim()) {
      combined = String(v1);
    } else if (v2 != null && String(v2).trim()) {
      combined = String(v2);
    }

    if (combined.trim()) headers[colLetter] = combined.trim();
  }
  return headers;
}

function peekExcelRows(wb, sheetName, headerRow) {
  const data = getSheetData(wb, sheetName);
  if (!data) return `ERROR: Hoja '${sheetName}' no encontrada.`;
  const lines = [`Hoja: ${sheetName}\n`];
  let shown = 0;
  for (let r = headerRow + 1; r < data.length && shown < 20; r++) {
    const row = data[r] || [];
    const firstVal = row[0];
    if (firstVal != null && String(firstVal).trim()) {
      lines.push(`  Fila ${r + 1}: ${String(firstVal).substring(0, 90)}`);
      shown++;
    }
  }
  if (shown >= 20) lines.push('  ...');
  return lines.join('\n');
}

function readExcelData(wb, sheetName, rowNumbers, mapping) {
  const data = getSheetData(wb, sheetName);
  if (!data) return [];
  const rows = [];

  rowNumbers.forEach((rowNum, seqIdx) => {
    const rowIdx = rowNum - 1;
    if (rowIdx < 0 || rowIdx >= data.length) return;

    const excelRow = data[rowIdx] || [];
    const rowData = {};
    const seq = seqIdx + 1;

    mapping.forEach(m => {
      const header = m.header;
      const source = m.source;

      if (!source || source === '(vacío)') {
        rowData[header] = '';
      } else if (source === '(auto-incremento)') {
        rowData[header] = String(seq);
      } else if (source === '(extraer país)') {
        const colIdx = colLetterToIndex(m.from_col || 'D') - 1;
        const raw = excelRow[colIdx] != null ? String(excelRow[colIdx]) : '';
        rowData[header] = extractCountry(raw);
      } else {
        const colIdx = colLetterToIndex(source) - 1;
        let raw = excelRow[colIdx];
        const formatType = m.format || 'valor_tal_cual';

        if (formatType === 'valor_tal_cual') {
          if (raw == null) rowData[header] = '';
          else if (raw instanceof Date) rowData[header] = formatExcelDate(raw);
          else if (typeof raw === 'number') rowData[header] = raw;
          else rowData[header] = String(raw);
        } else if (formatType === 'short_date') {
          const val = raw != null ? String(raw) : '';
          rowData[header] = convertDate(val);
        } else {
          rowData[header] = raw != null ? String(raw) : '';
        }
      }
    });
    rows.push(rowData);
  });

  return rows;
}

// ---------------------------------------------------------------------------
// Auto-mapping
// ---------------------------------------------------------------------------
function autoMap(templateCols, headers) {
  const mapping = [];
  const used = new Set();

  const sortedHeaders = Object.entries(headers).sort((a, b) => {
    if (a[0].length !== b[0].length) return a[0].length - b[0].length;
    return a[0].localeCompare(b[0]);
  });

  templateCols.forEach(col => {
    const headerNorm = normalize(col.header);
    let bestSource = '';
    let bestScore = 0;
    let fmt = 'valor_tal_cual';

    if (['no', 'no.', 'n', '#', 'numero'].includes(headerNorm)) {
      mapping.push({ ...col, source: '(auto-incremento)', format: '' });
      return;
    }
    if (['pais', 'country', 'paises'].includes(headerNorm)) {
      let entityCol = 'D';
      for (const [letter, eh] of sortedHeaders) {
        if (normalize(eh).includes('entidad') || normalize(eh).includes('contratante')) {
          entityCol = letter;
          break;
        }
      }
      mapping.push({ ...col, source: '(extraer país)', from_col: entityCol, format: '' });
      return;
    }

    const words = new Set(headerNorm.split(/\s+/));
    const isStartDate = [...DATE_KEYWORDS_START].some(w => words.has(w));
    const isEndDate = [...DATE_KEYWORDS_END].some(w => words.has(w));
    if (isStartDate || isEndDate) fmt = 'short_date';

    for (const [letter, excelHeader] of sortedHeaders) {
      if (used.has(letter)) continue;
      let score = wordOverlap(col.header, excelHeader);
      const ehNorm = normalize(excelHeader);
      if (isStartDate && (ehNorm.includes('desde') || ehNorm.includes('inicio') || ehNorm.includes('from'))) score += 0.4;
      if (isEndDate && (ehNorm.includes('hasta') || ehNorm.includes('fin') || ehNorm.includes('until'))) score += 0.4;

      if (score > bestScore) {
        bestScore = score;
        bestSource = letter;
      }
    }

    if (bestSource && bestScore > 0.1) {
      used.add(bestSource);
      mapping.push({ ...col, source: bestSource, format: fmt });
    } else {
      mapping.push({ ...col, source: '(vacío)', format: fmt });
    }
  });

  return mapping;
}

// ---------------------------------------------------------------------------
// Generate Word (docx library)
// ---------------------------------------------------------------------------
function buildDocument(dataRows, templateInfo, mapping) {
  const d = typeof docx !== 'undefined' ? docx : window.docx;
  if (!d) throw new Error('La librería docx no está cargada.');
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle } = d;

  const tableRows = [];
  const widths = mapping.map(m => m.width || 1500);

  // Header row
  const headerCells = mapping.map((m, i) =>
    new TableCell({
      children: [
        new Paragraph({
          children: [new TextRun({ text: m.header, bold: true })],
          alignment: AlignmentType.CENTER,
        }),
      ],
      shading: { fill: 'BFBFBF' },
      width: { size: widths[i] || 1500, type: WidthType.DXA },
    })
  );
  tableRows.push(new TableRow({ children: headerCells }));

  // Data rows
  dataRows.forEach(rowData => {
    const cells = mapping.map((m, colIdx) => {
      let value = rowData[m.header];
      const formatType = m.format || 'valor_tal_cual';

      if (value == null) value = '';
      else if (formatType === 'valor_tal_cual' && typeof value === 'number') {
        value = Number.isInteger(value) ? String(value) : String(value);
      } else if (typeof value !== 'string') value = String(value);

      return new TableCell({
        children: [
          new Paragraph({
            children: [new TextRun({ text: value, bold: m.bold })],
            alignment: m.align === 'JUSTIFY (3)' ? AlignmentType.JUSTIFIED :
              m.align === 'LEFT (0)' ? AlignmentType.LEFT :
              m.align === 'RIGHT (2)' ? AlignmentType.RIGHT : AlignmentType.CENTER,
          }),
        ],
        width: { size: widths[colIdx] || 1500, type: WidthType.DXA },
      });
    });
    tableRows.push(new TableRow({ children: cells }));
  });

  const table = new Table({
    rows: tableRows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4 },
      bottom: { style: BorderStyle.SINGLE, size: 4 },
      left: { style: BorderStyle.SINGLE, size: 4 },
      right: { style: BorderStyle.SINGLE, size: 4 },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4 },
      insideVertical: { style: BorderStyle.SINGLE, size: 4 },
    },
  });

  const doc = new Document({
    sections: [{
      properties: {},
      children: [
        new Paragraph({
          children: [new TextRun({ text: templateInfo.title || '', bold: true })],
          alignment: AlignmentType.CENTER,
        }),
        table,
      ],
    }],
  });

  return doc;
}

// ---------------------------------------------------------------------------
// UI
// ---------------------------------------------------------------------------
function setStatus(msg, type = '') {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className = 'status' + (type ? ` ${type}` : '');
}

function parseRowsInput(raw) {
  return raw.replace(/,|;/g, ' ').split(/\s+/)
    .map(s => s.trim()).filter(Boolean)
    .flatMap(p => {
      if (p.includes('-') && !p.startsWith('-')) {
        const [a, b] = p.split('-').map(Number);
        const nums = [];
        for (let i = a; i <= b; i++) nums.push(i);
        return nums;
      }
      return [parseInt(p, 10)];
    })
    .filter(n => !isNaN(n));
}

function getMappingFromUI() {
  const mapping = [];
  document.querySelectorAll('#mappingBody tr').forEach(tr => {
    const header = tr.dataset.header;
    const sourceSelect = tr.querySelector('.source-select');
    const formatSelect = tr.querySelector('.format-select');
    const width = parseInt(tr.dataset.width, 10) || 1500;
    const bold = tr.dataset.bold === 'true';
    const align = tr.dataset.align || 'CENTER (1)';

    const sourceVal = sourceSelect?.value || '(vacío)';
    let source = sourceVal;
    if (sourceVal.includes(':')) source = sourceVal.split(':')[0].trim();

    const formatVal = formatSelect?.value || 'valor_tal_cual';
    const format = formatVal === '(ninguno)' ? 'valor_tal_cual' : formatVal;

    let from_col = 'D';
    if (source === '(extraer país)') {
      for (const [letter, eh] of Object.entries(excelHeaders)) {
        if (normalize(eh).includes('entidad') || normalize(eh).includes('contratante')) {
          from_col = letter;
          break;
        }
      }
    }

    mapping.push({
      header,
      source: source === '(vacío)' ? '' : source,
      format,
      width,
      bold,
      align,
      from_col,
    });
  });
  return mapping;
}

function renderMapping(mapping) {
  const tbody = document.getElementById('mappingBody');
  tbody.innerHTML = '';

  const excelOptions = ['(vacío)', '(auto-incremento)', '(extraer país)'];
  Object.entries(excelHeaders).sort((a, b) => {
    if (a[0].length !== b[0].length) return a[0].length - b[0].length;
    return a[0].localeCompare(b[0]);
  }).forEach(([letter, h]) => excelOptions.push(`${letter}: ${h.substring(0, 40)}`));

  const formatOptions = [
    'valor_tal_cual',
    'short_date',
  ];

  mapping.forEach(m => {
    const tr = document.createElement('tr');
    tr.dataset.header = m.header;
    tr.dataset.width = m.width || 1500;
    tr.dataset.bold = String(m.bold || false);
    tr.dataset.align = m.align || 'CENTER (1)';

    const sourceSelect = document.createElement('select');
    sourceSelect.className = 'source-select';
    let selectedSource = '(vacío)';
    if (m.source === '(auto-incremento)') selectedSource = '(auto-incremento)';
    else if (m.source === '(extraer país)') selectedSource = '(extraer país)';
    else if (m.source) {
      const match = excelOptions.find(o => o.startsWith(m.source + ':'));
      if (match) selectedSource = match;
    }
    excelOptions.forEach(opt => {
      const optEl = document.createElement('option');
      optEl.value = opt;
      optEl.textContent = opt;
      if (opt === selectedSource) optEl.selected = true;
      sourceSelect.appendChild(optEl);
    });

    const formatSelect = document.createElement('select');
    formatSelect.className = 'format-select';
    const fmt = m.format || 'valor_tal_cual';
    formatOptions.forEach(opt => {
      const optEl = document.createElement('option');
      optEl.value = opt;
      optEl.textContent = opt === 'valor_tal_cual' ? 'valor_tal_cual (tal cual en Excel)' : 'short_date (ej: ago-21)';
      if (opt === fmt) optEl.selected = true;
      formatSelect.appendChild(optEl);
    });

    tr.innerHTML = `
      <td>${m.header}</td>
      <td>→</td>
      <td></td>
      <td></td>
    `;
    tr.querySelector('td:nth-child(3)').appendChild(sourceSelect);
    tr.querySelector('td:nth-child(4)').appendChild(formatSelect);
    tbody.appendChild(tr);
  });

  document.getElementById('mappingHint').textContent = `Se auto-mapearon ${mapping.length} columnas. Ajustá si es necesario.`;
}

function tryBuildMapping() {
  if (!templateInfo || !Object.keys(excelHeaders).length) return;
  const mapping = autoMap(templateInfo.columns, excelHeaders);
  renderMapping(mapping);
}

// Event handlers
document.getElementById('templateInput').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  templateFile = file;
  document.getElementById('templateName').textContent = file.name;
  try {
    setStatus('Leyendo template...');
    templateInfo = await parseTemplate(file);
    setStatus(`Template cargado: ${templateInfo.columns.length} columnas detectadas.`, 'success');
    tryBuildMapping();
  } catch (err) {
    setStatus(`Error: ${err.message}`, 'error');
    console.error(err);
  }
});

document.getElementById('excelInput').addEventListener('change', async (e) => {
  excelFile = e.target.files[0];
  if (!excelFile) return;
  document.getElementById('excelName').textContent = excelFile.name;
  document.getElementById('loadExcel').click();
});

document.getElementById('loadExcel').addEventListener('click', async () => {
  if (!excelFile) {
    setStatus('Seleccioná un archivo Excel primero.', 'error');
    return;
  }
  const sheetName = document.getElementById('sheetName').value.trim() || 'ESP';
  const headerRow = parseInt(document.getElementById('headerRow').value, 10) || 3;

  try {
    setStatus('Leyendo Excel...');
    const arrayBuffer = await excelFile.arrayBuffer();
    excelWorkbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false, cellNF: false });
    excelHeaders = readExcelHeaders(excelWorkbook, sheetName, headerRow);
    const preview = peekExcelRows(excelWorkbook, sheetName, headerRow);
    document.getElementById('excelPreview').innerHTML = `<pre>${preview}</pre>`;
    setStatus(`Excel cargado: ${Object.keys(excelHeaders).length} columnas detectadas.`, 'success');
    tryBuildMapping();
  } catch (err) {
    setStatus(`Error: ${err.message}`, 'error');
    console.error(err);
  }
});

document.getElementById('generateBtn').addEventListener('click', async () => {
  if (!templateInfo) {
    setStatus('Seleccioná un template Word primero.', 'error');
    return;
  }
  if (!excelWorkbook) {
    setStatus('Seleccioná y cargá un Excel primero.', 'error');
    return;
  }

  const rowsRaw = document.getElementById('rowsInput').value;
  const rowNumbers = parseRowsInput(rowsRaw);
  if (!rowNumbers.length) {
    setStatus('Ingresá al menos un número de fila.', 'error');
    return;
  }

  const mapping = getMappingFromUI();
  const hasMapped = mapping.some(m => m.source && m.source !== '(vacío)');
  if (!hasMapped) {
    setStatus('Mapeá al menos una columna del Excel.', 'error');
    return;
  }

  const sheetName = document.getElementById('sheetName').value.trim() || 'ESP';

  try {
    setStatus('Generando documento...');
    const dataRows = readExcelData(excelWorkbook, sheetName, rowNumbers, mapping);
    if (!dataRows.length) {
      setStatus('No se obtuvieron datos para esas filas.', 'error');
      return;
    }

    const doc = buildDocument(dataRows, templateInfo, mapping);
    const d = typeof docx !== 'undefined' ? docx : window.docx;
    const blob = await d.Packer.toBlob(doc);
    const filename = `Tabla_filas_${rowNumbers.join('_')}.docx`;
    saveAs(blob, filename);
    setStatus(`Listo: ${filename} (${dataRows.length} filas)`, 'success');
  } catch (err) {
    setStatus(`Error: ${err.message}`, 'error');
    console.error(err);
  }
});
