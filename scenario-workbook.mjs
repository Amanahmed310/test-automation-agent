import { posix } from 'node:path';

const XML_ENTRY_LIMIT = 10 * 1024 * 1024;
const XML_TOTAL_LIMIT = 20 * 1024 * 1024;
const OUTPUT_LIMIT = 80_000;
const REDACTED = '[REDACTED]';
const sensitive = (value) => /username|password|token|secret|email|credential/i.test(
  String(value).replace(/[\s_.-]/g, ''),
);
const normalized = (value) => String(value).toLowerCase().replace(/[^a-z0-9]/g, '');
// Migration workbooks commonly keep the old and replacement case IDs in
// separate columns. Select by the replacement ID; the old case ID is not the
// identifier Azure DevOps uses for the fetched work item.
const newIdHeader = (value) => ['newtestcase', 'newtestcaseid'].includes(normalized(value));
const oldIdHeader = (value) => ['oldtestcase', 'oldtestcaseid'].includes(normalized(value));
const idHeader = (value) => ['testcaseid', 'testid', 'id'].includes(normalized(value))
  || newIdHeader(value) || oldIdHeader(value);
const localName = (name) => name.split(':').at(-1);
const children = (node, name) => node.children.filter((child) => child.name === name);
const child = (node, name) => children(node, name)[0];
const attribute = (node, name) => node.attributes[name]
  ?? Object.entries(node.attributes).find(([key]) => localName(key) === name)?.[1];

function decodeXml(text) {
  return text.replace(/&([^;\s<&]+);/g, (_, entity) => {
    const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
    if (Object.hasOwn(named, entity)) return named[entity];
    if (!/^#(?:[0-9]+|x[0-9a-f]+)$/i.test(entity)) {
      throw new Error('Workbook XML contains an unsupported entity.');
    }
    const code = /^#x/i.test(entity) ? parseInt(entity.slice(2), 16) : Number(entity.slice(1));
    if (!Number.isInteger(code) || code < 1 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) {
      throw new Error('Workbook XML contains an invalid character entity.');
    }
    return String.fromCodePoint(code);
  });
}

// A bounded XML tokenizer, not a cell-body regex: a self-closing element is
// complete immediately and can never consume a following cell's value/type.
function parseXml(buffer, part) {
  const xml = buffer.toString('utf8').replace(/^\uFEFF/, '');
  const root = { name: '', attributes: {}, children: [], text: '' };
  const stack = [root];
  const tokens = /<!--[^]*?-->|<!\[CDATA\[[^]*?\]\]>|<\?[^]*?\?>|<\/?[A-Za-z_][\w.:-]*(?:[^<>"']|"[^"]*"|'[^']*')*>|[^<]+/gy;
  let offset = 0;
  let count = 0;
  while (offset < xml.length) {
    tokens.lastIndex = offset;
    const match = tokens.exec(xml);
    if (!match) throw new Error(`Unsupported or malformed workbook XML in ${part} (DTD is not supported).`);
    offset = tokens.lastIndex;
    const token = match[0];
    const parent = stack.at(-1);
    if (token.startsWith('<!--') || token.startsWith('<?')) continue;
    if (token.startsWith('<![CDATA[')) {
      parent.text += token.slice(9, -3);
    } else if (token.startsWith('</')) {
      if (stack.length === 1 || token.slice(2, -1).trim() !== parent.qualifiedName) {
        throw new Error(`Mismatched XML element in ${part}.`);
      }
      stack.pop();
    } else if (token.startsWith('<')) {
      const opening = /^<([\w.:-]+)/.exec(token);
      const selfClosing = /\/\s*>$/.test(token);
      const attrs = token.slice(opening[0].length, selfClosing ? token.lastIndexOf('/') : -1);
      const attributes = Object.create(null);
      const pattern = /\s+([\w.:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/gy;
      let cursor = 0;
      while (cursor < attrs.length && attrs.slice(cursor).trim()) {
        pattern.lastIndex = cursor;
        const attr = pattern.exec(attrs);
        if (!attr || Object.hasOwn(attributes, attr[1])) throw new Error(`Invalid XML attributes in ${part}.`);
        attributes[attr[1]] = decodeXml(attr[2] ?? attr[3]);
        cursor = pattern.lastIndex;
      }
      const node = { qualifiedName: opening[1], name: localName(opening[1]), attributes, children: [], text: '' };
      parent.children.push(node);
      count += 1;
      if (count > 500_000 || stack.length > 64) throw new Error(`Workbook XML structure exceeds resource bounds in ${part}.`);
      if (!selfClosing) stack.push(node);
    } else {
      parent.text += decodeXml(token);
    }
  }
  if (stack.length !== 1 || root.children.length !== 1 || root.text.trim()) {
    throw new Error(`Incomplete or malformed workbook XML in ${part}.`);
  }
  return root.children[0];
}

function richText(node) {
  if (!node) return '';
  // Phonetic annotations (rPh) are not part of the actual string.
  return node.children.map((item) => item.name === 't' ? item.text
    : item.name === 'r' ? children(item, 't').map((text) => text.text).join('') : '').join('');
}

function location(ref) {
  const match = /^([A-Z]{1,3})([1-9]\d*)$/.exec(ref ?? '');
  if (!match) throw new Error('Workbook contains a missing or invalid cell reference.');
  const columnIndex = [...match[1]].reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
  const row = Number(match[2]);
  if (columnIndex > 16384 || row > 1048576) throw new Error('Workbook cell reference exceeds XLSX bounds.');
  return { ref, column: match[1], columnIndex, row };
}

function parseCell(node, sharedStrings) {
  const cell = location(attribute(node, 'r'));
  const type = attribute(node, 't') ?? 'n';
  const cached = child(node, 'v');
  const formula = child(node, 'f');
  const inline = child(node, 'is');
  const missingCache = Boolean(formula && (!cached || !cached.text.trim()));
  const value = cached?.text ?? '';
  cell.type = { s: 'sharedString', inlineStr: 'inlineString', n: 'number', b: 'boolean', e: 'error', str: 'string', d: 'date' }[type] ?? type;
  cell.value = null;
  if (missingCache) {
    cell.formula = { status: 'missing-cache' };
  } else if (type === 's' && cached && value.trim()) {
    if (!/^\d+$/.test(value.trim()) || !Number.isSafeInteger(Number(value)) || Number(value) >= sharedStrings.length) {
      throw new Error(`Invalid shared-string reference at ${cell.ref}; an index is not scenario data.`);
    }
    cell.value = sharedStrings[Number(value)];
  } else if (type === 'inlineStr') {
    cell.value = inline ? richText(inline) : null;
  } else if (cached) {
    // Keep numeric lexical precision, dates, boolean encodings and whitespace.
    cell.value = value;
  }
  if (formula && !missingCache) cell.formula = { status: 'cached' };
  // Never expose formula source: it may embed credentials or external links.
  return cell;
}

function parseSheet(root, sharedStrings) {
  const rows = [];
  const cellsByRef = new Map();
  const seenRows = new Set();
  for (const node of children(child(root, 'sheetData') ?? { children: [] }, 'row')) {
    const cells = children(node, 'c').map((item) => parseCell(item, sharedStrings));
    const row = Number(attribute(node, 'r') ?? cells[0]?.row);
    if (!Number.isInteger(row) || row < 1 || row > 1048576 || seenRows.has(row)) {
      throw new Error('Workbook contains a missing, duplicate or invalid row number.');
    }
    seenRows.add(row);
    for (const cell of cells) {
      if (cell.row !== row || cellsByRef.has(cell.ref)) throw new Error('Workbook contains inconsistent or duplicate cell references.');
      cellsByRef.set(cell.ref, cell);
    }
    rows.push({ row, cells: cells.sort((a, b) => a.columnIndex - b.columnIndex) });
  }
  const merges = children(child(root, 'mergeCells') ?? { children: [] }, 'mergeCell').map((node) => {
    const range = attribute(node, 'ref');
    const refs = range?.split(':') ?? [];
    if (refs.length < 1 || refs.length > 2) throw new Error('Invalid workbook merged range.');
    const start = location(refs[0]);
    const end = location(refs[1] ?? refs[0]);
    if (end.row < start.row || end.columnIndex < start.columnIndex) throw new Error('Reversed workbook merged range.');
    return { range, start, end };
  });
  return { rows: rows.sort((a, b) => a.row - b.row), cellsByRef, merges };
}

const hasCellValue = (cell) => String(cell?.value ?? '').trim() !== '' || Boolean(cell?.formula);
const isHeading = (cell) => /^\s*test\s*case\s*(?:id\s*)?(?:#\s*)?\d/i.test(String(cell?.value ?? ''));
function headingId(cell) {
  // The entire numeric identifier must end here: 91825 is not 918250 or 91825A.
  return /^\s*test\s*case\s*(?:id\s*)?(?:#\s*)?(\d+)(?=$|\s|[:\-–—.)])/i.exec(String(cell?.value ?? ''))?.[1] ?? null;
}

function selectRows(sheet, target) {
  const firstHeader = sheet.rows.find((row) => !row.cells.some(isHeading)
    && row.cells.some((cell) => idHeader(cell.value)));
  const firstHeading = sheet.rows.find((row) => row.cells.some(isHeading));
  const selected = [];
  const usedHeaders = new Set();
  // A section may itself contain an "ID" field; that is not an ID table.
  // Conversely, an established table must not change its key column because
  // one of its scenario values happens to be "ID" or "TestCase ID".
  if (firstHeader && (!firstHeading || firstHeader.row < firstHeading.row)) {
    let header;
    let idColumn;
    let migrationIdColumn = false;
    let activeSection = false;
    let boundaryRow = 0;
    const select = (row, selection) => {
      usedHeaders.add(header);
      selected.push({ ...row, selection });
    };
    for (const row of sheet.rows) {
      let ids = row.cells.filter((cell) => idHeader(cell.value)
        && (idColumn === undefined || cell.columnIndex === idColumn));
      // Resolve migration headers before ambiguity checks, regardless of their
      // column order. Once chosen, the key column remains fixed for the table.
      if (idColumn === undefined && ids.some((cell) => newIdHeader(cell.value))) {
        ids = ids.filter((cell) => newIdHeader(cell.value));
      }
      if (ids.length) {
        if (ids.length > 1) throw new Error('Ambiguous workbook: multiple ID columns on one header row.');
        header = row;
        idColumn = ids[0].columnIndex;
        migrationIdColumn = newIdHeader(ids[0].value) || oldIdHeader(ids[0].value);
        activeSection = false;
        boundaryRow = row.row;
        continue;
      }
      if (!header) continue;
      let idCell = row.cells.find((cell) => cell.columnIndex === idColumn);
      if (isHeading(idCell)) {
        activeSection = headingId(idCell) === target;
        boundaryRow = row.row;
        if (activeSection) select(row, { mode: 'section' });
        continue;
      }
      // An explicit numeric key supersedes a section. A heading-looking value
      // in another field on that same data row cannot override the key.
      if (/^\d+$/.test(String(idCell?.value ?? '').trim())) {
        activeSection = false;
        boundaryRow = row.row;
        if (String(idCell.value).trim() === target) select(row, { mode: 'id-column', idRef: idCell.ref });
        continue;
      }
      // An unknown formula key or a non-exact numeric-looking key may mark
      // another case; do not inherit a section or merged ID through it.
      if ((idCell?.formula && idCell.value == null) || /^[+-]?\d/.test(String(idCell?.value ?? '').trim())) {
        activeSection = false;
        boundaryRow = row.row;
        continue;
      }
      let inheritedRange;
      if (!hasCellValue(idCell)) {
        const merges = sheet.merges.filter((merge) => merge.start.columnIndex === idColumn
          && merge.start.row > header.row && merge.start.row >= boundaryRow
          && merge.start.row < row.row && merge.end.row >= row.row);
        if (merges.length > 1) throw new Error('Ambiguous overlapping merged ID ranges in workbook.');
        if (merges.length) {
          idCell = sheet.cellsByRef.get(merges[0].start.ref);
          inheritedRange = merges[0].range;
        }
      }
      const titles = row.cells.filter(isHeading);
      if (titles.length && (!migrationIdColumn || !hasCellValue(idCell))) {
        // Outside the key column, only a standalone title with no competing
        // (including merged) ID is unambiguously a section. Ambiguous titles
        // close the previous section but cannot start a new one. Migration
        // tables never fall back to an old/other column to identify a case.
        activeSection = !migrationIdColumn && titles.length === 1 && row.cells.filter(hasCellValue).length === 1
          && !hasCellValue(idCell) && headingId(titles[0]) === target;
        boundaryRow = row.row;
        if (activeSection) select(row, { mode: 'section' });
        continue;
      }
      if (activeSection) {
        select(row, { mode: 'section' });
      } else if (String(idCell?.value ?? '').trim() === target) {
        select(row, { mode: 'id-column', idRef: idCell.ref, ...(inheritedRange ? { inheritedRange } : {}) });
      }
    }
  } else {
    let active = false;
    for (const row of sheet.rows) {
      const titles = row.cells.filter(isHeading);
      if (titles.length) active = titles.length === 1 && headingId(titles[0]) === target;
      if (active) selected.push({ ...row, selection: { mode: 'section' } });
    }
  }
  return { selected, headers: [...usedHeaders] };
}

function redactor(sheet) {
  // Scan the WHOLE sheet, including headers before a selected section. Apply
  // conservative column and key/value-row masking; never fill empty cells.
  const sensitiveColumns = new Map();
  const sensitiveRows = new Set();
  const markedCells = new Set();
  for (const row of sheet.rows) {
    for (const cell of row.cells) {
      if (!sensitive(cell.value ?? '')) continue;
      markedCells.add(cell.ref);
      sensitiveRows.add(row.row);
      sensitiveColumns.set(cell.columnIndex, Math.min(row.row, sensitiveColumns.get(cell.columnIndex) ?? Infinity));
    }
  }
  const directlyMasked = (cell) => markedCells.has(cell.ref) || sensitiveRows.has(cell.row)
    || cell.row >= (sensitiveColumns.get(cell.columnIndex) ?? Infinity);
  // Propagate redaction through actual merges, including their anchors. This
  // also handles a sensitive header merged horizontally across several columns.
  const maskedMerges = new Set();
  let changed = true;
  while (changed) {
    changed = false;
    for (const merge of sheet.merges) {
      if (maskedMerges.has(merge)) continue;
      const touchesColumn = [...sensitiveColumns].some(([column, startRow]) => column >= merge.start.columnIndex
        && column <= merge.end.columnIndex && startRow <= merge.end.row);
      const touchesRow = [...sensitiveRows].some((row) => row >= merge.start.row && row <= merge.end.row);
      if (!touchesColumn && !touchesRow) continue;
      maskedMerges.add(merge);
      for (let column = merge.start.columnIndex; column <= merge.end.columnIndex; column += 1) {
        sensitiveColumns.set(column, Math.min(merge.start.row, sensitiveColumns.get(column) ?? Infinity));
      }
      changed = true;
    }
  }
  return (cell) => {
    if (!cell) return null;
    return directlyMasked(cell) ? { ...cell, value: REDACTED, redacted: true } : { ...cell };
  };
}

function relationshipPath(target) {
  if (!target || /[\\?#]/.test(target) || /^[a-z][a-z0-9+.-]*:/i.test(target)) {
    throw new Error('Invalid or external workbook relationship target.');
  }
  let decoded;
  try { decoded = decodeURIComponent(target); } catch { throw new Error('Invalid workbook relationship path encoding.'); }
  if (/[\\?#\u0000]/.test(decoded)) throw new Error('Invalid workbook relationship path.');
  const path = posix.normalize(decoded.startsWith('/') ? decoded.slice(1)
    : decoded.startsWith('xl/') ? decoded : posix.join('xl', decoded));
  if (!path.startsWith('xl/') || !path.endsWith('.xml')) throw new Error('Workbook relationship escapes the XML resource directory.');
  return path;
}

/** Extract only one test case from an already-expanded XLSX ZIP entry map.
 * Pure: no filesystem/network access, mutation, formula evaluation or dependencies.
 * Returns a redacted model within the JSON output budget; numeric values remain
 * lossless lexical strings.
 */
export function extractWorkbookModel(entries, testCaseId) {
  if (!(entries instanceof Map)) throw new TypeError('Workbook entries must be a Map of ZIP paths to Buffers.');
  const target = String(testCaseId ?? '').trim();
  if (!/^\d+$/.test(target) || target.length > 64) throw new Error('A numeric, exact Test Case ID is required.');
  let total = 0;
  for (const [part, buffer] of entries) {
    if (typeof part !== 'string' || !Buffer.isBuffer(buffer)) throw new TypeError('Workbook entries must map XML paths to Buffers.');
    total += buffer.length;
    if (buffer.length > XML_ENTRY_LIMIT || total > XML_TOTAL_LIMIT) {
      throw new Error('The XLSX workbook contains oversized XML data (10 MiB per entry / 20 MiB total).');
    }
  }
  const read = (part) => {
    if (!entries.has(part)) throw new Error(`Required workbook XML part is missing: ${part}.`);
    return parseXml(entries.get(part), part);
  };
  const workbook = read('xl/workbook.xml');
  const rels = read('xl/_rels/workbook.xml.rels');
  const relationships = new Map();
  for (const rel of children(rels, 'Relationship')) {
    const id = attribute(rel, 'Id');
    if (!id || relationships.has(id)) throw new Error('Missing or duplicate workbook relationship ID.');
    relationships.set(id, rel);
  }
  const sharedRel = [...relationships.values()].find((rel) => /\/sharedStrings$/.test(attribute(rel, 'Type') ?? ''));
  if (sharedRel && attribute(sharedRel, 'TargetMode') === 'External') throw new Error('External shared strings are not supported.');
  const sharedPath = sharedRel ? relationshipPath(attribute(sharedRel, 'Target')) : 'xl/sharedStrings.xml';
  const sharedStrings = entries.has(sharedPath) ? children(read(sharedPath), 'si').map(richText) : [];
  const result = {
    testCaseId: target,
    status: 'no-match',
    message: `No exact Test Case ID ${target} row or section was found. No other case rows are included.`,
    formatNotes: [
      'Cells remain positional evidence, not inferred API paths. Keep field labels distinct from values; do not concatenate adjacent or merged columns into a field path.',
      'Original references and columns are retained. Missing cells are not filled; merged ranges identify their anchors only. ID inheritance requires an actual merge.',
      'Shared strings are decoded: a numeric shared-string index is not data. Numeric values retain their exact lexical text, not an inferred format or date.',
      'Formula source is omitted. Cached results are evidence, not recalculated values; missing-cache means no usable formula value was supplied.',
      'Sensitive headers/columns and key/value rows are conservatively redacted, including merge anchors. Headers outside the selected section are checked too.',
    ],
    sheets: [],
  };
  for (const sheetNode of children(child(workbook, 'sheets') ?? { children: [] }, 'sheet')) {
    const rel = relationships.get(attribute(sheetNode, 'id'));
    if (!rel) throw new Error('A workbook sheet has no matching relationship.');
    if (attribute(rel, 'TargetMode') === 'External') throw new Error('External workbook sheets are not supported.');
    const type = attribute(rel, 'Type');
    if (type && !/\/worksheet$/.test(type)) continue; // Chart/dialog sheets contain no worksheet cells.
    const part = relationshipPath(attribute(rel, 'Target'));
    const sheet = parseSheet(read(part), sharedStrings);
    const { selected, headers } = selectRows(sheet, target);
    if (!selected.length) continue;
    const redact = redactor(sheet);
    const selectedNumbers = new Set(selected.map((row) => row.row));
    const serializeRow = (row) => ({ ...row, cells: row.cells.map(redact) });
    const merges = sheet.merges.filter((merge) => selected.some((row) => row.row >= merge.start.row && row.row <= merge.end.row))
      .map((merge) => ({
        range: merge.range,
        anchorRef: merge.start.ref,
        anchor: selectedNumbers.has(merge.start.row) ? redact(sheet.cellsByRef.get(merge.start.ref)) : null,
        ...(selectedNumbers.has(merge.start.row) ? {} : { anchorOmitted: 'outside-selected-case' }),
      }));
    result.sheets.push({ name: attribute(sheetNode, 'name') ?? '', part, headers: headers.map(serializeRow), rows: selected.map(serializeRow), merges });
  }
  if (result.sheets.length) {
    result.status = 'matched';
    result.message = `Only exact Test Case ID ${target} rows or sections are included.`;
  }
  const text = JSON.stringify(result, null, 2);
  if (text.length > OUTPUT_LIMIT) {
    throw new Error(`Extracted Test Case ID ${target} exceeds the ${OUTPUT_LIMIT.toLocaleString('en-US')}-character budget (${text.length} characters). Narrow the workbook; no partial section was returned.`);
  }
  return result;
}

/** Compatibility API: retain the same readable JSON and complete-section budget. */
export function extractWorkbookEntries(entries, testCaseId) {
  return JSON.stringify(extractWorkbookModel(entries, testCaseId), null, 2);
}
