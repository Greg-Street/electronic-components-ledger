import * as XLSX from 'xlsx';
import { createIcons, icons } from 'lucide';
import seedData from './seed.json';
import './styles.css';

const STORAGE_KEY = 'partshelf-lite-v1';
const PURCHASE_ALIASES = ['商品编号', 'Supplier Part', 'LCSC', '立创编号', '料号', '物料编码'];
const MODEL_ALIASES = ['商品型号', '厂家型号', 'Manufacturer Part', '型号', 'Manufacturer part'];
const PURCHASE_QTY_ALIASES = ['订购数量（修改后）', '订购数量', 'Quantity', '数量', 'Qty'];
const BOM_QTY_ALIASES = ['Quantity', '数量', 'Qty', '用量'];
const TYPE_ALIASES = ['商品类型', '类型', 'Category'];
const NAME_ALIASES = ['商品名称', '名称', 'Comment', 'Value'];
const PACKAGE_ALIASES = ['封装格式', 'Footprint', '封装'];
const BRAND_ALIASES = ['品牌', 'Manufacturer', '厂商'];
const PASSIVE_CATEGORIES = new Set(['电阻', '电容', '电感', '磁珠']);

const ui = {
  view: 'overview',
  query: '',
  status: 'all',
  category: 'all',
  sortKey: 'remaining',
  sortDirection: 'asc',
  expandedCodes: new Set(),
  collapsedGroups: new Set(),
  selectedCodes: new Set(),
  toast: null,
};

let modalState = null;

const copy = (value) => JSON.parse(JSON.stringify(value));

function isChipRecord(record) {
  const text = [record.type, record.name, record.model, record.value, record.category, record.description, record.primaryCategory, record.secondaryCategory].filter(Boolean).join(' ');
  return /芯片|集成电路|运算放大器|放大器|隔离器|ADC|DAC|TPS\d|OPA\d|AD\d|ISO\d|MAX\d|LT\d|ADS\d/i.test(text);
}

function createInitialState() {
  return {
    purchases: seedData.purchases.map((item) => ({ ...item, qty: Number(item.qty) || 0, source: '采购对账单' })),
    usage: seedData.usage.map((item) => ({ ...item, qty: Number(item.qty) || 0, source: item.projects?.join('、') || 'BOM' })),
    plans: [],
    planSources: [],
    adjustments: [],
    imports: [],
  };
}

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) return createInitialState();
    const parsed = JSON.parse(saved);
    return {
      ...createInitialState(),
      ...parsed,
      purchases: Array.isArray(parsed.purchases) ? parsed.purchases : createInitialState().purchases,
      usage: Array.isArray(parsed.usage) ? parsed.usage : createInitialState().usage,
      plans: Array.isArray(parsed.plans) ? parsed.plans : [],
      planSources: Array.isArray(parsed.planSources) ? parsed.planSources : [],
      adjustments: Array.isArray(parsed.adjustments) ? parsed.adjustments : [],
      imports: Array.isArray(parsed.imports) ? parsed.imports : createInitialState().imports,
    };
  } catch {
    return createInitialState();
  }
}

let state = loadState();

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function numberValue(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const cleaned = String(value ?? '').replace(/[￥¥,]/g, '').trim();
  const parsed = Number(cleaned);
  if (Number.isFinite(parsed)) return parsed;
  const match = cleaned.match(/[-+]?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function normalize(value) {
  return String(value ?? '').trim().toLowerCase();
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function slugFromFilename(filename) {
  return filename.replace(/\.[^.]+$/, '').replace(/_?\d{4}-\d{2}-\d{2}$/, '');
}

function findHeader(rows, requiredAliases) {
  const required = requiredAliases.map(normalize);
  return rows.findIndex((row) => {
    const cells = row.map(normalize);
    return required.every((alias) => cells.some((cell) => cell === alias || cell.includes(alias)));
  });
}

function headerMap(row) {
  const map = new Map();
  row.forEach((cell, index) => map.set(normalize(cell), index));
  return map;
}

function indexByAliases(map, aliases) {
  for (const alias of aliases) {
    const normalizedAlias = normalize(alias);
    const index = map.get(normalizedAlias);
    if (index !== undefined) return index;
    const fuzzy = [...map.entries()].find(([key]) => key.includes(normalizedAlias) || normalizedAlias.includes(key));
    if (fuzzy) return fuzzy[1];
  }
  return -1;
}

function cellByAliases(row, map, aliases) {
  const index = indexByAliases(map, aliases);
  return index >= 0 ? row[index] : '';
}

function fallbackKey({ model, value, footprint, designator }) {
  const seed = [model, value, footprint, designator].find((item) => String(item ?? '').trim());
  return `MAN:${seed || `未命名-${Date.now()}`}`;
}

function parseWorkbook(file) {
  return file.arrayBuffer().then((buffer) => XLSX.read(buffer, { type: 'array', cellDates: true }));
}

function bufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return btoa(binary);
}

function parseBomWorkbook(workbook, sourceName) {
  const rows = [];
  const project = slugFromFilename(sourceName);
  workbook.SheetNames.forEach((sheetName) => {
    const sheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
    const headerIndex = findHeader(sheetRows, ['Quantity', 'Supplier Part']);
    if (headerIndex < 0) return;
    const map = headerMap(sheetRows[headerIndex]);
    sheetRows.slice(headerIndex + 1).forEach((row, rowOffset) => {
      const sourceRow = headerIndex + 1 + rowOffset;
      const model = String(cellByAliases(row, map, ['Manufacturer Part']) || '').trim();
      const value = String(cellByAliases(row, map, ['Value', 'Comment']) || '').trim();
      const footprint = String(cellByAliases(row, map, PACKAGE_ALIASES) || '').trim();
      const designator = String(cellByAliases(row, map, ['Designator']) || '').trim();
      const code = String(cellByAliases(row, map, ['Supplier Part', 'LCSC', '商品编号']) || '').trim();
      const description = String(cellByAliases(row, map, ['Description', '描述', 'BOM Description']) || '').trim();
      const voltageRatingMax = String(cellByAliases(row, map, ['Voltage Rating (Max)', '额定电压', 'Voltage Rating']) || '').trim();
      const voltageSupply = String(cellByAliases(row, map, ['Voltage - Supply', '供电电压', 'Supply Voltage']) || '').trim();
      const temperatureCoefficient = String(cellByAliases(row, map, ['Temperature Coefficient', '温度系数']) || '').trim();
      const powerWatts = String(cellByAliases(row, map, ['Power(Watts)', 'Power (Watts)', '功率']) || '').trim();
      const category = String(cellByAliases(row, map, ['Category', '分类']) || '').trim();
      const primaryCategory = String(cellByAliases(row, map, ['Primary Category', '一级分类']) || '').trim();
      const secondaryCategory = String(cellByAliases(row, map, ['Secondary Category', '二级分类']) || '').trim();
      const jlcpcbPrice = cellByAliases(row, map, ['JLCPCB Price', '嘉立创价格']);
      const lcscPrice = cellByAliases(row, map, ['LCSC Price', '立创价格']);
      const qty = numberValue(cellByAliases(row, map, BOM_QTY_ALIASES));
      if ((!code && !model && !value) || qty <= 0) return;
      rows.push({
        code: code || fallbackKey({ model, value, footprint, designator }),
        qty,
        value,
        footprint,
        designator,
        model,
        manufacturerPart: model,
        brand: String(cellByAliases(row, map, BRAND_ALIASES) || '').trim(),
        supplier: String(cellByAliases(row, map, ['Supplier', '供应商']) || 'LCSC').trim(),
        description,
        voltageRatingMax,
        voltageSupply,
        temperatureCoefficient,
        powerWatts,
        category,
        primaryCategory,
        secondaryCategory,
        jlcpcbPrice,
        lcscPrice,
        project,
        source: sourceName,
        sheetName,
        sourceRow,
      });
    });
  });
  if (!rows.length) throw new Error('没有找到 BOM 的“Quantity / Supplier Part”列。');
  return rows;
}

function parsePurchaseWorkbook(workbook, sourceName) {
  const rows = [];
  workbook.SheetNames.forEach((sheetName) => {
    const sheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
    const headerIndex = findHeader(sheetRows, ['商品编号', '订购数量']);
    if (headerIndex < 0) return;
    const map = headerMap(sheetRows[headerIndex]);
    sheetRows.slice(headerIndex + 1).forEach((row) => {
      const code = String(cellByAliases(row, map, PURCHASE_ALIASES) || '').trim();
      const model = String(cellByAliases(row, map, MODEL_ALIASES) || '').trim();
      if (!code && !model) return;
      const type = String(cellByAliases(row, map, TYPE_ALIASES) || '').trim();
      const name = String(cellByAliases(row, map, NAME_ALIASES) || '').trim();
      const item = {
        code: code || fallbackKey({ model, value: name }),
        brand: String(cellByAliases(row, map, BRAND_ALIASES) || '').trim(),
        type,
        name,
        model,
        manufacturerPart: model,
        package: String(cellByAliases(row, map, PACKAGE_ALIASES) || '').trim(),
        qty: numberValue(cellByAliases(row, map, PURCHASE_QTY_ALIASES)),
        unit: String(cellByAliases(row, map, ['商品单位', 'Unit', '单位']) || '个').trim(),
        price: numberValue(cellByAliases(row, map, ['商品单价', '单价（人民币含税）', '单价', 'Price'])),
        supplier: String(cellByAliases(row, map, ['供应商', 'Supplier']) || 'LCSC').trim(),
        source: sourceName,
      };
      item.type ||= inferCategory(item);
      if (item.qty > 0) rows.push(item);
    });
  });
  if (!rows.length) throw new Error('没有找到采购表的“商品编号 / 订购数量”列。');
  return rows;
}

async function importPurchaseFile(file) {
  const workbook = await parseWorkbook(file);
  const rows = parsePurchaseWorkbook(workbook, file.name);
  state.purchases.push(...rows);
  state.imports.unshift({ id: crypto.randomUUID(), kind: 'purchase', name: file.name, note: `新增 ${rows.length} 条采购记录`, stamp: new Date().toLocaleString('zh-CN', { hour12: false }) });
  saveState();
  return rows.length;
}

async function parseBomRows(file) {
  const workbook = await parseWorkbook(file);
  return parseBomWorkbook(workbook, file.name);
}

async function importBomFile(file) {
  const rows = await parseBomRows(file);
  state.usage.push(...rows);
  const total = rows.reduce((sum, item) => sum + item.qty, 0);
  state.imports.unshift({ id: crypto.randomUUID(), kind: 'bom', name: file.name, note: `新增 ${formatQty(total)} 个用量`, stamp: new Date().toLocaleString('zh-CN', { hour12: false }) });
  saveState();
  return rows.length;
}

async function importPlanFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
  const rows = parseBomWorkbook(workbook, file.name);
  const planId = crypto.randomUUID();
  state.planSources.push({ id: planId, name: file.name, data: bufferToBase64(buffer) });
  state.plans.push(...rows.map((item) => ({ ...item, planId, sourceId: planId })));
  const total = rows.reduce((sum, item) => sum + item.qty, 0);
  state.imports.unshift({ id: planId, kind: 'plan', name: file.name, note: `新板备料需求 ${formatQty(total)} 个`, stamp: new Date().toLocaleString('zh-CN', { hour12: false }) });
  saveState();
  return rows.length;
}

function workbookHasHeader(workbook, requiredAliases) {
  return workbook.SheetNames.some((sheetName) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
    return findHeader(rows, requiredAliases) >= 0;
  });
}

function resetLedger() {
  if (!window.confirm('确定清空当前台账吗？这会删除当前采购记录、已用 BOM、手工调整和备料计划，但不会删除初始空台账。')) return;
  state.purchases = [];
  state.usage = [];
  state.adjustments = [];
  state.plans = [];
  state.planSources = [];
  state.imports = [];
  ui.query = '';
  ui.status = 'all';
  ui.category = 'all';
  ui.selectedCodes.clear();
  ui.expandedCodes.clear();
  ui.collapsedGroups.clear();
  saveState();
  render();
  showToast('当前台账已清空');
}

async function resetAndImportFiles(files) {
  const purchases = [];
  const usage = [];
  const imports = [];

  for (const file of files) {
    const workbook = await parseWorkbook(file);
    if (workbookHasHeader(workbook, ['商品编号', '订购数量'])) {
      const rows = parsePurchaseWorkbook(workbook, file.name);
      purchases.push(...rows);
      imports.push({ id: crypto.randomUUID(), kind: 'purchase', name: file.name, note: `重新导入采购记录 ${rows.length} 条`, stamp: new Date().toLocaleString('zh-CN', { hour12: false }) });
      continue;
    }
    if (workbookHasHeader(workbook, ['Quantity', 'Supplier Part'])) {
      const rows = parseBomWorkbook(workbook, file.name);
      usage.push(...rows);
      const total = rows.reduce((sum, item) => sum + item.qty, 0);
      imports.push({ id: crypto.randomUUID(), kind: 'bom', name: file.name, note: `重新导入已用 BOM ${rows.length} 条，共 ${formatQty(total)} 个用量`, stamp: new Date().toLocaleString('zh-CN', { hour12: false }) });
      continue;
    }
    throw new Error(`无法识别“${file.name}”，请确认是采购表或 BOM 表。`);
  }

  if (!purchases.length && !usage.length) throw new Error('没有识别到有效的采购记录或 BOM 用量。');
  const purchaseTotal = purchases.reduce((sum, item) => sum + numberValue(item.qty), 0);
  const usageTotal = usage.reduce((sum, item) => sum + numberValue(item.qty), 0);
  if (!window.confirm(`即将用 ${files.length} 个文件替换当前台账：采购 ${formatQty(purchaseTotal)} 个，已用 BOM ${formatQty(usageTotal)} 个。继续吗？`)) return false;

  state.purchases = purchases;
  state.usage = usage;
  state.adjustments = [];
  state.plans = [];
  state.planSources = [];
  state.imports = imports;
  ui.query = '';
  ui.status = 'all';
  ui.category = 'all';
  ui.selectedCodes.clear();
  ui.expandedCodes.clear();
  ui.collapsedGroups.clear();
  saveState();
  render();
  return { purchaseRows: purchases.length, usageRows: usage.length, purchaseTotal, usageTotal };
}

async function handleResetImport(files) {
  try {
    const result = await resetAndImportFiles(files);
    if (result === false) return;
    showToast(`已重置并导入：采购 ${formatQty(result.purchaseTotal)} 个，已用 ${formatQty(result.usageTotal)} 个`);
  } catch (error) {
    showToast(error.message || '重置导入失败，当前台账未改变。', 'error');
  }
}

function formatQty(value) {
  const number = Number(value) || 0;
  return Number.isInteger(number) ? number.toLocaleString('zh-CN') : number.toLocaleString('zh-CN', { maximumFractionDigits: 2 });
}

function firstMatch(text, pattern) {
  const match = String(text || '').match(pattern);
  return match ? match[0].replace(/\s+/g, '') : '';
}

function firstLabeledValue(text, labels) {
  const match = String(text || '').match(new RegExp(`(?:${labels})\\s*[:：]\\s*([^;；,，]+)`, 'i'));
  return match ? match[1].trim() : '';
}

function extractSpecs(item) {
  const text = [
    item.name,
    item.value,
    item.model,
    item.description,
    item.voltageRatingMax,
    item.voltageSupply,
    item.temperatureCoefficient,
    item.powerWatts,
    item.category,
    item.primaryCategory,
    item.secondaryCategory,
  ].filter(Boolean).join(' ');
  return {
    value: String(item.value || firstMatch(text, /\d+(?:\.\d+)?\s*(?:[kKmM]?Ω|uF|μF|nF|pF|uH|mH|GHz|kHz)/i) || '').trim(),
    tolerance: firstMatch(text, /±\s*\d+(?:\.\d+)?\s*%/),
    power: firstMatch(text, /\d+(?:\.\d+)?\s*(?:mW|W)/i),
    voltage: firstMatch(text, /\d+(?:\.\d+)?\s*V(?:\b|$)/i),
    current: firstMatch(text, /\d+(?:\.\d+)?\s*A(?:\b|$)/i),
    frequency: firstMatch(text, /\d+(?:\.\d+)?\s*(?:GHz|MHz|kHz)/i),
    temperature: firstMatch(text, /[-−]\s*\d+\s*℃?\s*[~～至]\s*[+＋]?\s*\d+\s*℃?/),
    voltageRatingMax: String(item.voltageRatingMax || firstLabeledValue(item.description, '额定电压|Voltage\\s+Rating(?:\\s*\\(Max\\))?') || '').trim(),
    voltageSupply: String(item.voltageSupply || firstLabeledValue(item.description, '供电电压|Voltage\\s*-\\s*Supply|Supply\\s+Voltage') || '').trim(),
    temperatureCoefficient: String(item.temperatureCoefficient || firstLabeledValue(item.description, '温度系数|Temperature\\s+Coefficient') || firstMatch(text, /(?:X[567]R|Y[4575]V|NP0|C0G)/i) || '').trim(),
    powerWatts: String(item.powerWatts || firstLabeledValue(item.description, '功率|Power(?:\\s*\\(Watts\\))?') || '').trim(),
  };
}

function mergeDetails(details, item) {
  const parsed = extractSpecs(item);
  details.value ||= parsed.value;
  details.footprint ||= item.footprint || item.package || '';
  details.tolerance ||= parsed.tolerance;
  details.power ||= parsed.power || (parsed.powerWatts ? `${parsed.powerWatts} W` : '');
  details.powerWatts ||= parsed.powerWatts;
  details.voltage ||= parsed.voltage;
  details.voltageRatingMax ||= parsed.voltageRatingMax;
  details.voltageSupply ||= parsed.voltageSupply;
  details.current ||= parsed.current;
  details.frequency ||= parsed.frequency;
  details.temperature ||= parsed.temperature;
  details.temperatureCoefficient ||= parsed.temperatureCoefficient;
  if (item.description) details.description = item.description;
  else details.description ||= item.name || '';
  details.category ||= item.category || '';
  details.primaryCategory ||= item.primaryCategory || '';
  details.secondaryCategory ||= item.secondaryCategory || '';
  details.manufacturerPart ||= item.manufacturerPart || item.model || '';
  details.unit ||= item.unit || '个';
  details.unitPrice ||= numberValue(item.price) || numberValue(item.lcscPrice) || numberValue(item.jlcpcbPrice);
  details.jlcpcbPrice ||= numberValue(item.jlcpcbPrice);
  details.lcscPrice ||= numberValue(item.lcscPrice);
  details.supplier ||= item.supplier || 'LCSC';
}

function inferCategory(item) {
  const text = [item.type, item.name, item.model, item.value, item.designator, item.package, item.footprint, item.category, item.description, item.primaryCategory, item.secondaryCategory].filter(Boolean).join(' ');
  if (isChipRecord(item) || /(^|[,\s])U\d/i.test(item.designator || '')) return '芯片';
  if (/电阻|排阻|Ω|resistor|^R/i.test(text)) return '电阻';
  if (/电容|uF|nF|pF|电解|capacitor|^C/i.test(text)) return '电容';
  if (/电感|磁珠|uH|inductor|^L/i.test(text)) return '电感';
  if (/二极管|肖特基|稳压|LED|发光|^D|^LED|BAS\d|BZT\d|(?:^|\s)(?:1N|SS|MBR)\d/i.test(text)) return /LED|发光/i.test(text) ? 'LED' : '二极管';
  if (/电源模块|隔离电源|PWRM|^WR[A-Z]/i.test(text)) return '模块';
  if (/按键|开关|switch|^SW|SS-\d/i.test(text)) return '开关';
  if (/连接器|接线端子|针座|排针|排母|插座|插头|端子|直插排针|SMA|^CN|^RF|DB\d|HX-XH|HDR-TH/i.test(text)) return '连接器';
  if (/测试点|test(?:[\s_-])?point/i.test(text)) return '测试点';
  if (/电源|模块|隔离|变换/i.test(text)) return '模块';
  return item.type || '其他';
}

function normalizePartSpec(value) {
  return normalize(value)
    .replace(/[\u03bc\u00b5]/g, 'u')
    .replace(/[\u2212\u2013\u2014]/g, '-')
    .replace(/\s+/g, '');
}

function normalizeFootprint(value) {
  const footprint = normalizePartSpec(value);
  return footprint.replace(/^[rcl](?=\d{4}(?:$|[-_]))/, '');
}

function materialKey(category, item) {
  const details = item.details || item;
  const value = normalizePartSpec(details.value || item.value);
  const footprint = normalizeFootprint(details.footprint || item.footprint || item.package);
  if (PASSIVE_CATEGORIES.has(category) && value) return ['spec', category, value, footprint].join('|');
  const model = normalizePartSpec(details.manufacturerPart || item.manufacturerPart || item.model);
  if (model) return ['model', category, model, footprint].join('|');
  return ['code', item.code].join('|');
}

function getInventory() {
  const map = new Map();
  const ensure = (code, item) => {
    if (!map.has(code)) map.set(code, {
      code,
      name: item.name || item.value || item.model || '未命名物料',
      model: item.model || '',
      type: item.type || '',
      brand: item.brand || '',
      package: item.package || item.footprint || '',
      purchased: 0,
      used: 0,
      adjusted: 0,
      projects: new Set(),
      designators: new Set(),
      details: {
        value: '',
        footprint: '',
        tolerance: '',
        power: '',
        powerWatts: '',
        voltage: '',
        voltageRatingMax: '',
        voltageSupply: '',
        current: '',
        frequency: '',
        temperature: '',
        temperatureCoefficient: '',
        description: '',
        category: '',
        primaryCategory: '',
        secondaryCategory: '',
        manufacturerPart: '',
        unit: '个',
        unitPrice: 0,
        jlcpcbPrice: 0,
        lcscPrice: 0,
        supplier: 'LCSC',
      },
    });
    const result = map.get(code);
    result.name = result.name === '未命名物料' ? (item.name || item.value || item.model || result.name) : result.name;
    result.model ||= item.model || '';
    result.type ||= item.type || '';
    result.brand ||= item.brand || '';
    result.package ||= item.package || item.footprint || '';
    mergeDetails(result.details, item);
    return result;
  };

  state.purchases.forEach((item) => {
    const row = ensure(item.code, item);
    row.purchased += numberValue(item.qty);
  });
  state.usage.forEach((item) => {
    const row = ensure(item.code, item);
    row.used += numberValue(item.qty);
    if (item.project) row.projects.add(item.project);
    (item.projects || []).forEach((project) => row.projects.add(project));
    if (item.designator) row.designators.add(item.designator);
  });
  state.adjustments.forEach((item) => {
    const row = ensure(item.code, item);
    row.adjusted += numberValue(item.qty);
  });

  const rows = [...map.values()].map((item) => {
    const remaining = item.purchased - item.used + item.adjusted;
    const category = inferCategory({ ...item, ...item.details });
    const specKey = materialKey(category, item);
    const specLabel = [item.details.value, item.details.footprint, item.details.tolerance, item.details.power].filter(Boolean).join(' · ') || item.model || '未标注规格';
    return { ...item, remaining, groupRemaining: remaining, groupSize: 1, shared: false, category, projects: [...item.projects], designators: [...item.designators], specKey, specLabel, status: 'healthy' };
  });

  const groups = new Map();
  rows.forEach((row) => {
    if (!groups.has(row.specKey)) groups.set(row.specKey, { rows: [], purchased: 0, used: 0, adjusted: 0 });
    const group = groups.get(row.specKey);
    group.rows.push(row);
    group.purchased += row.purchased;
    group.used += row.used;
    group.adjusted += row.adjusted;
  });

  groups.forEach((group) => {
    const groupRemaining = group.purchased - group.used + group.adjusted;
    let demand = group.used;
    const stockRows = group.rows.filter((row) => row.purchased > 0 || row.adjusted !== 0);
    stockRows.forEach((row) => {
      const supply = row.purchased + row.adjusted;
      const usable = Math.max(0, supply);
      const consumed = Math.min(usable, demand);
      row.remaining = supply < 0 ? supply : supply - consumed;
      demand -= consumed;
    });

    group.rows.filter((row) => !stockRows.includes(row)).forEach((row) => { row.remaining = 0; });
    let uncovered = Math.max(0, demand);
    group.rows.filter((row) => row.used > 0).forEach((row) => {
      if (uncovered <= 0) return;
      const shortfall = Math.min(row.used, uncovered);
      row.remaining = row.remaining >= 0 ? -shortfall : row.remaining;
      uncovered -= shortfall;
    });

    group.rows.forEach((row) => {
      row.groupRemaining = groupRemaining;
      row.groupSize = group.rows.length;
      row.shared = group.rows.length > 1;
      const coveredOnly = row.purchased + row.adjusted <= 0 && row.used > 0 && row.remaining === 0 && groupRemaining >= 0;
      row.status = row.remaining < 0
        ? 'shortage'
        : coveredOnly
          ? 'covered'
          : row.remaining === 0
            ? 'depleted'
            : row.remaining <= Math.max(2, row.purchased * 0.1)
              ? 'low'
              : 'healthy';
    });
  });

  return rows;
}

function filteredInventory() {
  const query = normalize(ui.query);
  const rows = getInventory()
    .filter((item) => ui.status === 'all' || item.status === ui.status || (ui.status === 'healthy' && item.status === 'covered'))
    .filter((item) => ui.category === 'all' || item.category === ui.category)
    .filter((item) => !query || [item.code, item.name, item.model, item.brand, item.category, item.specLabel, item.details.value, item.details.footprint, item.details.description, item.details.category, item.details.primaryCategory, item.details.secondaryCategory, item.details.manufacturerPart].some((value) => normalize(value).includes(query)));
  return sortInventoryRows(rows);
}

const INVENTORY_SORT_KEYS = new Set(['model', 'spec', 'footprint', 'category', 'purchased', 'used', 'adjusted', 'remaining', 'status']);

function inventorySortValue(row, key) {
  if (key === 'model') return normalize(row.model || row.name);
  if (key === 'spec') return normalize(row.details.value || row.specLabel);
  if (key === 'footprint') return normalize(row.details.footprint);
  if (key === 'category') return normalize(row.category);
  if (key === 'purchased') return row.purchased;
  if (key === 'used') return row.used;
  if (key === 'adjusted') return row.adjusted;
  if (key === 'status') return ({ shortage: 0, depleted: 1, low: 2, covered: 3, healthy: 4 }[row.status] ?? 9);
  return row.remaining;
}

function compareInventoryValues(left, right) {
  if (typeof left === 'number' && typeof right === 'number') return left - right;
  return String(left).localeCompare(String(right), 'zh-CN', { numeric: true, sensitivity: 'base' });
}

function sortInventoryRows(rows) {
  const direction = ui.sortDirection === 'desc' ? -1 : 1;
  const key = INVENTORY_SORT_KEYS.has(ui.sortKey) ? ui.sortKey : 'remaining';
  return rows.sort((a, b) => {
    const primary = compareInventoryValues(inventorySortValue(a, key), inventorySortValue(b, key));
    if (primary) return primary * direction;
    const modelTie = compareInventoryValues(inventorySortValue(a, 'model'), inventorySortValue(b, 'model'));
    return modelTie || a.code.localeCompare(b.code, 'zh-CN', { numeric: true });
  });
}

function resolveInventorySpecKey(item, category, inventoryByCode) {
  // The same LCSC part can have different footprint labels in purchase exports and BOM exports.
  const code = normalizePartSpec(item.code);
  return inventoryByCode.get(code)?.specKey || materialKey(category, item);
}

function getProcurementSuggestions() {
  const inventory = getInventory();
  const inventoryByCode = new Map(inventory.map((row) => [normalizePartSpec(row.code), row]));
  const inventoryBySpec = new Map();
  inventory.forEach((row) => {
    if (!inventoryBySpec.has(row.specKey)) inventoryBySpec.set(row.specKey, row);
  });

  const groups = new Map();
  state.plans.forEach((item) => {
    const category = inferCategory(item);
    const specKey = resolveInventorySpecKey(item, category, inventoryByCode);
    if (!groups.has(specKey)) {
      groups.set(specKey, {
        specKey,
        category,
        qty: 0,
        value: item.value || '',
        footprint: item.footprint || '',
        model: item.model || '',
        unit: item.unit || '个',
        brands: new Set(),
        codes: new Set(),
        projects: new Set(),
      });
    }
    const group = groups.get(specKey);
    group.qty += numberValue(item.qty);
    if (item.value) group.value ||= item.value;
    if (item.footprint) group.footprint ||= item.footprint;
    if (item.model) group.model ||= item.model;
    if (item.brand) group.brands.add(item.brand);
    if (item.code) group.codes.add(item.code);
    if (item.project) group.projects.add(item.project);
  });

  return [...groups.values()]
    .map((group) => {
      const inventoryRow = inventoryBySpec.get(group.specKey);
      const currentAvailable = Math.max(0, inventoryRow?.groupRemaining || 0);
      const shortage = Math.max(0, group.qty - currentAvailable);
      return {
        ...group,
        currentAvailable,
        shortage,
        specLabel: [group.value, group.footprint].filter(Boolean).join(' · ') || group.model || [...group.codes][0] || '未标注规格',
        modelLabel: group.model || group.value || [...group.codes][0] || '未命名物料',
        brands: [...group.brands],
        codes: [...group.codes],
        projects: [...group.projects],
        status: shortage > 0 ? 'buy' : 'ready',
      };
    })
    .sort((a, b) => b.shortage - a.shortage || b.qty - a.qty || a.specLabel.localeCompare(b.specLabel));
}

function planRowKey(item, index) {
  if (item.sourceId && item.sheetName && Number.isInteger(item.sourceRow)) return `${item.sourceId}|${item.sheetName}|${item.sourceRow}`;
  return `legacy|${index}`;
}

function getPlanExportQuantities() {
  const inventory = getInventory();
  const inventoryByCode = new Map(inventory.map((row) => [normalizePartSpec(row.code), row]));
  const availableBySpec = new Map();
  inventory.forEach((row) => {
    if (!availableBySpec.has(row.specKey)) availableBySpec.set(row.specKey, Math.max(0, row.groupRemaining || 0));
  });

  const quantities = new Map();
  state.plans.forEach((item, index) => {
    const category = inferCategory(item);
    const specKey = resolveInventorySpecKey(item, category, inventoryByCode);
    const available = Math.max(0, availableBySpec.get(specKey) || 0);
    const demand = Math.max(0, numberValue(item.qty));
    const covered = Math.min(demand, available);
    availableBySpec.set(specKey, available - covered);
    quantities.set(planRowKey(item, index), demand - covered);
  });
  return quantities;
}

function shiftFormulaRows(formula, oldRow, newRow) {
  const oldExcelRow = oldRow + 1;
  const newExcelRow = newRow + 1;
  return String(formula).replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g, (match, column, row) => Number(row) === oldExcelRow ? `${column}${newExcelRow}` : match);
}

function cloneWorkbookCell(cell, oldRow, newRow) {
  if (!cell) return null;
  const cloned = copy(cell);
  if (cloned.f) cloned.f = shiftFormulaRows(cloned.f, oldRow, newRow);
  return cloned;
}

function rowHasContent(row) {
  return row.some((value) => String(value ?? '').trim() !== '');
}

function filterBomSheet(sheet, sheetRows, headerIndex, quantityIndex, outputBySourceRow) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  const keepRows = [];
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    if (rowIndex <= headerIndex || !outputBySourceRow.has(rowIndex)) {
      if (rowIndex <= headerIndex || rowHasContent(sheetRows[rowIndex] || [])) keepRows.push(rowIndex);
    } else if (outputBySourceRow.get(rowIndex) > 0) {
      keepRows.push(rowIndex);
    }
  }

  const rowMap = new Map(keepRows.map((oldRow, newRow) => [oldRow, newRow]));
  const output = {};
  Object.keys(sheet).forEach((key) => {
    if (key.startsWith('!') && !['!ref', '!merges', '!rows'].includes(key)) output[key] = copy(sheet[key]);
  });

  keepRows.forEach((oldRow) => {
    const newRow = rowMap.get(oldRow);
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const oldAddress = XLSX.utils.encode_cell({ r: oldRow, c: column });
      const cell = cloneWorkbookCell(sheet[oldAddress], oldRow, newRow);
      if (!cell) continue;
      if (outputBySourceRow.has(oldRow) && column === quantityIndex) {
        cell.v = outputBySourceRow.get(oldRow);
        cell.t = 'n';
        delete cell.w;
      }
      output[XLSX.utils.encode_cell({ r: newRow, c: column })] = cell;
    }
  });

  if (sheet['!cols']) output['!cols'] = copy(sheet['!cols']);
  if (sheet['!rows']) output['!rows'] = keepRows.map((oldRow) => copy(sheet['!rows'][oldRow] || {}));
  output['!merges'] = (sheet['!merges'] || []).flatMap((merge) => {
    const start = rowMap.get(merge.s.r);
    const end = rowMap.get(merge.e.r);
    return start === undefined || end === undefined ? [] : [{ s: { c: merge.s.c, r: start }, e: { c: merge.e.c, r: end } }];
  });
  output['!ref'] = XLSX.utils.encode_range({ s: { c: range.s.c, r: 0 }, e: { c: range.e.c, r: Math.max(0, keepRows.length - 1) } });
  return output;
}

function buildProcurementWorkbook(source, quantities) {
  const workbook = XLSX.read(source.data, { type: 'base64', cellDates: true, cellStyles: true });
  let exportedRows = 0;
  let exportedQuantity = 0;

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const sheetRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const headerIndex = findHeader(sheetRows, ['Quantity', 'Supplier Part']);
    if (headerIndex < 0) return;
    const map = headerMap(sheetRows[headerIndex]);
    const quantityIndex = indexByAliases(map, BOM_QTY_ALIASES);
    if (quantityIndex < 0) return;

    const outputBySourceRow = new Map();
    state.plans.forEach((item, index) => {
      if (item.sourceId !== source.id || item.sheetName !== sheetName || !Number.isInteger(item.sourceRow)) return;
      const outputQty = quantities.get(planRowKey(item, index)) || 0;
      outputBySourceRow.set(item.sourceRow, outputQty);
      if (outputQty > 0) {
        exportedRows += 1;
        exportedQuantity += outputQty;
      }
    });
    if (outputBySourceRow.size) workbook.Sheets[sheetName] = filterBomSheet(sheet, sheetRows, headerIndex, quantityIndex, outputBySourceRow);
  });

  const summarySheet = workbook.Sheets['Field汇总'];
  if (summarySheet) {
    const summaryRows = XLSX.utils.sheet_to_json(summarySheet, { header: 1, defval: '' });
    const totalRow = summaryRows.findIndex((row) => /total\s*quantity|总数量/i.test(String(row[0] || '')));
    if (totalRow >= 0) {
      const totalAddress = XLSX.utils.encode_cell({ r: totalRow, c: 1 });
      summarySheet[totalAddress] ||= {};
      summarySheet[totalAddress].v = exportedQuantity;
      summarySheet[totalAddress].t = 'n';
      delete summarySheet[totalAddress].f;
      delete summarySheet[totalAddress].w;
    }
  }

  return { workbook, exportedRows, exportedQuantity };
}

function sanitizeFilename(name) {
  return String(name || '采购建议').replace(/[\\/:*?"<>|]/g, '_').replace(/\.[^.]+$/, '') || '采购建议';
}

function downloadProcurementWorkbook(workbook, sourceName) {
  XLSX.writeFile(workbook, `${sanitizeFilename(sourceName)}-待采购.xlsx`, { bookType: 'xlsx', compression: true });
}

function buildLegacyProcurementWorkbook(rows) {
  const output = [['No.', 'Quantity', 'Comment', 'Designator', 'Footprint', 'Value', 'Manufacturer Part', 'Manufacturer', 'Supplier Part', 'Supplier']];
  rows.forEach((row, index) => output.push([index + 1, row.shortage, row.value || row.model || '', '', row.footprint || '', row.value || '', row.model || '', row.brands?.[0] || '', row.codes?.[0] || '', 'LCSC']));
  const worksheet = XLSX.utils.aoa_to_sheet(output);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'bom模板');
  return workbook;
}

function exportProcurement() {
  const quantities = getPlanExportQuantities();
  const sources = state.planSources.filter((source) => state.plans.some((item) => item.sourceId === source.id));
  let exportCount = 0;
  let exportedRows = 0;

  sources.forEach((source) => {
    const result = buildProcurementWorkbook(source, quantities);
    if (result.exportedRows <= 0) return;
    downloadProcurementWorkbook(result.workbook, source.name);
    exportCount += 1;
    exportedRows += result.exportedRows;
  });

  if (!sources.length) {
    const rows = getProcurementSuggestions().filter((row) => row.shortage > 0);
    if (rows.length) {
      downloadProcurementWorkbook(buildLegacyProcurementWorkbook(rows), '采购建议');
      exportCount = 1;
      exportedRows = rows.length;
    }
  }

  showToast(exportCount ? `已导出 ${exportedRows} 行待采购物料` : '当前没有需要采购的物料', exportCount ? 'success' : 'error');
}

function statusText(status) {
  return { healthy: '充足', covered: '已覆盖', low: '偏低', depleted: '已用尽', shortage: '缺料' }[status] || status;
}

function statusClass(status) {
  return `status-${status}`;
}

function icon(name, size = 16) {
  return `<i data-lucide="${name}" width="${size}" height="${size}" aria-hidden="true"></i>`;
}

function renderShell() {
  const nav = [
    ['overview', '总览', 'layout-dashboard'],
    ['inventory', '库存明细', 'list-filter'],
    ['planning', '采购建议', 'shopping-cart'],
    ['imports', '导入记录', 'file-stack'],
  ];
  return `
    <div class="app-shell">
      <aside class="sidebar">
        <div class="brand-mark"><span class="brand-dot"></span><span>PARTS<span class="brand-muted">/</span>LEDGER</span></div>
        <div class="sidebar-kicker">电子元件库存台账</div>
        <nav class="primary-nav" aria-label="主导航">
          ${nav.map(([id, label, iconName]) => `<button class="nav-link ${ui.view === id ? 'is-active' : ''}" data-view="${id}">${icon(iconName)}<span>${label}</span></button>`).join('')}
        </nav>
        <div class="sidebar-rule"></div>
        <div class="sidebar-note">
          <span class="eyebrow">当前规则</span>
          <strong>采购记录 + BOM 自动计算</strong>
          <span>允许手工调整与新建元件</span>
        </div>
        <button class="sidebar-reset" id="reset-demo">${icon('rotate-ccw')}<span>恢复初始空台账</span></button>
        <button class="sidebar-reset" data-open-reset-import>${icon('file-input')}<span>重置并导入新台账</span></button>
        <button class="sidebar-reset sidebar-reset-danger" data-clear-ledger>${icon('trash-2')}<span>清空当前台账</span></button>
        <div class="sidebar-footer">本地保存 · 不上传文件</div>
      </aside>
      <section class="workspace">
        <header class="topbar">
          <div class="topbar-context"><span class="live-dot"></span><span>库存状态</span><span class="slash">/</span><strong>${ui.view === 'overview' ? '总览' : ui.view === 'inventory' ? '库存明细' : ui.view === 'planning' ? '采购建议' : '导入记录'}</strong></div>
          <div class="topbar-actions">
            <button class="button button-quiet" id="export-csv" title="导出当前库存 CSV">${icon('download')}<span>导出清单</span></button>
            <button class="button button-primary" data-open-import="purchase">${icon('upload')}<span>导入文件</span></button>
          </div>
        </header>
        <div class="content" id="content"></div>
      </section>
    </div>
    <input class="sr-only" type="file" id="purchase-file" accept=".xls,.xlsx,.csv" />
    <input class="sr-only" type="file" id="bom-file" accept=".xls,.xlsx,.csv" />
    <input class="sr-only" type="file" id="plan-file" accept=".xls,.xlsx,.csv" />
    <input class="sr-only" type="file" id="reset-import-files" accept=".xls,.xlsx,.csv" multiple />
    <div id="toast" class="toast" aria-live="polite"></div>
    <div id="modal-root"></div>
  `;
}

function renderOverview() {
  const rows = getInventory();
  const purchased = rows.reduce((sum, row) => sum + row.purchased, 0);
  const used = rows.reduce((sum, row) => sum + row.used, 0);
  const adjusted = rows.reduce((sum, row) => sum + row.adjusted, 0);
  const remaining = rows.reduce((sum, row) => sum + Math.max(row.remaining, 0), 0);
  const shortageCount = rows.filter((row) => row.status === 'shortage').length;
  const shortageRows = rows.filter((row) => row.status === 'shortage').sort((a, b) => a.remaining - b.remaining).slice(0, 6);
  const categoryMap = new Map();
  rows.forEach((row) => categoryMap.set(row.category, (categoryMap.get(row.category) || 0) + 1));
  const categories = [...categoryMap.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6);
  const coverage = purchased ? Math.min(100, Math.round((used / purchased) * 100)) : 0;

  return `
    <div class="page-heading">
      <div>
        <div class="eyebrow">MATERIAL CONTROL</div>
        <h1>把余量看清楚，<em>再开始下一块板。</em></h1>
        <p>导入采购记录和已用 BOM，按物料规格合并并计算余量；采购建议只列出当前缺口。</p>
      </div>
      <div class="heading-stamp"><span>SYNCED</span><strong>${rows.length}</strong><small>物料种类</small></div>
    </div>
    <section class="rule-strip">
      <div class="rule-strip-main">${icon('scan-line', 18)}<span>自动核算已开启</span><strong>购买数量 − BOM 用量 = 当前余量</strong></div>
      <div class="rule-strip-meta"><span>最近计算</span><b>${new Date().toLocaleTimeString('zh-CN', { hour12: false })}</b></div>
    </section>
    <section class="stat-grid" aria-label="库存概览">
      <article class="stat-card stat-card-accent"><div class="stat-label">当前可用</div><div class="stat-value">${formatQty(remaining)}<small> pcs</small></div><div class="stat-foot">采购 ${formatQty(purchased)} ${adjusted ? `${adjusted > 0 ? '+' : '−'} 手工 ${formatQty(Math.abs(adjusted))}` : ''} − 使用 ${formatQty(used)}</div></article>
      <article class="stat-card"><div class="stat-label">已用 / 采购</div><div class="stat-value">${coverage}<small>%</small></div><div class="progress"><span style="width:${coverage}%"></span></div><div class="stat-foot">BOM 占用率</div></article>
      <article class="stat-card stat-card-warning"><div class="stat-label">需要补料</div><div class="stat-value">${shortageCount}<small> 种</small></div><div class="stat-foot">${shortageCount ? '含采购未覆盖的 BOM 物料' : '暂时没有缺料'}</div></article>
      <article class="stat-card"><div class="stat-label">芯片种类</div><div class="stat-value">${rows.filter((row) => row.category === '芯片').length}<small> 种</small></div><div class="stat-foot">按采购记录中的实际数量计算</div></article>
    </section>
    <section class="overview-grid">
      <div class="section-panel shortage-panel">
        <div class="panel-header"><div><div class="eyebrow">ACTION REQUIRED</div><h2>缺料清单</h2></div><button class="text-button" data-view="inventory" data-status="shortage">查看全部 ${icon('arrow-up-right', 14)}</button></div>
        ${shortageRows.length ? `<div class="shortage-list">${shortageRows.map(renderShortageRow).join('')}</div>` : `<div class="empty-state compact">${icon('circle-check', 20)}<strong>这批物料没有出现负库存</strong><span>继续导入采购记录或新的 BOM 即可更新。</span></div>`}
      </div>
      <div class="section-panel category-panel">
        <div class="panel-header"><div><div class="eyebrow">MATERIAL MIX</div><h2>物料构成</h2></div><span class="panel-count">${rows.length} SKUs</span></div>
        <div class="category-list">${categories.map(([name, count]) => `<div class="category-row"><span>${escapeHtml(name)}</span><div class="category-bar"><i style="width:${Math.max(8, (count / rows.length) * 100)}%"></i></div><b>${count}</b></div>`).join('')}</div>
        <div class="category-foot">按识别到的物料类型统计，不含库位字段。</div>
      </div>
    </section>
    <section class="section-panel recent-panel">
      <div class="panel-header"><div><div class="eyebrow">LEDGER SNAPSHOT</div><h2>最近导入</h2></div><button class="text-button" data-view="imports">打开导入记录 ${icon('arrow-up-right', 14)}</button></div>
      <div class="recent-grid">${state.imports.slice(0, 4).map(renderImportCard).join('')}</div>
    </section>
  `;
}

function renderShortageRow(row) {
  return `<button class="shortage-row" data-focus-code="${escapeHtml(row.code)}"><span class="shortage-icon">${icon('triangle-alert', 15)}</span><span class="shortage-main"><strong>${escapeHtml(row.model || row.name)}</strong><small>${escapeHtml(row.code)} · ${escapeHtml(row.category)}</small></span><span class="shortage-number">${row.remaining < 0 ? formatQty(row.remaining) : '0'} <small>余量</small></span></button>`;
}

function renderImportCard(item) {
  const isBom = item.kind === 'bom';
  const isPlan = item.kind === 'plan';
  return `<div class="import-card"><div class="import-icon ${isBom ? 'is-bom' : isPlan ? 'is-plan' : ''}">${icon(isBom ? 'file-code-2' : isPlan ? 'clipboard-check' : 'receipt-text', 17)}</div><div><strong>${escapeHtml(item.name)}</strong><span>${escapeHtml(item.note)}</span></div><time>${escapeHtml(item.stamp)}</time></div>`;
}

function renderInventorySortButton(label, key) {
  const active = ui.sortKey === key;
  const direction = active && ui.sortDirection === 'desc' ? 'arrow-down-wide-narrow' : active ? 'arrow-up-narrow-wide' : 'arrow-up-down';
  return `<button type="button" class="table-sort-button ${active ? 'is-active' : ''}" data-sort-key="${key}" title="点击按${label}排序，再次点击切换顺序"><span>${label}</span>${icon(direction, 12)}</button>`;
}

function getSelectedInventoryRows() {
  return getInventory().filter((row) => ui.selectedCodes.has(row.code));
}

function renderInventory() {
  const rows = filteredInventory();
  const allRows = getInventory();
  const categories = [...new Set(allRows.map((row) => row.category))].sort();
  ui.selectedCodes.forEach((code) => { if (!allRows.some((row) => row.code === code)) ui.selectedCodes.delete(code); });
  const selectedCount = allRows.filter((row) => ui.selectedCodes.has(row.code)).length;
  const allVisibleSelected = rows.length > 0 && rows.every((row) => ui.selectedCodes.has(row.code));
  const someVisibleSelected = rows.some((row) => ui.selectedCodes.has(row.code));
  return `
    <div class="page-heading page-heading-tight"><div><div class="eyebrow">INVENTORY LEDGER</div><h1>库存明细</h1><p>每一行都是一个可核对的物料编号；同规格的阻容电感共享库存，余量由采购记录、BOM 用量和手工调整共同计算。</p></div><div class="heading-actions"><button class="button button-primary" id="new-part">${icon('plus')}<span>新建元件</span></button></div></div>
    <section class="inventory-toolbar"><label class="search-box">${icon('search', 17)}<input id="inventory-search" type="search" placeholder="搜索 C 编号、型号、品牌或类型" value="${escapeHtml(ui.query)}"/></label><div class="filter-group"><select id="status-filter" aria-label="库存状态"><option value="all" ${ui.status === 'all' ? 'selected' : ''}>全部状态</option><option value="shortage" ${ui.status === 'shortage' ? 'selected' : ''}>缺料</option><option value="depleted" ${ui.status === 'depleted' ? 'selected' : ''}>已用尽</option><option value="low" ${ui.status === 'low' ? 'selected' : ''}>偏低</option><option value="healthy" ${ui.status === 'healthy' ? 'selected' : ''}>充足</option></select><select id="category-filter" aria-label="物料分类"><option value="all">全部分类</option>${categories.map((category) => `<option value="${escapeHtml(category)}" ${ui.category === category ? 'selected' : ''}>${escapeHtml(category)}</option>`).join('')}</select></div></section>
    <section class="table-shell"><div class="table-meta"><div class="table-meta-main"><label class="select-all-control"><input type="checkbox" data-select-all aria-label="选择当前筛选的全部元件" ${allVisibleSelected ? 'checked' : ''} ${someVisibleSelected && !allVisibleSelected ? 'data-indeterminate="true"' : ''}/><span>全选当前 ${rows.length} 行</span></label><span>显示 ${rows.length} / ${allRows.length} 种物料 · 表头可排序，分类和状态可筛选</span></div><div class="table-meta-actions ${selectedCount ? 'has-selection' : ''}">${selectedCount ? `<span class="selection-count">已选 ${selectedCount} 项</span><input class="bulk-adjust-input" data-bulk-adjust-qty type="number" min="1" step="1" value="1" aria-label="批量调整数量"/><button class="icon-button" data-bulk-adjust="-1" title="批量减少" aria-label="批量减少">${icon('minus', 14)}</button><button class="icon-button" data-bulk-adjust="1" title="批量增加" aria-label="批量增加">${icon('plus', 14)}</button><button class="icon-button danger" data-bulk-delete title="删除选中元件" aria-label="删除选中元件">${icon('trash-2', 14)}</button>` : ''}<span class="table-meta-right">${icon('database', 14)}本地台账</span></div></div><div class="table-scroll"><table class="inventory-table"><thead><tr><th class="select-column" aria-label="选择"></th><th>${renderInventorySortButton('物料 / 制造商型号', 'model')}</th><th>${renderInventorySortButton('关键参数', 'spec')}</th><th>${renderInventorySortButton('封装', 'footprint')}</th><th>${renderInventorySortButton('分类', 'category')}</th><th>${renderInventorySortButton('采购', 'purchased')}</th><th>${renderInventorySortButton('已用', 'used')}</th><th>${renderInventorySortButton('手工调整', 'adjusted')}</th><th>${renderInventorySortButton('当前余量', 'remaining')}</th><th>${renderInventorySortButton('状态', 'status')}</th><th>操作</th></tr></thead><tbody id="inventory-body">${rows.length ? renderInventoryTableRows(rows) : `<tr><td colspan="11"><div class="empty-state">${icon('search-x', 22)}<strong>没有匹配的物料</strong><span>换个关键词或清除筛选条件。</span></div></td></tr>`}</tbody></table></div></section>
  `;
}

function renderInventoryTableRows(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    if (!groups.has(row.specKey)) groups.set(row.specKey, { key: row.specKey, label: row.specLabel, rows: [] });
    groups.get(row.specKey).rows.push(row);
  });
  return [...groups.values()]
    .sort((a, b) => {
      const direction = ui.sortDirection === 'desc' ? -1 : 1;
      const pickExtreme = (groupRows) => groupRows.reduce((best, row) => {
        const comparison = compareInventoryValues(inventorySortValue(row, ui.sortKey), inventorySortValue(best, ui.sortKey));
        return (direction === 1 ? comparison < 0 : comparison > 0) ? row : best;
      }, groupRows[0]);
      const aValue = pickExtreme(a.rows);
      const bValue = pickExtreme(b.rows);
      return compareInventoryValues(inventorySortValue(aValue, ui.sortKey), inventorySortValue(bValue, ui.sortKey)) * direction || a.label.localeCompare(b.label, 'zh-CN', { numeric: true });
    })
    .map((group) => {
      const id = encodeURIComponent(group.key);
      const collapsed = ui.collapsedGroups.has(id);
      const sortedRows = sortInventoryRows(group.rows);
      const groupRemaining = group.rows.reduce((sum, row) => sum + row.groupRemaining, 0) / group.rows.length;
      const groupAvailability = groupRemaining < 0 ? `缺口 ${formatQty(groupRemaining)}` : `共用余量 ${formatQty(groupRemaining)}`;
      return `<tr class="spec-group-row"><td colspan="11"><button class="spec-group-toggle" data-toggle-group="${escapeHtml(id)}" aria-expanded="${!collapsed}">${icon(collapsed ? 'chevron-right' : 'chevron-down', 15)}<span class="spec-group-name">${escapeHtml(group.label)}</span><span class="spec-group-count">${sortedRows.length} 个品牌 / 规格行 · ${groupAvailability}</span></button></td></tr>${collapsed ? '' : sortedRows.map(renderInventoryRow).join('')}`;
    }).join('');
}

function renderDetailPair(label, value) {
  if (!value) return '';
  return `<div class="detail-pair"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`;
}

function renderInventoryRow(row) {
  const amountClass = row.remaining < 0 ? 'negative' : row.remaining === 0 ? 'zero' : '';
  const adjustmentText = row.adjusted > 0 ? `+${formatQty(row.adjusted)}` : formatQty(row.adjusted);
  const details = row.details;
  const expanded = ui.expandedCodes.has(row.code);
  const powerText = details.powerWatts ? `${details.powerWatts} W` : details.power;
  const coreDetails = [
    renderDetailPair('值 / 阻值 / 容值', details.value),
    renderDetailPair('精度', details.tolerance),
    renderDetailPair('额定电压', details.voltageRatingMax),
    renderDetailPair('供电电压', details.voltageSupply),
    renderDetailPair('电压（提取）', !details.voltageRatingMax && !details.voltageSupply ? details.voltage : ''),
    renderDetailPair('温度系数', details.temperatureCoefficient),
    renderDetailPair('功率', powerText),
    renderDetailPair('工作电流', details.current),
    renderDetailPair('频率', details.frequency),
    renderDetailPair('工作温度', details.temperature),
  ].join('');
  const sharedDetails = row.shared ? renderDetailPair('同规格共享余量', `${formatQty(row.groupRemaining)} ${details.unit}`) : '';
  const usageDetails = [renderDetailPair('采购数量', `${formatQty(row.purchased)} ${details.unit}`), renderDetailPair('BOM 已用', `${formatQty(row.used)} ${details.unit}`), renderDetailPair('手工调整', `${adjustmentText} ${details.unit}`), renderDetailPair('当前余量', `${formatQty(row.remaining)} ${details.unit}`), sharedDetails, renderDetailPair('来源项目', row.projects.join('、')), renderDetailPair('位号', row.designators.join('、'))].join('');
  const priceDetails = [
    details.unitPrice ? renderDetailPair('采购单价', `¥${details.unitPrice}`) : '',
    details.jlcpcbPrice ? renderDetailPair('JLCPCB 价格', `¥${details.jlcpcbPrice}`) : '',
    details.lcscPrice ? renderDetailPair('LCSC 价格', `¥${details.lcscPrice}`) : '',
  ].join('');
  return `<tr class="inventory-main-row" data-code="${escapeHtml(row.code)}"><td class="select-cell"><input type="checkbox" data-select-code="${escapeHtml(row.code)}" aria-label="选择 ${escapeHtml(row.code)}" ${ui.selectedCodes.has(row.code) ? 'checked' : ''}/></td><td><div class="part-cell"><button class="detail-toggle ${expanded ? 'is-open' : ''}" data-toggle-details="${escapeHtml(row.code)}" title="${expanded ? '收起详情' : '展开详情'}" aria-label="${expanded ? '收起详情' : '展开详情'}" aria-expanded="${expanded}">${icon('chevron-right', 15)}</button><span class="part-swatch ${row.category === '芯片' ? 'chip' : ''}"></span><div><strong>${escapeHtml(row.model || row.name)}</strong><span>${escapeHtml(row.code)} · ${escapeHtml(row.brand || row.name)}</span></div></div></td><td><span class="spec-summary">${escapeHtml(details.value || '未标注')}</span><small class="spec-subline">${escapeHtml(details.tolerance || details.power || '')}</small></td><td><span class="package-value">${escapeHtml(details.footprint || '未提供')}</span></td><td><span class="category-tag">${escapeHtml(row.category)}</span></td><td class="number-cell">${formatQty(row.purchased)}</td><td class="number-cell used-cell">${formatQty(row.used)}</td><td class="number-cell adjustment-cell ${row.adjusted < 0 ? 'negative' : row.adjusted > 0 ? 'positive' : ''}">${adjustmentText}</td><td class="number-cell remaining-cell ${amountClass}">${formatQty(row.remaining)}</td><td><span class="status-pill ${statusClass(row.status)}"><i></i>${statusText(row.status)}</span></td><td><div class="row-actions"><button class="icon-button" data-quick-adjust="${escapeHtml(row.code)}" data-delta="-1" title="按右侧数量减少" aria-label="按右侧数量减少">${icon('minus', 14)}</button><input class="inline-adjust-input" data-inline-adjust-qty="${escapeHtml(row.code)}" type="number" min="1" step="1" value="1" aria-label="${escapeHtml(row.code)}调整数量"/><button class="icon-button" data-quick-adjust="${escapeHtml(row.code)}" data-delta="1" title="按左侧数量增加" aria-label="按左侧数量增加">${icon('plus', 14)}</button><button class="icon-button danger" data-delete-part="${escapeHtml(row.code)}" title="删除物料" aria-label="删除物料">${icon('trash-2', 14)}</button></div></td></tr><tr class="inventory-detail-row ${expanded ? 'is-open' : ''}" data-detail-for="${escapeHtml(row.code)}"><td colspan="11"><div class="detail-panel"><section class="detail-block"><span class="detail-block-label">核心参数</span><div class="detail-pairs">${renderDetailPair('封装', details.footprint)}${coreDetails || '<span class="detail-empty">原始文件未提供更多电气参数</span>'}</div></section><section class="detail-block"><span class="detail-block-label">分类与料号</span><div class="detail-pairs">${renderDetailPair('识别分类', row.category)}${renderDetailPair('原始分类', details.category)}${renderDetailPair('主分类', details.primaryCategory)}${renderDetailPair('次分类', details.secondaryCategory)}${renderDetailPair('立创编号', row.code)}${renderDetailPair('品牌', row.brand)}${renderDetailPair('制造商型号', details.manufacturerPart)}${renderDetailPair('供应商', details.supplier)}${priceDetails}</div></section><section class="detail-block"><span class="detail-block-label">使用追踪</span><div class="detail-pairs">${usageDetails || '<span class="detail-empty">尚未进入 BOM</span>'}</div></section>${details.description ? `<div class="detail-description"><span>原始描述</span><strong>${escapeHtml(details.description)}</strong></div>` : ''}</div></td></tr>`;
}

function renderPlanning() {
  const rows = getProcurementSuggestions();
  const demand = rows.reduce((sum, row) => sum + row.qty, 0);
  const shortage = rows.reduce((sum, row) => sum + row.shortage, 0);
  const shortageKinds = rows.filter((row) => row.shortage > 0).length;
  const coveredKinds = rows.filter((row) => row.shortage === 0).length;
  const planFiles = new Set(state.plans.map((item) => item.source || item.project)).size;

  return `
    <div class="page-heading page-heading-tight"><div><div class="eyebrow">MATERIAL PLANNING</div><h1>采购建议</h1><p>导入新板 BOM 后，与当前可用库存进行对比。这里是备料预估，不会扣减实际库存。</p></div><div class="heading-actions"><button class="button button-quiet" data-open-import="plan">${icon('clipboard-check')}<span>导入新板 BOM</span></button>${rows.length ? `<button class="button button-primary" id="export-procurement" ${shortage ? '' : 'disabled'} title="只导出还需采购的物料，数量会扣除现有库存">${icon('file-down')}<span>导出嘉立创采购表</span></button><button class="button button-quiet" id="clear-plan">${icon('trash-2')}<span>清空计划</span></button>` : ''}</div></div>
    <section class="stat-grid plan-stat-grid" aria-label="采购建议概览">
      <article class="stat-card stat-card-warning"><div class="stat-label">还需采购</div><div class="stat-value">${formatQty(shortage)}<small> pcs</small></div><div class="stat-foot">所有新板 BOM 的缺口合计</div></article>
      <article class="stat-card"><div class="stat-label">新板需求</div><div class="stat-value">${formatQty(demand)}<small> pcs</small></div><div class="stat-foot">${planFiles} 份备料计划</div></article>
      <article class="stat-card"><div class="stat-label">缺料种类</div><div class="stat-value">${shortageKinds}<small> 种</small></div><div class="stat-foot">需要下单补充的规格</div></article>
      <article class="stat-card stat-card-accent"><div class="stat-label">库存可满足</div><div class="stat-value">${coveredKinds}<small> 种</small></div><div class="stat-foot">现有库存足够的规格</div></article>
    </section>
    ${rows.length ? `<section class="table-shell plan-shell"><div class="table-meta"><span>共 ${rows.length} 种规格 · 按同规格库存合并，不区分品牌</span><span class="table-meta-right">${icon('clipboard-list', 14)}备料计划</span></div><div class="table-scroll"><table class="plan-table"><thead><tr><th>物料 / 型号</th><th>规格 / 封装</th><th>分类</th><th>新板需求</th><th>现有可用</th><th>还需采购</th><th>状态</th><th>来源</th></tr></thead><tbody>${rows.map(renderPlanningRow).join('')}</tbody></table></div></section>` : `<section class="section-panel empty-plan"><div class="empty-state compact">${icon('clipboard-plus', 25)}<strong>还没有新板备料计划</strong><span>导入一份新板 BOM，系统会马上列出需要采购的元件。</span><button class="button button-primary" data-open-import="plan">${icon('upload')}<span>导入新板 BOM</span></button></div></section>`}
  `;
}

function renderPlanningRow(row) {
  const brandText = row.brands.length ? row.brands.join('、') : '未标注品牌';
  const codeText = row.codes.length ? row.codes.join('、') : '无 C 编号';
  const sourceText = row.projects.length ? row.projects.join('、') : '新板 BOM';
  return `<tr class="plan-row ${row.shortage > 0 ? 'plan-needs-buy' : 'plan-covered'}" data-plan-spec="${escapeHtml(row.specKey)}"><td><div class="plan-part"><span class="part-swatch ${row.category === '芯片' ? 'chip' : ''}"></span><div><strong>${escapeHtml(row.modelLabel)}</strong><span>${escapeHtml(codeText)} · ${escapeHtml(brandText)}</span></div></div></td><td><span class="spec-summary">${escapeHtml(row.value || '未标注')}</span><small class="spec-subline">${escapeHtml(row.footprint || '未提供')}</small></td><td><span class="category-tag">${escapeHtml(row.category)}</span></td><td class="number-cell">${formatQty(row.qty)} ${escapeHtml(row.unit)}</td><td class="number-cell plan-available">${formatQty(row.currentAvailable)} ${escapeHtml(row.unit)}</td><td class="number-cell plan-shortage ${row.shortage > 0 ? 'negative' : ''}">${formatQty(row.shortage)} ${escapeHtml(row.unit)}</td><td><span class="plan-status ${row.shortage > 0 ? 'plan-status-buy' : 'plan-status-ready'}"><i></i>${row.shortage > 0 ? '需要采购' : '库存足够'}</span></td><td><span class="plan-source">${escapeHtml(sourceText)}</span></td></tr>`;
}

function renderImports() {
  return `
    <div class="page-heading page-heading-tight"><div><div class="eyebrow">IMPORT PIPELINE</div><h1>导入记录</h1><p>支持采购记录、已使用 BOM 和新板备料 BOM。批量重置会先解析所选文件，再替换当前台账；1/2/3 可作为已用 BOM，01/02/03 可作为采购记录。</p></div><div class="heading-actions"><button class="button button-quiet" data-clear-ledger>${icon('trash-2')}<span>清空当前台账</span></button><button class="button button-quiet" data-open-reset-import>${icon('file-input')}<span>重置并导入</span></button><button class="button button-quiet" data-open-import="bom">${icon('file-up')}<span>导入已用 BOM</span></button><button class="button button-quiet" id="new-part">${icon('plus')}<span>新建元件</span></button><button class="button button-primary" data-open-import="purchase">${icon('shopping-bag')}<span>导入采购</span></button></div></div>
    <section class="import-drop-grid"><button class="drop-zone" data-open-import="purchase"><span class="drop-icon">${icon('receipt-text', 22)}</span><strong>采购记录</strong><span>导入 .xls / .xlsx / .csv</span><small>按“商品编号 + 订购数量”读取</small></button><button class="drop-zone drop-zone-bom" data-open-import="bom"><span class="drop-icon">${icon('file-code-2', 22)}</span><strong>已用 BOM</strong><span>导入 .xlsx / .csv</span><small>会计入实际已用数量</small></button><button class="drop-zone drop-zone-plan" data-open-import="plan"><span class="drop-icon">${icon('clipboard-check', 22)}</span><strong>新板备料 BOM</strong><span>导入 .xlsx / .csv</span><small>只对比库存，不扣实际库存</small></button></section>
    <section class="import-rules"><div class="rule-item"><span class="rule-number">01</span><div><strong>同规格共享库存</strong><span>阻容电感按值和封装匹配，忽略品牌与 C 编号；芯片等器件仍按型号区分。</span></div></div><div class="rule-item"><span class="rule-number">02</span><div><strong>手工增减可追溯</strong><span>库存明细支持快速加减，也可以输入任意数量和备注。</span></div></div><div class="rule-item"><span class="rule-number">03</span><div><strong>没有可替代库存才缺料</strong><span>只有同规格共享库存也无法覆盖 BOM 用量时，才会显示缺料。</span></div></div></section>
    <section class="section-panel import-history"><div class="panel-header"><div><div class="eyebrow">SOURCE FILES</div><h2>已载入文件</h2></div><span class="panel-count">${state.imports.length} 条记录</span></div><div class="history-list">${state.imports.map(renderImportCard).join('')}</div></section>
  `;
}

function renderContent() {
  const content = document.getElementById('content');
  if (!content) return;
  content.innerHTML = ui.view === 'overview' ? renderOverview() : ui.view === 'inventory' ? renderInventory() : ui.view === 'planning' ? renderPlanning() : renderImports();
}

function addAdjustment(code, qty, note = '快速调整', options = {}) {
  const amount = numberValue(qty);
  if (!amount) return;
  state.adjustments.push({ id: crypto.randomUUID(), code, qty: amount, note, createdAt: new Date().toISOString() });
  saveState();
  if (options.preserveScroll) refreshInventoryTable();
  else render();
  showToast(`${amount > 0 ? '已增加' : '已扣减'} ${formatQty(Math.abs(amount))} 个`);
}

function getInlineAdjustmentAmount(button) {
  const row = button.closest('tr.inventory-main-row');
  const input = row?.querySelector('[data-inline-adjust-qty]');
  return Math.max(1, Math.round(numberValue(input?.value || 1)));
}

function adjustSelectedParts(delta) {
  const selectedRows = getSelectedInventoryRows();
  const input = document.querySelector('[data-bulk-adjust-qty]');
  const amount = Math.max(1, Math.round(numberValue(input?.value || 1)));
  if (!selectedRows.length) return;
  selectedRows.forEach((row) => state.adjustments.push({ id: crypto.randomUUID(), code: row.code, qty: delta * amount, note: '批量调整', createdAt: new Date().toISOString() }));
  saveState();
  refreshInventoryTable();
  showToast(`${delta > 0 ? '已批量增加' : '已批量扣减'} ${formatQty(amount)} 个，共 ${selectedRows.length} 项`);
}

function deletePart(code) {
  const row = getInventory().find((item) => item.code === code);
  if (!row || !window.confirm(`确定删除“${row.model || row.name}”及其台账记录吗？`)) return;
  state.purchases = state.purchases.filter((item) => item.code !== code);
  state.usage = state.usage.filter((item) => item.code !== code);
  state.adjustments = state.adjustments.filter((item) => item.code !== code);
  ui.selectedCodes.delete(code);
  saveState();
  render();
  showToast(`已删除 ${row.model || row.name}`);
}

function deleteSelectedParts() {
  const selectedRows = getSelectedInventoryRows();
  if (!selectedRows.length || !window.confirm(`确定删除选中的 ${selectedRows.length} 个元件及其台账记录吗？`)) return;
  const codes = new Set(selectedRows.map((row) => row.code));
  state.purchases = state.purchases.filter((item) => !codes.has(item.code));
  state.usage = state.usage.filter((item) => !codes.has(item.code));
  state.adjustments = state.adjustments.filter((item) => !codes.has(item.code));
  ui.selectedCodes.clear();
  saveState();
  render();
  showToast(`已删除 ${selectedRows.length} 个元件`);
}

function clearPlan() {
  if (!state.plans.length || !window.confirm('确定清空当前新板备料计划吗？实际库存不会受到影响。')) return;
  state.plans = [];
  state.planSources = [];
  saveState();
  render();
  showToast('已清空新板备料计划');
}

function openAdjustModal(code) {
  const row = getInventory().find((item) => item.code === code);
  modalState = { kind: 'adjust', code, label: row?.model || row?.name || code };
  renderModal();
  bindModalEvents();
}

function openNewPartModal() {
  modalState = { kind: 'new' };
  renderModal();
  bindModalEvents();
}

function closeModal() {
  modalState = null;
  renderModal();
}

function renderModal() {
  const root = document.getElementById('modal-root');
  if (!root) return;
  if (!modalState) {
    root.innerHTML = '';
    return;
  }
  if (modalState.kind === 'adjust') {
    root.innerHTML = `<div class="modal-backdrop" data-close-modal><section class="modal" role="dialog" aria-modal="true" aria-labelledby="adjust-title"><div class="modal-head"><div><div class="eyebrow">MANUAL ADJUSTMENT</div><h2 id="adjust-title">调整数量</h2><p>${escapeHtml(modalState.label)} · ${escapeHtml(modalState.code)}</p></div><button class="icon-button" data-close-modal title="关闭" aria-label="关闭">${icon('x')}</button></div><form id="adjust-form"><label class="form-field"><span>操作</span><select id="adjust-operation"><option value="1">增加库存</option><option value="-1">扣减库存</option></select></label><label class="form-field"><span>数量</span><input id="adjust-quantity" type="number" min="1" step="1" value="1" required /></label><label class="form-field"><span>备注 <small>可选</small></span><input id="adjust-note" type="text" placeholder="例如：盘点发现多 2 个" maxlength="80" /></label><div class="modal-actions"><button type="button" class="button button-quiet" data-close-modal>取消</button><button type="submit" class="button button-primary">保存调整</button></div></form></section></div>`;
  } else {
    root.innerHTML = `<div class="modal-backdrop" data-close-modal><section class="modal modal-wide" role="dialog" aria-modal="true" aria-labelledby="new-part-title"><div class="modal-head"><div><div class="eyebrow">LOCAL PART</div><h2 id="new-part-title">新建元件</h2><p>录入后会作为一条采购数量加入库存。</p></div><button class="icon-button" data-close-modal title="关闭" aria-label="关闭">${icon('x')}</button></div><form id="new-part-form"><div class="form-grid"><label class="form-field"><span>立创编号 <small>可选</small></span><input id="new-code" type="text" placeholder="例如 C123456" /></label><label class="form-field"><span>型号 / 名称</span><input id="new-model" type="text" placeholder="例如 OPA564AIDWPR" required /></label><label class="form-field"><span>分类 <small>可选</small></span><input id="new-type" type="text" placeholder="例如 芯片 / 电阻" /></label><label class="form-field"><span>品牌 <small>可选</small></span><input id="new-brand" type="text" placeholder="例如 TI" /></label><label class="form-field"><span>封装 <small>可选</small></span><input id="new-package" type="text" placeholder="例如 SOIC-8" /></label><label class="form-field"><span>初始数量</span><input id="new-qty" type="number" min="0" step="1" value="1" required /></label></div><div class="modal-actions"><button type="button" class="button button-quiet" data-close-modal>取消</button><button type="submit" class="button button-primary">加入库存</button></div></form></section></div>`;
  }
  createIcons({ icons });
}

function bindModalEvents() {
  document.querySelectorAll('[data-close-modal]').forEach((element) => element.addEventListener('click', (event) => {
    if (event.currentTarget === event.target || event.currentTarget.tagName === 'BUTTON') closeModal();
  }));
  document.getElementById('adjust-form')?.addEventListener('submit', (event) => {
    event.preventDefault();
    const sign = Number(document.getElementById('adjust-operation').value);
    const qty = numberValue(document.getElementById('adjust-quantity').value);
    const note = document.getElementById('adjust-note').value.trim() || '手工调整';
    const code = modalState.code;
    modalState = null;
    if (qty > 0) addAdjustment(code, sign * qty, note);
  });
  document.getElementById('new-part-form')?.addEventListener('submit', (event) => {
    event.preventDefault();
    const model = document.getElementById('new-model').value.trim();
    const code = document.getElementById('new-code').value.trim() || `LOCAL-${Date.now()}`;
    const qty = numberValue(document.getElementById('new-qty').value);
    if (!model) return;
    state.purchases.push({ code, model, name: model, type: document.getElementById('new-type').value.trim(), brand: document.getElementById('new-brand').value.trim(), package: document.getElementById('new-package').value.trim(), qty, unit: '个', price: 0, source: '手动新增' });
    state.imports.unshift({ id: crypto.randomUUID(), kind: 'purchase', name: `手动新增 · ${model}`, note: `加入 ${formatQty(qty)} 个`, stamp: new Date().toLocaleString('zh-CN', { hour12: false }) });
    saveState();
    modalState = null;
    render();
    showToast(`已新建 ${model}`);
  });
}

function showToast(message, tone = 'success') {
  ui.toast = { message, tone };
  const toast = document.getElementById('toast');
  if (!toast) return;
  toast.className = `toast is-visible ${tone}`;
  toast.innerHTML = `${icon(tone === 'error' ? 'circle-alert' : 'circle-check', 16)}<span>${escapeHtml(message)}</span>`;
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => { toast.className = 'toast'; }, 3200);
}

function exportInventory() {
  const rows = getInventory();
  const output = [['立创编号', '型号', '分类', '采购数量', '已用数量', '手工调整', '当前余量', '状态', '来源项目']];
  rows.forEach((row) => output.push([row.code, row.model || row.name, row.category, row.purchased, row.used, row.adjusted, row.remaining, statusText(row.status), row.projects.join('、')]));
  const blob = new Blob([`\ufeff${output.map((row) => row.map((cell) => `"${String(cell ?? '').replaceAll('"', '""')}"`).join(',')).join('\n')}`], { type: 'text/csv;charset=utf-8' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `元件余量清单-${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  URL.revokeObjectURL(link.href);
  showToast('库存清单已导出');
}

async function handleImport(kind, file) {
  try {
    const count = kind === 'purchase' ? await importPurchaseFile(file) : kind === 'bom' ? await importBomFile(file) : await importPlanFile(file);
    render();
    showToast(kind === 'purchase' ? `已导入 ${count} 条采购记录` : kind === 'bom' ? `已导入 ${count} 条 BOM 用量` : `已生成 ${count} 条新板备料需求`);
  } catch (error) {
    showToast(error.message || '导入失败，请检查表格格式。', 'error');
  }
}

function openImport(kind) {
  const input = document.getElementById(kind === 'purchase' ? 'purchase-file' : kind === 'bom' ? 'bom-file' : 'plan-file');
  input?.click();
}

function openResetImport() {
  document.getElementById('reset-import-files')?.click();
}

function focusCode(code) {
  ui.view = 'inventory';
  ui.query = code;
  ui.status = 'all';
  render();
  window.setTimeout(() => document.querySelector(`[data-code="${CSS.escape(code)}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'center' }), 60);
}

function bindEvents() {
  document.getElementById('purchase-file')?.addEventListener('change', (event) => event.target.files[0] && handleImport('purchase', event.target.files[0]));
  document.getElementById('bom-file')?.addEventListener('change', (event) => event.target.files[0] && handleImport('bom', event.target.files[0]));
  document.getElementById('plan-file')?.addEventListener('change', (event) => event.target.files[0] && handleImport('plan', event.target.files[0]));
  document.getElementById('reset-import-files')?.addEventListener('change', (event) => {
    const files = [...event.target.files];
    event.target.value = '';
    if (files.length) handleResetImport(files);
  });
  document.getElementById('export-csv')?.addEventListener('click', exportInventory);
  document.querySelectorAll('[data-open-reset-import]').forEach((button) => button.addEventListener('click', openResetImport));
  document.querySelectorAll('[data-clear-ledger]').forEach((button) => button.addEventListener('click', resetLedger));
  document.getElementById('reset-demo')?.addEventListener('click', () => {
    state = createInitialState();
    saveState();
    ui.query = '';
    ui.status = 'all';
    ui.selectedCodes.clear();
    render();
    showToast('已恢复初始空台账');
  });
}

function bindContentEvents() {
  document.querySelectorAll('[data-view]').forEach((button) => button.addEventListener('click', () => {
    ui.view = button.dataset.view;
    if (button.dataset.status) ui.status = button.dataset.status;
    render();
  }));
  document.querySelectorAll('[data-open-import]').forEach((button) => button.addEventListener('click', () => openImport(button.dataset.openImport)));
  document.querySelectorAll('[data-focus-code]').forEach((button) => button.addEventListener('click', () => focusCode(button.dataset.focusCode)));
  document.getElementById('new-part')?.addEventListener('click', openNewPartModal);
  document.getElementById('clear-plan')?.addEventListener('click', clearPlan);
  document.getElementById('export-procurement')?.addEventListener('click', exportProcurement);
  document.querySelectorAll('[data-sort-key]').forEach((button) => button.addEventListener('click', () => {
    const key = button.dataset.sortKey;
    if (ui.sortKey === key) ui.sortDirection = ui.sortDirection === 'asc' ? 'desc' : 'asc';
    else { ui.sortKey = key; ui.sortDirection = 'asc'; }
    refreshInventoryView();
  }));
  bindInventoryRowEvents();
  document.querySelectorAll('[data-toggle-details]').forEach((button) => button.addEventListener('click', () => {
    const code = button.dataset.toggleDetails;
    if (ui.expandedCodes.has(code)) ui.expandedCodes.delete(code); else ui.expandedCodes.add(code);
    refreshInventoryTable();
  }));
  document.querySelectorAll('[data-toggle-group]').forEach((button) => button.addEventListener('click', () => {
    const groupId = button.dataset.toggleGroup;
    if (ui.collapsedGroups.has(groupId)) ui.collapsedGroups.delete(groupId); else ui.collapsedGroups.add(groupId);
    refreshInventoryTable();
  }));
  document.getElementById('inventory-search')?.addEventListener('input', (event) => { ui.query = event.target.value; refreshInventoryTable(); });
  document.getElementById('status-filter')?.addEventListener('change', (event) => { ui.status = event.target.value; renderContent(); bindContentEvents(); createIcons({ icons }); });
  document.getElementById('category-filter')?.addEventListener('change', (event) => { ui.category = event.target.value; renderContent(); bindContentEvents(); createIcons({ icons }); });
}

function bindInventoryRowEvents() {
  document.querySelectorAll('[data-quick-adjust]').forEach((button) => button.addEventListener('click', () => addAdjustment(button.dataset.quickAdjust, Number(button.dataset.delta) * getInlineAdjustmentAmount(button), '快速调整', { preserveScroll: true })));
  document.querySelectorAll('[data-delete-part]').forEach((button) => button.addEventListener('click', () => deletePart(button.dataset.deletePart)));
  document.querySelectorAll('[data-select-code]').forEach((checkbox) => checkbox.addEventListener('change', () => {
    if (checkbox.checked) ui.selectedCodes.add(checkbox.dataset.selectCode);
    else ui.selectedCodes.delete(checkbox.dataset.selectCode);
    refreshInventoryView();
  }));
  document.querySelector('[data-select-all]')?.addEventListener('change', (event) => {
    const visibleCodes = filteredInventory().map((row) => row.code);
    visibleCodes.forEach((code) => event.target.checked ? ui.selectedCodes.add(code) : ui.selectedCodes.delete(code));
    refreshInventoryView();
  });
  document.querySelectorAll('[data-bulk-adjust]').forEach((button) => button.addEventListener('click', () => adjustSelectedParts(Number(button.dataset.bulkAdjust))));
  document.querySelector('[data-bulk-delete]')?.addEventListener('click', deleteSelectedParts);
  updateInventorySelectionUi();
}

function refreshInventoryTable() {
  const body = document.getElementById('inventory-body');
  if (!body) return;
  const scroll = body.closest('.table-scroll');
  const scrollTop = scroll?.scrollTop || 0;
  const rows = filteredInventory();
  body.innerHTML = rows.length ? renderInventoryTableRows(rows) : `<tr><td colspan="11"><div class="empty-state">${icon('search-x', 22)}<strong>没有匹配的物料</strong><span>换个关键词或清除筛选条件。</span></div></td></tr>`;
  bindInventoryRowEvents();
  document.querySelectorAll('[data-toggle-details]').forEach((button) => button.addEventListener('click', () => { const code = button.dataset.toggleDetails; if (ui.expandedCodes.has(code)) ui.expandedCodes.delete(code); else ui.expandedCodes.add(code); refreshInventoryTable(); }));
  document.querySelectorAll('[data-toggle-group]').forEach((button) => button.addEventListener('click', () => { const groupId = button.dataset.toggleGroup; if (ui.collapsedGroups.has(groupId)) ui.collapsedGroups.delete(groupId); else ui.collapsedGroups.add(groupId); refreshInventoryTable(); }));
  if (scroll) scroll.scrollTop = scrollTop;
  createIcons({ icons });
}

function updateInventorySelectionUi() {
  const visibleRows = filteredInventory();
  const selectAll = document.querySelector('[data-select-all]');
  if (selectAll) {
    const selectedVisible = visibleRows.filter((row) => ui.selectedCodes.has(row.code)).length;
    selectAll.checked = visibleRows.length > 0 && selectedVisible === visibleRows.length;
    selectAll.indeterminate = selectedVisible > 0 && selectedVisible < visibleRows.length;
  }
}

function refreshInventoryView() {
  const scroll = document.querySelector('.table-scroll');
  const scrollTop = scroll?.scrollTop || 0;
  renderContent();
  bindContentEvents();
  createIcons({ icons });
  const nextScroll = document.querySelector('.table-scroll');
  if (nextScroll) nextScroll.scrollTop = scrollTop;
}

function render() {
  document.getElementById('app').innerHTML = renderShell();
  renderContent();
  renderModal();
  bindEvents();
  bindContentEvents();
  bindModalEvents();
  createIcons({ icons });
}

render();
