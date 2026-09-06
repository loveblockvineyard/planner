import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Droplets, Clock, AlertTriangle, Settings, LogOut, Plus, Trash2,
  Download, Upload, X, Check, RefreshCw, Users, Layers, Pencil,
  ChevronRight, MapPin, WifiOff, Beaker, ChevronLeft, Cloud, LayoutDashboard, Wind, Droplet, Thermometer, Wrench, Mic, Fuel, Truck
} from 'lucide-react';
import logoUrl from './assets/logo.png';
import Papa from 'papaparse';
import * as XLSX from 'xlsx-js-style';
import { loadJSON, saveJSON, subscribe } from './db';
import { ensureReady, isConfigured } from './firebase';
const CLOUD_ON = isConfigured;

/* ============================================================
   Loveblock Vineyard Ops — installable PWA, offline-first.
   Operator app + manager console. Data is cached on the device and
   syncs automatically when a connection returns.
   ============================================================ */

// Fire-and-forget POST to the manager's notification webhook (Apps Script / Zapier / Make).
// Uses no-cors + text/plain so it works cross-origin without the endpoint needing CORS headers.
async function postWebhook(url, payload) {
  if (!url) return false;
  try {
    await fetch(url, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
    });
    return true;
  } catch { return false; }
}

const K = {
  config: 'vineyard:config',
  spraysLegacy: 'vineyard:sprays',
  sprays: type => `vineyard:sprays:${type}`,
  work: 'vineyard:work',
  maint: 'vineyard:maintenance',
  fuel: 'vineyard:fuel',
  hours: 'vineyard:hours',
  rm: 'vineyard:rm',
  el: 'vineyard:el',
  disease: 'vineyard:disease',
  workDone: 'vineyard:work:done',
  ts: code => `vineyard:ts:${code}`,
  hz: code => `vineyard:hz:${code}`,
};

/* Live-update hook: re-reads a key whenever any device changes it.
   `guard` lets a component ignore echoes of its own just-saved write. */
function useLiveKey(key, apply, deps = []) {
  useEffect(() => {
    const off = subscribe(key, value => apply(value));
    return off;
  }, deps);
}

/* Modified E-L system growth stages */
const EL_STAGES = [
  [1, 'Winter bud'], [2, 'Bud scales opening'], [3, 'Wooly bud — green showing'],
  [4, 'Budburst — leaf tips visible'], [7, 'First leaf separated from shoot tip'],
  [9, '2–3 leaves separated, shoots 2–4 cm'], [11, '4 leaves separated'],
  [12, '5 leaves separated, shoots ~10 cm, inflorescence clear'], [13, '6 leaves separated'],
  [14, '7 leaves separated'], [15, '8 leaves separated, shoot elongating rapidly'],
  [16, '10 leaves separated'], [17, '12 leaves separated, inflorescences well developed'],
  [18, '14 leaves separated, cap colour fading from green'],
  [19, '~16 leaves separated, beginning of flowering'], [20, '10% caps off'],
  [21, '30% caps off'], [23, '17–20 leaves, 50% caps off — flowering'],
  [25, '80% caps off'], [26, 'Cap fall complete'],
  [27, 'Setting, young berries enlarging (>2 mm)'], [29, 'Berries pepper-corn size (4 mm)'],
  [31, 'Berries pea size (7 mm)'], [32, 'Beginning of bunch closure, berries touching'],
  [33, 'Berries still hard and green'], [34, 'Berries begin to soften, sugar increasing'],
  [35, 'Berries begin to colour and enlarge'], [36, 'Berries with intermediate sugar'],
  [37, 'Berries not quite ripe'], [38, 'Berries harvest ripe'], [39, 'Berries over ripe'],
  [41, 'After harvest: cane maturation complete'], [43, 'Beginning of leaf fall'],
  [47, 'End of leaf fall'],
];
const elLabel = code => { const f = EL_STAGES.find(x => x[0] === Number(code)); return f ? f[1] : ''; };

/* Disease monitoring vocabulary */
const DISEASES = ['Downy Mildew', 'Powdery Mildew', 'Botrytis', 'Mealybug'];
const INCIDENCE = ['None', 'Very low (1 per bay)', 'Low (2–3 per bay)', 'Medium (5–10 per bay)', 'High (over 10 per bay)'];
const FOUND_ON = ['None', 'Leaf', 'Bunch'];
const SEVERITY = ['0%', '5%', '10%', '30%', '50%', '70%', '100%'];

/* Block certification. Organic and in-conversion blocks may only be sprayed
   with BioGro-certified products that are approved for use. */
const BLOCK_CERT = {
  'Hill - C18 - SB': 'Conversion', 'Hill - C18 - CH': 'Conversion',
  'SB 01': 'Conversion', 'SB 02': 'Conversion', 'WB - SYR': 'Conversion',
  'Woolshed - Pinot Gris': 'Organic', 'Woolshed - Sauvignon Blanc': 'Organic',
  'Eros - Front PG': 'Organic', 'Eros - Back SB': 'Organic', 'Eros - Front SB': 'Organic',
  'Eros - SB2020': 'Organic', 'Eros - TG2015': 'Organic', 'Eros - TG2016': 'Organic', 'Eros - TG2017': 'Organic',
  'Hill - A 23': 'SWNZ', 'Hill - A SB': 'SWNZ', 'Hill - E GEW': 'SWNZ', 'Hill - E PG': 'SWNZ',
  'Hill - E SB': 'SWNZ', 'Hill - F PG': 'SWNZ', 'Hill - F RSL': 'SWNZ', 'Hill - F SB': 'SWNZ',
  'WB - RSL': 'SWNZ', 'WB - PG': 'SWNZ', 'WB - CHA': 'SWNZ',
  'SB 03': 'SWNZ', 'SB 04': 'SWNZ', 'SB 05': 'SWNZ',
};
const CERTS = ['', 'Organic', 'Conversion', 'SWNZ'];
const MACHINE_TYPES = ['Tractor', 'Sprayer', 'Vehicle', 'Equipment', 'Harvester', 'Other'];
const PRODUCT_CATEGORIES = ['Powdery Mildew', 'Downy Mildew', 'Botrytis', 'Mealy bug', 'Nutrition', 'Spreader/Adjuvant'];
const RATE_BASES = [{ key: 'per100', label: 'per 100 L' }, { key: 'perHa', label: 'per hectare' }];
// organic and conversion blocks are restricted to certified, approved products
const certRestricted = cert => cert === 'Organic' || cert === 'Conversion';
// products in the mix that an organic or in-conversion block can't take
function nonOrganicInMix(mix, products) {
  return (mix || []).map(m => (products || []).find(p => p.name === m.product) || { name: m.product })
    .filter(p => !p.biogro || p.approved === false);
}
const certOf = (blockName, config) => {
  const b = (config.blocks || []).find(x => x.name === blockName);
  return (b && b.cert) || '';
};
const certTone = cert => cert === 'Organic' ? 'bg-emerald-50 border-emerald-300 text-emerald-800'
  : cert === 'Conversion' ? 'bg-lime-50 border-lime-300 text-lime-800'
  : cert === 'SWNZ' ? 'bg-sky-50 border-sky-200 text-sky-800' : 'bg-stone-100 border-stone-200 text-stone-500';

/* ============================================================
   Work rates — how long a job should take
   Tractor tasks run on km/h over the block's vine rows, plus a turning
   allowance for the headlands. Hand tasks run on plants per hour.
   ============================================================ */
const DEFAULT_VINE_SPACING = 1.8;      // metres between vines, used when a block has no count
const DEFAULT_ROW_WIDTH = 2.7;         // metres between rows, used when a block has none
const WORK_DAY_HOURS = 8;

const rowWidthOf = (b, config) =>
  numOf((b || {}).rowWidth) || numOf((config || {}).rowWidth) || DEFAULT_ROW_WIDTH;

// km of vine row: the recorded figure, else worked out from area ÷ row width
function blockKm(b, config) {
  if (!b) return 0;
  if (numOf(b.km) > 0) return numOf(b.km);
  const w = rowWidthOf(b, config);
  return w > 0 ? Math.round((numOf(b.ha) * 10000 / w) / 1000 * 100) / 100 : 0;
}
// vines in a block: use the recorded count, else derive from km of vine row
function blockVines(b, config) {
  if (!b) return 0;
  if (numOf(b.vines) > 0) return numOf(b.vines);
  const spacing = numOf((config || {}).vineSpacing) || DEFAULT_VINE_SPACING;
  return spacing > 0 ? Math.round((blockKm(b, config) * 1000) / spacing) : 0;
}
// planned hours for one block of one task, null when no pace is set
function plannedHours(task, blockName, config) {
  const pace = ((config || {}).workPace || {})[task];
  if (!pace || !numOf(pace.value)) return null;
  const b = (config.blocks || []).find(x => x.name === blockName);
  if (!b) return null;
  if (pace.type === 'kmh') {
    const km = blockKm(b, config);
    if (!km) return null;
    const run = km / numOf(pace.value);
    return Math.round(run * (1 + numOf(pace.headland) / 100) * 100) / 100;   // + headland turning
  }
  const vines = blockVines(b, config);
  if (!vines) return null;
  return Math.round((vines / numOf(pace.value)) * 100) / 100;
}
const paceLabel = pace => !pace || !numOf(pace.value) ? ''
  : pace.type === 'kmh' ? `${fmtNum(pace.value)} km/h${numOf(pace.headland) ? ` +${fmtNum(pace.headland)}%` : ''}`
  : `${fmtNum(pace.value)} plants/h`;
// spread hours over working days of 8 h, returning day offsets
const daysFromHours = h => Math.max(1, Math.ceil(numOf(h) / WORK_DAY_HOURS));

/* Every spreadsheet we hand out gets the same treatment: columns wide enough to
   read without dragging, a gridded border, and a bold, slightly larger header. */
const XL_BORDER = { style: 'thin', color: { rgb: 'D6D3D1' } };
function dressSheet(ws) {
  if (!ws || !ws['!ref']) return ws;
  const range = XLSX.utils.decode_range(ws['!ref']);
  const cols = [];
  for (let C = range.s.c; C <= range.e.c; C++) {
    let width = 9;
    for (let R = range.s.r; R <= range.e.r; R++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (!cell) continue;
      const text = String(cell.v == null ? '' : cell.v);
      const isHeader = R === range.s.r;
      width = Math.max(width, Math.min(58, text.length + (isHeader ? 4 : 3)));
      cell.s = {
        font: isHeader ? { bold: true, sz: 13 } : { sz: 11 },
        alignment: { vertical: 'center', wrapText: false },
        border: { top: XL_BORDER, bottom: XL_BORDER, left: XL_BORDER, right: XL_BORDER },
        ...(isHeader ? { fill: { patternType: 'solid', fgColor: { rgb: 'EFEBE2' } } } : {}),
      };
    }
    cols.push({ wch: width });
  }
  ws['!cols'] = cols;
  ws['!rows'] = [{ hpt: 22 }];                 // taller header row
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };    // keep headers in view
  return ws;
}
// append a sheet with the house formatting applied
function addSheet(wb, ws, name) { return addSheet(wb, dressSheet(ws), name); }

/* ---------- defaults / seed ---------- */
const DEFAULT_CONFIG = {
  siteName: 'Loveblock',
  managerCode: '0000',
  techCode: '2222',
  techName: 'Maria Romero',
  operators: [{ code: '1234', name: 'Sample Operator' }],
  blocks: [
    { name: 'Hill - A SB', ha: 0.64, rows: '267-274', km: 2.13, cert: 'SWNZ' },
    { name: 'Hill - C18 - CH', ha: 1.08, rows: '61-79', km: 3.6, cert: 'Conversion' },
    { name: 'Hill - C18 - SB', ha: 1, rows: '38-60', km: 3.33, cert: 'Conversion' },
    { name: 'Hill - E GEW', ha: 2.96, rows: '58-98', km: 9.87, cert: 'SWNZ' },
    { name: 'Hill - E PG', ha: 1.55, rows: '18-57', km: 5.17, cert: 'SWNZ' },
    { name: 'Hill - E SB', ha: 0.28, rows: '1-17', km: 0.93, cert: 'SWNZ' },
    { name: 'Hill - F PG', ha: 8.424, rows: '19-113', km: 28.08, cert: 'SWNZ' },
    { name: 'Hill - F RSL', ha: 2.98, rows: '114-146', km: 9.93, cert: 'SWNZ' },
    { name: 'Hill - F SB', ha: 1.69, rows: '1-18', km: 5.63, cert: 'SWNZ' },
    { name: 'Hill - A 23', ha: 7.9, rows: '42-180', km: 26.69, cert: 'SWNZ' },
    { name: 'Eros - Front PG', ha: 5.86, rows: '201-139', km: 19.53, cert: 'Organic' },
    { name: 'Eros - Front SB', ha: 4.01, rows: '202-242', km: 13.37, cert: 'Organic' },
    { name: 'Eros - Back SB', ha: 9.13, rows: '129-244', km: 30.43, cert: 'Organic' },
    { name: 'Eros - SB2020', ha: 2.514, rows: '105-128', km: 8.38, cert: 'Organic' },
    { name: 'Eros - TG2017', ha: 9.246, rows: '1-104', km: 30.82, cert: 'Organic' },
    { name: 'Eros - TG2015', ha: 2.91, rows: '65-101', km: 9.7, cert: 'Organic' },
    { name: 'Eros - TG2016', ha: 8.31, rows: '1-64 102-138', km: 27.7, cert: 'Organic' },
    { name: 'Woolshed - Sauvignon Blanc', ha: 10.99, rows: '1-108', km: 36.63, cert: 'Organic' },
    { name: 'Woolshed - Pinot Gris', ha: 1.83, rows: '1-37', km: 6.1, cert: 'Organic' },
    { name: 'Eros/Loveblock farm', ha: 0, rows: '', km: 0 },
    { name: 'N/A', ha: 0, rows: '', km: 0 },
    { name: 'SB 01', ha: 3.56, rows: '1-43', km: 0, cert: 'Conversion' },
    { name: 'SB 02', ha: 3.14, rows: '1-43', km: 0, cert: 'Conversion' },
    { name: 'SB 03', ha: 1.68, rows: '92-138', km: 0, cert: 'SWNZ' },
    { name: 'SB 04', ha: 5.328, rows: '43-115', km: 0, cert: 'SWNZ' },
    { name: 'SB 05', ha: 3.9, rows: '1-50', km: 0, cert: 'SWNZ' },
    { name: 'WB - CHA', ha: 0.83, rows: '29-91', km: 0, cert: 'SWNZ' },
    { name: 'WB - PG', ha: 1, rows: '51-91', km: 0, cert: 'SWNZ' },
    { name: 'WB - RSL', ha: 2.98, rows: '', km: 0, cert: 'SWNZ' },
    { name: 'WB - SYR', ha: 0.1, rows: '1-3', km: 0, cert: 'Conversion' },
  ],
  jobs: [
    { name: 'Canopy Control', code: '21301' },
    { name: 'Bud Rubbing', code: '21302' },
    { name: 'Frost Protection', code: '21303' },
    { name: 'Fruit Thinning & Dropping', code: '21304' },
    { name: 'Leaf Plucking', code: '21305' },
    { name: 'Replacement Plants - Training', code: '21307' },
    { name: 'Shoot Thinning', code: '21308' },
    { name: 'Skirting', code: '21309' },
    { name: 'Spraying Canopy', code: '21310' },
    { name: 'Vine Trimming & Mowing', code: '21311' },
    { name: 'Wire Lifting & Dropping', code: '21312' },
    { name: 'Ground Control', code: '21313' },
    { name: 'Applying Fertilzer', code: '21314' },
    { name: 'Mowing', code: '21315' },
    { name: 'Composting / Mulching', code: '21316' },
    { name: 'Soil Work', code: '21317' },
    { name: 'Undervine Control', code: '21317/01' },
    { name: 'Weed Spraying', code: '21318' },
    { name: 'Harvesting', code: '21319' },
    { name: 'Irrigation - R&M', code: '21322' },
    { name: 'Bird Control - Bird nets', code: '21327' },
    { name: 'Net Removal & Clipping', code: '21328' },
    { name: 'Weta Guards', code: '21330' },
    { name: 'Pruning', code: '21331' },
    { name: 'R&M Posts & Wires - Clipping Fixing', code: '21339' },
    { name: 'Staff training', code: '24804' },
    { name: 'Farm supervisor', code: '24808' },
    { name: 'Track maintenance', code: '21343' },
    { name: 'Young vines', code: '21307' },
    { name: 'Canopy Control:Retrunking', code: '21346' },
    { name: 'Washdown tractor/sprayer, farm vehicle R&M, attach sprayer etc.', code: '24605' },
    { name: 'Other - Specify on notes', code: '372' },
    { name: 'Sick leave', code: '' },
    { name: 'Annual leave', code: '' },
  ],
  // Jason and Simon (machinery operators) pick from this shorter list instead
  machineryTasks: [
    { name: 'Leaf Plucking', code: '21305' },
    { name: 'Spraying Canopy', code: '21310' },
    { name: 'Vine Trimming & Mowing', code: '21311' },
    { name: 'Ground Control', code: '21313' },
    { name: 'Applying Fertilzer', code: '21314' },
    { name: 'Mowing', code: '21315' },
    { name: 'Composting / Mulching', code: '21316' },
    { name: 'Soil Work', code: '21317' },
    { name: 'Undervine Control', code: '21317/01' },
    { name: 'Weed Spraying', code: '21318' },
    { name: 'Harvesting', code: '21319' },
    { name: 'Irrigation - R&M', code: '21322' },
    { name: 'Bird Control - Bird nets', code: '21327' },
    { name: 'Net Removal & Clipping', code: '21328' },
    { name: 'Pruning', code: '21331' },
    { name: 'R&M Posts & Wires - Clipping Fixing', code: '21339' },
    { name: 'Staff training', code: '24804' },
    { name: 'Track maintenance', code: '21343' },
    { name: 'Washdown tractor/sprayer, farm vehicle R&M, attach sprayer etc.', code: '24605' },
    { name: 'Other - Specify on notes', code: '372' },
    { name: 'Sick leave', code: '' },
    { name: 'Annual leave', code: '' },
  ],
  statuses: ['To Spray', 'Jason', 'Simon'],
  laneTanks: { Jason: 3000, Simon: 2000 },   // legacy single-board fields (kept for safety)
  waterRate: 400,
  sprayTypes: [
    { key: 'canopy', label: 'Canopy Spray', statuses: ['To Spray', 'Jason', 'Simon'], laneTanks: { Jason: 3000, Simon: 2000 }, waterRate: 300, roundMix: [{ product: 'Microthiol Disperss', per100: 1.333 }, { product: 'NZBioActive', per100: 0.73 }, { product: 'Artemis Opti', per100: 0.05 }], roundDeducted: false },
    { key: 'ground', label: 'Ground Spray', statuses: ['To Spray', '300 L', '2000 L'], laneTanks: { '300 L': 300, '2000 L': 2000 }, waterRate: 400, roundMix: [], roundDeducted: false },
    { key: 'weed', label: 'Weed Spray', statuses: ['To Spray', 'Weed sprayer'], laneTanks: { 'Weed sprayer': 1000 }, waterRate: 200, roundMix: [{ product: 'Roundup UltraMAX', per100: 1.25 }, { product: 'LI 700', per100: 0.2 }, { product: 'Shark', per100: 0.1 }], roundDeducted: false },
  ],
  fuelTanks: ['Loveblock farm', 'Eros', 'Winery', 'Mobile tank'],
  // fleet: 'machine' services on hours run, 'vehicle' gets a weekly check
  vehicles: [
    { name: 'Fendt', kind: 'machine', machineType: 'Tractor', serviceEveryHours: 250, lastServiceHours: 0, checkEveryDays: 7 },
    { name: 'Loader Fendt', kind: 'machine', machineType: 'Tractor', serviceEveryHours: 250, lastServiceHours: 0 },
    { name: 'John Deer Loader', kind: 'machine', machineType: 'Tractor', serviceEveryHours: 250, lastServiceHours: 0 },
    { name: 'Fendt 209P', kind: 'machine', machineType: 'Tractor', serviceEveryHours: 250, lastServiceHours: 0 },
    { name: 'NPP 894', kind: 'vehicle', machineType: 'Vehicle', serviceEveryHours: 0, lastServiceHours: 0 },
    { name: 'GRM 565', kind: 'vehicle', machineType: 'Vehicle', serviceEveryHours: 0, lastServiceHours: 0 },
    { name: 'Waterblaster', kind: 'machine', machineType: 'Equipment', serviceEveryHours: 100, lastServiceHours: 0 },
    { name: 'Pellenc harvester', kind: 'machine', machineType: 'Harvester', serviceEveryHours: 200, lastServiceHours: 0 },
    { name: 'Mower', kind: 'machine', machineType: 'Equipment', serviceEveryHours: 200, lastServiceHours: 0, hoursSource: 'tasks', startHours: 0 },
    { name: 'Mulcher', kind: 'machine', machineType: 'Equipment', serviceEveryHours: 200, lastServiceHours: 0, hoursSource: 'tasks', startHours: 0 },
    { name: 'Undervine Plough', kind: 'machine', machineType: 'Equipment', serviceEveryHours: 200, lastServiceHours: 0, hoursSource: 'tasks', startHours: 0 },
    { name: 'Undervine Rollhacker', kind: 'machine', machineType: 'Equipment', serviceEveryHours: 200, lastServiceHours: 0, hoursSource: 'tasks', startHours: 0 },
    { name: 'Power harrow/seeder', kind: 'machine', machineType: 'Equipment', serviceEveryHours: 200, lastServiceHours: 0, hoursSource: 'tasks', startHours: 0 },
  ],
  // what the operator ticks off on a machine check
  vineSpacing: 1.8,
  rowWidth: 2.7,
  workPace: {
    'Mulching': { type: 'kmh', value: 5, headland: 15 },
    'Mowing - 2nd pass - driving row': { type: 'kmh', value: 6, headland: 15 },
    'Mowing/Topping  - Cover crop rows': { type: 'kmh', value: 6, headland: 15 },
    'Trimming - First pass': { type: 'kmh', value: 4.5, headland: 15 },
    'Undervine Blade': { type: 'kmh', value: 3, headland: 20 },
    'French plough -  Every row - 1st pass': { type: 'kmh', value: 3, headland: 20 },
    'Prunning': { type: 'plants', value: 55, headland: 0 },
    'Wire lift work': { type: 'plants', value: 220, headland: 0 },
    'Bud Rub': { type: 'plants', value: 180, headland: 0 },
    'Shoot thin': { type: 'plants', value: 120, headland: 0 },
    'Hand leaf plucking': { type: 'plants', value: 90, headland: 0 },
  },
  checklist: [
    'Engine oil level', 'Coolant level', 'Hydraulic oil level', 'Fuel / water trap',
    'Air filter', 'Greasing done', 'Tyres & pressures', 'Lights & beacon',
    'Brakes', 'Leaks (oil, fuel, water)', 'Guards & PTO cover', 'Seatbelt & ROPS',
    'Mirrors & windscreen', 'Fire extinguisher', 'General cleanliness',
  ],
  // Lookups used when exporting the timesheet — editable in Setup.
  blockCodes: {
    'Eros - Back SB': 'EROS MSB F&B', 'Eros - Front PG': 'EROS MPG', 'Eros - Front SB': 'EROS MSB F&B',
    'Eros - SB2020': 'EROS MSB 20', 'Eros - TG2015': 'EROS MSB 15&16', 'Eros - TG2016': 'EROS MSB 15&16',
    'Eros - TG2017': 'EROS MSB 17',
    'Hill - A 23': '760/03 - WIP A23', 'Hill - C - CBl': '760/01 - WIP Block C', 'Hill - B - Muscat': '760/05 - WIP Block B',
    'Hill - E GEW': 'Hill GEW', 'Hill - E PG': 'Hill MPG', 'Hill - E SB': 'Hill MSB', 'Hill - F PG': 'Hill MPG',
    'Hill - F RSL': 'Hill RIE', 'Hill - F SB': 'Hill MSB', 'Hill - C18 - CH': 'Hill CHD', 'Hill - C18 - SB': 'Hill MSB',
    'Hill - A SB': 'Hill MSB',
    'Woolshed - Pinot Gris': 'Woolshed MPG', 'Woolshed - Sauvignon Blanc': 'Woolshed MSB',
    'SB 01': 'WB SB 01/02 - Conversion', 'SB 02': 'WB SB 01/02 - Conversion',
    'SB 03': 'WB SB 03/04/05', 'SB 04': 'WB SB 03/04/05', 'SB 05': 'WB SB 03/04/05',
    'WB - CHA': 'WB - CHA', 'WB - PG': 'WB - PG', 'WB - RSL': 'WB - RSL', 'WB - SYR': 'WB - SYR',
  },
  jobAccounts: {
    'Pruning': 'Canopy Control:Pruning',
    'Pre pruning - Barrel pruning': 'Canopy Control:Pruning',
    'Bud Rubbing': 'Canopy Control:Bud Rubbing',
    'Rootstock': 'Canopy Control:Bud Rubbing',
    'Frost Protection': 'Canopy Control:Frost Protection',
    'Fruit Thinning & Dropping': 'Canopy Control:Fruit Thinning & Dropping',
    'Colour thin': 'Canopy Control:Fruit Thinning & Dropping',
    'Second sets removal': 'Canopy Control:Fruit Thinning & Dropping',
    'Leaf Plucking': 'Canopy Control:Leaf Plucking',
    'Hand leaf plucking': 'Canopy Control:Leaf Plucking',
    'Machine Leaf plucking - ERO Combi': 'Canopy Control:Leaf Plucking',
    'Replacement Plants - Training': 'Canopy Control:Replacement Plants incl Trainig',
    'Young plants care': 'Canopy Control:Replacement Plants incl Trainig',
    'Shoot Thinning': 'Canopy Control:Shoot Thinning',
    'Shoot thin': 'Canopy Control:Shoot Thinning',
    'Skirting': 'Canopy Control:Skirting',
    'Spraying Canopy': 'Canopy Control:Spraying Canopy',
    'Vine Trimming & Mowing': 'Canopy Control:Vine Trimming & Mowing',
    'Wire Lifting & Dropping': 'Canopy Control:Wire Lifting & Dropping',
    'Wire lift work': 'Canopy Control:Wire Lifting & Dropping',
    'Machine shaking': 'Canopy Control:Mechanical Shaking',
    'Ground Control': 'Ground Control:Soil Work',
    'Applying Fertilzer': 'Ground Control:Fertilzer',
    'Mowing': 'Ground Control:Mulching/Mowing',
    'Mulching': 'Ground Control:Mulching/Mowing',
    'Composting / Mulching': 'Ground Control:Compost spreading',
    'Compost spreading': 'Ground Control:Compost spreading',
    'Soil Work': 'Ground Control:Soil Work',
    'Undervine Control': 'Ground Control: Undervine Control',
    'Hand weed': 'Ground Control: Undervine Control',
    'Undervine Blade': 'Ground Control: Undervine Control',
    'Weed Spraying': 'Ground Control:Weed Spraying',
    'Harvesting': 'Harvesting',
    'Irrigation - R&M': 'Irrigation:Repairs and Maintenance',
    'Irrigation maintenance': 'Irrigation:Repairs and Maintenance',
    'Irrigation Flush': 'Irrigation:Repairs and Maintenance',
    'Bird Control - Bird nets': 'Pest Control:Bird Control',
    'Birds Netting - On': 'Pest Control:Bird Control',
    'Birds Netting - Off': 'Pest Control:Bird Control',
    'Bird scaring': 'Pest Control:Bird Control',
    'Net Removal & Clipping': 'Pest Control:Bird Control',
    'Weta Guards': 'Pest Control:Weta Guards',
    'Weta guards on retrunk, replants and vines': 'Pest Control:Weta Guards',
    'Bee Netting on - Clipping': 'Pest Control:Other',
    'R&M Posts & Wires - Clipping Fixing': 'Vineyard Exps - Other:Repairs & Maintenance:Posts & Wires',
    'Stays - all blocks': 'Vineyard Exps - Other:Repairs & Maintenance:Posts & Wires',
    'Remove dead vines': 'Vineyard Exps - Other:Repairs & Maintenance:Other',
    'Track maintenance': 'Vineyard Exps - Track Maintenace',
    'Washdown tractor/sprayer, farm vehicle R&M, attach sprayer etc.': 'Vineyard Expenses:Vehicle & Machinery Costs:Tractor:Tractor R&M',
    'Staff training': 'Vineyard Expenses:Staff Expenses:Viti Tech',
    'Canopy Control': 'Canopy Control:Pruning',
    'Canopy Control:Retrunking': 'Canopy Control:Pruning',
    'Leaf and petiole sample flowering': 'Vineyard Exps - Other:General Vineyard Expenses',
    'Other - Specify on notes': 'Vineyard Exps - Other:General Vineyard Expenses',
  },
  // which machine racks up hours when a task is worked
  taskMachines: {
    'Mowing/Topping  - Cover crop rows': 'Mower',
    'Mowing - 2nd pass - driving row': 'Mower',
    'Mulching': 'Mulcher',
    'French plough -  Every 2nd row - Driving rows': 'Undervine Plough',
    'French plough -  Every row - 1st pass': 'Undervine Plough',
    'French plough - Every row - Autumn': 'Undervine Plough',
    'Undervine cultivating - French Plough - All rows': 'Undervine Plough',
    'V Frame - Post French plough - Autumn': 'Undervine Plough',
    'Undervine Blade': 'Undervine Rollhacker',
    'UVC - Rollhack - Winter pass': 'Undervine Rollhacker',
    'UVC - Rollhack - 2nd pass': 'Undervine Rollhacker',
    'UVC - Rollhack - 3rd pass': 'Undervine Rollhacker',
    'Sowing Cover Crop - Autumn - Every 2nd row': 'Power harrow/seeder',
    'Sowing Cover Crop - Spring - Every 2nd row': 'Power harrow/seeder',
    'Sowing Cover Crop - Summer - Every 10th row': 'Power harrow/seeder',
    'Summer  Cultivation - POWER HARROW - Vigour': 'Power harrow/seeder',
  },
  weather: { lat: -41.62, lon: 174.08, label: 'Awatere Valley', stationUrl: '' },
  workTasks: [
    'Prunning', 'Pre pruning - Barrel pruning', 'Wire lift work', 'Wire drop',
    '1st Wire Lifting', '2nd Wire Lifting', '3rd Wire lift', 'Young plants care',
    'Compost spreading', 'Remove dead vines', 'Stays - all blocks',
    'Irrigation maintenance', 'Irrigation Flush',
    'Weta guards on retrunk, replants and vines', 'Hand weed',
    'Bee Netting on - Clipping', 'Shoot thin', 'Bud Rub',
    'Leaf and petiole sample flowering', 'Hand leaf plucking', 'Fruit thinning',
    'Rootstock', 'Birds Netting - On', 'Birds Netting - Off', 'Bird scaring',
    'Colour thin', 'Second sets removal',
    'Machine Leaf plucking - ERO Combi', 'Machine Leaf plucking - ERO Combi - 2nd pass',
    'Machine shaking',
    'Sowing Cover Crop - Autumn - Every 2nd row',
    'Sowing Cover Crop - Spring - Every 2nd row',
    'Sowing Cover Crop - Summer - Every 10th row',
    'Deep riper', 'French plough - Every row - Autumn',
    'French plough -  Every row - 1st pass', 'French plough -  Every 2nd row - Driving rows',
    'V Frame - Post Plough - Covercrop rows', 'Crimping   - Cover crop rows',
    'V Frame - Post French plough - Autumn', 'V Frame - Post Plough - Cover crop rows',
    'V Frame - Post Plough - Driving rows', 'Mowing/Topping  - Cover crop rows',
    'Mowing - 2nd pass - driving row', 'Mulching',
    'Summer  Cultivation - POWER HARROW - Vigour', 'Trimming - First pass',
    'Undervine Blade', 'Undervine cultivating - French Plough - All rows',
    'UVC - Rollhack - Winter pass', 'UVC - Rollhack - 2nd pass', 'UVC - Rollhack - 3rd pass',
  ],
  dataVersion: 17,
  products: [
    { name: 'Microthiol Disperss', unit: 'Kg', category: 'Powdery Mildew', rateBasis: 'per100', actives: 'sulphur — elemental', rate: 1.333, stock: '', minStock: '', biogro: true, approved: true },
    { name: 'NZBioActive', unit: 'L', category: 'Nutrition', rateBasis: 'per100', actives: 'fertiliser', rate: 0.73, stock: '', minStock: '', biogro: true, approved: true },
    { name: 'Artemis Opti', unit: 'L', category: 'Spreader/Adjuvant', rateBasis: 'per100', actives: 'polyether modified polysiloxane', rate: 0.05, stock: '', minStock: '', approved: true },
    { name: 'Roundup UltraMAX', unit: 'L', category: '', rateBasis: 'per100', actives: '570 g/L glyphosate', rate: 1.25, stock: 200, minStock: 50, approved: true },
    { name: 'LI 700', unit: 'L', category: 'Spreader/Adjuvant', rateBasis: 'per100', actives: 'penetrant/acidifier', rate: 0.2, stock: 40, minStock: 10, approved: true },
    { name: 'Shark', unit: 'L', category: '', rateBasis: 'per100', actives: '240 g/L carfentrazone', rate: 0.1, stock: 20, minStock: 5, approved: true },
  ],
  roundMix: [
    { product: 'Roundup UltraMAX', per100: 1.25 },
    { product: 'LI 700', per100: 0.2 },
    { product: 'Shark', per100: 0.1 },
  ],
  roundDeducted: false,
  webhookUrl: '',
  notifyEmail: '',
};
const OLD_DEFAULT_STATUSES = ['To Spray', 'In Progress', 'Completed'];
const HAZARD_TYPES = ['Slip / Trip / Fall', 'Machinery', 'Vehicle', 'Chemical / Spray',
  'Electrical', 'Manual Handling', 'Environmental', 'Other'];
const SEVERITIES = ['Low', 'Medium', 'High', 'Critical'];
// Seeded from GrapeLink Operator Job Sheet #766447 (Foliar/Full Canopy, 300 L/ha)
const SHEET_CANOPY_CARDS = [
  { id: 'js766447a', status: 'To Spray', done: false, fields: { Block: 'SB 01 - Home Block North', 'Total area': '3.15 ha', 'Water rate': '300 L/ha', Rows: '1 to 43', 'Vine row m': '11667', Vineyard: 'Loveblock Winery', Method: 'Foliar/Full Canopy', 'Job #': '766447' } },
  { id: 'js766447b', status: 'To Spray', done: false, fields: { Block: 'SB 02 - Home Block South', 'Total area': '3.30 ha', 'Water rate': '300 L/ha', Rows: '1 to 43', 'Vine row m': '12222', Vineyard: 'Loveblock Winery', Method: 'Foliar/Full Canopy', 'Job #': '766447' } },
  { id: 'js766447c', status: 'To Spray', done: false, fields: { Block: 'SYRAH - Home Block', 'Total area': '0.10 ha', 'Water rate': '300 L/ha', Rows: '1 to 3', 'Vine row m': '370', Vineyard: 'Loveblock Winery', Method: 'Foliar/Full Canopy', 'Job #': '766447' } },
];

/* ---------- helpers ---------- */
/* Everything date-related runs on New Zealand time, whatever the device is set
   to. toISOString() would give UTC — which is yesterday for most of an NZ day. */
const NZ_TZ = 'Pacific/Auckland';
function nzParts(d = new Date()) {
  const f = new Intl.DateTimeFormat('en-NZ', {
    timeZone: NZ_TZ, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
  }).formatToParts(d).reduce((a, p) => { a[p.type] = p.value; return a; }, {});
  return { y: +f.year, m: +f.month, d: +f.day, hh: f.hour === '24' ? '00' : f.hour, mm: f.minute, ss: f.second };
}
// a Date whose local fields read as the NZ wall clock — safe for day/week maths
const nzNow = () => { const p = nzParts(); return new Date(p.y, p.m - 1, p.d, +p.hh, +p.mm, +p.ss); };
const pad2 = n => String(n).padStart(2, '0');
const todayStr = () => { const p = nzParts(); return `${p.y}-${pad2(p.m)}-${pad2(p.d)}`; };
const nowTime = () => { const p = nzParts(); return `${p.hh}:${p.mm}`; };
// "2 h 15 m" from a millisecond span
const fmtDuration = ms => {
  const mins = Math.max(0, Math.round(ms / 60000));
  const h = Math.floor(mins / 60), m = mins % 60;
  return h ? `${h} h${m ? ` ${m} m` : ''}` : `${m} m`;
};
/* A job can run over several days. Each Start…Stop is one session, and the
   total is the sum of them — so a block picked up again the next morning adds
   a fresh session rather than counting the night in between. */
function cardSessions(card) {
  if (Array.isArray(card.sessions) && card.sessions.length) return card.sessions;
  // older cards recorded a single start/finish
  if (card.startedTs) return [{ startTs: card.startedTs, startAt: card.startedAt, startTime: card.startedTime,
    endTs: card.doneTs || null, endAt: card.doneAt || '', endTime: card.doneTime || '', by: card.startedBy || '' }];
  return [];
}
const openSession = card => cardSessions(card).find(x => !x.endTs) || null;
function cardWorkedMs(card, includeOpen = true) {
  return cardSessions(card).reduce((sum, x) => {
    if (x.endTs) return sum + Math.max(0, x.endTs - x.startTs);
    return includeOpen ? sum + Math.max(0, Date.now() - x.startTs) : sum;
  }, 0);
}
// where a job stands: finished, someone on it now, started but stopped, or not begun
function cardState(c) {
  if (c.done) return 'done';
  if (openSession(c)) return 'live';
  return cardSessions(c).some(x => x.endTs) ? 'paused' : 'planned';
}
const cardWorkedHours = card => Math.round(cardWorkedMs(card, false) / 36000) / 100;

const nowTimeNZ = () => { const p = nzParts(); return `${p.hh}:${p.mm}`; };
const todayNZ = () => { const p = nzParts(); return `${pad2(p.d)}/${pad2(p.m)}/${p.y}`; };
// times every 15 minutes, "HH:MM"
const TIME_OPTIONS = Array.from({ length: 96 }, (_, i) => `${String(Math.floor(i / 4)).padStart(2, '0')}:${String((i % 4) * 15).padStart(2, '0')}`);
// 24h of 15-min slots, ordered to begin at a given hour (then wraps past midnight)
const timeSlotsFrom = startHour => Array.from({ length: 96 }, (_, i) => {
  const mins = (startHour * 60 + i * 15) % 1440;
  return `${String(Math.floor(mins / 60)).padStart(2, '0')}:${String(mins % 60).padStart(2, '0')}`;
});
const START_OPTIONS = timeSlotsFrom(3);   // 03:00 → … → 02:45
const FINISH_OPTIONS = timeSlotsFrom(7);   // 07:00 → … → 06:45
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
function calcHours(start, finish) {
  if (!start || !finish) return 0;
  const [sh, sm] = start.split(':').map(Number);
  const [fh, fm] = finish.split(':').map(Number);
  let mins = (fh * 60 + fm) - (sh * 60 + sm);
  if (mins < 0) mins += 1440;
  return Math.round((mins / 60) * 100) / 100;
}
/* Task list for a person. Operators flagged as machinery (Jason, Simon) get the
   shorter tractor list; everyone else gets the full set. */
function tasksFor(config, session) {
  const all = config.jobs || [];
  const ops = config.operators || [];
  const me = ops.find(o => o.code === (session && session.code)) ||
             ops.find(o => o.name === (session && session.name));
  const picked = me && Array.isArray(me.tasks) ? me.tasks : null;
  if (!picked || !picked.length) return all;          // no selection = sees everything
  const set = new Set(picked);
  return all.filter(j => set.has(j.name));
}

/* Leave detection.
   Leave is still paid (it shows in Hrs) but isn't a day on site, so it's
   excluded from the "days" count on the timesheet. Matches on the task name,
   and on the note when the task is the catch-all "Other". */
const LEAVE_RE = /\b(annual\s*leave|sick\s*leave|sick\s*day|bereavement|parental\s*leave|public\s*holiday|stat\s*day|holiday\s*pay|unpaid\s*leave|on\s*leave|leave)\b/i;
function isLeave(entry) {
  if (!entry) return false;
  const job = String(entry.job || '');
  if (LEAVE_RE.test(job)) return true;
  if (/^\s*other\b/i.test(job) && LEAVE_RE.test(String(entry.note || ''))) return true;
  return false;
}

/* Lunch break check.
   Looks at everything logged for one day and works out whether there's a gap
   of at least 30 minutes somewhere in the middle of the day. Only complains
   when the day actually runs across lunchtime or is a long one — a short
   morning shift doesn't need a break. */
const minsOf = t => { if (!t) return null; const [h, m] = String(t).split(':').map(Number); return (h * 60 + (m || 0)); };
function lunchCheck(entries) {
  const spans = entries
    .map(e => ({ s: minsOf(e.start), f: minsOf(e.finish) }))
    .filter(x => x.s != null && x.f != null)
    .map(x => ({ s: x.s, f: x.f < x.s ? x.f + 1440 : x.f }))
    .sort((a, b) => a.s - b.s);
  if (!spans.length) return { needed: false, ok: true, gap: 0 };

  const dayStart = spans[0].s, dayEnd = Math.max(...spans.map(x => x.f));
  const worked = spans.reduce((sum, x) => sum + (x.f - x.s), 0);
  // only expected once the day totals six hours or more
  const needed = worked >= 6 * 60;
  if (!needed) return { needed: false, ok: true, gap: 0, worked };

  // biggest gap between consecutive spans, anywhere from late morning to mid-afternoon
  let gap = 0, gapAt = null;
  let cursor = spans[0].f;
  for (let i = 1; i < spans.length; i++) {
    const g = spans[i].s - cursor;
    if (g > gap) { gap = g; gapAt = cursor; }
    cursor = Math.max(cursor, spans[i].f);
  }
  return { needed: true, ok: gap >= 30, gap, gapAt, worked };
}
const hhmm = m => `${String(Math.floor((m % 1440) / 60)).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;

function fmtDate(d) {
  if (!d) return '';
  const dt = new Date(d + 'T00:00:00');
  if (isNaN(dt)) return d;
  return dt.toLocaleDateString('en-NZ', { day: '2-digit', month: 'short', year: 'numeric' });
}
function mondayOf(date) {
  const d = new Date(date); const day = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - day);
  // format from local fields, not toISOString — that would shift the day back
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

// ---- spray mix + chemical usage helpers ----
const numOf = v => { const n = parseFloat(String(v ?? '').replace(/[^0-9.\-]/g, '')); return isNaN(n) ? 0 : n; };
const fmtNum = n => { const r = Math.round(n * 100) / 100; return Number.isInteger(r) ? String(r) : String(r); };
function cardBlockName(card) {
  const f = card.fields || {}; const keys = Object.keys(f);
  if ('Block' in f) return f.Block;
  const k = keys.find(k => k.toLowerCase().includes('block') || k.toLowerCase().includes('task'));
  return k ? f[k] : (keys.length ? f[keys[0]] : '');
}
function cardArea(card, config) {
  const f = card.fields || {};
  const k = Object.keys(f).find(k => /area|\bha\b|hectare/i.test(k));
  let a = k ? numOf(f[k]) : 0;
  if (!a) { const b = (config.blocks || []).find(b => b.name === cardBlockName(card)); if (b) a = numOf(b.ha); }
  return a;
}
function cardWater(card, config) {
  const f = card.fields || {};
  const k = Object.keys(f).find(k => /water/i.test(k));
  const w = k ? numOf(f[k]) : 0;
  return w || numOf(config.waterRate) || 0;
}
function cardRows(card, config) {
  const f = card.fields || {};
  const k = Object.keys(f).find(k => /^rows?$/i.test(k) || /row range/i.test(k));
  if (k && f[k]) return String(f[k]);
  const b = (config.blocks || []).find(b => b.name === cardBlockName(card));
  return b && b.rows ? String(b.rows) : '';
}
const tankFor = (config, lane) => numOf((config.laneTanks || {})[lane]);
const productUnit = (config, name) => { const p = (config.products || []).find(p => p.name === name); return p ? (p.unit || '') : ''; };
// total of each product used across the given (done) cards, by label rate × volume
/* Rates are either per 100 L of spray or per hectare. Per-100 L products scale
   with the water rate; per-hectare products don't. */
function productBasis(config, name) {
  const p = ((config || {}).products || []).find(x => x.name === name);
  return (p && p.rateBasis) === 'perHa' ? 'perHa' : 'per100';
}
const mixRate = m => numOf(m.rate !== undefined && m.rate !== '' ? m.rate : m.per100);
function amountForVolume(config, m, volumeL, waterRate) {
  const rate = mixRate(m);
  if (productBasis(config, m.product) === 'perHa') {
    const w = numOf(waterRate);
    return w > 0 ? rate * (volumeL / w) : 0;      // litres ÷ L/ha = hectares
  }
  return rate * volumeL / 100;
}
function ratePer100(config, m, waterRate) {
  const rate = mixRate(m);
  if (productBasis(config, m.product) === 'perHa') {
    const w = numOf(waterRate);
    return w > 0 ? rate * 100 / w : 0;
  }
  return rate;
}

function roundUsage(cards, config) {
  const mix = config.roundMix || [];
  const used = {}; mix.forEach(m => { used[m.product] = 0; });
  cards.forEach(c => {
    const water = cardWater(c, config);
    const vol = cardArea(c, config) * water;
    mix.forEach(m => { used[m.product] += amountForVolume(config, m, vol, water); });
  });
  return used;
}
function areaProgress(cards, config) {
  let done = 0, total = 0;
  cards.forEach(c => { const a = cardArea(c, config); total += a; if (c.done) done += a; });
  return { done, total, pct: total > 0 ? Math.round((done / total) * 100) : 0 };
}

// ---- spreadsheet import helpers ----
const NOISE_COLS = ['grapelink', 'modified by', 'modified', 'duration calc', 'predecessors',
  'sprint', 'variance', 'clone', 'rootstock', 'id', 'row id', 'rowid', 'sheet', 'created by', 'created'];
const isTruthy = v => ['true', 'yes', 'y', '1', 'x', 'done', '✓', 'complete', 'completed', 'sprayed'].includes(String(v).trim().toLowerCase());
function cellToStr(v) {
  if (v == null) return '';
  if (v instanceof Date) {
    const dd = String(v.getDate()).padStart(2, '0'), mm = String(v.getMonth() + 1).padStart(2, '0');
    return `${dd}/${mm}/${v.getFullYear()}`;
  }
  if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(Math.round(v * 100) / 100);
  return String(v).replace(/\s+/g, ' ').trim();
}
// Excel stores dates as a serial number of days since 1899-12-30.
function serialToDate(n) {
  const d = new Date(Math.round((n - 25569) * 86400 * 1000));
  const dd = String(d.getUTCDate()).padStart(2, '0'), mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  return `${dd}/${mm}/${d.getUTCFullYear()}`;
}
// Turn an array-of-arrays (any leading title rows tolerated) into row objects keyed by header.
function aoaToRows(aoa) {
  if (!aoa || !aoa.length) return [];
  let headerIdx = 0, best = -1;
  for (let i = 0; i < Math.min(aoa.length, 15); i++) {
    const count = (aoa[i] || []).filter(c => c != null && String(c).trim() !== '').length;
    if (count > best && count >= 2) { best = count; headerIdx = i; }
  }
  const seen = {};
  const headers = (aoa[headerIdx] || []).map((c, i) => {
    let h = (c != null && String(c).trim() !== '') ? String(c).trim() : `Column ${i + 1}`;
    if (seen[h]) { seen[h]++; h = `${h} (${seen[h]})`; } else seen[h] = 1; // keep duplicate columns distinct
    return h;
  });
  const rows = [];
  for (let i = headerIdx + 1; i < aoa.length; i++) {
    const arr = aoa[i]; if (!arr) continue;
    if (!arr.some(c => c != null && String(c).trim() !== '')) continue;
    const o = {}; headers.forEach((h, j) => { o[h] = arr[j] != null ? arr[j] : ''; });
    rows.push(o);
  }
  return rows;
}

/* ---------- shared style tokens ---------- */
const CREAM = '#F4F1EA';
const cls = {
  input: 'w-full px-3.5 py-3 rounded-lg border border-stone-300 bg-white text-stone-900 text-[15px] focus:outline-none focus:ring-2 focus:ring-stone-500/40 focus:border-stone-500',
  label: 'block text-[11px] font-semibold uppercase tracking-[0.12em] text-stone-500 mb-1.5',
  primary: 'inline-flex items-center justify-center gap-2 px-4 py-3 rounded-lg bg-stone-900 text-stone-50 text-[15px] font-medium hover:bg-stone-800 active:bg-stone-700 transition-colors disabled:opacity-40',
  ghost: 'inline-flex items-center justify-center gap-2 px-4 py-3 rounded-lg bg-white border border-stone-300 text-stone-800 text-[15px] font-medium hover:bg-stone-50 transition-colors',
  card: 'bg-white border border-stone-200 rounded-xl',
};
const serif = { fontFamily: 'Georgia, "Times New Roman", serif' };

/* ============================================================
   Combobox (type to filter)
   ============================================================ */
function Combobox({ label, options, value, onChange, placeholder, icon: Icon }) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState('');
  const ref = useRef(null);
  useEffect(() => {
    const onDoc = e => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', onDoc);
    return () => document.removeEventListener('mousedown', onDoc);
  }, []);
  const filtered = options.filter(o => o.toLowerCase().includes(q.toLowerCase()));
  return (
    <div ref={ref} className="relative">
      {label && <label className={cls.label}>{label}</label>}
      <div className="relative">
        {Icon && <Icon size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />}
        <input
          className={cls.input + (Icon ? ' pl-9' : '')}
          value={open ? q : value}
          placeholder={value || placeholder || 'Type to search…'}
          onFocus={() => { setOpen(true); setQ(''); }}
          onChange={e => { setQ(e.target.value); setOpen(true); }}
        />
      </div>
      {open && (
        <div className="absolute z-30 mt-1 w-full max-h-56 overflow-auto bg-white border border-stone-300 rounded-lg shadow-lg">
          {filtered.length ? filtered.map(o => (
            <button key={o}
              onMouseDown={e => { e.preventDefault(); onChange(o); setOpen(false); setQ(''); }}
              className={'w-full text-left px-3.5 py-2.5 text-[15px] hover:bg-stone-100 ' + (o === value ? 'bg-stone-50 font-medium' : '')}>
              {o}
            </button>
          )) : <div className="px-3.5 py-3 text-stone-400 text-sm">No matches</div>}
        </div>
      )}
    </div>
  );
}

/* ============================================================
   Auth — numeric keypad
   ============================================================ */
function AuthScreen({ config, onSubmit }) {
  const [code, setCode] = useState('');
  const [err, setErr] = useState('');
  const press = d => { setErr(''); setCode(c => (c.length < 8 ? c + d : c)); };
  const submit = () => {
    const ok = onSubmit(code);
    if (!ok) { setErr('Code not recognised'); setCode(''); }
  };
  const keys = ['1', '2', '3', '4', '5', '6', '7', '8', '9', 'C', '0', '⌫'];
  return (
    <div className="min-h-screen flex flex-col items-center justify-center px-6" style={{ backgroundColor: CREAM }}>
      <img src={logoUrl} alt={config.siteName} className="w-64 max-w-[70vw] mb-8 select-none" draggable="false" />
      <p className="text-[11px] font-semibold uppercase tracking-[0.18em] text-stone-500 mb-2">Enter your code</p>
      <div className="h-12 flex items-center justify-center mb-5">
        <span style={serif} className="text-3xl tracking-[0.3em] text-stone-900">{code ? code.replace(/./g, '•') : '—'}</span>
      </div>
      {err && <p className="text-red-700 text-sm mb-3">{err}</p>}
      <div className="grid grid-cols-3 gap-2.5 w-full max-w-[280px]">
        {keys.map(k => (
          <button key={k}
            onClick={() => k === 'C' ? (setCode(''), setErr('')) : k === '⌫' ? setCode(c => c.slice(0, -1)) : press(k)}
            className="h-16 rounded-xl bg-white border border-stone-300 text-stone-900 text-2xl font-light hover:bg-stone-100 active:bg-stone-200 transition-colors">
            {k}
          </button>
        ))}
      </div>
      <button onClick={submit} disabled={!code}
        className={cls.primary + ' w-full max-w-[280px] mt-4 py-4 text-base'}>
        Continue <ChevronRight size={18} />
      </button>
      <p className="mt-6 text-xs text-stone-400">Ask your manager for your code.</p>
    </div>
  );
}

/* ============================================================
   Top bar
   ============================================================ */
function TopBar({ siteName, subtitle, onBack, onLogout, wide }) {
  return (
    <header className="sticky top-0 z-20 border-b border-stone-300" style={{ backgroundColor: CREAM }}>
      <div className={(wide ? 'max-w-[1800px]' : 'max-w-5xl') + ' mx-auto px-4 sm:px-6 h-16 flex items-center justify-between'}>
        <div className="flex items-center gap-3 min-w-0">
          {onBack && (
            <button onClick={onBack} className="p-2 -ml-2 rounded-lg hover:bg-stone-200/60 text-stone-700">
              <ChevronRight size={20} className="rotate-180" />
            </button>
          )}
          <img src={logoUrl} alt={siteName} className="h-9 w-auto shrink-0 select-none" draggable="false" />
          {subtitle && <span className="text-sm text-stone-500 truncate hidden sm:block">{subtitle}</span>}
        </div>
        <button onClick={onLogout} className="inline-flex items-center gap-1.5 text-sm text-stone-600 hover:text-stone-900 px-2 py-1">
          <LogOut size={16} /> <span className="hidden sm:inline">Sign out</span>
        </button>
      </div>
    </header>
  );
}

function Banner({ msg }) {
  if (!msg) return null;
  return (
    <div className="flex items-center gap-2 px-3.5 py-2.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 text-sm mb-4">
      <Check size={16} /> {msg}
    </div>
  );
}

/* ============================================================
   Spray board (kanban) — shared by operator & manager
   ============================================================ */
function RoundPanel({ tc, sprays, patchType, onApplyWater }) {
  const [open, setOpen] = useState(false);
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [showRange, setShowRange] = useState(false);
  // cards carry their actual date as dd/mm/yyyy — compare on ISO
  const isoOfCard = c => {
    const v = (c.fields && (c.fields['Actual date'] || c.fields['Planned date'])) || c.doneAt || '';
    const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : '';
  };
  const inRange = c => {
    if (!from && !to) return true;
    const d = isoOfCard(c);
    if (!d) return false;                       // undated cards drop out of a filtered export
    return (!from || d >= from) && (!to || d <= to);
  };
  const prog = areaProgress(sprays || [], tc);
  const products = tc.products || [];
  const mix = tc.roundMix || [];
  const laneTanks = tc.laneTanks || {};
  const setMix = next => patchType({ roundMix: next });

  const exportRound = () => {
    const m = tc.roundMix || [];
    const inScope = (sprays || []).filter(inRange);
    const rows = inScope.map(c => {
      const lane = c.status, area = cardArea(c, tc), water = cardWater(c, tc), vol = area * water;
      const row = {
        Block: cardBlockName(c), Operator: lane, 'Tank (L)': tankFor(tc, lane) || '',
        'Area (ha)': area, 'Water (L/ha)': water, 'Volume (L)': Math.round(vol),
        Done: c.done ? 'Yes' : 'No',
        'Actual date': (c.fields && c.fields['Actual date']) || c.doneAt || '',
        'Actual time': (c.fields && c.fields['Actual time']) || c.doneTime || '',
        'Completed by': c.doneBy || '',
      };
      m.forEach(x => { row[`${x.product} (${productUnit(tc, x.product)})`] = Math.round(amountForVolume(tc, x, vol, water) * 100) / 100; });
      return row;
    });
    const used = roundUsage(inScope.filter(c => c.done), tc);
    const usage = products.filter(p => used[p.name] != null).map(p => ({
      Product: p.name, Unit: p.unit || '',
      BioGro: p.biogro ? 'Certified' : '', Approved: p.approved === false ? 'NOT APPROVED' : 'Yes', 'Used this round': Math.round((used[p.name] || 0) * 100) / 100,
      'Opening stock': numOf(p.stock), Remaining: Math.round((numOf(p.stock) - (used[p.name] || 0)) * 100) / 100,
    }));
    const lanes = Object.keys(laneTanks);
    const mixSheet = m.map(x => {
      const row = {
        Product: x.product, Rate: mixRate(x),
        Basis: productBasis(tc, x.product) === 'perHa' ? 'per hectare' : 'per 100 L',
        Unit: productUnit(tc, x.product),
        'Per 100 L': Math.round(ratePer100(tc, x, tc.waterRate) * 1000) / 1000,
      };
      lanes.forEach(l => { row[`${l} · ${numOf(laneTanks[l])} L`] = Math.round(amountForVolume(tc, x, numOf(laneTanks[l]), tc.waterRate) * 100) / 100; });
      return row;
    });
    const wb = XLSX.utils.book_new();
    addSheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{}]), 'Spray round');
    addSheet(wb, XLSX.utils.json_to_sheet(usage.length ? usage : [{}]), 'Product usage');
    addSheet(wb, XLSX.utils.json_to_sheet(mixSheet.length ? mixSheet : [{}]), 'Mix');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = `${tc.label || 'spray'}-round_${todayStr()}.xlsx`.replace(/\s+/g, '-').toLowerCase(); document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  return (
    <div className={cls.card + ' p-4 mb-5'}>
      <div className="flex items-center justify-between gap-3 mb-2">
        <span className="text-[11px] font-semibold uppercase tracking-[0.14em] text-stone-500">Round progress · by area</span>
        <span className="text-sm text-stone-600">{fmtNum(prog.done)} / {fmtNum(prog.total)} ha</span>
      </div>
      <div className="flex items-center gap-3">
        <div className="flex-1 h-3 rounded-full bg-stone-100 overflow-hidden">
          <div className="h-full rounded-full transition-all" style={{ width: prog.pct + '%', backgroundColor: prog.pct === 100 ? '#059669' : '#57534e' }} />
        </div>
        <span style={serif} className="text-xl font-bold text-stone-900 tabular-nums w-12 text-right">{prog.pct}%</span>
      </div>
      {prog.pct === 100 && prog.total > 0 && (
        <div className="mt-2 text-sm text-emerald-700 flex items-center gap-1.5"><Check size={15} /> Round complete — export below, then deduct stock in the Shed tab.</div>
      )}

      <div className="flex gap-2 mt-3 flex-wrap">
        <button onClick={() => setOpen(v => !v)} className={cls.ghost + ' !py-2 !px-3'}><Layers size={15} /> Round mix ({mix.length})</button>
        <button onClick={() => setShowRange(v => !v)} className={cls.ghost + ' !py-2 !px-3'}>
          {from || to ? `Dates: ${from || 'start'} → ${to || 'today'}` : 'All dates'}
        </button>
        <button onClick={exportRound} className={cls.primary + ' !py-2 !px-3'}><Download size={15} /> Export round</button>
      </div>

      {showRange && (
        <div className="mt-3 pt-3 border-t border-stone-200">
          <p className="text-sm text-stone-500 mb-2">Limit the export to blocks sprayed in a date range. Blocks with no date are left out when a range is set.</p>
          <RangePicker from={from} to={to} onChange={(f, t) => { setFrom(f); setTo(t); }} compact />
        </div>
      )}

      {open && (
        <div className="mt-3 pt-3 border-t border-stone-200">
          <p className="text-sm text-stone-500 mb-2">This mix applies to every block in the <b>{tc.label}</b> round. Full-tank amounts scale to each sprayer automatically.</p>

          <div className="flex items-end gap-2 mb-3 pb-3 border-b border-stone-100 flex-wrap">
            <div>
              <label className={cls.label}>Water rate for this round</label>
              <div className="flex items-center gap-2">
                <input value={tc.waterRate ?? ''} inputMode="decimal"
                  onChange={e => patchType({ waterRate: e.target.value === '' ? '' : numOf(e.target.value) })}
                  className={cls.input + ' !w-28 text-right text-lg'} />
                <span className="text-sm text-stone-500">L/ha</span>
              </div>
            </div>
            <div className="text-[13px] text-stone-500 pb-2.5">
              {(() => {
                const prog = areaProgress(sprays || [], tc);
                const vol = Math.round(numOf(tc.waterRate) * prog.total);
                return prog.total > 0
                  ? <>{fmtNum(prog.total)} ha on this board · about <b className="text-stone-800">{fmtNum(vol)} L</b> of water for the round</>
                  : <>Sets the water volume used to work out product quantities.</>;
              })()}
            </div>
            {(() => {
              // blocks imported from a job sheet carry their own rate, which wins
              const own = (sprays || []).filter(c => {
                const f = c.fields || {}; const k = Object.keys(f).find(x => /water/i.test(x));
                return k && numOf(f[k]) > 0 && numOf(f[k]) !== numOf(tc.waterRate);
              });
              if (!own.length) return null;
              return (
                <button onClick={() => onApplyWater && onApplyWater(numOf(tc.waterRate))}
                  className={cls.ghost + ' !py-2 !px-3 pb-0'}>
                  Apply to {own.length} block{own.length > 1 ? 's' : ''} with their own rate
                </button>
              );
            })()}
          </div>
          <p className="text-xs text-stone-400 -mt-2 mb-3">
            Changing this rescales every block's usage and the round totals. Full-tank and per-100 L mixes are unaffected.
            Blocks loaded from a job sheet keep the rate on the sheet until you apply this one to them.
          </p>

          {products.length === 0 && <p className="text-sm text-amber-700 mb-2">Add products in the Shed tab first.</p>}
          {(() => {
            const offending = nonOrganicInMix(mix, products);
            if (!offending.length) return null;
            const blocked = (sprays || []).filter(c => certRestricted(certOf(cardBlockName(c), tc)));
            if (!blocked.length) return null;
            return (
              <div className="flex items-start gap-2 rounded-lg border border-red-300 bg-red-50 px-3 py-2.5 mb-3">
                <AlertTriangle size={17} className="text-red-600 shrink-0 mt-0.5" />
                <div className="text-[13.5px] text-red-900 leading-snug">
                  <b>This mix can't go on {blocked.length} organic / in-conversion block{blocked.length > 1 ? 's' : ''}.</b>
                  <div className="mt-1">{offending.map(p => p.name).join(', ')} {offending.length > 1 ? 'are' : 'is'} not BioGro certified and approved.</div>
                  <div className="text-[12px] text-red-800/80 mt-1">{blocked.map(c => cardBlockName(c)).join(', ')}</div>
                </div>
              </div>
            );
          })()}
          {(() => {
            const bad = mix.map(x => products.find(p => p.name === x.product)).filter(p => p && p.approved === false);
            if (!bad.length) return null;
            return (
              <div className="flex items-start gap-2 rounded-lg border border-red-300 bg-red-50 px-3 py-2.5 mb-3">
                <AlertTriangle size={17} className="text-red-600 shrink-0 mt-0.5" />
                <div className="text-[13.5px] text-red-900 leading-snug">
                  <b>{bad.map(p => p.name).join(', ')}</b> {bad.length > 1 ? 'are' : 'is'} not approved for use.
                  Check before spraying, or mark {bad.length > 1 ? 'them' : 'it'} approved in the Shed.
                </div>
              </div>
            );
          })()}
          <div className="space-y-2">
            {mix.map((x, i) => (
              <div key={i} className="flex items-center gap-2">
                <select value={x.product} onChange={e => setMix(mix.map((y, j) => j === i ? { ...y, product: e.target.value } : y))} className={cls.input}>
                  {!products.some(p => p.name === x.product) && <option value={x.product}>{x.product}</option>}
                  {products.map(p => <option key={p.name} value={p.name}>{p.name}</option>)}
                </select>
                <input value={x.rate !== undefined ? x.rate : x.per100} onChange={e => setMix(mix.map((y, j) => j === i ? { ...y, rate: e.target.value, per100: undefined } : y))}
                  inputMode="decimal" className={cls.input + ' !w-24 text-right'} />
                <span className="text-sm text-stone-500 w-24">
                  {productUnit(tc, x.product)}{productBasis(tc, x.product) === 'perHa' ? '/ha' : '/100L'}
                </span>
                {(() => {
                  const p = (tc.products || []).find(q => q.name === x.product);
                  if (!p) return null;
                  return (
                    <span className="flex gap-1 shrink-0">
                      {p.biogro && <span className="text-[10px] font-semibold uppercase tracking-wide text-emerald-800 bg-emerald-50 border border-emerald-200 rounded px-1.5 py-0.5">BioGro</span>}
                      {p.approved === false && <span className="text-[10px] font-semibold uppercase tracking-wide text-red-700 bg-red-50 border border-red-200 rounded px-1.5 py-0.5">Not approved</span>}
                    </span>
                  );
                })()}
                <button onClick={() => setMix(mix.filter((_, j) => j !== i))} className="p-2 rounded-lg hover:bg-red-50 text-red-500 shrink-0"><Trash2 size={15} /></button>
              </div>
            ))}
          </div>
          <button onClick={() => setMix([...mix, { product: products[0]?.name || '', rate: '' }])} className={cls.ghost + ' !py-2 !px-3 mt-2'}><Plus size={15} /> Add product to mix</button>
        </div>
      )}
    </div>
  );
}

function SprayHub({ config, setConfig, manager, operatorName }) {
  const [active, setActive] = useState(null);
  const types = config.sprayTypes || [];
  const type = types.find(t => t.key === active);
  if (active && type) return <SprayBoard config={config} setConfig={setConfig} manager={manager} type={type} typeKey={active} onBack={() => setActive(null)} operatorName={operatorName} />;
  return (
    <div>
      <div className="flex items-center gap-2 text-stone-700 mb-4"><Droplets size={18} /><h2 className="text-lg font-semibold text-stone-900">Spray</h2></div>
      <div className="grid gap-3 sm:grid-cols-3">
        {types.map(t => {
          const lanes = Object.entries(t.laneTanks || {}).map(([l, v]) => `${l} · ${numOf(v)} L`);
          return (
            <button key={t.key} onClick={() => setActive(t.key)}
              className="flex flex-col items-center text-center rounded-2xl border border-stone-200 bg-white p-6 hover:border-stone-400 hover:shadow-md transition-all active:scale-[0.99]">
              <div className="w-20 h-20 rounded-2xl bg-stone-900 text-stone-50 flex items-center justify-center mb-4"><Droplets size={40} /></div>
              <div className="font-semibold text-stone-900 text-[16px]">{t.label}</div>
              <div className="text-[13px] text-stone-500 mt-1.5 space-y-0.5">
                {lanes.length ? lanes.map(l => <div key={l}>{l}</div>) : <div>No sprayers set</div>}
              </div>
            </button>
          );
        })}
      </div>
    </div>
  );
}

/* Build a spray round by hand: pick the blocks, who sprays them and when */
function SprayRoundBuilder({ config, tc, statuses, onAdd, onClose }) {
  const [picked, setPicked] = useState([]);
  const [lane, setLane] = useState(statuses[1] || statuses[0] || 'To Spray');
  const [date, setDate] = useState(todayStr());
  const blocks = config.blocks || [];
  const toggle = n => setPicked(picked.includes(n) ? picked.filter(x => x !== n) : [...picked, n]);
  const totalHa = picked.reduce((s, n) => s + numOf((blocks.find(b => b.name === n) || {}).ha), 0);
  const water = numOf(tc.waterRate) || 0;

  const add = () => {
    if (!picked.length) return;
    const cards = picked.map(name => {
      const b = blocks.find(x => x.name === name) || {};
      return {
        id: uid(), status: lane, done: false,
        fields: {
          Block: name,
          'Total area': `${fmtNum(b.ha)} ha`,
          'Water rate': `${fmtNum(water)} L/ha`,
          Rows: b.rows || '',
          'Planned date': date ? date.split('-').reverse().join('/') : '',
        },
      };
    });
    onAdd(cards);
    setPicked([]);
  };

  const groups = {};
  blocks.forEach(b => { const v = vineyardOf(b.name); (groups[v] = groups[v] || []).push(b); });
  const order = [...VINEYARDS, ...Object.keys(groups).filter(k => !VINEYARDS.includes(k)).sort()].filter(v => (groups[v] || []).length);

  return (
    <div className={cls.card + ' p-4 mb-5'}>
      <div className="flex items-center justify-between gap-2 mb-3">
        <h3 className="font-semibold text-stone-900">Add blocks to this round</h3>
        <button onClick={onClose} className="p-1.5 rounded-lg hover:bg-stone-100 text-stone-500"><X size={16} /></button>
      </div>
      <p className="text-sm text-stone-500 mb-3">
        Set the mix under <b>Round mix</b> first — it applies to every block on this board. Then pick the blocks, who's spraying and when.
      </p>

      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4 mb-3">
        {order.map(v => {
          const list = groups[v];
          const allPicked = list.every(b => picked.includes(b.name));
          return (
            <div key={v} className="rounded-xl border border-stone-200 bg-stone-50/60 p-2.5">
              <div className="flex items-center justify-between gap-2 mb-2 px-0.5">
                <span className="text-[13px] font-bold text-stone-900">{v}</span>
                <button onClick={() => setPicked(allPicked
                  ? picked.filter(n => !list.some(b => b.name === n))
                  : [...new Set([...picked, ...list.map(b => b.name)])])}
                  className="text-[11px] px-2 py-1 rounded-md border border-stone-300 bg-white text-stone-600 hover:bg-stone-100">
                  {allPicked ? 'None' : 'All'}
                </button>
              </div>
              <div className="space-y-1.5">
                {list.map(b => (
                  <button key={b.name} onClick={() => toggle(b.name)}
                    className={'w-full text-left px-3 py-2 rounded-lg border text-sm transition-colors ' +
                      (picked.includes(b.name) ? 'bg-sky-100 border-sky-300 text-sky-900 font-medium' : 'bg-white border-stone-300 text-stone-700 hover:border-stone-400')}>
                    <span className="truncate block">{b.name}{numOf(b.ha) > 0 && <span className="opacity-60"> · {fmtNum(b.ha)} ha</span>}</span>
                  </button>
                ))}
              </div>
            </div>
          );
        })}
      </div>

      <div className="flex gap-3 flex-wrap items-end">
        <div><label className={cls.label}>Sprayer</label>
          <select value={lane} onChange={e => setLane(e.target.value)} className={cls.input + ' !w-auto'}>
            {statuses.map(st => <option key={st} value={st}>{st}</option>)}
          </select></div>
        <div><label className={cls.label}>Planned date</label>
          <input type="date" value={date} onChange={e => setDate(e.target.value)} className={cls.input + ' !w-auto'} /></div>
        <div className="text-sm text-stone-500 pb-2.5">
          {picked.length ? `${picked.length} block${picked.length > 1 ? 's' : ''} · ${fmtNum(Math.round(totalHa * 100) / 100)} ha · ${fmtNum(Math.round(totalHa * water))} L of water` : 'No blocks picked yet'}
        </div>
        <button onClick={add} disabled={!picked.length} className={cls.primary + ' !py-2.5 ml-auto'}>
          <Plus size={16} /> Add {picked.length ? `${picked.length} block${picked.length > 1 ? 's' : ''}` : 'blocks'}
        </button>
      </div>
    </div>
  );
}

function SprayBoard({ config, manager, setConfig, type, typeKey, onBack, operatorName }) {
  const tc = useMemo(() => ({
    ...config,
    statuses: type.statuses, laneTanks: type.laneTanks,
    roundMix: type.roundMix, waterRate: type.waterRate, roundDeducted: type.roundDeducted,
  }), [config, type]);
  const SK = K.sprays(typeKey);
  const patchType = patch => {
    const next = (config.sprayTypes || []).map(t => (t.key === typeKey ? { ...t, ...patch } : t));
    setConfig({ ...config, sprayTypes: next });
  };
  const [sprays, setSprays] = useState(null);
  const [showLoader, setShowLoader] = useState(false);
  const [showBuilder, setShowBuilder] = useState(false);
  const [editing, setEditing] = useState(null);
  const [view, setView] = useState('kanban');    // manager can switch to 'grid'
  const spraysRef = useRef([]);

  const load = async () => { const data = await loadJSON(SK, []); spraysRef.current = data; setSprays(data); };
  // another device ticking a block Done reflects here immediately
  useLiveKey(SK, v => { if (v) { spraysRef.current = v; setSprays(v); } }, [SK]);
  useEffect(() => { load(); }, []);
  useEffect(() => { spraysRef.current = sprays || []; }, [sprays]);

  const persist = async next => { setSprays(next); await saveJSON(SK, next); };
  const move = async (id, status) => {
    const fresh = await loadJSON(SK, sprays || []);
    await persist(fresh.map(c => (c.id === id ? { ...c, status } : c)));
  };
  const remove = async id => { await persist((sprays || []).filter(c => c.id !== id)); };
  // move a card one lane left/right — reliable on any device
  const shiftLane = async (card, dir) => {
    const from = statuses.indexOf(card.status);
    const to = Math.min(Math.max((from < 0 ? 0 : from) + dir, 0), statuses.length - 1);
    if (to === from) return;
    await persist((sprays || []).map(c => (c.id === card.id ? { ...c, status: statuses[to] } : c)));
  };

  const statuses = type.statuses;
  const grouped = useMemo(() => {
    const g = {}; statuses.forEach(s => (g[s] = []));
    (sprays || []).forEach(c => { (g[c.status] = g[c.status] || []).push(c); });
    Object.keys(g).forEach(s => {
      if (!statuses.includes(s) && g[s].length === 0) delete g[s];
      else g[s].sort((a, b) => (a.done ? 1 : 0) - (b.done ? 1 : 0)); // done cards sink to the bottom
    });
    return g;
  }, [sprays, statuses]);

  const toggleDone = async card => {
    const fresh = await loadJSON(SK, sprays || []);
    await persist(fresh.map(c => {
      if (c.id !== card.id) return c;
      const done = !c.done;
      const fields = { ...c.fields };
      if (done) { fields['Actual date'] = todayNZ(); fields['Actual time'] = nowTimeNZ(); }
      else { delete fields['Actual date']; delete fields['Actual time']; }
      return done
        ? { ...c, done, fields, doneAt: todayNZ(), doneTime: nowTimeNZ(), doneTs: Date.now(), doneBy: operatorName || 'Manager' }
        : { ...c, done, fields, doneAt: '', doneTime: '', doneTs: null, doneBy: '' };
    }));
  };

  // grid editing
  const columns = useMemo(() => {
    const set = [];
    (sprays || []).forEach(c => Object.keys(c.fields || {}).forEach(k => { if (!set.includes(k)) set.push(k); }));
    return set;
  }, [sprays]);

  // Drag: shared with the work board — grab anywhere on a card, move between
  // lanes and reorder. Mouse drags on movement; touch needs a short press-hold.
  const { drag, overCol, overIndex, boardRef, colRefs, listRefs, startPointer } = useKanbanDrag({
    items: sprays || [], lanes: statuses, laneKey: 'status',
    persist: next => { spraysRef.current = next; persist(next); },
  });

  // on an operator's phone, open with their own lane centred (they can still
  // swipe to To Spray and the others)
  const mineRef = useRef(null);
  const centred = useRef(false);
  useEffect(() => {
    if (manager || centred.current) return;
    let tries = 0;
    const go = () => {
      if (centred.current) return;
      if (centreLane(boardRef.current, mineRef.current)) { centred.current = true; return; }
      if (tries++ < 12) requestAnimationFrame(go);
    };
    requestAnimationFrame(go);
  }, [manager, operatorName, (sprays || []).length]);
  // switching to another spray board should centre again
  useEffect(() => { centred.current = false; }, [typeKey]);

  if (sprays === null) return <div className="p-8 text-center text-stone-400">Loading spray plan…</div>;

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700 min-w-0">
          {onBack && <button onClick={onBack} className="p-1.5 -ml-1.5 rounded-lg hover:bg-stone-100 text-stone-500 shrink-0"><ChevronLeft size={20} /></button>}
          <Droplets size={18} /><h2 className="text-lg font-semibold text-stone-900 truncate">{type.label}</h2>
          <span className="text-sm text-stone-400 shrink-0">· {sprays.length} cards</span>
        </div>
        <div className="flex items-center gap-2">
          {manager && (
            <div className="inline-flex rounded-lg border border-stone-300 overflow-hidden">
              <button onClick={() => setView('kanban')} className={'px-3 py-2 text-sm font-medium ' + (view === 'kanban' ? 'bg-stone-900 text-stone-50' : 'bg-white text-stone-600 hover:bg-stone-50')}>Kanban</button>
              <button onClick={() => setView('grid')} className={'px-3 py-2 text-sm font-medium border-l border-stone-300 ' + (view === 'grid' ? 'bg-stone-900 text-stone-50' : 'bg-white text-stone-600 hover:bg-stone-50')}>Grid</button>
            </div>
          )}
          <button onClick={load} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /> Refresh</button>
          {manager && <button onClick={() => { setShowBuilder(v => !v); setShowLoader(false); }} className={cls.primary + ' !py-2 !px-3'}><Plus size={15} /> Add blocks</button>}
          {manager && <button onClick={() => { setShowLoader(v => !v); setShowBuilder(false); }} className={cls.ghost + ' !py-2 !px-3'}><Upload size={15} /> Load data</button>}
        </div>
      </div>

      {manager && <RoundPanel tc={tc} sprays={sprays} patchType={patchType}
        onApplyWater={async rate => {
          const next = (sprays || []).map(c => {
            const f = { ...(c.fields || {}) };
            const k = Object.keys(f).find(x => /water/i.test(x)) || 'Water rate';
            f[k] = `${fmtNum(rate)} L/ha`;
            return { ...c, fields: f };
          });
          await persist(next);
        }} />}

      {(!manager || view === 'kanban') && (
        <p className="text-xs text-stone-400 mb-3">Drag a card to reorder it up/down or move it between lanes. On a phone, press and hold a card first, then drag. Tick <span className="text-emerald-600 font-medium">Done</span> when sprayed — it sinks to the bottom and turns green.</p>
      )}

      {manager && showBuilder && (
        <SprayRoundBuilder config={config} tc={tc} statuses={statuses}
          onClose={() => setShowBuilder(false)}
          onAdd={async cards => { const base = await loadJSON(SK, sprays || []); await persist([...base, ...cards]); }} />
      )}

      {manager && showLoader && (
        <SprayLoader config={tc} setConfig={cfg => patchType({ statuses: cfg.statuses })}
          onLoaded={async (cards, replace, meta) => {
            const base = replace ? [] : await loadJSON(SK, sprays);
            const patch = {};
            if (replace && type.roundDeducted) patch.roundDeducted = false;
            if (meta && meta.mix && meta.mix.length) patch.roundMix = meta.mix;
            if (meta && meta.waterRate) patch.waterRate = meta.waterRate;
            if (Object.keys(patch).length || (meta && meta.products && meta.products.length)) {
              const next = { ...config };
              if (meta && meta.products && meta.products.length) {
                const have = new Set((config.products || []).map(p => p.name));
                next.products = [...(config.products || []), ...meta.products.filter(p => !have.has(p.name))];
              }
              next.sprayTypes = (config.sprayTypes || []).map(t => (t.key === typeKey ? { ...t, ...patch } : t));
              setConfig(next);
            }
            await persist([...base, ...cards]);
            setShowLoader(false);
          }} />
      )}

      {manager && view === 'grid' ? (
        <SprayGrid sprays={sprays} statuses={statuses} columns={columns} onPersist={persist} onToggleDone={toggleDone} config={tc} />
      ) : (
        <div ref={boardRef} className="flex gap-4 overflow-x-auto pb-3 -mx-4 px-4">
          {Object.keys(grouped).map(status => {
            const isMine = !manager && laneIsUser(status, operatorName);
            return (
            <div key={status}
              ref={el => { colRefs.current[status] = el; if (isMine) mineRef.current = el; }}
              className={'shrink-0 w-[320px] rounded-xl transition-colors ' +
                (overCol === status ? 'bg-stone-200/60 ring-2 ring-stone-400' : isMine ? 'bg-white ring-2 ring-stone-800' : '')}>
              <div className="flex items-center justify-between mb-2.5 px-2 pt-2">
                <span className={'text-[18px] font-bold leading-tight ' + (isMine ? 'text-stone-900' : 'text-stone-800')}>
                  {status}{isMine && <span className="text-[13px] font-semibold text-stone-500"> · you</span>}
                </span>
                <span className="text-xs text-stone-400 bg-stone-200/70 rounded-full px-2 py-0.5">{grouped[status].length}</span>
              </div>
              <div ref={el => { listRefs.current[status] = el; }} className="space-y-3 px-2 pb-2 min-h-[80px]">
                {(() => {
                  const visible = grouped[status];
                  const showInd = !!drag && overCol === status;
                  const line = <div className="h-1.5 rounded-full bg-emerald-500/80 mx-1" />;
                  if (visible.length === 0 && !showInd) {
                    return <div className="text-sm text-stone-400 italic py-6 text-center border border-dashed border-stone-300 rounded-xl">Drop here</div>;
                  }
                  const items = [];
                  visible.forEach((card, i) => {
                    if (showInd && overIndex === i) items.push(<div key={'ind-' + i}>{line}</div>);
                    const isDragged = !!drag && drag.card.id === card.id;
                    items.push(
                      <div key={card.id} data-card-id={card.id} className={isDragged ? 'opacity-25' : ''}>
                        <SprayCard card={card} manager={manager}
                          tank={tankFor(tc, status)} roundMix={tc.roundMix || []} config={tc}
                          onShift={shiftLane}
                          canLeft={statuses.indexOf(card.status) > 0}
                          canRight={statuses.indexOf(card.status) < statuses.length - 1}
                          onStartDrag={e => startPointer(e, card)}
                          onToggleDone={() => toggleDone(card)}
                          onEdit={() => setEditing(card)} onDelete={() => remove(card.id)} />
                      </div>
                    );
                  });
                  if (showInd && (overIndex == null || overIndex >= visible.length)) items.push(<div key="ind-end">{line}</div>);
                  return items;
                })()}
              </div>
            </div>
            );
          })}
        </div>
      )}

      {/* drag ghost */}
      {drag && (
        <div className="fixed z-[60] pointer-events-none w-[296px] opacity-90"
          style={{ left: drag.x, top: drag.y, transform: 'translate(-30px, -20px) rotate(2deg)' }}>
          <SprayCard card={drag.card} manager={false} ghost tank={tankFor(tc, drag.card.status)} roundMix={tc.roundMix || []} config={tc} />
        </div>
      )}

      {editing && (
        <SprayEditor card={editing} statuses={statuses}
          onClose={() => setEditing(null)}
          onSave={async updated => {
            await persist((sprays || []).map(c => (c.id === updated.id ? updated : c)));
            setEditing(null);
          }} />
      )}
    </div>
  );
}

function SprayCard({ card, manager, onStartDrag, onEdit, onDelete, onToggleDone, dragging, ghost, tank = 0, roundMix = [], config = {}, onShift, canLeft, canRight }) {
  const [showPart, setShowPart] = useState(false);
  const f = card.fields || {};
  const keys = Object.keys(f);
  const findKey = names => keys.find(k => names.some(n => k.toLowerCase().includes(n)));
  const titleKey = ('Block' in f) ? 'Block' : (findKey(['block', 'task']) || keys[0]);
  const productKey = ('Product' in f) ? 'Product' : (findKey(['product', 'mix', 'chemical', 'spray']) || null);
  const block = (titleKey && f[titleKey]) || 'Card';
  const showMix = (roundMix || []).length > 0;
  const product = showMix ? '' : (productKey ? f[productKey] : '');

  // structured figures
  const area = cardArea(card, config);
  const wr = cardWater(card, config);
  const rows = cardRows(card, config);
  const totalWater = area * wr;
  const fullTanks = tank > 0 ? Math.floor((totalWater + 1e-6) / tank) : 0;
  const remainder = tank > 0 ? Math.max(0, totalWater - fullTanks * tank) : totalWater;
  const hasPart = tank > 0 && remainder > 0.5;
  const PART_EXTRA = 40;            // always mix 40 L more than the part tank needs
  const partVol = remainder + PART_EXTRA;

  // hide fields we now render structurally
  const shownKeys = new Set([titleKey, productKey]);
  const isStructural = k => /total area|^area$|hectare|^ha$|water|^rows?$/i.test(k);
  const rest = keys.filter(k => !shownKeys.has(k) && f[k] !== '' && f[k] != null && !(showMix && /mix|product/i.test(k)) && !isStructural(k));
  const done = !!card.done;
  const stop = e => e.stopPropagation();
  const MixLine = ({ name, amt }) => (
    <div className="flex justify-between gap-2 text-[15px] leading-snug">
      <span className="text-stone-700 truncate">{name}</span>
      <span className="text-stone-900 font-medium shrink-0 tabular-nums">{fmtNum(amt)} {productUnit(config, name)}</span>
    </div>
  );
  const Fact = ({ label, value }) => (
    <div className="min-w-0">
      <div className="text-[10.5px] uppercase tracking-wide text-stone-400 leading-tight">{label}</div>
      <div className={'text-[14px] font-medium leading-tight truncate ' + (done ? 'text-stone-500' : 'text-stone-900')}>{value}</div>
    </div>
  );
  const tanksLabel = tank <= 0 ? 'Assign a sprayer'
    : remainder <= 0.5 ? `${fullTanks} full tank${fullTanks === 1 ? '' : 's'}`
      : fullTanks === 0 ? `Part tank · ${fmtNum(remainder)} L`
        : `${fullTanks} full + part ${fmtNum(remainder)} L`;

  return (
    <div
      onPointerDown={ghost ? undefined : onStartDrag}
      style={ghost ? undefined : { touchAction: 'auto' }}
      className={
        'border rounded-xl shadow-sm select-none ' +
        (ghost ? 'bg-white shadow-xl ' : 'cursor-grab active:cursor-grabbing ') +
        (dragging ? 'opacity-30 ' : '') +
        (done ? 'bg-stone-50 border-stone-200 opacity-60 ' : 'bg-white border-stone-200 ')
      }>
      <div className="p-3.5">
        <div className="flex items-start justify-between gap-2">
          <div className="min-w-0">
            <div className={'font-bold leading-tight text-[18px] ' + (done ? 'text-stone-500 line-through' : 'text-stone-900')}>{block}</div>
            {(() => {
              const cert = certOf(block, config);
              if (!cert) return null;
              const offending = certRestricted(cert) ? nonOrganicInMix(roundMix, config.products) : [];
              return (
                <div className="flex items-center gap-1.5 mt-1 flex-wrap">
                  <span className={'text-[11px] font-semibold uppercase tracking-wide border rounded px-1.5 py-0.5 ' + certTone(cert)}>
                    {cert === 'Conversion' ? 'In conversion' : cert}
                  </span>
                  {offending.length > 0 && (
                    <span className="text-[11px] font-semibold uppercase tracking-wide text-red-700 bg-red-50 border border-red-200 rounded px-1.5 py-0.5">
                      Mix not permitted
                    </span>
                  )}
                </div>
              );
            })()}
            {product && <div className={'text-[16px] mt-1 break-words ' + (done ? 'text-stone-400' : 'text-stone-700')}>{product}</div>}
          </div>
          {manager && !ghost && (
            <div className="flex gap-1 shrink-0">
              <button onPointerDown={stop} onClick={onEdit} className="p-1.5 rounded-md hover:bg-stone-100 text-stone-500"><Pencil size={15} /></button>
              <button onPointerDown={stop} onClick={onDelete} className="p-1.5 rounded-md hover:bg-red-50 text-red-500"><Trash2 size={15} /></button>
            </div>
          )}
        </div>

        {/* key figures */}
        {(area > 0 || rows || wr > 0) && (
          <div className="mt-2.5 grid grid-cols-2 gap-x-3 gap-y-2">
            {area > 0 && <Fact label="Total area" value={`${fmtNum(area)} ha`} />}
            {rows && <Fact label="Rows" value={rows} />}
            {wr > 0 && <Fact label="Water rate" value={`${fmtNum(wr)} L/ha`} />}
            {totalWater > 0 && <Fact label="Total water" value={`${fmtNum(totalWater)} L`} />}
            {totalWater > 0 && <div className="col-span-2"><Fact label="Tanks" value={tanksLabel} /></div>}
          </div>
        )}

        {rest.length > 0 && (
          <dl className="mt-2.5 space-y-1">
            {rest.map(k => (
              <div key={k} className="flex gap-2 text-[15px] leading-snug">
                <dt className="text-stone-400 shrink-0 min-w-[84px]">{k}</dt>
                <dd className={done ? 'text-stone-500 break-words' : 'text-stone-800 break-words'}>{String(f[k])}</dd>
              </div>
            ))}
          </dl>
        )}

        {showMix && (
          <div className="mt-3 rounded-lg bg-stone-50 border border-stone-200 p-2.5 space-y-0.5">
            <div className="text-[12px] font-semibold uppercase tracking-[0.12em] text-stone-500 mb-1">100 L mix</div>
            {roundMix.map(m => <MixLine key={'p' + m.product} name={m.product} amt={ratePer100(config, m, water)} />)}
            {tank > 0 ? (
              <>
                <div className="text-[12px] font-semibold uppercase tracking-[0.12em] text-emerald-700 mt-2 mb-1">Full tank mix · {tank} L</div>
                {roundMix.map(m => <MixLine key={'t' + m.product} name={m.product} amt={amountForVolume(config, m, tank, water)} />)}
                {hasPart && !ghost && (
                  <div className="mt-2">
                    <button onPointerDown={stop} onClick={() => setShowPart(v => !v)}
                      className="inline-flex items-center gap-1.5 text-[12.5px] font-medium text-stone-700 bg-white border border-stone-300 rounded-full px-2.5 py-1 hover:bg-stone-50">
                      <Beaker size={13} /> {showPart ? 'Hide' : 'Part tank mix'} · {fmtNum(remainder)} L
                    </button>
                    {showPart && (
                      <div className="mt-2 rounded-lg bg-amber-50 border border-amber-200 p-2.5 space-y-0.5">
                        <div className="text-[12px] font-semibold uppercase tracking-[0.12em] text-amber-700 mb-1">Part tank · mix for {fmtNum(partVol)} L <span className="normal-case font-normal text-amber-600">({fmtNum(remainder)} L + 40 L)</span></div>
                        {roundMix.map(m => <MixLine key={'pt' + m.product} name={m.product} amt={amountForVolume(config, m, partVol, water)} />)}
                        <div className="text-[11.5px] text-amber-700/80 pt-1">Water to {fmtNum(partVol)} L.</div>
                      </div>
                    )}
                  </div>
                )}
              </>
            ) : (
              <div className="text-[12px] text-stone-400 mt-1.5">Drop into an operator's lane for the full-tank mix.</div>
            )}
          </div>
        )}
        {!ghost && onShift && (
          <div className="flex items-center gap-1.5 mt-3">
            <button onPointerDown={stop} onClick={() => onShift(card, -1)} disabled={!canLeft}
              title="Move to the lane on the left"
              className="px-2 py-2 rounded-lg border border-stone-300 bg-white text-stone-600 hover:bg-stone-50 disabled:opacity-30">
              <ChevronLeft size={15} />
            </button>
            <button onPointerDown={stop} onClick={() => onShift(card, 1)} disabled={!canRight}
              title="Move to the lane on the right"
              className="px-2 py-2 rounded-lg border border-stone-300 bg-white text-stone-600 hover:bg-stone-50 disabled:opacity-30">
              <ChevronRight size={15} />
            </button>
          </div>
        )}
        {!ghost && (
          <button onPointerDown={stop} onClick={onToggleDone}
            className={'mt-3 inline-flex items-center gap-2 pl-2 pr-3 py-1.5 rounded-full text-sm font-medium border transition-colors ' +
              (done ? 'bg-emerald-600 border-emerald-600 text-white hover:bg-emerald-700'
                : 'bg-white border-stone-300 text-stone-600 hover:bg-stone-50')}>
            <span className={'w-5 h-5 rounded-md flex items-center justify-center border ' + (done ? 'border-white/70 bg-white/20' : 'border-stone-400')}>
              {done && <Check size={13} />}
            </span>
            {done ? 'Done' : 'Mark done'}
          </button>
        )}
      </div>
    </div>
  );
}

function SprayGrid({ sprays, statuses, columns, onPersist, config = {} }) {
  const [rows, setRows] = useState(sprays || []);
  const editing = useRef(false);
  const [newCol, setNewCol] = useState('');
  useEffect(() => { if (!editing.current) setRows(sprays || []); }, [sprays]);

  const commit = next => { setRows(next); onPersist(next); };
  const setCell = (id, key, value) => setRows(rs => rs.map(c => (c.id === id ? { ...c, fields: { ...c.fields, [key]: value } } : c)));
  const persistNow = () => { editing.current = false; setRows(rs => { onPersist(rs); return rs; }); };
  const setLane = (id, status) => commit(rows.map(c => (c.id === id ? { ...c, status } : c)));
  const toggleDone = id => commit(rows.map(c => {
    if (c.id !== id) return c;
    const done = !c.done; const fields = { ...c.fields };
    if (done) { fields['Actual date'] = todayNZ(); fields['Actual time'] = nowTimeNZ(); }
    else { delete fields['Actual date']; delete fields['Actual time']; }
    return done
      ? { ...c, done, fields, doneAt: todayNZ(), doneTime: nowTimeNZ(), doneTs: Date.now() }
      : { ...c, done, fields, doneAt: '', doneTime: '', doneTs: null };
  }));
  const del = id => commit(rows.filter(c => c.id !== id));
  const add = () => { const fields = {}; columns.forEach(k => (fields[k] = '')); commit([...rows, { id: uid(), status: statuses[0], done: false, fields }]); };
  const addColumn = () => {
    const k = newCol.trim(); if (!k || columns.includes(k)) { setNewCol(''); return; }
    commit(rows.map(c => ({ ...c, fields: { ...c.fields, [k]: c.fields[k] ?? '' } }))); setNewCol('');
  };

  return (
    <div>
      <div className="overflow-x-auto border border-stone-200 rounded-xl bg-white">
        <table className="text-sm whitespace-nowrap">
          <thead>
            <tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200 bg-stone-50">
              <th className="px-3 py-2.5 font-semibold">Done</th>
              <th className="px-3 py-2.5 font-semibold">Lane</th>
              <th className="px-3 py-2.5 font-semibold">Mix</th>
              {columns.map(k => <th key={k} className="px-3 py-2.5 font-semibold">{k}</th>)}
              <th className="px-3 py-2.5"></th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr><td colSpan={columns.length + 4} className="px-3 py-6 text-center text-stone-400">No cards. Add one below or load your diary.</td></tr>
            ) : rows.map(c => (
              <tr key={c.id} className={'border-b border-stone-100 last:border-0 ' + (c.done ? 'bg-emerald-50/40' : '')}>
                <td className="px-3 py-2 text-center">
                  <input type="checkbox" checked={!!c.done} onChange={() => toggleDone(c.id)} className="w-4 h-4 accent-emerald-600" />
                </td>
                <td className="px-2 py-2">
                  <select value={c.status} onChange={e => setLane(c.id, e.target.value)}
                    className="px-2 py-1.5 rounded-md border border-stone-200 bg-white text-stone-800 text-[13px] focus:outline-none focus:ring-2 focus:ring-stone-400/40">
                    {statuses.map(s => <option key={s}>{s}</option>)}
                    {!statuses.includes(c.status) && <option>{c.status}</option>}
                  </select>
                </td>
                <td className="px-3 py-2 align-top">
                  {(config.roundMix || []).length === 0 ? <span className="text-stone-300">—</span> : (() => {
                    const tank = tankFor(config, c.status);
                    return (
                      <div className="space-y-0.5">
                        {(config.roundMix || []).map(m => {
                          const w = numOf(config.waterRate);
                          const amt = tank > 0 ? amountForVolume(config, m, tank, w) : ratePer100(config, m, w);
                          return (
                            <div key={m.product} className="text-[13px] leading-tight text-stone-800">
                              {m.product} - <span className="font-medium tabular-nums">{fmtNum(amt)} {productUnit(config, m.product)}{tank > 0 ? '' : '/100L'}</span>
                            </div>
                          );
                        })}
                        <div className="text-[10px] uppercase tracking-wide text-stone-400 pt-0.5">{tank > 0 ? `full tank · ${tank} L` : 'per 100 L'}</div>
                      </div>
                    );
                  })()}
                </td>
                {columns.map(k => (
                  <td key={k} className="px-2 py-2">
                    <input value={c.fields[k] ?? ''} onFocus={() => (editing.current = true)}
                      onChange={e => setCell(c.id, k, e.target.value)} onBlur={persistNow}
                      className="w-full min-w-[130px] px-2 py-1.5 rounded-md border border-stone-200 bg-white text-stone-800 text-[13px] focus:outline-none focus:ring-2 focus:ring-stone-400/40" />
                  </td>
                ))}
                <td className="px-2 py-2 text-right">
                  <button onClick={() => del(c.id)} className="p-1.5 rounded-md hover:bg-red-50 text-red-500"><Trash2 size={15} /></button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="flex items-center gap-2 mt-3 flex-wrap">
        <button onClick={add} className={cls.primary + ' !py-2 !px-3'}><Plus size={15} /> Add card</button>
        <div className="flex items-center gap-2 ml-auto">
          <input value={newCol} onChange={e => setNewCol(e.target.value)} onKeyDown={e => e.key === 'Enter' && addColumn()} placeholder="New column name" className={cls.input + ' !w-44 !py-2'} />
          <button onClick={addColumn} className={cls.ghost + ' !py-2 !px-3'}><Plus size={15} /> Column</button>
        </div>
      </div>
      <p className="text-xs text-stone-400 mt-2">Edits save when you click out of a cell. Ticking Done stamps today as the Actual date and sinks the card in Kanban view.</p>
    </div>
  );
}

/* ---------- GrapeLink Operator Job Sheet (PDF) import ---------- */
// Extract a PDF into rows; each row is an array of cell strings, left→right.
async function extractPdfRows(arrayBuffer) {
  const pdfjs = await import('pdfjs-dist');
  try {
    const workerUrl = (await import('pdfjs-dist/build/pdf.worker.min.mjs?url')).default;
    pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;
  } catch {
    pdfjs.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.7.76/build/pdf.worker.min.mjs';
  }
  const doc = await pdfjs.getDocument({ data: new Uint8Array(arrayBuffer) }).promise;
  const out = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    const map = new Map();
    for (const it of tc.items) {
      const s = (it.str || '').trim(); if (!s) continue;
      const x = it.transform[4], y = Math.round(it.transform[5]);
      let key = [...map.keys()].find(k => Math.abs(k - y) <= 3);
      if (key == null) { key = y; map.set(key, []); }
      map.get(key).push({ x, s });
    }
    [...map.entries()].sort((a, b) => b[0] - a[0]).forEach(([, items]) => out.push(items.sort((a, b) => a.x - b.x).map(i => i.s)));
  }
  return out;
}

const isNumStr = v => /^-?\d+(\.\d+)?$/.test(String(v).trim());
// normalise a rate+unit to L / Kg (mL→L, g→Kg)
function normRate(per100, unit) {
  const u = (unit || '').toLowerCase();
  if (u === 'ml') return { per100: per100 / 1000, unit: 'L' };
  if (u === 'g') return { per100: per100 / 1000, unit: 'Kg' };
  return { per100, unit: unit || '' };
}

// Parse extracted rows from a GrapeLink Operator Job Sheet.
function parseGrapeLink(rows) {
  const findAfter = label => {
    for (const cells of rows) {
      const i = cells.findIndex(c => c.toLowerCase().replace(/\s+/g, ' ').includes(label));
      if (i >= 0) for (let j = i + 1; j < cells.length; j++) if (isNumStr(cells[j])) return numOf(cells[j]);
    }
    return null;
  };
  const waterRate = findAfter('spray vol/ha');
  const tank = findAfter('tank');
  let method = '';
  for (const cells of rows) { const i = cells.findIndex(c => /application method/i.test(c)); if (i >= 0 && cells[i + 1]) { method = cells[i + 1]; break; } }

  const startM = rows.findIndex(c => c.join(' ').toLowerCase().includes('materials to apply'));
  const endM = rows.findIndex(c => c.join(' ').toLowerCase().includes('blocks included'));
  const mix = [], products = [];
  for (let r = startM; r < endM && r >= 0; r++) {
    const cells = rows[r]; if (!cells) continue;
    const name0 = cells[0] || '';
    if (/^(water|hasprayed|materials|cf|rate|dilute|mix)/i.test(name0) || !/[A-Za-z]/.test(name0)) continue;
    const nums = [], unitsAfter = [];
    for (let k = 1; k < cells.length; k++) if (isNumStr(cells[k])) { nums.push(numOf(cells[k])); unitsAfter.push(cells[k + 1] && !isNumStr(cells[k + 1]) ? cells[k + 1] : ''); }
    if (nums.length < 3) continue;                       // dilute, CF, mix-rate at least
    const { per100, unit } = normRate(nums[2], unitsAfter[2] || unitsAfter[0]);   // Mix Rate/100L
    const name = name0.replace(/\s*\[[^\]]*\]\s*$/, '').trim();
    mix.push({ product: name, per100: Math.round(per100 * 1000) / 1000 });
    products.push({ name, unit, concentration: '', rate: Math.round(per100 * 1000) / 1000, stock: '', minStock: '' });
  }

  const endB = rows.findIndex((c, i) => i > endM && /^totals/i.test(c[0] || ''));
  const cards = [];
  for (let r = endM + 1; r < (endB < 0 ? rows.length : endB); r++) {
    const cells = rows[r]; if (!cells || cells.length < 8) continue;
    const len = cells.length;
    if (!cells.slice(len - 7).every(isNumStr)) continue;  // 7 trailing numeric columns
    const ha = numOf(cells[len - 6]);
    const haVol = numOf(cells[len - 3]);
    const block = cells.slice(1, len - 7).join(' ').trim();
    if (!block || !ha) continue;
    cards.push({ block, vineyard: cells[0], ha, waterRate: haVol || waterRate });
  }
  return { waterRate, tank, method, mix, products, cards };
}

function SprayLoader({ config, setConfig, onLoaded }) {
  const [text, setText] = useState('');
  const [replace, setReplace] = useState(true);
  const [err, setErr] = useState('');
  const fileRef = useRef(null);

  const rowsToCards = rows => {
    if (!rows.length) return [];
    const headers = Object.keys(rows[0] || {});
    const hasData = {}; headers.forEach(h => { hasData[h] = rows.some(r => cellToStr(r[h]) !== ''); });
    const lc = h => h.toLowerCase();
    const assignedKey = headers.find(h => lc(h).includes('assigned'));
    const doneKey = headers.find(h => ['sprayed', 'done', 'complete', 'completed'].includes(lc(h)) || lc(h).includes('sprayed'));
    const statusKey = headers.find(h => lc(h) === 'status');

    const cards = []; const newStatuses = new Set();
    rows.forEach(r => {
      if (!headers.some(h => cellToStr(r[h]) !== '')) return;
      const sVal = statusKey ? cellToStr(r[statusKey]) : '';
      const aVal = assignedKey ? cellToStr(r[assignedKey]) : '';
      let status = sVal || aVal || config.statuses[0];
      if (status && !config.statuses.includes(status) && !newStatuses.has(status)) newStatuses.add(status);
      const done = doneKey ? isTruthy(r[doneKey]) : false;
      const fields = {};
      headers.forEach(h => {
        if (!hasData[h] || NOISE_COLS.includes(lc(h))) return;
        if (h === statusKey || h === assignedKey || h === doneKey) return;
        const raw = r[h];
        const val = (/date/i.test(h) && typeof raw === 'number' && raw > 20000 && raw < 80000)
          ? serialToDate(raw) : cellToStr(raw);
        if (val !== '') fields[h] = val;
      });
      cards.push({ id: uid(), status: status || config.statuses[0], done, fields });
    });
    if (newStatuses.size) setConfig({ ...config, statuses: [...config.statuses, ...newStatuses] });
    return cards;
  };

  const parseText = () => {
    setErr('');
    if (!text.trim()) { setErr('Paste some rows first.'); return; }
    const aoa = Papa.parse(text.replace(/\r/g, ''), { skipEmptyLines: true }).data;
    const cards = rowsToCards(aoaToRows(aoa));
    if (!cards.length) { setErr('No data rows found. Make sure one row holds your column headers.'); return; }
    onLoaded(cards, replace);
  };

  const parseFile = file => {
    setErr('');
    const name = file.name.toLowerCase();
    const reader = new FileReader();
    if (name.endsWith('.pdf')) {
      reader.onload = async e => {
        try {
          const rows = await extractPdfRows(e.target.result);
          const g = parseGrapeLink(rows);
          if (!g.cards.length) { setErr('Couldn’t read any blocks from that PDF. Is it a GrapeLink Operator Job Sheet?'); return; }
          const cards = g.cards.map(c => ({
            id: uid(), status: (config.statuses || [])[0] || 'To Spray', done: false,
            fields: {
              Block: c.block, 'Total area': `${fmtNum(c.ha)} ha`,
              'Water rate': `${fmtNum(c.waterRate || g.waterRate || 0)} L/ha`,
              Vineyard: c.vineyard || '', Method: g.method || '',
            },
          }));
          onLoaded(cards, replace, { mix: g.mix, waterRate: g.waterRate, products: g.products });
        } catch (err) {
          setErr('Couldn’t read that PDF in this view. PDF import runs in the installed app — or paste the rows instead.');
        }
      };
      reader.readAsArrayBuffer(file);
      return;
    }
    reader.onload = e => {
      try {
        let aoa;
        if (name.endsWith('.csv') || name.endsWith('.tsv')) {
          aoa = Papa.parse(String(e.target.result).replace(/\r/g, ''), { skipEmptyLines: true }).data;
        } else {
          const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
          aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
        }
        const cards = rowsToCards(aoaToRows(aoa));
        if (!cards.length) { setErr('No data rows found in that file.'); return; }
        onLoaded(cards, replace);
      } catch { setErr('Could not read that file.'); }
    };
    if (name.endsWith('.csv') || name.endsWith('.tsv')) reader.readAsText(file);
    else reader.readAsArrayBuffer(file);
  };

  return (
    <div className={cls.card + ' p-4 mb-5'}>
      <p className="text-sm text-stone-600 mb-3">
        Upload a <b>GrapeLink Operator Job Sheet</b> (.pdf) to load a round — blocks, areas, water rate and the product mix are read in automatically. You can also upload your <b>Spray Diary</b> (.xlsx/.csv) or paste rows.
      </p>
      <textarea value={text} onChange={e => setText(e.target.value)} rows={5}
        placeholder={'Block\tProduct\tRate\tTarget\tWater Rate\tDate\tWHP\tStatus\nBlock 1\tSulphur 80WG\t3 kg/ha\tPowdery Mildew\t300 L/ha\t2026-11-15\t0 days\tTo Spray'}
        className={cls.input + ' font-mono text-[12px] resize-y'} />
      <div className="flex items-center gap-4 mt-3 flex-wrap">
        <label className="inline-flex items-center gap-2 text-sm text-stone-700">
          <input type="checkbox" checked={replace} onChange={e => setReplace(e.target.checked)} className="w-4 h-4 accent-stone-800" />
          Replace existing cards
        </label>
        <div className="flex gap-2 ml-auto">
          <input ref={fileRef} type="file" accept=".pdf,.csv,.tsv,.xlsx,.xls" className="hidden"
            onChange={e => e.target.files[0] && parseFile(e.target.files[0])} />
          <button onClick={() => fileRef.current?.click()} className={cls.ghost + ' !py-2 !px-3'}><Upload size={15} /> Upload PDF / file</button>
          <button onClick={parseText} className={cls.primary + ' !py-2 !px-3'}><Check size={15} /> Load pasted rows</button>
        </div>
      </div>
      {err && <p className="text-red-700 text-sm mt-2">{err}</p>}
    </div>
  );
}

function SprayEditor({ card, statuses, onClose, onSave }) {
  const [fields, setFields] = useState({ ...card.fields });
  const [status, setStatus] = useState(card.status);
  const [newKey, setNewKey] = useState('');
  return (
    <div className="fixed inset-0 z-50 bg-stone-900/40 flex items-end sm:items-center justify-center p-0 sm:p-4" onClick={onClose}>
      <div className="bg-white w-full sm:max-w-md rounded-t-2xl sm:rounded-2xl max-h-[90vh] overflow-auto" onClick={e => e.stopPropagation()}>
        <div className="flex items-center justify-between p-4 border-b border-stone-200 sticky top-0 bg-white">
          <h3 className="font-semibold text-stone-900">Edit card</h3>
          <button onClick={onClose} className="p-1.5 rounded-md hover:bg-stone-100 text-stone-500"><X size={18} /></button>
        </div>
        <div className="p-4 space-y-3">
          <div>
            <label className={cls.label}>Status (lane)</label>
            <select value={status} onChange={e => setStatus(e.target.value)} className={cls.input}>
              {statuses.map(s => <option key={s}>{s}</option>)}
              {!statuses.includes(status) && <option>{status}</option>}
            </select>
          </div>
          {Object.keys(fields).map(k => (
            <div key={k}>
              <label className={cls.label + ' flex items-center justify-between'}>
                <span>{k}</span>
                <button onClick={() => setFields(f => { const n = { ...f }; delete n[k]; return n; })} className="text-red-500 normal-case tracking-normal text-xs">remove</button>
              </label>
              <input value={fields[k]} onChange={e => setFields(f => ({ ...f, [k]: e.target.value }))} className={cls.input} />
            </div>
          ))}
          <div className="flex gap-2 pt-1">
            <input value={newKey} onChange={e => setNewKey(e.target.value)} placeholder="Add a field (e.g. Operator)" className={cls.input} />
            <button onClick={() => { if (newKey.trim()) { setFields(f => ({ ...f, [newKey.trim()]: '' })); setNewKey(''); } }} className={cls.ghost + ' shrink-0'}><Plus size={16} /></button>
          </div>
        </div>
        <div className="p-4 border-t border-stone-200 flex gap-2 sticky bottom-0 bg-white">
          <button onClick={onClose} className={cls.ghost + ' flex-1'}>Cancel</button>
          <button onClick={() => onSave({ ...card, status, fields })} className={cls.primary + ' flex-1'}>Save card</button>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   Timesheet — operator entry + own hours
   ============================================================ */
function TimesheetOperator({ config, session }) {
  const [date, setDate] = useState(todayStr());
  const [start, setStart] = useState('');
  const [finish, setFinish] = useState('');
  const [block, setBlock] = useState('');
  const [job, setJob] = useState('');
  const [note, setNote] = useState('');
  const [suggested, setSuggested] = useState(false);
  const [dateTouched, setDateTouched] = useState(false);
  // keep the date on today unless they've deliberately chosen another one
  useEffect(() => {
    if (dateTouched) return;
    const tick = () => { const t = todayStr(); setDate(d => (d === t || dateTouched ? d : t)); };
    tick();
    const id = setInterval(tick, 60000);
    const onShow = () => { if (!document.hidden) tick(); };
    document.addEventListener('visibilitychange', onShow);
    window.addEventListener('focus', tick);
    return () => { clearInterval(id); document.removeEventListener('visibilitychange', onShow); window.removeEventListener('focus', tick); };
  }, [dateTouched]);
  const [entries, setEntries] = useState([]);
  const [msg, setMsg] = useState('');
  const [busy, setBusy] = useState(false);

  const load = async () => setEntries(await loadJSON(K.ts(session.code), []));
  useEffect(() => { load(); }, []);

  // the finish time of the most recent entry on the selected date (entries are newest-first)
  const lastFinishFor = d => { const e = entries.find(x => x.date === d); return e ? (e.finish || '') : ''; };
  const lastEntry = (entries || [])[0];   // newest first
  // pre-fill Start with the last finish so the day carries on without re-typing
  useEffect(() => { if (!start) { const f = lastFinishFor(date); if (f) setStart(f); } }, [entries, date]);
  // opening the form: offer the block and task from the last entry logged
  useEffect(() => {
    if (!lastEntry || block || job) return;
    if (lastEntry.block) setBlock(lastEntry.block);
    if (lastEntry.job) setJob(lastEntry.job);
    if (lastEntry.block || lastEntry.job) setSuggested(true);
  }, [entries]);

  const myTasks = tasksFor(config, session);   // Jason/Simon get the machinery list
  const hours = calcHours(start, finish);
  const valid = date && start && finish && block && job;

  // lunch break: check the day as it stands, including the row being filled in
  const dayEntries = (entries || []).filter(e => e.date === date);
  const lunch = lunchCheck(start && finish ? [...dayEntries, { start, finish }] : dayEntries);

  const submit = async () => {
    if (!valid) return;
    setBusy(true);
    const jobObj = myTasks.find(j => j.name === job) || (config.jobs || []).find(j => j.name === job);
    const entry = { id: uid(), operatorCode: session.code, operatorName: session.name, date, start, finish, hours, block, job, jobCode: jobObj ? jobObj.code : '', note: note.trim(), createdAt: Date.now() };
    const fresh = await loadJSON(K.ts(session.code), entries);
    const next = [entry, ...fresh];
    await saveJSON(K.ts(session.code), next);
    setEntries(next);
    // carry the finish time into the next start, and keep the block and task
    // as a suggestion — most of the time the next entry is on the same job
    setStart(finish); setFinish(''); setNote('');
    setSuggested(true);
    setMsg(`Saved ${hours} h on ${block}. Next task starts at ${finish}.`); setBusy(false);
    setTimeout(() => setMsg(''), 4000);
  };

  const removeEntry = async id => {
    const next = entries.filter(e => e.id !== id);
    await saveJSON(K.ts(session.code), next); setEntries(next);
  };

  // the operator only needs the current pay period in front of them; older
  // entries stay saved and still come through on the manager's export
  // show only the current pay period — it clears at noon on the Monday after
  // each fortnight ends, once they've had the weekend to check their hours
  const period = timesheetPeriod();
  const recent = entries.filter(e => e.date >= period.startISO && e.date <= period.endISO)
    .slice().sort((a, b) => (b.date + (b.start || '')).localeCompare(a.date + (a.start || '')));
  const fortnightTotal = recent.reduce((s, e) => s + (e.hours || 0), 0);
  const weekStart = mondayOf(nzNow());
  const weekTotal = entries.filter(e => e.date >= weekStart).reduce((s, e) => s + (e.hours || 0), 0);

  return (
    <div className="space-y-6">
      <Banner msg={msg} />
      <div className={cls.card + ' p-4'}>
        <div className="flex items-center gap-2 mb-4 text-stone-700"><Clock size={18} /><h2 className="text-lg font-semibold text-stone-900">Log time</h2></div>
        <div className="space-y-3.5">
          <div><label className={cls.label}>Date</label><input type="date" value={date} onChange={e => { setDate(e.target.value); setDateTouched(e.target.value !== todayStr()); }} className={cls.input} />
            {date !== todayStr() && (
              <button onClick={() => { setDate(todayStr()); setDateTouched(false); }} className="text-xs text-stone-500 underline mt-1">Back to today</button>
            )}</div>
          <div className="grid grid-cols-2 gap-3">
            <div><label className={cls.label}>Start</label>
              <select value={start} onChange={e => setStart(e.target.value)} className={cls.input}>
                <option value="">--:--</option>
                {START_OPTIONS.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
            <div><label className={cls.label}>Finish</label>
              <select value={finish} onChange={e => setFinish(e.target.value)} className={cls.input}>
                <option value="">--:--</option>
                {FINISH_OPTIONS.map(t => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>
          </div>
          {lastFinishFor(date) && start === lastFinishFor(date) && (
            <p className="text-xs text-stone-400 -mt-1">Start carried over from your last finish ({lastFinishFor(date)}) — change it if needed.</p>
          )}

          {lunch.needed && !lunch.ok && (
            <div className="flex items-start gap-2 rounded-lg border border-red-300 bg-red-50 px-3 py-2.5">
              <AlertTriangle size={17} className="text-red-600 shrink-0 mt-0.5" />
              <div className="text-[13.5px] text-red-800 leading-snug">
                <b>No 30 minute lunch break on this day.</b>{' '}
                {lunch.gap > 0
                  ? `The longest gap is ${lunch.gap} min${lunch.gapAt != null ? ` (from ${hhmm(lunch.gapAt)})` : ''}.`
                  : 'Your entries run straight through with no gap.'}
                {' '}Usually 12:30–13:00 — split your hours around the break, or adjust the times if you took it at a different time.
              </div>
            </div>
          )}
          {lunch.needed && lunch.ok && (
            <p className="text-xs text-emerald-700 -mt-1 inline-flex items-center gap-1.5"><Check size={14} /> {lunch.gap} min break recorded on this day.</p>
          )}
          <Combobox label="Block" options={(config.blocks || []).map(b => b.name)} value={block} onChange={v => { setBlock(v); setSuggested(false); }} icon={MapPin} placeholder="Search blocks…" />
          <Combobox label="Job" options={myTasks.map(j => j.name)} value={job} onChange={v => { setJob(v); setSuggested(false); }} icon={Layers} placeholder="Search jobs…" />
          {suggested && (block || job) && (
            <p className="text-xs text-stone-400 -mt-1 flex items-center gap-2">
              Carried over from your last entry — change it if you've moved on.
              <button onClick={() => { setBlock(''); setJob(''); setSuggested(false); }}
                className="underline hover:text-stone-600">Clear</button>
            </p>
          )}
          <div><label className={cls.label}>Note (optional)</label><input value={note} onChange={e => setNote(e.target.value)} className={cls.input} placeholder="Anything worth recording" /></div>
          <div className="flex items-center justify-between pt-1">
            <div className="text-sm text-stone-500">Total: <span className="font-semibold text-stone-900 text-base">{hours} h</span></div>
            <button onClick={submit} disabled={!valid || busy} className={cls.primary}>{busy ? 'Saving…' : 'Save entry'} <Check size={16} /></button>
          </div>
        </div>
      </div>

      <div className={cls.card + ' p-4 flex items-center justify-between gap-4'}>
        <div>
          <div className="text-[11px] font-semibold uppercase tracking-[0.14em] text-stone-400">This fortnight · {period.label}</div>
          <div className="flex items-baseline gap-1">
            <span style={serif} className="text-5xl font-bold text-stone-900 leading-none tabular-nums">{Math.round(fortnightTotal * 100) / 100}</span>
            <span className="text-2xl font-semibold text-stone-500">h</span>
          </div>
          <div className="text-sm text-stone-500 mt-1.5">This week {Math.round(weekTotal * 100) / 100} h</div>
        </div>
        <div className="w-14 h-14 rounded-2xl bg-stone-900 text-stone-50 flex items-center justify-center shrink-0"><Clock size={26} /></div>
      </div>

      <div>
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-base font-semibold text-stone-900">Your hours <span className="font-normal text-stone-400 text-sm">· {period.label}</span></h3>
        </div>
        {recent.length === 0 ? (
          <p className="text-stone-400 text-sm py-6 text-center border border-dashed border-stone-300 rounded-xl">
            {entries.length ? 'Nothing logged this fortnight yet.' : 'No entries yet. Log your first above.'}
          </p>
        ) : (
          <div className="space-y-2">
            {recent.map(e => (
              <div key={e.id} className="flex items-center gap-3 bg-white border border-stone-200 rounded-lg px-3.5 py-3">
                <div className="text-center shrink-0 w-14">
                  <div className="text-lg font-semibold text-stone-900 leading-none">{e.hours}</div>
                  <div className="text-[10px] uppercase tracking-wide text-stone-400">hours</div>
                </div>
                <div className="min-w-0 flex-1">
                  <div className="text-[15px] text-stone-900 truncate">{e.block} · {e.job}</div>
                  <div className="text-xs text-stone-500">{fmtDate(e.date)} · {e.start}–{e.finish}{e.note ? ` · ${e.note}` : ''}</div>
                </div>
                <button onClick={() => removeEntry(e.id)} className="p-1.5 rounded-md hover:bg-red-50 text-red-400 shrink-0"><Trash2 size={15} /></button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   Hazard — operator form + own reports
   ============================================================ */
const sevStyle = {
  Low: 'bg-stone-100 text-stone-700 border-stone-300',
  Medium: 'bg-amber-100 text-amber-800 border-amber-300',
  High: 'bg-orange-100 text-orange-800 border-orange-300',
  Critical: 'bg-red-100 text-red-800 border-red-300',
};

function HazardForm({ config, session }) {
  const [date, setDate] = useState(todayStr());
  const [time, setTime] = useState(nowTime());
  const [block, setBlock] = useState('');
  const [location, setLocation] = useState('');
  const [type, setType] = useState(HAZARD_TYPES[0]);
  const [severity, setSeverity] = useState('Medium');
  const [desc, setDesc] = useState('');
  const [reports, setReports] = useState([]);
  const [msg, setMsg] = useState('');
  const [busy, setBusy] = useState(false);

  const load = async () => setReports(await loadJSON(K.hz(session.code), []));
  useEffect(() => { load(); }, []);

  const valid = desc.trim() && (block || location.trim());
  const submit = async () => {
    if (!valid) return;
    setBusy(true);
    const r = { id: uid(), reportedByCode: session.code, reportedByName: session.name, date, time, block, location: location.trim(), type, severity, description: desc.trim(), status: 'Open', createdAt: Date.now() };
    const fresh = await loadJSON(K.hz(session.code), reports);
    const next = [r, ...fresh];
    await saveJSON(K.hz(session.code), next);
    setReports(next);
    // notify the manager by email (best effort — never blocks the report being saved)
    const sent = config.webhookUrl ? await postWebhook(config.webhookUrl, {
      type: 'hazard',
      siteName: config.siteName,
      notifyEmail: config.notifyEmail || '',
      severity, hazardType: type,
      block: block || '', location: location.trim(),
      description: desc.trim(),
      reportedBy: session.name, operatorCode: session.code,
      date, time, submittedAt: new Date().toISOString(),
    }) : false;
    setBlock(''); setLocation(''); setDesc(''); setSeverity('Medium'); setType(HAZARD_TYPES[0]); setTime(nowTime());
    setMsg(config.webhookUrl
      ? (sent ? 'Hazard reported — your manager has been emailed.' : 'Hazard saved. Email alert could not be sent — your manager can still see it in the console.')
      : 'Hazard reported. Your manager can see it in the console.');
    setBusy(false);
    setTimeout(() => setMsg(''), 6000);
  };

  return (
    <div className="space-y-6">
      <Banner msg={msg} />
      <div className={cls.card + ' p-4'}>
        <div className="flex items-center gap-2 mb-1 text-stone-700"><AlertTriangle size={18} /><h2 className="text-lg font-semibold text-stone-900">Report a hazard</h2></div>
        <p className="text-sm text-stone-500 mb-4">If it’s an emergency or someone is hurt, deal with that first and call for help.</p>
        <div className="space-y-3.5">
          <div className="grid grid-cols-2 gap-3">
            <div><label className={cls.label}>Date</label><input type="date" value={date} onChange={e => setDate(e.target.value)} className={cls.input} /></div>
            <div><label className={cls.label}>Time</label><input type="time" value={time} onChange={e => setTime(e.target.value)} className={cls.input} /></div>
          </div>
          <Combobox label="Block" options={(config.blocks || []).map(b => b.name)} value={block} onChange={setBlock} icon={MapPin} placeholder="Search blocks…" />
          <div><label className={cls.label}>Location detail (optional)</label><input value={location} onChange={e => setLocation(e.target.value)} className={cls.input} placeholder="e.g. row 14, near the headland" /></div>
          <div><label className={cls.label}>Hazard type</label><select value={type} onChange={e => setType(e.target.value)} className={cls.input}>{HAZARD_TYPES.map(t => <option key={t}>{t}</option>)}</select></div>
          <div>
            <label className={cls.label}>Severity</label>
            <div className="grid grid-cols-4 gap-2">
              {SEVERITIES.map(s => (
                <button key={s} onClick={() => setSeverity(s)}
                  className={'py-2.5 rounded-lg border text-sm font-medium transition-all ' + (severity === s ? sevStyle[s] + ' ring-2 ring-offset-1 ring-stone-400' : 'bg-white border-stone-300 text-stone-500 hover:bg-stone-50')}>
                  {s}
                </button>
              ))}
            </div>
          </div>
          <div><label className={cls.label}>What did you see?</label><textarea value={desc} onChange={e => setDesc(e.target.value)} rows={3} className={cls.input + ' resize-y'} placeholder="Describe the hazard and any action you’ve taken." /></div>
          <button onClick={submit} disabled={!valid || busy} className={cls.primary + ' w-full py-3.5'}>{busy ? 'Sending…' : 'Submit hazard report'}</button>
        </div>
      </div>

      <div>
        <h3 className="text-base font-semibold text-stone-900 mb-3">Your reports</h3>
        {reports.length === 0 ? (
          <p className="text-stone-400 text-sm py-6 text-center border border-dashed border-stone-300 rounded-xl">Nothing reported yet.</p>
        ) : (
          <div className="space-y-2">{reports.map(r => <HazardRow key={r.id} r={r} />)}</div>
        )}
      </div>
    </div>
  );
}

function HazardRow({ r, onResolve }) {
  return (
    <div className="bg-white border border-stone-200 rounded-lg px-3.5 py-3">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className={'text-[11px] font-semibold uppercase tracking-wide px-2 py-0.5 rounded-full border ' + sevStyle[r.severity]}>{r.severity}</span>
            <span className="text-[15px] font-medium text-stone-900">{r.type}</span>
            {r.status === 'Resolved' && <span className="text-[11px] text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-full px-2 py-0.5">Resolved</span>}
          </div>
          <div className="text-sm text-stone-700 mt-1">{r.description}</div>
          <div className="text-xs text-stone-500 mt-1">
            {[r.block, r.location].filter(Boolean).join(' · ')} · {fmtDate(r.date)} {r.time}
            {r.reportedByName ? ` · ${r.reportedByName}` : ''}
          </div>
        </div>
        {onResolve && (
          <button onClick={() => onResolve(r.status === 'Resolved' ? 'Open' : 'Resolved')}
            className={'shrink-0 text-xs font-medium px-2.5 py-1.5 rounded-md border ' + (r.status === 'Resolved' ? 'border-stone-300 text-stone-500 hover:bg-stone-50' : 'border-emerald-300 text-emerald-700 hover:bg-emerald-50')}>
            {r.status === 'Resolved' ? 'Reopen' : 'Mark resolved'}
          </button>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   Manager — timesheet dashboard
   ============================================================ */
function TimesheetDashboard({ config }) {
  const [all, setAll] = useState(null);
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');

  const load = async () => {
    const rows = [];
    for (const op of config.operators) {
      const arr = await loadJSON(K.ts(op.code), []);
      arr.forEach(t => rows.push({ ...t, operatorName: t.operatorName || op.name, operatorCode: t.operatorCode || op.code }));
    }
    rows.sort((a, b) => (b.date + b.start).localeCompare(a.date + a.start));
    setAll(rows);
  };
  useEffect(() => { load(); }, []);

  const filtered = useMemo(() => (all || []).filter(t => (!from || t.date >= from) && (!to || t.date <= to)), [all, from, to]);
  const byOp = useMemo(() => {
    const m = {};
    filtered.forEach(t => { const k = t.operatorName; if (!m[k]) m[k] = { hours: 0, count: 0 }; m[k].hours += t.hours || 0; m[k].count += 1; });
    return Object.entries(m).map(([name, v]) => ({ name, hours: Math.round(v.hours * 100) / 100, count: v.count })).sort((a, b) => b.hours - a.hours);
  }, [filtered]);
  const maxH = Math.max(1, ...byOp.map(o => o.hours));

  const exportXlsx = () => {
    /* Matches the Love Block farm timesheet layout exactly:
       title rows, the nine entry columns, and the per-person
       summary block off to the right. */
    const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const pad = n => String(n).padStart(2, '0');
    const dmy = iso => { const [y, m, d] = String(iso).split('-'); return y ? `${d}/${m}/${y.slice(2)}` : iso; };
    const ampm = t => {
      if (!t) return '';
      const [hs, ms] = String(t).split(':'); let h = Number(hs); const m = ms || '00';
      const suffix = h >= 12 ? 'PM' : 'AM';
      h = h % 12; if (h === 0) h = 12;
      return `${h}:${m} ${suffix}`;
    };
    // fortnight ending: the To filter if set, else the latest entry, else today —
    // always rolled forward to the Sunday that ends the fortnight
    const endISO = to || filtered.map(t => t.date).sort().slice(-1)[0] || todayStr();
    const endD = (() => {
      const [y, m, d] = endISO.split('-').map(Number);
      const dt = new Date(y, m - 1, d);
      const dow = dt.getDay();                 // 0 = Sunday
      if (dow !== 0) dt.setDate(dt.getDate() + (7 - dow));
      return dt;
    })();
    const endLabel = `${pad(endD.getDate())} ${['January','February','March','April','May','June','July','August','September','October','November','December'][endD.getMonth()]} ${endD.getFullYear()}`;
    const tabName = `FE${pad(endD.getDate())}${MONTHS[endD.getMonth()]}${String(endD.getFullYear()).slice(2)}`;

    const names = [...new Set(filtered.map(t => t.operatorName))].sort((a, b) => String(a).localeCompare(String(b)));
    const byDate = (a, b) => (a.date + (a.start || '')).localeCompare(b.date + (b.start || ''));

    const rows = [];
    rows.push(['LOVE BLOCK FARM TIMESHEET']);
    rows.push([`Fortnightly ending ${endLabel}`]);
    rows.push(['Date', 'Name', 'Start ', 'Finish', 'Total', 'Job code', 'Job description ', ' Block Code', 'Account', 'Notes ', '', '', 'Name', 'Hrs', 'days', 'Total days', 'Kilometers driven own car - Fuel allowance']);

    // one entry per row, grouped by person then date
    const ordered = [];
    names.forEach(n => filtered.filter(t => t.operatorName === n).sort(byDate).forEach(t => ordered.push(t)));
    // look up the accounting block code and account for each row
    const blockCodes = config.blockCodes || {};
    const jobAccounts = config.jobAccounts || {};
    const lookBlock = b => (b && blockCodes[b]) ? blockCodes[b] : (b === 'N/A' || !b ? '' : '#NO MATCH');
    const lookAccount = j => (j && jobAccounts[j]) ? jobAccounts[j] : (j ? '#NO MATCH' : '');
    ordered.forEach(t => {
      rows.push([
        dmy(t.date), t.operatorName, ampm(t.start), ampm(t.finish),
        (Number(t.hours) || 0).toFixed(2),
        t.jobCode || '', t.job || '', lookBlock(t.block), lookAccount(t.job), t.note || '',
      ]);
    });

    // per-person summary in columns K–O, alongside the entries
    names.forEach((n, i) => {
      const mine = filtered.filter(t => t.operatorName === n);
      const hrs = mine.reduce((s, t) => s + (Number(t.hours) || 0), 0);
      // "days" = days on site, so annual/sick leave and holidays don't count
      const days = new Set(mine.filter(t => !isLeave(t)).map(t => t.date)).size;
      const r = rows[3 + i] || (rows[3 + i] = []);
      while (r.length < 11) r.push('');
      r[11] = i + 1; r[12] = n; r[13] = hrs.toFixed(2); r[14] = days; r[15] = days;
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }];
    ws['!cols'] = [{ wch: 10 }, { wch: 16 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 12 },
      { wch: 34 }, { wch: 22 }, { wch: 46 }, { wch: 30 }, { wch: 3 }, { wch: 4 }, { wch: 16 }, { wch: 9 }, { wch: 7 }, { wch: 11 }, { wch: 34 }];

    const wb = XLSX.utils.book_new();
    addSheet(wb, ws, tabName);
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `Timesheet_${tabName}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  /* ---- Smartsheet push ---- */
  const SS_SHEET   = 2014856873987972;
  const SS_COLS = {
    name:    3451088539701124,
    date:    4371468221632388,
    start:   78974826794884,
    finish:  4582574454165380,
    hours:   1199288726015876,
    jobDesc: 636338772594564,
    codes:   4531792111423364,
    notes:   2119668407947140,
    block:   7954688167071620,
  };
  const [ssStatus, setSsStatus] = useState('');
  const [ssBusy, setSsBusy]     = useState(false);

  const pushToSmartsheet = async () => {
    if (!filtered.length) return;
    setSsBusy(true); setSsStatus('');
    const toSS  = t => `${t.padStart(2,'0')}:00:00`.replace(/^(\d{2}:\d{2})$/, '$1:00');
    const fmtT  = t => { if (!t) return ''; const [h,m] = t.split(':'); return `${h.padStart(2,'0')}:${(m||'00').padStart(2,'0')}:00`; };
    const names = [...new Set(filtered.map(t => t.operatorName))].sort((a,b) => a.localeCompare(b));
    const byDate = (a,b) => (a.date+(a.start||'')).localeCompare(b.date+(b.start||''));
    const ordered = [];
    names.forEach(n => filtered.filter(t => t.operatorName === n).sort(byDate).forEach(t => ordered.push(t)));

    const rows = ordered.map(t => {
      const code    = t.jobCode || '';
      const jobDesc = code ? `${code}\t${t.job}` : (t.job || '');
      return {
        toBottom: true,
        cells: [
          { columnId: SS_COLS.name,    value: t.operatorName, strict: false },
          { columnId: SS_COLS.date,    value: t.date },
          { columnId: SS_COLS.start,   value: fmtT(t.start),  strict: false },
          { columnId: SS_COLS.finish,  value: fmtT(t.finish), strict: false },
          { columnId: SS_COLS.hours,   value: t.hours },
          { columnId: SS_COLS.jobDesc, value: jobDesc, strict: false },
          { columnId: SS_COLS.codes,   value: code },
          { columnId: SS_COLS.notes,   value: t.note || '' },
          { columnId: SS_COLS.block,   value: t.block, strict: false },
        ],
      };
    });

    // send in batches of 100 (Smartsheet API limit is 500, but keeping it safe)
    const BATCH = 100;
    let pushed = 0, errors = 0;
    for (let i = 0; i < rows.length; i += BATCH) {
      const batch = rows.slice(i, i + BATCH);
      setSsStatus(`Uploading ${Math.min(i + BATCH, rows.length)} / ${rows.length} rows…`);
      try {
        const res = await window.__ssAddRows(SS_SHEET, batch);
        pushed += (res?.result?.length || batch.length);
      } catch (e) {
        console.error('Smartsheet batch error', e);
        errors += batch.length;
      }
    }
    setSsBusy(false);
    setSsStatus(errors ? `⚠ ${pushed} rows uploaded, ${errors} failed — check the console.` : `✓ ${pushed} rows added to Smartsheet.`);
    setTimeout(() => setSsStatus(''), 6000);
  };

  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700"><Clock size={18} /><h2 className="text-lg font-semibold text-stone-900">Hours dashboard</h2></div>
        <div className="flex gap-2 flex-wrap">
          <button onClick={load} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /> Refresh</button>
          <button onClick={exportXlsx} disabled={!filtered.length} className={cls.ghost + ' !py-2 !px-3'}><Download size={15} /> Export Excel</button>
          <button onClick={pushToSmartsheet} disabled={!filtered.length || ssBusy} className={cls.primary + ' !py-2 !px-3'}>
            {ssBusy ? <><RefreshCw size={15} className="animate-spin" /> Uploading…</> : <><Upload size={15} /> Push to Smartsheet</>}
          </button>
        </div>
      </div>
      {ssStatus && <div className={'text-sm px-3 py-2 rounded-lg ' + (ssStatus.startsWith('⚠') ? 'bg-amber-50 text-amber-800 border border-amber-200' : 'bg-emerald-50 text-emerald-800 border border-emerald-200')}>{ssStatus}</div>}

      <RangePicker from={from} to={to} onChange={(f, t) => { setFrom(f); setTo(t); }} />

      <div className={cls.card + ' p-4'}>
        <div className="mb-4">
          <span className="text-[11px] font-semibold uppercase tracking-[0.14em] text-stone-500">Total hours by operator</span>
        </div>
        {byOp.length === 0 ? <p className="text-stone-400 text-sm text-center py-4">No hours in this range.</p> : (
          <div className="space-y-3">
            {byOp.map(o => (
              <div key={o.name}>
                <div className="flex justify-between text-sm mb-1">
                  <span className="text-stone-800 font-medium">{o.name}</span>
                  <span className="text-stone-700"><b className="text-stone-900">{o.hours} h</b> <span className="text-stone-400">· {o.count} {o.count === 1 ? 'entry' : 'entries'}</span></span>
                </div>
                <div className="h-2.5 rounded-full bg-stone-100 overflow-hidden">
                  <div className="h-full rounded-full" style={{ width: (o.hours / maxH * 100) + '%', backgroundColor: '#57534e' }} />
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      <div>
        <h3 className="text-base font-semibold text-stone-900 mb-2">All entries <span className="text-stone-400 font-normal">· {filtered.length}</span></h3>
        <div className="overflow-x-auto border border-stone-200 rounded-xl bg-white">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
                <th className="px-3 py-2.5 font-semibold">Operator</th>
                <th className="px-3 py-2.5 font-semibold">Date</th>
                <th className="px-3 py-2.5 font-semibold">Time</th>
                <th className="px-3 py-2.5 font-semibold text-right">Hrs</th>
                <th className="px-3 py-2.5 font-semibold">Block</th>
                <th className="px-3 py-2.5 font-semibold">Job</th>
                <th className="px-3 py-2.5 font-semibold">Code</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr><td colSpan={7} className="px-3 py-6 text-center text-stone-400">No entries.</td></tr>
              ) : filtered.map(t => (
                <tr key={t.id} className="border-b border-stone-100 last:border-0">
                  <td className="px-3 py-2.5 text-stone-900 whitespace-nowrap">{t.operatorName}</td>
                  <td className="px-3 py-2.5 text-stone-600 whitespace-nowrap">{fmtDate(t.date)}</td>
                  <td className="px-3 py-2.5 text-stone-600 whitespace-nowrap">{t.start}–{t.finish}</td>
                  <td className="px-3 py-2.5 text-right font-medium text-stone-900">{t.hours}</td>
                  <td className="px-3 py-2.5 text-stone-700 whitespace-nowrap">{t.block}</td>
                  <td className="px-3 py-2.5 text-stone-700 whitespace-nowrap">{t.job}</td>
                  <td className="px-3 py-2.5 text-stone-500 whitespace-nowrap font-mono text-xs">{t.jobCode || ''}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   Manager — hazard log
   ============================================================ */
function HazardLog({ config }) {
  const [all, setAll] = useState(null);
  const [showResolved, setShowResolved] = useState(false);

  const load = async () => {
    const rows = [];
    for (const op of config.operators) {
      const arr = await loadJSON(K.hz(op.code), []);
      arr.forEach(h => rows.push({ ...h, reportedByName: h.reportedByName || op.name, reportedByCode: h.reportedByCode || op.code }));
    }
    rows.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
    setAll(rows);
  };
  useEffect(() => { load(); }, []);

  const resolve = async (r, status) => {
    const arr = await loadJSON(K.hz(r.reportedByCode), []);
    const idx = arr.findIndex(h => h.id === r.id);
    if (idx >= 0) { arr[idx] = { ...arr[idx], status }; await saveJSON(K.hz(r.reportedByCode), arr); }
    setAll(rows => rows.map(h => (h.id === r.id ? { ...h, status } : h)));
  };

  if (all === null) return <div className="p-8 text-center text-stone-400">Loading hazards…</div>;
  const list = showResolved ? all : all.filter(r => r.status !== 'Resolved');
  const openCount = all.filter(r => r.status !== 'Resolved').length;

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700">
          <AlertTriangle size={18} /><h2 className="text-lg font-semibold text-stone-900">Hazards</h2>
          <span className="text-sm text-stone-400">· {openCount} open</span>
        </div>
        <div className="flex gap-2 items-center">
          <label className="inline-flex items-center gap-2 text-sm text-stone-700">
            <input type="checkbox" checked={showResolved} onChange={e => setShowResolved(e.target.checked)} className="w-4 h-4 accent-stone-800" /> Show resolved
          </label>
          <button onClick={load} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /> Refresh</button>
        </div>
      </div>
      {list.length === 0 ? (
        <p className="text-stone-400 text-sm py-8 text-center border border-dashed border-stone-300 rounded-xl">No hazards to show.</p>
      ) : (
        <div className="space-y-2">{list.map(r => <HazardRow key={r.id} r={r} onResolve={s => resolve(r, s)} />)}</div>
      )}
    </div>
  );
}

/* ============================================================
   Manager — setup / configuration
   ============================================================ */
function TagEditor({ label, items, onChange }) {
  const [val, setVal] = useState('');
  return (
    <div>
      <label className={cls.label}>{label}</label>
      <div className="flex flex-wrap gap-2 mb-2">
        {items.map((it, i) => (
          <span key={it + i} className="inline-flex items-center gap-1.5 bg-stone-100 border border-stone-300 rounded-full pl-3 pr-1.5 py-1 text-sm text-stone-800">
            {it}
            <button onClick={() => onChange(items.filter((_, j) => j !== i))} className="rounded-full hover:bg-stone-300/60 p-0.5"><X size={13} /></button>
          </span>
        ))}
        {items.length === 0 && <span className="text-sm text-stone-400">None yet.</span>}
      </div>
      <div className="flex gap-2">
        <input value={val} onChange={e => setVal(e.target.value)} onKeyDown={e => { if (e.key === 'Enter' && val.trim()) { onChange([...items, val.trim()]); setVal(''); } }}
          placeholder={`Add ${label.toLowerCase()}…`} className={cls.input} />
        <button onClick={() => { if (val.trim()) { onChange([...items, val.trim()]); setVal(''); } }} className={cls.ghost + ' shrink-0'}><Plus size={16} /></button>
      </div>
    </div>
  );
}

function BlocksEditor({ blocks, onChange }) {
  const [showPaste, setShowPaste] = useState(false);
  const [paste, setPaste] = useState('');
  const setRow = (i, patch) => onChange(blocks.map((b, j) => (j === i ? { ...b, ...patch } : b)));
  const totalHa = blocks.reduce((s, b) => s + numOf(b.ha), 0);
  const totalKm = blocks.reduce((s, b) => s + numOf(b.km), 0);
  const load = replace => {
    const aoa = Papa.parse(paste.replace(/\r/g, ''), { skipEmptyLines: true }).data;
    const parsed = [];
    aoa.forEach((row, idx) => {
      const cells = (row || []).map(c => String(c == null ? '' : c).trim());
      if (!cells.length || !cells[0]) return;
      const joined = cells.join(' ').toLowerCase();
      if (idx === 0 && (joined.includes('block') || joined.includes('name')) && /ha|hectare|area/.test(joined)) return;
      parsed.push({ name: cells[0], ha: numOf(cells[1]), rows: cells[2] || '', km: numOf(cells[3]), rowWidth: cells[4] ? numOf(cells[4]) : '' });
    });
    if (!parsed.length) return;
    onChange(replace ? parsed : [...blocks, ...parsed]);
    setPaste(''); setShowPaste(false);
  };
  const gi = 'px-2 py-1.5 rounded-md border border-stone-200 bg-white text-stone-800 text-[13px] focus:outline-none focus:ring-2 focus:ring-stone-400/40';
  return (
    <div>
      <label className={cls.label + ' flex items-center justify-between'}>
        <span>Blocks</span>
        <span className="normal-case tracking-normal text-stone-400">{blocks.length} blocks · {fmtNum(Math.round(totalHa * 100) / 100)} ha{totalKm > 0 ? ` · ${fmtNum(Math.round(totalKm * 10) / 10)} km` : ''}</span>
      </label>
      <p className="text-xs text-stone-400 -mt-1 mb-2">Hectares drive the “percentage done” on the spray and work boards. Organic and in-conversion blocks may only be sprayed with BioGro-certified, approved products.</p>
      <div className="overflow-x-auto border border-stone-200 rounded-xl mb-2">
        <table className="w-full text-sm whitespace-nowrap">
          <thead>
            <tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200 bg-stone-50">
              <th className="px-3 py-2.5 font-semibold">Block name</th>
              <th className="px-3 py-2.5 font-semibold">Hectares</th>
              <th className="px-3 py-2.5 font-semibold">Rows</th>
              <th className="px-3 py-2.5 font-semibold">Row width</th>
              <th className="px-3 py-2.5 font-semibold">Km of vine row</th>
              <th className="px-3 py-2.5 font-semibold">Vines</th>
              <th className="px-3 py-2.5 font-semibold">Certification</th>
              <th className="px-3 py-2.5"></th>
            </tr>
          </thead>
          <tbody>
            {blocks.length === 0 ? (
              <tr><td colSpan={8} className="px-3 py-6 text-center text-stone-400">No blocks yet — add one or paste your list.</td></tr>
            ) : blocks.map((b, i) => (
              <tr key={i} className="border-b border-stone-100 last:border-0">
                <td className="px-2 py-1.5"><input value={b.name} onChange={e => setRow(i, { name: e.target.value })} placeholder="Block name" className={gi + ' w-full min-w-[190px] font-medium'} /></td>
                <td className="px-2 py-1.5"><input value={b.ha ?? ''} onChange={e => setRow(i, { ha: e.target.value })} inputMode="decimal" placeholder="0" className={gi + ' w-24 text-right'} /></td>
                <td className="px-2 py-1.5"><input value={b.rows ?? ''} onChange={e => setRow(i, { rows: e.target.value })} placeholder="e.g. 42-180" className={gi + ' w-36'} /></td>
                <td className="px-2 py-1.5">
                  <input value={b.rowWidth ?? ''} onChange={e => setRow(i, { rowWidth: e.target.value })} inputMode="decimal"
                    placeholder="2.7" title="Metres between rows — leave blank to use the vineyard default"
                    className={gi + ' w-20 text-right'} />
                </td>
                <td className="px-2 py-1.5">
                  <input value={b.km ?? ''} onChange={e => setRow(i, { km: e.target.value })} inputMode="decimal"
                    placeholder={String(blockKm({ ...b, km: '' }, { rowWidth: 2.7 }) || '0')}
                    title="Leave blank to work it out from hectares ÷ row width"
                    className={gi + ' w-24 text-right'} />
                </td>
                <td className="px-2 py-1.5">
                  <input value={b.vines ?? ''} onChange={e => setRow(i, { vines: e.target.value })} inputMode="numeric"
                    placeholder={String(blockVines({ ...b, vines: '' }, { vineSpacing: 1.8 }) || '')}
                    title="Leave blank to work it out from km of vine row" className={gi + ' w-24 text-right'} />
                </td>
                <td className="px-2 py-1.5">
                  <select value={b.cert || ''} onChange={e => setRow(i, { cert: e.target.value })} className={gi + ' w-32'}>
                    {CERTS.map(c => <option key={c} value={c}>{c || '—'}</option>)}
                  </select>
                </td>
                <td className="px-2 py-1.5 text-right"><button onClick={() => onChange(blocks.filter((_, j) => j !== i))} className="p-1.5 rounded-md hover:bg-red-50 text-red-500"><Trash2 size={15} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="flex gap-2">
        <button onClick={() => onChange([...blocks, { name: '', ha: '', rows: '', km: '' }])} className={cls.ghost + ' !py-2 !px-3'}><Plus size={16} /> Add block</button>
        <button onClick={() => setShowPaste(v => !v)} className={cls.ghost + ' !py-2 !px-3'}><Upload size={15} /> Load list</button>
      </div>
      {showPaste && (
        <div className="mt-3 p-3 rounded-lg border border-stone-200 bg-stone-50">
          <p className="text-sm text-stone-600 mb-2">Paste up to five columns: <b>Block name, Hectares, Rows, Km, Row width</b> (all but the name optional).</p>
          <textarea value={paste} onChange={e => setPaste(e.target.value)} rows={4} placeholder={'Hill - A 23\t7.9\t42-180\t26.69\nHill - E PG\t1.55\t18-57\t5.17'} className={cls.input + ' font-mono text-[12px] resize-y'} />
          <div className="flex gap-2 mt-2 justify-end">
            <button onClick={() => load(false)} className={cls.ghost + ' !py-2 !px-3'}>Add to list</button>
            <button onClick={() => load(true)} className={cls.primary + ' !py-2 !px-3'}>Replace list</button>
          </div>
        </div>
      )}
    </div>
  );
}

function SprayTypesEditor({ types, onChange }) {
  const setType = (i, patch) => onChange(types.map((t, j) => (j === i ? { ...t, ...patch } : t)));
  const addLane = i => {
    const t = types[i]; const name = (window.prompt('Sprayer / lane name (e.g. Jason, Weed sprayer)') || '').trim();
    if (!name || t.statuses.includes(name)) return;
    setType(i, { statuses: [...t.statuses, name], laneTanks: { ...t.laneTanks, [name]: '' } });
  };
  const removeLane = (i, lane) => {
    const t = types[i]; const lt = { ...t.laneTanks }; delete lt[lane];
    setType(i, { statuses: t.statuses.filter(s => s !== lane), laneTanks: lt });
  };
  return (
    <div>
      <label className={cls.label}>Spray boards & sprayers</label>
      <p className="text-xs text-stone-400 -mt-1 mb-2">Each board has its own lanes, tank sizes and water rate. Full-tank mixes scale to each tank.</p>
      <div className="space-y-4">
        {types.map((t, i) => {
          const lanes = (t.statuses || []).filter(s => s.toLowerCase() !== 'to spray');
          return (
            <div key={t.key} className="rounded-xl border border-stone-200 p-3">
              <input value={t.label} onChange={e => setType(i, { label: e.target.value })} className={cls.input + ' font-medium mb-3'} />
              <div className="space-y-2">
                {lanes.map(lane => (
                  <div key={lane} className="flex items-center gap-2">
                    <span className="text-sm text-stone-700 flex-1 min-w-0 truncate">{lane}</span>
                    <input value={t.laneTanks[lane] ?? ''} onChange={e => setType(i, { laneTanks: { ...t.laneTanks, [lane]: e.target.value === '' ? '' : numOf(e.target.value) } })}
                      inputMode="numeric" placeholder="litres" className={cls.input + ' !w-28'} />
                    <span className="text-sm text-stone-400">L</span>
                    <button onClick={() => removeLane(i, lane)} className="p-2 rounded-lg hover:bg-red-50 text-red-500 shrink-0"><Trash2 size={15} /></button>
                  </div>
                ))}
                {lanes.length === 0 && <p className="text-sm text-stone-400">No sprayers yet.</p>}
              </div>
              <button onClick={() => addLane(i)} className={cls.ghost + ' !py-2 !px-3 mt-2'}><Plus size={15} /> Add sprayer</button>
              <div className="flex items-center gap-2 mt-3 pt-3 border-t border-stone-100">
                <span className="text-sm text-stone-700 flex-1">Water rate</span>
                <input value={t.waterRate ?? ''} onChange={e => setType(i, { waterRate: e.target.value === '' ? '' : numOf(e.target.value) })} inputMode="numeric" className={cls.input + ' !w-28'} />
                <span className="text-sm text-stone-400">L/ha</span>
              </div>
            </div>
          );
        })}
        {types.length === 0 && <p className="text-sm text-stone-400">No spray boards configured.</p>}
      </div>
    </div>
  );
}

/* Two-column lookup editor — used for block codes and job accounts */
function LookupEditor({ label, hint, map, keys, onChange }) {
  const [showPaste, setShowPaste] = useState(false);
  const [paste, setPaste] = useState('');
  const rows = keys && keys.length ? keys : Object.keys(map || {}).sort();
  const missing = rows.filter(k => !(map || {})[k]).length;
  const load = () => {
    const next = { ...(map || {}) };
    paste.split(/\r?\n/).forEach(line => {
      const parts = line.split('\t').length > 1 ? line.split('\t') : line.split(/\s{2,}/);
      const k = (parts[0] || '').trim(), v = (parts[1] || '').trim();
      if (k && v) next[k] = v;
    });
    onChange(next); setPaste(''); setShowPaste(false);
  };
  const gi = 'px-2 py-1.5 rounded-md border border-stone-200 bg-white text-stone-800 text-[13px] w-full focus:outline-none focus:ring-2 focus:ring-stone-400/40';
  return (
    <div>
      <label className={cls.label + ' flex items-center justify-between'}>
        <span>{label}</span>
        <span className="normal-case tracking-normal text-stone-400">
          {rows.length} rows{missing ? ` · ${missing} with no code` : ''}
        </span>
      </label>
      <p className="text-xs text-stone-400 -mt-1 mb-2">{hint}</p>
      <div className="overflow-x-auto border border-stone-200 rounded-xl mb-2 max-h-80 overflow-y-auto">
        <table className="w-full text-sm">
          <thead className="sticky top-0 bg-stone-50">
            <tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
              <th className="px-3 py-2 font-semibold">In the app</th>
              <th className="px-3 py-2 font-semibold">Goes to the spreadsheet as</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(k => (
              <tr key={k} className={'border-b border-stone-100 last:border-0 ' + (!(map || {})[k] ? 'bg-amber-50/60' : '')}>
                <td className="px-3 py-1.5 text-stone-700 whitespace-nowrap">{k}</td>
                <td className="px-2 py-1.5">
                  <input value={(map || {})[k] || ''} placeholder="— no match —"
                    onChange={e => onChange({ ...(map || {}), [k]: e.target.value })} className={gi} />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <button onClick={() => setShowPaste(v => !v)} className={cls.ghost + ' !py-2 !px-3'}><Upload size={15} /> Paste mappings</button>
      {showPaste && (
        <div className="mt-3 p-3 rounded-lg border border-stone-200 bg-stone-50">
          <p className="text-sm text-stone-600 mb-2">Two columns, tab separated — name on the left, code on the right. Existing rows are updated, new ones added.</p>
          <textarea value={paste} onChange={e => setPaste(e.target.value)} rows={4} className={cls.input + ' font-mono text-[12px] resize-y'} />
          <div className="flex justify-end mt-2"><button onClick={load} className={cls.primary + ' !py-2 !px-3'}>Load</button></div>
        </div>
      )}
    </div>
  );
}

/* Work rate per task — km/h for tractor jobs, plants per hour for hand work */
function PaceEditor({ tasks, pace, vineSpacing, rowWidth, onChange, onSpacing, onRowWidth }) {
  const set = (task, patch) => onChange({ ...(pace || {}), [task]: { type: 'kmh', value: '', headland: 15, ...((pace || {})[task] || {}), ...patch } });
  const gi = 'px-2 py-1.5 rounded-md border border-stone-200 bg-white text-stone-800 text-[13px] focus:outline-none focus:ring-2 focus:ring-stone-400/40';
  const set_ = [...tasks].sort((a, b) => a.localeCompare(b));
  const done = set_.filter(t => numOf(((pace || {})[t] || {}).value) > 0).length;
  return (
    <div>
      <label className={cls.label + ' flex items-center justify-between'}>
        <span>Work rates</span>
        <span className="normal-case tracking-normal text-stone-400">{done} of {set_.length} tasks have a rate</span>
      </label>
      <p className="text-xs text-stone-400 -mt-1 mb-2">
        Used to work out how long a job should take. Tractor jobs use km/h across the block's vine rows plus a turning
        allowance for the headlands; hand jobs use plants per hour.
      </p>
      <div className="flex items-end gap-2 mb-2">
        <div><label className="text-[10px] uppercase tracking-wide text-stone-400 block">Vine spacing</label>
          <div className="flex items-center gap-1.5">
            <input value={vineSpacing ?? ''} onChange={e => onSpacing(e.target.value)} inputMode="decimal" className={gi + ' w-20 text-right'} />
            <span className="text-sm text-stone-500">m</span>
          </div></div>
        <div><label className="text-[10px] uppercase tracking-wide text-stone-400 block">Default row width</label>
          <div className="flex items-center gap-1.5">
            <input value={rowWidth ?? ''} onChange={e => onRowWidth(e.target.value)} inputMode="decimal" className={gi + ' w-20 text-right'} />
            <span className="text-sm text-stone-500">m</span>
          </div></div>
        <p className="text-xs text-stone-400 pb-2">Used for blocks that don't have their own figures. Row width sets the km of vine row per hectare.</p>
      </div>
      <div className="overflow-x-auto border border-stone-200 rounded-xl max-h-96 overflow-y-auto">
        <table className="w-full text-sm">
          <thead className="sticky top-0 bg-stone-50">
            <tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
              <th className="px-3 py-2 font-semibold">Task</th>
              <th className="px-3 py-2 font-semibold">Measured by</th>
              <th className="px-3 py-2 font-semibold">Rate</th>
              <th className="px-3 py-2 font-semibold">Headland %</th>
            </tr>
          </thead>
          <tbody>
            {set_.map(t => {
              const pc = (pace || {})[t] || {};
              const isKm = (pc.type || 'kmh') === 'kmh';
              return (
                <tr key={t} className={'border-b border-stone-100 last:border-0 ' + (numOf(pc.value) > 0 ? '' : 'bg-amber-50/40')}>
                  <td className="px-3 py-1.5 text-stone-700 whitespace-nowrap">{t}</td>
                  <td className="px-2 py-1.5">
                    <select value={pc.type || 'kmh'} onChange={e => set(t, { type: e.target.value })} className={gi + ' w-32'}>
                      <option value="kmh">km/h (tractor)</option>
                      <option value="plants">plants/hour</option>
                    </select>
                  </td>
                  <td className="px-2 py-1.5">
                    <div className="flex items-center gap-1.5">
                      <input value={pc.value ?? ''} onChange={e => set(t, { value: e.target.value })} inputMode="decimal" className={gi + ' w-20 text-right'} />
                      <span className="text-[12px] text-stone-400 w-16">{isKm ? 'km/h' : 'plants/h'}</span>
                    </div>
                  </td>
                  <td className="px-2 py-1.5">
                    <input value={isKm ? (pc.headland ?? '') : ''} disabled={!isKm}
                      onChange={e => set(t, { headland: e.target.value })} inputMode="decimal"
                      className={gi + ' w-20 text-right disabled:opacity-30'} />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* Editor for the vineyard work task list (the buttons on the Work planner) */
function WorkTasksEditor({ tasks, onChange }) {
  const [showPaste, setShowPaste] = useState(false);
  const [paste, setPaste] = useState('');
  const [nt, setNt] = useState('');
  const add = () => { const v = nt.trim(); if (!v || tasks.includes(v)) { setNt(''); return; } onChange([...tasks, v]); setNt(''); };
  const load = replace => {
    const lines = paste.split(/\r?\n/).map(l => l.split('\t')[0].trim()).filter(Boolean);
    if (!lines.length) return;
    const merged = replace ? lines : [...tasks, ...lines.filter(l => !tasks.includes(l))];
    onChange(merged); setPaste(''); setShowPaste(false);
  };
  return (
    <div>
      <label className={cls.label + ' flex items-center justify-between'}>
        <span>Vineyard work tasks</span>
        <span className="normal-case tracking-normal text-stone-400">{tasks.length} tasks</span>
      </label>
      <p className="text-xs text-stone-400 -mt-1 mb-2">These are the buttons on the Work planner. Timesheet tasks are separate — they're the two lists below.</p>
      <div className="flex flex-wrap gap-1.5 mb-2">
        {tasks.map((t, i) => (
          <span key={i} className="inline-flex items-center gap-1 pl-3 pr-1.5 py-1.5 rounded-lg border border-stone-300 bg-white text-sm text-stone-700">
            <input value={t} onChange={e => onChange(tasks.map((x, j) => (j === i ? e.target.value : x)))}
              className="bg-transparent border-0 p-0 focus:outline-none text-sm text-stone-800"
              style={{ width: `${Math.max(6, t.length) * 7.4}px` }} />
            <button onClick={() => onChange(tasks.filter((_, j) => j !== i))} className="p-1 rounded hover:bg-red-50 text-red-400 shrink-0"><X size={13} /></button>
          </span>
        ))}
        {tasks.length === 0 && <p className="text-sm text-stone-400">No work tasks yet.</p>}
      </div>
      <div className="flex items-center gap-2">
        <input value={nt} onChange={e => setNt(e.target.value)} onKeyDown={e => e.key === 'Enter' && add()} placeholder="New work task" className={cls.input} />
        <button onClick={add} className={cls.primary + ' shrink-0'}><Plus size={16} /></button>
        <button onClick={() => setShowPaste(v => !v)} className={cls.ghost + ' !py-2 !px-3 shrink-0'}><Upload size={15} /> Load list</button>
      </div>
      {showPaste && (
        <div className="mt-3 p-3 rounded-lg border border-stone-200 bg-stone-50">
          <p className="text-sm text-stone-600 mb-2">Paste one task per line.</p>
          <textarea value={paste} onChange={e => setPaste(e.target.value)} rows={4} placeholder={'Mulching\nShoot thin\nBud Rub'} className={cls.input + ' font-mono text-[12px] resize-y'} />
          <div className="flex gap-2 mt-2 justify-end">
            <button onClick={() => load(false)} className={cls.ghost + ' !py-2 !px-3'}>Add to list</button>
            <button onClick={() => load(true)} className={cls.primary + ' !py-2 !px-3'}>Replace list</button>
          </div>
        </div>
      )}
    </div>
  );
}

function TasksEditor({ tasks, onChange, label }) {
  const [showPaste, setShowPaste] = useState(false);
  const [paste, setPaste] = useState('');
  const [order, setOrder] = useState('code'); // 'code' = Code,Task ; 'task' = Task,Code
  const setRow = (i, patch) => onChange(tasks.map((t, j) => (j === i ? { ...t, ...patch } : t)));
  const load = replace => {
    const aoa = Papa.parse(paste.replace(/\r/g, ''), { skipEmptyLines: true }).data;
    const parsed = [];
    aoa.forEach((row, idx) => {
      if (!row || !row.length) return;
      const cells = row.map(c => String(c == null ? '' : c).trim());
      const joined = cells.join(' ').toLowerCase();
      if (idx === 0 && joined.includes('code') && /task|job|name/.test(joined)) return; // skip header row
      let code = '', name = '';
      if (cells.length >= 2) {
        if (order === 'code') { code = cells[0]; name = cells.slice(1).join(' '); }
        else { name = cells.slice(0, -1).join(' '); code = cells[cells.length - 1]; }
      } else { name = cells[0]; }
      if (name) parsed.push({ code, name });
    });
    if (!parsed.length) return;
    onChange(replace ? parsed : [...tasks, ...parsed]);
    setPaste(''); setShowPaste(false);
  };
  return (
    <div>
      <label className={cls.label}>{label || 'Tasks'}</label>
      <p className="text-xs text-stone-400 -mt-1 mb-2">Operators see the task name only. The code is exported in its own column on the timesheet.</p>
      <div className="space-y-2 mb-2">
        {tasks.map((t, i) => (
          <div key={i} className="flex items-center gap-2">
            <input value={t.code} onChange={e => setRow(i, { code: e.target.value })} placeholder="Code" className={cls.input + ' !w-24 font-mono uppercase text-center'} />
            <input value={t.name} onChange={e => setRow(i, { name: e.target.value })} placeholder="Task name" className={cls.input} />
            <button onClick={() => onChange(tasks.filter((_, j) => j !== i))} className="p-2.5 rounded-lg hover:bg-red-50 text-red-500 shrink-0"><Trash2 size={16} /></button>
          </div>
        ))}
        {tasks.length === 0 && <p className="text-sm text-stone-400">No tasks yet.</p>}
      </div>
      <div className="flex gap-2">
        <button onClick={() => onChange([...tasks, { code: '', name: '' }])} className={cls.ghost + ' !py-2 !px-3'}><Plus size={16} /> Add task</button>
        <button onClick={() => setShowPaste(v => !v)} className={cls.ghost + ' !py-2 !px-3'}><Upload size={15} /> Load list</button>
      </div>
      {showPaste && (
        <div className="mt-3 p-3 rounded-lg border border-stone-200 bg-stone-50">
          <div className="flex items-center gap-2 mb-2 text-sm text-stone-600">
            Paste two columns:
            <select value={order} onChange={e => setOrder(e.target.value)} className="px-2 py-1 rounded border border-stone-300 bg-white text-sm">
              <option value="code">Code, Task</option>
              <option value="task">Task, Code</option>
            </select>
          </div>
          <textarea value={paste} onChange={e => setPaste(e.target.value)} rows={4}
            placeholder={'PRU\tPruning\nWRP\tWrapping\nSPR\tSpraying'}
            className={cls.input + ' font-mono text-[12px] resize-y'} />
          <div className="flex gap-2 mt-2 justify-end">
            <button onClick={() => load(false)} className={cls.ghost + ' !py-2 !px-3'}>Add to list</button>
            <button onClick={() => load(true)} className={cls.primary + ' !py-2 !px-3'}>Replace list</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* Backup & restore — a full snapshot of everything, so nothing is ever
   one mistake away from being lost. */
function BackupPanel() {
  const [msg, setMsg] = useState('');
  const [busy, setBusy] = useState(false);
  const fileRef = useRef(null);

  const allKeys = async () => {
    const cfg = await loadJSON(K.config, DEFAULT_CONFIG);
    const keys = [K.config, K.work, K.workDone, K.spraysLegacy];
    (cfg.sprayTypes || []).forEach(t => keys.push(K.sprays(t.key)));
    (cfg.operators || []).forEach(o => { keys.push(K.ts(o.code)); keys.push(K.hz(o.code)); });
    keys.push(K.ts(cfg.managerCode || '0000'), K.hz(cfg.managerCode || '0000'));
    return [...new Set(keys)];
  };

  const backup = async () => {
    setBusy(true);
    try {
      const keys = await allKeys();
      const data = {};
      for (const k of keys) {
        const v = await loadJSON(k, null);
        if (v !== null) data[k] = v;
      }
      const payload = { app: 'vineyard-ops', savedAt: new Date().toISOString(), data };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = `vineyard-backup_${todayStr()}.json`;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setMsg(`Backup saved — ${Object.keys(data).length} records.`);
    } catch (e) { setMsg('Backup failed: ' + (e.message || e)); }
    setBusy(false);
    setTimeout(() => setMsg(''), 6000);
  };

  const restore = file => {
    const reader = new FileReader();
    reader.onload = async e => {
      setBusy(true);
      try {
        const payload = JSON.parse(String(e.target.result));
        const data = payload && payload.data;
        if (!data || typeof data !== 'object') throw new Error('not a Vineyard Ops backup');
        const count = Object.keys(data).length;
        if (!window.confirm(`Restore ${count} records from this backup?\n\nAnything currently in the app with the same name will be replaced by the backup's version.`)) { setBusy(false); return; }
        for (const [k, v] of Object.entries(data)) await saveJSON(k, v);
        setMsg(`Restored ${count} records. Reloading…`);
        setTimeout(() => window.location.reload(), 900);
      } catch (err) { setMsg('Could not read that file: ' + (err.message || err)); setBusy(false); }
    };
    reader.readAsText(file);
  };

  return (
    <div className={cls.card + ' p-4'}>
      <div className="flex items-center gap-2 mb-2"><Download size={16} className="text-stone-500" /><h3 className="font-semibold text-stone-900">Backup & restore</h3></div>
      <p className="text-sm text-stone-500 mb-3">
        Saves everything — operators, blocks, tasks, spray boards, work, timesheets and hazards — into one file.
        Worth doing before any app update, and every so often anyway.
      </p>
      <div className="flex gap-2 flex-wrap">
        <button onClick={backup} disabled={busy} className={cls.primary + ' !py-2 !px-3'}><Download size={15} /> Download backup</button>
        <input ref={fileRef} type="file" accept=".json" className="hidden"
          onChange={e => e.target.files[0] && restore(e.target.files[0])} />
        <button onClick={() => fileRef.current?.click()} disabled={busy} className={cls.ghost + ' !py-2 !px-3'}><Upload size={15} /> Restore from backup</button>
      </div>
      {msg && <div className="mt-3 text-sm px-3 py-2 rounded-lg bg-stone-100 text-stone-700 border border-stone-200">{msg}</div>}
    </div>
  );
}

function Setup({ config, onSave }) {
  const [openTasks, setOpenTasks] = useState(null);   // which operator's task picker is open
  const [draft, setDraft] = useState(JSON.parse(JSON.stringify(config)));
  const [msg, setMsg] = useState('');
  const [nc, setNc] = useState(''); const [nn, setNn] = useState('');
  const dirty = JSON.stringify(draft) !== JSON.stringify(config);

  const save = async () => { await onSave(draft); setMsg('Settings saved.'); setTimeout(() => setMsg(''), 3000); };
  const addOp = () => {
    if (!nc.trim() || !nn.trim()) return;
    if (draft.operators.some(o => o.code === nc.trim()) || nc.trim() === draft.managerCode || nc.trim() === draft.techCode) { setMsg('That code is already in use.'); return; }
    setDraft({ ...draft, operators: [...draft.operators, { code: nc.trim(), name: nn.trim() }] }); setNc(''); setNn('');
  };

  return (
    <div className="space-y-6 pb-24">
      <Banner msg={msg} />
      <div className="flex items-center gap-2 text-stone-700"><Settings size={18} /><h2 className="text-lg font-semibold text-stone-900">Setup</h2></div>

      <BackupPanel />

      <div className={cls.card + ' p-4 space-y-3.5'}>
        <h3 className="font-semibold text-stone-900">General</h3>
        <div className="grid sm:grid-cols-2 gap-3.5">
          <div><label className={cls.label}>App name</label><input value={draft.siteName} onChange={e => setDraft({ ...draft, siteName: e.target.value })} className={cls.input} /></div>
          <div><label className={cls.label}>Manager code</label><input value={draft.managerCode} onChange={e => setDraft({ ...draft, managerCode: e.target.value.replace(/\D/g, '') })} inputMode="numeric" className={cls.input} /></div>
          <div className="flex gap-3">
            <div className="flex-1"><label className={cls.label}>Technical Viticulturist</label>
              <input value={draft.techName || ''} onChange={e => setDraft({ ...draft, techName: e.target.value })} placeholder="Name" className={cls.input} /></div>
            <div className="w-32"><label className={cls.label}>Their code</label>
              <input value={draft.techCode || ''} onChange={e => setDraft({ ...draft, techCode: e.target.value.replace(/\D/g, '') })} inputMode="numeric" className={cls.input + ' font-mono text-center'} /></div>
          </div>
          <p className="text-xs text-stone-400 -mt-1">The Technical Viticulturist gets their own console — team timeline, E-L stages and disease monitoring.</p>
        </div>
      </div>

      <div className={cls.card + ' p-4'}>
        <div className="flex items-center gap-2 mb-3"><Users size={16} className="text-stone-500" /><h3 className="font-semibold text-stone-900">Operators</h3></div>
        <div className="space-y-2 mb-4">
          {draft.operators.map((o, i) => (
            <div key={i}>
            <div className="flex items-center gap-2">
              <input value={o.code} onChange={e => { const ops = [...draft.operators]; ops[i] = { ...o, code: e.target.value.replace(/\D/g, '') }; setDraft({ ...draft, operators: ops }); }}
                inputMode="numeric" className={cls.input + ' !w-24 font-mono text-center'} />
              <input value={o.name} onChange={e => { const ops = [...draft.operators]; ops[i] = { ...o, name: e.target.value }; setDraft({ ...draft, operators: ops }); }} className={cls.input} />
              <button onClick={() => setOpenTasks(openTasks === i ? null : i)}
                title="Choose which tasks this person sees on their timesheet"
                className={'px-3 py-2.5 rounded-lg border text-sm whitespace-nowrap shrink-0 ' +
                  (openTasks === i ? 'bg-stone-900 border-stone-900 text-stone-50' : 'bg-white border-stone-300 text-stone-700 hover:bg-stone-50')}>
                {Array.isArray(o.tasks) && o.tasks.length ? `${o.tasks.length} tasks` : 'All tasks'}
              </button>
              <button onClick={() => setDraft({ ...draft, operators: draft.operators.filter((_, j) => j !== i) })} className="p-2.5 rounded-lg hover:bg-red-50 text-red-500 shrink-0"><Trash2 size={16} /></button>
            </div>
            {openTasks === i && (() => {
              const allTasks = draft.jobs || [];
              const sel = Array.isArray(o.tasks) ? o.tasks : [];
              const isAll = sel.length === 0;
              const setSel = next => {
                const ops = [...draft.operators];
                ops[i] = { ...o, tasks: next.length === allTasks.length ? [] : next };   // all selected = no restriction
                setDraft({ ...draft, operators: ops });
              };
              const toggle = name => setSel(sel.includes(name) ? sel.filter(x => x !== name) : [...sel, name]);
              return (
                <div className="mt-2 mb-1 p-3 rounded-xl border border-stone-300 bg-stone-50">
                  <div className="flex items-center justify-between gap-2 mb-2 flex-wrap">
                    <span className="text-[13px] text-stone-600">
                      Tasks <b className="text-stone-900">{o.name || 'this operator'}</b> can pick on their timesheet
                      {isAll && <span className="text-stone-400"> — all {allTasks.length} (no restriction)</span>}
                    </span>
                    <div className="flex gap-2">
                      <button onClick={() => setSel(allTasks.map(t => t.name))} className="text-xs px-2.5 py-1.5 rounded-lg border border-stone-300 bg-white text-stone-600 hover:bg-stone-100">All</button>
                      <button onClick={() => setSel([])} className="text-xs px-2.5 py-1.5 rounded-lg border border-stone-300 bg-white text-stone-600 hover:bg-stone-100">Clear</button>
                    </div>
                  </div>
                  <div className="grid gap-1 sm:grid-cols-2 lg:grid-cols-3 max-h-72 overflow-auto">
                    {allTasks.map(t => {
                      const on = isAll || sel.includes(t.name);
                      return (
                        <label key={t.name} className={'flex items-center gap-2 px-2.5 py-1.5 rounded-lg border cursor-pointer text-[13px] ' +
                          (on ? 'bg-white border-stone-300 text-stone-800' : 'bg-stone-100/60 border-stone-200 text-stone-400')}>
                          <input type="checkbox" checked={on} onChange={() => toggle(t.name)} className="w-4 h-4 accent-stone-800 shrink-0" />
                          <span className="truncate" title={t.name}>{t.name}</span>
                          {t.code && <span className="ml-auto text-[11px] text-stone-400 font-mono shrink-0">{t.code}</span>}
                        </label>
                      );
                    })}
                  </div>
                </div>
              );
            })()}
            </div>
          ))}
          {draft.operators.length === 0 && <p className="text-sm text-stone-400">No operators yet.</p>}
        </div>
        <div className="flex items-center gap-2 pt-3 border-t border-stone-100">
          <input value={nc} onChange={e => setNc(e.target.value.replace(/\D/g, ''))} placeholder="Code" inputMode="numeric" className={cls.input + ' !w-24 font-mono text-center'} />
          <input value={nn} onChange={e => setNn(e.target.value)} placeholder="Operator name" className={cls.input} onKeyDown={e => e.key === 'Enter' && addOp()} />
          <button onClick={addOp} className={cls.primary + ' shrink-0'}><Plus size={16} /></button>
        </div>
        <p className="text-xs text-stone-400 mt-2">Each operator signs in with their code. Codes must be unique. Tap the tasks button beside anyone to choose exactly which timesheet tasks they see — leave it on "All tasks" for full access.</p>
      </div>

      <div className={cls.card + ' p-4 space-y-3.5'}>
        <div className="flex items-center gap-2"><AlertTriangle size={16} className="text-stone-500" /><h3 className="font-semibold text-stone-900">Hazard email alerts</h3></div>
        <p className="text-sm text-stone-500 -mt-1">Paste the Web app URL from your Apps Script (or a Zapier/Make webhook). Leave blank to turn email alerts off — hazards still appear in the console either way.</p>
        <div><label className={cls.label}>Webhook URL</label><input value={draft.webhookUrl} onChange={e => setDraft({ ...draft, webhookUrl: e.target.value.trim() })} className={cls.input + ' font-mono text-[13px]'} placeholder="https://script.google.com/macros/s/…/exec" /></div>
        <div><label className={cls.label}>Notify email(s)</label><input value={draft.notifyEmail} onChange={e => setDraft({ ...draft, notifyEmail: e.target.value })} className={cls.input} inputMode="email" placeholder="you@winery.co.nz, hs@winery.co.nz" /></div>
      </div>

      <div className={cls.card + ' p-4 space-y-5'}>
        <BlocksEditor blocks={draft.blocks} onChange={v => setDraft({ ...draft, blocks: v })} />
        <WorkTasksEditor tasks={draft.workTasks || []} onChange={v => setDraft({ ...draft, workTasks: v })} />
        <LookupEditor label="Machines used by each task"
          hint="When someone works one of these tasks, the hours go on that machine's clock — so mowing builds hours on the mower, mulching on the mulcher. Leave blank if a task uses no machine of its own."
          map={draft.taskMachines || {}} keys={(draft.workTasks || []).slice().sort((a, b) => a.localeCompare(b))}
          onChange={v => setDraft({ ...draft, taskMachines: v })} />
        <PaceEditor tasks={draft.workTasks || []} pace={draft.workPace || {}} vineSpacing={draft.vineSpacing} rowWidth={draft.rowWidth}
          onChange={v => setDraft({ ...draft, workPace: v })}
          onSpacing={v => setDraft({ ...draft, vineSpacing: v })}
          onRowWidth={v => setDraft({ ...draft, rowWidth: v })} />
        <TasksEditor tasks={draft.jobs} onChange={v => setDraft({ ...draft, jobs: v })} label="Timesheet tasks (all operators)" />
        <LookupEditor label="Block codes for the timesheet export"
          hint="Each block's name in the app, and the code payroll expects in the Block Code column."
          map={draft.blockCodes || {}} keys={(draft.blocks || []).map(b => b.name).filter(n => n !== 'N/A')}
          onChange={v => setDraft({ ...draft, blockCodes: v })} />
        <LookupEditor label="Accounts for the timesheet export"
          hint="Each timesheet task and the account it books to. Fills the Account column."
          map={draft.jobAccounts || {}} keys={(draft.jobs || []).map(j => j.name)}
          onChange={v => setDraft({ ...draft, jobAccounts: v })} />
        <SprayTypesEditor types={draft.sprayTypes || []} onChange={v => setDraft({ ...draft, sprayTypes: v })} />
      </div>

      <div className="fixed bottom-0 inset-x-0 border-t border-stone-300 px-4 py-3" style={{ backgroundColor: CREAM }}>
        <div className="max-w-[1800px] mx-auto flex items-center justify-between gap-3">
          <span className="text-sm text-stone-500">{dirty ? 'You have unsaved changes.' : 'All changes saved.'}</span>
          <div className="flex gap-2">
            <button onClick={() => setDraft(JSON.parse(JSON.stringify(config)))} disabled={!dirty} className={cls.ghost + ' disabled:opacity-40'}>Discard</button>
            <button onClick={save} disabled={!dirty} className={cls.primary}>Save changes</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   Operator shell
   ============================================================ */
/* ============================================================
   Fuel log + machine hours + R&M scheduling
   ============================================================ */
// current engine hours = the highest reading operators have entered
function hoursOf(vehicleName, hoursLog) {
  let h = 0;
  (hoursLog || []).forEach(r => { if (r.vehicle === vehicleName) h = Math.max(h, numOf(r.hours)); });
  return h;
}
// when that reading was last taken
function hoursAsOf(vehicleName, hoursLog) {
  const rows = (hoursLog || []).filter(r => r.vehicle === vehicleName).sort((a, b) => (b.ts || 0) - (a.ts || 0));
  return rows.length ? rows[0] : null;
}
/* REGO and WOF expiry. Warns a month out, and flags anything past its date. */
const EXPIRY_WARN_DAYS = 30;
function expiryStatus(iso) {
  if (!iso) return null;
  const [y, m, d] = String(iso).split('-').map(Number);
  if (!y) return null;
  const due = new Date(y, m - 1, d);
  const today = startOfDay(nzNow());
  const days = Math.round((due - today) / 86400000);
  return {
    iso, due, days,
    label: `${pad2(d)}/${pad2(m)}/${y}`,
    expired: days < 0,
    soon: days >= 0 && days <= EXPIRY_WARN_DAYS,
  };
}
// anything on the fleet that's expired or about to
function expiringItems(vehicles) {
  const out = [];
  (vehicles || []).forEach(v => {
    [['REGO', v.rego], ['WOF', v.wof]].forEach(([what, iso]) => {
      const st = expiryStatus(iso);
      if (st && (st.expired || st.soon)) out.push({ vehicle: v.name, what, ...st });
    });
  });
  return out.sort((a, b) => a.days - b.days);
}

// weekly check / hour service status for one vehicle
/* Two separate clocks per machine:
   • SERVICE — a workshop job, scheduled on engine hours run.
   • CHECKLIST — the operator's own inspection, scheduled on days.
   Engine hours come either from what operators type in, or (for implements
   like the mower and mulcher) from the hours worked on the tasks that use them. */

// hours a machine has accumulated through work sessions on its tasks
function implementHours(vehicleName, workCards, config) {
  const map = (config || {}).taskMachines || {};
  let h = 0;
  (workCards || []).forEach(c => {
    if (map[c.task] !== vehicleName) return;
    h += cardWorkedHours(c);
  });
  return Math.round(h * 100) / 100;
}
// the machine's current engine hours, whichever way it's tracked
function machineHours(v, hoursLog, workCards, config) {
  if (!v) return 0;
  if (v.hoursSource === 'tasks') {
    return Math.round((numOf(v.startHours) + implementHours(v.name, workCards, config)) * 100) / 100;
  }
  return hoursOf(v.name, hoursLog);
}

// next workshop service, on hours
function serviceStatus(v, hours, rmLog) {
  const every = numOf(v.serviceEveryHours) || 0;
  const last = (rmLog || []).filter(r => r.vehicle === v.name && r.kind === 'Service')
    .sort((a, b) => (b.ts || 0) - (a.ts || 0))[0];
  const base = last ? numOf(last.hours) : numOf(v.lastServiceHours);
  if (!every) return { tracked: false, detail: 'no service interval set', lastAt: last ? last.date : '', base };
  const nextAt = base + every;                 // e.g. serviced at 1500, every 500 → due at 2000
  const remaining = Math.round((nextAt - hours) * 10) / 10;
  return {
    tracked: true, base, nextAt, remaining, every,
    run: Math.max(0, Math.round((hours - base) * 10) / 10),
    due: remaining <= 0,
    soon: remaining > 0 && remaining <= Math.max(10, every * 0.1),
    detail: remaining <= 0 ? `overdue by ${fmtNum(Math.abs(remaining))} h` : `${fmtNum(remaining)} h to go`,
    lastAt: last ? last.date : '', lastBy: last ? last.by : '', lastHours: last ? numOf(last.hours) : numOf(v.lastServiceHours),
  };
}

// operator checklist, on days
function checkStatus(v, rmLog) {
  const every = numOf(v.checkEveryDays) || 0;
  const last = (rmLog || []).filter(r => r.vehicle === v.name && r.kind === 'Checklist')
    .sort((a, b) => (b.ts || 0) - (a.ts || 0))[0];
  const days = last ? Math.floor((Date.now() - last.ts) / 86400000) : null;
  if (!every) return { tracked: false, days, lastAt: last ? last.date : '', detail: 'no checklist interval set' };
  const remaining = days == null ? -1 : every - days;
  return {
    tracked: true, every, days, remaining,
    due: days == null || remaining <= 0,
    soon: remaining === 1,
    detail: days == null ? 'never done'
      : remaining <= 0 ? `due — last done ${days} day${days === 1 ? '' : 's'} ago`
      : `due in ${remaining} day${remaining === 1 ? '' : 's'}`,
    lastAt: last ? last.date : '', lastBy: last ? last.by : '',
  };
}

function FuelForm({ config, session }) {
  const [date, setDate] = useState(todayStr());
  const [tank, setTank] = useState('');
  const [vehicle, setVehicle] = useState('');
  const [meterStart, setMeterStart] = useState('');
  const [meterEnd, setMeterEnd] = useState('');
  const [override, setOverride] = useState('');     // only if the meter can't be used
  const [log, setLog] = useState([]);
  const [msg, setMsg] = useState('');
  const load = async () => setLog(await loadJSON(K.fuel, []));
  useEffect(() => { load(); }, []);
  useLiveKey(K.fuel, v => setLog(v || []));

  const tanks = config.fuelTanks || [];
  // each tank has its own pump meter — carry its last reading into the start
  const lastForTank = t => (log || []).filter(f => f.tank === t).sort((a, b) => (b.ts || 0) - (a.ts || 0))[0];
  useEffect(() => {
    if (!tank) return;
    const last = lastForTank(tank);
    setMeterStart(last && last.meterEnd !== '' && last.meterEnd != null ? String(last.meterEnd) : '');
  }, [tank, log]);

  const litres = meterStart !== '' && meterEnd !== ''
    ? Math.round((numOf(meterEnd) - numOf(meterStart)) * 100) / 100
    : numOf(override);
  const badMeter = meterStart !== '' && meterEnd !== '' && numOf(meterEnd) < numOf(meterStart);
  const valid = date && tank && vehicle && litres > 0 && !badMeter;

  const save = async () => {
    if (!valid) return;
    const entry = {
      id: uid(), date, tank, vehicle, litres,
      meterStart: meterStart === '' ? '' : numOf(meterStart),
      meterEnd: meterEnd === '' ? '' : numOf(meterEnd),
      by: session ? session.name : 'Manager', ts: Date.now(),
    };
    const fresh = await loadJSON(K.fuel, []);
    const next = [entry, ...fresh];
    await saveJSON(K.fuel, next); setLog(next);
    setMsg(`${fmtNum(litres)} L from ${tank} into ${vehicle}.`);
    setMeterStart(String(entry.meterEnd || '')); setMeterEnd(''); setOverride('');
    setTimeout(() => setMsg(''), 4000);
  };

  const recent = (log || []).slice(0, 8);
  return (
    <div className="space-y-4 pb-6">
      <div className="flex items-center gap-2 text-stone-700"><Fuel size={18} /><h2 className="text-lg font-semibold text-stone-900">Diesel fill</h2></div>
      {msg && <div className="text-sm px-3 py-2.5 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200">{msg}</div>}

      <div className={cls.card + ' p-4 space-y-3.5'}>
        <div><label className={cls.label}>Date</label><input type="date" value={date} onChange={e => setDate(e.target.value)} className={cls.input} /></div>

        <div><label className={cls.label}>Which tank</label>
          <div className="grid grid-cols-2 gap-2">
            {tanks.map(t => (
              <button key={t} onClick={() => setTank(t)}
                className={'px-3 py-3 rounded-lg border text-sm font-medium transition-colors ' +
                  (tank === t ? 'bg-stone-900 border-stone-900 text-stone-50' : 'bg-white border-stone-300 text-stone-700 hover:border-stone-400')}>
                {t}
              </button>
            ))}
          </div>
        </div>

        <div><label className={cls.label}>Vehicle being filled</label>
          <select value={vehicle} onChange={e => setVehicle(e.target.value)} className={cls.input}>
            <option value="">Choose a vehicle…</option>
            {(config.vehicles || []).map(v => <option key={v.name} value={v.name}>{v.name}</option>)}
          </select></div>

        <div>
          <label className={cls.label}>Tank meter</label>
          <div className="flex gap-3">
            <div className="flex-1">
              <input value={meterStart} onChange={e => setMeterStart(e.target.value)} inputMode="decimal" placeholder="start" className={cls.input} />
              <p className="text-[11px] text-stone-400 mt-1">Start {tank && lastForTank(tank) ? '— last reading on this tank' : ''}</p>
            </div>
            <div className="flex-1">
              <input value={meterEnd} onChange={e => setMeterEnd(e.target.value)} inputMode="decimal" placeholder="end" className={cls.input + ' text-lg'} />
              <p className="text-[11px] text-stone-400 mt-1">End — after filling</p>
            </div>
          </div>
        </div>

        <div className={'rounded-lg border px-3 py-3 flex items-center justify-between ' + (badMeter ? 'bg-red-50 border-red-300' : 'bg-stone-50 border-stone-200')}>
          <span className="text-sm text-stone-600">Diesel used</span>
          <span className={'text-2xl font-bold tabular-nums ' + (badMeter ? 'text-red-700' : 'text-stone-900')}>
            {badMeter ? 'check meter' : `${fmtNum(litres || 0)} L`}
          </span>
        </div>
        {badMeter && <p className="text-sm text-red-600 -mt-1">The end reading is lower than the start.</p>}

        {meterStart === '' && meterEnd === '' && (
          <div><label className={cls.label}>Or enter litres directly</label>
            <input value={override} onChange={e => setOverride(e.target.value)} inputMode="decimal" placeholder="litres" className={cls.input} />
            <p className="text-[11px] text-stone-400 mt-1">Use this only if the tank has no meter.</p></div>
        )}

        <button onClick={save} disabled={!valid} className={cls.primary + ' w-full !py-3.5 text-base'}><Check size={18} /> Save fill</button>
      </div>

      {recent.length > 0 && (
        <div>
          <h3 className="text-base font-semibold text-stone-900 mb-2">Recent fills</h3>
          <div className="space-y-2">
            {recent.map(f => (
              <div key={f.id} className="flex items-center gap-3 bg-white border border-stone-200 rounded-lg px-3.5 py-2.5">
                <div className="min-w-0">
                  <div className="font-semibold text-stone-900 text-[15px] truncate">{f.vehicle}</div>
                  <div className="text-[12px] text-stone-400">{f.date} · {f.tank || 'tank not set'} · {f.by}</div>
                </div>
                <div className="ml-auto text-right shrink-0">
                  <div className="font-semibold text-stone-900 tabular-nums">{fmtNum(f.litres)} L</div>
                  {f.meterEnd !== '' && <div className="text-[12px] text-stone-400 tabular-nums">meter {fmtNum(f.meterEnd)}</div>}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function MachineChecks({ config, session }) {
  const [rm, setRm] = useState([]);
  const [hoursLog, setHoursLog] = useState([]);
  const [workCards, setWorkCards] = useState([]);
  const [active, setActive] = useState(null);     // machine being checked
  const [results, setResults] = useState({});
  const [hours, setHours] = useState('');
  const [notes, setNotes] = useState('');
  const [hoursFor, setHoursFor] = useState(null); // machine having hours updated
  const [hoursVal, setHoursVal] = useState('');
  const [msg, setMsg] = useState('');
  const [showAll, setShowAll] = useState(false);

  const load = async () => { setRm(await loadJSON(K.rm, [])); setHoursLog(await loadJSON(K.hours, [])); setWorkCards(await loadJSON(K.work, [])); };
  useEffect(() => { load(); }, []);
  useLiveKey(K.rm, v => setRm(v || []));
  useLiveKey(K.hours, v => setHoursLog(v || []));
  useLiveKey(K.work, v => setWorkCards(v || []));

  const items = config.checklist || [];
  const all = config.vehicles || [];
  const mine = all.filter(v => Array.isArray(v.assignedTo) && v.assignedTo.includes(session.code));
  const others = all.filter(v => !mine.includes(v));
  const shown = showAll ? [...mine, ...others] : mine;

  const saveHours = async () => {
    const v = hoursFor; const val = numOf(hoursVal);
    if (!v || val <= 0) return;
    const current = machineHours(v, hoursLog, workCards, config);
    if (val < current && !window.confirm(`That's lower than the last reading of ${fmtNum(current)} h. Save anyway?`)) return;
    const entry = { id: uid(), vehicle: v.name, hours: val, date: todayNZ(), time: nowTimeNZ(), ts: Date.now(), by: session.name };
    const next = [entry, ...(hoursLog || [])];
    await saveJSON(K.hours, next); setHoursLog(next);
    setMsg(`${v.name} — ${fmtNum(val)} hours recorded.`);
    setHoursFor(null); setHoursVal('');
    setTimeout(() => setMsg(''), 4000);
  };

  const submitCheck = async () => {
    const v = active; if (!v) return;
    const failed = items.filter(it => results[it] === 'bad');
    const entry = {
      id: uid(), vehicle: v.name, kind: 'Checklist',
      date: todayNZ(), time: nowTimeNZ(), ts: Date.now(),
      hours: numOf(hours) || machineHours(v, hoursLog, workCards, config), by: session.name,
      note: [notes.trim(), failed.length ? `Needs attention: ${failed.join(', ')}` : ''].filter(Boolean).join(' — '),
      items: { ...results },
    };
    const nextRm = [entry, ...(rm || [])];
    await saveJSON(K.rm, nextRm); setRm(nextRm);
    // a new hours reading counts as this week's entry too
    if (numOf(hours) > 0) {
      const h = { id: uid(), vehicle: v.name, hours: numOf(hours), date: todayNZ(), time: nowTimeNZ(), ts: Date.now(), by: session.name };
      const nextH = [h, ...(hoursLog || [])];
      await saveJSON(K.hours, nextH); setHoursLog(nextH);
    }
    // anything failing becomes a maintenance job for the manager
    if (failed.length) {
      const list = await loadJSON(K.maint, []);
      const reports = failed.map(f => ({
        id: uid(), block: v.name, rowRef: '', kind: 'Machinery', urgent: false,
        detail: `${f} — found on ${entry.kind.toLowerCase()}`, heard: '',
        status: 'Open', reportedBy: session.name,
        reportedAt: todayNZ(), reportedTime: nowTimeNZ(), reportedTs: Date.now(),
      }));
      await saveJSON(K.maint, [...reports, ...list]);
    }
    setMsg(failed.length ? `Check saved — ${failed.length} item${failed.length > 1 ? 's' : ''} sent to maintenance.` : 'Check saved.');
    setActive(null); setResults({}); setHours(''); setNotes('');
    setTimeout(() => setMsg(''), 5000);
  };

  if (active) {
    const doneCount = items.filter(it => results[it]).length;
    return (
      <div className="space-y-4 pb-6">
        <div className="flex items-center gap-2">
          <button onClick={() => setActive(null)} className="p-1.5 -ml-1.5 rounded-lg hover:bg-stone-100 text-stone-500"><ChevronLeft size={20} /></button>
          <h2 className="text-lg font-semibold text-stone-900">{active.name}</h2>
          <span className="text-sm text-stone-400 ml-auto">{doneCount}/{items.length}</span>
        </div>
        <div className={cls.card + ' p-4 space-y-2'}>
          {items.map(it => (
            <div key={it} className="flex items-center gap-2 py-1.5 border-b border-stone-100 last:border-0">
              <span className="text-[15px] text-stone-800 flex-1 min-w-0">{it}</span>
              <button onClick={() => setResults({ ...results, [it]: 'ok' })}
                className={'px-3 py-1.5 rounded-lg border text-sm font-medium shrink-0 ' +
                  (results[it] === 'ok' ? 'bg-emerald-600 border-emerald-600 text-white' : 'bg-white border-stone-300 text-stone-600')}>OK</button>
              <button onClick={() => setResults({ ...results, [it]: 'bad' })}
                className={'px-3 py-1.5 rounded-lg border text-sm font-medium shrink-0 ' +
                  (results[it] === 'bad' ? 'bg-red-600 border-red-600 text-white' : 'bg-white border-stone-300 text-stone-600')}>Fix</button>
            </div>
          ))}
        </div>
        <div className={cls.card + ' p-4 space-y-3'}>
          <div><label className={cls.label}>Engine hours now</label>
            <input value={hours} onChange={e => setHours(e.target.value)} inputMode="decimal" className={cls.input + ' text-lg'} /></div>
          <div><label className={cls.label}>Notes</label>
            <textarea value={notes} onChange={e => setNotes(e.target.value)} rows={2} className={cls.input + ' resize-y'} /></div>
          <button onClick={submitCheck} className={cls.primary + ' w-full !py-3.5 text-base'}><Check size={18} /> Submit check</button>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-4 pb-6">
      <div className="flex items-center gap-2 text-stone-700"><Truck size={18} /><h2 className="text-lg font-semibold text-stone-900">My machines</h2></div>
      {msg && <div className="text-sm px-3 py-2.5 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200">{msg}</div>}
      {mine.length === 0 && !showAll && (
        <div className="rounded-xl border border-stone-200 bg-white p-5 text-center">
          <Truck size={26} className="text-stone-300 mx-auto mb-2" />
          <p className="text-stone-600 text-sm">No machines are assigned to you yet.</p>
          <p className="text-stone-400 text-[13px] mt-1">Ask for your machines to be assigned in Setup — or open any machine below to record hours or a check.</p>
          <button onClick={() => setShowAll(true)} className={cls.ghost + ' !py-2 !px-3 mt-3'}>Show all machines</button>
        </div>
      )}

      {mine.length > 0 && !showAll && others.length > 0 && (
        <button onClick={() => setShowAll(true)} className="text-[13px] text-stone-500 underline">Show the other machines too</button>
      )}
      {showAll && (
        <button onClick={() => setShowAll(false)} className="text-[13px] text-stone-500 underline">Show only my machines</button>
      )}

      {shown.map(v => {
        const hours = machineHours(v, hoursLog, workCards, config);
        const svc = serviceStatus(v, hours, rm);
        const chk = checkStatus(v, rm);
        const hr = hoursAsOf(v.name, hoursLog);
        const days = hr ? Math.floor((Date.now() - hr.ts) / 86400000) : null;
        const hoursDue = v.hoursSource !== 'tasks' && (days == null || days >= 7);
        return (
          <div key={v.name} className={'rounded-xl border p-4 ' + (chk.due ? 'bg-red-50 border-red-300' : svc.due ? 'bg-amber-50 border-amber-300' : 'bg-white border-stone-200')}>
            <div className="flex items-start justify-between gap-2">
              <div className="min-w-0">
                <div className="font-bold text-[17px] text-stone-900">{v.name}</div>
                <div className="text-[13px] text-stone-600 mt-0.5">{fmtNum(hours)} h on the clock</div>
                <div className={'text-[13px] mt-0.5 ' + (chk.due ? 'text-red-700 font-medium' : 'text-stone-500')}>
                  Checklist {chk.tracked ? chk.detail : '— no interval set'}
                </div>
                {svc.tracked && (
                  <div className={'text-[12px] mt-0.5 ' + (svc.due ? 'text-amber-800 font-medium' : 'text-stone-400')}>
                    Workshop service at {fmtNum(svc.nextAt)} h · {svc.detail}
                  </div>
                )}
              </div>
              {chk.due && <span className="text-[11px] font-semibold uppercase tracking-wide text-red-700 bg-red-100 border border-red-200 rounded-full px-2 py-0.5 shrink-0">Check due</span>}
            </div>

            {v.hoursSource !== 'tasks' && (
              <div className={'mt-3 rounded-lg border px-3 py-2.5 ' + (hoursDue ? 'bg-amber-50 border-amber-300' : 'bg-stone-50 border-stone-200')}>
                {hoursFor && hoursFor.name === v.name ? (
                  <div className="flex gap-2 items-end">
                    <div className="flex-1"><label className={cls.label}>Engine hours</label>
                      <input value={hoursVal} onChange={e => setHoursVal(e.target.value)} inputMode="decimal" autoFocus className={cls.input + ' text-lg'} /></div>
                    <button onClick={saveHours} className={cls.primary + ' !py-2.5'}><Check size={16} /></button>
                    <button onClick={() => { setHoursFor(null); setHoursVal(''); }} className={cls.ghost + ' !py-2.5 !px-3'}>Cancel</button>
                  </div>
                ) : (
                  <div className="flex items-center gap-2 flex-wrap">
                    <div className="text-[13px] min-w-0">
                      <div className={hoursDue ? 'text-amber-900 font-medium' : 'text-stone-600'}>
                        {hoursDue ? 'Weekly hours reading due' : 'Hours up to date'}
                      </div>
                      <div className="text-[12px] text-stone-400">{hr ? `${fmtNum(hr.hours)} h on ${hr.date}` : 'never entered'}</div>
                    </div>
                    <button onClick={() => { setHoursFor(v); setHoursVal(String(hoursOf(v.name, hoursLog) || '')); }}
                      className={cls.ghost + ' !py-2 !px-3 ml-auto'}>Update hours</button>
                  </div>
                )}
              </div>
            )}

            <button onClick={() => { setActive(v); setResults({}); setHours(String(hours || '')); setNotes(''); }}
              className={cls.primary + ' w-full justify-center !py-2.5 mt-3'}>
              <Check size={16} /> Do the checklist
            </button>
          </div>
        );
      })}
    </div>
  );
}

function FleetManager({ config, setConfig }) {
  const [fuel, setFuel] = useState(null);
  const [rm, setRm] = useState([]);
  const [hoursLog, setHoursLog] = useState([]);
  const [pane, setPane] = useState('status');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [openAssign, setOpenAssign] = useState(null);
  const [workCards, setWorkCards] = useState([]);   // implements take their hours from these
  const load = async () => {
    setFuel(await loadJSON(K.fuel, [])); setRm(await loadJSON(K.rm, []));
    setHoursLog(await loadJSON(K.hours, [])); setWorkCards(await loadJSON(K.work, []));
  };
  useEffect(() => { load(); }, []);
  useLiveKey(K.fuel, v => setFuel(v || []));
  useLiveKey(K.rm, v => setRm(v || []));
  useLiveKey(K.hours, v => setHoursLog(v || []));
  useLiveKey(K.work, v => setWorkCards(v || []));

  if (fuel === null) return <div className="p-8 text-center text-stone-400">Loading fleet…</div>;
  const vehicles = config.vehicles || [];
  const statuses = vehicles.map(v => {
    const hours = machineHours(v, hoursLog, workCards, config);
    return { v, hours, svc: serviceStatus(v, hours, rm), chk: checkStatus(v, rm) };
  });
  const dueNow = statuses.filter(x => x.svc.due || x.chk.due);

  // the workshop has serviced a machine — record it at the hours on the clock
  const recordService = async (v, currentHours) => {
    const at = window.prompt(`Service done on ${v.name}.\n\nEngine hours at the service?`, String(currentHours || ''));
    if (at === null) return;
    const note = window.prompt('Anything to note? (workshop, parts, leave blank if none)', '') || '';
    const entry = {
      id: uid(), vehicle: v.name, kind: 'Service', date: todayNZ(), time: nowTimeNZ(), ts: Date.now(),
      hours: numOf(at), by: 'Manager', note: note.trim(),
    };
    const next = [entry, ...rm];
    setRm(next); await saveJSON(K.rm, next);
  };

  const isoNZ = v => { const m = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : ''; };
  const inRange = iso => (!from && !to) ? true : (iso ? ((!from || iso >= from) && (!to || iso <= to)) : false);
  const fuelInRange = (fuel || []).filter(f => inRange(f.date));           // fuel dates are already ISO
  const rmInRange = (rm || []).filter(r => inRange(isoNZ(r.date)));

  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();
    addSheet(wb, XLSX.utils.json_to_sheet((fuelInRange || []).map(f => ({
      Date: f.date, Vehicle: f.vehicle, 'Diesel (L)': f.litres, 'Meter start': f.meterStart, 'Meter end': f.meterEnd,
      'Hours run': f.hoursRun, 'Logged by': f.by,
    })) || [{}]), 'Diesel');
    addSheet(wb, XLSX.utils.json_to_sheet((hoursLog || []).slice().sort((a, b) => (b.ts || 0) - (a.ts || 0)).map(h => ({
      Date: h.date, Time: h.time, Machine: h.vehicle, 'Engine hours': numOf(h.hours), 'Entered by': h.by,
    })) || [{}]), 'Engine hours');
    addSheet(wb, XLSX.utils.json_to_sheet((rmInRange || []).map(r => ({
      Date: r.date, Time: r.time, Vehicle: r.vehicle, What: r.kind, 'Hours at the time': r.hours, By: r.by, Notes: r.note || '',
    })) || [{}]), 'R&M');
    addSheet(wb, XLSX.utils.json_to_sheet(statuses.map(({ v, hours, svc, chk }) => ({
      Vehicle: v.name, Type: v.machineType || '', 'Scheduled by': v.kind === 'vehicle' ? 'Weekly check' : 'Hours run',
      'Actual hours now': hours,
      'Hours read on': (hoursAsOf(v.name, hoursLog) || {}).date || '', 'Read by': (hoursAsOf(v.name, hoursLog) || {}).by || '',
      'REGO due': (expiryStatus(v.rego) || {}).label || '', 'WOF due': (expiryStatus(v.wof) || {}).label || '',
      'Service every (h)': v.serviceEveryHours || '',
      'Next service at (h)': svc.tracked ? svc.nextAt : '', 'Hours to service': svc.tracked ? svc.remaining : '',
      'Service status': svc.due ? 'DUE' : svc.soon ? 'Due soon' : svc.tracked ? 'OK' : '',
      'Last serviced': svc.lastAt || '', 'Checklist every (days)': v.checkEveryDays || '',
      'Checklist status': chk.due ? 'DUE' : chk.tracked ? 'OK' : '', 'Last checklist': chk.lastAt || '',
    })) || [{}]), 'Status');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = `fleet_${todayStr()}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const setVeh = (i, patch) => setConfig({ ...config, vehicles: vehicles.map((v, j) => (j === i ? { ...v, ...patch } : v)) });
  const gi = 'px-2 py-1.5 rounded-md border border-stone-200 bg-white text-stone-800 text-[13px] focus:outline-none focus:ring-2 focus:ring-stone-400/40';

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700"><Truck size={18} />
          <h2 className="text-lg font-semibold text-stone-900">Fleet & R&amp;M</h2>
          {dueNow.length > 0 && <span className="text-sm text-red-700 bg-red-50 border border-red-200 rounded-full px-2.5 py-0.5">{dueNow.length} due</span>}
          {(() => {
            const ex = expiringItems(vehicles);
            if (!ex.length) return null;
            const bad = ex.filter(x => x.expired).length;
            return (
              <span className={'text-sm rounded-full px-2.5 py-0.5 border ' + (bad ? 'text-red-700 bg-red-50 border-red-200' : 'text-amber-800 bg-amber-50 border-amber-200')}>
                {bad ? `${bad} expired` : `${ex.length} expiring`}
              </span>
            );
          })()}
        </div>
        <div className="flex gap-2 items-center flex-wrap">
          <div className="inline-flex rounded-lg border border-stone-300 overflow-hidden text-sm">
            {[['status', 'Status'], ['hours', 'Hours'], ['fuel', 'Diesel log'], ['history', 'R&M history'], ['setup', 'Fleet setup']].map(([k, l]) => (
              <button key={k} onClick={() => setPane(k)}
                className={'px-3 py-1.5 font-medium border-l first:border-l-0 border-stone-300 ' + (pane === k ? 'bg-stone-900 text-stone-50' : 'bg-white text-stone-600 hover:bg-stone-50')}>{l}</button>
            ))}
          </div>
          <button onClick={load} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /></button>
          <button onClick={exportXlsx} className={cls.primary + ' !py-2 !px-3'}><Download size={15} /> Export</button>
        </div>
      </div>

      {(pane === 'fuel' || pane === 'history') && (
        <RangePicker from={from} to={to} onChange={(f, t) => { setFrom(f); setTo(t); }} compact />
      )}

      {pane === 'status' && (
        <div className="grid gap-3 md:grid-cols-2 lg:grid-cols-3">
          {statuses.map(({ v, hours, svc, chk }) => {
            const alert = svc.due || chk.due;
            const warn = !alert && (svc.soon || chk.soon);
            return (
              <div key={v.name} className={'rounded-xl border p-4 ' + (alert ? 'bg-red-50 border-red-300' : warn ? 'bg-amber-50 border-amber-300' : 'bg-white border-stone-200')}>
                <div className="flex items-start justify-between gap-2">
                  <div className="min-w-0">
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <span className="font-bold text-[16px] text-stone-900 truncate">{v.name}</span>
                      <span className="text-[10px] font-semibold uppercase tracking-wide text-stone-600 bg-stone-100 border border-stone-200 rounded px-1.5 py-0.5">{v.machineType || 'Other'}</span>
                    </div>
                    <div className="text-[12px] text-stone-500">
                      {fmtNum(hours)} h on the clock
                      {v.hoursSource === 'tasks' && <span className="text-stone-400"> · counted from work</span>}
                    </div>
                  </div>
                </div>

                {/* workshop service, on hours */}
                <div className="mt-3 rounded-lg border border-stone-200 bg-white/70 px-3 py-2">
                  <div className="flex items-baseline justify-between gap-2">
                    <span className="text-[11px] font-semibold uppercase tracking-wide text-stone-500">Workshop service</span>
                    {svc.due && <span className="text-[10px] font-semibold uppercase tracking-wide text-red-700 bg-red-100 border border-red-200 rounded px-1.5 py-0.5">Due</span>}
                    {svc.soon && !svc.due && <span className="text-[10px] font-semibold uppercase tracking-wide text-amber-700 bg-amber-100 border border-amber-200 rounded px-1.5 py-0.5">Soon</span>}
                  </div>
                  {svc.tracked ? (
                    <>
                      <div className="flex justify-between text-[13px] mt-1">
                        <span className="text-stone-700 font-medium">next at {fmtNum(svc.nextAt)} h</span>
                        <span className={svc.due ? 'text-red-700 font-medium' : 'text-stone-500'}>{svc.detail}</span>
                      </div>
                      <div className="h-2 rounded-full bg-stone-200 overflow-hidden mt-1.5">
                        <div className="h-full rounded-full" style={{
                          width: Math.min(100, Math.max(0, (svc.run / svc.every) * 100)) + '%',
                          backgroundColor: svc.due ? '#dc2626' : svc.soon ? '#d97706' : '#57534e',
                        }} />
                      </div>
                      <div className="text-[11px] text-stone-400 mt-1">
                        {svc.lastAt ? `Last serviced ${svc.lastAt} at ${fmtNum(svc.lastHours)} h` : `From ${fmtNum(svc.base)} h — no service recorded yet`}
                      </div>
                    </>
                  ) : <div className="text-[13px] text-stone-500 mt-1">{svc.detail}</div>}
                  <button onClick={() => recordService(v, hours)} className={cls.ghost + ' !py-2 !px-3 mt-2 w-full justify-center'}>
                    <Check size={15} /> Record a service
                  </button>
                </div>

                {/* operator checklist, on days */}
                <div className="mt-2 rounded-lg border border-stone-200 bg-white/70 px-3 py-2">
                  <div className="flex items-baseline justify-between gap-2">
                    <span className="text-[11px] font-semibold uppercase tracking-wide text-stone-500">Operator checklist</span>
                    {chk.due && <span className="text-[10px] font-semibold uppercase tracking-wide text-red-700 bg-red-100 border border-red-200 rounded px-1.5 py-0.5">Due</span>}
                  </div>
                  <div className="text-[13px] text-stone-600 mt-1">
                    {chk.tracked ? <>every {chk.every} days · {chk.detail}</> : chk.detail}
                  </div>
                  <div className="text-[11px] text-stone-400 mt-0.5">
                    {chk.lastAt ? `Last done ${chk.lastAt}${chk.lastBy ? ` by ${chk.lastBy}` : ''}` : 'Never done'}
                  </div>
                </div>

                {(() => {
                  const items = [['REGO', v.rego], ['WOF', v.wof]].map(([w, iso]) => [w, expiryStatus(iso)]).filter(([, x]) => x);
                  if (!items.length) return null;
                  return (
                    <div className="flex gap-1.5 flex-wrap mt-2">
                      {items.map(([what, x]) => (
                        <span key={what} className={'text-[11px] font-semibold rounded px-1.5 py-0.5 border ' +
                          (x.expired ? 'bg-red-100 border-red-300 text-red-800'
                            : x.soon ? 'bg-amber-100 border-amber-300 text-amber-900'
                            : 'bg-stone-100 border-stone-200 text-stone-600')}>
                          {what} {x.expired ? 'expired' : x.label}
                        </span>
                      ))}
                    </div>
                  );
                })()}
              </div>
            );
          })}
        </div>
      )}

      {pane === 'hours' && (
        <div className="overflow-x-auto border border-stone-200 rounded-xl bg-white">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
              <th className="px-3 py-2.5 font-semibold">Date</th><th className="px-3 py-2.5 font-semibold">Time</th>
              <th className="px-3 py-2.5 font-semibold">Machine</th><th className="px-3 py-2.5 font-semibold text-right">Engine hours</th>
              <th className="px-3 py-2.5 font-semibold">Entered by</th>
            </tr></thead>
            <tbody>
              {(hoursLog || []).length === 0
                ? <tr><td colSpan={5} className="px-3 py-6 text-center text-stone-400">No hours entered yet — operators add these from My machines.</td></tr>
                : hoursLog.slice().sort((a, b) => (b.ts || 0) - (a.ts || 0)).map(h => (
                  <tr key={h.id} className="border-b border-stone-100 last:border-0">
                    <td className="px-3 py-2.5 whitespace-nowrap">{h.date}</td>
                    <td className="px-3 py-2.5 text-stone-500">{h.time}</td>
                    <td className="px-3 py-2.5 font-medium text-stone-900">{h.vehicle}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums font-semibold">{fmtNum(h.hours)} h</td>
                    <td className="px-3 py-2.5 text-stone-500">{h.by}</td>
                  </tr>
                ))}
            </tbody>
          </table>
        </div>
      )}

      {pane === 'fuel' && (
        <div className="overflow-x-auto border border-stone-200 rounded-xl bg-white">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
              <th className="px-3 py-2.5 font-semibold">Date</th><th className="px-3 py-2.5 font-semibold">Vehicle</th>
              <th className="px-3 py-2.5 font-semibold text-right">Diesel</th><th className="px-3 py-2.5 font-semibold text-right">Meter start</th>
              <th className="px-3 py-2.5 font-semibold text-right">Meter end</th><th className="px-3 py-2.5 font-semibold text-right">Hours run</th>
              <th className="px-3 py-2.5 font-semibold">By</th>
            </tr></thead>
            <tbody>
              {fuelInRange.length === 0 ? <tr><td colSpan={7} className="px-3 py-6 text-center text-stone-400">No fills in this range.</td></tr>
                : fuelInRange.map(f => (
                  <tr key={f.id} className="border-b border-stone-100 last:border-0">
                    <td className="px-3 py-2.5 whitespace-nowrap">{f.date}</td>
                    <td className="px-3 py-2.5 font-medium text-stone-900">{f.vehicle}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums">{fmtNum(f.litres)} L</td>
                    <td className="px-3 py-2.5 text-right tabular-nums text-stone-500">{f.meterStart !== '' ? fmtNum(f.meterStart) : '—'}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums text-stone-500">{f.meterEnd !== '' ? fmtNum(f.meterEnd) : '—'}</td>
                    <td className="px-3 py-2.5 text-right tabular-nums">{f.hoursRun !== '' ? fmtNum(f.hoursRun) : '—'}</td>
                    <td className="px-3 py-2.5 text-stone-500">{f.by}</td>
                  </tr>
                ))}
            </tbody>
          </table>
        </div>
      )}

      {pane === 'history' && (
        <div className="space-y-2">
          {rmInRange.length === 0 ? <p className="text-stone-400 text-sm text-center py-10">No checks or services in this range.</p>
            : rmInRange.map(r => (
              <div key={r.id} className="flex items-center gap-3 bg-white border border-stone-200 rounded-lg px-4 py-3">
                <div className="min-w-0">
                  <div className="font-semibold text-stone-900">{r.vehicle} <span className="font-normal text-stone-500">· {r.kind}</span></div>
                  <div className="text-[12px] text-stone-400">{r.date} {r.time} · {r.by}{r.hours ? ` · ${fmtNum(r.hours)} h` : ''}{r.note ? ` · ${r.note}` : ''}</div>
                </div>
              </div>
            ))}
        </div>
      )}

      {pane === 'setup' && (
        <div className="overflow-x-auto border border-stone-200 rounded-xl bg-white">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200 bg-stone-50">
              <th className="px-3 py-2.5 font-semibold">Vehicle / machine</th>
              <th className="px-3 py-2.5 font-semibold">Type</th>
              <th className="px-3 py-2.5 font-semibold">Schedule by</th>
              <th className="px-3 py-2.5 font-semibold">Service every (h)</th>
              <th className="px-3 py-2.5 font-semibold">Hours at last service</th>
              <th className="px-3 py-2.5 font-semibold">Assigned to</th>
              <th className="px-3 py-2.5 font-semibold">Checklist every</th>
              <th className="px-3 py-2.5 font-semibold">Hours from</th>
              <th className="px-3 py-2.5 font-semibold">REGO due</th>
              <th className="px-3 py-2.5 font-semibold">WOF due</th>
              <th className="px-3 py-2.5 font-semibold">Hours now</th>
              <th className="px-3 py-2.5"></th>
            </tr></thead>
            <tbody>
              {vehicles.map((v, i) => (
                <React.Fragment key={i}>
                <tr className="border-b border-stone-100 last:border-0">
                  <td className="px-2 py-1.5"><input value={v.name} onChange={e => setVeh(i, { name: e.target.value })} className={gi + ' w-full min-w-[170px] font-medium'} /></td>
                  <td className="px-2 py-1.5">
                    <select value={v.machineType || 'Other'} onChange={e => setVeh(i, { machineType: e.target.value })} className={gi + ' w-32'}>
                      {MACHINE_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                    </select>
                  </td>
                  <td className="px-2 py-1.5">
                    <select value={v.kind} onChange={e => setVeh(i, { kind: e.target.value })} className={gi}>
                      <option value="machine">Hours run</option>
                      <option value="vehicle">Weekly check</option>
                    </select>
                  </td>
                  <td className="px-2 py-1.5"><input value={v.serviceEveryHours ?? ''} onChange={e => setVeh(i, { serviceEveryHours: e.target.value })} inputMode="decimal" disabled={v.kind === 'vehicle'} className={gi + ' w-24 text-right disabled:opacity-40'} /></td>
                  <td className="px-2 py-1.5"><input value={v.lastServiceHours ?? ''} onChange={e => setVeh(i, { lastServiceHours: e.target.value })} inputMode="decimal" className={gi + ' w-28 text-right'} /></td>
                  <td className="px-2 py-1.5">
                    <button onClick={() => setOpenAssign(openAssign === i ? null : i)}
                      className={'px-2.5 py-1.5 rounded-md border text-[13px] whitespace-nowrap ' +
                        (openAssign === i ? 'bg-stone-900 border-stone-900 text-stone-50' : 'bg-white border-stone-300 text-stone-700 hover:bg-stone-50')}>
                      {(v.assignedTo || []).length
                        ? (config.operators || []).filter(o => (v.assignedTo || []).includes(o.code)).map(o => o.name).join(', ') || `${v.assignedTo.length} people`
                        : 'Nobody'}
                    </button>
                  </td>
                  <td className="px-2 py-1.5">
                    <div className="flex items-center gap-1.5">
                      <input value={v.checkEveryDays ?? ''} onChange={e => setVeh(i, { checkEveryDays: e.target.value })} inputMode="numeric" className={gi + ' w-16 text-right'} />
                      <span className="text-[12px] text-stone-400">days</span>
                    </div>
                  </td>
                  <td className="px-2 py-1.5">
                    <select value={v.hoursSource || 'manual'} onChange={e => setVeh(i, { hoursSource: e.target.value })} className={gi + ' w-36'}>
                      <option value="manual">Operator enters</option>
                      <option value="tasks">Counted from work</option>
                    </select>
                  </td>
                  {[['rego', v.rego], ['wof', v.wof]].map(([field, val]) => {
                    const st = expiryStatus(val);
                    return (
                      <td key={field} className="px-2 py-1.5">
                        <input type="date" value={val || ''} onChange={e => setVeh(i, { [field]: e.target.value })}
                          className={gi + ' w-36 ' + (st && st.expired ? 'border-red-400 bg-red-50 text-red-800' : st && st.soon ? 'border-amber-400 bg-amber-50 text-amber-900' : '')} />
                        {st && (st.expired || st.soon) && (
                          <div className={'text-[10px] mt-0.5 font-semibold ' + (st.expired ? 'text-red-700' : 'text-amber-700')}>
                            {st.expired ? `expired ${Math.abs(st.days)} d ago` : `${st.days} d to go`}
                          </div>
                        )}
                      </td>
                    );
                  })}
                  <td className="px-2 py-1.5">
                    {(() => {
                      const hr = hoursAsOf(v.name, hoursLog);
                      return (
                        <div className="text-right">
                          <div className="text-[13px] font-semibold text-stone-900 tabular-nums">{fmtNum(hoursOf(v.name, hoursLog))} h</div>
                          <div className="text-[10px] text-stone-400 whitespace-nowrap">{hr ? `${hr.date} · ${hr.by}` : 'none entered'}</div>
                        </div>
                      );
                    })()}
                  </td>
                  <td className="px-2 py-1.5 text-right"><button onClick={() => setConfig({ ...config, vehicles: vehicles.filter((_, j) => j !== i) })} className="p-1.5 rounded-md hover:bg-red-50 text-red-500"><Trash2 size={15} /></button></td>
                </tr>
                {openAssign === i && (
                  <tr className="border-b border-stone-100 bg-stone-50">
                    <td colSpan={6} className="px-3 py-2.5">
                      <div className="text-[13px] text-stone-600 mb-2">Who checks and services <b className="text-stone-900">{v.name || 'this machine'}</b>?</div>
                      <div className="flex flex-wrap gap-1.5">
                        {(config.operators || []).map(o => {
                          const on = (v.assignedTo || []).includes(o.code);
                          return (
                            <button key={o.code} onClick={() => {
                              const cur = v.assignedTo || [];
                              setVeh(i, { assignedTo: on ? cur.filter(c => c !== o.code) : [...cur, o.code] });
                            }}
                              className={'px-3 py-1.5 rounded-lg border text-[13px] ' +
                                (on ? 'bg-stone-900 border-stone-900 text-stone-50 font-medium' : 'bg-white border-stone-300 text-stone-700 hover:border-stone-400')}>
                              {o.name}
                            </button>
                          );
                        })}
                        {(config.operators || []).length === 0 && <span className="text-sm text-stone-400">Add operators in Setup first.</span>}
                      </div>
                    </td>
                  </tr>
                )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
          <div className="p-2.5">
            <button onClick={() => setConfig({ ...config, vehicles: [...vehicles, { name: '', kind: 'machine', serviceEveryHours: 250, lastServiceHours: 0 }] })} className={cls.ghost + ' !py-2 !px-3'}><Plus size={16} /> Add vehicle or machine</button>
          </div>
          <p className="text-xs text-stone-400 px-3 pb-3">Hours come from the meter readings operators enter when they fill up, so the clock keeps itself up to date.</p>
        </div>
      )}
    </div>
  );
}

/* ============================================================
   Maintenance reports — spoken in from the vineyard
   ============================================================ */
const MAINT_KINDS = [
  { label: 'Irrigation', re: /irrigation|leak|dripper|sprinkler|water|pipe|valve|solenoid|filter/i },
  { label: 'Wires', re: /\bwire|fruiting wire|foliage wire|cordon/i },
  { label: 'Posts', re: /\bpost|strainer|stay\b|anchor/i },
  { label: 'Trellis', re: /trellis|clip|staple/i },
  { label: 'Netting', re: /net\b|netting/i },
  { label: 'Machinery', re: /tractor|sprayer|mower|machine|implement/i },
  { label: 'Track / access', re: /track|road|gate|culvert/i },
  { label: 'Other', re: /.*/ },
];
const normWords = t => String(t).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const WORD_NUMS = { one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10, eleven: 11, twelve: 12 };
const spokenDigits = t => String(t).replace(/\b(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve)\b/gi, m => WORD_NUMS[m.toLowerCase()]);

// best-guess block from what was said
function matchBlockName(text, blocks) {
  const t = normWords(spokenDigits(text));
  let best = null, score = 0;
  const names = (blocks || []).map(b => b.name).filter(n => n && n !== 'N/A');
  names.forEach(n => {
    const bn = normWords(n); if (!bn) return;
    const toks = bn.split(' ').filter(Boolean);
    let hit = 0;
    toks.forEach(tok => { if (new RegExp('\\b' + tok.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b').test(t)) hit++; });
    if (hit === toks.length) { const sc = hit / toks.length + (t.includes(bn) ? 1 : 0); if (sc > score) { score = sc; best = n; } }
  });
  if (!best) names.forEach(n => {
    const bn = normWords(n); if (!bn) return;
    const toks = bn.split(' ').filter(Boolean);
    let hit = 0; toks.forEach(tok => { if (t.includes(tok)) hit++; });
    const sc = hit / toks.length;
    if (sc >= 0.6 && sc > score) { score = sc; best = n; }
  });
  return best || '';
}
function parseRowRef(text) {
  const t = spokenDigits(String(text));
  let m = t.match(/\brows?\s*(\d+)\s*(?:to|through|-|and)\s*(\d+)/i); if (m) return `${m[1]}-${m[2]}`;
  m = t.match(/\brows?\s*(\d+)/i); if (m) return m[1];
  m = t.match(/\bposts?\s*(\d+)/i); if (m) return `post ${m[1]}`;
  m = t.match(/\bbays?\s*(\d+)/i); if (m) return `bay ${m[1]}`;
  return '';
}
const maintKind = t => (MAINT_KINDS.find(k => k.re.test(t)) || { label: 'Other' }).label;
const maintUrgent = t => /urgent|asap|emergency|flooding|burst|major|straight away|right now/i.test(t);

function MaintenanceForm({ config, session }) {
  const SR = typeof window !== 'undefined' && (window.SpeechRecognition || window.webkitSpeechRecognition);
  const [listening, setListening] = useState(false);
  const [heard, setHeard] = useState('');
  const [block, setBlock] = useState('');
  const [rowRef, setRowRef] = useState('');
  const [kind, setKind] = useState('Other');
  const [detail, setDetail] = useState('');
  const [urgent, setUrgent] = useState(false);
  const [msg, setMsg] = useState('');
  const [err, setErr] = useState('');
  const [existing, setExisting] = useState([]);
  useEffect(() => { (async () => setExisting(await loadJSON(K.maint, [])))(); }, []);
  useLiveKey(K.maint, v => setExisting(v || []));
  const recRef = useRef(null);

  const applyTranscript = text => {
    setHeard(text);
    const b = matchBlockName(text, config.blocks); if (b) setBlock(b);
    const r = parseRowRef(text); if (r) setRowRef(r);
    setKind(maintKind(text));
    if (maintUrgent(text)) setUrgent(true);
    setDetail(text);
  };

  const start = () => {
    setErr('');
    if (!SR) { setErr('This browser can’t do voice — type it in below instead. (Voice works in Chrome on Android and Safari on iPhone.)'); return; }
    try {
      const rec = new SR();
      recRef.current = rec;
      rec.lang = 'en-NZ'; rec.interimResults = true; rec.continuous = false; rec.maxAlternatives = 1;
      let final = '';
      rec.onresult = e => {
        let interim = '';
        for (let i = e.resultIndex; i < e.results.length; i++) {
          const chunk = e.results[i][0].transcript;
          if (e.results[i].isFinal) final += chunk; else interim += chunk;
        }
        setHeard((final + ' ' + interim).trim());
      };
      rec.onerror = ev => {
        setListening(false);
        setErr(ev.error === 'not-allowed'
          ? 'Microphone blocked — allow mic access for this site, then try again.'
          : 'Didn’t catch that. Try again, or type it below.');
      };
      rec.onend = () => { setListening(false); if (final.trim()) applyTranscript(final.trim()); };
      rec.start();
      setListening(true);
    } catch { setErr('Couldn’t start the microphone.'); }
  };
  const stop = () => { try { recRef.current && recRef.current.stop(); } catch { /* ignore */ } setListening(false); };

  const valid = block && (detail.trim() || kind !== 'Other');
  // already-open report for the same block, row and type?
  const sameRef = (a, b) => String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
  const duplicate = (existing || []).find(r => r.status !== 'Done'
    && sameRef(r.block, block) && sameRef(r.rowRef, rowRef) && sameRef(r.kind, kind));
  const save = async () => {
    if (!valid) return;
    const entry = {
      id: uid(), block, rowRef, kind, urgent,
      detail: detail.trim(), heard: heard.trim(),
      status: 'Open',
      reportedBy: session ? session.name : 'Manager',
      reportedAt: todayNZ(), reportedTime: nowTimeNZ(), reportedTs: Date.now(),
    };
    const list = await loadJSON(K.maint, []);
    const next = [entry, ...list];
    await saveJSON(K.maint, next); setExisting(next);
    if (config.webhookUrl && urgent) {
      postWebhook(config.webhookUrl, {
        type: 'maintenance', notifyEmail: config.notifyEmail || '', siteName: config.siteName,
        subject: `[${config.siteName}] Urgent maintenance — ${block}`,
        body: `${kind} — ${block}${rowRef ? ` (${rowRef})` : ''}\n\n${detail}\n\nReported by ${entry.reportedBy} at ${entry.reportedAt} ${entry.reportedTime}`,
      });
    }
    setMsg(`Reported: ${kind} at ${block}${rowRef ? ` · ${rowRef}` : ''}.`);
    setBlock(''); setRowRef(''); setKind('Other'); setDetail(''); setHeard(''); setUrgent(false);
    setTimeout(() => setMsg(''), 5000);
  };

  return (
    <div className="space-y-4 pb-6">
      <div className="flex items-center gap-2 text-stone-700"><Wrench size={18} /><h2 className="text-lg font-semibold text-stone-900">Report maintenance</h2></div>
      {msg && <div className="text-sm px-3 py-2.5 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200">{msg}</div>}

      <div className={cls.card + ' p-5 text-center'}>
        <p className="text-sm text-stone-500 mb-4">
          Tap and say the block, the row and what needs doing — for example<br />
          <span className="text-stone-700 italic">“Hill A 23, row 42, irrigation dripper leaking”</span>
        </p>
        <button onClick={listening ? stop : start}
          className={'w-28 h-28 rounded-full mx-auto flex items-center justify-center transition-all ' +
            (listening ? 'bg-red-600 text-white scale-105 animate-pulse shadow-lg' : 'bg-stone-900 text-stone-50 hover:bg-stone-800 shadow')}>
          <Mic size={44} />
        </button>
        <div className="mt-3 text-sm font-medium text-stone-700">{listening ? 'Listening — tap to stop' : 'Tap to speak'}</div>
        {heard && <div className="mt-3 text-[15px] text-stone-800 bg-stone-50 border border-stone-200 rounded-lg px-3 py-2.5 text-left">“{heard}”</div>}
        {err && <div className="mt-3 text-sm text-amber-800 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2 text-left">{err}</div>}
      </div>

      <div className={cls.card + ' p-4 space-y-3.5'}>
        <p className="text-sm text-stone-500 -mb-1">Check what it heard, fix anything it got wrong, then send.</p>
        <Combobox label="Block" options={(config.blocks || []).map(b => b.name)} value={block} onChange={setBlock} icon={MapPin} placeholder="Search blocks…" />
        <div className="flex gap-3">
          <div className="flex-1"><label className={cls.label}>Row / post</label>
            <input value={rowRef} onChange={e => setRowRef(e.target.value)} placeholder="e.g. 42 or 12-15" className={cls.input} /></div>
          <div className="flex-1"><label className={cls.label}>Type</label>
            <select value={kind} onChange={e => setKind(e.target.value)} className={cls.input}>
              {MAINT_KINDS.map(k => <option key={k.label}>{k.label}</option>)}
            </select></div>
        </div>
        <div><label className={cls.label}>What needs doing</label>
          <textarea value={detail} onChange={e => setDetail(e.target.value)} rows={3} className={cls.input + ' resize-y'} placeholder="Describe the problem" /></div>
        <label className="flex items-center gap-2.5 text-sm text-stone-700">
          <input type="checkbox" checked={urgent} onChange={e => setUrgent(e.target.checked)} className="w-4 h-4 accent-red-600" />
          Urgent — needs attention today
        </label>
        {duplicate && (
          <div className="flex items-start gap-2 rounded-lg border border-amber-300 bg-amber-50 px-3 py-2.5">
            <AlertTriangle size={17} className="text-amber-600 shrink-0 mt-0.5" />
            <div className="text-[13.5px] text-amber-900 leading-snug">
              <b>Already reported.</b> {duplicate.kind} at {duplicate.block}{duplicate.rowRef ? ` · ${duplicate.rowRef}` : ''} was logged
              by {duplicate.reportedBy} on {duplicate.reportedAt} and is still open — no need to send it again.
              <div className="text-[12px] text-amber-800/80 mt-1">“{duplicate.detail}”</div>
            </div>
          </div>
        )}
        <button onClick={save} disabled={!valid || !!duplicate} className={cls.primary + ' w-full !py-3.5 text-base'}>
          <Check size={18} /> Send report
        </button>
      </div>
    </div>
  );
}

function MaintenanceManager({ config }) {
  const [list, setList] = useState(null);
  const [filter, setFilter] = useState('open');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const load = async () => setList(await loadJSON(K.maint, []));
  useEffect(() => { load(); }, []);
  useLiveKey(K.maint, v => setList(v || []));
  const persist = async next => { setList(next); await saveJSON(K.maint, next); };

  if (list === null) return <div className="p-8 text-center text-stone-400">Loading reports…</div>;
  const isoOfNZ = v => { const m = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : ''; };
  const shown = list
    .filter(r => filter === 'all' || (filter === 'open' ? r.status !== 'Done' : r.status === 'Done'))
    .filter(r => { if (!from && !to) return true; const d = isoOfNZ(r.reportedAt); if (!d) return false; return (!from || d >= from) && (!to || d <= to); });
  const openCount = list.filter(r => r.status !== 'Done').length;

  const setStatus = (id, status) => persist(list.map(r => r.id === id
    ? { ...r, status, closedAt: status === 'Done' ? todayNZ() : '', closedTime: status === 'Done' ? nowTimeNZ() : '' } : r));
  const remove = id => { if (window.confirm('Delete this report?')) persist(list.filter(r => r.id !== id)); };
  // correct a report's date if it went in on the wrong day
  const amendDate = (id, iso) => {
    const [y, m, d] = String(iso).split('-');
    if (!y) return;
    persist(list.map(r => (r.id === id
      ? { ...r, dateISO: iso, reportedAt: `${d}/${m}/${y}`, reportedTs: new Date(+y, +m - 1, +d, 12).getTime() }
      : r)));
  };

  const exportXlsx = () => {
    const rows = shown.map(r => ({
      Status: r.status, Urgent: r.urgent ? 'Yes' : '', Type: r.kind, Block: r.block, 'Row / post': r.rowRef || '',
      'What needs doing': r.detail || '', 'Reported by': r.reportedBy || '',
      'Date reported': r.reportedAt || '', 'Time reported': r.reportedTime || '',
      'Date closed': r.closedAt || '', 'Time closed': r.closedTime || '', 'Spoken words': r.heard || '',
    }));
    const wb = XLSX.utils.book_new();
    addSheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{}]), 'Maintenance');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = `maintenance_${todayStr()}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700"><Wrench size={18} />
          <h2 className="text-lg font-semibold text-stone-900">Maintenance</h2>
          {openCount > 0 && <span className="text-sm text-red-700 bg-red-50 border border-red-200 rounded-full px-2.5 py-0.5">{openCount} open</span>}
        </div>
        <div className="flex gap-2 items-center flex-wrap">
          <div className="inline-flex rounded-lg border border-stone-300 overflow-hidden text-sm">
            {[['open', 'Open'], ['done', 'Done'], ['all', 'All']].map(([k, l]) => (
              <button key={k} onClick={() => setFilter(k)}
                className={'px-3 py-1.5 font-medium border-l first:border-l-0 border-stone-300 ' + (filter === k ? 'bg-stone-900 text-stone-50' : 'bg-white text-stone-600 hover:bg-stone-50')}>{l}</button>
            ))}
          </div>
          <button onClick={load} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /></button>
          <button onClick={exportXlsx} disabled={!shown.length} className={cls.primary + ' !py-2 !px-3'}><Download size={15} /> Export</button>
        </div>
      </div>

      <RangePicker from={from} to={to} onChange={(f, t) => { setFrom(f); setTo(t); }} compact />

      {shown.length === 0 ? (
        <p className="text-stone-400 text-sm text-center py-10">Nothing here.</p>
      ) : (() => {
        // one column per kind of job, so all the wire work sits together
        const groups = {};
        shown.forEach(r => { (groups[r.kind || 'Other'] = groups[r.kind || 'Other'] || []).push(r); });
        const order = [...MAINT_KINDS.map(k => k.label), ...Object.keys(groups).filter(k => !MAINT_KINDS.some(m => m.label === k))]
          .filter((k, i, a) => a.indexOf(k) === i && (groups[k] || []).length);
        return (
          <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 items-start">
            {order.map(kind => {
              const list = groups[kind].slice().sort((a, b) =>
                (b.urgent ? 1 : 0) - (a.urgent ? 1 : 0) || (b.reportedTs || 0) - (a.reportedTs || 0));
              const openN = list.filter(r => r.status !== 'Done').length;
              return (
                <div key={kind} className="rounded-xl border border-stone-200 bg-stone-50/60 p-2.5">
                  <div className="flex items-baseline justify-between gap-2 px-0.5 mb-2">
                    <span className="text-[14px] font-bold text-stone-900">{kind}</span>
                    <span className="text-[11px] text-stone-500">{openN ? `${openN} open` : 'all done'}</span>
                  </div>
                  <div className="space-y-2">
                    {list.map(r => (
                      <div key={r.id} className={'rounded-lg border p-2.5 ' +
                        (r.status === 'Done' ? 'bg-stone-100/70 border-stone-200' : r.urgent ? 'bg-red-50 border-red-300' : 'bg-white border-stone-200')}>
                        <div className="flex items-start justify-between gap-1.5">
                          <div className="min-w-0">
                            <div className={'text-[15px] font-bold truncate ' + (r.status === 'Done' ? 'text-stone-500' : 'text-stone-900')}>{r.block}</div>
                            {r.rowRef && <div className="text-[14px] font-semibold text-stone-700">{r.rowRef}</div>}
                          </div>
                          {r.urgent && r.status !== 'Done' && (
                            <span className="text-[10px] font-semibold uppercase tracking-wide text-red-700 bg-red-100 border border-red-200 rounded px-1.5 py-0.5 shrink-0">Urgent</span>
                          )}
                        </div>
                        <div className="text-[13px] text-stone-600 mt-1 leading-snug">{r.detail}</div>
                        <div className="text-[11px] text-stone-400 mt-1 flex items-center gap-1.5 flex-wrap">
                          <input type="date" value={r.dateISO || isoOfNZ(r.reportedAt) || ''} onChange={e => amendDate(r.id, e.target.value)} title="Change the date"
                            className="px-1 py-0.5 rounded border border-stone-200 text-[11px] text-stone-600 focus:outline-none focus:ring-2 focus:ring-stone-400/40" />
                          <span>{r.reportedBy}</span>
                          {r.pin && <a href={`https://www.google.com/maps?q=${r.pin.lat},${r.pin.lon}`} target="_blank" rel="noreferrer" className="underline inline-flex items-center gap-0.5"><MapPin size={11} /> map</a>}
                        </div>
                        {r.status === 'Done' && r.closedAt && <div className="text-[11px] text-emerald-700 mt-0.5">done {r.closedAt}</div>}
                        <div className="flex gap-1.5 mt-2">
                          <button onClick={() => setStatus(r.id, r.status === 'Done' ? 'Open' : 'Done')}
                            className={'flex-1 py-1.5 rounded-md border text-[12.5px] font-medium ' +
                              (r.status === 'Done' ? 'bg-white border-stone-300 text-stone-600' : 'bg-stone-900 border-stone-900 text-stone-50')}>
                            {r.status === 'Done' ? 'Reopen' : 'Mark done'}
                          </button>
                          <button onClick={() => remove(r.id)} className="p-1.5 rounded-md hover:bg-red-50 text-red-400 shrink-0"><Trash2 size={14} /></button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        );
      })()}
    </div>
  );
}

/* ============================================================
   Vineyard work — schedule any task onto blocks, assign operators,
   and track it on a kanban board (manager + operator views)
   ============================================================ */
/* Reusable kanban drag — same feel as the spray board: mouse drags immediately,
   touch needs a ~180ms press-and-hold (so a quick swipe still scrolls the list).
   Moves cards between lanes and reorders within a lane. */
// Scroll a lane into the middle of its board. Measured from live rectangles,
// not offsetLeft, so it doesn't depend on which ancestor happens to be positioned.
function centreLane(board, el) {
  if (!board || !el) return false;
  const b = board.getBoundingClientRect();
  const e = el.getBoundingClientRect();
  if (!b.width || !e.width) return false;
  board.scrollLeft += (e.left - b.left) - (b.width - e.width) / 2;
  return true;
}

function useKanbanDrag({ items, lanes, laneKey, persist }) {
  const [drag, setDrag] = useState(null);
  const [overCol, setOverCol] = useState(null);
  const [overIndex, setOverIndex] = useState(null);
  const boardRef = useRef(null);
  const colRefs = useRef({});
  const listRefs = useRef({});
  const itemsRef = useRef([]);
  const dragRef = useRef(null);
  const overColRef = useRef(null);
  const overIndexRef = useRef(null);
  const moveRef = useRef(null);
  const pendingRef = useRef(null);
  const dragActiveRef = useRef(false);
  const tmGuardRef = useRef(null);
  const lpTimerRef = useRef(null);
  const cancelRef = useRef(null);

  useEffect(() => { itemsRef.current = items || []; }, [items]);
  useEffect(() => () => { if (cancelRef.current) cancelRef.current(); }, []);

  const cleanupDrag = () => {
    if (lpTimerRef.current) clearTimeout(lpTimerRef.current);
    if (tmGuardRef.current) document.removeEventListener('touchmove', tmGuardRef.current);
    lpTimerRef.current = null; tmGuardRef.current = null;
    pendingRef.current = null; dragRef.current = null;
    overColRef.current = null; overIndexRef.current = null;
    dragActiveRef.current = false;
    cancelRef.current = null;
    document.body.style.userSelect = '';
    document.body.style.touchAction = '';
    document.body.style.overflow = '';
    setDrag(null); setOverCol(null); setOverIndex(null);
  };

  const updateTarget = (x, y) => {
    let lane = null;
    for (const s of Object.keys(colRefs.current)) {
      const el = colRefs.current[s]; if (!el) continue;
      const r = el.getBoundingClientRect();
      if (x >= r.left && x <= r.right) { lane = s; if (y >= r.top && y <= r.bottom) break; }
    }
    let index = 0;
    if (lane && listRefs.current[lane]) {
      const cards = Array.from(listRefs.current[lane].querySelectorAll('[data-card-id]'));
      index = cards.length;
      for (let i = 0; i < cards.length; i++) {
        const r = cards[i].getBoundingClientRect();
        if (y < r.top + r.height / 2) { index = i; break; }
      }
    }
    overColRef.current = lane; overIndexRef.current = index;
    setOverCol(lane); setOverIndex(index);
  };

  const drop = () => {
    const lane = overColRef.current; const d = dragRef.current;
    if (!lane || !d) return;
    const id = d.card.id;
    const arr = itemsRef.current.slice();
    const moving = arr.find(c => c.id === id); if (!moving) return;
    const laneCards = arr.filter(c => c[laneKey] === lane && c.id !== id)
      .sort((a, b) => (a.done ? 1 : 0) - (b.done ? 1 : 0));
    const idx = Math.min(Math.max(overIndexRef.current ?? laneCards.length, 0), laneCards.length);
    laneCards.splice(idx, 0, { ...moving, [laneKey]: lane });
    const byLane = {}; lanes.forEach(s => (byLane[s] = []));
    arr.filter(c => c[laneKey] !== lane && c.id !== id).forEach(c => { (byLane[c[laneKey]] = byLane[c[laneKey]] || []).push(c); });
    byLane[lane] = laneCards;
    const order = [...lanes, ...Object.keys(byLane).filter(s => !lanes.includes(s))];
    const next = []; order.forEach(s => (byLane[s] || []).forEach(c => next.push(c)));
    itemsRef.current = next; persist(next, { id, lane });
  };

  const autoScroll = (x, y) => {
    const board = boardRef.current;
    if (board) {
      const br = board.getBoundingClientRect();
      if (x > br.right - 48) board.scrollLeft += 16;
      else if (x < br.left + 48) board.scrollLeft -= 16;
    }
    if (y < 72) window.scrollBy(0, -12);
    else if (y > window.innerHeight - 72) window.scrollBy(0, 12);
  };

  const startPointer = (e, card) => {
    if (e.pointerType === 'mouse' && e.button !== 0) return;
    const isTouch = e.pointerType !== 'mouse';
    const pid = e.pointerId;
    pendingRef.current = { card, startX: e.clientX, startY: e.clientY, active: false, held: false, lastX: e.clientX, lastY: e.clientY };

    // Listeners go on window, NOT the card: once a drag starts the card is
    // replaced by the ghost, so any listener bound to it would die with it.
    const touchGuard = ev => { if (dragActiveRef.current && ev.cancelable) ev.preventDefault(); };
    document.addEventListener('touchmove', touchGuard, { passive: false });
    tmGuardRef.current = touchGuard;

    const activate = (x, y) => {
      const p = pendingRef.current; if (!p || p.active) return;
      p.active = true;
      dragRef.current = { card, fromLane: card[laneKey] };
      dragActiveRef.current = true;
      document.body.style.userSelect = 'none';
      // the card owns the gesture from here: stop the browser panning as well
      document.body.style.touchAction = 'none';
      document.body.style.overflow = 'hidden';
      if (isTouch && navigator.vibrate) { try { navigator.vibrate(18); } catch { /* ignore */ } }
      setDrag({ card, x, y });
      updateTarget(x, y);
    };

    const onMove = ev => {
      if (ev.pointerId != null && pid != null && ev.pointerId !== pid) return;
      const p = pendingRef.current; if (!p) return;
      const x = ev.clientX, y = ev.clientY;
      p.lastX = x; p.lastY = y;
      if (!p.active) {
        const dx = x - p.startX, dy = y - p.startY;
        if (isTouch) {
          // before the hold completes, real movement means they want to scroll
          if (Math.abs(dx) > 14 || Math.abs(dy) > 14) finish(false);
          return;
        }
        if (Math.abs(dx) < 5 && Math.abs(dy) < 5) return;
        activate(x, y);
      }
      if (ev.cancelable) ev.preventDefault();
      setDrag(d => (d ? { ...d, x, y } : { card, x, y }));
      updateTarget(x, y);
      autoScroll(x, y);
    };

    const finish = doDrop => {
      const p = pendingRef.current;
      if (doDrop && p && p.active) {
        // a quick sideways flick that ends off-column still moves one lane over
        if (!overColRef.current) {
          const dx = (p.lastX ?? p.startX) - p.startX;
          if (Math.abs(dx) > 40) {
            const from = lanes.indexOf(card[laneKey]);
            const to = Math.min(Math.max(from + (dx > 0 ? 1 : -1), 0), lanes.length - 1);
            if (to !== from) { overColRef.current = lanes[to]; overIndexRef.current = 0; }
          }
        }
        drop();
      }
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onCancel);
      cleanupDrag();
    };

    const onUp = ev => { if (ev.pointerId == null || pid == null || ev.pointerId === pid) finish(true); };
    const onCancel = () => { if (dragActiveRef.current) return; finish(false); };

    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onCancel);
    cancelRef.current = () => finish(false);

    if (isTouch) {
      // hold about three quarters of a second, then the card lifts
      lpTimerRef.current = setTimeout(() => {
        const p = pendingRef.current;
        if (p && !p.active) { p.held = true; activate(p.lastX, p.lastY); }
      }, 800);
    }
  };

  return { drag, overCol, overIndex, boardRef, colRefs, listRefs, startPointer };
}

/* Spray lanes are named after the sprayer ("Jason"), while an operator may log
   in as "Jason Badman" — match on the full name or the first name. */
function laneIsUser(lane, user) {
  if (!lane || !user) return false;
  const a = String(lane).trim().toLowerCase();
  const b = String(user).trim().toLowerCase();
  if (a === b) return true;
  if (a === 'to spray' || a === 'unassigned') return false;
  return b.startsWith(a + ' ') || a.startsWith(b + ' ') || a.split(' ')[0] === b.split(' ')[0];
}

const UNASSIGNED = 'Unassigned';

/* Work tasks collapse into families so the picker isn't a wall of 50 buttons.
   First rule that matches wins; anything unmatched stays a standalone button. */
const WORK_GROUPS = [
  { label: 'Wire work',        match: /\bwire\b/i },
  { label: 'French plough',    match: /french plough/i },
  { label: 'V Frame',          match: /^v frame/i },
  { label: 'Undervine / UVC',  match: /\buvc\b|undervine/i },
  { label: 'Mowing & topping', match: /mowing|topping|trimming|mulching/i },
  { label: 'Cover crop',       match: /cover crop|sowing|crimping/i },
  { label: 'Cultivation',      match: /cultivation|deep riper|power harrow|plough/i },
  { label: 'Irrigation',       match: /irrigation/i },
  { label: 'Netting & birds',  match: /netting|bird/i },
  { label: 'Machine work',     match: /^machine /i },
  { label: 'Leaf & fruit',     match: /leaf|fruit thin|colour thin|shoot thin|bud rub|second sets/i },
  { label: 'Pruning',          match: /pruning/i },
  { label: 'Vine care',        match: /young plants|replant|dead vines|weta|rootstock|stays|retrunk/i },
];
function groupWorkTasks(list) {
  const groups = new Map();      // label -> [tasks]
  const loose = [];
  list.forEach(t => {
    const g = WORK_GROUPS.find(r => r.match.test(t));
    if (!g) { loose.push(t); return; }
    if (!groups.has(g.label)) groups.set(g.label, []);
    groups.get(g.label).push(t);
  });
  // a group of one isn't worth hiding behind a dropdown
  const out = [];
  groups.forEach((items, label) => {
    if (items.length < 2) loose.push(...items);
    else out.push({ label, items: items.slice().sort((a, b) => a.localeCompare(b)) });
  });
  out.sort((a, b) => a.label.localeCompare(b.label));
  return { groups: out, loose: loose.sort((a, b) => a.localeCompare(b)) };
}

/* Blocks belong to a vineyard — used to lay the picker out in columns */
const VINEYARDS = ['Eros', 'Hill', 'Winery'];
function vineyardOf(name) {
  const n = String(name || '').toLowerCase();
  if (n.startsWith('eros')) return 'Eros';
  if (n.startsWith('hill')) return 'Hill';
  if (n.startsWith('woolshed')) return 'Woolshed';
  if (/^(sb |wb |winery)/.test(n)) return 'Winery';
  return 'Other';
}

function WorkPlanner({ config, cards, archived, onPersist }) {
  const [task, setTask] = useState('');
  const [picked, setPicked] = useState([]);      // block names
  const [ops, setOps] = useState([]);            // operator codes
  const [due, setDue] = useState('');
  const [note, setNote] = useState('');
  const [msg, setMsg] = useState('');
  const [taskQ, setTaskQ] = useState('');
  const [openGroup, setOpenGroup] = useState(null);

  const blocks = config.blocks || [];
  const operators = config.operators || [];
  const tasks = (config.workTasks || [])
    .filter(t => t.toLowerCase().includes(taskQ.toLowerCase()))
    .slice()
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));

  const toggle = (arr, v, set) => set(arr.includes(v) ? arr.filter(x => x !== v) : [...arr, v]);
  const totalHa = picked.reduce((s, n) => s + numOf((blocks.find(b => b.name === n) || {}).ha), 0);

  // colour state for a block, relative to the task currently selected
  const blockState = name => {
    if (picked.includes(name)) return 'picked';
    if (!task) return 'none';
    const rel = [...(cards || []), ...(archived || [])].filter(c => c.task === task && c.block === name);
    if (!rel.length) return 'none';
    return rel.every(c => c.done) ? 'done' : 'scheduled';
  };

  const create = async () => {
    if (!task || !picked.length) return;
    const chosen = operators.filter(o => ops.includes(o.code));
    const made = picked.map((name, i) => {
      const b = blocks.find(x => x.name === name) || {};
      const op = chosen.length ? chosen[i % chosen.length] : null;   // spread blocks across operators
      const ph = plannedHours(task, name, config);
      return {
        id: uid(), task, block: name, ha: numOf(b.ha), rows: b.rows || '',
        plannedHours: ph == null ? '' : ph,
        assignee: op ? op.name : UNASSIGNED, assigneeCode: op ? op.code : '',
        due, note: note.trim(), done: false, createdAt: Date.now(),
      };
    });
    await onPersist([...(cards || []), ...made]);
    setMsg(`${made.length} job${made.length > 1 ? 's' : ''} scheduled${chosen.length ? ` across ${chosen.length} operator${chosen.length > 1 ? 's' : ''}` : ''}.`);
    setPicked([]); setNote('');
    setTimeout(() => setMsg(''), 5000);
  };

  return (
    <div className={cls.card + ' p-4 space-y-5'}>
      <div>
        <div className="flex items-center justify-between gap-2 mb-2 flex-wrap">
          <label className={cls.label + ' !mb-0'}>1 · Choose the work</label>
          <input value={taskQ} onChange={e => setTaskQ(e.target.value)} placeholder="Filter tasks…"
            className="px-3 py-1.5 rounded-lg border border-stone-300 text-sm w-44 focus:outline-none focus:ring-2 focus:ring-stone-400/40" />
        </div>
        {(() => {
          const btn = active => 'px-3 py-2 rounded-lg border text-sm whitespace-nowrap transition-colors ' +
            (active ? 'bg-stone-900 border-stone-900 text-stone-50 font-medium'
                    : 'bg-white border-stone-300 text-stone-700 hover:border-stone-400');
          // while filtering, show a flat list — grouping only gets in the way
          if (taskQ.trim()) {
            return (
              <div className="flex flex-wrap gap-1.5 p-0.5">
                {tasks.map(t => <button key={t} onClick={() => setTask(t)} className={btn(task === t)}>{t}</button>)}
                {tasks.length === 0 && <p className="text-sm text-stone-400">No task matches that.</p>}
              </div>
            );
          }
          const { groups, loose } = groupWorkTasks(tasks);
          return (
            <div className="p-0.5">
              <div className="flex flex-wrap gap-1.5">
                {groups.map(g => {
                  const holdsSelected = g.items.includes(task);
                  const isOpen = openGroup === g.label;
                  return (
                    <button key={g.label} onClick={() => setOpenGroup(isOpen ? null : g.label)}
                      className={'inline-flex items-center gap-1.5 px-3 py-2 rounded-lg border text-sm whitespace-nowrap transition-colors ' +
                        (holdsSelected ? 'bg-stone-900 border-stone-900 text-stone-50 font-medium'
                          : isOpen ? 'bg-stone-100 border-stone-400 text-stone-900'
                          : 'bg-white border-stone-300 text-stone-700 hover:border-stone-400')}>
                      {g.label}
                      <span className={'text-[11px] rounded-full px-1.5 py-0.5 ' + (holdsSelected ? 'bg-white/20' : 'bg-stone-100 text-stone-500')}>{g.items.length}</span>
                      <ChevronRight size={14} className={'transition-transform ' + (isOpen ? 'rotate-90' : '')} />
                    </button>
                  );
                })}
                {loose.map(t => <button key={t} onClick={() => { setTask(t); setOpenGroup(null); }} className={btn(task === t)}>{t}</button>)}
              </div>
              {openGroup && (
                <div className="mt-2 p-2.5 rounded-xl border border-stone-300 bg-stone-50">
                  <div className="text-[11px] font-semibold uppercase tracking-wide text-stone-500 mb-1.5">{openGroup}</div>
                  <div className="flex flex-wrap gap-1.5">
                    {(groups.find(g => g.label === openGroup) || { items: [] }).items.map(t => (
                      <button key={t} onClick={() => setTask(t)} className={btn(task === t)}>{t}</button>
                    ))}
                  </div>
                </div>
              )}
              {task && <p className="text-[13px] text-stone-500 mt-2">Selected: <b className="text-stone-900">{task}</b></p>}
            </div>
          );
        })()}
      </div>

      <div>
        <div className="flex items-center justify-between gap-2 mb-2 flex-wrap">
          <label className={cls.label + ' !mb-0'}>2 · Choose the blocks</label>
          <div className="flex gap-2 items-center flex-wrap">
            {task && (
              <div className="flex gap-3 items-center text-[12px] text-stone-500 mr-1">
                <span className="inline-flex items-center gap-1.5"><span className="w-3 h-3 rounded border border-sky-300 bg-sky-100 inline-block" /> Selected</span>
                <span className="inline-flex items-center gap-1.5"><span className="w-3 h-3 rounded border border-emerald-200 bg-emerald-50 inline-block" /> Done</span>
                <span className="inline-flex items-center gap-1.5"><span className="w-3 h-3 rounded border border-amber-200 bg-amber-50 inline-block" /> Scheduled</span>
                <span className="inline-flex items-center gap-1.5"><span className="w-3 h-3 rounded border border-stone-300 bg-white inline-block" /> Not scheduled</span>
              </div>
            )}
            <button onClick={() => setPicked(blocks.map(b => b.name))} className="text-xs px-2.5 py-1.5 rounded-lg border border-stone-300 bg-white text-stone-600 hover:bg-stone-50">Select all</button>
            <button onClick={() => setPicked([])} className="text-xs px-2.5 py-1.5 rounded-lg border border-stone-300 bg-white text-stone-600 hover:bg-stone-50">Clear</button>
          </div>
        </div>
        {(() => {
          const groups = {};
          blocks.forEach(b => { const v = vineyardOf(b.name); (groups[v] = groups[v] || []).push(b); });
          const order = [...VINEYARDS, ...Object.keys(groups).filter(k => !VINEYARDS.includes(k)).sort()];
          const cols = order.filter(v => (groups[v] || []).length);
          const blockBtn = b => {
            const st = blockState(b.name);
            const tone =
              st === 'picked'    ? 'bg-sky-100 border-sky-300 text-sky-900 font-medium'
              : st === 'done'      ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
              : st === 'scheduled' ? 'bg-amber-50 border-amber-200 text-amber-800'
              : 'bg-white border-stone-300 text-stone-700 hover:border-stone-400';
            return (
              <button key={b.name} onClick={() => toggle(picked, b.name, setPicked)}
                className={'w-full text-left px-3 py-2 rounded-lg border text-sm transition-colors ' + tone}>
                <span className="truncate block">{b.name}{numOf(b.ha) > 0 && <span className="opacity-60"> · {fmtNum(b.ha)} ha</span>}</span>
              </button>
            );
          };
          return (
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4 p-0.5">
              {cols.map(v => {
                const list = groups[v];
                const ha = list.reduce((sum, b) => sum + numOf(b.ha), 0);
                const allPicked = list.every(b => picked.includes(b.name));
                return (
                  <div key={v} className="rounded-xl border border-stone-200 bg-stone-50/60 p-2.5">
                    <div className="flex items-center justify-between gap-2 mb-2 px-0.5">
                      <span className="text-[13px] font-bold text-stone-900">{v}</span>
                      <button onClick={() => setPicked(allPicked
                        ? picked.filter(n => !list.some(b => b.name === n))
                        : [...new Set([...picked, ...list.map(b => b.name)])])}
                        className="text-[11px] px-2 py-1 rounded-md border border-stone-300 bg-white text-stone-600 hover:bg-stone-100">
                        {allPicked ? 'None' : 'All'}
                      </button>
                    </div>
                    <div className="space-y-1.5">{list.map(blockBtn)}</div>
                    <div className="text-[11px] text-stone-400 mt-2 px-0.5">{list.length} blocks · {fmtNum(Math.round(ha * 100) / 100)} ha</div>
                  </div>
                );
              })}
            </div>
          );
        })()}
        {picked.length > 0 && (() => {
          const hrs = picked.map(n => plannedHours(task, n, config)).filter(h => h != null);
          const total = hrs.reduce((s2, h) => s2 + h, 0);
          const pace = (config.workPace || {})[task];
          return (
            <p className="text-sm text-stone-500 mt-2">
              {picked.length} block{picked.length > 1 ? 's' : ''} · {fmtNum(Math.round(totalHa * 100) / 100)} ha
              {task && total > 0 && (
                <> · about <b className="text-stone-800">{fmtNum(Math.round(total * 10) / 10)} h</b> at {paceLabel(pace)}
                  <span className="text-stone-400"> ({fmtNum(Math.round(total / WORK_DAY_HOURS * 10) / 10)} days of 8 h)</span></>
              )}
              {task && !pace && <span className="text-amber-700"> · no work rate set for this task</span>}
              {task && pace && hrs.length < picked.length && <span className="text-amber-700"> · {picked.length - hrs.length} block{picked.length - hrs.length > 1 ? 's' : ''} missing km or vine numbers</span>}
            </p>
          );
        })()}
      </div>

      <div>
        <label className={cls.label}>3 · Assign operators <span className="font-normal normal-case tracking-normal text-stone-400">— blocks are shared out between them; leave empty for Unassigned</span></label>
        <div className="flex gap-1.5 flex-wrap">
          {operators.map(o => (
            <button key={o.code} onClick={() => toggle(ops, o.code, setOps)}
              className={'px-3 py-2 rounded-lg border text-sm whitespace-nowrap transition-colors ' +
                (ops.includes(o.code) ? 'bg-stone-900 border-stone-900 text-stone-50 font-medium' : 'bg-white border-stone-300 text-stone-700 hover:border-stone-400')}>{o.name}</button>
          ))}
          {operators.length === 0 && <p className="text-sm text-stone-400">Add operators in Setup first.</p>}
        </div>
      </div>

      <div className="flex gap-3 flex-wrap items-end">
        <div><label className={cls.label}>Due date (optional)</label><input type="date" value={due} onChange={e => setDue(e.target.value)} className={cls.input + ' !w-auto'} /></div>
        <div className="flex-1 min-w-[200px]"><label className={cls.label}>Note (optional)</label><input value={note} onChange={e => setNote(e.target.value)} placeholder="Anything the crew should know" className={cls.input} /></div>
        <button onClick={create} disabled={!task || !picked.length} className={cls.primary + ' !py-2.5'}>
          <Plus size={16} /> Schedule {picked.length ? `${picked.length} job${picked.length > 1 ? 's' : ''}` : 'work'}
        </button>
      </div>
      {msg && <div className="text-sm px-3 py-2 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200">{msg}</div>}
    </div>
  );
}

function WorkCard({ card, onToggle, onStart, onRemove, onPointerDown, ghost, onShift, canLeft, canRight }) {
  return (
    <div data-card-id={card.id}
      onPointerDown={onPointerDown ? e => onPointerDown(e, card) : undefined}
      className={'rounded-xl border p-4 select-none ' +
        (onPointerDown ? 'cursor-grab active:cursor-grabbing ' : '') +
        (ghost ? 'shadow-2xl rotate-2 w-[300px] ' : '') +
        (card.done ? 'bg-emerald-50 border-emerald-200' : 'bg-white border-stone-200')}
      style={{ touchAction: onPointerDown ? 'auto' : undefined }}>
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <div className={'font-bold text-[19px] leading-tight ' + (card.done ? 'text-emerald-800' : 'text-stone-900')}>{card.task}</div>
          <div className={'font-semibold text-[16px] mt-0.5 leading-tight ' + (card.done ? 'text-emerald-700' : 'text-stone-700')}>{card.block}</div>
        </div>
        {onRemove && <button onPointerDown={e => e.stopPropagation()} onClick={() => onRemove(card.id)} className="p-1.5 rounded-lg hover:bg-red-50 text-red-400 shrink-0"><Trash2 size={15} /></button>}
      </div>
      <div className="mt-2 space-y-0.5">
        {numOf(card.ha) > 0 && <div className="text-[15px] font-semibold text-stone-800">{fmtNum(card.ha)} ha</div>}
        {card.rows && <div className="text-[15px] font-semibold text-stone-800">Rows {card.rows}</div>}
        {card.due && <div className="text-[15px] font-semibold text-stone-800">Due {card.due}</div>}
        {numOf(card.plannedHours) > 0 && (() => {
          const actual = cardWorkedHours(card);
          const over = actual > 0 && actual > numOf(card.plannedHours) * 1.1;
          const under = actual > 0 && actual < numOf(card.plannedHours) * 0.9;
          return (
            <div className={'text-[14px] font-semibold ' + (over ? 'text-amber-700' : under ? 'text-emerald-700' : 'text-stone-800')}>
              {fmtNum(card.plannedHours)} h planned{actual > 0 ? ` · ${fmtNum(actual)} h actual` : ''}
            </div>
          );
        })()}
      </div>
      {card.note && <div className="text-[12px] text-stone-500 mt-1.5 italic">{card.note}</div>}
      {cardState(card) === 'paused' && (
        <div className="mt-2 inline-flex items-center gap-1.5 text-[11px] font-semibold uppercase tracking-wide text-amber-800 bg-amber-50 border border-amber-200 rounded px-2 py-1">
          <span className="w-1.5 h-1.5 rounded-full bg-amber-500" /> In progress — stopped
        </div>
      )}
      {!card.done && card.lastRow && (
        <div className="mt-2 rounded-lg bg-sky-50 border border-sky-200 px-2.5 py-2">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-sky-700">Pick up from</div>
          <div className="text-[16px] font-bold text-sky-900 leading-tight">after row {card.lastRow}</div>
          <div className="text-[11px] text-sky-700/80">{card.lastRowBy}{card.lastRowAt ? ` · ${card.lastRowAt}` : ''}</div>
        </div>
      )}
      {(() => {
        const sessions = cardSessions(card);
        if (!sessions.length && !card.done) return null;
        const open = sessions.find(x => !x.endTs);
        const worked = cardWorkedMs(card);
        return (
          <div className="mt-2 space-y-0.5">
            {open && (
              <div className="text-[12px] text-sky-700 font-semibold">
                Running since {open.startTime} today{open.startAt !== todayNZ() ? ` (${open.startAt})` : ''}
              </div>
            )}
            {sessions.filter(x => x.endTs).slice(-3).map((x, i) => (
              <div key={i} className="text-[12px] text-stone-500">
                {x.startAt} · {x.startTime}–{x.endTime} <span className="text-stone-400">({fmtDuration(x.endTs - x.startTs)})</span>
              </div>
            ))}
            {sessions.filter(x => x.endTs).length > 3 && (
              <div className="text-[11px] text-stone-400">+{sessions.filter(x => x.endTs).length - 3} earlier day{sessions.filter(x => x.endTs).length - 3 > 1 ? 's' : ''}</div>
            )}
            {worked > 0 && (
              <div className="text-[12.5px] font-semibold text-stone-800">Total {fmtDuration(worked)}{open ? ' so far' : ''}</div>
            )}
            {card.done && (card.doneAt || card.doneBy) && (
              <div className="text-[12px] text-emerald-700 font-medium">
                Finished {card.doneAt}{card.doneTime ? ` ${card.doneTime}` : ''}{card.doneBy ? ` · ${card.doneBy}` : ''}
              </div>
            )}
          </div>
        );
      })()}
      {onToggle && (
        <div className="flex items-center gap-1.5 mt-3">
          {onShift && (
            <button onPointerDown={e => e.stopPropagation()} onClick={() => onShift(card, -1)} disabled={!canLeft}
              title="Move to the column on the left"
              className="px-2.5 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-600 hover:bg-stone-50 disabled:opacity-30 shrink-0">
              <ChevronLeft size={16} />
            </button>
          )}
          {onStart && !card.done && (() => {
            const running = !!openSession(card);
            return (
              <button onPointerDown={e => e.stopPropagation()} onClick={() => onStart(card)}
                className={'flex-1 py-2.5 rounded-lg text-[15px] font-medium border transition-colors ' +
                  (running ? 'bg-sky-600 border-sky-600 text-white' : 'bg-white border-stone-300 text-stone-700 hover:border-sky-500 hover:text-sky-700')}>
                {running ? 'Stop' : (cardSessions(card).length ? 'Start again' : 'Start')}
              </button>
            );
          })()}
          <button onPointerDown={e => e.stopPropagation()} onClick={() => onToggle(card)}
            className={'flex-1 py-2.5 rounded-lg text-[15px] font-medium border transition-colors ' +
              (card.done ? 'bg-emerald-600 border-emerald-600 text-white' : 'bg-white border-stone-300 text-stone-700 hover:border-emerald-500 hover:text-emerald-700')}>
            {card.done ? <span className="inline-flex items-center gap-1.5"><Check size={15} /> Done</span> : 'Done'}
          </button>
          {onShift && (
            <button onPointerDown={e => e.stopPropagation()} onClick={() => onShift(card, 1)} disabled={!canRight}
              title="Move to the column on the right"
              className="px-2.5 py-2.5 rounded-lg border border-stone-300 bg-white text-stone-600 hover:bg-stone-50 disabled:opacity-30 shrink-0">
              <ChevronRight size={16} />
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function WorkBoard({ config, cards, onPersist, onArchive, manager, currentUser }) {
  const [taskFilter, setTaskFilter] = useState('');
  const [stopFor, setStopFor] = useState(null);   // card being stopped
  const [stopRow, setStopRow] = useState('');
  const operatorNames = (config.operators || []).map(o => o.name);
  // everyone sees every lane; for an operator their own column sits first after Unassigned
  const lanes = manager
    ? [UNASSIGNED, ...operatorNames]
    : [UNASSIGNED, ...(operatorNames.includes(currentUser) ? [currentUser] : []), ...operatorNames.filter(n => n !== currentUser)];
  const mine = cards || [];
  const shown = taskFilter ? mine.filter(c => c.task === taskFilter) : mine;
  const tasksPresent = [...new Set(mine.map(c => c.task))].sort();

  const { drag, overCol, overIndex, boardRef, colRefs, listRefs, startPointer } = useKanbanDrag({
    items: cards || [], lanes, laneKey: 'assignee',
    persist: async (next, info) => {
      // keep assigneeCode in step with the lane the card was dropped into
      const op = (config.operators || []).find(o => o.name === (info && info.lane));
      await onPersist(next.map(c => (info && c.id === info.id ? { ...c, assigneeCode: op ? op.code : '' } : c)));
    },
  });

  // on an operator's phone, open with their own column centred on screen
  const mineRef = useRef(null);
  const centred = useRef(false);
  useEffect(() => {
    if (manager || centred.current) return;
    let tries = 0;
    const go = () => {
      if (centred.current) return;
      if (centreLane(boardRef.current, mineRef.current)) { centred.current = true; return; }
      if (tries++ < 12) requestAnimationFrame(go);   // wait for layout
    };
    requestAnimationFrame(go);
  }, [manager, currentUser, (cards || []).length]);

  const grouped = {};
  lanes.forEach(l => (grouped[l] = []));
  shown.forEach(c => {
    const lane = grouped[c.assignee] ? c.assignee : UNASSIGNED;
    (grouped[lane] = grouped[lane] || []).push(c);
  });
  Object.keys(grouped).forEach(l => grouped[l].sort((a, b) => (a.done ? 1 : 0) - (b.done ? 1 : 0)));

  const toggle = async card => {
    await onPersist((cards || []).map(c => c.id === card.id
      ? (!c.done
        ? (() => {
            const sessions = cardSessions(c).map(x => x.endTs ? x
              : { ...x, endTs: Date.now(), endAt: todayNZ(), endTime: nowTimeNZ() });
            return { ...c, sessions, done: true, doneAt: todayNZ(), doneTime: nowTimeNZ(), doneBy: currentUser || c.assignee || '', doneTs: Date.now() };
          })()
        : { ...c, done: false, doneAt: '', doneTime: '', doneBy: '', doneTs: null })
      : c));
  };
  const remove = async id => { await onPersist((cards || []).filter(c => c.id !== id)); };
  // Start opens a spell of work; pressing Stop closes it. A job picked up again
  // the next morning opens a fresh session, so each day is recorded separately.
  const start = async card => {
    // stopping? ask where they got to, so whoever picks it up knows
    if (openSession(card)) { setStopFor(card); setStopRow(card.lastRow || ''); return; }
    await onPersist((cards || []).map(c => {
      if (c.id !== card.id) return c;
      const sessions = cardSessions(c).slice();
      sessions.push({ startTs: Date.now(), startAt: todayNZ(), startTime: nowTimeNZ(), endTs: null, endAt: '', endTime: '', by: currentUser || c.assignee || '' });
      const first = sessions[0];
      return { ...c, sessions, startedAt: first.startAt, startedTime: first.startTime, startedTs: first.startTs, startedBy: first.by };
    }));
  };
  // close the open spell, recording the row reached
  const confirmStop = async (finishRow) => {
    const card = stopFor; if (!card) return;
    await onPersist((cards || []).map(c => {
      if (c.id !== card.id) return c;
      const sessions = cardSessions(c).slice();
      const open = sessions.findIndex(x => !x.endTs);
      if (open >= 0) sessions[open] = { ...sessions[open], endTs: Date.now(), endAt: todayNZ(), endTime: nowTimeNZ(), endRow: finishRow };
      return { ...c, sessions, lastRow: finishRow, lastRowBy: currentUser || c.assignee || '', lastRowAt: todayNZ() };
    }));
    setStopFor(null); setStopRow('');
  };
  // move a card one column left/right — always works, whatever the device
  const shift = async (card, dir) => {
    const from = lanes.indexOf(card.assignee);
    const to = Math.min(Math.max((from < 0 ? 0 : from) + dir, 0), lanes.length - 1);
    if (to === from) return;
    const lane = lanes[to];
    const op = (config.operators || []).find(o => o.name === lane);
    await onPersist((cards || []).map(c => c.id === card.id
      ? { ...c, assignee: lane, assigneeCode: op ? op.code : '' } : c));
  };
  const fileDone = async () => {
    const done = (cards || []).filter(c => c.done);
    if (!done.length) return;
    if (!window.confirm(`Move ${done.length} completed job${done.length > 1 ? 's' : ''} to the Completed record? They stay searchable there.`)) return;
    await onArchive(done);
    await onPersist((cards || []).filter(c => !c.done));
  };

  const doneCount = mine.filter(c => c.done).length;
  const haTotal = mine.reduce((s, c) => s + numOf(c.ha), 0);
  const haDone = mine.filter(c => c.done).reduce((s, c) => s + numOf(c.ha), 0);
  const pct = haTotal > 0 ? Math.round(haDone / haTotal * 100) : (mine.length ? Math.round(doneCount / mine.length * 100) : 0);

  return (
    <div>
      <div className="flex items-center justify-between gap-2 mb-3 flex-wrap">
        <div className="flex items-center gap-2 text-sm text-stone-600">
          <span><b className="text-stone-900">{fmtNum(Math.round(haDone * 100) / 100)}</b> / {fmtNum(Math.round(haTotal * 100) / 100)} ha</span>
          <div className="w-28 h-2 rounded-full bg-stone-200 overflow-hidden">
            <div className="h-full rounded-full" style={{ width: pct + '%', backgroundColor: pct === 100 ? '#059669' : '#57534e' }} />
          </div>
          <span className="tabular-nums font-semibold text-stone-900">{pct}%</span>
          <span className="text-stone-400">· {doneCount}/{mine.length} blocks</span>
        </div>
        <div className="flex gap-2 items-center flex-wrap">
          <select value={taskFilter} onChange={e => setTaskFilter(e.target.value)}
            className="px-2.5 py-2 rounded-lg border border-stone-300 bg-white text-stone-700 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400/40">
            <option value="">All tasks</option>
            {tasksPresent.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
          {manager && doneCount > 0 && <button onClick={fileDone} className={cls.ghost + ' !py-2 !px-3'}>File {doneCount} done</button>}
        </div>
      </div>

      <p className="text-xs text-stone-400 mb-3">Drag a card to reorder it or hand it to another operator. On a phone, press and hold first, then drag. Tick <span className="text-emerald-600 font-medium">Done</span> when finished — it records the date and who did it.</p>

      {mine.length === 0 ? (
        <p className="text-stone-400 text-sm text-center py-10">No work scheduled yet{manager ? ' — pick a task and some blocks above.' : '.'}</p>
      ) : (
        <div ref={boardRef} className="flex gap-4 overflow-x-auto pb-3 -mx-4 px-4">
          {lanes.map(lane => {
            const isMine = !manager && lane === currentUser;
            return (
            <div key={lane} ref={el => { colRefs.current[lane] = el; if (isMine) mineRef.current = el; }}
              className={'shrink-0 w-[320px] rounded-xl transition-colors ' +
                (overCol === lane ? 'bg-stone-200/60 ring-2 ring-stone-400' : isMine ? 'bg-white ring-2 ring-stone-800' : 'bg-stone-50/60')}>
              <div className="flex items-center justify-between mb-2.5 px-3 pt-2.5">
                <span className={'text-[18px] font-bold leading-tight ' + (isMine ? 'text-stone-900' : 'text-stone-800')}>
                  {lane}{isMine && <span className="text-[13px] font-semibold text-stone-500"> · you</span>}
                </span>
                <span className="text-xs text-stone-400 bg-stone-200/70 rounded-full px-2 py-0.5">{(grouped[lane] || []).length}</span>
              </div>
              <div ref={el => { listRefs.current[lane] = el; }} className="space-y-3 px-2.5 pb-2.5 min-h-[80px]">
                {(() => {
                  const visible = grouped[lane] || [];
                  const showInd = !!drag && overCol === lane;
                  const line = <div className="h-1.5 rounded-full bg-emerald-500/80 mx-1" />;
                  const out = [];
                  visible.forEach((c, i) => {
                    if (showInd && overIndex === i) out.push(<div key={'i' + i}>{line}</div>);
                    const isDragged = !!drag && drag.card.id === c.id;
                    out.push(
                      <div key={c.id} className={isDragged ? 'opacity-25' : ''}>
                        <WorkCard card={c} onToggle={toggle} onStart={start}
                          onRemove={manager ? remove : null} onPointerDown={startPointer}
                          onShift={shift} canLeft={lanes.indexOf(c.assignee) > 0}
                          canRight={lanes.indexOf(c.assignee) < lanes.length - 1} />
                      </div>
                    );
                  });
                  if (showInd && (overIndex ?? visible.length) >= visible.length) out.push(<div key="iend">{line}</div>);
                  if (!out.length) out.push(<div key="empty" className="text-xs text-stone-300 text-center py-4">Drop here</div>);
                  return out;
                })()}
              </div>
            </div>
            );
          })}
        </div>
      )}

      {drag && (
        <div className="fixed z-50 pointer-events-none" style={{ left: drag.x - 132, top: drag.y - 40 }}>
          <WorkCard card={drag.card} ghost />
        </div>
      )}

      {stopFor && (
        <div className="fixed inset-0 z-[60] bg-stone-900/40 flex items-end sm:items-center justify-center p-4"
          onClick={() => { setStopFor(null); setStopRow(''); }}>
          <div className="bg-white rounded-2xl w-full max-w-sm p-5" onClick={e => e.stopPropagation()}>
            <h3 className="font-bold text-[17px] text-stone-900">Where did you get to?</h3>
            <p className="text-sm text-stone-500 mt-1">{stopFor.task} · {stopFor.block}{stopFor.rows ? ` (rows ${stopFor.rows})` : ''}</p>
            <label className={cls.label + ' mt-4'}>Last row finished</label>
            <input value={stopRow} onChange={e => setStopRow(e.target.value)} autoFocus
              placeholder="e.g. 42" className={cls.input + ' text-lg'} />
            <p className="text-xs text-stone-400 mt-1.5">This shows on the card, so whoever picks it up next knows where to start.</p>
            <div className="flex gap-2 mt-4">
              <button onClick={() => confirmStop('')} className={cls.ghost + ' flex-1 justify-center !py-2.5'}>Skip</button>
              <button onClick={() => confirmStop(stopRow.trim())} className={cls.primary + ' flex-1 justify-center !py-2.5'}>
                <Check size={16} /> Stop
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* Completed work record — what was done, when, and by whom */
function WorkHistory({ config }) {
  const [archive, setArchive] = useState(null);
  const [live, setLive] = useState([]);
  const [q, setQ] = useState('');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');

  const load = async () => {
    setArchive(await loadJSON(K.workDone, []));
    setLive(await loadJSON(K.work, []));
  };
  useEffect(() => { load(); }, []);

  if (archive === null) return <div className="p-8 text-center text-stone-400">Loading record…</div>;

  // everything completed: filed archive + anything still ticked on the board
  const all = [...archive, ...live.filter(c => c.done)]
    .sort((a, b) => (b.doneTs || 0) - (a.doneTs || 0));
  const iso = c => { if (!c.doneAt) return ''; const [d, m, y] = String(c.doneAt).split('/'); return y ? `${y}-${m}-${d}` : ''; };
  const filtered = all.filter(c => {
    const hay = `${c.task} ${c.block} ${c.doneBy || ''} ${c.assignee || ''}`.toLowerCase();
    if (q && !hay.includes(q.toLowerCase())) return false;
    const d = iso(c);
    if (from && d && d < from) return false;
    if (to && d && d > to) return false;
    return true;
  });
  const haDone = filtered.reduce((s, c) => s + numOf(c.ha), 0);

  const exportXlsx = () => {
    const stampOf = c => {
      if (c.doneTs) { const d = new Date(c.doneTs); const p = n => String(n).padStart(2, '0');
        return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`; }
      return [c.doneAt, c.doneTime].filter(Boolean).join(' ');
    };
    const hoursTaken = c => { const h = cardWorkedHours(c); return h > 0 ? h : ''; };
    const rows = filtered.map(c => ({
      'Date started': c.startedAt || '', 'Time started': c.startedTime || '',
      'Date done': c.doneAt || '', 'Time done': c.doneTime || '',
      'Hours worked': hoursTaken(c), 'Days worked': cardSessions(c).filter(x => x.endTs).length || '',
      'Last row done': c.lastRow || '',
      'Completed at': stampOf(c), 'Completed by': c.doneBy || c.assignee || '',
      Task: c.task, Block: c.block, Hectares: numOf(c.ha), Rows: c.rows || '',
      'Assigned to': c.assignee || '', 'Due date': c.due || '', Note: c.note || '',
    }));
    const wb = XLSX.utils.book_new();
    addSheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{}]), 'Completed work');
    const byPerson = {};
    filtered.forEach(c => {
      const k = c.doneBy || c.assignee || '—';
      byPerson[k] = byPerson[k] || { jobs: 0, ha: 0, hrs: 0 };
      byPerson[k].jobs++; byPerson[k].ha += numOf(c.ha);
      byPerson[k].hrs += numOf(hoursTaken(c));
    });
    const sum = Object.entries(byPerson).map(([k, v]) => ({ Person: k, Jobs: v.jobs, Hectares: Math.round(v.ha * 100) / 100, 'Hours taken': Math.round(v.hrs * 100) / 100 }));
    addSheet(wb, XLSX.utils.json_to_sheet(sum.length ? sum : [{}]), 'By person');
    const daily = [];
    filtered.forEach(c => cardSessions(c).forEach(x => {
      if (!x.endTs) return;
      daily.push({
        Date: x.startAt, Start: x.startTime, Finish: x.endTime,
        Hours: Math.round((x.endTs - x.startTs) / 36000) / 100,
        Person: x.by || c.doneBy || c.assignee || '', Task: c.task, Block: c.block, Hectares: numOf(c.ha),
        'Finished at row': x.endRow || '',
      });
    }));
    daily.sort((a, b) => String(a.Date).split('/').reverse().join('').localeCompare(String(b.Date).split('/').reverse().join('')));
    addSheet(wb, XLSX.utils.json_to_sheet(daily.length ? daily : [{}]), 'Day by day');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = `completed-work_${todayStr()}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  return (
    <div className="space-y-4">
      <div className="flex items-end gap-3 flex-wrap">
        <div className="flex-1 min-w-[180px]"><label className={cls.label}>Search</label>
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Task, block or person…" className={cls.input} /></div>
        <RangePicker from={from} to={to} onChange={(f, t) => { setFrom(f); setTo(t); }} compact />
        <button onClick={load} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /> Refresh</button>
        <button onClick={exportXlsx} disabled={!filtered.length} className={cls.primary + ' !py-2 !px-3'}><Download size={15} /> Export</button>
      </div>

      <div className="text-sm text-stone-600">
        <b className="text-stone-900">{filtered.length}</b> completed job{filtered.length === 1 ? '' : 's'} · {fmtNum(Math.round(haDone * 100) / 100)} ha
      </div>

      {filtered.length === 0 ? (
        <p className="text-stone-400 text-sm text-center py-10">Nothing completed in this range yet.</p>
      ) : (
        <div className="overflow-x-auto border border-stone-200 rounded-xl bg-white">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
                <th className="px-3 py-2.5 font-semibold">Started</th>
                <th className="px-3 py-2.5 font-semibold">Date done</th>
                <th className="px-3 py-2.5 font-semibold">Time</th>
                <th className="px-3 py-2.5 font-semibold">Completed by</th>
                <th className="px-3 py-2.5 font-semibold">Task</th>
                <th className="px-3 py-2.5 font-semibold">Block</th>
                <th className="px-3 py-2.5 font-semibold text-right">Ha</th>
                <th className="px-3 py-2.5 font-semibold">Rows</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(c => (
                <tr key={c.id} className="border-b border-stone-100 last:border-0">
                  <td className="px-3 py-2.5 whitespace-nowrap text-stone-500">{c.startedAt ? `${c.startedAt}${c.startedTime ? ' ' + c.startedTime : ''}` : '—'}</td>
                  <td className="px-3 py-2.5 whitespace-nowrap text-stone-800 font-medium">{c.doneAt || '—'}</td>
                  <td className="px-3 py-2.5 whitespace-nowrap text-stone-600 tabular-nums">{c.doneTime || '—'}</td>
                  <td className="px-3 py-2.5 whitespace-nowrap text-stone-700">{c.doneBy || c.assignee || '—'}</td>
                  <td className="px-3 py-2.5 text-stone-700">{c.task}</td>
                  <td className="px-3 py-2.5 text-stone-900 font-medium whitespace-nowrap">{c.block}</td>
                  <td className="px-3 py-2.5 text-right tabular-nums text-stone-700">{numOf(c.ha) ? fmtNum(c.ha) : '—'}</td>
                  <td className="px-3 py-2.5 text-stone-500 whitespace-nowrap">{c.rows || '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function WorkManager({ config }) {
  const [cards, setCards] = useState(null);
  const [archived, setArchived] = useState([]);
  const [pane, setPane] = useState('board');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const load = async () => {
    setCards(await loadJSON(K.work, []));
    setArchived(await loadJSON(K.workDone, []));
  };
  useEffect(() => { load(); }, []);
  // an operator ticking Done shows up here within a second
  useLiveKey(K.work, v => setCards(v || []));
  useLiveKey(K.workDone, v => setArchived(v || []));
  const persist = async next => { setCards(next); await saveJSON(K.work, next); };
  const archive = async done => {
    const prev = await loadJSON(K.workDone, []);
    const next = [...prev, ...done];
    await saveJSON(K.workDone, next);
    setArchived(next);
  };
  // everything: what's on the board plus everything filed away
  const exportAll = () => {
    const stampOf = c => {
      if (c.doneTs) { const d = new Date(c.doneTs); const p = n => String(n).padStart(2, '0');
        return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`; }
      return [c.doneAt, c.doneTime].filter(Boolean).join(' ');
    };
    const created = c => { if (!c.createdAt) return ''; const d = new Date(c.createdAt); const p = n => String(n).padStart(2, '0');
      return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()}`; };
    const rowOf = (c, where) => ({
      Status: c.done ? 'Done' : (c.assignee && c.assignee !== UNASSIGNED ? 'Assigned' : 'Unassigned'),
      Where: where,
      Task: c.task, Block: c.block, Hectares: numOf(c.ha), Rows: c.rows || '',
      'Assigned to': c.assignee || '', 'Due date': c.due || '',
      Scheduled: created(c),
      'Date done': c.doneAt || '', 'Time done': c.doneTime || '', 'Completed at': stampOf(c),
      'Completed by': c.doneBy || '', Note: c.note || '',
    });
    const all = [
      ...(cards || []).map(c => rowOf(c, 'On the board')),
      ...(archived || []).map(c => rowOf(c, 'Filed')),
    ].sort((a, b) => String(a.Task).localeCompare(String(b.Task)) || String(a.Block).localeCompare(String(b.Block)));

    // per-task summary
    const byTask = {};
    [...(cards || []), ...(archived || [])].forEach(c => {
      const t = byTask[c.task] = byTask[c.task] || { blocks: 0, ha: 0, doneBlocks: 0, doneHa: 0 };
      t.blocks++; t.ha += numOf(c.ha);
      if (c.done) { t.doneBlocks++; t.doneHa += numOf(c.ha); }
    });
    const taskRows = Object.entries(byTask).sort((a, b) => a[0].localeCompare(b[0])).map(([task, t]) => ({
      Task: task, Blocks: t.blocks, 'Blocks done': t.doneBlocks,
      Hectares: Math.round(t.ha * 100) / 100, 'Hectares done': Math.round(t.doneHa * 100) / 100,
      'Percent done': t.ha > 0 ? Math.round(t.doneHa / t.ha * 100) : 0,
    }));

    // per-person summary
    const byWho = {};
    [...(cards || []), ...(archived || [])].forEach(c => {
      const w = c.doneBy || c.assignee || UNASSIGNED;
      const p = byWho[w] = byWho[w] || { blocks: 0, ha: 0, done: 0, doneHa: 0 };
      p.blocks++; p.ha += numOf(c.ha);
      if (c.done) { p.done++; p.doneHa += numOf(c.ha); }
    });
    const whoRows = Object.entries(byWho).sort((a, b) => a[0].localeCompare(b[0])).map(([who, p]) => ({
      Person: who, 'Blocks assigned': p.blocks, 'Blocks done': p.done,
      Hectares: Math.round(p.ha * 100) / 100, 'Hectares done': Math.round(p.doneHa * 100) / 100,
    }));

    const wb = XLSX.utils.book_new();
    addSheet(wb, XLSX.utils.json_to_sheet(all.length ? all : [{}]), 'All work');
    addSheet(wb, XLSX.utils.json_to_sheet(taskRows.length ? taskRows : [{}]), 'By task');
    addSheet(wb, XLSX.utils.json_to_sheet(whoRows.length ? whoRows : [{}]), 'By person');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = `vineyard-work_${todayStr()}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const isoOfMs = ms => { const d = new Date(ms); return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; };
  const exportBoard = () => {
    const inRange = c => {
      if (!from && !to) return true;
      const d = c.createdAt ? isoOfMs(c.createdAt) : '';
      if (!d) return false;
      return (!from || d >= from) && (!to || d <= to);
    };
    const rows = (cards || []).filter(inRange).map(c => {
      const st = cardState(c);
      return {
        Status: st === 'live' ? 'In progress (now)' : st === 'paused' ? 'In progress (stopped)' : st === 'done' ? 'Done' : 'Planned',
        Task: c.task, Block: c.block, Hectares: numOf(c.ha), Rows: c.rows || '',
        'Assigned to': c.assignee || '', 'Scheduled on': c.createdAt ? (() => { const d = new Date(c.createdAt); return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`; })() : '',
        'Due date': c.due || '', 'Planned hours': numOf(c.plannedHours) || '',
        'Hours worked': cardWorkedHours(c) || '', 'Last row done': c.lastRow || '',
        'Date done': c.doneAt || '', 'Completed by': c.doneBy || '', Note: c.note || '',
      };
    });
    const wb = XLSX.utils.book_new();
    addSheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{}]), 'Work board');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = `work-board_${todayStr()}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  if (cards === null) return <div className="p-8 text-center text-stone-400">Loading work…</div>;
  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700"><Layers size={18} /><h2 className="text-lg font-semibold text-stone-900">Vineyard work</h2></div>
        <div className="flex items-center gap-2">
          <div className="inline-flex rounded-lg border border-stone-300 overflow-hidden">
            <button onClick={() => setPane('board')} className={'px-3 py-2 text-sm font-medium ' + (pane === 'board' ? 'bg-stone-900 text-stone-50' : 'bg-white text-stone-600 hover:bg-stone-50')}>Board</button>
            <button onClick={() => setPane('done')} className={'px-3 py-2 text-sm font-medium border-l border-stone-300 ' + (pane === 'done' ? 'bg-stone-900 text-stone-50' : 'bg-white text-stone-600 hover:bg-stone-50')}>Completed</button>
          </div>
          <button onClick={load} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /> Refresh</button>
          <button onClick={exportAll} className={cls.primary + ' !py-2 !px-3'} title="Everything — scheduled, in progress and completed">
            <Download size={15} /> Export all
          </button>
        </div>
      </div>
      {pane === 'board' ? (
        <>
          <WorkPlanner config={config} cards={cards} archived={archived} onPersist={persist} />

          <div className={cls.card + ' p-3'}>
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <span className="text-[13px] font-semibold text-stone-700">Export the board</span>
              <button onClick={() => exportBoard()} className={cls.primary + ' !py-2 !px-3'}><Download size={15} /> Export</button>
            </div>
            <RangePicker from={from} to={to} onChange={(f, t) => { setFrom(f); setTo(t); }} compact />
            <p className="text-[11px] text-stone-400 mt-1.5">
              Everything currently on the board — planned, part done and finished — filtered by the date it was scheduled.
            </p>
          </div>

          <WorkBoard config={config} cards={cards} onPersist={persist} onArchive={archive} manager={true} currentUser="Manager" />
        </>
      ) : <WorkHistory config={config} />}
    </div>
  );
}

function WorkOperator({ config, session }) {
  const [cards, setCards] = useState(null);
  const load = async () => setCards(await loadJSON(K.work, []));
  useEffect(() => { load(); }, []);
  // jobs the manager schedules, or a mate hands over, appear straight away
  useLiveKey(K.work, v => setCards(v || []));
  const persist = async next => { setCards(next); await saveJSON(K.work, next); };
  if (cards === null) return <div className="p-8 text-center text-stone-400">Loading your work…</div>;
  return (
    <div>
      <div className="flex items-center justify-between gap-2 mb-4">
        <div className="flex items-center gap-2 text-stone-700"><Layers size={18} /><h2 className="text-lg font-semibold text-stone-900">My work</h2></div>
        <button onClick={load} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /> Refresh</button>
      </div>
      <WorkBoard config={config} cards={cards} onPersist={persist} manager={false} currentUser={session.name} />
    </div>
  );
}

/* Fortnight check-your-hours reminder.
   Fortnights end on a Sunday — 30 Aug 2026 was one. Operators are reminded
   from the Friday before at 3:30pm until noon on the Monday after. */
const FORTNIGHT_ANCHOR = new Date(2026, 7, 30);   // Sun 30 Aug 2026, local time
// The pay period an operator should be looking at. It stays on the finished
// fortnight until noon on the Monday after it ends — the moment the reminder
// clears — then rolls over to the new one, so their list starts fresh.
function timesheetPeriod(now = nzNow()) {
  const DAY = 86400000, FN = 14 * DAY;
  let end = new Date(FORTNIGHT_ANCHOR);
  end.setHours(0, 0, 0, 0);
  while (now.getTime() > end.getTime() + 1.5 * DAY) end = new Date(end.getTime() + FN);
  const cutover = new Date(end.getTime() + DAY); cutover.setHours(12, 0, 0, 0);   // Monday noon
  if (now > cutover) end = new Date(end.getTime() + FN);
  const start = new Date(end.getTime() - 13 * DAY); start.setHours(0, 0, 0, 0);
  const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const nz = d => `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}`;
  return { start, end, startISO: iso(start), endISO: iso(end), label: `${nz(start)} – ${nz(end)}` };
}

function fortnightWindow(now = nzNow()) {
  const DAY = 86400000, FN = 14 * DAY;
  let end = new Date(FORTNIGHT_ANCHOR);
  end.setHours(0, 0, 0, 0);
  while (now.getTime() > end.getTime() + 1.5 * DAY) end = new Date(end.getTime() + FN);
  const from = new Date(end.getTime() - 2 * DAY); from.setHours(15, 30, 0, 0);   // Friday 3:30pm
  const to = new Date(end.getTime() + DAY); to.setHours(12, 0, 0, 0);            // Monday noon
  return { end, from, to, open: now >= from && now <= to };
}

function OperatorApp({ config, session, onLogout }) {
  const [view, setView] = useState('home');
  const tiles = [
    { id: 'work', label: 'My work', sub: 'Tasks assigned to you', icon: Layers },
    { id: 'spray', label: 'Spray', sub: 'Today’s spray plan', icon: Droplets },
    { id: 'timesheet', label: 'Timesheet', sub: 'Log & check your hours', icon: Clock },
    { id: 'fuel', label: 'Diesel', sub: 'Log a fill and the meter', icon: Fuel },
    { id: 'machines', label: 'My machines', sub: 'Checks and servicing', icon: Truck },
    { id: 'maint', label: 'Maintenance', sub: 'Report a leak, wire or post', icon: Wrench },
    { id: 'hazard', label: 'Hazard', sub: 'Report something unsafe', icon: AlertTriangle },
  ];
  const titles = { work: 'My work', spray: 'Spray plan', timesheet: 'Timesheet', fuel: 'Diesel fill', machines: 'My machines', maint: 'Maintenance', hazard: 'Hazard report' };
  const fw = fortnightWindow();
  // machines assigned to this operator whose checklist has come round
  const [checksDue, setChecksDue] = useState([]);
  useEffect(() => {
    (async () => {
      const rmLog = await loadJSON(K.rm, []);
      const mine = (config.vehicles || []).filter(v => Array.isArray(v.assignedTo) && v.assignedTo.includes(session.code));
      setChecksDue(mine.filter(v => { const c = checkStatus(v, rmLog); return c.tracked && c.due; }));
    })();
  }, [config, session.code, view]);
  const fnEnd = `${String(fw.end.getDate()).padStart(2, '0')}/${String(fw.end.getMonth() + 1).padStart(2, '0')}/${fw.end.getFullYear()}`;

  return (
    <div className="min-h-screen" style={{ backgroundColor: CREAM }}>
      <TopBar siteName={config.siteName} subtitle={view === 'home' ? `Kia ora, ${session.name}` : titles[view]}
        onBack={view === 'home' ? null : () => setView('home')} onLogout={onLogout} />
      <main className="max-w-md mx-auto px-4 py-6 overflow-x-hidden">
        {checksDue.length > 0 && view === 'home' && (
          <button onClick={() => setView('machines')}
            className="w-full text-left mb-4 rounded-xl border border-red-300 bg-red-50 px-4 py-3 flex items-start gap-2.5 hover:bg-red-100 transition-colors">
            <Truck size={18} className="text-red-600 shrink-0 mt-0.5" />
            <div>
              <div className="font-semibold text-red-900 text-[15px]">
                {checksDue.length === 1 ? `${checksDue[0].name} needs its checklist` : `${checksDue.length} machines need their checklist`}
              </div>
              <div className="text-[13px] text-red-800/90 mt-0.5">
                {checksDue.map(v => v.name).join(', ')} — tap to run through it.
              </div>
            </div>
          </button>
        )}

        {fw.open && (
          <button onClick={() => setView('timesheet')}
            className="w-full text-left mb-4 rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 flex items-start gap-2.5 hover:bg-amber-100 transition-colors">
            <Clock size={18} className="text-amber-600 shrink-0 mt-0.5" />
            <div>
              <div className="font-semibold text-amber-900 text-[15px]">Check your hours for the fortnight ending {fnEnd}</div>
              <div className="text-[13px] text-amber-800/90 mt-0.5">
                Go through your timesheet and tell {config.siteName} if anything is wrong — before Monday midday.
              </div>
            </div>
          </button>
        )}
        {view === 'home' && (
          <div>
            <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-stone-500 mb-1">Signed in</p>
            <h1 style={serif} className="text-2xl text-stone-900 mb-6">Kia ora, {session.name}</h1>
            <div className="space-y-3">
              {tiles.map(t => (
                <button key={t.id} onClick={() => setView(t.id)}
                  className="w-full flex items-center gap-4 bg-white border border-stone-200 rounded-xl px-5 py-5 text-left hover:border-stone-400 hover:shadow-sm transition-all">
                  <span className="w-12 h-12 rounded-xl bg-stone-900 text-stone-50 flex items-center justify-center shrink-0"><t.icon size={22} /></span>
                  <span className="min-w-0">
                    <span className="block text-lg font-semibold text-stone-900">{t.label}</span>
                    <span className="block text-sm text-stone-500">{t.sub}</span>
                  </span>
                  <ChevronRight size={20} className="ml-auto text-stone-300 shrink-0" />
                </button>
              ))}
            </div>
          </div>
        )}
        {view === 'work' && <WorkOperator config={config} session={session} />}
        {view === 'spray' && <SprayHub config={config} setConfig={() => {}} manager={false} operatorName={session.name} />}
        {view === 'timesheet' && <TimesheetOperator config={config} session={session} />}
        {view === 'fuel' && <FuelForm config={config} session={session} />}
        {view === 'machines' && <MachineChecks config={config} session={session} />}
        {view === 'maint' && <MaintenanceForm config={config} session={session} />}
        {view === 'hazard' && <HazardForm config={config} session={session} />}
      </main>
    </div>
  );
}

/* ============================================================
   Manager shell
   ============================================================ */
/* ============================================================
   Chemical shed — product master, stock, usage, low-stock alerts
   ============================================================ */
function ChemicalShed({ config, setConfig }) {
  const products = config.products || [];
  const setProducts = next => setConfig({ ...config, products: next });
  const setRow = (i, patch) => setProducts(products.map((p, j) => (j === i ? { ...p, ...patch } : p)));
  const [msg, setMsg] = useState('');
  const [showPaste, setShowPaste] = useState(false);
  const [paste, setPaste] = useState('');

  const [byType, setByType] = useState(null);
  const types = config.sprayTypes || [];

  const loadSprays = async () => {
    const out = {};
    for (const t of types) out[t.key] = await loadJSON(K.sprays(t.key), []);
    setByType(out);
  };
  useEffect(() => { loadSprays(); }, []);

  const lowNow = products.filter(p => p.minStock !== '' && p.minStock != null && numOf(p.stock) < numOf(p.minStock));

  const loadProducts = replace => {
    const aoa = Papa.parse(paste.replace(/\r/g, ''), { skipEmptyLines: true }).data;
    const parsed = [];
    aoa.forEach((row, idx) => {
      const c = (row || []).map(x => String(x == null ? '' : x).trim());
      if (!c.length || !c[0]) return;
      const joined = c.join(' ').toLowerCase();
      if (idx === 0 && (joined.includes('name') || joined.includes('product')) && /stock|rate|min|unit/.test(joined)) return;
      parsed.push({ name: c[0], unit: c[1] || 'L', concentration: c[2] || '', rate: c[3] || '', stock: c[4] || '', minStock: c[5] || '' });
    });
    if (!parsed.length) return;
    setProducts(replace ? parsed : [...products, ...parsed]);
    setPaste(''); setShowPaste(false);
  };

  const deduct = type => {
    if (type.roundDeducted) { setMsg(`The ${type.label} round was already deducted. Load a new plan (Replace) on that board to start the next round.`); return; }
    const cards = ((byType && byType[type.key]) || []).filter(c => c.done);
    if (!cards.length) { setMsg(`No ${type.label} blocks are marked done yet — nothing to deduct.`); return; }
    if (!window.confirm(`Subtract the ${type.label} round usage from shed stock? Do this once, when that round is finished.`)) return;
    const tcfg = { ...config, roundMix: type.roundMix, waterRate: type.waterRate };
    const usage = roundUsage(cards, tcfg);
    const next = products.map(p => ({ ...p, stock: fmtNum(Math.max(0, numOf(p.stock) - (usage[p.name] || 0))) }));
    const low = next.filter(p => p.minStock !== '' && p.minStock != null && numOf(p.stock) < numOf(p.minStock));
    const nextTypes = types.map(t => (t.key === type.key ? { ...t, roundDeducted: true } : t));
    setConfig({ ...config, products: next, sprayTypes: nextTypes });
    if (low.length && config.webhookUrl) {
      postWebhook(config.webhookUrl, {
        type: 'low_stock',
        notifyEmail: config.notifyEmail || '',
        siteName: config.siteName,
        subject: `[${config.siteName}] Low chemical stock — ${low.length} product${low.length > 1 ? 's' : ''}`,
        body: 'These products are at or below their minimum and need reordering:\n\n' +
          low.map(p => `• ${p.name}: ${fmtNum(numOf(p.stock))} ${p.unit || ''} left (min ${fmtNum(numOf(p.minStock))})`).join('\n'),
        products: low.map(p => ({ name: p.name, stock: numOf(p.stock), min: numOf(p.minStock), unit: p.unit || '' })),
      });
    }
    setMsg(low.length
      ? `${type.label} stock updated. ${low.length} product${low.length > 1 ? 's are' : ' is'} below minimum${config.webhookUrl ? ' — a reorder alert was emailed.' : '.'}`
      : `${type.label} stock updated.`);
  };

  return (
    <div className="space-y-5 pb-6">
      <Banner msg={msg} />
      <div className="flex items-center gap-2 text-stone-700"><Beaker size={18} /><h2 className="text-lg font-semibold text-stone-900">Chemical shed</h2></div>

      {lowNow.length > 0 && (
        <div className="flex items-start gap-2 px-3.5 py-3 rounded-lg bg-red-50 border border-red-200 text-red-800 text-sm">
          <AlertTriangle size={16} className="mt-0.5 shrink-0" />
          <span><b>{lowNow.length} product{lowNow.length > 1 ? 's' : ''} below minimum:</b> {lowNow.map(p => p.name).join(', ')}. Time to reorder.</span>
        </div>
      )}

      {/* Product master */}
      <div className={cls.card + ' p-4'}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-semibold text-stone-900">Products & stock</h3>
          <button onClick={() => setShowPaste(v => !v)} className={cls.ghost + ' !py-2 !px-3'}><Upload size={15} /> Load list</button>
        </div>
        {showPaste && (
          <div className="mb-3 p-3 rounded-lg border border-stone-200 bg-stone-50">
            <p className="text-sm text-stone-600 mb-2">Paste columns: <b>Name, Unit, Concentration, Rate/100L, Stock, Min</b>.</p>
            <textarea value={paste} onChange={e => setPaste(e.target.value)} rows={4}
              placeholder={'Roundup UltraMAX\tL\t570 g/L\t1.25\t200\t50\nLI 700\tL\tpenetrant\t0.2\t40\t10'}
              className={cls.input + ' font-mono text-[12px] resize-y'} />
            <div className="flex gap-2 mt-2 justify-end">
              <button onClick={() => loadProducts(false)} className={cls.ghost + ' !py-2 !px-3'}>Add</button>
              <button onClick={() => loadProducts(true)} className={cls.primary + ' !py-2 !px-3'}>Replace</button>
            </div>
          </div>
        )}
        <div className="overflow-x-auto border border-stone-200 rounded-xl">
          <table className="w-full text-sm whitespace-nowrap">
            <thead>
              <tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200 bg-stone-50">
                <th className="px-3 py-2.5 font-semibold">Product</th>
                <th className="px-3 py-2.5 font-semibold">Unit</th>
                <th className="px-3 py-2.5 font-semibold">Category</th>
                <th className="px-3 py-2.5 font-semibold">Rate</th>
                <th className="px-3 py-2.5 font-semibold">Per</th>
                <th className="px-3 py-2.5 font-semibold">In stock</th>
                <th className="px-3 py-2.5 font-semibold">Min</th>
                <th className="px-3 py-2.5 font-semibold text-center">BioGro</th>
                <th className="px-3 py-2.5 font-semibold text-center">Approved</th>
                <th className="px-3 py-2.5 font-semibold">Active ingredients</th>
                <th className="px-3 py-2.5"></th>
              </tr>
            </thead>
            <tbody>
              {products.length === 0 ? (
                <tr><td colSpan={9} className="px-3 py-6 text-center text-stone-400">No products yet — add a row or paste your list.</td></tr>
              ) : products
                .map((p, i) => ({ p, i }))
                .sort((a, b) => {
                  const rank = c => { const n = PRODUCT_CATEGORIES.indexOf(c || ''); return n < 0 ? 99 : n; };
                  return rank(a.p.category) - rank(b.p.category) || String(a.p.name).localeCompare(String(b.p.name));
                })
                .map(({ p, i }, pos, arr) => {
                  const prev = pos > 0 ? arr[pos - 1].p.category || '' : null;
                  const showHead = (p.category || '') !== prev;
                  const low = p.minStock !== '' && p.minStock != null && numOf(p.stock) < numOf(p.minStock);
                const gi = 'px-2 py-1.5 rounded-md border border-stone-200 bg-white text-stone-800 text-[13px] focus:outline-none focus:ring-2 focus:ring-stone-400/40';
                  return (
                    <React.Fragment key={i}>
                    {showHead && (
                      <tr className="bg-stone-100/80 border-b border-stone-200">
                        <td colSpan={11} className="px-3 py-1.5 text-[11px] font-bold uppercase tracking-wide text-stone-600">
                          {p.category || 'Not categorised'}
                        </td>
                      </tr>
                    )}
                    <tr className={'border-b border-stone-100 ' + (low ? 'bg-red-50/50' : '')}>
                    <td className="px-2 py-1.5">
                      <div className="flex items-center gap-1.5">
                        <input value={p.name} onChange={e => setRow(i, { name: e.target.value })} placeholder="Product name" className={gi + ' w-full min-w-[150px] font-medium'} />
                        {low && <span className="text-[10px] font-semibold uppercase tracking-wide text-red-700 bg-red-100 border border-red-200 rounded-full px-1.5 py-0.5 shrink-0">Low</span>}
                      </div>
                    </td>
                    <td className="px-2 py-1.5"><input value={p.unit} onChange={e => setRow(i, { unit: e.target.value })} placeholder="L" className={gi + ' w-16'} /></td>
                    <td className="px-2 py-1.5">
                      <select value={p.category || ''} onChange={e => setRow(i, { category: e.target.value })} className={gi + ' w-40'}>
                        <option value="">—</option>
                        {PRODUCT_CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </td>
                    <td className="px-2 py-1.5"><input value={p.rate} onChange={e => setRow(i, { rate: e.target.value })} inputMode="decimal" className={gi + ' w-20 text-right'} /></td>
                    <td className="px-2 py-1.5">
                      <select value={p.rateBasis || 'per100'} onChange={e => setRow(i, { rateBasis: e.target.value })} className={gi + ' w-28'}>
                        {RATE_BASES.map(b => <option key={b.key} value={b.key}>{b.label}</option>)}
                      </select>
                    </td>
                    <td className="px-2 py-1.5"><input value={p.stock} onChange={e => setRow(i, { stock: e.target.value })} inputMode="decimal" className={gi + ' w-20 text-right'} /></td>
                    <td className="px-2 py-1.5"><input value={p.minStock} onChange={e => setRow(i, { minStock: e.target.value })} inputMode="decimal" className={gi + ' w-20 text-right'} /></td>
                    <td className="px-2 py-1.5 text-center">
                      <button onClick={() => setRow(i, { biogro: !p.biogro })} title="BioGro certified"
                        className={'px-2.5 py-1.5 rounded-md border text-[12px] font-semibold ' +
                          (p.biogro ? 'bg-emerald-600 border-emerald-600 text-white' : 'bg-white border-stone-300 text-stone-400')}>
                        {p.biogro ? 'Yes' : 'No'}
                      </button>
                    </td>
                    <td className="px-2 py-1.5 text-center">
                      <button onClick={() => setRow(i, { approved: !p.approved })} title="Approved for use"
                        className={'px-2.5 py-1.5 rounded-md border text-[12px] font-semibold ' +
                          (p.approved ? 'bg-emerald-600 border-emerald-600 text-white' : 'bg-red-50 border-red-300 text-red-700')}>
                        {p.approved ? 'Yes' : 'Not approved'}
                      </button>
                    </td>
                    <td className="px-2 py-1.5"><input value={p.actives ?? p.concentration ?? ''} onChange={e => setRow(i, { actives: e.target.value })} placeholder="—" className={gi + ' w-full min-w-[180px]'} /></td>
                    <td className="px-2 py-1.5 text-right"><button onClick={() => setProducts(products.filter((_, j) => j !== i))} className="p-1.5 rounded-md hover:bg-red-50 text-red-500"><Trash2 size={15} /></button></td>
                    </tr>
                    </React.Fragment>
                  );
                })}
            </tbody>
          </table>
        </div>
        <button onClick={() => setProducts([...products, { name: '', unit: 'L', concentration: '', rate: '', stock: '', minStock: '' }])} className={cls.ghost + ' !py-2 !px-3 mt-3'}><Plus size={16} /> Add product</button>
      </div>

      {/* Usage per spray round */}
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <h3 className="font-semibold text-stone-900">Usage by round</h3>
        <button onClick={loadSprays} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /> Refresh</button>
      </div>
      {byType === null ? <p className="text-stone-400 text-sm">Loading…</p> : types.map(type => {
        const cards = (byType[type.key] || []).filter(c => c.done);
        const usedNames = (type.roundMix || []).map(m => m.product);
        const tcfg = { ...config, roundMix: type.roundMix, waterRate: type.waterRate };
        const usage = roundUsage(cards, tcfg);
        return (
          <div key={type.key} className={cls.card + ' p-4'}>
            <div className="flex items-center justify-between mb-3 gap-2 flex-wrap">
              <h4 className="font-semibold text-stone-900">{type.label} <span className="text-stone-400 font-normal">· {cards.length} block{cards.length === 1 ? '' : 's'} done</span></h4>
              <button onClick={() => deduct(type)} disabled={type.roundDeducted || cards.length === 0} className={cls.primary + ' !py-2 !px-3'}>
                {type.roundDeducted ? 'Already deducted' : 'Deduct from stock'}
              </button>
            </div>
            {usedNames.length === 0 ? (
              <p className="text-sm text-amber-700">No mix set for this round yet — open the {type.label} board → Round mix.</p>
            ) : (
              <div className="overflow-x-auto border border-stone-200 rounded-xl">
                <table className="w-full text-sm">
                  <thead><tr className="text-left text-[11px] uppercase tracking-wide text-stone-500 border-b border-stone-200">
                    <th className="px-3 py-2.5 font-semibold">Product</th>
                    <th className="px-3 py-2.5 font-semibold text-right">Used</th>
                    <th className="px-3 py-2.5 font-semibold text-right">In stock</th>
                    <th className="px-3 py-2.5 font-semibold text-right">After round</th>
                  </tr></thead>
                  <tbody>
                    {usedNames.map(name => {
                      const p = products.find(p => p.name === name);
                      const u = usage[name] || 0;
                      const stock = p ? numOf(p.stock) : 0;
                      const after = stock - u;
                      const min = p ? numOf(p.minStock) : 0;
                      const unit = p ? (p.unit || '') : '';
                      return (
                        <tr key={name} className="border-b border-stone-100 last:border-0">
                          <td className="px-3 py-2.5 text-stone-900">{name}{!p && <span className="text-red-600 text-xs"> · not in shed</span>}</td>
                          <td className="px-3 py-2.5 text-right font-medium text-stone-900 tabular-nums">{fmtNum(u)} {unit}</td>
                          <td className="px-3 py-2.5 text-right text-stone-600 tabular-nums">{p ? `${fmtNum(stock)} ${unit}` : '—'}</td>
                          <td className={'px-3 py-2.5 text-right tabular-nums font-medium ' + (p && after < min ? 'text-red-700' : 'text-stone-700')}>{p ? `${fmtNum(after)} ${unit}` : '—'}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        );
      })}
      <p className="text-xs text-stone-400">Usage = label rate × area × water rate. Deduct each round once it's finished; loading a new plan (Replace) on a board starts its next round.</p>
    </div>
  );
}

/* ============================================================
   Manager dashboard — weather, low stock, hours, round progress
   ============================================================ */
const WMO = code => {
  const c = Number(code);
  if (c === 0) return 'Clear';
  if (c <= 3) return 'Partly cloudy';
  if (c <= 48) return 'Fog';
  if (c <= 57) return 'Drizzle';
  if (c <= 67) return 'Rain';
  if (c <= 77) return 'Snow';
  if (c <= 82) return 'Showers';
  if (c <= 86) return 'Snow showers';
  if (c <= 99) return 'Thunderstorm';
  return '—';
};
function pickNum(obj, re) {
  let found = null;
  const walk = o => {
    if (found != null || o == null || typeof o !== 'object') return;
    for (const k of Object.keys(o)) {
      const v = o[k];
      if (found != null) return;
      if (re.test(k) && (typeof v === 'number' || (typeof v === 'string' && v.trim() !== '' && !isNaN(numOf(v))))) { found = numOf(v); return; }
      if (v && typeof v === 'object') walk(v);
    }
  };
  walk(obj); return found;
}

/* ============================================================
   Date range picker — shared by every export
   ============================================================ */
const isoOf = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
function rangePreset(which) {
  const now = nzNow();
  if (which === 'all') return { from: '', to: '' };
  if (which === 'week') return { from: mondayOf(now), to: todayStr() };
  if (which === 'fortnight') { const p = timesheetPeriod(now); return { from: p.startISO, to: p.endISO }; }
  if (which === 'lastfortnight') {
    const p = timesheetPeriod(now);
    const end = new Date(p.start); end.setDate(end.getDate() - 1);
    const start = new Date(end); start.setDate(start.getDate() - 13);
    return { from: isoOf(start), to: isoOf(end) };
  }
  if (which === 'month') return { from: isoOf(new Date(now.getFullYear(), now.getMonth(), 1)), to: todayStr() };
  if (which === 'lastmonth') {
    const s = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const e = new Date(now.getFullYear(), now.getMonth(), 0);
    return { from: isoOf(s), to: isoOf(e) };
  }
  if (which === 'season') {
    // NZ vineyard season runs July to June
    const y = now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
    return { from: `${y}-07-01`, to: todayStr() };
  }
  return { from: '', to: '' };
}
const RANGE_PRESETS = [
  ['week', 'This week'], ['fortnight', 'This fortnight'], ['lastfortnight', 'Last fortnight'],
  ['month', 'This month'], ['lastmonth', 'Last month'], ['season', 'This season'], ['all', 'All time'],
];

function RangePicker({ from, to, onChange, compact }) {
  const active = RANGE_PRESETS.find(([k]) => {
    const r = rangePreset(k);
    return r.from === from && r.to === to;
  });
  return (
    <div className={'flex items-end gap-2 flex-wrap ' + (compact ? '' : 'mb-1')}>
      <div className="flex gap-1.5 flex-wrap">
        {RANGE_PRESETS.map(([k, label]) => (
          <button key={k} onClick={() => { const r = rangePreset(k); onChange(r.from, r.to); }}
            className={'px-2.5 py-1.5 rounded-lg border text-[13px] font-medium transition-colors ' +
              (active && active[0] === k ? 'bg-stone-900 border-stone-900 text-stone-50' : 'bg-white border-stone-300 text-stone-600 hover:bg-stone-50')}>
            {label}
          </button>
        ))}
      </div>
      <div className="flex items-end gap-2">
        <div><label className="text-[10px] uppercase tracking-wide text-stone-400 block">From</label>
          <input type="date" value={from} onChange={e => onChange(e.target.value, to)}
            className="px-2.5 py-1.5 rounded-lg border border-stone-300 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400/40" /></div>
        <div><label className="text-[10px] uppercase tracking-wide text-stone-400 block">To</label>
          <input type="date" value={to} onChange={e => onChange(from, e.target.value)}
            className="px-2.5 py-1.5 rounded-lg border border-stone-300 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400/40" /></div>
      </div>
    </div>
  );
}

/* ============================================================
   Gantt — what's running across the vineyard, work and sprays
   ============================================================ */
const DAY_MS = 86400000;
// accepts ms, yyyy-mm-dd, dd/mm/yyyy — returns a Date at midnight, or null
function anyDate(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') { const d = new Date(v); return isNaN(d) ? null : startOfDay(d); }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return startOfDay(new Date(+m[1], +m[2] - 1, +m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return startOfDay(new Date(+m[3], +m[2] - 1, +m[1]));
  const d = new Date(s);
  return isNaN(d) ? null : startOfDay(d);
}
const startOfDay = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const dayLabel = d => `${d.getDate()}/${d.getMonth() + 1}`;

/* Scheduling helpers — used to lay planned blocks out on the timeline.
   Working days only (Mon–Fri), 8 hours a day. */
const isWeekend = d => d.getDay() === 0 || d.getDay() === 6;
// nth working day on or after a date
function workingDayDate(from, n) {
  const d = new Date(from); let left = Math.floor(n);
  while (isWeekend(d)) d.setDate(d.getDate() + 1);
  while (left > 0) { d.setDate(d.getDate() + 1); if (!isWeekend(d)) left--; }
  return d;
}

function GanttPanel({ config }) {
  const [rows, setRows] = useState(null);
  const [running, setRunning] = useState([]);   // blocks with a spell open right now
  const [scope, setScope] = useState('all');   // all | work | spray
  const [open, setOpen] = useState({});        // which rows are expanded
  const toggleRow = k => setOpen(o => ({ ...o, [k]: !o[k] }));

  const build = async () => {
    const out = [];
    // ---- vineyard work, grouped by task ----
    // only what's still on the board: filing a job away clears it from the
    // timeline too (the full record stays under Work ▸ Completed)
    const all = await loadJSON(K.work, []);

    // Forward plan: each operator works their outstanding blocks one after
    // another, 8 h a day, weekends skipped. Gives every block a planned slot.
    const planned = {};
    const queues = {};
    // `all` is already in board order, so pushing in sequence preserves it
    all.filter(c => !c.done).forEach(c => { const w = c.assignee || UNASSIGNED; (queues[w] = queues[w] || []).push(c); });
    const todayD = startOfDay(nzNow());
    // keep the order the operator has them in on their board — dragging a card
    // up the lane moves it up the plan too
    Object.values(queues).forEach(list => {
      let cur = 0;
      list.forEach(c => {
        const ph = numOf(c.plannedHours) || null;
        const remaining = ph != null ? Math.max(0.25, ph - cardWorkedHours(c)) : null;
        const durWd = remaining != null ? remaining / WORK_DAY_HOURS : 0.5;
        planned[c.id] = {
          start: workingDayDate(todayD, cur),
          end: workingDayDate(todayD, Math.max(cur, cur + durWd - 0.001)),
          hours: remaining, noRate: ph == null,
        };
        cur += durWd;
      });
    });

    const byTask = {};
    all.forEach(c => { (byTask[c.task] = byTask[c.task] || []).push(c); });
    Object.entries(byTask).forEach(([task, cards]) => {
      const starts = cards.map(c => anyDate(c.createdAt)).filter(Boolean);
      const ends = cards.map(c => anyDate(c.doneAt) || anyDate(c.due)).filter(Boolean);
      if (!starts.length) return;
      const start = new Date(Math.min(...starts.map(d => d.getTime())));
      const end = ends.length ? new Date(Math.max(...ends.map(d => d.getTime()))) : null;
      const haTotal = cards.reduce((s, c) => s + numOf(c.ha), 0);
      const haDone = cards.filter(c => c.done).reduce((s, c) => s + numOf(c.ha), 0);
      // one child row per block: done bars sit on their real dates, outstanding
      // ones on their planned slot in that operator's queue
      const children = cards.map(c => {
        const state = cardState(c);
        const live = state === 'live';
        const paused = state === 'paused';
        const pl = planned[c.id];
        const doneAt = anyDate(c.doneAt);
        const firstS = cardSessions(c).filter(x => x.startTs).sort((a, b) => a.startTs - b.startTs)[0];
        const startD = c.done
          ? (firstS ? startOfDay(new Date(firstS.startTs)) : (doneAt || start))
          : (live && firstS ? startOfDay(new Date(firstS.startTs)) : (pl ? pl.start : start));
        const endD = c.done ? (doneAt || startD) : (pl ? pl.end : null);
        return {
          label: c.block, who: c.doneBy || c.assignee || '',
          ha: numOf(c.ha), done: !!c.done, live, paused,
          start: startD, end: endD,
          pct: c.done ? 100 : 0,
          hours: pl ? pl.hours : null, noRate: pl ? pl.noRate : false,
          detail: c.done ? `done ${c.doneAt || ''}`.trim()
            : live ? 'working on it now'
            : paused ? `part done${c.lastRow ? ` — to row ${c.lastRow}` : ''}${pl && pl.hours != null ? ` · ${fmtNum(Math.round(pl.hours * 10) / 10)} h left` : ''}`
            : (pl && pl.hours != null ? `${fmtNum(Math.round(pl.hours * 10) / 10)} h planned` : 'no work rate'),
          blocks: [c.block],
          kind: 'work',
        };
      }).sort((a, b) => (a.done ? 1 : 0) - (b.done ? 1 : 0));   // board order, done last
      out.push({
        kind: 'work', label: task, start, end,
        pct: haTotal > 0 ? Math.round(haDone / haTotal * 100) : Math.round(cards.filter(c => c.done).length / cards.length * 100),
        detail: `${cards.length} block${cards.length > 1 ? 's' : ''} · ${fmtNum(Math.round(haTotal * 100) / 100)} ha`,
        blockNames: cards.map(c => c.block),
        doneBlocks: cards.filter(c => c.done).map(c => c.block),
        liveBlocks: cards.filter(c => !c.done && openSession(c)).map(c => c.block),
        todoBlocks: cards.filter(c => !c.done && !openSession(c)).map(c => c.block),
        children,
      });
    });
    // ---- spray rounds, one row per board ----
    for (const t of (config.sprayTypes || [])) {
      const cards = await loadJSON(K.sprays(t.key), []);
      if (!cards.length) continue;
      const tcfg = { ...config, waterRate: t.waterRate };
      const dates = cards.map(c => anyDate((c.fields || {})['Actual date']) || anyDate((c.fields || {}).Date)).filter(Boolean);
      const prog = areaProgress(cards, tcfg);
      const start = dates.length ? new Date(Math.min(...dates.map(d => d.getTime()))) : startOfDay(new Date());
      const end = dates.length && cards.every(c => c.done) ? new Date(Math.max(...dates.map(d => d.getTime()))) : null;
      // one child per sprayer on this board
      const byLane = {};
      cards.forEach(c => { const l = c.status || 'To Spray'; (byLane[l] = byLane[l] || []).push(c); });
      const kids = Object.entries(byLane).map(([lane, laneCards]) => {
        const ld = laneCards.map(c => anyDate((c.fields || {})['Actual date']) || anyDate((c.fields || {}).Date)).filter(Boolean);
        const lp = areaProgress(laneCards, tcfg);
        return {
          label: lane, ha: lp.total, done: laneCards.every(c => c.done), who: '',
          start: ld.length ? new Date(Math.min(...ld.map(d => d.getTime()))) : start,
          end: ld.length && laneCards.every(c => c.done) ? new Date(Math.max(...ld.map(d => d.getTime()))) : null,
          pct: lp.pct,
          detail: `${laneCards.length} block${laneCards.length > 1 ? 's' : ''} · ${fmtNum(lp.total)} ha`,
          blocks: laneCards.map(c => cardBlockName(c)),
          kind: 'spray',
        };
      }).sort((a, b) => String(a.label).localeCompare(String(b.label)));
      out.push({
        kind: 'spray', label: t.label, start, end, pct: prog.pct,
        detail: `${cards.length} block${cards.length > 1 ? 's' : ''} · ${fmtNum(prog.total)} ha`,
        blockNames: cards.map(c => cardBlockName(c)),
        doneBlocks: cards.filter(c => c.done).map(c => cardBlockName(c)),
        todoBlocks: cards.filter(c => !c.done).map(c => cardBlockName(c)),
        children: kids,
      });
    }
    // anything being worked on right now comes first, then the rest lined up
    const liveCount = r => (r.liveBlocks || []).length;
    out.sort((a, b) =>
      (liveCount(b) > 0) - (liveCount(a) > 0) ||
      (a.kind === b.kind ? String(a.label).localeCompare(String(b.label)) : a.kind === 'work' ? -1 : 1));
    setRows(out);
    // who is on the tools right now — an open spell with no finish time
    const live = [];
    (await loadJSON(K.work, [])).forEach(c => {
      const open = openSession(c);
      if (!c.done && open) live.push({
        id: c.id, task: c.task, block: c.block, ha: numOf(c.ha),
        who: open.by || c.assignee || '', since: open.startTime || '', sinceTs: open.startTs,
        lastRow: c.lastRow || '',
      });
    });
    live.sort((a, b) => (a.sinceTs || 0) - (b.sinceTs || 0));
    setRunning(live);

  };
  useEffect(() => { build(); }, [config]);
  useLiveKey(K.work, () => build());
  useLiveKey(K.workDone, () => build());   // filing a job removes its bar

  if (rows === null) return <div className={cls.card + ' p-4 text-sm text-stone-400'}>Loading timeline…</div>;

  const shown = rows.filter(r => scope === 'all' || r.kind === scope);
  const today = startOfDay(new Date());
  if (!shown.length) {
    return (
      <div className={cls.card + ' p-4'}>
        <div className="flex items-center gap-2 text-stone-700 mb-2"><Layers size={17} /><h3 className="font-semibold text-stone-900">Vineyard timeline</h3></div>
        <p className="text-sm text-stone-400">Nothing scheduled yet — plan some work or load a spray round and it'll appear here.</p>
      </div>
    );
  }

  // window: a little before the earliest start, a little after the latest end
  const mins = shown.map(r => r.start.getTime());
  const maxs = shown.map(r => (r.end || today).getTime());
  let from = new Date(Math.min(...mins, today.getTime()));
  let to = new Date(Math.max(...maxs, today.getTime()));
  from = new Date(from.getTime() - 3 * DAY_MS);
  to = new Date(to.getTime() + 4 * DAY_MS);
  const span = Math.max(1, Math.round((to - from) / DAY_MS));
  const pxPerDay = span <= 21 ? 46 : span <= 60 ? 22 : span <= 140 ? 11 : 6;
  const width = span * pxPerDay;
  const xOf = d => ((startOfDay(d) - from) / DAY_MS) * pxPerDay;

  // month bands across the top
  const months = [];
  for (let i = 0; i < span; i++) {
    const d = new Date(from.getTime() + i * DAY_MS);
    const key = `${d.getFullYear()}-${d.getMonth()}`;
    const last = months[months.length - 1];
    if (!last || last.key !== key) months.push({ key, label: d.toLocaleDateString(undefined, { month: 'short', year: '2-digit' }), from: i, days: 1 });
    else last.days++;
  }
  const ticks = [];
  const step = pxPerDay >= 22 ? 1 : pxPerDay >= 11 ? 7 : 14;
  for (let i = 0; i < span; i += step) ticks.push(i);

  return (
    <div className={cls.card + ' p-4'}>
      <div className="flex items-center justify-between gap-2 mb-3 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700"><Layers size={17} /><h3 className="font-semibold text-stone-900">Vineyard timeline</h3></div>
        <div className="flex items-center gap-2">
          <div className="inline-flex rounded-lg border border-stone-300 overflow-hidden text-sm">
            {[['all', 'All'], ['work', 'Work'], ['spray', 'Spray']].map(([k, l]) => (
              <button key={k} onClick={() => setScope(k)}
                className={'px-3 py-1.5 font-medium border-l first:border-l-0 border-stone-300 ' + (scope === k ? 'bg-stone-900 text-stone-50' : 'bg-white text-stone-600 hover:bg-stone-50')}>{l}</button>
            ))}
          </div>
          <button onClick={build} className={cls.ghost + ' !py-1.5 !px-2.5'}><RefreshCw size={15} /></button>
        </div>
      </div>

      {running.length > 0 && (
        <div className="mb-4 rounded-xl border border-sky-200 bg-sky-50/70 p-3">
          <div className="flex items-center gap-2 mb-2">
            <span className="relative flex h-2.5 w-2.5">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-sky-400 opacity-75" />
              <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-sky-500" />
            </span>
            <span className="text-[13px] font-semibold text-sky-900">
              In progress now · {running.length} block{running.length > 1 ? 's' : ''}
            </span>
          </div>
          <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
            {running.map(r => (
              <div key={r.id} className="bg-white border border-sky-200 rounded-lg px-3 py-2">
                <div className="text-[14px] font-bold text-stone-900 truncate">{r.block}</div>
                <div className="text-[12.5px] text-stone-600 truncate">{r.task}</div>
                <div className="text-[11px] text-sky-700 mt-0.5">
                  {r.who}{r.since ? ` · since ${r.since}` : ''}{r.lastRow ? ` · from row ${r.lastRow}` : ''}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="flex">
        {/* fixed labels */}
        <div className="shrink-0 w-64 pr-3 border-r border-stone-200">
          <div className="h-10" />
          {shown.map((r, i) => {
            const rowKey = r.kind + '|' + r.label;
            const isOpen = !!open[rowKey];
            return (
              <div key={i}>
                <button onClick={() => toggleRow(rowKey)}
                  className="h-14 w-full text-left flex flex-col justify-center hover:bg-stone-50 rounded-lg px-1 -mx-1">
                  <div className="flex items-center gap-1.5 min-w-0">
                    <ChevronRight size={14} className={'shrink-0 text-stone-400 transition-transform ' + (isOpen ? 'rotate-90' : '')} />
                    {(r.liveBlocks || []).length > 0 ? (
                  <span className="relative flex h-2.5 w-2.5 shrink-0" title="Work in progress">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-sky-400 opacity-75" />
                    <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-sky-500" />
                  </span>
                ) : (
                  <span className={'w-2 h-2 rounded-full shrink-0 ' + (r.kind === 'spray' ? 'bg-yellow-400' : 'bg-stone-700')} />
                )}
                    <span className="text-[15px] font-bold text-stone-900 truncate" title={(r.blockNames || []).join(', ')}>{r.label}</span>
                  </div>
                  <div className="text-[11px] text-stone-400 truncate pl-[22px]">{r.detail}</div>
                </button>
                {isOpen && (r.children || []).map((c, j) => (
                  <div key={j} className="h-9 flex items-center gap-1.5 pl-[26px] pr-1" title={(c.blocks || []).join(', ')}>
                    {c.live ? (
                      <span className="relative flex h-2 w-2 shrink-0">
                        <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-sky-400 opacity-75" />
                        <span className="relative inline-flex rounded-full h-2 w-2 bg-sky-500" />
                      </span>
                    ) : (
                      <span className={'w-1.5 h-1.5 rounded-full shrink-0 ' + (c.pct === 100 ? 'bg-emerald-500' : c.paused ? 'bg-orange-500' : 'bg-stone-300')} />
                    )}
                    <span className={'text-[13px] font-semibold truncate ' + (c.live ? 'text-sky-900' : 'text-stone-800')}>{c.label}</span>
                    <span className={'text-[11px] shrink-0 ml-auto whitespace-nowrap ' + (c.live ? 'text-sky-700 font-medium' : 'text-stone-400')}>{c.detail}</span>
                  </div>
                ))}
              </div>
            );
          })}
        </div>

        {/* scrolling timeline */}
        <div className="flex-1 overflow-x-auto">
          <div style={{ width }} className="relative">
            {/* months + day ticks */}
            <div className="h-10 relative border-b border-stone-200">
              {months.map(m => (
                <div key={m.key} className="absolute top-0 text-[11px] font-semibold uppercase tracking-wide text-stone-500 px-1"
                  style={{ left: m.from * pxPerDay, width: m.days * pxPerDay }}>{m.label}</div>
              ))}
              {ticks.map(i => (
                <div key={i} className="absolute bottom-0.5 text-[10px] text-stone-400" style={{ left: i * pxPerDay + 2 }}>
                  {dayLabel(new Date(from.getTime() + i * DAY_MS))}
                </div>
              ))}
            </div>

            {/* today marker */}
            {today >= from && today <= to && (
              <div className="absolute top-10 bottom-0 w-px bg-red-400/70 z-10" style={{ left: xOf(today) }}>
                <div className="absolute -top-0.5 -left-1 w-2 h-2 rounded-full bg-red-400" />
              </div>
            )}

            {/* bars */}
            {shown.map((r, i) => {
              const rowKey = r.kind + '|' + r.label;
            const isOpen = !!open[rowKey];
              const bar = (item, height) => {
                const s0 = Math.max(0, xOf(item.start));
                const e0 = item.end ? Math.max(xOf(item.end) + pxPerDay, s0 + pxPerDay) : xOf(today) + pxPerDay;
                return { left: s0, width: Math.max(pxPerDay, e0 - s0), height, ongoing: !item.end };
              };
              const g = bar(r, 30);
              const grid = ticks.map(t => <div key={t} className="absolute top-0 bottom-0 w-px bg-stone-100" style={{ left: t * pxPerDay }} />);
              return (
                <div key={i}>
                  <div className="h-14 relative border-b border-stone-100">
                    {grid}
                    <div className="absolute top-1/2 -translate-y-1/2 rounded-lg overflow-hidden shadow-sm"
                      style={{ left: g.left, width: g.width, height: g.height }}
                      title={`${r.label} — ${r.pct}% done${g.ongoing ? ' (ongoing)' : ''}\n${(r.blockNames || []).join(', ')}`}>
                      <div className={'w-full h-full ' + (r.kind === 'spray' ? 'bg-sky-100 border border-sky-300' : 'bg-stone-200 border border-stone-300')}>
                        <div className={'h-full ' + (r.pct === 100 ? 'bg-emerald-500/70' : r.kind === 'spray' ? 'bg-sky-400/60' : 'bg-stone-500/50')}
                          style={{ width: r.pct + '%' }} />
                      </div>
                      <span className="absolute inset-0 flex items-center px-2 text-[12px] font-semibold text-stone-800 whitespace-nowrap">
                        {r.pct}%
                      </span>
                    </div>
                  </div>
                      {isOpen && (r.children || []).map((c, j) => {
                    const cb = bar(c, 16);
                    return (
                      <div key={j} className="h-9 relative border-b border-stone-100">
                        {grid}
                        <div className="absolute top-1/2 -translate-y-1/2 rounded-md overflow-hidden"
                          style={{ left: cb.left, width: cb.width, height: cb.height }}
                          title={`${c.label} — ${c.pct}% done\n${(c.blocks || []).join(', ')}`}>
                          <div className={'w-full h-full rounded-sm ' +
                            (c.done ? 'bg-emerald-500/80'
                              : c.live ? 'bg-sky-500'
                              : c.paused ? 'bg-orange-500'
                              : c.noRate ? 'bg-stone-400'
                              : c.kind === 'spray' ? 'bg-yellow-400' : 'bg-stone-900')} />
                        </div>
                        <span className="absolute top-1/2 -translate-y-1/2 text-[11px] whitespace-nowrap flex items-center gap-1.5"
                          style={{ left: cb.left + cb.width + 6 }}>
                          <span className={c.live ? 'text-sky-800 font-semibold' : c.paused ? 'text-amber-800 font-medium' : c.done ? 'text-emerald-700' : 'text-stone-700'}>{c.label}</span>
                          {c.paused && <span className="text-[10px] font-semibold uppercase tracking-wide text-amber-800 bg-amber-50 border border-amber-200 rounded px-1">part done</span>}
                          {c.live && (
                            <span className="relative flex h-2 w-2">
                              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-sky-400 opacity-75" />
                              <span className="relative inline-flex rounded-full h-2 w-2 bg-sky-500" />
                            </span>
                          )}
                        </span>
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        </div>
      </div>

      <div className="flex gap-4 items-center mt-3 text-[11px] text-stone-500 flex-wrap">
        <span className="inline-flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-stone-700" /> Work</span>
        <span className="inline-flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-yellow-400" /> Spray</span>
        <span className="inline-flex items-center gap-1.5"><span className="w-3 h-2 rounded bg-emerald-500/70" /> Complete</span>
        <span className="inline-flex items-center gap-1.5"><span className="w-px h-3 bg-red-400" /> Today</span>
        <span className="text-stone-400">Bars run from when the work was scheduled to its due or completion date; the fill is progress by hectares.</span>
      </div>
    </div>
  );
}

/* Growth stage across the vineyard — where every block sits, by E-L stage */
function GrowthSummary({ config }) {
  const [log, setLog] = useState(null);
  const [view, setView] = useState('stage');   // stage | vineyard
  useEffect(() => { (async () => setLog(await loadJSON(K.el, [])))(); }, []);
  useLiveKey(K.el, v => setLog(v || []));

  if (log === null) return <div className={cls.card + ' p-4 text-sm text-stone-400'}>Loading growth stages…</div>;

  const blocks = (config.blocks || []).filter(b => b.name !== 'N/A');
  const latestFor = name => (log || []).filter(r => r.block === name).sort((a, b) => (b.ts || 0) - (a.ts || 0))[0] || null;
  const rows = blocks.map(b => {
    const r = latestFor(b.name);
    const days = r ? Math.floor((Date.now() - r.ts) / 86400000) : null;
    return { block: b.name, ha: numOf(b.ha), vineyard: vineyardOf(b.name), stage: r ? Number(r.stage) : null, label: r ? r.label : '', date: r ? r.date : '', days };
  });
  const recorded = rows.filter(r => r.stage != null);
  const stale = recorded.filter(r => r.days >= 7).length;
  const range = recorded.length
    ? [Math.min(...recorded.map(r => r.stage)), Math.max(...recorded.map(r => r.stage))]
    : null;

  const chip = r => (
    <div key={r.block} className={'flex items-center gap-2 px-2.5 py-1.5 rounded-lg border ' +
      (r.stage == null ? 'bg-stone-50 border-stone-200' : r.days >= 7 ? 'bg-amber-50 border-amber-200' : 'bg-white border-stone-200')}>
      <span className={'w-8 h-8 rounded-lg flex items-center justify-center font-bold text-[14px] shrink-0 ' +
        (r.stage == null ? 'bg-stone-200 text-stone-400' : 'bg-stone-900 text-stone-50')}>
        {r.stage == null ? '—' : r.stage}
      </span>
      <div className="min-w-0">
        <div className="text-[13px] font-semibold text-stone-900 truncate">{r.block}</div>
        <div className={'text-[11px] truncate ' + (r.days >= 7 ? 'text-amber-700' : 'text-stone-400')}>
          {r.stage == null ? 'not recorded' : `${r.days === 0 ? 'today' : `${r.days} d ago`} · ${r.label}`}
        </div>
      </div>
    </div>
  );

  return (
    <div className={cls.card + ' p-4'}>
      <div className="flex items-center justify-between gap-2 mb-3 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700">
          <Layers size={17} /><h3 className="font-semibold text-stone-900">Growth stage</h3>
          {range && <span className="text-[13px] text-stone-500">E-L {range[0]}{range[1] !== range[0] ? `–${range[1]}` : ''} across the vineyard</span>}
        </div>
        <div className="flex items-center gap-2">
          {stale > 0 && (
            <span className="text-[12px] text-amber-700 bg-amber-50 border border-amber-200 rounded-full px-2.5 py-0.5">
              {stale} over a week old
            </span>
          )}
          <div className="inline-flex rounded-lg border border-stone-300 overflow-hidden text-sm">
            <button onClick={() => setView('vineyard')} className={'px-3 py-1.5 font-medium ' + (view === 'vineyard' ? 'bg-stone-900 text-stone-50' : 'bg-white text-stone-600 hover:bg-stone-50')}>By vineyard</button>
            <button onClick={() => setView('stage')} className={'px-3 py-1.5 font-medium border-l border-stone-300 ' + (view === 'stage' ? 'bg-stone-900 text-stone-50' : 'bg-white text-stone-600 hover:bg-stone-50')}>By stage</button>
          </div>
        </div>
      </div>

      {recorded.length === 0 ? (
        <p className="text-sm text-stone-400 py-6 text-center">No growth stages recorded yet — they appear here as the Technical Viticulturist walks the blocks.</p>
      ) : view === 'vineyard' ? (
        (() => {
          const groups = {};
          recorded.forEach(r => { (groups[r.vineyard] = groups[r.vineyard] || []).push(r); });
          const order = [...VINEYARDS, ...Object.keys(groups).filter(k => !VINEYARDS.includes(k)).sort()].filter(v => (groups[v] || []).length);
          return (
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
              {order.map(v => {
                const list = groups[v].filter(x => x.stage != null).slice().sort((a, b) => b.stage - a.stage);
                const rec = list;
                const avg = rec.length ? Math.round(rec.reduce((s, x) => s + x.stage, 0) / rec.length) : null;
                const ha = list.reduce((s, x) => s + x.ha, 0);
                return (
                  <div key={v} className="rounded-xl border border-stone-200 bg-stone-50/60 p-2.5">
                    <div className="flex items-baseline justify-between gap-2 mb-2 px-0.5">
                      <span className="text-[13px] font-bold text-stone-900">{v}</span>
                      <span className="text-[11px] text-stone-500">{avg != null ? `~E-L ${avg}` : '—'} · {fmtNum(Math.round(ha * 10) / 10)} ha</span>
                    </div>
                    <div className="space-y-1.5">{list.map(chip)}</div>
                  </div>
                );
              })}
            </div>
          );
        })()
      ) : (
        (() => {
          const byStage = {};
          recorded.forEach(r => { (byStage[r.stage] = byStage[r.stage] || []).push(r); });
          const stages = Object.keys(byStage).map(Number).sort((a, b) => a - b);
          return (
            <div className="space-y-2.5">
              {stages.map(st => {
                const list = byStage[st].slice().sort((a, b) => a.block.localeCompare(b.block));
                const ha = list.reduce((s, x) => s + x.ha, 0);
                return (
                  <div key={st} className="flex gap-3 items-start">
                    <div className="w-11 shrink-0 text-center">
                      <div className="w-11 h-11 rounded-xl bg-stone-900 text-stone-50 flex items-center justify-center font-bold text-[17px]">{st}</div>
                      <div className="text-[10px] text-stone-400 mt-1">{fmtNum(Math.round(ha * 10) / 10)} ha</div>
                    </div>
                    <div className="min-w-0 flex-1">
                      <div className="text-[13px] font-semibold text-stone-700 mb-1">{elLabel(st)}</div>
                      <div className="flex flex-wrap gap-1.5">
                        {list.map(r => (
                          <span key={r.block} className={'text-[12px] px-2 py-1 rounded-md border ' + (r.days >= 7 ? 'bg-amber-50 border-amber-200 text-amber-900' : 'bg-white border-stone-300 text-stone-700')}>
                            {r.block}
                          </span>
                        ))}
                      </div>
                    </div>
                  </div>
                );
              })}

            </div>
          );
        })()
      )}
    </div>
  );
}

function Dashboard({ config, setConfig, onNavigate }) {
  const weather = config.weather || {};
  const [wx, setWx] = useState(null);
  const [hours, setHours] = useState(null);
  const [rounds, setRounds] = useState(null);
  const [editWx, setEditWx] = useState(false);
  const products = config.products || [];
  const low = products.filter(p => p.minStock !== '' && p.minStock != null && numOf(p.stock) < numOf(p.minStock));

  const [maint, setMaint] = useState([]);
  useEffect(() => { (async () => setMaint(await loadJSON(K.maint, [])))(); }, []);
  useLiveKey(K.maint, v => setMaint(v || []));
  const [tick, setTick] = useState(0);
  useEffect(() => {
    (async () => {
      const monday = mondayOf(nzNow());
      const rh = [];
      for (const op of (config.operators || [])) {
        const ts = await loadJSON(K.ts(op.code), []);
        rh.push({ name: op.name, code: op.code, week: ts.filter(t => t.date >= monday).reduce((s, t) => s + numOf(t.hours), 0) });
      }
      setHours(rh);
      const rd = [];
      for (const t of (config.sprayTypes || [])) {
        const cards = await loadJSON(K.sprays(t.key), []);
        rd.push({ key: t.key, label: t.label, ...areaProgress(cards, config) });
      }
      setRounds(rd);
    })();
  }, [config, tick]);

  // refresh the dashboard when any spray board or timesheet changes anywhere
  useEffect(() => {
    const keys = [
      ...(config.sprayTypes || []).map(t => K.sprays(t.key)),
      ...(config.operators || []).map(o => K.ts(o.code)),
    ];
    const offs = keys.map(k => subscribe(k, () => setTick(n => n + 1)));
    return () => offs.forEach(off => off());
  }, [JSON.stringify((config.sprayTypes || []).map(t => t.key)), JSON.stringify((config.operators || []).map(o => o.code))]);

  const loadWx = async () => {
    setWx(null);
    try {
      if ((weather.stationUrl || '').trim()) {
        const r = await fetch(weather.stationUrl.trim()); const j = await r.json();
        setWx({ ok: true, source: 'Your station', temp: pickNum(j, /temp/i), humidity: pickNum(j, /humid/i), wind: pickNum(j, /wind.?speed|windspeed|wind/i), rain: pickNum(j, /rain|precip/i), desc: '' });
      } else {
        const u = `https://api.open-meteo.com/v1/forecast?latitude=${weather.lat}&longitude=${weather.lon}&current=temperature_2m,relative_humidity_2m,precipitation,wind_speed_10m,weather_code&timezone=auto`;
        const r = await fetch(u); const j = await r.json(); const c = j.current || {};
        setWx({ ok: true, source: 'Open‑Meteo', temp: c.temperature_2m, humidity: c.relative_humidity_2m, wind: c.wind_speed_10m, rain: c.precipitation, desc: WMO(c.weather_code) });
      }
    } catch { setWx({ ok: false }); }
  };
  useEffect(() => { loadWx(); }, [weather.lat, weather.lon, weather.stationUrl]);

  const Metric = ({ icon: Ic, label, value }) => (
    <div className="flex items-center gap-2">
      <Ic size={16} className="text-stone-400 shrink-0" />
      <div><div className="text-stone-900 font-medium leading-none">{value}</div><div className="text-[11px] text-stone-400 mt-0.5">{label}</div></div>
    </div>
  );

  return (
    <div className="space-y-5 pb-6">
      <div className="flex items-center gap-2 text-stone-700"><LayoutDashboard size={18} /><h2 className="text-lg font-semibold text-stone-900">Dashboard</h2></div>

      {/* low stock first — it's the thing that needs acting on */}
      {low.length > 0 && (
        <button onClick={() => onNavigate && onNavigate('shed')}
          className="w-full text-left rounded-2xl border border-red-300 bg-red-50 hover:bg-red-100 transition-colors p-4">
          <div className="flex items-center gap-2 mb-2.5">
            <AlertTriangle size={18} className="text-red-600" />
            <h3 className="font-semibold text-red-900">{low.length} product{low.length > 1 ? 's' : ''} low on stock — reorder</h3>
            <span className="ml-auto text-[13px] font-medium text-red-700 inline-flex items-center gap-1">Open shed <ChevronRight size={15} /></span>
          </div>
          <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
            {low.map(p => (
              <div key={p.name} className="flex justify-between items-baseline gap-2 bg-white border border-red-200 rounded-lg px-3 py-2">
                <span className="text-[15px] font-semibold text-stone-900 truncate">{p.name}</span>
                <span className="text-sm text-red-700 tabular-nums whitespace-nowrap">
                  {fmtNum(numOf(p.stock))} {p.unit} <span className="text-stone-400">/ min {fmtNum(numOf(p.minStock))}</span>
                </span>
              </div>
            ))}
          </div>
        </button>
      )}

      {(() => {
        const open = maint.filter(r => r.status !== 'Done');
        const urgent = open.filter(r => r.urgent);
        if (!open.length) return null;
        return (
          <button onClick={() => onNavigate && onNavigate('maint')}
            className={'w-full text-left rounded-2xl border p-4 transition-colors ' +
              (urgent.length ? 'border-red-300 bg-red-50 hover:bg-red-100' : 'border-amber-300 bg-amber-50 hover:bg-amber-100')}>
            <div className="flex items-center gap-2 mb-2.5">
              <Wrench size={18} className={urgent.length ? 'text-red-600' : 'text-amber-600'} />
              <h3 className={'font-semibold ' + (urgent.length ? 'text-red-900' : 'text-amber-900')}>
                {open.length} maintenance job{open.length > 1 ? 's' : ''} open{urgent.length ? ` · ${urgent.length} urgent` : ''}
              </h3>
              <span className={'ml-auto text-[13px] font-medium inline-flex items-center gap-1 ' + (urgent.length ? 'text-red-700' : 'text-amber-800')}>
                Open maintenance <ChevronRight size={15} />
              </span>
            </div>
            <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
              {open.slice(0, 6).map(r => (
                <div key={r.id} className="bg-white border border-stone-200 rounded-lg px-3 py-2">
                  <div className="text-[14px] font-semibold text-stone-900 truncate">{r.block}{r.rowRef ? ` · ${r.rowRef}` : ''}</div>
                  <div className="text-[12px] text-stone-500 truncate">{r.kind} — {r.detail}</div>
                </div>
              ))}
            </div>
            {open.length > 6 && <div className="text-[12px] text-stone-500 mt-2">and {open.length - 6} more…</div>}
          </button>
        );
      })()}

      {(() => {
        const ex = expiringItems(config.vehicles);
        if (!ex.length) return null;
        const expired = ex.filter(x => x.expired);
        return (
          <button onClick={() => onNavigate && onNavigate('fleet')}
            className={'w-full text-left rounded-2xl border p-4 transition-colors ' +
              (expired.length ? 'border-red-300 bg-red-50 hover:bg-red-100' : 'border-amber-300 bg-amber-50 hover:bg-amber-100')}>
            <div className="flex items-center gap-2 mb-2.5">
              <Truck size={18} className={expired.length ? 'text-red-600' : 'text-amber-600'} />
              <h3 className={'font-semibold ' + (expired.length ? 'text-red-900' : 'text-amber-900')}>
                {expired.length ? `${expired.length} REGO/WOF expired` : `${ex.length} REGO/WOF due within a month`}
              </h3>
              <span className={'ml-auto text-[13px] font-medium inline-flex items-center gap-1 ' + (expired.length ? 'text-red-700' : 'text-amber-800')}>
                Open fleet <ChevronRight size={15} />
              </span>
            </div>
            <div className="grid gap-2 sm:grid-cols-2 lg:grid-cols-3">
              {ex.slice(0, 6).map((x, i) => (
                <div key={i} className="bg-white border border-stone-200 rounded-lg px-3 py-2 flex items-baseline justify-between gap-2">
                  <span className="text-[14px] font-semibold text-stone-900 truncate">{x.vehicle}</span>
                  <span className={'text-[12px] whitespace-nowrap ' + (x.expired ? 'text-red-700 font-semibold' : 'text-amber-800')}>
                    {x.what} {x.expired ? `expired ${x.label}` : `${x.days} d · ${x.label}`}
                  </span>
                </div>
              ))}
            </div>
          </button>
        );
      })()}

      <GanttPanel config={config} />

      <div className="grid gap-4 lg:grid-cols-2">
        {/* Weather */}
        <div className={cls.card + ' p-4'}>
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2 text-stone-700"><Cloud size={17} /><h3 className="font-semibold text-stone-900">Weather</h3>
              <span className="text-[12px] text-stone-400">· {weather.label || 'Vineyard'}</span></div>
            <div className="flex gap-1">
              <button onClick={loadWx} className="p-2 rounded-lg hover:bg-stone-100 text-stone-500"><RefreshCw size={15} /></button>
              <button onClick={() => setEditWx(v => !v)} className="p-2 rounded-lg hover:bg-stone-100 text-stone-500"><Settings size={15} /></button>
            </div>
          </div>
          {wx === null ? <p className="text-sm text-stone-400">Loading conditions…</p>
            : !wx.ok ? <p className="text-sm text-amber-700">Couldn't reach the weather feed right now. It works on the deployed app; if you've set a station URL, check it returns JSON.</p>
              : (
                <div>
                  <div className="flex items-end gap-3 mb-3">
                    <span style={serif} className="text-4xl font-bold text-stone-900 leading-none">{wx.temp != null ? `${fmtNum(wx.temp)}°` : '—'}</span>
                    <span className="text-stone-500 mb-1">{wx.desc || ''}</span>
                  </div>
                  <div className="grid grid-cols-3 gap-3">
                    <Metric icon={Droplet} label="Humidity" value={wx.humidity != null ? `${fmtNum(wx.humidity)}%` : '—'} />
                    <Metric icon={Wind} label="Wind" value={wx.wind != null ? `${fmtNum(wx.wind)} km/h` : '—'} />
                    <Metric icon={Droplets} label="Rain" value={wx.rain != null ? `${fmtNum(wx.rain)} mm` : '—'} />
                  </div>
                  <div className="text-[11px] text-stone-400 mt-3">Source: {wx.source}</div>
                </div>
              )}
          {editWx && (
            <div className="mt-3 pt-3 border-t border-stone-200 space-y-2">
              <p className="text-xs text-stone-500">Default uses the vineyard's coordinates. To use your own weather station, paste a URL that returns its JSON readings.</p>
              <div className="flex gap-2">
                <div className="flex-1"><label className="text-[10px] uppercase tracking-wide text-stone-400">Latitude</label><input value={weather.lat ?? ''} onChange={e => setConfig({ ...config, weather: { ...weather, lat: e.target.value } })} className={cls.input + ' !py-2'} /></div>
                <div className="flex-1"><label className="text-[10px] uppercase tracking-wide text-stone-400">Longitude</label><input value={weather.lon ?? ''} onChange={e => setConfig({ ...config, weather: { ...weather, lon: e.target.value } })} className={cls.input + ' !py-2'} /></div>
              </div>
              <div><label className="text-[10px] uppercase tracking-wide text-stone-400">Weather station JSON URL (optional)</label><input value={weather.stationUrl ?? ''} onChange={e => setConfig({ ...config, weather: { ...weather, stationUrl: e.target.value } })} placeholder="https://…" className={cls.input + ' !py-2'} /></div>
            </div>
          )}
        </div>

        {/* Low stock */}
        <div className={cls.card + ' p-4'}>
          <div className="flex items-center gap-2 text-stone-700 mb-3"><Beaker size={17} /><h3 className="font-semibold text-stone-900">Chemical shed</h3></div>
          {(() => {
            const unapproved = products.filter(p => p.approved === false);
            if (!unapproved.length) return null;
            return (
              <div className="flex items-start gap-2 text-sm text-red-700 mb-2.5">
                <AlertTriangle size={16} className="shrink-0 mt-0.5" />
                <span>{unapproved.map(p => p.name).join(', ')} — not approved for use</span>
              </div>
            );
          })()}
          {low.length === 0 ? (
            <div className="flex items-center gap-2 text-sm text-emerald-700"><Check size={16} /> All products above their minimum.</div>
          ) : (
            <div className="space-y-2">
              <div className="flex items-center gap-2 text-sm text-red-700 font-medium"><AlertTriangle size={16} /> {low.length} product{low.length > 1 ? 's' : ''} to reorder</div>
              {low.map(p => (
                <div key={p.name} className="flex justify-between items-center text-sm px-3 py-2 rounded-lg bg-red-50 border border-red-200">
                  <span className="text-stone-800">{p.name}</span>
                  <span className="text-red-700 tabular-nums">{fmtNum(numOf(p.stock))} {p.unit} <span className="text-stone-400">/ min {fmtNum(numOf(p.minStock))}</span></span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Operator hours */}
        <div className={cls.card + ' p-4'}>
          <div className="flex items-center gap-2 text-stone-700 mb-3"><Clock size={17} /><h3 className="font-semibold text-stone-900">Operator hours <span className="text-stone-400 font-normal text-sm">· this week</span></h3></div>
          {hours === null ? <p className="text-sm text-stone-400">Loading…</p> : hours.length === 0 ? <p className="text-sm text-stone-400">No operators yet.</p> : (() => {
            const max = Math.max(1, ...hours.map(h => h.week));
            return (
              <div className="space-y-2.5">
                {hours.map(h => (
                  <div key={h.code}>
                    <div className="flex justify-between text-sm mb-1"><span className="text-stone-700">{h.name}</span><span className="text-stone-900 font-medium tabular-nums">{fmtNum(h.week)} h</span></div>
                    <div className="h-2 rounded-full bg-stone-100 overflow-hidden"><div className="h-full rounded-full bg-stone-700" style={{ width: (h.week / max * 100) + '%' }} /></div>
                  </div>
                ))}
              </div>
            );
          })()}
        </div>

        {/* Spray rounds */}
        <div className={cls.card + ' p-4'}>
          <div className="flex items-center gap-2 text-stone-700 mb-3"><Droplets size={17} /><h3 className="font-semibold text-stone-900">Spray rounds <span className="text-stone-400 font-normal text-sm">· by area</span></h3></div>
          {rounds === null ? <p className="text-sm text-stone-400">Loading…</p> : (
            <div className="space-y-3">
              {rounds.map(r => (
                <div key={r.key}>
                  <div className="flex justify-between text-sm mb-1"><span className="text-stone-700">{r.label}</span><span className="text-stone-500 tabular-nums">{fmtNum(r.done)} / {fmtNum(r.total)} ha · <span className="text-stone-900 font-medium">{r.pct}%</span></span></div>
                  <div className="h-2.5 rounded-full bg-stone-100 overflow-hidden"><div className="h-full rounded-full transition-all" style={{ width: r.pct + '%', backgroundColor: r.pct === 100 ? '#059669' : '#57534e' }} /></div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <GrowthSummary config={config} />
    </div>
  );
}

/* ============================================================
   Technical viticulturist console
   ============================================================ */
function ELStages({ config, session }) {
  const [log, setLog] = useState(null);
  const [active, setActive] = useState(null);     // block being scored
  const [showAll, setShowAll] = useState(false);
  const [note, setNote] = useState('');
  const [msg, setMsg] = useState('');
  const [recordDate, setRecordDate] = useState(todayStr());
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');

  const load = async () => setLog(await loadJSON(K.el, []));
  useEffect(() => { load(); }, []);
  useLiveKey(K.el, v => setLog(v || []));

  if (log === null) return <div className="p-8 text-center text-stone-400">Loading…</div>;

  const blocks = (config.blocks || []).filter(b => b.name !== 'N/A');
  const latestFor = name => (log || []).filter(r => r.block === name).sort((a, b) => (b.ts || 0) - (a.ts || 0))[0] || null;
  // most blocks sit within a stage or two of each other — start the picker there
  const median = (() => {
    const cur = blocks.map(b => latestFor(b.name)).filter(Boolean).map(r => Number(r.stage));
    if (!cur.length) return null;
    cur.sort((a, b) => a - b);
    return cur[Math.floor(cur.length / 2)];
  })();

  const nzOf = iso => { const [y, m, d] = String(iso).split('-'); return y ? `${d}/${m}/${y}` : ''; };
  const tsOf = iso => { const [y, m, d] = String(iso).split('-').map(Number); return y ? new Date(y, m - 1, d, 12, 0).getTime() : Date.now(); };

  const record = async (block, stage) => {
    const iso = recordDate || todayStr();
    const entry = {
      id: uid(), block, stage: Number(stage), label: elLabel(stage),
      date: nzOf(iso), dateISO: iso, time: iso === todayStr() ? nowTimeNZ() : '',
      ts: iso === todayStr() ? Date.now() : tsOf(iso),
      by: session.name, note: note.trim(),
    };
    const next = [entry, ...(log || [])];
    await saveJSON(K.el, next); setLog(next);
    setMsg(`${block} — E-L ${stage} on ${entry.date}`);
    setActive(null); setShowAll(false); setNote(''); setRecordDate(todayStr());
    setTimeout(() => setMsg(''), 3000);
  };

  // fix a record that went in wrong — change its date or stage, or remove it
  const amend = async (id, patch) => {
    const next = (log || []).map(r => {
      if (r.id !== id) return r;
      const merged = { ...r, ...patch };
      if (patch.dateISO) { merged.date = nzOf(patch.dateISO); merged.ts = tsOf(patch.dateISO); }
      if (patch.stage != null) { merged.stage = Number(patch.stage); merged.label = elLabel(patch.stage); }
      return merged;
    });
    await saveJSON(K.el, next); setLog(next);
  };
  const removeRecord = async id => {
    if (!window.confirm('Delete this reading?')) return;
    const next = (log || []).filter(r => r.id !== id);
    await saveJSON(K.el, next); setLog(next);
  };

  const exportXlsx = () => {
    const inRange = r => (!from || (r.dateISO || '') >= from) && (!to || (r.dateISO || '') <= to);
    const rows = (log || []).filter(inRange).map(r => ({
      Date: r.date, Time: r.time, Block: r.block, 'E-L stage': r.stage, Description: r.label, By: r.by, Note: r.note || '',
    }));
    // one column per block, one row per week — the shape she'll want for trends
    const wb = XLSX.utils.book_new();
    addSheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{}]), 'E-L records');
    const current = blocks.map(b => { const r = latestFor(b.name); return { Block: b.name, 'E-L stage': r ? r.stage : '', Description: r ? r.label : '', 'Last checked': r ? r.date : '' }; });
    addSheet(wb, XLSX.utils.json_to_sheet(current), 'Current stage');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = `el-stages_${todayStr()}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  if (active) {
    const last = latestFor(active);
    const anchorStage = last ? Number(last.stage) : (median != null ? median : 1);
    const idx = Math.max(0, EL_STAGES.findIndex(x => x[0] === anchorStage));
    const near = showAll ? EL_STAGES : EL_STAGES.slice(Math.max(0, idx - 1), idx + 6);
    return (
      <div className="space-y-4 pb-6">
        <div className="flex items-center gap-2">
          <button onClick={() => { setActive(null); setShowAll(false); }} className="p-1.5 -ml-1.5 rounded-lg hover:bg-stone-100 text-stone-500"><ChevronLeft size={20} /></button>
          <h2 className="text-lg font-semibold text-stone-900">{active}</h2>
          {last && <span className="text-sm text-stone-500 ml-auto">now E-L {last.stage} · {last.date}</span>}
        </div>
        <p className="text-sm text-stone-500">
          {last ? 'Starting from where this block was last time.' : median != null ? 'Starting from where the rest of the vineyard is.' : 'Pick the stage.'}
        </p>
        <div className="flex items-end gap-2">
          <div><label className={cls.label}>Reading date</label>
            <input type="date" value={recordDate} onChange={e => setRecordDate(e.target.value)} className={cls.input + ' !w-auto'} /></div>
          {recordDate !== todayStr() && (
            <button onClick={() => setRecordDate(todayStr())} className="text-xs text-stone-500 underline pb-3">Back to today</button>
          )}
        </div>
        <div className="space-y-1.5">
          {near.map(([code, label]) => (
            <button key={code} onClick={() => record(active, code)}
              className={'w-full text-left px-3 py-3 rounded-xl border transition-colors flex items-center gap-3 ' +
                (last && Number(last.stage) === code ? 'bg-sky-50 border-sky-300' : 'bg-white border-stone-300 hover:border-stone-400')}>
              <span className="w-9 h-9 rounded-lg bg-stone-900 text-stone-50 flex items-center justify-center font-bold text-[15px] shrink-0">{code}</span>
              <span className="text-[15px] text-stone-800 leading-snug">{label}</span>
            </button>
          ))}
        </div>
        <button onClick={() => setShowAll(v => !v)} className={cls.ghost + ' !py-2 !px-3'}>
          {showAll ? 'Show nearby stages only' : 'Show all stages'}
        </button>
        <div><label className={cls.label}>Note (optional)</label>
          <input value={note} onChange={e => setNote(e.target.value)} className={cls.input} placeholder="Anything worth recording" /></div>

        {(() => {
          const history = (log || []).filter(r => r.block === active).sort((a, b) => (b.ts || 0) - (a.ts || 0));
          if (!history.length) return null;
          return (
            <div className="pt-2">
              <h3 className="text-base font-semibold text-stone-900 mb-2">Previous readings</h3>
              <p className="text-xs text-stone-400 mb-2">Change a date or stage here if one went in wrong.</p>
              <div className="space-y-2">
                {history.map(r => (
                  <div key={r.id} className="flex items-center gap-2 bg-white border border-stone-200 rounded-lg px-2.5 py-2 flex-wrap">
                    <input type="date" value={r.dateISO || ''} onChange={e => amend(r.id, { dateISO: e.target.value })}
                      className="px-2 py-1.5 rounded-md border border-stone-200 text-[13px] focus:outline-none focus:ring-2 focus:ring-stone-400/40" />
                    <select value={r.stage} onChange={e => amend(r.id, { stage: e.target.value })}
                      className="px-2 py-1.5 rounded-md border border-stone-200 text-[13px] max-w-[190px] focus:outline-none focus:ring-2 focus:ring-stone-400/40">
                      {EL_STAGES.map(([code, label]) => <option key={code} value={code}>{code} — {label}</option>)}
                    </select>
                    <span className="text-[11px] text-stone-400">{r.by}{r.time ? ` · ${r.time}` : ''}</span>
                    <button onClick={() => removeRecord(r.id)} className="p-1.5 rounded-md hover:bg-red-50 text-red-400 ml-auto"><Trash2 size={15} /></button>
                  </div>
                ))}
              </div>
            </div>
          );
        })()}
      </div>
    );
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700"><Layers size={18} /><h2 className="text-lg font-semibold text-stone-900">E-L stages</h2></div>
        <div className="flex gap-2 items-center flex-wrap">
          <button onClick={load} className={cls.ghost + ' !py-2 !px-3'}><RefreshCw size={15} /></button>
          <button onClick={exportXlsx} className={cls.primary + ' !py-2 !px-3'}><Download size={15} /> Export</button>
        </div>
      </div>
      {msg && <div className="text-sm px-3 py-2 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200">{msg}</div>}
      <RangePicker from={from} to={to} onChange={(f, t) => { setFrom(f); setTo(t); }} compact />
      <p className="text-sm text-stone-500">Tap a block to record where it's at. Blocks not checked in the last seven days are marked.</p>

      {(() => {
        const groups = {};
        blocks.forEach(b => { const v = vineyardOf(b.name); (groups[v] = groups[v] || []).push(b); });
        const order = [...VINEYARDS, ...Object.keys(groups).filter(k => !VINEYARDS.includes(k)).sort()].filter(v => (groups[v] || []).length);
        return (
          <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
            {order.map(v => (
              <div key={v} className="rounded-xl border border-stone-200 bg-stone-50/60 p-2.5">
                <div className="text-[13px] font-bold text-stone-900 mb-2 px-0.5">{v}</div>
                <div className="space-y-1.5">
                  {groups[v].map(b => {
                    const r = latestFor(b.name);
                    const days = r ? Math.floor((Date.now() - r.ts) / 86400000) : null;
                    const stale = days == null || days >= 7;
                    return (
                      <button key={b.name} onClick={() => setActive(b.name)}
                        className={'w-full text-left px-3 py-2.5 rounded-lg border transition-colors ' +
                          (stale ? 'bg-white border-amber-300 hover:border-amber-400' : 'bg-white border-stone-300 hover:border-stone-400')}>
                        <div className="flex items-center gap-2">
                          <span className="text-[14px] font-semibold text-stone-900 truncate flex-1">{b.name}</span>
                          {r
                            ? <span className="text-[15px] font-bold text-stone-900 shrink-0">{r.stage}</span>
                            : <span className="text-[12px] text-stone-400 shrink-0">—</span>}
                        </div>
                        <div className={'text-[11px] mt-0.5 truncate ' + (stale ? 'text-amber-700' : 'text-stone-400')}>
                          {r ? `${days === 0 ? 'today' : `${days} d ago`} · ${r.label}` : 'not recorded yet'}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        );
      })()}
    </div>
  );
}

function DiseaseMonitor({ config, session }) {
  const [log, setLog] = useState(null);
  const [block, setBlock] = useState('');
  const [disease, setDisease] = useState('');
  const [incidence, setIncidence] = useState('');
  const [foundOn, setFoundOn] = useState('');
  const [severity, setSeverity] = useState('');
  const [rowRef, setRowRef] = useState('');
  const [note, setNote] = useState('');
  const [pin, setPin] = useState(null);
  const [pinMsg, setPinMsg] = useState('');
  const [msg, setMsg] = useState('');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');

  const load = async () => setLog(await loadJSON(K.disease, []));
  useEffect(() => { load(); }, []);
  useLiveKey(K.disease, v => setLog(v || []));

  const dropPin = () => {
    setPinMsg('Getting your position…');
    if (!navigator.geolocation) { setPinMsg('This device can’t give a location.'); return; }
    navigator.geolocation.getCurrentPosition(
      pos => { setPin({ lat: pos.coords.latitude, lon: pos.coords.longitude, acc: Math.round(pos.coords.accuracy) }); setPinMsg(''); },
      err => setPinMsg(err.code === 1 ? 'Location blocked — allow it for this site.' : 'Couldn’t get a position.'),
      { enableHighAccuracy: true, timeout: 10000 }
    );
  };

  const valid = block && disease && incidence;
  const save = async () => {
    if (!valid) return;
    const entry = {
      id: uid(), block, disease, incidence, foundOn: foundOn || 'None', severity: severity || '0%',
      rowRef: rowRef.trim(), note: note.trim(), pin,
      date: todayNZ(), dateISO: todayStr(), time: nowTimeNZ(), ts: Date.now(), by: session.name,
    };
    const next = [entry, ...(log || [])];
    await saveJSON(K.disease, next); setLog(next);
    setMsg(`${disease} recorded at ${block}.`);
    setDisease(''); setIncidence(''); setFoundOn(''); setSeverity(''); setRowRef(''); setNote(''); setPin(null);
    setTimeout(() => setMsg(''), 4000);
  };

  const removeFinding = async id => {
    if (!window.confirm('Delete this finding?')) return;
    const next = (log || []).filter(r => r.id !== id);
    await saveJSON(K.disease, next); setLog(next);
  };
  // correct the date if it went in on the wrong day
  const amendDate = async (id, iso) => {
    const nz = (() => { const [y, m, d] = String(iso).split('-'); return y ? `${d}/${m}/${y}` : ''; })();
    const ts = (() => { const [y, m, d] = String(iso).split('-').map(Number); return y ? new Date(y, m - 1, d, 12, 0).getTime() : Date.now(); })();
    const next = (log || []).map(r => (r.id === id ? { ...r, dateISO: iso, date: nz, ts } : r));
    await saveJSON(K.disease, next); setLog(next);
  };

  const exportXlsx = () => {
    const inRange = r => (!from || (r.dateISO || '') >= from) && (!to || (r.dateISO || '') <= to);
    const rows = (log || []).filter(inRange).map(r => ({
      Date: r.date, Time: r.time, Block: r.block, 'Row / bay': r.rowRef || '', Disease: r.disease,
      Incidence: r.incidence, 'Found on': r.foundOn, Severity: r.severity,
      Latitude: r.pin ? r.pin.lat : '', Longitude: r.pin ? r.pin.lon : '',
      'Map link': r.pin ? `https://www.google.com/maps?q=${r.pin.lat},${r.pin.lon}` : '',
      By: r.by, Note: r.note || '',
    }));
    const wb = XLSX.utils.book_new();
    addSheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{}]), 'Disease monitoring');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = `disease-monitoring_${todayStr()}.xlsx`; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  if (log === null) return <div className="p-8 text-center text-stone-400">Loading…</div>;
  const shown = (log || []).filter(r => (!from || (r.dateISO || '') >= from) && (!to || (r.dateISO || '') <= to));
  const pick = (value, set, options, tone) => (
    <div className="flex flex-wrap gap-1.5">
      {options.map(o => (
        <button key={o} onClick={() => set(value === o ? '' : o)}
          className={'px-3 py-2 rounded-lg border text-sm transition-colors ' +
            (value === o ? (tone || 'bg-stone-900 border-stone-900 text-stone-50 font-medium') : 'bg-white border-stone-300 text-stone-700 hover:border-stone-400')}>
          {o}
        </button>
      ))}
    </div>
  );

  return (
    <div className="space-y-4 pb-6">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-2 text-stone-700"><Beaker size={18} /><h2 className="text-lg font-semibold text-stone-900">Disease monitoring</h2></div>
        <button onClick={exportXlsx} className={cls.primary + ' !py-2 !px-3'}><Download size={15} /> Export</button>
      </div>
      {msg && <div className="text-sm px-3 py-2.5 rounded-lg bg-emerald-50 text-emerald-800 border border-emerald-200">{msg}</div>}

      <div className={cls.card + ' p-4 space-y-4'}>
        <div>
          <label className={cls.label}>Block</label>
          <Combobox label="" options={(config.blocks || []).map(b => b.name)} value={block} onChange={setBlock} icon={MapPin} placeholder="Search blocks…" />
        </div>

        <div><label className={cls.label}>What did you find</label>
          {pick(disease, setDisease, DISEASES, 'bg-red-600 border-red-600 text-white font-medium')}</div>

        {disease && (
          <>
            <div><label className={cls.label}>Incidence</label>{pick(incidence, setIncidence, INCIDENCE)}</div>
            <div><label className={cls.label}>Found on</label>{pick(foundOn, setFoundOn, FOUND_ON)}</div>
            <div><label className={cls.label}>Severity</label>{pick(severity, setSeverity, SEVERITY)}</div>
            <div className="flex gap-3">
              <div className="flex-1"><label className={cls.label}>Row / bay</label>
                <input value={rowRef} onChange={e => setRowRef(e.target.value)} placeholder="e.g. row 42" className={cls.input} /></div>
              <div className="flex-1"><label className={cls.label}>Pin the spot</label>
                <button onClick={dropPin} className={cls.ghost + ' !py-2.5 w-full justify-center'}>
                  <MapPin size={15} /> {pin ? 'Re-pin here' : 'Drop pin'}
                </button></div>
            </div>
            {pin && (
              <div className="text-[13px] text-stone-600 bg-stone-50 border border-stone-200 rounded-lg px-3 py-2 flex items-center gap-2 flex-wrap">
                <MapPin size={14} className="text-emerald-600" />
                {pin.lat.toFixed(5)}, {pin.lon.toFixed(5)} <span className="text-stone-400">±{pin.acc} m</span>
                <a href={`https://www.google.com/maps?q=${pin.lat},${pin.lon}`} target="_blank" rel="noreferrer" className="underline ml-auto">view map</a>
                <button onClick={() => setPin(null)} className="text-stone-400 hover:text-red-500"><X size={14} /></button>
              </div>
            )}
            {pinMsg && <p className="text-sm text-amber-700">{pinMsg}</p>}
            <div><label className={cls.label}>Note</label>
              <textarea value={note} onChange={e => setNote(e.target.value)} rows={2} className={cls.input + ' resize-y'} /></div>
          </>
        )}

        <button onClick={save} disabled={!valid} className={cls.primary + ' w-full !py-3.5 text-base'}><Check size={18} /> Record finding</button>
      </div>

      <div className="flex items-center justify-between gap-2 flex-wrap">
        <h3 className="text-base font-semibold text-stone-900">Recent findings</h3>
      </div>
      <RangePicker from={from} to={to} onChange={(f, t) => { setFrom(f); setTo(t); }} compact />
      {shown.length === 0 ? <p className="text-stone-400 text-sm text-center py-8">Nothing recorded in this range.</p> : (
        <div className="space-y-2">
          {shown.slice(0, 40).map(r => (
            <div key={r.id} className="bg-white border border-stone-200 rounded-lg px-4 py-3">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-[16px] font-bold text-stone-900">{r.block}</span>
                {r.rowRef && <span className="text-[14px] text-stone-600">· {r.rowRef}</span>}
                <span className="text-[12px] font-semibold uppercase tracking-wide text-red-700 bg-red-50 border border-red-200 rounded-full px-2 py-0.5">{r.disease}</span>
                {r.pin && <a href={`https://www.google.com/maps?q=${r.pin.lat},${r.pin.lon}`} target="_blank" rel="noreferrer" className="text-[12px] text-stone-500 underline inline-flex items-center gap-1"><MapPin size={12} /> map</a>}
                <button onClick={() => removeFinding(r.id)} title="Delete this finding"
                  className="ml-auto p-1.5 rounded-md hover:bg-red-50 text-red-400 shrink-0"><Trash2 size={15} /></button>
              </div>
              <div className="text-[13px] text-stone-600 mt-1">
                {r.incidence} · on {r.foundOn} · severity {r.severity}
              </div>
              <div className="text-[12px] text-stone-400 mt-1.5 flex items-center gap-2 flex-wrap">
                <input type="date" value={r.dateISO || ''} onChange={e => amendDate(r.id, e.target.value)}
                  title="Change the date"
                  className="px-1.5 py-1 rounded border border-stone-200 text-[12px] text-stone-600 focus:outline-none focus:ring-2 focus:ring-stone-400/40" />
                <span>{r.time} · {r.by}{r.note ? ` · ${r.note}` : ''}</span>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function TechApp({ config, onLogout, session }) {
  const [tab, setTab] = useState('dashboard');
  const tabs = [
    { id: 'dashboard', label: 'Where the team is', icon: LayoutDashboard },
    { id: 'el', label: 'E-L stages', icon: Layers },
    { id: 'disease', label: 'Disease', icon: Beaker },
  ];
  return (
    <div className="min-h-screen" style={{ backgroundColor: CREAM }}>
      <TopBar siteName={config.siteName} subtitle={`${session.name} · Technical Viticulturist`} onLogout={onLogout} wide />
      <nav className="sticky top-16 z-10 border-b border-stone-300" style={{ backgroundColor: CREAM }}>
        <div className="max-w-[1800px] mx-auto px-4 sm:px-6 flex gap-1 overflow-x-auto">
          {tabs.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)}
              className={'inline-flex items-center gap-2 px-3.5 py-3 text-sm font-medium border-b-2 -mb-px whitespace-nowrap transition-colors ' +
                (tab === t.id ? 'border-stone-900 text-stone-900' : 'border-transparent text-stone-500 hover:text-stone-800')}>
              <t.icon size={16} /> {t.label}
            </button>
          ))}
        </div>
      </nav>
      <main className="max-w-[1800px] mx-auto px-4 sm:px-6 py-6 overflow-x-hidden">
        {tab === 'dashboard' && (
          <div className="space-y-5">
            <div className="flex items-center gap-2 text-stone-700"><LayoutDashboard size={18} /><h2 className="text-lg font-semibold text-stone-900">Where the team is working</h2></div>
            <GanttPanel config={config} />
          </div>
        )}
        {tab === 'el' && <ELStages config={config} session={session} />}
        {tab === 'disease' && <DiseaseMonitor config={config} session={session} />}
      </main>
    </div>
  );
}

function ManagerApp({ config, setConfig, onLogout }) {
  const [tab, setTab] = useState('dashboard');
  const tabs = [
    { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { id: 'spray', label: 'Spray', icon: Droplets },
    { id: 'work', label: 'Work', icon: Layers },
    { id: 'timesheets', label: 'Timesheets', icon: Clock },
    { id: 'maint', label: 'Maintenance', icon: Wrench },
    { id: 'fleet', label: 'Fleet', icon: Truck },
    { id: 'hazards', label: 'Hazards', icon: AlertTriangle },
    { id: 'shed', label: 'Shed', icon: Beaker },
    { id: 'setup', label: 'Setup', icon: Settings },
  ];
  const saveConfig = async c => { setConfig(c); await saveJSON(K.config, c); };

  return (
    <div className="min-h-screen" style={{ backgroundColor: CREAM }}>
      <TopBar siteName={config.siteName} subtitle={CLOUD_ON ? "Manager console · synced" : "Manager console · this device only"} onLogout={onLogout} wide />
      <nav className="sticky top-16 z-10 border-b border-stone-300" style={{ backgroundColor: CREAM }}>
        <div className="max-w-[1800px] mx-auto px-4 sm:px-6 flex gap-1 overflow-x-auto">
          {tabs.map(t => (
            <button key={t.id} onClick={() => setTab(t.id)}
              className={'inline-flex items-center gap-2 px-3.5 py-3 text-sm font-medium border-b-2 -mb-px whitespace-nowrap transition-colors ' +
                (tab === t.id ? 'border-stone-900 text-stone-900' : 'border-transparent text-stone-500 hover:text-stone-800')}>
              <t.icon size={16} /> {t.label}
            </button>
          ))}
        </div>
      </nav>
      <main className="max-w-[1800px] mx-auto px-4 sm:px-6 py-6 overflow-x-hidden">
        {tab === 'dashboard' && <Dashboard config={config} setConfig={saveConfig} onNavigate={setTab} />}
        {tab === 'spray' && <SprayHub config={config} setConfig={saveConfig} manager={true} />}
        {tab === 'work' && <WorkManager config={config} />}
        {tab === 'timesheets' && <TimesheetDashboard config={config} />}
        {tab === 'maint' && <MaintenanceManager config={config} />}
        {tab === 'fleet' && <FleetManager config={config} setConfig={saveConfig} />}
        {tab === 'hazards' && <HazardLog config={config} />}
        {tab === 'shed' && <ChemicalShed config={config} setConfig={saveConfig} />}
        {tab === 'setup' && <Setup config={config} onSave={saveConfig} />}
      </main>
    </div>
  );
}

/* ============================================================
   Root
   ============================================================ */
function OfflineBanner() {
  const [offline, setOffline] = useState(typeof navigator !== 'undefined' && !navigator.onLine);
  useEffect(() => {
    const on = () => setOffline(false), off = () => setOffline(true);
    window.addEventListener('online', on); window.addEventListener('offline', off);
    return () => { window.removeEventListener('online', on); window.removeEventListener('offline', off); };
  }, []);
  if (!offline) return null;
  return (
    <div className="fixed bottom-0 inset-x-0 z-[70] bg-stone-900 text-stone-100 text-sm py-2 px-4 flex items-center justify-center gap-2">
      <WifiOff size={15} /> Offline — your changes are saved and will sync when you're back online.
    </div>
  );
}

export default function App() {
  const [booting, setBooting] = useState(true);
  const [config, setConfig] = useState(null);
  const [session, setSession] = useState(null);

  useEffect(() => {
    (async () => {
      await ensureReady();
      let cfg = await loadJSON(K.config, null);
      if (!cfg) { cfg = DEFAULT_CONFIG; await saveJSON(K.config, cfg); }
      cfg = { ...DEFAULT_CONFIG, ...cfg };
      // migrate jobs from plain strings to { name, code } objects
      cfg.jobs = (cfg.jobs || []).map(j => (typeof j === 'string' ? { name: j, code: '' } : { name: j.name || '', code: j.code || '' }));
      // migrate blocks from plain strings to { name, ha } objects
      cfg.blocks = (cfg.blocks || []).map(b => (typeof b === 'string' ? { name: b, ha: 0 } : { ...b, name: b.name || '', ha: Number(b.ha) || 0 }));
      // ensure newer fields exist
      if (!cfg.laneTanks) cfg.laneTanks = { ...DEFAULT_CONFIG.laneTanks };
      if (!Array.isArray(cfg.products)) cfg.products = [];
      if (!Array.isArray(cfg.roundMix)) cfg.roundMix = [];
      if (cfg.waterRate == null) cfg.waterRate = DEFAULT_CONFIG.waterRate;
      if (!Array.isArray(cfg.sprayTypes) || !cfg.sprayTypes.length) cfg.sprayTypes = JSON.parse(JSON.stringify(DEFAULT_CONFIG.sprayTypes));
      if (!cfg.weather) cfg.weather = { ...DEFAULT_CONFIG.weather };
      if (!Array.isArray(cfg.vehicles) || !cfg.vehicles.length) cfg.vehicles = JSON.parse(JSON.stringify(DEFAULT_CONFIG.vehicles));
      cfg.taskMachines = { ...DEFAULT_CONFIG.taskMachines, ...(cfg.taskMachines || {}) };
      cfg.vehicles = cfg.vehicles.map(v => ({ rego: '', wof: '', checkEveryDays: v.kind === 'vehicle' ? 7 : 14, hoursSource: 'manual', startHours: 0, ...v })).map(v => v.machineType ? v : { ...v, machineType: (DEFAULT_CONFIG.vehicles.find(d => d.name === v.name) || {}).machineType || (v.kind === 'vehicle' ? 'Vehicle' : 'Equipment') });
      if (!Array.isArray(cfg.checklist) || !cfg.checklist.length) cfg.checklist = [...DEFAULT_CONFIG.checklist];
      if (!Array.isArray(cfg.fuelTanks) || !cfg.fuelTanks.length) cfg.fuelTanks = [...DEFAULT_CONFIG.fuelTanks];
      cfg.products = (cfg.products || []).map(p => ({
        category: '', rateBasis: 'per100',
        actives: p.actives !== undefined ? p.actives : (p.concentration || ''),
        ...p, approved: p.approved === undefined ? true : p.approved,
      }));
      cfg.blocks = (cfg.blocks || []).map(b => (b.cert === undefined && BLOCK_CERT[b.name] ? { ...b, cert: BLOCK_CERT[b.name] } : b));
      if (cfg.vineSpacing == null) cfg.vineSpacing = DEFAULT_CONFIG.vineSpacing;
      if (cfg.rowWidth == null) cfg.rowWidth = DEFAULT_CONFIG.rowWidth;
      cfg.workPace = { ...DEFAULT_CONFIG.workPace, ...(cfg.workPace || {}) };
      cfg.blockCodes = { ...DEFAULT_CONFIG.blockCodes, ...(cfg.blockCodes || {}) };
      cfg.jobAccounts = { ...DEFAULT_CONFIG.jobAccounts, ...(cfg.jobAccounts || {}) };
      if (!Array.isArray(cfg.workTasks) || !cfg.workTasks.length) cfg.workTasks = [...DEFAULT_CONFIG.workTasks];
      if (!Array.isArray(cfg.machineryTasks) || !cfg.machineryTasks.length) cfg.machineryTasks = JSON.parse(JSON.stringify(DEFAULT_CONFIG.machineryTasks));
      // move any old machinery flag onto the per-operator task selection
      const machineryNames = (cfg.machineryTasks || DEFAULT_CONFIG.machineryTasks || []).map(t => t.name);
      cfg.operators = (cfg.operators || []).map(o => {
        if (Array.isArray(o.tasks) && o.tasks.length) return o;
        const wasMachinery = o.taskSet === 'machinery' || (!cfg.taskSetsAssigned && /\b(jason|simon)\b/i.test(o.name || ''));
        return wasMachinery ? { ...o, tasks: [...machineryNames], taskSet: undefined } : o;
      });
      cfg.taskSetsAssigned = true;
      // ground sprayer sizes -> 300 / 2000 (only if still on the old Jason/Simon defaults)
      cfg.sprayTypes = (cfg.sprayTypes || []).map(t => {
        if (t.key === 'ground' && JSON.stringify(t.statuses) === JSON.stringify(['To Spray', 'Jason', 'Simon'])) {
          const g = DEFAULT_CONFIG.sprayTypes.find(x => x.key === 'ground');
          return { ...t, statuses: [...g.statuses], laneTanks: { ...g.laneTanks } };
        }
        return t;
      });
      // version bump: pull the official block + task lists from the file into any older install
      if ((cfg.dataVersion || 0) < DEFAULT_CONFIG.dataVersion) {
        // MERGE ONLY — never replace or delete what's already there.
        // Add any official blocks/tasks that are missing; keep every existing
        // entry (including ones edited or added by hand) exactly as it is.
        const haveBlocks = new Set((cfg.blocks || []).map(b => b.name));
        DEFAULT_CONFIG.blocks.forEach(b => { if (!haveBlocks.has(b.name)) cfg.blocks.push({ ...b }); });
        const haveJobs = new Set((cfg.jobs || []).map(j => j.name));
        DEFAULT_CONFIG.jobs.forEach(j => { if (!haveJobs.has(j.name)) cfg.jobs.push({ ...j }); });
        const haveTasks = new Set(cfg.workTasks || []);
        DEFAULT_CONFIG.workTasks.forEach(t => { if (!haveTasks.has(t)) cfg.workTasks.push(t); });
        const have = new Set((cfg.products || []).map(p => p.name));
        DEFAULT_CONFIG.products.forEach(p => { if (!have.has(p.name)) cfg.products.push({ ...p }); });
        cfg.dataVersion = DEFAULT_CONFIG.dataVersion;
        await saveJSON(K.config, cfg);
        // spray boards and work boards are left alone — they hold real records
      }
      // fresh install: seed the canopy board from the job sheet
      if (await loadJSON(K.sprays('canopy'), null) === null) await saveJSON(K.sprays('canopy'), SHEET_CANOPY_CARDS);
      setConfig(cfg); setBooting(false);
    })();
  }, []);

  const handleCode = code => {
    if (!code) return false;
    if (code === config.managerCode) { setSession({ role: 'manager' }); return true; }
    if (config.techCode && code === config.techCode) { setSession({ role: 'tech', name: config.techName || 'Technical Viticulturist', code }); return true; }
    const op = config.operators.find(o => o.code === code);
    if (op) { setSession({ role: 'operator', code: op.code, name: op.name }); return true; }
    return false;
  };

  if (booting) return (
    <div className="min-h-screen flex items-center justify-center" style={{ backgroundColor: CREAM }}>
      <RefreshCw className="animate-spin text-stone-400" size={28} />
    </div>
  );

  let screen;
  if (!session) screen = <AuthScreen config={config} onSubmit={handleCode} />;
  else if (session.role === 'manager') screen = <ManagerApp config={config} setConfig={setConfig} onLogout={() => setSession(null)} />;
  else if (session.role === 'tech') screen = <TechApp config={config} session={session} onLogout={() => setSession(null)} />;
  else screen = <OperatorApp config={config} session={session} onLogout={() => setSession(null)} />;
  return (<>{screen}<OfflineBanner /></>);
}

