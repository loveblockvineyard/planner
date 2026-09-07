import React from 'react';
import { createRoot } from 'react-dom/client';
import { SprayHub } from './App.test.jsx';
const config = {
  siteName: 'Loveblock',
  blocks: [{ name: 'Hill - E PG', ha: 1.55, rows: '18-57', km: 5.17, cert: 'SWNZ' }],
  products: [
    { name: 'Microthiol Disperss', unit: 'Kg', category: 'Powdery Mildew', rateBasis: 'per100', rate: 1.333, approved: true, biogro: true },
  ],
  operators: [{ code: '1234', name: 'Jason' }],
  sprayTypes: [{ key: 'ground', label: 'Ground Spray', statuses: ['To Spray', '300 L', '2000 L'], laneTanks: { '300 L': 300, '2000 L': 2000 }, waterRate: 400,
    roundMix: [
      { product: 'Microthiol Disperss', per100: 1.333 },     // legacy shape, no 'rate'
      { product: 'Deleted Product', per100: 2 },              // no longer in shed
      { product: '', rate: '' },                              // blank row
    ] }],
};
window.__errs = [];
window.addEventListener('error', e => window.__errs.push((e.message||'') + (e.error&&e.error.stack ? ' | ' + e.error.stack.split('\n').slice(0,5).join(' > ') : '')));
const root = createRoot(document.getElementById('root'));
function tapText(re) { const el = [...document.querySelectorAll('button')].find(x => re.test(x.textContent)); if (el) el.click(); return !!el; }
root.render(React.createElement(SprayHub, { config, setConfig: () => {}, manager: true }));
setTimeout(() => {
  tapText(/Ground Spray/);
  setTimeout(() => {
    console.log('after open, errs:', window.__errs.join('\n') || 'none');
    tapText(/Round mix/);
    setTimeout(() => {
      console.log('after round mix open, errs:', window.__errs.join('\n') || 'none');
    }, 400);
  }, 500);
}, 400);
