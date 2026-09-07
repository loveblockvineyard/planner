import React from 'react';
import { createRoot } from 'react-dom/client';
import { SprayHub } from './App.test.jsx';
const config = {
  siteName: 'Loveblock',
  blocks: [{ name: 'Hill - E PG', ha: 1.55, rows: '18-57', km: 5.17, cert: 'SWNZ' }],
  products: [{ name: 'Microthiol Disperss', unit: 'Kg', category: 'Powdery Mildew', rateBasis: 'per100', rate: 1.333, approved: true, biogro: true }],
  operators: [{ code: '1234', name: 'Jason' }],
  sprayTypes: [{ key: 'ground', label: 'Ground Spray', statuses: ['To Spray', '300 L', '2000 L'], laneTanks: { '300 L': 300, '2000 L': 2000 }, waterRate: 400, roundMix: [] }],
};
window.__errs = [];
window.addEventListener('error', e => window.__errs.push((e.message||'')));
const root = createRoot(document.getElementById('root'));
function tapText(tag, re) { const el = [...document.querySelectorAll(tag)].find(x => re.test(x.textContent)); if (el) el.click(); return !!el; }
root.render(React.createElement(SprayHub, { config, setConfig: () => {}, manager: true }));
setTimeout(() => {
  tapText('button', /Ground Spray/);
  setTimeout(() => {
    tapText('button', /Add blocks/);
    setTimeout(() => {
      tapText('button', /Hill - E PG/);
      setTimeout(() => {
        const btns = [...document.querySelectorAll('button')].map(b=>b.textContent.trim()).filter(t=>/add/i.test(t));
        console.log('buttons with "add":', JSON.stringify(btns));
      }, 300);
    }, 300);
  }, 500);
}, 400);
