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
function tapText(tag, re) { const els = [...document.querySelectorAll(tag)].filter(x => re.test(x.textContent)); if (els[0]) els[0].click(); return els.length; }
root.render(React.createElement(SprayHub, { config, setConfig: () => {}, manager: true }));
setTimeout(() => {
  tapText('button', /Ground Spray/);
  setTimeout(() => {
    tapText('button', /^Add blocks$/);
    setTimeout(() => {
      const n = tapText('button', /Hill - E PG/);
      console.log('matched Hill buttons:', n);
      setTimeout(() => {
        const btns = [...document.querySelectorAll('button')].map(b=>b.textContent.trim());
        console.log('ALL BUTTONS NOW:', JSON.stringify(btns));
      }, 300);
    }, 300);
  }, 500);
}, 400);
