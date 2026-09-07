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
window.addEventListener('error', e => window.__errs.push((e.message||'') + (e.error&&e.error.stack ? ' | ' + e.error.stack.split('\n').slice(0,4).join(' > ') : '')));
const root = createRoot(document.getElementById('root'));
function tapText(re) { const el = [...document.querySelectorAll('button')].find(x => re.test(x.textContent)); if (el) el.click(); return !!el; }
root.render(React.createElement(SprayHub, { config, setConfig: () => {}, manager: true }));
setTimeout(() => {
  tapText(/Ground Spray/);
  setTimeout(() => {
    tapText(/Add blocks/);
    setTimeout(() => {
      console.log('after Add blocks click, errs:', window.__errs.join('\n') || 'none');
      const btns = [...document.querySelectorAll('button')].map(b=>b.textContent.trim());
      console.log('buttons:', JSON.stringify(btns));
      const clicked = tapText(/Hill - E PG/);
      console.log('clicked Hill button:', clicked);
      setTimeout(() => {
        console.log('after Hill click, errs:', window.__errs.join('\n') || 'none');
        const btns2 = [...document.querySelectorAll('button')].map(b=>b.textContent.trim());
        console.log('buttons2:', JSON.stringify(btns2));
        const clicked2 = tapText(/Add \d+ block/);
        console.log('clicked Add N block:', clicked2);
        setTimeout(() => {
          console.log('FINAL errs:', window.__errs.join('\n') || 'NONE');
        }, 400);
      }, 400);
    }, 400);
  }, 500);
}, 400);
