import React from 'react';
import { createRoot } from 'react-dom/client';
import { SprayHub } from './App.test.jsx';
const config = {
  siteName: 'Loveblock',
  blocks: [
    { name: 'Hill - A 23', ha: 7.9, km: 26.69, cert: 'SWNZ' },
    { name: 'Eros - TG2017', ha: 9.246, km: 30.82, cert: 'Organic' },
    { name: 'SB 04', ha: 5.328, rowWidth: 2.7 },
  ],
  products: [
    { name: 'Microthiol Disperss', unit: 'Kg', category: 'Powdery Mildew', rateBasis: 'per100', rate: 1.333, approved: true, biogro: true },
    { name: 'NZBioActive', unit: 'L', category: 'Nutrition', rateBasis: 'perHa', rate: 2.5, approved: true, biogro: true },
  ],
  operators: [{ code: '1234', name: 'Jason' }],
  sprayTypes: [{ key: 'ground', label: 'Ground Spray', statuses: ['To Spray', '300 L', '2000 L'], laneTanks: { '300 L': 300, '2000 L': 2000 }, waterRate: 400,
    roundMix: [{ product: 'Microthiol Disperss', rate: 1.333 }, { product: 'NZBioActive', rate: 2.5 }] }],
};
window.__errs = [];
window.addEventListener('error', e => window.__errs.push((e.message||'') + (e.error&&e.error.stack ? ' | ' + e.error.stack.split('\n').slice(0,6).join(' > ') : '')));
const root = createRoot(document.getElementById('root'));
function tapText(re) { const el = [...document.querySelectorAll('button')].find(x => re.test(x.textContent)); if (el) el.click(); return !!el; }
function setInput(re, val) { const el = [...document.querySelectorAll('input')].find(x => re.test(x.value) || re.test(x.placeholder||'')); if(!el) return false;
  const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,'value').set;
  setter.call(el, val); el.dispatchEvent(new Event('input',{bubbles:true})); return true; }
root.render(React.createElement(SprayHub, { config, setConfig: () => {}, manager: true }));
setTimeout(() => {
  tapText(/Ground Spray/);
  setTimeout(() => {
    tapText(/Round mix/);
    setTimeout(() => {
      console.log('mix open errs:', window.__errs.join('\n') || 'none');
      // change the water rate field
      const waterInput = [...document.querySelectorAll('input')].find(i => i.value === '400');
      console.log('found water input:', !!waterInput);
      if (waterInput) {
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,'value').set;
        setter.call(waterInput, '600'); waterInput.dispatchEvent(new Event('input',{bubbles:true}));
      }
      setTimeout(() => {
        console.log('after water change errs:', window.__errs.join('\n') || 'none');
        console.log('click Fill rates from shed:', tapText(/Fill rates/));
        setTimeout(() => {
          console.log('FINAL errs:', window.__errs.join('\n') || 'NONE');
        }, 400);
      }, 300);
    }, 400);
  }, 500);
}, 400);
