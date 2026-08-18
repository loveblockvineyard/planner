import React from 'react';
import { createRoot } from 'react-dom/client';
import { registerSW } from 'virtual:pwa-register';
import './index.css';
import App from './App.jsx';

/* ------------------------------------------------------------
   Instant updates.

   When a new version is deployed, apply it straight away instead
   of waiting for the person to close and reopen the app. The one
   exception: if they're mid-way through typing something, hold on
   until they stop, so a reload never eats a half-filled form.
   ------------------------------------------------------------ */
const busyTyping = () => {
  const el = document.activeElement;
  if (!el) return false;
  const tag = (el.tagName || '').toLowerCase();
  return tag === 'input' || tag === 'textarea' || tag === 'select' || el.isContentEditable;
};

const updateSW = registerSW({
  immediate: true,
  onNeedRefresh() { applyUpdate(); },
  onRegisteredSW(_url, reg) {
    if (!reg) return;
    // look for a new version regularly, and whenever the app comes back
    // to the foreground (phones left open all day in the vineyard)
    const check = () => { reg.update().catch(() => {}); };
    setInterval(check, 30000);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) check(); });
    window.addEventListener('focus', check);
    window.addEventListener('online', check);
  },
});

let applying = false;
function applyUpdate() {
  if (applying) return;
  applying = true;
  const go = () => {
    if (busyTyping()) { setTimeout(go, 1500); return; }   // wait for a quiet moment
    updateSW(true);   // activates the new version and reloads
  };
  go();
}

createRoot(document.getElementById('root')).render(<App />);
