/* ============================================================
   Storage layer.

   If Supabase env vars are present the app uses a shared cloud
   database — every phone and desktop sees the same data.
   If they're missing it quietly falls back to this device's
   local storage, so the app still runs (e.g. opened from a file).

   The whole app only ever calls loadJSON / saveJSON.
   ============================================================ */
import { createClient } from '@supabase/supabase-js';

const URL_ = import.meta.env.VITE_SUPABASE_URL;
const KEY_ = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const isCloud = Boolean(URL_ && KEY_);
const supabase = isCloud ? createClient(URL_, KEY_) : null;

const TABLE = 'kv';
const mem = new Map();

/* ---- local fallback ---- */
function lsGet(key) {
  try { return window.localStorage.getItem(key); }
  catch { return mem.has(key) ? mem.get(key) : null; }
}
function lsSet(key, val) {
  try { window.localStorage.setItem(key, val); }
  catch { mem.set(key, val); }
}

export async function loadJSON(key, fallback) {
  if (isCloud) {
    try {
      const { data, error } = await supabase.from(TABLE).select('value').eq('key', key).maybeSingle();
      if (error) throw error;
      if (data && data.value !== null && data.value !== undefined) return data.value;
      // nothing in the cloud yet — seed from this device if it has something
      const local = lsGet(key);
      return local == null ? fallback : JSON.parse(local);
    } catch (e) {
      console.error('loadJSON (cloud) failed, using local copy:', key, e.message || e);
      try { const v = lsGet(key); return v == null ? fallback : JSON.parse(v); } catch { return fallback; }
    }
  }
  try { const v = lsGet(key); return v == null ? fallback : JSON.parse(v); }
  catch { return fallback; }
}

export async function saveJSON(key, value) {
  // always keep a local copy: instant reads and an offline safety net
  try { lsSet(key, JSON.stringify(value)); } catch { /* ignore */ }
  if (isCloud) {
    try {
      const { error } = await supabase.from(TABLE)
        .upsert({ key, value, updated_at: new Date().toISOString() }, { onConflict: 'key' });
      if (error) throw error;
      return true;
    } catch (e) {
      console.error('saveJSON (cloud) failed, kept locally:', key, e.message || e);
      return false;
    }
  }
  return true;
}

/* ============================================================
   Live updates.

   subscribe(key, onChange) calls onChange(newValue) whenever that
   key changes on ANY device — so a manager watching the board sees
   an operator's "Done" the moment they tap it.

   Uses Supabase Realtime (a websocket push, near-instant). A slow
   poll runs alongside it as a safety net for flaky rural signal,
   and is the only mechanism when running local-only.

   Returns an unsubscribe function — always call it on unmount.
   ============================================================ */
export function subscribe(key, onChange, { pollMs = 20000 } = {}) {
  let stopped = false;
  let lastSeen = null;

  const emit = value => {
    if (stopped) return;
    const stamp = JSON.stringify(value);
    if (stamp === lastSeen) return;   // nothing actually changed
    lastSeen = stamp;
    onChange(value);
  };

  let channel = null;
  if (isCloud) {
    try {
      channel = supabase
        .channel(`kv:${key}`)
        .on('postgres_changes',
          { event: '*', schema: 'public', table: TABLE, filter: `key=eq.${key}` },
          payload => {
            const v = payload.new && payload.new.value;
            if (v !== undefined && v !== null) {
              try { lsSet(key, JSON.stringify(v)); } catch { /* ignore */ }
              emit(v);
            }
          })
        .subscribe();
    } catch (e) {
      console.error('realtime subscribe failed, polling instead:', key, e.message || e);
    }
  }

  const timer = setInterval(async () => {
    if (stopped) return;
    try { const v = await loadJSON(key, null); if (v !== null) emit(v); }
    catch { /* ignore */ }
  }, pollMs);

  return () => {
    stopped = true;
    clearInterval(timer);
    if (channel) { try { supabase.removeChannel(channel); } catch { /* ignore */ } }
  };
}
