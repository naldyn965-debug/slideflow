// Local mode: IndexedDB in this browser. Lets the app run before Firebase is configured. Not synced, not shared.
import { SCHEMA_VERSION } from './index.js';

const NAMES = ['projects', 'templates', 'files'];
let dbp;
function db() {
  dbp ||= new Promise((resolve, reject) => {
    if (!('indexedDB' in globalThis)) return reject(new Error('This browser does not support local storage (IndexedDB).'));
    const r = indexedDB.open('slideflow-local', 1);
    r.onupgradeneeded = () => NAMES.forEach((n) => r.result.createObjectStore(n, { keyPath: 'id' }));
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error || new Error('Could not open local storage.'));
  });
  return dbp;
}
const wrap = (req) => new Promise((res, rej) => { req.onsuccess = () => res(req.result); req.onerror = () => rej(req.error); });
async function read(name, fn) { const d = await db(); return wrap(fn(d.transaction(name).objectStore(name))); }
async function commit(names, fn) {
  const d = await db();
  return new Promise((res, rej) => {
    const t = d.transaction(names, 'readwrite');
    fn(...names.map((n) => t.objectStore(n)));
    t.oncomplete = () => res();
    t.onabort = t.onerror = () => rej(t.error?.name === 'QuotaExceededError' ? new Error('Browser storage is full. Delete a template or free up space.') : t.error || new Error('Could not write to local storage.'));
  });
}
const newest = (a, b) => b.updatedAt - a.updatedAt;

export async function createStore() {
  const user = { id: 'local', name: 'Local workspace' };
  return {
    mode: 'local',
    init: () => db(),
    session: { user, onChange: (cb) => { cb(user); return () => {}; }, signIn: async () => {}, signOut: async () => {} },

    projects: {
      list: async () => (await read('projects', (s) => s.getAll())).sort(newest),
      get: async (id) => (await read('projects', (s) => s.get(id))) || null,
      async create({ name, template }) {
        const now = Date.now();
        const p = { id: crypto.randomUUID(), name, templateId: template?.id ?? null, templateName: template?.name ?? null, schemaVersion: SCHEMA_VERSION, createdAt: now, updatedAt: now };
        await commit(['projects'], (s) => s.put(p));
        return p;
      },
      remove: (id) => commit(['projects'], (s) => s.delete(id)),
    },

    templates: {
      list: async () => (await read('templates', (s) => s.getAll())).sort(newest),
      get: async (id) => (await read('templates', (s) => s.get(id))) || null,
      async add(file, { name, summary }, onProgress) {
        const now = Date.now();
        const id = crypto.randomUUID();
        const t = { id, name, fileName: file.name, size: file.size, storagePath: `local/${id}`, ...summary, schemaVersion: SCHEMA_VERSION, createdAt: now, updatedAt: now };
        const bytes = await file.arrayBuffer();
        await commit(['templates', 'files'], (ts, fs) => { ts.put(t); fs.put({ id, bytes }); });
        onProgress?.(1);
        return t;
      },
      async file(t) {
        const rec = await read('files', (s) => s.get(t.id));
        if (!rec) throw new Error('The stored file for this template is missing.');
        return new Blob([rec.bytes]);
      },
      remove: (t) => commit(['templates', 'files'], (ts, fs) => { ts.delete(t.id); fs.delete(t.id); }),
    },
  };
}
