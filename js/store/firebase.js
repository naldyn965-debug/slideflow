// Firebase adapter. SDK is loaded from Google's CDN (pinned) only when Firebase is configured.
// Data layout:  Firestore  users/{uid}/templates/{id}, users/{uid}/projects/{id}
//               Storage    users/{uid}/templates/{id}.pptx
import { firebaseConfig, FIREBASE_SDK_VERSION, PPTX_MIME } from '../config.js';
import { SCHEMA_VERSION } from './index.js';

export async function createStore() {
  const base = `https://www.gstatic.com/firebasejs/${FIREBASE_SDK_VERSION}`;
  const [{ initializeApp }, A, F, S] = await Promise.all(['app', 'auth', 'firestore', 'storage'].map((m) => import(`${base}/firebase-${m}.js`)));
  const app = initializeApp(firebaseConfig);
  const auth = A.getAuth(app);
  const fs = F.getFirestore(app);
  const storage = S.getStorage(app);

  let user = null;
  const subs = new Set();
  const first = new Promise((resolve) => A.onAuthStateChanged(auth, (u) => {
    user = u && { id: u.uid, name: u.displayName || (u.isAnonymous ? 'Guest' : u.email || 'Account'), email: u.email, anonymous: u.isAnonymous };
    subs.forEach((cb) => cb(user));
    resolve();
  }));

  const col = (name) => {
    if (!user) throw new Error('You are signed out. Sign in again.');
    return F.collection(fs, 'users', user.id, name);
  };
  const rows = (snap) => snap.docs.map((d) => ({ id: d.id, ...d.data() }));
  const one = async (name, id) => { const s = await F.getDoc(F.doc(col(name), id)); return s.exists() ? { id, ...s.data() } : null; };
  const recent = (name) => F.getDocs(F.query(col(name), F.orderBy('updatedAt', 'desc'), F.limit(200))).then(rows);

  return {
    mode: 'firebase',
    init: () => first,
    session: {
      get user() { return user; },
      onChange(cb) { subs.add(cb); cb(user); return () => subs.delete(cb); },
      signIn: (kind) => (kind === 'guest' ? A.signInAnonymously(auth) : A.signInWithPopup(auth, new A.GoogleAuthProvider())),
      signOut: () => A.signOut(auth),
    },

    projects: {
      list: () => recent('projects'),
      get: (id) => one('projects', id),
      async create({ name, template, mode = 'preserve' }) {
        const id = crypto.randomUUID();
        const now = Date.now();
        const p = { name, templateId: template?.id ?? null, templateName: template?.name ?? null, mode, schemaVersion: SCHEMA_VERSION, createdAt: now, updatedAt: now };
        await F.setDoc(F.doc(col('projects'), id), p);
        return { id, ...p };
      },
      remove: (id) => F.deleteDoc(F.doc(col('projects'), id)),
    },

    templates: {
      list: () => recent('templates'),
      get: (id) => one('templates', id),
      async add(file, { name, summary, design }, onProgress) {
        const id = crypto.randomUUID();
        const path = `users/${user.id}/templates/${id}.pptx`;
        const ref = S.ref(storage, path);
        await new Promise((resolve, reject) => {
          const task = S.uploadBytesResumable(ref, file, { contentType: PPTX_MIME });
          task.on('state_changed', (s) => onProgress?.(s.bytesTransferred / s.totalBytes), reject, resolve);
        });
        const now = Date.now();
        const doc = { name, fileName: file.name, size: file.size, storagePath: path, ...summary, ...(design ? { design, designVersion: design.version } : {}), mode: 'preserve', schemaVersion: SCHEMA_VERSION, createdAt: now, updatedAt: now };
        try { await F.setDoc(F.doc(col('templates'), id), doc); }
        catch (e) { await S.deleteObject(ref).catch(() => {}); throw e; } // no orphaned files
        return { id, ...doc };
      },
      async update(t, patch) { // design / designVersion / mode only (enforced by firestore.rules)
        await F.updateDoc(F.doc(col('templates'), t.id), patch);
        return { ...t, ...patch };
      },
      file: (t) => S.getBlob(S.ref(storage, t.storagePath)), // needs bucket CORS, see README
      async remove(t) {
        // File first: if this fails the record stays, so nothing is lost silently.
        await S.deleteObject(S.ref(storage, t.storagePath)).catch((e) => { if (e.code !== 'storage/object-not-found') throw e; });
        await F.deleteDoc(F.doc(col('templates'), t.id));
      },
    },
  };
}
