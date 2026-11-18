// Storage contract used by every view (both adapters implement it):
//   mode, init(), session { user, onChange(cb), signIn('google'|'guest'), signOut() }
//   projects { list(), get(id), create({name, template}), remove(id) }
//   templates { list(), get(id), add(file, {name, summary, design?}, onProgress), update(template, patch), file(template), remove(template) }
import { useFirebase } from '../config.js';

export async function createStore() {
  const mod = useFirebase ? await import('./firebase.js') : await import('./local.js');
  const store = await mod.createStore();
  await store.init();
  return store;
}

export const SCHEMA_VERSION = 1;
