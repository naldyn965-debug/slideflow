// Firebase web config is NOT a secret: access is enforced by firestore.rules / storage.rules.
// Leave apiKey empty to run in Local mode (IndexedDB, this browser only).
export const firebaseConfig = {
  apiKey: '',
  authDomain: '',
  projectId: '',
  storageBucket: '',
  appId: '',
};

export const FIREBASE_SDK_VERSION = '10.14.1';
export const useFirebase = Boolean(firebaseConfig.apiKey && firebaseConfig.projectId && firebaseConfig.storageBucket);

// Keep in sync with storage.rules / firestore.rules.
export const LIMITS = {
  maxTemplateBytes: 50 * 1024 * 1024,
  maxSlides: 200,
  maxXmlBytes: 8 * 1024 * 1024,
  maxUncompressedBytes: 600 * 1024 * 1024,
  maxNameLength: 120,
};

export const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
