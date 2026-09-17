// Configuration Firebase — projet partagé avec les autres apps HT Maintenance.
// Nécessite : npm install firebase

import { initializeApp } from "firebase/app";
import { initializeFirestore } from "firebase/firestore";
import { getAuth, setPersistence, browserLocalPersistence, inMemoryPersistence } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyDDA5cCPZO2Wjfx-8YP4WFJQIUVIc-Qqb0",
  authDomain: "ht-maintenance.firebaseapp.com",
  projectId: "ht-maintenance",
  storageBucket: "ht-maintenance.firebasestorage.app",
  messagingSenderId: "273138950416",
  appId: "1:273138950416:web:eb32be0db419dbbfe8c595",
};

const app = initializeApp(firebaseConfig);

// experimentalForceLongPolling : force Firestore à utiliser de simples
// requêtes HTTP en polling, jamais le canal de streaming temps réel —
// nécessaire sur les anciennes versions de Safari/macOS (ex. Big Sur).
export const db = initializeFirestore(app, {
  experimentalForceLongPolling: true,
  useFetchStreams: false,
});

// Firebase Authentication utilise IndexedDB par défaut pour retenir la
// connexion — mal supporté sur les anciennes versions de Safari/macOS, ce
// qui peut faire "bloquer" silencieusement la vérification de connexion
// (écran vide, aucune erreur). On force explicitement localStorage, plus
// simple et plus largement compatible ; en dernier recours, en mémoire
// (fonctionne quand même, juste sans rester connecté après fermeture).
export const auth = getAuth(app);
setPersistence(auth, browserLocalPersistence).catch(() => {
  setPersistence(auth, inMemoryPersistence).catch(() => {});
});
