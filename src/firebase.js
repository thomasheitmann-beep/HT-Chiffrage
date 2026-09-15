// Configuration Firebase — projet partagé avec les autres apps HT Maintenance.
// Nécessite : npm install firebase

import { initializeApp } from "firebase/app";
import { initializeFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";

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

// Les règles Firestore de ce projet exigent un utilisateur authentifié
// (request.auth != null) — on utilise une connexion anonyme automatique,
// sans écran de connexion ni mot de passe pour l'utilisateur.
export const auth = getAuth(app);
