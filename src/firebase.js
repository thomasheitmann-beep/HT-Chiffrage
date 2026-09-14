// Configuration Firebase — projet partagé avec les autres apps HT Maintenance.
// Nécessite : npm install firebase

import { initializeApp } from "firebase/app";
import { initializeFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDDA5cCPZO2Wjfx-8YP4WFJQIUVIc-Qqb0",
  authDomain: "ht-maintenance.firebaseapp.com",
  projectId: "ht-maintenance",
  storageBucket: "ht-maintenance.firebasestorage.app",
  messagingSenderId: "273138950416",
  appId: "1:273138950416:web:eb32be0db419dbbfe8c595",
};

const app = initializeApp(firebaseConfig);

// experimentalAutoDetectLongPolling : bascule automatiquement en long-polling
// quand le canal temps réel habituel (WebChannel) est bloqué — corrige les
// erreurs "Fetch API cannot load .../Listen/channel... due to access control
// checks" observées sur Safari avec certains bloqueurs/réseaux.
export const db = initializeFirestore(app, {
  experimentalAutoDetectLongPolling: true,
  useFetchStreams: false,
});
