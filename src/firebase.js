// Configuration Firebase — projet partagé avec les autres apps HT Maintenance.
// Nécessite : npm install firebase

import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDDA5cCPZO2Wjfx-8YP4WFJQIUVIc-Qqb0",
  authDomain: "ht-maintenance.firebaseapp.com",
  projectId: "ht-maintenance",
  storageBucket: "ht-maintenance.firebasestorage.app",
  messagingSenderId: "273138950416",
  appId: "1:273138950416:web:eb32be0db419dbbfe8c595",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
