// Configuración de Firebase — completar con los valores reales del proyecto
// (Firebase console → Configuración del proyecto → tus apps → SDK setup and configuration)
import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";
import { getAuth } from "firebase/auth";

const firebaseConfig = {
  apiKey: "AIzaSyCRAYSJcyJkj-TMVXYcAjrb-P5ai0fMUYs",
  authDomain: "despertares-transporte.firebaseapp.com",
  projectId: "despertares-transporte",
  storageBucket: "despertares-transporte.firebasestorage.app",
  messagingSenderId: "658131316904",
  appId: "1:658131316904:web:1e7d709d4701e3f5f4529d",
  measurementId: "G-HKCE3N4RZ1",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);
