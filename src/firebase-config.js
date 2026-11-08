import { initializeApp } from "firebase/app";
import { getAuth, GoogleAuthProvider } from "firebase/auth";
import { getFirestore } from "firebase/firestore";
import { getStorage } from "firebase/storage";

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY || "AIzaSyA7U6RLAlPj22bdCKpSBycQ-8LuNiibHwc",
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN || "school-roster-3bfc1.firebaseapp.com",
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID || "school-roster-3bfc1",
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET || "school-roster-3bfc1.firebasestorage.app",
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID || "867702306001",
  appId: import.meta.env.VITE_FIREBASE_APP_ID || "1:867702306001:web:575f451bd79264cc5201f0",
  measurementId: import.meta.env.VITE_FIREBASE_MEASUREMENT_ID || "G-M6XQ7EYWCD"
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const googleProvider = new GoogleAuthProvider();
export const db = getFirestore(app);
export const storage = getStorage(app);
