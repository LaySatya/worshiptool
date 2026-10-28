/**
 * Firebase Client Configuration
 * Project: churchtool
 * Project ID: immanuel-songs
 * Project Number: 354565330550
 */
import { initializeApp, getApps } from "firebase/app";
import {
  getFirestore,
  collection,
  doc,
  getDocs,
  getDoc,
  setDoc,
  deleteDoc,
  query,
  where,
  orderBy,
} from "firebase/firestore";

const firebaseConfig = {
  projectId: "immanuel-songs",
  appId: "1:354565330550:web:churchtool",
  authDomain: "immanuel-songs.firebaseapp.com",
  storageBucket: "immanuel-songs.appspot.com",
  messagingSenderId: "354565330550",
  // Fallback placeholder apiKey if not provided by env/storage
  apiKey:
    typeof window !== "undefined" && window.localStorage.getItem("firebase_api_key")
      ? window.localStorage.getItem("firebase_api_key")
      : "AIzaSyDummyKeyForImmanuelSongs354565330550",
};

let app = null;
let db = null;
let isInitialized = false;

try {
  app = getApps().length ? getApps()[0] : initializeApp(firebaseConfig);
  db = getFirestore(app);
  isInitialized = true;
  console.log("[Firebase] Initialized for project:", firebaseConfig.projectId);
} catch (err) {
  console.warn("[Firebase] Initialization notice:", err.message);
  isInitialized = false;
}

export {
  app,
  db,
  isInitialized,
  firebaseConfig,
  collection,
  doc,
  getDocs,
  getDoc,
  setDoc,
  deleteDoc,
  query,
  where,
  orderBy,
};
