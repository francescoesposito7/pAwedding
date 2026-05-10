// ── FIREBASE CONFIG ──
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getFirestore, collection, addDoc, getDocs, serverTimestamp }
  from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged }
  from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";

const firebaseConfig = {
  apiKey:            "AIzaSyA2QLpNhaQQBbN82J8SOIs2i1DCv-FmUQ0",
  authDomain:        "pawedding-a73c2.firebaseapp.com",
  projectId:         "pawedding-a73c2",
  storageBucket:     "pawedding-a73c2.firebasestorage.app",
  messagingSenderId: "620641371010",
  appId:             "1:620641371010:web:2e986eac67835a15d2c59b"
};

const app  = initializeApp(firebaseConfig);
const db   = getFirestore(app);
const auth = getAuth(app);

export { db, auth, collection, addDoc, getDocs, serverTimestamp,
         GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged };
