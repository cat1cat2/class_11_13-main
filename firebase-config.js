// Paste the Firebase Web App config shown in Firebase Console here.
// This browser config is public; protect user data with Firestore Security Rules.
export const firebaseConfig = {
  apiKey: "AIzaSyBtIyXYBY5pIkErC0uwCWLo6shyosQyprc",
  authDomain: "web-app-9372c.firebaseapp.com",
  projectId: "web-app-9372c",
  storageBucket: "web-app-9372c.firebasestorage.app",
  messagingSenderId: "323358441627",
  appId: "1:323358441627:web:f6e80706d2d961b5f1463e",
  measurementId: "G-F1SYMFS9Y0"
};

// Fill this with the reCAPTCHA Enterprise site key for the registered Web App after
// enabling Firebase App Check. AI requests stay disabled until it is set.
export const appCheckSiteKey = '';

export const firebaseConfigured = !Object.values(firebaseConfig).some(value =>
  String(value).startsWith("PASTE_") || String(value).includes("YOUR_PROJECT_ID")
);
