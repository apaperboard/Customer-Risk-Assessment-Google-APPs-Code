import { initializeApp, type FirebaseApp } from 'firebase/app'
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged, type User } from 'firebase/auth'
import { getFirestore, doc, getDoc, setDoc, serverTimestamp } from 'firebase/firestore'

let app: FirebaseApp | null = null

export type FirebaseConfig = {
  apiKey: string
  authDomain: string
  projectId: string
  storageBucket?: string
  messagingSenderId?: string
  appId?: string
}

export function initFirebase(): boolean {
  if (app) return true
  const globalCfg = (window as any).__FIREBASE_CONFIG as FirebaseConfig | undefined
  if (!globalCfg || !globalCfg.apiKey || !globalCfg.projectId) {
    // Not configured; caller can fall back to local storage
    return false
  }
  app = initializeApp(globalCfg)
  return true
}

export function isFirebaseReady(): boolean {
  return !!app
}

export function onUser(callback: (u: User | null) => void): () => void {
  if (!app) throw new Error('Firebase not initialized')
  const auth = getAuth(app)
  return onAuthStateChanged(auth, callback)
}

export async function signInWithGoogle(): Promise<User> {
  if (!app) throw new Error('Firebase not initialized')
  const auth = getAuth(app)
  const provider = new GoogleAuthProvider()
  const res = await signInWithPopup(auth, provider)
  return res.user
}

export async function signOutUser(): Promise<void> {
  if (!app) throw new Error('Firebase not initialized')
  const auth = getAuth(app)
  await signOut(auth)
}

export async function saveLatestReport(customerKey: string, report: any): Promise<void> {
  if (!app) throw new Error('Firebase not initialized')
  const db = getFirestore(app)
  const key = customerKey.toUpperCase().trim()
  await setDoc(doc(db, 'latestReports', key), { report, updatedAt: serverTimestamp() }, { merge: true })
}

export async function loadLatestReport<T = any>(customerKey: string): Promise<T | null> {
  if (!app) throw new Error('Firebase not initialized')
  const db = getFirestore(app)
  const key = customerKey.toUpperCase().trim()
  const snap = await getDoc(doc(db, 'latestReports', key))
  if (!snap.exists()) return null
  const data = snap.data() as any
  return (data && data.report) ?? null
}

