// Συγχρονισμός με Firebase: login με Google και αποθήκευση στο Firestore.
// Δομή: users/{uid}/tx/{id} μία κίνηση ανά έγγραφο, users/{uid}/meta/settings {bud, rules, ign}.
// Το index.html καλεί CL.push() σε κάθε save(). Εδώ βρίσκουμε τι άλλαξε και γράφουμε μόνο αυτό.
import { initializeApp } from "https://www.gstatic.com/firebasejs/12.6.0/firebase-app.js";
import { getAuth, onAuthStateChanged, GoogleAuthProvider, signInWithPopup, signInWithRedirect, signOut } from "https://www.gstatic.com/firebasejs/12.6.0/firebase-auth.js";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, collection, doc, onSnapshot, writeBatch } from "https://www.gstatic.com/firebasejs/12.6.0/firebase-firestore.js";

const app = initializeApp({
  apiKey: "AIzaSyDV0J_r3YKOenBK31t9959Pk4ffYK51mgc",
  authDomain: "tetradio-64530.firebaseapp.com",
  projectId: "tetradio-64530",
  storageBucket: "tetradio-64530.firebasestorage.app",
  messagingSenderId: "567647197037",
  appId: "1:567647197037:web:c3bec4f86130c451662c64"
});
const auth = getAuth(app);
// Τοπική cache του Firestore: η εφαρμογή δουλεύει και χωρίς σύνδεση και συγχρονίζει όταν βρει ίντερνετ.
const db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });

const CL = window.CL = { user: null, ready: false, err: "" };
let have = new Set(), meta = "", unsub = [];
const clean = (o) => JSON.parse(JSON.stringify(o));
const metaOf = (s) => JSON.stringify({ bud: s.bud || {}, rules: s.rules || {}, ign: s.ign || [], acc: s.acc || {}, cash: s.cash || null });

CL.signIn = async () => {
  CL.err = ""; const p = new GoogleAuthProvider();
  try { await signInWithPopup(auth, p) }
  catch (e) {
    if (e.code === "auth/popup-blocked" || e.code === "auth/operation-not-supported-in-this-environment") return signInWithRedirect(auth, p);
    if (e.code !== "auth/popup-closed-by-user" && e.code !== "auth/cancelled-popup-request") { CL.err = "Η σύνδεση απέτυχε: " + e.code; render() }
  }
};
CL.signOut = () => signOut(auth);

// Γράφει στο Firestore τη διαφορά ανάμεσα στο S και σε όσα έχει ήδη το cloud.
CL.push = async () => {
  if (!CL.user || !CL.ready) return;
  const base = "users/" + CL.user.uid, ops = [];
  const ids = new Set(S.tx.map((t) => t.id));
  S.tx.forEach((t) => { if (!have.has(t.id)) { have.add(t.id); ops.push((b) => b.set(doc(db, base + "/tx/" + t.id), clean(t))) } });
  [...have].forEach((id) => { if (!ids.has(id)) { have.delete(id); ops.push((b) => b.delete(doc(db, base + "/tx/" + id))) } });
  const m = metaOf(S);
  if (m !== meta) { meta = m; ops.push((b) => b.set(doc(db, base + "/meta/settings"), JSON.parse(m))) }
  cache();
  // Ένα batch δέχεται έως 500 εγγραφές.
  for (let i = 0; i < ops.length; i += 450) {
    const b = writeBatch(db); ops.slice(i, i + 450).forEach((f) => f(b));
    b.commit().catch((e) => { CL.err = "Δεν αποθηκεύτηκε στο cloud: " + e.code; render() });
  }
};

// Μεταφέρει τα δεδομένα αυτής της συσκευής στον λογαριασμό. Όσα υπάρχουν ήδη δεν διπλασιάζονται.
CL.local = () => { try { const L = JSON.parse(localStorage.getItem("tetradio1")); return L && L.tx && L.tx.length && localStorage.getItem("tetradio1-moved") !== CL.user.uid ? L : null } catch (e) { return null } };
CL.migrate = () => {
  const L = CL.local(); if (!L) return;
  const ids = new Set(S.tx.map((t) => t.id)), uids = new Set(S.tx.map((t) => t.uid).filter(Boolean));
  L.tx.forEach((t) => { if (!ids.has(t.id) && !(t.uid && uids.has(t.uid))) S.tx.push(t) });
  S.bud = Object.assign({}, L.bud, S.bud);
  S.rules = Object.assign({}, L.rules, S.rules);
  S.ign = [...new Set([...(S.ign || []), ...(L.ign || [])])];
  S.acc = Object.assign({}, L.acc, S.acc); S.cash = S.cash || L.cash || null;
  localStorage.setItem("tetradio1-moved", CL.user.uid);
  CL.push(); render();
};
CL.skip = () => { localStorage.setItem("tetradio1-moved", CL.user.uid); render() };

function cache() { try { localStorage.setItem("tetradio-cloud", JSON.stringify(S)) } catch (e) {} }

onAuthStateChanged(auth, (u) => {
  unsub.forEach((f) => f()); unsub = [];
  CL.user = u; CL.ready = false; CL.known = true;
  if (!u) {
    // Αποσύνδεση: γυρνάμε στα δεδομένα της συσκευής.
    if (localStorage.getItem("tetradio-cloud")) { localStorage.removeItem("tetradio-cloud"); location.reload() }
    render(); return;
  }
  const base = "users/" + u.uid; let gotTx = false, gotMeta = false, shown = "";
  // Ξαναζωγραφίζουμε μόνο όταν άλλαξαν τα δεδομένα, για να μη σβήνει μια φόρμα που συμπληρώνεις.
  const done = () => {
    if (!gotTx || !gotMeta) return;
    CL.ready = true; cache();
    const key = [...have].sort().join() + meta;
    if (key !== shown) { shown = key; render() }
  };
  // Μια άδεια απάντηση από την cache δεν σημαίνει άδειο λογαριασμό: περιμένουμε τον server.
  unsub.push(onSnapshot(collection(db, base + "/tx"), (q) => {
    if (q.empty && q.metadata.fromCache) return;
    have = new Set(q.docs.map((d) => d.id));
    S.tx = q.docs.map((d) => Object.assign({ id: d.id }, d.data()));
    gotTx = true; done();
  }, (e) => { CL.err = "Σφάλμα συγχρονισμού: " + e.code; render() }));
  unsub.push(onSnapshot(doc(db, base + "/meta/settings"), (d) => {
    if (!d.exists() && d.metadata.fromCache) return;
    const v = d.exists() ? d.data() : {};
    S.bud = v.bud || {}; S.rules = v.rules || {}; S.ign = v.ign || []; S.acc = v.acc || {}; S.cash = v.cash || null;
    meta = metaOf(S); gotMeta = true; done();
  }, (e) => { CL.err = "Σφάλμα συγχρονισμού: " + e.code; render() }));
});
