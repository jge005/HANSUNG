import {
  initializeFirestore,
  getFirestore,
  doc,
  setDoc,
  getDoc,
  enableNetwork,
} from "https://www.gstatic.com/firebasejs/12.11.0/firebase-firestore.js";

const fsFbApp = window.firebaseApp;
let db = null;

if (fsFbApp) {
  try {
    db = initializeFirestore(fsFbApp, {
      experimentalForceLongPolling: true,
      useFetchStreams: false,
    });
  } catch (err) {
    console.warn("Firestore long-polling 초기화 실패, 기본 설정으로 재시도:", err);
    db = getFirestore(fsFbApp);
  }
}

if (db) {
  enableNetwork(db).catch(function (err) {
    console.warn("Firestore 네트워크 활성화 실패:", err);
  });
}

function isRetryableFirestoreError(err) {
  var code = err && err.code ? String(err.code) : "";
  var message = err && err.message ? String(err.message).toLowerCase() : "";
  return (
    code.indexOf("unavailable") >= 0 ||
    code.indexOf("failed-precondition") >= 0 ||
    message.indexOf("offline") >= 0 ||
    message.indexOf("network") >= 0
  );
}

async function withNetworkRetry(work) {
  if (!db) throw new Error("Firestore 초기화 실패: window.firebaseApp 없음");
  try {
    return await work();
  } catch (err) {
    if (!isRetryableFirestoreError(err)) throw err;
    await enableNetwork(db).catch(function () {});
    return await work();
  }
}

async function saveWorkDraft(draftId, payload) {
  await withNetworkRetry(function () {
    return setDoc(doc(db, "workDrafts", draftId), payload, { merge: true });
  });
}

async function loadWorkDraft(draftId) {
  return withNetworkRetry(async function () {
    const snap = await getDoc(doc(db, "workDrafts", draftId));
    return snap.exists() ? snap.data() : null;
  });
}

window.firebaseFirestoreApi = {
  saveWorkDraft: saveWorkDraft,
  loadWorkDraft: loadWorkDraft,
  isReady: function () { return !!db; },
  isRetryableFirestoreError: isRetryableFirestoreError,
};
