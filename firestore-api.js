import {
  initializeFirestore,
  getFirestore,
  doc,
  setDoc,
  getDoc,
  collection,
  getDocs,
  deleteDoc,
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

const DRAFT_CHUNK_COLLECTION = "chunks";
const DRAFT_MAX_CHUNK_CHARS = 80000;

function splitTextChunks(text, chunkChars) {
  var source = String(text || "");
  var size = Math.max(1000, Number(chunkChars) || DRAFT_MAX_CHUNK_CHARS);
  var chunks = [];
  for (var i = 0; i < source.length; i += size) {
    chunks.push(source.slice(i, i + size));
  }
  return chunks.length ? chunks : [""];
}

async function saveChunkedDraft(draftId, payload) {
  var json = JSON.stringify(payload || {});
  var chunks = splitTextChunks(json, DRAFT_MAX_CHUNK_CHARS);
  var draftRef = doc(db, "workDrafts", draftId);
  var chunksColRef = collection(draftRef, DRAFT_CHUNK_COLLECTION);

  await setDoc(
    draftRef,
    {
      __chunked: true,
      chunkCount: chunks.length,
      updatedAt: payload && payload.updatedAt ? payload.updatedAt : new Date().toISOString(),
      version: 2,
    }
  );

  await Promise.all(
    chunks.map(function (chunk, index) {
      var id = String(index).padStart(4, "0");
      return setDoc(doc(chunksColRef, id), { i: index, data: chunk }, { merge: false });
    })
  );

  var snap = await getDocs(chunksColRef);
  var deletes = [];
  snap.forEach(function (chunkDoc) {
    var data = chunkDoc.data() || {};
    var i = typeof data.i === "number" ? data.i : Number(chunkDoc.id);
    if (!isNaN(i) && i >= chunks.length) {
      deletes.push(deleteDoc(chunkDoc.ref));
    }
  });
  if (deletes.length) {
    await Promise.all(deletes);
  }
}

async function loadChunkedDraft(draftId, meta) {
  var draftRef = doc(db, "workDrafts", draftId);
  var chunksColRef = collection(draftRef, DRAFT_CHUNK_COLLECTION);
  var snap = await getDocs(chunksColRef);
  var pieces = [];
  snap.forEach(function (chunkDoc) {
    var data = chunkDoc.data() || {};
    pieces.push({
      i: typeof data.i === "number" ? data.i : Number(chunkDoc.id),
      data: String(data.data || ""),
    });
  });
  if (!pieces.length) return null;
  pieces.sort(function (a, b) { return a.i - b.i; });
  var count = meta && typeof meta.chunkCount === "number" ? meta.chunkCount : pieces.length;
  var json = pieces.slice(0, count).map(function (item) { return item.data; }).join("");
  if (!json) return null;
  return JSON.parse(json);
}

async function saveWorkDraft(draftId, payload) {
  await withNetworkRetry(function () {
    return saveChunkedDraft(draftId, payload);
  });
}

async function loadWorkDraft(draftId) {
  return withNetworkRetry(async function () {
    const ref = doc(db, "workDrafts", draftId);
    const snap = await getDoc(ref);
    if (!snap.exists()) return null;
    var data = snap.data() || {};
    if (data.__chunked) {
      try {
        return await loadChunkedDraft(draftId, data);
      } catch (err) {
        console.warn("분할 문서 로드 실패, 단일 문서 데이터로 재시도:", err);
      }
    }
    return data;
  });
}

window.firebaseFirestoreApi = {
  saveWorkDraft: saveWorkDraft,
  loadWorkDraft: loadWorkDraft,
  isReady: function () { return !!db; },
  isRetryableFirestoreError: isRetryableFirestoreError,
};
