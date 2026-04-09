    import {
      getFirestore,
      doc,
      setDoc,
      getDoc,
    } from "https://www.gstatic.com/firebasejs/12.11.0/firebase-firestore.js";

    const fsFbApp = window.firebaseApp;
    const db = fsFbApp ? getFirestore(fsFbApp) : null;

    async function saveWorkDraft(draftId, payload) {
      if (!db) throw new Error("Firestore 초기화 실패: window.firebaseApp 없음");
      await setDoc(doc(db, "workDrafts", draftId), payload, { merge: true });
    }

    async function loadWorkDraft(draftId) {
      if (!db) throw new Error("Firestore 초기화 실패: window.firebaseApp 없음");
      const snap = await getDoc(doc(db, "workDrafts", draftId));
      return snap.exists() ? snap.data() : null;
    }

    // 메인(일반) 스크립트에서 호출 가능하도록 노출
    window.firebaseFirestoreApi = {
      saveWorkDraft: saveWorkDraft,
      loadWorkDraft: loadWorkDraft,
    };
