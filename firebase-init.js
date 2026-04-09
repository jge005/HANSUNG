    // Import the functions you need from the SDKs you need
    import { initializeApp } from "https://www.gstatic.com/firebasejs/12.11.0/firebase-app.js";
    import {
      getAnalytics,
      isSupported as isAnalyticsSupported,
    } from "https://www.gstatic.com/firebasejs/12.11.0/firebase-analytics.js";

    // Your web app's Firebase configuration
    // (API key 등 값은 프론트엔드에 포함돼도 되며, 환경변수로 숨기는 방식은 서버에서 관리하는 게 일반적입니다.)
    const firebaseConfig = {
      apiKey: "AIzaSyA7ap55EZoEFYTgmmQOuPD9DC-4MfNI4c4",
      authDomain: "hansung-smart-ledger.firebaseapp.com",
      projectId: "hansung-smart-ledger",
      storageBucket: "hansung-smart-ledger.firebasestorage.app",
      messagingSenderId: "56040874401",
      appId: "1:56040874401:web:719e4890aeaa47b4b02be0",
      measurementId: "G-YSBZP17CKW",
    };

    // Initialize Firebase
    const fbApp = initializeApp(firebaseConfig);

    // 기존 non-module script에서도 접근 가능하도록 전역에 먼저 노출
    window.firebaseApp = fbApp;
    window.firebaseAnalytics = null;

    // 로컬 file:// 실행 등 Analytics가 지원되지 않는 환경에서도 Firestore는 계속 쓰도록 분리
    isAnalyticsSupported()
      .then(function (supported) {
        if (supported) {
          window.firebaseAnalytics = getAnalytics(fbApp);
        }
      })
      .catch(function (err) {
        console.warn("Firebase Analytics 초기화 건너뜀:", err);
      });
