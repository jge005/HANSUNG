(function () {
  var salesColumns = [
    { key: "date", label: "월일", w: "w-72", width: 72 },
    { key: "client", label: "업체", w: "w-120", width: 120 },
    { key: "code", label: "코드", w: "w-90", width: 90 },
    { key: "name", label: "품명", w: "w-170", width: 170 },
    { key: "qty", label: "수량", w: "w-70", width: 70 },
    { key: "price", label: "단가", w: "w-90", width: 90 },
    { key: "amount", label: "금액", w: "w-100", width: 100 },
    { key: "note1", label: "비고1", w: "w-110", width: 110 },
    { key: "note2", label: "비고2", w: "w-110", width: 110 },
  ];

  var roleCards = [
    { key: "accounting", title: "경리", icon: "scroll" },
    { key: "manager", title: "실장", icon: "sheet" },
    { key: "orders", title: "발주", icon: "folder" },
    { key: "work", title: "작업", icon: "settings" },
  ];

  var salesTabs = [
    { key: "manage", label: "매출관리", icon: "sheet" },
    { key: "status", label: "매출현황", icon: "clipboard" },
    { key: "work", label: "거래작업", icon: "scroll" },
    { key: "statement", label: "거래명세서", icon: "folder" },
    { key: "price", label: "매출단가", icon: "tags" },
    { key: "client", label: "업체리스트", icon: "building" },
  ];

  var icons = {
    lock: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>',
    arrowRight: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12h14M12 5l7 7-7 7"/></svg>',
    arrowLeft: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5M12 19l-7-7 7-7"/></svg>',
    sheet: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8"/></svg>',
    scroll: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg>',
    folder: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>',
    settings: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42"/></svg>',
    clipboard: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/></svg>',
    tags: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2H2v10l9.29 9.29a1 1 0 0 0 1.41 0l6.59-6.59a1 1 0 0 0 0-1.41L12 2z"/><circle cx="7.5" cy="7.5" r=".5" fill="currentColor" stroke="none"/></svg>',
    building: '<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 22V4a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v18M6 12h12M10 12v10M14 12v10"/></svg>',
  };

  function icon(name) {
    return icons[name] || icons.sheet;
  }

  var state = {
    entered: false,
    role: null,
    mainTab: null,
    subTab: "manage",
  };

  var DRAFT_ID_KEY = "ledgerDraftId";
  var LEGACY_DRAFT_ID_KEY = "workDraftId";
  var LOCAL_BACKUP_KEY = "ledgerDraftBackup";

  var ledgerState = {
    draftId: null,
    loadedFromFirebase: false,
    loadingPromise: null,
    saveTimer: null,
    saveRetryTimer: null,
    loadRetryTimer: null,
    dirty: false,
    localUpdatedAt: null,
    cloudOffline: false,
  };

  // 거래작업 탭에 표시할 “복사본” 데이터(매출현황 ST.rows를 절대 수정하지 않음)
  var workState = {
    items: [],
    info: {},
    draftId: null,
    loadedFromFirebase: false,
    loadingPromise: null,
    saveTimer: null,
  };

  var app = document.getElementById("app");

  function emptyRow() {
    return {
      __order: null,
      date: "", client: "", code: "", name: "",
      qty: "", price: "", amount: "", note1: "", note2: "",
    };
  }

  /* ========== Sales table engine ========== */
  var ST = {
    rows: [],
    selectedCell: { row: 0, col: 0 },
    selectedKeys: ["0:0"],
    dragging: false,
    isComposing: false,
    editMode: false,
    dragStart: null,
    dragMode: "replace",
    dragBaseKeys: [],
    undoStack: [],
    redoStack: [],
    editOrigin: { row: 0, col: 0, value: "" },
    editSnapshotRows: null,
    tableBuilt: false,
    globalListenersBound: false,
    host: null,
    scrollEl: null,
    inputEl: null,
    cellMap: {},
    rowHeadEls: [],
    colHeadEls: [],
    prevSelectedKeys: [],
    prevFillKeys: [],
    prevActiveKey: null,
    visibleRowsDirty: true,
    visibleRowIndexes: [],
    visibleRowOffsets: [0],
    visibleRowPosMap: {},
    renderRange: { start: -1, end: -1 },
    fillDragging: false,
    fillSourceRect: null,
    fillPreviewKeys: [],
    fillDirection: null,
    colWidths: [],
    rowHeights: [],
    resizing: null,
    filter: {
      colKey: "all",
      keyword: "",
    },
  };

  var ROW_HEADER_WIDTH = 46;
  var DEFAULT_ROW_HEIGHT = 28;
  var MIN_ROW_HEIGHT = 24;
  var MIN_COL_WIDTH = 56;
  var MIN_ROW_COUNT = 80;
  var ROW_GROWTH_CHUNK = 100;
  var COL_COUNT = salesColumns.length;

  function getRowCount() {
    return Math.max(ST.rows && ST.rows.length ? ST.rows.length : 0, MIN_ROW_COUNT);
  }

  function normalizeSalesRows(rows, minCount) {
    var normalized = Array.isArray(rows) ? rows.slice() : [];
    var targetCount = Math.max(normalized.length, minCount || MIN_ROW_COUNT);
    for (var i = 0; i < targetCount; i++) {
      normalized[i] = Object.assign(emptyRow(), normalized[i] || {});
      if (normalized[i].__order == null) normalized[i].__order = i;
      recalculateAmountForRow(normalized, i);
    }
    return normalized;
  }

  function ensureRowCapacity(minCount) {
    var targetCount = Math.max(minCount || 0, MIN_ROW_COUNT);
    if (getRowCount() >= targetCount) return false;

    var scrollTop = ST.scrollEl ? ST.scrollEl.scrollTop : 0;
    var scrollLeft = ST.scrollEl ? ST.scrollEl.scrollLeft : 0;

    ST.rows = normalizeSalesRows(ST.rows, targetCount);
    ST.visibleRowsDirty = true;
    initGridSizing();

    if (ST.tableBuilt && ST.host) {
      ST.tableBuilt = false;
      buildSalesTable(ST.host);
      if (ST.scrollEl) {
        ST.scrollEl.scrollTop = scrollTop;
        ST.scrollEl.scrollLeft = scrollLeft;
      }
    }

    return true;
  }

  function getLastUsedRowIndex(rows) {
    var source = Array.isArray(rows) ? rows : [];
    for (var i = source.length - 1; i >= 0; i--) {
      var row = source[i] || {};
      for (var j = 0; j < salesColumns.length; j++) {
        var key = salesColumns[j].key;
        if (row[key] != null && String(row[key]).trim() !== "") {
          return i;
        }
      }
    }
    return -1;
  }

  function getStoredRowCount() {
    return Math.max(MIN_ROW_COUNT, getLastUsedRowIndex(ST.rows) + 20);
  }

  function cloneItems(items) {
    if (!Array.isArray(items)) return [];
    return items.map(function (item) {
      return Object.assign({}, item || {});
    });
  }

  function cloneWorkInfo(info) {
    try {
      return JSON.parse(JSON.stringify(info || {}));
    } catch (err) {
      return {};
    }
  }

  function ensureDraftId() {
    if (ledgerState.draftId) return ledgerState.draftId;

    var id = null;
    try {
      id =
        localStorage.getItem(DRAFT_ID_KEY) ||
        localStorage.getItem(LEGACY_DRAFT_ID_KEY);
    } catch (err) {}

    if (!id) {
      id =
        window.crypto && crypto.randomUUID
          ? crypto.randomUUID()
          : String(Date.now());
    }

    ledgerState.draftId = id;
    workState.draftId = id;

    try {
      localStorage.setItem(DRAFT_ID_KEY, id);
      localStorage.setItem(LEGACY_DRAFT_ID_KEY, id);
    } catch (err) {}

    return id;
  }

  function getLedgerPayload() {
    var storedRowCount = getStoredRowCount();
    return {
      statusRows: normalizeSalesRows(ST.rows, storedRowCount).slice(0, storedRowCount),
      workItems: cloneItems(workState.items),
      workInfo: cloneWorkInfo(workState.info),
      sheetLayout: {
        colWidths: ST.colWidths.slice(),
        rowHeights: ST.rowHeights.slice(0, storedRowCount),
      },
      items: cloneItems(workState.items),
      updatedAt: new Date().toISOString(),
    };
  }

  function applyLedgerData(data) {
    if (!data) return;

    if (Array.isArray(data.statusRows)) {
      ST.rows = normalizeSalesRows(data.statusRows);
      ST.visibleRowsDirty = true;
    }

    if (Array.isArray(data.workItems)) {
      workState.items = cloneItems(data.workItems);
    } else if (Array.isArray(data.items)) {
      workState.items = cloneItems(data.items);
    }

    if (data.workInfo && typeof data.workInfo === "object") {
      workState.info = cloneWorkInfo(data.workInfo);
    }

    if (data.sheetLayout && typeof data.sheetLayout === "object") {
      if (Array.isArray(data.sheetLayout.colWidths)) {
        ST.colWidths = data.sheetLayout.colWidths.slice(0, COL_COUNT);
      }
      if (Array.isArray(data.sheetLayout.rowHeights)) {
        ST.rowHeights = data.sheetLayout.rowHeights.slice();
        ST.visibleRowsDirty = true;
      }
      initGridSizing();
    }

    if (ST.tableBuilt) {
      refreshGridValues();
      updateSelectionUI();
    }
  }

  function loadLocalLedgerBackup() {
    try {
      var raw = localStorage.getItem(LOCAL_BACKUP_KEY);
      if (raw) {
        var data = JSON.parse(raw);
        ledgerState.localUpdatedAt = data.updatedAt || null;
        applyLedgerData(data);
      }
    } catch (err) {
      console.warn("로컬 백업 불러오기 실패:", err);
    }
  }

  function saveLocalLedgerBackup() {
    try {
      var payload = getLedgerPayload();
      ledgerState.localUpdatedAt = payload.updatedAt;
      localStorage.setItem(LOCAL_BACKUP_KEY, JSON.stringify(payload));
    } catch (err) {
      console.warn("로컬 백업 저장 실패:", err);
    }
  }

  function isRetryableFirestoreError(err) {
    if (
      window.firebaseFirestoreApi &&
      typeof window.firebaseFirestoreApi.isRetryableFirestoreError === "function"
    ) {
      return window.firebaseFirestoreApi.isRetryableFirestoreError(err);
    }
    var code = err && err.code ? String(err.code) : "";
    var message = err && err.message ? String(err.message).toLowerCase() : "";
    return (
      code.indexOf("unavailable") >= 0 ||
      code.indexOf("failed-precondition") >= 0 ||
      message.indexOf("offline") >= 0 ||
      message.indexOf("network") >= 0
    );
  }

  function queueLedgerSaveRetry(delay) {
    if (ledgerState.saveRetryTimer) clearTimeout(ledgerState.saveRetryTimer);
    ledgerState.saveRetryTimer = setTimeout(function () {
      ledgerState.saveRetryTimer = null;
      if (ledgerState.dirty) saveLedgerDraftToFirebase();
    }, delay || 3000);
  }

  function queueLedgerLoadRetry(delay) {
    if (ledgerState.loadRetryTimer) clearTimeout(ledgerState.loadRetryTimer);
    ledgerState.loadRetryTimer = setTimeout(function () {
      ledgerState.loadRetryTimer = null;
      ledgerState.loadingPromise = null;
      ensureLedgerLoadedFromFirebase();
    }, delay || 3000);
  }

  function bindFirestoreRetryListeners() {
    if (window.__ledgerFirestoreRetryBound) return;
    window.__ledgerFirestoreRetryBound = true;

    window.addEventListener("online", function () {
      ledgerState.cloudOffline = false;
      ledgerState.loadingPromise = null;
      ensureLedgerLoadedFromFirebase();
      if (ledgerState.dirty) saveLedgerDraftToFirebase();
    });

    document.addEventListener("visibilitychange", function () {
      if (document.visibilityState !== "visible") return;
      if (!ledgerState.loadedFromFirebase) {
        ledgerState.loadingPromise = null;
        ensureLedgerLoadedFromFirebase();
      }
      if (ledgerState.dirty) saveLedgerDraftToFirebase();
    });
  }

  function saveLedgerDraftToFirebase(attempts) {
    attempts = attempts || 0;
    ensureDraftId();
    saveLocalLedgerBackup();

    if (!window.firebaseFirestoreApi || !window.firebaseFirestoreApi.saveWorkDraft) {
      if (attempts < 30) {
        setTimeout(function () {
          saveLedgerDraftToFirebase(attempts + 1);
        }, 100);
      }
      return;
    }

    window.firebaseFirestoreApi
      .saveWorkDraft(ledgerState.draftId, getLedgerPayload())
      .then(function () {
        ledgerState.loadedFromFirebase = true;
        ledgerState.dirty = false;
        ledgerState.cloudOffline = false;
      })
      .catch(function (err) {
        ledgerState.cloudOffline = isRetryableFirestoreError(err);
        if (ledgerState.cloudOffline) {
          console.warn("Firestore 저장 지연, 재시도 예정:", err);
          queueLedgerSaveRetry(3000);
          return;
        }
        console.error("Firestore 저장 실패:", err);
      });
  }

  function scheduleLedgerDraftSave() {
    ledgerState.dirty = true;
    saveLocalLedgerBackup();

    if (ledgerState.saveTimer) clearTimeout(ledgerState.saveTimer);
    ledgerState.saveTimer = setTimeout(function () {
      saveLedgerDraftToFirebase();
    }, 500);
  }

  function ensureLedgerLoadedFromFirebase() {
    if (ledgerState.loadedFromFirebase) return Promise.resolve();
    if (ledgerState.loadingPromise) return ledgerState.loadingPromise;

    loadLocalLedgerBackup();

    ledgerState.loadingPromise = new Promise(function (resolve) {
      var id = ensureDraftId();
      var tries = 0;

      function finish(success) {
        ledgerState.loadedFromFirebase = !!success;
        ledgerState.loadingPromise = null;
        resolve();
      }

      function tryLoad() {
        tries++;
        if (window.firebaseFirestoreApi && window.firebaseFirestoreApi.loadWorkDraft) {
          window.firebaseFirestoreApi
            .loadWorkDraft(id)
            .then(function (data) {
              if (data && !ledgerState.dirty) {
                var remoteTime = data.updatedAt ? Date.parse(data.updatedAt) : 0;
                var localTime = ledgerState.localUpdatedAt
                  ? Date.parse(ledgerState.localUpdatedAt)
                  : 0;
                if (!localTime || (remoteTime && remoteTime >= localTime)) {
                  applyLedgerData(data);
                  saveLocalLedgerBackup();
                }
              }
              ledgerState.cloudOffline = false;
              finish(true);
            })
            .catch(function (err) {
              ledgerState.cloudOffline = isRetryableFirestoreError(err);
              if (ledgerState.cloudOffline) {
                console.warn("Firestore 불러오기 지연, 재시도 예정:", err);
                queueLedgerLoadRetry(3000);
              } else {
                console.error("Firestore 불러오기 실패:", err);
              }
              finish(false);
            });
          return;
        }

        if (tries >= 30) {
          finish(false);
          return;
        }

        setTimeout(tryLoad, 100);
      }

      tryLoad();
    });

    return ledgerState.loadingPromise;
  }

  function setCellValueInRows(rows, r, c, value, normalize) {
    var key = salesColumns[c].key;
    var nextValue = normalize === false ? value : normalizeCellValue(c, value);
    if (normalize !== false) {
      nextValue = evaluateFormulaValue(nextValue, rows);
    }
    rows[r] = Object.assign({}, rows[r], (function () {
      var o = {};
      o[key] = nextValue;
      return o;
    })());
    recalculateAmountForRow(rows, r);
    return rows;
  }

  function parseCalcNumber(value) {
    if (value == null) return null;
    var text = String(value).trim().replace(/,/g, "");
    if (!text) return null;
    var num = Number(text);
    return isFinite(num) ? num : null;
  }

  function formatAmountValue(value) {
    if (!isFinite(value)) return "";
    if (Math.round(value) === value) {
      return String(value);
    }
    return String(Math.round(value * 100) / 100);
  }

  function formatDisplayNumber(value) {
    if (value == null || value === "") return "";
    var num = parseCalcNumber(value);
    if (num == null) return String(value);
    var rounded = Math.round(num * 100) / 100;
    return rounded.toLocaleString("ko-KR");
  }

  function columnIndexFromLabel(label) {
    var text = String(label || "").toUpperCase();
    if (!text) return -1;
    var value = 0;
    for (var i = 0; i < text.length; i++) {
      var code = text.charCodeAt(i);
      if (code < 65 || code > 90) return -1;
      value = value * 26 + (code - 64);
    }
    return value - 1;
  }

  function getRawValue(r, c) {
    var key = salesColumns[c].key;
    return ST.rows[r] && ST.rows[r][key] != null ? String(ST.rows[r][key]) : "";
  }

  function getCellNumericForFormula(rowIndex, colIndex, rows) {
    if (rowIndex < 0 || rowIndex >= rows.length || colIndex < 0 || colIndex >= COL_COUNT) return 0;
    var row = rows[rowIndex] || emptyRow();
    var raw = row[salesColumns[colIndex].key];
    var num = parseCalcNumber(raw);
    return num == null ? 0 : num;
  }

  function evaluateFormulaValue(text, rows) {
    if (text == null) return "";
    var expr = String(text).trim();
    if (!expr || expr.charAt(0) !== "=") return text;
    expr = expr.slice(1);

    expr = expr.replace(/SUM\(\s*([A-Z]+)(\d+)\s*:\s*([A-Z]+)(\d+)\s*\)/gi, function (_, c1, r1, c2, r2) {
      var startCol = columnIndexFromLabel(c1);
      var endCol = columnIndexFromLabel(c2);
      var startRow = Number(r1) - 1;
      var endRow = Number(r2) - 1;
      if (startCol < 0 || endCol < 0) return "0";
      var top = Math.min(startRow, endRow);
      var bottom = Math.max(startRow, endRow);
      var left = Math.min(startCol, endCol);
      var right = Math.max(startCol, endCol);
      var sum = 0;
      for (var rr = top; rr <= bottom; rr++) {
        for (var cc = left; cc <= right; cc++) {
          sum += getCellNumericForFormula(rr, cc, rows);
        }
      }
      return String(sum);
    });

    expr = expr.replace(/\b([A-Z]+)(\d+)\b/g, function (_, col, row) {
      return String(
        getCellNumericForFormula(Number(row) - 1, columnIndexFromLabel(col), rows)
      );
    });

    if (!/^[0-9+\-*/().,\s]+$/.test(expr)) return text;

    try {
      var result = Function('"use strict"; return (' + expr + ');')();
      if (typeof result === "number" && isFinite(result)) {
        return formatAmountValue(result);
      }
      return text;
    } catch (err) {
      return text;
    }
  }

  function formatCellDisplay(colIndex, value) {
    var key = salesColumns[colIndex].key;
    if (key === "amount") return formatDisplayNumber(value);
    return value == null ? "" : String(value);
  }

  function recalculateAmountForRow(rows, rowIndex) {
    if (!rows[rowIndex]) return;
    var qty = parseCalcNumber(rows[rowIndex].qty);
    var price = parseCalcNumber(rows[rowIndex].price);
    rows[rowIndex] = Object.assign({}, rows[rowIndex], {
      amount:
        qty != null && price != null
          ? formatAmountValue(qty * price)
          : "",
    });
  }

  function normalizeCellValue(colIndex, value) {
    var text = value == null ? "" : String(value);
    if (salesColumns[colIndex].key !== "date") return text;

    var slash = text.match(/^\s*(\d{1,2})\s*\/\s*(\d{1,2})\s*$/);
    if (slash) {
      return Number(slash[1]) + "월 " + Number(slash[2]) + "일";
    }

    var korean = text.match(/^\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일\s*$/);
    if (korean) {
      return Number(korean[1]) + "월 " + Number(korean[2]) + "일";
    }

    return text;
  }

  function commitCellValue(row, col) {
    var current = getRawValue(row, col);
    var normalized = normalizeCellValue(col, current);
    var evaluated = evaluateFormulaValue(normalized, ST.rows);
    if (current === evaluated) {
      ST.editSnapshotRows = null;
      return;
    }
    var beforeRows = ST.editSnapshotRows || cloneRows(ST.rows);
    var nextRows = ST.rows.slice();
    nextRows[row] = Object.assign({}, nextRows[row] || emptyRow());
    setCellValueInRows(nextRows, row, col, normalized, true);
    commitRowsMutation(nextRows, [row], true, beforeRows);
    ST.editSnapshotRows = null;
  }

  function commitActiveCellValue() {
    commitCellValue(ST.selectedCell.row, ST.selectedCell.col);
  }

  function parseMonthDay(value) {
    var text = value == null ? "" : String(value).trim();
    var match =
      text.match(/^(\d{1,2})\s*\/\s*(\d{1,2})$/) ||
      text.match(/^(\d{1,2})\s*월\s*(\d{1,2})\s*일$/);
    if (!match) return null;

    var month = Number(match[1]);
    var day = Number(match[2]);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;

    return new Date(2026, month - 1, day);
  }

  function formatMonthDay(date) {
    return date.getMonth() + 1 + "월 " + date.getDate() + "일";
  }

  function parseWeekday(value) {
    var text = value == null ? "" : String(value).trim();
    var shortDays = ["월", "화", "수", "목", "금", "토", "일"];
    var longDays = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"];
    var shortIndex = shortDays.indexOf(text);
    if (shortIndex >= 0) return { index: shortIndex, style: "short" };
    var longIndex = longDays.indexOf(text);
    if (longIndex >= 0) return { index: longIndex, style: "long" };
    return null;
  }

  function formatWeekday(index, style) {
    var shortDays = ["월", "화", "수", "목", "금", "토", "일"];
    var longDays = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"];
    var normalized = ((index % 7) + 7) % 7;
    return style === "long" ? longDays[normalized] : shortDays[normalized];
  }

  function buildFillSeries(sourceValues, count, direction) {
    var values = sourceValues.map(function (value) {
      return value == null ? "" : String(value);
    });
    if (!count) return [];

    var parsedDates = values.map(parseMonthDay);
    if (parsedDates.every(function (item) { return !!item; })) {
      var dateBase;
      var dateStep;
      if (direction === "up" || direction === "left") {
        dateBase = parsedDates[0];
        dateStep =
          parsedDates.length > 1
            ? Math.round((parsedDates[0] - parsedDates[1]) / 86400000)
            : -1;
      } else {
        dateBase = parsedDates[parsedDates.length - 1];
        dateStep =
          parsedDates.length > 1
            ? Math.round(
                (parsedDates[parsedDates.length - 1] -
                  parsedDates[parsedDates.length - 2]) /
                  86400000
              )
            : 1;
      }
      if (!dateStep) dateStep = direction === "up" || direction === "left" ? -1 : 1;
      return Array.from({ length: count }, function (_, index) {
        var next = new Date(dateBase);
        next.setDate(dateBase.getDate() + dateStep * (index + 1));
        return formatMonthDay(next);
      });
    }

    var parsedWeekdays = values.map(parseWeekday);
    if (parsedWeekdays.every(function (item) { return !!item; })) {
      var weekdayBase;
      var weekdayStep;
      var weekdayStyle =
        parsedWeekdays[0].style === "long" ||
        parsedWeekdays[parsedWeekdays.length - 1].style === "long"
          ? "long"
          : "short";
      if (direction === "up" || direction === "left") {
        weekdayBase = parsedWeekdays[0].index;
        weekdayStep =
          parsedWeekdays.length > 1
            ? parsedWeekdays[0].index - parsedWeekdays[1].index
            : -1;
      } else {
        weekdayBase = parsedWeekdays[parsedWeekdays.length - 1].index;
        weekdayStep =
          parsedWeekdays.length > 1
            ? parsedWeekdays[parsedWeekdays.length - 1].index -
              parsedWeekdays[parsedWeekdays.length - 2].index
            : 1;
      }
      if (!weekdayStep) weekdayStep = direction === "up" || direction === "left" ? -1 : 1;
      return Array.from({ length: count }, function (_, index) {
        return formatWeekday(weekdayBase + weekdayStep * (index + 1), weekdayStyle);
      });
    }

    if (direction === "up" || direction === "left") {
      return Array.from({ length: count }, function (_, index) {
        var sourceIndex =
          (values.length - 1 - (index % values.length) + values.length) %
          values.length;
        return values[sourceIndex];
      });
    }

    return Array.from({ length: count }, function (_, index) {
      return values[index % values.length];
    });
  }

  function initGridSizing() {
    var defaultWidths = salesColumns.map(function (col) {
      return col.width || 100;
    });
    ST.colWidths = defaultWidths.map(function (width, index) {
      return ST.colWidths[index] || width;
    });
    ST.rowHeights = Array.from({ length: getRowCount() }, function (_, index) {
      return ST.rowHeights[index] || DEFAULT_ROW_HEIGHT;
    });
  }

  function getColWidth(colIndex) {
    return ST.colWidths[colIndex] || salesColumns[colIndex].width || 100;
  }

  function getRowHeight(rowIndex) {
    return ST.rowHeights[rowIndex] || DEFAULT_ROW_HEIGHT;
  }

  function columnLabel(colIndex) {
    var label = "";
    var n = colIndex;
    while (n >= 0) {
      label = String.fromCharCode(65 + (n % 26)) + label;
      n = Math.floor(n / 26) - 1;
    }
    return label;
  }

  function cellAddress(row, col) {
    return columnLabel(col) + (row + 1);
  }

  function parseNumericValue(value) {
    if (value == null) return null;
    var text = String(value).trim().replace(/,/g, "");
    if (!text) return null;
    var num = Number(text);
    return isFinite(num) ? num : null;
  }

  function getSelectionMetrics() {
    var count = ST.selectedKeys.length;
    var numericCount = 0;
    var sum = 0;
    ST.selectedKeys.forEach(function (key) {
      var p = key.split(":");
      var value = parseNumericValue(getValue(Number(p[0]), Number(p[1])));
      if (value != null) {
        numericCount++;
        sum += value;
      }
    });
    return {
      count: count,
      numericCount: numericCount,
      sum: sum,
      avg: numericCount ? sum / numericCount : 0,
    };
  }

  function formatMetricNumber(value) {
    var rounded = Math.round(value * 100) / 100;
    return rounded.toLocaleString("ko-KR");
  }

  function updateStatusBar() {
    if (!ST.host) return;
    var metrics = getSelectionMetrics();
    var bounds = selectionBounds();
    var values = {
      address: cellAddress(ST.selectedCell.row, ST.selectedCell.col),
      range:
        ST.selectedKeys.length > 1
          ? cellAddress(bounds.top, bounds.left) + ":" + cellAddress(bounds.bottom, bounds.right)
          : "-",
      count: String(metrics.count),
      sum: formatMetricNumber(metrics.sum),
      avg: metrics.numericCount ? formatMetricNumber(metrics.avg) : "-",
    };
    Object.keys(values).forEach(function (key) {
      ST.host.querySelectorAll('[data-st-status="' + key + '"]').forEach(function (el) {
        el.textContent = values[key];
      });
    });
  }

  function applyGridSizingUI() {
    if (!ST.host) return;
    initGridSizing();
    ST.visibleRowsDirty = true;
    for (var c = 0; c < COL_COUNT; c++) {
      var colEl = ST.host.querySelector('col[data-col="' + c + '"]');
      if (colEl) colEl.style.width = getColWidth(c) + "px";
    }
    renderVirtualRows(true);
    if (ST.inputEl) {
      ST.inputEl.style.height = getRowHeight(ST.selectedCell.row) - 4 + "px";
      ST.inputEl.style.lineHeight = getRowHeight(ST.selectedCell.row) - 4 + "px";
    }
  }

  function applyColumnWidthUI(colIndex) {
    if (!ST.host) return;
    var colEl = ST.host.querySelector('col[data-col="' + colIndex + '"]');
    if (colEl) colEl.style.width = getColWidth(colIndex) + "px";
  }

  function applyRowHeightUI(rowIndex) {
    if (!ST.host) return;
    ST.visibleRowsDirty = true;
    renderVirtualRows(true);
    if (ST.inputEl && ST.selectedCell.row === rowIndex) {
      ST.inputEl.style.height = getRowHeight(rowIndex) - 4 + "px";
      ST.inputEl.style.lineHeight = getRowHeight(rowIndex) - 4 + "px";
    }
  }

  function selectAllCells() {
    ST.selectedCell = { row: 0, col: 0 };
    ST.selectedKeys = keysFromPoints(
      { row: 0, col: 0 },
      { row: getRowCount() - 1, col: COL_COUNT - 1 }
    );
    ST.editMode = false;
    updateSelectionUI();
    moveInputToCell(0, 0);
    syncEditOrigin();
    focusInput();
  }

  function selectRow(rowIndex) {
    ST.selectedCell = { row: rowIndex, col: 0 };
    ST.selectedKeys = keysFromPoints(
      { row: rowIndex, col: 0 },
      { row: rowIndex, col: COL_COUNT - 1 }
    );
    ST.editMode = false;
    updateSelectionUI();
    moveInputToCell(rowIndex, 0);
    syncEditOrigin();
    focusInput();
  }

  function selectColumn(colIndex) {
    ST.selectedCell = { row: 0, col: colIndex };
    ST.selectedKeys = keysFromPoints(
      { row: 0, col: colIndex },
      { row: getRowCount() - 1, col: colIndex }
    );
    ST.editMode = false;
    updateSelectionUI();
    moveInputToCell(0, colIndex);
    syncEditOrigin();
    focusInput();
  }

  function cutSelectionAsync() {
    var text = getSelectionText();
    if (!text) return;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(function () {});
    }
    clearSelectionCells();
  }

  function moveByTab(backward) {
    var row = ST.selectedCell.row;
    var col = ST.selectedCell.col + (backward ? -1 : 1);
    if (col < 0) {
      col = COL_COUNT - 1;
      row = Math.max(0, row - 1);
    } else if (col >= COL_COUNT) {
      col = 0;
      row = row + 1;
    }
    if (row >= getRowCount() - 1) ensureRowCapacity(row + ROW_GROWTH_CHUNK);
    moveSelection(row, col);
  }

  function startResize(kind, index, clientX, clientY) {
    ST.resizing = {
      kind: kind,
      index: index,
      startX: clientX,
      startY: clientY,
      startSize: kind === "col" ? getColWidth(index) : getRowHeight(index),
    };
  }

  function rowMatchesFilter(row) {
    var keyword = (ST.filter.keyword || "").trim().toLowerCase();
    if (!keyword) return true;
    var keys =
      ST.filter.colKey && ST.filter.colKey !== "all"
        ? [ST.filter.colKey]
        : salesColumns.map(function (col) { return col.key; });
    return keys.some(function (key) {
      var colIndex = salesColumns.findIndex(function (col) { return col.key === key; });
      var raw = row && row[key] != null ? row[key] : "";
      return formatCellDisplay(colIndex, raw).toLowerCase().indexOf(keyword) >= 0;
    });
  }

  function applyFilterUI() {
    if (!ST.host) return;
    ST.visibleRowsDirty = true;
    renderVirtualRows(true);
    updateSelectionUI();
  }

  function sortRowsByColumn(colKey, direction) {
    var colIndex = salesColumns.findIndex(function (col) { return col.key === colKey; });
    if (colIndex < 0 || !direction) return;
    applyRowsChange(function (rows) {
      rows.sort(function (a, b) {
        var av = a ? a[colKey] : "";
        var bv = b ? b[colKey] : "";
        var an = parseCalcNumber(av);
        var bn = parseCalcNumber(bv);
        var result = 0;
        if (an != null && bn != null) {
          result = an - bn;
        } else {
          result = String(formatCellDisplay(colIndex, av)).localeCompare(
            String(formatCellDisplay(colIndex, bv)),
            "ko"
          );
        }
        if (!result) {
          result = (a && a.__order != null ? a.__order : 0) - (b && b.__order != null ? b.__order : 0);
        }
        return direction === "desc" ? -result : result;
      });
      return rows;
    });
  }

  function resetSortOrder() {
    applyRowsChange(function (rows) {
      rows.sort(function (a, b) {
        return (a && a.__order != null ? a.__order : 0) - (b && b.__order != null ? b.__order : 0);
      });
      return rows;
    });
  }

  function cellKey(r, c) {
    return r + ":" + c;
  }

  function rectFromKeys(keys) {
    var coords = keys.map(function (k) {
      var p = k.split(":");
      return [Number(p[0]), Number(p[1])];
    });
    var rs = coords.map(function (x) { return x[0]; });
    var cs = coords.map(function (x) { return x[1]; });
    return {
      top: Math.min.apply(null, rs),
      bottom: Math.max.apply(null, rs),
      left: Math.min.apply(null, cs),
      right: Math.max.apply(null, cs),
    };
  }

  function keysFromPoints(a, b) {
    var top = Math.min(a.row, b.row);
    var bottom = Math.max(a.row, b.row);
    var left = Math.min(a.col, b.col);
    var right = Math.max(a.col, b.col);
    var keys = [];
    for (var r = top; r <= bottom; r++) {
      for (var c = left; c <= right; c++) {
        keys.push(cellKey(r, c));
      }
    }
    return keys;
  }

  function cloneRows(src) {
    return src.map(function (row) {
      var o = {};
      for (var k in row) o[k] = row[k];
      return o;
    });
  }

  function getValue(r, c) {
    return formatCellDisplay(c, getRawValue(r, c));
  }

  function selectionBounds() {
    return rectFromKeys(ST.selectedKeys);
  }

  function updateRenderedCell(r, c) {
    if (!ST.tableBuilt || !ST.host) return;
    var td = ST.cellMap[cellKey(r, c)];
    if (!td) return;
    var disp = td.querySelector(".st-display");
    var ghost = td.querySelector(".st-ghost");
    var value = getValue(r, c);
    if (disp) disp.textContent = value;
    if (ghost) ghost.textContent = value;
    if (ST.selectedCell.row === r && ST.selectedCell.col === c && ST.inputEl) {
      ST.inputEl.value = getRawValue(r, c);
    }
  }

  function refreshRenderedRows(rowIndexes) {
    if (!Array.isArray(rowIndexes)) return;
    var seen = {};
    rowIndexes.forEach(function (rowIndex) {
      if (seen[rowIndex]) return;
      seen[rowIndex] = true;
      for (var c = 0; c < COL_COUNT; c++) {
        updateRenderedCell(rowIndex, c);
      }
    });
    updateStatusBar();
  }

  function commitRowsMutation(nextRows, changedRows, trackHistory, beforeOverride) {
    if (trackHistory === undefined) trackHistory = true;
    var before = trackHistory ? (beforeOverride || cloneRows(ST.rows)) : null;
    if (trackHistory && Array.isArray(changedRows) && changedRows.length) {
      var changed = changedRows.some(function (rowIndex) {
        return JSON.stringify(before[rowIndex] || {}) !== JSON.stringify(nextRows[rowIndex] || {});
      });
      if (!changed) return false;
    } else if (trackHistory && JSON.stringify(before) === JSON.stringify(nextRows)) {
      return false;
    }
    if (!trackHistory && ST.rows === nextRows) return false;
    if (trackHistory) {
      ST.undoStack.push(before);
      if (ST.undoStack.length > 200) ST.undoStack.shift();
      ST.redoStack = [];
    }
    ST.rows = nextRows;
    if (ST.filter && ST.filter.keyword) {
      ST.visibleRowsDirty = true;
      renderVirtualRows(true);
      updateSelectionUI();
    } else {
      refreshRenderedRows(changedRows || []);
    }
    scheduleLedgerDraftSave();
    return true;
  }

  function applyRowsChange(producer, trackHistory) {
    if (trackHistory === undefined) trackHistory = true;
    var before = cloneRows(ST.rows);
    var after = producer(cloneRows(ST.rows));
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    if (trackHistory) {
      ST.undoStack.push(before);
      if (ST.undoStack.length > 200) ST.undoStack.shift();
      ST.redoStack = [];
    }
    ST.rows = after;
    ST.visibleRowsDirty = true;
    refreshGridValues();
    scheduleLedgerDraftSave();
  }

  function setValue(r, c, value) {
    if (getRawValue(r, c) === value) return;
    if (!ST.editSnapshotRows) {
      ST.editSnapshotRows = cloneRows(ST.rows);
    }
    var nextRows = ST.rows.slice();
    nextRows[r] = Object.assign({}, nextRows[r] || emptyRow());
    setCellValueInRows(nextRows, r, c, value, false);
    commitRowsMutation(nextRows, [r], false);
  }

  function rebuildVisibleRowsData() {
    if (!ST.visibleRowsDirty && ST.visibleRowIndexes.length) return;
    var indexes = [];
    var offsets = [0];
    var posMap = {};
    var total = 0;

    for (var r = 0; r < getRowCount(); r++) {
      if (!rowMatchesFilter(ST.rows[r])) continue;
      posMap[r] = indexes.length;
      indexes.push(r);
      total += getRowHeight(r);
      offsets.push(total);
    }

    if (!indexes.length) {
      posMap[0] = 0;
      indexes.push(0);
      total = getRowHeight(0);
      offsets = [0, total];
    }

    ST.visibleRowIndexes = indexes;
    ST.visibleRowOffsets = offsets;
    ST.visibleRowPosMap = posMap;
    ST.visibleRowsDirty = false;
  }

  function getVisibleRowPosition(actualRow) {
    return ST.visibleRowPosMap.hasOwnProperty(actualRow)
      ? ST.visibleRowPosMap[actualRow]
      : -1;
  }

  function findVisibleIndexByOffset(offset) {
    var offsets = ST.visibleRowOffsets;
    var low = 0;
    var high = Math.max(0, offsets.length - 2);

    while (low <= high) {
      var mid = Math.floor((low + high) / 2);
      var start = offsets[mid];
      var end = offsets[mid + 1];
      if (offset < start) {
        high = mid - 1;
      } else if (offset >= end) {
        low = mid + 1;
      } else {
        return mid;
      }
    }

    return Math.max(0, Math.min(offsets.length - 2, low));
  }

  function buildTableRowHtml(r) {
    var html = '<tr data-row="' + r + '" style="height:' + getRowHeight(r) + 'px">';
    html +=
      '<th class="st-row-head" data-row-head="' + r + '">' +
      (r + 1) +
      '<div class="st-resizer-row" data-row-resizer="' + r + '"></div>' +
      "</th>";
    for (var c = 0; c < COL_COUNT; c++) {
      var v = getValue(r, c);
      var active = r === ST.selectedCell.row && c === ST.selectedCell.col;
      html += '<td data-r="' + r + '" data-c="' + c + '" class="' + (active ? "sel-in sel-active" : "") + '">';
      html += '<div class="st-cell-inner" style="min-height:' + (getRowHeight(r) - 4) + 'px">';
      html +=
        '<div class="st-display" style="' +
        (active ? "display:none;" : "") +
        'height:' + (getRowHeight(r) - 4) + 'px;' +
        'line-height:' + (getRowHeight(r) - 4) + 'px;' +
        '">' +
        escapeHtml(v) +
        "</div>";
      html += '<div class="st-ghost" style="display:none;height:' + (getRowHeight(r) - 4) + 'px;line-height:' + (getRowHeight(r) - 4) + 'px"></div>';
      html += "</div></td>";
    }
    html += "</tr>";
    return html;
  }

  function renderVirtualRows(force) {
    if (!ST.tableBuilt || !ST.host || !ST.scrollEl) return;
    var tbody = ST.host.querySelector("tbody");
    if (!tbody) return;

    rebuildVisibleRowsData();

    var viewportHeight = ST.scrollEl.clientHeight || 600;
    var scrollTop = ST.scrollEl.scrollTop || 0;
    var startIndex = Math.max(0, findVisibleIndexByOffset(scrollTop) - 12);
    var endIndex = Math.min(
      ST.visibleRowIndexes.length - 1,
      findVisibleIndexByOffset(scrollTop + viewportHeight) + 16
    );

    if (!force && ST.renderRange.start === startIndex && ST.renderRange.end === endIndex) {
      return;
    }

    var topHeight = ST.visibleRowOffsets[startIndex] || 0;
    var bottomHeight =
      ST.visibleRowOffsets[ST.visibleRowOffsets.length - 1] -
      (ST.visibleRowOffsets[endIndex + 1] || 0);
    var html = "";

    if (topHeight > 0) {
      html += '<tr data-spacer="top"><td colspan="' + (COL_COUNT + 1) + '" style="height:' + topHeight + 'px;padding:0;border:none;background:transparent"></td></tr>';
    }

    for (var i = startIndex; i <= endIndex; i++) {
      html += buildTableRowHtml(ST.visibleRowIndexes[i]);
    }

    if (bottomHeight > 0) {
      html += '<tr data-spacer="bottom"><td colspan="' + (COL_COUNT + 1) + '" style="height:' + bottomHeight + 'px;padding:0;border:none;background:transparent"></td></tr>';
    }

    tbody.innerHTML = html;
    ST.cellMap = {};
    tbody.querySelectorAll("td[data-r][data-c]").forEach(function (td) {
      ST.cellMap[cellKey(Number(td.getAttribute("data-r")), Number(td.getAttribute("data-c")))] = td;
    });
    ST.rowHeadEls = Array.prototype.slice.call(tbody.querySelectorAll("[data-row-head]"));
    ST.renderRange = { start: startIndex, end: endIndex };
  }

  function ensureRowVisible(rowIndex) {
    if (!ST.scrollEl) return;
    var visiblePos = getVisibleRowPosition(rowIndex);
    if (visiblePos < 0) return;
    var top = ST.visibleRowOffsets[visiblePos] || 0;
    var bottom = ST.visibleRowOffsets[visiblePos + 1] || top + getRowHeight(rowIndex);
    var viewportTop = ST.scrollEl.scrollTop;
    var viewportBottom = viewportTop + ST.scrollEl.clientHeight;

    if (top < viewportTop) {
      ST.scrollEl.scrollTop = top;
    } else if (bottom > viewportBottom) {
      ST.scrollEl.scrollTop = Math.max(0, bottom - ST.scrollEl.clientHeight);
    }
  }

  function moveSelection(row, col) {
    commitActiveCellValue();
    if (row >= getRowCount() - 5) ensureRowCapacity(row + ROW_GROWTH_CHUNK);
    var nr = Math.max(0, Math.min(getRowCount() - 1, row));
    var nc = Math.max(0, Math.min(COL_COUNT - 1, col));
    ST.selectedCell = { row: nr, col: nc };
    ST.selectedKeys = [cellKey(nr, nc)];
    ST.editMode = false;
    ST.isComposing = false;
    ST.fillPreviewKeys = [];
    ST.fillSourceRect = null;
    ST.fillDirection = null;
    updateSelectionUI();
    moveInputToCell(nr, nc);
    syncEditOrigin();
    focusInput();
  }

  function clearSelectionCells() {
    applyRowsChange(function (current) {
      ST.selectedKeys.forEach(function (key) {
        var p = key.split(":");
        var r = Number(p[0]);
        var c = Number(p[1]);
        var k = salesColumns[c].key;
        var row = Object.assign({}, current[r]);
        row[k] = "";
        current[r] = row;
      });
      return current;
    });
  }

  function getSelectionText() {
    if (!ST.selectedKeys.length) return "";
    var b = selectionBounds();
    var lines = [];
    for (var r = b.top; r <= b.bottom; r++) {
      var cols = [];
      for (var c = b.left; c <= b.right; c++) {
        cols.push(getValue(r, c));
      }
      lines.push(cols.join("\t"));
    }
    return lines.join("\n");
  }

  function pasteTextIntoSelection(text) {
    if (!text) return;
    var startRow = ST.selectedCell.row;
    var startCol = ST.selectedCell.col;
    var lines = text.replace(/\r/g, "").split("\n");
    if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
    var grid = lines.map(function (line) {
      return line.split("\t");
    });
    var sourceRows = Math.max(grid.length, 1);
    var sourceCols = Math.max.apply(
      null,
      grid.map(function (line) {
        return line.length;
      })
    );
    var selectionRect = selectionBounds();
    var selectionIsRect =
      ST.selectedKeys.length ===
      (selectionRect.bottom - selectionRect.top + 1) *
        (selectionRect.right - selectionRect.left + 1);
    var requiredRowCount = ST.selectedKeys.length > 1 && selectionIsRect
      ? selectionRect.bottom + 1
      : startRow + sourceRows;
    ensureRowCapacity(requiredRowCount + 5);

    applyRowsChange(function (current) {
      if (ST.selectedKeys.length > 1 && sourceRows === 1 && sourceCols === 1) {
        ST.selectedKeys.forEach(function (key) {
          var pos = key.split(":");
          setCellValueInRows(
            current,
            Number(pos[0]),
            Number(pos[1]),
            grid[0][0],
            true
          );
        });
        return current;
      }

      if (ST.selectedKeys.length > 1 && selectionIsRect) {
        for (var r = selectionRect.top; r <= selectionRect.bottom; r++) {
          for (var c = selectionRect.left; c <= selectionRect.right; c++) {
            var rowOffset = r - selectionRect.top;
            var colOffset = c - selectionRect.left;
            var rowValues = grid[rowOffset % sourceRows] || [""];
            var nextValue = rowValues[colOffset % rowValues.length] || "";
            setCellValueInRows(current, r, c, nextValue, true);
          }
        }
        return current;
      }

      grid.forEach(function (parts, rowOffset) {
        parts.forEach(function (value, colOffset) {
          var row = startRow + rowOffset;
          var col = startCol + colOffset;
          if (row < current.length && col < salesColumns.length) {
            setCellValueInRows(current, row, col, value, true);
          }
        });
      });
      return current;
    });

    var endRow;
    var endCol;
    if (ST.selectedKeys.length > 1 && selectionIsRect) {
      endRow = selectionRect.bottom;
      endCol = selectionRect.right;
      ST.selectedKeys = keysFromPoints(
        { row: selectionRect.top, col: selectionRect.left },
        { row: endRow, col: endCol }
      );
    } else {
      var pastedRows = Math.max(grid.length, 1);
      var pastedCols = sourceCols;
      endRow = Math.max(0, Math.min(getRowCount() - 1, startRow + pastedRows - 1));
      endCol = Math.max(0, Math.min(COL_COUNT - 1, startCol + pastedCols - 1));
      ST.selectedKeys = keysFromPoints(
        { row: startRow, col: startCol },
        { row: endRow, col: endCol }
      );
    }

    ST.selectedCell = { row: endRow, col: endCol };
    ST.editMode = false;
    updateSelectionUI();
    moveInputToCell(ST.selectedCell.row, ST.selectedCell.col);
    syncEditOrigin();
    focusInput();
  }

  function undo() {
    var prev = ST.undoStack.pop();
    if (!prev) return;
    ST.redoStack.push(cloneRows(ST.rows));
    ST.rows = cloneRows(prev);
    ST.editMode = false;
    ST.isComposing = false;
    refreshGridValues();
    scheduleLedgerDraftSave();
    updateSelectionUI();
    moveInputToCell(ST.selectedCell.row, ST.selectedCell.col);
    syncEditOrigin();
    focusInput();
  }

  function redo() {
    var next = ST.redoStack.pop();
    if (!next) return;
    ST.undoStack.push(cloneRows(ST.rows));
    ST.rows = cloneRows(next);
    ST.editMode = false;
    ST.isComposing = false;
    refreshGridValues();
    scheduleLedgerDraftSave();
    updateSelectionUI();
    moveInputToCell(ST.selectedCell.row, ST.selectedCell.col);
    syncEditOrigin();
    focusInput();
  }

  function snapshotEditOrigin() {
    ST.editOrigin = {
      row: ST.selectedCell.row,
      col: ST.selectedCell.col,
      value: getRawValue(ST.selectedCell.row, ST.selectedCell.col),
    };
  }

  function syncEditOrigin() {
    if (!ST.editMode) snapshotEditOrigin();
  }

  function refreshGridValues() {
    if (!ST.tableBuilt || !ST.host) return;
    renderVirtualRows(true);
    if (ST.inputEl) {
      ST.inputEl.value = getRawValue(ST.selectedCell.row, ST.selectedCell.col);
    }
    syncInputOverlay();
    updateStatusBar();
  }

  function updateSelectionUI() {
    if (!ST.tableBuilt || !ST.host) return;
    var currentSelected = {};
    var currentFill = {};
    ST.selectedKeys.forEach(function (key) { currentSelected[key] = true; });
    ST.fillPreviewKeys.forEach(function (key) { currentFill[key] = true; });
    var activeKey = cellKey(ST.selectedCell.row, ST.selectedCell.col);
    var dirtyKeys = {};

    ST.prevSelectedKeys.forEach(function (key) { dirtyKeys[key] = true; });
    ST.prevFillKeys.forEach(function (key) { dirtyKeys[key] = true; });
    ST.selectedKeys.forEach(function (key) { dirtyKeys[key] = true; });
    ST.fillPreviewKeys.forEach(function (key) { dirtyKeys[key] = true; });
    if (ST.prevActiveKey) dirtyKeys[ST.prevActiveKey] = true;
    dirtyKeys[activeKey] = true;

    Object.keys(dirtyKeys).forEach(function (key) {
      var td = ST.cellMap[key];
      if (!td) return;
      td.classList.toggle("sel-in", !!currentSelected[key]);
      td.classList.toggle("sel-fill", !!currentFill[key]);
      td.classList.toggle("sel-active", key === activeKey);
    });

    var oldHandle = ST.host.querySelector(".st-fill-handle");
    if (oldHandle) oldHandle.remove();

    if (!ST.fillDragging && ST.selectedKeys.length) {
      var bounds = selectionBounds();
      var handleCell = ST.host.querySelector(
        '.st-table td[data-r="' + bounds.bottom + '"][data-c="' + bounds.right + '"]'
      );
      if (handleCell) {
        var handle = document.createElement("div");
        handle.className = "st-fill-handle";
        handle.title = "자동 채우기";
        handleCell.appendChild(handle);
      }
    }

    ST.colHeadEls.forEach(function (el) {
      el.classList.toggle(
        "st-head-active",
        ST.selectedKeys.some(function (key) {
          return Number(key.split(":")[1]) === Number(el.getAttribute("data-col-head"));
        })
      );
    });
    ST.rowHeadEls.forEach(function (el) {
      el.classList.toggle(
        "st-head-active",
        ST.selectedKeys.some(function (key) {
          return Number(key.split(":")[0]) === Number(el.getAttribute("data-row-head"));
        })
      );
    });
    ST.prevSelectedKeys = ST.selectedKeys.slice();
    ST.prevFillKeys = ST.fillPreviewKeys.slice();
    ST.prevActiveKey = activeKey;
    updateStatusBar();
  }

  function updateFillPreview(row, col) {
    if (!ST.fillSourceRect) return;
    var source = ST.fillSourceRect;
    var down = row - source.bottom;
    var up = source.top - row;
    var right = col - source.right;
    var left = source.left - col;
    var verticalDistance = Math.max(down, up, 0);
    var horizontalDistance = Math.max(right, left, 0);

    ST.fillPreviewKeys = [];
    ST.fillDirection = null;

    if (!verticalDistance && !horizontalDistance) {
      updateSelectionUI();
      return;
    }

    if (verticalDistance >= horizontalDistance) {
      if (down > 0) {
        ST.fillDirection = "down";
        for (var r = source.bottom + 1; r <= row; r++) {
          for (var c = source.left; c <= source.right; c++) {
            ST.fillPreviewKeys.push(cellKey(r, c));
          }
        }
      } else if (up > 0) {
        ST.fillDirection = "up";
        for (var r2 = row; r2 < source.top; r2++) {
          for (var c2 = source.left; c2 <= source.right; c2++) {
            ST.fillPreviewKeys.push(cellKey(r2, c2));
          }
        }
      }
    } else {
      if (right > 0) {
        ST.fillDirection = "right";
        for (var c3 = source.right + 1; c3 <= col; c3++) {
          for (var r3 = source.top; r3 <= source.bottom; r3++) {
            ST.fillPreviewKeys.push(cellKey(r3, c3));
          }
        }
      } else if (left > 0) {
        ST.fillDirection = "left";
        for (var c4 = col; c4 < source.left; c4++) {
          for (var r4 = source.top; r4 <= source.bottom; r4++) {
            ST.fillPreviewKeys.push(cellKey(r4, c4));
          }
        }
      }
    }

    updateSelectionUI();
  }

  function applyFillPreview() {
    if (!ST.fillSourceRect || !ST.fillPreviewKeys.length || !ST.fillDirection) return;

    var source = ST.fillSourceRect;
    var direction = ST.fillDirection;
    var bounds = rectFromKeys(ST.fillPreviewKeys);
    ensureRowCapacity(bounds.bottom + 5);

    applyRowsChange(function (rows) {
      if (direction === "down" || direction === "up") {
        var count = Math.abs(bounds.bottom - bounds.top) + 1;
        var targetRows =
          direction === "down"
            ? Array.from({ length: count }, function (_, index) {
                return source.bottom + index + 1;
              })
            : Array.from({ length: count }, function (_, index) {
                return source.top - index - 1;
              });

        for (var c = source.left; c <= source.right; c++) {
          var sourceValues = [];
          for (var r = source.top; r <= source.bottom; r++) {
            sourceValues.push(getValue(r, c));
          }
          var filled = buildFillSeries(sourceValues, count, direction);
          targetRows.forEach(function (targetRow, index) {
            setCellValueInRows(rows, targetRow, c, filled[index], true);
          });
        }
      } else {
        var count2 = Math.abs(bounds.right - bounds.left) + 1;
        var targetCols =
          direction === "right"
            ? Array.from({ length: count2 }, function (_, index) {
                return source.right + index + 1;
              })
            : Array.from({ length: count2 }, function (_, index) {
                return source.left - index - 1;
              });

        for (var r2 = source.top; r2 <= source.bottom; r2++) {
          var sourceValues2 = [];
          for (var c5 = source.left; c5 <= source.right; c5++) {
            sourceValues2.push(getValue(r2, c5));
          }
          var filled2 = buildFillSeries(sourceValues2, count2, direction);
          targetCols.forEach(function (targetCol, index2) {
            setCellValueInRows(rows, r2, targetCol, filled2[index2], true);
          });
        }
      }
      return rows;
    });

    var start;
    var end;
    if (direction === "down") {
      start = { row: source.top, col: source.left };
      end = { row: bounds.bottom, col: source.right };
    } else if (direction === "up") {
      start = { row: bounds.top, col: source.left };
      end = { row: source.bottom, col: source.right };
    } else if (direction === "right") {
      start = { row: source.top, col: source.left };
      end = { row: source.bottom, col: bounds.right };
    } else {
      start = { row: source.top, col: bounds.left };
      end = { row: source.bottom, col: source.right };
    }

    ST.selectedKeys = keysFromPoints(start, end);
    ST.selectedCell = { row: end.row, col: end.col };
    ST.fillPreviewKeys = [];
    ST.fillSourceRect = null;
    ST.fillDirection = null;
    ST.editMode = false;
    updateSelectionUI();
    moveInputToCell(ST.selectedCell.row, ST.selectedCell.col);
    syncEditOrigin();
    focusInput();
  }

  function syncInputOverlay() {
    if (!ST.tableBuilt || !ST.inputEl || !ST.host) return;
    var r = ST.selectedCell.row;
    var c = ST.selectedCell.col;
    ensureRowVisible(r);
    renderVirtualRows();
    var td = ST.host.querySelector('.st-table td[data-r="' + r + '"][data-c="' + c + '"]');
    if (!td) return;
    var inner = td.querySelector(".st-cell-inner");
    if (!inner) return;
    var v = getValue(r, c);
    ST.inputEl.value = v;
    ST.inputEl.style.height = getRowHeight(r) - 4 + "px";
    var gh = inner.querySelector(".st-ghost");
    if (ST.editMode) {
      ST.inputEl.classList.remove("no-edit");
      if (gh) {
        gh.textContent = "";
        gh.style.display = "none";
      }
    } else {
      ST.inputEl.classList.add("no-edit");
      if (gh) {
        gh.style.display = "block";
        gh.textContent = v;
      }
    }
    var len = ST.inputEl.value.length;
    requestAnimationFrame(function () {
      if (!ST.inputEl) return;
      if (ST.editMode) ST.inputEl.setSelectionRange(len, len);
      else ST.inputEl.setSelectionRange(0, len);
    });
  }

  function focusInput() {
    if (ST.scrollEl) ST.scrollEl.focus();
    requestAnimationFrame(function () {
      if (ST.inputEl) {
        ST.inputEl.focus();
        var len = ST.inputEl.value.length;
        if (ST.editMode) ST.inputEl.setSelectionRange(len, len);
        else ST.inputEl.setSelectionRange(0, len);
      }
    });
  }

  function buildSalesTable(host) {
    ST.host = host;
    initGridSizing();
    ST.rows = normalizeSalesRows(ST.rows);
    ST.selectedCell = {
      row: Math.max(0, Math.min(getRowCount() - 1, ST.selectedCell.row || 0)),
      col: Math.max(0, Math.min(COL_COUNT - 1, ST.selectedCell.col || 0)),
    };
    if (!ST.selectedKeys.length) {
      ST.selectedKeys = [cellKey(ST.selectedCell.row, ST.selectedCell.col)];
    }
    ST.editMode = false;
    ST.tableBuilt = true;
    ST.visibleRowsDirty = true;

    var html =
      '<div class="st-wrap">' +
      '<div class="st-toolbar">' +
      '<div>현재 셀 <strong data-st-status="address">A1</strong></div>' +
      '<div>선택 영역 <strong data-st-status="range">-</strong></div>' +
      '</div>' +
      '<div class="st-scroll" tabindex="0">' +
      '<table class="st-table"><colgroup>' +
      '<col style="width:' + ROW_HEADER_WIDTH + 'px" />';
    salesColumns.forEach(function (col, colIndex) {
      html += '<col data-col="' + colIndex + '" style="width:' + getColWidth(colIndex) + 'px" />';
    });
    html += '</colgroup><thead><tr>';
    html += '<th class="st-corner"></th>';
    salesColumns.forEach(function (col, colIndex) {
      html +=
        '<th class="st-col-head" data-col-head="' + colIndex + '">' +
        columnLabel(colIndex) +
        '<div style="font-size:11px;font-weight:600;color:#8b82a3">' + col.label + '</div>' +
        '<div class="st-resizer-col" data-col-resizer="' + colIndex + '"></div>' +
        "</th>";
    });
    html += "</tr></thead><tbody></tbody></table>";
    html +=
      '</div>' +
      '<div class="st-statusbar">' +
      '<div class="st-status-main">' +
      '<span class="st-status-pill">현재 셀 <strong data-st-status="address">A1</strong></span>' +
      '<span class="st-status-pill">선택 영역 <strong data-st-status="range">-</strong></span>' +
      '</div>' +
      '<div class="st-status-metrics">' +
      '<span class="st-status-pill">개수 <strong data-st-status="count">1</strong></span>' +
      '<span class="st-status-pill">합계 <strong data-st-status="sum">0</strong></span>' +
      '<span class="st-status-pill">평균 <strong data-st-status="avg">-</strong></span>' +
      '</div>' +
      '</div>' +
      '</div>';

    host.innerHTML = html;
    ST.scrollEl = host.querySelector(".st-scroll");
    ST.colHeadEls = Array.prototype.slice.call(host.querySelectorAll("[data-col-head]"));
    ST.rowHeadEls = [];
    ST.prevSelectedKeys = [];
    ST.prevFillKeys = [];
    ST.prevActiveKey = null;
    ST.renderRange = { start: -1, end: -1 };

    var tbody = host.querySelector("tbody");
    renderVirtualRows(true);
    ST.inputEl = document.createElement("input");
    ST.inputEl.type = "text";
    ST.inputEl.setAttribute("spellcheck", "false");
    ST.inputEl.className = "st-input no-edit";
    ST.inputEl.style.height = getRowHeight(0) - 4 + "px";
    ST.inputEl.style.lineHeight = getRowHeight(0) - 4 + "px";
    var firstInner = tbody.querySelector(
      'td[data-r="' + ST.selectedCell.row + '"][data-c="' + ST.selectedCell.col + '"] .st-cell-inner'
    );
    if (firstInner) firstInner.appendChild(ST.inputEl);
    ST.inputEl.value = getRawValue(ST.selectedCell.row, ST.selectedCell.col);
    syncInputOverlay();

    ST.scrollEl.addEventListener("keydown", onTableKeyDown);
    ST.scrollEl.addEventListener("copy", onCopy);
    ST.scrollEl.addEventListener("paste", onPaste);
    ST.scrollEl.addEventListener("dragstart", function (e) {
      e.preventDefault();
    });

    tbody.addEventListener("mousedown", onTbodyMouseDown);
    tbody.addEventListener("dblclick", onTbodyDblClick);
    if (!ST.globalListenersBound) {
      document.addEventListener("mousemove", onDocMouseMove);
      window.addEventListener("mouseup", onWindowMouseUp);
      ST.globalListenersBound = true;
    }
    ST.scrollEl.addEventListener("scroll", function () {
      if (ST.scrollEl.scrollTop + ST.scrollEl.clientHeight >= ST.scrollEl.scrollHeight - 120) {
        ensureRowCapacity(getRowCount() + ROW_GROWTH_CHUNK);
      }
      renderVirtualRows();
      updateSelectionUI();
    });
    ST.colHeadEls.forEach(function (el) {
      el.addEventListener("mousedown", function (e) {
        if (e.target.closest("[data-col-resizer]")) return;
        e.preventDefault();
        commitActiveCellValue();
        selectColumn(Number(el.getAttribute("data-col-head")));
      });
    });
    var corner = host.querySelector(".st-corner");
    if (corner) {
      corner.addEventListener("mousedown", function (e) {
        if (e.target.closest("[data-col-resizer], [data-row-resizer]")) return;
        e.preventDefault();
        commitActiveCellValue();
        selectAllCells();
      });
    }
    host.querySelectorAll("[data-col-resizer]").forEach(function (el) {
      el.addEventListener("mousedown", function (e) {
        if (e.detail >= 2) {
          e.preventDefault();
          e.stopPropagation();
          ST.colWidths[Number(el.getAttribute("data-col-resizer"))] =
            salesColumns[Number(el.getAttribute("data-col-resizer"))].width || 100;
          applyGridSizingUI();
          scheduleLedgerDraftSave();
          return;
        }
        e.preventDefault();
        e.stopPropagation();
        startResize("col", Number(el.getAttribute("data-col-resizer")), e.clientX, e.clientY);
      });
      el.addEventListener("dblclick", function (e) {
        e.preventDefault();
        e.stopPropagation();
        ST.colWidths[Number(el.getAttribute("data-col-resizer"))] =
          salesColumns[Number(el.getAttribute("data-col-resizer"))].width || 100;
        applyGridSizingUI();
        scheduleLedgerDraftSave();
      });
    });
    ST.inputEl.addEventListener("compositionstart", function () {
      ST.isComposing = true;
      if (!ST.editMode) {
        snapshotEditOrigin();
        ST.editMode = true;
        syncInputOverlay();
      }
    });
    ST.inputEl.addEventListener("compositionend", function (e) {
      ST.isComposing = false;
      setValue(ST.selectedCell.row, ST.selectedCell.col, e.target.value);
      syncInputOverlay();
    });
    ST.inputEl.addEventListener("input", function (e) {
      if (!ST.editMode) {
        snapshotEditOrigin();
        ST.editMode = true;
      }
      setValue(ST.selectedCell.row, ST.selectedCell.col, e.target.value);
      syncInputOverlay();
    });
    ST.inputEl.addEventListener("focus", function (e) {
      var len = e.target.value.length;
      if (ST.editMode) e.target.setSelectionRange(len, len);
      else e.target.setSelectionRange(0, len);
    });
    ST.inputEl.addEventListener("blur", function () {
      commitActiveCellValue();
    });
    ST.inputEl.addEventListener("mouseup", function (e) {
      if (!ST.editMode) {
        e.preventDefault();
        var len = e.target.value.length;
        e.target.setSelectionRange(0, len);
      }
    });
    ST.inputEl.addEventListener("select", function (e) {
      if (!ST.editMode) {
        var len = e.target.value.length;
        e.target.setSelectionRange(0, len);
      }
    });
    ST.inputEl.addEventListener("keydown", onInputKeyDown);

    syncEditOrigin();
    applyGridSizingUI();
    updateStatusBar();
    focusInput();
  }

  function moveInputToCell(r, c) {
    if (!ST.host || !ST.inputEl) return;
    ensureRowVisible(r);
    renderVirtualRows();
    var tbody = ST.host.querySelector("tbody");
    if (!tbody) return;
    var prevTd = ST.inputEl.closest("td");
    if (prevTd) {
      var disp = prevTd.querySelector(".st-display");
      var gh = prevTd.querySelector(".st-ghost");
      if (disp) {
        disp.style.display = "";
        disp.textContent = getValue(Number(prevTd.getAttribute("data-r")), Number(prevTd.getAttribute("data-c")));
      }
      if (gh) {
        gh.textContent = "";
        gh.style.display = "none";
      }
    }
    var td = tbody.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
    if (!td) return;
    var inner = td.querySelector(".st-cell-inner");
    var disp = td.querySelector(".st-display");
    var gh = td.querySelector(".st-ghost");
    if (disp) disp.style.display = "none";
    if (gh) gh.style.display = "none";
    inner.appendChild(ST.inputEl);
    ST.inputEl.value = getRawValue(r, c);
    syncInputOverlay();
  }

  function onCellDblClick(e) {
    var td = e.target.closest("td[data-r]");
    if (!td || !ST.host || !ST.host.contains(td)) return;
    commitActiveCellValue();
    var rowIndex = Number(td.getAttribute("data-r"));
    var colIndex = Number(td.getAttribute("data-c"));
    ST.selectedCell = { row: rowIndex, col: colIndex };
    ST.selectedKeys = [cellKey(rowIndex, colIndex)];
    snapshotEditOrigin();
    ST.editMode = true;
    ST.fillPreviewKeys = [];
    ST.fillSourceRect = null;
    ST.fillDirection = null;
    updateSelectionUI();
    moveInputToCell(rowIndex, colIndex);
    requestAnimationFrame(function () {
      if (!ST.inputEl) return;
      ST.inputEl.focus();
      var len = ST.inputEl.value.length;
      ST.inputEl.setSelectionRange(len, len);
    });
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }
  function escapeAttr(s) {
    return escapeHtml(s).replace(/"/g, "&quot;");
  }

  function onWindowMouseUp() {
    if (ST.resizing) {
      ST.resizing = null;
      scheduleLedgerDraftSave();
      return;
    }
    if (ST.fillDragging) {
      ST.fillDragging = false;
      applyFillPreview();
      return;
    }
    if (ST.dragging) {
      ST.dragging = false;
      moveInputToCell(ST.selectedCell.row, ST.selectedCell.col);
      syncEditOrigin();
      focusInput();
      return;
    }
    ST.dragging = false;
  }

  function onTbodyMouseDown(e) {
    var rowResizer = e.target.closest("[data-row-resizer]");
    if (rowResizer) {
      if (e.detail >= 2) {
        e.preventDefault();
        e.stopPropagation();
        ST.rowHeights[Number(rowResizer.getAttribute("data-row-resizer"))] = DEFAULT_ROW_HEIGHT;
        applyGridSizingUI();
        scheduleLedgerDraftSave();
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      startResize("row", Number(rowResizer.getAttribute("data-row-resizer")), e.clientX, e.clientY);
      return;
    }

    var rowHead = e.target.closest("[data-row-head]");
    if (rowHead) {
      e.preventDefault();
      commitActiveCellValue();
      selectRow(Number(rowHead.getAttribute("data-row-head")));
      return;
    }

    onCellMouseDown(e);
  }

  function onTbodyDblClick(e) {
    var rowResizer = e.target.closest("[data-row-resizer]");
    if (rowResizer) {
      e.preventDefault();
      e.stopPropagation();
      ST.rowHeights[Number(rowResizer.getAttribute("data-row-resizer"))] = DEFAULT_ROW_HEIGHT;
      applyGridSizingUI();
      scheduleLedgerDraftSave();
      return;
    }
    onCellDblClick(e);
  }

  function onCellMouseDown(e) {
    var td = e.target.closest("td[data-r]");
    if (!td || !td.getAttribute("data-r")) return;
    if (e.button !== 0) return;
    if (e.target.closest(".st-fill-handle")) {
      e.preventDefault();
      e.stopPropagation();
      ST.fillDragging = true;
      ST.dragging = false;
      ST.fillSourceRect = selectionBounds();
      ST.fillPreviewKeys = [];
      ST.fillDirection = null;
      updateSelectionUI();
      return;
    }

    commitActiveCellValue();
    e.preventDefault();
    var rowIndex = Number(td.getAttribute("data-r"));
    var colIndex = Number(td.getAttribute("data-c"));
    var pos = { row: rowIndex, col: colIndex };

    ST.dragging = true;
    ST.dragStart = pos;
    ST.selectedCell = pos;
    ST.editMode = false;
    ST.fillPreviewKeys = [];
    ST.fillSourceRect = null;
    ST.fillDirection = null;

    var key = cellKey(rowIndex, colIndex);
    var already = ST.selectedKeys.indexOf(key) >= 0;
    if (e.ctrlKey || e.metaKey) {
      ST.dragMode = already ? "remove" : "add";
      if (already) {
        ST.selectedKeys = ST.selectedKeys.filter(function (k) {
          return k !== key;
        });
      } else {
        ST.selectedKeys = Array.from(new Set(ST.selectedKeys.concat([key])));
      }
    } else {
      ST.dragMode = "replace";
      ST.selectedKeys = [key];
    }
    ST.dragBaseKeys = ST.selectedKeys.slice();

    updateSelectionUI();
    moveInputToCell(rowIndex, colIndex);
    syncEditOrigin();
    focusInput();

    requestAnimationFrame(function () {
      if (!ST.inputEl) return;
      var len = ST.inputEl.value.length;
      ST.inputEl.setSelectionRange(0, len);
    });
  }

  function getTdFromPoint(x, y) {
    var el = document.elementFromPoint(x, y);
    if (!el) return null;
    var td = el.closest("td[data-r]");
    if (!td || !ST.host || !ST.host.contains(td)) return null;
    return td;
  }

  function onDocMouseMove(e) {
    if (ST.resizing) {
      if (ST.resizing.kind === "col") {
        ST.colWidths[ST.resizing.index] = Math.max(
          MIN_COL_WIDTH,
          ST.resizing.startSize + (e.clientX - ST.resizing.startX)
        );
        applyColumnWidthUI(ST.resizing.index);
      } else {
        ST.rowHeights[ST.resizing.index] = Math.max(
          MIN_ROW_HEIGHT,
          ST.resizing.startSize + (e.clientY - ST.resizing.startY)
        );
        applyRowHeightUI(ST.resizing.index);
      }
      return;
    }

    if (ST.fillDragging) {
      var fillTd = getTdFromPoint(e.clientX, e.clientY);
      if (!fillTd) return;
      updateFillPreview(
        Number(fillTd.getAttribute("data-r")),
        Number(fillTd.getAttribute("data-c"))
      );
      return;
    }

    if (!ST.dragging || !ST.dragStart) return;
    var td = getTdFromPoint(e.clientX, e.clientY);
    if (!td) return;
    var rowIndex = Number(td.getAttribute("data-r"));
    var colIndex = Number(td.getAttribute("data-c"));
    var current = { row: rowIndex, col: colIndex };
    ST.selectedCell = current;
    ST.editMode = false;
    var rectKeys = keysFromPoints(ST.dragStart, current);
    if (ST.dragMode === "replace") ST.selectedKeys = rectKeys;
    else if (ST.dragMode === "add")
      ST.selectedKeys = Array.from(new Set(ST.dragBaseKeys.concat(rectKeys)));
    else
      ST.selectedKeys = ST.dragBaseKeys.filter(function (k) {
        return rectKeys.indexOf(k) < 0;
      });

    updateSelectionUI();
  }

  function onTableKeyDown(e) {
    if (ST.editMode) return;
    if (ST.inputEl && e.target === ST.inputEl) return;

    if (e.ctrlKey || e.metaKey) {
      var k = e.key.toLowerCase();
      if (k === "a") {
        e.preventDefault();
        selectAllCells();
        return;
      }
      if (k === "c") {
        e.preventDefault();
        copySelectionAsync();
        return;
      }
      if (k === "x") {
        e.preventDefault();
        cutSelectionAsync();
        return;
      }
      if (k === "v") {
        e.preventDefault();
        pasteFromClipboardAsync();
        return;
      }
      if (k === "z") {
        e.preventDefault();
        if (e.shiftKey) redo();
        else undo();
        return;
      }
      if (k === "y") {
        e.preventDefault();
        redo();
        return;
      }
      return;
    }

    if (e.altKey && e.key === "F2") {
      e.preventDefault();
      snapshotEditOrigin();
      ST.editMode = true;
      syncInputOverlay();
      focusInput();
      return;
    }

    if (e.key === "Delete") {
      e.preventDefault();
      clearSelectionCells();
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      moveByTab(e.shiftKey);
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      moveSelection(
        e.shiftKey ? ST.selectedCell.row - 1 : ST.selectedCell.row + 1,
        ST.selectedCell.col
      );
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      moveSelection(ST.selectedCell.row - 1, ST.selectedCell.col);
      return;
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      moveSelection(ST.selectedCell.row + 1, ST.selectedCell.col);
      return;
    }
    if (e.key === "ArrowLeft") {
      e.preventDefault();
      moveSelection(ST.selectedCell.row, ST.selectedCell.col - 1);
      return;
    }
    if (e.key === "ArrowRight") {
      e.preventDefault();
      moveSelection(ST.selectedCell.row, ST.selectedCell.col + 1);
      return;
    }
    if (e.key === "F2") {
      e.preventDefault();
      snapshotEditOrigin();
      ST.editMode = true;
      syncInputOverlay();
      focusInput();
      return;
    }
    if (e.key === "Escape") {
      e.preventDefault();
      var o = ST.editOrigin;
      applyRowsChange(function (current) {
        var k = salesColumns[o.col].key;
        current[o.row] = Object.assign({}, current[o.row], (function () { var x = {}; x[k] = o.value; return x; })());
        return current;
      });
      moveSelection(o.row, o.col);
    }
  }

  function copySelectionAsync() {
    var text = getSelectionText();
    if (!text) return;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(function () {});
    }
  }

  function pasteFromClipboardAsync() {
    if (navigator.clipboard && navigator.clipboard.readText) {
      navigator.clipboard.readText().then(pasteTextIntoSelection).catch(function () {});
    }
  }

  function onCopy(e) {
    var text = getSelectionText();
    if (!text) return;
    e.preventDefault();
    e.clipboardData.setData("text/plain", text);
  }

  function onPaste(e) {
    e.preventDefault();
    var text = e.clipboardData.getData("text/plain") || e.clipboardData.getData("text");
    pasteTextIntoSelection(text);
  }

  function onInputKeyDown(e) {
    var rowIndex = ST.selectedCell.row;
    var colIndex = ST.selectedCell.col;

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "c" && !ST.editMode) {
      e.preventDefault();
      copySelectionAsync();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "a" && !ST.editMode) {
      e.preventDefault();
      selectAllCells();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "x" && !ST.editMode) {
      e.preventDefault();
      cutSelectionAsync();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "v" && !ST.editMode) {
      e.preventDefault();
      pasteFromClipboardAsync();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
      e.preventDefault();
      if (e.shiftKey) redo();
      else undo();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "y") {
      e.preventDefault();
      redo();
      return;
    }
    if (e.key === "Escape") {
      e.preventDefault();
      var o = ST.editOrigin;
      applyRowsChange(function (current) {
        var k = salesColumns[o.col].key;
        current[o.row] = Object.assign({}, current[o.row], (function () { var x = {}; x[k] = o.value; return x; })());
        return current;
      });
      moveSelection(o.row, o.col);
      return;
    }
    if (ST.isComposing || e.isComposing) return;

    var input = e.target;
    var start = input.selectionStart != null ? input.selectionStart : 0;
    var end = input.selectionEnd != null ? input.selectionEnd : 0;
    var len = input.value.length;

    if (!ST.editMode) {
      if (e.key === "Delete") {
        e.preventDefault();
        clearSelectionCells();
        return;
      }
      if (e.key === "Enter") {
        e.preventDefault();
        moveSelection(e.shiftKey ? rowIndex - 1 : rowIndex + 1, colIndex);
        return;
      }
      if (e.key === "Tab") {
        e.preventDefault();
        moveByTab(e.shiftKey);
        return;
      }
      if (e.key === "ArrowUp") {
        e.preventDefault();
        moveSelection(rowIndex - 1, colIndex);
        return;
      }
      if (e.key === "ArrowDown") {
        e.preventDefault();
        moveSelection(rowIndex + 1, colIndex);
        return;
      }
      if (e.key === "ArrowLeft") {
        e.preventDefault();
        moveSelection(rowIndex, colIndex - 1);
        return;
      }
      if (e.key === "ArrowRight") {
        e.preventDefault();
        moveSelection(rowIndex, colIndex + 1);
        return;
      }
      if (e.key === "F2") {
        e.preventDefault();
        snapshotEditOrigin();
        ST.editMode = true;
        syncInputOverlay();
        focusInput();
      }
      return;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      moveSelection(e.shiftKey ? rowIndex - 1 : rowIndex + 1, colIndex);
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      moveByTab(e.shiftKey);
      return;
    }
    if (e.key === "ArrowUp") {
      e.preventDefault();
      moveSelection(rowIndex - 1, colIndex);
      return;
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      moveSelection(rowIndex + 1, colIndex);
      return;
    }
    if (e.key === "ArrowLeft" && start === 0 && end === 0) {
      e.preventDefault();
      moveSelection(rowIndex, colIndex - 1);
      return;
    }
    if (e.key === "ArrowRight" && start === len && end === len) {
      e.preventDefault();
      moveSelection(rowIndex, colIndex + 1);
      return;
    }
  }

  function attachSalesTableWhenNeeded(container) {
    if (!ST.tableBuilt || ST.host !== container) {
      ST.tableBuilt = false;
      buildSalesTable(container);
    } else {
      updateSelectionUI();
      moveInputToCell(ST.selectedCell.row, ST.selectedCell.col);
      syncEditOrigin();
      focusInput();
    }
  }

  function getWorkInfoValue(path) {
    var cur = workState.info || {};
    String(path).split(".").forEach(function (part) {
      cur = cur && cur[part] != null ? cur[part] : "";
    });
    return cur == null ? "" : String(cur);
  }

  function setWorkInfoValue(path, value) {
    var parts = String(path).split(".");
    var cur = workState.info || {};
    workState.info = cur;
    for (var i = 0; i < parts.length - 1; i++) {
      if (!cur[parts[i]] || typeof cur[parts[i]] !== "object") {
        cur[parts[i]] = {};
      }
      cur = cur[parts[i]];
    }
    cur[parts[parts.length - 1]] = value;
  }

  function formatBusinessNo(value) {
    var digits = String(value || "").replace(/\D/g, "");
    if (digits.length !== 10) return value == null ? "" : String(value);
    return digits.slice(0, 3) + "-" + digits.slice(3, 5) + "-" + digits.slice(5);
  }

  function formatDateLabel(value) {
    if (!value) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(value))) {
      var parts = String(value).split("-");
      return Number(parts[0]) + "년 " + Number(parts[1]) + "월 " + Number(parts[2]) + "일";
    }
    return String(value);
  }

  function getStatementTotals(items) {
    var totals = {
      qty: 0,
      supply: 0,
      tax: 0,
      grand: 0,
    };

    items.forEach(function (item) {
      totals.qty += parseCalcNumber(item.qty) || 0;
      totals.supply += parseCalcNumber(item.supply) || 0;
      totals.tax += parseCalcNumber(item.tax) || 0;
    });

    totals.grand = totals.supply + totals.tax;
    return totals;
  }

  function formatStatementItemDate(value) {
    var text = String(value || "").trim();
    var monthDay = text.match(/^(\d{1,2})월\s*(\d{1,2})일$/);
    if (monthDay) {
      return String(monthDay[1]).padStart(2, "0") + "월 " + String(monthDay[2]).padStart(2, "0") + "일";
    }
    return text;
  }

  function statementCell(text, className) {
    return '<td' + (className ? ' class="' + className + '"' : "") + '>' + escapeHtml(text == null ? "" : String(text)) + "</td>";
  }

  function statementHead(text, className, attrs) {
    return '<th' + (className ? ' class="' + className + '"' : "") + (attrs ? " " + attrs : "") + ">" + escapeHtml(text == null ? "" : String(text)) + "</th>";
  }

  function renderStatementItemsRows(items, totalRows) {
    var rows = "";
    for (var i = 0; i < totalRows; i++) {
      var item = items[i] || {};
      var nameText = item.name || "";
      var codeText = item.code || "";
      var noteText = item.note || "";
      var nameClass = "item-name";
      var codeClass = "item-code";
      var noteClass = "item-note";
      if (String(nameText).length > 38) nameClass += " tighter";
      else if (String(nameText).length > 24) nameClass += " tight";
      if (String(codeText).length > 14) codeClass += " tighter";
      else if (String(codeText).length > 10) codeClass += " tight";
      if (String(noteText).length > 14) noteClass += " tighter";
      else if (String(noteText).length > 10) noteClass += " tight";
      rows += "<tr>";
      rows += statementCell(formatStatementItemDate(item.date || ""), "center");
      rows += statementCell(nameText, nameClass);
      rows += statementCell(codeText, codeClass);
      rows += statementCell(noteText, noteClass);
      rows += statementCell(formatDisplayNumber(item.qty), "right");
      rows += statementCell(formatDisplayNumber(item.price), "right");
      rows += statementCell(formatDisplayNumber(item.supply), "right");
      rows += statementCell(formatDisplayNumber(item.tax), "right");
      rows += "</tr>";
    }
    return rows;
  }

  function renderStatementCopy(copyClass, tagLabel, items, totals) {
    var supplier = workState.info && workState.info.supplier ? workState.info.supplier : {};
    var receiver = workState.info && workState.info.receiver ? workState.info.receiver : {};
    var docDate = formatDateLabel(getWorkInfoValue("date")) || "";
    var totalRows = 14;

    return (
      '<section class="statement-copy ' + copyClass + '">' +
        '<div class="statement-copy-titlebar">' +
          '<div class="statement-copy-date">' + escapeHtml(docDate) + '</div>' +
          '<div class="statement-copy-title">거 래 명 세 표</div>' +
          '<div class="statement-copy-tag">(' + escapeHtml(tagLabel) + ')</div>' +
        '</div>' +
        '<table class="statement-table statement-header-table">' +
          '<colgroup>' +
            '<col style="width:4.5%" />' +
            '<col style="width:7%" />' +
            '<col style="width:22%" />' +
            '<col style="width:6%" />' +
            '<col style="width:10.5%" />' +
            '<col style="width:4.5%" />' +
            '<col style="width:7%" />' +
            '<col style="width:22%" />' +
            '<col style="width:6%" />' +
            '<col style="width:10.5%" />' +
          '</colgroup>' +
          '<tbody>' +
            '<tr>' +
              statementHead("등록번호", "center", 'colspan="2"') +
              '<td colspan="3" class="center">' + escapeHtml(formatBusinessNo(supplier.businessNo || "")) + '</td>' +
              statementHead("등록번호", "center", 'colspan="2"') +
              '<td colspan="3" class="center">' + escapeHtml(formatBusinessNo(receiver.businessNo || "")) + '</td>' +
            '</tr>' +
            '<tr>' +
              statementHead("공급자", "vhead", 'rowspan="3"') +
              statementHead("상호", "center") +
              statementCell(supplier.company || "") +
              statementHead("성명", "center") +
              statementCell(supplier.ceo || "", "center") +
              statementHead("공급받는자", "vhead", 'rowspan="3"') +
              statementHead("상호", "center") +
              statementCell(receiver.company || "", "") +
              statementHead("성명", "center") +
              statementCell(receiver.ceo || "", "center") +
            '</tr>' +
            '<tr>' +
              statementHead("주소", "center") +
              '<td colspan="3" class="center">' + escapeHtml(supplier.address || "") + '</td>' +
              statementHead("주소", "center") +
              '<td colspan="3" class="center">' + escapeHtml(receiver.address || "") + '</td>' +
            '</tr>' +
            '<tr>' +
              statementHead("업태", "center") +
              statementCell(supplier.businessType || "", "center") +
              statementHead("종목", "center") +
              statementCell(supplier.businessItem || "", "center") +
              statementHead("업태", "center") +
              statementCell(receiver.businessType || "", "center") +
              statementHead("종목", "center") +
              statementCell(receiver.businessItem || "", "center") +
            '</tr>' +
          '</tbody>' +
        '</table>' +
        '<table class="statement-table statement-items-table">' +
          '<colgroup>' +
            '<col style="width:10%" />' +
            '<col style="width:31%" />' +
            '<col style="width:10%" />' +
            '<col style="width:10%" />' +
            '<col style="width:8%" />' +
            '<col style="width:10%" />' +
            '<col style="width:13%" />' +
            '<col style="width:8%" />' +
          '</colgroup>' +
          '<thead><tr>' +
            statementHead("일자", "center") +
            statementHead("품명", "center") +
            statementHead("CODE", "center") +
            statementHead("비고", "center") +
            statementHead("수량", "center") +
            statementHead("단가", "center") +
            statementHead("금액", "center") +
            statementHead("세액", "center") +
          '</tr></thead>' +
          '<tbody>' + renderStatementItemsRows(items, totalRows) + '</tbody>' +
        '</table>' +
        '<table class="statement-table statement-footer-table">' +
          '<colgroup>' +
            '<col style="width:12%" />' +
            '<col style="width:23%" />' +
            '<col style="width:8%" />' +
            '<col style="width:12%" />' +
            '<col style="width:8%" />' +
            '<col style="width:14%" />' +
            '<col style="width:10%" />' +
            '<col style="width:13%" />' +
          '</colgroup>' +
          '<tbody><tr>' +
            statementHead("공급가액", "center") +
            statementCell(formatDisplayNumber(totals.supply), "right") +
            statementHead("세액", "center") +
            statementCell(formatDisplayNumber(totals.tax), "right") +
            statementHead("합계", "center") +
            statementCell(formatDisplayNumber(totals.grand), "right") +
            statementHead("인수자", "center") +
            statementCell("", "center") +
          '</tr></tbody>' +
        '</table>' +
      '</section>'
    );
  }

  function renderStatementTab() {
    var items = (workState.items || []).filter(function (item) {
      if (!item) return false;
      return ["date", "code", "name", "qty", "price", "supply", "tax", "note"].some(function (key) {
        return item[key] != null && String(item[key]).trim() !== "";
      });
    });
    var totals = getStatementTotals(items);

    return (
      '<div class="statement-print-root">' +
        '<div class="statement-toolbar">' +
          '<button type="button" class="soft-btn" id="btn-print-statement">' + icon("sheet") + '인쇄하기</button>' +
        '</div>' +
        '<div class="statement-page">' +
          '<div class="statement-stack">' +
            renderStatementCopy("blue", "공급받는자 보관용", items, totals) +
            renderStatementCopy("red", "공급자 보관용", items, totals) +
          '</div>' +
        '</div>' +
      '</div>'
    );
  }

  function renderWorkTab() {
    var items = workState.items || [];
    var totalRows = Math.max(items.length, 12);
    var itemsRows = "";

    for (var i = 0; i < totalRows; i++) {
      var it = items[i] || {};
      var vDate = it.date != null ? String(it.date) : "";
      var vCode = it.code != null ? String(it.code) : "";
      var vName = it.name != null ? String(it.name) : "";
      var vQty = it.qty != null ? String(it.qty) : "";
      var vPrice = it.price != null ? String(it.price) : "";
      var vSupply = it.supply != null ? String(it.supply) : "";
      var vTax = it.tax != null ? String(it.tax) : "";
      var vNote = it.note != null ? String(it.note) : "";

      itemsRows += "<tr>";
      itemsRows +=
        '<td><input class="cell-input" type="text" data-work-row="' +
        i +
        '" data-work-field="date" placeholder="일자" value="' +
        escapeAttr(vDate) +
        '" /></td>';
      itemsRows +=
        '<td><input class="cell-input" type="text" data-work-row="' +
        i +
        '" data-work-field="code" placeholder="코드" value="' +
        escapeAttr(vCode) +
        '" /></td>';
      itemsRows +=
        '<td><input class="cell-input" type="text" data-work-row="' +
        i +
        '" data-work-field="name" placeholder="품명" value="' +
        escapeAttr(vName) +
        '" /></td>';
      itemsRows +=
        '<td><input class="cell-input" type="text" inputmode="numeric" data-work-row="' +
        i +
        '" data-work-field="qty" placeholder="수량" value="' +
        escapeAttr(vQty) +
        '" /></td>';
      itemsRows +=
        '<td><input class="cell-input" type="text" inputmode="numeric" data-work-row="' +
        i +
        '" data-work-field="price" placeholder="단가" value="' +
        escapeAttr(vPrice) +
        '" /></td>';
      itemsRows +=
        '<td><input class="cell-input" type="text" inputmode="numeric" data-work-row="' +
        i +
        '" data-work-field="supply" placeholder="공급가액" value="' +
        escapeAttr(vSupply) +
        '" /></td>';
      itemsRows +=
        '<td><input class="cell-input" type="text" inputmode="numeric" data-work-row="' +
        i +
        '" data-work-field="tax" placeholder="세액" value="' +
        escapeAttr(vTax) +
        '" /></td>';
      itemsRows +=
        '<td><input class="cell-input" type="text" data-work-row="' +
        i +
        '" data-work-field="note" placeholder="비고" value="' +
        escapeAttr(vNote) +
        '" /></td>';
      itemsRows += "</tr>";
    }

    return (
      '<div class="workdoc">' +
        '<div class="workdoc-grid2">' +
          '<div class="workdoc-card">' +
            '<div class="workdoc-head">' +
              '<div class="workdoc-title">' + icon("building") + ' 공급자</div>' +
              '<div class="sub">필수항목만 먼저 UI 구성</div>' +
            "</div>" +
            '<div class="workdoc-form">' +
              '<div class="workdoc-label">상호</div><input class="workdoc-input" data-work-info="supplier.company" value="' + escapeAttr(getWorkInfoValue("supplier.company")) + '" />' +
              '<div class="workdoc-label">대표자</div><input class="workdoc-input" data-work-info="supplier.ceo" value="' + escapeAttr(getWorkInfoValue("supplier.ceo")) + '" />' +
              '<div class="workdoc-label">사업자등록번호</div><input class="workdoc-input" data-work-info="supplier.businessNo" placeholder="000-00-00000" value="' + escapeAttr(getWorkInfoValue("supplier.businessNo")) + '" />' +
              '<div class="workdoc-label">주소</div><input class="workdoc-input" data-work-info="supplier.address" value="' + escapeAttr(getWorkInfoValue("supplier.address")) + '" />' +
              '<div class="workdoc-label">업태</div><input class="workdoc-input" data-work-info="supplier.businessType" value="' + escapeAttr(getWorkInfoValue("supplier.businessType")) + '" />' +
              '<div class="workdoc-label">종목</div><input class="workdoc-input" data-work-info="supplier.businessItem" value="' + escapeAttr(getWorkInfoValue("supplier.businessItem")) + '" />' +
            "</div>" +
          "</div>" +
          '<div class="workdoc-card">' +
            '<div class="workdoc-head">' +
              '<div class="workdoc-title">' + icon("sheet") + ' 공급받는자</div>' +
              '<div class="sub">거래처 정보</div>' +
            "</div>" +
            '<div class="workdoc-form">' +
              '<div class="workdoc-label">상호</div><input class="workdoc-input" data-work-info="receiver.company" value="' + escapeAttr(getWorkInfoValue("receiver.company")) + '" />' +
              '<div class="workdoc-label">대표자</div><input class="workdoc-input" data-work-info="receiver.ceo" value="' + escapeAttr(getWorkInfoValue("receiver.ceo")) + '" />' +
              '<div class="workdoc-label">사업자등록번호</div><input class="workdoc-input" data-work-info="receiver.businessNo" placeholder="000-00-00000" value="' + escapeAttr(getWorkInfoValue("receiver.businessNo")) + '" />' +
              '<div class="workdoc-label">주소</div><input class="workdoc-input" data-work-info="receiver.address" value="' + escapeAttr(getWorkInfoValue("receiver.address")) + '" />' +
              '<div class="workdoc-label">업태</div><input class="workdoc-input" data-work-info="receiver.businessType" value="' + escapeAttr(getWorkInfoValue("receiver.businessType")) + '" />' +
              '<div class="workdoc-label">종목</div><input class="workdoc-input" data-work-info="receiver.businessItem" value="' + escapeAttr(getWorkInfoValue("receiver.businessItem")) + '" />' +
            "</div>" +
          "</div>" +
        "</div>" +
        '<div class="workdoc-card">' +
          '<div class="workdoc-date-line">' +
            '<div class="workdoc-title">' + icon("clipboard") + ' 작성일자</div>' +
            '<input type="date" class="workdoc-input" data-work-info="date" value="' + escapeAttr(getWorkInfoValue("date")) + '" />' +
            '<div class="sub">한 줄 영역</div>' +
          "</div>" +
        "</div>" +
        '<div class="workdoc-card">' +
          '<div class="workdoc-head">' +
            '<div class="workdoc-title">' + icon("scroll") + ' 품목</div>' +
            '<div class="sub">표 구조만 먼저 추가 (엔진은 추후)</div>' +
          "</div>" +
          '<div class="items-wrap">' +
            '<table class="items">' +
              "<thead><tr>" +
                '<th class="items-col-date">일자</th>' +
                '<th class="items-col-code">코드</th>' +
                '<th class="items-col-name">품명</th>' +
                '<th class="items-col-qty">수량</th>' +
                '<th class="items-col-price">단가</th>' +
                '<th class="items-col-supply">공급가액</th>' +
                '<th class="items-col-tax">세액</th>' +
                '<th class="items-col-note">비고</th>' +
              "</tr></thead>" +
              "<tbody>" + itemsRows + "</tbody>" +
            "</table>" +
          "</div>" +
        "</div>" +
      "</div>"
    );
  }

  function sendSelectedStatusRowsToWork() {
    // 1) 매출현황(ST)에서 선택된 셀들 -> row index 추출
    var keys = ST.selectedKeys || [];
    if (!keys.length) return;

    var rowSet = new Set();
    keys.forEach(function (key) {
      var p = String(key).split(":");
      if (p.length >= 2) rowSet.add(Number(p[0]));
    });

    // 2) 중복 row 제거 -> 정렬
    var rows = Array.from(rowSet).sort(function (a, b) {
      return a - b;
    });
    if (!rows.length) return;

    // 3) 최대 10줄 제한
    if (rows.length > 10) {
      alert("선택 행은 최대 10줄까지만 가능합니다.");
      return;
    }

    // 4) 선택 row들의 전체 데이터 복사(원본 ST.rows 수정 X)
    //    -> 거래작업 품목표 컬럼에 매핑
    var items = rows.map(function (r) {
      var row = ST.rows[r] || emptyRow();
      return {
        date: row.date,
        code: row.code,
        name: row.name,
        qty: row.qty,
        price: row.price,
        supply: row.amount,
        tax: "", // 매출현황 엔진에는 세액 필드가 없으므로 비움
        note: [row.note1, row.note2].filter(function (x) {
          return x != null && String(x).trim() !== "";
        }).join(" / "),
      };
    });

    workState.items = items;

    // 5) 거래작업 탭으로 이동 + 복사본 렌더링
    //    (동시에 로컬 백업/Firebase에도 “복사본” 저장)
    state.subTab = "work";

    ensureDraftId();
    workState.loadedFromFirebase = true;
    scheduleLedgerDraftSave();

    render();
  }

  function ensureWorkLoadedFromFirebase() {
    return ensureLedgerLoadedFromFirebase().then(function () {
      workState.loadedFromFirebase = true;
    });
  }

  function saveWorkDraftToFirebase() {
    saveLedgerDraftToFirebase();
  }

  /* ========== App render ========== */
  function render() {
    if (!state.entered) {
      app.innerHTML =
        '<div class="login-wrap" id="login-root" tabindex="0">' +
        '<div class="panel">' +
        '<div class="mb-4" style="display:flex;align-items:center;gap:12px">' +
        '<div class="icon-box">' +
        icons.lock +
        "</div>" +
        "<div>" +
        '<div class="title-lg">로그인</div>' +
        '<div class="sub">엔터만 눌러도 바로 넘어가도록 설정됨</div>' +
        "</div></div>" +
        '<input type="text" class="field" id="login-id" placeholder="아이디" />' +
        '<input type="password" class="field mt-3" id="login-pw" placeholder="비밀번호" />' +
        '<button type="button" class="primary-btn mt-4" id="login-btn">' +
        "들어가기 " +
        icons.arrowRight +
        "</button></div></div>";

      var lr = document.getElementById("login-root");
      function doEnter() {
        state.entered = true;
        render();
      }
      lr.addEventListener("keydown", function (e) {
        if (e.key === "Enter") doEnter();
      });
      document.getElementById("login-btn").addEventListener("click", doEnter);
      lr.focus();
      return;
    }

    if (!state.role) {
      var h = '<div class="role-grid">';
      roleCards.forEach(function (card) {
        h +=
          '<button type="button" class="role-card" data-role="' +
          card.key +
          '">' +
          '<div class="icon-box-sm" style="display:flex;align-items:center;justify-content:center">' +
          icon(card.icon) +
          "</div>" +
          '<div class="title-xl">' +
          card.title +
          "</div></button>";
      });
      h += "</div>";
      app.innerHTML = h;
      app.querySelectorAll(".role-card").forEach(function (btn) {
        btn.addEventListener("click", function () {
          state.role = btn.getAttribute("data-role");
          render();
        });
      });
      return;
    }

    if (state.role !== "accounting") {
      app.innerHTML =
        '<div class="shell-bg prep">' + state.role + " 화면 준비중</div>";
      return;
    }

    var top =
      '<div class="top-bar">' +
      '<button type="button" class="soft-btn" id="btn-back">' +
      icons.arrowLeft +
      " 이전</button>" +
      '<button type="button" class="soft-btn" id="btn-out">로그아웃</button></div>';

    app.className = "";

    if (!state.mainTab) {
      app.innerHTML =
        top +
        '<div class="content"><div class="main-pick">' +
        '<button type="button" class="main-card" data-main="sales">' +
        icon("scroll") +
        '<div class="mb-3"></div><div class="font-semibold">매출</div></button>' +
        '<button type="button" class="main-card" data-main="purchase">' +
        icon("folder") +
        '<div class="mb-3"></div><div class="font-semibold">매입</div></button>' +
        '<button type="button" class="main-card" data-main="closing">' +
        icon("sheet") +
        '<div class="mb-3"></div><div class="font-semibold">마감</div></button>' +
        "</div></div>";

      wireTopBar();
      app.querySelectorAll(".main-card").forEach(function (b) {
        b.addEventListener("click", function () {
          state.mainTab = b.getAttribute("data-main");
          render();
        });
      });
      return;
    }

    if (state.mainTab === "sales") {
      var tabs = '<div class="tabs">';
      salesTabs.forEach(function (t) {
        tabs +=
          '<button type="button" class="sub-tab' +
          (state.subTab === t.key ? " active" : "") +
          '" data-sub="' +
          t.key +
          '">' +
          icon(t.icon) +
          t.label +
          "</button>";
      });
      tabs += "</div>";

      var body;
      if (state.subTab === "status") {
        body =
          '<div class="toolbar-card" style="margin-bottom:12px">' +
          '<label>필터:</label>' +
          '<select id="status-filter-col"><option value="all">전체</option>' +
          salesColumns.map(function (col) { return '<option value="' + col.key + '">' + col.label + '</option>'; }).join("") +
          '</select>' +
          '<input type="text" class="field" id="status-filter-keyword" style="width:180px;height:36px" placeholder="검색어" />' +
          '<button type="button" class="soft-btn" id="btn-filter-apply">적용</button>' +
          '<button type="button" class="soft-btn" id="btn-filter-clear">해제</button>' +
          '<label style="margin-left:12px">정렬:</label>' +
          '<select id="status-sort-col">' +
          salesColumns.map(function (col) { return '<option value="' + col.key + '">' + col.label + '</option>'; }).join("") +
          '</select>' +
          '<button type="button" class="soft-btn" id="btn-sort-asc">오름차순</button>' +
          '<button type="button" class="soft-btn" id="btn-sort-desc">내림차순</button>' +
          '<button type="button" class="soft-btn" id="btn-sort-reset">입력순</button>' +
          '</div>' +
          '<div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:12px;">' +
          '<button type="button" class="soft-btn" id="btn-send-to-work">' +
          icon("arrowRight") +
          "거래작업으로 보내기" +
          "</button>" +
          "</div>" +
          '<div id="sales-table-host"></div>';
      } else if (state.subTab === "work") {
        body = renderWorkTab();
      } else if (state.subTab === "statement") {
        body = renderStatementTab();
      } else {
        var lab = salesTabs.find(function (x) {
          return x.key === state.subTab;
        });
        body =
          '<div class="empty-page">' +
          (lab ? lab.label : state.subTab) +
          " 페이지 (비워둠)</div>";
      }

      app.innerHTML = top + '<div class="content">' + tabs + body + "</div>";

      wireTopBar();

      app.querySelectorAll(".sub-tab").forEach(function (b) {
        b.addEventListener("click", function () {
          state.subTab = b.getAttribute("data-sub");
          render();
        });
      });

      if (state.subTab === "status") {
        var filterCol = document.getElementById("status-filter-col");
        var filterKeyword = document.getElementById("status-filter-keyword");
        var sortCol = document.getElementById("status-sort-col");
        if (filterCol) filterCol.value = ST.filter.colKey || "all";
        if (filterKeyword) filterKeyword.value = ST.filter.keyword || "";
        if (sortCol) sortCol.value = ST.filter.colKey && ST.filter.colKey !== "all" ? ST.filter.colKey : "date";
        attachSalesTableWhenNeeded(document.getElementById("sales-table-host"));
        ensureLedgerLoadedFromFirebase().then(function () {
          if (state.subTab === "status") {
            refreshGridValues();
            updateSelectionUI();
            applyFilterUI();
          }
        });
        var applyBtn = document.getElementById("btn-filter-apply");
        var clearBtn = document.getElementById("btn-filter-clear");
        var sortAscBtn = document.getElementById("btn-sort-asc");
        var sortDescBtn = document.getElementById("btn-sort-desc");
        var sortResetBtn = document.getElementById("btn-sort-reset");
        if (applyBtn) {
          applyBtn.addEventListener("click", function () {
            ST.filter.colKey = filterCol ? filterCol.value : "all";
            ST.filter.keyword = filterKeyword ? filterKeyword.value : "";
            applyFilterUI();
          });
        }
        if (clearBtn) {
          clearBtn.addEventListener("click", function () {
            ST.filter.colKey = "all";
            ST.filter.keyword = "";
            if (filterCol) filterCol.value = "all";
            if (filterKeyword) filterKeyword.value = "";
            applyFilterUI();
          });
        }
        if (sortAscBtn) {
          sortAscBtn.addEventListener("click", function () {
            sortRowsByColumn(sortCol ? sortCol.value : "date", "asc");
            applyFilterUI();
          });
        }
        if (sortDescBtn) {
          sortDescBtn.addEventListener("click", function () {
            sortRowsByColumn(sortCol ? sortCol.value : "date", "desc");
            applyFilterUI();
          });
        }
        if (sortResetBtn) {
          sortResetBtn.addEventListener("click", function () {
            resetSortOrder();
            applyFilterUI();
          });
        }
        if (filterKeyword) {
          filterKeyword.addEventListener("keydown", function (e) {
            if (e.key === "Enter") {
              ST.filter.colKey = filterCol ? filterCol.value : "all";
              ST.filter.keyword = filterKeyword.value;
              applyFilterUI();
            }
          });
        }
        var sendBtn = document.getElementById("btn-send-to-work");
        if (sendBtn) {
          sendBtn.addEventListener("click", function () {
            sendSelectedStatusRowsToWork();
          });
        }
      } else if (state.subTab === "work") {
        // Firestore에서 복사본을 자동 불러오기
        var wasLoaded = workState.loadedFromFirebase;
        ensureWorkLoadedFromFirebase().then(function () {
          // 처음 로딩인 경우에만 한 번 더 렌더링(값 반영)
          if (!wasLoaded) render();
        });

        // 거래작업 입력 변경 시 로컬 백업/Firebase 자동 저장
        if (workState.saveTimer) {
          clearTimeout(workState.saveTimer);
          workState.saveTimer = null;
        }
        app.querySelectorAll('input[data-work-info]').forEach(function (inp) {
          inp.addEventListener("input", function () {
            var path = inp.getAttribute("data-work-info");
            if (!path) return;
            setWorkInfoValue(path, inp.value);
            scheduleLedgerDraftSave();
          });
        });
        app.querySelectorAll('input[data-work-row]').forEach(function (inp) {
          inp.addEventListener("input", function () {
            var r = Number(inp.getAttribute("data-work-row"));
            var f = inp.getAttribute("data-work-field");
            if (isNaN(r) || !f) return;

            if (!workState.items[r]) workState.items[r] = emptyRow();
            if (f === "note") {
              workState.items[r].note = inp.value;
            } else {
              workState.items[r][f] = inp.value;
            }

            scheduleLedgerDraftSave();
          });
        });
      } else if (state.subTab === "statement") {
        var wasStatementLoaded = workState.loadedFromFirebase;
        ensureWorkLoadedFromFirebase().then(function () {
          if (!wasStatementLoaded && state.subTab === "statement") render();
        });
        var printBtn = document.getElementById("btn-print-statement");
        if (printBtn) {
          printBtn.addEventListener("click", function () {
            window.print();
          });
        }
      }
      return;
    }

    app.innerHTML =
      top +
      '<div class="content"><div class="empty-page">' +
      (state.mainTab === "purchase" ? "매입" : "마감") +
      "</div></div>";
    wireTopBar();
  }

  function wireTopBar() {
    var back = document.getElementById("btn-back");
    var out = document.getElementById("btn-out");
    if (back) {
      back.addEventListener("click", function () {
        if (state.mainTab) {
          state.mainTab = null;
          ST.tableBuilt = false;
          ST.host = null;
          ST.inputEl = null;
          ST.scrollEl = null;
        } else state.role = null;
        render();
      });
    }
    if (out) {
      out.addEventListener("click", function () {
        state.entered = false;
        state.role = null;
        state.mainTab = null;
        state.subTab = "manage";
        ST.tableBuilt = false;
        ST.host = null;
        ST.inputEl = null;
        ST.scrollEl = null;
        render();
      });
    }
  }

  bindFirestoreRetryListeners();
  loadLocalLedgerBackup();
  render();
})();
