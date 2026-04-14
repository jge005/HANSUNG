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
  var purchaseTabs = [
    { key: "manage", label: "매입관리", icon: "sheet" },
    { key: "status", label: "매입현황", icon: "clipboard" },
    { key: "price", label: "매입단가", icon: "tags" },
    { key: "client", label: "업체리스트", icon: "building" },
  ];
  var closingTabs = [
    { key: "attendance", label: "근무(급여)", icon: "sheet" },
    { key: "salesclose", label: "매출마감", icon: "scroll" },
    { key: "purchaseclose", label: "매입마감", icon: "folder" },
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
    closingSubTab: "attendance",
  };

  var DRAFT_ID_KEY = "ledgerDraftId";
  var LEGACY_DRAFT_ID_KEY = "workDraftId";
  var LOCAL_BACKUP_KEY = "ledgerDraftBackup";
  var SHARED_DRAFT_ID = "shared-ledger-main";

  var ledgerState = {
    draftId: null,
    legacyDraftId: null,
    loadedFromFirebase: false,
    initialLoadFinished: false,
    loadingPromise: null,
    saveTimer: null,
    localSaveTimer: null,
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
  var clientState = {
    rows: [],
  };
  var priceState = {
    rows: [],
    activeClient: "",
    sidebarScrollTop: 0,
  };
  var manageState = {
    startMonth: "",
    endMonth: "",
    client: "",
  };
  var closingState = {
    attendanceMonth: "",
    attendanceView: "employee",
    employeeViewMode: "entry",
    employeeRoster: [],
    employeeRowsByMonth: {},
    outsourceVendor: "leaders",
    outsourceRosterByVendor: {},
    outsourceRowsByKey: {},
    outsourceWarningsByKey: {},
    outsourceMetaByKey: {},
    salesCloseRowsByMonth: {},
    purchaseCloseRowsByMonth: {},
    attendanceSearch: "",
  };
  var activeLedgerKind = "sales";
  var salesLedgerBundle = null;
  var purchaseLedgerBundle = null;
  var appToastTimer = null;
  var workGridFields = ["date", "code", "name", "qty", "price", "supply", "tax", "note"];
  var clientGridFields = ["supplierMode", "company", "businessNo", "ceoName", "address", "businessType", "businessItem"];
  var clientDataFields = ["supplierMode", "company", "businessNo", "ceoName", "address", "businessType", "businessItem"];
  var priceGridFields = ["client", "code", "price", "name"];
  var priceEditorFields = ["code", "price", "name"];
  var closingAttendanceDayFields = [];
  for (var closingDayIndex = 1; closingDayIndex <= 31; closingDayIndex++) {
    closingAttendanceDayFields.push("d" + closingDayIndex);
  }
  var closingEmployeeNames = [
    "김은숙",
    "이지연",
    "정인숙",
    "안현수",
    "최성근 사원",
    "김규헌 차장",
    "고광운 과장",
    "박재현 대리",
    "김현규 대리",
    "조영빈 대리",
    "김주헌 대리",
    "장진우 대리",
    "최성규 사원",
    "김용원 사원",
    "김미은 주임",
    "김인선 사원",
    "박지안 사원",
    "이수안 사원",
    "김애정 사원"
  ];
  var closingEmployeeMarkers = ["야근", "특근", "지각", "조퇴", "결근"];
  var closingOutsourceVendors = [
    { value: "gonggam", label: "공감인/우리인컴/엑스큐솔루션/인네트웍스" },
    { value: "leaders", label: "리더스솔루션" },
    { value: "ieum", label: "이음플러스" }
  ];
  var closingOutsourceMarkers = ["정상", "연장", "심야", "특근"];
  var closingCloseTableRowCount = 52;
  var salesCloseCompaniesSeed = [
    "대영전자㈜ 광주지점", "대영전자 S/P는 신공장 사업자로 끊기", "대영 VINA", "대영일렉트릭", "디지티 (영세율)", "디지티 S/P",
    "HS일렉트릭", "대성포레스", "대정", "동아전기", "무등스크린", "삼광반도체", "성진전자기술", "세현전자", "승명이엔지",
    "아이테크코리아", "알.비.코리아", "에스디씨코퍼래이션", "에스제이아이 (구 성지산업)", "영림테크", "우인전자", "월드전자기술",
    "유미전자", "정신전자", "제이에스테크", "진흥전기", "창조테크", "플라텔"
  ];
  var purchaseCloseCompaniesSeed = [
    "아톰", "와이에스워터", "영테크", "하나에스피", "제이앤디테크", "서암", "엠텍", "태용", "효성앵글", "광성포장", "넥서스",
    "녹원씨엔아이", "다온씨엔티", "대흥인텍스", "덕성하이텍", "동영화성", "두리테크", "삼지산업", "성일테크", "성전목형",
    "에스에이치텍", "에이치티피", "엔씨씨통상", "오로라", "전원기획", "제일피복상사", "에스이테크", "보우테이프", "카고브릿지",
    "㈜휴먼스", "KGL (AIR)", "KGL (VSL)", "부천삼정 (경동택배)", "천일콤프레샤", "모세프린텍", "트랜스모바일", "우성ENG (구 티엔제이)"
  ];
  var closingAttendanceColumns = [
    { key: "employee", label: "이름", width: 118 },
    { key: "type", label: "구분", width: 76 }
  ];
  closingAttendanceDayFields.forEach(function (field, index) {
    closingAttendanceColumns.push({
      key: field,
      label: String(index + 1),
      width: 48
    });
  });
  var closingOutsourceColumns = [
    { key: "employee", label: "이름", width: 118 },
    { key: "type", label: "구분", width: 76 }
  ];
  closingAttendanceDayFields.forEach(function (field, index) {
    closingOutsourceColumns.push({
      key: field,
      label: String(index + 1),
      width: 52
    });
  });
  closingOutsourceColumns.push(
    { key: "weeklyAllowance", label: "주휴", width: 64 },
    { key: "perfectAllowance", label: "만근", width: 64 },
    { key: "retirementShare", label: "퇴직금", width: 72 },
    { key: "agencyFee", label: "수수료", width: 72 }
  );
  var workSheetColumns = [
    { key: "date", label: "일자", width: 88 },
    { key: "code", label: "코드", width: 104 },
    { key: "name", label: "품명", width: 240 },
    { key: "qty", label: "수량", width: 76 },
    { key: "price", label: "단가", width: 92 },
    { key: "supply", label: "공급가액", width: 108 },
    { key: "tax", label: "세액", width: 88 },
    { key: "note", label: "비고", width: 180 },
  ];
  var clientSheetColumns = [
    { key: "supplierMode", label: "구분", width: 84 },
    { key: "company", label: "상호", width: 180 },
    { key: "businessNo", label: "사업자등록번호", width: 148 },
    { key: "ceoName", label: "대표자명", width: 116 },
    { key: "address", label: "주소", width: 280 },
    { key: "businessType", label: "업태", width: 110 },
    { key: "businessItem", label: "종목", width: 180 },
  ];
  var priceSheetColumns = [
    { key: "code", label: "코드", width: 130 },
    { key: "price", label: "단가", width: 100 },
    { key: "name", label: "품목명", width: 360 },
  ];
  var workSheetEngine = null;
  var clientSheetEngine = null;
  var priceSheetEngine = null;
  var closingAttendanceSheetEngine = null;
  var closingOutsourceSheetEngine = null;
  var closingFingerprintSourceFile = null;

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
    autocompleteEl: null,
    autocompleteItems: [],
    autocompleteActiveIndex: -1,
    autocompleteForcedOpen: false,
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
      clientType: "all",
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
      normalized[i].client = getDisplayCompanyName(normalized[i].client);
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
    syncActiveLedgerBundle();
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

  function emptyPriceRow() {
    return {
      client: "",
      code: "",
      price: "",
      name: "",
    };
  }

  function emptyClosingSalesCloseRow(no) {
    return {
      no: no || "",
      division: "",
      company: "",
      closeDate: "",
      taxIssueDate: "",
      amount: "",
      mailSent: "",
      mailReply: "",
      issueConfirm: "",
      note: "",
    };
  }

  function emptyClosingPurchaseCloseRow(no) {
    return {
      no: no || "",
      division: "",
      company: "",
      closeIssueDate: "",
      supplyAmount: "",
      totalAmount: "",
      detailCheck: "",
      issueConfirm: "",
      note: "",
    };
  }

  function buildClosingSalesCloseSeedRows() {
    var rows = [];
    for (var i = 1; i <= closingCloseTableRowCount; i++) {
      var row = emptyClosingSalesCloseRow(String(i));
      row.company = salesCloseCompaniesSeed[i - 1] || "";
      if (i === 1) row.division = "법인";
      if (i === 8) row.division = "개인";
      rows.push(row);
    }
    return rows;
  }

  function buildClosingPurchaseCloseSeedRows() {
    var rows = [];
    for (var i = 1; i <= closingCloseTableRowCount; i++) {
      var row = emptyClosingPurchaseCloseRow(String(i));
      row.company = purchaseCloseCompaniesSeed[i - 1] || "";
      if (i === 1) row.division = "개인";
      if (i === 8) row.division = "법인";
      rows.push(row);
    }
    return rows;
  }

  function defaultClosingMonthLabel() {
    return new Date().getMonth() + 1 + "월";
  }

  function getClosingMonthOptions() {
    var options = [];
    for (var month = 1; month <= 12; month++) {
      options.push({ value: month + "월", label: month + "월" });
    }
    return options;
  }

  function emptyClosingEmployeeRow(name, type) {
    var row = {
      employee: name || "",
      type: type || "",
      count: "0",
    };
    closingAttendanceDayFields.forEach(function (field) {
      row[field] = "";
    });
    return row;
  }

  function buildClosingEmployeeTemplateRows() {
    var rows = [];
    closingEmployeeNames.forEach(function (name) {
      closingEmployeeMarkers.forEach(function (type, markerIndex) {
        rows.push(emptyClosingEmployeeRow(markerIndex === 0 ? name : "", type));
      });
    });
    return rows;
  }

  function parseClosingMonthNumber(label) {
    var match = String(label || "").match(/^(\d+)월$/);
    return match ? Number(match[1]) : 0;
  }

  function buildClosingEmployeeRowsFromRoster(sourceRows) {
    var normalized = normalizeClosingEmployeeRows(sourceRows);
    var rows = [];
    for (var groupStart = 0; groupStart < normalized.length; groupStart += closingEmployeeMarkers.length) {
      var header = normalized[groupStart] || {};
      var name = String(header.employee || "");
      for (var markerIndex = 0; markerIndex < closingEmployeeMarkers.length; markerIndex++) {
        rows.push(emptyClosingEmployeeRow(markerIndex === 0 ? name : "", closingEmployeeMarkers[markerIndex]));
      }
    }
    return rows.length ? rows : buildClosingEmployeeTemplateRows();
  }

  function getClosingEmployeeRosterFromRows(rows) {
    var normalized = normalizeClosingEmployeeRows(rows);
    var roster = [];
    for (var groupStart = 0; groupStart < normalized.length; groupStart += closingEmployeeMarkers.length) {
      roster.push(String((normalized[groupStart] && normalized[groupStart].employee) || ""));
    }
    if (!roster.length) {
      roster = closingEmployeeNames.slice();
    }
    return roster;
  }

  function isSameEmployeeRoster(a, b) {
    if (!Array.isArray(a) || !Array.isArray(b)) return false;
    if (a.length !== b.length) return false;
    for (var i = 0; i < a.length; i++) {
      if (String(a[i] || "") !== String(b[i] || "")) return false;
    }
    return true;
  }

  function buildClosingEmployeeRowsByRosterAndSource(roster, sourceRows) {
    var source = normalizeClosingEmployeeRows(sourceRows);
    var safeRoster = Array.isArray(roster) && roster.length ? roster : closingEmployeeNames.slice();
    var rows = [];
    for (var groupIndex = 0; groupIndex < safeRoster.length; groupIndex++) {
      for (var markerIndex = 0; markerIndex < closingEmployeeMarkers.length; markerIndex++) {
        var sourceRow = source[groupIndex * closingEmployeeMarkers.length + markerIndex] || {};
        var next = emptyClosingEmployeeRow(
          markerIndex === 0 ? String(safeRoster[groupIndex] || "") : "",
          closingEmployeeMarkers[markerIndex]
        );
        next.count = sourceRow.count != null && String(sourceRow.count).trim() !== "" ? String(sourceRow.count) : "0";
        closingAttendanceDayFields.forEach(function (field) {
          next[field] = sourceRow[field] != null ? String(sourceRow[field]) : "";
        });
        rows.push(next);
      }
    }
    return rows;
  }

  function applyClosingEmployeeRosterToAllMonths(roster) {
    var safeRoster = Array.isArray(roster) && roster.length ? roster : closingEmployeeNames.slice();
    closingState.employeeRoster = safeRoster.slice();
    if (!closingState.employeeRowsByMonth || typeof closingState.employeeRowsByMonth !== "object") {
      closingState.employeeRowsByMonth = {};
    }
    var keys = Object.keys(closingState.employeeRowsByMonth);
    if (!keys.length) {
      var monthKey = closingState.attendanceMonth || defaultClosingMonthLabel();
      closingState.employeeRowsByMonth[monthKey] = buildClosingEmployeeRowsByRosterAndSource(safeRoster, []);
      return;
    }
    keys.forEach(function (monthKey) {
      closingState.employeeRowsByMonth[monthKey] = buildClosingEmployeeRowsByRosterAndSource(
        safeRoster,
        closingState.employeeRowsByMonth[monthKey]
      );
    });
  }

  function getClosingEmployeeSeedRows(monthLabel) {
    var byMonth = closingState.employeeRowsByMonth || {};
    var targetMonth = parseClosingMonthNumber(monthLabel);
    if (!targetMonth) return null;
    for (var step = 1; step <= 11; step++) {
      var month = ((targetMonth - step - 1 + 120) % 12) + 1;
      var key = month + "월";
      if (key === monthLabel) continue;
      if (Array.isArray(byMonth[key]) && byMonth[key].length) {
        return byMonth[key];
      }
    }
    var keys = Object.keys(byMonth);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (key === monthLabel) continue;
      if (Array.isArray(byMonth[key]) && byMonth[key].length) {
        return byMonth[key];
      }
    }
    return null;
  }

  function normalizeClosingEmployeeRows(rows) {
    var sourceRows = Array.isArray(rows) && rows.length ? rows : buildClosingEmployeeTemplateRows();
    var groupCount = Math.max(1, Math.ceil(sourceRows.length / closingEmployeeMarkers.length));
    var normalized = [];
    for (var groupIndex = 0; groupIndex < groupCount; groupIndex++) {
      for (var markerIndex = 0; markerIndex < closingEmployeeMarkers.length; markerIndex++) {
        var source = sourceRows[groupIndex * closingEmployeeMarkers.length + markerIndex] || {};
        var next = emptyClosingEmployeeRow(markerIndex === 0 ? String(source.employee || "") : "", closingEmployeeMarkers[markerIndex]);
        next.employee = markerIndex === 0 ? String(source.employee || next.employee || "") : "";
        next.count = source.count != null && String(source.count).trim() !== "" ? String(source.count) : "0";
        closingAttendanceDayFields.forEach(function (field) {
          next[field] = source[field] != null ? String(source[field]) : "";
        });
        normalized.push(next);
      }
    }
    return normalized;
  }

  function emptyClosingOutsourceRow(name, type) {
    var row = {
      employee: name || "",
      joinDate: "",
      type: type || "",
      payCalc: "",
      bonus: "",
      payAmount: "",
    };
    closingAttendanceDayFields.forEach(function (field) {
      row[field] = "";
    });
    return row;
  }

  function buildClosingOutsourceTemplateRows() {
    var rows = [];
    for (var employeeIndex = 0; employeeIndex < 8; employeeIndex++) {
      closingOutsourceMarkers.forEach(function (type, markerIndex) {
        rows.push(emptyClosingOutsourceRow(markerIndex === 0 ? "" : "", type));
      });
    }
    return rows;
  }

  function normalizeClosingOutsourceRows(rows) {
    var sourceRows = Array.isArray(rows) && rows.length ? rows : buildClosingOutsourceTemplateRows();
    var groupCount = Math.max(1, Math.ceil(sourceRows.length / closingOutsourceMarkers.length));
    var normalized = [];
    for (var groupIndex = 0; groupIndex < groupCount; groupIndex++) {
      for (var markerIndex = 0; markerIndex < closingOutsourceMarkers.length; markerIndex++) {
        var source = sourceRows[groupIndex * closingOutsourceMarkers.length + markerIndex] || {};
        var next = emptyClosingOutsourceRow(markerIndex === 0 ? String(source.employee || "") : "", closingOutsourceMarkers[markerIndex]);
        next.employee = markerIndex === 0 ? String(source.employee || next.employee || "") : "";
        next.joinDate = markerIndex === 0 ? String(source.joinDate || "") : "";
        next.type = source.type != null ? String(source.type) : next.type;
        closingAttendanceDayFields.forEach(function (field) {
          next[field] = source[field] != null ? String(source[field]) : "";
        });
        next.payCalc = source.payCalc != null ? String(source.payCalc) : "";
        next.bonus = source.bonus != null ? String(source.bonus) : "";
        next.payAmount = source.payAmount != null ? String(source.payAmount) : "";
        normalized.push(next);
      }
    }
    return normalized;
  }

  function getClosingOutsourceRosterFromRows(rows) {
    var normalized = normalizeClosingOutsourceRows(rows);
    var roster = [];
    for (var groupStart = 0; groupStart < normalized.length; groupStart += closingOutsourceMarkers.length) {
      var header = normalized[groupStart] || {};
      roster.push({
        employee: String(header.employee || ""),
        joinDate: String(header.joinDate || ""),
      });
    }
    return roster.length ? roster : [{ employee: "", joinDate: "" }];
  }

  function buildClosingOutsourceRowsByRosterAndSource(roster, sourceRows) {
    var source = normalizeClosingOutsourceRows(sourceRows);
    var safeRoster = Array.isArray(roster) && roster.length ? roster : [{ employee: "", joinDate: "" }];
    var rows = [];
    for (var groupIndex = 0; groupIndex < safeRoster.length; groupIndex++) {
      for (var markerIndex = 0; markerIndex < closingOutsourceMarkers.length; markerIndex++) {
        var sourceRow = source[groupIndex * closingOutsourceMarkers.length + markerIndex] || {};
        var profile = safeRoster[groupIndex] || {};
        var next = emptyClosingOutsourceRow(
          markerIndex === 0 ? String(profile.employee || "") : "",
          closingOutsourceMarkers[markerIndex]
        );
        next.joinDate = markerIndex === 0 ? String(profile.joinDate || "") : "";
        next.type = sourceRow.type != null ? String(sourceRow.type) : next.type;
        closingAttendanceDayFields.forEach(function (field) {
          next[field] = sourceRow[field] != null ? String(sourceRow[field]) : "";
        });
        next.payCalc = sourceRow.payCalc != null ? String(sourceRow.payCalc) : "";
        next.bonus = sourceRow.bonus != null ? String(sourceRow.bonus) : "";
        next.payAmount = sourceRow.payAmount != null ? String(sourceRow.payAmount) : "";
        rows.push(next);
      }
    }
    return rows;
  }

  function getClosingOutsourceRoster(vendor) {
    ensureClosingAttendanceState();
    var nextVendor = vendor || closingState.outsourceVendor || "leaders";
    var roster = closingState.outsourceRosterByVendor[nextVendor];
    if (!Array.isArray(roster) || !roster.length) {
      var seedRows = getClosingOutsourceSeedRows(nextVendor, closingState.attendanceMonth);
      roster = getClosingOutsourceRosterFromRows(seedRows || buildClosingOutsourceTemplateRows());
      closingState.outsourceRosterByVendor[nextVendor] = roster;
    }
    return closingState.outsourceRosterByVendor[nextVendor];
  }

  function isSameClosingOutsourceRoster(a, b) {
    if (!Array.isArray(a) || !Array.isArray(b)) return false;
    if (a.length !== b.length) return false;
    for (var i = 0; i < a.length; i++) {
      var left = a[i] || {};
      var right = b[i] || {};
      if (String(left.employee || "") !== String(right.employee || "")) return false;
      if (String(left.joinDate || "") !== String(right.joinDate || "")) return false;
    }
    return true;
  }

  function getClosingOutsourceSeedRows(vendor, monthLabel) {
    var source = closingState.outsourceRowsByKey && typeof closingState.outsourceRowsByKey === "object"
      ? closingState.outsourceRowsByKey
      : {};
    var targetMonth = parseClosingMonthNumber(monthLabel);
    if (targetMonth) {
      for (var step = 1; step <= 11; step++) {
        var month = ((targetMonth - step - 1 + 120) % 12) + 1;
        var key = vendor + "::" + month + "월";
        if (Array.isArray(source[key]) && source[key].length) return source[key];
      }
    }
    var prefix = vendor + "::";
    var keys = Object.keys(source);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (key.indexOf(prefix) !== 0) continue;
      if (Array.isArray(source[key]) && source[key].length) return source[key];
    }
    return null;
  }

  function applyClosingOutsourceRosterToVendorAllMonths(vendor, roster) {
    var nextVendor = vendor || closingState.outsourceVendor || "leaders";
    var safeRoster = Array.isArray(roster) && roster.length ? roster : [{ employee: "", joinDate: "" }];
    closingState.outsourceRosterByVendor[nextVendor] = safeRoster.slice();
    var prefix = nextVendor + "::";
    var keys = Object.keys(closingState.outsourceRowsByKey || {}).filter(function (key) {
      return key.indexOf(prefix) === 0;
    });
    if (!keys.length) {
      var key = getClosingOutsourceKey(closingState.attendanceMonth, nextVendor);
      closingState.outsourceRowsByKey[key] = buildClosingOutsourceRowsByRosterAndSource(safeRoster, []);
      return;
    }
    keys.forEach(function (key) {
      closingState.outsourceRowsByKey[key] = buildClosingOutsourceRowsByRosterAndSource(
        safeRoster,
        closingState.outsourceRowsByKey[key]
      );
    });
  }

  function cloneClosingRowsMap(rowsByMonth) {
    var source = rowsByMonth && typeof rowsByMonth === "object" ? rowsByMonth : {};
    var cloned = {};
    Object.keys(source).forEach(function (monthKey) {
      cloned[monthKey] = normalizeClosingEmployeeRows(source[monthKey]);
    });
    return cloned;
  }

  function ensureClosingAttendanceState() {
    if (!closingState.attendanceMonth || !/^\d+월$/.test(closingState.attendanceMonth)) {
      closingState.attendanceMonth = defaultClosingMonthLabel();
    }
    if (closingState.attendanceView !== "outsource") {
      closingState.attendanceView = "employee";
    }
    if (closingState.employeeViewMode !== "summary") {
      closingState.employeeViewMode = "entry";
    }
    if (!Array.isArray(closingState.employeeRoster)) {
      closingState.employeeRoster = [];
    }
    if (!closingState.employeeRowsByMonth || typeof closingState.employeeRowsByMonth !== "object") {
      closingState.employeeRowsByMonth = {};
    }
    if (!closingState.outsourceVendor) {
      closingState.outsourceVendor = "leaders";
    }
    if (!closingState.outsourceRowsByKey || typeof closingState.outsourceRowsByKey !== "object") {
      closingState.outsourceRowsByKey = {};
    }
    if (!closingState.outsourceRosterByVendor || typeof closingState.outsourceRosterByVendor !== "object") {
      closingState.outsourceRosterByVendor = {};
    }
    if (!closingState.outsourceWarningsByKey || typeof closingState.outsourceWarningsByKey !== "object") {
      closingState.outsourceWarningsByKey = {};
    }
    if (!closingState.outsourceMetaByKey || typeof closingState.outsourceMetaByKey !== "object") {
      closingState.outsourceMetaByKey = {};
    }
    if (!closingState.salesCloseRowsByMonth || typeof closingState.salesCloseRowsByMonth !== "object") {
      closingState.salesCloseRowsByMonth = {};
    }
    if (!closingState.purchaseCloseRowsByMonth || typeof closingState.purchaseCloseRowsByMonth !== "object") {
      closingState.purchaseCloseRowsByMonth = {};
    }
    if (typeof closingState.attendanceSearch !== "string") {
      closingState.attendanceSearch = "";
    }
    if (!closingState.employeeRoster.length) {
      var seedRows = closingState.employeeRowsByMonth[closingState.attendanceMonth] || getClosingEmployeeSeedRows(closingState.attendanceMonth);
      closingState.employeeRoster = getClosingEmployeeRosterFromRows(seedRows || buildClosingEmployeeTemplateRows());
    }
    if (!closingState.employeeRowsByMonth[closingState.attendanceMonth]) {
      closingState.employeeRowsByMonth[closingState.attendanceMonth] = buildClosingEmployeeRowsByRosterAndSource(
        closingState.employeeRoster,
        []
      );
    } else {
      closingState.employeeRowsByMonth[closingState.attendanceMonth] = buildClosingEmployeeRowsByRosterAndSource(
        closingState.employeeRoster,
        closingState.employeeRowsByMonth[closingState.attendanceMonth]
      );
    }
    return closingState;
  }

  function normalizeClosingSalesCloseRows(rows) {
    var source = Array.isArray(rows) ? rows : [];
    var normalized = [];
    for (var i = 0; i < Math.max(closingCloseTableRowCount, source.length); i++) {
      var row = source[i] || emptyClosingSalesCloseRow(String(i + 1));
      normalized.push({
        no: String((row && row.no) || (i + 1)),
        division: String((row && row.division) || ""),
        company: String((row && row.company) || ""),
        closeDate: String((row && row.closeDate) || ""),
        taxIssueDate: String((row && row.taxIssueDate) || ""),
        amount: String((row && row.amount) || ""),
        mailSent: String((row && row.mailSent) || ""),
        mailReply: String((row && row.mailReply) || ""),
        issueConfirm: String((row && row.issueConfirm) || ""),
        note: String((row && row.note) || ""),
      });
    }
    return normalized.slice(0, closingCloseTableRowCount);
  }

  function normalizeClosingPurchaseCloseRows(rows) {
    var source = Array.isArray(rows) ? rows : [];
    var normalized = [];
    for (var i = 0; i < Math.max(closingCloseTableRowCount, source.length); i++) {
      var row = source[i] || emptyClosingPurchaseCloseRow(String(i + 1));
      normalized.push({
        no: String((row && row.no) || (i + 1)),
        division: String((row && row.division) || ""),
        company: String((row && row.company) || ""),
        closeIssueDate: String((row && row.closeIssueDate) || ""),
        supplyAmount: String((row && row.supplyAmount) || ""),
        totalAmount: String((row && row.totalAmount) || ""),
        detailCheck: String((row && row.detailCheck) || ""),
        issueConfirm: String((row && row.issueConfirm) || ""),
        note: String((row && row.note) || ""),
      });
    }
    return normalized.slice(0, closingCloseTableRowCount);
  }

  function getClosingSalesCloseRows(monthLabel) {
    ensureClosingAttendanceState();
    var key = monthLabel || closingState.attendanceMonth;
    if (!Array.isArray(closingState.salesCloseRowsByMonth[key])) {
      closingState.salesCloseRowsByMonth[key] = buildClosingSalesCloseSeedRows();
    } else {
      closingState.salesCloseRowsByMonth[key] = normalizeClosingSalesCloseRows(closingState.salesCloseRowsByMonth[key]);
    }
    return closingState.salesCloseRowsByMonth[key];
  }

  function setClosingSalesCloseRows(monthLabel, rows) {
    ensureClosingAttendanceState();
    var key = monthLabel || closingState.attendanceMonth;
    closingState.salesCloseRowsByMonth[key] = normalizeClosingSalesCloseRows(rows);
  }

  function getClosingPurchaseCloseRows(monthLabel) {
    ensureClosingAttendanceState();
    var key = monthLabel || closingState.attendanceMonth;
    if (!Array.isArray(closingState.purchaseCloseRowsByMonth[key])) {
      closingState.purchaseCloseRowsByMonth[key] = buildClosingPurchaseCloseSeedRows();
    } else {
      closingState.purchaseCloseRowsByMonth[key] = normalizeClosingPurchaseCloseRows(closingState.purchaseCloseRowsByMonth[key]);
    }
    return closingState.purchaseCloseRowsByMonth[key];
  }

  function setClosingPurchaseCloseRows(monthLabel, rows) {
    ensureClosingAttendanceState();
    var key = monthLabel || closingState.attendanceMonth;
    closingState.purchaseCloseRowsByMonth[key] = normalizeClosingPurchaseCloseRows(rows);
  }

  function getClosingEmployeeRows(monthLabel) {
    ensureClosingAttendanceState();
    var key = monthLabel || closingState.attendanceMonth;
    if (!closingState.employeeRowsByMonth[key]) {
      closingState.employeeRowsByMonth[key] = buildClosingEmployeeRowsByRosterAndSource(
        closingState.employeeRoster,
        []
      );
    } else {
      closingState.employeeRowsByMonth[key] = buildClosingEmployeeRowsByRosterAndSource(
        closingState.employeeRoster,
        closingState.employeeRowsByMonth[key]
      );
    }
    return closingState.employeeRowsByMonth[key];
  }

  function setClosingEmployeeRows(monthLabel, rows) {
    ensureClosingAttendanceState();
    var key = monthLabel || closingState.attendanceMonth;
    var normalized = normalizeClosingEmployeeRows(rows);
    closingState.employeeRowsByMonth[key] = normalized;
    var nextRoster = getClosingEmployeeRosterFromRows(normalized);
    if (!isSameEmployeeRoster(nextRoster, closingState.employeeRoster)) {
      applyClosingEmployeeRosterToAllMonths(nextRoster);
    } else {
      closingState.employeeRowsByMonth[key] = buildClosingEmployeeRowsByRosterAndSource(
        closingState.employeeRoster,
        closingState.employeeRowsByMonth[key]
      );
    }
  }

  function getClosingOutsourceKey(monthLabel, vendor) {
    var month = monthLabel || ensureClosingAttendanceState().attendanceMonth;
    var nextVendor = vendor || closingState.outsourceVendor || "leaders";
    return nextVendor + "::" + month;
  }

  function getClosingOutsourceRows(monthLabel, vendor) {
    ensureClosingAttendanceState();
    var key = getClosingOutsourceKey(monthLabel, vendor);
    var roster = getClosingOutsourceRoster(vendor);
    if (!closingState.outsourceRowsByKey[key]) {
      closingState.outsourceRowsByKey[key] = buildClosingOutsourceRowsByRosterAndSource(roster, []);
    } else {
      closingState.outsourceRowsByKey[key] =
        buildClosingOutsourceRowsByRosterAndSource(roster, closingState.outsourceRowsByKey[key]);
    }
    return closingState.outsourceRowsByKey[key];
  }

  function setClosingOutsourceRows(monthLabel, vendor, rows) {
    ensureClosingAttendanceState();
    var key = getClosingOutsourceKey(monthLabel, vendor);
    var nextVendor = vendor || closingState.outsourceVendor || "leaders";
    var normalized = normalizeClosingOutsourceRows(rows);
    closingState.outsourceRowsByKey[key] = normalized;
    var nextRoster = getClosingOutsourceRosterFromRows(normalized);
    var currentRoster = getClosingOutsourceRoster(nextVendor);
    if (!isSameClosingOutsourceRoster(nextRoster, currentRoster)) {
      applyClosingOutsourceRosterToVendorAllMonths(nextVendor, nextRoster);
    } else {
      closingState.outsourceRowsByKey[key] = buildClosingOutsourceRowsByRosterAndSource(
        currentRoster,
        closingState.outsourceRowsByKey[key]
      );
    }
  }

  function getClosingOutsourceWarnings(monthLabel, vendor) {
    ensureClosingAttendanceState();
    var key = getClosingOutsourceKey(monthLabel, vendor);
    if (!closingState.outsourceWarningsByKey[key] || typeof closingState.outsourceWarningsByKey[key] !== "object") {
      closingState.outsourceWarningsByKey[key] = {};
    }
    return closingState.outsourceWarningsByKey[key];
  }

  function setClosingOutsourceWarnings(monthLabel, vendor, warnings) {
    ensureClosingAttendanceState();
    closingState.outsourceWarningsByKey[getClosingOutsourceKey(monthLabel, vendor)] =
      warnings && typeof warnings === "object" ? warnings : {};
  }

  function getClosingOutsourceMeta(monthLabel, vendor) {
    ensureClosingAttendanceState();
    var key = getClosingOutsourceKey(monthLabel, vendor);
    if (!closingState.outsourceMetaByKey[key] || typeof closingState.outsourceMetaByKey[key] !== "object") {
      closingState.outsourceMetaByKey[key] = { retirement: "", hourlyWage: "", lockDay: "" };
    }
    return closingState.outsourceMetaByKey[key];
  }

  function setClosingOutsourceMeta(monthLabel, vendor, meta) {
    ensureClosingAttendanceState();
    closingState.outsourceMetaByKey[getClosingOutsourceKey(monthLabel, vendor)] = Object.assign(
      { retirement: "", hourlyWage: "", lockDay: "" },
      meta || {}
    );
  }

  function getClosingOutsourceLockDay(meta, daysInMonth) {
    var maxDay = daysInMonth || getClosingDaysInMonth();
    var value = parseCalcNumber(meta && meta.lockDay);
    if (value == null) return 0;
    value = Math.floor(value);
    if (value < 0) value = 0;
    if (value > maxDay) value = maxDay;
    return value;
  }

  function getClosingAttendanceYear() {
    return new Date().getFullYear();
  }

  function getClosingAttendanceMonthNumber() {
    ensureClosingAttendanceState();
    var match = String(closingState.attendanceMonth || "").match(/^(\d+)월$/);
    return match ? Number(match[1]) : new Date().getMonth() + 1;
  }

  function getClosingDaysInMonth() {
    var year = getClosingAttendanceYear();
    var month = getClosingAttendanceMonthNumber();
    return new Date(year, month, 0).getDate();
  }

  function isClosingWeekend(day) {
    var label = getClosingWeekdayLabel(day);
    return label === "토" || label === "일";
  }

  function getClosingWarningDayLimit() {
    var maxDays = getClosingDaysInMonth();
    var now = new Date();
    if (
      getClosingAttendanceYear() === now.getFullYear() &&
      getClosingAttendanceMonthNumber() === now.getMonth() + 1
    ) {
      return Math.max(1, Math.min(maxDays, now.getDate()));
    }
    return maxDays;
  }

  function getClosingWeekdayLabel(day) {
    var labels = ["일", "월", "화", "수", "목", "금", "토"];
    var year = getClosingAttendanceYear();
    var month = getClosingAttendanceMonthNumber();
    return labels[new Date(year, month - 1, day).getDay()];
  }

  function getClosingWeekdayClass(day) {
    var label = getClosingWeekdayLabel(day);
    if (label === "일") return " sunday";
    if (label === "토") return " saturday";
    return "";
  }

  function calculateClosingOutsourceTimeTotal(row) {
    var total = 0;
    var daysInMonth = getClosingDaysInMonth();
    closingAttendanceDayFields.forEach(function (field, index) {
      if (index + 1 > daysInMonth) return;
      total += parseCalcNumber(row && row[field]) || 0;
    });
    return total;
  }

  function normalizeJoinDateText(value) {
    var text = String(value || "").trim();
    if (!text || text === "근속 시작일") return "";
    var digits = text.replace(/\D/g, "");
    if (digits.length === 8) {
      return digits.slice(0, 4) + "-" + digits.slice(4, 6) + "-" + digits.slice(6);
    }
    return text;
  }

  function calculateClosingOutsourceGrandTotal(rows) {
    var grouped = [];
    for (var i = 0; i < rows.length; i += closingOutsourceMarkers.length) {
      var block = rows.slice(i, i + closingOutsourceMarkers.length);
      var amount = 0;
      block.forEach(function (row) {
        amount += parseCalcNumber(row && row.payAmount) || 0;
      });
      grouped.push(amount);
    }
    return grouped.reduce(function (sum, value) { return sum + value; }, 0);
  }

  var DEFAULT_CLOSING_HOURLY_WAGE_2026_KR = 10320;

  function getClosingOutsourceHourlyWage(meta) {
    var value = parseCalcNumber(meta && meta.hourlyWage);
    if (value == null || value <= 0) return DEFAULT_CLOSING_HOURLY_WAGE_2026_KR;
    return value;
  }

  function isOutsourceAbsenceValue(value) {
    var text = String(value || "").trim();
    if (!text) return false;
    if (/결근/.test(text)) return true;
    var num = parseCalcNumber(text);
    if (num == null) return false;
    return num <= 0;
  }

  function isOutsourceWeeklyAbsenceValue(value, hasWarning) {
    var text = String(value || "").trim();
    if (hasWarning && !text) return true;
    if (!text) return false;
    if (/결근/.test(text)) return true;
    var num = parseCalcNumber(text);
    if (num == null) return false;
    return num <= 0;
  }

  function getWeekdayLabelByDate(year, month, day) {
    var labels = ["일", "월", "화", "수", "목", "금", "토"];
    return labels[new Date(year, month - 1, day).getDay()];
  }

  function getPreviousClosingMonthInfo() {
    var year = getClosingAttendanceYear();
    var month = getClosingAttendanceMonthNumber();
    var prevMonth = month === 1 ? 12 : month - 1;
    var prevYear = month === 1 ? year - 1 : year;
    var prevDays = new Date(prevYear, prevMonth, 0).getDate();
    return {
      year: prevYear,
      month: prevMonth,
      monthLabel: prevMonth + "월",
      daysInMonth: prevDays,
    };
  }

  function findOutsourceNormalRowByEmployee(monthLabel, vendor, employeeName) {
    var rows = normalizeClosingOutsourceRows(getClosingOutsourceRows(monthLabel, vendor));
    var token = normalizeClosingSearchText(employeeName || "");
    if (!token) return null;
    for (var i = 0; i < rows.length; i += closingOutsourceMarkers.length) {
      var name = normalizeClosingSearchText(rows[i] && rows[i].employee);
      if (!name) continue;
      if (name === token || name.indexOf(token) >= 0 || token.indexOf(name) >= 0) {
        return rows[i];
      }
    }
    return null;
  }

  function getOutsourceNormalCellValueForWeekly(currentNormalRow, previousNormalRow, dayInCurrentMonth, previousMonthInfo) {
    if (dayInCurrentMonth >= 1) {
      return currentNormalRow ? currentNormalRow["d" + dayInCurrentMonth] : "";
    }
    var prevDay = previousMonthInfo.daysInMonth + dayInCurrentMonth;
    if (!previousNormalRow || prevDay < 1 || prevDay > previousMonthInfo.daysInMonth) return "";
    return previousNormalRow["d" + prevDay];
  }

  function calculateClosingOutsourceWeeklyHolidayHours(block, warnings, groupStart, daysInMonth, applyToNormalRow, lockDay) {
    var normalRow = block && block[0] ? block[0] : {};
    var employeeName = String(normalRow.employee || "");
    var previousMonthInfo = getPreviousClosingMonthInfo();
    var previousNormalRow = findOutsourceNormalRowByEmployee(
      previousMonthInfo.monthLabel,
      closingState.outsourceVendor,
      employeeName
    );
    var dayLimit = Math.min(daysInMonth, getClosingWarningDayLimit());
    var extraHours = 0;
    for (var sunday = 1; sunday <= dayLimit; sunday++) {
      if (getClosingWeekdayLabel(sunday) !== "일") continue;
      var hasWeekday = false;
      var absentInWeek = false;
      for (var day = sunday - 6; day <= sunday - 1; day++) {
        if (day > dayLimit) break;
        var weekday = day >= 1
          ? getClosingWeekdayLabel(day)
          : getWeekdayLabelByDate(previousMonthInfo.year, previousMonthInfo.month, previousMonthInfo.daysInMonth + day);
        if (weekday === "일") continue;
        hasWeekday = true;
        var value = getOutsourceNormalCellValueForWeekly(
          normalRow,
          previousNormalRow,
          day,
          previousMonthInfo
        );
        var hasWarning = false;
        if (day >= 1) {
          var field = "d" + day;
          var warningKey = groupStart + ":" + field;
          hasWarning = !!(warnings && warnings[warningKey]);
        }
        if (isOutsourceWeeklyAbsenceValue(value, hasWarning)) {
          absentInWeek = true;
          break;
        }
      }
      if (hasWeekday && !absentInWeek) {
        var sundayField = "d" + sunday;
        var sundayText = String(normalRow[sundayField] || "").trim();
        var sundayNum = parseCalcNumber(sundayText);
        if (sundayNum != null && sundayNum > 0) {
          continue;
        }
        if (sundayText && sundayNum == null) {
          continue;
        }
        var locked = applyToNormalRow && Number(lockDay || 0) >= sunday;
        if (locked) {
          continue;
        }
        if (applyToNormalRow) {
          normalRow[sundayField] = "8";
        }
        extraHours += 8;
      }
    }
    return extraHours;
  }

  function calculateClosingOutsourcePerfectAttendancePay(block, warnings, groupStart, daysInMonth, hourlyWage) {
    var normalRow = block && block[0] ? block[0] : {};
    var dayLimit = Math.min(daysInMonth, getClosingWarningDayLimit());
    for (var day = 1; day <= dayLimit; day++) {
      var weekday = getClosingWeekdayLabel(day);
      if (weekday === "일") continue;
      var field = "d" + day;
      if (isOutsourceAbsenceValue(normalRow[field])) {
        return 0;
      }
    }
    // 만근수당은 월 1회, 1일치(8시간)만 지급한다.
    return 8 * DEFAULT_CLOSING_HOURLY_WAGE_2026_KR;
  }

  function getClosingOutsourceWageMultiplier(type) {
    var marker = String(type || "").trim();
    if (marker === "연장" || marker === "심야" || marker === "특근") {
      return 1.5;
    }
    return 1;
  }

  function calculateClosingOutsourceDerived(block, warnings, groupStart, daysInMonth, hourlyWage) {
    var weeklyHolidayHours = calculateClosingOutsourceWeeklyHolidayHours(block, warnings, groupStart, daysInMonth);
    var perfectAttendancePay = calculateClosingOutsourcePerfectAttendancePay(
      block,
      warnings,
      groupStart,
      daysInMonth,
      hourlyWage
    );
    var derivedRows = [];
    var groupTotal = 0;
    for (var i = 0; i < closingOutsourceMarkers.length; i++) {
      var row = block[i] || emptyClosingOutsourceRow("", closingOutsourceMarkers[i]);
      var rawTimeTotal = calculateClosingOutsourceTimeTotal(row);
      // 시간은 입력값 그대로 합산한다. (주휴 자동 +8 가산 비활성화)
      var timeTotal = rawTimeTotal;
      var multiplier = getClosingOutsourceWageMultiplier(row.type || closingOutsourceMarkers[i]);
      var basePay = timeTotal * hourlyWage * multiplier;
      // 주휴는 시간/급여계산(기본급)으로만 반영하고, 만근수당은 별도 1일치만 유지한다.
      var bonusPay = i === 0 ? perfectAttendancePay : 0;
      var payAmount = basePay + bonusPay;
      groupTotal += payAmount;
      derivedRows.push({
        timeTotal: timeTotal,
        basePay: basePay,
        bonusPay: bonusPay,
        payAmount: payAmount,
      });
    }
    return {
      weeklyHolidayHours: weeklyHolidayHours,
      perfectAttendancePay: perfectAttendancePay,
      rows: derivedRows,
      groupTotal: groupTotal,
    };
  }

  function getClosingOutsourceMgmtRate(vendor) {
    return vendor === "ieum" ? 0.12 : 0.15;
  }

  function formatClosingCurrency(value) {
    return formatDisplayNumber(Math.round((parseCalcNumber(value) || 0) * 1000) / 1000);
  }

  function getClosingEmployeePeriodLabel() {
    var year = getClosingAttendanceYear();
    var month = getClosingAttendanceMonthNumber();
    var prevMonth = month === 1 ? 12 : month - 1;
    var prevYear = month === 1 ? year - 1 : year;
    return prevYear + "년 " + prevMonth + "월 25일 ~ " + year + "년 " + month + "월 24일";
  }

  function getClosingEmployeePeriodDays() {
    var year = getClosingAttendanceYear();
    var month = getClosingAttendanceMonthNumber();
    var prevMonth = month === 1 ? 12 : month - 1;
    var prevYear = month === 1 ? year - 1 : year;
    var start = new Date(prevYear, prevMonth - 1, 25);
    var end = new Date(year, month - 1, 24);
    var cursor = new Date(start.getTime());
    var days = [];
    while (cursor <= end && days.length < closingAttendanceDayFields.length) {
      var y = cursor.getFullYear();
      var m = cursor.getMonth() + 1;
      var d = cursor.getDate();
      var weekday = ["일", "월", "화", "수", "목", "금", "토"][cursor.getDay()];
      days.push({
        year: y,
        month: m,
        day: d,
        weekday: weekday,
      });
      cursor.setDate(cursor.getDate() + 1);
    }
    return days;
  }

  function getClosingEmployeeCount(row, daySlots) {
    var count = 0;
    var slotLength = Array.isArray(daySlots) ? daySlots.length : 0;
    for (var i = 0; i < slotLength; i++) {
      var field = closingAttendanceDayFields[i];
      var value = String((row && row[field]) || "").trim();
      if (!value || value === "0") continue;
      count++;
    }
    return count;
  }

  function syncClosingEmployeeCounts(rows, daySlots) {
    var normalized = normalizeClosingEmployeeRows(rows);
    for (var i = 0; i < normalized.length; i++) {
      normalized[i].count = String(getClosingEmployeeCount(normalized[i], daySlots));
    }
    return normalized;
  }

  function addClosingEmployeeGroup(name) {
    var rows = normalizeClosingEmployeeRows(getClosingEmployeeRows(closingState.attendanceMonth));
    closingEmployeeMarkers.forEach(function (type, markerIndex) {
      rows.push(emptyClosingEmployeeRow(markerIndex === 0 ? (name || "") : "", type));
    });
    setClosingEmployeeRows(closingState.attendanceMonth, rows);
  }

  function removeClosingEmployeeGroup(groupStart) {
    var rows = normalizeClosingEmployeeRows(getClosingEmployeeRows(closingState.attendanceMonth));
    rows.splice(groupStart, closingEmployeeMarkers.length);
    if (!rows.length) {
      addClosingEmployeeGroup("");
      return;
    }
    setClosingEmployeeRows(closingState.attendanceMonth, rows);
  }

  function addClosingOutsourceGroup(name) {
    var rows = normalizeClosingOutsourceRows(getClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor));
    closingOutsourceMarkers.forEach(function (type, markerIndex) {
      rows.push(emptyClosingOutsourceRow(markerIndex === 0 ? (name || "") : "", type));
    });
    setClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor, rows);
  }

  function removeClosingOutsourceGroup(groupStart) {
    var rows = normalizeClosingOutsourceRows(getClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor));
    rows.splice(groupStart, closingOutsourceMarkers.length);
    if (!rows.length) {
      addClosingOutsourceGroup("");
      return;
    }
    setClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor, rows);
  }

  function getClosingOutsourceVendorLabel() {
    var current = closingOutsourceVendors.find(function (item) {
      return item.value === closingState.outsourceVendor;
    });
    return current ? current.label : "아웃소싱";
  }

  function normalizeClosingSearchText(value) {
    return String(value || "").replace(/\s+/g, "").toLowerCase();
  }

  function matchesClosingAttendanceSearch(employee, joinDate) {
    var keyword = normalizeClosingSearchText(closingState.attendanceSearch || "");
    if (!keyword) return true;
    var employeeText = normalizeClosingSearchText(employee || "");
    var joinDateText = normalizeClosingSearchText(joinDate || "");
    return employeeText.indexOf(keyword) >= 0 || joinDateText.indexOf(keyword) >= 0;
  }

  function parseSpreadsheetTimeValue(value) {
    var text = String(value || "").trim();
    var match = text.match(/T(\d{2}):(\d{2}):(\d{2})/);
    if (!match) return null;
    return {
      hour: Number(match[1]),
      minute: Number(match[2]),
      second: Number(match[3]),
      totalSeconds: Number(match[1]) * 3600 + Number(match[2]) * 60 + Number(match[3]),
    };
  }

  function isTimeBetween(time, startHour, startMinute, endHour, endMinute) {
    if (!time) return false;
    var total = time.totalSeconds;
    var start = startHour * 3600 + startMinute * 60;
    var end = endHour * 3600 + endMinute * 60;
    return total >= start && total <= end;
  }

  function calculateLateDeductionHours(time) {
    if (!time) return 0;
    var nine = 9 * 3600;
    var diff = time.totalSeconds - nine;
    if (diff <= 359) return 0;
    return Math.ceil((diff - 359) / 1800) * 0.5;
  }

  function calculateEarlyDeductionHours(time) {
    if (!time) return 8;
    var endGrace = 17 * 3600 + 55 * 60;
    if (time.totalSeconds >= endGrace) return 0;
    var six = 18 * 3600;
    var diff = Math.max(0, six - time.totalSeconds);
    return Math.ceil(diff / 1800) * 0.5;
  }

  function parseSpreadsheetXmlRows(rowNodes) {
    return Array.prototype.map.call(rowNodes || [], function (rowNode) {
      var result = {};
      var col = 1;
      Array.prototype.forEach.call(rowNode.childNodes || [], function (child) {
        if (!child || child.nodeType !== 1 || child.localName !== "Cell") return;
        var idx = child.getAttribute("ss:Index") || child.getAttribute("Index") ||
          child.getAttributeNS("urn:schemas-microsoft-com:office:spreadsheet", "Index");
        if (idx) col = Number(idx);
        var text = "";
        Array.prototype.forEach.call(child.childNodes || [], function (cellChild) {
          if (!cellChild || cellChild.nodeType !== 1 || cellChild.localName !== "Data") return;
          text = String(cellChild.textContent || "").trim();
        });
        result[col] = text;
        col++;
      });
      return result;
    });
  }

  function parseFingerprintXml(text) {
    var parser = new DOMParser();
    var xml = parser.parseFromString(String(text || ""), "application/xml");
    if (xml.getElementsByTagName("parsererror").length) {
      throw new Error("지문 원본 파일 형식을 읽지 못했습니다.");
    }
    var worksheets = Array.prototype.filter.call(
      xml.getElementsByTagNameNS("urn:schemas-microsoft-com:office:spreadsheet", "Worksheet"),
      function (node) {
        var name =
          node.getAttribute("ss:Name") ||
          node.getAttribute("Name") ||
          node.getAttributeNS("urn:schemas-microsoft-com:office:spreadsheet", "Name");
        return name === "개별 보고서";
      }
    );
    if (!worksheets.length) {
      throw new Error("개별 보고서 시트를 찾지 못했습니다.");
    }
    var table = worksheets[0].getElementsByTagNameNS("urn:schemas-microsoft-com:office:spreadsheet", "Table")[0];
    if (!table) throw new Error("개별 보고서 표를 찾지 못했습니다.");
    var rows = parseSpreadsheetXmlRows(table.getElementsByTagNameNS("urn:schemas-microsoft-com:office:spreadsheet", "Row"));
    var byEmployee = {};
    var currentName = "";
    rows.forEach(function (row) {
      if (row[1] && /^ID:/i.test(row[1]) && row[2] && /^이름:/i.test(row[2])) {
        currentName = String(row[2]).replace(/^이름:/i, "").trim();
        if (currentName && !byEmployee[currentName]) byEmployee[currentName] = {};
        return;
      }
      if (!currentName || !row[1] || !/^\d{4}-\d{1,2}-\d{1,2}T/.test(row[1])) return;
      var dayMatch = String(row[1]).match(/^\d{4}-(\d{1,2})-(\d{1,2})T/);
      if (!dayMatch) return;
      var month = Number(dayMatch[1]);
      var day = Number(dayMatch[2]);
      byEmployee[currentName][month + "-" + day] = {
        weekday: String(row[2] || ""),
        in1: parseSpreadsheetTimeValue(row[3]),
        out1: parseSpreadsheetTimeValue(row[4]),
        in2: parseSpreadsheetTimeValue(row[5]),
        out2: parseSpreadsheetTimeValue(row[6]),
        in3: parseSpreadsheetTimeValue(row[7]),
        out3: parseSpreadsheetTimeValue(row[8]),
      };
    });
    return byEmployee;
  }

  function calculateOutsourceRecord(record) {
    if (!record) return { warning: true, reason: "no_record" };
    var firstIn = record.in1 || record.in2 || record.in3;
    var lastOut = record.out3 || record.out2 || record.out1;
    if (!firstIn || !lastOut) {
      return { warning: true, reason: "missing_in_or_out" };
    }
    var weekday = record.weekday;
    var isWeekend = weekday === "토" || weekday === "일";
    var hasAfternoonPunch = !!(record.in2 || record.in3);
    var result = {
      normal: "",
      overtime: "",
      night: "",
      special: "",
      warning: false,
    };

    if (!hasAfternoonPunch && isTimeBetween(lastOut, 12, 33, 13, 39)) {
      result[isWeekend ? "special" : "normal"] = "3.7";
      return result;
    }

    if (isTimeBetween(firstIn, 12, 40, 13, 40)) {
      result[isWeekend ? "special" : "normal"] = "4.3";
      if (!isWeekend && lastOut.totalSeconds >= (20 * 3600 + 50 * 60)) {
        result.overtime = "2.5";
      }
      return result;
    }

    var baseHours = Math.max(0, 8 - calculateLateDeductionHours(firstIn) - calculateEarlyDeductionHours(lastOut));
    if (baseHours > 0) {
      result[isWeekend ? "special" : "normal"] = formatDisplayNumber(baseHours);
    }
    if (!isWeekend && lastOut.totalSeconds >= (20 * 3600 + 50 * 60)) {
      result.overtime = "2.5";
    }
    return result;
  }

  function readFileAsText(file) {
    return new Promise(function (resolve, reject) {
      if (!file) {
        reject(new Error("지문 원본 파일을 먼저 선택해주세요."));
        return;
      }
      var reader = new FileReader();
      reader.onload = function () {
        resolve(String(reader.result || ""));
      };
      reader.onerror = function () {
        reject(new Error("파일을 읽는 중 문제가 생겼습니다."));
      };
      reader.readAsText(file);
    });
  }

  function readFileAsTextWithEncoding(file, encoding) {
    return new Promise(function (resolve, reject) {
      if (!file) {
        reject(new Error("파일을 먼저 선택해주세요."));
        return;
      }
      var reader = new FileReader();
      reader.onload = function () {
        resolve(String(reader.result || ""));
      };
      reader.onerror = function () {
        reject(new Error("파일을 읽는 중 문제가 생겼습니다."));
      };
      reader.readAsText(file, encoding || "utf-8");
    });
  }

  function readFileAsArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      if (!file) {
        reject(new Error("파일을 먼저 선택해주세요."));
        return;
      }
      var reader = new FileReader();
      reader.onload = function () {
        resolve(reader.result);
      };
      reader.onerror = function () {
        reject(new Error("파일을 읽는 중 문제가 생겼습니다."));
      };
      reader.readAsArrayBuffer(file);
    });
  }

  function parseSpreadsheetXmlToGrid(text) {
    var parser = new DOMParser();
    var xml = parser.parseFromString(String(text || ""), "application/xml");
    if (xml.getElementsByTagName("parsererror").length) return [];
    var table = xml.getElementsByTagNameNS("urn:schemas-microsoft-com:office:spreadsheet", "Table")[0];
    if (!table) return [];
    var xmlRows = parseSpreadsheetXmlRows(
      table.getElementsByTagNameNS("urn:schemas-microsoft-com:office:spreadsheet", "Row")
    );
    return xmlRows.map(function (rowObj) {
      var maxCol = 0;
      Object.keys(rowObj || {}).forEach(function (key) {
        var colNum = Number(key);
        if (colNum > maxCol) maxCol = colNum;
      });
      var row = [];
      for (var col = 1; col <= maxCol; col++) {
        row.push(String(rowObj[col] || "").trim());
      }
      return row;
    });
  }

  function parseHtmlTableToGrid(text) {
    var parser = new DOMParser();
    var doc = parser.parseFromString(String(text || ""), "text/html");
    var table = doc.querySelector("table");
    if (!table) return [];
    return Array.prototype.map.call(table.querySelectorAll("tr"), function (tr) {
      return Array.prototype.map.call(tr.querySelectorAll("th,td"), function (cell) {
        return String(cell.textContent || "").replace(/\s+/g, " ").trim();
      });
    });
  }

  function normalizeImportHeaderToken(value) {
    return String(value || "")
      .replace(/\s+/g, "")
      .replace(/[()\-_/]/g, "")
      .toLowerCase();
  }

  function detectStatusHeaderMap(headerRow) {
    var aliases = {
      date: ["월일", "일자", "날짜", "일"],
      client: ["업체", "거래처", "상호", "업체명", "거래처명"],
      code: ["코드", "code", "품번"],
      name: ["품명", "품목", "품명규격", "제품명"],
      qty: ["수량", "수량ea", "ea"],
      price: ["단가", "판매단가", "매입단가"],
      amount: ["금액", "공급가액", "매출금액", "매입금액", "합계"],
      note1: ["비고", "비고1", "적요"],
      note2: ["비고2"],
    };
    var normalizedHeaders = (headerRow || []).map(normalizeImportHeaderToken);
    var map = {};
    Object.keys(aliases).forEach(function (key) {
      var idx = normalizedHeaders.findIndex(function (head) {
        return aliases[key].some(function (alias) {
          return head === normalizeImportHeaderToken(alias) || head.indexOf(normalizeImportHeaderToken(alias)) >= 0;
        });
      });
      if (idx >= 0) map[key] = idx;
    });
    return map;
  }

  function scoreStatusHeaderMap(map) {
    var requiredKeys = ["date", "client", "code", "name", "qty", "price", "amount"];
    var score = 0;
    requiredKeys.forEach(function (key) {
      if (map[key] != null) score++;
    });
    return score;
  }

  function parseStatusRowsFromGrid(grid) {
    if (!grid || !grid.length) return [];
    var bestHeaderIndex = -1;
    var bestHeaderMap = {};
    var bestScore = 0;
    for (var i = 0; i < Math.min(grid.length, 40); i++) {
      var candidateMap = detectStatusHeaderMap(grid[i]);
      var score = scoreStatusHeaderMap(candidateMap);
      if (score > bestScore) {
        bestScore = score;
        bestHeaderIndex = i;
        bestHeaderMap = candidateMap;
      }
    }
    var startRow = bestHeaderIndex >= 0 && bestScore >= 3 ? bestHeaderIndex + 1 : 0;
    var colMap = bestHeaderIndex >= 0 && bestScore >= 3
      ? bestHeaderMap
      : { date: 0, client: 1, code: 2, name: 3, qty: 4, price: 5, amount: 6, note1: 7, note2: 8 };

    var rows = [];
    var emptyStreak = 0;
    for (var rowIndex = startRow; rowIndex < grid.length; rowIndex++) {
      var row = grid[rowIndex] || [];
      var item = {
        date: String(row[colMap.date] || "").trim(),
        client: String(row[colMap.client] || "").trim(),
        code: String(row[colMap.code] || "").trim(),
        name: String(row[colMap.name] || "").trim(),
        qty: String(row[colMap.qty] || "").trim(),
        price: String(row[colMap.price] || "").trim(),
        amount: String(row[colMap.amount] || "").trim(),
        note1: String(row[colMap.note1] || "").trim(),
        note2: String(row[colMap.note2] || "").trim(),
      };
      var hasAny = ["date", "client", "code", "name", "qty", "price", "amount", "note1", "note2"].some(function (key) {
        return !!item[key];
      });
      if (!hasAny) {
        emptyStreak++;
        if (emptyStreak >= 30 && rows.length) break;
        continue;
      }
      emptyStreak = 0;
      if (normalizeImportHeaderToken(item.date) === "월일" && normalizeImportHeaderToken(item.client) === "업체") {
        continue;
      }
      var joined = [item.date, item.client, item.code, item.name, item.qty, item.price, item.amount, item.note1, item.note2].join(" ");
      if (joined.indexOf("\uFFFD") >= 0) continue;
      var weirdMatches = joined.match(/[^\u0020-\u007E\u00A0-\u024F\u1100-\u11FF\u3130-\u318F\uAC00-\uD7A3]/g);
      if (weirdMatches && weirdMatches.length > Math.max(8, Math.floor(joined.length * 0.12))) continue;
      var plausibleCore =
        (/^\d{1,2}\s*월\s*\d{1,2}\s*일$/.test(item.date) || /^\d{1,2}[\/\.-]\d{1,2}$/.test(item.date) || /^\d{4}-\d{2}-\d{2}$/.test(item.date)) ||
        parseCalcNumber(item.qty) != null ||
        parseCalcNumber(item.price) != null ||
        parseCalcNumber(item.amount) != null ||
        /[A-Za-z]{1,}\d{1,}/.test(item.code);
      if (!plausibleCore) continue;
      rows.push(item);
    }
    return rows;
  }

  function parseStatusRowsFromWorkbookBuffer(arrayBuffer) {
    if (!(window.XLSX && window.XLSX.read && window.XLSX.utils && window.XLSX.utils.sheet_to_json)) {
      throw new Error("엑셀 엔진을 불러오지 못했습니다. 잠시 후 다시 시도해주세요.");
    }
    var workbook = window.XLSX.read(arrayBuffer, { type: "array", cellDates: false, raw: false });
    var sheetNames = workbook && workbook.SheetNames ? workbook.SheetNames : [];
    if (!sheetNames.length) return [];
    var mergedRows = [];
    for (var i = 0; i < sheetNames.length; i++) {
      var sheet = workbook.Sheets[sheetNames[i]];
      if (!sheet) continue;
      var grid = window.XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: "",
      });
      var parsed = parseStatusRowsFromGrid(grid);
      if (parsed.length) {
        mergedRows = mergedRows.concat(parsed);
      }
    }
    return mergedRows;
  }

  function parseStatusRowsFromFileText(text) {
    var content = String(text || "");
    var grid = [];
    if (
      /<\?xml/i.test(content) &&
      /schemas-microsoft-com:office:spreadsheet/i.test(content)
    ) {
      grid = parseSpreadsheetXmlToGrid(content);
    } else if (/<table/i.test(content)) {
      grid = parseHtmlTableToGrid(content);
    } else {
      grid = parseClipboardGrid(content);
    }
    return parseStatusRowsFromGrid(grid);
  }

  function appendImportedStatusRows(rows) {
    var incoming = Array.isArray(rows) ? rows : [];
    if (!incoming.length) return 0;
    var start = Math.max(0, getLastUsedRowIndex(ST.rows) + 1);
    var needed = start + incoming.length + 20;
    var nextRows = normalizeSalesRows(ST.rows, Math.max(MIN_ROW_COUNT, needed));
    incoming.forEach(function (row, index) {
      nextRows[start + index] = Object.assign({}, nextRows[start + index] || {}, {
        date: row.date || "",
        client: row.client || "",
        code: row.code || "",
        name: row.name || "",
        qty: row.qty || "",
        price: row.price || "",
        amount: row.amount || "",
        note1: row.note1 || "",
        note2: row.note2 || "",
      });
    });
    ST.rows = nextRows;
    ST.visibleRowsDirty = true;
    return incoming.length;
  }

  function applyFingerprintDataToOutsourceRows(parsedEmployees) {
    var rows = normalizeClosingOutsourceRows(getClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor));
    var previousWarnings = getClosingOutsourceWarnings(closingState.attendanceMonth, closingState.outsourceVendor);
    var warnings = {};
    var monthNumber = getClosingAttendanceMonthNumber();
    var daysInMonth = getClosingDaysInMonth();
    var meta = getClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor);
    var lockDay = getClosingOutsourceLockDay(meta, daysInMonth);
    var matchedGroups = 0;
    var warningCells = 0;
    for (var groupStart = 0; groupStart < rows.length; groupStart += closingOutsourceMarkers.length) {
      var headerRow = rows[groupStart] || {};
      var employeeName = normalizeClosingSearchText(headerRow.employee);
      if (!employeeName) continue;
      var matchedName = "";
      Object.keys(parsedEmployees || {}).some(function (candidate) {
        var normalized = normalizeClosingSearchText(candidate);
        if (normalized === employeeName || normalized.indexOf(employeeName) >= 0 || employeeName.indexOf(normalized) >= 0) {
          matchedName = candidate;
          return true;
        }
        return false;
      });
      if (matchedName) matchedGroups++;
      for (var day = 1; day <= 31; day++) {
        var field = "d" + day;
        var warningKey = groupStart + ":" + field;
        if (day <= lockDay) {
          if (previousWarnings && previousWarnings[warningKey]) {
            warnings[warningKey] = true;
          }
          continue;
        }
        for (var markerOffset = 0; markerOffset < closingOutsourceMarkers.length; markerOffset++) {
          rows[groupStart + markerOffset][field] = "";
        }
        if (!matchedName || day > daysInMonth) continue;
        var record = parsedEmployees[matchedName][monthNumber + "-" + day];
        var calculated = calculateOutsourceRecord(record);
        if (calculated.warning) {
          var warningLimit = getClosingWarningDayLimit();
          if (day > warningLimit || isClosingWeekend(day)) {
            continue;
          }
          warnings[warningKey] = true;
          warningCells++;
          continue;
        }
        rows[groupStart + 0][field] = calculated.normal;
        rows[groupStart + 1][field] = calculated.overtime;
        rows[groupStart + 2][field] = calculated.night;
        rows[groupStart + 3][field] = calculated.special;
      }
      if (matchedName) {
        var block = rows.slice(groupStart, groupStart + closingOutsourceMarkers.length);
        calculateClosingOutsourceWeeklyHolidayHours(block, warnings, groupStart, daysInMonth, true, lockDay);
      }
    }
    setClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor, rows);
    setClosingOutsourceWarnings(closingState.attendanceMonth, closingState.outsourceVendor, warnings);
    return {
      matchedGroups: matchedGroups,
      warningCells: warningCells,
      lockDay: lockDay,
    };
  }

  function applyFingerprintDataToEmployeeRows(parsedEmployees) {
    var rows = normalizeClosingEmployeeRows(getClosingEmployeeRows(closingState.attendanceMonth));
    var periodDays = getClosingEmployeePeriodDays();
    var markerMap = {};
    closingEmployeeMarkers.forEach(function (label, idx) {
      markerMap[label] = idx;
    });
    var now = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var matchedGroups = 0;
    for (var groupStart = 0; groupStart < rows.length; groupStart += closingEmployeeMarkers.length) {
      var headerRow = rows[groupStart] || {};
      var employeeName = normalizeClosingSearchText(headerRow.employee);
      if (!employeeName) continue;
      var matchedName = "";
      Object.keys(parsedEmployees || {}).some(function (candidate) {
        var normalized = normalizeClosingSearchText(candidate);
        if (normalized === employeeName || normalized.indexOf(employeeName) >= 0 || employeeName.indexOf(normalized) >= 0) {
          matchedName = candidate;
          return true;
        }
        return false;
      });
      for (var clearIndex = 0; clearIndex < closingEmployeeMarkers.length; clearIndex++) {
        periodDays.forEach(function (_, slotIndex) {
          rows[groupStart + clearIndex][closingAttendanceDayFields[slotIndex]] = "";
        });
      }
      if (!matchedName) continue;
      matchedGroups++;
      for (var dayIndex = 0; dayIndex < periodDays.length; dayIndex++) {
        var dayInfo = periodDays[dayIndex];
        var field = closingAttendanceDayFields[dayIndex];
        var slotDate = new Date(dayInfo.year, dayInfo.month - 1, dayInfo.day);
        if (slotDate > today) continue;
        var isWeekend = dayInfo.weekday === "토" || dayInfo.weekday === "일";
        var record = parsedEmployees[matchedName][dayInfo.month + "-" + dayInfo.day];
        var firstIn = record ? (record.in1 || record.in2 || record.in3) : null;
        var lastOut = record ? (record.out3 || record.out2 || record.out1) : null;
        if (!firstIn || !lastOut) {
          if (!isWeekend) {
            rows[groupStart + markerMap["결근"]][field] = "1";
          }
          continue;
        }
        if (lastOut.totalSeconds >= (20 * 3600 + 50 * 60)) {
          rows[groupStart + markerMap["야근"]][field] = "1";
        }
        if (isWeekend) {
          rows[groupStart + markerMap["특근"]][field] = "1";
          continue;
        }
        if (firstIn.totalSeconds > (9 * 3600 + 5 * 60 + 59)) {
          rows[groupStart + markerMap["지각"]][field] = "1";
        }
        if (lastOut.totalSeconds < (17 * 3600 + 55 * 60)) {
          rows[groupStart + markerMap["조퇴"]][field] = "1";
        }
      }
    }
    rows = syncClosingEmployeeCounts(rows, periodDays);
    setClosingEmployeeRows(closingState.attendanceMonth, rows);
    return { matchedGroups: matchedGroups };
  }

  function renderClosingOutsourceMatrix() {
    var rows = getClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor);
    var warnings = getClosingOutsourceWarnings(closingState.attendanceMonth, closingState.outsourceVendor);
    var meta = getClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor);
    var daysInMonth = getClosingDaysInMonth();
    var year = getClosingAttendanceYear();
    var monthNumber = getClosingAttendanceMonthNumber();
    var hourlyWage = getClosingOutsourceHourlyWage(meta);
    var wageTotal = 0;
    var headDays = "";
    for (var day = 1; day <= daysInMonth; day++) {
      var weekday = getClosingWeekdayLabel(day);
      var extra = weekday === "일" ? " sunday" : (weekday === "토" ? " saturday" : "");
      headDays +=
        '<th class="closing-matrix-day-head' + extra + '">' +
          '<div class="closing-matrix-weekday">' + escapeHtml(weekday) + '</div>' +
          '<div class="closing-matrix-daynum">' + day + '</div>' +
        '</th>';
    }
    var body = "";
    for (var index = 0; index < rows.length; index += closingOutsourceMarkers.length) {
      var block = rows.slice(index, index + closingOutsourceMarkers.length);
      var headerRow = block[0] || emptyClosingOutsourceRow("", "정상");
      if (!matchesClosingAttendanceSearch(headerRow.employee, headerRow.joinDate)) continue;
      var derived = calculateClosingOutsourceDerived(block, warnings, index, daysInMonth, hourlyWage);
      wageTotal += derived.groupTotal;
      for (var markerIndex = 0; markerIndex < closingOutsourceMarkers.length; markerIndex++) {
        var row = block[markerIndex] || emptyClosingOutsourceRow("", closingOutsourceMarkers[markerIndex]);
        var derivedRow = derived.rows[markerIndex] || {
          timeTotal: calculateClosingOutsourceTimeTotal(row),
          basePay: 0,
          bonusPay: 0,
          payAmount: 0,
        };
        body += '<tr class="closing-employee-matrix-row' + (markerIndex === 0 ? ' closing-employee-group-start' : '') + '">';
        if (markerIndex === 0) {
          body += '<th rowspan="' + closingOutsourceMarkers.length + '" class="closing-matrix-sticky left closing-matrix-name">';
            body += '<input type="text" class="closing-matrix-meta-input" data-outsource-group="' + index + '" data-outsource-meta="employee" value="' + escapeAttr(headerRow.employee || "") + '" />';
          body += '</th>';
          body += '<th rowspan="' + closingOutsourceMarkers.length + '" class="closing-matrix-sticky left closing-matrix-join">';
          body += '<input type="text" class="closing-matrix-meta-input" data-outsource-group="' + index + '" data-outsource-meta="joinDate" value="' + escapeAttr(normalizeJoinDateText(headerRow.joinDate || "")) + '" placeholder="근속 시작일" />';
          body += '</th>';
        }
        body += '<th class="closing-matrix-type">' + escapeHtml(row.type || closingOutsourceMarkers[markerIndex]) + '</th>';
        for (var dayCol = 1; dayCol <= daysInMonth; dayCol++) {
          var field = "d" + dayCol;
          var warningKey = index + ":" + field;
          var weekdayClass = getClosingWeekdayClass(dayCol);
          var normalText = String(row[field] || "").trim();
          var normalNum = parseCalcNumber(normalText);
          var isWarningCell = markerIndex === 0 && !!warnings[warningKey] && !normalText;
          var isCautionCell = markerIndex === 0 && !!normalText && (normalNum == null || normalNum < 8);
          body += '<td class="closing-matrix-cell' + weekdayClass + (isWarningCell ? ' warning' : '') + (isCautionCell ? ' caution' : '') + '">';
          body += '<input type="text" class="closing-matrix-input" data-outsource-row="' + (index + markerIndex) + '" data-outsource-field="' + field + '" value="' + escapeAttr(row[field] || "") + '" />';
          body += '</td>';
        }
        body += '<td class="closing-matrix-summary center">' + escapeHtml(formatDisplayNumber(derivedRow.timeTotal || 0)) + '</td>';
        body += '<td class="closing-matrix-summary right">' + escapeHtml(formatDisplayNumber(derivedRow.basePay || 0)) + '</td>';
        body += '<td class="closing-matrix-summary right">' + escapeHtml(formatDisplayNumber(derivedRow.bonusPay || 0)) + '</td>';
        body += '<td class="closing-matrix-summary right">' + escapeHtml(formatDisplayNumber(derivedRow.payAmount || 0)) + '</td>';
        if (markerIndex === 0) {
          body += '<td rowspan="' + closingOutsourceMarkers.length + '" class="closing-matrix-total right">';
          body += '<div>' + escapeHtml(formatDisplayNumber(derived.groupTotal || 0)) + '</div>';
          body += '<button type="button" class="closing-remove-mini" data-remove-outsource-group="' + index + '">-</button>';
          body += '</td>';
        }
        body += '</tr>';
      }
    }
    var mgmtRate = getClosingOutsourceMgmtRate(closingState.outsourceVendor);
    var mgmtFee = Math.round(wageTotal * mgmtRate);
    var retirement = parseCalcNumber(meta.retirement) || 0;
    var total = wageTotal + mgmtFee + retirement;
    var totalVat = Math.round(total * 1.1);
    return (
      '<div class="closing-matrix-wrap">' +
        '<div class="closing-matrix-summarybar">' +
          '<div class="closing-matrix-badge">' + escapeHtml(year + "년 " + monthNumber + "월 " + getClosingOutsourceVendorLabel() + " 급여 청구내역") + '</div>' +
          '<div class="closing-matrix-badge">시급 <input type="text" class="closing-inline-input right" style="min-width:100px;height:30px" id="closing-outsource-hourly-wage" value="' + escapeAttr(meta.hourlyWage || String(DEFAULT_CLOSING_HOURLY_WAGE_2026_KR)) + '" /></div>' +
          '<div class="closing-matrix-badge strong">총 급여지급액 ' + escapeHtml(formatClosingCurrency(wageTotal)) + ' <span class="muted">(시급 ' + escapeHtml(formatDisplayNumber(hourlyWage)) + '원)</span></div>' +
          '<div class="closing-matrix-badge">관리비(' + escapeHtml(String(Math.round(mgmtRate * 100))) + '%) ' + escapeHtml(formatClosingCurrency(mgmtFee)) + '</div>' +
          '<div class="closing-matrix-badge">퇴직금 <input type="text" class="closing-inline-input" style="min-width:110px;height:30px" id="closing-outsource-retirement" value="' + escapeAttr(meta.retirement || "") + '" /></div>' +
          '<div class="closing-matrix-badge">총 금액 ' + escapeHtml(formatClosingCurrency(total)) + '</div>' +
          '<div class="closing-matrix-badge strong">총 금액(VAT 포함) ' + escapeHtml(formatClosingCurrency(totalVat)) + '</div>' +
        '</div>' +
        '<div class="closing-matrix-table-wrap">' +
          '<table class="closing-matrix-table closing-matrix-table-compact">' +
            '<thead>' +
              '<tr>' +
                '<th class="closing-matrix-sticky left closing-matrix-name" rowspan="2">성명</th>' +
                '<th class="closing-matrix-sticky left closing-matrix-join" rowspan="2">근속 시작</th>' +
                '<th rowspan="2">구분</th>' +
                headDays +
                '<th rowspan="2">시간</th>' +
                '<th rowspan="2">급여계산</th>' +
                '<th rowspan="2">만근수당<br/>소급분</th>' +
                '<th rowspan="2">급여지급액</th>' +
                '<th rowspan="2">총 금액</th>' +
              '</tr>' +
            '</thead>' +
            '<tbody>' + body + '</tbody>' +
          '</table>' +
        '</div>' +
      '</div>'
    );
  }

  function renderClosingEmployeeEntryMatrix() {
    var rows = getClosingEmployeeRows(closingState.attendanceMonth);
    var year = getClosingAttendanceYear();
    var monthNumber = getClosingAttendanceMonthNumber();
    var periodDays = getClosingEmployeePeriodDays();
    var headDays = "";
    periodDays.forEach(function (dayInfo) {
      var extra = dayInfo.weekday === "일" ? " sunday" : (dayInfo.weekday === "토" ? " saturday" : "");
      headDays +=
        '<th class="closing-matrix-day-head' + extra + '">' +
          '<div class="closing-matrix-weekday">' + escapeHtml(dayInfo.weekday) + '</div>' +
          '<div class="closing-matrix-daynum">' + dayInfo.day + '</div>' +
        '</th>';
    });
    var body = "";
    for (var index = 0; index < rows.length; index += closingEmployeeMarkers.length) {
      var block = rows.slice(index, index + closingEmployeeMarkers.length);
      var headerRow = block[0] || emptyClosingEmployeeRow("", "야근");
      if (!matchesClosingAttendanceSearch(headerRow.employee, "")) continue;
      for (var markerIndex = 0; markerIndex < closingEmployeeMarkers.length; markerIndex++) {
        var row = block[markerIndex] || emptyClosingEmployeeRow("", closingEmployeeMarkers[markerIndex]);
        var rowCount = getClosingEmployeeCount(row, periodDays);
        body += '<tr class="closing-matrix-row' + (markerIndex === 0 ? ' closing-outsource-group-start' : '') + '">';
        if (markerIndex === 0) {
          body += '<th rowspan="' + closingEmployeeMarkers.length + '" class="closing-matrix-sticky left closing-matrix-name">';
          body += '<input type="text" class="closing-matrix-meta-input" data-employee-group="' + index + '" data-employee-meta="employee" value="' + escapeAttr(headerRow.employee || "") + '" placeholder="이름" />';
          body += '</th>';
        }
        body += '<th class="closing-matrix-type">' + escapeHtml(row.type || closingEmployeeMarkers[markerIndex]) + '</th>';
        periodDays.forEach(function (dayInfo, slotIndex) {
          var field = closingAttendanceDayFields[slotIndex];
          var extra = dayInfo.weekday === "일" ? " sunday" : (dayInfo.weekday === "토" ? " saturday" : "");
          body += '<td class="closing-matrix-cell' + extra + '">';
          body += '<input type="text" class="closing-matrix-input center" data-employee-row="' + (index + markerIndex) + '" data-employee-field="' + field + '" value="' + escapeAttr(row[field] || "") + '" />';
          body += '</td>';
        });
        body += '<td class="closing-matrix-summary center">' + escapeHtml(String(rowCount)) + '</td>';
        if (markerIndex === 0) {
          body += '<td rowspan="' + closingEmployeeMarkers.length + '" class="closing-matrix-total center">';
          body += '<button type="button" class="closing-remove-mini" data-remove-employee-group="' + index + '">-</button>';
          body += '</td>';
        }
        body += '</tr>';
      }
    }
    return (
      '<div class="closing-matrix-wrap">' +
        '<div class="closing-matrix-summarybar">' +
          '<div class="closing-matrix-badge">' + escapeHtml(year + "년 " + monthNumber + "월 정직원 수기 입력") + '</div>' +
          '<div class="closing-matrix-badge">' + escapeHtml(getClosingEmployeePeriodLabel()) + '</div>' +
          '<div class="closing-matrix-badge">일자별 수기표 (1은 해당 항목 발생)</div>' +
        '</div>' +
        '<div class="closing-matrix-table-wrap">' +
          '<table class="closing-matrix-table closing-matrix-table-compact">' +
            '<thead>' +
              '<tr>' +
                '<th class="closing-matrix-sticky left closing-matrix-name" rowspan="2">성명</th>' +
                '<th rowspan="2">구분</th>' +
                headDays +
                '<th rowspan="2">월합계</th>' +
                '<th rowspan="2">-</th>' +
              '</tr>' +
            '</thead>' +
            '<tbody>' + body + '</tbody>' +
          '</table>' +
        '</div>' +
      '</div>'
    );
  }

  function renderClosingEmployeeSummaryMatrix() {
    var rows = getClosingEmployeeRows(closingState.attendanceMonth);
    var year = getClosingAttendanceYear();
    var monthNumber = getClosingAttendanceMonthNumber();
    var periodDays = getClosingEmployeePeriodDays();
    var body = "";
    var visibleCount = 0;
    for (var index = 0; index < rows.length; index += closingEmployeeMarkers.length) {
      var block = rows.slice(index, index + closingEmployeeMarkers.length);
      var headerRow = block[0] || emptyClosingEmployeeRow("", "야근");
      if (!matchesClosingAttendanceSearch(headerRow.employee, "")) continue;
      visibleCount++;
      var counts = {};
      for (var markerIndex = 0; markerIndex < closingEmployeeMarkers.length; markerIndex++) {
        var row = block[markerIndex] || emptyClosingEmployeeRow("", closingEmployeeMarkers[markerIndex]);
        counts[row.type || closingEmployeeMarkers[markerIndex]] = String(getClosingEmployeeCount(row, periodDays));
      }
      body += '<div class="closing-employee-card">';
      body += '<div class="closing-employee-card-head">';
      body += '<div class="closing-employee-name-readonly">' + escapeHtml(headerRow.employee || "이름") + '</div>';
      body += '</div>';
      body += '<table class="closing-employee-mini-table"><tbody>';
      body += '<tr><th>야근</th><td class="right">' + escapeHtml(String(counts["야근"] || "0")) + '</td></tr>';
      body += '<tr><th>특근</th><td class="right">' + escapeHtml(String(counts["특근"] || "0")) + '</td></tr>';
      body += '<tr><th>지각</th><td class="right">' + escapeHtml(String(counts["지각"] || "0")) + '</td></tr>';
      body += '<tr><th>조퇴</th><td class="right">' + escapeHtml(String(counts["조퇴"] || "0")) + '</td></tr>';
      body += '<tr><th>결근</th><td class="right">' + escapeHtml(String(counts["결근"] || "0")) + '</td></tr>';
      body += '</tbody></table>';
      body += '</div>';
    }
    return (
      '<div class="closing-matrix-wrap">' +
        '<div class="closing-matrix-summarybar">' +
          '<div class="closing-matrix-badge">' + escapeHtml(year + "년 " + monthNumber + "월 정직원 연동 화면") + '</div>' +
          '<div class="closing-matrix-badge">' + escapeHtml(getClosingEmployeePeriodLabel()) + '</div>' +
          '<div class="closing-matrix-badge">야근 / 특근 / 지각 / 조퇴 / 결근 월 합계</div>' +
        '</div>' +
        '<div class="closing-employee-grid closing-employee-grid-compact">' + (visibleCount ? body : '<div class="muted">표시할 데이터가 없습니다.</div>') + '</div>' +
      '</div>'
    );
  }

  function todayIsoString() {
    var now = new Date();
    var year = now.getFullYear();
    var month = String(now.getMonth() + 1).padStart(2, "0");
    var day = String(now.getDate()).padStart(2, "0");
    return year + "-" + month + "-" + day;
  }

  function plusDaysIsoString(days) {
    var now = new Date();
    now.setDate(now.getDate() + (days || 0));
    var year = now.getFullYear();
    var month = String(now.getMonth() + 1).padStart(2, "0");
    var day = String(now.getDate()).padStart(2, "0");
    return year + "-" + month + "-" + day;
  }

  function defaultWorkDateIsoString() {
    return plusDaysIsoString(1);
  }

  function isoToMonthDayLabel(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) return "";
    var parts = String(value).split("-");
    return Number(parts[1]) + "월 " + Number(parts[2]) + "일";
  }

  function monthDayMatchesIso(value, isoDate) {
    var parsed = parseSalesMonthDay(value);
    if (!parsed || !/^\d{4}-\d{2}-\d{2}$/.test(String(isoDate || ""))) return false;
    var parts = String(isoDate).split("-");
    return parsed.month === Number(parts[1]) && parsed.day === Number(parts[2]);
  }

  function showAppToast(message, type) {
    var layer = document.getElementById("app-toast-layer");
    if (!layer) {
      layer = document.createElement("div");
      layer.id = "app-toast-layer";
      layer.className = "app-toast-layer";
      document.body.appendChild(layer);
    }
    layer.innerHTML = '<div class="app-toast' + (type ? " " + type : "") + '">' + escapeHtml(message || "") + "</div>";
    if (appToastTimer) clearTimeout(appToastTimer);
    appToastTimer = setTimeout(function () {
      if (layer) layer.innerHTML = "";
    }, 2600);
  }

  function emptySupplierInfo() {
    return {
      company: "",
      ceo: "",
      businessNo: "",
      address: "",
      businessType: "",
      businessItem: "",
    };
  }

  function defaultCorporateSupplierInfo() {
    return {
      company: "주식회사 한성엔터프라이즈",
      ceo: "김두식",
      businessNo: "361-81-01922",
      address: "경기도 부천시 석천로 380번길 17-7",
      businessType: "제조",
      businessItem: "전자부품 제조업",
    };
  }

  function defaultPersonalSupplierInfo() {
    return {
      company: "한성엔터프라이즈",
      ceo: "김두식",
      businessNo: "128-22-30482",
      address: "경기도 부천시 석천로 380번길 17-7",
      businessType: "제조",
      businessItem: "전자부품, 임가공",
    };
  }

  function getDefaultSupplierInfoForMode(mode) {
    return mode === "personal"
      ? defaultPersonalSupplierInfo()
      : defaultCorporateSupplierInfo();
  }

  function resetSupplierProfileToDefault(mode) {
    var info = ensureWorkInfoShape();
    var normalizedMode = mode === "personal" ? "personal" : "corporate";
    info.supplierProfiles[normalizedMode] = Object.assign(
      emptySupplierInfo(),
      getDefaultSupplierInfoForMode(normalizedMode)
    );
  }

  function ensureWorkInfoShape() {
    if (!workState.info || typeof workState.info !== "object") {
      workState.info = {};
    }

    var info = workState.info;
    if (!info.supplierProfiles || typeof info.supplierProfiles !== "object") {
      info.supplierProfiles = {
        corporate: defaultCorporateSupplierInfo(),
        personal: defaultPersonalSupplierInfo(),
      };
    }

    info.supplierProfiles.corporate = Object.assign(
      defaultCorporateSupplierInfo(),
      info.supplierProfiles.corporate || {}
    );
    info.supplierProfiles.personal = Object.assign(
      defaultPersonalSupplierInfo(),
      info.supplierProfiles.personal || {}
    );

    if (info.supplier && typeof info.supplier === "object") {
      var hasCorporateData = Object.keys(info.supplierProfiles.corporate).some(function (key) {
        return info.supplierProfiles.corporate[key] != null &&
          String(info.supplierProfiles.corporate[key]).trim() !== "";
      });
      if (!hasCorporateData) {
        info.supplierProfiles.corporate = Object.assign(emptySupplierInfo(), info.supplier);
      }
    }

    if (info.supplierMode !== "personal" && info.supplierMode !== "corporate") {
      info.supplierMode = "corporate";
    }

    if (!info.date || !/^\d{4}-\d{2}-\d{2}$/.test(String(info.date))) {
      info.date = defaultWorkDateIsoString();
    }

    return info;
  }

  function getCurrentSupplierMode() {
    return ensureWorkInfoShape().supplierMode;
  }

  function normalizeSupplierMode(value) {
    if (value === "personal") return "personal";
    if (value === "corporate") return "corporate";
    return "";
  }

  function parseSupplierModeLabel(value) {
    var text = String(value || "").trim().toLowerCase();
    if (text === "개인" || text === "personal") return "personal";
    if (text === "법인" || text === "corporate") return "corporate";
    return "";
  }

  function getCurrentSupplierInfo() {
    var info = ensureWorkInfoShape();
    return info.supplierProfiles[getCurrentSupplierMode()];
  }

  function emptyClientRow() {
    return {
      supplierMode: "",
      company: "",
      businessNo: "",
      ceoName: "",
      address: "",
      businessType: "",
      businessItem: "",
    };
  }

  function cloneClientRows(rows) {
    if (!Array.isArray(rows)) return [];
    return rows.map(function (row) {
      return Object.assign(emptyClientRow(), row || {});
    });
  }

  function ensureClientRows(rows, minCount) {
    var normalized = cloneClientRows(rows);
    var count = Math.max(minCount || 20, normalized.length);
    for (var i = 0; i < count; i++) {
      normalized[i] = Object.assign(emptyClientRow(), normalized[i] || {});
    }
    return normalized;
  }

  function parseClipboardGrid(text) {
    var normalized = String(text || "").replace(/\r/g, "");
    var lines = normalized.split("\n");
    if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
    return lines.map(function (line) {
      return line.split("\t");
    });
  }

  function pasteGridIntoWorkItems(text, startRow, startField) {
    var grid = parseClipboardGrid(text);
    if (!grid.length) return false;
    var startCol = workGridFields.indexOf(startField);
    if (startCol < 0) return false;

    var rows = cloneItems(workState.items);
    var needed = startRow + grid.length;
    while (rows.length < needed) rows.push({});

    grid.forEach(function (line, rowOffset) {
      var rowIndex = startRow + rowOffset;
      if (!rows[rowIndex]) rows[rowIndex] = {};
      line.forEach(function (value, colOffset) {
        var field = workGridFields[startCol + colOffset];
        if (!field) return;
        rows[rowIndex][field] = value;
      });
    });

    workState.items = rows;
    scheduleLedgerDraftSave();
    render();
    return true;
  }

  function pasteGridIntoClientRows(text, startRow, startField) {
    var grid = parseClipboardGrid(text);
    if (!grid.length) return false;

    var startCol = clientGridFields.indexOf(startField);
    if (startCol < 0) return false;

    if (startField === "supplierMode") {
      var firstMode = grid[0] && grid[0].length ? parseSupplierModeLabel(grid[0][0]) : "";
      if (!firstMode) startCol = 1;
    }

    var rows = ensureClientRows(clientState.rows, startRow + grid.length + 1);
    grid.forEach(function (line, rowOffset) {
      var rowIndex = startRow + rowOffset;
      if (!rows[rowIndex]) rows[rowIndex] = emptyClientRow();
      line.forEach(function (value, colOffset) {
        var field = clientGridFields[startCol + colOffset];
        if (!field) return;
        rows[rowIndex][field] =
          field === "supplierMode"
            ? (parseSupplierModeLabel(value) || rows[rowIndex][field] || "corporate")
            : value;
      });
    });

    clientState.rows = ensureClientRows(rows, rows.length);
    scheduleLedgerDraftSave();
    render();
    return true;
  }

  function rowHasAnyValue(row, keys) {
    if (!row) return false;
    return (keys || Object.keys(row)).some(function (key) {
      return row[key] != null && String(row[key]).trim() !== "";
    });
  }

  function normalizeWorkItemsRows(rows, minCount) {
    var normalized = Array.isArray(rows)
      ? rows.map(function (row) { return Object.assign({}, row || {}); })
      : [];
    var targetCount = Math.max(minCount || 12, normalized.length);
    for (var i = 0; i < targetCount; i++) {
      normalized[i] = Object.assign({}, normalized[i] || {});
    }
    while (
      normalized.length < targetCount ||
      rowHasAnyValue(normalized[normalized.length - 1], workGridFields)
    ) {
      normalized.push({});
    }
    return normalized;
  }

  function trimTrailingRows(rows, dataFields, minCount, factory) {
    var normalized = Array.isArray(rows)
      ? rows.map(function (row) { return Object.assign({}, row || {}); })
      : [];
    var lastUsed = -1;
    for (var i = normalized.length - 1; i >= 0; i--) {
      if (rowHasAnyValue(normalized[i], dataFields)) {
        lastUsed = i;
        break;
      }
    }
    var target = Math.max(minCount || 0, lastUsed + 2);
    if (!target) return [];
    normalized = normalized.slice(0, target);
    while (normalized.length < target) {
      normalized.push(factory());
    }
    return normalized;
  }

  function normalizeClientGridRows(rows, minCount) {
    var normalized = trimTrailingRows(rows, clientDataFields, minCount || 20, emptyClientRow);
    normalized = ensureClientRows(normalized, minCount || 20);
    return normalized.map(function (row) {
      var hasAnyClientValue = rowHasAnyValue(row, clientDataFields);
      var hasClientInfo = rowHasAnyValue(row, ["company", "businessNo", "ceoName", "address", "businessType", "businessItem"]);
      var normalizedMode = normalizeSupplierMode(row.supplierMode);
      row.supplierMode = hasAnyClientValue
        ? (normalizedMode || (hasClientInfo ? "corporate" : ""))
        : "";
      return row;
    });
  }

  function normalizePriceGridRows(rows, minCount) {
    var normalized = Array.isArray(rows)
      ? rows.map(function (row) { return Object.assign(emptyPriceRow(), row || {}); })
      : [];
    var lastUsed = -1;
    for (var i = normalized.length - 1; i >= 0; i--) {
      if (rowHasAnyValue(normalized[i], priceGridFields)) {
        lastUsed = i;
        break;
      }
    }
    var target = Math.max(minCount || 40, lastUsed + 21);
    for (var i = 0; i < target; i++) {
      normalized[i] = Object.assign(emptyPriceRow(), normalized[i] || {});
      normalized[i].client = getDisplayCompanyName(normalized[i].client);
    }
    while (normalized.length < target) {
      normalized.push(emptyPriceRow());
    }
    return normalized;
  }

  function normalizePriceEditorRows(rows, minCount) {
    var normalized = Array.isArray(rows)
      ? rows.map(function (row) {
          return {
            code: row && row.code != null ? String(row.code) : "",
            price: row && row.price != null ? String(row.price) : "",
            name: row && row.name != null ? String(row.name) : "",
          };
        })
      : [];
    var lastUsed = -1;
    for (var i = normalized.length - 1; i >= 0; i--) {
      if (rowHasAnyValue(normalized[i], priceEditorFields)) {
        lastUsed = i;
        break;
      }
    }
    var target = Math.max(minCount || 40, lastUsed + 21);
    for (var j = 0; j < target; j++) {
      normalized[j] = Object.assign({ code: "", price: "", name: "" }, normalized[j] || {});
    }
    while (normalized.length < target) {
      normalized.push({ code: "", price: "", name: "" });
    }
    return normalized;
  }

  function createLedgerBundle() {
    return {
      rows: normalizeSalesRows([], MIN_ROW_COUNT),
      clientState: {
        rows: normalizeClientGridRows([], 20),
      },
      priceState: {
        rows: normalizePriceGridRows([], 40),
        activeClient: "",
        sidebarScrollTop: 0,
      },
      manageState: {
        startMonth: "",
        endMonth: "",
        client: "",
      },
      filter: {
        colKey: "all",
        keyword: "",
        clientType: "all",
      },
      colWidths: salesColumns.map(function (col) {
        return col.width || 100;
      }),
      rowHeights: [],
    };
  }

  function ensureLedgerBundles() {
    if (!salesLedgerBundle) {
      salesLedgerBundle = {
        rows: ST.rows,
        clientState: clientState,
        priceState: priceState,
        manageState: manageState,
        filter: ST.filter,
        colWidths: ST.colWidths,
        rowHeights: ST.rowHeights,
      };
    }
    if (!purchaseLedgerBundle) {
      purchaseLedgerBundle = createLedgerBundle();
    }
  }

  function getLedgerBundle(kind) {
    ensureLedgerBundles();
    return kind === "purchase" ? purchaseLedgerBundle : salesLedgerBundle;
  }

  function bindLedgerBundle(bundle) {
    ST.rows = bundle.rows;
    clientState = bundle.clientState;
    priceState = bundle.priceState;
    manageState = bundle.manageState;
    ST.filter = bundle.filter;
    ST.colWidths = bundle.colWidths;
    ST.rowHeights = bundle.rowHeights;
    ST.tableBuilt = false;
    ST.host = null;
    ST.scrollEl = null;
    ST.inputEl = null;
    ST.autocompleteEl = null;
    ST.visibleRowsDirty = true;
    ST.renderRange = { start: -1, end: -1 };
    ST.selectedCell = { row: 0, col: 0 };
    ST.selectedKeys = ["0:0"];
    ST.editMode = false;
    ST.isComposing = false;
  }

  function switchLedgerModule(kind, force) {
    var nextKind = kind === "purchase" ? "purchase" : "sales";
    ensureLedgerBundles();
    if (!force && activeLedgerKind === nextKind) return;
    var bundle = getLedgerBundle(nextKind);
    activeLedgerKind = nextKind;
    bindLedgerBundle(bundle);
  }

  function syncActiveLedgerBundle() {
    var bundle = getLedgerBundle(activeLedgerKind);
    bundle.rows = ST.rows;
    bundle.clientState = clientState;
    bundle.priceState = priceState;
    bundle.manageState = manageState;
    bundle.filter = ST.filter;
    bundle.colWidths = ST.colWidths;
    bundle.rowHeights = ST.rowHeights;
  }

  function refreshWorkSummaryUi() {
    var items = (workState.items || []).filter(function (item) {
      return rowHasAnyValue(item, workGridFields);
    });
    var totals = getStatementTotals(items);
    var countEl = document.getElementById("work-summary-count");
    var supplyEl = document.getElementById("work-summary-supply");
    var taxEl = document.getElementById("work-summary-tax");
    var grandEl = document.getElementById("work-summary-grand");
    var qtyEl = document.getElementById("work-summary-qty");
    if (countEl) countEl.value = String(items.length);
    if (supplyEl) supplyEl.textContent = formatDisplayNumber(totals.supply);
    if (taxEl) taxEl.textContent = formatDisplayNumber(totals.tax);
    if (grandEl) grandEl.textContent = formatDisplayNumber(totals.grand);
    if (qtyEl) qtyEl.textContent = formatDisplayNumber(totals.qty);
  }

  function createMiniSheetEngine(options) {
    var engine = {
      host: null,
      scrollEl: null,
      inputEl: null,
      selectedCell: { row: 0, col: 0 },
      selectedKeys: ["0:0"],
      prevSelectedKeys: [],
      prevActiveKey: null,
      dragging: false,
      dragStart: null,
      dragMode: "replace",
      dragBaseKeys: [],
      editMode: false,
      isComposing: false,
      editOrigin: { row: 0, col: 0, value: "" },
      editSnapshotRows: null,
      undoStack: [],
      redoStack: [],
      colWidths: [],
      rowHeights: [],
      resizing: null,
      globalListenersBound: false,
      cellMap: {},
      rowHeadEls: [],
      colHeadEls: [],
    };

    function getColumns() {
      return options.columns;
    }

    function getRows() {
      return options.getRows();
    }

    function cloneRowList(rows) {
      return (rows || []).map(function (row) {
        return Object.assign({}, row || {});
      });
    }

    function getRowCount() {
      return getRows().length;
    }

    function getColCount() {
      return getColumns().length;
    }

    function getColWidth(index) {
      return engine.colWidths[index] || getColumns()[index].width || 100;
    }

    function getRowHeight(index) {
      return engine.rowHeights[index] || 28;
    }

    function ensureRowCount(minCount) {
      var current = getRows();
      var normalized = options.normalizeRows
        ? options.normalizeRows(cloneRowList(current), minCount)
        : cloneRowList(current);
      if (normalized.length !== current.length) {
        options.setRows(normalized);
        if (typeof options.onRowsChange === "function") options.onRowsChange(normalized);
      }
    }

    function getRawValue(r, c) {
      var rows = getRows();
      var key = getColumns()[c].key;
      return rows[r] && rows[r][key] != null ? String(rows[r][key]) : "";
    }

    function getDisplayValue(r, c) {
      var raw = getRawValue(r, c);
      if (typeof options.formatValue === "function") {
        return options.formatValue(r, c, raw, getRows());
      }
      return raw;
    }

    function setRawValue(rows, r, c, value) {
      var key = getColumns()[c].key;
      var nextValue = value == null ? "" : String(value);
      if (typeof options.normalizeValue === "function") {
        nextValue = options.normalizeValue(r, c, nextValue, rows);
      }
      rows[r] = Object.assign({}, rows[r] || {}, (function () {
        var o = {};
        o[key] = nextValue;
        return o;
      })());
      if (typeof options.afterCellChange === "function") {
        options.afterCellChange(rows, r, c);
      }
    }

    function setRows(nextRows, skipHistory, beforeRows, skipRefresh) {
      var normalized = options.normalizeRows
        ? options.normalizeRows(cloneRowList(nextRows))
        : cloneRowList(nextRows);
      if (!skipHistory) {
        engine.undoStack.push(beforeRows || cloneRowList(getRows()));
        if (engine.undoStack.length > 200) engine.undoStack.shift();
        engine.redoStack = [];
      }
      options.setRows(normalized);
      if (typeof options.onRowsChange === "function") {
        options.onRowsChange(normalized);
      }
      if (!skipRefresh) refreshGridValues();
    }

    function snapshotEditOrigin() {
      engine.editOrigin = {
        row: engine.selectedCell.row,
        col: engine.selectedCell.col,
        value: getRawValue(engine.selectedCell.row, engine.selectedCell.col),
      };
    }

    function syncEditOrigin() {
      if (!engine.editMode) snapshotEditOrigin();
    }

    function selectionBounds() {
      return rectFromKeys(engine.selectedKeys);
    }

    function updateStatusBar() {
      if (!engine.host) return;
      var addr = document.getElementById(options.idPrefix + "-status-cell");
      var range = document.getElementById(options.idPrefix + "-status-range");
      var count = document.getElementById(options.idPrefix + "-status-count");
      var sum = document.getElementById(options.idPrefix + "-status-sum");
      var avg = document.getElementById(options.idPrefix + "-status-avg");
      if (addr) {
        addr.textContent = columnLabel(engine.selectedCell.col) + (engine.selectedCell.row + 1);
      }
      if (range) {
        var bounds = selectionBounds();
        range.textContent =
          columnLabel(bounds.left) + (bounds.top + 1) +
          ":" +
          columnLabel(bounds.right) + (bounds.bottom + 1);
      }
      if (count) count.textContent = String(engine.selectedKeys.length);
      if (sum || avg) {
        var values = engine.selectedKeys
          .map(function (key) {
            var parts = key.split(":");
            return parseCalcNumber(getRawValue(Number(parts[0]), Number(parts[1])));
          })
          .filter(function (value) { return value != null; });
        var total = values.reduce(function (acc, value) { return acc + value; }, 0);
        if (sum) sum.textContent = values.length ? formatDisplayNumber(total) : "-";
        if (avg) avg.textContent = values.length ? formatDisplayNumber(total / values.length) : "-";
      }
    }

    function buildTableHtml() {
      var rows = getRows();
      var html =
        '<div class="st-wrap">' +
          '<div class="st-toolbar"><strong>' + escapeHtml(options.title || "") + '</strong><span>' + escapeHtml(options.subtitle || "") + '</span></div>' +
          '<div class="st-scroll" tabindex="0" style="max-height:' + (options.maxHeight || 420) + 'px">' +
            '<table class="st-table"><thead><tr>' +
              '<th class="st-corner" data-corner="1"></th>';
      getColumns().forEach(function (col, index) {
        html +=
          '<th class="st-col-head" data-col-head="' + index + '" style="width:' + getColWidth(index) + 'px;min-width:' + getColWidth(index) + 'px">' +
            escapeHtml(col.label) +
            '<div class="st-resizer-col" data-mini-col-resizer="' + index + '"></div>' +
          '</th>';
      });
      html += '</tr></thead><tbody>';
      for (var r = 0; r < rows.length; r++) {
        html += '<tr data-row="' + r + '" style="height:' + getRowHeight(r) + 'px">';
        html +=
          '<th class="st-row-head" data-row-head="' + r + '">' +
            (r + 1) +
            '<div class="st-resizer-row" data-mini-row-resizer="' + r + '"></div>' +
          '</th>';
        for (var c = 0; c < getColCount(); c++) {
          var active = r === engine.selectedCell.row && c === engine.selectedCell.col;
          var displayHeight = Math.max(24, getRowHeight(r) - 4);
          html += '<td data-r="' + r + '" data-c="' + c + '" class="' + (active ? 'sel-in sel-active' : '') + '">';
          html += '<div class="st-cell-inner" style="min-height:' + displayHeight + 'px">';
          html +=
            '<div class="st-display" style="' +
              (active ? 'display:none;' : '') +
              'height:' + displayHeight + 'px;line-height:' + displayHeight + 'px;">' +
              escapeHtml(getDisplayValue(r, c)) +
            '</div>';
          html += '<div class="st-ghost" style="display:none;height:' + displayHeight + 'px;line-height:' + displayHeight + 'px;"></div>';
          html += '</div></td>';
        }
        html += '</tr>';
      }
      html +=
            '</tbody></table>' +
          '</div>' +
          '<div class="st-statusbar">' +
            '<div class="st-status-main">' +
              '<span class="st-status-pill">현재셀 <strong id="' + options.idPrefix + '-status-cell">A1</strong></span>' +
              '<span class="st-status-pill">범위 <strong id="' + options.idPrefix + '-status-range">A1:A1</strong></span>' +
              '<span class="st-status-pill">개수 <strong id="' + options.idPrefix + '-status-count">1</strong></span>' +
            '</div>' +
            '<div class="st-status-metrics">' +
              '<span class="st-status-pill">합계 <strong id="' + options.idPrefix + '-status-sum">-</strong></span>' +
              '<span class="st-status-pill">평균 <strong id="' + options.idPrefix + '-status-avg">-</strong></span>' +
            '</div>' +
          '</div>' +
        '</div>';
      return html;
    }

    function rebuildCellMap() {
      if (!engine.host) return;
      engine.cellMap = {};
      engine.host.querySelectorAll("td[data-r][data-c]").forEach(function (td) {
        engine.cellMap[cellKey(Number(td.getAttribute("data-r")), Number(td.getAttribute("data-c")))] = td;
      });
      engine.rowHeadEls = Array.prototype.slice.call(engine.host.querySelectorAll("[data-row-head]"));
      engine.colHeadEls = Array.prototype.slice.call(engine.host.querySelectorAll("[data-col-head]"));
    }

    function refreshGridValues() {
      if (!engine.host) return;
      engine.host.innerHTML = buildTableHtml();
      engine.scrollEl = engine.host.querySelector(".st-scroll");
      rebuildCellMap();

      engine.inputEl = document.createElement("input");
      engine.inputEl.type = "text";
      engine.inputEl.setAttribute("spellcheck", "false");
      engine.inputEl.className = "st-input no-edit";
      moveInputToCell(engine.selectedCell.row, engine.selectedCell.col);

      engine.scrollEl.addEventListener("keydown", onTableKeyDown);
      engine.scrollEl.addEventListener("copy", onCopy);
      engine.scrollEl.addEventListener("paste", onPaste);

      engine.host.querySelector("tbody").addEventListener("mousedown", onCellMouseDown);
      engine.host.querySelector("tbody").addEventListener("dblclick", onCellDblClick);
      engine.host.querySelector("thead").addEventListener("mousedown", onHeadMouseDown);
      engine.host.querySelector("thead").addEventListener("dblclick", onHeadDoubleClick);

      engine.inputEl.addEventListener("compositionstart", function () {
        engine.isComposing = true;
        if (!engine.editMode) {
          engine.editSnapshotRows = cloneRowList(getRows());
          engine.editMode = true;
          syncInputOverlay();
        }
      });
      engine.inputEl.addEventListener("compositionend", function (e) {
        engine.isComposing = false;
        setValue(engine.selectedCell.row, engine.selectedCell.col, e.target.value);
        syncInputOverlay();
      });
      engine.inputEl.addEventListener("input", function (e) {
        if (!engine.editMode) {
          engine.editSnapshotRows = cloneRowList(getRows());
          engine.editMode = true;
        }
        setValue(engine.selectedCell.row, engine.selectedCell.col, e.target.value);
        syncInputOverlay();
      });
      engine.inputEl.addEventListener("blur", function () {
        commitActiveCellValue();
      });
      engine.inputEl.addEventListener("keydown", onInputKeyDown);
      engine.inputEl.addEventListener("copy", onCopy);
      engine.inputEl.addEventListener("paste", onPaste);
      engine.inputEl.addEventListener("cut", function (e) {
        onCopy(e);
        clearSelectionCells();
      });
      engine.inputEl.addEventListener("focus", function (e) {
        var len = e.target.value.length;
        if (engine.editMode) e.target.setSelectionRange(len, len);
        else e.target.setSelectionRange(0, len);
      });
      engine.inputEl.addEventListener("mouseup", function (e) {
        if (!engine.editMode) {
          e.preventDefault();
          var len = e.target.value.length;
          e.target.setSelectionRange(0, len);
        }
      });
      engine.inputEl.addEventListener("select", function (e) {
        if (!engine.editMode) {
          var len = e.target.value.length;
          e.target.setSelectionRange(0, len);
        }
      });
      updateSelectionUI();
      syncInputOverlay();
      focusInput();
    }

    function moveInputToCell(r, c) {
      if (!engine.host || !engine.inputEl) return;
      var prevTd = engine.inputEl.closest("td");
      if (prevTd) {
        var prevDisp = prevTd.querySelector(".st-display");
        var prevGhost = prevTd.querySelector(".st-ghost");
        if (prevDisp) {
          prevDisp.style.display = "";
          prevDisp.textContent = getDisplayValue(
            Number(prevTd.getAttribute("data-r")),
            Number(prevTd.getAttribute("data-c"))
          );
        }
        if (prevGhost) {
          prevGhost.textContent = "";
          prevGhost.style.display = "none";
        }
      }
      var td = engine.host.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
      if (!td) return;
      var inner = td.querySelector(".st-cell-inner");
      var disp = td.querySelector(".st-display");
      if (disp) disp.style.display = "none";
      inner.appendChild(engine.inputEl);
      engine.inputEl.value = getRawValue(r, c);
      syncInputOverlay();
    }

    function syncInputOverlay() {
      if (!engine.inputEl || !engine.host) return;
      var r = engine.selectedCell.row;
      var c = engine.selectedCell.col;
      var td = engine.host.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
      if (!td) return;
      var inner = td.querySelector(".st-cell-inner");
      var ghost = inner.querySelector(".st-ghost");
      var raw = getRawValue(r, c);
      engine.inputEl.value = raw;
      if (engine.editMode) {
        engine.inputEl.classList.remove("no-edit");
        if (ghost) {
          ghost.textContent = "";
          ghost.style.display = "none";
        }
      } else {
        engine.inputEl.classList.add("no-edit");
        if (ghost) {
          ghost.textContent = getDisplayValue(r, c);
          ghost.style.display = "block";
        }
      }
      var displayHeight = Math.max(24, getRowHeight(r) - 4);
      engine.inputEl.style.height = displayHeight + "px";
      engine.inputEl.style.lineHeight = displayHeight + "px";
    }

    function focusInput() {
      if (!engine.inputEl) return;
      requestAnimationFrame(function () {
        if (!engine.inputEl) return;
        if (engine.scrollEl) engine.scrollEl.focus();
        engine.inputEl.focus();
        var len = engine.inputEl.value.length;
        if (engine.editMode) engine.inputEl.setSelectionRange(len, len);
        else engine.inputEl.setSelectionRange(0, len);
      });
    }

    function commitActiveCellValue() {
      var row = engine.selectedCell.row;
      var col = engine.selectedCell.col;
      var current = getRawValue(row, col);
      var normalized = current;
      if (typeof options.normalizeValue === "function") {
        normalized = options.normalizeValue(row, col, current, getRows());
      }
      if (normalized === current) {
        if (
          engine.editSnapshotRows &&
          JSON.stringify(engine.editSnapshotRows) !== JSON.stringify(getRows())
        ) {
          engine.undoStack.push(engine.editSnapshotRows);
          if (engine.undoStack.length > 200) engine.undoStack.shift();
          engine.redoStack = [];
        }
        engine.editSnapshotRows = null;
        return;
      }
      var nextRows = cloneRowList(getRows());
      setRawValue(nextRows, row, col, normalized);
      setRows(nextRows, false, engine.editSnapshotRows || cloneRowList(getRows()));
      engine.editSnapshotRows = null;
    }

    function setValue(r, c, value) {
      if (getRawValue(r, c) === value) return;
      var nextRows = cloneRowList(getRows());
      setRawValue(nextRows, r, c, value);
      setRows(nextRows, true, null, true);
      updateStatusBar();
    }

    function clearSelectionCells() {
      var nextRows = cloneRowList(getRows());
      engine.selectedKeys.forEach(function (key) {
        var parts = key.split(":");
        setRawValue(nextRows, Number(parts[0]), Number(parts[1]), "");
      });
      setRows(nextRows, false);
      engine.editMode = false;
      updateSelectionUI();
      moveInputToCell(engine.selectedCell.row, engine.selectedCell.col);
    }

    function getSelectionText() {
      if (!engine.selectedKeys.length) return "";
      var bounds = selectionBounds();
      var lines = [];
      for (var r = bounds.top; r <= bounds.bottom; r++) {
        var cols = [];
        for (var c = bounds.left; c <= bounds.right; c++) {
          cols.push(getDisplayValue(r, c));
        }
        lines.push(cols.join("\t"));
      }
      return lines.join("\n");
    }

    function pasteTextIntoSelection(text) {
      if (!text) return;
      var grid = parseClipboardGrid(text);
      if (!grid.length) return;
      var startRow = engine.selectedCell.row;
      var startCol = engine.selectedCell.col;
      var selectionRect = selectionBounds();
      var selectionIsRect = engine.selectedKeys.length ===
        (selectionRect.bottom - selectionRect.top + 1) * (selectionRect.right - selectionRect.left + 1);
      var sourceRows = Math.max(grid.length, 1);
      var sourceCols = Math.max.apply(null, grid.map(function (line) { return line.length; }));
      ensureRowCount(Math.max(getRowCount(), startRow + sourceRows + 1, selectionRect.bottom + 2));
      var nextRows = cloneRowList(getRows());

      if (engine.selectedKeys.length > 1 && sourceRows === 1 && sourceCols === 1) {
        engine.selectedKeys.forEach(function (key) {
          var parts = key.split(":");
          setRawValue(nextRows, Number(parts[0]), Number(parts[1]), grid[0][0]);
        });
      } else if (engine.selectedKeys.length > 1 && selectionIsRect) {
        for (var r = selectionRect.top; r <= selectionRect.bottom; r++) {
          for (var c = selectionRect.left; c <= selectionRect.right; c++) {
            var rowOffset = r - selectionRect.top;
            var colOffset = c - selectionRect.left;
            var rowValues = grid[rowOffset % sourceRows] || [""];
            setRawValue(nextRows, r, c, rowValues[colOffset % rowValues.length] || "");
          }
        }
      } else {
        grid.forEach(function (line, rowOffset) {
          line.forEach(function (value, colOffset) {
            var rowIndex = startRow + rowOffset;
            var colIndex = startCol + colOffset;
            if (colIndex < getColCount()) {
              if (rowIndex >= nextRows.length) nextRows.push(options.emptyRow());
              setRawValue(nextRows, rowIndex, colIndex, value);
            }
          });
        });
      }

      setRows(nextRows, false);
      var endRow = selectionIsRect && engine.selectedKeys.length > 1
        ? selectionRect.bottom
        : Math.min(getRowCount() - 1, startRow + sourceRows - 1);
      var endCol = selectionIsRect && engine.selectedKeys.length > 1
        ? selectionRect.right
        : Math.min(getColCount() - 1, startCol + sourceCols - 1);
      engine.selectedCell = { row: endRow, col: endCol };
      engine.selectedKeys = keysFromPoints(
        selectionIsRect && engine.selectedKeys.length > 1
          ? { row: selectionRect.top, col: selectionRect.left }
          : { row: startRow, col: startCol },
        { row: endRow, col: endCol }
      );
      engine.editMode = false;
      updateSelectionUI();
      moveInputToCell(endRow, endCol);
      focusInput();
    }

    function moveSelection(row, col) {
      commitActiveCellValue();
      ensureRowCount(Math.max(getRowCount(), row + 2));
      var nr = Math.max(0, Math.min(getRowCount() - 1, row));
      var nc = Math.max(0, Math.min(getColCount() - 1, col));
      engine.selectedCell = { row: nr, col: nc };
      engine.selectedKeys = [cellKey(nr, nc)];
      engine.editMode = false;
      engine.isComposing = false;
      updateSelectionUI();
      moveInputToCell(nr, nc);
      syncEditOrigin();
      focusInput();
    }

    function moveByTab(backward) {
      var row = engine.selectedCell.row;
      var col = engine.selectedCell.col + (backward ? -1 : 1);
      if (col < 0) {
        col = getColCount() - 1;
        row = Math.max(0, row - 1);
      } else if (col >= getColCount()) {
        col = 0;
        row = row + 1;
      }
      moveSelection(row, col);
    }

    function cutSelectionAsync() {
      var text = getSelectionText();
      if (!text) return;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).catch(function () {});
      }
      clearSelectionCells();
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

    function undoSelection() {
      var prev = engine.undoStack.pop();
      if (!prev) return;
      engine.redoStack.push(cloneRowList(getRows()));
      options.setRows(options.normalizeRows ? options.normalizeRows(prev) : prev);
      if (typeof options.onRowsChange === "function") options.onRowsChange(getRows());
      refreshGridValues();
    }

    function redoSelection() {
      var next = engine.redoStack.pop();
      if (!next) return;
      engine.undoStack.push(cloneRowList(getRows()));
      options.setRows(options.normalizeRows ? options.normalizeRows(next) : next);
      if (typeof options.onRowsChange === "function") options.onRowsChange(getRows());
      refreshGridValues();
    }

    function updateSelectionUI() {
      if (!engine.host) return;
      var currentSelected = {};
      engine.selectedKeys.forEach(function (key) { currentSelected[key] = true; });
      var activeKey = cellKey(engine.selectedCell.row, engine.selectedCell.col);
      var dirtyKeys = {};
      engine.prevSelectedKeys.forEach(function (key) { dirtyKeys[key] = true; });
      engine.selectedKeys.forEach(function (key) { dirtyKeys[key] = true; });
      if (engine.prevActiveKey) dirtyKeys[engine.prevActiveKey] = true;
      dirtyKeys[activeKey] = true;
      Object.keys(dirtyKeys).forEach(function (key) {
        var td = engine.cellMap[key];
        if (!td) return;
        td.classList.toggle("sel-in", !!currentSelected[key]);
        td.classList.toggle("sel-active", key === activeKey);
      });
      engine.colHeadEls.forEach(function (el) {
        el.classList.toggle("st-head-active", engine.selectedKeys.some(function (key) {
          return Number(key.split(":")[1]) === Number(el.getAttribute("data-col-head"));
        }));
      });
      engine.rowHeadEls.forEach(function (el) {
        el.classList.toggle("st-head-active", engine.selectedKeys.some(function (key) {
          return Number(key.split(":")[0]) === Number(el.getAttribute("data-row-head"));
        }));
      });
      engine.prevSelectedKeys = engine.selectedKeys.slice();
      engine.prevActiveKey = activeKey;
      updateStatusBar();
    }

    function selectRow(rowIndex) {
      engine.selectedCell = { row: rowIndex, col: 0 };
      engine.selectedKeys = keysFromPoints({ row: rowIndex, col: 0 }, { row: rowIndex, col: getColCount() - 1 });
      engine.editMode = false;
      updateSelectionUI();
      moveInputToCell(rowIndex, 0);
      focusInput();
    }

    function selectColumn(colIndex) {
      engine.selectedCell = { row: 0, col: colIndex };
      engine.selectedKeys = keysFromPoints({ row: 0, col: colIndex }, { row: getRowCount() - 1, col: colIndex });
      engine.editMode = false;
      updateSelectionUI();
      moveInputToCell(0, colIndex);
      focusInput();
    }

    function selectAll() {
      engine.selectedCell = { row: 0, col: 0 };
      engine.selectedKeys = keysFromPoints({ row: 0, col: 0 }, { row: getRowCount() - 1, col: getColCount() - 1 });
      engine.editMode = false;
      updateSelectionUI();
      moveInputToCell(0, 0);
      focusInput();
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

    function onTableKeyDown(e) {
      if (engine.editMode) return;
      if (engine.inputEl && e.target === engine.inputEl) return;
      if (e.ctrlKey || e.metaKey) {
        var k = e.key.toLowerCase();
        if (k === "a") {
          e.preventDefault();
          selectAll();
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
          undoSelection();
          return;
        }
        if (k === "y") {
          e.preventDefault();
          redoSelection();
          return;
        }
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
        moveSelection(e.shiftKey ? engine.selectedCell.row - 1 : engine.selectedCell.row + 1, engine.selectedCell.col);
        return;
      }
      if (e.key === "ArrowUp") {
        e.preventDefault();
        moveSelection(engine.selectedCell.row - 1, engine.selectedCell.col);
        return;
      }
      if (e.key === "ArrowDown") {
        e.preventDefault();
        moveSelection(engine.selectedCell.row + 1, engine.selectedCell.col);
        return;
      }
      if (e.key === "ArrowLeft") {
        e.preventDefault();
        moveSelection(engine.selectedCell.row, engine.selectedCell.col - 1);
        return;
      }
      if (e.key === "ArrowRight") {
        e.preventDefault();
        moveSelection(engine.selectedCell.row, engine.selectedCell.col + 1);
        return;
      }
      if (e.key === "F2") {
        e.preventDefault();
        engine.editSnapshotRows = cloneRowList(getRows());
        engine.editMode = true;
        syncInputOverlay();
        focusInput();
      }
    }

    function onInputKeyDown(e) {
      var row = engine.selectedCell.row;
      var col = engine.selectedCell.col;
      var input = e.target;
      var start = input.selectionStart != null ? input.selectionStart : 0;
      var end = input.selectionEnd != null ? input.selectionEnd : 0;
      var len = input.value.length;

      if (e.key === "F2") {
        e.preventDefault();
        if (!engine.editMode) {
          engine.editSnapshotRows = cloneRowList(getRows());
          engine.editMode = true;
          syncInputOverlay();
        }
        focusInput();
        return;
      }

      if (e.key === "Delete" && engine.selectedKeys.length > 1) {
        e.preventDefault();
        engine.editMode = false;
        clearSelectionCells();
        return;
      }

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "a" && !engine.editMode) {
        e.preventDefault();
        selectAll();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "c" && !engine.editMode) {
        e.preventDefault();
        copySelectionAsync();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "x" && !engine.editMode) {
        e.preventDefault();
        cutSelectionAsync();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "v" && !engine.editMode) {
        e.preventDefault();
        pasteFromClipboardAsync();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z") {
        e.preventDefault();
        undoSelection();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "y") {
        e.preventDefault();
        redoSelection();
        return;
      }
      if (e.key === "Escape") {
        e.preventDefault();
        if (engine.editSnapshotRows) {
          options.setRows(options.normalizeRows ? options.normalizeRows(engine.editSnapshotRows) : engine.editSnapshotRows);
          if (typeof options.onRowsChange === "function") options.onRowsChange(getRows());
          engine.editSnapshotRows = null;
          refreshGridValues();
        }
        moveSelection(row, col);
        return;
      }
      if (engine.isComposing || e.isComposing) return;
      if (!engine.editMode) {
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
          moveSelection(e.shiftKey ? row - 1 : row + 1, col);
          return;
        }
        if (e.key === "ArrowUp") {
          e.preventDefault();
          moveSelection(row - 1, col);
          return;
        }
        if (e.key === "ArrowDown") {
          e.preventDefault();
          moveSelection(row + 1, col);
          return;
        }
        if (e.key === "ArrowLeft") {
          e.preventDefault();
          moveSelection(row, col - 1);
          return;
        }
        if (e.key === "ArrowRight") {
          e.preventDefault();
          moveSelection(row, col + 1);
          return;
        }
      } else {
        if (e.key === "Tab") {
          e.preventDefault();
          moveByTab(e.shiftKey);
          return;
        }
        if (e.key === "Enter") {
          e.preventDefault();
          moveSelection(e.shiftKey ? row - 1 : row + 1, col);
          return;
        }
        if (e.key === "ArrowUp") {
          e.preventDefault();
          moveSelection(row - 1, col);
          return;
        }
        if (e.key === "ArrowDown") {
          e.preventDefault();
          moveSelection(row + 1, col);
          return;
        }
        if (e.key === "ArrowLeft" && start === 0 && end === 0) {
          e.preventDefault();
          moveSelection(row, col - 1);
          return;
        }
        if (e.key === "ArrowRight" && start === len && end === len) {
          e.preventDefault();
          moveSelection(row, col + 1);
          return;
        }
      }
    }

    function onCellMouseDown(e) {
      var td = e.target.closest("td[data-r]");
      if (!td || e.button !== 0) return;
      e.preventDefault();
      commitActiveCellValue();
      var row = Number(td.getAttribute("data-r"));
      var col = Number(td.getAttribute("data-c"));
      engine.dragging = true;
      engine.dragStart = { row: row, col: col };
      engine.selectedCell = { row: row, col: col };
      engine.editMode = false;
      var key = cellKey(row, col);
      var already = engine.selectedKeys.indexOf(key) >= 0;
      if (e.ctrlKey || e.metaKey) {
        engine.dragMode = already ? "remove" : "add";
        if (already) {
          engine.selectedKeys = engine.selectedKeys.filter(function (item) { return item !== key; });
        } else {
          engine.selectedKeys = Array.from(new Set(engine.selectedKeys.concat([key])));
        }
      } else {
        engine.dragMode = "replace";
        engine.selectedKeys = [key];
      }
      engine.dragBaseKeys = engine.selectedKeys.slice();
      updateSelectionUI();
      moveInputToCell(row, col);
      syncEditOrigin();
      focusInput();
    }

    function onCellDblClick(e) {
      var td = e.target.closest("td[data-r]");
      if (!td) return;
      commitActiveCellValue();
      engine.selectedCell = { row: Number(td.getAttribute("data-r")), col: Number(td.getAttribute("data-c")) };
      engine.selectedKeys = [cellKey(engine.selectedCell.row, engine.selectedCell.col)];
      engine.editSnapshotRows = cloneRowList(getRows());
      engine.editMode = true;
      updateSelectionUI();
      moveInputToCell(engine.selectedCell.row, engine.selectedCell.col);
      focusInput();
    }

    function onHeadMouseDown(e) {
      var colResizer = e.target.closest("[data-mini-col-resizer]");
      if (colResizer) {
        e.preventDefault();
        if (e.detail >= 2) {
          engine.colWidths[Number(colResizer.getAttribute("data-mini-col-resizer"))] =
            getColumns()[Number(colResizer.getAttribute("data-mini-col-resizer"))].width || 100;
          refreshGridValues();
          return;
        }
        engine.resizing = {
          kind: "col",
          index: Number(colResizer.getAttribute("data-mini-col-resizer")),
          startX: e.clientX,
          startY: e.clientY,
          startSize: getColWidth(Number(colResizer.getAttribute("data-mini-col-resizer"))),
        };
        return;
      }
      var corner = e.target.closest("[data-corner]");
      if (corner) {
        e.preventDefault();
        selectAll();
        return;
      }
      var head = e.target.closest("[data-col-head]");
      if (head) {
        e.preventDefault();
        selectColumn(Number(head.getAttribute("data-col-head")));
      }
    }

    function onHeadDoubleClick(e) {
      var colResizer = e.target.closest("[data-mini-col-resizer]");
      if (colResizer) {
        var colIndex = Number(colResizer.getAttribute("data-mini-col-resizer"));
        engine.colWidths[colIndex] = getColumns()[colIndex].width || 100;
        refreshGridValues();
      }
    }

    function bindGlobalListeners() {
      if (engine.globalListenersBound) return;
      engine.globalListenersBound = true;
      document.addEventListener("mousemove", function (e) {
        if (engine.resizing) {
          if (engine.resizing.kind === "col") {
            engine.colWidths[engine.resizing.index] = Math.max(56, engine.resizing.startSize + (e.clientX - engine.resizing.startX));
          } else {
            engine.rowHeights[engine.resizing.index] = Math.max(24, engine.resizing.startSize + (e.clientY - engine.resizing.startY));
          }
          refreshGridValues();
          return;
        }
        if (!engine.dragging || !engine.host) return;
        var el = document.elementFromPoint(e.clientX, e.clientY);
        var td = el && el.closest ? el.closest("td[data-r]") : null;
        if (!td || !engine.host.contains(td)) return;
        var current = { row: Number(td.getAttribute("data-r")), col: Number(td.getAttribute("data-c")) };
        engine.selectedCell = current;
        var rectKeys = keysFromPoints(engine.dragStart, current);
        if (engine.dragMode === "replace") engine.selectedKeys = rectKeys;
        else if (engine.dragMode === "add") engine.selectedKeys = Array.from(new Set(engine.dragBaseKeys.concat(rectKeys)));
        else engine.selectedKeys = engine.dragBaseKeys.filter(function (key) { return rectKeys.indexOf(key) < 0; });
        updateSelectionUI();
        moveInputToCell(current.row, current.col);
      });
      window.addEventListener("mouseup", function () {
        engine.dragging = false;
        engine.resizing = null;
      });
      document.addEventListener("mousedown", function (e) {
        var rowResizer = e.target.closest && e.target.closest("[data-mini-row-resizer]");
        if (!rowResizer || !engine.host || !engine.host.contains(rowResizer)) return;
        e.preventDefault();
        if (e.detail >= 2) {
          engine.rowHeights[Number(rowResizer.getAttribute("data-mini-row-resizer"))] = 28;
          refreshGridValues();
          return;
        }
        engine.resizing = {
          kind: "row",
          index: Number(rowResizer.getAttribute("data-mini-row-resizer")),
          startX: e.clientX,
          startY: e.clientY,
          startSize: getRowHeight(Number(rowResizer.getAttribute("data-mini-row-resizer"))),
        };
      });
      document.addEventListener("mousedown", function (e) {
        var rowHead = e.target.closest && e.target.closest("[data-row-head]");
        if (rowHead && engine.host && engine.host.contains(rowHead) && !e.target.closest("[data-mini-row-resizer]")) {
          e.preventDefault();
          selectRow(Number(rowHead.getAttribute("data-row-head")));
        }
      });
    }

    engine.attach = function (host) {
      engine.host = host;
      ensureRowCount(options.minRows || 12);
      refreshGridValues();
      bindGlobalListeners();
    };

    engine.refresh = function () {
      if (engine.host) refreshGridValues();
    };

    return engine;
  }

  function getWorkSheetEngine() {
    if (workSheetEngine) return workSheetEngine;
    var createSheetEngine = window.createSheetEngine || createMiniSheetEngine;
    workSheetEngine = createSheetEngine({
      idPrefix: "work-grid",
      title: "품목 시트",
      subtitle: "매출현황과 같은 셀 선택/복사/붙여넣기 사용",
      maxHeight: 460,
      minRows: 12,
      columns: workSheetColumns,
      emptyRow: function () { return {}; },
      getRows: function () { return normalizeWorkItemsRows(workState.items, 12); },
      setRows: function (rows) { workState.items = normalizeWorkItemsRows(rows, 12); },
      normalizeRows: function (rows, minCount) { return normalizeWorkItemsRows(rows, minCount || 12); },
      normalizeValue: function (rowIndex, colIndex, value) {
        var key = workSheetColumns[colIndex].key;
        if (key === "date") return normalizeCellValue(0, value);
        return value;
      },
      formatValue: function (rowIndex, colIndex, raw) {
        var key = workSheetColumns[colIndex].key;
        if (key === "qty" || key === "price" || key === "supply" || key === "tax") {
          return formatDisplayNumber(raw);
        }
        return raw;
      },
      onRowsChange: function () {
        scheduleLedgerDraftSave();
        refreshWorkSummaryUi();
      },
    });
    return workSheetEngine;
  }

  function getClientSheetEngine() {
    if (clientSheetEngine) return clientSheetEngine;
    var createSheetEngine = window.createSheetEngine || createMiniSheetEngine;
    clientSheetEngine = createSheetEngine({
      idPrefix: "client-grid",
      title: "업체 시트",
      subtitle: "맨 왼쪽 구분 셀에서 법인/개인을 선택하거나 붙여넣기",
      maxHeight: 560,
      minRows: 20,
      columns: clientSheetColumns,
      emptyRow: function () { return emptyClientRow(); },
      getRows: function () { return normalizeClientGridRows(clientState.rows, 20); },
      setRows: function (rows) { clientState.rows = normalizeClientGridRows(rows, 20); },
      normalizeRows: function (rows, minCount) { return normalizeClientGridRows(rows, minCount || 20); },
      getCellEditor: function (rowIndex, colIndex) {
        var key = clientSheetColumns[colIndex].key;
        if (key === "supplierMode") {
          return {
            type: "select",
            options: [
              { value: "corporate", label: "법인" },
              { value: "personal", label: "개인" }
            ]
          };
        }
        return null;
      },
      normalizeValue: function (rowIndex, colIndex, value) {
        var key = clientSheetColumns[colIndex].key;
        if (key === "supplierMode") {
          var parsed = parseSupplierModeLabel(value);
          return parsed || (String(value || "").trim() ? "corporate" : "");
        }
        return value;
      },
      formatValue: function (rowIndex, colIndex, raw) {
        var key = clientSheetColumns[colIndex].key;
        if (key === "supplierMode") return raw === "personal" ? "개인" : raw === "corporate" ? "법인" : "";
        return raw;
      },
      onRowsChange: function () {
        scheduleLedgerDraftSave();
      },
    });
    return clientSheetEngine;
  }

  function attachWorkGridWhenNeeded(container) {
    if (!container) return;
    getWorkSheetEngine().attach(container);
  }

  function attachClientGridWhenNeeded(container) {
    if (!container) return;
    getClientSheetEngine().attach(container);
  }

  function getPriceSheetEngine() {
    if (priceSheetEngine) return priceSheetEngine;
    var createSheetEngine = window.createSheetEngine || createMiniSheetEngine;
    priceSheetEngine = createSheetEngine({
      idPrefix: "price-grid",
      title: "업체별 매출단가 시트",
      subtitle: "선택한 업체의 코드 / 단가 / 품목명을 엑셀처럼 복사·붙여넣기",
      maxHeight: 560,
      minRows: 40,
      columns: priceSheetColumns,
      emptyRow: function () { return { code: "", price: "", name: "" }; },
      getRows: function () { return getPriceRowsForClient(ensureActivePriceClient()); },
      setRows: function (rows) { setPriceRowsForClient(ensureActivePriceClient(), rows); },
      normalizeRows: function (rows, minCount) { return normalizePriceEditorRows(rows, minCount || 40); },
      formatValue: function (rowIndex, colIndex, raw) {
        var key = priceSheetColumns[colIndex].key;
        if (key === "price") return formatDisplayNumber(raw);
        return raw;
      },
      onRowsChange: function () {
        scheduleLedgerDraftSave();
      },
    });
    return priceSheetEngine;
  }

  function attachPriceGridWhenNeeded(container) {
    if (!container) return;
    getPriceSheetEngine().attach(container);
  }

  function getClosingAttendanceSheetEngine() {
    if (closingAttendanceSheetEngine) return closingAttendanceSheetEngine;
    var createSheetEngine = window.createSheetEngine || createMiniSheetEngine;
    closingAttendanceSheetEngine = createSheetEngine({
      idPrefix: "closing-attendance-grid",
      title: "정직원 월 시트",
      subtitle: "",
      maxHeight: 620,
      minRows: closingEmployeeNames.length * closingEmployeeMarkers.length,
      columns: closingAttendanceColumns,
      emptyRow: function () { return emptyClosingEmployeeRow("", ""); },
      getRows: function () { return getClosingEmployeeRows(closingState.attendanceMonth); },
      setRows: function (rows) { setClosingEmployeeRows(closingState.attendanceMonth, rows); },
      normalizeRows: function (rows) { return normalizeClosingEmployeeRows(rows); },
      onRowsChange: function () {
        scheduleLedgerDraftSave();
      },
    });
    return closingAttendanceSheetEngine;
  }

  function attachClosingAttendanceGridWhenNeeded(container) {
    if (!container) return;
    getClosingAttendanceSheetEngine().attach(container);
  }

  function getClosingOutsourceSheetEngine() {
    if (closingOutsourceSheetEngine) return closingOutsourceSheetEngine;
    var createSheetEngine = window.createSheetEngine || createMiniSheetEngine;
    closingOutsourceSheetEngine = createSheetEngine({
      idPrefix: "closing-outsource-grid",
      title: "아웃소싱 월 시트",
      subtitle: "",
      maxHeight: 620,
      minRows: 72,
      columns: closingOutsourceColumns,
      emptyRow: function () { return emptyClosingOutsourceRow("", ""); },
      getRows: function () {
        return getClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor);
      },
      setRows: function (rows) {
        setClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor, rows);
      },
      normalizeRows: function (rows) { return normalizeClosingOutsourceRows(rows); },
      onRowsChange: function () {
        scheduleLedgerDraftSave();
      },
    });
    return closingOutsourceSheetEngine;
  }

  function attachClosingOutsourceGridWhenNeeded(container) {
    if (!container) return;
    getClosingOutsourceSheetEngine().attach(container);
  }

  function ensureDraftId() {
    if (ledgerState.draftId) return ledgerState.draftId;

    var localId = null;
    try {
      localId =
        localStorage.getItem(DRAFT_ID_KEY) ||
        localStorage.getItem(LEGACY_DRAFT_ID_KEY);
    } catch (err) {}

    ledgerState.legacyDraftId =
      localId && localId !== SHARED_DRAFT_ID ? localId : null;
    ledgerState.draftId = SHARED_DRAFT_ID;
    workState.draftId = SHARED_DRAFT_ID;

    try {
      localStorage.setItem(DRAFT_ID_KEY, SHARED_DRAFT_ID);
      localStorage.setItem(LEGACY_DRAFT_ID_KEY, SHARED_DRAFT_ID);
    } catch (err) {}

    return SHARED_DRAFT_ID;
  }

  function getLedgerPayload() {
    ensureLedgerBundles();
    syncActiveLedgerBundle();
    var salesStoredRowCount = Math.max(MIN_ROW_COUNT, getLastUsedRowIndex(salesLedgerBundle.rows) + 20);
    var purchaseStoredRowCount = Math.max(MIN_ROW_COUNT, getLastUsedRowIndex(purchaseLedgerBundle.rows) + 20);
    var trimmedWorkItems = trimTrailingRows(workState.items, workGridFields, 12, function () { return {}; });
    var trimmedClientRows = normalizeClientGridRows(salesLedgerBundle.clientState.rows, 20);
    var trimmedPriceRows = trimTrailingRows(salesLedgerBundle.priceState.rows, priceGridFields, 30, emptyPriceRow);
    var trimmedPurchaseClientRows = normalizeClientGridRows(purchaseLedgerBundle.clientState.rows, 20);
    var trimmedPurchasePriceRows = trimTrailingRows(purchaseLedgerBundle.priceState.rows, priceGridFields, 30, emptyPriceRow);
    return {
      sales: {
        statusRows: normalizeSalesRows(salesLedgerBundle.rows, salesStoredRowCount).slice(0, salesStoredRowCount),
        clientRows: cloneClientRows(trimmedClientRows),
        priceRows: trimmedPriceRows.map(function (row) { return Object.assign({}, row); }),
        priceActiveClient: salesLedgerBundle.priceState.activeClient || "",
        sheetLayout: {
          colWidths: salesLedgerBundle.colWidths.slice(),
          rowHeights: salesLedgerBundle.rowHeights.slice(0, salesStoredRowCount),
        },
        manageState: Object.assign({}, salesLedgerBundle.manageState),
      },
      purchase: {
        statusRows: normalizeSalesRows(purchaseLedgerBundle.rows, purchaseStoredRowCount).slice(0, purchaseStoredRowCount),
        clientRows: cloneClientRows(trimmedPurchaseClientRows),
        priceRows: trimmedPurchasePriceRows.map(function (row) { return Object.assign({}, row); }),
        priceActiveClient: purchaseLedgerBundle.priceState.activeClient || "",
        sheetLayout: {
          colWidths: purchaseLedgerBundle.colWidths.slice(),
          rowHeights: purchaseLedgerBundle.rowHeights.slice(0, purchaseStoredRowCount),
        },
        manageState: Object.assign({}, purchaseLedgerBundle.manageState),
      },
      shared: {
        workItems: cloneItems(trimmedWorkItems),
        workInfo: cloneWorkInfo(workState.info),
      },
      closing: {
        attendanceMonth: closingState.attendanceMonth || defaultClosingMonthLabel(),
        attendanceView: closingState.attendanceView === "outsource" ? "outsource" : "employee",
        employeeViewMode: closingState.employeeViewMode === "summary" ? "summary" : "entry",
        attendanceSearch: String(closingState.attendanceSearch || ""),
        employeeRoster: Array.isArray(closingState.employeeRoster) ? closingState.employeeRoster.slice() : [],
        employeeRowsByMonth: cloneClosingRowsMap(closingState.employeeRowsByMonth),
        outsourceVendor: closingState.outsourceVendor || "leaders",
        outsourceRosterByVendor: (function () {
          var source = closingState.outsourceRosterByVendor && typeof closingState.outsourceRosterByVendor === "object"
            ? closingState.outsourceRosterByVendor
            : {};
          var cloned = {};
          Object.keys(source).forEach(function (vendor) {
            var roster = Array.isArray(source[vendor]) ? source[vendor] : [];
            cloned[vendor] = roster.map(function (profile) {
              return {
                employee: String((profile && profile.employee) || ""),
                joinDate: String((profile && profile.joinDate) || ""),
              };
            });
          });
          return cloned;
        })(),
        outsourceRowsByKey: (function () {
          var source = closingState.outsourceRowsByKey && typeof closingState.outsourceRowsByKey === "object"
            ? closingState.outsourceRowsByKey
            : {};
          var cloned = {};
          Object.keys(source).forEach(function (key) {
            cloned[key] = normalizeClosingOutsourceRows(source[key]);
          });
          return cloned;
        })(),
        outsourceWarningsByKey: (function () {
          var source = closingState.outsourceWarningsByKey && typeof closingState.outsourceWarningsByKey === "object"
            ? closingState.outsourceWarningsByKey
            : {};
          var cloned = {};
          Object.keys(source).forEach(function (key) {
            cloned[key] = Object.assign({}, source[key] || {});
          });
          return cloned;
        })(),
        outsourceMetaByKey: (function () {
          var source = closingState.outsourceMetaByKey && typeof closingState.outsourceMetaByKey === "object"
            ? closingState.outsourceMetaByKey
            : {};
          var cloned = {};
          Object.keys(source).forEach(function (key) {
            cloned[key] = Object.assign({ retirement: "", hourlyWage: "", lockDay: "" }, source[key] || {});
          });
          return cloned;
        })(),
        salesCloseRowsByMonth: (function () {
          var source = closingState.salesCloseRowsByMonth && typeof closingState.salesCloseRowsByMonth === "object"
            ? closingState.salesCloseRowsByMonth
            : {};
          var cloned = {};
          Object.keys(source).forEach(function (key) {
            cloned[key] = normalizeClosingSalesCloseRows(source[key]);
          });
          return cloned;
        })(),
        purchaseCloseRowsByMonth: (function () {
          var source = closingState.purchaseCloseRowsByMonth && typeof closingState.purchaseCloseRowsByMonth === "object"
            ? closingState.purchaseCloseRowsByMonth
            : {};
          var cloned = {};
          Object.keys(source).forEach(function (key) {
            cloned[key] = normalizeClosingPurchaseCloseRows(source[key]);
          });
          return cloned;
        })(),
      },
      updatedAt: new Date().toISOString(),
    };
  }

  function applyLedgerData(data) {
    if (!data) return;
    ensureLedgerBundles();

    var salesData = data.sales && typeof data.sales === "object" ? data.sales : data;
    var purchaseData = data.purchase && typeof data.purchase === "object" ? data.purchase : null;
    var sharedData = data.shared && typeof data.shared === "object" ? data.shared : data;
    var closingData = data.closing && typeof data.closing === "object" ? data.closing : null;

    if (Array.isArray(salesData.statusRows)) {
      var nextSalesRows = normalizeSalesRows(salesData.statusRows);
      var currentSalesLast = getLastUsedRowIndex(salesLedgerBundle.rows || []);
      var nextSalesLast = getLastUsedRowIndex(nextSalesRows);
      if (!(nextSalesLast < 0 && currentSalesLast >= 0)) {
        salesLedgerBundle.rows = nextSalesRows;
      }
    }

    if (Array.isArray(sharedData.workItems)) {
      workState.items = trimTrailingRows(sharedData.workItems, workGridFields, 12, function () { return {}; });
    } else if (Array.isArray(sharedData.items)) {
      workState.items = trimTrailingRows(sharedData.items, workGridFields, 12, function () { return {}; });
    }

    if (sharedData.workInfo && typeof sharedData.workInfo === "object") {
      workState.info = cloneWorkInfo(sharedData.workInfo);
    }

    if (salesData && Array.isArray(salesData.clientRows)) {
      salesLedgerBundle.clientState.rows = normalizeClientGridRows(salesData.clientRows, 20);
    } else if (Array.isArray(sharedData.clientRows)) {
      salesLedgerBundle.clientState.rows = normalizeClientGridRows(sharedData.clientRows, 20);
    }

    if (Array.isArray(salesData.priceRows)) {
      salesLedgerBundle.priceState.rows = normalizePriceGridRows(salesData.priceRows, 30);
    }
    if (typeof salesData.priceActiveClient === "string") {
      salesLedgerBundle.priceState.activeClient = salesData.priceActiveClient;
    }

    if (salesData.sheetLayout && typeof salesData.sheetLayout === "object") {
      if (Array.isArray(salesData.sheetLayout.colWidths)) {
        salesLedgerBundle.colWidths = salesData.sheetLayout.colWidths.slice(0, COL_COUNT);
      }
      if (Array.isArray(salesData.sheetLayout.rowHeights)) {
        salesLedgerBundle.rowHeights = salesData.sheetLayout.rowHeights.slice();
      }
    }
    if (salesData.manageState && typeof salesData.manageState === "object") {
      salesLedgerBundle.manageState = {
        startMonth: salesData.manageState.startMonth || "",
        endMonth: salesData.manageState.endMonth || "",
        client: salesData.manageState.client || "",
      };
    }

    if (purchaseData && Array.isArray(purchaseData.statusRows)) {
      var nextPurchaseRows = normalizeSalesRows(purchaseData.statusRows);
      var currentPurchaseLast = getLastUsedRowIndex(purchaseLedgerBundle.rows || []);
      var nextPurchaseLast = getLastUsedRowIndex(nextPurchaseRows);
      if (!(nextPurchaseLast < 0 && currentPurchaseLast >= 0)) {
        purchaseLedgerBundle.rows = nextPurchaseRows;
      }
    } else if (Array.isArray(data.purchaseStatusRows)) {
      var legacyPurchaseRows = normalizeSalesRows(data.purchaseStatusRows);
      var currentLegacyPurchaseLast = getLastUsedRowIndex(purchaseLedgerBundle.rows || []);
      var nextLegacyPurchaseLast = getLastUsedRowIndex(legacyPurchaseRows);
      if (!(nextLegacyPurchaseLast < 0 && currentLegacyPurchaseLast >= 0)) {
        purchaseLedgerBundle.rows = legacyPurchaseRows;
      }
    }
    if (purchaseData && Array.isArray(purchaseData.clientRows)) {
      purchaseLedgerBundle.clientState.rows = normalizeClientGridRows(purchaseData.clientRows, 20);
    } else if (Array.isArray(data.purchaseClientRows)) {
      purchaseLedgerBundle.clientState.rows = normalizeClientGridRows(data.purchaseClientRows, 20);
    } else if (Array.isArray(sharedData.clientRows)) {
      purchaseLedgerBundle.clientState.rows = normalizeClientGridRows(sharedData.clientRows, 20);
    }
    if (purchaseData && Array.isArray(purchaseData.priceRows)) {
      purchaseLedgerBundle.priceState.rows = normalizePriceGridRows(purchaseData.priceRows, 30);
    } else if (Array.isArray(data.purchasePriceRows)) {
      purchaseLedgerBundle.priceState.rows = normalizePriceGridRows(data.purchasePriceRows, 30);
    }
    if (purchaseData && typeof purchaseData.priceActiveClient === "string") {
      purchaseLedgerBundle.priceState.activeClient = purchaseData.priceActiveClient;
    } else if (typeof data.purchasePriceActiveClient === "string") {
      purchaseLedgerBundle.priceState.activeClient = data.purchasePriceActiveClient;
    }
    if (purchaseData && purchaseData.sheetLayout && typeof purchaseData.sheetLayout === "object") {
      if (Array.isArray(purchaseData.sheetLayout.colWidths)) {
        purchaseLedgerBundle.colWidths = purchaseData.sheetLayout.colWidths.slice(0, COL_COUNT);
      }
      if (Array.isArray(purchaseData.sheetLayout.rowHeights)) {
        purchaseLedgerBundle.rowHeights = purchaseData.sheetLayout.rowHeights.slice();
      }
    } else if (data.purchaseSheetLayout && typeof data.purchaseSheetLayout === "object") {
      if (Array.isArray(data.purchaseSheetLayout.colWidths)) {
        purchaseLedgerBundle.colWidths = data.purchaseSheetLayout.colWidths.slice(0, COL_COUNT);
      }
      if (Array.isArray(data.purchaseSheetLayout.rowHeights)) {
        purchaseLedgerBundle.rowHeights = data.purchaseSheetLayout.rowHeights.slice();
      }
    }
    if (purchaseData && purchaseData.manageState && typeof purchaseData.manageState === "object") {
      purchaseLedgerBundle.manageState = {
        startMonth: purchaseData.manageState.startMonth || "",
        endMonth: purchaseData.manageState.endMonth || "",
        client: purchaseData.manageState.client || "",
      };
    } else if (data.purchaseManageState && typeof data.purchaseManageState === "object") {
      purchaseLedgerBundle.manageState = {
        startMonth: data.purchaseManageState.startMonth || "",
        endMonth: data.purchaseManageState.endMonth || "",
        client: data.purchaseManageState.client || "",
      };
    }

    if (closingData) {
      closingState.attendanceMonth = /^\d+월$/.test(String(closingData.attendanceMonth || ""))
        ? closingData.attendanceMonth
        : defaultClosingMonthLabel();
      closingState.attendanceView = closingData.attendanceView === "outsource" ? "outsource" : "employee";
      closingState.employeeViewMode = closingData.employeeViewMode === "summary" ? "summary" : "entry";
      closingState.attendanceSearch = String(closingData.attendanceSearch || "");
      closingState.employeeRoster = Array.isArray(closingData.employeeRoster)
        ? closingData.employeeRoster.map(function (name) { return String(name || ""); })
        : [];
      closingState.employeeRowsByMonth = cloneClosingRowsMap(closingData.employeeRowsByMonth);
      closingState.outsourceVendor = String(closingData.outsourceVendor || "leaders");
      closingState.outsourceRosterByVendor = {};
      if (closingData.outsourceRosterByVendor && typeof closingData.outsourceRosterByVendor === "object") {
        Object.keys(closingData.outsourceRosterByVendor).forEach(function (vendor) {
          var roster = Array.isArray(closingData.outsourceRosterByVendor[vendor])
            ? closingData.outsourceRosterByVendor[vendor]
            : [];
          closingState.outsourceRosterByVendor[vendor] = roster.map(function (profile) {
            return {
              employee: String((profile && profile.employee) || ""),
              joinDate: String((profile && profile.joinDate) || ""),
            };
          });
        });
      }
      closingState.outsourceRowsByKey = {};
      if (closingData.outsourceRowsByKey && typeof closingData.outsourceRowsByKey === "object") {
        Object.keys(closingData.outsourceRowsByKey).forEach(function (key) {
          closingState.outsourceRowsByKey[key] =
            normalizeClosingOutsourceRows(closingData.outsourceRowsByKey[key]);
        });
      }
      closingState.outsourceWarningsByKey = {};
      if (closingData.outsourceWarningsByKey && typeof closingData.outsourceWarningsByKey === "object") {
        Object.keys(closingData.outsourceWarningsByKey).forEach(function (key) {
          closingState.outsourceWarningsByKey[key] = Object.assign({}, closingData.outsourceWarningsByKey[key] || {});
        });
      }
      closingState.outsourceMetaByKey = {};
      if (closingData.outsourceMetaByKey && typeof closingData.outsourceMetaByKey === "object") {
        Object.keys(closingData.outsourceMetaByKey).forEach(function (key) {
          closingState.outsourceMetaByKey[key] = Object.assign({ retirement: "", hourlyWage: "", lockDay: "" }, closingData.outsourceMetaByKey[key] || {});
        });
      }
      closingState.salesCloseRowsByMonth = {};
      if (closingData.salesCloseRowsByMonth && typeof closingData.salesCloseRowsByMonth === "object") {
        Object.keys(closingData.salesCloseRowsByMonth).forEach(function (key) {
          closingState.salesCloseRowsByMonth[key] = normalizeClosingSalesCloseRows(closingData.salesCloseRowsByMonth[key]);
        });
      }
      closingState.purchaseCloseRowsByMonth = {};
      if (closingData.purchaseCloseRowsByMonth && typeof closingData.purchaseCloseRowsByMonth === "object") {
        Object.keys(closingData.purchaseCloseRowsByMonth).forEach(function (key) {
          closingState.purchaseCloseRowsByMonth[key] = normalizeClosingPurchaseCloseRows(closingData.purchaseCloseRowsByMonth[key]);
        });
      }
    }
    ensureClosingAttendanceState();

    switchLedgerModule(activeLedgerKind, true);
    ST.visibleRowsDirty = true;
    initGridSizing();

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
      syncActiveLedgerBundle();
      var payload = getLedgerPayload();
      ledgerState.localUpdatedAt = payload.updatedAt;
      localStorage.setItem(LOCAL_BACKUP_KEY, JSON.stringify(payload));
    } catch (err) {
      console.warn("로컬 백업 저장 실패:", err);
    }
  }

  function flushLocalLedgerBackup() {
    if (ledgerState.localSaveTimer) {
      clearTimeout(ledgerState.localSaveTimer);
      ledgerState.localSaveTimer = null;
    }
    saveLocalLedgerBackup();
  }

  function scheduleLocalLedgerBackup(delay) {
    if (ledgerState.localSaveTimer) clearTimeout(ledgerState.localSaveTimer);
    ledgerState.localSaveTimer = setTimeout(function () {
      ledgerState.localSaveTimer = null;
      saveLocalLedgerBackup();
    }, delay || 600);
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

    window.addEventListener("beforeunload", function () {
      if (ledgerState.dirty) flushLocalLedgerBackup();
    });
  }

  function saveLedgerDraftToFirebase(attempts) {
    attempts = attempts || 0;
    ensureDraftId();
    flushLocalLedgerBackup();

    if (!ledgerState.initialLoadFinished && ledgerState.loadingPromise) {
      queueLedgerSaveRetry(1200);
      return;
    }

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
    scheduleLocalLedgerBackup(700);

    if (ledgerState.saveTimer) clearTimeout(ledgerState.saveTimer);
    ledgerState.saveTimer = setTimeout(function () {
      saveLedgerDraftToFirebase();
    }, 1200);
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
        ledgerState.initialLoadFinished = true;
        ledgerState.loadingPromise = null;
        if (ledgerState.dirty) {
          queueLedgerSaveRetry(300);
        }
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
                ledgerState.cloudOffline = false;
                finish(true);
                return;
              }
              if (
                !data &&
                ledgerState.legacyDraftId &&
                ledgerState.legacyDraftId !== id
              ) {
                window.firebaseFirestoreApi
                  .loadWorkDraft(ledgerState.legacyDraftId)
                  .then(function (legacyData) {
                    if (legacyData && !ledgerState.dirty) {
                      applyLedgerData(legacyData);
                      saveLocalLedgerBackup();
                      window.firebaseFirestoreApi
                        .saveWorkDraft(id, legacyData)
                        .catch(function (migrationErr) {
                          console.warn("기존 Firestore 문서를 공용 문서로 옮기지 못했습니다:", migrationErr);
                        });
                    }
                    ledgerState.cloudOffline = false;
                    finish(true);
                  })
                  .catch(function () {
                    ledgerState.cloudOffline = false;
                    finish(true);
                  });
                return;
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
    if (salesColumns[colIndex].key === "client") return getDisplayCompanyName(text);
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
    if (salesColumns[col].key === "client" || salesColumns[col].key === "code") {
      applyPriceLookupForSalesRow(nextRows, row);
    }
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

  function cellHasData(r, c) {
    return String(getRawValue(r, c) || "").trim() !== "";
  }

  function findCtrlJumpTarget(row, col, direction) {
    var maxRow = Math.max(0, getRowCount() - 1);
    var maxCol = Math.max(0, COL_COUNT - 1);
    var step = (direction === "up" || direction === "left") ? -1 : 1;
    var currentHasData = cellHasData(row, col);
    var cursor;
    var limit;

    if (direction === "up" || direction === "down") {
      cursor = row + step;
      if (cursor < 0) return { row: 0, col: col };
      if (cursor > maxRow) return { row: maxRow, col: col };

      if (currentHasData && cellHasData(cursor, col)) {
        while (cursor + step >= 0 && cursor + step <= maxRow && cellHasData(cursor + step, col)) {
          cursor += step;
        }
        return { row: cursor, col: col };
      }

      while (cursor >= 0 && cursor <= maxRow) {
        if (cellHasData(cursor, col)) {
          return { row: cursor, col: col };
        }
        cursor += step;
      }

      if (direction === "down") {
        limit = getLastUsedRowIndex(ST.rows);
        return { row: limit >= 0 ? limit : maxRow, col: col };
      }
      return { row: 0, col: col };
    }

    cursor = col + step;
    if (cursor < 0) return { row: row, col: 0 };
    if (cursor > maxCol) return { row: row, col: maxCol };

    if (currentHasData && cellHasData(row, cursor)) {
      while (cursor + step >= 0 && cursor + step <= maxCol && cellHasData(row, cursor + step)) {
        cursor += step;
      }
      return { row: row, col: cursor };
    }

    while (cursor >= 0 && cursor <= maxCol) {
      if (cellHasData(row, cursor)) {
        return { row: row, col: cursor };
      }
      cursor += step;
    }

    return { row: row, col: direction === "right" ? maxCol : 0 };
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
    var clientType = ST.filter.clientType || "all";
    if (clientType !== "all") {
      var matchedClientRow = findClientRowByCompany(row && row.client);
      var supplierMode = getClientRowSupplierMode(matchedClientRow);
      if (clientType === "corporate" && supplierMode !== "corporate") return false;
      if (clientType === "personal" && supplierMode !== "personal") return false;
    }
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
        if (colKey === "date") {
          var ad = parseSalesMonthDay(av);
          var bd = parseSalesMonthDay(bv);
          var adValue = ad ? ad.month * 100 + ad.day : null;
          var bdValue = bd ? bd.month * 100 + bd.day : null;
          if (adValue != null && bdValue != null) {
            var dateResult = adValue - bdValue;
            if (!dateResult) {
              dateResult = (a && a.__order != null ? a.__order : 0) - (b && b.__order != null ? b.__order : 0);
            }
            return direction === "desc" ? -dateResult : dateResult;
          }
          if (adValue != null || bdValue != null) {
            var mixedResult = adValue != null ? -1 : 1;
            return direction === "desc" ? -mixedResult : mixedResult;
          }
        }
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
    syncActiveLedgerBundle();
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
    syncActiveLedgerBundle();
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
    restoreActiveInputBinding();
  }

  function restoreActiveInputBinding() {
    if (!ST.inputEl || !ST.host) return;
    var activeTd = ST.cellMap[cellKey(ST.selectedCell.row, ST.selectedCell.col)];
    if (!activeTd) return;
    var inner = activeTd.querySelector(".st-cell-inner");
    var disp = activeTd.querySelector(".st-display");
    var ghost = activeTd.querySelector(".st-ghost");
    if (!inner) return;
    if (ST.inputEl.parentNode !== inner) {
      inner.appendChild(ST.inputEl);
    }
    if (disp) disp.style.display = "none";
    if (ghost) ghost.style.display = "none";
    syncInputOverlay();
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

  function focusLastUsedSalesRow() {
    if (!ST.tableBuilt || !ST.scrollEl) return;
    var lastUsedRow = getLastUsedRowIndex(ST.rows);
    if (lastUsedRow < 0) return;

    var targetCol = Math.max(0, Math.min(COL_COUNT - 1, ST.selectedCell.col || 0));
    ST.selectedCell = { row: lastUsedRow, col: targetCol };
    ST.selectedKeys = [cellKey(lastUsedRow, targetCol)];
    ST.editMode = false;
    ST.isComposing = false;
    ST.fillPreviewKeys = [];
    ST.fillSourceRect = null;
    ST.fillDirection = null;

    rebuildVisibleRowsData();
    var visiblePos = getVisibleRowPosition(lastUsedRow);
    if (visiblePos >= 0) {
      var rowTop = ST.visibleRowOffsets[visiblePos] || 0;
      var anchorOffset = Math.max(0, ST.scrollEl.clientHeight - getRowHeight(lastUsedRow) * 6);
      ST.scrollEl.scrollTop = Math.max(0, rowTop - anchorOffset);
    }

    renderVirtualRows(true);
    updateSelectionUI();
    moveInputToCell(lastUsedRow, targetCol);
    syncEditOrigin();
    focusInput();
  }

  function moveSelection(row, col) {
    commitActiveCellValue();
    hideSalesAutocomplete();
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
    ST.editMode = false;
    ST.isComposing = false;
    updateSelectionUI();
    moveInputToCell(ST.selectedCell.row, ST.selectedCell.col);
    syncEditOrigin();
    focusInput();
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
          var rowIndex = Number(pos[0]);
          var colIndex = Number(pos[1]);
          setCellValueInRows(current, rowIndex, colIndex, grid[0][0], true);
          if (salesColumns[colIndex].key === "client" || salesColumns[colIndex].key === "code") {
            applyPriceLookupForSalesRow(current, rowIndex);
          }
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
            if (salesColumns[c].key === "client" || salesColumns[c].key === "code") {
              applyPriceLookupForSalesRow(current, r);
            }
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
            if (salesColumns[col].key === "client" || salesColumns[col].key === "code") {
              applyPriceLookupForSalesRow(current, row);
            }
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
    syncActiveLedgerBundle();
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
    syncActiveLedgerBundle();
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

  function syncInputOverlay(options) {
    if (!ST.tableBuilt || !ST.inputEl || !ST.host) return;
    options = options || {};
    var r = ST.selectedCell.row;
    var c = ST.selectedCell.col;
    if (options.ensureVisible) {
      ensureRowVisible(r);
      renderVirtualRows();
    }
    var td = ST.host.querySelector('.st-table td[data-r="' + r + '"][data-c="' + c + '"]');
    if (!td) return;
    var inner = td.querySelector(".st-cell-inner");
    if (!inner) return;
    var raw = getRawValue(r, c);
    var v = getValue(r, c);
    ST.inputEl.value = raw;
    ST.inputEl.style.height = getRowHeight(r) - 4 + "px";
    ST.inputEl.style.lineHeight = getRowHeight(r) - 4 + "px";
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
    updateSalesAutocomplete();
    requestAnimationFrame(function () {
      if (!ST.inputEl) return;
      if (ST.editMode) ST.inputEl.setSelectionRange(len, len);
      else ST.inputEl.setSelectionRange(0, len);
    });
  }

  function focusInput() {
    requestAnimationFrame(function () {
      if (!ST.inputEl) return;
      ST.inputEl.focus();
      var len = ST.inputEl.value.length;
      if (ST.editMode) ST.inputEl.setSelectionRange(len, len);
      else ST.inputEl.setSelectionRange(0, len);
    });
  }

  function buildSalesTable(host) {
    ST.host = host;
    initGridSizing();
    ST.rows = normalizeSalesRows(ST.rows);
    var lastUsedOnOpen = getLastUsedRowIndex(ST.rows);
    if (
      lastUsedOnOpen >= 0 &&
      (!ST.selectedKeys.length ||
        (ST.selectedCell.row === 0 && ST.selectedCell.col === 0 && !rowHasAnyValue(ST.rows[0], salesColumns.map(function (col) { return col.key; }))))
    ) {
      ST.selectedCell = { row: lastUsedOnOpen, col: 0 };
      ST.selectedKeys = [cellKey(lastUsedOnOpen, 0)];
    }
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
      '<div id="sales-grid-datalist" class="autocomplete-popover hidden"></div>' +
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
    ST.inputEl.setAttribute("autocomplete", "off");
    ST.inputEl.className = "st-input no-edit";
    ST.inputEl.style.height = getRowHeight(0) - 4 + "px";
    ST.inputEl.style.lineHeight = getRowHeight(0) - 4 + "px";
    var firstInner = tbody.querySelector(
      'td[data-r="' + ST.selectedCell.row + '"][data-c="' + ST.selectedCell.col + '"] .st-cell-inner'
    );
    if (firstInner) firstInner.appendChild(ST.inputEl);
    ST.inputEl.value = getRawValue(ST.selectedCell.row, ST.selectedCell.col);
    ST.autocompleteEl = host.querySelector("#sales-grid-datalist");
    syncInputOverlay({ ensureVisible: true });

    ST.scrollEl.addEventListener("keydown", onTableKeyDown);
    ST.scrollEl.addEventListener("copy", onCopy);
    ST.scrollEl.addEventListener("paste", onPaste);
    ST.scrollEl.addEventListener("dragstart", function (e) {
      e.preventDefault();
    });
    ST.scrollEl.addEventListener("wheel", function (e) {
      if (!ST.scrollEl) return;
      var deltaY = e.deltaY || 0;
      var deltaX = e.deltaX || 0;
      if (!deltaY && !deltaX) return;
      e.preventDefault();
      e.stopPropagation();
      ST.scrollEl.scrollTop += deltaY;
      ST.scrollEl.scrollLeft += deltaX;
    }, { passive: false });

    tbody.addEventListener("mousedown", onTbodyMouseDown);
    tbody.addEventListener("dblclick", onTbodyDblClick);
    if (!ST.globalListenersBound) {
      document.addEventListener("mousemove", onDocMouseMove);
      window.addEventListener("mouseup", onWindowMouseUp);
      document.addEventListener("keydown", onGlobalSalesKeyDown, true);
      ST.globalListenersBound = true;
    }
    ST.scrollEl.addEventListener("scroll", function () {
      if (ST.scrollEl.scrollTop + ST.scrollEl.clientHeight >= ST.scrollEl.scrollHeight - 120) {
        ensureRowCapacity(getRowCount() + ROW_GROWTH_CHUNK);
      }
      renderVirtualRows();
      updateSelectionUI();
      positionSalesAutocomplete();
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
      applyPriceLookupForActiveSalesCell();
      syncInputOverlay();
      updateSalesAutocomplete();
    });
    ST.inputEl.addEventListener("input", function (e) {
      if (!ST.editMode) {
        snapshotEditOrigin();
        ST.editMode = true;
      }
      setValue(ST.selectedCell.row, ST.selectedCell.col, e.target.value);
      applyPriceLookupForActiveSalesCell();
      syncInputOverlay();
      updateSalesAutocomplete();
    });
    ST.inputEl.addEventListener("focus", function (e) {
      var len = e.target.value.length;
      if (ST.editMode) e.target.setSelectionRange(len, len);
      else e.target.setSelectionRange(0, len);
    });
    ST.inputEl.addEventListener("blur", function () {
      commitActiveCellValue();
      setTimeout(hideSalesAutocomplete, 120);
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
    hideSalesAutocomplete();
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
    hideSalesAutocomplete();
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
    hideSalesAutocomplete();
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

  function isEditableDomTarget(target) {
    if (!target || !target.closest) return false;
    if (target === ST.inputEl) return true;
    return !!target.closest('input, textarea, select, [contenteditable="true"]');
  }

  function isSalesTextEditingTarget(target) {
    if (!target || !target.closest) return false;
    if (target === ST.inputEl) return true;
    return !!target.closest('input:not([type="button"]):not([type="checkbox"]):not([type="radio"]), textarea, [contenteditable="true"]');
  }

  function onGlobalSalesKeyDown(e) {
    if (!ST.tableBuilt || state.mainTab !== "sales" || state.subTab !== "status") return;
    if (isSalesTextEditingTarget(e.target)) return;
    var key = e.key;
    var lowered = String(key || "").toLowerCase();
    var shouldHandle =
      key === "ArrowUp" ||
      key === "ArrowDown" ||
      key === "ArrowLeft" ||
      key === "ArrowRight" ||
      key === "Tab" ||
      key === "Enter" ||
      key === "Delete" ||
      key === "F2" ||
      key === "Escape" ||
      ((e.ctrlKey || e.metaKey) && (lowered === "a" || lowered === "c" || lowered === "x" || lowered === "v" || lowered === "z" || lowered === "y" || lowered === "f")) ||
      (e.altKey && (key === "ArrowDown" || key === "ArrowUp" || key === "F2"));
    if (!shouldHandle) return;
    onTableKeyDown(e);
  }

  function onTableKeyDown(e) {
    if (ST.editMode) return;
    if (ST.inputEl && e.target === ST.inputEl) return;

    if (e.altKey && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
      e.preventDefault();
      openSalesAutocompletePicker();
      return;
    }
    if (hasSalesAutocompleteOptions()) {
      if (e.key === "ArrowDown") {
        e.preventDefault();
        moveSalesAutocomplete(1);
        return;
      }
      if (e.key === "ArrowUp") {
        e.preventDefault();
        moveSalesAutocomplete(-1);
        return;
      }
      if (e.key === "Enter") {
        e.preventDefault();
        if (chooseSalesAutocomplete(ST.autocompleteActiveIndex >= 0 ? ST.autocompleteActiveIndex : 0)) return;
      }
      if (e.key === "Escape") {
        e.preventDefault();
        hideSalesAutocomplete();
        focusInput();
        return;
      }
    }

    if ((e.ctrlKey || e.metaKey) && /^Arrow(Up|Down|Left|Right)$/.test(e.key)) {
      e.preventDefault();
      var ctrlDirection = e.key.replace("Arrow", "").toLowerCase();
      var ctrlTarget = findCtrlJumpTarget(ST.selectedCell.row, ST.selectedCell.col, ctrlDirection);
      moveSelection(ctrlTarget.row, ctrlTarget.col);
      return;
    }

    if (e.ctrlKey || e.metaKey) {
      var k = e.key.toLowerCase();
      if (k === "f") {
        e.preventDefault();
        focusSalesFilterInput();
        return;
      }
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
      return;
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

    if (e.altKey && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
      e.preventDefault();
      openSalesAutocompletePicker();
      return;
    }
    if (hasSalesAutocompleteOptions()) {
      if (e.key === "ArrowDown") {
        e.preventDefault();
        moveSalesAutocomplete(1);
        return;
      }
      if (e.key === "ArrowUp") {
        e.preventDefault();
        moveSalesAutocomplete(-1);
        return;
      }
      if (e.key === "Enter") {
        e.preventDefault();
        if (chooseSalesAutocomplete(ST.autocompleteActiveIndex >= 0 ? ST.autocompleteActiveIndex : 0)) return;
      }
      if (e.key === "Escape") {
        e.preventDefault();
        hideSalesAutocomplete();
        return;
      }
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "f") {
      e.preventDefault();
      focusSalesFilterInput();
      return;
    }

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
      if ((e.ctrlKey || e.metaKey) && /^Arrow(Up|Down|Left|Right)$/.test(e.key)) {
        e.preventDefault();
        var ctrlJumpDirection = e.key.replace("Arrow", "").toLowerCase();
        var ctrlJumpTarget = findCtrlJumpTarget(rowIndex, colIndex, ctrlJumpDirection);
        moveSelection(ctrlJumpTarget.row, ctrlJumpTarget.col);
        return;
      }
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
      restoreActiveInputBinding();
      moveInputToCell(ST.selectedCell.row, ST.selectedCell.col);
      syncEditOrigin();
      focusInput();
    }
  }

  function getWorkInfoValue(path) {
    var parts = String(path).split(".");
    if (parts[0] === "supplier") {
      var supplierCur = getCurrentSupplierInfo();
      parts.slice(1).forEach(function (part) {
        supplierCur = supplierCur && supplierCur[part] != null ? supplierCur[part] : "";
      });
      return supplierCur == null ? "" : String(supplierCur);
    }

    var cur = ensureWorkInfoShape();
    parts.forEach(function (part) {
      cur = cur && cur[part] != null ? cur[part] : "";
    });
    return cur == null ? "" : String(cur);
  }

  function setWorkInfoValue(path, value) {
    var parts = String(path).split(".");
    if (parts[0] === "supplier") {
      var supplier = getCurrentSupplierInfo();
      for (var s = 1; s < parts.length - 1; s++) {
        if (!supplier[parts[s]] || typeof supplier[parts[s]] !== "object") {
          supplier[parts[s]] = {};
        }
        supplier = supplier[parts[s]];
      }
      supplier[parts[parts.length - 1]] = value;
      return;
    }

    var cur = ensureWorkInfoShape();
    for (var i = 0; i < parts.length - 1; i++) {
      if (!cur[parts[i]] || typeof cur[parts[i]] !== "object") {
        cur[parts[i]] = {};
      }
      cur = cur[parts[i]];
    }
    cur[parts[parts.length - 1]] = value;
  }

  function findClientRowByCompany(company) {
    var target = String(company || "").trim();
    if (!target) return null;
    var targetNormalized = normalizeCompanyMatchText(target);
    var partialMatch = null;
    var partialScore = Infinity;
    for (var i = 0; i < clientState.rows.length; i++) {
      var row = clientState.rows[i] || {};
      var companyName = String(row.company || "").trim();
      if (!companyName) continue;
      if (companyName === target) return row;
      var rowNormalized = normalizeCompanyMatchText(companyName);
      if (!rowNormalized || !targetNormalized) continue;
      if (rowNormalized === targetNormalized) return row;
      if (rowNormalized.indexOf(targetNormalized) >= 0 || targetNormalized.indexOf(rowNormalized) >= 0) {
        var score = Math.abs(rowNormalized.length - targetNormalized.length);
        if (!partialMatch || score < partialScore) {
          partialMatch = row;
          partialScore = score;
        }
      }
    }
    return partialMatch;
  }

  function getClientRowSupplierMode(clientRow) {
    if (!clientRow) return getCurrentSupplierMode();
    return (
      normalizeSupplierMode(clientRow.supplierMode) ||
      parseSupplierModeLabel(clientRow.supplierMode) ||
      getCurrentSupplierMode() ||
      "corporate"
    );
  }

  function applyClientToWorkInfo(clientRow) {
    if (!clientRow) return;
    ensureWorkInfoShape();
    workState.info.receiver = Object.assign({}, workState.info.receiver || {}, {
      company: clientRow.company || "",
      ceo: clientRow.ceoName || "",
      businessNo: clientRow.businessNo || "",
      address: clientRow.address || "",
      businessType: clientRow.businessType || "",
      businessItem: clientRow.businessItem || "",
    });
    workState.info.supplierMode = getClientRowSupplierMode(clientRow);
  }

  function syncWorkItemDatesToDocumentDate() {
    var docDate = getWorkInfoValue("date");
    var itemDate = isoToMonthDayLabel(docDate);
    if (!itemDate) return;
    workState.items = (workState.items || []).map(function (item) {
      var nextItem = Object.assign({}, item || {});
      var hasAnyValue = ["date", "code", "name", "qty", "price", "supply", "tax", "note"].some(function (key) {
        return nextItem[key] != null && String(nextItem[key]).trim() !== "";
      });
      if (hasAnyValue) nextItem.date = itemDate;
      return nextItem;
    });
    if (workSheetEngine && typeof workSheetEngine.setRows === "function") {
      workSheetEngine.setRows(workState.items);
    }
  }

  function normalizeLookupText(value) {
    return String(value || "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "")
      .replace(/[-_/().]/g, "");
  }

  function getDisplayCompanyName(value) {
    var text = String(value || "").trim();
    if (!text) return "";
    text = text
      .replace(/㈜/g, "")
      .replace(/\(주\)/gi, "")
      .replace(/주식회사/gi, "")
      .replace(/유한회사/gi, "")
      .replace(/유한책임회사/gi, "")
      .replace(/합자회사/gi, "")
      .replace(/합명회사/gi, "")
      .replace(/\([^)]*\)/g, " ")
      .replace(/\s+[^\s]+지점$/g, "")
      .replace(/\s+/g, " ")
      .trim();
    return text;
  }

  function normalizeCompanyMatchText(value) {
    return normalizeLookupText(getDisplayCompanyName(value));
  }

  function getPriceRowsWithValues() {
    return normalizePriceGridRows(priceState.rows, 30).filter(function (row) {
      return rowHasAnyValue(row, priceGridFields);
    });
  }

  function getDistinctPriceClients() {
    var seen = {};
    var items = [];
    getPriceRowsWithValues().forEach(function (row) {
      var company = getDisplayCompanyName(row.client);
      var normalized = normalizeCompanyMatchText(company);
      if (!company || !normalized || seen[normalized]) return;
      seen[normalized] = true;
      items.push(company);
    });
    return items.sort();
  }

  function getAllPriceClientOptions() {
    var seen = {};
    var items = [];

    function pushClient(name) {
      var label = getDisplayCompanyName(name);
      var normalized = normalizeCompanyMatchText(label);
      if (!label || !normalized || seen[normalized]) return;
      seen[normalized] = true;
      items.push(label);
    }

    getDistinctPriceClients().forEach(pushClient);
    clientState.rows.forEach(function (row) {
      pushClient(row && row.company);
    });

    return items.sort();
  }

  function ensureActivePriceClient() {
    var current = String(priceState.activeClient || "").trim();
    if (current) return current;
    var options = getAllPriceClientOptions();
    priceState.activeClient = options[0] || "";
    return priceState.activeClient;
  }

  function getPriceRowsForClient(clientName) {
    var target = normalizeCompanyMatchText(clientName);
    var rows = priceState.rows.filter(function (row) {
      return target && normalizeCompanyMatchText(row && row.client) === target;
    }).map(function (row) {
      return {
        code: row && row.code != null ? String(row.code) : "",
        price: row && row.price != null ? String(row.price) : "",
        name: row && row.name != null ? String(row.name) : "",
      };
    });
    return normalizePriceEditorRows(rows, 80);
  }

  function setPriceRowsForClient(clientName, rows) {
    var label = String(clientName || "").trim();
    var normalized = normalizeCompanyMatchText(label);
    if (!label || !normalized) {
      priceState.rows = normalizePriceGridRows(priceState.rows, 80);
      return;
    }

    var preserved = priceState.rows.filter(function (row) {
      return normalizeCompanyMatchText(row && row.client) !== normalized &&
        rowHasAnyValue(row, priceGridFields);
    }).map(function (row) {
      return Object.assign(emptyPriceRow(), row || {});
    });

    var nextRows = (rows || []).filter(function (row) {
      return rowHasAnyValue(row, priceEditorFields);
    }).map(function (row) {
      return {
        client: label,
        code: row.code || "",
        price: row.price || "",
        name: row.name || "",
      };
    });

    priceState.rows = normalizePriceGridRows(preserved.concat(nextRows), 80);
    priceState.activeClient = label;
  }

  function renamePriceClient(oldName, newName) {
    var oldLabel = String(oldName || "").trim();
    var newLabel = String(newName || "").trim();
    var oldNormalized = normalizeCompanyMatchText(oldLabel);
    var newNormalized = normalizeCompanyMatchText(newLabel);
    if (!oldLabel || !newLabel || !oldNormalized || !newNormalized) return false;
    if (oldNormalized === newNormalized) {
      priceState.activeClient = newLabel;
      priceState.rows = normalizePriceGridRows(
        priceState.rows.map(function (row) {
          var nextRow = Object.assign(emptyPriceRow(), row || {});
          if (normalizeCompanyMatchText(nextRow.client) === oldNormalized) {
            nextRow.client = newLabel;
          }
          return nextRow;
        }),
        80
      );
      return true;
    }

    priceState.rows = normalizePriceGridRows(
      priceState.rows.map(function (row) {
        var nextRow = Object.assign(emptyPriceRow(), row || {});
        if (normalizeCompanyMatchText(nextRow.client) === oldNormalized) {
          nextRow.client = newLabel;
        }
        return nextRow;
      }),
      80
    );
    priceState.activeClient = newLabel;
    return true;
  }

  function deletePriceClient(clientName) {
    var label = String(clientName || "").trim();
    var normalized = normalizeCompanyMatchText(label);
    if (!label || !normalized) return false;

    priceState.rows = normalizePriceGridRows(
      priceState.rows.filter(function (row) {
        return normalizeCompanyMatchText(row && row.client) !== normalized;
      }),
      80
    );

    var remainingClients = getAllPriceClientOptions();
    priceState.activeClient = remainingClients[0] || "";
    return true;
  }

  function findPriceMatches(clientInput, codeInput) {
    var clientToken = normalizeCompanyMatchText(clientInput);
    var codeToken = normalizeLookupText(codeInput);
    return getPriceRowsWithValues().filter(function (row) {
      var rowClient = normalizeCompanyMatchText(row.client);
      var rowCode = normalizeLookupText(row.code);
      var clientMatched =
        !clientToken ||
        rowClient === clientToken ||
        rowClient.indexOf(clientToken) >= 0 ||
        clientToken.indexOf(rowClient) >= 0;
      var codeMatched = !codeToken || rowCode === codeToken || rowCode.indexOf(codeToken) >= 0;
      return clientMatched && codeMatched;
    });
  }

  function findResolvedPriceRow(clientInput, codeInput) {
    var matches = findPriceMatches(clientInput, codeInput);
    if (!matches.length) return null;
    var clientToken = normalizeCompanyMatchText(clientInput);
    var codeToken = normalizeLookupText(codeInput);
    var exact = matches.filter(function (row) {
      var rowClient = normalizeCompanyMatchText(row.client);
      var rowCode = normalizeLookupText(row.code);
      var clientExact = !clientToken || rowClient === clientToken;
      var codeExact = !codeToken || rowCode === codeToken;
      return clientExact && codeExact;
    });
    if (exact.length === 1) return exact[0];
    if (matches.length === 1) return matches[0];
    return null;
  }

  function applyPriceLookupForSalesRow(rows, rowIndex) {
    var row = rows[rowIndex] || emptyRow();
    var client = row.client || "";
    var code = row.code || "";
    if (!code) return;
    var matched = findResolvedPriceRow(client, code);
    if (!matched) return;
    rows[rowIndex] = Object.assign({}, row, {
      client: getDisplayCompanyName(row.client || matched.client || ""),
      code: matched.code || row.code || "",
      name: matched.name || row.name || "",
      price: matched.price || row.price || "",
    });
    recalculateAmountForRow(rows, rowIndex);
  }

  function applyPriceLookupForActiveSalesCell() {
    var activeColumn = salesColumns[ST.selectedCell.col];
    var activeKey = activeColumn ? activeColumn.key : "";
    if (activeKey !== "client" && activeKey !== "code") return;
    applyRowsChange(function (current) {
      applyPriceLookupForSalesRow(current, ST.selectedCell.row);
      return current;
    }, false);
  }

  function updateSalesAutocomplete() {
    if (!ST.inputEl || !ST.autocompleteEl) return;
    var row = ST.rows[ST.selectedCell.row] || emptyRow();
    var key = salesColumns[ST.selectedCell.col].key;
    var raw = ST.inputEl.value || "";
    if (key !== "client" && key !== "code") {
      hideSalesAutocomplete();
      return;
    }
    if (!ST.autocompleteForcedOpen && !String(raw).trim()) {
      hideSalesAutocomplete();
      return;
    }
    var suggestions = [];
    if (key === "client") {
      suggestions = getClientSuggestionOptions(raw);
    } else if (key === "code") {
      suggestions = getCodeSuggestionOptions(row.client, raw);
    }
    var prevActive = ST.autocompleteActiveIndex;
    ST.autocompleteItems = suggestions.slice();
    ST.autocompleteActiveIndex = suggestions.length ? Math.max(0, Math.min(prevActive, suggestions.length - 1)) : -1;
    if (!suggestions.length) {
      ST.autocompleteEl.innerHTML = "";
      ST.autocompleteEl.classList.add("hidden");
      return;
    }
    ST.autocompleteEl.innerHTML = suggestions.map(function (item, index) {
      return (
        '<button type="button" class="autocomplete-item' + (index === ST.autocompleteActiveIndex ? ' active' : '') + '" data-sales-autocomplete="' + index + '">' +
          '<span class="autocomplete-item-main">' + escapeHtml(item.value || "") + '</span>' +
          (item.label ? '<span class="autocomplete-item-sub">' + escapeHtml(item.label) + '</span>' : "") +
        '</button>'
      );
    }).join("");
    ST.autocompleteEl.classList.remove("hidden");
    positionSalesAutocomplete();
    ST.autocompleteEl.querySelectorAll("[data-sales-autocomplete]").forEach(function (btn) {
      btn.addEventListener("mousedown", function (e) {
        e.preventDefault();
      });
      btn.addEventListener("click", function () {
        chooseSalesAutocomplete(Number(btn.getAttribute("data-sales-autocomplete")));
      });
    });
  }

  function hasSalesAutocompleteOptions() {
    return !!(ST.autocompleteEl && ST.autocompleteItems && ST.autocompleteItems.length);
  }

  function positionSalesAutocomplete() {
    if (!ST.autocompleteEl || !ST.inputEl || !ST.host) return;
    var inputRect = ST.inputEl.getBoundingClientRect();
    var hostRect = ST.host.getBoundingClientRect();
    ST.autocompleteEl.style.left = inputRect.left - hostRect.left + "px";
    ST.autocompleteEl.style.top = inputRect.bottom - hostRect.top + 4 + "px";
    ST.autocompleteEl.style.width = Math.max(220, inputRect.width) + "px";
  }

  function hideSalesAutocomplete() {
    if (!ST.autocompleteEl) return;
    ST.autocompleteItems = [];
    ST.autocompleteActiveIndex = -1;
    ST.autocompleteForcedOpen = false;
    ST.autocompleteEl.innerHTML = "";
    ST.autocompleteEl.classList.add("hidden");
  }

  function chooseSalesAutocomplete(index) {
    if (!ST.autocompleteItems || index < 0 || index >= ST.autocompleteItems.length) return false;
    var item = ST.autocompleteItems[index];
    if (!item || !ST.inputEl) return false;
    ST.inputEl.value = item.value || "";
    ST.autocompleteForcedOpen = false;
    setValue(ST.selectedCell.row, ST.selectedCell.col, ST.inputEl.value);
    applyPriceLookupForActiveSalesCell();
    hideSalesAutocomplete();
    syncInputOverlay();
    focusInput();
    return true;
  }

  function moveSalesAutocomplete(delta) {
    if (!hasSalesAutocompleteOptions()) return false;
    var next = ST.autocompleteActiveIndex + delta;
    if (next < 0) next = ST.autocompleteItems.length - 1;
    if (next >= ST.autocompleteItems.length) next = 0;
    ST.autocompleteActiveIndex = next;
    updateSalesAutocomplete();
    return true;
  }

  function openSalesAutocompletePicker() {
    var activeColumn = salesColumns[ST.selectedCell.col];
    var activeKey = activeColumn ? activeColumn.key : "";
    if (activeKey !== "client" && activeKey !== "code") return false;
    if (!ST.editMode) {
      snapshotEditOrigin();
      ST.editMode = true;
    }
    ST.autocompleteForcedOpen = true;
    syncInputOverlay({ ensureVisible: true });
    updateSalesAutocomplete();
    if (!hasSalesAutocompleteOptions()) {
      focusInput();
      return false;
    }
    requestAnimationFrame(function () {
      if (!ST.inputEl) return;
      ST.inputEl.focus();
    });
    return true;
  }

  function focusSalesFilterInput() {
    var filterKeyword = document.getElementById("status-filter-keyword");
    if (!filterKeyword) return false;
    filterKeyword.focus();
    filterKeyword.select();
    return true;
  }

  function getClientSuggestionOptions(keyword) {
    var token = normalizeCompanyMatchText(keyword);
    return getDistinctPriceClients()
      .filter(function (company) {
        var normalized = normalizeCompanyMatchText(company);
        return !token || normalized.indexOf(token) >= 0 || token.indexOf(normalized) >= 0;
      })
      .slice(0, 30)
      .map(function (company) {
        return { value: company, label: company };
      });
  }

  function getCodeSuggestionOptions(clientInput, codeInput) {
    var rows = findPriceMatches(clientInput, codeInput).slice(0, 40);
    return rows.map(function (row) {
      return {
        value: row.code || "",
        label: [row.name || "", row.price ? formatDisplayNumber(row.price) : ""].filter(Boolean).join(" / "),
      };
    });
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
      if (String(nameText).length > 42) nameClass += " tightest";
      else if (String(nameText).length > 34) nameClass += " tighter";
      else if (String(nameText).length > 22) nameClass += " tight";
      if (String(codeText).length > 14) codeClass += " tighter";
      else if (String(codeText).length > 10) codeClass += " tight";
      if (String(noteText).length > 14) noteClass += " tighter";
      else if (String(noteText).length > 10) noteClass += " tight";
      rows += "<tr>";
      rows += statementCell(formatStatementItemDate(item.date || ""), "item-date center");
      rows += statementCell(nameText, nameClass);
      rows += statementCell(codeText, codeClass);
      rows += statementCell(noteText, noteClass);
      rows += statementCell(formatDisplayNumber(item.qty), "item-num right");
      rows += statementCell(formatDisplayNumber(item.price), "item-num right");
      rows += statementCell(formatDisplayNumber(item.supply), "item-num right");
      rows += statementCell(formatDisplayNumber(item.tax), "item-num right");
      rows += "</tr>";
    }
    return rows;
  }

  function renderStatementCopy(copyClass, tagLabel, items, totals) {
    var supplier = getCurrentSupplierInfo();
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
            statementCell(formatDisplayNumber(totals.supply), "statement-value right") +
            statementHead("세액", "center") +
            statementCell(formatDisplayNumber(totals.tax), "statement-value right") +
            statementHead("합계", "center") +
            statementCell(formatDisplayNumber(totals.grand), "statement-value right") +
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
    ensureWorkInfoShape();
    var items = workState.items || [];
    var totals = getStatementTotals(items.filter(function (item) {
      if (!item) return false;
      return ["date", "code", "name", "qty", "price", "supply", "tax", "note"].some(function (key) {
        return item[key] != null && String(item[key]).trim() !== "";
      });
    }));

    return (
      '<div class="workdoc">' +
        '<div class="workdoc-top">' +
          '<div class="workdoc-card">' +
            '<div class="workdoc-head">' +
              '<div class="workdoc-title">' + icon("building") + ' 공급자(우리 회사)</div>' +
              '<div class="workdoc-head-actions">' +
                '<div class="sub">구분</div>' +
                '<select class="workdoc-input workdoc-select" id="supplier-mode">' +
                  '<option value="corporate"' + (getCurrentSupplierMode() === "corporate" ? " selected" : "") + '>법인</option>' +
                  '<option value="personal"' + (getCurrentSupplierMode() === "personal" ? " selected" : "") + '>개인</option>' +
                '</select>' +
                '<button type="button" class="soft-btn" id="btn-supplier-reset-current">기본값 불러오기</button>' +
              '</div>' +
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
          '<div class="workdoc-card">' +
            '<div class="workdoc-head">' +
              '<div class="workdoc-title">' + icon("clipboard") + ' 작업요약</div>' +
            "</div>" +
            '<div class="workdoc-meta">' +
              '<div class="workdoc-meta-row">' +
                '<div class="workdoc-label">작성일자</div>' +
                '<input type="date" class="workdoc-input" data-work-info="date" value="' + escapeAttr(getWorkInfoValue("date")) + '" />' +
              '</div>' +
              '<div class="workdoc-meta-row">' +
                '<div class="workdoc-label">품목 수</div>' +
                '<input type="text" class="workdoc-input" id="work-summary-count" value="' + escapeAttr(String(items.filter(function (item) { return item && ["date", "code", "name", "qty", "price", "supply", "tax", "note"].some(function (key) { return item[key] != null && String(item[key]).trim() !== ""; }); }).length)) + '" readonly />' +
              '</div>' +
              '<div class="workdoc-summary">' +
                '<div class="workdoc-summary-item">' +
                  '<div class="workdoc-summary-label">공급가액</div>' +
                  '<div class="workdoc-summary-value" id="work-summary-supply">' + escapeHtml(formatDisplayNumber(totals.supply)) + '</div>' +
                '</div>' +
                '<div class="workdoc-summary-item">' +
                  '<div class="workdoc-summary-label">세액</div>' +
                  '<div class="workdoc-summary-value" id="work-summary-tax">' + escapeHtml(formatDisplayNumber(totals.tax)) + '</div>' +
                '</div>' +
                '<div class="workdoc-summary-item">' +
                  '<div class="workdoc-summary-label">합계</div>' +
                  '<div class="workdoc-summary-value" id="work-summary-grand">' + escapeHtml(formatDisplayNumber(totals.grand)) + '</div>' +
                '</div>' +
                '<div class="workdoc-summary-item">' +
                  '<div class="workdoc-summary-label">수량합계</div>' +
                  '<div class="workdoc-summary-value" id="work-summary-qty">' + escapeHtml(formatDisplayNumber(totals.qty)) + '</div>' +
                '</div>' +
              '</div>' +
            '</div>' +
          "</div>" +
        "</div>" +
        '<div class="workdoc-card">' +
          '<div id="work-grid-host"></div>' +
        "</div>" +
      "</div>"
    );
  }

  function renderClientTab() {
    clientState.rows = normalizeClientGridRows(clientState.rows, 20);
    return (
      '<div class="clientdoc">' +
        '<div class="clientdoc-card">' +
          '<div class="clientdoc-head">' +
            '<div class="clientdoc-title">' + icon("building") + ' 업체 리스트</div>' +
            '<div class="sub">매출현황처럼 셀 단위 선택/복사/붙여넣기</div>' +
          '</div>' +
          '<div id="client-grid-host"></div>' +
        '</div>' +
      '</div>'
    );
  }

  function renderPriceTab() {
    priceState.rows = normalizePriceGridRows(priceState.rows, 40);
    var activeClient = ensureActivePriceClient();
    var clientOptions = getAllPriceClientOptions();
    var priceTitle = activeLedgerKind === "purchase" ? "매입단가 리스트" : "매출단가 리스트";
    return (
      '<div class="clientdoc">' +
        '<div class="price-layout">' +
          '<div class="price-sidebar">' +
            '<div class="clientdoc-title">' + icon("building") + ' 업체 목록</div>' +
            '<div style="position:relative">' +
            '<input type="text" class="field" id="price-active-client" style="height:36px" placeholder="업체명 입력 또는 선택" value="' + escapeAttr(activeClient) + '" />' +
            '<div id="price-client-list" class="autocomplete-popover hidden"></div>' +
            '</div>' +
            '<div class="price-main-actions">' +
              '<button type="button" class="soft-btn" id="btn-price-open">불러오기</button>' +
              '<button type="button" class="soft-btn" id="btn-price-new">새 업체</button>' +
            '</div>' +
            '<div class="price-client-list">' +
              clientOptions.map(function (company) {
                var activeClass = normalizeCompanyMatchText(company) === normalizeCompanyMatchText(activeClient) ? " active" : "";
                return '<button type="button" class="price-client-item' + activeClass + '" data-price-client="' + escapeAttr(company) + '">' + escapeHtml(company) + '</button>';
              }).join("") +
            '</div>' +
          '</div>' +
          '<div class="clientdoc-card">' +
            '<div class="clientdoc-head">' +
              '<div class="clientdoc-title">' + icon("tags") + " " + priceTitle + '</div>' +
              '<div class="price-main-actions">' +
                '<button type="button" class="soft-btn" id="btn-price-rename">업체명 수정</button>' +
                '<button type="button" class="soft-btn" id="btn-price-delete">업체 삭제</button>' +
              '</div>' +
            '</div>' +
            '<div class="sub" style="margin-bottom:10px">' + escapeHtml(activeClient ? activeClient + ' 단가표' : '왼쪽에서 업체를 선택하거나 새 업체를 입력하세요') + '</div>' +
            '<div id="price-grid-host"></div>' +
          '</div>' +
        '</div>' +
      '</div>'
    );
  }

  function capturePriceSidebarScroll() {
    var listEl = document.querySelector(".price-client-list");
    if (!listEl) return;
    priceState.sidebarScrollTop = listEl.scrollTop || 0;
  }

  function restorePriceSidebarScroll() {
    var listEl = document.querySelector(".price-client-list");
    if (!listEl) return;
    listEl.scrollTop = priceState.sidebarScrollTop || 0;
  }

  function getPriceClientSuggestionOptions(keyword) {
    var token = normalizeCompanyMatchText(keyword);
    return getAllPriceClientOptions()
      .filter(function (company) {
        var normalized = normalizeCompanyMatchText(company);
        return !token || normalized.indexOf(token) >= 0 || token.indexOf(normalized) >= 0;
      })
      .slice(0, 30)
      .map(function (company) {
        return { value: company, label: company };
      });
  }

  function updatePriceClientAutocomplete(inputEl) {
    var popover = document.getElementById("price-client-list");
    if (!inputEl || !popover) return;
    var suggestions = getPriceClientSuggestionOptions(inputEl.value || "");
    if (!suggestions.length) {
      popover.innerHTML = "";
      popover.classList.add("hidden");
      return;
    }
    popover.innerHTML = suggestions.map(function (item) {
      return '<button type="button" class="autocomplete-item" data-price-suggestion="' + escapeAttr(item.value) + '"><span class="autocomplete-item-main">' + escapeHtml(item.value) + '</span></button>';
    }).join("");
    popover.classList.remove("hidden");
    popover.style.left = "0";
    popover.style.top = "40px";
    popover.style.width = "100%";
    popover.querySelectorAll("[data-price-suggestion]").forEach(function (btn) {
      btn.addEventListener("mousedown", function (e) {
        e.preventDefault();
      });
      btn.addEventListener("click", function () {
        inputEl.value = btn.getAttribute("data-price-suggestion") || "";
        popover.classList.add("hidden");
      });
    });
  }

  function hidePriceClientAutocomplete() {
    var popover = document.getElementById("price-client-list");
    if (!popover) return;
    popover.innerHTML = "";
    popover.classList.add("hidden");
  }

  function parseSalesMonthDay(value) {
    var text = String(value || "").trim();
    if (!text) return null;
    var match =
      text.match(/^(\d{1,2})\s*월\s*(\d{1,2})\s*일$/) ||
      text.match(/^(\d{1,2})\s*\/\s*(\d{1,2})$/);
    if (!match) return null;
    var month = Number(match[1]);
    var day = Number(match[2]);
    if (!month || !day) return null;
    return { month: month, day: day };
  }

  function getManageSourceRows() {
    var keys = salesColumns.map(function (col) { return col.key; });
    return ST.rows.filter(function (row) {
      return rowHasAnyValue(row, keys);
    });
  }

  function getManageAvailableMonths(rows) {
    var seen = {};
    rows.forEach(function (row) {
      var parsed = parseSalesMonthDay(row.date);
      if (parsed) seen[parsed.month] = true;
    });
    var months = Object.keys(seen).map(function (value) { return Number(value); }).sort(function (a, b) { return a - b; });
    return months.length ? months : [new Date().getMonth() + 1];
  }

  function normalizeManageMonthValue(value, fallback) {
    if (String(value || "") === "all") return "all";
    var num = Number(value);
    if (!num || num < 1 || num > 12) return fallback || "all";
    return String(num);
  }

  function shiftManageRangeFromStart(nextStartValue) {
    var nextStart = normalizeManageMonthValue(nextStartValue, "all");
    var currentStart = normalizeManageMonthValue(manageState.startMonth, "all");
    var currentEnd = normalizeManageMonthValue(manageState.endMonth, "all");

    manageState.startMonth = nextStart;

    if (nextStart === "all") {
      manageState.endMonth = "all";
      return;
    }

    if (currentStart === "all" || currentEnd === "all") {
      manageState.endMonth = nextStart;
      return;
    }

    var span = Number(currentEnd) - Number(currentStart);
    var shiftedEnd = Number(nextStart) + span;
    if (shiftedEnd < 1) shiftedEnd = 1;
    if (shiftedEnd > 12) shiftedEnd = 12;
    manageState.endMonth = String(shiftedEnd);
  }

  function getManageClientOptions(rows) {
    var seen = {};
    var clients = [];
    rows.forEach(function (row) {
      var client = getDisplayCompanyName(row.client);
      if (!client) return;
      var normalized = normalizeCompanyMatchText(client);
      if (seen[normalized]) return;
      seen[normalized] = true;
      clients.push(client);
    });
    return clients.sort(function (a, b) {
      return a.localeCompare(b, "ko");
    });
  }

  function ensureManageStateDefaults() {
    var rows = getManageSourceRows();
    var months = getManageAvailableMonths(rows);
    var clients = getManageClientOptions(rows);
    if (!manageState.startMonth || (manageState.startMonth !== "all" && months.indexOf(Number(manageState.startMonth)) < 0)) {
      manageState.startMonth = String(months[0]);
    }
    if (!manageState.endMonth || (manageState.endMonth !== "all" && months.indexOf(Number(manageState.endMonth)) < 0)) {
      manageState.endMonth = String(months[months.length - 1]);
    }
    if (!manageState.client) {
      manageState.client = "";
    } else if (manageState.client && clients.indexOf(manageState.client) < 0) {
      manageState.client = "";
    }
  }

  function getManageFilteredRows() {
    ensureManageStateDefaults();
    var startMonth = normalizeManageMonthValue(manageState.startMonth, "all");
    var endMonth = normalizeManageMonthValue(manageState.endMonth, "all");
    var selectedClient = String(manageState.client || "").trim();
    return getManageSourceRows().filter(function (row) {
      var parsed = parseSalesMonthDay(row.date);
      if (startMonth !== "all" && endMonth !== "all") {
        var monthMin = Math.min(Number(startMonth), Number(endMonth));
        var monthMax = Math.max(Number(startMonth), Number(endMonth));
        if (parsed && (parsed.month < monthMin || parsed.month > monthMax)) {
          return false;
        }
      }
      if (selectedClient && normalizeCompanyMatchText(row.client) !== normalizeCompanyMatchText(selectedClient)) {
        return false;
      }
      return true;
    });
  }

  function buildManageProductSummary(rows) {
    var map = {};
    rows.forEach(function (row) {
      var key = String(row.name || row.code || "").trim();
      if (!key) return;
      if (!map[key]) {
        map[key] = {
          name: String(row.name || "").trim() || key,
          qty: 0,
          amount: 0,
        };
      }
      map[key].qty += parseCalcNumber(row.qty) || 0;
      map[key].amount += parseCalcNumber(row.amount) || 0;
    });
    return Object.keys(map)
      .map(function (key) { return map[key]; })
      .sort(function (a, b) { return b.amount - a.amount; })
      .slice(0, 6);
  }

  function renderManageTab() {
    ensureManageStateDefaults();
    var sourceRows = getManageSourceRows();
    var months = getManageAvailableMonths(sourceRows);
    var clients = getManageClientOptions(sourceRows);
    var filteredRows = getManageFilteredRows();
    var totalQty = filteredRows.reduce(function (sum, row) {
      return sum + (parseCalcNumber(row.qty) || 0);
    }, 0);
    var totalAmount = filteredRows.reduce(function (sum, row) {
      return sum + (parseCalcNumber(row.amount) || 0);
    }, 0);
    var manageLabel = activeLedgerKind === "purchase" ? "매입" : "매출";

    function renderMonthOptions(selectedValue) {
      return '<option value="all"' + (String(selectedValue) === "all" ? " selected" : "") + '>전체</option>' +
      months.map(function (month) {
        return '<option value="' + month + '"' + (String(selectedValue) === String(month) ? " selected" : "") + '>' + month + '월</option>';
      }).join("");
    }

    function renderClientOptions(selectedValue) {
      return '<option value="">전체 업체</option>' + clients.map(function (client) {
        return '<option value="' + escapeAttr(client) + '"' + (selectedValue === client ? " selected" : "") + '>' + escapeHtml(client) + '</option>';
      }).join("");
    }

    return (
      '<div class="manage-page">' +
        '<div class="manage-toolbar-card">' +
          '<div class="manage-toolbar-grid">' +
            '<div class="manage-toolbar-item"><label>시작월</label><select id="manage-start-month">' + renderMonthOptions(manageState.startMonth) + '</select></div>' +
            '<div class="manage-toolbar-item"><label>종료월</label><select id="manage-end-month">' + renderMonthOptions(manageState.endMonth) + '</select></div>' +
            '<div class="manage-toolbar-item manage-toolbar-client"><label>업체</label><select id="manage-client">' + renderClientOptions(manageState.client) + '</select></div>' +
            '<div class="manage-toolbar-actions">' +
              '<button type="button" class="soft-btn" id="btn-manage-search">조회</button>' +
              '<button type="button" class="soft-btn" id="btn-manage-reset">초기화</button>' +
            '</div>' +
          '</div>' +
          '<div class="manage-toolbar-help">조건 선택 후 <strong>조회</strong>를 누르면 집계가 갱신됩니다.</div>' +
        '</div>' +
        '<div class="manage-summary-grid">' +
          '<div class="manage-summary-card"><div class="manage-summary-label">조회건수</div><div class="manage-summary-value">' + escapeHtml(formatDisplayNumber(filteredRows.length)) + '</div></div>' +
          '<div class="manage-summary-card"><div class="manage-summary-label">수량합계</div><div class="manage-summary-value">' + escapeHtml(formatDisplayNumber(totalQty)) + '</div></div>' +
          '<div class="manage-summary-card"><div class="manage-summary-label">' + manageLabel + '합계</div><div class="manage-summary-value">' + escapeHtml(formatDisplayNumber(totalAmount)) + '</div></div>' +
        '</div>' +
        '<div class="manage-layout">' +
          '<div class="manage-table-card">' +
            '<div class="manage-card-title">조회 결과</div>' +
            '<div class="manage-table-wrap">' +
              '<table class="manage-table">' +
                '<thead><tr><th>월일</th><th>업체명</th><th>코드</th><th>품명</th><th class="right">수량</th><th class="right">단가</th><th class="right">매출금액</th></tr></thead>' +
                '<tbody>' +
                  (filteredRows.length ? filteredRows.map(function (row) {
                    return '<tr>' +
                      '<td>' + escapeHtml(String(row.date || "")) + '</td>' +
                      '<td>' + escapeHtml(getDisplayCompanyName(row.client || "")) + '</td>' +
                      '<td>' + escapeHtml(String(row.code || "")) + '</td>' +
                      '<td>' + escapeHtml(String(row.name || "")) + '</td>' +
                      '<td class="right">' + escapeHtml(formatDisplayNumber(row.qty || "")) + '</td>' +
                      '<td class="right">' + escapeHtml(formatDisplayNumber(row.price || "")) + '</td>' +
                      '<td class="right">' + escapeHtml(formatDisplayNumber(row.amount || "")) + '</td>' +
                    '</tr>';
                  }).join("") : '<tr><td colspan="7" class="center muted">조건에 맞는 데이터가 없습니다.</td></tr>') +
                '</tbody>' +
              '</table>' +
            '</div>' +
          '</div>' +
        '</div>' +
      '</div>'
    );
  }

  function renderClosingAttendanceTab() {
    ensureClosingAttendanceState();
    var outsourceMeta = getClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor);
    var outsourceLockDay = getClosingOutsourceLockDay(outsourceMeta, getClosingDaysInMonth());
    var monthOptions = getClosingMonthOptions().map(function (option) {
      return (
        '<option value="' + escapeHtml(option.value) + '"' +
        (option.value === closingState.attendanceMonth ? " selected" : "") +
        ">" + escapeHtml(option.label) + "</option>"
      );
    }).join("");
    var outsourceVendorOptions = closingOutsourceVendors.map(function (option) {
      return (
        '<option value="' + escapeHtml(option.value) + '"' +
        (option.value === closingState.outsourceVendor ? " selected" : "") +
        ">" + escapeHtml(option.label) + "</option>"
      );
    }).join("");
    var isEmployeeView = closingState.attendanceView !== "outsource";
    var employeeViewToggle = isEmployeeView
      ? '<div class="toolbar-segment" style="display:inline-flex;gap:6px;">' +
          '<button type="button" class="soft-btn' + (closingState.employeeViewMode !== "summary" ? ' active-filter' : '') + '" data-closing-employee-view="entry">수기 입력</button>' +
          '<button type="button" class="soft-btn' + (closingState.employeeViewMode === "summary" ? ' active-filter' : '') + '" data-closing-employee-view="summary">연동 화면</button>' +
        '</div>'
      : "";
    return (
      '<div class="closing-page">' +
        '<div class="closing-card closing-sheet-card">' +
          '<div class="closing-sheet-toolbar">' +
            '<div>' +
              '<div class="closing-title">' + icon("sheet") + ' 근무(급여)</div>' +
              '<div class="closing-copy">' + escapeHtml(isEmployeeView
                ? "정직원 수기표/연동표를 전환하고 지문 원본으로 자동 반영할 수 있습니다."
                : "아웃소싱 업체별 월 시트를 먼저 관리하고, 다음 단계에서 지문 원본 자동계산을 연결합니다.") + '</div>' +
            '</div>' +
            '<div class="closing-inline-controls">' +
              '<button type="button" class="soft-btn' + (isEmployeeView ? ' active-filter' : '') + '" data-closing-attendance-view="employee">정직원</button>' +
              '<button type="button" class="soft-btn' + (!isEmployeeView ? ' active-filter' : '') + '" data-closing-attendance-view="outsource">아웃소싱</button>' +
              employeeViewToggle +
              '<label class="closing-inline-label" for="closing-attendance-month">대상월</label>' +
              '<select id="closing-attendance-month" class="closing-inline-select">' +
                monthOptions +
              '</select>' +
              '<input type="text" id="closing-attendance-search" class="closing-inline-input" placeholder="이름 찾기" value="' + escapeAttr(closingState.attendanceSearch || "") + '" />' +
              '<button type="button" class="soft-btn" id="closing-attendance-add">' + (isEmployeeView ? "직원 추가" : "인원 추가") + '</button>' +
              (isEmployeeView
                ? '<label class="closing-inline-file">' +
                    '<span class="closing-inline-file-btn">지문 원본 선택</span>' +
                    '<input type="file" id="closing-fingerprint-file" accept=".xls,.xml,.xlsx" />' +
                  '</label>' +
                  '<button type="button" class="soft-btn active-filter" id="closing-fingerprint-fill">정직원 자동채우기</button>' +
                  '<span class="closing-inline-note">' + escapeHtml(closingFingerprintSourceFile ? closingFingerprintSourceFile.name : "선택된 파일 없음") + '</span>'
                : '') +
              (!isEmployeeView
                ? '<label class="closing-inline-label" for="closing-outsource-vendor">업체</label>' +
                  '<select id="closing-outsource-vendor" class="closing-inline-select">' +
                    outsourceVendorOptions +
                  '</select>' +
                  '<label class="closing-inline-label" for="closing-outsource-lock-day">수정일</label>' +
                  '<input type="text" id="closing-outsource-lock-day" class="closing-inline-input right" style="width:64px" placeholder="일" value="' + escapeAttr(outsourceLockDay ? String(outsourceLockDay) : "") + '" />' +
                  '<button type="button" class="soft-btn" id="closing-outsource-lock-today">오늘 고정</button>' +
                  '<label class="closing-inline-file">' +
                    '<span class="closing-inline-file-btn">지문 원본 선택</span>' +
                    '<input type="file" id="closing-fingerprint-file" accept=".xls,.xml,.xlsx" />' +
                  '</label>' +
                  '<button type="button" class="soft-btn active-filter" id="closing-fingerprint-fill">자동채우기</button>' +
                  '<span class="closing-inline-note">' + escapeHtml(closingFingerprintSourceFile ? closingFingerprintSourceFile.name : "선택된 파일 없음") + '</span>'
                : '') +
            '</div>' +
          '</div>' +
          (!isEmployeeView
            ? '<div class="closing-rule-list closing-rule-list-inline">' +
                '<div class="closing-rule-item"><strong>30분 단위:</strong> 5분 59초까지는 유지, 6분부터 30분 차감</div>' +
                '<div class="closing-rule-item"><strong>반일:</strong> 오전 3.7 / 오후 4.3 기준</div>' +
                '<div class="closing-rule-item"><strong>퇴근 인정:</strong> 17:55 이후는 18:00 인정</div>' +
                '<div class="closing-rule-item"><strong>야근:</strong> 20:50 이후 2.5시간</div>' +
                '<div class="closing-rule-item"><strong>수당:</strong> 결근이 없으면 주휴/만근 유지</div>' +
              '</div>'
            : '') +
        '</div>' +
        '<div class="closing-card closing-sheet-card">' +
          (isEmployeeView
            ? (closingState.employeeViewMode === "summary" ? renderClosingEmployeeSummaryMatrix() : renderClosingEmployeeEntryMatrix())
            : renderClosingOutsourceMatrix()) +
        '</div>' +
      '</div>'
    );
  }

  function getClosingInputCellInfo(input) {
    if (!(input instanceof HTMLInputElement)) return null;
    var outsourceRow = input.getAttribute("data-outsource-row");
    var outsourceField = input.getAttribute("data-outsource-field");
    if (outsourceRow != null && outsourceField) {
      var fieldMatch = String(outsourceField).match(/^d(\d+)$/);
      var col = fieldMatch ? Number(fieldMatch[1]) : (outsourceField === "payCalc" ? 1001 : (outsourceField === "bonus" ? 1002 : (outsourceField === "payAmount" ? 1003 : -1)));
      if (col >= 0) {
        return {
          type: "outsource",
          row: Number(outsourceRow),
          col: col,
        };
      }
    }
    var employeeRow = input.getAttribute("data-employee-row");
    var employeeField = input.getAttribute("data-employee-field");
    if (employeeRow != null && employeeField) {
      var employeeMatch = String(employeeField).match(/^d(\d+)$/);
      if (employeeMatch) {
        return {
          type: "employee",
          row: Number(employeeRow),
          col: Number(employeeMatch[1]),
        };
      }
    }
    return null;
  }

  function focusClosingInputByCellInfo(wrap, info) {
    if (!wrap || !info) return false;
    var selector = "";
    if (info.type === "outsource") {
      var fieldName = info.col <= 31 ? ("d" + info.col) : (info.col === 1001 ? "payCalc" : (info.col === 1002 ? "bonus" : "payAmount"));
      selector = 'input[data-outsource-row="' + info.row + '"][data-outsource-field="' + fieldName + '"]';
    } else if (info.type === "employee") {
      selector = 'input[data-employee-row="' + info.row + '"][data-employee-field="d' + info.col + '"]';
    }
    if (!selector) return false;
    var next = wrap.querySelector(selector);
    if (!(next instanceof HTMLInputElement)) return false;
    next.focus();
    next.select();
    return true;
  }

  function autoFocusClosingFirstInput(wrap) {
    if (!wrap) return;
    var active = document.activeElement;
    if (active instanceof HTMLInputElement && wrap.contains(active)) return;
    var first = wrap.querySelector(
      'input[data-outsource-row][data-outsource-field^="d"], input[data-employee-row][data-employee-field^="d"]'
    );
    if (first instanceof HTMLInputElement) {
      first.focus();
      first.select();
    }
  }

  function bindClosingInputNavigation(wrap) {
    if (!wrap || wrap.__closingNavigationBound) return;
    wrap.__closingNavigationBound = true;
    wrap.addEventListener("keydown", function (e) {
      var target = e.target;
      if (!(target instanceof HTMLInputElement)) return;
      var info = getClosingInputCellInfo(target);
      if (!info) return;
      var nextInfo = null;
      if (e.key === "Enter" || e.key === "ArrowDown") {
        nextInfo = { type: info.type, row: info.row + 1, col: info.col };
      } else if (e.key === "ArrowUp") {
        nextInfo = { type: info.type, row: info.row - 1, col: info.col };
      } else if (e.key === "ArrowLeft") {
        nextInfo = { type: info.type, row: info.row, col: info.col - 1 };
      } else if (e.key === "ArrowRight") {
        nextInfo = { type: info.type, row: info.row, col: info.col + 1 };
      }
      if (!nextInfo) return;
      if (nextInfo.row < 0) return;
      if (info.type === "employee") {
        if (nextInfo.col < 1) nextInfo.col = 1;
        if (nextInfo.col > 31) nextInfo.col = 31;
      } else if (info.type === "outsource") {
        if (nextInfo.col < 1) nextInfo.col = 1;
        if (nextInfo.col > 1003) nextInfo.col = 1003;
      }
      if (!focusClosingInputByCellInfo(wrap, nextInfo)) return;
      e.preventDefault();
    });
  }

  function renderClosingSalesCloseTab() {
    ensureClosingAttendanceState();
    var rows = getClosingSalesCloseRows(closingState.attendanceMonth);
    var monthOptions = getClosingMonthOptions().map(function (option) {
      return '<option value="' + escapeAttr(option.value) + '"' + (option.value === closingState.attendanceMonth ? " selected" : "") + ">" + escapeHtml(option.label) + "</option>";
    }).join("");
    var body = rows.map(function (row, index) {
      return '<tr>' +
        '<td class="center">' + escapeHtml(row.no || String(index + 1)) + '</td>' +
        '<td><input type="text" class="closing-inline-input" data-sales-close-row="' + index + '" data-sales-close-field="division" value="' + escapeAttr(row.division || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input" data-sales-close-row="' + index + '" data-sales-close-field="company" value="' + escapeAttr(row.company || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input" data-sales-close-row="' + index + '" data-sales-close-field="closeDate" value="' + escapeAttr(row.closeDate || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input" data-sales-close-row="' + index + '" data-sales-close-field="taxIssueDate" value="' + escapeAttr(row.taxIssueDate || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input right" data-sales-close-row="' + index + '" data-sales-close-field="amount" value="' + escapeAttr(row.amount || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input center" data-sales-close-row="' + index + '" data-sales-close-field="mailSent" value="' + escapeAttr(row.mailSent || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input center" data-sales-close-row="' + index + '" data-sales-close-field="mailReply" value="' + escapeAttr(row.mailReply || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input center" data-sales-close-row="' + index + '" data-sales-close-field="issueConfirm" value="' + escapeAttr(row.issueConfirm || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input" data-sales-close-row="' + index + '" data-sales-close-field="note" value="' + escapeAttr(row.note || "") + '" /></td>' +
      '</tr>';
    }).join("");
    return (
      '<div class="closing-page">' +
        '<div class="closing-card closing-card-wide">' +
          '<div class="closing-title">' + icon("scroll") + ' 매출 업체별 마감 확인</div>' +
          '<div class="closing-copy">' + escapeHtml(getClosingAttendanceYear() + "년 " + closingState.attendanceMonth + " 매출 마감내역/세금계산서 확인표") + '</div>' +
          '<div class="closing-inline-controls" style="margin-top:8px"><label class="closing-inline-label" for="closing-close-month">대상월</label><select id="closing-close-month" class="closing-inline-select">' + monthOptions + '</select></div>' +
        '</div>' +
        '<div class="closing-card closing-card-wide">' +
          '<div class="closing-table-simple-wrap">' +
            '<table class="closing-table-simple">' +
              '<thead><tr><th style="width:48px">NO.</th><th style="width:72px">구분</th><th style="width:240px">업체명</th><th style="width:110px">마감일</th><th style="width:110px">계산서발행</th><th style="width:150px">매출액</th><th style="width:90px">메일발송</th><th style="width:90px">회신메일</th><th style="width:90px">발행확인</th><th>비고</th></tr></thead>' +
              '<tbody>' + body + '</tbody>' +
            '</table>' +
          '</div>' +
        '</div>' +
      '</div>'
    );
  }

  function renderClosingPurchaseCloseTab() {
    ensureClosingAttendanceState();
    var rows = getClosingPurchaseCloseRows(closingState.attendanceMonth);
    var monthOptions = getClosingMonthOptions().map(function (option) {
      return '<option value="' + escapeAttr(option.value) + '"' + (option.value === closingState.attendanceMonth ? " selected" : "") + ">" + escapeHtml(option.label) + "</option>";
    }).join("");
    var body = rows.map(function (row, index) {
      return '<tr>' +
        '<td class="center">' + escapeHtml(row.no || String(index + 1)) + '</td>' +
        '<td><input type="text" class="closing-inline-input" data-purchase-close-row="' + index + '" data-purchase-close-field="division" value="' + escapeAttr(row.division || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input" data-purchase-close-row="' + index + '" data-purchase-close-field="company" value="' + escapeAttr(row.company || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input" data-purchase-close-row="' + index + '" data-purchase-close-field="closeIssueDate" value="' + escapeAttr(row.closeIssueDate || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input right" data-purchase-close-row="' + index + '" data-purchase-close-field="supplyAmount" value="' + escapeAttr(row.supplyAmount || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input right" data-purchase-close-row="' + index + '" data-purchase-close-field="totalAmount" value="' + escapeAttr(row.totalAmount || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input center" data-purchase-close-row="' + index + '" data-purchase-close-field="detailCheck" value="' + escapeAttr(row.detailCheck || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input center" data-purchase-close-row="' + index + '" data-purchase-close-field="issueConfirm" value="' + escapeAttr(row.issueConfirm || "") + '" /></td>' +
        '<td><input type="text" class="closing-inline-input" data-purchase-close-row="' + index + '" data-purchase-close-field="note" value="' + escapeAttr(row.note || "") + '" /></td>' +
      '</tr>';
    }).join("");
    return (
      '<div class="closing-page">' +
        '<div class="closing-card closing-card-wide">' +
          '<div class="closing-title">' + icon("folder") + ' 매입 업체별 마감 확인</div>' +
          '<div class="closing-copy">' + escapeHtml(getClosingAttendanceYear() + "년 " + closingState.attendanceMonth + " 매입 마감내역/세금계산서 확인표") + '</div>' +
          '<div class="closing-inline-controls" style="margin-top:8px"><label class="closing-inline-label" for="closing-close-month">대상월</label><select id="closing-close-month" class="closing-inline-select">' + monthOptions + '</select></div>' +
        '</div>' +
        '<div class="closing-card closing-card-wide">' +
          '<div class="closing-table-simple-wrap">' +
            '<table class="closing-table-simple">' +
              '<thead><tr><th style="width:48px">NO.</th><th style="width:72px">구분</th><th style="width:240px">업체명</th><th style="width:140px">마감일/발행일</th><th style="width:160px">매입금액(공급가액)</th><th style="width:130px">합계 금액</th><th style="width:90px">내역확인</th><th style="width:90px">발행확인</th><th>비고</th></tr></thead>' +
              '<tbody>' + body + '</tbody>' +
            '</table>' +
          '</div>' +
        '</div>' +
      '</div>'
    );
  }

  function renderClosingPlaceholderTab(title, description) {
    return (
      '<div class="closing-page">' +
        '<div class="closing-card closing-card-wide">' +
          '<div class="closing-title">' + icon("sheet") + ' ' + escapeHtml(title) + '</div>' +
          '<div class="closing-copy">' + escapeHtml(description) + '</div>' +
        '</div>' +
      '</div>'
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
      showAppToast("선택 행은 최대 10줄까지만 가능합니다.", "warning");
      return;
    }

    ensureWorkInfoShape();
    var targetDocDate = defaultWorkDateIsoString();
    workState.info.date = targetDocDate;
    var targetItemDate = isoToMonthDayLabel(targetDocDate);
    var hasDateMismatch = false;

    // 4) 선택 row들의 전체 데이터 복사(원본 ST.rows 수정 X)
    //    -> 거래작업 품목표 컬럼에 매핑
    var items = rows.map(function (r) {
      var row = ST.rows[r] || emptyRow();
      var matchedClientRow = findClientRowByCompany(row.client || "");
      var isPersonalClient = getClientRowSupplierMode(matchedClientRow) === "personal";
      if (!monthDayMatchesIso(row.date, targetDocDate)) {
        hasDateMismatch = true;
      }
      return {
        date: targetItemDate,
        code: isPersonalClient ? "" : row.code,
        name: isPersonalClient ? (row.code || row.name) : row.name,
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
    var firstClientName = "";
    for (var ri = 0; ri < rows.length; ri++) {
      var selectedRow = ST.rows[rows[ri]] || emptyRow();
      if (selectedRow.client && String(selectedRow.client).trim() !== "") {
        firstClientName = String(selectedRow.client).trim();
        break;
      }
    }
    if (firstClientName) {
      var matchedClient = findClientRowByCompany(firstClientName);
      if (matchedClient) {
        applyClientToWorkInfo(matchedClient);
      } else {
        ensureWorkInfoShape();
        workState.info.receiver = Object.assign({}, workState.info.receiver || {}, {
          company: firstClientName,
        });
      }
    }

    if (hasDateMismatch) {
      showAppToast("선택한 매출현황 일자 중 일부가 거래작업 작성일자(오늘+1일)와 다릅니다.", "warning");
    }

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

    if (state.mainTab === "sales" || state.mainTab === "purchase") {
      var isPurchaseTab = state.mainTab === "purchase";
      var currentTabs = isPurchaseTab ? purchaseTabs : salesTabs;
      switchLedgerModule(isPurchaseTab ? "purchase" : "sales");
      if (!currentTabs.some(function (tab) { return tab.key === state.subTab; })) {
        state.subTab = "manage";
      }
      var tabs = '<div class="tabs">';
      currentTabs.forEach(function (t) {
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
      if (state.subTab === "manage") {
        body = renderManageTab();
      } else if (state.subTab === "status") {
        var activeClientType = ST.filter.clientType || "all";
        body =
          '<div class="toolbar-card" style="margin-bottom:12px">' +
          '<label>필터:</label>' +
          '<select id="status-filter-col"><option value="all">전체</option>' +
          salesColumns.map(function (col) { return '<option value="' + col.key + '">' + col.label + '</option>'; }).join("") +
          '</select>' +
          '<input type="text" class="field" id="status-filter-keyword" style="width:180px;height:36px" placeholder="검색어" />' +
          '<button type="button" class="soft-btn filter-apply-btn" id="btn-filter-apply">적용</button>' +
          '<button type="button" class="soft-btn filter-clear-btn" id="btn-filter-clear">해제</button>' +
          '<input type="file" id="status-import-file" accept=".xls,.xlsx,.xml,.csv,.txt,.htm,.html" style="display:none" />' +
          '<button type="button" class="soft-btn" id="btn-status-import">' + (isPurchaseTab ? "매입대장 가져오기" : "매출대장 가져오기") + '</button>' +
          '<div class="toolbar-segment" style="display:inline-flex;gap:6px;margin-left:4px">' +
          '<button type="button" class="soft-btn' + (activeClientType === "all" ? ' active-filter' : '') + '" id="btn-filter-client-all">전체</button>' +
          '<button type="button" class="soft-btn' + (activeClientType === "corporate" ? ' active-filter' : '') + '" id="btn-filter-client-corporate">법인</button>' +
          '<button type="button" class="soft-btn' + (activeClientType === "personal" ? ' active-filter' : '') + '" id="btn-filter-client-personal">개인</button>' +
          '</div>' +
          '<label style="margin-left:12px">정렬:</label>' +
          '<select id="status-sort-col">' +
          salesColumns.map(function (col) { return '<option value="' + col.key + '">' + col.label + '</option>'; }).join("") +
          '</select>' +
          '<button type="button" class="soft-btn" id="btn-sort-asc">오름차순</button>' +
          '<button type="button" class="soft-btn" id="btn-sort-desc">내림차순</button>' +
          '<button type="button" class="soft-btn" id="btn-sort-reset">입력순</button>' +
          (isPurchaseTab ? "" :
          '<button type="button" class="soft-btn" id="btn-send-to-work" style="margin-left:auto">' +
          icon("arrowRight") +
          "거래작업으로 보내기" +
          "</button>") +
          '</div>' +
          '<div id="sales-table-host"></div>';
      } else if (!isPurchaseTab && state.subTab === "work") {
        body = renderWorkTab();
      } else if (!isPurchaseTab && state.subTab === "statement") {
        body = renderStatementTab();
      } else if (state.subTab === "price") {
        body = renderPriceTab();
      } else if (state.subTab === "client") {
        body = renderClientTab();
      } else {
        var lab = currentTabs.find(function (x) {
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

      if (state.subTab === "manage") {
        var wasManageLoaded = ledgerState.loadedFromFirebase;
        ensureLedgerLoadedFromFirebase().then(function () {
          if (!wasManageLoaded && state.subTab === "manage") render();
        });
        var manageStartMonth = document.getElementById("manage-start-month");
        var manageEndMonth = document.getElementById("manage-end-month");
        var manageClient = document.getElementById("manage-client");
        var manageSearchBtn = document.getElementById("btn-manage-search");
        var manageResetBtn = document.getElementById("btn-manage-reset");
        if (manageSearchBtn) {
          manageSearchBtn.addEventListener("click", function () {
            manageState.startMonth = manageStartMonth ? manageStartMonth.value : manageState.startMonth;
            manageState.endMonth = manageEndMonth ? manageEndMonth.value : manageState.endMonth;
            manageState.client = manageClient ? manageClient.value : manageState.client;
            render();
          });
        }
        if (manageStartMonth) {
          manageStartMonth.addEventListener("change", function () {
            shiftManageRangeFromStart(manageStartMonth.value);
            render();
          });
        }
        if (manageEndMonth) {
          manageEndMonth.addEventListener("change", function () {
            manageState.endMonth = normalizeManageMonthValue(manageEndMonth.value, manageState.endMonth || "all");
            if (manageState.endMonth === "all") {
              manageState.startMonth = "all";
            }
            render();
          });
        }
        if (manageResetBtn) {
          manageResetBtn.addEventListener("click", function () {
            manageState.startMonth = "";
            manageState.endMonth = "";
            manageState.client = "";
            render();
          });
        }
      } else if (state.subTab === "status") {
        var filterCol = document.getElementById("status-filter-col");
        var filterKeyword = document.getElementById("status-filter-keyword");
        var sortCol = document.getElementById("status-sort-col");
        if (filterCol) filterCol.value = ST.filter.colKey || "all";
        if (filterKeyword) filterKeyword.value = ST.filter.keyword || "";
        if (sortCol) sortCol.value = ST.filter.colKey && ST.filter.colKey !== "all" ? ST.filter.colKey : "date";
        attachSalesTableWhenNeeded(document.getElementById("sales-table-host"));
        requestAnimationFrame(function () {
          if (state.subTab === "status") focusLastUsedSalesRow();
        });
        ensureLedgerLoadedFromFirebase().then(function () {
          if (state.subTab === "status") {
            refreshGridValues();
            updateSelectionUI();
            applyFilterUI();
            requestAnimationFrame(function () {
              if (state.subTab === "status") focusLastUsedSalesRow();
            });
          }
        });
        var applyBtn = document.getElementById("btn-filter-apply");
        var clearBtn = document.getElementById("btn-filter-clear");
        var filterAllBtn = document.getElementById("btn-filter-client-all");
        var filterCorporateBtn = document.getElementById("btn-filter-client-corporate");
        var filterPersonalBtn = document.getElementById("btn-filter-client-personal");
        var sortAscBtn = document.getElementById("btn-sort-asc");
        var sortDescBtn = document.getElementById("btn-sort-desc");
        var sortResetBtn = document.getElementById("btn-sort-reset");
        var importBtn = document.getElementById("btn-status-import");
        var importFileInput = document.getElementById("status-import-file");
        function setStatusClientTypeFilter(type) {
          ST.filter.clientType = type;
          render();
        }
        if (filterAllBtn) {
          filterAllBtn.addEventListener("click", function () {
            setStatusClientTypeFilter("all");
          });
        }
        if (filterCorporateBtn) {
          filterCorporateBtn.addEventListener("click", function () {
            setStatusClientTypeFilter("corporate");
          });
        }
        if (filterPersonalBtn) {
          filterPersonalBtn.addEventListener("click", function () {
            setStatusClientTypeFilter("personal");
          });
        }
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
            ST.filter.clientType = "all";
            if (filterCol) filterCol.value = "all";
            if (filterKeyword) filterKeyword.value = "";
            render();
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
        if (sendBtn && !isPurchaseTab) {
          sendBtn.addEventListener("click", function () {
            sendSelectedStatusRowsToWork();
          });
        }
        if (importBtn && importFileInput) {
          importBtn.addEventListener("click", function () {
            importFileInput.click();
          });
          importFileInput.addEventListener("change", function () {
            var file = importFileInput.files && importFileInput.files[0] ? importFileInput.files[0] : null;
            if (!file) return;
            var lowerName = String(file.name || "").toLowerCase();
            var isLegacyXls = /\.xls$/.test(lowerName) && !/\.xlsx$/.test(lowerName);
            readFileAsArrayBuffer(file)
              .then(function (buffer) {
                var workbookRows = [];
                try {
                  workbookRows = parseStatusRowsFromWorkbookBuffer(buffer);
                } catch (workbookErr) {
                  console.warn("워크북 파서 실패, 텍스트 파서로 재시도:", workbookErr);
                  if (isLegacyXls) {
                    throw new Error("암호/보호된 .xls 파일은 직접 가져오기 어렵습니다. 엑셀에서 암호 해제 후 .xlsx로 저장해서 가져와주세요.");
                  }
                  workbookRows = [];
                }
                if (workbookRows.length) return workbookRows;
                if (isLegacyXls) {
                  throw new Error(".xls 파일을 해석하지 못했습니다. 엑셀에서 .xlsx로 저장 후 다시 가져와주세요.");
                }
                return readFileAsTextWithEncoding(file, "utf-8").then(function (utfText) {
                  var utfRows = [];
                  try {
                    utfRows = parseStatusRowsFromFileText(utfText);
                  } catch (utfErr) {
                    console.warn("UTF-8 텍스트 파서 실패:", utfErr);
                    utfRows = [];
                  }
                  if (utfRows.length) return utfRows;
                  return readFileAsTextWithEncoding(file, "euc-kr").then(function (eucText) {
                    try {
                      return parseStatusRowsFromFileText(eucText);
                    } catch (eucErr) {
                      console.warn("EUC-KR 텍스트 파서 실패:", eucErr);
                      return [];
                    }
                  });
                });
              })
              .then(function (parsedRows) {
                if (!parsedRows || !parsedRows.length) {
                  showAppToast("가져올 수 있는 현황 행을 찾지 못했어요. 헤더(월일/업체/코드/품명/수량/단가/금액)를 확인해주세요.", "warning");
                  return;
                }
                var importedCount = appendImportedStatusRows(parsedRows);
                refreshGridValues();
                updateSelectionUI();
                applyFilterUI();
                scheduleLedgerDraftSave();
                showAppToast((isPurchaseTab ? "매입현황" : "매출현황") + "으로 " + importedCount + "행을 가져왔어요.", "success");
              })
              .catch(function (err) {
                console.error("현황 파일 가져오기 실패:", err);
                showAppToast(err && err.message ? err.message : "현황 파일 가져오기에 실패했어요.", "warning");
              })
              .finally(function () {
                importFileInput.value = "";
              });
          });
        }
      } else if (!isPurchaseTab && state.subTab === "work") {
        // Firestore에서 복사본을 자동 불러오기
        var wasLoaded = workState.loadedFromFirebase;
        ensureWorkLoadedFromFirebase().then(function () {
          // 처음 로딩인 경우에만 한 번 더 렌더링(값 반영)
          if (!wasLoaded) render();
        });
        var existingReceiverCompany = getWorkInfoValue("receiver.company");
        if (existingReceiverCompany) {
          var autoMatchedClient = findClientRowByCompany(existingReceiverCompany);
          if (autoMatchedClient) {
            applyClientToWorkInfo(autoMatchedClient);
          }
        }
        attachWorkGridWhenNeeded(document.getElementById("work-grid-host"));
        refreshWorkSummaryUi();

        // 거래작업 입력 변경 시 로컬 백업/Firebase 자동 저장
        if (workState.saveTimer) {
          clearTimeout(workState.saveTimer);
          workState.saveTimer = null;
        }
        var supplierModeSelect = document.getElementById("supplier-mode");
        if (supplierModeSelect) {
          supplierModeSelect.addEventListener("change", function () {
            ensureWorkInfoShape().supplierMode = supplierModeSelect.value === "personal" ? "personal" : "corporate";
            scheduleLedgerDraftSave();
            render();
          });
        }
        var supplierResetBtn = document.getElementById("btn-supplier-reset-current");
        if (supplierResetBtn) {
          supplierResetBtn.addEventListener("click", function () {
            var currentMode = getCurrentSupplierMode();
            resetSupplierProfileToDefault(currentMode);
            scheduleLedgerDraftSave();
            render();
          });
        }
        app.querySelectorAll('input[data-work-info]').forEach(function (inp) {
          inp.addEventListener("input", function () {
            var path = inp.getAttribute("data-work-info");
            if (!path) return;
            setWorkInfoValue(path, inp.value);
            if (path === "date") {
              syncWorkItemDatesToDocumentDate();
            }
            if (path === "receiver.company") {
              var matchedClient = findClientRowByCompany(inp.value);
              if (matchedClient) {
                applyClientToWorkInfo(matchedClient);
                scheduleLedgerDraftSave();
                render();
                return;
              }
            }
            scheduleLedgerDraftSave();
          });
        });
      } else if (!isPurchaseTab && state.subTab === "statement") {
        var wasStatementLoaded = workState.loadedFromFirebase;
        ensureWorkLoadedFromFirebase().then(function () {
          if (!wasStatementLoaded && state.subTab === "statement") render();
        });
        var printBtn = document.getElementById("btn-print-statement");
        if (printBtn) {
          printBtn.addEventListener("click", function () {
            window.print();
          });
        }      } else if (state.subTab === "price") {
        var wasPriceLoaded = ledgerState.loadedFromFirebase;
        ensureLedgerLoadedFromFirebase().then(function () {
          if (!wasPriceLoaded && state.subTab === "price") render();
        });
        attachPriceGridWhenNeeded(document.getElementById("price-grid-host"));
        restorePriceSidebarScroll();
        var priceClientInput = document.getElementById("price-active-client");
        var priceOpenBtn = document.getElementById("btn-price-open");
        var priceNewBtn = document.getElementById("btn-price-new");
        var priceRenameBtn = document.getElementById("btn-price-rename");
        var priceDeleteBtn = document.getElementById("btn-price-delete");
        var priceClientListEl = document.querySelector(".price-client-list");
        if (priceClientListEl) {
          priceClientListEl.addEventListener("scroll", function () {
            priceState.sidebarScrollTop = priceClientListEl.scrollTop || 0;
          });
        }
        function applyActivePriceClient() {
          if (!priceClientInput) return;
          capturePriceSidebarScroll();
          priceState.activeClient = String(priceClientInput.value || "").trim();
          scheduleLedgerDraftSave();
          render();
        }
        if (priceOpenBtn) {
          priceOpenBtn.addEventListener("click", applyActivePriceClient);
        }
        if (priceNewBtn) {
          priceNewBtn.addEventListener("click", function () {
            if (!priceClientInput) return;
            priceState.activeClient = String(priceClientInput.value || "").trim();
            scheduleLedgerDraftSave();
            render();
          });
        }
        if (priceRenameBtn) {
          priceRenameBtn.addEventListener("click", function () {
            if (!priceClientInput) return;
            var nextName = String(priceClientInput.value || "").trim();
            var previousName = String(priceState.activeClient || "").trim();
            if (!previousName || !nextName) return;
            if (renamePriceClient(previousName, nextName)) {
              capturePriceSidebarScroll();
              scheduleLedgerDraftSave();
              render();
            }
          });
        }
        if (priceDeleteBtn) {
          priceDeleteBtn.addEventListener("click", function () {
            var currentName = String(priceState.activeClient || "").trim();
            if (!currentName) return;
            if (deletePriceClient(currentName)) {
              capturePriceSidebarScroll();
              scheduleLedgerDraftSave();
              render();
            }
          });
        }
        if (priceClientInput) {
          priceClientInput.addEventListener("change", applyActivePriceClient);
          priceClientInput.addEventListener("input", function () {
            updatePriceClientAutocomplete(priceClientInput);
          });
          priceClientInput.addEventListener("keydown", function (e) {
            if (e.key === "Enter") {
              e.preventDefault();
              hidePriceClientAutocomplete();
              applyActivePriceClient();
            } else if (e.key === "Escape") {
              hidePriceClientAutocomplete();
            }
          });
          priceClientInput.addEventListener("blur", function () {
            setTimeout(hidePriceClientAutocomplete, 120);
          });
        }
        app.querySelectorAll("[data-price-client]").forEach(function (btn) {
          btn.addEventListener("click", function () {
            capturePriceSidebarScroll();
            priceState.activeClient = btn.getAttribute("data-price-client") || "";
            scheduleLedgerDraftSave();
            render();
          });
        });
      } else if (state.subTab === "client") {
        var wasClientLoaded = ledgerState.loadedFromFirebase;
        ensureLedgerLoadedFromFirebase().then(function () {
          if (!wasClientLoaded && state.subTab === "client") render();
        });
        attachClientGridWhenNeeded(document.getElementById("client-grid-host"));
      }
      return;
    }

    if (state.mainTab === "closing") {
      if (!closingTabs.some(function (tab) { return tab.key === state.closingSubTab; })) {
        state.closingSubTab = "attendance";
      }
      var closingTabsHtml = '<div class="tabs">';
      closingTabs.forEach(function (t) {
        closingTabsHtml +=
          '<button type="button" class="sub-tab' +
          (state.closingSubTab === t.key ? " active" : "") +
          '" data-closing-sub="' +
          t.key +
          '">' +
          icon(t.icon) +
          t.label +
          "</button>";
      });
      closingTabsHtml += "</div>";

      var closingBody = "";
      if (state.closingSubTab === "attendance") {
        closingBody = renderClosingAttendanceTab();
      } else if (state.closingSubTab === "salesclose") {
        closingBody = renderClosingSalesCloseTab();
      } else if (state.closingSubTab === "purchaseclose") {
        closingBody = renderClosingPurchaseCloseTab();
      }

      app.innerHTML = top + '<div class="content">' + closingTabsHtml + closingBody + "</div>";
      wireTopBar();
      app.querySelectorAll("[data-closing-sub]").forEach(function (btn) {
        btn.addEventListener("click", function () {
          state.closingSubTab = btn.getAttribute("data-closing-sub");
          render();
        });
      });
      if (state.closingSubTab === "attendance") {
        var wasClosingLoaded = ledgerState.loadedFromFirebase;
        ensureLedgerLoadedFromFirebase().then(function () {
          if (state.mainTab === "closing" && state.closingSubTab === "attendance") {
            if (!wasClosingLoaded) {
              render();
            }
          }
        });
        app.querySelectorAll("[data-closing-attendance-view]").forEach(function (btn) {
          btn.addEventListener("click", function () {
            var nextView = btn.getAttribute("data-closing-attendance-view") === "outsource" ? "outsource" : "employee";
            if (closingState.attendanceView === nextView) return;
            closingState.attendanceView = nextView;
            render();
          });
        });
        app.querySelectorAll("[data-closing-employee-view]").forEach(function (btn) {
          btn.addEventListener("click", function () {
            var nextMode = btn.getAttribute("data-closing-employee-view") === "summary" ? "summary" : "entry";
            if (closingState.employeeViewMode === nextMode) return;
            closingState.employeeViewMode = nextMode;
            scheduleLedgerDraftSave();
            render();
          });
        });
        var closingMonthSelect = document.getElementById("closing-attendance-month");
        if (closingMonthSelect) {
          closingMonthSelect.value = closingState.attendanceMonth || defaultClosingMonthLabel();
          closingMonthSelect.addEventListener("change", function () {
            closingState.attendanceMonth = closingMonthSelect.value || defaultClosingMonthLabel();
            ensureClosingAttendanceState();
            scheduleLedgerDraftSave();
            render();
          });
        }
        var closingSearchInput = document.getElementById("closing-attendance-search");
        if (closingSearchInput) {
          closingSearchInput.value = closingState.attendanceSearch || "";
          function applyClosingSearch() {
            closingState.attendanceSearch = closingSearchInput.value || "";
            render();
          }
          closingSearchInput.addEventListener("change", applyClosingSearch);
          closingSearchInput.addEventListener("keydown", function (e) {
            if (e.key === "Enter") {
              e.preventDefault();
              applyClosingSearch();
            } else if (e.key === "Escape") {
              closingSearchInput.value = "";
              applyClosingSearch();
            }
          });
        }
        var closingAddButton = document.getElementById("closing-attendance-add");
        if (closingAddButton) {
          closingAddButton.addEventListener("click", function () {
            if (closingState.attendanceView === "outsource") {
              addClosingOutsourceGroup("");
            } else {
              addClosingEmployeeGroup("");
            }
            scheduleLedgerDraftSave();
            render();
          });
        }
        var closingOutsourceVendorSelect = document.getElementById("closing-outsource-vendor");
        if (closingOutsourceVendorSelect) {
          closingOutsourceVendorSelect.value = closingState.outsourceVendor || "leaders";
          closingOutsourceVendorSelect.addEventListener("change", function () {
            closingState.outsourceVendor = closingOutsourceVendorSelect.value || "leaders";
            ensureClosingAttendanceState();
            render();
          });
        }
        var closingOutsourceLockDayInput = document.getElementById("closing-outsource-lock-day");
        if (closingOutsourceLockDayInput) {
          closingOutsourceLockDayInput.addEventListener("change", function () {
            var meta = getClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor);
            var parsed = parseCalcNumber(closingOutsourceLockDayInput.value);
            if (parsed == null) {
              meta.lockDay = "";
            } else {
              var normalized = Math.floor(parsed);
              if (normalized < 0) normalized = 0;
              if (normalized > getClosingDaysInMonth()) normalized = getClosingDaysInMonth();
              meta.lockDay = String(normalized);
            }
            setClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor, meta);
            scheduleLedgerDraftSave();
            render();
          });
        }
        var closingOutsourceLockTodayButton = document.getElementById("closing-outsource-lock-today");
        if (closingOutsourceLockTodayButton) {
          closingOutsourceLockTodayButton.addEventListener("click", function () {
            var lockDay = getClosingWarningDayLimit();
            var meta = getClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor);
            meta.lockDay = String(lockDay);
            setClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor, meta);
            scheduleLedgerDraftSave();
            render();
            showAppToast("수정일을 " + lockDay + "일로 고정했어요. 다음 자동채우기부터 이전 날짜는 유지됩니다.", "success");
          });
        }
        var closingFingerprintFileInput = document.getElementById("closing-fingerprint-file");
        if (closingFingerprintFileInput) {
          closingFingerprintFileInput.addEventListener("change", function () {
            closingFingerprintSourceFile = closingFingerprintFileInput.files && closingFingerprintFileInput.files[0]
              ? closingFingerprintFileInput.files[0]
              : null;
            render();
          });
        }
        var closingFingerprintFillButton = document.getElementById("closing-fingerprint-fill");
        if (closingFingerprintFillButton) {
          closingFingerprintFillButton.addEventListener("click", function () {
            if (!closingFingerprintSourceFile) {
              showAppToast("지문 원본 파일을 먼저 선택해주세요.", "warning");
              return;
            }
            readFileAsText(closingFingerprintSourceFile)
              .then(function (text) {
                var parsed = parseFingerprintXml(text);
                if (closingState.attendanceView === "outsource") {
                  var outsourceResult = applyFingerprintDataToOutsourceRows(parsed);
                  scheduleLedgerDraftSave();
                  render();
                  showAppToast(
                    "자동채우기를 완료했어요. 적용 직원 " + outsourceResult.matchedGroups + "명, 확인 필요 칸 " + outsourceResult.warningCells + "개 (수정일 " + outsourceResult.lockDay + "일까지 보호)",
                    outsourceResult.warningCells ? "warning" : "success"
                  );
                } else {
                  var employeeResult = applyFingerprintDataToEmployeeRows(parsed);
                  scheduleLedgerDraftSave();
                  render();
                  showAppToast("정직원 자동채우기를 완료했어요. 적용 직원 " + employeeResult.matchedGroups + "명", "success");
                }
              })
              .catch(function (err) {
                console.error("지문 원본 자동채우기 실패:", err);
                showAppToast(err && err.message ? err.message : "지문 원본 자동채우기에 실패했어요.", "warning");
              });
          });
        }
        if (closingState.attendanceView === "outsource") {
          var outsourceWrap = app.querySelector(".closing-matrix-wrap");
          if (outsourceWrap) {
            bindClosingInputNavigation(outsourceWrap);
            autoFocusClosingFirstInput(outsourceWrap);
            outsourceWrap.querySelectorAll("[data-remove-outsource-group]").forEach(function (btn) {
              btn.addEventListener("click", function () {
                removeClosingOutsourceGroup(Number(btn.getAttribute("data-remove-outsource-group")));
                scheduleLedgerDraftSave();
                render();
              });
            });
            outsourceWrap.addEventListener("input", function (e) {
              var target = e.target;
              if (!(target instanceof HTMLInputElement)) return;
              var rowIndex = Number(target.getAttribute("data-outsource-row"));
              var field = target.getAttribute("data-outsource-field");
              var groupIndex = Number(target.getAttribute("data-outsource-group"));
              var metaField = target.getAttribute("data-outsource-meta");
              var rows = normalizeClosingOutsourceRows(getClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor));
              if (!isNaN(rowIndex) && field) {
                rows[rowIndex] = Object.assign({}, rows[rowIndex] || {}, (function () {
                  var next = {};
                  next[field] = target.value;
                  return next;
                })());
                setClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor, rows);
                scheduleLedgerDraftSave();
                return;
              }
              if (!isNaN(groupIndex) && metaField) {
                var baseIndex = groupIndex;
                var nextMetaValue = target.value;
                if (metaField === "joinDate") {
                  nextMetaValue = normalizeJoinDateText(nextMetaValue);
                  target.value = nextMetaValue;
                }
                rows[baseIndex] = Object.assign({}, rows[baseIndex] || {}, (function () {
                  var next = {};
                  next[metaField] = nextMetaValue;
                  return next;
                })());
                setClosingOutsourceRows(closingState.attendanceMonth, closingState.outsourceVendor, rows);
                scheduleLedgerDraftSave();
              }
            });
            var outsourceRetirementInput = document.getElementById("closing-outsource-retirement");
            if (outsourceRetirementInput) {
              outsourceRetirementInput.addEventListener("input", function () {
                var meta = getClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor);
                meta.retirement = outsourceRetirementInput.value || "";
                setClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor, meta);
                scheduleLedgerDraftSave();
              });
              outsourceRetirementInput.addEventListener("change", function () {
                render();
              });
            }
            var outsourceHourlyWageInput = document.getElementById("closing-outsource-hourly-wage");
            if (outsourceHourlyWageInput) {
              outsourceHourlyWageInput.addEventListener("input", function () {
                var meta = getClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor);
                meta.hourlyWage = outsourceHourlyWageInput.value || "";
                setClosingOutsourceMeta(closingState.attendanceMonth, closingState.outsourceVendor, meta);
                scheduleLedgerDraftSave();
              });
              outsourceHourlyWageInput.addEventListener("change", function () {
                render();
              });
            }
            outsourceWrap.addEventListener("change", function () {
              render();
            });
          }
        } else {
          var employeeWrap = app.querySelector(".closing-matrix-wrap");
          if (employeeWrap) {
            bindClosingInputNavigation(employeeWrap);
            autoFocusClosingFirstInput(employeeWrap);
            employeeWrap.querySelectorAll("[data-remove-employee-group]").forEach(function (btn) {
              btn.addEventListener("click", function () {
                removeClosingEmployeeGroup(Number(btn.getAttribute("data-remove-employee-group")));
                scheduleLedgerDraftSave();
                render();
              });
            });
            employeeWrap.addEventListener("input", function (e) {
              var target = e.target;
              if (!(target instanceof HTMLInputElement)) return;
              var groupIndex = Number(target.getAttribute("data-employee-group"));
              var metaField = target.getAttribute("data-employee-meta");
              var rowIndex = Number(target.getAttribute("data-employee-row"));
              var field = target.getAttribute("data-employee-field");
              var rows = normalizeClosingEmployeeRows(getClosingEmployeeRows(closingState.attendanceMonth));
              if (!isNaN(groupIndex) && metaField) {
                rows[groupIndex] = Object.assign({}, rows[groupIndex] || {}, (function () {
                  var next = {};
                  next[metaField] = target.value;
                  return next;
                })());
                setClosingEmployeeRows(closingState.attendanceMonth, rows);
                scheduleLedgerDraftSave();
                return;
              }
              if (isNaN(rowIndex) || !field) return;
              rows[rowIndex] = Object.assign({}, rows[rowIndex] || {}, (function () {
                var next = {};
                next[field] = target.value;
                return next;
              })());
              rows[rowIndex].count = String(getClosingEmployeeCount(rows[rowIndex], getClosingEmployeePeriodDays()));
              setClosingEmployeeRows(closingState.attendanceMonth, rows);
              scheduleLedgerDraftSave();
            });
          }
        }
      } else if (state.closingSubTab === "salesclose") {
        var closeMonthSelectForSales = document.getElementById("closing-close-month");
        if (closeMonthSelectForSales) {
          closeMonthSelectForSales.value = closingState.attendanceMonth || defaultClosingMonthLabel();
          closeMonthSelectForSales.addEventListener("change", function () {
            closingState.attendanceMonth = closeMonthSelectForSales.value || defaultClosingMonthLabel();
            ensureClosingAttendanceState();
            scheduleLedgerDraftSave();
            render();
          });
        }
        app.querySelectorAll("[data-sales-close-row]").forEach(function (inputEl) {
          inputEl.addEventListener("input", function () {
            var rowIndex = Number(inputEl.getAttribute("data-sales-close-row"));
            var field = inputEl.getAttribute("data-sales-close-field");
            if (isNaN(rowIndex) || !field) return;
            var rows = normalizeClosingSalesCloseRows(getClosingSalesCloseRows(closingState.attendanceMonth));
            rows[rowIndex] = Object.assign({}, rows[rowIndex] || {}, (function () {
              var next = {};
              next[field] = inputEl.value || "";
              return next;
            })());
            setClosingSalesCloseRows(closingState.attendanceMonth, rows);
            scheduleLedgerDraftSave();
          });
        });
      } else if (state.closingSubTab === "purchaseclose") {
        var closeMonthSelectForPurchase = document.getElementById("closing-close-month");
        if (closeMonthSelectForPurchase) {
          closeMonthSelectForPurchase.value = closingState.attendanceMonth || defaultClosingMonthLabel();
          closeMonthSelectForPurchase.addEventListener("change", function () {
            closingState.attendanceMonth = closeMonthSelectForPurchase.value || defaultClosingMonthLabel();
            ensureClosingAttendanceState();
            scheduleLedgerDraftSave();
            render();
          });
        }
        app.querySelectorAll("[data-purchase-close-row]").forEach(function (inputEl) {
          inputEl.addEventListener("input", function () {
            var rowIndex = Number(inputEl.getAttribute("data-purchase-close-row"));
            var field = inputEl.getAttribute("data-purchase-close-field");
            if (isNaN(rowIndex) || !field) return;
            var rows = normalizeClosingPurchaseCloseRows(getClosingPurchaseCloseRows(closingState.attendanceMonth));
            rows[rowIndex] = Object.assign({}, rows[rowIndex] || {}, (function () {
              var next = {};
              next[field] = inputEl.value || "";
              return next;
            })());
            setClosingPurchaseCloseRows(closingState.attendanceMonth, rows);
            scheduleLedgerDraftSave();
          });
        });
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

