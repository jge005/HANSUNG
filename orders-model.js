(function (global) {
  "use strict";

  var ORDER_SOURCE_TYPES = ["excel", "pdf", "image", "text", "manual", "phone", "kakao"];
  var ORDER_INPUT_METHODS = ["auto_upload", "manual_entry"];
  var ORDER_RECEIVED_CHANNELS = ["email", "phone", "kakao", "direct", "file_upload"];
  var ORDER_STATUSES = [
    "received",
    "review_required",
    "confirmed",
    "scheduled",
    "in_progress",
    "completed",
    "cancelled",
  ];
  var SCHEDULE_CHANGE_DIRECTIONS = ["advanced", "delayed"];
  var SCHEDULE_CHANGE_REQUESTERS = ["customer", "internal", "unknown"];

  function nowIso() {
    return new Date().toISOString();
  }

  function createId(prefix) {
    var head = prefix || "id";
    try {
      if (global.crypto && typeof global.crypto.randomUUID === "function") {
        return head + "_" + global.crypto.randomUUID();
      }
    } catch (err) {}
    return head + "_" + Date.now() + "_" + Math.floor(Math.random() * 1000000);
  }

  function isObject(value) {
    return value != null && typeof value === "object" && !Array.isArray(value);
  }

  function toCleanString(value) {
    if (value == null) return "";
    return String(value).trim();
  }

  function toNullableString(value) {
    var text = toCleanString(value);
    return text || null;
  }

  function normalizeEnum(value, allowed, fallback) {
    var text = toCleanString(value);
    if (allowed.indexOf(text) >= 0) return text;
    return fallback;
  }

  function normalizeBoolean(value) {
    if (value === true) return true;
    if (value === false) return false;
    var text = toCleanString(value).toLowerCase();
    return text === "true" || text === "1" || text === "y" || text === "yes";
  }

  function normalizeNumber(value, fallback) {
    if (value == null || value === "") return fallback;
    var text = String(value).replace(/,/g, "").trim();
    if (!text) return fallback;
    var num = Number(text);
    return isFinite(num) ? num : fallback;
  }

  function normalizeConfidence(value) {
    var num = normalizeNumber(value, 0);
    if (!isFinite(num)) return 0;
    if (num < 0) return 0;
    if (num > 1) return 1;
    return num;
  }

  function normalizeDate(value) {
    var text = toCleanString(value);
    if (!text) return null;
    var m = text.match(/^(\d{4})[-/.]?(\d{1,2})[-/.]?(\d{1,2})$/);
    if (m) {
      var year = Number(m[1]);
      var month = Number(m[2]);
      var day = Number(m[3]);
      if (year >= 1900 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        var mm = month < 10 ? "0" + month : String(month);
        var dd = day < 10 ? "0" + day : String(day);
        return year + "-" + mm + "-" + dd;
      }
    }
    var date = new Date(text);
    if (!isNaN(date.getTime())) {
      var y = date.getFullYear();
      var mo = date.getMonth() + 1;
      var d = date.getDate();
      return (
        y +
        "-" +
        (mo < 10 ? "0" + mo : String(mo)) +
        "-" +
        (d < 10 ? "0" + d : String(d))
      );
    }
    return null;
  }

  function normalizeUnresolvedFields(value) {
    if (!Array.isArray(value)) return [];
    var out = [];
    for (var i = 0; i < value.length; i++) {
      var text = toCleanString(value[i]);
      if (!text) continue;
      if (out.indexOf(text) >= 0) continue;
      out.push(text);
    }
    return out;
  }

  function createEmptyOrderItem(orderId) {
    return {
      id: createId("oi"),
      order_id: orderId || null,
      product_name: "",
      spec: "",
      quantity: 0,
      unit: "",
      unit_price: null,
      remarks: "",
    };
  }

  function normalizeOrderItem(input, orderId) {
    var src = isObject(input) ? input : {};
    var item = createEmptyOrderItem(orderId || src.order_id || null);
    item.id = toCleanString(src.id) || item.id;
    item.order_id = toNullableString(src.order_id) || item.order_id;
    item.product_name = toCleanString(src.product_name);
    item.spec = toCleanString(src.spec);
    item.quantity = normalizeNumber(src.quantity, 0);
    item.unit = toCleanString(src.unit);
    item.unit_price = normalizeNumber(src.unit_price, null);
    item.remarks = toCleanString(src.remarks);
    return item;
  }

  function createEmptyScheduleChange(orderId) {
    return {
      id: createId("sc"),
      order_id: orderId || null,
      previous_ship_date: null,
      new_ship_date: null,
      change_direction: "unknown",
      change_requested_by: "unknown",
      change_reason: "",
      note: "",
      changed_at: nowIso(),
      changed_by: "",
    };
  }

  function computeDirection(previousShipDate, newShipDate) {
    if (!previousShipDate || !newShipDate) return "unknown";
    var prev = new Date(previousShipDate).getTime();
    var next = new Date(newShipDate).getTime();
    if (!isFinite(prev) || !isFinite(next)) return "unknown";
    if (next < prev) return "advanced";
    if (next > prev) return "delayed";
    return "unknown";
  }

  function normalizeScheduleChange(input, orderId) {
    var src = isObject(input) ? input : {};
    var row = createEmptyScheduleChange(orderId || src.order_id || null);
    row.id = toCleanString(src.id) || row.id;
    row.order_id = toNullableString(src.order_id) || row.order_id;
    row.previous_ship_date = normalizeDate(src.previous_ship_date);
    row.new_ship_date = normalizeDate(src.new_ship_date);
    row.change_direction = normalizeEnum(
      src.change_direction,
      SCHEDULE_CHANGE_DIRECTIONS,
      computeDirection(row.previous_ship_date, row.new_ship_date)
    );
    row.change_requested_by = normalizeEnum(
      src.change_requested_by,
      SCHEDULE_CHANGE_REQUESTERS,
      "unknown"
    );
    row.change_reason = toCleanString(src.change_reason);
    row.note = toCleanString(src.note);
    row.changed_at = toCleanString(src.changed_at) || nowIso();
    row.changed_by = toCleanString(src.changed_by);
    return row;
  }

  function createEmptyOrder() {
    var now = nowIso();
    return {
      id: createId("ord"),
      source_type: "manual",
      input_method: "manual_entry",
      received_channel: "direct",
      received_by: "",
      customer_name: "",
      order_date: null,
      requested_ship_date: null,
      current_ship_date: null,
      original_ship_date: null,
      due_date: null,
      is_urgent: false,
      status: "received",
      notes: "",
      raw_text: "",
      unresolved_fields: [],
      confidence: 1,
      created_at: now,
      updated_at: now,
      order_items: [],
      schedule_changes: [],
    };
  }

  function normalizeOrder(input) {
    var src = isObject(input) ? input : {};
    var base = createEmptyOrder();
    base.id = toCleanString(src.id) || base.id;
    base.source_type = normalizeEnum(src.source_type, ORDER_SOURCE_TYPES, base.source_type);
    base.input_method = normalizeEnum(src.input_method, ORDER_INPUT_METHODS, base.input_method);
    base.received_channel = normalizeEnum(
      src.received_channel,
      ORDER_RECEIVED_CHANNELS,
      base.received_channel
    );
    base.received_by = toCleanString(src.received_by);
    base.customer_name = toCleanString(src.customer_name);
    base.order_date = normalizeDate(src.order_date);
    base.requested_ship_date = normalizeDate(src.requested_ship_date);
    base.current_ship_date = normalizeDate(src.current_ship_date);
    base.original_ship_date = normalizeDate(src.original_ship_date);
    base.due_date = normalizeDate(src.due_date);
    base.is_urgent = normalizeBoolean(src.is_urgent);
    base.status = normalizeEnum(src.status, ORDER_STATUSES, base.status);
    base.notes = toCleanString(src.notes);
    base.raw_text = toCleanString(src.raw_text);
    base.unresolved_fields = normalizeUnresolvedFields(src.unresolved_fields);
    base.confidence = normalizeConfidence(src.confidence);
    base.created_at = toCleanString(src.created_at) || base.created_at;
    base.updated_at = toCleanString(src.updated_at) || nowIso();

    var itemSource = Array.isArray(src.order_items) ? src.order_items : [];
    base.order_items = itemSource.map(function (row) {
      return normalizeOrderItem(row, base.id);
    });

    var changeSource = Array.isArray(src.schedule_changes) ? src.schedule_changes : [];
    base.schedule_changes = changeSource.map(function (row) {
      return normalizeScheduleChange(row, base.id);
    });

    if (!base.original_ship_date) {
      base.original_ship_date = base.current_ship_date || base.requested_ship_date || null;
    }
    if (!base.current_ship_date) {
      base.current_ship_date = base.requested_ship_date || base.original_ship_date || null;
    }
    return base;
  }

  function createOrderFromManualEntry(input) {
    var source = isObject(input) ? input : {};
    source.source_type = source.source_type || "manual";
    source.input_method = "manual_entry";
    source.confidence = source.confidence == null ? 1 : source.confidence;
    return normalizeOrder(source);
  }

  function createOrderFromAutoUpload(input) {
    var source = isObject(input) ? input : {};
    source.input_method = "auto_upload";
    if (!source.source_type) source.source_type = "text";
    if (source.confidence == null) source.confidence = 0.6;
    return normalizeOrder(source);
  }

  function addOrderItem(order, itemInput) {
    var normalized = normalizeOrder(order);
    var row = normalizeOrderItem(itemInput, normalized.id);
    normalized.order_items.push(row);
    normalized.updated_at = nowIso();
    return normalized;
  }

  function addScheduleChange(order, changeInput) {
    var normalized = normalizeOrder(order);
    var change = normalizeScheduleChange(changeInput, normalized.id);
    if (!change.previous_ship_date) {
      change.previous_ship_date = normalized.current_ship_date || normalized.requested_ship_date || null;
      change.change_direction = normalizeEnum(
        change.change_direction,
        SCHEDULE_CHANGE_DIRECTIONS,
        computeDirection(change.previous_ship_date, change.new_ship_date)
      );
    }
    normalized.schedule_changes.push(change);
    if (change.new_ship_date) {
      if (!normalized.original_ship_date) {
        normalized.original_ship_date = normalized.current_ship_date || change.previous_ship_date || change.new_ship_date;
      }
      normalized.current_ship_date = change.new_ship_date;
    }
    normalized.updated_at = nowIso();
    return normalized;
  }

  function mergeExtractionResult(order, extracted) {
    var current = normalizeOrder(order);
    var patch = isObject(extracted) ? extracted : {};
    var merged = normalizeOrder(Object.assign({}, current, patch, {
      id: current.id,
      created_at: current.created_at,
      updated_at: nowIso(),
      input_method: current.input_method === "manual_entry" ? current.input_method : "auto_upload",
    }));
    if (Array.isArray(patch.order_items)) {
      merged.order_items = patch.order_items.map(function (row) {
        return normalizeOrderItem(row, merged.id);
      });
    }
    return merged;
  }

  function toFirestoreWriteBundle(order) {
    var normalized = normalizeOrder(order);
    var orderDoc = Object.assign({}, normalized);
    delete orderDoc.order_items;
    delete orderDoc.schedule_changes;
    return {
      order_doc: orderDoc,
      order_items: normalized.order_items.slice(),
      schedule_changes: normalized.schedule_changes.slice(),
    };
  }

  function sampleOrderStructure() {
    var order = createOrderFromManualEntry({
      received_by: "담당자명",
      customer_name: "거래처명",
      order_date: "2026-04-17",
      requested_ship_date: "2026-04-20",
      current_ship_date: "2026-04-20",
      original_ship_date: "2026-04-20",
      due_date: "2026-04-22",
      status: "received",
      notes: "",
      unresolved_fields: [],
      confidence: 1,
    });
    order = addOrderItem(order, {
      product_name: "FILM A",
      spec: "0.38T",
      quantity: 100,
      unit: "EA",
      unit_price: 1200,
      remarks: "",
    });
    return order;
  }

  global.OrderModel = {
    ORDER_SOURCE_TYPES: ORDER_SOURCE_TYPES.slice(),
    ORDER_INPUT_METHODS: ORDER_INPUT_METHODS.slice(),
    ORDER_RECEIVED_CHANNELS: ORDER_RECEIVED_CHANNELS.slice(),
    ORDER_STATUSES: ORDER_STATUSES.slice(),
    SCHEDULE_CHANGE_DIRECTIONS: SCHEDULE_CHANGE_DIRECTIONS.slice(),
    SCHEDULE_CHANGE_REQUESTERS: SCHEDULE_CHANGE_REQUESTERS.slice(),

    createId: createId,
    normalizeDate: normalizeDate,
    normalizeOrderItem: normalizeOrderItem,
    normalizeScheduleChange: normalizeScheduleChange,
    normalizeOrder: normalizeOrder,
    createEmptyOrder: createEmptyOrder,
    createOrderFromManualEntry: createOrderFromManualEntry,
    createOrderFromAutoUpload: createOrderFromAutoUpload,
    addOrderItem: addOrderItem,
    addScheduleChange: addScheduleChange,
    mergeExtractionResult: mergeExtractionResult,
    toFirestoreWriteBundle: toFirestoreWriteBundle,
    sampleOrderStructure: sampleOrderStructure,
  };
})(window);

