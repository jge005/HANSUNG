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
  var SCHEDULE_CHANGE_DIRECTIONS = ["advanced", "delayed", "unknown"];
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

  function detectOrderSourceType(input) {
    if (typeof input === "string") {
      var trimmed = input.trim();
      if (!trimmed) return "text";
      if (trimmed.indexOf("%PDF") === 0) return "pdf";
      return "text";
    }
    if (!input || typeof input !== "object") return "manual";

    var explicit = normalizeEnum(input.source_type, ORDER_SOURCE_TYPES, "");
    if (explicit) return explicit;

    var name = toCleanString(input.name || input.fileName).toLowerCase();
    var mime = toCleanString(input.type || input.mimeType).toLowerCase();
    if (mime.indexOf("sheet") >= 0 || /\.(xlsx|xls|csv|tsv)$/i.test(name)) return "excel";
    if (mime.indexOf("pdf") >= 0 || /\.pdf$/i.test(name)) return "pdf";
    if (mime.indexOf("image") >= 0 || /\.(png|jpg|jpeg|bmp|gif|webp|tif|tiff)$/i.test(name)) return "image";
    if (mime.indexOf("text") >= 0 || /\.(txt|md|log|json)$/i.test(name)) return "text";
    return "manual";
  }

  function normalizeText(rawText) {
    var text = String(rawText || "");
    text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    text = text.replace(/\u00a0/g, " ");
    text = text.replace(/[ \t]+/g, " ");
    text = text.replace(/\n{3,}/g, "\n\n");
    return text.trim();
  }

  function extractRawText(input) {
    return new Promise(function (resolve, reject) {
      try {
        var sourceType = detectOrderSourceType(input);

        if (typeof input === "string") {
          resolve({
            source_type: sourceType,
            raw_text: normalizeText(input),
            meta: { mode: "direct_text" },
          });
          return;
        }

        if (isObject(input) && typeof input.raw_text === "string") {
          resolve({
            source_type: sourceType,
            raw_text: normalizeText(input.raw_text),
            meta: { mode: "raw_text_field" },
          });
          return;
        }

        var file = input && input.file ? input.file : input;
        if (!file || typeof FileReader === "undefined") {
          resolve({
            source_type: sourceType,
            raw_text: "",
            meta: { mode: "no_file_reader" },
          });
          return;
        }

        if (sourceType === "excel" && global.XLSX && typeof file.arrayBuffer === "function") {
          file
            .arrayBuffer()
            .then(function (buffer) {
              var wb = global.XLSX.read(buffer, { type: "array" });
              var lines = [];
              (wb.SheetNames || []).forEach(function (sheetName) {
                var ws = wb.Sheets[sheetName];
                if (!ws) return;
                var rows = global.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
                lines.push("#" + sheetName);
                rows.forEach(function (row) {
                  if (!Array.isArray(row)) return;
                  var line = row.map(function (v) { return String(v == null ? "" : v).trim(); }).join("\t").trim();
                  if (line) lines.push(line);
                });
              });
              resolve({
                source_type: sourceType,
                raw_text: normalizeText(lines.join("\n")),
                meta: { mode: "xlsx_sheet_to_text" },
              });
            })
            .catch(reject);
          return;
        }

        if ((sourceType === "text" || sourceType === "manual" || sourceType === "kakao") && typeof file.text === "function") {
          file
            .text()
            .then(function (text) {
              resolve({
                source_type: sourceType,
                raw_text: normalizeText(text),
                meta: { mode: "file_text" },
              });
            })
            .catch(reject);
          return;
        }

        // PDF/Image는 OCR 엔진 연결 전까지 raw_text 없이 반환
        resolve({
          source_type: sourceType,
          raw_text: "",
          meta: { mode: "no_ocr_stub" },
        });
      } catch (err) {
        reject(err);
      }
    });
  }

  function parseCustomerName(text) {
    var normalized = normalizeText(text);
    var lines = normalized.split("\n");
    var patterns = [
      /(?:거래처|발주처|수신|업체명|고객명|customer)\s*[:：]\s*(.+)$/i,
      /(?:to|buyer)\s*[:：]\s*(.+)$/i,
    ];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      for (var j = 0; j < patterns.length; j++) {
        var m = line.match(patterns[j]);
        if (m && toCleanString(m[1])) {
          return toCleanString(m[1]).replace(/[)\]]+$/, "").trim();
        }
      }
    }
    return "";
  }

  function extractDateCandidates(text) {
    var out = [];
    var normalized = normalizeText(text);
    var lines = normalized.split("\n");
    var dateRegex = /(\d{4}[.\-/년]\s*\d{1,2}[.\-/월]\s*\d{1,2}\s*일?|\d{8})/g;
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      var match;
      while ((match = dateRegex.exec(line))) {
        var raw = toCleanString(match[1]).replace(/년|월/g, "-").replace(/일/g, "").replace(/[.\/]/g, "-").replace(/\s+/g, "");
        var date = normalizeDate(raw);
        if (!date) continue;
        out.push({ date: date, line: line, lineIndex: i });
      }
    }
    return out;
  }

  function parseDates(text) {
    var candidates = extractDateCandidates(text);
    var result = {
      order_date: null,
      requested_ship_date: null,
      due_date: null,
      current_ship_date: null,
      original_ship_date: null,
      candidates: candidates,
    };
    if (!candidates.length) return result;

    var lowerLines = normalizeText(text).toLowerCase().split("\n");
    function pickByLabel(labels) {
      for (var i = 0; i < candidates.length; i++) {
        var c = candidates[i];
        var l = lowerLines[c.lineIndex] || "";
        for (var j = 0; j < labels.length; j++) {
          if (l.indexOf(labels[j]) >= 0) return c.date;
        }
      }
      return null;
    }

    result.order_date =
      pickByLabel(["발주일", "작성일", "order date", "order_date"]) ||
      candidates[0].date;
    result.requested_ship_date =
      pickByLabel(["납기", "요청출고", "출고요청", "희망출고", "ship", "납품요청"]) ||
      (candidates[1] ? candidates[1].date : null);
    result.due_date =
      pickByLabel(["due", "마감", "납기일"]) ||
      result.requested_ship_date;
    result.current_ship_date = result.requested_ship_date;
    result.original_ship_date = result.requested_ship_date;
    return result;
  }

  function parseUrgency(text) {
    var t = normalizeText(text).toLowerCase();
    return /긴급|급함|당일|최우선|urgent|asap|rush/.test(t);
  }

  function parseItems(text) {
    var normalized = normalizeText(text);
    var lines = normalized.split("\n");
    var items = [];

    function tryMakeItemFromParts(parts) {
      if (!parts || !parts.length) return null;
      var cleaned = parts.map(function (p) { return toCleanString(p); }).filter(Boolean);
      if (cleaned.length < 2) return null;

      var qty = null;
      var unit = "";
      var unitPrice = null;
      var qtyMatch = cleaned.join(" ").match(/(?:수량|qty)?\s*[:：]?\s*([0-9][0-9,\.]*)\s*(EA|PCS|SET|KG|M|ROLL|BOX|개|박스|롤|kg|m)?/i);
      if (qtyMatch) {
        qty = normalizeNumber(qtyMatch[1], null);
        unit = toCleanString(qtyMatch[2] || "");
      }
      var priceMatch = cleaned.join(" ").match(/(?:단가|price|unit\s*price)\s*[:：]?\s*([0-9][0-9,\.]*)/i);
      if (priceMatch) unitPrice = normalizeNumber(priceMatch[1], null);

      var codeLike = cleaned.find(function (p) { return /^[A-Z]{1,4}\d{2,}[-A-Z0-9]*$/i.test(p); }) || "";
      var productName = cleaned[0];
      if (productName.length <= 2 && cleaned[1]) productName = cleaned[1];

      if (!productName) return null;
      if (qty == null && !/\t|,|\|/.test(parts.join(" "))) return null;

      return {
        product_name: codeLike ? productName + " " + codeLike : productName,
        spec: cleaned.length >= 3 ? cleaned[1] : "",
        quantity: qty == null ? 0 : qty,
        unit: unit,
        unit_price: unitPrice,
        remarks: "",
      };
    }

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      var lower = line.toLowerCase();
      if (!line.trim()) continue;
      if (/발주일|작성일|거래처|납기|긴급|customer|order date/.test(lower)) continue;

      var parts;
      if (line.indexOf("\t") >= 0) parts = line.split("\t");
      else if (line.indexOf("|") >= 0) parts = line.split("|");
      else if (line.indexOf(",") >= 0) parts = line.split(",");
      else parts = line.split(/\s{2,}/);

      var item = tryMakeItemFromParts(parts);
      if (item) items.push(item);
    }

    if (!items.length) {
      // fallback: 품목명 + 수량 패턴
      var fallbackRegex = /([A-Za-z0-9가-힣\-\(\)\/\.\s]{3,}?)\s+([0-9][0-9,\.]*)\s*(EA|PCS|SET|KG|M|ROLL|BOX|개|박스|롤)?/g;
      var m;
      while ((m = fallbackRegex.exec(normalized))) {
        items.push({
          product_name: toCleanString(m[1]),
          spec: "",
          quantity: normalizeNumber(m[2], 0),
          unit: toCleanString(m[3] || ""),
          unit_price: null,
          remarks: "",
        });
        if (items.length >= 200) break;
      }
    }

    return items;
  }

  function parseScheduleChange(text, currentShipDate) {
    var normalized = normalizeText(text);
    if (!/변경|연기|앞당김|납기변경|reschedule|delay|advance/i.test(normalized)) {
      return null;
    }
    var dates = extractDateCandidates(normalized);
    if (!dates.length) return null;
    var previous = currentShipDate || (dates[0] ? dates[0].date : null);
    var newer = dates[1] ? dates[1].date : dates[0].date;
    var dir = computeDirection(previous, newer);
    var requester = /요청|customer|고객|거래처/i.test(normalized) ? "customer" : "unknown";
    if (/내부|생산|자재|internal/i.test(normalized)) requester = "internal";
    return normalizeScheduleChange(
      {
        previous_ship_date: previous,
        new_ship_date: newer,
        change_direction: dir === "unknown" ? "delayed" : dir,
        change_requested_by: requester,
        change_reason: "텍스트에서 일정 변경 키워드 감지",
        note: normalized.slice(0, 300),
      },
      null
    );
  }

  function buildNormalizedOrder(options) {
    var opts = isObject(options) ? options : {};
    var rawText = normalizeText(opts.raw_text || "");
    var sourceType = normalizeEnum(opts.source_type, ORDER_SOURCE_TYPES, detectOrderSourceType(opts.file || rawText || "text"));
    var inputMethod = normalizeEnum(opts.input_method, ORDER_INPUT_METHODS, "auto_upload");
    var receivedChannel = normalizeEnum(opts.received_channel, ORDER_RECEIVED_CHANNELS, "file_upload");

    var customerName = parseCustomerName(rawText);
    var dates = parseDates(rawText);
    var items = parseItems(rawText);
    var urgent = parseUrgency(rawText);
    var scheduleChange = parseScheduleChange(rawText, dates.current_ship_date);

    var unresolved = [];
    var confidence = {
      overall: 0,
      customer_name: customerName ? 0.9 : 0.2,
      dates: dates.order_date || dates.requested_ship_date ? 0.75 : 0.2,
      items: items.length ? 0.8 : 0.2,
      urgency: urgent ? 0.9 : 0.6,
    };

    if (!customerName) unresolved.push("customer_name");
    if (!dates.order_date) unresolved.push("order_date");
    if (!dates.requested_ship_date) unresolved.push("requested_ship_date");
    if (!items.length) unresolved.push("order_items");
    if (!rawText) unresolved.push("raw_text");

    confidence.overall =
      (confidence.customer_name + confidence.dates + confidence.items + confidence.urgency) / 4;

    var baseOrder = normalizeOrder({
      source_type: sourceType,
      input_method: inputMethod,
      received_channel: receivedChannel,
      received_by: toCleanString(opts.received_by),
      customer_name: customerName,
      order_date: dates.order_date,
      requested_ship_date: dates.requested_ship_date,
      current_ship_date: dates.current_ship_date,
      original_ship_date: dates.original_ship_date,
      due_date: dates.due_date,
      is_urgent: urgent,
      status: unresolved.length ? "review_required" : "received",
      notes: toCleanString(opts.notes),
      raw_text: rawText,
      unresolved_fields: unresolved,
      confidence: confidence.overall,
      order_items: items.map(function (row) { return normalizeOrderItem(row, null); }),
      schedule_changes: scheduleChange ? [scheduleChange] : [],
    });

    baseOrder.order_items = baseOrder.order_items.map(function (it) {
      return normalizeOrderItem(it, baseOrder.id);
    });
    baseOrder.schedule_changes = baseOrder.schedule_changes.map(function (sc) {
      return normalizeScheduleChange(sc, baseOrder.id);
    });

    return {
      order: baseOrder,
      items: baseOrder.order_items.slice(),
      confidence: confidence,
      unresolved_fields: unresolved.slice(),
    };
  }

  function runExtractionFromInput(input, context) {
    var ctx = isObject(context) ? context : {};
    return extractRawText(input).then(function (rawResult) {
      return buildNormalizedOrder({
        source_type: rawResult.source_type,
        input_method: ctx.input_method || "auto_upload",
        received_channel: ctx.received_channel || (rawResult.source_type === "phone" ? "phone" : "file_upload"),
        received_by: ctx.received_by || "",
        notes: ctx.notes || "",
        raw_text: rawResult.raw_text || "",
        file: input && input.file ? input.file : input,
      });
    });
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
    detectOrderSourceType: detectOrderSourceType,
    extractRawText: extractRawText,
    normalizeText: normalizeText,
    parseCustomerName: parseCustomerName,
    parseDates: parseDates,
    parseItems: parseItems,
    parseUrgency: parseUrgency,
    parseScheduleChange: parseScheduleChange,
    buildNormalizedOrder: buildNormalizedOrder,
    runExtractionFromInput: runExtractionFromInput,
  };
})(window);
