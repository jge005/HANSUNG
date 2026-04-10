(function () {
  function parseClipboardGrid(text) {
    var normalized = String(text || "").replace(/\r/g, "");
    var lines = normalized.split("\n");
    if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
    return lines.map(function (line) {
      return line.split("\t");
    });
  }

  function parseCalcNumber(value) {
    if (value == null) return null;
    var text = String(value).trim().replace(/,/g, "");
    if (!text) return null;
    var num = Number(text);
    return isFinite(num) ? num : null;
  }

  function formatDisplayNumber(value) {
    if (value == null || value === "") return "";
    var num = parseCalcNumber(value);
    if (num == null) return String(value);
    var rounded = Math.round(num * 100) / 100;
    return rounded.toLocaleString("ko-KR");
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
      right: Math.max.apply(null, cs)
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

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }
  function escapeAttr(s) {
    return escapeHtml(s).replace(/"/g, "&quot;");
  }
  function createMiniSheetEngine(options) {
    var engine = {
      host: null,
      scrollEl: null,
      inputEl: null,
      selectEl: null,
      selectedCell: { row: 0, col: 0 },
      selectedKeys: ["0:0"],
      prevSelectedKeys: [],
      prevActiveKey: null,
      dragging: false,
      dragMoved: false,
      suppressNextClick: false,
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
      findQuery: "",
      findMatches: [],
      findIndex: -1,
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

    function getCellEditor(r, c) {
      if (typeof options.getCellEditor === "function") {
        return options.getCellEditor(r, c, getRows()) || null;
      }
      return null;
    }

    function isSelectEditor(r, c) {
      var editor = getCellEditor(r, c);
      return !!(editor && editor.type === "select" && Array.isArray(editor.options));
    }

    function populateSelectOptions(r, c) {
      if (!engine.selectEl) return;
      var editor = getCellEditor(r, c);
      if (!editor || editor.type !== "select" || !Array.isArray(editor.options)) return;
      var currentValue = getRawValue(r, c);
      engine.selectEl.innerHTML = editor.options.map(function (item) {
        var option = typeof item === "string"
          ? { value: item, label: item }
          : { value: item.value, label: item.label || item.value };
        return '<option value="' + escapeAttr(option.value) + '">' + escapeHtml(option.label) + '</option>';
      }).join("");
      engine.selectEl.value = currentValue;
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
        addr.textContent = String.fromCharCode(65 + engine.selectedCell.col) + (engine.selectedCell.row + 1);
      }
      if (range) {
        var bounds = selectionBounds();
        range.textContent =
          String.fromCharCode(65 + bounds.left) + (bounds.top + 1) +
          ":" +
          String.fromCharCode(65 + bounds.right) + (bounds.bottom + 1);
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
          '<div class="st-toolbar">' +
            '<div style="display:flex;align-items:center;gap:10px;min-width:0;flex:1">' +
              '<strong>' + escapeHtml(options.title || "") + '</strong>' +
              '<span>' + escapeHtml(options.subtitle || "") + '</span>' +
            '</div>' +
            '<div id="' + options.idPrefix + '-find-wrap" style="display:' + (engine.findQuery ? 'flex' : 'none') + ';align-items:center;gap:8px;flex-shrink:0">' +
              '<input type="text" id="' + options.idPrefix + '-find-input" class="field" placeholder="찾기 (Ctrl+F)" style="width:180px;height:32px" value="' + escapeAttr(engine.findQuery) + '" />' +
              '<button type="button" class="soft-btn" id="' + options.idPrefix + '-find-prev">이전</button>' +
              '<button type="button" class="soft-btn" id="' + options.idPrefix + '-find-next">다음</button>' +
              '<span class="sub" id="' + options.idPrefix + '-find-status"></span>' +
              '<button type="button" class="soft-btn" id="' + options.idPrefix + '-find-close">닫기</button>' +
            '</div>' +
          '</div>' +
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
          var editor = getCellEditor(r, c);
          html += '<td data-r="' + r + '" data-c="' + c + '" class="' + (active ? 'sel-in sel-active' : '') + '">';
          html += '<div class="st-cell-inner" style="min-height:' + displayHeight + 'px">';
          if (editor && editor.type === "select" && Array.isArray(editor.options)) {
            html +=
              '<select class="st-cell-select" data-inline-select="' + r + ':' + c + '" style="height:' + displayHeight + 'px;line-height:' + displayHeight + 'px;">' +
              editor.options.map(function (item) {
                var option = typeof item === "string"
                  ? { value: item, label: item }
                  : { value: item.value, label: item.label || item.value };
                return '<option value="' + escapeAttr(option.value) + '"' +
                  (String(getRawValue(r, c)) === String(option.value) ? ' selected' : '') +
                  '>' + escapeHtml(option.label) + '</option>';
              }).join("") +
              '</select>';
          } else {
            html +=
              '<div class="st-display" style="' +
                (active ? 'display:none;' : '') +
                'height:' + displayHeight + 'px;line-height:' + displayHeight + 'px;">' +
                escapeHtml(getDisplayValue(r, c)) +
              '</div>';
          }
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
      engine.selectEl = document.createElement("select");
      engine.selectEl.className = "st-input st-select";
      engine.selectEl.style.display = "none";
      moveInputToCell(engine.selectedCell.row, engine.selectedCell.col);

      engine.scrollEl.addEventListener("keydown", onTableKeyDown);
      engine.scrollEl.addEventListener("copy", onCopy);
      engine.scrollEl.addEventListener("paste", onPaste);

      engine.host.querySelector("tbody").addEventListener("mousedown", onCellMouseDown);
      engine.host.querySelector("tbody").addEventListener("click", onCellClick);
      engine.host.querySelector("tbody").addEventListener("dblclick", onCellDblClick);
      engine.host.querySelector("thead").addEventListener("mousedown", onHeadMouseDown);
      engine.host.querySelector("thead").addEventListener("dblclick", onHeadDoubleClick);
      var findInputEl = document.getElementById(options.idPrefix + "-find-input");
      var findPrevBtn = document.getElementById(options.idPrefix + "-find-prev");
      var findNextBtn = document.getElementById(options.idPrefix + "-find-next");
      var findCloseBtn = document.getElementById(options.idPrefix + "-find-close");
      if (findInputEl) {
        findInputEl.addEventListener("input", function () {
          engine.findQuery = findInputEl.value;
          collectFindMatches();
          if (engine.findMatches.length) moveToFindMatch(1);
        });
        findInputEl.addEventListener("keydown", function (e) {
          if (e.key === "Enter") {
            e.preventDefault();
            moveToFindMatch(e.shiftKey ? -1 : 1);
            return;
          }
          if (e.key === "Escape") {
            e.preventDefault();
            closeFindBar();
            focusInput();
          }
        });
      }
      if (findPrevBtn) {
        findPrevBtn.addEventListener("click", function () {
          moveToFindMatch(-1);
        });
      }
      if (findNextBtn) {
        findNextBtn.addEventListener("click", function () {
          moveToFindMatch(1);
        });
      }
      if (findCloseBtn) {
        findCloseBtn.addEventListener("click", function () {
          closeFindBar();
          focusInput();
        });
      }
      engine.host.querySelectorAll("[data-inline-select]").forEach(function (selectEl) {
        selectEl.addEventListener("mousedown", function (e) {
          e.stopPropagation();
        });
        selectEl.addEventListener("focus", function () {
          var parts = String(selectEl.getAttribute("data-inline-select")).split(":");
          var row = Number(parts[0]);
          var col = Number(parts[1]);
          engine.selectedCell = { row: row, col: col };
          engine.selectedKeys = [cellKey(row, col)];
          engine.editMode = false;
          updateSelectionUI();
        });
        selectEl.addEventListener("keydown", function (e) {
          if (e.key === "F2") {
            e.preventDefault();
            return;
          }
          if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "f") {
            e.preventDefault();
            openFindBar(engine.findQuery);
            return;
          }
          if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "c") {
            e.preventDefault();
            copySelectionAsync();
            return;
          }
          if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "x") {
            e.preventDefault();
            cutSelectionAsync();
            return;
          }
          if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "v") {
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
          if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "a") {
            e.preventDefault();
            selectAll();
            return;
          }
          if (e.key === "Delete" && engine.selectedKeys.length > 1) {
            e.preventDefault();
            clearSelectionCells();
            return;
          }
        });
        selectEl.addEventListener("change", function () {
          var parts = String(selectEl.getAttribute("data-inline-select")).split(":");
          var row = Number(parts[0]);
          var col = Number(parts[1]);
          var beforeRows = cloneRowList(getRows());
          var nextRows = cloneRowList(getRows());
          setRawValue(nextRows, row, col, selectEl.value);
          setRows(nextRows, false, beforeRows);
          engine.selectedCell = { row: row, col: col };
          engine.selectedKeys = [cellKey(row, col)];
          engine.editMode = false;
          updateSelectionUI();
        });
      });

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
      engine.selectEl.addEventListener("change", function (e) {
        if (!engine.editMode) {
          engine.editSnapshotRows = cloneRowList(getRows());
          engine.editMode = true;
        }
        setValue(engine.selectedCell.row, engine.selectedCell.col, e.target.value);
        commitActiveCellValue();
        engine.editMode = false;
        syncInputOverlay();
        updateSelectionUI();
        focusInput();
      });
      engine.selectEl.addEventListener("keydown", onInputKeyDown);
      engine.selectEl.addEventListener("blur", function () {
        if (engine.editMode) {
          commitActiveCellValue();
          engine.editMode = false;
          syncInputOverlay();
        }
      });
      syncInputOverlay();
      refreshFindStatus();
      focusInput();
    }

    function moveInputToCell(r, c) {
      if (!engine.host || !engine.inputEl) return;
      var prevTd = engine.inputEl.closest("td") || (engine.selectEl ? engine.selectEl.closest("td") : null);
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
      if (isSelectEditor(r, c)) {
        if (engine.inputEl.parentNode === inner) inner.removeChild(engine.inputEl);
        syncInputOverlay();
        return;
      }
      inner.appendChild(engine.inputEl);
      if (engine.selectEl && engine.selectEl.parentNode === inner) {
        inner.removeChild(engine.selectEl);
      }
      engine.inputEl.value = getRawValue(r, c);
      syncInputOverlay();
    }

    function activateEditorForCell(r, c) {
      engine.selectedCell = { row: r, col: c };
      engine.selectedKeys = [cellKey(r, c)];
      if (!engine.editMode) {
        engine.editSnapshotRows = cloneRowList(getRows());
      }
      engine.editMode = true;
      updateSelectionUI();
      moveInputToCell(r, c);
      syncEditOrigin();
      focusInput();
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
      var inlineSelect = td.querySelector("[data-inline-select]");
      var useSelect = engine.editMode && isSelectEditor(r, c);
      if (inlineSelect) inlineSelect.value = raw;
      if (isSelectEditor(r, c)) {
        engine.selectEl.style.display = "none";
        engine.inputEl.style.display = "none";
        if (ghost) ghost.style.display = "none";
        return;
      }
      if (useSelect && engine.selectEl.parentNode !== inner) {
        inner.appendChild(engine.selectEl);
      } else if (!useSelect && engine.inputEl.parentNode !== inner) {
        inner.appendChild(engine.inputEl);
      }
      engine.inputEl.value = raw;
      if (engine.editMode) {
        if (useSelect) {
          populateSelectOptions(r, c);
          engine.inputEl.style.display = "none";
          engine.selectEl.style.display = "";
        } else {
          engine.selectEl.style.display = "none";
          engine.inputEl.style.display = "";
          engine.inputEl.classList.remove("no-edit");
        }
        if (ghost) {
          ghost.textContent = "";
          ghost.style.display = "none";
        }
      } else {
        engine.selectEl.style.display = "none";
        engine.inputEl.style.display = "";
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
        if (isSelectEditor(engine.selectedCell.row, engine.selectedCell.col)) {
          var td = engine.host && engine.host.querySelector(
            'td[data-r="' + engine.selectedCell.row + '"][data-c="' + engine.selectedCell.col + '"]'
          );
          var inlineSelect = td ? td.querySelector("[data-inline-select]") : null;
          if (inlineSelect) {
            inlineSelect.focus();
            return;
          }
        }
        if (engine.editMode && isSelectEditor(engine.selectedCell.row, engine.selectedCell.col)) {
          populateSelectOptions(engine.selectedCell.row, engine.selectedCell.col);
          engine.selectEl.focus();
          return;
        }
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
        if (k === "f") {
          e.preventDefault();
          openFindBar(engine.findQuery);
          return;
        }
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

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "f") {
        e.preventDefault();
        openFindBar(engine.findQuery);
        return;
      }

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
      if (e.target.closest("[data-inline-select]")) return;
      e.preventDefault();
      commitActiveCellValue();
      var row = Number(td.getAttribute("data-r"));
      var col = Number(td.getAttribute("data-c"));
      if (isSelectEditor(row, col)) {
        activateEditorForCell(row, col);
        return;
      }
      engine.dragging = true;
      engine.dragMoved = false;
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

    function onCellClick(e) {
      var td = e.target.closest("td[data-r]");
      if (!td) return;
      if (e.target.closest("[data-inline-select]")) return;
      if (engine.suppressNextClick) {
        engine.suppressNextClick = false;
        return;
      }
      var row = Number(td.getAttribute("data-r"));
      var col = Number(td.getAttribute("data-c"));
      var key = cellKey(row, col);
      if (
        !engine.dragging &&
        !engine.editMode &&
        !isSelectEditor(row, col) &&
        engine.selectedKeys.length === 1 &&
        engine.selectedKeys[0] === key &&
        engine.selectedCell.row === row &&
        engine.selectedCell.col === col
      ) {
        activateEditorForCell(row, col);
      }
    }

    function refreshFindStatus() {
      if (!engine.host) return;
      var statusEl = document.getElementById(options.idPrefix + "-find-status");
      var inputEl = document.getElementById(options.idPrefix + "-find-input");
      if (inputEl && inputEl.value !== engine.findQuery) {
        inputEl.value = engine.findQuery;
      }
      if (statusEl) {
        if (!engine.findQuery) {
          statusEl.textContent = "";
        } else if (!engine.findMatches.length) {
          statusEl.textContent = "0건";
        } else {
          statusEl.textContent = String(engine.findIndex + 1) + " / " + String(engine.findMatches.length);
        }
      }
    }

    function collectFindMatches() {
      var query = String(engine.findQuery || "").trim().toLowerCase();
      engine.findMatches = [];
      engine.findIndex = -1;
      if (!query) {
        refreshFindStatus();
        return;
      }
      for (var r = 0; r < getRowCount(); r++) {
        for (var c = 0; c < getColCount(); c++) {
          var haystack = String(getDisplayValue(r, c) || getRawValue(r, c) || "").toLowerCase();
          if (haystack.indexOf(query) >= 0) {
            engine.findMatches.push({ row: r, col: c });
          }
        }
      }
      refreshFindStatus();
    }

    function moveToFindMatch(step) {
      if (!engine.findMatches.length) {
        refreshFindStatus();
        return false;
      }
      if (engine.findIndex < 0) {
        engine.findIndex = 0;
      } else {
        engine.findIndex =
          (engine.findIndex + step + engine.findMatches.length) % engine.findMatches.length;
      }
      var match = engine.findMatches[engine.findIndex];
      moveSelection(match.row, match.col);
      refreshFindStatus();
      return true;
    }

    function openFindBar(initialValue) {
      if (!engine.host) return;
      var wrapEl = document.getElementById(options.idPrefix + "-find-wrap");
      var inputEl = document.getElementById(options.idPrefix + "-find-input");
      if (!wrapEl || !inputEl) return;
      wrapEl.style.display = "flex";
      if (typeof initialValue === "string") {
        engine.findQuery = initialValue;
        collectFindMatches();
      } else {
        refreshFindStatus();
      }
      inputEl.value = engine.findQuery;
      inputEl.focus();
      inputEl.select();
    }

    function closeFindBar() {
      if (!engine.host) return;
      var wrapEl = document.getElementById(options.idPrefix + "-find-wrap");
      if (wrapEl) wrapEl.style.display = "none";
    }

    function onCellDblClick(e) {
      var td = e.target.closest("td[data-r]");
      if (!td) return;
      commitActiveCellValue();
      activateEditorForCell(Number(td.getAttribute("data-r")), Number(td.getAttribute("data-c")));
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
        engine.dragMoved = true;
        var rectKeys = keysFromPoints(engine.dragStart, current);
        if (engine.dragMode === "replace") engine.selectedKeys = rectKeys;
        else if (engine.dragMode === "add") engine.selectedKeys = Array.from(new Set(engine.dragBaseKeys.concat(rectKeys)));
        else engine.selectedKeys = engine.dragBaseKeys.filter(function (key) { return rectKeys.indexOf(key) < 0; });
        updateSelectionUI();
        moveInputToCell(current.row, current.col);
      });
      window.addEventListener("mouseup", function () {
        var hadDrag = engine.dragging;
        var hadDragMoved = engine.dragMoved;
        var hadResize = !!engine.resizing;
        engine.dragging = false;
        engine.dragMoved = false;
        engine.resizing = null;
        if (hadDragMoved) {
          engine.suppressNextClick = true;
        }
        if (hadDrag || hadResize) {
          focusInput();
        }
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

  window.createSheetEngine = createMiniSheetEngine;
})();
