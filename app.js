(() => {
  "use strict";

  const SVG_NS = "http://www.w3.org/2000/svg";
  const XHTML_NS = "http://www.w3.org/1999/xhtml";

  const LEVEL_LABELS = ["ประเภทสินค้า", "รุ่น", "ความจุ", "สี"];
  const LEVEL_PROMPTS = [
    "ลากปุ่มกลางไปที่หมวดสินค้า หรือแตะเพื่อเลือก",
    "ลากปุ่มกลางไปเลือกรุ่นที่ต้องการ",
    "ลากปุ่มกลางไปเลือกความจุ",
    "ลากปุ่มกลางไปเลือกสี"
  ];
  const DEFAULT_REMOTE_API_BASE = "https://adizjust.pythonanywhere.com";
  const STOCK_API_BASE = (
    window.STOCK_API_BASE ||
    (isLocalAppServer() ? "" : DEFAULT_REMOTE_API_BASE)
  ).replace(/\/+$/, "");

  const el = (id) => document.getElementById(id);
  const wheelSvg = el("wheel");
  const hubWrap = el("hubWrap");
  const hub = el("hub");
  const hubTitle = el("hubTitle");
  const hubSub = el("hubSub");
  const crumbsBox = el("crumbs");
  const promptEl = el("prompt");
  const backBtn = el("backBtn");
  const resetBtn = el("resetBtn");
  const importBtn = el("importBtn");
  const importInput = el("importInput");
  const toastEl = el("toast");
  const overlay = el("overlay");
  const overlayText = el("overlayText");
  const overlayHelp = el("overlayHelp");
  const overlaySpinner = el("overlaySpinner");
  const wheelWrap = el("wheelWrap");
  const productGrid = el("productGrid");
  const productGridTitle = el("productGridTitle");
  const productGridCount = el("productGridCount");
  const productGridItems = el("productGridItems");
  const searchToggle = el("searchToggle");
  const searchPanel = el("searchPanel");
  const searchInput = el("searchInput");
  const searchResults = el("searchResults");
  let searchMatches = [];
  let searchActive = -1;

  /** @type {Object} nested tree: {category:{model:{storage:{color: {qty, itemCode}}}}} */
  let tree = {};
  /** current selection path, e.g. ["Smartphone","iPhone 15 Pro","128GB"] */
  let path = [];
  /** filename of the data currently loaded, shown for context */
  let dataSourceName = "stock.xlsx";
  /** metadata for the ring currently on screen, used for drag hit-testing */
  let currentSegmentsMeta = [];
  /** active pointer-drag session on the center hub, or null */
  let dragState = null;
  /** svg-space point ({x,y}) the *next* buildRing() should expand out from —
   *  set right before a drag-triggered selection, consumed once and cleared;
   *  null means "use the wheel's resting center" (the default look) */
  let ringOriginPoint = null;
  let toastTimer = null;
  let stockRequestBusy = false;
  let refreshCooldownUntil = 0;
  let importCooldownUntil = 0;
  let lastNavigationAt = -Infinity;

  function updateCommandButtons() {
    const now = Date.now();
    resetBtn.disabled = stockRequestBusy || now < refreshCooldownUntil;
    importBtn.disabled = stockRequestBusy || now < importCooldownUntil;
  }

  function finishCommand(button) {
    stockRequestBusy = false;
    if (button === resetBtn) refreshCooldownUntil = Date.now() + 3000;
    if (button === importBtn) importCooldownUntil = Date.now() + 3000;
    updateCommandButtons();
    setTimeout(updateCommandButtons, 3000);
  }

  function allowNavigation() {
    const now = performance.now();
    if (now - lastNavigationAt < 180) return false;
    lastNavigationAt = now;
    return true;
  }

  // ---------------------------------------------------------------
  // Data loading
  // ---------------------------------------------------------------

  async function loadStock(showFailureOverlay = true) {
    try {
      let res = STOCK_API_BASE
        ? await fetch(apiUrl("/api/stock"), { cache: "no-store" })
        : null;
      if (!res || res.status === 404) {
        res = await fetch(`stock.xlsx?t=${Date.now()}`, { cache: "no-store" });
      }
      if (!res.ok) throw new Error("HTTP " + res.status);
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const rows = mergeWorkbookRows(wb);
      const nextTree = buildTree(rows);
      if (!Object.keys(nextTree).length) throw new Error("ไม่พบข้อมูล stock ที่ใช้ได้");
      tree = nextTree;
      renderSearchResults();
      path = [];
      dataSourceName = "stock.xlsx";
      hideOverlay();
      renderLevel();
      return true;
    } catch (err) {
      if (showFailureOverlay) {
        showError(err);
      } else {
        showToast("โหลดข้อมูลใหม่ไม่สำเร็จ: " + (err && err.message ? err.message : String(err)), "err");
      }
      return false;
    }
  }

  function isLocalAppServer() {
    return ["localhost", "127.0.0.1", ""].includes(window.location.hostname);
  }

  function apiUrl(pathname) {
    return `${STOCK_API_BASE}${pathname}`;
  }

  function connectStockSocket() {
    if (STOCK_API_BASE) return;
    if (!("WebSocket" in window)) return;

    const protocol = window.location.protocol === "https:" ? "wss:" : "ws:";
    const ws = new WebSocket(`${protocol}//${window.location.host}/ws`);

    ws.onmessage = async (event) => {
      try {
        const message = JSON.parse(event.data);
        if (message.type !== "STOCK_UPDATED") return;

        path = [];
        await loadStock();
        showToast("อัปเดตข้อมูล stock.xlsx แล้ว", "ok");
      } catch (err) {
        showToast("โหลดข้อมูลใหม่ไม่สำเร็จ: " + (err && err.message ? err.message : String(err)), "err");
      }
    };

    ws.onclose = () => {
      setTimeout(connectStockSocket, 2500);
    };
  }

  // Reads every sheet in a workbook and merges them into one flat row list.
  // Lets a workbook organize data "one category per sheet" (sheet name is
  // used as the Category when a row doesn't already specify one), while
  // still supporting a single flat sheet (plain CSV) with an explicit
  // Category column.
  function mergeWorkbookRows(wb) {
    const rows = [];
    wb.SheetNames.forEach((sheetName) => {
      if (/^(คำแนะนำ|instructions?|readme|info)$/i.test(sheetName.trim())) return;
      const sheet = wb.Sheets[sheetName];
      const sheetRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      sheetRows.forEach((row) => {
        const r = Object.assign({}, row);
        if (!String(r.Category ?? "").trim()) r.Category = sheetName.trim();
        rows.push(r);
      });
    });
    return rows;
  }

  function buildTree(rows) {
    const t = {};
    for (const row of rows) {
      const cat = String(row.Category ?? "").trim();
      const model = String(row.Model ?? "").trim();
      const storage = String(row.Storage ?? "").trim();
      const color = String(row.Color ?? "").trim();
      const itemCode = getRowValue(row, ["Item Code", "ItemCode", "Code", "SKU"]);
      const qty = Number(row.Qty);
      if (!cat || !model || !storage || !color) continue;
      if (!Number.isFinite(qty) || qty <= 0) continue; // out of stock -> never shown

      if (!t[cat]) t[cat] = {};
      if (!t[cat][model]) t[cat][model] = {};
      if (!t[cat][model][storage]) t[cat][model][storage] = {};
      const existing = t[cat][model][storage][color];
      if (existing) {
        existing.qty += qty;
        existing.itemCode = mergeItemCodes(existing.itemCode, itemCode);
      } else {
        t[cat][model][storage][color] = { qty, itemCode };
      }
    }
    return t;
  }

  function getRowValue(row, names) {
    const wanted = names.map((name) => name.toLowerCase());
    for (const [key, value] of Object.entries(row)) {
      if (wanted.includes(String(key).trim().toLowerCase())) {
        return String(value ?? "").trim();
      }
    }
    return "";
  }

  function mergeItemCodes(a, b) {
    const codes = new Set();
    [a, b].forEach((value) => {
      String(value || "")
        .split(",")
        .map((code) => code.trim())
        .filter(Boolean)
        .forEach((code) => codes.add(code));
    });
    return Array.from(codes).sort().join(", ");
  }

  function resolve(p) {
    let node = tree;
    for (const key of p) {
      if (node == null) return null;
      node = node[key];
    }
    return node;
  }

  function showError(err) {
    overlay.classList.remove("hidden");
    overlaySpinner.style.display = "none";
    overlayText.textContent = "โหลด stock.xlsx ไม่สำเร็จ";
    overlayHelp.innerHTML =
      "เบราว์เซอร์บางตัว (เช่น Chrome) บล็อกการอ่านไฟล์ในเครื่องเมื่อเปิดผ่าน <b>file://</b> โดยตรง<br><br>" +
      "วิธีแก้ที่ง่ายที่สุด: เปิดไฟล์ index.html นี้ด้วย <b>Firefox</b> แทน (รองรับการอ่านไฟล์ในโฟลเดอร์เดียวกันได้ทันที)<br><br>" +
      "รายละเอียดข้อผิดพลาด: " + (err && err.message ? err.message : String(err));
  }

  function hideOverlay() {
    overlay.classList.add("hidden");
  }

  // ---------------------------------------------------------------
  // Import (.csv / .xlsx / .xls)
  // ---------------------------------------------------------------

  importBtn.addEventListener("click", () => {
    if (!stockRequestBusy && Date.now() >= importCooldownUntil) importInput.click();
  });

  importInput.addEventListener("change", async (ev) => {
    const file = ev.target.files && ev.target.files[0];
    importInput.value = ""; // allow re-selecting the same file again later
    if (!file || stockRequestBusy || Date.now() < importCooldownUntil) return;
    stockRequestBusy = true;
    updateCommandButtons();
    try {
      if (isPdfFile(file)) {
        await uploadPdf(file);
        showToast(`นำเข้า PDF "${file.name}" สำเร็จ`, "ok");
        return;
      }

      const synced = await uploadStockFile(file);
      if (synced) {
        if (!await loadStock()) throw new Error("โหลดข้อมูลหลังอัปโหลดไม่สำเร็จ");
        showToast(`ซิงก์ stock.xlsx จาก "${file.name}" สำเร็จ`, "ok");
        return;
      }

      await loadLocalStockFile(file);
      showToast(`นำเข้า "${file.name}" สำเร็จในเครื่องนี้`, "ok");
    } catch (err) {
      showToast("นำเข้าไฟล์ไม่สำเร็จ: " + (err && err.message ? err.message : String(err)), "err");
    } finally {
      finishCommand(importBtn);
    }
  });

  function isPdfFile(file) {
    return file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
  }

  async function uploadPdf(file) {
    const formData = new FormData();
    formData.append("file", file);

    const res = await fetch(apiUrl("/api/convert"), {
      method: "POST",
      body: formData
    });

    if (!res.ok) {
      let message = "HTTP " + res.status;
      try {
        const data = await res.json();
        if (data.detail) message = data.detail;
      } catch (err) {
        // keep the HTTP status as the fallback message
      }
      throw new Error(message);
    }

    const contentType = res.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
      if (!await loadStock()) throw new Error("โหลดข้อมูลหลังอัปโหลดไม่สำเร็จ");
      return;
    }

    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const rows = mergeWorkbookRows(wb);
    const newTree = buildTree(rows);
    if (Object.keys(newTree).length === 0) {
      throw new Error("แปลง PDF สำเร็จ แต่ไม่พบข้อมูล stock ที่ใช้ได้");
    }
    tree = newTree;
    path = [];
    dataSourceName = file.name.replace(/\.pdf$/i, ".xlsx");
    hideOverlay();
    renderLevel();
  }

  async function uploadStockFile(file) {
    const formData = new FormData();
    if (STOCK_API_BASE) {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const rows = mergeWorkbookRows(wb);
      if (!Object.keys(buildTree(rows)).length) {
        throw new Error("ไม่พบข้อมูลที่ใช้ได้ในไฟล์นี้ (ตรวจคอลัมน์ Category/Model/Storage/Color/Qty)");
      }
      const normalized = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(normalized, XLSX.utils.json_to_sheet(rows), "Stock");
      const bytes = XLSX.write(normalized, { bookType: "xlsx", type: "array" });
      formData.append("file", new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "stock.xlsx");
    } else {
      if (file.name.toLowerCase().endsWith(".xls")) return false;
      formData.append("file", file);
    }

    try {
      const res = await fetch(apiUrl("/api/upload-stock"), {
        method: "POST",
        body: formData
      });

      if (!STOCK_API_BASE && (res.status === 404 || res.status === 405)) return false;

      if (!res.ok) {
        let message = "HTTP " + res.status;
        try {
          const data = await res.json();
          if (data.detail) message = data.detail;
        } catch (err) {
          // keep the HTTP status as the fallback message
        }
        throw new Error(message);
      }

      return true;
    } catch (err) {
      if (!STOCK_API_BASE && err instanceof TypeError) return false;
      throw err;
    }
  }

  async function loadLocalStockFile(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const rows = mergeWorkbookRows(wb);
    const newTree = buildTree(rows);
    if (Object.keys(newTree).length === 0) {
      throw new Error("ไม่พบข้อมูลที่ใช้ได้ในไฟล์นี้ (ตรวจคอลัมน์ Category/Model/Storage/Color/Qty)");
    }
    tree = newTree;
    path = [];
    dataSourceName = file.name;
    renderLevel();
  }

  function showToast(msg, type) {
    toastEl.textContent = msg;
    toastEl.className = "toast" + (type === "ok" ? " toast--ok" : type === "err" ? " toast--err" : "");
    // restart the show animation even if a toast is already visible
    void toastEl.offsetWidth;
    toastEl.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toastEl.classList.remove("show"), 3200);
  }

  // ---------------------------------------------------------------
  // Geometry helpers
  // ---------------------------------------------------------------

  const CX = 200, CY = 200, R_OUTER = 190, R_INNER = 96;
  const GAP_DEG = 2.2; // visual gap between segments
  const ARM_RADIUS = R_INNER + (R_OUTER - R_INNER) * 0.5; // drag distance needed to "arm" a segment

  function polar(cx, cy, r, angleDeg) {
    const a = ((angleDeg - 90) * Math.PI) / 180;
    return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
  }

  // Inverse of polar(): given an offset from center (in the same svg units),
  // returns the angle in degrees using the same convention (0 = top, clockwise).
  function angleFromDelta(dx, dy) {
    const a = (Math.atan2(dy, dx) * 180) / Math.PI + 90;
    return ((a % 360) + 360) % 360;
  }

  function donutPath(startAngle, endAngle) {
    const large = endAngle - startAngle <= 180 ? 0 : 1;
    const oStart = polar(CX, CY, R_OUTER, endAngle);
    const oEnd = polar(CX, CY, R_OUTER, startAngle);
    const iStart = polar(CX, CY, R_INNER, endAngle);
    const iEnd = polar(CX, CY, R_INNER, startAngle);
    return [
      "M", oStart.x.toFixed(2), oStart.y.toFixed(2),
      "A", R_OUTER, R_OUTER, 0, large, 0, oEnd.x.toFixed(2), oEnd.y.toFixed(2),
      "L", iEnd.x.toFixed(2), iEnd.y.toFixed(2),
      "A", R_INNER, R_INNER, 0, large, 1, iStart.x.toFixed(2), iStart.y.toFixed(2),
      "Z"
    ].join(" ");
  }

  // bounding box (square) for a foreignObject centered at the segment's mid-radius/mid-angle
  function labelBox(midAngle) {
    const midR = (R_OUTER + R_INNER) / 2;
    const p = polar(CX, CY, midR, midAngle);
    const size = R_OUTER - R_INNER + 30;
    return { x: p.x - size / 2, y: p.y - size / 2, size };
  }

  // ---------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------

  function clearSvg() {
    while (wheelSvg.firstChild) wheelSvg.removeChild(wheelSvg.firstChild);
  }

  function buildRing(items) {
    clearSvg();
    currentSegmentsMeta = [];

    // where this ring's entrance animation bursts outward from: wherever the
    // finger currently is mid-drag, or the wheel's resting center otherwise.
    // Consumed once so the *next* normal render (tap, back, reset) falls
    // back to the default center point.
    const origin = ringOriginPoint || { x: CX, y: CY };
    ringOriginPoint = null;

    const ringBg = document.createElementNS(SVG_NS, "circle");
    ringBg.setAttribute("cx", CX);
    ringBg.setAttribute("cy", CY);
    ringBg.setAttribute("r", (R_OUTER + R_INNER) / 2);
    ringBg.setAttribute("class", "segment-ring");
    wheelSvg.appendChild(ringBg);

    const n = items.length;
    const step = 360 / n;
    const groups = [];

    items.forEach((item, i) => {
      const rangeStart = i * step;
      const rangeEnd = (i + 1) * step;
      const start = rangeStart + GAP_DEG / 2;
      const end = rangeEnd - GAP_DEG / 2;
      const mid = (start + end) / 2;

      const g = document.createElementNS(SVG_NS, "g");
      g.setAttribute("class", "seg-group");
      g.style.transformBox = "view-box";
      g.style.transformOrigin = origin.x.toFixed(2) + "px " + origin.y.toFixed(2) + "px";
      g.style.opacity = "0";

      const segPath = document.createElementNS(SVG_NS, "path");
      segPath.setAttribute("d", donutPath(start, end));
      segPath.setAttribute("class", "segment");
      g.appendChild(segPath);

      const box = labelBox(mid);
      const fo = document.createElementNS(SVG_NS, "foreignObject");
      fo.setAttribute("x", box.x);
      fo.setAttribute("y", box.y);
      fo.setAttribute("width", box.size);
      fo.setAttribute("height", box.size);

      const wrapper = document.createElementNS(XHTML_NS, "div");
      wrapper.setAttribute("class", "seg-label");
      wrapper.innerHTML = renderLabelContent(item);
      fo.appendChild(wrapper);
      g.appendChild(fo);

      const activate = (ev) => {
        ev.preventDefault();
        pickValue(item.value);
      };
      g.addEventListener("click", activate);
      g.style.cursor = "pointer";

      wheelSvg.appendChild(g);
      groups.push(g);

      currentSegmentsMeta.push({ value: item.value, rangeStart, rangeEnd, group: g });
    });

    anime({
      targets: groups,
      scale: [0, 1],
      opacity: [0, 1],
      easing: "easeOutElastic(1, .65)",
      duration: 740,
      delay: anime.stagger(55)
    });

    return groups;
  }

  function renderLabelContent(item) {
    const main = escapeHtml(formatItemLabel(item));
    const sub = item.sub ? `<span class="seg-label-sub">${escapeHtml(item.sub)}</span>` : "";
    return `<span class="seg-label-inner">${main}${sub}</span>`;
  }

  function formatItemLabel(item) {
    return Number.isFinite(item.qty) ? `${item.label} (${item.qty})` : item.label;
  }

  function buildImageSearchUrl(parts) {
    const query = parts.map((part) => String(part || "").trim()).filter(Boolean).join(" ");
    return `https://www.google.com/search?tbm=isch&q=${encodeURIComponent(query)}`;
  }

  function getTotalQty(node) {
    if (!node) return 0;
    if (typeof node === "number") return node;
    if (typeof node === "object" && Number.isFinite(node.qty)) return node.qty;
    if (typeof node !== "object") return 0;
    return Object.values(node).reduce((sum, child) => sum + getTotalQty(child), 0);
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
    }[c]));
  }

  function transitionToRing(items) {
    const oldGroups = Array.from(wheelSvg.querySelectorAll(".seg-group"));
    if (oldGroups.length === 0) {
      buildRing(items);
      return;
    }
    anime({
      targets: oldGroups,
      scale: 0,
      opacity: 0,
      easing: "easeInBack",
      duration: 190,
      complete: () => buildRing(items)
    });
  }

  // ---------------------------------------------------------------
  // Selection (shared by tap-to-select and drag-to-select)
  // ---------------------------------------------------------------

  function pickValue(value) {
    if (!allowNavigation()) return;
    path.push(value);
    renderLevel();
  }

  function modelMatches(query) {
    const words = query.toLocaleLowerCase().trim().split(/\s+/).filter(Boolean);
    if (!words.length) return [];
    const matches = [];
    for (const [category, models] of Object.entries(tree)) {
      for (const [model, variants] of Object.entries(models)) {
        const modelText = model.toLocaleLowerCase();
        const fullText = `${category} ${model}`.toLocaleLowerCase();
        const compactText = fullText.replace(/\s+/g, "");
        const tokens = fullText.split(/[^\p{L}\p{N}]+/u).filter(Boolean);
        if (!words.every((word) => fullText.includes(word) || compactText.includes(word) || tokens.some((token) => isNearWord(word, token)))) continue;
        let score = words.reduce((sum, word) => sum + (modelText.startsWith(word) ? 5 : modelText.includes(word) ? 3 : compactText.includes(word) ? 2 : 0), 0);
        if (modelText === words.join(" ")) score += 20;
        matches.push({ category, model, qty: getTotalQty(variants), score });
      }
    }
    return matches.sort((a, b) => b.score - a.score || a.model.length - b.model.length || a.model.localeCompare(b.model)).slice(0, 6);
  }

  function isNearWord(a, b) {
    if (a.length < 4 || Math.abs(a.length - b.length) > 1) return false;
    let i = 0, j = 0, edits = 0;
    while (i < a.length && j < b.length) {
      if (a[i] === b[j]) { i++; j++; continue; }
      if (++edits > 1) return false;
      if (a.length >= b.length) i++;
      if (b.length >= a.length) j++;
    }
    return edits + (i < a.length || j < b.length ? 1 : 0) <= 1;
  }

  function renderSearchResults() {
    searchMatches = modelMatches(searchInput.value);
    searchActive = -1;
    searchResults.replaceChildren();
    searchInput.setAttribute("aria-expanded", searchMatches.length ? "true" : "false");
    if (!searchInput.value.trim()) return;
    if (!searchMatches.length) {
      const empty = document.createElement("div");
      empty.className = "search-empty";
      empty.textContent = "ไม่พบรุ่นที่ตรงกับคำค้น";
      searchResults.appendChild(empty);
      return;
    }
    searchMatches.forEach((item, index) => {
      const row = document.createElement("button");
      row.type = "button";
      row.className = "search-result";
      row.setAttribute("role", "option");
      row.id = `search-option-${index}`;
      row.setAttribute("aria-selected", "false");
      const name = document.createElement("span");
      name.className = "search-result-name";
      name.textContent = item.model;
      const meta = document.createElement("span");
      meta.className = "search-result-meta";
      meta.textContent = `${item.category} · ${item.qty} เครื่อง`;
      row.append(name, meta);
      row.addEventListener("click", () => selectSearchResult(index));
      searchResults.appendChild(row);
    });
  }

  function setSearchActive(index) {
    searchActive = index;
    Array.from(searchResults.children).forEach((row, i) => {
      row.classList.toggle("search-result--active", i === index);
      row.setAttribute("aria-selected", i === index ? "true" : "false");
    });
    if (index >= 0) searchInput.setAttribute("aria-activedescendant", `search-option-${index}`);
    else searchInput.removeAttribute("aria-activedescendant");
  }

  function closeSearch() {
    searchPanel.classList.remove("search-panel--open");
    searchPanel.setAttribute("aria-hidden", "true");
    searchToggle.setAttribute("aria-expanded", "false");
    searchInput.blur();
  }

  function selectSearchResult(index) {
    const item = searchMatches[index];
    if (!item) return;
    path = [item.category, item.model];
    closeSearch();
    renderLevel();
  }

  searchToggle.addEventListener("click", () => {
    const open = !searchPanel.classList.contains("search-panel--open");
    searchPanel.classList.toggle("search-panel--open", open);
    searchPanel.setAttribute("aria-hidden", open ? "false" : "true");
    searchToggle.setAttribute("aria-expanded", open ? "true" : "false");
    if (open) searchInput.focus();
  });
  searchInput.addEventListener("input", renderSearchResults);
  searchInput.addEventListener("keydown", (event) => {
    if (event.key === "Escape") { closeSearch(); searchToggle.focus(); return; }
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      if (!searchMatches.length) return;
      event.preventDefault();
      const step = event.key === "ArrowDown" ? 1 : -1;
      setSearchActive((searchActive + step + searchMatches.length) % searchMatches.length);
    }
    if (event.key === "Enter" && searchMatches.length) {
      event.preventDefault();
      selectSearchResult(searchActive < 0 ? 0 : searchActive);
    }
  });
  document.addEventListener("pointerdown", (event) => {
    if (!searchPanel.contains(event.target) && !searchToggle.contains(event.target)) closeSearch();
  });

  // ---------------------------------------------------------------
  // Draggable center hub — drag it out to a category/segment; releasing
  // once it reaches that segment's ring counts as selecting it.
  // ---------------------------------------------------------------

  function setArmed(segMeta) {
    currentSegmentsMeta.forEach((s) => {
      if (s.group) s.group.classList.toggle("seg-armed", s === segMeta);
    });
  }

  function onHubPointerDown(ev) {
    if (currentSegmentsMeta.length === 0) return;
    ev.preventDefault();
    try { hub.setPointerCapture(ev.pointerId); } catch (e) { /* ignore */ }
    const rect = wheelWrap.getBoundingClientRect();
    dragState = {
      pointerId: ev.pointerId,
      startX: ev.clientX,
      startY: ev.clientY,
      scale: 400 / rect.width // svg units per screen px (wheel is a square)
    };
    anime.remove(hub);
    hub.classList.add("hub--dragging");
  }

  function onHubPointerMove(ev) {
    if (!dragState || ev.pointerId !== dragState.pointerId) return;

    let dxPx = ev.clientX - dragState.startX;
    let dyPx = ev.clientY - dragState.startY;
    const maxSvgR = R_OUTER - 20;
    const maxPxR = maxSvgR / dragState.scale;
    const distPx = Math.hypot(dxPx, dyPx);
    if (distPx > maxPxR) {
      const k = maxPxR / distPx;
      dxPx *= k;
      dyPx *= k;
    }

    anime.set(hub, { translateX: dxPx, translateY: dyPx });

    const dxSvg = dxPx * dragState.scale;
    const dySvg = dyPx * dragState.scale;
    const radiusSvg = Math.hypot(dxSvg, dySvg);

    if (radiusSvg >= ARM_RADIUS) {
      const angleDeg = angleFromDelta(dxSvg, dySvg);
      const armed = currentSegmentsMeta.find((s) => angleDeg >= s.rangeStart && angleDeg < s.rangeEnd) || null;
      // Reaching a category selects it immediately — the user never has to
      // release the button. commitDragSelection() re-anchors the drag origin
      // to the current finger position afterwards, so the distance below
      // resets to ~0 next frame; that's what stops a steady hold from firing
      // the same selection over and over (no rapid-click retriggering).
      if (armed) commitDragSelection(armed, dxSvg, dySvg, ev);
    }
  }

  function commitDragSelection(segMeta, dxSvg, dySvg, ev) {
    // snap the hub back instantly (no elastic bounce mid-gesture) and clear
    // any hover highlight before the ring swaps out from under it
    anime.remove(hub);
    anime.set(hub, { translateX: 0, translateY: 0 });
    setArmed(null);

    // the next ring should visually burst outward from right where the
    // finger currently is, not from the wheel's static center
    ringOriginPoint = { x: CX + dxSvg, y: CY + dySvg };

    pickValue(segMeta.value);

    const reachedEnd = currentSegmentsMeta.length === 0; // leaf result or dead end
    if (reachedEnd) {
      try { hub.releasePointerCapture(dragState.pointerId); } catch (e) { /* ignore */ }
      hub.classList.remove("hub--dragging");
      dragState = null;
      return;
    }

    // keep the drag alive so the user can carry straight into the next ring
    // without lifting their finger — a fresh, deliberate push past
    // ARM_RADIUS is required for every level, by design.
    dragState.startX = ev.clientX;
    dragState.startY = ev.clientY;
  }

  function releaseHub() {
    hub.classList.remove("hub--dragging");
    setArmed(null);
    dragState = null;
    anime({
      targets: hub,
      translateX: 0,
      translateY: 0,
      duration: 560,
      easing: "easeOutElastic(1, .55)"
    });
  }

  function onHubPointerUp(ev) {
    if (!dragState || ev.pointerId !== dragState.pointerId) return;
    releaseHub();
  }

  function onHubPointerCancel(ev) {
    if (!dragState || ev.pointerId !== dragState.pointerId) return;
    releaseHub();
  }

  hub.addEventListener("pointerdown", onHubPointerDown);
  hub.addEventListener("pointermove", onHubPointerMove);
  hub.addEventListener("pointerup", onHubPointerUp);
  hub.addEventListener("pointercancel", onHubPointerCancel);

  // ---------------------------------------------------------------
  // Level flow
  // ---------------------------------------------------------------

  function renderLevel() {
    removeResultCard();
    backBtn.disabled = path.length === 0;

    const node = resolve(path);

    if (path.length === 4) {
      // leaf: node is a stock object. Clear the wheel so stray taps on the
      // dimmed ring behind the result card can't register a phantom selection.
      hideProductGrid(true);
      clearSvg();
      currentSegmentsMeta = [];
      showResult(node);
      wheelWrap.classList.add("wheel-hidden-behind");
      promptEl.textContent = "";
      updateHub();
      updateCrumbs();
      return;
    }

    wheelWrap.classList.remove("wheel-hidden-behind");

    const keys = node ? Object.keys(node) : [];
    const items = keys.map((k) => ({
      value: k,
      label: k,
      qty: getTotalQty(node[k])
    }));

    if (items.length === 0) {
      promptEl.textContent = "ไม่มีสินค้าคงเหลือในหมวดนี้";
      currentSegmentsMeta = [];
      hideProductGrid(false);
      transitionToRing([]);
    } else {
      promptEl.textContent = LEVEL_PROMPTS[path.length];
      // Radial stays fast for <= 8 choices. When a level has more than 8,
      // temporarily replace the wheel with a readable product grid.
      if (items.length > 8) {
        transitionToProductGrid(items);
      } else {
        transitionToRadial(items);
      }
    }

    updateHub();
    updateCrumbs();
  }

  function transitionToRadial(items) {
    const gridVisible = productGrid && !productGrid.classList.contains("hidden");
    if (gridVisible) {
      animateGridOut(() => {
        hideProductGrid(true);
        wheelWrap.classList.remove("wheel-hidden-behind");
        transitionToRing(items, true);
      });
      return;
    }

    hideProductGrid(true);
    wheelWrap.classList.remove("wheel-hidden-behind");
    transitionToRing(items);
  }

  function transitionToProductGrid(items) {
    const oldGroups = Array.from(wheelSvg.querySelectorAll(".seg-group"));
    const show = () => {
      clearSvg();
      currentSegmentsMeta = [];
      wheelWrap.classList.add("wheel-hidden-for-grid");
      hubWrap.classList.add("hub-hidden-for-grid");
      renderProductGrid(items);
      animateGridIn();
    };

    if (oldGroups.length > 0) {
      anime({
        targets: oldGroups,
        scale: 0.15,
        opacity: 0,
        rotate: 10,
        easing: "easeInBack",
        duration: 180,
        delay: anime.stagger(12, { from: "center" }),
        complete: show
      });
    } else {
      show();
    }
  }

  function renderProductGrid(items) {
    if (!productGrid || !productGridItems) return;

    productGridTitle.textContent = path.length === 0
      ? "รายการสินค้า"
      : truncate(path[path.length - 1], 28);
    productGridCount.textContent = `${items.length} รายการ · ${items.reduce((sum, item) => sum + (item.qty || 0), 0)} เครื่อง`;

    productGridItems.innerHTML = "";
    items.forEach((item, index) => {
      const card = document.createElement("button");
      card.className = "product-grid-card";
      card.type = "button";
      card.style.setProperty("--grid-i", index);
      card.innerHTML = `
        <span class="product-grid-index">${String(index + 1).padStart(2, "0")}</span>
        <span class="product-grid-name">${escapeHtml(formatItemLabel(item))}</span>
        <span class="product-grid-arrow">›</span>
      `;
      card.addEventListener("click", (ev) => {
        ev.preventDefault();
        pickValue(item.value);
      });
      productGridItems.appendChild(card);
    });

    productGrid.classList.remove("hidden");
  }

  function animateGridIn() {
    if (!productGrid || !productGridItems) return;
    const cards = Array.from(productGridItems.children);

    anime.remove(cards);
    anime({
      targets: cards,
      opacity: [0, 1],
      translateY: [18, 0],
      scale: [0.94, 1],
      rotateX: [8, 0],
      easing: "easeOutCubic",
      duration: 260,
      delay: anime.stagger(16, { start: 20 }),
      complete: () => cards.forEach((card) => card.style.removeProperty("transform"))
    });

    anime({
      targets: productGrid,
      opacity: [0, 1],
      translateY: [26, 0],
      scale: [0.985, 1],
      easing: "easeOutExpo",
      duration: 260
    });
  }

  function animateGridOut(done) {
    if (!productGrid || productGrid.classList.contains("hidden")) {
      if (done) done();
      return;
    }

    const cards = Array.from(productGridItems.children);
    anime.remove(cards);
    anime({
      targets: cards,
      opacity: 0,
      translateY: -12,
      scale: 0.96,
      easing: "easeInCubic",
      duration: 130,
      delay: anime.stagger(10, { from: "last" }),
      complete: () => {
        anime({
          targets: productGrid,
          opacity: 0,
          translateY: -22,
          scale: 0.985,
          easing: "easeInCubic",
          duration: 150,
          complete: () => {
            if (done) done();
          }
        });
      }
    });
  }

  function hideProductGrid(immediate = false) {
    if (!productGrid) return;
    if (immediate) {
      anime.remove(productGrid);
      anime.remove(productGridItems ? Array.from(productGridItems.children) : []);
      productGrid.classList.add("hidden");
      productGrid.style.opacity = "";
      productGrid.style.transform = "";
      wheelWrap.classList.remove("wheel-hidden-for-grid");
      hubWrap.classList.remove("hub-hidden-for-grid");
      return;
    }
    productGrid.classList.add("hidden");
  }

  function updateHub() {
    if (path.length === 0) {
      hubTitle.textContent = "STOCK";
      hubSub.textContent = dataSourceName;
    } else if (path.length < 4) {
      hubTitle.textContent = truncate(path[path.length - 1], 12);
      hubSub.textContent = `ขั้นตอน ${path.length}/4`;
    } else {
      hubTitle.textContent = truncate(path[2], 10);
      hubSub.textContent = truncate(path[3], 12);
    }
  }

  function truncate(s, n) {
    s = String(s);
    return s.length > n ? s.slice(0, n - 1) + "…" : s;
  }

  function updateCrumbs() {
    crumbsBox.innerHTML = "";
    if (path.length === 0) {
      const c = document.createElement("span");
      c.className = "crumb crumb--active";
      c.textContent = "เริ่มต้น";
      crumbsBox.appendChild(c);
      return;
    }
    path.forEach((p, i) => {
      if (i > 0) {
        const sep = document.createElement("span");
        sep.className = "crumb-sep";
        sep.textContent = "›";
        crumbsBox.appendChild(sep);
      }
      const c = document.createElement("span");
      c.className = "crumb" + (i === path.length - 1 ? " crumb--active" : "");
      c.textContent = p;
      crumbsBox.appendChild(c);
    });
  }

  // ---------------------------------------------------------------
  // Result card
  // ---------------------------------------------------------------

  function removeResultCard() {
    const existing = document.querySelector(".result-card-pos");
    if (existing) existing.remove();
    hubWrap.style.display = "flex";
  }

  function showResult(stock) {
    const qty = typeof stock === "number" ? stock : Number(stock && stock.qty) || 0;
    const itemCode = typeof stock === "object" && stock ? String(stock.itemCode || "").trim() : "";
    hubWrap.style.display = "none";

    const pos = document.createElement("div");
    pos.className = "result-card-pos";

    const card = document.createElement("div");
    card.className = "result-card";
    card.style.opacity = "0";
    card.style.transformBox = "border-box";
    card.style.transformOrigin = "center";

    const maxScale = 12; // reference scale for the stock bar (purely visual)
    const pct = Math.max(6, Math.min(100, Math.round((qty / maxScale) * 100)));
    const imageSearchUrl = buildImageSearchUrl(path);

    card.innerHTML = `
      <div class="result-actions">
        <button class="result-action-btn result-expand-btn" type="button" aria-pressed="false">ขยาย</button>
        <a class="result-action-btn" href="${escapeHtml(imageSearchUrl)}" target="_blank" rel="noopener noreferrer">ดูรูป</a>
      </div>
      <div class="result-eyebrow">${escapeHtml(path[0])} · ${escapeHtml(path[2])}</div>
      <div class="result-model">${escapeHtml(path[1])}</div>
      <div class="result-variant">${escapeHtml(path[3])}</div>
      ${itemCode ? `<div class="result-item-code">Item Code: ${escapeHtml(itemCode)}</div>` : ""}
      <div class="result-qty-row">
        <div class="result-qty-num">${qty}</div>
        <div class="result-qty-label">เครื่องคงเหลือ</div>
      </div>
      <div class="result-bar-track"><div class="result-bar-fill" id="resultBarFill"></div></div>
      <div class="result-note">แตะ "ย้อนกลับ" เพื่อเลือกสีอื่น หรือ "เริ่มใหม่" เพื่อดูสินค้าอื่น</div>
    `;
    pos.appendChild(card);
    wheelWrap.appendChild(pos);

    const expandBtn = card.querySelector(".result-expand-btn");
    if (expandBtn) {
      expandBtn.addEventListener("click", () => {
        const expanded = card.classList.toggle("result-card--expanded");
        expandBtn.setAttribute("aria-pressed", expanded ? "true" : "false");
        expandBtn.textContent = expanded ? "ย่อ" : "ขยาย";
      });
    }

    anime({ targets: card, opacity: [0, 1], scale: [0.85, 1], easing: "easeOutElastic(1, .7)", duration: 620 });
    requestAnimationFrame(() => {
      const fill = document.getElementById("resultBarFill");
      if (fill) fill.style.transition = "width 0.55s ease";
      if (fill) requestAnimationFrame(() => (fill.style.width = pct + "%"));
    });
  }

  // ---------------------------------------------------------------
  // Controls
  // ---------------------------------------------------------------

  backBtn.addEventListener("click", () => {
    if (path.length === 0 || !allowNavigation()) return;
    path.pop();
    renderLevel();
  });

  resetBtn.addEventListener("click", async () => {
    if (stockRequestBusy || Date.now() < refreshCooldownUntil) return;
    stockRequestBusy = true;
    updateCommandButtons();
    path = [];
    try {
      if (await loadStock(false)) {
        showToast("โหลดข้อมูลสต็อกล่าสุดแล้ว", "ok");
      } else {
        renderLevel();
      }
    } finally {
      finishCommand(resetBtn);
    }
  });

  // kick off
  connectStockSocket();
  stockRequestBusy = true;
  updateCommandButtons();
  loadStock().finally(() => {
    stockRequestBusy = false;
    updateCommandButtons();
  });
})();
