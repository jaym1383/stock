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
  const syncBtn = el("syncBtn");
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
  const stage = document.querySelector(".stage");
  const importControl = el("importControl");
  const importCollapse = el("importCollapse");
  const codeOverlay = el("codeOverlay");
  const expandedItemCode = el("expandedItemCode");
  const selectionOpen = el("selectionOpen");
  const selectionExit = el("selectionExit");
  const selectionCount = el("selectionCount");
  const selectionModal = el("selectionModal");
  const selectionList = el("selectionList");
  const selectionText = el("selectionText");
  const includeQuantity = el("includeQuantity");
  const selectionCopy = el("selectionCopy");
  const selectedPaths = new Map();
  const excludedLeaves = new Set();
  let selectionMode = false;
  let searchMatches = [];
  let searchActive = -1;
  let pendingOverviewState = null;
  let codeReturnFocus = null;
  let importExpanded = false;
  let pickerOpen = false;

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
  let hubReturning = false;
  let suppressTileClickUntil = 0;
  /** svg-space point ({x,y}) the *next* buildRing() should expand out from —
   *  set right before a drag-triggered selection, consumed once and cleared;
   *  null means "use the wheel's resting center" (the default look) */
  let ringOriginPoint = null;
  let toastTimer = null;
  let stockRequestBusy = false;
  let refreshCooldownUntil = 0;
  let importCooldownUntil = 0;
  let lastNavigationAt = -Infinity;
  let offlineMode = false;
  let lastSnapshot = null;
  const offlineNotice = el("offlineNotice");
  function setOffline(value) {
    offlineMode = value;
    offlineNotice.hidden = !value;
    offlineNotice.open = false;
    const hasData = Object.keys(tree).length > 0;
    el("offlineDescription").textContent = hasData
      ? "กำลังใช้ข้อมูลล่าสุดในเครื่อง ไม่สามารถอัปโหลดได้ชั่วคราว"
      : "ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้ และยังไม่มีข้อมูลสำรองในเครื่อง";
    el("offlineTimestamp").textContent = lastSnapshot
      ? "ข้อมูลล่าสุด: " + new Date(lastSnapshot.savedAt).toLocaleString("th-TH") : "";
    el("offlineEmpty").hidden = hasData;
    wheelWrap.hidden = !hasData;
    promptEl.hidden = !hasData;
    searchToggle.disabled = !hasData;
    if (value) setImportExpanded(false);
    updateCommandButtons();
  }
  document.addEventListener("pointerdown", event => { if (!offlineNotice.contains(event.target)) offlineNotice.open = false; });
  document.addEventListener("keydown", event => { if (event.key === "Escape") offlineNotice.open = false; });
  window.addEventListener("offline", () => setOffline(true));

  function updateCommandButtons() {
    const now = Date.now();
    resetBtn.disabled = stockRequestBusy;
    syncBtn.disabled = stockRequestBusy || now < refreshCooldownUntil;
    syncBtn.setAttribute("aria-busy", String(stockRequestBusy));
    importBtn.disabled = offlineMode || stockRequestBusy || now < importCooldownUntil;
  }

  function finishCommand(button) {
    stockRequestBusy = false;
    if (button === syncBtn) refreshCooldownUntil = Date.now() + 3000;
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
    const startedAt = performance.now();
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 8000);
    try {
      if (!navigator.onLine) throw new Error("offline");
      let res = STOCK_API_BASE
        ? await fetch(apiUrl("/api/stock"), { cache: "no-store", signal: controller.signal })
        : null;
      if (!res) {
        res = await fetch(`stock.xlsx?t=${Date.now()}`, { cache: "no-store", signal: controller.signal });
      }
      if (!res.ok) throw new Error("HTTP " + res.status);
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const rows = mergeWorkbookRows(wb);
      const nextTree = buildTree(rows);
      if (!Object.keys(nextTree).length) throw new Error("ไม่พบข้อมูล stock ที่ใช้ได้");
      tree = nextTree;
      reconcileSelection();
      lastSnapshot = { tree: nextTree, savedAt: Date.now() };
      try { await window.stockCache.write(lastSnapshot); }
      catch (error) { console.warn("Unable to save offline stock", error); }
      renderSearchResults();
      path = [];
      dataSourceName = "stock.xlsx";
      if (!overlay.classList.contains("hidden")) {
        await new Promise((resolve) => setTimeout(resolve, Math.max(0, 1200 - (performance.now() - startedAt))));
      }
      hideOverlay();
      renderLevel();
      setOffline(false);
      return true;
    } catch (err) {
      if (!lastSnapshot) {
        try { lastSnapshot = await window.stockCache.read(); } catch (_) { /* Storage may be unavailable. */ }
      }
      if (lastSnapshot?.tree && Object.keys(lastSnapshot.tree).length) tree = lastSnapshot.tree;
      reconcileSelection();
      path = [];
      closeSearch();
      renderSearchResults();
      renderLevel();
      hideOverlay();
      setOffline(true);
      return false;
    } finally { clearTimeout(timeout); }
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
    overlay.classList.add("overlay--error");
    el("teamLoaderFrame")?.remove();
    overlaySpinner.style.display = "none";
    overlayText.textContent = "โหลด stock.xlsx ไม่สำเร็จ";
    overlayHelp.innerHTML =
      "เบราว์เซอร์บางตัว (เช่น Chrome) บล็อกการอ่านไฟล์ในเครื่องเมื่อเปิดผ่าน <b>file://</b> โดยตรง<br><br>" +
      "วิธีแก้ที่ง่ายที่สุด: เปิดไฟล์ index.html นี้ด้วย <b>Firefox</b> แทน (รองรับการอ่านไฟล์ในโฟลเดอร์เดียวกันได้ทันที)<br><br>" +
      "รายละเอียดข้อผิดพลาด: " + (err && err.message ? err.message : String(err));
  }

  function hideOverlay() {
    overlay.classList.add("hidden");
    setTimeout(() => el("teamLoaderFrame")?.remove(), 300);
  }

  // ---------------------------------------------------------------
  // Import (.csv / .xlsx / .xls)
  // ---------------------------------------------------------------

  function setImportExpanded(expanded) {
    importExpanded = expanded;
    importControl.classList.toggle("import-control--expanded", expanded);
    importBtn.setAttribute("aria-expanded", String(expanded));
    importBtn.setAttribute("aria-label", expanded ? "เลือกไฟล์นำเข้า" : "นำเข้าข้อมูล");
    importCollapse.hidden = !expanded;
  }

  importBtn.addEventListener("click", () => {
    if (offlineMode || stockRequestBusy || Date.now() < importCooldownUntil) return;
    if (!importExpanded) { setImportExpanded(true); return; }
    pickerOpen = true;
    importInput.click();
  });
  importCollapse.addEventListener("click", () => setImportExpanded(false));
  importInput.addEventListener("cancel", () => { pickerOpen = false; setImportExpanded(false); });
  window.addEventListener("focus", () => {
    if (pickerOpen) setTimeout(() => {
      if (pickerOpen) { pickerOpen = false; setImportExpanded(false); }
    }, 300);
  });
  document.addEventListener("pointerdown", (event) => {
    if (importExpanded && !importControl.contains(event.target)) setImportExpanded(false);
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && importExpanded) setImportExpanded(false);
  });

  importInput.addEventListener("change", async (ev) => {
    const file = ev.target.files && ev.target.files[0];
    pickerOpen = false;
    setImportExpanded(false);
    importInput.value = ""; // allow re-selecting the same file again later
    if (!file || offlineMode || stockRequestBusy || Date.now() < importCooldownUntil) return;
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
    reconcileSelection();
    path = [];
    dataSourceName = file.name.replace(/\.pdf$/i, ".xlsx");
    lastSnapshot = { tree: newTree, savedAt: Date.now() };
    try { await window.stockCache.write(lastSnapshot); } catch (error) { console.warn("Unable to save offline stock", error); }
    setOffline(false);
    renderSearchResults();
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
    renderSearchResults();
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
    const corner = 7;
    const outerInset = Math.asin(corner / R_OUTER) * 180 / Math.PI;
    const innerInset = Math.asin(corner / R_INNER) * 180 / Math.PI;
    const point = (radius, angle) => {
      const p = polar(CX, CY, radius, angle);
      return `${p.x.toFixed(2)} ${p.y.toFixed(2)}`;
    };
    return [
      "M", point(R_OUTER, startAngle + outerInset),
      "A", R_OUTER, R_OUTER, 0, large, 1, point(R_OUTER, endAngle - outerInset),
      "Q", point(R_OUTER, endAngle), point(R_OUTER - corner, endAngle),
      "L", point(R_INNER + corner, endAngle),
      "Q", point(R_INNER, endAngle), point(R_INNER, endAngle - innerInset),
      "A", R_INNER, R_INNER, 0, large, 0, point(R_INNER, startAngle + innerInset),
      "Q", point(R_INNER, startAngle), point(R_INNER + corner, startAngle),
      "L", point(R_OUTER - corner, startAngle),
      "Q", point(R_OUTER, startAngle), point(R_OUTER, startAngle + outerInset),
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
    const defs = document.createElementNS(SVG_NS, "defs");
    const surface = document.createElementNS(SVG_NS, "linearGradient");
    surface.setAttribute("id", "tileSurface");
    surface.setAttribute("x2", "0");
    surface.setAttribute("y2", "1");
    for (const [offset, color] of [["0%", "#ffffff"], ["100%", "#e7e7e7"]]) {
      const stop = document.createElementNS(SVG_NS, "stop");
      stop.setAttribute("offset", offset);
      stop.setAttribute("stop-color", color);
      surface.appendChild(stop);
    }
    defs.appendChild(surface);
    wheelSvg.appendChild(defs);

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
      g.dataset.selectionPath = JSON.stringify([...path, item.value]);
      if (selectedPaths.has(JSON.stringify([...path, item.value]))) g.classList.add("seg-selected");
      g.style.transformBox = "view-box";
      g.style.transformOrigin = origin.x.toFixed(2) + "px " + origin.y.toFixed(2) + "px";
      g.style.opacity = "0";
      const visual = document.createElementNS(SVG_NS, "g");
      const orbit = document.createElementNS(SVG_NS, "g");
      orbit.setAttribute("class", "tile-orbit");
      g.appendChild(orbit);
      visual.setAttribute("class", "tile-visual");
      const tileCenter = polar(CX, CY, ARM_RADIUS, mid);
      visual.style.transformBox = "view-box";
      visual.style.transformOrigin = `${tileCenter.x}px ${tileCenter.y}px`;
      orbit.appendChild(visual);

      const segPath = document.createElementNS(SVG_NS, "path");
      segPath.setAttribute("d", donutPath(start, end));
      segPath.setAttribute("class", "segment");
      visual.appendChild(segPath);

      const box = labelBox(mid);
      const fo = document.createElementNS(SVG_NS, "foreignObject");
      fo.setAttribute("x", box.x);
      fo.setAttribute("y", box.y);
      fo.setAttribute("width", box.size);
      fo.setAttribute("height", box.size);

      const wrapper = document.createElementNS(XHTML_NS, "div");
      wrapper.setAttribute("class", "seg-label");
      const labelLength = String(item.label).length;
      wrapper.style.setProperty("--label-size", `${Math.max(11, (n > 6 ? 16 : 18) - Math.max(0, Math.ceil((labelLength - 9) / 4)))}px`);
      wrapper.style.setProperty("--label-width", n > 6 ? "88px" : "110px");
      wrapper.innerHTML = renderLabelContent(item);
      fo.appendChild(wrapper);
      visual.appendChild(fo);

      const activate = (ev) => {
        ev.preventDefault();
        if (g.dataset.longPressConsumed === "true") { g.dataset.longPressConsumed = "false"; return; }
        if (dragState || performance.now() < suppressTileClickUntil) return;
        if (selectionMode) { toggleSelectedPath([...path, item.value]); return; }
        pickValue(item.value);
      };
      g.addEventListener("click", activate);
      bindLongPress(g, [...path, item.value]);
      g.style.cursor = "pointer";

      wheelSvg.appendChild(g);
      groups.push(g);

      currentSegmentsMeta.push({ value: item.value, label: item.label, rangeStart: start, rangeEnd: end, group: g, orbit, visual, tileCenter });
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
    return `<span class="seg-label-inner"><span class="seg-label-name">${escapeHtml(item.label)}</span>${Number.isFinite(item.qty) ? `<span class="seg-label-qty">(${item.qty})</span>` : ""}</span>`;
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
  // Hover uses the original sector geometry even while other tiles shrink.
  // ---------------------------------------------------------------

  function setArmed(segMeta, focus = 0) {
    const selectedIndex = currentSegmentsMeta.indexOf(segMeta);
    const active = selectedIndex >= 0;
    if (active && wheelSvg.lastElementChild !== segMeta.group) {
      wheelSvg.appendChild(segMeta.group);
    }
    currentSegmentsMeta.forEach((s) => {
      if (s.group) s.group.classList.toggle("seg-armed", s === segMeta);
      s.orbit.style.transform = "rotate(0deg)";
      s.visual.style.transform = `scale(${active && s === segMeta ? 1.16 : 1})`;
      s.visual.style.opacity = String(active && s !== segMeta ? .62 : 1);
    });
    const readout = el("pointerReadout");
    readout.textContent = segMeta ? segMeta.label : "";
    readout.classList.toggle("pointer-readout--visible", !!segMeta);
    readout.style.fontSize = `${Math.max(17, 29 - Math.max(0, String(segMeta?.label || "").length - 15) * .45)}px`;
  }

  function onHubPointerDown(ev) {
    if (currentSegmentsMeta.length === 0 || dragState || hubReturning || !ev.isPrimary || ev.button !== 0) return;
    ev.preventDefault();
    try { hub.setPointerCapture(ev.pointerId); } catch (e) { /* ignore */ }
    const rect = wheelWrap.getBoundingClientRect();
    dragState = {
      pointerId: ev.pointerId,
      startX: rect.left + rect.width / 2,
      startY: rect.top + rect.height / 2,
      lastX: 0,
      lastY: 0,
      pointed: null,
      scale: 400 / rect.width // svg units per screen px (wheel is a square)
    };
    anime.remove(hub);
    anime.remove(currentSegmentsMeta.map((s) => s.group));
    anime.set(currentSegmentsMeta.map((s) => s.group), { scale: 1, opacity: 1 });
    anime.set(hub, { translateX: ev.clientX - dragState.startX, translateY: ev.clientY - dragState.startY, scale: .256 });
    hub.classList.add("hub--dragging");
    wheelWrap.classList.add("pointer-active");
  }

  function onHubPointerMove(ev) {
    if (!dragState || ev.pointerId !== dragState.pointerId) return;

    ev.preventDefault();
    const dxPx = ev.clientX - dragState.startX;
    const dyPx = ev.clientY - dragState.startY;
    anime.set(hub, { translateX: dxPx, translateY: dyPx, scale: .256 });
    const targetX = dxPx * dragState.scale;
    const targetY = dyPx * dragState.scale;
    const fromX = dragState.lastX, fromY = dragState.lastY;
    const steps = Math.max(1, Math.ceil(Math.hypot(targetX - fromX, targetY - fromY) / 8));
    for (let step = 1; step <= steps && dragState; step++) {
      const dx = fromX + (targetX - fromX) * step / steps;
      const dy = fromY + (targetY - fromY) * step / steps;
      const radius = Math.hypot(dx, dy);
      const angle = angleFromDelta(dx, dy);
      const tile = currentSegmentsMeta.find((s) => angle >= s.rangeStart && angle < s.rangeEnd) || null;
      if (radius >= R_INNER && radius <= R_OUTER) {
        dragState.pointed = tile;
        setArmed(tile, tile ? 1 : 0);
        if (tile && radius >= R_INNER + (R_OUTER - R_INNER) * .75) {
          commitDragSelection(tile);
        }
      } else if (radius < R_INNER || !tile) {
        dragState.pointed = null;
        setArmed(null);
      } else if (tile && tile === dragState.pointed) {
        commitDragSelection(tile);
      } else {
        setArmed(tile === dragState.pointed ? tile : null);
      }
    }
    if (dragState) { dragState.lastX = targetX; dragState.lastY = targetY; }
  }

  function commitDragSelection(segMeta) {
    // Consume this gesture before changing the menu; held pointers cannot select again.
    releaseHub();
    ringOriginPoint = null;
    pickValue(segMeta.value);
  }

  function releaseHub() {
    suppressTileClickUntil = performance.now() + 400;
    const pointerId = dragState && dragState.pointerId;
    dragState = null;
    if (pointerId != null && hub.hasPointerCapture(pointerId)) hub.releasePointerCapture(pointerId);
    hub.classList.remove("hub--dragging");
    wheelWrap.classList.remove("pointer-active");
    setArmed(null);
    hubReturning = true;
    anime.remove(hub);
    anime({
      targets: hub,
      translateX: 0,
      translateY: 0,
      scale: 1,
      duration: 320,
      easing: "easeOutCubic",
      complete: () => { hubReturning = false; }
    });
  }

  function onHubPointerUp(ev) {
    if (!dragState || ev.pointerId !== dragState.pointerId) return;
    onHubPointerMove(ev);
    if (!dragState) return;
    const radius = Math.hypot(dragState.lastX, dragState.lastY);
    if (dragState.pointed && radius >= R_INNER && radius <= R_OUTER + 20) {
      commitDragSelection(dragState.pointed);
    } else {
      releaseHub();
    }
  }

  function onHubPointerCancel(ev) {
    if (!dragState || ev.pointerId !== dragState.pointerId) return;
    releaseHub();
  }

  hub.addEventListener("pointerdown", onHubPointerDown);
  hub.addEventListener("pointermove", onHubPointerMove);
  hub.addEventListener("pointerup", onHubPointerUp);
  hub.addEventListener("pointercancel", onHubPointerCancel);
  hub.addEventListener("lostpointercapture", (ev) => {
    if (dragState && ev.pointerId === dragState.pointerId) releaseHub();
  });

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
      card.dataset.selectionPath = JSON.stringify([...path, item.value]);
      if (selectedPaths.has(JSON.stringify([...path, item.value]))) card.classList.add("is-selected");
      card.type = "button";
      card.style.setProperty("--grid-i", index);
      const initial = Array.from(item.label.trim()).find((char) => /[\p{L}\p{N}]/u.test(char)) || "#";
      const tone = (initial.toUpperCase().codePointAt(0) * 5) % 26;
      const background = 222 + tone;
      const accent = 130 + tone * 3;
      card.style.setProperty("--grid-gray", `rgb(${background} ${background} ${background})`);
      card.style.setProperty("--grid-gray-edge", `rgb(${accent} ${accent} ${accent})`);
      card.innerHTML = `
        <span class="product-grid-name">${escapeHtml(item.label)}</span>
        <span class="product-grid-qty">${Number.isFinite(item.qty) ? `(${item.qty})` : ""}</span>
        <span class="product-grid-arrow">›</span>
      `;
      card.addEventListener("click", (ev) => {
        ev.preventDefault();
        if (card.dataset.longPressConsumed === "true") { card.dataset.longPressConsumed = "false"; return; }
        if (selectionMode) { toggleSelectedPath([...path, item.value]); return; }
        pickValue(item.value);
      });
      bindLongPress(card, [...path, item.value]);
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
      hubSub.textContent = "";
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
    if (existing) {
      if (existing.codeResizeObserver) existing.codeResizeObserver.disconnect();
      existing.remove();
    }
    closeCodeOverlay(false);
    stage.classList.remove("stage--result");
    wheelWrap.classList.remove("wheel-wrap--result");
    hubWrap.style.display = "flex";
  }

  function fitCode(element, maximum) {
    if (!element || !element.isConnected || !element.clientWidth) return;
    let size = maximum;
    element.style.fontSize = `${size}px`;
    while (element.scrollWidth > element.clientWidth && size > 1) {
      size -= 0.5;
      element.style.fontSize = `${size}px`;
    }
  }

  function closeCodeOverlay(restoreFocus = true) {
    if (codeOverlay.hidden) return;
    codeOverlay.hidden = true;
    if (restoreFocus && codeReturnFocus && codeReturnFocus.isConnected) codeReturnFocus.focus();
    codeReturnFocus = null;
  }

  function openCodeOverlay(code, returnFocus) {
    codeReturnFocus = returnFocus;
    expandedItemCode.textContent = code;
    codeOverlay.hidden = false;
    expandedItemCode.focus();
    requestAnimationFrame(() => fitCode(expandedItemCode, 48));
    document.fonts.ready.then(() => fitCode(expandedItemCode, 48));
  }

  codeOverlay.addEventListener("click", (event) => {
    if (event.target === codeOverlay || event.target === expandedItemCode) closeCodeOverlay();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !codeOverlay.hidden) closeCodeOverlay();
  });
  window.addEventListener("resize", () => {
    if (!codeOverlay.hidden) fitCode(expandedItemCode, 48);
  });

  function captureOverviewState(card) {
    const toggle = card.querySelector(".overview-toggle");
    return {
      open: toggle.getAttribute("aria-expanded") === "true",
      stageScrollTop: stage.scrollTop,
      expanded: new Set(Array.from(card.querySelectorAll(".overview-group-toggle[aria-expanded='true']"), (button) => button.dataset.storage))
    };
  }

  function renderOverview(card, previousState) {
    const modelStock = tree[path[0]] && tree[path[0]][path[1]];
    const toggle = card.querySelector(".overview-toggle");
    const panel = card.querySelector(".overview-panel");
    const list = card.querySelector(".overview-list");

    for (const [storage, colors] of Object.entries(modelStock || {})) {
      const available = Object.entries(colors).filter(([, stock]) => getTotalQty(stock) > 0);
      if (!available.length) continue;
      const group = document.createElement("div");
      group.className = "overview-group";
      const groupToggle = document.createElement("button");
      groupToggle.type = "button";
      groupToggle.className = "overview-group-toggle";
      groupToggle.dataset.storage = storage;
      const groupOpen = previousState ? previousState.expanded.has(storage) : true;
      groupToggle.setAttribute("aria-expanded", String(groupOpen));
      const groupTitle = document.createElement("span");
      groupTitle.textContent = storage;
      const groupCount = document.createElement("span");
      groupCount.className = "overview-group-count";
      groupCount.textContent = `${available.reduce((sum, [, stock]) => sum + getTotalQty(stock), 0)} เครื่อง`;
      groupToggle.append(groupTitle, groupCount);
      const colorsBox = document.createElement("div");
      colorsBox.className = "overview-colors";
      colorsBox.hidden = !groupOpen;
      groupToggle.addEventListener("click", () => {
        const next = groupToggle.getAttribute("aria-expanded") !== "true";
        groupToggle.setAttribute("aria-expanded", String(next));
        colorsBox.hidden = !next;
      });
      for (const [color, stock] of available) {
        const choice = document.createElement("button");
        choice.type = "button";
        choice.className = "overview-color" + (storage === path[2] && color === path[3] ? " overview-color--selected" : "");
        const colorName = document.createElement("span");
        colorName.textContent = color;
        const colorQty = document.createElement("span");
        colorQty.className = "overview-color-qty";
        colorQty.textContent = `(${getTotalQty(stock)})`;
        choice.append(colorName, colorQty);
        choice.addEventListener("click", () => {
          pendingOverviewState = captureOverviewState(card);
          path = [path[0], path[1], storage, color];
          renderLevel();
        });
        colorsBox.appendChild(choice);
      }
      group.append(groupToggle, colorsBox);
      list.appendChild(group);
    }

    const open = previousState ? previousState.open : false;
    toggle.setAttribute("aria-expanded", String(open));
    panel.hidden = !open;
    toggle.addEventListener("click", () => {
      const next = toggle.getAttribute("aria-expanded") !== "true";
      toggle.setAttribute("aria-expanded", String(next));
      panel.hidden = !next;
    });
    requestAnimationFrame(() => {
      if (previousState) stage.scrollTop = previousState.stageScrollTop;
    });
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
    const overviewState = pendingOverviewState;
    pendingOverviewState = null;

    card.innerHTML = `
      <div class="result-head">
        <div class="result-eyebrow">${escapeHtml(path[0])} · ${escapeHtml(path[2])}</div>
        <div class="result-model">${escapeHtml(path[1])}</div>
        <a class="result-action-btn" href="${escapeHtml(imageSearchUrl)}" target="_blank" rel="noopener noreferrer">ดูรูป</a>
        <button class="result-select-btn" type="button">เพิ่มในรายการ</button>
        ${itemCode ? `<div class="result-code-section"><button class="result-item-code" type="button" aria-label="ขยายรหัสสินค้า">${escapeHtml(itemCode)}</button><button class="code-zoom-btn" type="button" aria-label="ขยายรหัสสินค้า" title="ขยายรหัสสินค้า"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="10.8" cy="10.8" r="6.8" fill="none" stroke="currentColor" stroke-width="2"/><path d="m16 16 5 5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg></button></div>` : ""}
      </div>
      <div class="result-body">
        <div class="result-variant">${escapeHtml(path[3])}</div>
        <div class="result-qty-row">
          <div class="result-qty-num">${qty}</div>
          <div class="result-qty-label">เครื่องคงเหลือ</div>
        </div>
        <div class="result-bar-track"><div class="result-bar-fill" id="resultBarFill"></div></div>
        <div class="stock-overview">
          <button class="overview-toggle" type="button" aria-expanded="false">ความจุและสีที่มีสต็อก <span class="overview-chevron" aria-hidden="true">⌄</span></button>
          <div class="overview-panel" hidden><div class="overview-list"></div></div>
        </div>
      </div>
    `;
    pos.appendChild(card);
    wheelWrap.appendChild(pos);
    stage.classList.add("stage--result");
    wheelWrap.classList.add("wheel-wrap--result");
    renderOverview(card, overviewState);
    const resultSelect = card.querySelector(".result-select-btn");
    resultSelect.addEventListener("click", () => toggleSelectedPath([...path]));
    resultSelect.textContent = selectedPaths.has(JSON.stringify(path)) ? "เอาออกจากรายการ" : "เพิ่มในรายการ";

    const codeButton = card.querySelector(".result-item-code");
    if (codeButton) {
      codeButton.addEventListener("click", () => openCodeOverlay(itemCode, codeButton));
      card.querySelector(".code-zoom-btn").addEventListener("click", () => openCodeOverlay(itemCode, codeButton));
      const observer = new ResizeObserver(() => fitCode(codeButton, 36));
      observer.observe(codeButton);
      pos.codeResizeObserver = observer;
      document.fonts.ready.then(() => fitCode(codeButton, 36));
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

  function bindLongPress(element, targetPath) {
    let timer = null, startX = 0, startY = 0;
    element.addEventListener("pointerdown", (event) => {
      if (!event.isPrimary || event.button !== 0) return;
      if (selectionMode) return;
      startX = event.clientX; startY = event.clientY;
      clearTimeout(timer);
      timer = setTimeout(() => {
        element.dataset.longPressConsumed = "true";
        toggleSelectedPath(targetPath);
        if (navigator.vibrate) navigator.vibrate(25);
      }, 500);
    });
    element.addEventListener("pointermove", (event) => {
      if (Math.hypot(event.clientX - startX, event.clientY - startY) > 12) clearTimeout(timer);
    });
    ["pointerup", "pointercancel", "pointerleave"].forEach((name) => element.addEventListener(name, () => clearTimeout(timer)));
    element.addEventListener("pointerup", () => {
      // Keep suppression for the entire hold, then consume its release click.
      setTimeout(() => { element.dataset.longPressConsumed = "false"; }, 400);
    });
    element.addEventListener("pointercancel", () => { element.dataset.longPressConsumed = "false"; });
    element.addEventListener("contextmenu", (event) => event.preventDefault());
  }

  function toggleSelectedPath(targetPath) {
    const enteringMode = !selectionMode;
    selectionMode = true;
    const key = JSON.stringify(targetPath);
    if (selectedPaths.has(key)) selectedPaths.delete(key);
    else selectedPaths.set(key, targetPath);
    if (selectedPaths.size === 0) { exitSelectionMode(); return; }
    updateSelectionButton();
    if (enteringMode) showToast("โหมดเลือก: แตะรายการเพื่อเลือกหรือยกเลิก", "ok");
    document.querySelectorAll("[data-selection-path]").forEach((element) => {
      const active = selectedPaths.has(element.dataset.selectionPath);
      element.classList.toggle("seg-selected", active && element.classList.contains("seg-group"));
      element.classList.toggle("is-selected", active && element.classList.contains("product-grid-card"));
    });
    const resultSelect = document.querySelector(".result-select-btn");
    if (resultSelect) resultSelect.textContent = selectedPaths.has(JSON.stringify(path)) ? "เอาออกจากรายการ" : "เพิ่มในรายการ";
    showToast(selectedPaths.has(key) ? "เพิ่มในรายการที่เลือกแล้ว" : "เอาออกจากรายการแล้ว", "ok");
  }

  function updateSelectionButton() {
    selectionOpen.hidden = !selectionMode;
    selectionExit.hidden = !selectionMode;
    selectionCount.textContent = String(selectedPaths.size);
  }

  function exitSelectionMode() {
    const modalWasOpen = !selectionModal.hidden;
    selectedPaths.clear();
    excludedLeaves.clear();
    selectionMode = false;
    selectionModal.hidden = true;
    updateSelectionButton();
    document.querySelectorAll(".seg-selected,.product-grid-card.is-selected").forEach((element) => {
      element.classList.remove("seg-selected", "is-selected");
    });
    const resultSelect = document.querySelector(".result-select-btn");
    if (resultSelect) resultSelect.textContent = "เพิ่มในรายการ";
    if (modalWasOpen) searchToggle.focus();
  }

  function selectedLeaves() {
    const leaves = new Map();
    const visit = (parts, node) => {
      if (!node) return;
      if (parts.length === 4) {
        if (getTotalQty(node) > 0) leaves.set(JSON.stringify(parts), { parts, qty: getTotalQty(node) });
        return;
      }
      Object.entries(node).forEach(([key, child]) => visit([...parts, key], child));
    };
    selectedPaths.forEach((parts) => visit(parts, resolve(parts)));
    return Array.from(leaves.values()).sort((a, b) => a.parts.join("\u0000").localeCompare(b.parts.join("\u0000"), "th"));
  }

  function previewText(leaves) {
    const included = leaves.filter((leaf) => !excludedLeaves.has(JSON.stringify(leaf.parts)));
    const groups = new Map();
    included.forEach(({ parts: [category, model, storage, color], qty }) => {
      if (!groups.has(category)) groups.set(category, new Map());
      const models = groups.get(category);
      if (!models.has(model)) models.set(model, new Map());
      const storages = models.get(model);
      if (!storages.has(storage)) storages.set(storage, []);
      storages.get(storage).push(includeQuantity.checked ? `${color} (${qty})` : color);
    });
    const lines = ["สินค้าที่มีสต็อก"];
    groups.forEach((models, category) => {
      lines.push("", category);
      models.forEach((storages, model) => {
        lines.push(model);
        storages.forEach((colors, storage) => lines.push(`- ${storage}: ${colors.join(", ")}`));
        lines.push("");
      });
    });
    return included.length ? lines.join("\n").trim() : "";
  }

  function renderSelectionDialog() {
    const leaves = selectedLeaves();
    if (!leaves.some(leaf => !excludedLeaves.has(JSON.stringify(leaf.parts)))) {
      exitSelectionMode();
      return;
    }
    selectionList.replaceChildren();
    const categories = new Map();
    leaves.forEach((leaf) => {
      const [category, model, storage] = leaf.parts;
      if (!categories.has(category)) categories.set(category, new Map());
      const models = categories.get(category);
      if (!models.has(model)) models.set(model, new Map());
      const storages = models.get(model);
      if (!storages.has(storage)) storages.set(storage, []);
      storages.get(storage).push(leaf);
    });
    const checkbox = (name, subset, className) => {
      const label = document.createElement("label");
      label.className = className;
      const input = document.createElement("input");
      input.type = "checkbox";
      const checked = subset.filter((leaf) => !excludedLeaves.has(JSON.stringify(leaf.parts))).length;
      input.checked = checked === subset.length;
      input.indeterminate = checked > 0 && checked < subset.length;
      input.addEventListener("change", () => {
        subset.forEach((leaf) => {
          const key = JSON.stringify(leaf.parts);
          if (input.checked) excludedLeaves.delete(key); else excludedLeaves.add(key);
        });
        if (leaves.every((leaf) => excludedLeaves.has(JSON.stringify(leaf.parts)))) {
          exitSelectionMode();
          return;
        }
        renderSelectionDialog();
      });
      label.append(input, document.createTextNode(name));
      return label;
    };
    categories.forEach((models, category) => {
      const section = document.createElement("section");
      section.className = "selection-category";
      const title = document.createElement("h3"); title.textContent = category; section.append(title);
      models.forEach((storages, model) => {
        const modelBox = document.createElement("div"); modelBox.className = "selection-model";
        const modelLeaves = Array.from(storages.values()).flat();
        modelBox.append(checkbox(model, modelLeaves, "selection-model-label"));
        storages.forEach((stock, storage) => {
          const storageBox = document.createElement("div"); storageBox.className = "selection-storage";
          storageBox.append(checkbox(storage, stock, "selection-storage-label"));
          const colors = document.createElement("div"); colors.className = "selection-colors";
          stock.forEach((leaf) => colors.append(checkbox(includeQuantity.checked ? `${leaf.parts[3]} · ${leaf.qty} เครื่อง` : leaf.parts[3], [leaf], "selection-color-label")));
          storageBox.append(colors); modelBox.append(storageBox);
        });
        section.append(modelBox);
      });
      selectionList.append(section);
    });
    if (!leaves.length) selectionList.textContent = "ไม่มีรายการที่มีสต็อกในชุดที่เลือก";
    selectionText.textContent = previewText(leaves) || "ไม่มีรายการสำหรับคัดลอก";
    selectionCopy.disabled = !previewText(leaves);
  }

  selectionOpen.addEventListener("click", () => {
    renderSelectionDialog();
    if (!selectionMode) return;
    selectionModal.hidden = false;
    el("selectionClose").focus();
  });
  function closeSelectionDialog() { selectionModal.hidden = true; selectionOpen.focus(); }
  el("selectionClose").addEventListener("click", closeSelectionDialog);
  selectionModal.addEventListener("click", (event) => { if (event.target === selectionModal) closeSelectionDialog(); });
  document.addEventListener("keydown", (event) => { if (event.key === "Escape" && !selectionModal.hidden) closeSelectionDialog(); });
  includeQuantity.addEventListener("change", renderSelectionDialog);
  selectionModal.addEventListener("keydown", event => {
    if (event.key !== "Tab") return;
    const controls = Array.from(selectionModal.querySelectorAll('button:not(:disabled),input:not(:disabled)'));
    const first = controls[0], last = controls[controls.length - 1];
    if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
    if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
  });
  function reconcileSelection() {
    if (!selectionMode) return;
    const oldCount = selectedPaths.size;
    for (const [key, parts] of selectedPaths) {
      if (!getTotalQty(resolve(parts))) selectedPaths.delete(key);
    }
    const available = new Set(selectedLeaves().map(leaf => JSON.stringify(leaf.parts)));
    for (const key of excludedLeaves) if (!available.has(key)) excludedLeaves.delete(key);
    if (!selectedPaths.size || !selectedLeaves().some(leaf => !excludedLeaves.has(JSON.stringify(leaf.parts)))) exitSelectionMode();
    else {
      updateSelectionButton();
      if (!selectionModal.hidden) renderSelectionDialog();
    }
    if (oldCount !== selectedPaths.size) showToast("รายการที่เลือกเปลี่ยนตามข้อมูลสต็อกล่าสุด", "ok");
  }
  selectionExit.addEventListener("click", exitSelectionMode);
  selectionCopy.addEventListener("click", async () => {
    const value = previewText(selectedLeaves());
    if (!value) return;
    try {
      if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(value);
      else {
        const input = document.createElement("textarea");
        input.value = value; document.body.append(input); input.select();
        if (!document.execCommand("copy")) throw new Error("Copy failed");
        input.remove();
      }
      showToast("คัดลอกข้อความแล้ว", "ok");
    } catch (error) { showToast("คัดลอกไม่สำเร็จ กรุณาลองอีกครั้ง", "err"); }
  });


  backBtn.addEventListener("click", () => {
    if (path.length === 0 || !allowNavigation()) return;
    exitSelectionMode();
    path.pop();
    renderLevel();
  });

  resetBtn.addEventListener("click", () => {
    if (!allowNavigation()) return;
    exitSelectionMode();
    closeSearch();
    path = [];
    renderLevel();
  });

  syncBtn.addEventListener("click", async () => {
    if (stockRequestBusy || Date.now() < refreshCooldownUntil) return;
    exitSelectionMode();
    stockRequestBusy = true;
    updateCommandButtons();
    closeSearch();
    path = [];
    try {
      if (await loadStock(false)) showToast("โหลดข้อมูลสต็อกล่าสุดแล้ว", "ok");
    } finally {
      finishCommand(syncBtn);
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
