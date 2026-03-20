// PDF Masking Tool Application
(function () {
  'use strict';

  // Configure PDF.js worker
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  // --- State ---
  const state = {
    pdfDoc: null,
    pdfBytes: null,
    pdfFileName: '',
    currentPage: 1,
    totalPages: 0,
    scale: 1.5,
    currentTool: 'rect',
    maskColor: '#000000',
    isDrawing: false,
    startX: 0,
    startY: 0,
    // Per-page mask data: { pageNum: [{ type, data, color, size }] }
    masks: {},
    currentPath: [],
    // For rect preview
    rectPreview: null,
    // Selection state
    selectedMaskIndex: -1,
    selDragMode: null, // 'move' | 'resize-nw' | 'resize-ne' | 'resize-sw' | 'resize-se' | null
    selDragStart: null,
    selOrigData: null,
  };

  // --- DOM ---
  const $ = (sel) => document.querySelector(sel);
  const dropZone = $('#dropZone');
  const fileInput = $('#fileInput');
  const fileBtn = $('#fileBtn');
  const toolbar = $('#toolbar');
  const pageNav = $('#pageNav');
  const canvasWrapper = $('#canvasWrapper');
  const pdfCanvas = $('#pdfCanvas');
  const maskCanvas = $('#maskCanvas');
  const pdfCtx = pdfCanvas.getContext('2d');
  const maskCtx = maskCanvas.getContext('2d');
  const loadingOverlay = $('#loadingOverlay');
  const loadingText = $('#loadingText');

  // --- Reset to top ---
  $('#headerLogo').addEventListener('click', () => {
    if (!state.pdfDoc) return;
    if (!confirm('現在の編集内容は破棄されます。トップに戻りますか？')) return;
    resetApp();
  });

  function resetApp() {
    state.pdfDoc = null;
    state.pdfBytes = null;
    state.currentPage = 1;
    state.totalPages = 0;
    state.masks = {};
    state.currentPath = [];
    state.isDrawing = false;

    pdfCtx.clearRect(0, 0, pdfCanvas.width, pdfCanvas.height);
    maskCtx.clearRect(0, 0, maskCanvas.width, maskCanvas.height);
    fileInput.value = '';

    toolbar.style.display = 'none';
    pageNav.style.display = 'none';
    canvasWrapper.style.display = 'none';
    dropZone.style.display = 'flex';
  }

  // --- File Loading ---
  fileBtn.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) loadPDF(e.target.files[0]);
  });

  // Drag & Drop
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });
  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('drag-over');
  });
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') {
      loadPDF(file);
    }
  });

  async function loadPDF(file) {
    showLoading('PDFを読み込み中...');
    try {
      state.pdfFileName = file.name.replace(/\.pdf$/i, '');
      const arrayBuffer = await file.arrayBuffer();
      state.pdfBytes = new Uint8Array(arrayBuffer);
      state.pdfDoc = await pdfjsLib.getDocument({ data: state.pdfBytes.slice() }).promise;
      state.totalPages = state.pdfDoc.numPages;
      state.currentPage = 1;
      state.masks = {};

      // Show UI
      dropZone.style.display = 'none';
      toolbar.style.display = 'flex';
      pageNav.style.display = 'flex';
      canvasWrapper.style.display = 'flex';

      $('#totalPages').textContent = state.totalPages;
      state.scale = 1.0; // デフォルト100%
      await renderPage();
    } catch (err) {
      alert('PDFの読み込みに失敗しました: ' + err.message);
    } finally {
      hideLoading();
    }
  }

  // --- Page Rendering ---
  async function renderPage() {
    const page = await state.pdfDoc.getPage(state.currentPage);
    const viewport = page.getViewport({ scale: state.scale });

    pdfCanvas.width = viewport.width;
    pdfCanvas.height = viewport.height;
    maskCanvas.width = viewport.width;
    maskCanvas.height = viewport.height;

    await page.render({ canvasContext: pdfCtx, viewport }).promise;

    // Redraw saved masks for this page
    redrawMasks();
    $('#currentPage').textContent = state.currentPage;
    updateZoomLabel();
  }

  function redrawMasks() {
    maskCtx.clearRect(0, 0, maskCanvas.width, maskCanvas.height);
    const pageMasks = state.masks[state.currentPage] || [];
    for (const mask of pageMasks) {
      drawMask(maskCtx, mask);
    }
  }

  function drawMask(ctx, mask) {
    ctx.save();
    if (mask.type === 'rect') {
      ctx.fillStyle = mask.color;
      ctx.fillRect(mask.data.x, mask.data.y, mask.data.w, mask.data.h);
    } else if (mask.type === 'pen') {
      ctx.strokeStyle = mask.color;
      ctx.lineWidth = mask.size;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.beginPath();
      const pts = mask.data;
      if (pts.length > 0) {
        ctx.moveTo(pts[0].x, pts[0].y);
        for (let i = 1; i < pts.length; i++) {
          ctx.lineTo(pts[i].x, pts[i].y);
        }
      }
      ctx.stroke();
    } else if (mask.type === 'free') {
      // Freeform filled shape
      ctx.fillStyle = mask.color;
      ctx.beginPath();
      const pts = mask.data;
      if (pts.length > 0) {
        ctx.moveTo(pts[0].x, pts[0].y);
        for (let i = 1; i < pts.length; i++) {
          ctx.lineTo(pts[i].x, pts[i].y);
        }
        ctx.closePath();
      }
      ctx.fill();

      // Also draw thick stroke to cover edges
      ctx.strokeStyle = mask.color;
      ctx.lineWidth = mask.size;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.beginPath();
      if (pts.length > 0) {
        ctx.moveTo(pts[0].x, pts[0].y);
        for (let i = 1; i < pts.length; i++) {
          ctx.lineTo(pts[i].x, pts[i].y);
        }
        ctx.closePath();
      }
      ctx.stroke();
    }
    ctx.restore();
  }

  // --- Fit to screen (cycle: 縦フィット → 横フィット → 100%) ---
  let fitMode = 0; // 0=縦, 1=横, 2=100%
  const fitLabels = ['縦フィット', '横フィット', '100%'];

  async function cycleFit() {
    if (!state.pdfDoc) return;
    const currentMode = fitMode;
    // 即座に次のモードに更新（非同期待ち前に）
    fitMode = (fitMode + 1) % 3;
    $('#zoomFit').textContent = fitLabels[fitMode];

    const page = await state.pdfDoc.getPage(state.currentPage);
    const viewport = page.getViewport({ scale: 1 });
    const maxW = window.innerWidth - 60;
    const maxH = window.innerHeight - 220;

    if (currentMode === 0) {
      // 縦フィット: 高さに合わせる
      state.scale = Math.min(maxH / viewport.height, 4);
    } else if (currentMode === 1) {
      // 横フィット: 幅に合わせる
      state.scale = Math.min(maxW / viewport.width, 4);
    } else {
      // 100%
      state.scale = 1.0;
    }

    renderPage();
  }

  // --- Tool Selection ---
  document.querySelectorAll('.tool-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tool-btn').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      state.currentTool = btn.dataset.tool;

      const isSelect = state.currentTool === 'select';
      $('#drawOptionsGroup').style.display = isSelect ? 'none' : 'flex';
      $('#selectInfoGroup').style.display = isSelect ? 'flex' : 'none';
      maskCanvas.style.cursor = isSelect ? 'default' : 'crosshair';

      if (!isSelect) {
        deselectMask();
      }
    });
  });

  // Color
  $('#maskColor').addEventListener('input', (e) => {
    state.maskColor = e.target.value;
  });
  // --- Drawing ---
  function getPos(e) {
    const rect = maskCanvas.getBoundingClientRect();
    const scaleX = maskCanvas.width / rect.width;
    const scaleY = maskCanvas.height / rect.height;
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    return {
      x: (clientX - rect.left) * scaleX,
      y: (clientY - rect.top) * scaleY,
    };
  }

  // --- Auto-scroll during draw/drag ---
  const SCROLL_EDGE = 40; // px from container edge to trigger scroll
  const SCROLL_SPEED = 10; // px per frame
  let autoScrollRAF = null;
  let lastClientX = 0, lastClientY = 0;

  function autoScrollContainer(el) {
    // Check if el is scrollable and cursor is near its edges; scroll if so
    if (!el) return false;
    const isScrollableY = el.scrollHeight > el.clientHeight;
    const isScrollableX = el.scrollWidth > el.clientWidth;
    if (!isScrollableY && !isScrollableX) return false;

    const r = el.getBoundingClientRect();
    let didScroll = false;

    if (isScrollableY) {
      if (lastClientY > r.bottom - SCROLL_EDGE && lastClientY >= r.top) {
        const before = el.scrollTop;
        el.scrollTop += SCROLL_SPEED;
        if (el.scrollTop !== before) didScroll = true;
      } else if (lastClientY < r.top + SCROLL_EDGE && lastClientY <= r.bottom) {
        const before = el.scrollTop;
        el.scrollTop -= SCROLL_SPEED;
        if (el.scrollTop !== before) didScroll = true;
      }
    }
    if (isScrollableX) {
      if (lastClientX > r.right - SCROLL_EDGE && lastClientX >= r.left) {
        const before = el.scrollLeft;
        el.scrollLeft += SCROLL_SPEED;
        if (el.scrollLeft !== before) didScroll = true;
      } else if (lastClientX < r.left + SCROLL_EDGE && lastClientX <= r.right) {
        const before = el.scrollLeft;
        el.scrollLeft -= SCROLL_SPEED;
        if (el.scrollLeft !== before) didScroll = true;
      }
    }
    return didScroll;
  }

  function startAutoScroll() {
    if (autoScrollRAF) return;
    function tick() {
      if (!state.isDrawing && !state.selDragMode) {
        autoScrollRAF = requestAnimationFrame(tick);
        return;
      }

      // Try each scrollable container in order
      const scrolled =
        autoScrollContainer(canvasWrapper) ||
        autoScrollContainer(document.querySelector('.main-content')) ||
        autoScrollContainer(document.documentElement);

      if (scrolled) {
        const syntheticPos = getPosFromClient(lastClientX, lastClientY);
        if (state.currentTool === 'select') {
          handleSelectDrag(syntheticPos);
        } else {
          drawingAt(syntheticPos);
        }
      }

      autoScrollRAF = requestAnimationFrame(tick);
    }
    autoScrollRAF = requestAnimationFrame(tick);
  }

  function stopAutoScroll() {
    if (autoScrollRAF) {
      cancelAnimationFrame(autoScrollRAF);
      autoScrollRAF = null;
    }
  }

  function getPosFromClient(clientX, clientY) {
    const rect = maskCanvas.getBoundingClientRect();
    const scaleX = maskCanvas.width / rect.width;
    const scaleY = maskCanvas.height / rect.height;
    return {
      x: (clientX - rect.left) * scaleX,
      y: (clientY - rect.top) * scaleY,
    };
  }

  maskCanvas.addEventListener('mousedown', startDraw);
  // Use document-level listeners so drawing continues outside canvas
  document.addEventListener('mousemove', (e) => {
    lastClientX = e.clientX;
    lastClientY = e.clientY;
    if (state.isDrawing || state.selDragMode) {
      drawing(e);
    }
  });
  document.addEventListener('mouseup', (e) => {
    if (state.isDrawing || state.selDragMode) {
      endDraw(e);
    }
  });

  // Touch support — 描画ツール使用時のみスクロールを無効化
  maskCanvas.addEventListener('touchstart', (e) => {
    if (state.currentTool !== 'select') {
      // 描画ツール → 常にスクロール無効化
      e.preventDefault();
    } else {
      // selectツール → マスク上タッチ時のみ無効化
      const pos = getPos(e);
      const onMask = state.masks.some(m => hitTestMask(pos, m));
      if (onMask || state.selectedMaskIndex >= 0) {
        e.preventDefault();
      }
    }
    startDraw(e);
  }, { passive: false });
  maskCanvas.addEventListener('touchmove', (e) => {
    if (state.isDrawing || state.selDragMode) {
      e.preventDefault();
    }
    drawing(e);
  }, { passive: false });
  maskCanvas.addEventListener('touchend', (e) => {
    if (state.isDrawing || state.selDragMode) {
      e.preventDefault();
    }
    endDraw(e);
  }, { passive: false });

  function startDraw(e) {
    const pos = getPos(e);
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    lastClientX = clientX;
    lastClientY = clientY;

    // --- Select tool ---
    if (state.currentTool === 'select') {
      handleSelectStart(pos);
      startAutoScroll();
      return;
    }

    state.isDrawing = true;
    state.startX = pos.x;
    state.startY = pos.y;

    if (state.currentTool === 'pen-thick' || state.currentTool === 'pen-thin' || state.currentTool === 'free') {
      state.currentPath = [{ x: pos.x, y: pos.y }];
    }
    startAutoScroll();
  }

  function getPenSize(tool) {
    if (tool === 'pen-thick') return 8;
    if (tool === 'pen-thin') return 2;
    return 8; // fallback for free
  }

  function drawing(e) {
    const pos = getPos(e);

    // --- Select tool drag ---
    if (state.currentTool === 'select') {
      handleSelectDrag(pos);
      return;
    }

    if (!state.isDrawing) return;
    drawingAt(pos);
  }

  function drawingAt(pos) {
    if (state.currentTool === 'rect') {
      redrawMasks();
      maskCtx.save();
      maskCtx.fillStyle = state.maskColor;
      maskCtx.globalAlpha = 0.5;
      const x = Math.min(state.startX, pos.x);
      const y = Math.min(state.startY, pos.y);
      const w = Math.abs(pos.x - state.startX);
      const h = Math.abs(pos.y - state.startY);
      maskCtx.fillRect(x, y, w, h);
      maskCtx.restore();
    } else if (state.currentTool === 'pen-thick' || state.currentTool === 'pen-thin') {
      state.currentPath.push({ x: pos.x, y: pos.y });
      redrawMasks();
      maskCtx.save();
      maskCtx.strokeStyle = state.maskColor;
      maskCtx.lineWidth = getPenSize(state.currentTool);
      maskCtx.lineCap = 'round';
      maskCtx.lineJoin = 'round';
      maskCtx.beginPath();
      maskCtx.moveTo(state.currentPath[0].x, state.currentPath[0].y);
      for (let i = 1; i < state.currentPath.length; i++) {
        maskCtx.lineTo(state.currentPath[i].x, state.currentPath[i].y);
      }
      maskCtx.stroke();
      maskCtx.restore();
    } else if (state.currentTool === 'free') {
      state.currentPath.push({ x: pos.x, y: pos.y });
      redrawMasks();
      maskCtx.save();
      maskCtx.strokeStyle = state.maskColor;
      maskCtx.lineWidth = 2;
      maskCtx.setLineDash([5, 5]);
      maskCtx.beginPath();
      maskCtx.moveTo(state.currentPath[0].x, state.currentPath[0].y);
      for (let i = 1; i < state.currentPath.length; i++) {
        maskCtx.lineTo(state.currentPath[i].x, state.currentPath[i].y);
      }
      maskCtx.closePath();
      maskCtx.stroke();
      maskCtx.restore();
    }
  }

  function endDraw(e) {
    stopAutoScroll();

    // --- Select tool end ---
    if (state.currentTool === 'select') {
      handleSelectEnd();
      return;
    }

    if (!state.isDrawing) return;
    state.isDrawing = false;

    if (!state.masks[state.currentPage]) {
      state.masks[state.currentPage] = [];
    }

    if (state.currentTool === 'rect') {
      const pos = e.changedTouches ? getTouchEndPos(e) : getEndPos(e);
      const x = Math.min(state.startX, pos.x);
      const y = Math.min(state.startY, pos.y);
      const w = Math.abs(pos.x - state.startX);
      const h = Math.abs(pos.y - state.startY);
      if (w > 2 && h > 2) {
        state.masks[state.currentPage].push({
          type: 'rect',
          data: { x, y, w, h },
          color: state.maskColor,
        });
      }
    } else if (state.currentTool === 'pen-thick' || state.currentTool === 'pen-thin') {
      const penSize = getPenSize(state.currentTool);
      if (state.currentPath.length > 1) {
        state.masks[state.currentPage].push({
          type: 'pen',
          data: [...state.currentPath],
          color: state.maskColor,
          size: penSize,
        });
      }
    } else if (state.currentTool === 'free') {
      if (state.currentPath.length > 2) {
        state.masks[state.currentPage].push({
          type: 'free',
          data: [...state.currentPath],
          color: state.maskColor,
          size: 8,
        });
      }
    }

    state.currentPath = [];
    redrawMasks();

    // Auto-switch to select mode after drawing
    switchToSelect();
  }

  function switchToSelect() {
    document.querySelectorAll('.tool-btn').forEach((b) => b.classList.remove('active'));
    const selectBtn = document.querySelector('[data-tool="select"]');
    if (selectBtn) {
      selectBtn.classList.add('active');
      state.currentTool = 'select';
      $('#drawOptionsGroup').style.display = 'none';
      $('#selectInfoGroup').style.display = 'flex';
      maskCanvas.style.cursor = 'default';
    }
    // Select the last added mask
    const pageMasks = state.masks[state.currentPage] || [];
    if (pageMasks.length > 0) {
      selectMask(pageMasks.length - 1);
    }
  }

  function getEndPos(e) {
    const rect = maskCanvas.getBoundingClientRect();
    const scaleX = maskCanvas.width / rect.width;
    const scaleY = maskCanvas.height / rect.height;
    return {
      x: (e.clientX - rect.left) * scaleX,
      y: (e.clientY - rect.top) * scaleY,
    };
  }

  function getTouchEndPos(e) {
    const touch = e.changedTouches[0];
    const rect = maskCanvas.getBoundingClientRect();
    const scaleX = maskCanvas.width / rect.width;
    const scaleY = maskCanvas.height / rect.height;
    return {
      x: (touch.clientX - rect.left) * scaleX,
      y: (touch.clientY - rect.top) * scaleY,
    };
  }

  // =============================================
  // --- Select Tool: Move & Resize Masks ---
  // =============================================
  const HANDLE_SIZE = 8;

  function getMaskBounds(mask) {
    if (mask.type === 'rect') {
      return { x: mask.data.x, y: mask.data.y, w: mask.data.w, h: mask.data.h };
    } else if (mask.type === 'pen' || mask.type === 'free') {
      const pts = mask.data;
      let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      for (const p of pts) {
        if (p.x < minX) minX = p.x;
        if (p.y < minY) minY = p.y;
        if (p.x > maxX) maxX = p.x;
        if (p.y > maxY) maxY = p.y;
      }
      const pad = (mask.size || 0) / 2;
      return { x: minX - pad, y: minY - pad, w: maxX - minX + pad * 2, h: maxY - minY + pad * 2 };
    }
    return null;
  }

  function getHandlePositions(b) {
    return [
      { name: 'resize-nw', x: b.x, y: b.y },
      { name: 'resize-n',  x: b.x + b.w / 2, y: b.y },
      { name: 'resize-ne', x: b.x + b.w, y: b.y },
      { name: 'resize-w',  x: b.x, y: b.y + b.h / 2 },
      { name: 'resize-e',  x: b.x + b.w, y: b.y + b.h / 2 },
      { name: 'resize-sw', x: b.x, y: b.y + b.h },
      { name: 'resize-s',  x: b.x + b.w / 2, y: b.y + b.h },
      { name: 'resize-se', x: b.x + b.w, y: b.y + b.h },
    ];
  }

  function hitTestHandle(pos, bounds) {
    const hs = HANDLE_SIZE;
    const handles = getHandlePositions(bounds);
    for (const h of handles) {
      if (Math.abs(pos.x - h.x) < hs * 2 && Math.abs(pos.y - h.y) < hs * 2) {
        return h.name;
      }
    }
    return null;
  }

  function hitTestMask(pos, mask) {
    const b = getMaskBounds(mask);
    if (!b) return false;
    return pos.x >= b.x && pos.x <= b.x + b.w && pos.y >= b.y && pos.y <= b.y + b.h;
  }

  function deselectMask() {
    state.selectedMaskIndex = -1;
    state.selDragMode = null;
    $('#deleteSelBtn').style.display = 'none';
    redrawMasks();
  }

  function selectMask(idx) {
    state.selectedMaskIndex = idx;
    $('#deleteSelBtn').style.display = idx >= 0 ? 'inline-flex' : 'none';
    redrawMasks();
    drawSelectionUI();
  }

  function drawSelectionUI() {
    if (state.selectedMaskIndex < 0) return;
    const pageMasks = state.masks[state.currentPage] || [];
    const mask = pageMasks[state.selectedMaskIndex];
    if (!mask) return;

    const b = getMaskBounds(mask);
    if (!b) return;

    maskCtx.save();
    // Selection border
    maskCtx.strokeStyle = '#6366f1';
    maskCtx.lineWidth = 2;
    maskCtx.setLineDash([6, 3]);
    maskCtx.strokeRect(b.x, b.y, b.w, b.h);
    maskCtx.setLineDash([]);

    // Resize handles (8 points: corners + edge midpoints)
    const hs = HANDLE_SIZE;
    const handles = getHandlePositions(b);
    for (const h of handles) {
      maskCtx.fillStyle = '#ffffff';
      maskCtx.fillRect(h.x - hs / 2, h.y - hs / 2, hs, hs);
      maskCtx.strokeStyle = '#6366f1';
      maskCtx.lineWidth = 2;
      maskCtx.strokeRect(h.x - hs / 2, h.y - hs / 2, hs, hs);
    }
    maskCtx.restore();
  }

  function handleSelectStart(pos) {
    const pageMasks = state.masks[state.currentPage] || [];

    // If already selected, check handles first
    if (state.selectedMaskIndex >= 0) {
      const mask = pageMasks[state.selectedMaskIndex];
      if (mask) {
        const b = getMaskBounds(mask);
        const handle = hitTestHandle(pos, b);
        if (handle && mask.type === 'rect') {
          state.selDragMode = handle;
          state.selDragStart = { x: pos.x, y: pos.y };
          state.selOrigData = JSON.parse(JSON.stringify(mask.data));
          return;
        }
        // Check if clicking inside selected mask to move
        if (hitTestMask(pos, mask)) {
          state.selDragMode = 'move';
          state.selDragStart = { x: pos.x, y: pos.y };
          state.selOrigData = JSON.parse(JSON.stringify(mask.data));
          return;
        }
      }
    }

    // Try to select a mask (reverse order = topmost first)
    for (let i = pageMasks.length - 1; i >= 0; i--) {
      if (hitTestMask(pos, pageMasks[i])) {
        selectMask(i);
        state.selDragMode = 'move';
        state.selDragStart = { x: pos.x, y: pos.y };
        state.selOrigData = JSON.parse(JSON.stringify(pageMasks[i].data));
        return;
      }
    }

    // Clicked empty area
    deselectMask();
  }

  function handleSelectDrag(pos) {
    if (!state.selDragMode || state.selectedMaskIndex < 0) return;
    const pageMasks = state.masks[state.currentPage] || [];
    const mask = pageMasks[state.selectedMaskIndex];
    if (!mask) return;

    const dx = pos.x - state.selDragStart.x;
    const dy = pos.y - state.selDragStart.y;
    const orig = state.selOrigData;

    if (state.selDragMode === 'move') {
      if (mask.type === 'rect') {
        mask.data.x = orig.x + dx;
        mask.data.y = orig.y + dy;
      } else if (Array.isArray(orig)) {
        // pen / free — move all points
        for (let i = 0; i < orig.length; i++) {
          mask.data[i].x = orig[i].x + dx;
          mask.data[i].y = orig[i].y + dy;
        }
      }
    } else if (mask.type === 'rect' && state.selDragMode.startsWith('resize-')) {
      const dir = state.selDragMode.replace('resize-', '');
      let nx = orig.x, ny = orig.y, nw = orig.w, nh = orig.h;

      // Horizontal
      if (dir.includes('w')) { nx = orig.x + dx; nw = orig.w - dx; }
      if (dir.includes('e')) { nw = orig.w + dx; }

      // Vertical
      if (dir.includes('n')) { ny = orig.y + dy; nh = orig.h - dy; }
      if (dir.includes('s')) { nh = orig.h + dy; }

      // Enforce minimum size
      if (nw < 10) { nw = 10; }
      if (nh < 10) { nh = 10; }
      mask.data.x = nx;
      mask.data.y = ny;
      mask.data.w = nw;
      mask.data.h = nh;
    }

    redrawMasks();
    drawSelectionUI();
  }

  function handleSelectEnd() {
    state.selDragMode = null;
    state.selDragStart = null;
    state.selOrigData = null;
  }

  // Delete selected mask
  $('#deleteSelBtn').addEventListener('click', () => {
    if (state.selectedMaskIndex >= 0) {
      const pageMasks = state.masks[state.currentPage];
      if (pageMasks) {
        pageMasks.splice(state.selectedMaskIndex, 1);
      }
      deselectMask();
    }
  });

  // --- Undo ---
  $('#undoBtn').addEventListener('click', () => {
    const pageMasks = state.masks[state.currentPage];
    if (pageMasks && pageMasks.length > 0) {
      pageMasks.pop();
      redrawMasks();
    }
  });

  // --- Clear Page ---
  $('#clearPageBtn').addEventListener('click', () => {
    state.masks[state.currentPage] = [];
    deselectMask();
  });

  // --- Full Page Mask ---
  $('#clearAllBtn').addEventListener('click', () => {
    if (!state.masks[state.currentPage]) {
      state.masks[state.currentPage] = [];
    }
    // Add a rect covering the entire canvas
    state.masks[state.currentPage].push({
      type: 'rect',
      data: { x: 0, y: 0, w: maskCanvas.width, h: maskCanvas.height },
      color: state.maskColor,
    });
    // Switch to select mode with the new mask selected
    switchToSelect();
  });

  // --- Page Navigation ---
  $('#prevPage').addEventListener('click', () => {
    if (state.currentPage > 1) {
      state.currentPage--;
      renderPage();
    }
  });

  $('#nextPage').addEventListener('click', () => {
    if (state.currentPage < state.totalPages) {
      state.currentPage++;
      renderPage();
    }
  });

  // --- Zoom ---
  function updateZoomLabel() {
    $('#zoomLevel').textContent = Math.round(state.scale * 100) + '%';
  }

  $('#zoomIn').addEventListener('click', () => {
    state.scale = Math.min(state.scale + 0.25, 4);
    renderPage();
  });

  $('#zoomOut').addEventListener('click', () => {
    state.scale = Math.max(state.scale - 0.25, 0.5);
    renderPage();
  });

  $('#zoomFit').addEventListener('click', cycleFit);

  // --- Export PDF with Masks ---
  $('#exportBtn').addEventListener('click', exportPDF);

  async function exportPDF() {
    showLoading('マスク済みPDFを生成中（フラット化処理）...');
    try {
      const { PDFDocument } = PDFLib;
      // Create a brand new PDF — no original text/data carried over
      const newPdf = await PDFDocument.create();

      // High-resolution scale for export (2x for sharp output)
      const exportScale = 2;

      for (let pageNum = 1; pageNum <= state.totalPages; pageNum++) {
        loadingText.textContent =
          `フラット化中... ${pageNum} / ${state.totalPages} ページ`;

        const pdfPage = await state.pdfDoc.getPage(pageNum);
        const viewport = pdfPage.getViewport({ scale: exportScale });

        // 1. Render PDF page to canvas
        const pageCanvas = document.createElement('canvas');
        pageCanvas.width = viewport.width;
        pageCanvas.height = viewport.height;
        const pageCtx = pageCanvas.getContext('2d');
        await pdfPage.render({ canvasContext: pageCtx, viewport }).promise;

        // 2. Draw masks on top of the rendered page
        const pageMasks = state.masks[pageNum] || [];
        if (pageMasks.length > 0) {
          // Scale masks from display scale to export scale
          const scaleFactor = exportScale / state.scale;
          for (const mask of pageMasks) {
            drawMaskToContext(pageCtx, mask, scaleFactor);
          }
        }

        // 3. Convert flattened canvas to JPEG image
        const jpegDataUrl = pageCanvas.toDataURL('image/jpeg', 0.92);
        const jpegBytes = dataURLtoBytes(jpegDataUrl);
        const img = await newPdf.embedJpg(jpegBytes);

        // 4. Add page with same dimensions as original
        const origViewport = pdfPage.getViewport({ scale: 1 });
        const page = newPdf.addPage([origViewport.width, origViewport.height]);
        page.drawImage(img, {
          x: 0,
          y: 0,
          width: origViewport.width,
          height: origViewport.height,
        });
      }

      const resultBytes = await newPdf.save();
      downloadBlob(resultBytes, state.pdfFileName + '_masked.pdf', 'application/pdf');
    } catch (err) {
      alert('PDF出力に失敗しました: ' + err.message);
      console.error(err);
    } finally {
      hideLoading();
    }
  }

  // Draw a mask onto a canvas context with a scale factor
  function drawMaskToContext(ctx, mask, scaleFactor) {
    ctx.save();
    if (mask.type === 'rect') {
      ctx.fillStyle = mask.color;
      ctx.fillRect(
        mask.data.x * scaleFactor,
        mask.data.y * scaleFactor,
        mask.data.w * scaleFactor,
        mask.data.h * scaleFactor,
      );
    } else if (mask.type === 'pen') {
      ctx.strokeStyle = mask.color;
      ctx.lineWidth = mask.size * scaleFactor;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.beginPath();
      const pts = mask.data;
      if (pts.length > 0) {
        ctx.moveTo(pts[0].x * scaleFactor, pts[0].y * scaleFactor);
        for (let i = 1; i < pts.length; i++) {
          ctx.lineTo(pts[i].x * scaleFactor, pts[i].y * scaleFactor);
        }
      }
      ctx.stroke();
    } else if (mask.type === 'free') {
      const pts = mask.data;
      if (pts.length < 3) { ctx.restore(); return; }
      ctx.fillStyle = mask.color;
      ctx.beginPath();
      ctx.moveTo(pts[0].x * scaleFactor, pts[0].y * scaleFactor);
      for (let i = 1; i < pts.length; i++) {
        ctx.lineTo(pts[i].x * scaleFactor, pts[i].y * scaleFactor);
      }
      ctx.closePath();
      ctx.fill();

      ctx.strokeStyle = mask.color;
      ctx.lineWidth = mask.size * scaleFactor;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.beginPath();
      ctx.moveTo(pts[0].x * scaleFactor, pts[0].y * scaleFactor);
      for (let i = 1; i < pts.length; i++) {
        ctx.lineTo(pts[i].x * scaleFactor, pts[i].y * scaleFactor);
      }
      ctx.closePath();
      ctx.stroke();
    }
    ctx.restore();
  }

  function dataURLtoBytes(dataURL) {
    const base64 = dataURL.split(',')[1];
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  function downloadBlob(bytes, filename, mime) {
    const blob = new Blob([bytes], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  }

  // --- Loading ---
  function showLoading(text) {
    loadingText.textContent = text;
    loadingOverlay.style.display = 'flex';
  }

  function hideLoading() {
    loadingOverlay.style.display = 'none';
  }

  // --- Keyboard shortcuts ---
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === 'z') {
      e.preventDefault();
      $('#undoBtn').click();
    }
    if (e.key === 'Delete' && state.selectedMaskIndex >= 0) {
      e.preventDefault();
      $('#deleteSelBtn').click();
    }
  });
})();
