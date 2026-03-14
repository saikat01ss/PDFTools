// ──────────────────────────────────────────────
// PDF Organizer – script.js
// ──────────────────────────────────────────────

// Setup PDF.js Worker
pdfjsLib.GlobalWorkerOptions.workerSrc = './js/pdf.worker.min.js';

// ── State ──────────────────────────────────────
let uploadedFiles = {};   // fileId -> { name, bytes, label }   label = 'A','B','C'…
let pagesState = [];   // current visual order of pages
let fileIdCounter = 0;
let pageIdCounter = 0;

// Zoom state: number of grid columns (1-8)
const BASE_HEIGHT = 256;   // px
const BASE_GRID_COLUMN = Math.max(1, Math.round((window.innerWidth - document.querySelector('.sidebar').offsetWidth - 136) / BASE_HEIGHT));
let gridColumns = BASE_GRID_COLUMN;
const MIN_COLS = 1;
const MAX_COLS = 8;

// thumbnailCache[pageId] = { scale, dataURL }
let thumbnailCache = {};

// Selection state
let selectedPageIds = new Set();
let lastSelectedIndex = -1;
// Tracks the ORDER in which pages were selected (for custom range field)
let selectionOrder = [];

// Guard flag: true while programmatically updating the custom-range field
// from the selection – prevents the field's input listener from looping back.
let _updatingFieldFromSelection = false;

// ── File label helpers ─────────────────────────
// Each file gets a letter label: 1st file → 'A', 2nd → 'B', … 'Z', then 'AA', 'AB'…
function fileIndexToLabel(idx) {
    let label = '';
    idx++; // 1-based
    while (idx > 0) {
        idx--;
        label = String.fromCharCode(65 + (idx % 26)) + label;
        idx = Math.floor(idx / 26);
    }
    return label;
}

// Returns the label for a given fileId (e.g. 'A', 'B')
function getFileLabel(fileId) {
    return uploadedFiles[fileId] ? uploadedFiles[fileId].label : '?';
}

// Returns the page label string for a page object (e.g. 'A3', 'B1')
function getPageLabel(page) {
    if (page.isBlank) return '—';
    return getFileLabel(page.fileId) + page.originalPageNum;
}

// ── Undo / Redo stack ──────────────────────────
let undoStack = [];
let redoStack = [];

function snapshot() {
    return pagesState.map(p => ({ ...p }));
}

function saveUndo() {
    undoStack.push(snapshot());
    redoStack = [];
    if (undoStack.length > 100) undoStack.shift();
}

function undo() {
    if (undoStack.length === 0) return;
    redoStack.push(snapshot());
    pagesState = undoStack.pop();
    const validIds = new Set(pagesState.map(p => p.id));
    selectedPageIds = new Set([...selectedPageIds].filter(id => validIds.has(id)));
    lastSelectedIndex = -1;
    updateUI();
}

function redo() {
    if (redoStack.length === 0) return;
    undoStack.push(snapshot());
    pagesState = redoStack.pop();
    const validIds = new Set(pagesState.map(p => p.id));
    selectedPageIds = new Set([...selectedPageIds].filter(id => validIds.has(id)));
    lastSelectedIndex = -1;
    updateUI();
}

// Keyboard shortcuts
document.addEventListener('keydown', e => {
    const tag = document.activeElement.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;
    if ((e.ctrlKey || e.metaKey) && e.key === 'z') { e.preventDefault(); undo(); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'y') { e.preventDefault(); redo(); }
    // Instruction 5: Ctrl+A selects all pages
    if ((e.ctrlKey || e.metaKey) && e.key === 'a') { e.preventDefault(); selectAllPages(); }
    if (e.key === 'Escape') closePagePreview();
});

// ── UI Elements ────────────────────────────────
const grid = document.getElementById('page-grid');
const emptyState = document.getElementById('empty-state');
const fileListElement = document.getElementById('file-list');
const badge = document.getElementById('file-count-badge');
const customInput = document.getElementById('custom-range-input');
const extractionFilename = document.getElementById('extraction-filename');

// Toggle extraction input (manual radio change by user)
document.querySelectorAll('input[name="extract-mode"]').forEach(radio => {
    radio.addEventListener('change', e => {
        if (e.target.value === 'all') {
            // Instruction 2: switching to "extract all" clears the field and deselects all pages
            customInput.disabled = true;
            customInput.value = '';
            if (!_updatingFieldFromSelection) {
                selectedPageIds.clear();
                selectionOrder = [];
                lastSelectedIndex = -1;
                updateUI();
            }
        } else {
            customInput.disabled = false;
            customInput.focus();
        }
    });
});

// Feature 3: custom range input → page selection sync
// When the user types in the field, parse labels and update selectedPageIds.
customInput.addEventListener('input', () => {
    if (_updatingFieldFromSelection) return; // avoid feedback loop
    const pages = parseCustomRangeByLabel(customInput.value);
    selectedPageIds = new Set(pages.map(p => p.id));
    lastSelectedIndex = pages.length > 0
        ? pagesState.findIndex(p => p.id === pages[pages.length - 1].id)
        : -1;
    // Rebuild UI without updating the field again
    _updatingFieldFromSelection = true;
    updateUI();
    _updatingFieldFromSelection = false;
});

// Feature 4: Scroll wheel scrolls the grid; only allow Ctrl+wheel for browser
// zoom (which the user can suppress via browser settings). We prevent the
// default only when Ctrl is NOT held so native grid scrolling works normally.
document.getElementById('main-area').addEventListener('wheel', e => {
    if (!e.ctrlKey) {
        // Let the browser handle natural scrolling – don't interfere.
        return;
    }
    // Ctrl+wheel: prevent browser page zoom; map to our zoom buttons instead.
    e.preventDefault();
    if (e.deltaY < 0) zoomIn(); else zoomOut();
}, { passive: false });

// ── Zoom ───────────────────────────────────────
function cardHeight() {
    return Math.round((window.innerWidth - document.querySelector('.sidebar').offsetWidth - 136 - (gridColumns - 1) * 16) / gridColumns);
}

function desiredRenderScale() {
    const ratio = cardHeight() / BASE_HEIGHT;
    return Math.min(2.0, Math.max(0.3, 0.5 * ratio));
}

function applyGridStyle() {
    grid.style.gridTemplateColumns = `repeat(${gridColumns}, minmax(0, 1fr))`;
    document.querySelectorAll('.page-card').forEach(card => {
        card.style.height = cardHeight() + 'px';
    });
}

function zoomIn() {
    if (gridColumns <= MIN_COLS) return;
    gridColumns--;
    applyGridStyle();
    rerenderVisible();
}

function zoomOut() {
    if (gridColumns >= MAX_COLS) return;
    gridColumns++;
    applyGridStyle();
    rerenderVisible();
}

async function rerenderVisible() {
    const scale = desiredRenderScale();
    document.body.classList.add('loading');
    const promises = pagesState.map(async (page) => {
        if (page.isBlank) return;
        const cached = thumbnailCache[page.id];
        if (cached && cached.scale >= scale - 0.01) return;
        await renderPageThumbnail(page, scale);
    });
    await Promise.all(promises);
    document.body.classList.remove('loading');
    updateUI();
}

// ── Thumbnail rendering ────────────────────────
async function renderPageThumbnail(pageObj, scale) {
    if (!uploadedFiles[pageObj.fileId]) return;
    try {
        const bytes = uploadedFiles[pageObj.fileId].bytes.slice(0);
        const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
        const page = await pdf.getPage(pageObj.originalPageNum);
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
        const dataURL = canvas.toDataURL();
        thumbnailCache[pageObj.id] = { scale, dataURL };
        pageObj.thumbnail = dataURL;
    } catch (e) {
        console.warn('Re-render failed for', pageObj.id, e);
    }
}

// ── Blank page utility ─────────────────────────
function createBlankPageThumbnail() {
    const canvas = document.createElement('canvas');
    canvas.width = 595;
    canvas.height = 842;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = 'white';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.strokeStyle = '#e2e8f0';
    ctx.lineWidth = 1;
    ctx.strokeRect(0, 0, canvas.width, canvas.height);
    return canvas.toDataURL();
}

// ── File upload ────────────────────────────────
async function handleFileUpload(event) {
    const files = event.target.files;
    if (files.length === 0) return;
    document.body.classList.add('loading');
    const scale = desiredRenderScale();
    try {
        for (const file of files) {
            const arrayBuffer = await file.arrayBuffer();
            const fileId = 'file_' + fileIdCounter++;
            // Assign label based on how many files we have so far
            const fileIndex = Object.keys(uploadedFiles).length;
            const fileLabel = fileIndexToLabel(fileIndex);
            uploadedFiles[fileId] = { name: file.name, bytes: arrayBuffer.slice(0), label: fileLabel };

            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const viewport = page.getViewport({ scale });
                const canvas = document.createElement('canvas');
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
                const dataURL = canvas.toDataURL();
                const pageId = 'page_' + pageIdCounter++;
                thumbnailCache[pageId] = { scale, dataURL };
                pagesState.push({
                    id: pageId,
                    fileId,
                    originalPageNum: i,
                    rotation: 0,
                    isBlank: false,
                    thumbnail: dataURL
                });
            }
        }
        updateUI();
    } catch (err) {
        console.error('Error parsing PDF:', err);
        alert('Failed to read the PDF. It might be corrupted or protected.');
    }
    document.body.classList.remove('loading');
    event.target.value = '';
}

// ── Main UI renderer ───────────────────────────
function updateUI() {
    // Preserve scroll position so selecting a page doesn't jump to top
    const mainArea = document.getElementById('main-area');
    const savedScrollTop = mainArea ? mainArea.scrollTop : 0;

    // Empty state
    if (pagesState.length > 0) {
        emptyState.classList.add('hidden');
        grid.classList.remove('hidden');
    } else {
        emptyState.classList.remove('hidden');
        grid.classList.add('hidden');
    }

    // File list sidebar
    const fileKeys = Object.keys(uploadedFiles);
    if (fileKeys.length > 0) {
        badge.classList.remove('hidden');
        badge.innerText = fileKeys.length;
        extractionFilename.innerText = `(${uploadedFiles[fileKeys[0]].name})`;
        fileListElement.innerHTML = fileKeys.map(key => `
            <div class="file-item">
                <span class="file-item-name" title="${uploadedFiles[key].name}">
                    <i class="fas fa-file-pdf file-item-icon"></i>${uploadedFiles[key].label}: ${uploadedFiles[key].name}
                </span>
                <button onclick="deleteFile('${key}')" class="file-delete-btn" title="Remove this file">
                    <i class="fas fa-trash-alt"></i>
                </button>
            </div>
        `).join('');
    } else {
        badge.classList.add('hidden');
        fileListElement.innerHTML = 'No files selected.';
        extractionFilename.innerText = '';
    }

    // Selection bar
    updateSelectionBar();

    // When multiple pages are selected, per-card action buttons should only
    // show on hover (not always-visible). Use a data attribute for CSS targeting.
    const multiSelected = selectedPageIds.size > 1;

    // Page grid
    grid.innerHTML = '';
    grid.style.gridTemplateColumns = `repeat(${gridColumns}, minmax(0, 1fr))`;
    const h = cardHeight();

    pagesState.forEach((page, index) => {
        const isSelected = selectedPageIds.has(page.id);
        const pageLabel = getPageLabel(page);
        const pageEl = document.createElement('div');
        // Add 'multi-selected' class when multiple are selected so CSS hides
        // the per-card actions unless hovered
        pageEl.className = 'relative group'
            + (isSelected ? ' page-selected' : '')
            + (isSelected && multiSelected ? ' multi-selected' : '');
        pageEl.id = page.id;
        pageEl.dataset.index = index;

        pageEl.innerHTML = `
            <div draggable="true" class="page-card${isSelected ? ' page-card--selected' : ''}"
                 style="height:${h}px;">

                <!-- Top-right: rotate CCW + rotate CW + delete (instruction 3: CCW left of CW) -->
                <div class="page-actions">
                    <button data-action="rotateCCW" data-id="${page.id}"
                            class="page-action-btn rotate-ccw" title="Rotate anti-clockwise">
                        <i class="fas fa-rotate-left"></i>
                    </button>
                    <button data-action="rotate" data-id="${page.id}"
                            class="page-action-btn rotate" title="Rotate clockwise">
                        <i class="fas fa-rotate-right"></i>
                    </button>
                    <button data-action="delete" data-id="${page.id}"
                            class="page-action-btn del" title="Delete">
                        <i class="fas fa-times"></i>
                    </button>
                </div>

                <!-- Bottom-left: copy (instruction 4: stays bottom-left) -->
                <button data-action="copy" data-id="${page.id}" data-index="${index}"
                        class="copy-page-btn" title="Duplicate page">
                    <i class="fas fa-copy"></i>
                </button>

                <!-- Bottom-right: full page preview (instruction 4: moved here) -->
                <button data-action="previewPage" data-id="${page.id}"
                        class="preview-page-btn" title="Full-screen preview">
                    <i class="fas fa-expand"></i>
                </button>

                <!-- Insert blank before/after -->
                <button data-action="insertBefore" data-index="${index}"
                        class="insert-btn left" title="Insert blank page before">
                    <i class="fas fa-plus"></i>
                </button>
                <button data-action="insertAfter" data-index="${index}"
                        class="insert-btn right" title="Insert blank page after">
                    <i class="fas fa-plus"></i>
                </button>

                <!-- Top-left: selection circle -->
                <div class="select-indicator${isSelected ? ' select-indicator--on' : ''}"
                     data-action="toggleSelect" data-id="${page.id}" data-index="${index}">
                    <i class="fas fa-check"></i>
                </div>

                <div class="thumb-area">
                    <img src="${page.thumbnail}"
                         class="page-thumbnail"
                         style="transform:rotate(${page.rotation}deg);">
                </div>

                <span class="page-num">${pageLabel}</span>
            </div>
        `;

        grid.appendChild(pageEl);

        // ── Event delegation on the wrapper ──
        pageEl.addEventListener('click', e => {
            const btn = e.target.closest('[data-action]');
            if (!btn) {
                handlePageClick(page.id, index, e);
                return;
            }
            const action = btn.dataset.action;
            const id = btn.dataset.id;
            const idx = parseInt(btn.dataset.index, 10);

            switch (action) {
                case 'rotate':
                    rotatePage(id, 1); break;
                case 'rotateCCW':
                    rotatePage(id, -1); break;
                case 'previewPage':
                    openPagePreview(page); break;
                case 'delete':
                    deletePage(id); break;
                case 'copy':
                    copyPage(id, idx); break;
                case 'insertBefore':
                    insertBlankPage(idx); break;
                case 'insertAfter':
                    insertBlankPage(idx + 1); break;
                case 'toggleSelect':
                    toggleSelectPage(id, idx); break;
            }
        });

        addDnDHandlers(pageEl);
    });

    // Feature 1: Ghost end-drop zone (always present when there are pages)
    if (pagesState.length > 0) {
        const ghost = document.createElement('div');
        ghost.id = 'ghost-drop-zone';
        ghost.className = 'ghost-drop-zone';
        ghost.style.height = h + 'px';
        ghost.addEventListener('dragover', e => {
            e.preventDefault();
            ghost.classList.add('drag-over');
            e.dataTransfer.dropEffect = 'move';
        });
        ghost.addEventListener('dragleave', () => ghost.classList.remove('drag-over'));
        ghost.addEventListener('drop', e => {
            e.preventDefault();
            ghost.classList.remove('drag-over');
            if (dragSrcEl === null) return;
            saveUndo();
            if (isDraggingMultiple) {
                const selected = pagesState.filter(p => selectedPageIds.has(p.id));
                const remaining = pagesState.filter(p => !selectedPageIds.has(p.id));
                pagesState = [...remaining, ...selected];
            } else {
                const draggingId = e.dataTransfer.getData('text/plain');
                const fromIndex = pagesState.findIndex(p => p.id === draggingId);
                if (fromIndex !== -1) {
                    const [moved] = pagesState.splice(fromIndex, 1);
                    pagesState.push(moved);
                }
            }
            if (dragSelectionSnapshot) selectedPageIds = new Set(dragSelectionSnapshot);
            dragSelectionSnapshot = null;
            updateUI();
        });
        grid.appendChild(ghost);
    }

    if (pagesState.length === 0) {
        selectedPageIds.clear();
        lastSelectedIndex = -1;
    }

    // Feature 2: Sync selection → custom range field
    syncCustomRangeToSelection();

    // Restore scroll position after DOM rebuild
    if (mainArea) mainArea.scrollTop = savedScrollTop;
}

// ── Selection helpers ──────────────────────────

/**
 * Build a compact range string from selected pages IN THE ORDER THEY WERE SELECTED.
 * Instruction 1: respects insertion order, e.g. selecting A12 then A2-A7 then A1
 * produces "A12, A2-A7, A1".
 * Consecutive pages (same file, sequential numbers) are compressed to ranges.
 */
function buildSelectionRangeString() {
    // Follow selectionOrder (insertion order) rather than pagesState order
    const orderedIds = selectionOrder.filter(id => selectedPageIds.has(id));
    const idToPage = {};
    pagesState.forEach(p => { idToPage[p.id] = p; });
    const selected = orderedIds.map(id => idToPage[id]).filter(p => p && !p.isBlank);
    if (selected.length === 0) return '';

    const parts = [];
    let rangeStart = selected[0];
    let rangePrev  = selected[0];

    const flushRange = (start, end) => {
        const sl = getPageLabel(start);
        const el = getPageLabel(end);
        const sm = sl.match(/^([A-Z]+)(\d+)$/);
        const em = el.match(/^([A-Z]+)(\d+)$/);
        if (sm && em && sm[1] === em[1] && parseInt(em[2]) - parseInt(sm[2]) > 0) {
            parts.push(sl + '-' + el);
        } else {
            parts.push(sl);
            if (sl !== el) parts.push(el); // shouldn't happen but safety net
        }
    };

    for (let i = 1; i < selected.length; i++) {
        const cur = selected[i];
        const prevLabel = getPageLabel(rangePrev);
        const curLabel  = getPageLabel(cur);
        const pm = prevLabel.match(/^([A-Z]+)(\d+)$/);
        const cm = curLabel.match(/^([A-Z]+)(\d+)$/);
        const consecutive = pm && cm
            && pm[1] === cm[1]
            && parseInt(cm[2]) === parseInt(pm[2]) + 1;
        if (consecutive) {
            rangePrev = cur;
        } else {
            flushRange(rangeStart, rangePrev);
            rangeStart = cur;
            rangePrev  = cur;
        }
    }
    flushRange(rangeStart, rangePrev);
    return parts.join(', ');
}

/**
 * Feature 2: Keep the custom-range field in sync with the current selection.
 * Called at the end of updateUI().
 */
function syncCustomRangeToSelection() {
    if (_updatingFieldFromSelection) return; // guard against feedback
    _updatingFieldFromSelection = true;

    const customRadio = document.getElementById('custom-range-radio');
    const allRadio    = document.querySelector('input[name="extract-mode"][value="all"]');

    if (selectedPageIds.size > 0) {
        // Switch to custom mode and populate field
        if (customRadio) customRadio.checked = true;
        customInput.disabled = false;
        customInput.value = buildSelectionRangeString();
    } else {
        // Switch back to 'all' mode and clear field
        if (allRadio) allRadio.checked = true;
        customInput.disabled = true;
        customInput.value = '';
    }

    _updatingFieldFromSelection = false;
}

function toggleSelectPage(pageId, index) {
    if (selectedPageIds.has(pageId)) {
        selectedPageIds.delete(pageId);
        selectionOrder = selectionOrder.filter(id => id !== pageId);
        if (selectedPageIds.size === 0) lastSelectedIndex = -1;
    } else {
        selectedPageIds.add(pageId);
        selectionOrder.push(pageId);
        lastSelectedIndex = index;
    }
    updateUI();
}

function handlePageClick(pageId, index, event) {
    if (event.shiftKey && lastSelectedIndex !== -1) {
        const start = Math.min(lastSelectedIndex, index);
        const end = Math.max(lastSelectedIndex, index);
        for (let i = start; i <= end; i++) {
            const id = pagesState[i].id;
            if (!selectedPageIds.has(id)) {
                selectedPageIds.add(id);
                selectionOrder.push(id);
            }
        }
    } else if (event.ctrlKey || event.metaKey) {
        if (selectedPageIds.has(pageId)) {
            selectedPageIds.delete(pageId);
            selectionOrder = selectionOrder.filter(id => id !== pageId);
        } else {
            selectedPageIds.add(pageId);
            selectionOrder.push(pageId);
            lastSelectedIndex = index;
        }
    } else {
        if (selectedPageIds.has(pageId) && selectedPageIds.size === 1) {
            selectedPageIds.clear();
            selectionOrder = [];
            lastSelectedIndex = -1;
        } else {
            selectedPageIds.clear();
            selectionOrder = [];
            selectedPageIds.add(pageId);
            selectionOrder.push(pageId);
            lastSelectedIndex = index;
        }
    }
    updateUI();
}

function updateSelectionBar() {
    const bar = document.getElementById('selection-bar');
    if (!bar) return;
    if (selectedPageIds.size > 0) {
        bar.classList.remove('hidden');
        document.getElementById('selection-count').textContent =
            `${selectedPageIds.size} page${selectedPageIds.size > 1 ? 's' : ''} selected`;
    } else {
        bar.classList.add('hidden');
    }
    // Instruction 5: keep select-all checkbox state in sync
    const selectAllCb = document.getElementById('select-all-cb');
    if (selectAllCb) {
        const allCount = pagesState.length;
        selectAllCb.checked = allCount > 0 && selectedPageIds.size === allCount;
        selectAllCb.indeterminate = selectedPageIds.size > 0 && selectedPageIds.size < allCount;
    }
}

// direction: 1 = clockwise, -1 = anti-clockwise
function rotateSelected(direction = 1) {
    saveUndo();
    const deg = direction === -1 ? 270 : 90;
    for (const id of selectedPageIds) {
        const p = pagesState.find(p => p.id === id);
        if (p) p.rotation = (p.rotation + deg) % 360;
    }
    // Keep selection intact
    updateUI();
}

function deleteSelected() {
    saveUndo();
    pagesState = pagesState.filter(p => !selectedPageIds.has(p.id));
    selectedPageIds.clear();
    selectionOrder = [];
    lastSelectedIndex = -1;
    updateUI();
}

function clearSelection() {
    selectedPageIds.clear();
    selectionOrder = [];
    lastSelectedIndex = -1;
    updateUI();
}

// Instruction 5: select all pages
function selectAllPages() {
    pagesState.forEach((p, i) => {
        if (!selectedPageIds.has(p.id)) {
            selectedPageIds.add(p.id);
            selectionOrder.push(p.id);
        }
    });
    lastSelectedIndex = pagesState.length - 1;
    updateUI();
}

// ── Drag & Drop ────────────────────────────────
let dragSrcEl = null;
let isDraggingMultiple = false;
// Snapshot of selection before drag (to restore after)
let dragSelectionSnapshot = null;

function addDnDHandlers(el) {
    el.addEventListener('dragstart', handleDragStart, false);
    el.addEventListener('dragover', handleDragOver, false);
    el.addEventListener('dragleave', handleDragLeave, false);
    el.addEventListener('drop', handleDrop, false);
    el.addEventListener('dragend', handleDragEnd, false);
}

function handleDragStart(e) {
    dragSrcEl = this;
    this.style.opacity = '0.5';
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', this.id);
    isDraggingMultiple = selectedPageIds.has(this.id) && selectedPageIds.size > 1;
    // Snapshot the selection so we can restore it after drop
    dragSelectionSnapshot = new Set(selectedPageIds);
}

function handleDragOver(e) {
    e.preventDefault();
    this.classList.add('drag-over');
    e.dataTransfer.dropEffect = 'move';
    return false;
}

function handleDragLeave() {
    this.classList.remove('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    this.classList.remove('drag-over');

    if (dragSrcEl === null || dragSrcEl === this) return false;

    const draggingId = e.dataTransfer.getData('text/plain');
    const droppedOnId = this.id;

    saveUndo();

    if (isDraggingMultiple) {
        const selected = pagesState.filter(p => selectedPageIds.has(p.id));
        let remaining = pagesState.filter(p => !selectedPageIds.has(p.id));
        const insertAt = remaining.findIndex(p => p.id === droppedOnId);
        if (insertAt === -1) {
            pagesState = [...remaining, ...selected];
        } else {
            remaining.splice(insertAt, 0, ...selected);
            pagesState = remaining;
        }
    } else {
        const fromIndex = pagesState.findIndex(p => p.id === draggingId);
        const toIndex = pagesState.findIndex(p => p.id === droppedOnId);
        if (fromIndex !== -1 && toIndex !== -1 && fromIndex !== toIndex) {
            const [moved] = pagesState.splice(fromIndex, 1);
            const newTo = pagesState.findIndex(p => p.id === droppedOnId);
            pagesState.splice(newTo !== -1 ? newTo : pagesState.length, 0, moved);
        }
    }

    // Restore selection (do NOT clear after drag-drop)
    if (dragSelectionSnapshot) {
        selectedPageIds = new Set(dragSelectionSnapshot);
    }
    dragSelectionSnapshot = null;
    updateUI();
    return false;
}

function handleDragEnd() {
    if (dragSrcEl) dragSrcEl.style.opacity = '';
    dragSrcEl = null;
    document.querySelectorAll('#page-grid > div').forEach(el => {
        el.classList.remove('drag-over');
        el.style.opacity = '';
    });
}

// ── Page CRUD ──────────────────────────────────
// direction: 1 = clockwise, -1 = anti-clockwise
function rotatePage(id, direction = 1) {
    saveUndo();
    const p = pagesState.find(p => p.id === id);
    if (p) {
        const deg = direction === -1 ? 270 : 90;
        p.rotation = (p.rotation + deg) % 360;
        // Keep selection intact - do not touch selectedPageIds
        updateUI();
    }
}

function deletePage(id) {
    saveUndo();
    pagesState = pagesState.filter(p => p.id !== id);
    selectedPageIds.delete(id);
    selectionOrder = selectionOrder.filter(sid => sid !== id);
    updateUI();
}

// ── Non-blocking inline confirmation ──────────
function showConfirm(message, onYes) {
    const existing = document.getElementById('confirm-overlay');
    if (existing) existing.remove();

    const overlay = document.createElement('div');
    overlay.id = 'confirm-overlay';
    overlay.style.cssText = [
        'position:fixed', 'inset:0', 'z-index:200',
        'display:flex', 'align-items:center', 'justify-content:center',
        'background:rgba(0,0,0,0.35)'
    ].join(';');

    overlay.innerHTML = `
        <div style="background:#fff;border-radius:12px;padding:1.5rem 2rem;
                    box-shadow:0 20px 40px rgba(0,0,0,0.25);max-width:360px;
                    width:90%;text-align:center;">
            <p style="margin-bottom:1.25rem;font-size:0.95rem;color:#1a202c;line-height:1.5;">
                ${message}
            </p>
            <div style="display:flex;gap:0.75rem;justify-content:center;">
                <button id="confirm-yes"
                    style="background:#e53e3e;color:#fff;border:none;padding:0.5rem 1.25rem;
                           border-radius:6px;cursor:pointer;font-weight:600;font-size:0.9rem;">
                    Yes, remove
                </button>
                <button id="confirm-no"
                    style="background:#f3f4f6;color:#374151;border:none;padding:0.5rem 1.25rem;
                           border-radius:6px;cursor:pointer;font-size:0.9rem;">
                    Cancel
                </button>
            </div>
        </div>
    `;

    document.body.appendChild(overlay);
    overlay.querySelector('#confirm-yes').onclick = () => { overlay.remove(); onYes(); };
    overlay.querySelector('#confirm-no').onclick = () => overlay.remove();
    overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
}

function deleteFile(fileId) {
    showConfirm(`Remove "${uploadedFiles[fileId]?.name}" and all its pages?`, () => {
        saveUndo();
        delete uploadedFiles[fileId];
        const removed = pagesState.filter(p => p.fileId === fileId).map(p => p.id);
        const removedSet = new Set(removed);
        removed.forEach(id => { selectedPageIds.delete(id); delete thumbnailCache[id]; });
        selectionOrder = selectionOrder.filter(id => !removedSet.has(id));
        pagesState = pagesState.filter(p => p.fileId !== fileId);
        // Re-assign labels for remaining files to keep them sequential
        reassignFileLabels();
        updateUI();
    });
}

// Re-assign A/B/C labels after a file is removed
function reassignFileLabels() {
    const keys = Object.keys(uploadedFiles);
    keys.forEach((key, idx) => {
        uploadedFiles[key].label = fileIndexToLabel(idx);
    });
}

function insertBlankPage(index) {
    saveUndo();
    pagesState.splice(index, 0, {
        id: 'page_' + pageIdCounter++,
        isBlank: true,
        rotation: 0,
        thumbnail: createBlankPageThumbnail()
    });
    updateUI();
}

function copyPage(id, index) {
    const original = pagesState.find(p => p.id === id);
    if (!original) return;
    saveUndo();
    const newId = 'page_' + pageIdCounter++;
    const copy = { ...original, id: newId };
    if (thumbnailCache[id]) thumbnailCache[newId] = thumbnailCache[id];
    pagesState.splice(index + 1, 0, copy);
    updateUI();
}

function resetAll() {
    showConfirm('Clear all files and start over?', () => {
        saveUndo();
        uploadedFiles = {};
        pagesState = [];
        thumbnailCache = {};
        selectedPageIds.clear();
        lastSelectedIndex = -1;
        fileIdCounter = 0;
        updateUI();
    });
}

// ── Fullscreen Page Preview ────────────────────
async function openPagePreview(pageObj) {
    if (pageObj.isBlank) return;
    const overlay = document.getElementById('page-preview-overlay');
    const canvas = document.getElementById('page-preview-canvas');
    overlay.classList.remove('hidden');
    document.body.classList.add('loading');
    try {
        const bytes = uploadedFiles[pageObj.fileId].bytes.slice(0);
        const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
        const page = await pdf.getPage(pageObj.originalPageNum);
        // Render at high resolution (scale 2.5 for crisp preview)
        const scale = 2.5;
        const viewport = page.getViewport({ scale, rotation: pageObj.rotation });
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
    } catch (e) {
        console.error('Preview render failed', e);
    }
    document.body.classList.remove('loading');
}

function closePagePreview() {
    const overlay = document.getElementById('page-preview-overlay');
    if (overlay) overlay.classList.add('hidden');
}

// ── PDF Export ─────────────────────────────────

/**
 * Parse a custom range string using letter-prefixed page labels.
 * Format: "A1-A5, A8, B1-B5"
 * Returns an array of page objects from pagesState that match.
 */
function parseCustomRangeByLabel(rangeStr) {
    // Build a map from label (e.g. 'A3') -> page object
    const labelMap = {};
    pagesState.forEach(page => {
        if (!page.isBlank) {
            const label = getPageLabel(page);
            labelMap[label.toUpperCase()] = page;
        }
    });

    const result = [];
    const seen = new Set();

    for (let part of rangeStr.split(',')) {
        part = part.trim().toUpperCase();
        if (!part) continue;

        if (part.includes('-')) {
            // Range like A3-A7 or A3-B2 (cross-file ranges)
            const dashIdx = part.indexOf('-');
            const startLabel = part.slice(0, dashIdx).trim();
            const endLabel = part.slice(dashIdx + 1).trim();

            // Parse start: letter prefix + number
            const startMatch = startLabel.match(/^([A-Z]+)(\d+)$/);
            const endMatch = endLabel.match(/^([A-Z]+)(\d+)$/);
            if (!startMatch || !endMatch) continue;

            const [, sPrefix, sNumStr] = startMatch;
            const [, ePrefix, eNumStr] = endMatch;
            const sNum = parseInt(sNumStr, 10);
            const eNum = parseInt(eNumStr, 10);

            if (sPrefix === ePrefix) {
                // Same file: iterate page numbers
                for (let n = sNum; n <= eNum; n++) {
                    const lbl = sPrefix + n;
                    if (labelMap[lbl] && !seen.has(lbl)) {
                        result.push(labelMap[lbl]);
                        seen.add(lbl);
                    }
                }
            } else {
                // Cross-file range: collect all pages from startLabel to endLabel in
                // the order they appear in pagesState
                let inRange = false;
                for (const page of pagesState) {
                    if (page.isBlank) continue;
                    const lbl = getPageLabel(page).toUpperCase();
                    if (lbl === startLabel) inRange = true;
                    if (inRange && !seen.has(lbl)) {
                        result.push(page);
                        seen.add(lbl);
                    }
                    if (lbl === endLabel) { inRange = false; }
                }
            }
        } else {
            // Single label like A3
            if (labelMap[part] && !seen.has(part)) {
                result.push(labelMap[part]);
                seen.add(part);
            }
        }
    }
    return result;
}

async function organizeAndExport() {
    if (pagesState.length === 0) return alert('Please add at least one PDF.');
    document.body.classList.add('loading');
    try {
        const mode = document.querySelector('input[name="extract-mode"]:checked').value;
        let toExport = pagesState;
        if (mode === 'custom') {
            const parsed = parseCustomRangeByLabel(customInput.value);
            if (parsed.length === 0) {
                document.body.classList.remove('loading');
                return alert('Invalid custom range. Use format like: A1-A5, A8, B1-B5');
            }
            toExport = parsed;
        }

        const { PDFDocument, degrees } = PDFLib;
        const finalPdf = await PDFDocument.create();
        const loaded = {};

        for (const cfg of toExport) {
            if (cfg.isBlank) {
                const pg = finalPdf.addPage([595.28, 841.89]);
                if (cfg.rotation) pg.setRotation(degrees(cfg.rotation));
            } else {
                if (!loaded[cfg.fileId]) {
                    loaded[cfg.fileId] = await PDFDocument.load(uploadedFiles[cfg.fileId].bytes);
                }
                const [pg] = await finalPdf.copyPages(loaded[cfg.fileId], [cfg.originalPageNum - 1]);
                if (cfg.rotation) pg.setRotation(degrees(pg.getRotation().angle + cfg.rotation));
                finalPdf.addPage(pg);
            }
        }

        const bytes = await finalPdf.save();
        const a = Object.assign(document.createElement('a'), {
            href: URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' })),
            download: 'Organized_Document.pdf'
        });
        a.click();
    } catch (err) {
        console.error('Export error:', err);
        alert('An error occurred while generating the PDF.');
    }
    document.body.classList.remove('loading');
}