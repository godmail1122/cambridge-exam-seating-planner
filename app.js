// ============================================================
//  Exam Seating Planner - app.js
// ============================================================

const state = {
  mode: null,
  plans: [],
  currentPlanIndex: 0,
  singlePlan: null,
  multiRoomLayouts: [],
  multiCurrentRoomIndex: 0,
  previewPlans: [],
  selectedId: null,
  selectedHeaderKey: null,
  idCounter: 0,
  multiWorkbook: null,
  multiRows: [],
  previewEditable: false,
  previewMode: null,
  canvasZoom: 1,
  clipboard: null,
  history: { stack: [], pointer: -1 },
};

const SEAT_LAYOUT_PIXELS_PER_CM = 37.8;
const SEAT_LAYOUT_SIDE_MARGIN_CM = 1;
const SEAT_LAYOUT_MIN_SPAN_CM = 4;
const BROWSER_LAYOUTS_STORAGE_KEY = 'exam-seating-browser-layouts';
const BROWSER_LAYOUTS_TTL_MS = 30 * 24 * 60 * 60 * 1000;

// ============================================================
//  INIT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  bindSidebarEvents();
  bindToolbarEvents();
  setupCanvasViewportEvents();
  restoreBrowserSavedLayouts();
  updateLayoutActionState();
  document.addEventListener('keydown', onKeyDown);
  window.addEventListener('resize', syncCanvasZoom);
});

// ============================================================
//  UNDO / REDO HISTORY
// ============================================================
function pushHistory() {
  const plan = state.plans[state.currentPlanIndex];
  if (!plan) return;
  const snapshot = JSON.parse(JSON.stringify(plan));
  const h = state.history;
  h.stack = h.stack.slice(0, h.pointer + 1);
  h.stack.push(snapshot);
  if (h.stack.length > 60) h.stack.shift();
  h.pointer = h.stack.length - 1;
  persistBrowserSavedLayouts();
}

function undo() {
  if (!state.plans[state.currentPlanIndex]) return;
  const h = state.history;
  if (h.pointer <= 0) { setStatus('Nothing to undo.'); return; }
  h.pointer--;
  state.plans[state.currentPlanIndex] = JSON.parse(JSON.stringify(h.stack[h.pointer]));
  syncActivePlanStore();
  reRenderElements(state.plans[state.currentPlanIndex]);
  persistBrowserSavedLayouts();
  setStatus('Undo.');
}

function redo() {
  if (!state.plans[state.currentPlanIndex]) return;
  const h = state.history;
  if (h.pointer >= h.stack.length - 1) { setStatus('Nothing to redo.'); return; }
  h.pointer++;
  state.plans[state.currentPlanIndex] = JSON.parse(JSON.stringify(h.stack[h.pointer]));
  syncActivePlanStore();
  reRenderElements(state.plans[state.currentPlanIndex]);
  persistBrowserSavedLayouts();
  setStatus('Redo.');
}

function buildBrowserLayoutsPayload() {
  return {
    type: 'exam-seating-browser-layouts',
    version: 1,
    expiresAt: new Date(Date.now() + BROWSER_LAYOUTS_TTL_MS).toISOString(),
    data: buildLayoutsPayload('all'),
  };
}

function persistBrowserSavedLayouts() {
  if (typeof window === 'undefined' || !window.localStorage) return;
  if (state.mode === 'multiple' && state.previewMode === 'generated') return;

  try {
    if (!state.multiRoomLayouts.length) {
      window.localStorage.removeItem(BROWSER_LAYOUTS_STORAGE_KEY);
      return;
    }
    window.localStorage.setItem(BROWSER_LAYOUTS_STORAGE_KEY, JSON.stringify(buildBrowserLayoutsPayload()));
  } catch (error) {
    console.error(error);
  }
}

function restoreBrowserSavedLayouts() {
  if (typeof window === 'undefined' || !window.localStorage) return;

  try {
    const raw = window.localStorage.getItem(BROWSER_LAYOUTS_STORAGE_KEY);
    if (!raw) {
      updateSavedRoomSelect();
      updateSingleLoadedRoomSelect();
      return;
    }

    const payload = JSON.parse(raw);
    const expiresAt = Date.parse(payload?.expiresAt || '');
    if (!payload || payload.type !== 'exam-seating-browser-layouts' || !Number.isFinite(expiresAt) || expiresAt <= Date.now()) {
      window.localStorage.removeItem(BROWSER_LAYOUTS_STORAGE_KEY);
      updateSavedRoomSelect();
      updateSingleLoadedRoomSelect();
      return;
    }

    const storedPayload = payload.data;
    if (!storedPayload || storedPayload.type !== 'exam-seating-layouts' || !Array.isArray(storedPayload.plans)) {
      window.localStorage.removeItem(BROWSER_LAYOUTS_STORAGE_KEY);
      updateSavedRoomSelect();
      updateSingleLoadedRoomSelect();
      return;
    }

    const loadedPlans = storedPayload.plans.map(normalizeLoadedPlan).filter(Boolean);
    if (!loadedPlans.length) {
      window.localStorage.removeItem(BROWSER_LAYOUTS_STORAGE_KEY);
      updateSavedRoomSelect();
      updateSingleLoadedRoomSelect();
      return;
    }

    state.multiRoomLayouts = loadedPlans;
    state.idCounter = Math.max(state.idCounter, getMaxElementId(state.multiRoomLayouts));
    updateSavedRoomSelect();
    updateSingleLoadedRoomSelect();
  } catch (error) {
    console.error(error);
    window.localStorage.removeItem(BROWSER_LAYOUTS_STORAGE_KEY);
    updateSavedRoomSelect();
    updateSingleLoadedRoomSelect();
  }
}

function syncActivePlanStore() {
  if (state.mode === 'single') {
    state.singlePlan = state.plans[0] || null;
    return;
  }
  if (state.mode === 'multiple') {
    if (state.previewMode === 'generated') {
      state.previewPlans = state.plans;
      return;
    }
    state.multiRoomLayouts = state.plans;
    state.multiCurrentRoomIndex = state.currentPlanIndex;
  }
}

function reRenderElements(plan) {
  state.selectedId = null;
  state.selectedHeaderKey = null;
  renderCurrentPlan();
}

// ============================================================
//  SIDEBAR FLOW
// ============================================================
function bindSidebarEvents() {
  document.querySelectorAll('.mode-btn').forEach((btn) =>
    btn.addEventListener('click', () => selectMode(btn.dataset.mode))
  );
  document.querySelectorAll('.back-btn').forEach((btn) =>
    btn.addEventListener('click', () => showStep(btn.dataset.step))
  );
  document.getElementById('generateSingleBtn').addEventListener('click', generateSingle);
  document.getElementById('singleLoadedRoomSelect').addEventListener('change', loadSingleLoadedRoom);
  document.getElementById('updateSingleLayoutBtn').addEventListener('click', saveSingleRoomLayout);
  document.getElementById('multiFileInput').addEventListener('change', handleMultiFile);
  document.getElementById('multiSheetSelect').addEventListener('change', () =>
    loadMultiSheet(document.getElementById('multiSheetSelect').value)
  );
  document.getElementById('loadRoomLayoutBtn').addEventListener('click', () => {
    document.getElementById('loadLayoutsInput').click();
  });
  document.getElementById('saveRoomLayoutBtn').addEventListener('click', saveRoomLayout);
  document.getElementById('deleteRoomLayoutBtn').addEventListener('click', deleteCurrentRoomLayout);
  document.getElementById('saveAllRoomLayoutsBtn').addEventListener('click', () => saveLayouts('all'));
  document.getElementById('previewMultiLayoutsBtn').addEventListener('click', previewMultiLayouts);
  document.getElementById('savePreviewEditsBtn').addEventListener('click', savePreviewRoomEdits);
  document.getElementById('multiRoomNumberInput').addEventListener('input', updateLayoutActionState);
  document.getElementById('multiRoomNumberInput').addEventListener('change', handleMultiRoomNumberChange);
  document.getElementById('multiSavedRoomSelect').addEventListener('change', loadSelectedSavedRoom);
  document.getElementById('generateMultiBtn').addEventListener('click', generateMultiple);
  document.getElementById('downloadSampleLink').addEventListener('click', (e) => {
    e.preventDefault();
    downloadSample();
  });
  document.getElementById('loadLayoutsInput').addEventListener('change', handleLoadLayoutsFile);
  document.getElementById('exportPdfBtn').addEventListener('click', exportToPdf);
  document.getElementById('deleteBtn').addEventListener('click', deleteSelectedElement);
}
function selectMode(mode) {
  state.mode = mode;
  state.previewPlans = [];
  state.previewEditable = false;
  state.previewMode = null;
  showStep(mode === 'single' ? 'step-single' : 'step-multiple');
  if (mode === 'single') {
    activateSinglePlanView();
  } else {
    activateMultiLayoutView(state.multiCurrentRoomIndex);
  }
}

function showStep(stepId) {
  document.querySelectorAll('.sidebar-step').forEach((el) => el.classList.remove('active'));
  const target = document.getElementById(stepId);
  if (target) target.classList.add('active');
  updateLayoutActionState();
}

function activateSinglePlanView() {
  state.previewPlans = [];
  state.previewEditable = false;
  state.previewMode = null;
  state.plans = state.singlePlan ? [state.singlePlan] : [];
  state.currentPlanIndex = 0;
  updateSingleLoadedRoomSelect();
  renderCurrentPlan();
  document.getElementById('roomNav').style.display = 'none';
}

function activateMultiLayoutView(index = state.multiCurrentRoomIndex) {
  state.previewPlans = [];
  state.previewEditable = false;
  state.previewMode = null;
  state.plans = state.multiRoomLayouts;
  state.currentPlanIndex = state.multiRoomLayouts.length
    ? clamp(index, 0, state.multiRoomLayouts.length - 1)
    : 0;
  state.multiCurrentRoomIndex = state.currentPlanIndex;
  syncCurrentRoomNumberInput();
  updateSavedRoomSelect();
  renderCurrentPlan();
  renderRoomNav();
}

function syncCurrentRoomNumberInput() {
  const input = document.getElementById('multiRoomNumberInput');
  if (!input) return;
  const plan = state.multiRoomLayouts[state.multiCurrentRoomIndex];
  input.value = plan?.meta?.room || input.value || '';
}

function syncSingleInputsFromPlan(plan) {
  if (!plan) return;
  const meta = plan.meta || {};
  document.getElementById('schoolInput').value = meta.school || '';
  document.getElementById('examSeriesInput').value = meta.examSeries || '';
  document.getElementById('roomInput').value = meta.room || '';
  document.getElementById('paperInput').value = meta.paper || '';
  document.getElementById('syllabusInput').value = meta.syllabusCode || '';
  document.getElementById('componentInput').value = meta.componentCode || '';
  document.getElementById('candidateTextarea').value = (plan.seatingConfig?.candidates || []).join('\n');
  document.getElementById('colsInput').value = String(Math.max(1, Number(plan.seatingConfig?.columns) || 4));
  document.getElementById('rowsInput').value = String(Math.max(1, Number(plan.seatingConfig?.rows) || 4));
  document.getElementById('manualSeatInput').checked = !!plan.seatingConfig?.manualSeatMode;
}

function syncSingleInputsFromRoomLayout(plan) {
  if (!plan) return;
  const meta = plan.meta || {};
  document.getElementById('roomInput').value = meta.room || '';
  document.getElementById('colsInput').value = String(Math.max(1, Number(plan.seatingConfig?.columns) || 4));
  document.getElementById('rowsInput').value = String(Math.max(1, Number(plan.seatingConfig?.rows) || 4));
  document.getElementById('manualSeatInput').checked = !!plan.seatingConfig?.manualSeatMode;
}

function updateSingleLoadedRoomSelect() {
  const select = document.getElementById('singleLoadedRoomSelect');
  if (!select) return;

  const currentRoom = state.singlePlan?.meta?.room || '';
  select.innerHTML = ['<option value="">-- Select loaded room --</option>']
    .concat(state.multiRoomLayouts.map((plan) => {
      const room = escHtml(plan?.meta?.room || 'Room');
      return `<option value="${room}">${room}</option>`;
    }))
    .join('');
  select.disabled = state.multiRoomLayouts.length === 0;
  select.value = state.multiRoomLayouts.some((plan) => String(plan?.meta?.room || '').trim() === currentRoom) ? currentRoom : '';
}

function loadSingleLoadedRoom() {
  const select = document.getElementById('singleLoadedRoomSelect');
  if (!select || !select.value) return;

  const roomName = select.value;
  const sourcePlan = state.multiRoomLayouts.find((plan) => String(plan?.meta?.room || '').trim() === roomName);
  if (!sourcePlan) return;

  state.singlePlan = normalizeLoadedPlan(sourcePlan);
  state.idCounter = Math.max(state.idCounter, getMaxElementId([state.singlePlan]));
  syncSingleInputsFromRoomLayout(state.singlePlan);
  activateSinglePlanView();
  setStatus(`Loaded room layout for ${roomName} into Single Room.`);
}

function updateSavedRoomSelect() {
  const select = document.getElementById('multiSavedRoomSelect');
  if (!select) return;

  const currentRoom = state.multiRoomLayouts[state.multiCurrentRoomIndex]?.meta?.room || '';
  select.innerHTML = ['<option value="">-- Select saved room --</option>']
    .concat(state.multiRoomLayouts.map((plan) => {
      const room = escHtml(plan?.meta?.room || 'Room');
      return `<option value="${room}">${room}</option>`;
    }))
    .join('');
  select.value = currentRoom || '';
}

function handleMultiRoomNumberChange() {
  const roomName = document.getElementById('multiRoomNumberInput').value.trim();
  if (!roomName) return;

  const existingIndex = state.multiRoomLayouts.findIndex((plan) => String(plan?.meta?.room || '').trim() === roomName);
  if (existingIndex >= 0) {
    state.multiCurrentRoomIndex = existingIndex;
    activateMultiLayoutView(existingIndex);
    setStatus(`Loaded room layout for ${roomName}.`);
    return;
  }

  const options = getMultiLayoutOptions();
  const plan = buildPlan({ room: roomName }, [], options.columns, options.rows, null, { manualSeatMode: options.manualSeatMode });
  state.multiRoomLayouts.push(plan);
  state.multiCurrentRoomIndex = state.multiRoomLayouts.length - 1;
  activateMultiLayoutView(state.multiCurrentRoomIndex);
  setStatus(`Created a new room layout for ${roomName}.`);
}

function loadSelectedSavedRoom() {
  const select = document.getElementById('multiSavedRoomSelect');
  if (!select || !select.value) return;

  const roomName = select.value;
  const index = state.multiRoomLayouts.findIndex((plan) => String(plan?.meta?.room || '').trim() === roomName);
  if (index < 0) return;
  state.multiCurrentRoomIndex = index;
  activateMultiLayoutView(index);
  setStatus(`Loaded room layout for ${roomName}.`);
}

function getSavedRoomLayout(roomName) {
  const normalizedRoomName = String(roomName || '').trim();
  if (!normalizedRoomName) return null;

  const sourcePlan = state.multiRoomLayouts.find(
    (plan) => String(plan?.meta?.room || '').trim() === normalizedRoomName
  );
  return sourcePlan ? normalizeLoadedPlan(sourcePlan) : null;
}

function buildRoomLayoutTemplate(plan, roomName) {
  if (!plan) return null;

  const normalizedPlan = normalizeLoadedPlan(plan);
  const normalizedRoomName = String(roomName || normalizedPlan.meta?.room || '').trim() || 'Room';
  const classroom = normalizedPlan.classroom
    ? { ...normalizedPlan.classroom }
    : { width: 690, height: 740 };

  return {
    meta: { room: normalizedRoomName },
    headerLines: buildHeaderLines({ room: normalizedRoomName }),
    headerStyles: normalizedPlan.headerStyles
      ? JSON.parse(JSON.stringify(normalizedPlan.headerStyles))
      : buildHeaderStyles(),
    classroom,
    seatingConfig: {
      candidates: [],
      columns: Math.max(1, Number(normalizedPlan.seatingConfig?.columns) || 4),
      rows: Math.max(1, Number(normalizedPlan.seatingConfig?.rows) || 4),
      manualSeatMode: !!normalizedPlan.seatingConfig?.manualSeatMode,
    },
    seatingArea: normalizedPlan.seatingArea
      ? { ...normalizedPlan.seatingArea }
      : buildInitialSeatingArea(classroom),
    elements: buildPersistentLayout(normalizedPlan).map((element) => (
      element.type === 'seat' && element.manualSeat
        ? { ...element, label: '' }
        : element
    )),
  };
}

function getSingleRoomTemplate(roomName) {
  const normalizedRoomName = String(roomName || '').trim();
  const currentPlanRoom = String(state.singlePlan?.meta?.room || '').trim();
  if (normalizedRoomName && currentPlanRoom === normalizedRoomName) {
    return state.singlePlan;
  }
  return getSavedRoomLayout(normalizedRoomName);
}

function upsertSingleRoomLayout(plan) {
  const roomName = document.getElementById('roomInput')?.value.trim() || '';
  if (!roomName) {
    return { ok: false, reason: 'missing-room' };
  }
  if (!plan) {
    return { ok: false, reason: 'missing-plan' };
  }

  const savedPlan = buildRoomLayoutTemplate(plan, roomName);
  if (!savedPlan) {
    return { ok: false, reason: 'missing-plan' };
  }

  const existingIndex = state.multiRoomLayouts.findIndex(
    (entry) => String(entry?.meta?.room || '').trim() === roomName
  );

  if (existingIndex >= 0) {
    state.multiRoomLayouts[existingIndex] = savedPlan;
  } else {
    state.multiRoomLayouts.push(savedPlan);
  }

  updateSavedRoomSelect();
  updateSingleLoadedRoomSelect();
  persistBrowserSavedLayouts();

  return {
    ok: true,
    action: existingIndex >= 0 ? 'updated' : 'added',
    roomName,
  };
}

function saveSingleRoomLayout() {
  if (!state.singlePlan) {
    setStatus('Generate or load a single-room plan first.');
    return;
  }

  const result = upsertSingleRoomLayout(state.singlePlan);
  if (!result.ok) {
    setStatus('Enter a room number first.');
    return;
  }

  setStatus(`${result.action === 'updated' ? 'Updated' : 'Saved'} room layout for ${result.roomName}.`);
}

// ============================================================
//  SINGLE ROOM
// ============================================================
function generateSingle() {
  const meta = {
    school: document.getElementById('schoolInput').value.trim() || 'School Name',
    examSeries: document.getElementById('examSeriesInput').value.trim() || 'Exam Series',
    room: document.getElementById('roomInput').value.trim() || 'Room',
    paper: document.getElementById('paperInput').value.trim() || 'Paper',
    syllabusCode: document.getElementById('syllabusInput').value.trim() || '',
    componentCode: document.getElementById('componentInput').value.trim() || '',
  };
  const previousPlan = getSingleRoomTemplate(meta.room);
  const manualSeatMode = !!document.getElementById('manualSeatInput')?.checked;
  const candidates = document.getElementById('candidateTextarea').value
    .split('\n').map((s) => normalizeCandidateNumber(s)).filter(Boolean);
  if (!candidates.length) { setStatus('Enter at least one candidate number.'); return; }

  const cols = Math.max(1, parseInt(document.getElementById('colsInput').value, 10) || 4);
  const rows = Math.max(1, parseInt(document.getElementById('rowsInput').value, 10) || 4);
  const previousManualSeats = getOrderedManualSeats(previousPlan?.elements || []);
  if (!manualSeatMode && !canFitSeatGrid(candidates.length, cols, rows)) {
    setStatus(`Increase rows or columns. ${candidates.length} candidate(s) do not fit in a ${rows} x ${cols} grid.`);
    return;
  }
  if (manualSeatMode && previousManualSeats.length && previousManualSeats.length < candidates.length) {
    setStatus(`Add more manual seats. You have ${previousManualSeats.length} manual seat(s) for ${candidates.length} candidate(s).`);
    return;
  }
  state.idCounter = 0;
  state.history = { stack: [], pointer: -1 };
  const plan = buildPlan(meta, candidates, cols, rows, previousPlan, { manualSeatMode });
  state.singlePlan = plan;
  state.plans = [state.singlePlan];
  state.previewPlans = [];
  state.previewEditable = false;
  state.previewMode = null;
  state.currentPlanIndex = 0;
  renderCurrentPlan();
  const saveResult = upsertSingleRoomLayout(plan);
  pushHistory();
  document.getElementById('roomNav').style.display = 'none';
  if (manualSeatMode) {
    const manualSeatCount = getOrderedManualSeats(plan.elements).length;
    const layoutMessage = saveResult.ok
      ? ` ${saveResult.action === 'updated' ? 'Updated' : 'Saved'} room layout for ${saveResult.roomName}.`
      : ' Enter a room number to save this layout.';
    setStatus(
      manualSeatCount
        ? `Manual seat mode active. Filled ${Math.min(candidates.length, manualSeatCount)} manual seat(s).${layoutMessage}`
        : `Manual seat mode active. Add seats from the toolbar, place them, then generate again to fill candidate numbers.${layoutMessage}`
    );
    return;
  }
  setStatus(
    saveResult.ok
      ? `Plan generated with ${candidates.length} seat(s). ${saveResult.action === 'updated' ? 'Updated' : 'Saved'} room layout for ${saveResult.roomName}.`
      : `Plan generated with ${candidates.length} seat(s). Enter a room number to save this layout.`
  );
}

// ============================================================
//  MULTIPLE ROOMS
// ============================================================
async function handleMultiFile(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    const buffer = await file.arrayBuffer();
    state.multiWorkbook = XLSX.read(buffer, { type: 'array' });
    const names = state.multiWorkbook.SheetNames;
    const sel = document.getElementById('multiSheetSelect');
    sel.innerHTML = names.map((n) => `<option value="${escHtml(n)}">${escHtml(n)}</option>`).join('');
    document.getElementById('multiSheetRow').style.display = '';
    loadMultiSheet(names[0]);
  } catch (err) {
    console.error(err);
    setStatus('Could not read the file.');
  }
}

function loadMultiSheet(name) {
  const sheet = state.multiWorkbook.Sheets[name];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
  state.multiRows = rows;
  const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
  populateMultiSelects(headers);
  autoMapMultiSelects(headers);
}

function populateMultiSelects(headers) {
  const opts = ['<option value="">-- Not mapped --</option>']
    .concat(headers.map((h) => `<option value="${escHtml(h)}">${escHtml(h)}</option>`))
    .join('');
  ['multiSchoolCol','multiSeriesCol','multiRoomCol','multiPaperCol',
   'multiSyllabusCol','multiComponentCol','multiCandidateCol'].forEach((id) => {
    document.getElementById(id).innerHTML = opts;
  });
}

function findHeaderMatch(headers, keywords) {
  const normalizedHeaders = headers.map((header) => ({
    original: header,
    lower: String(header).trim().toLowerCase(),
  }));

  const exactMatch = normalizedHeaders.find((header) => keywords.some((keyword) => header.lower === keyword));
  if (exactMatch) return exactMatch.original;

  const partialMatch = normalizedHeaders.find((header) => keywords.some((keyword) => header.lower.includes(keyword)));
  return partialMatch ? partialMatch.original : '';
}

function autoMapMultiSelects(headers) {
  const mappings = {
    multiSchoolCol: ['school name','school','centre name','centre'],
    multiSeriesCol: ['exam series','series','session','exam title'],
    multiRoomCol: ['room number','room','hall','venue'],
    multiPaperCol: ['paper','syllabus title','subject title'],
    multiSyllabusCol: ['syllabus code','subject code'],
    multiComponentCol: ['component code','component'],
    multiCandidateCol: ['candidate number','candidate no','candidate','candidates','index'],
  };
  Object.entries(mappings).forEach(([selectId, keywords]) => {
    const match = findHeaderMatch(headers, keywords);
    if (match) document.getElementById(selectId).value = match;
  });
}

function getImportedSeatCount(value, fallback = 4) {
  const parsed = parseInt(String(value ?? '').trim(), 10);
  return Math.max(1, Number.isFinite(parsed) ? parsed : fallback);
}

function parseCandidateCell(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => normalizeCandidateNumber(entry)).filter(Boolean);
  }
  const text = String(value ?? '').trim();
  if (!text) return [];
  return text
    .split(/[\n,;]+/)
    .map((entry) => normalizeCandidateNumber(entry))
    .filter(Boolean);
}

function buildImportedPlanEntries(cols) {
  const candidateHeader = String(cols.candidate || '').trim().toLowerCase();
  const rowBasedFormat = !!(cols.rows || cols.cols || candidateHeader === 'candidates' || candidateHeader === 'candidate list');

  if (rowBasedFormat) {
    return state.multiRows
      .map((row, index) => {
        const candidates = parseCandidateCell(cols.candidate ? row[cols.candidate] : '');
        if (!candidates.length) return null;

        const meta = {
          school: cols.school ? String(row[cols.school] ?? '').trim() : 'School Name',
          examSeries: cols.series ? String(row[cols.series] ?? '').trim() : 'Exam Series',
          room: cols.room ? String(row[cols.room] ?? '').trim() || `Room ${index + 1}` : `Room ${index + 1}`,
          paper: cols.paper ? String(row[cols.paper] ?? '').trim() : '',
          syllabusCode: cols.syllabus ? String(row[cols.syllabus] ?? '').trim() : '',
          componentCode: cols.component ? String(row[cols.component] ?? '').trim() : '',
        };

        return {
          meta,
          roomName: meta.room,
          candidates,
          rows: getImportedSeatCount(cols.rows ? row[cols.rows] : '', Math.ceil(candidates.length / 4) || 1),
          columns: getImportedSeatCount(cols.cols ? row[cols.cols] : '', Math.min(4, Math.max(1, candidates.length))),
          hasExplicitRows: !!cols.rows,
          hasExplicitColumns: !!cols.cols,
          requiresTemplate: false,
          sourceLabel: `${meta.room}${meta.paper ? ` / ${meta.paper}` : ''}`,
        };
      })
      .filter(Boolean);
  }

  const perRoom = new Map();
  state.multiRows.forEach((row) => {
    const roomKey = cols.room ? String(row[cols.room] ?? '').trim() || 'Room 1' : 'Room 1';
    if (!perRoom.has(roomKey)) perRoom.set(roomKey, []);
    perRoom.get(roomKey).push(row);
  });

  return Array.from(perRoom.entries()).map(([roomName, rows]) => {
    const first = rows[0] || {};
    return {
      meta: {
        school: cols.school ? String(first[cols.school] ?? '').trim() : 'School Name',
        examSeries: cols.series ? String(first[cols.series] ?? '').trim() : 'Exam Series',
        room: roomName,
        paper: cols.paper ? String(first[cols.paper] ?? '').trim() : '',
        syllabusCode: cols.syllabus ? String(first[cols.syllabus] ?? '').trim() : '',
        componentCode: cols.component ? String(first[cols.component] ?? '').trim() : '',
      },
      roomName,
      candidates: rows.map((entry) => normalizeCandidateNumber(entry[cols.candidate])).filter(Boolean),
      rows: null,
      columns: null,
      hasExplicitRows: false,
      hasExplicitColumns: false,
      requiresTemplate: true,
      sourceLabel: roomName,
    };
  });
}

function getImportedGrid(entry, defaultColumns = 4, defaultRows = 4) {
  const columns = Math.max(1, Number(entry?.columns) || defaultColumns);
  const requestedRows = Math.max(1, Number(entry?.rows) || defaultRows);
  const minimumRowsToFit = Math.max(1, Math.ceil((entry?.candidates?.length || 0) / columns));
  return {
    columns,
    rows: Math.max(requestedRows, minimumRowsToFit),
  };
}

function generateMultiple() {
  if (!state.multiRows.length) { setStatus('Load a file first.'); return; }

  const cols = getExcelColumnMapping();
  if (!cols.candidate) { setStatus('Map the Candidate Number column first.'); return; }
  const planEntries = buildImportedPlanEntries(cols);
  if (!planEntries.length) { setStatus('No valid rows were found in the file.'); return; }

  const templatesByRoom = new Map(
    state.multiRoomLayouts.map((plan) => [String(plan?.meta?.room || '').trim(), plan])
  );

  for (const entry of planEntries) {
    const template = templatesByRoom.get(entry.roomName);
    if (!template && entry.requiresTemplate) {
      setStatus(`No saved room layout was found for ${entry.sourceLabel}.`);
      return;
    }

    if (!template) {
      const { columns: seatsPerRow, rows: rowCount } = getImportedGrid(entry, 4, 1);
      if (!canFitSeatGrid(entry.candidates.length, seatsPerRow, rowCount)) {
        setStatus(`Increase rows or columns in the file for ${entry.sourceLabel}. ${entry.candidates.length} candidate(s) do not fit in a ${rowCount} x ${seatsPerRow} grid.`);
        return;
      }
      continue;
    }

    const seatingConfig = template.seatingConfig || {};
    const manualSeatMode = !!seatingConfig.manualSeatMode;
    const importedGrid = getImportedGrid(entry, 4, 1);
    const seatsPerRow = entry.hasExplicitColumns
      ? importedGrid.columns
      : Math.max(1, Number(seatingConfig.columns) || importedGrid.columns);
    const rowCount = entry.hasExplicitRows
      ? importedGrid.rows
      : Math.max(1, Number(seatingConfig.rows) || importedGrid.rows);
    const manualSeatCount = getOrderedManualSeats(template.elements || []).length;

    if (!manualSeatMode && !canFitSeatGrid(entry.candidates.length, seatsPerRow, rowCount)) {
      setStatus(`Increase rows or columns in the saved layout for ${entry.sourceLabel}. ${entry.candidates.length} candidate(s) do not fit in a ${rowCount} x ${seatsPerRow} grid.`);
      return;
    }
    if (manualSeatMode && manualSeatCount && manualSeatCount < entry.candidates.length) {
      setStatus(`Add more manual seats in the saved layout for ${entry.sourceLabel}. ${manualSeatCount} seat(s) found for ${entry.candidates.length} candidate(s).`);
      return;
    }
  }

  const generatedPlans = [];

  planEntries.forEach((entry) => {
    const template = templatesByRoom.get(entry.roomName);
    const seatingConfig = template?.seatingConfig || {};
    const importedGrid = getImportedGrid(entry, 4, 1);
    const seatsPerRow = entry.hasExplicitColumns
      ? importedGrid.columns
      : Math.max(1, Number(seatingConfig.columns) || importedGrid.columns);
    const rowCount = entry.hasExplicitRows
      ? importedGrid.rows
      : Math.max(1, Number(seatingConfig.rows) || importedGrid.rows);
    generatedPlans.push(
      buildPlan(
        entry.meta,
        entry.candidates,
        seatsPerRow,
        rowCount,
        template,
        { manualSeatMode: !!seatingConfig.manualSeatMode }
      )
    );
  });

  state.plans = generatedPlans;
  state.currentPlanIndex = 0;
  state.previewPlans = generatedPlans;
  state.previewEditable = true;
  state.previewMode = 'generated';
  renderCurrentPlan();
  renderRoomNav();
  setStatus(`Generated ${generatedPlans.length} plan(s) from Excel.`);
}

function getExcelColumnMapping() {
  const headers = state.multiRows.length ? Object.keys(state.multiRows[0]) : [];
  const autoDetectedExtras = {
    rows: findHeaderMatch(headers, ['rows', 'row count']),
    cols: findHeaderMatch(headers, ['cols', 'columns', 'column count']),
  };
  const selectMapping = {
    school: document.getElementById('multiSchoolCol')?.value,
    series: document.getElementById('multiSeriesCol')?.value,
    room: document.getElementById('multiRoomCol')?.value,
    paper: document.getElementById('multiPaperCol')?.value,
    syllabus: document.getElementById('multiSyllabusCol')?.value,
    component: document.getElementById('multiComponentCol')?.value,
    candidate: document.getElementById('multiCandidateCol')?.value,
  };
  if (Object.values(selectMapping).some(Boolean)) {
    return {
      ...selectMapping,
      ...autoDetectedExtras,
    };
  }

  const resolved = {};
  const mappings = {
    school: ['school name', 'school'],
    series: ['exam series', 'series', 'exam title'],
    room: ['room number', 'room'],
    paper: ['paper', 'syllabus title', 'subject title'],
    syllabus: ['syllabus code', 'syllabus'],
    component: ['component code', 'component'],
    candidate: ['candidate number', 'candidate', 'candidates'],
    rows: ['rows', 'row count'],
    cols: ['cols', 'columns', 'column count'],
  };
  Object.entries(mappings).forEach(([key, names]) => {
    resolved[key] = findHeaderMatch(headers, names);
  });
  return resolved;
}

function getMultiLayoutOptions() {
  return {
    columns: Math.max(1, parseInt(document.getElementById('multiColsInput').value, 10) || 4),
    rows: Math.max(1, parseInt(document.getElementById('multiRowsInput').value, 10) || 4),
    manualSeatMode: !!document.getElementById('multiManualSeatInput')?.checked,
  };
}

function saveRoomLayout() {
  const roomName = document.getElementById('multiRoomNumberInput').value.trim();
  if (!roomName) {
    setStatus('Enter a room number first.');
    return;
  }

  let plan = state.multiRoomLayouts[state.multiCurrentRoomIndex];
  if (!plan) {
    handleMultiRoomNumberChange();
    plan = state.multiRoomLayouts[state.multiCurrentRoomIndex];
  }
  if (!plan) return;

  const options = getMultiLayoutOptions();
  plan.meta = { ...(plan.meta || {}), room: roomName };
  plan.headerLines = buildHeaderLines(plan.meta);
  plan.seatingConfig = {
    ...(plan.seatingConfig || {}),
    columns: options.columns,
    rows: options.rows,
    manualSeatMode: options.manualSeatMode,
    candidates: Array.isArray(plan.seatingConfig?.candidates) ? plan.seatingConfig.candidates : [],
  };

  const duplicateIndex = state.multiRoomLayouts.findIndex((entry, index) => index !== state.multiCurrentRoomIndex && String(entry?.meta?.room || '').trim() === roomName);
  if (duplicateIndex >= 0) {
    state.multiRoomLayouts[duplicateIndex] = plan;
    state.multiRoomLayouts.splice(state.multiCurrentRoomIndex, 1);
    state.multiCurrentRoomIndex = duplicateIndex;
  }

  activateMultiLayoutView(state.multiCurrentRoomIndex);
  pushHistory();
  setStatus(`Saved room layout for ${roomName}.`);
}

function deleteCurrentRoomLayout() {
  if (state.mode !== 'multiple') return;

  const currentPlan = state.multiRoomLayouts[state.multiCurrentRoomIndex];
  if (!currentPlan) {
    setStatus('No saved room layout to delete.');
    return;
  }

  const deletedRoomName = currentPlan.meta?.room || 'room';
  state.previewPlans = [];
  state.previewEditable = false;
  state.previewMode = null;
  state.selectedId = null;
  state.selectedHeaderKey = null;
  state.multiRoomLayouts.splice(state.multiCurrentRoomIndex, 1);

  if (state.multiRoomLayouts.length) {
    const nextIndex = clamp(state.multiCurrentRoomIndex, 0, state.multiRoomLayouts.length - 1);
    state.multiCurrentRoomIndex = nextIndex;
    activateMultiLayoutView(nextIndex);
  } else {
    state.multiCurrentRoomIndex = 0;
    state.plans = [];
    state.currentPlanIndex = 0;
    renderCurrentPlan();
    renderRoomNav();
    updateSavedRoomSelect();
    const roomInput = document.getElementById('multiRoomNumberInput');
    if (roomInput) roomInput.value = '';
    const roomSelect = document.getElementById('multiSavedRoomSelect');
    if (roomSelect) roomSelect.value = '';
    updateLayoutActionState();
  }

  setStatus(`Deleted room layout for ${deletedRoomName}.`);
}

function previewMultiLayouts() {
  if (!state.multiRoomLayouts.length) {
    setStatus('Create or load room layouts first.');
    return;
  }

  state.plans = state.multiRoomLayouts;
  state.currentPlanIndex = clamp(state.multiCurrentRoomIndex, 0, state.multiRoomLayouts.length - 1);
  state.previewPlans = state.multiRoomLayouts;
  state.previewEditable = true;
  state.previewMode = 'layouts';
  renderCurrentPlan();
  renderRoomNav();
  setStatus(`Previewing ${state.previewPlans.length} room layout(s).`);
}

function savePreviewRoomEdits() {
  if (state.mode !== 'multiple' || !state.multiRoomLayouts.length) {
    setStatus('Create or load room layouts first.');
    return;
  }

  if (!state.previewEditable || !state.previewPlans.length || state.previewMode !== 'layouts') {
    setStatus('Open View All Rooms to edit and save stacked room layouts.');
    return;
  }

  const savedCount = state.multiRoomLayouts.length;
  state.previewPlans = [];
  state.previewEditable = false;
  state.previewMode = null;
  activateMultiLayoutView(state.multiCurrentRoomIndex);
  pushHistory();
  setStatus(`Saved edits for ${savedCount} room layout(s).`);
}

// ============================================================
//  BUILD PLAN
// ============================================================
function buildPlan(meta, candidates, columns, rows, previousPlan = null, seatOptions = {}) {
  const classroom = previousPlan?.classroom
    ? { ...previousPlan.classroom }
    : { width: 690, height: 740 };
  const manualSeatMode = !!seatOptions.manualSeatMode;

  const seatingConfig = {
    candidates: [...candidates],
    columns,
    rows,
    manualSeatMode,
  };
  const seatingArea = previousPlan?.seatingArea
    ? { ...previousPlan.seatingArea }
    : buildInitialSeatingArea(classroom, seatOptions.maxSeatWidthCm);
  const elements = buildPersistentLayout(previousPlan);
  if (manualSeatMode) {
    applyCandidatesToManualSeats(getOrderedManualSeats(elements), candidates);
  } else {
    elements.push(...buildSeatElements(candidates, columns, rows, classroom, { ...seatOptions, seatingArea }));
  }

  return {
    meta,
    headerLines: buildHeaderLines(meta),
    headerStyles: previousPlan?.headerStyles
      ? JSON.parse(JSON.stringify(previousPlan.headerStyles))
      : buildHeaderStyles(),
    classroom,
    seatingConfig,
    seatingArea,
    elements,
  };
}

function buildSeatElements(candidates, columns, rows, classroom, seatOptions = {}) {
  if (!candidates.length) return [];

  const defaultSeatWidth = 100;
  const defaultSeatHeight = 95;
  const minimumGapX = 18;
  const minimumGapY = 18;
  const minimumSeatWidth = 25;
  const minimumSeatHeight = 20;
  const seatsPerRow = Math.max(1, Math.min(columns, candidates.length));
  const rowCount = Math.max(1, Math.min(rows || Math.ceil(candidates.length / seatsPerRow), candidates.length));
  const seatAreaMetrics = getSeatAreaMetrics(classroom, seatOptions.seatingArea, seatOptions.maxSeatWidthCm);
  const { availableWidth, availableHeight, seatAreaLeft, seatAreaTop } = seatAreaMetrics;
  const rowSeatCounts = buildRowSeatCounts(candidates.length, rowCount, seatsPerRow);

  let seatScale = Math.min(
    1,
    (availableWidth - minimumGapX * (seatsPerRow - 1)) / (seatsPerRow * defaultSeatWidth),
    (availableHeight - minimumGapY * (rowCount - 1)) / (rowCount * defaultSeatHeight)
  );
  seatScale = Number.isFinite(seatScale) ? seatScale : 1;
  const seatWidth = Math.max(minimumSeatWidth, Math.floor(defaultSeatWidth * seatScale));
  const seatHeight = Math.max(minimumSeatHeight, Math.floor(defaultSeatHeight * seatScale));

  const verticalGap = rowCount > 1
    ? Math.max(0, (availableHeight - rowCount * seatHeight) / (rowCount - 1))
    : 0;

  const elements = [];
  let candidateIndex = 0;
  const fullRowHorizontalGap = seatsPerRow > 1
    ? Math.max(0, (availableWidth - seatsPerRow * seatWidth) / (seatsPerRow - 1))
    : 0;
  const fullRowSlotPositions = Array.from({ length: seatsPerRow }, (_, slot) =>
    seatAreaLeft + slot * (seatWidth + fullRowHorizontalGap)
  );

  for (let row = 0; row < rowCount; row++) {
    const seatsInThisRow = rowSeatCounts[row] || 0;
    if (!seatsInThisRow) continue;
    const y = Math.round(seatAreaTop + row * (seatHeight + verticalGap));

    for (let seatIndex = 0; seatIndex < seatsInThisRow; seatIndex++) {
      const candidate = candidates[candidateIndex++];
      const x = Math.round(fullRowSlotPositions[seatIndex]);
      elements.push(makeEl('seat', candidate, x, y, seatWidth, seatHeight, { manualSeat: false }));
    }
  }

  return elements;
}

function getOrderedManualSeats(elements = []) {
  return elements
    .filter((element) => element.type === 'seat' && element.manualSeat)
    .sort((left, right) => (left.y - right.y) || (left.x - right.x) || left.id.localeCompare(right.id));
}

function applyCandidatesToManualSeats(seats, candidates) {
  seats.forEach((seat, index) => {
    seat.label = candidates[index] || '';
  });
}

function buildRowSeatCounts(candidateCount, rowCount, seatsPerRow) {
  const counts = [];
  let remainingCandidates = candidateCount;

  for (let row = 0; row < rowCount; row++) {
    const seatsInRow = Math.min(seatsPerRow, remainingCandidates);
    counts.push(seatsInRow);
    remainingCandidates -= seatsInRow;
  }

  return counts;
}

function canFitSeatGrid(candidateCount, columns, rows) {
  return candidateCount <= columns * rows;
}

function buildInitialSeatingArea(classroom, maxSeatWidthCm) {
  const sideMarginPx = Math.round(SEAT_LAYOUT_SIDE_MARGIN_CM * SEAT_LAYOUT_PIXELS_PER_CM);
  const topMarginPx = Math.round(3 * SEAT_LAYOUT_PIXELS_PER_CM);
  const bottomMarginPx = Math.round(2 * SEAT_LAYOUT_PIXELS_PER_CM);
  const baseAvailableWidth = Math.max(0, classroom.width - sideMarginPx * 2);
  const baseAvailableHeight = Math.max(0, classroom.height - topMarginPx - bottomMarginPx);
  const configuredWidth = maxSeatWidthCm
    ? Math.min(baseAvailableWidth, Math.round(maxSeatWidthCm * SEAT_LAYOUT_PIXELS_PER_CM))
    : baseAvailableWidth;
  const extraWidth = Math.max(0, baseAvailableWidth - configuredWidth);
  const configuredHeight = baseAvailableHeight;
  const extraHeight = Math.max(0, baseAvailableHeight - configuredHeight);

  return {
    leftInsetPx: Math.round(extraWidth / 2),
    rightInsetPx: Math.round(extraWidth / 2),
    topInsetPx: Math.round(extraHeight / 2),
    bottomInsetPx: Math.round(extraHeight / 2),
  };
}

function getSeatAreaMetrics(classroom, seatingArea, maxSeatWidthCm) {
  const sideMarginPx = Math.round(SEAT_LAYOUT_SIDE_MARGIN_CM * SEAT_LAYOUT_PIXELS_PER_CM);
  const topMarginPx = Math.round(3 * SEAT_LAYOUT_PIXELS_PER_CM);
  const bottomMarginPx = Math.round(2 * SEAT_LAYOUT_PIXELS_PER_CM);
  const baseAvailableWidth = Math.max(0, classroom.width - sideMarginPx * 2);
  const baseAvailableHeight = Math.max(0, classroom.height - topMarginPx - bottomMarginPx);
  const defaultArea = seatingArea || buildInitialSeatingArea(classroom, maxSeatWidthCm);
  const leftInsetPx = clamp(Math.round(defaultArea.leftInsetPx || 0), 0, baseAvailableWidth);
  const maxRightInset = Math.max(0, baseAvailableWidth - leftInsetPx - Math.round(SEAT_LAYOUT_MIN_SPAN_CM * SEAT_LAYOUT_PIXELS_PER_CM));
  const rightInsetPx = clamp(Math.round(defaultArea.rightInsetPx || 0), 0, maxRightInset);
  const availableWidth = Math.max(0, baseAvailableWidth - leftInsetPx - rightInsetPx);
  const topInsetPx = clamp(Math.round(defaultArea.topInsetPx || 0), 0, baseAvailableHeight);
  const minVerticalSpanPx = Math.round(SEAT_LAYOUT_MIN_SPAN_CM * SEAT_LAYOUT_PIXELS_PER_CM);
  const maxBottomInset = Math.max(0, baseAvailableHeight - topInsetPx - minVerticalSpanPx);
  const bottomInsetPx = clamp(Math.round(defaultArea.bottomInsetPx || 0), 0, maxBottomInset);
  const availableHeight = Math.max(0, baseAvailableHeight - topInsetPx - bottomInsetPx);

  return {
    sideMarginPx,
    topMarginPx,
    bottomMarginPx,
    baseAvailableWidth,
    baseAvailableHeight,
    availableWidth,
    availableHeight,
    seatAreaLeft: sideMarginPx + leftInsetPx,
    seatAreaRight: sideMarginPx + leftInsetPx + availableWidth,
    seatAreaTop: topMarginPx + topInsetPx,
    seatAreaBottom: topMarginPx + topInsetPx + availableHeight,
    leftInsetPx,
    rightInsetPx,
    topInsetPx,
    bottomInsetPx,
  };
}

function normalizeCandidateNumber(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  if (/^\d+$/.test(text) && text.length < 4) {
    return text.padStart(4, '0');
  }
  return text;
}

function getDefaultElementSpec(type) {
  const specs = {
    door: { label: 'DOOR', width: 60, height: 22 },
    window: { label: 'WINDOW', width: 110, height: 30 },
    whiteboard: { label: 'WHITEBOARD', width: 220, height: 28 },
    softboard: { label: 'SOFT BOARD', width: 170, height: 24 },
    invigdesk: { label: "INVIGILATOR'S DESK", width: 200, height: 30 },
    seat: { label: '0012', width: 100, height: 95 },
    table: { label: 'TABLE', width: 400, height: 44 },
    signature: { label: "Exam Officer's Signature", width: 240, height: 42 },
    facingdirection: { label: 'Candidates facing this way', width: 54, height: 260 },
    text: { label: 'Text here', width: 130, height: 32 },
    box: { label: '', width: 110, height: 60 },
  };

  return specs[type] || { label: '', width: 100, height: 50 };
}

function buildPersistentLayout(previousPlan) {
  if (!previousPlan) {
    const whiteboard = getDefaultElementSpec('whiteboard');
    const softboard = getDefaultElementSpec('softboard');
    const invigdesk = getDefaultElementSpec('invigdesk');
    const signature = getDefaultElementSpec('signature');
    return [
      makeEl('whiteboard', whiteboard.label, 20, 16, whiteboard.width, whiteboard.height),
      makeEl('softboard', softboard.label, 258, 16, softboard.width, softboard.height),
      makeEl('invigdesk', invigdesk.label, 446, 16, invigdesk.width, invigdesk.height),
      makeEl('signature', signature.label, 450, -24, signature.width, signature.height, { region: 'footer' }),
    ];
  }

  const elements = previousPlan.elements
    .filter((element) => element.type !== 'seat' || element.manualSeat)
    .map(cloneLayoutElement);

  elements.forEach((element) => {
    if (element.type === 'signature' && element.region === 'footer' && element.y === 6) {
      element.y = -24;
    }
  });

  if (!elements.some((element) => element.type === 'signature')) {
    const signature = getDefaultElementSpec('signature');
    elements.push(makeEl('signature', signature.label, 450, -24, signature.width, signature.height, { region: 'footer' }));
  }

  return elements;
}

function cloneLayoutElement(element) {
  return {
    ...JSON.parse(JSON.stringify(element)),
    id: `el-${++state.idCounter}`,
  };
}

function buildHeaderLines(meta) {
  return {
    school: meta.school || 'School Name',
    examSeries: `Exam Seating Plan - ${meta.examSeries || 'Exam Series'}`,
    paper: `Paper - ${meta.paper || ''}`.trim(),
    roomNumber: `Room number - ${meta.room || ''}`.trim(),
    syllabusCode: `Syllabus code - ${meta.syllabusCode || ''}`.trim(),
    componentCode: `Component code - ${meta.componentCode || ''}`.trim(),
  };
}

function buildHeaderStyles() {
  return {
    school: { fontSize: 12, fontWeight: 'bold', fontStyle: 'normal', textAlign: 'center', borderWidth: 0 },
    examSeries: { fontSize: 11, fontWeight: 'bold', fontStyle: 'normal', textAlign: 'center', borderWidth: 0 },
    paper: { fontSize: 11, fontWeight: 'bold', fontStyle: 'normal', textAlign: 'center', borderWidth: 0 },
    roomNumber: { fontSize: 11, fontWeight: 'bold', fontStyle: 'normal', textAlign: 'center', borderWidth: 0 },
    syllabusCode: { fontSize: 11, fontWeight: 'bold', fontStyle: 'normal', textAlign: 'center', borderWidth: 0 },
    componentCode: { fontSize: 11, fontWeight: 'bold', fontStyle: 'normal', textAlign: 'center', borderWidth: 0 },
  };
}

function makeEl(type, label, x, y, w, h, options = {}) {
  const boldTypes = new Set(['seat','whiteboard','invigdesk','door']);
  const element = {
    id: `el-${++state.idCounter}`,
    type, label, x, y,
    width: w, height: h,
    region: options.region || 'classroom',
    fontSize: 11,
    fontWeight: boldTypes.has(type) ? 'bold' : 'normal',
    fontStyle: 'normal',
    textAlign: 'center',
    borderWidth: (type === 'facingdirection' || type === 'signature') ? 0 : (type === 'softboard') ? 1 : (type === 'door' || type === 'whiteboard' || type === 'invigdesk') ? 2 : 1,
    borderStyle: (type === 'softboard') ? 'dashed' : 'solid',
    fillColor: options.fillColor ?? getDefaultFillColor(type),
    rotation: 0,
  };

  if (type === 'seat') {
    element.manualSeat = !!options.manualSeat;
  }

  return element;
}

function getDefaultFillColor(type) {
  return type === 'seat' ? '#ffffff' : 'transparent';
}

function wrapPlannerCanvasHtml(content = '') {
  return `<div class="planner-canvas-inner">${content}</div>`;
}

function getPlannerCanvasInner() {
  return document.querySelector('#plannerCanvas .planner-canvas-inner');
}

function getTouchDistance(touches) {
  if (!touches || touches.length < 2) return 0;
  const dx = touches[0].clientX - touches[1].clientX;
  const dy = touches[0].clientY - touches[1].clientY;
  return Math.hypot(dx, dy);
}

function getTouchCenter(touches, rect) {
  return {
    x: ((touches[0].clientX + touches[1].clientX) / 2) - rect.left,
    y: ((touches[0].clientY + touches[1].clientY) / 2) - rect.top,
  };
}

function syncCanvasZoom() {
  const canvas = document.getElementById('plannerCanvas');
  const inner = getPlannerCanvasInner();
  if (!canvas) return;

  if (!inner || !inner.children.length) {
    canvas.style.width = '100%';
    canvas.style.height = '100%';
    return;
  }

  const zoom = clamp(state.canvasZoom || 1, 0.6, 2.5);
  state.canvasZoom = zoom;
  inner.style.transform = `scale(${zoom})`;
  inner.style.transformOrigin = 'top left';

  const baseWidth = Math.max(inner.scrollWidth, inner.offsetWidth, 1);
  const baseHeight = Math.max(inner.scrollHeight, inner.offsetHeight, 1);
  canvas.style.width = `${Math.ceil(baseWidth * zoom)}px`;
  canvas.style.height = `${Math.ceil(baseHeight * zoom)}px`;
}

function setCanvasZoom(nextZoom, anchor = null) {
  const canvasArea = document.getElementById('canvasArea');
  const previousZoom = state.canvasZoom || 1;
  const clampedZoom = clamp(nextZoom, 0.6, 2.5);
  if (Math.abs(clampedZoom - previousZoom) < 0.001) return;

  state.canvasZoom = clampedZoom;

  let unscaledFocus = null;
  if (canvasArea && anchor) {
    unscaledFocus = {
      x: (canvasArea.scrollLeft + anchor.x) / previousZoom,
      y: (canvasArea.scrollTop + anchor.y) / previousZoom,
    };
  }

  syncCanvasZoom();

  if (canvasArea && unscaledFocus) {
    canvasArea.scrollLeft = Math.max(0, unscaledFocus.x * clampedZoom - anchor.x);
    canvasArea.scrollTop = Math.max(0, unscaledFocus.y * clampedZoom - anchor.y);
  }
}

function setupCanvasViewportEvents() {
  const canvasArea = document.getElementById('canvasArea');
  if (!canvasArea) return;

  let pinchState = null;

  canvasArea.addEventListener('touchstart', (event) => {
    if (event.touches.length !== 2 || !getPlannerCanvasInner()?.children.length) return;
    const areaRect = canvasArea.getBoundingClientRect();
    pinchState = {
      distance: getTouchDistance(event.touches),
      zoom: state.canvasZoom || 1,
      center: getTouchCenter(event.touches, areaRect),
    };
    event.preventDefault();
  }, { passive: false });

  canvasArea.addEventListener('touchmove', (event) => {
    if (!pinchState || event.touches.length !== 2) return;
    const areaRect = canvasArea.getBoundingClientRect();
    const nextDistance = getTouchDistance(event.touches);
    if (!nextDistance || !pinchState.distance) return;
    const nextCenter = getTouchCenter(event.touches, areaRect);
    setCanvasZoom(pinchState.zoom * (nextDistance / pinchState.distance), nextCenter);
    pinchState.center = nextCenter;
    event.preventDefault();
  }, { passive: false });

  const clearPinchState = () => {
    pinchState = null;
  };

  canvasArea.addEventListener('touchend', (event) => {
    if (event.touches.length < 2) clearPinchState();
  }, { passive: true });
  canvasArea.addEventListener('touchcancel', clearPinchState, { passive: true });
}

function setupOutsidePageDeselection(plan, root = document) {
  root.addEventListener('pointerdown', (event) => {
    if (event.target.closest('.plan-element, .header-editable, .rh, .rh-rotate, .resize-handle, .seat-width-ruler-preview, .seat-height-ruler-preview')) {
      return;
    }

    const insideDocument = event.target.closest('.exam-document');
    const insideSafeArea = event.target.closest('.page-safe-area');
    if (!insideDocument || !insideSafeArea) {
      selectElement(null, plan);
    }
  }, true);
}

// ============================================================
//  RENDER
// ============================================================
function renderCurrentPlan() {
  const canvas = document.getElementById('plannerCanvas');
  const emptyHint = document.getElementById('emptyHint');

  if (state.previewPlans.length) {
    state.selectedId = null;
    state.selectedHeaderKey = null;
    canvas.innerHTML = wrapPlannerCanvasHtml(state.previewPlans
      .map((plan, index) => `<div class="all-plan-preview-item" data-preview-index="${index}">${buildDocumentHtml(plan)}</div>`)
      .join(''));
    const inner = getPlannerCanvasInner();
    inner.querySelectorAll('.all-plan-preview-item').forEach((item, index) => {
      setupPlanInteractions(state.previewPlans[index], item, {
        editable: state.previewEditable,
        onActivate: state.previewEditable
          ? () => activatePreviewRoom(index)
          : null,
      });
    });
    emptyHint.style.display = 'none';
    syncCanvasZoom();
    return;
  }

  const plan = state.plans[state.currentPlanIndex];
  if (!plan) {
    state.selectedId = null;
    state.selectedHeaderKey = null;
    canvas.innerHTML = '';
    emptyHint.style.display = '';
    syncCanvasZoom();
    return;
  }

  state.selectedId = null;
  state.selectedHeaderKey = null;
  if (!plan.headerStyles) {
    plan.headerStyles = buildHeaderStyles();
  }
  if (!plan.headerLines) {
    plan.headerLines = buildHeaderLines(plan.meta || {});
  }
  canvas.innerHTML = wrapPlannerCanvasHtml(buildDocumentHtml(plan));
  emptyHint.style.display = 'none';
  setupPlanInteractions(plan, getPlannerCanvasInner(), { editable: true });
  syncCanvasZoom();
}

function activatePreviewRoom(index) {
  if (state.mode !== 'multiple') return;
  if (index < 0 || index >= state.previewPlans.length) return;
  state.plans = state.previewPlans;
  state.currentPlanIndex = index;
  if (state.previewMode === 'layouts') {
    state.multiCurrentRoomIndex = index;
    syncCurrentRoomNumberInput();
    updateSavedRoomSelect();
  } else {
    const roomInput = document.getElementById('multiRoomNumberInput');
    if (roomInput) roomInput.value = state.previewPlans[index]?.meta?.room || '';
  }
  renderRoomNav();
}

function buildDocumentHtml(plan) {
  const { classroom } = plan;
  const elHtml = plan.elements.filter((element) => element.region !== 'footer').map(buildElementHtml).join('');
  const footerHtml = plan.elements.filter((element) => element.region === 'footer').map(buildElementHtml).join('');
  const hs = plan.headerStyles || buildHeaderStyles();
  const lines = plan.headerLines || buildHeaderLines(plan.meta || {});
  const seatAreaMetrics = getSeatAreaMetrics(classroom, plan.seatingArea);

  const lineStyle = (cfg) => [
    `font-size:${cfg.fontSize}px`,
    `font-weight:${cfg.fontWeight}`,
    `font-style:${cfg.fontStyle || 'normal'}`,
    `text-align:${cfg.textAlign || 'left'}`,
    `border:${Math.max(0, cfg.borderWidth || 0)}px solid #000`,
  ].join(';');

  return `
    <div class="layout-preview-shell">
      ${buildSeatWidthRulerHtml(classroom, seatAreaMetrics)}
      <div class="layout-preview-body">
        ${buildSeatHeightRulerHtml(classroom, seatAreaMetrics)}
        <div class="exam-document">
          <div class="page-safe-area">
            <div class="doc-header">
              <div class="header-line header-editable school-line" data-header-key="school" style="${lineStyle(hs.school)}">
                <span class="header-value">${escHtml(lines.school || '')}</span>
              </div>
              <div class="header-line header-editable" data-header-key="examSeries" style="${lineStyle(hs.examSeries)}">
                <span class="header-value">${escHtml(lines.examSeries || '')}</span>
              </div>
              <div class="header-line header-editable" data-header-key="paper" style="${lineStyle(hs.paper)}">
                <span class="header-value">${escHtml(lines.paper || '')}</span>
              </div>
              <div class="header-line header-editable" data-header-key="roomNumber" style="${lineStyle(hs.roomNumber)}">
                <span class="header-value">${escHtml(lines.roomNumber || '')}</span>
              </div>
              <div class="header-line header-editable" data-header-key="syllabusCode" style="${lineStyle(hs.syllabusCode)}">
                <span class="header-value">${escHtml(lines.syllabusCode || '')}</span>
              </div>
              <div class="header-line header-editable" data-header-key="componentCode" style="${lineStyle(hs.componentCode)}">
                <span class="header-value">${escHtml(lines.componentCode || '')}</span>
              </div>
            </div>
            <div class="classroom-slot">
              <div class="classroom-shell" style="width:${classroom.width}px;height:${classroom.height}px;">
                <div class="classroom-content">${elHtml}</div>
                <div class="resize-handle" title="Drag to resize classroom"></div>
              </div>
            </div>
            <div class="doc-footer">
              <div class="footer-content">${footerHtml}</div>
            </div>
          </div>
        </div>
      </div>
    </div>`;
}

function buildSeatWidthRulerHtml(classroom, seatAreaMetrics) {
  const activeWidth = Math.max(0, seatAreaMetrics.seatAreaRight - seatAreaMetrics.seatAreaLeft);
  const spanCm = (activeWidth / SEAT_LAYOUT_PIXELS_PER_CM).toFixed(1);

  return `
    <div class="seat-width-ruler-preview" style="width:${classroom.width}px;" title="Drag the left and right handles to define the seat placement area">
      <div class="seat-width-ruler-track"></div>
      <div class="seat-width-ruler-active" style="left:${seatAreaMetrics.seatAreaLeft}px;width:${activeWidth}px;"></div>
      <div class="seat-width-ruler-label">${spanCm} cm</div>
      <div class="seat-width-handle left" data-seat-ruler-handle="left" style="left:${seatAreaMetrics.seatAreaLeft}px;"></div>
      <div class="seat-width-handle right" data-seat-ruler-handle="right" style="left:${seatAreaMetrics.seatAreaRight}px;"></div>
    </div>`;
}

function buildSeatHeightRulerHtml(classroom, seatAreaMetrics) {
  const activeHeight = Math.max(0, seatAreaMetrics.seatAreaBottom - seatAreaMetrics.seatAreaTop);
  const spanCm = (activeHeight / SEAT_LAYOUT_PIXELS_PER_CM).toFixed(1);

  return `
    <div class="seat-height-ruler-preview" style="height:${classroom.height}px;" title="Drag the top and bottom handles to define the seat placement area">
      <div class="seat-height-ruler-track"></div>
      <div class="seat-height-ruler-active" style="top:${seatAreaMetrics.seatAreaTop}px;height:${activeHeight}px;"></div>
      <div class="seat-height-ruler-label">${spanCm} cm</div>
      <div class="seat-height-handle top" data-seat-ruler-handle="top" style="top:${seatAreaMetrics.seatAreaTop}px;"></div>
      <div class="seat-height-handle bottom" data-seat-ruler-handle="bottom" style="top:${seatAreaMetrics.seatAreaBottom}px;"></div>
    </div>`;
}

function buildElementHtml(el) {
  const transform = el.rotation ? `rotate(${el.rotation}deg)` : '';
  const fillColor = el.fillColor ?? getDefaultFillColor(el.type);
  const style = [
    `left:${el.x}px`, `top:${el.y}px`,
    `width:${el.width}px`, `height:${el.height}px`,
    `font-size:${el.fontSize}px`,
    `font-weight:${el.fontWeight}`,
    `font-style:${el.fontStyle || 'normal'}`,
    `text-align:${el.textAlign}`,
    `background-color:${el.type === 'seat' ? 'transparent' : fillColor}`,
    `--seat-fill:${el.type === 'seat' ? fillColor : getDefaultFillColor('seat')}`,
    (el.type === 'seat' || el.type === 'facingdirection' || el.type === 'signature')
      ? 'border:0'
      : `border:${el.borderWidth}px ${el.borderStyle || 'solid'} #000`,
    transform ? `transform:${transform}` : '',
    `transform-origin:center center`,
  ].filter(Boolean).join(';');

  return `<div class="plan-element type-${escHtml(el.type)}" data-id="${el.id}" style="${style}">
    ${buildElementInnerHtml(el)}
    <div class="rh rh-nw" data-resize="nw"></div>
    <div class="rh rh-n"  data-resize="n"></div>
    <div class="rh rh-ne" data-resize="ne"></div>
    <div class="rh rh-e"  data-resize="e"></div>
    <div class="rh rh-se" data-resize="se"></div>
    <div class="rh rh-s"  data-resize="s"></div>
    <div class="rh rh-sw" data-resize="sw"></div>
    <div class="rh rh-w"  data-resize="w"></div>
    <div class="rh-rotate" data-rotate="true" title="Rotate (double-click for +90 degrees, hold Shift for 15-degree steps)"></div>
  </div>`;
}

function buildElementInnerHtml(el) {
  if (el.type === 'seat') return buildSeatHtml(el);
  if (el.type === 'facingdirection') return buildFacingDirectionHtml(el);
  if (el.type === 'signature') return buildSignatureHtml(el);
  return `<div class="el-text">${escHtml(el.label)}</div>`;
}

function buildSeatHtml(el) {
  return `
    <div class="seat-body">
      <div class="seat-caption">CANDIDATE</div>
      <div class="seat-caption">NUMBER</div>
      <div class="el-text seat-number">${escHtml(el.label)}</div>
    </div>
    <div class="seat-chair">CHAIR</div>`;
}

function buildFacingDirectionHtml(el) {
  return `
    <div class="facing-direction-arrow" aria-hidden="true">
      <div class="facing-direction-line"></div>
      <div class="facing-direction-head"></div>
    </div>
    <div class="el-text facing-direction-text">${escHtml(el.label)}</div>`;
}

function buildSignatureHtml(el) {
  return `
    <div class="signature-body">
      <div class="el-text signature-text">${escHtml(el.label)}</div>
      <div class="signature-underline" aria-hidden="true"></div>
    </div>`;
}

// ============================================================
//  INTERACTIONS
// ============================================================
function setupPlanInteractions(plan, root = document, options = {}) {
  const { editable = true, onActivate = null } = options;
  const content = root.querySelector('.classroom-content');
  const footerContent = root.querySelector('.footer-content');
  const shell = root.querySelector('.classroom-shell');
  const slot = root.querySelector('.classroom-slot');
  const resizeHandle = root.querySelector('.classroom-shell > .resize-handle');
  const seatWidthRuler = root.querySelector('.seat-width-ruler-preview');
  const seatHeightRuler = root.querySelector('.seat-height-ruler-preview');

  if (!content || !shell || !slot) return;

  clampClassroomToSlot(shell, slot, plan);
  setupOutsidePageDeselection(plan, root);

  if (typeof onActivate === 'function') {
    root.addEventListener('pointerdown', onActivate, true);
  }

  if (!editable) {
    syncSeatWidthRuler(plan, seatWidthRuler);
    syncSeatHeightRuler(plan, seatHeightRuler, root);
    return;
  }

  content.addEventListener('pointerdown', (e) => {
    if (e.target === content) selectElement(null, plan);
  });
  if (footerContent) {
    footerContent.addEventListener('pointerdown', (e) => {
      if (e.target === footerContent) selectElement(null, plan);
    });
  }

  setupHeaderInteractions(plan, root);

  content.querySelectorAll('.plan-element').forEach((node) => {
    setupElementDrag(node, plan, content);
    setupElementResize(node, plan, content);
    setupElementRotate(node, plan, content);
    setupTextEdit(node, plan);
  });
  if (footerContent) {
    footerContent.querySelectorAll('.plan-element').forEach((node) => {
      setupElementDrag(node, plan, footerContent);
      setupElementResize(node, plan, footerContent);
      setupElementRotate(node, plan, footerContent);
      setupTextEdit(node, plan);
    });
  }

  if (seatWidthRuler) setupSeatWidthRuler(plan, seatWidthRuler, content, seatHeightRuler, root);
  if (seatHeightRuler) {
    setupSeatHeightRuler(plan, seatHeightRuler, content, seatWidthRuler, root);
    syncSeatHeightRuler(plan, seatHeightRuler, root);
  }
  if (resizeHandle) setupClassroomResize(shell, slot, resizeHandle, plan, content, seatWidthRuler, seatHeightRuler, root);
}

function clampClassroomToSlot(shell, slot, plan) {
  const maxW = Math.max(320, slot.clientWidth);
  const maxH = Math.max(260, slot.clientHeight);
  plan.classroom.width = clamp(plan.classroom.width, 320, maxW);
  plan.classroom.height = clamp(plan.classroom.height, 260, maxH);
  shell.style.width = `${plan.classroom.width}px`;
  shell.style.height = `${plan.classroom.height}px`;
}

function setupSeatWidthRuler(plan, ruler, content, heightRuler, root = document) {
  ruler.querySelectorAll('[data-seat-ruler-handle]').forEach((handle) => {
    handle.addEventListener('pointerdown', (event) => {
      event.preventDefault();
      event.stopPropagation();

      if (typeof handle.setPointerCapture === 'function') {
        try {
          handle.setPointerCapture(event.pointerId);
        } catch (error) {
          // Ignore capture errors from synthetic/non-capturable pointers.
        }
      }

      const handleType = handle.dataset.seatRulerHandle;
      const minSpanPx = Math.round(SEAT_LAYOUT_MIN_SPAN_CM * SEAT_LAYOUT_PIXELS_PER_CM);
      const activePointerId = event.pointerId;

      const onMove = (moveEvent) => {
        if (moveEvent.pointerId !== activePointerId) return;
        const rulerRect = ruler.getBoundingClientRect();
        const metrics = getSeatAreaMetrics(plan.classroom, plan.seatingArea);
        const pointerX = clamp(moveEvent.clientX - rulerRect.left, metrics.sideMarginPx, plan.classroom.width - metrics.sideMarginPx);

        if (handleType === 'left') {
          const nextLeft = clamp(pointerX, metrics.sideMarginPx, metrics.seatAreaRight - minSpanPx);
          plan.seatingArea.leftInsetPx = Math.round(nextLeft - metrics.sideMarginPx);
        } else {
          const nextRight = clamp(pointerX, metrics.seatAreaLeft + minSpanPx, plan.classroom.width - metrics.sideMarginPx);
          plan.seatingArea.rightInsetPx = Math.round((plan.classroom.width - metrics.sideMarginPx) - nextRight);
        }

        relayoutSeats(plan, content, ruler, heightRuler, root);
      };

      const cleanup = () => {
        handle.removeEventListener('pointermove', onMove);
        handle.removeEventListener('pointerup', onUp);
        handle.removeEventListener('pointercancel', onUp);
        handle.removeEventListener('lostpointercapture', onLostPointerCapture);
        if (typeof handle.releasePointerCapture === 'function' && handle.hasPointerCapture?.(activePointerId)) {
          try {
            handle.releasePointerCapture(activePointerId);
          } catch (error) {
            // Ignore capture release errors once the pointer lifecycle is complete.
          }
        }
      };

      const onUp = (upEvent) => {
        if (upEvent.pointerId !== activePointerId) return;
        cleanup();
        pushHistory();
      };

      const onLostPointerCapture = () => {
        cleanup();
      };

      handle.addEventListener('pointermove', onMove);
      handle.addEventListener('pointerup', onUp);
      handle.addEventListener('pointercancel', onUp);
      handle.addEventListener('lostpointercapture', onLostPointerCapture);
    });
  });
}

function setupSeatHeightRuler(plan, ruler, content, widthRuler, root = document) {
  ruler.querySelectorAll('[data-seat-ruler-handle]').forEach((handle) => {
    handle.addEventListener('pointerdown', (event) => {
      event.preventDefault();
      event.stopPropagation();

      if (typeof handle.setPointerCapture === 'function') {
        try {
          handle.setPointerCapture(event.pointerId);
        } catch (error) {
          // Ignore capture errors from synthetic/non-capturable pointers.
        }
      }

      const handleType = handle.dataset.seatRulerHandle;
      const minSpanPx = Math.round(SEAT_LAYOUT_MIN_SPAN_CM * SEAT_LAYOUT_PIXELS_PER_CM);
      const activePointerId = event.pointerId;

      const onMove = (moveEvent) => {
        if (moveEvent.pointerId !== activePointerId) return;
        const rulerRect = ruler.getBoundingClientRect();
        const metrics = getSeatAreaMetrics(plan.classroom, plan.seatingArea);
        const pointerY = clamp(moveEvent.clientY - rulerRect.top, metrics.topMarginPx, plan.classroom.height - metrics.bottomMarginPx);

        if (handleType === 'top') {
          const nextTop = clamp(pointerY, metrics.topMarginPx, metrics.seatAreaBottom - minSpanPx);
          plan.seatingArea.topInsetPx = Math.round(nextTop - metrics.topMarginPx);
        } else {
          const nextBottom = clamp(pointerY, metrics.seatAreaTop + minSpanPx, plan.classroom.height - metrics.bottomMarginPx);
          plan.seatingArea.bottomInsetPx = Math.round((plan.classroom.height - metrics.bottomMarginPx) - nextBottom);
        }

        relayoutSeats(plan, content, widthRuler, ruler, root);
      };

      const cleanup = () => {
        handle.removeEventListener('pointermove', onMove);
        handle.removeEventListener('pointerup', onUp);
        handle.removeEventListener('pointercancel', onUp);
        handle.removeEventListener('lostpointercapture', onLostPointerCapture);
        if (typeof handle.releasePointerCapture === 'function' && handle.hasPointerCapture?.(activePointerId)) {
          try {
            handle.releasePointerCapture(activePointerId);
          } catch (error) {
            // Ignore capture release errors once the pointer lifecycle is complete.
          }
        }
      };

      const onUp = (upEvent) => {
        if (upEvent.pointerId !== activePointerId) return;
        cleanup();
        pushHistory();
      };

      const onLostPointerCapture = () => {
        cleanup();
      };

      handle.addEventListener('pointermove', onMove);
      handle.addEventListener('pointerup', onUp);
      handle.addEventListener('pointercancel', onUp);
      handle.addEventListener('lostpointercapture', onLostPointerCapture);
    });
  });
}

function relayoutSeats(plan, content, widthRuler, heightRuler, root = document) {
  if (!plan?.seatingConfig || !content) return;

  const persistedElements = plan.elements.filter((element) => element.type !== 'seat' || element.manualSeat);
  const manualSeats = getOrderedManualSeats(persistedElements);
  manualSeats.forEach((seat) => {
    clampElementIntoConstraints(seat, plan, content);
  });
  if (plan.seatingConfig.manualSeatMode) {
    applyCandidatesToManualSeats(manualSeats, plan.seatingConfig.candidates);
  }
  const autoSeats = plan.seatingConfig.manualSeatMode
    ? []
    : buildSeatElements(
      plan.seatingConfig.candidates,
      plan.seatingConfig.columns,
      plan.seatingConfig.rows,
      plan.classroom,
      { seatingArea: plan.seatingArea }
    );
  const seatElements = [...manualSeats, ...autoSeats];

  plan.elements = [...persistedElements, ...autoSeats];

  if (state.selectedId && !plan.elements.some((element) => element.id === state.selectedId)) {
    state.selectedId = null;
  }

  content.querySelectorAll('.plan-element.type-seat').forEach((node) => node.remove());
  content.insertAdjacentHTML('beforeend', seatElements.map(buildElementHtml).join(''));
  content.querySelectorAll('.plan-element.type-seat').forEach((node) => {
    setupElementDrag(node, plan, content);
    setupElementResize(node, plan, content);
    setupElementRotate(node, plan, content);
    setupTextEdit(node, plan);
  });
  syncElementOrder(content, plan);

  syncSeatWidthRuler(plan, widthRuler);
  syncSeatHeightRuler(plan, heightRuler, root);
  syncToolbarSelection(plan);
}

function syncSeatWidthRuler(plan, ruler) {
  if (!plan || !ruler) return;
  const metrics = getSeatAreaMetrics(plan.classroom, plan.seatingArea);
  const active = ruler.querySelector('.seat-width-ruler-active');
  const label = ruler.querySelector('.seat-width-ruler-label');
  const leftHandle = ruler.querySelector('[data-seat-ruler-handle="left"]');
  const rightHandle = ruler.querySelector('[data-seat-ruler-handle="right"]');
  const activeWidth = Math.max(0, metrics.seatAreaRight - metrics.seatAreaLeft);

  ruler.style.width = `${plan.classroom.width}px`;
  if (active) {
    active.style.left = `${metrics.seatAreaLeft}px`;
    active.style.width = `${activeWidth}px`;
  }
  if (label) label.textContent = `${(activeWidth / SEAT_LAYOUT_PIXELS_PER_CM).toFixed(1)} cm`;
  if (leftHandle) leftHandle.style.left = `${metrics.seatAreaLeft}px`;
  if (rightHandle) rightHandle.style.left = `${metrics.seatAreaRight}px`;
}

function syncSeatHeightRuler(plan, ruler, root = document) {
  if (!plan || !ruler) return;
  const metrics = getSeatAreaMetrics(plan.classroom, plan.seatingArea);
  const shell = root.querySelector('.classroom-shell');
  const doc = root.querySelector('.exam-document');
  const active = ruler.querySelector('.seat-height-ruler-active');
  const label = ruler.querySelector('.seat-height-ruler-label');
  const topHandle = ruler.querySelector('[data-seat-ruler-handle="top"]');
  const bottomHandle = ruler.querySelector('[data-seat-ruler-handle="bottom"]');
  const activeHeight = Math.max(0, metrics.seatAreaBottom - metrics.seatAreaTop);

  ruler.style.height = `${plan.classroom.height}px`;
  if (shell && doc) {
    const shellRect = shell.getBoundingClientRect();
    const docRect = doc.getBoundingClientRect();
    ruler.style.marginTop = `${Math.max(0, shellRect.top - docRect.top)}px`;
  }
  if (active) {
    active.style.top = `${metrics.seatAreaTop}px`;
    active.style.height = `${activeHeight}px`;
  }
  if (label) label.textContent = `${(activeHeight / SEAT_LAYOUT_PIXELS_PER_CM).toFixed(1)} cm`;
  if (topHandle) topHandle.style.top = `${metrics.seatAreaTop}px`;
  if (bottomHandle) bottomHandle.style.top = `${metrics.seatAreaBottom}px`;
}

function selectElement(id, plan) {
  state.selectedId = id;
  state.selectedHeaderKey = null;
  document.querySelectorAll('.header-editable').forEach((el) => {
    el.classList.remove('selected-header');
  });
  document.querySelectorAll('.plan-element').forEach((el) => {
    el.classList.toggle('selected', el.dataset.id === id);
  });
  if (plan) syncToolbarSelection(plan);
}

function syncElementOrder(content, plan) {
  if (!content || !plan) return;
  plan.elements.forEach((element) => {
    const node = content.querySelector(`.plan-element[data-id="${element.id}"]`);
    if (node) content.appendChild(node);
  });
}

function setupHeaderInteractions(plan, root = document) {
  const headerLines = root.querySelectorAll('.header-editable');
  headerLines.forEach((line) => {
    line.addEventListener('pointerdown', (event) => {
      event.stopPropagation();
      selectHeader(line.dataset.headerKey, plan);
    });

    const valueNode = line.querySelector('.header-value');
    if (!valueNode) return;

    valueNode.addEventListener('dblclick', (event) => {
      event.stopPropagation();
      valueNode.contentEditable = 'true';
      line.classList.add('editing');
      valueNode.focus();
      const range = document.createRange();
      range.selectNodeContents(valueNode);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    });

    valueNode.addEventListener('blur', () => {
      if (valueNode.contentEditable !== 'true') return;
      valueNode.contentEditable = 'false';
      line.classList.remove('editing');
      const key = line.dataset.headerKey;
      if (!plan.headerLines) {
        plan.headerLines = buildHeaderLines(plan.meta || {});
      }
      if (key && plan.headerLines[key] !== undefined) {
        plan.headerLines[key] = valueNode.textContent.trim();
      }
      pushHistory();
    });

    valueNode.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        valueNode.blur();
      }
      event.stopPropagation();
    });
  });
}

function selectHeader(key, plan) {
  state.selectedHeaderKey = key;
  state.selectedId = null;
  document.querySelectorAll('.plan-element').forEach((el) => {
    el.classList.remove('selected');
  });
  document.querySelectorAll('.header-editable').forEach((el) => {
    el.classList.toggle('selected-header', el.dataset.headerKey === key);
  });
  syncToolbarSelection(plan);
}

function getSelectedStyleTarget(plan) {
  if (!plan) return null;
  if (state.selectedHeaderKey) {
    if (!plan.headerStyles) plan.headerStyles = buildHeaderStyles();
    return { kind: 'header', model: plan.headerStyles[state.selectedHeaderKey], key: state.selectedHeaderKey };
  }
  if (state.selectedId) {
    const model = plan.elements.find((entry) => entry.id === state.selectedId);
    if (model) return { kind: 'element', model };
  }
  return null;
}

function syncToolbarSelection(plan) {
  const target = getSelectedStyleTarget(plan);
  const model = target ? target.model : null;
  document.getElementById('fontSizeSelect').value = model ? String(model.fontSize) : '11';
  document.getElementById('boldBtn').classList.toggle('active', !!(model && model.fontWeight === 'bold'));
  document.getElementById('italicBtn').classList.toggle('active', !!(model && model.fontStyle === 'italic'));
  document.getElementById('borderWidthInput').value = model ? String(model.borderWidth ?? 1) : '1';
  document.getElementById('fillColorInput').value = normalizeToolbarColor(model?.fillColor, model?.type);
}

function normalizeToolbarColor(fillColor, type) {
  const color = fillColor ?? getDefaultFillColor(type);
  if (!color || color === 'transparent') return '#ffffff';
  return color;
}

// ---- Drag ----
function setupElementDrag(node, plan, room) {
  node.addEventListener('pointerdown', (e) => {
    if (e.target.classList.contains('rh') || e.target.classList.contains('rh-rotate')) return;
    if (node.classList.contains('editing')) return;

    e.preventDefault();
    e.stopPropagation();
    selectElement(node.dataset.id, plan);

    const startX = e.clientX, startY = e.clientY;
    const startLeft = parseFloat(node.style.left) || 0;
    const startTop  = parseFloat(node.style.top)  || 0;

    const onMove = (mv) => {
      const dx = mv.clientX - startX;
      const dy = mv.clientY - startY;
      const next = constrainElementRect(
        startLeft + dx,
        startTop + dy,
        node.offsetWidth,
        node.offsetHeight,
        getElementRotation(node, plan),
        room,
        plan,
        plan.elements.find((element) => element.id === node.dataset.id)
      );
      node.style.left = next.left + 'px';
      node.style.top = next.top + 'px';
    };

    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      persistEl(node, plan);
      pushHistory();
    };

    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp, { once: true });
  });
}

// ---- Resize (8 handles, rotation-aware) ----
function setupElementResize(node, plan, room) {
  node.querySelectorAll('.rh').forEach((handle) => {
    handle.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      e.stopPropagation();

      const dir = handle.dataset.resize;
      const startX = e.clientX, startY = e.clientY;
      const startW = node.offsetWidth,  startH = node.offsetHeight;
      const startL = parseFloat(node.style.left) || 0;
      const startT = parseFloat(node.style.top)  || 0;
      const startCenterX = startL + startW / 2;
      const startCenterY = startT + startH / 2;
      const el = plan.elements.find((el2) => el2.id === node.dataset.id);
      const rot = (el && el.rotation) ? el.rotation : 0;

      const onMove = (mv) => {
        const sdx = mv.clientX - startX;
        const sdy = mv.clientY - startY;
        // Rotate delta to element-local space
        const rad = -rot * Math.PI / 180;
        const dx = sdx * Math.cos(rad) - sdy * Math.sin(rad);
        const dy = sdx * Math.sin(rad) + sdy * Math.cos(rad);

        let w = startW;
        let h = startH;
        let centerShiftX = 0;
        let centerShiftY = 0;

        if (dir.includes('e')) {
          w = Math.max(28, startW + dx);
          centerShiftX += (w - startW) / 2;
        }
        if (dir.includes('w')) {
          w = Math.max(28, startW - dx);
          centerShiftX += (startW - w) / 2;
        }
        if (dir.includes('s')) {
          h = Math.max(18, startH + dy);
          centerShiftY += (h - startH) / 2;
        }
        if (dir.includes('n')) {
          h = Math.max(18, startH - dy);
          centerShiftY += (startH - h) / 2;
        }

        const worldRad = rot * Math.PI / 180;
        const worldShiftX = centerShiftX * Math.cos(worldRad) - centerShiftY * Math.sin(worldRad);
        const worldShiftY = centerShiftX * Math.sin(worldRad) + centerShiftY * Math.cos(worldRad);

        let l = startCenterX + worldShiftX - w / 2;
        let t = startCenterY + worldShiftY - h / 2;

        const next = constrainElementRect(
          l,
          t,
          w,
          h,
          rot,
          room,
          plan,
          el
        );
        w = next.width;
        h = next.height;
        l = next.left;
        t = next.top;

        node.style.width  = w + 'px';
        node.style.height = h + 'px';
        node.style.left   = l + 'px';
        node.style.top    = t + 'px';
      };

      const onUp = () => {
        window.removeEventListener('pointermove', onMove);
        window.removeEventListener('pointerup', onUp);
        persistEl(node, plan);
        pushHistory();
      };

      window.addEventListener('pointermove', onMove);
      window.addEventListener('pointerup', onUp, { once: true });
    });
  });
}

// ---- Rotate ----
function setupElementRotate(node, plan, room) {
  const handle = node.querySelector('.rh-rotate');
  if (!handle) return;

  const applyRotation = (rotation) => {
    const el = plan.elements.find((el2) => el2.id === node.dataset.id);
    if (!el) return;

    el.rotation = rotation;
    node.style.transform = `rotate(${rotation}deg)`;

    const next = constrainElementRect(
      parseFloat(node.style.left) || 0,
      parseFloat(node.style.top) || 0,
      node.offsetWidth,
      node.offsetHeight,
      rotation,
      room,
      plan,
      el
    );
    node.style.left = next.left + 'px';
    node.style.top = next.top + 'px';
  };

  handle.addEventListener('dblclick', (e) => {
    e.preventDefault();
    e.stopPropagation();

    const el = plan.elements.find((el2) => el2.id === node.dataset.id);
    if (!el) return;

    applyRotation(normalizeRotation((el.rotation || 0) + 90));
    pushHistory();
  });

  handle.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    e.stopPropagation();

    const el = plan.elements.find((el2) => el2.id === node.dataset.id);
    if (!el) return;

    // Center of element in viewport
    const getCenter = () => {
      const rect = node.getBoundingClientRect();
      return { cx: rect.left + rect.width / 2, cy: rect.top + rect.height / 2 };
    };

    const { cx, cy } = getCenter();
    const startAngle = Math.atan2(e.clientY - cy, e.clientX - cx) * 180 / Math.PI;
    const startRot = el.rotation || 0;

    const onMove = (mv) => {
      const { cx: ccx, cy: ccy } = getCenter();
      const angle = Math.atan2(mv.clientY - ccy, mv.clientX - ccx) * 180 / Math.PI;
      const rawRotation = startRot + angle - startAngle;
      let newRot = rawRotation;

      if (mv.shiftKey) {
        newRot = Math.round(rawRotation / 15) * 15;
      } else {
        newRot = snapRotation(rawRotation);
      }

      applyRotation(newRot);
    };

    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      pushHistory();
    };

    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp, { once: true });
  });
}

function snapRotation(rotation) {
  const normalized = normalizeRotation(rotation);
  const snapTargets = [0, 90, 180, 270, 360];
  const snapThreshold = 8;

  for (const target of snapTargets) {
    if (Math.abs(normalized - target) <= snapThreshold) {
      return target === 360 ? 0 : target;
    }
  }

  return rotation;
}

function normalizeRotation(rotation) {
  return ((rotation % 360) + 360) % 360;
}

// ---- Text edit ----
function setupTextEdit(node, plan) {
  const textNode = node.querySelector('.el-text');
  if (!textNode) return;

  node.addEventListener('dblclick', (e) => {
    e.stopPropagation();
    textNode.contentEditable = 'true';
    node.classList.add('editing');
    textNode.focus();
    const range = document.createRange();
    range.selectNodeContents(textNode);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  });

  textNode.addEventListener('blur', () => {
    if (textNode.contentEditable !== 'true') return;
    textNode.contentEditable = 'false';
    node.classList.remove('editing');
    const el = plan.elements.find((e) => e.id === node.dataset.id);
    if (el) el.label = textNode.textContent.trim() || el.label;
    pushHistory();
  });

  textNode.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); textNode.blur(); }
    e.stopPropagation();
  });
}

// ---- Classroom resize ----
function setupClassroomResize(shell, slot, handle, plan, content, widthRuler, heightRuler, root = document) {
  handle.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX, startY = e.clientY;
    const startW = plan.classroom.width, startH = plan.classroom.height;

    const onMove = (mv) => {
      const maxW = Math.max(320, slot.clientWidth);
      const ruler = widthRuler;
      const rulerGap = ruler ? 8 : 0;
      const maxH = Math.max(260, slot.clientHeight - ((ruler?.offsetHeight || 0) + rulerGap));
      plan.classroom.width  = clamp(startW + mv.clientX - startX, 320, maxW);
      plan.classroom.height = clamp(startH + mv.clientY - startY, 260, maxH);
      shell.style.width  = plan.classroom.width  + 'px';
      shell.style.height = plan.classroom.height + 'px';
      relayoutSeats(
        plan,
        content,
        widthRuler,
        heightRuler,
        root
      );
    };

    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
    };

    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp, { once: true });
  });
}

// ---- Persist element ----
function persistEl(node, plan) {
  const el = plan.elements.find((e) => e.id === node.dataset.id);
  if (!el) return;
  el.width = node.offsetWidth;
  el.height = node.offsetHeight;
  el.x = parseFloat(node.style.left) || 0;
  el.y = parseFloat(node.style.top) || 0;
  clampElementIntoConstraints(el, plan, el.region === 'footer'
    ? document.querySelector('.footer-content')
    : document.querySelector('.classroom-content'));
  node.style.left = `${el.x}px`;
  node.style.top = `${el.y}px`;
}

function getElementRotation(node, plan) {
  const el = plan.elements.find((entry) => entry.id === node.dataset.id);
  return el && typeof el.rotation === 'number' ? el.rotation : 0;
}

function clampElementIntoConstraints(element, plan, room) {
  if (!element || !room) return;
  const next = constrainElementRect(
    element.x,
    element.y,
    element.width,
    element.height,
    element.rotation || 0,
    room,
    plan,
    element
  );
  element.x = next.left;
  element.y = next.top;
  element.width = next.width;
  element.height = next.height;
}

function getElementConstraintArea(plan, room, element) {
  const roomWidth = Math.max(0, room?.clientWidth || 0);
  const roomHeight = Math.max(0, room?.clientHeight || 0);

  if (!(element?.type === 'seat' && element.manualSeat && element.region !== 'footer' && plan?.classroom)) {
    return {
      left: 0,
      top: 0,
      width: roomWidth,
      height: roomHeight,
    };
  }

  const metrics = getSeatAreaMetrics(plan.classroom, plan.seatingArea, plan.seatingConfig?.maxSeatWidthCm);
  return {
    left: metrics.seatAreaLeft,
    top: metrics.seatAreaTop,
    width: Math.max(0, metrics.seatAreaRight - metrics.seatAreaLeft),
    height: Math.max(0, metrics.seatAreaBottom - metrics.seatAreaTop),
  };
}

function constrainElementRect(left, top, width, height, rotation, room, plan = null, element = null) {
  const constraintArea = getElementConstraintArea(plan, room, element);
  const bounds = getRotatedBounds(width, height, rotation);
  const minLeft = constraintArea.left - bounds.minX;
  const maxLeft = constraintArea.left + constraintArea.width - bounds.maxX;
  const minTop = constraintArea.top - bounds.minY;
  const maxTop = constraintArea.top + constraintArea.height - bounds.maxY;
  const clampedLeft = maxLeft < minLeft ? minLeft : clamp(left, minLeft, maxLeft);
  const clampedTop = maxTop < minTop ? minTop : clamp(top, minTop, maxTop);

  return {
    left: clampedLeft,
    top: clampedTop,
    width,
    height,
  };
}

function getRotatedBounds(width, height, rotation) {
  const rad = (rotation || 0) * Math.PI / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const halfWidth = width / 2;
  const halfHeight = height / 2;
  const corners = [
    { x: -halfWidth, y: -halfHeight },
    { x: halfWidth, y: -halfHeight },
    { x: halfWidth, y: halfHeight },
    { x: -halfWidth, y: halfHeight },
  ].map((corner) => ({
    x: corner.x * cos - corner.y * sin + halfWidth,
    y: corner.x * sin + corner.y * cos + halfHeight,
  }));

  return {
    minX: Math.min(...corners.map((corner) => corner.x)),
    maxX: Math.max(...corners.map((corner) => corner.x)),
    minY: Math.min(...corners.map((corner) => corner.y)),
    maxY: Math.max(...corners.map((corner) => corner.y)),
  };
}

// ============================================================
//  TOOLBAR
// ============================================================
function bindToolbarEvents() {
  document.querySelectorAll('.add-btn').forEach((btn) =>
    btn.addEventListener('click', () => addElementToCanvas(btn.dataset.type))
  );
  document.getElementById('fontSizeSelect').addEventListener('change', (e) =>
    applyFormat('fontSize', parseInt(e.target.value, 10))
  );
  document.getElementById('boldBtn').addEventListener('click', () =>
    toggleFormat('fontWeight', 'bold', 'normal')
  );
  document.getElementById('italicBtn').addEventListener('click', () =>
    toggleFormat('fontStyle', 'italic', 'normal')
  );
  document.getElementById('alignLeftBtn').addEventListener('click',   () => applyFormat('textAlign', 'left'));
  document.getElementById('alignCenterBtn').addEventListener('click', () => applyFormat('textAlign', 'center'));
  document.getElementById('alignRightBtn').addEventListener('click',  () => applyFormat('textAlign', 'right'));
  document.getElementById('borderWidthInput').addEventListener('change', (e) =>
    applyFormat('borderWidth', Math.max(0, parseInt(e.target.value, 10) || 0))
  );
  document.getElementById('fillColorInput').addEventListener('input', (e) =>
    applyFormat('fillColor', e.target.value)
  );
  document.getElementById('clearFillBtn').addEventListener('click', () => {
    const plan = state.plans[state.currentPlanIndex];
    const target = getSelectedStyleTarget(plan);
    if (!target || target.kind !== 'element' || !target.model) return;
    applyFormat('fillColor', getDefaultFillColor(target.model.type));
  });
  document.getElementById('undoBtn').addEventListener('click', undo);
  document.getElementById('redoBtn').addEventListener('click', redo);
  document.getElementById('cutBtn').addEventListener('click', cutSelected);
  document.getElementById('copyBtn').addEventListener('click', copySelected);
  document.getElementById('pasteBtn').addEventListener('click', pasteClipboard);
  document.getElementById('sendToBackBtn').addEventListener('click', () => moveSelectedLayer('back'));
  document.getElementById('bringToFrontBtn').addEventListener('click', () => moveSelectedLayer('front'));
  document.getElementById('deleteBtn').addEventListener('click', deleteSelectedElement);
}

function moveSelectedLayer(direction) {
  if (!state.selectedId || state.selectedHeaderKey) return;
  const plan = state.plans[state.currentPlanIndex];
  if (!plan) return;

  const currentIndex = plan.elements.findIndex((element) => element.id === state.selectedId);
  if (currentIndex < 0) return;
  if (direction === 'front' && currentIndex === plan.elements.length - 1) return;
  if (direction === 'back' && currentIndex === 0) return;

  const [selectedElement] = plan.elements.splice(currentIndex, 1);
  if (direction === 'front') plan.elements.push(selectedElement);
  else plan.elements.unshift(selectedElement);

  const content = document.querySelector('.classroom-content');
  if (content) syncElementOrder(content, plan);
  const footerContent = document.querySelector('.footer-content');
  if (footerContent) syncElementOrder(footerContent, plan);
  selectElement(selectedElement.id, plan);
  pushHistory();
}

function moveSelectedElement(dx, dy) {
  if (!state.selectedId || state.selectedHeaderKey) return false;
  const plan = state.plans[state.currentPlanIndex];
  if (!plan) return false;

  const selectedElement = plan.elements.find((element) => element.id === state.selectedId);
  const content = selectedElement?.region === 'footer'
    ? document.querySelector('.footer-content')
    : document.querySelector('.classroom-content');
  const node = document.querySelector(`.plan-element[data-id="${state.selectedId}"]`);
  if (!content || !node) return false;

  const currentLeft = parseFloat(node.style.left) || 0;
  const currentTop = parseFloat(node.style.top) || 0;
  const nextRect = constrainElementRect(
    currentLeft + dx,
    currentTop + dy,
    node.offsetWidth,
    node.offsetHeight,
    getElementRotation(node, plan),
    content,
    plan,
    selectedElement
  );

  if (nextRect.left === currentLeft && nextRect.top === currentTop) return false;

  node.style.left = `${nextRect.left}px`;
  node.style.top = `${nextRect.top}px`;
  persistEl(node, plan);
  return true;
}

function applyFormat(prop, value) {
  const plan = state.plans[state.currentPlanIndex];
  if (!plan) return;

  const target = getSelectedStyleTarget(plan);
  if (!target || !target.model) return;
  target.model[prop] = value;

  if (target.kind === 'element') {
    const node = document.querySelector(`.plan-element[data-id="${state.selectedId}"]`);
    if (node) {
      if (prop === 'fontSize') node.style.fontSize = value + 'px';
      else if (prop === 'fontWeight') node.style.fontWeight = value;
      else if (prop === 'fontStyle') node.style.fontStyle = value;
      else if (prop === 'textAlign') node.style.textAlign = value;
      else if (prop === 'fillColor') {
        if (target.model.type === 'seat') node.style.setProperty('--seat-fill', value);
        else node.style.backgroundColor = value;
      }
      else if (prop === 'borderWidth') {
        const bs = target.model.borderStyle || 'solid';
        node.style.border = `${value}px ${bs} #000`;
      }
    }
  }

  if (target.kind === 'header') {
    const node = document.querySelector(`.header-editable[data-header-key="${target.key}"]`);
    if (node) {
      if (prop === 'fontSize') node.style.fontSize = value + 'px';
      else if (prop === 'fontWeight') node.style.fontWeight = value;
      else if (prop === 'fontStyle') node.style.fontStyle = value;
      else if (prop === 'textAlign') node.style.textAlign = value;
      else if (prop === 'borderWidth') node.style.border = `${value}px solid #000`;
    }
  }

  syncToolbarSelection(plan);
  pushHistory();
}

function toggleFormat(prop, on, off) {
  const plan = state.plans[state.currentPlanIndex];
  const target = getSelectedStyleTarget(plan);
  if (!target || !target.model) return;
  applyFormat(prop, target.model[prop] === on ? off : on);
}

function addElementToCanvas(type) {
  const plan = state.plans[state.currentPlanIndex];
  if (!plan) { setStatus('Generate a plan first, then add elements.'); return; }

  const defaults = getDefaultElementSpec(type);
  const el = makeEl(type, defaults.label, 40, 40, defaults.width, defaults.height, { manualSeat: type === 'seat' });
  plan.elements.push(el);

  const content = document.querySelector('.classroom-content');
  if (!content) return;
  if (type === 'seat' && plan.seatingConfig?.manualSeatMode) {
    relayoutSeats(
      plan,
      content,
      document.querySelector('.seat-width-ruler-preview'),
      document.querySelector('.seat-height-ruler-preview')
    );
    selectElement(el.id, plan);
    pushHistory();
    return;
  }
  content.insertAdjacentHTML('beforeend', buildElementHtml(el));

  const node = content.querySelector(`.plan-element[data-id="${el.id}"]`);
  if (node) {
    setupElementDrag(node, plan, content);
    setupElementResize(node, plan, content);
    setupElementRotate(node, plan, content);
    setupTextEdit(node, plan);
    selectElement(el.id, plan);
  }
  pushHistory();
}

// ============================================================
//  DELETE / CUT / COPY / PASTE
// ============================================================
function deleteSelectedElement() {
  if (state.selectedHeaderKey) {
    setStatus('Header lines cannot be deleted.');
    return;
  }
  if (!state.selectedId) return;
  const plan = state.plans[state.currentPlanIndex];
  if (!plan) return;
  const selectedElement = plan.elements.find((element) => element.id === state.selectedId);
  const relayoutManualSeats = !!(selectedElement?.type === 'seat' && selectedElement.manualSeat && plan.seatingConfig?.manualSeatMode);
  plan.elements = plan.elements.filter((e) => e.id !== state.selectedId);
  const node = document.querySelector(`.plan-element[data-id="${state.selectedId}"]`);
  if (node) node.remove();
  state.selectedId = null;
  syncToolbarToElement(null, null);
  if (relayoutManualSeats) {
    const content = document.querySelector('.classroom-content');
    if (content) {
      relayoutSeats(
        plan,
        content,
        document.querySelector('.seat-width-ruler-preview'),
        document.querySelector('.seat-height-ruler-preview')
      );
    }
  }
  pushHistory();
}

function copySelected() {
  if (state.selectedHeaderKey) {
    setStatus('Copy/Cut/Paste is available for classroom objects only.');
    return;
  }
  if (!state.selectedId) return;
  const plan = state.plans[state.currentPlanIndex];
  if (!plan) return;
  const el = plan.elements.find((e) => e.id === state.selectedId);
  if (el) { state.clipboard = JSON.parse(JSON.stringify(el)); setStatus('Copied.'); }
}

function cutSelected() {
  if (!state.selectedId) return;
  copySelected();
  deleteSelectedElement();
  setStatus('Cut.');
}

function pasteClipboard() {
  if (!state.clipboard) { setStatus('Nothing on clipboard.'); return; }
  const plan = state.plans[state.currentPlanIndex];
  if (!plan) return;

  const copy = JSON.parse(JSON.stringify(state.clipboard));
  copy.id = `el-${++state.idCounter}`;
  copy.x += 20;
  copy.y += 20;
  plan.elements.push(copy);

  const content = document.querySelector('.classroom-content');
  if (!content) return;
  if (copy.type === 'seat' && copy.manualSeat && plan.seatingConfig?.manualSeatMode) {
    relayoutSeats(
      plan,
      content,
      document.querySelector('.seat-width-ruler-preview'),
      document.querySelector('.seat-height-ruler-preview')
    );
    selectElement(copy.id, plan);
    pushHistory();
    setStatus('Pasted.');
    return;
  }
  content.insertAdjacentHTML('beforeend', buildElementHtml(copy));

  const node = content.querySelector(`.plan-element[data-id="${copy.id}"]`);
  if (node) {
    setupElementDrag(node, plan, content);
    setupElementResize(node, plan, content);
    setupElementRotate(node, plan, content);
    setupTextEdit(node, plan);
    selectElement(copy.id, plan);
  }
  pushHistory();
  setStatus('Pasted.');
}

// ============================================================
//  KEYBOARD SHORTCUTS
// ============================================================
function onKeyDown(e) {
  const active = document.activeElement;
  const inInput = active && (active.isContentEditable ||
    active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.tagName === 'SELECT');

  if (e.key === 'Delete' && !inInput) { deleteSelectedElement(); return; }
  if (e.key === 'Escape') {
    state.selectedHeaderKey = null;
    selectElement(null, state.plans[state.currentPlanIndex]);
    return;
  }

  if (!inInput && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) {
    const step = e.shiftKey ? 10 : 1;
    const deltaByKey = {
      ArrowLeft: [-step, 0],
      ArrowRight: [step, 0],
      ArrowUp: [0, -step],
      ArrowDown: [0, step],
    };
    const [dx, dy] = deltaByKey[e.key];
    e.preventDefault();
    if (moveSelectedElement(dx, dy)) pushHistory();
    return;
  }

  if ((e.ctrlKey || e.metaKey) && !inInput) {
    switch (e.key.toLowerCase()) {
      case 'z': e.preventDefault(); undo(); break;
      case 'y': e.preventDefault(); redo(); break;
      case 'c': e.preventDefault(); copySelected(); break;
      case 'x': e.preventDefault(); cutSelected(); break;
      case 'v': e.preventDefault(); pasteClipboard(); break;
    }
  }
}

// ============================================================
//  ROOM NAVIGATION
// ============================================================
function renderRoomNav() {
  const nav = document.getElementById('roomNav');
  const list = document.getElementById('roomList');
  if (state.mode !== 'multiple' || state.multiRoomLayouts.length <= 1) {
    nav.style.display = 'none';
    list.innerHTML = '';
    updateLayoutActionState();
    return;
  }
  nav.style.display = '';
  list.innerHTML = state.multiRoomLayouts
    .map((p, i) =>
      `<button class="room-nav-btn${i === state.multiCurrentRoomIndex ? ' active' : ''}" data-index="${i}">${escHtml(p.meta.room)}</button>`
    ).join('');
  list.querySelectorAll('.room-nav-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.previewPlans = [];
      state.multiCurrentRoomIndex = parseInt(btn.dataset.index, 10);
      activateMultiLayoutView(state.multiCurrentRoomIndex);
    });
  });
  updateLayoutActionState();
}

function updateLayoutActionState() {
  const saveRoomLayoutBtn = document.getElementById('saveRoomLayoutBtn');
  const deleteRoomLayoutBtn = document.getElementById('deleteRoomLayoutBtn');
  const saveAllRoomLayoutsBtn = document.getElementById('saveAllRoomLayoutsBtn');
  const previewMultiLayoutsBtn = document.getElementById('previewMultiLayoutsBtn');
  const savePreviewEditsBtn = document.getElementById('savePreviewEditsBtn');
  const loadRoomLayoutBtn = document.getElementById('loadRoomLayoutBtn');
  const roomInput = document.getElementById('multiRoomNumberInput');
  if (!saveRoomLayoutBtn || !deleteRoomLayoutBtn || !saveAllRoomLayoutsBtn || !previewMultiLayoutsBtn || !savePreviewEditsBtn || !loadRoomLayoutBtn || !roomInput) return;

  const hasLayouts = state.multiRoomLayouts.length > 0;
  const hasRoomNumber = !!roomInput.value.trim();
  saveRoomLayoutBtn.disabled = !hasRoomNumber;
  deleteRoomLayoutBtn.disabled = !hasLayouts;
  saveAllRoomLayoutsBtn.disabled = !hasLayouts;
  previewMultiLayoutsBtn.disabled = !hasLayouts;
  savePreviewEditsBtn.disabled = !(state.previewMode === 'layouts' && state.previewEditable && state.previewPlans.length);
  loadRoomLayoutBtn.disabled = false;
}

function saveLayouts(scope) {
  if (!state.multiRoomLayouts.length) {
    setStatus('Generate or load a layout first.');
    return;
  }

  const payload = buildLayoutsPayload(scope);
  const roomLabel = scope === 'current'
    ? slugifyFilePart(state.plans[state.currentPlanIndex]?.meta?.room || 'room-layout')
    : 'all-rooms';
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `exam-seating-${roomLabel}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  setStatus(scope === 'current' ? 'Current room layout saved.' : 'All room layouts saved.');
}

function buildLayoutsPayload(scope) {
  const plans = scope === 'current'
    ? [state.multiRoomLayouts[state.multiCurrentRoomIndex]]
    : state.multiRoomLayouts;

  return {
    type: 'exam-seating-layouts',
    version: 1,
    scope,
    mode: 'multiple',
    currentPlanIndex: scope === 'current' ? 0 : state.multiCurrentRoomIndex,
    savedAt: new Date().toISOString(),
    plans: JSON.parse(JSON.stringify(plans)),
  };
}

async function handleLoadLayoutsFile(event) {
  const [file] = event.target.files || [];
  event.target.value = '';
  if (!file) return;

  try {
    const text = await file.text();
    const payload = JSON.parse(text);
    loadLayoutsPayload(payload);
  } catch (err) {
    console.error(err);
    setStatus('Could not load the saved layout file.');
  }
}

function loadLayoutsPayload(payload) {
  if (!payload || payload.type !== 'exam-seating-layouts' || !Array.isArray(payload.plans) || !payload.plans.length) {
    setStatus('This saved layout file is not valid.');
    return;
  }

  const loadedPlans = payload.plans.map(normalizeLoadedPlan).filter(Boolean);
  if (!loadedPlans.length) {
    setStatus('No valid room layouts were found in the saved file.');
    return;
  }

  state.mode = 'multiple';
  state.multiRoomLayouts = loadedPlans;
  state.plans = state.multiRoomLayouts;
  state.multiCurrentRoomIndex = clamp(parseInt(payload.currentPlanIndex, 10) || 0, 0, loadedPlans.length - 1);
  state.currentPlanIndex = state.multiCurrentRoomIndex;
  state.previewPlans = [];
  state.selectedId = null;
  state.selectedHeaderKey = null;
  state.history = { stack: [], pointer: -1 };
  state.idCounter = getMaxElementId(state.multiRoomLayouts);

  showStep('step-multiple');
  activateMultiLayoutView(state.multiCurrentRoomIndex);
  pushHistory();
  setStatus(`Loaded ${loadedPlans.length} saved room layout(s).`);
}

function normalizeLoadedPlan(plan) {
  if (!plan || typeof plan !== 'object') return null;

  const loadedPlan = JSON.parse(JSON.stringify(plan));
  loadedPlan.meta = loadedPlan.meta || {};
  loadedPlan.classroom = loadedPlan.classroom || { width: 690, height: 740 };
  loadedPlan.headerLines = loadedPlan.headerLines || buildHeaderLines(loadedPlan.meta);
  loadedPlan.headerStyles = loadedPlan.headerStyles || buildHeaderStyles();
  loadedPlan.seatingConfig = loadedPlan.seatingConfig || { candidates: [], columns: 4, rows: 4, manualSeatMode: false };
  loadedPlan.seatingArea = loadedPlan.seatingArea || buildInitialSeatingArea(loadedPlan.classroom, loadedPlan.seatingConfig.maxSeatWidthCm);
  loadedPlan.elements = Array.isArray(loadedPlan.elements)
    ? loadedPlan.elements.map(normalizeLoadedElement)
    : buildPersistentLayout(null);

  return loadedPlan;
}

function normalizeLoadedElement(element) {
  const loadedElement = JSON.parse(JSON.stringify(element || {}));
  loadedElement.id = typeof loadedElement.id === 'string' ? loadedElement.id : `el-${++state.idCounter}`;
  loadedElement.type = loadedElement.type || 'text';
  loadedElement.label = String(loadedElement.label ?? '');
  loadedElement.x = Number(loadedElement.x) || 0;
  loadedElement.y = Number(loadedElement.y) || 0;
  loadedElement.width = Number(loadedElement.width) || 100;
  loadedElement.height = Number(loadedElement.height) || 50;
  loadedElement.region = loadedElement.region || 'classroom';
  loadedElement.fontSize = Number(loadedElement.fontSize) || 11;
  loadedElement.fontWeight = loadedElement.fontWeight || 'normal';
  loadedElement.fontStyle = loadedElement.fontStyle || 'normal';
  loadedElement.textAlign = loadedElement.textAlign || 'center';
  loadedElement.borderWidth = Number.isFinite(Number(loadedElement.borderWidth)) ? Number(loadedElement.borderWidth) : 1;
  loadedElement.borderStyle = loadedElement.borderStyle || 'solid';
  loadedElement.fillColor = loadedElement.fillColor ?? getDefaultFillColor(loadedElement.type);
  loadedElement.rotation = Number(loadedElement.rotation) || 0;
  if (loadedElement.type === 'seat') loadedElement.manualSeat = !!loadedElement.manualSeat;
  return loadedElement;
}

function getMaxElementId(plans) {
  return plans.reduce((maxId, plan) => Math.max(maxId, ...((plan.elements || []).map((element) => {
    const match = /^el-(\d+)$/.exec(String(element.id || ''));
    return match ? parseInt(match[1], 10) : 0;
  }))), 0);
}

function slugifyFilePart(value) {
  return String(value || 'layout')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'layout';
}

// ============================================================
//  SAMPLE CSV
// ============================================================
function downloadSample() {
  const a = document.createElement('a');
  a.href = new URL('Seating_Plan_ALevel.xlsx', window.location.href).href;
  a.download = 'Seating_Plan_ALevel.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
}

// ============================================================
//  PDF EXPORT
// ============================================================
function getPlansForCurrentExport() {
  if (state.previewPlans.length) return state.previewPlans.filter(Boolean);
  return [state.plans[state.currentPlanIndex]].filter(Boolean);
}

function getRelativeRect(node, rootRect) {
  const rect = node.getBoundingClientRect();
  return {
    x: rect.left - rootRect.left,
    y: rect.top - rootRect.top,
    width: rect.width,
    height: rect.height,
  };
}

function parsePdfColor(color) {
  if (!color || color === 'transparent') return null;
  const value = String(color).trim();

  if (/^#([\da-f]{3}|[\da-f]{6})$/i.test(value)) {
    const hex = value.slice(1);
    const expanded = hex.length === 3
      ? hex.split('').map((part) => part + part).join('')
      : hex;
    return {
      r: parseInt(expanded.slice(0, 2), 16),
      g: parseInt(expanded.slice(2, 4), 16),
      b: parseInt(expanded.slice(4, 6), 16),
    };
  }

  const rgbMatch = value.match(/^rgba?\(([^)]+)\)$/i);
  if (!rgbMatch) return null;
  const parts = rgbMatch[1].split(',').map((part) => Number(part.trim()));
  if (parts.length < 3 || parts.some((part, index) => index < 3 && !Number.isFinite(part))) {
    return null;
  }
  return { r: parts[0], g: parts[1], b: parts[2] };
}

function applyPdfStrokeColor(pdf, color) {
  const rgb = parsePdfColor(color);
  if (!rgb) return false;
  pdf.setDrawColor(rgb.r, rgb.g, rgb.b);
  return true;
}

function applyPdfFillColor(pdf, color) {
  const rgb = parsePdfColor(color);
  if (!rgb) return false;
  pdf.setFillColor(rgb.r, rgb.g, rgb.b);
  return true;
}

function getPdfFontStyle(model = {}) {
  const isBold = model.fontWeight === 'bold';
  const isItalic = model.fontStyle === 'italic';
  if (isBold && isItalic) return 'bolditalic';
  if (isBold) return 'bold';
  if (isItalic) return 'italic';
  return 'normal';
}

function setPdfTextStyle(pdf, model = {}, fontSize = model.fontSize || 11) {
  pdf.setFont('helvetica', getPdfFontStyle(model));
  pdf.setFontSize(fontSize);
  pdf.setTextColor(0, 0, 0);
}

function rotatePoint(x, y, centerX, centerY, angle) {
  if (!angle) return { x, y };
  const radians = angle * Math.PI / 180;
  const dx = x - centerX;
  const dy = y - centerY;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);

  return {
    x: centerX + dx * cos - dy * sin,
    y: centerY + dx * sin + dy * cos,
  };
}

function getRectPoints(x, y, width, height) {
  return [
    { x, y },
    { x: x + width, y },
    { x: x + width, y: y + height },
    { x, y: y + height },
  ];
}

function rotatePoints(points, centerX, centerY, angle) {
  if (!angle) return points;
  return points.map((point) => rotatePoint(point.x, point.y, centerX, centerY, angle));
}

function getPdfPaintMode(hasStroke, hasFill) {
  if (hasStroke && hasFill) return 'DF';
  if (hasFill) return 'F';
  return 'S';
}

function drawPdfPolygon(pdf, points, options = {}) {
  if (!points.length) return;

  const {
    fillColor = null,
    strokeColor = '#000000',
    lineWidth = 1,
    dash = null,
  } = options;

  const hasFill = applyPdfFillColor(pdf, fillColor);
  const hasStroke = applyPdfStrokeColor(pdf, strokeColor);
  pdf.setLineWidth(lineWidth);
  pdf.setLineDashPattern(dash || [], 0);

  const vectors = [];
  for (let index = 1; index < points.length; index++) {
    vectors.push([
      points[index].x - points[index - 1].x,
      points[index].y - points[index - 1].y,
    ]);
  }

  pdf.lines(vectors, points[0].x, points[0].y, [1, 1], getPdfPaintMode(hasStroke, hasFill), true);
  pdf.setLineDashPattern([], 0);
}

function drawPdfRect(pdf, rect, options = {}) {
  const {
    angle = 0,
    rotationCenter = null,
    fillColor = null,
    strokeColor = '#000000',
    lineWidth = 1,
    dash = null,
  } = options;

  const center = rotationCenter || {
    x: rect.x + rect.width / 2,
    y: rect.y + rect.height / 2,
  };
  const points = rotatePoints(getRectPoints(rect.x, rect.y, rect.width, rect.height), center.x, center.y, angle);
  drawPdfPolygon(pdf, points, { fillColor, strokeColor, lineWidth, dash });
}

function drawPdfLine(pdf, start, end, options = {}) {
  const {
    angle = 0,
    rotationCenter = null,
    lineWidth = 1,
    strokeColor = '#000000',
    dash = null,
  } = options;

  const center = rotationCenter || {
    x: (start.x + end.x) / 2,
    y: (start.y + end.y) / 2,
  };
  const rotatedStart = rotatePoint(start.x, start.y, center.x, center.y, angle);
  const rotatedEnd = rotatePoint(end.x, end.y, center.x, center.y, angle);
  applyPdfStrokeColor(pdf, strokeColor);
  pdf.setLineWidth(lineWidth);
  pdf.setLineDashPattern(dash || [], 0);
  pdf.line(rotatedStart.x, rotatedStart.y, rotatedEnd.x, rotatedEnd.y);
  pdf.setLineDashPattern([], 0);
}

function drawPdfCenteredText(pdf, text, x, y, model = {}, options = {}) {
  if (!text) return;
  const {
    fontSize = model.fontSize || 11,
    angle = 0,
    rotationCenter = { x, y },
    align = model.textAlign || 'center',
  } = options;

  setPdfTextStyle(pdf, model, fontSize);
  const point = rotatePoint(x, y, rotationCenter.x, rotationCenter.y, angle);
  pdf.text(String(text), point.x, point.y, { align, angle });
}

function drawPdfTextBlock(pdf, text, rect, model = {}, options = {}) {
  if (!text) return;

  const {
    angle = 0,
    rotationCenter = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
    paddingX = 4,
    paddingY = 4,
    fontSize = model.fontSize || 11,
    valign = 'middle',
    noWrap = false,
    underline = false,
    lineHeightFactor = 1.18,
    textAlign = model.textAlign || 'center',
  } = options;

  setPdfTextStyle(pdf, model, fontSize);
  const maxWidth = Math.max(8, rect.width - paddingX * 2);
  const textLines = noWrap
    ? [String(text)]
    : pdf.splitTextToSize(String(text), maxWidth).flatMap((line) => Array.isArray(line) ? line : [line]);

  if (!textLines.length) return;

  const lineHeight = fontSize * lineHeightFactor;
  const totalHeight = Math.max(fontSize, lineHeight * textLines.length);
  let startY = rect.y + paddingY + fontSize;
  if (valign === 'middle') {
    startY = rect.y + (rect.height - totalHeight) / 2 + fontSize;
  } else if (valign === 'bottom') {
    startY = rect.y + rect.height - totalHeight + fontSize - paddingY;
  }

  textLines.forEach((line, index) => {
    let baseX = rect.x + paddingX;
    if (textAlign === 'center') baseX = rect.x + rect.width / 2;
    if (textAlign === 'right') baseX = rect.x + rect.width - paddingX;
    const baseY = startY + index * lineHeight;
    const point = rotatePoint(baseX, baseY, rotationCenter.x, rotationCenter.y, angle);
    pdf.text(line, point.x, point.y, { align: textAlign, angle });

    if (!underline) return;
    const textWidth = pdf.getTextWidth(line);
    const startX = textAlign === 'center'
      ? baseX - textWidth / 2
      : textAlign === 'right'
        ? baseX - textWidth
        : baseX;
    const underlineStart = rotatePoint(startX, baseY + 1.5, rotationCenter.x, rotationCenter.y, angle);
    const underlineEnd = rotatePoint(startX + textWidth, baseY + 1.5, rotationCenter.x, rotationCenter.y, angle);
    pdf.setLineWidth(Math.max(0.75, fontSize * 0.05));
    pdf.line(underlineStart.x, underlineStart.y, underlineEnd.x, underlineEnd.y);
  });
}

function drawHeaderLineToPdf(pdf, docEl, plan, key) {
  const headerNode = docEl.querySelector(`.header-editable[data-header-key="${key}"]`);
  if (!headerNode) return;

  const docRect = docEl.getBoundingClientRect();
  const rect = getRelativeRect(headerNode, docRect);
  const styles = plan.headerStyles || buildHeaderStyles();
  const lines = plan.headerLines || buildHeaderLines(plan.meta || {});
  const model = styles[key] || buildHeaderStyles()[key];
  const borderWidth = Math.max(0, model.borderWidth || 0);

  if (borderWidth) {
    drawPdfRect(pdf, rect, {
      strokeColor: '#000000',
      lineWidth: borderWidth,
    });
  }

  drawPdfTextBlock(pdf, lines[key] || '', rect, model, {
    paddingX: 4,
    paddingY: 2,
    valign: 'middle',
    underline: key === 'school',
    noWrap: true,
    textAlign: model.textAlign || 'center',
  });
}

function drawGenericElementToPdf(pdf, element, rect) {
  const angle = element.rotation || 0;
  const center = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
  const fillColor = element.type === 'seat' ? '#ffffff' : (element.fillColor ?? getDefaultFillColor(element.type));
  const borderColor = element.type === 'text' ? '#bbbbbb' : '#000000';
  const dash = element.borderStyle === 'dashed' ? [4, 3] : null;

  if (element.borderWidth > 0 || fillColor !== 'transparent') {
    drawPdfRect(pdf, rect, {
      angle,
      rotationCenter: center,
      fillColor,
      strokeColor: element.borderWidth > 0 ? borderColor : null,
      lineWidth: Math.max(0.75, element.borderWidth || 0.75),
      dash,
    });
  }

  drawPdfTextBlock(pdf, element.label, rect, element, {
    angle,
    rotationCenter: center,
    paddingX: 4,
    paddingY: 3,
    textAlign: element.textAlign || 'center',
  });
}

function drawSeatElementToPdf(pdf, element, rect) {
  const angle = element.rotation || 0;
  const center = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
  const bodyRect = {
    x: rect.x,
    y: rect.y,
    width: rect.width,
    height: rect.height * 0.8,
  };
  const chairRect = {
    x: rect.x + rect.width * 0.05,
    y: rect.y + rect.height * 0.72,
    width: rect.width * 0.9,
    height: rect.height * 0.2,
  };

  drawPdfRect(pdf, bodyRect, {
    angle,
    rotationCenter: center,
    fillColor: element.fillColor ?? '#ffffff',
    strokeColor: '#000000',
    lineWidth: 1.5,
  });
  drawPdfRect(pdf, chairRect, {
    angle,
    rotationCenter: center,
    fillColor: '#ffffff',
    strokeColor: '#888888',
    lineWidth: 1.5,
  });

  const captionModel = { ...element, fontWeight: 'bold', textAlign: 'center' };
  drawPdfTextBlock(pdf, 'CANDIDATE', {
    x: bodyRect.x + 4,
    y: bodyRect.y + bodyRect.height * 0.16,
    width: bodyRect.width - 8,
    height: bodyRect.height * 0.16,
  }, captionModel, {
    angle,
    rotationCenter: center,
    fontSize: element.fontSize * 0.95,
    noWrap: true,
  });
  drawPdfTextBlock(pdf, 'NUMBER', {
    x: bodyRect.x + 4,
    y: bodyRect.y + bodyRect.height * 0.3,
    width: bodyRect.width - 8,
    height: bodyRect.height * 0.16,
  }, captionModel, {
    angle,
    rotationCenter: center,
    fontSize: element.fontSize * 0.95,
    noWrap: true,
  });
  drawPdfTextBlock(pdf, element.label, {
    x: bodyRect.x + 6,
    y: bodyRect.y + bodyRect.height * 0.46,
    width: bodyRect.width - 12,
    height: bodyRect.height * 0.22,
  }, { ...element, fontWeight: 'bold', textAlign: 'center' }, {
    angle,
    rotationCenter: center,
    fontSize: element.fontSize * 1.08,
    noWrap: true,
  });
  drawPdfTextBlock(pdf, 'CHAIR', chairRect, { ...element, fontWeight: 'bold', textAlign: 'center' }, {
    angle,
    rotationCenter: center,
    fontSize: element.fontSize * 0.92,
    noWrap: true,
  });
}

function drawSignatureElementToPdf(pdf, element, rect) {
  const angle = element.rotation || 0;
  const center = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
  const textRect = {
    x: rect.x,
    y: rect.y,
    width: rect.width,
    height: Math.max(16, rect.height - 30),
  };

  drawPdfTextBlock(pdf, element.label, textRect, { ...element, textAlign: 'left' }, {
    angle,
    rotationCenter: center,
    paddingX: 0,
    paddingY: 0,
    valign: 'top',
    noWrap: true,
    textAlign: 'left',
  });
  drawPdfLine(pdf, {
    x: rect.x,
    y: rect.y + rect.height - 1.5,
  }, {
    x: rect.x + rect.width,
    y: rect.y + rect.height - 1.5,
  }, {
    angle,
    rotationCenter: center,
    lineWidth: 1.5,
  });
}

function drawFacingDirectionElementToPdf(pdf, element, rect) {
  const angle = element.rotation || 0;
  const center = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
  const arrowTop = rect.y + rect.height * 0.04;
  const arrowBottom = rect.y + rect.height * 0.4;
  const arrowX = rect.x + rect.width / 2;

  drawPdfLine(pdf, { x: arrowX, y: arrowTop + 8 }, { x: arrowX, y: arrowBottom }, {
    angle,
    rotationCenter: center,
    lineWidth: 1.5,
  });

  const headPoints = rotatePoints([
    { x: arrowX, y: arrowTop },
    { x: arrowX - 5, y: arrowTop + 8 },
    { x: arrowX + 5, y: arrowTop + 8 },
  ], center.x, center.y, angle);
  drawPdfPolygon(pdf, headPoints, { fillColor: '#000000', strokeColor: '#000000', lineWidth: 1 });

  const characters = String(element.label || '').split('');
  const charSize = element.fontSize || 11;
  const totalHeight = characters.length * charSize * 0.95;
  const startY = rect.y + rect.height * 0.5 + Math.max(0, (rect.height * 0.42 - totalHeight) / 2) + charSize;

  characters.forEach((character, index) => {
    drawPdfCenteredText(pdf, character, rect.x + rect.width / 2, startY + index * charSize * 0.95, {
      ...element,
      fontWeight: element.fontWeight || 'normal',
    }, {
      fontSize: charSize,
      angle,
      rotationCenter: center,
      align: 'center',
    });
  });
}

function drawElementToPdf(pdf, docEl, element) {
  const node = docEl.querySelector(`.plan-element[data-id="${element.id}"]`);
  if (!node) return;

  const docRect = docEl.getBoundingClientRect();
  const rect = getRelativeRect(node, docRect);

  if (element.type === 'seat') {
    drawSeatElementToPdf(pdf, element, rect);
    return;
  }
  if (element.type === 'signature') {
    drawSignatureElementToPdf(pdf, element, rect);
    return;
  }
  if (element.type === 'facingdirection') {
    drawFacingDirectionElementToPdf(pdf, element, rect);
    return;
  }

  drawGenericElementToPdf(pdf, element, rect);
}

function drawDocumentToPdf(pdf, docEl, plan) {
  const docRect = docEl.getBoundingClientRect();
  const documentStyles = window.getComputedStyle(docEl);
  const paddingLeft = parseFloat(documentStyles.paddingLeft) || 38;
  const paddingTop = parseFloat(documentStyles.paddingTop) || 38;
  const pageBorderRect = {
    x: paddingLeft,
    y: paddingTop,
    width: docRect.width - paddingLeft * 2,
    height: docRect.height - paddingTop * 2,
  };

  pdf.setFillColor(255, 255, 255);
  pdf.rect(0, 0, docRect.width, docRect.height, 'F');
  drawPdfRect(pdf, pageBorderRect, { strokeColor: '#000000', lineWidth: 1.5 });

  ['school', 'examSeries', 'paper', 'roomNumber', 'syllabusCode', 'componentCode']
    .forEach((key) => drawHeaderLineToPdf(pdf, docEl, plan, key));

  const classroomNode = docEl.querySelector('.classroom-shell');
  if (classroomNode) {
    const classroomRect = getRelativeRect(classroomNode, docRect);
    drawPdfRect(pdf, classroomRect, { fillColor: '#ffffff', strokeColor: '#000000', lineWidth: 2 });
  }

  plan.elements
    .filter((element) => element.region !== 'footer')
    .forEach((element) => drawElementToPdf(pdf, docEl, element));

  plan.elements
    .filter((element) => element.region === 'footer')
    .forEach((element) => drawElementToPdf(pdf, docEl, element));
}

async function exportToPdf() {
  const docEls = Array.from(document.querySelectorAll('.exam-document'));
  if (!docEls.length) { setStatus('Generate a plan first.'); return; }
  const exportPlans = getPlansForCurrentExport();
  setStatus('Preparing editable PDF export…');

  // Temporarily hide selection UI
  const selectedNodes = Array.from(document.querySelectorAll('.plan-element.selected'));
  selectedNodes.forEach((node) => node.classList.remove('selected'));

  try {
    const { jsPDF } = window.jspdf;
    const firstRect = docEls[0].getBoundingClientRect();
    const firstOrient = firstRect.width > firstRect.height ? 'landscape' : 'portrait';
    const pdf = new jsPDF({
      orientation: firstOrient,
      unit: 'px',
      format: [firstRect.width, firstRect.height],
      compress: true,
    });

    docEls.forEach((docEl, index) => {
      const rect = docEl.getBoundingClientRect();
      const plan = exportPlans[index] || exportPlans[0];
      if (index > 0) {
        const orient = rect.width > rect.height ? 'landscape' : 'portrait';
        pdf.addPage([rect.width, rect.height], orient);
      }
      drawDocumentToPdf(pdf, docEl, plan);
    });

    const fileName = 'exam-seating-plan-editable.pdf';
    pdf.save(fileName);
    setStatus(`Editable PDF saved as ${fileName}${docEls.length > 1 ? ` with ${docEls.length} pages` : ''}.`);
  } catch (err) {
    console.error(err);
    setStatus('Editable PDF export failed. Check console for details.');
  } finally {
    selectedNodes.forEach((node) => node.classList.add('selected'));
  }
}

// ============================================================
//  UTILITIES
// ============================================================
function setStatus(msg) {
  document.getElementById('statusBox').textContent = msg;
}

function clamp(v, min, max) {
  return Math.min(max, Math.max(min, v));
}

function escHtml(v) {
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
