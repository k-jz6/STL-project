// ============================================
// 履歴管理 (Undo/Redo)
// ============================================
const HistoryManager = {
    stack: [],
    currentIndex: -1,
    limit: 30,
    isRestoring: false,

    init(initialState) {
        this.stack = [JSON.stringify(initialState)];
        this.currentIndex = 0;
        this.updateButtons();
    },

    record(state) {
        if (this.isRestoring) return;
        // 現在位置より先の履歴は破棄
        if (this.currentIndex < this.stack.length - 1) {
            this.stack = this.stack.slice(0, this.currentIndex + 1);
        }
        // 新しい状態を追加
        const json = JSON.stringify(state);
        // 直前と同じなら保存しない
        if (this.stack[this.currentIndex] === json) return;

        this.stack.push(json);
        if (this.stack.length > this.limit) {
            this.stack.shift();
        } else {
            this.currentIndex++;
        }
        this.updateButtons();
    },

    undo() {
        if (this.currentIndex > 0) {
            this.currentIndex--;
            this.performRestore();
        }
    },

    redo() {
        if (this.currentIndex < this.stack.length - 1) {
            this.currentIndex++;
            this.performRestore();
        }
    },

    performRestore() {
        this.isRestoring = true;
        const data = JSON.parse(this.stack[this.currentIndex]);
        
        const scrollContainer = document.querySelector(".gantt-scroll-container");
        const savedScrollLeft = scrollContainer ? scrollContainer.scrollLeft : 0;
        const savedScrollTop = scrollContainer ? scrollContainer.scrollTop : 0;

        restoreFromData(data); 
        updateCurrentPlanEntryFromAppData();
        DataManager.save(appStore);

        if (scrollContainer) {
            scrollContainer.scrollLeft = savedScrollLeft;
            scrollContainer.scrollTop = savedScrollTop;
        }

        this.isRestoring = false;
        this.updateButtons();
    },

    updateButtons() {
        const undoBtn = document.getElementById("undoBtn");
        const redoBtn = document.getElementById("redoBtn");
        if(undoBtn) undoBtn.disabled = (this.currentIndex <= 0);
        if(redoBtn) redoBtn.disabled = (this.currentIndex >= this.stack.length - 1);
    }
};

// ============================================
// データ保存管理
// ============================================
const DataManager = {
    dbName: "GanttAppDB",
    storeName: "appData",
    useLocalStorage: false,
    STORE_KEY: "main",

    async init() {
        try {
            await new Promise((resolve, reject) => {
                const req = indexedDB.open(this.dbName, 1);
                req.onupgradeneeded = (e) => {
                    const db = e.target.result;
                    if (!db.objectStoreNames.contains(this.storeName)) {
                        db.createObjectStore(this.storeName, { keyPath: "id" });
                    }
                };
                req.onsuccess = (e) => { this.db = e.target.result; resolve(); };
                req.onerror = (e) => reject(e);
            });
        } catch (err) {
            console.warn("IndexedDB fallback -> LocalStorage");
            this.useLocalStorage = true;
        }
    },

    async load() {
        if (this.useLocalStorage) {
            const json = localStorage.getItem(this.dbName + "_" + this.STORE_KEY);
            return json ? JSON.parse(json) : null;
        }
        return new Promise((resolve) => {
            const tx = this.db.transaction([this.storeName], "readonly");
            const req = tx.objectStore(this.storeName).get(this.STORE_KEY);
            req.onsuccess = (e) => resolve(e.target.result ? e.target.result.data : null);
            req.onerror = () => resolve(null);
        });
    },

    async save(data) {
        const ind = document.getElementById("statusIndicator");
        ind.style.opacity = 1;
        setTimeout(() => ind.style.opacity = 0, 1500);

        if (this.useLocalStorage) {
            localStorage.setItem(this.dbName + "_" + this.STORE_KEY, JSON.stringify(data));
            return;
        }
        return new Promise((resolve) => {
            const tx = this.db.transaction([this.storeName], "readwrite");
            tx.objectStore(this.storeName).put({ id: this.STORE_KEY, data: data });
            tx.oncomplete = () => resolve();
        });
    }
};

// ============================================
// アプリ状態・定数
// ============================================
const CELL_WIDTH = 28;
const MAIN_LINE_Y = 32;
const MAIN_LABEL_BASE_TOP = MAIN_LINE_Y - 20;
const MAIN_DIVIDER_Y = 66;
const SUB_SCHEDULE_TOP = 103;
const BASE_ROW_HEIGHT = 129;
const COLLAPSED_ROW_HEIGHT = 68;
const SEGMENT_OFFSET_Y = 48; 

const now = new Date();
const todayISO = dateToISO(now);
const defaultStart = new Date(now.getFullYear(), now.getMonth(), 1);
const defaultEnd = new Date(now.getFullYear(), now.getMonth() + 3, 0);
const DEFAULT_TODO_COLUMNS = "項目1, 項目2, 担当, 実施内容, 計画, 実績, メモ";
const LEGACY_TODO_COLUMNS = "項目1, 項目2, 時間, 実施内容, 計画, 実績";
const LEGACY_TODO_COLUMNS_V2 = "項目1, 項目2, 時間, 実施内容, 計画, 実績, メモ";
// 旧バージョンの固定列名 -> 項目列の位置
const LEGACY_ITEM_COLUMN_INDEX = { "項目1": 0, "項目2": 1, "担当": 2, "時間": 2 };
const DEFAULT_FREE_MEMO_HEIGHT = 116;
const MIN_FREE_MEMO_HEIGHT = 38;
const MAX_IMPORT_FILE_BYTES = 2 * 1024 * 1024;
// 行の色（項目欄のみ）の選択肢。文字が読みやすい淡い色に限定する
const ROW_COLOR_PALETTE = [
    { value: "", name: "なし" },
    { value: "#fee2e2", name: "赤" },
    { value: "#ffedd5", name: "橙" },
    { value: "#fef9c3", name: "黄" },
    { value: "#dcfce7", name: "緑" },
    { value: "#e0f2fe", name: "水色" },
    { value: "#dbeafe", name: "青" },
    { value: "#ede9fe", name: "紫" },
    { value: "#fce7f3", name: "桃" }
];
const EXCEL_SHEET_PLAN = "計画";
const EXCEL_SHEET_MEMO = "メモ";
const EXCEL_FILE_PREFIX = "日程表：";
const EXCEL_COL = {
    done: "完",
    hidden: "非表示",
    memo: "項目1メモ",
    type: "M/S",
    comment: "コメント",
    start: "開始予定",
    end: "終了予定",
    progress: "進捗(1-10)"
};
const EXCEL_TYPE = { mainAll: "M全", mainStart: "M始", mainEnd: "M終", milestone: "M", sub: "S" };
const MAX_PLAN_TASKS = 500;
const MAX_TASK_SEGMENTS = 300;
const MAX_MAIN_SCHEDULES = 100;
const MAX_DEPENDENCIES = 2000;
const MAX_STRING_LENGTH = 3000;
const MAX_ITEM_COLUMNS = 9;
const MIN_ITEM_COLUMNS = 1;
const DEFAULT_HEADERS = ["項目1", "項目2", "担当"];
const DEFAULT_COLUMN_WIDTHS = [30, 120, 90, 40];
const GRIP_COLUMN_WIDTH = 30;
const DEFAULT_ITEM_COLUMN_WIDTH = 90;
const FIXED_TODO_COLUMNS = ["実施内容", "計画", "実績", "メモ"];

let appData = {
    projectName: "標準の計画",
    memoFormat: "plain",
    settings: {
        startDate: dateToISO(defaultStart),
        endDate: dateToISO(defaultEnd),
        holidays: [],
        vacations: [],
        showMainLine: true,
        guideMode: "main",
        hideHolidays: false,
        memoCollapsed: false,
        memoWidth: 0
    },
    headers: DEFAULT_HEADERS.slice(),
    todoColumns: DEFAULT_TODO_COLUMNS,
    columnWidths: DEFAULT_COLUMN_WIDTHS.slice(),
    dependencies: [],
    tasks: [],
    memo: "",
    memoHeight: DEFAULT_FREE_MEMO_HEIGHT
};
let appStore = {
    version: 2,
    currentPlanId: null,
    plans: []
};
let currentPlanId = null;

let allTimelineDays = [];
let timelineDays = [];
let taskObjects = [];
let activeTaskId = null;
let activeProgressSegmentId = null; 
let activeProgressTaskId = null;
let selectionMode = 0; 
let isCtrlSelectionMode = false;
let activeDailyValueTarget = null;
let suppressNextClickAfterProgressCancel = false;
let suppressNextClickAfterMenuDismiss = false;

let currentTodoDate = new Date();
let todoSelectionState = false; 
let pendingGuideRefreshFrame = null;
let dependencyDraft = null;
let dependencyMousePosition = null;
let isRestoringData = false;
let isApplyingFreeMemoHeight = false;
let freeMemoResizeObserver = null;

let dragState = {
    isDragging: false,
    type: null,
    taskId: null,
    segId: null,
    milestoneId: null,
    selectedSegRefs: [],
    startX: 0,
    originalLeft: 0,
    originalWidth: 0,
    originalStartDate: null,
    originalEndDate: null,
    el: null
};

// DOM要素
const headerRow = document.getElementById("headerRow");
const rowsContainer = document.getElementById("rowsContainer");
const leftRowsContainer = document.getElementById("leftRows");
const rangeLabel = document.getElementById("rangeLabel");
const ganttRight = document.getElementById("ganttRight");
const freeMemo = document.getElementById("freeMemo");
const freeMemoArea = document.getElementById("freeMemoArea");
const memoToggleBtn = document.getElementById("memoToggleBtn");
const memoSplitter = document.getElementById("memoSplitter");
let isResizingMemo = false;
let memoResizeStartX = 0;
let memoResizeStartWidth = 0;
const showHiddenCheck = document.getElementById("showHiddenCheck");
const guideModeSelect = document.getElementById("guideModeSelect");
const collapseAllSubsBtn = document.getElementById("collapseAllSubsBtn");
const expandAllSubsBtn = document.getElementById("expandAllSubsBtn");
const toggleHolidaysBtn = document.getElementById("toggleHolidaysBtn");
const projectNameInput = document.getElementById("projectNameInput");
const planSelect = document.getElementById("planSelect");
const newPlanBtn = document.getElementById("newPlanBtn");
const deletePlanBtn = document.getElementById("deletePlanBtn");

const contextMenu = document.getElementById("contextMenu");
let contextMenuTargetTaskId = null;

const segmentContextMenu = document.getElementById("segmentContextMenu");
let contextMenuTargetSegId = null;
let contextMenuTargetTaskForSeg = null;
let contextMenuTargetSegAnchor = null;

const dependencyContextMenu = document.getElementById("dependencyContextMenu");
let contextMenuTargetDependencyId = null;

const headerContextMenu = document.getElementById("headerContextMenu");
let contextMenuTargetHeaderIndex = null;

const settingsPanel = document.getElementById("settingsPanel");
const totalRow = document.getElementById("totalRow");
const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

const taskMemoPanel = document.getElementById("taskMemoPanel");
const taskMemoHeader = document.getElementById("taskMemoHeader");
const taskMemoTitle = document.getElementById("taskMemoTitle");
const taskMemoTextarea = document.getElementById("taskMemoTextarea");
const taskMemoCount = document.getElementById("taskMemoCount");
const taskMemoClose = document.getElementById("taskMemoClose");
let memoPanelTaskId = null;
let memoPanelPinned = false;

let leftColumnWidths = DEFAULT_COLUMN_WIDTHS.slice();
let isResizingCol = false;
let resizeColIndex = null;
let resizeStartX = 0;
let resizeStartWidth = 0;

let isResizingRow = false;
let resizeRowTaskId = null;
let resizeStartY = 0;
let resizeStartHeight = 0;

// ============================================
// ヘルパー関数
// ============================================
function nextItemColumnName(existing) {
    for (let n = 1; n <= MAX_ITEM_COLUMNS; n++) {
        const name = "項目" + n;
        if (!existing.includes(name)) return name;
    }
    return "項目";
}

function normalizeHeaders(headers) {
    const source = Array.isArray(headers) ? headers : [];
    const list = [];
    source.slice(0, MAX_ITEM_COLUMNS).forEach(h => {
        const text = (typeof h === "string") ? h.slice(0, 40) : "";
        list.push(text.trim() === "" ? nextItemColumnName(list) : text);
    });
    while (list.length < MIN_ITEM_COLUMNS) list.push(nextItemColumnName(list));
    return list;
}

function normalizeColumnWidths(widths, headerCount) {
    const source = Array.isArray(widths) ? widths : [];
    const gripWidth = Number(source[0]);
    const out = [Number.isFinite(gripWidth) ? gripWidth : GRIP_COLUMN_WIDTH];
    for (let i = 0; i < headerCount; i++) {
        const w = Number(source[i + 1]);
        const fallback = DEFAULT_COLUMN_WIDTHS[i + 1] ?? DEFAULT_ITEM_COLUMN_WIDTH;
        out.push(Number.isFinite(w) ? w : fallback);
    }
    return out;
}

// 旧形式(label1/label2/label3)も読めるようにする
function readTaskLabels(taskData, count) {
    let labels = Array.isArray(taskData?.labels)
        ? taskData.labels.map(v => (typeof v === "string" ? v : ""))
        : [taskData?.label1, taskData?.label2, taskData?.label3].map(v => (typeof v === "string" ? v : ""));
    labels = labels.slice(0, count);
    while (labels.length < count) labels.push("");
    return labels;
}

function getTaskLabelEditables(task) {
    if (!task?.leftRowEl) return [];
    return Array.from(task.leftRowEl.querySelectorAll(".label-cell > .editable"));
}

function getTaskLabels(task) {
    return getTaskLabelEditables(task).map(el => el.textContent);
}

// 貼り付けや編集で入り込んだ装飾(色・太字など)を取り除き、保存内容と同じプレーンテキストに揃える
function stripEditableFormatting(el) {
    if (!el) return;
    const text = el.textContent;
    if (el.children.length > 0 || el.innerHTML !== text) {
        el.textContent = text;
    }
}

// contenteditable に装飾が入らないようにする（貼り付けはプレーンテキスト、離れるときに整形）
function setupPlainTextEditing(el, onBlur) {
    el.addEventListener("paste", (e) => {
        e.preventDefault();
        const text = ((e.clipboardData || window.clipboardData)?.getData("text/plain") || "").replace(/\r?\n/g, " ");
        document.execCommand("insertText", false, text);
    });
    el.addEventListener("blur", () => {
        stripEditableFormatting(el);
        if (onBlur) onBlur();
    });
}

function getHeaderCells() {
    return Array.from(document.querySelectorAll(".left-header .left-header-cell"));
}

function getHeaderTexts() {
    const cells = getHeaderCells();
    if (cells.length === 0) return normalizeHeaders(appData.headers);
    return cells.map(el => el.textContent);
}

function getDefaultTodoColumns(headers) {
    return [...(headers || getHeaderTexts()), ...FIXED_TODO_COLUMNS].join(", ");
}

// 保存済みの列名（旧固定名や改名前の名前）を現在のヘッダー名に読み替える
function normalizeTodoColumns(cols, itemHeaders) {
    const headers = itemHeaders || getHeaderTexts();
    return cols.map(col => {
        if (headers.includes(col)) return col;
        const legacy = LEGACY_ITEM_COLUMN_INDEX[col];
        if (legacy != null && legacy < headers.length) return headers[legacy];
        return col;
    });
}

function todoColumnKey(col, itemHeaders) {
    const headers = itemHeaders || getHeaderTexts();
    const idx = headers.indexOf(col);
    if (idx !== -1) return "item:" + idx;
    if (col === "実施内容") return "desc";
    if (col === "計画") return "plan";
    if (col === "実績") return "actual";
    if (col === "メモ") return "memo";
    const legacy = LEGACY_ITEM_COLUMN_INDEX[col];
    if (legacy != null && legacy < headers.length) return "item:" + legacy;
    return "unknown";
}
function pad2(n) { return String(n).padStart(2, "0"); }
function dateToISO(d) { return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }
function isoToDate(iso) { const [y, m, d] = iso.split("-").map(Number); return new Date(y, m - 1, d); }
function shiftDateStr(str, delta) {
    const d = isoToDate(str);
    d.setDate(d.getDate() + delta);
    return dateToISO(d);
}
function formatTimestamp(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const h = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');
    return `${y}${m}${day}${h}${min}`;
}
function dateToIndex(str) { return timelineDays.findIndex((d) => d.iso === str); }
function centerX(index) { return index * CELL_WIDTH + CELL_WIDTH / 2; }
function dateToVisibleIndexAtOrAfter(str) {
    const exact = dateToIndex(str);
    if (exact !== -1) return exact;
    return timelineDays.findIndex((d) => d.iso >= str);
}
function dateToVisibleIndexAtOrBefore(str) {
    const exact = dateToIndex(str);
    if (exact !== -1) return exact;
    for (let i = timelineDays.length - 1; i >= 0; i--) {
        if (timelineDays[i].iso <= str) return i;
    }
    return -1;
}
function getVisibleRangeIndices(startIso, endIso) {
    const s = startIso <= endIso ? startIso : endIso;
    const e = startIso <= endIso ? endIso : startIso;
    const sIdx = dateToVisibleIndexAtOrAfter(s);
    const eIdx = dateToVisibleIndexAtOrBefore(e);
    if (sIdx === -1 || eIdx === -1 || sIdx > eIdx) return null;
    return { sIdx, eIdx };
}
function shiftDateByVisibleColumns(str, delta) {
    if (appData.settings.hideHolidays === true && timelineDays.length > 0) {
        let idx = dateToIndex(str);
        if (idx === -1) {
            idx = delta >= 0 ? dateToVisibleIndexAtOrAfter(str) : dateToVisibleIndexAtOrBefore(str);
        }
        if (idx !== -1) {
            const nextIdx = Math.max(0, Math.min(timelineDays.length - 1, idx + delta));
            return timelineDays[nextIdx].iso;
        }
    }
    return shiftDateStr(str, delta);
}
function getByteLength(str) {
    let len = 0;
    for (let i = 0; i < str.length; i++) {
        const c = str.charCodeAt(i);
        len += ((c >= 0x0 && c <= 0x7f) || (c >= 0xff61 && c <= 0xff9f)) ? 1 : 2;
    }
    return len;
}

function extractPlainTextFromHTML(html) {
    const temp = document.createElement("div");
    temp.innerHTML = html || "";
    return temp.innerText || temp.textContent || "";
}

function getFreeMemoText() {
    if (!freeMemo) return "";
    return freeMemo.value || "";
}

function setFreeMemoText(text) {
    if (!freeMemo) return;
    freeMemo.value = text || "";
}

// メモ・備考欄の幅の下限と、日程表側に必ず残す幅
const MIN_MEMO_WIDTH = 200;   // style.css の .free-memo-area { min-width } と揃える
const MIN_GANTT_WIDTH = 320;

// 今そのときに指定できるメモ欄の幅の範囲
function getMemoWidthLimits() {
    const mainArea = document.querySelector(".main-area");
    const total = mainArea ? mainArea.getBoundingClientRect().width : window.innerWidth;
    const splitter = memoSplitter ? memoSplitter.getBoundingClientRect().width : 12;
    const max = Math.max(MIN_MEMO_WIDTH, total - splitter - MIN_GANTT_WIDTH);
    return { min: MIN_MEMO_WIDTH, max };
}

// 幅を反映する。0 や未設定なら既定（全体の1/4）に戻す
function applyMemoWidth(width) {
    if (!freeMemoArea) return 0;
    // 閉じているときはCSS側の指定を使うので、個別指定を外す
    if (freeMemoArea.classList.contains("memo-collapsed")) {
        freeMemoArea.style.flex = "";
        freeMemoArea.style.maxWidth = "";
        return Number(width) || 0;
    }
    const raw = Number(width);
    if (!Number.isFinite(raw) || raw <= 0) {
        freeMemoArea.style.flex = "";
        freeMemoArea.style.maxWidth = "";
        return 0;
    }
    const { min, max } = getMemoWidthLimits();
    const next = Math.round(Math.max(min, Math.min(max, raw)));
    freeMemoArea.style.flex = `0 0 ${next}px`;
    freeMemoArea.style.maxWidth = "none";
    return next;
}

function setMemoCollapsed(collapsed, options = {}) {
    const isCollapsed = !!collapsed;
    if (freeMemoArea) freeMemoArea.classList.toggle("memo-collapsed", isCollapsed);
    if (memoSplitter) memoSplitter.style.display = isCollapsed ? "none" : "";
    if (memoToggleBtn) {
        memoToggleBtn.textContent = isCollapsed ? "◀" : "▶";
        memoToggleBtn.title = isCollapsed ? "メモ・備考を開く" : "メモ・備考を閉じる";
        memoToggleBtn.setAttribute("aria-expanded", String(!isCollapsed));
    }
    if (appData?.settings) {
        appData.settings.memoCollapsed = isCollapsed;
        applyMemoWidth(appData.settings.memoWidth);
    }
    if (!options.skipSave) triggerSave();
    scheduleProgressGuideRefresh();
}

function createEmptyPlanData(name = "標準の計画") {
    return {
        projectName: name,
        memoFormat: "plain",
        settings: {
            startDate: dateToISO(defaultStart),
            endDate: dateToISO(defaultEnd),
            holidays: [],
            vacations: [],
            showMainLine: true,
            guideMode: "main",
            hideHolidays: false,
            memoCollapsed: false,
            memoWidth: 0
        },
        headers: DEFAULT_HEADERS.slice(),
        todoColumns: DEFAULT_TODO_COLUMNS,
        columnWidths: DEFAULT_COLUMN_WIDTHS.slice(),
        dependencies: [],
        tasks: [],
        memo: "",
        memoHeight: DEFAULT_FREE_MEMO_HEIGHT
    };
}

function clonePlanData(data) {
    return JSON.parse(JSON.stringify(data));
}

function createPlanEntry(name = "標準の計画", data = null) {
    const planData = data ? clonePlanData(data) : createEmptyPlanData(name);
    const displayName = name || planData.projectName || "標準の計画";
    planData.projectName = displayName;
    return {
        id: "plan_" + Date.now() + "_" + Math.random().toString(36).slice(2),
        name: displayName,
        data: planData
    };
}

function sanitizeString(value, fallback = "", maxLen = MAX_STRING_LENGTH) {
    if (typeof value !== "string") return fallback;
    return value.slice(0, maxLen);
}

function sanitizeIsoDate(value, fallback) {
    if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
    return fallback;
}

function sanitizeStringArray(values, maxItems = 366, maxLen = 32) {
    if (!Array.isArray(values)) return [];
    return values
        .map(v => sanitizeString(v, "", maxLen))
        .filter(v => /^\d{4}-\d{2}-\d{2}$/.test(v))
        .slice(0, maxItems);
}

function sanitizeDependencies(dependencies, tasksById) {
    if (!Array.isArray(dependencies)) return [];
    return dependencies
        .slice(0, MAX_DEPENDENCIES)
        .map(dep => ({
            id: sanitizeString(dep?.id, "dep_" + Date.now() + "_" + Math.random().toString(36).slice(2), 120),
            fromTaskId: sanitizeString(dep?.fromTaskId, "", 120),
            fromMainId: sanitizeString(dep?.fromMainId, "", 120),
            fromAnchor: dep?.fromAnchor === "end" ? "end" : "start",
            toTaskId: sanitizeString(dep?.toTaskId, "", 120),
            toMainId: sanitizeString(dep?.toMainId, "", 120),
            toAnchor: dep?.toAnchor === "end" ? "end" : "start"
        }))
        .filter(dep => {
            const fromTask = tasksById.get(dep.fromTaskId);
            const toTask = tasksById.get(dep.toTaskId);
            if (!fromTask || !toTask) return false;
            const hasFromMain = fromTask.mainSchedules.some(main => main.id === dep.fromMainId);
            const hasToMain = toTask.mainSchedules.some(main => main.id === dep.toMainId);
            return hasFromMain && hasToMain;
        });
}

function sanitizeImportedPlanData(raw) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
        throw new Error("計画データの形式が不正です。");
    }

    const base = createEmptyPlanData(sanitizeString(raw.projectName, "標準の計画", 120));
    const importedHeaders = normalizeHeaders(Array.isArray(raw.headers)
        ? raw.headers.map((h, i) => sanitizeString(h, base.headers[i] || "", 40))
        : base.headers.slice());
    const startDate = sanitizeIsoDate(raw?.settings?.startDate, base.settings.startDate);
    const endDate = sanitizeIsoDate(raw?.settings?.endDate, base.settings.endDate);

    const data = {
        projectName: sanitizeString(raw.projectName, base.projectName, 120),
        memoFormat: "plain",
        settings: {
            startDate,
            endDate,
            holidays: sanitizeStringArray(raw?.settings?.holidays, 366, 16),
            vacations: sanitizeStringArray(raw?.settings?.vacations, 366, 16),
            showMainLine: raw?.settings?.showMainLine !== false,
            guideMode: normalizeGuideMode(raw?.settings?.guideMode, raw?.settings?.showMainLine),
            hideHolidays: raw?.settings?.hideHolidays === true,
            memoCollapsed: raw?.settings?.memoCollapsed === true,
            memoWidth: Number.isFinite(Number(raw?.settings?.memoWidth)) ? Math.max(0, Number(raw.settings.memoWidth)) : 0
        },
        headers: importedHeaders,
        todoColumns: sanitizeString(raw.todoColumns, DEFAULT_TODO_COLUMNS, 500),
        columnWidths: normalizeColumnWidths(raw.columnWidths, importedHeaders.length),
        dependencies: [],
        tasks: [],
        memo: sanitizeString(raw.memo, "", 200000),
        memoHeight: Math.max(MIN_FREE_MEMO_HEIGHT, Number(raw.memoHeight) || DEFAULT_FREE_MEMO_HEIGHT)
    };

    const tasks = Array.isArray(raw.tasks) ? raw.tasks.slice(0, MAX_PLAN_TASKS) : [];
    const tasksById = new Map();
    data.tasks = tasks.map((task, taskIndex) => {
        const mainSchedules = Array.isArray(task?.mainSchedules)
            ? task.mainSchedules.slice(0, MAX_MAIN_SCHEDULES).map((main, mainIndex) => ({
                id: sanitizeString(main?.id, `main_${taskIndex}_${mainIndex}`, 120),
                startDate: sanitizeIsoDate(main?.startDate, startDate),
                endDate: sanitizeIsoDate(main?.endDate, endDate),
                label: sanitizeString(main?.label, "メイン計画", 200),
                startLabel: sanitizeString(main?.startLabel, "", 200),
                endLabel: sanitizeString(main?.endLabel, "", 200),
                progressEndDate: main?.progressEndDate ? sanitizeIsoDate(main.progressEndDate, null) : null,
                milestones: Array.isArray(main?.milestones)
                    ? main.milestones.slice(0, 200).map((ms, msIndex) => ({
                        id: sanitizeString(ms?.id, `ms_${taskIndex}_${mainIndex}_${msIndex}`, 120),
                        date: sanitizeIsoDate(ms?.date, startDate),
                        label: sanitizeString(ms?.label, "マイルストーン", 200)
                    }))
                    : []
            }))
            : [];

        const normalizedTask = {
            id: sanitizeString(task?.id, `task_${taskIndex}`, 120),
            labels: readTaskLabels(task, importedHeaders.length)
                .map(v => sanitizeString(v, "", 200)),
            mainSchedules,
            segments: Array.isArray(task?.segments)
                ? task.segments.slice(0, MAX_TASK_SEGMENTS).map((seg, segIndex) => ({
                    id: sanitizeString(seg?.id, `seg_${taskIndex}_${segIndex}`, 120),
                    startDate: sanitizeIsoDate(seg?.startDate, startDate),
                    endDate: sanitizeIsoDate(seg?.endDate, endDate),
                    type: seg?.type === "point" ? "point" : "range",
                    label: sanitizeString(seg?.label, "", 200),
                    progressEndDate: seg?.progressEndDate ? sanitizeIsoDate(seg.progressEndDate, null) : null,
                    dailyValues: (seg?.dailyValues && typeof seg.dailyValues === "object" && !Array.isArray(seg.dailyValues)) ? seg.dailyValues : {},
                    dailyResults: (seg?.dailyResults && typeof seg.dailyResults === "object" && !Array.isArray(seg.dailyResults)) ? seg.dailyResults : {}
                }))
                : [],
            memo: sanitizeString(task?.memo, "", 3000),
            customHeight: Number.isFinite(Number(task?.customHeight)) ? Number(task.customHeight) : 0,
            isDone: !!task?.isDone,
            isHidden: !!task?.isHidden,
            isSubCollapsed: !!task?.isSubCollapsed,
            color: normalizeRowColor(task?.color)
        };
        tasksById.set(normalizedTask.id, normalizedTask);
        return normalizedTask;
    });

    data.dependencies = sanitizeDependencies(raw.dependencies, tasksById);
    return data;
}

function normalizeLoadedStore(raw) {
    if (!raw) {
        const entry = createPlanEntry("標準の計画");
        return {
            version: 2,
            currentPlanId: entry.id,
            plans: [entry]
        };
    }

    if (Array.isArray(raw.plans)) {
        const plans = raw.plans.map(plan => {
            const planData = plan?.data ? clonePlanData(plan.data) : createEmptyPlanData(plan?.name || "標準の計画");
            const planName = plan?.name || planData.projectName || "標準の計画";
            planData.projectName = planName;
            return {
                id: plan?.id || ("plan_" + Date.now() + "_" + Math.random().toString(36).slice(2)),
                name: planName,
                data: planData
            };
        });
        const fallbackEntry = createPlanEntry("標準の計画");
        const normalizedPlans = plans.length > 0 ? plans : [fallbackEntry];
        const candidateId = raw.currentPlanId;
        const currentId = normalizedPlans.some(plan => plan.id === candidateId)
            ? candidateId
            : normalizedPlans[0].id;
        return {
            version: 2,
            currentPlanId: currentId,
            plans: normalizedPlans
        };
    }

    const migratedEntry = createPlanEntry(raw.projectName || "標準の計画", raw);
    return {
        version: 2,
        currentPlanId: migratedEntry.id,
        plans: [migratedEntry]
    };
}

function getCurrentPlanEntry() {
    return appStore.plans.find(plan => plan.id === currentPlanId) || null;
}

function updateCurrentPlanEntryFromAppData() {
    const entry = getCurrentPlanEntry();
    if (!entry) return;
    entry.name = appData.projectName || "標準の計画";
    entry.data = clonePlanData(appData);
    appStore.currentPlanId = currentPlanId;
}

function refreshPlanSelect() {
    if (!planSelect) return;
    planSelect.innerHTML = "";
    appStore.plans.forEach(plan => {
        const option = document.createElement("option");
        option.value = plan.id;
        option.textContent = plan.name || "標準の計画";
        if (plan.id === currentPlanId) option.selected = true;
        planSelect.appendChild(option);
    });
    if (deletePlanBtn) {
        deletePlanBtn.disabled = appStore.plans.length <= 1;
    }
}

function renameCurrentPlanWithPrompt() {
    const entry = getCurrentPlanEntry();
    if (!entry) return;
    const nextNameRaw = prompt("計画名を変更:", entry.name || appData.projectName || "標準の計画");
    if (nextNameRaw === null) return;
    const nextName = nextNameRaw.trim();
    if (!nextName) return;
    projectNameInput.value = nextName;
    document.title = nextName + " | 工程表";
    triggerSave();
}

function persistStore() {
    updateCurrentPlanEntryFromAppData();
    return DataManager.save(appStore);
}

function switchPlan(planId, options = {}) {
    const { saveCurrent = true, resetHistory = true } = options;
    if (saveCurrent) {
        syncDataModel();
        updateCurrentPlanEntryFromAppData();
    }

    const entry = appStore.plans.find(plan => plan.id === planId);
    if (!entry) return false;

    currentPlanId = entry.id;
    appStore.currentPlanId = entry.id;
    refreshPlanSelect();
    restoreFromData(clonePlanData(entry.data));
    requestAnimationFrame(() => requestAnimationFrame(scrollToToday));
    if (resetHistory) HistoryManager.init(appData);
    return true;
}

// ============================================
// 項目列（左側）の描画・追加・削除
// ============================================
function renderLeftHeader() {
    const header = document.querySelector(".left-header");
    if (!header) return;
    header.querySelectorAll(".left-header-cell").forEach(el => el.remove());

    const headers = normalizeHeaders(appData.headers);
    appData.headers = headers;

    headers.forEach((text, i) => {
        const cell = document.createElement("div");
        cell.className = "left-header-cell";
        cell.id = "lh" + (i + 1);
        cell.contentEditable = "true";
        cell.dataset.colIndex = String(i);
        cell.textContent = text;
        cell.title = "右クリックで項目列を追加・削除できます";

        cell.addEventListener("paste", (e) => {
            e.preventDefault();
            const text = ((e.clipboardData || window.clipboardData)?.getData("text/plain") || "").replace(/\r?\n/g, " ");
            document.execCommand("insertText", false, text);
        });
        cell.addEventListener("blur", () => {
            stripEditableFormatting(cell);
            const gridIndex = i + 1;
            const previousName = appData.headers[i];
            const nextName = cell.textContent.trim();
            if (nextName === "") {
                cell.textContent = previousName;
            } else if (nextName !== previousName) {
                renameTodoColumn(previousName, nextName);
                appData.headers[i] = nextName;
            }
            leftColumnWidths[gridIndex] = clampColumnWidth(gridIndex, leftColumnWidths[gridIndex]);
            applyLeftColumnWidths();
            triggerSave();
        });
        cell.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            e.stopPropagation();
            showHeaderContextMenu(e, i);
        });

        header.appendChild(cell);
    });
}

// ヘッダー名を変えたら、ToDoの出力項目の指定も追随させる
function renameTodoColumn(previousName, nextName) {
    if (!previousName || previousName === nextName) return;
    const input = document.getElementById("todoColumnsInput");
    if (!input) return;
    const cols = input.value.split(",").map(v => v.trim()).filter(Boolean);
    if (!cols.includes(previousName)) return;
    input.value = cols.map(c => (c === previousName ? nextName : c)).join(", ");
    appData.todoColumns = input.value;
}

function addItemColumn(atIndex) {
    syncDataModel();
    if (appData.headers.length >= MAX_ITEM_COLUMNS) {
        alert(`項目列は最大${MAX_ITEM_COLUMNS}列までです。`);
        return;
    }
    const index = Math.max(0, Math.min(appData.headers.length, atIndex));
    appData.headers.splice(index, 0, nextItemColumnName(appData.headers));
    appData.columnWidths.splice(index + 1, 0, DEFAULT_ITEM_COLUMN_WIDTH);
    appData.tasks.forEach(t => {
        if (!Array.isArray(t.labels)) t.labels = [];
        t.labels.splice(index, 0, "");
    });
    const input = document.getElementById("todoColumnsInput");
    if (input) {
        const cols = input.value.split(",").map(v => v.trim()).filter(Boolean);
        const name = appData.headers[index];
        if (!cols.includes(name)) {
            const anchor = index > 0 ? appData.headers[index - 1] : null;
            const at = anchor ? cols.indexOf(anchor) : -1;
            cols.splice(at >= 0 ? at + 1 : (index > 0 ? cols.length : 0), 0, name);
            input.value = cols.join(", ");
            appData.todoColumns = input.value;
        }
    }
    restoreFromData(appData);
    triggerSave();
}

function removeItemColumn(index) {
    syncDataModel();
    if (index < 0 || index >= appData.headers.length) return;
    if (appData.headers.length <= MIN_ITEM_COLUMNS) {
        alert("項目列は最低1列必要です。");
        return;
    }
    const name = appData.headers[index];
    const hasContent = appData.tasks.some(t => (t.labels || [])[index]);
    const warning = hasContent
        ? `項目列「${name}」を削除します。
この列に入力済みの内容もすべて削除されます。よろしいですか？`
        : `項目列「${name}」を削除しますか？`;
    if (!confirm(warning)) return;

    appData.headers.splice(index, 1);
    appData.columnWidths.splice(index + 1, 1);
    appData.tasks.forEach(t => {
        if (Array.isArray(t.labels)) t.labels.splice(index, 1);
    });
    const input = document.getElementById("todoColumnsInput");
    if (input) {
        input.value = input.value.split(",").map(v => v.trim()).filter(Boolean)
            .filter(c => c !== name).join(", ");
        appData.todoColumns = input.value;
    }
    restoreFromData(appData);
    triggerSave();
}

function showHeaderContextMenu(e, index) {
    if (!headerContextMenu) return;
    contextMenuTargetHeaderIndex = index;
    const canAdd = appData.headers.length < MAX_ITEM_COLUMNS;
    const canDelete = appData.headers.length > MIN_ITEM_COLUMNS;
    const addLeft = document.getElementById("ctxHeaderAddLeft");
    const addRight = document.getElementById("ctxHeaderAddRight");
    const del = document.getElementById("ctxHeaderDelete");
    if (addLeft) addLeft.classList.toggle("menu-disabled", !canAdd);
    if (addRight) addRight.classList.toggle("menu-disabled", !canAdd);
    if (del) del.classList.toggle("menu-disabled", !canDelete);

    hideContextMenus();
    headerContextMenu.style.display = "block";
    headerContextMenu.style.left = e.pageX + "px";
    headerContextMenu.style.top = e.pageY + "px";
}

function applyLeftColumnWidths() {
    const cols = leftColumnWidths.map(w => `${w}px`).join(" ");
    const header = document.querySelector(".left-header");
    if (header) header.style.gridTemplateColumns = cols;
    const rows = document.querySelectorAll(".left-row");
    rows.forEach(r => {
        r.style.gridTemplateColumns = cols;
        const handle = r.querySelector(".row-resize-handle");
        if (handle) handle.style.left = `${leftColumnWidths[0]}px`;
    });
    const left = document.querySelector(".gantt-left");
    if (left) {
        const total = leftColumnWidths.reduce((a, b) => a + b, 0);
        left.style.flex = `0 0 ${total}px`;
        left.style.width = `${total}px`;
    }
    updateLeftResizeHandles();
}

function measureMinWidthForHeader(index) {
    const el = (index === 0)
        ? document.getElementById("rowSelectHeader")
        : getHeaderCells()[index - 1];
    if (!el) return 40;
    const text = el.textContent || "";
    const probe = document.createElement("span");
    probe.style.position = "absolute";
    probe.style.visibility = "hidden";
    probe.style.whiteSpace = "nowrap";
    probe.style.fontSize = "11px";
    probe.style.fontWeight = "600";
    probe.textContent = text;
    document.body.appendChild(probe);
    const width = probe.getBoundingClientRect().width;
    document.body.removeChild(probe);
    return Math.ceil(width + 16);
}

function clampColumnWidth(index, width) {
    const min = (index === 0) ? 26 : measureMinWidthForHeader(index);
    const max = (index === 0) ? 60 : 420;
    return Math.max(min, Math.min(max, width));
}

function ensureLeftResizeOverlay() {
    const left = document.querySelector(".gantt-left");
    if (!left) return null;
    let overlay = left.querySelector(".left-resize-overlay");
    if (!overlay) {
        overlay = document.createElement("div");
        overlay.className = "left-resize-overlay";
        left.appendChild(overlay);
    }
    return overlay;
}

function updateLeftResizeHandles() {
    const overlay = ensureLeftResizeOverlay();
    if (!overlay) return;
    overlay.innerHTML = "";
    let acc = 0;
    for (let i = 0; i < leftColumnWidths.length; i++) {
        acc += leftColumnWidths[i];
        const handle = document.createElement("div");
        handle.className = "left-resize-handle";
        handle.style.left = `${acc - 4}px`;
        handle.dataset.colIndex = String(i);
        handle.addEventListener("mousedown", (e) => {
            e.preventDefault();
            e.stopPropagation();
            isResizingCol = true;
            resizeColIndex = i;
            resizeStartX = e.clientX;
            resizeStartWidth = leftColumnWidths[i];
            document.body.style.cursor = "col-resize";
            document.body.style.userSelect = "none";
        });
        overlay.appendChild(handle);
    }
}

function computeTaskBaseHeight(task) {
    if (task?.isSubCollapsed) return COLLAPSED_ROW_HEIGHT;
    if (!timelineDays.length) return BASE_ROW_HEIGHT;
    if (!task.segments || task.segments.length === 0) return BASE_ROW_HEIGHT;

    const taskDates = {};
    task.segments.forEach(seg => seg._lane = 0);
    const sortedSegs = [...task.segments].sort((a, b) => (a.startDate !== b.startDate) ? (a.startDate < b.startDate ? -1 : 1) : (a.endDate < b.endDate ? -1 : 1));
    let maxLaneUsed = 0;
    sortedSegs.forEach(seg => {
        let requiredLane = 0;
        const visibleRange = getVisibleRangeIndices(seg.startDate, seg.endDate);
        if (!visibleRange) return;
        let { sIdx, eIdx } = visibleRange;
        while (true) {
            let overlap = false;
            for (let i = Math.min(sIdx, eIdx); i <= Math.max(sIdx, eIdx); i++) {
                const iso = timelineDays[i].iso;
                if (taskDates[iso] && taskDates[iso].includes(requiredLane)) { overlap = true; break; }
            }
            if (!overlap) break;
            requiredLane++;
        }
        seg._lane = requiredLane;
        maxLaneUsed = Math.max(maxLaneUsed, requiredLane);
        for (let i = Math.min(sIdx, eIdx); i <= Math.max(sIdx, eIdx); i++) {
            const iso = timelineDays[i].iso;
            if (!taskDates[iso]) taskDates[iso] = [];
            taskDates[iso].push(requiredLane);
        }
    });
    const laneCount = maxLaneUsed + 1;
    return Math.max(BASE_ROW_HEIGHT, SUB_SCHEDULE_TOP + ((laneCount - 1) * SEGMENT_OFFSET_Y) + 26);
}

function updateTaskSubToggleButton(task) {
    const btn = task?.leftRowEl?.querySelector(".sub-toggle-btn");
    if (!btn) return;
    const isCollapsed = !!task.isSubCollapsed;
    btn.textContent = isCollapsed ? "＋" : "−";
    btn.title = isCollapsed ? "この行のサブスケジュールを開く" : "この行のサブスケジュールを閉じる";
    btn.setAttribute("aria-label", btn.title);
}

function applyTaskSubScheduleState(task) {
    if (!task) return;
    const isCollapsed = !!task.isSubCollapsed;
    task.leftRowEl.classList.toggle("sub-collapsed", isCollapsed);
    task.rowEl.classList.toggle("sub-collapsed", isCollapsed);
    updateTaskSubToggleButton(task);
}

function setTaskSubScheduleCollapsed(task, collapsed, options = {}) {
    if (!task) return;
    const nextCollapsed = !!collapsed;
    if (task.isSubCollapsed === nextCollapsed) return;
    task.isSubCollapsed = nextCollapsed;
    task.pendingStartIndex = null;
    task.pendingStartDate = null;
    task.pendingStartLane = 0;
    applyTaskSubScheduleState(task);
    renderAllSegments();
    if (!options.skipSave) triggerSave();
}

function setAllTaskSubSchedulesCollapsed(collapsed) {
    let changed = false;
    taskObjects.forEach(task => {
        const nextCollapsed = !!collapsed;
        if (task.isSubCollapsed === nextCollapsed) return;
        task.isSubCollapsed = nextCollapsed;
        task.pendingStartIndex = null;
        task.pendingStartDate = null;
        task.pendingStartLane = 0;
        applyTaskSubScheduleState(task);
        changed = true;
    });
    if (!changed) return;
    renderAllSegments();
    triggerSave();
}

function getSubScheduleLaneFromY(y) {
    const relativeY = y - SUB_SCHEDULE_TOP;
    if (relativeY <= 0) return 0;
    return Math.max(0, Math.floor((relativeY + (SEGMENT_OFFSET_Y / 2)) / SEGMENT_OFFSET_Y));
}

function normalizeMainSchedule(mainSchedule) {
    if (!mainSchedule || !mainSchedule.startDate || !mainSchedule.endDate) return null;
    return {
        id: mainSchedule.id || ("main_" + Date.now() + "_" + Math.random().toString(36).slice(2)),
        startDate: mainSchedule.startDate,
        endDate: mainSchedule.endDate,
        label: typeof mainSchedule.label === "string" ? mainSchedule.label : "",
        startLabel: mainSchedule.startLabel || "",
        endLabel: mainSchedule.endLabel || "",
        progressEndDate: mainSchedule.progressEndDate || null,
        milestones: Array.isArray(mainSchedule.milestones)
            ? mainSchedule.milestones.map(ms => ({
                id: ms.id || ("ms_" + Date.now() + "_" + Math.random().toString(36).slice(2)),
                date: ms.date,
                label: ms.label || "マイルストーン"
            }))
            : []
    };
}

function normalizeMainSchedules(taskData) {
    if (Array.isArray(taskData.mainSchedules)) {
        return taskData.mainSchedules
            .map(normalizeMainSchedule)
            .filter(Boolean)
            .sort((a, b) => a.startDate.localeCompare(b.startDate));
    }
    const legacyMain = normalizeMainSchedule(taskData.mainSchedule);
    return legacyMain ? [legacyMain] : [];
}

function serializeMainSchedule(main) {
    return {
        id: main.id,
        startDate: main.startDate,
        endDate: main.endDate,
        label: main.label || "",
        startLabel: main.startLabel || "",
        endLabel: main.endLabel || "",
        progressEndDate: main.progressEndDate || null,
        milestones: (main.milestones || []).map(ms => ({
            id: ms.id,
            date: ms.date,
            label: ms.label || ""
        }))
    };
}

function findMainScheduleById(task, mainId) {
    return (task.mainSchedules || []).find(main => main.id === mainId) || null;
}

// 実績が終了日まで達しているメイン計画か（＝未実施の部分が残っていない）
function isMainFullyDone(main) {
    if (!main?.progressEndDate) return false;
    return isoToDate(main.progressEndDate).getTime() >= isoToDate(main.endDate).getTime();
}

// 表示されている日の並びの中での位置。その日が無ければ直前の日の位置を返す
function visibleDayIndex(days, iso) {
    const exact = days.indexOf(iso);
    if (exact !== -1) return exact;
    for (let k = days.length - 1; k >= 0; k--) {
        if (days[k] <= iso) return k;
    }
    return 0;
}

// マイルストーンを「開始日のコメント」「終了日のコメント」「途中のコメント」に分ける。
// 端点かどうかは、伸縮する前の端点日付で判定する。
function classifyMainMilestones(main, startIso, endIso) {
    const list = Array.isArray(main?.milestones) ? [...main.milestones] : [];
    list.sort((a, b) => a.date.localeCompare(b.date));
    const startComment = list.find(ms => ms.date === startIso) || null;
    const endComment = list.find(ms => ms !== startComment && ms.date === endIso) || null;
    const others = list.filter(ms => ms !== startComment && ms !== endComment);
    return { startComment, endComment, others };
}

// 途中のコメントは端点に重ねないので、端点ぶんの日数を別に確保する
function requiredMainDayCount(main, startIso = main?.startDate, endIso = main?.endDate) {
    const { startComment, endComment, others } = classifyMainMilestones(main, startIso, endIso);
    if (others.length === 0) return (startComment && endComment) ? 2 : 1;
    return others.length + 2;
}

// これ以上は縮められない終了日（開始日を固定して縮める場合）
function minEndDateForMain(main) {
    return shiftDateByVisibleColumns(main.startDate, requiredMainDayCount(main) - 1);
}

// これ以上は縮められない開始日（終了日を固定して縮める場合）
function maxStartDateForMain(main) {
    return shiftDateByVisibleColumns(main.endDate, -(requiredMainDayCount(main) - 1));
}

// 伸縮後に各コメントを置き直す。
// 端点のコメントは端点に追従させ、途中のコメントは端点を避けて玉突きで内側に詰める。
function layoutMainMilestones(main, previousStart, previousEnd) {
    if (!main || !Array.isArray(main.milestones) || main.milestones.length === 0) return;
    const { startComment, endComment, others } = classifyMainMilestones(main, previousStart, previousEnd);
    if (startComment) startComment.date = main.startDate;
    if (endComment) endComment.date = main.endDate;

    if (others.length > 0) {
        const bounds = [main.startDate, main.endDate, ...others.map(ms => ms.date)];
        const days = listVisibleDays(
            bounds.reduce((a, b) => (a < b ? a : b)),
            bounds.reduce((a, b) => (a > b ? a : b))
        );
        // 途中のコメントが使えるのは端点の内側だけ
        const innerFirst = visibleDayIndex(days, main.startDate) + 1;
        const innerLast = visibleDayIndex(days, main.endDate) - 1;
        if (innerLast >= innerFirst) {
            // 終了日側からはみ出したものを、後ろから順に手前へ詰める
            let limit = innerLast;
            for (let i = others.length - 1; i >= 0; i--) {
                const idx = visibleDayIndex(days, others[i].date);
                if (idx <= limit) break;
                const nextIdx = Math.max(innerFirst, limit);
                others[i].date = days[nextIdx];
                limit = nextIdx - 1;
            }
            // 開始日側からはみ出したものを、前から順に後ろへ詰める
            let floor = innerFirst;
            for (let i = 0; i < others.length; i++) {
                const idx = visibleDayIndex(days, others[i].date);
                if (idx >= floor) break;
                const nextIdx = Math.min(innerLast, floor);
                others[i].date = days[nextIdx];
                floor = nextIdx + 1;
            }
        }
    }

    main.milestones.sort((a, b) => a.date.localeCompare(b.date));
}

function findMainScheduleByDate(task, iso) {
    return (task.mainSchedules || []).find(main => iso >= main.startDate && iso <= main.endDate) || null;
}

function ensureDependenciesArray() {
    if (!Array.isArray(appData.dependencies)) appData.dependencies = [];
    return appData.dependencies;
}

function findDependencyById(depId) {
    return ensureDependenciesArray().find(dep => dep.id === depId) || null;
}

function findEndpointTask(taskId) {
    return taskObjects.find(task => task.id === taskId) || null;
}

function isSameEndpoint(a, b) {
    return !!a && !!b
        && a.taskId === b.taskId
        && a.mainId === b.mainId
        && a.anchor === b.anchor;
}

function hasDuplicateDependency(candidate) {
    return ensureDependenciesArray().some(dep =>
        dep.fromTaskId === candidate.fromTaskId
        && dep.fromMainId === candidate.fromMainId
        && dep.fromAnchor === candidate.fromAnchor
        && dep.toTaskId === candidate.toTaskId
        && dep.toMainId === candidate.toMainId
        && dep.toAnchor === candidate.toAnchor
    );
}

function clearDependencyDraft(shouldRender = true) {
    dependencyDraft = null;
    dependencyMousePosition = null;
    if (shouldRender) renderDependencies();
}

function cleanupDependencies() {
    if (isRestoringData) {
        return ensureDependenciesArray();
    }
    ensureDependenciesArray();
    appData.dependencies = appData.dependencies.filter(dep => {
        if (!dep || !dep.id) return false;
        if (!["start", "end"].includes(dep.fromAnchor) || !["start", "end"].includes(dep.toAnchor)) return false;
        const fromTask = findEndpointTask(dep.fromTaskId);
        const toTask = findEndpointTask(dep.toTaskId);
        if (!fromTask || !toTask) return false;
        if (!findMainScheduleById(fromTask, dep.fromMainId)) return false;
        if (!findMainScheduleById(toTask, dep.toMainId)) return false;
        return !(dep.fromTaskId === dep.toTaskId && dep.fromMainId === dep.toMainId && dep.fromAnchor === dep.toAnchor);
    });

    if (dependencyDraft) {
        const fromTask = findEndpointTask(dependencyDraft.taskId);
        if (!fromTask || !findMainScheduleById(fromTask, dependencyDraft.mainId)) {
            clearDependencyDraft(false);
        }
    }

    return appData.dependencies;
}

function beginDependencyDraft(task, main, anchor) {
    dependencyDraft = {
        taskId: task.id,
        mainId: main.id,
        anchor
    };
    const sourceEl = getMainEndpointElement(dependencyDraft);
    if (sourceEl) {
        const rect = sourceEl.getBoundingClientRect();
        dependencyMousePosition = {
            x: rect.left + (rect.width / 2),
            y: rect.top + (rect.height / 2)
        };
    }
    renderDependencies();
}

function completeDependencyDraft(task, main, anchor) {
    if (!dependencyDraft) return false;
    const targetRef = { taskId: task.id, mainId: main.id, anchor };
    if (isSameEndpoint(dependencyDraft, targetRef)) {
        alert("始点と終点に同じ丸は指定できません。");
        return false;
    }

    const nextDep = {
        id: "dep_" + Date.now() + "_" + Math.random().toString(36).slice(2),
        fromTaskId: dependencyDraft.taskId,
        fromMainId: dependencyDraft.mainId,
        fromAnchor: dependencyDraft.anchor,
        toTaskId: task.id,
        toMainId: main.id,
        toAnchor: anchor
    };

    if (hasDuplicateDependency(nextDep)) {
        alert("同じ依存線はすでに存在します。");
        return false;
    }

    ensureDependenciesArray().push(nextDep);
    clearDependencyDraft(false);
    renderDependencies();
    triggerSave();
    return true;
}

function getMainEndpointElements() {
    return Array.from(rowsContainer.querySelectorAll(".main-point[data-task-id][data-main-id][data-anchor]"));
}

function getMainEndpointElement(ref) {
    return getMainEndpointElements().find(el =>
        el.dataset.taskId === ref.taskId
        && el.dataset.mainId === ref.mainId
        && el.dataset.anchor === ref.anchor
    ) || null;
}

function hasOverlappingMainSchedule(task, startIso, endIso, excludeId = null) {
    return (task.mainSchedules || []).some(main => {
        if (excludeId && main.id === excludeId) return false;
        return !(endIso < main.startDate || startIso > main.endDate);
    });
}

function getTaskTitle(task) {
    const parts = getTaskLabels(task).map(v => v.trim()).filter(Boolean);
    if (parts.length === 0) return "メモ";
    return "メモ: " + parts.join(" / ");
}

function updateMemoCount(text) {
    const len = text.length;
    taskMemoCount.textContent = `${len} / 3000`;
}

function positionMemoPanel(anchorEl) {
    const rect = anchorEl.getBoundingClientRect();
    const panelRect = taskMemoPanel.getBoundingClientRect();
    const padding = 12;
    let left = rect.right + 10;
    let top = rect.top;
    if (left + panelRect.width > window.innerWidth - padding) {
        left = rect.left - panelRect.width - 10;
    }
    if (left < padding) left = padding;
    if (top + panelRect.height > window.innerHeight - padding) {
        top = window.innerHeight - panelRect.height - padding;
    }
    if (top < padding) top = padding;
    taskMemoPanel.style.left = `${left}px`;
    taskMemoPanel.style.top = `${top}px`;
}

function openTaskMemo(task, anchorEl, pin = false) {
    if (!task || !taskMemoPanel) return;
    memoPanelTaskId = task.id;
    memoPanelPinned = pin;
    taskMemoTitle.textContent = getTaskTitle(task);
    const memoText = task.memo || "";
    taskMemoTextarea.value = memoText;
    updateMemoCount(memoText);
    taskMemoPanel.classList.remove("memo-hidden");
    taskMemoPanel.setAttribute("aria-hidden", "false");
    positionMemoPanel(anchorEl || task.leftRowEl);
}

function closeTaskMemoPanel() {
    memoPanelTaskId = null;
    memoPanelPinned = false;
    taskMemoPanel.classList.add("memo-hidden");
    taskMemoPanel.setAttribute("aria-hidden", "true");
}

function clearSegmentSelection() {
    taskObjects.forEach(task => {
        task.segments.forEach(seg => {
            seg.isSelected = false;
        });
    });
}

function getSelectedSubSegments() {
    const refs = [];
    taskObjects.forEach(task => {
        task.segments.forEach(seg => {
            if (seg.isSelected) refs.push({ task, seg });
        });
    });
    return refs;
}

function updateCtrlSelectionMode(enabled) {
    isCtrlSelectionMode = enabled;
    document.body.classList.toggle("ctrl-select-mode", enabled);
}

function isSubSegmentSelectableTarget(target) {
    return !!target.closest("[data-sub-selectable='true']");
}

function cancelActiveProgressSelection() {
    if (!activeProgressSegmentId) return false;
    activeProgressSegmentId = null;
    activeProgressTaskId = null;
    suppressNextClickAfterProgressCancel = true;
    renderAllSegments();
    return true;
}

function hideContextMenus() {
    contextMenu.style.display = "none";
    segmentContextMenu.style.display = "none";
    if (headerContextMenu) headerContextMenu.style.display = "none";
    if (dependencyContextMenu) dependencyContextMenu.style.display = "none";
    if (contextMenuTargetDependencyId) {
        contextMenuTargetDependencyId = null;
        renderDependencies();
    }
}

function dismissContextMenusWithClickSuppression() {
    hideContextMenus();
    suppressNextClickAfterMenuDismiss = true;
}

function isOverlayCancelStateActive() {
    return segmentContextMenu.style.display === "block" ||
        contextMenu.style.display === "block" ||
        (dependencyContextMenu && dependencyContextMenu.style.display === "block") ||
        !!activeProgressSegmentId;
}

function findTaskForActiveProgressSelection() {
    if (!activeProgressSegmentId) return null;
    if (activeProgressTaskId) {
        return taskObjects.find(task => task.id === activeProgressTaskId) || null;
    }
    return taskObjects.find(task => task.segments.some(seg => seg.id === activeProgressSegmentId)) || null;
}

function isValidProgressClickTarget(target) {
    const task = findTaskForActiveProgressSelection();
    if (!task) return false;
    if (!task.rowEl.contains(target)) return false;
    if (!target.closest(".cell")) return false;
    if (target.closest(".segment") || target.closest(".point") || target.closest(".segment-label") || target.closest(".daily-val")) {
        return false;
    }
    return true;
}

function setActiveDailyValueTarget(task, seg, iso) {
    activeDailyValueTarget = {
        taskId: task.id,
        segId: seg.id,
        iso
    };
}

function clearActiveDailyValueTarget() {
    activeDailyValueTarget = null;
}

function getActiveDailyValueRef() {
    if (!activeDailyValueTarget) return null;
    const task = taskObjects.find(t => t.id === activeDailyValueTarget.taskId);
    const seg = task ? task.segments.find(s => s.id === activeDailyValueTarget.segId) : null;
    if (!task || !seg) return null;
    return { task, seg, iso: activeDailyValueTarget.iso };
}

function moveActiveDailyValue(delta) {
    const ref = getActiveDailyValueRef();
    if (!ref) return;
    const nextIso = shiftDateByVisibleColumns(ref.iso, delta);
    if (nextIso < ref.seg.startDate || nextIso > ref.seg.endDate) return;
    if (dateToIndex(nextIso) === -1) return;
    activeDailyValueTarget.iso = nextIso;
    renderAllSegments();
}

function editDailyValue(task, seg, iso) {
    setActiveDailyValueTarget(task, seg, iso);
    const curVal = (seg.dailyValues && seg.dailyValues[iso]) || "";
    let input = prompt("工数 (例: 1, 0.5) または文字:", curVal);
    if (input !== null) {
        input = input.trim();
        if (input === "") {
            if (seg.dailyValues) delete seg.dailyValues[iso];
        } else {
            if (!seg.dailyValues) seg.dailyValues = {};
            if (/^\d(\.\d)?$/.test(input) || (!isNaN(parseFloat(input)) && getByteLength(input) <= 4)) {
                seg.dailyValues[iso] = input;
            } else {
                if (getByteLength(input) <= 4) seg.dailyValues[iso] = input;
                else {
                    alert("全角2文字(半角4文字)以内で入力してください。");
                    renderAllSegments();
                    return;
                }
            }
        }
        renderAllSegments();
        triggerSave();
    } else {
        renderAllSegments();
    }
}

// ============================================
// データ同期 & 保存
// ============================================
function syncDataModel() {
    appData.tasks = taskObjects.map(t => {
        return {
            id: t.id,
            labels: getTaskLabels(t),
            mainSchedule: (t.mainSchedules && t.mainSchedules.length > 0) ? serializeMainSchedule(t.mainSchedules[0]) : null,
            mainSchedules: (t.mainSchedules || []).map(serializeMainSchedule),
            segments: t.segments.map(seg => ({
                id: seg.id,
                startDate: seg.startDate,
                endDate: seg.endDate,
                type: seg.type,
                label: seg.label || "",
                progressEndDate: seg.progressEndDate || null,
                dailyValues: seg.dailyValues ? { ...seg.dailyValues } : {},
                dailyResults: seg.dailyResults ? { ...seg.dailyResults } : {}
            })),
            memo: t.memo || "",
            customHeight: t.customHeight || 0,
            isDone: t.isDone || false,
            isHidden: t.isHidden || false,
            isSubCollapsed: !!t.isSubCollapsed,
            color: t.color || ""
        };
    });
    appData.memo = getFreeMemoText();
    appData.memoFormat = "plain";
    appData.projectName = projectNameInput.value;
    
    appData.headers = getHeaderTexts();
    appData.todoColumns = document.getElementById("todoColumnsInput").value;
    appData.columnWidths = leftColumnWidths.slice();
    appData.dependencies = cleanupDependencies().map(dep => ({ ...dep }));
    appData.projectName = projectNameInput.value || "標準の計画";
    // memoHeight は旧レイアウト（メモ欄が下段）用の値。読み込んだ値をそのまま保持する
    updateCurrentPlanEntryFromAppData();
    refreshPlanSelect();
}

function triggerSave() {
    syncDataModel();
    calculateTotals();
    HistoryManager.record(appData);
    DataManager.save(appStore);
}

projectNameInput.addEventListener("change", () => {
    document.title = projectNameInput.value + " | 工程表";
    triggerSave();
});
freeMemo.addEventListener("input", () => {
    triggerSave();
});


// ============================================
// 初期化 & 復元
// ============================================
async function initializeApp() {
    await DataManager.init();
    appStore = normalizeLoadedStore(await DataManager.load());
    currentPlanId = appStore.currentPlanId;
    refreshPlanSelect();
    const currentEntry = getCurrentPlanEntry();
    restoreFromData(clonePlanData(currentEntry ? currentEntry.data : createEmptyPlanData()));
    setTimeout(scrollToToday, 100);
    HistoryManager.init(appData);
    setupControlEvents();
    await DataManager.save(appStore);
}

function restoreFromData(data) {
    isRestoringData = true;
    appData = data;
    if (!data.settings.startDate) {
        appData.settings.startDate = dateToISO(defaultStart);
        appData.settings.endDate = dateToISO(defaultEnd);
    }
    if (typeof appData.settings.showMainLine !== "boolean") {
        appData.settings.showMainLine = true;
    }
    appData.settings.guideMode = normalizeGuideMode(appData.settings.guideMode, appData.settings.showMainLine);
    appData.settings.showMainLine = appData.settings.guideMode !== "off";
    if (typeof appData.settings.hideHolidays !== "boolean") {
        appData.settings.hideHolidays = false;
    }
    if (!Array.isArray(appData.settings.holidays)) appData.settings.holidays = [];
    if (!Array.isArray(appData.settings.vacations)) appData.settings.vacations = [];
    if (typeof appData.settings.memoCollapsed !== "boolean") {
        appData.settings.memoCollapsed = false;
    }
    if (!Number.isFinite(Number(appData.settings.memoWidth))) {
        appData.settings.memoWidth = 0;
    }
    appData.headers = normalizeHeaders(appData.headers);
    appData.columnWidths = normalizeColumnWidths(appData.columnWidths, appData.headers.length);
    if (!Array.isArray(appData.dependencies)) appData.dependencies = [];
    if (!appData.memoFormat) appData.memoFormat = "legacy-html";
    if (!appData.memoHeight) appData.memoHeight = DEFAULT_FREE_MEMO_HEIGHT;
    
    if (!appData.todoColumns
        || appData.todoColumns === LEGACY_TODO_COLUMNS
        || appData.todoColumns === LEGACY_TODO_COLUMNS_V2) {
        appData.todoColumns = DEFAULT_TODO_COLUMNS;
    }
    // 旧既定のヘッダー名「時間」は「担当」へ移行する（変更済みの名前はそのまま）
    if (appData.headers[2] === "時間") appData.headers[2] = "担当";

    // ToDoの出力項目に残っている旧列名を、現在のヘッダー名に置き換えておく
    appData.todoColumns = normalizeTodoColumns(
        appData.todoColumns.split(",").map(v => v.trim()).filter(Boolean),
        appData.headers
    ).join(", ");

    projectNameInput.value = data.projectName || "標準の計画";
    document.title = projectNameInput.value + " | 工程表";
    if (typeof data.memo === "string") {
        const memoText = appData.memoFormat === "plain"
            ? data.memo
            : extractPlainTextFromHTML(data.memo);
        setFreeMemoText(memoText);
    } else {
        setFreeMemoText("");
    }

    renderLeftHeader();

    document.getElementById("todoColumnsInput").value = appData.todoColumns;
    if (guideModeSelect) guideModeSelect.value = appData.settings.guideMode;
    updateHolidayToggleButton();
    setMemoCollapsed(appData.settings.memoCollapsed === true, { skipSave: true });
    applyMemoWidth(appData.settings.memoWidth);

    leftRowsContainer.innerHTML = "";
    rowsContainer.innerHTML = "";
    taskObjects = [];

    buildTimeline();
    buildHeader();
    leftColumnWidths = appData.columnWidths.slice();
    applyLeftColumnWidths();

    if (appData.tasks && appData.tasks.length > 0) {
        appData.tasks.forEach(tData => {
            tData.mainSchedules = normalizeMainSchedules(tData);
            addTaskRow(tData);
        });
    } else {
        addTaskRow();
    }
    isRestoringData = false;
    cleanupDependencies();
    clearDependencyDraft(false);
}

function scrollToToday() {
    let todayIdx = timelineDays.findIndex(d => d.iso === todayISO);
    if (todayIdx === -1 && appData.settings.hideHolidays === true) {
        todayIdx = dateToVisibleIndexAtOrAfter(todayISO);
    }
    if (todayIdx !== -1) {
        const x = todayIdx * CELL_WIDTH;
        const scrollContainer = document.querySelector(".gantt-scroll-container");
        if (scrollContainer) {
            scrollContainer.scrollLeft = x - (scrollContainer.clientWidth / 2) + 280; 
        }
    }
}

function buildTimeline() {
    allTimelineDays = [];
    timelineDays = [];
    const startDt = isoToDate(appData.settings.startDate);
    const endDt = isoToDate(appData.settings.endDate);
    const curr = new Date(startDt);

    while (curr <= endDt) {
        const iso = dateToISO(curr);
        const dow = curr.getDay();
        allTimelineDays.push({
            index: allTimelineDays.length,
            date: new Date(curr),
            iso,
            day: curr.getDate(),
            dow,
            month: curr.getMonth() + 1,
            year: curr.getFullYear(),
            isWeekend: dow === 0 || dow === 6,
            isHoliday: appData.settings.holidays.includes(iso),
            isVacation: (appData.settings.vacations || []).includes(iso),
            isToday: iso === todayISO
        });
        curr.setDate(curr.getDate() + 1);
    }
    timelineDays = appData.settings.hideHolidays === true
        ? allTimelineDays.filter((d) => !d.isWeekend && !d.isHoliday)
        : allTimelineDays.slice();
    if (timelineDays.length === 0) {
        timelineDays = allTimelineDays.slice();
    }
    timelineDays.forEach((d, index) => {
        d.index = index;
    });
    updateRangeLabel();
}

function updateRangeLabel() {
    if (!timelineDays.length) { rangeLabel.textContent = ""; return; }
    rangeLabel.textContent = `${dateToISO(timelineDays[0].date)} 〜 ${dateToISO(timelineDays[timelineDays.length - 1].date)}`;
}

function buildHeader() {
    const total = timelineDays.length;
    const showMainLine = appData.settings.showMainLine !== false;
    headerRow.innerHTML = "";
    headerRow.style.gridTemplateColumns = `repeat(${total}, ${CELL_WIDTH}px)`;
    timelineDays.forEach((d) => {
        const c = document.createElement("div");
        c.className = "header-day";
        if (d.isWeekend) c.classList.add("weekend");
        if (d.isHoliday) c.classList.add("holiday");
        if (d.isVacation) c.classList.add("vacation");
        if (d.isToday) c.classList.add("today");
        if (d.isToday && showMainLine) c.classList.add("today-boundary");
        c.innerHTML = `<div class="header-day-num">${d.month}/${d.day}</div><div class="header-day-week">${WEEKDAYS[d.dow]}</div>`;
        headerRow.appendChild(c);
    });

    totalRow.innerHTML = "";
    totalRow.style.gridTemplateColumns = `repeat(${total}, ${CELL_WIDTH}px)`;
    timelineDays.forEach((d) => {
        const c = document.createElement("div");
        c.className = "total-cell";
        if (d.isWeekend) c.classList.add("weekend");
        if (d.isHoliday) c.classList.add("holiday");
        if (d.isVacation) c.classList.add("vacation");
        if (d.isToday) c.classList.add("today");
        c.dataset.iso = d.iso;
        totalRow.appendChild(c);
    });
}

function updateHolidayToggleButton() {
    if (!toggleHolidaysBtn) return;
    const isHidden = appData.settings.hideHolidays === true;
    toggleHolidaysBtn.textContent = isHidden ? "休日表示" : "休日非表示";
    toggleHolidaysBtn.title = isHidden ? "休日列を表示する" : "休日列を閉じて稼働日のみ表示する";
    toggleHolidaysBtn.setAttribute("aria-pressed", String(isHidden));
    toggleHolidaysBtn.classList.toggle("is-active", isHidden);
}

function rebuildTaskRowCells(task) {
    if (!task?.cellRowEl || !task?.rowEl) return;
    const total = timelineDays.length;
    task.rowEl.style.gridTemplateColumns = `repeat(${total}, ${CELL_WIDTH}px)`;
    task.cellRowEl.innerHTML = "";
    timelineDays.forEach((d, i) => {
        const c = document.createElement("div");
        c.className = "cell";
        if (d.isWeekend) c.classList.add("weekend");
        if (d.isHoliday) c.classList.add("holiday");
        if (d.isVacation) c.classList.add("vacation");
        if (d.isToday) c.classList.add("today");
        c.dataset.index = i;
        task.cellRowEl.appendChild(c);
    });
}

function refreshTimelineDisplay() {
    buildTimeline();
    buildHeader();
    taskObjects.forEach((task) => {
        task.pendingStartIndex = null;
        task.pendingStartDate = null;
        task.pendingStartLane = 0;
        task.pendingMainStartIndex = null;
        task.pendingMainStartDate = null;
        rebuildTaskRowCells(task);
    });
    updateHolidayToggleButton();
    renderAllSegments();
    requestAnimationFrame(() => {
        scrollToToday();
        scheduleProgressGuideRefresh();
    });
}

function ensureProgressGuideOverlay() {
    let overlay = rowsContainer.querySelector(".progress-guide-overlay");
    if (overlay) return overlay;

    overlay = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    overlay.classList.add("progress-guide-overlay");
    overlay.setAttribute("aria-hidden", "true");
    rowsContainer.appendChild(overlay);
    return overlay;
}

function ensureTotalGuideOverlay() {
    let overlay = totalRow.querySelector(".total-guide-overlay");
    if (overlay) return overlay;

    overlay = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    overlay.classList.add("total-guide-overlay");
    overlay.setAttribute("aria-hidden", "true");
    totalRow.appendChild(overlay);
    return overlay;
}

function getVisibleTaskRows() {
    return taskObjects.filter(task => task.rowEl && task.rowEl.offsetParent !== null);
}

// カミナリ線の表示方法: "off"（表示しない） / "main"（メインのみ） / "mainSub"（メイン＋サブ）
function normalizeGuideMode(mode, showMainLine) {
    if (mode === "off" || mode === "main" || mode === "mainSub") return mode;
    return showMainLine === false ? "off" : "main";
}

function guidePointFromElement(el, overlayRect, side) {
    if (!el) return null;
    const rect = el.getBoundingClientRect();
    return {
        x: (side === "start" ? rect.left : rect.right) - overlayRect.left,
        y: rect.top + (rect.height / 2) - overlayRect.top
    };
}

// 実績バーの右端（＝どこまで進んだか）の座標
function getMainProgressEdgePoint(task, main, overlayRect) {
    const el = task.segLayerEl.querySelector(`[data-guide-role="main-progress"][data-main-id="${main.id}"]`);
    return guidePointFromElement(el, overlayRect, "end");
}

// メイン計画のバーの端の座標（side: "start" | "end"）
function getMainBaseEdgePoint(task, main, overlayRect, side) {
    const el = task.segLayerEl.querySelector(`[data-guide-role="main-base"][data-main-id="${main.id}"]`);
    return guidePointFromElement(el, overlayRect, side);
}

// 計画の並び（同じ段に並ぶメイン計画、またはサブ計画の1段分）から、カミナリ線の折れ位置を求める。
// getProgress(bar) は実績バー右端の座標、getBase(bar, side) は計画バー端の座標を返す。
function computeGuideAnchor(bars, getProgress, getBase) {
    const list = [...bars].sort((a, b) => a.startDate.localeCompare(b.startDate));
    if (list.length === 0) return null;

    // まだ終わっていない最初の計画
    const target = list.find(bar => {
        const end = bar.endDate || bar.startDate;
        if (!bar.progressEndDate) return true;
        return bar.progressEndDate < end;
    });

    // 実績が今日より先まで進んでいる場合だけ、その先端に合わせる。
    // 今日より手前で終わっているものは「遅れ」ではないので、縦のままにする。
    const aheadPoint = (bar) => {
        if (!bar || !bar.progressEndDate || bar.progressEndDate <= todayISO) return null;
        return getProgress(bar) || getBase(bar, "end");
    };

    // すべて終わっている場合は、いちばん先まで進んでいる実績の先端に合わせる
    if (!target) {
        const furthest = list.reduce((acc, bar) =>
            (!acc || (bar.progressEndDate || "") > (acc.progressEndDate || "")) ? bar : acc, null);
        return aheadPoint(furthest);
    }

    // 進捗があるなら、着手時期の前後にかかわらずその先端に合わせる
    if (target.progressEndDate) {
        const p = getProgress(target);
        if (p) return p;
    }

    // 進捗がまだ無く、着手時期も来ていない場合。
    // 手前に今日より先まで終えた計画があれば、そこに合わせる
    if (target.startDate > todayISO) {
        const before = list.slice(0, list.indexOf(target));
        for (let i = before.length - 1; i >= 0; i--) {
            const p = aheadPoint(before[i]);
            if (p) return p;
        }
        return null;
    }

    return getBase(target, "start");
}

function getGuideAnchorPoint(task, overlayRect) {
    // 完了にした行は、実際の進捗に関わらず完了扱いとして縦に引く
    if (task.isDone) return null;
    return computeGuideAnchor(
        task.mainSchedules || [],
        (main) => getMainProgressEdgePoint(task, main, overlayRect),
        (main, side) => getMainBaseEdgePoint(task, main, overlayRect, side)
    );
}

// サブ計画の段ごとの折れ位置（上の段から順）。段に何もなければ x は null
function getSubGuidePoints(task, overlayRect) {
    if (task.isSubCollapsed) return [];
    const rowTop = task.rowEl.getBoundingClientRect().top - overlayRect.top;
    const lanes = new Map();
    (task.segments || []).forEach(seg => {
        const lane = seg._lane || 0;
        if (!lanes.has(lane)) lanes.set(lane, []);
        lanes.get(lane).push(seg);
    });
    const findEl = (seg, role) =>
        task.segLayerEl.querySelector(`[data-guide-role="${role}"][data-seg-id="${seg.id}"]`);

    return [...lanes.keys()].sort((a, b) => a - b).map(lane => {
        const laneY = rowTop + SUB_SCHEDULE_TOP + (lane * SEGMENT_OFFSET_Y);
        const anchor = task.isDone ? null : computeGuideAnchor(
            lanes.get(lane),
            (seg) => guidePointFromElement(findEl(seg, "sub-progress"), overlayRect, "end"),
            (seg, side) => guidePointFromElement(findEl(seg, "sub-base"), overlayRect, side)
        );
        return { y: anchor ? anchor.y : laneY, x: anchor ? anchor.x : null };
    });
}

function appendGuideForTask(commands, task, overlayRect, boundaryX, previousPoint, isFirstTask = false) {
    const rowRect = task.rowEl.getBoundingClientRect();
    const rowTop = rowRect.top - overlayRect.top;
    const rowBottom = rowRect.bottom - overlayRect.top;
    const anchorPoint = getGuideAnchorPoint(task, overlayRect);

    // メイン＋サブ: メインの段、サブの各段の順に、上から折れ位置をつないでいく
    if (normalizeGuideMode(appData.settings.guideMode, appData.settings.showMainLine) === "mainSub") {
        const subPoints = getSubGuidePoints(task, overlayRect);
        if (anchorPoint || subPoints.some(p => p.x !== null)) {
            const points = [
                anchorPoint || { x: boundaryX, y: rowTop + MAIN_LINE_Y },
                ...subPoints.map(p => ({ x: p.x !== null ? p.x : boundaryX, y: p.y }))
            ];
            points.forEach(p => commands.push(`L ${p.x} ${p.y}`));
            return points[points.length - 1];
        }
    }

    if (anchorPoint) {
        commands.push(`L ${anchorPoint.x} ${anchorPoint.y}`);
        return { x: anchorPoint.x, y: anchorPoint.y };
    }

    if (!isFirstTask && (previousPoint.x !== boundaryX || previousPoint.y !== rowTop)) {
        commands.push(`L ${boundaryX} ${rowTop}`);
    }

    if (previousPoint.x !== boundaryX || previousPoint.y !== rowBottom) {
        commands.push(`L ${boundaryX} ${rowBottom}`);
    }

    return { x: boundaryX, y: rowBottom };
}

function renderProgressGuide() {
    const overlay = ensureProgressGuideOverlay();
    const totalOverlay = ensureTotalGuideOverlay();
    overlay.innerHTML = "";
    totalOverlay.innerHTML = "";

    if (normalizeGuideMode(appData.settings.guideMode, appData.settings.showMainLine) === "off") return;
    if (!timelineDays.length) return;

    const visibleTasks = getVisibleTaskRows();
    if (visibleTasks.length === 0) return;

    const todayHeader = headerRow.querySelector(".header-day.today-boundary");
    if (!todayHeader) return;

    const overlayWidth = Math.max(headerRow.scrollWidth, rowsContainer.scrollWidth, totalRow.scrollWidth);
    const overlayHeight = rowsContainer.scrollHeight;
    overlay.setAttribute("width", String(overlayWidth));
    overlay.setAttribute("height", String(overlayHeight));
    overlay.setAttribute("viewBox", `0 0 ${overlayWidth} ${overlayHeight}`);
    totalOverlay.setAttribute("width", String(overlayWidth));
    totalOverlay.setAttribute("height", String(totalRow.offsetHeight));
    totalOverlay.setAttribute("viewBox", `0 0 ${overlayWidth} ${totalRow.offsetHeight}`);

    const overlayRect = overlay.getBoundingClientRect();
    const todayRect = todayHeader.getBoundingClientRect();
    const startX = todayRect.left - overlayRect.left;
    const startY = 0;

    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.classList.add("progress-guide-path");
    const commands = [`M ${startX} ${startY}`];
    let lastGuidePoint = { x: startX, y: startY };

    visibleTasks.forEach((task, index) => {
        lastGuidePoint = appendGuideForTask(
            commands,
            task,
            overlayRect,
            startX,
            lastGuidePoint,
            index === 0
        );
    });

    const lastTask = visibleTasks[visibleTasks.length - 1];
    let finalBoundaryY = null;
    let finalTotalBottomY = null;
    if (lastTask) {
        const lastRowRect = lastTask.rowEl.getBoundingClientRect();
        finalBoundaryY = lastRowRect.bottom - overlayRect.top;
        const totalRect = totalRow.getBoundingClientRect();
        finalTotalBottomY = totalRect.bottom - overlayRect.top;
    }

    path.setAttribute("d", commands.join(" "));
    overlay.appendChild(path);

    if (finalBoundaryY != null) {
        const connectorPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
        connectorPath.classList.add("progress-guide-path");
        connectorPath.setAttribute("d", `M ${lastGuidePoint.x} ${lastGuidePoint.y} L ${startX} ${finalBoundaryY}`);
        overlay.appendChild(connectorPath);

        if (finalTotalBottomY != null && finalTotalBottomY !== finalBoundaryY) {
            const tailPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
            tailPath.classList.add("progress-guide-path");
            tailPath.setAttribute("d", `M ${startX} 0 L ${startX} ${totalRow.offsetHeight}`);
            totalOverlay.appendChild(tailPath);
        }
    }
}

function ensureDependencyOverlay() {
    let overlay = ganttRight.querySelector(".dependency-overlay");
    if (!overlay) {
        overlay = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        overlay.classList.add("dependency-overlay");
        ganttRight.appendChild(overlay);
    }
    return overlay;
}

function createDependencyPathD(startPoint, endPoint) {
    const dx = endPoint.x - startPoint.x;
    const cp = Math.max(42, Math.min(120, Math.abs(dx) * 0.45));
    const cpDir = dx >= 0 ? 1 : -1;
    const c1x = startPoint.x + (cp * cpDir);
    const c2x = endPoint.x - (cp * cpDir);
    return `M ${startPoint.x} ${startPoint.y} C ${c1x} ${startPoint.y}, ${c2x} ${endPoint.y}, ${endPoint.x} ${endPoint.y}`;
}

function renderDependencies() {
    const overlay = ensureDependencyOverlay();
    overlay.innerHTML = "";
    cleanupDependencies();

    const overlayWidth = Math.max(headerRow.scrollWidth, rowsContainer.scrollWidth, totalRow.scrollWidth);
    const overlayHeight = headerRow.offsetHeight + rowsContainer.scrollHeight + totalRow.offsetHeight;
    overlay.setAttribute("width", String(overlayWidth));
    overlay.setAttribute("height", String(overlayHeight));
    overlay.setAttribute("viewBox", `0 0 ${overlayWidth} ${overlayHeight}`);

    const defs = document.createElementNS("http://www.w3.org/2000/svg", "defs");

    const marker = document.createElementNS("http://www.w3.org/2000/svg", "marker");
    marker.setAttribute("id", "dependencyArrow");
    marker.setAttribute("viewBox", "0 0 10 10");
    marker.setAttribute("refX", "9");
    marker.setAttribute("refY", "5");
    marker.setAttribute("markerWidth", "7");
    marker.setAttribute("markerHeight", "7");
    marker.setAttribute("orient", "auto-start-reverse");
    const markerPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
    markerPath.setAttribute("d", "M 0 0 L 10 5 L 0 10 z");
    markerPath.setAttribute("fill", "#475569");
    marker.appendChild(markerPath);
    defs.appendChild(marker);

    const draftMarker = document.createElementNS("http://www.w3.org/2000/svg", "marker");
    draftMarker.setAttribute("id", "dependencyArrowDraft");
    draftMarker.setAttribute("viewBox", "0 0 10 10");
    draftMarker.setAttribute("refX", "9");
    draftMarker.setAttribute("refY", "5");
    draftMarker.setAttribute("markerWidth", "7");
    draftMarker.setAttribute("markerHeight", "7");
    draftMarker.setAttribute("orient", "auto-start-reverse");
    const draftMarkerPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
    draftMarkerPath.setAttribute("d", "M 0 0 L 10 5 L 0 10 z");
    draftMarkerPath.setAttribute("fill", "#94a3b8");
    draftMarker.appendChild(draftMarkerPath);
    defs.appendChild(draftMarker);
    overlay.appendChild(defs);

    const overlayRect = overlay.getBoundingClientRect();

    ensureDependenciesArray().forEach(dep => {
        const startEl = getMainEndpointElement({
            taskId: dep.fromTaskId,
            mainId: dep.fromMainId,
            anchor: dep.fromAnchor
        });
        const endEl = getMainEndpointElement({
            taskId: dep.toTaskId,
            mainId: dep.toMainId,
            anchor: dep.toAnchor
        });
        if (!startEl || !endEl) return;

        const startRect = startEl.getBoundingClientRect();
        const endRect = endEl.getBoundingClientRect();
        const startPoint = {
            x: startRect.left + (startRect.width / 2) - overlayRect.left,
            y: startRect.top + (startRect.height / 2) - overlayRect.top
        };
        const endPoint = {
            x: endRect.left + (endRect.width / 2) - overlayRect.left,
            y: endRect.top + (endRect.height / 2) - overlayRect.top
        };
        const d = createDependencyPathD(startPoint, endPoint);

        const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
        path.classList.add("dependency-path");
        if (contextMenuTargetDependencyId === dep.id) {
            path.classList.add("selected");
        }
        path.setAttribute("d", d);
        path.setAttribute("marker-end", "url(#dependencyArrow)");
        overlay.appendChild(path);

        const hitPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
        hitPath.classList.add("dependency-hit-path");
        hitPath.setAttribute("d", d);
        hitPath.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            e.stopPropagation();
            hideContextMenus();
            contextMenuTargetDependencyId = dep.id;
            renderDependencies();
            dependencyContextMenu.style.display = "block";
            dependencyContextMenu.style.left = e.pageX + "px";
            dependencyContextMenu.style.top = e.pageY + "px";
        });
        overlay.appendChild(hitPath);
    });

    if (dependencyDraft) {
        const sourceEl = getMainEndpointElement(dependencyDraft);
        if (sourceEl && dependencyMousePosition) {
            const sourceRect = sourceEl.getBoundingClientRect();
            const startPoint = {
                x: sourceRect.left + (sourceRect.width / 2) - overlayRect.left,
                y: sourceRect.top + (sourceRect.height / 2) - overlayRect.top
            };
            const endPoint = {
                x: dependencyMousePosition.x - overlayRect.left,
                y: dependencyMousePosition.y - overlayRect.top
            };
            const draftPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
            draftPath.classList.add("dependency-path", "draft");
            draftPath.setAttribute("d", createDependencyPathD(startPoint, endPoint));
            draftPath.setAttribute("marker-end", "url(#dependencyArrowDraft)");
            overlay.appendChild(draftPath);
        }
    }

    getMainEndpointElements().forEach(el => {
        const ref = {
            taskId: el.dataset.taskId,
            mainId: el.dataset.mainId,
            anchor: el.dataset.anchor
        };
        el.classList.toggle("dependency-source", !!dependencyDraft && isSameEndpoint(ref, dependencyDraft));
        el.classList.toggle("dependency-target", !!dependencyDraft && !isSameEndpoint(ref, dependencyDraft));
    });
}

function scheduleProgressGuideRefresh() {
    if (pendingGuideRefreshFrame !== null) {
        cancelAnimationFrame(pendingGuideRefreshFrame);
    }
    pendingGuideRefreshFrame = requestAnimationFrame(() => {
        pendingGuideRefreshFrame = null;
        renderProgressGuide();
        renderDependencies();
    });
}

// ============================================
// 行操作 (Drag & Drop)
// ============================================
let dragSrcEl = null;
function handleRowDragStart(e) {
    dragSrcEl = this.closest('.left-row');
    dragSrcEl.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    const rows = Array.from(leftRowsContainer.children);
    e.dataTransfer.setData('text/plain', rows.indexOf(dragSrcEl));
}
function handleRowDrop(e) {
    e.stopPropagation();
    const targetRow = this.closest('.left-row');
    if (dragSrcEl !== targetRow) {
        const rows = Array.from(leftRowsContainer.children);
        const srcIdx = parseInt(e.dataTransfer.getData('text/plain'));
        const targetIdx = rows.indexOf(targetRow);
        const movedItem = taskObjects.splice(srcIdx, 1)[0];
        taskObjects.splice(targetIdx, 0, movedItem);
        refreshRowsDOM();
        renderAllSegments();
        triggerSave();
    }
    return false;
}
function handleRowDragEnd(e) {
    leftRowsContainer.querySelectorAll('.left-row').forEach(r => { r.classList.remove('over'); r.classList.remove('dragging'); });
}
function refreshRowsDOM() {
    taskObjects.forEach(task => { leftRowsContainer.appendChild(task.leftRowEl); rowsContainer.appendChild(task.rowEl); });
}

function addTaskRow(initialData = null) {
    const id = initialData ? initialData.id : "task_" + Date.now() + "_" + Math.random().toString(36).slice(2);
    const total = timelineDays.length;

    const leftRow = document.createElement("div");
    leftRow.className = "left-row";
    if (initialData && initialData.isDone) leftRow.classList.add("task-done");
    if (initialData && initialData.isHidden) leftRow.classList.add("task-hidden");

    const grip = document.createElement("div");
    grip.className = "drag-handle"; 
    grip.dataset.taskId = id;

    const gripIcon = document.createElement("span");
    gripIcon.className = "grip-icon";
    gripIcon.textContent = "⠿";
    grip.appendChild(gripIcon);

    const insertBtn = document.createElement("div");
    insertBtn.className = "row-insert-btn";
    insertBtn.textContent = "+";
    insertBtn.title = "この下に行を追加";
    const stopEvt = (e) => { e.stopPropagation(); };
    insertBtn.addEventListener("mousedown", stopEvt);
    insertBtn.addEventListener("dragstart", stopEvt);
    insertBtn.addEventListener("click", (e) => { e.stopPropagation(); insertTaskAfter(id); });
    
    grip.appendChild(insertBtn);

    const subToggleBtn = document.createElement("button");
    subToggleBtn.className = "sub-toggle-btn";
    subToggleBtn.type = "button";
    subToggleBtn.addEventListener("mousedown", stopEvt);
    subToggleBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        const t = taskObjects.find(task => task.id === id);
        if (!t) return;
        setTaskSubScheduleCollapsed(t, !t.isSubCollapsed);
    });
    grip.appendChild(subToggleBtn);

    grip.addEventListener('dragstart', (e) => {
        if (selectionMode !== 0) {
            e.preventDefault(); 
            return;
        }
        handleRowDragStart.call(grip, e);
    });
    
    grip.addEventListener("click", (e) => {
        if (selectionMode !== 0) {
            e.stopPropagation();
            const t = taskObjects.find(task => task.id === id);
            if (t) {
                t.isSelected = !t.isSelected;
                renderGrip(t);
            }
        }
    });

    leftRow.appendChild(grip);

    leftRow.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; return false; });
    leftRow.addEventListener('dragenter', function() { this.classList.add('over'); });
    leftRow.addEventListener('dragleave', function() { this.classList.remove('over'); });
    leftRow.addEventListener('drop', handleRowDrop);
    leftRow.addEventListener('dragend', handleRowDragEnd);

    const createInput = (ph, text, isFirst) => {
        const cell = document.createElement("div"); cell.className = "label-cell";
        const ed = document.createElement("div"); ed.className = "editable"; ed.contentEditable = "true"; ed.dataset.placeholder = ph;
        if (text) ed.textContent = text;
        setupPlainTextEditing(ed, triggerSave);
        if (isFirst) {
            const openMemo = (e) => {
                e.stopPropagation();
                openTaskMemo(task, task.leftRowEl, true);
            };
            cell.addEventListener("click", openMemo);
            ed.addEventListener("click", openMemo);
        }
        cell.appendChild(ed);
        return cell;
    };
    const headerNames = normalizeHeaders(appData.headers);
    const initialLabels = readTaskLabels(initialData, headerNames.length);
    headerNames.forEach((name, i) => {
        leftRow.appendChild(createInput(name, initialData ? initialLabels[i] : "", i === 0));
    });
    leftRow.style.gridTemplateColumns = leftColumnWidths.map(w => `${w}px`).join(" ");

    const rowResize = document.createElement("div");
    rowResize.className = "row-resize-handle";
    rowResize.style.left = `${leftColumnWidths[0]}px`;
    rowResize.addEventListener("mousedown", (e) => {
        e.preventDefault();
        e.stopPropagation();
        isResizingRow = true;
        resizeRowTaskId = id;
        resizeStartY = e.clientY;
        resizeStartHeight = leftRow.getBoundingClientRect().height;
        document.body.style.cursor = "row-resize";
        document.body.style.userSelect = "none";
    });
    leftRow.appendChild(rowResize);

    leftRow.addEventListener('contextmenu', (e) => { e.preventDefault(); showContextMenu(e, id); });
    leftRowsContainer.appendChild(leftRow);

    const row = document.createElement("div");
    row.className = "task-row";
    row.dataset.id = id;
    row.style.gridTemplateColumns = `repeat(${total}, ${CELL_WIDTH}px)`;
    if (initialData && initialData.isDone) row.classList.add("task-done");
    if (initialData && initialData.isHidden) row.classList.add("task-hidden");

    const cellRow = document.createElement("div"); cellRow.style.display = "contents";
    for (let i = 0; i < total; i++) {
        const d = timelineDays[i];
        const c = document.createElement("div"); c.className = "cell";
        if (d.isWeekend) c.classList.add("weekend");
        if (d.isHoliday) c.classList.add("holiday");
        if (d.isVacation) c.classList.add("vacation");
        if (d.isToday) c.classList.add("today");
        c.dataset.index = i; cellRow.appendChild(c);
    }
    row.appendChild(cellRow);
    const segLayer = document.createElement("div"); segLayer.className = "segments-layer"; row.appendChild(segLayer);
    rowsContainer.appendChild(row);

    const task = {
        id, rowEl: row, leftRowEl: leftRow, cellRowEl: cellRow, segLayerEl: segLayer,
        mainSchedules: initialData ? normalizeMainSchedules(initialData) : [],
        segments: initialData ? initialData.segments.map(seg => ({
            ...seg,
            isSelected: false,
            _lane: 0
        })) : [],
        memo: initialData ? (initialData.memo || "") : "",
        customHeight: initialData ? (initialData.customHeight || 0) : 0,
        isDone: initialData ? !!initialData.isDone : false,
        isHidden: initialData ? !!initialData.isHidden : false,
        isSubCollapsed: initialData ? !!initialData.isSubCollapsed : false,
        color: initialData ? normalizeRowColor(initialData.color) : "",
        pendingMainStartIndex: null, pendingMainStartDate: null,
        pendingStartIndex: null, pendingStartDate: null, pendingStartLane: 0,
        isSelected: false
    };
    taskObjects.push(task);
    
    applyTaskRowColor(task);
    applyTaskSubScheduleState(task);
    renderGrip(task); 

    setupRowInteraction(task);
    activeTaskId = id;
    renderAllSegments();
    if (!initialData) triggerSave();
}

function renderGrip(task) {
    const grip = task.leftRowEl.querySelector(".drag-handle");
    if (!grip) return;
    const iconSpan = grip.querySelector(".grip-icon");
    if (!iconSpan) return;

    let iconText = "⠿";
    if (selectionMode === 1 || selectionMode === 2) {
        iconText = task.isSelected ? "☑️" : "□";
    }

    iconSpan.textContent = iconText;
    grip.draggable = (selectionMode === 0);
    grip.style.cursor = (selectionMode === 0) ? "grab" : "pointer";
}

function insertTaskAfter(targetTaskId) {
    addTaskRow();
    const newTask = taskObjects[taskObjects.length - 1];
    const targetIndex = taskObjects.findIndex(t => t.id === targetTaskId);
    if (targetIndex === -1) return;

    const targetTask = taskObjects[targetIndex];
    if (targetTask.leftRowEl.nextSibling) {
        leftRowsContainer.insertBefore(newTask.leftRowEl, targetTask.leftRowEl.nextSibling);
        rowsContainer.insertBefore(newTask.rowEl, targetTask.rowEl.nextSibling);
    }
    taskObjects.pop(); 
    taskObjects.splice(targetIndex + 1, 0, newTask);
    triggerSave();
}

// ============================================
// 描画ロジック
// ============================================
function renderAllSegments() {
    if (!timelineDays.length) return;
    const rangeStart = timelineDays[0].date;
    const rangeEnd = timelineDays[timelineDays.length - 1].date;

    taskObjects.forEach((task) => {
        const isSubCollapsed = !!task.isSubCollapsed;
        task.segLayerEl.innerHTML = "";
        drawScheduleDivider(task);

        const taskDates = {};
        task.segments.forEach(seg => seg._lane = 0);
        const sortedSegs = [...task.segments].sort((a, b) => (a.startDate !== b.startDate) ? (a.startDate < b.startDate ? -1 : 1) : (a.endDate < b.endDate ? -1 : 1));
        
        let maxLaneUsed = 0;
        sortedSegs.forEach(seg => {
            let requiredLane = 0;
            const visibleRange = getVisibleRangeIndices(seg.startDate, seg.endDate);
            if (!visibleRange) return;
            let { sIdx, eIdx } = visibleRange;
            while (true) {
                let overlap = false;
                for (let i = Math.min(sIdx, eIdx); i <= Math.max(sIdx, eIdx); i++) {
                    const iso = timelineDays[i].iso;
                    if (taskDates[iso] && taskDates[iso].includes(requiredLane)) { overlap = true; break; }
                }
                if (!overlap) break;
                requiredLane++;
            }
            seg._lane = requiredLane;
            maxLaneUsed = Math.max(maxLaneUsed, requiredLane);
            for (let i = Math.min(sIdx, eIdx); i <= Math.max(sIdx, eIdx); i++) {
                const iso = timelineDays[i].iso;
                if (!taskDates[iso]) taskDates[iso] = [];
                taskDates[iso].push(requiredLane);
            }
        });

        const draftLaneCount = (!isSubCollapsed && task.pendingStartIndex != null) ? ((task.pendingStartLane || 0) + 1) : 0;
        const laneCount = Math.max(task.segments.length > 0 ? (maxLaneUsed + 1) : 1, draftLaneCount || 1);
        const newHeight = isSubCollapsed
            ? COLLAPSED_ROW_HEIGHT
            : Math.max(BASE_ROW_HEIGHT, SUB_SCHEDULE_TOP + ((laneCount - 1) * SEGMENT_OFFSET_Y) + 26);
        task.baseHeight = newHeight;
        const finalHeight = isSubCollapsed ? newHeight : Math.max(newHeight, task.customHeight || 0);
        task.rowEl.style.height = finalHeight + "px";
        task.leftRowEl.style.height = finalHeight + "px";
        applyTaskSubScheduleState(task);

        (task.mainSchedules || []).forEach((main, mainIndex) => {
            drawMainSchedule(task, main, mainIndex);
        });

        if (!isSubCollapsed) {
            task.segments.forEach((seg) => {
                const lane = seg._lane || 0;
                const topPx = SUB_SCHEDULE_TOP + (lane * SEGMENT_OFFSET_Y);
                if (seg.type === "point") {
                    const idx = dateToIndex(seg.startDate);
                    if (idx !== -1) drawPointSegment(task, seg, idx, topPx);
                } else {
                    const sdt = isoToDate(seg.startDate), edt = isoToDate(seg.endDate);
                    if (edt >= rangeStart && sdt <= rangeEnd) {
                        const visibleRange = getVisibleRangeIndices(seg.startDate, seg.endDate);
                        if (visibleRange) drawRangeSegment(task, seg, visibleRange.sIdx, visibleRange.eIdx, topPx);
                    }
                }
            });
        }
        if (task.pendingMainStartIndex != null) drawDraftStart(task, task.pendingMainStartIndex, MAIN_LINE_Y);
        if (!isSubCollapsed && task.pendingStartIndex != null) drawDraftStart(task, null, SUB_SCHEDULE_TOP + ((task.pendingStartLane || 0) * SEGMENT_OFFSET_Y));
        adjustLabelPositions(task);
    });
    calculateTotals();
    updateBottomRowBorders();
    renderProgressGuide();
    renderDependencies();
}

function drawScheduleDivider(task) {
    const divider = document.createElement("div");
    divider.className = "schedule-divider";
    if (task.isSubCollapsed) divider.classList.add("is-collapsed");
    divider.style.top = MAIN_DIVIDER_Y + "px";
    task.segLayerEl.appendChild(divider);
}

function attachMainEndpointInteractions(pt, task, main, anchor, editHandler) {
    pt.dataset.taskId = task.id;
    pt.dataset.mainId = main.id;
    pt.dataset.anchor = anchor;
    pt.addEventListener("mousedown", (e) => {
        if (dependencyDraft) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        initDrag(e, task, main, "move", pt, "main");
    });
    pt.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (!dependencyDraft) return;
        completeDependencyDraft(task, main, anchor);
    });
    pt.addEventListener("dblclick", (e) => {
        e.preventDefault();
        e.stopPropagation();
        editHandler();
    });
    pt.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        e.stopPropagation();
        showSegmentContextMenu(e, task, { id: main.id, scheduleScope: "main", anchor });
    });
}

function drawMainSchedule(task, main, mainIndex = 0) {
    if (!main) return;
    const isProgressSelected = activeProgressSegmentId === main.id;

    const rangeStart = timelineDays[0]?.date;
    const rangeEnd = timelineDays[timelineDays.length - 1]?.date;
    const startDate = isoToDate(main.startDate);
    const endDate = isoToDate(main.endDate);
    if (!rangeStart || !rangeEnd || endDate < rangeStart || startDate > rangeEnd) return;

    const visibleStart = startDate < rangeStart ? rangeStart : startDate;
    const visibleEnd = endDate > rangeEnd ? rangeEnd : endDate;
    const visibleRange = getVisibleRangeIndices(dateToISO(visibleStart), dateToISO(visibleEnd));
    if (!visibleRange) return;
    const { sIdx, eIdx } = visibleRange;

    const sc = centerX(sIdx);
    const ec = centerX(eIdx);
    const left = Math.min(sc, ec);
    const width = Math.max(1, Math.abs(sc - ec));

    const line = document.createElement("div");
    line.className = "segment main-segment" + (isProgressSelected ? " progress-active" : "");
    line.dataset.guideRole = "main-base";
    line.dataset.mainId = main.id;
    line.style.left = left + "px";
    line.style.width = width + "px";
    line.style.top = MAIN_LINE_Y + "px";

    // 実績があると開始日は動かせないので、サブ計画と同じく左ハンドルを作らない
    if (main.progressEndDate) line.classList.add("fixed");
    if (!main.progressEndDate) {
        const lHandle = document.createElement("div");
        lHandle.className = "resize-handle left";
        lHandle.addEventListener("mousedown", (e) => initDrag(e, task, main, "resize-left", line, "main"));
        line.appendChild(lHandle);
    }

    // 実績が終了日まで達している（＝未実施の部分が無い）ときは右ハンドルも作らない
    if (!isMainFullyDone(main)) {
        const rHandle = document.createElement("div");
        rHandle.className = "resize-handle right";
        rHandle.addEventListener("mousedown", (e) => initDrag(e, task, main, "resize-right", line, "main"));
        line.appendChild(rHandle);
    }

    line.addEventListener("mousedown", (e) => {
        if (e.target.classList.contains("resize-handle")) return;
        initDrag(e, task, main, "move", line, "main");
    });
    line.addEventListener("dblclick", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (isCtrlSelectionMode || e.ctrlKey) return;
        editMainScheduleLabel(task, main);
    });
    line.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        e.stopPropagation();
        showSegmentContextMenu(e, task, { id: main.id, scheduleScope: "main" });
    });
    task.segLayerEl.appendChild(line);

    if (main.progressEndDate) {
        const progressDate = isoToDate(main.progressEndDate);
        const progressVisibleEnd = progressDate > visibleEnd ? visibleEnd : progressDate;
        if (progressVisibleEnd >= visibleStart) {
            const pIdx = dateToVisibleIndexAtOrBefore(dateToISO(progressVisibleEnd));
            if (pIdx !== -1) {
                const progressLeft = centerX(sIdx);
                const progressRight = (progressVisibleEnd < endDate && pIdx < eIdx)
                    ? (pIdx + 1) * CELL_WIDTH
                    : centerX(pIdx);
                const doneWidth = progressRight - progressLeft;
                if (doneWidth > 0) {
                    const doneLine = document.createElement("div");
                    doneLine.className = "segment main-segment done" + (isProgressSelected ? " progress-active" : "");
                    doneLine.dataset.guideRole = "main-progress";
                    doneLine.dataset.mainId = main.id;
                    doneLine.style.left = progressLeft + "px";
                    doneLine.style.width = doneWidth + "px";
                    doneLine.style.top = MAIN_LINE_Y + "px";
                    doneLine.style.pointerEvents = "none";
                    task.segLayerEl.appendChild(doneLine);
                }
            }
        }
    }

    if (startDate >= rangeStart && startDate <= rangeEnd) {
        const startIdx = dateToIndex(main.startDate);
        if (startIdx !== -1) {
            const pt = document.createElement("div");
            const startDone = main.progressEndDate && isoToDate(main.progressEndDate).getTime() >= startDate.getTime();
            pt.className = "point main-point" + (startDone ? " done" : "") + (isProgressSelected ? " progress-active" : "");
            pt.style.left = centerX(startIdx) + "px";
            pt.style.top = MAIN_LINE_Y + "px";
            pt.style.cursor = "grab";
            attachMainEndpointInteractions(pt, task, main, "start", () => editMainEndpointLabel(task, main, "start"));
            task.segLayerEl.appendChild(pt);
            drawMainEndpointLabel(task, main, "start", centerX(startIdx), main.startLabel || "", isProgressSelected);
        }
    }

    if (endDate >= rangeStart && endDate <= rangeEnd) {
        const endIdx = dateToIndex(main.endDate);
        if (endIdx !== -1) {
            const pt = document.createElement("div");
            const endDone = main.progressEndDate && isoToDate(main.progressEndDate).getTime() >= endDate.getTime();
            pt.className = "point main-point" + (endDone ? " done" : "") + (isProgressSelected ? " progress-active" : "");
            pt.style.left = centerX(endIdx) + "px";
            pt.style.top = MAIN_LINE_Y + "px";
            pt.style.cursor = "grab";
            attachMainEndpointInteractions(pt, task, main, "end", () => editMainEndpointLabel(task, main, "end"));
            task.segLayerEl.appendChild(pt);
            drawMainEndpointLabel(task, main, "end", centerX(endIdx), main.endLabel || "", isProgressSelected);
        }
    }

    if (main.label) {
        const lab = document.createElement("div");
        lab.className = "segment-label main-label" + (isProgressSelected ? " progress-active" : "");
        lab.textContent = main.label;
        lab.style.left = ((sc + ec) / 2) + "px";
        lab.dataset.baseTop = String(MAIN_LABEL_BASE_TOP);
        lab.style.top = MAIN_LABEL_BASE_TOP + "px";
        lab.addEventListener("click", (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (isCtrlSelectionMode || e.ctrlKey || dependencyDraft) return;
            editMainScheduleLabel(task, main);
        });
        lab.addEventListener("dblclick", (e) => {
            e.preventDefault();
            e.stopPropagation();
        });
        lab.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            e.stopPropagation();
            showSegmentContextMenu(e, task, { id: main.id, scheduleScope: "main" });
        });
        task.segLayerEl.appendChild(lab);
    }

    const milestones = [...(main.milestones || [])].sort((a, b) => a.date.localeCompare(b.date));
    milestones.forEach((milestone, index) => drawMainMilestone(task, main, milestone, index + (mainIndex * 100), line));
}

function drawMainMilestone(task, main, milestone, index, mainLineEl = null) {
    const idx = dateToIndex(milestone.date);
    if (idx === -1) return;
    const x = centerX(idx);
    const isAbove = index % 2 === 0;
    const labelTop = isAbove ? (MAIN_LINE_Y - 23) : (MAIN_LINE_Y + 9);

    const pt = document.createElement("div");
    const isDone = main?.progressEndDate && isoToDate(main.progressEndDate).getTime() >= isoToDate(milestone.date).getTime();
    pt.className = "point milestone-point" + (isDone ? " done" : "");
    pt.style.left = x + "px";
    pt.style.top = MAIN_LINE_Y + "px";
    pt.style.cursor = "grab";
    // 端点に重なっているマイルストーンは端点のコメントなので、掴んだら端点ごと動かす
    const atStart = milestone.date === main.startDate;
    const atEnd = milestone.date === main.endDate;
    pt.addEventListener("mousedown", (e) => {
        if (atStart && atEnd) {
            initDrag(e, task, main, "move", mainLineEl || pt, "main");
        } else if (atEnd) {
            initDrag(e, task, main, "resize-right", mainLineEl || pt, "main");
        } else if (atStart) {
            initDrag(e, task, main, "resize-left", mainLineEl || pt, "main");
        } else {
            initDrag(e, task, { ...milestone, mainId: main.id }, "move", pt, "milestone");
        }
    });
    pt.addEventListener("dblclick", (e) => {
        e.preventDefault();
        e.stopPropagation();
        editMilestone(task, main, milestone);
    });
    pt.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (confirm("このマイルストーンを削除しますか？")) {
            main.milestones = (main.milestones || []).filter(ms => ms.id !== milestone.id);
            renderAllSegments();
            triggerSave();
        }
    });
    task.segLayerEl.appendChild(pt);

    const label = document.createElement("div");
    label.className = "segment-label milestone-label";
    label.textContent = milestone.label || "マイルストーン";
    label.style.left = x + "px";
    label.dataset.baseTop = String(labelTop);
    label.style.top = labelTop + "px";
    label.addEventListener("dblclick", (e) => {
        e.preventDefault();
        e.stopPropagation();
        editMilestone(task, main, milestone);
    });
    label.addEventListener("click", (e) => {
        if (isCtrlSelectionMode || e.ctrlKey) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        e.preventDefault();
        e.stopPropagation();
        editMilestone(task, main, milestone);
    });
    label.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (confirm("このマイルストーンを削除しますか？")) {
            main.milestones = (main.milestones || []).filter(ms => ms.id !== milestone.id);
            renderAllSegments();
            triggerSave();
        }
    });
    task.segLayerEl.appendChild(label);
}

function editMilestone(task, main, milestone) {
    const nextLabel = prompt("マイルストーン名:", milestone.label || "マイルストーン");
    if (nextLabel === null) return;
    milestone.label = nextLabel.trim() || "マイルストーン";
    renderAllSegments();
    triggerSave();
}

function editSubSegmentLabel(task, seg) {
    const nextLabel = window.prompt("計画内容:", seg.label || "");
    if (nextLabel === null) return;
    seg.label = nextLabel.trim();
    renderAllSegments();
    triggerSave();
}

function handleSubLabelMouseDown(e, task, seg, dragEl) {
    if (!(isCtrlSelectionMode && seg.isSelected)) return;
    initDrag(e, task, seg, "move", dragEl);
}

function drawMainEndpointLabel(task, main, side, x, text, isProgressSelected = false) {
    if (!text) return;
    const label = document.createElement("div");
    label.className = "segment-label main-endpoint-label" + (isProgressSelected ? " progress-active" : "");
    label.textContent = text;
    label.style.left = x + "px";
    label.dataset.baseTop = String(MAIN_LABEL_BASE_TOP);
    label.style.top = MAIN_LABEL_BASE_TOP + "px";
    label.addEventListener("dblclick", (e) => {
        e.preventDefault();
        e.stopPropagation();
        editMainEndpointLabel(task, main, side);
    });
    label.addEventListener("click", (e) => {
        if (isCtrlSelectionMode || e.ctrlKey) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        e.preventDefault();
        e.stopPropagation();
        editMainEndpointLabel(task, main, side);
    });
    task.segLayerEl.appendChild(label);
}

function editMainScheduleLabel(task, main) {
    if (!main) return;
    const nextLabel = window.prompt("計画内容:", main.label || "");
    if (nextLabel === null) return;
    main.label = nextLabel.trim();
    renderAllSegments();
    triggerSave();
}

function editMainEndpointLabel(task, main, side) {
    if (!main) return;
    const key = side === "start" ? "startLabel" : "endLabel";
    const promptLabel = side === "start" ? "開始コメント" : "終了コメント";
    const nextLabel = prompt(`${promptLabel}:`, main[key] || "");
    if (nextLabel === null) return;
    main[key] = nextLabel.trim();
    renderAllSegments();
    triggerSave();
}

function calculateTotals() {
    const totals = {};
    timelineDays.forEach(d => totals[d.iso] = 0);
    taskObjects.forEach(task => {
        if (task.isHidden) return;
        task.segments.forEach(seg => {
            if (seg.dailyValues) {
                for (const [iso, val] of Object.entries(seg.dailyValues)) {
                    if (iso < seg.startDate || iso > seg.endDate) continue;
                    const num = parseFloat(val);
                    if (!isNaN(num)) totals[iso] += num;
                }
            }
        });
    });
    const cells = totalRow.children;
    for (let i = 0; i < cells.length; i++) {
        const cell = cells[i];
        const iso = cell.dataset.iso;
        let val = totals[iso];
        if (val > 0) {
            if (val > 99.9) val = 99.9;
            cell.textContent = (val % 1 === 0) ? val : val.toFixed(1);
        } else {
            cell.textContent = "";
        }
    }
}

function drawRangeSegment(task, seg, sIdx, eIdx, topPx) {
    const sc = centerX(sIdx), ec = centerX(eIdx);
    const baseLeft = Math.min(sc, ec), baseWidth = Math.max(1, Math.abs(sc - ec));
    const isProgressSelected = activeProgressSegmentId === seg.id;
    const isSelected = !!seg.isSelected;

    const hitbox = document.createElement("div");
    hitbox.className = "segment-hitbox" + (isSelected ? " segment-selected" : "");
    hitbox.dataset.subSelectable = "true";
    hitbox.style.left = baseLeft + "px";
    hitbox.style.width = baseWidth + "px";
    hitbox.style.top = topPx + "px";
    hitbox.addEventListener("mousedown", (e) => {
        if (e.ctrlKey) return;
        if (isCtrlSelectionMode && !seg.isSelected) return;
        initDrag(e, task, seg, "move", div);
    });
    addSegEvents(hitbox, task, seg);
    task.segLayerEl.appendChild(hitbox);

    const div = document.createElement("div");
    div.className = "segment" + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
    div.dataset.subSelectable = "true";
    div.dataset.guideRole = "sub-base";
    div.dataset.segId = seg.id;
    div.style.left = baseLeft + "px";
    div.style.width = baseWidth + "px";
    div.style.top = topPx + "px";

    if (seg.progressEndDate) {
        div.classList.add("fixed");
    }

    // [修正] 実績(progressEndDate)がある場合は開始日が固定されるため、左ハンドルは生成しない
    if (!seg.progressEndDate) {
        const lHandle = document.createElement("div"); lHandle.className = "resize-handle left";
        lHandle.addEventListener("mousedown", (e) => initDrag(e, task, seg, "resize-left", div));
        div.appendChild(lHandle);
    }

    // [修正] 完全に完了している(progress >= end)場合のみ右ハンドルを生成しない（未完了なら生成する）
    const isFullyDone = seg.progressEndDate && (isoToDate(seg.progressEndDate).getTime() >= isoToDate(seg.endDate).getTime());
    if (!isFullyDone) {
        const rHandle = document.createElement("div"); rHandle.className = "resize-handle right";
        rHandle.addEventListener("mousedown", (e) => initDrag(e, task, seg, "resize-right", div));
        div.appendChild(rHandle);
    }

    div.addEventListener("mousedown", (e) => {
        if (e.target.classList.contains("resize-handle")) return;
        initDrag(e, task, seg, "move", div);
    });

    addSegEvents(div, task, seg);
    task.segLayerEl.appendChild(div);

    const visibleRange = getVisibleRangeIndices(seg.startDate, seg.endDate);
    if (visibleRange) {
        const minI = visibleRange.sIdx;
        const maxI = visibleRange.eIdx;
        for (let i = minI; i <= maxI; i++) {
            const iso = timelineDays[i].iso;
            const x = centerX(i);
            const valDiv = document.createElement("div");
            valDiv.className = "daily-val";
            if (activeDailyValueTarget && activeDailyValueTarget.segId === seg.id && activeDailyValueTarget.iso === iso) {
                valDiv.classList.add("active");
            }
            valDiv.style.left = x + "px";
            valDiv.style.top = (topPx + 4) + "px";
            if (seg.dailyValues && seg.dailyValues[iso] != null) {
                const rawV = seg.dailyValues[iso];
                const v = parseFloat(rawV);
                valDiv.textContent = (!isNaN(v) && /^\d(\.\d)?$/.test(rawV)) ? ((v % 1 === 0) ? v : v.toFixed(1)) : rawV;
            }
            valDiv.addEventListener("click", (e) => handleDailyValueClick(e, task, seg, iso));
            task.segLayerEl.appendChild(valDiv);
        }
    }

    if (seg.progressEndDate) {
        const sIdxRaw = dateToVisibleIndexAtOrAfter(seg.startDate);
        const pIdxRaw = dateToVisibleIndexAtOrBefore(seg.progressEndDate);
        if (sIdxRaw !== -1 && pIdxRaw !== -1 && pIdxRaw >= sIdxRaw) {
            const left = centerX(sIdxRaw);
            const eIdxRaw = dateToVisibleIndexAtOrBefore(seg.endDate);
            let right = (eIdxRaw !== -1 && pIdxRaw < eIdxRaw) ? (pIdxRaw + 1) * CELL_WIDTH : centerX(pIdxRaw);
            const w = right - left;
            if (w > 0) {
                const dDiv = document.createElement("div");
                dDiv.className = "segment done";
                dDiv.dataset.guideRole = "sub-progress";
                dDiv.dataset.segId = seg.id;
                dDiv.style.left = left + "px";
                dDiv.style.width = w + "px";
                dDiv.style.pointerEvents = "none";
                dDiv.style.top = topPx + "px";
                task.segLayerEl.appendChild(dDiv);
            }
        }
    }

    const pointsData = [ { x: sc, d: isoToDate(seg.startDate), isEnd: false }, { x: ec, d: isoToDate(seg.endDate), isEnd: true } ];
    pointsData.forEach((ptData) => {
        const pt = document.createElement("div");
        let isDone = seg.progressEndDate && (isoToDate(seg.progressEndDate).getTime() >= ptData.d.getTime());
        pt.className = "point" + (isDone ? " done" : "") + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
        pt.dataset.subSelectable = "true";
        pt.style.left = ptData.x + "px"; pt.style.top = topPx + "px";
        
        pt.style.cursor = "grab";
        pt.addEventListener("mousedown", (e) => initDrag(e, task, seg, "move", div));

        addSegEvents(pt, task, seg);
        task.segLayerEl.appendChild(pt);
    });

    if (seg.label) {
        const lab = document.createElement("div");
        const isCompletedFull = seg.progressEndDate && isoToDate(seg.progressEndDate).getTime() >= isoToDate(seg.endDate).getTime();
        lab.className = "segment-label" + (isCompletedFull ? " done" : "") + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
        lab.dataset.subSelectable = "true";
        lab.textContent = seg.label;
        lab.style.left = (sc + ec) / 2 + "px";
        lab.addEventListener("mousedown", (e) => handleSubLabelMouseDown(e, task, seg, div));
        
        const baseTop = topPx - 19;
        lab.dataset.baseTop = baseTop; 
        lab.style.top = baseTop + "px";
        
        addSegEvents(lab, task, seg);
        task.segLayerEl.appendChild(lab);
    }
}

function drawPointSegment(task, seg, idx, topPx) {
    const c = centerX(idx);
    const isProgressSelected = activeProgressSegmentId === seg.id;
    const isDone = seg.progressEndDate && isoToDate(seg.progressEndDate).getTime() >= isoToDate(seg.startDate).getTime();
    const isSelected = !!seg.isSelected;

    const pt = document.createElement("div");
    pt.className = "point" + (isDone ? " done" : "") + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
    pt.dataset.subSelectable = "true";
    pt.dataset.guideRole = "sub-base";
    pt.dataset.segId = seg.id;
    pt.style.left = c + "px"; pt.style.top = topPx + "px";
    
    pt.style.cursor = "grab";
    pt.addEventListener("mousedown", (e) => initDrag(e, task, seg, "move", pt));

    addSegEvents(pt, task, seg);
    task.segLayerEl.appendChild(pt);

    const iso = timelineDays[idx].iso;
    const valDiv = document.createElement("div");
    valDiv.className = "daily-val";
    if (activeDailyValueTarget && activeDailyValueTarget.segId === seg.id && activeDailyValueTarget.iso === iso) {
        valDiv.classList.add("active");
    }
    valDiv.style.left = c + "px"; 
    valDiv.style.top = (topPx + 4) + "px";
    if (seg.dailyValues && seg.dailyValues[iso] != null) {
        valDiv.textContent = seg.dailyValues[iso];
    }
    valDiv.addEventListener("click", (e) => handleDailyValueClick(e, task, seg, iso));
    task.segLayerEl.appendChild(valDiv);

    if (seg.label) {
        const lab = document.createElement("div");
        lab.className = "segment-label" + (isDone ? " done" : "") + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
        lab.dataset.subSelectable = "true";
        lab.textContent = seg.label;
        lab.style.left = c + "px"; 
        lab.addEventListener("mousedown", (e) => handleSubLabelMouseDown(e, task, seg, pt));
        
        const baseTop = topPx - 19;
        lab.dataset.baseTop = baseTop;
        lab.style.top = baseTop + "px";
        
        addSegEvents(lab, task, seg);
        task.segLayerEl.appendChild(lab);
    }
}

function drawDraftStart(task, index = null, topPx = 30) {
    const targetIndex = index == null ? task.pendingStartIndex : index;
    if (targetIndex == null) return;
    const c = centerX(targetIndex);
    const pt = document.createElement("div"); pt.className = "point draft";
    pt.style.left = c + "px"; pt.style.top = topPx + "px";
    pt.title = "キャンセル";
    pt.addEventListener("click", (e) => {
        e.stopPropagation();
        if (index == null) {
            task.pendingStartIndex = null;
            task.pendingStartDate = null;
            task.pendingStartLane = 0;
        } else {
            task.pendingMainStartIndex = null;
            task.pendingMainStartDate = null;
        }
        renderAllSegments();
    });
    task.segLayerEl.appendChild(pt);
}

function adjustLabelPositions(task) {
    const labels = Array.from(task.segLayerEl.querySelectorAll(".segment-label:not(.milestone-label)"));
    if (labels.length === 0) return;

    const groups = {};
    labels.forEach(el => {
        const baseTop = parseFloat(el.dataset.baseTop);
        const key = Math.round(baseTop);
        if (!groups[key]) groups[key] = [];
        groups[key].push({ el, baseTop, left: parseFloat(el.style.left) });
    });

    Object.values(groups).forEach(items => {
        items.sort((a, b) => a.left - b.left);
        items.forEach((item, index) => {
            if (index % 2 === 0) {
                item.el.style.top = item.baseTop + "px";
                item.el.style.zIndex = "20";
            } else {
                item.el.style.top = (item.baseTop - 13.5) + "px";
                item.el.style.zIndex = "30";
            }
        });
    });
}

// ============================================
// Drag & Drop ロジック (移動・伸縮)
// ============================================
function initDrag(e, task, seg, type, el, scope = "sub") {
    if (e.button !== 0) return;
    if (isOverlayCancelStateActive()) return;
    if (e.ctrlKey) return;
    if (scope === "sub" && isCtrlSelectionMode && !seg.isSelected) return;

    let dragType = type;
    // [修正] 完了済み(progressEndDateあり)の場合、「blocked」という状態でドラッグを開始する。
    // 即座にreturnせず、グローバルなマウスイベントをあえて設定することで、
    // ドラッグ中のマウス操作をこの機能が「乗っ取る」形にし、裏側のセルに反応させないようにする。
    if (scope === "sub" && type === "move" && seg.progressEndDate) {
        dragType = "blocked";
    }

    // メイン計画も同様に、実績が1日でもあれば全体移動と開始日の伸縮を禁止する
    if (scope === "main" && seg.progressEndDate) {
        if (type === "move" || type === "resize-left") {
            dragType = "blocked";
        } else if (type === "resize-right" && isMainFullyDone(seg)) {
            dragType = "blocked";
        }
    }

    // デフォルト動作(テキスト選択など)と伝播を阻止
    e.preventDefault();
    e.stopPropagation();

    let selectedSegRefs = [];
    if (scope === "sub" && type === "move") {
        const currentSelected = getSelectedSubSegments();
        if (seg.isSelected && currentSelected.length > 0) {
            selectedSegRefs = currentSelected.map(ref => ({
                taskId: ref.task.id,
                segId: ref.seg.id,
                originalStartDate: ref.seg.startDate,
                originalEndDate: ref.seg.endDate
            }));
        }
    }

    dragState = {
        isDragging: true, type: dragType, taskId: task.id, segId: seg.id || null, milestoneId: scope === "milestone" ? (seg.id || null) : null, mainId: scope === "milestone" ? (seg.mainId || null) : (scope === "main" ? (seg.id || null) : null), selectedSegRefs, scope, startX: e.clientX,
        originalLeft: parseFloat(el.style.left), originalWidth: parseFloat(el.style.width),
        originalStartDate: seg.startDate || seg.date, originalEndDate: seg.endDate || seg.date, el: el
    };
    el.classList.add("dragging");
    
    // blockedの場合は「掴んでいる」ことを示すカーソルにするが、位置は更新されない
    if (dragType === "blocked") {
        document.body.style.cursor = "grabbing"; // 掴んでいるが動かせない
    } else {
        document.body.style.cursor = type === "move" ? "grabbing" : "col-resize";
    }
    
    document.addEventListener("mousemove", handleGlobalMouseMove);
    document.addEventListener("mouseup", handleGlobalMouseUp);
}

function handleGlobalMouseMove(e) {
    if (!dragState.isDragging) return;
    e.preventDefault(); // これにより裏側のテキスト選択やセルの反応を防ぐ

    // [修正] blocked状態（実績ありの移動）なら、座標計算もスタイル更新もしない
    if (dragState.type === "blocked") {
        return;
    }

    const diffPx = e.clientX - dragState.startX;
    if (dragState.type === "move") {
        dragState.el.style.left = (dragState.originalLeft + diffPx) + "px";
    } else if (dragState.type === "resize-right") {
        const newW = Math.max(0, dragState.originalWidth + diffPx); 
        dragState.el.style.width = newW + "px";
    } else if (dragState.type === "resize-left") {
        const newLeft = dragState.originalLeft + diffPx;
        const newWidth = dragState.originalWidth - diffPx;
        if (newWidth >= 0) { 
            dragState.el.style.left = newLeft + "px";
            dragState.el.style.width = newWidth + "px";
        }
    }
}

function handleGlobalMouseUp(e) {
    if (!dragState.isDragging) return;

    // [修正] blocked状態なら、何も計算せずにクリーンアップへ進む
    if (dragState.type !== "blocked") {
        const diffPx = e.clientX - dragState.startX;
        const dayDelta = Math.round(diffPx / CELL_WIDTH);
        const task = taskObjects.find(t => t.id === dragState.taskId);
        const seg = !task ? null : (
            dragState.scope === "main"
                ? findMainScheduleById(task, dragState.segId)
                : dragState.scope === "milestone"
                    ? (findMainScheduleById(task, dragState.mainId)?.milestones || []).find(ms => ms.id === dragState.milestoneId)
                    : task.segments.find(s => s.id === dragState.segId)
        );

        if (task && seg && dayDelta !== 0) {
            if (dragState.scope === "milestone" && dragState.type === "move") {
                const shiftedDate = shiftDateByVisibleColumns(dragState.originalStartDate, dayDelta);
                const main = findMainScheduleById(task, dragState.mainId);
                if (main) {
                    const clampedDate = shiftedDate < main.startDate
                        ? main.startDate
                        : shiftedDate > main.endDate
                            ? main.endDate
                            : shiftedDate;
                    seg.date = clampedDate;
                }
            } else if (dragState.scope === "sub" && dragState.type === "move" && dragState.selectedSegRefs.length > 0) {
                dragState.selectedSegRefs.forEach(ref => {
                    const refTask = taskObjects.find(t => t.id === ref.taskId);
                    const refSeg = refTask ? refTask.segments.find(s => s.id === ref.segId) : null;
                    if (!refSeg) return;
                    refSeg.startDate = shiftDateByVisibleColumns(ref.originalStartDate, dayDelta);
                    refSeg.endDate = shiftDateByVisibleColumns(ref.originalEndDate, dayDelta);
                    if (refSeg.dailyValues) {
                        const newVals = {};
                        Object.keys(refSeg.dailyValues).forEach(iso => newVals[shiftDateByVisibleColumns(iso, dayDelta)] = refSeg.dailyValues[iso]);
                        refSeg.dailyValues = newVals;
                    }
                    if (refSeg.dailyResults) {
                        const newRes = {};
                        Object.keys(refSeg.dailyResults).forEach(iso => newRes[shiftDateByVisibleColumns(iso, dayDelta)] = refSeg.dailyResults[iso]);
                        refSeg.dailyResults = newRes;
                    }
                });
            } else if (dragState.type === "move") {
                const nextStartDate = shiftDateByVisibleColumns(dragState.originalStartDate, dayDelta);
                const nextEndDate = shiftDateByVisibleColumns(dragState.originalEndDate, dayDelta);

                if (dragState.scope === "main" && hasOverlappingMainSchedule(task, nextStartDate, nextEndDate, seg.id)) {
                    alert("メイン計画は重複して設定できません。別の日付範囲を指定してください。");
                } else {
                    seg.startDate = nextStartDate;
                    seg.endDate = nextEndDate;

                    if (dragState.scope === "main" && Array.isArray(seg.milestones)) {
                        seg.milestones = seg.milestones.map(ms => ({
                            ...ms,
                            date: shiftDateByVisibleColumns(ms.date, dayDelta)
                        }));
                    }

                    if (dragState.scope !== "main" && seg.dailyValues) {
                        const newVals = {};
                        Object.keys(seg.dailyValues).forEach(iso => newVals[shiftDateByVisibleColumns(iso, dayDelta)] = seg.dailyValues[iso]);
                        seg.dailyValues = newVals;
                    }
                    if (dragState.scope !== "main" && seg.dailyResults) {
                        const newRes = {};
                        Object.keys(seg.dailyResults).forEach(iso => newRes[shiftDateByVisibleColumns(iso, dayDelta)] = seg.dailyResults[iso]);
                        seg.dailyResults = newRes;
                    }
                }

            } else if (dragState.type === "resize-right") {
                const newEnd = shiftDateByVisibleColumns(dragState.originalEndDate, dayDelta);
                let nextEnd = newEnd < seg.startDate ? seg.startDate : newEnd;
                if (dragState.scope === "main") {
                    // コメントが収まらなくなる手前で止める（縮める方向のみ）
                    const minEnd = minEndDateForMain(seg);
                    if (nextEnd < minEnd) nextEnd = minEnd > seg.endDate ? seg.endDate : minEnd;
                }
                if (dragState.scope === "main" && hasOverlappingMainSchedule(task, seg.startDate, nextEnd, seg.id)) {
                    alert("メイン計画は重複して設定できません。");
                } else {
                    const previousEnd = seg.endDate;
                    seg.endDate = nextEnd;
                    if (dragState.scope === "main") {
                        layoutMainMilestones(seg, seg.startDate, previousEnd);
                    }
                }
            } else if (dragState.type === "resize-left") {
                const newStart = shiftDateByVisibleColumns(dragState.originalStartDate, dayDelta);
                let nextStart = newStart > seg.endDate ? seg.endDate : newStart;
                if (dragState.scope === "main") {
                    const maxStart = maxStartDateForMain(seg);
                    if (nextStart > maxStart) nextStart = maxStart < seg.startDate ? seg.startDate : maxStart;
                }
                if (dragState.scope === "main" && hasOverlappingMainSchedule(task, nextStart, seg.endDate, seg.id)) {
                    alert("メイン計画は重複して設定できません。");
                } else {
                    const previousStart = seg.startDate;
                    seg.startDate = nextStart;
                    if (dragState.scope === "main") {
                        layoutMainMilestones(seg, previousStart, seg.endDate);
                    }
                }
            }
            triggerSave();
        }
    }

    if (dragState.el) dragState.el.classList.remove("dragging");
    document.body.style.cursor = "";
    document.removeEventListener("mousemove", handleGlobalMouseMove);
    document.removeEventListener("mouseup", handleGlobalMouseUp);
    dragState.isDragging = false; dragState.el = null;
    renderAllSegments();
}

// ============================================
// インタラクション (クリック等)
// ============================================
function setupRowInteraction(task) {
    task.rowEl.addEventListener("click", (e) => {
        if (dependencyDraft) return;
        if (segmentContextMenu.style.display === "block" || contextMenu.style.display === "block") return;
        if (e.target.closest(".segment") || e.target.closest(".point") || e.target.closest(".segment-label") || e.target.closest(".daily-val")) return;
        if (isCtrlSelectionMode || e.ctrlKey) return;
        clearSegmentSelection();
        const rect = task.rowEl.getBoundingClientRect();
        const idx = Math.max(0, Math.min(timelineDays.length - 1, Math.floor((e.clientX - rect.left) / CELL_WIDTH)));
        const y = e.clientY - rect.top;
        if (task.isSubCollapsed || y <= MAIN_DIVIDER_Y) handleMainCellClick(task, idx);
        else handleCellClick(task, idx, y);
    });
    task.leftRowEl.addEventListener("click", () => {
        if (dependencyDraft) return;
        if (segmentContextMenu.style.display === "block" || contextMenu.style.display === "block" || activeProgressSegmentId) return;
        if (isCtrlSelectionMode) return;
        activeTaskId = task.id;
        taskObjects.forEach(t => {
            t.pendingStartDate = null;
            t.pendingStartIndex = null;
            t.pendingStartLane = 0;
            t.pendingMainStartDate = null;
            t.pendingMainStartIndex = null;
        });
        renderAllSegments();
    });
}

function updateBottomRowBorders() {
    const rows = rowsContainer.querySelectorAll(".task-row");
    if (rows.length === 0) return;
    rows.forEach(r => r.classList.remove("is-last-visible"));
    let last = null;
    for (let i = rows.length - 1; i >= 0; i--) {
        if (!rows[i].classList.contains("task-hidden")) { last = rows[i]; break; }
    }
    if (last) last.classList.add("is-last-visible");

    const leftRows = leftRowsContainer.querySelectorAll(".left-row");
    leftRows.forEach(r => r.classList.remove("is-last-visible"));
    let lastLeft = null;
    for (let i = leftRows.length - 1; i >= 0; i--) {
        if (!leftRows[i].classList.contains("task-hidden")) { lastLeft = leftRows[i]; break; }
    }
    if (lastLeft) lastLeft.classList.add("is-last-visible");
}

function handleCellClick(task, index, y = SUB_SCHEDULE_TOP) {
    if (isCtrlSelectionMode) return;
    if (task.isSubCollapsed) return;
    const clickedIso = timelineDays[index].iso;
    task.pendingMainStartIndex = null;
    task.pendingMainStartDate = null;

    if (activeProgressSegmentId) {
        const targetSeg = findMainScheduleById(task, activeProgressSegmentId)
            || task.segments.find(s => s.id === activeProgressSegmentId);
        if (targetSeg) {
            targetSeg.progressEndDate = clickedIso;
            activeProgressSegmentId = null; 
            activeProgressTaskId = null;
            renderAllSegments(); 
            triggerSave();
        } else {
            alert("選択中のバーはこの行にありません。");
        }
    } else {
        if (task.pendingStartIndex === null) {
            task.pendingStartIndex = index;
            task.pendingStartDate = clickedIso;
            task.pendingStartLane = getSubScheduleLaneFromY(y);
            renderAllSegments();
        } else {
            const startIso = task.pendingStartDate;
            const endIso = clickedIso;
            const s = startIso < endIso ? startIso : endIso;
            const e = startIso < endIso ? endIso : startIso;
            
            const newSeg = {
                id: "seg_" + Date.now() + "_" + Math.random().toString(36).slice(2),
                startDate: s, endDate: e, type: "range", 
                label: "...", 
                progressEndDate: null, dailyValues: {}, dailyResults: {}
            };
            task.segments.push(newSeg);

            task.pendingStartIndex = null; 
            task.pendingStartDate = null;
            task.pendingStartLane = 0;
            
            renderAllSegments();

            setTimeout(() => {
                const initialLabel = "新規作業";
                const inputLabel = prompt("計画内容を入力してください:", initialLabel);
                
                if (inputLabel === null) {
                    task.segments.pop(); 
                    renderAllSegments();
                } else {
                    newSeg.label = (inputLabel.trim() === "") ? initialLabel : inputLabel;
                    renderAllSegments();
                    triggerSave();
                }
            }, 10);
        }
    }
}

function handleSegClick(task, seg, addMode) {
    if (isOverlayCancelStateActive()) return;
    if (!seg.id) return;
    activeTaskId = task.id;
    if (addMode && !isCtrlSelectionMode) {
        updateCtrlSelectionMode(true);
        seg.isSelected = true;
        renderAllSegments();
        return;
    }
    if (isCtrlSelectionMode) {
        if (addMode) {
            seg.isSelected = !seg.isSelected;
        } else if (!seg.isSelected) {
            seg.isSelected = true;
        }
        renderAllSegments();
        return;
    }
}

function addSegEvents(el, task, seg) {
    el.addEventListener("click", (e) => {
        e.stopPropagation();
        if (isOverlayCancelStateActive()) return;
        if (el.classList.contains("segment-label") && !isCtrlSelectionMode && !e.ctrlKey) {
            editSubSegmentLabel(task, seg);
            return;
        }
        handleSegClick(task, seg, e.ctrlKey || isCtrlSelectionMode);
    });
    
    el.addEventListener("dblclick", (e) => {
        if (isOverlayCancelStateActive()) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        if (isCtrlSelectionMode || e.ctrlKey) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        e.preventDefault();
        e.stopPropagation();
        
        if (window.getSelection) {
            window.getSelection().removeAllRanges();
        }
        
        editSubSegmentLabel(task, seg);
    });
    
    el.addEventListener("contextmenu", (e) => {
        if (activeProgressSegmentId) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        if (isCtrlSelectionMode || e.ctrlKey) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        e.preventDefault();
        e.stopPropagation(); 
        showSegmentContextMenu(e, task, seg);
    });
}

function handleDailyValueClick(e, task, seg, iso) {
    e.stopPropagation();
    if (isOverlayCancelStateActive()) return;
    if (isCtrlSelectionMode || e.ctrlKey) return;
    if (
        activeDailyValueTarget &&
        activeDailyValueTarget.taskId === task.id &&
        activeDailyValueTarget.segId === seg.id &&
        activeDailyValueTarget.iso === iso
    ) {
        clearActiveDailyValueTarget();
        renderAllSegments();
        return;
    }
    editDailyValue(task, seg, iso);
}

// ============================================
// コンテキストメニューなど
// ============================================
function normalizeRowColor(value) {
    return ROW_COLOR_PALETTE.some(c => c.value === value) ? value : "";
}

// 行の色を項目欄に反映する
function applyTaskRowColor(task) {
    if (!task?.leftRowEl) return;
    const color = normalizeRowColor(task.color);
    task.leftRowEl.classList.toggle("has-row-color", !!color);
    if (color) task.leftRowEl.style.setProperty("--row-color", color);
    else task.leftRowEl.style.removeProperty("--row-color");
}

// 右クリックメニュー内の色見本を描く
function renderRowColorPalette(task) {
    const palette = document.getElementById("cmColorPalette");
    if (!palette) return;
    palette.innerHTML = "";
    ROW_COLOR_PALETTE.forEach(c => {
        const sw = document.createElement("div");
        sw.className = "menu-color-swatch" + ((task.color || "") === c.value ? " is-current" : "");
        sw.title = c.name;
        if (c.value) sw.style.background = c.value;
        else { sw.style.background = "#fff"; sw.textContent = "×"; }
        sw.addEventListener("click", () => {
            const t = taskObjects.find(x => x.id === contextMenuTargetTaskId);
            if (!t) return;
            t.color = c.value;
            applyTaskRowColor(t);
            hideContextMenus();
            triggerSave();
        });
        palette.appendChild(sw);
    });
}

function showContextMenu(e, taskId) {
    contextMenuTargetTaskId = taskId;
    const task = taskObjects.find(t => t.id === taskId);
    const hideBtn = document.getElementById("cmHide");
    const unhideBtn = document.getElementById("cmUnhide");
    if (task.isHidden) { hideBtn.style.display = "none"; unhideBtn.style.display = "block"; }
    else { hideBtn.style.display = "block"; unhideBtn.style.display = "none"; }
    renderRowColorPalette(task);

    segmentContextMenu.style.display = "none";
    if (headerContextMenu) headerContextMenu.style.display = "none";
    contextMenu.style.display = "block";
    // 画面の下端・右端ではみ出さないように収める
    const w = contextMenu.offsetWidth;
    const h = contextMenu.offsetHeight;
    const vw = window.innerWidth || document.documentElement.clientWidth;
    const vh = window.innerHeight || document.documentElement.clientHeight;
    let x = e.pageX;
    let y = e.pageY;
    if (vw > 0) x = Math.max(4, Math.min(x, vw - w - 4));
    if (vh > 0) y = Math.max(4, Math.min(y, vh - h - 4));
    contextMenu.style.left = x + "px";
    contextMenu.style.top = y + "px";
}

function showSegmentContextMenu(e, task, seg) {
    contextMenuTargetSegId = seg.id;
    contextMenuTargetTaskForSeg = task;
    segmentContextMenu.dataset.scope = seg.scheduleScope || "sub";
    contextMenuTargetSegAnchor = seg.anchor || null;

    const depBtn = document.getElementById("ctxMainDependency");
    const depCancelBtn = document.getElementById("ctxMainDependencyCancel");
    const showDependencyControls = segmentContextMenu.dataset.scope === "main" && !!contextMenuTargetSegAnchor;
    if (depBtn) {
        depBtn.style.display = showDependencyControls ? "flex" : "none";
        depBtn.textContent = dependencyDraft ? "🔗 依存線を引き直す" : "🔗 依存線を開始";
    }
    if (depCancelBtn) {
        depCancelBtn.style.display = (showDependencyControls && dependencyDraft) ? "flex" : "none";
    }
    
    contextMenu.style.display = "none";
    if (dependencyContextMenu) dependencyContextMenu.style.display = "none";
    if (headerContextMenu) headerContextMenu.style.display = "none";
    segmentContextMenu.style.display = "block";
    segmentContextMenu.style.left = e.pageX + "px";
    segmentContextMenu.style.top = e.pageY + "px";
}

document.addEventListener("click", () => { 
    if (!isCtrlSelectionMode) {
        clearSegmentSelection();
        renderAllSegments();
    }
    hideContextMenus();
});

document.getElementById("cmComplete").addEventListener("click", () => {
    const t = taskObjects.find(t => t.id === contextMenuTargetTaskId);
    if (t) {
        t.isDone = !t.isDone;
        t.leftRowEl.classList.toggle("task-done", t.isDone);
        t.rowEl.classList.toggle("task-done", t.isDone);
        scheduleProgressGuideRefresh();
        triggerSave();
    }
});
document.getElementById("cmHide").addEventListener("click", () => {
    const t = taskObjects.find(t => t.id === contextMenuTargetTaskId);
    if (t) { t.isHidden = true; t.leftRowEl.classList.add("task-hidden"); t.rowEl.classList.add("task-hidden"); triggerSave(); }
});
document.getElementById("cmUnhide").addEventListener("click", () => {
    const t = taskObjects.find(t => t.id === contextMenuTargetTaskId);
    if (t) { t.isHidden = false; t.leftRowEl.classList.remove("task-hidden"); t.rowEl.classList.remove("task-hidden"); triggerSave(); }
});

document.getElementById("ctxSegProgress").addEventListener("click", () => {
    if (contextMenuTargetSegId) {
        activeProgressSegmentId = contextMenuTargetSegId;
        activeProgressTaskId = contextMenuTargetTaskForSeg ? contextMenuTargetTaskForSeg.id : null;
        renderAllSegments();
    }
});

document.getElementById("ctxMainDependency").addEventListener("click", () => {
    if (!contextMenuTargetTaskForSeg || !contextMenuTargetSegId || !contextMenuTargetSegAnchor) return;
    const main = findMainScheduleById(contextMenuTargetTaskForSeg, contextMenuTargetSegId);
    if (!main) return;
    beginDependencyDraft(contextMenuTargetTaskForSeg, main, contextMenuTargetSegAnchor);
    hideContextMenus();
});

document.getElementById("ctxMainDependencyCancel").addEventListener("click", () => {
    clearDependencyDraft();
    hideContextMenus();
});

document.getElementById("ctxSegDelete").addEventListener("click", () => {
    if (contextMenuTargetTaskForSeg && contextMenuTargetSegId) {
        if (confirm("選択の計画を削除しますか？")) {
            if (segmentContextMenu.dataset.scope === "main") {
                contextMenuTargetTaskForSeg.mainSchedules = (contextMenuTargetTaskForSeg.mainSchedules || [])
                    .filter(main => main.id !== contextMenuTargetSegId);
                if (activeProgressSegmentId === contextMenuTargetSegId) {
                    activeProgressSegmentId = null;
                    activeProgressTaskId = null;
                }
            } else {
                contextMenuTargetTaskForSeg.segments = contextMenuTargetTaskForSeg.segments.filter(s => s.id !== contextMenuTargetSegId);
                if (activeProgressSegmentId === contextMenuTargetSegId) {
                    activeProgressSegmentId = null;
                    activeProgressTaskId = null;
                }
            }
            renderAllSegments(); 
            triggerSave();
        }
    }
});

document.getElementById("ctxHeaderAddLeft").addEventListener("click", () => {
    if (contextMenuTargetHeaderIndex === null) return;
    const index = contextMenuTargetHeaderIndex;
    hideContextMenus();
    addItemColumn(index);
});
document.getElementById("ctxHeaderAddRight").addEventListener("click", () => {
    if (contextMenuTargetHeaderIndex === null) return;
    const index = contextMenuTargetHeaderIndex + 1;
    hideContextMenus();
    addItemColumn(index);
});
document.getElementById("ctxHeaderDelete").addEventListener("click", () => {
    if (contextMenuTargetHeaderIndex === null) return;
    const index = contextMenuTargetHeaderIndex;
    hideContextMenus();
    removeItemColumn(index);
});

document.getElementById("ctxDependencyDelete").addEventListener("click", () => {
    if (!contextMenuTargetDependencyId) return;
    if (!confirm("この依存線を削除しますか？")) return;
    appData.dependencies = ensureDependenciesArray().filter(dep => dep.id !== contextMenuTargetDependencyId);
    contextMenuTargetDependencyId = null;
    hideContextMenus();
    renderDependencies();
    triggerSave();
});

showHiddenCheck.addEventListener("change", (e) => {
    document.body.classList.toggle("show-hidden-mode", e.target.checked);
});
guideModeSelect.addEventListener("change", (e) => {
    appData.settings.guideMode = normalizeGuideMode(e.target.value, true);
    appData.settings.showMainLine = appData.settings.guideMode !== "off";
    buildHeader();
    renderAllSegments();
    scheduleProgressGuideRefresh();
    triggerSave();
});
if (collapseAllSubsBtn) {
    collapseAllSubsBtn.addEventListener("click", () => setAllTaskSubSchedulesCollapsed(true));
}
if (expandAllSubsBtn) {
    expandAllSubsBtn.addEventListener("click", () => setAllTaskSubSchedulesCollapsed(false));
}
if (memoToggleBtn) {
    memoToggleBtn.addEventListener("click", () => {
        setMemoCollapsed(!(appData.settings.memoCollapsed === true));
    });
}
if (memoSplitter) {
    memoSplitter.addEventListener("mousedown", (e) => {
        if (e.button !== 0) return;
        if (appData.settings.memoCollapsed === true) return;
        e.preventDefault();
        isResizingMemo = true;
        memoResizeStartX = e.clientX;
        memoResizeStartWidth = freeMemoArea.getBoundingClientRect().width;
        memoSplitter.classList.add("dragging");
        document.body.style.cursor = "col-resize";
        document.body.style.userSelect = "none";
    });
    // ダブルクリックで既定の幅（全体の1/4）に戻す
    memoSplitter.addEventListener("dblclick", () => {
        if (appData.settings.memoCollapsed === true) return;
        appData.settings.memoWidth = 0;
        applyMemoWidth(0);
        scheduleProgressGuideRefresh();
        triggerSave();
    });
}
// ウインドウの大きさが変わったら、はみ出さないように収め直す
window.addEventListener("resize", () => {
    if (!appData?.settings?.memoWidth) return;
    applyMemoWidth(appData.settings.memoWidth);
    scheduleProgressGuideRefresh();
});
if (toggleHolidaysBtn) {
    toggleHolidaysBtn.addEventListener("click", () => {
        appData.settings.hideHolidays = appData.settings.hideHolidays !== true;
        refreshTimelineDisplay();
        triggerSave();
    });
}

// ============================================
// イベント設定
// ============================================
function setupControlEvents() {
    if (planSelect) {
        planSelect.addEventListener("change", async (e) => {
            const nextId = e.target.value;
            if (!nextId || nextId === currentPlanId) return;
            switchPlan(nextId, { saveCurrent: true, resetHistory: true });
            await DataManager.save(appStore);
        });
        planSelect.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            e.stopPropagation();
            planSelect.blur();
            renameCurrentPlanWithPrompt();
        }, true);
    }

    if (newPlanBtn) {
        newPlanBtn.addEventListener("click", async () => {
            syncDataModel();
            const nextNameRaw = prompt("新しい計画名:", "新しい計画");
            if (nextNameRaw === null) return;
            const nextName = nextNameRaw.trim() || `新しい計画 ${appStore.plans.length + 1}`;
            const entry = createPlanEntry(nextName);
            appStore.plans.push(entry);
            currentPlanId = entry.id;
            appStore.currentPlanId = entry.id;
            refreshPlanSelect();
            restoreFromData(clonePlanData(entry.data));
            HistoryManager.init(appData);
            await DataManager.save(appStore);
        });
    }

    if (deletePlanBtn) {
        deletePlanBtn.addEventListener("click", async () => {
            const currentEntry = getCurrentPlanEntry();
            if (!currentEntry) return;
            if (appStore.plans.length <= 1) {
                alert("最後の1件は削除できません。");
                return;
            }
            if (!confirm(`計画「${currentEntry.name}」を削除しますか？`)) return;
            const currentIndex = appStore.plans.findIndex(plan => plan.id === currentPlanId);
            appStore.plans = appStore.plans.filter(plan => plan.id !== currentPlanId);
            const fallback = appStore.plans[Math.max(0, currentIndex - 1)] || appStore.plans[0];
            currentPlanId = fallback.id;
            appStore.currentPlanId = fallback.id;
            refreshPlanSelect();
            restoreFromData(clonePlanData(fallback.data));
            HistoryManager.init(appData);
            await DataManager.save(appStore);
        });
    }

    document.addEventListener("mousedown", (e) => {
        const clickedInsideSegmentMenu = segmentContextMenu.style.display === "block" && segmentContextMenu.contains(e.target);
        const clickedInsideTaskMenu = contextMenu.style.display === "block" && contextMenu.contains(e.target);
        if (clickedInsideSegmentMenu || clickedInsideTaskMenu) return;

        if (segmentContextMenu.style.display === "block" || contextMenu.style.display === "block") {
            e.preventDefault();
            e.stopPropagation();
            dismissContextMenusWithClickSuppression();
            return;
        }

        if (!activeProgressSegmentId) return;
        if (isValidProgressClickTarget(e.target)) return;
        e.preventDefault();
        e.stopPropagation();
        cancelActiveProgressSelection();
    }, true);

    document.addEventListener("click", (e) => {
        if (suppressNextClickAfterMenuDismiss) {
            e.preventDefault();
            e.stopPropagation();
            suppressNextClickAfterMenuDismiss = false;
            return;
        }
        if (suppressNextClickAfterProgressCancel) {
            e.preventDefault();
            e.stopPropagation();
            suppressNextClickAfterProgressCancel = false;
            return;
        }
        const clickedInsideSegmentMenu = segmentContextMenu.style.display === "block" && segmentContextMenu.contains(e.target);
        const clickedInsideTaskMenu = contextMenu.style.display === "block" && contextMenu.contains(e.target);
        if (clickedInsideSegmentMenu || clickedInsideTaskMenu) return;

        if (segmentContextMenu.style.display === "block" || contextMenu.style.display === "block") {
            e.preventDefault();
            e.stopPropagation();
            return;
        }

        if (!activeProgressSegmentId) return;
        if (isValidProgressClickTarget(e.target)) return;
        e.preventDefault();
        e.stopPropagation();
    }, true);

    document.addEventListener("keydown", (e) => {
        const tag = document.activeElement?.tagName;
        const isEditable = document.activeElement?.isContentEditable;
        if (tag === "INPUT" || tag === "TEXTAREA" || isEditable) return;

        if (dependencyDraft && e.key === "Escape") {
            e.preventDefault();
            clearDependencyDraft();
            return;
        }

        const ref = getActiveDailyValueRef();
        if (!ref) return;

        if (e.key === "ArrowRight") {
            e.preventDefault();
            moveActiveDailyValue(1);
        } else if (e.key === "ArrowLeft") {
            e.preventDefault();
            moveActiveDailyValue(-1);
        } else if (e.key === "Enter") {
            e.preventDefault();
            editDailyValue(ref.task, ref.seg, ref.iso);
        } else if (e.key === "Escape") {
            e.preventDefault();
            clearActiveDailyValueTarget();
            renderAllSegments();
        }
    });

    window.addEventListener("blur", () => updateCtrlSelectionMode(false));
    document.addEventListener("mousemove", (e) => {
        if (!dependencyDraft) return;
        dependencyMousePosition = { x: e.clientX, y: e.clientY };
        renderDependencies();
    });
    document.addEventListener("click", (e) => {
        if (!isCtrlSelectionMode) return;
        if (isSubSegmentSelectableTarget(e.target)) return;
        e.preventDefault();
        e.stopPropagation();
        clearSegmentSelection();
        renderAllSegments();
        updateCtrlSelectionMode(false);
    }, true);

    document.getElementById("undoBtn").addEventListener("click", () => HistoryManager.undo());
    document.getElementById("redoBtn").addEventListener("click", () => HistoryManager.redo());

    document.getElementById("settingsButton").addEventListener("click", () => {
        document.getElementById("settingsStartDate").value = appData.settings.startDate;
        document.getElementById("settingsEndDate").value = appData.settings.endDate;
        document.getElementById("settingsHolidays").value = appData.settings.holidays.join(", ");
        document.getElementById("settingsVacations").value = (appData.settings.vacations || []).join(", ");
        settingsPanel.classList.remove("settings-hidden");
    });
    document.getElementById("settingsCancel").addEventListener("click", () => settingsPanel.classList.add("settings-hidden"));
    document.getElementById("settingsSave").addEventListener("click", () => {
        if (!confirm("期間を変更しますか？")) return;
        appData.settings.startDate = document.getElementById("settingsStartDate").value;
        appData.settings.endDate = document.getElementById("settingsEndDate").value;
        const hText = document.getElementById("settingsHolidays").value.trim();
        appData.settings.holidays = hText ? hText.split(",").map(s => s.trim()).filter(s => s) : [];
        const vText = document.getElementById("settingsVacations").value.trim();
        appData.settings.vacations = vText ? vText.split(",").map(s => s.trim()).filter(s => s) : [];
        settingsPanel.classList.add("settings-hidden");
        restoreFromData(appData); triggerSave();
    });
    document.getElementById("clearAllBtn").addEventListener("click", () => {
        if (confirm("現在の計画を破棄し、新しい空の計画を開始しますか？")) {
            const nextData = createEmptyPlanData(appData.projectName || "新しい計画");
            nextData.settings = { ...appData.settings };
            restoreFromData(nextData);
            activeTaskId = null; activeProgressSegmentId = null; triggerSave();
            settingsPanel.classList.add("settings-hidden");
        }
    });

    document.getElementById("downloadBtn").addEventListener("click", () => {
        syncDataModel();
        const blob = new Blob([JSON.stringify(appData, null, 2)], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a"); a.href = url;
        const safeName = (appData.projectName || "schedule").replace(/[\\/:*?"<>|]/g, "_");
        a.download = `${safeName}_${formatTimestamp(new Date())}.json`;
        document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    });
    const fileInput = document.getElementById("fileInput");
    document.getElementById("uploadBtn").addEventListener("click", () => { fileInput.click(); });
    fileInput.addEventListener("change", (e) => {
        const file = e.target.files[0]; if (!file) return;
        if (file.size > MAX_IMPORT_FILE_BYTES) {
            alert(`ファイルサイズが大きすぎます。${Math.floor(MAX_IMPORT_FILE_BYTES / (1024 * 1024))}MB 以下の JSON を読み込んでください。`);
            fileInput.value = "";
            return;
        }
        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const raw = JSON.parse(evt.target.result);
                let data = raw;
                if (Array.isArray(raw?.plans)) {
                    const importedStore = normalizeLoadedStore(raw);
                    const importedCurrent = importedStore.plans.find(plan => plan.id === importedStore.currentPlanId) || importedStore.plans[0];
                    data = clonePlanData(importedCurrent.data);
                }
                const validatedData = sanitizeImportedPlanData(data);
                if (confirm("現在の計画を上書きして読み込みますか？")) { restoreFromData(validatedData); triggerSave(); alert("完了"); }
            } catch (err) { alert("読込失敗"); }
            fileInput.value = "";
        };
        reader.readAsText(file);
    });

    const excelExportBtn = document.getElementById("excelExportBtn");
    const excelImportBtn = document.getElementById("excelImportBtn");
    const excelFileInput = document.getElementById("excelFileInput");
    if (excelExportBtn) excelExportBtn.addEventListener("click", exportToExcel);
    if (excelImportBtn && excelFileInput) {
        excelImportBtn.addEventListener("click", () => { excelFileInput.click(); });
        excelFileInput.addEventListener("change", (e) => {
            const file = e.target.files[0];
            excelFileInput.value = "";
            if (file) importExcelFile(file);
        });
    }

    const rowSelectHeader = document.getElementById("rowSelectHeader");
    if (rowSelectHeader) {
        rowSelectHeader.textContent = "□";

        rowSelectHeader.addEventListener("click", () => {
            selectionMode = (selectionMode + 1) % 3;
            
            if (selectionMode === 0) rowSelectHeader.textContent = "□";
            else if (selectionMode === 1) rowSelectHeader.textContent = "☑️";
            else if (selectionMode === 2) rowSelectHeader.textContent = "🔳"; 

            const delBtn = document.getElementById("deleteSelectedBtn");
            if(delBtn) delBtn.style.display = (selectionMode !== 0) ? "inline-block" : "none";

            taskObjects.forEach(t => {
                if (selectionMode === 1) t.isSelected = true; 
                else if (selectionMode === 2) t.isSelected = false;
                else t.isSelected = false;
                renderGrip(t);
            });
        });
    }

    const deleteSelectedBtn = document.getElementById("deleteSelectedBtn");
    if (deleteSelectedBtn) {
        deleteSelectedBtn.addEventListener("click", () => {
            const selectedTasks = taskObjects.filter(t => t.isSelected);
            if (selectedTasks.length === 0) {
                alert("削除対象が選択されていません。");
                return;
            }
            if (confirm(`${selectedTasks.length} 件の行を削除しますか？`)) {
                for (let i = taskObjects.length - 1; i >= 0; i--) {
                    if (taskObjects[i].isSelected) {
                        taskObjects[i].leftRowEl.remove();
                        taskObjects[i].rowEl.remove();
                        taskObjects.splice(i, 1);
                    }
                }
                triggerSave();
            }
        });
    }

    document.getElementById("todoColumnsInput").addEventListener("change", triggerSave);
    
    document.getElementById("outlookBtn").addEventListener("click", exportTodoToOutlookCSV);
    
    document.getElementById("todoCsvBtn").addEventListener("click", exportTodoToCSV);

    const todoBtn = document.getElementById("todoBtn");
    if(todoBtn) todoBtn.addEventListener("click", () => {
         const todoPanel = document.getElementById("todoPanel");
         currentTodoDate = new Date(); 
         todoPanel.classList.remove("settings-hidden"); 
         updateTodoTable(currentTodoDate);
    });
    
    initTodoFeature();

    if (taskMemoTextarea) {
        taskMemoTextarea.addEventListener("input", () => {
            let text = taskMemoTextarea.value || "";
            if (text.length > 3000) {
                text = text.slice(0, 3000);
                taskMemoTextarea.value = text;
            }
            updateMemoCount(text);
            const task = taskObjects.find(t => t.id === memoPanelTaskId);
            if (task) {
                task.memo = text;
                triggerSave();
            }
        });
        taskMemoTextarea.addEventListener("keydown", (e) => {
            if (e.key === "Escape") closeTaskMemoPanel();
        });
    }

    updateLeftResizeHandles();

    if (taskMemoClose) {
        taskMemoClose.addEventListener("click", () => closeTaskMemoPanel());
    }

    if (taskMemoPanel && taskMemoHeader) {
        let isDragging = false;
        let startX = 0;
        let startY = 0;
        let initL = 0;
        let initT = 0;

        taskMemoHeader.addEventListener("mousedown", (e) => {
            if (e.target.closest("button")) return;
            isDragging = true;
            startX = e.clientX;
            startY = e.clientY;
            const r = taskMemoPanel.getBoundingClientRect();
            initL = r.left;
            initT = r.top;
            document.body.style.userSelect = "none";
        });
        document.addEventListener("mousemove", (e) => {
            if (!isDragging) return;
            const left = initL + (e.clientX - startX);
            const top = initT + (e.clientY - startY);
            taskMemoPanel.style.left = `${left}px`;
            taskMemoPanel.style.top = `${top}px`;
        });
        document.addEventListener("mouseup", () => {
            if (!isDragging) return;
            isDragging = false;
            document.body.style.userSelect = "";
        });
    }

    document.addEventListener("mousemove", (e) => {
        if (isResizingMemo) {
            // 右へドラッグ＝メモ欄が狭くなる。0以下は「既定に戻す」の意味になるので下限で止める
            const delta = memoResizeStartX - e.clientX;
            applyMemoWidth(Math.max(MIN_MEMO_WIDTH, memoResizeStartWidth + delta));
            scheduleProgressGuideRefresh();
        } else if (isResizingCol) {
            const delta = e.clientX - resizeStartX;
            const newWidth = clampColumnWidth(resizeColIndex, resizeStartWidth + delta);
            leftColumnWidths[resizeColIndex] = newWidth;
            applyLeftColumnWidths();
            scheduleProgressGuideRefresh();
        } else if (isResizingRow) {
            const task = taskObjects.find(t => t.id === resizeRowTaskId);
            if (!task) return;
            const delta = e.clientY - resizeStartY;
            const baseHeight = task.baseHeight || computeTaskBaseHeight(task);
            const newHeight = Math.max(baseHeight, resizeStartHeight + delta);
            task.customHeight = newHeight;
            task.rowEl.style.height = newHeight + "px";
            task.leftRowEl.style.height = newHeight + "px";
        }
    });
    document.addEventListener("mouseup", () => {
        if (isResizingMemo) {
            isResizingMemo = false;
            memoSplitter.classList.remove("dragging");
            document.body.style.cursor = "";
            document.body.style.userSelect = "";
            appData.settings.memoWidth = Math.round(freeMemoArea.getBoundingClientRect().width);
            scheduleProgressGuideRefresh();
            triggerSave();
        }
        if (isResizingCol) {
            isResizingCol = false;
            resizeColIndex = null;
            document.body.style.cursor = "";
            document.body.style.userSelect = "";
            scheduleProgressGuideRefresh();
            triggerSave();
        }
        if (isResizingRow) {
            isResizingRow = false;
            resizeRowTaskId = null;
            document.body.style.cursor = "";
            document.body.style.userSelect = "";
            renderAllSegments();
            triggerSave();
        }
    });
}

function handleMainCellClick(task, index) {
    if (isCtrlSelectionMode) return;
    const clickedIso = timelineDays[index].iso;
    const clickedMain = findMainScheduleByDate(task, clickedIso);

    if (activeProgressSegmentId) {
        const targetMain = findMainScheduleById(task, activeProgressSegmentId);
        if (!targetMain) {
            alert("この行にメインスケジュールがありません。");
            return;
        }
        targetMain.progressEndDate = clickedIso;
        activeProgressSegmentId = null;
        activeProgressTaskId = null;
        renderAllSegments();
        triggerSave();
        return;
    }

    if (task.pendingMainStartIndex === null && !clickedMain) {
        task.pendingMainStartIndex = index;
        task.pendingMainStartDate = clickedIso;
        task.pendingStartIndex = null;
        task.pendingStartDate = null;
        renderAllSegments();
        return;
    }

    if (task.pendingMainStartIndex !== null) {
        const startIso = task.pendingMainStartDate;
        const endIso = clickedIso;
        const s = startIso < endIso ? startIso : endIso;
        const e = startIso < endIso ? endIso : startIso;
        task.pendingMainStartIndex = null;
        task.pendingMainStartDate = null;

        if (hasOverlappingMainSchedule(task, s, e)) {
            alert("メイン計画は重複して設定できません。既存のメイン計画の外側で設定してください。");
            renderAllSegments();
            return;
        }

        const newMain = {
            id: "main_" + Date.now() + "_" + Math.random().toString(36).slice(2),
            startDate: s,
            endDate: e,
            label: "メイン計画",
            startLabel: "",
            endLabel: "",
            progressEndDate: null,
            milestones: []
        };
        task.mainSchedules.push(newMain);
        task.mainSchedules.sort((a, b) => a.startDate.localeCompare(b.startDate));
        renderAllSegments();

        setTimeout(() => {
            const initialLabel = "メイン計画";
            const inputLabel = prompt("計画内容を入力してください:", initialLabel);

            if (inputLabel === null) {
                task.mainSchedules = task.mainSchedules.filter(m => m.id !== newMain.id);
                renderAllSegments();
            } else {
                newMain.label = (inputLabel.trim() === "") ? initialLabel : inputLabel.trim();
                renderAllSegments();
                triggerSave();
            }
        }, 10);
        return;
    }

    const main = clickedMain;
    if (!main) {
        alert("既存メイン計画の外側をクリックすると新しいメイン計画を追加できます。範囲内をクリックするとマイルストーンを追加できます。");
        return;
    }

    let milestone = (main.milestones || []).find(ms => ms.date === clickedIso);
    const defaultLabel = milestone ? (milestone.label || "マイルストーン") : "マイルストーン";
    const labelInput = prompt("マイルストーン名を入力してください:", defaultLabel);
    if (labelInput === null) return;

    if (!milestone) {
        milestone = {
            id: "ms_" + Date.now() + "_" + Math.random().toString(36).slice(2),
            date: clickedIso,
            label: labelInput.trim() || "マイルストーン"
        };
        if (!Array.isArray(main.milestones)) main.milestones = [];
        main.milestones.push(milestone);
        main.milestones.sort((a, b) => a.date.localeCompare(b.date));
    } else {
        milestone.label = labelInput.trim() || "マイルストーン";
    }
    renderAllSegments();
    triggerSave();
}

function updateTodoTable(dateObj) {
    const todoDateDisplay = document.getElementById("todoDateDisplay");
    todoDateDisplay.textContent = `${dateToISO(dateObj)} (${WEEKDAYS[dateObj.getDay()]})`;
    
    const table = document.querySelector(".todo-table");
    let colgroup = table.querySelector("colgroup");
    if (colgroup) table.removeChild(colgroup); 
    colgroup = document.createElement("colgroup");
    table.insertBefore(colgroup, table.firstChild);
    
    const colWidths = [
        "40px",  // 選択
        "15%",   // 項目1
        "15%",   // 項目2
        "80px",  // 担当
        "auto",  // 実施内容
        "70px",  // 計画
        "70px"   // 実績
    ];
    colWidths.forEach(w => {
        const col = document.createElement("col");
        col.style.width = w;
        colgroup.appendChild(col);
    });

    const h1 = document.getElementById("lh1").textContent;
    const h2 = document.getElementById("lh2").textContent;
    const h3 = document.getElementById("lh3").textContent;
    
    const displayCols = ["□", h1, h2, h3, "実施内容", "計画", "実績"];
    const colKeys = ["select", "h1", "h2", "h3", "desc", "plan", "actual"];

    const thead = document.getElementById("todoThead");
    thead.innerHTML = "";
    const trH = document.createElement("tr");

    displayCols.forEach((colName, idx) => {
        const th = document.createElement("th");
        th.textContent = colName;
        if (idx === 0) {
            th.style.cursor = "pointer";
            th.style.textAlign = "center";
            th.textContent = todoSelectionState ? "☑️" : "□";
            th.addEventListener("click", () => {
                todoSelectionState = !todoSelectionState; 
                th.textContent = todoSelectionState ? "☑️" : "□";
                
                const checkboxes = document.querySelectorAll(".todo-row-checkbox");
                checkboxes.forEach(cb => {
                    cb.textContent = todoSelectionState ? "☑️" : "□";
                    cb.dataset.checked = todoSelectionState ? "true" : "false";
                });
                checkTodoDeleteBtnVisibility();
            });
        }
        trH.appendChild(th);
    });
    thead.appendChild(trH);

    const iso = dateToISO(dateObj);
    const tbody = document.getElementById("todoTableBody"); 
    tbody.innerHTML = "";
    let hasItem = false;

    // [修正] 合計計算用変数
    let totalPlan = 0;
    let totalActual = 0;

    taskObjects.forEach(task => {
        if (task.isHidden) return;
        
        const editable1 = task.leftRowEl.children[1].querySelector(".editable");
        const editable2 = task.leftRowEl.children[2].querySelector(".editable");
        const editable3 = task.leftRowEl.children[3].querySelector(".editable");

        const t1 = editable1.textContent;
        const t2 = editable2.textContent;
        const t3 = editable3.textContent;

        task.segments.forEach(seg => {
            if (seg.startDate <= iso && seg.endDate >= iso) {
                const tr = document.createElement("tr");
                tr.dataset.taskId = task.id;
                tr.dataset.segId = seg.id;

                // [修正] 合計計算
                const pv = (seg.dailyValues || {})[iso];
                const av = (seg.dailyResults || {})[iso];
                if (pv && !isNaN(parseFloat(pv))) totalPlan += parseFloat(pv);
                if (av && !isNaN(parseFloat(av))) totalActual += parseFloat(av);

                colKeys.forEach((key) => {
                    const td = document.createElement("td");
                    
                    if (key === "select") {
                        td.style.textAlign = "center";
                        td.style.cursor = "pointer";
                        td.className = "todo-row-checkbox";
                        td.textContent = todoSelectionState ? "☑️" : "□"; 
                        td.dataset.checked = todoSelectionState ? "true" : "false";
                        
                        td.addEventListener("click", (e) => {
                            e.stopPropagation();
                            const isChecked = td.dataset.checked === "true";
                            td.dataset.checked = isChecked ? "false" : "true";
                            td.textContent = (td.dataset.checked === "true") ? "☑️" : "□";
                            checkTodoDeleteBtnVisibility();
                        });
                    } else {
                        const input = document.createElement("input");
                        let val = "";

                        if (key === "h1") {
                            val = t1;
                            input.addEventListener("change", (e) => { editable1.textContent = e.target.value; triggerSave(); });
                        }
                        else if (key === "h2") {
                            val = t2;
                            input.addEventListener("change", (e) => { editable2.textContent = e.target.value; triggerSave(); });
                        }
                        else if (key === "h3") {
                            val = t3;
                            input.style.textAlign = "center";
                            input.addEventListener("change", (e) => { editable3.textContent = e.target.value; triggerSave(); });
                        }
                        else if (key === "desc") {
                            val = seg.label || "";
                            input.addEventListener("change", (e) => { seg.label = e.target.value; renderAllSegments(); triggerSave(); });
                        }
                        else if (key === "plan") {
                            val = (seg.dailyValues || {})[iso] || "";
                            input.style.textAlign = "center";
                            input.addEventListener("change", (e) => {
                                if(!seg.dailyValues) seg.dailyValues = {};
                                seg.dailyValues[iso] = e.target.value;
                                if(!e.target.value) delete seg.dailyValues[iso];
                                renderAllSegments();
                                triggerSave();
                                updateTodoTable(dateObj); // 合計再計算のため
                            });
                        }
                        else if (key === "actual") {
                            val = (seg.dailyResults || {})[iso] || "";
                            input.style.textAlign = "center";
                            input.addEventListener("change", (e) => {
                                if(!seg.dailyResults) seg.dailyResults = {};
                                seg.dailyResults[iso] = e.target.value;
                                if(!e.target.value) delete seg.dailyResults[iso];
                                triggerSave();
                                updateTodoTable(dateObj); // 合計再計算のため
                            });
                        }
                        input.value = val;
                        td.appendChild(input);
                    }
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
                hasItem = true;
            }
        });
    });
    document.getElementById("todoEmptyMsg").style.display = hasItem ? "none" : "block";

    // [修正] 合計行（tfoot）の追加
    let tfoot = table.querySelector("tfoot");
    if(tfoot) table.removeChild(tfoot);
    tfoot = document.createElement("tfoot");
    const trF = document.createElement("tr");
    
    const tdLabel = document.createElement("td");
    tdLabel.colSpan = 5; 
    tdLabel.textContent = "合計";
    tdLabel.style.textAlign = "right";
    trF.appendChild(tdLabel);

    const tdPlan = document.createElement("td");
    tdPlan.textContent = (totalPlan % 1 === 0) ? totalPlan : totalPlan.toFixed(1);
    tdPlan.style.textAlign = "center";
    trF.appendChild(tdPlan);

    const tdActual = document.createElement("td");
    tdActual.textContent = (totalActual % 1 === 0) ? totalActual : totalActual.toFixed(1);
    tdActual.style.textAlign = "center";
    trF.appendChild(tdActual);

    tfoot.appendChild(trF);
    table.appendChild(tfoot);

    checkTodoDeleteBtnVisibility();
}

function checkTodoDeleteBtnVisibility() {
    const delBtn = document.getElementById("todoDeleteBtn");
    if(!delBtn) return;
    const checkedItems = document.querySelectorAll(".todo-row-checkbox[data-checked='true']");
    delBtn.style.display = (checkedItems.length > 0) ? "inline-block" : "none";
}

function exportTodoToOutlookCSV() {
    const iso = dateToISO(currentTodoDate);
    const dateStr = iso.replace(/-/g, '/');

    const items = [];
    taskObjects.forEach(task => {
        if (task.isHidden) return;
        const t1 = (getTaskLabelEditables(task)[0] || {}).textContent || "";
        task.segments.forEach(seg => {
            if (seg.startDate <= iso && seg.endDate >= iso) {
                items.push({
                    item1: t1,
                    desc: seg.label || ""
                });
            }
        });
    });

    if (items.length === 0) {
        alert("出力するデータがありません");
        return;
    }

    const headers = ["件名","開始日","開始時刻","終了日","終了時刻","プライベート","公開する時間帯の種類","秘密度","優先度"];
    const rows = [];

    let currentMin = 510; 

    items.forEach((item) => {
        const subject = `${item.item1}：${item.desc}`;

        const hStart = Math.floor(currentMin / 60);
        const mStart = currentMin % 60;
        const startTimeStr = `${hStart}:${pad2(mStart)}:00`;

        const endMin = currentMin + 30;
        const hEnd = Math.floor(endMin / 60);
        const mEnd = endMin % 60;
        const endTimeStr = `${hEnd}:${pad2(mEnd)}:00`;

        currentMin += 30;

        const rowData = [
            `"${subject.replace(/"/g, '""')}"`,
            `"${dateStr}"`,
            `"${startTimeStr}"`,
            `"${dateStr}"`,
            `"${endTimeStr}"`,
            `"FALSE"`,
            `"2"`,
            `"標準"`,
            `"標準"`
        ];
        rows.push(rowData.join(","));
    });

    const csvContent = headers.map(h => `"${h}"`).join(",") + "\r\n" + rows.join("\r\n");

    downloadAsShiftJIS(csvContent, `Outlook_${formatTimestamp(new Date())}.csv`);
}

function exportTodoToCSV() {
    const filename = `ToDoList_${formatTimestamp(new Date())}.csv`;
    const tbody = document.getElementById("todoTableBody");

    const columnsRaw = document.getElementById("todoColumnsInput").value || "";
    let columns = columnsRaw.split(",").map(s => s.trim()).filter(Boolean);
    if (columns.length === 0) {
        columns = getDefaultTodoColumns().split(",").map(s => s.trim());
    }
    columns = normalizeTodoColumns(columns);

    const headers = columns.map(c => '"' + c.replace(/"/g, '""') + '"');
    const iso = dateToISO(currentTodoDate);

    const rows = [];
    tbody.querySelectorAll("tr").forEach(tr => {
        const taskId = tr.dataset.taskId;
        const segId = tr.dataset.segId;
        const task = taskObjects.find(t => t.id === taskId);
        const seg = task ? task.segments.find(s => s.id === segId) : null;

        const t1 = task ? task.leftRowEl.children[1].querySelector(".editable").textContent : "";
        const t2 = task ? task.leftRowEl.children[2].querySelector(".editable").textContent : "";
        const t3 = task ? task.leftRowEl.children[3].querySelector(".editable").textContent : "";
        const desc = seg ? (seg.label || "") : "";
        const plan = seg && seg.dailyValues ? (seg.dailyValues[iso] || "") : "";
        const actual = seg && seg.dailyResults ? (seg.dailyResults[iso] || "") : "";
        const memo = task ? (task.memo || "") : "";

        const rowData = columns.map(col => {
            let v = "";
            if (col === "項目1") v = t1;
            else if (col === "項目2") v = t2;
            else if (col === "担当") v = t3;
            else if (col === "実施内容") v = desc;
            else if (col === "計画") v = plan;
            else if (col === "実績") v = actual;
            else if (col === "メモ") v = memo;
            return '"' + String(v).replace(/"/g, '""') + '"';
        });
        rows.push(rowData.join(","));
    });

    if (rows.length === 0) {
        alert("出力するデータがありません");
        return;
    }

    const csvContent = headers.join(",") + "\r\n" + rows.join("\r\n");
    downloadAsShiftJIS(csvContent, filename);
}

function downloadAsShiftJIS(content, filename) {
    if (typeof Encoding === "undefined") {
        alert("文字コード変換ライブラリが読み込まれていません。インターネット接続を確認してください。\nとりあえずUTF-8(BOM付)で出力します。");
        const blob = new Blob(["\uFEFF" + content], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        return;
    }

    const unicodeList = [];
    for (let i = 0; i < content.length; i++) {
        unicodeList.push(content.charCodeAt(i));
    }
    
    const sjisCodeList = Encoding.convert(unicodeList, {
        to: 'SJIS',
        from: 'UNICODE'
    });
    
    const u8Array = new Uint8Array(sjisCodeList);
    const blob = new Blob([u8Array], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ============================================
// Excel入出力
// ============================================
// 「表示されている日」か（休日非表示のときは土日・会社休日を除く）
// settings を省略すると現在の計画の設定を使う（インポート時は読込先の計画の設定を渡す）
function isVisibleDay(iso, settings = appData.settings) {
    if (settings.hideHolidays !== true) return true;
    const dow = isoToDate(iso).getDay();
    if (dow === 0 || dow === 6) return false;
    return !(settings.holidays || []).includes(iso);
}

// 期間内で表示されている日の一覧。全日非表示なら暦日で返す
function listVisibleDays(startIso, endIso, settings = appData.settings) {
    const s = startIso <= endIso ? startIso : endIso;
    const e = startIso <= endIso ? endIso : startIso;
    const all = [];
    const visible = [];
    const cur = isoToDate(s);
    const end = isoToDate(e);
    let guard = 0;
    while (cur <= end && guard < 5000) {
        const iso = dateToISO(cur);
        all.push(iso);
        if (isVisibleDay(iso, settings)) visible.push(iso);
        cur.setDate(cur.getDate() + 1);
        guard++;
    }
    return visible.length > 0 ? visible : all;
}

// 進捗日 -> 1〜10 の段階（表示日数に対する消化日数の割合）
function progressToScale(startIso, endIso, progressIso) {
    if (!progressIso) return null;
    const days = listVisibleDays(startIso, endIso);
    const done = days.filter(iso => iso <= progressIso).length;
    if (done <= 0) return null;
    return Math.max(1, Math.min(10, Math.round((done / days.length) * 10)));
}

// 1〜10 の段階 -> 進捗日
function scaleToProgressIso(startIso, endIso, scale, settings = appData.settings) {
    const n = Number(scale);
    if (!Number.isFinite(n) || n <= 0) return null;
    const days = listVisibleDays(startIso, endIso, settings);
    const count = Math.max(1, Math.min(days.length, Math.round((Math.min(n, 10) / 10) * days.length)));
    return days[count - 1];
}

function isValidIso(iso) {
    if (typeof iso !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return false;
    return dateToISO(isoToDate(iso)) === iso;
}

// Excelセルの値（日付型・シリアル値・文字列）を YYYY-MM-DD に変換
function excelCellToIso(value, settings = appData.settings) {
    if (value == null || value === "") return null;
    if (value instanceof Date) {
        if (Number.isNaN(value.getTime())) return null;
        return dateToISO(value);
    }
    if (typeof value === "number") {
        const parsed = XLSX.SSF.parse_date_code(value);
        if (!parsed) return null;
        const iso = `${parsed.y}-${pad2(parsed.m)}-${pad2(parsed.d)}`;
        return isValidIso(iso) ? iso : null;
    }
    const text = String(value).trim();
    let m = text.match(/^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
    if (m) {
        const iso = `${m[1]}-${pad2(Number(m[2]))}-${pad2(Number(m[3]))}`;
        return isValidIso(iso) ? iso : null;
    }
    m = text.match(/^(\d{1,2})[\/\-.月](\d{1,2})日?$/);
    if (m) {
        const year = isoToDate(settings.startDate).getFullYear();
        const iso = `${year}-${pad2(Number(m[1]))}-${pad2(Number(m[2]))}`;
        return isValidIso(iso) ? iso : null;
    }
    return null;
}

function excelDate(iso) {
    return iso ? isoToDate(iso) : "";
}

function cellText(value) {
    if (value == null) return "";
    return String(value).trim();
}

// 「日程表：202606_202609122224.xlsx」-> 「202606」
// 先頭の「日程表：」と、末尾に付けた出力日時（_YYYYMMDDHHMM）を取り除く
function planNameFromExcelFileName(fileName) {
    let name = String(fileName || "").replace(/\.xlsx$/i, "");
    name = name.replace(/^日程表[：:]/, "");
    name = name.replace(/_\d{12}$/, "");
    return name.trim();
}

// ---- エクスポート ----
// 行ごとの列構成: 項目列の並びに「項目1メモ」を差し込む（2列目の後ろ。項目列が1つなら1列目の後ろ）
function buildExcelTaskColumns(headers) {
    const cols = headers.map((name, index) => ({ kind: "item", index, name }));
    cols.splice(Math.min(2, headers.length), 0, { kind: "memo", name: EXCEL_COL.memo });
    return cols;
}

function buildExcelPlanRows() {
    const headers = getHeaderTexts();
    const taskColumns = buildExcelTaskColumns(headers);
    const rows = [[
        EXCEL_COL.done, EXCEL_COL.hidden, ...taskColumns.map(c => c.name),
        EXCEL_COL.type, EXCEL_COL.comment, EXCEL_COL.start, EXCEL_COL.end, EXCEL_COL.progress
    ]];
    const blankTaskCols = () => ["", "", ...taskColumns.map(() => "")];

    taskObjects.forEach(task => {
        const scheduleRows = [];

        [...(task.mainSchedules || [])]
            .sort((a, b) => a.startDate.localeCompare(b.startDate))
            .forEach(main => {
                const days = listVisibleDays(main.startDate, main.endDate);
                const milestones = [...(main.milestones || [])].sort((a, b) => a.date.localeCompare(b.date));
                const startMs = milestones.find(ms => ms.date === main.startDate) || null;
                const endMs = (main.endDate !== main.startDate)
                    ? (milestones.find(ms => ms.date === main.endDate) || null)
                    : null;
                scheduleRows.push([
                    EXCEL_TYPE.mainAll, main.label || "",
                    excelDate(days[0]), excelDate(days[days.length - 1]),
                    progressToScale(main.startDate, main.endDate, main.progressEndDate)
                ]);
                // 開始日・終了日のマイルストーンは M始/M終 として出力（無ければ端点コメント）
                scheduleRows.push([EXCEL_TYPE.mainStart, startMs ? (startMs.label || "") : (main.startLabel || ""), "", "", null]);
                scheduleRows.push([EXCEL_TYPE.mainEnd, endMs ? (endMs.label || "") : (main.endLabel || ""), "", "", null]);
                milestones
                    .filter(ms => ms !== startMs && ms !== endMs)
                    .forEach(ms => {
                        scheduleRows.push([EXCEL_TYPE.milestone, ms.label || "", excelDate(ms.date), excelDate(ms.date), null]);
                    });
            });

        [...(task.segments || [])]
            .sort((a, b) => (a.startDate !== b.startDate)
                ? a.startDate.localeCompare(b.startDate)
                : (a.endDate || a.startDate).localeCompare(b.endDate || b.startDate))
            .forEach(seg => {
                const endIso = seg.endDate || seg.startDate;
                const days = listVisibleDays(seg.startDate, endIso);
                scheduleRows.push([
                    EXCEL_TYPE.sub, seg.label || "",
                    excelDate(days[0]), excelDate(days[days.length - 1]),
                    progressToScale(seg.startDate, endIso, seg.progressEndDate)
                ]);
            });

        if (scheduleRows.length === 0) scheduleRows.push(["", "", "", "", null]);

        const labels = getTaskLabels(task);
        scheduleRows.forEach((scheduleRow, index) => {
            const taskCols = index === 0
                ? [
                    task.isDone ? EXCEL_COL.done : "",
                    task.isHidden ? EXCEL_COL.hidden : "",
                    ...taskColumns.map(c => (c.kind === "item" ? (labels[c.index] || "") : (task.memo || "")))
                ]
                : blankTaskCols();
            rows.push([...taskCols, ...scheduleRow]);
        });
    });
    return rows;
}

// メモ・備考: 1行 = 1行、タブ = 列
function buildExcelMemoRows() {
    const text = getFreeMemoText();
    if (!text) return [];
    return text.split(/\r?\n/).map(line => line.split("\t").map(cell => {
        const trimmed = cell.replace(/^[\s\u3000]+|[\s\u3000]+$/g, "");
        if (trimmed === "") return null;
        if (/^-?\d+(\.\d+)?$/.test(trimmed)) return Number(trimmed);
        return trimmed;
    }));
}

function buildExcelWorkbook() {
    syncDataModel();

    const headers = getHeaderTexts();
    const planRows = buildExcelPlanRows().map(row => row.map(v => (v === "" ? null : v)));
    const wsPlan = XLSX.utils.aoa_to_sheet(planRows, { cellDates: true });
    Object.keys(wsPlan).forEach(key => {
        if (key[0] !== "!" && wsPlan[key].t === "d") wsPlan[key].z = "yyyy/m/d";
    });
    wsPlan["!cols"] = [
        { wch: 4 }, { wch: 8 },
        ...buildExcelTaskColumns(headers).map(c => ({ wch: c.kind === "memo" ? 18 : 16 })),
        { wch: 7 }, { wch: 20 }, { wch: 12 }, { wch: 12 }, { wch: 11 }
    ];

    const memoRows = buildExcelMemoRows();
    const wsMemo = XLSX.utils.aoa_to_sheet(memoRows.length > 0 ? memoRows : [[null]]);
    wsMemo["!cols"] = [{ wch: 30 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsPlan, EXCEL_SHEET_PLAN);
    XLSX.utils.book_append_sheet(wb, wsMemo, EXCEL_SHEET_MEMO);
    return wb;
}

function exportToExcel() {
    if (typeof XLSX === "undefined") {
        alert("Excel出力ライブラリ (vendor/xlsx.full.min.js) が読み込まれていません。");
        return;
    }
    const wb = buildExcelWorkbook();
    const safeName = (appData.projectName || "schedule").replace(/[\\/:*?"<>|]/g, "_");
    // 日程表：<計画名>_YYYYMMDDHHMM.xlsx
    XLSX.writeFile(wb, `${EXCEL_FILE_PREFIX}${safeName}_${formatTimestamp(new Date())}.xlsx`);
}

// ---- インポート ----
// ワークブック -> 計画データ（restoreFromData に渡せる形）
// basePlan: 期間・休日・列幅などを引き継ぐ元の計画データ（読込先の計画。無ければ現在の計画）
function parseExcelWorkbook(wb, fileName, basePlan = appData) {
    const baseSettings = basePlan.settings || appData.settings;
    const planSheet = wb.Sheets[EXCEL_SHEET_PLAN] || wb.Sheets[wb.SheetNames[0]];
    if (!planSheet) throw new Error("「計画」シートが見つかりません。");

    const rows = XLSX.utils.sheet_to_json(planSheet, { header: 1, raw: true, defval: null });
    const headerRow = (rows[0] || []).map(cellText);
    const col = {};
    Object.entries(EXCEL_COL).forEach(([key, name]) => { col[key] = headerRow.indexOf(name); });
    if (col.type < 0) throw new Error(`1行目に「${EXCEL_COL.type}」列が見つかりません。`);

    // 項目列 = 「完」「非表示」の右から「M/S」の左まで（「項目1メモ」列は除く）
    const firstItemCol = Math.max(col.done, col.hidden) + 1;
    const itemCols = [];
    for (let c = firstItemCol; c < col.type; c++) {
        if (c !== col.memo) itemCols.push(c);
    }
    if (itemCols.length === 0) throw new Error("項目列が見つかりません（「非表示」と「M/S」の間に項目列を置いてください）。");
    if (itemCols.length > MAX_ITEM_COLUMNS) throw new Error(`項目列は最大${MAX_ITEM_COLUMNS}列までです（${itemCols.length}列あります）。`);
    const headers = normalizeHeaders(itemCols.map(c => headerRow[c]));

    const tasks = [];
    let currentTask = null;
    let currentMain = null;
    const warnings = [];
    let minIso = null;
    let maxIso = null;
    const noteDate = (iso) => {
        if (!iso) return;
        if (!minIso || iso < minIso) minIso = iso;
        if (!maxIso || iso > maxIso) maxIso = iso;
    };
    const readCell = (row, index) => (index >= 0 ? row[index] : null);

    for (let r = 1; r < rows.length; r++) {
        const row = rows[r] || [];
        const excelRowNo = r + 1;
        const doneText = cellText(readCell(row, col.done));
        const hiddenText = cellText(readCell(row, col.hidden));
        const labels = itemCols.map(c => cellText(row[c]));
        // 行メモは改行を含み得るので trim しない
        const memoValue = readCell(row, col.memo);
        const memoText = memoValue == null ? "" : String(memoValue);
        const startsTask = doneText !== "" || hiddenText !== "" || memoText.trim() !== "" || labels.some(v => v !== "");

        if (startsTask || !currentTask) {
            if (tasks.length >= MAX_PLAN_TASKS) {
                warnings.push(`行数が上限(${MAX_PLAN_TASKS})を超えたため、${excelRowNo}行目以降は読み込みませんでした。`);
                break;
            }
            currentTask = {
                id: "task_" + Date.now() + "_" + Math.random().toString(36).slice(2) + "_" + r,
                labels,
                mainSchedules: [],
                segments: [],
                memo: memoText,
                customHeight: 0,
                isDone: doneText !== "",
                isHidden: hiddenText !== "",
                isSubCollapsed: false
            };
            tasks.push(currentTask);
            currentMain = null;
        }

        const type = cellText(readCell(row, col.type));
        if (type === "") continue;
        const comment = cellText(readCell(row, col.comment));
        const startIso = excelCellToIso(readCell(row, col.start), baseSettings);
        const endIso = excelCellToIso(readCell(row, col.end), baseSettings) || startIso;
        const progress = readCell(row, col.progress);

        if (type === EXCEL_TYPE.mainAll) {
            if (!startIso) { warnings.push(`${excelRowNo}行目: M全の開始予定が読めないため飛ばしました。`); currentMain = null; continue; }
            const sIso = startIso <= endIso ? startIso : endIso;
            const eIso = startIso <= endIso ? endIso : startIso;
            const overlaps = currentTask.mainSchedules.some(m => !(eIso < m.startDate || sIso > m.endDate));
            if (overlaps) { warnings.push(`${excelRowNo}行目: メイン計画「${comment}」が同じ行の別のメイン計画と重なるため飛ばしました。`); currentMain = null; continue; }
            currentMain = {
                id: "main_" + Date.now() + "_" + Math.random().toString(36).slice(2) + "_" + r,
                startDate: sIso, endDate: eIso,
                label: comment, startLabel: "", endLabel: "",
                progressEndDate: scaleToProgressIso(sIso, eIso, progress, baseSettings),
                milestones: []
            };
            currentTask.mainSchedules.push(currentMain);
            noteDate(sIso); noteDate(eIso);
        } else if (type === EXCEL_TYPE.mainStart || type === EXCEL_TYPE.mainEnd || type === EXCEL_TYPE.milestone) {
            if (!currentMain) {
                if (comment !== "") warnings.push(`${excelRowNo}行目: 直前にM全が無いため「${comment}」を飛ばしました。`);
                continue;
            }
            let msIso = null;
            if (type === EXCEL_TYPE.mainStart) msIso = currentMain.startDate;
            else if (type === EXCEL_TYPE.mainEnd) msIso = currentMain.endDate;
            else msIso = startIso;
            if (type !== EXCEL_TYPE.milestone && comment === "") continue;  // M始/M終 の空欄は「無し」
            if (!msIso) { warnings.push(`${excelRowNo}行目: マイルストーンの日付が読めないため飛ばしました。`); continue; }
            if (msIso < currentMain.startDate || msIso > currentMain.endDate) {
                warnings.push(`${excelRowNo}行目: マイルストーン「${comment}」がメイン計画の期間外のため飛ばしました。`);
                continue;
            }
            if (currentMain.milestones.some(ms => ms.date === msIso)) continue;
            currentMain.milestones.push({
                id: "ms_" + Date.now() + "_" + Math.random().toString(36).slice(2) + "_" + r,
                date: msIso,
                label: comment || "マイルストーン"
            });
        } else if (type === EXCEL_TYPE.sub) {
            if (!startIso) { warnings.push(`${excelRowNo}行目: Sの開始予定が読めないため飛ばしました。`); continue; }
            const sIso = startIso <= endIso ? startIso : endIso;
            const eIso = startIso <= endIso ? endIso : startIso;
            if (currentTask.segments.length >= MAX_TASK_SEGMENTS) { warnings.push(`${excelRowNo}行目: サブ計画が上限(${MAX_TASK_SEGMENTS})を超えたため飛ばしました。`); continue; }
            currentTask.segments.push({
                id: "seg_" + Date.now() + "_" + Math.random().toString(36).slice(2) + "_" + r,
                startDate: sIso, endDate: eIso, type: "range",
                label: comment,
                progressEndDate: scaleToProgressIso(sIso, eIso, progress, baseSettings),
                dailyValues: {}, dailyResults: {}
            });
            noteDate(sIso); noteDate(eIso);
        } else {
            warnings.push(`${excelRowNo}行目: M/S「${type}」は不明な種別のため飛ばしました。`);
        }
    }
    tasks.forEach(t => t.mainSchedules.sort((a, b) => a.startDate.localeCompare(b.startDate)));

    // メモ・備考シート
    let memo = "";
    const memoSheet = wb.Sheets[EXCEL_SHEET_MEMO];
    if (memoSheet) {
        const memoRows = XLSX.utils.sheet_to_json(memoSheet, { header: 1, raw: true, defval: null });
        memo = memoRows.map(row => {
            const cells = (row || []).map(v => (v == null ? "" : String(v)));
            while (cells.length > 0 && cells[cells.length - 1] === "") cells.pop();
            return cells.join("\t");
        }).join("\n");
    }

    // 期間: 読込先の計画の設定を引き継ぎ、読み込んだ日付を含むように広げる
    const settings = { ...baseSettings, holidays: [...(baseSettings.holidays || [])], vacations: [...(baseSettings.vacations || [])] };
    if (minIso && minIso < settings.startDate) {
        const d = isoToDate(minIso);
        settings.startDate = dateToISO(new Date(d.getFullYear(), d.getMonth(), 1));
    }
    if (maxIso && maxIso > settings.endDate) {
        const d = isoToDate(maxIso);
        settings.endDate = dateToISO(new Date(d.getFullYear(), d.getMonth() + 1, 0));
    }

    const baseHeaders = normalizeHeaders(basePlan.headers);
    const keepWidths = baseHeaders.length === headers.length;
    const planName = planNameFromExcelFileName(fileName) || basePlan.projectName || "標準の計画";

    return {
        data: {
            projectName: planName,
            memoFormat: "plain",
            settings,
            headers,
            todoColumns: getDefaultTodoColumns(headers),
            columnWidths: keepWidths
                ? normalizeColumnWidths(basePlan.columnWidths, headers.length)
                : normalizeColumnWidths([], headers.length),
            dependencies: [],
            tasks,
            memo,
            memoHeight: basePlan.memoHeight || DEFAULT_FREE_MEMO_HEIGHT
        },
        warnings
    };
}

function importExcelFile(file) {
    if (typeof XLSX === "undefined") {
        alert("Excel読込ライブラリ (vendor/xlsx.full.min.js) が読み込まれていません。");
        return;
    }
    if (file.size > MAX_IMPORT_FILE_BYTES) {
        alert(`ファイルサイズが大きすぎます。${Math.floor(MAX_IMPORT_FILE_BYTES / (1024 * 1024))}MB 以下の Excel を読み込んでください。`);
        return;
    }
    const reader = new FileReader();
    reader.onload = async (evt) => {
        // 読込先の計画はファイル名（日程表：B.xlsx -> B）で決める
        syncDataModel();
        updateCurrentPlanEntryFromAppData();
        const planName = planNameFromExcelFileName(file.name);
        if (!planName) {
            alert("ファイル名から日程表の名前を判断できません。\n「日程表：<名前>.xlsx」の形式で保存してください。");
            return;
        }
        const existingEntry = appStore.plans.find(plan => plan.name === planName) || null;
        const basePlan = existingEntry ? existingEntry.data : appData;

        let parsed;
        try {
            const wb = XLSX.read(new Uint8Array(evt.target.result), { type: "array", cellDates: false });
            parsed = parseExcelWorkbook(wb, file.name, basePlan);
        } catch (err) {
            alert("Excelの読込に失敗しました。\n" + (err && err.message ? err.message : ""));
            return;
        }
        const validated = sanitizeImportedPlanData(parsed.data);
        validated.projectName = planName;
        const taskCount = validated.tasks.length;

        let targetEntry;
        if (existingEntry) {
            if (!confirm(`日程表「${planName}」は既にあります。\n${taskCount} 行を読み込んで上書きしてよろしいですか？`)) return;
            existingEntry.name = planName;
            existingEntry.data = clonePlanData(validated);
            targetEntry = existingEntry;
        } else {
            if (!confirm(`日程表「${planName}」は無いため、新しく作成して ${taskCount} 行を読み込みます。\nよろしいですか？`)) return;
            targetEntry = createPlanEntry(planName, validated);
            appStore.plans.push(targetEntry);
        }

        currentPlanId = targetEntry.id;
        appStore.currentPlanId = targetEntry.id;
        refreshPlanSelect();
        restoreFromData(clonePlanData(targetEntry.data));
        activeTaskId = null;
        activeProgressSegmentId = null;
        HistoryManager.init(appData);
        await persistStore();
        requestAnimationFrame(() => requestAnimationFrame(scrollToToday));

        let message = existingEntry
            ? `完了: 日程表「${planName}」を上書きしました（${taskCount} 行）。`
            : `完了: 日程表「${planName}」を新しく作成しました（${taskCount} 行）。`;
        if (parsed.warnings.length > 0) {
            const shown = parsed.warnings.slice(0, 10);
            message += `\n\n注意 (${parsed.warnings.length}件):\n・` + shown.join("\n・");
            if (parsed.warnings.length > shown.length) message += `\n・…ほか ${parsed.warnings.length - shown.length} 件`;
        }
        alert(message);
    };
    reader.onerror = () => alert("ファイルを読み込めませんでした。");
    reader.readAsArrayBuffer(file);
}

function ensureTodoDateControlLayout() {
    const todoPanel = document.getElementById("todoPanel");
    const prevBtn = document.getElementById("todoPrevDay");
    const nextBtn = document.getElementById("todoNextDay");
    const todayBtn = document.getElementById("todoTodayBtn");
    let label = document.getElementById("todoDateLabel")
        || todoPanel?.querySelector(".todo-date-label, .todo-date, [data-todo-date]");
    const controlsRow = prevBtn?.parentElement;
    if (!controlsRow || !prevBtn || !nextBtn || !todayBtn) {
        return label || null;
    }

    let group = document.getElementById("todoDateControlGroup");
    if (!group) {
        group = document.createElement("div");
        group.id = "todoDateControlGroup";
        controlsRow.innerHTML = "";
        controlsRow.appendChild(group);
    }

    controlsRow.style.display = "flex";
    controlsRow.style.justifyContent = "center";
    controlsRow.style.alignItems = "center";
    controlsRow.style.width = "100%";
    controlsRow.style.minWidth = "0";
    controlsRow.style.padding = "0 8px";
    controlsRow.style.boxSizing = "border-box";

    group.style.display = "flex";
    group.style.alignItems = "center";
    group.style.justifyContent = "center";
    group.style.gap = "10px";
    group.style.width = "fit-content";
    group.style.maxWidth = "100%";
    group.style.flex = "0 1 auto";
    group.style.minWidth = "0";

    if (!label) {
        label = document.createElement("div");
        label.id = "todoDateLabel";
        label.className = "todo-date-label";
    }

    [prevBtn, label, nextBtn, todayBtn].forEach(el => {
        if (el.parentElement !== group) {
            group.appendChild(el);
        }
    });

    prevBtn.style.flex = "0 0 auto";
    nextBtn.style.flex = "0 0 auto";
    todayBtn.style.flex = "0 0 auto";
    prevBtn.style.margin = "0";
    nextBtn.style.margin = "0";
    todayBtn.style.margin = "0";

    label.style.margin = "0";
    label.style.fontWeight = "600";
    label.style.fontSize = "clamp(15px, 2.6vw, 18px)";
    label.style.lineHeight = "1.2";
    label.style.textAlign = "center";
    label.style.whiteSpace = "nowrap";
    label.style.flex = "0 0 auto";
    label.style.minWidth = "210px";

    return label;
}

function updateTodoTable(dateObj) {
    const todoPanel = document.getElementById("todoPanel");
    let label = ensureTodoDateControlLayout()
        || document.getElementById("todoDateLabel")
        || todoPanel?.querySelector(".todo-date-label, .todo-date, [data-todo-date]");
    const table = document.getElementById("todoTable")
        || todoPanel?.querySelector("table");
    const emptyMsg = document.getElementById("todoEmptyMsg")
        || todoPanel?.querySelector(".todo-empty, .empty-message, [data-todo-empty]");
    if (!table) return;

    let thead = table.querySelector("thead");
    if (!thead) {
        thead = document.createElement("thead");
        table.prepend(thead);
    }

    let tbody = document.getElementById("todoTableBody") || table.querySelector("tbody");
    if (!tbody) {
        tbody = document.createElement("tbody");
        table.appendChild(tbody);
    }
    if (!tbody.id) tbody.id = "todoTableBody";

    const dateText = `${dateObj.getFullYear()}年${dateObj.getMonth() + 1}月${dateObj.getDate()}日(${WEEKDAYS[dateObj.getDay()]})`;
    if (label) label.textContent = dateText;

    const columnsInput = document.getElementById("todoColumnsInput");
    const columnsRaw = columnsInput ? columnsInput.value : "";
    let displayCols = columnsRaw.split(",").map(s => s.trim()).filter(Boolean);
    if (displayCols.length === 0) {
        displayCols = getDefaultTodoColumns(itemHeaders).split(",").map(s => s.trim());
    }
    const itemHeaders = getHeaderTexts();
    displayCols = normalizeTodoColumns(displayCols, itemHeaders);
    displayCols = displayCols.filter(col => col !== "メモ");
    const requiredTodoCols = [...itemHeaders, "実施内容", "計画", "実績"];
    requiredTodoCols.forEach(col => {
        if (!displayCols.includes(col)) displayCols.push(col);
    });

    const colKeys = displayCols.map(col => todoColumnKey(col, itemHeaders));
    const colWidths = {
        desc: "120px",
        plan: "70px",
        actual: "70px",
        memo: "100px",
        unknown: "100px"
    };
    const itemColWidth = (idx) => (idx === 0 ? "90px" : (idx === 1 ? "150px" : "90px"));
    const widthOf = (key) => key.startsWith("item:")
        ? itemColWidth(Number(key.slice(5)))
        : (colWidths[key] || colWidths.unknown);

    table.style.tableLayout = "fixed";
    table.style.width = "100%";
    thead.innerHTML = "";
    const trH = document.createElement("tr");
    displayCols.forEach((colName, idx) => {
        const th = document.createElement("th");
        th.textContent = colName;
        th.style.width = widthOf(colKeys[idx]);
        trH.appendChild(th);
    });
    thead.appendChild(trH);

    const iso = dateToISO(dateObj);
    tbody.innerHTML = "";

    let hasItem = false;
    let totalPlan = 0;
    let totalActual = 0;

    taskObjects.forEach(task => {
        if (task.isHidden) return;

        const itemEditables = getTaskLabelEditables(task);
        const activeSegments = task.segments.filter(seg => seg.startDate <= iso && seg.endDate >= iso);
        const rowDefs = activeSegments.length
            ? activeSegments.map(seg => ({ seg, hasTask: true }))
            : [{ seg: null, hasTask: false }];

        rowDefs.forEach(({ seg, hasTask }) => {
            const tr = document.createElement("tr");
            tr.dataset.taskId = task.id;
            if (seg) tr.dataset.segId = seg.id;

            const planValue = seg ? ((seg.dailyValues || {})[iso] || "") : "";
            const actualValue = seg ? ((seg.dailyResults || {})[iso] || "") : "";
            if (planValue && !isNaN(parseFloat(planValue))) totalPlan += parseFloat(planValue);
            if (actualValue && !isNaN(parseFloat(actualValue))) totalActual += parseFloat(actualValue);

            colKeys.forEach((key) => {
                const td = document.createElement("td");
                td.style.width = widthOf(key);

                const input = document.createElement("input");
                let val = "";
                let isEditable = false;
                input.style.width = "100%";
                input.style.boxSizing = "border-box";

                if (key.startsWith("item:")) {
                    const itemIndex = Number(key.slice(5));
                    const editable = itemEditables[itemIndex];
                    // 1列目は予定が無い行でも表示・編集できる（従来どおり）
                    const available = editable && (itemIndex === 0 || hasTask);
                    val = available ? editable.textContent : "";
                    if (itemIndex >= 2) input.style.textAlign = "center";
                    if (available) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            editable.textContent = e.target.value;
                            triggerSave();
                        });
                    }
                } else if (key === "desc") {
                    val = hasTask && seg ? (seg.label || "") : "";
                    if (hasTask && seg) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            seg.label = e.target.value;
                            renderAllSegments();
                            triggerSave();
                        });
                    }
                } else if (key === "plan") {
                    val = planValue;
                    input.style.textAlign = "center";
                    if (hasTask && seg) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            if (!seg.dailyValues) seg.dailyValues = {};
                            seg.dailyValues[iso] = e.target.value;
                            if (!e.target.value) delete seg.dailyValues[iso];
                            renderAllSegments();
                            triggerSave();
                            updateTodoTable(dateObj);
                        });
                    }
                } else if (key === "actual") {
                    val = actualValue;
                    input.style.textAlign = "center";
                    if (hasTask && seg) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            if (!seg.dailyResults) seg.dailyResults = {};
                            seg.dailyResults[iso] = e.target.value;
                            if (!e.target.value) delete seg.dailyResults[iso];
                            triggerSave();
                            updateTodoTable(dateObj);
                        });
                    }
                } else if (key === "memo") {
                    val = hasTask ? (task.memo || "") : "";
                    if (hasTask) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            task.memo = e.target.value;
                            triggerSave();
                        });
                    }
                }

                input.value = val;
                if (!isEditable) {
                    input.readOnly = true;
                    input.tabIndex = -1;
                }
                td.appendChild(input);
                tr.appendChild(td);
            });

            tbody.appendChild(tr);
            hasItem = true;
        });
    });

    if (emptyMsg) {
        emptyMsg.style.display = hasItem ? "none" : "block";
    }

    let tfoot = table.querySelector("tfoot");
    if (tfoot) table.removeChild(tfoot);
    tfoot = document.createElement("tfoot");
    const trF = document.createElement("tr");

    const tdLabel = document.createElement("td");
    tdLabel.colSpan = Math.max(1, colKeys.length - 2);
    tdLabel.textContent = "合計";
    tdLabel.style.textAlign = "right";
    trF.appendChild(tdLabel);

    const tdPlan = document.createElement("td");
    tdPlan.textContent = (totalPlan % 1 === 0) ? totalPlan : totalPlan.toFixed(1);
    tdPlan.style.textAlign = "center";
    trF.appendChild(tdPlan);

    const tdActual = document.createElement("td");
    tdActual.textContent = (totalActual % 1 === 0) ? totalActual : totalActual.toFixed(1);
    tdActual.style.textAlign = "center";
    trF.appendChild(tdActual);

    tfoot.appendChild(trF);
    table.appendChild(tfoot);
    checkTodoDeleteBtnVisibility();
}

function exportTodoToCSV() {
    const filename = `ToDoList_${formatTimestamp(new Date())}.csv`;
    const tbody = document.getElementById("todoTableBody");

    const columnsRaw = document.getElementById("todoColumnsInput").value || "";
    let columns = columnsRaw.split(",").map(s => s.trim()).filter(Boolean);
    if (columns.length === 0) {
        columns = getDefaultTodoColumns().split(",").map(s => s.trim());
    }
    columns = normalizeTodoColumns(columns);

    const headers = columns.map(c => '"' + c.replace(/"/g, '""') + '"');
    const iso = dateToISO(currentTodoDate);
    const rows = [];

    tbody.querySelectorAll("tr").forEach(tr => {
        const taskId = tr.dataset.taskId;
        const segId = tr.dataset.segId;
        const task = taskObjects.find(t => t.id === taskId);
        const seg = task ? task.segments.find(s => s.id === segId) : null;
        const hasTask = !!seg;

        const itemValues = task ? getTaskLabels(task) : [];
        const desc = hasTask && seg ? (seg.label || "") : "";
        const plan = hasTask && seg && seg.dailyValues ? (seg.dailyValues[iso] || "") : "";
        const actual = hasTask && seg && seg.dailyResults ? (seg.dailyResults[iso] || "") : "";
        const memo = hasTask && task ? (task.memo || "") : "";

        const rowData = columns.map(col => {
            const key = todoColumnKey(col);
            let v = "";
            if (key.startsWith("item:")) {
                const i = Number(key.slice(5));
                v = (i === 0 || hasTask) ? (itemValues[i] || "") : "";
            }
            else if (key === "desc") v = desc;
            else if (key === "plan") v = plan;
            else if (key === "actual") v = actual;
            else if (key === "memo") v = memo;
            return '"' + String(v).replace(/"/g, '""') + '"';
        });
        rows.push(rowData.join(","));
    });

    if (rows.length === 0) {
        alert("出力するデータがありません");
        return;
    }

    const csvContent = headers.join(",") + "\r\n" + rows.join("\r\n");
    downloadAsShiftJIS(csvContent, filename);
}

function addTodoRow(dateObj) {
    addTaskRow();
    const newTask = taskObjects[taskObjects.length - 1];
    const iso = dateToISO(dateObj);
    const newSeg = {
        id: "seg_" + Date.now() + "_" + Math.random().toString(36).slice(2),
        startDate: iso,
        endDate: iso, 
        type: "point",
        label: "", 
        progressEndDate: null,
        dailyValues: {},
        dailyResults: {}
    };
    newTask.segments.push(newSeg);
    renderAllSegments();
    triggerSave();
    updateTodoTable(dateObj);
}

function initTodoFeature() {
    const todoPanel = document.getElementById("todoPanel");
    if (!todoPanel) return;
    ensureTodoDateControlLayout();
    
    const update = () => updateTodoTable(currentTodoDate);

    document.getElementById("todoCloseBtn").addEventListener("click", () => todoPanel.classList.add("settings-hidden"));
    
    document.getElementById("todoPrevDay").addEventListener("click", () => { 
        currentTodoDate.setDate(currentTodoDate.getDate() - 1); 
        todoSelectionState = false; 
        update(); 
    });
    document.getElementById("todoNextDay").addEventListener("click", () => { 
        currentTodoDate.setDate(currentTodoDate.getDate() + 1); 
        todoSelectionState = false;
        update(); 
    });
    document.getElementById("todoTodayBtn").addEventListener("click", () => { 
        currentTodoDate = new Date(); 
        todoSelectionState = false;
        update(); 
    });

    document.getElementById("todoAddRowBtn").addEventListener("click", () => {
        addTodoRow(currentTodoDate);
    });
    
    const footerControls = document.querySelector(".todo-footer > div:nth-child(2)");
    if (!document.getElementById("todoDeleteBtn")) {
        const delBtn = document.createElement("button");
        delBtn.id = "todoDeleteBtn";
        delBtn.className = "btn-secondary";
        delBtn.style.color = "#ef4444";
        delBtn.style.borderColor = "#fca5a5";
        delBtn.style.fontSize = "12px";
        delBtn.style.marginLeft = "8px";
        delBtn.textContent = "🗑️ 選択行を削除";
        delBtn.style.display = "none"; 
        
        delBtn.addEventListener("click", () => {
            const checkboxes = document.querySelectorAll(".todo-row-checkbox[data-checked='true']");
            if (checkboxes.length === 0) {
                alert("削除する項目を選択してください。");
                return;
            }
            
            if (confirm(`${checkboxes.length} 件の項目を削除しますか？`)) {
                const itemsToDelete = [];
                checkboxes.forEach(cb => {
                    const tr = cb.closest("tr");
                    itemsToDelete.push({ taskId: tr.dataset.taskId, segId: tr.dataset.segId });
                });

                let changeOccurred = false;
                itemsToDelete.forEach(item => {
                    const task = taskObjects.find(t => t.id === item.taskId);
                    if (task) {
                        const originalLen = task.segments.length;
                        task.segments = task.segments.filter(s => s.id !== item.segId);
                        if (task.segments.length !== originalLen) changeOccurred = true;
                    }
                });

                if (changeOccurred) {
                    renderAllSegments();
                    triggerSave();
                    update(); 
                }
            }
        });
        
        const addBtn = document.getElementById("todoAddRowBtn");
        if(addBtn) {
            addBtn.insertAdjacentElement('afterend', delBtn);
        }
    }

    const win = todoPanel.querySelector(".todo-window"), header = todoPanel.querySelector(".todo-header");
    if (win) {
        win.style.minWidth = "540px";
    }
    let isDragging = false, startX, startY, initL, initT;
    header.addEventListener("mousedown", (e) => {
        if(e.target.closest("button")) return;
        isDragging = true; startX = e.clientX; startY = e.clientY;
        const r = win.getBoundingClientRect(); initL = r.left; initT = r.top;
        header.style.cursor = "grabbing"; document.body.style.userSelect = "none";
    });
    document.addEventListener("mousemove", (e) => {
        if (!isDragging) return;
        win.style.left = (initL + e.clientX - startX) + "px"; 
        win.style.top = (initT + e.clientY - startY) + "px";
        
        win.style.width = win.offsetWidth + "px";
        win.style.height = win.offsetHeight + "px";
    });
    document.addEventListener("mouseup", () => { isDragging = false; header.style.cursor = "grab"; document.body.style.userSelect = ""; });
}

window.addEventListener("resize", renderAllSegments);
initializeApp();
