const APP_VERSION = "stage3-2026-04-29";
const STORAGE_KEY = "exam-workshop-stage3-v1";
const BANK_STORAGE_KEY = "exam-workshop-bank-v1";
const LEGACY_STORAGE_KEYS = ["exam-workshop-stage1-v1", "exam-workshop-studio-v2"];

const questionList = document.getElementById("question-list");
const bankList = document.getElementById("bank-list");
const patternList = document.getElementById("patternList");
const totalQuestions = document.getElementById("totalQuestions");
const totalPoints = document.getElementById("totalPoints");
const maxScoreDisplay = document.getElementById("maxScoreDisplay");
const pointCheckStatus = document.getElementById("pointCheckStatus");
const pointCheckMessage = document.getElementById("pointCheckMessage");
const validationList = document.getElementById("validationList");
const validationCount = document.getElementById("validationCount");
const paperSheet = document.getElementById("paper-sheet");
const answerSheet = document.getElementById("answer-sheet");
const answerOnlySheet = document.getElementById("answer-only-sheet");
const analysisSheet = document.getElementById("analysis-sheet");
const previewPaperLabel = document.getElementById("previewPaperLabel");
const previewScope = document.getElementById("previewScope");
const printPreviewButton = document.getElementById("printPreview");
const saveState = document.getElementById("saveState");
const editorStatus = document.getElementById("editorStatus");
const bankStatus = document.getElementById("bankStatus");
const patternStatus = document.getElementById("patternStatus");
const templateSummary = document.getElementById("templateSummary");
const paperBaseSizeDisplay = document.getElementById("paperBaseSizeDisplay");
const paperTotalSizeDisplay = document.getElementById("paperTotalSizeDisplay");
const layoutSummary = document.getElementById("layoutSummary");
const previewVariantButtons = Array.from(document.querySelectorAll(".variant-button"));

const textMetaFieldIds = [
  "templateId",
  "schoolName",
  "examTitle",
  "subject",
  "grade",
  "unit",
  "className",
  "durationMinutes",
  "maxScore",
  "examDate",
  "paperSize",
  "customPaperWidthMm",
  "customPaperHeightMm",
  "paperTilesX",
  "paperTilesY",
  "layoutMode",
  "globalFontSize",
  "questionGap",
  "lineHeight",
  "sideMarginMm",
  "topBottomMarginMm",
  "nameLabel",
  "groupLabel",
  "numberLabel",
  "instructions"
];

const checkboxMetaFieldIds = [
  "showNameField",
  "showGroupNumberFields"
];

const bankFilterFieldIds = [
  "bankTagFilter",
  "bankDifficultyFilter",
  "bankFavoritesOnly"
];

const questionTypes = {
  multiple: {
    label: "選択問題",
    placeholder: "例: 江戸幕府を開いた人物を選びなさい。",
    usesChoices: true,
    choiceLabel: "選択肢",
    choicePlaceholder: "1行に1つずつ入力",
    choiceHelp: "選択問題は選択肢を2つ以上入れてください。正解欄には正しい選択肢そのものか、1・2・A などを入れられます。",
    sampleChoices: "徳川家康\n織田信長\n豊臣秀吉\n足利尊氏",
    answerPlaceholder: "例: 徳川家康"
  },
  truefalse: {
    label: "○×問題",
    placeholder: "例: 1603年に江戸幕府が開かれた。",
    usesChoices: false,
    choiceLabel: "",
    choicePlaceholder: "",
    choiceHelp: "",
    sampleChoices: "",
    answerPlaceholder: "例: ○"
  },
  short: {
    label: "記述問題",
    placeholder: "例: 参勤交代の目的を40字程度で説明しなさい。",
    usesChoices: false,
    choiceLabel: "",
    choicePlaceholder: "",
    choiceHelp: "",
    sampleChoices: "",
    answerPlaceholder: "模範解答を入力"
  },
  fill: {
    label: "穴埋め問題",
    placeholder: "例: 江戸幕府の政治の中心は（　　　）である。",
    usesChoices: false,
    choiceLabel: "",
    choicePlaceholder: "",
    choiceHelp: "",
    sampleChoices: "",
    answerPlaceholder: "例: 江戸"
  },
  order: {
    label: "並び替え問題",
    placeholder: "例: 次の出来事を古い順に並べなさい。",
    usesChoices: true,
    choiceLabel: "並び替え用の項目",
    choicePlaceholder: "1行に1つずつ入力",
    choiceHelp: "並び替えの項目は2つ以上入れてください。正解欄には正しい順番の項目名か、1→2→3 のような番号も使えます。",
    sampleChoices: "鎖国\n参勤交代\n大政奉還",
    answerPlaceholder: "例: 鎖国 → 参勤交代 → 大政奉還"
  }
};

const difficultyLabels = {
  easy: "やさしめ",
  standard: "ふつう",
  hard: "ややむずかしめ"
};

const answerSpaceOptions = {
  small: { label: "小さめ", lines: 1 },
  medium: { label: "ふつう", lines: 3 },
  large: { label: "大きめ", lines: 5 },
  xlarge: { label: "かなり大きめ", lines: 8 }
};

const paperSizes = {
  A4: { label: "A4", widthMm: 210, heightMm: 297, previewWidthPx: 710 },
  A3: { label: "A3", widthMm: 297, heightMm: 420, previewWidthPx: 860 },
  A5: { label: "A5", widthMm: 148, heightMm: 210, previewWidthPx: 560 },
  B4: { label: "B4", widthMm: 257, heightMm: 364, previewWidthPx: 820 },
  B5: { label: "B5", widthMm: 182, heightMm: 257, previewWidthPx: 620 },
  Letter: { label: "Letter", widthMm: 216, heightMm: 279, previewWidthPx: 730 },
  Legal: { label: "Legal", widthMm: 216, heightMm: 356, previewWidthPx: 730 },
  Postcard: { label: "はがき", widthMm: 100, heightMm: 148, previewWidthPx: 420 },
  Custom: { label: "自由サイズ", widthMm: 210, heightMm: 297, previewWidthPx: 710 }
};

const testTemplates = {
  quiz: {
    label: "小テスト",
    maxScore: "20",
    durationMinutes: "10",
    paperSize: "A4",
    layoutMode: "compact",
    showNameField: true,
    showGroupNumberFields: true,
    recommendedQuestionCount: 5,
    globalFontSize: "15",
    questionGap: "5",
    lineHeight: "1.7",
    sideMarginMm: "12",
    topBottomMarginMm: "12"
  },
  unit: {
    label: "単元テスト",
    maxScore: "50",
    durationMinutes: "30",
    paperSize: "A4",
    layoutMode: "standard",
    showNameField: true,
    showGroupNumberFields: true,
    recommendedQuestionCount: 8,
    globalFontSize: "16",
    questionGap: "7",
    lineHeight: "1.8",
    sideMarginMm: "14",
    topBottomMarginMm: "14"
  },
  term: {
    label: "定期テスト",
    maxScore: "100",
    durationMinutes: "50",
    paperSize: "A4",
    layoutMode: "spacious",
    showNameField: true,
    showGroupNumberFields: true,
    recommendedQuestionCount: 10,
    globalFontSize: "16",
    questionGap: "8",
    lineHeight: "1.9",
    sideMarginMm: "14",
    topBottomMarginMm: "15"
  },
  retry: {
    label: "再テスト",
    maxScore: "100",
    durationMinutes: "40",
    paperSize: "A4",
    layoutMode: "standard",
    showNameField: true,
    showGroupNumberFields: true,
    recommendedQuestionCount: 8,
    globalFontSize: "16",
    questionGap: "7",
    lineHeight: "1.8",
    sideMarginMm: "14",
    topBottomMarginMm: "14"
  },
  homework: {
    label: "宿題プリント",
    maxScore: "50",
    durationMinutes: "20",
    paperSize: "A4",
    layoutMode: "compact",
    showNameField: true,
    showGroupNumberFields: true,
    recommendedQuestionCount: 6,
    globalFontSize: "15",
    questionGap: "5",
    lineHeight: "1.7",
    sideMarginMm: "12",
    topBottomMarginMm: "12"
  },
  checksheet: {
    label: "確認プリント",
    maxScore: "30",
    durationMinutes: "15",
    paperSize: "A4",
    layoutMode: "compact",
    showNameField: true,
    showGroupNumberFields: true,
    recommendedQuestionCount: 6,
    globalFontSize: "15",
    questionGap: "5",
    lineHeight: "1.7",
    sideMarginMm: "12",
    topBottomMarginMm: "12"
  }
};

const defaultEditorMessage = "問題は入力するたびに自動保存されます。";
const defaultInstructions = "ていねいな字で書き、見直しをしてから提出しましょう。";

const uiState = {
  activeTab: "paper",
  previewVariant: "base",
  editorMessage: defaultEditorMessage,
  saveFailedWarned: false,
  bankFailedWarned: false
};

const state = {
  meta: createDefaultMeta(),
  questions: [],
  bank: [],
  bankFilters: createDefaultBankFilters(),
  patterns: createEmptyPatterns()
};

function createDefaultMeta() {
  const template = testTemplates.unit;
  return {
    templateId: "unit",
    schoolName: "南中学校",
    examTitle: "歴史の確認テスト",
    subject: "社会",
    grade: "2年",
    unit: "江戸時代",
    className: "A組",
    durationMinutes: template.durationMinutes,
    maxScore: template.maxScore,
    examDate: "",
    paperSize: template.paperSize,
    customPaperWidthMm: "210",
    customPaperHeightMm: "297",
    paperTilesX: "1",
    paperTilesY: "1",
    globalFontSize: template.globalFontSize,
    questionGap: template.questionGap,
    lineHeight: template.lineHeight,
    sideMarginMm: template.sideMarginMm,
    topBottomMarginMm: template.topBottomMarginMm,
    nameLabel: "名前",
    groupLabel: "組",
    numberLabel: "番号",
    instructions: defaultInstructions,
    showNameField: template.showNameField,
    showGroupNumberFields: template.showGroupNumberFields,
    layoutMode: template.layoutMode
  };
}

function createDefaultBankFilters() {
  return {
    tag: "",
    difficulty: "all",
    favoritesOnly: false
  };
}

function createEmptyPatterns() {
  return {
    shuffleQuestions: true,
    shuffleChoices: true,
    generated: {}
  };
}

function createId() {
  if (window.crypto && typeof window.crypto.randomUUID === "function") {
    return window.crypto.randomUUID();
  }
  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createSeededRandom(seedText) {
  let seed = 2166136261;
  const text = String(seedText || "seed");
  for (let index = 0; index < text.length; index += 1) {
    seed ^= text.charCodeAt(index);
    seed = Math.imul(seed, 16777619);
  }
  let stateValue = seed >>> 0;
  return () => {
    stateValue ^= stateValue << 13;
    stateValue ^= stateValue >>> 17;
    stateValue ^= stateValue << 5;
    return (stateValue >>> 0) / 4294967296;
  };
}

function shuffleArray(items, random) {
  const next = [...items];
  for (let index = next.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(random() * (index + 1));
    [next[index], next[swapIndex]] = [next[swapIndex], next[index]];
  }
  return next;
}

function getDefaultAnswerSpace(type) {
  return type === "short" || type === "order" ? "medium" : "small";
}

function getDefaultAnswerAreaHeightMm(type) {
  if (type === "multiple" || type === "truefalse") {
    return "12";
  }
  if (type === "fill") {
    return "18";
  }
  if (type === "order") {
    return "28";
  }
  return "36";
}

function convertAnswerLinesToSpace(value, type) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 1) {
    return getDefaultAnswerSpace(type);
  }
  if (number <= 3) {
    return "medium";
  }
  if (number <= 5) {
    return "large";
  }
  return "xlarge";
}

function normalizeQuestion(raw = {}) {
  const type = questionTypes[raw.type] ? raw.type : "multiple";
  const choices = typeof raw.choices === "string"
    ? raw.choices
    : Array.isArray(raw.choices)
      ? raw.choices.join("\n")
      : "";

  return {
    id: raw.id || createId(),
    type,
    prompt: String(raw.prompt ?? ""),
    points: raw.points === undefined || raw.points === null ? "10" : String(raw.points),
    answer: String(raw.answer ?? ""),
    explanation: String(raw.explanation ?? raw.note ?? ""),
    difficulty: Object.prototype.hasOwnProperty.call(difficultyLabels, raw.difficulty) ? raw.difficulty : "standard",
    unitTag: String(raw.unitTag ?? raw.tag ?? raw.section ?? ""),
    answerSpace: Object.prototype.hasOwnProperty.call(answerSpaceOptions, raw.answerSpace)
      ? raw.answerSpace
      : convertAnswerLinesToSpace(raw.answerLines, type),
    questionFontSize: String(raw.questionFontSize ?? raw.fontSize ?? "16"),
    choiceFontSize: String(raw.choiceFontSize ?? "15"),
    answerAreaHeightMm: String(raw.answerAreaHeightMm ?? getDefaultAnswerAreaHeightMm(type)),
    blockPaddingMm: String(raw.blockPaddingMm ?? "4"),
    blockWidthPercent: String(raw.blockWidthPercent ?? "100"),
    figureSpaceHeightMm: String(raw.figureSpaceHeightMm ?? "0"),
    choices: questionTypes[type].usesChoices ? choices : ""
  };
}

function createEmptyQuestion(type = "multiple", overrides = {}) {
  return normalizeQuestion({
    type,
    prompt: overrides.prompt ?? "",
    points: overrides.points ?? "10",
    answer: overrides.answer ?? "",
    explanation: overrides.explanation ?? "",
    difficulty: overrides.difficulty ?? "standard",
    unitTag: overrides.unitTag ?? state.meta.unit ?? createDefaultMeta().unit,
    answerSpace: overrides.answerSpace ?? getDefaultAnswerSpace(type),
    questionFontSize: overrides.questionFontSize ?? "16",
    choiceFontSize: overrides.choiceFontSize ?? "15",
    answerAreaHeightMm: overrides.answerAreaHeightMm ?? getDefaultAnswerAreaHeightMm(type),
    blockPaddingMm: overrides.blockPaddingMm ?? "4",
    blockWidthPercent: overrides.blockWidthPercent ?? "100",
    figureSpaceHeightMm: overrides.figureSpaceHeightMm ?? "0",
    choices: overrides.choices ?? questionTypes[type].sampleChoices
  });
}

function cloneQuestion(question, overrides = {}) {
  return normalizeQuestion({
    ...question,
    ...overrides
  });
}

function parsePositiveIntStrict(value) {
  const text = String(value ?? "").trim();
  if (!text || !/^\d+$/.test(text)) {
    return { valid: false, value: null };
  }
  const number = Number(text);
  if (!Number.isFinite(number) || number <= 0) {
    return { valid: false, value: null };
  }
  return { valid: true, value: number };
}

function parseDurationValue(value) {
  const match = String(value ?? "").match(/\d+/);
  return match ? match[0] : "";
}

function normalizeDateValue(value) {
  const text = String(value ?? "").trim();
  if (!text) {
    return "";
  }
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : "";
}

function inferMaxScore(questions = []) {
  const total = questions.reduce((sum, question) => {
    const parsed = parsePositiveIntStrict(question.points);
    return sum + (parsed.valid ? parsed.value : 0);
  }, 0);
  return total > 0 ? String(total) : createDefaultMeta().maxScore;
}

function getQuestionPointsValue(question) {
  const parsed = parsePositiveIntStrict(question.points);
  return parsed.valid ? parsed.value : null;
}

function getQuestionPointsLabel(question) {
  const value = getQuestionPointsValue(question);
  return value === null ? "配点未設定" : `${value}点`;
}

function getChoiceItems(question) {
  return String(question.choices || "")
    .split("\n")
    .map((choice) => choice.trim())
    .filter(Boolean);
}

function getAnswerLinesCount(question) {
  const preset = answerSpaceOptions[question.answerSpace] || answerSpaceOptions.medium;
  if (question.type === "multiple" || question.type === "truefalse") {
    return 1;
  }
  return preset.lines;
}

function normalizeSpace(text) {
  return String(text ?? "")
    .replace(/\r\n/g, "\n")
    .replace(/[ \t]+/g, " ")
    .trim();
}

function buildQuestionFingerprint(question) {
  const normalized = normalizeQuestion(question);
  return JSON.stringify({
    type: normalized.type,
    prompt: normalizeSpace(normalized.prompt),
    choices: getChoiceItems(normalized).map((item) => normalizeSpace(item)),
    answer: normalizeSpace(normalized.answer)
  });
}

function normalizeBankEntry(raw = {}) {
  const question = normalizeQuestion(raw.question ?? raw);
  return {
    id: String(raw.id ?? createId()),
    fingerprint: String(raw.fingerprint ?? buildQuestionFingerprint(question)),
    favorite: Boolean(raw.favorite),
    createdAt: String(raw.createdAt ?? new Date().toISOString()),
    updatedAt: String(raw.updatedAt ?? raw.createdAt ?? new Date().toISOString()),
    question
  };
}

function escapeHtml(text) {
  return String(text || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeAttribute(text) {
  return escapeHtml(text).replaceAll("`", "&#96;");
}

function formatMultilineHtml(text, fallback = "") {
  const safeText = String(text || "").trim() || fallback;
  return escapeHtml(safeText).replaceAll("\n", "<br>");
}

function formatDateLabel(value) {
  if (!value) {
    return "未設定";
  }
  const [year, month, day] = String(value).split("-");
  if (!year || !month || !day) {
    return value;
  }
  return `${Number(year)}年${Number(month)}月${Number(day)}日`;
}

function getPaperConfig(paperSize = state.meta.paperSize) {
  return paperSizes[paperSize] || paperSizes.A4;
}

function getMetaValue(key, fallback = "") {
  const value = state.meta[key];
  if (typeof value === "boolean") {
    return value;
  }
  const text = String(value ?? "").trim();
  return text || fallback;
}

function getTemplateConfig(templateId = state.meta.templateId) {
  return testTemplates[templateId] || testTemplates.unit;
}

function renderTemplateSummary() {
  const template = getTemplateConfig();
  templateSummary.textContent = `${template.label} / 満点 ${template.maxScore}点 / ${template.durationMinutes}分 / 問題数のめやす ${template.recommendedQuestionCount}問 / ${template.paperSize} / ${template.layoutMode === "compact" ? "詰めて印刷" : template.layoutMode === "spacious" ? "ゆったり印刷" : "標準印刷"}`;
}

function syncMetaFromInputs() {
  textMetaFieldIds.forEach((id) => {
    const element = document.getElementById(id);
    state.meta[id] = element.value.trim();
  });
  checkboxMetaFieldIds.forEach((id) => {
    const element = document.getElementById(id);
    state.meta[id] = element.checked;
  });
}

function populateFieldsFromState() {
  textMetaFieldIds.forEach((id) => {
    const element = document.getElementById(id);
    if (element) {
      element.value = state.meta[id] ?? "";
    }
  });
  checkboxMetaFieldIds.forEach((id) => {
    const element = document.getElementById(id);
    if (element) {
      element.checked = Boolean(state.meta[id]);
    }
  });
  document.getElementById("shuffleQuestionOrder").checked = state.patterns.shuffleQuestions;
  document.getElementById("shuffleChoiceOrder").checked = state.patterns.shuffleChoices;
  document.getElementById("bankTagFilter").value = state.bankFilters.tag;
  document.getElementById("bankDifficultyFilter").value = state.bankFilters.difficulty;
  document.getElementById("bankFavoritesOnly").checked = state.bankFilters.favoritesOnly;
}

function buildBackupPayload() {
  return {
    version: APP_VERSION,
    savedAt: new Date().toISOString(),
    meta: { ...state.meta },
    questions: state.questions.map((question) => ({ ...question }))
  };
}

function parseStoredJson(raw, onError) {
  try {
    return JSON.parse(raw);
  } catch (error) {
    console.warn(error);
    if (typeof onError === "function") {
      onError(error);
    }
    return null;
  }
}

function showStorageWarning(label) {
  window.alert(`${label}の保存に失敗しました。ブラウザの保存領域がいっぱいの可能性があります。不要なバックアップやタブを整理してからもう一度お試しください。`);
}

function updateSaveState(success, message) {
  saveState.textContent = message;
  if (success) {
    uiState.saveFailedWarned = false;
  }
}

function persistDraft() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(buildBackupPayload()));
    const timeLabel = new Date().toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit" });
    updateSaveState(true, `${timeLabel} に自動保存しました。`);
    return true;
  } catch (error) {
    console.error(error);
    updateSaveState(false, "自動保存に失敗しました。");
    if (!uiState.saveFailedWarned) {
      uiState.saveFailedWarned = true;
      showStorageWarning("下書き");
    }
    return false;
  }
}

function persistBank() {
  try {
    const payload = {
      version: APP_VERSION,
      savedAt: new Date().toISOString(),
      items: state.bank
    };
    localStorage.setItem(BANK_STORAGE_KEY, JSON.stringify(payload));
    uiState.bankFailedWarned = false;
    return true;
  } catch (error) {
    console.error(error);
    if (!uiState.bankFailedWarned) {
      uiState.bankFailedWarned = true;
      showStorageWarning("問題バンク");
    }
    return false;
  }
}

function loadQuestionBank() {
  const raw = localStorage.getItem(BANK_STORAGE_KEY);
  if (!raw) {
    state.bank = [];
    return;
  }

  const data = parseStoredJson(raw, () => {
    localStorage.removeItem(BANK_STORAGE_KEY);
    window.alert("問題バンクの保存データが読み込めなかったため、問題バンクを初期化しました。");
  });

  if (!data) {
    state.bank = [];
    return;
  }

  const items = Array.isArray(data.items)
    ? data.items
    : Array.isArray(data.bank)
      ? data.bank
      : Array.isArray(data)
        ? data
        : [];

  state.bank = items.map(normalizeBankEntry);
}

function applyTemplateConfig(templateId, message = "") {
  const nextTemplateId = testTemplates[templateId] ? templateId : "unit";
  const template = getTemplateConfig(nextTemplateId);
  state.meta.templateId = nextTemplateId;
  state.meta.maxScore = template.maxScore;
  state.meta.durationMinutes = template.durationMinutes;
  state.meta.paperSize = template.paperSize;
  state.meta.layoutMode = template.layoutMode;
  state.meta.showNameField = template.showNameField;
  state.meta.showGroupNumberFields = template.showGroupNumberFields;
  populateFieldsFromState();
  renderTemplateSummary();
  if (message) {
    setEditorMessage(message);
  }
}

function createVariantBundle(variantKey) {
  if (variantKey === "base") {
    return {
      key: "base",
      label: "元のテスト",
      questions: state.questions.map((question) => cloneQuestion(question)),
      isPattern: false
    };
  }

  const pattern = state.patterns.generated[variantKey];
  if (!pattern) {
    return createVariantBundle("base");
  }

  return {
    key: variantKey,
    label: `${variantKey}パターン`,
    questions: pattern.questions.map((question) => cloneQuestion(question)),
    isPattern: true
  };
}

function getActiveBundle() {
  return createVariantBundle(uiState.previewVariant);
}

function setEditorMessage(message) {
  uiState.editorMessage = message;
  editorStatus.textContent = message;
}

function setPatternMessage(message) {
  patternStatus.textContent = message;
}

function resetPatterns(message = "A/B/Cパターンを消しました。") {
  const hadPatterns = Object.keys(state.patterns.generated).length > 0;
  state.patterns.generated = {};
  if (uiState.previewVariant !== "base") {
    uiState.previewVariant = "base";
  }
  if (hadPatterns) {
    setPatternMessage(message);
  } else {
    setPatternMessage("まだA/B/Cパターンは作っていません。");
  }
}

function invalidatePatterns(message = "問題が変わったので、A/B/Cパターンは作り直してください。") {
  if (!Object.keys(state.patterns.generated).length) {
    return;
  }
  resetPatterns(message);
}

function getValidationState() {
  const errors = [];
  const warnings = [];
  const questionIssueMap = new Map();

  const maxScoreCheck = parsePositiveIntStrict(state.meta.maxScore);
  const durationCheck = parsePositiveIntStrict(state.meta.durationMinutes);

  if (!maxScoreCheck.valid) {
    errors.push({ severity: "error", message: "満点は1以上の数字で入力してください。" });
  }

  if (!durationCheck.valid) {
    errors.push({ severity: "error", message: "制限時間は1以上の数字で入力してください。" });
  }

  const totalPointsValue = state.questions.reduce((sum, question) => {
    const points = getQuestionPointsValue(question);
    return sum + (points ?? 0);
  }, 0);

  if (!state.questions.length) {
    warnings.push({ severity: "warning", message: "問題がまだありません。右の追加ボタンから問題を入れてください。" });
  }

  state.questions.forEach((question, index) => {
    const messages = [];
    const questionNumber = index + 1;

    if (!question.prompt.trim()) {
      messages.push(`第${questionNumber}問の問題文を入力してください。`);
    }

    if (!question.answer.trim()) {
      messages.push(`第${questionNumber}問の正解を入力してください。`);
    }

    if (!parsePositiveIntStrict(question.points).valid) {
      messages.push(`第${questionNumber}問の配点は1以上の数字で入力してください。`);
    }

    if ((question.type === "multiple" || question.type === "order") && getChoiceItems(question).length < 2) {
      messages.push(`第${questionNumber}問の${questionTypes[question.type].choiceLabel}は2つ以上必要です。`);
    }

    if (messages.length) {
      questionIssueMap.set(index, messages);
      messages.forEach((message) => {
        errors.push({ severity: "error", message });
      });
    }
  });

  if (maxScoreCheck.valid && state.questions.length && totalPointsValue !== maxScoreCheck.value) {
    warnings.push({
      severity: "warning",
      message: `合計点は${totalPointsValue}点です。設定した満点の${maxScoreCheck.value}点と一致していません。`
    });
  }

  return {
    errors,
    warnings,
    questionIssueMap,
    totalPointsValue,
    maxScoreValue: maxScoreCheck.valid ? maxScoreCheck.value : null,
    durationValue: durationCheck.valid ? durationCheck.value : null
  };
}

function renderValidationItemsHtml(items, emptyLabel) {
  if (!items.length) {
    return `<div class="validation-item is-success">${escapeHtml(emptyLabel)}</div>`;
  }
  return items.map((item) => `
    <div class="validation-item ${item.severity === "error" ? "is-error" : "is-warning"}">
      ${escapeHtml(item.message)}
    </div>
  `).join("");
}

function renderStats(validation) {
  totalQuestions.textContent = String(state.questions.length);
  totalPoints.textContent = `${validation.totalPointsValue}点`;
  maxScoreDisplay.textContent = validation.maxScoreValue === null ? "未設定" : `${validation.maxScoreValue}点`;

  pointCheckStatus.classList.remove("point-check-ok", "point-check-warn");

  if (validation.errors.length) {
    pointCheckStatus.textContent = "修正が必要";
    pointCheckStatus.classList.add("point-check-warn");
    pointCheckMessage.textContent = `印刷前に直したい項目が ${validation.errors.length} 件あります。`;
  } else if (validation.warnings.length) {
    pointCheckStatus.textContent = "注意あり";
    pointCheckStatus.classList.add("point-check-warn");
    pointCheckMessage.textContent = validation.warnings[0].message;
  } else {
    pointCheckStatus.textContent = "このまま使えます";
    pointCheckStatus.classList.add("point-check-ok");
    pointCheckMessage.textContent = "配点と入力内容はそろっています。このまま印刷や保存に進めます。";
  }

  validationCount.textContent = `${validation.errors.length + validation.warnings.length}件`;
  validationList.innerHTML = renderValidationItemsHtml(
    [...validation.errors, ...validation.warnings],
    "気になる点はありません。"
  );
}

function renderTypeOptions(selectedType) {
  return Object.entries(questionTypes)
    .map(([value, config]) => `<option value="${value}"${value === selectedType ? " selected" : ""}>${config.label}</option>`)
    .join("");
}

function renderDifficultyOptions(selectedDifficulty) {
  return Object.entries(difficultyLabels)
    .map(([value, label]) => `<option value="${value}"${value === selectedDifficulty ? " selected" : ""}>${label}</option>`)
    .join("");
}

function renderAnswerSpaceOptions(selectedValue) {
  return Object.entries(answerSpaceOptions)
    .map(([value, config]) => `<option value="${value}"${value === selectedValue ? " selected" : ""}>${config.label}</option>`)
    .join("");
}

function renderQuestionCard(question, index, validation) {
  const config = questionTypes[question.type];
  const issues = validation.questionIssueMap.get(index) || [];
  const choiceHiddenClass = config.usesChoices ? "" : "is-hidden";

  return `
    <article class="question-card ${issues.length ? "is-invalid" : ""}" data-index="${index}">
      <div class="question-card-top">
        <div>
          <p class="mini-type">${escapeHtml(config.label)}</p>
          <h4>第${index + 1}問</h4>
          <div class="meta-badges">
            <span>${escapeHtml(getQuestionPointsLabel(question))}</span>
            <span>${escapeHtml(difficultyLabels[question.difficulty])}</span>
            <span>${escapeHtml(question.unitTag || "単元タグなし")}</span>
            ${issues.length ? "<span>要確認</span>" : ""}
          </div>
        </div>
        <div class="card-actions">
          <button type="button" class="small-button" data-action="move-up">上へ</button>
          <button type="button" class="small-button" data-action="move-down">下へ</button>
          <button type="button" class="small-button" data-action="duplicate">複製</button>
          <button type="button" class="small-button" data-action="save-bank">バンク保存</button>
          <button type="button" class="small-button" data-action="delete">削除</button>
        </div>
      </div>

      <div class="question-grid">
        <label>
          問題形式
          <select data-field="type">
            ${renderTypeOptions(question.type)}
          </select>
        </label>
        <label>
          配点
          <input type="number" min="1" step="1" data-field="points" value="${escapeAttribute(question.points)}" placeholder="例: 5">
        </label>
        <label>
          難しさ
          <select data-field="difficulty">
            ${renderDifficultyOptions(question.difficulty)}
          </select>
        </label>
        <label>
          単元タグ
          <input type="text" data-field="unitTag" value="${escapeAttribute(question.unitTag)}" placeholder="例: 江戸時代">
        </label>
        <label class="full-width">
          解答欄の大きさ
          <select data-field="answerSpace">
            ${renderAnswerSpaceOptions(question.answerSpace)}
          </select>
        </label>
        <label class="full-width">
          問題文
          <textarea rows="4" data-field="prompt" placeholder="${escapeAttribute(config.placeholder)}">${escapeHtml(question.prompt)}</textarea>
        </label>
        <label class="full-width choice-wrap ${choiceHiddenClass}">
          ${escapeHtml(config.choiceLabel)}
          <textarea rows="4" data-field="choices" placeholder="${escapeAttribute(config.choicePlaceholder)}">${escapeHtml(question.choices)}</textarea>
          <small class="field-tip">${escapeHtml(config.choiceHelp)}</small>
        </label>
        <label class="full-width">
          正解
          <textarea rows="2" data-field="answer" placeholder="${escapeAttribute(config.answerPlaceholder)}">${escapeHtml(question.answer)}</textarea>
        </label>
        <label class="full-width">
          解説
          <textarea rows="3" data-field="explanation" placeholder="先生用の解説や採点メモを入れられます。">${escapeHtml(question.explanation)}</textarea>
        </label>
      </div>
    </article>
  `;
}

function renderQuestionList(validation) {
  if (!state.questions.length) {
    questionList.innerHTML = `
      <div class="empty-state">
        <div>
          <h3>まだ問題がありません</h3>
          <p>追加ボタンから問題を作るか、問題バンクから再利用してください。</p>
        </div>
      </div>
    `;
    return;
  }

  questionList.innerHTML = state.questions
    .map((question, index) => renderQuestionCard(question, index, validation))
    .join("");
}

function updateBankFiltersFromInputs() {
  state.bankFilters.tag = document.getElementById("bankTagFilter").value.trim();
  state.bankFilters.difficulty = document.getElementById("bankDifficultyFilter").value;
  state.bankFilters.favoritesOnly = document.getElementById("bankFavoritesOnly").checked;
}

function getFilteredBankEntries() {
  const normalizedTag = state.bankFilters.tag.trim().toLowerCase();
  return state.bank.filter((entry) => {
    const tagMatch = !normalizedTag || entry.question.unitTag.toLowerCase().includes(normalizedTag);
    const difficultyMatch = state.bankFilters.difficulty === "all" || entry.question.difficulty === state.bankFilters.difficulty;
    const favoriteMatch = !state.bankFilters.favoritesOnly || entry.favorite;
    return tagMatch && difficultyMatch && favoriteMatch;
  });
}

function renderBankCard(entry) {
  const question = entry.question;
  const previewText = question.prompt.trim() || "問題文はまだ入っていません。";

  return `
    <article class="bank-card ${entry.favorite ? "is-favorite" : ""}" data-bank-id="${escapeAttribute(entry.id)}">
      <div class="bank-card-top">
        <div>
          <p class="mini-type">${escapeHtml(questionTypes[question.type].label)}</p>
          <h4>${escapeHtml(previewText.slice(0, 48) || "保存済み問題")}${previewText.length > 48 ? "…" : ""}</h4>
          <div class="bank-card-tags">
            <span>${escapeHtml(getQuestionPointsLabel(question))}</span>
            <span>${escapeHtml(difficultyLabels[question.difficulty])}</span>
            <span>${escapeHtml(question.unitTag || "単元タグなし")}</span>
            ${entry.favorite ? "<span>お気に入り</span>" : ""}
          </div>
        </div>
        <div class="card-actions">
          <button type="button" class="small-button" data-bank-action="reuse">テストに追加</button>
          <button type="button" class="small-button" data-bank-action="favorite">${entry.favorite ? "お気に入り解除" : "お気に入り"}</button>
          <button type="button" class="small-button" data-bank-action="delete">削除</button>
        </div>
      </div>
      <p class="bank-card-copy">${escapeHtml(previewText)}</p>
    </article>
  `;
}

function renderBankList() {
  const filteredEntries = getFilteredBankEntries();

  if (!state.bank.length) {
    bankStatus.textContent = "問題バンクはまだ空です。問題カードの「バンク保存」か「今の問題を全部バンク保存」でためていけます。";
    bankList.innerHTML = `
      <div class="empty-state">
        <div>
          <h3>問題バンクは空です</h3>
          <p>一度保存すると、次のテストでもすぐ再利用できます。</p>
        </div>
      </div>
    `;
    return;
  }

  bankStatus.textContent = `保存済み ${state.bank.length} 件 / 表示中 ${filteredEntries.length} 件`;

  if (!filteredEntries.length) {
    bankList.innerHTML = `
      <div class="empty-state">
        <div>
          <h3>条件に合う問題がありません</h3>
          <p>単元タグや難しさの絞り込みをゆるめてみてください。</p>
        </div>
      </div>
    `;
    return;
  }

  bankList.innerHTML = filteredEntries.map(renderBankCard).join("");
}

function describePatternOptions() {
  const labels = [];
  labels.push(state.patterns.shuffleQuestions ? "問題順シャッフル" : "問題順はそのまま");
  labels.push(state.patterns.shuffleChoices ? "選択肢シャッフル" : "選択肢はそのまま");
  return labels.join(" / ");
}

function renderPatternCard(key) {
  const pattern = state.patterns.generated[key];
  if (!pattern) {
    return "";
  }
  const activeClass = uiState.previewVariant === key ? " pattern-card-active" : "";
  return `
    <article class="pattern-card${activeClass}" data-pattern-key="${key}">
      <div class="pattern-card-top">
        <div>
          <p class="mini-type">Pattern</p>
          <h4>${key}パターン</h4>
          <div class="meta-badges">
            <span>${pattern.questions.length}問</span>
            <span>${pattern.totalPoints}点</span>
            <span>${escapeHtml(describePatternOptions())}</span>
          </div>
        </div>
        <div class="card-actions">
          <button type="button" class="small-button" data-pattern-action="show">中央に表示</button>
          <button type="button" class="small-button" data-pattern-action="student">生徒用</button>
          <button type="button" class="small-button" data-pattern-action="teacher">解答付き</button>
          <button type="button" class="small-button" data-pattern-action="answeronly">解答用紙</button>
          <button type="button" class="small-button" data-pattern-action="csv">CSV</button>
        </div>
      </div>
      <p class="pattern-card-copy">元の問題は残したまま、このパターンだけを印刷・CSV出力できます。</p>
    </article>
  `;
}

function syncPreviewVariantButtons() {
  previewVariantButtons.forEach((button) => {
    const variant = button.dataset.variant;
    const enabled = variant === "base" || Boolean(state.patterns.generated[variant]);
    button.disabled = !enabled;
    button.classList.toggle("active", uiState.previewVariant === variant);
  });
}

function renderPatternList() {
  const keys = ["A", "B", "C"].filter((key) => state.patterns.generated[key]);
  syncPreviewVariantButtons();

  if (!keys.length) {
    patternList.innerHTML = `
      <div class="empty-state">
        <div>
          <h3>A/B/Cパターンはまだありません</h3>
          <p>左の「A/B/Cパターンを作る」を押すと、元の問題を残したまま別パターンを作れます。</p>
        </div>
      </div>
    `;
    return;
  }

  patternList.innerHTML = keys.map(renderPatternCard).join("");
}

function updatePreviewScope() {
  const bundle = getActiveBundle();
  previewScope.textContent = `${bundle.label}を中央表示しています。CSVと印刷もこの表示内容に合わせて出せます。`;
}

function renderQuestionBadges(question) {
  const badges = [
    question.unitTag ? `<span class="paper-tag">${escapeHtml(question.unitTag)}</span>` : "",
    question.difficulty ? `<span class="paper-tag is-difficulty">${escapeHtml(difficultyLabels[question.difficulty])}</span>` : ""
  ].filter(Boolean);
  return badges.length ? `<div class="paper-badges">${badges.join("")}</div>` : "";
}

function renderChoiceList(question) {
  const items = getChoiceItems(question);
  if (!items.length) {
    return "";
  }
  if (question.type === "order") {
    return `<ul class="choice-list">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
  }
  return `<ol class="choice-list">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ol>`;
}

function renderAnswerLinesHtml(lineCount) {
  return `
    <div class="answer-lines">
      ${Array.from({ length: lineCount }, () => '<div class="answer-line"></div>').join("")}
    </div>
  `;
}

function renderStudentRowHtml() {
  const showNameField = Boolean(state.meta.showNameField);
  const showGroupNumberFields = Boolean(state.meta.showGroupNumberFields);
  const blocks = [];

  if (showGroupNumberFields) {
    blocks.push(`
      <div>
        <small>${escapeHtml(getMetaValue("groupLabel", "組"))}</small>
        <div class="line-box"></div>
      </div>
    `);
    blocks.push(`
      <div>
        <small>${escapeHtml(getMetaValue("numberLabel", "番号"))}</small>
        <div class="line-box"></div>
      </div>
    `);
  }

  if (showNameField) {
    blocks.push(`
      <div>
        <small>${escapeHtml(getMetaValue("nameLabel", "名前"))}</small>
        <div class="line-box"></div>
      </div>
    `);
  }

  if (!blocks.length) {
    return "";
  }

  const rowClass = blocks.length === 1 ? "student-row is-name-only" : "student-row";
  return `<div class="${rowClass}">${blocks.join("")}</div>`;
}

function renderCommonPaperHeader(bundle, extraTitle = "", extraNote = "") {
  const paperLabel = getPaperConfig().label;
  const duration = parsePositiveIntStrict(state.meta.durationMinutes);
  const maxScore = parsePositiveIntStrict(state.meta.maxScore);

  return `
    <header class="paper-header">
      <div class="paper-kicker">
        <span>${escapeHtml(getMetaValue("schoolName", "学校名"))}</span>
        <span>${escapeHtml(getMetaValue("subject", "教科"))}</span>
      </div>
      <h2 class="paper-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))}</h2>
      <div class="paper-kicker">
        <span>学年: ${escapeHtml(getMetaValue("grade", "未設定"))}</span>
        <span>単元: ${escapeHtml(getMetaValue("unit", "未設定"))}</span>
        <span>クラス: ${escapeHtml(getMetaValue("className", "未設定"))}</span>
      </div>
      <div class="paper-kicker">
        <span>実施日: ${escapeHtml(formatDateLabel(state.meta.examDate))}</span>
        <span>制限時間: ${escapeHtml(duration.valid ? `${duration.value}分` : "未設定")}</span>
        <span>満点: ${escapeHtml(maxScore.valid ? `${maxScore.value}点` : "未設定")}</span>
        <span>用紙: ${escapeHtml(paperLabel)}</span>
        <span>${escapeHtml(bundle.label)}</span>
      </div>
      <p class="paper-notes">${formatMultilineHtml(state.meta.instructions, defaultInstructions)}</p>
      ${extraNote ? `<p class="paper-notes">${escapeHtml(extraNote)}</p>` : ""}
      ${extraTitle ? `<p class="paper-notes">${escapeHtml(extraTitle)}</p>` : ""}
      ${renderStudentRowHtml()}
    </header>
  `;
}

function renderPaperView(bundle) {
  const questions = bundle.questions;
  if (!questions.length) {
    paperSheet.innerHTML = `
      <div class="empty-state">
        <div>
          <h3>まだ問題がありません</h3>
          <p>問題を追加すると、生徒用の見た目がここに表示されます。</p>
        </div>
      </div>
    `;
    return;
  }

  const questionBlocks = questions.map((question, index) => `
    <section class="paper-question">
      <div class="paper-question-top">
        <span>第${index + 1}問 ${escapeHtml(questionTypes[question.type].label)}</span>
        <span>${escapeHtml(getQuestionPointsLabel(question))}</span>
      </div>
      <p>${formatMultilineHtml(question.prompt, "ここに問題文を入力してください。")}</p>
      ${renderQuestionBadges(question)}
      ${renderChoiceList(question)}
      ${renderAnswerLinesHtml(getAnswerLinesCount(question))}
    </section>
  `).join("");

  paperSheet.innerHTML = `${renderCommonPaperHeader(bundle)}${questionBlocks}`;
}

function renderAnswerView(bundle) {
  const questions = bundle.questions;
  if (!questions.length) {
    answerSheet.innerHTML = `
      <div class="empty-state">
        <div>
          <h3>解答付きプレビューはまだ空です</h3>
          <p>問題と正解を入れると、先生用の確認画面になります。</p>
        </div>
      </div>
    `;
    return;
  }

  const total = questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0);
  const questionBlocks = questions.map((question, index) => `
    <section class="paper-question">
      <div class="paper-question-top">
        <span>第${index + 1}問 ${escapeHtml(questionTypes[question.type].label)}</span>
        <span>${escapeHtml(getQuestionPointsLabel(question))}</span>
      </div>
      <p>${formatMultilineHtml(question.prompt, "ここに問題文を入力してください。")}</p>
      ${renderQuestionBadges(question)}
      ${renderChoiceList(question)}
      <div class="answer-box">
        <strong>正解</strong>
        <p>${formatMultilineHtml(question.answer, "未入力")}</p>
      </div>
      <div class="note-box">
        <strong>解説</strong>
        <p>${formatMultilineHtml(question.explanation, "解説はまだありません。")}</p>
      </div>
    </section>
  `).join("");

  answerSheet.innerHTML = `${renderCommonPaperHeader(bundle, `解答付き / 合計 ${total}点`, "印刷すると先生用の答え付き資料になります。")}${questionBlocks}`;
}

function renderAnswerOnlyView(bundle) {
  const questions = bundle.questions;
  if (!questions.length) {
    answerOnlySheet.innerHTML = `
      <div class="empty-state">
        <div>
          <h3>解答用紙はまだ空です</h3>
          <p>問題を追加すると、解答欄だけの見た目がここに表示されます。</p>
        </div>
      </div>
    `;
    return;
  }

  const blocks = questions.map((question, index) => `
    <section class="paper-question">
      <div class="paper-question-top">
        <span>第${index + 1}問</span>
        <span>${escapeHtml(getQuestionPointsLabel(question))}</span>
      </div>
      <p>${escapeHtml(questionTypes[question.type].label)}の解答欄</p>
      ${renderAnswerLinesHtml(getAnswerLinesCount(question))}
    </section>
  `).join("");

  answerOnlySheet.innerHTML = `${renderCommonPaperHeader(bundle, "解答用紙", "問題文なしで、解答欄だけを印刷できます。")}${blocks}`;
}

function getTypeCountMap(questions) {
  return questions.reduce((result, question) => {
    result[question.type] = (result[question.type] || 0) + 1;
    return result;
  }, { multiple: 0, truefalse: 0, short: 0, fill: 0, order: 0 });
}

function getDifficultyCountMap(questions) {
  return questions.reduce((result, question) => {
    result[question.difficulty] = (result[question.difficulty] || 0) + 1;
    return result;
  }, { easy: 0, standard: 0, hard: 0 });
}

function getTagCountMap(questions) {
  const result = {};
  questions.forEach((question) => {
    String(question.unitTag || "")
      .split(/[,\n、/]/)
      .map((tag) => tag.trim())
      .filter(Boolean)
      .forEach((tag) => {
        result[tag] = (result[tag] || 0) + 1;
      });
  });
  return result;
}

function formatSummaryList(entries, emptyLabel) {
  if (!entries.length) {
    return `<p>${escapeHtml(emptyLabel)}</p>`;
  }
  return `<ul class="summary-list">${entries.map((entry) => `<li>${entry}</li>`).join("")}</ul>`;
}

function renderCheckView(validation, bundle) {
  const typeMap = getTypeCountMap(bundle.questions);
  const difficultyMap = getDifficultyCountMap(bundle.questions);
  const tagMap = getTagCountMap(bundle.questions);

  const typeEntries = Object.keys(questionTypes).map((key) => `${escapeHtml(questionTypes[key].label)}: ${typeMap[key] || 0}問`);
  const difficultyEntries = Object.keys(difficultyLabels).map((key) => `${escapeHtml(difficultyLabels[key])}: ${difficultyMap[key] || 0}問`);
  const tagEntries = Object.entries(tagMap)
    .sort((left, right) => right[1] - left[1])
    .slice(0, 8)
    .map(([tag, count]) => `${escapeHtml(tag)}: ${count}問`);

  analysisSheet.innerHTML = `
    <header class="paper-header">
      <div class="paper-kicker">
        <span>印刷前チェック</span>
        <span>${escapeHtml(bundle.label)}</span>
      </div>
      <h2 class="paper-title">確認メモ</h2>
      <p class="paper-notes">合計点、入力不足、問題形式の偏りをまとめて確認できます。</p>
    </header>
    <div class="analysis-grid">
      <section class="analysis-card">
        <h3>入力チェック</h3>
        <div class="validation-list">
          ${renderValidationItemsHtml([...validation.errors, ...validation.warnings], "気になる点はありません。")}
        </div>
      </section>
      <section class="analysis-card">
        <h3>点数まとめ</h3>
        ${formatSummaryList([
          `問題数: ${bundle.questions.length}問`,
          `合計点: ${bundle.questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0)}点`,
          `設定した満点: ${validation.maxScoreValue === null ? "未設定" : `${validation.maxScoreValue}点`}`,
          `制限時間: ${validation.durationValue === null ? "未設定" : `${validation.durationValue}分`}`
        ], "まだ集計できません。")}
      </section>
      <section class="analysis-card">
        <h3>問題形式</h3>
        ${formatSummaryList(typeEntries, "問題がありません。")}
      </section>
      <section class="analysis-card">
        <h3>難しさ</h3>
        ${formatSummaryList(difficultyEntries, "問題がありません。")}
      </section>
      <section class="analysis-card">
        <h3>単元タグ</h3>
        ${formatSummaryList(tagEntries, "単元タグはまだありません。")}
      </section>
      <section class="analysis-card">
        <h3>現在の状態</h3>
        ${formatSummaryList([
          validation.errors.length ? `修正が必要: ${validation.errors.length}件` : "修正が必要な項目はありません。",
          validation.warnings.length ? `注意したい項目: ${validation.warnings.length}件` : "注意だけの項目はありません。",
          bundle.isPattern ? `${bundle.label}を表示中です。` : "元のテストを表示中です。"
        ], "確認情報はありません。")}
      </section>
    </div>
  `;
}

function applyPaperSizeToPreview() {
  const config = getPaperConfig();
  const previewHeight = Math.round(config.previewWidthPx * (config.heightMm / config.widthMm));

  [paperSheet, answerSheet, answerOnlySheet, analysisSheet].forEach((sheet) => {
    sheet.style.setProperty("--paper-screen-width", `${config.previewWidthPx}px`);
    sheet.style.setProperty("--paper-screen-height", `${previewHeight}px`);
    sheet.dataset.paperSize = config.label;
    sheet.dataset.layoutMode = state.meta.layoutMode || "standard";
  });

  previewPaperLabel.textContent = `${config.label} / 中央の紙と印刷に反映`;
}

function updatePreviewPrintButton() {
  const bundle = getActiveBundle();
  if (uiState.activeTab === "answer") {
    printPreviewButton.textContent = `${bundle.label}の解答付きPDF保存・印刷`;
    return;
  }
  if (uiState.activeTab === "answeronly") {
    printPreviewButton.textContent = `${bundle.label}の解答用紙だけ印刷`;
    return;
  }
  printPreviewButton.textContent = `${bundle.label}の生徒用PDF保存・印刷`;
}

function renderDerivedViews(validation = getValidationState()) {
  const bundle = getActiveBundle();
  applyPaperSizeToPreview();
  renderTemplateSummary();
  renderStats(validation);
  renderPaperView(bundle);
  renderAnswerView(bundle);
  renderAnswerOnlyView(bundle);
  renderCheckView(validation, bundle);
  renderBankList();
  renderPatternList();
  updatePreviewScope();
  updatePreviewPrintButton();
  editorStatus.textContent = uiState.editorMessage;
}

function updatePreviewPrintButton() {
  const bundle = getActiveBundle();
  if (uiState.activeTab === "answer") {
    printPreviewButton.textContent = `${bundle.label}の解答付き教師用を印刷またはPDF保存`;
    return;
  }
  if (uiState.activeTab === "answeronly") {
    printPreviewButton.textContent = `${bundle.label}の解答用紙を印刷またはPDF保存`;
    return;
  }
  printPreviewButton.textContent = `${bundle.label}の問題用紙を印刷またはPDF保存`;
}

function buildPrintableHtml(mode, bundle) {
  const normalizedMode = mode === "student" ? "paper" : mode;
  const paperConfig = getPaperConfig();
  const layout = getEffectiveLayoutConfig();
  const totalPointsValue = bundle.questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0);
  const labelMap = {
    paper: "問題用紙",
    answeronly: "解答用紙",
    teacher: "解答付き教師用",
    "paper-answer-set": "問題用紙 + 解答用紙"
  };
  const label = labelMap[normalizedMode] || "問題用紙";
  const printMetaColumns = state.meta.showGroupNumberFields && state.meta.showNameField ? "1fr 1fr 1.4fr" : state.meta.showGroupNumberFields ? "1fr 1fr" : "1fr";

  const buildHeader = (sectionLabel) => {
    const duration = parsePositiveIntStrict(state.meta.durationMinutes);
    const maxScore = parsePositiveIntStrict(state.meta.maxScore);
    const studentFields = [];

    if (state.meta.showGroupNumberFields) {
      studentFields.push(`<div class="meta-box"><div class="meta-box-label">${escapeHtml(getMetaValue("groupLabel", "組"))}</div><div class="meta-box-line"></div></div>`);
      studentFields.push(`<div class="meta-box"><div class="meta-box-label">${escapeHtml(getMetaValue("numberLabel", "番号"))}</div><div class="meta-box-line"></div></div>`);
    }
    if (state.meta.showNameField) {
      studentFields.push(`<div class="meta-box wide"><div class="meta-box-label">${escapeHtml(getMetaValue("nameLabel", "名前"))}</div><div class="meta-box-line"></div></div>`);
    }

    return `
      <header class="print-header">
        <div class="print-header-top">
          <span>${escapeHtml(getMetaValue("schoolName", "学校名"))}</span>
          <span>${escapeHtml(getMetaValue("subject", "教科"))}</span>
        </div>
        <h1 class="print-header-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))}</h1>
        <div class="print-header-subtitle">${escapeHtml(sectionLabel)} / ${escapeHtml(bundle.label)}</div>
        <div class="print-summary">
          <span><strong>学年:</strong> ${escapeHtml(getMetaValue("grade", "未設定"))}</span>
          <span><strong>単元:</strong> ${escapeHtml(getMetaValue("unit", "未設定"))}</span>
          <span><strong>クラス:</strong> ${escapeHtml(getMetaValue("className", "未設定"))}</span>
          <span><strong>実施日:</strong> ${escapeHtml(formatDateLabel(state.meta.examDate))}</span>
          <span><strong>制限時間:</strong> ${escapeHtml(duration.valid ? `${duration.value}分` : "未設定")}</span>
          <span><strong>満点:</strong> ${escapeHtml(maxScore.valid ? `${maxScore.value}点` : "未設定")}</span>
          <span><strong>用紙:</strong> ${escapeHtml(paperConfig.label)}</span>
          <span><strong>枚数:</strong> ${escapeHtml(paperConfig.tileLabel)}</span>
          <span><strong>合計サイズ:</strong> ${escapeHtml(formatMeasure(paperConfig.widthMm))}mm × ${escapeHtml(formatMeasure(paperConfig.heightMm))}mm</span>
          <span><strong>合計点:</strong> ${totalPointsValue}点</span>
        </div>
        ${studentFields.length ? `<div class="print-meta-grid">${studentFields.join("")}</div>` : ""}
        <div class="print-instructions">${formatMultilineHtml(state.meta.instructions, defaultInstructions)}</div>
      </header>
    `;
  };

  const renderPrintQuestion = (question, index, sectionMode) => {
    const questionLayout = getQuestionLayout(question, layout);
    const style = [
      `padding:${formatMeasure(questionLayout.blockPaddingMm)}mm`,
      `margin-bottom:${formatMeasure(layout.questionGap)}mm`,
      `font-size:${formatMeasure(questionLayout.questionFontSize)}px`,
      `width:min(100%, ${formatMeasure(questionLayout.blockWidthPercent)}%)`
    ].join("; ");

    if (sectionMode === "answeronly") {
      return `
        <section class="print-question" style="${style}">
          <div class="print-question-top">
            <div class="print-question-title">第${index + 1}問</div>
            <div class="print-question-points">${escapeHtml(getQuestionPointsLabel(question))}</div>
          </div>
          <div class="print-prompt">${escapeHtml(questionTypes[question.type].label)}の解答欄</div>
          ${renderPrintableStudentAnswerArea(question, layout)}
        </section>
      `;
    }

    const teacherBox = sectionMode === "teacher"
      ? `
        <div class="teacher-answer">
          <div class="teacher-answer-label">正解</div>
          <div class="teacher-answer-value">${formatMultilineHtml(question.answer, "未入力")}</div>
        </div>
        <div class="teacher-note">
          <strong>解説:</strong>
          <span>${formatMultilineHtml(question.explanation, "解説はまだありません。")}</span>
        </div>
      `
      : "";

    return `
      <section class="print-question" style="${style}">
        <div class="print-question-top">
          <div class="print-question-title">第${index + 1}問 ${escapeHtml(questionTypes[question.type].label)}</div>
          <div class="print-question-points">${escapeHtml(getQuestionPointsLabel(question))}</div>
        </div>
        <div class="print-badges">
          <span>${escapeHtml(difficultyLabels[question.difficulty])}</span>
          <span>${escapeHtml(question.unitTag || "単元タグなし")}</span>
        </div>
        <p class="print-prompt">${formatMultilineHtml(question.prompt, "ここに問題文を入力してください。")}</p>
        ${renderPrintableChoices(question, layout)}
        ${questionLayout.figureSpaceHeightMm > 0 ? `<div class="figure-space" style="height:${formatMeasure(questionLayout.figureSpaceHeightMm)}mm;">図や画像のスペース</div>` : ""}
        ${renderPrintableStudentAnswerArea(question, layout)}
        ${teacherBox}
      </section>
    `;
  };

  const buildSection = (sectionLabel, sectionMode) => `
    <section class="sheet">
      ${buildHeader(sectionLabel)}
      ${bundle.questions.map((question, index) => renderPrintQuestion(question, index, sectionMode)).join("")}
      ${sectionMode === "teacher" ? `<div class="footer-total">合計点: ${totalPointsValue}点</div>` : ""}
    </section>
  `;

  const sections = normalizedMode === "paper-answer-set"
    ? [buildSection("問題用紙", "paper"), buildSection("解答用紙", "answeronly")]
    : [buildSection(label, normalizedMode)];

  return `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${escapeHtml(getMetaValue("examTitle", "テスト名"))} ${escapeHtml(label)} ${escapeHtml(bundle.label)}</title>
      <style>
        :root { color-scheme: light; }
        * { box-sizing: border-box; }
        body { margin: 0; font-family: "Yu Gothic", "Hiragino Sans", sans-serif; background: #edf1ef; color: #111; }
        .toolbar { position: sticky; top: 0; z-index: 10; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 16px 20px; background: #24343b; color: #fff; }
        .toolbar-title { font-size: 1rem; font-weight: 700; }
        .toolbar-note { font-size: 0.88rem; opacity: 0.88; }
        .toolbar-button { border: none; border-radius: 999px; padding: 12px 18px; background: #fff; color: #24343b; font-size: 0.95rem; font-weight: 700; cursor: pointer; }
        .sheet { width: ${formatMeasure(paperConfig.widthMm)}mm; min-height: ${formatMeasure(paperConfig.heightMm)}mm; margin: 18px auto; padding: ${formatMeasure(layout.topBottomMarginMm)}mm ${formatMeasure(layout.sideMarginMm)}mm; background: #fff; box-shadow: 0 12px 30px rgba(0, 0, 0, 0.08); font-size: ${formatMeasure(layout.globalFontSize)}px; line-height: ${formatMeasure(layout.lineHeight)}; }
        .sheet + .sheet { page-break-before: always; break-before: page; }
        .print-header { border-bottom: 2px solid #222; padding-bottom: 10mm; margin-bottom: 8mm; }
        .print-header-top, .print-summary { display: flex; flex-wrap: wrap; gap: 12px 18px; font-size: 0.94rem; }
        .print-header-title { margin: 8mm 0 3mm; text-align: center; font-size: 1.8rem; font-weight: 700; letter-spacing: 0.04em; }
        .print-header-subtitle { text-align: center; font-size: 1rem; font-weight: 700; margin-bottom: 6mm; }
        .print-meta-grid { display: grid; grid-template-columns: ${printMetaColumns}; gap: 12px; margin: 6mm 0; }
        .meta-box { border: 1px solid #333; padding: 10px 12px; min-height: 16mm; }
        .meta-box-label { font-size: 0.82rem; margin-bottom: 6px; font-weight: 700; }
        .meta-box-line { border-bottom: 1px solid #222; height: 18px; }
        .print-instructions { margin-top: 4mm; padding: 10px 12px; border: 1px solid #444; background: #fafafa; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .print-question { border-bottom: 1px solid #bbb; page-break-inside: avoid; break-inside: avoid; }
        .print-question:last-child { border-bottom: none; }
        .print-question-top { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 4px; }
        .print-question-title { font-size: 1.08rem; font-weight: 700; }
        .print-question-points { min-width: 68px; text-align: right; font-size: 0.98rem; font-weight: 700; }
        .print-badges { display: flex; flex-wrap: wrap; gap: 8px; margin: 6px 0 10px; }
        .print-badges span { border: 1px solid #666; padding: 3px 8px; font-size: 0.8rem; }
        .print-prompt, .teacher-answer-value, .teacher-note { margin: 0 0 10px; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .print-choices { margin: 0 0 10px 1.4em; padding: 0; white-space: pre-wrap; overflow-wrap: anywhere; }
        .response-lines { display: grid; gap: 10px; margin-top: 12px; }
        .write-line { border-bottom: 1px solid #222; min-height: 24px; }
        .teacher-answer { margin-top: 12px; border: 1px solid #222; background: #f8f8f8; padding: 10px 12px; }
        .teacher-answer-label { font-size: 0.82rem; font-weight: 700; margin-bottom: 6px; }
        .teacher-note strong { display: inline-block; margin-right: 6px; }
        .figure-space { display: grid; place-items: center; margin-top: 8px; border: 1px dashed #aaa; background: #fafafa; color: #566; font-size: 0.84rem; }
        .footer-total { margin-top: 8mm; text-align: right; font-size: 1rem; font-weight: 700; }
        @page { size: ${formatMeasure(paperConfig.widthMm)}mm ${formatMeasure(paperConfig.heightMm)}mm; margin: 0; }
        @media print { body { background: #fff; } .toolbar { display: none; } .sheet { width: auto; min-height: auto; margin: 0; box-shadow: none; } }
      </style>
    </head>
    <body>
      <div class="toolbar">
        <div>
          <div class="toolbar-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))} / ${escapeHtml(label)} / ${escapeHtml(bundle.label)}</div>
          <div class="toolbar-note">上のボタンから印刷またはPDF保存に進めます。</div>
        </div>
        <button class="toolbar-button" onclick="window.print()">印刷またはPDF保存</button>
      </div>
      ${sections.join("")}
    </body>
    </html>
  `;
}

function openPrintableDocument(mode, bundle = getActiveBundle()) {
  const normalizedMode = mode === "student" ? "paper" : mode;
  const labelMap = {
    paper: "問題用紙を印刷",
    answeronly: "解答用紙を印刷",
    teacher: "解答付き教師用を印刷",
    "paper-answer-set": "問題用紙と解答用紙を印刷"
  };
  const label = labelMap[normalizedMode] || "問題用紙を印刷";
  syncMetaFromInputs();
  if (!ensureReadyForOutput(label, bundle.questions)) {
    return;
  }
  const exportWindow = window.open("", "_blank");
  if (!exportWindow) {
    window.alert("新しい画面を開けませんでした。ポップアップブロックを一時的に解除して、もう一度お試しください。");
    return;
  }
  exportWindow.document.open();
  exportWindow.document.write(buildPrintableHtml(normalizedMode, bundle));
  exportWindow.document.close();
  setEditorMessage(`${bundle.label}の${label}を開きました。上のボタンから印刷またはPDF保存に進めます。`);
}

function renderDerivedViews(validation = getValidationState()) {
  const bundle = getActiveBundle();
  applyPaperSizeToPreview();
  renderTemplateSummary();
  renderStats(validation);
  renderPaperView(bundle);
  renderAnswerView(bundle);
  renderAnswerOnlyView(bundle);
  renderCheckView(validation, bundle);
  renderBankList();
  renderPatternList();
  updatePreviewScope();
  updatePreviewPrintButton();
  editorStatus.textContent = uiState.editorMessage;
}

function stage3EnsureCombinedExportButton() {
  const answerButton = document.getElementById("exportAnswerSheet");
  if (!answerButton || document.getElementById("exportStudentAndAnswerPdf")) {
    return;
  }
  const button = document.createElement("button");
  button.type = "button";
  button.id = "exportStudentAndAnswerPdf";
  button.className = "secondary";
  button.textContent = "問題用紙と解答用紙を印刷";
  answerButton.insertAdjacentElement("afterend", button);
  button.addEventListener("click", () => openPrintableDocument("paper-answer-set"));
}

stage3EnsureCombinedExportButton();
renderDerivedViews(getValidationState());

function stage3GetNormalizedPrintMode(mode) {
  if (mode === "student") {
    return "paper";
  }
  return mode;
}

function stage3BuildPrintableHeader(label, totalPointsValue, bundle, paperConfig) {
  const duration = parsePositiveIntStrict(state.meta.durationMinutes);
  const maxScore = parsePositiveIntStrict(state.meta.maxScore);
  const studentFields = [];

  if (state.meta.showGroupNumberFields) {
    studentFields.push(`<div class="meta-box"><div class="meta-box-label">${escapeHtml(getMetaValue("groupLabel", "組"))}</div><div class="meta-box-line"></div></div>`);
    studentFields.push(`<div class="meta-box"><div class="meta-box-label">${escapeHtml(getMetaValue("numberLabel", "番号"))}</div><div class="meta-box-line"></div></div>`);
  }
  if (state.meta.showNameField) {
    studentFields.push(`<div class="meta-box wide"><div class="meta-box-label">${escapeHtml(getMetaValue("nameLabel", "名前"))}</div><div class="meta-box-line"></div></div>`);
  }

  return `
    <header class="print-header">
      <div class="print-header-top">
        <span>${escapeHtml(getMetaValue("schoolName", "学校名"))}</span>
        <span>${escapeHtml(getMetaValue("subject", "教科"))}</span>
      </div>
      <h1 class="print-header-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))}</h1>
      <div class="print-header-subtitle">${escapeHtml(label)} / ${escapeHtml(bundle.label)}</div>
      <div class="print-summary">
        <span><strong>学年:</strong> ${escapeHtml(getMetaValue("grade", "未設定"))}</span>
        <span><strong>単元:</strong> ${escapeHtml(getMetaValue("unit", "未設定"))}</span>
        <span><strong>クラス:</strong> ${escapeHtml(getMetaValue("className", "未設定"))}</span>
        <span><strong>実施日:</strong> ${escapeHtml(formatDateLabel(state.meta.examDate))}</span>
        <span><strong>制限時間:</strong> ${escapeHtml(duration.valid ? `${duration.value}分` : "未設定")}</span>
        <span><strong>満点:</strong> ${escapeHtml(maxScore.valid ? `${maxScore.value}点` : "未設定")}</span>
        <span><strong>用紙:</strong> ${escapeHtml(paperConfig.label)}</span>
        <span><strong>枚数:</strong> ${escapeHtml(paperConfig.tileLabel)}</span>
        <span><strong>合計サイズ:</strong> ${escapeHtml(formatMeasure(paperConfig.widthMm))}mm × ${escapeHtml(formatMeasure(paperConfig.heightMm))}mm</span>
        <span><strong>合計点:</strong> ${totalPointsValue}点</span>
      </div>
      ${studentFields.length ? `<div class="print-meta-grid">${studentFields.join("")}</div>` : ""}
      <div class="print-instructions">${formatMultilineHtml(state.meta.instructions, defaultInstructions)}</div>
    </header>
  `;
}

function stage3RenderPrintableQuestion(question, index, mode, layout) {
  const questionLayout = getQuestionLayout(question, layout);
  const style = [
    `padding:${formatMeasure(questionLayout.blockPaddingMm)}mm`,
    `margin-bottom:${formatMeasure(layout.questionGap)}mm`,
    `font-size:${formatMeasure(questionLayout.questionFontSize)}px`,
    `width:min(100%, ${formatMeasure(questionLayout.blockWidthPercent)}%)`
  ].join("; ");

  if (mode === "answeronly") {
    return `
      <section class="print-question" style="${style}">
        <div class="print-question-top">
          <div class="print-question-title">第${index + 1}問</div>
          <div class="print-question-points">${escapeHtml(getQuestionPointsLabel(question))}</div>
        </div>
        <div class="print-prompt">${escapeHtml(questionTypes[question.type].label)}の解答欄</div>
        ${renderPrintableStudentAnswerArea(question, layout)}
      </section>
    `;
  }

  const teacherBox = mode === "teacher"
    ? `
      <div class="teacher-answer">
        <div class="teacher-answer-label">正解</div>
        <div class="teacher-answer-value">${formatMultilineHtml(question.answer, "未入力")}</div>
      </div>
      <div class="teacher-note">
        <strong>解説:</strong>
        <span>${formatMultilineHtml(question.explanation, "解説はまだありません。")}</span>
      </div>
    `
    : "";

  return `
    <section class="print-question" style="${style}">
      <div class="print-question-top">
        <div class="print-question-title">第${index + 1}問 ${escapeHtml(questionTypes[question.type].label)}</div>
        <div class="print-question-points">${escapeHtml(getQuestionPointsLabel(question))}</div>
      </div>
      <div class="print-badges">
        <span>${escapeHtml(difficultyLabels[question.difficulty])}</span>
        <span>${escapeHtml(question.unitTag || "単元タグなし")}</span>
      </div>
      <p class="print-prompt">${formatMultilineHtml(question.prompt, "ここに問題文を入力してください。")}</p>
      ${renderPrintableChoices(question, layout)}
      ${questionLayout.figureSpaceHeightMm > 0 ? `<div class="figure-space" style="height:${formatMeasure(questionLayout.figureSpaceHeightMm)}mm;">図や画像のスペース</div>` : ""}
      ${renderPrintableStudentAnswerArea(question, layout)}
      ${teacherBox}
    </section>
  `;
}

function stage3BuildPrintableSection(label, mode, bundle, layout, paperConfig, totalPointsValue) {
  const questionsHtml = bundle.questions
    .map((question, index) => stage3RenderPrintableQuestion(question, index, mode, layout))
    .join("");

  return `
    <section class="sheet ${mode === "answeronly" ? "sheet-answeronly" : ""}">
      ${stage3BuildPrintableHeader(label, totalPointsValue, bundle, paperConfig)}
      ${questionsHtml}
      ${mode === "teacher" ? `<div class="footer-total">合計点: ${totalPointsValue}点</div>` : ""}
    </section>
  `;
}

function buildPrintableHtml(mode, bundle) {
  const normalizedMode = stage3GetNormalizedPrintMode(mode);
  const paperConfig = getPaperConfig();
  const layout = getEffectiveLayoutConfig();
  const totalPointsValue = bundle.questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0);
  const labelMap = {
    paper: "問題用紙",
    answeronly: "解答用紙",
    teacher: "解答付き教師用",
    "paper-answer-set": "問題用紙 + 解答用紙"
  };
  const label = labelMap[normalizedMode] || "問題用紙";
  const printMetaColumns = state.meta.showGroupNumberFields && state.meta.showNameField ? "1fr 1fr 1.4fr" : state.meta.showGroupNumberFields ? "1fr 1fr" : "1fr";
  const sections = normalizedMode === "paper-answer-set"
    ? [
        stage3BuildPrintableSection("問題用紙", "paper", bundle, layout, paperConfig, totalPointsValue),
        stage3BuildPrintableSection("解答用紙", "answeronly", bundle, layout, paperConfig, totalPointsValue)
      ]
    : [stage3BuildPrintableSection(label, normalizedMode, bundle, layout, paperConfig, totalPointsValue)];

  return `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${escapeHtml(getMetaValue("examTitle", "テスト名"))} ${escapeHtml(label)} ${escapeHtml(bundle.label)}</title>
      <style>
        :root { color-scheme: light; }
        * { box-sizing: border-box; }
        body { margin: 0; font-family: "Yu Gothic", "Hiragino Sans", sans-serif; background: #edf1ef; color: #111; }
        .toolbar { position: sticky; top: 0; z-index: 10; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 16px 20px; background: #24343b; color: #fff; }
        .toolbar-title { font-size: 1rem; font-weight: 700; }
        .toolbar-note { font-size: 0.88rem; opacity: 0.88; }
        .toolbar-button { border: none; border-radius: 999px; padding: 12px 18px; background: #fff; color: #24343b; font-size: 0.95rem; font-weight: 700; cursor: pointer; }
        .sheet { width: ${formatMeasure(paperConfig.widthMm)}mm; min-height: ${formatMeasure(paperConfig.heightMm)}mm; margin: 18px auto; padding: ${formatMeasure(layout.topBottomMarginMm)}mm ${formatMeasure(layout.sideMarginMm)}mm; background: #fff; box-shadow: 0 12px 30px rgba(0, 0, 0, 0.08); font-size: ${formatMeasure(layout.globalFontSize)}px; line-height: ${formatMeasure(layout.lineHeight)}; }
        .sheet + .sheet { page-break-before: always; break-before: page; }
        .print-header { border-bottom: 2px solid #222; padding-bottom: 10mm; margin-bottom: 8mm; }
        .print-header-top, .print-summary { display: flex; flex-wrap: wrap; gap: 12px 18px; font-size: 0.94rem; }
        .print-header-title { margin: 8mm 0 3mm; text-align: center; font-size: 1.8rem; font-weight: 700; letter-spacing: 0.04em; }
        .print-header-subtitle { text-align: center; font-size: 1rem; font-weight: 700; margin-bottom: 6mm; }
        .print-meta-grid { display: grid; grid-template-columns: ${printMetaColumns}; gap: 12px; margin: 6mm 0; }
        .meta-box { border: 1px solid #333; padding: 10px 12px; min-height: 16mm; }
        .meta-box-label { font-size: 0.82rem; margin-bottom: 6px; font-weight: 700; }
        .meta-box-line { border-bottom: 1px solid #222; height: 18px; }
        .print-instructions { margin-top: 4mm; padding: 10px 12px; border: 1px solid #444; background: #fafafa; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .print-question { border-bottom: 1px solid #bbb; page-break-inside: avoid; break-inside: avoid; }
        .print-question:last-child { border-bottom: none; }
        .print-question-top { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 4px; }
        .print-question-title { font-size: 1.08rem; font-weight: 700; }
        .print-question-points { min-width: 68px; text-align: right; font-size: 0.98rem; font-weight: 700; }
        .print-badges { display: flex; flex-wrap: wrap; gap: 8px; margin: 6px 0 10px; }
        .print-badges span { border: 1px solid #666; padding: 3px 8px; font-size: 0.8rem; }
        .print-prompt, .teacher-answer-value, .teacher-note { margin: 0 0 10px; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .print-choices { margin: 0 0 10px 1.4em; padding: 0; white-space: pre-wrap; overflow-wrap: anywhere; }
        .response-lines { display: grid; gap: 10px; margin-top: 12px; }
        .write-line { border-bottom: 1px solid #222; min-height: 24px; }
        .teacher-answer { margin-top: 12px; border: 1px solid #222; background: #f8f8f8; padding: 10px 12px; }
        .teacher-answer-label { font-size: 0.82rem; font-weight: 700; margin-bottom: 6px; }
        .teacher-note strong { display: inline-block; margin-right: 6px; }
        .figure-space { display: grid; place-items: center; margin-top: 8px; border: 1px dashed #aaa; background: #fafafa; color: #566; font-size: 0.84rem; }
        .footer-total { margin-top: 8mm; text-align: right; font-size: 1rem; font-weight: 700; }
        @page { size: ${formatMeasure(paperConfig.widthMm)}mm ${formatMeasure(paperConfig.heightMm)}mm; margin: 0; }
        @media print { body { background: #fff; } .toolbar { display: none; } .sheet { width: auto; min-height: auto; margin: 0; box-shadow: none; } }
      </style>
    </head>
    <body>
      <div class="toolbar">
        <div>
          <div class="toolbar-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))} / ${escapeHtml(label)} / ${escapeHtml(bundle.label)}</div>
          <div class="toolbar-note">上のボタンから印刷またはPDF保存に進めます。</div>
        </div>
        <button class="toolbar-button" onclick="window.print()">印刷またはPDF保存</button>
      </div>
      ${sections.join("")}
    </body>
    </html>
  `;
}

function openPrintableDocument(mode, bundle = getActiveBundle()) {
  const normalizedMode = stage3GetNormalizedPrintMode(mode);
  const labelMap = {
    paper: "問題用紙を印刷",
    answeronly: "解答用紙を印刷",
    teacher: "解答付き教師用を印刷",
    "paper-answer-set": "問題用紙と解答用紙を印刷"
  };
  const label = labelMap[normalizedMode] || "問題用紙を印刷";
  syncMetaFromInputs();
  if (!ensureReadyForOutput(label, bundle.questions)) {
    return;
  }
  const exportWindow = window.open("", "_blank");
  if (!exportWindow) {
    window.alert("新しい画面を開けませんでした。ポップアップブロックを一時的に解除して、もう一度お試しください。");
    return;
  }
  exportWindow.document.open();
  exportWindow.document.write(buildPrintableHtml(normalizedMode, bundle));
  exportWindow.document.close();
  setEditorMessage(`${bundle.label}の${label}を開きました。上のボタンから印刷またはPDF保存に進めます。`);
}

function stage3EnsureCombinedExportButton() {
  const exportActions = document.querySelector(".export-actions");
  const answerButton = document.getElementById("exportAnswerSheet");
  if (!exportActions || !answerButton || document.getElementById("exportStudentAndAnswerPdf")) {
    return;
  }
  const button = document.createElement("button");
  button.type = "button";
  button.id = "exportStudentAndAnswerPdf";
  button.className = "secondary";
  button.textContent = "問題用紙と解答用紙を印刷";
  answerButton.insertAdjacentElement("afterend", button);
  button.addEventListener("click", () => openPrintableDocument("paper-answer-set"));
}

stage3EnsureCombinedExportButton();
renderDerivedViews(getValidationState());

function ensureQuestionsForOutput(label, questions) {
  if (questions.length) {
    return true;
  }
  window.alert(`${label}の前に、少なくとも1問は問題を追加してください。`);
  setEditorMessage(`${label}の前に問題を追加してください。`);
  return false;
}

function ensureReadyForOutput(label, questions) {
  if (!ensureQuestionsForOutput(label, questions)) {
    return false;
  }
  const validation = getValidationState();
  if (!validation.errors.length) {
    if (validation.warnings.length) {
      setEditorMessage(`${label}を開きました。注意メッセージはありますが、印刷はできます。`);
    }
    return true;
  }
  switchTab("check");
  const messages = validation.errors.slice(0, 5).map((item) => `- ${item.message}`).join("\n");
  window.alert(`${label}の前に次の項目を確認してください。\n${messages}`);
  setEditorMessage(`${label}の前に、チェック欄の内容を確認してください。`);
  return false;
}

function renderPrintableStudentAnswerArea(question, layout = getEffectiveLayoutConfig()) {
  const questionLayout = getQuestionLayout(question, layout);
  return `
    <div class="response-lines" style="min-height:${formatMeasure(questionLayout.answerAreaHeightMm)}mm;">
      ${Array.from({ length: getAnswerLinesCount(question, layout) }, () => '<div class="write-line"></div>').join("")}
    </div>
  `;
}

function renderPrintableChoices(question, layout = getEffectiveLayoutConfig()) {
  const items = getChoiceItems(question);
  if (!items.length) {
    return "";
  }
  const questionLayout = getQuestionLayout(question, layout);
  if (question.type === "order") {
    return `<ul class="print-choices" style="font-size:${formatMeasure(questionLayout.choiceFontSize)}px;">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
  }
  return `<ol class="print-choices" style="font-size:${formatMeasure(questionLayout.choiceFontSize)}px;">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ol>`;
}

function renderPrintableQuestion(question, index, mode, layout = getEffectiveLayoutConfig()) {
  const questionLayout = getQuestionLayout(question, layout);
  const style = [
    `padding:${formatMeasure(questionLayout.blockPaddingMm)}mm`,
    `margin-bottom:${formatMeasure(layout.questionGap)}mm`,
    `font-size:${formatMeasure(questionLayout.questionFontSize)}px`,
    `width:min(100%, ${formatMeasure(questionLayout.blockWidthPercent)}%)`
  ].join("; ");

  if (mode === "answeronly") {
    return `
      <section class="print-question" style="${style}">
        <div class="print-question-top">
          <div class="print-question-title">第${index + 1}問</div>
          <div class="print-question-points">${escapeHtml(getQuestionPointsLabel(question))}</div>
        </div>
        <div class="print-prompt">${escapeHtml(questionTypes[question.type].label)}の解答欄</div>
        ${renderPrintableStudentAnswerArea(question, layout)}
      </section>
    `;
  }

  const teacherBox = mode === "teacher"
    ? `
      <div class="teacher-answer">
        <div class="teacher-answer-label">正解</div>
        <div class="teacher-answer-value">${formatMultilineHtml(question.answer, "未入力")}</div>
      </div>
      <div class="teacher-note">
        <strong>解説:</strong>
        <span>${formatMultilineHtml(question.explanation, "解説はまだありません。")}</span>
      </div>
    `
    : "";

  return `
    <section class="print-question" style="${style}">
      <div class="print-question-top">
        <div class="print-question-title">第${index + 1}問 ${escapeHtml(questionTypes[question.type].label)}</div>
        <div class="print-question-points">${escapeHtml(getQuestionPointsLabel(question))}</div>
      </div>
      <div class="print-badges">
        <span>${escapeHtml(difficultyLabels[question.difficulty])}</span>
        <span>${escapeHtml(question.unitTag || "単元タグなし")}</span>
      </div>
      <p class="print-prompt">${formatMultilineHtml(question.prompt, "ここに問題文を入力してください。")}</p>
      ${renderPrintableChoices(question, layout)}
      ${questionLayout.figureSpaceHeightMm > 0 ? `<div class="figure-space" style="height:${formatMeasure(questionLayout.figureSpaceHeightMm)}mm;">図や画像のスペース</div>` : ""}
      ${renderPrintableStudentAnswerArea(question, layout)}
      ${teacherBox}
    </section>
  `;
}

function renderPrintableAnswerOnlyQuestion(question, index, layout = getEffectiveLayoutConfig()) {
  return renderPrintableQuestion(question, index, "answeronly", layout);
}

function buildPrintableHeader(label, totalPointsValue, bundle) {
  const paperConfig = getPaperConfig();
  const duration = parsePositiveIntStrict(state.meta.durationMinutes);
  const maxScore = parsePositiveIntStrict(state.meta.maxScore);

  const studentFields = [];
  if (state.meta.showGroupNumberFields) {
    studentFields.push(`<div class="meta-box"><div class="meta-box-label">${escapeHtml(getMetaValue("groupLabel", "組"))}</div><div class="meta-box-line"></div></div>`);
    studentFields.push(`<div class="meta-box"><div class="meta-box-label">${escapeHtml(getMetaValue("numberLabel", "番号"))}</div><div class="meta-box-line"></div></div>`);
  }
  if (state.meta.showNameField) {
    studentFields.push(`<div class="meta-box wide"><div class="meta-box-label">${escapeHtml(getMetaValue("nameLabel", "名前"))}</div><div class="meta-box-line"></div></div>`);
  }

  return `
    <header class="print-header">
      <div class="print-header-top">
        <span>${escapeHtml(getMetaValue("schoolName", "学校名"))}</span>
        <span>${escapeHtml(getMetaValue("subject", "教科"))}</span>
      </div>
      <h1 class="print-header-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))}</h1>
      <div class="print-header-subtitle">${escapeHtml(label)} / ${escapeHtml(bundle.label)}</div>
      <div class="print-summary">
        <span><strong>学年:</strong> ${escapeHtml(getMetaValue("grade", "未設定"))}</span>
        <span><strong>単元:</strong> ${escapeHtml(getMetaValue("unit", "未設定"))}</span>
        <span><strong>クラス:</strong> ${escapeHtml(getMetaValue("className", "未設定"))}</span>
        <span><strong>実施日:</strong> ${escapeHtml(formatDateLabel(state.meta.examDate))}</span>
        <span><strong>制限時間:</strong> ${escapeHtml(duration.valid ? `${duration.value}分` : "未設定")}</span>
        <span><strong>満点:</strong> ${escapeHtml(maxScore.valid ? `${maxScore.value}点` : "未設定")}</span>
        <span><strong>用紙:</strong> ${escapeHtml(paperConfig.label)}</span>
        <span><strong>枚数:</strong> ${escapeHtml(paperConfig.tileLabel)}</span>
        <span><strong>合計サイズ:</strong> ${escapeHtml(formatMeasure(paperConfig.widthMm))}mm × ${escapeHtml(formatMeasure(paperConfig.heightMm))}mm</span>
        <span><strong>合計点:</strong> ${totalPointsValue}点</span>
      </div>
      ${studentFields.length ? `<div class="print-meta-grid">${studentFields.join("")}</div>` : ""}
      <div class="print-instructions">${formatMultilineHtml(state.meta.instructions, defaultInstructions)}</div>
    </header>
  `;
}

function buildPrintableHtml(mode, bundle) {
  const paperConfig = getPaperConfig();
  const layout = getEffectiveLayoutConfig();
  const totalPointsValue = bundle.questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0);
  const label = mode === "teacher" ? "解答付き教師用" : mode === "answeronly" ? "解答用紙" : "問題用紙";
  const questionsHtml = mode === "answeronly"
    ? bundle.questions.map((question, index) => renderPrintableAnswerOnlyQuestion(question, index, layout)).join("")
    : bundle.questions.map((question, index) => renderPrintableQuestion(question, index, mode, layout)).join("");
  const printMetaColumns = state.meta.showGroupNumberFields && state.meta.showNameField ? "1fr 1fr 1.4fr" : state.meta.showGroupNumberFields ? "1fr 1fr" : "1fr";

  return `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${escapeHtml(getMetaValue("examTitle", "テスト名"))} ${escapeHtml(label)} ${escapeHtml(bundle.label)}</title>
      <style>
        :root { color-scheme: light; }
        * { box-sizing: border-box; }
        body { margin: 0; font-family: "Yu Gothic", "Hiragino Sans", sans-serif; background: #edf1ef; color: #111; }
        .toolbar { position: sticky; top: 0; z-index: 10; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 16px 20px; background: #24343b; color: #fff; }
        .toolbar-title { font-size: 1rem; font-weight: 700; }
        .toolbar-note { font-size: 0.88rem; opacity: 0.88; }
        .toolbar-button { border: none; border-radius: 999px; padding: 12px 18px; background: #fff; color: #24343b; font-size: 0.95rem; font-weight: 700; cursor: pointer; }
        .sheet { width: ${formatMeasure(paperConfig.widthMm)}mm; min-height: ${formatMeasure(paperConfig.heightMm)}mm; margin: 18px auto; padding: ${formatMeasure(layout.topBottomMarginMm)}mm ${formatMeasure(layout.sideMarginMm)}mm; background: #fff; box-shadow: 0 12px 30px rgba(0, 0, 0, 0.08); font-size: ${formatMeasure(layout.globalFontSize)}px; line-height: ${formatMeasure(layout.lineHeight)}; }
        .print-header { border-bottom: 2px solid #222; padding-bottom: 10mm; margin-bottom: 8mm; }
        .print-header-top, .print-summary { display: flex; flex-wrap: wrap; gap: 12px 18px; font-size: 0.94rem; }
        .print-header-title { margin: 8mm 0 3mm; text-align: center; font-size: 1.8rem; font-weight: 700; letter-spacing: 0.04em; }
        .print-header-subtitle { text-align: center; font-size: 1rem; font-weight: 700; margin-bottom: 6mm; }
        .print-meta-grid { display: grid; grid-template-columns: ${printMetaColumns}; gap: 12px; margin: 6mm 0; }
        .meta-box { border: 1px solid #333; padding: 10px 12px; min-height: 16mm; }
        .meta-box-label { font-size: 0.82rem; margin-bottom: 6px; font-weight: 700; }
        .meta-box-line { border-bottom: 1px solid #222; height: 18px; }
        .print-instructions { margin-top: 4mm; padding: 10px 12px; border: 1px solid #444; background: #fafafa; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .print-question { border-bottom: 1px solid #bbb; page-break-inside: avoid; break-inside: avoid; }
        .print-question:last-child { border-bottom: none; }
        .print-question-top { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 4px; }
        .print-question-title { font-size: 1.08rem; font-weight: 700; }
        .print-question-points { min-width: 68px; text-align: right; font-size: 0.98rem; font-weight: 700; }
        .print-badges { display: flex; flex-wrap: wrap; gap: 8px; margin: 6px 0 10px; }
        .print-badges span { border: 1px solid #666; padding: 3px 8px; font-size: 0.8rem; }
        .print-prompt, .teacher-answer-value, .teacher-note { margin: 0 0 10px; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .print-choices { margin: 0 0 10px 1.4em; padding: 0; white-space: pre-wrap; overflow-wrap: anywhere; }
        .response-lines { display: grid; gap: 10px; margin-top: 12px; }
        .write-line { border-bottom: 1px solid #222; min-height: 24px; }
        .teacher-answer { margin-top: 12px; border: 1px solid #222; background: #f8f8f8; padding: 10px 12px; }
        .teacher-answer-label { font-size: 0.82rem; font-weight: 700; margin-bottom: 6px; }
        .teacher-note strong { display: inline-block; margin-right: 6px; }
        .figure-space { display: grid; place-items: center; margin-top: 8px; border: 1px dashed #aaa; background: #fafafa; color: #566; font-size: 0.84rem; }
        .footer-total { margin-top: 8mm; text-align: right; font-size: 1rem; font-weight: 700; }
        @page { size: ${formatMeasure(paperConfig.widthMm)}mm ${formatMeasure(paperConfig.heightMm)}mm; margin: 0; }
        @media print { body { background: #fff; } .toolbar { display: none; } .sheet { width: auto; min-height: auto; margin: 0; box-shadow: none; } }
      </style>
    </head>
    <body>
      <div class="toolbar">
        <div>
          <div class="toolbar-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))} / ${escapeHtml(label)} / ${escapeHtml(bundle.label)}</div>
          <div class="toolbar-note">上のボタンから印刷またはPDF保存に進めます。</div>
        </div>
        <button class="toolbar-button" onclick="window.print()">印刷またはPDF保存</button>
      </div>
      <main class="sheet">
        ${buildPrintableHeader(label, totalPointsValue, bundle)}
        ${questionsHtml}
        ${mode === "teacher" ? `<div class="footer-total">合計点: ${totalPointsValue}点</div>` : ""}
      </main>
    </body>
    </html>
  `;
}

function openPrintableDocument(mode, bundle = getActiveBundle()) {
  const label = mode === "teacher" ? "解答付き教師用を印刷" : mode === "answeronly" ? "解答用紙を印刷" : "問題用紙を印刷";
  syncMetaFromInputs();
  if (!ensureReadyForOutput(label, bundle.questions)) {
    return;
  }
  const exportWindow = window.open("", "_blank");
  if (!exportWindow) {
    window.alert("新しい画面を開けませんでした。ポップアップブロックを一時的に解除して、もう一度お試しください。");
    return;
  }
  exportWindow.document.open();
  exportWindow.document.write(buildPrintableHtml(mode, bundle));
  exportWindow.document.close();
  setEditorMessage(`${bundle.label}の${label}を開きました。上のボタンから印刷またはPDF保存に進めます。`);
}

function renderAll() {
  const validation = getValidationState();
  renderQuestionList(validation);
  renderDerivedViews(validation);
}

function buildExportFilename(suffix, extension, bundle = getActiveBundle()) {
  const base = getMetaValue("examTitle", "test")
    .replace(/[\\/:*?"<>|]+/g, "_")
    .replace(/\s+/g, "_");
  const variantPart = bundle.key === "base" ? "base" : bundle.key;
  return `${base}_${variantPart}_${suffix}.${extension}`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function switchTab(tabName) {
  uiState.activeTab = tabName;
  document.querySelectorAll(".tab-button").forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === tabName);
  });
  document.querySelectorAll(".paper-view").forEach((view) => {
    view.classList.toggle("active", view.id === `${tabName}-tab`);
  });
  updatePreviewPrintButton();
}

function switchPreviewVariant(variantKey) {
  if (variantKey !== "base" && !state.patterns.generated[variantKey]) {
    return;
  }
  uiState.previewVariant = variantKey;
  renderDerivedViews();
}

function ensureQuestionsForOutput(label, questions) {
  if (questions.length) {
    return true;
  }
  window.alert(`${label}の前に、少なくとも1問は作成してください。`);
  setEditorMessage(`${label}の前に問題を追加してください。`);
  return false;
}

function ensureReadyForOutput(label, questions) {
  if (!ensureQuestionsForOutput(label, questions)) {
    return false;
  }

  const validation = getValidationState();
  if (!validation.errors.length) {
    if (validation.warnings.length) {
      setEditorMessage(`${label}を開きました。満点とのずれなど注意だけの項目はありますが、出力はできます。`);
    }
    return true;
  }

  switchTab("check");
  const messages = validation.errors.slice(0, 5).map((item) => `- ${item.message}`).join("\n");
  window.alert(`${label}の前に次の項目を直してください。\n${messages}`);
  setEditorMessage(`${label}の前に、チェックの内容を直してください。`);
  return false;
}

function renderPrintableStudentAnswerArea(question) {
  return `
    <div class="response-lines">
      ${Array.from({ length: getAnswerLinesCount(question) }, () => '<div class="write-line"></div>').join("")}
    </div>
  `;
}

function renderPrintableChoices(question) {
  const items = getChoiceItems(question);
  if (!items.length) {
    return "";
  }
  if (question.type === "order") {
    return `<ul class="print-choices">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
  }
  return `<ol class="print-choices">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ol>`;
}

function renderPrintableQuestion(question, index, mode) {
  const teacherBox = mode === "teacher"
    ? `
      <div class="teacher-answer">
        <div class="teacher-answer-label">正解</div>
        <div class="teacher-answer-value">${formatMultilineHtml(question.answer, "未入力")}</div>
      </div>
      <div class="teacher-note">
        <strong>解説:</strong>
        <span>${formatMultilineHtml(question.explanation, "解説はまだありません。")}</span>
      </div>
    `
    : "";

  return `
    <section class="print-question">
      <div class="print-question-top">
        <div class="print-question-title">第${index + 1}問 ${escapeHtml(questionTypes[question.type].label)}</div>
        <div class="print-question-points">${escapeHtml(getQuestionPointsLabel(question))}</div>
      </div>
      <div class="print-badges">
        <span>${escapeHtml(difficultyLabels[question.difficulty])}</span>
        <span>${escapeHtml(question.unitTag || "単元タグなし")}</span>
      </div>
      <p class="print-prompt">${formatMultilineHtml(question.prompt, "ここに問題文を入力してください。")}</p>
      ${renderPrintableChoices(question)}
      ${renderPrintableStudentAnswerArea(question)}
      ${teacherBox}
    </section>
  `;
}

function renderPrintableAnswerOnlyQuestion(question, index) {
  return `
    <section class="print-question">
      <div class="print-question-top">
        <div class="print-question-title">第${index + 1}問 ${escapeHtml(questionTypes[question.type].label)}</div>
        <div class="print-question-points">${escapeHtml(getQuestionPointsLabel(question))}</div>
      </div>
      ${renderPrintableStudentAnswerArea(question)}
    </section>
  `;
}

function buildPrintableHeader(label, totalPoints, bundle) {
  const paperConfig = getPaperConfig(state.meta.paperSize);
  const duration = parsePositiveIntStrict(state.meta.durationMinutes);
  const maxScore = parsePositiveIntStrict(state.meta.maxScore);

  const studentFields = [];
  if (state.meta.showGroupNumberFields) {
    studentFields.push(`
      <div class="meta-box">
        <div class="meta-box-label">${escapeHtml(getMetaValue("groupLabel", "組"))}</div>
        <div class="meta-box-line"></div>
      </div>
    `);
    studentFields.push(`
      <div class="meta-box">
        <div class="meta-box-label">${escapeHtml(getMetaValue("numberLabel", "番号"))}</div>
        <div class="meta-box-line"></div>
      </div>
    `);
  }
  if (state.meta.showNameField) {
    studentFields.push(`
      <div class="meta-box wide">
        <div class="meta-box-label">${escapeHtml(getMetaValue("nameLabel", "名前"))}</div>
        <div class="meta-box-line"></div>
      </div>
    `);
  }

  return `
    <header class="print-header">
      <div class="print-header-top">
        <span>${escapeHtml(getMetaValue("schoolName", "学校名"))}</span>
        <span>${escapeHtml(getMetaValue("subject", "教科"))}</span>
      </div>
      <h1 class="print-header-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))}</h1>
      <div class="print-header-subtitle">${escapeHtml(label)} / ${escapeHtml(bundle.label)}</div>
      <div class="print-summary">
        <span><strong>学年:</strong> ${escapeHtml(getMetaValue("grade", "未設定"))}</span>
        <span><strong>単元:</strong> ${escapeHtml(getMetaValue("unit", "未設定"))}</span>
        <span><strong>クラス:</strong> ${escapeHtml(getMetaValue("className", "未設定"))}</span>
        <span><strong>実施日:</strong> ${escapeHtml(formatDateLabel(state.meta.examDate))}</span>
        <span><strong>制限時間:</strong> ${escapeHtml(duration.valid ? `${duration.value}分` : "未設定")}</span>
        <span><strong>満点:</strong> ${escapeHtml(maxScore.valid ? `${maxScore.value}点` : "未設定")}</span>
        <span><strong>用紙:</strong> ${escapeHtml(paperConfig.label)}</span>
        <span><strong>合計点:</strong> ${totalPoints}点</span>
      </div>
      ${studentFields.length ? `<div class="print-meta-grid">${studentFields.join("")}</div>` : ""}
      <div class="print-instructions">${formatMultilineHtml(state.meta.instructions, defaultInstructions)}</div>
    </header>
  `;
}

function buildPrintableHtml(mode, bundle) {
  const paperConfig = getPaperConfig(state.meta.paperSize);
  const questions = bundle.questions;
  const totalPointsValue = questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0);
  const label = mode === "teacher"
    ? "解答付き"
    : mode === "answeronly"
      ? "解答用紙"
      : "生徒用";

  const questionsHtml = mode === "answeronly"
    ? questions.map((question, index) => renderPrintableAnswerOnlyQuestion(question, index)).join("")
    : questions.map((question, index) => renderPrintableQuestion(question, index, mode)).join("");

  const printMetaColumns = state.meta.showGroupNumberFields && state.meta.showNameField ? "1fr 1fr 1.4fr" : state.meta.showGroupNumberFields ? "1fr 1fr" : "1fr";

  return `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${escapeHtml(getMetaValue("examTitle", "テスト名"))} ${escapeHtml(label)} ${escapeHtml(bundle.label)}</title>
      <style>
        :root { color-scheme: light; }
        * { box-sizing: border-box; }
        body {
          margin: 0;
          font-family: "Yu Gothic", "Hiragino Sans", sans-serif;
          background: #eef1f4;
          color: #111;
        }
        .toolbar {
          position: sticky;
          top: 0;
          z-index: 10;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 16px;
          padding: 16px 20px;
          background: #24343b;
          color: #fff;
        }
        .toolbar-title { font-size: 1rem; font-weight: 700; }
        .toolbar-note { font-size: 0.88rem; opacity: 0.88; }
        .toolbar-button {
          border: none;
          border-radius: 999px;
          padding: 12px 18px;
          background: #fff;
          color: #24343b;
          font-size: 0.95rem;
          font-weight: 700;
          cursor: pointer;
        }
        .sheet {
          width: ${paperConfig.widthMm}mm;
          min-height: ${paperConfig.heightMm}mm;
          margin: 18px auto;
          padding: 14mm 14mm 16mm;
          background: #fff;
          box-shadow: 0 12px 30px rgba(0, 0, 0, 0.08);
        }
        .sheet[data-layout-mode="compact"] .print-question {
          padding-bottom: 5mm;
          margin-bottom: 5mm;
        }
        .sheet[data-layout-mode="spacious"] .print-question {
          padding-bottom: 9mm;
          margin-bottom: 9mm;
        }
        .print-header {
          border-bottom: 2px solid #222;
          padding-bottom: 10mm;
          margin-bottom: 8mm;
        }
        .print-header-top, .print-summary {
          display: flex;
          flex-wrap: wrap;
          gap: 12px 18px;
          font-size: 0.94rem;
        }
        .print-header-title {
          margin: 8mm 0 3mm;
          text-align: center;
          font-size: 1.8rem;
          font-weight: 700;
          letter-spacing: 0.04em;
        }
        .print-header-subtitle {
          text-align: center;
          font-size: 1rem;
          font-weight: 700;
          margin-bottom: 6mm;
        }
        .print-meta-grid {
          display: grid;
          grid-template-columns: ${printMetaColumns};
          gap: 12px;
          margin: 6mm 0;
        }
        .meta-box {
          border: 1px solid #333;
          padding: 10px 12px;
          min-height: 16mm;
        }
        .meta-box-label {
          font-size: 0.82rem;
          margin-bottom: 6px;
          font-weight: 700;
        }
        .meta-box-line {
          border-bottom: 1px solid #222;
          height: 18px;
        }
        .wide { grid-column: auto; }
        .print-instructions {
          margin-top: 4mm;
          padding: 10px 12px;
          border: 1px solid #444;
          background: #fafafa;
          line-height: 1.8;
          font-size: 0.92rem;
          white-space: pre-wrap;
          overflow-wrap: anywhere;
          word-break: break-word;
        }
        .print-question {
          padding: 0 0 7mm;
          margin-bottom: 7mm;
          border-bottom: 1px solid #bbb;
          page-break-inside: avoid;
          break-inside: avoid;
        }
        .print-question:last-child {
          margin-bottom: 0;
          border-bottom: none;
        }
        .print-question-top {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 12px;
          margin-bottom: 4px;
        }
        .print-question-title {
          font-size: 1.08rem;
          font-weight: 700;
        }
        .print-question-points {
          min-width: 68px;
          text-align: right;
          font-size: 0.98rem;
          font-weight: 700;
        }
        .print-badges {
          display: flex;
          flex-wrap: wrap;
          gap: 8px;
          margin: 6px 0 10px;
        }
        .print-badges span {
          border: 1px solid #666;
          padding: 3px 8px;
          font-size: 0.8rem;
        }
        .print-prompt, .teacher-answer-value, .teacher-note {
          margin: 0 0 10px;
          line-height: 1.9;
          font-size: 1rem;
          white-space: pre-wrap;
          overflow-wrap: anywhere;
          word-break: break-word;
        }
        .print-choices {
          margin: 0 0 10px 1.4em;
          padding: 0;
          line-height: 1.9;
          white-space: pre-wrap;
          overflow-wrap: anywhere;
        }
        .response-lines {
          display: grid;
          gap: 12px;
          margin-top: 12px;
        }
        .write-line {
          border-bottom: 1px solid #222;
          min-height: 28px;
        }
        .teacher-answer {
          margin-top: 12px;
          border: 1px solid #222;
          background: #f8f8f8;
          padding: 10px 12px;
        }
        .teacher-answer-label {
          font-size: 0.82rem;
          font-weight: 700;
          margin-bottom: 6px;
        }
        .teacher-note strong {
          display: inline-block;
          margin-right: 6px;
        }
        .footer-total {
          margin-top: 8mm;
          text-align: right;
          font-size: 1rem;
          font-weight: 700;
        }
        @page {
          size: ${paperConfig.widthMm}mm ${paperConfig.heightMm}mm;
          margin: 12mm;
        }
        @media print {
          body { background: #fff; }
          .toolbar { display: none; }
          .sheet {
            width: auto;
            min-height: auto;
            margin: 0;
            padding: 0;
            box-shadow: none;
          }
        }
      </style>
    </head>
    <body>
      <div class="toolbar">
        <div>
          <div class="toolbar-title">${escapeHtml(bundle.label)}の${escapeHtml(label)}を開きました</div>
          <div class="toolbar-note">上のボタンから「印刷またはPDF保存」を押してください。用紙サイズは ${escapeHtml(paperConfig.label)} です。</div>
        </div>
        <button class="toolbar-button" onclick="window.print()">印刷またはPDF保存</button>
      </div>

      <main class="sheet" data-layout-mode="${escapeAttribute(state.meta.layoutMode || "standard")}">
        ${buildPrintableHeader(label, totalPointsValue, bundle)}
        ${questionsHtml}
        ${mode === "teacher" ? `<div class="footer-total">合計 ${totalPointsValue}点</div>` : ""}
      </main>

      <script>
        window.addEventListener("load", () => {
          setTimeout(() => window.print(), 300);
        });
      <\/script>
    </body>
    </html>
  `;
}

function openPrintableDocument(mode, bundle = getActiveBundle()) {
  const label = mode === "teacher"
    ? "解答付きPDF保存・印刷"
    : mode === "answeronly"
      ? "解答用紙だけ印刷"
      : "生徒用PDF保存・印刷";

  syncMetaFromInputs();

  if (!ensureReadyForOutput(label, bundle.questions)) {
    return;
  }

  const exportWindow = window.open("", "_blank");
  if (!exportWindow) {
    window.alert("新しい画面を開けませんでした。ポップアップブロックを一時的に解除して、もう一度お試しください。");
    return;
  }

  exportWindow.document.open();
  exportWindow.document.write(buildPrintableHtml(mode, bundle));
  exportWindow.document.close();

  setEditorMessage(`${bundle.label}の${label}を開きました。上のボタンから印刷またはPDF保存に進めます。`);
}

function csvEscape(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function exportCsv(bundle = getActiveBundle()) {
  syncMetaFromInputs();

  if (!ensureReadyForOutput("CSV出力", bundle.questions)) {
    return;
  }

  const rows = [
    ["問題番号", "問題文", "問題形式", "選択肢", "正解", "解説", "配点", "難易度", "単元タグ"],
    ...bundle.questions.map((question, index) => [
      index + 1,
      question.prompt,
      questionTypes[question.type].label,
      getChoiceItems(question).join("\n"),
      question.answer,
      question.explanation,
      getQuestionPointsValue(question) ?? "",
      difficultyLabels[question.difficulty],
      question.unitTag
    ])
  ];

  const csv = `\uFEFF${rows.map((row) => row.map(csvEscape).join(",")).join("\r\n")}`;
  downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8;" }), buildExportFilename("questions", "csv", bundle));
  setEditorMessage(`${bundle.label}のCSVを出力しました。Excelで開きやすいUTF-8 BOM付きです。`);
}

function exportJson() {
  syncMetaFromInputs();
  downloadBlob(
    new Blob([JSON.stringify(buildBackupPayload(), null, 2)], { type: "application/json" }),
    buildExportFilename("backup", "json", createVariantBundle("base"))
  );
  setEditorMessage("バックアップJSONを出力しました。復元用として保管できます。");
}

function normalizeLoadedMeta(meta = {}, questions = []) {
  const templateId = testTemplates[meta.templateId] ? meta.templateId : "unit";
  const template = getTemplateConfig(templateId);

  return {
    templateId,
    schoolName: String(meta.schoolName ?? createDefaultMeta().schoolName),
    examTitle: String(meta.examTitle ?? createDefaultMeta().examTitle),
    subject: String(meta.subject ?? createDefaultMeta().subject),
    grade: String(meta.grade ?? ""),
    unit: String(meta.unit ?? ""),
    className: String(meta.className ?? meta.gradeClass ?? ""),
    durationMinutes: String(meta.durationMinutes ?? parseDurationValue(meta.duration) ?? template.durationMinutes),
    maxScore: String(meta.maxScore ?? inferMaxScore(questions)),
    examDate: normalizeDateValue(meta.examDate),
    paperSize: paperSizes[meta.paperSize] ? meta.paperSize : template.paperSize,
    nameLabel: String(meta.nameLabel ?? "名前"),
    groupLabel: String(meta.groupLabel ?? "組"),
    numberLabel: String(meta.numberLabel ?? "番号"),
    instructions: String(meta.instructions ?? defaultInstructions),
    showNameField: meta.showNameField === undefined ? true : Boolean(meta.showNameField),
    showGroupNumberFields: meta.showGroupNumberFields === undefined ? true : Boolean(meta.showGroupNumberFields),
    layoutMode: ["compact", "standard", "spacious"].includes(meta.layoutMode) ? meta.layoutMode : template.layoutMode
  };
}

function applyLoadedData(data = {}) {
  const source = data && typeof data === "object" ? data : {};
  const meta = source.meta && typeof source.meta === "object" ? source.meta : {};
  const questions = Array.isArray(source.questions) ? source.questions : [];

  state.meta = normalizeLoadedMeta(meta, questions);
  state.questions = questions.length ? questions.map(normalizeQuestion) : [createEmptyQuestion("multiple")];
  state.bankFilters = createDefaultBankFilters();
  state.patterns = createEmptyPatterns();
  uiState.previewVariant = "base";

  populateFieldsFromState();
  renderAll();
  persistDraft();
}

function importJson(event) {
  const [file] = event.target.files;
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      if (!data || typeof data !== "object") {
        throw new Error("Invalid JSON payload");
      }
      applyLoadedData(data);
      setEditorMessage("バックアップJSONを読み込みました。");
    } catch (error) {
      console.warn(error);
      window.alert("バックアップJSONの読み込みに失敗しました。ファイルが壊れていないか確認してください。");
      setEditorMessage("バックアップJSONを読み込めませんでした。");
    }
  };
  reader.readAsText(file, "utf-8");
  event.target.value = "";
}

function moveItem(array, index, direction) {
  const nextIndex = index + direction;
  if (nextIndex < 0 || nextIndex >= array.length) {
    return;
  }
  const [item] = array.splice(index, 1);
  array.splice(nextIndex, 0, item);
}

function updateQuestionField(index, field, value) {
  const question = state.questions[index];
  if (!question) {
    return;
  }

  if (field === "type") {
    const nextType = questionTypes[value] ? value : question.type;
    question.type = nextType;
    question.answerSpace = getDefaultAnswerSpace(nextType);
    if (questionTypes[nextType].usesChoices) {
      if (!question.choices.trim()) {
        question.choices = questionTypes[nextType].sampleChoices;
      }
    } else {
      question.choices = "";
    }
    return;
  }

  question[field] = value;
}

function saveQuestionToBank(question, messagePrefix = "") {
  const normalizedQuestion = cloneQuestion(question, { id: createId() });
  const fingerprint = buildQuestionFingerprint(normalizedQuestion);
  const existingIndex = state.bank.findIndex((entry) => entry.fingerprint === fingerprint);
  const now = new Date().toISOString();

  if (existingIndex >= 0) {
    state.bank[existingIndex] = {
      ...state.bank[existingIndex],
      updatedAt: now,
      question: normalizedQuestion
    };
    persistBank();
    renderBankList();
    setEditorMessage(`${messagePrefix}同じ問題があったため、問題バンクの内容を上書きしました。`);
    return;
  }

  state.bank.unshift({
    id: createId(),
    fingerprint,
    favorite: false,
    createdAt: now,
    updatedAt: now,
    question: normalizedQuestion
  });
  persistBank();
  renderBankList();
  setEditorMessage(`${messagePrefix}問題を問題バンクに保存しました。`);
}

function saveAllQuestionsToBank() {
  if (!state.questions.length) {
    window.alert("保存する問題がありません。");
    return;
  }
  state.questions.forEach((question) => saveQuestionToBank(question));
  setEditorMessage(`今の問題 ${state.questions.length} 問を問題バンクへ確認しながら保存しました。`);
}

function reuseBankQuestion(bankId) {
  const entry = state.bank.find((item) => item.id === bankId);
  if (!entry) {
    return;
  }
  state.questions.push(cloneQuestion(entry.question, { id: createId() }));
  invalidatePatterns();
  renderAll();
  persistDraft();
  setEditorMessage("問題バンクから問題を追加しました。");
}

function toggleBankFavorite(bankId) {
  const entry = state.bank.find((item) => item.id === bankId);
  if (!entry) {
    return;
  }
  entry.favorite = !entry.favorite;
  entry.updatedAt = new Date().toISOString();
  persistBank();
  renderBankList();
  setEditorMessage(entry.favorite ? "問題をお気に入りに入れました。" : "お気に入りを外しました。");
}

function deleteBankQuestion(bankId) {
  const entry = state.bank.find((item) => item.id === bankId);
  if (!entry) {
    return;
  }
  const confirmed = window.confirm("この問題を問題バンクから削除しますか？");
  if (!confirmed) {
    return;
  }
  state.bank = state.bank.filter((item) => item.id !== bankId);
  persistBank();
  renderBankList();
  setEditorMessage("問題バンクから削除しました。");
}

function resetBank() {
  if (!state.bank.length) {
    window.alert("問題バンクはすでに空です。");
    return;
  }
  const confirmed = window.confirm("問題バンクを初期化しますか？保存済みの問題はすべて消えます。");
  if (!confirmed) {
    return;
  }
  state.bank = [];
  persistBank();
  renderBankList();
  setEditorMessage("問題バンクを初期化しました。");
}

function getAnswerTokenChoice(token, choices) {
  const trimmed = String(token ?? "").trim();
  if (!trimmed) {
    return "";
  }

  if (/^\d+$/.test(trimmed)) {
    const index = Number(trimmed) - 1;
    if (choices[index]) {
      return choices[index];
    }
  }

  if (/^[A-Za-z]$/.test(trimmed)) {
    const index = trimmed.toUpperCase().charCodeAt(0) - 65;
    if (choices[index]) {
      return choices[index];
    }
  }

  const exact = choices.find((choice) => normalizeSpace(choice) === normalizeSpace(trimmed));
  return exact || trimmed;
}

function normalizeChoiceBasedAnswer(question, choices) {
  const answer = String(question.answer || "").trim();
  if (!answer) {
    return "";
  }

  if (question.type === "multiple") {
    return getAnswerTokenChoice(answer, choices);
  }

  if (question.type === "order") {
    const looksLikeTokens = /^[0-9A-Za-z\s,、>\-→]+$/.test(answer);
    if (!looksLikeTokens) {
      return answer;
    }

    const tokens = answer
      .split(/[\s,、>\-→]+/)
      .map((item) => item.trim())
      .filter(Boolean);

    if (!tokens.length) {
      return answer;
    }

    return tokens.map((token) => getAnswerTokenChoice(token, choices)).join(" → ");
  }

  return answer;
}

function shuffleQuestionForPattern(question, random, shuffleChoices) {
  const nextQuestion = cloneQuestion(question, { id: createId() });
  const originalChoices = getChoiceItems(nextQuestion);
  nextQuestion.answer = normalizeChoiceBasedAnswer(nextQuestion, originalChoices);

  if (!shuffleChoices || !originalChoices.length) {
    return nextQuestion;
  }

  const shuffledChoices = shuffleArray(originalChoices, random);
  nextQuestion.choices = shuffledChoices.join("\n");
  return nextQuestion;
}

function createPatternQuestions(baseQuestions, variantKey) {
  const random = createSeededRandom(`${variantKey}:${JSON.stringify(baseQuestions)}:${state.patterns.shuffleQuestions}:${state.patterns.shuffleChoices}`);
  let questions = baseQuestions.map((question) => shuffleQuestionForPattern(question, random, state.patterns.shuffleChoices));

  if (state.patterns.shuffleQuestions) {
    questions = shuffleArray(questions, random);
  }

  return questions.map((question) => cloneQuestion(question, { id: createId() }));
}

function generatePatterns() {
  syncMetaFromInputs();
  state.patterns.shuffleQuestions = document.getElementById("shuffleQuestionOrder").checked;
  state.patterns.shuffleChoices = document.getElementById("shuffleChoiceOrder").checked;

  if (!ensureReadyForOutput("A/B/Cパターン作成", state.questions)) {
    return;
  }

  const generated = {};
  ["A", "B", "C"].forEach((key) => {
    const questions = createPatternQuestions(state.questions, key);
    generated[key] = {
      questions,
      totalPoints: questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0)
    };
  });

  state.patterns.generated = generated;
  uiState.previewVariant = "A";
  renderDerivedViews();
  setPatternMessage(`A/B/Cパターンを作りました。${describePatternOptions()} です。`);
  setEditorMessage("A/B/Cパターンを作成しました。中央表示や印刷に切り替えられます。");
}

function getPatternBundle(patternKey) {
  const generated = state.patterns.generated[patternKey];
  if (!generated) {
    return null;
  }
  return {
    key: patternKey,
    label: `${patternKey}パターン`,
    questions: generated.questions.map((question) => cloneQuestion(question)),
    isPattern: true
  };
}

function getTargetScoreForHelper(fallback) {
  const parsed = parsePositiveIntStrict(state.meta.maxScore);
  return parsed.valid ? parsed.value : fallback;
}

function assignPointsByWeights(weights, targetTotal) {
  if (!weights.length) {
    return [];
  }
  if (targetTotal < weights.length) {
    return null;
  }

  const basePoints = Array(weights.length).fill(1);
  const remaining = targetTotal - weights.length;
  const totalWeight = weights.reduce((sum, weight) => sum + weight, 0);

  if (remaining === 0) {
    return basePoints;
  }

  const scaled = weights.map((weight) => (weight / totalWeight) * remaining);
  const increments = scaled.map((value) => Math.floor(value));
  let distributed = increments.reduce((sum, value) => sum + value, 0);
  const fractions = scaled
    .map((value, index) => ({ index, fraction: value - increments[index] }))
    .sort((left, right) => right.fraction - left.fraction);

  while (distributed < remaining) {
    const target = fractions[distributed % fractions.length];
    increments[target.index] += 1;
    distributed += 1;
  }

  return basePoints.map((value, index) => value + increments[index]);
}

function applyAssignedPoints(pointsArray, nextMaxScore = null, message = "") {
  pointsArray.forEach((points, index) => {
    if (state.questions[index]) {
      state.questions[index].points = String(points);
    }
  });

  if (nextMaxScore !== null) {
    state.meta.maxScore = String(nextMaxScore);
    document.getElementById("maxScore").value = String(nextMaxScore);
  }

  invalidatePatterns();
  renderAll();
  persistDraft();
  if (message) {
    setEditorMessage(message);
  }
}

function autoAdjustPointsTo(targetTotal) {
  if (!state.questions.length) {
    window.alert("配点を調整する前に問題を追加してください。");
    return;
  }

  const currentWeights = state.questions.map((question) => getQuestionPointsValue(question) ?? 1);
  const assigned = assignPointsByWeights(currentWeights, targetTotal);
  if (!assigned) {
    window.alert(`問題数が${state.questions.length}問あるため、${targetTotal}点に整数で調整できません。`);
    return;
  }

  applyAssignedPoints(assigned, targetTotal, `配点を${targetTotal}点に自動調整しました。`);
}

function equalizePoints() {
  if (!state.questions.length) {
    window.alert("配点をそろえる前に問題を追加してください。");
    return;
  }

  const targetTotal = getTargetScoreForHelper(state.questions.length * 10);
  const weights = state.questions.map(() => 1);
  const assigned = assignPointsByWeights(weights, targetTotal);
  if (!assigned) {
    window.alert("現在の満点では、全問を1点以上の同じ配点にできません。");
    return;
  }

  applyAssignedPoints(assigned, null, "全問の配点をできるだけ同じになるように調整しました。");
}

function boostShortAnswerPoints() {
  if (!state.questions.length) {
    window.alert("配点を調整する前に問題を追加してください。");
    return;
  }

  const hasShort = state.questions.some((question) => question.type === "short");
  if (!hasShort) {
    window.alert("記述問題がないため、この調整は使えません。");
    return;
  }

  const targetTotal = getTargetScoreForHelper(
    state.questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0) || 50
  );
  const weights = state.questions.map((question) => (question.type === "short" ? 2.2 : 1));
  const assigned = assignPointsByWeights(weights, targetTotal);
  if (!assigned) {
    window.alert("現在の満点では、記述問題だけを高配点にできません。");
    return;
  }

  applyAssignedPoints(assigned, null, "記述問題をやや高配点に調整しました。");
}

function applyQuestionChange(message) {
  invalidatePatterns();
  renderAll();
  persistDraft();
  if (message) {
    setEditorMessage(message);
  }
}

function loadDraft() {
  let loaded = false;
  for (const key of [STORAGE_KEY, ...LEGACY_STORAGE_KEYS]) {
    const raw = localStorage.getItem(key);
    if (!raw) {
      continue;
    }

    const data = parseStoredJson(raw, () => {
      if (key === STORAGE_KEY) {
        window.alert("自動保存データの読み込みに失敗したため、新しい状態で開きました。");
      }
    });

    if (!data) {
      continue;
    }

    applyLoadedData(data);
    loaded = true;
    if (key !== STORAGE_KEY) {
      setEditorMessage("前の保存形式を読み込み、第2段階の形に変換しました。");
    } else {
      setEditorMessage("前回の入力内容を復元しました。");
    }
    break;
  }

  if (!loaded) {
    if (!state.questions.length) {
      state.questions = [createEmptyQuestion("multiple")];
    }
    populateFieldsFromState();
    renderAll();
    persistDraft();
  }
}

function handleQuestionCardClick(event) {
  const button = event.target.closest("button[data-action]");
  if (!button) {
    return;
  }

  const card = button.closest(".question-card");
  const index = Number(card?.dataset.index);
  if (!Number.isInteger(index)) {
    return;
  }

  if (button.dataset.action === "move-up") {
    moveItem(state.questions, index, -1);
    applyQuestionChange(`第${index + 1}問を上へ移動しました。`);
    return;
  }

  if (button.dataset.action === "move-down") {
    moveItem(state.questions, index, 1);
    applyQuestionChange(`第${index + 1}問を下へ移動しました。`);
    return;
  }

  if (button.dataset.action === "duplicate") {
    state.questions.splice(index + 1, 0, cloneQuestion(state.questions[index], { id: createId() }));
    applyQuestionChange(`第${index + 1}問を複製しました。`);
    return;
  }

  if (button.dataset.action === "save-bank") {
    saveQuestionToBank(state.questions[index], `第${index + 1}問を`);
    return;
  }

  if (button.dataset.action === "delete") {
    state.questions.splice(index, 1);
    applyQuestionChange(`第${index + 1}問を削除しました。`);
  }
}

function handleQuestionFieldUpdate(event) {
  const field = event.target.dataset.field;
  const card = event.target.closest(".question-card");
  if (!field || !card) {
    return;
  }
  const index = Number(card.dataset.index);
  if (!Number.isInteger(index)) {
    return;
  }
  updateQuestionField(index, field, event.target.value);
}

function registerEventListeners() {
  questionList.addEventListener("click", handleQuestionCardClick);

  questionList.addEventListener("input", (event) => {
    handleQuestionFieldUpdate(event);
    invalidatePatterns();
    renderDerivedViews();
    persistDraft();
  });

  questionList.addEventListener("change", (event) => {
    handleQuestionFieldUpdate(event);
    invalidatePatterns();
    renderAll();
    persistDraft();
  });

  textMetaFieldIds.forEach((id) => {
    const element = document.getElementById(id);
    const eventName = element.tagName === "SELECT" ? "change" : "input";
    element.addEventListener(eventName, () => {
      syncMetaFromInputs();
      if (id === "templateId") {
        applyTemplateConfig(element.value, "テンプレートを変更しました。入力済みの問題はそのまま残しています。");
      }
      renderDerivedViews();
      persistDraft();
    });
  });

  checkboxMetaFieldIds.forEach((id) => {
    const element = document.getElementById(id);
    element.addEventListener("change", () => {
      syncMetaFromInputs();
      renderDerivedViews();
      persistDraft();
    });
  });

  bankFilterFieldIds.forEach((id) => {
    const element = document.getElementById(id);
    const eventName = element.type === "checkbox" ? "change" : "input";
    element.addEventListener(eventName, () => {
      updateBankFiltersFromInputs();
      renderBankList();
    });
  });

  document.querySelectorAll("[data-add-type]").forEach((button) => {
    button.addEventListener("click", () => {
      state.questions.push(createEmptyQuestion(button.dataset.addType));
      applyQuestionChange(`${questionTypes[button.dataset.addType].label}を追加しました。`);
    });
  });

  document.getElementById("saveAllToBank").addEventListener("click", saveAllQuestionsToBank);

  document.getElementById("clearQuestions").addEventListener("click", () => {
    if (!state.questions.length) {
      window.alert("消す問題がありません。");
      return;
    }
    const confirmed = window.confirm("今の問題をすべて消しますか？");
    if (!confirmed) {
      return;
    }
    state.questions = [];
    applyQuestionChange("今の問題をすべて消しました。");
  });

  document.getElementById("autoScore100").addEventListener("click", () => autoAdjustPointsTo(100));
  document.getElementById("autoScore50").addEventListener("click", () => autoAdjustPointsTo(50));
  document.getElementById("equalizePoints").addEventListener("click", equalizePoints);
  document.getElementById("boostShortAnswer").addEventListener("click", boostShortAnswerPoints);

  document.getElementById("exportStudentPdf").addEventListener("click", () => openPrintableDocument("student"));
  document.getElementById("exportTeacherPdf").addEventListener("click", () => openPrintableDocument("teacher"));
  document.getElementById("exportAnswerSheet").addEventListener("click", () => openPrintableDocument("answeronly"));
  document.getElementById("exportCsv").addEventListener("click", () => exportCsv());
  document.getElementById("exportJson").addEventListener("click", exportJson);
  document.getElementById("importJson").addEventListener("change", importJson);

  document.getElementById("generatePatterns").addEventListener("click", generatePatterns);
  document.getElementById("clearPatterns").addEventListener("click", () => {
    resetPatterns();
    renderDerivedViews();
    setEditorMessage("A/B/Cパターンを消しました。");
  });

  document.getElementById("shuffleQuestionOrder").addEventListener("change", (event) => {
    state.patterns.shuffleQuestions = event.target.checked;
  });

  document.getElementById("shuffleChoiceOrder").addEventListener("change", (event) => {
    state.patterns.shuffleChoices = event.target.checked;
  });

  printPreviewButton.addEventListener("click", () => {
    if (uiState.activeTab === "answer") {
      openPrintableDocument("teacher");
      return;
    }
    if (uiState.activeTab === "answeronly") {
      openPrintableDocument("answeronly");
      return;
    }
    openPrintableDocument("student");
  });

  document.querySelectorAll(".tab-button").forEach((button) => {
    button.addEventListener("click", () => switchTab(button.dataset.tab));
  });

  previewVariantButtons.forEach((button) => {
    button.addEventListener("click", () => switchPreviewVariant(button.dataset.variant));
  });

  bankList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-bank-action]");
    if (!button) {
      return;
    }
    const card = button.closest("[data-bank-id]");
    const bankId = card?.dataset.bankId;
    if (!bankId) {
      return;
    }
    if (button.dataset.bankAction === "reuse") {
      reuseBankQuestion(bankId);
      return;
    }
    if (button.dataset.bankAction === "favorite") {
      toggleBankFavorite(bankId);
      return;
    }
    if (button.dataset.bankAction === "delete") {
      deleteBankQuestion(bankId);
    }
  });

  document.getElementById("resetBank").addEventListener("click", resetBank);

  patternList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-pattern-action]");
    if (!button) {
      return;
    }
    const card = button.closest("[data-pattern-key]");
    const patternKey = card?.dataset.patternKey;
    if (!patternKey) {
      return;
    }
    const bundle = getPatternBundle(patternKey);
    if (!bundle) {
      return;
    }
    if (button.dataset.patternAction === "show") {
      switchPreviewVariant(patternKey);
      return;
    }
    if (button.dataset.patternAction === "student") {
      openPrintableDocument("student", bundle);
      return;
    }
    if (button.dataset.patternAction === "teacher") {
      openPrintableDocument("teacher", bundle);
      return;
    }
    if (button.dataset.patternAction === "answeronly") {
      openPrintableDocument("answeronly", bundle);
      return;
    }
    if (button.dataset.patternAction === "csv") {
      exportCsv(bundle);
    }
  });
}

function init() {
  state.questions = [createEmptyQuestion("multiple")];
  loadQuestionBank();
  registerEventListeners();
  populateFieldsFromState();
  applyTemplateConfig(state.meta.templateId);
  loadDraft();
  updateBankFiltersFromInputs();
  renderDerivedViews();
}

init();

function getQuestionPointsLabel(question) {
  const value = getQuestionPointsValue(question);
  return value === null ? "配点未設定" : `${value}点`;
}

function getAnswerLinesCount(question, layout = null) {
  const preset = answerSpaceOptions[question.answerSpace] || answerSpaceOptions.medium;
  if (question.type === "multiple" || question.type === "truefalse") {
    return 1;
  }
  if (!layout) {
    return preset.lines;
  }
  const questionLayout = getQuestionLayout(question, layout);
  const lineHeightMm = Math.max(6.2, (questionLayout.questionFontSize * layout.lineHeight * 0.264583) + 1.8);
  const numericLines = Math.max(1, Math.round(questionLayout.answerAreaHeightMm / lineHeightMm));
  return Math.max(preset.lines, numericLines);
}

function formatDateLabel(value) {
  if (!value) {
    return "未設定";
  }
  const [year, month, day] = String(value).split("-");
  if (!year || !month || !day) {
    return value;
  }
  return `${Number(year)}年${Number(month)}月${Number(day)}日`;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function formatMeasure(value) {
  if (!Number.isFinite(value)) {
    return "";
  }
  return Number.isInteger(value) ? String(value) : String(Math.round(value * 10) / 10);
}

function parsePositiveNumberStrict(value) {
  const text = String(value ?? "").trim();
  if (!text || !/^\d+(\.\d+)?$/.test(text)) {
    return { valid: false, value: null };
  }
  const number = Number(text);
  if (!Number.isFinite(number) || number <= 0) {
    return { valid: false, value: null };
  }
  return { valid: true, value: number };
}

function parseNumberWithFallback(value, fallback, min, max) {
  const parsed = parsePositiveNumberStrict(value);
  if (!parsed.valid) {
    return clamp(fallback, min, max);
  }
  return clamp(parsed.value, min, max);
}

function getLayoutModeConfig(modeId = state.meta.layoutMode) {
  const layoutModes = {
    standard: { label: "標準", fontOffset: 0, gapOffset: 0, lineHeightOffset: 0, sideMarginOffset: 0, topMarginOffset: 0, answerScale: 1 },
    compact: { label: "詰め込み", fontOffset: -1, gapOffset: -2, lineHeightOffset: -0.12, sideMarginOffset: -2, topMarginOffset: -2, answerScale: 0.85 },
    spacious: { label: "ゆったり", fontOffset: 1, gapOffset: 3, lineHeightOffset: 0.12, sideMarginOffset: 1, topMarginOffset: 1, answerScale: 1.2 },
    largeText: { label: "大きめ文字", fontOffset: 3, gapOffset: 1, lineHeightOffset: 0.15, sideMarginOffset: 0, topMarginOffset: 0, answerScale: 1.15 },
    smallText: { label: "小さめ文字", fontOffset: -2, gapOffset: -1, lineHeightOffset: -0.1, sideMarginOffset: -1, topMarginOffset: -1, answerScale: 0.9 },
    wideAnswers: { label: "解答欄広め", fontOffset: 0, gapOffset: 1, lineHeightOffset: 0.05, sideMarginOffset: 0, topMarginOffset: 0, answerScale: 1.45 },
    onePage: { label: "1ページに収める", fontOffset: -1, gapOffset: -2, lineHeightOffset: -0.18, sideMarginOffset: -2, topMarginOffset: -2, answerScale: 0.78 }
  };
  return layoutModes[modeId] || layoutModes.standard;
}

function getPaperConfig(meta = state.meta) {
  const paperSize = paperSizes[meta.paperSize] ? meta.paperSize : "A4";
  const base = paperSizes[paperSize] || paperSizes.A4;
  const widthCheck = paperSize === "Custom" ? parsePositiveNumberStrict(meta.customPaperWidthMm) : { valid: true, value: base.widthMm };
  const heightCheck = paperSize === "Custom" ? parsePositiveNumberStrict(meta.customPaperHeightMm) : { valid: true, value: base.heightMm };
  const tilesXCheck = parsePositiveIntStrict(meta.paperTilesX);
  const tilesYCheck = parsePositiveIntStrict(meta.paperTilesY);
  const tilesX = tilesXCheck.valid ? tilesXCheck.value : 1;
  const tilesY = tilesYCheck.valid ? tilesYCheck.value : 1;
  const baseWidthMm = widthCheck.valid ? widthCheck.value : Number.NaN;
  const baseHeightMm = heightCheck.valid ? heightCheck.value : Number.NaN;
  const widthMm = Number.isFinite(baseWidthMm) ? baseWidthMm * tilesX : Number.NaN;
  const heightMm = Number.isFinite(baseHeightMm) ? baseHeightMm * tilesY : Number.NaN;
  const previewWidthPx = clamp(Math.round((Number.isFinite(widthMm) ? widthMm : 210) * 3.1), 380, 920);
  const previewHeightPx = Math.round(previewWidthPx * ((Number.isFinite(heightMm) ? heightMm : 297) / (Number.isFinite(widthMm) ? widthMm : 210)));
  const tileLabel = tilesX === 1 && tilesY === 1
    ? "1枚分"
    : tilesX > 1 && tilesY > 1
      ? `${tilesX}×${tilesY}枚分`
      : tilesX > 1
        ? `横${tilesX}枚分`
        : `縦${tilesY}枚分`;
  const issues = [];

  if (paperSize === "Custom" && (!widthCheck.valid || !heightCheck.valid)) {
    issues.push({ severity: "error", message: "自由サイズを選んだときは、幅mmと高さmmを入力してください。" });
  }
  if (!tilesXCheck.valid) {
    issues.push({ severity: "error", message: "横に何枚分は1以上の数字で入力してください。" });
  }
  if (!tilesYCheck.valid) {
    issues.push({ severity: "error", message: "縦に何枚分は1以上の数字で入力してください。" });
  }

  return {
    key: paperSize,
    label: base.label,
    baseWidthMm,
    baseHeightMm,
    widthMm,
    heightMm,
    tilesX,
    tilesY,
    tileLabel,
    previewWidthPx,
    previewHeightPx,
    issues,
    paperBadge: Number.isFinite(widthMm) && Number.isFinite(heightMm)
      ? `${base.label} ${formatMeasure(widthMm)}×${formatMeasure(heightMm)}mm`
      : `${base.label} サイズ確認`,
    headerLabel: Number.isFinite(widthMm) && Number.isFinite(heightMm)
      ? `${base.label} / ${tileLabel} / ${formatMeasure(widthMm)}mm × ${formatMeasure(heightMm)}mm`
      : `${base.label} / サイズ未完成`
  };
}

function getEffectiveLayoutConfig(meta = state.meta) {
  const mode = getLayoutModeConfig(meta.layoutMode);
  return {
    modeId: meta.layoutMode,
    label: mode.label,
    globalFontSize: clamp(parseNumberWithFallback(meta.globalFontSize, 16, 8, 40) + mode.fontOffset, 8, 40),
    questionGap: clamp(parseNumberWithFallback(meta.questionGap, 7, 0, 40) + mode.gapOffset, 0, 40),
    lineHeight: clamp(parseNumberWithFallback(meta.lineHeight, 1.8, 1, 3) + mode.lineHeightOffset, 1, 3),
    sideMarginMm: clamp(parseNumberWithFallback(meta.sideMarginMm, 14, 4, 60) + mode.sideMarginOffset, 4, 60),
    topBottomMarginMm: clamp(parseNumberWithFallback(meta.topBottomMarginMm, 14, 4, 60) + mode.topMarginOffset, 4, 60),
    answerScale: mode.answerScale
  };
}

function getQuestionLayout(question, layout = getEffectiveLayoutConfig()) {
  const defaultAnswerHeight = Number(getDefaultAnswerAreaHeightMm(question.type));
  return {
    questionFontSize: clamp(parseNumberWithFallback(question.questionFontSize, layout.globalFontSize, 8, 40), 8, 40),
    choiceFontSize: clamp(parseNumberWithFallback(question.choiceFontSize, Math.max(8, layout.globalFontSize - 1), 8, 40), 8, 40),
    answerAreaHeightMm: clamp(parseNumberWithFallback(question.answerAreaHeightMm, defaultAnswerHeight, 6, 240) * layout.answerScale, 6, 240),
    blockPaddingMm: clamp(parseNumberWithFallback(question.blockPaddingMm, 4, 0, 30), 0, 30),
    blockWidthPercent: clamp(parseNumberWithFallback(question.blockWidthPercent, 100, 35, 100), 35, 100),
    figureSpaceHeightMm: clamp(parseNumberWithFallback(question.figureSpaceHeightMm, 0, 0, 180), 0, 180)
  };
}

function estimateBundleHeightMm(bundle, paperConfig = getPaperConfig(), layout = getEffectiveLayoutConfig()) {
  const headerBase = state.meta.showNameField || state.meta.showGroupNumberFields ? 64 : 48;
  const questionHeight = bundle.questions.reduce((sum, question) => {
    const questionLayout = getQuestionLayout(question, layout);
    const contentWidthFactor = questionLayout.blockWidthPercent / 100;
    const widthPenalty = 1 / Math.max(0.45, contentWidthFactor);
    const promptChars = normalizeSpace(question.prompt).length || 18;
    const promptLines = Math.max(1, Math.ceil((promptChars / 28) * widthPenalty));
    const promptHeight = promptLines * (questionLayout.questionFontSize * layout.lineHeight * 0.264583);
    const choiceCount = getChoiceItems(question).length;
    const choiceHeight = choiceCount ? choiceCount * Math.max(6, questionLayout.choiceFontSize * layout.lineHeight * 0.264583) : 0;
    return sum + promptHeight + choiceHeight + questionLayout.answerAreaHeightMm + questionLayout.figureSpaceHeightMm + (questionLayout.blockPaddingMm * 2) + layout.questionGap + 14;
  }, 0);
  return {
    contentHeight: headerBase + questionHeight,
    availableHeight: Number.isFinite(paperConfig.heightMm) ? Math.max(40, paperConfig.heightMm - (layout.topBottomMarginMm * 2)) : 0
  };
}

function updateCustomPaperVisibility() {
  const isCustom = state.meta.paperSize === "Custom";
  const widthWrap = document.getElementById("customPaperWidthWrap");
  const heightWrap = document.getElementById("customPaperHeightWrap");
  if (widthWrap) {
    widthWrap.hidden = !isCustom;
  }
  if (heightWrap) {
    heightWrap.hidden = !isCustom;
  }
}

function renderTemplateSummary() {
  const template = getTemplateConfig();
  const layout = getLayoutModeConfig(state.meta.layoutMode);
  templateSummary.textContent = `${template.label} / 満点 ${state.meta.maxScore}点 / ${state.meta.durationMinutes}分 / 問題数の目安 ${template.recommendedQuestionCount}問 / ${state.meta.paperSize} / ${layout.label}`;
}

function applyTemplateConfig(templateId, message = "") {
  const nextTemplateId = testTemplates[templateId] ? templateId : "unit";
  const template = getTemplateConfig(nextTemplateId);
  state.meta.templateId = nextTemplateId;
  state.meta.maxScore = template.maxScore;
  state.meta.durationMinutes = template.durationMinutes;
  state.meta.paperSize = template.paperSize;
  state.meta.layoutMode = template.layoutMode;
  state.meta.showNameField = template.showNameField;
  state.meta.showGroupNumberFields = template.showGroupNumberFields;
  state.meta.globalFontSize = template.globalFontSize;
  state.meta.questionGap = template.questionGap;
  state.meta.lineHeight = template.lineHeight;
  state.meta.sideMarginMm = template.sideMarginMm;
  state.meta.topBottomMarginMm = template.topBottomMarginMm;
  populateFieldsFromState();
  renderTemplateSummary();
  if (message) {
    setEditorMessage(message);
  }
}

function normalizeLoadedMeta(meta = {}, questions = []) {
  const templateId = testTemplates[meta.templateId] ? meta.templateId : "unit";
  const template = getTemplateConfig(templateId);
  const defaults = createDefaultMeta();
  return {
    templateId,
    schoolName: String(meta.schoolName ?? defaults.schoolName),
    examTitle: String(meta.examTitle ?? defaults.examTitle),
    subject: String(meta.subject ?? defaults.subject),
    grade: String(meta.grade ?? defaults.grade),
    unit: String(meta.unit ?? defaults.unit),
    className: String(meta.className ?? meta.gradeClass ?? defaults.className),
    durationMinutes: String(meta.durationMinutes ?? parseDurationValue(meta.duration) ?? template.durationMinutes),
    maxScore: String(meta.maxScore ?? inferMaxScore(questions)),
    examDate: normalizeDateValue(meta.examDate),
    paperSize: paperSizes[meta.paperSize] ? meta.paperSize : template.paperSize,
    customPaperWidthMm: String(meta.customPaperWidthMm ?? defaults.customPaperWidthMm),
    customPaperHeightMm: String(meta.customPaperHeightMm ?? defaults.customPaperHeightMm),
    paperTilesX: String(meta.paperTilesX ?? "1"),
    paperTilesY: String(meta.paperTilesY ?? "1"),
    nameLabel: String(meta.nameLabel ?? "名前"),
    groupLabel: String(meta.groupLabel ?? "組"),
    numberLabel: String(meta.numberLabel ?? "番号"),
    instructions: String(meta.instructions ?? defaultInstructions),
    showNameField: meta.showNameField === undefined ? true : Boolean(meta.showNameField),
    showGroupNumberFields: meta.showGroupNumberFields === undefined ? true : Boolean(meta.showGroupNumberFields),
    layoutMode: ["standard", "compact", "spacious", "largeText", "smallText", "wideAnswers", "onePage"].includes(meta.layoutMode)
      ? meta.layoutMode
      : template.layoutMode,
    globalFontSize: String(meta.globalFontSize ?? template.globalFontSize),
    questionGap: String(meta.questionGap ?? template.questionGap),
    lineHeight: String(meta.lineHeight ?? template.lineHeight),
    sideMarginMm: String(meta.sideMarginMm ?? template.sideMarginMm),
    topBottomMarginMm: String(meta.topBottomMarginMm ?? template.topBottomMarginMm)
  };
}

function updateQuestionField(index, field, value) {
  const question = state.questions[index];
  if (!question) {
    return;
  }

  if (field === "type") {
    const nextType = questionTypes[value] ? value : question.type;
    question.type = nextType;
    question.answerSpace = getDefaultAnswerSpace(nextType);
    question.answerAreaHeightMm = getDefaultAnswerAreaHeightMm(nextType);
    if (questionTypes[nextType].usesChoices) {
      if (!question.choices.trim()) {
        question.choices = questionTypes[nextType].sampleChoices;
      }
    } else {
      question.choices = "";
    }
    return;
  }

  question[field] = value;
}

function getValidationState(bundle = getActiveBundle()) {
  const errors = [];
  const warnings = [];
  const questionIssueMap = new Map();
  const paperConfig = getPaperConfig();
  const layout = getEffectiveLayoutConfig();

  const maxScoreCheck = parsePositiveIntStrict(state.meta.maxScore);
  const durationCheck = parsePositiveIntStrict(state.meta.durationMinutes);
  if (!maxScoreCheck.valid) {
    errors.push({ severity: "error", message: "満点は1以上の数字で入力してください。" });
  }
  if (!durationCheck.valid) {
    errors.push({ severity: "error", message: "制限時間は1以上の数字で入力してください。" });
  }
  paperConfig.issues.forEach((issue) => {
    (issue.severity === "error" ? errors : warnings).push(issue);
  });

  const totalPointsValue = state.questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0);
  if (!state.questions.length) {
    warnings.push({ severity: "warning", message: "問題がまだありません。右側の追加ボタンから問題を作ってください。" });
  }

  state.questions.forEach((question, index) => {
    const messages = [];
    const questionNumber = index + 1;
    if (!question.prompt.trim()) {
      messages.push(`第${questionNumber}問の問題文を入力してください。`);
    }
    if (!question.answer.trim()) {
      messages.push(`第${questionNumber}問の正解を入力してください。`);
    }
    if (!parsePositiveIntStrict(question.points).valid) {
      messages.push(`第${questionNumber}問の配点は1以上の数字で入力してください。`);
    }
    if ((question.type === "multiple" || question.type === "order") && getChoiceItems(question).length < 2) {
      messages.push(`第${questionNumber}問の${questionTypes[question.type].choiceLabel}は2つ以上必要です。`);
    }

    const layoutChecks = [
      { label: "問題の文字サイズ", value: question.questionFontSize, min: 8, max: 40 },
      { label: "選択肢の文字サイズ", value: question.choiceFontSize, min: 8, max: 40 },
      { label: "解答欄の高さ", value: question.answerAreaHeightMm, min: 6, max: 200 },
      { label: "問題ブロックの余白", value: question.blockPaddingMm, min: 0, max: 30 },
      { label: "問題の横幅", value: question.blockWidthPercent, min: 35, max: 100 },
      { label: "図や画像のスペース", value: question.figureSpaceHeightMm, min: 0, max: 180 }
    ];

    layoutChecks.forEach((check) => {
      const parsed = parsePositiveNumberStrict(check.value);
      const isZeroAllowed = check.min === 0 && String(check.value).trim() === "0";
      if (!parsed.valid && !isZeroAllowed) {
        messages.push(`第${questionNumber}問の${check.label}を数字で入力してください。`);
        return;
      }
      const numericValue = isZeroAllowed ? 0 : parsed.value;
      if (numericValue < check.min || numericValue > check.max) {
        messages.push(`第${questionNumber}問の${check.label}は${check.min}〜${check.max}の範囲で入力してください。`);
      }
    });

    if (messages.length) {
      questionIssueMap.set(index, messages);
      messages.forEach((message) => errors.push({ severity: "error", message }));
    }
  });

  if (maxScoreCheck.valid && state.questions.length && totalPointsValue !== maxScoreCheck.value) {
    warnings.push({ severity: "warning", message: `合計点は${totalPointsValue}点です。設定した満点の${maxScoreCheck.value}点と一致していません。` });
  }

  const estimated = estimateBundleHeightMm(bundle, paperConfig, layout);
  if (layout.modeId === "onePage" && estimated.availableHeight > 0 && estimated.contentHeight > estimated.availableHeight) {
    warnings.push({ severity: "warning", message: "1ページに収める設定ですが、今の内容では入りきらない可能性があります。" });
  }
  if (layout.modeId === "onePage" && layout.globalFontSize <= 10) {
    warnings.push({ severity: "warning", message: "1ページに収めるため文字がかなり小さくなっています。読みやすさも確認してください。" });
  }

  return {
    errors,
    warnings,
    questionIssueMap,
    totalPointsValue,
    maxScoreValue: maxScoreCheck.valid ? maxScoreCheck.value : null,
    durationValue: durationCheck.valid ? durationCheck.value : null
  };
}

function renderValidationItemsHtml(items, emptyLabel) {
  if (!items.length) {
    return `<div class="validation-item is-success">${escapeHtml(emptyLabel)}</div>`;
  }
  return items.map((item) => `
    <div class="validation-item ${item.severity === "error" ? "is-error" : "is-warning"}">
      ${escapeHtml(item.message)}
    </div>
  `).join("");
}

function renderStats(validation) {
  totalQuestions.textContent = String(state.questions.length);
  totalPoints.textContent = `${validation.totalPointsValue}点`;
  maxScoreDisplay.textContent = validation.maxScoreValue === null ? "未設定" : `${validation.maxScoreValue}点`;
  pointCheckStatus.classList.remove("point-check-ok", "point-check-warn");

  if (validation.errors.length) {
    pointCheckStatus.textContent = "要修正";
    pointCheckStatus.classList.add("point-check-warn");
    pointCheckMessage.textContent = `印刷前に直したい項目が ${validation.errors.length} 件あります。`;
  } else if (validation.warnings.length) {
    pointCheckStatus.textContent = "注意あり";
    pointCheckStatus.classList.add("point-check-warn");
    pointCheckMessage.textContent = validation.warnings[0].message;
  } else {
    pointCheckStatus.textContent = "このまま使えます";
    pointCheckStatus.classList.add("point-check-ok");
    pointCheckMessage.textContent = "配点と入力内容はそろっています。このまま印刷や保存に進めます。";
  }

  validationCount.textContent = `${validation.errors.length + validation.warnings.length}件`;
  validationList.innerHTML = renderValidationItemsHtml(
    [...validation.errors, ...validation.warnings],
    "気になる点はありません。"
  );
}

function renderQuestionCard(question, index, validation) {
  const config = questionTypes[question.type];
  const issues = validation.questionIssueMap.get(index) || [];
  const choiceHiddenClass = config.usesChoices ? "" : "is-hidden";

  return `
    <article class="question-card ${issues.length ? "is-invalid" : ""}" data-index="${index}">
      <div class="question-card-top">
        <div>
          <p class="mini-type">${escapeHtml(config.label)}</p>
          <h4>第${index + 1}問</h4>
          <div class="meta-badges">
            <span>${escapeHtml(getQuestionPointsLabel(question))}</span>
            <span>${escapeHtml(difficultyLabels[question.difficulty])}</span>
            <span>${escapeHtml(question.unitTag || "単元タグなし")}</span>
            ${issues.length ? "<span>要確認</span>" : ""}
          </div>
        </div>
        <div class="card-actions">
          <button type="button" class="small-button" data-action="move-up">上へ</button>
          <button type="button" class="small-button" data-action="move-down">下へ</button>
          <button type="button" class="small-button" data-action="duplicate">複製</button>
          <button type="button" class="small-button" data-action="save-bank">バンク保存</button>
          <button type="button" class="small-button" data-action="delete">削除</button>
        </div>
      </div>
      <div class="question-grid">
        <label>
          問題形式
          <select data-field="type">${renderTypeOptions(question.type)}</select>
        </label>
        <label>
          配点
          <input type="number" min="1" step="1" data-field="points" value="${escapeAttribute(question.points)}">
        </label>
        <label>
          難しさ
          <select data-field="difficulty">${renderDifficultyOptions(question.difficulty)}</select>
        </label>
        <label>
          単元タグ
          <input type="text" data-field="unitTag" value="${escapeAttribute(question.unitTag)}" placeholder="例: 歴史総合">
        </label>
        <label>
          線の目安
          <select data-field="answerSpace">${renderAnswerSpaceOptions(question.answerSpace)}</select>
        </label>
        <label>
          解答欄の高さ(mm)
          <input type="number" min="6" step="1" data-field="answerAreaHeightMm" value="${escapeAttribute(question.answerAreaHeightMm)}">
        </label>
        <label class="full-width">
          問題文
          <textarea rows="4" data-field="prompt" placeholder="${escapeAttribute(config.placeholder)}">${escapeHtml(question.prompt)}</textarea>
        </label>
        <label class="full-width choice-wrap ${choiceHiddenClass}">
          ${escapeHtml(config.choiceLabel)}
          <textarea rows="4" data-field="choices" placeholder="${escapeAttribute(config.choicePlaceholder)}">${escapeHtml(question.choices)}</textarea>
          <small class="field-tip">${escapeHtml(config.choiceHelp)}</small>
        </label>
        <label class="full-width">
          正解
          <textarea rows="2" data-field="answer" placeholder="${escapeAttribute(config.answerPlaceholder)}">${escapeHtml(question.answer)}</textarea>
        </label>
        <label class="full-width">
          解説
          <textarea rows="3" data-field="explanation" placeholder="採点や見直し用のメモを入力できます。">${escapeHtml(question.explanation)}</textarea>
        </label>
        <div class="meta-grid full-width question-layout-grid">
          <label>
            問題の文字サイズ
            <input type="number" min="8" max="40" step="1" data-field="questionFontSize" value="${escapeAttribute(question.questionFontSize)}">
          </label>
          <label>
            選択肢の文字サイズ
            <input type="number" min="8" max="40" step="1" data-field="choiceFontSize" value="${escapeAttribute(question.choiceFontSize)}">
          </label>
          <label>
            問題ブロックの余白(mm)
            <input type="number" min="0" max="30" step="1" data-field="blockPaddingMm" value="${escapeAttribute(question.blockPaddingMm)}">
          </label>
          <label>
            問題の横幅(%)
            <input type="number" min="35" max="100" step="1" data-field="blockWidthPercent" value="${escapeAttribute(question.blockWidthPercent)}">
          </label>
          <label>
            図や画像のスペース(mm)
            <input type="number" min="0" max="180" step="1" data-field="figureSpaceHeightMm" value="${escapeAttribute(question.figureSpaceHeightMm)}">
          </label>
        </div>
      </div>
    </article>
  `;
}

function renderQuestionList(validation) {
  if (!state.questions.length) {
    questionList.innerHTML = `
      <div class="empty-state">
        <div>
          <h3>まだ問題がありません</h3>
          <p>追加ボタンから問題を作るか、問題バンクから呼び出してください。</p>
        </div>
      </div>
    `;
    return;
  }
  questionList.innerHTML = state.questions.map((question, index) => renderQuestionCard(question, index, validation)).join("");
}

function renderQuestionBadges(question) {
  const badges = [
    question.unitTag ? `<span class="paper-tag">${escapeHtml(question.unitTag)}</span>` : "",
    question.difficulty ? `<span class="paper-tag is-difficulty">${escapeHtml(difficultyLabels[question.difficulty])}</span>` : ""
  ].filter(Boolean);
  return badges.length ? `<div class="paper-badges">${badges.join("")}</div>` : "";
}

function renderChoiceList(question, layout = getEffectiveLayoutConfig()) {
  const items = getChoiceItems(question);
  if (!items.length) {
    return "";
  }
  const questionLayout = getQuestionLayout(question, layout);
  const listTag = question.type === "order" ? "ul" : "ol";
  return `<${listTag} class="choice-list" style="font-size:${questionLayout.choiceFontSize}px;">${items.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</${listTag}>`;
}

function renderAnswerLinesHtml(question, layout = getEffectiveLayoutConfig()) {
  const lineCount = getAnswerLinesCount(question, layout);
  const questionLayout = getQuestionLayout(question, layout);
  return `
    <div class="answer-lines" style="min-height:${formatMeasure(questionLayout.answerAreaHeightMm)}mm;">
      ${Array.from({ length: lineCount }, () => '<div class="answer-line"></div>').join("")}
    </div>
  `;
}

function renderFigureSpace(questionLayout) {
  if (questionLayout.figureSpaceHeightMm <= 0) {
    return "";
  }
  return `<div class="question-figure-space" style="height:${formatMeasure(questionLayout.figureSpaceHeightMm)}mm;">図や画像を入れるスペース</div>`;
}

function renderStudentRowHtml() {
  const blocks = [];
  if (state.meta.showGroupNumberFields) {
    blocks.push(`<div><small>${escapeHtml(getMetaValue("groupLabel", "組"))}</small><div class="line-box"></div></div>`);
    blocks.push(`<div><small>${escapeHtml(getMetaValue("numberLabel", "番号"))}</small><div class="line-box"></div></div>`);
  }
  if (state.meta.showNameField) {
    blocks.push(`<div><small>${escapeHtml(getMetaValue("nameLabel", "名前"))}</small><div class="line-box"></div></div>`);
  }
  if (!blocks.length) {
    return "";
  }
  return `<div class="${blocks.length === 1 ? "student-row is-name-only" : "student-row"}">${blocks.join("")}</div>`;
}

function renderCommonPaperHeader(bundle, mode, paperConfig, layout) {
  const duration = parsePositiveIntStrict(state.meta.durationMinutes);
  const maxScore = parsePositiveIntStrict(state.meta.maxScore);
  const modeLabel = mode === "answer" ? "解答付き教師用" : mode === "answeronly" ? "解答用紙" : "問題用紙";
  return `
    <header class="paper-header">
      <div class="paper-kicker">
        <span>${escapeHtml(getMetaValue("schoolName", "学校名"))}</span>
        <span>${escapeHtml(getMetaValue("subject", "教科"))}</span>
      </div>
      <h2 class="paper-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))}</h2>
      <div class="paper-kicker">
        <span>学年: ${escapeHtml(getMetaValue("grade", "未設定"))}</span>
        <span>単元: ${escapeHtml(getMetaValue("unit", "未設定"))}</span>
        <span>クラス: ${escapeHtml(getMetaValue("className", "未設定"))}</span>
      </div>
      <div class="paper-kicker">
        <span>実施日: ${escapeHtml(formatDateLabel(state.meta.examDate))}</span>
        <span>制限時間: ${escapeHtml(duration.valid ? `${duration.value}分` : "未設定")}</span>
        <span>満点: ${escapeHtml(maxScore.valid ? `${maxScore.value}点` : "未設定")}</span>
      </div>
      <div class="paper-kicker">
        <span>用紙: ${escapeHtml(paperConfig.label)}</span>
        <span>${escapeHtml(paperConfig.tileLabel)}</span>
        <span>${escapeHtml(formatMeasure(paperConfig.widthMm))}mm × ${escapeHtml(formatMeasure(paperConfig.heightMm))}mm</span>
        <span>${escapeHtml(modeLabel)}</span>
        <span>${escapeHtml(bundle.label)}</span>
      </div>
      <p class="paper-notes">${formatMultilineHtml(state.meta.instructions, defaultInstructions)}</p>
      ${renderStudentRowHtml()}
    </header>
  `;
}

function renderPreviewQuestion(question, index, mode, layout = getEffectiveLayoutConfig()) {
  const questionLayout = getQuestionLayout(question, layout);
  const blockStyle = [
    `padding:${formatMeasure(questionLayout.blockPaddingMm)}mm`,
    `margin-bottom:${formatMeasure(layout.questionGap)}mm`,
    `font-size:${formatMeasure(questionLayout.questionFontSize)}px`,
    `width:min(100%, ${formatMeasure(questionLayout.blockWidthPercent)}%)`
  ].join("; ");

  if (mode === "answeronly") {
    return `
      <section class="paper-question" style="${blockStyle}">
        <div class="paper-question-top">
          <span>第${index + 1}問</span>
          <span>${escapeHtml(getQuestionPointsLabel(question))}</span>
        </div>
        <p>${escapeHtml(questionTypes[question.type].label)}の解答欄</p>
        ${renderAnswerLinesHtml(question, layout)}
      </section>
    `;
  }

  const teacherBox = mode === "answer"
    ? `
      <div class="answer-box">
        <strong>正解</strong>
        <p>${formatMultilineHtml(question.answer, "未入力")}</p>
      </div>
      <div class="note-box">
        <strong>解説</strong>
        <p>${formatMultilineHtml(question.explanation, "解説はまだありません。")}</p>
      </div>
    `
    : "";

  return `
    <section class="paper-question" style="${blockStyle}">
      <div class="paper-question-top">
        <span>第${index + 1}問 ${escapeHtml(questionTypes[question.type].label)}</span>
        <span>${escapeHtml(getQuestionPointsLabel(question))}</span>
      </div>
      <p>${formatMultilineHtml(question.prompt, "ここに問題文を入力してください。")}</p>
      ${renderQuestionBadges(question)}
      ${renderChoiceList(question, layout)}
      ${renderFigureSpace(questionLayout)}
      ${renderAnswerLinesHtml(question, layout)}
      ${teacherBox}
    </section>
  `;
}

function renderPaperView(bundle) {
  const validation = getValidationState(bundle);
  const layout = getEffectiveLayoutConfig();
  if (!bundle.questions.length) {
    paperSheet.innerHTML = `<div class="empty-state"><div><h3>まだ問題がありません</h3><p>問題を追加すると、ここに印刷イメージが表示されます。</p></div></div>`;
    return;
  }
  paperSheet.innerHTML = `${renderCommonPaperHeader(bundle, "paper", getPaperConfig(), layout)}${bundle.questions.map((question, index) => renderPreviewQuestion(question, index, "paper", layout)).join("")}`;
  renderStats(validation);
}

function renderAnswerView(bundle) {
  const layout = getEffectiveLayoutConfig();
  if (!bundle.questions.length) {
    answerSheet.innerHTML = `<div class="empty-state"><div><h3>解答付き教師用はまだ空です</h3><p>問題と正解を入れると、ここに表示されます。</p></div></div>`;
    return;
  }
  answerSheet.innerHTML = `${renderCommonPaperHeader(bundle, "answer", getPaperConfig(), layout)}${bundle.questions.map((question, index) => renderPreviewQuestion(question, index, "answer", layout)).join("")}`;
}

function renderAnswerOnlyView(bundle) {
  const layout = getEffectiveLayoutConfig();
  if (!bundle.questions.length) {
    answerOnlySheet.innerHTML = `<div class="empty-state"><div><h3>解答用紙はまだ空です</h3><p>問題を追加すると、ここに表示されます。</p></div></div>`;
    return;
  }
  answerOnlySheet.innerHTML = `${renderCommonPaperHeader(bundle, "answeronly", getPaperConfig(), layout)}${bundle.questions.map((question, index) => renderPreviewQuestion(question, index, "answeronly", layout)).join("")}`;
}

function renderCheckView(validation, bundle) {
  const typeMap = getTypeCountMap(bundle.questions);
  const difficultyMap = getDifficultyCountMap(bundle.questions);
  const tagMap = getTagCountMap(bundle.questions);
  const paperConfig = getPaperConfig();
  const layout = getEffectiveLayoutConfig();
  const estimate = estimateBundleHeightMm(bundle, paperConfig, layout);

  const typeEntries = Object.keys(questionTypes).map((key) => `${escapeHtml(questionTypes[key].label)}: ${typeMap[key] || 0}問`);
  const difficultyEntries = Object.keys(difficultyLabels).map((key) => `${escapeHtml(difficultyLabels[key])}: ${difficultyMap[key] || 0}問`);
  const tagEntries = Object.entries(tagMap).sort((left, right) => right[1] - left[1]).slice(0, 8).map(([tag, count]) => `${escapeHtml(tag)}: ${count}問`);

  analysisSheet.innerHTML = `
    <header class="paper-header">
      <div class="paper-kicker"><span>印刷前チェック</span><span>${escapeHtml(bundle.label)}</span></div>
      <h2 class="paper-title">設定の確認</h2>
      <p class="paper-notes">用紙サイズ、余白、配点、問題の種類をまとめて確認できます。</p>
    </header>
    <div class="analysis-grid">
      <section class="analysis-card">
        <h3>チェック結果</h3>
        <div class="validation-list">${renderValidationItemsHtml([...validation.errors, ...validation.warnings], "気になる点はありません。")}</div>
      </section>
      <section class="analysis-card">
        <h3>用紙とレイアウト</h3>
        ${formatSummaryList([
          `用紙サイズ: ${escapeHtml(paperConfig.label)}`,
          `横に何枚分: ${paperConfig.tilesX}枚`,
          `縦に何枚分: ${paperConfig.tilesY}枚`,
          `合計サイズ: ${formatMeasure(paperConfig.widthMm)}mm × ${formatMeasure(paperConfig.heightMm)}mm`,
          `レイアウト: ${escapeHtml(layout.label)}`,
          `文字の大きさ: ${formatMeasure(layout.globalFontSize)}px`,
          `問題の間隔: ${formatMeasure(layout.questionGap)}mm`,
          `左右余白: ${formatMeasure(layout.sideMarginMm)}mm`,
          `上下余白: ${formatMeasure(layout.topBottomMarginMm)}mm`
        ], "まだ確認項目がありません。")}
      </section>
      <section class="analysis-card">
        <h3>配点と時間</h3>
        ${formatSummaryList([
          `問題数: ${bundle.questions.length}問`,
          `合計点: ${bundle.questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0)}点`,
          `設定した満点: ${validation.maxScoreValue === null ? "未設定" : `${validation.maxScoreValue}点`}`,
          `制限時間: ${validation.durationValue === null ? "未設定" : `${validation.durationValue}分`}`,
          `1ページ想定: ${Math.round(estimate.contentHeight)}mm / 使える高さ ${Math.round(estimate.availableHeight)}mm`
        ], "まだ集計できていません。")}
      </section>
      <section class="analysis-card">
        <h3>問題形式</h3>
        ${formatSummaryList(typeEntries, "問題がありません。")}
      </section>
      <section class="analysis-card">
        <h3>難しさ</h3>
        ${formatSummaryList(difficultyEntries, "問題がありません。")}
      </section>
      <section class="analysis-card">
        <h3>単元タグ</h3>
        ${formatSummaryList(tagEntries, "単元タグはまだありません。")}
      </section>
    </div>
  `;
}

function applyPaperSizeToPreview() {
  const paperConfig = getPaperConfig();
  const layout = getEffectiveLayoutConfig();
  [paperSheet, answerSheet, answerOnlySheet, analysisSheet].forEach((sheet) => {
    sheet.style.setProperty("--paper-screen-width", `${paperConfig.previewWidthPx}px`);
    sheet.style.setProperty("--paper-screen-height", `${paperConfig.previewHeightPx}px`);
    sheet.dataset.paperSize = paperConfig.paperBadge;
    sheet.dataset.layoutMode = state.meta.layoutMode || "standard";
    sheet.style.padding = `${formatMeasure(layout.topBottomMarginMm)}mm ${formatMeasure(layout.sideMarginMm)}mm`;
    sheet.style.fontSize = `${formatMeasure(layout.globalFontSize)}px`;
    sheet.style.lineHeight = String(layout.lineHeight);
  });
  previewPaperLabel.textContent = paperConfig.headerLabel;
  paperBaseSizeDisplay.textContent = Number.isFinite(paperConfig.baseWidthMm) && Number.isFinite(paperConfig.baseHeightMm)
    ? `もとの用紙サイズ: ${paperConfig.label} ${formatMeasure(paperConfig.baseWidthMm)}mm × ${formatMeasure(paperConfig.baseHeightMm)}mm`
    : "もとの用紙サイズ: 自由サイズを入力してください";
  paperTotalSizeDisplay.textContent = Number.isFinite(paperConfig.widthMm) && Number.isFinite(paperConfig.heightMm)
    ? `合計サイズ: ${formatMeasure(paperConfig.widthMm)}mm × ${formatMeasure(paperConfig.heightMm)}mm`
    : "合計サイズ: サイズがまだ完成していません";
  layoutSummary.textContent = `レイアウト: ${layout.label} / 文字の大きさ ${formatMeasure(layout.globalFontSize)}px / 問題の間隔 ${formatMeasure(layout.questionGap)}mm / 行間 ${formatMeasure(layout.lineHeight)} / 左右余白 ${formatMeasure(layout.sideMarginMm)}mm / 上下余白 ${formatMeasure(layout.topBottomMarginMm)}mm`;
  updateCustomPaperVisibility();
}

function updatePreviewPrintButton() {
  const bundle = getActiveBundle();
  if (uiState.activeTab === "answer") {
    printPreviewButton.textContent = `${bundle.label}の解答付き教師用を印刷またはPDF保存`;
    return;
  }
  if (uiState.activeTab === "answeronly") {
    printPreviewButton.textContent = `${bundle.label}の解答用紙を印刷またはPDF保存`;
    return;
  }
  printPreviewButton.textContent = `${bundle.label}の問題用紙を印刷またはPDF保存`;
}

function renderDerivedViews(validation = getValidationState()) {
  const bundle = getActiveBundle();
  applyPaperSizeToPreview();
  renderTemplateSummary();
  renderStats(validation);
  renderPaperView(bundle);
  renderAnswerView(bundle);
  renderAnswerOnlyView(bundle);
  renderCheckView(validation, bundle);
  renderBankList();
  renderPatternList();
  updatePreviewScope();
  updatePreviewPrintButton();
  editorStatus.textContent = uiState.editorMessage;
}

function updatePreviewPrintButton() {
  const bundle = getActiveBundle();
  if (uiState.activeTab === "answer") {
    printPreviewButton.textContent = `${bundle.label}の解答付き教師用を印刷またはPDF保存`;
    return;
  }
  if (uiState.activeTab === "answeronly") {
    printPreviewButton.textContent = `${bundle.label}の解答用紙を印刷またはPDF保存`;
    return;
  }
  printPreviewButton.textContent = `${bundle.label}の問題用紙を印刷またはPDF保存`;
}

function buildPrintableHtml(mode, bundle) {
  const normalizedMode = stage3GetNormalizedPrintMode(mode);
  const paperConfig = getPaperConfig();
  const layout = getEffectiveLayoutConfig();
  const totalPointsValue = bundle.questions.reduce((sum, question) => sum + (getQuestionPointsValue(question) ?? 0), 0);
  const labelMap = {
    paper: "問題用紙",
    answeronly: "解答用紙",
    teacher: "解答付き教師用",
    "paper-answer-set": "問題用紙 + 解答用紙"
  };
  const label = labelMap[normalizedMode] || "問題用紙";
  const printMetaColumns = state.meta.showGroupNumberFields && state.meta.showNameField ? "1fr 1fr 1.4fr" : state.meta.showGroupNumberFields ? "1fr 1fr" : "1fr";
  const sections = normalizedMode === "paper-answer-set"
    ? [
        stage3BuildPrintableSection("問題用紙", "paper", bundle, layout, paperConfig, totalPointsValue),
        stage3BuildPrintableSection("解答用紙", "answeronly", bundle, layout, paperConfig, totalPointsValue)
      ]
    : [stage3BuildPrintableSection(label, normalizedMode, bundle, layout, paperConfig, totalPointsValue)];

  return `
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${escapeHtml(getMetaValue("examTitle", "テスト名"))} ${escapeHtml(label)} ${escapeHtml(bundle.label)}</title>
      <style>
        :root { color-scheme: light; }
        * { box-sizing: border-box; }
        body { margin: 0; font-family: "Yu Gothic", "Hiragino Sans", sans-serif; background: #edf1ef; color: #111; }
        .toolbar { position: sticky; top: 0; z-index: 10; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 16px 20px; background: #24343b; color: #fff; }
        .toolbar-title { font-size: 1rem; font-weight: 700; }
        .toolbar-note { font-size: 0.88rem; opacity: 0.88; }
        .toolbar-button { border: none; border-radius: 999px; padding: 12px 18px; background: #fff; color: #24343b; font-size: 0.95rem; font-weight: 700; cursor: pointer; }
        .sheet { width: ${formatMeasure(paperConfig.widthMm)}mm; min-height: ${formatMeasure(paperConfig.heightMm)}mm; margin: 18px auto; padding: ${formatMeasure(layout.topBottomMarginMm)}mm ${formatMeasure(layout.sideMarginMm)}mm; background: #fff; box-shadow: 0 12px 30px rgba(0, 0, 0, 0.08); font-size: ${formatMeasure(layout.globalFontSize)}px; line-height: ${formatMeasure(layout.lineHeight)}; }
        .sheet + .sheet { page-break-before: always; break-before: page; }
        .print-header { border-bottom: 2px solid #222; padding-bottom: 10mm; margin-bottom: 8mm; }
        .print-header-top, .print-summary { display: flex; flex-wrap: wrap; gap: 12px 18px; font-size: 0.94rem; }
        .print-header-title { margin: 8mm 0 3mm; text-align: center; font-size: 1.8rem; font-weight: 700; letter-spacing: 0.04em; }
        .print-header-subtitle { text-align: center; font-size: 1rem; font-weight: 700; margin-bottom: 6mm; }
        .print-meta-grid { display: grid; grid-template-columns: ${printMetaColumns}; gap: 12px; margin: 6mm 0; }
        .meta-box { border: 1px solid #333; padding: 10px 12px; min-height: 16mm; }
        .meta-box-label { font-size: 0.82rem; margin-bottom: 6px; font-weight: 700; }
        .meta-box-line { border-bottom: 1px solid #222; height: 18px; }
        .print-instructions { margin-top: 4mm; padding: 10px 12px; border: 1px solid #444; background: #fafafa; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .print-question { border-bottom: 1px solid #bbb; page-break-inside: avoid; break-inside: avoid; }
        .print-question:last-child { border-bottom: none; }
        .print-question-top { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 4px; }
        .print-question-title { font-size: 1.08rem; font-weight: 700; }
        .print-question-points { min-width: 68px; text-align: right; font-size: 0.98rem; font-weight: 700; }
        .print-badges { display: flex; flex-wrap: wrap; gap: 8px; margin: 6px 0 10px; }
        .print-badges span { border: 1px solid #666; padding: 3px 8px; font-size: 0.8rem; }
        .print-prompt, .teacher-answer-value, .teacher-note { margin: 0 0 10px; white-space: pre-wrap; overflow-wrap: anywhere; word-break: break-word; }
        .print-choices { margin: 0 0 10px 1.4em; padding: 0; white-space: pre-wrap; overflow-wrap: anywhere; }
        .response-lines { display: grid; gap: 10px; margin-top: 12px; }
        .write-line { border-bottom: 1px solid #222; min-height: 24px; }
        .teacher-answer { margin-top: 12px; border: 1px solid #222; background: #f8f8f8; padding: 10px 12px; }
        .teacher-answer-label { font-size: 0.82rem; font-weight: 700; margin-bottom: 6px; }
        .teacher-note strong { display: inline-block; margin-right: 6px; }
        .figure-space { display: grid; place-items: center; margin-top: 8px; border: 1px dashed #aaa; background: #fafafa; color: #566; font-size: 0.84rem; }
        .footer-total { margin-top: 8mm; text-align: right; font-size: 1rem; font-weight: 700; }
        @page { size: ${formatMeasure(paperConfig.widthMm)}mm ${formatMeasure(paperConfig.heightMm)}mm; margin: 0; }
        @media print { body { background: #fff; } .toolbar { display: none; } .sheet { width: auto; min-height: auto; margin: 0; box-shadow: none; } }
      </style>
    </head>
    <body>
      <div class="toolbar">
        <div>
          <div class="toolbar-title">${escapeHtml(getMetaValue("examTitle", "テスト名"))} / ${escapeHtml(label)} / ${escapeHtml(bundle.label)}</div>
          <div class="toolbar-note">上のボタンから印刷またはPDF保存に進めます。</div>
        </div>
        <button class="toolbar-button" onclick="window.print()">印刷またはPDF保存</button>
      </div>
      ${sections.join("")}
    </body>
    </html>
  `;
}

function openPrintableDocument(mode, bundle = getActiveBundle()) {
  const normalizedMode = stage3GetNormalizedPrintMode(mode);
  const labelMap = {
    paper: "問題用紙を印刷",
    answeronly: "解答用紙を印刷",
    teacher: "解答付き教師用を印刷",
    "paper-answer-set": "問題用紙と解答用紙を印刷"
  };
  const label = labelMap[normalizedMode] || "問題用紙を印刷";
  syncMetaFromInputs();
  if (!ensureReadyForOutput(label, bundle.questions)) {
    return;
  }
  const exportWindow = window.open("", "_blank");
  if (!exportWindow) {
    window.alert("新しい画面を開けませんでした。ポップアップブロックを一時的に解除して、もう一度お試しください。");
    return;
  }
  exportWindow.document.open();
  exportWindow.document.write(buildPrintableHtml(normalizedMode, bundle));
  exportWindow.document.close();
  setEditorMessage(`${bundle.label}の${label}を開きました。上のボタンから印刷またはPDF保存に進めます。`);
}

function renderDerivedViews(validation = getValidationState()) {
  const bundle = getActiveBundle();
  applyPaperSizeToPreview();
  renderTemplateSummary();
  renderStats(validation);
  renderPaperView(bundle);
  renderAnswerView(bundle);
  renderAnswerOnlyView(bundle);
  renderCheckView(validation, bundle);
  renderBankList();
  renderPatternList();
  updatePreviewScope();
  updatePreviewPrintButton();
  editorStatus.textContent = uiState.editorMessage;
}

function stage3EnsureIntegratedPaperButton() {
  const exportActions = document.querySelector(".export-actions");
  const studentButton = document.getElementById("exportStudentPdf");
  if (!exportActions || !studentButton || document.getElementById("exportIntegratedSheet")) {
    return;
  }
  const button = document.createElement("button");
  button.type = "button";
  button.id = "exportIntegratedSheet";
  button.className = "secondary";
  button.textContent = "問題と回答欄を一体化して印刷";
  button.addEventListener("click", () => openPrintableDocument("paper"));
  studentButton.insertAdjacentElement("afterend", button);
}

stage3EnsureIntegratedPaperButton();
stage3EnsureCombinedExportButton();
renderDerivedViews(getValidationState());
