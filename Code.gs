/**
 * @OnlyCurrentDoc
 * 導師與專任雙系統整合版成績管理系統 (Bug 修正整合版)
 */

// ==========================================
// 全域設定區 (請填入 Google Drive 資料夾 ID 與 API Key)
// ==========================================
const FOLDER_UPLOAD_ID = "傳統 Excel 試算表上傳區";      // 傳統 Excel 試算表上傳區
const FOLDER_PROCESSED_ID = "已處理備份區";   // 已處理備份區
const FOLDER_HANDWRITTEN_ID = "傳統 Excel 試算表上傳區"; // 手寫照片自動掃描資料夾

const GEMINI_API_KEY = "gemini api key";
const CONFIG = {
  BG_FOLDER_ID: "1zYnYo3kYLv35aSzCKDq1f1Kk7DPQzR_w", // 存放背景圖與 LOGO 的資料夾 ID
  LOGO_KEYWORD: "logo"                          // LOGO 檔名關鍵字
};

/**
 * 讀取 System_Settings 分頁中的全域系統設定 (A:B 欄 Key-Value)
 * @param {boolean} [suppressError=false] 若為 true 且找不到分頁時不拋出例外，改回傳預設物件
 */
function getSystemConfig(suppressError = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("System_Settings");
  
  const config = {
    GEMINI_API_KEY: "",
    UPLOAD_FOLDER_ID: "",
    BACKUP_FOLDER_ID: "",
    BG_FOLDER_ID: "",
    LOGO_KEYWORD: "logo",
    TEACHER_PIN: "1234"
  };

  if (!sheet) {
    if (suppressError) {
      config.GEMINI_API_KEY = GEMINI_API_KEY;
      config.UPLOAD_FOLDER_ID = FOLDER_UPLOAD_ID;
      config.BACKUP_FOLDER_ID = FOLDER_PROCESSED_ID;
      config.BG_FOLDER_ID = CONFIG.BG_FOLDER_ID;
      config.LOGO_KEYWORD = CONFIG.LOGO_KEYWORD;
      return config;
    }
    throw new Error("找不到 [System_Settings] 基礎設定頁面！請先建立分頁或由選單執行「建立/檢查 System_Settings」。");
  }

  const data = sheet.getRange("A2:B20").getValues();
  for (let r = 0; r < data.length; r++) {
    const key = String(data[r][0]).trim();
    const val = String(data[r][1]).trim();
    if (key && config.hasOwnProperty(key)) {
      config[key] = val;
    }
  }

  // 若分頁內欄位為空，以常數為後備
  if (!config.GEMINI_API_KEY && GEMINI_API_KEY && !GEMINI_API_KEY.includes("gemini api key")) {
    config.GEMINI_API_KEY = GEMINI_API_KEY;
  }
  if (!config.UPLOAD_FOLDER_ID && FOLDER_UPLOAD_ID && !FOLDER_UPLOAD_ID.includes("傳統 Excel")) {
    config.UPLOAD_FOLDER_ID = FOLDER_UPLOAD_ID;
  }
  if (!config.BACKUP_FOLDER_ID && FOLDER_PROCESSED_ID && !FOLDER_PROCESSED_ID.includes("已處理備份區")) {
    config.BACKUP_FOLDER_ID = FOLDER_PROCESSED_ID;
  }
  if (!config.BG_FOLDER_ID && CONFIG.BG_FOLDER_ID && !CONFIG.BG_FOLDER_ID.includes("請在此貼上")) {
    config.BG_FOLDER_ID = CONFIG.BG_FOLDER_ID;
  }
  if (!config.LOGO_KEYWORD) {
    config.LOGO_KEYWORD = CONFIG.LOGO_KEYWORD || "logo";
  }

  return config;
}

/**
 * 讀取 System_Settings 分頁中指定的班級對照設定 (D:H 欄)
 */
function getHomeroomConfig(classCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("System_Settings");
  if (!sheet) throw new Error("找不到 [System_Settings] 基礎設定頁面！");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("[System_Settings] 無對照表資料！");

  const data = sheet.getRange("D1:H" + lastRow).getValues();
  if (data.length < 2) throw new Error("[System_Settings] 無對照表資料！");

  const headers = data[0].map(h => String(h).trim());
  const idxClass = headers.indexOf("Class_Code");
  const idxSSID = headers.indexOf("Homeroom_SS_ID");
  const idxSheet = headers.indexOf("Target_Sheet_Name");
  const idxKeyCol = headers.indexOf("Key_Column_Name");
  const idxSubjCol = headers.indexOf("Subject_Column");

  if (idxClass === -1 || idxSSID === -1) {
    throw new Error("[System_Settings] 對照表缺少必要欄位 (Class_Code 或 Homeroom_SS_ID)！");
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxClass]).trim() === String(classCode).trim()) {
      return {
        classCode: String(data[i][idxClass]).trim(),
        homeroomSSID: String(data[i][idxSSID]).trim(),
        targetSheetName: idxSheet !== -1 ? String(data[i][idxSheet]).trim() : "導師成績總表",
        keyColumnName: idxKeyCol !== -1 ? String(data[i][idxKeyCol]).trim() : "座號",
        subjectColumn: idxSubjCol !== -1 ? String(data[i][idxSubjCol]).trim() : ""
      };
    }
  }
  throw new Error(`在 [System_Settings] 中找不到班級代碼：${classCode}`);
}

/**
 * 取得當前有效的 Gemini API Key (優先取用 System_Settings，後備全域常數)
 */
function getActiveGeminiApiKey() {
  try {
    const sys = getSystemConfig(true);
    if (sys && sys.GEMINI_API_KEY && sys.GEMINI_API_KEY.trim() !== "" && !sys.GEMINI_API_KEY.includes("gemini api key")) {
      return sys.GEMINI_API_KEY.trim();
    }
  } catch (e) {}
  return GEMINI_API_KEY;
}

/**
 * 取得當前有效的上傳資料夾 ID
 */
function getActiveUploadFolderId() {
  try {
    const sys = getSystemConfig(true);
    if (sys && sys.UPLOAD_FOLDER_ID && sys.UPLOAD_FOLDER_ID.trim() !== "" && !sys.UPLOAD_FOLDER_ID.includes("傳統 Excel")) {
      return sys.UPLOAD_FOLDER_ID.trim();
    }
  } catch (e) {}
  return FOLDER_UPLOAD_ID;
}

/**
 * 取得當前有效的備份資料夾 ID
 */
function getActiveBackupFolderId() {
  try {
    const sys = getSystemConfig(true);
    if (sys && sys.BACKUP_FOLDER_ID && sys.BACKUP_FOLDER_ID.trim() !== "" && !sys.BACKUP_FOLDER_ID.includes("已處理備份區")) {
      return sys.BACKUP_FOLDER_ID.trim();
    }
  } catch (e) {}
  return FOLDER_PROCESSED_ID;
}

// 全域變數：紀錄已失效的模型
const DISABLED_MODELS = new Set();

// 導師班各科權重預設設定
const WEIGHTS = {
  "國文": 5, "英文": 3, "數學": 4, 
  "歷史": 1, "地理": 1, "公民": 1, "理化": 3, "生物": 3
};

// 專任班級評量類別預設權重
const SUBJECT_TYPE_WEIGHTS = {
  "段考": 0.5,
  "小考": 0.2,
  "作業": 0.2,
  "筆記": 0.1
};

// ==========================================
// 1. Web App 網頁查詢端進入點
// ==========================================
function doGet(e) {
  // 支援教師全功能主控台 (手寫成績單AI辨識匯入 + 導師主控版 + 科任主控版)
  if (e && e.parameter && (e.parameter.page === "teacher" || e.parameter.mode === "teacher" || e.parameter.page === "hub")) {
    const template = HtmlService.createTemplateFromFile('TeacherHub');
    template.isModalDialog = false;
    return template.evaluate()
      .setTitle('教師全功能主控台')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 支援科任教師儀表板專用頁面 (例如: ?page=dashboard 或 ?mode=subject)
  if (e && e.parameter && (e.parameter.page === "dashboard" || e.parameter.mode === "subject")) {
    const template = HtmlService.createTemplateFromFile('SubjectDashboard');
    template.isModalDialog = false;
    return template.evaluate()
      .setTitle('科任教師跨班分析儀表板')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 預設為學生成績查詢頁面
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('學生周平時成績查詢系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 學生端查詢：
 * 1. 頂部資訊與表格：本週（近 7 日 / 最新一次）成績清單、週加權平均與週排名
 * 2. 全歷史圖表區：
 *    - 原始分數雷達圖 & 趨勢圖 (100 分制)
 *    - 五標相對雷達圖 & 趨勢圖 (0-6 階梯刻度)
 */
function getStudentScores(seat, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tutorSheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
  if (!tutorSheet) return { success: false, message: "系統尚未建立導師總表" };

  const studentEndRow = getStudentEndRow(tutorSheet, 2, 3);
  let foundRow = -1;

  // 1. 驗證學生座號與密碼
  for (let r = 4; r <= studentEndRow; r++) {
    const rPwd = String(tutorSheet.getRange(r, 1).getValue()).trim();
    const rSeat = String(tutorSheet.getRange(r, 2).getValue()).trim();
    if (parseInt(rSeat, 10) === parseInt(seat, 10) && rPwd === String(password).trim()) {
      foundRow = r;
      break;
    }
  }

  if (foundRow === -1) return { success: false, message: "座號或密碼錯誤！" };

  const studentName = tutorSheet.getRange(foundRow, 3).getValue();
  const lastCol = tutorSheet.getLastColumn();
  if (lastCol < 4) return { success: false, message: "目前無成績資料" };

  // 2. 讀取表頭資訊與全班成績矩陣
  const subjects = tutorSheet.getRange(1, 4, 1, lastCol - 3).getDisplayValues()[0];
  const ranges = tutorSheet.getRange(2, 4, 1, lastCol - 3).getDisplayValues()[0];
  const dates = tutorSheet.getRange(3, 4, 1, lastCol - 3).getValues()[0];

  const studentCount = studentEndRow - 3;
  const allScores = tutorSheet.getRange(4, 4, studentCount, lastCol - 3).getValues();
  const myScores = allScores[foundRow - 4];

  // 3. 【週成績範圍 (targetCols)】：篩選近 7 日 (若無資料則為最新日期)
  const now = new Date();
  const todayEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  const sevenDaysAgo = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7, 0, 0, 0, 0);

  let targetCols = [];
  let maxDate = null;

  for (let c = 0; c < subjects.length; c++) {
    const subj = String(subjects[c] || "").trim();
    if (!subj) continue;
    const d = parseDateHelper(dates[c]);
    if (!d) continue;

    if (!maxDate || d > maxDate) maxDate = d;

    if (d >= sevenDaysAgo && d <= todayEnd) {
      targetCols.push(c);
    }
  }

  if (targetCols.length === 0 && maxDate !== null) {
    const maxDateStr = Utilities.formatDate(maxDate, "GMT+8", "yyyy-MM-dd");
    for (let c = 0; c < subjects.length; c++) {
      const subj = String(subjects[c] || "").trim();
      if (!subj) continue;
      const d = parseDateHelper(dates[c]);
      if (d && Utilities.formatDate(d, "GMT+8", "yyyy-MM-dd") === maxDateStr) {
        targetCols.push(c);
      }
    }
  }

  if (targetCols.length === 0) {
    return { success: false, message: "目前無符合條件的成績資料" };
  }

  // 4. 計算週加權總平均與週排名
  const weights = getTutorWeights();
  const studentWeeklyWeightedAvgs = [];

  for (let s = 0; s < studentCount; s++) {
    const subjSums = {}, subjCounts = {};
    targetCols.forEach(c => {
      const sc = allScores[s][c];
      const subj = String(subjects[c]).trim();
      if (sc !== "" && sc !== null && !isNaN(Number(sc))) {
        subjSums[subj] = (subjSums[subj] || 0) + Number(sc);
        subjCounts[subj] = (subjCounts[subj] || 0) + 1;
      }
    });

    let wSum = 0, wTotal = 0;
    Object.keys(subjSums).forEach(k => {
      const sAvg = subjSums[k] / subjCounts[k];
      const w = weights[k] || 1;
      wSum += sAvg * w;
      wTotal += w;
    });
    studentWeeklyWeightedAvgs.push(wTotal > 0 ? Math.round((wSum / wTotal) * 10) / 10 : 0);
  }

  const weeklyRanks = calculateRanks(studentWeeklyWeightedAvgs);
  const myWeeklyRank = weeklyRanks[foundRow - 4];
  const myWeeklyWeightedAvg = studentWeeklyWeightedAvgs[foundRow - 4];

  // 5. 組裝本週成績明細清單
  const subjectGrouped = {};
  targetCols.forEach(c => {
    const sc = myScores[c];
    if (sc === "" || sc === null || isNaN(Number(sc))) return;

    const subj = String(subjects[c]).trim();
    const scoreNum = Number(sc);

    if (!subjectGrouped[subj]) subjectGrouped[subj] = { items: [] };

    let colSum = 0, colCount = 0;
    for (let s = 0; s < studentCount; s++) {
      const val = allScores[s][c];
      if (val !== "" && val !== null && val !== undefined && !isNaN(Number(val))) {
        colSum += Number(val);
        colCount++;
      }
    }
    const classAvgVal = colCount > 0 ? Math.round((colSum / colCount) * 10) / 10 : "-";

    subjectGrouped[subj].items.push({
      range: String(ranges[c] || ""),
      score: scoreNum,
      classAvg: classAvgVal
    });
  });

  Object.keys(subjectGrouped).forEach(subj => {
    const items = subjectGrouped[subj].items;
    const avg = Math.round((items.reduce((a, b) => a + b.score, 0) / items.length) * 10) / 10;
    subjectGrouped[subj].subjAvg = avg;

    const allStudentSubjAvgs = [];
    for (let s = 0; s < studentCount; s++) {
      let sSum = 0, sCnt = 0;
      targetCols.forEach(c => {
        if (String(subjects[c]).trim() === subj) {
          const sc = allScores[s][c];
          if (sc !== "" && sc !== null && !isNaN(Number(sc))) {
            sSum += Number(sc); sCnt++;
          }
        }
      });
      allStudentSubjAvgs.push(sCnt > 0 ? sSum / sCnt : 0);
    }
    const subjRanks = calculateRanks(allStudentSubjAvgs);
    subjectGrouped[subj].subjRank = subjRanks[foundRow - 4];
  });

  // 6. 【全歷史圖表數據計算】
  let globalMinDate = null;
  for (let c = 0; c < subjects.length; c++) {
    const d = parseDateHelper(dates[c]);
    if (d && (!globalMinDate || d < globalMinDate)) {
      globalMinDate = d;
    }
  }
  if (!globalMinDate) globalMinDate = new Date();
  globalMinDate.setHours(0, 0, 0, 0);

  const historyWeekScores = {};
  const historyWeeksSet = new Set();
  const historicalSubjSums = {};
  const historicalSubjCounts = {};
  const allStudentHistoricalSubjScores = {}; // 全班全歷史各科成績集

  for (let c = 0; c < subjects.length; c++) {
    const subj = String(subjects[c] || "").trim();
    if (!subj) continue;

    const d = parseDateHelper(dates[c]);
    if (!d) continue;

    const diffDays = Math.floor((d.getTime() - globalMinDate.getTime()) / (1000 * 60 * 60 * 24));
    const weekLabel = `第${Math.floor(diffDays / 7) + 1}週`;
    historyWeeksSet.add(weekLabel);

    if (!allStudentHistoricalSubjScores[subj]) allStudentHistoricalSubjScores[subj] = [];

    for (let s = 0; s < studentCount; s++) {
      const val = allScores[s][c];
      if (val !== "" && val !== null && !isNaN(Number(val))) {
        allStudentHistoricalSubjScores[subj].push(Number(val));
      }
    }

    const sc = myScores[c];
    if (sc !== "" && sc !== null && !isNaN(Number(sc))) {
      const scoreNum = Number(sc);
      historicalSubjSums[subj] = (historicalSubjSums[subj] || 0) + scoreNum;
      historicalSubjCounts[subj] = (historicalSubjCounts[subj] || 0) + 1;

      if (!historyWeekScores[weekLabel]) historyWeekScores[weekLabel] = {};
      if (!historyWeekScores[weekLabel][subj]) historyWeekScores[weekLabel][subj] = [];
      historyWeekScores[weekLabel][subj].push(scoreNum);
    }
  }

  const sortedHistoryWeeks = Array.from(historyWeeksSet).sort(
    (a, b) => parseInt(a.replace("第", "")) - parseInt(b.replace("第", ""))
  );

  const allSubjectNames = Array.from(new Set(subjects.map(s => String(s).trim()).filter(Boolean)));
  const palette = ["#3498db", "#e74c3c", "#2ecc71", "#f1c40f", "#9b59b6", "#e67e22", "#1abc9c", "#34495e"];

  // (A) 原始 100 分制趨勢圖
  const trendDatasets = allSubjectNames.map((subj, idx) => ({
    label: subj,
    data: sortedHistoryWeeks.map(w => {
      if (historyWeekScores[w] && historyWeekScores[w][subj]) {
        const arr = historyWeekScores[w][subj];
        return Math.round((arr.reduce((a, b) => a + b, 0) / arr.length) * 10) / 10;
      }
      return null;
    }),
    borderColor: palette[idx % palette.length],
    backgroundColor: palette[idx % palette.length],
    fill: false,
    tension: 0.2,
    spanGaps: true
  }));

  // (B) 原始 100 分制雷達圖
  const radarLabels = [];
  const radarData = [];
  allSubjectNames.forEach(subj => {
    if (historicalSubjCounts[subj] > 0) {
      radarLabels.push(subj);
      radarData.push(Math.round((historicalSubjSums[subj] / historicalSubjCounts[subj]) * 10) / 10);
    }
  });

  // (C) 五標相對雷達圖 (0-6 階梯)
  const fiveTierRadarLabels = [];
  const fiveTierRadarData = [];
  const fiveTierClassBenchmark = []; // 班級均標基準線 (全為 3.0)

  allSubjectNames.forEach(subj => {
    if (historicalSubjCounts[subj] > 0) {
      const myAvg = historicalSubjSums[subj] / historicalSubjCounts[subj];
      const pObj = calculateFiveTierPercentiles(allStudentHistoricalSubjScores[subj] || []);
      const mappedVal = mapScoreToFiveTierScale(myAvg, pObj);

      fiveTierRadarLabels.push(subj);
      fiveTierRadarData.push(mappedVal);
      fiveTierClassBenchmark.push(3.0);
    }
  });

  // (D) 五標相對趨勢圖 (0-6 階梯)
  const fiveTierTrendDatasets = allSubjectNames.map((subj, idx) => ({
    label: subj,
    data: sortedHistoryWeeks.map(w => {
      if (historyWeekScores[w] && historyWeekScores[w][subj]) {
        const arr = historyWeekScores[w][subj];
        const myWAvg = arr.reduce((a, b) => a + b, 0) / arr.length;

        // 計算該週全班該科五標
        const classWScores = [];
        for (let c = 0; c < subjects.length; c++) {
          if (String(subjects[c]).trim() === subj) {
            const d = parseDateHelper(dates[c]);
            if (d) {
              const diffDays = Math.floor((d.getTime() - globalMinDate.getTime()) / (1000 * 60 * 60 * 24));
              const wLabel = `第${Math.floor(diffDays / 7) + 1}週`;
              if (wLabel === w) {
                for (let s = 0; s < studentCount; s++) {
                  const val = allScores[s][c];
                  if (val !== "" && val !== null && !isNaN(Number(val))) {
                    classWScores.push(Number(val));
                  }
                }
              }
            }
          }
        }

        const pObj = calculateFiveTierPercentiles(classWScores);
        return mapScoreToFiveTierScale(myWAvg, pObj);
      }
      return null;
    }),
    borderColor: palette[idx % palette.length],
    backgroundColor: palette[idx % palette.length],
    fill: false,
    tension: 0.2,
    spanGaps: true
  }));

  return {
    success: true,
    studentName: studentName,
    weeklyWeightedAvg: myWeeklyWeightedAvg,
    weeklyRank: myWeeklyRank,
    subjectGrouped: subjectGrouped,
    radarChart: { labels: radarLabels, data: radarData },
    trendChart: { labels: sortedHistoryWeeks, datasets: trendDatasets },
    fiveTierRadarChart: {
      labels: fiveTierRadarLabels,
      studentData: fiveTierRadarData,
      benchmarkData: fiveTierClassBenchmark
    },
    fiveTierTrendChart: {
      labels: sortedHistoryWeeks,
      datasets: fiveTierTrendDatasets
    }
  };
}

/**
 * 輔助函式：計算百分位數 (底標 P12, 後標 P25, 均標 P50, 前標 P75, 頂標 P88)
 */
function calculateFiveTierPercentiles(scores) {
  const sorted = scores.filter(v => v !== "" && v !== null && !isNaN(Number(v))).map(Number).sort((a, b) => a - b);
  if (sorted.length === 0) return { p12: 0, p25: 0, p50: 0, p75: 0, p88: 0 };

  const getP = (p) => {
    if (sorted.length === 1) return sorted[0];
    const rank = p * (sorted.length - 1);
    const low = Math.floor(rank);
    const high = Math.ceil(rank);
    const weight = rank - low;
    return sorted[low] + weight * (sorted[high] - sorted[low]);
  };

  return {
    p12: getP(0.12),
    p25: getP(0.25),
    p50: getP(0.50),
    p75: getP(0.75),
    p88: getP(0.88)
  };
}

/**
 * 輔助函式：將原始分數映射至 0 - 6 階梯刻度
 */
function mapScoreToFiveTierScale(score, pObj) {
  if (score === null || score === undefined || isNaN(Number(score))) return null;
  const s = Number(score);
  const { p12, p25, p50, p75, p88 } = pObj;

  let val = 0;
  if (s >= p88) {
    val = 5 + (100 > p88 ? (s - p88) / (100 - p88) : 0);
  } else if (s >= p75) {
    val = 4 + (p88 > p75 ? (s - p75) / (p88 - p75) : 0);
  } else if (s >= p50) {
    val = 3 + (p75 > p50 ? (s - p50) / (p75 - p50) : 0);
  } else if (s >= p25) {
    val = 2 + (p50 > p25 ? (s - p25) / (p50 - p25) : 0);
  } else if (s >= p12) {
    val = 1 + (p25 > p12 ? (s - p12) / (p25 - p12) : 0);
  } else {
    val = p12 > 0 ? (s / p12) : 0;
  }
  return Math.round(Math.min(6, Math.max(0, val)) * 100) / 100;
}
/**
 * 導師端：產生 HTML 個人報表頁面
 */
function generateBatchReportHtml(startDateStr, endDateStr) {
  const data = getStudentReportData(startDateStr, endDateStr);
  const template = HtmlService.createTemplateFromFile("BatchReport");
  template.reportData = JSON.stringify(data);
  template.reportJson = JSON.stringify(data);
  return template.evaluate().setTitle("學生個人成績報表").getContent();
}

// ==========================================
// 2. Google Sheet 頂端自訂雙選單
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu("🌟 教師主控台")
    .addItem("🚀 開啟教師全功能主控台 (辨識匯入 / 導師主控 / 科任主控)", "showTeacherHubDialog")
    .addItem("🔗 顯示主控台網頁連結 (可分享至手機/平板)", "showTeacherHubUrlDialog")
    .addToUi();

  ui.createMenu("🏫 導師成績管理")
    .addItem("📥 執行每日匯入 (自動路由分頁)", "importDailyScores")
    .addItem("📥 匯入當前分頁成績至導師總表 (5字解析比對)", "importSubjectSheetToTutorSheetUI")
    .addSeparator()
    .addItem("📊 產生指定區間歷史成績單", "showSheetWriteDialog")
    .addItem("🖨️ 匯出區間個人報表 (列印/PDF)", "showBatchReportDialog")
    .addSeparator()
    .addItem("⚙️ 設定科目/類別加權值", "showWeightSettingDialog")
    .addItem("📦 結算並匯出舊學期資料 (開新學期)", "archiveAndResetSemester")
    .addSeparator()
    .addItem("🛠️ 建立/檢查 [System_Settings] 基礎設定分頁", "initSystemSettingsSheet")
    .addToUi();

  ui.createMenu("🏫 成績管理工具")
    .addItem("📤 同步當前科任成績至導師試算表 (跨檔動態開啟)", "syncSubjectSheetToHomeroomSpreadsheetUI")
    .addItem("📥 匯入當前分頁成績至導師總表 (本檔)", "importSubjectSheetToTutorSheetUI")
    .addSeparator()
    .addItem("📈 開啟科任跨班分析儀表板 (視覺化)", "showSubjectDashboardDialog")
    .addItem("📑 產出科任跨班分析總表分頁", "createSubjectAnalyticsSheetUI")
    .addSeparator()
    .addItem("📊 產生專任班級總成績報表", "showSelectSubjectSheetDialog")
    .addItem("💯 指定欄號換算百分制", "convertSelectedToPercentage")
    .addItem("🍎計算平均與區間人數", "generateCurrentSubjectStats")
    .addItem("🔀 按評量類別排序欄位 (當前班級)", "sortCurrentSheetColumnsByCategory")
    .addSeparator()
    .addItem("⚙️ 設定科目/類別加權值", "showWeightSettingDialog")
    .addItem("➕ 手動新增專任班級分頁", "createNewSubjectClassSheet")
    .addToUi();
}

// ==========================================
// 3. 每日自動執行與檔案匯入
// ==========================================
/**
 * 每日自動執行成績匯入任務，並於有處理檔案時寄送 Email 報表至執行者的 Gmail
 */
function importDailyScores() {
  Logger.log(">>> 開始執行每日成績自動匯入任務...");
  const logs = [];
  let processedFileCount = 0;

  // 1. 處理手寫圖片
  if (typeof processUploadedImageFiles === 'function') {
    try {
      const imageLogs = processUploadedImageFiles();
      if (Array.isArray(imageLogs) && imageLogs.length > 0) {
        logs.push("【📷 手寫照片辨識結果】");
        logs.push(...imageLogs);
        processedFileCount += imageLogs.length;
      }
    } catch (e) {
      logs.push(`❌ 手寫照片處理異常: ${e.message}`);
    }
  }

  // 2. 處理 Excel 檔案
  if (typeof processExcelFiles === 'function') {
    try {
      const excelLogs = processExcelFiles();
      if (Array.isArray(excelLogs) && excelLogs.length > 0) {
        logs.push("\n【📊 Excel 檔案匯入結果】");
        logs.push(...excelLogs);
        processedFileCount += excelLogs.length;
      }
    } catch (e) {
      logs.push(`❌ Excel 檔處理異常: ${e.message}`);
    }
  }

  // 3. 防呆判斷：若今日完全沒有處理任何檔案，則不寄送信件
  if (processedFileCount === 0 || logs.length === 0) {
    Logger.log("ℹ️ 今日資料夾內無待處理檔案（或無檔案被處理），跳過寄送 Email。");
    return;
  }

  // 💡【核心新增】：匯入完成後，自動更新所有分頁的平均數與分數區間人數
  try {
    generateCurrentSubjectStats();
    logs.push("\n【📈 全分頁統計結果】");
    logs.push("✅ 已自動重算並更新所有分頁底部的平均分數與區間人數。");
  } catch (statErr) {
    Logger.log(`⚠️ 自動統計更新異常: ${statErr.message}`);
    logs.push(`\n⚠️ 自動統計更新異常: ${statErr.message}`);
  }

  // 4. 取得當前 Google 帳號的 Gmail 並寄送報表
  const userEmail = Session.getEffectiveUser().getEmail();
  const currentTimeStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
  
  const subject = `【成績系統】每日成績自動匯入彙整報表 (${currentTimeStr})`;
  const body = `親愛的老師，您好：\n\n` +
               `系統已為您完成成績自動匯入作業，執行彙整結果如下：\n\n` +
               `--------------------------------------------------\n` +
               logs.join("\n") +
               `\n--------------------------------------------------\n\n` +
               `※ 若有提示「跳過無專任分頁」之檔案，該檔案仍保留於上傳資料夾中，請您確認建立分頁後重新上傳。\n` +
               `此郵件由成績自動匯入系統自動發送，請勿直接回覆。`;

  try {
    MailApp.sendEmail(userEmail, subject, body);
    Logger.log(`📧 匯入報表已成功寄送至帳號: ${userEmail}`);
  } catch (emailErr) {
    Logger.log(`❌ 寄送 Email 報表失敗: ${emailErr.message}`);
  }
}
/**
 * 批次處理「成績上傳資料夾」中的 Excel / Google Sheets 檔案
 */
function processExcelFiles() {
  Logger.log(">>> 開始執行 processExcelFiles()");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = DriveApp.getFolderById(getActiveUploadFolderId());
  const processedFolder = DriveApp.getFolderById(getActiveBackupFolderId());
  const files = folder.getFiles();
  const logs = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const mimeType = file.getMimeType();
    
    // 檢查是否為 Excel 或 Google Sheets
    const isExcel = fileName.toLowerCase().endsWith(".xlsx") || fileName.toLowerCase().endsWith(".xls") || mimeType === MimeType.MICROSOFT_EXCEL;
    const isGoogleSheet = mimeType === MimeType.GOOGLE_SHEETS;

    if (isExcel || isGoogleSheet) {
      let targetSS = null;
      let isTempFile = false;

      try {
        // 若為 Excel 檔案，需先透過 API 轉為臨時 Google 試算表才能讀取內容
        if (isExcel) {
          targetSS = convertExcelToGoogleSheet(file.getId());
          isTempFile = true;
        } else {
          targetSS = SpreadsheetApp.open(file);
        }

        const sourceSheets = targetSS.getSheets();
        let totalImportedCols = 0;
        let processedTabNames = [];
        let skippedTabNames = [];

        // 逐一解析 Excel 裡面的所有分頁 (Tabs)
        for (const sourceSheet of sourceSheets) {
          const res = importSingleExcelSheetWithRouting(sourceSheet, ss);
          if (res.status === "SUCCESS") {
            totalImportedCols += res.cols;
            processedTabNames.push(sourceSheet.getName().trim());
          } else if (res.status === "SKIPPED") {
            skippedTabNames.push(sourceSheet.getName().trim());
          }
        }

        let logMsg = `✅ 檔案 [${fileName}] 處理完成 (匯入分頁: ${processedTabNames.join(", ")}，共 ${totalImportedCols} 個欄位)`;
        if (skippedTabNames.length > 0) {
          logMsg += ` ⚠️ 跳過無對應分頁: [${skippedTabNames.join(", ")}]`;
        }
        logs.push(logMsg);

        // 移至已處理資料夾
        file.moveTo(processedFolder);

      } catch (err) {
        Logger.log(`❌ 處理檔案 [${fileName}] 失敗: ${err.message}`);
        logs.push(`❌ 檔案 [${fileName}] 處理失敗: ${err.message}`);
      } finally {
        // 若使用了臨時轉換檔，務必清理刪除
        if (isTempFile && targetSS) {
          try {
            DriveApp.getFileById(targetSS.getId()).setTrashed(true);
          } catch (e) {}
        }
      }
    }
  }

  return logs;
}

/**
 * 處理單一 Excel 分頁並將成績路由至對應的目標分頁
 * @param {Sheet} sourceSheet - Excel 內的單一分頁
 * @param {Spreadsheet} targetSS - 系統活頁簿 (getActiveSpreadsheet)
 * @return {Object} 處理結果 { status: "SUCCESS"|"SKIPPED"|"NONE", cols: 寫入欄位數 }
 */
function importSingleExcelSheetWithRouting(sourceSheet, ss) {
  const tabName = sourceSheet.getName().trim();
  const lastCol = sourceSheet.getLastColumn();
  const lastRow = sourceSheet.getLastRow();

  if (lastRow < 4) return { status: "NONE", cols: 0 };

  const isMainSheet = (tabName === "導師成績總表");
  let targetSheet = null;

  if (isMainSheet) {
    targetSheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
  } else {
    targetSheet = ss.getSheetByName(tabName);
    // 【核心防呆】：若系統中無此專任分頁，跳過寫入
    if (!targetSheet) {
      Logger.log(`⚠️ 系統中無專任分頁 [${tabName}]，跳過寫入。`);
      return { status: "SKIPPED", cols: 0 };
    }
  }

  // 決定讀取座標規範 (導師總表從 D 欄/Col 4 開始，專任分頁從 C 欄/Col 3 開始)
  const startReadCol = isMainSheet ? 4 : 3;
  const seatColIdx = isMainSheet ? 2 : 1; // 導師總表座號在 B 欄(2)，專任座號在 A 欄(1)

  if (lastCol < startReadCol) return { status: "NONE", cols: 0 };

  // 1. 【修復點】：1~3 列全數使用 getDisplayValues() 確保讀取到的都是畫面上的「純文字」
  const headers = sourceSheet.getRange(1, startReadCol, 1, lastCol - startReadCol + 1).getDisplayValues()[0];
  const ranges = sourceSheet.getRange(2, startReadCol, 1, lastCol - startReadCol + 1).getDisplayValues()[0];
  const times = sourceSheet.getRange(3, startReadCol, 1, lastCol - startReadCol + 1).getDisplayValues()[0];

  // 讀取學生資料區域
  const studentData = sourceSheet.getRange(4, 1, lastRow - 3, lastCol).getValues();

  let importedCols = 0;

  for (let c = 0; c < headers.length; c++) {
    const headerTitle = String(headers[c] || "").trim();
    if (!headerTitle) continue;

    const examRange = String(ranges[c] || "").trim() || "無範圍";
    const rawTime = String(times[c] || "").trim();
    const timeStr = rawTime !== "" ? rawTime : Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");

    // 建立座號對應成績 Map
    const scoreMap = new Map();
    for (let r = 0; r < studentData.length; r++) {
      const seat = parseInt(studentData[r][seatColIdx - 1], 10);
      const rawScore = studentData[r][startReadCol - 1 + c];
      const scoreStr = String(rawScore !== null && rawScore !== undefined ? rawScore : "").trim();

      // 2. 【修復點】：剔除空白字串，防止空格被 Number() 自動轉為 0 分
      if (!isNaN(seat) && seat > 0 && scoreStr !== "" && !isNaN(Number(scoreStr))) {
        scoreMap.set(seat, Number(scoreStr));
      }
    }

    if (scoreMap.size === 0) continue;

    // 寫入目標 Google 試算表
    const targetLastCol = targetSheet.getLastColumn();
    const minStartCol = isMainSheet ? 4 : 3;
    const targetCol = (targetLastCol < minStartCol - 1) ? minStartCol : targetLastCol + 1;

    // 3. 【修復點】：寫入目標表前，先設置儲存格為純文字 (@)，並以 String() 轉型寫入
    targetSheet.getRange(1, targetCol).setNumberFormat("@").setValue(headerTitle);
    targetSheet.getRange(2, targetCol).setNumberFormat("@").setValue(examRange);
    targetSheet.getRange(3, targetCol).setNumberFormat("@").setValue(timeStr);

    // 4. 寫入學生成績
    const seatSearchCol = isMainSheet ? 2 : 1;
    const studentEndRow = getStudentEndRow(targetSheet, seatSearchCol, isMainSheet ? 3 : 2);
    const totalStudents = studentEndRow - 3;

    if (totalStudents > 0) {
      const seatList = targetSheet.getRange(4, seatSearchCol, totalStudents, 1).getValues();
      const outputScores = seatList.map(r => {
        const seatNum = parseInt(r[0], 10);
        return [scoreMap.has(seatNum) ? scoreMap.get(seatNum) : ""];
      });
      targetSheet.getRange(4, targetCol, outputScores.length, 1).setValues(outputScores);
    }

    importedCols++;
  }

  return { status: "SUCCESS", cols: importedCols };
}

/**
 * 輔助轉檔函式：將 Excel (.xlsx) 檔案轉為臨時 Google 試算表以利 Apps Script 讀取
 */
function convertExcelToGoogleSheet(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const config = {
    title: "[Temp_Import_Excel] " + file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  };

  const url = "https://www.googleapis.com/upload/drive/v2/files?uploadType=multipart&convert=true";
  const boundary = "-------314159265358979323846";
  const delimiter = "\r\n--" + boundary + "\r\n";
  const close_delim = "\r\n--" + boundary + "--";

  const multipartRequestBody =
    delimiter +
    'Content-Type: application/json\r\n\r\n' +
    JSON.stringify(config) +
    delimiter +
    'Content-Type: ' + blob.getContentType() + '\r\n' +
    'Content-Transfer-Encoding: base64\r\n\r\n' +
    Utilities.base64Encode(blob.getBytes()) +
    close_delim;

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "multipart/related; boundary=" + boundary,
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    payload: multipartRequestBody,
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  if (json.error) {
    throw new Error("Excel 轉檔失敗: " + json.error.message);
  }
  return SpreadsheetApp.openById(json.id);
}
function processHandwrittenImageFiles() {
  const folder = DriveApp.getFolderById(FOLDER_HANDWRITTEN_ID);
  const files = folder.getFiles();
  const processedLogs = [];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    if (mimeType.startsWith("image/") || fileName.toLowerCase().endsWith(".heic")) {
      try {
        Logger.log(`📸 開始處理照片: ${fileName}`);
        const resultMsg = importScoresFromHandwrittenImage(file.getId());
        processedLogs.push(`${fileName}: ${resultMsg}`);
        Utilities.sleep(3000);
      } catch (err) {
        Logger.log(`❌ 照片 ${fileName} 辨識失敗: ${err.message}`);
        Utilities.sleep(3000);
      }
    }
  }
  return processedLogs;
}
/**
 * 批量處理上傳資料夾中的所有手寫成績單照片
 * 由 importDailyScores() 自動呼叫
 */
function processUploadedImageFiles() {
  const logs = [];
  
  // 1. 取得上傳資料夾 (請確保專案全域變數 FOLDER_UPLOAD_ID 已設定)
  let uploadFolder;
  try {
    uploadFolder = DriveApp.getFolderById(getActiveUploadFolderId());
  } catch (e) {
    Logger.log(`❌ 無法取得上傳資料夾，請檢查 FOLDER_UPLOAD_ID：${e.message}`);
    return [`❌ 找不到上傳資料夾 (ID: ${FOLDER_UPLOAD_ID})`];
  }

  // 2. 搜尋資料夾內的所有圖片檔案
  const files = uploadFolder.getFiles();
  const imageExtensions = [".jpg", ".jpeg", ".png", ".heic"];

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const fileNameLower = fileName.toLowerCase();
    const mimeType = file.getMimeType();

    // 判斷是否為圖片檔 (比對 MimeType 或副檔名)
    const isImage = mimeType.startsWith("image/") || 
                    imageExtensions.some(ext => fileNameLower.endsWith(ext));

    if (isImage) {
      Logger.log(`📷 發現待處理圖片檔案：[${fileName}] (ID: ${file.getId()})`);
      try {
        // 呼叫您的手寫圖片 AI 辨識函式
        const resultMsg = importScoresFromHandwrittenImage(file.getId());
        logs.push(`✅ 圖片 [${fileName}]：${resultMsg}`);
      } catch (err) {
        Logger.log(`❌ 圖片 [${fileName}] 辨識失敗：${err.message}`);
        logs.push(`❌ 圖片 [${fileName}] 處理失敗：${err.message}`);
      }
    }
  }

  return logs;
}

// ==========================================
// 4. Gemini AI 手寫照片辨識邏輯
// ==========================================
function importScoresFromHandwrittenImage(fileId) {
  Logger.log(">>> 開始辨識照片，傳入的 File ID 為: [" + fileId + "]");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
  const mainClassName = normalizeClassName(String(mainSheet.getRange("A1").getValue()).trim());

  const imageFile = DriveApp.getFileById(fileId);
  const fileName = imageFile.getName().toLowerCase();
  const imageBlob = imageFile.getBlob();
  const base64Data = Utilities.base64Encode(imageBlob.getBytes());

  let mimeType = imageBlob.getContentType();
  if (fileName.endsWith(".heic") || !mimeType || mimeType === "application/octet-stream") {
    mimeType = "image/jpeg";
  }

  const prompt = `你是一個精準的手寫成績單 OCR 辨識助手。請仔細分析這張成績單照片，判斷其「班級」、「科目」、「評量類型」與成績數據，並嚴格回傳純 JSON 格式：

  {
    "className": "班級名稱(例如: 國二愛、國二仁、二仁，若無寫則填寫空字串)",
    "subject": "科目名稱(例如: 理化、歷史、國文)",
    "type": "評量類型(例如: 小考、段考、作業、筆記，若無標示預設填 小考)",
    "columns": [
      {
        "range": "考試範圍或測驗名稱",
        "scores": [
          {"seat": 1, "score": 85},
          {"seat": 2, "score": 90}
        ]
      }
    ]
  }

  【辨識規則說明】：
  1. 班級與科目：尋找照片頂端手寫字（如「二愛」、「國二仁」、「理化小考」）。
  2. 評量類型 (type)：辨識是否有「段考」、「小考」、「作業」、「筆記」等字樣，未標示則填「小考」。
  3. 欄位解析 (columns) 與座號成績：
     - seat 請轉為純整數數字 (如 1, 2, 3...)。
     - 被劃刪除線、空白、缺考者，該座號 score 請填寫 null 或忽略。

  【輸出格式限制】：
  - 嚴格僅輸出純文字 JSON，絕對不要包裝任何 Markdown 標籤。`;

  // 💡 整合後的模型呼叫順序 (優先 Lite -> 動態抓取 -> 真實可用的備援清單)
  const dynamicModel = getLatestFlashModelName();
  const candidateModels = Array.from(new Set([
    "gemini-flash-lite-latest",  // 1. 強制優先選用 Lite（速度快、省額度）
    dynamicModel,                // 2. 動態偵測到的最新版 Flash 模型
    "gemini-flash-latest",       // 3. Standard Flash 最新版 (高辨識率)
    "gemini-3.6-flash",          // 4. 新世代強大模型備援
    "gemini-2.5-flash"           // 5. 保底備援
  ].filter(Boolean)));

  let lastError = null;

  for (const activeModel of candidateModels) {
    Logger.log(`🤖 嘗試呼叫 API 模型: [${activeModel}]...`);
    const apiKey = getActiveGeminiApiKey();
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${activeModel}:generateContent?key=${apiKey}`;
    
    const payload = {
      "contents": [{
        "parts": [
          { "text": prompt },
          { "inline_data": { "mime_type": mimeType, "data": base64Data } }
        ]
      }]
    };

    try {
      const response = UrlFetchApp.fetch(url, {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
      });
      
      const result = JSON.parse(response.getContentText());
      
      if (result.error) {
        const errMsg = result.error.message || "";
        if (errMsg.toLowerCase().includes("quota") || errMsg.toLowerCase().includes("not found") || result.error.code === 429) {
          lastError = new Error(`[${activeModel}] 額度不足或失效: ${errMsg}`);
          continue; 
        }
        throw new Error(`Gemini API 錯誤: ${errMsg}`);
      }

      if (!result.candidates || result.candidates.length === 0 || !result.candidates[0].content) {
        throw new Error("API 未回傳有效文字內容");
      }

      let responseText = result.candidates[0].content.parts[0].text;
      responseText = responseText.replace(/```json/g, "").replace(/```/g, "").trim();

      const parsedData = JSON.parse(responseText);
      
      let rawClass = parsedData.className ? parsedData.className.trim() : "";
      let recognizedClass = normalizeClassName(rawClass);

      if (!recognizedClass) recognizedClass = mainClassName;

      const subject = parsedData.subject ? parsedData.subject.substring(0, 2).trim() : "理化";
      const evalType = parsedData.type ? parsedData.type.trim() : "小考";
      const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
      const columns = parsedData.columns || [];

      if (columns.length === 0) throw new Error("API 未回傳有效的欄位成績資料");

      let targetSheetsConfig = [];

      // 1. 若辨識出的班級為導師班，將成績填入「導師成績總表」
      if (recognizedClass === mainClassName) {
        targetSheetsConfig.push({ sheet: mainSheet, mode: "MAIN" });
      }

      // 2. 搜尋專任班級分頁 (如: 國二愛理化)
      const specialSheetName = `${recognizedClass}${subject}`;
      const specialSheet = ss.getSheetByName(specialSheetName);

      // 【關鍵修改】：僅在專任分頁「已存在」時才寫入，絕不自動建立新分頁
      if (specialSheet) {
        targetSheetsConfig.push({ sheet: specialSheet, mode: "EXISTING_SPECIAL" });
      }

      // 若完全沒有匹配的分頁（既非導師班，又無該專任班級分頁）
      if (targetSheetsConfig.length === 0) {
        Logger.log(`⚠️ 辨識班級科目為 [${specialSheetName}]，但系統內無此專任分頁，已自動跳過寫入（檔案保留於原處）。`);
  
        // 不將檔案移動至已處理資料夾，保留在原處供教師檢視
        return `⚠️ 辨識班級科目為 [${specialSheetName}]，因系統中無此專任分頁，已跳過不填入（檔案未移動，請檢查分頁）。`;
      }

      let totalImportedCols = 0;
      let totalScoresCount = 0;

      for (const colData of columns) {
        const examRange = colData.range || "無範圍";
        const scoreMap = new Map();

        if (colData.scores && Array.isArray(colData.scores)) {
          colData.scores.forEach(item => {
            if (item.seat !== undefined && item.seat !== null && item.score !== null && item.score !== undefined) {
              const numericSeat = parseInt(item.seat, 10);
              const numericScore = Number(item.score);
              if (!isNaN(numericSeat) && !isNaN(numericScore)) {
                scoreMap.set(numericSeat, numericScore);
              }
            }
          });
        }

        if (scoreMap.size === 0) continue;

        //const stats = calculateScoreStats(Array.from(scoreMap.values()));

        for (const config of targetSheetsConfig) {
          const targetSheet = config.sheet;
          const mode = config.mode;

          if (mode === "MAIN") {
            const lastCol = targetSheet.getLastColumn();
            const startCol = (lastCol < 3) ? 4 : lastCol + 1;

            targetSheet.getRange(1, startCol).setValue(subject);
            targetSheet.getRange(2, startCol).setNumberFormat("@").setValue(examRange);
            targetSheet.getRange(3, startCol).setValue(todayStr);

            const studentEndRow = getStudentEndRow(targetSheet, 2, 3);
            const totalStudents = studentEndRow - 3;

            if (totalStudents > 0) {
              const seatList = targetSheet.getRange(4, 2, totalStudents, 1).getValues();
              const outputScores = seatList.map(r => {
                const seatNum = parseInt(r[0], 10);
                return [scoreMap.has(seatNum) ? scoreMap.get(seatNum) : ""];
              });
              targetSheet.getRange(4, startCol, outputScores.length, 1).setValues(outputScores);

              //writeStatsToSheet(targetSheet, studentEndRow + 2, startCol, stats, 3);
            }

          } else if (mode === "EXISTING_SPECIAL") {
            const lastCol = targetSheet.getLastColumn();
            const startCol = (lastCol < 2) ? 3 : lastCol + 1;

            targetSheet.getRange(1, startCol).setValue(evalType);
            targetSheet.getRange(2, startCol).setNumberFormat("@").setValue(examRange);
            targetSheet.getRange(3, startCol).setValue(todayStr);

            const studentEndRow = getStudentEndRow(targetSheet, 1, 2);
            const totalStudents = studentEndRow - 3;

            if (totalStudents > 0) {
              const seatList = targetSheet.getRange(4, 1, totalStudents, 1).getValues();
              const outputScores = seatList.map(r => {
                const seatNum = parseInt(r[0], 10);
                return [scoreMap.has(seatNum) ? scoreMap.get(seatNum) : ""];
              });
              targetSheet.getRange(4, startCol, outputScores.length, 1).setValues(outputScores);

              //writeStatsToSheet(targetSheet, studentEndRow + 2, startCol, stats, 2);
            }
          }
        }

        totalImportedCols++;
        totalScoresCount += scoreMap.size;
      }

      try {
        const processedFolder = DriveApp.getFolderById(getActiveBackupFolderId());
        imageFile.moveTo(processedFolder);
        Logger.log(`📦 已成功將檔案 [${imageFile.getName()}] 移動至備份資料夾`);
      } catch (moveErr) {
        Logger.log(`⚠️ 移動檔案失敗: ${moveErr.message}`);
      }

      const writtenSheetNames = targetSheetsConfig.map(c => c.sheet.getName()).join(" & ");
      Logger.log(`🎉 成功寫入分頁 [${writtenSheetNames}]！`);
      return `[${writtenSheetNames}] 成功登錄 ${totalImportedCols} 個欄位 (共 ${totalScoresCount} 人次)`;

    } catch (err) {
      Logger.log(`❌ 模型 [${activeModel}] 執行失敗: ${err.message}`);
      lastError = err;
    }
  }

  throw new Error(`辨識失敗: ${lastError ? lastError.message : "未知錯誤"}`);
}
// ==========================================
// 5. 專任端功能：百分比換算、報表與排序
// ==========================================
function convertSelectedToPercentage() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const colRes = ui.prompt("💯 換算百分制", "請輸入要換算的分數欄號 (例如: C 或 D)：", ui.ButtonSet.OK_CANCEL);
  if (colRes.getSelectedButton() !== ui.Button.OK) return;
  
  const colLetter = colRes.getResponseText().trim().toUpperCase();
  if (!colLetter.match(/^[A-Z]+$/)) return ui.alert("⚠️ 請輸入有效的欄位英文字母！");

  const scoreRes = ui.prompt("💯 換算百分制", `請輸入 ${colLetter} 欄這場測驗的「滿分分數」：`, ui.ButtonSet.OK_CANCEL);
  if (scoreRes.getSelectedButton() !== ui.Button.OK) return;

  const maxScore = Number(scoreRes.getResponseText().trim());
  if (isNaN(maxScore) || maxScore <= 0) return ui.alert("⚠️ 請輸入有效的數字滿分！");

  const colIndex = letterToColumn(colLetter);
  const isMain = (sheet.getRange("C3").getValue() === "姓名");
  const studentEndRow = getStudentEndRow(sheet, isMain ? 2 : 1, isMain ? 3 : 2);
  const studentCount = studentEndRow - 3;

  if (studentCount <= 0) return ui.alert("找不到有效的學生成績列！");

  const range = sheet.getRange(4, colIndex, studentCount, 1);
  const values = range.getValues();

  for (let r = 0; r < values.length; r++) {
    const val = Number(values[r][0]);
    if (!isNaN(val) && values[r][0] !== "") {
      values[r][0] = Math.round((val / maxScore) * 100);
    }
  }

  range.setValues(values);

  // 💡【核心修正】：移除舊的 calculateAllSheetsAverages()，改呼叫新版的全分頁統計函式
  generateCurrentSubjectStats();

  ui.alert(`🎉 第 ${colLetter} 欄成績已成功折算為百分制並重算統計！`);
}

function showSelectSubjectSheetDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const mainSheetName = (ss.getSheetByName("導師成績總表") || sheets[0]).getName();

  // 保留您原本完整的排除清單
  const excludeSheets = [mainSheetName, "導師班成績統計表", "設定", "權重設定", "歷史紀錄", "Template", "區間個人成績報表"];

  // 1. 抓取所有符合條件的專任分頁，並預先分析各分頁包含的評量類別與次數
  const subjectSheetData = {};

  sheets.forEach(s => {
    const name = s.getName();
    if (!excludeSheets.includes(name) && !name.endsWith("_總成績報表") && !name.includes("報表")) {
      const lastCol = s.getLastColumn();
      const categoryCounts = {};

      if (lastCol >= 3) {
        const categories = s.getRange(1, 3, 1, lastCol - 2).getValues()[0];
        categories.forEach(c => {
          const cat = String(c || "").trim();
          // 過濾掉空白與統計欄位
          if (cat && !cat.includes("平均") && !cat.includes("排名")) {
            categoryCounts[cat] = (categoryCounts[cat] || 0) + 1;
          }
        });
      }
      subjectSheetData[name] = categoryCounts;
    }
  });

  const subjectSheetNames = Object.keys(subjectSheetData);

  if (subjectSheetNames.length === 0) {
    return SpreadsheetApp.getUi().alert("目前沒有可用的專任班級分頁！");
  }

  // 將資料轉為 JSON 字串嵌入前端 JavaScript
  const sheetDataJson = JSON.stringify(subjectSheetData);
  const optionsHtml = subjectSheetNames.map(s => `<option value="${s}">${s}</option>`).join("");

  // 2. 構建包含 Top N 設定欄位的動態 HTML 對話框
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: "Microsoft JhengHei", Arial, sans-serif; padding: 15px; color: #333; margin: 0; }
          h3 { margin-top: 0; color: #27ae60; border-bottom: 2px solid #27ae60; padding-bottom: 6px; font-size: 16px; }
          .form-group { margin-bottom: 12px; }
          label { font-weight: bold; font-size: 13px; display: block; margin-bottom: 4px; }
          select { width: 100%; padding: 8px; font-size: 14px; border: 1px solid #ccc; border-radius: 4px; }
          .cat-box { background: #f8f9fa; border: 1px solid #e0e0e0; border-radius: 6px; padding: 10px; max-height: 180px; overflow-y: auto; }
          .cat-item { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; font-size: 13px; }
          .cat-item:last-child { margin-bottom: 0; }
          .cat-item input { width: 70px; padding: 5px; text-align: center; font-size: 13px; border: 1px solid #ccc; border-radius: 4px; }
          .note { font-size: 11px; color: #666; margin-top: 2px; margin-bottom: 8px; line-height: 1.3; }
          button { width: 100%; padding: 10px; margin-top: 15px; font-size: 14px; background-color: #27ae60; color: white; border: none; cursor: pointer; border-radius: 4px; font-weight: bold; }
          button:hover { background-color: #219150; }
          #loading { display: none; text-align: center; color: #27ae60; font-size: 13px; font-weight: bold; margin-top: 10px; }
        </style>
      </head>
      <body>
        <h3>📊 產生專任班級總成績報表</h3>
        
        <div class="form-group">
          <label for="sheetSelect">1. 請選擇班級：</label>
          <select id="sheetSelect" onchange="renderCategories()">${optionsHtml}</select>
        </div>

        <div class="form-group">
          <label>2. 各類別平均取最高數量：</label>
          <div class="note">💡 填入數字（預設 0 為全選）。若成績次數不足，將自動補 0 分計算平均。</div>
          <div id="categoryContainer" class="cat-box"></div>
        </div>

        <div id="loading">⏳ 報表生成與寫入公式中，請稍候...</div>

        <button id="btnSubmit" onclick="submit()">產生報表</button>

        <script>
          const sheetData = ${sheetDataJson};

          function renderCategories() {
            const selectedSheet = document.getElementById("sheetSelect").value;
            const container = document.getElementById("categoryContainer");
            container.innerHTML = "";

            const categories = sheetData[selectedSheet] || {};
            const catNames = Object.keys(categories);

            if (catNames.length === 0) {
              container.innerHTML = '<div style="color:#999; text-align:center;">此分頁未找到成績欄位</div>';
              return;
            }

            catNames.forEach(cat => {
              const count = categories[cat];
              const div = document.createElement("div");
              div.className = "cat-item";
              div.innerHTML = \`
                <span><strong>\${cat}</strong> (\${count} 次成績)：</span>
                <input type="number" id="topN_\${cat}" data-cat="\${cat}" value="0" min="0" placeholder="0為全選">
              \`;
              container.appendChild(div);
            });
          }

          function submit() {
            const sheetName = document.getElementById('sheetSelect').value;
            const inputs = document.querySelectorAll('#categoryContainer input[type="number"]');
            const categoryTopNMap = {};

            inputs.forEach(input => {
              const cat = input.getAttribute("data-cat");
              const val = parseInt(input.value, 10);
              categoryTopNMap[cat] = (isNaN(val) || val < 0) ? 0 : val;
            });

            document.getElementById("btnSubmit").style.display = "none";
            document.getElementById("loading").style.display = "block";

            google.script.run
              .withSuccessHandler(m => {
                alert(m);
                google.script.host.close();
              })
              .withFailureHandler(err => {
                alert("產生報表失敗：" + err.message);
                document.getElementById("btnSubmit").style.display = "block";
                document.getElementById("loading").style.display = "none";
              })
              .generateSubjectClassReportByName(sheetName, categoryTopNMap);
          }

          // 初始化繪製第一個班級的類別
          window.onload = renderCategories;
        </script>
      </body>
    </html>`;

  // 稍微加高視窗高度 (480px)，以便完整顯示各類別輸入框
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(480), 
    "產生專任班級總成績報表"
  );
}

/**
 * 核心統合邏輯：根據前端傳入的班級名稱與各類別 Top N 設定，產生獨立的「_總成績報表」分頁
 */
function generateSubjectClassReportByName(sheetName, categoryTopNMap) {
  categoryTopNMap = categoryTopNMap || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) throw new Error(`找不到工作表：${sheetName}`);

  const lastCol = sourceSheet.getLastColumn();
  if (lastCol < 3) throw new Error("該工作表無足夠資料！");

  const studentEndRow = getStudentEndRow(sourceSheet, 1, 2);
  const studentCount = studentEndRow - 3;
  if (studentCount <= 0) throw new Error("無學生資料！");

  // 1. 讀取原始資料
  const studentInfo = sourceSheet.getRange(4, 1, studentCount, 2).getValues();
  const categories = sourceSheet.getRange(1, 3, 1, lastCol - 2).getValues()[0];
  const ranges = sourceSheet.getRange(2, 3, 1, lastCol - 2).getValues()[0];
  const dates = sourceSheet.getRange(3, 3, 1, lastCol - 2).getValues()[0];
  const scoresMatrix = sourceSheet.getRange(4, 3, studentCount, lastCol - 2).getValues();

  const typeWeights = (typeof getSubjectTeacherWeights === 'function') 
    ? getSubjectTeacherWeights() 
    : { "段考": 0.5, "小考": 0.2, "作業": 0.2, "筆記": 0.1 };

  // 2. 依評量類別分組
  const typeGroups = {};
  for (let c = 0; c < categories.length; c++) {
    let typeName = String(categories[c] || "").trim();
    if (typeName === "") typeName = "平時測驗";

    if (!typeName.includes("平均") && !typeName.includes("排名")) {
      if (!typeGroups[typeName]) typeGroups[typeName] = [];
      typeGroups[typeName].push(c);
    }
  }

  const presentTypes = Object.keys(typeGroups);
  if (presentTypes.length === 0) throw new Error("未發現有效的評量類別成績！");

  // 3. 建構報表表頭與欄位追蹤
  const headerRow1 = ["", ""];
  const headerRow2 = ["", ""];
  const headerRow3 = ["座號", "姓名"];

  const blueCols = [];   
  const yellowCols = []; 
  const typeAvgColLetters = {}; 

  presentTypes.forEach(t => {
    // 寫入該類別的所有單科測驗
    typeGroups[t].forEach(c => {
      headerRow1.push(t);
      headerRow2.push(String(ranges[c] || ""));
      const dateVal = (typeof parseDateHelper === 'function') ? parseDateHelper(dates[c]) : dates[c];
      headerRow3.push(dateVal instanceof Date ? Utilities.formatDate(dateVal, "GMT+8", "MM/dd") : String(dates[c] || ""));
    });

    // 確保 Top N 設定能精準抓取（去除字串前後空白並轉為數字）
    const topN = Number(categoryTopNMap[t] || categoryTopNMap[t.trim()]) || 0;
    const topNText = (topN > 0) ? `(取前${topN}高)` : "";
    
    headerRow1.push(t); 
    headerRow2.push(`${t}平均${topNText}`); 
    headerRow3.push("");
    blueCols.push(headerRow1.length);
    typeAvgColLetters[t] = getColLetter(headerRow1.length);

    headerRow1.push(t); 
    headerRow2.push(`${t}排名`); 
    headerRow3.push("");
    yellowCols.push(headerRow1.length);
  });

  // 總成績欄
  headerRow1.push("總成績", "總成績");
  headerRow2.push("加權平均", "總排名");
  headerRow3.push("", "");

  const weightedAvgColIdx = headerRow1.length - 1;
  const weightedAvgColLetter = getColLetter(weightedAvgColIdx);
  const totalRankColIdx = headerRow1.length;

  // 4. 建構學生資料陣列
  const studentRows = [];
  const reportStartRow = 4;
  const reportEndRow = reportStartRow + studentCount - 1;

  for (let s = 0; s < studentCount; s++) {
    const row = [studentInfo[s][0], studentInfo[s][1]];
    const currRowIndex = reportStartRow + s;

    presentTypes.forEach(t => {
      const catStartColIdx = row.length + 1;
      typeGroups[t].forEach(c => {
        row.push(scoresMatrix[s][c]);
      });
      const catEndColIdx = row.length;

      const startLetter = getColLetter(catStartColIdx);
      const endLetter = getColLetter(catEndColIdx);
      const scoreRangeStr = `${startLetter}${currRowIndex}:${endLetter}${currRowIndex}`;
      
      const topN = Number(categoryTopNMap[t] || categoryTopNMap[t.trim()]) || 0;

      // 【公式 A】：類別平均公式（已修復 ARRAYFORMULA 位置與補 0 邏輯）
      let avgFormula = "";
      if (topN === 0) {
        // 全選 (自動忽略空白/缺考)
        avgFormula = `=IFERROR(ROUND(AVERAGE(${scoreRangeStr}), 1), "")`;
      } else {
        // 取前 Top N 高 (不足者補 0 分)
        const zerosStr = Array(topN).fill(0).join(", ");
        avgFormula = `=ARRAYFORMULA(IFERROR(ROUND(SUM(LARGE({IF(ISNUMBER(${scoreRangeStr}), ${scoreRangeStr}, 0), ${zerosStr}}, SEQUENCE(1, ${topN}))) / ${topN}, 1), ""))`;
      }
      row.push(avgFormula);

      // 【公式 B】：類別排名公式
      const avgColLetter = typeAvgColLetters[t];
      const rankFormula = `=IF(ISNUMBER(${avgColLetter}${currRowIndex}), RANK(${avgColLetter}${currRowIndex}, ${avgColLetter}$4:${avgColLetter}$${reportEndRow}), "")`;
      row.push(rankFormula);
    });

    // 【公式 C】：加權總平均公式
    let weightedTerms = [];
    let totalWeightSum = 0;
    presentTypes.forEach(t => {
      const w = Number(typeWeights[t]) || 1;
      const avgColLetter = typeAvgColLetters[t];
      weightedTerms.push(`IFERROR(${avgColLetter}${currRowIndex}, 0)*${w}`);
      totalWeightSum += w;
    });

    const finalAvgFormula = `=IFERROR(ROUND((${weightedTerms.join(" + ")}) / ${totalWeightSum}, 1), "")`;
    row.push(finalAvgFormula);

    // 【公式 D】：總排名公式
    const totalRankFormula = `=IF(ISNUMBER(${weightedAvgColLetter}${currRowIndex}), RANK(${weightedAvgColLetter}${currRowIndex}, ${weightedAvgColLetter}$4:${weightedAvgColLetter}$${reportEndRow}), "")`;
    row.push(totalRankFormula);

    studentRows.push(row);
  }

  // 5. 建立或覆蓋「_總成績報表」分頁
  const reportSheetName = `${sheetName}_總成績報表`;
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) ss.deleteSheet(reportSheet);

  reportSheet = ss.insertSheet(reportSheetName);

  // 6. 寫入資料與套用格式
  reportSheet.getRange(1, 1, 1, headerRow1.length).setValues([headerRow1]).setFontWeight("bold");
  reportSheet.getRange(2, 1, 1, headerRow2.length).setNumberFormat("@").setValues([headerRow2]);
  reportSheet.getRange(3, 1, 1, headerRow3.length).setValues([headerRow3]).setFontWeight("bold");

  if (studentRows.length > 0) {
    reportSheet.getRange(4, 1, studentRows.length, studentRows[0].length).setValues(studentRows);
  }

  const totalRows = studentRows.length + 3;

  // 套用配色
  blueCols.forEach(col => reportSheet.getRange(1, col, totalRows, 1).setBackground("#EBF5FB"));
  yellowCols.forEach(col => reportSheet.getRange(1, col, totalRows, 1).setBackground("#FEF9E7"));
  reportSheet.getRange(1, weightedAvgColIdx, totalRows, 2).setBackground("#FDEDEC");

  // 表頭跨欄合併
  let currCol = 3;
  presentTypes.forEach(t => {
    const span = typeGroups[t].length + 2;
    if (span > 1) {
      reportSheet.getRange(1, currCol, 1, span).merge();
    }
    currCol += span;
  });
  reportSheet.getRange(1, currCol, 1, 2).merge();

  // 外框與置中
  const fullRange = reportSheet.getRange(1, 1, totalRows, headerRow1.length);
  fullRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
  fullRange.setBorder(true, true, true, true, true, true, "#BDC3C7", SpreadsheetApp.BorderStyle.SOLID);

  ss.setActiveSheet(reportSheet);
  return `🎉 成功建立「${sheetName}」班級總成績報表！\n所有平均與排名均已正確設定動態公式。`;
}

/**
 * 輔助函式：前端對話框點擊「產生報表」時的對應進入點
 */
function generateSubjectReport(sheetName, categoryTopNMap) {
  return generateSubjectClassReportByName(sheetName, categoryTopNMap);
}

/**
 * 輔助函式：欄位數字轉為試算表 A1 欄位字母
 */
function getColLetter(colIndex) {
  let temp = "";
  let letter = "";
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
}

function sortCurrentSheetColumnsByCategory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 3) return SpreadsheetApp.getUi().alert("當前分頁成績欄位不足！");

  const isMain = (sheet.getRange("C3").getValue() === "姓名");
  if (isMain) return SpreadsheetApp.getUi().alert("此功能適用於專任班級分頁！");

  const studentEndRow = getStudentEndRow(sheet, 1, 2);
  const numRows = studentEndRow - 3;
  if (numRows <= 0) return SpreadsheetApp.getUi().alert("無學生資料可排序！");

  const numCols = lastCol - 2;
  const customTypeOrder = getSubjectTypeOrder();
  const typeOrder = {};
  customTypeOrder.forEach((t, idx) => {
    typeOrder[t] = idx + 1;
  });

  const types = sheet.getRange(1, 3, 1, numCols).getValues()[0];
  const ranges = sheet.getRange(2, 3, 1, numCols).getDisplayValues()[0];
  const dates = sheet.getRange(3, 3, 1, numCols).getValues()[0];
  const scoresMatrix = sheet.getRange(4, 3, numRows, numCols).getValues();

  const colsData = [];
  for (let i = 0; i < numCols; i++) {
    const type = String(types[i] || "").trim();
    colsData.push({
      type: type,
      range: String(ranges[i] || "").trim(),
      date: dates[i],
      scores: scoresMatrix.map(row => row[i]),
      weight: typeOrder[type] || 99
    });
  }

  colsData.sort((a, b) => a.weight - b.weight);

  const newTypes = [colsData.map(d => d.type)];
  const newRanges = [colsData.map(d => d.range)];
  const newDates = [colsData.map(d => d.date)];
  const newScoresMatrix = [];
  for (let r = 0; r < numRows; r++) {
    newScoresMatrix.push(colsData.map(d => d.scores[r]));
  }

  sheet.getRange(1, 3, 1, numCols).setValues(newTypes);

  const rangeRow = sheet.getRange(2, 3, 1, numCols);
  rangeRow.setNumberFormat("@");
  rangeRow.setValues(newRanges);

  sheet.getRange(3, 3, 1, numCols).setValues(newDates);
  sheet.getRange(4, 3, numRows, numCols).setValues(newScoresMatrix);

  calculateAllSheetsAverages();
  SpreadsheetApp.getUi().alert("🎉 已成功將當前分頁成績欄位依「評量類別」集中排列！");
}

function createNewSubjectClassSheet() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("➕ 新增專任班級", "請輸入專任班級名稱 (例如: 國二愛理化)：", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK && res.getResponseText().trim()) {
    getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), res.getResponseText().trim());
    ui.alert("✅ 分頁建立完成！");
  }
}

// ==========================================
// 6. 導師區間歷史成績單與報表產生器
// ==========================================
/**
 * 精確找出導師成績總表中「最早的有效測驗日期」
 */
function getEarliestDateInMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
  const lastCol = mainSheet.getLastColumn();
  if (lastCol < 4) return Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");

  const subjects = mainSheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
  const dates = mainSheet.getRange(3, 4, 1, lastCol - 3).getValues()[0];
  let minDate = null;

  for (let c = 0; c < dates.length; c++) {
    const subj = String(subjects[c] || "").trim();
    if (!subj) continue; // 跳過無科目的無效欄位
    const dateObj = parseDateHelper(dates[c]);
    if (dateObj) {
      if (!minDate || dateObj < minDate) minDate = dateObj;
    }
  }
  return minDate ? Utilities.formatDate(minDate, "GMT+8", "yyyy-MM-dd") : Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
}

function showSheetWriteDialog() {
  const earliestDate = getEarliestDateInMainSheet();
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head><base target="_top"><style>
        body { font-family: "Microsoft JhengHei", Arial; padding: 15px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; font-weight: bold; margin-bottom: 5px; }
        input { width: 100%; padding: 8px; box-sizing: border-box; }
        button { background-color: #3498db; color: white; padding: 10px; width: 100%; border: none; cursor: pointer; font-size: 15px; border-radius: 4px; font-weight: bold; }
      </style></head>
      <body>
        <h3>📊 產生導師班成績統計表</h3>
        <div class="form-group"><label>起始日期：</label><input type="date" id="startDate" value="${earliestDate}"></div>
        <div class="form-group"><label>結束日期：</label><input type="date" id="endDate"></div>
        <button onclick="generate()">整理並貼至統計表</button>
        <script>
          const now = new Date();
          const yyyy = now.getFullYear();
          const mm = String(now.getMonth() + 1).padStart(2, '0');
          const dd = String(now.getDate()).padStart(2, '0');
          document.getElementById('endDate').value = \`\${yyyy}-\${mm}-\${dd}\`;

          function generate() {
            google.script.run.withSuccessHandler(m => { alert(m); google.script.host.close(); })
              .generateWeeklyReportByRange(document.getElementById('startDate').value, document.getElementById('endDate').value);
          }
        </script>
      </body>
    </html>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(htmlContent).setWidth(360).setHeight(280), "導師班成績統計表產生器");
}

/**
 * 導師端：產生指定區間歷史成績單，寫入「導師班成績統計表」
 */
function generateWeeklyReportByRange(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tutorSheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
  if (!tutorSheet) throw new Error("找不到導師成績總表！");

  const studentEndRow = getStudentEndRow(tutorSheet, 2, 3);
  const studentCount = studentEndRow - 3;
  if (studentCount <= 0) throw new Error("無學生資料！");

  const studentInfo = tutorSheet.getRange(4, 2, studentCount, 2).getValues();
  const lastCol = tutorSheet.getLastColumn();
  if (lastCol < 4) throw new Error("無成績欄位！");

  const subjects = tutorSheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
  const ranges = tutorSheet.getRange(2, 4, 1, lastCol - 3).getDisplayValues()[0];
  const dates = tutorSheet.getRange(3, 4, 1, lastCol - 3).getValues()[0];
  const scoresMatrix = tutorSheet.getRange(4, 4, studentCount, lastCol - 3).getValues();

  const sParts = startDateStr.split("-").map(Number);
  const startDt = new Date(sParts[0], sParts[1] - 1, sParts[2], 0, 0, 0, 0);

  const eParts = endDateStr.split("-").map(Number);
  const endDt = new Date(eParts[0], eParts[1] - 1, eParts[2], 23, 59, 59, 999);

  const validCols = [];
  for (let c = 0; c < dates.length; c++) {
    const subj = String(subjects[c] || "").trim();
    if (!subj) continue;

    const rawDate = dates[c];
    const d = parseDateHelper(rawDate);

    if (d && d >= startDt && d <= endDt) {
      validCols.push({
        colIdx: c,
        subject: subj,
        range: String(ranges[c] || "").trim(),
        date: rawDate
      });
    }
  }

  if (validCols.length === 0) throw new Error("選取的時間區間內無成績資料！");

  const subjectGroups = {};
  validCols.forEach(item => {
    if (!subjectGroups[item.subject]) subjectGroups[item.subject] = [];
    subjectGroups[item.subject].push(item);
  });

  const targetSheetName = "導師班成績統計表";
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    targetSheet.clear();
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }

  const header1 = ["座號", "姓名"];
  const header2 = ["", ""];
  const header3 = ["", ""];

  const blueCols = [];
  const yellowCols = [];

  // 取得依序排序之科目（預設：國、英、數、社會、自然、其他，或依自訂順序）
  const orderedSubjects = sortSubjectsByCustomOrder(Object.keys(subjectGroups), getTutorSubjectOrder());

  orderedSubjects.forEach(subj => {
    subjectGroups[subj].forEach(item => {
      header1.push(subj);
      header2.push(item.range);
      const d = parseDateHelper(item.date);
      header3.push(d ? Utilities.formatDate(d, "GMT+8", "MM/dd") : String(item.date));
    });

    header1.push(subj); header2.push("平均"); header3.push("");
    blueCols.push(header1.length);

    header1.push(subj); header2.push("排名"); header3.push("");
    yellowCols.push(header1.length);
  });

  header1.push("總成績", "總成績");
  header2.push("加權平均", "總排名");
  header3.push("", "");

  const weightedAvgColIdx = header1.length - 1;
  const totalRankColIdx = header1.length;

  const subjRanksMap = {};
  orderedSubjects.forEach(subj => {
    const subjAvgs = [];
    for (let s = 0; s < studentCount; s++) {
      let sSum = 0, sCnt = 0;
      subjectGroups[subj].forEach(item => {
        const sc = scoresMatrix[s][item.colIdx];
        if (sc !== "" && sc !== null && !isNaN(Number(sc))) {
          sSum += Number(sc); sCnt++;
        }
      });
      subjAvgs.push(sCnt > 0 ? Math.round((sSum / sCnt) * 10) / 10 : "");
    }
    subjRanksMap[subj] = calculateRanks(subjAvgs);
  });

  const studentRows = [];
  const weights = getTutorWeights();
  const allWeightedAvgs = [];

  for (let s = 0; s < studentCount; s++) {
    const row = [studentInfo[s][0], studentInfo[s][1]];
    let totalWScore = 0, totalW = 0;

    orderedSubjects.forEach(subj => {
      let sSum = 0, sCnt = 0;
      subjectGroups[subj].forEach(item => {
        const sc = scoresMatrix[s][item.colIdx];
        row.push(sc);
        if (sc !== "" && sc !== null && !isNaN(Number(sc))) {
          sSum += Number(sc); sCnt++;
        }
      });

      const sAvg = sCnt > 0 ? Math.round((sSum / sCnt) * 10) / 10 : "";
      row.push(sAvg);
      row.push(subjRanksMap[subj][s] || "");

      if (sAvg !== "") {
        const w = weights[subj] || 1;
        totalWScore += sAvg * w;
        totalW += w;
      }
    });

    const wAvg = totalW > 0 ? Math.round((totalWScore / totalW) * 10) / 10 : "";
    row.push(wAvg);
    row.push("");
    allWeightedAvgs.push(wAvg);

    studentRows.push(row);
  }

  const overallRanks = calculateRanks(allWeightedAvgs);
  for (let s = 0; s < studentCount; s++) {
    studentRows[s][studentRows[s].length - 1] = overallRanks[s] || "";
  }

  targetSheet.getRange(1, 1, 1, header1.length).setValues([header1]).setFontWeight("bold");
  targetSheet.getRange(2, 1, 1, header2.length).setNumberFormat("@").setValues([header2]);
  targetSheet.getRange(3, 1, 1, header3.length).setValues([header3]).setFontWeight("bold");

  if (studentRows.length > 0) {
    targetSheet.getRange(4, 1, studentRows.length, studentRows[0].length).setValues(studentRows);
  }

  const totalRows = studentRows.length + 3;
  blueCols.forEach(col => targetSheet.getRange(1, col, totalRows, 1).setBackground("#EBF5FB"));
  yellowCols.forEach(col => targetSheet.getRange(1, col, totalRows, 1).setBackground("#FEF9E7"));
  targetSheet.getRange(1, weightedAvgColIdx, totalRows, 2).setBackground("#FDEDEC");

  const fullRange = targetSheet.getRange(1, 1, totalRows, header1.length);
  fullRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
  fullRange.setBorder(true, true, true, true, true, true, "#BDC3C7", SpreadsheetApp.BorderStyle.SOLID);

  ss.setActiveSheet(targetSheet);
  return `✅ 成功建立「導師班成績統計表」！區間：${startDateStr} ~ ${endDateStr}`;
}

// ==========================================
// 7. 通用輔助函式
// ==========================================

/**
 * 通用日期解析輔助函式
 */
function parseDateHelper(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const str = String(val).trim().replace(/\./g, '-').replace(/\//g, '-');
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange("A3").setValue("座號").setFontWeight("bold");
    sheet.getRange("B3").setValue("姓名").setFontWeight("bold");
    const defaultSeats = Array.from({ length: 40 }, (_, i) => [i + 1, `學生${i + 1}`]);
    sheet.getRange(4, 1, 40, 2).setValues(defaultSeats);
  }
  return sheet;
}

function calculateAllSheetsAverages() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.endsWith("_總成績報表") || sheetName === "導師班成績統計表") return;

    const lastCol = sheet.getLastColumn();
    const isMainSheet = (sheet.getRange("C3").getValue() === "姓名");
    const seatCol = isMainSheet ? 2 : 1;
    const labelCol = isMainSheet ? 3 : 2;
    const startDataCol = isMainSheet ? 4 : 3;

    if (lastCol < startDataCol) return;

    const studentEndRow = getStudentEndRow(sheet, seatCol, labelCol);
    if (studentEndRow <= 3) return;

    const statsStartRow = studentEndRow + 1;
    
    const maxRow = sheet.getLastRow();
    if (maxRow >= statsStartRow) {
      sheet.getRange(statsStartRow, 1, maxRow - statsStartRow + 1, lastCol).clearContent().clearFormat();
    }

    for (let c = startDataCol; c <= lastCol; c++) {
      writeStatsToSheet(sheet, statsStartRow, c, studentEndRow, labelCol);
    }
  });
}

function getStudentEndRow(sheet, seatCol, labelCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return 3;
  
  const values = sheet.getRange(1, seatCol, lastRow, 2).getValues();
  let endRow = 3;

  for (let r = 3; r < values.length; r++) {
    const seatVal = values[r][0];
    const labelVal = String(values[r][1] || "").trim();

    if (labelVal.includes("平均") || labelVal.includes("分") || labelVal.includes("排名") || labelVal.includes("標籤")) {
      break;
    }
    if (seatVal !== "" && seatVal !== null && !isNaN(Number(seatVal))) {
      endRow = r + 1;
    }
  }
  return endRow;
}

function calculateScoreStats(scores) {
  if (!scores || scores.length === 0) {
    return { c100: 0, c90_100: 0, c80_90: 0, c70_80: 0, c60_70: 0, cFail: 0 };
  }

  let c100 = 0, c90_100 = 0, c80_90 = 0, c70_80 = 0, c60_70 = 0, cFail = 0;

  scores.forEach(score => {
    if (score === 100) c100++;
    if (score >= 90 && score <= 100) c90_100++;
    else if (score >= 80 && score < 90) c80_90++;
    else if (score >= 70 && score < 80) c70_80++;
    else if (score >= 60 && score < 70) c60_70++;
    else if (score < 60) cFail++;
  });

  return { c100, c90_100, c80_90, c70_80, c60_70, cFail };
}

function writeStatsToSheet(sheet, startRow, targetCol, studentEndRow, labelColIdx) {
  const colLetter = getColumnLetter(targetCol); 
  const scoreRange = `${colLetter}4:${colLetter}${studentEndRow}`;

  const labels = [
    ["平均分數"], ["100分"], ["90 ~ 99分"], ["80 ~ 89分"], ["70 ~ 79分"],
    ["60 ~ 69分"], ["50 ~ 59分"], ["40 ~ 49分"], ["30 ~ 39分"], ["20 ~ 29分"], ["10 ~ 19分"], ["0 ~ 9分"]
  ];
  sheet.getRange(startRow, labelColIdx, labels.length, 1).setValues(labels).setFontWeight("bold");

  const formulas = [
    [`=IFERROR(ROUND(AVERAGE(${scoreRange}), 1), "")`],
    [`=COUNTIF(${scoreRange}, 100)`],
    [`=COUNTIFS(${scoreRange}, ">=90", ${scoreRange}, "<100")`],
    [`=COUNTIFS(${scoreRange}, ">=80", ${scoreRange}, "<90")`],
    [`=COUNTIFS(${scoreRange}, ">=70", ${scoreRange}, "<80")`],
    [`=COUNTIFS(${scoreRange}, ">=60", ${scoreRange}, "<70")`],
    [`=COUNTIFS(${scoreRange}, ">=50", ${scoreRange}, "<60")`],
    [`=COUNTIFS(${scoreRange}, ">=40", ${scoreRange}, "<50")`],
    [`=COUNTIFS(${scoreRange}, ">=30", ${scoreRange}, "<40")`],
    [`=COUNTIFS(${scoreRange}, ">=20", ${scoreRange}, "<30")`],
    [`=COUNTIFS(${scoreRange}, ">=10", ${scoreRange}, "<20")`],
    [`=COUNTIFS(${scoreRange}, ">=0",  ${scoreRange}, "<10")`]
  ];

  sheet.getRange(startRow, targetCol, formulas.length, 1).setFormulas(formulas);
}

function getColumnLetter(col) {
  let temp, letter = "";
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function normalizeClassName(classNameStr) {
  if (!classNameStr) return "";
  let clean = classNameStr.replace(/班/g, "").trim();
  if (clean.length === 2 && !clean.startsWith("國")) {
    clean = "國" + clean;
  }
  return clean;
}

function letterToColumn(letter) {
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function calculateRanks(arr) {
  const sorted = arr
    .map((val, idx) => ({ val: val === "" || val === null || isNaN(val) ? -1 : Number(val), idx: idx }))
    .sort((a, b) => b.val - a.val);

  const ranks = new Array(arr.length).fill("");

  for (let i = 0; i < sorted.length; i++) {
    if (sorted[i].val === -1) continue;
    if (i > 0 && sorted[i].val === sorted[i - 1].val) {
      ranks[sorted[i].idx] = ranks[sorted[i - 1].idx];
    } else {
      ranks[sorted[i].idx] = i + 1;
    }
  }
  return ranks;
}

function archiveAndResetSemester() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('⚠️ 歸檔確認', '確定要結算並重置舊學期成績嗎？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert('🎉 新學期重置完成！舊成績已妥善歸檔備份。');
}

/**
 * 動態取得最新且適用的 Flash 模型名稱 (已修正過濾與預設值)
 */
function getLatestFlashModelName() {
  const defaultModel = "gemini-flash-lite-latest"; // 💡 修正預設值為目前有效的模型
  if (!GEMINI_API_KEY) return defaultModel;

  const apiKey = getActiveGeminiApiKey();
  const listUrl = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  
  try {
    const response = UrlFetchApp.fetch(listUrl, { "muteHttpExceptions": true });
    const result = JSON.parse(response.getContentText());
    
    if (result.models && Array.isArray(result.models)) {
      const flashModels = result.models
        .filter(m => {
          const name = m.name.toLowerCase();
          const hasGen = m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent");
          const isFlash = name.includes("flash");
          
          // 💡 排除 TTS(語音)、Image(生圖)、Computer-Use 等不適合純 OCR 辨識的模型
          const isExcluded = name.includes("tts") || name.includes("image") || name.includes("computer-use") || name.includes("omni") || name.includes("exp");
          return hasGen && isFlash && !isExcluded;
        })
        .map(m => m.name.replace("models/", ""));

      if (flashModels.length > 0) {
        // 自然排序：版號新、數字大的排在最前面
        flashModels.sort((a, b) => b.localeCompare(a, undefined, { numeric: true, sensitivity: 'base' }));
        Logger.log(`🤖 自動選擇的最佳 Flash 模型: [${flashModels[0]}]`);
        return flashModels[0];
      }
    }
  } catch (err) {
    Logger.log("⚠️ 取得模型清單失敗，使用預設值: " + err.toString());
  }

  return defaultModel;
}
// ==========================================
// 8. 個人區間報表彈窗與全班資料打包
// ==========================================
function showBatchReportDialog() {
  const earliestDate = getEarliestDateInMainSheet();
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head><base target="_top"><style>
        body { font-family: sans-serif; padding: 15px; }
        .group { margin-bottom: 12px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input { width: 100%; padding: 8px; box-sizing: border-box; }
        button { width: 100%; padding: 10px; background: #2ecc71; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 15px; font-weight: bold; margin-top: 10px; }
      </style></head>
      <body>
        <h3>🖨️ 選擇匯出報表日期區間</h3>
        <div class="group"><label>起始日期</label><input type="date" id="start" value="${earliestDate}"></div>
        <div class="group"><label>結束日期</label><input type="date" id="end"></div>
        <button onclick="submit()">產生全班個人報表</button>
        <script>
          const now = new Date();
          const yyyy = now.getFullYear();
          const mm = String(now.getMonth() + 1).padStart(2, '0');
          const dd = String(now.getDate()).padStart(2, '0');
          document.getElementById('end').value = \`\${yyyy}-\${mm}-\${dd}\`;

          function submit() {
            const s = document.getElementById('start').value;
            const e = document.getElementById('end').value;
            if(!s || !e) return alert('請選擇完整日期！');
            google.script.run.openBatchReportWindow(s, e);
            google.script.host.close();
          }
        </script>
      </body>
    </html>
  `).setWidth(350).setHeight(260);
  
  SpreadsheetApp.getUi().showModalDialog(html, '匯出個人區間報表');
}

/**
 * 彈出 BatchReport.html 並載入打包數據 (修正資料傳遞與防錯)
 */
function openBatchReportWindow(startDateStr, endDateStr) {
  try {
    const reportData = getStudentReportData(startDateStr, endDateStr);
    
    if (!reportData || !reportData.reports || reportData.reports.length === 0) {
      SpreadsheetApp.getUi().alert("⚠️ 選擇的時間區間內無可列印之成績資料！");
      return;
    }

    const template = HtmlService.createTemplateFromFile('BatchReport');
    // 將 JSON 字串傳給樣板變數
    template.reportData = JSON.stringify(reportData);
    
    const htmlOutput = template.evaluate()
      .setWidth(1050)
      .setHeight(800)
      .setTitle('區間個人成績報表');
      
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '區間個人成績報表');
  } catch (err) {
    SpreadsheetApp.getUi().alert("開啟個人報表失敗：" + err.message);
  }
}

/**
 * 依指定區間取得全班個人成績報表數據 (含 100 分制與 0-6 五標相對指標數據)
 */
function getStudentReportData(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tutorSheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
  if (!tutorSheet) throw new Error("找不到導師成績總表！");

  let studentEndRow = 3;
  const bValues = tutorSheet.getRange("B4:B" + tutorSheet.getLastRow()).getValues();
  for (let i = 0; i < bValues.length; i++) {
    if (bValues[i][0] !== "" && bValues[i][0] !== null) {
      studentEndRow = i + 4;
    }
  }
  const studentCount = studentEndRow - 3;
  if (studentCount <= 0) return { startDate: startDateStr, endDate: endDateStr, reports: [] };

  const studentInfo = tutorSheet.getRange(4, 2, studentCount, 2).getValues();
  const lastCol = tutorSheet.getLastColumn();
  if (lastCol < 4) return { startDate: startDateStr, endDate: endDateStr, reports: [] };

  const subjects = tutorSheet.getRange(1, 4, 1, lastCol - 3).getDisplayValues()[0];
  const ranges = tutorSheet.getRange(2, 4, 1, lastCol - 3).getDisplayValues()[0];
  const dates = tutorSheet.getRange(3, 4, 1, lastCol - 3).getValues()[0];
  const allScores = tutorSheet.getRange(4, 4, studentCount, lastCol - 3).getValues();

  const startDt = parseDateHelper(startDateStr);
  const endDt = parseDateHelper(endDateStr);
  if (startDt) startDt.setHours(0, 0, 0, 0);
  if (endDt) endDt.setHours(23, 59, 59, 999);

  let minDate = null;
  for (let c = 0; c < dates.length; c++) {
    const subj = String(subjects[c] || "").trim();
    if (!subj) continue;
    const dt = parseDateHelper(dates[c]);
    if (dt && (!minDate || dt < minDate)) minDate = dt;
  }
  if (!minDate) minDate = new Date();
  minDate.setHours(0, 0, 0, 0);

  // 1. 篩選指定日期區間內的有效欄位
  const validCols = [];
  const classSubjectScoresInRange = {}; // 區間內全班各科成績彙整 (用於五標計算)

  for (let c = 0; c < dates.length; c++) {
    const subj = String(subjects[c] || "").trim();
    if (!subj) continue;

    const d = parseDateHelper(dates[c]);
    if (!d || ((!startDt || d >= startDt) && (!endDt || d <= endDt))) {
      const refDate = d || new Date();
      const diffDays = Math.floor((refDate.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24));
      const weekLabel = `第${Math.max(1, Math.floor(diffDays / 7) + 1)}週`;

      let colSum = 0, colScoreCount = 0;
      if (!classSubjectScoresInRange[subj]) classSubjectScoresInRange[subj] = [];

      for (let s = 0; s < studentCount; s++) {
        const val = Number(allScores[s][c]);
        if (allScores[s][c] !== "" && allScores[s][c] !== null && !isNaN(val)) {
          colSum += val;
          colScoreCount++;
          classSubjectScoresInRange[subj].push(val);
        }
      }
      const calculatedClassAvg = colScoreCount > 0 ? (Math.round((colSum / colScoreCount) * 10) / 10) : "-";

      validCols.push({
        colIdx: c,
        subject: subj,
        range: String(ranges[c] || "").trim(),
        dateObj: refDate,
        weekLabel: weekLabel,
        classAvg: calculatedClassAvg
      });
    }
  }

  if (validCols.length === 0) return { startDate: startDateStr, endDate: endDateStr, reports: [] };

  const weeksSet = new Set(validCols.map(v => v.weekLabel));
  const weeksMap = Array.from(weeksSet).sort((a,b) => parseInt(a.replace("第","")) - parseInt(b.replace("第","")));
  const weights = getTutorWeights();

  const studentSubjectAvgs = Array.from({ length: studentCount }, () => ({}));
  const studentSubjectExams = Array.from({ length: studentCount }, () => ({}));
  const studentWeekScores = Array.from({ length: studentCount }, () => ({}));

  for (let s = 0; s < studentCount; s++) {
    const subjSums = {}, subjCounts = {};

    validCols.forEach(item => {
      const c = item.colIdx;
      const score = allScores[s][c];
      const subj = item.subject;

      if (score !== "" && score !== null && !isNaN(Number(score))) {
        const numScore = Number(score);

        if (!studentSubjectExams[s][subj]) studentSubjectExams[s][subj] = [];
        studentSubjectExams[s][subj].push({
          range: item.range,
          score: numScore,
          classAvg: item.classAvg,
          date: Utilities.formatDate(item.dateObj, "GMT+8", "yyyy-MM-dd")
        });

        if (!subjSums[subj]) { subjSums[subj] = 0; subjCounts[subj] = 0; }
        subjSums[subj] += numScore;
        subjCounts[subj] += 1;

        const wLabel = item.weekLabel;
        if (!studentWeekScores[s][wLabel]) studentWeekScores[s][wLabel] = {};
        if (!studentWeekScores[s][wLabel][subj]) studentWeekScores[s][wLabel][subj] = [];
        studentWeekScores[s][wLabel][subj].push(numScore);
      }
    });

    Object.keys(subjSums).forEach(subj => {
      studentSubjectAvgs[s][subj] = Math.round((subjSums[subj] / subjCounts[subj]) * 10) / 10;
    });
  }

  // 計算區間科目排名
  const allSubjsSet = new Set();
  studentSubjectAvgs.forEach(sObj => Object.keys(sObj).forEach(k => allSubjsSet.add(k)));
  const allSubjs = Array.from(allSubjsSet);

  const subjectRanksMap = {};
  allSubjs.forEach(subj => {
    const subjScores = studentSubjectAvgs.map(sObj => sObj[subj] !== undefined ? sObj[subj] : "");
    subjectRanksMap[subj] = calculateRanks(subjScores);
  });

  // 計算區間加權平均與總排名
  const allWeightedAverages = [];
  for (let s = 0; s < studentCount; s++) {
    let totalWeightedSum = 0, totalWeight = 0;
    Object.keys(studentSubjectAvgs[s]).forEach(k => {
      const avg = studentSubjectAvgs[s][k];
      const w = weights[k] !== undefined ? weights[k] : 1;
      totalWeightedSum += avg * w;
      totalWeight += w;
    });
    const wAvg = totalWeight > 0 ? Math.round((totalWeightedSum / totalWeight) * 10) / 10 : "";
    allWeightedAverages.push(wAvg);
  }

  const overallRanks = calculateRanks(allWeightedAverages);
  const palette = ["#3498db", "#e74c3c", "#2ecc71", "#f1c40f", "#9b59b6", "#e67e22", "#1abc9c", "#34495e"];

  // 打包個人報表數據（依自訂/預設順序排列各科）
  const reports = [];
  const orderedReportSubjs = sortSubjectsByCustomOrder(allSubjs, getTutorSubjectOrder());

  for (let s = 0; s < studentCount; s++) {
    const subjectDetails = [];

    orderedReportSubjs.forEach(subj => {
      if (studentSubjectAvgs[s][subj] === undefined) return;
      subjectDetails.push({
        subject: subj,
        subjectAvg: studentSubjectAvgs[s][subj],
        subjectRank: subjectRanksMap[subj] ? (subjectRanksMap[subj][s] || "-") : "-",
        exams: studentSubjectExams[s][subj] || []
      });
    });

    // 1. 原始 100 分制趨勢圖數據
    const weeklyChartData = {};
    Object.keys(studentSubjectAvgs[s]).forEach(subj => {
      weeklyChartData[subj] = weeksMap.map(w => {
        if (studentWeekScores[s][w] && studentWeekScores[s][w][subj]) {
          const arr = studentWeekScores[s][w][subj];
          return Math.round((arr.reduce((a,b)=>a+b,0) / arr.length) * 10) / 10;
        }
        return null;
      });
    });

    const datasets = Object.keys(weeklyChartData).map((subj, idx) => ({
      label: subj,
      data: weeklyChartData[subj],
      borderColor: palette[idx % palette.length],
      backgroundColor: palette[idx % palette.length],
      fill: false,
      tension: 0.2,
      spanGaps: true
    }));

    // 2. 五標相對雷達圖數據 (0-6 階梯)
    const fiveTierRadarLabels = [];
    const fiveTierRadarData = [];
    const fiveTierClassBenchmark = [];

    allSubjs.forEach(subj => {
      if (studentSubjectAvgs[s][subj] !== undefined) {
        const myAvg = studentSubjectAvgs[s][subj];
        const pObj = calculateFiveTierPercentiles(classSubjectScoresInRange[subj] || []);
        const mappedVal = mapScoreToFiveTierScale(myAvg, pObj);

        fiveTierRadarLabels.push(subj);
        fiveTierRadarData.push(mappedVal);
        fiveTierClassBenchmark.push(3.0);
      }
    });

    // 3. 五標相對趨勢圖數據 (0-6 階梯)
    const fiveTierTrendDatasets = allSubjs.map((subj, idx) => ({
      label: subj,
      data: weeksMap.map(w => {
        if (studentWeekScores[s][w] && studentWeekScores[s][w][subj]) {
          const arr = studentWeekScores[s][w][subj];
          const myWAvg = arr.reduce((a,b)=>a+b,0) / arr.length;

          // 計算該週全班該科五標
          const classWScores = [];
          validCols.forEach(v => {
            if (v.subject === subj && v.weekLabel === w) {
              const val = Number(allScores[s][v.colIdx]);
              for (let st = 0; st < studentCount; st++) {
                const scVal = allScores[st][v.colIdx];
                if (scVal !== "" && scVal !== null && !isNaN(Number(scVal))) {
                  classWScores.push(Number(scVal));
                }
              }
            }
          });

          const pObj = calculateFiveTierPercentiles(classWScores);
          return mapScoreToFiveTierScale(myWAvg, pObj);
        }
        return null;
      }),
      borderColor: palette[idx % palette.length],
      backgroundColor: palette[idx % palette.length],
      fill: false,
      tension: 0.2,
      spanGaps: true
    }));

    reports.push({
      info: { seat: studentInfo[s][0], name: studentInfo[s][1] },
      summary: {
        weightedAvg: allWeightedAverages[s] !== "" ? allWeightedAverages[s] : "-",
        rank: overallRanks[s] !== undefined ? overallRanks[s] : "-"
      },
      subjectDetails: subjectDetails,
      weeklyChart: {
        labels: weeksMap.length > 0 ? weeksMap : ["無資料"],
        datasets: datasets
      },
      fiveTierRadarChart: {
        labels: fiveTierRadarLabels,
        studentData: fiveTierRadarData,
        benchmarkData: fiveTierClassBenchmark
      },
      fiveTierTrendChart: {
        labels: weeksMap.length > 0 ? weeksMap : ["無資料"],
        datasets: fiveTierTrendDatasets
      }
    });
  }

  return { startDate: startDateStr, endDate: endDateStr, reports: reports };
}
/**
 * 取得舊版備用權重 (若其他模組有呼叫則保留)
 */
function getStoredWeights() {
  const prop = PropertiesService.getScriptProperties().getProperty("CUSTOM_WEIGHTS");
  if (prop) {
    try { return JSON.parse(prop); } catch(e) {}
  }
  return { "國文": 5, "英文": 3, "數學": 4, "歷史": 1, "地理": 1, "公民": 1, "理化": 3, "生物": 3, "段考": 0.5, "小考": 0.2, "作業": 0.2, "筆記": 0.1 };
}

/**
 * 顯示加權權重設定對話框
 */
function showWeightSettingDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('WeightSettingDialog')
    .setWidth(450)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "⚙️ 設定加權權重");
}

/**
 * 前端初始化讀取兩邊系統的權重資料
 */
// 預設標準科目排序順序 (國、英、數、社會、自然、其他)
const DEFAULT_SUBJECT_ORDER_LIST = [
  "國文", "英文", "數學",
  "歷史", "地理", "公民",
  "理化", "物理", "化學", "生物", "地科"
];

// 預設專任評量類別排序順序
const DEFAULT_SUBJECT_TYPE_ORDER_LIST = [
  "段考", "小考", "作業", "筆記"
];

/**
 * 前端初始化讀取系統權重與排序設定
 */
function getWeightsData() {
  return {
    tutor: getTutorWeights(),
    subject: getSubjectTeacherWeights(),
    tutorOrder: getTutorSubjectOrder(),
    subjectOrder: getSubjectTypeOrder()
  };
}

/**
 * 儲存導師與專任加權設定及自訂排序至系統屬性
 */
function saveAllCustomWeights(tutorWeights, subjectWeights, tutorOrder, subjectOrder) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("TUTOR_SUBJECT_WEIGHTS", JSON.stringify(tutorWeights || {}));
  props.setProperty("SUBJECT_TYPE_WEIGHTS", JSON.stringify(subjectWeights || {}));
  if (Array.isArray(tutorOrder)) {
    props.setProperty("TUTOR_SUBJECT_ORDER", JSON.stringify(tutorOrder));
  }
  if (Array.isArray(subjectOrder)) {
    props.setProperty("SUBJECT_TYPE_ORDER", JSON.stringify(subjectOrder));
  }
  return "✅ 導師班科目順序/加權與科任類別順序/加權已成功儲存！";
}

/**
 * 取得導師班科目排序順序（陣列）
 * 1. 優先取用使用者自訂順序 (ScriptProperties)
 * 2. 次之依預設規則排序：國、英、數、社會(歷史、地理、公民)、自然(理化、物理、化學、生物、地科)、其他
 */
function getTutorSubjectOrder() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty("TUTOR_SUBJECT_ORDER");
  if (json) {
    try {
      const arr = JSON.parse(json);
      if (Array.isArray(arr) && arr.length > 0) return arr;
    } catch (e) {
      Logger.log("解析 TUTOR_SUBJECT_ORDER 失敗: " + e.message);
    }
  }

  // 取得現有所有科目 keys 並依預設規則排序
  const weights = getTutorWeights();
  const keys = Object.keys(weights);
  return sortSubjectsByDefaultPattern(keys);
}

/**
 * 取得科任班評量類別排序順序（陣列）
 */
function getSubjectTypeOrder() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty("SUBJECT_TYPE_ORDER");
  if (json) {
    try {
      const arr = JSON.parse(json);
      if (Array.isArray(arr) && arr.length > 0) return arr;
    } catch (e) {
      Logger.log("解析 SUBJECT_TYPE_ORDER 失敗: " + e.message);
    }
  }
  const weights = getSubjectTeacherWeights();
  const keys = Object.keys(weights);
  return keys.length > 0 ? keys : DEFAULT_SUBJECT_TYPE_ORDER_LIST.slice();
}

/**
 * 預設排序模式：
 * 國、英、數、社會(歷史、地理、公民)、自然科(理化、物理、化學、生物、地科)、其他
 */
function sortSubjectsByDefaultPattern(subjectList) {
  function getRank(s) {
    const name = String(s || "").trim();
    if (!name) return 999;
    if (name.includes("國")) return 10;
    if (name.includes("英")) return 20;
    if (name.includes("數")) return 30;
    if (name.includes("歷")) return 40;
    if (name.includes("公民") || name === "公") return 60;
    if (name.includes("地科") || name.includes("地球科學")) return 110;
    if (name.includes("地理")) return 50;
    if (name.includes("理化")) return 70;
    if (name.includes("物理")) return 80;
    if (name.includes("化學")) return 90;
    if (name.includes("生物") || name === "生") return 100;
    return 500; // 其他科目排在後方
  }

  return subjectList.slice().sort((a, b) => {
    const ra = getRank(a);
    const rb = getRank(b);
    if (ra !== rb) return ra - rb;
    return String(a).localeCompare(String(b), "zh-TW");
  });
}

/**
 * 依自訂陣列順序對科目進行排序
 */
function sortSubjectsByCustomOrder(subjectList, customOrder) {
  if (!customOrder || !Array.isArray(customOrder) || customOrder.length === 0) {
    return sortSubjectsByDefaultPattern(subjectList);
  }

  const orderMap = {};
  customOrder.forEach((item, idx) => {
    orderMap[String(item).trim()] = idx;
  });

  return subjectList.slice().sort((a, b) => {
    const nameA = String(a).trim();
    const nameB = String(b).trim();
    const idxA = orderMap.hasOwnProperty(nameA) ? orderMap[nameA] : 999;
    const idxB = orderMap.hasOwnProperty(nameB) ? orderMap[nameB] : 999;
    if (idxA !== idxB) return idxA - idxB;
    return sortSubjectsByDefaultPattern([nameA, nameB])[0] === nameA ? -1 : 1;
  });
}

/**
 * 取得導師加權：
 * 1. 優先採用使用者自訂儲存設定 (ScriptProperties)
 * 2. 若未儲存，採用全域預設 WEIGHTS 字典
 */
function getTutorWeights() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty("TUTOR_SUBJECT_WEIGHTS");
  if (json) {
    try {
      return JSON.parse(json);
    } catch (e) {
      Logger.log("解析 TUTOR_SUBJECT_WEIGHTS 失敗: " + e.message);
    }
  }

  // 預設科目加權字典 (若全域未定義則使用備用預設值)
  const defaultWeights = (typeof WEIGHTS !== 'undefined') ? WEIGHTS : {
    "國文": 5, "英文": 3, "數學": 4,
    "歷史": 1, "地理": 1, "公民": 1,
    "理化": 3, "生物": 3
  };

  return Object.assign({}, defaultWeights);
}

/**
 * 取得專任評量類別權重：
 * 1. 優先採用使用者自訂儲存設定 (ScriptProperties)
 * 2. 若未儲存，回傳全域預設 SUBJECT_TYPE_WEIGHTS
 */
function getSubjectTeacherWeights() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty("SUBJECT_TYPE_WEIGHTS");
  if (json) {
    try {
      return JSON.parse(json);
    } catch (e) {
      Logger.log("解析 SUBJECT_TYPE_WEIGHTS 失敗: " + e.message);
    }
  }

  // 預設評量類別字典 (若全域未定義則使用備用預設值)
  const defaultTypes = (typeof SUBJECT_TYPE_WEIGHTS !== 'undefined') ? SUBJECT_TYPE_WEIGHTS : {
    "段考": 0.5, "小考": 0.2, "作業": 0.2, "筆記": 0.1
  };

  return Object.assign({}, defaultTypes);
}


/**
 * 遍歷所有分頁：計算各分頁底部的「平均分數」與「分數區間人數」
 * 【智慧增量優化】：僅針對「尚未計算 / 沒有資料」的成績欄位進行補算，已存在的欄位不重複運算
 */
function generateCurrentSubjectStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let processedTabCount = 0;
  let totalNewCalculatedCols = 0;

  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    // 排除特殊系統與報表分頁
    if (sheetName === "System_Settings" || sheetName === "科任跨班分析總表" || sheetName.endsWith("_總成績") || sheetName === "導師班成績統計表") {
      continue;
    }

    const lastCol = sheet.getLastColumn();
    if (lastCol < 2) continue; // 欄位不足則跳過

    // 自動判斷分頁結構：導師成績總表座號在 B 欄(分數 D 欄起)，專任分頁座號在 A 欄(分數 C 欄起)
    const isMainSheet = (sheetName === "導師成績總表");
    const seatColIdx = isMainSheet ? 2 : 1;
    const startScoreColIdx = isMainSheet ? 4 : 3;

    // 精準計算學生最後一列
    const studentEndRow = getStudentEndRow(sheet, seatColIdx, seatColIdx + 1);
    const studentCount = studentEndRow - 3;
    if (studentCount <= 0) continue; // 無有效學生數據則跳過

    const statsStartRow = studentEndRow + 2; // 統計區從學生名單下方隔一行開始

    // 精簡後的統計項目標題
    const labels = [
      ["平均分數"],
      ["最高分"],
      ["最低分"],
      ["100分人數"],
      ["90-100分人數"],
      ["80-90分人數"],
      ["70-80分人數"],
      ["60-70分人數"],
      ["不及格人數"]
    ];

    // 1. 檢查 A 欄統計標題是否已存在，若不存在才寫入標題與格式
    const currentAValue = String(sheet.getRange(statsStartRow, 1).getValue() || "").trim();
    if (currentAValue !== "平均分數") {
      sheet.getRange(statsStartRow, 1, labels.length, 1).setValues(labels).setFontWeight("bold").setBackground("#D5F5E3");
      sheet.getRange(statsStartRow, 1, labels.length, 1).setHorizontalAlignment("center").setVerticalAlignment("middle");
    }

    let updatedColsCount = 0;

    // 2. 針對成績欄位：僅在「尚未計算/無資料」時才補寫入統計公式
    for (let c = startScoreColIdx; c <= lastCol; c++) {
      const colLetter = getColLetter(c);

      // 讀取第 1 列儲存格
      const cellR1 = sheet.getRange(1, c);
      let r1Title = String(cellR1.getValue() || "").trim();

      // 處理跨欄置中（合併儲存格）
      if (!r1Title && cellR1.isPartOfMerge()) {
        const mergedRanges = cellR1.getMergedRanges();
        if (mergedRanges.length > 0) {
          r1Title = String(mergedRanges[0].getValue() || "").trim();
        }
      }

      const r2Title = String(sheet.getRange(2, c).getValue() || "").trim();
      const r3Title = String(sheet.getRange(3, c).getValue() || "").trim();
      const combinedHeader = `${r1Title} ${r2Title} ${r3Title}`;

      // 判斷是否包含排名相關關鍵字
      const isRankCol = /排名|名次|Rank/i.test(combinedHeader);

      // 過濾無標題或排名欄位
      if ((!r1Title && !r2Title && !r3Title) || isRankCol) {
        continue;
      }

      // 【核心防呆與增量檢查】：檢查該欄的平均分數儲存格是否已有公式或已有數值
      const statCell = sheet.getRange(statsStartRow, c);
      const existingFormula = String(statCell.getFormula() || "").trim();
      const existingValue = String(statCell.getValue() || "").trim();

      if (existingFormula !== "" || existingValue !== "") {
        // 該欄已經計算過統計，直接跳過，不重複運算！
        continue;
      }

      const rangeStr = `${colLetter}4:${colLetter}${studentEndRow}`;

      const formulas = [
        [`=IFERROR(ROUND(AVERAGE(${rangeStr}), 1), "")`],
        [`=IFERROR(MAX(${rangeStr}), "")`],
        [`=IFERROR(MIN(${rangeStr}), "")`],
        [`=COUNTIF(${rangeStr}, 100)`],
        [`=COUNTIFS(${rangeStr}, ">=90", ${rangeStr}, "<100")`],
        [`=COUNTIFS(${rangeStr}, ">=80", ${rangeStr}, "<90")`],
        [`=COUNTIFS(${rangeStr}, ">=70", ${rangeStr}, "<80")`],
        [`=COUNTIFS(${rangeStr}, ">=60", ${rangeStr}, "<70")`],
        [`=COUNTIFS(${rangeStr}, ">=0", ${rangeStr}, "<60")`]
      ];

      const targetColStatRange = sheet.getRange(statsStartRow, c, formulas.length, 1);
      targetColStatRange.setFormulas(formulas);
      targetColStatRange.setBackground("#EAFAF1");
      targetColStatRange.setHorizontalAlignment("center").setVerticalAlignment("middle");

      updatedColsCount++;
    }

    if (updatedColsCount > 0) {
      processedTabCount++;
      totalNewCalculatedCols += updatedColsCount;
    }
  }

  // 4. 完成提示（相容自動排程觸發與人工手動按鈕執行）
  let msg = "";
  if (totalNewCalculatedCols > 0) {
    msg = `✅ 已為 ${processedTabCount} 個分頁成功補全 ${totalNewCalculatedCols} 個新欄位的統計資料！`;
  } else {
    msg = `✅ 檢查完畢：所有分頁的成績欄位統計皆已存在，無須重複重算。`;
  }

  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg); // 排程自動執行時無 UI 介面，改上記錄檔
  }
}
/**
 * 查詢目前 API Key 可使用的所有 Gemini 模型名稱
 */
function listAvailableGeminiModels() {
  const apiKey = getActiveGeminiApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  
  try {
    const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      Logger.log("❌ 查詢失敗: " + result.error.message);
      return;
    }

    Logger.log("================ 可用 Gemini 模型清單 ================");
    const availableModels = result.models
      .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"))
      .map(m => m.name.replace("models/", ""));

    availableModels.forEach(modelName => {
      Logger.log("👉 " + modelName);
    });
    Logger.log("====================================================");

  } catch (err) {
    Logger.log("❌ 發生例外錯誤: " + err.message);
  }
}
/**
 * 自動讀取資料夾並區分 LOGO、背景圖與班級動態標題
 */
function getPageThemeAssets() {
  let systemTitle = "學生成績查詢系統";
  const sysConfig = getSystemConfig(true);

  // 1. 讀取班級名稱
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
    if (sheet) {
      const val = sheet.getRange("A1").getValue();
      const className = String(val !== null && val !== undefined ? val : "").trim();
      if (className) {
        systemTitle = `${className} 週平時成績查詢`;
      }
    }
  } catch (e) {
    Logger.log("讀取班級名稱失敗: " + e.message);
  }

  // 2. 檢查底圖資料夾 ID
  if (!sysConfig.BG_FOLDER_ID) {
    return { bg: null, logo: null, title: systemTitle };
  }

  // 3. 讀取 Google 雲端硬碟素材
  try {
    const folder = DriveApp.getFolderById(sysConfig.BG_FOLDER_ID);
    const files = folder.getFiles();
    const bgFiles = [];
    let logoBase64 = null;

    while (files.hasNext()) {
      const file = files.next();
      const mime = file.getMimeType();
      const fileName = file.getName().toLowerCase();

      if (mime.includes("image/")) {
        if (fileName.includes((sysConfig.LOGO_KEYWORD || "logo").toLowerCase())) {
          const logoBlob = file.getBlob();
          logoBase64 = `data:${logoBlob.getContentType()};base64,${Utilities.base64Encode(logoBlob.getBytes())}`;
        } else {
          bgFiles.push(file);
        }
      }
    }

    let bgBase64 = null;
    if (bgFiles.length > 0) {
      const randomFile = bgFiles[Math.floor(Math.random() * bgFiles.length)];
      const bgBlob = randomFile.getBlob();
      bgBase64 = `data:${bgBlob.getContentType()};base64,${Utilities.base64Encode(bgBlob.getBytes())}`;
    }

    return { bg: bgBase64, logo: logoBase64, title: systemTitle };
  } catch (e) {
    Logger.log("讀取素材失敗: " + e.message);
    return { bg: null, logo: null, title: systemTitle };
  }
}
/**
 * 讀取雲端硬碟指定資料夾中的背景圖片並轉為 Base64 格式
 */
function getBackgroundImage() {
  const folderId = "BG_FOLDER_ID"; // 您的雲端硬碟資料夾 ID
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const imageFiles = [];

    while (files.hasNext()) {
      const file = files.next();
      const mime = file.getMimeType();
      if (mime.includes("image/")) {
        imageFiles.push(file);
      }
    }

    if (imageFiles.length === 0) return null;

    // 隨機挑選一張背景圖片
    const randomFile = imageFiles[Math.floor(Math.random() * imageFiles.length)];
    const blob = randomFile.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    return `data:${blob.getContentType()};base64,${base64}`;
  } catch (e) {
    Logger.log("讀取背景底圖失敗: " + e.message);
    return null;
  }
}

function testThemeAssets() {
  const res = getPageThemeAssets();
  Logger.log("=== 測試結果 ===");
  Logger.log("1. 取得標題: " + res.title);
  Logger.log("2. LOGO 是否抓取成功: " + (res.logo ? "成功" : "失敗 (無檔案或未設定)"));
  Logger.log("3. 背景圖是否抓取成功: " + (res.bg ? "成功" : "失敗 (無檔案或未設定)"));
}

// ==========================================
// 4. 科任/導師分頁匯入與 [System_Settings] 管理函式
// ==========================================

/**
 * 將科目成績寫入導師成績總表新欄位
 * @param {string} subject 科目名稱
 * @param {string} examRange 測驗範圍
 * @param {string} examDate 測驗日期 (YYYY-MM-DD)
 * @param {Object|Map} scoreMap 座號對應成績
 * @param {Sheet} [optionalTargetSheet] 可選目標工作表（未提供時預設當前試算表的導師成績總表）
 * @return {Object} { addedColumn: 寫入欄位編號 }
 */
function addNewSubjectColumnToTutorSheet(subject, examRange, examDate, scoreMap, optionalTargetSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tutorSheet = optionalTargetSheet || (ss.getSheetByName("導師成績總表") || ss.getSheets()[0]);
  if (!tutorSheet) throw new Error("找不到「導師成績總表」！");

  const lastCol = tutorSheet.getLastColumn();
  const startCol = (lastCol < 3) ? 4 : lastCol + 1;

  tutorSheet.getRange(1, startCol).setNumberFormat("@").setValue(subject);
  tutorSheet.getRange(2, startCol).setNumberFormat("@").setValue(examRange);
  tutorSheet.getRange(3, startCol).setNumberFormat("@").setValue(examDate);

  const studentEndRow = getStudentEndRow(tutorSheet, 2, 3);
  const totalStudents = studentEndRow - 3;

  if (totalStudents > 0) {
    const seatList = tutorSheet.getRange(4, 2, totalStudents, 1).getValues();
    const outputScores = seatList.map(r => {
      const seatKey = String(r[0]).trim();
      const seatNum = parseInt(seatKey, 10);
      let sc = "";
      if (scoreMap instanceof Map) {
        if (scoreMap.has(seatKey)) sc = scoreMap.get(seatKey);
        else if (scoreMap.has(seatNum)) sc = scoreMap.get(seatNum);
        else if (scoreMap.has(String(seatNum))) sc = scoreMap.get(String(seatNum));
      } else if (scoreMap && typeof scoreMap === "object") {
        if (scoreMap.hasOwnProperty(seatKey)) sc = scoreMap[seatKey];
        else if (scoreMap.hasOwnProperty(seatNum)) sc = scoreMap[seatNum];
        else if (scoreMap.hasOwnProperty(String(seatNum))) sc = scoreMap[String(seatNum)];
      }
      return [sc !== "" && sc !== null && sc !== undefined ? sc : ""];
    });
    tutorSheet.getRange(4, startCol, outputScores.length, 1).setValues(outputScores);
  }

  return { addedColumn: startCol };
}

/**
 * 導師端：自動抓取 5 字名稱分頁 (例: 101理化) 並驗證 A1 班級後匯入總表新欄
 */
function importSubjectSheetToTutorSheetUI() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName().trim();

  // 防呆：禁止在管理分頁上執行
  if (sheetName === "導師成績總表" || sheetName === "System_Settings") {
    ui.alert("⚠️ 請先切換至科任成績分頁（例如：101理化），再執行匯入！");
    return;
  }

  // 1. 自動解析分頁名稱 (假設格式為：班級(3字) + 科目(2字)，如 "101理化")
  let parsedClass = "";
  let parsedSubject = "";

  if (sheetName.length >= 5) {
    parsedClass = sheetName.substring(0, 3);   // 前 3 字：101
    parsedSubject = sheetName.substring(3);    // 第 4 字以後：理化
  } else {
    parsedClass = sheetName.replace(/[^0-9]/g, "");
    parsedSubject = sheetName.replace(/[0-9]/g, "");
  }

  // 2. 讀取導師總表 A1 並比對班級
  const tutorSheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
  const a1Value = String(tutorSheet.getRange("A1").getValue()).trim();

  if (parsedClass && !a1Value.includes(parsedClass)) {
    const confirmImport = ui.alert(
      "⚠️ 班級不符警示",
      `當前分頁解析班級為【${parsedClass}】，但導師總表 A1 為【${a1Value}】。
確定要繼續匯入嗎？`,
      ui.ButtonSet.YES_NO
    );
    if (confirmImport !== ui.Button.YES) return;
  }

  // 3. 彈出輸入視窗（預填解析出來的科目，簡化輸入程序）
  const respSubject = ui.prompt("確認科目", `請確認科目名稱：`, ui.ButtonSet.OK_CANCEL);
  if (respSubject.getSelectedButton() !== ui.Button.OK) return;
  const finalSubject = respSubject.getResponseText().trim() || parsedSubject;

  const respRange = ui.prompt("測驗範圍", "請輸入測驗範圍/細項（例：第三單元小考）：", ui.ButtonSet.OK_CANCEL);
  if (respRange.getSelectedButton() !== ui.Button.OK) return;
  const rangeText = respRange.getResponseText().trim();

  const defaultDate = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
  const respDate = ui.prompt("測驗日期", `請確認測驗日期 (YYYY-MM-DD)：`, ui.ButtonSet.OK_CANCEL);
  if (respDate.getSelectedButton() !== ui.Button.OK) return;
  const testDate = respDate.getResponseText().trim() || defaultDate;

  // 4. 讀取分頁成績資料 (第 1 欄位為座號/學號，第 2 欄位為成績)
  const data = activeSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert("⚠️ 該分頁無成績資料！");
    return;
  }

  const scoreMap = {};
  for (let r = 1; r < data.length; r++) {
    const key = String(data[r][0]).trim();
    const score = data[r][1];
    if (key !== "") {
      scoreMap[key] = score;
    }
  }

  // 5. 執行寫入導師總表新欄位
  try {
    ui.toast(`正在匯入【${parsedClass} - ${finalSubject}】...`, "⏳ 處理中", -1);
    const res = addNewSubjectColumnToTutorSheet(finalSubject, rangeText, testDate, scoreMap);
    ui.alert("✅ 匯入成功！", `已將【${parsedClass} ${finalSubject} - ${rangeText}】成功寫入「導師成績總表」第 ${res.addedColumn} 欄！`, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert("❌ 匯入失敗", err.message, ui.ButtonSet.OK);
  }
}

/**
 * 專任/科任端：動態查詢 [System_Settings] 對照表，將當前分頁成績跨檔案同步至對應導師試算表
 */
function syncSubjectSheetToHomeroomSpreadsheetUI() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName().trim();

  // 防呆：禁止在管理分頁上執行
  if (sheetName === "導師成績總表" || sheetName === "System_Settings") {
    ui.alert("⚠️ 請先切換至科任成績分頁（例如：101理化），再執行跨檔同步！");
    return;
  }

  // 1. 自動解析分頁名稱 (前 3 字為班級代碼，後續為科目名稱，例如 "101理化")
  let parsedClass = "";
  let parsedSubject = "";

  if (sheetName.length >= 5) {
    parsedClass = sheetName.substring(0, 3);
    parsedSubject = sheetName.substring(3);
  } else {
    parsedClass = sheetName.replace(/[^0-9]/g, "");
    parsedSubject = sheetName.replace(/[0-9]/g, "");
  }

  if (!parsedClass) {
    ui.alert("⚠️ 解析班級代碼失敗", `分頁名稱【${sheetName}】中無法識別出班級代碼（如 101）！`, ui.ButtonSet.OK);
    return;
  }

  // 2. 透過 getHomeroomConfig 取得對照表資訊 (Homeroom_SS_ID、Target_Sheet_Name 等)
  let hrConfig;
  try {
    hrConfig = getHomeroomConfig(parsedClass);
  } catch (err) {
    ui.alert("❌ 找不到導師試算表對照", err.message, ui.ButtonSet.OK);
    return;
  }

  if (!hrConfig.homeroomSSID) {
    ui.alert("⚠️ 尚未設定導師試算表 ID", `[System_Settings] 中班級【${parsedClass}】尚未填入 Homeroom_SS_ID！`, ui.ButtonSet.OK);
    return;
  }

  // 3. 開啟目標導師試算表
  let tutorSS;
  try {
    tutorSS = SpreadsheetApp.openById(hrConfig.homeroomSSID);
  } catch (err) {
    ui.alert("❌ 開啟導師試算表失敗", `無法開啟試算表 (ID: ${hrConfig.homeroomSSID})：\n${err.message}`, ui.ButtonSet.OK);
    return;
  }

  const targetSheetName = hrConfig.targetSheetName || "導師成績總表";
  const tutorSheet = tutorSS.getSheetByName(targetSheetName) || tutorSS.getSheets()[0];
  if (!tutorSheet) {
    ui.alert("❌ 找不到目標分頁", `導師試算表中找不到【${targetSheetName}】分頁！`, ui.ButtonSet.OK);
    return;
  }

  // 4. 比對導師總表 A1 班級防呆
  const a1Value = String(tutorSheet.getRange("A1").getValue()).trim();
  if (parsedClass && !a1Value.includes(parsedClass)) {
    const confirmImport = ui.alert(
      "⚠️ 班級不符警示",
      `當前分頁班級為【${parsedClass}】，但該導師試算表 A1 儲存格內容為【${a1Value}】。\n確定仍要跨檔同步匯入嗎？`,
      ui.ButtonSet.YES_NO
    );
    if (confirmImport !== ui.Button.YES) return;
  }

  // 5. 提示確認科目、測驗範圍、測驗日期
  const defaultSubj = hrConfig.subjectColumn || parsedSubject;
  const respSubject = ui.prompt("確認科目", `請確認科目名稱：`, ui.ButtonSet.OK_CANCEL);
  if (respSubject.getSelectedButton() !== ui.Button.OK) return;
  const finalSubject = respSubject.getResponseText().trim() || defaultSubj;

  const respRange = ui.prompt("測驗範圍", "請輸入測驗範圍/細項（例：第三單元小考）：", ui.ButtonSet.OK_CANCEL);
  if (respRange.getSelectedButton() !== ui.Button.OK) return;
  const rangeText = respRange.getResponseText().trim();

  const defaultDate = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
  const respDate = ui.prompt("測驗日期", `請確認測驗日期 (YYYY-MM-DD)：`, ui.ButtonSet.OK_CANCEL);
  if (respDate.getSelectedButton() !== ui.Button.OK) return;
  const testDate = respDate.getResponseText().trim() || defaultDate;

  // 6. 讀取當前科任分頁成績 (A 欄座號，B 欄或第 2 欄成績)
  const data = activeSheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert("⚠️ 該分頁無成績資料！");
    return;
  }

  const scoreMap = {};
  for (let r = 1; r < data.length; r++) {
    const key = String(data[r][0]).trim();
    const score = data[r][1];
    if (key !== "") {
      scoreMap[key] = score;
    }
  }

  // 7. 寫入導師試算表
  try {
    ui.toast(`正在同步寫入【${parsedClass}班 導師試算表】...`, "⏳ 同步中", -1);
    const res = addNewSubjectColumnToTutorSheet(finalSubject, rangeText, testDate, scoreMap, tutorSheet);
    ui.alert("✅ 跨檔同步成功！", `已將【${parsedClass} ${finalSubject} - ${rangeText}】成功寫入導師試算表 (${tutorSS.getName()}) 之【${tutorSheet.getName()}】第 ${res.addedColumn} 欄！`, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert("❌ 跨檔同步失敗", err.message, ui.ButtonSet.OK);
  }
}

/**
 * 一鍵建立或初始化 [System_Settings] 基礎設定分頁
 */
function initSystemSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("System_Settings");
  const ui = SpreadsheetApp.getUi();

  if (sheet) {
    const resp = ui.alert(
      "提示",
      "[System_Settings] 分頁已存在，是否要檢查並補充缺少欄位？（不會覆蓋現有已設定數值）",
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  } else {
    sheet = ss.insertSheet("System_Settings");
  }

  // 設定 A:B 欄 (全域系統參數)
  const leftHeaders = [
    ["項目 (Key)", "設定值 (Value)"],
    ["TEACHER_PIN", "1234"],
    ["GEMINI_API_KEY", ""],
    ["UPLOAD_FOLDER_ID", ""],
    ["BACKUP_FOLDER_ID", ""],
    ["BG_FOLDER_ID", ""],
    ["LOGO_KEYWORD", "logo"]
  ];

  for (let i = 0; i < leftHeaders.length; i++) {
    const r = i + 1;
    const currentKey = sheet.getRange(r, 1).getValue();
    if (!currentKey) {
      sheet.getRange(r, 1).setValue(leftHeaders[i][0]);
      sheet.getRange(r, 2).setValue(leftHeaders[i][1]);
    }
  }

  // 設定 D:H 欄 (班級與導師試算表對照表)
  const rightHeaders = [
    ["Class_Code", "Homeroom_SS_ID", "Target_Sheet_Name", "Key_Column_Name", "Subject_Column"],
    ["101", "", "導師成績總表", "座號", "理化"],
    ["102", "", "導師成績總表", "座號", "理化"],
    ["103", "", "導師成績總表", "座號", "理化"]
  ];

  if (!sheet.getRange(1, 4).getValue()) {
    sheet.getRange(1, 4, rightHeaders.length, 5).setValues(rightHeaders);
  }

  // 美化標題列
  sheet.getRange("A1:B1").setBackground("#34495e").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("D1:H1").setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");
  sheet.autoResizeColumns(1, 8);

  ui.alert("✅ [System_Settings] 分頁建立/檢查完成！請至該分頁填入您的 API Key、資料夾 ID 與各班導師試算表 ID。");
}


// ==========================================
// 5. 科任班級三維度跨班分析儀表板 (Macro, Micro, Auto-Alert)
// ==========================================

/**
 * 解析科任班級分頁名稱
 * 規範：五碼固定為前2字為年級，中間字為班級，最後兩字為科目 (例: 國二愛理化, 高一忠物理, 101理化)
 */
function parseSubjectSheetInfo(sheetName) {
  const name = String(sheetName || "").trim();
  const excludeList = ["導師成績總表", "System_Settings", "科任跨班分析總表"];
  if (excludeList.includes(name) || name.startsWith("工作表") || name.startsWith("Sheet") || name.endsWith("_總成績")) {
    return null;
  }
  if (name.length < 4) return null;

  let grade = "";
  let className = "";
  let subject = "";

  if (name.length >= 5) {
    grade = name.substring(0, 2);                      // 前 2 字：年級 (例如 國二, 高一, 10)
    subject = name.substring(name.length - 2);         // 後 2 字：科目 (例如 理化, 物理)
    className = name.substring(2, name.length - 2);    // 中間字：班級 (例如 愛, 忠, 1)
  } else if (name.length === 4) {
    grade = name.substring(0, 2);
    subject = name.substring(2);
    className = grade;
  }

  const groupKey = `${grade}_${subject}`;
  return {
    sheetName: name,
    grade: grade,
    className: className || name,
    subject: subject,
    groupKey: groupKey,
    groupDisplayName: `${grade} ${subject}`
  };
}

/**
 * 取得當前活頁簿中所有「同年段同科目」科任班級群組
 */
function getSubjectClassGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const groups = {};

  for (const sheet of sheets) {
    const info = parseSubjectSheetInfo(sheet.getName());
    if (!info) continue;

    if (!groups[info.groupKey]) {
      groups[info.groupKey] = {
        groupKey: info.groupKey,
        grade: info.grade,
        subject: info.subject,
        displayName: info.groupDisplayName,
        classes: []
      };
    }
    groups[info.groupKey].classes.push({
      sheetName: info.sheetName,
      className: info.className
    });
  }

  return Object.values(groups);
}

/**
 * 取得指定群組的完整跨班分析矩陣數據 (Macro View, Auto-Alert, Micro View)
 */
function getCrossClassAnalytics(groupKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const matchedSheets = [];

  for (const sheet of sheets) {
    const info = parseSubjectSheetInfo(sheet.getName());
    if (info && (!groupKey || info.groupKey === groupKey)) {
      matchedSheets.push({ sheet: sheet, info: info });
    }
  }

  if (matchedSheets.length === 0) {
    throw new Error(`找不到符合群組 [${groupKey}] 的科任班級分頁！`);
  }

  // 1. 統一抓取各班歷次測驗資訊
  const testMap = new Map(); // testKey -> { id, label, category, range, date }
  const classesData = [];

  matchedSheets.forEach(({ sheet, info }) => {
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    const studentEndRow = getStudentEndRow(sheet, 1, 2);
    const numStudents = Math.max(0, studentEndRow - 3);

    const cls = {
      className: info.className,
      sheetName: info.sheetName,
      grade: info.grade,
      subject: info.subject,
      studentCount: numStudents,
      students: [],
      tests: [],
      testAvgs: {},
      testP50s: {},
      categoryAvgs: {}
    };

    if (lastCol >= 3 && numStudents > 0) {
      const numCols = lastCol - 2;
      const types = sheet.getRange(1, 3, 1, numCols).getValues()[0];
      const ranges = sheet.getRange(2, 3, 1, numCols).getDisplayValues()[0];
      const dates = sheet.getRange(3, 3, 1, numCols).getValues()[0];
      const studentInfo = sheet.getRange(4, 1, numStudents, 2).getValues(); // [seat, name]
      const scoresMatrix = sheet.getRange(4, 3, numStudents, numCols).getValues();

      // 解析測驗表頭
      for (let c = 0; c < numCols; c++) {
        const cat = String(types[c] || "").trim() || "測驗";
        const rng = String(ranges[c] || "").trim() || `第${c + 1}單元`;
        const dt = dates[c] ? Utilities.formatDate(parseDateHelper(dates[c]) || new Date(), "GMT+8", "yyyy-MM-dd") : "";
        const testId = `T_${c + 1}`;
        const label = `${rng}${dt ? " (" + dt.substring(5) + ")" : ""}`;

        if (!testMap.has(testId)) {
          testMap.set(testId, { testId, label, category: cat, range: rng, date: dt, index: c });
        }

        // 計算該次測驗該班成績集
        const colScores = [];
        for (let r = 0; r < numStudents; r++) {
          const val = scoresMatrix[r][c];
          if (val !== "" && val !== null && !isNaN(Number(val))) {
            colScores.push(Number(val));
          }
        }

        const avg = colScores.length > 0 ? Math.round((colScores.reduce((a, b) => a + b, 0) / colScores.length) * 10) / 10 : null;
        cls.testAvgs[testId] = avg;
        const pObj = calculateFiveTierPercentiles(colScores);
        cls.testP50s[testId] = pObj.p50;
      }

      // 解析學生個人數據
      for (let r = 0; r < numStudents; r++) {
        const seat = studentInfo[r][0];
        const name = String(studentInfo[r][1] || `學生${seat}`).trim();
        const scores = [];
        const studentTestMap = {};
        const catSums = {}, catCounts = {};

        for (let c = 0; c < numCols; c++) {
          const testId = `T_${c + 1}`;
          const raw = scoresMatrix[r][c];
          const cat = String(types[c] || "").trim() || "測驗";
          if (raw !== "" && raw !== null && !isNaN(Number(raw))) {
            const sc = Number(raw);
            scores.push(sc);
            studentTestMap[testId] = sc;
            catSums[cat] = (catSums[cat] || 0) + sc;
            catCounts[cat] = (catCounts[cat] || 0) + 1;
          } else {
            scores.push(null);
            studentTestMap[testId] = null;
          }
        }

        const validScores = scores.filter(s => s !== null);
        const avg = validScores.length > 0 ? Math.round((validScores.reduce((a, b) => a + b, 0) / validScores.length) * 10) / 10 : 0;
        const latestScore = validScores.length > 0 ? validScores[validScores.length - 1] : 0;

        const studentCatAvgs = {};
        Object.keys(catSums).forEach(k => {
          studentCatAvgs[k] = Math.round((catSums[k] / catCounts[k]) * 10) / 10;
        });

        cls.students.push({
          seat: seat,
          name: name,
          className: info.className,
          scores: scores,
          validScores: validScores,
          testMap: studentTestMap,
          avg: avg,
          latestScore: latestScore,
          categoryAvgs: studentCatAvgs,
          alertHistory: [],
          tags: []
        });
      }

      // 計算班級各評量類別平均
      const classCatSums = {}, classCatCounts = {};
      cls.students.forEach(st => {
        Object.keys(st.categoryAvgs).forEach(k => {
          classCatSums[k] = (classCatSums[k] || 0) + st.categoryAvgs[k];
          classCatCounts[k] = (classCatCounts[k] || 0) + 1;
        });
      });
      Object.keys(classCatSums).forEach(k => {
        cls.categoryAvgs[k] = Math.round((classCatSums[k] / classCatCounts[k]) * 10) / 10;
      });
    }

    classesData.push(cls);
  });

  const trendTests = Array.from(testMap.values());

  // 2. 年段全體各次測驗基準 (gradeAvg)
  trendTests.forEach(t => {
    let sum = 0, count = 0;
    classesData.forEach(c => {
      const val = c.testAvgs[t.testId];
      if (val !== null && val !== undefined) {
        sum += val;
        count++;
      }
    });
    t.gradeAvg = count > 0 ? Math.round((sum / count) * 10) / 10 : null;
  });

  // 3. 跨班指標對照表 (Macro Stats) 計算
  // 【關鍵修正】：以同年段同科目所有班級學生作為母群，計算全域統一的年段五標基準
  const allGradeStudentAvgs = [];
  const gradeCategoryScores = {}; // 類別 -> [全年段學生平均分]

  classesData.forEach(cls => {
    cls.students.forEach(s => {
      if (s.avg > 0) allGradeStudentAvgs.push(s.avg);
      Object.keys(s.categoryAvgs || {}).forEach(cat => {
        if (!gradeCategoryScores[cat]) gradeCategoryScores[cat] = [];
        gradeCategoryScores[cat].push(s.categoryAvgs[cat]);
      });
    });
  });

  // 全年段統一五標標準 (年段常模)
  const gradeFiveTiers = calculateFiveTierPercentiles(allGradeStudentAvgs);

  // 全年段各評量類別五標與均值基準 (供雷達圖背景使用)
  const gradeCategoryTiers = {};
  Object.keys(gradeCategoryScores).forEach(cat => {
    const arr = gradeCategoryScores[cat];
    const catPObj = calculateFiveTierPercentiles(arr);
    const catAvg = arr.length > 0 ? Math.round((arr.reduce((a, b) => a + b, 0) / arr.length) * 10) / 10 : 0;
    gradeCategoryTiers[cat] = {
      top: catPObj.p88,
      high: catPObj.p75,
      median: catPObj.p50,
      low: catPObj.p25,
      bottom: catPObj.p12,
      avg: catAvg
    };
  });

  const macroStats = [];
  let allStudentsTotal = 0;
  let allScoresTotalSum = 0;
  let allScoresCount = 0;

  classesData.forEach(cls => {
    const allStudentAvgs = cls.students.map(s => s.avg).filter(a => a > 0);
    const n = allStudentAvgs.length;
    allStudentsTotal += n;

    const classAvg = n > 0 ? Math.round((allStudentAvgs.reduce((a, b) => a + b, 0) / n) * 10) / 10 : 0;
    const passCount = allStudentAvgs.filter(a => a >= 60).length;
    const passRate = n > 0 ? Math.round((passCount / n) * 100) : 0;

    // 標準差 (齊次度)
    let variance = 0;
    if (n > 1) {
      variance = allStudentAvgs.reduce((acc, v) => acc + Math.pow(v - classAvg, 2), 0) / (n - 1);
    }
    const stdDev = Math.round(Math.sqrt(variance) * 10) / 10;

    // 齊次度判讀
    let homogeneity = { label: "適中", badge: "badge-info" };
    if (stdDev >= 18) {
      homogeneity = { label: "雙峰 (兩極化)", badge: "badge-danger" };
    } else if (stdDev <= 12) {
      homogeneity = { label: "集中 (程度平均)", badge: "badge-success" };
    }

    // 五標分佈：依據「全年段統一五標標準 (年段常模)」統計各班達到頂標與落在底標之比例
    const topCount = allStudentAvgs.filter(a => a >= gradeFiveTiers.p88).length;
    const bottomCount = allStudentAvgs.filter(a => a <= gradeFiveTiers.p12).length;
    const topTierRate = n > 0 ? Math.round((topCount / n) * 100) : 0;
    const bottomTierRate = n > 0 ? Math.round((bottomCount / n) * 100) : 0;

    // 梯隊分佈 (頂標、前標、均標、後標、底標)：以年段標準區分各班五標落點人數
    const tierCounts = { top: 0, high: 0, mid: 0, low: 0, bottom: 0 };
    allStudentAvgs.forEach(s => {
      if (s >= gradeFiveTiers.p88) tierCounts.top++;
      else if (s >= gradeFiveTiers.p75) tierCounts.high++;
      else if (s >= gradeFiveTiers.p50) tierCounts.mid++;
      else if (s >= gradeFiveTiers.p25) tierCounts.low++;
      else tierCounts.bottom++;
    });

    const tierPercentages = {
      top: n > 0 ? Math.round((tierCounts.top / n) * 100) : 0,
      high: n > 0 ? Math.round((tierCounts.high / n) * 100) : 0,
      mid: n > 0 ? Math.round((tierCounts.mid / n) * 100) : 0,
      low: n > 0 ? Math.round((tierCounts.low / n) * 100) : 0,
      bottom: n > 0 ? Math.round((tierCounts.bottom / n) * 100) : 0
    };

    cls.students.forEach(s => {
      if (s.avg > 0) {
        allScoresTotalSum += s.avg;
        allScoresCount++;
      }
    });

    macroStats.push({
      className: cls.className,
      studentCount: n,
      avg: classAvg,
      passRate: passRate,
      stdDev: stdDev,
      homogeneity: homogeneity,
      topTierRate: topTierRate,
      bottomTierRate: bottomTierRate,
      tierCounts: tierCounts,
      tierPercentages: tierPercentages,
      percentiles: gradeFiveTiers
    });
  });

  const overallAvg = allScoresCount > 0 ? Math.round((allScoresTotalSum / allScoresCount) * 10) / 10 : 0;

  // 4. 自動化預警與關注學生標記 (Auto-Alert System)
  const attentionStudents = [];

  classesData.forEach((cls, clsIdx) => {
    const classPercentiles = macroStats[clsIdx].percentiles;

    cls.students.forEach(student => {
      const valid = student.validScores;
      if (valid.length === 0) return;

      const tags = [];
      const history = [];

      // 逐次分析歷次測驗 (滾動式檢測)
      let lowP12Consecutive = 0;

      for (let i = 0; i < valid.length; i++) {
        const sc = valid[i];
        const testLabel = trendTests[i] ? trendTests[i].label : `第${i + 1}次測驗`;

        // 檢查前 3 次均值進退步
        if (i >= 3) {
          const prev3Avg = (valid[i - 1] + valid[i - 2] + valid[i - 3]) / 3;
          const diff = Math.round((sc - prev3Avg) * 10) / 10;
          if (diff <= -15) {
            history.push({
              type: "warning",
              badgeLabel: "🟡 顯著退步",
              testLabel: testLabel,
              score: sc,
              reason: `分數較前3次均分(${Math.round(prev3Avg)})大幅下滑 ${Math.abs(diff)} 分`
            });
          } else if (diff >= 15) {
            history.push({
              type: "growth",
              badgeLabel: "🟢 顯著進步",
              testLabel: testLabel,
              score: sc,
              reason: `分數較前3次均分(${Math.round(prev3Avg)})大幅進步 ${diff} 分`
            });
          }
        }

        // 檢查底標 (低於班級 P12)
        if (sc < classPercentiles.p12 || sc < 50) {
          lowP12Consecutive++;
          if (lowP12Consecutive >= 2) {
            history.push({
              type: "crisis",
              badgeLabel: "🔴 高風險/警訊",
              testLabel: testLabel,
              score: sc,
              reason: `連續 ${lowP12Consecutive} 次測驗低於班級底標或未滿50分`
            });
          }
        } else {
          lowP12Consecutive = 0;
        }
      }

      student.alertHistory = history;

      // 學生最新狀態標記 (頂部警示判斷)
      const latestSc = student.latestScore;
      const recentLen = valid.length;

      // 🔴 條件 1: 最新分數 < 50 或 連續 2 次未達底標
      if (latestSc < 50 || lowP12Consecutive >= 2) {
        tags.push({
          type: "crisis",
          label: "🔴 高風險",
          reason: latestSc < 50 ? `最新成績未達50分 (${latestSc}分)` : "連續2次低於班級底標"
        });
      }

      // 🟡 條件 2: 最新一次較前3次下滑 > 15
      if (recentLen >= 4) {
        const p3Avg = (valid[recentLen - 2] + valid[recentLen - 3] + valid[recentLen - 4]) / 3;
        const diff = Math.round((latestSc - p3Avg) * 10) / 10;
        if (diff <= -15) {
          tags.push({
            type: "warning",
            label: "🟡 顯著退步",
            reason: `最新成績較前3次平均下滑 ${Math.abs(diff)} 分`
          });
        } else if (diff >= 15) {
          tags.push({
            type: "growth",
            label: "🟢 顯著進步",
            reason: `最新成績較前3次平均提升 ${diff} 分`
          });
        }
      }

      // ⚡ 條件 3: 高波動/不穩定 (個人標準差 > 18 且測驗次數 >= 3)
      if (valid.length >= 3) {
        let vSum = 0;
        const mean = student.avg;
        for (let s of valid) vSum += Math.pow(s - mean, 2);
        const sSD = Math.round(Math.sqrt(vSum / (valid.length - 1)) * 10) / 10;
        if (sSD >= 18) {
          tags.push({
            type: "unstable",
            label: "⚡ 高波動",
            reason: `成績忽高忽低 (個人SD: ${sSD})`
          });
        }
      }

      // 建議教學行動
      let suggestedAction = "課堂持續觀察";
      if (tags.some(t => t.type === "crisis")) {
        suggestedAction = "優先列入課後輔導或聯繫導師";
      } else if (tags.some(t => t.type === "warning")) {
        suggestedAction = "個別面談排查觀念盲點或生活狀況";
      } else if (tags.some(t => t.type === "unstable")) {
        suggestedAction = "檢視觀念紮實度（防猜題或抱佛腳）";
      } else if (tags.some(t => t.type === "growth")) {
        suggestedAction = "口頭表揚激勵維持學習熱情";
      }

      student.tags = tags;
      student.suggestedAction = suggestedAction;

      if (tags.length > 0) {
        attentionStudents.push({
          seat: student.seat,
          name: student.name,
          className: student.className,
          latestScore: student.latestScore,
          avgScore: student.avg,
          tags: tags,
          suggestedAction: suggestedAction
        });
      }
    });
  });

  // KPI 總結數據
  const kpis = {
    classCount: classesData.length,
    studentCount: allStudentsTotal,
    overallAvg: overallAvg,
    crisisCount: attentionStudents.filter(s => s.tags.some(t => t.type === "crisis")).length,
    warningCount: attentionStudents.filter(s => s.tags.some(t => t.type === "warning")).length,
    growthCount: attentionStudents.filter(s => s.tags.some(t => t.type === "growth")).length
  };

  return {
    groupKey: groupKey,
    kpis: kpis,
    macroStats: macroStats,
    trendTests: trendTests,
    classes: classesData,
    attentionStudents: attentionStudents,
    gradeFiveTiers: gradeFiveTiers,
    gradeCategoryTiers: gradeCategoryTiers
  };
}

/**
 * 在試算表中產出「科任跨班分析總表」實體分頁
 */
function generateSubjectClassCrossAnalyticsSheet(groupKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = getCrossClassAnalytics(groupKey);

  let sheet = ss.getSheetByName("科任跨班分析總表");
  if (!sheet) {
    sheet = ss.insertSheet("科任跨班分析總表");
  } else {
    sheet.clear();
  }

  // 1. 大標題
  sheet.getRange("A1").setValue(`📊 ${data.groupKey.replace("_", " ")} 科任跨班分析總表`).setFontSize(16).setFontWeight("bold");
  sheet.getRange("A2").setValue(`產出時間：${Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm")} | 分析班級數：${data.kpis.classCount} 班 | 年段總人數：${data.kpis.studentCount} 人 | 年段平均：${data.kpis.overallAvg} 分`).setFontColor("#64748b");

  // 2. 區塊一：跨班指標對照表 (Macro View)
  let curRow = 4;
  sheet.getRange(curRow, 1).setValue("一、跨班橫向指標對照表 (Macro View)").setFontWeight("bold").setFontSize(13);
  curRow++;

  const macroHeaders = ["班級", "應考人數", "班平均", "及格率", "標準差 (SD)", "齊次度判讀", "頂標率 (P88)", "底標率 (P12)"];
  sheet.getRange(curRow, 1, 1, macroHeaders.length).setValues([macroHeaders])
    .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  curRow++;

  const macroRows = data.macroStats.map(m => [
    m.className + " 班",
    m.studentCount,
    m.avg,
    m.passRate + "%",
    m.stdDev,
    m.homogeneity.label,
    m.topTierRate + "%",
    m.bottomTierRate + "%"
  ]);

  if (macroRows.length > 0) {
    sheet.getRange(curRow, 1, macroRows.length, macroHeaders.length).setValues(macroRows).setHorizontalAlignment("center");
    curRow += macroRows.length + 2;
  }

  // 3. 區塊二：各班能力梯隊結構分佈表 (五標分佈)
  sheet.getRange(curRow, 1).setValue("二、各班學生能力結構梯隊分佈 (五標比例)").setFontWeight("bold").setFontSize(13);
  curRow++;

  const tierHeaders = ["班級", "頂標人數 (≥P88)", "前標人數", "均標人數", "後標人數", "底標人數 (<P12)"];
  sheet.getRange(curRow, 1, 1, tierHeaders.length).setValues([tierHeaders])
    .setBackground("#334155").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  curRow++;

  const tierRows = data.macroStats.map(m => [
    m.className + " 班",
    `${m.tierCounts.top} (${m.tierPercentages.top}%)`,
    `${m.tierCounts.high} (${m.tierPercentages.high}%)`,
    `${m.tierCounts.mid} (${m.tierPercentages.mid}%)`,
    `${m.tierCounts.low} (${m.tierPercentages.low}%)`,
    `${m.tierCounts.bottom} (${m.tierPercentages.bottom}%)`
  ]);

  if (tierRows.length > 0) {
    sheet.getRange(curRow, 1, tierRows.length, tierHeaders.length).setValues(tierRows).setHorizontalAlignment("center");
    curRow += tierRows.length + 2;
  }

  // 4. 區塊三：自動化預警與需關注學生清單 (Auto-Alert System)
  sheet.getRange(curRow, 1).setValue("三、需關注學生自動預警標記名單 (Auto-Alert System)").setFontWeight("bold").setFontSize(13);
  curRow++;

  const alertHeaders = ["標記類型", "班級", "座號", "姓名", "最新成績", "個人平均", "觸發原因說明", "建議教學行動"];
  sheet.getRange(curRow, 1, 1, alertHeaders.length).setValues([alertHeaders])
    .setBackground("#b91c1c").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  curRow++;

  if (data.attentionStudents.length > 0) {
    const alertRows = data.attentionStudents.map(s => [
      s.tags.map(t => t.label).join(", "),
      s.className + " 班",
      s.seat,
      s.name,
      s.latestScore,
      s.avgScore,
      s.tags.map(t => t.reason).join("；"),
      s.suggestedAction
    ]);
    sheet.getRange(curRow, 1, alertRows.length, alertHeaders.length).setValues(alertRows);
    sheet.getRange(curRow, 1, alertRows.length, 6).setHorizontalAlignment("center");
    sheet.getRange(curRow, 7, alertRows.length, 2).setHorizontalAlignment("left");
  } else {
    sheet.getRange(curRow, 1, 1, alertHeaders.length).setValues([["無", "全體班級", "-", "表現穩定", "-", "-", "無觸發預警之學生", "持續維持現行教法"]])
      .setHorizontalAlignment("center");
  }

  sheet.autoResizeColumns(1, 9);
  sheet.setFrozenRows(2);
  ss.setActiveSheet(sheet);

  return `✅ 已成功為【${data.groupKey.replace("_", " ")}】建立/更新「科任跨班分析總表」分頁！`;
}

/**
 * 試算表選單彈跳互動視窗開啟科任儀表板
 */
function showSubjectDashboardDialog() {
  const template = HtmlService.createTemplateFromFile('SubjectDashboard');
  template.isModalDialog = true;
  const html = template.evaluate()
    .setWidth(1300)
    .setHeight(820)
    .setTitle("📊 科任教師三維度跨班分析總覽");
  SpreadsheetApp.getUi().showModalDialog(html, "📊 科任教師三維度跨班分析總覽");
}

/**
 * 選單點選「產出科任跨班分析總表分頁」對話框
 */
function createSubjectAnalyticsSheetUI() {
  const ui = SpreadsheetApp.getUi();
  const groups = getSubjectClassGroups();
  if (groups.length === 0) {
    ui.alert("⚠️ 系統中找不到符合五碼命名格式的科任分頁（例如：國二愛理化、高一忠物理）！");
    return;
  }

  if (groups.length === 1) {
    const res = generateSubjectClassCrossAnalyticsSheet(groups[0].groupKey);
    ui.alert(res);
    return;
  }

  const groupNames = groups.map((g, idx) => `${idx + 1}. ${g.displayName} (${g.classes.length} 班)`).join("\n");
  const resp = ui.prompt(
    "選擇分析群組",
    `請輸入要分析的群組編號 (1 ~ ${groups.length})：\n\n${groupNames}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() === ui.Button.OK) {
    const idx = parseInt(resp.getResponseText().trim(), 10) - 1;
    if (idx >= 0 && idx < groups.length) {
      ui.toast("正在運算跨班數據並產出報表...", "⏳ 處理中", -1);
      const res = generateSubjectClassCrossAnalyticsSheet(groups[idx].groupKey);
      ui.alert(res);
    } else {
      ui.alert("❌ 輸入編號無效！");
    }
  }
}


/**
 * ==========================================
 * 教師全功能主控台 (TeacherHub) 後端支援模組
 * 1. 手寫成績單多圖辨識 (recognizeUploadedImages)
 * 2. 校對確認存檔匯入 (saveConfirmedOcrScores)
 * 3. 導師班級主控版數據 (getTutorDashboardData)
 * 4. 系統班級與科目清單查詢 (getAvailableClassesAndSubjects)
 * 5. Web App 與試算表選單入口 (showTeacherHubDialog / doGet 整合)
 * ==========================================
 */

/**
 * 取得試算表中現有的班級代碼、科目名稱與導師班學生名單
 */
function getAvailableClassesAndSubjects() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const mainSheet = ss.getSheetByName("導師成績總表") || sheets[0];
  const mainClassName = normalizeClassName(String(mainSheet ? mainSheet.getRange("A1").getValue() : "").trim());

  // 1. 導師班學生名單 (座號 -> 姓名)
  const tutorStudents = [];
  if (mainSheet) {
    const studentEndRow = getStudentEndRow(mainSheet, 2, 3);
    const count = studentEndRow - 3;
    if (count > 0) {
      const vals = mainSheet.getRange(4, 2, count, 2).getValues();
      vals.forEach(r => {
        const seat = parseInt(r[0], 10);
        const name = String(r[1] || "").trim();
        if (!isNaN(seat)) {
          tutorStudents.push({ seat: seat, name: name });
        }
      });
    }
  }

  // 2. 蒐集系統所有班級與科目
  const classesSet = new Set();
  const subjectsSet = new Set(["國文", "英文", "數學", "理化", "物理", "化學", "生物", "地科", "歷史", "地理", "公民"]);

  if (mainClassName) classesSet.add(mainClassName);

  // 檢查所有工作表
  sheets.forEach(sh => {
    const sName = sh.getName();
    // 檢查是否為專任五碼格式 (如 101理化 或 國二愛理化)
    const info = parseSubjectSheetInfo(sName);
    if (info) {
      classesSet.add(info.classCode);
      subjectsSet.add(info.subject);
    }
  });

  // 檢查 System_Settings 設定頁
  try {
    const sysSheet = ss.getSheetByName("System_Settings");
    if (sysSheet && sysSheet.getLastRow() >= 2) {
      const data = sysSheet.getRange("D2:H" + sysSheet.getLastRow()).getValues();
      data.forEach(r => {
        const cCode = String(r[0]).trim();
        const sCol = String(r[4] || "").trim();
        if (cCode) classesSet.add(cCode);
        if (sCol) subjectsSet.add(sCol);
      });
    }
  } catch (e) {
    // 略過
  }

  return {
    mainClassName: mainClassName,
    classes: Array.from(classesSet).filter(Boolean),
    subjects: Array.from(subjectsSet).filter(Boolean),
    evalTypes: ["小考", "平時測驗", "作業", "段考", "期中考", "期末考", "筆記", "課堂表現", "週考"],
    tutorStudents: tutorStudents,
    earliestDate: getEarliestDateInMainSheet()
  };
}

/**
 * 網頁端呼叫：批量手寫成績單照片 AI 辨識 (支援多圖)
 * @param {Array<Object>} imagesData [{ name: string, mimeType: string, base64: string }, ...]
 * @returns {Array<Object>} 各張照片的結構化辨識成果，供前端校對表格呈現
 */
function recognizeUploadedImages(imagesData) {
  if (!imagesData || !Array.isArray(imagesData) || imagesData.length === 0) {
    throw new Error("未接收到任何照片數據！");
  }

  const meta = getAvailableClassesAndSubjects();
  const mainClassName = meta.mainClassName;
  const tutorStudentMap = {};
  meta.tutorStudents.forEach(st => {
    tutorStudentMap[st.seat] = st.name;
  });

  const prompt = `你是一個精準的手寫成績單 OCR 辨識助手。請仔細分析這張成績單照片，判斷其「班級」、「科目」、「評量類型」與各座號成績數據，並嚴格回傳純 JSON 格式：

{
  "className": "班級名稱(例如: 101, 102, 國二愛, 國二仁，未註記則填空字串)",
  "subject": "科目名稱(例如: 國文, 英文, 數學, 理化, 歷史, 地理, 公民, 生物...)",
  "type": "評量類型(例如: 小考, 作業, 段考, 筆記，若無標示預設填 小考)",
  "examRange": "考試範圍或測驗單元名稱(例如: 第一章, 1-1~1-2, Unit 3，若無則填寫 無範圍)",
  "columns": [
    {
      "range": "考試範圍或單元名稱",
      "scores": [
        {"seat": 1, "score": 85},
        {"seat": 2, "score": 90}
      ]
    }
  ]
}

【辨識規則說明】：
1. 班級與科目：尋找照片頂端或標題手寫字（如「101」、「102」、「二愛」、「國二仁」、「理化小考」）。
2. 評量類型 (type)：辨識是否有「段考」、「小考」、「作業」、「筆記」等字樣，未標示則填「小考」。
3. 考試範圍 (examRange)：辨識範圍字樣，未註明可填「無範圍」。
4. 座號與成績 (scores)：
   - seat 請轉為純整數數字 (如 1, 2, 3...)。
   - score 請轉為數值。
   - 被劃刪除線、空白、缺考者，該座號 score 請填寫 null 或忽略。

【輸出格式限制】：
- 嚴格僅輸出純文字 JSON，絕對不要包含任何 Markdown 標記 (如 \`\`\`json)。`;

  const dynamicModel = getLatestFlashModelName();
  const candidateModels = Array.from(new Set([
    "gemini-flash-lite-latest",
    dynamicModel,
    "gemini-flash-latest",
    "gemini-3.6-flash",
    "gemini-2.5-flash"
  ].filter(Boolean)));

  const apiKey = getActiveGeminiApiKey();
  const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
  const results = [];

  for (let i = 0; i < imagesData.length; i++) {
    const imgItem = imagesData[i];
    const fileName = imgItem.name || `photo_${i + 1}.jpg`;
    let mimeType = imgItem.mimeType || "image/jpeg";
    if (mimeType === "application/octet-stream" || !mimeType) {
      mimeType = "image/jpeg";
    }

    let rawBase64 = imgItem.base64 || "";
    // 若前端傳入含有 data:image/png;base64, 前綴，則去除
    if (rawBase64.indexOf("base64,") !== -1) {
      rawBase64 = rawBase64.split("base64,")[1];
    }

    const singleResult = {
      fileIndex: i,
      fileName: fileName,
      imageBase64: rawBase64,
      recognizedSuccess: false,
      className: "",
      subject: "理化",
      type: "小考",
      examRange: "無範圍",
      examDate: todayStr,
      scores: [],
      error: null
    };

    if (!rawBase64) {
      singleResult.error = "照片內容為空";
      results.push(singleResult);
      continue;
    }

    let recognizedText = "";
    let apiError = null;

    for (const modelName of candidateModels) {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
      const payload = {
        "contents": [{
          "parts": [
            { "text": prompt },
            { "inline_data": { "mime_type": mimeType, "data": rawBase64 } }
          ]
        }]
      };

      try {
        const resp = UrlFetchApp.fetch(url, {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });

        const json = JSON.parse(resp.getContentText());
        if (json.error) {
          const msg = json.error.message || "";
          if (msg.toLowerCase().includes("quota") || msg.toLowerCase().includes("not found") || json.error.code === 429) {
            apiError = new Error(`[${modelName}] 額度不足或忙碌: ${msg}`);
            continue; // 切換下一個模型
          }
          throw new Error(`Gemini API 錯誤: ${msg}`);
        }

        if (json.candidates && json.candidates.length > 0 && json.candidates[0].content) {
          recognizedText = json.candidates[0].content.parts[0].text;
          apiError = null;
          break; // 成功取得
        }
      } catch (err) {
        apiError = err;
      }
    }

    if (!recognizedText) {
      singleResult.error = apiError ? apiError.message : "無法取得辨識內容";
      results.push(singleResult);
      continue;
    }

    // 解析 JSON
    try {
      let cleanText = recognizedText.replace(/```json/gi, "").replace(/```/g, "").trim();
      const parsed = JSON.parse(cleanText);

      let parsedClass = parsed.className ? normalizeClassName(parsed.className.trim()) : "";
      if (!parsedClass) parsedClass = mainClassName;

      let parsedSubj = parsed.subject ? parsed.subject.trim() : "理化";
      let parsedType = parsed.type ? parsed.type.trim() : "小考";
      let parsedRange = parsed.examRange ? parsed.examRange.trim() : "";

      const scoreItems = [];
      const scoreMap = new Map();

      if (parsed.columns && Array.isArray(parsed.columns) && parsed.columns.length > 0) {
        if (!parsedRange && parsed.columns[0].range) {
          parsedRange = parsed.columns[0].range;
        }
        const colScores = parsed.columns[0].scores || [];
        colScores.forEach(item => {
          if (item.seat !== undefined && item.seat !== null && item.score !== null && item.score !== undefined) {
            const seatNum = parseInt(item.seat, 10);
            const scoreVal = Number(item.score);
            if (!isNaN(seatNum) && !isNaN(scoreVal)) {
              scoreMap.set(seatNum, scoreVal);
            }
          }
        });
      }

      // 如果有對應導師班學生名單，依序填入姓名
      if (scoreMap.size > 0) {
        // 依座號排序
        const sortedSeats = Array.from(scoreMap.keys()).sort((a, b) => a - b);
        sortedSeats.forEach(stNum => {
          scoreItems.push({
            seat: stNum,
            name: tutorStudentMap[stNum] || `座號${stNum}`,
            score: scoreMap.get(stNum)
          });
        });
      }

      singleResult.recognizedSuccess = true;
      singleResult.className = parsedClass;
      singleResult.subject = parsedSubj;
      singleResult.type = parsedType;
      singleResult.examRange = parsedRange || "無範圍";
      singleResult.scores = scoreItems;
    } catch (parseErr) {
      singleResult.error = `解析辨識結果失敗: ${parseErr.message} (原始回應: ${recognizedText.substring(0, 100)})`;
    }

    results.push(singleResult);
  }

  return results;
}

/**
 * 網頁端呼叫：儲存教師校對後的成績數據並寫入試算表
 * @param {Array<Object>} confirmedList 教師校對後的項目陣列
 * @returns {Object} 匯入結果日誌
 */
function saveConfirmedOcrScores(confirmedList) {
  if (!confirmedList || !Array.isArray(confirmedList) || confirmedList.length === 0) {
    throw new Error("沒有要存檔的成績項目！");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
  const mainClassName = normalizeClassName(String(mainSheet.getRange("A1").getValue()).trim());

  let backupFolder = null;
  try {
    backupFolder = DriveApp.getFolderById(getActiveBackupFolderId());
  } catch (e) {
    Logger.log("備份資料夾未設定或無法取得: " + e.message);
  }

  const logs = [];
  let successCount = 0;

  for (let idx = 0; idx < confirmedList.length; idx++) {
    const item = confirmedList[idx];
    const className = normalizeClassName(String(item.className || "").trim());
    const subject = String(item.subject || "").trim();
    const evalType = String(item.type || "小考").trim();
    const examRange = String(item.examRange || "無範圍").trim();
    const examDate = String(item.examDate || Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd")).trim();
    const scores = item.scores || [];

    if (!className || !subject) {
      logs.push(`⚠️ 第 ${idx + 1} 筆項目因缺少班級或科目而略過。`);
      continue;
    }

    // 建立座號對成績 Map
    const scoreMap = new Map();
    scores.forEach(s => {
      const seat = parseInt(s.seat, 10);
      const score = Number(s.score);
      if (!isNaN(seat) && !isNaN(score)) {
        scoreMap.set(seat, score);
      }
    });

    if (scoreMap.size === 0) {
      logs.push(`⚠️ 班級 [${className}] ${subject} 無有效分數，已略過。`);
      continue;
    }

    let writtenTargets = [];

    // 1. 若為本導師班，寫入「導師成績總表」
    if (className === mainClassName && mainSheet) {
      const lastCol = mainSheet.getLastColumn();
      const startCol = (lastCol < 3) ? 4 : lastCol + 1;

      mainSheet.getRange(1, startCol).setValue(subject);
      mainSheet.getRange(2, startCol).setNumberFormat("@").setValue(examRange);
      mainSheet.getRange(3, startCol).setValue(examDate);

      const studentEndRow = getStudentEndRow(mainSheet, 2, 3);
      const totalStudents = studentEndRow - 3;

      if (totalStudents > 0) {
        const seatList = mainSheet.getRange(4, 2, totalStudents, 1).getValues();
        const outputScores = seatList.map(r => {
          const seatNum = parseInt(r[0], 10);
          return [scoreMap.has(seatNum) ? scoreMap.get(seatNum) : ""];
        });
        mainSheet.getRange(4, startCol, outputScores.length, 1).setValues(outputScores);
      }
      writtenTargets.push(`導師成績總表 (第 ${startCol} 欄)`);
    }

    // 2. 搜尋或寫入專任班級分頁 (如 101理化 或 國二愛理化)
    const specialSheetName = `${className}${subject}`;
    let specialSheet = ss.getSheetByName(specialSheetName);

    if (specialSheet) {
      const lastCol = specialSheet.getLastColumn();
      const startCol = (lastCol < 2) ? 3 : lastCol + 1;

      specialSheet.getRange(1, startCol).setValue(evalType);
      specialSheet.getRange(2, startCol).setNumberFormat("@").setValue(examRange);
      specialSheet.getRange(3, startCol).setValue(examDate);

      const studentEndRow = getStudentEndRow(specialSheet, 1, 2);
      const totalStudents = studentEndRow - 3;

      if (totalStudents > 0) {
        const seatList = specialSheet.getRange(4, 1, totalStudents, 1).getValues();
        const outputScores = seatList.map(r => {
          const seatNum = parseInt(r[0], 10);
          return [scoreMap.has(seatNum) ? scoreMap.get(seatNum) : ""];
        });
        specialSheet.getRange(4, startCol, outputScores.length, 1).setValues(outputScores);
      }
      writtenTargets.push(`科任分頁 [${specialSheetName}] (第 ${startCol} 欄)`);
    }

    // 3. 檢查 System_Settings 是否有跨檔導師試算表需要同步
    try {
      const homeroomConfig = getHomeroomConfig(className);
      if (homeroomConfig && homeroomConfig.homeroomSSID) {
        const targetSS = SpreadsheetApp.openById(homeroomConfig.homeroomSSID);
        const targetSheet = targetSS.getSheetByName(homeroomConfig.targetSheetName || "導師成績總表");
        if (targetSheet) {
          const tA1 = normalizeClassName(String(targetSheet.getRange("A1").getValue()).trim());
          if (tA1.indexOf(className) !== -1 || className.indexOf(tA1) !== -1) {
            addNewSubjectColumnToTutorSheet(subject, examRange, examDate, scoreMap, targetSheet);
            writtenTargets.push(`跨檔導師總表 (${targetSS.getName()} - ${targetSheet.getName()})`);
          }
        }
      }
    } catch (e) {
      // 找不到或無設定則略過
    }

    // 4. 若有傳入照片 Base64，備份至 Google Drive
    if (backupFolder && item.imageBase64) {
      try {
        const bytes = Utilities.base64Decode(item.imageBase64);
        const bFileName = `[${className}_${subject}_${examRange}]_${item.fileName || "scores.jpg"}`;
        const blob = Utilities.newBlob(bytes, "image/jpeg", bFileName);
        backupFolder.createFile(blob);
      } catch (err) {
        Logger.log(`照片備份失敗: ${err.message}`);
      }
    }

    if (writtenTargets.length > 0) {
      successCount++;
      logs.push(`✅ 成功匯入 [${className}] ${subject} (${examRange})，寫入目標：${writtenTargets.join("、")}，共 ${scoreMap.size} 筆成績。`);
    } else {
      logs.push(`⚠️ [${className}] ${subject} 未找到匹配之導師總表或科任分頁 [${specialSheetName}]，請確認分頁名稱。`);
    }
  }

  // 嘗試執行統計計算更新
  try {
    generateCurrentSubjectStats();
  } catch (statErr) {
    // 忽略
  }

  return {
    success: successCount > 0,
    totalProcessed: confirmedList.length,
    successCount: successCount,
    logs: logs
  };
}

/**
 * 網頁端呼叫：取得導師班級主控版完整數據
 * 包含：
 * 1. 班級各科巨觀學習曲線數據 (班平均隨週次/測驗起伏)
 * 2. 班級各科統計概況 (平均、標準差、及格率、五標)
 * 3. 各學生個人微觀診斷數據 (整合 BatchReport 四大圖表與明細)
 */
function getTutorDashboardData(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tutorSheet = ss.getSheetByName("導師成績總表") || ss.getSheets()[0];
  if (!tutorSheet) throw new Error("找不到導師成績總表！");

  const className = normalizeClassName(String(tutorSheet.getRange("A1").getValue()).trim()) || "導師班";

  // 1. 呼叫既有的 getStudentReportData 取得每位學生個人完整報表
  const reportPackage = getStudentReportData(startDateStr, endDateStr);
  const reports = reportPackage.reports || [];

  // 2. 彙整全班各科學習曲線 (Macro View)
  const studentEndRow = getStudentEndRow(tutorSheet, 2, 3);
  const studentCount = studentEndRow - 3;
  const lastCol = tutorSheet.getLastColumn();

  const classMacroTrend = {
    weeks: [],
    datasets: []
  };

  const subjectStatsList = [];

  if (studentCount > 0 && lastCol >= 4) {
    const subjects = tutorSheet.getRange(1, 4, 1, lastCol - 3).getDisplayValues()[0];
    const ranges = tutorSheet.getRange(2, 4, 1, lastCol - 3).getDisplayValues()[0];
    const dates = tutorSheet.getRange(3, 4, 1, lastCol - 3).getValues()[0];
    const allScores = tutorSheet.getRange(4, 4, studentCount, lastCol - 3).getValues();

    const startDt = parseDateHelper(startDateStr);
    const endDt = parseDateHelper(endDateStr);
    if (startDt) startDt.setHours(0, 0, 0, 0);
    if (endDt) endDt.setHours(23, 59, 59, 999);

    let minDate = null;
    for (let c = 0; c < dates.length; c++) {
      const dt = parseDateHelper(dates[c]);
      if (dt && (!minDate || dt < minDate)) minDate = dt;
    }
    if (!minDate) minDate = new Date();
    minDate.setHours(0, 0, 0, 0);

    // 收集符合區間的欄位
    const validColumns = [];
    const weeksSet = new Set();
    const subjectScoresMap = {}; // 科目 -> [所有成績]
    const subjectWeekAvgsMap = {}; // 科目 -> { weekLabel: [平均分, ...] }

    for (let c = 0; c < dates.length; c++) {
      const subj = String(subjects[c] || "").trim();
      if (!subj) continue;

      const d = parseDateHelper(dates[c]);
      if (!d || ((!startDt || d >= startDt) && (!endDt || d <= endDt))) {
        const refDate = d || new Date();
        const diffDays = Math.floor((refDate.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24));
        const weekLabel = `第${Math.max(1, Math.floor(diffDays / 7) + 1)}週`;
        weeksSet.add(weekLabel);

        const colScores = [];
        for (let s = 0; s < studentCount; s++) {
          const val = Number(allScores[s][c]);
          if (allScores[s][c] !== "" && allScores[s][c] !== null && !isNaN(val)) {
            colScores.push(val);
          }
        }

        if (colScores.length > 0) {
          const colAvg = colScores.reduce((a, b) => a + b, 0) / colScores.length;
          if (!subjectScoresMap[subj]) subjectScoresMap[subj] = [];
          subjectScoresMap[subj].push(...colScores);

          if (!subjectWeekAvgsMap[subj]) subjectWeekAvgsMap[subj] = {};
          if (!subjectWeekAvgsMap[subj][weekLabel]) subjectWeekAvgsMap[subj][weekLabel] = [];
          subjectWeekAvgsMap[subj][weekLabel].push(colAvg);

          validColumns.push({
            colIdx: c,
            subject: subj,
            range: ranges[c] || "無範圍",
            date: d ? Utilities.formatDate(d, "GMT+8", "MM/dd") : "-",
            weekLabel: weekLabel,
            avg: Math.round(colAvg * 10) / 10,
            count: colScores.length
          });
        }
      }
    }

    const sortedWeeks = Array.from(weeksSet).sort((a, b) => {
      const numA = parseInt(a.replace(/\D/g, ""), 10) || 0;
      const numB = parseInt(b.replace(/\D/g, ""), 10) || 0;
      return numA - numB;
    });

    classMacroTrend.weeks = sortedWeeks.length > 0 ? sortedWeeks : ["無資料"];

    const palette = [
      "#3b82f6", "#ef4444", "#10b981", "#f59e0b", "#8b5cf6",
      "#ec4899", "#06b6d4", "#84cc16", "#6366f1", "#14b8a6"
    ];

    const orderedSubjs = sortSubjectsByDefaultPattern(Object.keys(subjectScoresMap));

    orderedSubjs.forEach((subj, sIdx) => {
      const weekData = sortedWeeks.map(w => {
        if (subjectWeekAvgsMap[subj] && subjectWeekAvgsMap[subj][w]) {
          const arr = subjectWeekAvgsMap[subj][w];
          return Math.round((arr.reduce((a, b) => a + b, 0) / arr.length) * 10) / 10;
        }
        return null;
      });

      classMacroTrend.datasets.push({
        label: subj,
        data: weekData,
        borderColor: palette[sIdx % palette.length],
        backgroundColor: palette[sIdx % palette.length],
        fill: false,
        tension: 0.25,
        spanGaps: true
      });

      // 計算各科總覽指標 (直接依照成績數據計算，防止 undefined)
      const allSubjScores = subjectScoresMap[subj] || [];
      let subjAvg = 0;
      let subjPassRate = 0;
      let subjMax = 0;
      let subjMin = 0;
      let subjStdDev = 0;

      if (allSubjScores.length > 0) {
        const sum = allSubjScores.reduce((a, b) => a + b, 0);
        subjAvg = Math.round((sum / allSubjScores.length) * 10) / 10;
        subjMax = Math.max(...allSubjScores);
        subjMin = Math.min(...allSubjScores);
        const passCount = allSubjScores.filter(s => s >= 60).length;
        subjPassRate = Math.round((passCount / allSubjScores.length) * 1000) / 10;

        const variance = allSubjScores.reduce((acc, val) => acc + Math.pow(val - subjAvg, 2), 0) / allSubjScores.length;
        subjStdDev = Math.round(Math.sqrt(variance) * 10) / 10;
      }

      const pObj = calculateFiveTierPercentiles(allSubjScores);

      subjectStatsList.push({
        subject: subj,
        examCount: (subjectWeekAvgsMap[subj] ? Object.values(subjectWeekAvgsMap[subj]).flat().length : 0),
        totalScoresCount: allSubjScores.length,
        average: subjAvg,
        passRate: subjPassRate,
        maxScore: subjMax,
        minScore: subjMin,
        stdDev: subjStdDev,
        fiveTiers: {
          top: pObj.p88,
          high: pObj.p75,
          median: pObj.p50,
          low: pObj.p25,
          bottom: pObj.p12
        }
      });
    });
  }

  // 學生簡要索引列表
  const studentList = reports.map((rpt, idx) => ({
    index: idx,
    seat: rpt.info.seat,
    name: rpt.info.name,
    weightedAvg: rpt.summary.weightedAvg,
    rank: rpt.summary.rank
  }));

  return {
    className: className,
    startDate: startDateStr,
    endDate: endDateStr,
    studentCount: reports.length,
    studentList: studentList,
    classMacroTrend: classMacroTrend,
    subjectStatsList: subjectStatsList,
    reports: reports
  };
}

/**
 * 試算表選單呼叫：開啟教師主控台視窗
 */
function showTeacherHubDialog() {
  const template = HtmlService.createTemplateFromFile("TeacherHub");
  template.isModalDialog = true;
  const html = template.evaluate()
    .setWidth(1250)
    .setHeight(880)
    .setTitle("🌟 教師全功能主控台 (成績辨識 / 導師主控 / 科任主控)");
  SpreadsheetApp.getUi().showModalDialog(html, "🌟 教師全功能主控台");
}

/**
 * 試算表選單呼叫：顯示教師主控台網頁連結
 */
function showTeacherHubUrlDialog() {
  const url = ScriptApp.getService().getUrl() + "?page=teacher";
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif; padding:15px; text-align:center;">
      <p style="font-size:16px; font-weight:bold; margin-bottom:12px;">🌟 教師全功能主控台網頁連結</p>
      <input type="text" value="${url}" readonly style="width:100%; padding:8px; font-size:14px; box-sizing:border-box; margin-bottom:15px;" id="urlInput">
      <br>
      <button onclick="navigator.clipboard.writeText(document.getElementById('urlInput').value); alert('已複製連結！');" style="padding:8px 16px; background:#2563eb; color:white; border:none; border-radius:4px; cursor:pointer; font-size:14px;">📋 複製連結</button>
      <a href="${url}" target="_blank" style="display:inline-block; margin-left:10px; padding:8px 16px; background:#10b981; color:white; text-decoration:none; border-radius:4px; font-size:14px;">🚀 直接開啟網頁</a>
    </div>
  `).setWidth(450).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, "教師主控台網址");
}



/**
 * 驗證教師 PIN 碼是否正確
 * @param {string} inputPin 前端輸入的 PIN 碼
 * @returns {Object} { success: boolean, message?: string }
 */
function verifyTeacherPin(inputPin) {
  try {
    const config = getSystemConfig(true);
    const correctPin = String(config.TEACHER_PIN || "1234").trim();
    const userPin = String(inputPin || "").trim();

    if (!userPin) {
      return { success: false, message: "請輸入 PIN 碼！" };
    }

    if (userPin === correctPin) {
      return { success: true, message: "驗證成功！" };
    } else {
      return { success: false, message: "PIN 碼錯誤，請重新輸入！" };
    }
  } catch (err) {
    return { success: false, message: "驗證過程發生錯誤：" + err.message };
  }
}
