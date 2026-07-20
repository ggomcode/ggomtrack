/**
 * Attendance Management System - Main Backend
 * Author: Antigravity AI
 */

/**
 * Handle Web App Access
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('AI 출결 관리 시스템')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Include separate HTML/JS files in index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Initialize Sheets if they don't exist
 */
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
    if (!sheet) return "AttendanceDB 시트를 찾을 수 없습니다.";
    const data = sheet.getDataRange().getValues().slice(1);
    
    const debugRows = [];
    data.forEach((row, i) => {
      const dateVal = row[0];
      const date = parseDateSafe(dateVal);
      if (!date) return;
      const month = getMonthSafe(date);
      if (month === 5) {
        debugRows.push({
          rowNum: i + 2,
          dateRaw: String(row[0]),
          dateParsed: Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd"),
          studentIdRaw: String(row[1]),
          name: String(row[2]),
          category: String(row[4]),
          subCategory: String(row[5]),
          days: String(row[8]),
          ruleCheck: String(row[19])
        });
      }
    });
    return "5월 데이터 분석 결과:\n" + JSON.stringify(debugRows, null, 2);
  } catch (e) {
    return "에러 발생: " + e.toString();
  }
}

/**
 * Initialize Log Sheet
 */
function initializeLogSheet(ss) {
  const sheetName = CONFIG.SHEETS.LOG || 'SystemLog';
  let logSheet = ss.getSheetByName(sheetName);
  if (!logSheet) {
    logSheet = ss.insertSheet(sheetName);
    const headers = ['일시', '모델', '파일명', '상태', '상세내용'];
    logSheet.appendRow(headers);
    logSheet.getRange(1, 1, 1, headers.length).setBackground('#999999').setFontColor('white').setFontWeight('bold');
  }
  return logSheet;
}

/**
 * Log to SystemLog sheet
 */
function logToSystem(filename, status, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG) || initializeLogSheet(ss);
    logSheet.appendRow([new Date(), CONFIG.GEMINI_MODEL, filename, status, message]);
  } catch (e) {
    Logger.log("로그 기록 실패: " + e.toString());
  }
}

/**
 * Upload Image/PDF to Drive and Process with Gemini
 */
function processAttendanceFile(filename, base64Data) {
  try {
    // 1. Save to Drive
    const folder = getTargetFolder();
    const contentType = base64Data.split(',')[0].split(':')[1].split(';')[0];
    const base64Content = base64Data.split(',')[1];
    const bytes = Utilities.base64Decode(base64Content);
    const fileBlob = Utilities.newBlob(bytes, contentType, filename);
    const file = folder.createFile(fileBlob);
    
    // 2. Call Gemini Vision API
    let studentNames = [];
    try {
      studentNames = getTargetGradeStudentNames();
    } catch (e) {
      Logger.log("대상 학생 명단 가져오기 실패: " + e.toString());
    }
    const extractedDataArray = callGeminiVision(base64Content, contentType, filename, studentNames);
    
    // 3. Post-processing: Correct Names & Link File
    const fileUrl = file.getUrl();
    const resultWithUrl = [];
    const allSplitItems = [];
    
    extractedDataArray.forEach(rawItem => {
      const item = { ...rawItem };
      const studentInfo = verifyStudent(item.studentId, item.name);
      
      // Normalize categories, subcategories, and reason details
      const norm = normalizeCategories(item.category, item.subCategory, filename, item.reasonDetail);
      item.category = norm.category;
      item.subCategory = norm.subCategory;
      item.reasonDetail = norm.reasonDetail;
      
      // Normalize periodStart and periodEnd format (e.g. 01교시 -> 1교시)
      item.periodStart = formatPeriodSafe(item.periodStart);
      item.periodEnd = formatPeriodSafe(item.periodEnd);
      
      // Date Override: Record actual absence date (periodStart) instead of document issuance date (date)
      if (item.periodStart && item.periodStart.includes('-')) {
        item.date = item.periodStart;
      }
      
      // OCR Correction: AI frequently misinterprets '1' in '1일간' as '|', '/', 'I', 'l', '(', '0' or completely misses it
      if (item.totalDays && typeof item.totalDays === 'string') {
        item.totalDays = item.totalDays.replace(/\s+/g, ""); // Remove spaces
        if (/^[\(]?([|/Il01(]|)\)?일간[\)]?$/i.test(item.totalDays)) {
           item.totalDays = "1일간";
        } else {
           item.totalDays = item.totalDays.replace(/[\(\)]/g, "");
        }
      }
      
      // Terminology Normalization: Unify various experiential learning terms
      if (item.reasonDetail && typeof item.reasonDetail === 'string') {
        item.reasonDetail = item.reasonDetail.replace(/(학교장허가\s*|현장\s*|교외\s*)*체험\s*(학습)?/g, "학교장허가체험학습");
      }
      
      // Clean up 'None', 'null', '없음', '해당없음' strings from AI output
      const cleanNone = (str) => (typeof str === 'string' && /^(none|null|(해당\s*)?없음|첨부서류\s*누락)$/i.test(str.trim())) ? "" : str;
      item.name = cleanNone(item.name);
      item.parentName = cleanNone(item.parentName);
      item.teacherName = cleanNone(item.teacherName);
      
      // Normalize signature fields to "있음" or "없음"
      const normalizeSign = (val) => {
        if (!val) return "없음";
        const s = String(val).trim();
        if (/^(무|x|×|no|none|n|false|서명\s*안됨|안됨|미서명|서명\s*없음|없음|null)$/i.test(s)) {
          return "없음";
        }
        return "있음";
      };
      item.studentSigned = normalizeSign(item.studentSigned);
      item.parentSigned = normalizeSign(item.parentSigned);
      item.teacherSigned = normalizeSign(item.teacherSigned);
      
      item.attachments = cleanNone(item.attachments);
      
      // Apply correct ID and Name
      if (studentInfo.match && studentInfo.correctId && studentInfo.correctName) {
        item.studentId = studentInfo.correctId;
        item.name = studentInfo.correctName;
      }

      // 4. Force NEIS string standard format: YYYY-MM-DD~YYYY-MM-DD 대분류소분류(사유상세)(X일간)
      const normDate = item.date || "";
      const normCat = item.category || "";
      const normSub = item.subCategory || "";
      const normReason = cleanNone(item.reasonDetail) || "";
      
      let normDaysNum = 1;
      const normDaysRaw = item.totalDays || "1일간";
      const normDaysMatch = normDaysRaw.match(/(\d+)/);
      if (normDaysMatch) {
        normDaysNum = parseInt(normDaysMatch[1], 10);
      }
      
      let normDateRangeStr = normDate;
      if (normDaysNum > 1 && normDate) {
        let endDateStr = "";
        if (item.periodEnd && typeof item.periodEnd === 'string') {
          const endMatch = item.periodEnd.match(/(\d{4}-\d{2}-\d{2})/);
          if (endMatch) endDateStr = endMatch[1];
        }
        if (!endDateStr) {
          try {
            const parts = normDate.split('-');
            if (parts.length === 3) {
              const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
              const endDateObj = new Date(startDateObj.getTime() + (normDaysNum - 1) * 24 * 60 * 60 * 1000);
              endDateStr = Utilities.formatDate(endDateObj, "GMT+9", "yyyy-MM-dd");
            }
          } catch (e) {}
        }
        if (endDateStr && endDateStr !== normDate) {
          normDateRangeStr = `${normDate}~${endDateStr}`;
        }
      }
      
      if (normDate && normCat && normSub) {
        const normDaySuffix = normDaysNum > 1 ? `(${normDaysNum}일간)` : "";
        if (normReason) {
          item.neisString = `${normDateRangeStr} ${normCat}${normSub}(${normReason})${normDaySuffix}`;
        } else {
          item.neisString = `${normDateRangeStr} ${normCat}${normSub}${normDaySuffix}`;
        }
      } else {
        item.neisString = cleanNone(item.neisString);
      }

      // Split item if it crosses months
      const splitList = splitMultiMonthItem(item);
      const isSplit = splitList.length > 1;

      splitList.forEach(splitItem => {
        if (isSplit) {
          splitItem.preserveNeisString = true;
          splitItem.neisString = item.neisString; // Inherit the original unsplit neisString
        } else {
          splitItem.neisString = item.neisString;
        }
        splitItem.studentInfo = studentInfo; // Temporarily attach student directory info
        allSplitItems.push(splitItem);
      });
    });

    // Second Pass: Run rule check and validation across all split items in the batch
    allSplitItems.forEach(splitItem => {
      const studentInfo = splitItem.studentInfo || {};

      // Rule checks
      splitItem.ruleCheck = checkAttendanceRule(splitItem);
      
      if (splitItem.ruleCheck) {
        splitItem.ruleCheck = splitItem.ruleCheck.replace(/⚠️ 보완 필요: (학생 |학부모 )?(이름|성명|학번) 불일치/g, "")
                                       .replace(/\//g, " ⚠️ ")
                                       .replace(/  +/g, " ")
                                       .trim();
        if (splitItem.ruleCheck.startsWith("오류") || splitItem.ruleCheck.startsWith("보완")) {
          splitItem.ruleCheck = "⚠️ " + splitItem.ruleCheck;
        }
      }
      
      // Cross-Day Validation for Disease Attachments
      const isDisease = (splitItem.category && splitItem.category.includes("질병"));
      if (isDisease) {
        const cleanAttach = splitItem.attachments ? String(splitItem.attachments).replace(/\s+/g, "") : "";
        const cleanIssuer = splitItem.issuer ? String(splitItem.issuer).replace(/\s+/g, "") : "";
        const hospitalKeywords = ["진료", "진단", "처방", "통원", "입원", "퇴원", "입퇴원", "입·퇴원", "입/퇴원", "입-퇴원", "소견", "약", "영수증", "보건", "보건실", "입실", "확인증", "병원", "약국", "의원", "치과", "한의원"];
        const hasHospitalDoc = hospitalKeywords.some(kw => cleanAttach.includes(kw) || cleanIssuer.includes(kw));
        const hasSubstituteDoc = ["학부모", "담임", "확인서", "의견서", "서명"].some(kw => cleanAttach.includes(kw));
        
        if (hasHospitalDoc) {
          splitItem.ruleCheck = ""; // valid
        } else {
          // If the item itself doesn't have a hospital document, check if there is a hospital document for the SAME day
          const targetDate = splitItem.date || splitItem.periodStart;
          const sameDayHasHospitalDoc = targetDate ? checkSameDayHospitalDoc(splitItem.studentId, targetDate, allSplitItems) : false;
          
          if (sameDayHasHospitalDoc) {
            splitItem.ruleCheck = ""; // valid
          } else if (hasSubstituteDoc) {
            const prevHasHospitalDoc = targetDate ? checkPreviousDayHospitalDoc(splitItem.studentId, targetDate, allSplitItems) : false;
            if (prevHasHospitalDoc) {
              splitItem.ruleCheck = "";
            } else {
              splitItem.ruleCheck = "체크필요 [질병 첫날 또는 반복된 결석의 홀수번째 날은 병원/약국 서류 필수]";
            }
          } else {
            splitItem.ruleCheck = "체크필요 [첨부서류 누락]";
          }
        }
      }

      // 학적 변동 기준일(자퇴, 전출, 전입) 출결 체크
      if (studentInfo.match && studentInfo.status && studentInfo.statusDate) {
        const checkDate = splitItem.date || splitItem.periodStart;
        const isActive = isStudentActiveOnDate(studentInfo.status, studentInfo.statusDate, checkDate);
        if (!isActive) {
          const sDateStr = studentInfo.statusDate instanceof Date ? Utilities.formatDate(studentInfo.statusDate, "GMT+9", "yyyy-MM-dd") : String(studentInfo.statusDate);
          if (studentInfo.status.includes('자퇴') || studentInfo.status.includes('전출') || studentInfo.status.includes('퇴학')) {
            splitItem.ruleCheck = `체크필요 [${studentInfo.status} 학생 - 기준일(${sDateStr}) 이후 출결 기록]`;
          } else if (studentInfo.status.includes('전입')) {
            splitItem.ruleCheck = `체크필요 [전입 학생 - 기준일(${sDateStr}) 이전 출결 기록]`;
          }
        }
      }

      // 출결신고서 누락 검증 (단, 고등학교 발행 정식 공문은 제외)
      const hasReport = splitItem.hasReportCard === true;
      const isOfficial = splitItem.isOfficialDocument === true;
      const issuer = String(splitItem.issuer || "").trim();
      const isSchoolOfficial = isOfficial && 
                               (issuer.includes("고등학교") || issuer.includes("학교")) && 
                               !/대학|협회|연맹|학원|병원|의원/.test(issuer);
                               
      if (!hasReport && !isSchoolOfficial) {
        if (!splitItem.ruleCheck) {
          splitItem.ruleCheck = "체크필요 [출결신고서 누락]";
        } else if (!splitItem.ruleCheck.includes("출결신고서 누락")) {
          splitItem.ruleCheck = splitItem.ruleCheck + " ⚠️ 체크필요 [출결신고서 누락]";
        }
      }

      if (splitItem.ruleCheck) {
        splitItem.ruleCheck = splitItem.ruleCheck.replace(/⚠️ 보완 필요: (학생 |학부모 )?(이름|성명|학번) 불일치/g, "")
                                       .replace(/⚠️ 보완 필요 여부 검토: 일수 기재 누락 여부에 대한 재검토 추천/g, "")
                                       .replace(/⚠️ 보완 필요: 일수 기재 누락/g, "")
                                       .trim();
        if (splitItem.ruleCheck === "규정 준수" || splitItem.ruleCheck === "✅ 정상" || splitItem.ruleCheck === "체크필요" || splitItem.ruleCheck === "체크필요:") {
          splitItem.ruleCheck = "";
        }
      }

      splitItem.fileUrl = fileUrl;
      splitItem.filename = filename;
      
      delete splitItem.studentInfo; // Remove temporary property
      resultWithUrl.push(splitItem);
    });
    
    // 4. Auto-folder Organization (Move file based on first student's info if multiple)
    if (resultWithUrl.length > 0) {
      organizeFile(file, resultWithUrl[0]);
    }
    
    return {
      success: true,
      fileUrl: fileUrl,
      data: resultWithUrl
    };
  } catch (e) {
    Logger.log(e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Splits an attendance item into multiple segments if its date range crosses month boundaries.
 */
function splitMultiMonthItem(item) {
  const periodStart = item.periodStart || item.date || "";
  const periodEnd = item.periodEnd || "";
  
  let daysNum = 1;
  const daysRaw = item.totalDays || "1일간";
  const daysMatch = daysRaw.match(/(\d+)/);
  if (daysMatch) {
    daysNum = parseInt(daysMatch[1], 10);
  }
  
  if (!periodStart) return [item];
  
  const parseDateTime = (str) => {
    if (!str) return { date: "", time: "" };
    const match = str.match(/^(\d{4}-\d{2}-\d{2})(.*)$/);
    if (match) {
      return { date: match[1], time: match[2].trim() };
    }
    return { date: str, time: "" };
  };
  
  const startInfo = parseDateTime(periodStart);
  const endInfo = parseDateTime(periodEnd);
  
  if (!startInfo.date) return [item];
  
  let actualEndStr = endInfo.date;
  if (!actualEndStr) {
    try {
      const parts = startInfo.date.split('-');
      if (parts.length === 3) {
        const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
        const endDateObj = new Date(startDateObj.getTime() + (daysNum - 1) * 24 * 60 * 60 * 1000);
        actualEndStr = Utilities.formatDate(endDateObj, "GMT+9", "yyyy-MM-dd");
      }
    } catch (e) {}
  }
  
  if (!actualEndStr) return [item];
  
  try {
    const startParts = startInfo.date.split('-');
    const endParts = actualEndStr.split('-');
    if (startParts.length !== 3 || endParts.length !== 3) return [item];
    
    const startDateObj = new Date(parseInt(startParts[0], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[2], 10), 12, 0, 0);
    const endDateObj = new Date(parseInt(endParts[0], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[2], 10), 12, 0, 0);
    
    if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime()) || startDateObj.getTime() > endDateObj.getTime()) {
      return [item];
    }
    
    // Check if same month
    if (startDateObj.getFullYear() === endDateObj.getFullYear() && startDateObj.getMonth() === endDateObj.getMonth()) {
      return [item];
    }
    
    // Generate all dates
    const allDates = [];
    const diffTime = endDateObj.getTime() - startDateObj.getTime();
    const calculatedDays = Math.round(diffTime / (24 * 60 * 60 * 1000)) + 1;
    for (let d = 0; d < calculatedDays; d++) {
      const currentDateObj = new Date(startDateObj.getTime() + d * 24 * 60 * 60 * 1000);
      allDates.push(Utilities.formatDate(currentDateObj, "GMT+9", "yyyy-MM-dd"));
    }
    
    const groups = {};
    allDates.forEach(dateStr => {
      const ym = dateStr.substring(0, 7);
      if (!groups[ym]) groups[ym] = [];
      groups[ym].push(dateStr);
    });
    
    const ymKeys = Object.keys(groups).sort();
    if (ymKeys.length <= 1) return [item];
    
    const splitItems = [];
    ymKeys.forEach((ym, index) => {
      const groupDates = groups[ym];
      const segStartDate = groupDates[0];
      const segEndDate = groupDates[groupDates.length - 1];
      const segDays = groupDates.length;
      
      const newItem = JSON.parse(JSON.stringify(item));
      
      newItem.date = segStartDate;
      newItem.periodStart = segStartDate + (index === 0 && startInfo.time ? " " + startInfo.time : "");
      newItem.periodEnd = segEndDate + (index === ymKeys.length - 1 && endInfo.time ? " " + endInfo.time : "");
      newItem.totalDays = segDays + "일간";
      
      splitItems.push(newItem);
    });
    
    return splitItems;
  } catch (e) {
    return [item];
  }
}

/**
 * Helper: Get Target Folder (By ID or Spreadsheet Parent)
 */
function getTargetFolder() {
  const folderId = CONFIG.UPLOAD_FOLDER_ID;
  if (folderId && folderId.trim().length > 0) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      Logger.log("경고: 폴더 ID가 잘못되었습니다. 시트 상위 폴더를 사용합니다.");
    }
  }
  
  // Fallback: Current Spreadsheet's Parent Folder
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  if (parents.hasNext()) return parents.next();
  
  return DriveApp.getRootFolder(); // Last resort
}

/**
 * Call Gemini API with Retry and Logging
 */
function callGeminiVision(base64Content, mimeType, filename, studentNames = []) {
  const apiKey = CONFIG.GEMINI_API_KEY;
  if (!apiKey) throw new Error("GEMINI_API_KEY가 설정되지 않았습니다. 설정(Settings) 페이지에서 API 키를 입력해 주세요.");

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;
  
  let prompt = `
    신고서 및 첨부서류를 분석하여 정보를 추출하고 규정 준수 여부를 검토하라.
    모든 응답은 반드시 지정된 JSON 구조의 배열(ARRAY)로 반환되어야 한다.

    [추출 가이드]
    - **category (대분류):** 반드시 '질병', '출석인정', '기타', '미인정' 중 정확히 하나만 선택하여 기입하라. 
      - '출석인정': 생리통(생리결석/생리조퇴/생리지각), 학교장허가체험학습, 전염병/감염병(독감, 인플루엔자, 코로나19 등 격리가 필요한 병명), 경조사, 공가 등의 사유일 때 선택하라.
      - '질병': 감기, 몸살, 장염, 두통, 복통, 치과 치료 등 일반적인 질병이나 통원 치료일 때 선택하라.
      - '미인정': 태만, 가출, 고의적 출석 거부 등 정당한 사유가 없을 때 선택하라.
      - 절대 '질병결석'이나 '출석인정(생리통)' 처럼 소분류나 상세사유를 결합해 기입하지 마라.
    - **subCategory (소분류):** 반드시 '결석', '지각', '조퇴', '결과' 중 정확히 하나만 선택하여 기입하라. 절대 '생리통', '경조사', '깸석' 등을 기입하지 마라. 생리통이나 경조사 같은 상세 사유는 반드시 사유상세(reasonDetail) 필드에만 기입해야 한다. (주의: 문서 제목이 '결석계'이더라도 본문 내부의 선택 항목(체크박스 등)에서 '지각', '조퇴', '결과' 등에 체크되어 있거나 사유상세에 지각/조퇴가 명시되어 있다면 해당 소분류로 정확하게 분류하여 추출하라. 결석이 아닌 건을 결석으로 오분류하지 않도록 정밀히 판별하라.)
    - **서명(학생/학부모/담임):** 각 서명란(주로 이름 옆이나 '(서명 또는 인)', '(인)', '(서명)' 표시가 있는 곳)을 정밀하게 확인하여 물리적인 서명(수필 서명, 사인, 흘려 쓴 이름, 도장/날인 등)이 존재하면 "있음", 서명란이 비어 있거나 인쇄된 성명 텍스트 외에 친필 서명/사인이 없는 경우 "없음"으로 매우 엄격하게 판별하여 추출하라. 학생서명, 학부모서명, 담임서명 모두 동일하게 판별한다.
    - **학번 추출 규칙:** 5자리 학번의 첫 번째 숫자는 항상 1~6 사이이다. 절대 0으로 시작할 수 없다. 만약 0으로 보인다면 보정하여 추출하라.
    - **이름/학번/학부모명 불일치 판정 금지:** 이미지 판독 결과가 명부와 다르거나 판독이 어렵더라도 성함/학번 불일치와 관련된 보완 필요 경고를 절대 생성하지 마라.
    - **기간 (시작일시/종료일시):** 결석계 폼 내부에 명시된 '결석/지각/조퇴/결과 시작일'과 '종료일'을 찾아 "YYYY-MM-DD X교시" 형식으로 추출하라. (작성일이나 첨부서류 발급일이 아님). 기재되어 있지 않으면 "×"로 추출하라.
    - **일수:** 숫자 1을 괄호와 함께 흘려 쓰거나('|', 'I', 'l', '/', '(' 등), 아예 비워두는 경우가 많습니다. '0일간'이거나 형태가 불분명하면 무조건 "1일간"으로 추출하고, 그 외 숫자가 명확하면 "X일간"으로 추출하라. 괄호()는 빼고 숫자와 글자만 추출하라.
    - **사유상세:** 대분류가 질병이나 기타인 경우 학생이 기록한 구체적 사유를, 출석인정인 경우 생리통, 학교장허가체험학습, 전염병 또는 구체적 사유를 그대로 추출하라.
    - **첨부서류:** PDF 전체를 정밀하게 분석하여 실제로 첨부된 서류의 정확한 명칭(예: 진료확인서, 처방전 등)을 하나만 추출하고, 없으면 "×"로 추출하라. (안내 문구 목록을 적지 마라).
    - **통원일(docStartDate)/종료일(docEndDate):** 병원/약국 발행 서류가 기준이다. 서류에 명시된 '통원 기간' 또는 '진료 기간'의 시작일이 발급일과 다르다면 그 시작일을 통원일로, 종료일이 있다면 통원종료일로 추출하라. 기간 표시 없이 발급일만 있다면 발급일을 통원일로 추출하고 통원종료일은 빈칸("")으로 추출하라.
    - **발급처 (issuer):** 서류를 발급한 기관명(병원, 약국 등)을 추출하라. 단, 학부모 확인서는 "학부모", 담임 확인서는 "담임"이라고 추출하라.
    - **규정위반여부 (ruleCheck):** 서류가 완벽히 규정을 준수하면 빈칸("")으로 반환하라. 위반사항이 있다면 "체크필요: [짧은 사유]" 형식으로 반환하라. 단, 질병 결석의 경우 병원/약국 서류 유무에 대한 판정만 진행하라. 서류 제출 지연(제출 기한 도과)에 대한 판정은 절대 수행하지 마라.
    - **NEIS 문구 (neisString):** 반드시 "[일자] [대분류][소분류]([사유상세])" 형태로 생성하라.
      - 예시: 2026-03-10 질병결석(복통)
      - 만약 사유상세(reasonDetail)가 없거나 '없음', '해당없음'인 경우 괄호와 사유를 생략하고 "[일자] [대분류][소분류]"로만 생성하라. (예: 2026-03-10 질병결석)
      - 일자는 date 필드값(YYYY-MM-DD)을 활용하며, 대분류와 소분류는 띄어쓰기 없이 붙여서 기입하라.
    - **수정 및 정정 표시 처리:**
      - 문서 본문에 두 줄을 긋고 수정도장(보통 빨간색/파란색 원형 또는 사각형 개인 도장)을 찍거나 글씨를 덧써서 수정(정정)한 경우, 두 줄로 지워진 과거의 텍스트나 항목은 완전히 무시하고 **최종 수정(정정)된 올바른 텍스트 및 표시만을 기준으로 정보를 추출**하라.
      - **도장 겹침 오판 주의:** 수정도장이 체크박스(예: '지각', '조퇴' 등)나 특정 글자 위에 찍혀 있는 경우, 도장의 붉은색/푸른색 잉크나 테두리 선이 진하게 표시되어 마치 해당 체크박스에 체크 표시(V)가 된 것처럼 AI가 오인하기 쉽다. 수정도장의 잉크 자국이나 날인 형태는 체크 표시가 아님을 인지하고, 도장 겹침으로 인해 진하게 보이는 부분에 현혹되지 마라. 실제 손글씨로 직접 체크(V)되거나 동그라미 쳐진 최종 선택 항목(비록 더 연하게 표시되어 있더라도)을 신중하고 정확하게 판별하여 추출하라.
    - **hasReportCard (출결신고서 포함 여부):** PDF 전체(특히 해당 학생의 서류 영역)에 결석계, 결석신고서, 지각·조퇴·결과 신고서 등의 출결신고서 양식이 포함되어 있는 경우 true, 포함되어 있지 않고 증빙서류(진료확인서, 처방전 등)만 있는 경우 false로 판별하여 반환하라.
    - **isOfficialDocument (정식 공문 여부):** PDF 내에 공무 수행, 대회 참가 등의 증빙을 위해 학교나 기관 등에서 발행한 '공문'(수신/발신처 및 관인/직인이 날인된 공식 문서 포맷)이 포함되어 있는 경우 true, 그렇지 않으면 false로 판별하여 반환하라.

    [복수 학생 서류 병합 파일 처리 지침]
    - **대상 상황:** 하나의 PDF 파일 내에 여러 학생의 서류가 합쳐져서 업로드되는 경우 이 지침을 적용한다.
    - **학생별 독립 인지:** 문서 내의 모든 페이지를 정밀 분석하여, 각 페이지마다 다르게 나타나는 학번과 학생 이름을 가장 먼저 파악하라.
    - **개별 레코드 분리:** 학번이나 이름이 서로 다르다면 완전히 별개의 학생 결석/출결 건으로 간주하여, 각각 독립된 JSON 객체(레코드)로 분리하고 배열에 포함하라.
    - **데이터 격리:** 앞 페이지에 나온 특정 학생의 신고서 정보(시작일시, 종료일시, 구체적 사유 등)를 뒤 페이지에 등장하는 다른 학생의 증빙 서류에 결합하거나 적용(상속)하여 오판하지 않도록 철저히 격리하라.
    - **신고서 없는 학생 처리:** 뒤 페이지 학생은 신고서 양식이 없으므로, 해당 학생의 JSON 객체에는 날짜/일수 등을 증빙서류 기준으로만 추출하고, 결석계 양식 관련 필드('studentSigned', 'parentName', 'parentSigned' 등)는 "없음"으로 처리하고, 'hasReportCard'는 반드시 'false'로 설정하라.

    [공문 및 목록/표 형태의 서류 처리 지침]
    - **대상 문서:** 단일 결석계 양식이 아니라, 공문(공식 문서), 참가자 명단, 월간 근무상황표, 출석표, 활동일지 등 목록이나 표 형태의 서류인 경우 이 지침을 적용한다.
    - **개별 레코드 분리:** 대상 학생 명단에 포함된 학생의 이름을 본 문서(예: 표의 참여자/근무자/참가자 열 등)에서 찾은 뒤, 해당 학생이 참여/활동한 **각각의 날짜(일자)별로 개별적인 JSON 레코드(객체)**를 생성하여 배열에 포함하라.
      - 예: 학생 '장명준'이 3월 5일, 3월 10일, 3월 12일에 참여한 기록이 표에 있다면, '장명준'에 대해 날짜가 각각 2026-03-05, 2026-03-10, 2026-03-12인 총 3개의 JSON 객체를 각각 생성하여 배열에 담아야 한다.
    - **날짜 필드 지정:** 'date'와 'periodStart', 'periodEnd' 필드는 모두 해당 활동이 발생한 날짜(형식: "YYYY-MM-DD")로 설정하라. 연도가 명시되지 않았다면 문서의 발급일/작성일의 연도(예: 2026년)를 기준으로 채워라.
    - **출결 구분 기본값:** 목록/표에서 해당 활동 참여가 지각, 조퇴, 결과, 결석 중 어느 것인지 명시되어 있지 않거나 단순히 '참여', '근무', '출근' 등으로 되어 있어 모호한 경우, **대분류(category)는 "출석인정", 소분류(subCategory)는 "결석"으로 기본 설정**하라. (단, 문서에 지각, 조퇴, 결과가 명확히 표시된 경우는 해당 값을 사용하라.)
    - **서명 및 증빙 기본값:** 표/목록 형태 서류의 특성상 개별 학생/학부모 친필 서명은 없으므로, 불필요한 서류 미비 경고가 발생하지 않도록 **'studentSigned', 'parentSigned', 'teacherSigned' 필드는 모두 "있음"**으로 설정하라.
    - **사유 상세 및 첨부서류:** 'reasonDetail'에는 해당 행의 구체적인 활동 내용 또는 업무내용(예: "장애인 맞춤형 일자리 오리엔테이션", "커피바리스타 1강" 등)을 추출하여 기입하라. 'attachments' 필드는 "공문" 또는 해당 서류의 명칭(예: "근무상황표")을 적어라. 'totalDays'는 "1일간"으로 설정하고, 'issuer'는 문서의 확인자/담당자 성함(예: "지예나") 또는 발급 기관명을 적어라.
  `;

  if (studentNames && studentNames.length > 0) {
    prompt += `
    \n[대상 학생 명단]
    현재 분석 대상 학년의 학생 명단은 다음과 같다: ${studentNames.join(', ')}
    문서 내에서 참여자 또는 대상자 목록을 확인할 때, 반드시 이 명단에 포함된 학생의 정보만 추출하라. 명단에 없는 이름(예: 확인자/검토자 '지예나' 등)은 절대로 학생 이름으로 추출해선 안 된다.
    `;
  }

  const payload = {
    contents: [{
      parts: [
        { inlineData: { mimeType: mimeType, data: base64Content } },
        { text: String(prompt).trim() }
      ]
    }],
    generationConfig: {
      response_mime_type: "application/json",
      response_schema: {
        type: "ARRAY",
        items: {
          type: "OBJECT",
          properties: {
            date: { type: "STRING" },
            studentId: { type: "STRING" },
            name: { type: "STRING" },
            studentSigned: { type: "STRING" },
            category: { type: "STRING" },
            subCategory: { type: "STRING" },
            periodStart: { type: "STRING" },
            periodEnd: { type: "STRING" },
            totalDays: { type: "STRING" },
            reasonDetail: { type: "STRING" },
            parentName: { type: "STRING" },
            parentSigned: { type: "STRING" },
            teacherName: { type: "STRING" },
            teacherSigned: { type: "STRING" },
            attachments: { type: "STRING" },
            docStartDate: { type: "STRING" },
            docEndDate: { type: "STRING" },
            issuer: { type: "STRING" },
            ruleCheck: { type: "STRING" },
            neisString: { type: "STRING" },
            hasReportCard: { type: "BOOLEAN" },
            isOfficialDocument: { type: "BOOLEAN" }
          },
          required: ["date", "studentId", "name", "studentSigned", "category", "subCategory", "periodStart", "periodEnd", "totalDays", "reasonDetail", "parentName", "parentSigned", "teacherName", "teacherSigned", "attachments", "docStartDate", "docEndDate", "issuer", "ruleCheck", "neisString", "hasReportCard", "isOfficialDocument"]
        }
      }
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = fetchWithRetry(url, options, filename);
  const responseCode = response.getResponseCode();
  const contentText = response.getContentText();
  
  if (responseCode !== 200) {
    let errorMsg = contentText;
    try {
      const errorObj = JSON.parse(contentText);
      errorMsg = errorObj.error ? errorObj.error.message : contentText;
    } catch (e) {}
    logToSystem(filename, "ERROR " + responseCode, errorMsg);
    throw new Error(errorMsg);
  }

  const json = JSON.parse(contentText);
  let resultText = json.candidates[0].content.parts[0].text;
  
  logToSystem(filename, "SUCCESS", "분석 완료");

  try {
    // Schema enforcement usually returns clean JSON, but just in case:
    const jsonMatch = resultText.match(/\[[\s\S]*\]/);
    const finalJson = jsonMatch ? jsonMatch[0] : (resultText.startsWith('{') ? `[${resultText}]` : resultText);
    return JSON.parse(finalJson);
  } catch (e) {
    throw new Error("AI 데이터 파싱 실패: " + e.message + "\nRaw Content: " + resultText);
  }
}

/**
 * Fetch with Exponential Backoff Retry
 */
function fetchWithRetry(url, options, filename, maxRetries = 3) {
  let retryCount = 0;
  let waitTime = 2000; // Start with 2 seconds

  while (retryCount <= maxRetries) {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    // If success or non-retryable error (not 429 or 5xx), return
    if (code === 200 || (code !== 429 && code < 500)) {
      return response;
    }

    // If 429 or 5xx, retry with delay
    if (retryCount < maxRetries) {
      Logger.log(`Retry ${retryCount + 1}/${maxRetries} for ${filename} after ${waitTime}ms (Status: ${code})`);
      logToSystem(filename, "RETRY " + code, `${retryCount + 1}차 재시도 대기중... (${waitTime}ms)`);
      
      Utilities.sleep(waitTime);
      retryCount++;
      waitTime *= 2; // Exponential backoff
    } else {
      return response; // Final failure
    }
  }
}

/**
 * Process NEIS Data (Array of Arrays)
 */
function processNeisData(month, data, filename) {
  try {
    // data is already a 2D array parsed by frontend SheetJS
    let globalGrade = "";
    let globalClass = "";
    
    for (let i = 0; i < Math.min(15, data.length); i++) {
       const rowStr = (data[i] || []).join(" ");
       const match = rowStr.match(/(\d+)학년\s*(\d+)반/);
       if (match) {
         globalGrade = match[1];
         globalClass = match[2].padStart(2, '0');
         break;
       }
    }

    const settings = getSystemSettings();
    if (globalGrade && String(globalGrade) !== String(settings.grade)) {
      throw new Error(`이 시스템은 ${settings.grade}학년만 관리하도록 설정되어 있습니다. 업로드한 파일은 ${globalGrade}학년 데이터입니다.`);
    }

    const students = [];
    let isDataRowStarted = false;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // Skip empty or short rows
      if (!row || row.length < 5) continue;
      
      const seqRaw = String(row[0]).trim();
      
      // Look for a row that starts with a number (SEQ) to identify data rows
      if (/^\d+$/.test(seqRaw)) {
        isDataRowStarted = true;
        const name = String(row[1]).trim();
        if (!name) continue; 
        
        const sNum = seqRaw.padStart(2, '0');
        let stId = "";
        if (globalGrade && globalClass) {
          stId = globalGrade + globalClass + sNum;
        } else {
          stId = seqRaw; // fallback
        }
        
        const parseStat = (val) => {
          const parsed = parseInt(val, 10);
          return isNaN(parsed) ? 0 : parsed;
        };

        students.push({
          seq: seqRaw,
          studentId: stId,
          name: name,
          absent: [parseStat(row[3]), parseStat(row[4]), parseStat(row[5]), parseStat(row[6])],
          tardy: [parseStat(row[7]), parseStat(row[8]), parseStat(row[9]), parseStat(row[10])],
          early: [parseStat(row[11]), parseStat(row[12]), parseStat(row[13]), parseStat(row[14])],
          skipped: [parseStat(row[15]), parseStat(row[16]), parseStat(row[17]), parseStat(row[18])],
          stats: [parseStat(row[19]), parseStat(row[20]), parseStat(row[21]), parseStat(row[22])]
        });
      }
    }

    // Sort students by studentId in ascending order
    students.sort((a, b) => {
      const aId = parseInt(a.studentId, 10);
      const bId = parseInt(b.studentId, 10);
      if (!isNaN(aId) && !isNaN(bId)) {
        return aId - bId;
      }
      return a.studentId.localeCompare(b.studentId);
    });

    if (students.length === 0) {
      logToSystem(filename, "EMPTY_DATA", "조건에 부합하는 학생(출결기록 발생)이 없거나 파싱할 데이터가 없습니다.");
      return { success: true, month: month, count: 0 };
    }

    saveNeisStatsSheet(month, students);
    logToSystem(filename, "SUCCESS", `${month}월 데이터 ${students.length}건 저장 완료`);
    
    // Backup is handled at frontend or ignored because data is just extracted

    return { success: true, month: month, count: students.length };
  } catch(e) {
    Logger.log(e.toString());
    logToSystem(filename, "ERROR", e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Save Parsed NEIS Statistics to NeisData sheet
 */
function saveNeisStatsSheet(month, students) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  if (!sheet) throw new Error("NEIS 시트가 존재하지 않습니다.");
  
  const safeGet = (arr, idx) => (arr && Array.isArray(arr) && arr.length > idx) ? arr[idx] : 0;
  
  const rows = students.map(st => [
    st.seq || '',
    month || '',
    st.studentId || '',
    st.name || '',
    safeGet(st.absent, 0), safeGet(st.absent, 1), safeGet(st.absent, 2), safeGet(st.absent, 3),
    safeGet(st.tardy, 0), safeGet(st.tardy, 1), safeGet(st.tardy, 2), safeGet(st.tardy, 3),
    safeGet(st.early, 0), safeGet(st.early, 1), safeGet(st.early, 2), safeGet(st.early, 3),
    safeGet(st.skipped, 0), safeGet(st.skipped, 1), safeGet(st.skipped, 2), safeGet(st.skipped, 3),
    safeGet(st.stats, 0), safeGet(st.stats, 1), safeGet(st.stats, 2), safeGet(st.stats, 3)
  ]);

  if (rows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    const targetRange = sheet.getRange(startRow, 1, rows.length, 24);
    targetRange.setValues(rows);
    
    // Create background color matrix to highlight non-zero attendance records
    const bgColors = [];
    rows.forEach(row => {
      const bgRow = Array(24).fill(null);
      let hasNonZero = false;
      for (let c = 4; c < 24; c++) {
        if (typeof row[c] === 'number' && row[c] > 0) {
          hasNonZero = true;
        }
      }
      
      if (hasNonZero) {
        // Highlight columns A~D (0~3) in soft yellow
        for (let c = 0; c < 4; c++) {
          bgRow[c] = '#fff2cc';
        }
        // Highlight only cells with non-zero values in soft yellow
        for (let c = 4; c < 24; c++) {
          if (typeof row[c] === 'number' && row[c] > 0) {
            bgRow[c] = '#fff2cc';
          }
        }
      }
      bgColors.push(bgRow);
    });
    
    targetRange.setBackgrounds(bgColors);
    
    // [New] Automatically reconcile and update Y, Z columns after saving
    reconcileNeisWithAttendance(month, "all");
  }
}

/**
 * Reconcile NeisData sheet with AttendanceDB records and write results to AttendanceData
 */
function reconcileNeisWithAttendance(targetMonth, targetClass) {
  targetMonth = targetMonth || "all";
  targetClass = targetClass || "all";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  let dataSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
  
  if (!attSheet || !neisSheet) return "필수 시트가 존재하지 않습니다.";
  
  if (!dataSheet) {
    dataSheet = ss.insertSheet(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
    const dataHeaders = ['순번', '월', '학번', '이름', '결석_질병', '결석_미인정', '결석_기타', '결석_인정', '지각_질병', '지각_미인정', '지각_기타', '지각_인정', '조퇴_질병', '조퇴_미인정', '조퇴_기타', '조퇴_인정', '결과_질병', '결과_미인정', '결과_기타', '결과_인정', '결석통계', '지각통계', '조퇴통계', '결과통계', '일치여부', '불일치내용'];
    dataSheet.appendRow(dataHeaders);
    dataSheet.getRange(1, 1, 1, dataHeaders.length).setBackground('#e67e22').setFontColor('white').setFontWeight('bold');
  }

  const settings = getSystemSettings();

  const attData = attSheet.getDataRange().getValues().slice(1);
  const neisRows = neisSheet.getDataRange().getValues().slice(1);
  
  // 1. Aggregate AttendanceDB by StudentId_Month
  // Map<SID_Month, { stats: { "Category_SubCategory": Value }, processedFiles: Set }>
  const aggMap = new Map();
  
  attData.forEach((row, rowIndex) => {
    // Determine the start date of the period from G열 (시작일시) first, with fallback to A열 (일자)
    const dateValRaw = String(row[6] || "").trim();
    let dateVal = "";
    if (dateValRaw && dateValRaw.includes('-')) {
      dateVal = dateValRaw.split(' ')[0];
    }
    if (!dateVal || dateVal.split('-').length !== 3) {
      dateVal = row[0];
    }
    
    const date = parseDateSafe(dateVal);
    if (!date || isNaN(date.getTime())) return;
    
    const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
    const sid = String(row[1]).trim();
    const gradeNum = sid.length >= 3 ? parseInt(sid.charAt(0), 10) : -1;
    const cNum = sid.length >= 3 ? parseInt(sid.substring(1,3), 10) : -1;
    if (gradeNum !== settings.grade || cNum > settings.classes) return;
    const catRaw = String(row[5] || "").trim();      // F열: 소분류 (결석, 지각, 조퇴, 결과) -> NEIS category
    const subCatRaw = String(row[4] || "").trim();   // E열: 대분류 (질병, 출석인정, 기타, 미인정) -> NEIS subcategory
    const daysRaw = String(row[8] || "1").trim();
    
    // Normalize subCategory labels to match NEIS (질병, 미인정, 기타, 인정)
    let subCat = "기타";
    if (subCatRaw.includes("질병")) subCat = "질병";
    else if (subCatRaw.includes("미인정")) subCat = "미인정";
    else if (subCatRaw.includes("기타")) subCat = "기타";
    else if (subCatRaw.includes("출석인정") || subCatRaw.includes("인정")) subCat = "인정";

    let cat = "결석";
    if (catRaw.includes("결석")) cat = "결석";
    else if (catRaw.includes("지각")) cat = "지각";
    else if (catRaw.includes("조퇴")) cat = "조퇴";
    else if (catRaw.includes("결과")) cat = "결과";

    // Parse days: "3일간" -> 3
    let days = 1;
    const match = daysRaw.match(/(\d+)/);
    if (match) days = parseInt(match[1], 10);
    
    const reasonDetail = String(row[9] || "").trim();
    const isExperiential = reasonDetail.includes("학교장허가체험학습") || 
                           reasonDetail.includes("현장체험") || 
                           reasonDetail.includes("교외체험") || 
                           reasonDetail.includes("체험학습");

    // Calculate covered dates range
    const coveredDates = [];
    try {
      const parts = dateStr.split('-');
      if (parts.length === 3) {
        const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
        for (let d = 0; d < days; d++) {
          const currentDateObj = new Date(startDateObj.getTime() + d * 24 * 60 * 60 * 1000);
          coveredDates.push(Utilities.formatDate(currentDateObj, "GMT+9", "yyyy-MM-dd"));
        }
      } else {
        coveredDates.push(dateStr);
      }
    } catch (e) {
      coveredDates.push(dateStr);
    }

    const statKey = `${cat}_${subCat}`;

    // Distribute each covered date to its respective calendar month
    coveredDates.forEach(dStr => {
      const parts = dStr.split('-');
      if (parts.length !== 3) return;
      const actualMonth = parseInt(parts[1], 10);
      
      const key = `${sid}_${actualMonth}`;
      if (!aggMap.has(key)) {
        aggMap.set(key, {
          stats: {},
          processedDates: {},
          experientialRowNums: new Set()
        });
      }
      const studentData = aggMap.get(key);
      const studentStats = studentData.stats;

      // 생리통 예외: 생리통(출석인정)은 서류 제출 의무가 없으므로 로컬 서류 건수 계산에서 제외
      const isMenstrual = subCat === "인정" && /생리/.test(reasonDetail);

      // Handle experiential learning document count per month (prevent counting multiple days in the same document as separate docs)
      if (subCat === "인정" && isExperiential) {
        if (!studentData.experientialRowNums.has(rowIndex)) {
          studentData.experientialRowNums.add(rowIndex);
          studentStats["체험학습_건수"] = (studentStats["체험학습_건수"] || 0) + 1;
        }
      }

      if (cat === "결석") {
        if (!studentData.processedDates[statKey]) {
          studentData.processedDates[statKey] = new Set();
        }
        if (!studentData.processedDates[statKey].has(dStr)) {
          studentData.processedDates[statKey].add(dStr);
          if (isMenstrual) {
            studentStats[statKey + "_menstrual"] = (studentStats[statKey + "_menstrual"] || 0) + 1;
          } else {
            studentStats[statKey] = (studentStats[statKey] || 0) + 1;
          }
        }
      } else {
        // For 지각/조퇴/결과, count unique days
        if (isMenstrual) {
          const mKey = statKey + "_menstrual";
          if (!studentStats[mKey]) studentStats[mKey] = new Set();
          if (studentStats[mKey] instanceof Set) {
            studentStats[mKey].add(dStr);
          }
        } else {
          if (!studentStats[statKey]) studentStats[statKey] = new Set();
          if (studentStats[statKey] instanceof Set) {
            studentStats[statKey].add(dStr);
          }
        }
      }

      // Track dates for each category in the actual month (For visual display in NEIS mismatch details)
      if (!isMenstrual) {
        const dateKey = `${statKey}_dates`;
        if (!studentStats[dateKey]) studentStats[dateKey] = [];
        if (!studentStats[dateKey].includes(dStr)) {
          studentStats[dateKey].push(dStr);
        }
      }
    });
  });

  // 1.5. Read existing AttendanceData to prepare for non-destructive overwriting
  const lastRowData = dataSheet.getLastRow();
  let existingData = [];
  let existingBg = [];
  if (lastRowData > 1) {
    existingData = dataSheet.getRange(2, 1, lastRowData - 1, 26).getValues();
    existingBg = dataSheet.getRange(2, 1, lastRowData - 1, 26).getBackgrounds();
  }
  
  // Create a map of existing rows by key "월_학번_이름" -> { index, values, bg }
  const existingMap = new Map();
  existingData.forEach((row, i) => {
    const month = String(row[1] || "").trim();
    const sid = String(row[2] || "").trim();
    const name = String(row[3] || "").trim();
    if (month && sid && name) {
      const key = `${month}_${sid}_${name}`;
      existingMap.set(key, { index: i, values: row, bg: existingBg[i] });
    }
  });

  const studentMap = getStudentDirectoryMap();
  const examPeriods = getExamPeriods();

  // 2. Iterate NeisData and Compare -> Build AttendanceData
  const mapping = [
    { col: 4, cat: "결석", sub: "질병" }, { col: 5, cat: "결석", sub: "미인정" }, { col: 6, cat: "결석", sub: "기타" }, { col: 7, cat: "결석", sub: "인정" },
    { col: 8, cat: "지각", sub: "질병" }, { col: 9, cat: "지각", sub: "미인정" }, { col: 10, cat: "지각", sub: "기타" }, { col: 11, cat: "지각", sub: "인정" },
    { col: 12, cat: "조퇴", sub: "질병" }, { col: 13, cat: "조퇴", sub: "미인정" }, { col: 14, cat: "조퇴", sub: "기타" }, { col: 15, cat: "조퇴", sub: "인정" },
    { col: 16, cat: "결과", sub: "질병" }, { col: 17, cat: "결과", sub: "미인정" }, { col: 18, cat: "결과", sub: "기타" }, { col: 19, cat: "결과", sub: "인정" }
  ];

  const newRowsToAppend = [];

  neisRows.forEach((row) => {
    if (!row[1] || !row[2]) return;
    const rowMonth = parseInt(row[1]);
    const sid = String(row[2]).trim();
    
    const gradeNum = sid.length >= 3 ? parseInt(sid.charAt(0), 10) : -1;
    // Parse class number from studentId (studentId format: G-CC-NN, e.g., 30101 -> class 1)
    const cNum = sid.length >= 3 ? parseInt(sid.substring(1,3), 10) : -1;
    
    if (gradeNum !== settings.grade || cNum > settings.classes) return;
    
    // Apply Month & Class Filter
    const matchesMonth = (targetMonth === "all" || String(rowMonth) === String(targetMonth));
    const matchesClass = (targetClass === "all" || String(cNum) === String(targetClass));
    if (!matchesMonth || !matchesClass) return;

    const name = String(row[3]).trim();
    const key = `${sid}_${rowMonth}`;
    const studentData = aggMap.get(key) || { stats: {} };
    const studentStats = studentData.stats;
    
    const studentInfo = studentMap.get(sid);
    const isFemale = studentInfo && (String(studentInfo.gender).trim() === '여' || String(studentInfo.gender).trim() === '여자');
    
    // Check if there are any exam periods falling in this rowMonth
    const hasExamInMonth = examPeriods.some(period => {
      if (!period.startStr) return false;
      const startMonth = parseInt(period.startStr.split('-')[1], 10);
      const endMonth = period.endStr ? parseInt(period.endStr.split('-')[1], 10) : startMonth;
      return startMonth === rowMonth || endMonth === rowMonth;
    });

    let mismatches = [];
    let isMatch = true;
    
    // 순번, 월, 학번, 성명
    const newRow = [row[0], row[1], row[2], row[3]];
    const bgRow = [null, null, null, null];
    
    let absTotal = 0;
    let tardyTotal = 0;
    let earlyTotal = 0;
    let skipTotal = 0;

    mapping.forEach(m => {
      const neisVal = parseInt(row[m.col]) || 0;
      const statKey = `${m.cat}_${m.sub}`;
      const localValRaw = studentStats[statKey];
      let localVal = 0;
      
      if (m.cat === "결석") {
        localVal = (typeof localValRaw === 'number') ? localValRaw : 0;
        if (m.sub !== "인정") absTotal += localVal;
      } else if (m.cat === "지각") {
        localVal = (localValRaw instanceof Set) ? localValRaw.size : 0;
        if (m.sub !== "인정") tardyTotal += localVal;
      } else if (m.cat === "조퇴") {
        localVal = (localValRaw instanceof Set) ? localValRaw.size : 0;
        if (m.sub !== "인정") earlyTotal += localVal;
      } else if (m.cat === "결과") {
        localVal = (localValRaw instanceof Set) ? localValRaw.size : 0;
        if (m.sub !== "인정") skipTotal += localVal;
      }
      
      newRow.push(localVal);
      
      let isMismatch = neisVal !== localVal;
      
      // Unexcused (미인정) exception:
      // Unexcused absences/tardies/earlies/skips do not have scanned excuse forms.
      // Therefore, do not flag a mismatch for "미인정" category.
      if (isMismatch && m.sub === "미인정") {
        isMismatch = false;
      }
      
      // Menstrual pain (생리통) exception:
      // If female, subCategory is "인정":
      // 1) If NEIS count is greater than local count, the difference (neisVal - localVal) is assumed to be undocumented menstrual pain absences (since no document is required outside exam periods).
      //    We suppress the mismatch if the total menstrual count (local menstrual docs + undocumented diff) is within the allowed monthly limit (1 for 결석, 3 for 지각/조퇴/결과).
      // 2) If NEIS count is less than local count, we suppress the mismatch if the difference (localVal - neisVal) is within the local menstrual records.
      if (isMismatch && m.sub === "인정" && isFemale) {
        const mKey = `${statKey}_menstrual`;
        const localMenstrualRaw = studentStats[mKey];
        const localMenstrualVal = (localMenstrualRaw instanceof Set) ? localMenstrualRaw.size : (typeof localMenstrualRaw === 'number' ? localMenstrualRaw : 0);
        
        if (neisVal > localVal) {
          const diff = neisVal - localVal;
          const limit = (m.cat === "결석") ? 1 : 3;
          if (diff + localMenstrualVal <= limit) {
            isMismatch = false;
          }
        } else if (neisVal < localVal) {
          const diff = localVal - neisVal;
          if (diff <= localMenstrualVal) {
            isMismatch = false;
          }
        }
      }
      
      // School Experiential Learning (학교장허가체험학습) Exception:
      // If category is "인정", and there are at least 2 scanned documents (신청서 + 결과보고서) 
      // marked as "학교장허가체험학습", we suppress the mismatch error regardless of the day count difference.
      if (isMismatch && m.sub === "인정" && studentStats["체험학습_건수"] >= 2) {
        isMismatch = false;
      }

      if (isMismatch) {
        isMatch = false;
        const dateKey = `${statKey}_dates`;
        const localDates = studentStats[dateKey] || [];
        const dateSuffix = localDates.length > 0 ? ` [보관서류 날짜: ${localDates.join(", ")}]` : " [보관서류 날짜: 없음]";
        mismatches.push(`${m.cat}(${m.sub}) 나이스:${neisVal} / 서류:${localVal}${dateSuffix}`);
        bgRow.push('#fff2cc'); // Yellow for mismatch
      } else {
        bgRow.push(null);
      }
    });
    
    // 총계 (결석, 지각, 조퇴, 결과)
    newRow.push(absTotal, tardyTotal, earlyTotal, skipTotal);
    bgRow.push(null, null, null, null);
    
    // 일치여부, 불일치내용
    newRow.push(isMatch ? "✅ 일치" : "❌ 불일치", isMatch ? "-" : mismatches.join("\n"));
    bgRow.push(isMatch ? null : '#fce5cd', null); // Subtle orange for mismatch status
    
    // Check if duplicate (same month, studentId, name) exists in current AttendanceData
    const dataKey = `${rowMonth}_${sid}_${name}`;
    const matched = existingMap.get(dataKey);
    if (matched) {
      // Keep original sequence number (순번)
      newRow[0] = matched.values[0];
      
      // Overwrite the entry in existing data array
      existingData[matched.index] = newRow;
      existingBg[matched.index] = bgRow;
    } else {
      newRowsToAppend.push({ newRow, bgRow });
    }
  });
  
  // Append new non-existing rows
  newRowsToAppend.forEach(item => {
    existingData.push(item.newRow);
    existingBg.push(item.bgRow);
  });
  
  // Clear the existing content completely before writing back the merged data
  const lastRow = dataSheet.getLastRow();
  if (lastRow > 1) {
    const numRows = lastRow - 1;
    const clearRange = dataSheet.getRange(2, 1, numRows, 26);
    clearRange.clearContent();
    clearRange.setBackground(null);
  }
  
  // Write the merged data back to the sheet in a single call
  if (existingData.length > 0) {
    const targetRange = dataSheet.getRange(2, 1, existingData.length, 26);
    targetRange.setValues(existingData);
    targetRange.setBackgrounds(existingBg);
  }
  
  return "대조 및 AttendanceData 시트 업데이트 완료 (" + (targetMonth === "all" ? "전체 월" : targetMonth + "월") + " / " + (targetClass === "all" ? "전체 학급" : targetClass + "반") + ")";
}

/**
 * Save confirmed data to Sheet
 */
function saveToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  
  // 0. Final Normalization of Categories & Reason Details
  const norm = normalizeCategories(data.category, data.subCategory, data.filename, data.reasonDetail);
  data.category = norm.category;
  data.subCategory = norm.subCategory;
  data.reasonDetail = norm.reasonDetail;
  
  // 1. Duplicate Detection (Check Date & StudentId)
  const dateStr = data.date || "";
  const idStr = data.studentId || "";
  const existingRecords = sheet.getDataRange().getValues();
  
  const isDuplicate = existingRecords.some(row => {
    const d = parseDateSafe(row[0]);
    if (!d) return false;
    const rDate = Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
    return rDate === dateStr && row[1] == idStr;
  });
  
  if (isDuplicate) return "중복: [" + data.name + "] 이미 해당 날짜 및 학번의 기록이 존재합니다.";

  // 2. Student Directory Verification (Optional Info)
  let verificationNote = "";
  const studentInfo = verifyStudent(idStr, data.name);
  if (studentInfo.match && studentInfo.correctId && studentInfo.correctName) {
    // [Force] Use directory values for saving, not raw AI data
    data.studentId = studentInfo.correctId;
    data.name = studentInfo.correctName;
  } else {
    verificationNote = "⚠️ 명부 불일치";
  }

  // 3. Attachments format rule
  let attachments = data.attachments !== undefined ? data.attachments : "×";
  let subCat = (data.subCategory || "").trim();
  let reason = (data.reasonDetail || "").trim();
  
  // 3.5. Force NEIS string standard format: YYYY-MM-DD~YYYY-MM-DD 대분류소분류(사유상세)(X일간)
  const finalDate = data.date || "";
  const finalCat = data.category || "";
  const finalSub = data.subCategory || "";
  const finalReason = reason;
  
  // Parse total days
  let daysNum = 1;
  const daysRaw = data.totalDays || "1일간";
  const daysMatch = daysRaw.match(/(\d+)/);
  if (daysMatch) {
    daysNum = parseInt(daysMatch[1], 10);
  }
  
  // Calculate date range string
  let dateRangeStr = finalDate;
  if (daysNum > 1 && finalDate) {
    let endDateStr = "";
    if (data.periodEnd && typeof data.periodEnd === 'string') {
      const endMatch = data.periodEnd.match(/(\d{4}-\d{2}-\d{2})/);
      if (endMatch) {
        endDateStr = endMatch[1];
      }
    }
    
    // Fallback: calculate end date if not explicitly in periodEnd
    if (!endDateStr) {
      try {
        const parts = finalDate.split('-');
        if (parts.length === 3) {
          const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
          const endDateObj = new Date(startDateObj.getTime() + (daysNum - 1) * 24 * 60 * 60 * 1000);
          endDateStr = Utilities.formatDate(endDateObj, "GMT+9", "yyyy-MM-dd");
        }
      } catch (e) {}
    }
    
    if (endDateStr && endDateStr !== finalDate) {
      dateRangeStr = `${finalDate}~${endDateStr}`;
    }
  }
  
  const isReasonNone = (str) => !str || /^(none|null|(해당\s*)?없음|첨부서류\s*누락|×)$/i.test(String(str).trim());
  
  if (finalDate && finalCat && finalSub) {
    if (data.preserveNeisString && data.neisString) {
      // Preserve original unsplit neisString
    } else {
      const daySuffix = daysNum > 1 ? `(${daysNum}일간)` : "";
      if (!isReasonNone(finalReason)) {
        data.neisString = `${dateRangeStr} ${finalCat}${finalSub}(${finalReason})${daySuffix}`;
      } else {
        data.neisString = `${dateRangeStr} ${finalCat}${finalSub}${daySuffix}`;
      }
    }
  }

  // 4. Save as Hyperlink for premium experience
  const fileNameCell = `=HYPERLINK("${data.fileUrl}", "${data.filename}")`;

  const safeVal = (val) => val !== undefined ? val : '×';

  const formatSignForSheet = (val) => {
    if (val === "있음") return "";
    if (val === "없음") return "×";
    if (!val) return ""; // 빈칸 유지 (서명 있음)
    const s = String(val).trim();
    if (/^(무|x|×|no|none|n|false|서명\s*안됨|안됨|미서명|서명\s*없음|없음)$/i.test(s)) {
      return "×";
    }
    return ""; // 그 외에는 서명된 것으로 간주하여 빈칸 처리
  };

  sheet.appendRow([
    data.date || '',                // A (1) 일자
    data.studentId || '',           // B (2) 학번
    data.name || '',                // C (3) 성명
    formatSignForSheet(data.studentSigned),    // D (4) 학생서명
    data.category || '',            // E (5) 대분류
    subCat,                         // F (6) 소분류
    formatPeriodSafe(data.periodStart),      // G (7) 시작일시
    formatPeriodSafe(data.periodEnd),        // H (8) 종료일시
    data.totalDays || '',           // I (9) 일수
    (data.reasonDetail || '') + (verificationNote ? " ["+verificationNote+"]" : ""), // J (10) 사유상세
    data.parentName || '',          // K (11) 학부모성함
    formatSignForSheet(data.parentSigned),     // L (12) 학부모서명
    data.teacherName || '',         // M (13) 담임성함
    formatSignForSheet(data.teacherSigned),    // N (14) 담임서명
    attachments,                    // O (15) 첨부서류
    data.docStartDate || '',        // P (16) 통원일
    data.docEndDate || '',          // Q (17) 통원종료일
    data.issuer || '',              // R (18) 발급처
    fileNameCell,                   // S (19) 파일명
    data.ruleCheck || '',           // T (20) 규정위반여부
    new Date(),                     // U (21) 스캔일시
    data.neisString || '',          // V (22) NEIS문구
    ''                              // W (23) 보완완료
  ]);
  return "저장 완료!";
}

/**
 * Save Multiple Records Sequentially to Preserve Order
 */
function saveMultipleToSheet(items) {
  if (!items || items.length === 0) return "저장할 데이터가 없습니다.";
  
  let successCount = 0;
  let skipCount = 0;
  let messages = [];
  
  items.forEach(item => {
    const result = saveToSheet(item);
    if (result.includes("저장 완료")) {
      successCount++;
    } else {
      skipCount++;
      messages.push(result);
    }
  });
  
  let finalMsg = `총 ${items.length}건 중 ${successCount}건 저장 완료!`;
  if (skipCount > 0) finalMsg += ` (${skipCount}건 제외: ${messages.join(', ')})`;
  
  return finalMsg;
}

/**
 * Organize File into Subfolders (Year/Month/Class/Name)
 */
function organizeFile(file, data) {
  try {
    const rootFolder = getTargetFolder();
    const date = parseDateSafe(data.date) || new Date();
    const year = getYearSafe(date) + "년";
    const month = getMonthSafe(date) + "월";
    
    // Class folder (studentId: G-CC-NN)
    let classNum = "미분류";
    if (data.studentId && data.studentId.length >= 3) {
      classNum = parseInt(data.studentId.substring(1, 3)) + "반";
    }
    
    // Create folders if they don't exist
    const yearFolder = getSubFolder(rootFolder, year);
    const monthFolder = getSubFolder(yearFolder, month);
    const classFolder = getSubFolder(monthFolder, classNum);
    const studentFolder = getSubFolder(classFolder, data.name || "무명이");
    
    file.moveTo(studentFolder);
  } catch (e) {
    Logger.log("폴더 이동 실패: " + e.toString());
  }
}

function getSubFolder(parent, folderName) {
  const folders = parent.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(folderName);
}

/**
 * Helper: Get NEIS Folder
 */
function getNeisFolder() {
  const folderId = CONFIG.NEIS_FOLDER_ID;
  if (folderId && folderId.trim().length > 0) {
    try {
      // Split the ID if it contains query parameters (like ?dmr)
      const cleanId = folderId.split('?')[0]; 
      return DriveApp.getFolderById(cleanId);
    } catch (e) {
      Logger.log("경고: NEIS 폴더 ID가 잘못되었습니다. 기본 업로드 폴더를 사용합니다.");
    }
  }
  return getTargetFolder(); // Fallback
}

/**
 * Safe date parsing utility
 */
function parseDateSafe(val) {
  if (val instanceof Date) return val;
  if (!val) return null;
  
  const str = String(val).trim();
  if (str.length === 0) return null;
  
  // Extract date part (e.g., "2026-05-21 1교시" -> "2026-05-21")
  // Normalize dots to hyphens (e.g., "2026.05.21" -> "2026-05-21")
  const datePart = str.split(' ')[0].replace(/\./g, '-');
  const d = new Date(datePart);
  if (!isNaN(d.getTime())) return d;
  
  // Fallback for standard date parsing
  const fallback = new Date(str);
  if (!isNaN(fallback.getTime())) return fallback;
  
  return null;
}

/**
 * Safe month extraction (GMT+9 timezone consistent)
 */
function getMonthSafe(date) {
  if (!date || isNaN(date.getTime())) return null;
  const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  return parseInt(dateStr.split('-')[1], 10);
}

/**
 * Safe year extraction (GMT+9 timezone consistent)
 */
function getYearSafe(date) {
  if (!date || isNaN(date.getTime())) return null;
  const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  return parseInt(dateStr.split('-')[0], 10);
}

/**
 * Helper: Safely formats period input (Dates or strings) to 'YYYY-MM-DD X교시'
 */
function formatPeriodSafe(val) {
  if (!val) return "-";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+9", "yyyy-MM-dd");
  }
  
  let str = String(val).trim();
  
  // Clean up any leading zeros in "교시" (e.g. 01교시 -> 1교시)
  str = str.replace(/(\s+|[^0-9])0+(\d+)\s*교시/g, "$1$2교시");

  // If it's already in the correct format, return immediately
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str;
  }
  
  const parsed = parseDateSafe(str);
  if (parsed) {
    const matchLesson = str.match(/(\d+\s*교시)/);
    const datePart = Utilities.formatDate(parsed, "GMT+9", "yyyy-MM-dd");
    return matchLesson ? `${datePart} ${matchLesson[1]}` : datePart;
  }
  
  return str;
}

/**
 * Check if the student is active on a specific attendance date based on remarks and statusDate
 */
function isStudentActiveOnDate(status, statusDateRaw, checkDateRaw) {
  if (!status || !statusDateRaw) return true; // No special status -> active
  const statusDate = parseDateSafe(statusDateRaw);
  const checkDate = parseDateSafe(checkDateRaw);
  if (!statusDate || !checkDate) return true; // Safe fallback if dates cannot be parsed
  
  const sStr = Utilities.formatDate(statusDate, "GMT+9", "yyyy-MM-dd");
  const cStr = Utilities.formatDate(checkDate, "GMT+9", "yyyy-MM-dd");
  const sParts = sStr.split('-');
  const cParts = cStr.split('-');
  const sTime = new Date(parseInt(sParts[0], 10), parseInt(sParts[1], 10) - 1, parseInt(sParts[2], 10)).getTime();
  const cTime = new Date(parseInt(cParts[0], 10), parseInt(cParts[1], 10) - 1, parseInt(cParts[2], 10)).getTime();
  
  if (status.includes('자퇴') || status.includes('전출') || status.includes('퇴학')) {
    // Active up to the statusDate (inclusive)
    return cTime <= sTime;
  } else if (status.includes('전입')) {
    // Active from the statusDate (inclusive)
    return cTime >= sTime;
  }
  return true;
}

/**
 * Verify Student in Directory
 */
function verifyStudent(id, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return { match: true }; 

  const settings = getSystemSettings();

  const data = sheet.getDataRange().getValues().slice(1);
  const studentList = [];
  data.forEach(row => {
    if (!row[0] || !row[1] || !row[2]) return;
    const grade = parseInt(row[0], 10);
    if (grade !== settings.grade) return; // Only verify for target grade
    studentList.push({
      grade: row[0],
      cls: row[1],
      num: row[2],
      name: String(row[3]).trim(),
      id: `${row[0]}${row[1].toString().padStart(2, '0')}${row[2].toString().padStart(2, '0')}`,
      status: String(row[5] || "").trim(), // 비고 (자퇴, 전출, 전입 등)
      statusDate: row[6] // 기준일 (YYYY-MM-DD)
    });
  });

  const inputIdRaw = String(id || "").trim();
  const inputNameRaw = String(name || "").trim();

  // Strip '(공동)' from student name for comparison with StudentDirectory
  const inputNameClean = inputNameRaw.replace(/\(공동\)/g, "").trim();

  // 0. Clean Input ID (Handle common OCR errors: S->5, O->0, I/L->1, Z->2, B->8)
  const inputIdClean = inputIdRaw.toUpperCase()
    .replace(/S/g, "5")
    .replace(/O/g, "0")
    .replace(/[IL]/g, "1")
    .replace(/Z/g, "2")
    .replace(/B/g, "8")
    .replace(/\D/g, ""); // Remove any remaining non-digits

  // 1. Exact Name Match (Strongest fallback for messy digits)
  const nameMatches = studentList.filter(s => s.name === inputNameClean);
  if (nameMatches.length === 1) {
    return { match: true, correctId: nameMatches[0].id, correctName: nameMatches[0].name, status: nameMatches[0].status, statusDate: nameMatches[0].statusDate };
  } else if (nameMatches.length > 1) {
    // If namesake, try to find the one with closest ID similarity
    let closest = nameMatches[0];
    let maxSimilarity = 0;
    
    nameMatches.forEach(s => {
      let similarity = 0;
      for (let i = 0; i < Math.min(s.id.length, inputIdClean.length); i++) {
        if (s.id[i] === inputIdClean[i]) similarity++;
      }
      if (similarity > maxSimilarity) { maxSimilarity = similarity; closest = s; }
    });
    return { match: true, correctId: closest.id, correctName: closest.name, status: closest.status, statusDate: closest.statusDate };
  }

  // 2. ID Match (Using Cleaned ID)
  const idMatch = studentList.find(s => s.id === inputIdClean);
  if (idMatch) {
    // If ID is correct, we accept most names unless they are completely different
    // (e.g. at least one character match or name is short/long)
    return { match: true, correctId: idMatch.id, correctName: idMatch.name, status: idMatch.status, statusDate: idMatch.statusDate };
  }

  // 3. Last Resort: Partial ID match (4 out of 5 digits match) + Name partial match
  for (const s of studentList) {
    let idDiffCount = 0;
    if (s.id.length === inputIdClean.length) {
      for (let i = 0; i < s.id.length; i++) {
        if (s.id[i] !== inputIdClean[i]) idDiffCount++;
      }
    }
    
    // If 4/5 digits match and Name contains or is contained in Directory Name
    const nameSimilarity = inputNameClean.length >= 2 && (s.name.includes(inputNameClean) || inputNameClean.includes(s.name));
    if (idDiffCount <= 1 && nameSimilarity) {
      return { match: true, correctId: s.id, correctName: s.name, status: s.status, statusDate: s.statusDate };
    }
  }

  return { match: false };
}

/**
 * Get Student Directory gender map
 */
function getStudentDirectoryMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  const map = new Map();
  if (!sheet) return map;

  const settings = getSystemSettings();

  const data = sheet.getDataRange().getValues().slice(1);
  data.forEach(row => {
    if (!row[0] || !row[1] || !row[2]) return;
    const grade = parseInt(row[0], 10);
    if (grade !== settings.grade) return; // Only load target grade
    const cls = row[1];
    const num = row[2];
    const name = String(row[3]).trim();
    const gender = String(row[4] || "").trim(); // 성별
    const id = `${grade}${cls.toString().padStart(2, '0')}${num.toString().padStart(2, '0')}`;
    map.set(id, { name, gender });
  });
  return map;
}

/**
 * Retrieve all exam periods from Sheet
 */
function getExamPeriods() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.EXAM_PERIODS || 'ExamPeriods');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues().slice(1);
  const periods = [];
  data.forEach((row, index) => {
    if (row[0]) {
      let startStr = "";
      let endStr = "";
      try {
        const start = new Date(row[0]);
        const end = row[1] ? new Date(row[1]) : new Date(row[0]);
        if (!isNaN(start.getTime()) && !isNaN(end.getTime())) {
          startStr = Utilities.formatDate(start, "GMT+9", "yyyy-MM-dd");
          endStr = Utilities.formatDate(end, "GMT+9", "yyyy-MM-dd");
          periods.push({
            index: index + 2, // 1-indexed and skip header row
            startStr: startStr,
            endStr: endStr,
            name: String(row[2] || "").trim()
          });
        }
      } catch (e) {
        Logger.log("Error parsing exam period row: " + e.toString());
      }
    }
  });
  return periods;
}

/**
 * Add a new exam period to Sheet
 */
function addExamPeriod(startDateStr, endDateStr, name) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EXAM_PERIODS || 'ExamPeriods');
    if (!sheet) throw new Error("시험기간 시트가 없습니다.");

    // Validate dates
    const start = new Date(startDateStr);
    const end = new Date(endDateStr || startDateStr);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      throw new Error("올바른 날짜 형식이 아닙니다.");
    }

    const startFormatted = Utilities.formatDate(start, "GMT+9", "yyyy-MM-dd");
    const endFormatted = Utilities.formatDate(end, "GMT+9", "yyyy-MM-dd");

    sheet.appendRow([startFormatted, endFormatted, String(name || "").trim()]);
    
    // Auto-sort by start date in ascending order
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, 3).sort({ column: 1, ascending: true });
    }

    return { success: true, message: "시험 일정이 성공적으로 추가되었습니다." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Delete an exam period from Sheet
 */
function deleteExamPeriod(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EXAM_PERIODS || 'ExamPeriods');
    if (!sheet) throw new Error("시험기간 시트가 없습니다.");

    const lastRow = sheet.getLastRow();
    const idx = parseInt(rowIndex, 10);
    if (isNaN(idx) || idx < 2 || idx > lastRow) {
      throw new Error("유효하지 않은 일련번호입니다.");
    }

    sheet.deleteRow(idx);
    return { success: true, message: "시험 일정이 성공적으로 삭제되었습니다." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Cross-Day Validation: Check if the previous school day had a hospital document
 */
/**
 * Cross-Day Validation: Check if the previous school day had a hospital document
 */
function checkPreviousDayHospitalDoc(studentId, currentDateStr, currentBatchItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (!sheet) return false;
  
  const cleanDateStr = currentDateStr ? String(currentDateStr).split(' ')[0] : '';
  const current = new Date(cleanDateStr);
  if (isNaN(current.getTime())) return false; // Invalid date
  
  const hospitalKeywords = ["진료", "진단", "처방", "통원", "입원", "퇴원", "입퇴원", "입·퇴원", "입/퇴원", "입-퇴원", "소견", "약", "영수증", "보건", "보건실", "입실", "확인증", "병원", "약국", "의원", "치과", "한의원"];
  
  let mostRecentPriorDate = null;
  let mostRecentPriorHasHospital = false;

  const checkItem = (rId, rDateStr, rCat, rAttach, rIssuer) => {
    if (String(rId).trim() !== String(studentId).trim()) return;
    if (!rCat.includes('질병')) return;
    
    const rDate = new Date(rDateStr.split(' ')[0]);
    if (isNaN(rDate.getTime())) return;
    
    const diffTime = current.getTime() - rDate.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    // Must be a prior day and within 4 days (matching the consecutive disease rule in DataCleanup)
    if (diffDays > 0 && diffDays <= 4) {
      if (!mostRecentPriorDate || rDate > mostRecentPriorDate) {
        mostRecentPriorDate = rDate;
        const cleanAttach = String(rAttach || "").replace(/\s+/g, "");
        const cleanIssuer = String(rIssuer || "").replace(/\s+/g, "");
        mostRecentPriorHasHospital = hospitalKeywords.some(kw => cleanAttach.includes(kw) || cleanIssuer.includes(kw));
      }
    }
  };

  // 1. Check in the current batch first
  if (currentBatchItems && Array.isArray(currentBatchItems)) {
    currentBatchItems.forEach(item => {
      if (!item) return;
      const rId = String(item.studentId || "").trim();
      const rCat = String(item.category || "").trim();
      const rAttach = String(item.attachments || "").trim();
      const rIssuer = String(item.issuer || "").trim();
      const itemDate = item.date || "";
      
      // Calculate covered dates range
      let daysNum = 1;
      const daysRaw = item.totalDays || "1일간";
      const daysMatch = daysRaw.match(/(\d+)/);
      if (daysMatch) daysNum = parseInt(daysMatch[1], 10);

      try {
        const parts = itemDate.split('-');
        if (parts.length === 3) {
          const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
          for (let d = 0; d < daysNum; d++) {
            const currentDateObj = new Date(startDateObj.getTime() + d * 24 * 60 * 60 * 1000);
            const dStr = Utilities.formatDate(currentDateObj, "GMT+9", "yyyy-MM-dd");
            checkItem(rId, dStr, rCat, rAttach, rIssuer);
          }
        } else {
          checkItem(rId, itemDate, rCat, rAttach, rIssuer);
        }
      } catch (e) {
        checkItem(rId, itemDate, rCat, rAttach, rIssuer);
      }
    });
  }

  // 2. Check in the database (AttendanceDB)
  const data = sheet.getDataRange().getValues().slice(1);
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || !row[1]) continue;
    
    const d = parseDateSafe(row[0]);
    if (!d) continue;
    const rDateStr = Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
    const rId = String(row[1]).trim();
    const rCat = String(row[4] || "").trim();
    const rIssuer = String(row[13] || "").trim();
    const rAttach = String(row[14] || "").trim();
    
    checkItem(rId, rDateStr, rCat, rAttach, rIssuer);
  }

  return mostRecentPriorDate ? mostRecentPriorHasHospital : false;
}

/**
 * Cross-Day Validation: Check if the same date had a hospital document
 */
/**
 * Cross-Day Validation: Check if the same date had a hospital document
 */
function checkSameDayHospitalDoc(studentId, currentDateStr, currentBatchItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (!sheet) return false;
  
  const cleanDateStr = currentDateStr ? String(currentDateStr).split(' ')[0] : '';
  const current = new Date(cleanDateStr);
  if (isNaN(current.getTime())) return false; // Invalid date
  const targetDateStr = Utilities.formatDate(current, "GMT+9", "yyyy-MM-dd");
  const hospitalKeywords = ["진료", "진단", "처방", "통원", "입원", "퇴원", "입퇴원", "입·퇴원", "입/퇴원", "입-퇴원", "소견", "약", "영수증", "보건", "보건실", "입실", "확인증", "병원", "약국", "의원", "치과", "한의원"];

  // 1. Check in the current batch first
  if (currentBatchItems && Array.isArray(currentBatchItems)) {
    for (let i = 0; i < currentBatchItems.length; i++) {
      const item = currentBatchItems[i];
      if (!item) continue;
      const rId = String(item.studentId || "").trim();
      const rCat = String(item.category || "").trim();
      const rAttach = String(item.attachments || "").trim();
      const rIssuer = String(item.issuer || "").trim();
      
      const itemDate = item.date || "";
      if (rId === String(studentId).trim() && rCat.includes('질병')) {
        let daysNum = 1;
        const daysRaw = item.totalDays || "1일간";
        const daysMatch = daysRaw.match(/(\d+)/);
        if (daysMatch) daysNum = parseInt(daysMatch[1], 10);

        const coveredDates = [];
        try {
          const parts = itemDate.split('-');
          if (parts.length === 3) {
            const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
            for (let d = 0; d < daysNum; d++) {
              const currentDateObj = new Date(startDateObj.getTime() + d * 24 * 60 * 60 * 1000);
              coveredDates.push(Utilities.formatDate(currentDateObj, "GMT+9", "yyyy-MM-dd"));
            }
          } else {
            coveredDates.push(itemDate);
          }
        } catch (e) {
          coveredDates.push(itemDate);
        }

        if (coveredDates.includes(targetDateStr)) {
          const cleanAttach = String(rAttach).replace(/\s+/g, "");
          const cleanIssuer = String(rIssuer).replace(/\s+/g, "");
          if (hospitalKeywords.some(kw => cleanAttach.includes(kw) || cleanIssuer.includes(kw))) {
            return true;
          }
        }
      }
    }
  }

  // 2. Check in the database (AttendanceDB)
  const data = sheet.getDataRange().getValues().slice(1);
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || !row[1]) continue;
    
    const d = parseDateSafe(row[0]);
    if (!d) continue;
    const rDate = Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
    const rId = String(row[1]).trim();
    const rCat = String(row[4] || "").trim(); // 대분류
    const rIssuer = String(row[13] || "").trim();
    const rAttach = String(row[14] || "").trim();
    
    if (rId === String(studentId).trim() && rDate === targetDateStr && rCat.includes('질병')) {
      const cleanAttach = String(rAttach).replace(/\s+/g, "");
      const cleanIssuer = String(rIssuer).replace(/\s+/g, "");
      if (hospitalKeywords.some(kw => cleanAttach.includes(kw) || cleanIssuer.includes(kw))) {
        return true;
      }
    }
  }
  return false;
}


/**
 * Reset Database (For testing)
 */
function resetDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Reset AttendanceDB
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (attSheet && attSheet.getLastRow() > 1) {
    attSheet.getRange(2, 1, attSheet.getLastRow() - 1, attSheet.getLastColumn()).clearContent();
  }
  
  // 2. Reset NeisData
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  if (neisSheet && neisSheet.getLastRow() > 1) {
    neisSheet.getRange(2, 1, neisSheet.getLastRow() - 1, neisSheet.getLastColumn()).clearContent();
  }

  // 3. Reset SystemLog
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  if (logSheet && logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).clearContent();
  }
  
  return "데이터베이스 및 시스템 로그가 초기화되었습니다. (헤더 제외)";
}

/**
 * Compare local data with NEIS data (Detailed Version)
 */
function compareWithNeis(targetMonth) {
  targetMonth = targetMonth || "all";
  // Recalculate reconciliation first to update matching status in AttendanceData sheet
  reconcileNeisWithAttendance(targetMonth, "all");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
  
  if (!dataSheet) return [];
  
  const discrepancies = [];
  if (dataSheet.getLastRow() > 1) {
    const dData = dataSheet.getDataRange().getValues().slice(1);
    dData.forEach(row => {
      const rowMonth = String(row[1]).trim();
      
      // Filter by targetMonth
      if (targetMonth !== "all" && String(rowMonth) !== String(targetMonth)) {
        return;
      }
      
      const sidStr = String(row[2]).trim();
      const gradeNum = sidStr.length >= 3 ? parseInt(sidStr.charAt(0), 10) : -1;
      const cNum = sidStr.length >= 3 ? parseInt(sidStr.substring(1,3), 10) : -1;
      const settings = getSystemSettings();
      if (gradeNum !== settings.grade || cNum > settings.classes) return;

      const name = String(row[3]).trim();
      const matchStatus = String(row[24] || "");
      const mismatchDetails = String(row[25] || "");
      
      if (matchStatus.includes("❌")) {
        // Determine type of mismatch based on details
        let type = 'CONTENT_MISMATCH';
        
        // Highlight "미인정" or "서류누락" or "나이스누락"
        if (mismatchDetails.includes("나이스:0") || mismatchDetails.includes("서류:0")) {
          // If NEIS has records but local doesn't, it is MISSING_SCAN (스캔본 누락)
          // If NEIS has 0 and local has > 0, it is MISSING_NEIS (NEIS 쪽에 서류누락)
          if (/나이스:[1-9]\d* \/ 서류:0/.test(mismatchDetails)) {
            type = 'MISSING_SCAN';
          } else if (/나이스:0 \/ 서류:[1-9]\d*/.test(mismatchDetails)) {
            type = 'MISSING_NEIS';
          }
        }
        
        discrepancies.push({
          type: type,
          date: `${rowMonth}월`,
          studentId: sidStr,
          name: name,
          details: mismatchDetails
        });
      }
    });
  }
  
  return discrepancies;
}

/**
 * Retrieve dynamic dashboard data (Total scans, active month, acknowledged counts, top 5 rows)
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  
  const totalScans = attSheet ? Math.max(0, attSheet.getLastRow() - 1) : 0;
  
  let acknowledgedThisMonth = 0;
  let recentRecords = [];
  
  if (attSheet && attSheet.getLastRow() > 1) {
    const data = attSheet.getDataRange().getValues().slice(1);
    const sorted = [...data].reverse();
    
    for (let i = 0; i < Math.min(5, sorted.length); i++) {
       const row = sorted[i];
       const d = parseDateSafe(row[0]);
       const dateStr = d ? Utilities.formatDate(d, "GMT+9", "yyyy/MM/dd") : "-";
       const issuedStr = row[15] ? String(row[15]) : "-";
       
       recentRecords.push({
         date: dateStr,
         studentId: row[1] || '-',
         name: row[2] || '-',
         category: `${row[4] || ''} (${row[5] || ''})`,
         issuedDate: issuedStr,
         reason: row[9] || ''
       });
    }
    
    const currentMonth = getMonthSafe(new Date());
    data.forEach(row => {
       const dateVal = parseDateSafe(row[0]);
       if (dateVal && getMonthSafe(dateVal) === currentMonth) {
          if ((row[5] || '').toString().includes('인정')) {
            acknowledgedThisMonth++;
          }
       }
    });
  }
  
  let lastNeisMonth = '-';
  if (neisSheet && neisSheet.getLastRow() > 1) {
    const neisData = neisSheet.getDataRange().getValues().slice(1);
    const months = neisData.map(r => parseInt(r[1])).filter(x => !isNaN(x));
    if (months.length > 0) {
       lastNeisMonth = Math.max(...months);
    }
  }
  
  return {
    totalScans,
    lastNeisMonth,
    acknowledgedThisMonth,
    recentRecords
  };
}

/**
 * Filtered Dashboard Analytics for Unified UI
 */
function getDashboardAnalytics(targetMonth, targetClass) {
  // 대시보드 조회 시 최신 수식으로 결석통계(인정 제외) 등 U~Z열을 동적으로 재계산하여 동기화
  try {
    reconcileAttendanceDataSheetDirectly(true);
  } catch(e) {
    Logger.log("대시보드 조회 중 자동 재계산 실패: " + e.toString());
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  const studentSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  
  const settings = getSystemSettings();

  let allStudents = [];
  if (studentSheet && studentSheet.getLastRow() > 1) {
    const sData = studentSheet.getDataRange().getValues().slice(1);
    allStudents = sData.map(row => {
      const remarks = String(row[5] || "").trim();
      const statusDateRaw = row[6];
      const grade = parseInt(row[0], 10);
      
      let isExcluded = false;
      if (remarks && statusDateRaw) {
        const statusDate = parseDateSafe(statusDateRaw);
        if (statusDate) {
          const sYear = getYearSafe(statusDate);
          const sMonth = getMonthSafe(statusDate);
          const sDateOnly = new Date(sYear, sMonth - 1, statusDate.getDate()).getTime();
          
          if (targetMonth !== 'all') {
            const filterMonth = parseInt(targetMonth, 10);
            const filterYear = 2026; // Current academic year
            
            const sTimeMonth = sYear * 12 + sMonth;
            const fTimeMonth = filterYear * 12 + filterMonth;
            
            if (remarks.includes('자퇴') || remarks.includes('전출') || remarks.includes('퇴학')) {
              if (fTimeMonth > sTimeMonth) {
                isExcluded = true;
              }
            } else if (remarks.includes('전입')) {
              if (fTimeMonth < sTimeMonth) {
                isExcluded = true;
              }
            }
          } else {
            // 'all' (전체 월) 조회 시 학적 변동생 예외 처리
            const semesterStart = new Date(2026, 2, 2).getTime(); // 2026년 3월 2일 (새 학년도 첫 수업일 부근)
            if (remarks.includes('전입')) {
              // 학기 시작일 이후 중도 전입해 온 학생은 전체 개근 대상에서 제외
              if (sDateOnly > semesterStart) {
                isExcluded = true;
              }
            } else if (remarks.includes('자퇴') || remarks.includes('전출') || remarks.includes('퇴학')) {
              // 학기 도중 전출/자퇴/퇴학한 학생도 전체 기간을 채우지 못했으므로 제외
              isExcluded = true;
            }
          }
        }
      } else {
        // Fallback if no date is specified
        if (remarks.includes('자퇴') || remarks.includes('전출') || remarks.includes('퇴학')) {
          isExcluded = true;
        }
      }
      
      return {
        classNum: String(row[1]),
        name: String(row[3]),
        studentId: String(row[0]) + String(row[1]).padStart(2, '0') + String(row[2]).padStart(2, '0'),
        excluded: isExcluded || (grade !== settings.grade)
      };
    }).filter(s => s.name && s.name.trim() !== '' && !s.excluded);
  }

  if (targetClass && targetClass !== 'all') {
    allStudents = allStudents.filter(s => s.classNum === String(targetClass));
  }
  
  const hitStudentIds = new Set();
  
  let attCount = 0;
  let neisCount = 0;
  let filteredRecords = [];
  
  if (attSheet && attSheet.getLastRow() > 1) {
    const aData = attSheet.getDataRange().getValues().slice(1);
    aData.forEach(row => {
      const d = parseDateSafe(row[0]);
      if (d && !isNaN(d.getTime())) {
         const rowMonth = String(getMonthSafe(d));
         if (targetMonth === 'all' || rowMonth === String(targetMonth)) {
            const sidStr = String(row[1]);
            const gradeNum = sidStr.length >= 3 ? parseInt(sidStr.charAt(0), 10) : -1;
            const cNum = sidStr.length >= 3 ? parseInt(sidStr.substring(1,3), 10) : -1;
            if (gradeNum === settings.grade && (targetClass === 'all' || cNum === parseInt(targetClass, 10)) && cNum <= settings.classes) {
               hitStudentIds.add(sidStr);
               attCount++;
               
               let formattedIssuedDate = "-";
               if (row[15]) {
                 const idate = new Date(row[15]);
                 if (!isNaN(idate.getTime())) {
                   formattedIssuedDate = Utilities.formatDate(idate, "GMT+9", "yyyy-MM-dd");
                 } else {
                   formattedIssuedDate = String(row[15]);
                 }
               }

               filteredRecords.push({
                 date: row[0] ? Utilities.formatDate(d, "GMT+9", "yyyy/MM/dd") : "-",
                 studentId: row[1] || '-',
                 name: row[2] || '-',
                 category: `${row[4] || ''} (${row[5] || ''})`,
                 issuedDate: formattedIssuedDate,
                 reason: row[9] || '',
                 attachments: row[14] || '',
                 ruleCheck: row[19] || '',
                 actionStatus: row[22] || '' // 23rd column (W)
               });
            }
         }
      }
    });
  }
  filteredRecords.reverse();
  
  if (neisSheet && neisSheet.getLastRow() > 1) {
    const nData = neisSheet.getDataRange().getValues().slice(1);
    nData.forEach(row => {
      const rowMonth = String(row[1]);
      if (targetMonth === 'all' || rowMonth === String(targetMonth)) {
        const sidStr = String(row[2]);
        const gradeNum = sidStr.length >= 3 ? parseInt(sidStr.charAt(0), 10) : -1;
        const cNum = sidStr.length >= 3 ? parseInt(sidStr.substring(1,3), 10) : -1;
        if (gradeNum === settings.grade && (targetClass === 'all' || cNum === parseInt(targetClass, 10)) && cNum <= settings.classes) {
           hitStudentIds.add(sidStr);
           
           // columns 20 to 23: stats
           const sum = (parseInt(row[20])||0) + (parseInt(row[21])||0) + (parseInt(row[22])||0) + (parseInt(row[23])||0);
           neisCount += sum;
        }
      }
    });
  }
  
  const perfectStudents = allStudents.filter(s => !hitStudentIds.has(s.studentId));
  perfectStudents.sort((a,b) => parseInt(a.studentId) - parseInt(b.studentId));
  
  let mismatches = [];
  let targetMismatchedUniqueIds = new Set();
  
  const dataSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
  if (dataSheet && dataSheet.getLastRow() > 1) {
    const dData = dataSheet.getDataRange().getValues().slice(1);
    dData.forEach(row => {
      const rowMonth = String(row[1]);
      const sidStr = String(row[2]);
      const name = String(row[3]);
      const matchStatus = String(row[24] || "");
      const mismatchDetails = String(row[25] || "");
      
      if (matchStatus.includes("❌")) {
        if (targetMonth === 'all' || rowMonth === String(targetMonth)) {
          const gradeNum = sidStr.length >= 3 ? parseInt(sidStr.charAt(0), 10) : -1;
          const cNum = sidStr.length >= 3 ? parseInt(sidStr.substring(1,3), 10) : -1;
          if (gradeNum === settings.grade && (targetClass === 'all' || cNum === parseInt(targetClass, 10)) && cNum <= settings.classes) {
            mismatches.push({
              type: 'CONTENT_MISMATCH',
              date: `${rowMonth}월`,
              studentId: sidStr,
              name: name,
              details: mismatchDetails
            });
            targetMismatchedUniqueIds.add(sidStr);
          }
        }
      }
    });
  }
  
  // Calculate matchRate
  // Match Rate = (Total students mapped - uniquely mismatched students) / Total students mapped * 100
  let matchRate = "100%";
  if (allStudents.length > 0) {
    const correctCount = allStudents.length - targetMismatchedUniqueIds.size;
    matchRate = Math.round((correctCount / allStudents.length) * 100) + "%";
  } else if (allStudents.length === 0 && targetMismatchedUniqueIds.size > 0) {
    matchRate = "0%";
  }

  return {
    perfectList: perfectStudents,
    mismatchList: mismatches,
    attCount: attCount,
    neisCount: neisCount,
    matchRate: matchRate,
    filteredRecords: filteredRecords
  };
}

/**
 * Batch Loader for PDF Generative Loop
 */
function getAllClassesDashboardAnalytics(targetMonth) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const studentSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
   let classes = new Set();
   const settings = getSystemSettings();
   if (studentSheet && studentSheet.getLastRow() > 1) {
     const sData = studentSheet.getDataRange().getValues().slice(1);
     sData.forEach(r => {
       const grade = parseInt(r[0], 10);
       const cls = parseInt(r[1], 10);
       if (grade === settings.grade && cls <= settings.classes) {
         classes.add(cls);
       }
     });
   }
   
   let classArr = Array.from(classes)
                  .map(c => parseInt(c))
                  .filter(c => !isNaN(c))
                  .sort((a,b)=>a-b);
                  
   if (classArr.length === 0) {
     classArr = [];
     for (let i = 1; i <= settings.classes; i++) {
       classArr.push(i);
     }
   } 
   
   const results = [];
   classArr.forEach(c => {
     results.push({
       className: `${c}반`,
       data: getDashboardAnalytics(targetMonth, String(c))
     });
   });
   
   return results;
}

/**
 * Get all document records that cover multiple days (e.g. totalDays > 1) from AttendanceDB.
 * Used for the premium "Integrated Proof Documents Lookup" dashboard feature.
 */
function getMultiDayDocuments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const range = sheet.getDataRange();
  const values = range.getValues().slice(1);
  const formulas = range.getCell(2, 1).offset(0, 0, values.length, range.getLastColumn()).getFormulas();
  
  const results = [];
  
  values.forEach((row, i) => {
    const daysRaw = String(row[8] || "").trim(); // Column I (index 8) is 일수
    const matchDays = daysRaw.match(/(\d+)/);
    if (matchDays) {
      const days = parseInt(matchDays[1], 10);
      if (days > 1) {
        // Parse filename hyperlink from formulas
        const fileFormula = formulas[i][18] || ""; // Column S (index 18) is 파일명
        let fileUrl = "";
        let filename = String(row[18] || "-");
        
        if (fileFormula) {
          const matchLink = fileFormula.match(/=HYPERLINK\("([^"]+)"\s*,\s*"([^"]+)"\)/i);
          if (matchLink) {
            fileUrl = matchLink[1];
            filename = matchLink[2];
          }
        }
        
        // Extract Month from Date (Column A / index 0)
        let month = "-";
        if (row[0]) {
          const d = parseDateSafe(row[0]);
          if (d && !isNaN(d.getTime())) {
            month = `${getMonthSafe(d)}월`;
          }
        }
        
        const startFormatted = formatPeriodSafe(row[6]);
        const endFormatted = formatPeriodSafe(row[7]);
        
        let daysFormatted = daysRaw;
        if (daysRaw) {
          const match = daysRaw.match(/(\d+)/);
          if (match) {
            daysFormatted = `(${match[1]}일간)`;
          }
        }
        
        results.push({
          month: month,
          studentId: String(row[1]).trim(),
          name: String(row[2]).trim(),
          category: `${row[4] || ""} (${row[5] || ""})`,
          period: `${startFormatted} ~ ${endFormatted} ${daysFormatted}`,
          days: daysRaw,
          filename: filename,
          fileUrl: fileUrl
        });
      }
    }
  });
  
  // Sort results: month ascending, studentId ascending
  results.sort((a, b) => {
    const aM = parseInt(a.month) || 0;
    const bM = parseInt(b.month) || 0;
    if (aM !== bM) return aM - bM;
    return a.studentId.localeCompare(b.studentId);
  });
  
  return results;
}

/**
 * Normalize and repair 대분류 (category) and 소분류 (subCategory).
 * Also extracts unrecognized detail terms like 생리통 or 경조사 and merges them into J열 (reasonDetail).
 */
function normalizeCategories(catRaw, subCatRaw, filename, reasonDetail) {
  let category = String(catRaw || "").trim();
  let subCategory = String(subCatRaw || "").trim();
  let reason = String(reasonDetail || "").trim();
  const file = String(filename || "").trim();

  if (!category && !subCategory) {
    return { category: "질병", subCategory: "결석", reasonDetail: reason };
  }

  // 1. Extract extra details (non-standard keywords)
  let extCatDetail = category;
  let extSubDetail = subCategory;

  const catKeywords = ["출석인정", "미인정", "질병", "기타"];
  catKeywords.forEach(kw => { extCatDetail = extCatDetail.replace(kw, ""); });
  extCatDetail = extCatDetail.replace(/인정/g, ""); // "출석인정" 외에 단독 "인정"도 제거

  const subKeywords = ["결석", "지각", "조퇴", "결과", "깸석", "결서", "지가", "조태"];
  subKeywords.forEach(kw => { extSubDetail = extSubDetail.replace(kw, ""); });

  const cleanDetail = (str) => str.replace(/[\(\)\[\]\{\}\s,\/\-_~:;]/g, "").trim();
  extCatDetail = cleanDetail(extCatDetail);
  extSubDetail = cleanDetail(extSubDetail);

  // 2. Determine standardized Category and SubCategory
  const combined = (category + " " + subCategory + " " + file).trim();
  const lowerReason = reason.toLowerCase();
  const lowerCombined = combined.toLowerCase();

  // If the reason or filename/categories contain recognized absence keywords, force Category to "출석인정"
  const hasRecognizedKeyword = /생리|체험학습|독감|인플루엔자|코로나|경조|전염병|감염병|공가|학교장허가/.test(lowerReason) ||
                               /생리|체험학습|독감|인플루엔자|코로나|경조|전염병|감염병|공가|학교장허가/.test(lowerCombined);

  let finalCategory = "질병"; // default
  if (combined.includes("미인정")) {
    finalCategory = "미인정";
  } else if (combined.includes("출석인정") || combined.includes("출석 인정") || (combined.includes("인정") && !combined.includes("미인정")) || hasRecognizedKeyword) {
    finalCategory = "출석인정";
  } else if (combined.includes("질병")) {
    finalCategory = "질병";
  } else if (combined.includes("기타")) {
    finalCategory = "기타";
  }

  let finalSubCategory = "결석"; // default
  if (combined.includes("지각") || combined.includes("지가")) {
    finalSubCategory = "지각";
  } else if (combined.includes("조퇴") || combined.includes("조태")) {
    finalSubCategory = "조퇴";
  } else if (combined.includes("결과")) {
    finalSubCategory = "결과";
  } else if (combined.includes("결석") || combined.includes("깸석") || combined.includes("결서")) {
    finalSubCategory = "결석";
  }

  // 3. Extract and merge any residual details to J열 (reasonDetail)
  let extraDetails = [];
  const allStandardKeywords = ["출석인정", "미인정", "질병", "기타", "결석", "지각", "조퇴", "결과", "깸석", "결서", "지가", "조태", "인정"];
  
  if (extCatDetail && !allStandardKeywords.includes(extCatDetail)) {
    extraDetails.push(extCatDetail);
  }
  if (extSubDetail && !allStandardKeywords.includes(extSubDetail)) {
    extraDetails.push(extSubDetail);
  }

  if (extraDetails.length > 0) {
    const extraStr = extraDetails.filter((v, idx, self) => self.indexOf(v) === idx).join(", ");
    if (extraStr) {
      if (!reason) {
        reason = extraStr;
      } else {
        const cleanReason = reason.replace(/\s+/g, "");
        const cleanExtra = extraStr.replace(/\s+/g, "");
        if (!cleanReason.includes(cleanExtra) && !cleanExtra.includes(cleanReason)) {
          reason = reason + " (" + extraStr + ")";
        }
      }
    }
  }

  // 체험학습 관련 사유를 표준명칭인 '학교장허가체험학습'으로 통일하여 저장되도록 보완
  if (reason && /(학교장허가\s*|현장\s*|교외\s*)*체험\s*(학습)?/g.test(reason)) {
    reason = reason.replace(/(학교장허가\s*|현장\s*|교외\s*)*체험\s*(학습)?/g, "학교장허가체험학습");
  }

  return {
    category: finalCategory,
    subCategory: finalSubCategory,
    reasonDetail: reason
  };
}

/**
 * 스프레드가 열릴 때 상단에 편리한 사용자 맞춤 메뉴를 추가합니다.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ggomtrack AI')
    .addItem('🔄 NEIS 데이터 대조 재생성 (NeisData + AttendanceDB)', 'reconcileNeisWithAttendance')
    .addItem('✏️ 직접 수정한 AttendanceData 기준으로 Y:Z열 재계산', 'reconcileAttendanceDataSheetDirectly')
    .addSeparator()
    .addItem('🔑 임시: 하드코딩된 설정값을 스크립트 속성으로 이전', 'migrateLegacyConfigToProperties')
    .addToUi();
}

/**
 * 사용자가 AttendanceData 시트에서 직접 수정한 E~T열 수치(서류 건수)를 기반으로
 * U~X열(통계 합계) 및 Y:Z열(일치여부, 불일치내용)을 재계산하고 색상을 칠합니다.
 */
function reconcileAttendanceDataSheetDirectly(isSilent) {
  isSilent = (isSilent === true);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS || "NeisData");
  const dataSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
  
  if (!neisSheet || !dataSheet) {
    if (!isSilent) {
      SpreadsheetApp.getUi().alert("NeisData 또는 AttendanceData 시트가 존재하지 않습니다.");
    } else {
      Logger.log("오류: NeisData 또는 AttendanceData 시트가 존재하지 않습니다.");
    }
    return;
  }
  
  const neisRows = neisSheet.getDataRange().getValues().slice(1);
  const dataRows = dataSheet.getDataRange().getValues();
  const headers = dataRows[0];
  const dataBody = dataRows.slice(1);
  
  if (dataBody.length === 0) {
    if (!isSilent) {
      SpreadsheetApp.getUi().alert("대조할 데이터가 AttendanceData 시트에 존재하지 않습니다.");
    } else {
      Logger.log("경고: 대조할 데이터가 AttendanceData 시트에 존재하지 않습니다.");
    }
    return;
  }
  
  // 1. NEIS 데이터 맵화 (Key: 학번_월)
  const neisMap = new Map();
  neisRows.forEach(row => {
    if (!row[1] || !row[2]) return;
    const month = parseInt(row[1]);
    const sid = String(row[2]).trim();
    neisMap.set(`${sid}_${month}`, row);
  });

  // Count experiential learning and menstrual pain documents in AttendanceDB for each student_month
  const experientialCountMap = new Map();
  const menstrualCountMap = new Map(); // Key: studentId_month_category
  const settings = getSystemSettings();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (attSheet) {
    const attData = attSheet.getDataRange().getValues().slice(1);
    attData.forEach(row => {
      const date = parseDateSafe(row[0]);
      if (!date || isNaN(date.getTime())) return;
      const month = getMonthSafe(date);
      const sid = String(row[1]).trim();
      const gradeNum = sid.length >= 3 ? parseInt(sid.charAt(0), 10) : -1;
      const cNum = sid.length >= 3 ? parseInt(sid.substring(1,3), 10) : -1;
      if (gradeNum !== settings.grade || cNum > settings.classes) return;

      const subCatRaw = String(row[4] || "").trim();
      const catRaw = String(row[5] || "").trim();
      const reasonDetail = String(row[9] || "").trim();
      
      let subCat = "기타";
      if (subCatRaw.includes("인정")) subCat = "인정";

      let cat = "결석";
      if (catRaw.includes("지각")) cat = "지각";
      else if (catRaw.includes("조퇴")) cat = "조퇴";
      else if (catRaw.includes("결과")) cat = "결과";

      const isExperiential = reasonDetail.includes("학교장허가체험학습") || 
                             reasonDetail.includes("현장체험") || 
                             reasonDetail.includes("교외체험") || 
                             reasonDetail.includes("체험학습");
      if (subCat === "인정" && isExperiential) {
        const key = `${sid}_${month}`;
        experientialCountMap.set(key, (experientialCountMap.get(key) || 0) + 1);
      }

      const isMenstrual = subCat === "인정" && /생리/.test(reasonDetail);
      if (isMenstrual) {
        const key = `${sid}_${month}_${cat}`;
        menstrualCountMap.set(key, (menstrualCountMap.get(key) || 0) + 1);
      }
    });
  }
  
  const studentMap = getStudentDirectoryMap();
  const examPeriods = getExamPeriods();
  
  // NEIS와 AttendanceData 열 매핑
  const mapping = [
    { col: 4, cat: "결석", sub: "질병" }, { col: 5, cat: "결석", sub: "미인정" }, { col: 6, cat: "결석", sub: "기타" }, { col: 7, cat: "결석", sub: "인정" },
    { col: 8, cat: "지각", sub: "질병" }, { col: 9, cat: "지각", sub: "미인정" }, { col: 10, cat: "지각", sub: "기타" }, { col: 11, cat: "지각", sub: "인정" },
    { col: 12, cat: "조퇴", sub: "질병" }, { col: 13, cat: "조퇴", sub: "미인정" }, { col: 14, cat: "조퇴", sub: "기타" }, { col: 15, cat: "조퇴", sub: "인정" },
    { col: 16, cat: "결과", sub: "질병" }, { col: 17, cat: "결과", sub: "미인정" }, { col: 18, cat: "결과", sub: "기타" }, { col: 19, cat: "결과", sub: "인정" }
  ];
  
  const updatedRows = [];
  const bgRows = [];
  
  dataBody.forEach((row) => {
    const rowMonth = parseInt(row[1]);
    const sid = String(row[2]).trim();
    
    // 만약 월이나 학번이 비어있으면 그대로 둠
    if (isNaN(rowMonth) || !sid) {
      updatedRows.push(row);
      bgRows.push(new Array(26).fill(null));
      return;
    }

    const gradeNum = sid.length >= 3 ? parseInt(sid.charAt(0), 10) : -1;
    const cNum = sid.length >= 3 ? parseInt(sid.substring(1,3), 10) : -1;
    if (gradeNum !== settings.grade || cNum > settings.classes) {
      updatedRows.push(row);
      bgRows.push(new Array(26).fill(null));
      return;
    }
    
    const neisRow = neisMap.get(`${sid}_${rowMonth}`);
    const studentInfo = studentMap.get(sid);
    const isFemale = studentInfo && (String(studentInfo.gender).trim() === '여' || String(studentInfo.gender).trim() === '여자');
    
    // 시험기간 확인
    const hasExamInMonth = examPeriods.some(period => {
      if (!period.startStr) return false;
      const startMonth = parseInt(period.startStr.split('-')[1], 10);
      const endMonth = period.endStr ? parseInt(period.endStr.split('-')[1], 10) : startMonth;
      return startMonth === rowMonth || endMonth === rowMonth;
    });
    
    // 순번(0), 월(1), 학번(2), 성명(3) 복사
    const newRow = [row[0], row[1], row[2], row[3]];
    const bgRow = new Array(26).fill(null);
    
    let absTotal = 0;
    let tardyTotal = 0;
    let earlyTotal = 0;
    let skipTotal = 0;
    
    let mismatches = [];
    let isMatch = true;
    
    // E(4) ~ T(19)열 데이터 검사 및 비교
    mapping.forEach((m) => {
      const localVal = parseInt(row[m.col]) || 0; // 사용자가 직접 수정한 로컬 값 읽기
      newRow.push(localVal); // 그대로 저장
      
      if (m.cat === "결석" && m.sub !== "인정") absTotal += localVal;
      else if (m.cat === "지각" && m.sub !== "인정") tardyTotal += localVal;
      else if (m.cat === "조퇴" && m.sub !== "인정") earlyTotal += localVal;
      else if (m.cat === "결과" && m.sub !== "인정") skipTotal += localVal;
      
      let neisVal = 0;
      if (neisRow) {
        neisVal = parseInt(neisRow[m.col]) || 0;
      }
      
      let isMismatch = neisVal !== localVal;
      
      // 미인정 예외
      if (isMismatch && m.sub === "미인정") {
        isMismatch = false;
      }
      
      // 생리통 예외
      if (isMismatch && m.sub === "인정" && isFemale) {
        const mKey = `${sid}_${rowMonth}_${m.cat}`;
        const menstrualCount = menstrualCountMap.get(mKey) || 0;
        if (neisVal > localVal) {
          const diff = neisVal - localVal;
          const limit = (m.cat === "결석") ? 1 : 3;
          if (diff + menstrualCount <= limit) {
            isMismatch = false;
          }
        } else if (neisVal < localVal) {
          const diff = localVal - neisVal;
          if (diff <= menstrualCount) {
            isMismatch = false;
          }
        }
      }
      
      // School Experiential Learning (학교장허가체험학습) Exception:
      // If category is "인정", and there are at least 2 scanned documents (신청서 + 결과보고서)
      // marked as "학교장허가체험학습" in the database, suppress mismatch error regardless of the day difference.
      const expCount = experientialCountMap.get(`${sid}_${rowMonth}`) || 0;
      if (isMismatch && m.sub === "인정" && expCount >= 2) {
        isMismatch = false;
      }
      
      if (isMismatch) {
        isMatch = false;
        mismatches.push(`${m.cat}(${m.sub}) 나이스:${neisVal} / 서류:${localVal}`);
        bgRow[m.col] = '#fff2cc'; // Mismatch cell yellow background
      }
    });
    
    // U~X열 (결석, 지각, 조퇴, 결과 총계) 수식 업데이트
    newRow.push(absTotal, tardyTotal, earlyTotal, skipTotal);
    
    // Y: 일치여부, Z: 불일치내용
    newRow.push(isMatch ? "✅ 일치" : "❌ 불일치", isMatch ? "-" : mismatches.join("\n"));
    
    if (!isMatch) {
      bgRow[24] = '#fce5cd'; // Y열 배경색
    }
    
    updatedRows.push(newRow);
    bgRows.push(bgRow);
  });
  
  // 시트에 일괄 저장
  const targetRange = dataSheet.getRange(2, 1, updatedRows.length, 26);
  targetRange.setValues(updatedRows);
  targetRange.setBackgrounds(bgRows);
  
  if (!isSilent) {
    SpreadsheetApp.getUi().alert("🎉 수동 수정된 수치를 기반으로 Y:Z열 대조 재계산이 완료되었습니다!");
  }
}

function runDiagnosticForStudent(name) {
  name = name || "권예빈";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const rows = attSheet.getDataRange().getValues();
  const matched = [];
  rows.forEach((row, i) => {
    if (row[2] && String(row[2]).includes(name)) {
      matched.push({
        rowNum: i + 1,
        date: row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT+9", "yyyy-MM-dd") : String(row[0]),
        sid: row[1],
        name: row[2],
        category: row[4],
        subCategory: row[5],
        start: row[6],
        end: row[7],
        days: row[8],
        neisString: row[21]
      });
    }
  });
  Logger.log(JSON.stringify(matched, null, 2));
  return JSON.stringify(matched, null, 2);
}

/**
 * Validate attendance document rules (retrieved from AI metadata)
 */
function checkAttendanceRule(item) {
  return item.ruleCheck || "";
}

/**
 * Retrieve System Settings (Grade and Classes) from PropertiesService
 */
function getSystemSettings() {
  const props = PropertiesService.getScriptProperties();
  const grade = props.getProperty('SYSTEM_GRADE') || '3';
  const classes = props.getProperty('SYSTEM_CLASSES') || '11';
  const geminiApiKey = props.getProperty('GEMINI_API_KEY') || '';
  const uploadFolderId = props.getProperty('UPLOAD_FOLDER_ID') || '';
  const neisFolderId = props.getProperty('NEIS_FOLDER_ID') || '';
  return {
    grade: parseInt(grade, 10),
    classes: parseInt(classes, 10),
    geminiApiKey: geminiApiKey,
    uploadFolderId: uploadFolderId,
    neisFolderId: neisFolderId
  };
}

/**
 * Save System Settings (Grade and Classes) to PropertiesService
 */
function saveSystemSettings(grade, classes) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SYSTEM_GRADE', String(grade));
    props.setProperty('SYSTEM_CLASSES', String(classes));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Retrieve names of students in the target grade from StudentDirectory
 */
function getTargetGradeStudentNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
    if (!sheet) return [];

    const settings = getSystemSettings();
    const data = sheet.getDataRange().getValues().slice(1);
    const studentNames = [];
    data.forEach(row => {
      if (!row[0] || !row[3]) return;
      const grade = parseInt(row[0], 10);
      if (grade === settings.grade) {
        const name = String(row[3]).trim();
        if (name && !studentNames.includes(name)) {
          studentNames.push(name);
        }
      }
    });
    return studentNames;
  } catch (e) {
    Logger.log("getTargetGradeStudentNames 오류: " + e.toString());
    return [];
  }
}

/**
 * Save API Key and Folder IDs to PropertiesService
 */
function saveApiSettings(geminiApiKey, uploadFolderId, neisFolderId) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('GEMINI_API_KEY', String(geminiApiKey || '').trim());
    props.setProperty('UPLOAD_FOLDER_ID', String(uploadFolderId || '').trim());
    props.setProperty('NEIS_FOLDER_ID', String(neisFolderId || '').trim());
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Temporary migration utility to transfer legacy hardcoded credentials in Constants.gs to Script Properties
 */
function migrateLegacyConfigToProperties() {
  try {
    const props = PropertiesService.getScriptProperties();
    // Use the values currently loaded in CONFIG (which are hardcoded in Constants.gs right now)
    props.setProperty('GEMINI_API_KEY', String(CONFIG.GEMINI_API_KEY || '').trim());
    props.setProperty('UPLOAD_FOLDER_ID', String(CONFIG.UPLOAD_FOLDER_ID || '').trim());
    props.setProperty('NEIS_FOLDER_ID', String(CONFIG.NEIS_FOLDER_ID || '').trim());
    SpreadsheetApp.getUi().alert(
      "🎉 성공: 기존 하드코딩 설정값들이 구글 스크립트 속성(Script Properties)으로 이전되었습니다.\n\n" +
      "이제 로컬 Constants.gs의 GEMINI_API_KEY, UPLOAD_FOLDER_ID, NEIS_FOLDER_ID 값을 빈 문자열('')로 비우고 소스코드를 푸시해도 문제없이 작동합니다."
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert("오류 발생: 스크립트 속성 이전 실패 (" + e.toString() + ")");
  }
}
/**
 * Handle Web App Access
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('AI 출결 관리 시스템')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Include separate HTML/JS files in index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Initialize Sheets if they don't exist
 */
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
    if (!sheet) return "AttendanceDB 시트를 찾을 수 없습니다.";
    const data = sheet.getDataRange().getValues().slice(1);
    
    const debugRows = [];
    data.forEach((row, i) => {
      const dateVal = row[0];
      const date = parseDateSafe(dateVal);
      if (!date) return;
      const month = getMonthSafe(date);
      if (month === 5) {
        debugRows.push({
          rowNum: i + 2,
          dateRaw: String(row[0]),
          dateParsed: Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd"),
          studentIdRaw: String(row[1]),
          name: String(row[2]),
          category: String(row[4]),
          subCategory: String(row[5]),
          days: String(row[8]),
          ruleCheck: String(row[19])
        });
      }
    });
    return "5월 데이터 분석 결과:\n" + JSON.stringify(debugRows, null, 2);
  } catch (e) {
    return "에러 발생: " + e.toString();
  }
}

/**
 * Initialize Log Sheet
 */
function initializeLogSheet(ss) {
  const sheetName = CONFIG.SHEETS.LOG || 'SystemLog';
  let logSheet = ss.getSheetByName(sheetName);
  if (!logSheet) {
    logSheet = ss.insertSheet(sheetName);
    const headers = ['일시', '모델', '파일명', '상태', '상세내용'];
    logSheet.appendRow(headers);
    logSheet.getRange(1, 1, 1, headers.length).setBackground('#999999').setFontColor('white').setFontWeight('bold');
  }
  return logSheet;
}

/**
 * Log to SystemLog sheet
 */
function logToSystem(filename, status, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG) || initializeLogSheet(ss);
    logSheet.appendRow([new Date(), CONFIG.GEMINI_MODEL, filename, status, message]);
  } catch (e) {
    Logger.log("로그 기록 실패: " + e.toString());
  }
}

/**
 * Upload Image/PDF to Drive and Process with Gemini
 */
function processAttendanceFile(filename, base64Data) {
  try {
    // 1. Save to Drive
    const folder = getTargetFolder();
    const contentType = base64Data.split(',')[0].split(':')[1].split(';')[0];
    const base64Content = base64Data.split(',')[1];
    const bytes = Utilities.base64Decode(base64Content);
    const fileBlob = Utilities.newBlob(bytes, contentType, filename);
    const file = folder.createFile(fileBlob);
    
    // 2. Call Gemini Vision API
    let studentNames = [];
    try {
      studentNames = getTargetGradeStudentNames();
    } catch (e) {
      Logger.log("대상 학생 명단 가져오기 실패: " + e.toString());
    }
    const extractedDataArray = callGeminiVision(base64Content, contentType, filename, studentNames);
    
    // 3. Post-processing: Correct Names & Link File
    const fileUrl = file.getUrl();
    const resultWithUrl = [];
    const allSplitItems = [];
    
    extractedDataArray.forEach(rawItem => {
      const item = { ...rawItem };
      const studentInfo = verifyStudent(item.studentId, item.name);
      
      // Normalize categories, subcategories, and reason details
      const norm = normalizeCategories(item.category, item.subCategory, filename, item.reasonDetail);
      item.category = norm.category;
      item.subCategory = norm.subCategory;
      item.reasonDetail = norm.reasonDetail;
      
      // Normalize periodStart and periodEnd format (e.g. 01교시 -> 1교시)
      item.periodStart = formatPeriodSafe(item.periodStart);
      item.periodEnd = formatPeriodSafe(item.periodEnd);
      
      // Date Override: Record actual absence date (periodStart) instead of document issuance date (date)
      if (item.periodStart && item.periodStart.includes('-')) {
        item.date = item.periodStart;
      }
      
      // OCR Correction: AI frequently misinterprets '1' in '1일간' as '|', '/', 'I', 'l', '(', '0' or completely misses it
      if (item.totalDays && typeof item.totalDays === 'string') {
        item.totalDays = item.totalDays.replace(/\s+/g, ""); // Remove spaces
        if (/^[\(]?([|/Il01(]|)\)?일간[\)]?$/i.test(item.totalDays)) {
           item.totalDays = "1일간";
        } else {
           item.totalDays = item.totalDays.replace(/[\(\)]/g, "");
        }
      }
      
      // Terminology Normalization: Unify various experiential learning terms
      if (item.reasonDetail && typeof item.reasonDetail === 'string') {
        item.reasonDetail = item.reasonDetail.replace(/(학교장허가\s*|현장\s*|교외\s*)*체험\s*(학습)?/g, "학교장허가체험학습");
      }
      
      // Clean up 'None', 'null', '없음', '해당없음' strings from AI output
      const cleanNone = (str) => (typeof str === 'string' && /^(none|null|(해당\s*)?없음|첨부서류\s*누락)$/i.test(str.trim())) ? "" : str;
      item.name = cleanNone(item.name);
      item.parentName = cleanNone(item.parentName);
      item.teacherName = cleanNone(item.teacherName);
      
      // Normalize signature fields to "있음" or "없음"
      const normalizeSign = (val) => {
        if (!val) return "없음";
        const s = String(val).trim();
        if (/^(무|x|×|no|none|n|false|서명\s*안됨|안됨|미서명|서명\s*없음|없음|null)$/i.test(s)) {
          return "없음";
        }
        return "있음";
      };
      item.studentSigned = normalizeSign(item.studentSigned);
      item.parentSigned = normalizeSign(item.parentSigned);
      item.teacherSigned = normalizeSign(item.teacherSigned);
      
      item.attachments = cleanNone(item.attachments);
      
      // Apply correct ID and Name
      if (studentInfo.match && studentInfo.correctId && studentInfo.correctName) {
        item.studentId = studentInfo.correctId;
        item.name = studentInfo.correctName;
      }

      // 4. Force NEIS string standard format: YYYY-MM-DD~YYYY-MM-DD 대분류소분류(사유상세)(X일간)
      const normDate = item.date || "";
      const normCat = item.category || "";
      const normSub = item.subCategory || "";
      const normReason = cleanNone(item.reasonDetail) || "";
      
      let normDaysNum = 1;
      const normDaysRaw = item.totalDays || "1일간";
      const normDaysMatch = normDaysRaw.match(/(\d+)/);
      if (normDaysMatch) {
        normDaysNum = parseInt(normDaysMatch[1], 10);
      }
      
      let normDateRangeStr = normDate;
      if (normDaysNum > 1 && normDate) {
        let endDateStr = "";
        if (item.periodEnd && typeof item.periodEnd === 'string') {
          const endMatch = item.periodEnd.match(/(\d{4}-\d{2}-\d{2})/);
          if (endMatch) endDateStr = endMatch[1];
        }
        if (!endDateStr) {
          try {
            const parts = normDate.split('-');
            if (parts.length === 3) {
              const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
              const endDateObj = new Date(startDateObj.getTime() + (normDaysNum - 1) * 24 * 60 * 60 * 1000);
              endDateStr = Utilities.formatDate(endDateObj, "GMT+9", "yyyy-MM-dd");
            }
          } catch (e) {}
        }
        if (endDateStr && endDateStr !== normDate) {
          normDateRangeStr = `${normDate}~${endDateStr}`;
        }
      }
      
      if (normDate && normCat && normSub) {
        const normDaySuffix = normDaysNum > 1 ? `(${normDaysNum}일간)` : "";
        if (normReason) {
          item.neisString = `${normDateRangeStr} ${normCat}${normSub}(${normReason})${normDaySuffix}`;
        } else {
          item.neisString = `${normDateRangeStr} ${normCat}${normSub}${normDaySuffix}`;
        }
      } else {
        item.neisString = cleanNone(item.neisString);
      }

      // Split item if it crosses months
      const splitList = splitMultiMonthItem(item);
      const isSplit = splitList.length > 1;

      splitList.forEach(splitItem => {
        if (isSplit) {
          splitItem.preserveNeisString = true;
          splitItem.neisString = item.neisString; // Inherit the original unsplit neisString
        } else {
          splitItem.neisString = item.neisString;
        }
        splitItem.studentInfo = studentInfo; // Temporarily attach student directory info
        allSplitItems.push(splitItem);
      });
    });

    // Second Pass: Run rule check and validation across all split items in the batch
    allSplitItems.forEach((splitItem, splitIndex) => {
      const studentInfo = splitItem.studentInfo || {};

      // Rule checks
      splitItem.ruleCheck = checkAttendanceRule(splitItem);
      
      if (splitItem.ruleCheck) {
        splitItem.ruleCheck = splitItem.ruleCheck.replace(/⚠️ 보완 필요: (학생 |학부모 )?(이름|성명|학번) 불일치/g, "")
                                       .replace(/\//g, " ⚠️ ")
                                       .replace(/  +/g, " ")
                                       .trim();
        if (splitItem.ruleCheck.startsWith("오류") || splitItem.ruleCheck.startsWith("보완")) {
          splitItem.ruleCheck = "⚠️ " + splitItem.ruleCheck;
        }
      }
      
      // 생리결석 한도 검증 로직 (1개월: 결석 1회 또는 지각/조퇴/결과 3회. 혼용 불가/한도 초과 시 강등)
      if (splitItem.category && splitItem.category.includes("출석인정") && splitItem.reasonDetail && splitItem.reasonDetail.includes("생리")) {
        const targetDate = splitItem.date || splitItem.periodStart || "";
        const yearMonthStr = targetDate.substring(0, 7);
        if (yearMonthStr) {
          const usage = getMonthlyMenstrualUsage(splitItem.studentId, yearMonthStr, allSplitItems.slice(0, splitIndex));
          let isLimitExceeded = false;
          
          if (splitItem.subCategory && splitItem.subCategory.includes("결석")) {
            // 결석: 기존 결석 1회 이상 또는 부분결석(조퇴 등) 1회 이상이면 초과
            if (usage.absenceCount >= 1 || usage.partialCount > 0) {
              isLimitExceeded = true;
            }
          } else if (splitItem.subCategory && (splitItem.subCategory.includes("지각") || splitItem.subCategory.includes("조퇴") || splitItem.subCategory.includes("결과"))) {
            // 지각/조퇴/결과: 기존 결석 1회 이상 또는 부분결석 3회 이상이면 초과
            if (usage.absenceCount >= 1 || usage.partialCount >= 3) {
              isLimitExceeded = true;
            }
          }
          
          if (isLimitExceeded) {
            // 서류 유무에 따라 질병 또는 미인정으로 전환
            const cleanAttach = splitItem.attachments ? String(splitItem.attachments).replace(/\s+/g, "") : "";
            const cleanIssuer = splitItem.issuer ? String(splitItem.issuer).replace(/\s+/g, "") : "";
            const hospitalKeywords = ["진료", "진단", "처방", "통원", "입원", "퇴원", "입퇴원", "입·퇴원", "입/퇴원", "입-퇴원", "소견", "약", "영수증", "보건", "보건실", "입실", "확인증", "병원", "약국", "의원", "치과", "한의원"];
            const hasHospitalDoc = hospitalKeywords.some(kw => cleanAttach.includes(kw) || cleanIssuer.includes(kw));
            const hasSubstituteDoc = ["학부모", "담임", "확인서", "의견서", "서명"].some(kw => cleanAttach.includes(kw));
            
            if (hasHospitalDoc || hasSubstituteDoc) {
              splitItem.category = "질병";
              splitItem.ruleCheck = "⚠️ 체크필요 [생리결석 허용 횟수 초과로 질병결석 전환됨]";
            } else {
              splitItem.category = "미인정";
              splitItem.ruleCheck = "⚠️ 체크필요 [생리결석 허용 횟수 초과 및 증빙서류 누락으로 미인정결석 전환됨]";
            }
            
            // neisString 재구성
            const normCat = splitItem.category;
            const normSub = splitItem.subCategory;
            const normReason = splitItem.reasonDetail;
            let daysNum = 1;
            const daysMatch = String(splitItem.totalDays || "1일간").match(/(\d+)/);
            if (daysMatch) daysNum = parseInt(daysMatch[1], 10);
            const normDaySuffix = daysNum > 1 ? `(${daysNum}일간)` : "";
            const dateRangeStr = splitItem.neisString.split(' ')[0];
            
            if (normReason) {
              splitItem.neisString = `${dateRangeStr} ${normCat}${normSub}(${normReason})${normDaySuffix}`;
            } else {
              splitItem.neisString = `${dateRangeStr} ${normCat}${normSub}${normDaySuffix}`;
            }
          } else {
            // 한도 내라면 AI가 잘못 달아둔 누락 경고 제거
            splitItem.ruleCheck = "";
          }
        }
      }

      // Cross-Day Validation for Disease Attachments
      const isDisease = (splitItem.category && splitItem.category.includes("질병"));
      if (isDisease) {
        const cleanAttach = splitItem.attachments ? String(splitItem.attachments).replace(/\s+/g, "") : "";
        const cleanIssuer = splitItem.issuer ? String(splitItem.issuer).replace(/\s+/g, "") : "";
        const hospitalKeywords = ["진료", "진단", "처방", "통원", "입원", "퇴원", "입퇴원", "입·퇴원", "입/퇴원", "입-퇴원", "소견", "약", "영수증", "보건", "보건실", "입실", "확인증", "병원", "약국", "의원", "치과", "한의원"];
        const hasHospitalDoc = hospitalKeywords.some(kw => cleanAttach.includes(kw) || cleanIssuer.includes(kw));
        const hasSubstituteDoc = ["학부모", "담임", "확인서", "의견서", "서명"].some(kw => cleanAttach.includes(kw));
        
        if (hasHospitalDoc) {
          splitItem.ruleCheck = ""; // valid
        } else {
          // If the item itself doesn't have a hospital document, check if there is a hospital document for the SAME day
          const targetDate = splitItem.date || splitItem.periodStart;
          const sameDayHasHospitalDoc = targetDate ? checkSameDayHospitalDoc(splitItem.studentId, targetDate, allSplitItems) : false;
          
          if (sameDayHasHospitalDoc) {
            splitItem.ruleCheck = ""; // valid
          } else if (hasSubstituteDoc) {
            splitItem.ruleCheck = "질병 대체 서류 검증 필요";
          } else {
            splitItem.ruleCheck = "체크필요 [첨부서류 누락]";
          }
        }
      }

      // 학적 변동 기준일(자퇴, 전출, 전입) 출결 체크
      if (studentInfo.match && studentInfo.status && studentInfo.statusDate) {
        const checkDate = splitItem.date || splitItem.periodStart;
        const isActive = isStudentActiveOnDate(studentInfo.status, studentInfo.statusDate, checkDate);
        if (!isActive) {
          const sDateStr = studentInfo.statusDate instanceof Date ? Utilities.formatDate(studentInfo.statusDate, "GMT+9", "yyyy-MM-dd") : String(studentInfo.statusDate);
          if (studentInfo.status.includes('자퇴') || studentInfo.status.includes('전출') || studentInfo.status.includes('퇴학')) {
            splitItem.ruleCheck = `체크필요 [${studentInfo.status} 학생 - 기준일(${sDateStr}) 이후 출결 기록]`;
          } else if (studentInfo.status.includes('전입')) {
            splitItem.ruleCheck = `체크필요 [전입 학생 - 기준일(${sDateStr}) 이전 출결 기록]`;
          }
        }
      }

      // 출결신고서 누락 검증 (단, 고등학교 발행 정식 공문은 제외)
      const hasReport = splitItem.hasReportCard === true;
      const isOfficial = splitItem.isOfficialDocument === true;
      const issuer = String(splitItem.issuer || "").trim();
      const isSchoolOfficial = isOfficial && 
                               (issuer.includes("고등학교") || issuer.includes("학교")) && 
                               !/대학|협회|연맹|학원|병원|의원/.test(issuer);
                               
      if (!hasReport && !isSchoolOfficial) {
        if (!splitItem.ruleCheck) {
          splitItem.ruleCheck = "체크필요 [출결신고서 누락]";
        } else if (!splitItem.ruleCheck.includes("출결신고서 누락")) {
          splitItem.ruleCheck = splitItem.ruleCheck + " ⚠️ 체크필요 [출결신고서 누락]";
        }
      }

      if (splitItem.ruleCheck) {
        splitItem.ruleCheck = splitItem.ruleCheck.replace(/⚠️ 보완 필요: (학생 |학부모 )?(이름|성명|학번) 불일치/g, "")
                                       .replace(/⚠️ 보완 필요 여부 검토: 일수 기재 누락 여부에 대한 재검토 추천/g, "")
                                       .replace(/⚠️ 보완 필요: 일수 기재 누락/g, "")
                                       .trim();
        if (splitItem.ruleCheck === "규정 준수" || splitItem.ruleCheck === "✅ 정상" || splitItem.ruleCheck === "체크필요" || splitItem.ruleCheck === "체크필요:") {
          splitItem.ruleCheck = "";
        } else if (splitItem.ruleCheck && !splitItem.ruleCheck.startsWith("⚠️") && (splitItem.ruleCheck.startsWith("오류") || splitItem.ruleCheck.startsWith("보완") || splitItem.ruleCheck.startsWith("체크필요"))) {
          splitItem.ruleCheck = "⚠️ " + splitItem.ruleCheck;
        }
      }

      splitItem.fileUrl = fileUrl;
      splitItem.filename = filename;
      
      delete splitItem.studentInfo; // Remove temporary property
      resultWithUrl.push(splitItem);
    });
    
    // 4. Auto-folder Organization (Move file based on first student's info if multiple)
    if (resultWithUrl.length > 0) {
      organizeFile(file, resultWithUrl[0]);
    }
    
    return {
      success: true,
      fileUrl: fileUrl,
      data: resultWithUrl
    };
  } catch (e) {
    Logger.log(e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Splits an attendance item into multiple segments if its date range crosses month boundaries.
 */
function splitMultiMonthItem(item) {
  const periodStart = item.periodStart || item.date || "";
  const periodEnd = item.periodEnd || "";
  
  let daysNum = 1;
  const daysRaw = item.totalDays || "1일간";
  const daysMatch = daysRaw.match(/(\d+)/);
  if (daysMatch) {
    daysNum = parseInt(daysMatch[1], 10);
  }
  
  if (!periodStart) return [item];
  
  const parseDateTime = (str) => {
    if (!str) return { date: "", time: "" };
    const match = str.match(/^(\d{4}-\d{2}-\d{2})(.*)$/);
    if (match) {
      return { date: match[1], time: match[2].trim() };
    }
    return { date: str, time: "" };
  };
  
  const startInfo = parseDateTime(periodStart);
  const endInfo = parseDateTime(periodEnd);
  
  if (!startInfo.date) return [item];
  
  let actualEndStr = endInfo.date;
  if (!actualEndStr) {
    try {
      const parts = startInfo.date.split('-');
      if (parts.length === 3) {
        const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
        let current = new Date(startDateObj);
        let addedDays = 1;
        const holidays = getHolidaysSet(current.getFullYear(), current.getFullYear() + 1);
        const schoolHolidays = getDiscretionaryHolidays();
        schoolHolidays.forEach(h => holidays.add(h));
        
        // 주말과 공휴일을 제외한 수업일 기준으로 종료일 계산
        while (addedDays < daysNum) {
          current.setDate(current.getDate() + 1);
          if (!isHolidayOrWeekend(current, holidays)) {
            addedDays++;
          }
        }
        actualEndStr = Utilities.formatDate(current, "GMT+9", "yyyy-MM-dd");
      }
    } catch (e) {}
  }
  
  if (!actualEndStr) return [item];
  
  try {
    const startParts = startInfo.date.split('-');
    const endParts = actualEndStr.split('-');
    if (startParts.length !== 3 || endParts.length !== 3) return [item];
    
    const startDateObj = new Date(parseInt(startParts[0], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[2], 10), 12, 0, 0);
    const endDateObj = new Date(parseInt(endParts[0], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[2], 10), 12, 0, 0);
    
    if (isNaN(startDateObj.getTime()) || isNaN(endDateObj.getTime()) || startDateObj.getTime() > endDateObj.getTime()) {
      return [item];
    }
    
    // Check if same month
    if (startDateObj.getFullYear() === endDateObj.getFullYear() && startDateObj.getMonth() === endDateObj.getMonth()) {
      return [item];
    }
    
    // Generate all dates (only school days)
    const allDates = [];
    let currentObj = new Date(startDateObj);
    const holidays = getHolidaysSet(startDateObj.getFullYear() - 1, endDateObj.getFullYear() + 1);
    const schoolHolidays = getDiscretionaryHolidays();
    schoolHolidays.forEach(h => holidays.add(h));
    
    let maxIterations = 365;
    while (currentObj <= endDateObj && maxIterations > 0) {
      if (!isHolidayOrWeekend(currentObj, holidays)) {
        allDates.push(Utilities.formatDate(currentObj, "GMT+9", "yyyy-MM-dd"));
      }
      currentObj.setDate(currentObj.getDate() + 1);
      maxIterations--;
    }
    
    const groups = {};
    allDates.forEach(dateStr => {
      const ym = dateStr.substring(0, 7);
      if (!groups[ym]) groups[ym] = [];
      groups[ym].push(dateStr);
    });
    
    const ymKeys = Object.keys(groups).sort();
    if (ymKeys.length <= 1) return [item];
    
    const splitItems = [];
    ymKeys.forEach((ym, index) => {
      const groupDates = groups[ym];
      const segStartDate = groupDates[0];
      const segEndDate = groupDates[groupDates.length - 1];
      const segDays = groupDates.length;
      
      const newItem = JSON.parse(JSON.stringify(item));
      
      newItem.date = segStartDate;
      newItem.periodStart = segStartDate + (index === 0 && startInfo.time ? " " + startInfo.time : "");
      newItem.periodEnd = segEndDate + (index === ymKeys.length - 1 && endInfo.time ? " " + endInfo.time : "");
      newItem.totalDays = segDays + "일간";
      
      splitItems.push(newItem);
    });
    
    return splitItems;
  } catch (e) {
    return [item];
  }
}

/**
 * Helper: Get Target Folder (By ID or Spreadsheet Parent)
 */
function getTargetFolder() {
  const folderId = CONFIG.UPLOAD_FOLDER_ID;
  if (folderId && folderId.trim().length > 0) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      Logger.log("경고: 폴더 ID가 잘못되었습니다. 시트 상위 폴더를 사용합니다.");
    }
  }
  
  // Fallback: Current Spreadsheet's Parent Folder
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  if (parents.hasNext()) return parents.next();
  
  return DriveApp.getRootFolder(); // Last resort
}

/**
 * Call Gemini API with Retry and Logging
 */
function callGeminiVision(base64Content, mimeType, filename, studentNames = []) {
  const apiKey = CONFIG.GEMINI_API_KEY;
  if (!apiKey) throw new Error("API Key is missing in Constants.gs");

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;
  
  let prompt = `
    신고서 및 첨부서류를 분석하여 정보를 추출하고 규정 준수 여부를 검토하라.
    모든 응답은 반드시 지정된 JSON 구조의 배열(ARRAY)로 반환되어야 한다.

    [추출 가이드]
    - **category (대분류):** 반드시 '질병', '출석인정', '기타', '미인정' 중 정확히 하나만 선택하여 기입하라.
      - **[최우선 규칙] 체크된 분류 우선 반영**: 문서 내의 출결 종류 선택란(체크박스 등)에서 '질병', '출석인정' 등의 항목에 명시적으로 체크(V) 또는 표시가 되어 있다면, 작성된 사유의 단어와 무관하게 **반드시 체크된 분류를 대분류로 설정**해야 한다.
      - 예시: 출결 종류 체크박스 중 '질병'에 체크가 되어 있고, 사유에는 '생리통으로 인한 병원 진료'라고 적혀 있다면 대분류(category)는 반드시 **'질병'**으로 판정해야 한다. 생리통이라는 단어가 있다고 해서 마음대로 '출석인정'으로 바꿔 판정해서는 안 된다.
      - '출석인정': 출결 종류 항목에서 '출석인정'에 체크되어 있거나, 생리통(단, 질병에 체크되지 않고 생리통만으로 단독 결석계를 제출한 경우), 학교장허가체험학습, 전염병/감염병(독감, 인플루엔자, 코로나19 등 격리가 필요한 병명), 경조사, 공가 등의 사유일 때 선택하라.
      - '질병': 출결 종류 항목에서 '질병'에 체크되어 있거나, 감기, 몸살, 장염, 두통, 복통, 치과 치료 등 일반적인 질병이나 통원 치료일 때 선택하라.
      - '미인정': 태만, 가출, 고의적 출석 거부 등 정당한 사유가 없을 때 선택하라.
      - 절대 '질병결석'이나 '출석인정(생리통)' 처럼 소분류나 상세사유를 결합해 기입하지 마라.
    - **subCategory (소분류):** 반드시 '결석', '지각', '조퇴', '결과' 중 정확히 하나만 선택하여 기입하라. 절대 '생리통', '경조사', '깸석' 등을 기입하지 마라. 생리통이나 경조사 같은 상세 사유는 반드시 사유상세(reasonDetail) 필드에만 기입해야 한다. (주의: 문서 제목이 '결석계'이더라도 본문 내부의 선택 항목(체크박스 등)에서 '지각', '조퇴', '결과' 등에 체크되어 있거나 사유상세에 지각/조퇴가 명시되어 있다면 해당 소분류로 정확하게 분류하여 추출하라. 결석이 아닌 건을 결석으로 오분류하지 않도록 정밀히 판별하라.)
    - **서명(학생/학부모/담임):** 각 서명란(주로 이름 옆이나 '(서명 또는 인)', '(인)', '(서명)' 표시가 있는 곳)을 정밀하게 확인하여 물리적인 서명(수필 서명, 사인, 흘려 쓴 이름, 도장/날인 등)이 존재하면 "있음", 서명란이 비어 있거나 인쇄된 성명 텍스트 외에 친필 서명/사인이 없는 경우 "없음"으로 매우 엄격하게 판별하여 추출하라. 학생서명, 학부모서명, 담임서명 모두 동일하게 판별한다.
    - **학번 추출 규칙:** 5자리 학번의 첫 번째 숫자는 항상 1~6 사이이다. 절대 0으로 시작할 수 없다. 만약 0으로 보인다면 보정하여 추출하라.
    - **이름/학번/학부모명 불일치 판정 금지:** 이미지 판독 결과가 명부와 다르거나 판독이 어렵더라도 성함/학번 불일치와 관련된 보완 필요 경고를 절대 생성하지 마라.
    - **기간 (시작일시/종료일시):** 결석계 폼 내부에 명시된 '결석/지각/조퇴/결과 시작일'과 '종료일'을 찾아 "YYYY-MM-DD X교시" 형식으로 추출하라. (작성일이나 첨부서류 발급일이 아님). 기재되어 있지 않으면 "×"로 추출하라.
    - **일수:** 숫자 1을 괄호와 함께 흘려 쓰거나('|', 'I', 'l', '/', '(' 등), 아예 비워두는 경우가 많습니다. '0일간'이거나 형태가 불분명하면 무조건 "1일간"으로 추출하고, 그 외 숫자가 명확하면 "X일간"으로 추출하라. 괄호()는 빼고 숫자와 글자만 추출하라.
    - **사유상세:** 대분류가 질병이나 기타인 경우 학생이 기록한 구체적 사유를, 출석인정인 경우 생리통, 학교장허가체험학습, 전염병 또는 구체적 사유를 그대로 추출하라.
    - **첨부서류:** PDF의 모든 페이지를 정밀하게 분석하여, 신고서 양식(폼) 이외에 실제로 추가 촬영/스캔되어 첨부된 증빙 서류(예: 진료확인서, 처방전, 약 봉투 등)가 존재하는지 철저히 확인하라.
      - [주의] 신고서 양식 하단이나 본문에 인쇄된 '첨부서류 안내 문구(예: 진료확인서, 처방전 등을 제출 바람)'는 단순 안내일 뿐 실제 첨부된 서류가 아니므로 절대 추출해선 안 된다.
      - 실제 추가된 서류 이미지가 발견된 경우에만 해당 서류의 명칭을 하나 추출하고, 실제 첨부 서류가 없다면 반드시 "×"로 추출하라.
    - **통원일(docStartDate)/종료일(docEndDate):** 
      - 질병 관련 병원/약국 서류의 경우: 서류에 명시된 '통원 기간' 또는 '진료 기간'의 시작일이 발급일과 다르다면 그 시작일을 통원일로, 종료일이 있다면 통원종료일로 추출하라. 기간 표시 없이 발급일만 있다면 발급일을 통원일로 추출하고 통원종료일은 빈칸("")으로 추출하라.
      - 경조사 관련 서류(사망진단서, 청첩장 등)의 경우: 경조사 발생일(사망일, 결혼식일, 입양일 등)을 통원일(docStartDate)로 추출하고, 통원종료일(docEndDate)은 빈칸("")으로 추출하라.
    - **발급처 (issuer):** 서류를 발급한 기관명(병원, 약국 등)을 추출하라. 단, 학부모 확인서는 "학부모", 담임 확인서는 "담임"이라고 추출하라.
    - **규정위반여부 (ruleCheck):** 서류가 완벽히 규정을 준수하면 빈칸("")으로 반환하라. 위반사항이 있다면 "체크필요: [짧은 사유]" 형식으로 반환하라. 단, 서류 제출 지연(제출 기한 도과)에 대한 판정은 절대 수행하지 마라.
      - **질병 결석인 경우**:
        - 병원/약국 서류 외에도 '보건실 입실 확인증(보건실 방문 확인서)'이 첨부되어 있다면 병원 서류와 동급의 정상 서류로 인정하여 규정 준수(빈칸 "")로 반환하라.
        - 만약 병원 서류(또는 보건실 확인서)가 없고 학부모 확인서나 담임 확인서만 첨부된 경우, 시스템에서 전날 병원진료 기록을 조회하여 보정해야 하므로 반드시 **"질병 대체 서류 검증 필요"**라고 반환하라. (바로 오류 경고를 내지 말 것)
    - **NEIS 문구 (neisString):** 반드시 "[일자] [대분류][소분류]([사유상세])" 형태로 생성하라.
      - 예시: 2026-03-10 질병결석(복통)
      - 만약 사유상세(reasonDetail)가 없거나 '없음', '해당없음'인 경우 괄호와 사유를 생략하고 "[일자] [대분류][소분류]"로만 생성하라. (예: 2026-03-10 질병결석)
      - 일자는 date 필드값(YYYY-MM-DD)을 활용하며, 대분류와 소분류는 띄어쓰기 없이 붙여서 기입하라.
    - **수정 및 정정 표시 처리:**
      - 문서 본문에 두 줄을 긋고 수정도장(보통 빨간색/파란색 원형 또는 사각형 개인 도장)을 찍거나 글씨를 덧써서 수정(정정)한 경우, 두 줄로 지워진 과거의 텍스트나 항목은 완전히 무시하고 **최종 수정(정정)된 올바른 텍스트 및 표시만을 기준으로 정보를 추출**하라.
      - **도장 겹침 오판 주의:** 수정도장이 체크박스(예: '지각', '조퇴' 등)나 특정 글자 위에 찍혀 있는 경우, 도장의 붉은색/푸른색 잉크나 테두리 선이 진하게 표시되어 마치 해당 체크박스에 체크 표시(V)가 된 것처럼 AI가 오인하기 쉽다. 수정도장의 잉크 자국이나 날인 형태는 체크 표시가 아님을 인지하고, 도장 겹침으로 인해 진하게 보이는 부분에 현혹되지 마라. 실제 손글씨로 직접 체크(V)되거나 동그라미 쳐진 최종 선택 항목(비록 더 연하게 표시되어 있더라도)을 신중하고 정확하게 판별하여 추출하라.
    - **hasReportCard (출결신고서 포함 여부):** PDF 전체(특히 해당 학생의 서류 영역)에 결석계, 결석신고서, 지각·조퇴·결과 신고서 등의 출결신고서 양식이 포함되어 있는 경우 true, 포함되어 있지 않고 증빙서류(진료확인서, 처방전 등)만 있는 경우 false로 판별하여 반환하라.
    - **isOfficialDocument (정식 공문 여부):** PDF 내에 공무 수행, 대회 참가 등의 증빙을 위해 학교나 기관 등에서 발행한 '공문'(수신/발신처 및 관인/직인이 날인된 공식 문서 포맷)이 포함되어 있는 경우 true, 그렇지 않으면 false로 판별하여 반환하라.

    [복수 학생 서류 병합 파일 처리 지침]
    - **대상 상황:** 하나의 PDF 파일 내에 여러 학생의 서류가 합쳐져서 업로드되는 경우 이 지침을 적용한다.
    - **학생별 독립 인지:** 문서 내의 모든 페이지를 정밀 분석하여, 각 페이지마다 다르게 나타나는 학번과 학생 이름을 가장 먼저 파악하라.
    - **개별 레코드 분리:** 학번이나 이름이 서로 다르다면 완전히 별개의 학생 결석/출결 건으로 간주하여, 각각 독립된 JSON 객체(레코드)로 분리하고 배열에 포함하라.
    - **데이터 격리:** 앞 페이지에 나온 특정 학생의 신고서 정보(시작일시, 종료일시, 구체적 사유 등)를 뒤 페이지에 등장하는 다른 학생의 증빙 서류에 결합하거나 적용(상속)하여 오판하지 않도록 철저히 격리하라.
    - **신고서 없는 학생 처리:** 뒤 페이지 학생은 신고서 양식이 없으므로, 해당 학생의 JSON 객체에는 날짜/일수 등을 증빙서류 기준으로만 추출하고, 결석계 양식 관련 필드('studentSigned', 'parentName', 'parentSigned' 등)는 "없음"으로 처리하고, 'hasReportCard'는 반드시 'false'로 설정하라.

    [복수 출결 상황(복수 체크) 발생 시 처리 지침]
    - **대상 상황:** 신고서 상단의 출결 구분(결석, 지각, 조퇴, 결과) 중 2개 이상의 항목에 체크(표시)가 되어 있고, 수정도장이나 두 줄 긋기 등의 정정 흔적이 없는 경우(즉, 한 개의 결석계로 여러 종류의 출결 상황을 동시에 신고한 경우) 이 지침을 적용한다.
    - **사유 및 확인 내용 분석:** 작성된 사유나 담임 확인 의견서의 구체적인 본문 내용(예: "두통으로 26일 조퇴하였으나, 27일 호전되지 않아 처방약 복용 및 요양함")을 정밀하게 읽어내어, 날짜별로 발생한 구체적인 출결 상황을 파악하라.
    - **개별 레코드 분리:** 판별한 각 날짜별 출결 종류(소분류)에 맞추어 **각각 개별적인 JSON 객체(레코드)를 생성**하여 배열에 포함하라.
      - 예시: 26일 조퇴 후 27일 결석을 한 장의 서류에 신고한 경우:
        - 객체 1: date="[해당연도]-MM-26", subCategory="조퇴", reasonDetail="두통으로 인한 조퇴", totalDays="1일간"
        - 객체 2: date="[해당연도]-MM-27", subCategory="결석", reasonDetail="두통으로 인한 요양 및 결석", totalDays="1일간"
      - 만약 단순 나열식으로 "26일 지각, 27일 조퇴"와 같이 적혀 있다면, 이에 맞추어 각각의 날짜와 해당 소분류(지각 / 조퇴)를 갖는 JSON 객체들을 각각 나누어 리턴하라.

    [공문 및 목록/표 형태의 서류 처리 지침]
    - **대상 문서:** 단일 결석계 양식이 아니라, 공문(공식 문서), 참가자 명단, 월간 근무상황표, 출석표, 활동일지 등 목록이나 표 형태의 서류인 경우 이 지침을 적용한다.
    - **개별 레코드 분리:** 대상 학생 명단에 포함된 학생의 이름을 본 문서(예: 표의 참여자/근무자/참가자 열 등)에서 찾은 뒤, 해당 학생이 참여/활동한 **각각의 날짜(일자)별로 개별적인 JSON 레코드(객체)**를 생성하여 배열에 포함하라.
      - 예: 학생 '장명준'이 3월 5일, 3월 10일, 3월 12일에 참여한 기록이 표에 있다면, '장명준'에 대해 날짜가 각각 2026-03-05, 2026-03-10, 2026-03-12인 총 3개의 JSON 객체를 각각 생성하여 배열에 담아야 한다.
    - **날짜 필드 지정:** 'date'와 'periodStart', 'periodEnd' 필드는 모두 해당 활동이 발생한 날짜(형식: "YYYY-MM-DD")로 설정하라. 연도가 명시되지 않았다면 문서의 발급일/작성일의 연도(예: 2026년)를 기준으로 채워라.
    - **출결 구분 기본값:** 목록/표에서 해당 활동 참여가 지각, 조퇴, 결과, 결석 중 어느 것인지 명시되어 있지 않거나 단순히 '참여', '근무', '출근' 등으로 되어 있어 모호한 경우, **대분류(category)는 "출석인정", 소분류(subCategory)는 "결석"으로 기본 설정**하라. (단, 문서에 지각, 조퇴, 결과가 명확히 표시된 경우는 해당 값을 사용하라.)
    - **서명 및 증빙 기본값:** 표/목록 형태 서류의 특성상 개별 학생/학부모 친필 서명은 없으므로, 불필요한 서류 미비 경고가 발생하지 않도록 **'studentSigned', 'parentSigned', 'teacherSigned' 필드는 모두 "있음"**으로 설정하라.
    - **사유 상세 및 첨부서류:** 'reasonDetail'에는 해당 행의 구체적인 활동 내용 또는 업무내용(예: "장애인 맞춤형 일자리 오리엔테이션", "커피바리스타 1강" 등)을 추출하여 기입하라. 'attachments' 필드는 "공문" 또는 해당 서류의 명칭(예: "근무상황표")을 적어라. 'totalDays'는 "1일간"으로 설정하고, 'issuer'는 문서의 확인자/담당자 성함(예: "지예나") 또는 발급 기관명을 적어라.
  `;

  if (studentNames && studentNames.length > 0) {
    prompt += `
    \n[대상 학생 명단]
    현재 분석 대상 학년의 학생 명단은 다음과 같다: ${studentNames.join(', ')}
    문서 내에서 참여자 또는 대상자 목록을 확인할 때, 반드시 이 명단에 포함된 학생의 정보만 추출하라. 명단에 없는 이름(예: 확인자/검토자 '지예나' 등)은 절대로 학생 이름으로 추출해선 안 된다.
    `;
  }

  const payload = {
    contents: [{
      parts: [
        { inlineData: { mimeType: mimeType, data: base64Content } },
        { text: String(prompt).trim() }
      ]
    }],
    generationConfig: {
      response_mime_type: "application/json",
      response_schema: {
        type: "ARRAY",
        items: {
          type: "OBJECT",
          properties: {
            date: { type: "STRING" },
            studentId: { type: "STRING" },
            name: { type: "STRING" },
            studentSigned: { type: "STRING" },
            category: { type: "STRING" },
            subCategory: { type: "STRING" },
            periodStart: { type: "STRING" },
            periodEnd: { type: "STRING" },
            totalDays: { type: "STRING" },
            reasonDetail: { type: "STRING" },
            parentName: { type: "STRING" },
            parentSigned: { type: "STRING" },
            teacherName: { type: "STRING" },
            teacherSigned: { type: "STRING" },
            attachments: { type: "STRING" },
            docStartDate: { type: "STRING" },
            docEndDate: { type: "STRING" },
            issuer: { type: "STRING" },
            ruleCheck: { type: "STRING" },
            neisString: { type: "STRING" },
            hasReportCard: { type: "BOOLEAN" },
            isOfficialDocument: { type: "BOOLEAN" }
          },
          required: ["date", "studentId", "name", "studentSigned", "category", "subCategory", "periodStart", "periodEnd", "totalDays", "reasonDetail", "parentName", "parentSigned", "teacherName", "teacherSigned", "attachments", "docStartDate", "docEndDate", "issuer", "ruleCheck", "neisString", "hasReportCard", "isOfficialDocument"]
        }
      }
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = fetchWithRetry(url, options, filename);
  const responseCode = response.getResponseCode();
  const contentText = response.getContentText();
  
  if (responseCode !== 200) {
    let errorMsg = contentText;
    try {
      const errorObj = JSON.parse(contentText);
      errorMsg = errorObj.error ? errorObj.error.message : contentText;
    } catch (e) {}
    logToSystem(filename, "ERROR " + responseCode, errorMsg);
    throw new Error(errorMsg);
  }

  const json = JSON.parse(contentText);
  let resultText = json.candidates[0].content.parts[0].text;
  
  logToSystem(filename, "SUCCESS", "분석 완료");

  try {
    // Schema enforcement usually returns clean JSON, but just in case:
    const jsonMatch = resultText.match(/\[[\s\S]*\]/);
    const finalJson = jsonMatch ? jsonMatch[0] : (resultText.startsWith('{') ? `[${resultText}]` : resultText);
    return JSON.parse(finalJson);
  } catch (e) {
    throw new Error("AI 데이터 파싱 실패: " + e.message + "\nRaw Content: " + resultText);
  }
}

/**
 * Fetch with Exponential Backoff Retry
 */
function fetchWithRetry(url, options, filename, maxRetries = 3) {
  let retryCount = 0;
  let waitTime = 2000; // Start with 2 seconds

  while (retryCount <= maxRetries) {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    // If success or non-retryable error (not 429 or 5xx), return
    if (code === 200 || (code !== 429 && code < 500)) {
      return response;
    }

    // If 429 or 5xx, retry with delay
    if (retryCount < maxRetries) {
      Logger.log(`Retry ${retryCount + 1}/${maxRetries} for ${filename} after ${waitTime}ms (Status: ${code})`);
      logToSystem(filename, "RETRY " + code, `${retryCount + 1}차 재시도 대기중... (${waitTime}ms)`);
      
      Utilities.sleep(waitTime);
      retryCount++;
      waitTime *= 2; // Exponential backoff
    } else {
      return response; // Final failure
    }
  }
}

/**
 * Process NEIS Data (Array of Arrays)
 */
function processNeisData(month, data, filename) {
  try {
    // data is already a 2D array parsed by frontend SheetJS
    let globalGrade = "";
    let globalClass = "";
    
    for (let i = 0; i < Math.min(15, data.length); i++) {
       const rowStr = (data[i] || []).join(" ");
       const match = rowStr.match(/(\d+)학년\s*(\d+)반/);
       if (match) {
         globalGrade = match[1];
         globalClass = match[2].padStart(2, '0');
         break;
       }
    }

    const settings = getSystemSettings();
    if (globalGrade && String(globalGrade) !== String(settings.grade)) {
      throw new Error(`이 시스템은 ${settings.grade}학년만 관리하도록 설정되어 있습니다. 업로드한 파일은 ${globalGrade}학년 데이터입니다.`);
    }

    const students = [];
    let isDataRowStarted = false;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // Skip empty or short rows
      if (!row || row.length < 5) continue;
      
      const seqRaw = String(row[0]).trim();
      
      // Look for a row that starts with a number (SEQ) to identify data rows
      if (/^\d+$/.test(seqRaw)) {
        isDataRowStarted = true;
        const name = String(row[1]).trim();
        if (!name) continue; 
        
        const sNum = seqRaw.padStart(2, '0');
        let stId = "";
        if (globalGrade && globalClass) {
          stId = globalGrade + globalClass + sNum;
        } else {
          stId = seqRaw; // fallback
        }
        
        const parseStat = (val) => {
          const parsed = parseInt(val, 10);
          return isNaN(parsed) ? 0 : parsed;
        };

        students.push({
          seq: seqRaw,
          studentId: stId,
          name: name,
          absent: [parseStat(row[3]), parseStat(row[4]), parseStat(row[5]), parseStat(row[6])],
          tardy: [parseStat(row[7]), parseStat(row[8]), parseStat(row[9]), parseStat(row[10])],
          early: [parseStat(row[11]), parseStat(row[12]), parseStat(row[13]), parseStat(row[14])],
          skipped: [parseStat(row[15]), parseStat(row[16]), parseStat(row[17]), parseStat(row[18])],
          stats: [parseStat(row[19]), parseStat(row[20]), parseStat(row[21]), parseStat(row[22])]
        });
      }
    }

    // Sort students by studentId in ascending order
    students.sort((a, b) => {
      const aId = parseInt(a.studentId, 10);
      const bId = parseInt(b.studentId, 10);
      if (!isNaN(aId) && !isNaN(bId)) {
        return aId - bId;
      }
      return a.studentId.localeCompare(b.studentId);
    });

    if (students.length === 0) {
      logToSystem(filename, "EMPTY_DATA", "조건에 부합하는 학생(출결기록 발생)이 없거나 파싱할 데이터가 없습니다.");
      return { success: true, month: month, count: 0 };
    }

    saveNeisStatsSheet(month, students);
    logToSystem(filename, "SUCCESS", `${month}월 데이터 ${students.length}건 저장 완료`);
    
    // Backup is handled at frontend or ignored because data is just extracted

    return { success: true, month: month, count: students.length };
  } catch(e) {
    Logger.log(e.toString());
    logToSystem(filename, "ERROR", e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Save Parsed NEIS Statistics to NeisData sheet
 */
function saveNeisStatsSheet(month, students) {
  if (!students || students.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  if (!sheet) throw new Error("NEIS 시트가 존재하지 않습니다.");
  
  const safeGet = (arr, idx) => (arr && Array.isArray(arr) && arr.length > idx) ? arr[idx] : 0;
  
  // 1. Prepare new rows to write
  const newRows = students.map(st => [
    st.seq || '',
    month || '',
    st.studentId || '',
    st.name || '',
    safeGet(st.absent, 0), safeGet(st.absent, 1), safeGet(st.absent, 2), safeGet(st.absent, 3),
    safeGet(st.tardy, 0), safeGet(st.tardy, 1), safeGet(st.tardy, 2), safeGet(st.tardy, 3),
    safeGet(st.early, 0), safeGet(st.early, 1), safeGet(st.early, 2), safeGet(st.early, 3),
    safeGet(st.skipped, 0), safeGet(st.skipped, 1), safeGet(st.skipped, 2), safeGet(st.skipped, 3),
    safeGet(st.stats, 0), safeGet(st.stats, 1), safeGet(st.stats, 2), safeGet(st.stats, 3)
  ]);

  // 2. Identify the target class from studentId (e.g. "30101" -> "301")
  const firstStudentId = String(students[0].studentId || "");
  if (!firstStudentId) {
    // If studentId format is empty, fallback to simple append to avoid data loss
    const startRow = sheet.getLastRow() + 1;
    const targetRange = sheet.getRange(startRow, 1, newRows.length, 24);
    targetRange.setValues(newRows);
    reconcileNeisWithAttendance(month, "all");
    return;
  }

  const targetClassPrefix = firstStudentId.substring(0, Math.max(1, firstStudentId.length - 2));
  const targetMonthStr = String(month).trim();

  // 3. Read existing data and filter out duplicate month & class records
  const lastRow = sheet.getLastRow();
  let mergedData = [];
  
  if (lastRow > 1) {
    const existingData = sheet.getRange(2, 1, lastRow - 1, 24).getValues();
    mergedData = existingData.filter(row => {
      const rowMonth = String(row[1]).trim();
      const rowStudentId = String(row[2]).trim();
      const rowClassPrefix = rowStudentId.substring(0, Math.max(1, rowStudentId.length - 2));
      
      // Filter out (delete) records matching the same month and same class prefix
      const isTargetDuplicate = (rowMonth === targetMonthStr && rowClassPrefix === targetClassPrefix);
      return !isTargetDuplicate;
    });
  }

  // 4. Append new rows to filtered list
  mergedData = mergedData.concat(newRows);

  // 5. Clear old contents and write merged dataset
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 24).clearContent().setBackground(null);
  }
  
  const targetRange = sheet.getRange(2, 1, mergedData.length, 24);
  targetRange.setValues(mergedData);

  // 6. Generate and apply backgrounds
  const bgColors = mergedData.map(row => {
    const bgRow = Array(24).fill(null);
    let hasNonZero = false;
    for (let c = 4; c < 24; c++) {
      if (typeof row[c] === 'number' && row[c] > 0) {
        hasNonZero = true;
      }
    }
    
    if (hasNonZero) {
      for (let c = 0; c < 4; c++) bgRow[c] = '#fff2cc';
      for (let c = 4; c < 24; c++) {
        if (typeof row[c] === 'number' && row[c] > 0) {
          bgRow[c] = '#fff2cc';
        }
      }
    }
    return bgRow;
  });
  
  targetRange.setBackgrounds(bgColors);
  
  // 7. Auto-reconcile and update AttendanceData
  reconcileNeisWithAttendance(month, "all");
}

/**
 * Reconcile NeisData sheet with AttendanceDB records and write results to AttendanceData
 */
function reconcileNeisWithAttendance(targetMonth, targetClass) {
  targetMonth = targetMonth || "all";
  targetClass = targetClass || "all";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  let dataSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
  
  if (!attSheet || !neisSheet) return "필수 시트가 존재하지 않습니다.";
  
  if (!dataSheet) {
    dataSheet = ss.insertSheet(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
    const dataHeaders = ['순번', '월', '학번', '이름', '결석_질병', '결석_미인정', '결석_기타', '결석_인정', '지각_질병', '지각_미인정', '지각_기타', '지각_인정', '조퇴_질병', '조퇴_미인정', '조퇴_기타', '조퇴_인정', '결과_질병', '결과_미인정', '결과_기타', '결과_인정', '결석통계', '지각통계', '조퇴통계', '결과통계', '일치여부', '불일치내용'];
    dataSheet.appendRow(dataHeaders);
    dataSheet.getRange(1, 1, 1, dataHeaders.length).setBackground('#e67e22').setFontColor('white').setFontWeight('bold');
  }

  const settings = getSystemSettings();
  const examPeriods = getExamPeriods();

  const attData = attSheet.getDataRange().getValues().slice(1);
  const neisRows = neisSheet.getDataRange().getValues().slice(1);
  
  // 1. Aggregate AttendanceDB by StudentId_Month
  // Map<SID_Month, { stats: { "Category_SubCategory": Value }, processedFiles: Set }>
  const aggMap = new Map();
  
  attData.forEach((row, rowIndex) => {
    // Determine the start date of the period from G열 (시작일시) first, with fallback to A열 (일자)
    const dateValRaw = String(row[6] || "").trim();
    let dateVal = "";
    if (dateValRaw && dateValRaw.includes('-')) {
      dateVal = dateValRaw.split(' ')[0];
    }
    if (!dateVal || dateVal.split('-').length !== 3) {
      dateVal = row[0];
    }
    
    const date = parseDateSafe(dateVal);
    if (!date || isNaN(date.getTime())) return;
    
    const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
    const sid = String(row[1]).trim();
    const gradeNum = sid.length >= 3 ? parseInt(sid.charAt(0), 10) : -1;
    const cNum = sid.length >= 3 ? parseInt(sid.substring(1,3), 10) : -1;
    if (gradeNum !== settings.grade || cNum > settings.classes) return;
    const catRaw = String(row[5] || "").trim();      // F열: 소분류 (결석, 지각, 조퇴, 결과) -> NEIS category
    const subCatRaw = String(row[4] || "").trim();   // E열: 대분류 (질병, 출석인정, 기타, 미인정) -> NEIS subcategory
    const daysRaw = String(row[8] || "1").trim();
    
    // Normalize subCategory labels to match NEIS (질병, 미인정, 기타, 인정)
    let subCat = "기타";
    if (subCatRaw.includes("질병")) subCat = "질병";
    else if (subCatRaw.includes("미인정")) subCat = "미인정";
    else if (subCatRaw.includes("기타")) subCat = "기타";
    else if (subCatRaw.includes("출석인정") || subCatRaw.includes("인정")) subCat = "인정";

    let cat = "결석";
    if (catRaw.includes("결석")) cat = "결석";
    else if (catRaw.includes("지각")) cat = "지각";
    else if (catRaw.includes("조퇴")) cat = "조퇴";
    else if (catRaw.includes("결과")) cat = "결과";

    // Parse days: "3일간" -> 3
    let days = 1;
    const match = daysRaw.match(/(\d+)/);
    if (match) days = parseInt(match[1], 10);
    
    const reasonDetail = String(row[9] || "").trim();
    const isExperiential = reasonDetail.includes("학교장허가체험학습") || 
                           reasonDetail.includes("현장체험") || 
                           reasonDetail.includes("교외체험") || 
                           reasonDetail.includes("체험학습");

    // Calculate covered dates range
    const coveredDates = [];
    try {
      const parts = dateStr.split('-');
      if (parts.length === 3) {
        const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
        for (let d = 0; d < days; d++) {
          const currentDateObj = new Date(startDateObj.getTime() + d * 24 * 60 * 60 * 1000);
          coveredDates.push(Utilities.formatDate(currentDateObj, "GMT+9", "yyyy-MM-dd"));
        }
      } else {
        coveredDates.push(dateStr);
      }
    } catch (e) {
      coveredDates.push(dateStr);
    }

    const statKey = `${cat}_${subCat}`;

    // Distribute each covered date to its respective calendar month
    coveredDates.forEach(dStr => {
      const parts = dStr.split('-');
      if (parts.length !== 3) return;
      const actualMonth = parseInt(parts[1], 10);
      
      const key = `${sid}_${actualMonth}`;
      if (!aggMap.has(key)) {
        aggMap.set(key, {
          stats: {},
          processedDates: {},
          experientialRowNums: new Set()
        });
      }
      const studentData = aggMap.get(key);
      const studentStats = studentData.stats;

      // Check if this specific date is an exam date
      const isExamDate = examPeriods.some(period => {
        if (!period.startStr) return false;
        return dStr >= period.startStr && dStr <= period.endStr;
      });

      // 생리통 예외: 생리통(출석인정)이고, 해당 날짜가 시험기간이 아닌 경우에만 서류 제출 의무가 없으므로 로컬 서류 건수 계산에서 제외
      const isMenstrual = subCat === "인정" && /생리/.test(reasonDetail) && !isExamDate;

      // Handle experiential learning document count per month (prevent counting multiple days in the same document as separate docs)
      if (subCat === "인정" && isExperiential) {
        if (!studentData.experientialRowNums.has(rowIndex)) {
          studentData.experientialRowNums.add(rowIndex);
          studentStats["체험학습_건수"] = (studentStats["체험학습_건수"] || 0) + 1;
        }
      }

      if (cat === "결석") {
        if (!studentData.processedDates[statKey]) {
          studentData.processedDates[statKey] = new Set();
        }
        if (!studentData.processedDates[statKey].has(dStr)) {
          studentData.processedDates[statKey].add(dStr);
          // 생리통 여부와 관계없이 로컬 서류 건수에 합산 (사용자 화면에 정상 표시되도록)
          studentStats[statKey] = (studentStats[statKey] || 0) + 1;
          
          if (isMenstrual) {
            studentStats[statKey + "_menstrual"] = (studentStats[statKey + "_menstrual"] || 0) + 1;
          }
        }
      } else {
        // For 지각/조퇴/결과, count unique days
        // 생리통 여부와 관계없이 로컬 서류 건수에 합산
        if (!studentStats[statKey]) studentStats[statKey] = new Set();
        if (studentStats[statKey] instanceof Set) {
          studentStats[statKey].add(dStr);
        }
        
        if (isMenstrual) {
          const mKey = statKey + "_menstrual";
          if (!studentStats[mKey]) studentStats[mKey] = new Set();
          if (studentStats[mKey] instanceof Set) {
            studentStats[mKey].add(dStr);
          }
        }
      }

      // Track dates for each category in the actual month (For visual display in NEIS mismatch details)
      if (!isMenstrual) {
        const dateKey = `${statKey}_dates`;
        if (!studentStats[dateKey]) studentStats[dateKey] = [];
        if (!studentStats[dateKey].includes(dStr)) {
          studentStats[dateKey].push(dStr);
        }
      }
    });
  });

  // 1.5. Read existing AttendanceData to prepare for non-destructive overwriting
  const lastRowData = dataSheet.getLastRow();
  let existingData = [];
  let existingBg = [];
  if (lastRowData > 1) {
    existingData = dataSheet.getRange(2, 1, lastRowData - 1, 26).getValues();
    existingBg = dataSheet.getRange(2, 1, lastRowData - 1, 26).getBackgrounds();
  }
  
  // Create a map of existing rows by key "월_학번_이름" -> { index, values, bg }
  const existingMap = new Map();
  existingData.forEach((row, i) => {
    const month = String(row[1] || "").trim();
    const sid = String(row[2] || "").trim();
    const name = String(row[3] || "").trim();
    if (month && sid && name) {
      const key = `${month}_${sid}_${name}`;
      existingMap.set(key, { index: i, values: row, bg: existingBg[i] });
    }
  });

  const studentMap = getStudentDirectoryMap();

  // 2. Iterate NeisData and Compare -> Build AttendanceData
  const mapping = [
    { col: 4, cat: "결석", sub: "질병" }, { col: 5, cat: "결석", sub: "미인정" }, { col: 6, cat: "결석", sub: "기타" }, { col: 7, cat: "결석", sub: "인정" },
    { col: 8, cat: "지각", sub: "질병" }, { col: 9, cat: "지각", sub: "미인정" }, { col: 10, cat: "지각", sub: "기타" }, { col: 11, cat: "지각", sub: "인정" },
    { col: 12, cat: "조퇴", sub: "질병" }, { col: 13, cat: "조퇴", sub: "미인정" }, { col: 14, cat: "조퇴", sub: "기타" }, { col: 15, cat: "조퇴", sub: "인정" },
    { col: 16, cat: "결과", sub: "질병" }, { col: 17, cat: "결과", sub: "미인정" }, { col: 18, cat: "결과", sub: "기타" }, { col: 19, cat: "결과", sub: "인정" }
  ];

  const newRowsToAppend = [];

  neisRows.forEach((row) => {
    if (!row[1] || !row[2]) return;
    const rowMonth = parseInt(row[1]);
    const sid = String(row[2]).trim();
    
    const gradeNum = sid.length >= 3 ? parseInt(sid.charAt(0), 10) : -1;
    // Parse class number from studentId (studentId format: G-CC-NN, e.g., 30101 -> class 1)
    const cNum = sid.length >= 3 ? parseInt(sid.substring(1,3), 10) : -1;
    
    if (gradeNum !== settings.grade || cNum > settings.classes) return;
    
    // Apply Month & Class Filter
    const matchesMonth = (targetMonth === "all" || String(rowMonth) === String(targetMonth));
    const matchesClass = (targetClass === "all" || String(cNum) === String(targetClass));
    if (!matchesMonth || !matchesClass) return;

    const name = String(row[3]).trim();
    const key = `${sid}_${rowMonth}`;
    const studentData = aggMap.get(key) || { stats: {} };
    const studentStats = studentData.stats;
    
    const studentInfo = studentMap.get(sid);
    const genderStr = studentInfo ? String(studentInfo.gender || "").trim().toLowerCase() : "";
    const isFemale = studentInfo && (genderStr === '여' || genderStr === '여자' || genderStr === 'f' || genderStr === 'female');
    
    // Check if there are any exam periods falling in this rowMonth
    const hasExamInMonth = examPeriods.some(period => {
      if (!period.startStr) return false;
      const startMonth = parseInt(period.startStr.split('-')[1], 10);
      const endMonth = period.endStr ? parseInt(period.endStr.split('-')[1], 10) : startMonth;
      return startMonth === rowMonth || endMonth === rowMonth;
    });

    let mismatches = [];
    let isMatch = true;
    
    // 순번, 월, 학번, 성명
    const newRow = [row[0], row[1], row[2], row[3]];
    const bgRow = [null, null, null, null];
    
    let absTotal = 0;
    let tardyTotal = 0;
    let earlyTotal = 0;
    let skipTotal = 0;

    mapping.forEach(m => {
      const neisVal = parseInt(row[m.col]) || 0;
      const statKey = `${m.cat}_${m.sub}`;
      const localValRaw = studentStats[statKey];
      let localVal = 0;
      
      if (m.cat === "결석") {
        localVal = (typeof localValRaw === 'number') ? localValRaw : 0;
        if (m.sub !== "인정") absTotal += localVal;
      } else if (m.cat === "지각") {
        localVal = (localValRaw instanceof Set) ? localValRaw.size : 0;
        if (m.sub !== "인정") tardyTotal += localVal;
      } else if (m.cat === "조퇴") {
        localVal = (localValRaw instanceof Set) ? localValRaw.size : 0;
        if (m.sub !== "인정") earlyTotal += localVal;
      } else if (m.cat === "결과") {
        localVal = (localValRaw instanceof Set) ? localValRaw.size : 0;
        if (m.sub !== "인정") skipTotal += localVal;
      }
      
      newRow.push(localVal);
      
      let isMismatch = neisVal !== localVal;
      
      // Unexcused (미인정) exception:
      // Unexcused absences/tardies/earlies/skips do not have scanned excuse forms.
      // Therefore, do not flag a mismatch for "미인정" category.
      if (isMismatch && m.sub === "미인정") {
        isMismatch = false;
      }
      
      // Menstrual pain (생리통) exception:
      // If female, subCategory is "인정":
      // 1) If NEIS count is greater than local count, the difference (neisVal - localVal) is assumed to be undocumented menstrual pain absences (since no document is required outside exam periods).
      //    We suppress the mismatch if the total menstrual count (local menstrual docs + undocumented diff) is within the allowed monthly limit (1 for 결석, 3 for 지각/조퇴/결과).
      // 2) If NEIS count is less than local count, we suppress the mismatch if the difference (localVal - neisVal) is within the local menstrual records.
      if (isMismatch && m.sub === "인정" && isFemale) {
        const mKey = `${statKey}_menstrual`;
        const localMenstrualRaw = studentStats[mKey];
        const localMenstrualVal = (localMenstrualRaw instanceof Set) ? localMenstrualRaw.size : (typeof localMenstrualRaw === 'number' ? localMenstrualRaw : 0);
        
        if (neisVal > localVal) {
          const diff = neisVal - localVal;
          const limit = (m.cat === "결석") ? 1 : 3;
          if (diff + localMenstrualVal <= limit) {
            isMismatch = false;
          }
        } else if (neisVal < localVal) {
          const diff = localVal - neisVal;
          if (diff <= localMenstrualVal) {
            isMismatch = false;
          }
        }
      }
      
      // School Experiential Learning (학교장허가체험학습) Exception:
      // If category is "인정", and there are at least 2 scanned documents (신청서 + 결과보고서) 
      // marked as "학교장허가체험학습", we suppress the mismatch error regardless of the day count difference.
      if (isMismatch && m.sub === "인정" && studentStats["체험학습_건수"] >= 2) {
        isMismatch = false;
      }

      if (isMismatch) {
        isMatch = false;
        const dateKey = `${statKey}_dates`;
        const localDates = studentStats[dateKey] || [];
        const dateSuffix = localDates.length > 0 ? ` [보관서류 날짜: ${localDates.join(", ")}]` : " [보관서류 날짜: 없음]";
        mismatches.push(`${m.cat}(${m.sub}) 나이스:${neisVal} / 서류:${localVal}${dateSuffix}`);
        bgRow.push('#fff2cc'); // Yellow for mismatch
      } else {
        bgRow.push(null);
      }
    });
    
    // 총계 (결석, 지각, 조퇴, 결과)
    newRow.push(absTotal, tardyTotal, earlyTotal, skipTotal);
    bgRow.push(null, null, null, null);
    
    // 일치여부, 불일치내용
    newRow.push(isMatch ? "✅ 일치" : "❌ 불일치", isMatch ? "-" : mismatches.join("\n"));
    bgRow.push(isMatch ? null : '#fce5cd', null); // Subtle orange for mismatch status
    
    // Check if duplicate (same month, studentId, name) exists in current AttendanceData
    const dataKey = `${rowMonth}_${sid}_${name}`;
    const matched = existingMap.get(dataKey);
    if (matched) {
      // Keep original sequence number (순번)
      newRow[0] = matched.values[0];
      
      // Overwrite the entry in existing data array
      existingData[matched.index] = newRow;
      existingBg[matched.index] = bgRow;
    } else {
      newRowsToAppend.push({ newRow, bgRow });
    }
  });
  
  // Append new non-existing rows
  newRowsToAppend.forEach(item => {
    existingData.push(item.newRow);
    existingBg.push(item.bgRow);
  });
  
  // Clear the existing content completely before writing back the merged data
  const lastRow = dataSheet.getLastRow();
  if (lastRow > 1) {
    const numRows = lastRow - 1;
    const clearRange = dataSheet.getRange(2, 1, numRows, 26);
    clearRange.clearContent();
    clearRange.setBackground(null);
  }
  
  // Write the merged data back to the sheet in a single call
  if (existingData.length > 0) {
    const targetRange = dataSheet.getRange(2, 1, existingData.length, 26);
    targetRange.setValues(existingData);
    targetRange.setBackgrounds(existingBg);
  }
  
  return "대조 및 AttendanceData 시트 업데이트 완료 (" + (targetMonth === "all" ? "전체 월" : targetMonth + "월") + " / " + (targetClass === "all" ? "전체 학급" : targetClass + "반") + ")";
}

/**
 * Save confirmed data to Sheet
 */
function saveToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  
  // 0. Final Normalization of Categories & Reason Details
  const norm = normalizeCategories(data.category, data.subCategory, data.filename, data.reasonDetail);
  data.category = norm.category;
  data.subCategory = norm.subCategory;
  data.reasonDetail = norm.reasonDetail;
  
  // 1. Duplicate Detection (Check Date & StudentId)
  const dateStr = data.date || "";
  const idStr = data.studentId || "";
  const existingRecords = sheet.getDataRange().getValues();
  
  const isDuplicate = existingRecords.some(row => {
    const d = parseDateSafe(row[0]);
    if (!d) return false;
    const rDate = Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
    return rDate === dateStr && row[1] == idStr;
  });
  
  if (isDuplicate) return "중복: [" + data.name + "] 이미 해당 날짜 및 학번의 기록이 존재합니다.";

  // 2. Student Directory Verification (Optional Info)
  let verificationNote = "";
  const studentInfo = verifyStudent(idStr, data.name);
  if (studentInfo.match && studentInfo.correctId && studentInfo.correctName) {
    // [Force] Use directory values for saving, not raw AI data
    data.studentId = studentInfo.correctId;
    data.name = studentInfo.correctName;
  } else {
    verificationNote = "⚠️ 명부 불일치";
  }

  // 3. Attachments format rule
  let attachments = data.attachments !== undefined ? data.attachments : "×";
  let subCat = (data.subCategory || "").trim();
  let reason = (data.reasonDetail || "").trim();
  
  // 3.5. Force NEIS string standard format: YYYY-MM-DD~YYYY-MM-DD 대분류소분류(사유상세)(X일간)
  const finalDate = data.date || "";
  const finalCat = data.category || "";
  const finalSub = data.subCategory || "";
  const finalReason = reason;
  
  // Parse total days
  let daysNum = 1;
  const daysRaw = data.totalDays || "1일간";
  const daysMatch = daysRaw.match(/(\d+)/);
  if (daysMatch) {
    daysNum = parseInt(daysMatch[1], 10);
  }
  
  // Calculate date range string
  let dateRangeStr = finalDate;
  if (daysNum > 1 && finalDate) {
    let endDateStr = "";
    if (data.periodEnd && typeof data.periodEnd === 'string') {
      const endMatch = data.periodEnd.match(/(\d{4}-\d{2}-\d{2})/);
      if (endMatch) {
        endDateStr = endMatch[1];
      }
    }
    
    // Fallback: calculate end date if not explicitly in periodEnd
    if (!endDateStr) {
      try {
        const parts = finalDate.split('-');
        if (parts.length === 3) {
          const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
          const endDateObj = new Date(startDateObj.getTime() + (daysNum - 1) * 24 * 60 * 60 * 1000);
          endDateStr = Utilities.formatDate(endDateObj, "GMT+9", "yyyy-MM-dd");
        }
      } catch (e) {}
    }
    
    if (endDateStr && endDateStr !== finalDate) {
      dateRangeStr = `${finalDate}~${endDateStr}`;
    }
  }
  
  const isReasonNone = (str) => !str || /^(none|null|(해당\s*)?없음|첨부서류\s*누락|×)$/i.test(String(str).trim());
  
  if (finalDate && finalCat && finalSub) {
    if (data.preserveNeisString && data.neisString) {
      // Preserve original unsplit neisString
    } else {
      const daySuffix = daysNum > 1 ? `(${daysNum}일간)` : "";
      if (!isReasonNone(finalReason)) {
        data.neisString = `${dateRangeStr} ${finalCat}${finalSub}(${finalReason})${daySuffix}`;
      } else {
        data.neisString = `${dateRangeStr} ${finalCat}${finalSub}${daySuffix}`;
      }
    }
  }

  // 4. Save as Hyperlink for premium experience
  const fileNameCell = `=HYPERLINK("${data.fileUrl}", "${data.filename}")`;

  const safeVal = (val) => val !== undefined ? val : '×';

  const formatSignForSheet = (val) => {
    if (val === "있음") return "";
    if (val === "없음") return "×";
    if (!val) return ""; // 빈칸 유지 (서명 있음)
    const s = String(val).trim();
    if (/^(무|x|×|no|none|n|false|서명\s*안됨|안됨|미서명|서명\s*없음|없음)$/i.test(s)) {
      return "×";
    }
    return ""; // 그 외에는 서명된 것으로 간주하여 빈칸 처리
  };

  sheet.appendRow([
    data.date || '',                // A (1) 일자
    data.studentId || '',           // B (2) 학번
    data.name || '',                // C (3) 성명
    formatSignForSheet(data.studentSigned),    // D (4) 학생서명
    data.category || '',            // E (5) 대분류
    subCat,                         // F (6) 소분류
    formatPeriodSafe(data.periodStart),      // G (7) 시작일시
    formatPeriodSafe(data.periodEnd),        // H (8) 종료일시
    data.totalDays || '',           // I (9) 일수
    (data.reasonDetail || '') + (verificationNote ? " ["+verificationNote+"]" : ""), // J (10) 사유상세
    data.parentName || '',          // K (11) 학부모성함
    formatSignForSheet(data.parentSigned),     // L (12) 학부모서명
    data.teacherName || '',         // M (13) 담임성함
    formatSignForSheet(data.teacherSigned),    // N (14) 담임서명
    attachments,                    // O (15) 첨부서류
    data.docStartDate || '',        // P (16) 통원일
    data.docEndDate || '',          // Q (17) 통원종료일
    data.issuer || '',              // R (18) 발급처
    fileNameCell,                   // S (19) 파일명
    data.ruleCheck || '',           // T (20) 규정위반여부
    new Date(),                     // U (21) 스캔일시
    data.neisString || '',          // V (22) NEIS문구
    ''                              // W (23) 보완완료
  ]);
  return "저장 완료!";
}

/**
 * Save Multiple Records Sequentially to Preserve Order
 */
function saveMultipleToSheet(items) {
  if (!items || items.length === 0) return "저장할 데이터가 없습니다.";
  
  let successCount = 0;
  let skipCount = 0;
  let messages = [];
  
  items.forEach(item => {
    const result = saveToSheet(item);
    if (result.includes("저장 완료")) {
      successCount++;
    } else {
      skipCount++;
      messages.push(result);
    }
  });
  
  let finalMsg = `총 ${items.length}건 중 ${successCount}건 저장 완료!`;
  if (skipCount > 0) finalMsg += ` (${skipCount}건 제외: ${messages.join(', ')})`;
  
  return finalMsg;
}

/**
 * Organize File into Subfolders (Year/Month/Class/Name)
 */
function organizeFile(file, data) {
  try {
    const rootFolder = getTargetFolder();
    const date = parseDateSafe(data.date) || new Date();
    const year = getYearSafe(date) + "년";
    const month = getMonthSafe(date) + "월";
    
    // Class folder (studentId: G-CC-NN)
    let classNum = "미분류";
    if (data.studentId && data.studentId.length >= 3) {
      classNum = parseInt(data.studentId.substring(1, 3)) + "반";
    }
    
    // Create folders if they don't exist
    const yearFolder = getSubFolder(rootFolder, year);
    const monthFolder = getSubFolder(yearFolder, month);
    const classFolder = getSubFolder(monthFolder, classNum);
    const studentFolder = getSubFolder(classFolder, data.name || "무명이");
    
    file.moveTo(studentFolder);
  } catch (e) {
    Logger.log("폴더 이동 실패: " + e.toString());
  }
}

function getSubFolder(parent, folderName) {
  const folders = parent.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(folderName);
}

/**
 * Helper: Get NEIS Folder
 */
function getNeisFolder() {
  const folderId = CONFIG.NEIS_FOLDER_ID;
  if (folderId && folderId.trim().length > 0) {
    try {
      // Split the ID if it contains query parameters (like ?dmr)
      const cleanId = folderId.split('?')[0]; 
      return DriveApp.getFolderById(cleanId);
    } catch (e) {
      Logger.log("경고: NEIS 폴더 ID가 잘못되었습니다. 기본 업로드 폴더를 사용합니다.");
    }
  }
  return getTargetFolder(); // Fallback
}

/**
 * Safe date parsing utility
 */
function parseDateSafe(val) {
  if (val instanceof Date) return val;
  if (!val) return null;
  
  const str = String(val).trim();
  if (str.length === 0) return null;
  
  // Extract date part (e.g., "2026-05-21 1교시" -> "2026-05-21")
  // Normalize dots to hyphens (e.g., "2026.05.21" -> "2026-05-21")
  const datePart = str.split(' ')[0].replace(/\./g, '-');
  const d = new Date(datePart);
  if (!isNaN(d.getTime())) return d;
  
  // Fallback for standard date parsing
  const fallback = new Date(str);
  if (!isNaN(fallback.getTime())) return fallback;
  
  return null;
}

/**
 * Safe month extraction (GMT+9 timezone consistent)
 */
function getMonthSafe(date) {
  if (!date || isNaN(date.getTime())) return null;
  const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  return parseInt(dateStr.split('-')[1], 10);
}

/**
 * Safe year extraction (GMT+9 timezone consistent)
 */
function getYearSafe(date) {
  if (!date || isNaN(date.getTime())) return null;
  const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  return parseInt(dateStr.split('-')[0], 10);
}

/**
 * Helper: Safely formats period input (Dates or strings) to 'YYYY-MM-DD X교시'
 */
function formatPeriodSafe(val) {
  if (!val) return "-";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+9", "yyyy-MM-dd");
  }
  
  let str = String(val).trim();
  
  // Clean up any leading zeros in "교시" (e.g. 01교시 -> 1교시)
  str = str.replace(/(\s+|[^0-9])0+(\d+)\s*교시/g, "$1$2교시");

  // If it's already in the correct format, return immediately
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str;
  }
  
  const parsed = parseDateSafe(str);
  if (parsed) {
    const matchLesson = str.match(/(\d+\s*교시)/);
    const datePart = Utilities.formatDate(parsed, "GMT+9", "yyyy-MM-dd");
    return matchLesson ? `${datePart} ${matchLesson[1]}` : datePart;
  }
  
  return str;
}

/**
 * Check if the student is active on a specific attendance date based on remarks and statusDate
 */
function isStudentActiveOnDate(status, statusDateRaw, checkDateRaw) {
  if (!status || !statusDateRaw) return true; // No special status -> active
  const statusDate = parseDateSafe(statusDateRaw);
  const checkDate = parseDateSafe(checkDateRaw);
  if (!statusDate || !checkDate) return true; // Safe fallback if dates cannot be parsed
  
  const sStr = Utilities.formatDate(statusDate, "GMT+9", "yyyy-MM-dd");
  const cStr = Utilities.formatDate(checkDate, "GMT+9", "yyyy-MM-dd");
  const sParts = sStr.split('-');
  const cParts = cStr.split('-');
  const sTime = new Date(parseInt(sParts[0], 10), parseInt(sParts[1], 10) - 1, parseInt(sParts[2], 10)).getTime();
  const cTime = new Date(parseInt(cParts[0], 10), parseInt(cParts[1], 10) - 1, parseInt(cParts[2], 10)).getTime();
  
  if (status.includes('자퇴') || status.includes('전출') || status.includes('퇴학')) {
    // Active up to the statusDate (inclusive)
    return cTime <= sTime;
  } else if (status.includes('전입')) {
    // Active from the statusDate (inclusive)
    return cTime >= sTime;
  }
  return true;
}

/**
 * Verify Student in Directory
 */
function verifyStudent(id, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  if (!sheet) return { match: true }; 

  const settings = getSystemSettings();

  const data = sheet.getDataRange().getValues().slice(1);
  const studentList = [];
  data.forEach(row => {
    if (!row[0] || !row[1] || !row[2]) return;
    const grade = parseInt(row[0], 10);
    if (grade !== settings.grade) return; // Only verify for target grade
    studentList.push({
      grade: row[0],
      cls: row[1],
      num: row[2],
      name: String(row[3]).trim(),
      id: `${row[0]}${row[1].toString().padStart(2, '0')}${row[2].toString().padStart(2, '0')}`,
      status: String(row[5] || "").trim(), // 비고 (자퇴, 전출, 전입 등)
      statusDate: row[6] // 기준일 (YYYY-MM-DD)
    });
  });

  const inputIdRaw = String(id || "").trim();
  const inputNameRaw = String(name || "").trim();

  // Strip '(공동)' from student name for comparison with StudentDirectory
  const inputNameClean = inputNameRaw.replace(/\(공동\)/g, "").trim();

  // 0. Clean Input ID (Handle common OCR errors: S->5, O->0, I/L->1, Z->2, B->8)
  const inputIdClean = inputIdRaw.toUpperCase()
    .replace(/S/g, "5")
    .replace(/O/g, "0")
    .replace(/[IL]/g, "1")
    .replace(/Z/g, "2")
    .replace(/B/g, "8")
    .replace(/\D/g, ""); // Remove any remaining non-digits

  // 1. Exact Name Match (Strongest fallback for messy digits)
  const nameMatches = studentList.filter(s => s.name === inputNameClean);
  if (nameMatches.length === 1) {
    const matchedStudent = nameMatches[0];
    let isIdCompatible = true;
    
    // AI가 판독한 학번이 있고, 그 학번의 학년/반(앞 3자리, 예: "305")이 매칭된 학생과 다르면 오보정 방지를 위해 자동 매칭을 차단합니다.
    if (inputIdClean.length >= 3) {
      const aiClassPrefix = inputIdClean.substring(0, 3);
      const dbClassPrefix = matchedStudent.id.substring(0, 3);
      if (aiClassPrefix !== dbClassPrefix) {
        isIdCompatible = false;
      }
    }
    
    if (isIdCompatible) {
      return { match: true, correctId: matchedStudent.id, correctName: matchedStudent.name, status: matchedStudent.status, statusDate: matchedStudent.statusDate };
    }
  } else if (nameMatches.length > 1) {
    // If namesake, try to find the one with closest ID similarity within the same class/grade
    let closest = null;
    let maxSimilarity = -1;
    
    nameMatches.forEach(s => {
      // AI가 판독한 학번의 학년/반과 다르면 동명이인 후보에서 배제합니다.
      if (inputIdClean.length >= 3) {
        const aiClassPrefix = inputIdClean.substring(0, 3);
        const dbClassPrefix = s.id.substring(0, 3);
        if (aiClassPrefix !== dbClassPrefix) return;
      }
      
      let similarity = 0;
      for (let i = 0; i < Math.min(s.id.length, inputIdClean.length); i++) {
        if (s.id[i] === inputIdClean[i]) similarity++;
      }
      if (similarity > maxSimilarity) { maxSimilarity = similarity; closest = s; }
    });
    
    if (closest) {
      return { match: true, correctId: closest.id, correctName: closest.name, status: closest.status, statusDate: closest.statusDate };
    }
  }

  // 2. ID Match (Using Cleaned ID)
  const idMatch = studentList.find(s => s.id === inputIdClean);
  if (idMatch) {
    // If ID is correct, we accept most names unless they are completely different
    // (e.g. at least one character match or name is short/long)
    return { match: true, correctId: idMatch.id, correctName: idMatch.name, status: idMatch.status, statusDate: idMatch.statusDate };
  }

  // 3. Last Resort: Partial ID match (4 out of 5 digits match) + Name partial match
  for (const s of studentList) {
    let idDiffCount = 0;
    if (s.id.length === inputIdClean.length) {
      for (let i = 0; i < s.id.length; i++) {
        if (s.id[i] !== inputIdClean[i]) idDiffCount++;
      }
    }
    
    // If 4/5 digits match and Name contains or is contained in Directory Name
    const nameSimilarity = inputNameClean.length >= 2 && (s.name.includes(inputNameClean) || inputNameClean.includes(s.name));
    if (idDiffCount <= 1 && nameSimilarity) {
      return { match: true, correctId: s.id, correctName: s.name, status: s.status, statusDate: s.statusDate };
    }
  }

  return { match: false };
}

/**
 * Get Student Directory gender map
 */
function getStudentDirectoryMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  const map = new Map();
  if (!sheet) return map;

  const settings = getSystemSettings();

  const data = sheet.getDataRange().getValues().slice(1);
  data.forEach(row => {
    if (!row[0] || !row[1] || !row[2]) return;
    const grade = parseInt(row[0], 10);
    if (grade !== settings.grade) return; // Only load target grade
    const cls = row[1];
    const num = row[2];
    const name = String(row[3]).trim();
    const gender = String(row[4] || "").trim(); // 성별
    const id = `${grade}${cls.toString().padStart(2, '0')}${num.toString().padStart(2, '0')}`;
    map.set(id, { name, gender });
  });
  return map;
}

/**
 * Retrieve all exam periods from Sheet
 */
function getExamPeriods() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.EXAM_PERIODS || 'ExamPeriods');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues().slice(1);
  const periods = [];
  data.forEach((row, index) => {
    if (row[0]) {
      let startStr = "";
      let endStr = "";
      try {
        const start = new Date(row[0]);
        const end = row[1] ? new Date(row[1]) : new Date(row[0]);
        if (!isNaN(start.getTime()) && !isNaN(end.getTime())) {
          startStr = Utilities.formatDate(start, "GMT+9", "yyyy-MM-dd");
          endStr = Utilities.formatDate(end, "GMT+9", "yyyy-MM-dd");
          periods.push({
            index: index + 2, // 1-indexed and skip header row
            startStr: startStr,
            endStr: endStr,
            name: String(row[2] || "").trim()
          });
        }
      } catch (e) {
        Logger.log("Error parsing exam period row: " + e.toString());
      }
    }
  });
  return periods;
}

/**
 * Add a new exam period to Sheet
 */
function addExamPeriod(startDateStr, endDateStr, name) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EXAM_PERIODS || 'ExamPeriods');
    if (!sheet) throw new Error("시험기간 시트가 없습니다.");

    // Validate dates
    const start = new Date(startDateStr);
    const end = new Date(endDateStr || startDateStr);
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      throw new Error("올바른 날짜 형식이 아닙니다.");
    }

    const startFormatted = Utilities.formatDate(start, "GMT+9", "yyyy-MM-dd");
    const endFormatted = Utilities.formatDate(end, "GMT+9", "yyyy-MM-dd");

    sheet.appendRow([startFormatted, endFormatted, String(name || "").trim()]);
    
    // Auto-sort by start date in ascending order
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, 3).sort({ column: 1, ascending: true });
    }

    return { success: true, message: "시험 일정이 성공적으로 추가되었습니다." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Delete an exam period from Sheet
 */
function deleteExamPeriod(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.EXAM_PERIODS || 'ExamPeriods');
    if (!sheet) throw new Error("시험기간 시트가 없습니다.");

    const lastRow = sheet.getLastRow();
    const idx = parseInt(rowIndex, 10);
    if (isNaN(idx) || idx < 2 || idx > lastRow) {
      throw new Error("유효하지 않은 일련번호입니다.");
    }

    sheet.deleteRow(idx);
    return { success: true, message: "시험 일정이 성공적으로 삭제되었습니다." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Cross-Day Validation: Check if the previous school day had a hospital document
 */
/**
 * Cross-Day Validation: Check if the previous school day had a hospital document
 */
function checkPreviousDayHospitalDoc(studentId, currentDateStr, currentBatchItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (!sheet) return false;
  
  const cleanDateStr = currentDateStr ? String(currentDateStr).split(' ')[0] : '';
  const current = new Date(cleanDateStr);
  if (isNaN(current.getTime())) return false; // Invalid date
  
  const hospitalKeywords = ["진료", "진단", "처방", "통원", "입원", "퇴원", "입퇴원", "입·퇴원", "입/퇴원", "입-퇴원", "소견", "약", "영수증", "보건", "보건실", "입실", "확인증", "병원", "약국", "의원", "치과", "한의원"];
  
  let mostRecentPriorDate = null;
  let mostRecentPriorHasHospital = false;

  const checkItem = (rId, rDateStr, rCat, rAttach, rIssuer) => {
    if (String(rId).trim() !== String(studentId).trim()) return;
    if (!rCat.includes('질병')) return;
    
    const rDate = new Date(rDateStr.split(' ')[0]);
    if (isNaN(rDate.getTime())) return;
    
    const diffTime = current.getTime() - rDate.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    // Must be a prior day and within 4 days (matching the consecutive disease rule in DataCleanup)
    if (diffDays > 0 && diffDays <= 4) {
      if (!mostRecentPriorDate || rDate > mostRecentPriorDate) {
        mostRecentPriorDate = rDate;
        const cleanAttach = String(rAttach || "").replace(/\s+/g, "");
        const cleanIssuer = String(rIssuer || "").replace(/\s+/g, "");
        mostRecentPriorHasHospital = hospitalKeywords.some(kw => cleanAttach.includes(kw) || cleanIssuer.includes(kw));
      }
    }
  };

  // 1. Check in the current batch first
  if (currentBatchItems && Array.isArray(currentBatchItems)) {
    currentBatchItems.forEach(item => {
      if (!item) return;
      const rId = String(item.studentId || "").trim();
      const rCat = String(item.category || "").trim();
      const rAttach = String(item.attachments || "").trim();
      const rIssuer = String(item.issuer || "").trim();
      const itemDate = item.date || "";
      
      // Calculate covered dates range
      let daysNum = 1;
      const daysRaw = item.totalDays || "1일간";
      const daysMatch = daysRaw.match(/(\d+)/);
      if (daysMatch) daysNum = parseInt(daysMatch[1], 10);

      try {
        const parts = itemDate.split('-');
        if (parts.length === 3) {
          const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
          for (let d = 0; d < daysNum; d++) {
            const currentDateObj = new Date(startDateObj.getTime() + d * 24 * 60 * 60 * 1000);
            const dStr = Utilities.formatDate(currentDateObj, "GMT+9", "yyyy-MM-dd");
            checkItem(rId, dStr, rCat, rAttach, rIssuer);
          }
        } else {
          checkItem(rId, itemDate, rCat, rAttach, rIssuer);
        }
      } catch (e) {
        checkItem(rId, itemDate, rCat, rAttach, rIssuer);
      }
    });
  }

  // 2. Check in the database (AttendanceDB)
  const data = sheet.getDataRange().getValues().slice(1);
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || !row[1]) continue;
    
    const d = parseDateSafe(row[0]);
    if (!d) continue;
    const rDateStr = Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
    const rId = String(row[1]).trim();
    const rCat = String(row[4] || "").trim();
    const rIssuer = String(row[13] || "").trim();
    const rAttach = String(row[14] || "").trim();
    
    checkItem(rId, rDateStr, rCat, rAttach, rIssuer);
  }

  return mostRecentPriorDate ? mostRecentPriorHasHospital : false;
}

/**
 * Cross-Day Validation: Check if the same date had a hospital document
 */
/**
 * Cross-Day Validation: Check if the same date had a hospital document
 */
function checkSameDayHospitalDoc(studentId, currentDateStr, currentBatchItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (!sheet) return false;
  
  const cleanDateStr = currentDateStr ? String(currentDateStr).split(' ')[0] : '';
  const current = new Date(cleanDateStr);
  if (isNaN(current.getTime())) return false; // Invalid date
  const targetDateStr = Utilities.formatDate(current, "GMT+9", "yyyy-MM-dd");
  const hospitalKeywords = ["진료", "진단", "처방", "통원", "입원", "퇴원", "입퇴원", "입·퇴원", "입/퇴원", "입-퇴원", "소견", "약", "영수증", "보건", "보건실", "입실", "확인증", "병원", "약국", "의원", "치과", "한의원"];

  // 1. Check in the current batch first
  if (currentBatchItems && Array.isArray(currentBatchItems)) {
    for (let i = 0; i < currentBatchItems.length; i++) {
      const item = currentBatchItems[i];
      if (!item) continue;
      const rId = String(item.studentId || "").trim();
      const rCat = String(item.category || "").trim();
      const rAttach = String(item.attachments || "").trim();
      const rIssuer = String(item.issuer || "").trim();
      
      const itemDate = item.date || "";
      if (rId === String(studentId).trim() && rCat.includes('질병')) {
        let daysNum = 1;
        const daysRaw = item.totalDays || "1일간";
        const daysMatch = daysRaw.match(/(\d+)/);
        if (daysMatch) daysNum = parseInt(daysMatch[1], 10);

        const coveredDates = [];
        try {
          const parts = itemDate.split('-');
          if (parts.length === 3) {
            const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
            for (let d = 0; d < daysNum; d++) {
              const currentDateObj = new Date(startDateObj.getTime() + d * 24 * 60 * 60 * 1000);
              coveredDates.push(Utilities.formatDate(currentDateObj, "GMT+9", "yyyy-MM-dd"));
            }
          } else {
            coveredDates.push(itemDate);
          }
        } catch (e) {
          coveredDates.push(itemDate);
        }

        if (coveredDates.includes(targetDateStr)) {
          const cleanAttach = String(rAttach).replace(/\s+/g, "");
          const cleanIssuer = String(rIssuer).replace(/\s+/g, "");
          if (hospitalKeywords.some(kw => cleanAttach.includes(kw) || cleanIssuer.includes(kw))) {
            return true;
          }
        }
      }
    }
  }

  // 2. Check in the database (AttendanceDB)
  const data = sheet.getDataRange().getValues().slice(1);
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || !row[1]) continue;
    
    const d = parseDateSafe(row[0]);
    if (!d) continue;
    const rDate = Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
    const rId = String(row[1]).trim();
    const rCat = String(row[4] || "").trim(); // 대분류
    const rIssuer = String(row[13] || "").trim();
    const rAttach = String(row[14] || "").trim();
    
    if (rId === String(studentId).trim() && rDate === targetDateStr && rCat.includes('질병')) {
      const cleanAttach = String(rAttach).replace(/\s+/g, "");
      const cleanIssuer = String(rIssuer).replace(/\s+/g, "");
      if (hospitalKeywords.some(kw => cleanAttach.includes(kw) || cleanIssuer.includes(kw))) {
        return true;
      }
    }
  }
  return false;
}


/**
 * Reset Database (For testing)
 */
function resetDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Reset AttendanceDB
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (attSheet && attSheet.getLastRow() > 1) {
    attSheet.getRange(2, 1, attSheet.getLastRow() - 1, attSheet.getLastColumn()).clearContent();
  }
  
  // 2. Reset NeisData
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  if (neisSheet && neisSheet.getLastRow() > 1) {
    neisSheet.getRange(2, 1, neisSheet.getLastRow() - 1, neisSheet.getLastColumn()).clearContent();
  }

  // 3. Reset SystemLog
  const logSheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  if (logSheet && logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).clearContent();
  }
  
  return "데이터베이스 및 시스템 로그가 초기화되었습니다. (헤더 제외)";
}

/**
 * Compare local data with NEIS data (Detailed Version)
 */
function compareWithNeis(targetMonth) {
  targetMonth = targetMonth || "all";
  // Recalculate reconciliation first to update matching status in AttendanceData sheet
  reconcileNeisWithAttendance(targetMonth, "all");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
  
  if (!dataSheet) return [];
  
  const discrepancies = [];
  if (dataSheet.getLastRow() > 1) {
    const dData = dataSheet.getDataRange().getValues().slice(1);
    dData.forEach(row => {
      const rowMonth = String(row[1]).trim();
      
      // Filter by targetMonth
      if (targetMonth !== "all" && String(rowMonth) !== String(targetMonth)) {
        return;
      }
      
      const sidStr = String(row[2]).trim();
      const gradeNum = sidStr.length >= 3 ? parseInt(sidStr.charAt(0), 10) : -1;
      const cNum = sidStr.length >= 3 ? parseInt(sidStr.substring(1,3), 10) : -1;
      const settings = getSystemSettings();
      if (gradeNum !== settings.grade || cNum > settings.classes) return;

      const name = String(row[3]).trim();
      const matchStatus = String(row[24] || "");
      const mismatchDetails = String(row[25] || "");
      
      if (matchStatus.includes("❌")) {
        // Determine type of mismatch based on details
        let type = 'CONTENT_MISMATCH';
        
        // Highlight "미인정" or "서류누락" or "나이스누락"
        if (mismatchDetails.includes("나이스:0") || mismatchDetails.includes("서류:0")) {
          // If NEIS has records but local doesn't, it is MISSING_SCAN (스캔본 누락)
          // If NEIS has 0 and local has > 0, it is MISSING_NEIS (NEIS 쪽에 서류누락)
          if (/나이스:[1-9]\d* \/ 서류:0/.test(mismatchDetails)) {
            type = 'MISSING_SCAN';
          } else if (/나이스:0 \/ 서류:[1-9]\d*/.test(mismatchDetails)) {
            type = 'MISSING_NEIS';
          }
        }
        
        discrepancies.push({
          type: type,
          date: `${rowMonth}월`,
          studentId: sidStr,
          name: name,
          details: mismatchDetails
        });
      }
    });
  }
  
  return discrepancies;
}

/**
 * Retrieve dynamic dashboard data (Total scans, active month, acknowledged counts, top 5 rows)
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  
  const totalScans = attSheet ? Math.max(0, attSheet.getLastRow() - 1) : 0;
  
  let acknowledgedThisMonth = 0;
  let recentRecords = [];
  
  if (attSheet && attSheet.getLastRow() > 1) {
    const data = attSheet.getDataRange().getValues().slice(1);
    const sorted = [...data].reverse();
    
    for (let i = 0; i < Math.min(5, sorted.length); i++) {
       const row = sorted[i];
       const d = parseDateSafe(row[0]);
       const dateStr = d ? Utilities.formatDate(d, "GMT+9", "yyyy/MM/dd") : "-";
       const issuedStr = row[15] ? String(row[15]) : "-";
       
       recentRecords.push({
         date: dateStr,
         studentId: row[1] || '-',
         name: row[2] || '-',
         category: `${row[4] || ''} (${row[5] || ''})`,
         issuedDate: issuedStr,
         reason: row[9] || ''
       });
    }
    
    const currentMonth = getMonthSafe(new Date());
    data.forEach(row => {
       const dateVal = parseDateSafe(row[0]);
       if (dateVal && getMonthSafe(dateVal) === currentMonth) {
          if ((row[5] || '').toString().includes('인정')) {
            acknowledgedThisMonth++;
          }
       }
    });
  }
  
  let lastNeisMonth = '-';
  if (neisSheet && neisSheet.getLastRow() > 1) {
    const neisData = neisSheet.getDataRange().getValues().slice(1);
    const months = neisData.map(r => parseInt(r[1])).filter(x => !isNaN(x));
    if (months.length > 0) {
       lastNeisMonth = Math.max(...months);
    }
  }
  
  return {
    totalScans,
    lastNeisMonth,
    acknowledgedThisMonth,
    recentRecords
  };
}

/**
 * Filtered Dashboard Analytics for Unified UI
 */
function getDashboardAnalytics(targetMonth, targetClass) {
  // [Performance Optimization] 단순 대시보드 조회 시마다 무거운 전체 대조 연산을 수행하지 않도록 주석 처리합니다.
  // 데이터 동기화는 NEIS 업로드 시점에 자동 수행되며, 필요 시 메뉴에서 수동으로 실행 가능합니다.
  /*
  try {
    reconcileAttendanceDataSheetDirectly(true);
  } catch(e) {
    Logger.log("대시보드 조회 중 자동 재계산 실패: " + e.toString());
  }
  */

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS);
  const studentSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
  
  const settings = getSystemSettings();

  let allStudents = [];
  if (studentSheet && studentSheet.getLastRow() > 1) {
    const sData = studentSheet.getDataRange().getValues().slice(1);
    allStudents = sData.map(row => {
      const remarks = String(row[5] || "").trim();
      const statusDateRaw = row[6];
      const grade = parseInt(row[0], 10);
      
      let isExcluded = false;
      if (remarks && statusDateRaw) {
        const statusDate = parseDateSafe(statusDateRaw);
        if (statusDate) {
          const sYear = getYearSafe(statusDate);
          const sMonth = getMonthSafe(statusDate);
          const sDateOnly = new Date(sYear, sMonth - 1, statusDate.getDate()).getTime();
          
          if (targetMonth !== 'all') {
            const filterMonth = parseInt(targetMonth, 10);
            const filterYear = 2026; // Current academic year
            
            const sTimeMonth = sYear * 12 + sMonth;
            const fTimeMonth = filterYear * 12 + filterMonth;
            
            if (remarks.includes('자퇴') || remarks.includes('전출') || remarks.includes('퇴학')) {
              if (fTimeMonth > sTimeMonth) {
                isExcluded = true;
              }
            } else if (remarks.includes('전입')) {
              if (fTimeMonth < sTimeMonth) {
                isExcluded = true;
              }
            }
          } else {
            // 'all' (전체 월) 조회 시 학적 변동생 예외 처리
            const semesterStart = new Date(2026, 2, 2).getTime(); // 2026년 3월 2일 (새 학년도 첫 수업일 부근)
            if (remarks.includes('전입')) {
              // 학기 시작일 이후 중도 전입해 온 학생은 전체 개근 대상에서 제외
              if (sDateOnly > semesterStart) {
                isExcluded = true;
              }
            } else if (remarks.includes('자퇴') || remarks.includes('전출') || remarks.includes('퇴학')) {
              // 학기 도중 전출/자퇴/퇴학한 학생도 전체 기간을 채우지 못했으므로 제외
              isExcluded = true;
            }
          }
        }
      } else {
        // Fallback if no date is specified
        if (remarks.includes('자퇴') || remarks.includes('전출') || remarks.includes('퇴학')) {
          isExcluded = true;
        }
      }
      
      return {
        classNum: String(row[1]),
        name: String(row[3]),
        studentId: String(row[0]) + String(row[1]).padStart(2, '0') + String(row[2]).padStart(2, '0'),
        excluded: isExcluded || (grade !== settings.grade)
      };
    }).filter(s => s.name && s.name.trim() !== '' && !s.excluded);
  }

  if (targetClass && targetClass !== 'all') {
    allStudents = allStudents.filter(s => s.classNum === String(targetClass));
  }
  
  const hitStudentIds = new Set();
  
  let attCount = 0;
  let neisCount = 0;
  let filteredRecords = [];
  
  if (attSheet && attSheet.getLastRow() > 1) {
    const aData = attSheet.getDataRange().getValues().slice(1);
    aData.forEach(row => {
      const d = parseDateSafe(row[0]);
      if (d && !isNaN(d.getTime())) {
         const rowMonth = String(getMonthSafe(d));
         if (targetMonth === 'all' || rowMonth === String(targetMonth)) {
            const sidStr = String(row[1]);
            const gradeNum = sidStr.length >= 3 ? parseInt(sidStr.charAt(0), 10) : -1;
            const cNum = sidStr.length >= 3 ? parseInt(sidStr.substring(1,3), 10) : -1;
            if (gradeNum === settings.grade && (targetClass === 'all' || cNum === parseInt(targetClass, 10)) && cNum <= settings.classes) {
               hitStudentIds.add(sidStr);
               attCount++;
               
               let formattedIssuedDate = "-";
               if (row[15]) {
                 const idate = new Date(row[15]);
                 if (!isNaN(idate.getTime())) {
                   formattedIssuedDate = Utilities.formatDate(idate, "GMT+9", "yyyy-MM-dd");
                 } else {
                   formattedIssuedDate = String(row[15]);
                 }
               }

               filteredRecords.push({
                 date: row[0] ? Utilities.formatDate(d, "GMT+9", "yyyy/MM/dd") : "-",
                 studentId: row[1] || '-',
                 name: row[2] || '-',
                 category: `${row[4] || ''} (${row[5] || ''})`,
                 issuedDate: formattedIssuedDate,
                 periodStart: row[6] ? formatPeriodSafe(row[6]) : '-',
                 periodEnd: row[7] ? formatPeriodSafe(row[7]) : '-',
                 totalDays: row[8] || '-',
                 reason: row[9] || '',
                 attachments: row[14] || '',
                 ruleCheck: row[19] || '',
                 actionStatus: row[22] || '' // 23rd column (W)
               });
            }
         }
      }
    });
  }
  filteredRecords.reverse();
  
  if (neisSheet && neisSheet.getLastRow() > 1) {
    const nData = neisSheet.getDataRange().getValues().slice(1);
    nData.forEach(row => {
      const rowMonth = String(row[1]);
      if (targetMonth === 'all' || rowMonth === String(targetMonth)) {
        const sidStr = String(row[2]);
        const gradeNum = sidStr.length >= 3 ? parseInt(sidStr.charAt(0), 10) : -1;
        const cNum = sidStr.length >= 3 ? parseInt(sidStr.substring(1,3), 10) : -1;
        if (gradeNum === settings.grade && (targetClass === 'all' || cNum === parseInt(targetClass, 10)) && cNum <= settings.classes) {
           hitStudentIds.add(sidStr);
           
           // columns 20 to 23: stats
           const sum = (parseInt(row[20])||0) + (parseInt(row[21])||0) + (parseInt(row[22])||0) + (parseInt(row[23])||0);
           neisCount += sum;
        }
      }
    });
  }
  
  const perfectStudents = allStudents.filter(s => !hitStudentIds.has(s.studentId));
  perfectStudents.sort((a,b) => parseInt(a.studentId) - parseInt(b.studentId));
  
  let mismatches = [];
  let targetMismatchedUniqueIds = new Set();
  let attendanceDataList = [];
  
  const dataSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
  if (dataSheet && dataSheet.getLastRow() > 1) {
    const dData = dataSheet.getDataRange().getValues().slice(1);
    dData.forEach(row => {
      const rowMonth = String(row[1]).trim();
      const sidStr = String(row[2]).trim();
      const name = String(row[3]).trim();
      const matchStatus = String(row[24] || "");
      const mismatchDetails = String(row[25] || "");
      
      if (targetMonth === 'all' || rowMonth === String(targetMonth)) {
        const gradeNum = sidStr.length >= 3 ? parseInt(sidStr.charAt(0), 10) : -1;
        const cNum = sidStr.length >= 3 ? parseInt(sidStr.substring(1,3), 10) : -1;
        if (gradeNum === settings.grade && (targetClass === 'all' || cNum === parseInt(targetClass, 10)) && cNum <= settings.classes) {
          if (matchStatus.includes("❌")) {
            mismatches.push({
              type: 'CONTENT_MISMATCH',
              date: `${rowMonth}월`,
              studentId: sidStr,
              name: name,
              details: mismatchDetails
            });
            targetMismatchedUniqueIds.add(sidStr);
          }
          
          let sum = 0;
          for (let col = 4; col <= 19; col++) {
            sum += (parseInt(row[col]) || 0);
          }
          if (sum > 0) {
            attendanceDataList.push({
              studentId: sidStr,
              name: name,
              month: rowMonth,
              absent_disease: parseInt(row[4]) || 0,
              absent_unauthorized: parseInt(row[5]) || 0,
              absent_other: parseInt(row[6]) || 0,
              absent_recognized: parseInt(row[7]) || 0,
              tardy_disease: parseInt(row[8]) || 0,
              tardy_unauthorized: parseInt(row[9]) || 0,
              tardy_other: parseInt(row[10]) || 0,
              tardy_recognized: parseInt(row[11]) || 0,
              early_disease: parseInt(row[12]) || 0,
              early_unauthorized: parseInt(row[13]) || 0,
              early_other: parseInt(row[14]) || 0,
              early_recognized: parseInt(row[15]) || 0,
              skipped_disease: parseInt(row[16]) || 0,
              skipped_unauthorized: parseInt(row[17]) || 0,
              skipped_other: parseInt(row[18]) || 0,
              skipped_recognized: parseInt(row[19]) || 0
            });
          }
        }
      }
    });
  }
  
  // Calculate matchRate
  // Match Rate = (Total students mapped - uniquely mismatched students) / Total students mapped * 100
  let matchRate = "100%";
  if (allStudents.length > 0) {
    const correctCount = allStudents.length - targetMismatchedUniqueIds.size;
    matchRate = Math.round((correctCount / allStudents.length) * 100) + "%";
  } else if (allStudents.length === 0 && targetMismatchedUniqueIds.size > 0) {
    matchRate = "0%";
  }

  return {
    perfectList: perfectStudents,
    mismatchList: mismatches,
    attCount: attCount,
    neisCount: neisCount,
    matchRate: matchRate,
    filteredRecords: filteredRecords,
    attendanceDataList: attendanceDataList
  };
}

/**
 * Batch Loader for PDF Generative Loop
 */
function getAllClassesDashboardAnalytics(targetMonth) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const studentSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
   let classes = new Set();
   const settings = getSystemSettings();
   if (studentSheet && studentSheet.getLastRow() > 1) {
     const sData = studentSheet.getDataRange().getValues().slice(1);
     sData.forEach(r => {
       const grade = parseInt(r[0], 10);
       const cls = parseInt(r[1], 10);
       if (grade === settings.grade && cls <= settings.classes) {
         classes.add(cls);
       }
     });
   }
   
   let classArr = Array.from(classes)
                  .map(c => parseInt(c))
                  .filter(c => !isNaN(c))
                  .sort((a,b)=>a-b);
                  
   if (classArr.length === 0) {
     classArr = [];
     for (let i = 1; i <= settings.classes; i++) {
       classArr.push(i);
     }
   } 
   
   const results = [];
   classArr.forEach(c => {
     results.push({
       className: `${c}반`,
       data: getDashboardAnalytics(targetMonth, String(c))
     });
   });
   
   return results;
}

/**
 * Get all document records that cover multiple days (e.g. totalDays > 1) from AttendanceDB.
 * Used for the premium "Integrated Proof Documents Lookup" dashboard feature.
 */
function getMultiDayDocuments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const range = sheet.getDataRange();
  const values = range.getValues().slice(1);
  const formulas = range.getCell(2, 1).offset(0, 0, values.length, range.getLastColumn()).getFormulas();
  
  const results = [];
  
  values.forEach((row, i) => {
    const daysRaw = String(row[8] || "").trim(); // Column I (index 8) is 일수
    const matchDays = daysRaw.match(/(\d+)/);
    if (matchDays) {
      const days = parseInt(matchDays[1], 10);
      if (days > 1) {
        // Parse filename hyperlink from formulas
        const fileFormula = formulas[i][18] || ""; // Column S (index 18) is 파일명
        let fileUrl = "";
        let filename = String(row[18] || "-");
        
        if (fileFormula) {
          const matchLink = fileFormula.match(/=HYPERLINK\("([^"]+)"\s*,\s*"([^"]+)"\)/i);
          if (matchLink) {
            fileUrl = matchLink[1];
            filename = matchLink[2];
          }
        }
        
        // Extract Month from Date (Column A / index 0)
        let month = "-";
        if (row[0]) {
          const d = parseDateSafe(row[0]);
          if (d && !isNaN(d.getTime())) {
            month = `${getMonthSafe(d)}월`;
          }
        }
        
        const startFormatted = formatPeriodSafe(row[6]);
        const endFormatted = formatPeriodSafe(row[7]);
        
        let daysFormatted = daysRaw;
        if (daysRaw) {
          const match = daysRaw.match(/(\d+)/);
          if (match) {
            daysFormatted = `(${match[1]}일간)`;
          }
        }
        
        results.push({
          month: month,
          studentId: String(row[1]).trim(),
          name: String(row[2]).trim(),
          category: `${row[4] || ""} (${row[5] || ""})`,
          period: `${startFormatted} ~ ${endFormatted} ${daysFormatted}`,
          days: daysRaw,
          filename: filename,
          fileUrl: fileUrl
        });
      }
    }
  });
  
  // Sort results: month ascending, studentId ascending
  results.sort((a, b) => {
    const aM = parseInt(a.month) || 0;
    const bM = parseInt(b.month) || 0;
    if (aM !== bM) return aM - bM;
    return a.studentId.localeCompare(b.studentId);
  });
  
  return results;
}

/**
 * Normalize and repair 대분류 (category) and 소분류 (subCategory).
 * Also extracts unrecognized detail terms like 생리통 or 경조사 and merges them into J열 (reasonDetail).
 */
function normalizeCategories(catRaw, subCatRaw, filename, reasonDetail) {
  let category = String(catRaw || "").trim();
  let subCategory = String(subCatRaw || "").trim();
  let reason = String(reasonDetail || "").trim();
  const file = String(filename || "").trim();

  if (!category && !subCategory) {
    return { category: "질병", subCategory: "결석", reasonDetail: reason };
  }

  // 1. Extract extra details (non-standard keywords)
  let extCatDetail = category;
  let extSubDetail = subCategory;

  const catKeywords = ["출석인정", "미인정", "질병", "기타"];
  catKeywords.forEach(kw => { extCatDetail = extCatDetail.replace(kw, ""); });
  extCatDetail = extCatDetail.replace(/인정/g, ""); // "출석인정" 외에 단독 "인정"도 제거

  const subKeywords = ["결석", "지각", "조퇴", "결과", "깸석", "결서", "지가", "조태"];
  subKeywords.forEach(kw => { extSubDetail = extSubDetail.replace(kw, ""); });

  const cleanDetail = (str) => str.replace(/[\(\)\[\]\{\}\s,\/\-_~:;]/g, "").trim();
  extCatDetail = cleanDetail(extCatDetail);
  extSubDetail = cleanDetail(extSubDetail);

  // 2. Determine standardized Category and SubCategory
  const combined = (category + " " + subCategory + " " + file).trim();
  const lowerReason = reason.toLowerCase();
  const lowerCombined = combined.toLowerCase();

  // If the reason or filename/categories contain recognized absence keywords, force Category to "출석인정"
  const hasRecognizedKeyword = /생리|체험학습|독감|인플루엔자|코로나|경조|전염병|감염병|공가|학교장허가/.test(lowerReason) ||
                               /생리|체험학습|독감|인플루엔자|코로나|경조|전염병|감염병|공가|학교장허가/.test(lowerCombined);

  let finalCategory = "질병"; // default
  if (combined.includes("미인정")) {
    finalCategory = "미인정";
  } else if (combined.includes("출석인정") || combined.includes("출석 인정") || (combined.includes("인정") && !combined.includes("미인정")) || hasRecognizedKeyword) {
    finalCategory = "출석인정";
  } else if (combined.includes("질병")) {
    finalCategory = "질병";
  } else if (combined.includes("기타")) {
    finalCategory = "기타";
  }

  let finalSubCategory = "결석"; // default
  if (combined.includes("지각") || combined.includes("지가")) {
    finalSubCategory = "지각";
  } else if (combined.includes("조퇴") || combined.includes("조태")) {
    finalSubCategory = "조퇴";
  } else if (combined.includes("결과")) {
    finalSubCategory = "결과";
  } else if (combined.includes("결석") || combined.includes("깸석") || combined.includes("결서")) {
    finalSubCategory = "결석";
  }

  // 3. Extract and merge any residual details to J열 (reasonDetail)
  let extraDetails = [];
  const allStandardKeywords = ["출석인정", "미인정", "질병", "기타", "결석", "지각", "조퇴", "결과", "깸석", "결서", "지가", "조태", "인정"];
  
  if (extCatDetail && !allStandardKeywords.includes(extCatDetail)) {
    extraDetails.push(extCatDetail);
  }
  if (extSubDetail && !allStandardKeywords.includes(extSubDetail)) {
    extraDetails.push(extSubDetail);
  }

  if (extraDetails.length > 0) {
    const extraStr = extraDetails.filter((v, idx, self) => self.indexOf(v) === idx).join(", ");
    if (extraStr) {
      if (!reason) {
        reason = extraStr;
      } else {
        const cleanReason = reason.replace(/\s+/g, "");
        const cleanExtra = extraStr.replace(/\s+/g, "");
        if (!cleanReason.includes(cleanExtra) && !cleanExtra.includes(cleanReason)) {
          reason = reason + " (" + extraStr + ")";
        }
      }
    }
  }

  // 체험학습 관련 사유를 표준명칭인 '학교장허가체험학습'으로 통일하여 저장되도록 보완
  if (reason && /(학교장허가\s*|현장\s*|교외\s*)*체험\s*(학습)?/g.test(reason)) {
    reason = reason.replace(/(학교장허가\s*|현장\s*|교외\s*)*체험\s*(학습)?/g, "학교장허가체험학습");
  }

  return {
    category: finalCategory,
    subCategory: finalSubCategory,
    reasonDetail: reason
  };
}

/**
 * 스프레드가 열릴 때 상단에 편리한 사용자 맞춤 메뉴를 추가합니다.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AttenTrack AI')
    .addItem('🔄 NEIS 데이터 대조 재생성 (NeisData + AttendanceDB)', 'reconcileNeisWithAttendance')
    .addItem('✏️ 직접 수정한 AttendanceData 기준으로 Y:Z열 재계산', 'reconcileAttendanceDataSheetDirectly')
    .addToUi();
}

/**
 * 사용자가 AttendanceData 시트에서 직접 수정한 E~T열 수치(서류 건수)를 기반으로
 * U~X열(통계 합계) 및 Y:Z열(일치여부, 불일치내용)을 재계산하고 색상을 칠합니다.
 */
function reconcileAttendanceDataSheetDirectly(isSilent) {
  isSilent = (isSilent === true);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const neisSheet = ss.getSheetByName(CONFIG.SHEETS.NEIS || "NeisData");
  const dataSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_DATA || "AttendanceData");
  
  if (!neisSheet || !dataSheet) {
    if (!isSilent) {
      SpreadsheetApp.getUi().alert("NeisData 또는 AttendanceData 시트가 존재하지 않습니다.");
    } else {
      Logger.log("오류: NeisData 또는 AttendanceData 시트가 존재하지 않습니다.");
    }
    return;
  }
  
  const neisRows = neisSheet.getDataRange().getValues().slice(1);
  const dataRows = dataSheet.getDataRange().getValues();
  const headers = dataRows[0];
  const dataBody = dataRows.slice(1);
  
  if (dataBody.length === 0) {
    if (!isSilent) {
      SpreadsheetApp.getUi().alert("대조할 데이터가 AttendanceData 시트에 존재하지 않습니다.");
    } else {
      Logger.log("경고: 대조할 데이터가 AttendanceData 시트에 존재하지 않습니다.");
    }
    return;
  }
  
  // 1. NEIS 데이터 맵화 (Key: 학번_월)
  const neisMap = new Map();
  neisRows.forEach(row => {
    if (!row[1] || !row[2]) return;
    const month = parseInt(row[1]);
    const sid = String(row[2]).trim();
    neisMap.set(`${sid}_${month}`, row);
  });

  // Count experiential learning and menstrual pain documents, and gather date sets in AttendanceDB for each student_month
  const experientialCountMap = new Map();
  const menstrualCountMap = new Map(); // Key: studentId_month_category_subcategory
  const studentDatesMap = new Map();   // Key: studentId_month_category_subcategory -> Set of dates
  
  const settings = getSystemSettings();
  const examPeriods = getExamPeriods();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  
  if (attSheet) {
    const attData = attSheet.getDataRange().getValues().slice(1);
    attData.forEach((row, rowIndex) => {
      // G열 (시작일시) 우선, A열 (일자) 차선
      const dateValRaw = String(row[6] || "").trim();
      let dateVal = "";
      if (dateValRaw && dateValRaw.includes('-')) {
        dateVal = dateValRaw.split(' ')[0];
      }
      if (!dateVal || dateVal.split('-').length !== 3) {
        dateVal = row[0];
      }
      
      const date = parseDateSafe(dateVal);
      if (!date || isNaN(date.getTime())) return;
      
      const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
      const sid = String(row[1]).trim();
      const gradeNum = sid.length >= 3 ? parseInt(sid.charAt(0), 10) : -1;
      const cNum = sid.length >= 3 ? parseInt(sid.substring(1,3), 10) : -1;
      if (gradeNum !== settings.grade || cNum > settings.classes) return;

      const subCatRaw = String(row[4] || "").trim();   // E열: 대분류 (질병, 인정 등)
      const catRaw = String(row[5] || "").trim();      // F열: 소분류 (결석, 지각 등)
      const daysRaw = String(row[8] || "1").trim();    // I열: 일수
      const reasonDetail = String(row[9] || "").trim(); // J열: 사유상세
      
      let subCat = "기타";
      if (subCatRaw.includes("질병")) subCat = "질병";
      else if (subCatRaw.includes("미인정")) subCat = "미인정";
      else if (subCatRaw.includes("기타")) subCat = "기타";
      else if (subCatRaw.includes("출석인정") || subCatRaw.includes("인정")) subCat = "인정";

      let cat = "결석";
      if (catRaw.includes("결석")) cat = "결석";
      else if (catRaw.includes("지각")) cat = "지각";
      else if (catRaw.includes("조퇴")) cat = "조퇴";
      else if (catRaw.includes("결과")) cat = "결과";

      let days = 1;
      const match = daysRaw.match(/(\d+)/);
      if (match) days = parseInt(match[1], 10);

      const isExperiential = reasonDetail.includes("학교장허가체험학습") || 
                             reasonDetail.includes("현장체험") || 
                             reasonDetail.includes("교외체험") || 
                             reasonDetail.includes("체험학습");

      const coveredDates = [];
      try {
        const parts = dateStr.split('-');
        if (parts.length === 3) {
          const startDateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
          for (let d = 0; d < days; d++) {
            const currentDateObj = new Date(startDateObj.getTime() + d * 24 * 60 * 60 * 1000);
            coveredDates.push(Utilities.formatDate(currentDateObj, "GMT+9", "yyyy-MM-dd"));
          }
        } else {
          coveredDates.push(dateStr);
        }
      } catch (e) {
        coveredDates.push(dateStr);
      }

      coveredDates.forEach(dStr => {
        const parts = dStr.split('-');
        if (parts.length !== 3) return;
        const actualMonth = parseInt(parts[1], 10);
        
        const isExamDate = examPeriods.some(period => {
          if (!period.startStr) return false;
          return dStr >= period.startStr && dStr <= period.endStr;
        });

        const isMenstrual = subCat === "인정" && /생리/.test(reasonDetail) && !isExamDate;

        // 체험학습 건수 누적 (중복 방지)
        if (subCat === "인정" && isExperiential) {
          const expKey = `${sid}_${actualMonth}`;
          const isRowAlreadyCountedKey = `${expKey}_row_${rowIndex}`;
          if (!experientialCountMap.has(isRowAlreadyCountedKey)) {
            experientialCountMap.set(isRowAlreadyCountedKey, true);
            experientialCountMap.set(expKey, (experientialCountMap.get(expKey) || 0) + 1);
          }
        }

        // 생리통 건수 누적
        if (isMenstrual) {
          const mKey = `${sid}_${actualMonth}_${cat}_인정`;
          menstrualCountMap.set(mKey, (menstrualCountMap.get(mKey) || 0) + 1);
        }

        // 날짜 정보 수집
        if (!isMenstrual) {
          const dateKey = `${sid}_${actualMonth}_${cat}_${subCat}`;
          if (!studentDatesMap.has(dateKey)) {
            studentDatesMap.set(dateKey, new Set());
          }
          studentDatesMap.get(dateKey).add(dStr);
        }
      });
    });
  }
  
  const studentMap = getStudentDirectoryMap();
  
  // NEIS와 AttendanceData 열 매핑
  const mapping = [
    { col: 4, cat: "결석", sub: "질병" }, { col: 5, cat: "결석", sub: "미인정" }, { col: 6, cat: "결석", sub: "기타" }, { col: 7, cat: "결석", sub: "인정" },
    { col: 8, cat: "지각", sub: "질병" }, { col: 9, cat: "지각", sub: "미인정" }, { col: 10, cat: "지각", sub: "기타" }, { col: 11, cat: "지각", sub: "인정" },
    { col: 12, cat: "조퇴", sub: "질병" }, { col: 13, cat: "조퇴", sub: "미인정" }, { col: 14, cat: "조퇴", sub: "기타" }, { col: 15, cat: "조퇴", sub: "인정" },
    { col: 16, cat: "결과", sub: "질병" }, { col: 17, cat: "결과", sub: "미인정" }, { col: 18, cat: "결과", sub: "기타" }, { col: 19, cat: "결과", sub: "인정" }
  ];
  
  const updatedRows = [];
  const bgRows = [];
  
  dataBody.forEach((row) => {
    const rowMonth = parseInt(row[1]);
    const sid = String(row[2]).trim();
    
    // 만약 월이나 학번이 비어있으면 그대로 둠
    if (isNaN(rowMonth) || !sid) {
      updatedRows.push(row);
      bgRows.push(new Array(26).fill(null));
      return;
    }

    const gradeNum = sid.length >= 3 ? parseInt(sid.charAt(0), 10) : -1;
    const cNum = sid.length >= 3 ? parseInt(sid.substring(1,3), 10) : -1;
    if (gradeNum !== settings.grade || cNum > settings.classes) {
      updatedRows.push(row);
      bgRows.push(new Array(26).fill(null));
      return;
    }
    
    const neisRow = neisMap.get(`${sid}_${rowMonth}`);
    const studentInfo = studentMap.get(sid);
    const genderStr = studentInfo ? String(studentInfo.gender || "").trim().toLowerCase() : "";
    const isFemale = studentInfo && (genderStr === '여' || genderStr === '여자' || genderStr === 'f' || genderStr === 'female');
    
    // 시험기간 확인
    const hasExamInMonth = examPeriods.some(period => {
      if (!period.startStr) return false;
      const startMonth = parseInt(period.startStr.split('-')[1], 10);
      const endMonth = period.endStr ? parseInt(period.endStr.split('-')[1], 10) : startMonth;
      return startMonth === rowMonth || endMonth === rowMonth;
    });
    
    // 순번(0), 월(1), 학번(2), 성명(3) 복사
    const newRow = [row[0], row[1], row[2], row[3]];
    const bgRow = new Array(26).fill(null);
    
    let absTotal = 0;
    let tardyTotal = 0;
    let earlyTotal = 0;
    let skipTotal = 0;
    
    let mismatches = [];
    let isMatch = true;
    
    // E(4) ~ T(19)열 데이터 검사 및 비교
    mapping.forEach((m) => {
      const localVal = parseInt(row[m.col]) || 0; // 사용자가 직접 수정한 로컬 값 읽기
      newRow.push(localVal); // 그대로 저장
      
      if (m.cat === "결석" && m.sub !== "인정") absTotal += localVal;
      else if (m.cat === "지각" && m.sub !== "인정") tardyTotal += localVal;
      else if (m.cat === "조퇴" && m.sub !== "인정") earlyTotal += localVal;
      else if (m.cat === "결과" && m.sub !== "인정") skipTotal += localVal;
      
      let neisVal = 0;
      if (neisRow) {
        neisVal = parseInt(neisRow[m.col]) || 0;
      }
      
      let isMismatch = neisVal !== localVal;
      
      // 미인정 예외
      if (isMismatch && m.sub === "미인정") {
        isMismatch = false;
      }
      
      // 생리통 예외
      if (isMismatch && m.sub === "인정" && isFemale) {
        const mKey = `${sid}_${rowMonth}_${m.cat}_인정`;
        const menstrualCount = menstrualCountMap.get(mKey) || 0;
        if (neisVal > localVal) {
          const diff = neisVal - localVal;
          const limit = (m.cat === "결석") ? 1 : 3;
          if (diff + menstrualCount <= limit) {
            isMismatch = false;
          }
        } else if (neisVal < localVal) {
          const diff = localVal - neisVal;
          if (diff <= menstrualCount) {
            isMismatch = false;
          }
        }
      }
      
      // School Experiential Learning (학교장허가체험학습) Exception:
      // If category is "인정", and there are at least 2 scanned documents (신청서 + 결과보고서)
      // marked as "학교장허가체험학습" in the database, suppress mismatch error regardless of the day difference.
      const expCount = experientialCountMap.get(`${sid}_${rowMonth}`) || 0;
      if (isMismatch && m.sub === "인정" && expCount >= 2) {
        isMismatch = false;
      }
      
      if (isMismatch) {
        isMatch = false;
        const dateKey = `${sid}_${rowMonth}_${m.cat}_${m.sub}`;
        const localDatesSet = studentDatesMap.get(dateKey);
        const localDates = localDatesSet ? Array.from(localDatesSet) : [];
        const dateSuffix = localDates.length > 0 ? ` [보관서류 날짜: ${localDates.join(", ")}]` : " [보관서류 날짜: 없음]";
        
        mismatches.push(`${m.cat}(${m.sub}) 나이스:${neisVal} / 서류:${localVal}${dateSuffix}`);
        bgRow[m.col] = '#fff2cc'; // Mismatch cell yellow background
      }
    });
    
    // U~X열 (결석, 지각, 조퇴, 결과 총계) 수식 업데이트
    newRow.push(absTotal, tardyTotal, earlyTotal, skipTotal);
    
    // Y: 일치여부, Z: 불일치내용
    newRow.push(isMatch ? "✅ 일치" : "❌ 불일치", isMatch ? "-" : mismatches.join("\n"));
    
    if (!isMatch) {
      bgRow[24] = '#fce5cd'; // Y열 배경색
    }
    
    updatedRows.push(newRow);
    bgRows.push(bgRow);
  });
  
  // 시트에 일괄 저장
  const targetRange = dataSheet.getRange(2, 1, updatedRows.length, 26);
  targetRange.setValues(updatedRows);
  targetRange.setBackgrounds(bgRows);
  
  if (!isSilent) {
    SpreadsheetApp.getUi().alert("🎉 수동 수정된 수치를 기반으로 Y:Z열 대조 재계산이 완료되었습니다!");
  }
}

function runDiagnosticForStudent(name) {
  name = name || "권예빈";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const rows = attSheet.getDataRange().getValues();
  const matched = [];
  rows.forEach((row, i) => {
    if (row[2] && String(row[2]).includes(name)) {
      matched.push({
        rowNum: i + 1,
        date: row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT+9", "yyyy-MM-dd") : String(row[0]),
        sid: row[1],
        name: row[2],
        category: row[4],
        subCategory: row[5],
        start: row[6],
        end: row[7],
        days: row[8],
        neisString: row[21]
      });
    }
  });
  Logger.log(JSON.stringify(matched, null, 2));
  return JSON.stringify(matched, null, 2);
}

/**
 * Validate attendance document rules (retrieved from AI metadata & custom calculations)
 */
function checkAttendanceRule(item) {
  let ruleCheck = item.ruleCheck || "";

  // 1. 카테고리가 "출석인정"인 경우 경조사 규정 검증 수행
  const isApproved = (item.category && item.category.includes("출석인정"));
  if (isApproved) {
    const reason = item.reasonDetail || "";
    const eventRules = getFamilyEventRules(reason);
    if (eventRules) {
      const startYear = 2025;
      const endYear = 2028;
      const holidays = getHolidaysSet(startYear, endYear);
      const schoolHolidays = getDiscretionaryHolidays();
      schoolHolidays.forEach(h => holidays.add(h));

      // 시작/종료일의 날짜 부분 추출
      const startDateStr = item.periodStart ? item.periodStart.split(' ')[0] : "";
      const endDateStr = item.periodEnd ? item.periodEnd.split(' ')[0] : "";
      
      // 실제 소요 학업일수 계산
      const schoolDaysTaken = countSchoolDays(startDateStr, endDateStr, holidays);
      
      const errors = [];
      
      // 규정 1: 경조사 허용일수 초과 검증
      if (schoolDaysTaken > eventRules.allowedDays) {
        errors.push(`허용일수 초과 (허용 ${eventRules.allowedDays}일 / 신청 ${schoolDaysTaken}일)`);
      }
      
      // 규정 2: 발생일 동반 검증 (docStartDate가 발생일)
      const eventDateStr = item.docStartDate ? item.docStartDate.split(' ')[0] : "";
      if (eventDateStr && startDateStr) {
        const isValidStart = isValidFamilyEventStartDate(eventDateStr, startDateStr, eventRules.isDeath, holidays);
        if (!isValidStart) {
          errors.push(`시작일 규정 위반 (발생일: ${eventDateStr}, 신청일: ${startDateStr})`);
        }
      }
      
      if (errors.length > 0) {
        ruleCheck = "체크필요 [" + errors.join(", ") + "]";
      } else {
        ruleCheck = ""; // 정상 준수 시 빈칸
      }
    }
  }

  return ruleCheck;
}

/**
 * Classifies family events (경조사) and returns the rule settings
 */
function getFamilyEventRules(reason) {
  if (!reason) return null;
  const text = String(reason).replace(/\s+/g, ""); // 공백 제거하여 매칭

  // 2. 학생 본인의 입양: 20일
  if (text.includes("입양")) {
    return { event: "입양", allowedDays: 20, isDeath: false };
  }

  // 3-6. 사망 (사망, 부고, 별세, 장례, 상 등)
  const isDeath = /사망|부고|별세|장례|상$|상[^장]|국장|국민장/.test(reason) || /상구|상제|상주|영결식|발인/.test(reason);
  if (isDeath) {
    // 4) 부모의 조부모, 외조부모, 증조부모, 외증조부모, 진외증조부모, 외외증조부모의 사망: 3일
    // "부모의 조부모", "부모의 외조부모" 및 증조/외증조 등은 조부모/외조부모 매칭 전에 선점해야 함
    if (/부모의조부모|부모의외조부모/.test(text)) {
      return { event: "부모의 조부모/외조부모 사망", allowedDays: 3, isDeath: true };
    }
    if (/증조|외증조|진외증조|외외증조/.test(text)) {
      return { event: "증조부모/외증조부모 등 사망", allowedDays: 3, isDeath: true };
    }
    // 3) 부모, 조부모, 외조부모의 사망: 5일
    if (/부모|조부모|외조부모|친조부|친조모|외조부|외조모|할아버지|할머니|외할아버지|외할머니|부친|모친|아버지|어머니|조모|조부|외조모|외조부/.test(text)) {
      return { event: "부모/조부모/외조부모 사망", allowedDays: 5, isDeath: true };
    }
    // 5) 형제, 자매, 형제의 배우자, 자매의 배우자 사망: 3일
    if (/형제|자매|남동생|여동생|오빠|형[^기]|누나|언니|형수|제수|매형|매제|올케/.test(text)) {
      return { event: "형제/자매 또는 그 배우자 사망", allowedDays: 3, isDeath: true };
    }
    // 6) 부모의 형제, 부모의 자매, 부모의 형제의 배우자, 부모의 자매의 배우자 사망: 3일
    if (/부모의형제|부모의자매|삼촌|백부|숙부|고모|이모|외삼촌|백모|숙모|고모부|이모부|외숙모/.test(text)) {
      return { event: "부모의 형제/자매 또는 그 배우자 사망", allowedDays: 3, isDeath: true };
    }
  }

  // 1) 부, 모, 형제, 자매의 결혼: 1일
  if (/결혼|혼인|웨딩/.test(reason)) {
    if (/부|모|형제|자매|형재|오빠|형[^기]|누나|언니|남동생|여동생/.test(text)) {
      return { event: "부/모/형제/자매의 결혼", allowedDays: 1, isDeath: false };
    }
  }

  // 7) 조부모의 재혼, 회갑, 칠순, 팔순: 1일
  if (/재혼|회갑|칠순|팔순|환갑|고희/.test(reason)) {
    if (/조부모|외조부모|할아버지|할머니|외할아버지|외할머니|친조부|친조모|외조부|외조모|부모|아버지|어머니|부|모/.test(text)) {
      return { event: "조부모 재혼 또는 부모/조부모 회갑/칠순/팔순", allowedDays: 1, isDeath: false };
    }
  }

  return null;
}

/**
 * South Korean holidays list fetch & cache (with fallback for 2026/2027)
 */
function getHolidaysSet(startYear, endYear) {
  const holidays = new Set();
  try {
    const cal = CalendarApp.getCalendarById("ko.south_korean#holiday@group.v.calendar.google.com");
    if (cal) {
      const events = cal.getEvents(new Date(startYear, 0, 1), new Date(endYear, 11, 31));
      events.forEach(ev => {
        const d = ev.getStartTime();
        const end = ev.getEndTime();
        let cur = new Date(d);
        while (cur < end) {
          holidays.add(Utilities.formatDate(cur, "GMT+9", "yyyy-MM-dd"));
          cur.setDate(cur.getDate() + 1);
        }
      });
    }
  } catch (e) {
    Logger.log("Failed to fetch holidays from Google Calendar: " + e.toString());
  }
  
  // 2026년 공휴일 및 대체공휴일 수동 추가 (백업)
  const hardcoded2026 = [
    "2026-01-01",
    "2026-02-16", "2026-02-17", "2026-02-18", // 설날
    "2026-03-01", "2026-03-02", // 삼일절 + 대체공휴일
    "2026-05-05", // 어린이날
    "2026-05-24", "2026-05-25", // 부처님오신날 + 대체공휴일
    "2026-06-06", // 현충일
    "2026-08-15", "2026-08-17", // 광복절 + 대체공휴일
    "2026-09-24", "2026-09-25", "2026-09-26", "2026-09-28", // 추석 + 대체공휴일
    "2026-10-03", "2026-10-05", // 개천절 + 대체공휴일
    "2026-10-09", // 한글날
    "2026-12-25" // 성탄절
  ];
  
  // 2027년 공휴일 및 대체공휴일 수동 추가 (백업)
  const hardcoded2027 = [
    "2027-01-01",
    "2027-02-05", "2027-02-06", "2027-02-07", "2027-02-08", // 설날 + 대체공휴일
    "2027-03-01", // 삼일절
    "2027-05-05", // 어린이날
    "2027-05-13", // 부처님오신날
    "2027-06-06", // 현충일
    "2027-08-15", "2027-08-16", // 광복절 + 대체공휴일
    "2027-09-14", "2027-09-15", "2027-09-16", // 추석
    "2027-10-03", "2027-10-04", // 개천절 + 대체공휴일
    "2027-10-09", "2027-10-11", // 한글날 + 대체공휴일
    "2027-12-25", "2027-12-27" // 성탄절 + 대체공휴일
  ];

  hardcoded2026.forEach(h => holidays.add(h));
  hardcoded2027.forEach(h => holidays.add(h));
  
  return holidays;
}

/**
 * Reads discretionary school holidays (재량휴업일) from a sheet named '재량휴업일' or 'SchoolHolidays'
 */
function getDiscretionaryHolidays() {
  const holidays = new Set();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("SchoolHolidays") || ss.getSheetByName("재량휴업일");
    if (sheet) {
      const data = sheet.getDataRange().getValues().slice(1);
      data.forEach(row => {
        if (row[0]) {
          const d = parseDateSafe(row[0]);
          if (d) holidays.add(Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd"));
        }
      });
    }
  } catch (e) {
    Logger.log("Failed to read discretionary holidays sheet: " + e.toString());
  }
  return holidays;
}

/**
 * Checks if a date is a Saturday, Sunday, or public holiday
 */
function isHolidayOrWeekend(date, holidaysCached) {
  const day = date.getDay();
  if (day === 0 || day === 6) return true; // 토요일, 일요일
  
  const dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  if (holidaysCached && holidaysCached.has(dateStr)) return true;
  
  return false;
}

/**
 * Counts actual school days between two dates, excluding weekends and holidays
 */
function countSchoolDays(startDateStr, endDateStr, holidaysCached) {
  const start = new Date(startDateStr);
  const end = new Date(endDateStr || startDateStr);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return 1;

  let count = 0;
  let cur = new Date(start);
  while (cur <= end) {
    if (!isHolidayOrWeekend(cur, holidaysCached)) {
      count++;
    }
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

/**
 * Validates family event start date relative to the event occurrence date
 */
function isValidFamilyEventStartDate(eventDateStr, startDateStr, isDeath, holidaysCached) {
  if (!eventDateStr || !startDateStr) return true; // 발생일/시작일이 없으면 판단 보류

  const evDate = new Date(eventDateStr);
  const stDate = new Date(startDateStr);
  if (isNaN(evDate.getTime()) || isNaN(stDate.getTime())) return true;

  const evDateStr = Utilities.formatDate(evDate, "GMT+9", "yyyy-MM-dd");
  const stDateStr = Utilities.formatDate(stDate, "GMT+9", "yyyy-MM-dd");

  // 1. 발생 당일 시작하면 정상
  if (stDateStr === evDateStr) return true;

  // 2. 사망의 경우 다음날부터 시작 가능
  const nextDay = new Date(evDate);
  nextDay.setDate(evDate.getDate() + 1);
  const nextDayStr = Utilities.formatDate(nextDay, "GMT+9", "yyyy-MM-dd");
  if (isDeath && stDateStr === nextDayStr) return true;

  // 3. 발생일(또는 다음날)이 공휴일/주말인 경우, 그 다음 첫 학업일(school day)부터 실시 가능
  let firstSchoolDayAfterEv = new Date(evDate);
  do {
    firstSchoolDayAfterEv.setDate(firstSchoolDayAfterEv.getDate() + 1);
  } while (isHolidayOrWeekend(firstSchoolDayAfterEv, holidaysCached));
  
  const firstSchoolDayAfterEvStr = Utilities.formatDate(firstSchoolDayAfterEv, "GMT+9", "yyyy-MM-dd");
  if (stDateStr === firstSchoolDayAfterEvStr) return true;

  if (isDeath) {
    let firstSchoolDayAfterNext = new Date(nextDay);
    do {
      firstSchoolDayAfterNext.setDate(firstSchoolDayAfterNext.getDate() + 1);
    } while (isHolidayOrWeekend(firstSchoolDayAfterNext, holidaysCached));
    
    const firstSchoolDayAfterNextStr = Utilities.formatDate(firstSchoolDayAfterNext, "GMT+9", "yyyy-MM-dd");
    if (stDateStr === firstSchoolDayAfterNextStr) return true;
  }

  return false;
}

/**
 * Retrieve System Settings (Grade and Classes) from PropertiesService
 */
function getSystemSettings() {
  const props = PropertiesService.getScriptProperties();
  const grade = props.getProperty('SYSTEM_GRADE') || '3';
  const classes = props.getProperty('SYSTEM_CLASSES') || '11';
  return {
    grade: parseInt(grade, 10),
    classes: parseInt(classes, 10)
  };
}

/**
 * Save System Settings (Grade and Classes) to PropertiesService
 */
function saveSystemSettings(grade, classes) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SYSTEM_GRADE', String(grade));
    props.setProperty('SYSTEM_CLASSES', String(classes));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Retrieve names of students in the target grade from StudentDirectory
 */
function getTargetGradeStudentNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);
    if (!sheet) return [];

    const settings = getSystemSettings();
    const data = sheet.getDataRange().getValues().slice(1);
    const studentNames = [];
    data.forEach(row => {
      if (!row[0] || !row[3]) return;
      const grade = parseInt(row[0], 10);
      if (grade === settings.grade) {
        const name = String(row[3]).trim();
        if (name && !studentNames.includes(name)) {
          studentNames.push(name);
        }
      }
    });
    return studentNames;
  } catch (e) {
    Logger.log("getTargetGradeStudentNames 오류: " + e.toString());
    return [];
  }
}

/**
 * Get monthly usage of menstrual absence (결석) and partials (지각/조퇴/결과)
 * for a specific student in a specific month.
 */
function getMonthlyMenstrualUsage(studentId, yearMonthStr, priorBatchItems) {
  let absenceCount = 0;
  let partialCount = 0;

  // 1. Check in the database
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (sheet) {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    for (let i = 1; i < values.length; i++) {
      const rDate = String(values[i][0] || "").trim();
      const rId = String(values[i][1] || "").trim();
      const rCat = String(values[i][4] || "").trim();
      const rSub = String(values[i][5] || "").trim();
      const rReason = String(values[i][9] || "").trim();

      if (rId === String(studentId).trim() && rDate.startsWith(yearMonthStr) && rCat.includes("출석인정") && rReason.includes("생리")) {
        if (rSub.includes("결석")) {
          let daysNum = 1;
          const daysRaw = String(values[i][8] || "1").trim();
          const daysMatch = daysRaw.match(/(\d+)/);
          if (daysMatch) daysNum = parseInt(daysMatch[1], 10);
          absenceCount += daysNum;
        } else if (rSub.includes("지각") || rSub.includes("조퇴") || rSub.includes("결과")) {
          partialCount += 1;
        }
      }
    }
  }

  // 2. Check in the currently processing batch (prior items only)
  if (priorBatchItems && Array.isArray(priorBatchItems)) {
    for (let i = 0; i < priorBatchItems.length; i++) {
      const item = priorBatchItems[i];
      if (!item) continue;
      const rId = String(item.studentId || "").trim();
      const rDate = String(item.date || item.periodStart || "").trim();
      const rCat = String(item.category || "").trim();
      const rSub = String(item.subCategory || "").trim();
      const rReason = String(item.reasonDetail || "").trim();

      if (rId === String(studentId).trim() && rDate.startsWith(yearMonthStr) && rCat.includes("출석인정") && rReason.includes("생리")) {
        if (rSub.includes("결석")) {
          let daysNum = 1;
          const daysRaw = String(item.totalDays || "1").trim();
          const daysMatch = daysRaw.match(/(\d+)/);
          if (daysMatch) daysNum = parseInt(daysMatch[1], 10);
          absenceCount += daysNum;
        } else if (rSub.includes("지각") || rSub.includes("조퇴") || rSub.includes("결과")) {
          partialCount += 1;
        }
      }
    }
  }

  return { absenceCount, partialCount };
}
