/**
 * 🛠️ 1회성 데이터 보정 도구: A~U열 텍스트 자동 교정
 * E열, F열, J열 분류뿐만 아니라 날짜, 서명, 규정위반 여부 등을 일괄 정리합니다.
 * 실행 방법: 편집기 상단에서 'fixLegacyAttendanceData' 함수를 선택하고 [실행] 버튼 클릭
 */
function fixLegacyAttendanceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("AttendanceDB 시트를 찾을 수 없습니다.");
    return;
  }
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // 날짜 변환 헬퍼 (YYYY-MM-DD 형식으로 변환)
  const formatDateOnly = (val) => {
     if (!val) return "";
     if (val instanceof Date) {
        return Utilities.formatDate(val, "GMT+9", "yyyy-MM-dd");
     }
     const str = String(val).trim();
     const m = str.match(/(\d{4})[-.](\d{1,2})[-.](\d{1,2})/);
     if (m) {
        return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
     }
     return str;
  };

  // 서명 변환 헬퍼 (무, X, 서명 안됨 등 -> × / 그 외 O, Signed, 빈칸, 학생이름 등 -> 빈칸)
  const formatSign = (val) => {
     if (!val) return ""; 
     const s = String(val).trim().toLowerCase();
     if (/무|x|×|안됨|없음|없다/i.test(s)) return '×';
     return ""; // O, Signed, Yes, 사인있음 등은 모두 서명된 것으로 간주하여 빈칸 처리
  };

  // 질병 연속성 체크용 Map (학생 학번 기준)
  // studentId -> { date: Date 객체, count: 연속 일수 }
  const diseaseState = new Map();
  
  // 첫 번째 행은 헤더이므로 제외
  for (let i = 1; i < values.length; i++) {
    // 1. A열 (일자)에서 날짜와 교시 분리
    let aVal = String(values[i][0] || "").trim();
    let aPeriodMatch = aVal.match(/(\d+)\s*교시/);
    let aPeriod = aPeriodMatch ? `${parseInt(aPeriodMatch[1], 10)}교시` : "";
    values[i][0] = formatDateOnly(values[i][0]); // A열은 날짜만 남김
    
    // 2. D열, L열, N열 (서명) 처리
    values[i][3] = formatSign(values[i][3]); // D열 학생서명
    values[i][11] = formatSign(values[i][11]); // L열 학부모서명
    values[i][13] = formatSign(values[i][13]); // N열 담임서명
    
    // 3. E열, F열, J열 정리 (기존 로직)
    let eVal = String(values[i][4] || "").trim();
    let fVal = String(values[i][5] || "").trim();
    let jVal = String(values[i][9] || "").trim();
    
    let cat = eVal;
    let sub = fVal;
    let reason = jVal;
    
    const combined = eVal + " " + fVal;
    let extractedSub = "";
    if (combined.includes("결석")) extractedSub = "결석";
    else if (combined.includes("지각")) extractedSub = "지각";
    else if (combined.includes("조퇴")) extractedSub = "조퇴";
    else if (combined.includes("결과")) extractedSub = "결과";
    
    let extractedCat = "";
    if (combined.includes("질병")) extractedCat = "질병";
    else if (combined.includes("인정") || combined.includes("출석인정") || combined.includes("체험") || combined.includes("생리")) extractedCat = "출석인정";
    else if (combined.includes("미인정") || combined.includes("무단")) extractedCat = "미인정";
    else if (combined.includes("기타")) extractedCat = "기타";
    
    const cleanTerms = (text) => text.replace(/결석|지각|조퇴|결과|질병|출석인정|기타|미인정/g, "").replace(/[\[\]\(\)]/g, "").trim();
    if (fVal && !["결석", "지각", "조퇴", "결과"].includes(fVal)) {
       const cleanedF = cleanTerms(fVal);
       if (cleanedF.length > 0 && !reason.includes(cleanedF)) reason = (cleanedF + " " + reason).trim();
    }
    if (eVal && !["질병", "출석인정", "기타", "미인정"].includes(eVal)) {
       const cleanedE = cleanTerms(eVal);
       if (cleanedE.length > 0 && !reason.includes(cleanedE)) reason = (cleanedE + " " + reason).trim();
    }
    
    cat = extractedCat || cat || "기타";
    if (cat === "인정") cat = "출석인정";
    sub = extractedSub || sub || "결석";
    
    // 용어 통일
    reason = reason.replace(/(학교장허가\s*|현장\s*|교외\s*)*체험\s*(학습)?/g, "학교장허가체험학습");
    
    values[i][4] = cat;
    values[i][5] = sub;
    values[i][9] = reason;
    
    // 4. G열 (시작일시), H열 (종료일시) 표준화
    let gVal = String(values[i][6] || "").trim();
    let hVal = String(values[i][7] || "").trim();
    
    // G열
    let gPeriodMatch = gVal.match(/(\d+)\s*교시/);
    let gPeriod = gPeriodMatch ? `${parseInt(gPeriodMatch[1], 10)}교시` : "";
    let gDate = formatDateOnly(values[i][6]);
    if (!gPeriod && aPeriod && gDate === values[i][0]) gPeriod = aPeriod; // A열에 교시가 있었고 날짜가 같다면 G로 가져옴
    values[i][6] = gDate ? (gDate + (gPeriod ? ` ${gPeriod}` : "")) : "×";
    if (values[i][6].trim() === "") values[i][6] = "×";
    
    // H열
    let hPeriodMatch = hVal.match(/(\d+)\s*교시/);
    let hPeriod = hPeriodMatch ? `${parseInt(hPeriodMatch[1], 10)}교시` : "";
    let hDate = formatDateOnly(values[i][7]);
    values[i][7] = hDate ? (hDate + (hPeriod ? ` ${hPeriod}` : "")) : "×";
    if (values[i][7].trim() === "") values[i][7] = "×";
    
    // 5. U열 (입력일시) YYYY-MM-DD 포맷
    values[i][20] = formatDateOnly(values[i][20]);
    
    // 6. T열 (규정위반여부)
    let tVal = "";
    const docStr = String(values[i][14] || "").trim(); // O열 첨부서류
    const sid = String(values[i][1]); // B열 학번
    
    if (cat === "질병") {
        const hasHospital = /진료|진단|소견|처방|통원|입원|병원|약국|약|영수증/.test(docStr);
        const hasSubstitute = /학부모|담임|확인서|의견서|서명/.test(docStr);
        
        const curDateObj = new Date(values[i][0]); // A열 기준
        let state = diseaseState.get(sid) || { date: null, count: 0 };
        
        let isConsecutive = false;
        if (state.date) {
           const diffTime = curDateObj.getTime() - state.date.getTime();
           const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
           // 주말 포함하여 최대 4일 이내면 연속 결석으로 간주
           if (diffDays > 0 && diffDays <= 4) {
               isConsecutive = true;
           }
        }
        
        if (isConsecutive) {
           state.count++;
        } else {
           state.count = 1;
        }
        state.date = curDateObj;
        diseaseState.set(sid, state);
        
        if (state.count % 2 === 1) { // 1, 3, 5일차
           if (!hasHospital) tVal = "체크필요 [질병 결석(홀수일차)은 병원/약국 서류 필수]";
        } else { // 2, 4, 6일차
           if (!hasHospital && !hasSubstitute) tVal = "체크필요 [질병 결석(짝수일차)은 학부모확인서 또는 병원서류 필수]";
        }
    } else if (cat !== "미인정" && !reason.includes("생리")) {
        // 출석인정, 기타결석 등 (미인정과 생리결석 제외)
        if (docStr === "×" || docStr === "") {
            tVal = "체크필요 [질병 외 사유(출석인정 등) 관련 첨부서류 누락 의심]";
        }
    }
    
    values[i][19] = tVal;
    
    // 7. 'None', 'null', '첨부서류 누락', '해당없음' 이라는 텍스트가 들어간 셀을 모두 빈칸으로 정리
    for (let c = 0; c < values[i].length; c++) {
       if (typeof values[i][c] === 'string' && /^(none|null|(해당\s*)?없음|첨부서류\s*누락)$/i.test(values[i][c].trim())) {
           values[i][c] = "";
       }
    }
    
    // 8. I열 (일수) 괄호 제거
    if (typeof values[i][8] === 'string') {
       values[i][8] = values[i][8].replace(/[\(\)]/g, "").trim();
    }
  }
  
  dataRange.setValues(values);
  SpreadsheetApp.getUi().alert("날짜, 교시, 서명, 질병 결석 반복 규정 등 모든 항목의 일괄 보정이 완료되었습니다!");
}
