/**
 * Attendance Management System - Configuration
 */

const CONFIG = {
  // Gemini API Key (스크립트 속성 'GEMINI_API_KEY'에서 안전하게 로드)
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '', 
  
  // 사용할 Gemini 모델명
  GEMINI_MODEL: 'gemini-3.5-flash',
  
  // 구글 드라이브 스캔본 업로드 폴더 ID (스크립트 속성 'UPLOAD_FOLDER_ID'에서 로드)
  UPLOAD_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('UPLOAD_FOLDER_ID') || '', 
  
  // NEIS 엑셀 파일(.xlsx) 업로드 폴더 ID (스크립트 속성 'NEIS_FOLDER_ID'에서 로드)
  NEIS_FOLDER_ID: PropertiesService.getScriptProperties().getProperty('NEIS_FOLDER_ID') || '', 
  
  // 시트명 설정
  SHEETS: {
    ATTENDANCE: 'AttendanceDB',
    NEIS: 'NeisData',
    STUDENTS: 'StudentDirectory',
    LOG: 'SystemLog',
    ATTENDANCE_DATA: 'AttendanceData',
    EXAM_PERIODS: 'ExamPeriods'
  }
};
