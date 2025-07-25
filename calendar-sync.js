//ㄴㄷㄹㄴ아ㅓ롸너우하ㅓㅜㅏㅓ


/** 
 * ==================== 전역 설정 구성 ====================
 * Google Apps Script 캘린더 동기화 시스템의 핵심 설정값들을 관리합니다.
 * 이 설정을 통해 시트 구조, 캘린더 정보, 트리거 조건 등을 중앙에서 관리할 수 있습니다.
 */
const CALENDAR_SYNC_CONFIG = {
  // 연동할 구글 캘린더의 고유 ID - 캘린더 설정에서 확인 가능
  TARGET_CALENDAR_ID: '62feee307ee8a80b5ffa58a66f540833f146d3edee4bda3bd6da5326884d3a36@group.calendar.google.com',
  
  // 스프레드시트의 시트 이름들을 정의 - 시트 이름 변경 시 여기서만 수정하면 됨
  SHEET_NAMES: {
    SCHEDULE_SHEET: '수업표',        // 수업 일정이 입력되는 메인 시트
    STUDENT_DATABASE: '학생정보DB'    // 학생 상세 정보가 저장된 데이터베이스 시트
  },
  
  // 수업표 시트의 컬럼 인덱스 정의 - 컬럼 순서 변경 시 여기서만 수정
  SCHEDULE_COLUMNS: {
    TRIGGER_COMMAND: 1,    // A열: '캘린더이동' 명령어 입력 컬럼
    STUDENT_NAME: 2,       // B열: 학생 이름 (학생정보DB와 매칭되는 키값)
    LESSON_CONTENT: 3,     // C열: 수업 내용 (현재 미사용)
    REPEAT_FREQUENCY: 4,   // D열: 반복 주기 (예: '매주 월요일')
    LESSON_TIME: 5,        // E열: 수업 시간 정보
    SYNC_STATUS: 14,       // N열: 동기화 상태 ('전송완료' 또는 '전송실패')
    EVENT_ID: 15          // O열: 생성된 이벤트의 고유 식별자
  },
  
  // 학생정보DB 시트의 컬럼 인덱스 정의
  STUDENT_DB_COLUMNS: {
    STUDENT_NAME: 4,       // D열: 학생 이름 (수업표 시트와 매칭)
    ADDITIONAL_INFO_F: 6,  // F열: 추가 정보 1 (캘린더 이벤트 설명에 포함)
    ADDITIONAL_INFO_G: 7,  // G열: 추가 정보 2 (캘린더 이벤트 설명에 포함)
    ADDITIONAL_INFO_M: 13, // M열: 추가 정보 3 (캘린더 이벤트 설명에 포함)
    ADDITIONAL_INFO_N: 14  // N열: 추가 정보 4 (캘린더 이벤트 설명에 포함)
  },
  
  // 동기화를 실행하는 트리거 명령어
  SYNC_TRIGGER_WORD: '캘린더이동',
  
  // 기본 이벤트 시간 설정 (시간 데이터가 없을 경우 사용)
  DEFAULT_EVENT_TIME: {
    START_HOUR: 9,    // 기본 시작 시간: 오전 9시
    END_HOUR: 10      // 기본 종료 시간: 오전 10시 (1시간 수업)
  }
};

/**
 * ==================== 시트 편집 감지 함수 (Simple Trigger) ====================
 * 스프레드시트에서 셀이 편집될 때 자동으로 실행되는 함수입니다.
 * Simple Trigger는 권한이 제한되어 있어 캘린더 API 접근이 불가능하므로
 * 편집 감지만 하고 실제 동기화는 Installable Trigger에서 처리합니다.
 */
function detectSheetEdit(editEvent) {
  try {
    // 편집 이벤트에서 필요한 정보 추출
    const editedSheet = editEvent.source.getActiveSheet();
    const editedRange = editEvent.range;
    const editedValue = editEvent.value;
    const editedColumn = editedRange.getColumn();
    const editedRow = editedRange.getRow();
    
    // A열에 캘린더 동기화 트리거 명령어가 입력되었는지 확인
    const isTriggerColumn = editedColumn === CALENDAR_SYNC_CONFIG.SCHEDULE_COLUMNS.TRIGGER_COMMAND;
    const isTriggerWord = editedValue === CALENDAR_SYNC_CONFIG.SYNC_TRIGGER_WORD;
    
    if (isTriggerColumn && isTriggerWord) {
      console.log(`[편집 감지] ${editedRow}행에 "${CALENDAR_SYNC_CONFIG.SYNC_TRIGGER_WORD}" 입력 감지됨`);
      console.log('[편집 감지] Installable Trigger에서 실제 캘린더 동기화가 실행됩니다.');
    }
    
  } catch (error) {
    console.error('[편집 감지 오류] 시트 편집 감지 중 예외 발생:', error);
    console.error('[편집 감지 오류] 오류 상세:', error.stack);
  }
}

// 기존 함수명과의 호환성을 위한 별칭
const onEdit = detectSheetEdit;

/**
 * ==================== 캘린더 동기화 메인 실행 함수 (Installable Trigger) ====================
 * Installable Trigger로 등록되어 실제 캘린더 API 권한을 가지고 동기화를 수행하는 함수입니다.
 * 사용자가 A열에 '캘린더이동'을 입력하면 이 함수가 트리거되어 전체 동기화 프로세스를 실행합니다.
 */
function executeCalendarSynchronization(editEvent) {
  try {
    // 편집 이벤트에서 상세 정보 추출
    const editedSheet = editEvent.source.getActiveSheet();
    const editedRange = editEvent.range;
    const editedValue = editEvent.value;
    const sheetName = editedSheet.getName();
    const editedRow = editedRange.getRow();
    const editedColumn = editedRange.getColumn();
    
    // 동기화 실행 조건 로깅
    console.log(`[동기화 트리거] 편집 상세 - 시트: ${sheetName}, 행: ${editedRow}, 열: ${editedColumn}, 값: "${editedValue}"`);
    
    // 동기화 실행 조건 검증
    const isScheduleSheet = sheetName === CALENDAR_SYNC_CONFIG.SHEET_NAMES.SCHEDULE_SHEET;
    const isTriggerColumn = editedColumn === CALENDAR_SYNC_CONFIG.SCHEDULE_COLUMNS.TRIGGER_COMMAND;
    const isTriggerWord = editedValue === CALENDAR_SYNC_CONFIG.SYNC_TRIGGER_WORD;
    
    if (isScheduleSheet && isTriggerColumn && isTriggerWord) {
      console.log(`[동기화 시작] ${editedRow}행에서 캘린더 동기화 프로세스를 시작합니다.`);
      
      // 실제 캘린더 이벤트 생성 및 동기화 실행
      processScheduleRowToCalendarEvent(editedSheet, editedRow);
      
    } else {
      // 조건에 맞지 않는 경우 상세 로깅 (디버깅용)
      console.log(`[동기화 조건 불일치] 시트명: ${sheetName} (${isScheduleSheet ? '일치' : '불일치'}), ` +
                  `컬럼: ${editedColumn} (${isTriggerColumn ? '일치' : '불일치'}), ` +
                  `값: "${editedValue}" (${isTriggerWord ? '일치' : '불일치'})`);
    }
    
  } catch (error) {
    console.error('[동기화 실행 오류] 캘린더 동기화 실행 중 예외 발생:', error);
    console.error('[동기화 실행 오류] 오류 상세:', error.stack);
    
    // 사용자에게 오류 상황 알림 (가능한 경우)
    if (editEvent && editEvent.source) {
      try {
        SpreadsheetApp.getUi().alert('캘린더 동기화 중 오류가 발생했습니다. 로그를 확인해주세요.');
      } catch (uiError) {
        console.error('[UI 알림 실패] 사용자 알림 표시 중 오류:', uiError);
      }
    }
  }
}

// 기존 함수명과의 호환성을 위한 별칭
const processCalendarSync = executeCalendarSynchronization;

/**
 * ==================== 동기화 결과 업데이트 함수 ====================
 * 수업표 시트의 특정 행에 동기화 상태와 이벤트 ID를 기록합니다.
 * 
 * @param {Sheet} sheet - 수업표 시트 객체
 * @param {number} row - 업데이트할 행 번호
 * @param {string} status - 동기화 상태 ('전송완료' 또는 '전송실패')
 * @param {string} eventId - 생성된 이벤트의 고유 식별자 또는 오류 메시지
 */
function updateSyncResult(sheet, row, status, eventId) {
  try {
    // N열(14)에 상태, O열(15)에 이벤트 ID 기록
    sheet.getRange(row, CALENDAR_SYNC_CONFIG.SCHEDULE_COLUMNS.SYNC_STATUS).setValue(status);
    sheet.getRange(row, CALENDAR_SYNC_CONFIG.SCHEDULE_COLUMNS.EVENT_ID).setValue(eventId);
    
    console.log(`[동기화 결과 업데이트] ${row}행 - 상태: ${status}, ID: ${eventId}`);
    
  } catch (error) {
    console.error('[동기화 결과 업데이트] 업데이트 중 오류 발생:', error);
  }
}

/**
 * ==================== 학생정보DB에서 학생 조회 ====================
 * 학생 이름을 기준으로 학생정보DB 시트에서 해당 학생의 상세 정보를 조회합니다.
 * 
 * @param {string} studentName - 조회할 학생 이름
 * @returns {Object|null} 학생 정보 객체 또는 null (학생을 찾지 못한 경우)
 */
function findStudentInDatabase(studentName) {
  try {
    console.log(`[학생 조회] 학생명: "${studentName}"`);
    
    // 현재 스프레드시트에서 학생정보DB 시트 가져오기
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const studentDbSheet = spreadsheet.getSheetByName(CALENDAR_SYNC_CONFIG.SHEET_NAMES.STUDENT_DATABASE);
    
    if (!studentDbSheet) {
      console.error('[학생 조회] 학생정보DB 시트를 찾을 수 없음');
      return null;
    }
    
    // 학생정보DB의 데이터 범위 확인
    const lastRow = studentDbSheet.getLastRow();
    if (lastRow < 2) {
      console.error('[학생 조회] 학생정보DB에 데이터가 없음 (헤더만 존재)');
      return null;
    }
    
    // 2행부터 마지막 행까지 데이터 읽기 (헤더 제외)
    const dbData = studentDbSheet.getRange(2, 1, lastRow - 1, 14).getValues();
    console.log(`[학생 조회] 총 ${dbData.length}명의 학생 데이터 조회됨`);
    
    // 학생 이름으로 해당 행 찾기
    for (let i = 0; i < dbData.length; i++) {
      const rowData = dbData[i];
      const dbStudentName = rowData[CALENDAR_SYNC_CONFIG.STUDENT_DB_COLUMNS.STUDENT_NAME - 1]; // D열 (4-1=3)
      
      // 학생 이름 비교 (공백 제거 후 비교)
      if (dbStudentName && dbStudentName.toString().trim() === studentName.toString().trim()) {
        console.log(`[학생 조회] "${studentName}" 학생 정보 발견됨`);
        
        // 추가 정보 컬럼들에서 데이터 추출
        const additionalInfoF = rowData[CALENDAR_SYNC_CONFIG.STUDENT_DB_COLUMNS.ADDITIONAL_INFO_F - 1] || '';
        const additionalInfoG = rowData[CALENDAR_SYNC_CONFIG.STUDENT_DB_COLUMNS.ADDITIONAL_INFO_G - 1] || '';
        const additionalInfoM = rowData[CALENDAR_SYNC_CONFIG.STUDENT_DB_COLUMNS.ADDITIONAL_INFO_M - 1] || '';
        const additionalInfoN = rowData[CALENDAR_SYNC_CONFIG.STUDENT_DB_COLUMNS.ADDITIONAL_INFO_N - 1] || '';
        
        // 헤더 정보도 함께 가져와서 설명에 포함
        const headers = studentDbSheet.getRange(1, 1, 1, 14).getValues()[0];
        const headerF = headers[CALENDAR_SYNC_CONFIG.STUDENT_DB_COLUMNS.ADDITIONAL_INFO_F - 1] || 'F열';
        const headerG = headers[CALENDAR_SYNC_CONFIG.STUDENT_DB_COLUMNS.ADDITIONAL_INFO_G - 1] || 'G열';
        const headerM = headers[CALENDAR_SYNC_CONFIG.STUDENT_DB_COLUMNS.ADDITIONAL_INFO_M - 1] || 'M열';
        const headerN = headers[CALENDAR_SYNC_CONFIG.STUDENT_DB_COLUMNS.ADDITIONAL_INFO_N - 1] || 'N열';
        
        // 캘린더 이벤트 설명 구성
        const eventDescription = `${headerF}: ${additionalInfoF}\n${headerG}: ${additionalInfoG}\n${headerM}: ${additionalInfoM}\n${headerN}: ${additionalInfoN}`;
        
        console.log(`[학생 조회] 이벤트 설명 구성 완료`);
        
        return {
          name: dbStudentName,
          eventDescription: eventDescription,
          additionalInfoF,
          additionalInfoG,
          additionalInfoM,
          additionalInfoN
        };
      }
    }
    
    console.log(`[학생 조회] "${studentName}" 학생을 학생정보DB에서 찾을 수 없음`);
    return null;
    
  } catch (error) {
    console.error('[학생 조회] 학생 정보 조회 중 오류 발생:', error);
    return null;
  }
}

/**
 * ==================== 반복 일정 날짜 계산 ====================
 * 반복 주기 문자열을 분석하여 현재 월의 해당 요일 날짜들을 계산합니다.
 * 
 * @param {string} repeatFrequency - 반복 주기 (예: '매주 월요일')
 * @returns {Array<Date>} 해당 월의 모든 대상 요일 날짜 배열
 */
function calculateRecurringDates(repeatFrequency) {
  try {
    console.log(`[날짜 계산] 반복 주기 분석: "${repeatFrequency}"`);
    
    // 현재 날짜 기준으로 계산
    const currentDate = new Date();
    const targetYear = currentDate.getFullYear();
    const targetMonth = currentDate.getMonth(); // 0부터 시작 (1월=0)
    
    console.log(`[날짜 계산] 대상 년월: ${targetYear}년 ${targetMonth + 1}월`);
    
    // '매주 X요일' 형식에서 요일 추출
    const dayMatch = repeatFrequency.match(/매주\s*([월화수목금토일])요일/);
    if (!dayMatch) {
      console.error('[날짜 계산] 올바르지 않은 반복 주기 형식:', repeatFrequency);
      return [];
    }
    
    // 한글 요일을 숫자로 변환 (일요일=0, 월요일=1, ...)
    const targetDayKorean = dayMatch[1];
    const dayMapping = {
      '일': 0, '월': 1, '화': 2, '수': 3,
      '목': 4, '금': 5, '토': 6
    };
    const targetDayNumber = dayMapping[targetDayKorean];
    
    if (targetDayNumber === undefined) {
      console.error('[날짜 계산] 알 수 없는 요일:', targetDayKorean);
      return [];
    }
    
    console.log(`[날짜 계산] ${targetDayKorean}요일 → ${targetDayNumber}`);
    
    // 해당 월의 모든 대상 요일 날짜 찾기
    const targetDates = [];
    const firstDay = new Date(targetYear, targetMonth, 1);
    const lastDay = new Date(targetYear, targetMonth + 1, 0); // 다음 달 0일 = 이번 달 마지막 날
    
    console.log(`[날짜 계산] 검색 범위: ${firstDay.toDateString()} ~ ${lastDay.toDateString()}`);
    
    // 1일부터 마지막 날까지 반복하면서 해당 요일 찾기
    for (let day = 1; day <= lastDay.getDate(); day++) {
      const currentDate = new Date(targetYear, targetMonth, day);
      
      if (currentDate.getDay() === targetDayNumber) {
        targetDates.push(new Date(currentDate)); // 새 Date 객체로 복사
        console.log(`[날짜 계산] 대상 날짜 발견: ${currentDate.toDateString()}`);
      }
    }
    
    console.log(`[날짜 계산] 총 ${targetDates.length}개의 날짜 계산 완료`);
    return targetDates;
    
  } catch (error) {
    console.error('[날짜 계산] 날짜 계산 중 오류 발생:', error);
    return [];
  }
}

/**
 * ==================== 캘린더 이벤트 일괄 생성 ====================
 * 여러 날짜에 대해 동일한 내용의 캘린더 이벤트를 생성합니다.
 * 
 * @param {string} studentName - 학생 이름 (이벤트 제목으로 사용)
 * @param {string} eventDescription - 이벤트 설명
 * @param {Array<Date>} targetDates - 이벤트를 생성할 날짜 배열
 * @param {string} lessonTime - 수업 시간 정보
 * @returns {Object} 생성 결과 (성공 개수, 실패 개수)
 */
function createCalendarEventsForDates(studentName, eventDescription, targetDates, lessonTime) {
  try {
    console.log(`[이벤트 생성] ${studentName} 학생의 ${targetDates.length}개 이벤트 생성 시작`);
    
    let successCount = 0;
    let failureCount = 0;
    
    // 각 날짜별로 이벤트 생성
    targetDates.forEach((date, index) => {
      console.log(`[이벤트 생성] ${index + 1}/${targetDates.length} 진행 중: ${date.toDateString()}`);
      
      const eventCreated = createSingleCalendarEvent(
        studentName,        // 이벤트 제목
        eventDescription,   // 이벤트 설명
        date,              // 이벤트 날짜
        lessonTime         // 수업 시간
      );
      
      if (eventCreated) {
        successCount++;
        console.log(`✓ [이벤트 생성 성공] ${date.toDateString()}`);
      } else {
        failureCount++;
        console.error(`✗ [이벤트 생성 실패] ${date.toDateString()}`);
      }
    });
    
    console.log(`[이벤트 생성 완료] 성공: ${successCount}개, 실패: ${failureCount}개`);
    
    return {
      successCount,
      failureCount,
      totalCount: targetDates.length
    };
    
  } catch (error) {
    console.error('[이벤트 생성] 이벤트 일괄 생성 중 오류 발생:', error);
    return {
      successCount: 0,
      failureCount: targetDates.length,
      totalCount: targetDates.length
    };
  }
}

/**
 * ==================== 단일 캘린더 이벤트 생성 ====================
 * 특정 날짜와 시간에 캘린더 이벤트를 생성합니다.
 * 
 * @param {string} title - 이벤트 제목
 * @param {string} description - 이벤트 설명
 * @param {Date} date - 이벤트 날짜
 * @param {string} timeString - 수업 시간 문자열
 * @returns {boolean} 생성 성공 여부
 */
function createSingleCalendarEvent(title, description, date, timeString) {
  try {
    // 캘린더 객체 가져오기
    const calendar = CalendarApp.getCalendarById(CALENDAR_SYNC_CONFIG.TARGET_CALENDAR_ID);
    
    if (!calendar) {
      console.error('[이벤트 생성] 캘린더를 찾을 수 없음:', CALENDAR_SYNC_CONFIG.TARGET_CALENDAR_ID);
      return false;
    }
    
    // 시간 정보 파싱
    let startHour = CALENDAR_SYNC_CONFIG.DEFAULT_EVENT_TIME.START_HOUR;
    let startMinute = 0;
    
    if (timeString && timeString.toString().trim()) {
      const parsedTime = parseTimeString(timeString.toString().trim());
      if (parsedTime) {
        startHour = parsedTime.hour;
        startMinute = parsedTime.minute;
        console.log(`[시간 파싱] "${timeString}" → ${startHour}시 ${startMinute}분`);
      } else {
        console.log(`[시간 파싱] "${timeString}" 파싱 실패 - 기본 시간(${startHour}시) 사용`);
      }
    }
    
    // 시작/종료 시간 설정
    const startTime = new Date(date);
    startTime.setHours(startHour, startMinute, 0, 0);
    
    const endTime = new Date(startTime);
    endTime.setHours(startTime.getHours() + 1); // 1시간 수업
    
    // 캘린더 이벤트 생성
    const createdEvent = calendar.createEvent(
      title,
      startTime,
      endTime,
      {
        description: description
      }
    );
    
    if (createdEvent) {
      console.log(`[이벤트 생성] 성공 - ID: ${createdEvent.getId()}, 시간: ${startTime.toLocaleString()}`);
      return true;
    } else {
      console.error('[이벤트 생성] createEvent 반환값이 없음');
      return false;
    }
    
  } catch (error) {
    console.error('[이벤트 생성] 캘린더 이벤트 생성 중 오류:', error);
    return false;
  }
}

/**
 * ==================== 이벤트 고유 ID 생성 ====================
 * 캘린더 이벤트 추적을 위한 고유 식별자를 생성합니다.
 * 
 * @returns {string} 고유 식별자 문자열
 */
function generateEventId() {
  const timestamp = Date.now();
  const randomString = Math.random().toString(36).substring(2, 11); // 9자리 랜덤 문자열
  
  const eventId = `SYNC_${timestamp}_${randomString}`;
  console.log(`[ID 생성] 새 이벤트 ID: ${eventId}`);
  
  return eventId;
}

/**
 * ==================== 수업표 행 데이터를 캘린더 이벤트로 변환 ====================
 * 수업표 시트의 특정 행 데이터를 읽어서 구글 캘린더에 반복 일정으로 생성하는 핵심 함수입니다.
 * 학생 정보, 수업 시간, 반복 주기를 분석하여 한 달간의 모든 해당 요일에 일정을 생성합니다.
 * 
 * @param {Sheet} scheduleSheet - 수업표 스프레드시트 객체
 * @param {number} targetRow - 처리할 행 번호
 */
function processScheduleRowToCalendarEvent(scheduleSheet, targetRow) {
  try {
    console.log(`[일정 처리 시작] ${targetRow}행 데이터를 캘린더 이벤트로 변환을 시작합니다.`);
    
    // 수업표 행 데이터 읽기 및 파싱
    const scheduleData = extractScheduleDataFromRow(scheduleSheet, targetRow);
    if (!scheduleData.isValid) {
      updateSyncResult(scheduleSheet, targetRow, '전송실패', scheduleData.errorMessage);
      return;
    }
    
    // 학생 정보 조회 및 검증
    const studentInfo = findStudentInDatabase(scheduleData.studentName);
    if (!studentInfo) {
      updateSyncResult(scheduleSheet, targetRow, '전송실패', '학생 정보 없음');
      console.error(`[학생 조회 실패] "${scheduleData.studentName}" 학생을 학생정보DB에서 찾을 수 없습니다.`);
      return;
    }
    
    // 반복 일정 날짜 계산
    const targetDates = calculateRecurringDates(scheduleData.repeatFrequency);
    if (targetDates.length === 0) {
      updateSyncResult(scheduleSheet, targetRow, '전송실패', '날짜 계산 실패');
      return;
    }
    
    // 각 날짜별 캘린더 이벤트 생성
    const eventCreationResults = createCalendarEventsForDates(
      scheduleData.studentName,
      studentInfo.eventDescription,
      targetDates,
      scheduleData.lessonTime
    );
    
    // 결과 처리 및 상태 업데이트
    const uniqueEventId = generateEventId();
    if (eventCreationResults.successCount === targetDates.length) {
      updateSyncResult(scheduleSheet, targetRow, '전송완료', uniqueEventId);
      console.log(`[일정 생성 완료] 총 ${eventCreationResults.successCount}개 이벤트가 성공적으로 생성되었습니다.`);
    } else {
      updateSyncResult(scheduleSheet, targetRow, '전송실패', `${eventCreationResults.successCount}/${targetDates.length} 성공`);
      console.error(`[일정 생성 부분실패] ${eventCreationResults.successCount}/${targetDates.length}개 이벤트만 생성되었습니다.`);
    }
    
  } catch (error) {
    console.error('[일정 처리 오류] 캘린더 이벤트 처리 중 예상치 못한 오류 발생:', error);
    console.error('[일정 처리 오류] 오류 상세:', error.stack);
    updateSyncResult(scheduleSheet, targetRow, '전송실패', '시스템 오류');
  }
}

// 기존 함수명과의 호환성을 위한 별칭
const handleCalendarEvent = processScheduleRowToCalendarEvent;

/**
 * ==================== 수업표 행 데이터 추출 및 검증 ====================
 * 수업표의 특정 행에서 필요한 모든 데이터를 추출하고 유효성을 검증합니다.
 * 
 * @param {Sheet} sheet - 수업표 시트 객체
 * @param {number} row - 데이터를 읽을 행 번호
 * @returns {Object} 추출된 데이터와 유효성 검증 결과
 */
function extractScheduleDataFromRow(sheet, row) {
  try {
    // 해당 행의 모든 데이터 읽기 (A열부터 O열까지)
    const rowData = sheet.getRange(row, 1, 1, 15).getValues()[0];
    
    // 각 컬럼별 데이터 추출
    const extractedData = {
      triggerCommand: rowData[CALENDAR_SYNC_CONFIG.SCHEDULE_COLUMNS.TRIGGER_COMMAND - 1],
      studentName: rowData[CALENDAR_SYNC_CONFIG.SCHEDULE_COLUMNS.STUDENT_NAME - 1],
      lessonContent: rowData[CALENDAR_SYNC_CONFIG.SCHEDULE_COLUMNS.LESSON_CONTENT - 1],
      repeatFrequency: rowData[CALENDAR_SYNC_CONFIG.SCHEDULE_COLUMNS.REPEAT_FREQUENCY - 1],
      lessonTime: rowData[CALENDAR_SYNC_CONFIG.SCHEDULE_COLUMNS.LESSON_TIME - 1]
    };
    
    console.log(`[데이터 추출] 학생명: "${extractedData.studentName}", 주기: "${extractedData.repeatFrequency}", 시간: "${extractedData.lessonTime}"`);
    
    console.log(`[수업표 데이터] 트리거:"${extractedData.triggerCommand}", 학생명:"${extractedData.studentName}", 주기:"${extractedData.repeatFrequency}", 시간:"${extractedData.lessonTime}"`); // << 데이터 로그
    
    // 필수 데이터 유효성 검사
    const validationResult = {
      isValid: true,
      errorMessage: ''
    };
    
    if (!extractedData.studentName || !extractedData.repeatFrequency || !extractedData.lessonTime) {
      validationResult.isValid = false;
      validationResult.errorMessage = '필수 데이터 누락 (학생명, 반복주기, 수업시간 중 일부 누락)';
      console.error('[검증 실패] 필수 데이터 누락:', { 
        studentName: extractedData.studentName, 
        repeatFrequency: extractedData.repeatFrequency, 
        lessonTime: extractedData.lessonTime 
      });
      return validationResult;
    }
    
    // 트리거 값 재확인 (안전장치)
    if (extractedData.triggerCommand !== CALENDAR_SYNC_CONFIG.SYNC_TRIGGER_WORD) {
      validationResult.isValid = false;
      validationResult.errorMessage = '트리거 값이 일치하지 않음';
      console.log('[검증 실패] 트리거 값이 일치하지 않음 - 처리 중단');
      return validationResult;
    }
    
    // 반환할 데이터 구성
    validationResult.studentName = extractedData.studentName;
    validationResult.repeatFrequency = extractedData.repeatFrequency;
    validationResult.lessonTime = extractedData.lessonTime;
    validationResult.lessonContent = extractedData.lessonContent;
    
    console.log('[데이터 추출 완료] 유효성 검사 통과');
    return validationResult;
    
  } catch (error) {
    console.error('[데이터 추출 오류] 수업표 행 데이터 추출 중 오류:', error);
    return {
      isValid: false,
      errorMessage: '데이터 추출 중 시스템 오류: ' + error.message
    };
  }
}

// 이 함수는 상단에 calculateRecurringDates 함수로 대체되었습니다.

/******** 5. 구글 캘린더 이벤트 생성 ********/ // << 실제 캘린더 이벤트 생성 함수
function createCalendarEvent(title, description, date, timeString) {
  try {
    // 캘린더 객체 가져오기 (권한 필요)
    const calendar = CalendarApp.getCalendarById(CALENDAR_SYNC_CONFIG.TARGET_CALENDAR_ID); // << 캘린더 객체
    
    if (!calendar) {
      console.error('[캘린더 오류] 캘린더를 찾을 수 없음:', CALENDAR_SYNC_CONFIG.TARGET_CALENDAR_ID); // << 캘린더 없음 로그
      return false;
    }
    
    // E열 시간 데이터 파싱 (예: "오후 1:30:00" 또는 "13:30:00")
    let startHour = CALENDAR_SYNC_CONFIG.DEFAULT_EVENT_TIME.START_HOUR; // << 기본 시작 시간
    let startMinute = 0; // << 기본 시작 분
    
    if (timeString && timeString.toString().trim()) {
      const parsedTime = parseTimeString(timeString.toString().trim()); // << 시간 문자열 파싱
      if (parsedTime) {
        startHour = parsedTime.hour; // << 파싱된 시간
        startMinute = parsedTime.minute; // << 파싱된 분
        console.log(`[시간 파싱] "${timeString}" → ${startHour}시 ${startMinute}분`); // << 파싱 로그
      } else {
        console.log(`[시간 파싱] "${timeString}" 파싱 실패 - 기본 시간 사용`); // << 파싱 실패 로그
      }
    }
    
    // 시작 시간 설정
    const startTime = new Date(date); // << 시작 시간 복사
    startTime.setHours(startHour, startMinute, 0, 0); // << 파싱된 시간으로 설정
    
    // 종료 시간 설정 (1시간 후)
    const endTime = new Date(startTime); // << 종료 시간 복사
    endTime.setHours(startTime.getHours() + 1); // << 1시간 후로 설정
    
    // 캘린더 이벤트 생성
    const createdEvent = calendar.createEvent(
      title,           // << 이벤트 제목
      startTime,       // << 시작 시간
      endTime,         // << 종료 시간
      {
        description: description    // << 이벤트 설명
      }
    );
    
    if (createdEvent) {
      console.log(`[이벤트 생성] 성공 - ID: ${createdEvent.getId()}, 시간: ${startTime.toLocaleString()}`);
      return true;
    } else {
      console.error('[이벤트 생성] createEvent 반환값이 없음');
      return false;
    }
    
  } catch (error) {
    console.error('[이벤트 생성] 캘린더 이벤트 생성 중 오류:', error); // << 생성 오류 로그
    return false;
  }
}

/******** 5-1. 시간 문자열 파싱 함수 ********/ // << 시간 문자열을 시간/분으로 변환
function parseTimeString(timeData) {
  try {
    console.log(`[시간 파싱] 입력 데이터:`, timeData, `타입: ${typeof timeData}`); // << 입력 데이터 로그
    
    // Date 객체인 경우 직접 시간 추출
    if (timeData instanceof Date) {
      const hour = timeData.getHours(); // << Date 객체에서 시간 추출
      const minute = timeData.getMinutes(); // << Date 객체에서 분 추출
      console.log(`[시간 파싱] Date 객체에서 추출: ${hour}시 ${minute}분`); // << 추출 로그
      return { hour, minute }; // << 추출된 시간 반환
    }
    
    // 문자열로 변환하여 처리
    const timeStr = timeData.toString().trim(); // << 문자열 변환
    console.log(`[시간 파싱] 문자열 변환: "${timeStr}"`); // << 변환 로그
    
    // "오후 1:30:00" 형식 파싱
    const ampmMatch = timeStr.match(/^(오전|오후)\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/); // << 오전/오후 형식
    if (ampmMatch) {
      const isAM = ampmMatch[1] === '오전'; // << 오전 여부
      let hour = parseInt(ampmMatch[2], 10); // << 시간
      const minute = parseInt(ampmMatch[3], 10); // << 분
      
      // 12시간 → 24시간 변환
      if (!isAM && hour !== 12) {
        hour += 12; // << 오후 시간 변환 (12시 제외)
      } else if (isAM && hour === 12) {
        hour = 0; // << 오전 12시 → 0시
      }
      
      console.log(`[시간 파싱] 오전/오후 파싱: ${hour}시 ${minute}분`); // << 파싱 로그
      return { hour, minute }; // << 변환된 시간 반환
    }
    
    // "13:30:00" 또는 "13:30" 형식 파싱
    const timeMatch = timeStr.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/); // << 24시간 형식
    if (timeMatch) {
      const hour = parseInt(timeMatch[1], 10); // << 시간
      const minute = parseInt(timeMatch[2], 10); // << 분
      
      if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
        console.log(`[시간 파싱] 24시간 파싱: ${hour}시 ${minute}분`); // << 파싱 로그
        return { hour, minute }; // << 유효한 시간 반환
      }
    }
    
    // GMT 시간 문자열에서 시간 추출 시도
    const gmtMatch = timeStr.match(/(\d{2}):(\d{2}):(\d{2})\s*GMT/); // << GMT 형식
    if (gmtMatch) {
      const hour = parseInt(gmtMatch[1], 10); // << 시간
      const minute = parseInt(gmtMatch[2], 10); // << 분
      console.log(`[시간 파싱] GMT 파싱: ${hour}시 ${minute}분`); // << GMT 파싱 로그
      return { hour, minute }; // << GMT 시간 반환
    }
    
    // 숫자만 있는 경우 (시리얼 넘버로 들어온 시간)
    const numericValue = parseFloat(timeData); // << 숫자 변환
    if (!isNaN(numericValue) && numericValue > 0 && numericValue < 1) {
      // 시리얼 시간 (0.5 = 12:00:00)
      const totalMinutes = Math.round(numericValue * 24 * 60); // << 총 분 계산
      const hour = Math.floor(totalMinutes / 60); // << 시간 계산
      const minute = totalMinutes % 60; // << 분 계산
      console.log(`[시간 파싱] 시리얼 넘버 파싱: ${hour}시 ${minute}분`); // << 시리얼 파싱 로그
      return { hour, minute }; // << 계산된 시간 반환
    }
    
    console.log(`[시간 파싱] 모든 형식 파싱 실패`); // << 파싱 실패 로그
    return null; // << 파싱 실패
    
  } catch (error) {
    console.error('[시간 파싱] 파싱 중 오류:', error); // << 파싱 오류 로그
    return null; // << 오류 시 null 반환
  }
}

// 이 함수는 상단에 findStudentInDatabase 함수로 대체되었습니다.

/******** 5-2. 제목용 시간 형식 변환 함수 ********/ // << 제목에 표시할 깔끔한 시간 형식 생성
function formatTimeForTitle(timeData) {
  try {
    console.log(`[제목 시간 변환] 입력:`, timeData); // << 입력 로그
    
    // 시간 파싱
    const parsedTime = parseTimeString(timeData); // << 기존 파싱 함수 사용
    if (!parsedTime) {
      return '시간미정'; // << 파싱 실패 시 기본값
    }
    
    const { hour, minute } = parsedTime; // << 시간, 분 추출
    
    // 오전/오후 형식으로 변환
    let displayHour = hour; // << 표시용 시간
    let ampm = '오전'; // << 오전/오후
    
    if (hour === 0) {
      displayHour = 12; // << 0시 → 오전 12시
      ampm = '오전';
    } else if (hour < 12) {
      displayHour = hour; // << 1~11시 → 오전
      ampm = '오전';
    } else if (hour === 12) {
      displayHour = 12; // << 12시 → 오후 12시
      ampm = '오후';
    } else {
      displayHour = hour - 12; // << 13~23시 → 오후
      ampm = '오후';
    }
    
    // 분이 0이면 생략, 아니면 포함
    const formattedTime = minute === 0 
      ? `${ampm} ${displayHour}시` // << 정각인 경우
      : `${ampm} ${displayHour}:${minute.toString().padStart(2, '0')}`; // << 분 포함
    
    console.log(`[제목 시간 변환] 결과: "${formattedTime}"`); // << 변환 결과 로그
    return formattedTime; // << 변환된 시간 반환
    
  } catch (error) {
    console.error('[제목 시간 변환] 변환 중 오류:', error); // << 변환 오류 로그
    return '시간미정'; // << 오류 시 기본값
  }
}

// 이 함수는 상단에 updateSyncResult 함수로 대체되었습니다.

// 이 함수는 상단에 generateEventId 함수로 대체되었습니다.

/**
 * ==================== 트리거 설정 함수 ====================
 * 스프레드시트 편집 이벤트에 대한 Installable Trigger를 설정합니다.
 * 
 * @returns {string} 설정 결과 메시지
 */
function setupCalendarTrigger() {
  try {
    // 기존 트리거 중복 방지를 위해 삭제
    const existingTriggers = ScriptApp.getProjectTriggers();
    existingTriggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'executeCalendarSynchronization') {
        ScriptApp.deleteTrigger(trigger);
        console.log('[트리거 설정] 기존 트리거 삭제됨');
      }
    });
    
    // 새 편집 트리거 등록
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const newTrigger = ScriptApp.newTrigger('executeCalendarSynchronization')
      .forSpreadsheet(currentSpreadsheet)
      .onEdit()
      .create();
    
    console.log('✅ [트리거 설정] 편집 트리거가 성공적으로 등록되었습니다!');
    console.log('💡 [사용법] 이제 A열에 "캘린더이동"을 입력하면 자동으로 캘린더 동기화가 실행됩니다.');
    
    return '트리거 등록 완료';
    
  } catch (error) {
    console.error('[트리거 설정] 트리거 등록 중 오류:', error);
    return '트리거 등록 실패: ' + error.toString();
  }
}

/**
 * ==================== 권한 테스트 함수 ====================
 * 캘린더 API 접근 권한이 올바르게 설정되어 있는지 확인합니다.
 * 
 * @returns {string} 테스트 결과 메시지
 */
function testCalendarPermission() {
  try {
    const calendar = CalendarApp.getCalendarById(CALENDAR_SYNC_CONFIG.TARGET_CALENDAR_ID);
    
    if (calendar) {
      const calendarName = calendar.getName();
      console.log('✅ [권한 테스트] 캘린더 액세스 성공:', calendarName);
      return `권한 테스트 성공 - 캘린더: ${calendarName}`;
    } else {
      console.error('❌ [권한 테스트] 캘린더를 찾을 수 없음');
      return '권한 테스트 실패 - 캘린더 없음';
    }
    
  } catch (error) {
    console.error('❌ [권한 테스트] 권한 테스트 실패:', error);
    return '권한 테스트 실패: ' + error.toString();
  }
}

// ============================================================================
// << 사용 방법 및 예제
// ============================================================================
/*
🚀 설정 및 사용 가이드:

1️⃣ 권한 설정:
   - 먼저 'testCalendarPermission()' 함수를 실행하여 캘린더 접근 권한 승인

2️⃣ 트리거 등록:
   - 'setupCalendarTrigger()' 함수를 실행하여 자동 트리거 설정

3️⃣ 사용법:
   - 구글 시트 A열에 "캘린더이동" 입력
   - B열: 제목 앞부분 (예: "회의")
   - C열: 일정 내용 (예: "주간 팀 미팅")
   - D열: 반복 주기 (예: "매주 월요일")
   - E열: 제목 뒷부분 (예: "팀회의")
   
   📋 결과: "회의_팀회의" 제목으로 해당 월의 모든 월요일에 일정 생성

4️⃣ 결과 확인:
   - N열: 전송완료/전송실패 상태
   - O열: 고유 이벤트 ID

5️⃣ 트러블슈팅:
   - 권한 오류 시: appsscript.json에 캘린더 권한 확인
   - 트리거 오류 시: setupCalendarTrigger() 재실행
   - 날짜 오류 시: D열 형식 확인 ("매주 X요일")

📝 주의사항:
   - Simple Trigger(onEdit)는 권한 제한으로 캘린더 API 사용 불가
   - Installable Trigger(processCalendarSync) 등록 필수
   - 이벤트 시간은 오전 9시~10시 고정 (CFG.TIME에서 수정 가능)
*/