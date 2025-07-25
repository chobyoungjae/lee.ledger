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
    
    console.log(`[수업표 데이터] 트리거:"${trigger}", 학생명:"${studentName}", 주기:"${frequency}", 시간:"${timeData}"`); // << 데이터 로그
    
    // 트리거 값 재확인 (안전장치)
    if (trigger !== CFG.TRIGGER_WORD) {
      console.log('[검증 실패] 트리거 값이 일치하지 않음 - 처리 중단'); // << 검증 실패 로그
      return;
    }
    
    // 필수 데이터 유효성 검사
    if (!studentName || !frequency || !timeData) {
      updateResult(sheet, row, '전송실패', '필수 데이터 누락'); // << 결과 업데이트
      console.error('[검증 실패] 필수 데이터 누락:', { studentName, frequency, timeData }); // << 에러 로그
      return;
    }
    
    // 학생정보DB에서 해당 학생 정보 찾기
    const studentInfo = getStudentInfo(studentName); // << 학생 정보 조회
    if (!studentInfo) {
      updateResult(sheet, row, '전송실패', '학생 정보 없음'); // << 결과 업데이트
      console.error('[검증 실패] 학생 정보를 찾을 수 없음:', studentName); // << 에러 로그
      return;
    }
    
    // 이벤트 제목과 내용 생성
    const eventTitle = studentName; // << 학생 이름만 제목으로 사용
    const eventContent = studentInfo.content; // << 학생정보DB에서 조회한 내용
    
    // 현재 날짜 기준 년월 계산
    const currentDate = new Date(); // << 현재 날짜
    const targetYear = currentDate.getFullYear(); // << 대상 년도
    const targetMonth = currentDate.getMonth(); // << 대상 월 (0부터 시작)
    
    console.log(`[처리 대상] ${targetYear}년 ${targetMonth + 1}월, 제목: "${eventTitle}"`); // << 처리 대상 로그
    
    // 요일 정보 추출 (정규표현식)
    const dayMatch = frequency.match(/매주\s*([월화수목금토일])요일/); // << 요일 매칭
    if (!dayMatch) {
      updateResult(sheet, row, '전송실패', '요일 형식 오류'); // << 결과 업데이트
      console.error('[형식 오류] 요일 형식이 올바르지 않음:', frequency); // << 형식 오류 로그
      return;
    }
    
    // 한글 요일을 숫자로 변환
    const targetDayKorean = dayMatch[1]; // << 한글 요일
    const dayMapping = { // << 요일 변환 맵
      '일': 0, '월': 1, '화': 2, '수': 3, 
      '목': 4, '금': 5, '토': 6
    };
    const targetDayNumber = dayMapping[targetDayKorean]; // << 요일 숫자
    
    if (targetDayNumber === undefined) {
      updateResult(sheet, row, '전송실패', '요일 변환 실패'); // << 결과 업데이트
      console.error('[변환 실패] 요일 변환 실패:', targetDayKorean); // << 변환 실패 로그
      return;
    }
    
    console.log(`[요일 변환] ${targetDayKorean}요일 → ${targetDayNumber}`); // << 변환 로그
    
    // 고유 이벤트 ID 생성 (같은 반복 일정은 동일 ID)
    const uniqueEventId = generateUniqueId(); // << 고유 ID 생성
    
    // 해당 월의 모든 대상 요일 날짜 계산
    const targetDates = getTargetDatesInMonth(targetYear, targetMonth, targetDayNumber); // << 대상 날짜 목록
    
    console.log(`[일정 생성] 생성할 일정 수: ${targetDates.length}개`); // << 생성 수 로그
    
    // 각 날짜에 캘린더 이벤트 생성
    let successCount = 0; // << 성공 카운터
    targetDates.forEach((date, index) => {
      console.log(`[진행률] ${index + 1}/${targetDates.length} 일정 생성 중...`); // << 진행률 로그
      
      if (createCalendarEvent(eventTitle, eventContent, date, timeData)) {
        successCount++; // << 성공 카운트 증가
        console.log(`✓ [성공] ${date.toDateString()} 일정 생성 완료`); // << 성공 로그
      } else {
        console.error(`✗ [실패] ${date.toDateString()} 일정 생성 실패`); // << 실패 로그
      }
    });
    
    // 최종 결과에 따라 상태 업데이트
    if (successCount === targetDates.length) {
      updateResult(sheet, row, '전송완료', uniqueEventId); // << 성공 결과 업데이트
      console.log(`🎉 [완료] 모든 일정 생성 성공: ${successCount}/${targetDates.length}`); // << 완료 로그
    } else {
      updateResult(sheet, row, '전송실패', uniqueEventId); // << 실패 결과 업데이트
      console.error(`⚠️ [부분실패] 일부 일정 생성 실패: ${successCount}/${targetDates.length}`); // << 부분실패 로그
    }
    
  } catch (error) {
    console.error('[치명적 오류] 캘린더 이벤트 처리 중 오류:', error); // << 치명적 오류 로그
    updateResult(sheet, row, '전송실패', '시스템 오류'); // << 시스템 오류 결과
  }
}

/******** 4. 특정 월의 요일 날짜 계산 ********/ // << 월별 특정 요일 날짜 계산 함수
function getTargetDatesInMonth(year, month, dayOfWeek) {
  const dates = []; // << 날짜 목록 배열
  
  // 해당 월의 첫날과 마지막날 계산
  const firstDay = new Date(year, month, 1); // << 월 첫날
  const lastDay = new Date(year, month + 1, 0); // << 월 마지막날 (다음달 0일)
  
  console.log(`[날짜 계산] 범위: ${firstDay.toDateString()} ~ ${lastDay.toDateString()}`); // << 계산 범위 로그
  
  // 1일부터 마지막날까지 반복하면서 해당 요일 찾기
  for (let day = 1; day <= lastDay.getDate(); day++) {
    const currentDate = new Date(year, month, day); // << 현재 확인 날짜
    
    // 현재 날짜의 요일이 대상 요일과 일치하면 목록에 추가
    if (currentDate.getDay() === dayOfWeek) {
      dates.push(new Date(currentDate)); // << 새 Date 객체로 추가 (참조 문제 방지)
    }
  }
  
  console.log(`[찾은 날짜] ${dates.map(d => d.getDate()).join(', ')}일`); // << 찾은 날짜 로그
  return dates; // << 날짜 배열 반환
}

/******** 5. 구글 캘린더 이벤트 생성 ********/ // << 실제 캘린더 이벤트 생성 함수
function createCalendarEvent(title, description, date, timeString) {
  try {
    // 캘린더 객체 가져오기 (권한 필요)
    const calendar = CalendarApp.getCalendarById(CFG.CALENDAR_ID); // << 캘린더 객체
    
    if (!calendar) {
      console.error('[캘린더 오류] 캘린더를 찾을 수 없음:', CFG.CALENDAR_ID); // << 캘린더 없음 로그
      return false;
    }
    
    // E열 시간 데이터 파싱 (예: "오후 1:30:00" 또는 "13:30:00")
    let startHour = CFG.TIME.START_HOUR; // << 기본 시작 시간
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
    
    // 참조 코드와 동일한 방식으로 이벤트 생성 (createEvent 사용)
    const createdEvent = calendar.createEvent(
      title,           // << 이벤트 제목
      startTime,       // << 시작 시간
      endTime,         // << 종료 시간
      {
        description: description    // << 이벤트 설명
      }
    );
    
    console.log(`[이벤트 생성] 성공, 이벤트 ID: ${createdEvent.getId()}`); // << 성공 로그
    
    if (createdEvent) {
      console.log(`[이벤트 생성] 성공: "${title}" (${startTime.toLocaleString()})`); // << 생성 성공 로그
      return true;
    } else {
      console.error('[이벤트 생성] 실패: createEvent 반환값 없음'); // << 생성 실패 로그
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

/******** 4-1. 학생정보 조회 함수 ********/ // << 학생정보DB에서 학생 정보 조회
function getStudentInfo(studentName) {
  try {
    console.log(`[학생정보 조회] 학생명: "${studentName}"`); // << 조회 시작 로그
    
    // 학생정보DB 시트 가져오기
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // << 현재 스프레드시트
    const studentDbSheet = ss.getSheetByName(CFG.SHEETS.STUDENT_DB); // << 학생정보DB 시트
    
    if (!studentDbSheet) {
      console.error('[학생정보 조회] 학생정보DB 시트를 찾을 수 없음'); // << 시트 없음 로그
      return null;
    }
    
    // 학생정보DB의 모든 데이터 가져오기
    const lastRow = studentDbSheet.getLastRow(); // << 마지막 행
    if (lastRow < 2) {
      console.error('[학생정보 조회] 학생정보DB에 데이터가 없음'); // << 데이터 없음 로그
      return null;
    }
    
    const dbData = studentDbSheet.getRange(2, 1, lastRow - 1, 14).getValues(); // << 2행부터 N열까지 데이터
    console.log(`[학생정보 조회] 총 ${dbData.length}명의 학생 데이터 로드`); // << 데이터 로드 로그
    
    // 학생 이름으로 해당 행 찾기
    for (let i = 0; i < dbData.length; i++) {
      const rowData = dbData[i]; // << 현재 행 데이터
      const dbStudentName = rowData[CFG.STUDENT_DB_COL.NAME - 1]; // << D열: 학생 이름
      
      if (dbStudentName && dbStudentName.toString().trim() === studentName.toString().trim()) {
        // 일치하는 학생 찾음, 필요한 데이터 추출
        const fData = rowData[CFG.STUDENT_DB_COL.F_DATA - 1] || ''; // << F열 데이터
        const gData = rowData[CFG.STUDENT_DB_COL.G_DATA - 1] || ''; // << G열 데이터
        const mData = rowData[CFG.STUDENT_DB_COL.M_DATA - 1] || ''; // << M열 데이터
        const nData = rowData[CFG.STUDENT_DB_COL.N_DATA - 1] || ''; // << N열 데이터
        
        // F1, G1, M1, N1은 헤더를 의미하는 것 같으니 실제 헤더 값 가져오기
        const headers = studentDbSheet.getRange(1, 1, 1, 14).getValues()[0]; // << 헤더 행
        const fHeader = headers[CFG.STUDENT_DB_COL.F_DATA - 1] || 'F1'; // << F열 헤더
        const gHeader = headers[CFG.STUDENT_DB_COL.G_DATA - 1] || 'G1'; // << G열 헤더
        const mHeader = headers[CFG.STUDENT_DB_COL.M_DATA - 1] || 'M1'; // << M열 헤더
        const nHeader = headers[CFG.STUDENT_DB_COL.N_DATA - 1] || 'N1'; // << N열 헤더
        
        // 캘린더 이벤트 내용 구성
        const content = `${fHeader} : ${fData}\n${gHeader} : ${gData}\n${mHeader} : ${mData}\n${nHeader} : ${nData}`; // << 내용 조합
        
        console.log(`[학생정보 조회] "${studentName}" 학생 정보 찾음`); // << 찾음 로그
        console.log(`[학생정보 내용] ${content}`); // << 내용 로그
        
        return {
          name: dbStudentName, // << 학생 이름
          content: content,    // << 캘린더 이벤트 내용
          fData, gData, mData, nData // << 개별 데이터 (필요시 사용)
        };
      }
    }
    
    console.log(`[학생정보 조회] "${studentName}" 학생을 찾을 수 없음`); // << 찾지 못함 로그
    return null; // << 학생 정보 없음
    
  } catch (error) {
    console.error('[학생정보 조회] 조회 중 오류:', error); // << 조회 오류 로그
    return null; // << 오류 시 null 반환
  }
}

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

/******** 6. 처리 결과 업데이트 ********/ // << 시트에 처리 결과 기록 함수
function updateResult(sheet, row, status, uniqueId) {
  try {
    // N열에 상태, O열에 고유 ID 기록
    sheet.getRange(row, CFG.COL.STATUS).setValue(status); // << N열: 상태 기록
    sheet.getRange(row, CFG.COL.UNIQUE_ID).setValue(uniqueId); // << O열: 고유 ID 기록
    
    console.log(`[결과 업데이트] ${row}행 - ${status} (${uniqueId})`); // << 업데이트 로그
    
  } catch (error) {
    console.error('[결과 업데이트] 업데이트 중 오류:', error); // << 업데이트 오류 로그
  }
}

/******** 7. 고유 ID 생성 ********/ // << 고유 ID 생성 함수
function generateUniqueId() {
  const timestamp = Date.now(); // << 현재 시간 (밀리초)
  const randomString = Math.random().toString(36).substr(2, 9); // << 랜덤 문자열 (9자리)
  
  const uniqueId = `EVENT_${timestamp}_${randomString}`; // << 고유 ID 조합
  
  console.log('[ID 생성] 새 고유 ID:', uniqueId); // << ID 생성 로그
  return uniqueId; // << 고유 ID 반환
}

/******** 8. 트리거 설정 함수 ********/ // << 수동 트리거 등록 함수
function setupCalendarTrigger() {
  try {
    // 기존 트리거 중복 방지를 위해 삭제
    const existingTriggers = ScriptApp.getProjectTriggers(); // << 기존 트리거 목록
    existingTriggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'processCalendarSync') {
        ScriptApp.deleteTrigger(trigger); // << 기존 트리거 삭제
        console.log('[트리거 설정] 기존 트리거 삭제됨'); // << 삭제 로그
      }
    });
    
    // 새 편집 트리거 등록
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // << 현재 스프레드시트
    const newTrigger = ScriptApp.newTrigger('processCalendarSync') // << 트리거 함수명 지정
      .forSpreadsheet(currentSpreadsheet) // << 스프레드시트 지정
      .onEdit() // << 편집 이벤트
      .create(); // << 트리거 생성
    
    console.log('✅ [트리거 설정] 편집 트리거가 성공적으로 등록되었습니다!'); // << 성공 로그
    console.log('💡 [사용법] 이제 A열에 "캘린더이동"을 입력하면 자동으로 캘린더 동기화가 실행됩니다.'); // << 사용법 안내
    
    return '트리거 등록 완료'; // << 성공 반환
    
  } catch (error) {
    console.error('[트리거 설정] 트리거 등록 중 오류:', error); // << 등록 오류 로그
    return '트리거 등록 실패: ' + error.toString(); // << 실패 반환
  }
}

/******** 9. 권한 테스트 함수 ********/ // << 캘린더 API 권한 확인 함수
function testCalendarPermission() {
  try {
    const calendar = CalendarApp.getCalendarById(CFG.CALENDAR_ID); // << 캘린더 객체 조회
    
    if (calendar) {
      const calendarName = calendar.getName(); // << 캘린더 이름 조회
      console.log('✅ [권한 테스트] 캘린더 액세스 성공:', calendarName); // << 성공 로그
      return `권한 테스트 성공 - 캘린더: ${calendarName}`; // << 성공 반환
    } else {
      console.error('❌ [권한 테스트] 캘린더를 찾을 수 없음'); // << 캘린더 없음 로그
      return '권한 테스트 실패 - 캘린더 없음'; // << 실패 반환
    }
    
  } catch (error) {
    console.error('❌ [권한 테스트] 권한 테스트 실패:', error); // << 권한 실패 로그
    return '권한 테스트 실패: ' + error.toString(); // << 실패 반환
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