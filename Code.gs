/**
 * 4학년 국어 온라인 대화 수업용 채팅 프로그램
 * 스프레드시트 ID: 1BUwIIddGcqYENhht3EMDOynm2Zbe5idXpgyJMQ0E59U
 */

const SPREADSHEET_ID = '1BUwIIddGcqYENhht3EMDOynm2Zbe5idXpgyJMQ0E59U';

// 웹앱 진입점
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('우리반 대화방')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// 스프레드시트 가져오기
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// 시트 초기화 (처음 실행 시)
function initializeSheets() {
  const ss = getSpreadsheet();
  
  // 회원정보 시트
  let memberSheet = ss.getSheetByName('회원정보');
  if (!memberSheet) {
    memberSheet = ss.insertSheet('회원정보');
    memberSheet.getRange(1, 1, 1, 4).setValues([['학번', '이름', '비밀번호', '가입일시']]);
    memberSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#FEE500');
  }
  
  // 대화내용 시트
  let chatSheet = ss.getSheetByName('대화내용');
  if (!chatSheet) {
    chatSheet = ss.insertSheet('대화내용');
    chatSheet.getRange(1, 1, 1, 5).setValues([['대화방번호', '작성자', '내용', '작성시간', '타임스탬프']]);
    chatSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#FEE500');
  }
  
  return { success: true, message: '시트가 초기화되었습니다.' };
}

// 회원가입
function registerUser(name, password) {
  try {
    if (!name || !password) {
      return { success: false, message: '이름과 비밀번호를 모두 입력해주세요.' };
    }
    
    name = String(name).trim();
    password = String(password).trim();
    
    if (name.length > 10) {
      return { success: false, message: '이름은 10자 이내로 입력해주세요.' };
    }
    
    if (password.length < 2 || password.length > 10) {
      return { success: false, message: '비밀번호는 2~10자로 입력해주세요.' };
    }
    
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('회원정보');
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName('회원정보');
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 중복 이름 체크
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === name) {
        return { success: false, message: '이미 사용 중인 이름입니다.' };
      }
    }
    
    // 새 회원 추가
    const newRow = data.length;
    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    
    sheet.getRange(newRow + 1, 1, 1, 4).setValues([[newRow, name, password, dateStr]]);
    sheet.getRange(newRow + 1, 3).setNumberFormat('@');
    
    return { success: true, message: '회원가입이 완료되었습니다!', userName: name };
  } catch (error) {
    return { success: false, message: '오류가 발생했습니다: ' + error.toString() };
  }
}

// 로그인
function loginUser(name, password) {
  try {
    if (!name || !password) {
      return { success: false, message: '이름과 비밀번호를 모두 입력해주세요.' };
    }
    
    name = String(name).trim();
    password = String(password).trim();
    
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('회원정보');
    
    if (!sheet) {
      return { success: false, message: '회원 정보가 없습니다. 먼저 회원가입을 해주세요.' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const storedName = String(data[i][1]).trim();
      const storedPw = String(data[i][2]).trim();
      
      if (storedName === name && storedPw === password) {
        return { success: true, message: '로그인 성공!', userName: name };
      }
    }
    
    return { success: false, message: '이름 또는 비밀번호가 일치하지 않습니다.' };
  } catch (error) {
    return { success: false, message: '오류가 발생했습니다: ' + error.toString() };
  }
}

// 메시지 전송
function sendMessage(roomNumber, userName, content) {
  try {
    if (roomNumber === undefined || roomNumber === null || !userName || !content) {
      return { success: false, message: '필수 정보가 누락되었습니다.' };
    }
    
    // 타입 변환
    roomNumber = Number(roomNumber);
    userName = String(userName).trim();
    content = String(content).trim();
    
    if (content.length === 0) {
      return { success: false, message: '메시지를 입력해주세요.' };
    }
    
    if (content.length > 500) {
      return { success: false, message: '메시지는 500자 이내로 입력해주세요.' };
    }
    
    // 특수문자 필터링 (XSS 방지)
    content = content.replace(/</g, '&lt;').replace(/>/g, '&gt;');
    
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('대화내용');
    
    if (!sheet) {
      initializeSheets();
      sheet = ss.getSheetByName('대화내용');
    }
    
    const now = new Date();
    const timeStr = Utilities.formatDate(now, 'Asia/Seoul', 'HH:mm');
    const timestamp = now.getTime();
    
    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, 5).setValues([[roomNumber, userName, content, timeStr, timestamp]]);
    
    return { success: true, message: '전송 완료', timestamp: timestamp };
  } catch (error) {
    return { success: false, message: '오류가 발생했습니다: ' + error.toString() };
  }
}

// 메시지 불러오기 (새 메시지만)
function getMessages(roomNumber, lastTimestamp) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('대화내용');
    
    if (!sheet) {
      return { success: true, messages: [], lastTimestamp: 0 };
    }
    
    // 타입 변환
    roomNumber = Number(roomNumber);
    lastTimestamp = Number(lastTimestamp) || 0;
    
    const data = sheet.getDataRange().getValues();
    const messages = [];
    let newLastTimestamp = lastTimestamp;
    
    for (let i = 1; i < data.length; i++) {
      const rowRoomNum = Number(data[i][0]);
      const msgTimestamp = Number(data[i][4]) || 0;
      
      if (rowRoomNum === roomNumber && msgTimestamp > lastTimestamp) {
        messages.push({
          userName: String(data[i][1]),
          content: String(data[i][2]),
          time: String(data[i][3]),
          timestamp: msgTimestamp
        });
        
        if (msgTimestamp > newLastTimestamp) {
          newLastTimestamp = msgTimestamp;
        }
      }
    }
    
    return { success: true, messages: messages, lastTimestamp: newLastTimestamp };
  } catch (error) {
    return { success: false, messages: [], lastTimestamp: lastTimestamp || 0, error: error.toString() };
  }
}

// 모든 메시지 불러오기 (처음 입장 시)
function getAllMessages(roomNumber) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('대화내용');
    
    if (!sheet) {
      return { success: true, messages: [], lastTimestamp: 0 };
    }
    
    // 타입 변환
    roomNumber = Number(roomNumber);
    
    const data = sheet.getDataRange().getValues();
    const messages = [];
    let lastTimestamp = 0;
    
    for (let i = 1; i < data.length; i++) {
      const rowRoomNum = Number(data[i][0]);
      
      if (rowRoomNum === roomNumber) {
        const msgTimestamp = Number(data[i][4]) || 0;
        
        messages.push({
          userName: String(data[i][1]),
          content: String(data[i][2]),
          time: String(data[i][3]),
          timestamp: msgTimestamp
        });
        
        if (msgTimestamp > lastTimestamp) {
          lastTimestamp = msgTimestamp;
        }
      }
    }
    
    return { success: true, messages: messages, lastTimestamp: lastTimestamp };
  } catch (error) {
    return { success: false, messages: [], lastTimestamp: 0, error: error.toString() };
  }
}

// 대화방 목록 정보 (각 방의 최신 메시지)
function getRoomList() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('대화내용');
    
    const rooms = [
      { id: 1, name: '1번 대화방', lastMessage: '', lastTime: '', count: 0 },
      { id: 2, name: '2번 대화방', lastMessage: '', lastTime: '', count: 0 },
      { id: 3, name: '3번 대화방', lastMessage: '', lastTime: '', count: 0 },
      { id: 4, name: '4번 대화방', lastMessage: '', lastTime: '', count: 0 },
      { id: 5, name: '5번 대화방', lastMessage: '', lastTime: '', count: 0 }
    ];
    
    if (!sheet) {
      return { success: true, rooms: rooms };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const roomNum = Number(data[i][0]);
      if (roomNum >= 1 && roomNum <= 5) {
        rooms[roomNum - 1].count++;
        rooms[roomNum - 1].lastMessage = String(data[i][2]);
        rooms[roomNum - 1].lastTime = String(data[i][3]);
      }
    }
    
    return { success: true, rooms: rooms };
  } catch (error) {
    return { success: false, rooms: [], error: error.toString() };
  }
}

// 테스트 함수 (디버깅용)
function testSendMessage() {
  const result = sendMessage(1, '테스트', '안녕하세요!');
  Logger.log(result);
  return result;
}

function testGetMessages() {
  const result = getAllMessages(1);
  Logger.log(result);
  return result;
}
