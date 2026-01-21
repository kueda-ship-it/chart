// シートを取得または作成する関数
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    // ヘッダー行を設定
    if (sheetName === 'users') {
      sheet.getRange(1, 1, 1, 5).setValues([['id', 'name', 'pass', 'role', 'data']]);
    } else if (sheetName === 'properties') {
      sheet.getRange(1, 1, 1, 2).setValues([['id', 'data']]);
    } else if (sheetName === 'masters') {
      sheet.getRange(1, 1, 1, 4).setValues([['key', 'value', 'type', 'data']]);
    }
  }
  return sheet;
}

function doPost(e) {
  // OPTIONSリクエスト（プリフライト）を処理
  // Google Apps Scriptでは、OPTIONSリクエストはpostDataが空で来る可能性がある
  if (!e.postData || !e.postData.contents) {
    // プリフライトリクエストの可能性がある場合は空のレスポンスを返す
    return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.JSON);
  }
  
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // CORSヘッダーを設定
    const output = ContentService.createTextOutput();
    
    const contents = JSON.parse(e.postData.contents);
    const type = contents.type; 
    
    if (!type) {
      return output.setContent(JSON.stringify({result: 'error', error: 'typeパラメータが指定されていません'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      return output.setContent(JSON.stringify({result: 'error', error: 'スプレッドシートが見つかりません'})).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (type === 'save') {
      if (!contents.state) {
        return output.setContent(JSON.stringify({result: 'error', error: 'stateデータがありません'})).setMimeType(ContentService.MimeType.JSON);
      }
      
      const state = contents.state;
      
      // ユーザーデータを保存
      const usersSheet = getOrCreateSheet(spreadsheet, 'users');
      usersSheet.clear();
      usersSheet.getRange(1, 1, 1, 5).setValues([['id', 'name', 'pass', 'role', 'data']]);
      if (state.masterUsers && state.masterUsers.length > 0) {
        const userData = state.masterUsers.map(u => [u.id, u.name, u.pass || '', u.role || 'user', JSON.stringify(u)]);
        usersSheet.getRange(2, 1, userData.length, 5).setValues(userData);
      }
      
      // プロパティデータを保存
      const propsSheet = getOrCreateSheet(spreadsheet, 'properties');
      propsSheet.clear();
      propsSheet.getRange(1, 1, 1, 2).setValues([['id', 'data']]);
      if (state.properties && state.properties.length > 0) {
        const propData = state.properties.map(p => [p.id, JSON.stringify(p)]);
        propsSheet.getRange(2, 1, propData.length, 2).setValues(propData);
      }
      
      // マスターデータとその他の設定を保存
      const mastersSheet = getOrCreateSheet(spreadsheet, 'masters');
      mastersSheet.clear();
      mastersSheet.getRange(1, 1, 1, 4).setValues([['key', 'value', 'type', 'data']]);
      const masterData = [
        ['masterL1', JSON.stringify(state.masterL1 || []), 'array', JSON.stringify(state.masterL1 || [])],
        ['masterL2', JSON.stringify(state.masterL2 || []), 'array', JSON.stringify(state.masterL2 || [])],
        ['masterL3', JSON.stringify(state.masterL3 || []), 'array', JSON.stringify(state.masterL3 || [])],
        ['masterProperties', JSON.stringify(state.masterProperties || []), 'array', JSON.stringify(state.masterProperties || [])],
        ['viewStart', state.viewStart || '', 'string', state.viewStart || ''],
        ['viewEnd', state.viewEnd || '', 'string', state.viewEnd || ''],
        ['zoom', state.zoom || 1.0, 'number', state.zoom || 1.0],
        ['collapsedIds', JSON.stringify(state.collapsedIds ? Array.from(state.collapsedIds) : []), 'array', JSON.stringify(state.collapsedIds ? Array.from(state.collapsedIds) : [])],
        ['cloudUrl', state.cloudUrl || '', 'string', state.cloudUrl || ''],
        ['lastSync', state.lastSync || '', 'string', state.lastSync || ''],
        ['masterTab', state.masterTab || 'users', 'string', state.masterTab || 'users']
      ];
      mastersSheet.getRange(2, 1, masterData.length, 4).setValues(masterData);
      
      return output.setContent(JSON.stringify({result: 'success'})).setMimeType(ContentService.MimeType.JSON);
      
    } else if (type === 'load') {
      // ユーザーデータを読み込み
      const usersSheet = spreadsheet.getSheetByName('users');
      let masterUsers = [];
      if (usersSheet && usersSheet.getLastRow() > 1) {
        const userRows = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 5).getValues();
        masterUsers = userRows.map(row => {
          try {
            return JSON.parse(row[4]); // data列からJSONをパース
          } catch (e) {
            // フォールバック: 個別の列から構築
            return { id: row[0], name: row[1], pass: row[2] || '', role: row[3] || 'user' };
          }
        });
      }
      
      // プロパティデータを読み込み
      const propsSheet = spreadsheet.getSheetByName('properties');
      let properties = [];
      if (propsSheet && propsSheet.getLastRow() > 1) {
        const propRows = propsSheet.getRange(2, 1, propsSheet.getLastRow() - 1, 2).getValues();
        properties = propRows.map(row => {
          try {
            return JSON.parse(row[1]); // data列からJSONをパース
          } catch (e) {
            return null;
          }
        }).filter(p => p !== null);
      }
      
      // マスターデータとその他の設定を読み込み
      const mastersSheet = spreadsheet.getSheetByName('masters');
      const state = {
        masterUsers: masterUsers.length > 0 ? masterUsers : [],
        properties: properties.length > 0 ? properties : [],
        masterL1: [],
        masterL2: [],
        masterL3: [],
        masterProperties: [],
        viewStart: null,
        viewEnd: null,
        zoom: 1.0,
        collapsedIds: [],
        cloudUrl: '',
        lastSync: null,
        masterTab: 'users'
      };
      
      if (mastersSheet && mastersSheet.getLastRow() > 1) {
        const masterRows = mastersSheet.getRange(2, 1, mastersSheet.getLastRow() - 1, 4).getValues();
        masterRows.forEach(row => {
          const key = row[0];
          const value = row[1];
          try {
            if (key === 'masterL1') state.masterL1 = JSON.parse(value);
            else if (key === 'masterL2') state.masterL2 = JSON.parse(value);
            else if (key === 'masterL3') state.masterL3 = JSON.parse(value);
            else if (key === 'masterProperties') state.masterProperties = JSON.parse(value);
            else if (key === 'viewStart') state.viewStart = value || null;
            else if (key === 'viewEnd') state.viewEnd = value || null;
            else if (key === 'zoom') state.zoom = parseFloat(value) || 1.0;
            else if (key === 'collapsedIds') state.collapsedIds = JSON.parse(value);
            else if (key === 'cloudUrl') state.cloudUrl = value || '';
            else if (key === 'lastSync') state.lastSync = value || null;
            else if (key === 'masterTab') state.masterTab = value || 'users';
          } catch (e) {
            console.error('Error parsing master data for key:', key, e);
          }
        });
      }
      
      // データが存在しない場合はnullを返す
      if (state.masterUsers.length === 0 && state.properties.length === 0) {
        return output.setContent(JSON.stringify({result: 'success', data: null})).setMimeType(ContentService.MimeType.JSON);
      }
      
      return output.setContent(JSON.stringify({result: 'success', data: state})).setMimeType(ContentService.MimeType.JSON);
      
    } else {
      return output.setContent(JSON.stringify({result: 'error', error: '不明なtype: ' + type})).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({result: 'error', error: error.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput("GAS is running. Use POST for data operations.");
}
