var ss = SpreadsheetApp.getActiveSpreadsheet();
var url = "https://xxxxxx.cybozu.com/k/v1/records.json"; // [サブドメイン]を適宜設定

var appConfig = {
  '1':   { sheetName: '企業DB', 
           apiToken : 'xxxx', 
           fields   : ['企業ID', '名前', '業種'],
           orderBy  : '企業ID'  
         },
  '2':   { sheetName: '支援実績', 
           apiToken : 'yyyy', 
           fields   : ['支援日', '企業ID', '支援分類'],
           orderBy  : '支援日'  
         }
};

// 現在の日付から、1ヶ月前の月初～今月末までのデータを取得する。
// 月初に実行
function monthlyRefleshKintoneRecords() {
  var daterange_start = Moment.moment().subtract(1, 'months').startOf("months").format("YYYY-MM-DD");
  var daterange_end = Moment.moment().endOf("months").format("YYYY-MM-DD");
  refleshKintoneRecords(daterange_start,daterange_end)
}

function refleshKintoneRecords(daterange_start, daterange_end) {
  // APIトークンをappConfigから動的に取得
  Object.keys(appConfig).forEach(function(appId) {
    var config = appConfig[appId];
    var sheetName = config.sheetName;
    var fields = ['レコード番号', ...config.fields]; // レコード番号を先頭に追加
    var apiToken = config.apiToken; // APIトークンを取得
    var query = '';
    if (sheetName === '支援実績') {
      query = `${config.orderBy} >= "${daterange_start}" and ${config.orderBy} < "${daterange_end}"`;
      deleteSeetDataBetween(appId,daterange_start,daterange_end)
    }
    writeRecordsToSheet(sheetName, appId, apiToken, query, fields);
    sortSheet(appId)
  });
}

function writeRecordsToSheet(sheetName, appId, apiToken, query, fieldNames) {
  var sheet = ss.getSheetByName(sheetName);
  var records = getKintoneRecord(appId, apiToken, query);
  records.forEach(record => {
    var rowValues = fieldNames.map(fieldName => {
      // record[fieldName] が存在し、valueがあるか確認
      if (record[fieldName] && record[fieldName].value !== undefined) {
        // valueが配列の場合、要素をコンマで繋げる
        if (Array.isArray(record[fieldName].value)) {
          return record[fieldName].value.join(', ');
        } else {
          // 配列ではない場合はそのまま文字列に変換
          return String(record[fieldName].value);
        }
      } else {
        // record[fieldName] が存在しないかvalueがない場合は空文字
        return '';
      }
    });
    var desRow = findRow(sheet, Number(record['$id'].value), 1) || sheet.getLastRow() + 1;
    sheet.getRange(desRow, 1, 1, rowValues.length).setValues([rowValues]);
  });
}

function findRow(sheet, val, col) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][col-1] === val) return i + 1;
  }
  return 0;
}

function getKintoneRecord(appId, apiToken, query) {
  var offset = 0;
  var limit = 100;
  var records = [];
  var params = {
    method: "GET",
    headers: { "X-Cybozu-API-Token": apiToken },
    muteHttpExceptions: true
  };

  do {
    var fetchUrl = `${url}?app=${appId}&query=${encodeURIComponent(query)} order by $id asc limit ${limit} offset ${offset}`;
    var response = UrlFetchApp.fetch(fetchUrl, params);
    var data = JSON.parse(response.getContentText());

    if (data.records.length > 0) {
      records = records.concat(data.records);
      offset += data.records.length;
    } else {
      break;
    }
  } while (data.records.length === limit);

  return records;
}

function deleteSeetDataBetween(recordId, daterange_start, daterange_end) {
  var sheet = ss.getSheetByName(appConfig[recordId].sheetName);
  var orderBy = appConfig[recordId].orderBy;
  var columnIndex = appConfig[recordId].fields.indexOf(orderBy) + 2; // インデックスは1ベースで計算
  var lastRow = sheet.getLastRow();
  var resRange_top = 1;
  var resRange_bottom = 1;
  
  for (var i = 2; i < lastRow; i++) {
    var cur_date = Moment.moment(sheet.getRange(i, columnIndex).getValue());
    
    if (resRange_top == 1 && cur_date.isBefore(daterange_end)) {
      resRange_top = i;
    } else if (cur_date.isBefore(daterange_start)) {
      resRange_bottom = i - 1;
      break;
    }
  }

  if (resRange_top != 1 && resRange_bottom == 1) {
    resRange_bottom = lastRow;  // Adjust to delete to the end if no end point is found
  }

  if (resRange_top != 1 && resRange_bottom != 1) {
    sheet.deleteRows(resRange_top, resRange_bottom - resRange_top + 1);
  }
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var appId = data.app.id;
  var config = appConfig[appId];
  var sheet = ss.getSheetByName(config.sheetName);
  var action = data.type;

  switch (action) {
    case 'ADD_RECORD':
    case 'UPDATE_RECORD':
      processAddOrUpdateRecord(data, sheet, config, action);
      break;
    case 'DELETE_RECORD':
      var recordId = data.recordId;
      deleteRecord(sheet, recordId);
      break;
  }
  sortSheet(appId)
}

function processAddOrUpdateRecord(data, sheet, config, action) {
  var record = data.record;
  var recordId = record.$id.value;
  var values = [recordId];

  config.fields.forEach(function(fieldName) {
    let valueToAdd = ''; // デフォルト値として空文字を設定
    if (record[fieldName] && record[fieldName].value !== undefined) {
      // record[fieldName].valueが配列かどうかを確認
      if (Array.isArray(record[fieldName].value)) {
        if (record[fieldName].value.length === 0) {
          // 配列が空の場合は空文字を設定
          valueToAdd = '';
        } else {
          // 配列の値がある場合はコンマで繋ぐ
          valueToAdd = record[fieldName].value.join(', ');
        }
      } else {
        // 配列ではない場合は元の値をそのまま設定
        valueToAdd = record[fieldName].value;
      }
    }
    // 決定された値をvaluesに追加
    values.push(valueToAdd);
  });

  if (action === 'ADD_RECORD') {
    sheet.appendRow(values);
  } else {
    updateRecord(sheet, values);
  }
}

function updateRecord(sheet, values) {
  var range = sheet.getDataRange();
  var data = range.getValues();
  var recordId = values[0];

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == recordId) {
      sheet.getRange(i + 1, 1, 1, values.length).setValues([values]);
      return;
    }
  }
  // 一致するレコードがない場合は追加
  sheet.appendRow(values);
}

function deleteRecord(sheet, recordId) {
  var range = sheet.getDataRange();
  var data = range.getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == recordId) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

function sortSheet(recordId) {
  var sheetName = appConfig[recordId].sheetName;
  var orderBy = appConfig[recordId].orderBy;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var range = sheet.getDataRange();
  var data = range.getValues();
  
  var columnIndex = appConfig[recordId].fields.indexOf(orderBy) + 2; // インデックスは1ベースで計算
  if (columnIndex > 0) {
    // 降順にソートする
    range.sort({column: columnIndex, ascending: false});
    console.log("Sorted sheet " + sheetName + " by " + orderBy + " in descending order.");
  } else {
    console.log("Sort column " + orderBy + " not found in the sheet " + sheetName);
  }
}
