// 画面を表示する
function doGet() {
  // 現在のスプレッドシートの名前を取得
  const ssName = SpreadsheetApp.getActiveSpreadsheet().getName();
  
  const template = HtmlService.createTemplateFromFile('index');
  // HTML側でスプレッドシート名を使えるように変数を渡す
  template.ssName = ssName;
  
  return template.evaluate()
    .setTitle(ssName) // ブラウザのタブ名もファイル名にする
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// スプレッドシートからデータを取得
function getSongData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('リスト');
  const values = sheet.getDataRange().getValues();
  values.shift(); // ヘッダーを削除
  return values.map(row => ({
    artist: row[0],
    title: row[1],
    votes: row[2] || 0,
    isDone: row[3] === true
  }));
}

// 新しい曲を追加
function addSong(artist, title) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リスト');
  sheet.appendRow([artist, title, 0, false]);
  return getSongData();
}

// 投票数を+1
function voteSong(title) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リスト');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === title) {
      sheet.getRange(i + 1, 3).setValue(Number(data[i][2] || 0) + 1);
      break;
    }
  }
  return getSongData();
}
// 全ての投票を0にリセットする（コード.gsに追加）
function resetAllVotes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リスト');
  const lastRow = sheet.getLastRow();
  
  // データがある場合のみ実行
  if (lastRow > 1) {
    // 2行目から最終行までの、3列目（C列：投票数）の範囲を取得
    const range = sheet.getRange(2, 3, lastRow - 1, 1);
    
    // 全て0で埋めた配列を作成して一括書き込み
    const resetData = new Array(lastRow - 1).fill([0]);
    range.setValues(resetData);
  }
  
  // 最新のデータを読み直して画面に返す
  return getSongData();
}
// 歌唱済みを切り替え
function toggleDone(title, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リスト');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === title) {
      sheet.getRange(i + 1, 4).setValue(status);
      break;
    }
  }
  return getSongData();
}
// 曲を削除する
function deleteSong(title) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リスト');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === title) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return getSongData();
}

function restoreSongData(title, votes, isDone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リスト');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === title) {
      sheet.getRange(i + 1, 3).setValue(votes);     // votes
      sheet.getRange(i + 1, 4).setValue(isDone);    // isDone
      break;
    }
  }
  return getSongData();
}
