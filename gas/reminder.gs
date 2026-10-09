/**
 * 耳つぼ「シールを外す日」お知らせメール 自動送信スクリプト
 * (Customer Card から予約を受け取り、当日の朝にあなたのGmailから送信します)
 *
 * 【セットアップ手順】は gas/README.md を見てください。
 *
 * 保存先: このスクリプトを作ったGoogleアカウントのスプレッドシート「reminders」シート
 * 送信元: このスクリプトを実行するGoogleアカウント(Gmail)
 */

var SHEET_NAME = 'reminders';
var TZ = 'Asia/Tokyo';
// 列: A=key, B=sendDate, C=to, D=subject, E=body, F=sender, G=status, H=sentAt
var HEADER = ['key', 'sendDate', 'to', 'subject', 'body', 'sender', 'status', 'sentAt'];

function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('耳つぼお知らせメール予約');
    props.setProperty('SHEET_ID', ss.getId());
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Customer Card からの依頼を受け取る
function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!token || req.token !== token) return json_({ ok: false, error: 'unauthorized' });

    var sh = getSheet_();
    var rows = sh.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === req.key) { rowIndex = i + 1; break; }
    }

    if (req.action === 'schedule') {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(req.sendDate) || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(req.to)) {
        return json_({ ok: false, error: 'bad_request' });
      }
      // 日付が文字列のままになるよう、先頭にアポストロフィ相当の書式を指定
      var values = [[req.key, req.sendDate, req.to, req.subject, req.body, req.sender || '', 'pending', '']];
      if (rowIndex === -1) {
        sh.appendRow(values[0]);
        sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(req.sendDate);
      } else {
        sh.getRange(rowIndex, 1, 1, 8).setValues(values);
        sh.getRange(rowIndex, 2).setNumberFormat('@').setValue(req.sendDate);
      }
      return json_({ ok: true });
    }

    if (req.action === 'cancel') {
      if (rowIndex !== -1 && rows[rowIndex - 1][6] === 'pending') {
        sh.getRange(rowIndex, 7).setValue('cancelled');
      }
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// 毎朝実行: 送信日が今日以前で、まだ送っていない予約を送信する
function sendDueReminders() {
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var sh = getSheet_();
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (r[6] !== 'pending') continue;
    var sendDate = r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1]);
    if (sendDate > today) continue;
    // 送信日から3日以上過ぎたものは送らず「期限切れ」にする(古いお知らせを送らないため)
    var late = (new Date(today) - new Date(sendDate)) / 86400000;
    if (late > 3) { sh.getRange(i + 1, 7).setValue('expired'); continue; }
    try {
      var opts = {};
      if (r[5]) opts.name = r[5];
      GmailApp.sendEmail(r[2], r[3], r[4], opts);
      sh.getRange(i + 1, 7, 1, 2).setValues([['sent', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')]]);
    } catch (err) {
      sh.getRange(i + 1, 7).setValue('error: ' + String(err).slice(0, 80));
    }
  }
}

// 最初に1回だけ手動で実行: 毎朝9時台に sendDueReminders を動かす設定
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDueReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDueReminders').timeBased().everyDays(1).atHour(9).inTimezone(TZ).create();
}

// 最初に1回だけ手動で実行: 合言葉を設定(下の文字列を自分で決めた合言葉に書き換えてから実行)
function setToken() {
  PropertiesService.getScriptProperties().setProperty('TOKEN', 'ここを自分の合言葉に書き換える');
}
