/**
 * Handles form submissions from both games (Scope Detective + ESG Pillars Quest).
 * Routes to the correct sheet tab based on the "quest" field sent from the game.
 */
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);

  if (data.quest === 'ESG_Pillars') {
    var sheet = ss.getSheetByName('ESG_Responses') || ss.insertSheet('ESG_Responses');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Dept', 'Q1Answer', 'Q2Links',
        'AutoScore', 'Q3Answer', 'Q3Score', 'Q4Answer', 'Q4Score', 'FinishedAt', 'Language']);
    }
    sheet.appendRow([
      new Date(), data.name, data.email, data.dept,
      data.q1Answer, data.q2Links, data.autoScore,
      data.q3Answer, data.q3Score, data.q4Answer, data.q4Score,
      data.finishedAt, data.lang
    ]);
  } else {
    // Original Scope Detective game — writes to the "Responses" tab, unchanged.
    var sheet = ss.getSheetByName('Responses') || ss.insertSheet('Responses');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Dept', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6',
        'QuizScore', 'MatchScore', 'AutoTotal', 'Part3Answer', 'Part3Score', 'FinishedAt', 'Language']);
    }
    var q = data.quizAnswers || {};
    sheet.appendRow([
      new Date(), data.name, data.email, data.dept,
      q.q1, q.q2, q.q3, q.q4, q.q5, q.q6,
      data.quizScore, data.matchScore, data.autoTotal,
      data.part3Answer, '', data.finishedAt, data.lang
    ]);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Serves the leaderboard as JSON (name / dept / score only — no email).
 * Called by the game with: WEBHOOK_URL?quest=ESG_Pillars
 * The sheet itself never needs to be published or made public.
 */
function doGet(e) {
  var quest = (e.parameter.quest || 'ESG_Pillars');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = (quest === 'ESG_Pillars') ? 'ESG_Responses' : 'Responses';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var headers = data[0];
  var nameIdx = headers.indexOf('Name');
  var deptIdx = headers.indexOf('Dept');
  var scoreIdx = headers.indexOf(quest === 'ESG_Pillars' ? 'AutoScore' : 'AutoTotal');

  var players = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[nameIdx]) continue;
    players.push({
      name: row[nameIdx],
      dept: deptIdx >= 0 ? row[deptIdx] : '',
      score: scoreIdx >= 0 ? row[scoreIdx] : 0
    });
  }
  players.sort(function (a, b) { return b.score - a.score; });
  players = players.slice(0, 20);

  return ContentService.createTextOutput(JSON.stringify(players))
    .setMimeType(ContentService.MimeType.JSON);
}
