// ===== Google Apps Script for IE Survey Data Collection =====
// Paste this into Extensions > Apps Script in your Google Sheet
// Deploy as: Web app | Execute as: Me | Access: Anyone

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'Responses';

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var headers = getHeaders();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var row = flattenData(data);

    // Find existing row by sessionId (column A, starting row 2)
    var sessionId = data.sessionId || '';
    var lastRow = sheet.getLastRow();
    var existingRow = -1;

    if (lastRow >= 2) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (ids[i][0] === sessionId) {
          existingRow = i + 2;
          break;
        }
      }
    }

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getHeaders() {
  return [
    'sessionId',
    'startTime',
    'endTime',
    'totalSeconds',
    'currentScreen',

    // S1: About You
    'q1_role',
    'q2_seniority',
    'q3_decision',
    'q4_orgsize',
    'q5_sector',
    'q5a_subsector',

    // S2: IE Relationship
    'q6_ie_relationship',

    // S3: Name Perception
    'q7_services',
    'q8_strategic_stretch',
    'q9_name_limitation',

    // S4: People Priorities
    'q10_priorities',
    'q11_inclusion_trend',
    'q12_trend_reason',

    // S5: Budget
    'q13_budget',
    'q14_headcount',
    'q15_budget_reason',
    'q16_roi_pressure',
    'q17_approval_difficulty',

    // S6: Language
    'q18_terms',
    'q19_language_shifting',
    'q19_dei_reframing',
    'q19_ticking_boxes',
    'q19_words_misunderstood',
    'q19_diversity_negative',
    'q19_inclusion_positive',

    // S7: Membership
    'q20_relationship_pref',
    'q21_membership_values',

    // S8: IE Brand (routed)
    'q22_credible',
    'q22_trusted',
    'q22_expert',
    'q22_forward_thinking',
    'q22_good_value',
    'q23_modern',
    'q23_professional',
    'q23_clear',
    'q23_engaging',
    'q23_work_with',
    'q23a_visual_thoughts',
    'q24_confusion',

    // S9: Opt-ins
    'q25_report_optin',
    'q25a_email',
    'q26_webinar_optin',

    'timestamp'
  ];
}

function flattenData(data) {
  var headers = getHeaders();
  var row = [];

  for (var i = 0; i < headers.length; i++) {
    var key = headers[i];

    if (key === 'timestamp') {
      row.push(new Date().toISOString());
    } else if (key === 'startTime' && data.startTime) {
      row.push(new Date(data.startTime).toISOString());
    } else if (key === 'endTime' && data.endTime) {
      row.push(new Date(data.endTime).toISOString());
    } else {
      row.push(data[key] != null ? data[key] : '');
    }
  }

  return row;
}

function doGet(e) {
  return ContentService.createTextOutput('IE Survey data collector is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}
