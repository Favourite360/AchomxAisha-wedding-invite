/* ============================================================
   Nnamdi & Aisha — wedding site form collector
   ------------------------------------------------------------
   Receives both the RSVP form and the Gift Registry "Contribute"
   form, and writes each to its own tab in ONE spreadsheet.

   SETUP (about 10 minutes)
   ------------------------------------------------------------
   1. Go to sheets.new to create a spreadsheet. Name it something
      like "Nnamdi & Aisha — Wedding Responses".

   2. Copy its ID from the address bar. In a URL like
        https://docs.google.com/spreadsheets/d/1AbC...XyZ/edit?gid=0
      the ID is ONLY the part between /d/ and the next slash —
      not the whole address. Paste it into SHEET_ID below.
      (Pasting the full URL also works; the script pulls the ID
      out for you.)

   3. In that spreadsheet: Extensions -> Apps Script.
      Delete whatever is in the editor, paste this whole file in,
      and click the save icon.

   4. Click Deploy -> New deployment.
        - Click the gear next to "Select type", choose "Web app"
        - Description:    wedding forms
        - Execute as:     Me
        - Who has access: Anyone            <- MUST be "Anyone",
                                               not "Anyone with
                                               Google account"
      Click Deploy. Google will ask you to authorise it — that is
      it asking permission to write to your own sheet. Approve it.

   5. Copy the Web app URL it gives you. It ends in /exec.
      That URL goes into index.html as ENDPOINT.

   NOTE — THE ONE THAT CATCHES EVERYONE:
   Editing and saving this file does NOT change what the live URL
   runs. A deployment is pinned to a numbered version. After any
   edit you MUST do:
       Deploy -> Manage deployments -> pencil icon
              -> Version: New version -> Deploy
   Saving alone leaves the old code serving, which looks exactly
   like the script being broken.

   To check which code is live, open the /exec URL in a browser.
   It reports its own version plus whether it can reach the sheet.

   The two tabs ("RSVP" and "Gifts") are created automatically on
   the first submission, with bold frozen headers. You do not need
   to make them yourself.
   ============================================================ */

const SHEET_ID = '14qG_Ni9zK2nXxF-EzJKYPxTULax5B2YP703NKJTLYI4';

/* Accepts either a bare ID or a full spreadsheet URL, so pasting
   the whole address from the tab bar still works. */
function resolveSheetId(value) {
  const found = String(value).match(/\/d\/([a-zA-Z0-9-_]+)/);
  return found ? found[1] : String(value).trim();
}

/* Which tab each form goes to, and the columns it writes. */
const TABS = {
  rsvp: {
    name: 'RSVP',
    headers: ['Received', 'Name', 'Attending']
  },
  gift: {
    name: 'Gifts',
    headers: ['Received', 'Gift item', 'Name', 'Message']
  }
};

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return reply('error', 'no data received');
    }

    const data = JSON.parse(e.postData.contents);

    /* Honeypot: a real person never sees or fills this field, so
       anything with it filled in is a bot. Return success so the
       bot does not learn it was rejected, but write nothing. */
    if (data.website) {
      return reply('ok', 'ignored');
    }

    const config = TABS[data.form];
    if (!config) {
      return reply('error', 'unknown form: ' + data.form);
    }

    const sheet = getTab(config);
    const now = Utilities.formatDate(new Date(), 'Africa/Lagos', 'yyyy-MM-dd HH:mm:ss');

    const row = (data.form === 'rsvp')
      ? [now, clean(data.name), clean(data.attending)]
      : [now, clean(data.item), clean(data.name), clean(data.message)];

    sheet.appendRow(row);
    return reply('ok', 'saved');

  } catch (err) {
    return reply('error', String(err));
  }
}

/* Bump this whenever you edit the file. Opening the /exec URL shows
   it back, which is how you tell whether the deployment is serving
   your latest code or an older pinned version. */
const SCRIPT_VERSION = 'v3';

/* Self-test. Open the /exec URL in a browser and it reports whether
   the running code is current AND whether it can actually reach the
   spreadsheet — the two things that fail silently on POST, because
   a no-cors request cannot read the response. */
function doGet() {
  const report = {
    version: SCRIPT_VERSION,
    resolvedId: resolveSheetId(SHEET_ID),
    sheetOpened: false,
    tabs: [],
    canWrite: false,
    error: null
  };

  try {
    const book = SpreadsheetApp.openById(resolveSheetId(SHEET_ID));
    report.sheetOpened = true;
    report.spreadsheetName = book.getName();
    report.tabs = book.getSheets().map(function (s) {
      return s.getName() + ' (' + s.getLastRow() + ' rows)';
    });

    /* Prove write access for real, then immediately undo it. */
    const probe = getTab(TABS.rsvp);
    probe.appendRow(['SELF-TEST', 'delete me', '']);
    SpreadsheetApp.flush();
    probe.deleteRow(probe.getLastRow());
    report.canWrite = true;

  } catch (err) {
    report.error = String(err);
  }

  return ContentService
    .createTextOutput(JSON.stringify(report, null, 1))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------
   ONE-OFF CLEANUP — removes the test rows left behind while the
   forms were being wired up.

   HOW TO RUN (takes 10 seconds, no deployment needed):
     1. In the Apps Script editor, pick "cleanupTestRows" from the
        function dropdown in the toolbar.
     2. Click Run.
     3. Read the result in the Execution log at the bottom.

   Running a function in the editor uses the code as saved, so this
   does NOT disturb the live deployment. Your forms keep working
   throughout.

   It only deletes rows whose Name cell begins with one of the
   markers below. A real guest's row can never match, so this is
   safe to run more than once.
   ------------------------------------------------------------ */
const TEST_MARKERS = ['ZZ TEST', 'ZZ FINAL TEST', 'SELF-TEST', 'ZZ CORS PROBE'];

function cleanupTestRows() {
  const book = SpreadsheetApp.openById(resolveSheetId(SHEET_ID));
  const summary = [];

  ['RSVP', 'Gifts'].forEach(function (tabName) {
    const sheet = book.getSheetByName(tabName);
    if (!sheet) { summary.push(tabName + ': tab not found'); return; }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { summary.push(tabName + ': nothing to do'); return; }

    /* Name sits in column B on RSVP, column C on Gifts. */
    const nameCol = (tabName === 'RSVP') ? 2 : 3;
    const names = sheet.getRange(2, nameCol, lastRow - 1, 1).getValues();

    /* Walk bottom-up so deleting a row cannot shift the ones we
       have not checked yet. */
    let removed = 0;
    for (let i = names.length - 1; i >= 0; i--) {
      const value = String(names[i][0] || '').trim();
      const isTest = TEST_MARKERS.some(function (marker) {
        return value.indexOf(marker) === 0;
      });
      if (isTest) {
        sheet.deleteRow(i + 2);
        removed++;
      }
    }
    summary.push(tabName + ': removed ' + removed +
                 ', ' + (sheet.getLastRow() - 1) + ' real rows left');
  });

  const message = summary.join('  |  ');
  Logger.log(message);
  return message;
}

/* Finds the tab, creating it (with headers) the first time. */
function getTab(config) {
  const book = SpreadsheetApp.openById(resolveSheetId(SHEET_ID));
  let sheet = book.getSheetByName(config.name);

  if (!sheet) {
    sheet = book.insertSheet(config.name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(config.headers);
    sheet.getRange(1, 1, 1, config.headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, config.headers.length);
  }

  return sheet;
}

/* Trim, cap length, and prefix anything Sheets would treat as a
   formula. Without this a submitted name like "=1+1" or "+A1"
   becomes a live formula in the couple's spreadsheet. */
function clean(value) {
  let text = String(value == null ? '' : value).trim().slice(0, 2000);
  if (/^[=+\-@]/.test(text)) {
    text = "'" + text;
  }
  return text;
}

function reply(result, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ result: result, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
