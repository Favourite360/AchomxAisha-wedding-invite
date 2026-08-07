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

   NOTE: every time you edit this script you must Deploy -> Manage
   deployments -> edit -> Version: New version, or the live site
   keeps running the old code.

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

/* Lets you confirm the deployment is alive by opening the /exec
   URL in a browser. Should show {"result":"ok",...}. */
function doGet() {
  return reply('ok', 'endpoint is live — post to this URL');
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
