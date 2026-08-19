/* ============================================================
   Nnamdi & Aisha — wedding site backend
   ------------------------------------------------------------
   One spreadsheet, three tabs:

     Guests  — the invitation list. YOU fill this in. Nobody whose
               number is not here can RSVP.
     RSVP    — who replied, filled in automatically.
     Gifts   — registry contributions, filled in automatically.

   ------------------------------------------------------------
   SETUP
   ------------------------------------------------------------
   1. Paste this whole file into the Apps Script editor
      (Extensions -> Apps Script), replacing everything. Save.

   2. Run setupTabs once:
        - pick "setupTabs" in the function dropdown
        - click Run
      That creates all three tabs with the right headers.

   3. Fill in the Guests tab. One row per invited guest:

        Name             Phone         Code     Admits
        Emmanuel         09138714023   NA-001   1

      Phone can be written any way — 0803 123 4567,
      +2348031234567, 234-803-123-4567 all work.
      Admits blank means 1. Set 2 or 3 for guests who may
      bring someone.
      Code is what prints large on their access card.

   4. Deploy -> New deployment -> Web app
        Execute as:     Me
        Who has access: Anyone      <- plain "Anyone"

   ------------------------------------------------------------
   AFTER ANY EDIT to this file you must redeploy, or the live
   site keeps running the old code:
        Deploy -> Manage deployments -> pencil
               -> Version: New version -> Deploy
   The URL does not change.

   To check what is live, open the /exec URL in a browser. It
   reports its own version and whether it can reach the sheet.
   ============================================================ */

const SHEET_ID = '14qG_Ni9zK2nXxF-EzJKYPxTULax5B2YP703NKJTLYI4';

const SCRIPT_VERSION = 'v5 — guest verification';

/* Accepts a bare ID or a full spreadsheet URL. */
function resolveSheetId(value) {
  const found = String(value).match(/\/d\/([a-zA-Z0-9-_]+)/);
  return found ? found[1] : String(value).trim();
}

const GUESTS_TAB = {
  name: 'Guests',
  headers: ['Name', 'Phone', 'Code', 'Admits']
};

const TABS = {
  rsvp: {
    name: 'RSVP',
    headers: ['Received', 'Name', 'Phone', 'Attending', 'Code', 'Admits']
  },
  gift: {
    name: 'Gifts',
    headers: ['Received', 'Gift item', 'Name', 'Message']
  }
};

/* ------------------------------------------------------------
   Run this once from the editor to build all three tabs.
   ------------------------------------------------------------ */
function setupTabs() {
  const made = [GUESTS_TAB, TABS.rsvp, TABS.gift].map(function (cfg) {
    const sheet = getTab(cfg);
    return cfg.name + ' (' + (sheet.getLastRow() - 1) + ' rows)';
  });
  const msg = 'Ready: ' + made.join(', ') + '. Now fill in the Guests tab.';
  Logger.log(msg);
  return msg;
}

/* ------------------------------------------------------------
   Phone matching. Nigerian numbers are written every which way,
   so reduce each to its last 10 digits before comparing —
   otherwise most real guests would be turned away.
   ------------------------------------------------------------ */
function normalisePhone(value) {
  const digits = String(value == null ? '' : value).replace(/\D/g, '');
  return digits.length > 10 ? digits.slice(-10) : digits;
}

/* Looks one number up in Guests. Returns that guest or null.
   Never returns the list. */
function findGuest(phone) {
  const key = normalisePhone(phone);
  if (key.length < 10) return null;

  const sheet = getTab(GUESTS_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (normalisePhone(rows[i][1]) === key) {
      const admits = parseInt(rows[i][3], 10);
      return {
        name:   String(rows[i][0] || '').trim(),
        phone:  String(rows[i][1] || '').trim(),
        code:   String(rows[i][2] || '').trim(),
        admits: (admits > 0) ? admits : 1
      };
    }
  }
  return null;
}

/* One row per guest, keyed on phone. Someone replying again — to
   change their answer or re-download their card — updates their
   existing row rather than adding a second, so the row count is
   always the real headcount. */
function upsertRsvp(guest, givenName, attending) {
  const sheet = getTab(TABS.rsvp);
  const now = Utilities.formatDate(new Date(), 'Africa/Lagos', 'yyyy-MM-dd HH:mm:ss');
  const values = [now, givenName, guest.phone, attending, guest.code, guest.admits];

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const phones = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    const key = normalisePhone(guest.phone);
    for (let i = 0; i < phones.length; i++) {
      if (normalisePhone(phones[i][0]) === key) {
        sheet.getRange(i + 2, 1, 1, values.length).setValues([values]);
        return 'updated';
      }
    }
  }
  sheet.appendRow(values);
  return 'added';
}

/* ------------------------------------------------------------
   Requests from the website.
   ------------------------------------------------------------ */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json({ error: 'no data received' });
    }

    const data = JSON.parse(e.postData.contents);

    /* Honeypot: invisible to people, filled in by bots. Answer
       normally so the bot learns nothing, but write nothing. */
    if (data.website) {
      return json({ ok: true, ignored: true });
    }

    /* --- guest checks their number -----------------------------
       Answers yes or no and nothing else. A stranger guessing
       numbers learns only whether one is on the list — no name,
       no code, no admits. */
    if (data.action === 'verify') {
      return json({ found: !!findGuest(data.phone) });
    }

    /* --- the RSVP itself ---------------------------------------
       Verified again here. The browser claiming it already passed
       means nothing, since anyone can post straight to this URL. */
    if (data.form === 'rsvp') {
      const guest = findGuest(data.phone);
      if (!guest) return json({ found: false, saved: false });

      const attending = (String(data.attending) === 'Yes') ? 'Yes' : 'No';

      /* The guest types their own name, so that is what goes on
         the record and the card. Code and admits always come from
         the sheet, never from the browser. */
      const givenName = clean(data.name) || guest.name;
      const outcome = upsertRsvp(guest, givenName, attending);

      return json({
        found:   true,
        saved:   true,
        outcome: outcome,
        code:    guest.code,
        admits:  guest.admits
      });
    }

    if (data.form === 'gift') {
      const sheet = getTab(TABS.gift);
      const now = Utilities.formatDate(new Date(), 'Africa/Lagos', 'yyyy-MM-dd HH:mm:ss');
      sheet.appendRow([now, clean(data.item), clean(data.name), clean(data.message)]);
      return json({ ok: true, saved: true });
    }

    return json({ error: 'unknown request' });

  } catch (err) {
    return json({ error: String(err) });
  }
}

/* ------------------------------------------------------------
   Self-test. Open the /exec URL in a browser: it reports which
   version is live and whether it can actually reach the sheet —
   the two things that otherwise fail silently.
   ------------------------------------------------------------ */
function doGet() {
  const report = {
    version: SCRIPT_VERSION,
    resolvedId: resolveSheetId(SHEET_ID),
    sheetOpened: false,
    tabs: [],
    guestsListed: 0,
    canWrite: false,
    error: null
  };

  try {
    const book = SpreadsheetApp.openById(resolveSheetId(SHEET_ID));
    report.sheetOpened = true;
    report.spreadsheetName = book.getName();
    report.tabs = book.getSheets().map(function (s) {
      return s.getName() + ' (' + Math.max(0, s.getLastRow() - 1) + ' rows)';
    });

    const guests = getTab(GUESTS_TAB);
    report.guestsListed = Math.max(0, guests.getLastRow() - 1);

    /* Prove write access for real, then undo it. */
    const probe = getTab(TABS.rsvp);
    probe.appendRow(['SELF-TEST', 'delete me', '', '', '', '']);
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
   Housekeeping.
   ------------------------------------------------------------ */

/* Finds a tab, creating it with headers the first time. */
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

/* Removes test submissions. Pick cleanupTestRows in the editor
   and click Run. Only deletes rows whose Name starts with one of
   the markers, so a real guest can never match. */
const TEST_MARKERS = ['ZZ TEST', 'ZZ FINAL TEST', 'SELF-TEST', 'ZZ CORS PROBE'];

function cleanupTestRows() {
  const book = SpreadsheetApp.openById(resolveSheetId(SHEET_ID));
  const summary = [];

  [['RSVP', 2], ['Gifts', 3]].forEach(function (pair) {
    const sheet = book.getSheetByName(pair[0]);
    if (!sheet) { summary.push(pair[0] + ': no tab'); return; }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { summary.push(pair[0] + ': already empty'); return; }

    const names = sheet.getRange(2, pair[1], lastRow - 1, 1).getValues();
    let removed = 0;
    for (let i = names.length - 1; i >= 0; i--) {
      const value = String(names[i][0] || '').trim();
      if (TEST_MARKERS.some(function (m) { return value.indexOf(m) === 0; })) {
        sheet.deleteRow(i + 2);
        removed++;
      }
    }
    summary.push(pair[0] + ': removed ' + removed +
                 ', ' + Math.max(0, sheet.getLastRow() - 1) + ' left');
  });

  const msg = summary.join('  |  ');
  Logger.log(msg);
  return msg;
}

/* Trim, cap length, and defuse anything Sheets would run as a
   formula. Without this a guest called "=1+1" becomes a live
   formula in the couple's spreadsheet. */
function clean(value) {
  let text = String(value == null ? '' : value).trim().slice(0, 2000);
  if (/^[=+\-@]/.test(text)) {
    text = "'" + text;
  }
  return text;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
