// ======================
// HMF Sub-Campaign Approvals — Google Apps Script
// Repository: https://github.com/scfeldberg/hmf-sc-approvals
// ======================

// ======================
// CONFIGURATION
// ======================
const CONFIG = {
  QB_BASE_URL:   'https://thehousemajoritypac.quickbase.com/db/bvs4z8ykz',
  QB_USER_TOKEN: 'b7mi6q_h848_0_ceb89dqcctc4vcdree3bzbpt8aqu',
  GITHUB_RAW:    'https://raw.githubusercontent.com/scfeldberg/hmf-sc-approvals/main/Code.gs',

  PASSWORD_DIGITAL:  'Digital',
  PASSWORD_RESEARCH: 'Research',
  PASSWORD_ADMIN:    'Admin',

  // QuickBase field IDs
  FID_DIGITAL_STATUS:         55,
  FID_DIGITAL_KEYWORD:        48,
  FID_DIGITAL_CHECKBOX:      137,
  FID_DIGITAL_CONTENT_KW:    52,
  FID_RESEARCH_STATUS:        50,
  FID_RESEARCH_KEYWORD:       53,
  FID_RESEARCH_CHECKBOX:     138,
  FID_RESEARCH_CONTENT_KW:   54,
  FID_APPROVAL_LOG:          133,
};


// ======================
// Session role — stored in memory for the duration of the script
// execution so internal menu rebuilds don't re-prompt the user.
// Reset to '' on each fresh document open via onOpen().
// ======================
let SESSION_ROLE = '';



// Builds menus dynamically based on document state and user role:
//   - Digital role  → "Digital" menu only
//   - Research role → "Research" menu only
//   - Admin role    → "Digital", "Research", and "Admin" menus
//
// Menu content depends on whether an HMF ID is present:
//   - No HMF ID: show "Add HMF ID" in Digital/Research menus
//   - HMF ID present: show full approval submenus
//   - Content submenu only shown once Script is approved
// ======================
function onOpen() {
  const ui = DocumentApp.getUi();

  const response = ui.prompt(
    'HMF Approval Access',
    'Enter your access password to load your approval menus.\n\n(Cancel to skip — you will be prompted again next time.)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const entered = response.getResponseText().trim();
  let cachedRole = '';

  if      (entered === CONFIG.PASSWORD_ADMIN)    cachedRole = 'admin';
  else if (entered === CONFIG.PASSWORD_DIGITAL)  cachedRole = 'digital';
  else if (entered === CONFIG.PASSWORD_RESEARCH) cachedRole = 'research';
  else {
    ui.alert('Incorrect Password',
      'The password you entered was not recognised.\n\nPlease reopen the document and try again.',
      ui.ButtonSet.OK);
    return;
  }

  SESSION_ROLE = cachedRole;
  buildMenus(cachedRole);
}


// ======================
// Builds menus for the given role. Extracted so that switchRole()
// can rebuild menus mid-session without re-prompting via onOpen().
// ======================
function buildMenus(role) {
  const isAdmin    = (role === 'admin');
  const isDigital  = (role === 'digital');
  const isResearch = (role === 'research');

  const ui  = DocumentApp.getUi();
  const rid = getRIDFromDoc();

  // ---------------------------------------------------------------
  // "Digital" menu — shown to admin and digital users
  // ---------------------------------------------------------------
  if (isAdmin || isDigital) {
    const digiMenu = ui.createMenu('Digital');

    if (!rid) {
      digiMenu.addItem('Add HMF ID', 'promptAddHmfId');
    } else {
      digiMenu.addSubMenu(ui.createMenu('Script')
        .addItem('Script: Pending Digital review',  'digiScript_Pending')
        .addItem('Script: Change request',          'digiScript_ChangeRequest')
        .addItem('Script: Rejected',                'digiScript_Rejected')
        .addItem('Script: Approved w/ comments',    'digiScript_ApprovedWithComments')
        .addItem('Script: Approved',                'digiScript_Approved'));

      const digiStatus = getApprovalStatusText('Digital Approval');
      if (digiStatus.includes('Script: Approved')) {
        digiMenu.addSubMenu(ui.createMenu('Content')
          .addItem('Content: Pending Digital review',  'digiContent_Pending')
          .addItem('Content: Change request',          'digiContent_ChangeRequest')
          .addItem('Content: Rejected',                'digiContent_Rejected')
          .addItem('Content: Approved w/ comments',    'digiContent_ApprovedWithComments')
          .addItem('Content: Approved',                'digiContent_Approved'));
      }
    }

    if (isDigital) {
      digiMenu
        .addSeparator()
        .addItem('Switch Role', 'switchRole');
    }

    digiMenu.addToUi();
  }

  // ---------------------------------------------------------------
  // "Research" menu — shown to admin and research users
  // ---------------------------------------------------------------
  if (isAdmin || isResearch) {
    const resMenu = ui.createMenu('Research');

    if (!rid) {
      resMenu.addItem('Add HMF ID', 'promptAddHmfId');
    } else {
      resMenu.addSubMenu(ui.createMenu('Script')
        .addItem('Script: Pending Research review',  'resScript_Pending')
        .addItem('Script: Change request',           'resScript_ChangeRequest')
        .addItem('Script: Rejected',                 'resScript_Rejected')
        .addItem('Script: Needs Legal review',        'resScript_NeedsLegal')
        .addItem('Script: Approved w/ comments',     'resScript_ApprovedWithComments')
        .addItem('Script: Approved',                 'resScript_Approved'));

      const resStatus = getApprovalStatusText('Research Approval');
      if (resStatus.includes('Script: Approved')) {
        resMenu.addSubMenu(ui.createMenu('Content')
          .addItem('Content: Pending Research review',  'resContent_Pending')
          .addItem('Content: Change request',           'resContent_ChangeRequest')
          .addItem('Content: Rejected',                 'resContent_Rejected')
          .addItem('Content: Needs Legal review',        'resContent_NeedsLegal')
          .addItem('Content: Approved w/ comments',     'resContent_ApprovedWithComments')
          .addItem('Content: Approved',                 'resContent_Approved'));
      }
    }

    if (isResearch) {
      resMenu
        .addSeparator()
        .addItem('Switch Role', 'switchRole');
    }

    resMenu.addToUi();
  }

  // ---------------------------------------------------------------
  // "Admin" menu — shown to admin users only, appears after Research
  // ---------------------------------------------------------------
  if (isAdmin) {
    ui.createMenu('Admin')
      .addItem('Initialize', 'adminInitialize')
      .addItem('Setup',      'promptAddHmfId')
      .addItem('Refresh',    'adminRefresh')
      .addItem('Update',     'adminUpdate')
      .addToUi();
  }
}


// ======================
// Switch Role — prompts for a new password and rebuilds menus
// without requiring the document to be closed and reopened.
// Available in the Digital and Research menus.
// ======================
function switchRole() {
  const ui       = DocumentApp.getUi();
  const response = ui.prompt(
    'Switch Role',
    'Enter your access password for the role you want to switch to:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const entered = response.getResponseText().trim();
  let newRole = '';

  if      (entered === CONFIG.PASSWORD_ADMIN)    newRole = 'admin';
  else if (entered === CONFIG.PASSWORD_DIGITAL)  newRole = 'digital';
  else if (entered === CONFIG.PASSWORD_RESEARCH) newRole = 'research';
  else {
    ui.alert('Incorrect Password',
      'The password you entered was not recognised. Your current menus remain active.',
      ui.ButtonSet.OK);
    return;
  }

  SESSION_ROLE = newRole;
  buildMenus(newRole);
}


// ======================
// ADMIN — Initialize
// Clears the HMF ID from the doc, resets both approval status cells
// to their pending defaults, then rebuilds the menus so the user
// can begin the Setup process.
// Does NOT touch QuickBase — Initialize is a doc-only reset.
// ======================
function adminInitialize() {
  const ui = DocumentApp.getUi();

  const confirm = ui.alert(
    '⚠️ Initialize — Are you sure?',
    'This will:\n\n' +
    '  • Clear the HMF ID from this document\n' +
    '  • Reset both approval statuses to Pending\n\n' +
    'This action cannot be undone. Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  const body = DocumentApp.getActiveDocument().getBody();

  // --- 1. Clear the HMF ID cell ---
  let hmfCell = null;
  outer:
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
    const table = el.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
        const cell = table.getRow(r).getCell(c);
        if (/HMF\s+ID\s*:/i.test(cell.getText())) {
          hmfCell = cell;
          break outer;
        }
      }
    }
  }

  if (!hmfCell) {
    ui.alert('❌ Table Not Found',
      'Could not find the "HMF ID:" cell in the approval table.',
      ui.ButtonSet.OK);
    return;
  }

  // Detect font from the HMF ID cell before clearing it
  let fontFamily = 'Arial';
  let fontSize   = 10;
  try {
    const ft = hmfCell.getChild(0).asParagraph().editAsText();
    if (ft.getFontFamily(0)) fontFamily = ft.getFontFamily(0);
    if (ft.getFontSize(0))   fontSize   = ft.getFontSize(0);
  } catch (e) { /* fall back to defaults */ }

  // Write back the label only, no ID
  writeCellText(hmfCell, 'HMF ID:', fontFamily, fontSize);

  // --- 2. Reset approval status cells ---
  const DIGITAL_DEFAULT  = 'Script: Pending Digital review';
  const RESEARCH_DEFAULT = 'Script: Pending Research review';

  const digiCell = getStatusCell('Digital Approval');
  const resCell  = getStatusCell('Research Approval');

  if (digiCell) writeCellText(digiCell, DIGITAL_DEFAULT,  fontFamily, fontSize);
  if (resCell)  writeCellText(resCell,  RESEARCH_DEFAULT, fontFamily, fontSize);

  // --- 3. Clean up any stray legacy status lines outside tables ---
  const oldHeader = body.findText('Digital Review Status');
  if (oldHeader) {
    const el = oldHeader.getElement().getParent();
    if (el && el.getType() === DocumentApp.ElementType.PARAGRAPH) el.removeFromParent();
  }
  for (let i = body.getNumChildren() - 1; i >= 0; i--) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const txt = el.asParagraph().getText().trim();
    if ([...DIGITAL_STATUSES, ...RESEARCH_STATUSES].some(s => s === txt)) {
      el.removeFromParent();
    }
  }

  ui.alert('✅ Initialized',
    'HMF ID has been cleared and approval statuses have been reset.\n\nUse Admin > Setup to assign a new HMF ID.',
    ui.ButtonSet.OK);

  // --- 4. Rebuild menus — will now show "Add HMF ID" since ID is cleared ---
  SESSION_ROLE = '';
  buildMenus(SESSION_ROLE);
}


// ======================
// ADMIN — Refresh
// Reads the current Digital and Research approval statuses directly
// from QuickBase (fid 55 and fid 50) and writes them into the doc
// status cells, overwriting whatever is currently displayed.
// ======================
function adminRefresh() {
  const ui  = DocumentApp.getUi();
  const rid = getRIDFromDoc();

  if (!rid) {
    ui.alert('⚠️ No HMF ID',
      'No HMF ID is set in this document.\n\nPlease run Admin > Setup first.',
      ui.ButtonSet.OK);
    return;
  }

  try {
    // QB XML API — API_GetRecordInfo fetches all field values for a record
    const apiUrl = `${CONFIG.QB_BASE_URL}?a=API_GetRecordInfo&rid=${rid}&usertoken=${CONFIG.QB_USER_TOKEN}`;
    const response = UrlFetchApp.fetch(apiUrl, { method: 'get', muteHttpExceptions: true });
    const xml      = XmlService.parse(response.getContentText());
    const root     = xml.getRootElement();

    const errcode = root.getChildText('errcode') || '999';
    if (errcode !== '0') {
      ui.alert('❌ QuickBase Error',
        `API returned error code ${errcode}.\n\nPlease check your connection and try again.`,
        ui.ButtonSet.OK);
      return;
    }

    // Extract fid 55 (Digital status) and fid 50 (Research status)
    let digitalStatus  = '';
    let researchStatus = '';

    const fields = root.getChildren('field');
    for (const field of fields) {
      const fid = field.getChildText('fid');
      if (fid === String(CONFIG.FID_DIGITAL_STATUS))  digitalStatus  = field.getChildText('value') || '';
      if (fid === String(CONFIG.FID_RESEARCH_STATUS)) researchStatus = field.getChildText('value') || '';
    }

    if (!digitalStatus && !researchStatus) {
      ui.alert('⚠️ No Data',
        'QuickBase returned no approval status values for this record.\n\nPlease verify the HMF ID is correct.',
        ui.ButtonSet.OK);
      return;
    }

    // Write values to doc cells
    const fontFamily = 'Arial';
    const fontSize   = 10;

    if (digitalStatus) {
      const digiCell = getStatusCell('Digital Approval');
      if (digiCell) writeCellText(digiCell, digitalStatus, fontFamily, fontSize);
    }

    if (researchStatus) {
      const resCell = getStatusCell('Research Approval');
      if (resCell) writeCellText(resCell, researchStatus, fontFamily, fontSize);
    }

    ui.alert('✅ Refreshed',
      `Approval statuses updated from QuickBase for Record ${rid}:\n\n` +
      `Digital:  ${digitalStatus  || '(no value)'}\n` +
      `Research: ${researchStatus || '(no value)'}`,
      ui.ButtonSet.OK);

    // Rebuild menus — Content submenus may now need to appear/disappear
    buildMenus(SESSION_ROLE);

  } catch (error) {
    ui.alert('Script Error', error.toString(), ui.ButtonSet.OK);
  }
}


// ======================
// ADMIN — Update
// Fetches the latest Code.gs from the public GitHub repository and
// replaces the content of this script file with the downloaded code.
// Uses the Apps Script API via UrlFetchApp — the script rewrites itself.
//
// Note: After Update runs, the user must reload the document (close
// and reopen) for the new code to take effect in the menu system.
// ======================
function adminUpdate() {
  const ui = DocumentApp.getUi();

  const confirm = ui.alert(
    '🔄 Update Script',
    'This will download the latest version of this script from GitHub and replace the current code.\n\n' +
    'After the update completes, please close and reopen this document for changes to take effect.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  try {
    // --- 1. Fetch latest code from GitHub ---
    const ghResponse = UrlFetchApp.fetch(CONFIG.GITHUB_RAW, {
      method:             'get',
      muteHttpExceptions: true,
    });

    if (ghResponse.getResponseCode() !== 200) {
      ui.alert('❌ Download Failed',
        `Could not reach GitHub (HTTP ${ghResponse.getResponseCode()}).\n\nPlease check your connection and try again.`,
        ui.ButtonSet.OK);
      return;
    }

    const newCode = ghResponse.getContentText();

    if (!newCode || newCode.trim().length < 100) {
      ui.alert('❌ Invalid Response',
        'The file downloaded from GitHub appears to be empty or invalid. Update cancelled.',
        ui.ButtonSet.OK);
      return;
    }

    // --- 2. Overwrite this script file via the Apps Script API ---
    // The script ID is available via ScriptApp.getScriptId()
    const scriptId  = ScriptApp.getScriptId();
    const oauthToken = ScriptApp.getOAuthToken();

    // First GET the project to find the file name for Code.gs
    const projectUrl = `https://script.googleapis.com/v1/projects/${scriptId}/content`;
    const getResp    = UrlFetchApp.fetch(projectUrl, {
      method:  'get',
      headers: { Authorization: `Bearer ${oauthToken}` },
      muteHttpExceptions: true,
    });

    if (getResp.getResponseCode() !== 200) {
      ui.alert('❌ Script API Error',
        'Could not read the current script project from the Apps Script API.\n\nEnsure this script has the "script.projects" OAuth scope.',
        ui.ButtonSet.OK);
      return;
    }

    const project = JSON.parse(getResp.getContentText());
    const files   = project.files || [];

    // Find Code.gs (or the first .gs file) and update its source
    let updated = false;
    for (const file of files) {
      if (file.type === 'SERVER_JS' && file.name === 'Code') {
        file.source = newCode;
        updated = true;
        break;
      }
    }

    if (!updated) {
      // Fallback: update the first SERVER_JS file found
      for (const file of files) {
        if (file.type === 'SERVER_JS') {
          file.source = newCode;
          updated = true;
          break;
        }
      }
    }

    if (!updated) {
      ui.alert('❌ File Not Found',
        'Could not find a Google Apps Script file to update in this project.',
        ui.ButtonSet.OK);
      return;
    }

    // PUT the updated project back
    const putResp = UrlFetchApp.fetch(projectUrl, {
      method:      'put',
      headers:     { Authorization: `Bearer ${oauthToken}`, 'Content-Type': 'application/json' },
      payload:     JSON.stringify({ files }),
      muteHttpExceptions: true,
    });

    if (putResp.getResponseCode() !== 200) {
      ui.alert('❌ Update Failed',
        'The script was downloaded but could not be saved.\n\nApps Script API returned: ' + putResp.getResponseCode(),
        ui.ButtonSet.OK);
      return;
    }

    ui.alert('✅ Update Complete',
      'The script has been updated to the latest version from GitHub.\n\n' +
      'Please close and reopen this document for the changes to take effect.',
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('Script Error', error.toString(), ui.ButtonSet.OK);
  }
}


// ======================
// Reads the HMF ID number from the approval table without prompting.
// Returns the ID string if found, or '' if missing/blank.
// ======================
function getRIDFromDoc() {
  const body = DocumentApp.getActiveDocument().getBody();
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
    const table = el.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
        const cellText = table.getRow(r).getCell(c).getText();
        const match = cellText.match(/HMF\s+ID\s*:\s*(\d+)/i);
        if (match) return match[1].trim();
      }
    }
  }
  return '';
}


// ======================
// Returns the current text of a status cell identified by its label.
// Used by onOpen() to decide whether to show Content submenus.
// ======================
function getApprovalStatusText(labelText) {
  const body = DocumentApp.getActiveDocument().getBody();
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
    const table = el.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        if (row.getCell(c).getText().trim() === labelText) {
          if (c + 1 < row.getNumCells()) return row.getCell(c + 1).getText().trim();
        }
      }
    }
  }
  return '';
}


// ======================
// Read the QuickBase Record ID from the "HMF ID:" cell.
// Falls back to a manual prompt if not found.
// ======================
function getOrPromptForRID() {
  const rid = getRIDFromDoc();
  if (rid) return rid;

  const response = DocumentApp.getUi().prompt(
    'QuickBase Record ID Required',
    '"HMF ID:" not found or has no number in the approval table.\n\nPlease enter the Record ID manually:',
    DocumentApp.getUi().ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== DocumentApp.getUi().Button.OK) return null;

  const entered = response.getResponseText().trim();
  if (!entered || isNaN(Number(entered))) {
    DocumentApp.getUi().alert('Invalid Input',
      'Please enter a valid numerical Record ID.',
      DocumentApp.getUi().ButtonSet.OK);
    return null;
  }

  return entered;
}


// ======================
// Opens a modal dialog for the user to enter the Sub-Campaign Record ID.
// Validates numerical input then writes it to the "HMF ID:" cell and
// rebuilds menus via onOpen().
// Accessible via Admin > Setup.
// ======================
function promptAddHmfId() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body {
        font-family: Arial, sans-serif;
        font-size: 13px;
        padding: 16px;
        margin: 0;
      }
      label {
        display: block;
        margin-bottom: 6px;
        font-weight: normal;
      }
      input[type="text"] {
        width: 100%;
        padding: 6px 8px;
        font-size: 13px;
        border: 1px solid #ccc;
        border-radius: 3px;
        box-sizing: border-box;
        margin-bottom: 10px;
      }
      input[type="text"].error { border-color: #c0392b; }
      .error-msg {
        color: #c0392b;
        font-size: 12px;
        margin-bottom: 10px;
        display: none;
      }
      .buttons {
        display: flex;
        justify-content: flex-end;
        gap: 8px;
        margin-top: 4px;
      }
      button {
        padding: 6px 16px;
        font-size: 13px;
        border: none;
        border-radius: 3px;
        cursor: pointer;
      }
      .btn-cancel { background: #e0e0e0; color: #333; }
      .btn-save   { background: #1a73e8; color: white; }
      .btn-save:hover   { background: #1558b0; }
      .btn-cancel:hover { background: #c8c8c8; }
    </style>

    <label for="rid">Sub-Campaign Record ID#:</label>
    <input type="text" id="rid" placeholder="Enter numerical ID" autofocus />
    <div class="error-msg" id="errMsg">Please enter a valid numerical ID.</div>

    <div class="buttons">
      <button class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
      <button class="btn-save"   onclick="submitId()">Save</button>
    </div>

    <script>
      document.getElementById('rid').addEventListener('keydown', function(e) {
        if (e.key === 'Enter') submitId();
      });

      function submitId() {
        const val    = document.getElementById('rid').value.trim();
        const input  = document.getElementById('rid');
        const errMsg = document.getElementById('errMsg');

        if (!val || !/^\\d+$/.test(val)) {
          input.classList.add('error');
          errMsg.style.display = 'block';
          input.focus();
          return;
        }

        input.classList.remove('error');
        errMsg.style.display = 'none';

        google.script.run
          .withSuccessHandler(function() { google.script.host.close(); })
          .withFailureHandler(function(err) {
            errMsg.textContent = 'Error saving: ' + err.message;
            errMsg.style.display = 'block';
          })
          .writeHmfIdToDoc(val);
      }
    </script>
  `)
  .setWidth(320)
  .setHeight(150)
  .setTitle('Setup — Add HMF ID');

  DocumentApp.getUi().showModalDialog(html, 'Setup — Add HMF ID');
}


// ======================
// Writes "HMF ID: <rid>" into the HMF ID cell, preserving font.
// Rebuilds menus after saving.
// ======================
function writeHmfIdToDoc(rid) {
  const body = DocumentApp.getActiveDocument().getBody();

  let hmfCell = null;
  outer:
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
    const table = el.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
        const cell = table.getRow(r).getCell(c);
        if (/HMF\s+ID\s*:/i.test(cell.getText())) {
          hmfCell = cell;
          break outer;
        }
      }
    }
  }

  if (!hmfCell) throw new Error('"HMF ID:" cell not found in the approval table.');

  let fontFamily = 'Arial';
  let fontSize   = 10;
  try {
    const ft = hmfCell.getChild(0).asParagraph().editAsText();
    if (ft.getFontFamily(0)) fontFamily = ft.getFontFamily(0);
    if (ft.getFontSize(0))   fontSize   = ft.getFontSize(0);
  } catch (e) { /* fall back to defaults */ }

  writeCellText(hmfCell, `HMF ID: ${rid}`, fontFamily, fontSize);

  buildMenus(SESSION_ROLE);
  adminRefresh();
}


// ======================
// STATUS CONFIG
// ======================
const DIGITAL_STATUSES = [
  'Script: Pending Digital review',
  'Script: Change request',
  'Script: Rejected',
  'Script: Approved w/ comments',
  'Script: Approved',
  'Content: Pending Digital review',
  'Content: Change request',
  'Content: Rejected',
  'Content: Approved w/ comments',
  'Content: Approved',
];

const RESEARCH_STATUSES = [
  'Script: Pending Research review',
  'Script: Change request',
  'Script: Rejected',
  'Script: Needs Legal review',
  'Script: Approved w/ comments',
  'Script: Approved',
  'Content: Pending Research review',
  'Content: Change request',
  'Content: Rejected',
  'Content: Needs Legal review',
  'Content: Approved w/ comments',
  'Content: Approved',
];


// ======================
// Returns the status cell for a given label.
// col 0 = "Digital Approval"  → status in col 1
// col 2 = "Research Approval" → status in col 3
// ======================
function getStatusCell(labelText) {
  const body = DocumentApp.getActiveDocument().getBody();
  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.TABLE) continue;
    const table = el.asTable();
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        if (row.getCell(c).getText().trim() === labelText) {
          if (c + 1 < row.getNumCells()) return row.getCell(c + 1);
        }
      }
    }
  }
  return null;
}


// ======================
// Writes text into a table cell's first paragraph in place.
// Always non-bold.
// ======================
function writeCellText(cell, text, fontFamily, fontSize) {
  while (cell.getNumChildren() > 1) {
    cell.getChild(cell.getNumChildren() - 1).removeFromParent();
  }
  const para   = cell.getChild(0).asParagraph();
  para.clear();
  para.appendText(text);
  const textEl = para.editAsText();
  textEl.setFontFamily(fontFamily || 'Arial');
  textEl.setFontSize(fontSize || 10);
  textEl.setBold(false);
  return textEl;
}


// ======================
// Generic status updater used by both Digital and Research menus.
// ======================
function updateApprovalStatus(labelText, statusValue, fidStatus, fidKeyword, fidCheckbox, pendingContent) {
  const rid = getOrPromptForRID();
  if (!rid) return;

  const colonIdx   = statusValue.indexOf(': ');
  const rawKeyword = statusValue.substring(colonIdx + 2);

  const CONTENT_KEYWORD_MAP = {
    'Pending Digital review':  'Pending review',
    'Pending Research review': 'Pending review',
    'Change request':          'Change Request',
    'Rejected':                'Rejected',
    'Needs Legal review':      'Needs Legal review',
    'Approved w/ comments':    'Approved w/ comments',
    'Approved':                'Approved',
  };

  const isContent     = statusValue.startsWith('Content:');
  const statusKeyword = isContent
    ? (CONTENT_KEYWORD_MAP[rawKeyword] || rawKeyword)
    : rawKeyword;

  try {
    const approvalPrefix      = labelText.replace(' Approval', '');
    const logValue            = `${approvalPrefix} ${statusValue}`;
    const effectiveKeywordFid = isContent
      ? (labelText === 'Digital Approval' ? CONFIG.FID_DIGITAL_CONTENT_KW : CONFIG.FID_RESEARCH_CONTENT_KW)
      : fidKeyword;

    const encodedStatus  = encodeURIComponent(statusValue);
    const encodedKeyword = encodeURIComponent(statusKeyword);
    const encodedLog     = encodeURIComponent(logValue);

    const scriptApproved     = statusValue.startsWith('Script: Approved');
    const contentKeywordFid  = labelText === 'Digital Approval'
      ? CONFIG.FID_DIGITAL_CONTENT_KW
      : CONFIG.FID_RESEARCH_CONTENT_KW;
    const pendingReviewParam = scriptApproved
      ? `&_fid_${contentKeywordFid}=${encodeURIComponent('Pending review')}`
      : '';

    const apiUrl = `${CONFIG.QB_BASE_URL}?a=API_EditRecord&rid=${rid}` +
      `&_fid_${fidStatus}=${encodedStatus}` +
      `&_fid_${effectiveKeywordFid}=${encodedKeyword}` +
      `&_fid_${fidCheckbox}=1` +
      `&_fid_${CONFIG.FID_APPROVAL_LOG}=${encodedLog}` +
      `${pendingReviewParam}` +
      `&usertoken=${CONFIG.QB_USER_TOKEN}`;

    const response = UrlFetchApp.fetch(apiUrl, { method: 'get', muteHttpExceptions: true });

    let success = false;
    try {
      const xml     = XmlService.parse(response.getContentText());
      const errcode = xml.getRootElement().getChildText('errcode') || '999';
      success = (errcode === '0');
    } catch (e) {
      success = (response.getResponseCode() >= 200 && response.getResponseCode() < 300);
    }

    if (!success) {
      DocumentApp.getUi().alert('❌ Update Failed',
        'The QuickBase API call did not succeed. Please try again.',
        DocumentApp.getUi().ButtonSet.OK);
      return;
    }

    const statusCell = getStatusCell(labelText);
    if (!statusCell) {
      DocumentApp.getUi().alert('⚠️ Cell Not Found',
        `Could not find the ${labelText} status cell.\n\nPlease run Admin > Initialize, then Admin > Setup.`,
        DocumentApp.getUi().ButtonSet.OK);
      return;
    }

    let cellValue;
    if (statusValue.startsWith('Script: Approved')) {
      cellValue = `${statusValue} / ${pendingContent}`;
    } else if (isContent) {
      const currentText  = statusCell.getText().trim();
      const slashIdx     = currentText.indexOf(' / ');
      const scriptPrefix = slashIdx !== -1 ? currentText.substring(0, slashIdx) : '';
      cellValue = scriptPrefix ? `${scriptPrefix} / ${statusValue}` : statusValue;
    } else {
      cellValue = statusValue;
    }

    writeCellText(statusCell, cellValue, 'Arial', 10);

    DocumentApp.getUi().alert('✅ Success',
      `"${statusValue}" has been set for Record ${rid}.`,
      DocumentApp.getUi().ButtonSet.OK);

    buildMenus(SESSION_ROLE);

  } catch (error) {
    DocumentApp.getUi().alert('Script Error', error.toString(), DocumentApp.getUi().ButtonSet.OK);
  }
}


// ======================
// Digital — Script handlers
// ======================
function digiScript_Pending()              { updateApprovalStatus('Digital Approval', 'Script: Pending Digital review',  CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }
function digiScript_ChangeRequest()        { updateApprovalStatus('Digital Approval', 'Script: Change request',          CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }
function digiScript_Rejected()             { updateApprovalStatus('Digital Approval', 'Script: Rejected',                CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }
function digiScript_ApprovedWithComments() { updateApprovalStatus('Digital Approval', 'Script: Approved w/ comments',   CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }
function digiScript_Approved()             { updateApprovalStatus('Digital Approval', 'Script: Approved',                CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }

// ======================
// Digital — Content handlers
// ======================
function digiContent_Pending()              { updateApprovalStatus('Digital Approval', 'Content: Pending Digital review',  CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }
function digiContent_ChangeRequest()        { updateApprovalStatus('Digital Approval', 'Content: Change request',          CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }
function digiContent_Rejected()             { updateApprovalStatus('Digital Approval', 'Content: Rejected',                CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }
function digiContent_ApprovedWithComments() { updateApprovalStatus('Digital Approval', 'Content: Approved w/ comments',   CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }
function digiContent_Approved()             { updateApprovalStatus('Digital Approval', 'Content: Approved',                CONFIG.FID_DIGITAL_STATUS, CONFIG.FID_DIGITAL_KEYWORD, CONFIG.FID_DIGITAL_CHECKBOX, 'Content: Pending Digital review'); }

// ======================
// Research — Script handlers
// ======================
function resScript_Pending()              { updateApprovalStatus('Research Approval', 'Script: Pending Research review',  CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resScript_ChangeRequest()        { updateApprovalStatus('Research Approval', 'Script: Change request',           CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resScript_Rejected()             { updateApprovalStatus('Research Approval', 'Script: Rejected',                 CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resScript_NeedsLegal()           { updateApprovalStatus('Research Approval', 'Script: Needs Legal review',       CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resScript_ApprovedWithComments() { updateApprovalStatus('Research Approval', 'Script: Approved w/ comments',    CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resScript_Approved()             { updateApprovalStatus('Research Approval', 'Script: Approved',                 CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }

// ======================
// Research — Content handlers
// ======================
function resContent_Pending()              { updateApprovalStatus('Research Approval', 'Content: Pending Research review',  CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resContent_ChangeRequest()        { updateApprovalStatus('Research Approval', 'Content: Change request',           CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resContent_Rejected()             { updateApprovalStatus('Research Approval', 'Content: Rejected',                 CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resContent_NeedsLegal()           { updateApprovalStatus('Research Approval', 'Content: Needs Legal review',       CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resContent_ApprovedWithComments() { updateApprovalStatus('Research Approval', 'Content: Approved w/ comments',    CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }
function resContent_Approved()             { updateApprovalStatus('Research Approval', 'Content: Approved',                 CONFIG.FID_RESEARCH_STATUS, CONFIG.FID_RESEARCH_KEYWORD, CONFIG.FID_RESEARCH_CHECKBOX, 'Content: Pending Research review'); }