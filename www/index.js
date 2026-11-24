// OAuth client ID from the Google Cloud console. It's public, and Google only
// accepts it from its authorized origins (localhost:8000, andrewmacheret.com)
const CLIENT_ID = '248237455343-92pm91eir0o4mh86obkq2v60u8a30dkp.apps.googleusercontent.com';
// The file picker needs the Google Cloud project number and an API key. The
// key is public too, and only works from andrewmacheret.com for the picker
const APP_ID = '248237455343';
const PICKER_API_KEY = 'AIzaSyAxkzeJQJKcw5kZ_P-NXqG15scFqMIrf3A';
// drive.file only reaches spreadsheets the user picks or Miler Lite creates
const SCOPES = 'https://www.googleapis.com/auth/drive.file';
const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';

const formName = 'mileage-form';
const form = document.forms[formName];

const $id = document.getElementById.bind(document);

const year = new Date().getFullYear();
const mileageSheet = `${year} Mileage`;
const settingsSheet = `${year} Settings`;
const startingRow = 4;

// The spreadsheets this browser has used, newest first
const sheetsKey = 'miler-lite-sheets';
const maxRecentSheets = 10;
let sheetList = loadSheetList();

let settings = {
  lastRowNum: 0,
  presets: [],
  maxMileage: 0
}
let submittingMileage = false;
let tokenClient;
// Goalie and Miler Lite share one Google sign-in, so they share the saved token
const tokenKey = 'daily-tracking-token';
let accessToken = loadSavedToken();
let signedOut = false;

// Older versions kept settings under a key that Goalie used too
window.localStorage.removeItem('settings');

function removeTime(date) {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
}
$id('date').valueAsDate = removeTime(new Date());

setMessage('info', 'Loading Google sign-in...');
$id('remove-button').onclick = handleRemoveClick;
renderSheetControls();
loadSheetsCached();

// Escape a string for HTML interpolation.
function escapeHTML(string) {
  return ('' + string).replace(/[&<>"'\/]/g, match => {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#x27;',
      '/': '&#x2F;'
    }[match];
  });
}

function loadSheetList() {
  try {
    const list = JSON.parse(window.localStorage.getItem(sheetsKey));
    if (list && Array.isArray(list.recent)) return list;
  } catch (error) {
    console.error(error);
  }
  return {current: null, recent: []};
}

function currentSheetId() {
  return sheetList.current;
}

function spreadsheetLink() {
  return `https://docs.google.com/spreadsheets/d/${currentSheetId()}/edit`;
}

function rememberSheet(id, name) {
  sheetList.recent = [{id, name}, ...sheetList.recent.filter(sheet => sheet.id !== id)].slice(0, maxRecentSheets);
  sheetList.current = id;
  window.localStorage.setItem(sheetsKey, JSON.stringify(sheetList));
  renderSheetControls();
}

function renderSheetControls() {
  const select = $id('sheet-select');
  select.innerHTML = sheetList.recent.map(sheet =>
    `<option value="${escapeHTML(sheet.id)}">${escapeHTML(sheet.name)}</option>`).join('');
  select.value = currentSheetId() || '';
  select.style.display = sheetList.recent.length ? '' : 'none';
  $id('remove-button').style.display = currentSheetId() ? '' : 'none';

  const link = $id('spreadsheet-link');
  link.style.display = currentSheetId() ? 'inline-block' : 'none';
  link.href = spreadsheetLink();
}

function handleSheetSelect(event) {
  const sheet = sheetList.recent.find(sheet => sheet.id === event.target.value);
  if (!sheet) return;
  rememberSheet(sheet.id, sheet.name);
  form.style.display = 'none';
  loadSheetsCached();
  if (accessToken) loadSheets();
}

// Takes the spreadsheet off this browser's list. It stays in Google Drive
function handleRemoveClick() {
  const id = currentSheetId();
  if (!id) return;
  window.localStorage.removeItem(settingsKey());
  sheetList.recent = sheetList.recent.filter(sheet => sheet.id !== id);
  sheetList.current = sheetList.recent.length ? sheetList.recent[0].id : null;
  window.localStorage.setItem(sheetsKey, JSON.stringify(sheetList));
  renderSheetControls();
  form.style.display = 'none';
  loadSheetsCached();
  updateSigninStatus(!!accessToken);
}


/**
 *  Called when the Google Identity Services script loads.
 */
function handleClientLoad() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: handleTokenResponse,
    error_callback: error => setMessage('danger', 'Sign-in failed: ' + escapeHTML(error.message || error.type)),
  });

  $id('authorize-button').onclick = handleAuthClick;
  $id('signout-button').onclick = handleSignoutClick;
  $id('sheet-select').onchange = handleSheetSelect;
  $id('pick-button').onclick = handlePickClick;
  $id('create-button').onclick = handleCreateClick;
  updateSigninStatus(!!accessToken);
}

// Access tokens last an hour. Keeping it until then means a reload doesn't
// ask the user to sign in again
function loadSavedToken() {
  try {
    const saved = JSON.parse(window.localStorage.getItem(tokenKey));
    if (saved && saved.expiresAt > Date.now()) return saved.token;
  } catch (error) {
    console.error(error);
  }
  return null;
}

function saveToken(token, expiresInSeconds) {
  // Stop a minute early so a request doesn't start with an almost-expired token
  const expiresAt = Date.now() + (expiresInSeconds - 60) * 1000;
  window.localStorage.setItem(tokenKey, JSON.stringify({token, expiresAt}));
}

function forgetToken() {
  accessToken = null;
  window.localStorage.removeItem(tokenKey);
}

function handleClientLoadError() {
  setMessage('danger', 'Could not load Google sign-in. Check your connection and reload.');
}

function handleTokenResponse(response) {
  if (response.error) {
    setMessage('danger', 'Sign-in failed: ' + escapeHTML(response.error_description || response.error));
    return;
  }
  if (!google.accounts.oauth2.hasGrantedAllScopes(response, SCOPES)) {
    setMessage('warning', 'Miler Lite needs access to its spreadsheets to work.');
    return;
  }
  accessToken = response.access_token;
  saveToken(accessToken, Number(response.expires_in) || 3600);
  updateSigninStatus(true);
}

/**
 *  Called when the signed in status changes, to update the UI
 *  appropriately. After a sign-in, the API is called.
 */
function updateSigninStatus(isSignedIn) {
  $id('authorize-button').style.display = isSignedIn ? 'none' : '';
  $id('signout-button').style.display = isSignedIn ? '' : 'none';
  $id('sheet-buttons').style.display = isSignedIn ? '' : 'none';
  if (!isSignedIn) {
    setMessage('warning', 'Click Authorize to sign in with Google.');
  } else if (currentSheetId()) {
    loadSheets();
  } else {
    setMessage('info', 'Click <b>Open</b> to choose your Miler Lite spreadsheet, or <b>New</b> to create one.');
  }
}

/**
 *  Sign in the user upon button click.
 */
function handleAuthClick(event) {
  // An empty prompt skips the consent screen once the user has agreed to it
  tokenClient.requestAccessToken({prompt: signedOut ? 'select_account' : ''});
}

/**
 *  Sign out the user upon button click.
 */
function handleSignoutClick(event) {
  // Forget the token without revoking it. Revoking would also drop access to
  // the spreadsheets the user picked, and they'd have to pick them again
  forgetToken();
  signedOut = true;
  updateSigninStatus(false);
}

function handlePickClick() {
  if (typeof gapi === 'undefined') {
    setMessage('warning', 'The file picker is still loading. Try again in a moment.');
    return;
  }
  gapi.load('picker', () => {
    const view = new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS)
      .setMode(google.picker.DocsViewMode.LIST);
    new google.picker.PickerBuilder()
      .setTitle('Choose your Miler Lite spreadsheet')
      .addView(view)
      .setOAuthToken(accessToken)
      .setDeveloperKey(PICKER_API_KEY)
      .setAppId(APP_ID)
      .setCallback(handlePicked)
      .build()
      .setVisible(true);
  });
}

function handlePicked(data) {
  if (data[google.picker.Response.ACTION] !== google.picker.Action.PICKED) return;
  const doc = data[google.picker.Response.DOCUMENTS][0];
  rememberSheet(doc[google.picker.Document.ID], doc[google.picker.Document.NAME]);
  form.style.display = 'none';
  loadSheetsCached();
  loadSheets();
}

function handleCreateClick() {
  setMessage('info', 'Creating a new Miler Lite spreadsheet...');
  const body = {
    properties: {title: 'Miler Lite mileage'},
    sheets: yearTabs().map(tab => ({properties: {title: tab.title}, data: [{rowData: tab.rows}]})),
  };
  sheetsFetch('', {method: 'POST', body: JSON.stringify(body)})
  .then(result => {
    rememberSheet(result.spreadsheetId, result.properties.title);
    loadSheets();
  }).catch(error => {
    console.error(error);
    setMessage('danger', 'Error: ' + error);
  });
}

// Adds this year's tabs to the current spreadsheet
function addYearTabs() {
  setMessage('info', `Adding the ${year} tabs...`);
  const requests = [];
  yearTabs().forEach((tab, index) => {
    // Pick sheet IDs up front so the cells can go in the same request
    const sheetId = year * 10 + index;
    requests.push({addSheet: {properties: {sheetId, title: tab.title, index}}});
    requests.push({updateCells: {start: {sheetId, rowIndex: 0, columnIndex: 0}, rows: tab.rows, fields: 'userEnteredValue'}});
  });
  sheetsFetch(`/${currentSheetId()}:batchUpdate`, {method: 'POST', body: JSON.stringify({requests})})
  .then(() => loadSheets())
  .catch(error => {
    console.error(error);
    setMessage('danger', 'Error: ' + error);
  });
}

function starterCell(value) {
  if (value === null) return {};
  return {userEnteredValue: typeof value === 'number' ? {numberValue: value}
    : value.startsWith('=') ? {formulaValue: value} : {stringValue: value}};
}

// A year's mileage and settings tabs in the layout Miler Lite reads
function yearTabs() {
  const tabs = [
    {title: mileageSheet, rows: [
      [`Miler Lite adds a row here each time you submit.`],
      [],
      ['Date', 'Business', 'Purpose', 'Destination', 'Start', 'End', 'Miles', 'Amount'],
    ]},
    {title: settingsSheet, rows: [
      ['List businesses in column A and saved trips in columns C to E, from row 3 down. Set your rate per mile in G3.'],
      ['Business', null, 'Purpose', 'Destination', 'Miles', null, 'Rate per mile', null, 'Highest odometer', 'Next row'],
      ['Work', null, 'Office', 'Main office', 10, null, 0.7, null,
        `=MAX('${mileageSheet}'!F${startingRow}:F)`, `=COUNTA('${mileageSheet}'!A${startingRow}:A)+${startingRow}`],
      ['Side business', null, 'Client visit', 'Client office', 25],
    ]},
  ];
  return tabs.map(tab => ({title: tab.title, rows: tab.rows.map(row => ({values: row.map(starterCell)}))}));
}

/**
 * Append a pre element to the body containing the given message
 * as its text node. Used to display the results of the API call.
 *
 * @param {string} message Text to be placed in pre element.
 */
function setMessage(level, message) {
  const messageElement = $id('message');
  messageElement.innerHTML = message;
  messageElement.className = 'alert alert-' + level;

  const messageElement2 = $id('message-below');
  messageElement2.innerHTML = message;
  messageElement2.className = 'alert alert-' + level;

  console.log(level, message);
}

function recalcMileage() {
  const mileage = (parseInt($id('end').value, 10) || 0) - (parseInt($id('start').value, 10) || 0);
  $id('mileage').value = mileage || '';
}

function startChanged() {
  $id('end').value = ((parseInt($id('start').value, 10) || 0) + (parseInt($id('mileage').value, 10) || 0)) || '';
  recalcMileage();
}

function endChanged() {
  recalcMileage();
}

// Calls the Sheets API and rejects with the error message
async function sheetsFetch(path, options = {}) {
  const headers = {...options.headers, Authorization: `Bearer ${accessToken}`};
  if (options.body) headers['Content-Type'] = 'application/json';
  const response = await fetch(`${SHEETS_API}${path}`, {...options, headers});
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    // Access tokens expire after an hour, so ask to sign in again
    if (response.status === 401) {
      forgetToken();
      updateSigninStatus(false);
      throw 'Sign-in expired. Click Authorize to sign in again.';
    }
    throw escapeHTML(body.error ? body.error.message : response.statusText);
  }
  return body;
}

function getSpreadsheet(range) {
  const params = new URLSearchParams({includeGridData: 'true',
    fields: 'properties.title,sheets.data.rowData.values.formattedValue'});
  // Tab names have spaces, which URLSearchParams would turn into +
  return sheetsFetch(`/${currentSheetId()}?ranges=${encodeURIComponent(range)}&${params}`);
}

function appendSpreadsheetRow(range, row) {
  const params = new URLSearchParams({
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    includeValuesInResponse: 'false'
  });
  return sheetsFetch(`/${currentSheetId()}/values/${encodeURIComponent(range)}:append?${params}`, {
    method: 'POST',
    body: JSON.stringify({values: [row]}),
  }).then(result => result.updates);
}

function settingsKey() {
  return `miler-lite-settings-${currentSheetId()}`;
}

function loadSheetsCached() {
  if (!currentSheetId()) return;
  try {
    const cached = JSON.parse(window.localStorage.getItem(settingsKey()));
    if (cached) {
      settings = cached;
      displaySettings();
    }
  } catch (error) {
    console.error(error);
  }
}

function loadSheets() {
  setMessage('info', 'Loading spreadsheet...');

  getSpreadsheet(`'${settingsSheet}'!A1:J`)
  .then(result => {
    // Keep the saved name up to date if the spreadsheet was renamed
    rememberSheet(currentSheetId(), result.properties.title);

    const rowData = result.sheets[0].data[0].rowData || [];
    const data = rowData.map(row => (row.values || []).map(cell => cell.formattedValue));
    if (data.length === 0) {
      setMessage('danger', 'No data found.');
      return;
    }

    loadSettings(data);

    setMessage('success', 'Loaded!');
  }).catch(error => {
    console.error(error);
    if (String(error).includes('Unable to parse range')) {
      setMessage('warning', `This spreadsheet has no ${year} tabs yet. ` +
        `<button type="button" class="btn btn-sm btn-primary ml-2" onclick="addYearTabs()">Add ${year} tabs</button>`);
    } else if (/permission|not found/i.test(error)) {
      // drive.file only reaches spreadsheets this Google account picked in Miler Lite
      setMessage('danger', 'Miler Lite can\'t open this spreadsheet with this Google account. Click <b>Open</b> to choose it again.');
    } else {
      setMessage('danger', 'Error: ' + error);
    }
  })
}

function dontSubmit(event) {
  if (event.keyCode == 13) {
    const focusable = Array.from(form.querySelectorAll('input:not([readonly]):not([type="radio"]):not([type="checkbox"]),button[type="submit"]'));
    const next = focusable[focusable.indexOf(event.target) + 1];
    if (next) {
      next.focus();
      return false;
    }
  }
  return true;
}

function loadSettings(values) {
  // business
  settings.businessHtml = '';
  for (let r = 2; r < values.length; r++) {
    const business = (values[r][0] || '').trim();
    if (business !== '') {
      // add business

      const index = r - 2;
      settings.businessHtml += $id('template-business-choice').innerHTML
        .replace(/\{\{INDEX\}\}/g, ''+index)
        .replace(/\{\{VALUE\}\}/g, escapeHTML(business))
        .replace(/\{\{LABEL\}\}/g, escapeHTML(business))
        .replace(/\{\{BUTTON_ACTIVE\}\}/g, index === 0 ? 'active' : '');
    }
  }

  // presets
  settings.presets = [];
  settings.presets.push({purpose: '', destination: '', mileage: 0});
  settings.presetsHtml = $id('template-preset').innerHTML
    .replace(/\{\{INDEX\}\}/g, 0)
    .replace(/\{\{PURPOSE\}\}/g, 'Reset')
    .replace(/\{\{BUTTON_CLASS\}\}/g, 'btn-outline-secondary');
  for (let r = 2; r < values.length; r++) {
    const purpose = (values[r][2] || '').trim();
    const destination = (values[r][3] || '').trim();
    const mileage = parseInt((values[r][4] || '').trim(), 10) || 0;
    if (purpose !== '' || destination !== '') {
      // add preset button

      const index = r - 1;
      settings.presets[index] = {purpose, destination, mileage};
      settings.presetsHtml += $id('template-preset').innerHTML
        .replace(/\{\{INDEX\}\}/g, ''+index)
        .replace(/\{\{PURPOSE\}\}/g, escapeHTML(purpose))
        .replace(/\{\{BUTTON_CLASS\}\}/g, 'btn-outline-primary');
    }
  }

  // values
  const valuesRow = values[2] || [];
  settings.maxMileage = parseInt(valuesRow[8], 10) || 0;
  settings.lastRowNum = parseInt(valuesRow[9], 10) || 4;

  window.localStorage.setItem(settingsKey(), JSON.stringify(settings));

  displaySettings();
}

function displaySettings() {
  $id('business-choices').innerHTML = settings.businessHtml;
  // make sure at least one is checked
  form.business[0].checked = true;

  $id('presets').innerHTML = settings.presetsHtml;

  $id('start').value = settings.maxMileage || '';
  $id('end').value = settings.maxMileage || '';
  recalcMileage();

  form.style.display = '';
}

function loadPreset(index) {
  const {purpose, destination, mileage} = settings.presets[index];
  $id('purpose').value = purpose;
  $id('destination').value = destination;
  $id('end').value = ((parseInt($id('start').value, 10) || 0) + mileage) || '';
  recalcMileage();
}

function validate({date, business, purpose, destination, start, end}) {
  $id('date').classList.remove('is-invalid');
  $id('business-choices').classList.remove('is-invalid');
  $id('purpose').classList.remove('is-invalid');
  $id('destination').classList.remove('is-invalid');
  $id('start').classList.remove('is-invalid');
  $id('end').classList.remove('is-invalid');
  $id('mileage').classList.remove('is-invalid');

  if (!date.match(/^\d{4}-\d{2}-\d{2}$/)) {
    setMessage('warning', 'Date is not valid.');
    $id('date').focus();
    $id('date').classList.add('is-invalid');
    return false;
  }
  if (date.substring(0, 4) !== (''+year)) {
    setMessage('warning', `Date is not in ${year}. ${date}`);
    $id('date').focus();
    $id('date').classList.add('is-invalid');
    return false;
  }

  if (business === '') {
    setMessage('warning', `Business is not specified.`);
    $id('business-0').focus();
    $id('business-choices').classList.add('is-invalid');
    return false;
  }

  if (purpose === '') {
    setMessage('warning', `Purpose is not specified.`);
    $id('purpose').focus();
    $id('purpose').classList.add('is-invalid');
    return false;
  }

  if (destination === '') {
    setMessage('warning', `Destination is not specified.`);
    $id('destination').focus();
    $id('destination').classList.add('is-invalid');
    return false;
  }

  if (start < 0) {
    setMessage('warning', `Start must be non-negative.`);
    $id('start').focus();
    $id('start').classList.add('is-invalid');
    return false;
  }

  if (start < settings.maxMileage) {
    setMessage('warning', `Start must be at least ${settings.maxMileage}.`);
    $id('start').focus();
    $id('start').classList.add('is-invalid');
    return false;
  }
  
  if (end < 0) {
    setMessage('warning', `End must be non-negative.`);
    $id('end').focus();
    $id('end').classList.add('is-invalid');
    return false;
  }

  if (start >= end) {
    setMessage('warning', `Mileage must be non-negative.`);
    $id('end').focus();
    $id('end').classList.add('is-invalid');
    $id('mileage').classList.add('is-invalid');
    return false;
  }

  return true;
}

function submitMileage() {
  try {
    setMessage('info', `Saving...`);

    const date = form['date'].valueAsDate.toISOString().substring(0, 10);
    const business = form['business'].value.trim();
    const purpose = form['purpose'].value.trim();
    const destination = form['destination'].value.trim();
    const start = parseInt(form['start'].value, 10) || 0;
    const end = parseInt(form['end'].value, 10) || 0;

    if (!validate({date, business, purpose, destination, start, end})) {
      return;
    }

    const row = [date, business, purpose, destination, start, end, `=F${settings.lastRowNum}-E${settings.lastRowNum}`, `=G${settings.lastRowNum}*'${settingsSheet}'!\$G\$3`];

    setSubmitEnabled(false);

    appendSpreadsheetRow(`'${mileageSheet}'!A${startingRow}:G`, row)
    .then(updates => {
      console.log(updates);
      setSubmitEnabled(true);
      setMessage('success', `Saved <a href="${spreadsheetLink()}" target="_blank" class="alert-link">${updates.updatedRange}</a>`);

      settings.lastRowNum += 1;
      if (settings.maxMileage < end) settings.maxMileage = end
      $id('start').value = end || '';
      $id('end').value = (end + (end - start)) || '';
      recalcMileage();

      window.localStorage.setItem(settingsKey(), JSON.stringify(settings));
    }).catch(error => {
      setSubmitEnabled(true);
      setMessage('danger', 'Error: ' + error);
    });
  } catch(error) {
    setSubmitEnabled(true);
    setMessage('danger', 'Error: ' + error);
  }
}

function setSubmitEnabled(shouldBeEnabled) {
  if (shouldBeEnabled) {
    $id('submit').removeAttribute('disabled');
    submittingMileage = false;
  } else {
    $id('submit').setAttribute('disabled', 'disabled');
    submittingMileage = true;
  }
}

