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
const DRIVE_API = 'https://www.googleapis.com/drive/v3/files';

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

// Saved settings hold the form's HTML, so bump this when the templates change
const settingsVersion = 3;
let settings = {
  settingsVersion,
  lastRowNum: 0,
  businesses: [],
  trips: [],
  rate: null,
  maxMileage: 0
}
let submittingMileage = false;
let tokenClient;
// Goalie and Miler Lite share one Google sign-in, so they share the saved token
const tokenKey = 'daily-tracking-token';
let accessToken = loadSavedToken();
let signedOut = false;
// The editor only opens on settings freshly loaded from this spreadsheet
let loadedSheetId = null;

// Older versions kept settings under a key that Goalie used too
window.localStorage.removeItem('settings');

function removeTime(date) {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
}
$id('date').valueAsDate = removeTime(new Date());

setMessage('info', 'Loading Google sign-in...');
$id('remove-button').onclick = handleRemoveClick;
$id('remove-dialog').querySelector('form').onsubmit = handleRemoveChoice;
$id('editor-add-business').onclick = () => addEditorBusiness('', true);
$id('editor-add-trip').onclick = () => addEditorTrip(null, true);
$id('editor-save').onclick = handleEditorSave;
$id('editor-cancel').onclick = displaySettings;
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

function currentSheetName() {
  const sheet = sheetList.recent.find(sheet => sheet.id === currentSheetId());
  return sheet ? sheet.name : '';
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
  link.style.display = currentSheetId() ? '' : 'none';
  link.href = spreadsheetLink();
}

function handleSheetSelect(event) {
  const sheet = sheetList.recent.find(sheet => sheet.id === event.target.value);
  if (!sheet) return;
  rememberSheet(sheet.id, sheet.name);
  hideSheet();
  loadSheetsCached();
  if (accessToken) loadSheets();
}

// Hides the form and the editor while another spreadsheet loads
function hideSheet() {
  form.style.display = 'none';
  $id('editor').style.display = 'none';
}

// Asks whether to just take the spreadsheet off the list, or trash it too
function handleRemoveClick() {
  if (!currentSheetId()) return;
  const dialog = $id('remove-dialog');
  $id('remove-name').textContent = currentSheetName();
  $id('remove-trash').hidden = !accessToken;
  $id('remove-signin-note').hidden = !!accessToken;
  dialog.showModal();
}

// Each dialog button closes it. Escape closes it without a submit
function handleRemoveChoice(event) {
  const choice = event.submitter && event.submitter.value;
  if (choice === 'list') removeCurrentSheet();
  if (choice === 'trash') trashCurrentSheet();
}

// Takes the spreadsheet off this browser's list and loads the next one
function removeCurrentSheet(doneMessage) {
  window.localStorage.removeItem(settingsKey());
  sheetList.recent = sheetList.recent.filter(sheet => sheet.id !== currentSheetId());
  sheetList.current = sheetList.recent.length ? sheetList.recent[0].id : null;
  window.localStorage.setItem(sheetsKey, JSON.stringify(sheetList));
  renderSheetControls();
  hideSheet();
  loadSheetsCached();
  if (accessToken && currentSheetId()) {
    loadSheets(doneMessage);
  } else {
    updateSigninStatus(!!accessToken);
    if (doneMessage) setMessage('success', doneMessage);
  }
}

// Trashed files stay in Google Drive's trash for 30 days
async function trashCurrentSheet() {
  const name = currentSheetName();
  setMessage('info', 'Moving the spreadsheet to the trash...');
  try {
    await googleFetch(`${DRIVE_API}/${currentSheetId()}`, {method: 'PATCH', body: JSON.stringify({trashed: true})});
  } catch (error) {
    console.error(error);
    return setMessage('danger', 'Could not move it to the trash: ' + error);
  }
  removeCurrentSheet(`Moved "${escapeHTML(name)}" to your Google Drive trash.`);
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
  $id('edit-button').onclick = handleEditClick;
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
  hideSheet();
  loadSheetsCached();
  loadSheets();
}

function handleCreateClick() {
  setMessage('info', 'Creating a new Miler Lite spreadsheet...');
  const body = {
    properties: {title: `Miler Lite ${year}`},
    sheets: yearTabs().map(tab => ({properties: {title: tab.title}, data: [{rowData: tab.rows}]})),
  };
  sheetsFetch('', {method: 'POST', body: JSON.stringify(body)})
  .then(result => {
    rememberSheet(result.spreadsheetId, result.properties.title);
    return loadSheets();
  }).then(loaded => {
    // Start in the editor so the example businesses and trips get replaced
    if (!loaded) return;
    handleEditClick();
    setMessage('success', 'Created! Set your rate and change the example businesses and saved trips to your own, then click Save.');
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
      // Submit inserts rows at the top of the range, which would shift plain
      // A4:A and F4:F references down and miss the first trip. INDIRECT stays put
      ['Work', null, 'Office', 'Main office', 10, null, 0.7, null,
        `=MAX(INDIRECT("'${mileageSheet}'!F${startingRow}:F"))`,
        `=COUNTA(INDIRECT("'${mileageSheet}'!A${startingRow}:A"))+${startingRow}`],
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
  for (const id of ['message', 'message-below']) {
    $id(id).innerHTML = message;
    $id(id).className = 'message ' + level;
  }

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

function sheetsFetch(path, options) {
  return googleFetch(`${SHEETS_API}${path}`, options);
}

// Calls a Google API and rejects with the error message
async function googleFetch(url, options = {}) {
  const headers = {...options.headers, Authorization: `Bearer ${accessToken}`};
  if (options.body) headers['Content-Type'] = 'application/json';
  const response = await fetch(url, {...options, headers});
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
    if (cached && cached.settingsVersion === settingsVersion) {
      settings = cached;
      displaySettings();
    }
  } catch (error) {
    console.error(error);
  }
}

// Resolves to true once the form is ready
function loadSheets(doneMessage = 'Loaded!') {
  setMessage('info', 'Loading spreadsheet...');

  return getSpreadsheet(`'${settingsSheet}'!A1:J`)
  .then(result => {
    // Keep the saved name up to date if the spreadsheet was renamed
    rememberSheet(currentSheetId(), result.properties.title);

    const rowData = result.sheets[0].data[0].rowData || [];
    const data = rowData.map(row => (row.values || []).map(cell => cell.formattedValue));
    if (data.length === 0) {
      setMessage('danger', `The ${settingsSheet} tab is empty.`);
      return false;
    }

    loadSettings(data);
    loadedSheetId = currentSheetId();

    setMessage('success', doneMessage);
    return true;
  }).catch(error => {
    console.error(error);
    if (String(error).includes('Unable to parse range')) {
      setMessage('warning', `This spreadsheet has no ${year} tabs yet. ` +
        `<button type="button" class="control small primary" onclick="addYearTabs()">Add ${year} tabs</button>`);
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
  const rows = values.slice(2);
  const cell = (row, column) => (row[column] || '').trim();

  settings.businesses = rows.map(row => cell(row, 0)).filter(Boolean);
  settings.businessHtml = settings.businesses.map((business, index) =>
    $id('template-business-choice').innerHTML
      .replace(/\{\{INDEX\}\}/g, ''+index)
      .replace(/\{\{VALUE\}\}/g, escapeHTML(business))
      .replace(/\{\{LABEL\}\}/g, escapeHTML(business))).join('');

  settings.trips = rows
    .map(row => ({purpose: cell(row, 2), destination: cell(row, 3), mileage: parseInt(cell(row, 4), 10) || 0}))
    .filter(trip => trip.purpose || trip.destination);
  // Button 0 is Reset, so the saved trips start at 1
  const presetButton = (index, label, className) => $id('template-preset').innerHTML
    .replace(/\{\{INDEX\}\}/g, ''+index)
    .replace(/\{\{PURPOSE\}\}/g, escapeHTML(label))
    .replace(/\{\{BUTTON_CLASS\}\}/g, className);
  settings.presetsHtml = presetButton(0, 'Reset', 'reset') +
    settings.trips.map((trip, index) => presetButton(index + 1, trip.purpose || trip.destination, '')).join('');

  const valuesRow = values[2] || [];
  // The rate may be formatted as money, like $0.70
  const rate = parseFloat(String(valuesRow[6] || '').replace(/[^\d.]/g, ''));
  settings.rate = isNaN(rate) ? null : rate;
  settings.maxMileage = parseInt(valuesRow[8], 10) || 0;
  settings.lastRowNum = parseInt(valuesRow[9], 10) || 4;

  window.localStorage.setItem(settingsKey(), JSON.stringify(settings));

  displaySettings();
}

function displaySettings() {
  $id('business-choices').innerHTML = settings.businessHtml;
  // make sure at least one is checked. form.business isn't a list when there's only one
  const firstBusiness = form.querySelector('input[name="business"]');
  if (firstBusiness) firstBusiness.checked = true;

  $id('presets').innerHTML = settings.presetsHtml;

  $id('start').value = settings.maxMileage || '';
  $id('end').value = settings.maxMileage || '';
  recalcMileage();

  $id('editor').style.display = 'none';
  form.style.display = '';
}

function loadPreset(index) {
  const reset = {purpose: '', destination: '', mileage: 0};
  const {purpose, destination, mileage} = index ? settings.trips[index - 1] : reset;
  $id('purpose').value = purpose;
  $id('destination').value = destination;
  $id('end').value = ((parseInt($id('start').value, 10) || 0) + mileage) || '';
  recalcMileage();
}

function handleEditClick() {
  if (!loadedSheetId || loadedSheetId !== currentSheetId()) {
    setMessage('warning', 'Wait for the spreadsheet to load, then click Edit.');
    return;
  }
  $id('editor-title').value = currentSheetName();
  $id('editor-rate').value = settings.rate === null ? '' : settings.rate;
  $id('editor-businesses').innerHTML = '';
  settings.businesses.forEach(business => addEditorBusiness(business, false));
  $id('editor-trips').innerHTML = '';
  settings.trips.forEach(trip => addEditorTrip(trip, false));
  form.style.display = 'none';
  $id('editor').style.display = '';
  setMessage('info', 'Rename the spreadsheet, change the rate, or rename, drag to reorder, add or remove businesses and saved trips, then click Save.');
}

function addEditorItem(listId, templateId, focus) {
  const item = document.createElement('div');
  item.className = 'item';
  item.innerHTML = $id(templateId).innerHTML;
  item.querySelector('.e-remove').onclick = () => item.remove();
  makeDraggable(item, item.querySelector('.drag-handle'));
  $id(listId).appendChild(item);
  if (focus) item.querySelector('input').focus();
  return item;
}

function addEditorBusiness(name, focus) {
  const item = addEditorItem('editor-businesses', 'template-editor-business', focus);
  item.querySelector('.e-name').value = name;
}

function addEditorTrip(trip, focus) {
  const item = addEditorItem('editor-trips', 'template-editor-trip', focus);
  if (!trip) return;
  item.querySelector('.e-purpose').value = trip.purpose;
  item.querySelector('.e-destination').value = trip.destination;
  item.querySelector('.e-miles').value = trip.mileage || '';
}

// Drag by the handle to reorder. Pointer events work for both mouse and touch,
// where the browser's own drag and drop doesn't work on phones
function makeDraggable(item, handle) {
  handle.onpointerdown = event => {
    event.preventDefault();
    item.classList.add('dragging');
    const move = event => {
      moveItem(item, event.clientY);
      // Scroll when dragging near the top or bottom of the screen
      if (event.clientY < 60) window.scrollBy(0, -10);
      if (event.clientY > window.innerHeight - 60) window.scrollBy(0, 10);
    };
    const stop = event => {
      if (event.type === 'pointerup') moveItem(item, event.clientY);
      item.classList.remove('dragging');
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', stop);
      window.removeEventListener('pointercancel', stop);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', stop);
    window.addEventListener('pointercancel', stop);
  };
  // The browser's own drag would swallow the pointer moves
  handle.ondragstart = event => event.preventDefault();
  // The arrow keys move it too. Moving the neighbors keeps the handle focused
  handle.onkeydown = event => {
    const prev = item.previousElementSibling;
    const next = item.nextElementSibling;
    if (event.key === 'ArrowUp' && prev) item.after(prev);
    else if (event.key === 'ArrowDown' && next) item.before(next);
    else return;
    event.preventDefault();
  };
}

// Moves the neighbors around the dragged item rather than the item itself,
// since moving it would cancel the drag on some phones
function moveItem(item, y) {
  const middle = other => other.getBoundingClientRect().top + other.offsetHeight / 2;
  let prev = item.previousElementSibling;
  while (prev && y < middle(prev)) {
    item.after(prev);
    prev = item.previousElementSibling;
  }
  let next = item.nextElementSibling;
  while (next && y > middle(next)) {
    item.before(next);
    next = item.nextElementSibling;
  }
}

async function handleEditorSave() {
  const value = (item, className) => item.querySelector(className).value.trim();
  const title = $id('editor-title').value.trim();
  const rateText = $id('editor-rate').value.trim();
  const rate = Number(rateText);
  const businesses = [...$id('editor-businesses').children].map(item => value(item, '.e-name'));
  const trips = [...$id('editor-trips').children].map(item => ({
    purpose: value(item, '.e-purpose'),
    destination: value(item, '.e-destination'),
    miles: value(item, '.e-miles'),
  }));

  if (!title) return setMessage('warning', 'The spreadsheet needs a name.');
  if (rateText === '' || !(rate >= 0)) return setMessage('warning', 'The rate per mile needs to be a number, like 0.70.');
  if (!businesses.length) return setMessage('warning', 'Add at least one business.');
  if (businesses.includes('')) return setMessage('warning', 'Every business needs a name.');
  if (trips.some(trip => !trip.purpose)) return setMessage('warning', 'Every saved trip needs a purpose.');
  if (trips.some(trip => !/^\d*$/.test(trip.miles))) return setMessage('warning', 'Miles need to be a whole number.');

  // Each saved trip's amount is a formula that uses the rate, so a new rate
  // changes the amounts of trips already saved this year too
  if (rate !== settings.rate) {
    try {
      const range = encodeURIComponent(`'${mileageSheet}'!A${startingRow}:A`);
      const saved = await sheetsFetch(`/${currentSheetId()}/values/${range}`);
      const count = saved.values ? saved.values.length : 0;
      if (count && !window.confirm(`Changing the rate also changes the amount of the ${count} trip${count === 1 ? '' : 's'} already saved in ${year}.\n\nSave anyway?`)) return;
    } catch (error) {
      console.error(error);
      return setMessage('danger', 'Error: ' + error);
    }
  }

  $id('editor-save').disabled = true;
  setMessage('info', 'Saving...');
  try {
    const sheetIds = await sheetsFetch(`/${currentSheetId()}?fields=sheets.properties(sheetId,title)`);
    const settingsId = sheetIds.sheets.find(sheet => sheet.properties.title === settingsSheet).properties.sheetId;
    const requests = editorRequests(settingsId, businesses, trips, rate);
    if (title !== currentSheetName()) {
      requests.push({updateSpreadsheetProperties: {properties: {title}, fields: 'title'}});
    }
    await sheetsFetch(`/${currentSheetId()}:batchUpdate`, {
      method: 'POST',
      body: JSON.stringify({requests}),
    });
    loadSheets('Saved!');
  } catch (error) {
    console.error(error);
    setMessage('danger', 'Error: ' + error);
  } finally {
    $id('editor-save').disabled = false;
  }
}

// Writes businesses to column A and saved trips to C to E from row 3 down,
// and the rate to G3. The formulas in I3 and J3 stay as they are
function editorRequests(sheetId, businesses, trips, rate) {
  // Strings are written as text, so a name starting with = isn't a formula
  const text = value => value ? {userEnteredValue: {stringValue: value}} : {};
  const number = value => value === '' ? {} : {userEnteredValue: {numberValue: Number(value)}};
  // With no end row, any old rows past the new list are cleared
  const write = (startColumnIndex, endColumnIndex, rows, endRowIndex) => ({updateCells: {
    range: {sheetId, startRowIndex: 2, endRowIndex, startColumnIndex, endColumnIndex},
    rows: rows.map(values => ({values})),
    fields: 'userEnteredValue',
  }});
  return [
    write(0, 1, businesses.map(business => [text(business)])),
    write(2, 5, trips.map(trip => [text(trip.purpose), text(trip.destination), number(trip.miles)])),
    write(6, 7, [[number(rate)]], 3),
  ];
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
      setMessage('success', `Saved <a href="${spreadsheetLink()}" target="_blank">${updates.updatedRange}</a>`);

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

