// Client ID and API key from the Developer Console
//localhost:8000, andrewmacheret.com
const CLIENT_ID = '303548077940-ov8iafec5pqhrd457fhe8sb2q5ak6o8s.apps.googleusercontent.com';
const API_KEY = 'AIzaSyAF3bxFGxELOH7aK3imyc__dpRuo9M2j5M';

// Array of API discovery doc URLs for APIs used by the quickstart
const DISCOVERY_DOCS = ["https://sheets.googleapis.com/$discovery/rest?version=v4"];

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
const SCOPES = "https://www.googleapis.com/auth/spreadsheets";

const formName = 'mileage-form';
const form = document.forms[formName];

const $id = document.getElementById.bind(document);

const spreadsheetId = '1witWYvqss5Y_7dDARZjXuixkh8ySaaEtnLyHkBk4lYc';
const spreadsheetLink = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
const year = new Date().getFullYear();
const mileageSheet = `${year} Mileage`;
const settingsSheet = `${year} Settings`;
const startingRow = 4;

let settings = {
  lastRowNum: 0,
  presets: [],
  maxMileage: 0
}
let submittingMileage = false;

function removeTime(date) {
  return new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
}
$id('date').valueAsDate = removeTime(new Date());
$id('spreadsheet-link').href = spreadsheetLink;

setMessage('info', 'Loading Google APIs...');
let autoreload = setTimeout(function() {
  setMessage('info', 'Google APIs timed out, reloading...');
  window.location.reload();
}, 5000);

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


/**
 *  On load, called to load the auth2 library and API client library.
 */
function handleClientLoad() {
  gapi.load('client:auth2', initClient);
}

/**
 *  Initializes the API client library and sets up sign-in state
 *  listeners.
 */
function initClient() {
  gapi.client.init({
    apiKey: API_KEY,
    clientId: CLIENT_ID,
    discoveryDocs: DISCOVERY_DOCS,
    scope: SCOPES
  }).then(() => {
    window.clearTimeout(autoreload);

    // Listen for sign-in state changes.
    gapi.auth2.getAuthInstance().isSignedIn.listen(updateSigninStatus);

    // Handle the initial sign-in state.
    updateSigninStatus(gapi.auth2.getAuthInstance().isSignedIn.get());
    $id('authorize-button').onclick = handleAuthClick;
    $id('signout-button').onclick = handleSignoutClick;
  });
}

/**
 *  Called when the signed in status changes, to update the UI
 *  appropriately. After a sign-in, the API is called.
 */
function updateSigninStatus(isSignedIn) {
  if (isSignedIn) {
    $id('authorize-button').style.display = 'none';
    $id('signout-button').style.display = '';
    loadSheets();
  } else {
    $id('authorize-button').style.display = '';
    $id('signout-button').style.display = 'none';
    setMessage('warning', 'Need authorization.');
  }
}

/**
 *  Sign in the user upon button click.
 */
function handleAuthClick(event) {
  gapi.auth2.getAuthInstance().signIn();
}

/**
 *  Sign out the user upon button click.
 */
function handleSignoutClick(event) {
  gapi.auth2.getAuthInstance().signOut();
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

function getSpreadsheetValues(range) {
  return new Promise((resolve, reject) => {
    gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId,
      range
    }).then(response => {
      resolve(response.result.values);
    }, response => {
      reject(response.result.error.message);
    });
  });
}

function appendSpreadsheetRow(range, row) {

  return new Promise((resolve, reject) => {
    const params = {
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      includeValuesInResponse: false
    };

    const valueRange = {
      'values': [
        row
      ],
    };

    gapi.client.sheets.spreadsheets.values.append(params, valueRange).then(response => {
      resolve(response.result.updates);
    }, response => {
      reject(response.result.error.message);
    });
  });
}

function loadSheetsCached() {
  let cached = window.localStorage.getItem('settings');
  if (!cached) return;
  
  settings = JSON.parse(window.localStorage.getItem('settings'));

  displaySettings();
}

function loadSheets() {
  setMessage('info', 'Loading sheets...');

  getSpreadsheetValues(`'${settingsSheet}'!A1:J`)
  .then(data => {
    if (data.length === 0) {
      setMessage('danger', 'No data found.');
      return;
    }

    loadSettings(data);

    setMessage('success', 'Loaded!');
  }).catch(error => {
    setMessage('danger', 'Error: ' + error);
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
  settings.maxMileage = parseInt(values[2][8], 10) || 0;
  settings.lastRowNum = parseInt(values[2][9], 10) || 4;

  window.localStorage.setItem('settings', JSON.stringify(settings));

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
      setMessage('success', `Saved <a href="${spreadsheetLink}" target="_blank" class="alert-link">${updates.updatedRange}</a>`);

      settings.lastRowNum += 1;
      if (settings.maxMileage < end) settings.maxMileage = end
      $id('start').value = end || '';
      $id('end').value = (end + (end - start)) || '';
      recalcMileage();

      window.localStorage.setItem('settings', JSON.stringify(settings));
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

