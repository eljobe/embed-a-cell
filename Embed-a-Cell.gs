/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi()
      .createAddonMenu()
      .addItem('Embedded Cells', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Embed-a-Cell');
  DocumentApp.getUi().showSidebar(ui);
}

function choose(sheetId, cell) {
  var sheet = SpreadsheetApp.openById(sheetId);
  var value = sheet.getRangeByName(cell).getDisplayValue();

  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();
  var element = cursor.getSurroundingText();
  var offset = cursor.getSurroundingTextOffset();

  cursor.insertText(value);
  recordNamedRange(doc, element, offset, offset + value.length - 1, sheetId, cell);
}

function refresh() {
  var doc = DocumentApp.getActiveDocument();
  var props = PropertiesService.getDocumentProperties();
  var keys = props.getKeys();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].substring(0, "embedacell-".length) === "embedacell-") {
      var rangeId = keys[i].substring("embedacell-".length);
      var range = doc.getNamedRangeById(rangeId);
      if (range != null) {
        var parts = props.getProperty(keys[i]).split(",");
        var sheetId = parts[0];
        var cell = parts[1];

        var sheet = SpreadsheetApp.openById(sheetId);
        var value = sheet.getRangeByName(cell).getDisplayValue();

        var rangeElement = range.getRange().getRangeElements()[0];
        var text = rangeElement.getElement().asText();
        text.deleteText(rangeElement.getStartOffset(), rangeElement.getEndOffsetInclusive());
        text.insertText(rangeElement.getStartOffset(), value);

        // Record a new range because deleting the text in the range effectively deletes the range.
        recordNamedRange(doc, text, rangeElement.getStartOffset(), rangeElement.getStartOffset() + value.length -1, sheetId, cell);
      }
      props.deleteProperty(keys[i]);
    }
  }
}

function addSheet(sheetId) {
  PropertiesService.getDocumentProperties().setProperty("embedacell-" + sheetId, sheetId);
  showSidebar();
}

function getSavedData() {
  var doc = DocumentApp.getActiveDocument();
  var props = PropertiesService.getDocumentProperties();
  var data = {}
  for (var key in props.getProperties()) {
    if (key.substring(0, "embedacell-".length) === "embedacell-") {
      var parts = props.getProperty(key).split(',');
      var sheetId = parts[0];
      var cell = "";
      if (parts.length > 1) {
        cell = parts[1];
      }

      if (!(sheetId in data)) {
        var sheet = SpreadsheetApp.openById(sheetId);
        var sheetData = {name: sheet.getName(), url: sheet.getUrl(), cells: []}
        data[sheetId] = sheetData;
      }
      Logger.log("Checking Cell: " + cell);
      if (cell !== "" && !contains(data[sheetId].cells, cell)) {
        data[sheetId].cells.push(cell);
      }
    }
  }
  if (Object.keys(data).length === 0 && data.constructor === Object) {
    throw "No data found!";
  }
  return data;
}

function recordNamedRange(doc, element, startOffset, endOffset, sheetId, cell) {
  var range = doc.newRange().addElement(element, startOffset, endOffset).build();
  var namedRange = doc.addNamedRange('embedacell-range', range);

  PropertiesService.getDocumentProperties()
      .setProperty("embedacell-" + namedRange.getId(), [sheetId, cell].join(","));
}

/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi().showModalDialog(html, 'Select a Spreadsheet');
}

/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

function clearLinks() {
  var keys = PropertiesService.getDocumentProperties().getKeys();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].substring(0, "embedacell-".length) === "embedacell-") {
      PropertiesService.getDocumentProperties().deleteProperty(keys[i]);
    }
  }
}

function contains(array, value) {
  for (var i = 0; i < array.length; i++) {
    if (array[i] === value) {
      return true;
    }
  }
  return false;
}
