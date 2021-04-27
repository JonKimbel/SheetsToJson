// TODO:
// https://support.google.com/cloud/answer/9110914#verification-requirements
// https://developers.google.com/workspace/add-ons/how-tos/publish-add-on-overview
// https://developers.google.com/workspace/marketplace/how-to-publish
// https://developers.google.com/apps-script/guides/cloud-platform-projects
// Add more info to jonkimbel.com about oauth stuff, how i'll have access to gist oauth key and can read gists
// Make the menu more reasonable

// Defining a schema in the sheet:
//  1. Only cells with font Courier New will be read as data.
//  2. Each key name is a named range.
//  3. If you need to repeat a name, append some number of '_'s to the end.
//  4. If you need to group together keys within a list, just name it '_' (or __, or ___, etc).
//  5. If you're getting an object and you need a list, add an empty cell with font Courier New to the named range.

const DATA_FONT_FAMILY = "Courier New";
const DATA_BACKGROUND = "yellow";
const DATA_TYPE = {
  "list":1, // A list
  "dict":2, // A dict
  "key":3, // A named piece of data in a dict
  "value":4}; // Primitive data, either in a list or keyed in a dict
  // Key in document properties. If present, user has authorized use of their user-tied Github auth key for this
  // doc.
const OAUTH_AWARENESS_KEY = "isAwareOfOAuthLink";

// Called by Apps Script.
function onInstall() {
  onOpen();
}

// Called by Apps Script.
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Sheets to JSON", [
    {name: "Import", functionName: "fromJson"},
    {name: "Export", functionName: "toJson"},
    {name: "Export 'JSON' sheet", functionName: "toJsonNamedSheet"},
    {name: "Export 'JSON' to PUBLIC Gist", functionName: "toJsonNamedSheetGist"},
    {name: "Name range", functionName: "name"},
    {name: "Print range type", functionName: "printRangeType"},
    {name: "Purge named ranges", functionName: "purgeNamedRanges"},
    {name: "Revoke Github access", functionName: "logout"},
    {name: "[DEBUG] Export", functionName: "toJsonDebug"},
    {name: "[DEBUG] Highlight data fields", functionName: "highlight"},
  ]);
}

/**
 * Called by OAuth2 library, this function name is specified by GithubGistClient.
 * 
 * Requires https://www.googleapis.com/auth/script.external_request
 */
function authCallback(request) {
  var client = new GithubGistClient();
  return client.handleCallback(request);
}

/**
 * Requires
 * https://www.googleapis.com/auth/script.external_request
 * https://www.googleapis.com/auth/script.container.ui
 */
function logout() {
  new GithubGistClient().logout();
  openUrl("https://github.com/settings/applications");
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function name() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter a name:');

  if (response.getSelectedButton() == ui.Button.OK) {
    var selection = SpreadsheetApp.getActiveSheet().getSelection();
    nameRange(selection.getActiveRange(), response.getResponseText());
  }
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function printRangeType() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('Range type', typeof SpreadsheetApp.getActiveSheet().getSelection().getActiveRange().getValue(), ui.ButtonSet.OK) ;
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function fromJson() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirmation', 'Are you sure you want to overwrite the active sheet?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange(1,1);
  assert(range.getValue().toString().length > 0, "Please enter your JSON in cell A1 of this sheet.");
  var object = JSON.parse(range.getValue());
  range.setValue("")
  writeObjectToSheet(object);
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function purgeNamedRanges() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirmation', 'Are you sure you want to delete all named ranges?', ui.ButtonSet.YES_NO);
  if (response != ui.Button.YES) {
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nrs = ss.getNamedRanges();

  nrs.forEach(nr => ss.removeNamedRange(nr.getName()));
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function nameRange(range, name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nrs = ss.getNamedRanges();
  name = trimKeyName(name);

  var dupeSet = new Set();
  nrs.forEach(nr => {
    if (trimKeyName(nr.getName()) == name) {
      dupeSet.add(nr.getName());
    }
  });

  name = name.length == 0 ? "_" : name;
  if (name.match(/\d+/)) {
    name = "REKT" + name;
  }
  while (dupeSet.has(name)) {
    name += "_";
  }
  console.log("trying to name a range '" + name + "'");
  ss.setNamedRange(name, range);
  console.log("success!");
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function writeObjectToSheet(object, loc = {"col": 1, "row": 1, "maxCol": 1}, name = null) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  console.log("click");
  
  if (object == null
      || typeof object == 'number'
      || typeof object == 'string'
      || typeof object == 'boolean') {
    var range = sheet.getRange(loc.row, loc.col);
    range.setFontFamily(DATA_FONT_FAMILY);
    range.setBackground(DATA_BACKGROUND);
    if (typeof object == 'string') {
      range.setNumberFormat("@");
    }
    if (object != null) {
      // HACK: we should probably actually write NULL to the sheet, but we don't because we use this behavior to make sure
      // imported arrays are read as arrays by the exporter.
      range.setValue(object);
    }
    if (name != null) {
      nameRange(range, name);
    }
    loc.row++;
    return loc;
  }

  var startRow = loc.row;

  if (Array.isArray(object)) {
    for (var i = 0; i < object.length; i++) {
      loc = writeObjectToSheet(object[i], loc);
    }
    // Write one null object to the sheet to ensure the one-length arrays are read by the exporter as an array.
    loc = writeObjectToSheet(null, loc);
  } else { // Must be an object
    var keys = Object.getOwnPropertyNames(object);
    for (var i = 0; i < keys.length; i++) {
      var labelRange = sheet.getRange(loc.row, loc.col);
      labelRange.setValue(keys[i]);

      loc.col++;
      loc.maxCol = loc.col > loc.maxCol ? loc.col : loc.maxCol;
      
      loc = writeObjectToSheet(object[keys[i]], loc, keys[i]);
      loc.col--;
    }
  }
  
  if (loc.col > 1 || Array.isArray(object)) { // If we wrap a root dict in an anonymous range it'll be wrapped in an array at export time.
    var containingRange = sheet.getRange(startRow, loc.col, loc.row - startRow, loc.maxCol - loc.col + 1);
    nameRange(containingRange, name != null ? name : "_");
  }
  return loc;
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function highlight() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var namedRanges = sheet.getNamedRanges();

  for (var i = 0; i < namedRanges.length; i++) {
    var rect = rangeToRect(namedRanges[i].getRange());
    for (var col = rect.left; col <= rect.right; col++) {
      for (var row = rect.top; row <= rect.bottom; row++) {
        var range = sheet.getRange(row,col);
        if (range.getFontFamily() == DATA_FONT_FAMILY) {
          if (range.getBackground() != DATA_BACKGROUND) {
            range.setBackground(DATA_BACKGROUND);
          }
        }
      }
    }
  }
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function toJsonDebug() {
  toJson(true);
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function toJsonNamedSheet() {
  toJson(false, SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JSON"));
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function toJsonNamedSheetGist() {
  toJson(false, SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JSON"), true);
}

/**
 * Requires
 * https://www.googleapis.com/auth/spreadsheets.currentonly
 * https://www.googleapis.com/auth/script.external_request
 * https://www.googleapis.com/auth/script.container.ui
 */
function toJson(debug=false, sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(), toGist=false) {
  var client;
  if (toGist) {
    client = new GithubGistClient();
    if (!client.hasAccess()) {
      openUrl(client.getAuthorizationUrl());
      return;
    } else if (!client.tokenHasBeenUsedInThisSheet()) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert('Confirmation', 'Would you like to publish using the Github account you\'ve linked with the Sheets to JSON addon?', ui.ButtonSet.YES_NO);
      if (response != ui.Button.YES) {
        logout();
        toJson(debug, sheet, toGist);
        return;
      }
      client.logTokenUseInThisSheet();
    }
  }

  SpreadsheetApp.flush();
  var namedRanges = sheet.getNamedRanges();
  var intermediates = [];
  for (var i = 0; i < namedRanges.length; i++) {
    intermediates.push(makeIntermediate(namedRanges, i));
  }
  for (var i = 0; i < intermediates.length; i++) {
    populateChildren(intermediates, i);
  }

  var root = {};
  root.rangeName = 'root';
  root.objects = Array.from(trimChildrenAndFindRootIntermediates(new Set(intermediates)));

  for (var i = 0; i < intermediates.length; i++) {
    setValues(intermediates[i], sheet);
  }

  setDataTypes(root);
  sortRecursive(root);
  trimEmpty(root);

  if (debug) {
    displayText(traverseForDebug(root, ''));
  } else {
    var jsonString = JSON.stringify(createJsonObjectOf(root, /* forDict = */ true), null, 2);
    if (toGist) {
      var gistUrl = client.newGist(
        /* content = */ jsonString,
        /* filename = */ "sheets.json",
        /* language = */ "json",
        /* description = */ "JSON object exported by the 'Sheets to JSON' Google Sheets addon.",
        /* public = */ true);
      displayText(gistUrl);
    } else {
      displayText(jsonString);
    }
  }
}

function sortRecursive(root) {
  if (root.dataType == DATA_TYPE.value) {
    return;
  }

  for (var i = 0; i < root.objects.length; i++) {
    sortRecursive(root.objects[i]);
  }

  root.objects.sort(topLeftLocationComparator);
}

function setDataTypes(root) {
  root.dataType = getDataType(root);
  if (root.dataType == DATA_TYPE.value) {
    return;
  }

  for (var i = 0; i < root.objects.length; i++) {
    setDataTypes(root.objects[i]);
  }
}

function trimEmpty(root) {
  if (root.dataType == DATA_TYPE.value) {
    return false;
  }

  for (var i = root.objects.length - 1; i >= 0; i--) {
    if (trimEmpty(root.objects[i])) {
      if (root.objects[i].objects.length == 0) {
        root.objects.splice(i, 1);
      }
    }
  }
  return true;
}

function traverseForDebug(intermediate, prefix) {
  if (intermediate.dataType == DATA_TYPE.value) {
    return "";
  }
  var str;
  str = prefix + 'rangeName:' + intermediate.rangeName + ', type:'
      + (intermediate.dataType == DATA_TYPE.key ? "key" : "")
      + (intermediate.dataType == DATA_TYPE.dict ? "dict" : "")
      + (intermediate.dataType == DATA_TYPE.list ? "list" : "") 
      + "\n";
  for (var i = 0; i < intermediate.objects.length; i++) {
    str += traverseForDebug(intermediate.objects[i], prefix + ' ');
  }
  return str;
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function makeIntermediate(nrs, index) {
  var intermediate = {};
  intermediate.rangeName = nrs[index].getName();
  intermediate.keyName = trimKeyName(nrs[index].getName());
  intermediate.rect = rangeToRect(nrs[index].getRange());
  intermediate.objects = [];
  intermediate.numberCellsForValues = 0;
  intermediate.dataType = 0;
  return intermediate;
}

function populateChildren(intermediates, index) {
  var intermediate = intermediates[index];
  for (var i = 0; i < intermediates.length; i++) {
    if (i == index) {
      continue;
    }
    if (contains(intermediate.rect, intermediates[i].rect)) {
      intermediate.objects.push(intermediates[i]);
    }
  }
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function setValues(intermediate, sheet) {
  for (var col = intermediate.rect.left; col <= intermediate.rect.right; col++) {
    for (var row = intermediate.rect.top; row <= intermediate.rect.bottom; row++) {
      var range = sheet.getRange(row, col);
      if (range.getFontFamily() != DATA_FONT_FAMILY) {
        continue;
      }
      var childOwned = false;
      for (var i = 0; !childOwned && i < intermediate.objects.length; i++) {
        if (containsCell(intermediate.objects[i].rect, row, col)) {
          childOwned = true;
        }
      }
      if (!childOwned) {
        intermediate.numberCellsForValues++;
        if (!isCellEmpty(range.getValue())) {
          intermediate.objects.push({ "rect": rangeToRect(range), "value": range.getValue(), "dataType": DATA_TYPE.value });
        }
      }
    }
  }
}

function trimChildrenAndFindRootIntermediates(intermediateSet, depth = 0) {
  if (intermediateSet.size == 0) {
    return intermediateSet;
  }

  // Find roots.
  var roots = new Set(intermediateSet);
  intermediateSet.forEach(i =>
    i.objects.forEach(child => roots.delete(child)));

  // Cut off the roots from intermediateSet and recurse to find direct children.
  roots.forEach(r => intermediateSet.delete(r));
  var directChildren = trimChildrenAndFindRootIntermediates(intermediateSet);

  // Remove everything but direct children from roots.
  roots.forEach(r => {
    for (var i = r.objects.length - 1; i >= 0; i--) {
      if (!directChildren.has(r.objects[i])) {
        r.objects.splice(i, 1);
      }
    }
  });

  return roots;
}

function createJsonObjectOf(intermediate, forDict) {
  if (intermediate.dataType == DATA_TYPE.value) {
    return intermediate.value;
  }
  if (intermediate.dataType == DATA_TYPE.key) {
    return intermediate.objects.length > 0 ? createJsonObjectOf(intermediate.objects[0]) : null;
  }
  if (intermediate.dataType == DATA_TYPE.dict) {
    var rootDict = {};
    intermediate.objects.forEach(i => rootDict[i.keyName] = createJsonObjectOf(i));
    return rootDict;
  }
  assert(intermediate.dataType == DATA_TYPE.list, "Unknown dataType: " + intermediate.dataType + "!");
  var rootArray = [];
  intermediate.objects.forEach(i => rootArray.push(createJsonObjectOf(i)));
  return rootArray;
}

function getDataType(intermediate) {
  if (intermediate.objects == undefined) {
    return DATA_TYPE.value;
  }
  if (isKey(intermediate)) {
    return DATA_TYPE.key;
  }
  if (isDictObject(intermediate)) {
    return DATA_TYPE.dict;
  }
  return DATA_TYPE.list;
}

function isKey(intermediate) {
  if (intermediate.numberCellsForValues == 1) {
    if (intermediate.objects.length == 0) {
      return true;
    }
    return intermediate.objects.length == 1 && getDataType(intermediate.objects[0]) == DATA_TYPE.value;
  }
  return false;
}

function isDictObject(intermediate) {
  if (intermediate.numberCellsForValues > 0) {
    return false;
  }
  var objectKeyNames = new Set();
  for (var i = 0; i < intermediate.objects.length; i++) {
    if (intermediate.objects[i].keyName.length == 0) {
      return false;
    }
    if (objectKeyNames.has(intermediate.objects[i].keyName)) {
      return false;
    }
    objectKeyNames.add(intermediate.objects[i].keyName);
  }
  return true;
}

function trimKeyName(name) {
  return name.replace(/^(REKT)?(.*[^_])_*$/, "$2").replace(/^_*$/, "");
}

/**
 * Requires https://www.googleapis.com/auth/spreadsheets.currentonly
 */
function rangeToRect(range) {
  var rect = {};
  rect.left = range.getColumn();
  rect.right = range.getLastColumn();
  rect.top = range.getRow();
  rect.bottom = range.getLastRow();
  return rect;
}

function topLeftLocationComparator(intermediate1, intermediate2) {
  if (intermediate1.rect.top < intermediate2.rect.top) {
    return -1;
  }
  if (intermediate1.rect.top > intermediate2.rect.top) {
    return 1;
  }
  if (intermediate1.rect.left < intermediate2.rect.left) {
    return -1;
  }
  if (intermediate1.rect.left > intermediate2.rect.left) {
    return -1;
  }
  return 0;
}

function contains(rect1, rect2) {
  if (rect1.top <= rect2.top
      && rect1.left <= rect2.left
      && rect1.bottom >= rect2.bottom
      && rect1.right >= rect2.right) {
    return rect1.top < rect2.top
      || rect1.left < rect2.left
      || rect1.bottom > rect2.bottom
      || rect1.right > rect2.right;
  }
  return false;
}

function containsCell(rect1, row, col) {
  return rect1.top <= row
      && rect1.left <= col
      && rect1.bottom >= row
      && rect1.right >= col;
}

/**
 * Open a URL in a new tab.
 * 
 * Requires https://www.googleapis.com/auth/script.container.ui
 */
// https://stackoverflow.com/a/47098533
function openUrl(url){
  var template = HtmlService.createTemplateFromFile('openUrl');
  template.data = url;
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setHeight(50),
    'Opening URL');
}

/**
 * Requires https://www.googleapis.com/auth/script.container.ui
 */
function displayText(text) {
  var template = HtmlService.createTemplateFromFile('displayText');
  template.data = text;
  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(700).setHeight(500),
    'Sheets to JSON output');
}

function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

function assert(condition, message) {
    if (!condition) {
        throw message || "Assertion failed";
    }
}