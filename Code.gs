// universal wrapper functions
function makeProper() { transformText('proper'); }
function makeUpper() { transformText('upper'); }
function makeLower() { transformText('lower'); }
function makeSentence() { transformText('sentence'); }

//universal menu launcher
function onOpen() {
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    try {
      ui = DocumentApp.getUi();
    } catch (e2) {
      try {
        ui = SlidesApp.getUi();
      } catch (e3) {
        return;
      }
    }
  }

  ui.createMenu('Capitalization')
    .addItem('Capitalize All Words', 'makeProper')
    .addItem('Capitalize First Letters', 'makeSentence')
    .addSeparator()
    .addItem('Make Uppercase', 'makeUpper')
    .addItem('Make Lowercase', 'makeLower')
    .addToUi();
}

// router function based on container type
function transformText(mode) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    if (sheet) {
      transformSheetsSelection(mode);
      return;
    }
  } catch (e) {}

  try {
    var doc = DocumentApp.getActiveDocument();
    if (doc) {
      transformDocsSelection(mode, doc);
      return;
    }
  } catch (e) {}

  try {
    var presentation = SlidesApp.getActivePresentation();
    if (presentation) {
      transformSlidesSelection(mode);
      return;
    }
  } catch (e) {}
}

//casing helper
function applyCase(text, mode) {
  if (!text || text.length === 0) return text;
  if (mode === 'upper') return text.toUpperCase();
  if (mode === 'lower') return text.toLowerCase();
  if (mode === 'proper') {
    return text.replace(/\b\w/g, function(l) { return l.toUpperCase(); });
  }
  if (mode === 'sentence') {
    return text.charAt(0).toUpperCase() + text.slice(1).toLowerCase();
  }
  return text;
}

//sheets handler
function transformSheetsSelection(mode) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var formulas = range.getFormulas();
  var values = range.getValues();
  var newFormulas = [];
  
  for (var r = 0; r < formulas.length; r++) {
    var row = [];
    for (var c = 0; c < formulas[r].length; c++) {
      var cellFormula = formulas[r][c];
      if (cellFormula !== "") {
        var cleanFormula = cellFormula.substring(1);
        if (mode === 'proper') row.push("=PROPER(" + cleanFormula + ")");
        else if (mode === 'upper') row.push("=UPPER(" + cleanFormula + ")");
        else if (mode === 'lower') row.push("=LOWER(" + cleanFormula + ")");
        else if (mode === 'sentence') row.push("=UPPER(LEFT(" + cleanFormula + ", 1)) & LOWER(MID(" + cleanFormula + ", 2, LEN(" + cleanFormula + ")))");
      } else {
        var text = values[r][c].toString();
        row.push(applyCase(text, mode));
      }
    }
    newFormulas.push(row);
  }
  range.setValues(newFormulas);
}

//docs handler
function transformDocsSelection(mode, doc) {
  var selection = doc.getSelection();
  if (!selection) return;
  
  var elements = selection.getRangeElements();
  for (var i = 0; i < elements.length; i++) {
    var element = elements[i];
    var textElement = element.getElement().asText();
    
    if (element.isPartial()) {
      var start = element.getStartOffset();
      var end = element.getEndOffsetInclusive();
      var originalText = textElement.getText().substring(start, end + 1);
      var newText = applyCase(originalText, mode);
      
      textElement.deleteText(start, end);
      textElement.insertText(start, newText);
    } else {
      var text = textElement.getText();
      textElement.setText(applyCase(text, mode));
    }
  }
}

//slides handler
function transformSlidesSelection(mode) {
  var selection = SlidesApp.getActivePresentation().getSelection();
  var textRange = selection.getTextRange();
  
  if (textRange) {
    var text = textRange.asString();
    textRange.setText(applyCase(text, mode));
  } else if (selection.getSelectionType() === SlidesApp.SelectionType.PAGE_ELEMENT) {
    var pageElements = selection.getPageElementRange().getPageElements();
    for (var i = 0; i < pageElements.length; i++) {
      if (pageElements[i].getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        var shapeText = pageElements[i].asShape().getText();
        shapeText.setText(applyCase(shapeText.asString(), mode));
      }
    }
  }
}
