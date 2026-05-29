/**
 * Creates a custom menu in the Google Docs UI when the document is opened.
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Canon Checker')
    .addItem('Open', 'showSidebar')
    .addToUi()
}

//Actually handles showing the sidebar, called above.
function showSidebar() {
  const html = HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Canon Checker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)

  DocumentApp.getUi()
    .showSidebar(html)
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function getDocumentText() {
  return DocumentApp.getActiveDocument().getBody().getText();
}

/**
 * Searches for a specific quote and moves the user's cursor to it.
 * @param {string} textToFind - The sentence to locate.
 */
function jumpToContradiction(textToFind) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // 1. Search for the text
  const searchResult = body.findText(textToFind);
  
  if (searchResult) {
    // 2. Get the specific text element and the offset
    const element = searchResult.getElement();
    const startOffset = searchResult.getStartOffset();
    const endOffset = searchResult.getEndOffsetInclusive();
    
    // 3. Create a "Range" (a selection)
    const range = doc.newRange()
        .addElement(element, startOffset, endOffset)
        .build();
  
    // 4. Set the user's active selection to this range
    doc.setSelection(range);

    doc.toast("HELLO THERE!")
    
    return true; // Tell the sidebar we found it
  }
  return false; // Not found
}