// TODO: Need clarification and implementation!
// BEGIN: FORMAT DATE
function GENERATE_DATE(body) {
  body.appendParagraph(`\n\n${
    (new Date()).toLocaleDateString(undefined, {
      year: "numeric",
      month: "long",
      day: "numeric",
    })
  }\n\n`);
}
// END: FORMAT DATE

// NOTE: Setting attributes in this fashion seems so hacky.
// there has to be a better way.

// BEGIN: HEADER FORMATTING FUNCTIONS
function GENERATE_HEADER(contact, body) {
  const {first, last, comp, street, geo, phone, email} = contact;
  // TODO: Don't forget to handle this edge case in body generation.
  if (first || last) {
    const nameLine = body.appendParagraph(
      `${first || ""}` + " " + `${last || ""}`
    );
    nameLine.setAttributes(STYLE_BOLD);
  }
  if (comp) {
    const rest = body.appendParagraph(comp);
    rest.setAttributes(STYLE_DEFAULT);
  }
  if (street) body.appendParagraph(street);
  if (geo) body.appendParagraph(geo);
  if (phone) body.appendParagraph(phone);
  if (email) body.appendParagraph(email);
}
// END: HEADER FORMATTING FUNCTIONS

// BEGIN: TEXT FORMATTING FUNCTIONS
const formatCompStr = (title, comp) => title 
  ? comp 
      ? `${title}, ${comp}`
      : title
  : comp;

function GENERATE_RECIPIENT(recpt, body) {
  const {first, last, title, comp, street, geo} = recpt;
  // TODO: Don't forget to handle this edge case in body generation.
  if (first || last) {
    const nameLine = body.appendParagraph(
      `${first || ""}` + " " + `${last || ""}`
    );
    nameLine.setAttributes(STYLE_BOLD);
  }
  if (title || comp) {
    const compStr = formatCompStr(title, comp);
    const compLine = body.appendParagraph(compStr);
    compLine.setAttributes(STYLE_DEFAULT);
  }
  if (street) body.appendParagraph(street);
  if (geo) body.appendParagraph(geo);
}
// END: TEXT FORMATTING FUNCTIONS

function GENERATE_SALUTATION({first}, body) {
  body.appendParagraph("\n");
  body.appendParagraph(`Dear ${first},\n`);
}

function GENERATE_RULE(body) {
  const rule = body.appendHorizontalRule();
  rule.setAttributes(configureRuleStyle());
}

function GENERATE_LETTER(body) {
  body.appendParagraph(LETTER_CONTENT);
}

function GENERATE_SIGNATURE({first, last}, body) {
  body.appendParagraph("\n\n\n\n");
  const sigName = body.appendParagraph(`${first} ${last}`);
  sigName.setAttributes(STYLE_LARGE);
};

const openLogSheet = (id, name) => SpreadsheetApp
  .openById(id)
  .getSheetByName(name);

function loadColAsArr(colNum) {
  const sheet = openLogSheet(LOG_SHEET_ID, LOG_SHEET_NAME);
  const rowEnd = sheet.getLastRow();
  const rowRange = sheet.getRange(1, colNum, rowEnd, 1);
  return rowRange.getValues();
}

// BEGIN: GOOGLE DOC INTERFACE
// NOTE TO SELF:
// Functional style does cause second reading into memory after test.
// Perhaps all these objects should be held once and for all upon document
// entry.
const docExists = (docName) => DriveApp
  .getFilesByName(docName)
  .hasNext();

const fetchDocByName = (docName) => DocumentApp
  .openById(
    DriveApp.getFilesByName(docName).next().getId()
  );

const createDoc = (docName) => DocumentApp
  .create(docName);

// TODO: Have this function take a generalized object which formats
// output in the form of a GS object.
const createOrFetchDoc = (docName) => docExists(docName) 
  ? fetchDocByName(docName)
  : createDoc(docName);

const configureRuleStyle = () => ({
  [DocumentApp.Attribute.BOLD]: true, 
});

// SINGLE PAGE LOGIC
function compilePage(body, contact, client) {
  GENERATE_RULE(body);
  GENERATE_HEADER(contact, body);
  GENERATE_DATE(body);
  GENERATE_RECIPIENT(client, body);
  GENERATE_RULE(body);
  GENERATE_SALUTATION(client, body);
  GENERATE_LETTER(body);
  GENERATE_SIGNATURE(contact, body);
}

function compileFooter(doc) {
  if (!doc.getFooter()) doc.addFooter();
  const footer = doc.getFooter();
  footer.clear();
  const footerLine = footer.appendParagraph("Rogers - 11/03/2057");
  footer.setAttributes({
    [DocumentApp.Attribute.BOLD]: false,
    [DocumentApp.Attribute.FONT_SIZE]: 8,
  });
  footerLine.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function toClickablePath(doc) {
  const url = doc.getUrl();
  return HtmlService
    .createHtmlOutput(
      `<p>Document "${doc.getName()}" was successfullly created.</p>` +
      `<h2>Follow the link below to open document.</h2>` +
      `<a href="${url}" target="_blank">${url}</a>`
    );
}

// FULL COMPILATION
function compileDoc() {
  const doc = createOrFetchDoc(`sales-log-${(new Date()).toLocaleDateString()}`);
  const body = doc.getBody();
  body.clear();

  const {contact, clients} = fetchData();

  clients.forEach((client, n) => {
    if (n > 0) body.insertPageBreak(body.getNumChildren());
    compilePage(body, contact, client);
  });

  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(toClickablePath(doc), "Compilation successful.");
  doc.saveAndClose();
}
// END: GOOGLE DOC INTERFACE

const cellsBounded = (row, col, aspect, bound) => col < aspect 
  && (row * aspect + col + 1) <= bound;

const clientToLabel = ({first, last, street, geo}) => `${first} `
  + `${last}\n${street}\n${geo}`;

function compileLbl() {
  const doc = createOrFetchDoc(
    `sales-labels-${(new Date()).toLocaleDateString()}`
  );
  
  const body = doc.getBody();
  body.clear();

  const {clients} = fetchData();

  const margin = FONT_SIZE_MED;
  const padding = 6;

  body.setMarginTop(0);
  body.setMarginBottom(margin);
  body.setMarginLeft(margin);
  body.setMarginRight(margin);

  const pageWidth = body.getPageWidth();
  const pageHeight = body.getPageHeight();

  const numLabels  = clients.length;
  const aspectRows = 10;
  const aspectCols = 3;
  const numRows = Math.ceil(numLabels / aspectCols);

  const availableWidth = pageWidth - 2 * margin;
  const availableHeight = pageHeight - 3 * margin;

  const cellWidth = availableWidth / aspectCols;
  const cellHeight = availableHeight / aspectRows;

  const table = body.appendTable();

  // Instantiate table cells
  for (let row = 0; row < numRows; row++) {
    const tr = table.appendTableRow();
    tr.setMinimumHeight(cellHeight);
    for (let col = 0; cellsBounded(row,col,aspectCols,numLabels); col++) {
      const td = tr.appendTableCell(
        clientToLabel(clients[row * aspectCols + col])
      );
      td.setWidth(cellWidth);
      td.setAttributes({
        [DocumentApp.Attribute.BOLD]: true,
        [DocumentApp.Attribute.FONT_SIZE]: 12,
        [DocumentApp.Attribute.VERTICAL_ALIGNMENT]: DocumentApp.VerticalAlignment.CENTER
      });
    }
  }

  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(
    toClickablePath(doc), 
    `Success: ${numLabels} labels generated.`
  );
  doc.saveAndClose();
}

// BEGIN: FETCHING SPREADSHEET
function fetchData() {
  const ss = SpreadsheetApp.openById(LOG_SHEET_ID);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  const raw = log.getDataRange().getValues();
  const data = {
    contact: {
      phone:  log.getRange(LOG_PHONE_CELL).getValue(),
      street: log.getRange(LOG_STREET_CELL).getValue(),
      geo:    log.getRange(LOG_GEO_CELL).getValue(),
      email:  log.getRange(LOG_EMAIL_CELL).getValue(),
      comp:   log.getRange(LOG_COMP_CELL).getValue(),
      first:  log.getRange(LOG_NAME_CELL).getValue().split(" ")[0],
      last:   log.getRange(LOG_NAME_CELL).getValue().split(' ')[1],
    },
  };
  data["clients"] = raw
    .slice(LOG_LABEL_END)
    .reduce((cList, [first, last, comp, street, geo, phone, email]) => [
      ...cList, 
      {first, last, comp, street, geo, phone, email}
    ], []);
  return data;
}
// END: FETCHING SPREADSHEET

// Generate UI element to execute main code.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("COMPILE")
    .addItem("Letters", "compileDoc")
    .addItem("Labels", "compileLbl")
    .addToUi();
}