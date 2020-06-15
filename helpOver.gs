function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Helpover')
      .addItem('Create new handover', 'askQuestions')
      .addToUi();
}

function main(_name, _itMember, _tableData, _extras) {
  
  // --------------------- STORE DATA --------------------- //
  var extraCharger = {
    number : "",
    name : "Extra charger",
    serial : "",
    quantity : 1
  }
  
  var usbc = {
    number : "",
    name : "Usb-c hub",
    serial : "",
    quantity : 1
  }
  
  var nonda = {
    number : "",
    name : "Nonda usb-c to usb-a adapter",
    serial : "",
    quantity : 1
  }
  
  var extrasData = [extraCharger, usbc, nonda];

  var date = {
    year : Utilities.formatDate(new Date(), "GMT+1", "yyyy"),
    month : Utilities.formatDate(new Date(), "GMT+1", "MM"),
    day : Utilities.formatDate(new Date(), "GMT+1", "dd")                             
  }
  // ----------------------------------------------------- //
  
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  // insert the information
  mailMerge(_name, _itMember, date, body);
  
  // insert the tables
  var tables = body.getTables();
  addToTable(tables[0], _tableData);
  addToTable(tables[1], _tableData);
  // add the extras if checked, in a not-so elegant way
  for (var i = 0; i < _extras.length; i++) {
    if (_extras[i]) {
      addToTable(tables[0], extrasData[i]);
      addToTable(tables[1], extrasData[i]);
    }
  }
  
  // create the new document
  var newDoc = DocumentApp.create(_name + " handover");
  copyDocWithoutScript(doc.getId(), newDoc.getId());
  
  // do I need this?
  doc.saveAndClose();
  newDoc.saveAndClose();
}

function mailMerge(_name, _itMember, _date, _body) {
  _body.replaceText("{{name}}", _name);
  _body.replaceText("{{year}}", _date.year);
  _body.replaceText("{{month}}", _date.month);
  _body.replaceText("{{day}}", _date.day);
  _body.replaceText("{{itMember}}", _itMember);
}

function addToTable(_table, _data) {
  var tr = _table.appendTableRow();
  for (var data in _data) {
    tr.appendTableCell(_data[data]).setBold(false);
  }
}

function askQuestions() {
  var ui = DocumentApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('form.html').setHeight(400);
  ui.showModalDialog(html, "Handover");
}

function processForm(form) {
  var name = form.name;
  var itMember = form.itMember;
  
  var tableData = {
    inventoryNum : form.inventoryNum,
    assetName : form.assetName,
    serial : form.serial,
    quantity : form.quantity,
  }
  
  var extras = [form.extraCharger, form.usbc, form.nonda]
  
  main(name, itMember, tableData, extras);
}

// https://stackoverflow.com/questions/10692669/how-can-i-generate-a-multipage-text-document-from-a-single-page-template-in-goog
// if you copy it with DriveApp, the bounded script gets copied to
function copyDocWithoutScript(_fromId, _toId) {
  
  var to = DocumentApp.openById(_toId);
  var toBody = to.getActiveSection();
  var fromBody = DocumentApp.openById(_fromId).getActiveSection();
  var totalElements = fromBody.getNumChildren();
  
  for (var i = 0; i < totalElements; ++i ) {
    var element = fromBody.getChild(i).copy();
    var type = element.getType();
    if( type == DocumentApp.ElementType.PARAGRAPH )
      toBody.appendParagraph(element);
    else if( type == DocumentApp.ElementType.TABLE )
      toBody.appendTable(element);
    else if( type == DocumentApp.ElementType.LIST_ITEM )
      toBody.appendListItem(element);
  }
}