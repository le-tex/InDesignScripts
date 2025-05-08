#targetengine "session"

var version = "v1.0.0";
/*
 * set language
 */
var lang = {
  pre: app.locale == 1279477613 ? 1 : 0 // en = 0, de = 1
}
/*
 * options object
 */
options = {
  schemaDir:"export",
  exportPackageSchema:true
}
/*
 * set panel preferences
 */
panel = {
  title:["le-tex – Generate IDML schema", "le-tex – IDML-Schema erzeugen"][lang.pre],
  selectDirMenuTitle:["Save schema as", "Schema speichern unter"][lang.pre],
  selectDirButtonTitle:["Choose", "Auswählen"][lang.pre],
  chooseSchema:["Choose schema", "Schema auswählen"][lang.pre],
  optionsTitle:["Choose schema-type", "Schema-Typ auswählen"][lang.pre],
  exportPackageSchemaValues:[["Package", "File"], ["Paket", "Datei"]][lang.pre],
  buttonOK:"OK",
  buttonCancel:["Cancel", "Abbrechen"][lang.pre],
  finishedMessage:["Schema exported!", "Schema exportiert!"][lang.pre]
}
/*
 * start
 */
main();
/*
 * main pipeline
 */
function main(){
  var userLevel = app.scriptPreferences.userInteractionLevel;
  app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;    
  // open file dialog and load a document
  if (app.layoutWindows.length == 0) {
    var file = File.openDialog ("Select a file", "InDesign:*.indd;*.indb;*.idml, InDesign Document:*.indd, InDesign Book:*.indb, InDesign Markup:*.idml", true)
    try {
      app.open(File(file));
    } catch (e) {
      alert(e);
      return;
    };
  }
  var doc = app.documents[0];
  // check if document is saved
  if ((!doc.saved || doc.modified)) {
    if ( confirm ("The document needs to be saved.", undefined, "Document not saved.")) {
      try {
        doc = doc.save();
      } catch (e) {
        alert ("The document couldn't be saved.\n" + e);
        return;
      }
    } else {
      return;
    }
  }
  try{
    drawWindow(doc) == 1
  } catch(e){
    alert(e);
  }
}
/*
 * User Interface
 */
function drawWindow(doc) {
  var myWindow = new Window("palette", panel.title + " " + version, undefined);
  myWindow.orientation = "column";
  myWindow.alignChildren ="fill";
  var panelSelectDir = myWindow.add("panel", undefined, panel.selectDirMenuTitle);
  panelSelectDir.alignChildren = "left";
  var panelSelectDirPath = panelSelectDir.add("edittext");
  panelSelectDirPath.preferredSize.width = 255;
  panelSelectDirPath.text = getDefaultExportPath();
  var panelSelectDirButtonGroup = panelSelectDir.add("group");
  panelSelectDirButtonGroup.alignChildren = "left";
  var panelSelectDirButton = panelSelectDirButtonGroup.add("button", undefined, panel.selectDirButtonTitle);
  // dropdown
  var panelOptions = myWindow.add("panel", undefined, panel.optionsTitle);
  var panelOptionsDropdown = panelOptions.add("dropdownlist", undefined, panel.exportPackageSchemaValues);
  panelOptionsDropdown.selection = 0;
  // buttons OK/Cancel
  panelOptions.alignChildren = "left";
  var panelButtonGroup = myWindow.add("group");
  panelButtonGroup.orientation = "row";
  var buttonOK = panelButtonGroup.add("button", undefined, panel.buttonOK, {name: "ok"});
  var buttonCancel = panelButtonGroup.add("button", undefined, panel.buttonCancel, {name: "cancel"} );
  // change text to selected file path
  panelSelectDirButton.onClick  = function() {
    panelSelectDirPath.text = "hirz"
    var result = Folder.selectDialog ();
    if (result) {
      panelSelectDirPath.text = result;
    }
  }
  buttonOK.onClick = function (){
    //overwrite values with form input
    options.schemaDir = Folder(panelSelectDirPath.text);
    options.exportPackageSchema = (panelOptionsDropdown.selection.index == 0) ? true : false;
    myWindow.close(1);
    generateRNCSchema();
  }
  buttonCancel.onClick = function() {
    myWindow.close();
  }
  return myWindow.show();
}
function generateRNCSchema(){
  app.generateIDMLSchema(Folder(options.schemaDir), options.exportPackageSchema);
  alert(panel.finishedMessage);
}
// get path relative to indesign file location
function getDefaultExportPath() {
  var exportPath = String(app.activeDocument.fullName);
  exportPath = exportPath.substring(0, exportPath.lastIndexOf('/')) + '/' + options.schemaDir;
  return exportPath
}
