#targetengine "session"
/*
 * insert-alt-text.jsx
 *
 *
 * Insert alt text from an XML document 
 *
 *
 * Authors: Martin Kraetke (@mkraetke)
 *
 */
jsExtensions();
var version = "v1.2.0";
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
  exportDir:"export",
  altTextXml:"alt-text.xml",
  altLabel:"letex:altText",
  artifactLabel:"letex:artifact",
  filenameLabel:"letex:fileName",
  logFilename:"alt-text.log",
  overrideExistingAltTexts:true
}
/*
 * set panel preferences
 */
panel = {
  title:["le-tex – Insert Alt Text", "le-tex – Alternativtexte einfügen"][lang.pre],
  selectDirButtonTitle:["Choose", "Auswählen"][lang.pre],
  generateAltXMLTitle:["Create XML file", "XML-Datei erstellen"][lang.pre],
  selectDirMenuTitle:["Choose XML file with alt texts", "XML-Datei mit Alternativtexten auswählen"][lang.pre],
  selectDirOpenFileDialogTitle:["Open XML file with alt texts", "XML-Datei mit Alternativtexten öffnen"][lang.pre],
  optionsTitle:["Options", "Optionen"][lang.pre],
  overrideExistingAltTextsTitle:["Override alt texts?", "Alt-Texte überschreiben"][lang.pre],
  buttonOK:"OK",
  buttonCancel:["Cancel", "Abbrechen"][lang.pre],
  lockedLayerWarning:["All layers with images must be unlocked.","Alle Ebenen mit Bildern müssen entsperrt sein."][lang.pre],
  xmlNotFound:["No XML file found", "Keine XML-Datei gefunden!"][lang.pre],
  lockedLayerWarning:["All layers with images must be unlocked.","Alle Ebenen mit Bildern müssen entsperrt sein."][lang.pre],
  generateAltXMLFinishedMsg:["Empty alt text XML file created", "Leere Alt-Text-XML-Datei erstellt"][lang.pre],
  finishedMessage:["Alt texts inserted", "Alt-Texte eingefügt!"][lang.pre]
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
  // show window
  try {
    initializeWindow(doc);
  } catch (e) {
    alert ("Error:\n" + e);
  }
  app.scriptPreferences.userInteractionLevel = userLevel;
}
/*
 * add JavaScript extensions
 */
function jsExtensions(){
  // indexOf
  if (!Array.prototype.indexOf) {
    Array.prototype.indexOf = function(elt /*, from*/)
    {
      var len = this.length;

      var from = Number(arguments[1]) || 0;
      from = (from < 0)
        ? Math.ceil(from)
        : Math.floor(from);

      if (from < 0)
        from += len;

      for (; from < len; from++) {
        if (from in this &&
            this[from] === elt)
          return from;
      }
      return -1;
    };
  }
  // endsWith
  String.prototype.endsWith = function(suffix) {
    return this.indexOf(suffix, this.length - suffix.length) !== -1;
  };
}
function initializeWindow(doc){
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
  //var myWindow = new Window("dialog", panel.title + " " + version, undefined, {resizable:true});
  var myWindow = new Window("palette", panel.title + " " + version, undefined);
  myWindow.orientation = "column";
  myWindow.alignChildren ="fill";
  var panelSelectDir = myWindow.add("panel", undefined, panel.selectDirMenuTitle);
  panelSelectDir.alignChildren = "left";
  var panelSelectDirInputPath = panelSelectDir.add("edittext");
  panelSelectDirInputPath.preferredSize.width = 255;
  panelSelectDirInputPath.text = getDefaultExportPath();
  var panelSelectDirButtonGroup = panelSelectDir.add("group");
  panelSelectDirButtonGroup.alignChildren = "left";
  var panelSelectDirButton = panelSelectDirButtonGroup.add("button", undefined, panel.selectDirButtonTitle);
  var generateAltXMLButton = panelSelectDirButtonGroup.add("button", undefined, panel.generateAltXMLTitle);
  // buttons OK/Cancel
  var panelOptions = myWindow.add("panel", undefined, panel.optionsTitle);
  panelOptions.alignChildren = "left";
  var overrideExistingAltTexts = panelOptions.add("checkbox", undefined, panel.overrideExistingAltTextsTitle);
  overrideExistingAltTexts.value = options.overrideExistingAltTexts;
  var panelButtonGroup = myWindow.add("group");
  panelButtonGroup.orientation = "row";
  var buttonOK = panelButtonGroup.add("button", undefined, panel.buttonOK, {name: "ok"});
  var buttonCancel = panelButtonGroup.add("button", undefined, panel.buttonCancel, {name: "cancel"} );
  // change text to selected file path
  panelSelectDirButton.onClick  = function() {
    var result = Folder(panelSelectDirInputPath.text).openDlg(panel.selectDirOpenFileDialogTitle, "XML:*.xml", true);
    if (result) {
      panelSelectDirInputPath.text = result;
    }
  }
  generateAltXMLButton.onClick  = function() {
    options.altTextXml = File(panelSelectDirInputPath.text + "/alt-text.xml");
    myWindow.close(1);
    generateAltXML(doc);
  }
  buttonOK.onClick = function (){
    //overwrite values with form input
    options.exportDir = Folder(getDefaultExportPath() + '/' + options.exportDir);
    options.altTextXml = File(panelSelectDirInputPath.text);
    options.overrideExistingAltTexts = overrideExistingAltTexts.value;
    myWindow.close(1);
    prepareAltTexts(doc);
  }
  buttonCancel.onClick = function() {
    myWindow.close();
  }
  return myWindow.show();
}
function prepareAltTexts(doc) {
  var xmlExtension = String(options.altTextXml).split(".").pop().toLowerCase()
  if(xmlExtension == 'xml' && File(options.altTextXml).exists){
    // clear log
    clearLog(options.exportDir, options.logFilename);
    var currentDate = new Date();
    var dateTime = currentDate.getDate() + '/'
        + (currentDate.getMonth()+1)  + '/' 
        + currentDate.getFullYear() + ' '  
        + currentDate.getHours() + ':'
        + currentDate.getMinutes() + ':' 
        + currentDate.getSeconds();
    writeLog("le-tex insert-alt-text " + version + "\nstarted at " + dateTime + "\n", options.exportDir, options.logFilename);
    // open XML file
    var xmlFile = File(options.altTextXml);
    xmlFile.open("r");
    var xml = XML(xmlFile.read());
    // iterate over file links
    var docLinks = doc.links;
    var altLinks = [];
    for (var i = 0; i < docLinks.length; i++) {
      var link = docLinks[i];
      writeLog("(" + i + ") --------------------------------------------------------------------------------\n"
               + link.name
               + "\n"
               + link.filePath, options.exportDir, options.logFilename);
      var rectangle = link.parent.parent;
      var filename = link.name;
      var filenameLabel = rectangle.extractLabel(options.filenameLabel);
      var xpath;
      if(String(filenameLabel).length > 0) {
        xpath = '/links/link[@name = \'' + filename + '\' or @name = \'' + filenameLabel + '\']';
      } else {
        xpath = '/links/link[@name = \'' + filename + '\']';
      }
      writeLog(xpath, options.exportDir, options.logFilename);
      var altText = String(xml.xpath(xpath + '/@alt'));
      var artifact = String(xml.xpath(xpath + '/@artifact'));
      if (altText.length != 0) {
        writeLog('alt: ' + altText, options.exportDir, options.logFilename);
        altLinkObject = {
          rectangle:rectangle,
          filename:filename,
          altText:altText
        }
        altLinks.push(altLinkObject);
      } else if (artifact == 'true') {
        writeLog('artifact: ' + artifact, options.exportDir, options.logFilename);
        altLinkObject = {
          rectangle:rectangle,
          filename:filename,
          altText:"",
          artifact:"true"
        }
        altLinks.push(altLinkObject);
      } else {
        writeLog('WARNING: no alt text found!', options.exportDir, options.logFilename);
      }
    }
    var counter = insertAltTexts(altLinks);
    alert (counter  + " " + panel.finishedMessage);
    writeLog("\n===============================================================================================\nFinished! Inserted alt links for " + altLinks.length + " of " + docLinks.length + " images.\nPlease check messages above for further details.", options.exportDir, options.logFilename);
    xmlFile.close();
    doc.save();
  } else {
    alert(panel.xmlNotFound);
  }
}
function insertAltTexts(altLinks){
  var counter = 0;
  for (i = 0; i < altLinks.length; i++) {
    if(options.overrideExistingAltTexts == true || String(altLinks[i].rectangle.objectExportOptions.customAltText).length == 0) {
      altLinks[i].rectangle.insertLabel(options.altLabel, altLinks[i].altText);
      altLinks[i].rectangle.objectExportOptions.altTextSourceType = SourceType.SOURCE_CUSTOM;
      altLinks[i].rectangle.objectExportOptions.customAltText = altLinks[i].altText;
      counter++;
    } else {
      altLinks[i].rectangle.insertLabel(options.altLabel, altLinks[i].rectangle.objectExportOptions.customAltText);
      writeLog('WARNING: alt text found but not overriden!', options.exportDir, options.logFilename);
    }
    if(altLinks[i].artifact == "true") {
      altLinks[i].rectangle.insertLabel(options.altLabel, "");
      altLinks[i].rectangle.insertLabel(options.artifactLabel, altLinks[i].artifact);
      altLinks[i].rectangle.objectExportOptions.applyTagType = TagType.TAG_ARTIFACT;
    }
  }
  return counter;
}
function generateAltXML(doc){
  var myAltXML = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r<links>\r";
  var docLinks = doc.links;
  for (var i = 0; i < docLinks.length; i++) {
    var link = docLinks[i];
    var rectangle = link.parent.parent;
    var filename = link.name;
    var filenameLabel = rectangle.extractLabel(options.filenameLabel);
    var nameAtt = filenameLabel.length > 0 ? filenameLabel : filename;
    var altAtt = String(rectangle.objectExportOptions.customAltText).length > 0 ? rectangle.objectExportOptions.customAltText : "";
    myAltXML += "  <link name=\"" + nameAtt + "\" "
                      + "alt=\"" + altAtt + "\"/>\r";
  }
  myAltXML += "</links>"
  myTargetFile = new File(options.altTextXml);
    myTargetFile.encoding = "UTF-8";
    myTargetFile.open('w');
    myTargetFile.write(myAltXML);
    myTargetFile.close();    
    alert(panel.generateAltXMLFinishedMsg + ": " + options.altTextXml + ".");
}
// get path relative to indesign file location
function getDefaultExportPath() {
  var exportPath = String(app.activeDocument.fullName);
  exportPath = exportPath.substring(0, exportPath.lastIndexOf('/')) + '/';
  return exportPath
}
function createDir (folder) {
  try {
    folder.create();
    return;
  } catch (e) {
    alert (e);
  }
}
// simple logging
function writeLog(message, dir, filename){
  var path = dir + '/' + filename;
  createDir(dir);
  var write_file = File(path);
  if (!write_file.exists) {
    write_file = new File(path);
  }
  d = new Date();
  var timestr = "[" + d.getHours() + ":" + d.getMinutes() + ":" + d.getSeconds() + "] "
  write_file.open('a', undefined, undefined);
  write_file.encoding = "UTF-8";
  write_file.lineFeed = "Unix";
  write_file.writeln(timestr + message);  
  write_file.close();
}
function clearLog(dir, filename){
  var path = dir + '/' + filename;
  del_file = File(path);
  if (del_file.exists) {
    del_file.remove();
  }
}
