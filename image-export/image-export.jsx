#targetengine "session"
/*
 * image-export.jsx
 *
 *
 * Export images from an InDesign document to web-friendly formats.
 *
 *
 * Note: this script requires at least InDesign Version 8.0 (CS6). For 
 * versions prior to 8.0, please try image-export_pre-cs6.jsx which 
 * comes with a limited feature set.
 *
 * Authors: Gregor Fellenz (twitter: @grefel), Martin Kraetke (@mkraetke)
 *
 */
jsExtensions();
var version = "v1.2.3";
var doc = app.documents[0];
/*
 * set language
 */
var lang = {
  pre: app.locale == 1279477613 ? 1 : 0 // en = 0, de = 1
}
/*
 * image options object
 */
image = {
  minExportDPI:1,
  maxExportDPI:2400,
  exportDPI:parseInt(getConfigValue("letex:exportDPI", "144")[0], 10),
  baseDPI:72,
  maxResolution:parseInt(getConfigValue("letex:maxResolution", "4000000")[0], 10),
  pngTransparency:getConfigValue("letex:pngTransparency", "true")[0] == "true",
  objectExportOptions:getConfigValue("letex:objectExportOptions", "true")[0] == "true",
  objectExportDensityFactorValues:[1, 2, 3, 4],
  objectExportDensityFactor:parseInt(getConfigValue("letex:objectExportDensityFactor", "1")[0], 10) - 1,
  overrideExportFilenames:getConfigValue("letex:overrideExportFilenames", "false")[0] == "true",
  exportFromHiddenLayers:getConfigValue("letex:exportFromHiddenLayers", "false")[0] == "true",
  relinkToExportPaths:getConfigValue("letex:relinkToExportPaths", "false")[0] == "true",
  exportGroupsAsSingleImage:getConfigValue("letex:exportGroupsAsSingleImage", "false")[0] == "true",
  cropImageToPage:getConfigValue("letex:cropImageToPage", "true")[0] == "true",
  removeRectangleStroke:getConfigValue("letex:removeRectangleStroke", "false")[0] == "true",
  exportDir:"export",
  exportQuality:parseInt(getConfigValue("letex:exportQuality", "2")[0], 10),
  exportFormat:["JPG", "PNG"].indexOf(getConfigValue("letex:exportFormat", "JPG")[0]), // 0 = JPG | 1 = PNG
  pageItemLabel:"letex:fileName",
  logFilename:"export.log"
}
function getConfigValue(label, defaultValue) {
  var value = [doc.extractLabel(label), defaultValue].filter( function(text) {
    return text != ""
  })
  return value;
}
/*
 * image options object
 */
imageInfo = {
  filename:null,
  format:null,
  width:null,
  height:null,
  anchorPos:null,
}
/*
 * set panel preferences
 */
panel = {
  title:["le-tex – Export Images", "le-tex – Bilder exportieren"][lang.pre],
  tabGeneralTitle:["General", "Allgemein"][lang.pre],
  tabAdvancedTitle:["Advanced", "Erweitert"][lang.pre],
  tabInfoTitle:["Info", "Informationen"][lang.pre],
  densityTitle:["Density (ppi)", "Auflösung (ppi)"][lang.pre],
  qualityTitle:["Quality", "Qualität"][lang.pre],
  qualityValues:[["max", "high", "medium", "low"], ["Maximum", "Hoch", "Mittel", "Niedrig"]][lang.pre],
  formatTitle:"Format",
  formatValues:["JPG", "PNG"],
  formatDescriptionPNG:["for line art and text", "Für Strichzeichnungen und Text"][lang.pre],
  formatDescriptionJPEG:["for photographs and gradients", "für Fotos und Verläufe"][lang.pre],
  objectExportOptionsTitle:["Object export options", "Objektexportoptionen"][lang.pre],
  objectExportDensityFactorTitle:["Density Multiplier", "Multiplikator Auflösung"][lang.pre],
  overrideExportFilenamesTitle:["Override embedded export filenames", "Eingebettete Export-Dateinamen überschreiben"][lang.pre],
  pngTransparencyTitle:["PNG Transparency", "PNG Transparenz"][lang.pre],
  exportFromHiddenLayersTitle:["Export images from hidden layers", "Bilder von versteckten Ebenen exportieren"][lang.pre],
  relinkToExportPathsTitle:["Relink to export path", "Verknüpfung zu Exportpfad ändern"][lang.pre],
  exportGroupsAsSingleImageTitle:["Export grouped images as single image", "Gruppierte Bilder als einzelnes Bild exportieren"][lang.pre],
  relinkToExportPathsWarning:["Warning! Each link will be replaced with its export path.", "Warnung! Jede Verknüpfung wird durch ihren Exportpfad ersetzt."][lang.pre],
  removeRectangleStrokeTitle:["Remove Stroke", "Rahmen entfernen"][lang.pre],
  maxResolutionTitle:["Max Resolution (px)", "Maximale Auflösung (px)"][lang.pre],
  selectDirButtonTitle:["Choose", "Auswählen"][lang.pre],
  selectDirMenuTitle:["Choose a directory", "Verzeichnis auswählen"][lang.pre],
  panelFilenameOptionsTitle:["Filenames", "Dateinamen"][lang.pre],
  miscellaneousOptionsTitle:["Miscellaneous Options", "Sonstige Optionen"][lang.pre],
  infoCharacters:[40],
  infoFilename:["Filename", "Dateiname"][lang.pre],
  infoWidth:["width", "Breite"][lang.pre],
  infoHeight:["height", "Höhe"][lang.pre],
  cropImageToPageTitle:["Crop image to page margins", "Bild auf Seite beschneiden"][lang.pre],
  infoAnchorPosTitle:["Anchor Position", "Verankerungsposition"][lang.pre],
  infoAnchorPos:[["above line", "anchored", "inline position"], ["Über Zeile", "Benutzerdefiniert", "Eingebunden"]][lang.pre],
  infoNoImage:["No image selected.", "Kein Bild ausgewählt."][lang.pre],
  progressBarTitle:["export Images", "Bilder exportieren"][lang.pre],
  noValidLinks:["No valid links found.", "Keine Bild-Verknüpfungen gefunden"][lang.pre],
  finishedMessage:["images exported.", "Bilder exportiert."][lang.pre],
  buttonOK:"OK",
  buttonCancel:["Cancel", "Abbrechen"][lang.pre],
  buttonSaveSettings:["Save Config", "Konfiguration speichern"][lang.pre],
  errorMissingImage:["Warning! Image cannot be found: ", "Warnung! Bild konnte nicht gefunden werden: "][lang.pre],
  errorEmbeddedImage:["Warning! Embedded Image cannot be exported: ", "Warnung! Eingebettetes Bild kann nicht exportiert werden: "][lang.pre],
  promptMissingImages:["images cannot be exported. Proceed?","Bilder können nicht exportiert werden. Fortfahren?"][lang.pre],
  lockedLayerWarning:["All layers with images must be unlocked.","Alle Ebenen mit Bildern müssen entsperrt sein."][lang.pre]
}
/*
 * start
 */
main();
/*
 * main pipeline
 */
function main(){
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
        var userLevel = app.scriptPreferences.userInteractionLevel;
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
        doc = doc.save();
        app.scriptPreferences.userInteractionLevel = userLevel;
        document.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
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
    // register event listener
    exportImages(doc);
  } catch (e) {
    alert ("Error:\n" + e);
  }
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
  // find
  if (!Array.prototype.find) {
    Array.prototype.find = function(predicate) {
      'use strict';
      if (this == null) {
        throw new TypeError('Array.prototype.find called on null or undefined');
      }
      if (typeof predicate !== 'function') {
        throw new TypeError('predicate must be a function');
      }
      var list = Object(this);
      var length = list.length >>> 0;
      var thisArg = arguments[1];
      var value;

      for (var i = 0; i < length; i++) {
        value = list[i];
        if (predicate.call(thisArg, value, i, list)) {
          return value;
        }
      }
      return undefined;
    };
  }
  // filter
  if (!Array.prototype.filter) {
    Array.prototype.filter = function(func, thisArg) {
      'use strict';
      if ( ! ((typeof func === 'Function' || typeof func === 'function') && this) )
        throw new TypeError();
      
      var len = this.length >>> 0,
          res = new Array(len), // preallocate array
          t = this, c = 0, i = -1;
      if (thisArg === undefined) {
        while (++i !== len){
          // checks to see if the key was set
          if (i in this){
            if (func(t[i], i, t)){
              res[c++] = t[i];
            }
          }
        }
      }
      else{
        while (++i !== len){
          // checks to see if the key was set
          if (i in this){
            if (func.call(thisArg, t[i], i, t)){
              res[c++] = t[i];
            }
          }
        }
      }
      
      res.length = c; // shrink down array to proper size
      return res;
    }
  }
}
function exportImages(doc){
  try{
    drawWindow() == 1
  } catch(e){
    alert(e);
  }
}
/*
 * User Interface
 */
function drawWindow() {
  //var myWindow = new Window("dialog", panel.title, undefined, {resizable:true});
  var myWindow = new Window("palette", panel.title + " " + version, undefined);
  myWindow.orientation = "column";
  myWindow.alignChildren ="fill";
  var tpanel = myWindow.add("tabbedpanel");
  tpanel.alignChildren = "fill";
  /*
   * General Tab
   */
  var tabGeneral = tpanel.add ("tab", undefined, panel.tabGeneralTitle);
  tabGeneral.alignChildren = "fill";
  var panelSelectDir = tabGeneral.add("panel", undefined, panel.selectDirMenuTitle);
  panelSelectDir.alignChildren = "left";
  var panelSelectDirInputPath = panelSelectDir.add("edittext");
  panelSelectDirInputPath.preferredSize.width = 255;
  panelSelectDirInputPath.text = getDefaultExportPath();
  var panelSelectDirButton = panelSelectDir.add("button", undefined, panel.selectDirButtonTitle);  
  var panelQuality = tabGeneral.add("panel", undefined, panel.qualityTitle);
  panelQuality.orientation = "column";
  panelQuality.alignChildren = "left";
  var panelQualityRadiobuttons = panelQuality.add("group");
  panelQualityRadiobuttons.orientation = "row";
  for (i = 0; i < panel.qualityValues.length; i++) {
    panelQualityRadiobuttons.add ("radiobutton", undefined, (panel.qualityValues[i]));
  }
  panelQualityRadiobuttons.children[image.exportQuality].value =true;
  var densityGroup = panelQuality.add("group");  
  var densityInputDPI = densityGroup.add ("edittext", undefined, image.exportDPI);
  densityInputDPI.characters = 4;
  var densitySliderDPI = densityGroup.add("slider", undefined, image.exportDPI, image.minExportDPI, image.maxExportDPI);
  densityGroup.add ("statictext", undefined, panel.densityTitle);
  // synchronize slider and input field
  var outDPI
  densitySliderDPI.onChanging = function () { densityInputDPI.text   = densitySliderDPI.value; };
  densityInputDPI.onChanging  = function () { densitySliderDPI.value = densityInputDPI.text };
  var panelMaxResolutionGroup = panelQuality.add("group");
  panelMaxResolutionGroup.add("statictext", undefined, panel.maxResolutionTitle);  
  var inputMaxRes = panelMaxResolutionGroup.add("edittext", undefined, image.maxResolution);
  var panelFormat = tabGeneral.add("panel", undefined, panel.formatTitle);
  var panelFormatGroup = panelFormat.add("group");
  panelFormat.alignChildren = "left";
  var formatDropdown = panelFormatGroup.add("dropdownlist", undefined, panel.formatValues);
  formatDropdown.selection = image.exportFormat;
  var formatDropdownDescription = panelFormatGroup.add("statictext", undefined, undefined);
  formatDropdownDescription.text = formatDropdown.selection.text == "PNG" ? panel.formatDescriptionPNG : panel.formatDescriptionJPEG;
  /*
   * Advanced Tab
   */
  var tabAdvanced = tpanel.add ("tab", undefined, panel.tabAdvancedTitle);
  tabAdvanced.alignChildren = "fill";
  var panelObjectExportOptions = tabAdvanced.add("panel", undefined, panel.objectExportOptionsTitle);
  panelObjectExportOptions.alignChildren = "left";
  var objectExportOptions = panelObjectExportOptions.add("group");
  var objectExportOptionsCheckbox = objectExportOptions.add("checkbox", undefined, panel.objectExportOptionsTitle);
  objectExportOptionsCheckbox.value = image.objectExportOptions;
  var objectExportOptionsDensity = panelObjectExportOptions.add("group");
  var objectExportOptionsDensityDropdown = objectExportOptionsDensity.add("dropdownlist", undefined, image.objectExportDensityFactorValues);
  objectExportOptionsDensityDropdown.selection = image.objectExportDensityFactor;
  var resolutionFactor = objectExportOptionsDensity.add("statictext", undefined, panel.objectExportDensityFactorTitle);
  var panelFilenameOptions = tabAdvanced.add("panel", undefined, panel.panelFilenameOptionsTitle);
  panelFilenameOptions.alignChildren = "left";
  var overrideExportFilenames = panelFilenameOptions.add("group");
  var overrideExportFilenamesCheckbox = overrideExportFilenames.add("checkbox", undefined, panel.overrideExportFilenamesTitle);
  overrideExportFilenamesCheckbox.value = image.overrideExportFilenames;
  var panelMiscellaneousOptions = tabAdvanced.add("panel", undefined, panel.miscellaneousOptionsTitle)
  panelMiscellaneousOptions.alignChildren = "left";
  var pngTransparencyGroupCheckbox = panelMiscellaneousOptions.add("checkbox", undefined, panel.pngTransparencyTitle);
  pngTransparencyGroupCheckbox.value = image.pngTransparency;
  var pngFormatActive = formatDropdown.selection.text == "PNG" || isFormatOverrideActive("PNG");
  pngTransparencyGroupCheckbox.enabled = pngFormatActive;
  // disable checkbox if selected format is not png
  formatDropdown.onChange = function(){
    pngFormatActive = formatDropdown.selection.text == "PNG" || isFormatOverrideActive("PNG");
    pngTransparencyGroupCheckbox.enabled = pngFormatActive;
    formatDropdownDescription.text = pngFormatActive ? panel.formatDescriptionPNG : panel.formatDescriptionJPEG;
  };
  var cropImageToPageCheckbox = panelMiscellaneousOptions.add("checkbox", undefined, panel.cropImageToPageTitle);
  cropImageToPageCheckbox.value = image.cropImageToPage;
  var removeRectangleStrokeCheckbox = panelMiscellaneousOptions.add("checkbox", undefined, panel.removeRectangleStrokeTitle);
  removeRectangleStrokeCheckbox.value = image.removeRectangleStroke;
  var exportFromHiddenLayersCheckbox = panelMiscellaneousOptions.add("checkbox", undefined, panel.exportFromHiddenLayersTitle);
  exportFromHiddenLayersCheckbox.value = image.exportFromHiddenLayers;
  var relinkToExportPathsCheckbox = panelMiscellaneousOptions.add("checkbox", undefined, panel.relinkToExportPathsTitle);
  relinkToExportPathsCheckbox.value = image.relinkToExportPaths;
  relinkToExportPathsCheckbox.onClick = function(){
    if(relinkToExportPathsCheckbox.value == true) {
      alert(panel.relinkToExportPathsWarning)
    }
  }
  var exportGroupsAsSingleImageCheckbox = panelMiscellaneousOptions.add("checkbox", undefined, panel.exportGroupsAsSingleImageTitle);
  exportGroupsAsSingleImageCheckbox.value = image.exportGroupsAsSingleImage;
  exportGroupsAsSingleImageCheckbox.onClick = function(){
    overrideExportFilenamesCheckbox.value = (exportGroupsAsSingleImageCheckbox.value == true) ? true : overrideExportFilenamesCheckbox.value;    
  }
  /*
   * Info Tab
   */
  var tabInfo = tpanel.add ("tab", undefined, panel.tabInfoTitle);
  tabInfo.alignChildren = "left";
  tabInfo.iFilename = tabInfo.add("statictext", undefined, panel.infoNoImage + "\n\n", {multiline:true});
  tabInfo.iPosX = tabInfo.add("statictext", undefined, "");
  tabInfo.iPosY = tabInfo.add("statictext", undefined, "");
  tabInfo.iWidth = tabInfo.add("statictext", undefined, "");
  tabInfo.iHeight = tabInfo.add("statictext", undefined, "");
  tabInfo.iPosX.characters = tabInfo.iPosY.characters = tabInfo.iWidth.characters = tabInfo.iHeight.characters = tabInfo.iFilename.characters = panel.infoCharacters;
  // listen to each selection change and update info tab
  var afterSelectChanged = app.addEventListener(Event.AFTER_SELECTION_CHANGED, function(){
    
    if(app.selection[0] != undefined && app.selection[0].constructor.name == "Rectangle"){
      var rectangle = app.selection[0];
      outDensity = getMaxDensity(densitySliderDPI.value, rectangle, image.maxResolution, image.baseDPI)
      width = Math.round((rectangle.geometricBounds[3] - rectangle.geometricBounds[1])  * 100 / 100 * outDensity / image.baseDPI);
      height = Math.round((rectangle.geometricBounds[2] - rectangle.geometricBounds[0]) * 100 / 100 * outDensity / image.baseDPI);
      posX = Math.round(rectangle.geometricBounds[1] * 100) / 100;
      posY = Math.round(rectangle.geometricBounds[0] * 100) / 100;
      var fileNameString = splitStringToArray(rectangle.extractLabel(image.pageItemLabel), panel.infoCharacters).join("\n");
      tabInfo.iFilename.text = panel.infoFilename + ":\n" + fileNameString;
      tabInfo.iPosX.text = "x: " + posX + "px";
      tabInfo.iPosY.text = "y: " + posY + "px";
      tabInfo.iWidth.text = panel.infoWidth + ": " + width + "px";
      tabInfo.iHeight.text = panel.infoHeight + ": " + height + "px";
    } else {
      tabInfo.iFilename.text = panel.infoNoImage;
      tabInfo.iPosX.text = "";
      tabInfo.iPosY.text = "";
      tabInfo.iHeight.text = "";
      tabInfo.iWidth.text = "";
    }
  }, false);
  // buttons OK/Cancel
  var panelButtonGroup = myWindow.add("group");
  panelButtonGroup.orientation = "row";
  var buttonOK = panelButtonGroup.add("button", undefined, panel.buttonOK, {name: "ok"});
  var buttonCancel = panelButtonGroup.add("button", undefined, panel.buttonCancel, {name: "cancel"} );
  var buttonSaveSettings = panelButtonGroup.add("button", undefined, panel.buttonSaveSettings );
  // change text to selected file path
  panelSelectDirButton.onClick  = function() {
    var result = Folder.selectDialog ();
    if (result) {
      panelSelectDirInputPath.text = result;
    }
  }
  buttonOK.onClick = function (){
    //overwrite values with form input
    image.exportDir = Folder(panelSelectDirInputPath.text);
    image.exportDPI = Number(densityInputDPI.text);
    image.exportQuality = selectedRadiobutton(panelQualityRadiobuttons);
    image.exportFormat = formatDropdown.selection.text;
    image.maxResolution = Number(inputMaxRes.text);
    image.objectExportOptions = objectExportOptionsCheckbox.value;
    image.objectExportDensityFactor = objectExportOptionsDensityDropdown.selection.text;
    image.overrideExportFilenames = overrideExportFilenamesCheckbox.value;
    image.pngTransparency = pngTransparencyGroupCheckbox.value;
    image.cropImageToPage = cropImageToPageCheckbox.value;
    image.exportFromHiddenLayers = exportFromHiddenLayersCheckbox.value;
    image.relinkToExportPaths = relinkToExportPathsCheckbox.value;
    image.exportGroupsAsSingleImage = exportGroupsAsSingleImageCheckbox.value;
    image.removeRectangleStroke = removeRectangleStrokeCheckbox.value;
    myWindow.close(1);
    afterSelectChanged.remove();
    getFilelinks(app.documents[0]);
  }
  buttonCancel.onClick = function() {
    afterSelectChanged.remove();
    myWindow.close();
  }
  buttonSaveSettings.onClick = function() {
    var doc = app.documents[0];
    doc.insertLabel("letex:exportDPI", densityInputDPI.text);
    doc.insertLabel("letex:exportQuality", String(selectedRadiobutton(panelQualityRadiobuttons)));
    doc.insertLabel("letex:exportFormat", formatDropdown.selection.text);
    doc.insertLabel("letex:maxResolution", inputMaxRes.text);
    doc.insertLabel("letex:objectExportOptions", String(objectExportOptionsCheckbox.value));
    doc.insertLabel("letex:objectExportDensityFactor", objectExportOptionsDensityDropdown.selection.text);
    doc.insertLabel("letex:overrideExportFilenames", String(overrideExportFilenamesCheckbox.value));
    doc.insertLabel("letex:pngTransparency", String(pngTransparencyGroupCheckbox.value));
    doc.insertLabel("letex:cropImageToPage", String(cropImageToPageCheckbox.value));
    doc.insertLabel("letex:exportFromHiddenLayers", String(exportFromHiddenLayersCheckbox.value));
    doc.insertLabel("letex:relinkToExportPaths", String(relinkToExportPathsCheckbox.value));
    doc.insertLabel("letex:exportGroupsAsSingleImage", String(exportGroupsAsSingleImageCheckbox.value));
    doc.insertLabel("letex:removeRectangleStroke", String(removeRectangleStrokeCheckbox.value));
    afterSelectChanged.remove();
    myWindow.close();
  }
  return myWindow.show();
}
function selectedRadiobutton (rbuttons){
  for(var i = 0; i < rbuttons.children.length; i++) {
    if(rbuttons.children[i].value == true) {
      return i;
    }
  }
}
function getFilelinks(doc) {
  app.scriptPreferences.measurementUnit = MeasurementUnits.PIXELS;
  var docLinks = linksToSortedArray(doc.links);
  var uniqueBasenames = [];
  var exportLinks = [];
  var missingLinks = [];
  var imageGroupIds = [];
  var imageGroupIterator = 0;
  // delete filename labels, if option is set
  if(image.overrideExportFilenames == true){
    deleteLabel(doc);
  }
  // clear log
  clearLog(image.exportDir, image.logFilename);
  var currentDate = new Date();
  var dateTime = currentDate.getDate() + '/'
      + (currentDate.getMonth()+1)  + '/' 
      + currentDate.getFullYear() + ' '  
      + currentDate.getHours() + ':'
      + currentDate.getMinutes() + ':' 
      + currentDate.getSeconds();
  writeLog("le-tex image-export " + version + "\nstarted at " + dateTime + "\n", image.exportDir, image.logFilename);
  
  // iterate over file links
  for (var i = 0; i < docLinks.length; i++) {
    var link = docLinks[i];
    writeLog("\n" + link.name + "\n" + link.filePath, image.exportDir, image.logFilename);
    if(isValidLink(link)){
      var rectangle = link.parent.parent;
      var linkname = link.name;
      // if a group should be exported as single image, replace rectangle with group object
      if(rectangle.parent.constructor.name == "Group" && image.exportGroupsAsSingleImage){
        // use always the filename of the first graphic to avoid duplicates
        linkname = rectangle.parent.rectangles[0].graphics[0].itemLink.name;
        rectangle = getTopmostGroup(rectangle);
      }
      rectangle = disableLocks(rectangle);
      parentPage = (rectangle.parentPage != null) ? rectangle.parentPage.name : "null";
      // restore the frame of anchored objects which overlaps the page
      // after running cropRectangleToPage()
      var originalBounds = rectangle.geometricBounds;
      var rotationAngle = rectangle.rotationAngle;
      var shearAngle = rectangle.absoluteShearAngle;
      var anchored = rectangle.parent.constructor.name == "Character";
      var anchorXoffset = (anchored) ? rectangle.anchoredObjectSettings.anchorXoffset : null;
      var anchorYoffset = (anchored) ? rectangle.anchoredObjectSettings.anchorYoffset : null;
      var textWrapMode = rectangle.textWrapPreferences.textWrapMode;
      writeLog(  "page: " + parentPage
                 + "\nshear angle: "    + shearAngle
                 + "\nrotation angle: " + rotationAngle
                 + "\ny1: " + originalBounds[0]
                 + ", x1: " + originalBounds[1]
                 + ", y2: " + originalBounds[2]
                 + ", x2: " + originalBounds[3]
                 + "\nanchor offsets: " + "x: " + anchorXoffset + ", y: " + anchorYoffset
                 + "\ntext wrap mode: "  + textWrapMode
                 , image.exportDir, image.logFilename);
      var exportFromHiddenLayers = rectangle.itemLayer.visible ? true : image.exportFromHiddenLayers;
      // offsets y1, x1, y2, x2 (top, left, bottom, right)
      var boundOffsets = [originalBounds[0] - rectangle.parentPage.bounds[0], 
                          originalBounds[1] - rectangle.parentPage.bounds[1], 
                          originalBounds[2] - rectangle.parentPage.bounds[2],
                          originalBounds[3] - rectangle.parentPage.bounds[3]];
      var absoluteAccuracy = 5; // in px
      var exceedsPage = boundOffsets[0] < (0 - absoluteAccuracy)
                     || boundOffsets[1] < (0 - absoluteAccuracy)
                     || boundOffsets[2] > absoluteAccuracy
                     || boundOffsets[3] > absoluteAccuracy
      // ignore images in overset text and rectangles with zero width or height
      if(exportFromHiddenLayers
         && originalBounds[0] - originalBounds[2] != 0
         && originalBounds[1] - originalBounds[3] != 0
        ){
        if(rectangle.itemLayer.locked == true) alert(panel.lockedLayerWarning);
        /* since cropping works not well with anchored images, we
         * create a duplicate of the rectangle where the cropping is applied
         */ 
        var rectangleCopy = null;
        if((image.cropImageToPage && exceedsPage)
           || (image.removeRectangleStroke && (rectangle.strokeWeight > 0))
          ){
          // disable text wrap temporarily, otherwise duplicate will be suppressed
          rectangle.textWrapPreferences.textWrapMode = 1852796517 // NONE
          // create duplicate of image
          rectangleCopy = rectangle.duplicate( [originalBounds[0], originalBounds[1]] , [0, 0] );
          // copy rotation angle
          rectangleCopy.rotationAngle = rectangle.rotationAngle;
          rectangleCopy = cropRectangleToPage(rectangleCopy);
          rectangle.textWrapPreferences.textWrapMode = textWrapMode;
          // setting strokeTint to 0 is more robust. Setting strokeWeight to 0 adds sometimes a stroke
          rectangleCopy.strokeTint = image.removeRectangleStroke ? 0 : rectangle.strokeTint;
        }
        var objectExportOptions = rectangle.objectExportOptions;
        // use format override in objectExportOptions if active. Check InDesign version because the property changed.
        var customImageConversion = isObjectExportOptionActive(objectExportOptions);
        var overrideBool = image.objectExportOptions && customImageConversion;
        var localFormat = overrideBool ? objectExportOptions.imageConversionType.toString().replace(/^JPEG/g, "JPG") : image.exportFormat;
        var localDensity = overrideBool ? Number(objectExportOptions.imageExportResolution.toString().replace(/^PPI_/g, "")) * image.objectExportDensityFactor : image.exportDPI;
        var normalizedDensity = getMaxDensity(localDensity, rectangle, image.maxResolution, image.baseDPI);
        var objectExportQualityInt = ["MAXIMUM", "HIGH", "MEDIUM", "LOW" ].indexOf(objectExportOptions.jpegOptionsQuality.toString());
        var localQuality = overrideBool && localFormat != "PNG" ? objectExportQualityInt : image.exportQuality;
        /*
         * set export filename
         */
        var filenameLabel = rectangle.extractLabel(image.pageItemLabel);
        var basename = (filenameLabel.length > 0 && image.overrideExportFilenames == false) ? getBasename(filenameLabel) : getBasename(linkname);
        var newFilename;
        var duplicates = hasDuplicates(link, docLinks, i);
        if( rectangle.constructor.name == "Group" ){
          newFilename = (filenameLabel.length > 0 && image.overrideExportFilenames == false) ? renameFile(basename, localFormat, false) : renameFile(getBasename(linkname) + "_group", localFormat, false);
       } else if(inArray(basename, uniqueBasenames) && (!duplicates)){
          newFilename = renameFile(basename, localFormat, true);
        } else {
          newFilename = renameFile(basename, localFormat, false);
        }
        uniqueBasenames.push(getBasename(newFilename));
        /*
         * construct link object
         */
        linkObject = {
          link:link,
          pageItem:rectangle,
          format:localFormat,
          quality:localQuality,
          density:normalizedDensity,
          newFilename:newFilename,
          newFilepath:File(image.exportDir + "/" + newFilename),
          objectExportOptions:objectExportOptions,
          anchored:anchored,
          rectangleCopy:rectangleCopy,
          exceedsPage:exceedsPage,
          originalBounds:originalBounds,
          group:rectangle.constructor.name == "Group",
          id:rectangle.id
        }
        exportLinks.push(linkObject);
        writeLog("=> stored to: " + linkObject.newFilepath, image.exportDir, image.logFilename);
      } else {
        missingLinks.push(linkname);
      }
    }
  }
  if (missingLinks.length > 0) {
    var result = confirm (missingLinks.length + " " + panel.promptMissingImages + "\n\n" + missingLinks.toString());
    if (!result) return;
  }
  // create directory
  createDir(image.exportDir);

  if (exportLinks.length  > 0) {
    //var progressBar = getProgressBar(exportLinks.length);
    var progressBar = getProgressBar("export " + exportLinks.length + " images");
    progressBar.reset(exportLinks.length);

    // iterate over files and store to specific location
    for (i = 0; i  < exportLinks.length; i++) {
      var exportFormat = exportLinks[i].format == "PNG" ? ExportFormat.PNG_FORMAT : ExportFormat.JPG;
      var exportResolution = exportLinks[i].density;
      var exportQuality = exportLinks[i].quality;
      // JPEG export options
      app.jpegExportPreferences.antiAlias = false;
      app.jpegExportPreferences.embedColorProfile = false;
      app.jpegExportPreferences.exportResolution = exportResolution;
      app.jpegExportPreferences.jpegColorSpace = JpegColorSpaceEnum.RGB;
      jpegExportQualityValues = [JPEGOptionsQuality.MAXIMUM ,JPEGOptionsQuality.HIGH, JPEGOptionsQuality.MEDIUM, JPEGOptionsQuality.LOW];
      app.jpegExportPreferences.jpegQuality = jpegExportQualityValues[exportQuality];
      app.jpegExportPreferences.jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING;
      app.jpegExportPreferences.simulateOverprint = true;
      app.jpegExportPreferences.useDocumentBleeds = true;
      // PNG export options
      app.pngExportPreferences.antiAlias = false;
      app.pngExportPreferences.exportResolution = exportResolution;
      app.pngExportPreferences.pngColorSpace = PNGColorSpaceEnum.RGB;
      pngExportQualityValues = [PNGQualityEnum.MAXIMUM, PNGQualityEnum.HIGH, PNGQualityEnum.MEDIUM, PNGQualityEnum.LOW];
      app.pngExportPreferences.pngQuality = pngExportQualityValues[exportQuality];
      app.pngExportPreferences.transparentBackground = image.pngTransparency;
      app.pngExportPreferences.simulateOverprint = true;
      app.pngExportPreferences.useDocumentBleeds = true;

      progressBar.hit("export " + exportLinks[i].newFilename, i);
      if((image.cropImageToPage && exportLinks[i].exceedsPage)
         || (image.removeRectangleStroke && (rectangle.strokeWeight > 0))
        ){
        exportLinks[i].rectangleCopy.exportFile(exportFormat, exportLinks[i].newFilepath);
        exportLinks[i].rectangleCopy.remove();
      } else {
        exportLinks[i].pageItem.exportFile(exportFormat, exportLinks[i].newFilepath);
      }
      // insert label with new file link for postprocessing
      exportLinks[i].pageItem.insertLabel(image.pageItemLabel, exportLinks[i].newFilename);
    }
    progressBar.close();

    /*
     * danger zone: relink all images to their respective export paths
     */
    if(image.relinkToExportPaths == true) {
      relinkToExportPaths(doc, exportLinks);
    }
    alert (exportLinks.length  + " " + panel.finishedMessage);
    writeLog("\nFinished! Exported " + exportLinks.length + " of " + docLinks.length + " images.\nPlease check messages above for further details.", image.exportDir, image.logFilename);
    doc.save();
  }
  else {
    alert (panel.noValidLinks);
  }
}
// draw simple progress bar
function getProgressBar (title){
  var progressBarWindow = new Window("palette", title, {x:0, y:0, width:400, height:50});
  var pbar = progressBarWindow.add("progressbar", {x:10, y:10, width:380, height:6}, 0, 100);
  var stext = progressBarWindow.add("statictext", {x:10, y:26, width:380, height:20}, '');
  progressBarWindow.center();
  progressBarWindow.reset = function (maxValue) {
    stext.text = "export export export export export export export export ";
    pbar.value = 0;
    pbar.maxvalue = maxValue||0;
    pbar.visible = !!maxValue;
    this.show();
  }
  progressBarWindow.hit = function(msg, value) {
    ++pbar.value;
    stext.text = msg;
  }
  return progressBarWindow;
}
// create directory
function createDir (folder) {
  try {
    folder.create();
    return;
  } catch (e) {
    alert (e);
  }
}
// check if image is missing or embedded
function isValidLink (link) {
  var rectangle = link.parent.parent;
  try {
    // script would crash when geometricBounds not available, e.g. image is placed on overset text
    var bounds = rectangle.geometricBounds;
    if(rectangle.hasOwnProperty("parentPage") && rectangle.parentPage == null){
      writeLog('=> FAILED: image is on pasteboard', image.exportDir, image.logFilename);
      return false;
    } else if(link.parent.constructor.name == 'Story'){
      writeLog('=> WARNING: text-only link found: ' + link.name, image.exportDir, image.logFilename);
      return false;
    } else if(rectangle.parent.constructor.name == "Group" && image.overrideExportFilenames == true) {
      writeLog('=> INFO: part of image group.', image.exportDir, image.logFilename);
      return true;
    } else {
      switch (link.status) {
      case LinkStatus.LINK_MISSING:
        writeLog('=> FAILED: image file is missing.', image.exportDir, image.logFilename);
        return false; break;
      case LinkStatus.LINK_EMBEDDED:
        writeLog('=> FAILED: embedded image.', image.exportDir, image.logFilename);
        return false; break;
      default:
        if(link != null) return true else return false;
      }
    }
  } catch (e) {
    writeLog('=> FAILED: image is placed in overset text', image.exportDir, image.logFilename);
    return false;
  }
}
// return filename with new extension and conditionally attach random string
function renameFile(basename, extension, rename) {
  var normalizedBasename = basename.replace(/[%\x00-\x1f\x80-\x9f\s\/\?<>\\:\*\|":]/g, '_');
  if(rename) {
    hash = ((1 + Math.random())*0x1000).toString(36).slice(1, 6);
    var renameFile = normalizedBasename + '-' + hash + '.' + extension.toLowerCase();
    return renameFile
  } else {
    var renameFile = normalizedBasename + '.' + extension.toLowerCase();
    return renameFile
  }
}
// get file basename
function getBasename(filename) {
  var basename = filename.match( /^(.*?)\.[a-z]{2,4}$/i);
  if(basename != null){
    return basename[1];
  } else {
    // no file extension
    return filename;
  }
}
// check if string exists in array
function inArray(string, array) {
  var length = array.length;
  for(var i = 0; i < length; i++) {
    if(array[i] == string)
      return true;
  }
  return false;
}
// get density limit according to maximal resolution value
function getMaxDensity(density, rectangle, maxResolution, baseDensity) {
  var bounds = rectangle.geometricBounds;
  var densityFactor = density / baseDensity;
  var width =  (bounds[3] - bounds[1]);
  var height = (bounds[2] - bounds[0]);
  var resolution = Math.round(width * height * Math.pow(densityFactor, 2));
  if(resolution > maxResolution) {
    var maxDensity =  Math.floor(Math.sqrt(maxResolution * Math.pow(densityFactor, 2) / resolution) * baseDensity);
    return maxDensity;
  } else {
    return density;
  }
}
// crop a rectangle to page bleeds
function cropRectangleToPage (rectangle){
  var bounds = rectangle.geometricBounds;   // bounds: [y1, x1, y2, x2], e.g. top left / bottom right
  var page = rectangle.parentPage;
  // release anchors to avoid displaced images. we need to restore the anchor later
  if(rectangle.parent.constructor.name == "Character"){
    rectangle.anchoredObjectSettings.releaseAnchoredObject();
  }  
  // page is null if the object is on the pasteboard
  if(page != null){
    // rectangle.geometricBounds = [bounds[0], bounds[1], bounds[2], bounds[3]];
    // iterate over each corner and fit them into page
    var newBounds = [];
    for(var i = 0; i <= 3; i++) {
      // y1 (top-left)
      if(i == 0 && bounds[i] < page.bounds[i]){
        newBounds[i] = page.bounds[i];
      // y2 (bottom-right)
      } else if(i == 2 && bounds[i] > page.bounds[i]){
        newBounds[i] = page.bounds[i];
      // x1 (top-left)
      } else if(i == 1 && bounds[i] < page.bounds[i] && page.side.toString() != "RIGHT_HAND"){
        newBounds[i] = page.bounds[i];
      // x2 (bottom-right)
      } else if(i == 3 && bounds[i] > page.bounds[i] && page.side.toString() != "LEFT_HAND"){
        newBounds[i] = page.bounds[i];
      } else {
        newBounds[i] = bounds[i];
      }
    }
    // assign new bounds
    rectangle.geometricBounds = newBounds;
  }
  return rectangle;
}
// get anchor position, needed for cropRectangleToPage()
function getAnchoredPosition(rectangle){
  // 1095716961: AnchorPosition.ABOVE_LINE
  // 1097814113: AnchorPosition.ANCHORED
  // 1095716969: AnchorPosition.INLINE_POSITION
  anchoredPosIndex = [1095716961, 1097814113, 1095716969].indexOf(rectangle.anchoredObjectSettings.anchoredPosition);
  return anchoredPosIndex;
}

// get path relative to indesign file location
function getDefaultExportPath() {
  var exportPath = String(app.activeDocument.fullName);
  exportPath = exportPath.substring(0, exportPath.lastIndexOf('/')) + '/' + image.exportDir;
  return exportPath
}
// delete all image file labels 
function deleteLabel(doc){  
  for (var i = 0; i < doc.links.length; i++) {
    var link = doc.links[i];
    var rectangle = link.parent.parent;
    rectangle.insertLabel(image.pageItemLabel, '');
  }
  for (var i = 0; i < doc.groups.length; i++) {
    var group = doc.groups[i];
    group.insertLabel(image.pageItemLabel, '');
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
  write_file.open('a', undefined, undefined);
  write_file.encoding = "UTF-8";
  write_file.lineFeed = "Unix";
  write_file.writeln(message);
  write_file.close();
}
function clearLog(dir, filename){
  var path = dir + '/' + filename;
  del_file = File(path);
  if (del_file.exists) {
    del_file.remove();
  }
}
// create array from links object
function linksToSortedArray(links){
  arr = [];
  // put links in an array
  for (var i = 0; i < links.length; i++) {
    arr.push(links[i]);
  }
  // sort the array by name
  arr.sort(function(a, b){
    var x = a.name.toLowerCase();
    var y = b.name.toLowerCase();
    return x.localeCompare(y);
  });
  return arr;
}
function isObjectExportOptionActive(objectExportOptions){
  var active = (parseFloat(app.version) <= 10) ? objectExportOptions.customImageConversion :
      objectExportOptions.preserveAppearanceFromLayout == PreserveAppearanceFromLayoutEnum.PRESERVE_APPEARANCE_RASTERIZE_CONTENT ||
      objectExportOptions.preserveAppearanceFromLayout == PreserveAppearanceFromLayoutEnum.PRESERVE_APPEARANCE_RASTERIZE_CONTAINER;
  return active;
}
function isFormatOverrideActive(searchFormat){
  var result = false;
  var links = app.documents[0].links;
  for(var i = 0; i < links.length; i++){
    var rectangle = links[i].parent.parent;
    if(rectangle.hasOwnProperty("objectExportOptions")
       && isObjectExportOptionActive(rectangle.objectExportOptions)
       && rectangle.objectExportOptions.imageConversionType.toString() == searchFormat
      ){
      result = true;
    }
  }
  return result;
}


// compare two links if they have equal dimensions
function hasDuplicates(link, docLinks, index) {
  var rectangle = link.parent.parent;
  var nextLink;
  var result = [];
  var i = 0;
  do{
    nextLink = docLinks[i];
    // check whether link names match. The index var is used to prevent that an image is compared with itself.
    if( link.name == nextLink.name && i != index && isValidLink(nextLink)) {
      var nextRectangle = nextLink.parent.parent;
      // InDesign calculates widths and heights not precisely, so we have to round them
      var rectangleWidth = Math.round((rectangle.geometricBounds[3] - rectangle.geometricBounds[1]) * 100) / 100;
      var rectangleHeight = Math.round((rectangle.geometricBounds[2] - rectangle.geometricBounds[0]) * 100) / 100;
      var nextRectangleWidth = Math.round((nextRectangle.geometricBounds[3] - nextRectangle.geometricBounds[1]) * 100) / 100;
      var nextRectangleHeight = Math.round((nextRectangle.geometricBounds[2] - nextRectangle.geometricBounds[0]) * 100) / 100;
      var equalWidth  = rectangleWidth == nextRectangleWidth;
      var equalHeight = rectangleHeight == nextRectangleHeight;
      var equalFlip = rectangle.absoluteFlip == nextRectangle.absoluteFlip;
      var equalRotationAngle = rectangle.absoluteRotationAngle == nextRectangle.absoluteRotationAngle;
      var equalShearAngle = rectangle.absoluteShearAngle == nextRectangle.absoluteShearAngle;
      var equalHorizontalScale = rectangle.absoluteHorizontalScale == nextRectangle.absoluteHorizontalScale;
      var equalVerticalScale = rectangle.absoluteVerticalScale == nextRectangle.absoluteVerticalScale;
      var inGroup = rectangle.parent.constructor.name == "Group";
      // note: either objectExportOptions are not active, then we safely ignore them or we
      // check if they are active for the two images
      var objectExportOptionsActive = !image.objectExportOptions || isObjectExportOptionActive(rectangle.objectExportOptions) == isObjectExportOptionActive(nextRectangle.objectExportOptions);
      result.push(equalFlip && equalRotationAngle && equalWidth && equalHeight && equalShearAngle && equalHorizontalScale && equalVerticalScale && inGroup && objectExportOptionsActive);
    }
    i++;
  }
  while(i < docLinks.length);
  if(inArray(true, result)){
    writeLog("Found duplicate: " + link.name, image.exportDir, image.logFilename);
  }
  return inArray(true, result);
}
function relinkToExportPaths (doc, exportLinks) {
  for(var i = 0; i < exportLinks.length; i++) {
    var linkId = exportLinks[i].link.id;
    var exportPath = exportLinks[i].newFilepath;
    var link = doc.links.itemByID(linkId);
    var rectangle = link.parent.parent;
    if(exportLinks[i].group){
      var x = exportLinks[i].pageItem.geometricBounds[1];
      var y = exportLinks[i].pageItem.geometricBounds[0];
      var group =  doc.groups.itemByID(exportLinks[i].id);
      var spread = group.parent;
      var image = spread.place(new File(exportPath), [x,y], doc.layers[0]);
      group.remove();
    } else {
      // relink to export path
      link.relink(exportPath);
      // fit content to frame, necessary because export crops, flips, etc
      rectangle.fit(FitOptions.CONTENT_TO_FRAME);
    }
  }
}
function splitStringToArray (string, index) {
  var tokens = [];
  for(var i = 0; i < string.length; i++){
    tokens.push(string.substring(i,  i + index - 1));
    i = i + index - 2;
  }
  return tokens;
}
function getTopmostGroup(rectangle){
  var p = rectangle.parent;
  while(p.parent.constructor.name == "Group"){
    p = p.parent;
  }
  return p;
}
// disable lock since this prevents images to be exported
// note that just the group itself has a lock state, not their children
function disableLocks(rectangle){
  if(rectangle.parent.constructor.name == "Group"
     && rectangle.parent.hasOwnProperty("locked")
     && rectangle.parent.locked != false){
    rectangle.parent.locked = false;
  }
  if(rectangle.itemLayer.hasOwnProperty("locked")
     && rectangle.itemLayer.locked != false){
    rectangle.itemLayer.locked = false;
  }
  if(rectangle.hasOwnProperty("locked")
     && rectangle.locked != false){
    rectangle.locked = false;
  }
  return rectangle;
}
