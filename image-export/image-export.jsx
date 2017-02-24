#targetengine "session"
/*
 * image-export.jsx
 *
 *
 * Export images from an InDesign document to web-friendly formats.
 *
 *
 * Note: this script requires at least InDesign Version 8.0 (CS6).
 *
 *
 * Authors: Gregor Fellenz (twitter: @grefel), Martin Kraetke (@mkraetke)
 *
 *
 * LICENSE
 *
 * Copyright (c) 2015, Gregor Fellenz and le-tex publishing services GmbH
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * 1. Redistributions of source code must retain the above copyright notice,
 * this list of conditions and the following disclaimer.
 *
 * 2. Redistributions in binary form must reproduce the above copyright notice,
 * this list of conditions and the following disclaimer in the documentation
 * and/or other materials provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
 * ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
 * IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT,
 * INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
 * BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA,
 * OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
 * WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 *
 */

/*
 * set language
 */
lang = {
  pre: 0 // en = 0, de = 1
}
/*
 * image options object
 */
image = {
    minExportDPI:1,
    maxExportDPI:2400,
    exportDPI:144,
    maxResolution:4000000,
    pngTransparency:true,
    objectExportOptions:true,
    objectExportDensityFactor:0,
    overrideExportFilenames:false,
    exportDir:"export",
    exportQuality:2,
    exportFormat:0, // 0 = PNG | 1 = JPG
    pageItemLabel:"letex:fileName",
    logFilename:"export.log"
}
/*
 * image options object
 */
imageInfo = {
    filename:null,
    format:null,
    width:null,
    height:null,
}
/*
 * set panel preferences
 */
panel = {
    title:["Export Images", "Bilder exportieren"][lang.pre],
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
    objectExportDensityFactorTitle:["Resolution Multiplier", "Multiplikator Auflösung"][lang.pre],
    objectExportDensityFactorValues:[1, 2, 3, 4],
    overrideExportFilenamesTitle:["Override embedded export filenames", "Eingebettete Export-Dateinamen überschreiben"][lang.pre],
    pngTransparencyTitle:["PNG Transparency", "PNG Transparenz"][lang.pre],
    maxResolutionTitle:["Max Resolution (px)", "Maximale Auflösung (px)"][lang.pre],
    selectDirButtonTitle:["Choose", "Auswählen"][lang.pre],
    selectDirMenuTitle:["Choose a directory", "Verzeichnis auswählen"][lang.pre],
    panelFilenameOptionsTitle:["Filenames", "Dateinamen"][lang.pre],
    miscellaneousOptionsTitle:["Miscellaneous Options", "Sonstige Optionen"][lang.pre],
    infoFilename:["Filename", "Dateiname"][lang.pre],
    infoWidth:["Width", "Breite"][lang.pre],
    infoHeight:["Height", "Höhe"][lang.pre],
    infoNoImage:["No image selected.", "Kein Bild ausgewählt."][lang.pre],
    progressBarTitle:["export Images", "Bilder exportieren"][lang.pre],
    noValidLinks:["No valid links found.", "Keine Bild-Verknüpfungen gefunden"][lang.pre],
    finishedMessage:["images exported.", "Bilder exportiert."][lang.pre],
    buttonOK:"OK",
    buttonCancel:["Cancel", "Abbrechen"][lang.pre],
    errorPasteboardImage:["Warning! Images on pasteboard will not be exported: ", "Warnung! Bild auf Montagefläche wird nicht exportiert: "][lang.pre],
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
        jsExtensions();
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
}
function exportImages(doc){
    if(drawWindow() == 1){
        // overwrite internal Measurement Units to pixels
        app.scriptPreferences.measurementUnit = MeasurementUnits.PIXELS;
    } else {
        app.removeEventListener(Event.AFTER_SELECTION_CHANGED, function(){}, false);
        return;
    }
    app.removeEventListener(Event.AFTER_SELECTION_CHANGED, function(){}, false);
}
/*
 * User Interface
 */
function drawWindow() {
    //var myWindow = new Window("dialog", panel.title, undefined, {resizable:true});
    var myWindow = new Window("palette", panel.title, undefined);
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
    densitySliderDPI.onChanging = function () {densityInputDPI.text = densitySliderDPI.value};
    densityInputDPI.onChanging = function () {densitySliderDPI.value = densityInputDPI.text};
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
    var objectExportOptionsCheckBox = objectExportOptions.add("checkbox", undefined, panel.objectExportOptionsTitle);
    var objectExportOptionsDensity = panelObjectExportOptions.add("group");
    var objectExportOptionsDensityDropdown = objectExportOptionsDensity.add("dropdownlist", undefined, panel.objectExportDensityFactorValues);
    objectExportOptionsDensityDropdown.selection = image.objectExportDensityFactor;
    var resolutionFactor = objectExportOptionsDensity.add("statictext", undefined, panel.objectExportDensityFactorTitle);
    objectExportOptionsCheckBox.value = image.objectExportOptions;
    var panelFilenameOptions = tabAdvanced.add("panel", undefined, panel.panelFilenameOptionsTitle);
    panelFilenameOptions.alignChildren = "left";
    var overrideExportFilenames = panelFilenameOptions.add("group");
    var overrideExportFilenamesCheckbox = overrideExportFilenames.add("checkbox", undefined, panel.overrideExportFilenamesTitle);
    overrideExportFilenamesCheckbox.value = image.overrideExportFilenames;
    var panelMiscellaneousOptions = tabAdvanced.add("panel", undefined, panel.miscellaneousOptionsTitle)
    panelMiscellaneousOptions.alignChildren = "left";
    var pngTransparencyGroup = panelMiscellaneousOptions.add("group");
    var pngTransparencyGroupCheckbox = pngTransparencyGroup.add("checkbox", undefined, panel.pngTransparencyTitle);
    pngTransparencyGroupCheckbox.value = image.pngTransparency;
    pngTransparencyGroupCheckbox.enabled = formatDropdown.selection.text == "PNG" ? true : false;
    // disable checkbox if selected format is not png
    formatDropdown.onChange = function(){
        pngTransparencyGroupCheckbox.enabled = formatDropdown.selection.text == "PNG" ? true : false;
        formatDropdownDescription.text = formatDropdown.selection.text == "PNG" ? panel.formatDescriptionPNG : panel.formatDescriptionJPEG;
    };
    /*
     * Info Tab
     */
    var tabInfo = tpanel.add ("tab", undefined, panel.tabInfoTitle);
    tabInfo.alignChildren = "left";
    tabInfo.iFilename = tabInfo.add("statictext", undefined, panel.infoNoImage);
    tabInfo.iPosX = tabInfo.add("statictext", undefined, "");
    tabInfo.iPosY = tabInfo.add("statictext", undefined, "");
    tabInfo.iWidth = tabInfo.add("statictext", undefined, "");
    tabInfo.iHeight = tabInfo.add("statictext", undefined, "");
    tabInfo.iPosX.characters = tabInfo.iPosY.characters = tabInfo.iWidth.characters = tabInfo.iHeight.characters = tabInfo.iFilename.characters = 40;
    app.removeEventListener(Event.AFTER_SELECTION_CHANGED, function(){}, false);
    app.addEventListener(Event.AFTER_SELECTION_CHANGED,
                         function(){
                             if(app.selection[0] != undefined && app.selection[0].constructor.name == "Rectangle"){
                                 var rectangle = app.selection[0];
                                 width = Math.round((rectangle.geometricBounds[3] - rectangle.geometricBounds[1])  * 100) / 100;
                                 height = Math.round((rectangle.geometricBounds[2] - rectangle.geometricBounds[0]) * 100) / 100;
                                 posX = Math.round(rectangle.geometricBounds[1] * 100) / 100;
                                 posY = Math.round(rectangle.geometricBounds[0] * 100) / 100;
                                 tabInfo.iFilename.text = panel.infoFilename + ": " + rectangle.extractLabel(image.pageItemLabel);
                                 tabInfo.iPosX.text = "x: " + posX;
                                 tabInfo.iPosY.text = "y: " + posY;
                                 tabInfo.iWidth.text = panel.infoWidth + ": " + width;
                                 tabInfo.iHeight.text = panel.infoHeight + ": " + height;
                             } else {
                                 tabInfo.iFilename.text = panel.infoNoImage;
                                 tabInfo.iPosX.text = "";
                                 tabInfo.iPosY.text = "";
                                 tabInfo.iHeight.text = "";
                                 tabInfo.iWidth.text = "";
                                 app.removeEventListener(Event.AFTER_SELECTION_CHANGED, function(){}, false);
                             }
                         }, false);
    // buttons OK/Cancel
    var panelButtonGroup = myWindow.add("group");
    panelButtonGroup.orientation = "row";
    var buttonOK = panelButtonGroup.add("button", undefined, panel.buttonOK, {name: "ok"});
    var buttonCancel = panelButtonGroup.add("button", undefined, panel.buttonCancel, {name: "cancel"} );
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
        image.objectExportOptions = objectExportOptionsCheckBox.value;
        image.objectExportDensityFactor = objectExportOptionsDensityDropdown.selection.text;
        image.overrideExportFilenames = overrideExportFilenamesCheckbox.value;
        image.pngTransparency = pngTransparencyGroupCheckbox.value;
        app.removeEventListener(Event.AFTER_SELECTION_CHANGED, function(){}, false)
        myWindow.close(1);

        function selectedRadiobutton (rbuttons){
            for(var i = 0; i < rbuttons.children.length; i++) {
                if(rbuttons.children[i].value == true) {
                    return i;
                }
            }
        }
        getFilelinks(app.documents[0]);
    }
    buttonCancel.onClick = function(){
        app.removeEventListener(Event.AFTER_SELECTION_CHANGED, function(){}, false)
        myWindow.close();
    }
    return myWindow.show();
}
function getFilelinks(doc) {
    var docLinks = linksToSortedArray(doc.links);
    var uniqueBasenames = [];
    var exportLinks = [];
    var missingLinks = [];
    // delete filename labels, if option is set
    if(image.overrideExportFilenames == true){
        deleteLabel(docLinks);
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
    writeLog("Image export started at " + dateTime + "\n", image.exportDir, image.logFilename);
    
    // iterate over file links
    for (var i = 0; i < docLinks.length; i++) {
        var link = docLinks[i];
        
        writeLog(link.name, image.exportDir, image.logFilename);
        
        var rectangle = link.parent.parent;
        // disable lock since this prevents images to be exported
        // note that just the group itself has a lock state, not their childs
        if(rectangle.parent.constructor.name == 'Group'){
            if(rectangle.parent.locked != false){
                rectangle.parent.locked = false; 
            }
        } else {
            if(rectangle.locked != false){
                rectangle.locked = false;
            }
        }
        var originalBounds = (rectangle.parentPage != null) ? rectangle.geometricBounds : [0, 0, 0, 0];
        // ignore images in overset text and rectangles with zero width or height 
        if(rectangle.parentPage != null && originalBounds[0] - originalBounds[2] != 0 && originalBounds[1] - originalBounds[3] != 0 ){
            if(rectangle.itemLayer.locked == true) alert(panel.lockedLayerWarning);
            // this is necessary to avoid moving of anchored objects with Y-Offset
            var imageExceedsPageY = rectangle.parentPage.bounds[0] > originalBounds[0];
            if(imageExceedsPageY){
                originalBounds[0] = originalBounds[0] + getAnchoredObjectOffset(rectangle)[0]; //y1
                originalBounds[2] = originalBounds[2] + getAnchoredObjectOffset(rectangle)[0]; //y2
            }
            rectangle = cropRectangleToBleeds(rectangle);
            var objectExportOptions = rectangle.objectExportOptions;
            // use format override in objectExportOptions if active. Check InDesign version because the property changed.
            var customImageConversion = (parseFloat(app.version) < 10) ? objectExportOptions.customImageConversion :
                objectExportOptions.preserveAppearanceFromLayout == PreserveAppearanceFromLayoutEnum.PRESERVE_APPEARANCE_RASTERIZE_CONTENT ||
                objectExportOptions.preserveAppearanceFromLayout == PreserveAppearanceFromLayoutEnum.PRESERVE_APPEARANCE_RASTERIZE_CONTAINER;
            var overrideBool = image.objectExportOptions && customImageConversion;
            var localFormat = overrideBool ? objectExportOptions.imageConversionType.toString() : image.exportFormat;
            var localDensity = overrideBool ? Number(objectExportOptions.imageExportResolution.toString().replace(/^PPI_/g, "")) * image.objectExportDensityFactor : image.exportDPI;
            var normalizedDensity = getMaxDensity(localDensity, rectangle, image.maxResolution);
            var objectExportQualityInt = ["MAXIMUM", "HIGH", "MEDIUM", "LOW" ].indexOf(objectExportOptions.jpegOptionsQuality.toString());
            var localQuality = overrideBool && localFormat != "PNG" ? objectExportQualityInt : image.exportQuality;
            if(isValidLink(link)){
                /*
                 * set export filename
                 */
                var filenameLabel = rectangle.extractLabel(image.pageItemLabel);
                var basename = (filenameLabel.length > 0 && image.overrideExportFilenames == false) ? getBasename(filenameLabel) : getBasename(link.name);
                var newFilename;
                // just adjacent images are compared.
                var duplicates = hasDuplicates(link, docLinks, i);
                if(inArray(basename, uniqueBasenames) && (!duplicates)){
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
                    originalBounds:originalBounds
                }
                /*
                 * check for custom object export options
                 */
                exportLinks.push(linkObject);
                writeLog("=> stored to: " + linkObject.newFilepath, image.exportDir, image.logFilename);
            } else {
                missingLinks.push(link.name);
            }
        } else {
            writeLog("=> FAILED: image is not placed on a page.", image.exportDir, image.logFilename);
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

            exportLinks[i].pageItem.exportFile(exportFormat, exportLinks[i].newFilepath);
            // insert label with new file link for postprocessing
            exportLinks[i].pageItem.insertLabel(image.pageItemLabel, exportLinks[i].newFilename);
            // restore original bounds
            exportLinks[i].pageItem.geometricBounds = exportLinks[i].originalBounds;
        }
        progressBar.close();

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
    if(link.parent.parent.parentPage == null) {
        return false;
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
}
// return filename with new extension and conditionally attach random string
function renameFile(basename, extension, rename) {
    var cleanBasename = cleanString(basename, '_');
    if(rename) {
        hash = ((1 + Math.random())*0x1000).toString(36).slice(1, 6);
        var renameFile = cleanBasename + '-' + hash + '.' + extension.toLowerCase();
        return renameFile
    } else {
        var renameFile = cleanBasename + '.' + extension.toLowerCase();
        return renameFile
    }
}
// clean string from illegal characters
function cleanString(string, replacement){
    var illegalCharsRegex = /[\x00-\x1f\x80-\x9f\s\/\?<>\\:\*\|":]/g;
    var replace = string.replace(/[\x00-\x1f\x80-\x9f\s\/\?<>\\:\*\|":]/g, replacement);
    return replace;
}
// get file basename
function getBasename(filename) {
    var basename = filename.match( /^(.*?)\.[a-z]{2,4}$/i);
    return basename[1];
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
function getMaxDensity(density, rectangle, maxResolution) {
    var bounds = rectangle.geometricBounds;
    var baseMultiplier = 72;
    var densityFactor = density / baseMultiplier;
    var width =  (bounds[3] - bounds[1]);
    var height = (bounds[2] - bounds[0]);
    var resolution = width * height * Math.pow(densityFactor, 2);
    if(resolution > maxResolution) {
        var maxDensity =  Math.floor(Math.sqrt(maxResolution * Math.pow(densityFactor, 2) / resolution) * baseMultiplier);
        return maxDensity;
    } else {
        return density;
    }
}
// crop rectangle to page bleeds
function cropRectangleToBleeds (rectangle){
    document.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;
    var rect = rectangle;
    var bounds = rect.geometricBounds;   // bounds: [y1, x1, y2, x2], e.g. top left / bottom right
    var page = rect.parentPage;
    // page is null if the object is on the pasteboard
    var rulerOrigin = document.viewPreferences.rulerOrigin;
    if(page != null){
        // iterate over corners and fit them into page
        var newBounds = [];
        for(var i = 0; i <= 3; i++) {
            // y1
            if(i == 0 && bounds[i] < page.bounds[i]){
                newBounds[i] = page.bounds[i];
                // x1
            } else if(i == 2 && bounds[i] > page.bounds[i]){
                newBounds[i] = page.bounds[i];
                // left edge, do not crop images which touch the spine
            } else if(i == 1 && bounds[i] < page.bounds[i] && page.side.toString() != "RIGHT_HAND"){
                newBounds[i] = page.bounds[i];
                // right edge
            } else if(i == 3 && bounds[i] > page.bounds[i] && page.side.toString() != "LEFT_HAND"){
                newBounds[i] = page.bounds[i];
            } else {
                newBounds[i] = bounds[i];
            }
        }
        // restore old bounds
        rect.geometricBounds = newBounds;
    }
    document.viewPreferences.rulerOrigin =  rulerOrigin;
    return rect;
}
function getAnchoredObjectOffset (obj){
    if(obj.parent.constructor === Character){
        anchorYoffset = obj.anchoredObjectSettings.anchorYoffset;
        anchorXoffset = obj.anchoredObjectSettings.anchorXoffset;
        return [anchorYoffset, anchorXoffset];
    } else {
        return [0, 0];
    }
}
// get path relative to indesign file location
function getDefaultExportPath() {
    var exportPath = String(app.activeDocument.fullName);
    exportPath = exportPath.substring(0, exportPath.lastIndexOf('/')) + '/' + image.exportDir;
    return exportPath
}
// delete all image file labels 
function deleteLabel(docLinks){
    for (var i = 0; i < docLinks.length; i++) {
        var link = docLinks[i];
        var rectangle = link.parent.parent;
        rectangle.insertLabel(image.pageItemLabel, '');
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
// compare two links if they have equal dimensions
function hasDuplicates(link, docLinks, index) {
    var rectangle = link.parent.parent;
    var nextLink;
    var result = [];
    var i = index + 1;
    do{
        nextLink = docLinks[i];
        i++;
        if(nextLink != undefined && link.name == nextLink.name) {
            var nextRectangle = nextLink.parent.parent;
            // InDesign calculates widths and heights not precisely, so we have to round them
            var rectangleWidth = Math.round((rectangle.geometricBounds[3] - rectangle.geometricBounds[1]) * 100) / 100;
            var rectangleHeight = Math.round((rectangle.geometricBounds[2] - rectangle.geometricBounds[0]) * 100) / 100;
            var nextRectangleWidth = Math.round((nextRectangle.geometricBounds[3] - nextRectangle.geometricBounds[1]) * 100) / 100;
            var nextRectangleHeight = Math.round((nextRectangle.geometricBounds[2] - nextRectangle.geometricBounds[0]) * 100) / 100;        
            var equalWidth  = rectangleWidth == nextRectangleWidth;
            var equalHeight = rectangleHeight == nextRectangleHeight;
            var equalFlip = rectangle.absoluteFlip == nextRectangle.absoluteFlip;
            var equalRotation = rectangle.absoluteRotationAngle == nextRectangle.absoluteRotationAngle;
            result.push(equalFlip && equalRotation && equalWidth && equalHeight);
        }
    }
    while(i < docLinks.length);
    if(inArray(true, result)){
        writeLog("Found duplicate: " + link.name + ". Generate new filename: " + inArray(true, result), image.exportDir, image.logFilename);
    }
    return inArray(true, result);
    return false;
}
