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
 * set image export preferences
 */

lang = {
  pre: 1 // en = 0, de = 1
}

image = {
    minExportDPI:1,
    maxExportDPI:2400,
    exportDPI:144,
    maxResolution:4000000,
    pngTransparency:true,
    objectExportOptions:true,
    exportDir:"export",
    exportQuality:2,
    exportFormat:0, // 0 = PNG | 1 = JPG
    pageItemLabel:"letex:fileName"
}
/*
 * set panel preferences
 */
panel = {
    title:["Export Images", "Bilder exportieren"][lang.pre],
    densityTitle:["Density (ppi)", "Auflösung (ppi)"][lang.pre],
    qualityTitle:["Quality", "Qualität"][lang.pre],
    qualityValues:[["max", "high", "medium", "low"], ["Maximum", "Hoch", "Mittel", "Niedrig"]][lang.pre],
    formatTitle:"Format",
    formatValues:["JPG", "PNG"],
    optionsTitle:["Options", "Optionen"][lang.pre],
    objectExportOptionsTitle:["Object export options", "Objektexportoptionen"][lang.pre],
    pngTransparencyTitle:["PNG Transparency", "PNG Transparenz"][lang.pre],
    maxResolutionTitle:["Max Resolution (px)", "Maximale Auflösung (px)"][lang.pre],
    selectDirButtonTitle:["Choose", "Auswählen"][lang.pre],
    selectDirMenuTitle:["Choose a directory", "Verzeichnis auswählen"][lang.pre],
    progressBarTitle:["export Images", "Bilder exportieren"][lang.pre],
    noValidLinks:["No valid links found.", "Keine Bild-Verknüpfungen gefunden"][lang.pre],
    finishedMessage:["images exported.", "Bilder exportiert."][lang.pre],
    buttonOK:"OK",
    buttonCancel:["Cancel", "Abbrechen"][lang.pre]
}

/*
 * start
 */
main();

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
        }
        else {
            return;
        }
    }
    // show window
    try {
        jsExtensions();
        exportImages(doc);
    } catch (e) {
        alert ("Error:\n" + e);
    }
}

function jsExtensions(){
  // indexOf
  if (!Array.prototype.indexOf)
  {
     Array.prototype.indexOf = function(elt /*, from*/)
     {
        var len = this.length;

        var from = Number(arguments[1]) || 0;
        from = (from < 0)
        ? Math.ceil(from)
        : Math.floor(from);

        if (from < 0)
        from += len;

        for (; from < len; from++)
        {
           if (from in this &&
           this[from] === elt)
           return from;
        }
        return -1;
     };
  }
}

function exportImages(doc){
    if(drawWindow()){
        getFilelinks(doc);
    } else {
        return;
    }
}

function drawWindow(){
    var myWindow = new Window("dialog", panel.title, undefined, {resizable:true});
        with (myWindow) {
            myWindow.orientation = "column";
            myWindow.alignChildren ="fill";
            myWindow.formPath = add("panel", undefined, panel.selectDirMenuTitle);
            with(myWindow.formPath){
                myWindow.formPath.alignChildren = "left";
                myWindow.formPath.inputPath = add("edittext");
                myWindow.formPath.inputPath.preferredSize.width = 255;
                myWindow.formPath.inputPath.text = getDefaultExportPath();
                myWindow.formPath.buttonChoosePath = add("button", undefined, panel.selectDirButtonTitle);
            }
            myWindow.qualityGroup = add("panel", undefined, panel.qualityTitle);
            with(myWindow.qualityGroup){
                qualityGroup.orientation = "column";
                qualityGroup.alignChildren = "left";
                qualityGroup.formQuality = add("group");
                with(myWindow.qualityGroup.formQuality){
                    formQuality.orientation = "row";
                    for (i = 0; i < panel.qualityValues.length; i++) {
                        formQuality.add ("radiobutton", undefined, (panel.qualityValues[i]));
                    }
                    formQuality.children[image.exportQuality].value =true;
                }
                myWindow.qualityGroup.formDensity = add( "group");
                with(myWindow.qualityGroup.formDensity){
                    formDensity.inputDPI = add ("edittext", undefined, image.exportDPI);
                    formDensity.inputDPI.characters = 4;
                    formDensity.sliderDPI = add("slider", undefined, image.exportDPI, image.minExportDPI, image.maxExportDPI);
                    formDensity.add ("statictext", undefined, panel.densityTitle);
                    // synchronize slider and input field
                    formDensity.sliderDPI.onChanging = function () {myWindow.qualityGroup.formDensity.inputDPI.text = myWindow.qualityGroup.formDensity.sliderDPI.value};
                    formDensity.inputDPI.onChanging = function () {myWindow.qualityGroup.formDensity.sliderDPI.value = myWindow.qualityGroup.formDensity.inputDPI.text};
                }
            }
            myWindow.optionsGroup = add("panel", undefined, panel.optionsTitle);
            with(myWindow.optionsGroup){
              optionsGroup.formatGroup = add("group");
              optionsGroup.alignChildren = "left";
              with(myWindow.optionsGroup.formatGroup){
                formatGroup.dropdown = add("dropdownlist", undefined, panel.formatValues );
                formatGroup.add ("statictext", undefined, panel.formatTitle);
                formatGroup.dropdown.selection = image.exportFormat;
              }
              myWindow.optionsGroup.maxResolutionGroup = add("group");
              with(myWindow.optionsGroup.maxResolutionGroup){
                maxResolutionGroup.inputMaxRes = maxResolutionGroup.add("edittext", undefined, image.maxResolution);
                maxResolutionGroup.add("statictext", undefined, panel.maxResolutionTitle);
              }
              myWindow.optionsGroup.objectExportOptions = add("group");
              with(myWindow.optionsGroup.objectExportOptions){
                myWindow.optionsGroup.objectExportOptions.checkbox = objectExportOptions.add ("checkbox", undefined, panel.objectExportOptionsTitle);
                myWindow.optionsGroup.objectExportOptions.checkbox.value = image.objectExportOptions;
              }
              myWindow.optionsGroup.pngTransparencyGroup = add("group");
              with(myWindow.optionsGroup.pngTransparencyGroup){
                myWindow.optionsGroup.pngTransparencyGroup.checkbox = pngTransparencyGroup.add ("checkbox", undefined, panel.pngTransparencyTitle);
                myWindow.optionsGroup.pngTransparencyGroup.checkbox.value = image.pngTransparency;
                myWindow.optionsGroup.pngTransparencyGroup.checkbox.enabled = myWindow.optionsGroup.formatGroup.dropdown.selection.text == "PNG" ? true : false;
                // disable checkbox if selected format is not PNG
                myWindow.optionsGroup.formatGroup.dropdown.onChange = function(){
                  myWindow.optionsGroup.pngTransparencyGroup.checkbox.enabled = myWindow.optionsGroup.formatGroup.dropdown.selection.text == "PNG" ? true : false;
                };
              }
            }
            myWindow.buttonGroup = add( "group");
            with(myWindow.buttonGroup){
                myWindow.buttonGroup.orientation = "row";
                myWindow.buttonGroup.buttonOK = add ("button", undefined, panel.buttonOK, {name: "ok"});
                myWindow.buttonGroup.buttonCancel = add ("button", undefined, panel.buttonCancel, {name: "cancel"} );
            }
        }
    myWindow.formPath.buttonChoosePath.onClick  = function () {
        var result = Folder.selectDialog ();
        if (result) {
            myWindow.formPath.inputPath.text = result;

        }
    }
    myWindow.buttonGroup.buttonOK.onClick = function(){
        //overwrite values with form input
        image.exportDir = Folder(myWindow.formPath.inputPath.text);
        image.exportDPI = Number(myWindow.qualityGroup.formDensity.inputDPI.text);
        image.exportQuality = selectedRadiobutton(myWindow.qualityGroup.formQuality);
        image.exportFormat = myWindow.optionsGroup.formatGroup.dropdown.selection.text;
        image.maxResolution = Number(myWindow.optionsGroup.maxResolutionGroup.inputMaxRes.text);
        image.pngTransparency = myWindow.optionsGroup.pngTransparencyGroup.checkbox.value;
        myWindow.close(1);

        function selectedRadiobutton (rbuttons){
            for(var i = 0; i < rbuttons.children.length; i++) {
                if(rbuttons.children[i].value == true) {
                    return i;
                }
            }
        }
    }
    myWindow.buttonGroup.buttonCancel.onClick = function() {
        myWindow.close();
    }
    return myWindow.show();
}


function getFilelinks(doc){
    // overwrite internal Measurement Units
    app.scriptPreferences.measurementUnit = MeasurementUnits.PIXELS;

    var docLinks = doc.links;
    var uniqueBasenames = [];
    var exportLinks = []
    /*
     * rename files if a basename exists twice
     */
    for (var i = 0; i < docLinks.length; i++) {
        var link = docLinks[i];
        var rectangle = link.parent.parent;
        var objectExportOptions = rectangle.objectExportOptions;
        // use format override in objectExportOptions if active
        var overrideBool = objectExportOptions.customImageConversion == true && image.objectExportOptions == true;
        var localFormat = overrideBool ? objectExportOptions.imageConversionType.toString() : image.exportFormat;
        var localDensity = overrideBool ? Number(objectExportOptions.imageExportResolution.toString().replace(/^PPI_/g, "")) : image.exportDPI;
        var normalizedDensity = getMaxDensity(localDensity, rectangle, image.maxResolution);
        var objectExportQualityInt = ["MAXIMUM", "HIGH", "MEDIUM", "LOW" ].indexOf(objectExportOptions.jpegOptionsQuality.toString());
        var localQuality = overrideBool && localFormat != "PNG" ? objectExportQualityInt : image.exportQuality;

        if(isValidLink(link)){
            var basename = getBasename(link.name);
            if(inArray(basename, uniqueBasenames)) {
                var newFilename = renameFile(basename, localFormat, true);
            } else {
                uniqueBasenames.push(basename);
                var newFilename = renameFile(basename, localFormat, false);
            }
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
                customImageConversion:objectExportOptions.customImageConversion,
                objectExportOptions:objectExportOptions
            }
            /*
             * check for custom object export options
             */

            exportLinks.push(linkObject);

        } else {
            var missingLinks = true;
        }
    }
    if (missingLinks) {
        var result = confirm ("Missing Image Links found. Proceed?");
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
        }
        progressBar.close();

        alert (exportLinks.length  + " " + panel.finishedMessage);
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
    switch (link.status) {
        case LinkStatus.LINK_MISSING:
            alert('Error: missing image: ' + link.name); return false; break;
        case LinkStatus.LINK_EMBEDDED:
            alert('Error: embedded image: ' + link.name); return false; break;
        default:
            if(link != null) return true else return false;
    }
}
// return filename with new extension and conditionally attach random string
function renameFile(basename, extension, rename) {
    if(rename) {
        hash = ((1 + Math.random())*0x1000).toString(36).slice(1, 6);
        var renameFile = basename + '-' + hash + '.' + extension.toLowerCase();
        return renameFile
    } else {
        var renameFile = basename + '.' + extension.toLowerCase();
        return renameFile
    }
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

// get path relative to indesign file location
function getDefaultExportPath() {
    var exportPath = String(app.activeDocument.fullName);
    exportPath = exportPath.substring(0, exportPath.lastIndexOf('/')) + '/' + image.exportDir;
    return exportPath
}
