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
image = {
    minExportDPI:1,
    maxExportDPI:2400,
    exportDPI:144,
    exportDir:"export",
    exportQuality:2,
    exportFormat:"JPG",
    pageItemLabel:"letex:fileName"
}
/*
 * set panel preferences 
 */
panel = {
    title:"export images",
    exportPathTitle:"Store images to",
    densityTitle:"Density (ppi)",
    qualityTitle:"Quality",
    qualityValues:["max", "high", "medium", "low"],
    selectDirButtonTitle:"Choose",
    selectDirMenuTitle:"Choose a directory",
    progressBarTitle:"export Images",
    noValidLinks:"No valid links found.",
    buttonOK:"OK",
    buttonCancel:"Cancel"
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
        exportImages(doc);
    } catch (e) {
        alert ("Error:\n" + e);
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
            myWindow.alignChildren ="left";
            myWindow.formPath = add("panel", undefined, panel.exportPathTitle);
            with(myWindow.formPath){
                myWindow.formPath.alignChildren = "left";
                myWindow.formPath.inputPath = add("edittext");
                myWindow.formPath.inputPath.preferredSize.width = 255;
                myWindow.formPath.inputPath.text = getDefaultExportPath();
                myWindow.formPath.buttonChoosePath = add("button", undefined, panel.selectDirButtonTitle);
            }
            myWindow.qualityGroup = add("panel", undefined, panel.qualityTitle);
            with(myWindow.qualityGroup){
                myWindow.qualityGroup.orientation = "column";
                myWindow.qualityGroup.alignChildren = "left";
                myWindow.qualityGroup.formQuality = add("group");
                with(myWindow.qualityGroup.formQuality){
                    myWindow.qualityGroup.formQuality.orientation = "row";
                    for (i = 0; i < panel.qualityValues.length; i++) {
                        myWindow.qualityGroup.formQuality.add ("radiobutton", undefined, (panel.qualityValues[i]));
                    }
                    myWindow.qualityGroup.formQuality.children[image.exportQuality].value =true;
                }
                myWindow.qualityGroup.formDensity = add( "group");
                with(myWindow.qualityGroup.formDensity){
                    myWindow.qualityGroup.formDensity.add ("statictext", undefined, panel.densityTitle);
                    myWindow.qualityGroup.formDensity.inputDPI = add ("edittext", undefined, image.exportDPI);
                    myWindow.qualityGroup.formDensity.inputDPI.characters = 4;
                    myWindow.qualityGroup.formDensity.sliderDPI = add("slider", undefined, image.exportDPI, image.minExportDPI, image.maxExportDPI);
                    // synchronize slider and input field
                    myWindow.qualityGroup.formDensity.sliderDPI.onChanging = function () {myWindow.qualityGroup.formDensity.inputDPI.text = myWindow.qualityGroup.formDensity.sliderDPI.value};
                    myWindow.qualityGroup.formDensity.inputDPI.onChanging = function () {myWindow.qualityGroup.formDensity.sliderDPI.value = myWindow.qualityGroup.formDensity.inputDPI.text};
                }
            }
            myWindow.buttonGroup = add( "group");
            with(myWindow.buttonGroup){
                myWindow.buttonGroup.orientation = "row";
                myWindow.buttonGroup.buttonOK = add ("button", undefined, "OK");
                myWindow.buttonGroup.buttonCancel = add ("button", undefined, "Cancel");
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
    var docLinks = doc.links;
    var uniqueBasenames = [];
    var exportLinks = []    
    /*
     * rename files if a basename exists twice
     */
    for (var i = 0; i < docLinks.length; i++) {
        var link = docLinks[i];
        if(isValidLink(link)){
            var basename = getBasename(link.name);
            if(inArray(basename, uniqueBasenames)) {
                var newFilename = renameFile(basename, image.exportFormat, true);
            } else {
                uniqueBasenames.push(basename);
                var newFilename = renameFile(basename, image.exportFormat, false);
            }
            /*
             * construct link object
             */
            linkObject = {
                link:link,
                pageItem:link.parent.parent,
                newFilename:newFilename,
                newFilepath:File(image.exportDir + "/" + newFilename)
            }

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
        with (app.jpegExportPreferences) {
            antiAlias = false;
            embedColorProfile = false;
            exportResolution= image.exportDPI;
            jpegColorSpace = JpegColorSpaceEnum.RGB;
            jpegExportQualityValues = [JPEGOptionsQuality.MAXIMUM ,JPEGOptionsQuality.HIGH, JPEGOptionsQuality.MEDIUM, JPEGOptionsQuality.LOW];
            jpegQuality = jpegExportQualityValues[image.exportQuality];
            jpegRenderingStyle = JPEGOptionsFormat.BASELINE_ENCODING  
            simulateOverprint = true; 
            useDocumentBleeds = true;
        }

        //var progressBar = getProgressBar(exportLinks.length);
        var progressBar = getProgressBar("export " + exportLinks.length + " images");
        progressBar.reset(exportLinks.length);

        // select export format
        var exportFormat = ExportFormat.JPG
        
        // iterate over files and store to specific location
        for (i = 0; i  < exportLinks.length; i++) {
            
            progressBar.hit("export " + exportLinks[i].newFilename, i);

            exportLinks[i].pageItem.exportFile(exportFormat, exportLinks[i].newFilepath);
            // insert label with new file link for postprocessing
            exportLinks[i].pageItem.insertLabel(image.pageItemLabel, exportLinks[i].newFilename);
        }
        progressBar.close();

        alert (exportLinks.length  + " images exported.");
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
// get path relative to indesign file location
function getDefaultExportPath() {
    var exportPath = String(app.activeDocument.fullName);
    exportPath = exportPath.substring(0, exportPath.lastIndexOf('/')) + '/' + image.exportDir;
    return exportPath
}