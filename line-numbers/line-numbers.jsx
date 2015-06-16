/*
 * line-numbers.jsx
 * 
 * 
 * This script adds line numbers to text frames.
 * 
 * Author: Martin Kraetke (Twitter: @mkraetke)
 *
 * 
 * LICENSE
 * 
 * Copyright (c) 2015, le-tex publishing services GmbH
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
#targetengine "session"

options = {
    debug:1,
    styles:[],
    restartNumPage: 1,
    restartNumFrame: 1,
    ignoreEmptyLines: 1,
    applyToStyle: 0,
    applyTo: 0,
    interval: 1,
    startFrom: 1,
    counter:0,
    pageOffset: app.documents[0].pages[0].documentOffset, // start with 1st page
    lineNumberObjStyleName: "Obj_LineNumber",
    lineNumberParaStyleName: "Para_LineNumber"
}

panel = {
    title:"Insert Line Numbers",
    stylesTitle:"Select Paragraph Styles",
    optionsTitle:"Options",
    applyTitle:"Apply to",
    startTitle:"Start at",
    intervalTitle:"Interval",
    restartNumPageTitle:"Restart numbering each page",
    restartNumFrameTitle:"Restart numbering each frame",
    ignoreEmptyLinesTitle:"Ignore empty lines",
    radioDocumentTitle:"Document",
    radioStoryTitle:"Selected Story",
    radioTextFrameTitle:"Selected Text Frame",
    lineNumberStylesTitle:"Line Number Styles",
    lineNumberObjStyleTitle:"Object Style",
    lineNumberParaStyleTitle:"Paragraph Style",
    buttonAddNumbersTitle:"Add Numbers",
    buttonRemoveNumbersTitle:"Remove all numbers",
    buttonCancelTitle:"Cancel",
    successMessage:" line numbers inserted.",
    removeMessage:" line numbers removed.",
    styleListLength:12
}

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
        drawWindow(doc);
    } catch (e) {
        alert ("Error:\n" + e);
    }
}

function drawWindow (doc) {

    var paraStyles = app.documents[0].allParagraphStyles.sort();                
    var rows = panel.styleListLength;

    var myWindow = new Window ("palette", panel.title, undefined);
    myWindow.orientation = "row";
    myWindow.alignChildren = "fill";
    // select styles panel
    var stylesGroup = myWindow.add ("group {alignChildren: ['', 'fill']}");
    var stylesGroupPanel = stylesGroup.add ("panel", undefined, panel.stylesTitle);
    var stylesGroupScrollbar = stylesGroup.add ("scrollbar", undefined, 0, 0, paraStyles.length-rows);
    stylesGroupScrollbar.preferredSize.width = 10;
    var stylesGroupCheckboxes = stylesGroupPanel.add ("group {orientation: 'column', alignChildren: ['fill', 'top']}");
    for (var i = 0; i < rows; i++){
        if(i < paraStyles.length && paraStyles[i].name != options.lineNumberParaStyleName){
            var checkbox = stylesGroupCheckboxes.add ("checkbox", undefined, paraStyles[i].name);
            checkbox.value = false;
        }
    }
    stylesGroupScrollbar.onChanging = function(){
        var start = Math.round (this.value);
        var stop = start+rows;
        var r = 0;
        for (var i = start; i < stop; i++){
            if(i < paraStyles.length){
                stylesGroupCheckboxes.children[r++].text = paraStyles[i].name;
            }
        }
    }
    // options panel
    var optionsGroup = myWindow.add ("group {orientation: 'column'}");
    var optionsGroupPanel = optionsGroup.add ("panel", undefined, panel.optionsTitle);
    optionsGroupPanel.alignment = "fill";
    optionsGroupPanel.alignChildren = "fill";
    var optionsGroupPanelStartGroup = optionsGroupPanel.add ("group {orientation: 'row'}");
    var optionsGroupPanelInputStart = optionsGroupPanelStartGroup.add ("edittext", undefined, options.startFrom);
    optionsGroupPanelInputStart.characters = 3;
    optionsGroupPanelStartGroup.add ("statictext", undefined, panel.startTitle);    
    var optionsGroupPanelIntervalGroup = optionsGroupPanel.add ("group {orientation: 'row'}");
    var optionsGroupPanelInputInterval = optionsGroupPanelIntervalGroup.add ("edittext", undefined, options.interval);
    optionsGroupPanelInputInterval.characters = 3;
    optionsGroupPanelIntervalGroup.add ("statictext", undefined, panel.intervalTitle);
    var optionsGroupPanelRestartNumPage = optionsGroupPanel.add ("checkbox", undefined, panel.restartNumPageTitle);
    optionsGroupPanelRestartNumPage.value = options.restartNumPage;
    optionsGroupPanelRestartNumPage.enabled = (options.restartNumFrame == 1) ? false : true;
    var optionsGroupPanelRestartNumFrame = optionsGroupPanel.add ("checkbox", undefined, panel.restartNumFrameTitle);
    optionsGroupPanelRestartNumFrame.value = options.restartNumFrame;
    optionsGroupPanelRestartNumFrame.onClick = function(){
        if(optionsGroupPanelRestartNumFrame.value == 1){
            optionsGroupPanelRestartNumPage.enabled = false;
            optionsGroupPanelRestartNumPage.value = 1;
        }else{
            optionsGroupPanelRestartNumPage.enabled = true;
        }
    }
    var optionsGroupPanelIgnoreEmptyLines = optionsGroupPanel.add ("checkbox", undefined, panel.ignoreEmptyLinesTitle);
    optionsGroupPanelIgnoreEmptyLines.value = options.ignoreEmptyLines;
    // panel apply to 
    var optionsGroupApplyToGroup = optionsGroup.add ("panel", undefined, panel.applyTitle);
    optionsGroupApplyToGroup.alignment = "fill";
    optionsGroupApplyToGroup.alignChildren = "left";
    optionsGroupApplyToGroup.add ("radiobutton", undefined, panel.radioDocumentTitle);
    optionsGroupApplyToGroup.add ("radiobutton", undefined, panel.radioStoryTitle);
    optionsGroupApplyToGroup.add ("radiobutton", undefined, panel.radioTextFrameTitle);
    optionsGroupApplyToGroup.children[options.applyTo].value =true;
    // panel style names
    var optionsGroupPanelLineNumberStyles = optionsGroup.add ("panel", undefined, panel.lineNumberStylesTitle);
    optionsGroupPanelLineNumberStyles.orientation = "column";
    optionsGroupPanelLineNumberStyles.alignment = "fill";
    optionsGroupPanelLineNumberStyles.alignChildren = "fill";
    var optionsGroupPanelLineObjStyleGroup = optionsGroupPanelLineNumberStyles.add("group");    
    var lineNumberObjStyleInput = optionsGroupPanelLineObjStyleGroup.add ("edittext", undefined, options.lineNumberObjStyleName);
    var optionsGroupPanelParaStyle = optionsGroupPanelLineObjStyleGroup.add ("statictext", undefined, panel.lineNumberObjStyleTitle);
    lineNumberObjStyleInput.characters = 20;
    var optionsGroupPanelLineParaStyleGroup = optionsGroupPanelLineNumberStyles.add("group");
    var lineNumberParaStyleInput = optionsGroupPanelLineParaStyleGroup.add ("edittext", undefined, options.lineNumberParaStyleName);
    var optionsGroupPanelParaStyle = optionsGroupPanelLineParaStyleGroup.add ("statictext", undefined, panel.lineNumberParaStyleTitle);

    lineNumberParaStyleInput.characters = 20;
    // button group    
    var buttonGroup = myWindow.add ("group {orientation: 'column', alignChildren: 'fill'}");
    var buttonAddNumbers = buttonGroup.add("button", undefined, panel.buttonAddNumbersTitle);
    buttonAddNumbers.enabled = true;
    var buttonRemoveNumbers = buttonGroup.add("button", undefined, panel.buttonRemoveNumbersTitle);
    var buttonCancel = buttonGroup.add("button", undefined, panel.buttonCancelTitle);
    buttonCancel.alignment = ["fill","bottom"];
    buttonAddNumbers.onClick = function(){

        options.styles = selectedInput(stylesGroupCheckboxes);
        options.restartNumPage = optionsGroupPanelRestartNumPage.value;
        options.restartNumFrame = optionsGroupPanelRestartNumFrame.value;
        options.ignoreEmptyLines = optionsGroupPanelIgnoreEmptyLines.value;
        options.startFrom = parseInt(optionsGroupPanelInputStart.text);
        options.interval = parseInt(optionsGroupPanelInputInterval.text);
        options.applyTo = selectedInput(optionsGroupApplyToGroup);
        options.lineNumberObjStyleName = lineNumberObjStyleInput.text;
        options.lineNumberParaStyleName = lineNumberParaStyleInput.text;
        // intialize line numbering
        if(options.styles.length == 0){
            alert("Please select one or more paragraph styles.")
        } else {
            try{
                addNumbers(options);
                alert(options.counter + " line numbers inserted");
            } catch(e){
                alert(e);
            }
        }
    }
    buttonRemoveNumbers.onClick = function(){
        removeNumbers(options);
    }

    buttonCancel.onClick = function() {
        myWindow.close();
    }
    return myWindow.show();
}

function removeNumbers(options){
    var doc = app.activeDocument;
    var stories = doc.stories;
    var counter = 0;
    for (i = 0; i < stories.length; i++){
        var currentStory = stories[i];
        var textFrames = currentStory.textFrames.everyItem().getElements();
        for (j = 0; j < textFrames.length; j++){
            var currentTextFrame = textFrames[j];
            if(currentTextFrame.appliedObjectStyle.name == options.lineNumberObjStyleName){
                currentTextFrame.remove();
                counter = counter +=1
            }
        }
    }
    alert(counter + panel.removeMessage)
    return;
}


function addNumbers(options) {
    var doc = app.activeDocument;
    // startFromTemp needed for saving last line number of frame 
    options.startFromTemp = options.startFrom;

    // create styles, later they are applied
    if(!doc.objectStyles.item(options.lineNumberObjStyleName)){
        var lineNumberObjStyle = doc.objectStyles.add();
        lineNumberObjStyle.name = options.lineNumberObjStyleName;
    }
    if(!doc.paragraphStyles.item(options.lineNumberParaStyleName)){
        var lineNumberParaStyle = doc.paragraphStyles.add();
        lineNumberParaStyle.name = options.lineNumberParaStyleName;
    }
    // add line numbers to entire document
    if(options.applyTo == 0){
        var scope = addNumbersToDocument(doc, options);
    // .. to story
    } else if (options.applyTo == 1){
        var scope = addNumbersToStory(doc, app.selection[0].parentStory, options);
    // .. to text frame
    } else if(app.selection[0]) {
        var scope = addNumbersToTextFrame(doc, app.selection[0], options);
    } else {
        return alert("No Text Frame selected.");
    }
    return;
}

function addNumbersToDocument(doc, options) {
    var stories = doc.stories;
    for (i = 0; i < stories.length; i++){
        currentStory = stories[i];
        addNumbersToStory(doc, currentStory, options);
    }
}

function addNumbersToStory(doc, story, options){
    var textContainers = story.textContainers;
    for (j = 0; j < textContainers.length; j++){
        currentTextFrame = textContainers[j];

        // exclude all frames with a line number object style
        if(currentTextFrame.appliedObjectStyle.name != options.lineNumberObjStyleName){
            addNumbersToTextFrame(doc, currentTextFrame, options);
        }
    }
}

function addNumbersToTextFrame(doc, textFrame, options){
    var currentPageOffset = textFrame.parentPage.documentOffset;
    var pageBreakBoolean = currentPageOffset != options.pageOffset;
    options.pageOffset = currentPageOffset;

    var lines = textFrame.lines;
    var penalty = 0;
    // discontinue line numbering at page breaks when corresponding option is set
    var start = (pageBreakBoolean && options.restartNumPage == 1) ? options.startFrom : options.startFromTemp;
    for (k = 0; k < lines.length; k++){
        var currentLine = lines[k];
        // penalty for empty lines
        penalty = (isLineEmpty(lines[k-1].contents) && options.ignoreEmptyLines == 1) ? penalty +=1 : penalty;
        
        // calculate line number
        var lineNumber = k + start - penalty;
        
        // remember last line if numbering should be contiuned for next frame
        options.startFromTemp = (options.restartNumFrame == 1) ? options.startFromTemp : lineNumber + 1;
        var lineNumberStr = lineNumber.toString();
        var lineParaStyle = currentLine.appliedParagraphStyle.name;
        if(!Boolean(isLineEmpty(currentLine.contents) && options.ignoreEmptyLines == 1) 
          && inArray(lineParaStyle,options.styles)
          && lineNumber % options.interval == 0 ){
            options.counter = options.counter += 1

            var anchoredFrame = currentLine.insertionPoints[0].textFrames.add();

            anchoredFrame.contents = lineNumberStr;
            anchoredFrame.paragraphs[0].appliedParagraphStyle = doc.paragraphStyles.itemByName(options.lineNumberParaStyleName);

            fitTextFrameProportionally(anchoredFrame, 0.01);
            anchoredFrame.fit(FitOptions.FRAME_TO_CONTENT);

            // apply object style
            anchoredFrame.appliedObjectStyle = doc.objectStyles.itemByName(options.lineNumberObjStyleName);

            // anchored object settings, can be overwritten with object style
            var anchoredFrameObjSettings = anchoredFrame.anchoredObjectSettings
            anchoredFrameObjSettings.anchoredPosition = AnchorPosition.ANCHORED;
            anchoredFrameObjSettings.horizontalReferencePoint = AnchoredRelativeTo.TEXT_FRAME;
            anchoredFrameObjSettings.horizontalAlignment = HorizontalAlignment.LEFT_ALIGN;
            anchoredFrameObjSettings.verticalReferencePoint = VerticallyRelativeTo.LINE_BASELINE;
            anchoredFrameObjSettings.verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
            // place relative to spine
            anchoredFrameObjSettings.spineRelative = true;
            anchoredFrameObjSettings.anchorXoffset = 5;
            anchoredFrameObjSettings.anchorYoffset = 0;

            // apply para style
            anchoredFrame.strokeWeight = 0;

        }
    }
    return;
}

// return array from groups with children input elements
function selectedInput (inputGroup) {
    var arr = []
    for(var i = 0; i < inputGroup.children.length; i++) {
        if(inputGroup.children[i].value == true){
            if(inputGroup.children[i].type == "radiobutton"){
               arr.push(i);
            } else {
               arr.push(inputGroup.children[i].text);
            }
        }
    }
    return arr;
}

function inArray(item,array){
    var count=array.length;
    for(var i=0;i<count;i++)
    {
        if(array[i]===item){return true;}
    }
    return false;
}

function isLineEmpty(line){
    if(line.match(/^(\n|\s)*$/)){
        return true;
    } else {
        return false;
    }
}


function resize(textFrame, by) {  
    textFrame.resize(CoordinateSpaces.INNER_COORDINATES, AnchorPoint.TOP_LEFT_ANCHOR, ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY, [by, by]);  
}  
function fitTextFrameProportionally(textFrame, factor) {  
    if (textFrame.overflows) {  
       while (textFrame.overflows) {  
            resize(textFrame, 1 + factor);  
       }  
    }  
    else {  
       while (!textFrame.overflows) {  
            resize(textFrame, 1 - factor);  
       }  
       resize(textFrame, 1 / (1 - factor));  
    }  
}