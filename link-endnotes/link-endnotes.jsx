#target indesign

//LinkEndnotes.jsx
//author: Anna Schmalfuß, le-tex publishing services GmbH
//version: 1.0
//date: 2015-08-27

//version 1.1
//modified 2015-08-31: create backlinks
//modified 2016-05-23: exactly match style names + specified conditions to process paras

/*
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


options = {
    endnoteTitleStyleName: "Endnote_U1",
    endnoteParaStyleName: "Endnote",
    chapterStyleName: "U1",
    endnoteNumberStyleName: "Endnote_Ziffer",
    endnoteRefStyleName: "Endnote_Ref",
    endnotePrefixLabel: "en_",
    run: 0
    }

var linkCount = 0;
var endnoteNrDestArray = [];
var endnoteRefDestArray = [];


if(app.documents.length != 0) {
    
    var doc = app.activeDocument;
    main(doc);
    
}	

else{
	alert("Error! \r\rNo active document!");
}





// ######################  functions 

function main(doc){
   if ((!doc.saved || doc.modified)) {
        if ( confirm ("The document needs to be saved.", undefined, "Document not saved.")) {
            try {
                var userLevel = app.scriptPreferences.userInteractionLevel;
                app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
                doc = doc.save();
                app.scriptPreferences.userInteractionLevel = userLevel;
            }
            catch (e) {
                alert ("The document couldn't be saved.\n" + e);
                return;
            }   
        }
        else {
            return;
        }
    }

    var safetyAlert = confirm("This script deletes any hyperlink and hyperlink\ndestination with the name prefix '" + options.endnotePrefixLabel + "'. \rDo you want to run the script?");
    if (safetyAlert == true){
        
        //remove hyperlinks with prefix en_
        for (var i = doc.hyperlinks.length - 1; i >= 0; i-- ) {
            var link = doc.hyperlinks[i];
            var source = link.source;
             if(link.name.match(/^en_/g) ){
                link.remove();        
                source.remove();
            }
        }
    
        // remove destinations with prefix en_
        for (var i = doc.hyperlinkTextDestinations.length - 1; i >= 0; i--) {
            var dest = doc.hyperlinkTextDestinations[i]
            if(dest.name.match(/^en_/g)){
               dest.remove();
               }
        }

        // draw UI
        try {
            ui(doc);
        } 
        catch (e) {
            alert (e);
        }

        //run script
       if(options.run != 0){
            run (options);
       }
    }
}


function ui(doc){
    var dialog = new Window ("dialog", "Link endnotes");
    dialog.alignChildren = "left";
        
    var pstyles = doc.allParagraphStyles;
    var pStylesArray = [];
    for (a = 0; a < pstyles.length; a ++){
        pStylesArray.push (pstyles[a].name);
        }
    
    var cstyles = doc.allCharacterStyles;
    var cStylesArray = [];
    for (a = 0; a < cstyles.length; a ++){
        cStylesArray.push (cstyles[a].name);
        }
    
    var panel1 = dialog.add( "panel", undefined, "Paragraph styles for sectionwise splitting")
    panel1.preferredSize.width = 450;
    panel1.orientation = 'column';
    panel1.spacing = 10;
    panel1.alignChildren = "left";
    
    var row1 = panel1.add("group");
    row1.add("statictext", undefined, "Headline style in main text");
    var list1 = row1.add ("dropdownlist", undefined, pStylesArray);
    list1.selection = 0;
    
     var row2 = panel1.add("group");
    row2.add("statictext", undefined, "Headline style in appendix");
    var list2 = row2.add ("dropdownlist", undefined, pStylesArray);
    list2.selection = 0;
    
    var panel2 = dialog.add( "panel", undefined, "Endnote styles")
    panel2.preferredSize.width = 450;
    panel2.orientation = 'column';
    panel2.spacing = 10;
    panel2.alignChildren = "left";
    
    var row3 = panel2.add("group");
    row3.add("statictext", undefined, "Paragraph style of endnotes");
    var list3 = row3.add ("dropdownlist", undefined, pStylesArray);
    list3.preferredSize.width = 175;
    list3.selection = 0;
    
    var row4 = panel2.add("group");
    row4.add("statictext", undefined, "Character style of endnote reference");
    var list4 = row4.add ("dropdownlist", undefined, cStylesArray);
    list4.preferredSize.width = 175;
    list4.selection = 0;
    
    var row5 = panel2.add("group");
    row5.add("statictext", undefined, "Character style of endnote number");
    var list5 = row5.add ("dropdownlist", undefined, cStylesArray);
    list5.preferredSize.width = 175;
    list5.selection = 0;
     
    var buttons = dialog.add ("group")
    var ok = buttons.add ("button", undefined, "OK", {name: "ok"});
    var esc = buttons.add ("button", undefined, "Cancel", {name: "esc"});
    
    ok.onClick = function() {
        if (list1.selection === null || list2.selection === null || list3.selection === null || list4.selection === null || list5.selection === null){
            alert("Error while selecting paragraph and character styles.");
       }
        else {
           options.chapterStyleName = list1.selection.text.replace(/(~.+)?/g, "");
           options.endnoteTitleStyleName = list2.selection.text.replace(/(~.+)?/g, "");
           options.endnoteParaStyleName = list3.selection.text.replace(/(~.+)?/g, "");
           options.endnoteRefStyleName = list4.selection.text.replace(/(~.+)?/g, "");
           options.endnoteNumberStyleName = list5.selection.text.replace(/(~.+)?/g, "");
           checkSelection(options);
           options.run = 1;
           dialog.close();
       };
    }
    return dialog.show();
}


function checkSelection(options){
    if(options.chapterStyleName == options.endnoteTitleStyleName){
        alert("Section and endnote headline styles are identical. This may cause incorrect linking.");
    }     
    if(options.endnoteRefStyleName == options.endnoteNumberStyleName){
        alert("Endote reference and number styles are identical. This may cause incorrect linking.");
    }
    return;
}


function run(options){
    // collect styles with identical base name (without tilde)
    options.endnoteParaStylePattern = eval("/^" + options.endnoteParaStyleName + "(~.+)?$/g");
    options.chapterStylePattern = eval("/^" + options.chapterStyleName + "(~.+)?$/g");
    options.endnoteTitleStylePattern = eval("/^" + options.endnoteTitleStyleName + "(~.+)?$/g");
    
    // progress bar
    var msgTitle = {de:"Schritt 1/2: Find and create link destinations"};
    var progress = new Window('palette', localize(msgTitle) );
    progress.frameLocation = [400,275];
    progress.bar = progress.add('progressbar');
    progress.bar.size = [400,15];
    progress.bar.value = 0;
    progress.bar.maxvalue = doc.allPageItems.length;
    progress.file = progress.add('statictext');
    progress.file.size = [400,15];
    progress.show();
    
    for(var i = 0; i < doc.stories.length; i++){
        var story = doc.stories[i];
        
        for (var k = 0; k < story.paragraphs.length; k++){
            var para = story.paragraphs[k];
        
                if(para.appliedParagraphStyle.name.match(options.endnoteTitleStylePattern) && para.contents.match(/\S/g)){
                // Whitespace und Umbruchzeichen entfernen (000A = harter Umbruch, 00AD = bedingter Zeilenumbruch, 200B = bedingter Zeilenumbruch ohne Trennstrich, FEFF = Tagklammern)            
                var endnotesTitle = para.contents.replace(/\s|[\u000A\u00AD\u200B\uFEFF]/g, "");
                var endnotesTitleForProgress = para.contents.replace(/[\u000A\u00AD\u200B\uFEFF]/g, "");
                //alert("endnotesTitle: " + endnotesTitle);
            }
                
                
                if(para.appliedParagraphStyle.name.match(options.endnoteParaStylePattern) && endnotesTitle){
                    endnoteNrDestArray = findLinkdestinations(para, options.endnoteNumberStyleName, options.endnotePrefixLabel, "a", endnoteNrDestArray, endnotesTitle);
                    progress.bar.value++;
                    progress.file.text = endnotesTitleForProgress;
                }            
        
       
                 if(para.appliedParagraphStyle.name.match(options.chapterStylePattern) && para.contents.match(/\S/g)){
                     var chapterTitle = para.contents.replace(/\s|[\u000A\u00AD\u200B\uFEFF]/g, ""); 
                     var chapterTitleForProgress = para.contents.replace(/[\u000A\u00AD\u200B\uFEFF]/g, "");
                  }                
               
                if(!para.appliedParagraphStyle.name.match(options.chapterStylePattern) && 
                    !para.appliedParagraphStyle.name.match(options.endnoteParaStylePattern) && 
                    !para.appliedParagraphStyle.name.match(options.endnoteTitleStylePattern) && 
                    chapterTitle && para.contents.match(/\S/g)){
                    //alert("Kapiteltitel: " + chapterTitle);
                    endnoteRefDestArray = findLinkdestinations(para, options.endnoteRefStyleName, options.endnotePrefixLabel, "b", endnoteRefDestArray, chapterTitle);
                    progress.bar.value++;
                    progress.file.text = chapterTitleForProgress;
                }            
         
       }
    }
   progress.close();
   
  
     // progress bar
    var msgTitle = {de:"Schritt 2/2: Find and link references"};
    var progress = new Window('palette', localize(msgTitle) );
    progress.frameLocation = [400,275];
    progress.bar = progress.add('progressbar');
    progress.bar.size = [400,15];
    progress.bar.value = 0;
    progress.bar.maxvalue = doc.allPageItems.length;
    progress.file = progress.add('statictext');
    progress.file.size = [400,15];
    progress.show();
    
     for(var i = 0; i < doc.stories.length; i++){
        var story = doc.stories[i]
        
        for (var k = 0; k < story.paragraphs.length; k++){
            var para = story.paragraphs[k];
            
            
                if(para.appliedParagraphStyle.name.match(options.chapterStylePattern) && para.contents.match(/\S/g)){
                        var chapterTitle = para.contents.replace(/\s|[\u000A\u00AD\u200B\uFEFF]/g, ""); 
                        var chapterTitleForProgress = para.contents.replace(/[\u000A\u00AD\u200B\uFEFF]/g, "");
                 }

                //alert("Kapiteltitel: " + chapterTitle);
                if(!para.appliedParagraphStyle.name.match(options.chapterStylePattern) && 
                    !para.appliedParagraphStyle.name.match(options.endnoteParaStylePattern) && 
                    !para.appliedParagraphStyle.name.match(options.endnoteTitleStylePattern) && 
                    chapterTitle && para.contents.match(/\S/g)){
                    linkCount = createLinks(para, options.endnoteRefStyleName, options.endnotePrefixLabel, "a", chapterTitle, linkCount);
                    progress.bar.value++;
                    progress.file.text = chapterTitleForProgress;
                    }
        
            if(para.appliedParagraphStyle.name.match(options.endnoteTitleStylePattern) && para.contents.match(/\S/g)){
            // Whitespace und Umbruchzeichen entfernen (000A = harter Umbruch, 00AD = bedingter Zeilenumbruch, 200B = bedingter Zeilenumbruch ohne Trennstrich, FEFF = Tagklammern)            
            var endnotesTitle = para.contents.replace(/\s|[\u000A\u00AD\u200B\uFEFF]/g, "");
            var endnotesTitleForProgress = para.contents.replace(/[\u000A\u00AD\u200B\uFEFF]/g, "");
            }
                
            //alert("endnotesTitle: " + endnotesTitle);
            if(para.appliedParagraphStyle.name.match(options.endnoteParaStylePattern) && endnotesTitle){
               linkCount = createLinks(para, options.endnoteNumberStyleName, options.endnotePrefixLabel, "b", endnotesTitle, linkCount);
               progress.bar.value++;
               progress.file.text = endnotesTitleForProgress;
            }

                
      }
    
    }  
progress.close();

alert(linkCount + " hyperlinks have been created.")
}



function findLinkdestinations (para, grepStyleName, prefixLabel, id,  destArray, title){
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        app.findChangeGrepOptions.properties = {includeFootnotes:true, includeMasterPages:false, includeHiddenLayers:true, includeLockedStoriesForFind:true, includeLockedLayersForFind:true};
        app.findGrepPreferences.findWhat = "\\d+";
        var foundItems = new Array();
        foundItems = para.findGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        var foundItemsArray = [];
        for (var i = 0; i < foundItems.length; i++ ) {
            var pattern = eval("/^" + grepStyleName + "/g");
            var charStyleCheck = foundItems[i].appliedCharacterStyle.name.match(pattern);
           if(charStyleCheck) {
                  var content = foundItems[i].contents;
				var ip = foundItems[i].insertionPoints[-1];
				foundItemsArray.push( {INDEX: i, Content: content, IP: ip, Title:title, source:foundItems[i].texts } );
			}
            }
        
	for (var i = foundItemsArray.length - 1; i >= 0; i--) {
        var label = prefixLabel + foundItemsArray[i].Content + id + "_"+ title;
		try {
			var linkDestination = doc.hyperlinkTextDestinations.item(label);
			linkDestination.name;
		}
		catch(e) {
			var linkDestination = doc.hyperlinkTextDestinations.add(foundItemsArray[i].IP, {name: label});
 /*  
foundItemsArray[i].IP.select();
app.activeWindow.zoomPercentage = app.activeWindow.zoomPercentage;
alert(foundItemsArray[i].ID + "\n\n" + e); 
  */
		}
		destArray.push( foundItemsArray[i] );
	}

return destArray;
}


function createLinks(para, grepStyleName, prefixLabel, id, title, counter){
       app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        app.findChangeGrepOptions.properties = {includeFootnotes:true, includeMasterPages:false, includeHiddenLayers:true, includeLockedStoriesForFind:true, includeLockedLayersForFind:true};
        app.findGrepPreferences.findWhat = "\\d+";
        var foundItems = new Array();
        foundItems = para.findGrep();
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
       
       for (var i = 0; i < foundItems.length; i++ ) {
            var pattern = eval("/^" + grepStyleName + "/g");
            var charStyleCheck = foundItems[i].appliedCharacterStyle.name.match(pattern);
           if(charStyleCheck) {
               
                  var linkDest = doc.hyperlinkTextDestinations.item(prefixLabel + foundItems[i].contents + id + "_"+ title);
                
                  if(linkDest != null){
                      try{
                    //alert("linkname: " + prefixLabel + foundItems[i].contents + id + "_"+ title);
                    var linkSource = doc.hyperlinkTextSources.add(foundItems[i].texts);
                    var linkName = prefixLabel+ foundItems[i].contents + id + "_"+ title.slice(0,20);
                    doc.hyperlinks.add(linkSource, linkDest, {name: linkName, visible:false});
                    counter++;
                    }
                    catch(e) {
                        alert(e);
                    }
                    
                  }
                  else{
                        alert("Error! \n\nThe hyperlink destination \"" + prefixLabel + foundItems[i].contents + id + "_" + title.slice(0,20) + "\" was not found. \nNo link will be created.");
                  }
                   
            }
       }
   return counter;
}
