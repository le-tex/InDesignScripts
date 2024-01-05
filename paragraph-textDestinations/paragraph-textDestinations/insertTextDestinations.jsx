﻿#target indesign

//insertTextDestinations.jsx
//author: Anna Schmalfuß, le-tex publishing services GmbH
//version: 1.0
//date: 2016-05-31

//checks for pre-existing text destinations
//creates a text destination at the beginning of each paragraph

var counter = 0;

if(app.documents.length != 0) {
    var doc = app.activeDocument;
    var progress = new Window('palette', localize("Creating text destinations") );
    progress.frameLocation = [400,275];
    progress.bar = progress.add('progressbar');
    progress.bar.size = [400,15];
    progress.bar.value = 0;
    progress.bar.maxvalue = doc.stories.length;
    progress.file = progress.add('statictext');
    progress.file.size = [400,15];
    progress.show();
    for(var i = 0; i < doc.stories.length; i++){
        progress.bar.value++;
        progress.file.text = "Story " + i + "/" + doc.stories.length;
        
        var story = doc.stories[i];
        counter = insertTextDest(story, i, counter);
     }
    progress.close();
    alert("Number of created text destinations: " + counter);
    
}	

else{
	alert("Error! \r\rNo active document!");
}



// ********************** functions **********************

function insertTextDest(text, storyID, destCount){
    var lengths = text.paragraphs.everyItem().length;
    for (var j = lengths.length - 1; j >= 0; j--) {
        var para = text.paragraphs[j];
        if (lengths[j] != 1 && checkForDest(para, storyID)) {
            try{
                var name = String(Math.random()).substr(2, 7) + lengths[j] + storyID;
                name = name.slice(-8);
                doc.hyperlinkTextDestinations.add(para.insertionPoints[0], {name: name + "_" + storyID});
                destCount ++;
            }
            catch(e){
                var zoom = app.activeWindow.zoomPercentage;
                app.activeWindow.zoomPercentage = zoom;
                alert(e);
             }
         }
    }
    return destCount;
}


function checkForDest(para, storyID){
    var check = true;
    for (var k = doc.hyperlinkTextDestinations.length - 1; k >= 0; k--) {
        var dest = doc.hyperlinkTextDestinations[k];
        if(dest.destinationText.index == para.insertionPoints[0].index && dest.name.match("_" + storyID)) check = false
    }
    return check;
}