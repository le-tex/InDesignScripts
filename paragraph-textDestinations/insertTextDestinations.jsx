#target indesign

//insertTextDestinations.jsx
//author: Anna Schmalfuß, le-tex publishing services GmbH
//version: 1.0
//date: 2016-05-31

//checks for pre-existing text destinations
//creates a text destination at the beginning of each paragraph

var counter = 0;

if(app.documents.length != 0) {
    var doc = app.activeDocument;
    for(var i = 0; i < doc.stories.length; i++){
        var story = doc.stories[i];
        counter = insertTextDest(story, i, counter);
     }
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
                doc.hyperlinkTextDestinations.add(para.insertionPoints[0], {name: String(Math.random()).substr(2, 6) + "_" + storyID});
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