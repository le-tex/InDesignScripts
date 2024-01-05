#target indesign

//
// PageNamesToStory_Notes.jsx
// Version 1.3
//
// created by: maqui  |  10:18 28.02.2013
// modified by: maqui |  16:59 28.02.2013 (Löschen der Bedingungstexte ohne Einblenden --> Formatierung bleibt 1:1 erhalten)
// modified by: maqui |  16:40 16.06.2014 Alle Textabschnitte erhalten die Seitenzahl-Infos
// modified by: maqui |  11:00 21.05.2015 Debug (bedingten Text "PageStart" und "PageEnd" löschen, auch wenn er sichtbar ist)
// modified by: gimsieke |  2015-06-07 Add 'CellPage_XY' to table cells
// modified by: aschmalfuss | 2018-02-26 replace conditional text with notes
//
// Trägt in jedem Textrahmen (bei jedem Seitenwechsel) den Seitenbeginn bzw. das Seitenende als Notiz mit der jeweiligen Seitenzahl ein.
//
//

var myPageStartTag = "PageStart";
var myPageEndTag = "PageEnd";

//var myStartSwatchName = "PageStart";
//var myEndSwatchName = "PageEnd";

main(); // ==>

function main(){
	if(app.documents.length != 0){
		var myDoc = app.activeDocument;
		var myPage;

		var myStartCounter = 0;
		var myEndCounter = 0;
		for (var i = 0; i < myDoc.stories.length; i++) {
			var myStory = myDoc.stories[i];
            myStory.notes.everyItem().remove();
			var myMasterSpreadFlag = true;
			for (var j = 0; j < myStory.textContainers.length; j++) {
				var myTestFrame = myStory.textContainers[j];
				// Wenn ein Textrahmen nicht auf der Montagefläche ...
				if(myTestFrame.hasOwnProperty("parentPage") && myTestFrame.parentPage != null) {
					// ... überprüfen, ob auf Musterseite
					if(myTestFrame.parentPage.parent.constructor.name != "MasterSpread" && myStory.textContainers.length > 0) {
						myMasterSpreadFlag = false;
						break;
					}
				}
			}
			// Wenn Textabschnitt sich nicht auf einer Musterseite befindet, ...
			if(!myMasterSpreadFlag && myStory.textContainers.length > 0) {
				// Alle Textrahmen vom Textabschnitt durchlaufen ...
				var myFrame = myStory.textContainers[myStory.textContainers.length - 1];
				var myOldPage = false; // wenn nicht als "false" definiert, ist myOldPage letzter "Page"-Wert???????????????????????
				var myOldFirstIP = false;
				while (myFrame != null) {
					var myPage = myFrame.parentPage;
					// Wenn Textabschnitt sich nicht auf Montagefläche ...
					if(myPage != null && (myFrame.constructor.name == "TextFrame" 
                                                     || myFrame.constructor.name == "EndnoteTextFrame") && myFrame.insertionPoints.length > 0) {
						// PageEnd
						// Ist Rahmen auf anderer Seite? ... Wenn ja, dann PageEnd-Tag setzen
						if(myPage != myOldPage) {
							var myLastContent = myPageEndTag + "_" + myPage.name;
							if(myOldFirstIP) {
								//var myLastIP = myOldFirstIP.parentStory.insertionPoints.previousItem(myOldFirstIP);
								var myLastIP = myOldFirstIP.parentStory.insertionPoints[myOldID - 1];
								//var myLastID = myLastIP.parentStory.insertionPoints.itemByRange(0, myLastIP.index).insertionPoints.length;
								//alert(myOldID + "\n" + myLastID);
							}
							else {
								var myLastIP = myFrame.insertionPoints.lastItem();
//myFrame.select();
//app.activeWindow.zoomPercentage = app.activeWindow.zoomPercentage;
//alert(myLastContent + " :: " + myOldPage.constructor.name);
							}
							//myLastIP.fillColor = myEndSwatch;
							var note = myLastIP.notes.add();
                                note.label=myPageEndTag;
                               note.texts[0].contents=myLastContent;
							myEndCounter++;
						}
						// PageStart
						// Ist vorheriger Rahmen auf anderer Seite? ... Wenn ja, dann PageStart-Tag setzen
						if (checkPrevFrame(myFrame, myPage) && myFrame.contents != "") {
							var myFirstContent = myPageStartTag + "_" + myPage.name;
							var myFirstIP = myFrame.insertionPoints.firstItem();
							var myOldFirstIP = myFirstIP; // IP merken, da er sich bei evtl. Umbruchveränderungen verschiebt
							var myOldID = myOldFirstIP.parentStory.insertionPoints.itemByRange(0, myOldFirstIP.index).insertionPoints.length;
							var note = myFirstIP.notes.add();
                                note.label=myPageStartTag;
                               note.texts[0].contents=myFirstContent;
							myStartCounter++;
						}
						myOldPage = myPage;
					}
					var myFrame = myFrame.previousTextFrame;
				}
			}

      // Process all table cells in the story since they won't be included in the frame iteration:
      if(!myMasterSpreadFlag && myStory.tables.length > 0) {
        for (var t = 0; t < myStory.tables.length; t++) {
		 myStory.tables[t].notes.everyItem().remove();
          for (var c = 0; c < myStory.tables[t].cells.length; c++) {
            var myCell = myStory.tables[t].cells[c];
            var firstCellChar = myCell.characters[0];
            if (firstCellChar != null) {
              var myFrame = firstCellChar.parentTextFrames[0];
              if (myFrame != null) {
     					  var myPage = myFrame.parentPage;
					      var myIP = myCell.insertionPoints.firstItem();
					      //myIP.fillColor = myStartSwatch;
					     var note = myIP.notes.add();
                            note.label="PageName";
                            note.texts[0].contents="CellPage_" + myPage.name;
              }
            }
          }
        }
			}
		}
		myDoc.select(NothingEnum.NOTHING);
		//alert("FERTIG! ;-) \r\rVergebene Startseiten-Label: " + myStartCounter + "\rVergebene Endseiten-Label:   " + myEndCounter);
	}
	else{
		alert("FEHLER! \r\rEs ist kein Dokument geöffnet.");
	}
}

function checkPrevFrame(thisFrame, thisPage) {
	if(thisFrame.previousTextFrame != null) {
		var thisPrevFrame = thisFrame.previousTextFrame;
		// Wenn Rahmen auf Montagefläche, vorherigen überprüfen, u.s.w.
		while (thisPrevFrame.parentPage == null) {
/* thisPrevFrame.select();
app.activeWindow.zoomPercentage = app.activeWindow.zoomPercentage;
alert("Rahmen auf Montagefläche!"); */
			if(thisPrevFrame.previousTextFrame != null) {
				var thisPrevFrame = thisPrevFrame.previousTextFrame;
			}
			else return true;
		}
		if (thisPrevFrame.parentPage != thisPage) return true;
		else return false;
	}
	else return true;
}

