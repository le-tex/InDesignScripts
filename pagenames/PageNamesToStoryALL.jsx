#target indesign

//
// PageNamesToStoryALL.jsx
// Version 1.3
//
// created by: maqui  |  10:18 28.02.2013
// modified by: maqui |  16:59 28.02.2013 (Löschen der Bedingungstexte ohne Einblenden --> Formatierung bleibt 1:1 erhalten)
// modified by: maqui |  16:40 16.06.2014 Alle Textabschnitte erhalten die Seitenzahl-Infos
// modified by: maqui |  11:00 21.05.2015 Debug (bedingten Text "PageStart" und "PageEnd" löschen, auch wenn er sichtbar ist)
// modified by: gimsieke |  2015-06-07 Add 'CellPage_XY' to table cells 
// 
// Trägt in jedem Textrahmen (bei jedem Seitenwechsel) den Seitenbeginn bzw. das Seitenende als bedingten Text mit der jeweiligen Seitenzahl ein.
//
//

var myPageStartTag = "PageStart";
var myPageEndTag = "PageEnd";

var myStartSwatchName = "PageStart";
var myEndSwatchName = "PageEnd";

var myCondSetOnName = "PageNumber_On";
var myCondSetOffName = "PageNumber_Off";

main(); // ==>

function main(){
	if(app.documents.length != 0){
		var myDoc = app.activeDocument;
		var myPage;

		// Bedingungen definieren
		var myStartCond = newCondition(myDoc, myPageStartTag); // ==>
		var myEndCond = newCondition(myDoc, myPageEndTag); // ==>
		
		// Sets definieren
		if(checkConditionSet(myDoc, myCondSetOnName)) myDoc.conditionSets.item(myCondSetOnName).remove();
		if(checkConditionSet(myDoc, myCondSetOffName)) myDoc.conditionSets.item(myCondSetOffName).remove();
		var myOnSettingArray = [[myStartCond, true], [myEndCond, true]];
		var myOffSettingArray = [[myStartCond, false], [myEndCond, false]];
		var myCondSetOn = myDoc.conditionSets.add({name:myCondSetOnName, setConditions:myOnSettingArray});
		var myCondSetOff = myDoc.conditionSets.add({name:myCondSetOffName, setConditions:myOffSettingArray});
		myDoc.conditionalTextPreferences.properties = {activeConditionSet:myCondSetOn};
		myCondSetOn.redefine();
		myDoc.conditionalTextPreferences.properties = {activeConditionSet:myCondSetOff};
		myCondSetOff.redefine();
		
		// Farben für bedingten Text (als Orientierungshilfe)
		if (checkSwatch(myDoc, myStartSwatchName) == false) myDoc.colors.add({name:myStartSwatchName, colorValue:[0,0,255], model:ColorModel.PROCESS, space:ColorSpace.RGB});
		if (checkSwatch(myDoc, myEndSwatchName) == false) myDoc.colors.add({name:myEndSwatchName, colorValue:[255,0,0], model:ColorModel.PROCESS, space:ColorSpace.RGB});
		var myStartSwatch = myDoc.swatches.itemByName(myStartSwatchName);
		var myEndSwatch = myDoc.swatches.itemByName(myEndSwatchName);
		var myStartCounter = 0;
		var myEndCounter = 0;
		for (var i = 0; i < myDoc.stories.length; i++) {
			var myStory = myDoc.stories[i];
			var myMasterSpreadFlag = true;
			for (var j = 0; j < myStory.textContainers.length; j++) {
				var myTestFrame = myStory.textContainers[j];
				// Wenn ein Textrahmen nicht auf der Montagefläche ...
				if(myTestFrame.parentPage != null) {
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
					if(myPage != null && myFrame.constructor.name == "TextFrame" && myFrame.insertionPoints.length > 0) {
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
							myLastIP.contents = myLastContent;
							myLastIP.fillColor = myEndSwatch;
							myLastIP.applyConditions(myEndCond, true);
							myEndCounter++;
						}
						// PageStart
						// Ist vorheriger Rahmen auf anderer Seite? ... Wenn ja, dann PageStart-Tag setzen
						if (checkPrevFrame(myFrame, myPage)) {
							var myFirstContent = myPageStartTag + "_" + myPage.name;
							var myFirstIP = myFrame.insertionPoints.firstItem();
							var myOldFirstIP = myFirstIP; // IP merken, da er sich bei evtl. Umbruchveränderungen verschiebt
							var myOldID = myOldFirstIP.parentStory.insertionPoints.itemByRange(0, myOldFirstIP.index).insertionPoints.length;
							myFirstIP.contents = myFirstContent;
							myFirstIP.fillColor = myStartSwatch;
							myFirstIP.applyConditions(myStartCond, true);
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
          for (var c = 0; c < myStory.tables[t].cells.length; c++) {
            var myCell = myStory.tables[t].cells[c];
            var firstCellChar = myCell.characters[0];
            if (firstCellChar != null) {
              var myFrame = firstCellChar.parentTextFrames[0];
              if (myFrame != null) {
     					  var myPage = myFrame.parentPage;
					      var myIP = myCell.insertionPoints.firstItem();
					      myIP.contents = "CellPage_" + myPage.name;
					      myIP.fillColor = myStartSwatch;
					      myIP.applyConditions(myStartCond, true);
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


function newCondition(myDoc, myCondName) {
	if(checkCondition(myDoc, myCondName)) {
		var myCond = myDoc.conditions.item(myCondName)
		myCond.visible = false;
		// Text mit dieser Bedingung löschen ...
		try {
			var myHiddenText = myDoc.stories.everyItem().hiddenTexts.everyItem().texts.everyItem().getElements(); 
		}
		catch(e) {}
		if(myHiddenText){
			for (var i = myHiddenText.length-1; i >= 0; i--) { 
				// Prüfung: versteckter Text hat mindestens eine Bedingung zugewiesen 
				if (myHiddenText[i].appliedConditions.length > 0) { 
					for (var x = myHiddenText[i].appliedConditions.length-1; x >= 0; x--) { 
						// Prüfung der dem versteckten zugewiesenen Bedingung 
						if (myHiddenText[i].appliedConditions[x].name == myCondName) {
	//alert(myHiddenText[i].texts[0].contents);
							myHiddenText[i].remove(); 
						}
					}
				}
			}
	  }
	      // How do I treat story hiddentext and cell hiddentext in a single pass? I.e., how can 
    // I merge the two lists? concat() did not work.
    try {
			var myHiddenText = myDoc.stories.everyItem().tables.everyItem().cells.everyItem().hiddenTexts.everyItem().texts.everyItem().getElements(); 
		}
		catch(e) {}
		if(myHiddenText){
			for (var i = myHiddenText.length-1; i >= 0; i--) { 
				if (myHiddenText[i].appliedConditions.length > 0) { 
					for (var x = myHiddenText[i].appliedConditions.length-1; x >= 0; x--) { 
						if (myHiddenText[i].appliedConditions[x].name == myCondName) {
							myHiddenText[i].remove(); 
						}
					}
				}
			}
		}
	}
	else{
		var myCond = myDoc.conditions.add({name:myCondName, visible:false});
	}
	return myCond;
}

function checkCondition(myDoc, myConditionName) {
	try {
		var myCond = myDoc.conditions.item(myConditionName);
		myCond.name;
		return true;
	}
	catch(e) { 
		return false; 
	}
}

function checkConditionSet(myDoc, myConditionSetName) {
	try {
		var myCondSet = myDoc.conditionSets.item(myConditionSetName);
		myCondSet.name;
		//alert(myConditionSetName + " gibt's schon!");
		return true;
	}
	catch(e) { 
		return false; 
	}
}

function checkSwatch(myDoc, mySwatchName) {
	try {
		myTestSwatch = myDoc.swatches.itemByName(mySwatchName);
		myTestSwatch.name;
		return true;
	}
	catch(e) {
		return false; 
	}
}
