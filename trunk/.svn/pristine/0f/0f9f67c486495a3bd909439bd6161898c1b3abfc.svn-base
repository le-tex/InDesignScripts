#target indesign

//
// PageNamesToStoryRevolution.jsx
// Version 1.1
//
// created by: maqui  |  10:18 28.02.2013
// modified by: maqui |  16:59 28.02.2013 (Löschen der Bedingungstexte ohne Einblenden --> Formatierung bleibt 1:1 erhalten)
// 
// Erstellt auf jedem Seitenbeginn und jedem Seitenende bedingten Text mit der jeweiligen Seitenzahl
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
  var myDoc = app.documents[0];
  var myPage;
  var mainStory = determineMainStory(myDoc); // ==>

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
	
  // Alle Textrahmen vom Hauptext durchlaufen
	var myFrame = mainStory.textContainers[mainStory.textContainers.length-1];
	var myOldPage;
	while (myFrame != null) {
		var myPage = myFrame.parentPage;
		if(myFrame.constructor.name == "TextFrame" && myPage != null && myFrame.insertionPoints.length > 0) {
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
				}
				myLastIP.contents = myLastContent;
				myLastIP.fillColor = myEndSwatch;
				myLastIP.applyConditions(myEndCond, true);
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
			}
			myOldPage = myPage;
		}
		var myFrame = myFrame.previousTextFrame;
	}
	myDoc.select(NothingEnum.NOTHING);
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


// Pick the story that is parent story to the largest number of text frames:
function determineMainStory(doc) {
  var textFrameCount = new Object;
  var max = 0;
  var mainStory = false;
  for (var p = 1; p <= doc.pages.length; p++) {
    page = doc.pages[p-1];
    for (var t = 0; t < page.textFrames.length; t++) {
			if(mainStory == false) mainStory = page.textFrames[t].parentStory;
      var key = String(page.textFrames[t].parentStory.id);
      if (textFrameCount[key] == null) {
        textFrameCount[key] = 1;
      } else {
        textFrameCount[key] += 1;
        if (textFrameCount[key] > max) {
          max = textFrameCount[key];
          mainStory = page.textFrames[t].parentStory;
        }
      }
    }
  }
  return mainStory;
}

function newCondition(myDoc, myCondName) {
	if(checkCondition(myDoc, myCondName)) {
		var myCond = myDoc.conditions.item(myCondName)
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
