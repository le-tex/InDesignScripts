#target indesign
// ReplaceSelectedTextAsImage.jsx
// written by Philipp Glatza, le-tex publishing services GmbH
// version 0.9.1 (2020-06-10)
//  - anchor exported image inline with better size and fit options
// version 0.9 (2019-11-27)

// user variables
var strImageFilenamePrefix = "img_p"
var searchNextMathToolsFormat = false; // values: true or false: when 'true', jump to next MathTools markup


// functions

myDoc = app.activeDocument;
selection = app.selection[0];
if(selection) {
  var myTf = selection.insertionPoints.lastItem().textFrames.add({geometricBounds:[0, 0, 1, 1 ]});
  var d  = new Date(),
    random = Math.floor(Math.random()*Math.floor(d / 1000)),
    strFilename = strImageFilenamePrefix + selection.parentTextFrames[0].parentPage.name + '_' + random + '.png';
  selection.duplicate(LocationOptions.AT_BEGINNING, myTf.insertionPoints.item(0));
  myTf.textFramePreferences.useNoLineBreaksForAutoSizing = true
  myTf.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_AND_WIDTH
  myTf.fit (FitOptions.FRAME_TO_CONTENT);
  var myBounds = myTf.geometricBounds
  myTf.exportFile(ExportFormat.PNG_FORMAT, File(Folder.myDocuments+'/' + strFilename))
  var rect = selection.insertionPoints[0].rectangles.add( {geometricBounds:[0,0, 10, 10 ]});
  rect.place (File(Folder.myDocuments+'/' + strFilename));
  rect.geometricBounds = myBounds;
  rect.fit (FitOptions.CONTENT_TO_FRAME);
  rect.anchoredObjectSettings.anchoredPosition = AnchorPosition.INLINE_POSITION;
  rect.anchoredObjectSettings.anchorYoffset = 0;
  selection.remove()

  // search next MathTools cstyle
  if(searchNextMathToolsFormat) {
    MathToolsStyles = myDoc.characterStyleGroups.itemByName("~~~~~(MathTools)").characterStyles
    for(i = 0, j = MathToolsStyles.length; i<j ;i++) {
      app.findGrepPreferences= app.changeGrepPreferences = null; 
      app.findGrepPreferences.appliedCharacterStyle = app.documents[0].characterStyleGroups.itemByName("~~~~~(MathTools)").characterStyles.itemByName(MathToolsStyles[i].name); 
      app.findGrepPreferences.findWhat="."; // any character 
      found = myDoc.findGrep(); 
      if(found .length){
        found[0].select();
        i = j // break current for
      }
    app.findGrepPreferences= app.changeGrepPreferences = null; 
    }
  }
} else {
  alert("Please select text.");
}
