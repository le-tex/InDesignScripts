#target indesign
// ReplaceSelectedTextAsImage.jsx
// written by Philipp Glatza, le-tex publishing services GmbH
// version 0.9.4 (2022-09-07)
//  - remove the created textframe, if not already removed
// version 0.9.3 (2021-01-18)
//  - remove any list number/label from paragraph
// version 0.9.2 (2020-12-15)
//  - better support for text in table cells
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
    strFilename = strImageFilenamePrefix + selection.parentTextFrames[0].parentPage.name + '_' + random + '.png'
    intStartPosNormalized = 0
    intEndPosNormalized = -1;
  if(selection.insertionPoints.itemByRange(0, 1).contents.toString().match(/^\s/g)) {
    intStartPosNormalized = 1;
  }
  if(selection.insertionPoints.itemByRange(-2, -1).contents.toString().match(/\s$/g)) {
    intEndPosNormalized = -2;
  }
  selection.insertionPoints.itemByRange(intStartPosNormalized, intEndPosNormalized).select()
  selection = app.selection[0];
  selection.duplicate(LocationOptions.AT_BEGINNING, myTf.insertionPoints.item(0));
  for(p = 0, ps = myTf.paragraphs.length; p < ps; p++) {
    myTf.paragraphs[p].bulletsAndNumberingListType = ListType.NO_LIST
    myTf.paragraphs[p].firstLineIndent = 0
    myTf.paragraphs[p].leftIndent = 0
    myTf.paragraphs[p].lastLineIndent = 0
  }
  myTf.textFramePreferences.useNoLineBreaksForAutoSizing = true
  myTf.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_AND_WIDTH
  myTf.fit (FitOptions.FRAME_TO_CONTENT);
  var myBounds = myTf.geometricBounds,
         myRect = null;
  myTf.exportFile(ExportFormat.PNG_FORMAT, File(Folder.myDocuments+'/' + strFilename))
  // text in table cell
  if(selection.parent.name.search("^[0-9]+:[0-9]+$") == 0) {
    var cwidth = selection.parent.width;
    myRect = selection.insertionPoints[0].rectangles.add();
    myRect.place (File(Folder.myDocuments+'/' + strFilename));
    selection.parent.width = cwidth;
  } else {
    myRect = selection.insertionPoints[0].rectangles.add( {geometricBounds:[0,0, 10, 10 ]});
    myRect.place (File(Folder.myDocuments+'/' + strFilename));
    myRect.geometricBounds = myBounds;
  }
  myRect.fit (FitOptions.CONTENT_TO_FRAME);
  myRect.anchoredObjectSettings.anchoredPosition = AnchorPosition.INLINE_POSITION;
  myRect.anchoredObjectSettings.anchorYoffset = 0;
  selection.remove()
  try{myTf.remove()}catch(e){}

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
