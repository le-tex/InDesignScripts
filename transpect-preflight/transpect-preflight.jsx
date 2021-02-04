#targetengine'foo';

// transpect-preflight.jsx
// Valentina Horsch

// Blendet alle bedingten Texte bis auf myAllowedConditions aus.
// Gleicht alle namen von Absatz-, Zeichen-, Tabellen-, Zellen- und Objektformaten mit native-names einer cssa.xml (Speicherort wie Skript) ab. 
// Gibt Liste mit allen verwendeten und nicht in der cssa.xml gefunden Formaten aus

// ------------------------------ main ------------------------------   


var doc = openDoc();
if( doc != undefined ) {
    var saveStatus = saveDoc( doc );
    if( saveStatus == true ) {
        // this is where the "fun" starts
        var myDoc = app.activeDocument;
        var myFile = myDoc.filePath

        var myAllowdConditions = ["EOnly", "EpubAlternative", "PrintOnly"];
        var myConditions = myDoc.conditions;
        var conditionArray = [];
        for (var i = 0; i < myConditions.length; i++) {
            conditionArray.push(myConditions[i].visible)
            } // for
        
        var myRegEx = "\\r";

        startDialog();

        } // if
    else {
        alert("Datei muss gespeichert sein! Skript wird beendet.")
        } // else
    } // if
else {
    alert("Das Öffnen der Datei schieterte!")
    } // else



// ------------------------------ functions ------------------------------

function startDialog() {
    var dialog = new Window("palette", undefined, undefined); 
    dialog.text = "PreExportCheck"; 
    dialog.orientation = "column"; 
    dialog.alignChildren = ["fill","top"]; 
    dialog.spacing = 10; 
    dialog.margins = 16; 

    var statictext1 = dialog.add("statictext", undefined, undefined, {name: "statictext1"}); 
    statictext1.text = "Was soll geprüft werden?"; 

    var checkbox1 = dialog.add("checkbox", undefined, undefined, {name: "checkbox1"}); 
    checkbox1.text = "bedingte Texte"; 
    checkbox1.value = true; 

    var checkbox2 = dialog.add("checkbox", undefined, undefined, {name: "checkbox2"}); 
    checkbox2.text = "Text-, Objekt- und Tabellenformate"; 
    checkbox2.value = true; 

    var group1 = dialog.add("group", undefined, {name: "group1"}); 
    group1.orientation = "row"; 
    group1.alignChildren = ["center","center"]; 
    group1.spacing = 10; 
    group1.margins = 0; 

    var button1 = group1.add("button", undefined, undefined, {name: "button1"}); 
    button1.text = "Prüfung starten"; 
    button1.onClick = function() {
        var boolCheckboxOne = checkbox1.value;
        var boolCheckboxTwo = checkbox2.value;
        dialog.close();
        $.sleep(400);
        checkForParts(boolCheckboxOne, boolCheckboxTwo)
        } // function

    var button2 = group1.add("button", undefined, undefined, {name: "button2"}); 
    button2.text = "Datei exportieren"; 
    button2.onClick = function() {
        dialog.close();
        $.sleep(400);
        exportIdml()
        } // function
    
    var button3 = group1.add("button", undefined, undefined, {name: "cancel"}); 
    button3.text = "Abbrechen"; 
    button3.onClick = function() {
        dialog.close();
        } // function

    dialog.show();
    } // function


// --------------------------------------------------------------------------

function checkForParts(checkbox1, checkbox2) {
    if (checkbox1 == true && checkbox2 == true) {
        setConditions();
        checkStyles()
        } // if
    else if (checkbox1 == true && checkbox2 == false) {
        setConditions()
        askForExport()
        } // else if
    else if (checkbox1 == false && checkbox2 == true) {
        checkStyles()
        } // else if
    else if (checkbox1 == false && checkbox2 == false) {
        var confi = confirm("Es wird keine Prüfung durchgeführt.\nSoll ein *.idml exportiert werden?");
        if (confi) {
            exportIdml()
            } // function
        else {
            alert("Skript beendet.")
            } // else
        } // else if
    } // function


// --------------------------------------------------------------------------

function setConditions () {
    // alle bedingten Texte einblenden

    for (var i = 0; i < myConditions.length; i++) {
        myConditions[i].visible = true
        } // for
    
    // bedingten Text auf Absatzmarken suchen und löschen
    var myFoundArray = [];
    var myText = "An folgenden Stellen wurde bedingter Text von Absatzmarken entfernt: \n\n";
    
    app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = myRegEx;
    var foundItems = app.findGrep(true);
    
    var x = new Window ('palette');
    x.text = "Textbedingungen werden geprüft...";
    x.pbar = x.add ('progressbar', undefined, 0, foundItems.length);
    x.pbar.preferredSize.width = 300;
    x.show();
    x.pbar.value = 0;

    if(foundItems.length > 0) {
        for (var k = foundItems.length - 1; k >= 0; k--) {
            if(foundItems[k].appliedConditions.length != 0) {
                var myItem = foundItems[k];
                myFoundArray.push(myItem);
                var myPage = foundItems[k].parentTextFrames[0].parentPage;
                var myLine = foundItems[k].insertionPoints[0].lines[0];
                var myItemText = "S. " + myPage.name + ":\t[...]" + myLine.contents;
                myText += myItemText + "\n";
                } // if
            x.pbar.value = x.pbar.value + 1;
            } // for
        } // if
    for (var j = 0; j < myFoundArray.length; j++) {
        myFoundArray[j].appliedConditions = [];
        } // for
    //Log-Datei schreiben
    
    if (myText.length > 71) { // entspricht Mindestlänge vpn befüllter Textdatei
        var logName = "/PreExportCeck_bedingte-Texte.txt";
        saveAs(myText , logName);
        alert("Es wurden bedingte Texte von Absatzmarken gelöscht. \n\nEine Liste aller geänderten Textstellen finden sie unter\n" + myFile + logName)
        } // if
    else {
        alert("Prüfung der bedingten Texte beendet: \nAlles in Ordnung!");
        } 
    } // function


// --------------------------------------------------------------------------

function askForExport() {
    var confi = confirm("Soll die Datei exportiert werden?");
    if (confi) {
        exportIdml()
        } // if
    else {
        // Bedingungen wieder ein- bzw. ausblenden wie vor Export
        for (var i = 0; i < myConditions.length; i++) {
            var visible = conditionArray[i]
            myConditions[i].visible = visible
            } // for
        } // else
    } // function


// --------------------------------------------------------------------------

function checkStyles() {
    
    var allMissingParaStyles = [];
    var allMissingCharStyles = [];
    var allMissingTableStyles = [];
    var allMissingCellStyles = [];
    var allMissingObjStyles = [];
    
    
    var myScriptFile = myGetScriptPath();
    var end = myScriptFile.lastIndexOf("/");
    var myScriptString  = myScriptFile.slice(0, end + 1); 
    var myXMLFile = File(myScriptString + "cssa.xml");
    if (!myXMLFile.exists) {
        alert(myScriptString + "cssa.xml " + " \nkonnte nicht gefunden werden.");
        myScriptFile = File(myScriptFile);
        myXMLFile = myScriptFile.openDlg("Bitte wählen Sie die cssa.xml aus:", ["Xml-Dateien:*.xml", "Alle Dateien:*.*"])
        } // if

    if (myXMLFile != null) {
        var myXMLRoot = loadJS(myXMLFile, "/css:rules"); 
        
        //alle native-names sammeln ==> in XML suchen?
        var myNativeNames = String(myXMLRoot.xpath("/css:rule/@native-name"));
        var myCssRegExpArray = [];

        n = 0;
        while (myXMLRoot.xpath("/css:rule/@regex")[n] != null) {
            var myCssRegExp = myXMLRoot.xpath("/css:rule/@regex")[n];
            myCssString = String(myCssRegExp);
            myCssRegExpArray.push(myCssString);
            n++
            } // while
        
        //alle Absatzformate sammeln
        var myParaStyles = getAllStyles(myDoc, "Paragraph");
        var myCharStyles = getAllStyles(myDoc, "Character");
        var myTableStyles = getAllStyles(myDoc, "Table");        
        var myCellStyles = getAllStyles(myDoc, "Cell");
        var myObjStyles = getAllStyles(myDoc, "Object");
        
        var myStyleArray = [];
        myStyleArray.push(myParaStyles);
        myStyleArray.push(myCharStyles);
        myStyleArray.push(myTableStyles);
        myStyleArray.push(myCellStyles);
        myStyleArray.push(myObjStyles);
        
// ==============================================        
// NEU
        var progressLength = 0;
        for (var y = 0; y < myStyleArray.length; y++) { 
            var progressLength = progressLength + myStyleArray[y].length
            } // for

        var w = new Window ('palette');
        w.text = "Formatnamen werden geprüft..."
        w.pbar = w.add ('progressbar', undefined, 0, progressLength);
        w.pbar.preferredSize.width = 300;
        w.show();
        w.pbar.value = 0;

        //Absatzformte mit NativeNames vergleichen
        for (var j = 0; j < myStyleArray.length; j++) { 

            var myArray = new Array;
            var myArray = myStyleArray[j];
            var myIndex = j; 

            for (var i = 0; i < myArray.length; i++) {
            
                var myCurrentStyleTilde = myArray[i];
                //Tildenformate
       
                if (myCurrentStyleTilde.indexOf("~") != -1) {
                    var myCurrentIndexOf = myCurrentStyleTilde.indexOf("~"); 
                    var myCurrentStyle = myCurrentStyleTilde.slice(0, myCurrentIndexOf);
                    } // if
                else {
                    var myCurrentStyle = myCurrentStyleTilde
                    } // else
                
                //wird das Format verwendet?
                if (myIndex == 0) {
                    var type = "Paragraph"
                    var myStyleList = allMissingParaStyles
                    } // if
                else if (myIndex == 1) {
                    var type = "Character"
                    var myStyleList = allMissingCharStyles
                    } // else if
                else if (myIndex == 2) {
                    var type = "Table"
                    var myStyleList = allMissingTableStyles
                    } // else if
                else if (myIndex == 3) {
                    var type = "Cell"
                    var myStyleList = allMissingCellStyles
                    } // else if
                else if (myIndex == 4) {
                    var type = "Object"
                    var myStyleList = allMissingObjStyles
                    } // else if
                
                var myMissingStyle = getStyle(myCurrentStyleTilde, myDoc, type);
                app.changeTextPreferences = NothingEnum.nothing;
                app.findTextPreferences = NothingEnum.nothing;
                
            if (type == "Paragraph") {
                    app.findTextPreferences.appliedParagraphStyle = myMissingStyle
                    myFoundItems = [];
                    myFoundItems = app.findText(true)
                    if (myFoundItems.length != 0) {
                        myMissingStyle = getFullStyleName(myMissingStyle, type);
                        var found = true;
                        } // if
                    else {
                        found = false;
                        } // else
                    } // if

                else if (type == "Character") {
                    app.findTextPreferences.appliedCharacterStyle = myMissingStyle
                    myFoundItems = [];
                    myFoundItems = app.findText(true)
                    if (myFoundItems.length != 0) {
                        myMissingStyle = getFullStyleName(myMissingStyle, type);
                        var found = true;
                        } // if
                    else {
                        found = false;
                        } // else
                    } // else if
                
                else if (type == "Object") {
                    app.findObjectPreferences.appliedObjectStyles = myMissingStyle;
                    myFoundItems = [];
                    myFoundItems = app.findObject(true)
                    if (myFoundItems.length != 0) {
                        myMissingStyle = getFullStyleName(myMissingStyle, type);
                        var found = true;
                        } // if
                    else {
                        found = false;
                        } // else
                    } // else if                
                
                else if (type == "Table") {
                    var result = findTableStyle(myMissingStyle);
                    if (result == true) {
                        myMissingStyle = getFullStyleName(myMissingStyle, type);
                        var found = true;
                        } // if
                    else {
                        found = false;
                        } // else
                    } // else if
                
                else if (type == "Cell") {
                    var result = findCellStyle(myMissingStyle);
                    if (result == true) {
                        myMissingStyle = getFullStyleName(myMissingStyle, type);
                        myMissingStyle = myMissingStyle.replace(/\s\|\s/g, ":")
                        var found = true;
                        } // if
                    else {
                        found = false;
                        } // else
                    } // else if
                
                                
                if (found == true) { // wird im Text verwendet
                    if (myCurrentStyle.indexOf(":") != -1) {
                        var myRegExpStyle = myCurrentStyle.replace(/:/g, "_"); 
                        }
                    else {
                        var myRegExpStyle = myCurrentStyle;
                        }
                    var bool = false;    
                    // mit native-names abgleichen
                    if (myNativeNames.match(myCurrentStyle)) {
                        bool = true;
                        } 
                    // mit RegExp abgleichen
                    if (bool == false) {

// ACHTUNG: Funktion regex.test() löst Bug in ID aus -> System stürzt ab. Daher der Abgleich der Regexe über Grep-Suche/Ersetzen in ID

                    var newFrame = app.activeDocument.pages[0].textFrames.add();     // neune temporären Rahmen anlegen
                    newFrame.contents = myRegExpStyle;                                              // Namen des AF einfügen    

                    for (var k = 0; k < myCssRegExpArray.length; k++) {                     // durch alle cssaRegExp loopen
                        app.changeGrepPreferences = NothingEnum.nothing;    
                        app.findGrepPreferences = NothingEnum.nothing;
                        app.findGrepPreferences.findWhat = myCssRegExpArray[k];     
                        myFoundItems = [];
                        myFoundItems = newFrame.paragraphs[0].findGrep(true);           // Abgleich mit dem Text im neuen Textrahmen = Name des verwendeten AF
                        if (myFoundItems.length != 0) {                                                   // wenn myFoundItems.length > 1, matcht aktuelle RegExp auf den Namen des AF
                            bool = true;                                                                           // -> Verwendung des Namens erlaubt -> bool = true wird weitergegeben        
                            break;
                            } // if // auf Anhieb gefunden: weiter
                        else {}
                        app.changeGrepPreferences = NothingEnum.nothing;
                        app.findGrepPreferences = NothingEnum.nothing;
                        } // for   
                    
                    newFrame.remove()                                                                        //  temporären Textrahmen wieder löschen  
                    } // if
                    else  {}
                    
                    // nirgendwo gefunden
                    if (bool == false) {                        
                        myStyleList.push(myMissingStyle)
                        } // if
                    } // if
                else {} // else // AF wird nicht verwendet
                w.pbar.value = w.pbar.value + 1;
                } // for
            w.pbar.value = w.pbar.value + 1;
            } // for
        
        myDoc.save();
        
        if (allMissingParaStyles.length != 0) {
            var logText = createLogText(allMissingParaStyles, "Absatzformate:\n");
            }
        if (allMissingCharStyles.length != 0) {
            logText = createLogText(allMissingCharStyles, logText += "\nZeichenformate:\n");
            }
        if (allMissingTableStyles.length != 0) {
            logText = createLogText(allMissingTableStyles, logText += "\nTabellenformate:\n");
            }
        if (allMissingCellStyles.length != 0) {
            logText = createLogText(allMissingCellStyles, logText += "\nZellenformate:\n");
            }
        if (allMissingObjStyles.length != 0) {
            logText = createLogText(allMissingObjStyles, logText += "\nObjektformate:\n");
            }
        if (allMissingParaStyles.length == 0 && allMissingCharStyles.length == 0 && allMissingTableStyles.length == 0 && allMissingCellStyles.length == 0 && allMissingObjStyles.length == 0) {
            w.close();
            alert("Prüfung der Formatnamen beendet: \nAlles in Ordnung!")
            }
        else {
            var logName = "/PreExportCeck_falsche-Formatnamen.txt"
            saveAs(logText, logName);
            w.close();
            showList(allMissingParaStyles, allMissingCharStyles, allMissingTableStyles, allMissingCellStyles, allMissingObjStyles)
            }
        } // if
    else {} // else
    } // function


// --------------------------------------------------------------------------

function getStyle(_styleName, _doc, _type) {
    var _indexOfColon = _styleName.lastIndexOf(":")
    
    if (_type == "Paragraph") {
        var _allStyles = _doc.paragraphStyles
        var _allGroups = _doc.paragraphStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.paragraphStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.paragraphStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // if
    
    else if (_type == "Character") {
        var _allStyles = _doc.characterStyles
        var _allGroups = _doc.characterStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.characterStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.characterStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // else if
    else if (_type == "Table") {
        var _allStyles = _doc.tableStyles
        var _allGroups = _doc.tableStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.tableStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.tableStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // else if
    else if (_type == "Cell") {
        var _allStyles = _doc.cellStyles
        var _allGroups = _doc.cellStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.cellStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.cellStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // else if
    else if (_type == "Object") {
        var _allStyles = _doc.objectStyles
        var _allGroups = _doc.objectStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.objectStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.objectStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // else if
    return _styleName
    } // function

// --------------------------------------------------------------------------

function getStyleOne(_styleName, _doc, _type) {
    var _indexOfColon = _styleName.lastIndexOf(":")

// old
    if (_type == "Paragraph") {
        var _allStyles = _doc.paragraphStyles
        var _allGroups = _doc.paragraphStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.paragraphStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.paragraphStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // if
    
    else if (_type == "Character") {
        var _allStyles = _doc.characterStyles
        var _allGroups = _doc.characterStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.characterStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.characterStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // else if
    else if (_type == "Table") {
        var _allStyles = _doc.tableStyles
        var _allGroups = _doc.tableStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.tableStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.tableStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // else if
    else if (_type == "Cell") {
        var _allStyles = _doc.cellStyles
        var _allGroups = _doc.cellStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.cellStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.cellStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // else if
    else if (_type == "Object") {
        var _allStyles = _doc.objectStyles
        var _allGroups = _doc.objectStyleGroups
        if (_indexOfColon <= 0) { // keine Gruppe
            var _styleName = _allStyles.item(_styleName);
            } // if
        else { // in Gruppen
            var styleNameArray = _styleName.split(":");
            var group = _allGroups.item(styleNameArray[0]);
            for (var l = 1; l < styleNameArray.length - 1; l++) {
                group = group.objectStyleGroups.item(styleNameArray[l]);
                } // for
            var _styleName = group.objectStyles.item(styleNameArray[styleNameArray.length - 1]);
            } // else
        } // else if
    return _styleName
    } // function


// --------------------------------------------------------------------------

function findMatch(_currentParaStyle, _allNativeNames, _bool) {
    var _bool = false;
    for (var i = 0; i < _allNativeNames.length; i++) {
        var _currentNativeName = _allNativeNames[i];
        if (_currentParaStyle.match(_currentNativeName)) {
            _bool = true;
            break
            } // if
        else {
            _bool = false
            } // else
        } // for
    return _bool
    } // function


// --------------------------------------------------------------------------
      
function getAllStyles(myDoc, type) {
    if (type == "Paragraph") {
        var all_styles = myDoc.allParagraphStyles
        } // if
    else if (type == "Character") {
        var all_styles = myDoc.allCharacterStyles
        } // else if
    else if (type == "Table") {
        var all_styles = myDoc.allTableStyles
        } // else if
    else if (type == "Cell") {
        var all_styles = myDoc.allCellStyles
        } // else if
    else if (type == "Object") {
        var all_styles = myDoc.allObjectStyles
        } // else if
    
    var all_style_names = new Array;
    
    for (var i = 0; i < all_styles.length; i++) {
        a_name = all_styles[i].name;
        var g = all_styles[i]; 
        while (g.parent.constructor !== Document) {
            g = g.parent;
            a_name = g.name + ':' + a_name
            } // while
        a_name = a_name;
        all_style_names.push(a_name)
        } // for
    return all_style_names
    } // function


// --------------------------------------------------------------------------

function myGetScriptPath() {
    try {
        myFile = app.activeScript.path
        } // try
    catch (myError) {
        myFile = myError.fileName
        } // catch
    return myFile
    } // function


// --------------------------------------------------------------------------

function loadJS(xmlFile, type) {        
    xmlFile = new File(xmlFile);
    xmlFile.encoding = "UTF-8";
    xmlFile.open("r");
    var xmlStr = xmlFile.read();    
    xmlFile.close();
    var xml = new XML(xmlStr);      
    var root = xml.xpath(type);
    return root;
    } // function 


// --------------------------------------------------------------------------

function getFullStyleName(_style, _type) {
    var _styleName = _style.name
    var _group = _style;
    while (_group.parent.constructor !== Document) {
        _group = _group.parent;
        var _styleName = _group.name + ' | ' + _styleName
        } // while
    _styleName = _styleName;
    return _styleName
    } // function


// --------------------------------------------------------------------------

function showList(_allMissingParaStyles, _allMissingCharStyles, _allMissingTableStyles, _allMissingCellStyles, _allMissingObjStyles) {
    var myWindow = new Window ("palette", undefined, undefined);
    myWindow.text = "Folgende Formate konnten in cssa.xml nicht gefunden werden:"; 
    myWindow.orientation = "row";
    myWindow.alignChildren = "fill";
    var stylesGroup = myWindow.add ("group {alignChildren: ['left', 'fill']}");
    var stylesGroupPanel = stylesGroup.add ("panel", undefined, undefined);
    var stylesGroupList = stylesGroupPanel.add ("listbox", undefined, undefined, {multiselect: false}); 
    stylesGroupList.alignment = "fill";
    stylesGroupList.preferredSize = [500, 400];
    
    
    if (_allMissingParaStyles.length != 0) {
        for (var i = 0; i < _allMissingParaStyles.length; i++) {
            var listitem = stylesGroupList.add("item");
            listitem.text = "Absatzformat       " + _allMissingParaStyles[i];
            } // for
        } // if
    else {
        var listitem = stylesGroupList.add ("item");
        listitem.enabled = false;
        listitem.text = "Absatzformate:     Alles in Ordnung!";
        }
    
    var listitem = stylesGroupList.add("item");
    listitem.text = "";
    listitem.enabled = false;
    if (_allMissingCharStyles.length != 0) {
        for (var i = 0; i < _allMissingCharStyles.length; i++){
            var listitem = stylesGroupList.add ("item");
            listitem.text = "Zeichenformat      " +_allMissingCharStyles[i]
            } // for
        } // if
    else {
        var listitem = stylesGroupList.add ("item");
        listitem.enabled = false;
        listitem.text = "Zeichenformate:        Alles in Ordnung!";
        }


    var listitem = stylesGroupList.add("item");
    listitem.text = "";
    listitem.enabled = false;
    if (_allMissingObjStyles.length != 0) {
        for (var i = 0; i < _allMissingObjStyles.length; i++) {
            var listitem = stylesGroupList.add ("item");
            listitem.text = "Objektformat       " + _allMissingObjStyles[i]
            } // for
        } // if
    else {
        var listitem = stylesGroupList.add ("item");
        listitem.enabled = false;
        listitem.text = "Objektformate:     Alles in Ordnung!";
        }

    var listitem = stylesGroupList.add("item");
    listitem.text = "";
    listitem.enabled = false;
    var listitem = stylesGroupList.add("item");
    listitem.text = "----------------------------------------------------------------------------------------";
    listitem.enabled = false;
    var listitem = stylesGroupList.add ("item");
    listitem.text = "Hinweis: Automatische Suche von Tabellen- und Zellenformate nicht per Skript möglich."; // anpassen???
    listitem.enabled = false;
    var listitem = stylesGroupList.add ("item");
    listitem.text = "----------------------------------------------------------------------------------------";
    listitem.enabled = false;

    
    var listitem = stylesGroupList.add("item");
    listitem.text = "";
    listitem.enabled = false;
    if (_allMissingTableStyles.length != 0) {
        for (var i = 0; i < _allMissingTableStyles.length; i++) {
            var listitem = stylesGroupList.add ("item");
            listitem.enabled = false;
            listitem.text = "Tabellenformat     " + _allMissingTableStyles[i]
            } // for
        } // if
    else {
        var listitem = stylesGroupList.add ("item");
        listitem.enabled = false;
        listitem.text = "Tabellenformate:       Alles in Ordnung!";
        }
    
    var listitem = stylesGroupList.add("item");
    listitem.text = "";
    listitem.enabled = false;
    if (_allMissingCellStyles.length != 0) {
        for (var i = 0; i < _allMissingCellStyles.length; i++) {
            var listitem = stylesGroupList.add ("item");
            listitem.enabled = false;
            listitem.text =  "Zellenformat      " +_allMissingCellStyles[i];
            } // for
        } // if
    else {
        var listitem = stylesGroupList.add ("item");
        listitem.enabled = false;
        listitem.text = "Zellenformate:     Alles in Ordnung!";
        }
    
    var group1 = myWindow.add("group", undefined, {name: "group1"}); 
    group1.orientation = "column"; 
    group1.alignChildren = ["left","top"]; 
    group1.spacing = 10; 
    group1.margins = 0; 

    var button1 = group1.add("button", undefined, undefined, {name: "button1"}); 
    button1.text = "Format suchen"; 
    button1.preferredSize.width = 130;
    button1.onClick = function() {userSearch(stylesGroupList.selection)} // function

    var button4 = group1.add("button", undefined, undefined, {name: "button4"}); 
    button4.text = "Liste aktualisieren"; 
    button4.preferredSize.width = 130;
    button4.onClick = function() {myWindow.close();
        $.sleep(400);
        checkStyles()
        } // function    

    var button3 = group1.add("button", undefined, undefined, {name: "button3"}); 
    button3.text = "idml exportieren";
    button3.preferredSize.width = 130;
    button3.onClick = function() {myWindow.close();
        $.sleep(400);
        exportIdml()
        } // function

    var button2 = group1.add("button", undefined, undefined, {name: "cancel"}); 
    button2.text = "Schließen"; 
    button2.preferredSize.width = 130;
    button2.onClick = function() {
        myDoc.save();
        myWindow.close()
        } // function
   
    myWindow.show();    
    } // function


// --------------------------------------------------------------------------

function showDialog(_conditionName, _bool) {
    var dialog = new Window ("dialog", undefined, undefined); 
    dialog.text = "Bedingung entfernen"; 
    dialog.orientation = "row"; 
    dialog.alignChildren = ["center","top"]; 
    dialog.spacing = 10; 
    dialog.margins = 16; 
    
    var group1 = dialog.add("group", undefined, {name: "group1"}); 
    group1.orientation = "column"; 
    group1.alignChildren = ["center","center"]; 
    group1.spacing = 10; 
    group1.margins = 0; 
    
    var statictext1 = group1.add("statictext", undefined, undefined, {name: "statictext1"}); 
    statictext1.text = "Soll der bedingte Text " + _conditionName + " von markierter Stelle entfernt werden?"; 
    statictext1.justify = "center"; 
    
    var group2 = group1.add("group", undefined, {name: "group2"}); 
    group2.orientation = "row"; 
    group2.alignChildren = ["center","center"]; 
    group2.spacing = 10; 
    group2.margins = 0; 
    
    
    var button1 = group2.add("button", undefined, undefined, {name: "button1"}); 
    button1.text = "Ja"; 
    button1.onClick = function() {
        dialog.close(); 
        $.sleep (300);
        }
        
    var button2 = group2.add("button", undefined, undefined, {name: "cancel"}); 
    button2.text = "Nein"; 
    button2.onClick = function() {
        dialog.close();
        $.sleep (300);
        } // function
    
    dialog.show();
    } // function


// --------------------------------------------------------------------------

function exportIdml() {
    
    for (var i = 0; i < myConditions.length; i++) {
        for (var j = 0; j < myAllowdConditions.length; j++) {
            if (myConditions[i].name.match(myAllowdConditions[j])) {
                myConditions[i].visible = true
                break
                } // if
            else {
                myConditions[i].visible = false
                } // else
            } // for
        } // for
    
    // Hier export 
    var chapters = [];
    chapters.push(myDoc);
    var p_list = progress_list (create_list (chapters), '*.idml wird exportiert ...');
    for (var i = 0; i < chapters.length; i++) {
        try {
            p_list[i].text = '+';
            myFile = new File (myDoc.filePath + "\\" + myDoc.name.replace (/indd$/, 'idml'));
            myDoc.exportFile(ExportFormat.INDESIGN_MARKUP, myFile, false);
            } // try
        catch(e){alert(e);};
        } // for 
    p_list[0].parent.parent.close ();
    
    // Bedingungen wieder ein- bzw. ausblenden wie vor Export
    for (var i = 0; i < myConditions.length; i++) {
        var visible = conditionArray[i]
        myConditions[i].visible = visible
        } // for
    
    alert("Fertig!")
    } // function


// --------------------------------------------------------------------------

function progress_list (array, title)
    {
    var txt = [];
    dlg___ = new Window ('palette', title);
    dlg___.orientation = 'row';
    var gr1 = dlg___.add ('group');
    gr1.orientation = 'column';
    gr1.alignChildren = ['left','top'];
    for (var i = 0; i < array.length; i++)
        {
        txt[i] = gr1.add ('statictext', undefined, '');
        txt[i].characters = 1
        }
    var gr2 = dlg___.add ('group');
    gr2.minimumSize.width = 200;
    gr2.orientation = 'column';
    gr2.alignChildren = ['left','top'];
    for (var i = 0; i < array.length; i++)
        gr2.add ('statictext', undefined, array[i]);
    dlg___.show();
    return txt;
    }


// --------------------------------------------------------------------------

function create_list (f)
    {
    var array = [];
    for (i = 0; i < f.length; i++)
        array.push (f[i].name);
    return array
    }


// --------------------------------------------------------------------------

function userSearch(_selection) {
    if (_selection == null || _selection.text.match("\-\-")) {
        alert("Bitte wählen Sie ein Format aus.")
        } // if
    else {
        app.changeObjectPreferences = NothingEnum.nothing;
        app.findObjectPreferences = NothingEnum.nothing; 
        app.changeGrepPreferences = NothingEnum.nothing;
        app.findGrepPreferences = NothingEnum.nothing;        
        app.changeTextPreferences = NothingEnum.nothing;
        app.findTextPreferences = NothingEnum.nothing;  

        app.scriptMenuActions.itemByID(18694).invoke()
        if (_selection.text.match("Absatzformat")) {
            var type = "Paragraph"
            var mySelectedText = _selection.text;
            var myNewSelectedText = mySelectedText.slice(19, mySelectedText.length)
            myNewReplacedText = myNewSelectedText.replace(/\s/g, ""); // whitespace entfernen
            myNewReplacedText = myNewSelectedText.replace(/\s\|\s/g, ":")
            var mySelectedStyle = getStyleOne(myNewReplacedText, myDoc, type)
            app.findGrepPreferences.appliedParagraphStyle = mySelectedStyle
            app.findTextPreferences.appliedParagraphStyle = mySelectedStyle
            } // if
        
        else if (_selection.text.match("Zeichenformat")) {
            var type = "Character"
            var mySelectedText = _selection.text;
            var myNewSelectedText = mySelectedText.slice(19, mySelectedText.length)
            myNewReplacedText = myNewSelectedText.replace(/\s\|\s/g, ":")
            var mySelectedStyle = getStyle(myNewReplacedText, myDoc, type)

            app.findGrepPreferences.appliedCharacterStyle = mySelectedStyle
            app.findTextPreferences.appliedCharacterStyle = mySelectedStyle
            } // else if
        
        else if (_selection.text.match("Objektformat")) {
            var type = "Object"
            var mySelectedText = _selection.text;
            var myNewSelectedText = mySelectedText.slice(19, mySelectedText.length)
            myNewReplacedText = myNewSelectedText.replace(/\s\|\s/g, ":")
            var mySelectedStyle = getStyle(myNewReplacedText, myDoc, type)
            
            app.findObjectPreferences.appliedObjectStyles = mySelectedStyle
            } // else if
        } // else
    } // function


// --------------------------------------------------------------------------

function createLogText(myArray, myText) {
    for (var k = 0; k < myArray.length; k++ ) {
        myText += myArray[k] + "\n";
        } // for
    return myText
    } // function


// --------------------------------------------------------------------------

function saveAs(logText, logName) {
    var myTargetFileName = myDoc.filePath;
    var myTargetFile = new File(myTargetFileName + logName);
    myTargetFile.encoding = "UTF-8";
    storeFile(myTargetFile, logText); // ==> Datei mit Text "füllen"
    } // function
    

// --------------------------------------------------------------------------

function storeFile(thisFile, thisText) {
    thisFile.open('w');
    thisFile.write(thisText);
    thisFile.close();
    } // function
    

// --------------------------------------------------------------------------

function findTableStyle(_missingStyle) {
    var _allTables = getTables();
    var _bool = false;
    for (k = 0; k < _allTables.length; k++) {
        if (_allTables[k].appliedTableStyle == _missingStyle) {
            _bool = true;
            return _bool
            break
            } // if
        else {} // else
        } // for
    } // function


// --------------------------------------------------------------------------

function findCellStyle (_missingStyle) {
    var _allTables = getTables();
    var _bool = false;
    for (var k = 0; k < _allTables.length; k++) {
        var _allCells = _allTables[k].cells.everyItem().getElements();
        for (var n = 0; n < _allCells.length; n++) {
            var _cell = _allCells[n];
            if (_cell.appliedCellStyle == _missingStyle) {
                _bool = true;
                return _bool
                break
                } // if
            else {} // else
            } // for
        } // for
    } // function


// --------------------------------------------------------------------------

function getTables() {
    if (app.selection.length == 0) {
        return app.documents[0].stories.everyItem().tables.everyItem().getElements();
        } // if
    else if (app.selection[0].parent instanceof Cell) {
        return [app.selection[0].parent.parent]
        } // else if
    else if (app.selection[0].hasOwnProperty ('parentStory')) {
        return app.selection[0].parentStory.tables.everyItem().getElements()
        } // else if
    else {
        return app.documents[0].stories.everyItem().tables.everyItem().getElements()
        } // else
    } // function  


// --------------------------------------------------------------------------

function openDoc() {
    if (app.layoutWindows.length == 0) {
    var file = File.openDialog ("Select a file", "InDesign:*.indd;*.indb;*.idml, InDesign Document:*.indd, InDesign Book:*.indb, InDesign Markup:*.idml", true)
    try {
      app.open(File(file));
      return app.documents[0];
    } catch (e) {
      alert(e);
      return undefined;
    };
  } else {
    return app.documents[0];
  }
}


// --------------------------------------------------------------------------

function saveDoc ( doc ) {
  // check if document is saved
  if ( ( !doc.saved || doc.modified ) ) {
    if ( confirm ("Das Dokument muss gespeichert werden.", undefined)) {
      try {
        doc = doc.save();
        // file successfully saved
        return true;
      } catch ( e ) {
        alert ("The document couldn't be saved.\n" + e);
        // file couldn't be saved
        return false;
      }
    } else {
      // user cancelled the save dialog
      return false;
    }
  } else {
      // file not modified, go on
      return true;
  }
}


// --------------------------------------------------------------------------


function writeLog(message, dir, filename){
  var path = dir + '/' + filename;
  createDir(dir);
  var write_file = File(path);
  if (!write_file.exists) {
    write_file = new File(path);
  }
  d = new Date();
  var timestr = "[" + d.getHours() + ":" + d.getMinutes() + ":" + d.getSeconds() + "] "
  write_file.open('a', undefined, undefined);
  write_file.encoding = "UTF-8";
  write_file.lineFeed = "Unix";
  write_file.writeln(timestr + message);
  write_file.close();
}

// --------------------------------------------------------------------------

 function createDir (folder) {
  try {
    folder.create();
    return;
  } catch (e) {
    alert (e);
  }
} 
