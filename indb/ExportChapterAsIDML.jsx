/*
 * ExportChapterAsIDML.jsx
 *
 *
 * Exports selected chapters from an INDB file to IDML and creates a result XML file.
 *
 *
 * Note: this script requires at least InDesign Version 6.0 (CS4).
 *
 * Authors: Matthias Quiering, Martin Kraetke (@mkraetke)
 *
 * Parts of this script originate from the work of Peter Kahrel.
 */
#target indesign;

version = "v1.0.2";

lang = {
    pre: app.locale == 1279477613 ? 1 : 0 // en = 0, de = 1
}
panel = {
  title:["le-tex – Export INDB as IDML", "le-tex – INDB als IDML exportieren"][lang.pre],
  selectDirTitle:["Select output directory", "Speicherort auswählen"][lang.pre],
  buttonDirSelect:["Select", "Auswählen"][lang.pre],
  selectChaptersTitle:["Select book parts", "Buchteile auswählen"][lang.pre],
  buttonSelectAll:["Select all", "Alle markieren"][lang.pre],
  buttonDeselect:["Deselect all", "Auswahl aufheben"][lang.pre],
  buttonInvertSelection:["Invert selection", "Auswahl umkehren"][lang.pre],
  buttonOk:"OK",
  buttonCancel:["Cancel", "Abbrechen"][lang.pre],
  exportMsg:["Exporting...", "Exportiere..."][lang.pre],
  versionErrorMsg:["InDesign versions prior to CS4 are not supported.", "InDesign-Versionen vor CS4 werden nicht unterstützt."][lang.pre],
  finishedMsg:["Finished!", "Fertig!"][lang.pre],
  selectBookTitle:["Select INDB file", "INDB-Datei auswählen"][lang.pre],
  selectBookErrorMsg:["No INDB file selected", "Keine INDB-Datei geöffnet"][lang.pre]
}
values = {
  outputDir:null,
  selectedChapters:null
}
// Nur ausführbar, ab CS4 aufwärts
var myVersion = app.scriptPreferences.version;
if(Number(myVersion.slice(0, myVersion.indexOf("."))) > 5) {
  export_to_idml (app.books.item (get_book ()));
} else {
  alert(panel.versionErrorMsg)
}
function export_to_idml (book) {
  var win = draw_window(book);
  var myDefaultInteraction = app.scriptPreferences.userInteractionLevel;
  app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
  chapters_to_idml (book, values.selectedChapters.chapters, values.outputDir);
  app.scriptPreferences.userInteractionLevel = myDefaultInteraction;
}
function chapters_to_idml (thisbook, chapters, dir) {
  var myScriptFileName = app.activeScript;
  var myScriptFile = File(myScriptFileName);
  var myFolder = Folder(dir);
  if(myFolder != null){
    var myFolderName = myFolder.fsName; // Windows-spezifische Namenskonvention ( /f/TEMP/  ==>  F:\TEMP )
    var myDefaultViewPDF = app.pdfExportPreferences.viewPDF;
    var f;
    var myDocList = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r";
    var myFolderPath = "";
    if (myFolder.fullName.match(/^\/\w\//g)) myFolderPath = "/" + myFolder.fullName.substr(1, 1) + ":" + myFolder.fullName.slice(2,); // A) Netzlaufwerk (z.B. "/i/indesign")
    else if (myFolder.fullName.match(/^~/g)) myFolderPath = "/" + myFolder.fsName.replace(/\\/g, "/"); // B) Nutzerverzeichnis (z.B. "~/Desktop/...")
    else myFolderPath = myFolder.fullName; // C) Serverpfad (z.B. "//Rennratte/indesign/...")
    myFolderPath = myFolderPath.replace(/[%&,; \(\)]/g, '_');
    myDocList += "<collection stable=\"true\" xml:base=\"file:" + myFolderPath + "/\">\r";
    var p_list = progress_list (create_list (chapters), panel.exportMsg);
    for (var i = 0; i < chapters.length; i++){
      try{
        p_list[i].text = '+';
        var myDoc = app.open (chapters[i], false);
        f = new File (myFolderName + "\\" + chapters[i].name.replace(/indd$/, 'idml'));
        myDoc.exportFile(ExportFormat.INDESIGN_MARKUP, f, false);
        myDoc.close(SaveOptions.NO);
        myDocList += "	<doc href=\"" + chapters[i].name.replace(/indd$/, 'idml') + "\"/>\r";
      } catch(e) {
        alert(e);
      }
    }
    p_list[0].parent.parent.close ();
    myDocList += "</collection>";
    myTargetFile = new File(myFolderName + "\\" + thisbook.name.replace(/\.indb$/, '') + ".indb.xml");
    myTargetFile.encoding = "UTF-8";
    myTargetFile.open('w');
    myTargetFile.write(myDocList);
    myTargetFile.close();    
    alert(panel.finishedMsg);
  }
}
function first_page (doc){
  return String (doc.pages[0].name)
}
function create_list (f){
  var array = [];
  for (i = 0; i < f.length; i++)
    array.push (f[i].name);
  return array
}
// draw window
function draw_window (book){
  var array = [];
  var book_contents = book.bookContents.everyItem().fullName;
  for (var i = 0; i < book_contents.length; i++)
    array[i] = File (book_contents[i]).name;
  var w = new Window ("dialog", panel.title + " " + version + " | " + book.name);
  w.orientation = "column";
  w.alignChildren = ["fill", "fill"];
  var panelSelectDir = w.add("panel", undefined, panel.selectDirTitle);
  panelSelectDir.orientation = ["row"];
  panelSelectDir.alignChildren = ["left", "top"];
  var panelSelectDirInputPath = panelSelectDir.add("edittext");
  panelSelectDirInputPath.preferredSize.width = 255;
  panelSelectDirInputPath.text = Folder(book.filePath);
  var panelSelectDirButton = panelSelectDir.add("button", undefined, panel.buttonDirSelect);
  var panelChooseChapters = w.add("panel", undefined, panel.selectChaptersTitle);
  panelChooseChapters.alignChildren = ["left", "top"];
  panelChooseChapters.orientation = "row";
  var g1 = panelChooseChapters.add("group");
  g1.alignChildren = "fill";
  g1.orientation = "column";
  var list = g1.add ('listbox', undefined, array, {multiselect: true});
  list.maximumSize.height = 700;
  list.minimumSize.width = 250;
  var buttons = panelChooseChapters.add ('group');
  buttons.orientation = 'column';
  var select_all = buttons.add ('button', undefined, panel.buttonSelectAll);
  var deselect_all = buttons.add ('button', undefined, panel.buttonDeselect);
  var invert = buttons.add ('button', undefined, panel.buttonInvertSelection);
  var buttonOkCancel = w.add("group");
  buttonOkCancel.orientation = "row"
  buttonOkCancel.childrenAlign = "right"
  var ok = buttonOkCancel.add ('button', undefined, panel.buttonOk, {name: 'ok'});
  var cancel = buttonOkCancel.add ('button', undefined, panel.buttonCancel, {name:'cancel'});
  panelSelectDirButton.preferredSize = select_all.preferredSize = deselect_all.preferredSize = invert.preferredSize
    = ok.preferredSize = cancel.preferredSize = [120,20];
  panelSelectDirButton.onClick  = function(){
    var result = Folder.selectDialog ();
    if (result) {
      panelSelectDirInputPath.text = result;
    }
  }
  select_all.onClick = function (){
    var all_items = new Array();
    var L = list.items.length;
    for (var i = 0; i < L; i++)
      all_items[i] = list.items[i];
    list.selection = all_items;
  }
  // select all items on start-up
  select_all.notify();
  invert.onClick = function (){
    var selected_items = new Array();
    var L = list.items.length;
    for (var i = 0; i < L; i++)
      if (list.items[i].selected == false)
        selected_items.push (list.items[i]);
    list.selection = null;
    list.selection = selected_items;
  }  
  deselect_all.onClick = function (){
    list.selection = null;
  }
  if(w.show () == 1){
    var selected_docs = get_selected (list, book_contents);
    values.outputDir = panelSelectDirInputPath.text;
    values.selectedChapters = {chapters: selected_docs};
    w.close ();
  } else {
    w.close ();
    exit();
  }
}
function get_selected (selected_list, booklist){
  var array = [];
  for (var i = 0; i < selected_list.items.length; i++)
    if (selected_list.items[i].selected)
      array.push (booklist[selected_list.items[i].index]);
  return array
}
function array_pos (item, array){
  for (var i = 0; i < array.length; i++)
    if (item == array[i])
      return i;
  return 0
}
function read_history (s){
  var temp = "";
  var f = File (script_dir() + s);
  if (f.exists){
    f.open ('r');
    var temp = f.read ();
    f.close ();
  }
  return temp
}
function store_history (s, picked){
  var f = File (script_dir() + s);
  f.open ('w');
  f.write (picked);
  f.close ()
}
function progress_bar (stop, title){
  w___ = new Window ('palette', title);
  pb___ = w___.add ('progressbar', undefined, 0, stop);
  pb___.preferredSize = [300,20];
  w___.show()
  return pb___;
}
function progress_list (array, title){
  var txt = [];
  dlg___ = new Window ('palette', title);
  dlg___.orientation = 'row';
  var gr1 = dlg___.add ('group');
  gr1.orientation = 'column';
  gr1.alignChildren = ['left','top'];
  for (var i = 0; i < array.length; i++){
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
function script_dir(){
  try {return File (app.activeScript).path + '/'}
  catch (e) {return File (e.fileName).path + '/'}
}
function errorM (m){
  alert (m, 'Error', true);
  exit();
}
function get_book (){
  switch (app.books.length) {
    case 0: alert(panel.selectBookErrorMsg); exit();
    case 1: return app.books[0].name;
    default: return pick_book();
  }
}
function pick_book (){
  var w = new Window ("dialog", panel.selectBookTitle);
  w.alignChildren = "right";
  var g = w.add ("group");
  var list = g.add ("listbox", undefined, app.books.everyItem().name);
  list.minimumSize.width = 250;
  list.selection = 0;
  var b = w.add ("group");
  b.add ("button", undefined, panel.buttonOK, {name: "ok"})
  b.add ("button", undefined, panel.buttonCancel, {name: "cancel"})
  if (w.show () == 1)
    return list.selection.text;
  else
    exit ();
}
