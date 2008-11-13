// $Id$

Drupal.sheetnode = {};

Drupal.sheetnode.onclicks = [];
Drupal.sheetnode.onunclicks = [];

Drupal.sheetnode.absoluteShowHide = function(tab) {
  this.onclicks[this.sheet.tabs[tab].name] = this.sheet.tabs[tab].onclick;
  this.sheet.tabs[tab].onclick = function(s, t) {
    if (Drupal.sheetnode.onclicks[t]) {
      Drupal.sheetnode.onclicks[t](s, t);
    }
    $(".absolute").show();
  }
  this.onunclicks[this.sheet.tabs[tab].name] = this.sheet.tabs[tab].onunclick;
  this.sheet.tabs[tab].onunclick = function(s, t) {
    if (Drupal.sheetnode.onunclicks[t]) {
      Drupal.sheetnode.onunclicks[t](s, t);
    }
    $(".absolute").hide();
  }
}

Drupal.sheetnode.startUp = function() {
  SocialCalc.Constants.defaultImagePrefix = Drupal.settings.sheetnode.imageprefix;
  SocialCalc.Constants.defaultCommentStyle = "background-repeat:no-repeat;background-position:top right;background-image:url("+ Drupal.settings.sheetnode.imageprefix +"-commentbg.gif);"
  SocialCalc.Constants.TCendcapClass = 
  SocialCalc.Constants.TCpanesliderClass = 
  SocialCalc.Constants.TClessbuttonClass = 
  SocialCalc.Constants.TCmorebuttonClass = 
  SocialCalc.Constants.TCscrollareaClass = 
  SocialCalc.Constants.TCthumbClass = "absolute";

  this.sheet = new SocialCalc.SpreadsheetControl();
  Drupal.sheetnode.absoluteShowHide(this.sheet.tabnums.edit);
  Drupal.sheetnode.absoluteShowHide(this.sheet.tabnums.sort);
  Drupal.sheetnode.absoluteShowHide(this.sheet.tabnums.comment);
  Drupal.sheetnode.absoluteShowHide(this.sheet.tabnums.names);
  this.sheet.ParseSheetSave(Drupal.settings.sheetnode.value);
  this.sheet.InitializeSpreadsheetControl(Drupal.settings.sheetnode.element);
  this.sheet.currentTab = this.sheet.tabnums.edit;
}

Drupal.sheetnode.resize = function() {
  if (this.sheet.SizeSSDiv()) {
    this.sheet.editor.ResizeTableEditor(this.sheet.width, this.sheet.height-
      (this.sheet.spreadsheetDiv.firstChild.offsetHeight + this.sheet.formulabarDiv.offsetHeight));
  }
}

$(document).ready(function() {
  $('#edit-submit').click(function() {
    $('#edit-save').val(Drupal.sheetnode.sheet.CreateSheetSave());
  });
  $(window).resize(function() {
    Drupal.sheetnode.resize();
  });
});

