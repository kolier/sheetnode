// $Id$

Drupal.sheetnode = {};

Drupal.sheetnode.focusSetup = function() {
  $(".form-text,.form-textarea,.form-select").focus(function(e) {
    SocialCalc.CmdGotFocus(this);
  });
}

Drupal.sheetnode.startUp = function() {
  SocialCalc.Constants.defaultImagePrefix = Drupal.settings.sheetnode.imagePrefix;
  SocialCalc.Constants.defaultCommentStyle = "background-repeat:no-repeat;background-position:top right;background-image:url("+ Drupal.settings.sheetnode.imagePrefix +"-commentbg.gif);"
  SocialCalc.Constants.defaultCommentClass = "tooltip";
  SocialCalc.Constants.TCendcapClass = 
  SocialCalc.Constants.TCpanesliderClass = 
  SocialCalc.Constants.TClessbuttonClass = 
  SocialCalc.Constants.TCmorebuttonClass = 
  SocialCalc.Constants.TCscrollareaClass = 
  SocialCalc.Constants.TCthumbClass = "absolute";

  this.sheet = new SocialCalc.SpreadsheetControl();
  if (!Drupal.settings.sheetnode.editMode) {
    this.sheet.tabbackground="display:none;";
    this.sheet.toolbarbackground="display:none;";
  }
  
  this.sheet.ParseSheetSave(Drupal.settings.sheetnode.value);
  this.sheet.FullRefreshAndRender();
  this.sheet.InitializeSpreadsheetControl(Drupal.settings.sheetnode.element, Drupal.settings.sheetnode.editMode ? 700 : 0);
  this.sheet.currentTab = this.sheet.tabnums.edit;

  Drupal.sheetnode.focusSetup();
}

Drupal.sheetnode.resize = function() {
  if (this.sheet.SizeSSDiv()) {
    this.sheet.editor.ResizeTableEditor(this.sheet.width, this.sheet.height-
      (this.sheet.spreadsheetDiv.firstChild.offsetHeight + this.sheet.formulabarDiv.offsetHeight));
  }
}

Drupal.sheetnode.save = function() {
  $('#edit-sheetsave').val(Drupal.sheetnode.sheet.CreateSheetSave());
  log = $('#edit-log').val();
  audit = Drupal.sheetnode.sheet.sheet.CreateAuditString();
  if (!log.length) {
    $('#edit-log').val(audit);
  }
  else {
    $('#edit-log').val(log + '\n' + audit);
  }
}

$(document).ready(function() {
  $('.sheetnode-submit').click(function() {
    Drupal.sheetnode.save();
  });
  $(window).resize(function() {
    Drupal.sheetnode.resize();
  });
});

