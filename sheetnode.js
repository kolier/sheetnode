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
  Drupal.sheetnode.absoluteShowHide(this.sheet.tabnums.edit);
  Drupal.sheetnode.absoluteShowHide(this.sheet.tabnums.sort);
  Drupal.sheetnode.absoluteShowHide(this.sheet.tabnums.comment);
  Drupal.sheetnode.absoluteShowHide(this.sheet.tabnums.names);
  this.sheet.ParseSheetSave(Drupal.settings.sheetnode.value);
  this.sheet.InitializeSpreadsheetControl(Drupal.settings.sheetnode.element, Drupal.settings.sheetnode.editMode ? 700 : 0);
  this.sheet.currentTab = this.sheet.tabnums.edit;

  Drupal.sheetnode.tooltip();
}

Drupal.sheetnode.resize = function() {
  if (this.sheet.SizeSSDiv()) {
    this.sheet.editor.ResizeTableEditor(this.sheet.width, this.sheet.height-
      (this.sheet.spreadsheetDiv.firstChild.offsetHeight + this.sheet.formulabarDiv.offsetHeight));
  }
}

/*
 * Taken from: Easiest Tooltip and Image Preview Using jQuery by Alen Grakalic
 * http://cssglobe.com/post/1695/easiest-tooltip-and-image-preview-using-jquery
 */
Drupal.sheetnode.tooltip = function() {
  /* CONFIG */     
  xOffset = 10;
  yOffset = 20;
  // these 2 variable determine popup's distance from the cursor
  // you might want to adjust to get the right result   
  /* END CONFIG */
  $(".tooltip").hover(function(e) {
    this.t = this.title;
    this.title = "";     
    $("body").append("<p id='tooltip'>"+ this.t +"</p>");
    $("#tooltip")
      .css("top",(e.pageY - xOffset) + "px")
      .css("left",(e.pageX + yOffset) + "px")
      .fadeIn("fast");
  },
  function() {
    this.title = this.t;
    $("#tooltip").remove();
  });
  $(".tooltip").mousemove(function(e) {
    $("#tooltip")
      .css("top",(e.pageY - xOffset) + "px")
      .css("left",(e.pageX + yOffset) + "px");
  });
}

$(document).ready(function() {
  $('#edit-submit').click(function() {
    $('#edit-save').val(Drupal.sheetnode.sheet.CreateSheetSave());
  });
  $(window).resize(function() {
    Drupal.sheetnode.resize();
  });
});

