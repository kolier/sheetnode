// $Id$
(function ($) {
// START jQuery

Drupal.sheetnode = Drupal.sheetnode || {};

Drupal.sheetnode.functionsSetup = function() {
  SocialCalc.Formula.FunctionList["SUMPRODUCT"] = [Drupal.sheetnode.functionSumProduct, -1, "rangen", "", "stat"];
  SocialCalc.Constants.s_fdef_SUMPRODUCT = 'Sums the pairwise products of 2 or more ranges. The ranges must be of equal length.';
  SocialCalc.Constants.s_farg_rangen = 'range1, range2, ...';

  // Override IF function to allow optional false value.
  SocialCalc.Formula.FunctionList["IF"] = [Drupal.sheetnode.functionIf, -2, "iffunc2", "", "test"];
  SocialCalc.Constants.s_farg_iffunc2 = "logical-expression, true-value[, false-value]";
}

Drupal.sheetnode.functionIf = function(fname, operand, foperand, sheet) {
  var cond, t;

  cond = SocialCalc.Formula.OperandValueAndType(sheet, foperand);
  t = cond.type.charAt(0);
  if (t != "n" && t != "b") {
    operand.push({type: "e#VALUE!", value: 0});
    return;
  }

  var op1, op2;

  op1 = foperand.pop();
  if (foperand.length == 1) {
    op2 = foperand.pop();
  }
  else if (foperand.length == 0) {
    op2 = {type: "n", value: 0};
  }
  else {
    scf.FunctionArgsError(fname, operand);
    return;
  }
  
  operand.push(cond.value ? op1 : op2);
  return;
}

Drupal.sheetnode.functionSumProduct = function(fname, operand, foperand, sheet) {
  var range, products = [], sum = 0;
  var scf = SocialCalc.Formula;
  var ncols = 0, nrows = 0;

  var PushOperand = function(t, v) {operand.push({type: t, value: v});};

  while (foperand.length > 0) {
    range = scf.TopOfStackValueAndType(sheet, foperand);
    if (range.type != "range") {
      PushOperand("e#VALUE!", 0);
      return;
    }
    rangeinfo = scf.DecodeRangeParts(sheet, range.value);
    if (!ncols) ncols = rangeinfo.ncols;
    else if (ncols != rangeinfo.ncols) {
      PushOperand("e#VALUE!", 0);
      return;
    }
    if (!nrows) nrows = rangeinfo.nrows;
    else if (nrows != rangeinfo.nrows) {
      PushOperand("e#VALUE!", 0);
      return;
    }
    for (i=0; i<rangeinfo.ncols; i++) {
      for (j=0; j<rangeinfo.nrows; j++) {
        k = i * rangeinfo.nrows + j;
        cellcr = SocialCalc.crToCoord(rangeinfo.col1num + i, rangeinfo.row1num + j);
        cell = rangeinfo.sheetdata.GetAssuredCell(cellcr);
        value = cell.valuetype == "n" ? cell.datavalue : 0;
        products[k] = (products[k] || 1) * value;
      }
    }
  }
  for (i=0; i<products.length; i++) {
    sum += products[i];
  }
  PushOperand("n", sum);

  return;
}

Drupal.sheetnode.loadsheetSetup = function() {
  SocialCalc.RecalcInfo.LoadSheet = function(sheetname) {
    data = $.ajax({
      type: 'POST',
      url: Drupal.settings.basePath+'sheetnode/load',
      data: 'sheetname='+escape(sheetname),
      datatype: 'text',
      async: false
    }).responseText;
    if (data !== null) {
      SocialCalc.RecalcLoadedSheet(sheetname, data, true);
    }
    return data;
  }
}

Drupal.sheetnode.focusSetup = function() {
  $(".form-text,.form-textarea,.form-select").focus(function(e) {
    SocialCalc.CmdGotFocus(this);
  });
}

Drupal.sheetnode.start = function(context) {
  // Just exit if the sheetnode is not in the new context or if it has already been processed.
  if ($('div#'+Drupal.settings.sheetnode.viewId, context).length == 0) return;
  if ($('div.sheetview-processed', context).length != 0) return;

  // DOM initialization.
  $('#'+Drupal.settings.sheetnode.editId, context).parents('form').submit(function() {
    Drupal.sheetnode.save();
    return true;
  });
  $('.collapsed').each(function() {
    var ev = 'DOMAttrModified';
    if ($.browser.msie) {
      ev = 'propertychange';
    }
    $(this).bind(ev, function(e) {
      if (Drupal.sheetnode.spreadsheet) {
        Drupal.sheetnode.spreadsheet.editor.SchedulePositionCalculations();
      }
    }, false);
  });
  $(window).resize(function() {
    Drupal.sheetnode.resize();
  });

  // SocialCalc initialization.
  SocialCalc.Popup.Controls = {};
  SocialCalc.Constants.defaultImagePrefix = Drupal.settings.sheetnode.imagePrefix;
  SocialCalc.Constants.defaultCommentStyle = "background-repeat:no-repeat;background-position:top right;background-image:url("+ Drupal.settings.sheetnode.imagePrefix +"commentbg.gif);"
  SocialCalc.Constants.defaultCommentClass = "cellcomment";
  this.spreadsheet = new SocialCalc.SpreadsheetControl();

  // Remove unwanted tabs.
  this.spreadsheet.tabs.splice(this.spreadsheet.tabnums.clipboard, 1);
  this.spreadsheet.tabs.splice(this.spreadsheet.tabnums.audit, 1);
  this.spreadsheet.tabnums = {};
  for (var i=0; i<this.spreadsheet.tabs.length; i++) {
    this.spreadsheet.tabnums[this.spreadsheet.tabs[i].name] = i;
  }
  
  // Hide toolbar if we're just viewing.
  if (!Drupal.settings.sheetnode.editMode) {
    this.spreadsheet.tabbackground="display:none;";
    this.spreadsheet.toolbarbackground="display:none;";
  }

  // Read in data and recompute.
  parts = this.spreadsheet.DecodeSpreadsheetSave(Drupal.settings.sheetnode.value);
  if (parts && parts.sheet) {
    this.spreadsheet.ParseSheetSave(Drupal.settings.sheetnode.value.substring(parts.sheet.start, parts.sheet.end));
  }
  this.spreadsheet.InitializeSpreadsheetControl(Drupal.settings.sheetnode.viewId, 700, $('div#'+Drupal.settings.sheetnode.viewId).width());
  if (parts && parts.edit) {
    this.spreadsheet.editor.LoadEditorSettings(Drupal.settings.sheetnode.value.substring(parts.edit.start, parts.edit.end));
  }

  // Special handling for Views AJAX.
  try {
    $('input[type=submit]', Drupal.settings.views.ajax.id).click(function() {
      Drupal.sheetnode.save();
    });
  }
  catch (e) {
    // Do nothing.
  }

  // Call our setup functions.
  Drupal.sheetnode.focusSetup();
  Drupal.sheetnode.functionsSetup();
  Drupal.sheetnode.loadsheetSetup();

  // Fix DOM where needed.
  $('div#SocialCalc-edittools', context).parent('div').attr('id', 'SocialCalc-toolbar');
  $('td#SocialCalc-edittab', context).parents('div:eq(0)').attr('id', 'SocialCalc-tabbar');
  $('div#'+Drupal.settings.sheetnode.viewId+' input:text', context).addClass('form-text');
  $('div#'+Drupal.settings.sheetnode.viewId+' input:radio', context).addClass('form-radio');
  $('div#'+Drupal.settings.sheetnode.viewId+' input:checkbox', context).addClass('form-checkbox');
  $('div#'+Drupal.settings.sheetnode.viewId+' textarea', context).addClass('form-textarea');
  $('div#'+Drupal.settings.sheetnode.viewId+' select', context).addClass('form-select');
  $('div#'+Drupal.settings.sheetnode.viewId+' input:button', context).addClass('form-submit');
  $('div#SocialCalc-sorttools td:first').css('width', 'auto');

  // Prepare for fullscreen handling when clicking the SocialCalc icon.
  $('td#'+SocialCalc.Constants.defaultTableEditorIDPrefix+'logo img').attr('title', Drupal.t('Fullscreen')).click(function() {
    div = $('div#'+Drupal.settings.sheetnode.viewId);
    if (div.hasClass('sheetview-fullscreen')) { // Going back to normal:
      // Restore saved values.
      div.removeClass('sheetview-fullscreen');
      Drupal.sheetnode.beforeFullscreen.parentElement.append(div);
      Drupal.sheetnode.spreadsheet.requestedHeight = Drupal.sheetnode.beforeFullscreen.requestedHeight;
      Drupal.sheetnode.resize();
      window.scroll(Drupal.sheetnode.beforeFullscreen.x, Drupal.sheetnode.beforeFullscreen.y);
    }
    else { // Going fullscreen:
      // Save current values.
      Drupal.sheetnode.beforeFullscreen = {
        parentElement: div.parent(),
        x: window.pageXOffset, y: window.pageYOffset,
        requestedHeight: Drupal.sheetnode.spreadsheet.requestedHeight
      };

      // Set values needed to go fullscreen.
      $('body').append(div);
      div.addClass('sheetview-fullscreen');
      Drupal.sheetnode.resize();
      window.scroll(0,0);
    }
  });
  
  // Signal that we've processed this instance of sheetnode.
  $('div#'+Drupal.settings.sheetnode.viewId, context).addClass('sheetview-processed');

  this.spreadsheet.ExecuteCommand('recalc');
}

Drupal.sheetnode.resize = function() {
  // Adjust width and height if needed.
  div = $('div#'+Drupal.settings.sheetnode.viewId);
  if (div.hasClass('sheetview-fullscreen')) {
    this.spreadsheet.requestedHeight = div.height();
  }
  this.spreadsheet.requestedWidth = div.width();
  
  // Call SocialCalc for resizing.
  if (this.spreadsheet.SizeSSDiv()) {
    this.spreadsheet.editor.ResizeTableEditor(this.spreadsheet.width, this.spreadsheet.height-
      (this.spreadsheet.spreadsheetDiv.firstChild.offsetHeight + this.spreadsheet.formulabarDiv.offsetHeight));
  }
}

Drupal.sheetnode.save = function() {
  $('#'+Drupal.settings.sheetnode.editId).val(this.spreadsheet.CreateSpreadsheetSave());
  log = $('#edit-log').val();
  if (log != undefined) {
    audit = this.spreadsheet.sheet.CreateAuditString();
    if (!log.length) {
      $('#edit-log').val(audit);
    }
    else {
      $('#edit-log').val(log + '\n' + audit);
    }
  }
}

// END jQuery
})(jQuery);

