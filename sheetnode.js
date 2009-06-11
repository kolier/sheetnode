// $Id$

Drupal.sheetnode = {};

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
  SocialCalc.Formula.SheetCache.loadsheet = function(sheetname) {
    data = $.ajax({
      type: 'POST',
      url: Drupal.settings.basePath+'sheetnode/load',
      data: 'sheetname='+escape(sheetname),
      datatype: 'text',
      async: false
    }).responseText;
    return data;
  }
}

Drupal.sheetnode.focusSetup = function() {
  $(".form-text,.form-textarea,.form-select").focus(function(e) {
    SocialCalc.CmdGotFocus(this);
  });
}

Drupal.sheetnode.startUp = function() {
  SocialCalc.Constants.defaultImagePrefix = Drupal.settings.sheetnode.imagePrefix;
  SocialCalc.Constants.defaultCommentStyle = "background-repeat:no-repeat;background-position:top right;background-image:url("+ Drupal.settings.sheetnode.imagePrefix +"commentbg.gif);"
  SocialCalc.Constants.defaultCommentClass = "cellcomment";

  this.sheet = new SocialCalc.SpreadsheetControl();

  // Remove audit tab.
  this.sheet.tabs.splice(this.sheet.tabnums.audit, 1);
  this.sheet.tabnums = {};
  for (var i=0; i<this.sheet.tabs.length; i++) {
    this.sheet.tabnums[this.sheet.tabs[i].name] = i;
  }
  
  // Hide toolbar if we're just viewing.
  if (!Drupal.settings.sheetnode.editMode) {
    this.sheet.tabbackground="display:none;";
    this.sheet.toolbarbackground="display:none;";
  }

  // Read in data and recompute.
  this.sheet.ParseSheetSave(Drupal.settings.sheetnode.value);
  this.sheet.ExecuteCommand('redisplay');
  this.sheet.InitializeSpreadsheetControl(Drupal.settings.sheetnode.element, Drupal.settings.sheetnode.editMode ? 700 : 0);
  
  Drupal.sheetnode.focusSetup();
  Drupal.sheetnode.functionsSetup();
  Drupal.sheetnode.loadsheetSetup();

  this.sheet.ExecuteCommand('recalc');
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
  $('.sheetnode-submit').parents('form').submit(function() {
    Drupal.sheetnode.save();
    return true;
  });
  $('.collapsed').each(function() {
    var ev = 'DOMAttrModified';
    if ($.browser.msie) {
      ev = 'propertychange';
    }
    $(this).bind(ev, function(e) {
      if (Drupal.sheetnode.sheet) {
        Drupal.sheetnode.sheet.editor.SchedulePositionCalculations();
      }
    }, false);
  });
  $(window).resize(function() {
    Drupal.sheetnode.resize();
  });
});

