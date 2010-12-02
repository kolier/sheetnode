// $Id$
(function ($) {
// START jQuery

Drupal.sheetnode = Drupal.sheetnode || {};

Drupal.sheetnode.functionsSetup = function() {
  // SUMPRODUCT function.
  SocialCalc.Formula.FunctionList["SUMPRODUCT"] = [Drupal.sheetnode.functionSumProduct, -1, "rangen", "", "stat"];
  SocialCalc.Constants.s_fdef_SUMPRODUCT = 'Sums the pairwise products of 2 or more ranges. The ranges must be of equal length.';
  SocialCalc.Constants.s_farg_rangen = 'range1, range2, ...';

  // Override IF function to allow optional false value.
  SocialCalc.Formula.FunctionList["IF"] = [Drupal.sheetnode.functionIf, -2, "iffunc2", "", "test"];
  SocialCalc.Constants.s_farg_iffunc2 = 'logical-expression, true-value[, false-value]';

  // DRUPALFIELD server-side function.
  SocialCalc.Formula.FunctionList["DRUPALFIELD"] = [Drupal.sheetnode.functionDrupalField, 3, "drupalfield"];
  SocialCalc.Constants.s_fdef_DRUPALFIELD = 'Returns a field from the specified Drupal entity (node, user, etc.)';
  SocialCalc.Constants.s_farg_drupalfield = 'oid, entity-name, field-name';

  // CEILING function.
  SocialCalc.Formula.FunctionList["CEILING"] = [Drupal.sheetnode.functionCeilingFloor, -1, "vsig", "", "math"];
  SocialCalc.Constants.s_fdef_CEILING = 'Rounds the given number up to the nearest integer or multiple of significance. Significance is the value to whose multiple of ten the value is to be rounded up (.01, .1, 1, 10, etc.)';
  SocialCalc.Constants.s_farg_vsig = 'value, [significance]';
  
  // FLOOR function.
  SocialCalc.Formula.FunctionList["FLOOR"] = [Drupal.sheetnode.functionCeilingFloor, -1, "vsig", "", "math"];
  SocialCalc.Constants.s_fdef_FLOOR = 'Rounds the given number down to the nearest multiple of significance. Significance is the value to whose multiple of ten the number is to be rounded down (.01, .1, 1, 10, etc.)';
}

Drupal.sheetnode.functionCeilingFloor = function(fname, operand, foperand, sheet) {
  var scf = SocialCalc.Formula;
  var val, sig, t;

  var PushOperand = function(t, v) {operand.push({type: t, value: v});};

  val = scf.OperandValueAndType(sheet, foperand);
  t = val.type.charAt(0);
  if (t != "n") {
    PushOperand("e#VALUE!", 0);
    return;
  }
  if (val.value == 0) { 
    PushOperand("n", 0); 
    return;
  }

  if (foperand.length == 1) {
    sig = scf.OperandValueAndType(sheet, foperand);
    t = val.type.charAt(0);
    if (t != "n") {
      PushOperand("e#VALUE!", 0);
      return;
    }
  }
  else if (foperand.length == 0) {
    sig = {type: "n", value: val.value > 0 ? 1 : -1};
  }
  else {
    PushOperand("e#VALUE!", 0);
    return;
  }
  if (sig.value == 0) {
    PushOperand("n", 0); 
    return;
  }
  if (sig.value * val.value < 0) {
    PushOperand("e#NUM!", 0);
    return;
  }
  
  switch (fname) {
    case "CEILING":
      PushOperand("n", Math.ceil(val.value / sig.value) * sig.value);
      break;
    case "FLOOR":
      PushOperand("n", Math.floor(val.value / sig.value) * sig.value);
      break;
  } 
}

Drupal.sheetnode.functionDrupalField = function(fname, operand, foperand, sheet) {
  var scf = SocialCalc.Formula;
  var oid, entity, field;

  oid = scf.OperandValueAndType(sheet, foperand);
  entity = scf.OperandValueAndType(sheet, foperand);
  if (isNaN(parseInt(oid.value))) {
    oid.value = Drupal.settings.sheetnode.context['oid'];
    entity.value = Drupal.settings.sheetnode.context['entity-name'];
  }
  field = scf.OperandValueAndType(sheet, foperand);

  $.ajax({
    type: 'POST',
    url: Drupal.settings.basePath+'sheetnode/field',
    data: 'oid='+oid.value+'&entity='+escape(entity.value)+'&field='+escape(field.value),
    datatype: 'json',
    async: false,
    success: function (data) {
      var result = Drupal.parseJson(data);
      operand.push(result);
    }
  });
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
  $('.form-text,.form-textarea,.form-select').not('.socialcalc-input').focus(function(e) {
    SocialCalc.CmdGotFocus(this);
  });
}

Drupal.sheetnode.viewModes = {
  readOnly: 0,
  fiddleMode: 1,
  htmlTable: 2
}

Drupal.sheetnode.start = function(context) {
  // Just exit if the sheetnode is not in the new context or if it has already been processed.
  if ($('div#'+Drupal.settings.sheetnode.view_id, context).length == 0) return;
  if ($('div.sheetview-processed', context).length != 0) return;
  
  // DOM initialization.
  $('#'+Drupal.settings.sheetnode.edit_id, context).parents('form').submit(function() {
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
  SocialCalc.Constants.defaultImagePrefix = Drupal.settings.sheetnode.image_prefix;
  SocialCalc.Constants.defaultCommentStyle = "background-repeat:no-repeat;background-position:top right;background-image:url("+ Drupal.settings.sheetnode.image_prefix +"commentbg.gif);"
  SocialCalc.Constants.defaultCommentClass = "cellcomment";
  SocialCalc.Constants.defaultReadonlyStyle = "background-repeat:no-repeat;background-position:top right;background-image:url("+ Drupal.settings.sheetnode.image_prefix +"lockbg.gif);"
  SocialCalc.Constants.defaultReadonlyClass = "readonly";
  SocialCalc.Constants.defaultUnhideLeftStyle = "float:left;width:9px;height:12px;cursor:col-resize;background-image:url("+ Drupal.settings.sheetnode.image_prefix +"unhideright.gif);padding:0px;";
  SocialCalc.Constants.defaultUnhideRightStyle = "float:right;width:9px;height:12px;cursor:col-resize;background-image:url("+ Drupal.settings.sheetnode.image_prefix +"unhideleft.gif);padding:0px;";
  this.spreadsheet = (Drupal.settings.sheetnode.editing || Drupal.settings.sheetnode.fiddling == Drupal.sheetnode.viewModes.fiddleMode) ? new SocialCalc.SpreadsheetControl() : new SocialCalc.SpreadsheetViewer();

  if (Drupal.settings.sheetnode.editing || Drupal.settings.sheetnode.fiddling == Drupal.sheetnode.viewModes.fiddleMode) {
    // Remove unwanted tabs.
    this.spreadsheet.tabs.splice(this.spreadsheet.tabnums.clipboard, 1);
    this.spreadsheet.tabs.splice(this.spreadsheet.tabnums.audit, 1);
    if (!Drupal.settings.sheetnode.perm_edit_sheet_settings) {
      this.spreadsheet.tabs.splice(this.spreadsheet.tabnums.settings, 1);
    }
    this.spreadsheet.tabnums = {};
    for (var i=0; i<this.spreadsheet.tabs.length; i++) {
      this.spreadsheet.tabnums[this.spreadsheet.tabs[i].name] = i;
    }
    
    // Hide toolbar if we're just viewing.
    if (!Drupal.settings.sheetnode.editing) {
      this.spreadsheet.tabbackground="display:none;";
      this.spreadsheet.toolbarbackground="display:none;";
    }
  }

  // Read in data and recompute.
  parts = this.spreadsheet.DecodeSpreadsheetSave(Drupal.settings.sheetnode.value);
  if (parts && parts.sheet) {
    this.spreadsheet.ParseSheetSave(Drupal.settings.sheetnode.value.substring(parts.sheet.start, parts.sheet.end));
  }
  if (Drupal.settings.sheetnode.editing || Drupal.settings.sheetnode.fiddling == Drupal.sheetnode.viewModes.fiddleMode) {
    this.spreadsheet.InitializeSpreadsheetControl(Drupal.settings.sheetnode.view_id, 700, $('div#'+Drupal.settings.sheetnode.view_id).width());
  }
  else {
    this.spreadsheet.InitializeSpreadsheetViewer(Drupal.settings.sheetnode.view_id, 700, $('div#'+Drupal.settings.sheetnode.view_id).width());
  }
  if (parts && parts.edit) {
    this.spreadsheet.editor.LoadEditorSettings(Drupal.settings.sheetnode.value.substring(parts.edit.start, parts.edit.end));
  }
  if (!Drupal.settings.sheetnode.editing && Drupal.settings.sheetnode.fiddling == Drupal.sheetnode.viewModes.htmlTable) {
    $('div#'+Drupal.settings.sheetnode.view_id).html(SocialCalc.SpreadsheetViewerCreateSheetHTML(this.spreadsheet));
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
  div = $('div#'+Drupal.settings.sheetnode.view_id, context);
  $('div#SocialCalc-edittools', div).parent('div').attr('id', 'SocialCalc-toolbar');
  $('td#SocialCalc-edittab', div).parents('div:eq(0)').attr('id', 'SocialCalc-tabbar');
  $('input:text', div).addClass('form-text socialcalc-input');
  $('input:radio', div).addClass('form-radio socialcalc-input');
  $('input:checkbox', div).addClass('form-checkbox socialcalc-input');
  $('textarea', div).addClass('form-textarea socialcalc-input');
  $('select', div).addClass('form-select socialcalc-input');
  $('input:button', div).addClass('form-submit socialcalc-input');
  $('div#SocialCalc-sorttools td:first', div).css('width', 'auto');
  $('div#SocialCalc-settingsview', div).css('border', 'none').css('width', 'auto').css('height', 'auto');

  // Lock cells requires special permission.
  if (Drupal.settings.sheetnode.editing && !Drupal.settings.sheetnode.perm_edit_sheet_settings) {
    $('#'+Drupal.sheetnode.spreadsheet.idPrefix+'locktools').css('display', 'none');
  }

  // Prepare for fullscreen handling when clicking the SocialCalc icon.
  $('td#'+SocialCalc.Constants.defaultTableEditorIDPrefix+'logo img', div).attr('title', Drupal.t('Fullscreen')).click(function() {
    if (div.hasClass('sheetview-fullscreen')) { // Going back to normal:
      // Restore saved values.
      div.removeClass('sheetview-fullscreen');
      if (Drupal.sheetnode.beforeFullscreen.index >= Drupal.sheetnode.beforeFullscreen.parentElement.children().length) {
        Drupal.sheetnode.beforeFullscreen.parentElement.append(div);
      } else {
        div.insertBefore(Drupal.sheetnode.beforeFullscreen.parentElement.children().get(Drupal.sheetnode.beforeFullscreen.index));
      }
      Drupal.sheetnode.spreadsheet.requestedHeight = Drupal.sheetnode.beforeFullscreen.requestedHeight;
      Drupal.sheetnode.resize();
      $('body').css('overflow', 'auto');
      window.scroll(Drupal.sheetnode.beforeFullscreen.x, Drupal.sheetnode.beforeFullscreen.y);
    }
    else { // Going fullscreen:
      // Save current values.
      Drupal.sheetnode.beforeFullscreen = {
        parentElement: div.parent(),
        index: div.parent().children().index(div),
        x: $(window).scrollLeft(), y: $(window).scrollTop(),
        requestedHeight: Drupal.sheetnode.spreadsheet.requestedHeight
      };

      // Set values needed to go fullscreen.
      $('body').append(div).css('overflow', 'hidden');
      div.addClass('sheetview-fullscreen');
      Drupal.sheetnode.resize();
      window.scroll(0,0);
    }
  });
  
  // Signal that we've processed this instance of sheetnode.
  div.addClass('sheetview-processed');

  // Force a recalc to refresh all values and scrollbars.
  this.spreadsheet.editor.EditorScheduleSheetCommands("recalc");
}

Drupal.sheetnode.resize = function() {
  // Adjust width and height if needed.
  div = $('div#'+Drupal.settings.sheetnode.view_id);
  if (div.hasClass('sheetview-fullscreen')) {
    this.spreadsheet.requestedHeight = div.height();
  }
  this.spreadsheet.requestedWidth = div.width();
  this.spreadsheet.DoOnResize();
}

Drupal.sheetnode.save = function() {
  $('#'+Drupal.settings.sheetnode.edit_id).val(this.spreadsheet.CreateSpreadsheetSave());
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

