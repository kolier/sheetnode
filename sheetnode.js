(function ($) {
// START jQuery

Drupal.sheetnode = Drupal.sheetnode || {};

Drupal.sheetnode.functionsSetup = function() {
  // ORG.DRUPAL.FIELD server-side function.
  SocialCalc.Formula.FunctionList["ORG.DRUPAL.FIELD"] = [Drupal.sheetnode.functionDrupalField, -1, "drupalfield", "", "drupal"];
  SocialCalc.Constants["s_fdef_ORG.DRUPAL.FIELD"] = 'Returns a field from the specified Drupal entity (node, user, etc.)';
  SocialCalc.Constants.s_farg_drupalfield = 'field-name, [oid, entity-name]';

  // Update function classes.
  SocialCalc.Constants.function_classlist.push('drupal');
  SocialCalc.Constants.s_fclass_drupal = "Drupal";
}

Drupal.sheetnode.functionDrupalField = function(fname, operand, foperand, sheet) {
  var scf = SocialCalc.Formula;
  var oid, entity, field;

  field = scf.OperandValueAndType(sheet, foperand);
  oid = scf.OperandValueAndType(sheet, foperand);
  entity = scf.OperandValueAndType(sheet, foperand);
  if (isNaN(parseInt(oid.value))) {
    oid.value = Drupal.settings.sheetnode.context['oid'];
    entity.value = Drupal.settings.sheetnode.context['entity-name'];
  }

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
  if ($('div#'+Drupal.settings.sheetnode.containerElement, context).length == 0) return;
  if ($('div.sheetview-processed', context).length != 0) return;

  // Settings initialization.
  var containerElement = $('div#'+Drupal.settings.sheetnode.containerElement, context);
  var showEditor = Drupal.settings.sheetnode.saveElement || Drupal.settings.sheetnode.viewMode == Drupal.sheetnode.viewModes.fiddleMode;
  var showToolbar = showEditor && Drupal.settings.sheetnode.showToolbar;
  
  // SocialCalc initialization.
  SocialCalc.Popup.Controls = {};
  SocialCalc.ConstantsSetImagePrefix(Drupal.settings.sheetnode.imagePrefix);
  SocialCalc.Constants.defaultCommentClass = "cellcomment";
  SocialCalc.Constants.defaultReadonlyClass = "readonly";
  this.spreadsheet = showEditor ? new SocialCalc.SpreadsheetControl() : new SocialCalc.SpreadsheetViewer();
  if (showToolbar) {
    // Remove unwanted tabs.
    this.spreadsheet.tabs.splice(this.spreadsheet.tabnums.clipboard, 1);
    this.spreadsheet.tabs.splice(this.spreadsheet.tabnums.audit, 1);
    if (!Drupal.settings.sheetnode.permissions['edit sheet settings']) {
      this.spreadsheet.tabs.splice(this.spreadsheet.tabnums.settings, 1);
    }
    this.spreadsheet.tabnums = {};
    for (var i=0; i<this.spreadsheet.tabs.length; i++) {
      this.spreadsheet.tabnums[this.spreadsheet.tabs[i].name] = i;
    }
  }
  else {
    this.spreadsheet.tabbackground="display:none;";
    this.spreadsheet.toolbarbackground="display:none;";
  }

  // Trigger event to alert plugins that we've created the spreadsheet.
  containerElement.trigger('sheetnodeCreated', {spreadsheet: this.spreadsheet});

  // Read in data and recompute.
  var parts = this.spreadsheet.DecodeSpreadsheetSave(Drupal.settings.sheetnode.value);
  if (parts && parts.sheet) {
    this.spreadsheet.ParseSheetSave(Drupal.settings.sheetnode.value.substring(parts.sheet.start, parts.sheet.end));
  }
  if (showEditor) {
    this.spreadsheet.InitializeSpreadsheetControl(Drupal.settings.sheetnode.containerElement, 700, containerElement.width());
  }
  else {
    this.spreadsheet.InitializeSpreadsheetViewer(Drupal.settings.sheetnode.containerElement, 700, containerElement.width());
  }
  if (parts && parts.edit) {
    this.spreadsheet.editor.LoadEditorSettings(Drupal.settings.sheetnode.value.substring(parts.edit.start, parts.edit.end));
  }
  if (!Drupal.settings.sheetnode.saveElement && Drupal.settings.sheetnode.viewMode == Drupal.sheetnode.viewModes.htmlTable) {
    containerElement.html(SocialCalc.SpreadsheetViewerCreateSheetHTML(this.spreadsheet));
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

  // DOM initialization.
  if (Drupal.settings.sheetnode.saveElement) {
    $('#'+Drupal.settings.sheetnode.saveElement, context).parents('form').submit(function() {
      Drupal.sheetnode.save();
      return true;
    });
  }
  $(window).resize(function() {
    Drupal.sheetnode.resize();
  });
  $('div#SocialCalc-edittools', containerElement).parent('div').attr('id', 'SocialCalc-toolbar');
  $('td#SocialCalc-edittab', containerElement).parents('div:eq(0)').attr('id', 'SocialCalc-tabbar');
  $('input:text', containerElement).addClass('form-text socialcalc-input');
  $('input:radio', containerElement).addClass('form-radio socialcalc-input');
  $('input:checkbox', containerElement).addClass('form-checkbox socialcalc-input');
  $('textarea', containerElement).addClass('form-textarea socialcalc-input');
  $('select', containerElement).addClass('form-select socialcalc-input');
  $('input:button', containerElement).addClass('form-submit socialcalc-input');
  $('div#SocialCalc-sorttools td:first', containerElement).css('width', 'auto');
  $('div#SocialCalc-settingsview', containerElement).css('border', 'none').css('width', 'auto').css('height', 'auto');

  // Lock cells requires special permission.
  if (showToolbar && !Drupal.settings.sheetnode.permissions['edit sheet settings']) {
    $('span#SocialCalc-locktools', containerElement).css('display', 'none');
  }

  // Prepare for fullscreen handling when clicking the SocialCalc icon.
  $('td#'+SocialCalc.Constants.defaultTableEditorIDPrefix+'logo img', containerElement).attr('title', Drupal.t('Fullscreen')).click(function() {
    if (containerElement.hasClass('sheetview-fullscreen')) { // Going back to normal:
      // Restore saved values.
      containerElement.removeClass('sheetview-fullscreen');
      if (Drupal.sheetnode.beforeFullscreen.index >= Drupal.sheetnode.beforeFullscreen.parentElement.children().length) {
        Drupal.sheetnode.beforeFullscreen.parentElement.append(containerElement);
      } else {
        containerElement.insertBefore(Drupal.sheetnode.beforeFullscreen.parentElement.children().get(Drupal.sheetnode.beforeFullscreen.index));
      }
      Drupal.sheetnode.spreadsheet.requestedHeight = Drupal.sheetnode.beforeFullscreen.requestedHeight;
      Drupal.sheetnode.resize();
      $('body').css('overflow', 'auto');
      window.scroll(Drupal.sheetnode.beforeFullscreen.x, Drupal.sheetnode.beforeFullscreen.y);
    }
    else { // Going fullscreen:
      // Save current values.
      Drupal.sheetnode.beforeFullscreen = {
        parentElement: containerElement.parent(),
        index: containerElement.parent().children().index(containerElement),
        x: $(window).scrollLeft(), y: $(window).scrollTop(),
        requestedHeight: Drupal.sheetnode.spreadsheet.requestedHeight
      };

      // Set values needed to go fullscreen.
      $('body').append(containerElement).css('overflow', 'hidden');
      containerElement.addClass('sheetview-fullscreen');
      Drupal.sheetnode.resize();
      window.scroll(0,0);
    }
  });
  
  // Signal that we've processed this instance of sheetnode.
  containerElement.addClass('sheetview-processed');

  // Trigger event to alert plugins that we've built the spreadsheet.
  containerElement.trigger('sheetnodeReady', {spreadsheet: this.spreadsheet});

  // Force a recalc to refresh all values and scrollbars.
  this.spreadsheet.editor.EditorScheduleSheetCommands("recalc");
}

Drupal.sheetnode.resize = function() {
  // Adjust width and height if needed.
  containerElement = $('div#'+Drupal.settings.sheetnode.containerElement);
  if (containerElement.hasClass('sheetview-fullscreen')) {
    this.spreadsheet.requestedHeight = containerElement.height();
  }
  this.spreadsheet.requestedWidth = containerElement.width();
  this.spreadsheet.DoOnResize();
}

Drupal.sheetnode.save = function() {
  $('#'+Drupal.settings.sheetnode.saveElement).val(this.spreadsheet.CreateSpreadsheetSave());
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

