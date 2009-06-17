<?php
// $Id$

require_once('socialcalc.inc');

class sheetnode_plugin_style extends views_plugin_style {
  function render() {
    // Hand-make SocialCalc structure based on views results.
    $col = 1;
    foreach ($this->view->field as $field => $info) {
      if ($info->options['exclude']) continue;
      $cell['pos'] = array($col, 1);
      $cell['datavalue'] = $info->options['label'];
      $cell['datatype'] = 't';
      $cell['valuetype'] = 't';

      $sc['cells'][] = $cell;
      $col++;
    }
    $row = 2;
    foreach ($this->view->result as $result) {
      $col = 1;
      foreach ($this->view->field as $field => $info) {
        if ($info->options['exclude']) continue;
        $cell['pos'] = array($col, $row);
        $value = $this->view->field[$field]->theme($result);
        $cell['datavalue'] = $value;
        $cell['datatype'] = is_numeric($value) ? 'v' : 't';
        $cell['valuetype'] = is_numeric($value) ? 'n' : 'th';

        $sc['cells'][] = $cell;
        $col++;
      }
      $row++;
    }
    $sc['attribs']['lastcol'] = $col-1;
    $sc['attribs']['lastrow'] = $row-1;

    // Inject the Sheetnode code.
    return _sheetnode_inject(socialcalc_save($sc), FALSE);
  }
}

