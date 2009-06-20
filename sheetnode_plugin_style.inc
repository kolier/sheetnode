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
      $cell['datavalue'] = $info->label();
      $cell['datatype'] = 't';
      $cell['valuetype'] = 't';

      $sheet['cells'][] = $cell;
      $col++;
    }
    $row = 2;
    foreach ($this->view->result as $result) {
      $col = 1;
      foreach ($this->view->field as $field => $info) {
        if ($info->options['exclude']) continue;
        $cell['pos'] = array($col, $row);
        $value = $info->theme($result);
        $cell['datavalue'] = $value;
        $cell['datatype'] = is_numeric($value) ? 'v' : 't';
        $cell['valuetype'] = is_numeric($value) ? 'n' : 'th';

        $sheet['cells'][] = $cell;
        $col++;
      }
      $row++;
    }
    $sheet['attribs']['lastcol'] = $col-1;
    $sheet['attribs']['lastrow'] = $row-1;

    $edit['rowpanes'] = array(
      0 => array('first' => 1, 'last' => 1),
      1 => array('first' => 2, 'last' => $row-1),
    );
    $edit['colpanes'] = array(
      0 => array('first' => 1, 'last' => $col-1),
    );
    $edit['ecell'] = array(
      'coord' => 'A1',
    );

    // Inject the Sheetnode code.
    $socialcalc = array(
      'sheet' => $sheet,
      'edit' => $edit,
      'audit' => socialcalc_default_audit($sheet),
    );
    return _sheetnode_inject(socialcalc_save($socialcalc), FALSE);
  }
}

