<?php
// $Id$

require_once(drupal_get_path('module', 'sheetnode') .'/socialcalc.inc');

function sheetnode_get_values($nid, $cells) {
  $sheet = _sheetnode_load($nid);
  $sc = socialcalc_parse_values($sheet);
  $values = array();
  if (is_array($cells)) foreach ($cells as $coord) {
    $values[$coord] = $sc['cells'][$coord]['datavalue'];
  }
  else {
    $values[$cells] = $sc['cells'][$cells]['datavalue'];
  }
  return $values;
}

