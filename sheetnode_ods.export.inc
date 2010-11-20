<?php
// $Id$

function _sheetnode_ods_export($nid) {
  require_once(drupal_get_path('module', 'sheetnode') . '/socialcalc.inc');

  $node = node_load($nid);
  if (!$node || $node->type != 'sheetnode') return drupal_not_found();

  return _sheetnode_ods_export_do($node->title, socialcalc_parse(_sheetnode_load($node->nid, $node->vid)));
}

function _sheetnode_ods_export_do($title, $socialcalc) {
  drupal_set_message(t('OOCalc (.ods) export not yet implemented!'), 'warning');
  return '';
}

