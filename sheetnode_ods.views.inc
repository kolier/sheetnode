<?php
// $Id$

/**
 * Implementation of hook_views_plugins().
 */
function sheetnode_ods_views_plugins() {
  return array(
    'style' => array(
      'sheet_ods' => array(
        'title' => t('ODS file'),
        'help' => t('Display the view as an OpenOffice.org Calc (ODS) file.'),
        'handler' => 'sheetnode_ods_plugin_style',
        'parent' => 'sheet',
        'uses row plugin' => FALSE,
        'uses fields' => TRUE,
        'uses options' => TRUE,
        'type' => 'feed',
      ),
    ),
  );
}

