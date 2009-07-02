<?php
// $Id$

/**
 * Implementation of hook_views_plugins().
 */
function sheetnode_xls_views_plugins() {
  return array(
    'style' => array(
      'sheet_xls' => array(
        'title' => t('XLS file'),
        'help' => t('Display the view as a Microsoft Excel (XLS) file.'),
        'handler' => 'sheetnode_xls_plugin_style',
        'parent' => 'sheet',
        'uses row plugin' => FALSE,
        'uses fields' => TRUE,
        'uses options' => TRUE,
        'type' => 'feed',
      ),
    ),
  );
}

