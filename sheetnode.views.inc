<?php
// $Id$

function sheetnode_views_plugins() {
  return array(
    'style' => array(
      'sheet' => array(
        'title' => t('Spreadsheet'),
        'help' => t('Displays results in a browser spreadsheet.'),
        'handler' => 'sheetnode_plugin_style',
        'uses row plugin' => FALSE,
        'uses fields' => TRUE,
        'uses options' => TRUE,
        'type' => 'normal',
        'theme' => 'views_view_sheet',
      ),
    ),
  );
}

