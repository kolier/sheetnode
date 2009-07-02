<?php
// $Id$

/**
 * Implementation of hook_views_plugins().
 */
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
        'js' => array(
          drupal_get_path('module', 'sheetnode').'/socialcalc/socialcalcconstants.js',
          drupal_get_path('module', 'sheetnode').'/socialcalc/socialcalc-3.js',
          drupal_get_path('module', 'sheetnode').'/socialcalc/socialcalctableeditor.js',
          drupal_get_path('module', 'sheetnode').'/socialcalc/formatnumber2.js',
          drupal_get_path('module', 'sheetnode').'/socialcalc/formula1.js',
          drupal_get_path('module', 'sheetnode').'/socialcalc/socialcalcpopup.js',
          drupal_get_path('module', 'sheetnode').'/socialcalc/socialcalcspreadsheetcontrol.js',
          drupal_get_path('module', 'sheetnode').'/sheetnode.js',
        ),
      ),
    ),
  );
}

