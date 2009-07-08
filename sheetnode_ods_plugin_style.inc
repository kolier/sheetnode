<?php
// $Id$

class sheetnode_ods_plugin_style extends sheetnode_plugin_style {
  function init(&$view, &$display, $options = NULL) {
    parent::init($view, $display, $options = NULL);
    $this->feed_image = drupal_get_path('module', 'sheetnode_ods') . '/images/ods.png';
  }

  function render() {
    require_once(drupal_get_path('module', 'sheetnode_ods') .'/sheetnode_ods.export.inc');
    $title = $this->display->handler->get_option('title');
    if (empty($title)) $title = $this->view->name;
    _sheetnode_ods_export_do($title, parent::render_sheet());
  }
};

