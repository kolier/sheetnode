<?php
// $Id$

class sheetnode_xls_plugin_style extends sheetnode_feed_plugin_style {
  function init(&$view, &$display, $options = NULL) {
    parent::init($view, $display, $options = NULL);
    $this->feed_image = drupal_get_path('module', 'sheetnode_xls') . '/images/xls.gif';
  }

  function options(&$options) {
    parent::options($options);
    $options['format'] = 'xls';
  }

  function options_form(&$form, &$form_values) {
    parent::options_form($form, $form_values);
    $form['format'] = array(
      '#type' => 'radios',
      '#title' => t('Export as'),
      '#options' => array(
        'xls' => t('XLS'),
        'xlsx' => t('XLSX'),
      ),
      '#default_value' => $this->options['format'],
    );
  }

  function render() {
    require_once(drupal_get_path('module', 'sheetnode_xls') .'/sheetnode_xls.export.inc');
    $title = $this->display->handler->get_option('title');
    if (empty($title)) $title = $this->view->name;
    $this->inherit_options();
    _sheetnode_xls_export_do($this->options['format'], $title, parent::render_sheet());
  }
};

