<?php
// $Id$

class sheetnode_handler_field_date extends views_handler_field_date {
  function options_form(&$form, &$form_state) {
    parent::options_form($form, $form_state);
    $form['date_format']['#options']['socialcalc'] = t('SocialCalc');
  }

  function render($values) {
    $format = $this->options['date_format'];
    if ($format != 'socialcalc') {
      return parent::render($values);
    }
    require_once(drupal_get_path('module', 'sheetnode') . '/socialcalc.inc');
    $value = $values->{$this->field_alias};
    return socialcalc_import_date($value);
  }
}

