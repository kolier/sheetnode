<?php
// $Id$

require_once('socialcalc.inc');

define('SHEETNODE_EXPANSION_VERTICAL', 0);
define('SHEETNODE_EXPANSION_HORIZONTAL', 1);

class sheetnode_plugin_style extends views_plugin_style {
  function options(&$options) {
    parent::options($options);
    $options['expansion'] = SHEETNODE_EXPANSION_VERTICAL; 
    $options['template'] = TRUE;
    $options['sheetsave'] = '';
  }

  function options_form(&$form, &$form_values) {
    parent::options_form($form, $form_values);
    $form['expansion'] = array(
      '#type' => 'radios',
      '#title' => t('Expansion of results'),
      '#description' => t('You can specify whether view results should be expanded horizontally or vertically.'),
      '#options' => array(
        SHEETNODE_EXPANSION_VERTICAL => t('Consecutive rows'),
        SHEETNODE_EXPANSION_HORIZONTAL => t('Consecutive columns'),
      ),
      '#default_value' => $this->options['expansion'], 
    );
    $form['template'] = array(
      '#type' => 'checkbox',
      '#title' => t('Use template'),
      '#description' => t('Check this box to use the spreadsheet below as template for your view. 
                           To place view results in the template, use the following syntax in the cells:
                           <ul>
                           <li><code>${row.field_label}</code> to indicate that this cell and subsequent ones should be filled with values of this field.</li>
                           <li><code>$[FUNCTION(@row.field_A_label@, @row.field_B_label@, ...)]</code> to indicate that a function should be placed in this cell. 
                           Field names will be replaced with the corresponding cell ranges.</li>
                           </ul>'),
      '#default_value' => $this->options['template'],
    );
    global $base_url;
    $imagePrefix = drupal_to_js($base_url .'/' .drupal_get_path('module', 'sheetnode') .'/socialcalc/images/sc-');
    $value = drupal_to_js($this->options['sheetsave']);
    $form['sheetview']['#value'] = <<<EOS
<div class="sheetview" id="sheetview">
<script language="javascript" type="text/javascript">
  Drupal.behaviors.sheetnode = function(context) {
    if ($('div.sheetview').length == 0) return;

    jQuery.extend(Drupal.settings, { "sheetnode": { "value": $value, "imagePrefix": $imagePrefix, "element": "sheetview", "sheetsave": "edit-style-options-sheetsave", "editMode": true } });
    var ajax_area = Drupal.settings.views.ajax.id;
    $('input[type=submit]', ajax_area).click(function() {
      Drupal.sheetnode.save();
    });
    Drupal.sheetnode.start();
  }
</script>
</div>
EOS;
    $form['sheetsave'] = array('#type' => 'hidden');
  }

  function render() {
    if (strpos($_REQUEST['q'], 'admin/build/views/ajax/preview') === 0) return t('<div class="messages error">Spreadsheet previewing is not supported yet.</div>');

    $tangent = $this->options['expansion'];
    $normal  = 1-$tangent;
    if ($this->options['template']) {
      $socialcalc = socialcalc_parse($this->options['sheetsave']);
      $lastpos = array(0,0);

      // Iterate through cells, gathering placeholder values.
      foreach ($socialcalc['sheet']['cells'] as $coord => $cell) {
        // Field placeholder?
        $matches = array();
        if (isset($cell['datavalue']) && $cell['datatype'] == 't') {
          if (preg_match('/^\$\{row\.(\w+)\}$/', $cell['datavalue'], $matches)) {
            $field_name = $this->get_real_field_name($matches[1]);
            if ($field_name && !$this->view->field[$field_name]->options['exclude']) {
              $fields[$field_name] = array(
                'coord' => $coord,
                'cell' => $cell,
                'info' => $this->view->field[$field_name],
              );
            }
          }
          else if (preg_match('/^\$\[(\w+)\((.*?)\)]$/', $cell['datavalue'], $matches)) {
            $formula = $matches[1];
            $expression = $matches[2];
            $formulas[] = array(
              'coord' => $coord,
              'cell' => $cell,
              'formula' => $formula,
              'expression' => $expression,
            );
          }
        }
        $lastpos = array(max($lastpos[0], $cell['pos'][0]), max($lastpos[1], $cell['pos'][1]));
      }

      // Replace field placeholders with actual values.
      $tangent_increment = array(SHEETNODE_EXPANSION_VERTICAL => 0, SHEETNODE_EXPANSION_HORIZONTAL => 1);
      if (!empty($fields)) foreach ($fields as $field_name => $field) {
        unset($socialcalc['sheet']['cells'][$field['coord']]);
        $newcell = $field['cell'];
        $pos = $newcell['pos'];
        foreach ($this->view->result as $result) {
          $newcell['pos'] = $pos;
          $value = $field['info']->theme($result);
          $newcell['datavalue'] = $value;
          $newcell['datatype'] = is_numeric($value) ? 'v' : 't';
          $newcell['valuetype'] = is_numeric($value) ? 'n' : 'th';
          $newcells[socialcalc_cr_to_coord($pos[0], $pos[1])] = $newcell;
          $pos[$normal]++;
        }
        $fields[$field_name]['endpos'] = array($pos[0]-$tangent_increment[$tangent], $pos[1]-$tangent_increment[$normal]);
        $lastpos = array(max($lastpos[0], $pos[0]), max($lastpos[1], $pos[1]));
      }

      // Replace formula placeholders with actual values.
      if (!empty($formulas)) foreach ($formulas as $formula) {
        $newcell = $formula['cell'];
        $expression = $formula['expression'];
        $matches = array();
        $count = (int)preg_match_all('/@row\.(\w+)@/', $formula['expression'], $matches);
        for ($i=0; $i<$count; $i++) {
          $field_name = $matches[1][$i];
          if (!isset($fields[$field_name])) { 
            $replace = 0;
          }
          else {
            $field = $fields[$field_name];
            $replace = socialcalc_cr_to_coord($field['cell']['pos'][0], $field['cell']['pos'][1]) .':'. socialcalc_cr_to_coord($field['endpos'][0], $field['endpos'][1]);
          }
          $expression = preg_replace('/@row\.(?:\w+)@/', $replace, $expression, 1);
        }
        $newcell['formula'] = $formula['formula'] .'('. $expression .')';
        $newcell['datatype'] = 'f';
        $newcell['valuetype'] = 'n';
        $newcell['datavalue'] = 0;
        $socialcalc['sheet']['cells'][$formula['coord']] = $newcell;
        $lastpos = array(max($lastpos[0], $newcell['pos'][0]), max($lastpos[1], $newcell['pos'][1]));
      }

      // Adjust positions of cells based on expanded values.
      foreach ($socialcalc['sheet']['cells'] as $coord => $cell) {
        $pos = $cell['pos'];
        foreach ($fields as $field_name => $field) {
          if ($pos[$tangent] == $field['cell']['pos'][$tangent]) { // supposed to be in the same column/row?
            if ($pos[$normal] >= $field['cell']['pos'][$normal] && $pos[$normal] <= $field['endpos'][$normal]) {
              $pos[$normal] += $field['endpos'][$normal] - $field['cell']['pos'][$normal];
              unset($socialcalc['sheet']['cells'][$coord]);
              $cell['pos'] = $pos;
              $socialcalc['sheet']['cells'][socialcalc_cr_to_coord($pos[0], $pos[1])] = $cell;
              $lastpos = array(max($lastpos[0], $pos[0]), max($lastpos[1], $pos[1]));
              break;
            }
          }
        }
      }

      // Replace values inside the sheet.
      $socialcalc['sheet']['cells'] += $newcells;
      $socialcalc['sheet']['attribs']['lastcol'] = $lastpos[0];
      $socialcalc['sheet']['attribs']['lastrow'] = $lastpos[1];
    }
    else {
      // Hand-make default SocialCalc structure based on views results.
      $pos = array(1,1);
      foreach ($this->view->field as $field => $info) {
        if ($info->options['exclude']) continue;
        $cell['pos'] = $pos;
        $cell['datavalue'] = $info->label();
        $cell['datatype'] = 't';
        $cell['valuetype'] = 't';

        $sheet['cells'][socialcalc_cr_to_coord($pos[0], $pos[1])] = $cell;
        $pos[$tangent]++;
      }
      $pos[$normal] = 2;
      foreach ($this->view->result as $result) {
        $pos[$tangent] = 1;
        foreach ($this->view->field as $field => $info) {
          if ($info->options['exclude']) continue;
          $cell['pos'] = $pos;
          $value = $info->theme($result);
          $cell['datavalue'] = $value;
          $cell['datatype'] = is_numeric($value) ? 'v' : 't';
          $cell['valuetype'] = is_numeric($value) ? 'n' : 'th';

          $sheet['cells'][socialcalc_cr_to_coord($pos[0], $pos[1])] = $cell;
          $pos[$tangent]++;
        }
        $pos[$normal]++;
      }
      $sheet['attribs']['lastcol'] = $pos[0]-1;
      $sheet['attribs']['lastrow'] = $pos[1]-1;

      if ($tangent == SHEETNODE_EXPANSION_VERTICAL) { 
        $edit['rowpanes'] = array(
          0 => array('first' => 1, 'last' => 1),
          1 => array('first' => 2, 'last' => $sheet['attribs']['lastrow']),
        );
        $edit['colpanes'] = array(
          0 => array('first' => 1, 'last' => $sheet['attribs']['lastcol']),
        );
      }
      else {
        $edit['colpanes'] = array(
          0 => array('first' => 1, 'last' => 1),
          1 => array('first' => 2, 'last' => $sheet['attribs']['lastcol']),
        );
        $edit['rowpanes'] = array(
          0 => array('first' => 1, 'last' => $sheet['attribs']['lastrow']),
        );
      }
      $edit['ecell'] = array(
        'coord' => 'A1',
      );

      // Inject the Sheetnode code.
      $socialcalc = array(
        'sheet' => $sheet,
        'edit' => $edit,
        'audit' => socialcalc_default_audit($sheet),
      );
    }
    return _sheetnode_inject(socialcalc_save($socialcalc), FALSE);
  }

  function get_real_field_name($field_name) {
    if (isset($this->view->field[$field_name])) return $field_name;
    foreach ($this->view->field as $field => $info) {
      if ($info->field_alias == $field_name) return $field;
      if (strcasecmp($info->label(), $field_name) == 0) return $field;
    }
    return FALSE;
  }
}

