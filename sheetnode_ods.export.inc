<?php
// $Id$

require_once(drupal_get_path('module', 'sheetnode') .'/socialcalc.inc');
$JAVA_BASE = variable_get('pjb_java_base', '');
require_once("$JAVA_BASE/java/Java.inc");
java_autoload(variable_get('sheetnode_ods_jar', ''));

function _sheetnode_ods_export($nid) {
  $node = node_load($nid);
  if (!$node || $node->type != 'sheetnode') return drupal_not_found();

  return _sheetnode_ods_export_do($node->title, socialcalc_parse(_sheetnode_load($node->nid, $node->vid)));
}

function _sheetnode_ods_cell_cmp($a, $b) {
  return $a['pos'][1] < $b['pos'][1] ? -1 : (
         $a['pos'][1] > $b['pos'][1] ?  1 : (
         $a['pos'][0] < $b['pos'][0] ? -1 : (
         $a['pos'][0] > $b['pos'][0] ?  1 : 0
  )));
}

function _sheetnode_ods_export_do($title, $socialcalc) {
  $sc = $socialcalc['sheet'];
  $sheetname = $title;
  $spreadsheet = org_odftoolkit_odfdom_doc_OdfSpreadsheetDocument::type()->newSpreadsheetDocument();
  $dom = $spreadsheet->getContentDom();
  $table = $dom->getElementsByTagNameNS(org_odftoolkit_odfdom_doc_table_OdfTable::type()->ELEMENT_NAME->getUri(), 
                                        org_odftoolkit_odfdom_doc_table_OdfTable::type()->ELEMENT_NAME->getLocalName())->item(0);
  $table->removeChild($table->getFirstChild()->getNextSibling()); // ODFDOM creates first row by default
  $table->setTableNameAttribute($sheetname);
  $root = $table->getParentNode();

  // Rows and columns
  $rows = array();
  for ($i=1; $i<=$sc['attribs']['lastrow']; $i++) {
    $row = $table->appendChild(new org_odftoolkit_odfdom_doc_table_OdfTableRow($dom));
    if (isset($sc['rowattribs']['height'][$i])) {
      $row->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTableRowProperties::type()->RowHeight, $sc['rowattribs']['height'][$i].'pt');
    }
    $rows[] = $row;
  }
  $cols = array();
  for ($i=1; $i<=$sc['attribs']['lastcol']; $i++) {
    $col = $i==1 ? $table->getTableColumn(0) : $table->addTableColumn(); // ODFDOM creates first column by default
    if (isset($sc['colattribs']['width'][$i])) {
      $col->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTableColumnProperties::type()->ColumnWidth, $sc['colattribs']['width'][$i].'px');
    }
    $cols[] = $col;
  }

  // Cells
  $rowNum = 0;
  uasort($sc['cells'], '_sheetnode_ods_cell_cmp');
  if ($sc['cells']) foreach ($sc['cells'] as $c) {
    if ($rowNum < $c['pos'][1]) { // new row
      $rowNum = $c['pos'][1];
      $row = $rows[$rowNum-1];
      $colNum = 1;
    }

    while ($colNum < $c['pos'][0]) { // append blank cells
      $row->appendChild(new org_odftoolkit_odfdom_doc_table_OdfTableCell($dom));
      $colNum++;
    }
    $cell = $row->appendChild(new org_odftoolkit_odfdom_doc_table_OdfTableCell($dom));
    
    // Value
    $p = $cell->appendChild(new org_odftoolkit_odfdom_doc_text_OdfTextParagraph($dom));
    $ovtav = new java_class('org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase$OfficeValueTypeAttributeValue');
    if (@$c['datatype'] == 't' || (@$c['datatype'] == 'f' && $c['valuetype'] != 'n')) {
      $cell->setOfficeValueTypeAttribute($ovtav->valueOf('STRING'));
      if ($c['valuetype'] == 'th') { // text-html
        $cell->setOfficeStringValueAttribute(strip_tags($c['datavalue']));
        $p->appendChild($dom->createTextNode(strip_tags($c['datavalue'])));
      }
      else {
        $cell->setOfficeStringValueAttribute($c['datavalue']);
        $p->appendChild($dom->createTextNode($c['datavalue']));
      }
    }
    else if (@$c['datatype'] == 'v' || (@$c['datatype'] == 'f' && $c['valuetype'] == 'n')) {
      if (!isset($c['nontextvalueformat'])) { // default number type is float but the format might dictate another type
        $cell->setOfficeValueTypeAttribute($ovtav->valueOf('FLOAT'));
        $cell->setOfficeValueAttribute((float)$c['datavalue']);
      }
      else {
        $formatinfo = socialcalc_parse_numberformat($sc['valueformats'][$c['nontextvalueformat']]);
        $sectioninfo = $formatinfo['sectioninfo'][0]; // TODO: Does it make sense to have sectioninfos with different cell types?
        if (!empty($sectioninfo['hascurrency'])) {
          $cell->setOfficeValueTypeAttribute($ovtav->valueOf('CURRENCY'));
          $cell->setOfficeValueAttribute((float)$c['datavalue']);
        }
        else if (!empty($sectioninfo['hasdate'])) {
          $cell->setOfficeValueTypeAttribute($ovtav->valueOf('DATE'));
          $cell->setOfficeDateValueAttribute($c['datavalue']);
        }
        else {
          $cell->setOfficeValueTypeAttribute($ovtav->valueOf('FLOAT'));
          $cell->setOfficeValueAttribute((float)$c['datavalue']);
        }
        // TODO: handle OdfPercentageStyle
      }
      $p->appendChild($dom->createTextNode($c['datavalue']));
    }
    if (@$c['datatype'] == 'f' && strpos($c['formula'], '!') === FALSE) {
      $cell->setTableFormulaAttribute('='.str_replace(',', ';', $c['formula']));
    }

    // Comment
    if (!empty($c['comment'])) {
      $note = $cell->appendChild(new org_odftoolkit_odfdom_doc_office_OdfOfficeAnnotation($dom));
      $notep = $note->appendChild(new org_odftoolkit_odfdom_doc_text_OdfTextParagraph($dom));
      $notep->appendChild($dom->createTextNode($c['comment']));
    }

    // Style
    _sheetnode_ods_export_style($c, $sc, $spreadsheet, $cell, $p);

    // Cell merges
    if (@$c['colspan'] > 1) $cell->setTableNumberColumnsSpannedAttribute($c['colspan']);
    if (@$c['rowspan'] > 1) $cell->setTableNumberRowsSpannedAttribute($c['rowspan']);
    
    $colNum++;
  }

  // Named expressions
  if (isset($sc['names'])) {
    $names = $root->appendChild(new org_odftoolkit_odfdom_doc_table_OdfTableNamedExpressions($dom));
    foreach ($sc['names'] as $n => $info) {
      $definition = $info['definition'];
      if ($definition[0] == '=') { // formula
        $definition = str_replace('(', "($'$sheetname'.", $definition);
        $definition = str_replace(',', ";$'$sheetname'.", $definition);
        $name = $names->appendChild(new org_odftoolkit_odfdom_doc_table_OdfTableNamedExpression($dom));
        $name->setTableExpressionAttribute($definition);
      }
      else {
        $definition = "$'$sheetname'.$definition";
        $name = $names->appendChild(new org_odftoolkit_odfdom_doc_table_OdfTableNamedRange($dom));
        $name->setTableCellRangeAddressAttribute($definition);
        $name->setTableBaseCellAddressAttribute($definition);
      }
      $name->setTableNameAttribute($n);
    }
  }
 
  // Write out to client.
  header('Content-type: application/vnd.oasis.opendocument.spreadsheet');
  header('Content-Disposition: attachment; filename='. _sheetnode_sanitize_filename($title, 'ods'));
  $memoryStream = new java_io_ByteArrayOutputStream();
  $spreadsheet->save($memoryStream);
  $memoryStream->close();
  echo java_values($memoryStream->toByteArray());
  exit();
}

function _sheetnode_ods_export_style($c, $sc, $spreadsheet, $cell, $p) {
  static $stylecount = 0;
  static $cellstyles = array();

  // Return if no formatting set.
  $cs = array_filter(array(
    isset($c['font']) ? 'f:'. $c['font'] : NULL,
    isset($c['color']) ? 'c:'. $c['color'] : NULL,
    isset($c['bt']) ? 'bt:'. $c['bt'] : NULL,
    isset($c['br']) ? 'br:'. $c['br'] : NULL,
    isset($c['bb']) ? 'bb:'. $c['bb'] : NULL,
    isset($c['bl']) ? 'bl:'. $c['bl'] : NULL,
    isset($c['layout']) ? 'l:'. $c['layout'] : NULL,
    isset($c['bgcolor']) ? 'bg:'. $c['bgcolor'] : NULL,
    isset($c['cellformat']) ? 'cf:'. $c['cellformat'] : NULL,
    isset($c['nontextvalueformat']) ? 'ntvf:'. $c['nontextvalueformat'] : NULL,
//    isset($c['textvalueformat']) ? 'tvf:'. $c['textvalueformat'] : NULL,
  ));
  if (empty($cs)) return;

  // Find cached cellstyle.
  $h = md5(implode(',', $cs));
  if (isset($cellstyles[$h])) {
    $cell->setTableStyleNameAttribute($cellstyles[$h]);
    return;
  }

  // Create new style
  $styles = $spreadsheet->getOrCreateDocumentStyles();
  $stylename = 'cell'.$stylecount++;
  $cellstyles[$h] = $stylename;
  $style = $styles->newStyle($stylename, org_odftoolkit_odfdom_dom_style_OdfStyleFamily::type()->TableCell);
  $cell->setTableStyleNameAttribute($stylename);
  
  // Horizontal alignment
  if (isset($c['cellformat'])) {
    $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleParagraphProperties::type()->TextAlign, $sc['cellformats'][$c['cellformat']]);
    $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTableCellProperties::type()->TextAlignSource, 'fix');
  }

  // Vertical alignment and padding
  $layout = socialcalc_cellformat_parselayout($c, $sc);
  if (isset($layout['alignvert'])) {
    $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTableCellProperties::type()->VerticalAlign, $layout['alignvert']);
  }
  $paddings = array(
    'padtop' => 'MarginTop',
    'padright' => 'MarginRight',
    'padbottom' => 'MarginBottom',
    'padleft' => 'MarginLeft',
  );
  foreach ($paddings as $pad => $property) {
    if (!empty($layout[$pad])) {
      $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleParagraphProperties::type()->$property, $layout[$pad].'px');
      $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleParagraphProperties::type()->TextAlign, 'start');
      $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTableCellProperties::type()->TextAlignSource, 'fix');
    }
  }

  // Background color
  $bg = socialcalc_cellformat_parsecolor($c, $sc, 'bgcolor');
  if ($bg) {
    $color = sprintf('#%02X%02X%02X', $bg['r'], $bg['g'], $bg['b']);
    $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTableCellProperties::type()->BackgroundColor, $color);
  }

  // Color
  $fg = socialcalc_cellformat_parsecolor($c, $sc, 'color');
  if ($fg) {
    $color = sprintf('#%02X%02X%02X', $fg['r'], $fg['g'], $fg['b']);
    $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTextProperties::type()->Color, $color);
  }

  // Font
  $ff = socialcalc_cellformat_parsefont($c, $sc);
  if (isset($ff['family'])) $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTextProperties::type()->FontFamily, $ff['family']);
  if (isset($ff['size'])) $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTextProperties::type()->FontSize, $ff['size']);
  if (!empty($ff['bold'])) $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTextProperties::type()->FontWeight, 'bold');
  if (!empty($ff['italic'])) $style->setProperty(org_odftoolkit_odfdom_doc_style_OdfStyleTextProperties::type()->FontStyle, 'italic');

  // Number format
  if (isset($c['nontextvalueformat'])) {
    _sheetnode_ods_export_numberformat($c, $sc, $spreadsheet, $cell, $p, $style);
  }
}

function _sheetnode_ods_export_numberformat($c, $sc, $spreadsheet, &$cell, $p, $style) {
  static $stylecount = 0;

  $dom = $spreadsheet->getStylesDom();
  $styles = $spreadsheet->getOrCreateDocumentStyles();
  $formatinfo = socialcalc_parse_numberformat($sc['valueformats'][$c['nontextvalueformat']]);
  $numsections = count($formatinfo['sectioninfo']);
  $conditions = array();
  $stylenames = array();
  foreach ($formatinfo['sectioninfo'] as $i => $sectioninfo) {
    $sectionend = ($i+1) == $numsections ? count($formatinfo['operators']) : $formatinfo['sectioninfo'][$i+1]['sectionstart']-1;
    $section = array_slice($formatinfo['operators'], $sectioninfo['sectionstart'], $sectionend - $sectioninfo['sectionstart']);
    $stylenames[$i] = 'number'. $stylecount .'n'. $i;

    // Create parent style
    if (!empty($sectioninfo['hascurrency'])) {
      $numstyle = new org_odftoolkit_odfdom_doc_number_OdfNumberCurrencyStyle($dom);
    }
    else if (!empty($sectioninfo['hasdate'])) {
      $numstyle = new org_odftoolkit_odfdom_doc_number_OdfNumberDateStyle($dom);
    }
    else {
      $numstyle = new org_odftoolkit_odfdom_doc_number_OdfNumberStyle($dom);
    }
    // TODO: handle OdfPercentageStyle
    $numstyle->setStyleNameAttribute($stylenames[$i]);

    // Populate number format
    $minintdigits = 0;
    $lastnumber = FALSE;
    $lasttext = '';
    $section[] = array(SOCIALCALC_COMMANDS_SECTION => ';'); // Manually add this to make sure we don't miss anything
    foreach ($section as $o) {
      $operator = key($o);
      $operand  = current($o);
      
      if ($lasttext && $operator != SOCIALCALC_COMMANDS_COPY) {
        $text = $numstyle->appendChild(new org_odftoolkit_odfdom_doc_number_OdfNumberText($dom));
        $text->setTextContent($lasttext);
        $lasttext = '';
      }
      if ($lastnumber && !in_array($operator, array(SOCIALCALC_COMMANDS_INTEGERPLACEHOLDER, SOCIALCALC_COMMANDS_FRACTIONPLACEHOLDER, SOCIALCALC_COMMANDS_DECIMAL))) {
        $number = $numstyle->appendChild(new org_odftoolkit_odfdom_doc_number_OdfNumber($dom));
        $number->setNumberGroupingAttribute($sectioninfo['thousandssep']);
        $number->setNumberDecimalPlacesAttribute($sectioninfo['fractiondigits']);
        $number->setNumberMinIntegerDigitsAttribute($minintdigits);
        $lastnumber = FALSE;
      }

      switch ($operator) {
      case SOCIALCALC_COMMANDS_CURRENCY:
        $currencysymbol = $numstyle->appendChild(new org_odftoolkit_odfdom_doc_number_OdfNumberCurrencySymbol($dom));
        $currencysymbol->appendChild($dom->createTextNode($operand));
        $currencysymbol->setNumberCountryAttribute('US'); // TODO
        $currencysymbol->setNumberRfcLanguageTagAttribute('en'); // TODO
        break;
      case SOCIALCALC_COMMANDS_COPY:
        $lasttext .= $operand;
        break;
      case SOCIALCALC_COMMANDS_INTEGERPLACEHOLDER:
        if ($operand == '0') $minintdigits++;
        // pass through
      case SOCIALCALC_COMMANDS_FRACTIONPLACEHOLDER:
      case SOCIALCALC_COMMANDS_DECIMAL:
        $lastnumber = TRUE;
        break;
      case SOCIALCALC_COMMANDS_COLOR: // TODO
        //$textprop = $numstyle->appendChild(new org_odftoolkit_odfdom_doc_style_OdfTextProperties($dom));
        //$numstyle->setProperty(org_odftoolkit_odfdom_doc_style_OdfTextProperties::type()->Color, $operand);
        break;
      case SOCIALCALC_COMMANDS_COMPARISON:
        $conditions[$i] = $operand;
        break;
      }
    }

    // Insert into spreadsheet
    $styles->appendChild($numstyle);
    if (($i+1) == $numsections) { // Last section connects to original cell style and to other sections
      $style->setStyleDataStyleNameAttribute($stylenames[$i]);
      for ($j=0; $j<$i; $j++) {
        $condition = isset($conditions[$j]) ? $conditions[$j] :
                     (($numsections == 2) ? '>=0' : 
                     (($numsections == 3 && $j == 0) ? '>0' :
                     (($numsections == 3 && $j == 1) ? '<0' : '=0'
                     )));
        $numstyle->newStyleMapElement($stylenames[$j], 'value()'.$condition);
      }
    }
    else { // Earlier sections get volatile=TRUE
      $numstyle->setStyleVolatileAttribute(TRUE);
    }
  }

  $stylecount++;
}

