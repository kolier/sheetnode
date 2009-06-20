<?php
// $Id$

require_once('socialcalc.inc');
$JAVA_BASE = variable_get('pjb_java_base', '');
require_once("$JAVA_BASE/java/Java.inc");
java_autoload(variable_get('sheetnode_ods_jar', ''));

function _sheetnode_ods_export($nid) {
  $node = node_load($nid);
  if (!$node || $node->type != 'sheetnode') return drupal_not_found();
 
  $socialcalc = socialcalc_parse(_sheetnode_load($nid));
  $sc = $socialcalc['sheet'];
  $sheetname = $node->title;
  $spreadsheet = org_odftoolkit_odfdom_doc_OdfSpreadsheetDocument::type()->createSpreadsheetDocument();
  $dom = $spreadsheet->getContentDom();
  $table = $dom->getElementsByTagNameNS(org_odftoolkit_odfdom_doc_element_table_OdfTable::type()->ELEMENT_NAME->getUri(), 
                                        org_odftoolkit_odfdom_doc_element_table_OdfTable::type()->ELEMENT_NAME->getLocalName())->item(0);
  $table->removeChild($table->getFirstChild()->getNextSibling()); // ODFDOM creates first row by default
  $table->setName($sheetname);
  $root = $table->getParentNode();

  // Rows and columns
  $rows = array();
  for ($i=1; $i<=$sc['attribs']['lastrow']; $i++) {
    $row = $table->appendChild(new org_odftoolkit_odfdom_doc_element_table_OdfTableRow($dom));
    if (isset($sc['rowattribs']['height'][$i])) {
      $row->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTableRowProperties::type()->RowHeight, 
        org_odftoolkit_odfdom_dom_util_OdfUnitConverter::type()->mapUnitToCm($sc['rowattribs']['height'][$i].'pt'));
    }
    $rows[] = $row;
  }
  $cols = array();
  for ($i=1; $i<=$sc['attribs']['lastcol']; $i++) {
    $col = $i==1 ? $table->getTableColumn(0) : $table->addTableColumn(); // ODFDOM creates first column by default
    if (isset($sc['colattribs']['width'][$i])) {
      $col->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTableColumnProperties::type()->ColumnWidth, 
        org_odftoolkit_odfdom_dom_util_OdfUnitConverter::type()->mapUnitToCm($sc['colattribs']['width'][$i].'px'));
    }
    $cols[] = $col;
  }

  // Cells
  $rowNum = 0;
  if ($sc['cells']) foreach ($sc['cells'] as $c) {
    if ($rowNum < $c['pos'][1]) { // new row
      $rowNum = $c['pos'][1];
      $row = $rows[$rowNum-1];
      $colNum = 1;
    }

    while ($colNum < $c['pos'][0]) { // append blank cells
      $row->appendChild(new org_odftoolkit_odfdom_doc_element_table_OdfTableCell($dom));
      $colNum++;
    }
    $cell = $row->appendChild(new org_odftoolkit_odfdom_doc_element_table_OdfTableCell($dom));
    
    // Value
    $p = $cell->appendChild(new org_odftoolkit_odfdom_doc_element_text_OdfParagraph($dom));
    if (isset($c['datatype'])) switch ($c['datatype']) {
    case 't':
      $cell->SetValueType(org_odftoolkit_odfdom_dom_type_office_OdfValueType::type()->valueOf('STRING'));
      if ($c['valuetype'] == 'th') { // text-html
        $cell->setStringValue(strip_tags($c['datavalue']));
        $p->appendChild($dom->createTextNode(strip_tags($c['datavalue'])));
      }
      else {
        $cell->setStringValue($c['datavalue']);
        $p->appendChild($dom->createTextNode($c['datavalue']));
      }
      break;
    case 'v':
      $cell->SetValueType(org_odftoolkit_odfdom_dom_type_office_OdfValueType::type()->valueOf('FLOAT'));
      $cell->SetValue($c['datavalue']);
      $p->appendChild($dom->createTextNode($c['datavalue']));
      break;
    case 'f':
      $cell->SetFormula('='.str_replace(',', ';', $c['formula']));
      break;
    }

    // Comment
    if (!empty($c['comment'])) {
      $note = $cell->appendChild(new org_odftoolkit_odfdom_doc_element_office_OdfAnnotation($dom));
      $notep = $note->appendChild(new org_odftoolkit_odfdom_doc_element_text_OdfParagraph($dom));
      $notep->appendChild($dom->createTextNode($c['comment']));
    }

    // Style
    $style = _sheetnode_ods_export_style($c, $sc, $spreadsheet, $cell, $p);

    // Number format
    if (isset($c['nontextvalueformat'])) {
      _sheetnode_ods_export_numberformat($c, $sc, $spreadsheet, $cell, $p, $style);
    }

    // Cell merges
    if (@$c['colspan'] > 1) $cell->setNumberColumnsSpanned($c['colspan']);
    if (@$c['rowspan'] > 1) $cell->setNumberRowsSpanned($c['rowspan']);
    
    $colNum++;
  }

  // Named expressions
  if (isset($sc['names'])) {
    $names = $root->appendChild(new org_odftoolkit_odfdom_doc_element_table_OdfNamedExpressions($dom));
    foreach ($sc['names'] as $n => $info) {
      $definition = $info['definition'];
      if ($definition[0] == '=') { // formula
        $definition = str_replace('(', "($'$sheetname'.", $definition);
        $definition = str_replace(',', ";$'$sheetname'.", $definition);
        $name = $names->appendChild(new org_odftoolkit_odfdom_doc_element_table_OdfNamedExpression($dom));
        $name->setExpression($definition);
      }
      else {
        $definition = "$'$sheetname'.$definition";
        $name = $names->appendChild(new org_odftoolkit_odfdom_doc_element_table_OdfNamedRange($dom));
        $name->setCellRangeAddress(new org_odftoolkit_odfdom_dom_type_OdfCellRangeAddress($definition));
        $name->setBaseCellAddress(new org_odftoolkit_odfdom_dom_type_OdfCellAddress($definition));
      }
      $name->setName($n);
    }
  }
  
  // Write out to client.
  header('Content-type: application/vnd.oasis.opendocument.spreadsheet');
  header('Content-Disposition: attachment; filename='. _sheetnode_sanitize_filename($node->title, 'ods'));
  $memoryStream = new java_io_ByteArrayOutputStream();
  $spreadsheet->save($memoryStream);
  $memoryStream->close();
  echo java_values($memoryStream->toByteArray());
  exit();
}

function _sheetnode_ods_export_numberformat($c, $sc, $spreadsheet, $cell, $p, $style) {
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
      $numstyle = new org_odftoolkit_odfdom_doc_element_number_OdfCurrencyStyle($dom);
    }
    else if (!empty($sectioninfo['hasdate'])) {
      $numstyle = new org_odftoolkit_odfdom_doc_element_number_OdfDateStyle($dom);
    }
    else {
      $numstyle = new org_odftoolkit_odfdom_doc_element_number_OdfNumberStyle($dom);
    }
    // TODO: handle OdfPercentageStyle.
    $numstyle->setName(new org_odftoolkit_odfdom_dom_type_OdfStyleName($stylenames[$i]));

    // Populate number format
    $minintdigits = 0;
    $lastnumber = false;
    $lasttext = '';
    $section[] = array(SOCIALCALC_COMMANDS_SECTION => ';'); // Manually add this to make sure we don't miss anything
    foreach ($section as $o) {
      $operator = key($o);
      $operand  = current($o);
      
      if ($lasttext && $operator != SOCIALCALC_COMMANDS_COPY) {
        $text = $numstyle->appendChild(new org_odftoolkit_odfdom_doc_element_number_OdfText($dom));
        $text->setTextContent($lasttext);
        $lasttext = '';
      }
      if ($lastnumber && !in_array($operator, array(SOCIALCALC_COMMANDS_INTEGERPLACEHOLDER, SOCIALCALC_COMMANDS_FRACTIONPLACEHOLDER, SOCIALCALC_COMMANDS_DECIMAL))) {
        $number = $numstyle->appendChild(new org_odftoolkit_odfdom_doc_element_number_OdfNumber($dom));
        $number->setGrouping(new org_odftoolkit_odfdom_dom_type_OdfBoolean($sectioninfo['thousandssep']));
        $number->setDecimalPlaces($sectioninfo['fractiondigits']);
        $number->setMinIntegerDigits($minintdigits);
        $lastnumber = false;
      }

      switch ($operator) {
      case SOCIALCALC_COMMANDS_CURRENCY:
        $currencysymbol = $numstyle->appendChild(new org_odftoolkit_odfdom_doc_element_number_OdfCurrencySymbol($dom));
        $currencysymbol->appendChild($dom->createTextNode($operand));
        $currencysymbol->setCountry('US'); // TODO
        $currencysymbol->setLanguage('en'); // TODO
        break;
      case SOCIALCALC_COMMANDS_COPY:
        $lasttext .= $operand;
        break;
      case SOCIALCALC_COMMANDS_INTEGERPLACEHOLDER:
        if ($operand == '0') $minintdigits++;
        // pass through
      case SOCIALCALC_COMMANDS_FRACTIONPLACEHOLDER:
      case SOCIALCALC_COMMANDS_DECIMAL:
        $lastnumber = true;
        break;
      case SOCIALCALC_COMMANDS_COLOR: // TODO
        //$textprop = $numstyle->appendChild(new org_odftoolkit_odfdom_doc_element_style_OdfTextProperties($dom));
        //$numstyle->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTextProperties::type()->Color, $operand);
        break;
      case SOCIALCALC_COMMANDS_COMPARISON:
        $conditions[$i] = $operand;
        break;
      }
    }

    // Insert into spreadsheet
    $styles->appendChild($numstyle);
    if (($i+1) == $numsections) { // Last section connects to original cell style and to other sections
      $style->setDataStyleName(new org_odftoolkit_odfdom_dom_type_OdfStyleNameRef($stylenames[$i]));
      for ($j=0; $j<$i; $j++) {
        $map = $numstyle->appendChild(new org_odftoolkit_odfdom_doc_element_style_OdfMap($dom));
        $map->setApplyStyleName(new org_odftoolkit_odfdom_dom_type_OdfStyleNameRef($stylenames[$j]));
        $condition = isset($conditions[$j]) ? $conditions[$j] :
                     (($numsections == 2) ? '>=0' : 
                     (($numsections == 3 && $j == 0) ? '>0' :
                     (($numsections == 3 && $j == 1) ? '<0' : '=0'
                     )));
        $map->setCondition('value()'.$condition);
      }
    }
    else { // Earlier sections get volatile=true
      $numstyle->setVolatile(new org_odftoolkit_odfdom_dom_type_OdfBoolean(true));
    }
  }

  $stylecount++;
}

function _sheetnode_ods_export_style($c, $sc, $spreadsheet, $cell, $p) {
  static $stylecount = 0;
 
  // Create new style
  $styles = $spreadsheet->getOrCreateDocumentStyles();
  $stylename = 'cell'.$stylecount++;
  $style = $styles->createStyle($stylename, org_odftoolkit_odfdom_dom_style_OdfStyleFamily::type()->TableCell);
  $cell->setStyleName($stylename);
  
  // Horizontal alignment
  if (isset($c['cellformat'])) {
    $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfParagraphProperties::type()->TextAlign, $sc['cellformats'][$c['cellformat']]);
    $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTableCellProperties::type()->TextAlignSource, 'fix');
  }

  // Vertical alignment and padding
  $layout = socialcalc_cellformat_parselayout($c, $sc);
  if (isset($layout['alignvert'])) {
    $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTableCellProperties::type()->VerticalAlign, $layout['alignvert']);
  }
  $paddings = array(
    'padtop' => 'MarginTop',
    'padright' => 'MarginRight',
    'padbottom' => 'MarginBottom',
    'padleft' => 'MarginLeft',
  );
  foreach ($paddings as $pad => $property) {
    if (isset($layout[$pad])) {
      $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfParagraphProperties::type()->$property, 
        org_odftoolkit_odfdom_dom_util_OdfUnitConverter::type()->mapUnitToCm($layout[$pad].'px'));
      $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfParagraphProperties::type()->TextAlign, 'start');
      $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTableCellProperties::type()->TextAlignSource, 'fix');
    }
  }

  // Background color
  $bg = socialcalc_cellformat_parsecolor($c, $sc, 'bgcolor');
  if ($bg) {
    $color = sprintf('#%02X%02X%02X', $bg['r'], $bg['g'], $bg['b']);
    $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTableCellProperties::type()->BackgroundColor, $color);
  }

  // Color
  $fg = socialcalc_cellformat_parsecolor($c, $sc, 'color');
  if ($fg) {
    $color = sprintf('#%02X%02X%02X', $fg['r'], $fg['g'], $fg['b']);
    $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTextProperties::type()->Color, $color);
  }

  // Font
  $ff = socialcalc_cellformat_parsefont($c, $sc);
  if (isset($ff['family'])) $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTextProperties::type()->FontFamily, $ff['family']);
  if (isset($ff['size'])) $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTextProperties::type()->FontSize, $ff['size']);
  if (!empty($ff['bold'])) $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTextProperties::type()->FontWeight, 'bold');
  if (!empty($ff['italic'])) $style->setProperty(org_odftoolkit_odfdom_doc_element_style_OdfTextProperties::type()->FontStyle, 'italic');

  return $style;
}

