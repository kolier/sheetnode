<?php
// $Id$

require_once('socialcalc.inc');
$JAVA_BASE = variable_get('pjb_java_base', '');
require_once("$JAVA_BASE/Java.inc");
java_autoload(variable_get('sheetnode_ods_jar', ''));

function _sheetnode_ods_export($nid) {
  $node = node_load($nid);
  if (!$node || $node->type != 'sheetnode') return drupal_not_found();
 
  $sc = socialcalc_parse(_sheetnode_load($nid));
  $sheetname = $node->title;
  $spreadsheet = org_openoffice_odf_doc_OdfSpreadsheetDocument::type()->createSpreadsheetDocument();
  $dom = $spreadsheet->getContentDom();
  $table = $dom->getElementsByTagNameNS(org_openoffice_odf_doc_element_table_OdfTable::type()->ELEMENT_NAME->getUri(), 
                                        org_openoffice_odf_doc_element_table_OdfTable::type()->ELEMENT_NAME->getLocalName())->item(0);
  $table->removeChild($table->getFirstChild()->getNextSibling()); // ODFDOM creates first row by default
  $table->setName($sheetname);
  $root = $table->getParentNode();

  // Rows and columns
  $rows = array();
  for ($i=1; $i<=$sc['attribs']['lastrow']; $i++) {
    $row = $table->appendChild(new org_openoffice_odf_doc_element_table_OdfTableRow($dom));
    if (isset($sc['rowattribs']['height'][$i])) {
      $row->setProperty(org_openoffice_odf_doc_element_style_OdfTableRowProperties::type()->RowHeight, 
        org_openoffice_odf_dom_util_OdfUnitConverter::type()->mapUnitToCm($sc['rowattribs']['height'][$i].'pt'));
    }
    $rows[] = $row;
  }
  $cols = array();
  for ($i=1; $i<=$sc['attribs']['lastcol']; $i++) {
    $col = $i==1 ? $table->getTableColumn(0) : $table->addTableColumn(); // ODFDOM creates first column by default
    if (isset($sc['colattribs']['width'][$i])) {
      $col->setProperty(org_openoffice_odf_doc_element_style_OdfTableColumnProperties::type()->ColumnWidth, 
        org_openoffice_odf_dom_util_OdfUnitConverter::type()->mapUnitToCm($sc['colattribs']['width'][$i].'px'));
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
      $row->appendChild(new org_openoffice_odf_doc_element_table_OdfTableCell($dom));
      $colNum++;
    }
    $cell = $row->appendChild(new org_openoffice_odf_doc_element_table_OdfTableCell($dom));
    
    // Value
    $p = $cell->appendChild(new org_openoffice_odf_doc_element_text_OdfParagraph($dom));
    switch ($c['datatype']) {
    case 't':
      $cell->SetValueType(org_openoffice_odf_dom_type_office_OdfValueType::type()->valueOf('STRING'));
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
      $cell->SetValueType(org_openoffice_odf_dom_type_office_OdfValueType::type()->valueOf('FLOAT'));
      $cell->SetValue($c['datavalue']);
      $p->appendChild($dom->createTextNode($c['datavalue']));
      break;
    case 'f':
      $cell->SetFormula('='.str_replace(',', ';', $c['formula']));
      break;
    }

    // Comment
    if ($c['comment']) {
      $note = $cell->appendChild(new org_openoffice_odf_doc_element_office_OdfAnnotation($dom));
      $notep = $note->appendChild(new org_openoffice_odf_doc_element_text_OdfParagraph($dom));
      $notep->appendChild($dom->createTextNode($c['comment']));
    }

    // Style
    _sheetnode_ods_export_style($c, $sc, $spreadsheet, $dom, $cell, $p);

    // Number format
    if ($c['nontextvalueformat']) {
      _sheetnode_ods_export_numberformat($c, $sc, $spreadsheet, $dom, $cell, $p);
    }

    // Cell merges
    if ($c['colspan'] > 1) $cell->setNumberColumnsSpanned($c['colspan']);
    if ($c['rowspan'] > 1) $cell->setNumberRowsSpanned($c['rowspan']);
    
    $colNum++;
  }

  // Named expressions
  if (isset($sc['names'])) {
    $names = $root->appendChild(new org_openoffice_odf_doc_element_table_OdfNamedExpressions($dom));
    foreach ($sc['names'] as $n => $info) {
      $definition = $info['definition'];
      if ($definition[0] == '=') { // formula
        $definition = str_replace('(', "($'$sheetname'.", $definition);
        $definition = str_replace(',', ";$'$sheetname'.", $definition);
        $name = $names->appendChild(new org_openoffice_odf_doc_element_table_OdfNamedExpression($dom));
        $name->setExpression($definition);
      }
      else {
        $definition = "$'$sheetname'.$definition";
        $name = $names->appendChild(new org_openoffice_odf_doc_element_table_OdfNamedRange($dom));
        $name->setCellRangeAddress($definition);
        $name->setBaseCellAddress($definition);
      }
      $name->setName($n);
    }
  }
/*
  // Write out to client.
  header('Content-type: application/vnd.oasis.opendocument.spreadsheet');
  header('Content-Disposition: attachment; filename='. _sheetnode_sanitize_filename($node->title, 'ods'));
  $memoryStream = new java_io_ByteArrayOutputStream();
  $spreadsheet->save($memoryStream);
  $memoryStream->close();
  echo java_values($memoryStream->toByteArray());
  exit();
*/
}

function _sheetnode_ods_export_numberformat($c, $sc, $spreadsheet, $dom, $cell, $p) {
  var_dump($sc['valueformats'][$c['nontextvalueformat']]);
  $formatinfo = socialcalc_parse_numberformat($sc['valueformats'][$c['nontextvalueformat']]);
  foreach ($formatinfo['sectioninfo'] as $i => $sectioninfo) {
    var_dump($sectioninfo);
    echo '<br/>';
    $sectionend = ($i+1) == count($formatinfo['sectioninfo']) ? count($formatinfo['operators']) : $formatinfo['sectioninfo'][$i+1]['sectionstart']-1;
    for ($j=$sectioninfo['sectionstart']; $j<$sectionend; $j++) {
      var_dump($formatinfo['operators'][$j]);
    }
    echo '<p/>';
  }
}

function _sheetnode_ods_export_style($c, $sc, $spreadsheet, $dom, $cell, $p) {
  static $stylecount = 0;
 
  // Create new style
  $styles = $spreadsheet->getOrCreateDocumentStyles();
  $stylename = 'cell'.$stylecount;
  $style = $styles->createStyle($stylename, org_openoffice_odf_dom_style_OdfStyleFamily::type()->TableCell);
  $cell->setStyleName($stylename);
  
  // Horizontal alignment
  if ($c['cellformat']) {
    $style->setProperty(org_openoffice_odf_doc_element_style_OdfParagraphProperties::type()->TextAlign, $sc['cellformats'][$c['cellformat']]);
    $style->setProperty(org_openoffice_odf_doc_element_style_OdfTableCellProperties::type()->TextAlignSource, 'fix');
  }

  // Vertical alignment and padding
  $layout = socialcalc_cellformat_parselayout($c, $sc);
  if ($layout['alignvert']) {
    $style->setProperty(org_openoffice_odf_doc_element_style_OdfTableCellProperties::type()->VerticalAlign, $layout['alignvert']);
  }
  $paddings = array(
    'padtop' => 'MarginTop',
    'padright' => 'MarginRight',
    'padbottom' => 'MarginBottom',
    'padleft' => 'MarginLeft',
  );
  foreach ($paddings as $pad => $property) {
    if ($layout[$pad]) {
      $style->setProperty(org_openoffice_odf_doc_element_style_OdfParagraphProperties::type()->$property, 
        org_openoffice_odf_dom_util_OdfUnitConverter::type()->mapUnitToCm($layout[$pad].'px'));
      $style->setProperty(org_openoffice_odf_doc_element_style_OdfParagraphProperties::type()->TextAlign, 'start');
      $style->setProperty(org_openoffice_odf_doc_element_style_OdfTableCellProperties::type()->TextAlignSource, 'fix');
    }
  }

  // Background color
  $bg = socialcalc_cellformat_parsecolor($c, $sc, 'bgcolor');
  if ($bg) {
    $color = sprintf('#%02X%02X%02X', $bg['r'], $bg['g'], $bg['b']);
    $style->setProperty(org_openoffice_odf_doc_element_style_OdfTableCellProperties::type()->BackgroundColor, $color);
  }

  // Color
  $fg = socialcalc_cellformat_parsecolor($c, $sc, 'color');
  if ($fg) {
    $color = sprintf('#%02X%02X%02X', $fg['r'], $fg['g'], $fg['b']);
    $style->setProperty(org_openoffice_odf_doc_element_style_OdfTextProperties::type()->Color, $color);
  }

  // Font
  $ff = socialcalc_cellformat_parsefont($c, $sc);
  if ($ff['family']) $style->setProperty(org_openoffice_odf_doc_element_style_OdfTextProperties::type()->FontFamily, $ff['family']);
  if ($ff['size']) $style->setProperty(org_openoffice_odf_doc_element_style_OdfTextProperties::type()->FontSize, $ff['size']);
  if ($ff['bold']) $style->setProperty(org_openoffice_odf_doc_element_style_OdfTextProperties::type()->FontWeight, 'bold');
  if ($ff['italic']) $style->setProperty(org_openoffice_odf_doc_element_style_OdfTextProperties::type()->FontStyle, 'italic');

  $stylecount++;
}

