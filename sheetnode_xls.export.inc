<?php
// $Id$

require_once('socialcalc.inc');
$JAVA_BASE = variable_get('pjb_java_base', NULL);
require_once("$JAVA_BASE/java/Java.inc");
java_autoload(variable_get('sheetnode_xls_jar', NULL));

function _sheetnode_xls_export($type, $nid) {
  $node = node_load($nid);
  if (!$node || $node->type != 'sheetnode') return drupal_not_found();

  // Create the Excel sheet
  $workbook = $type == 'xls' ? new org_apache_poi_hssf_usermodel_HSSFWorkbook() : new org_apache_poi_xssf_usermodel_XSSFWorkbook();
  $sheetname = $node->title;
  $sheet = $workbook->createSheet($sheetname);
  $patriarch = $sheet->createDrawingPatriarch();
  $sc = socialcalc_parse(_sheetnode_load($nid));
  
  // Names
  if (isset($sc['names'])) foreach ($sc['names'] as $n => $info) {
    $name = $workbook->createName();
    $name->setNameName($n);
    $definition = $info['definition'];
    if (strncmp('=', $definition, 1) == 0) { // formula
      $definition = substr($definition, 1);
      if (strpos($definition, '!') === FALSE) {
        $definition = str_replace('(', "('$sheetname'!", $definition);
        $definition = str_replace(',', ",'$sheetname'!", $definition);
      }
    }
    else if (strpos($definition, '!') === FALSE) {
      $definition = "'$sheetname'!$definition";
    }
    try { // Not 100% safe
      $name->setRefersToFormula($definition);
    }
    catch (Exception $e) {
      $msg = 'Failed to export name %name with definition %definition: %error';
      $arg = array('%name' => $n, '%definition' => $definition, '%error' => $e->getMessage());
      drupal_set_message(t($msg, $arg), 'error');
      watchdog('sheetnode', $msg, $arg, WATCHDOG_WARNING);
    }
  }

  if ($sc['cells']) foreach ($sc['cells'] as $c) {
    // Rows
    $row = $sheet->getRow($c['pos'][1]-1);
    if (java_is_null($row)) {
      $row = $sheet->createRow($c['pos'][1]-1);
      if (isset($sc['rowattribs']) && isset($sc['rowattribs']['height']) && isset($sc['rowattribs']['height'][$c['pos'][1]])) {
        $row->setHeightInPoints($sc['rowattribs']['height'][$c['pos'][1]]);
      }
    }

    // Cells
    $cell = $row->createCell($c['pos'][0]-1);
    _sheetnode_xls_export_cell($cell, $sheet, $patriarch, $workbook, $c, $sc); 
  }

  // Columns
  if (isset($sc['colattribs']) && isset($sc['colattribs']['width'])) {
    foreach ($sc['colattribs']['width'] as $col => $width) {
      $sheet->setColumnWidth($col-1, $width * 256 / WIDTH_CHAR_TO_PX);
    }
  }

  // Write and return the excel sheet to the client
  if ($type == 'xls') {
    header('Content-type: application/vnd.ms-excel');
  }
  else {
    header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }
  header('Content-Disposition: attachment; filename='. _sheetnode_sanitize_filename($node->title, $type));
  $memoryStream = new java_io_ByteArrayOutputStream();
  $workbook->write($memoryStream);
  $memoryStream->close();
  echo java_values($memoryStream->toByteArray());
  exit();
}

function _sheetnode_xls_export_cell($cell, $sheet, $patriarch, $workbook, $c, $sc) {
  // Cell data
  switch ($c['datatype']) {
  case 't':
    if ($c['valuetype'] == 'th') { // text-html
      $cell->setCellValue(strip_tags($c['datavalue']));
    }
    else {
      $cell->setCellValue($c['datavalue']);
    }
    $cell->setCellType(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_STRING);
    break;
  case 'v':
    $cell->setCellValue($c['datavalue']);
    $cell->setCellType(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_NUMERIC);
    break;
  case 'f':
    try { // Not 100% safe
      $cell->setCellFormula($c['formula']);
    }
    catch (Exception $e) {
      $msg = 'Failed to export formula %formula: %error';
      $arg = array('%formula' => $c['formula'], '%error' => $e->getMessage());
      drupal_set_message(t($msg, $arg), 'error');
      watchdog('sheetnode', $msg, $arg, $WATCHDOG_WARNING);
    }
    $cell->setCellType(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_FORMULA);
    break;
  }

  // Comment
  if ($c['comment']) {
    $string = $workbook->getCreationHelper()->createRichTextString($c['comment']);
    if (java_instanceof($workbook, org_apache_poi_hssf_usermodel_HSSFWorkbook::type())) { // HSSFWorkbook
      $anchor = $workbook->getCreationHelper()->createClientAnchor();
      $anchor->setCol1(java_values($cell->getColumnIndex())+1);
      $anchor->setRow1(java_values($cell->getRowIndex())+1);
      $anchor->setCol2(java_values($cell->getColumnIndex())+4);
      $anchor->setRow2(java_values($cell->getRowIndex())+4);
      $comment = $patriarch->createComment($anchor);
      $comment->setString($string);
      $cell->setCellComment($comment);
    }
    else { // XSSFWorkbook
//      $comment = $sheet->createComment();
    }
  }

  // Cellstyle
  _sheetnode_xls_export_cellstyle($cell, $sheet, $workbook, $c, $sc);

  // Merged regions
  if ($c['colspan'] > 1 || $c['rowspan'] > 1) {
    $cellrange = new org_apache_poi_ss_util_CellRangeAddress(
      java_values($cell->getRowIndex()), 
      java_values($cell->getRowIndex()) + max($c['rowspan']-1, 0),
      java_values($cell->getColumnIndex()), 
      java_values($cell->getColumnIndex()) + max($c['colspan']-1, 0)
    );
    $sheet->addMergedRegion($cellrange);
  }
}

function _sheetnode_xls_export_cellstyle($cell, $sheet, $workbook, $c, $sc) {
  static $cellstyles = array();
  
  // Return if no formatting set.
  $cs = array_filter(array(
    $c['font'] ? 'f:'. $c['font'] : NULL, 
    $c['color'] ? 'c:'. $c['color'] : NULL, 
    $c['bt'] ? 'bt:'. $c['bt'] : NULL, 
    $c['br'] ? 'br:'. $c['br'] : NULL, 
    $c['bb'] ? 'bb:'. $c['bb'] : NULL, 
    $c['bl'] ? 'bl:'. $c['bl'] : NULL,
    $c['layout'] ? 'l:'. $c['layout'] : NULL, 
    $c['bgcolor'] ? 'bg:'. $c['bgcolor'] : NULL,
    $c['cellformat'] ? 'cf:'. $c['cellformat'] : NULL, 
    $c['nontextvalueformat'] ? 'ntvf:'. $c['nontextvalueformat'] : NULL, 
    $c['textvalueformat'] ? 'tvf:'. $c['textvalueformat'] : NULL,
  ));
  if (empty($cs)) return;
  
  // Find cached cellstyle.
  $h = md5(implode(',', $cs));
  if (isset($cellstyles[$h])) {
    $cell->setCellStyle($cellstyles[$h]);
    return;
  }
  
  // Create new cellstyle.
  $cellstyle = $workbook->createCellStyle();
  $cellstyles[$h] = $cellstyle;

  // Font and color
  _sheetnode_xls_export_font($cellstyle, $cell, $sheet, $workbook, $c, $sc);
  
  // Background color
  $bg = socialcalc_cellformat_parsecolor($c, $sc, 'bgcolor');
  if ($bg) {
    $color = _sheetnode_xls_export_color($cell, $sheet, $workbook, $bg);
    $cellstyle->setFillPattern(1); // solid color
    $cellstyle->setFillForegroundColor($color);
  }

  // Horizontal alignment
  if ($c['cellformat']) {
    $alignments = array(
      'left' => org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_LEFT,
      'right' => org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_RIGHT,
      'center' => org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_CENTER,
    );
    $alignment = $alignments[$sc['cellformats'][$c['cellformat']]] | org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_GENERAL;
    $cellstyle->setAlignment($alignment);
  }

  // Vertical alignment and padding
  $layout = socialcalc_cellformat_parselayout($c, $sc);
  if ($layout) {
    if ($layout['padleft']) {
      $indent = intval(ceil($layout['padleft'] / WIDTH_CHAR_TO_PX)); 
      $cellstyle->setIndention($indent);
    }
    if ($layout['alignvert']) {
      $valignments = array(
        'top' => org_apache_poi_ss_usermodel_CellStyle::type()->VERTICAL_TOP,
        'bottom' => org_apache_poi_ss_usermodel_CellStyle::type()->VERTICAL_BOTTOM,
        'middle' => org_apache_poi_ss_usermodel_CellStyle::type()->VERTICAL_CENTER,
      );
      $valignment = $valignments[$layout['alignvert']] | org_apache_poi_ss_usermodel_CellStyle::type()->VERTICAL_JUSTIFY;
      $cellstyle->setVerticalAlignment($valignment);
    }
  }

  // Borders
  _sheetnode_xls_export_border($cellstyle, $cell, $sheet, $workbook, $c, $sc, 'Top');
  _sheetnode_xls_export_border($cellstyle, $cell, $sheet, $workbook, $c, $sc, 'Right');
  _sheetnode_xls_export_border($cellstyle, $cell, $sheet, $workbook, $c, $sc, 'Left');
  _sheetnode_xls_export_border($cellstyle, $cell, $sheet, $workbook, $c, $sc, 'Bottom');

  // Data format
  if ($c['nontextvalueformat']) {
    $dataformat = $workbook->createDataFormat();
    $format = $dataformat->getFormat($sc['valueformats'][$c['nontextvalueformat']]);
    $cellstyle->setDataFormat($format);
  }

  $cell->setCellStyle($cellstyle);
}

function _sheetnode_xls_export_font($cellstyle, $cell, $sheet, $workbook, $c, $sc) {
  static $fonts = array();

  // Return if no font set. 
  $f = array(
    $c['font'] ? 'f:'. $c['font'] : NULL, 
    $c['color'] ? 'c:'. $c['color'] : NULL, 
  );
  if (empty($f)) return;

  // Find cached font.
  $h = md5(implode(',', $f));
  if (isset($fonts[$h])) {
    $cellstyle->setFont($fonts[$h]);
    return;
  }

  // Create new font.
  $font = $workbook->createFont();
  $fonts[$h] = $font;

  $ff = socialcalc_cellformat_parsefont($c, $sc);
  if ($ff['family']) $font->setFontName($ff['family']);
  if ($ff['size']) $font->setFontHeightInPoints(intval($ff['size']));
  if ($ff['bold']) $font->setBoldweight(org_apache_poi_ss_usermodel_Font::type()->BOLDWEIGHT_BOLD);
  if ($ff['italic']) $font->setItalic(true);

  $fc = socialcalc_cellformat_parsecolor($c, $sc, 'color');
  if ($fc) {
    $color = _sheetnode_xls_export_color($cell, $sheet, $workbook, $fc);
    $font->setColor($color);
  }

  $cellstyle->setFont($font);
}

function _sheetnode_xls_export_border($cellstyle, $cell, $sheet, $workbook, $c, $sc, $pos) {
  $attribs = array(
    'Top' => 'bt',
    'Right' => 'br',
    'Bottom' => 'bb',
    'Left' => 'bl',
  );
  $border = socialcalc_cellformat_parseborder($c, $sc, $attribs[$pos]);
  if ($border) {
    if ($border['thickness'] || $border['style']) {
      $thickness = str_replace('px', '', $border['thickness']); // TODO: what about other units?
      if ($thickness > 0 && $thickness < 7) {
        $thickness = 'thin';
      }
      else if ($thickness > 7 && $thickness < 15) {
        $thickness = 'medium';
      }
      else if ($thickness > 15) {
        $thickness = 'thick';
      }
      $borderstyles = array(
        'thin' => array(
          'solid' => org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_THIN,
          'dashed' => org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_DASHED,
          'dotted' => org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_DOTTED,
          'double' => org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_DOUBLE,
        ),
        'medium' => array(
          'solid' => org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_MEDIUM,
          'dashed' => org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_MEDIUM_DASHED,
        ),
        'thick' => array(
          'solid' => org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_THICK,
        ),
      );
      // TODO: what about other combinations?
      $borderstyle = $borderstyles[$thickness][$border['style']] | org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_NONE;
      $setstyle = "setBorder{$pos}";
      $cellstyle->$setstyle($borderstyle);
    }
    if ($border['color']) {
      $color = _sheetnode_xls_export_color($cell, $sheet, $workbook, $border['color']);
      $setcolor = "set{$pos}BorderColor";
      $cellstyle->$setcolor($color);
    }
  }
}

function _sheetnode_xls_export_color($cell, $sheet, $workbook, $c) {
  if (java_instanceof($workbook, org_apache_poi_hssf_usermodel_HSSFWorkbook::type())) { // HSSFWorkbook
    static $used_palette = array();

    $palette = $workbook->getCustomPalette();
    $color = $palette->findColor($c['r'], $c['g'], $c['b']);
    if (java_is_null($color)) {
      // Replace an unused existing color with the new one.
      for ($i=0x8; $i<0x40; $i++) {
        if ($used_palette[$i]) continue;
        $palette->setColorAtIndex($i, $c['r'], $c['g'], $c['b']);
        $used_palette[$i] = true;
        return $i;
      }

      // Could not find an empty slot: find closest color.
      $color = $palette->findSimilarColor($c['r'], $c['g'], $c['b']);
      if (java_is_null($color)) {
        // Could not even find a close color: return default color.
        return org_apache_poi_ss_usermodel_Font::type()->COLOR_NORMAL;
      }
    }

    $used_palette[java_values($color->getIndex())] = true;
    return $color->getIndex();
  }
  else { // XSSFWorkbook
    $color = new org_apache_poi_xssf_usermodel_XSSFColor();
    $argb = array(1, $c['r'], $c['g'], $c['b']);
    $color->setRgb($argb);
    return $color;
  }
}

