<?php
// $Id$

require_once('socialcalc.inc');
$JAVA_BASE = drupal_get_path('module', 'sheetnode').'/META-INF/java';
require_once('META-INF/java/Java.inc');
java_autoload(drupal_get_path('module', 'sheetnode').'/poi-3.2-FINAL-20081019.jar');

function _sheetnode_xls_export($nid) {
  $node = node_load($nid);
  if (!$node || $node->type != 'sheetnode') return drupal_not_found();

  header('Content-type: application/vnd.ms-excel');
  header('Content-Disposition: attachment; filename='. _sheetnode_sanitize_filename($node->title, 'xls'));

  // Create the Excel sheet
  $workbook = new org_apache_poi_hssf_usermodel_HSSFWorkbook();
  $sheet = $workbook->createSheet($node->title);
  $patriarch = $sheet->createDrawingPatriarch();

  $sc = socialcalc_parse(_sheetnode_load($nid));
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
    $cell->setCellType(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_STRING);
    $cell->setCellValue($c['datavalue']);
    break;
  case 'v':
    $cell->setCellType(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_NUMERIC);
    $cell->setCellValue($c['datavalue']);
    break;
  case 'f':
    $cell->setCellType(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_FORMULA);
    $cell->setCellFormula($c['formula']);
    break;
  }

  // Comment
  if ($c['comment']) {
    $comment = $patriarch->createComment(
      new org_apache_poi_hssf_usermodel_HSSFClientAnchor(0, 0, 0, 0, 
        java_values($cell->getColumnIndex())+1, 
        java_values($cell->getRowIndex())+1, 
        java_values($cell->getColumnIndex())+4, 
        java_values($cell->getRowIndex())+4));
    $string = new org_apache_poi_hssf_usermodel_HSSFRichTextString($c['comment']);
    $comment->setString($string);
    $cell->setCellComment($comment);
  }

  // Cellstyle
  _sheetnode_xls_export_cellstyle($cell, $sheet, $workbook, $c, $sc);
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
      'left' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_LEFT,
      'right' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_RIGHT,
      'center' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_CENTER,
    );
    $alignment = $alignments[$sc['cellformats'][$c['cellformat']]];
    if (!$alignment) $alignment = org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_GENERAL;
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
        'top' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->VERTICAL_TOP,
        'bottom' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->VERTICAL_BOTTOM,
        'middle' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->VERTICAL_CENTER,
      );
      $valignment = $valignments[$layout['alignvert']];
      if (!$valignment) $valignment = org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->VERTICAL_JUSTIFY;
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
  if ($ff['bold']) $font->setBoldweight(org_apache_poi_hssf_usermodel_HSSFFont::type()->BOLDWEIGHT_BOLD);
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
      $thickness = str_replace('px', '', $border['thickness']);
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
          'solid' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_THIN,
          'dashed' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_DASHED,
          'dotted' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_DOTTED,
          'double' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_DOUBLE,
        ),
        'medium' => array(
          'solid' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_MEDIUM,
          'dashed' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_MEDIUM_DASHED,
        ),
        'thick' => array(
          'solid' => org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_THICK,
        ),
      );
      $borderstyle = $borderstyles[$thickness][$border['style']];
      if (!$borderstyle) $borderstyle = org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_NONE;
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
      return org_apache_poi_hssf_usermodel_HSSFFont::type()->COLOR_NORMAL;
     }
  }

  $used_palette[java_values($color->getIndex())] = true;
  return $color->getIndex();
}

