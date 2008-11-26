<?php
// $Id$

require_once('socialcalc.inc');
require_once('Spreadsheet/Excel/Writer.php');

function _sheetnode_xls_export($nid) {
  $node = node_load($nid);
  if (!$node || $node->type != 'sheetnode') return drupal_not_found();

  // Create the Excel sheet.
  $workbook =& new Spreadsheet_Excel_Writer();
  $workbook->send(_sheetnode_sanitize_filename($node->title, 'xls', 'spreadsheet'));
  $worksheet =& $workbook->addWorksheet($node->title);

  $sheet = socialcalc_parse(_sheetnode_load($nid));
  if ($sheet['cells']) foreach ($sheet['cells'] as $cell) {
    $format = _sheetnode_xls_format($cell, $sheet, $workbook);
    $row = $cell['pos'][1]-1;
    $col = $cell['pos'][0]-1;
    if (isset($cell['formula'])) {
      $worksheet->writeFormula($row, $col, '='. $cell['formula'], $format !== FALSE ? $format : NULL);
    }
    else {
      $worksheet->write($row, $col, trim($cell['datavalue']), $format !== FALSE ? $format : NULL);
    }
    if (isset($cell['comment'])) {
      $worksheet->writeNote($row, $col, trim($cell['comment']));
    }  
  }
  if ($sheet['colattribs'] && $sheet['colattribs']['width']) foreach ($sheet['colattribs']['width'] as $col => $width) {
    // @see http://pear.php.net/manual/en/package.fileformats.spreadsheet-excel-writer.spreadsheet-excel-writer-worksheet.setcolumn.php#6454
    $worksheet->setColumn($col-1, $col-1, floatval($width) / WIDTH_CHAR_TO_PX); 
  }

  $workbook->close();
  exit();
}

function _sheetnode_xls_format($cell, $sheet, &$workbook) {
  static $formats = array();
  
  // Return if no formatting set.
  $f = array_filter(array(
    $cell['bt'] ? 'bt:'. $cell['bt'] : NULL, 
    $cell['br'] ? 'br:'. $cell['br'] : NULL, 
    $cell['bb'] ? 'bb:'. $cell['bb'] : NULL, 
    $cell['bl'] ? 'bl:'. $cell['bl'] : NULL,
    $cell['layout'] ? 'l:'. $cell['layout'] : NULL, 
    $cell['font'] ? 'f:'. $cell['font'] : NULL, 
    $cell['color'] ? 'c:'. $cell['color'] : NULL, 
    $cell['bgcolor'] ? 'bg:'. $cell['bgcolor'] : NULL,
    $cell['cellformat'] ? 'cf:'. $cell['cellformat'] : NULL, 
    $cell['nontextvalueformat'] ? 'ntvf:'. $cell['nontextvalueformat'] : NULL, 
    $cell['textvalueformat'] ? 'tvf:'. $cell['textvalueformat'] : NULL,
    $cell['colspan'] ? 'cs:'. $cell['colspan'] : NULL, 
    $cell['rowspan'] ? 'rs:'. $cell['rowspan'] : NULL, 
    $cell['cssc'] ? 'cssc:'. $cell['cssc'] : NULL, 
    $cell['csss'] ? 'csss:'. $cell['csss'] : NULL,
  ));
  if (empty($f)) return NULL;
  
  // Find cached format.
  $h = md5(implode(',', $f));
  if (isset($formats[$h])) return $formats[$h];
  
  // Create new format.
  $format =& $workbook->addFormat();
  $formats[$h] = $format;
  
  $font = socialcalc_cellformat_parsefont($cell, $sheet);
  if ($font['family']) $format->setFontFamily($font['family']);
  if ($font['size']) $format->setSize($font['size']);
  if ($font['bold']) $format->setBold();
  if ($font['italic']) $format->setItalic();

  $bgcolor = socialcalc_cellformat_parsecolor($cell, $sheet, 'bgcolor');
  if ($bgcolor) {
    $c = _sheetnode_xls_color($cell, $sheet, $workbook, $bgcolor);
    $format->setPattern(1); // solid color
    $format->setFgColor($c);
  }

  $color = socialcalc_cellformat_parsecolor($cell, $sheet, 'color');
  if ($color) {
    $c = _sheetnode_xls_color($cell, $sheet, $workbook, $color);
    $format->setColor($c);
  }

  return $format;
}

function _sheetnode_xls_color($cell, $sheet, &$workbook, $color) {
  /*
  static $colors = array();

  $h = md5(implode(':',$color));
  if (isset($colors[$h])) {
    $c = $colors[$h];
  }
  else {
    $c = 16 + count($colors); // http://pear.php.net/manual/en/package.fileformats.spreadsheet-excel-writer.spreadsheet-excel-writer-workbook.setcustomcolor.php
    $c = $workbook->setCustomColor($c, $color['r'], $color['g'], $color['b']);
    $colors[$h] = $c;
  }
  */
  
  $c = 0; // index of closest color
  $d = 500; // closest distance, start off at sqrt(255^2 + 255^2 + 255^2) which is largest distance
  foreach ($workbook->_palette as $i => $entry) {
    if ($color['r'] == $entry[0] && 
        $color['g'] == $entry[1] &&
        $color['b'] == $entry[2]) { // exact match: we found our color
      return $i+8;
    }
    
    // minimize distance
    $d1 = sqrt( pow($color['r'] - $entry[0], 2) + 
                pow($color['g'] - $entry[1], 2) +
                pow($color['b'] - $entry[2], 2) );
    if ($d1 < $d) { // this color is closer than the current one
      $d = $d1;
      $c = $i;
    }
  }
  return $c+8;
}

