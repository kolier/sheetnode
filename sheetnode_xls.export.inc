<?php
// $Id$

function _sheetnode_xls_export($type, $nid) {
  require_once(drupal_get_path('module', 'sheetnode') . '/socialcalc.inc');

  $node = node_load($nid);
  if (!$node || $node->type != 'sheetnode') return drupal_not_found();

  return _sheetnode_xls_export_do($type, $node->title, socialcalc_parse(_sheetnode_load($node->nid, $node->vid)));
}

function _sheetnode_xls_export_do($type, $title, $socialcalc) {
  require_once(variable_get('sheetnode_xls_library_path', '') . '/Classes/PHPExcel.php');

  // Create new PHPExcel object
  $sc = $socialcalc['sheet'];
  $workbook = new PHPExcel();

  // Title
  try {
    // TODO: Split titles with format "Spreadsheet > Workbook" and sanitize.
    $workbook->getActiveSheet()->SetTitle($title);
  }
  catch (Exception $e) {
    watchdog('sheetnode_xls', 'Error setting worksheet title to "!title": @error', array('!title' => $title, '@error' => $e->getMessage()), WATCHDOG_WARNING);
  }

/*
  // Properties
  $workbook->getProperties()->setCreator("Maarten Balliauw")
                 ->setLastModifiedBy("Maarten Balliauw")
                 ->setTitle("Office 2007 XLSX Test Document")
                 ->setSubject("Office 2007 XLSX Test Document")
                 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
                 ->setKeywords("office 2007 openxml php")
                 ->setCategory("Test result file");
*/
  // Names
  // TODO: Not supported in Excel5.
  if (!empty($sc['names'])) foreach ($sc['names'] as $n => $info) {
    // TODO: Support external references.
    $workbook->addNamedRange(new PHPExcel_NamedRange($n, $workbook->getActiveSheet(), $info['definition']));
  }

  // Cells
  if ($sc['cells']) foreach ($sc['cells'] as $c) {
    $value = _sheetnode_xls_cell_value($c, $sc); 
    $coord = socialcalc_cr_to_coord($c['pos'][0], $c['pos'][1]);
    $workbook->getActiveSheet()->setCellValue($coord, $value);

    // Comment
    // TODO: Not supported in Excel5.
    if (!empty($c['comment'])) {
      $workbook->getActiveSheet()->getComment($coord)->getText()->createTextRun($c['comment']);
    }

    // Style
    $style = $workbook->getActiveSheet()->getStyle($coord);
    // Font
    if (!empty($c['font'])) {
      $scf = socialcalc_cellformat_parsefont($c, $sc);
      $font = $style->getFont();
      if (!empty($scf['family'])) $font->setName($scf['family']);
      if (!empty($scf['size'])) $font->setSize(floatval($scf['size']));
      if (!empty($scf['bold'])) $font->setBold(true);
      if (!empty($scf['italic'])) $font->setItalic(true);
      if (!empty($c['color'])) {
        $scc = socialcalc_cellformat_parsecolor($c, $sc, 'color');
        $rgb = sprintf('%02X%02X%02X', $scc['r'], $scc['g'], $scc['b']);
        $font->getColor()->setRGB($rgb);
      }
    }

    // Borders
    foreach (array(
      'br' => 'getRight', 
      'bl' => 'getLeft', 
      'bt' => 'getTop', 
      'bb' => 'getBottom',
    ) as $pos => $method) {
      $border = socialcalc_cellformat_parseborder($c, $sc, $pos);
      if (empty($border)) continue;
      $borderobj = $style->getBorders()->$method();
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
          'solid' => PHPExcel_Style_Border::BORDER_THIN,
          'dashed' => PHPExcel_Style_Border::BORDER_DASHED,
          'dotted' => PHPExcel_Style_Border::BORDER_DOTTED,
          'double' => PHPExcel_Style_Border::BORDER_DOUBLE,
        ),
        'medium' => array(
          'solid' => PHPExcel_Style_Border::BORDER_MEDIUM,
          'dashed' => PHPExcel_Style_Border::BORDER_MEDIUMDASHED,
        ),
        'thick' => array(
          'solid' => PHPExcel_Style_Border::BORDER_THICK,
        ),
      );
      // TODO: what about other combinations?
      $borderstyle = isset($borderstyles[$thickness][$border['style']]) ? $borderstyles[$thickness][$border['style']] : PHPExcel_Style_Border::BORDER_THIN;
      $borderobj->setBorderStyle($borderstyle);
      $scc = $border['color'];
      $rgb = sprintf('%02X%02X%02X', $scc['r'], $scc['g'], $scc['b']);
      $borderobj->getColor()->setRGB($rgb);
    }

    // Background color
    if (!empty($c['bgcolor'])) {
      $scc = socialcalc_cellformat_parsecolor($c, $sc, 'bgcolor');
      $rgb = sprintf('%02X%02X%02X', $scc['r'], $scc['g'], $scc['b']);
      $style->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
      $style->getFill()->getStartColor()->setRGB($rgb);
    }

    // Alignment
    if (!empty($c['cellformat'])) {
      $alignments = array(
        'left' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
        'right' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
        'center' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
      );
      $alignment = isset($alignments[$sc['cellformats'][$c['cellformat']]]) ? $alignments[$sc['cellformats'][$c['cellformat']]] : NULL;
      if ($alignment) {
        $style->getAlignment()->setHorizontal($alignment);
      }
    }

    // Vertical Alignment
    $layout = socialcalc_cellformat_parselayout($c, $sc);
    if (!empty($layout['alignvert'])) {
      $valignments = array(
        'top' => PHPExcel_Style_Alignment::VERTICAL_TOP,
        'bottom' => PHPExcel_Style_Alignment::VERTICAL_BOTTOM,
        'middle' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
      );
      $valignment = isset($valignments[$layout['alignvert']]) ? $valignments[$layout['alignvert']] : NULL;
      if ($valignment) {
        $style->getAlignment()->setVertical($valignment);
      }
    }

    // Data format
    if (!empty($c['nontextvalueformat'])) {
      $style->getNumberFormat()->setFormatCode($sc['valueformats'][$c['nontextvalueformat']]);
    }

    // Merged regions
    if ($c['colspan'] > 1 || $c['rowspan'] > 1) {
      $coord2 = socialcalc_cr_to_coord($c['pos'][0]+max($c['colspan']-1, 0), $c['pos'][1]+max($c['rowspan']-1, 0));
      $workbook->getActiveSheet()->mergeCells($coord .':'. $coord2);
    }
  }

  // Columns
  $font = $workbook->getActiveSheet()->getDefaultStyle()->getFont();
  if (!empty($sc['colattribs']['width'])) foreach($sc['colattribs']['width'] as $c => $width) {
    $workbook->getActiveSheet()->getColumnDimension(PHPExcel_Cell::stringFromColumnIndex($c-1))->setWidth(PHPExcel_Shared_Drawing::pixelsToCellDimension($width, $font));
  }

  // Rows
  if (!empty($sc['rowattribs']['height'])) foreach($sc['rowattribs']['height'] as $r => $height) {
    $workbook->getActiveSheet()->getRowDimension($col)->setRowHeight($height);
  }

  // Redirect output to a client’s web browser.
  $types = array(
    'xls' => array(
      'content-type' => 'application/vnd.ms-excel',
      'php-excel-type' => 'Excel5',
    ),
    'xlsx' => array(
      'content-type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'php-excel-type' => 'Excel2007',
    ),
  );
  header('Content-Type: ' . $types[$type]['content-type']);
  header('Content-Disposition: attachment;filename="' . _sheetnode_sanitize_filename($title, $type) . '"');
  header('Cache-Control: max-age=0');

  $writer = PHPExcel_IOFactory::createWriter($workbook, $types[$type]['php-excel-type']);
  $writer->save('php://output');
  exit;
}

function _sheetnode_xls_cell_value($c, $sc) {
  if ($c['datatype'] == 't' || ($c['datatype'] == 'f' && $c['valuetype'] != 'n')) {
    // TODO: Preserve multiline strings.
    if ($c['valuetype'] == 'th') { // text-html
      $value = strip_tags($c['datavalue']);
    }
    else {
      $value = $c['datavalue'];
    }
  }
  else if ($c['datatype'] == 'v' || ($c['datatype'] == 'f' && $c['valuetype'] == 'n')) {
    $value = floatval($c['datavalue']);
  }
  if ($c['datatype'] == 'f' && strpos($c['formula'], '!') === FALSE) {
    $value = '=' . $c['formula'];
  }
  return $value;
}

