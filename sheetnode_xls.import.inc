<?php
// $Id$

function _sheetnode_xls_import_form(&$form_state) {
  $form['#attributes'] = array('enctype' => "multipart/form-data");
  $form['import'] = array(
    '#type' => 'file',
    '#title' => t('XLS file to import'),
    '#size' => 40,
  );
  $form['submit'] = array(
    '#type' => 'submit',
    '#value' => t('Submit'),
  );
  return $form;
}

require_once('socialcalc.inc');
$JAVA_BASE = drupal_get_path('module', 'sheetnode').'/META-INF/java';
require_once('META-INF/java/Java.inc');
java_autoload(drupal_get_path('module', 'sheetnode').'/poi-3.2-FINAL-20081019.jar');

function _sheetnode_xls_import_form_validate($form, &$form_state) {
  $source = $_FILES['files']['name']['import'];

  // Handle PHP file upload errors.
  switch ($_FILES['files']['error']['import']) {
  case UPLOAD_ERR_OK:
    break;
  case UPLOAD_ERR_FORM_SIZE:
  case UPLOAD_ERR_INI_SIZE:
    form_set_error('import', t('The file %file could not be saved, because it exceeds %maxsize, the maximum allowed size for uploads. Please choose a smaller file or notify your site administrator.', 
      array('%file' => $source, '%maxsize' => format_size(file_upload_max_size()))));
    return;
  case UPLOAD_ERR_PARTIAL:
    form_set_error('import', t('The file %file could not be saved, because the upload did not complete. Please try again or notify your site administrator.', array('%file' => $source)));
    return;
  case UPLOAD_ERR_NO_FILE:
    form_set_error('import', t('No file was uploaded. Please choose a file to upload.'));
    return;
  case UPLOAD_ERR_NO_TMP_DIR:
    form_set_error('import', t('The file %file could not be saved, because the temporary upload folder is missing. Please notify your site administrator.', array('%file' => $source)));
    return;
  case UPLOAD_ERR_CANT_WRITE:
    form_set_error('import', t('The file %file could not be saved, because a failure occurred while writing it to the temporary folder. Please notify your site administrator.', array('%file' => $source)));
    return;
  case UPLOAD_ERR_EXTENSION:
    form_set_error('import', t('The file %file could not be saved, because its extension is disallowed. Please notify your site administrator.', array('%file' => $source)));
    return;
  default:
    form_set_error('import', t('The file %file could not be saved. An unknown error %error has occurred. Please notify your site administrator.', 
      array('%file' => $source, '%error' => $_FILES['files']['error']['import'])));
    return;
  }
  
  // Verify that the uploaded file is indeed XLS.
  try {
    $inp = new java_io_FileInputStream($_FILES['files']['tmp_name']['import']);
    $wb = new org_apache_poi_hssf_usermodel_HSSFWorkbook(new org_apache_poi_poifs_filesystem_POIFSFileSystem($inp));
  }
  catch (Exception $e) {
    form_set_error('import', t('The file %file appears to be invalid: %error. Please upload a valid Excel (.xls) file.', array('%file' => $source, '%error' => $e->getMessage())));
  }
}

// @see http://drupal.org/node/293663
function _sheetnode_xls_import_form_submit($form, &$form_state) {
  global $user;
  $form_state_node = array();
  module_load_include('inc', 'node', 'node.pages');
  $node = array('type' => 'sheetnode');
  $form_state_node['values']['title'] = pathinfo($_FILES['files']['name']['import'], PATHINFO_FILENAME);
  $form_state_node['values']['body'] = '';
  $form_state_node['values']['name'] = $user->name;
  $form_state_node['values']['sheetsave'] = _sheetnode_xls_import_do($_FILES['files']['tmp_name']['import']);
  $form_state_node['values']['op'] = t('Save');
  drupal_execute('sheetnode_node_form', $form_state_node, (object)$node);
  $form_state['redirect'] = isset($form_state_node['nid']) ? 'node/'. $form_state_node['nid'] .'/edit' : '';
}

function _sheetnode_xls_import_do($file) {
  $inp = new java_io_FileInputStream($_FILES['files']['tmp_name']['import']);
  $wb = new org_apache_poi_hssf_usermodel_HSSFWorkbook(new org_apache_poi_poifs_filesystem_POIFSFileSystem($inp));
  $sheet = $wb->getSheetAt(0);
  $sc = array(); // SocialCalc array structure

  // Default values.
  $sc['attribs']['defaultcolwidth'] = round(java_values($sheet->getDefaultColumnWidth()) * WIDTH_CHAR_TO_PX, 2); // width in chars * factor
  $sc['attribs']['defaultrowheight'] = round(java_values($sheet->getDefaultRowHeightInPoints()), 2);
  $sc['attribs']['lastrow'] = java_values($sheet->getLastRowNum())+1;
  $sc['attribs']['lastcol'] = 1; // POI doesn't tell us last col, so we'll compute it

  // Fonts
  for ($i=0; $i <= java_values($wb->getNumberOfFonts()); $i++) { // notice the <= 
    _sheetnode_xls_import_font($wb, $wb->getFontAt($i), $sc);
  }

  // Cell styles
  for ($i=0; $i < java_values($wb->getNumCellStyles()); $i++) {
    _sheetnode_xls_import_cellstyle($wb, $wb->getCellStyleAt($i), $sc);
  }

  // Rows
  for ($rit = $sheet->rowIterator(); java_is_true($rit->hasNext()); ) {
    $row = $rit->next();
    $rowNum = intval(java_values($row->getRowNum()));
    $height = round(java_values($row->getHeightInPoints()), 2);
    if ($height != $sc['attribs']['defaultrowheight']) {
      $sc['rowattribs']['height'][$rowNum+1] = $height;
    }

    // Columns and cells
    for ($cit = $row->cellIterator(); java_is_true($cit->hasNext()); ) {
      $cell = $cit->next();

      $colNum = intval(java_values($cell->getColumnIndex()));
      $width = round(java_values($sheet->getColumnWidth($colNum)) * WIDTH_CHAR_TO_PX / 256, 2); // width in 1/256 of a char!! 
      if ($width != $sc['attribs']['defaultcolwidth']) {
        $sc['colattribs']['width'][$colNum+1] = $width;
      }
      if (java_is_true($sheet->isColumnHidden($colNum))) {
        $sc['colattribs']['hide'][$colNum+1] = 'yes';
      }
      
      // Last column
      if ($colNum >= $sc['attribs']['lastcol']) $sc['attribs']['lastcol'] = $colNum+1;

      _sheetnode_xls_import_cell($wb, $cell, $sc);
    }
  }

  // Cell merges
  for ($i=0; $i < java_values($sheet->getNumMergedRegions()); $i++) {
    $range = $sheet->getMergedRegion($i);
    $c = java_values($range->getFirstColumn())+1;
    $r = java_values($range->getFirstRow())+1;
    $coord = socialcalc_cr_to_coord($c,$r);
    if (!isset($sc['cells'][$coord])) continue;
    $sc['cells'][$coord]['colspan'] = java_values($range->getLastColumn()) - java_values($range->getFirstColumn()) + 1;
    $sc['cells'][$coord]['rowspan'] = java_values($range->getLastRow()) - java_values($range->getFirstRow()) + 1;
  }

  // Names
  for ($i=0; $i < java_values($wb->getNumberOfNames()); $i++) {
    $name = $wb->getNameAt($i);
    $sc['names'][java_values($name->getNameName())] = array(
      'desc' => '',
      'definition' => (java_is_true($name->isFunctionName()) ? '=' : ''). java_values($name->getReference()),
    );
  }

  return socialcalc_save($sc);
}

function _sheetnode_xls_import_cell($wb, $cell, &$sc) {
  $type = java_values($cell->getCellType());
  $c = array(); // output cell
  
  // Position
  $pos = array(java_values($cell->getColumnIndex())+1, java_values($cell->getRowIndex())+1);
  $c['pos'] = $pos;

  // Value and type
  switch ($type) {
  case java_values(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_BOOLEAN):
    $c['datavalue'] = java_values($cell->getBooleanCellValue());
    $c['datatype'] = 'v';
    $c['valuetype'] = 'n';
    break;
  case java_values(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_NUMERIC):
    $c['datavalue'] = java_values($cell->getNumericCellValue());
    $c['datatype'] = 'v';
    $c['valuetype'] = 'n';
    break;
  case java_values(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_STRING):
    $c['datavalue'] = java_values($cell->getStringCellValue());
    $c['datatype'] = 't';
    $c['valuetype'] = 't';
    break;
  case java_values(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_FORMULA):
    $c['formula'] = java_values($cell->getCellFormula());
    $c['datatype'] = 'f';
    switch (java_values($cell->getCachedFormulaResultType())) {
    case java_values(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_BOOLEAN):
      $c['datavalue'] = java_values($cell->getBooleanCellValue());
      $c['valuetype'] = 'n';
      break;
    case java_values(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_NUMERIC):
      $c['datavalue'] = java_values($cell->getNumericCellValue());
      $c['valuetype'] = 'n';
      break;
    case java_values(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_STRING):
      $c['datavalue'] = java_values($cell->getStringCellValue());
      $c['valuetype'] = 't';
      break;
    default:
      $c['datavalue'] = 0;
      $c['valuetype'] = 'n';
      break;
    }
    break;
  case java_values(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_ERROR):
    $c['errors'] = java_values(org_apache_poi_hssf_usermodel_FormulaError::forInt(java_values($cell->getErrorCellValue()))->getString());
    break;
  case java_values(org_apache_poi_hssf_usermodel_HSSFCell::type()->CELL_TYPE_BLANK):
    break;
  }

  // Comment
  $comment = $cell->getCellComment();
  if (!java_is_null($comment)) {
    $c['comment'] = java_values($comment->getString()->getString());
  }

  // Style
  $cellstyle = $cell->getCellStyle();
  if (!java_is_null($cellstyle)) { 

    // Font
    $font = $cellstyle->getFont($wb);
    $fonthash = $sc['fonthash'][_sheetnode_xls_get_font($font)];
    if ($fonthash != $sc['attribs']['defaultfont']) $c['font'] = $fonthash;
    
    // Borders
    $bt = _sheetnode_xls_get_borderstyle($wb, $cellstyle, 'Top');
    if ($bt) $c['bt'] = $sc['borderstylehash'][$bt];
    $bl = _sheetnode_xls_get_borderstyle($wb, $cellstyle, 'Left');
    if ($bl) $c['bl'] = $sc['borderstylehash'][$bl];
    $bb = _sheetnode_xls_get_borderstyle($wb, $cellstyle, 'Bottom');
    if ($bb) $c['bb'] = $sc['borderstylehash'][$bb];
    $br = _sheetnode_xls_get_borderstyle($wb, $cellstyle, 'Right');
    if ($br) $c['br'] = $sc['borderstylehash'][$br];

    // Cell format (horizontal alignment)
    $c['cf'] = $sc['cellformathash'][_sheetnode_xls_get_cellformat($cellstyle)];

    // Layout (padding + vertical alignment)
    $layouthash = $sc['layouthash'][_sheetnode_xls_get_layout($cellstyle)];
    if ($layouthash != $sc['attribs']['defaultlayout']) $c['layout'] = $layouthash;

    // Color
    $color = _sheetnode_xls_get_color($wb, java_values($font->getColor()));
    if ($color) {
      $colorhash = $sc['colorhash'][$color];
      if ($colorhash != $sc['attribs']['defaultcolor']) $c['color'] = $colorhash;
    }

    // Background color
    $bgcolor = _sheetnode_xls_get_color($wb, java_values($cellstyle->getFillForegroundColor()));
    if ($bgcolor) {
      $bgcolorhash = $sc['colorhash'][$bgcolor];
      if ($bgcolorhash != $sc['attribs']['defaultbgcolor']) $c['bgcolor'] = $bgcolorhash;
    }
    
    // Non-text value format
    $ntvf = _sheetnode_xls_get_valueformat(java_values($cellstyle->getDataFormatString()));
    if ($ntvf) {
      $ntvfhash = $sc['valueformathash'][$ntvf];
      if ($ntvfhash != $sc['attribs']['defaultnontextvalueformat']) $c['nontextvalueformat'] = $ntvfhash;
    }
  }

  $sc['cells'][socialcalc_cr_to_coord($pos[0],$pos[1])] = $c;
}

function _sheetnode_xls_import_font($wb, $font, &$sc) {
  $fontstyle = _sheetnode_xls_get_font($font);
  if (!isset($sc['fonthash'][$fontstyle])) {
    $index = count($sc['fonts'])+1;
    $sc['fonts'][$index] = $fontstyle;
    $sc['fonthash'][$fontstyle] = $index;

    // Default font
    if (!isset($sc['attribs']['defaultfont'])) $sc['attribs']['defaultfont'] = $index;
  }
  
  $color = _sheetnode_xls_get_color($wb, java_values($font->getColor()));
  if ($color && !isset($sc['colorhash'][$color])) {
    $index = count($sc['colors'])+1;
    $sc['colors'][$index] = $color;
    $sc['colorhash'][$color] = $index;
  }
}

function _sheetnode_xls_import_cellstyle($wb, $style, &$sc) {
  // Border styles
  _sheetnode_xls_import_borderstyle($wb, $style, 'Top', $sc);
  _sheetnode_xls_import_borderstyle($wb, $style, 'Right', $sc);
  _sheetnode_xls_import_borderstyle($wb, $style, 'Left', $sc);
  _sheetnode_xls_import_borderstyle($wb, $style, 'Bottom', $sc);

  // Cell format (horizontal alignment)
  $cellformat = _sheetnode_xls_get_cellformat($style);
  if (!isset($sc['cellformathash'][$cellformat])) {
    $index = count($sc['cellformats'])+1;
    $sc['cellformats'][$index] = $cellformat;
    $sc['cellformathash'][$cellformat] = $index;
  }

  // Layout (padding + vertical alignment)
  $layout = _sheetnode_xls_get_layout($style);
  if (!isset($sc['layouthash'][$layout])) {
    $index = count($sc['layouts'])+1;
    $sc['layouts'][$index] = $layout;
    $sc['layouthash'][$layout] = $index;

    // Default layout
    if (!isset($sc['attribs']['defaultlayout'])) $sc['attribs']['defaultlayout'] = $index;
  }

  // Background color
  $bgcolor = _sheetnode_xls_get_color($wb, java_values($style->getFillForegroundColor()));
  if ($bgcolor && !isset($sc['colorhash'][$bgcolor])) {
    $index = count($sc['colorhash'])+1;
    $sc['colors'][$index] = $bgcolor;
    $sc['colorhash'][$bgcolor] = $index;
  }

  // Value format
  $valueformat = _sheetnode_xls_get_valueformat(java_values($style->getDataFormatString()));
  if ($valueformat && !isset($sc['valueformathash'][$valueformat])) {
    $index = count($sc['valueformats'])+1;
    $sc['valueformats'][$index] = $valueformat;
    $sc['valueformathash'][$valueformat] = $index;
  }
}

function _sheetnode_xls_import_borderstyle($wb, $style, $pos, &$sc) {
  $borderstyle = _sheetnode_xls_get_borderstyle($wb, $style, $pos);
  if ($borderstyle && !isset($sc['borderstylehash'][$borderstyle])) {
    $index = count($sc['borderstyles'])+1;
    $sc['borderstyles'][$index] = $borderstyle;
    $sc['borderstylehash'][$borderstyle] = $index;
  }
}

function _sheetnode_xls_get_font($font) {
  return (java_is_true($font->getItalic()) ? 'italic' : 'normal') .' '.
         (java_values($font->getBoldWeight()) == java_values(org_apache_poi_hssf_usermodel_HSSFFont::type()->BOLDWEIGHT_BOLD) ? 'bold' : 'normal') .' '.
          java_values($font->getFontHeightInPoints()) .'pt '.
          java_values($font->getFontName());
}

function _sheetnode_xls_get_borderstyle($wb, $style, $pos) {
  $borderstyles = array(
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_NONE) => FALSE, // TODO: is it?
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_THIN) => 'thin solid', 
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_MEDIUM) => 'medium solid', 
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_DASHED) => 'thin dashed', 
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_DOTTED) => 'thin dotted', 
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_THICK) => 'thick solid', 
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_DOUBLE) => 'thin double', 
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_HAIR) => 'thin dotted',  // TODO: is it?
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_MEDIUM_DASHED) => 'medium dashed', 
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_DASH_DOT) => 'thin dashed', // TODO: is it?
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_MEDIUM_DASH_DOT) => 'medium dashed', //TODO: is it?
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_DASH_DOT_DOT) => 'thin dashed', // TODO: is it?
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_MEDIUM_DASH_DOT_DOT) => 'medium dashed', // TODO: is it?
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->BORDER_SLANTED_DASH_DOT) => 'thin dashed', // TODO: is it?
  );
  $getstyle = "getBorder{$pos}";
  $borderstyle = $borderstyles[java_values($style->$getstyle())];
  if (!$borderstyle) return FALSE;
  $getcolor = "get{$pos}BorderColor";
  $bordercolor = _sheetnode_xls_get_color($wb, java_values($style->$getcolor()));
  return $borderstyle .' '. $bordercolor;
}

function _sheetnode_xls_get_cellformat($style) {
  $cellformats = array(
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_GENERAL) => 'left', // TODO: is it?
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_LEFT) => 'left',
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_CENTER) => 'center',
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_RIGHT) => 'right',
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_FILL) => 'justify', // TODO: is it?
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_JUSTIFY) => 'justify',
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->ALIGN_CENTER_SELECTION) => 'center', // TODO: is it?
  );
  return $cellformats[java_values($style->getAlignment())];
}

function _sheetnode_xls_get_layout($style) {
  $valignments = array(
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->VERTICAL_TOP) => 'top',
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->VERTICAL_CENTER) => 'middle',
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->VERTICAL_BOTTOM) => 'bottom', 
    java_values(org_apache_poi_hssf_usermodel_HSSFCellStyle::type()->VERTICAL_JUSTIFY) => 'middle', // TODO: is it?
  );
  $valignment = $valignments[java_values($style->getVerticalAlignment())];
  $padding_left = java_values($style->getIndention()) * WIDTH_CHAR_TO_PX; // simulate indentation with padding
  return 'padding:* * * '. $padding_left .'px;vertical-align:'. $valignment .';';
}

function _sheetnode_xls_get_valueformat($dataformat) {
  if (strtolower($dataformat) == 'general') return FALSE;
  return $dataformat;
}

function _sheetnode_xls_get_color($wb, $index) {
  if ($index == java_values(org_apache_poi_hssf_usermodel_HSSFFont::type()->COLOR_NORMAL)) return FALSE;
  $automatic = new java_class('org.apache.poi.hssf.util.HSSFColor$AUTOMATIC');
  if ($index == java_values($automatic->index)) return FALSE;
  $palette = $wb->getCustomPalette();
  $color = $palette->getColor($index);
  $triplet = java_values($color->getTriplet());
  return 'rgb('. $triplet[0] .','. $triplet[1] .','. $triplet[2] .')';
}

