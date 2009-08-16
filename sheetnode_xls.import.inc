<?php
// $Id$

function _sheetnode_xls_import_form(&$form_state) {
  $form['#attributes'] = array('enctype' => "multipart/form-data");
  $form['import'] = array(
    '#type' => 'file',
    '#title' => t('Excel spreadsheet (.xls or .xlsx) to import'),
    '#size' => 40,
  );
  $form['submit'] = array(
    '#type' => 'submit',
    '#value' => t('Submit'),
  );
  return $form;
}

require_once(drupal_get_path('module', 'sheetnode') .'/socialcalc.inc');
$JAVA_BASE = variable_get('pjb_java_base', NULL);
require_once("$JAVA_BASE/java/Java.inc");
java_autoload(variable_get('sheetnode_xls_jar', NULL));

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
    $workbook = org_apache_poi_ss_usermodel_WorkbookFactory::type()->create($inp);
  }
  catch (Exception $e) {
    form_set_error('import', t('The file %file appears to be invalid: %error. Please upload a valid Excel (.xls or .xlsx) file.', array('%file' => $source, '%error' => $e->getMessage())));
  }
}

abstract class SheetnodeXLS {
  public static function create($workbook) {
    if (java_instanceof($workbook, org_apache_poi_hssf_usermodel_HSSFWorkbook::type())) {
      return new SheetnodeHSSF($workbook);
    }
    else {
      return new SheetnodeXSSF($workbook);
    }
  }

  protected $_properties;
  public function getProperties($field = null) {
    if (!$this->_properties) {
      $this->_properties = $this->importProperties();
    }
    return $field ? $this->_properties[$field]['value'] : $this->_properties;
  }

  abstract public function getFont($cellstyle);
  abstract public function getFontOfFormattingRun($rtf, $run);
  abstract public function getColor($font);
  abstract public function getFillForegroundColor($cellstyle);
  abstract public function getBorderColor($cellstyle, $pos);
  abstract public function importFormula($cell, &$c);
  abstract public function importColor($color);
  abstract protected function importProperties();

  private $_cellstyles = array();
  public function existsCellstyle($cellstyle) {
    $index = java_values($cellstyle->getIndex()); 
    if (isset($this->_cellstyles[$index])) return true;
    $this->_cellstyles[$index] = true;
    return false;
  }

  private $_fonts = array();
  public function existsFont($font) {
    $index = java_values($font->getIndex());
    if (isset($this->_fonts[$index])) return true;
    $this->_fonts[$index] = true;
    return false;
  }

  protected $_workbook;
}

class SheetnodeHSSF extends SheetnodeXLS {
  public function SheetnodeHSSF($workbook) {
    $this->_workbook = $workbook;
  }
  protected function importProperties() {
    $properties = array();
    $si = $this->_workbook->getSummaryInformation();
    $dsi = $this->_workbook->getDocumentSummaryInformation();
    $fields_objects = array(
      array('object' => $si, 'fields' => array(
        'Title' => t('Title'),
        'Author' => t('Author'),
        'Subject' => t('Subject'),
        'Keywords' => t('Keywords'),
        'Comments' => array('label' => t('Body'), 'field' => 'Body'),
        'CreateDateTime' => t('Creation date'),
        'LastSaveDateTime' => t('Last modification date'),
        'EditTime' => array('label' => t('Total time spent on editing'), 'null' => 0),
        'LastAuthor' => t('Last modified by'),
        'LastPrinted' => t('Last printing date'),
        'PageCount' => array('label' => t('Page count'), 'null' => 0),
        'WordCount' => array('label' => t('Word count'), 'null' => 0),
        'CharCount' => array('label' => t('Character count'), 'null' => 0),
        'RevNumber' => t('Revision number'),
        'Template' => t('Template'),
        'ApplicationName' => t('Application'),
        'Security' => array('label' => t('Security'), 'map' => array(
          0 => null,
          1 => t('Password protected'),
          2 => t('Read-only recommended'),
          4 => t('Read-only enforced'),
          8 => t('Locked for annotations'),
        )),
      )),
      array('object' => $dsi, 'fields' => array(
        'Category' => t('Category'),
        'Company' => t('Company'),
        'Manager' => t('Manager'),
      )),
    );
    foreach ($fields_objects as $propertyset) {
      $object = $propertyset['object'];
      foreach ($propertyset['fields'] as $field => $label) {
        $function = 'get'. $field;
        $value = java_values($object->$function());
        $null = null;
        if (is_array($label)) {
          if (isset($label['field'])) $field = $label['field'];
          if (isset($label['null'])) $null = $label['null'];
          if (isset($label['map'])) $value = $label['map'][$value];
          $label = $label['label'];
        }
        if ($value != $null) $properties[$field] = array('label' => $label, 'value' => $value);
      }
    }
    foreach ($dsi->getCustomProperties() as $field => $value) {
      if (!java_is_null($value)) $properties[$field] = array('label' => t(java_values($value->getName())), 'value' => java_values($value->getValue()));
    }
    return $properties;
  }
  public function getFont($cellstyle) {
    return $cellstyle->getFont($this->_workbook);
  }
  public function getFontOfFormattingRun($rtf, $run) {
    return $this->_workbook->getFontAt(java_values($rtf->getFontOfFormattingRun($run)));
  }
  public function getColor($font) {
    return $font->getColor();
  }
  public function getFillForegroundColor($cellstyle) {
    return $cellstyle->getFillForegroundColor();
  }
  public function getBorderColor($cellstyle, $pos) {
    $getcolor = "get{$pos}BorderColor";
    return $cellstyle->$getcolor();
  }
  public function importFormula($cell, &$c) {
    $c['formula'] = java_values($cell->getCellFormula());
    $c['datatype'] = 'f';
    switch (java_values($cell->getCachedFormulaResultType())) {
    case java_values(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_BOOLEAN):
      $c['datavalue'] = java_values($cell->getBooleanCellValue());
      $c['valuetype'] = 'n';
      break;
    case java_values(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_NUMERIC):
      $c['datavalue'] = java_values($cell->getNumericCellValue());
      $c['valuetype'] = 'n';
      break;
    case java_values(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_STRING):
      $c['datavalue'] = java_values($cell->getStringCellValue());
      $c['valuetype'] = 't';
      break;
    default:
      $c['datavalue'] = 0;
      $c['valuetype'] = 'n';
      break;
    }
  }
  public function importColor($color) {
    $index = java_values($color);
    if ($index == java_values(org_apache_poi_hssf_usermodel_HSSFFont::type()->COLOR_NORMAL)) return FALSE;
    $automatic = new java_class('org.apache.poi.hssf.util.HSSFColor$AUTOMATIC');
    if ($index == java_values($automatic->index)) return FALSE;
    $palette = $this->_workbook->getCustomPalette();
    $triplet = java_values($palette->getColor($index)->getTriplet());
    return $triplet;
  }
}

class SheetnodeXSSF extends SheetnodeXLS {
  public function SheetnodeXSSF($workbook) {
    $this->_workbook = $workbook;
  }
  protected function importProperties() {
    $properties = array();
    $fields_objects = array(
      'CoreProperties' => array(
        'Category' => t('Category'),
        'ContentStatus' => t('Content status'),
        'ContentType' => t('Content type'),
        'Created' => array('label' => t('Creation date'), 'null' => 0),
        'Creator' => t('Author'),
        'Description' => array('label' => t('Body'), 'field' => 'Body'),
        'Identifier' => t('Identifier'),
        'Keywords' => t('Keywords'),
        'Language' => t('Language'),
        'LastModifiedBy' => t('Last modified by'),
        'LastPrinted' => t('Last printing date'),
        'Modified' => t('Last modification date'),
        'Revision' => t('Revision'),
        'Subject' => t('Subject'),
        'Title' => t('Title'),
        'Version' => t('Version'),
      ),
      'ExtendedProperties' => array(
        'Application' => t('Application'),
        'AppVersion' => t('Application version'),
        'Characters' => t('Character count'),
        'CharactersWithSpaces' => t('Character count (with spaces)'),
        'Company' => t('Company'),
        'DigSig' => t('Digital signature'),
        'DocSecurity' => array('label' => t('Security'), 'field' => 'Security', 'map' => array(
          0 => null,  
          1 => t('Password protected'),
          2 => t('Read-only recommended'),
          4 => t('Read-only enforced'),
          8 => t('Locked for annotations'),
        )),
        'Lines' => array('label' => t('Line count'), 'null' => 0),
        'Manager' => t('Manager'),
        'Pages' => array('label' => t('Page count'), 'null' => 0),
        'Paragraphs' => array('label' => t('Paragraph count'), 'null' => 0),
        'Template' => t('Template'),
        'TotalTime' => array('label' => t('Total time spent on editing'), 'null' => 0),
      ),
    );
    foreach ($fields_objects as $propertyset => $fields) {
      $function = 'get'. $propertyset;
      $object = $this->_workbook->getProperties()->$function()->getUnderlyingProperties();
      foreach ($fields as $field => $label) {
        if ($propertyset == 'CoreProperties') {
          $function = 'get'. $field .'Property';
          $value = java_values($object->$function()->getValue());
        }
        else {
          $function = 'get'. $field;
          $value = java_values($object->$function());
        }
        $null = null;
        if (is_array($label)) {
          if (isset($label['field'])) $field = $label['field'];
          if (isset($label['null'])) $null = $label['null'];
          if (isset($label['map'])) $value = $label['map'][$value];
          $label = $label['label'];
        }
        if ($value != $null) $properties[$field] = array('label' => $label, 'value' => $value);
      }
    }
    return $properties;
  }
  public function getFont($cellstyle) {
    return $cellstyle->getFont();
  }
  public function getFontOfFormattingRun($rtf, $run) {
    return $this->_workbook->getFontAt(0); // TODO
    return $rtf->getFontOfFormattingRun($run);
  }
  public function getColor($font) {
    return $font->getXSSFColor();
  }
  public function getFillForegroundColor($cellstyle) {
    return $cellstyle->getFillForegroundXSSFColor();
  }
  public function getBorderColor($cellstyle, $pos) {
    $getcolor = "get{$pos}BorderXSSFColor";
    return $cellstyle->$getcolor();
  }
  public function importFormula($cell, &$c) {
    $c['formula'] = java_values($cell->getCellFormula());
    $c['datatype'] = 'f';
    $v = java_values($cell->getStringCellValue());
    if (empty($v)) {
      $c['datavalue'] = java_values($cell->getNumericCellValue());
      $c['valuetype'] = 'n';
    }
    else {
      $c['datavalue'] = $v;
      $c['valuetype'] = 't';
    }
  }
  public function importColor($color) {
    if (java_is_null($color)) return FALSE;
    $argb = java_values($color->getRgb());
    if (!$argb) return FALSE;
    $triplet = array(ord($argb[1]),ord($argb[2]),ord($argb[3]));
    return $triplet;
  }
}

// @see http://drupal.org/node/293663
function _sheetnode_xls_import_form_submit($form, &$form_state) {
  set_time_limit(0);
  
  global $user;
  module_load_include('inc', 'node', 'node.pages');
  
  // Open sheet file
  $inp = new java_io_FileInputStream($_FILES['files']['tmp_name']['import']);
  $workbook = org_apache_poi_ss_usermodel_WorkbookFactory::type()->create($inp);
  $sheetnode = SheetnodeXLS::create($workbook);

  // Create a book if there are more than one sheet
  $numsheets = java_values($workbook->getNumberOfSheets());
  if ($numsheets > 1 && module_exists('book')) {
    $node = array('type' => 'book');
    $form_state_book = array();
    $form_state_book['values']['title'] = $sheetnode->getProperties('Title') ? $sheetnode->getProperties('Title') : pathinfo($_FILES['files']['name']['import'], PATHINFO_FILENAME);
    $body = '';
    foreach ($sheetnode->getProperties() as $field => $property) {
      if ($field != 'Title' && $field != 'Body') {
        $body .= $property['label'] .': <strong>'. $property['value'] ."</strong>\n";
      }
    }
    $body .= "\n". $sheetnode->getProperties('Body');
    $form_state_book['values']['body'] = $body;
    $form_state_book['values']['name'] = $user->name;
    $form_state_book['values']['op'] = t('Save');
    $form_state_book['values']['book']['bid'] = 'new';
    $form_state_book['values']['book']['plid'] = -1;
    $form_state_book['values']['book']['weight'] = 0;
    drupal_execute('book_node_form', $form_state_book, (object)$node);
    $book = node_load($form_state_book['nid']);
  }

  // Import all sheets
  for ($s = 0; $s < $numsheets; $s++) {
    $sheet = $workbook->getSheetAt($s);
    $sheetname = java_values($workbook->getSheetName($s));

    $node = array('type' => 'sheetnode');
    $form_state_node = array();
    $form_state_node['values']['title'] = $sheetname;
    $body = '';
    if (!isset($book)) {
      foreach ($sheetnode->getProperties() as $field => $property) {
        $body .= $property['label'] .': <strong>'. $property['value'] ."</strong>\n";
      }
    }
    $form_state_node['values']['body'] = $body;
    $form_state_node['values']['name'] = $user->name;
    $form_state_node['values']['sheetsave'] = _sheetnode_xls_import_do($workbook, $sheet, $sheetname, $sheetnode);
    $form_state_node['values']['op'] = t('Save');
    if (isset($book)) { // insert inside the book
      $form_state_node['values']['book']['bid'] = $book->nid;
      $form_state_node['values']['book']['plid'] = $book->book['mlid'];
      $form_state_node['values']['book']['weight'] = $s;
    }
    drupal_execute('sheetnode_node_form', $form_state_node, (object)$node);
  }

  // If more than one sheet, create a book to hold them
  if (isset($book)) {
    $form_state['redirect'] = 'node/'. $book->nid;
  }
  else {
    $form_state['redirect'] = 'node/'. $form_state_node['nid'];
  }
}

function _sheetnode_xls_import_do($workbook, $sheet, $sheetname, $sheetnode) {
  $sc = array(); // SocialCalc array structure

  // Default values.
  $sc['attribs']['defaultcolwidth'] = round(java_values($sheet->getDefaultColumnWidth()) * WIDTH_CHAR_TO_PX, 2); // width in chars * factor
  $sc['attribs']['defaultrowheight'] = round(java_values($sheet->getDefaultRowHeightInPoints()), 2);
  $sc['attribs']['lastrow'] = java_values($sheet->getLastRowNum())+1;
  $sc['attribs']['lastcol'] = 1; // POI doesn't tell us last col, so we'll compute it
  
  // Rows
  for ($rit = $sheet->rowIterator(); java_is_true($rit->hasNext()); ) {
    $row = $rit->next();
    $rowNum = intval(java_values($row->getRowNum()));
    $height = round(java_values($row->getHeightInPoints()), 2);
    if ($height != $sc['attribs']['defaultrowheight']) {
      $sc['rowattribs']['height'][$rowNum+1] = $height;
    }
    if (java_is_true($row->getZeroHeight())) {
      $sc['rowattribs']['hide'][$rowNum+1] = true;
    }

    // Columns and cells
    for ($cit = $row->cellIterator(); java_is_true($cit->hasNext()); ) {
      $cell = $cit->next();

      $colNum = intval(java_values($cell->getColumnIndex()));
      $width = round(java_values($sheet->getColumnWidth($colNum)) * WIDTH_CHAR_TO_PX / 256, 2); // width in 1/256 of a char!! 
      if ($width > 0 && $width != $sc['attribs']['defaultcolwidth']) {
        $sc['colattribs']['width'][$colNum+1] = $width;
      }
//      if (java_is_true($sheet->IsColumnHidden($colNum))) {
//        $sc['colattribs']['hide'][$colNum+1] = true;
//      }
      
      // Last column
      if ($colNum >= $sc['attribs']['lastcol']) $sc['attribs']['lastcol'] = $colNum+1;

      _sheetnode_xls_import_cell($workbook, $cell, $sc, $sheetnode);
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
  for ($i=0; $i < java_values($workbook->getNumberOfNames()); $i++) {
    $name = $workbook->getNameAt($i);
    $definition = java_values($name->getRefersToFormula());
    $definition = str_replace("'$sheetname'!", '', $definition);
    $definition = str_replace("$sheetname!", '', $definition);
    $sc['names'][java_values($name->getNameName())] = array(
      'desc' => '',
      'definition' => (java_is_true($name->isFunctionName()) ? '=' : ''). $definition,
    );
  }

  $socialcalc = array(
    'sheet' => $sc,
    'edit' => socialcalc_default_edit($sc),
    'audit' => socialcalc_default_audit($sc),
  );
  return socialcalc_save($socialcalc);
}

function _sheetnode_xls_import_cell($workbook, $cell, &$sc, $sheetnode) {
  $type = java_values($cell->getCellType());
  $c = array(); // output cell
  
  // Position
  $pos = array(java_values($cell->getColumnIndex())+1, java_values($cell->getRowIndex())+1);
  $c['pos'] = $pos;

  // Value and type
  switch ($type) {
  case java_values(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_BOOLEAN):
    $c['datavalue'] = java_values($cell->getBooleanCellValue());
    $c['datatype'] = 'v';
    $c['valuetype'] = 'n';
    break;
  case java_values(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_NUMERIC):
    $c['datavalue'] = java_values($cell->getNumericCellValue());
    $c['datatype'] = 'v';
    $c['valuetype'] = 'n';
    break;
  case java_values(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_STRING):
    $rtf = $cell->getRichStringCellValue();
    $runs = java_values($rtf->numFormattingRuns());
    if ($runs > 1) { // it is rich text
      $c['datavalue'] = '';
      $text = java_values($rtf->getString());
      $i = 0;
      for ($r=0; $r <= $runs; $r++) {
        if ($r < $runs) { 
          $j = java_values($rtf->getIndexOfFormattingRun($r));
        }
        else {
          $j = strlen($text);
        }
        $font = $r == 0 ? $sheetnode->getFont($cell->getCellStyle()) : $sheetnode->getFontOfFormattingRun($rtf, $r-1);
        $fontstyle = 'font:'. _sheetnode_xls_get_fontstyle($font) .';';
        $color = _sheetnode_xls_get_color($sheetnode, $sheetnode->getColor($font), true);
        $colorstyle = $color ? 'color:'. $color .';': '';
        $c['datavalue'] .= "<span style=\"$fontstyle$colorstyle\">". substr($text, $i, $j-$i) .'</span>';
        $i = $j;
      }
      $c['valuetype'] = 'th';

      // Add text-html format
      $valueformat = 'text-html';
      if (!($index = $sc['valueformathash'][$valueformat])) {
        $index = count($sc['valueformats'])+1;
        $sc['valueformats'][$index] = $valueformat;
        $sc['valueformathash'][$valueformat] = $index;
      }
      $c['textvalueformat'] = $index;
    }
    else { // it is plain text
      $c['datavalue'] = java_values($cell->getStringCellValue());
      $c['valuetype'] = 't';
    }
    $c['datatype'] = 't';
    break;
  case java_values(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_FORMULA):
    $sheetnode->importFormula($cell, $c);
    break;
  case java_values(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_ERROR):
    $c['errors'] = java_values(org_apache_poi_ss_usermodel_FormulaError::type()->forInt($cell->getErrorCellValue())->getString());
    break;
  case java_values(org_apache_poi_ss_usermodel_Cell::type()->CELL_TYPE_BLANK):
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
    // Import this style
    _sheetnode_xls_import_cellstyle($workbook, $cellstyle, $sc, $sheetnode);

    // Font
    $font = $sheetnode->getFont($cellstyle);
    if (!java_is_null($font)) {
      // Import this font
      _sheetnode_xls_import_font($workbook, $font, $sc, $sheetnode);
      
      $fonthash = $sc['fonthash'][_sheetnode_xls_get_fontstyle($font)];
      if (!isset($sc['attribs']['defaultfont']) || $fonthash != $sc['attribs']['defaultfont']) $c['font'] = $fonthash;
    }
    
    // Borders
    $bt = _sheetnode_xls_get_borderstyle($workbook, $cellstyle, 'Top', $sheetnode);
    if ($bt) $c['bt'] = $sc['borderstylehash'][$bt];
    $bl = _sheetnode_xls_get_borderstyle($workbook, $cellstyle, 'Left', $sheetnode);
    if ($bl) $c['bl'] = $sc['borderstylehash'][$bl];
    $bb = _sheetnode_xls_get_borderstyle($workbook, $cellstyle, 'Bottom', $sheetnode);
    if ($bb) $c['bb'] = $sc['borderstylehash'][$bb];
    $br = _sheetnode_xls_get_borderstyle($workbook, $cellstyle, 'Right', $sheetnode);
    if ($br) $c['br'] = $sc['borderstylehash'][$br];

    // Cell format (horizontal alignment)
    $cellformat = _sheetnode_xls_get_cellformat($cellstyle);
    if ($cellformat) {
      $c['cellformat'] = $sc['cellformathash'][$cellformat];
    }

    // Layout (padding + vertical alignment)
    $layouthash = $sc['layouthash'][_sheetnode_xls_get_layout($cellstyle)];
    if (!isset($sc['attribs']['defaultlayout']) || $layouthash != $sc['attribs']['defaultlayout']) $c['layout'] = $layouthash;

    // Color
    $color = _sheetnode_xls_get_color($sheetnode, $sheetnode->getColor($font));
    if ($color) {
      $colorhash = $sc['colorhash'][$color];
      if ($colorhash != $sc['attribs']['defaultcolor']) $c['color'] = $colorhash;
    }

    // Background color
    $bgcolor = _sheetnode_xls_get_color($sheetnode, $sheetnode->getFillForegroundColor($cellstyle));
    if ($bgcolor) {
      $bgcolorhash = $sc['colorhash'][$bgcolor];
      if ($bgcolorhash != $sc['attribs']['defaultbgcolor']) $c['bgcolor'] = $bgcolorhash;
    }
    
    // Non-text value format
    $ntvf = _sheetnode_xls_get_valueformat(java_values($cellstyle->getDataFormatString()));
    if ($ntvf) {
      $ntvfhash = $sc['valueformathash'][$ntvf];
      if (!isset($sc['attribs']['defaultnontextvalueformat']) || $ntvfhash != $sc['attribs']['defaultnontextvalueformat']) $c['nontextvalueformat'] = $ntvfhash;
    }
  }

  $sc['cells'][socialcalc_cr_to_coord($pos[0],$pos[1])] = $c;
}

function _sheetnode_xls_import_font($workbook, $font, &$sc, $sheetnode) {
  if ($sheetnode->existsFont($font)) return;

  $fontstyle = _sheetnode_xls_get_fontstyle($font);
  if (!isset($sc['fonthash'][$fontstyle])) {
    $index = count($sc['fonts'])+1;
    $sc['fonts'][$index] = $fontstyle;
    $sc['fonthash'][$fontstyle] = $index;
  }
  
  $color = _sheetnode_xls_get_color($sheetnode, $sheetnode->getColor($font));
  if ($color && !isset($sc['colorhash'][$color])) {
    $index = count($sc['colors'])+1;
    $sc['colors'][$index] = $color;
    $sc['colorhash'][$color] = $index;
  }
}

function _sheetnode_xls_import_cellstyle($workbook, $cellstyle, &$sc, $sheetnode) {
  if ($sheetnode->existsCellstyle($cellstyle)) return;

  // Border styles
  _sheetnode_xls_import_borderstyle($workbook, $cellstyle, 'Top', $sc, $sheetnode);
  _sheetnode_xls_import_borderstyle($workbook, $cellstyle, 'Right', $sc, $sheetnode);
  _sheetnode_xls_import_borderstyle($workbook, $cellstyle, 'Left', $sc, $sheetnode);
  _sheetnode_xls_import_borderstyle($workbook, $cellstyle, 'Bottom', $sc, $sheetnode);

  // Cell format (horizontal alignment)
  $cellformat = _sheetnode_xls_get_cellformat($cellstyle);
  if ($cellformat && !isset($sc['cellformathash'][$cellformat])) {
    $index = count($sc['cellformats'])+1;
    $sc['cellformats'][$index] = $cellformat;
    $sc['cellformathash'][$cellformat] = $index;
  }

  // Layout (padding + vertical alignment)
  $layout = _sheetnode_xls_get_layout($cellstyle);
  if (!isset($sc['layouthash'][$layout])) {
    $index = count($sc['layouts'])+1;
    $sc['layouts'][$index] = $layout;
    $sc['layouthash'][$layout] = $index;
  }

  // Background color
  $bgcolor = _sheetnode_xls_get_color($sheetnode, $sheetnode->getFillForegroundColor($cellstyle));
  if ($bgcolor && !isset($sc['colorhash'][$bgcolor])) {
    $index = count($sc['colorhash'])+1;
    $sc['colors'][$index] = $bgcolor;
    $sc['colorhash'][$bgcolor] = $index;
  }

  // Value format
  $valueformat = _sheetnode_xls_get_valueformat(java_values($cellstyle->getDataFormatString()));
  if ($valueformat && !isset($sc['valueformathash'][$valueformat])) {
    $index = count($sc['valueformats'])+1;
    $sc['valueformats'][$index] = $valueformat;
    $sc['valueformathash'][$valueformat] = $index;
  }
}

function _sheetnode_xls_import_borderstyle($workbook, $cellstyle, $pos, &$sc, $sheetnode) {
  $borderstyle = _sheetnode_xls_get_borderstyle($workbook, $cellstyle, $pos, $sheetnode);
  if ($borderstyle && !isset($sc['borderstylehash'][$borderstyle])) {
    $index = count($sc['borderstyles'])+1;
    $sc['borderstyles'][$index] = $borderstyle;
    $sc['borderstylehash'][$borderstyle] = $index;
  }
}

function _sheetnode_xls_get_fontstyle($font) {
  return (java_is_true($font->getItalic()) ? 'italic' : 'normal') .' '.
         (java_values($font->getBoldWeight()) == java_values(org_apache_poi_ss_usermodel_Font::type()->BOLDWEIGHT_BOLD) ? 'bold' : 'normal') .' '.
          java_values($font->getFontHeightInPoints()) .'pt '.
          java_values($font->getFontName());
}

function _sheetnode_xls_get_borderstyle($workbook, $cellstyle, $pos, $sheetnode) {
  $borderstyles = array(
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_NONE) => FALSE,
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_THIN) => 'thin solid', 
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_MEDIUM) => 'medium solid', 
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_DASHED) => 'thin dashed', 
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_DOTTED) => 'thin dotted', 
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_THICK) => 'thick solid', 
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_DOUBLE) => 'thin double', 
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_HAIR) => 'thin dotted',  // TODO: is it?
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_MEDIUM_DASHED) => 'medium dashed', 
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_DASH_DOT) => 'thin dashed', // TODO: is it?
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_MEDIUM_DASH_DOT) => 'medium dashed', //TODO: is it?
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_DASH_DOT_DOT) => 'thin dashed', // TODO: is it?
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_MEDIUM_DASH_DOT_DOT) => 'medium dashed', // TODO: is it?
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->BORDER_SLANTED_DASH_DOT) => 'thin dashed', // TODO: is it?
  );
  $getstyle = "getBorder{$pos}";
  $borderstyle = $borderstyles[java_values($cellstyle->$getstyle())];
  if (!$borderstyle) return FALSE;
  $bordercolor = _sheetnode_xls_get_color($sheetnode, $sheetnode->getBorderColor($cellstyle, $pos));
  return $borderstyle .' '. $bordercolor;
}

function _sheetnode_xls_get_cellformat($cellstyle) {
  $cellformats = array(
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_GENERAL) => FALSE,
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_LEFT) => 'left',
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_CENTER) => 'center',
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_RIGHT) => 'right',
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_FILL) => 'justify', // TODO: is it?
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_JUSTIFY) => 'justify',
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->ALIGN_CENTER_SELECTION) => 'center', // TODO: is it?
  );
  return $cellformats[java_values($cellstyle->getAlignment())];
}

function _sheetnode_xls_get_layout($cellstyle) {
  $valignments = array(
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->VERTICAL_TOP) => 'top',
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->VERTICAL_CENTER) => 'middle',
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->VERTICAL_BOTTOM) => 'bottom', 
    java_values(org_apache_poi_ss_usermodel_CellStyle::type()->VERTICAL_JUSTIFY) => 'middle', // TODO: is it?
  );
  $valignment = $valignments[java_values($cellstyle->getVerticalAlignment())];
  $padding_left = java_values($cellstyle->getIndention()) * WIDTH_CHAR_TO_PX; // simulate indentation with padding
  return 'padding:* * * '. $padding_left .'px;vertical-align:'. $valignment .';';
}

function _sheetnode_xls_get_valueformat($dataformat) {
  if (strtolower($dataformat) == 'general') return FALSE;
  return $dataformat;
}

function _sheetnode_xls_get_color($sheetnode, $color, $hex = FALSE) {
  $triplet = $sheetnode->importColor($color);
  if (!$triplet) return FALSE;
  return $hex ? 
    sprintf('#%02X%02X%02X', $triplet[0], $triplet[1], $triplet[2]) : 
    'rgb('. $triplet[0] .','. $triplet[1] .','. $triplet[2] .')';
}

