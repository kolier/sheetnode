<?php
// $Id$

/**
 * @file SocialCalc manipulation functions
 * Translated from socialcalc-3.js
 */

function socialcalc_parse($data) {
  $line = strtok($data, "\n");
  $sheet = array();
  while ($line !== FALSE) {
    $parts = explode(':', $line);
    switch ($parts[0]) {
    case 'cell':
      $coord = $parts[1];
      $cell = isset($sheet['cells'][$coord]) ? $sheet['cells'][$coord] : array('pos' => socialcalc_coord_to_cr($coord));
      $cell += socialcalc_parse_cell($parts, 2);
      $sheet['cells'][$coord] = $cell;
      break;

    case "col":
      $coord = $parts[1];
      $pos = socialcalc_coord_to_cr($coord .'1'); // convert to col number
      $j = 2;
      while ($t = $parts[$j++]) {
        switch ($t) {
        case "w":
          $sheet['colattribs']['width'][$pos[0]] = strval($parts[$j++]); // must be text - could be auto or %, etc.
          break;
        case "hide":
          $sheet['colattribs']['hide'][$pos[0]] = $parts[$j++];
          break;
        }
      }
      break;

    case "row":
      $coord = intval($parts[1]);
      $j = 2;
      while ($t = $parts[$j++]) {
        switch ($t) {
        case "h":
          $sheet['rowattribs']['height'][$coord] = intval($parts[$j++]);
          break;
        case "hide":
          $sheet['rowlattribs']['hide'][$coord] = $parts[$j++];
          break;
        }
      }
      break;

    case "sheet":
      $j = 1;
      while ($t = $parts[$j++]) {
        switch ($t) {
        case "c":
          $sheet['attribs']['lastcol'] = intval($parts[$j++]);
          break;
        case "r":
          $sheet['attribs']['lastrow'] = intval($parts[$j++]);
          break;
        case "w":
          $sheet['attribs']['defaultcolwidth'] = intval($parts[$j++]);
          break;
        case "h":
          $sheet['attribs']['defaultrowheight'] = intval($parts[$j++]);
          break;
        case "tf":
          $sheet['attribs']['defaulttextformat'] = intval($parts[$j++]);
          break;
        case "ntf":
          $sheet['attribs']['defaultnontextformat'] = intval($parts[$j++]);
          break;
        case "layout":
          $sheet['attribs']['defaultlayout'] = intval($parts[$j++]);
          break;
        case "font":
          $sheet['attribs']['defaultfont'] = intval($parts[$j++]);
          break;
        case "tvf":
          $sheet['attribs']['defaulttextvalueformat'] = intval($parts[$j++]);
          break;
        case "ntvf":
          $sheet['attribs']['defaultnontextvalueformat'] = intval($parts[$j++]);
          break;
        case "color":
          $sheet['attribs']['defaultcolor'] = intval($parts[$j++]);
          break;
        case "bgcolor":
          $sheet['attribs']['defaultbgcolor'] = intval($parts[$j++]);
          break;
        case "circularreferencecell":
          $sheet['attribs']['circularreferencecell'] = $parts[$j++];
          break;
        case "recalc":
          $sheet['attribs']['recalc'] = $parts[$j++];
          break;
        case "needsrecalc":
          $sheet['attribs']['needsrecalc'] = $parts[$j++];
          break;
        default:
          $j += 1;
          break;
        }
      }
    break;

    case "name":
      $name = strtoupper(socialcalc_decode_value($parts[1]));
      $sheet['names'][$name] = array('desc' => socialcalc_decode_value($parts[2]),
                                     'definition' => socialcalc_decode_value($parts[3]));
      break;

    case "layout":
      $parts = array();
      preg_match('/^layout\:(\d+)\:(.+)$/', $line, $parts); // layouts can have ":" in them
      $sheet['layouts'][intval($parts[1])] = $parts[2];
      $sheet['layouthash'][$parts[2]] = intval($parts[1]);
      break;

    case "font":
      $sheet['fonts'][intval($parts[1])] = $parts[2];
      $sheet['fonthash'][$parts[2]] = intval($parts[1]);
      break;

    case "color":
      $sheet['colors'][intval($parts[1])] = $parts[2];
      $sheet['colorhash'][$parts[2]] = intval($parts[1]);
      break;

    case "border":
      $sheet['borderstyles'][intval($parts[1])] = $parts[2];
      $sheet['borderstylehash'][$parts[2]] = intval($parts[1]);
      break;

    case "cellformat":
      $sheet['cellformats'][intval($parts[1])] = $parts[2];
      $sheet['cellformathash'][$parts[2]] = intval($parts[1]);
      break;

    case "valueformat":
      $v = socialcalc_decode_value($parts[2]);
      $sheet['valueformats'][intval($parts[1])] = $v;
      $sheet['valueformathash'][$v] = intval($parts[1]);
      break;

    case "version":
      break;

    case "copiedfrom":
      $sheetobj['copiedfrom'] = $parts[1] .':'. $parts[2];
      break;

    case "clipboardrange": // in save versions up to 1.3. Ignored.
    case "clipboard":
      break;

    case "":
      break;

    default:
      break;
    }
    $line = strtok("\n");
  }
  return $sheet;
}

function socialcalc_save($sc) {
  $sheet = array();
  $sheet[] = 'version:1.5';
  if ($sc['cells']) foreach ($sc['cells'] as $cell) {
    $sheet[] = socialcalc_save_cell($cell);
  }
  if ($sc['colattribs']) {
    if ($sc['colattribs']['width']) foreach ($sc['colattribs']['width'] as $col => $width) {
      $coord = socialcalc_cr_to_coord($col, 1);
      $sheet[] = 'col:'. substr($coord, 0, -1) . ':w:'. $width;
    }
  }
  if ($sc['rowattribs']) {
    if ($sc['rowattribs']['height']) foreach ($sc['rowattribs']['height'] as $row => $height) {
      $sheet[] = 'row:'. $row .':h:'. $height;
    }
  }
  if ($sc['fonts']) foreach ($sc['fonts'] as $fid => $font) {
    $sheet[] = 'font:'. $fid .':'. $font;
  }
  if ($sc['borderstyles']) foreach ($sc['borderstyles'] as $bsid => $borderstyle) {
    $sheet[] = 'border:'. $bsid .':'. $borderstyle;
  }
  if ($sc['cellformats']) foreach ($sc['cellformats'] as $cfid => $cellformat) {
    $sheet[] = 'cellformat:'. $cfid .':'. $cellformat;
  }

  $sheetfields = array(
    'lastcol' => 'c', 
    'lastrow' => 'r', 
    'defaultcolwidth' => 'w', 
    'defaultrowheight' => 'h',
    'defaulttextformat' => 'tf', 
    'defaultnontextformat' => 'ntf', 
    'defaulttextvalueformat' => 'tvf', 
    'defaultnontextvalueformat' => 'ntvf',
    'defaultlayout' => 'layout', 
    'defaultfont' => 'font', 
    'defaultcolor' => 'color', 
    'defaultbgcolor' => 'bgcolor',
    'circularreferencecell' => 'circularreferencecell', 
    'recalc' => 'recalc', 
    'needsrecalc' => 'needsrecalc'
  );
  $sheetattribs = array();
  foreach ($sheetfields as $key => $attrib) {
    if (isset($sc['attribs'][$key])) {
      $sheetattribs[] = $attrib .':'. $sc['attribs'][$key];
    }
  }
  if ($sheetattribs) {
    $sheet[] = 'sheet:'. implode(':', $sheetattribs);
  }

  $sheetsave = implode($sheet, "\n") ."\n";
//  var_dump($sheetsave);
  return $sheetsave;
}

function socialcalc_save_cell($cell) {
  $line = 'cell:'. socialcalc_cr_to_coord($cell['pos'][0], $cell['pos'][1]);
  $value = socialcalc_encode_value($cell['datavalue']);
  if ($cell['datatype']=='v') {
    if ($cell['valuetype']=='n') $line .= ':v:'. $value;
    else $line .= ':vt:'. $cell['valuetype'] .':' .$value;
  }
  else if ($cell['datatype']=='t') {
    if ($cell['valuetype']=='t') $line .= ':t:'. $value;
    else $line .= ':vt:'. $cell['valuetype'] .':' .$value;
  }
  else {
    $formula = socialcalc_encode_value($cell['formula']);
    if ($cell['datatype']=='f')
      $line .= ':vtf:'. $cell['valuetype'] .':'. $value .':'. $formula;
    else if ($cell['datatype']=='c')
      $line .= ':vtc:'. $cell['valuetype'] .':'. $value .':'. $formula;
  }
  if ($cell['errors']) {
    $line .= ':e:'. socialcalc_encode_value($cell['errors']);
  }
  $t = $cell['bt'] || '';
  $r = $cell['br'] || '';
  $b = $cell['bb'] || '';
  $l = $cell['bl'] || '';
  if ($t || $r || $b || $l)
    $line .= ':b:'. $t .':'. $r .':'. $b .':'. $l;
  if ($cell['layout']) $line .= ':l:'. $cell['layout'];
  if ($cell['font']) $line .= ':f:'. $cell['font'];
  if ($cell['color']) $line .= ':c:'. $cell['color'];
  if ($cell['bgcolor']) $line .= ':bg:'. $cell['bgcolor'];
  if ($cell['cellformat']) $line .= ':cf:'. $cell['cellformat'];
  if ($cell['textvalueformat']) $line .= ':tvf:'. $cell['textvalueformat'];
  if ($cell['nontextvalueformat']) $line .= ':ntvf:'. $cell['nontextvalueformat'];
  if ($cell['colspan']) $line .= ':colspan:'. $cell['colspan'];
  if ($cell['rowspan']) $line .= ':rowspan:'. $cell['rowspan'];
  if ($cell['cssc']) $line .= ':cssc:' .$cell['cssc'];
  if ($cell['csss']) $line .= ':csss:' .socialcalc_encode_value($cell['csss']);
  if ($cell['mod']) $line .= ':mod:'. $cell['mod'];
  if ($cell['comment']) $line .= ':comment:'. socialcalc_encode_value($cell['comment']);

  return $line;
}

function socialcalc_encode_value($s) {
  if (!is_string($s)) return $s;
  return str_replace(
    array('\\', ':', "\n"),
    array('\\b', '\\c', '\\n'),
    $s
  );
}

function socialcalc_parse_cell($parts, $j) {
  $cell = array();
  while ($t = $parts[$j++]) {
    switch ($t) {
    case "v":
      $cell['datavalue'] = doubleval(socialcalc_decode_value($parts[$j++]));
      $cell['datatype'] = "v";
      $cell['valuetype'] = "n";
      break;
    case "t":
      $cell['datavalue'] = strval(socialcalc_decode_value($parts[$j++]));
      $cell['datatype'] = "t";
      $cell['valuetype'] = "t"; 
      break;
    case "vt":
      $v = $parts[$j++];
      $cell['valuetype'] = $v;
      $cell['datatype'] = $v[0]=="n" ? "v" : "t";
      $cell['datavalue'] = socialcalc_decode_value($parts[$j++]);
      break;
    case "vtf":
      $cell['valuetype'] = strval($parts[$j++]);
      $cell['datavalue'] = socialcalc_decode_value($parts[$j++]);
      $cell['formula'] = strval(socialcalc_decode_value($parts[$j++]));
      $cell['datatype'] = "f";
      break;
    case "vtc":
      $cell['valuetype'] = strval($parts[$j++]);
      $cell['datavalue'] = socialcalc_decode_value($parts[$j++]);
      $cell['formula'] = strval(socialcalc_decode_value($parts[$j++]));
      $cell['datatype'] = "c";
      break;
    case "e":
      $cell['errors'] = strval(socialcalc_decode_value($parts[$j++]));
      break;
    case "b":
      $cell['bt'] = intval($parts[$j++]);
      $cell['br'] = intval($parts[$j++]);
      $cell['bb'] = intval($parts[$j++]);
      $cell['bl'] = intval($parts[$j++]);
      break;
    case "l":
      $cell['layout'] = intval($parts[$j++]);
      break;
    case "f":
      $cell['font'] = intval($parts[$j++]);
      break;
    case "c":
      $cell['color'] = intval($parts[$j++]);
      break;
    case "bg":
      $cell['bgcolor'] = intval($parts[$j++]);
      break;
    case "cf":
      $cell['cellformat'] = intval($parts[$j++]);
      break;
    case "ntvf":
      $cell['nontextvalueformat'] = intval($parts[$j++]);
      break;
    case "tvf":
      $cell['textvalueformat'] = intval($parts[$j++]);
      break;
    case "colspan":
      $cell['colspan'] = intval($parts[$j++]);
      break;
    case "rowspan":
      $cell['rowspan'] = intval($parts[$j++]);
      break;
    case "cssc":
      $cell['cssc'] = strval($parts[$j++]);
      break;
    case "csss":
      $cell['csss'] = strval(socialcalc_decode_value($parts[$j++]));
      break;
    case "mod":
      $j+=1;
      break;
    case "comment":
      $cell['comment'] = strval(socialcalc_decode_value($parts[$j++]));
      break;
    }
  }
  return $cell;
}

function socialcalc_coord_to_cr($coord) {
  static $coord_to_cr = array();
  if (isset($coord_to_cr[$coord])) return $coord_to_cr[$coord];
  $c=0;$r=0;
  for ($i=0; $i<strlen($coord); $i++) { // this was faster than using regexes; assumes well-formed
    $ch = ord(substr($coord, $i, 1));
    if ($ch==36) ; // skip $'s
    else if ($ch<=57) $r = 10*$r + $ch-48;
    else if ($ch>=97) $c = 26*$c + $ch-96;
    else if ($ch>=65) $c = 26*$c + $ch-64;
  }
  $coord_to_cr[$coord] = array($c, $r);
  return $coord_to_cr[$coord];
}


function socialcalc_cr_to_coord($c, $r) {
  $letters = array( "A","B","C","D","E","F","G","H","I","J","K","L","M",
                    "N","O","P","Q","R","S","T","U","V","W","X","Y","Z" );
  if ($c < 1) $c = 1;
  if ($c > 702) $c = 702; // maximum number of columns - ZZ
  if ($r < 1) $r = 1;
  $collow = ($c - 1) % 26;
  $colhigh = floor(($c - 1) / 26);
  if ($colhigh) {
    $result = $letters[$colhigh-1] . $letters[$collow] . $r;
  }
  else {
    $result = $letters[$collow] . $r;
  }
  return $result;
}

function socialcalc_decode_value($s) {
  if (!is_string($s)) return $s;
  if (strstr($s, '\\') === FALSE) return $s; // for performace reasons: replace nothing takes up time
  $r = str_replace('\\c', ':', $s);
  $r = str_replace('\\n', "\n", $r);
  return str_replace('\\b', '\\', $r);
}

function socialcalc_cellformat_parsefont($cell, $sheet) {
  if (!isset($cell['font'])) return FALSE;
  $parts = array();
  preg_match('/^(\*|\S+? \S+?) (\S+?) (\S.*)$/', $sheet['fonts'][$cell['font']], $parts);
  $font = array();
  if ($parts[3] != '*') $font['family'] = $parts[3];
  if ($parts[2] != '*') $font['size'] = $parts[2];
  if ($parts[1] != '*') {
    $font['bold'] = strpos($parts[1], 'bold') !== FALSE;
    $font['italic'] = strpos($parts[1], 'italic') !== FALSE;
  }
  return $font;
}

function socialcalc_cellformat_parsecolor($cell, $sheet, $color) {
  if (!isset($cell[$color])) return FALSE;
  $parts = array();
  preg_match('/^rgb\((\d+?),(\d+?),(\d+?)\)\s*$/', $sheet['colors'][$cell[$color]], $parts);
  return array(
    'r' => $parts[1],
    'g' => $parts[2],
    'b' => $parts[3],
  );
}

