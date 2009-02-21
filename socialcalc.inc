<?php
// $Id$

/**
 * @file SocialCalc manipulation functions
 * Translated from socialcalc-3.js and companion files
 */

function socialcalc_parse($data) {
  $line = strtok($data, "\n");
  $sheet = array();
  while ($line !== FALSE) {
    $line = rtrim($line);
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
          $sheet['rowattribs']['hide'][$coord] = $parts[$j++];
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

function socialcalc_parse_values($data) {
  $line = strtok($data, "\n");
  $sheet = array();
  while ($line !== FALSE) {
    $line = rtrim($line);
    $parts = explode(':', $line);
    if ($parts[0] == 'cell') {
      $coord = $parts[1];
      $cell = isset($sheet['cells'][$coord]) ? $sheet['cells'][$coord] : array();
      $cell += socialcalc_parse_cell_value($parts, 2);
      $sheet['cells'][$coord] = $cell;
    }
    $line = strtok("\n");
  }
  return $sheet;
}

function socialcalc_parse_cell_value($parts, $j) {
  $cell = array();
  while ($t = $parts[$j++]) {
    switch ($t) {
    case "v":
      $cell['datavalue'] = doubleval(socialcalc_decode_value($parts[$j++]));
      break 2; // first time for me :-)
    case "t":
      $cell['datavalue'] = strval(socialcalc_decode_value($parts[$j++]));
      break 2;
    case "vt":
      $j++;
      $cell['datavalue'] = socialcalc_decode_value($parts[$j++]);
      break 2;
    case "vtf":
      $j++;
      $cell['datavalue'] = socialcalc_decode_value($parts[$j++]);
      break 2;
    case "vtc":
      $j++;
      $cell['datavalue'] = socialcalc_decode_value($parts[$j++]);
      break 2;
    }
  }
  return $cell;
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
      $sheet[] = 'col:'. substr($coord, 0, -1) .':w:'. $width;
    }
    if ($sc['colattribs']['hide']) foreach ($sc['colattribs']['hide'] as $col => $hide) {
      $coord = socialcalc_cr_to_coord($col, 1);
      $sheet[] = 'col:'. substr($coord, 0, -1) .':hide:'. $hide;
    }
  }
  if ($sc['rowattribs']) {
    if ($sc['rowattribs']['height']) foreach ($sc['rowattribs']['height'] as $row => $height) {
      $sheet[] = 'row:'. $row .':h:'. $height;
    }
    if ($sc['rowattribs']['hide']) foreach ($sc['rowattribs']['hide'] as $row => $hide) {
      $sheet[] = 'row:'. $row .':hide:'. $hide;
    }
  }
  if ($sc['fonts']) foreach ($sc['fonts'] as $fid => $font) {
    $sheet[] = 'font:'. $fid .':'. $font;
  }
  if ($sc['borderstyles']) foreach ($sc['borderstyles'] as $bsid => $borderstyle) {
    $sheet[] = 'border:'. $bsid .':'. $borderstyle;
  }
  if ($sc['cellformats']) foreach ($sc['cellformats'] as $cfid => $cellformat) {
    $sheet[] = 'cellformat:'. $cfid .':'. socialcalc_encode_value($cellformat);
  }
  if ($sc['layouts']) foreach ($sc['layouts'] as $lid => $layout) {
    $sheet[] = 'layout:'. $lid .':'. $layout;
  }
  if ($sc['colors']) foreach ($sc['colors'] as $cid => $color) {
    $sheet[] = 'color:'. $cid .':'. $color;
  }
  if ($sc['valueformats']) foreach ($sc['valueformats'] as $vfid => $valueformat) {
    $sheet[] = 'valueformat:'. $vfid .':'. socialcalc_encode_value($valueformat);
  }
  if ($sc['names']) foreach ($sc['names'] as $name => $nameinfo) {
    $sheet[] = 'name:'. socialcalc_encode_value($name) .':'. socialcalc_encode_value($nameinfo['desc']) .':'. socialcalc_encode_value($nameinfo['definition']);
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
    'r' => intval($parts[1]),
    'g' => intval($parts[2]),
    'b' => intval($parts[3]),
  );
}

function socialcalc_cellformat_parselayout($cell, $sheet) {
  if (!isset($cell['layout'])) return FALSE;
  $parts = array();
  preg_match('/^padding:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+);vertical-align:\s*(\S+);$/', $sheet['layouts'][$cell['layout']], $parts);
  $layout = array();
  if ($parts[1] != '*') $layout['padtop'] = str_replace('px', '', $parts[1]);
  if ($parts[2] != '*') $layout['padright'] = str_replace('px', '', $parts[2]);
  if ($parts[3] != '*') $layout['padbottom'] = str_replace('px', '', $parts[3]);
  if ($parts[4] != '*') $layout['padleft'] = str_replace('px', '', $parts[4]);
  if ($parts[5] != '*') $layout['alignvert'] = $parts[5];
  return $layout;
}

function socialcalc_cellformat_parseborder($cell, $sheet, $attrib) {
  if (!isset($cell[$attrib])) return FALSE;
  $parts = array();
  preg_match('/^(\S+)\s+(\S+)\s+(\S.+)$/', $sheet['borderstyles'][$cell[$attrib]], $parts);
  $border['thickness'] = $parts[1];
  $border['style'] = $parts[2];
  $color = array();
  preg_match('/^rgb\((\d+?),(\d+?),(\d+?)\)\s*$/', $parts[3], $color);
  $border['color'] = array(
    'r' => intval($color[1]),
    'g' => intval($color[2]),
    'b' => intval($color[3]),
  );
  return $border;
}

/**
 * formatinfo = socialcalc_parse_numberformat(format_string)
 *
 * format_string: (e.g., "#,##0.00_);(#,##0.00)")
 * formatinfo:
       .operators - array of operators from parsing the format string (each a number) and corresponding operands (each usually a string)
       .sectioninfo - one hash for each section of the format
          .sectionstart
          .integerdigits
          .fractiondigits
          .commas
          .percent
          .thousandssep
          .hasdate
          .hascurrency
       .hascomparison - true if any section has [<100], etc.

 **************************/

define(SOCIALCALC_COMMANDS_COPY, 'COPY');
define(SOCIALCALC_COMMANDS_COLOR, 'COLOR');
define(SOCIALCALC_COMMANDS_INTEGERPLACEHOLDER, 'INTEGER');
define(SOCIALCALC_COMMANDS_FRACTIONPLACEHOLDER, 'FRACTION');
define(SOCIALCALC_COMMANDS_DECIMAL, 'DECIMAL');
define(SOCIALCALC_COMMANDS_CURRENCY, 'CURRENCY');
define(SOCIALCALC_COMMANDS_GENERAL, 'GENERAL');
define(SOCIALCALC_COMMANDS_SEPARATOR, 'SEPARATOR');
define(SOCIALCALC_COMMANDS_DATE, 'DATE');
define(SOCIALCALC_COMMANDS_COMPARISON, 'COMPARISON');
define(SOCIALCALC_COMMANDS_SECTION, 'SECTION');
define(SOCIALCALC_COMMANDS_STYLE, 'STYLE');

function socialcalc_parse_numberformat($format_string) {
  $integerpart = true; // start out in integer part
  $lastwasinteger = false; // last char was an integer placeholder
  $lastwasslash = false; // last char was a backslash - escaping following character
  $lastwasasterisk = false; // repeat next char
  $lastwasunderscore = false; // last char was _ which picks up following char for width
  $inquote = false;
  $quotestr = ''; // processing a quoted string
  $inbracket = false;
  $bracketstr = '';
  $bracketdata = null; // processing a bracketed string
  $ingeneral = 0; // checks for characters "General"
  $ampmstr = ''; // checks for characters "A/P" and "AM/PM"
  $indate = ''; // keeps track of date/time placeholders

  $thisformat = array(
    'operators' => array(),
    'sectioninfo' => array(),
   ); // create info structure for this format

  $sectioninfo = array(); // get reference to info for current section
  $sectioninfo['sectionstart'] = 0; // position in operands that starts this section
  $sectioninfo['integerdigits'] = 0; // number of integer-part placeholders
  $sectioninfo['fractiondigits'] = 0; // fraction placeholders
  $sectioninfo['commas'] = 0; // commas encountered, to handle scaling
  $sectioninfo['percent'] = 0; // times to scale by 100

  for ($chpos=0; $chpos<strlen($format_string); $chpos++) { // parse
    $ch = $format_string[$chpos]; // get next char to examine
    if ($inquote) {
      if ($ch == '"') {
        $inquote = false;
        $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_COPY => $quotestr);
        continue;
      }
      $quotestr .= $ch;
      continue;
    }
    if ($inbracket) {
      if ($ch==']') {
        $inbracket = false;
        $bracketdata = socialcalc_parse_bracketformat($bracketstr);
        if (key($bracketdata['operator']) == SOCIALCALC_COMMANDS_SEPARATOR) {
          $sectioninfo['thousandssep'] = true; // explicit [,]
          continue;
        }
        if (key($bracketdata['operator']) == SOCIALCALC_COMMANDS_DATE) {
          $sectioninfo['hasdate'] = true;
        }
        if (key($bracketdata['operator']) == SOCIALCALC_COMMANDS_COMPARISON) {
          $thisformat['hascomparison'] = true;
        }
        if (key($bracketdata['operator']) == SOCIALCALC_COMMANDS_CURRENCY) {
          $sectioninfo['hascurrency'] = true;
        }
        $thisformat['operators'][] = $bracketdata['operator'];
        continue;
      }
      $bracketstr .= $ch;
      continue;
    }
    if ($lastwasslash) {
      $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_COPY => $ch);
      $lastwasslash = false;
      continue;
    }
    if ($lastwasasterisk) {
      $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_COPY => $ch.$ch.$ch.$ch.$ch); // do 5 of them since no real tabs
      $lastwasasterisk = false;
      continue;
    }
    if ($lastwasunderscore) {
      $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_COPY => '&nbsp;');
      $lastwasunderscore = false;
      continue;
    }
    if ($ingeneral) {
      if (substr('general', $ingeneral, 1) == strtolower($ch)) {
        $ingeneral++;
        if ($ingeneral == 7) {
          $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_GENERAL => $ch);
          $ingeneral = 0;
        }
        continue;
      }
      $ingeneral = 0;
    }
    if ($indate) { // last char was part of a date placeholder
      if (substr($indate, 0, 1) == $ch) { // another of the same char
        $indate .= $ch; // accumulate it
        continue;
      }
      $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_DATE => $indate); // something else, save date info
      $sectioninfo['hasdate'] = true;
      $indate = '';
    }
    if ($ampmstr) {
      $ampmstr .= $ch;
      $part = strtolower($ampmstr);
      if ($part != substr('am/pm', 0, strlen($part)) && $part != substr('a/p', 0, strlen($part))) {
        $ampstr = '';
      }
      else if ($part == 'am/pm' || $part == 'a/p') {
        $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_DATE => $ampmstr);
        $ampmstr = '';
      }
      continue;
    }
    if ($ch=='#' || $ch=='0' || $ch=='?') { // placeholder
      if ($integerpart) {
        $sectioninfo['integerdigits']++;
        if ($sectioninfo['commas']) { // comma inside of integer placeholders
          $sectioninfo['thousandssep'] = true; // any number is thousands separator
          $sectioninfo['commas'] = 0; // reset count of "thousand" factors
        }
        $lastwasinteger = true;
        $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_INTEGERPLACEHOLDER => $ch);
      }
      else {
        $sectioninfo['fractiondigits']++;
        $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_FRACTIONPLACEHOLDER => $ch);
      }
    }
    else if ($ch=='.') { // decimal point
      $lastwasinteger = false;
      $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_DECIMAL => $ch);
      $integerpart = false;
    }
    else if ($ch=='$') { // currency char
      $lastwasinteger = false;
      $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_CURRENCY => $ch);
      $sectioninfo['hascurrency'] = true;
    }
    else if ($ch==',') {
      if ($lastwasinteger) {
        $sectioninfo['commas']++;
      }
      else {
        $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_COPY => $ch);
      }
    }
    else if ($ch=='%') {
      $lastwasinteger = false;
      $sectioninfo['percent']++;
      $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_COPY => $ch);
    }
    else if ($ch=='"') {
      $lastwasinteger = false;
      $inquote = true;
      $quotestr = '';
    }
    else if ($ch=='[') {
      $lastwasinteger = false;
      $inbracket = true;
      $bracketstr = '';
    }
    else if ($ch=='\\') {
      $lastwasslash = true;
      $lastwasinteger = false;
    }
    else if ($ch=='*') {
      $lastwasasterisk = true;
      $lastwasinteger = false;
    }
    else if ($ch=='_') {
      $lastwasunderscore = true;
      $lastwasinteger = false;
    }
    else if ($ch==';') {
      $thisformat['sectioninfo'][] = $sectioninfo; // store current section
      $sectioninfo = array(); // create new section
      $sectioninfo['sectionstart'] = 1 + count($thisformat['operators']); // remember where it starts
      $sectioninfo['integerdigits'] = 0; // number of integer-part placeholders
      $sectioninfo['fractiondigits'] = 0; // fraction placeholders
      $sectioninfo['commas'] = 0; // commas encountered, to handle scaling
      $sectioninfo['percent'] = 0; // times to scale by 100
      $integerpart = true; // reset for new section
      $lastwasinteger = false;
      $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_SECTION => $ch);
    }
    else if (strtolower($ch)=='g') {
      $ingeneral = 1;
      $lastwasinteger = false;
    }
    else if (strtolower($ch)=='a') {
      $ampmstr = $ch;
      $lastwasinteger = false;
    }
    else if (strpos('dmyhHs', $ch) !== false) {
      $indate = $ch;
    }
    else {
      $lastwasinteger = false;
      $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_COPY => $ch);
    }
  }

  if ($indate) { // last char was part of unsaved date placeholder
    $thisformat['operators'][] = array(SOCIALCALC_COMMANDS_DATE => $indate);
    $sectioninfo['hasdate'] = true;
  }
  $thisformat['sectioninfo'][] = $sectioninfo;

  return $thisformat;
}


/**
 * bracketdata = socialcalc_parse_bracketformat(bracketstr)
 *
 * bracketstr: (e.g., "RED", ">10")
 * bracketdata: array(operator => operand)
 *
 ************************* */

function socialcalc_parse_bracketformat($bracketstr) {
  $allowedcolors = array(
    'BLACK' => "#000000", 
    'BLUE' => "#0000FF", 
    'CYAN' => "#00FFFF", 
    'GREEN' => "#00FF00", 
    'MAGENTA' => "#FF00FF",
    'RED' => "#FF0000", 
    'WHITE' => "#FFFFFF", 
    'YELLOW' => "#FFFF00",
  );
  $alloweddates = array(
    'H' => "h]", 
    'M' => "m]", 
    'MM' => "mm]", 
    'S' => "s]", 
    'SS' => "ss]",
  );

  $bracketdata = array();
  $parts = array();

  if (substr($bracketstr, 0, 1)=='$') { // currency
    if (preg_match('/^\$(.+?)(\-.+?){0,1}$/', $bracketstr, $parts)) {
      $operand = isset($parts[1]) ? $parts[1] : '$';
    }
    else {
      $operand = substr($bracketstr, 1) ? substr($bracketstr, 1) : '$';
    }
    $bracketdata['operator'] = array(SOCIALCALC_COMMANDS_CURRENCY => $operand);
  }
  else if ($bracketstr=='?$') {
    $bracketdata['operator'] = array(SOCIALCALC_COMMANDS_CURRENCY => '[?$]');
  }
  else if (isset($allowedcolors[strtoupper($bracketstr)])) {
    $bracketdata['operator'] = array(SOCIALCALC_COMMANDS_COLOR => $allowedcolors[strtoupper($bracketstr)]);
  }
  else if (preg_match('/^style=([^"]*)$/', $bracketstr, $parts)) { // [style=...]
    $bracketdata['operator'] = array(SOCIALCALC_COMMANDS_STYLE => $parts[1]);
  }
  else if ($bracketstr==',') {
    $bracketdata['operator'] = array(SOCIALCALC_COMMANDS_SEPARATOR => $bracketstr);
  }
  else if ($alloweddates[strtoupper($bracketstr)]) {
    $bracketdata['operator'] = array(SOCIALCALC_COMMANDS_DATE => $alloweddates[strtoupper($bracketstr)]);
  }
  else if (preg_match('/^[<>=]/', $bracketstr)) { // comparison operator
//    preg_match('/^([<>=]+)(.+)$/', $bracketstr, $parts); // split operator and value
//    $bracketdata['operator'] = array(SOCIALCALC_COMMANDS_COMPARISON => $parts[1].':'.$parts[2]);
    $bracketdata['operator'] = array(SOCIALCALC_COMMANDS_COMPARISON => $bracketstr);    
  }
  else { // unknown bracket
    $bracketdata['operator'] = array(SOCIALCALC_COMMANDS_COPY => '['.$bracketstr.']');
  }
  return $bracketdata;
}

