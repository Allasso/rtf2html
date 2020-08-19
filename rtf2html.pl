#!/usr/bin/perl

## This file is part of RTF2HTML by Allasso Travesser.
##
## RTF2HTML by Allasso Travesser
## Copyright 2011 by Allasso Travesser under the GNU GPL license
##
## RTF2HTML by Allasso Travesser is free software: you can redistribute
## it and/or modify it under the terms of the GNU General Public License
## as published by the Free Software Foundation, either version 3 of the
## License, or (at your option) any later version.
##
## RTF2HTML by Allasso Travesser is distributed in the hope that it will
## be useful, but WITHOUT ANY WARRANTY; without even the implied warranty
## of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
## GNU General Public License for more details.
##
## You should have received a copy of the GNU General Public License
## along with RTF2HTML by Allasso Travesser.
## If not, see <http://www.gnu.org/licenses/>.
##
## If you use this file in another program, you must include
## this blurb.

use Encode;
use strict;
use Data::Dumper;

####################################################
####### Usage and default output

my $USAGE = q(
  Usage:

  rtf2html <sourcefile> <args>

  <sourcefile> required - The RTF file to convert.  Must have .rtf or .rtfd extension (case insensitive).
                          Will write output file to same directory, with same name appended with .html

  <args> optional:

  -html_entities         convert non-ascii charaters to HTML entities
  -no_smartquotes        convert smart quotes to ascii quotes
  -ascii_only_punct      attempt to convert non-ascii punctuation to ascii eqivalent
  -templatefile          name of template file to use
  -mult_spaces           allow multi-spaces by converting them to &nbsp;.
                         Default collapses sequential spaces to a single space
  -sngl_line_pars_use_break  consecutive paragraphs separated by single line break are converted to a single par
                         and <br> is inserted at the end of each line (currently disabled)
  -resolve_to_bq         resolve ambiguous paragraphs to <p class="bq"> (default: resolve to <p>)
  -allow_leading_spaces  allow leading spaces at beginning of paragraph (default no)
  -expand_tabs           expand tabs to n number of &nbsp;'s. If n = 0 or empty, treat tab as space
  -default_color         color string in the form of r,g,b.  Leave empty for default, 0,0,0 (black)
  -default_font_size     font size value in half-points (as used by the RTF spec)
  -default_font_family   font family in camel case format

  Indents: (Indents are set as: css margin-left for RTF left-indent, css margin-right for RTF right-indent,
  and css  for RTF first-indent.)

  -indent_thresholds     set thresholds for indent, where value at or above threshold will be set
                         to a fixed value.  Settings are comma-separated pairs, where the first value
                         is the threshold in pixels, and the second value is the fixed value in pixels
                         that will be used if at or above the threshold value.  Multiple settings can
                         be used by separating with a . (dot), where the fixed value will be used for
                         the highest threshold.  Example:

                         thr1,fxd1.thr2,fxd2.thr3,fxd3...

                         Note, any value below the lowest threshold will default to 0 (no indent). Also,
                         threshold/fixed value pairs will apply to first-indent, left-indent, and right-indent.
                         Default will be to use the values in the RTF.
  -blockquote_factor      when set, if left-indent value is provided in the RTF for a paragraph, right-indent will
                         be set to a value proportional to left-indent. This behavior will occur even in the absence
                         -use_right_indent, and any right-indent supplied in the RTF will be overridden by this value.
                         If paragraph has no left-indent value, right-indent will be handled normally according to
                         RTF values and -use_right_indent.
  -use_first_indent      use first-indent values (set as  in css)
  -use_right_indent      use right-indent values (set as margin-right in css), otherwise ignore.  If not set and
                         a value is set for -blockquote_factor, -blockquote_factor will still be calculated and used.
  -text_align            default text-align property.  text-align properties are determined from the RTF and css classes
                         will be generated for all those properties except the default.
                         Can be l, r, j, or c, where l is left, r is right, j is justified, and c is center.  If not set,
                         default is l (left).  Values of "left", "right", "justified" or "center" may also be used.
  -resolve_ambiguities   attempt to resolve ambiguities in paragraph formatting
  -resolve_amb
  -print_raw_text        print raw text after conversion from marked-up RTF string
  -raw_text
  -compare_rtf           compare RTF output
 );

if (scalar @ARGV == 0 || get_arg("-help")) {
  print $USAGE;
  exit;
}

####################################################
####### GET ARGS AND SET INTO GLOBAL CONSTANTS

my $_sourcefile = shift(@ARGV);

my $HTMLENTITIES = get_arg("-html_entities");
my $NO_SMARTQUOTES = get_arg("-no_smartquotes");
my $ASCII_ONLY_PUNCT = get_arg("-ascii_only_punct");
my $TEMPLATEFILE = get_arg("-templatefile", 2);
my $MULTSPACES = get_arg("-mult_spaces");
my $SNGL_LINE_PARS_USE_BR = get_arg("-sngl_line_pars_use_break");
my $RESOLVE_TO_BQ = get_arg("-resolve_to_bq");
my $ALLOW_LEAD_SPACES = get_arg("-allow_leading_spaces");
my $EXPAND_TABS = get_arg("-expand_tabs");
my $DEFAULT_COLOR = get_arg("-default_color", 2);
my $DEFAULT_FONT_SIZE = get_arg("-default_font_size", 2);
my $DEFAULT_FONT_FAMILY = get_arg("-default_font_family", 2);
my $INDENT_THRESHOLDS = get_arg("-indent_thresholds", 2);
my $BLOCKQUOTE_FACTOR = get_arg("-blockquote_factor", 2);
my $FIRST_INDENT = get_arg("-use_first_indent");
my $RIGHT_INDENT = get_arg("-use_right_indent");
my $DEFAULT_TEXT_ALIGN = get_arg("-text_align", 2);
my $RESOLVE_PARAGRAPH_AMBIGUITIES = get_arg("-resolve_ambiguities") || get_arg("-resolve_amb");
my $RAW_TEXT_OUTPUT = get_arg("-print_raw_text") || get_arg("-raw_text");
my $RTF_COMPARE_OUTPUT = get_arg("-compare_rtf");

if ($DEFAULT_TEXT_ALIGN eq 'left') {
  $DEFAULT_TEXT_ALIGN = 'l';
}
if ($DEFAULT_TEXT_ALIGN eq 'right') {
  $DEFAULT_TEXT_ALIGN = 'r';
}
if ($DEFAULT_TEXT_ALIGN eq 'justify') {
  $DEFAULT_TEXT_ALIGN = 'j';
}
if ($DEFAULT_TEXT_ALIGN eq 'center') {
  $DEFAULT_TEXT_ALIGN = 'c';
}

#print("WARNING : this is a test\n\n");

####################################################


####################################################
####### SET FILENAME VARIABLES

my $temp_rtf_file_1 = '/tmp/748559688736_tmp_0999_1.rtf';
$_sourcefile =~ s@/+$@@;

my $outfile = $_sourcefile.'.html';

if ($_sourcefile =~ /\.rtfd$/i) {
  $_sourcefile = $_sourcefile.'/TXT.rtf';
}

####################################################

if (!$_sourcefile) { print $USAGE; exit; }


####################################################
####### SET GLOBAL CONSTANTS

## don't want to use the 'use constant' pragma so
##   it will make it easier to use these in m// & s///

my $DIMFACTOR   = "20";    # dimension factor for \li, \ri, \fi
my $DIMUNITS    = "px";    # dimension units

my $LB          = "<br />";
my $LISTITEM    = "<li>";
my $LISTITEMOFF = "</li>";
my $BQ_ON       = "<blockquote>";
my $BQ_OFF      = "</blockquote>";
my $PAR_ON      = "<p>";
my $PAR_BQ      = "<p class=\"bq\">";
my $PAR_OFF     = "</p>";
my $IND_ON      = "<p class=\"indent\">";
my $IND_OFF     = "</p>";
my $ITL_ON      = "<em>";
my $ITL_OFF     = "</em>";
my $BLD_ON      = "<strong>";
my $BLD_OFF     = "</strong>";
my $UNL_ON      = "<u>";
my $UNL_OFF     = "</u>";
my $SUP_ON      = "<sup>";
my $SUP_OFF     = "</sup>";
my $SUB_ON      = "<sub>";
my $SUB_OFF     = "</sub>";
my $CLR_ON      = "<span class=\"clrCLASSREF\">";
my $CLR_OFF     = "</span>";


## commonly used REGEXES:

my $WS_HOR = '(?:\h|\[TAB_SUB\]|&nbsp;)';
my $WS_ALL = '(?:\s|\[TAB_SUB\]|&nbsp;)';
my $NC_HOR = '(?:\h|\[TAB_SUB\]|&nbsp;|<[^<>]*>)';
my $NC_ALL = '(?:\s|\[TAB_SUB\]|&nbsp;|<[^<>]*>)';
my $P_OPEN = '(?:<p>|<p [^<>]+>)';

####################################################

####################################################
####### DECLARE GLOBAL VARIABLES

my @__color_table_array;
my %__color_table_hash;
my %__par_table_hash;
my $__css_str = '';
my $__font_css = "";
my %__font_hash = ();

####################################################

####################################################
####### READ IN RTF FILE

my $holdTerminator = $/;
local $/;
open(FH, $_sourcefile);
my $_filestring = <FH>;
$/ = $holdTerminator;
close(FH);

####################################################

## EXTRACT COLORTBL FROM RTF AND CREATE CSS
extractColortbl($_filestring);

## EXTRACT FONTTBL FROM RTF AND CREATE CSS, CONSOLIDATE REDUNDANT FONT REFS
$_filestring = extractFonttbl($_filestring);

## EXTRACT LISTS
my $listinfo = extractLists($_filestring);
$__css_str .= $listinfo->[0];


## REMOVE NON-SECTION GROUPS
$_filestring = removeNonSections($_filestring);


## CONVERT HTML CODE CHARACTERS TO ENTITIES
$_filestring =~ s@\x26@&#38;@g;
$_filestring =~ s@\x3c@&#60;@g;
$_filestring =~ s@\x3e@&#62;@g;


## INSERT LIST HTML
$_filestring = insertListHTML($_filestring,$listinfo);

# # #print $_filestring;
# # #exit;

# # #interceptRTF($_filestring);


## MULTI SPACE \sbN and \saN PARAGRAPHS
$_filestring = multiSpacePars($_filestring);


## INSERT PAARAGRAPH HTML
$_filestring = insertParHTMLmultiBQ($_filestring);

## INSERT COLOR SPANS
$_filestring = insertColorSpans2($_filestring);     ## we're using method #2

## INSERT FONT SIZE SPANS
my $_filestring = insertFontSizeSpans($_filestring);

## INSERT FONT FAMILY SPANS
my $_filestring = insertFontFamilySpans($_filestring);

## INSERT IMAGE PLACEHOLDERS (as html comments)
#$_filestring =~ s@(\{\\\*\\shppict\{\\pict)@ IMAGE_PLACEHOLDER $1@g;


## INSERT TOGGLE STYLES
$_filestring = insertInlineTags($_filestring,'super','nosupersub',$SUP_ON,$SUP_OFF);
$_filestring = insertInlineTags($_filestring,'sub','nosupersub',$SUB_ON,$SUB_OFF);
$_filestring = insertInlineTags($_filestring,'i','i0',$ITL_ON,$ITL_OFF);
$_filestring = insertInlineTags($_filestring,'b','b0',$BLD_ON,$BLD_OFF);
$_filestring = insertInlineTags($_filestring,'ul','ulnone',$UNL_ON,$UNL_OFF);


## REPLACE '\U8232' CODE SECTIONS WITH '\NEWLINE' (makes up for deficiency in textutil)
#$_filestring =~ s@\\u8232@\\\n@g;



####################################################


#print STDERR $__css_str;
#print $_filestring;


####################################################
## CONVERT MARKED UP RTF FILE TO PARTIALLY MARKED UP TEXT

open (FH, '>', $temp_rtf_file_1);
print FH $_filestring;
close (FH);

my $_textstring = `textutil -convert txt -stdout $temp_rtf_file_1`;

if ($RAW_TEXT_OUTPUT) {
  $_textstring =~ s@\n\n\n+@\n\n@g;
  print $_textstring;
  exit;
}

if ($RTF_COMPARE_OUTPUT) {
  $_textstring = formatForRTFCompare($_textstring);
  $_textstring = utf8_to_num_html($_textstring);

  print $_textstring;
  exit;
}

#print $_textstring;

####################################################


####################################################
## PROCESS PARTIAL HTML STRING

#produceOutput("$_textstring");

## REPLACE SPECIAL CHAR SEQUENCES
$_textstring = utf8_to_num_html($_textstring);

$_textstring =~ s@\f+@\n\n@g;
$_textstring =~ s@\n\n\n+@\n\n@g;

$_textstring = defineParagraphBlocks($_textstring);
$_textstring = formatHTML($_textstring);

## CONVERT SPECIAL TAGS TO STANDARD HTML
$_textstring = convertTagsAndGenerateCss($_textstring);

sub convertTagsAndGenerateCss {
  my $ts = shift;

  my $ts_out = '';
  my %color_styles_used_hash;
  my %font_styles_used_hash;
  my @font_styles_used_array;
  my %p_styles_used_hash;
  my @p_styles_used_array;
  my %text_align_used_hash;
  my %sbp_used_hash;

  my $arr_str = $ts;
  $arr_str =~ s@(<clr[0-9]+>)@RCRDSPRTR$1RCRDSPRTR@g;
  $arr_str =~ s@(<p[0-9][^>]*>)@RCRDSPRTRBEGINPAR1$1RCRDSPRTR@g;
  $arr_str =~ s@(<p [^>]+>)@RCRDSPRTRBEGINPAR2$1RCRDSPRTR@g;
  $arr_str =~ s@(<fs[0-9]+>)@RCRDSPRTR$1RCRDSPRTR@g;
  $arr_str =~ s@(<ff[0-9]+>)@RCRDSPRTR$1RCRDSPRTR@g;

  for my $seg (split(/RCRDSPRTR/, $arr_str)) {
    if (index($seg, '<clr') == 0) {
      my $ref = $seg;
      $ref =~ s@<clr@@;
      $ref =~ s@>@@;
      my $val = $__color_table_hash{$ref};
      $seg = '<span class="clr_' . $val . '">';
      $color_styles_used_hash{$val} = 1;
    }
    if (index($seg, '<fs') == 0) {
      my $val = $seg;
      $val =~ s@<@@;
      $val =~ s@>@@;
      $seg = '<span class="' . $val . '">';

      if (!$font_styles_used_hash{$val}) {
        push(@font_styles_used_array, $val);
      }
      $font_styles_used_hash{$val} = 1;
    }
    if (index($seg, '<ff') == 0) {
      my $val = $seg;
      $val =~ s@<@@;
      $val =~ s@>@@;
      $seg = '<span class="ff_' .$__font_hash{$val} . '">';
    }
    if ($seg =~ s@BEGINPAR1@@) {
      my $clas_str = $seg;
      $clas_str =~ s@<p@@;
      $clas_str =~ s@>@@;
      my @clas_arr = split(/ /, $clas_str);
      my $ref = $clas_arr[0];
      my $val = $__par_table_hash{$ref};

      if (!$p_styles_used_hash{$val}) {
        push(@p_styles_used_array, $val);
      }
      $p_styles_used_hash{$val} = 1;

      my $ta_val = '';

      if ($val =~ s@_(ta[lrjc])@@) {
        $text_align_used_hash{$1} = 1;
        $ta_val = $1;
      }

      if (!$val and !$ta_val and !$clas_arr[1]) {
        next;
      }

      my $class_exp = '';
      if ($val) {
        $class_exp = 'p' . $val;
      }
      if ($ta_val) {
        if ($class_exp) {
          $class_exp .= ' ';
        }
        $class_exp .= $ta_val;
      }
      if ($clas_arr[1]) {
        if ($class_exp) {
          $class_exp .= ' ';
        }
        $class_exp .= $clas_arr[1];
        $sbp_used_hash{$clas_arr[1]} = 1;
      }
      $seg = '<p class="' . $class_exp . '">';
    }
    if ($seg =~ s@BEGINPAR2@@) {
      my $val = $seg;
      $val =~ s@<p @@;
      $val =~ s@>@@;
      $seg = '<p class="' . $val . '">';
    }
    $ts_out .= $seg;
  }

  # Generate css for paragraph formatting:

  for my $p_class (sort @p_styles_used_array) {
    $p_class =~ s@_ta[lrjc]@@;
    if (!$p_class) {
      next;
    }
    my $val = $p_class;
    $val =~ s@_ml([-0-9]+)@margin-left: $1px ; @;
    $val =~ s@_mr([-0-9]+)@margin-right: $1px ; @;
    $val =~ s@_ti([-0-9]+)@text-indent: $1px ; @;
    $__css_str .= 'p.p' . $p_class . ' { ' . $val . "}\n";
print($p_class . '    ' . $val . "\n");
  }

  if ($text_align_used_hash{'tal'}) {
print("p.tal\n");
    $__css_str .= "p.tal { text-align: left ; }\n";
  }
  if ($text_align_used_hash{'tar'}) {
print("p.tar\n");
    $__css_str .= "p.tar { text-align: right ; }\n";
  }
  if ($text_align_used_hash{'taj'}) {
print("p.taj\n");
    $__css_str .= "p.taj { text-align: justify ; }\n";
  }
  if ($text_align_used_hash{'tac'}) {
print("p.tac\n");
    $__css_str .= "p.tac { text-align: center ; }\n";
  }

  if ($sbp_used_hash{'sbpf'}) {
print("p.sbpf\n");
    $__css_str .= "p.sbpf { margin-bottom: 0 ; }\n";
  }
  if ($sbp_used_hash{'sbpm'}) {
print("p.sbpm\n");
    $__css_str .= "p.sbpm { margin-top: 0 ; margin-bottom: 0 ; }\n";
  }
  if ($sbp_used_hash{'sbpl'}) {
print("p.sbpl\n");
    $__css_str .= "p.sbpl { margin-top: 0 ; }\n";
  }

  # Generate css for text color formatting:

  for my $clr_class (sort @__color_table_array) {
    if ($color_styles_used_hash{$clr_class}) {
      my $val = $clr_class;
      $val =~ s@_@, @g;
      $__css_str .= '.clr_' . $clr_class . ' { color: rgb(' . $val . "); }\n";
print('clr: ' . $val . "\n");
    }
  }

  # Generate css for font formatting:

  for my $font_class (sort @font_styles_used_array) {
    my $units;
    my $val = $font_class;
    $val =~ s@fs@@g;

    if ($DEFAULT_FONT_SIZE) {
      $val = int(100 * $val / $DEFAULT_FONT_SIZE);
      $units = '%';
    } else {
      $val = $val / 2;
      $units = 'pt';
    }

    $__css_str .= '.' . $font_class . ' { font-size: ' . $val . $units . " ; }\n";
print($font_class . ": " . $val . "\n");
  }

  $__css_str .= $__font_css;

print($__css_str);

  # Convert closing tags:
  $ts_out =~ s@</clr[0-9]+>@</span>@g;
  $ts_out =~ s@</p[0-9]+>@</p>@g;
  $ts_out =~ s@</fs[0-9]+>@</span>@g;
  $ts_out =~ s@</ff[0-9]+>@</span>@g;

  # A fair amount of fs and ff spans are adjacent, so we'll catch at least those:
  $ts_out =~ s@<span class="(fs[0-9]+)"><span class="(ff_[a-zA-Z]+)">((?:(?!<span|</span>).)*)</span></span>@<span class="$2 $1">$3</span>@gs;
  $ts_out =~ s@<span class="(ff_[a-zA-Z]+)"><span class="(fs[0-9]+)">((?:(?!<span|</span>).)*)</span></span>@<span class="$1 $2">$3</span>@gs;

  return $ts_out;
}

#print $_textstring . "\n";

## TODO: Don't know if this stuff is needed any longer.

## ORGANIZE WHITESPACE
my $rtrn_ref = whitespace($_textstring);
my $TAB_REPLACEMENT = $rtrn_ref->[0];
$_textstring        = $rtrn_ref->[1];

## SINGL-FY BACK TO BACK CONTAINERS
#$_textstring = singlfy_back_2_back_containers($_textstring);


## RESOLVE CONTAINERS WHICH SPAN ACROSS BLOCK ELEMENTS (<p>)
#$_textstring = formatHTML($_textstring);


## PULL OPENING AND CLOSING P TAGS TO PARAGRAPH PERIMETERS (again)
$_textstring =~ s@(</p>)(.+?)(\n)@$2$1$3@g;
$_textstring =~ s@(\n)(.+?)(<p(?= |>)[^>]*>)@$1$3$2@g;


## NEST TAGS IN LIST BLOCKS
$_textstring = nestTagsList($_textstring);


## MARK UP LISTS
$_textstring = markUpLists($_textstring);


## RESOLVE PARAGRAPH AMIBIGUITIES
# TODO: logic on this is reversed for development.
if (!$RESOLVE_PARAGRAPH_AMBIGUITIES) { $_textstring = paragraphAbiguities($_textstring); }

####################################################
####### FINALLY.......


## INSERT INTER-PARAGRAPH LINEBREAKS
$_textstring =~ s@(?<!</p>)(?<!\n)\n@$LB\n@g;
$_textstring =~ s@^$LB\n@@;


## COLLAPSE SBPs TO SINGLE LINE SPACING
$_textstring =~ s@(<p[^>]*sbp[fm][^>]*>(?:(?!<p|</p>).)*</p>)\n@$1@sg;


## EXPAND TABS INTO SPACES
if ($EXPAND_TABS) { $_textstring =~ s@\[TAB_SUB\]@$TAB_REPLACEMENT@g; }


## TRIM
$_textstring =~ s@^\n+@@;
$_textstring =~ s@\n+$@\n@;


## REMOVE UNWANTED HORIZONTAL WHITESPACE
$_textstring =~ s@$WS_HOR+\n@\n@g;
if (!$ALLOW_LEAD_SPACES) {
  $_textstring =~ s@($P_OPEN)$WS_HOR+@$1@g;
  $_textstring =~ s@\n$WS_HOR+@\n@g;
}


## INSERT INTO TEMPLATE (IF DESIRED) AND PRINT TO OUTPUT FILE
my $outputstring = '';

if ($TEMPLATEFILE && -f $TEMPLATEFILE) {

  open (FH, "$TEMPLATEFILE");
  my $holdTerminator = $/;
  local $/;
  my $templatestring = <FH>;
  $/ = $holdTerminator;
  close (FH);

  if ($templatestring =~ s@<!-- *ADD STYLES HERE *-->\n{0,1}@$__css_str@ && $templatestring =~ s@<!-- *ADD CONTENT HERE *-->\n{0,1}@$_textstring@) {
      $outputstring = "$templatestring";
  }else{
    $outputstring = "TEMPLATE SELECT HAD INVALID FORMAT  --  NOT USED\n";
    $outputstring .= "template must include the strings:\n";
    $outputstring .= "<!--ADD STYLES HERE-->\n";
    $outputstring .= "<!--ADD CONTENT HERE-->\n";
    $outputstring .= "for proper style and content insertion\n\n\n";
    $outputstring .= "<style>\n$__css_str</style>\n\n\n$_textstring";
  }
}else{
  $outputstring = '<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />' . "\n";
  $outputstring .= '<style type="text/css">'."\n".$__css_str.'</style>'.
    "\n\n<!-- BEGIN HTML SECTION -->\n\n".$_textstring;
}

produceOutput("$outputstring");


########################################################################################################
########################################################################################################
########################################################################################################
########################################################################################################
##
##      SUBROUTINES
##
#################################################################
#################################################################
#################################################################
#################################################################


#################################################################
#################################################################
#################################################################
##
##      PRODUCE OUTPUT
##

sub produceOutput() {

  my $os = shift;

  open (FH, ">$outfile");
  print FH "$os";
  close (FH);


  ## UNLINK TEMP FILE
  unlink("$temp_rtf_file_1");


  ## PRINT MESSAGE FOR JAVA APP
  print("OutputFile: $outfile\n");        # _NOT_ DIAGNOSTIC !!!      NEEDED BY JAVA APP


  ## GOODBYE !

  print("goodbye !!\n\n");
  exit;

}


#################################################################
#################################################################
#################################################################
##
##      RTF PROCESSING
##


####################################################
####### SUBROUTINE@  EXTRACT FONTTBL FROM RTF AND CREATE CSS -
####### ELIMINATE REDUNDANT FONT FAMILY REFERENCES

#{\fonttbl\f0\fnil\fcharset0 Verdana-Bold;\f1\fnil\fcharset0 Verdana;\f2\fnil\fcharset0 Verdana-Italic;
#\f3\fnil\fcharset0 GillSans-Italic;\f4\fnil\fcharset0 GillSans;\f5\fnil\fcharset0 Verdana-BoldItalic;
#\f6\fnil\fcharset0 GillSans-Bold;\f7\fnil\fcharset0 GillSans-BoldItalic;}

#\fonttbl\f0\fnil\fcharset0 Verdana-Bold
#\f1\fnil\fcharset0 Verdana
#\f2\fnil\fcharset0 Verdana-Italic
#\f3\fnil\fcharset0 GillSans-Italic
#\f4\fnil\fcharset0 GillSans
#\f5\fnil\fcharset0 Verdana-BoldItalic
#\f6\fnil\fcharset0 GillSans-Bold
#\f7\fnil\fcharset0 GillSans-BoldItalic

sub extractFonttbl() {
  my $fs = shift;
  my $ftbl = "";

  if ($fs =~ m@({\\fonttbl[^}]*})@) {
    $ftbl = $1;
    $ftbl =~ s@{|}@@g;
    $ftbl =~ s@\\fonttbl@@g;
    $ftbl =~ s@\n@@g;

    my %families_hash = ();

    for my $seg (split(/;/, $ftbl)) {
      my $fam_name = $seg;
      $fam_name =~ s@.* @@;
      $fam_name =~ s@-.*@@;
      my $fam_ref = $seg;
      $fam_ref =~ s@^\\@@;
      $fam_ref =~ s@\\.*@@;

      my @subsegs = split(/ /, $seg);

      if (!$families_hash{$fam_name}) {
        $families_hash{$fam_name} = [];
      }
      push($families_hash{$fam_name}, $fam_ref);
    }

    for my $fam_name (keys %families_hash) {
      my @fam_arr = @{$families_hash{$fam_name}};
      my $len = scalar(@fam_arr);

      # Make family name css compatible by adding spaces between camelcase'ed caps.
      my $fam_css_name = $fam_name;
      $fam_css_name =~ s@([A-Z])@ $1@g;
      $fam_css_name =~ s@^ @@g;

      $__font_css .= ".ff_" . $fam_name . " { font-family: " . $fam_css_name . " ; }\n";

      my $repl = "";
      for my $ref (@fam_arr) {
        if (!$repl) {
          $repl = $ref;
          my $index = $ref;
          $index =~ s@\\@@;
          $__font_hash{"f" . $ref} = $fam_name;
        } else {
          $fs =~ s@\\$ref([^0-9])@\\$repl$1@g;
        }
      }
    }
  }

  return $fs;
}

####################################################
####### SUBROUTINE@  EXTRACT COLORTBL FROM RTF AND CREATE CSS

my $DEFAULT_COLOR_REFERENCE;

sub extractColortbl() {
  my $fs = shift;
  my $cw = 'colortbl';
  my $seg = '';

  my $red_default = 0;
  my $green_default = 0;
  my $blue_default = 0;

  if ($DEFAULT_COLOR) {
    my @default_arr = split(',', $DEFAULT_COLOR);

    $red_default = $default_arr[0];
    $green_default = $default_arr[1];
    $blue_default = $default_arr[2];

    $red_default =~ s@^\s+|\s+$@@g;
    $green_default =~ s@^\s+|\s+$@@g;
    $blue_default =~ s@^\s+|\s+$@@g;
  }

  if ($fs =~ m@(\\$cw)@) {

    #print $1."\n";

    $fs =~ s@.*?({[^{}]*\\$cw)@@s;
    $seg = $1;
    #print $leader."\n\n";
    my $cnt = 1;
    until ($cnt == 0) {

      if    ($fs =~ s@^([^{}]*{)@@) {
        $seg .= $1;
        $cnt++;
      }
      elsif ($fs =~ s@^([^{}]*})@@) {
        $seg .= $1;
        $cnt--;
      }
    }

    my $cnt = 1;
    if ($seg =~ m@\\colortbl[^;]*(?:red|green|blue)@) { $cnt = 0; } ## if clr 0 is not autocolor
    my %used_classes = ();

    for my $cd (split /;/, $seg) {

      my $red;
      my $grn;
      my $blu;

      if ($cd =~ /red([0-9]+)/)   { $red = $1; }
      if ($cd =~ /green([0-9]+)/) { $grn = $1; }
      if ($cd =~ /blue([0-9]+)/)  { $blu = $1; }
      if ($red ne '' && $grn ne '' && $blu ne '') {
        # Flag insertColorSpans2 to not create tags for this class.
        if ($red eq $red_default && $grn eq $green_default && $blu eq $blue_default) {
          $DEFAULT_COLOR_REFERENCE = $cnt . '';
        }

        # We're still going to let it create the css even if it's the default.

        #my $red_padded = "00" . $red;
        #my $grn_padded = "00" . $grn;
        #my $blu_padded = "00" . $blu;
        #$red_padded =~ s@.*(...)@$1@;
        #$grn_padded =~ s@.*(...)@$1@;
        #$blu_padded =~ s@.*(...)@$1@;
        #my $padded_val = $red_padded . $grn_padded . $blu_padded;

        #my $red_hex = sprintf("%02x", $red);
        #my $grn_hex = sprintf("%02x", $grn);
        #my $blu_hex = sprintf("%02x", $blu);
        #my $hex_val = $red_hex . $grn_hex . $blu_hex;

        my $rgb_class = $red . "_" . $grn . "_" . $blu;
        my $rgb_val = $red . "," . $grn . "," . $blu;

        $__color_table_hash{$cnt} = $rgb_class;
        if (!$used_classes{$rgb_class}) {
          push(@__color_table_array, $rgb_class);
        }
        $used_classes{$rgb_class} = 1;

        $cnt++;
      }
    }
  }
}

####################################################


####################################################
####### SUBROUTINE@  EXTRACT LISTS

sub extractLists () {

  my %ul_mrkr_hash = (
    'circle'  => 'circle',
    'square'  => 'square',
    'box'     => 'square',
    'diamond' => 'square',
  );

  my %ol_styl_hash = (
    '1'       => 'upper-roman',
    '2'       => 'lower-roman',
    '3'       => 'upper-alpha',
    '4'       => 'lower-alpha',
    '22'      => 'decimal-leading-zero',
  );

  ## import RTF _filestring
  my $fs = shift;

  ## first, let's create a map from the \listoverridetable
  ## which associates a \lsN control word with \listidN
  ##
  ## each list in the document has a reference \lsN.
  ## The \listoverridetable maps this reference to the \listidN,
  ## which points us to the proper \list in the \listtable

  my $cw          = 'listoverridetable';
  my $seg         = '';
  my %ovrd_hash   = ();

  if ($fs =~ m@(\\$cw)@) {

    my $tmp_fs = $fs;      ## preserve $fs, work on $tmp_fs instead

    ## extract the $cw group

    $tmp_fs =~ s@.*?({[^{}]*\\$cw)@@s;
    $seg = $1;
    my $cnt = 1;
    until ($cnt == 0) {
      if    ($tmp_fs =~ s@^([^{}]*{)@@) {
        $seg .= $1;
        $cnt++;
      }
      elsif ($tmp_fs =~ s@^([^{}]*})@@) {
        $seg .= $1;
        $cnt--;
      }
    }

    $seg =~ s@\\listoverride(?![a-zA-Z0-9-])@RCRDSPRTRBEGINLIST\\listoverride@g;

    for my $ovrd_cand (split(/RCRDSPRTR/, $seg)) {

      if ($ovrd_cand =~ m@^BEGINLIST@) {
        if ($ovrd_cand =~ m@\\listid([0-9]+)@) {
          my $list_id = $1;
          if ($ovrd_cand =~ m@\\ls([0-9]+)@) {
            $ovrd_hash{$1} = $list_id;
          }
        }
      }
    }
  }

  my $cw          = 'listtable';
  my $seg         = '';
  my $ct          = '';
  my %ltyp_hash   = ();
  my $list_css    = '';

  if ($fs =~ m@(\\$cw)@) {

    my $tmp_fs = $fs;      ## preserve $fs, work on $tmp_fs instead

    ## extract the $cw group

    $tmp_fs =~ s@.*?({[^{}]*\\$cw)@@s;
    $seg = $1;
    my $cnt = 1;
    until ($cnt == 0) {
      if    ($tmp_fs =~ s@^([^{}]*{)@@) {
        $seg .= $1;
        $cnt++;
      }
      elsif ($tmp_fs =~ s@^([^{}]*})@@) {
        $seg .= $1;
        $cnt--;
      }
    }

    $seg =~ s@\\list(?![a-zA-Z0-9-])@RCRDSPRTRBEGINLIST\\list@g;

    for my $list_cand (split(/RCRDSPRTR/, $seg)) {

      if ($list_cand =~ m@^BEGINLIST@) {

        my $list_styl = '';
        my $list_type = 'ul';

        if ($list_cand =~ m@(?:levelnfcn|levelnfc)([0-9]+)@) {

          my $styl_num = $1;

          if ($styl_num == 23 || $styl_num == 255) {

            if ($list_cand =~ m@.*?\\levelmarker[^{}]*\\{([^{}]+)\\}@) {
              if ($ul_mrkr_hash{$1}) { $list_styl = $ul_mrkr_hash{$1}; }
            }
          }else{

            $list_type = 'ol';
            if ($ol_styl_hash{$styl_num}) { $list_styl = $ol_styl_hash{$styl_num}; }
          }
        }
        if ($list_cand =~ m@\\listid([0-9]+)@) {

          my $list_id = $1;
          $ltyp_hash{$list_id} = $list_type;

          if ($list_styl) {
            $list_css .= $list_type.".l".$list_id." { list-style-type: ".$list_styl." ; }\n";
          }
        }
      }
    }
  }

  my @rtrn_array  = ();

  $rtrn_array[0]  = $list_css;
  $rtrn_array[1]  = \%ovrd_hash;
  $rtrn_array[2]  = \%ltyp_hash;

  return \@rtrn_array;

}

####################################################


####################################################
####### SUBROUTINE@  REMOVE NON-SECTION GROUPS

sub removeNonSections() {

  my $fs = shift;

  my @cwArray = (

    'themedata',
    'colorschememapping',
    'fonttbl',
    'filetbl',
    'colortbl',
    'stylesheet',
    'latentstyles',
    'listtable',
    'listoverridetable',
    'rsidtbl',
    'old[cpts]props',
    'protusertbl',
    'mmathPr',
    'generator',
    'info',
  );

  #print "@cwArray";

  for my $cw (@cwArray) {

    if ($fs =~ m@(\\$cw)@) {

      #print $1."\n";

      $fs =~ s@(.*?){[^{}]*\\$cw@@s;
      my $leader = $1;
      #print $leader."\n\n";
      my $cnt = 1;
      until ($cnt == 0) {

        if    ($fs =~ s@^[^{}]*{@@) { $cnt++; }
        elsif ($fs =~ s@^[^{}]*}@@) { $cnt--; }
      }
      $fs = $leader.$fs;
    }
  }

  return $fs;

}

####################################################


####################################################
####### SUBROUTINE@  INSERT LIST HTML

sub insertListHTML {
  my $fs      = shift;
  my $l_info  = shift;

  ## split into paragraphs (RTF definition)
  $fs =~ s@(\\pard(?![a-zA-Z0-9-]))@RCRDSPRTR$1@g;

  my $doc_str_bldr    = '';
  my $p_clas_cnt      = 0;

  ## look for lists and mark up acccordingly:
  for my $seg (split(/RCRDSPRTR/, $fs)) {

    if ($seg =~ m@\\ls([0-9]+)@) {

      my $ref_num     = $1;
      my $list_id     = $l_info->[1]->{$ref_num};
      my $list_typ    = $l_info->[2]->{$list_id};
      my $cw_ptrn     = '\\listtext(?![a-zA-Z0-9-])';

      if ($seg =~ m@(.*?){[^{}]*$cw_ptrn@s) {

        my $seg_ldr = $1;
        my $str_bldr = '';

        while ($seg =~ m@{[^{}]*$cw_ptrn(?:(?!{[^{}]*$cw_ptrn).)*@sg) {

          my $list_item = $&;

          #print $list_item."\n";

          $list_item =~ s@(.*?){[^{}]*\\$cw_ptrn@@s;
          my $li_ldr = $1;
          my $cnt = 1;
          until ($cnt == 0) {
            if    ($list_item =~ s@^[^{}]*{@@) { $cnt++; }
            elsif ($list_item =~ s@^[^{}]*}@@) { $cnt--; }
            else  { last; }
          }
          $str_bldr .= $li_ldr.$list_item;

        }

        if (!$list_typ) { $list_typ = 'ul'; }
        my $list_delim = "LSTSTUBCODE_TYPE_".$list_typ."_ID_".$list_id;
        $seg = $seg_ldr."\\\n\\\n".$list_delim."\\\n".$str_bldr."LSTSTUBCODE_END\\\n\\\n";
      }

    }
    $doc_str_bldr .= $seg;
  }

  return $doc_str_bldr;
}

####################################################


####################################################
####### SUBROUTINE@  MULTI SPACE \sbN and \saN PARAGRAPHS

sub multiSpacePars() {

  my $fs = shift;

  #print $fs;
  ## split into paragraphs (RTF definition)
  $fs =~ s@(\\pard(?![a-zA-Z0-9-]))@RCRDSPRTR$1@g;

  my $doc_str_bldr    = '';
  my $last_sa         = 0;

  ## look for pars that aren't lists and mark up acccordingly:
  for my $seg (split(/RCRDSPRTR/, $fs)) {

    if ($seg !~ m@LSTSTUBCODE_TYPE_@) {

      my $sb = 0;
      my $sa = 0;

      ## multi-break pars with \sb or \sa (space before/after) >= 40
      if ($seg =~ m@\\sb([0-9]+)@) { $sb = $1; }
      if ($seg =~ m@\\sa([0-9]+)@) { $sa = $1; }
      if ($sa + $sb >= 40) {
        $seg =~ s@\\\n@\\\n\\\n@g;
        $seg =~ s@\\\n((?:.(?!\\\n))*)$@$1@s;
      }
      if ($last_sa + $sb >= 40) {
        $seg = "\\\n\\\n".$seg;
      }
      $last_sa = $sa;
    }
    $doc_str_bldr .= $seg;
  }

  return $doc_str_bldr;
}

####################################################


####################################################
####### SUBROUTINE@  INSERT PAARAGRAPH HTML - MULTI BQ MARGIN

## Helper method for insertParHTMLmultiBQ below:
sub getIndentValueAtThreshold {
  my $value = shift;
  my $value_out = 0;

  for my $pair (split(/\./, $INDENT_THRESHOLDS)) {
    my @pair_arr = split(/,/, $pair);
    if ($value >= $pair_arr[0]) {
      $value_out = $pair_arr[1];
    }
  }
  return $value_out;
}

sub insertParHTMLmultiBQ {
  my $fs = shift;
  my $bq_factor = $BLOCKQUOTE_FACTOR;

  #print $fs;
  ## split into paragraphs (RTF definition)
  $fs =~ s@(\\pard(?![a-zA-Z0-9-]))@RCRDSPRTR$1@g;

  my $doc_str_bldr    = '';
  my $p_clas_cnt      = 1;
  my %p_css_hash      = ();

  ## look for pars that aren't lists and mark up acccordingly:
  for my $seg (split(/RCRDSPRTR/, $fs)) {

    if ($seg eq '') { next; }

    if ($seg =~ m@^\\pard(?![a-zA-Z0-9-])@ && $seg !~ m@LSTSTUBCODE_TYPE_@) {

      my $p_clas_name = '';
      my $calc_right_indent = 0;

      if ($seg =~ m@\\li(-?[0-9]+)@) {
        my $spc = sprintf("%.0f", ($1 / $DIMFACTOR));
        if ($INDENT_THRESHOLDS) {
          $spc = getIndentValueAtThreshold($spc);
        }
        if ($spc) {
          $p_clas_name .= '_ml' . $spc;
          # While using $spc in this calculation may introduce rounding errors
          # if not using $INDENT_THRESHOLDS, the significance of it in context
          # does not justify the complexity of dealing with both cases.
          $calc_right_indent = sprintf("%.0f", ($spc * $bq_factor))
        }
      }
      if ($calc_right_indent) {
        # We can ignore $INDENT_THRESHOLDS in this case as that would have
        # already been dealt with in the left-indent case.
        $p_clas_name .= '_mr' . $calc_right_indent;
      } elsif ($RIGHT_INDENT && $seg =~ m@\\ri(-?[0-9]+)@) {
        my $spc = sprintf("%.0f", ($1 / $DIMFACTOR));
        if ($INDENT_THRESHOLDS) {
          $spc = getIndentValueAtThreshold($spc);
        }
        if ($spc) {
          $p_clas_name .= '_mr' . $spc;
        }
      }
      if ($FIRST_INDENT && $seg =~ m@\\fi(-?[0-9]+)@) {
        my $spc = sprintf("%.0f", ($1 / $DIMFACTOR));
        if ($spc) {
          $p_clas_name .= '_ti' . $spc;
        }
      }

      # Text alignment:
      if ($DEFAULT_TEXT_ALIGN ne 'l' && $seg =~ m@\\ql@) {
        $p_clas_name .= '_tal';
      }
      if ($DEFAULT_TEXT_ALIGN ne 'r' && $seg =~ m@\\qr@) {
        $p_clas_name .= '_tar';
      }
      if ($DEFAULT_TEXT_ALIGN ne 'j' && $seg =~ m@\\qj@) {
        $p_clas_name .= '_taj';
      }
      if ($DEFAULT_TEXT_ALIGN ne 'c' && $seg =~ m@\\qc@) {
        $p_clas_name .= '_tac';
      }
      # Left align is the default in RTF, so if we explicitly have another default
      # and there is no alignment value, we'll need to explicitly set alignment to left
      if ($DEFAULT_TEXT_ALIGN && $DEFAULT_TEXT_ALIGN ne 'l' && $seg !~ m@\\q[lrjc]@) {
        $p_clas_name .= '_tal';
      }

      my $p_open = '<p>';
      my $p_close = '</p>';
      if ($p_clas_name) {
        if (!$p_css_hash{$p_clas_name}) {
          $p_css_hash{$p_clas_name} = $p_clas_cnt;
          $__par_table_hash{$p_clas_cnt} = $p_clas_name;
          $p_clas_cnt++;
        }
        my $p_clas = $p_css_hash{$p_clas_name};
        $p_open = '<p' . $p_clas . '>';
        $p_close = '</p' . $p_clas . '>';
      }
      if ($RTF_COMPARE_OUTPUT) {
        $p_close = '';
      }
      $seg = $p_open . $seg . $p_close;
    }
    $doc_str_bldr .= $seg;
  }

  return $doc_str_bldr;
}

####################################################


####################################################
####### SUBROUTINE@  INSERT COLOR SPANS - METHOD 1

## method 1 of 2 - double pass, but easier to see what is going on.

## In RTF-eze, color "spans" are not opened and closed, but rather,
## We start with a color, and use that color until encountering
## another color control word (\cfN) which changes it.
## However, color changes can also occur when a color is explicitly
## changed inside a group, and we come to the end of the group,
## in which the color reverts back to the color we had before entering
## the group.
##
## in the first pass, we are dealing with reaching the end of a group
## in which the color was changed, and we now need to revert back to
## the color we had before we entered the group.
## So we are going to place explicit \cfn's on those places.
## The reason for this will be understood upon study of the
## second pass algorithm.
##
## So, for pass 1:
## We split out (and capture) a \cfN, a {, or a } in our file string.
## ) When we come across a \cfN, we mark it as $crnt_clr_ref
##   (current color reference).
## ) When we come across a {, we push the current color reference onto
##   the stack. (now we are inside a group)
## ) If we come across another \cfN while we are in the group, that becomes
##   our new $crnt_clr_ref.
## ) So when we come to a } (now at the end of our group), we can pop the
##   stack and compare that to the $crnt_clr_ref, and see if it has changed
##   since entering the group.  If it has changed, then we use the value we
##   popped off the stack and place an explicit \cfN just after the } stating
##   that we have returned to the color used before entering the group.
##   If it hasn't changed, we don't need to do anything.
##
## This makes it all easier for the second pass. Now we don't have to be
## concerned with brackets, we only have to look for \cfN control words and
## act accordingly.
##
## In the second pass, we are now going to place span containers around color
## sections. However, we don't want to place containers around the default color
## (\cf0). The algorithm is simple:
## ) Upon the beginning of a color section, create two variables -
##   one for the closing tag of the previous section - $prev_cls ,
##   one for the opening tag of the current section  - $crnt_opn .
##   Initialize them with an empty string.
## (We will always be keeping a record of the previous color reference
##   This is initially set to 0 (the default color reference))
## ) see if the previous color reference was !0 (not default color)
##   If not, $crnt_opn = '</span>'
## ) see if the current color reference is !0 (not default color)
##   If not, $prev_cls = '<span class="clrN">' where N = current color ref.
## ) Place $prev_cls.$crnt_opn at this point.
##
## As you can see, $prev_cls and $crnt_opn will always be empty when
## corresponding to \cf0, thus, no span containers around
## the default color


sub insertColorSpans1() {

  my $fs          = shift;
  my $str_bldr    = '';
  my @stack;
  my $crnt_clr_ref    = '';

  for my $seg (split(/(\\cf[0-9]+|{|})/, $fs)) {

    if ($seg eq '') { next; }

    if ($seg =~ m@\\cf([0-9]+)@) { $crnt_clr_ref = $1; }
    if ($seg eq '{')            { push(@stack, $crnt_clr_ref); }
    if ($seg eq '}')            {
      my $last_clr = pop(@stack);
      if ($last_clr != $crnt_clr_ref) { $seg .= '\cf'.$last_clr; }
    }
    $str_bldr .= $seg;
  }
  $fs = $str_bldr;

  my $last_clr_ref    = 0;

  $fs =~ s{(\\cf([0-9]+))}
  {
    my $code        = $1;
    my $clr_ref     = $2;
    my $prev_cls    = '';
    my $crnt_opn    = '';
    if ($last_clr_ref != 0) { $prev_cls = '</span>'; };
    if ($clr_ref      != 0) { $crnt_opn = '<span class="clr'.$clr_ref.'">'; };
    $last_clr_ref = $clr_ref;
    my $rplcmnt = $code.$prev_cls.$crnt_opn;
    $rplcmnt;
  }ge;

  return $fs;
}

####################################################


####################################################
####### SUBROUTINE@  INSERT COLOR SPANS - METHOD 2

## method 2 of 2 - single pass, more efficient

## This algorithm would be best understood by studying its
## evolutionary predecessor insertColorSpans1() above.
## In this case, however, rather than placing explicit
## \cfN control words in the file in a preliminary pass,
## we are doing it all in one pass, putting the $prev_cls.$crnt_opn
## in directly.


sub insertColorSpans2() {
  my $fs              = shift;

  my $str_bldr        = '';
  my $crnt_clr_ref    = '';
  my $last_clr_ref    = '-1';
  my @stack;

  for my $seg (split(/(\\cf[0-9]+ ?|{|})/, $fs)) {

    if ($seg eq '') { next; }

    if ($seg eq '{') { push(@stack, $crnt_clr_ref); }
    if ($seg eq '}') {
      my $last_clr = pop(@stack);
      if ($last_clr ne $crnt_clr_ref) {
        $seg .= '\cf' . $last_clr;
      }
    }
    if ($seg =~ m@\\cf([0-9]+)@) {
      $crnt_clr_ref = $1;
      if ($crnt_clr_ref ne $last_clr_ref) {
        my $prev_cls    = '';
        my $crnt_opn    = '';
        # Don't insert closing tags for rtf compare
        if (!$RTF_COMPARE_OUTPUT && $last_clr_ref ne $DEFAULT_COLOR_REFERENCE && $last_clr_ref ne '-1') {
          $prev_cls = '</clr' . $last_clr_ref . '>';
        };
        # If rtf compare, ignore default exclusion.
        if ($RTF_COMPARE_OUTPUT || $crnt_clr_ref ne $DEFAULT_COLOR_REFERENCE) {
          $crnt_opn = '<clr' . $crnt_clr_ref . '>';
        };
        $last_clr_ref = $crnt_clr_ref;
        $seg .= $prev_cls.$crnt_opn;
      }
    }
    $str_bldr .= $seg;
  }

  return $str_bldr;
}

sub insertFontSizeSpans() {
  my $fs = shift;
  my $str_bldr = '';
  my $crnt_fs_ref = '';
  my $last_fs_ref = '-1';
  my $default_fs = '0';
  my %fs_sizes;

  if (!$RTF_COMPARE_OUTPUT && $DEFAULT_FONT_SIZE) {
    $default_fs = $DEFAULT_FONT_SIZE;
  }

  my @stack;

  for my $seg (split(/(\\fs[0-9]+ ?|{|})/, $fs)) {

    if ($seg eq '') { next; }

    my $apnd_tgs = 0;

    if ($seg eq '{') { push(@stack, $crnt_fs_ref); }
    if ($seg eq '}') {
      my $last_fs = pop(@stack);
      if ($last_fs ne $crnt_fs_ref) {
        $seg .= '\fs' . $last_fs;
      }
    }
    if ($seg =~ m@\\fs([0-9]+)@ || $apnd_tgs) {
      $crnt_fs_ref = $1;
      if ($crnt_fs_ref ne $last_fs_ref) {
        my $prev_cls    = '';
        my $crnt_opn    = '';
        # Don't insert closing tags for rtf compare
        if (!$RTF_COMPARE_OUTPUT && $last_fs_ref ne '0' &&
            $last_fs_ref ne $default_fs && $last_fs_ref ne '-1') {
          $prev_cls = '</fs' . $last_fs_ref . '>';
        };
        if ($crnt_fs_ref ne '0' && $crnt_fs_ref ne $default_fs) {
          $crnt_opn = '<fs' . $crnt_fs_ref . '>';
          $fs_sizes{$crnt_fs_ref} = 1;
        };
        $last_fs_ref = $crnt_fs_ref;
        $seg .= $prev_cls.$crnt_opn;
      }
    }
    $str_bldr .= $seg;
  }

  return $str_bldr;
}

sub insertFontFamilySpans() {
  my $fs = shift;
  my $str_bldr = '';
  my $crnt_ff_ref = '';
  my $last_ff_ref = '-1';
  my $default_ff = '0';
  my %ff_sizes;

  if (!$RTF_COMPARE_OUTPUT && $DEFAULT_FONT_FAMILY) {
    $default_ff = $DEFAULT_FONT_FAMILY;
  }

  my @stack;

  for my $seg (split(/(\\f[0-9]+ ?|{|})/, $fs)) {

    if ($seg eq '') { next; }

    my $apnd_tgs = 0;

    if ($seg eq '{') { push(@stack, $crnt_ff_ref); }
    if ($seg eq '}') {
      my $last_ff = pop(@stack);
      if ($last_ff ne $crnt_ff_ref) {
        $seg .= '\f' . $last_ff;
      }
    }
    if ($seg =~ m@\\f([0-9]+)@ || $apnd_tgs) {
      $crnt_ff_ref = $1;
      if ($crnt_ff_ref ne $last_ff_ref) {
        my $prev_cls    = '';
        my $crnt_opn    = '';
        # Don't insert closing tags for rtf compare
        if (!$RTF_COMPARE_OUTPUT && $last_ff_ref ne '0' &&
            $last_ff_ref ne $default_ff && $last_ff_ref ne '-1') {
          $prev_cls = '</ff' . $last_ff_ref . '>';
        };
        if ($crnt_ff_ref ne '0' && $crnt_ff_ref ne $default_ff) {
          $crnt_opn = '<ff' . $crnt_ff_ref . '>';
          $ff_sizes{$crnt_ff_ref} = 1;
        };
        $last_ff_ref = $crnt_ff_ref;
        $seg .= $prev_cls.$crnt_opn;
      }
    }
    $str_bldr .= $seg;
  }

  return $str_bldr;
}

####################################################


####################################################
####### SUBROUTINE@  INSERT INLINE STYLE TAGS

## Turn on and off styles - we must take into account that within
## the RTF, a style within a group may be turned off by reaching the
## end of the group, without an explicit 'off' control word
##
## algorithm:
##
## ($active = state of whether the style has been turned on or off)
## ) encounter 'on' control word:
##     If $active is not set,
##       1. insert HTML start tag
##       2. set $active
##       3. set group level to 1 (no matter where it is currently)
##     If active is already set, don't do anything.
## ) encounter either a } or an 'off' control word:
##     If $active and $grp_lvl == 1
##       1. insert HTML end tag
##       2. reset $active
## ) encountering { or } increases or decreases group level accordingly


sub insertInlineTags() {

  my $fs          = shift;
  my $cw_on       = shift;    # the nominal segment of an "on" control word
  my $cw_off      = shift;    # the nominal segment of an "off" control word
  my $tag_on      = shift;    # the string used for the HTML start tag
  my $tag_off     = shift;    # the string used for the HTML end tag
  my $active      = 0;        # turn on style sets $active, turn off style resets $active
  my $grp_lvl     = 0;        # group level - the level of the group we are in
  my $str_bldr    = '';

  for my $seg (split(/(\\$cw_on-?[0-9]*(?: |(?![a-zA-Z0-9-]))|\\$cw_off(?: |(?![a-zA-Z0-9-]))|{|})/, $fs)) {

    if (!$active && $seg =~ m@^\\$cw_on ?$@) {
      $seg .= $tag_on ;
      $active++;
      $grp_lvl = 1;
    }
    if ($active && $grp_lvl == 1 && ($seg eq '}' || $seg =~ m@^\\$cw_off ?$@)) {
      $seg .= $tag_off;
      $active = 0;
    }
    if ($seg eq '{') { $grp_lvl++; }
    if ($seg eq '}') { $grp_lvl--; }

    $str_bldr .= $seg;
  }

  return $str_bldr;
}

####################################################


#################################################################
#################################################################
#################################################################
##
##      HTML PROCESSING
##


####################################################
####### @SUBROUTINE: CONVERT UTF-8 NON ASCII CHARS TO HTML NUMERICAL ENCODING

sub utf8_to_num_html () {

  ## tables of generic subs :

  my %gen_sq_table = (
    '8216' => "\x27",
    '8217' => "\x27",
    '8218' => "\x2C",
    '8219' => "\x27",
    '8220' => "\x22",
    '8221' => "\x22",
    '8223' => "\x22",
    '8242' => "\x27",
    '8243' => "\x22",
    '8245' => "\x27",
    '8246' => "\x22",
  ),

  my %gen_op_table = (
    '8208' => "-",
    '8209' => "-",
    '8210' => "--",
    '8211' => "--",
    '8212' => "--",
    '8213' => "--",
    '8214' => "||",
    '8215' => "_",
    '8224' => "+",
    '8225' => "+",
    '8226' => "*",
    '8228' => ".",
    '8229' => "..",
    '8230' => "...",
    '8231' => ".",
  ),

  my %gen_ws_table = (
    '160'  => " ",
    '8192' => " ",
    '8193' => " ",
    '8194' => " ",
    '8195' => " ",
    '8196' => " ",
    '8197' => " ",
    '8198' => " ",
    '8199' => " ",
    '8200' => " ",
    '8201' => " ",
    '8202' => " ",

    '8203' => "",
    '8204' => "",
    '8205' => "",
    '8206' => "",
    '8207' => "",

    '8239' => " ",
    '8232' => "\n",
    '8233' => "\n\n",
    '8287' => " ",
    '12288' => " ",
  ),

  ## end tables of generic subs


  my $ts = shift @_;

  $ts = decode('UTF-8', $ts);

  my $os = '';
  my $err_msg = '';

  for my $seg (split(/([\x{0080}-\x{10ffff}])/, $ts)) {

    if ($seg =~ m@([\x{0080}-\x{10ffff}])@) {

      my $u_val = ord($1);

      unless($u_val == 65533) {

        if ( $NO_SMARTQUOTES && $gen_sq_table{$u_val}) {
          $os .= $gen_sq_table{$u_val};
        }elsif ( $ASCII_ONLY_PUNCT && $gen_op_table{$u_val}) {
          $os .= $gen_op_table{$u_val};
        }elsif ($gen_ws_table{$u_val}) {
          $os .= $gen_ws_table{$u_val};
        }else{
          if ($HTMLENTITIES) {
            $os .= '&#'.$u_val.';';
          }else{
            $os .= $seg;
          }
        }

      }else{

        $ts =~ m@.{0,20}$seg.{0,20}@;
        $err_msg .= $&."\n\n";
        $os .= $seg;
      }

    }else{

      $os .= $seg;
    }
  }

  if ($err_msg) {
    print "\nWARNING : INVALID UTF-8 ENCODING IN : \n\n";
    print $outfile."\n\n";
    $err_msg = encode('UTF-8', $err_msg);
    print $err_msg;
  }

  $os = encode('UTF-8', $os);

  return $os;

}

####################################################


####################################################
####### @SUBROUTINE: ORGANIZE WHITESPACE FOR PROCESSING

sub whitespace() {

  my $ts = shift;

#print $ts;

  $ts =~ s@\r\n@\n@g;
  $ts =~ s@\r@\n@g;
  $ts =~ s@\n\n\n+@\n\n@g;

  my $tab_rplcmnt;

  if ($EXPAND_TABS) {
    $ts =~ s@\t@[TAB_SUB]@g;
    my $cnt = 0;
    while ($cnt < $EXPAND_TABS) { $tab_rplcmnt .= '&nbsp;'; $cnt++; }
  }else{
    $ts =~ s@\t@ @g;
  }

  $ts =~ s@$WS_HOR+(\n|</p>|</div>)@$1@g;
  if (!$ALLOW_LEAD_SPACES) {
    $ts =~ s@\n$WS_HOR+@\n@g;
  }else{
    while ($ts =~ s@(\n|$P_OPEN)((?:\&nbsp;)*) @$1$2&nbsp;@g) {}
  }

  if ($MULTSPACES) {
    while($ts =~ s@( +) @$1&nbsp;@g) {}
  }else{
    $ts =~ s@ +@ @g;
  }

  $ts =~ s@^@\n\n@g;
  $ts =~ s@$@\n\n@g;

  my @rtrn_array;

  $rtrn_array[0] = $tab_rplcmnt;
  $rtrn_array[1] = $ts;

  return \@rtrn_array;
}

####################################################


####################################################
####### @SUBROUTINE: SINGLE-FY BACK TO BACK CONTAINERS

sub singlfy_back_2_back_containers() {

  ## ie, <a>some text</a> <a>some more text</a> becomes
  ##
  ##     <a>some text some more text</b>

  my $ts = shift;

  my $os = '';

  for my $seg (split(/\n\n/, $ts)) {

    if ($seg eq '') { next; }

    while ($seg =~ s@<([a-z]+) ([^<>]*)>((?:(?!<\1|</\1>).)*)</\1>@<$1 $2>$3</$1 $2>@gs) {}
    while ($seg =~ s@</([a-z]+) ([^<>]*)>($NC_ALL*)<\1 \2>@$3@g) {}
    while ($seg =~ s@</([a-z]+)>($NC_ALL*)<\1>@$2@g) {}
    $seg =~ s@</([a-z]+) ([^<>]*)>@</$1>@g;
    $os .= $seg."\n\n";
  }

  return "\n\n".$os;
}


####################################################
####### @SUBROUTINE: DEFINE PARAGRAPH BLOCKS

sub defineParagraphBlocks {
  my $ts = shift;

#print $ts;

  $ts =~ s@\n\n\n+@\n\n@g;

  ####################################################
  ## Minimize containers - ie, shift out whitespace so opening tag is as far to
  ## the right as possible, and closing tag is as far to the left.
  ## This can eliminate "paragraphs" which only consist of tags, and combined
  ## with the subsequent methods will eventually eliminates "paragraphs"
  ## which are only tags and whitespace.

  while($ts =~ s@(<[a-z0-9][^>]*>)($WS_ALL+)@$2$1@g) {}
  while($ts =~ s@($WS_ALL+)(</[^>]+>)@$2$1@g) {}

  $ts =~ s@\n\n\n+@\n\n@g;

  ####################################################
  ## Remove all whitespace between newlines which would prevent a paragraph
  ## break, all trailing whitespace, and if allowed, all leading whitespace.

  while ($ts =~ s@$WS_HOR+\n@\n@g) {}
  if (!$ALLOW_LEAD_SPACES) {
    $ts =~ s@\n$WS_HOR+@\n@g;
  }

  $ts =~ s@\n\n\n+@\n\n@g;

  ####################################################
  ## Another minimize containers?
  ## TODO: Find out how much of this stuff is redundant, (though repeating in
  ## some cases may be necessary.)

  while($ts =~ s@(<[a-z0-9][^>]*>)($WS_ALL+)@$2$1@g) {}
  while($ts =~ s@($WS_ALL+)(</[^>]+>)@$2$1@g) {}

  $ts =~ s@\n\n\n+@\n\n@g;

  ####################################################
  ## Remove containers which only contain tags or whitespace.
  ## (the regex says: find a matching opening and closing element which
  ##  is separated only by tag(s) or whitespace, and eliminate the opening
  ##  and closing tag, and be non-greedy about it (or else
  ##  <a></a><a></a> will become </a><a> - not good))
  while ($ts =~ s@<([a-z0-9]+)[^<>]*>((?:<(?!\1|/\1)[^<>]*>|$WS_ALL)*)</\1>@$2@) {}

  ####################################################
  ## "POLARIZE" TAGS
  ## "polarize" tags adjacent to \n / \n\n areas:
  ## 1) bubble sort all tags around a \n  =>  closing tags on left, opening tags on right
  ## 2) pull all closing tags to left of \n...
  ## 3) pull all opening tags to right of \n
  ## More stuff to ensure elimination of "empty" paragraphs.

  while($ts =~ s@(<[a-zA-Z][^>]*>)($WS_ALL*)(</[^>]*>)@$3$2$1@sg) {}
  while($ts =~ s@(\n$WS_ALL*)(</[^>]*>)@$2$1@sg) {}
  while($ts =~ s@(<[a-zA-Z][^>]*>)($WS_ALL*\n)@$2$1@sg) {}

  ####################################################
  ## Preceding actions may have left whitespace between newlines, so once
  ## again, remove it.

  while ($ts =~ s@$WS_HOR+\n@\n@g) {}
  if (!$ALLOW_LEAD_SPACES) {
    $ts =~ s@\n$WS_HOR+@\n@g;
  }

  $ts =~ s@\n\n\n+@\n\n@g;

  my @segs = split(/\n\n/, $ts);
  $ts = '';

  for my $seg (@segs) {
    if ($seg =~ m@\n@) {
      my @subsegs = split(/\n/, $seg);
      my $len = scalar @subsegs;
      my $lastindex = $len - 1;
      $seg = '';

      for (my $i = 0; $i < $len; $i++) {
        my $subseg = $subsegs[$i];
        if ($i == 0) {
          $seg .= '__RTF2HTML_SINGLE_RETURN_PARAGRAPH_FLAG_FIRST__' . $subseg . "\n\n";
        } elsif ($i == $lastindex) {
          $seg .= '__RTF2HTML_SINGLE_RETURN_PARAGRAPH_FLAG_LAST__' . $subseg;
        } else {
          $seg .= '__RTF2HTML_SINGLE_RETURN_PARAGRAPH_FLAG_MID__' . $subseg . "\n\n";
        }
      }
    }
    $ts .= $seg . "\n\n";
  }

  $ts =~ s@\n\n\n+@\n\n@g;

  ####################################################
  ## Apply cross-block containers.  RTF just turns stuff on and turns it off or
  ## changes it in serial fashion.  It does not care about "containers" or
  ## nesting.  The tags that exist at this point only reflect that, and will
  ## need to be treated so they end up being nested properly.
  ##
  ## This is the first step toward nesting containers in HTML.  We keep track
  ## of when styling is turned on, and ensure that if it doesn't get turned
  ## back off before the end of a paragraph, we'll close the container within
  ## that paragraph then re-open it again in the next.  We continue to do this
  ## until we finally hit a closing tag.
  ##
  ## At this point, with all the previous treatment, we can define our
  ## paragraphs to be blocks of text separated by two newlines.
  ## (Note that rtf2html does not have treatment for RTF users who create
  ## paragraphs by setting their spacing attribute to 2.0 (double) and just
  ## use single returns.)

  my $os = '';
  my %open_containers;

my $switch = 0;

  for my $seg (split(/\n\n/, $ts)) {
    $seg =~ s@^\s+|\s+$@@g;

if ($seg =~ m@they go on contriving the mischief of their hearts, opening their shameless mouths@) {
  $switch = 1;
}
if ($seg !~ m@they go on contriving the mischief of their hearts, opening their shameless mouths@) {
  $switch = 0;
}
$switch = 0;

if ($switch) {
  print "============================================\n";
  print "input string:\n\n";
  print $seg . "\n\n";
}

    # Prepend the paragraph with any open containers.
    for my $tagname (keys %open_containers) {
      $seg = $open_containers{$tagname} . $seg;
    }

    $seg =~ s@<@SPRTRSPRTR<@g;
    $seg =~ s@>@>SPRTRSPRTR@g;

    my @seg2s = split(/SPRTRSPRTR/, $seg);
    my $pstrbuilder = '';
    my @tags;
    my @accum = ();

    for (my $i = 0; $i < scalar @seg2s; $i++) {
      my $seg2 = $seg2s[$i];
      if ($seg2 eq '') {
        next;
      }

      if ($seg2 =~ m/^</) {
        my $tagname = $seg2;
        $tagname =~ s@</?([a-z0-9]+).*@$1@;
        my $tagtype = 1;
        if ($seg2 =~ m@^</@) {
          $tagtype = -1;
          # If we found a closing tag for an open container, delete
          # from open containers.
          if ($open_containers{$tagname}) {
            delete $open_containers{$tagname};
          }
        }
        # "tag" property just aids in debugging.
        push(@tags, { tagname => $tagname,
                      tagtype => $tagtype,
                      tag => $seg2 });
      }
      $pstrbuilder .= $seg2;
    }

    # If a closer wasn't found for an opener, enter it into %open_containers.
    for (my $i = 0; $i < scalar @tags; $i++) {
      if ($tags[$i]{tagtype} != 1) {
        next;
      }
      my $tagname = $tags[$i]{tagname};
      my $found = 0;
      my %foundindexes;
      for (my $j = $i + 1; $j < scalar @tags; $j++) {
        if ($tags[$j]{tagname} eq $tagname && $tags[$j]{tagtype} == -1 && !$foundindexes{$j}) {
          $found = 1;
          $foundindexes{$j} = 1;
        }
      }
      if (!$found) {
        $open_containers{$tagname} = '<' . $tagname. '>';
      }
    }

    for my $tagname (keys %open_containers) {
      $pstrbuilder .= '</' . $tagname . '>';
    }

if ($switch) {
  print "============================================\n";
  print "after distribute cross block containers:\n\n";
  print $pstrbuilder . "\n\n";
}

    # Check to see if we have more than one <p> containers.
    my $openps = 0;
    my $closeps = 0;
    my $teststr = $pstrbuilder;
    while ($teststr =~ s@<p@@) {
      $openps++;
    }
    while ($teststr =~ s@</p@@) {
      $closeps--;
    }

    # Print error if there is a open/close tag mismatch.
    if ($openps != 1 or $closeps != -1) {
      print "ERROR: p tag mismatch\n";
    }

    ## Treatment for blocks with multiple paragraphs (according to RTF
    ## definition.)  Some blocks may contain multiple RTF paragraph definitions.
    ## This is usually due to the user doing something like, single return
    ## for a new line, but this line my be aligned differently, or have a
    ## different indent.  With Bean, new definitions seem to always occur at
    ## a line break.
    ##
    ## Our approach here is to separate lines with new different paragraph
    ## definitions by a double newline.  We have already done that with the
    ## input string, and flagged those strings with unique text, which will
    ## now be processed.
    ##
    ## However, we don't want to ignore the fact that these were only single
    ## spaced, and the consumer will probably want them to appear single
    ## spaced in the browser as well.  So we add flags to the <p> tags which
    ## will signal post processing methods to insert classes based on the
    ## whether the container is the first, the last, or if 3 or more instances,
    ## a middle one.

    if ($pstrbuilder =~ s@__RTF2HTML_SINGLE_RETURN_PARAGRAPH_FLAG_FIRST__@@) {
      $pstrbuilder =~ s@(<p[0-9]*)>@$1 sbpf>@;
    } elsif ($pstrbuilder =~ s@__RTF2HTML_SINGLE_RETURN_PARAGRAPH_FLAG_MID__@@) {
      $pstrbuilder =~ s@(<p[0-9]*)>@$1 sbpm>@;
    } elsif ($pstrbuilder =~ s@__RTF2HTML_SINGLE_RETURN_PARAGRAPH_FLAG_LAST__@@) {
      $pstrbuilder =~ s@(<p[0-9]*)>@$1 sbpl>@;
    }

    $os .= $pstrbuilder . "\n\n";
  }

  $os =~ s@\n\n\n+@\n\n@g;

  # All double-spaced blocks by now should only have one <p> container.
  # Pull the opening and closing p's to the outside of the block.
  # Of course we want them there, but this should also cause subsequent
  # functions from messing with them and they should remain in position.
  @segs = split(/\n\n/, $os);
  $os = '';

  for my $seg (@segs) {
    $seg =~ s@(.*)(<p[^>]*>)(.*)@$2$1$3@;
    $seg =~ s@(.*)(</p[^>]*>)(.*)@$1$3$2@;
    $os .= $seg . "\n\n";
  }

  return $os;
}

####################################################
####### @SUBROUTINE: FORMAT HTML

sub formatHTML () {

  my $ts = shift;

  ## MAIN TESTERS
  #$ts = '<a> text1 <b> text2 <c><d> text3 </a> text4 </c> text5 <e> text6 <f></b> text7 <g></e> text8 </d> text9 </f> text10 </g>' . "\n\n";

  #$ts = '<a> text <b> text <c> text <d></c><e></a><f></b> text </d> text <g> text </e> text </f> text <h> text <i> text </g><j></h><k></i> text </j> text </k>';

#print $ts . "\n\n";

  #$ts = '<p>okay<em>then<clr2>now<fs18>President Barack Obama, February 13, 2015</em></fs18></clr2></p>';
  #$ts = '<p><em><clr2><fs18>President Barack Obama, February 13, 2015</em></fs18></clr2></p>';


  my $os = '';

my $switch = 0;

  for my $seg (split(/\n\n/, $ts)) {
    $seg =~ s@^\s+|\s+$@@g;


if ($seg =~ m@No one in the United States|because of who they are|President Barack Obama@) {
  $switch = 1;
}
if ($seg !~ m@No one in the United States|because of who they are|President Barack Obama@) {
  $switch = 0;
}
$switch = 0;

if ($switch) {
  print "============================================\n";
  print "input string:\n\n";
  print $seg . "\n\n";
}

    ######################################################################
    ## Remove all containers that are broken by only whitespace.
    ## We'll mask <p>'s however as this creates some undesired effects.

    $seg =~ s@<p@<P@g;
    $seg =~ s@</p@</P@g;
    while ($seg =~ s@</([a-z0-9]+)>($WS_HOR*)<\1>@$2@) {}
    while ($seg =~ s@</([a-z0-9]+)>($WS_HOR*)<\1 [^>]+>@$2@) {}
    $seg =~ s@<P@<p@g;
    $seg =~ s@</P@</p@g;

if ($switch) {
  print " --------------------------\n";
  print "after remove containers broken by whitespace:\n\n";
  print $seg . "\n\n";
}

    ######################################################################
    ## Generate tags list, and while we're at it, remove containers that open
    ## and close within the same tag cluster.
    ## TODO: would it be better to simplify this and do the remove containers
    ## thing with a regex?

    $seg =~ s@<@SPRTRSPRTR<@g;
    $seg =~ s@>@>SPRTRSPRTR@g;
    my @seg2s = split(/SPRTRSPRTR/, $seg);
    my @tags = ();
    my @accumtags = ();
    my @accumws = ();
    my $pstrbuilder = '';

    my $len = scalar @seg2s;
    my $lasti = $len - 1;

    for (my $i = 0; $i < $len; $i++) {
      my $seg2 = $seg2s[$i];
      if ($seg2 eq '') {
        next;
      }

      my $istext = 0;

      if ($seg2 =~ m/^</) {
        my $tagname = $seg2;
        $tagname =~ s@</?([a-z0-9]+).*@$1@;
        my $tagtype = 1;
        if ($seg2 =~ m@^</@) {
          $tagtype = -1;
        }
        push(@accumtags, { tagname => $tagname,
                           tag => $seg2,
                           type => $tagtype });

      } else {
        $istext = 1;
      }
      if ($istext or $i == $lasti) {
        if ($seg2 =~ m@^\s+$@) {
          push(@accumws, $seg2);
          next;
        }
        # Must be text - empty @accumtags into @tags/$pstrbuilder, then the text
        # into $pstrbuilder, reset @accumtags.

        my $accumtagslen = scalar @accumtags;
        for (my $j = 0; $j < $accumtagslen; $j++) {
          my $type = $accumtags[$j]{type};
          my $tagname = $accumtags[$j]{tagname};
          # "Remove" containers that exist within the same tag cluster.
          my $removecontainer = 0;
          if ($type == 1) {
            for (my $k = $j + 1; $k < $accumtagslen; $k++) {
              if ($accumtags[$k]{tagname} eq $tagname && $accumtags[$k]{type} == -1) {
                $accumtags[$k]{removetag} = 1;
                $removecontainer = 1;
              }
            }
          }
          if (!$removecontainer && !$accumtags[$j]{removetag}) {
            push(@tags, { tagname => $accumtags[$j]{tagname},
                          tag => $accumtags[$j]{tag},
                          type => $accumtags[$j]{type} });
            $pstrbuilder .= $accumtags[$j]{tag};
          }
        }
        for my $text (@accumws) {
          $pstrbuilder .= $text;
        }
        # May be here because the last seg was a tag, so test for istext.
        if ($istext) {
          $pstrbuilder .= $seg2;
        }
        @accumtags = ();
        @accumws = ();
      }
    }

if ($switch) {
  print " --------------------------\n";
  print "after remove containers that only exist within the same cluster:\n\n";
  print $pstrbuilder . "\n\n";
}

    ######################################################################
    ## Shift whitespace embedded in clusters to rhs.
    ## If whitespace occurs within a cluster of tags, move it to the rhs.
    ## This way any containers styling only whitespace will become empty,
    ## and will get removed in later process.

    $pstrbuilder =~ s@<@SPRTRSPRTR<@g;
    $pstrbuilder =~ s@>@>SPRTRSPRTR@g;
    @seg2s = split(/SPRTRSPRTR/, $pstrbuilder);
    $pstrbuilder = '';
    my $tagcount = 0;
    my @tags2 = ();
    my @accum = ();
    $len = scalar @seg2s;
    $lasti = $len - 1;

    for (my $i = 0; $i < $len; $i++) {
      my $seg2 = $seg2s[$i];
      if ($seg2 eq '') {
        next;
      }
      my $istext = 0;

      if ($seg2 =~ m/^</) {
        my $tagname = $tags[$tagcount]{tagname};
        my $tagtype = $tags[$tagcount]{type};
        my $tag = $tags[$tagcount]{tag};
        if ($tagtype == 1) {
          push(@accum, { tagname => $tagname,
                         tag => $tag,
                         type => $tagtype });
        } else {
          # Push/build closing tags immediately, this way they will come before
          # opening tags in the cluster.
          push(@tags2, { tagname => $tagname,
                         tag => $tag,
                         type => $tagtype });
          $pstrbuilder .= $tag;
        }
        $tagcount++;
      } else {
        $istext = 1;
      }
      if ($istext or $i == $lasti) {
        # Must be text - empty @accum into @tags/$pstrbuilder, then the text
        # into $pstrbuilder, reset @accum.

        for my $tabopen (@accum) {
          my $tagtype = @{$tabopen}{type};
          my $tagname = @{$tabopen}{tagname};
          my $tag = @{$tabopen}{tag};
          push(@tags2, { tagname => $tagname,
                         tag => $tag,
                         type => $tagtype });
          $pstrbuilder .= $tag;
        }
        # May be here because the last seg was a tag, so test for istext.
        if ($istext) {
          $pstrbuilder .= $seg2;
        }
        @accum = ();
      }
    }
    @tags = @tags2;

if ($switch) {
  print " --------------------------\n";
  print "after shift whitespace embedded in clusters to rhs:\n\n";
  print $pstrbuilder . "\n\n";
}
    ######################################################################
    ## "Polarize" tags within clusters, ie, in a cluster of tags, move all
    ## the closing tabs to the lhs, opening tags to the rhs.

    $pstrbuilder =~ s@<@SPRTRSPRTR<@g;
    $pstrbuilder =~ s@>@>SPRTRSPRTR@g;
    @seg2s = split(/SPRTRSPRTR/, $pstrbuilder);
    $pstrbuilder = '';
    my $tagcount = 0;
    my $clustercount = 0;
    my @tags2 = ();
    my @accumopen = ();
    my @accumclose = ();
    $len = scalar @seg2s;
    $lasti = $len - 1;

    for (my $i = 0; $i < $len; $i++) {
      my $seg2 = $seg2s[$i];
      if ($seg2 eq '') {
        next;
      }
      my $istext = 0;

      if ($seg2 =~ m/^</) {
        my $tagname = $tags[$tagcount]{tagname};
        my $tagtype = $tags[$tagcount]{type};
        my $tag = $tags[$tagcount]{tag};
        if ($tagtype == 1) {
          push(@accumopen, { tagname => $tagname,
                             tag => $tag,
                             type => $tagtype });
        } else {
          # Push/build closing tags immediately, this way they will come before
          # opening tags in the cluster.
          push(@accumclose, { tagname => $tagname,
                              tag => $tag,
                              type => $tagtype });
        }
        $tagcount++;
      } else {
        $istext = 1;
      }
      if ($istext or $i == $lasti) {
        # Must be text - empty accumulators into @tags/$pstrbuilder, then the text
        # into $pstrbuilder, reset @accum.

        for my $tabclose (@accumclose) {
          my $tagtype = @{$tabclose}{type};
          my $tagname = @{$tabclose}{tagname};
          my $tag = @{$tabclose}{tag};
          push(@tags2, { tagname => $tagname,
                         tag => $tag,
                         type => $tagtype,
                         cluster => $clustercount });
          $pstrbuilder .= $tag;
        }
        for my $tabopen (@accumopen) {
          my $tagtype = @{$tabopen}{type};
          my $tagname = @{$tabopen}{tagname};
          my $tag = @{$tabopen}{tag};
          push(@tags2, { tagname => $tagname,
                         tag => $tag,
                         type => $tagtype,
                         cluster => $clustercount });
          $pstrbuilder .= $tag;
        }
        # May be here because the last seg was a tag, so test for istext.
        if ($istext) {
          $pstrbuilder .= $seg2;
        }
        @accumopen = ();
        @accumclose = ();
        $clustercount++;
      }
    }
    @tags = @tags2;

if ($switch) {
  print " --------------------------\n";
  print "after polarize tags within clusters:\n\n";
  print $pstrbuilder . "\n\n";
}

    ##########################################################
    ## Determine the position of tags, both where they occur within the block,
    ## and relative positions to each other.  See further comments below.
    ## (This has no effect on current document string.)

    $len = scalar @tags;
    for (my $i = 0; $i < $len; $i++) {
      if ($tags[$i]{type} == 1) {
        my $tagname = $tags[$i]{tagname};
        my $tag = $tags[$i]{tag};
        my $level = 1;
        for (my $j = ($i + 1); $j < $len; $j++) {
          if ($tags[$j]{tagname} eq $tagname) {
            my $type = $tags[$j]{type};
            $level = $level + $type;
            if ($level == 0) {
              # If opening and or closing tags are in a cluster of tags, make
              # openpos and closepos to be the outermost boundaries of the cluster(s):
              #
              # eg, for: <a><b><c>text</b></a></c>, for cluster "b", make the
              # open position at <a> (farthest to the left within the cluster),
              # and close position at </c> (farthest to the right.)
              #
              # Otherwise a construct like this:
              # <a> text1 <b> text2 </a></b> text3
              # will end up like this:
              # <a> text1 </a><b><a> text2 </a></b> text3
              # which is correct, but not optimal, as the nesting problem can
              # be dealt with just by swapping tags in the cluster.  We will do
              # this at the end of the process.

              my $openpos = $i;
              my $closepos = $j;
              my $docpos = $tags[$j]{docpos};

              # Walk backward from the actual closing tag and query the docpos
              # of each tag, to see if it is adjacent to the previous one within
              # the document (and thus, within the same cluster.)
              for (my $k = ($j + 1); $k < $len; $k++) {
                $docpos++;
                if ($tags[$k]{docpos} == $docpos) {
                  $closepos++;
                } else {
                  # We found a gap in docpos, that means our last closepos
                  # was the last tag in the cluster.
                  last;
                }
              }
              $docpos = $tags[$i]{docpos};
              for (my $k = ($i - 1); $k > -1; $k--) {
                $docpos--;
                if ($tags[$k]{docpos} == $docpos) {
                  $openpos--;
                } else {
                  # We found a gap in docpos, that means our last openpos
                  # was the last tag in the cluster.
                  last;
                }
              }
              $tags[$i]{openpos} = $openpos;
              $tags[$i]{closepos} = $closepos;
              $tags[$j]{openpos} = $openpos;
              $tags[$j]{closepos} = $closepos;

              last;
            }
          }
        }
      }
    }

    ##########################################################
    ## Reorder tags within clusters to reduce container overlap as much
    ## as possible before adding new tags for this purpose.  In some cases
    ## this will reduce the amount of tags we'll have to add.

    if (1) {

      @tags2 = @tags;
      my @cluster = ();

      my $len = scalar @tags;
      my $lasttagsindex = $len - 1;

      for (my $i = 0; $i < $len; $i++) {
        if ($tags[$i]{type} == 1) {
          push(@cluster, { openpos => $tags[$i]{openpos}, closepos => $tags[$i]{closepos} });
        } else {
          $tags2[$i] = { tag => $tags[$i]{tag},
                         tagname => $tags[$i]{tagname},
                         type => $tags[$i]{type},
                         tag => $tags[$i]{tag},
                         openpos => $tags[$i]{openpos},
                         closepos => $tags[$i]{closepos},
                         cluster => $tags[$i]{cluster} };
        }
        if ($i == $lasttagsindex or $tags[$i + 1]{cluster} > $tags[$i]{cluster}) {
          my @sorted = sort { $b->{closepos} <=> $a->{closepos} } @cluster;
          my $tagpos = $cluster[0]{openpos};
          for my $mdata (@sorted) {
            my $index = @{$mdata}{openpos};
            $tags2[$tagpos] = { tag => $tags[$index]{tag},
                           tagname => $tags[$index]{tagname},
                           type => $tags[$index]{type},
                           tag => $tags[$index]{tag},
                           openpos => $tagpos,
                           closepos => $tags[$index]{closepos},
                           cluster => $tags[$i]{cluster} };
            $tags2[$tags[$index]{closepos}]{openpos} = $tagpos;
            $tagpos++;
          }
          @cluster = ();
        }
      }
      @tags = @tags2;

if ($switch) {
  print " --------------------------\n";
  print "after reorder opening tags:\n\n";
  #print Dumper(\@tags);
}

      @tags2 = @tags;
      @cluster = ();

      $len = scalar @tags;
      $lasttagsindex = $len - 1;

      my %openers_hash;

      for (my $i = 0; $i < $len; $i++) {
        if ($tags[$i]{type} == -1) {
          push(@cluster, { openpos => $tags[$i]{openpos}, closepos => $tags[$i]{closepos} });
        } else {
          @tags2[$i] = { tag => $tags[$i]{tag},
                         tagname => $tags[$i]{tagname},
                         type => $tags[$i]{type},
                         tag => $tags[$i]{tag},
                         openpos => $tags[$i]{openpos},
                         closepos => $tags[$i]{closepos},
                         cluster => $tags[$i]{cluster} };
        }

        if ($i == $lasttagsindex or $tags[$i + 1]{cluster} > $tags[$i]{cluster}) {
          if (scalar @cluster) {
            my @sorted = sort { $b->{openpos} <=> $a->{openpos} } @cluster;
            my $tagpos = $cluster[0]{closepos};
            for my $mdata (@sorted) {
              my $index = @{$mdata}{closepos};
              @tags2[$tagpos] = { tag => $tags[$index]{tag},
                             tagname => $tags[$index]{tagname},
                             type => $tags[$index]{type},
                             tag => $tags[$index]{tag},
                             openpos => $tags[$index]{openpos},
                             closepos => $tagpos,
                             cluster => $tags[$index]{cluster} };
              my $openerindex = $tags[$index]{openpos};
              $openers_hash{$openerindex} = $tagpos;
              $tags2[$openerindex]{closepos} = $tagpos;
              $openers_hash{$openerindex} = $tagpos;
              $tagpos++;
            }
          }
          @cluster = ();
        }
      }

      @tags = @tags2;

if ($switch) {
  print " --------------------------\n";
  print "after reorder closing tags:\n\n";
  #print Dumper(\@tags);
}

      ## Rebuild the doc string with the reordered tag objects.

      my $arrstr = $pstrbuilder;
      $arrstr =~ s@<@SPRTRSPRTR<@g;
      $arrstr =~ s@>@>SPRTRSPRTR@g;
      @seg2s = split(/SPRTRSPRTR/, $arrstr);

      $pstrbuilder = '';
      my $tagcount = 0;

      for (my $i = 0; $i < scalar @seg2s; $i++) {
        my $seg2 = $seg2s[$i];
        if ($seg2 eq '') {
          next;
        }
        if ($seg2 =~ m/^</) {
          $pstrbuilder .= $tags[$tagcount]{tag};
          $tagcount++;
        } else {
          $pstrbuilder .= $seg2;
        }
      }
    }

if ($switch) {
  print " --------------------------\n";
  print "after reorder tags:\n\n";
  print $pstrbuilder . "\n\n";
}

    ##########################################################
    ## Minimize containers (shift whitespace on rhs of opening tag to lhs,
    ## shift whitespace on lhs of closing tag to rhs.)

    while($pstrbuilder =~ s@(<[^/][^>]*>)($WS_ALL+)@$2$1@g) {}
    while($pstrbuilder =~ s@($WS_ALL+)(</[^>]+>)@$2$1@g) {}

if ($switch) {
  print " --------------------------\n";
  print "after minimize containers:\n\n";
  print $pstrbuilder . "\n\n";
}

    ##########################################################
    ## Insert close/open tags in order to nest properly, based on the openpos
    ## and closepos findings recorded in the $tags data structure.
    ## The point of this is for when containers overlap, one of the containers
    ## must be broken into two (or more) to maintain proper nesting.
    ##
    ## We will actually insert them into the document string, then generate a
    ## new tags list later.

    $pstrbuilder =~ s@<@SPRTRSPRTR<@g;
    $pstrbuilder =~ s@>@>SPRTRSPRTR@g;
    my @seg3s = split(/SPRTRSPRTR/, $pstrbuilder);
    my $mstrbuilder = '';
    my $count = 0;
    my %founds;

    $len = scalar @seg3s;
    for (my $i = 0; $i < $len; $i++) {
      my $seg2 = $seg3s[$i];
      if ($seg2 eq '') {
        next;
      }

      if ($seg2 =~ m/^</) {
        if ($tags[$count]{type} == 1) {
          my $closepos = $tags[$count]{closepos};
          my $tag = $tags[$count]{tag};

          for (my $j = ($count - 1); $j > -1; $j--) {
            if ($tags[$j]{type} == 1 && $founds{$j} != 1) {
              # Determine if containers overlap.  If so, break the previous
              # container at the subject tag.
              if ($tags[$j]{closepos} < $tags[$count]{closepos} && $tags[$j]{closepos} > $count) {
                $seg2 = '</' . $tags[$j]{tagname} . '>' . $seg2 . $tags[$j]{tag};
                $founds{$j} = 1;
              }
            }
          }
          $founds{$count} = 1;
        }
        $count++;
      } else {
        %founds = ();
      }
      $mstrbuilder .= $seg2;
    }

if ($switch) {
  print " --------------------------\n";
  print "after major nesting by inserting tags if needed:\n\n";
  print $mstrbuilder . "\n\n";
}

    ##########################################################
    ## Now that we have all of our tags inserted, we go through it all again
    ## (sorta), generating a new tags list which will reflect the updated doc.
    ## (No change in the doc string.)

    $mstrbuilder =~ s@<@SPRTRSPRTR<@g;
    $mstrbuilder =~ s@>@>SPRTRSPRTR@g;
    @seg2s = split(/SPRTRSPRTR/, $mstrbuilder);
    @tags = ();
    my $docposcnt = 0;

    for (my $i = 0; $i < scalar @seg2s; $i++) {
      my $seg2 = $seg2s[$i];
      if ($seg2 eq '') {
        next;
      }

      if ($seg2 =~ m/^</) {
        my $tagtype = 1;
        if ($seg2 =~ m@^</@) {
          $tagtype = -1;
        }
        my $tagname = $seg2;
        $tagname =~ s@</?([a-z0-9]+).*@$1@;

        # Unlike before, we don't need to "polarize" here because this will get
        # handled later when we sort the clusters.
        push(@tags, { tagname => $tagname,
                      tag => $seg2,
                      type => $tagtype,
                      docpos => $docposcnt });
        $pstrbuilder .= $seg2;
      } else {
        $pstrbuilder .= $seg2;
      }
      $docposcnt++;
    }

    ##########################################################
    ## (Again) Determine open and close position of containers
    ## (Now that we have inserted containers, everything has changed.)
    ## (No change in the doc string.)

    $len = scalar @tags;
    for (my $i = 0; $i < $len; $i++) {
      if ($tags[$i]{type} == 1) {
        my $tagname = $tags[$i]{tagname};
        my $tag = $tags[$i]{tag};
        my $level = 1;
        for (my $j = ($i + 1); $j < $len; $j++) {
          if ($tags[$j]{tagname} eq $tagname) {
            my $type = $tags[$j]{type};
            $level = $level + $type;
            if ($level == 0) {

              # This round, we want the actual openpos and closepos of the tag,
              # as we will be using it to re-arrange tags within clusters if
              # they are not nested properly.
              my $openpos = $i;
              my $closepos = $j;
              $tags[$i]{openpos} = $openpos;
              $tags[$i]{closepos} = $closepos;
              $tags[$j]{openpos} = $openpos;
              $tags[$j]{closepos} = $closepos;

              last;
            }
          }
        }
      }
    }

    ##########################################################
    ## (Again) Remove containers which only contain other tags.
    ## (No change in the doc string.)

    my @accum = ();
    my $tagcount = 0;
    $len = scalar @seg2s;
    $lasti = $len - 1;

    for (my $i = 0; $i < $len; $i++) {
      my $seg2 = $seg2s[$i];
      if ($seg2 eq '') {
        next;
      }

      my $istext = 0;
      my $tagislast = 0;

      if ($seg2 =~ m/^</) {
        push(@accum, $tagcount);
        $tagcount++;

        # If last segment is a tag, final tags won't process nor print because
        # we never hit a text segment where we dump it.  This will force that.
        # Note that we won't print $seg2 in this case.
        if ($i == $lasti) {
          $tagislast = 1;
        }
      } else {
        $istext = 1;
      }
      if ($istext or $tagislast) {
        my @openaccum = ();
        my @closeaccum = ();
        for (my $j = 0; $j < scalar @accum; $j++) {
          my $tagcounta = $accum[$j];
          if ($tags[$tagcounta]{type} == 1) {
            for (my $k = $j + 1; $k < scalar @accum; $k++) {
              my $tagcountb = $accum[$k];
              # Again, "remove" containers that exist within the same tag cluster.
              if ($tags[$tagcountb]{type} == -1 &&
                  $tags[$tagcounta]{tagname} eq $tags[$tagcountb]{tagname}) {
                $tags[$tagcounta]{removetag} = 1;
                $tags[$tagcountb]{removetag} = 1;
              }
            }
          }
        }
        @accum = ();
      }
    }

    ##########################################################
    ## We will be rebuilding the tags list/doc string while applying two operations:
    ## Exclude tags flagged for removal.
    ## "Polarized" closing and opening tags.

    my $ostrbuilder = '';
    my @tags2 = ();

    $len = scalar @seg2s;
    my @openaccum = ();
    my @closeaccum = ();
    my $tagcount = 0;
    my $lasti = $len - 1;

    for (my $i = 0; $i < $len; $i++) {
      my $seg2 = $seg2s[$i];
      if ($seg2 eq '') {
        next;
      }

      my $istext = 0;
      my $tagislast = 0;

      if ($seg2 =~ m/^</) {
        my $tagtype = $tags[$tagcount]{type};
        my $tagname = $tags[$tagcount]{tagname};
        my $docposcnt = $tags[$tagcount]{docpos};
        my $openpos = $tags[$tagcount]{openpos};
        my $closepos = $tags[$tagcount]{closepos};
        if ($tagtype == 1) {
          if (!$tags[$tagcount]{removetag}) {
            push(@openaccum, { tagname => $tagname,
                               tag => $seg2,
                               type => $tagtype,
                               docpos => $docposcnt,
                               openpos => $openpos,
                               closepos => $closepos
                               });
          }
        } else {
          if (!$tags[$tagcount]{removetag}) {
            push(@closeaccum, { tagname => $tagname,
                                tag => $seg2,
                                type => $tagtype,
                                docpos => $docposcnt,
                                openpos => $openpos,
                                closepos => $closepos
                                });
          }
        }

        $tagcount++;

      } else {
        $istext = 1;
      }
      if ($istext or $i == $lasti) {
        for (my $j = 0; $j < scalar @closeaccum; $j++) {
          my $tagname = $closeaccum[$j]{tagname};
          my $tag = $closeaccum[$j]{tag};
          my $docposcnt = $closeaccum[$j]{docpos};
          my $openpos = $closeaccum[$j]{openpos};
          my $closepos = $closeaccum[$j]{closepos};
          push(@tags2, { tagname => $tagname,
                        tag => $tag,
                        type => -1,
                        docpos => $docposcnt,
                        openpos => $openpos,
                        closepos => $closepos
                        });
          $ostrbuilder .= $tag;
        }
        for (my $j = 0; $j < scalar @openaccum; $j++) {
          my $tagname = $openaccum[$j]{tagname};
          my $tag = $openaccum[$j]{tag};
          my $docposcnt = $openaccum[$j]{docpos};
          my $openpos = $openaccum[$j]{openpos};
          my $closepos = $openaccum[$j]{closepos};
          push(@tags2, { tagname => $tagname,
                        tag => $tag,
                        type => 1,
                        docpos => $docposcnt,
                        openpos => $openpos,
                        closepos => $closepos
                        });
          $ostrbuilder .= $tag;
        }
        if ($istext) {
          $ostrbuilder .= $seg2;
        }
        @openaccum = ();
        @closeaccum = ();
      }
    }
    @tags = @tags2;

if ($switch) {
  print " --------------------------\n";
  print "after rebuild tags list, removing and polarizing:\n\n";
  print $ostrbuilder . "\n\n";
  #print Dumper(\@tags);
}

    ##########################################################
    ## Again: Reorder tags within clusters to reduce container overlap as much
    ## as possible before adding new tags for this purpose.  Since we have added
    ## tags, clusters may need reordering.

    if (1) {

      ## Set a cluster indexes.

      my $arrstr = $ostrbuilder;
      $arrstr =~ s@<@SPRTRSPRTR<@g;
      $arrstr =~ s@>@>SPRTRSPRTR@g;
      @seg2s = split(/SPRTRSPRTR/, $arrstr);

      my $tagcount = 0;
      my $clustercount = 0;

      for (my $i = 0; $i < scalar @seg2s; $i++) {
        my $seg2 = $seg2s[$i];
        if ($seg2 eq '') {
          next;
        }
        if ($seg2 =~ m/^</) {
          $tags[$tagcount]{cluster} = $clustercount;
          $tagcount++;
        } else {
          $clustercount++;
        }
      }

      ## Reorder opening tags within clusters.

      @tags2 = @tags;
      my @cluster = ();

      my $len = scalar @tags;
      my $lasttagsindex = $len - 1;

      for (my $i = 0; $i < $len; $i++) {
        if ($tags[$i]{type} == 1) {
          push(@cluster, { openpos => $tags[$i]{openpos}, closepos => $tags[$i]{closepos} });
        } else {
          $tags2[$i] = { tag => $tags[$i]{tag},
                         tagname => $tags[$i]{tagname},
                         type => $tags[$i]{type},
                         tag => $tags[$i]{tag},
                         openpos => $tags[$i]{openpos},
                         closepos => $tags[$i]{closepos},
                         cluster => $tags[$i]{cluster} };
        }
        if ($i == $lasttagsindex or $tags[$i + 1]{cluster} > $tags[$i]{cluster}) {
          my @sorted = sort { $b->{closepos} <=> $a->{closepos} } @cluster;
          my $tagpos = $cluster[0]{openpos};
          for my $mdata (@sorted) {
            my $index = @{$mdata}{openpos};
            $tags2[$tagpos] = { tag => $tags[$index]{tag},
                           tagname => $tags[$index]{tagname},
                           type => $tags[$index]{type},
                           tag => $tags[$index]{tag},
                           openpos => $tagpos,
                           closepos => $tags[$index]{closepos},
                           cluster => $tags[$i]{cluster} };
            $tags2[$tags[$index]{closepos}]{openpos} = $tagpos;
            $tagpos++;
          }
          @cluster = ();
        }
      }
      @tags = @tags2;

if ($switch) {
  print " --------------------------\n";
  print "after reorder opening tags 2:\n\n";
  #print Dumper(\@tags);
}

      ## Reorder closing tags within clusters

      @tags2 = @tags;
      @cluster = ();

      $len = scalar @tags;
      $lasttagsindex = $len - 1;

      my %openers_hash;

      for (my $i = 0; $i < $len; $i++) {
        if ($tags[$i]{type} == -1) {
          push(@cluster, { openpos => $tags[$i]{openpos}, closepos => $tags[$i]{closepos} });
        } else {
          @tags2[$i] = { tag => $tags[$i]{tag},
                         tagname => $tags[$i]{tagname},
                         type => $tags[$i]{type},
                         tag => $tags[$i]{tag},
                         openpos => $tags[$i]{openpos},
                         closepos => $tags[$i]{closepos},
                         cluster => $tags[$i]{cluster} };
        }

        if ($i == $lasttagsindex or $tags[$i + 1]{cluster} > $tags[$i]{cluster}) {
          if (scalar @cluster) {
            my @sorted = sort { $b->{openpos} <=> $a->{openpos} } @cluster;
            my $tagpos = $cluster[0]{closepos};
            for my $mdata (@sorted) {
              my $index = @{$mdata}{closepos};
              @tags2[$tagpos] = { tag => $tags[$index]{tag},
                             tagname => $tags[$index]{tagname},
                             type => $tags[$index]{type},
                             tag => $tags[$index]{tag},
                             openpos => $tags[$index]{openpos},
                             closepos => $tagpos,
                             cluster => $tags[$index]{cluster} };
              my $openerindex = $tags[$index]{openpos};
              $openers_hash{$openerindex} = $tagpos;
              $tags2[$openerindex]{closepos} = $tagpos;
              $openers_hash{$openerindex} = $tagpos;
              $tagpos++;
            }
          }
          @cluster = ();
        }
      }

      @tags = @tags2;

if ($switch) {
  print " --------------------------\n";
  print "after reorder closing tags 2:\n\n";
  #print Dumper(\@tags);
}

      ## Rebuild the doc string with the reordered tag objects.

      my $arrstr = $ostrbuilder;
      $arrstr =~ s@<@SPRTRSPRTR<@g;
      $arrstr =~ s@>@>SPRTRSPRTR@g;
      @seg2s = split(/SPRTRSPRTR/, $arrstr);

      $ostrbuilder = '';
      my $tagcount = 0;

      for (my $i = 0; $i < scalar @seg2s; $i++) {
        my $seg2 = $seg2s[$i];
        if ($seg2 eq '') {
          next;
        }
        if ($seg2 =~ m/^</) {
          $ostrbuilder .= $tags[$tagcount]{tag};
          $tagcount++;
        } else {
          $ostrbuilder .= $seg2;
        }
      }
    }

if ($switch) {
  print " --------------------------\n";
  print "after reorder tags 2:\n\n";
  print $ostrbuilder . "\n\n";
}














    ######################################################################
    ## (Again) Remove all containers that are broken by only whitespace.
    ## (Not sure if we still need to mask <p>'s here, but paranoia overrides.)

    $ostrbuilder =~ s@<p@<P@g;
    $ostrbuilder =~ s@</p@</P@g;
    while ($ostrbuilder =~ s@</([a-z0-9]+)>($WS_HOR*)<\1>@$2@) {}
    $ostrbuilder =~ s@<P@<p@g;
    $ostrbuilder =~ s@</P@</p@g;

if ($switch) {
  print " --------------------------\n";
  print "after remove all containers that are broken by only whitespace 2:\n\n";
  print $ostrbuilder . "\n\n";
}

    ######################################################################
    ## (Again) Minimize containers (shift whitespace on rhs of opening tag to lhs,
    ## shift whitespace on lhs of closing tag to rhs.)

    while($ostrbuilder =~ s@(<[^/][^>]*>)($WS_ALL+)@$2$1@g) {}
    while($ostrbuilder =~ s@($WS_ALL+)(</[^>]+>)@$2$1@g) {}

if ($switch) {
  print " --------------------------\n";
  print "after minimize containers 2:\n\n";
  print $ostrbuilder . "\n\n";
}

    ######################################################################
    ## Remove containers that contain no content.
    ## (the regex says: find a matching opening and closing element which
    ##  is separated only by tag(s) or whitespace, and eliminate the opening
    ##  and closing tag, and be non-greedy about it (or else
    ##  <a></a><a></a> will become </a><a> - not good))
    ## NOTE: I believe this gets superceded (for some reason)
    ## NOTE: If this does get used, will need to update for <p>'s which can have
    ## a space and more text after the tagname, but also taking into account
    ## you can have a tagnames like fs2 and fs20.
    #while ($ostrbuilder =~ s@<([a-z0-9]+)[^<>]*>((?:<(?!\1|/\1)[^<>]*>|$WS_ALL)*)</\1>@$2@) {}

    ######################################################################
    ## Remove containers that only contain whitespace.

    $seg =~ s@<p@<P@g;
    $seg =~ s@</p@</P@g;
    while ($seg =~ s@</([a-z0-9]+)>($WS_HOR*)<\1>@$2@) {}
    while ($seg =~ s@</([a-z0-9]+)>($WS_HOR*)<\1 [^>]+>@$2@) {}
    $seg =~ s@<P@<p@g;
    $seg =~ s@</P@</p@g;

if ($switch) {
  print " --------------------------\n";
  print "(final) after remove containers that contain no content:\n\n";
  print $ostrbuilder . "\n\n";
}


    $os .= $ostrbuilder . "\n\n";
  }

#print $os . "\n\n";
  return "\n\n".$os;
}


####################################################
####### @SUBROUTINE: NEST TAGS IN LISTS

## this works by the same principle as nestTags, only it is
## is done individually for lists, and each new line is
## considered a block container instead of \n\n.

sub nestTagsList () {

  my $ts = shift;

  my $os = '';

  my @tag_arr_1 = ();
  my @tag_arr_2 = ();

  for my $sect (split(/\n\n/, $ts)) {

    if ($sect eq  '') { next; }

    if ($sect !~ m@LSTSTUBCODE_TYPE_@) { $os .= $sect."\n\n"; next; }

    my $sect_out = '';

    for my $seg (split(/\n/, $sect)) {

      if ($seg eq  '') { next; }

      ##  make span outermost elements
      while ($seg =~ s@^($NC_ALL*)(<(?!span)[^<>]+>)($WS_ALL*)(<span[^<>]*>)@$1$4$3$2@) {}
      while ($seg =~ s@(</span>)($WS_ALL*)(</(?!span)[^<>]+>)($NC_ALL*)@$3$2$1$4@) {}

      my $seg_out = join('', @tag_arr_1);

      $seg =~ s@<@FLDSPRTR<@g;
      $seg =~ s@>@>FLDSPRTR@g;

      for my $seg2 (split(/FLDSPRTR/, $seg)) {

        ## if it's an opening tag, push it and the
        ## element name onto respective tag arrays
        if ($seg2 =~ m@^(<([^/<>][^ <>]*)[^<>]*>)@) {

          push(@tag_arr_1, $1);
          push(@tag_arr_2, $2);

        ## if it's a closing tag, we want to check
        ## and see that it is nested properly...
        }elsif ($seg2 =~ m@^</([^<>]+)>@) {

          my $tag_name = $1;

          ## create temp arrays to store what we
          ## push off of the tag arrays - this is
          ## done in a corresponding fashion
          my @tmp_arr_1 = ();
          my @tmp_arr_2 = ();

          ## declare some variables we'll be using
          my $insert;
          my $len = @tag_arr_1;
          my $cnt = 0;

          ## pop the last tag and name off the tag arrays
          my $test_tag  = pop @tag_arr_1;
          my $test_name = pop @tag_arr_2;

          while ($cnt < $len) {

            ## compare our closing tag name with the element
            ## we just popped off the tag array

            ## if no:
            if ($tag_name ne $test_name) {

              ## build our insert string with closing tags
              ## needed to preserve nesting
              $insert .= '</'.$test_name.'>';

              ## place unqualified opening tag and name
              ## on the beginning of respective arrays
              unshift (@tmp_arr_1, $test_tag);
              unshift (@tmp_arr_2, $test_name);

              ## pop again for the next round
              $test_tag  = pop @tag_arr_1;
              $test_name = pop @tag_arr_2;

            ## else, we're done (almost)
            }else{
              last;
            }
            $cnt++;
          }

          ## prepend our principle closing tag (this segment)
          ## with our insert string
          $seg2 = $insert.$seg2;

          ## append tag arrays with the temp arrays
          ## (put back the elements that need to be re-opened)
          push(@tag_arr_1, @tmp_arr_1);
          push(@tag_arr_2, @tmp_arr_2);

          my $len;

          ## append our principle closing tag (this segment)
          ## with the elements of the temp array
          while ($len = @tmp_arr_1) {
            $seg2 .= shift(@tmp_arr_1);
            shift(@tmp_arr_2);
          }
        }

        ## string builder
        $seg_out .= $seg2;
      }

      my $len = @tag_arr_2;
      my $cnt = $len - 1;

      while ($cnt > -1) {
        $seg_out .= '</'.$tag_arr_2[$cnt].'>';
        $cnt--;
      }

      $seg_out .= "\n";

      $sect_out .= $seg_out;
    }

    ##  remove empty containers
    while ($sect_out =~ s@<([a-z]+)[^<>]*>((?:<(?!\1|/\1)[^<>]*>|$WS_ALL)*)</\1>@$2@) {}

    $os .= $sect_out."\n";
  }

  return "\n\n".$os;

}


####################################################
####### @SUBROUTINE:  MARK UP LISTS

sub markUpLists () {

  my $ts = shift;

  my $os = '';

  for my $sect (split(/\n\n/, $ts)) {

    if ($sect eq  '') { next; }

    if ($sect =~ s@(LSTSTUBCODE_TYPE_.*\n)@@) {

      my $list_code = $1;
      $sect =~ s@\nLSTSTUBCODE_END@@;
      $list_code =~ m@LSTSTUBCODE_TYPE_([uo]l)_ID_([0-9]*)@;
      my $list_type = $1;
      my $list_id   = $2;
      my $tag_opn   = '<'.$list_type.'>';
      if ($list_id) { $tag_opn = '<'.$list_type.' class="l'.$list_id.'">'; }
      $sect =~ s@\n@</li>\n<li>@g;
      $sect = $tag_opn."\n<li>".$sect."</li>\n</".$list_type.'>';

    }

    $os .= $sect."\n\n";

  }

  return "\n\n".$os;

}

####### END @SUBROUTINE:  MARK UP LISTS
####################################################


####################################################
####### @SUBROUTINE: RESOLVE PARAGRAPH AMIBIGUITIES

## if there are ambiguities in paragraph blocks,
## ie, <p> and <p class="bq"> sections in the same
## paragraph*, attempt to resolve them to one or the
## other, based on whether $RESOLVE_TO_BQ is set or not.
##   * paragraph is defined here as being a block of text
##     that does not contain two or more consecutive newlines

sub paragraphAbiguities () {

  my($ts) = @_;

  my $os = '';
  my $msg = '';

  for my $seg (split(/\n\n/, $ts)) {

    if ($seg eq '') { next; }

    if ($seg =~ m@</p>.*</p>@s || $seg =~ m@<p.*<p@s) {

      my $p_start = '<p>';
      if ($RESOLVE_TO_BQ && $seg =~ m@<p class="bq@) { $p_start = '<p class="bq">'; }

      $seg =~ s@(?<!^)<p[^<>]*>@@g;
      $seg =~ s@</p>(?!$)@@g;
      $seg =~ s@^<p[^<>]*>@$p_start@;

      $seg =~ m@^(.{6,70})@s;

      $msg .= $1."...\n\n";

    }
    $os .= $seg."\n\n";
  }

  if ($msg) {
    print "\nWARNING : CONVERTER HAD TO MAKE SOME GUESSES IN : \n"
    .$outfile."\n\n".$msg."END WARNING\n";
  }

  return("\n\n".$os);

}

####### END @SUBROUTINE: RESOLVE PARAGRAPH AMIBIGUITIES
####################################################


####################################################
####### @SUBROUTINE: CREATE BQ BLOCKS

## find blocks of bq style paragraphs, and group them
## into single bq style div blocks
## NOTE: A lone paragraph will also be put in a div

sub createBQBlocks () {

  my($ts) = @_;
  my @bq_array;
  my @ts_array = split(/\n\n/, $ts);

  $ts =~ s@^\n+@@;
  $ts =~ s@\n+$@@;

  for my $seg (@ts_array) {

    if ($seg =~ m@^<p class="bq">@) {
      push(@bq_array, 1);
    }else{
      push(@bq_array, 0);
    }
  }

  my $cnt  = 0;
  my $last = @ts_array - 1;
  my $os = '';
  for my $seg (@ts_array) {

    if ($bq_array[$cnt]) {

      $seg =~ s@^<p class="bq">@<p>@;

      if ($cnt == 0 || !$bq_array[$cnt - 1]) {

        $seg =~ s@^<p>@<div class="bq">\n<p>@;
      }
      if ($cnt == $last || !$bq_array[$cnt + 1]) {

        $seg .= "\n</div>";
      }
    }
    $os .= $seg."\n\n";
    $cnt++
  }

  $os =~ s@^\n+@@;
  $os =~ s@\n+$@@;

  return "\n\n".$os."\n\n";
}

####### END @SUBROUTINE: CREATE BQ BLOCKS
####################################################


#################################################################
#################################################################
#################################################################
##
##      TOOLS
##


####################################################
####### @SUBROUTINE:  INTERCEPT RTF

sub interceptRTF() {

  my $fs = shift;

  open (FH, '>', $temp_rtf_file_1);
  print FH $fs;
  close (FH);

  my $ts = `textutil -convert txt -stdout $temp_rtf_file_1`;

  print $ts;
  exit;
}

sub formatForRTFCompare {
  my $fs = shift;

#return $fs;

  $fs =~ s@<@__SPRTR__<@g;
  $fs =~ s@>@>__SPRTR__@g;
  my @fsarr = split('__SPRTR__', $fs);
  my $len = scalar @fsarr;
  my @segarr = ();

#print Dumper(\@fsarr);

  # Sanitize into a nice array.
  for (my $i = 0; $i < $len; $i++) {
    my $seg = $fsarr[$i];
    if ($seg eq '') {
      next;
    }
#print " >> " . $seg . "\n";
    if ($seg =~ m@^</@) {
      push(@segarr, $seg);
    } else {
      # Separate whitespace from non-whitespace.
      for my $subseg (split(/(\s+)/, $seg)) {
        push(@segarr, $subseg);
      }
    }
  }

  my %ignore;
  my $len = scalar @segarr;
  my $lastindex = $len - 1;

  my $lasttextindex = -2;

  for (my $i = 0; $i < $len; $i++) {
    my $seg = $segarr[$i];

    if ($seg =~ m@^<@) {
      # segment is tag
      if ($seg =~ m@^</@) {
        if ($lasttextindex != $i - 1) {
          $segarr[$lasttextindex] .= $seg;
          $ignore{$i} = 1;
        }
      }
    } else {
      # segment is text
      if ($seg !~ m@^\s*$@) {
        $lasttextindex = $i;
      }
    }
  }

  my $prevtextindex = -2;

  for (my $i = $lastindex; $i > -1; $i--) {
    my $seg = $segarr[$i];

    if ($seg =~ m@^<@) {
      # segment is tag
      if ($seg !~ m@^</@) {
        if ($prevtextindex != $i - 1) {
          $segarr[$prevtextindex] = $seg . $segarr[$prevtextindex];
          $ignore{$i} = 1;
        }
      }
    } else {
      # segment is text
      if ($seg !~ m@^\s*$@) {
        $prevtextindex = $i;
      }
    }
  }

  my $os = '';

  for (my $i = 0; $i < $len; $i++) {
    my $seg = $segarr[$i];
    if (!$ignore{$i}) {
      $os .= $seg;
    }
  }

  $os =~ s@\n\n\n+@\n\n@g;

  return $os;
}

sub get_arg {

    my $x;
    my @a = @ARGV;
    my($needle, $mode) = @_;

    while ($x = shift(@a)) {
        if ($x eq $needle) {
            if ($mode == 2) {
                my $y = shift(@a);
                if ($y =~ /^-/ || $y eq "") {
                    print "invalid parameter for $x\n\n";
                    exit;
                }else{
                    return $y;
                }
            }else{
                return 1;
            }
            last;
        }
    }
    return 0;
}
