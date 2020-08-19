# rtf2html
RTF to HTML converter

### A basic RTF to HTML converter written in Perl.

This works for rich text RTF, but does not handle attachments.  However you can supply an RTFD file and it will deal with the text and just ignore the attachments.

It was designed around RTF files generated from Apple products such as Bean or Pages.  Non-Apple products handle file generation with slight differences, so if there are any quirks when feeding it RTF files generated on non-Apple products, you can simply open the non-Apple file in Bean, make a change (seems to be necessary) then save, in order to "launder" the file to an Apple-generated file.

### There are lots of other converters out there, why use this one?

RTF files record the history of every single change the user makes, even if subsequent changes render the previous ones moot.  So you end up with artifacts from doing-undoing font size, color changes etc, copy and paste etc, which results in lots of redundant code eg recording style changes of whitespace between words, redundant styling splitting words, etc.  Most RTF to HTML converters just follow what the RTF coding is and thus spit out a gazillion span or font tags, with a gazillion style declarations, resulting in a human-unreadable mess.

The advantage to this script is that it endeavors with great pains to spit out very clean HTML, more like one would do if they were hand-coding.  Paragraphs are separated with P tags which are given appropriate classes for indent, text-align, etc. There are no font tags, and all inline styles eg font-size changes etc are done with span tags given appropriate classes.

Its limitations are that it does not deal with tables and more complex RTF formatting.  It is great for things such as articles which are primarily text, and while not thoroughly tested, should be able to deal with lists to some degree.

It is all handled by the rtf2html.pl script.  Just make the script executable, then type rtf2html with no args to get a detailed help page.  All the instructions are there.  Note that the sourcefile arg comes first.

### Usage:

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
