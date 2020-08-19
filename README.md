# rtf2html
RTF to HTML converter

### A basic RTF to HTML converter written in Perl.

This works for rich text RTF, but does not handle attachments.  However you can supply an RTFD file and it will deal with the text and just ignore the attachments.

It was designed around RTF files generated from Apple products such as Bean or Pages.  Non-Apple products handle file generation with slight differences, so if there are any quirks when feeding it RTF files generated on non-Apple products, you can simply open the non-Apple file in Bean, make a change (seems to be necessary) then save, in order to "launder" the file to an Apple-generated file.

### There are lots of other converters out there, why use this one?

RTF files record the history of every single change the user makes, even if subsequent changes render the previous ones moot.  So you end up with artifacts from doing-undoing font size, color changes etc, copy and paste etc, which results in lots of redundant code eg recording style changes of whitespace between words, redundant styling splitting words, etc.  Most RTF to HTML converters just follow what the RTF coding is and thus spit out a gazillion span or font tags, with a gazillion style declarations, resulting in a human-unreadable mess.

The advantage to this script is that it endeavors with great pains to spit out very clean HTML, more like one would do if they were hand-coding.  Paragraphs are separated with P tags which are given appropriate classes for indent, text-align, etc. There are no font tags, and all inline styles eg font-size changes etc are done with span tags given appropriate classes.

Its limitations are that it does not deal with tables and more complex RTF formatting.  It is great for things such as articles which are primarily text, and while not thoroughly tested, should be able to deal with lists to some degree.

It is all handled by the rtf2html perl script.  Just make the script executable, then type rtf2html with no args to get a detailed help page.  All the instructions are there.
