#!/usr/bin/env tclsh
#-*- mode: tcl; coding: utf-8-unix; fill-column: 80; -*-

# -----------------------------------------------------------------------------
# note:
# to comments all ::log::log lines
# M-x query-replace-regexp  ^\([[:space:]]*\)\(::log::log .*\)$ → \1# \2
# to uncomment all of them
# M-x query-replace-regexp  ^\([[:space:]]*\)#[[:space:]]*\(::log::log .*\)$ → \1\2
# -----------------------------------------------------------------------------


#╔═════════════════════════╗
#║ *** tdomods-1.0.tm  *** ║
#╚═════════════════════════╝

# A package to transfer a table represented as list of lines, into an ods file using tdom
# ODS (Open Document Spreadsheets) is are part of the ODF (Open Document Format).
# (https://docs.oasis-open.org/office/OpenDocument/v1.3/)
# Such file are typically open and treated by the programme libreoffice.calc.
# (https://fr.libreoffice.org/discover/calc/)
# (https://us.libreoffice.org/discover/calc/)

package provide tdomods 1.0

# -----------------------------------------------------------------------------
#  This package provides a TclOO class also called tdomods, which generates
#  objects, that read tcl tables stored as a list of lines and write them into
#  ODS files.

#  It is understood that each line represents a record, and that each column
#  stands for a possible field of the records. Each member of the same column is
#  of then of the same type and each column has its own format of values.

#  If you seek for other possibilities to recover tables in Tcl and on the best
#  way to transfer them into list of lines, see the package datatable, that can
#  do it for you.
# -----------------------------------------------------------------------------

# -----------------------------------------------------------------------------

#  Main commands for practical use:

# +----------------------------------+----------------------------------+
# | set obj [tdomods new ?options?]  | create the object*               |
# +----------------------------------+----------------------------------+
# | $obj read $data                  | read the list of lines $data     |
# +----------------------------------+----------------------------------+
# | $obj insertHeaders hdLst         | Insert columns headers           |
# +----------------------------------+----------------------------------+
# | $obj addAutoFilter               | Add auto filter to the table     |
# +----------------------------------+----------------------------------+
# | $obj write file.ods              | Write table in the ods file      |
# +----------------------------------+----------------------------------+
# | $obj destroy                     | Released the memory (erase all   |
# |                                  | except the written file)         |
# +----------------------------------+----------------------------------+

# note*: alternatively one can use this instruction to create a TclOO object
#          tdomods create obj ?options?
# In that case, if created a object command 'obj' instead of an object address '$obj'
# (see TclOO help for further details)

# -----------------------------------------------------------------------------

# Limitations:

# The program try to recognize the type of data silently. If it doesn't succeed,
# it simply retrieves the data as strings.

# The decimals values are recognized either when they can be recognized as such
# by Tcl, or when they are using the local representation of decimals (see
# commands of configuration hereafter if you want to change this behavior).

# The financial valyes are recognized with a limited number of currencies and so
# far only when the currency is after the number (and not the reverse as it is
# done in US style).

# The date and times are recognized and the program does it best for it with the
# different possibilities. Only the ones using only figures are working.

# The table cannot have more than 702 columns (limitation due to private method
# ColAB, to keep it simple).


# -----------------------------------------------------------------------------

# Commands of configuration:

# +----------------------------------+----------------------------------+
# | $obj deleteDatabseRanges ?$name? | if no name, delete all ranges    |
# |                                  | defined. Practical way ro remove |
# |                                  | autofilter if needed.            |
# +----------------------------------+----------------------------------+
# | $obj set list_of_NA              | set list of NA indictations**    |
# +----------------------------------+----------------------------------+
# | $obj set accuracy                | number of decimals by default    |
# +----------------------------------+----------------------------------+
# | $obj set decimal_sep             | for localization***              |
# | $obj set thousand_sep            |                                  |
# +----------------------------------+----------------------------------+
# | $obj listFonts                   | fonts declared in ods file       |
# +----------------------------------+----------------------------------+
# | $obj addFont $fontName           | add a font in the ods file       |
# +----------------------------------+----------------------------------+
# | $obj deleteFont $fontName        | delete a font from ods file      |
# +----------------------------------+----------------------------------+
# | $obj deleteAllFonts              | delete all fonts from ods file   |
# +----------------------------------+----------------------------------+
# | $obj asXML                       | retrieve what is (or will be) in |
# |                                  | content.xml)                     |
# +----------------------------------+----------------------------------+

# note**: NA are strings such as 'NA' (for non applicable), '-' …
# The programme won't try to recognize them as numeric data and just pass
# them to try the recognition on the next item in the same column.

# note***: for French speaking users, using different decimals and thousands
# separators. It is not recommended to use different settings from what is in
# your LibreOffice settings. By default, the package use the system settings, as
# LibreOffice is normally also doing. So do not use this option unless you know
# what you do.

# options of creation:
# the majority of the input command can be used, when creating the object.

# +------------------------------------+
# | set obj [tdomods new -head $hdLst] |
# +------------------------------------+
# | set obj [tdomods new -na $naLst]   |
# +------------------------------------+
# | set obj [tdomods new -data $data]  |
# +------------------------------------+
# | set obj [tdomods new -sheet $name] |
# +------------------------------------+
# | set obj [tdomods new -fmt mftLst]  |
# +------------------------------------+
# | set obj [tdomods new -decimal_sep  |
# +------------------------------------+
# | set obj [tdomods new -thousand_sep |
# +------------------------------------+

# One can add several options together on the same line.

# -----------------------------------------------------------------------------


# -----------------------------------------------------------------------------
# dependencies
# -----------------------------------------------------------------------------

package require tdom

#  The package is using tdom extensively to build the xml files required in an
#  ods file. An ods file is a renammed zip file containing a full directory
#  structure containing mainly XML files. The minimum expected files needed, so
#  the ods file can be opened by libreoffice.calc are the followings:
#
# ODS file
#   │
#   ├── content.xml       : le contenu du fichier
#   │
#   ├── META-INF
#   │    │
#   │    └── manifest.xml : déclaration du contenu
#   │
#   ├── mimetype          : déclaration du type MIME
#   ┆
#   └╌╌ styles.xml        : declaration of styles
#
#  The file styles.xml isn't strictly necessary and it hasn't been developped
#  here: all styles are declared in the preamble of content.xml as
#  automatic-styles elements.

#  When the ods file is opened a restored by libreoffice.calc, some other files
#  are automatically appended to this list and the files themselves are also
#  completed with other informations. So the first opening of the file may takes
#  a bit more time, but the file is finally totally useable.


# -----------------------------------------------------------------------------
# create procedure lshift if missing. It is used to analyze optional arguments.
if {[string length [info procs lshift]] == 0} {
    uplevel #0 {
	# ---------------------------------------------------------------------
	# lshift list
	# ---------------------------------------------------------------------
	# from an orginal idea taken here: https://wiki.tcl-lang.org/page/lshift
	# ! arrgument by adress ...
	# Used in loops
	# ---------------------------------------------------------------------
	proc lshift list {
	    upvar 1 $list L
	    set R [lindex $L 0]
	    set L [lreplace $L [set L 0] 0]
	    # set L [lrange $L 1 end]
	    return $R
	}
	# ---------------------------------------------------------------------
    }
}
# -----------------------------------------------------------------------------

#------------------------------------------------------------------------------
#  package for supporting debugging
#------------------------------------------------------------------------------
# you may comment the two active lines
package require log
# activate all levels of message for log
::log::lvSuppressLE emergency 0
# desactivate debug message
# ::log::lvSuppress debug
#┌──────────────────────────────────┐
#│ levels of log records:      	    │
#├──────────────────────────────────┤
#│ - emergency                      │
#│ - alert                          │
#│ - critical                       │
#│ - error                          │
#│ - warning                        │
#│ - notice <--- logLevel variable  │
#│ - info........x		    │
#│ - debug.......x		    │
#│                                  │
#│https://wiki.tcl-lang.org/page/log│
#└──────────────────────────────────┘
#------------------------------------------------------------------------------



# -----------------------------------------------------------------------------
# Class to create domdoc object containing a data table with the formating
# used by content.xml in ods file
# -----------------------------------------------------------------------------
::oo::class create tdomods {

    variable ODFV     ; # odf version, 1.2 or 1.3

    variable MIMETYPE ; # content of mimetype file
    variable MANIFEST ; # content file content.xml
    variable CONTENT ;  # domdoc object for the file content.xml
    variable AUTOSTYLES ;  # domnode object for the list of automatic styles applicable
    # variable STYLES ; domndoc object for the file style.xml

    variable TABLE  ;  # domnode object of the table

    variable NA     ;  # list of string recognised as NA
    variable DS     ; # default decimal separator
    variable TS     ; # default thousand seperator
    variable AC     ; # default number of digit after coma

    variable COLFMT  ;  # array giving the format name for each column name

    # The elements of the FMT list will follow the following naming principles
    #
    # |---|-------------|-----------------------|-------------------------------|
    # | # | styles for: | format of style name  | what it defines               |
    # |---|-------------|-----------------------|-------------------------------|
    # | 1 | values      | start with V, then    | defines the way the value     |
    # |   |             | defines with codes    | is presented.                 |
    # |   |             | (see below).          |                               |
    # |---|-------------|-----------------------|-------------------------------|
    # | 2 | columns     | start with 'ce', then | link to value format, defines |
    # |   |             | number column number. | alignment (background...)     |
    # |---|-------------|-----------------------|-------------------------------|
    # | 3 | rows        | only two rows, header | for rows as for columns.      |
    # |   |             | and other. ro1, ro2   |                               |
    # |---|-------------|-----------------------|-------------------------------|
    #
    # For value format:
    #
    #  |---|------------|--------------------------------|----------|------------|
    #  | 1 | 2nd pos.   | 3th and following pos.         | format   | example    |
    #  |---|------------|--------------------------------|----------|------------|
    #  | V | F:float    | D2: 2 decimals                 | VFD2     |   "456,30" |
    #  |   |            | SS: separated by space         | VFD2SS   | "1 456,30" |
    #  |   |            | M: minimum digit               | VFD2M0   |       ,00  |
    #  |   |------------|--------------------------------|----------|------------|
    #  |   | I:integer  | D4: 4 digits minimum (with 0   | VID4     | "0001"     |
    #  |   |            | upfront)                       |          |            |
    #  |   |            | SS: thousands separatd by space| VISS     | "1 450"    |
    #  |   |            |                                |          |            |
    #  |   |------------|--------------------------------|----------|------------|
    #  |   | C:currency | D2SSC€ : 2 decimals, thousands | VCD2SSCE |"1 345,45 €"|
    #  |   |            | separated by space, currency € |          |            |
    #  |   |------------|--------------------------------|----------|------------|
    #  |   | D:date     | D2M2Y4 : 2 digits for days, 2  | VDD2M2Y4 |"02/04/2010"|
    #  |   |            | for month, 4 for years         |          |            |
    #  |   |------------|--------------------------------|----------|------------|
    #  |   | T:time     | H2: hours on two digits        | VTD2M2   | "15:35"    |
    #  |   |            | M2: minutes on two digits      |          |            |
    #  |   |------------|--------------------------------|----------|------------|
    #  |   | R:datetime |  VRD2M2Y4TH2M2                 |          |            |
    #  |   |            |                                |          |            |
    #  |   |------------|--------------------------------|----------|------------|
    #  |   | S: string  | C10 : 10 charaters             | VSC10    |"du texte!" |
    #  |---|------------|--------------------------------|----------|------------|
    #


    # ------------------------------------------------------------------------
    # domcontent create doc $args ?options?
    # set obj [domcontent new $args ?options?]
    # ------------------------------------------------------------------------
    # options are:
    #  -fmt          : list of format (tcl style) for columns
    #  -head         : list of colums headers
    #  -na           : list of non-applicable values (pass when recognizing data type)
    #  -data         : the datatable (a list of lines)
    #  -sheet        : name of the worksheet to be built
    #  -decimal_sep  : default decimal separator (default is to take the system value)
    #  -thousand_sep : default thousand separator (default is to take the system value)
    # ------------------------------------------------------------------------
    constructor args {

	# default values for options
	set NA {"" " " "NA" "-"}
	set SheetName Feuille1
	set ODFV 1.3  ; # by default we declare version 1.3 for ODF file
	while {[::string index $args 0] eq "-"} {
	    set WORD [lshift args]
	    switch -glob -nocase -- $WORD {
		"-fmt"  {set FMT [lshift args]}
		"-head*" - "tit*" {set HEADERS [lshift args]}
		"-na"    {set NA [lshift args]}
		"-data"  {set DATA [lshift args]}
		"-sheet" {set SheetName [lshift args]}
		"-decimal_sep"  {set chDS [lshift args]}
		"-thousand_sep" {set ChTS [lshift args]}
		"-odf_v*" {set ODFV [lshift args]}
	    }
	}
	# ::log::log debug "args://$args//"
	if {[llength $args] > 0} {set DATA $args}
	# --------------------------------------------------------------------

	# --------------------------------------------------------------------
	# Get decimal and thousand separators from system to accomodate French
	# or US notation of decimals. Grouping is always by 3.
	# This default can be update by arguments (see above).
	foreach L [exec locale -k LC_NUMERIC] {
	    set L [string map {= \ } $L]
	    switch [lindex $L 0] {
		"decimal_point" {set DS [lindex $L 1]}
		"thousands_sep" {set TS [lindex $L 1]}
	    }
	}
	if [info exists chDS] {
	    if {$chDS ni {, .}} {
		error "-decimal_sep mus be `.` or `,` and not $chDS" "wrong -decimal_sep given to build object tdomods" 403
	    }
	    set DS $chDS
	}
	if [info exists chTS] {
	    if {$chTS ni {, .}} {
		error "-thousand_sep mus be `,` or ` `, not $chTS" "wrong -thousand_sep given to build object tdomods" 403
	    }
	    set TS $chTS
	}
	# If everything fails, we still want to initialize both variables, French by default :-)
	if {[string length $DS] != 1} {set DS ","}
	if {[string length $TS] != 1} {set TS " "}

	# défault accuracy is always 2
	set AC 2
	# --------------------------------------------------------------------

	set MIMETYPE {application/vnd.oasis.opendocument.spreadsheet}

	set MANIFEST [dom createDocumentNS \
			  "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"\
			  manifest:manifest]
	set ROOT [$MANIFEST documentElement]
	$ROOT setAttribute \
	    xmlns:loext "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"\
	    manifest:version "$ODFV"

	set NODE [$MANIFEST createElement manifest:file-entry]
	$NODE setAttribute \
	    manifest:full-path "/"\
	    manifest:version "$ODFV"\
	    manifest:media-type "application/vnd.oasis.opendocument.spreadsheet"
	$ROOT appendChild $NODE

	set NODE [$MANIFEST createElement manifest:file-entry]
	$NODE setAttribute \
	    manifest:full-path "content.xml"\
	    manifest:media-type "text/xml"
	$ROOT appendChild $NODE

	# -------------------------------------------------------------------
	set CONTENT [dom createDocumentNS \
		     "urn:oasis:names:tc:opendocument:xmlns:office:1.0" \
		     office:document-content]
	set ROOT [$CONTENT documentElement]
	$ROOT setAttribute \
	    xmlns:table "urn:oasis:names:tc:opendocument:xmlns:table:1.0"\
	    xmlns:number "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"\
	    xmlns:text "urn:oasis:names:tc:opendocument:xmlns:text:1.0"\
	    xmlns:style "urn:oasis:names:tc:opendocument:xmlns:style:1.0"\
	    xmlns:fo "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"\
	    office:version "$ODFV"

	# the declaration of the font
	set FONTS [$CONTENT createElement office:font-face-decls]
	$ROOT appendChild $FONTS
	my addFont

	# the container for the automatic styles
	set AUTOSTYLES [$CONTENT createElement office:automatic-styles]
	$ROOT appendChild $AUTOSTYLES

	# the container for the body, a spreadsheet and a table
	set BODY [$CONTENT createElement office:body]
	$ROOT appendChild $BODY
	set SHEET [$CONTENT createElement office:spreadsheet]
	$BODY appendChild $SHEET
	set TABLE [$CONTENT createElement table:table]
	$TABLE setAttribute table:name $SheetName
	$SHEET appendChild $TABLE

	# insert a line of headers if applicable
	if [info exists HEADERS] {my insertHeaders $HEADERS}
	
	# recover the style from the DATA
	if [info exists DATA] {my read $DATA}
    }
    # ------------------------------------------------------------------------


    # ------------------------------------------------------------------------
    # $obj destroy
    # ------------------------------------------------------------------------
    destructor {
	unset NA
	unset -nocomplain COLFMT
	unset MIMETYPE
	$MANIFEST delete
	$CONTENT delete
    }
    # ------------------------------------------------------------------------


    # ------------------------------------------------------------------------
    # $obj set decimal_sep ?val?
    # $obj set thousand_sep ?val?
    # $obj set accuracy ?val?
    # $obj set list_of_NA ?val?
    # ------------------------------------------------------------------------
    # set the decimal or thousand separator for the object. If no value given,
    # return the actual value stored in the object (as the `set` tcl command).
    # ------------------------------------------------------------------------
    method set args {
	set what [lshift args]
	# ::log::log debug "what: $what || args: $args"
	switch -nocase -glob -- $what {
	    "dec*" {
		switch [llength $args] {
		    "0" {return $DS}
		    "1" {set $DS $args}
		    default {
			error "command is `obj set decimal_sep` with only one char"\
			    "error with obj set" 403
		    }
		}
	    }
	    "thou*" {
		switch [llength $args] {
		    "0" {return $TS}
		    "1" {set $TS $args}
		    default {
			error "command is `obj set thousand_sep` with only one char"\
			    "error with $obj set" 403
		    }
		}
	    }
	    "acc*" {
		switch [llength $args] {
		    "0" {return $AC}
		    "1" {
			if ![string is integer $args] {
			    error "command is `obj set accuracy` with one integer"\
				"error with obj set" 403
			}
			set $AC $args
		    }
		    default {
			error "command is `obj set accuracy` with only one integer"\
			    "error with obj set" 403
		    }
		}
	    }
	    "list_of_NA" - "NA" {
		if {[llength $args] == 0} {
		    return $NA
		} {	
		    set NA $args
		}
	    }
	    "odf*" {
		if {[llength $args] == 0} {
		    return $ODFV
		} {
		    if {$args ni {1.2 1.3}} {
			error "version of ODF is 1.2 or 1.3"\
			    "error in setting ODF version" 403		
		    }
		    set ODFV $args
		    set ROOT [$CONTENT documentElement]
		    $ROOT setAttribute office:version "$ODFV"
		    set ROOT [$MANIFEST documentElement]
		    $ROOT setAttribute manifest:version "$ODFV"
		}
	    }
	    default {
		error "command is `obj set decimal_sep` or \
		    `obj set thousand_sep` or `obj set accuracy` or \
		    `obj set list_of_NA` or `obj set odf_version`\
		    and not $what" "error with obj set" 403
	    }
	}
    }
    # ------------------------------------------------------------------------

    # ------------------------------------------------------------------------
    # RETRIEVING INFORMATION
    # ------------------------------------------------------------------------

    # ------------------------------------------------------------------------
    #  $obj asXML
    # ------------------------------------------------------------------------
    # retrieve the content formatted in XML
    # ------------------------------------------------------------------------
    method asXML {} {
	return [$CONTENT asXML]
    }
    # ------------------------------------------------------------------------

    # ------------------------------------------------------------------------
    # FONTS MANAGEMENT
    # ------------------------------------------------------------------------

    # ------------------------------------------------------------------------
    # $obj fontList
    # ------------------------------------------------------------------------
    # return the list of fonts declared
    # ------------------------------------------------------------------------
    method listFonts {} {
	set ROOT [$CONTENT documentElement]
	set FONTS [$ROOT getElementsByTagName office:font-face-decls]

	if ![$FONTS hasChildNodes] return
	foreach F [$FONTS childNodes] {
	    lappend RES [$F @style:name]
	}
	return $RES
    }

    # ------------------------------------------------------------------------
    #  $obj addFont ?name?
    # ------------------------------------------------------------------------
    # return the fragment domnode in the DOM tree CONTENT for the font element
    # ------------------------------------------------------------------------
    method addFont {{name "Liberation Sans"}} {
	set ROOT [$CONTENT documentElement]
	set FONTS [$ROOT getElementsByTagName office:font-face-decls]

	if {$name in [my listFonts]} return

	set FONT [$CONTENT createElement style:font-face]
	# --------------------------------------------------------------------
	# >> more choice of fonts to be inserted here
	# --------------------------------------------------------------------
	switch -nocase -- $name {
	    "Liberation Sans" {
		set generic swiss
		set pitch variable
	    }
	    "FreeSans" {
		set generic swiss
		set pitch variable
	    }
	    "Lating Modern Sans" {
		set generic swiss
		set pitch variable
	    }
	    "Courier" {
		set generic monospace
		set pitch fixed	
	    }
	    "FreeMono" {
		set generic monospace
		set pitch fixed	
	    }
	    "Latin Modern Mono" {
		set generic monospace
		set pitch fixed	
	    }
	    "Times" {
		set generic roman	
		set pitch variable
	    }
	    "FreeSerif" {
		set generic roman	
		set pitch variable
	    }
	    "Latin Modern Roman" {
		set generic roman	
		set pitch variable
	    }
	    "Century Schoolbook" {
		set generic roman
		set pitch variable
	    }
	    default {
		::log::log debug "use default parameters for fonts"
		set generic swiss
		set pitch variable
	    }
	}
	# --------------------------------------------------------------------
	$FONT setAttribute style:name $name\
	    style:font-family-generic $generic \
	    style:font-pitch $pitch
	$FONTS appendChild $FONT
	return $FONT
    }
    # ------------------------------------------------------------------------

    # ------------------------------------------------------------------------
    method deleteAllFonts {} {
	set ROOT [$CONTENT documentElement]
	set FONTS [$ROOT getElementsByTagName office:font-face-decls]

	if ![$FONTS hasChildNodes] return
	foreach F [$FONTS childNodes] {
	    $F delete
	}
	return
    }
    # ------------------------------------------------------------------------

    # ------------------------------------------------------------------------
    #  $obj deleteFont ?name?
    # ------------------------------------------------------------------------
    # remove a font
    # ------------------------------------------------------------------------
    method deleteFont fontname {
	set ROOT [$CONTENT documentElement]
	set FONTS [$ROOT getElementsByTagName office:font-face-decls]

	if ![$FONTS hasChildNodes] {return t}
	foreach F [$FONTS childNodes] {
	    if {[$F @style:name] eq $fontname} {
		$F delete
		return t
	    }
	}
	# the font was not found, then not deleted
	return f
    }
    # -------------------------------------------------------------------------

    # -------------------------------------------------------------------------
    #  PRIVATE METHODS NEEDED TO BUILT CONTENT.XML
    # -------------------------------------------------------------------------

    # -------------------------------------------------------------------------
    # $obj DateFmt $data
    # -------------------------------------------------------------------------
    # recognize a date and retuns itsformat, returns its code (see intro above)
    # and update the dom tree AUTOSTYLES
    # -------------------------------------------------------------------------
    method DateFmt data {

	# we have string, but have we a date or time ?
	# Only recognized formats for dates are :
	#   YYYY-MM-DD   ,   DD/MM/YYYY ,   DD/MM/YY , MM-DD-YYYY ,  MM-DD-YY
	# case of only one digit for day or month is recognized, but not followed (return 2 digits)
	if [regexp {([0-9]{4}|[0-9]{1,2})([-/])([0-9]{1,2})[-/]([0-9]{2}|[0-9]{4})} $data -> F1 SEP F2 F3] {
	    # The regexp is not meant to check the date is correct,
	    # but to recognize it as much as possible.
	    if {[string length $F1] == 4} {
		set FNAME VDY4M2D2

		# ::log::log debug "date recognized"
	
		# start with year, interpretation is: YYYY-mm-dd
		# we insert the style in the style dom tree if this
		# style of value is not existing
		if {$FNAME ni [dict values [array get COLFMT]]} {
		    set DATE [$CONTENT createElement number:date-style]
		    $DATE setAttribute style:name $FNAME
		    # $AUTOSTYLES appendChild $DATE

		    # ::log::log debug "DATE: [$DATE asXML]"
	
		    set NUM [$CONTENT createElement number:year]
		    $NUM setAttribute number:style long
		    $DATE appendChild $NUM
	
		    set NUM [$CONTENT createElement number:text]
		    $DATE appendChild $NUM
		    set TXT [$CONTENT createTextNode $SEP]
		    $NUM appendChild $TXT
	
		    set NUM [$CONTENT createElement number:month]
		    $NUM setAttribute number:style long
		    $DATE appendChild $NUM

		    set NUM [$CONTENT createElement number:text]
		    $DATE appendChild $NUM
		    set TXT [$CONTENT createTextNode $SEP]
		    $NUM appendChild $TXT
	
		    set NUM [$CONTENT createElement number:day]
		    $NUM setAttribute number:style long
		    $DATE appendChild $NUM
		} {
		    # retrieve the link to the correct DATE style node
		    set DSTY [$AUTOSTYLES getElementsByTagName number:date-style]
		    foreach STY $DTSY {
			if {[$STY @style:name ] eq $FNAME} {
			    set DATE $STY
			    break
			}
		    }
		}
		return [list $FNAME $DATE]
	
	    } elseif {$SEP eq "-"} {
		# non starting with YYYY, but a - as separator, US date style : MM-DD-YY(YY)
		set LEN [string length $F3]
		switch $LEN {
		    "2" {set FNAME VDM2D2Y2}
		    "4" {set FNAME VDM2D2Y4}
		    default {
			error "date $data unrecognized by DateFmt"\
			    "error in class todoms, private method DateFmt, unrecognized date format ($data)" 401
		    }
		}
		if {$FNAME ni [dict values [array get COLFMT]]} {
		    set DATE [$CONTENT createElement number:date-style]
		    $DATE setAttribute style:name $FNAME
		    # $AUTOSTYLES appendChild $DATE

		    set NUM [$CONTENT createElement number:month]
		    $NUM setAttribute number:style long
		    $DATE appendChild $NUM
		    	
		    set NUM [$CONTENT createElement number:text]
		    $DATE appendChild $NUM
		    set TXT [$CONTENT createTextNode $SEP]
		    $NUM appendChild $TXT
	
		    set NUM [$CONTENT createElement number:day]
		    $NUM setAttribute number:style long
		    $DATE appendChild $NUM

		    set NUM [$CONTENT createElement number:text]
		    $DATE appendChild $NUM
		    set TXT [$CONTENT createTextNode $SEP]
		    $NUM appendChild $TXT
	
		    set NUM [$CONTENT createElement number:year]
		    $NUM setAttribute number:style long
		    $DATE appendChild $NUM
		} {
		    # retrieve the link to the correct DATE style node
		    set DSTY [$AUTOSTYLES getElementsByTagName number:date-style]
		    foreach STY $DTSY {
			if {[$STY @style:name ] eq $FNAME} {
			    set DATE $STY
			    break
			}
		    }	
		}	
		return [list $FNAME $DATE]
	
	    } else {
		# European style starting with days : DD/MM/YY(YY)
		set LEN [string length $F3]
		# ::log::log debug "date ? : $data"
		switch $LEN {
		    "2" {set FNAME VDD2M2Y2}
		    "4" {set FNAME VDD2M2Y4}
		    default {
			error "date $data unrecognized by DateFmt"\
			    "error in class todoms, private method DateFmt, unrecognized date format ($data)" 401
		    }
		}
	
		if {$FNAME ni [dict values [array get COLFMT]]} {
		    # ::log::log debug "enregistrement de $FNAME"
		    set DATE [$CONTENT createElement number:date-style]
		    $DATE setAttribute style:name $FNAME
		    # $AUTOSTYLES appendChild $DATE

		    set NUM [$CONTENT createElement number:day]
		    $NUM setAttribute number:style long
		    $DATE appendChild $NUM
		    	
		    set NUM [$CONTENT createElement number:text]
		    $DATE appendChild $NUM
		    set TXT [$CONTENT createTextNode $SEP]
		    $NUM appendChild $TXT

		    set NUM [$CONTENT createElement number:month]
		    $NUM setAttribute number:style long
		    $DATE appendChild $NUM

		    set NUM [$CONTENT createElement number:text]
		    $DATE appendChild $NUM
		    set TXT [$CONTENT createTextNode $SEP]
		    $NUM appendChild $TXT
	
		    set NUM [$CONTENT createElement number:year]
		    $NUM setAttribute number:style long
		    $DATE appendChild $NUM
		} {
		    # retrieve the link to the correct DATE style node
		    set DSTY [$AUTOSTYLES getElementsByTagName number:date-style]

		    # ::log::log debug "DSTY: $DSTY"
	
		    foreach STY $DSTY {
			if {[$STY @style:name ] eq $FNAME} {
			    set DATE $STY
			    break
			}
		    }	
		}
		# ::log::log debug "DATE: $DATE"
		return [list $FNAME $DATE]
	    }
	} {
	    # no data recognition
	    return f
	}
    }
    # -------------------------------------------------------------------------


    # -------------------------------------------------------------------------
    # $obj TimeFmt $data
    # -------------------------------------------------------------------------
    # recognize a time format, returns its code (see intro above)
    # and update the dom tree AUTOSTYLES
    # -------------------------------------------------------------------------
    method TimeFmt data {

	if [regexp {[0-2][0-9]\:[0-5][0-9](\:[0-5][0-9])?} $data ->] {

	    set FNAME VTH2M2

	    if {$FNAME ni [dict values [array get COLFMT]]} {
		set TIME [$CONTENT createElement number:time-style]
		$TIME setAttribute style:name $FNAME
		# $AUTOSTYLES appendChild $TIME

		set NUM [$CONTENT createElement number:hours]
		$TIME appendChild $NUM
	
		set NUM [$CONTENT createElement number:text]
		$TIME appendChild $NUM
		set TXT [$CONTENT createTextNode \:]
		$NUM appendChild $TXT

		set NUM [$CONTENT createElement number:minutes]
		$TIME appendChild $NUM

	    } {
		# retrieve the link to the correct TIME style node
		set TSTY [$AUTOSTYLES getElementsByTagName number:time-style]
		foreach STY $TSTY {
		    if {[$STY @style:name ] eq $FNAME} {
			set TIME $STY
			break
		    }
		}	
	    }	
	    return [list $FNAME $TIME]
	} {
	    return f
	}
    }
    # -------------------------------------------------------------------------


    # -------------------------------------------------------------------------
    # $obj DateTimeFmt $data
    # -------------------------------------------------------------------------
    # recognize a time format, returns its code (see intro above)
    # and update the dom tree AUTOSTYLES
    # -------------------------------------------------------------------------
    method DateTimeFmt data {

	if [regexp {((([0-2][0-9])|([0-9]{4}))([-/])([0-5][0-9])\5(([0-9]{4})|([0-2][0-9])))[\s\t]+([0-2][0-9]:[0-5][0-9](:[0-5][0-9])?)} \
		$data -> d1 v2 m3 y3 s5 v6 y7 m8 m9 tA mB] {
	
	    if {[set RES [my DateFmt $d1]] ne "f" } {	
		set FNAME [lindex $RES 0]
		set DATE [lindex $RES 1]
	    } {
		return f
	    }
	    if {[set RES [my TimeFmt $tA]] ne "f" } {
		append FNAME [string range [lindex $RES 0] 1 end]
		# replace the D by a R to indicate a daytime
		set FNAME [string replace $FNAME 1 1 R]
		$DATE setAttribute style:name $FNAME

	
		set NUM [$CONTENT createElement number:text]
		$DATE appendChild $NUM
		set TXT [$CONTENT createTextNode " "]
		$NUM appendChild $TXT

		set NUM [$CONTENT createElement number:hours]
		$DATE appendChild $NUM
	
		set NUM [$CONTENT createElement number:text]
		$DATE appendChild $NUM
		set TXT [$CONTENT createTextNode ":"]
		$NUM appendChild $TXT

		set NUM [$CONTENT createElement number:minutes]
		$DATE appendChild $NUM
	    } {
		return f
	    }
	    return [list $FNAME $DATE]
	} {
	    return f
	}
    }
    # -------------------------------------------------------------------------
       

    # -------------------------------------------------------------------------
    #  $obj LstStyle lst
    #  ------------------------------------------------------------------------
    #  Recognize the style to apply on the elements a list (! must be a list),
    #  excluding the member of list NA from recognition. All elements are
    #  supposed to be compatible with a single style and the first positive
    #  recognition returns result.  Result are the tags presented in
    #  introduction and that will become the names of value styles.  This
    #  procedure update the dom tree AUTOSTYLES.
    #  ------------------------------------------------------------------------
    method LstStyle lst {
	
	foreach CELL $lst {
	    if {$CELL in $NA} continue

	    if [string is integer $CELL] {
	
		if {[string range $CELL 0 0] eq "0"} {
		    # integer start with 0, assuming fixed length
		    set LEN [string length $CELL]
		    set FNAME VID$LEN

		    # we insert the style in the style dom tree if this
		    # style of value is not existing
		    if {$FNAME ni [dict values [array get COLFMT]]} {
			set NUM [$CONTENT createElement number:number-style]
			$NUM setAttribute style:name $FNAME
			set DECI [$CONTENT createElement number:number]
			$DECI setAttribute number:decimal-places 0
			$DECI setAttribute number:min-decimal-places 0
			$DECI setAttribute number:min-integer-digits 4
			$NUM appendChild $DECI
			$AUTOSTYLES appendChild $NUM
		    }		    		    		   
		    return $FNAME
		   
		} {
		    # by defaut, thousand separed by space
		    set FNAME VISSM1

		    if {$FNAME ni [dict values [array get COLFMT]]} {
			set NUM [$CONTENT createElement number:number-style]
			$NUM setAttribute style:name $FNAME
			set DECI [$CONTENT createElement number:number]
			$DECI setAttribute number:decimal-places 0
			$DECI setAttribute number:min-decimal-places 0
			$DECI setAttribute number:min-integer-digits 1
			$DECI setAttribute number:grouping true
			$NUM appendChild $DECI
			$AUTOSTYLES appendChild $NUM
		    }
		    return $FNAME
		}
	
	    } elseif [string is double $CELL] {
		# convert the float to the locale stored in the object	
		set FNAME [join "VFD $AC  SSM1" ""]

		if {$FNAME ni [dict values [array get COLFMT]]} {
		    set NUM [$CONTENT createElement number:number-style]
		    $NUM setAttribute style:name $FNAME
		    set DECI [$CONTENT createElement number:number]
		    $DECI setAttribute number:decimal-places $AC
		    $DECI setAttribute number:min-decimal-places $AC
		    $DECI setAttribute number:min-integer-digits 1
		    $DECI setAttribute number:grouping true
		    $NUM appendChild $DECI
		    $AUTOSTYLES appendChild $NUM
		}
		return $FNAME
	
	    } else {
		# ::log::log debug "CELL is string: $CELL, to analyze"
	
		if {[set RES [my DateFmt $CELL]] ne "f" } {
		    # proc 'dataFmt' is update the variable 'style'
		    $AUTOSTYLES appendChild [lindex $RES 1]
		    return [lindex $RES 0]
		   
		}
		if {[set RES [my TimeFmt $CELL]] ne "f"} {		   
		    $AUTOSTYLES appendChild [lindex $RES 1]
		    return [lindex $RES 0]
		}
		if {[set RES [my DateTimeFmt $CELL]] ne "f"} {
		    $AUTOSTYLES appendChild [lindex $RES 1]
		    return [lindex $RES 0]	   
		}

		# >> recongnize the currency only in last position
		if [regexp -all {((\d{0,3}[ ]?)*),?(\d{0,3}[ ]?)*([€£F$])} $CELL RES INT X3 X4 CUR] {

		    # ::log::log debug "RES: $RES | INT: $INT | X3: $X3 | X4: $X4 | CUR: $CUR"
		   
		    # the string represents an amount with currency
		    switch $CUR {
			"€" {set FNAME VCD2SSCE}
			"$" {set FNAME VCD2SSCD}
			"F" {set FNAME VCD2SSCF}
			"£" {set FNAME VCD2SSCP}
			default {set FNAME VCD2SSCE}
		    }		
		   
		    if {$FNAME ni [dict values [array get COLFMT]]} {
			set NUM [$CONTENT createElement number:currency-style]
			$NUM setAttribute style:name $FNAME
			set DECI [$CONTENT createElement number:number]
			$DECI setAttribute number:decimal-places $AC
			$DECI setAttribute number:min-decimal-places $AC
			$DECI setAttribute number:min-integer-digits 1
			$DECI setAttribute number:grouping true
			$NUM appendChild $DECI
			set NTXT [$CONTENT createElement number:text]
			$NUM appendChild $NTXT
			set TXT [$CONTENT createTextNode " "]
			$NTXT appendChild $TXT
			set CTXT [$CONTENT createElement number:currency-symbol]
			$NUM appendChild $CTXT
			set TXT [$CONTENT createTextNode $CUR]
			$CTXT appendChild $TXT
		
			$AUTOSTYLES appendChild $NUM
		    } 		    		   
		    return $FNAME	   
		}
	
		if [regexp {\m((\d{0,3}[ ]?)+),?((\d{0,3}[ ]?)*)\M} $CELL NUM INT X2 DEC X4] {
		    # recognize a decimal in French style
		    # (otherwise the $CELL would have pass the preceding recognition test)

		    # ::log::log debug "$CELL // $NUM // $INT // $DEC"
		   
		    set FNAME [join "VFD $AC SSM1" ""]
		    # SS: thousand separated by space
		    # SC: thousand separated by coma
		    # ::log::log debug "FNAME: $FNAME"		   
		   
		    if {$FNAME ni [dict values [array get COLFMT]]} {
			set NUM [$CONTENT createElement number:currency-style]
			$NUM setAttribute style:name $FNAME
			set DECI [$CONTENT createElement number:number]
			$DECI setAttribute number:decimal-places $AC
			$DECI setAttribute number:min-decimal-places $AC
			$DECI setAttribute number:min-integer-digits 1
			$DECI setAttribute number:grouping true
		
			$NUM appendChild $DECI
			$AUTOSTYLES appendChild $NUM
		    }
		    return $FNAME	   
		}
	
		# return "string"
		# no particular value format for strings
		return "VS"	
	    }
	}
    }
    # -------------------------------------------------------------------------
   
   
    # -------------------------------------------------------------------------
    # $obj Fmt2style $fmtLst
    # -------------------------------------------------------------------------
    # Interprete a list of tcl format to update the variable 'AUTOSTYLES'.  This
    # option recognition is coming after a first path on data, if we specified
    # expressely such a list of format as done with the datatable package.  Now,
    # the tcl formating style (as used in format command) is very rough compared
    # to the style used in ods file. It is just adjusting number of decimal
    # digits or possible minimum length of an integer.
    # -------------------------------------------------------------------------
    method Fmt2style fmtLst {

	set LEN [llength $fmtLst]

	if ![$AUTOSTYLES hasChildNodes] {
	    error "read first the data, to have a first identification of styles"\
		"incorrect use pf private method Fmt2style of class tdomds, Fmt identified only after reading data first " 401
	}

	set N 0
	foreach FMT $fmtLst {
	    incr N
	    # set ceN ce[format %02d $N]
	    set ceN ce$N
	   
	    # position on the style forcasted for this column	   
	    foreach NODE [$AUTOSTYLES childNodes] {
		if {[$NODE @style:name] eq $ceN} {set CSTY $NODE}
	    }
	    # recover former style node
	    set FMT0 [$CSTY @style:data-style-name]	   
	   
	    # retrieve keys parameters of the format
	    regexp {\{?%[ +-0]?(\d*)?.?(\d*)([cdiuefgGs])\}?} $FMT -> LEN DEC TYP
	  
	   
	    switch  $TYP {
		f {
		    # only part, which may change is number of decimal digits
		    set COLFMT($ceN) [join "VFD $DEC M1SS" ""]
		   
		    # if the style has not changed, do nothing
		    if [string equal $COLFMT(ce$N) $FMT0] {
			continue
		    } {
			set NSTY [$CONTENT createElement number:number-style]
			$NSTY setAttribute style:name $COLFMT(ce$N)
			$AUTOSTYLES appendChild $NSTY
			set NODE [$CONTENT createElement number:number]
			$NODE setAttribute number:decimal-places $DEC
			$NODE setAttribute number:min-decimal-places $DEC
			$NODE setAttribute number:min-integer-digits 1
			$NSTY appendChild $NODE

			$CSTY setAttribute style:data-style-name $COLFMT(ce$N)
		    }
		}
		d - i - u {
		    # only part, which may change is length of the integer
		    set COLFMT($ceN) [join "VI $LEN" ""]
		   
		    # if the style has not changed, do nothing
		    if [string equal $COLFMT(ce$N) $FMT0] {
			continue
		    } {
			set NSTY [$CONTENT createElement number:number-style]
			$NSTY setAttribute style:name $COLFMT(ce$N)
			$AUTOSTYLES appendChild $NSTY
			set NODE [$CONTENT createElement number:number]
			$NODE setAttribute number:decimal-places 0
			$NODE setAttribute number:min-decimal-places 0
			$NODE setAttribute number:min-integer-digits $LEN
			$NSTY appendChild $NODE

			$CSTY setAttribute style:data-style-name $COLFMT(ce$N)
		    }
		}	
	    }	
	}
	return $AUTOSTYLES
    }  
    # -------------------------------------------------------------------------  

    # ------------------------------------------------------------------------
    #  $obj autoStyles {asXML/namesList}
    # ------------------------------------------------------------------------
    # retrieve the styles of the data formatted in XML or as a list
    # ------------------------------------------------------------------------
    method autoStyles fmt {
	switch $fmt {
	    "asXML" {return [$AUTOSTYLES asXML]}
	    "asList"  {	
		set RES {}
		foreach N [$AUTOSTYLES childNodes] {
		    lappend RES [$N getAttribute style:name "no-name"]
		}
		return $RES
	    }
	}	   
    }
    # ------------------------------------------------------------------------

    # -------------------------------------------------------------------------
    # $obj RawValue str FMT
    # -------------------------------------------------------------------------
    # In the ods file, we need to also archive the raw value corresponding to
    # a given formatted value:
    #  - float: convert to no internal separation and decimal is the point '.'
    #    starting from the hypothesis included in package amount.
    #  - currency: convert to internal row format for floats.
    #  - date: convert to YYYY-mm-dd
    # The type is given by a code name FMT (see in introductive comments above)
    # -------------------------------------------------------------------------
    method RawValue {str FMT} {

	switch [string index $FMT 1] {
	    "F" {	
		return [string map [subst {" " "" " " "" $TS "" $DS "."}] $str]
	    }
	    "I" {
		set RES [string map {" " ""} $str]
		set RES [string trimleft $RES 0]
		return $RES
	    }
	    "C" {
		switch [string index $FMT end] {
		    "E" {set CUR €}
		    "D" {set CUR $}
		    "F" {set CUR F}
		    "£" {set CUR £}
		}
		return [string map [subst {" " "" " " "" $TS "" $DS "." $CUR ""}] $str]
	    }
	    "D" {
		set date [clock scan $str]
		set RES [clock format $date -format {%Y-%m-%d}]
		return $RES
	    }
	    "T" {	
		set time [clock scan $str]
		set RES [clock format $time -format {%H:%M}]
		return $RES
	    }
	    "R" {
		set date [clock scan $str]
		set RES [clock format $date -format {%Y-%m-%d %H:%M}]
		return $RES
	    }
	    "S" {
		return $str
	    }
	    default {
		# type not recognized
		return $str
	    }
	}
    }
    # -------------------------------------------------------------------------

    # -------------------------------------------------------------------------
    #  $obj ValueType $FMT
    # -------------------------------------------------------------------------
    # Return the office:value-type requested to fill cell according to FMT
    # OpenDocument part1 § 19.385 office:value-type
    # -------------------------------------------------------------------------
    method ValueType FMT {
	switch [string index $FMT 1] {
	    "F" - "I" {return "float"}
	    "C"       {return "currency"}
	    "D" - "R" {return "date"}
	    "T"       {return "time"}	
	    "S"       {return "string"}
	    default   {return "string"}
	}
    }
   
    # -------------------------------------------------------------------------
    # $obj ColAB x
    # -------------------------------------------------------------------------
    # convert a column number to its version in AB as used in spreadsheet
    # column 1 -> A .. 26 -> Z .. 27 -> AA
    # -------------------------------------------------------------------------
    # to recover ASCII number of a char:    scan A %c
    # to transform an interger into a char: format %c 65
    # A -> 65  /  Z -> 90  / 26 letters
    method ColAB x {
	if {$x > 702} {
	    error "Table cannot have more than 702 columns" \
		"limitation reached for method ColAB of class tdomods" 401	   
	}
	if {$x > 26} {
	    set D [expr {($x-1)/26 + 64}]
	    set R [expr {($x -1) % 26 + 65}]
	    return [format %c $D][format %c $R]
	} {
	    return [format %c [expr {($x -1) % 26 + 65}]]
	}
    }
    # foreach I {1 26 27 28 29 30 31 51 52 53 54 55 56} {puts $I=[colNb $I]}
    # -------------------------------------------------------------------------

   
    # -------------------------------------------------------------------------
    #  MAIN METHODS
    # -------------------------------------------------------------------------

   
    # -------------------------------------------------------------------------
    # $obj read $data
    # -------------------------------------------------------------------------
    # read a datatable and store it in the CONTENT dom tree.
    #
    # We don't check if the column style is not alredy existing user shall empty
    # the table (see next methode) before doing a second entry.   
    # -------------------------------------------------------------------------
    method read data {

	if {[llength $data] == 0} {return}

	# --------------------------------------------------------------------
	# define the style of the data rows
	# --------------------------------------------------------------------
	set NODE [$CONTENT createElement style:style]
	$NODE setAttribute \
	    style:name ro2 \
	    style:family table-row \
	    style:parent-style-name Default \
	    style:data-style-name string
	$AUTOSTYLES appendChild $NODE

	set PROP [$CONTENT createElement style:table-row-properties]
	$PROP setAttribute \
	    fo:break-before auto \
	    style:use-optimal-row-height true
	$NODE appendChild $PROP

	# --------------------------------------------------------------------
	# define an empty column definition to be compliant
	# --------------------------------------------------------------------
	set COL [$CONTENT createElement table:table-column]
	$COL setAttribute table:style-name co1
	$TABLE appendChild $COL
	# ::log::log debug "column style created, styles: [my styles namesList]"


	# --------------------------------------------------------------------
        # Get data organized by columns
	# --------------------------------------------------------------------
	# rewrite the command `datatable column count`
	set NBCOL [llength [lindex $data 0]]
	# rewrite the command `datatable transpose $data`
	set LCOL {}
	for {set I 0} {$I < $NBCOL} {incr I} {
	    lappend LCOL [lsearch -all -inline -subindices -index $I $data *]
	}
	# --------------------------------------------------------------------
	# ::log::log debug "columns: $LCOL"

	# --------------------------------------------------------------------
        # Loop in the columns to determine the required formats
	# --------------------------------------------------------------------
	set N 0
	set COLNAMES {}
	foreach COL $LCOL {
	    incr N
	    # set ceN ce[format %02d $N]
	    set ceN ce$N
	    lappend COLNAMES $ceN
	    # recognized the style (update AUTOSTYLES)
	    # ::log::log debug "column $ceN"

	    # define the style attached to the cells of this columns
	    set FMT [my LstStyle $COL]	    	   
	    set COLFMT($ceN) $FMT
	    # ::log::log debug "FMT: $FMT"
	   
	    set NODE [$CONTENT createElement style:style]
	    $NODE setAttribute \
		style:name $ceN \
		style:family table-cell \
		style:parent-style-name Default \
		style:data-style-name $FMT
	    $AUTOSTYLES appendChild $NODE
	   
	    set SUB [$CONTENT createElement style:table-cell-properties]
	    $SUB setAttribute style:text-align-source value-type
	    $NODE appendChild $SUB
	   
	    set SUB [$CONTENT createElement style:paragraph-properties]
	    $SUB setAttribute fo:margin-right 0.2cm
	    $NODE appendChild $SUB	   
	}
	# ::log::log debug "AUTOSTYLES created: [my styles namesList]"

	# --------------------------------------------------------------------
        # Loop in the lines to fill the table
	# --------------------------------------------------------------------
	foreach L $data {
	    # ::log::log debug "Line: $L"
	   
	    set LINE [$CONTENT createElement table:table-row]
	    $LINE setAttribute table:style-name ro2
	    $TABLE appendChild $LINE

	    foreach ELE $L COLNAME $COLNAMES {
	
		set CELL [$CONTENT createElement table:table-cell]
		$CELL setAttribute table:style-name $COLNAME
		set VLTYP [my ValueType $COLFMT($COLNAME)]
		$CELL setAttribute office:value-type $VLTYP
		if {$VLTYP eq "currency"} {
		    $CELL setAttribute office:currency €
		}
		if {$VLTYP eq "date"} {
		    $CELL setAttribute office:date-value [my RawValue $ELE $COLFMT($COLNAME)]
		} {
		    $CELL setAttribute office:value [my RawValue $ELE $COLFMT($COLNAME)]
		}
		$LINE appendChild $CELL
		set NODE [$CONTENT createElement text:p]
		$CELL appendChild $NODE
		set TXT [$CONTENT createTextNode $ELE]
		$NODE appendChild $TXT
	    }
	}
    }
    # -------------------------------------------------------------------------


    # ----------------------------------------------------------------------
    # $obj empty
    # ----------------------------------------------------------------------
    # Remove all the table content, but keep all preliminary information.
    # Object is then ready for a new read command.
    # ----------------------------------------------------------------------
    method empty {} {
	# empty the data
	if [$TABLE hasChildNodes] {
	    foreach N [$TABLE childNodes] {
		$N delete
	    }
	}
	# empty the styles
	if [$AUTOSTYLES hasChildNodes] {
	    foreach N [$AUTOSTYLES childNodes] {
		$N delete
	    }
	}
	unset -nocomplain COLFMT
	return $CONTENT
    }
    # ----------------------------------------------------------------------


    # ----------------------------------------------------------------------
    # Otions for creating files :
    # /(see https://www.tcl.tk/man/tcl8.4/TclCmd/open.html)/
    #  CREAT  : create the file if not exists
    #  WRONLY : open in write only
    #  TRUNC  : truncate the file to 0 if exists (replace)
    #
    # ----------------------------------------------------------------------
    method write filename {

	# ------------------------------------------------------------------
	# create the temporary directories
	# ------------------------------------------------------------------
	# file name without extension
	set FNWOEXT [string range $filename 0 end-[string length [file extension $filename]]]
	set TEMPDIR [file join [file dirname $filename] #$FNWOEXT]
	file mkdir $TEMPDIR
	file mkdir $TEMPDIR/META-INF
	# ------------------------------------------------------------------

	# ------------------------------------------------------------------
	# create file mimetype
	# ------------------------------------------------------------------
	set FH [open $TEMPDIR/mimetype "CREAT WRONLY TRUNC"]
	puts -nonewline $FH $MIMETYPE
	close $FH
	# ------------------------------------------------------------------

	# ------------------------------------------------------------------
	# create file manifest.xml
	# ------------------------------------------------------------------
	regsub -all {>[\s\n]*<} [$MANIFEST asXML] {><} XML_MANIF
	set XML_MANIF "<?xml version=\"1.0\" encoding=\"UTF-8\"?>$XML_MANIF"
	       
	set FH [open $TEMPDIR/META-INF/manifest.xml "CREAT WRONLY TRUNC"]
	puts -nonewline $FH $XML_MANIF
	close $FH
	# ------------------------------------------------------------------

	# ------------------------------------------------------------------
	# create file content.xml
	# ------------------------------------------------------------------
	set XML_CONT [my asXML]
	set $XML_CONT "<?xml version=\"1.0\" encoding=\"UTF-8\"?>$XML_CONT"

	set FH [open $TEMPDIR/content.xml "CREAT WRONLY TRUNC"]
	puts -nonewline $FH  $XML_CONT
	close $FH
	# ------------------------------------------------------------------

	# ------------------------------------------------------------------
	# create the ods file by compressing files and directories
	# ------------------------------------------------------------------
	# zip creation was only tested in Linux and WinXP
	# get zip.exe from http://stahlworks.com/dev/zip.exe
	# option r : recurse into directories
	# option m : move files (delete them in orginal directories)
	# ! : it is important to be in the folder, since we don't want to zip
	#     the directory tempzip, but its content (and subfolder)
	#
	cd $TEMPDIR
	# exec {*}{zip -mr ../data.ods mimetype META-INF/manifest.xml content.xml}
	exec zip -mr ../$filename mimetype META-INF/manifest.xml content.xml
	cd ..

	file delete -force $TEMPDIR
	# ------------------------------------------------------------------
    }


    # -------------------------------------------------------------------------
    #  $obj insertHeaders $lst
    # -------------------------------------------------------------------------
    # insert a line of header
    # -------------------------------------------------------------------------
    method insertHeaders lst {

	set CE0 ce0
	set HEAD [$CONTENT createElement table:table-row]
	$HEAD setAttribute table:style-name $CE0

	# built the line of headers
	foreach C $lst {
	    set CEL [$CONTENT createElement table:table-cell]
	    $CEL setAttribute \
		table:style-name $CE0 \
		office:value-type string
	    $HEAD appendChild $CEL
	    set NODE [$CONTENT createElement text:p]
	    $CEL appendChild $NODE
	    set TXT [$CONTENT createTextNode $C]
	    $NODE appendChild $TXT	    	   
	}

	# insert the line of headers at first place
	set ROWS [$TABLE getElementsByTagName table:table-row]
	set ROW0 [lindex $ROWS 0]
	$TABLE insertBefore $HEAD $ROW0

	# now define the style for the head
	set HSTY [$CONTENT createElement style:style]
	$HSTY setAttribute \
	    style:name $CE0 \
	    style:family table-cell \
	    style:parent-style-name Default
	$AUTOSTYLES appendChild $HSTY

	set NODE [$CONTENT createElement style:table-cell-properties]
	$NODE setAttribute \
	    style:text-align-source fix \
	    fo:background-color #cccccc
	$HSTY appendChild $NODE
	set NODE [$CONTENT createElement style:text-properties]
	$NODE setAttribute fo:font-weight bold
	$HSTY appendChild $NODE
       
	return $HEAD
    }
    # -------------------------------------------------------------------------  

    # -------------------------------------------------------------------------
    # $obj addAutoFilter {}
    # -------------------------------------------------------------------------
    # add an auto filter to the table
    # -------------------------------------------------------------------------
    method addAutoFilter {} {

	# find back the spreadsheet containing the TABLE
        set SHEET [$TABLE parentNode]

	if ![string length [set RANGES [$SHEET getElementsByTagName table:database-ranges]]] {
	    # no chunk for storing ranges, we create one
	    set RANGES [$CONTENT createElement table:database-ranges]
	    # ::log::log debug "SHEET: [$SHEET nodeName]"
	    $SHEET appendChild $RANGES
	}

	set RANGE [$CONTENT createElement table:database-range]
	$RANGE setAttribute table:name data
	# recover number of rows of the table
	set NBROW [llength [$TABLE getElementsByTagName table:table-row]]
	# ::log::log debug "NBROW: $NBROW"
	# recover the number of (contiguous) columns
	set ROW1 [lindex [$TABLE getElementsByTagName table:table-row] 0]
	# ::log::log debug "ROW1: [$ROW1 asList]"
	set NBCOL [llength [$ROW1 getElementsByTagName table:table-cell]]
	# ::log::log debug "NBCOL: $NBCOL"
	set ADDR "[$TABLE @table:name].A1:[$TABLE @table:name].[my ColAB $NBCOL]$NBROW"
	# ::log::log debug "ADDR: $ADDR"
	$RANGE setAttribute table:target-range-address $ADDR
	$RANGE setAttribute table:display-filter-buttons true
	$RANGES appendChild $RANGE

	return $RANGES
    }
    # -------------------------------------------------------------------------  

    # -------------------------------------------------------------------------
    # $obj deleteDatabaseRanges name
    # -------------------------------------------------------------------------
    # Remove a range defined with the given name. If no name given, remove all.
    # This can remove an automatic filter
    # -------------------------------------------------------------------------
    method deleteDatabaseRanges {} {
	# find back the chunk containing ranges or return nothing
	if ![string length [set RANGES [$SHEET getElementsByTagName table:database-ranges]]] {
	    return
	}
	$RANGES delete
    }
    # -------------------------------------------------------------------------     
}
# -----------------------------------------------------------------------------
