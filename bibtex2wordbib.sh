#!/bin/bash

#
# Copyright (c) 2020, msdm
# https://github.com/msdm/bibtex2wordbib
# 
# All rights reserved.
# 
# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:
# 
#     * Redistributions of source code must retain the above copyright notice,
#       this list of conditions and the following disclaimer.
#     * Redistributions in binary form must reproduce the above copyright notice,
#       this list of conditions and the following disclaimer in the documentation
#       and/or other materials provided with the distribution.
#     * Neither the name of rhythmdbsync nor the names of its contributors
#       may be used to endorse or promote products derived from this software
#       without specific prior written permission.
# 
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
# "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
# LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
# A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR
# CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
# EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
# PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
# PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
# LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
# NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
# SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#
#
# bib(la)tex to Microsoft Word Bibliography convertor
#
# This bash script relies on every bibtex field is followed by a comma and 
# newline (,\n). So, make sure your .bib file is formatted properly
# (one field per line)
#
# This code is written based on ECMA-376 standard
# (https://www.ecma-international.org/publications/standards/Ecma-376.htm),
# which defines Microsoft Office XML file format including the bibliography XML.
# If language field is not set for an item, it will be set to English by 
# default.
#
# Usage: bibtex2wordbib.sh <bibtex file> [<output file>]
#
# If output file is not provided, the generated Bibliography XML will be
# printed in the stdout. Otherwise, with writing the generated XML in the file,
# conversion report will be printed in the stdout.

HELP="Usage: bibtex2wordbib.sh <bibtex file> [<output file>]"

if [ $# -lt 1 ]; then
    echo $HELP
    exit 0
elif [ $# -lt 2 ]; then
    OUT="/dev/stdout"
else
    OUT=$2
fi

awk -v out="$OUT" '
function logMsg(str) {
    if(out != "/dev/stdout") print str
}

function echo(str) {
    print str > out
}

function isInList(needle, haystack,     arr) {
    split(haystack, arr, ",")
    for(i in arr) {
        if(arr[i] == needle) return 1
    }
    return 0
}

function stripExtraSpaces(str) {
    gsub(/^[ \t\n]+/, "", str)
    gsub(/[ \t\n]+$/, "", str)
    return str
}

function insertField(type, value) {
    echo("    <b:" type ">" substr(value, 1, 255) "</b:" type ">")
}

function insertAuthor(name,     fs) {
    split(name, fs, ",")
    echo("          <b:Person>")
    if(length(fs) == 1) {
        echo("            <b:Last>" stripExtraSpaces(fs[1]) "</b:Last>")
    } else {
        echo("            <b:Last>" stripExtraSpaces(fs[1]) "</b:Last>")
        echo("            <b:First>" stripExtraSpaces(fs[2]) "</b:First>")
    }
    echo("          </b:Person>")
}

function insertAuthorBlock(type, authors,   i) {
    echo("      <b:" type ">")
    echo("        <b:NameList>")
    for(i in authors) insertAuthor(authors[i])
    echo("        </b:NameList>")
    echo("      </b:" type ">")
}

function parseAuthors(fields,   authors,    editors,    inventors,  translators) {
    if("author" in fields) split(fields["author"], authors, " and ")
    else if("bookauthor" in fields) split(fields["bookauthor"], authors, " and ")
    if("editor" in fields) split(fields["editor"], editors, " and ")
    if("holder" in fields) split(fields["inventor"], inventors, " and ")
    if("translator" in fields) split(fields["translator"], translators, " and ")
    
    if(length(authors) < 1 && length(editors) < 1) return
    
    echo("    <b:Author>")
    if(length(authors) > 0) insertAuthorBlock("Author", authors)
    if(length(editors) > 0) insertAuthorBlock("Editor", editors)
    if(length(inventors) > 0) insertAuthorBlock("Inventor", inventors)
    if(length(translators) > 0) insertAuthorBlock("Translator", translators)
    echo("    </b:Author>")
}

function parseDate(strDate,     ymd) {
    split(strDate, ymd, "-")
    if(length(ymd) > 0) insertField("Year", ymd[1])
    if(length(ymd) > 1) insertField("Month", ymd[2])
    if(length(ymd) > 2) insertField("Day", ymd[3])
}

function parseLang(lang) {

    if(isInList(stripExtraSpaces(lang), "arabic,aeb,ar,ar-IQ,ar-JO,ar-LB,ar-MR,ar-PS,ar-SY,ar-YE,arq,ary,arz,ayl")) insertField("LCID", "1025")
    else if(isInList(stripExtraSpaces(lang), "bulgarian,bg")) insertField("LCID", "1026")
    else if(isInList(stripExtraSpaces(lang), "catalan,ca")) insertField("LCID", "1027")
    else if(isInList(stripExtraSpaces(lang), "czech,cz")) insertField("LCID", "1029")
    else if(isInList(stripExtraSpaces(lang), "danish,da")) insertField("LCID", "1030")
    else if(isInList(stripExtraSpaces(lang), "german,naustrian,ngerman,nswissgerman,swissgerman,de,de-AT,de-CH,de-DE,de-Latf,de-Latf-AT,de-Latf-CH,de-Latf-DE")) insertField("LCID", "1031")
    else if(isInList(stripExtraSpaces(lang), "greek,polutonikogreek,el")) insertField("LCID", "1032")
    else if(isInList(stripExtraSpaces(lang), "english,american,australian,british,canadian,newzealand,en,en-AU,en-CA,en-GB,en-NZ,en-US")) insertField("LCID", "1033")
    else if(isInList(stripExtraSpaces(lang), "spanish,spanishmx,es,es-ES,es-MX")) insertField("LCID", "1034")
    else if(isInList(stripExtraSpaces(lang), "finnish,fi")) insertField("LCID", "1035")
    else if(isInList(stripExtraSpaces(lang), "french,acadien,canadian,fr,fr-CA,fr-CH,fr-FR")) insertField("LCID", "1036")
    else if(isInList(stripExtraSpaces(lang), "hebrew,he")) insertField("LCID", "1037")
    else if(isInList(stripExtraSpaces(lang), "hungarian,magyar,hu")) insertField("LCID", "1038")
    else if(isInList(stripExtraSpaces(lang), "icelandic,is")) insertField("LCID", "1039")
    else if(isInList(stripExtraSpaces(lang), "italian,it")) insertField("LCID", "1040")
    else if(isInList(stripExtraSpaces(lang), "japanese,ja")) insertField("LCID", "1041")
    else if(isInList(stripExtraSpaces(lang), "korean,ko")) insertField("LCID", "1042")
    else if(isInList(stripExtraSpaces(lang), "dutch,nl")) insertField("LCID", "1043")
    else if(isInList(stripExtraSpaces(lang), "norwegian,norsk,nynorsk,nb")) insertField("LCID", "1044")
    else if(isInList(stripExtraSpaces(lang), "polish,pl")) insertField("LCID", "1045")
    else if(isInList(stripExtraSpaces(lang), "portuguese,brazil,portuges,pt,pt-BR,pt-PT")) insertField("LCID", "1046")
    else if(isInList(stripExtraSpaces(lang), "romansh,rm")) insertField("LCID", "1047")
    else if(isInList(stripExtraSpaces(lang), "romanian,ro")) insertField("LCID", "1048")
    else if(isInList(stripExtraSpaces(lang), "russian,ru")) insertField("LCID", "1049")
    else if(isInList(stripExtraSpaces(lang), "croatian,hr")) insertField("LCID", "1050")
    else if(isInList(stripExtraSpaces(lang), "slovak,sk")) insertField("LCID", "1051")
    else if(isInList(stripExtraSpaces(lang), "albanian,sq")) insertField("LCID", "1052")
    else if(isInList(stripExtraSpaces(lang), "swedish,sv")) insertField("LCID", "1053")
    else if(isInList(stripExtraSpaces(lang), "thai,th")) insertField("LCID", "1054")
    else if(isInList(stripExtraSpaces(lang), "turkish,tr")) insertField("LCID", "1055")
    else if(isInList(stripExtraSpaces(lang), "urdu,ur")) insertField("LCID", "1056")
    else if(isInList(stripExtraSpaces(lang), "malay,bahasa,bahasam,id")) insertField("LCID", "1057")
    else if(isInList(stripExtraSpaces(lang), "ukrainian,uk")) insertField("LCID", "1058")
    else if(isInList(stripExtraSpaces(lang), "belarusian,be")) insertField("LCID", "1059")
    else if(isInList(stripExtraSpaces(lang), "slovenian,slovene,sl")) insertField("LCID", "1060")
    else if(isInList(stripExtraSpaces(lang), "estonian,et")) insertField("LCID", "1061")
    else if(isInList(stripExtraSpaces(lang), "latvian,lv")) insertField("LCID", "1062")
    else if(isInList(stripExtraSpaces(lang), "lithuanian,lt")) insertField("LCID", "1063")
    else if(isInList(stripExtraSpaces(lang), "farsi,persian")) insertField("LCID", "1065")
    else if(isInList(stripExtraSpaces(lang), "vietnamese,vi")) insertField("LCID", "1066")
    else if(isInList(stripExtraSpaces(lang), "armenian,hy")) insertField("LCID", "1067")
    else if(isInList(stripExtraSpaces(lang), "basque,eu")) insertField("LCID", "1069")
    else if(isInList(stripExtraSpaces(lang), "macedonian,mk")) insertField("LCID", "1071")
    else if(isInList(stripExtraSpaces(lang), "afrikaans,af")) insertField("LCID", "1078")
    else if(isInList(stripExtraSpaces(lang), "georgian,ka")) insertField("LCID", "1079")
    else if(isInList(stripExtraSpaces(lang), "hindi,hi")) insertField("LCID", "1081")
    else if(isInList(stripExtraSpaces(lang), "sami,samin,se")) insertField("LCID", "1083")
    else if(isInList(stripExtraSpaces(lang), "gaelic,irish,scottish,gd")) insertField("LCID", "1084")
    else if(isInList(stripExtraSpaces(lang), "turkmen,tk")) insertField("LCID", "1090")
    else if(isInList(stripExtraSpaces(lang), "bengali,bn")) insertField("LCID", "1093")
    else if(isInList(stripExtraSpaces(lang), "tamil,ta")) insertField("LCID", "1097")
    else if(isInList(stripExtraSpaces(lang), "telugu,te")) insertField("LCID", "1098")
    else if(isInList(stripExtraSpaces(lang), "malayalam,ml")) insertField("LCID", "1100")
    else if(isInList(stripExtraSpaces(lang), "marathi,mr")) insertField("LCID", "1102")
    else if(isInList(stripExtraSpaces(lang), "sanskrit,sa")) insertField("LCID", "1103")
    else if(isInList(stripExtraSpaces(lang), "mongolian,mn")) insertField("LCID", "1104")
    else if(isInList(stripExtraSpaces(lang), "tibetan,bo")) insertField("LCID", "1105")
    else if(isInList(stripExtraSpaces(lang), "welsh,cy")) insertField("LCID", "1106")
    else if(isInList(stripExtraSpaces(lang), "khmer,km")) insertField("LCID", "1107")
    else if(isInList(stripExtraSpaces(lang), "lao,lo")) insertField("LCID", "1108")
    else if(isInList(stripExtraSpaces(lang), "galician,gl")) insertField("LCID", "1110")
    else if(isInList(stripExtraSpaces(lang), "syriac,syr")) insertField("LCID", "1114")
}

function parseField(entryType, entryId, type, value) {
    if(type == "author") {
        # Do nothing. Author field is handled in a different way.
    } else if(type == "institution") {
        insertField("Institution", value)
    } else if(type == "booktitle") {
        insertField("BookTitle", value)
    } else if(type == "publisher") {
        insertField("Publisher", value)
    } else if(type == "title") {
        insertField("Title", value)
    } else if(type == "booktitle") {
        insertField("BookTitle", value)
    } else if(type == "journaltitle") {
        insertField("JournalName", value)
    } else if(type == "volume") {
        insertField("Volume", value)
    } else if(type == "issue") {
        insertField("Issue", value)
    } else if(type == "volumes") {
        insertField("NumberVolumes", value)
    } else if(type == "edition") {
        insertField("Edition", value)
    } else if(type == "version") {
        insertField("Version", value)
    } else if(type == "pages") {
        insertField("Pages", value)
    } else if(type == "date") {
        parseDate(value)
    } else if(type == "type") {
        if(entryType == "thesis") insertField("ThesisType", value)
        else insertField("type", value)
    } else if(type == "langid") {
        parseLang(value)
    } else if(isInList(type, "isan,isbn,ismn,isrn,issn,iswc")) {
        insertField("StandardNumber", value)
    } else if(type == "shorttitle") {
        insertField("ShortTitle", value)
    } else if(type == "url") {
        insertField("URL", value)
    } else {
        logMsg("Ignoring unsupported field: " type " (" entryType ", " entryId ")")
    }
}

BEGIN {
    RS="@"
    FS=",\n"
    nConvEntries = 0
    echo("<?xml version=\"1.0\" encoding=\"UTF-8\" ?>")
    echo("<b:Sources SelectedStyle=\"\" xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\">")

}

{
    if(NF < 2) next
    #gsub(/[\n\t ]+/, " ", $0)
    split($1, tok, /[{]/)
    if(length(tok) < 2) {
        echo("Malformed entry detected in bibtex file." > "/dev/stderr")
        exit 1
    }

    delete fields
    entryType = stripExtraSpaces(tok[1])
    entryId   = stripExtraSpaces(tok[2])

    for (i=2; i<=NF; i++) {
        split($i, tok, "=")

        if(length(tok) < 2) {
            echo("Malformed field detected in bibtex file." > "/dev/stderr")
            exit 1
        }

        field = stripExtraSpaces(tok[1])
        value = stripExtraSpaces(tok[2])
        gsub(/[{}]/, "", value)
        fields[field] = value
    }
        
    if(isInList(entryType, "article,periodical,suppperiodical,book,suppbook,bookinbook,mvbook,inbook,incollection,suppcollection,collection,"\
      "mbcollection,inproceedings,proceedings,mvproceedings,inproceedings,proceedings,mvproceedings,inreference,reference,mvreference,"\
      "manual,report,thesis,inreference,reference,mvreference,manual,report,thesis,patent,online,booklet,unpublished,misc")) {
        echo("  <b:Source>")
        echo("    <b:Tag>" entryId "</b:Tag>")
    } else {
        logMsg("Ignoring unsupported entry: " entryType ", " entryId)
        next
    }

    nConvEntries += 1
    if(entryType == "article") {
        insertField("SourceType", "JournalArticle")
    } else if(isInList(entryType, "periodical,suppperiodical")) {
        insertField("SourceType", "ArticleInAPeriodical")
    } else if(isInList(entryType, "book,suppbook,bookinbook,mvbook")) {
       insertField("SourceType", "Book")
    } else if(isInList(entryType, "inbook,incollection,suppcollection,collection,mbcollection")) {
        insertField("SourceType", "BookSection")
    } else if(isInList(entryType, "inproceedings,proceedings,mvproceedings")) {
        insertField("SourceType", "ConferenceProceedings")
    } else if(isInList(entryType, "inreference,reference,mvreference,manual,report,thesis")) {
        insertField("SourceType", "Report")
    } else if(entryType == "patent") {
        insertField("SourceType", "Patent")
    } else if(entryType == "online") {
        insertField("SourceType", "InternetSite")
    } else if(isInList(entryType, "booklet,unpublished,misc")) {
        insertField("SourceType", "Misc")
    }
    
    for(field in fields) {
        parseField(entryType, entryId, field, fields[field])
    }
    parseAuthors(fields)
    
    # Set language to English if it is not provided
    if(!("langid" in fields)) insertField("LCID", "1033")
  
    echo("  </b:Source>")
}

END {
    echo("</b:Sources>")
    
    if (nConvEntries == 0) logMsg("\nNo entry was converted.")
    else if (nConvEntries == 1) logMsg("\n" nConvEntries " entry was converted.")
    else logMsg("\n" nConvEntries " entries were converted.")
}
' "$1"
