# bibtex2wordbib

## bib(la)tex to Microsoft Word Bibliography convertor

This bash script relies on every bibtex field is followed by a comma and newline (,\n). So, make sure your .bib file is formatted properly (one field per line)

This code is written based on [ECMA-376 standard](https://www.ecma-international.org/publications/standards/Ecma-376.htm), which defines Microsoft Office XML file format including the bibliography XML. If language field is not set for an item, it will be set to English by default.

## Usage
`bibtex2wordbib.sh <bibtex file> [<output file>]`

If output file is not provided, the generated Bibliography XML will be printed in the stdout. Otherwise, with writing the generated XML in the file, conversion report will be printed in the stdout.
