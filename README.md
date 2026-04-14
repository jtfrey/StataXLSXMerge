# StataXLSXMerge

Python script that merges multiple .xlsx documents produced by Stata's putexcel command into a single optimized .xlsx document

The Stata program has a command, `putexcel …`, that exports data to files in the modern OpenXML format used by Microsoft Excel.  The Stata command is a bit naive in its operation:  redundant font, fill, border, and number format records are exported to the .xlsx file — one for every cell to which a style is applied.  The structure of the `styles.xml` component in the OpenXML format is designed to consolidate unique style elements and have one instance be referenced by the cell styles that use it.  The same approach is used for eliminating duplicate text in cells:  a 'sharedStrings.xml` component contains only the unique text strings, and zero or more cells reference the single instance.

This project had two goals:

1. Can the styles, shared strings, and worksheets from multiple .xlsx documents exported by Stata be consolidated into a single document?
2. Can the process of consolidation produce significantly smaller .xlsx files?

Goal 2 turns out to be of particular concern when the source document(s) have too many font, fill, border, or number format records (thanks to the naive way `putexcel …` works w.r.t. formatting).  While Apple's Numbers app or Google's Sheets app can open a document with e.g. 90k font records, Microsoft Excel will claim the file is corrupted.  Despite that OpenXML standard's stating that style item indices are unsigned 32-bit integers [1][2], anecdotal evidence is such that Excel must implement those ids using unsigned 16-bit integers:  thus, there is a limit of 65535 font records in one document, for example.

## Using the program

Usage is pretty straightfoward:  import one or more .xlsx files, export one .xlsx file.  The script allows for command-line arguments:

```bash
$ ./StataXLSXMerge -h
usage: StataXLSXMerge [-h] [--verbose] [--one-is-okay] --output <output-file-path>
                      <source-file-path> [<source-file-path> ...]

Merge multiple Stata-generated .xslx files

positional arguments:
  <source-file-path>    a .xlsx source file to be merged

optional arguments:
  -h, --help            show this help message and exit
  --verbose, -v         produce additional output as the script runs
  --one-is-okay, -1     go ahead with the merge even if just one source file is provided
  --output <output-file-path>, -o <output-file-path>
                        the .xlsx output file to create (any existing file will be overwritten)
```

Merging seven Stata-produced .xlsx files is accomplished as:

```bash
$ ./StataXLSXMerge -o SummaryReport_All.xlsx SummaryReport.xlsx SummaryReportP{2,3,4,5,6,7}.xlsx

$ ls -l SummaryReport.xlsx SummaryReportP*.xlsx | awk '{s+= $5;}END{printf("%d\n",s);}'
2732476

$ ls -l SummaryReport_All.xlsx| awk '{print $5;}'
940222
```

The seven original files are ca. 2.6 MiB in size; the merged document is just 918 KiB.  The process of merging — and eliminating all those duplicate string and style elements — reduced the file size by about 66%.

### What's with the -1 flag?

The idea of merging a single document into a single document sounds silly at first blush, but recall that merging seven documents managed to decrease the size of the single output file relative to the original files.  So what happens if I tell **StataXLSXMerge** that I'm not crazy and I do want to merge just one document:

```bash
$ ./StataXLSXMerge -1 -o SummaryReport-Optimized.xlsx SummaryReport.xlsx 

$ ls -l SummaryReport.xlsx| awk '{print $5;}'
423526

$ ls -l SummaryReport-Optimized.xlsx| awk '{print $5;}' 
152867
```

Again, the elimination of redundant strings and styles slimmed the file, this time by 64% relative to the original!


[1] https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fonts?view=openxml-3.0.1
[2] https://www.w3.org/TR/xmlschema-2/#unsignedInt
