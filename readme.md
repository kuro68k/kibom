# Kibom
**Pretty BOM generated for Kicad**

Status: **Beta**
Licence: GPL v3

Language: C#, Visual Studio 2013, Mono

This application uses PDFSharp and MigraDoc.

To build with Mono: `xbuild kibom.sln /property:Configuration=Release /property:Platform=x64`


### Generate normalized, readable BOMs for Kicad ###

BOM generation and management can be a bit of a pain. The ideal solution is to have all information contained in the schematic file, and the completed BOM automatically generated from it. The BOM is never edited directly, all changes are made to the schematic and thus the two documents never get out of sync.


Can be used as a plug-in for Kicad. Creates tab separated (TSV) or PDF/RTF BOMs.

- Items grouped by type
- Default parameters for each type (e.g. all caps X7R unless noted)
- Removal of no-fit and no-part items
- Sort by value (including SI units) and designation
- Substitutions file to conver Kicad footprint names to human format
- Custom fields


`Bom_defaults.txt` and bom_subs.txt are used to clean up the output. The defaults file matches designators with names and default types (e.g. R = Resistor, default 1% 0.125W SMD). The subs file is used to do a simple search and replace on footprint strings (e.g. `Resistors_SMD:R_0603` becomes `0603`).

`Custom_fields.txt` lists custom fields that can be added to components in Kicad, and which will then appear on the BOM.

Tab Separated Value (TSV) output can be copy/pasted into Google Sheets and most other spreadsheet programs. PDF output needs headers and footers adding to it. A notes section at the end would also be a good idea.

For XLSX output a header and footer can be specified. See the examples.

### Usage ###

Open the BOM window in Kicad and create a new plugin. For the command line enter:

`"<path to kibom>" "%I" "%O.BOM.txt" <output type>`

For example, to produce a text file you might use:

`"C:\Kibom\kibom.exe" "%I" "%O.BOM.txt" -pretty`

Which will generate `<project name>.BOM.txt` in the Kicad project directory. For Excel output with a template/footer you need to specify the template/footer location:

`"C:\Kibom\kibom.exe" "%I" "%O.BOM.xlsx" -xlsx -t "C:\Kibom\template.xlsx" -f "bom_footer.xlsx" -nfl`

Typically the template is generic (e.g. for your company) and the footer is unique to each project, containing notes and the like.

### Command line arguments ###

`-debug` Generate debug output

`-tsv` Generate tab separated value output

`-pretty` Generate pretty formatted text output

`-pdf` Generate PDF output

`-rtf` Generate RTF output (for LibreOffice/Word etc.)

`-xlsx` Generate Excel output

`-template,-t` Set the template file for Excel output

`-footer,-f` Set the footer file for Excel output

`-no-fit-list,-nfl` Generate a list of no-fit components at the end of the BOM
