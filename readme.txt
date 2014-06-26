fpspreadsheet
=============
The fpSpreadsheet library offers a convenient way to generate and read 
spreadsheet documents in various formats:
- Excel 2.x .xls
- Excel 5.0/Excel 95 .xls
- Excel 8.0 (Excel 97-XP) .xls
- Microsoft OOXML .xlsx
- LibreOffice/OpenOffice OpenDocument .ods
- wikimedia wikitable formats

The library is written in a very flexible manner, capable of being extended to 
support any number of formats easily.

Installation
============
If you only need non-GUI components: in Lazarus: 
- Package/Open Package File 
- select laz_fpspreadsheet.lpk
- click Compile. 
Now the package is known to Lazarus (and should e.g. show up in Package/Package Links). 
Add it to your project like you add other packages.

If you also want GUI components (grid and chart): 
- Package/Open Package File
- seleect laz_fpspreadsheet_visual.lpk
- click Compile
- then click Use, Install and follow the prompts to rebuild Lazarus with the new package.
Drop needed grid/chart components on your forms as usual
		
License
=======
LGPL with static linking exception. This is the same license as is used in the Lazarus Component Library. 

More information
================
FPSpreadsheet documentation in fpspreadsheet.chm (open e.g. with Lazarus lhelp)

The fpspreadsheet article on the Lazarus wiki with lots of example:
http://wiki.lazarus.freepascal.org/FPSpreadsheet

The demo programs in the examples folder
