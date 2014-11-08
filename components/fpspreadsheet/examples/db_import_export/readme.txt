This example program shows how a large database table can be exported to a 
spreadsheet using virtual mode or fpspreadsheet's fpsexport.
It also shows importing a spreadsheet file into a database using virtual mode.

First, run the section 1 to create a dBase file with random data.
Then, in section 2, the dBase file can be converted to any spreadsheet format
supported. Finally, in section 3, another dBase file can be created from a
selected spreadsheet file.

Export using virtual mode has the advantage that this takes less memory for the
spreadsheet contents, but requires some more coding. It is also quite fast.
Exporting using fpsexport needs less code but takes more memory (important for
large amounts of data) and seems slower.

Please note that this example is mainly educational to show a "real-world"
application of virtual mode, but, strictly speaking, virtual mode would not
be absolutely necessary due to the small number of columns.

