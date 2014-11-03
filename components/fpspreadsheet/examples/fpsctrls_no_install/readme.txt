FPSpreadsheetControls are a set of visual components which facilitate creation
of a spreadsheet application.

TsWorkbookSource is the base of package, it links the visual components to a
worksheet instance.

TsWorkbookTabControl is a TabControl which displays tabs for each worksheet
of the workbook. Changing the active tab selects the corresponding worksheet.

TsWorksheetGrid is a grid component which displays the contents of a worksheet.
It communicates with the TsWorkbookSource by receiving and sending messages 
on the selected cell.

TsCellEdit is a multi-line edit control (memo) for entering for cell values
and formulas. Pressing ENTER transfers the current text into the worksheet.

TsCellIndicator is a simple edit used to display the address of the currently
selected cell. Editing the text allows to jump to the cell address.

TsSpreadsheetInspector is a StringGrid (ValueListEditor, to be precise) which
displays details on the workbook, the selected worksheet, and the selected
cell values and properties.

Linking these controls to a TsWorkbookSource results in a working spreadsheet
appliation without writing any line of code.


The demo application in the folder "fpsctrls_no_install" can be run without
installing the FPSpreadsheet package.