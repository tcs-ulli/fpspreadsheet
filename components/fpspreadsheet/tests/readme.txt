Tests for fpspreadsheet

spreadtestgui
=============
Lets you quickly run tests in a GUI.
If there are problems, you can open the spreadtestgui.lpr in Lazarus, compile it with debug mode, and trace through the offending test and the fpspreadsheet code it calls.
More details: FPCUnit documentation

spreadtestcli
=============
Command line version of the above, extended with database output. Useful for scripting use (use e.g. --all --format=plain.

For output to an embedded Firebird database, make sure the required dlls/packages are present and run the program, e.g:
spreadtestcli --comment="Hoped to have fixed that string issue" --revision="482"
(the revision is the SVN revision number, so you can keep track of regresssions)

More details: FPCUnit documentation and
https://bitbucket.org/reiniero/testdbwriter

The tests
=========
Basic tests read XLS files and check the retrieved values against a list. This tests whether reading dates, text etc works.

Another test is to take that list of normative values, write it to an xls file, then read back and compare with the original list. This basically tests whether write support is correct.
The files are written to the temp directory. They are deleted on succesful test completion; otherwise they are kept so you can open them up with a spreadsheet application/mso dumper tool/hex editor and check what exactly got written.

Finally, there is a manual test unit: these tests write out cells to a spreadsheet file (testmanual.xls) that the user should inspect himself. Examples are tests for colors, formatting etc.

Adding tests
============
For most tests:
- Add new cells to the A column in the relevant xls files; see comments in files.
- Add corresponding normative/expected value in the relevant test unit; increase array size
- Add your tests that read the data from xls and checks against the norm array.

Note that tests that check for known failures are quite valuable. You can indicate you expect an exception etc. 

Ideas for more tests:
- add more tests to internaltests to explicitly tests fpspreadsheet functions/procedures/properties that don't read/write to xls/xml files
- writing RPN formulas in the manualtests unit
- more xls/xml file formats tested
- more corner cases
- writing all data available to various sheets and reading it back to test whether complex sheets work
- reading faulty files to test exception handling

For more details, please see the FPCUnit documentation.
