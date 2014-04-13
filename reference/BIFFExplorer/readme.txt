BIFF Explorer
--------------------------------------------------------------------------------

"BIFF Explorer" is a tool to peek into the internal structure of a binary
Excel file in biff format.

It displays a list of the BIFF records contained in the xls file along with 
name and explanation as described in various documentation files (see "About"). 
Selecting one of the records loads its bytes into a simple hex viewer; for the 
most important records I tried to decipher the contents of the hex values and 
display their meaning in a grid and a memo (page "Analysis"). For the other 
records select a byte in the hex viewer, and the program will display the 
contents of that byte and the following ones as integer, double, string 
(page "Values").

For compiling, note that the program requires the package "VirtualTreeview-new"
from ccr (which in turn requires the package "lclextensions" from
http://code.google.com/p/luipack/downloads/list).
