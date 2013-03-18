A console app that will convert a .xls file to XML or CVS.
In XML mode you can say in the arguments what the title will be of each column.

How to use:

C:\XlsToXml.exe help

** HELP **<br>
argument 0: xml,csv<br>
argument 1: input filename<br>
argument 2: page number to process<br>
argument 3: output filename<br>
argument 4 to x: param names in xml file only<br>


arg 0, is one of the two supported formats (maybe JSON later on?)<br>
arg 1, the .xls file you want to parse (if you have xlsx you need to convert to xls)<br>
arg 2, what page to process, 0 - n (yeah one page at the time)<br>
arg 3, output filename ex: out.csv<br>
arg 4, this is the tiles for each column so you can name the node names in the XML. Ex. Id Name Date<br>
