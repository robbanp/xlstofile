A console app that will convert a .xls file to XML or CVS.
In XML mode you can say in the arguments what the title will be of each column.

How to use:

C:\XlsToXml.exe help
** HELP **
argument 0: xml,csv
argument 1: input filename
argument 2: page number to process
argument 3: output filename
argument 4 to x: param names in xml file only

arg 0, is one of the two supported formats (maybe JSON later on?)
arg 1, the .xls file you want to parse (if you have xlsx you need to convert to xls)
arg 2, what page to process, 0 - n (yeah one page at the time)
arg 3, output filename ex: out.csv
arg 4, this is the tiles for each column so you can name the node names in the XML. Ex. Id Name Date
