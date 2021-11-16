Outlook to Yeastar converter
==============================

Python script to convert Outlook .CSV to Yeastar compatible .CSV
Convert an exported CSV from the format used in Outlook to one compatible with Yeastar. Rows that cannot be converted
are put into Removed.csv if not defined otherwise.

Getting Started
------------
Place script and exported contacts.csv from Outlook in an empty directory

From within the directory, run the python script 

`python xxx.py -h`
##  Usage 
`contacts.py [-h] [-o OUTPUTFILE] [-l LOGFILE] inputfile`

#### positional arguments:
  `inputfile `     The csv file in Outlook format

#### optional arguments:
  `-h, --help`     show this help message and exit
  `-o OUTPUTFILE`  The csv file in to be written to in Yeastar format
  `-l LOGFILE`     The csv file to be filled with log data
