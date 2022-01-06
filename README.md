Outlook to Yeastar converter
==============================

A Python script to convert Outlook/Exchange/Microsoft 365 Contact .CSVs to Yeastar PBX compatible .CSVs
Convert an exported CSV from the format used in Outlook to one compatible with Yeastar. Rows that cannot be converted
are put into Removed.csv if not defined otherwise.

Getting Started
------------
Place script and exported .CSV from Outlook in an empty directory

From within the directory, run the python script 

`python office_to_yeastar.py -h`
##  Usage 
`office_to_yeastar.py [-h] [-o OUTPUTFILE] [-l LOGFILE] inputfile`

#### positional arguments:
`inputfile `     The csv file that needs to be converted

#### optional arguments:
`-h, --help`     show this help message and exits

`-o OUTPUTFILE`  The csv file to be written with the converted data

`-l LOGFILE`     The csv file to be filled with log data

------------
![Logo](https://github.com/Kautenburger-IT/Kautenburger-IT/raw/main/Logo_Kautenburger-IT.png)
##  License 
https://github.com/Kautenburger-IT/Office-to-Yeastar-Converter/blob/main/LICENSE
