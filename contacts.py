import sys
import csv
import argparse
import re

#Define maximum length of fields
LEN_NUMBER = 31
LEN_FAX = 31
LEN_NAME = 127
LEN_STRING = 255
LEN_EMAIL = 127

#Define Outlook index numbers for required fields
I_VN = 1
I_NN = 3
I_FM = 5
I_EM = 56
I_TG1 = 31
I_TG2 = 32
I_MT1 = 40
I_MT2 = 45
I_TP1 = 37
I_TP2 = 38
I_FG = 30
I_FP = 36
I_OT = [29,33,34,35,42,44]
I_PLZ = 13
I_SG = 8
I_OG = 11
I_RG = 12
I_LRG = 14

#Define number fields
numbers = [31,32,40,45,37,38]

## Define field names for Yeastar format
ofieldnames = ['First Name','Last Name','Company','Email','Work Number','Work Number 2','Mobile','Mobile 2','Home','Home 2','Work Fax','Home Fax','Other','ZIP Code','Street','City','State','Country']

## Read a value formatted as a string
def readString(row,index,length):
## Replace invalid characters for strings
    string = row[index].replace("\n"," ")
    string = string.replace("Ä","Ae")
    string = string.replace("Ö","Oe")
    string = string.replace("Ü","Ue")
    string = string.replace("ä","ae")
    string = string.replace("ö","oe")
    string = string.replace("ü","ue")
    string = string.replace("ẞ","SS")
    string = string.replace("ß","ss")
## Shorten string on maximum length if necessary
    if(len(string) > length):
        return string[:length]
    return string

## Read a value formatted as a name
def readName(row,index,length):
    string = readString(row,index,length)
## replace invalid characters: !%.@:&"'\<>`$
    string = re.sub("[!%\.@:&\"\'\\<>`$]","",string)
    return string

## Read a value formatted as a company
def readCompany(row,index,length):
    string = readString(row,index,length)
## replace invalid characters: &"'\<>
    string = string.replace("&","+")
    string = re.sub("[\"\'\\<>]","",string)
    return string

## Read a value formatted as an email. Invalid entries will be deleted
def readEmail(row,index):
    string = row[index]
## Remove spaces
    string = re.sub("\s","",string)
## Check invalid characters: #,[]=&"'\<>`$
    if re.match("[#,\[\]=&\"\'\\<>`$]",string):
        return ""
## Check Email length
    if(len(string) > LEN_EMAIL):
        return ""
## Check Email format
    if re.match("^[\w\.+^\?*\-!%/{}~|]+@[^@]+\.[^@]+",string):
        return string
    else:
        return ""

## Read a value formatted as a number
def readNumber(row,index,length):
    string = row[index]
## Check if number is too long
    if(len(string) > length):
        return ""
## Remove spaces
    string = re.sub("\s","",string)
## Check number format
    if re.match("[^\d\s\(\)\.\-\+]",string):
        return ""
    else:
        return string

## Read other numbers
def readOther(row,indexlist):
    rvalue = []
    for f in indexlist:
        n = readNumber(row,f,LEN_NUMBER)
        if n:
            rvalue.append(n)
    if len(rvalue) == 0:
        return ""
    elif len(rvalue) == 1:
        return rvalue[0]
    else:
        return rvalue

## Write row into file of removed entries
def writeLog(file,values,message):
    row = ""
    row = row + message + "Full Row Data:\n"
    for v in values:
        row = row + str(v).replace("\n","") + ","
    file.write(str(row)+"\n")

## Write row into outputfile
def writeRow(writer,r,log):
    othernumber = readOther(r,I_OT)
    if type(othernumber) == list:
        writeLog(log,r,"\nLost Data - Too many numbers\n"+ str(othernumber[1:]))
        othernumber = othernumber[0]
    writer.writerow({'First Name':readName(r,I_VN,LEN_NAME),
                     'Last Name':readName(r,I_NN,LEN_NAME),
                     'Company':readCompany(r,I_FM,LEN_STRING),
                     'Email':readEmail(r,I_EM),
                     'Work Number':readNumber(r,I_TG1,LEN_NUMBER),
                     'Work Number 2':readNumber(r,I_TG2,LEN_NUMBER),
                     'Mobile':readNumber(r,I_MT1,LEN_NUMBER),
                     'Mobile 2':readNumber(r,I_MT2,LEN_NUMBER),
                     'Home':readNumber(r,I_TP1,LEN_NUMBER),
                     'Home 2':readNumber(r,I_TP2,LEN_NUMBER),
                     'Work Fax':readNumber(r,I_FG,LEN_FAX),
                     'Home Fax':readNumber(r,I_FP,LEN_FAX),
                     'Other':readOther(r,I_OT),
                     'ZIP Code':readString(r,I_PLZ,LEN_STRING),
                     'Street':readString(r,I_SG,LEN_STRING),
                     'City':readString(r,I_OG,LEN_STRING),
                     'State':readString(r,I_RG,LEN_STRING),
                     'Country':readString(r,I_LRG,LEN_STRING)})

## Convert file from Outlook to Yeastar format
def convert(inputfile,outputfile,logfile):
    with inputfile as infile, outputfile as output, logfile as log:
        oWriter = csv.DictWriter(output,ofieldnames)
        ## Write Yeastar Header into output file
        for s in ofieldnames:
            output.write(s)
            if s != ofieldnames[len(ofieldnames)-1]:
                output.write(",")
            else:
                output.write("\n")
        ## Write rows into correct files
        seen = []
        for r in csv.DictReader(infile):
            dicts = list(r.items())
            values = []
            for e in dicts:
                values.append(e[1])
            ## Remove rows without any name, put Last Name into First Name if necessary
            if values[I_VN] == "" and values[I_NN] == "" :
                writeLog(log,values,"\nRemoved - Missing Names\n")
                continue
            elif values[I_VN] == "":
                values[I_VN] = values[I_NN]
                values[I_NN] = ""
            ## Remove rows without atleast one necessary number
            if any(readNumber(values,s,LEN_NUMBER) for s in numbers) or readOther(values,I_OT):
                ## Remove duplicate rows
                (v,n) = (readName(values,I_VN,LEN_NAME),readName(values,I_NN,LEN_NAME))
                if (v,n) in seen:
                    writeLog(log,values,"\nRemoved - Duplicate Name\n")
                    continue
                writeRow(oWriter,values,log)
                seen.append((v,n))              
            else:
                writeLog(log,values,"\nRemoved - No valid entries in necessary number fields\n")  

## Get filenames and call convert function
class C():
    pass
c = C()
parser = argparse.ArgumentParser(description='Convert an exported CSV from the format used in Outlook to one compatible with Yeastar. Rows that cannot be converted are put into Removed.csv if not defined otherwise')
parser.add_argument('inputfile',type=argparse.FileType('r',encoding="utf-8"),help='The csv file in Outlook format')
parser.add_argument('-o',default='Output.csv',type=argparse.FileType('w',encoding='utf-8-sig'),help='The csv file in to be written to in Yeastar format',dest='outputfile')
parser.add_argument('-l',default='Log.txt',type=argparse.FileType('w',encoding='utf-8'),help='The csv file to be filled with log data',dest='logfile')
args = parser.parse_args(namespace=c)
if(c.inputfile == ""):
    c.inputfile = None
    sys.exit(1)
if(c.outputfile == ""):
    c.outputfile = None
    sys.exit(1)
if(c.logfile == ""):
    c.logfile = None
    sys.exit(1)
convert(c.inputfile,c.outputfile,c.logfile)
sys.exit(0)