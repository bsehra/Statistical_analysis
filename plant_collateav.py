#!usr/bin/env python

"""Program to produce average scores across sibling lines from data obtained by analysing expression of YFP in transgenic lines"""
#Written by Bhupinder Sehra; created 8/29/16
#Updated 6/1/17

import sys
import os
import pandas as pd
import numpy as np
import itertools
import openpyxl
import collections

"""
xfin = Excel workbook input with scoring data 
sheetin = worksheet name with data to process
devstagenum = number of developmental stages across which expression was analysed
xfout = Excel workbook to write out average scores
xsheetout = worksheet to write data output
newwb = 'y' for yes if data to be written to a new workbook; 'n' for writing out data to the same workbook as input
"""

xfin = sys.argv[1]
sheetin = str(sys.argv[2])
devstagenum = sys.argv[3]
xfout = sys.argv[4]
xsheetout = sys.argv[5]
newwb = str(sys.argv[6])

def getcol(matrix, i):
    """Returns a column from a 2D array"""
    return [row[i] for row in matrix]

def getavscores(devstagelist, onesiblist, stagecolnum):
    """Returns 
    avscorelist = [] #will contain average scores                                                                                                                                                                     
    templ = []
    collate = []
    zcollate = []
    avs = []
    newlist = []
    denom = 0
    counter = 0
    for stg in devstagelist:
        for srow in onesiblist:
            if str(srow[stagecolnum]) == str(stg):
                collate.append(srow[3:])
                #print "This is srow", srow[3:]
            #avs.append(srow[:2])
        for index in xrange(2):
            avs.append(srow[index])
        avs.append(stg)
        #print avs, "This is avs"
        zcollate = zip(*collate)
            #print "This is zcollate", zcollate
        #print zcollate[0], "This is zcollate[0]"
        #dealing with style expression (y/n values)
        for i in xrange(len(zcollate)):
                #print "row in zcollate", zcollate
            if 'n' in zcollate[i] or 'y' in zcollate[i]:
                if 'y' in zcollate[i]:
                    avs.append("y")
                else:
                    avs.append("n")
            else:
                avs.append(float(sum([fl for fl in zcollate[i]]))/float(len(zcollate[i])))
                #print "This is what was added:", "sum", float(sum([fl for fl in zcollate[i]])), "len",  float(len(zcollate[i])), "sum/len", float(sum([fl for fl in zcollate[i]]))/float(len(zcollate[i]))
        newlist.append(avs)
            #print newlist, "This is newlist"
        avs = []
        collate = []
    return newlist

#declaring variables
df2 = {}
sorteddf = {}
stagelist = []
labels = []
vsetsib = []
usiblist = []
sibrows = []
scorerows = []
finalrows = []
rcol = []
count = []
tissues=[]
stagecount = []
nptissuecount = []
sctrans = []
sctlen = 0
numrows= 0
numcols = 0
sibav =[]
allsibav = []
df = pd.read_excel(xfin, sheetname=None)
df2 = df.get(sheetin)
#print df2
tempdf = {}

"""Main method"""
for i in range(0, int(devstagenum)):  
    stagelist.append(str(df2.iloc[i, 2]))
tissues = df2.columns.values.tolist()
print "These are the tissues assayed", tissues
setsib = set(df2.iloc[:, 1].values.tolist())
setsib = set(filter(lambda x: x == x , setsib))
temp =[]
#collect alike sibling rows in main df
for sib in setsib:
   print "dealing with sib ", sib
   for row in df2.itertuples():
       try:
           if str(row[2]) == "":
               continue
           if row[2] == sib:
               sibrows.append(list(row)[1:])
       except:
           print "What you have cannot be converted to a float value."  
   if len(sibrows) == int(devstagenum):
       tempdf = pd.DataFrame(data=sibrows, columns=tissues)
       tempdf.set_value(0, 'line', sib)
       print "These are the scores for lines with only sibling assayed", tempdf     
   if len(sibrows) > int(devstagenum):
       print "yes more than 7", sibrows
       sibav = getavscores(stagelist, sibrows, 2)
       tempdf = pd.DataFrame(data=sibav, columns=tissues)
       tempdf.set_value(0, 'line', sib)
       print "These are the scores for lines with > 1 plant", tempdf
   allsibav.append(tempdf)
   sibrows = []
   print len(allsibav), "This is len of allsibav"

sorteddf = pd.concat(allsibav, ignore_index=False)

"""Write out to Excel Workbook (.xlsx)"""
if newwb == 'n':
    from openpyxl import load_workbook
    book = load_workbook(xfout)
    writer = pd.ExcelWriter(xfout, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    sorteddf.to_excel(writer, xsheetout)
    writer.save()

if newwb == "y":
    sorteddf.to_excel(xfout, sheet_name=xsheetout)














