#!usr/bin/env python

"""Program to process scoring data"""
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

#numtissues = sys.argv[4]

def getcol(matrix, i):
    """Return a column from a 2d array"""
    return [row[i] for row in matrix]

def bin_count(sclist, fineq1, fineq2):
    """Sorts an input list of values and returns a value corresponding to a bin"""
    if fineq1 == 0.0 and fineq2 == 0.0:
        return sum(1 for x in sclist if float(x) == fineq1)
    if fineq1 == 0.0 and fineq2 > 0.0:
        return sum(1 for x in sclist if float(x) > fineq1  and float(x) <= fineq2)
    else:
        return sum(1 for x in sclist if float(x) > fineq1  and float(x) <= fineq2)

def make_float(listnum):
    """Return float value from input element (normally int)"""
    return [float(element) for element in listnum]

#read in all sheets as dict
#declaring variables
df = {}
df2 = {}
df3 = {}
labels = []
stagelist = []
tissues = []
bins = []
scores = []
rcol = []
count = []
tissuecount=[]
stagecount = []
nptissuecount = []
sctrans = []
sctlen = 0
scoredict = {}
numrows= 0
numcols = 0

"""Main method"""
df = pd.read_excel(xfin, sheetname=None)
df2 = df.get(sheetin)
df2 = df2.drop("Style_ap_ring_(y/n)",  1)
tissues = df2.columns.values
print "these are df2 col vals", tissues, len(tissues)
#for i in range(0, devstagenum):
    #labels.append(df2.ix[:devstagenum-1,[2]].values
for i in range(0, int(devstagenum)):  
    stagelist.append(str(df2.iloc[i, 2]))
bins = ["n=0","0<n=<1", "1<n=<2", "2<n=<3", "3<n=<4", "4<n=<5"]
numrows = len(df2.index)
numcols = len(df2.transpose().index)
for s in stagelist:
    print "dealing with stage ", s
    for row in df2.itertuples():
        if str(row[3]) == str(s):
            scores.append(row)
            #print [x for x in scores]
    sctrans = map(list, itertools.izip_longest(*scores, fillvalue='-'))
    print "this is sctrans", [y for y in sctrans]
    scores = []
   #iterate over sctran from rows 4 to len sctrans
    for num in range(4, len(sctrans)):
        try:
            currcol = make_float(sctrans[num])
            #print "This is currcol", currcol
        except:
            print "This list is full of strings"
       #perform count function over this and all following rows. Count is a list.
       #index 0 (i <1); [1] (1<=n <2) etc all the way up to 5
        count.append(bin_count(currcol, float(0.0), float(0.0)))
        count.append(bin_count(currcol, float(0.0), float(1.0)))
        count.append(bin_count(currcol, float(1.0), float(2.0)))
        count.append(bin_count(currcol, float(2.0), float(3.0)))
        count.append(bin_count(currcol, float(3.0), float(4.0)))
        count.append(bin_count(currcol, float(4.0), float(5.0)))
        #5 scores gen/tissue
        tissuecount.append(count)
        nptissuecount=np.array(tissuecount)
        count = []
    scoredict[s] = [tissuecount]
    tissuecount = []
#Create dataframe of counts to write out to excel file
dataout = [] 
dataflist = []

for stage in stagelist:
    dataout = (scoredict.get(stage)[0])
    print stage
    print "This is what the matrix dataout looks like"
    print dataout, "raw"
    #print [row for row in dataout], "This is dataout"
    #print [row[0] for row in dataout], "This is dataout [0]"
    #print [row[0][0] for row in dataout], "This is dataout [0][0]"
    print len(dataout)
    countdf = pd.DataFrame(data=dataout, index=tissues[3:], columns=bins)
    print countdf
    countdf2 = countdf.transpose()
    #print countdf2
    dataflist.append(countdf2)
concatdf = pd.concat(dataflist, keys=stagelist)
print concatdf

#write out to Excel workbook
if newwb == 'n':
    from openpyxl import load_workbook
    book = load_workbook(xfout)
    writer = pd.ExcelWriter(xfout, engine='openpyxl') 
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    concatdf.to_excel(writer, xsheetout)
    writer.save()

if newwb == "y":
    concatdf.to_excel(xfout, sheet_name=xsheetout)      












