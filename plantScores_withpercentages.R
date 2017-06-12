#set working directory which contains Excel workbooks with count data
setwd("C:/Users/bsehra/Documents/Research/Promoter_analysis/Scoring_sheets_transcribed_122316v")
require(XLConnect)

filein <- "scoringsheet_final_1100_1800lines.xlsx" #file with counts
fileout <- "1100_1800_count_stat_nozerobin_area2.xlsx" #file to write out stats

noZero = 1

#load count files
countwb <- loadWorkbook(filein)
#readWorkSheet returns a list of Excel worksheets
cslist = readWorksheet(countwb, sheet = getSheets(countwb))
#getSheets() returns a list of sheet names
csnames <- getSheets(countwb)
print (csnames)

#change these worksheet numbers to worksheets containing counts which are to be compared
sht1=8
sht2=4
sibNums <- c(22, 21)

#Function to produce combined worksheet name derived from worksheets containing count data to be compared
#compiles
#input for getShtName: n1 and n2 are worksheet names
getShtName <- function(n1, n2){
  namev <- c(toString(n1), toString(n2))
  final <- paste(namev, collapse = "_")}

#function to extract counts over 2 different transgenic lines and perform Mann Whitney U Test
#input list of sheets and index of worksheets to compare; wksht1 and wksht2 are integers
#returns a p value despite ties. Includes zero bin counts
testf <- function(wbList, wksht1, wksht2, r, c){
  a = wbList[[wksht1]][r:(r+5), c]
  #cat(a, "This is a\n")
  b = wbList[[wksht2]][r:(r+5), c]
  #cat(b, "This is b\n")
  x = suppressWarnings((wilcox.test(a, b, correct=TRUE)$p.value))
  #cat(x, "\n")
  x
  }

#leave out the zero bin in the test (hence 'r+1')
#r refers to row at which n=0 info is located; c is column with values to use in test
testfWz <- function(wbList, wksht1, wksht2, r, c){
  a = wbList[[wksht1]][(r+1):(r+5), c]
  cat(a, "This is a\n")
  b = wbList[[wksht2]][(r+1):(r+5), c]
  cat(b, "This is b\n")
  x = suppressWarnings((wilcox.test(a, b, correct=TRUE)$p.value))
  x
}
#function to obtain percentage of independant lines with no observable reporter expression per transgenic line
getZeropc <- function(wbList, wksht, s1, zbr, zbc){
  zbinval = wbList[[wksht]][zbr, zbc]
  cat(zbinval, "this is num of lines with no expression")
  zval = (zbinval/s1)*100
  zval
}

#function to obtain percentage of independant lines with observable reporter expression per transgenic line
getnonZeropc <- function(wbList, wksht, s1, zbr, zbc){
  zbinval = wbList[[wksht]][zbr, zbc]
  nzvalpc = ((s1-zbinval)/s1)*100
  zval
}

#Returns a list of developmental stages where expression was analysed
getStageList <- function(endint, increment, wList, sIndex, sCol){
  sList <- list()
  for (n in seq(1, endint, increment)){
    sList <- c(sList, wList[[sIndex]][n,sCol])}
  sList}

#create file to write out results
sheetNames <- list()
foutwb <- loadWorkbook(fileout, create=TRUE)

#sheetnames of count tables to compare statistically
#concatenate with sheetNames first to retain right order
#csnames is a list of sheets in the workbook from above; [sht1], [sht2m] are integers calling sheet in list defined above
sName = getShtName(csnames[sht1], csnames[sht2])
sheetNames <- c(sheetNames, sName)
cat(sName, "this is next sheet name\n")


#Obtain list of tissues across which expression was analysed
numt <- length(colnames(cslist[[sht1]][-c(1,2)]))
print(numt)
tissueList <- colnames(cslist[[sht1]][-c(1,2)])
print(tissueList)
stageList <- list()
#42 = 7 x 6 ie 7 stages, cslist is the list of worksheets, 1st int is the number of the worksheet referenced
# and 2nd int is the column of that worksheet that contains the stages: tested
stageList <- getStageList(42, 6, cslist, sht1, 1)
print (stageList)

#Print stages out every 3 lines: tested
for (name in sheetNames){
  createSheet(foutwb, name = name)
  print(name)
  #write out stages to sheet making room for percentage lines to put in later
  for (s in 1:length(stageList)){
      writeWorksheet(foutwb, data=stageList[s], sheet = name, startRow = (s*3-1), startCol = 1, header=FALSE)
       print(s*3-1)}
    #write out tissues to sheet
  for (t in 1:length(tissueList)){
    writeWorksheet(foutwb, data=tissueList[t], sheet = name, startRow = 1, startCol = (t+1), header=FALSE)
    print(t)}
}

saveWorkbook(foutwb)
foutwb <- loadWorkbook(fileout, create = FALSE)
fslist <- getSheets(foutwb)
print(fslist)

#sibNums contains total sample numbers for transgenic lines to be compared

if (noZero==1){
probs <- list()
for (j in 1:numt){
  probs <- 1
  for (i in seq(1,42,6)){
    y <- testfWz(cslist, sht1, sht2, i, (j+2))
    print (i)
    print (j)
    print (tissueList[[j]])
    probs <- c(probs, y)
    cat(y, "this is prob from mann whitney U")
    writeWorksheet(foutwb, data=y, sheet = sheetNames[[1]], startRow = length(probs), startCol = (j+1), header=FALSE)
    zpl1 <- getZeropc(cslist, sht1, sibNums[[1]], i, (j+2))
    nzpl1 = 100-zpl1
    zpl2 <- getZeropc(cslist, sht2, sibNums[[2]], i, (j+2))
    nzpl2 = 100-zpl2
    cat(nzpl2, "this is nzpl2 for 2082\n")
    cat(zpl2, "This is zpl2 for 2082\n")
    writeWorksheet(foutwb, data=nzpl1, sheet = sheetNames[[1]], startRow = length(probs)+1, startCol = (j+1), header=FALSE)
    writeWorksheet(foutwb, data=nzpl2, sheet = sheetNames[[1]], startRow = length(probs)+2, startCol = (j+1), header=FALSE)
    probs <- c(probs, zpl1, zpl2)
  }
}
saveWorkbook(foutwb)
}

