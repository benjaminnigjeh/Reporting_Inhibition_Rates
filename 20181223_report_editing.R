require(xlsx)

rm(list=ls())
setwd(dir = choose.dir())

# the input file has to be CSV format and no need for a header
Input_DF <- read.csv(file.choose(), header = T)


# all numbers in the data frame is rounded to a single digit
Input_DF <- round(Input_DF, digits = 1)


# number of columns, we'll use this later
cols <- length(Input_DF[1, ]) 


# exporting data.frame to excel is easy with xlsx package
sheetname <- "mysheet"
write.xlsx(Input_DF, "Formated_Report.xlsx", sheetName=sheetname)
file <- "Formated_Report.xlsx"

# we want to highlight cells based on their value
# load workbook
wb <- loadWorkbook(file)              

# Dark Red
fo1 <- Fill(foregroundColor="#8B0000")   
font1 <-  Font(wb, color="whitesmoke", isItalic=FALSE)
a1 <-  Alignment(h="ALIGN_CENTER")
cs1 <- CellStyle(wb, fill=fo1, font = font1, alignment = a1)

# Red
fo2 <- Fill(foregroundColor="#FF0000")    
a2 <-  Alignment(h="ALIGN_CENTER")
cs2 <- CellStyle(wb, fill=fo2, alignment = a2) 

# Orange
fo3 <- Fill(foregroundColor="#FF8C00")    
a3 <-  Alignment(h="ALIGN_CENTER")
cs3 <- CellStyle(wb, fill=fo3, alignment = a3)        

#Yellow
fo4 <- Fill(foregroundColor="#FFFF00")    
a4 <-  Alignment(h="ALIGN_CENTER")
cs4 <- CellStyle(wb, fill=fo4, alignment = a4) 

#Green
fo5 <- Fill(foregroundColor="#00FF00")
a5 <-  Alignment(h="ALIGN_CENTER")
cs5 <- CellStyle(wb, fill=fo5, alignment = a5)         

#Blue
fo6 <- Fill(foregroundColor="#0000FF")    
font2 <-  Font(wb, color="whitesmoke", isItalic=FALSE)
a6 <-  Alignment(h="ALIGN_CENTER")
cs6 <- CellStyle(wb, fill=fo6, font = font2, alignment = a6)         

# get all sheets
sheets <- getSheets(wb)  

# get specific sheet
sheet <- sheets[[sheetname]]

# get rows: 1st row is "not" headers
rows <- getRows(sheet, rowIndex=1:(nrow(Input_DF)+1))     

# get cells: data begins from first column and goes on
cells <- getCells(rows, colIndex = 1:cols+1)          


# extract the cell values
values <- lapply(cells, getCellValue) 

# find cells meeting conditional criteria more than 90%
highlightdarkred <- NULL
for (i in names(values)) {
  x <- as.numeric(values[i])
  if (x > 90 && !is.na(x)) {
    highlightdarkred <- c(highlightdarkred, i)
  }    
}

# find cells meeting conditional criteria more than 75% less than 90%
highlightred <- NULL
for (i in names(values)) {
  x <- as.numeric(values[i])
  if (90 > x && x > 75 && !is.na(x)) {
    highlightred <- c(highlightred, i)
  }    
}

# find cells meeting conditional criteria more than 50% less than 75%
highlightpink <- NULL
for (i in names(values)) {
  x <- as.numeric(values[i])
  if (75 > x && x > 50 && !is.na(x)) {
    highlightpink <- c(highlightpink, i)
  }    
}

# find cells meeting conditional criteria more than 35% less than 50%
highlightyellow <- NULL
for (i in names(values)) {
  x <- as.numeric(values[i])
  if (50 > x && x > 35 && !is.na(x)) {
    highlightyellow <- c(highlightyellow, i)
  }    
}

# find cells meeting conditional criteria more than -200% less than 35%
highlightgreen <- NULL
for (i in names(values)) {
  x <- as.numeric(values[i])
  if (35 > x && x > -200 && !is.na(x)) {
    highlightgreen <- c(highlightgreen, i)
  }    
}

# find cells meeting conditional criteria less than -200%
highlightblue <- NULL
for (i in names(values)) {
  x <- as.numeric(values[i])
  if (x < -200 && !is.na(x)) {
    highlightblue <- c(highlightblue, i)
  }    
}


lapply(names(cells[highlightdarkred]),
       function(ii) setCellStyle(cells[[ii]], cs1))

lapply(names(cells[highlightred]),
       function(ii) setCellStyle(cells[[ii]], cs2))

lapply(names(cells[highlightpink]),
       function(ii) setCellStyle(cells[[ii]], cs3))

lapply(names(cells[highlightyellow]),
       function(ii) setCellStyle(cells[[ii]], cs4))

lapply(names(cells[highlightgreen]),
       function(ii) setCellStyle(cells[[ii]], cs5))

lapply(names(cells[highlightblue]),
       function(ii) setCellStyle(cells[[ii]], cs6))

saveWorkbook(wb, file)
