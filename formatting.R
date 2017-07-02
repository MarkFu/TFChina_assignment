#    formatting requests:
#    Add an extra row between months

args = commandArgs(T)
filename = args[1]

Required_Packages <- c("xlsx", "xtable", "zoo", "lubridate", "ggplot2", "gridExtra", "gdata", "brew") #"XLConnet"
Remaining_Packages <- Required_Packages[!(Required_Packages %in% installed.packages()[,"Package"])];
if (length(Remaining_Packages)) install.packages(Remaining_Packages)
for(package_name in Required_Packages) suppressMessages(library(package_name,character.only=TRUE,quietly=TRUE))

prune <- function(var, total = T, to.keep = c("Date", "Description", "Category", "Amount", "Account Balance")) {
  result <- subset(var, select = to.keep)
  result$Date <- as.character(result$Date)
  if (nrow(result) && total) result <- rbind(result, c("Total", sum(result$Amount), 0, 0, 0))
  return(result)
}

rename_single = function(data, oldname, newname){
  
  colnames(data)[which(colnames(data) == oldname)] = newname
  #print(names(data))
  return(data)
}

asNumeric <- function(x) as.numeric(as.character(x))
adj <- function(x) ifelse(!is.na(x), format(round(x, 2), nsmall = 2, big.mark = ","), "N/A")

####################################################################

#setwd("/Users/shuoxie/Dropbox/bank_statement/Yodlee")

sheetNum = sheetCount(filename)

to.keep = c("account_balance", "description", "post_date", "transactionBaseType", "category", "amount", "account_number", "status")
to.write = c("Status", "Date", "Description", "Type", "Amount", "Category", "Account Balance")

all_accounts = list()
account_numbers = c()

for (i in seq_len(sheetNum)) {
  # data = c()
  # data <- rbind(data, read.xlsx2(filename, sheetIndex = i))
  # acccount_list <- levels(data$account_number)
  
  data <- read.xlsx2(filename, sheetIndex = i)
  account_numbers = c(account_numbers, as.character(data$account_number[1]))
  
  data = subset(data, select = to.keep)
  
  ## remove pending transactions
  pending.index = which(data$status == "pending")
  
  data = rename_single(data, "transactionBaseType", "transaction_type")
  data$post_date <- as.Date(substr(data$post_date, 1, 10))
  data$account_balance <- asNumeric(data$account_balance)
  balance.at.import = mean(asNumeric(data$account_balance))
  data$amount <- asNumeric(data$amount)
  
  ##  use the following code to exclude pending transactions when calculating running balances  
  #   if(length(pending.index)) {
  #     pending.transactions = data[pending.index, ]
  #     data = data[-pending.index, ]
  #   }
  #   
  # backfill running account balance
  data = data[rev(order(data[, c("post_date")], data$description, data$amount)), ]
  sub_data = data[, c("post_date", "amount", "description")]
  
  if(nrow(sub_data) > 1){
    duplicate_idx = which(sapply(seq_len(nrow(sub_data) - 1), function(row) all(sub_data[row, ] == sub_data[row + 1, ])))
    if (length(duplicate_idx)) data <- data[-duplicate_idx, ]
    
    amount_sign <- 2 * as.numeric(data[, "transaction_type"] == "debit") - 1
    amount_flow <- data$amount * amount_sign
    balance_flow <- cumsum(amount_flow)
    
    balance_flow <- balance_flow + balance.at.import
    balance_flow = c(balance.at.import, balance_flow[-length(balance_flow)])
    
    data$account_balance = balance_flow 
  }
  
  # format dollar amount
  for(dollar_var in c("amount", "account_balance")){
    data[ , which(colnames(data) == dollar_var)] = adj(data[ , which(colnames(data) == dollar_var)])
  }
  
  data = data[rev(order(data[, c("post_date")], data$description, data$amount)), ]
  
  ##  use the following code to exclude pending transactions when calculating running balances    
  #   if(length(pending.index)) {
  #     pending.transactions$account_balance = "/"
  #     data = rbind(pending.transactions, data)
  #   }
  
  # change header
  data = rename_single(data, "post_date", "Date")
  data = rename_single(data, "description", "Description")
  data = rename_single(data, "amount", "Amount")
  data = rename_single(data, "transaction_type", "Type")
  data = rename_single(data, "category", "Category")
  data = rename_single(data, "account_balance", "Account Balance")
  data = rename_single(data, "status", "Status")
  
  data = subset(data, select = to.write)
  all_accounts[[i]] = data
}

output = xlsx::createWorkbook(type="xlsx")

# Title and sub title styles
TITLE_STYLE <- CellStyle(output)+ Font(output, heightInPoints = 16, 
                                       color = "black", isBold = TRUE, underline = 1)

HIGHLIGHT_CELL_STYLE <- CellStyle(output) + Font(output, heightInPoints = 16, isBold = F, isItalic = F, color = "red")
#+ Fill(backgroundColor="lavender", foregroundColor="lavender", pattern="SOLID_FOREGROUND") + Alignment(h="ALIGN_RIGHT")

TABLE_ROWNAMES_STYLE <- CellStyle(output) + Font(output, isBold=TRUE)
TABLE_COLNAMES_STYLE <- CellStyle(output) + Font(output, isBold=TRUE) +
  Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
  Border(color="black", position=c("TOP", "BOTTOM"), 
         pen=c("BORDER_THIN", "BORDER_THICK")) 

xlsx.addTitle<-function(sheet, rowIndex, title, titleStyle){
  rows <-createRow(sheet,rowIndex=rowIndex)
  sheetTitle <-createCell(rows, colIndex=1)
  setCellValue(sheetTitle[[1,1]], title)
  setCellStyle(sheetTitle[[1,1]], titleStyle)
}

#xlsx.addTitle(sheet, rowIndex = 1, title = filename, titleStyle = TITLE_STYLE)

for(i in seq_len(length(all_accounts))){
  sheet = xlsx::createSheet(output, sheetName = gsub("[[:punct:]]", "x", account_numbers[i]))
  
  xlsx::addDataFrame(all_accounts[[i]], sheet, startRow = 1, startColumn=1, 
                     colnamesStyle = TABLE_COLNAMES_STYLE,
                     row.names = F )
  
  #createFreezePane(sheet, 1, 1, 2, 8)
  xlsx::autoSizeColumn(sheet, colIndex = c(1:ncol(all_accounts[[i]])))
  #setColumnWidth(sheet, colIndex=c(1:ncol(all_accounts[[i]])), colWidth = 11)
  
  rowIndex <- which(all_accounts[[i]]$`Account Balance` < 0)
  colIndex <- which(names(all_accounts[[i]]) == "Account Balance")
  #setCellStyle(output, sheet = account_numbers[i], row = rowIndex, col = colIndex, cellstyle = HIGHLIGHT_CELL_STYLE) 
}

xlsx::saveWorkbook(output, paste("formatted_", filename, sep = ""))

#################################
### Add highlighting
#################################

# highlight balance amount if < 0      
wb = loadWorkbook(paste("formatted_", filename, sep = ""))  # load workbook
# create cell style
fill_object = Fill(foregroundColor="yellow")             
highlight_style = CellStyle(wb, fill = fill_object)

sheets = getSheets(wb)                               # get all sheets
Max.Row = 3000

# loop over sheets
for (sh in names(sheets)) {
  sheet = sheets[[sh]]   
  cols = 6
  rows <- getRows(sheet, rowIndex = 1:Max.Row)     # get rows
  cells <- getCells(rows, colIndex = cols)    # get cells
  # change start column if loadings do not start in second column of excel
  values <- lapply(cells, getCellValue)             # extract the values
  # find cells meeting conditional criteria 
  highlight <- "test"
  for (i in names(values)) {
    x <- (values[i])
    if (length(grep("-", x))>0 & !is.na(x)) {
      highlight <- c(highlight, i)
    }    
  }
  highlight <- highlight[-1]
  # apply style to cells that meet criteria
  lapply(names(cells[highlight]),
         function(ii) setCellStyle(cells[[ii]], highlight_style))
}

# save
saveWorkbook(wb, paste("formatted_", filename, sep = "")) 
print("Formatted Yodlee data file generated.")
