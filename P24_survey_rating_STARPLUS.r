## Run this file in R.
## You may need to install packages: cluster, data.table, tidyverse, openxlsx.


# Define possible script locations for different machines
paths <- c("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO report card - 2024\\Program\\", "C:\\Users\\liqian\\Dropbox (UFL)\\Project_QL\\2022_MCO_Report_Cards\\SFY2023-2024\\MCO Report Card - 2024\\Program\\")

# Loop through the paths and set the working directory if the path exists
for (path in paths) {
  if (file.exists(path)) {
    setwd(path)
    break
  }
}

# Verify that the working directory has been changed
current_wd <- getwd()
print(paste('Current Working Directory:', current_wd))

## Define the custom functions.
source(file.path(current_wd, '2. Admin', 'Program', 'Function', 'RetainReliable.R'))
source(file.path(current_wd, '2. Admin', 'Program', 'Function', 'PercentileRate.R'))


## Read in source file
outdir <- file.path(current_wd, '3. Survey', 'Data', 'Temp_Data')
infile <- "SP24_out.xlsx"
SP24_wb <- loadWorkbook(paste(outdir, infile, sep = "/"))

## merge tabs output by CAHPS macro into single data table
surv_var <- c("ATC", "HWDC", "PDRat", "HPRat")
SP24_surv <- as.data.frame(read.xlsx(SP24_wb), sheet = 1)
# for (invar in 2:length(surv_var)){
#   SP24_surv <- merge(SP24_surv, as.data.frame(read.xlsx(SP24_wb), sheet = invar))
# }

## Omit total and LD (if necessary)
SP24_surv <- subset(SP24_surv, !(SP24_surv$PHI_Plan_Code %in% c("TX", "Texas", "Total", NA)))
for (svar in surv_var){
  SP24_surv[[match(svar, names(SP24_surv))]][which(SP24_surv[match(paste0(svar, "_den"), names(SP24_surv))] < 30)] <- NA
}

## calculate reliability of each score wrt distribution of all scores
SP24_surv <- RetainReliable(SP24_surv, "ATC", 0.6, varlist = "ATC")
SP24_surv <- RetainReliable(SP24_surv, "HWDC", 0.6, varlist = "HWDC")
SP24_surv <- RetainReliable(SP24_surv, "PDRat", 0.6, varlist ="PDRat")
SP24_surv <- RetainReliable(SP24_surv, "HPRat", 0.6, varlist = "HPRat")


## Rate according to percentile band, adjusted for reliability and significance
SP24_surv <- PercentileRate(SP24_surv, varlist = surv_var)


##new output sheet, move to front of workbook
addWorksheet(wb = SP24_wb, sheetName = "SP24_rate")
# worksheetOrder(SP24_wb) <- c(length(names(SP24_wb)), 1:length(surv_var))
writeData(wb = SP24_wb, sheet = "SP24_rate", x = SP24_surv)
saveWorkbook(SP24_wb, paste(outdir, "SP24_out_rate.xlsx", sep = "/"), overwrite = TRUE)

