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
infile <- "SC24_out.xlsx"
SC24_wb <- loadWorkbook(paste(outdir, infile, sep = "/"))

## merge tabs output by CAHPS macro into single data table
surv_var <- c("GCQ", "HWDC", "PDRat", "HPRat")
SC24_surv <- as.data.frame(read.xlsx(SC24_wb), sheet = 1)
# for (invar in 2:length(surv_var)){
#   SC24_surv <- merge(SC24_surv, as.data.frame(read.xlsx(SC24_wb), sheet = invar))
# }

## Omit total and LD (if necessary)
SC24_surv <- subset(SC24_surv, !(SC24_surv$PHI_PLAN_CODE %in% c("TX", "Texas", "Total", NA)))
for (svar in surv_var){
  SC24_surv[[match(svar, names(SC24_surv))]][which(SC24_surv[match(paste0(svar, "_den"), names(SC24_surv))] < 30)] <- NA
}


## calculate reliability of each score wrt distribution of all scores
SC24_surv <- RetainReliable(SC24_surv, "GCQ", 0.6, varlist = "GCQ")
SC24_surv <- RetainReliable(SC24_surv, "HWDC", 0.6, varlist = "HWDC")
SC24_surv <- RetainReliable(SC24_surv, "PDRat", 0.6, varlist ="PDRat")
SC24_surv <- RetainReliable(SC24_surv, "HPRat", 0.6, varlist = "HPRat")


## Rate according to percentile band, adjusted for reliability and significance
SC24_surv <- PercentileRate(SC24_surv, varlist = surv_var)


##new output sheet, move to front of workbook
addWorksheet(wb = SC24_wb, sheetName = "SC24_rate")
# worksheetOrder(SC24_wb) <- c(length(names(SC24_wb)), 1:length(surv_var))
writeData(wb = SC24_wb, sheet = "SC24_rate", x = SC24_surv)
saveWorkbook(SC24_wb, paste(outdir, "SC24_out_rate.xlsx", sep = "/"), overwrite = TRUE)