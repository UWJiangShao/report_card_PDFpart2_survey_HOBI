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
infile <- "SA24_out.xlsx"
SA24_wb <- loadWorkbook(paste(outdir, infile, sep = "/"))

## merge tabs output by CAHPS macro into single data table
surv_var <- c("AtC", "HWDC", "PDRat", "HPRat")
SA24_surv <- as.data.frame(read.xlsx(SA24_wb), sheet = 1)
# for (invar in 2:length(surv_var)){
#   SA24_surv <- merge(SA24_surv, as.data.frame(read.xlsx(paste(outdir, infile, sep = "/")), sheet = invar))
# }

## Omit total and LD (if necessary)
SA24_surv <- subset(SA24_surv, !(SA24_surv$PHI_Plan_Code %in% c("TX", "Texas", "Total", NA)))
for (svar in surv_var){
  SA24_surv[[match(svar, names(SA24_surv))]][which(SA24_surv[match(paste0(svar, "_den"), names(SA24_surv))] < 30)] <- NA
}


## calculate reliability of each score wrt distribution of all scores
SA24_surv <- RetainReliable(SA24_surv, "AtC", 0.6, varlist = "AtC")
SA24_surv <- RetainReliable(SA24_surv, "HWDC", 0.6, varlist = "HWDC")
SA24_surv <- RetainReliable(SA24_surv, "PDRat", 0.6, varlist ="PDRat")
SA24_surv <- RetainReliable(SA24_surv, "HPRat", 0.6, varlist = "HPRat")


## Rate according to percentile band, adjusted for reliability and significance
SA24_surv <- PercentileRate(SA24_surv, varlist = surv_var)


##new output sheet, move to front of workbook
addWorksheet(wb = SA24_wb, sheetName = "SA24_rate")
# worksheetOrder(SA24_wb) <- c(length(names(SA24_wb)), 1:length(surv_var))
writeData(wb = SA24_wb, sheet = "SA24_rate", x = SA24_surv)
saveWorkbook(SA24_wb, paste(outdir, "SA24_out_rate.xlsx", sep = "/"), overwrite = TRUE)