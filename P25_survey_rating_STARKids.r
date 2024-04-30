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
infile <- "SK24_out.xlsx"
SK24_wb <- loadWorkbook(paste(outdir, infile, sep = "/"))

## merge tabs output by CAHPS macro into single data table
surv_var <- c("HPRat", "AtC", "SpecTher", "APM", "coord", "GNI", "transit", "BHCoun")
SK24_surv <- as.data.frame(read.xlsx(SK24_wb), sheet = 1)
# for (invar in 2:length(surv_var)){
#   SK24_surv <- merge(SK24_surv, as.data.frame(read.xlsx(SK24_wb), sheet = invar))
# }

# read in eligible survey population
# SK24_pop <- as.data.frame(read.xlsx(paste(outdir, "SK24_weights.xlsx", sep = "/"), sheet = 5)) %>% select(PHI_Plan_Code = PLAN_CD, sample_pool_members)
# SK24_surv <- merge(SK24_surv, SK24_pop)
# rm(SK24_pop)

## Omit total and LD (if necessary)
SK24_surv <- subset(SK24_surv, !(SK24_surv$PHI_Plan_Code %in% c("TX", "Texas", "Total", NA)))
for (svar in surv_var){
  SK24_surv[[match(svar, names(SK24_surv))]][which(SK24_surv[match(paste0(svar, "_den"), names(SK24_surv))] < 30)] <- NA
}

  
## calculate reliability of each score wrt distribution of all scores
SK24_surv <- RetainReliable(SK24_surv, "HPRat", 0.6, varlist = "HPRat")
SK24_surv <- RetainReliable(SK24_surv, "AtC", 0.6, varlist = "AtC")
SK24_surv <- RetainReliable(SK24_surv, "SpecTher", 0.6, varlist = "SpecTher")
SK24_surv <- RetainReliable(SK24_surv, "APM", 0.6, varlist ="APM")
SK24_surv <- RetainReliable(SK24_surv, "coord", 0.6, varlist ="coord")
SK24_surv <- RetainReliable(SK24_surv, "GNI", 0.6, varlist ="GNI")
SK24_surv <- RetainReliable(SK24_surv, "transit", 0.6, varlist ="transit")
SK24_surv <- RetainReliable(SK24_surv, "BHCoun", 0.6, varlist ="BHCoun")


## Rate according to percentile band, adjusted for reliability and significance
SK24_surv <- PercentileRate(SK24_surv, varlist = surv_var)


##new output sheet, move to front of workbook
addWorksheet(wb = SK24_wb, sheetName = "SK24_rate")
# worksheetOrder(SK24_wb) <- c(length(names(SK24_wb)), 1:length(surv_var))
writeData(wb = SK24_wb, sheet = "SK24_rate", x = SK24_surv)
saveWorkbook(SK24_wb, paste(outdir, "SK24_out_rate.xlsx", sep = "/"), overwrite = TRUE)
