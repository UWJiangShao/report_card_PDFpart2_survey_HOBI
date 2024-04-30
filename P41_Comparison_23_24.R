library(dplyr)
library(openxlsx)

# Read in the data
# Report card 2022
SK24 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\3. Survey\\Data\\Temp_Data\\SK24_out_rate.xlsx", sheet = 2)
SK23 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\3. Survey\\Data\\Temp_Data\\SK23_out.xlsx", sheet = 2)


# Only keep certain dataset:
measures <- c("HPRat", "AtC", "SpecTher", "APM", "coord", "GNI", "transit", "BHCoun")

temp <- SK24[,measures]
SK24 <- SK24[grep("(_den$|_rat$|_Code$)", names(SK24))]
SK24 <- cbind(SK24, temp)

temp <- SK23[,measures]
SK23 <- SK23[grep("(_den$|_rat$|_Code$)", names(SK23))]
SK23 <- cbind(SK23, temp)



compare_datasets <- function(DF23, DF24, threshold) {

  merged_data <- merge(DF23, DF24, by = "PHI_Plan_Code", suffixes = c("_23", "_24"), all = TRUE)
  
  diff_cols <- list()
  sig_cols <- list()
  
  for(colname in names(DF23)[-1]) {  
    diff_colname <- paste0(colname, "_diff")
    diff_cols[[diff_colname]] <- merged_data[[paste0(colname, "_24")]] - merged_data[[paste0(colname, "_23")]]
    
    sig_colname <- paste0(colname, "_sig")
    sig_cols[[sig_colname]] <- ifelse(abs(diff_cols[[diff_colname]] / merged_data[[paste0(colname, "_23")]]) > threshold, "Yes", "No")
  }
  
  diff_sig_df <- data.frame(diff_cols, sig_cols)
  

  final_data <- cbind(merged_data["PHI_Plan_Code"], merged_data[, -1], diff_sig_df)
  
  return(final_data)
}

final_dataset_SK <- compare_datasets(SK23, SK24, threshold = 0.25)



# compute_difference <- function(df1, df2, threshold) {
#   columns_to_compare <- c("Denominator", "Score", "Nearest.Cluster.Center", "Component.Rating", "Composite.score", "Reliability", "Final.Rating")
#   columns_to_convert <- columns_to_compare
#   
#   for (col in columns_to_convert) {
#     df1[[col]] <- as.numeric(as.character(df1[[col]]))
#   }
#   
#   for (col in columns_to_convert) {
#     df2[[col]] <- as.numeric(as.character(df2[[col]]))
#   }
#   
#   merged_df <- full_join(df1, df2, by = c("MCO", "Service.Area", "Plan.Code", "Measurement"))
#   
#   for (col in columns_to_compare) {
#     # 2022
#     col_x <- paste(col, ".x", sep = "")
#     # 2023
#     col_y <- paste(col, ".y", sep = "")
#     
#     if (is.numeric(merged_df[[col_x]]) && is.numeric(merged_df[[col_y]])) {
#       diff_value <- abs(as.numeric(merged_df[[col_y]]) - as.numeric(merged_df[[col_x]]))
#       merged_df[paste(col, "diff", sep = "_")] <- diff_value
#       
#       sig_column <- paste(col, "sig", sep = "_")
#       condition <- !is.na(diff_value) & !is.na(merged_df[[col_x]]) & merged_df[[col_x]] != 0 & (diff_value / abs(merged_df[[col_x]]) > threshold)
#       merged_df[sig_column] <- ifelse(condition, "Yes", "No")
#       
#     } else {
#       merged_df[paste(col, "diff", sep = "_")] <- NA
#     }
#     
#   }
#   
#   return(merged_df)
# }


SA24 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\3. Survey\\Data\\Temp_Data\\SA24_out_rate.xlsx", sheet = 2)
SA23 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\3. Survey\\Data\\Temp_Data\\SA23_out.xlsx", sheet = 2)


# Only keep certain dataset:
measures <- c("AtC", "HWDC", "PDRat", "HPRat")

temp <- SA24[,measures]
SA24 <- SA24[grep("(_den$|_rat$|_Code$)", names(SA24))]
SA24 <- cbind(SA24, temp)

temp <- SA23[,measures]
SA23 <- SA23[grep("(_den$|_rat$|_Code$)", names(SA23))]
SA23 <- cbind(SA23, temp)

final_dataset_SA <- compare_datasets(SA23, SA24, threshold = 0.25)





#########################################################################
# SC
SC24 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\3. Survey\\Data\\Temp_Data\\SC24_out_rate.xlsx", sheet = 2)
SC23 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\3. Survey\\Data\\Temp_Data\\SC23_out.xlsx", sheet = 2)

names(SC23)[names(SC23) == "RoutC_den"] <- "RoutCare_den"
names(SC23)[names(SC23) == "UrgC_den"] <- "UrgCare_den"

# Only keep certain dataset:
measures <- c("GCQ", "HWDC", "PDRat", "HPRat")


temp <- SC24[,measures]
SC24 <- SC24[grep("(_den$|_rat$|_Code$)", names(SC24))]
SC24 <- cbind(SC24, temp)

temp <- SC23[,measures]
SC23 <- SC23[grep("(_den$|_rat$|_Code$)", names(SC23))]
SC23 <- cbind(SC23, temp)

final_dataset_SC <- compare_datasets(SC23, SC24, threshold = 0.25)





######################################################################
#########################################################################
#########################################################################
# SP
SP24 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\3. Survey\\Data\\Temp_Data\\SP24_out_rate.xlsx", sheet = 2)
SP23 <- read.xlsx("C:\\Users\\jiang.shao\\Dropbox (UFL)\\3. Survey\\Data\\Temp_Data\\SP23_out.xlsx", sheet = 2)


# Only keep certain dataset:
measures <- c("ATC", "HWDC", "PDRat", "HPRat")


temp <- SP24[,measures]
SP24 <- SP24[grep("(_den$|_rat$|_Code$)", names(SP24))]
SP24 <- cbind(SP24, temp)

temp <- SP23[,measures]
SP23 <- SP23[grep("(_den$|_rat$|_Code$)", names(SP23))]
SP23 <- cbind(SP23, temp)

final_dataset_SP <- compare_datasets(SP23, SP24, threshold = 0.25)



########################
write_dataframes_to_excel <- function(df_list, file_name) {
  # Create a new workbook
  wb <- createWorkbook()
  
  # Loop over the list of data frames
  for (df_name in names(df_list)) {
    # Add a worksheet to the workbook with the name of the data frame
    addWorksheet(wb, df_name)
    # Write the data frame to the worksheet
    writeData(wb, sheet = df_name, x = df_list[[df_name]])
  }
  
  # Save the workbook to a file
  saveWorkbook(wb, file = file_name, overwrite = TRUE)
}

# List of data frames to write to the Excel file
dataframes <- list(
  final_dataset_SP = final_dataset_SP,
  final_dataset_SK = final_dataset_SK,
  final_dataset_SA = final_dataset_SA,
  final_dataset_SC = final_dataset_SC
)


write_dataframes_to_excel(dataframes, "C:\\Users\\jiang.shao\\Dropbox (UFL)\\MCO Report Card - 2024\\Program\\3. Survey\\Output\\comparison_23_24.xlsx")







