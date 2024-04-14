#creator: Eve McAleer
#last edited: 21/11/23
#Current run time for entire database : ~3 minutes

#Aspects of this code have been deleted/anonymised

# This code is for the Data Validation Project being undertaken in *team_name*. 
# The code automates running SQL checks on the *central*
# database. The SQL is read in, run on SQL Developer, and the results output to
# worksheets by Layer, for cleaning purposes.

#1. Prep 
#2. Function to execute SQL check
#3. Function to write output results to sheets in workbook 
#4. Function to generate summary report 
#5. Function to generate an acronym for long clean names 
#6. Connect to Database
#7. Loop through layers 
#8. Disconnect from database 

############################################################################################

#Testing: (last working fully 21/11/23)
#Output written to each sheet, with the sheet name corresponding to the clean name
#When results are null, blank sheet added with clean name, and disclaimer still added
#Try-Catch handling database-side error
#Disclaimer being added to top of each sheet
#Output Messages working for null results, blank worksheets, output files, acronyms, and invalid code
#Descriptions of each clean being added to cell G3 of the respective workbooks
#Added in result count per clean as an output message to track the progress of the run
#Added Summary Report functionality back in

#Testing: (last tested 15/11/23)


##################################################################################

# 1.0 Load required libraries (ensure they have been installed first)
library(readxl)
library(RJDBC)
library(xlsx)
library(openxlsx)
library(rJava)
library(writexl)

#1.1 Set working directory
setwd('Path_to_folder')

#2. Function to execute SQL checks and return the results 
executeSQLChecks <- function(connection, sql_code, layer, clean_name) {
  tryCatch({
    result_set <- dbSendQuery(connection, sql_code)
    results <- dbFetch(result_set, n = -1)  # Fetch all rows
    dbClearResult(result_set)
    
    if (nrow(results) == 0) {
      #cat("No data that requires cleaning for '", layer, ":", clean_name, "'.\n")
      return(NULL)  # Return NULL to indicate no results
    }
    
    return(results)
  }, warning = function(w) {
    # Handle warnings (e.g., developer messages)
    cat("Warning for '", layer, ":", clean_name, "': ", w$message, "\n")
    # Output the warning message to the worksheet
    return(data.frame(Error_Message = w$message))
  }, error = function(e) {
    # Handle SQL errors here
    cat("SQL error for '", layer, ":", clean_name, "': ", e$message, "\n")
    return(NULL)
  })
}

#3. Function to write output results to sheets in workbook 
writeWorksheet <- function(layer_workbook, results, clean_name, disclaimer, layer, desc, results_sum) {
  # Check if clean_name exceeds 31 characters
  if (nchar(clean_name) > 31) {
    # Generate an acronym and message for the clean name
    acronym_result <- generateAcronymWithMessage(clean_name)
    abbreviation <- acronym_result$abbreviation
    message <- acronym_result$message
    
    cat(message, "\n")  # Print the message with full name and abbreviation
    
    # Use the abbreviation for the sheet name
    clean_name <- abbreviation
  }
  
  # Always add a worksheet with the clean name
  addWorksheet(layer_workbook, sheetName = clean_name)
  
  
  # Print a message with the check name and the number of results
  cat(results_sum, "result(s) found for ", layer, ":", clean_name,  "\n")
  
  if (!is.null(disclaimer)) {
    # Add the disclaimer to the top line of the worksheet
    writeData(layer_workbook, sheet = clean_name, x = data.frame(disclaimer))
    
  }
  
  if (is.null(results)) {
    cat(" - Added an empty worksheet for '", layer, ":", clean_name, "'.\n")
    results <- data.frame()  # Create an empty data frame
  } else {
    # Write the results for this clean to the corresponding sheet
    writeData(layer_workbook, sheet = clean_name, results, startRow = if (!is.null(disclaimer)) 2 else 1)
  }
  
  # Add desc to cell G3 without overwriting existing data
  writeData(layer_workbook, sheet = clean_name, x = data.frame(Description = desc), startCol = 7, startRow = 3)
}

# 4. Function to generate the summary report for a layer 
generateSummaryReport <- function(layer, all_results) {

  # Save the summary report workbook
  summary_output_file <- file.path('Testing', paste0(layer, "_Summary_Report.xlsx"))
  write.xlsx(all_results, summary_output_file, rowNames = FALSE, colNames = TRUE)
  
  cat(paste0("Summary Report for ", layer, " saved at:", summary_output_file, "\n"))
}

#5. Function to, for longer clean_names, create an acronym and generate a message with the actual name
generateAcronymWithMessage <- function(long_name) {
  # Split the long name into words using underscores as separators
  words <- unlist(strsplit(long_name, "_"))
  
  # Extract the first letter from each word and capitalize it
  acronym <- toupper(substr(words, 1, 1))
  
  # Combine the acronym letters to form the abbreviation
  abbreviation <- paste(acronym, collapse = "")
  
  # Create an output message with both the acronym and full name
  message <- sprintf("Using acronym '%s' for clean name '%s'", abbreviation, long_name)
  
  # Return a list with the abbreviation and message
  return(list(abbreviation = abbreviation, message = message))
}

#6.0 create driver and connection objects
jdbcDriver =JDBC("oracle.jdbc.OracleDriver",classPath=c("Driver_Path/ojdbc8.jar"))

#6.1 create connection to Oracle database
con =dbConnect(jdbcDriver, 
               "jdbc:oracle:thin:@//DB_Details", 
               "DB_Name", 
               "Password")

#7.0 User Input: Select the level(s) for validation 
#selected_layers <- c("SECTION","COMPONENT","DISTRIBUTIONS")  # Modify as per requirements

#7.0.1 Automatic run of all layers - Anonymised
selected_layers <- c("SQUARE", "SECTION","COMPONENT", "PLOTPOINT",
                     "BROWSING DAMAGE", "POINT FEATURES", "LINE FEATURES", "STEMS","DISTRIBUTIONS")

#7.0.2 Pull disclaimer from sql_store
disclaimer_read <- read_excel("SQL_Store.xlsx", sheet = "DISCLAIMER")
disclaimer <- disclaimer_read[1, 1]

# 7.1 Loop through selected layers, execute SQL checks, and save output workbooks
for (layer in selected_layers) {
  cat("Processing layer:", layer, "\n")
  # Read the SQL code from the worksheet
  data <- read_excel("SQL_Store.xlsx", sheet = layer)
  
  # Creating workbook - layer
  layer_workbook <- createWorkbook()
  
  all_results <- data.frame(Check = character(), `Number_of_Cleans` = integer())  # Initialize df with column names
  
  for (row in 1:nrow(data)) {
    # Extract unique clean names from the "sqlstore" sheet
    clean_name <- unique(data$Check[row])
    
    # Extract description for each clean
    desc <- (data$Description[row])
    
    # Execute SQL checks and get the results
    results <- executeSQLChecks(con, data$SQL[row], layer, clean_name)
    
    # Calculate the number of results
    results_sum <- if (is.null(results) || nrow(results) == 0) 0 else nrow(results)
    
    # Write each clean to sheet
    writeWorksheet(layer_workbook, results, clean_name, disclaimer, layer, desc, results_sum)
    
    # Add clean_name and results_sum to all_results
    all_results <- rbind(all_results, data.frame(Check = clean_name, `Number_of_Cleans` = results_sum))
  }
  
  # Save the layer workbook with multiple sheets
  output_file_name <- 'Testing'
  output_file <- file.path(output_file_name, paste0(layer, "_output.xlsx"))
  saveWorkbook(layer_workbook, output_file)
  
  # Generate and save the summary report for the layer
  generateSummaryReport(layer, all_results)

}

# 8. Disconnect from the database
dbDisconnect(con)