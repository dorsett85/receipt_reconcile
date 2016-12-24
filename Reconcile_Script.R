
##### BRS CLER Reconcile Generator Script #####

# Setup
suppressMessages(library(dplyr))
suppressMessages(library(tidyr))
suppressMessages(library(lubridate))
suppressMessages(library(XLConnect))
options(stringsAsFactors = F)

# Write intro header
cat("\n")
cat(paste0(rep("#", 78), collapse = ""), "\n\n")
cat(paste0(rep(" ", 22), collapse = ""), "**Behavioral Research Services**\n")
cat(paste0(rep(" ", 24), collapse = ""), "**CLER Reconcile Generator**\n\n")
cat(paste0(rep("#", 78), collapse = ""), "\n\n")

#-----------------------------------------------------------------------------#
# 1. Create and export receipt_data from given date range

# Import "signups_analysis" & "prescreen_data"

signups_analysis <- read.csv("signups_analysis.csv", skip = 1)
prescreen_data <- read.csv("prescreen_data.csv", skip = 1)

# Add email to prescreen_data

user_email <- read.csv("users_analysis.csv", skip = 1)

prescreen_data <- prescreen_data %>%
  
  # Join prescreen and user_email
  left_join(user_email, by = c("first_name", "last_name")) %>%
  
  # Combine first and last name and select prescreen_data columns
  mutate(student_name = paste(first_name, last_name)) %>%
  select(student_name, sona_id = id_code, email = alt_email)

rm(user_email)

# Format signups_analysis

signups_analysis <- signups_analysis %>%
  
  # Add email and Sona ID's and remove duplicates
  left_join(prescreen_data, "student_name") %>%
  distinct(timeslot_date, sona_id, .keep_all = T) %>%
  distinct(timeslot_date, student_name, email, .keep_all = T) %>%
  
  # Create separate columns for last and first names
  separate(student_name_last, c("last_name", "first_name"), sep = ", ")

# Fill in start_date and end_date to filter a date range

## Create function to get date range user input when running from R or the
## command line.

date_range <- function() {
  cat("Enter Date Range...\n")
  if (interactive()) {
    start_date <- mdy(readline("Start Date (mm/dd/yy): "))
    end_date <- mdy(readline("End Date (mm/dd/yy): ")) + 1
    } else {
    cat("Start Date (mm/dd/yy): ")
    start_date <- mdy(readLines("stdin", n = 1))
    cat("End Date (mm/dd/yy): ")
    end_date <- mdy(readLines("stdin", n = 1)) + 1
  }
  return(signups_analysis %>%
           filter(mdy_hms(timeslot_date) >= start_date &
                    mdy_hms(timeslot_date) <= end_date))
}
signups_dates <- date_range()

rm(date_range)

# Join receipt_data and noshow_report

## Import no-show data

noshow_report <- read.csv("noshow_report.csv", skip = 1, dec = "_")

## Change no show payment amount and remove unexused no shows

noshow_report$payment <- ifelse(grepl("\\$10", noshow_report$Comments) == T, 10, 0)

## Create columns to join and join dataframes

cols_to_join <- c("first_name" = "First.Name",
                  "last_name" = "Last.Name",
                  "experiment_name" = "Study",
                  "timeslot_date" = "Date")

receipt_data <- signups_dates %>% 
  
  # Select and sort columns
  select(sona_id, last_name, first_name, timeslot_date, experiment_name, email) %>%
  arrange(experiment_name, timeslot_date, last_name) %>%
  
  # Join, remove no shows, and select columns
  left_join(noshow_report, by = cols_to_join) %>%
  replace_na(list(payment = 20)) %>%
  filter(payment != 0) %>%
  select(-Email, -Credits, -Unexcused.No.Show., -Comments) %>%
  
  # Add separate columns for date and time
  separate(timeslot_date, c("date", "time", "hours"), sep = " ") %>%
  unite(time, time, hours, sep = " ")

rm(cols_to_join)

# Create function to get input experiment authors when running from R or the
# command line.

author <- function() {
  receipt_data$author <- NA
  cat("\n")
  if (interactive()) {
    z <- unique(receipt_data$experiment_name)
    cat("Who are the authors for...")
    for (i in z) {
      receipt_data[receipt_data$experiment_name == i, "author"] <- readline(paste0(i, ": "))
    }
  } else {
    z <- unique(receipt_data$experiment_name)
    cat("Who are the authors for...\n")
    for (i in z) {
      cat(paste0(i, ": "))
      receipt_data[receipt_data$experiment_name == i, "author"] <- readLines("stdin", n = 1)
    }
  }
  return(receipt_data$author)
}
receipt_data$author <- author()

rm(author)

# Remove unneeded objects

rm(noshow_report, prescreen_data, signups_analysis, signups_dates)

# Export and open receipt_data.csv for payment entry

write.csv(receipt_data, "receipt_data.csv", row.names = F)
cat("\nOpening receipt_data.csv")
Sys.sleep(0.5)
cat(".")
Sys.sleep(0.5)
cat(".")
Sys.sleep(0.5)
cat(".\n")
shell("receipt_data.csv")

# Prompt to fill in payment data before running the next part of the script

if (interactive()) {
  cat("\n")
  readline("Have you filled in payment data? If yes press Enter. ")
  cat("\n")
} else {
  cat("\n")
  cat("Have you filled in payment data? If yes press Enter. ")
  invisible(readLines("stdin", n = 1))
  cat("\n")
}

#-----------------------------------------------------------------------------#
# 2. Create SessionFinancialReports

# Import receipt_data.csv after entering payment information

reconcile_data <- read.csv("receipt_data.csv")

# Change time column format

reconcile_data$time <- tolower(format(strptime(reconcile_data$time, "%I:%M:%S %p"), 
                                      "%l:%M%p"))

# Create list of dataframes for each study name and date_time

reconcile_list <- split(reconcile_data, 
                        list(factor(reconcile_data$experiment_name),
                             factor(paste(reconcile_data$date,
                                          reconcile_data$time))))

# Filter dataframes with 0 observations

reconcile_clean <- reconcile_list[sapply(reconcile_list, function(x) nrow(x)) > 0]

# Modify names so they can be exported to Excel

names(reconcile_clean) <- gsub("\\.", " ", names(reconcile_clean))
names(reconcile_clean) <- gsub(":|\\/", ".", names(reconcile_clean))

# Create participant report and summary worksheets

## Create and apply function to modify dataframes into participant report worksheet

participantWorksheet <- function(df){
  names <- paste(df$last_name, df$first_name, sep = ", ")
  email <- df$email
  payments <- df$payment
  col_one <- c("Session Name:", "Session Date", "Session Time:", 
               "Researchers:", "Payment Source:", "Payment Type:",
               NA, NA, "Record", 1:length(names), NA)
  col_two <- c(as.vector(df$experiment_name[1]),
               as.vector(df$date[1]),
               as.vector(df$time[1]),
               as.vector(df$author[1]), 
               "HBS", "Cash",
               NA, NA, "Name",
               names, NA)
  col_three <- c(rep(NA, 8), "Email", email, NA)
  col_four <- c(rep(NA, 8), "Payment Amount", payments, sum(payments))
  return(data.frame(col_one, col_two, col_three, col_four))
}

reconcile_output <- lapply(names(reconcile_clean), function(x) {
  participantWorksheet(reconcile_clean[[x]])
})

## Create and apply function to modify dataframes into summary worksheet

SummaryReport <- function(df) {
  col_one <- rep(NA, 12)
  col_two <- c(rep(NA, 3),
               "Session Name:",
               "Session Date:",
               "Session Time:",
               "Researchers:",
               rep(NA, 5))
  col_three <- c(NA,
                 "Session Financial Report",
                 NA,
                 df$experiment_name[1],
                 df$date[1],
                 df$time[1],
                 df$author[1],
                 NA,
                 "Payment Source Desc",
                 "HBS",
                 NA,
                 "Grand Total:")
  col_four <- c(rep(NA, 8),
                "Payment Desc",
                "Cash",
                rep(NA, 2))
  col_five <- c(rep(NA, 8),
                "Payment Amount",
                sum(df$payment),
                NA,
                sum(df$payment))
  return(data.frame(col_one, col_two, col_three, col_four, col_five))
}

reconcile_sum <- lapply(names(reconcile_clean), function(x) {
  SummaryReport(reconcile_clean[[x]])
})

## Create Excel workbooks

reconcile_excel <- lapply(1:length(reconcile_output), function (x) {
  loadWorkbook(paste0(names(reconcile_clean[x]), ".xlsx"), create = T)
})

## Format and enter data for all reconcile_excel worksheets
  
cat("Formatting Workbooks:\n\n")

invisible(lapply(1:length(reconcile_excel), function(x) {
  
  cat(paste0(names(reconcile_clean[x]), " ... ", x, "/", length(reconcile_clean)))
  cat("\n")
  
  # Formart participant_report worksheet
  
  ## Create participant_report and summary worksheets for each workbook
  createSheet(reconcile_excel[[x]], c("participant_report", "summary"))
  
  ## Add reconcile_output data to each worksheet
  writeWorksheet(reconcile_excel[[x]], reconcile_output[[x]], "participant_report", header = F)
  
  ## Create and apply function to change column width
  col_width <- function(col_num, width, sheet) {
    setColumnWidth(reconcile_excel[[x]], sheet, 
                   column = col_num, width = width)
  }
  col_width(1, 4100, "participant_report")
  col_width(2, 9700, "participant_report")
  col_width(3, 9700, "participant_report")
  col_width(4, 4300, "participant_report")
  
  ## Create and apply function to change border
  cell_style <- createCellStyle(reconcile_excel[[x]])
  
  setBorder(cell_style, side = "all", type = XLC$"BORDER.THIN",
            color = c(XLC$"COLOR.BLACK"))
  
  border <- function(col_num) {
    setCellStyle(reconcile_excel[[x]], sheet = "participant_report",
                 row = c(9:nrow(reconcile_output[[x]])), 
                 col = col_num, cellstyle = cell_style)
  }
  border(1)
  border(2)
  border(3)
  border(4)
  
  # Formart summary worksheet
  
  ## Add reconcile_sum data to reconcile_excel summary worksheets
  writeWorksheet(reconcile_excel[[x]], reconcile_sum[[x]], "summary", header = F)
  
  ## Add border
  cell_style <- createCellStyle(reconcile_excel[[x]])
  
  setBorder(cell_style, side = "all", type = XLC$"BORDER.THIN",
            color = c(XLC$"COLOR.BLACK"))
  
  border <- function(col_num, row_start, row_end) {
    setCellStyle(reconcile_excel[[x]], sheet = "summary",
                 row = c(row_start:row_end), 
                 col = col_num, cellstyle = cell_style)
  }
  border(3, 9, 10)
  border(4, 9, 10)
  border(5, 9, 10)
  
  # Change cell width
  col_width(2, 4000, "summary")
  col_width(3, 5000, "summary")
  col_width(4, 4000, "summary")
  col_width(5, 4100, "summary")
}))

# Prompt to fill in percent owed data before running the next part of the script

if (interactive()) {
  cat("\n")
  readline("Have you filled in percent owed data? If yes press Enter. ")
  cat("\n")
} else {
  cat("\n")
  cat("Have you filled in percent owed data? If yes press Enter. ")
  invisible(readLines("stdin", n = 1))
  cat("\n")
}
# Remove unneeded objects

rm(participantWorksheet, reconcile_clean, reconcile_list, 
   reconcile_sum, SummaryReport)

#-----------------------------------------------------------------------------#
# 3. Create CLERCoverSheets

# Summarise receipt_data by day of the previous week

coversheet_data <- reconcile_data %>%
  group_by(date, experiment_name) %>%
  summarise(payment = sum(payment),
            researcher = unique(author),
            portion = "To Be Determined",
            root = "To Be Determined")

# Import CLERCoverSheet templates into XLConnect interface

xl_cover_sheets <- rep(lapply("CLERCoverSheetTemplate.xlsx", loadWorkbook), 
                       nrow(coversheet_data))

# Modify names so they can be exported to Excel

names(xl_cover_sheets) <- paste(coversheet_data$experiment_name,
                                gsub("\\/", ".", coversheet_data$date))

# Add experiment_name, payment portion, and root number


#-----------------------------------------------------------------------------#
# 4. Export SessionFinancialReports and CLERCoverSheets

# Function to Create directories and export workbooks

export <- function() {
  
  # Create directory names
  start_date <- format(mdy(min(coversheet_data$date)), "%Y-%b-%d")
  end_date <- format(mdy(max(coversheet_data$date)), "%Y-%b-%d")
  date_range <- paste0(start_date, ".", end_date)
  payment <- paste0(date_range, "/")
  sub_fins <- paste0(payment, "SessionFinancialReports", "/")
  sub_cvrs <- paste0(payment, "CLERCoverSheets", "/")
  
  # Create directories
  dir.create(payment, showWarnings = F)
  dir.create(sub_fins, showWarnings = F)
  dir.create(sub_cvrs, showWarnings = F)
  
  # Export workbooks
  
  ## SessionFinancialReports
  cat("Exporting SessionFinancialReports:\n\n")
  
  invisible(lapply(1:length(reconcile_output), function(x) {
    cat(paste0(reconcile_excel[[x]]@filename, " ... ", x, "/", length(reconcile_excel)))
    cat("\n")
    saveWorkbook(reconcile_excel[[x]], file = paste0(sub_fins, 
                                                     reconcile_excel[[x]]@filename))
  }))
  ## CLERCoverSheets
  cat("\n")
  cat("Exporting CLERCoverSheets:\n\n")
  
  invisible(lapply(1:length(xl_cover_sheets), function(x) {
    cat(paste0(names(xl_cover_sheets[x]), " ... ", x, "/", length(xl_cover_sheets)))
    cat("\n")
    saveWorkbook(xl_cover_sheets[[x]], file = paste0(sub_cvrs,
                                                     names(xl_cover_sheets[x]), ".xlsx"))
  }))
}
export()

# Remove unneeded objects

rm(export, reconcile_excel, reconcile_output, xl_cover_sheets)

# Exit message
cat("\n")
cat("Success!\n")
Sys.sleep(2)
