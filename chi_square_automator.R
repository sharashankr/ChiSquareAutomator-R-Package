# Load necessary libraries
library(openxlsx)
library(readxl)
library(dplyr)

#' Load Data from Excel
#'
#' This function loads the Excel workbook and converts the specified sheet into a dataframe.
#'
#' @param file_path Path to the Excel file.
#' @param sheet_name Name of the worksheet to read (default is "Standard Stat Testing").
#' @return A list containing the workbook object and the dataframe.
#' @export
load_data <- function(file_path, sheet_name = "Standard Stat Testing") {
  wb <- loadWorkbook(file_path)
  df <- as.data.frame(read.xlsx(wb, sheet = sheet_name))
  return(list(wb = wb, df = df))
}

#' Perform Chi-square Tests
#'
#' This function performs the chi-square test on the dataset and returns the results.
#'
#' @param df Dataframe that contains the data from the Excel sheet.
#' @return A list of test results, including chi-square statistics, degrees of freedom, expected values, and p-values.
#' @export
perform_chi_square <- function(df) {
  results <- list()
  for (i in 1:nrow(df)) {
    a_size <- df[i, 8]
    a_label <- df[i, 5]
    b_size <- df[i, 12]
    b_label <- df[i, 9]
    cohort_a <- df[i, 6]
    cohort_b <- df[i, 10]
    
    observed <- matrix(c(a_size, cohort_a - a_size, b_size, cohort_b - b_size), nrow = 2, byrow = TRUE)
    
    test_result <- if (min(observed) <= 7) {
      chisq.test(observed, correct = TRUE)
    } else {
      chisq.test(observed, correct = FALSE)
    }
    
    results[[i]] <- list(
      chi2_stat = test_result$statistic,
      p_value = test_result$p.value,
      dof = test_result$parameter,
      expected = test_result$expected
    )
  }
  return(results)
}

#' Write Results to Excel
#'
#' This function writes the results (p-values) back to the Excel sheet and applies formatting.
#'
#' @param wb The workbook object.
#' @param df The original dataframe.
#' @param results The list of test results generated by `perform_chi_square()`.
#' @export
write_results <- function(wb, df, results) {
  for (i in 1:length(results)) {
    p_value <- results[[i]]$p_value
    p_value_t <- ifelse(p_value < 0.01, "<0.01", round(p_value, 4))
    
    writeData(wb, sheet = "Standard Stat Testing", x = p_value_t, startRow = i + 2, startCol = 13)
    
    if (p_value > 0.05) {
      style <- createStyle(fgFill = "#90EE90", halign = "center", fontSize = 12, border = "TopBottomLeftRight")
      addStyle(wb, sheet = "Standard Stat Testing", style = style, rows = i + 2, cols = 13, gridExpand = TRUE)
    }
  }
  
  saveWorkbook(wb, "Input_File.xlsx", overwrite = TRUE)
}

#' Run Automation
#'
#' This is the main function that coordinates loading data, performing the chi-square tests, and writing the results.
#'
#' @param file_path Path to the Excel file.
#' @export
run_automation <- function(file_path) {
  data <- load_data(file_path)
  results <- perform_chi_square(data$df)
  write_results(data$wb, data$df, results)
}
