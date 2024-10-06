# ChiSquareAutomator R Package

**ChiSquareAutomator** is an R package designed to automate the process of performing Chi-square tests on categorical data stored in Excel files. It reads data from an Excel sheet, performs Chi-square tests between two groups (cohorts), writes the p-values back into the Excel file, and formats the results.

## Table of Contents
- [Introduction](#introduction)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Input File Format](#input-file-format)
- [Functions](#functions)
- [Example](#example)
- [Contributing](#contributing)
- [License](#license)

## Introduction

**ChiSquareAutomator** helps researchers and data analysts to automate the statistical testing process for categorical data. Specifically, it automates the Chi-square testing between two sample groups (Segment A and Segment B) and writes the results back to the same Excel sheet. It also applies formatting for easier interpretation of results, such as highlighting cells with non-significant p-values.

## Features

- Automatically reads data from a designated Excel file.
- Performs Chi-square tests with or without continuity correction based on the data.
- Writes computed p-values back to the Excel file.
- Highlights non-significant p-values (p > 0.05) with green cells for better visibility.
- Allows for flexibility in input sheet format and quick statistical analysis.

## Installation

To use this package, you can install it directly from GitHub. Ensure that you have the `devtools` package installed first.

```r
# Install devtools if you don't have it already
install.packages("devtools")

# Install the ChiSquareAutomator package
devtools::install_github("your-repo/ChiSquareAutomator")

## Usage
After installing the package, you can use it to run chi-square tests on your dataset stored in an Excel file.

### Basic Usage
```r
# Load the package
library(ChiSquareAutomator)

# Run the automation on your Excel file
run_automation("path/to/Input_File.xlsx")

The function `run_automation()` will handle everything: reading the data, performing chi-square tests, and writing the results back to the file.
