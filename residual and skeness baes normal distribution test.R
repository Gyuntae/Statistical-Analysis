library(ggplot2)
library(moments)
library(readxl)
library(dplyr)
library(gridExtra)
library(openxlsx)  # For Excel export

# Create PDF with full path
pdf_path <- "C:/Users/rabae1g1/Desktop/all_histograms.pdf"
pdf(pdf_path, width=10, height=8)

# Load the dataset
file_path <- ("C:/Users/rabae1g1/Desktop/Human NASH/final!!/final/final/clinical metadata and labparameters/ald/ALD_imputed_data.xlsx")
dataset <- read_excel(file_path)

# Create dataframe to store skewness results
skewness_df <- data.frame(
  Variable = character(),
  Skewness = numeric(),
  WithinRange = logical(),
  stringsAsFactors = FALSE
)

# Loop through columns (excluding ID and Group)
for (col in 4:24) {
  var_name <- names(dataset)[col]
  
  # Get the column data
  var_data <- dataset[[col]]
  
  # Calculate residuals (deviations from mean)
  residuals <- var_data - mean(var_data, na.rm = TRUE)
  
  # Calculate skewness
  skewness_value <- skewness(residuals, na.rm = TRUE)
  
  # Add to skewness dataframe
  skewness_df <- rbind(skewness_df, data.frame(
    Variable = var_name,
    Skewness = skewness_value,
    WithinRange = (skewness_value >= -1 && skewness_value <= 1)
  ))
  
  # Determine skewness status
  if (skewness_value >= -1 && skewness_value <= 1) {
    status <- "Within [-1, 1]"
    color <- "darkgreen"
  } else {
    status <- "Outside [-1, 1]"
    color <- "darkred"
  }
  
  # Calculate mean and standard deviation
  mean_res <- mean(residuals, na.rm = TRUE)
  sd_res <- sd(residuals, na.rm = TRUE)
  
  # Use Freedman-Diaconis rule for optimal bin width
  IQR_val <- IQR(residuals, na.rm = TRUE)
  bin_width <- 2 * IQR_val / (length(na.omit(residuals))^(1/3))
  if (is.na(bin_width) || bin_width == 0) bin_width <- sd_res / 5
  
  # Calculate appropriate x-axis limits
  x_min <- min(residuals, na.rm = TRUE) - sd_res
  x_max <- max(residuals, na.rm = TRUE) + sd_res
  
  # Create and print histogram with skewness annotation
  hist_data <- data.frame(residuals = residuals)
  p <- ggplot(hist_data, aes(x = residuals)) +
    geom_histogram(aes(y = after_stat(density)), 
                   binwidth = bin_width, 
                   color = "black", 
                   fill = "skyblue") +
    stat_function(fun = dnorm, 
                  args = list(mean = mean_res, sd = sd_res), 
                  color = "red", 
                  size = 1) +
    labs(title = paste("Histogram of", var_name),
         subtitle = paste("Skewness:", round(skewness_value, 3), 
                          paste0("(", status, ")")),
         x = "Residuals",
         y = "Density") +
    theme_minimal() +
    xlim(x_min, x_max) +
    geom_vline(xintercept = mean_res, linetype = "dashed", color = "blue") +
    # Add skewness annotation directly in the plot
    annotate("text", x = mean_res + 0.5*sd_res, 
             y = max(density(residuals, na.rm=TRUE)$y, na.rm=TRUE) * 0.9,
             label = paste("Skewness:", round(skewness_value, 3)), 
             color = color,
             size = 5, hjust = 0) +
    theme(plot.subtitle = element_text(color = color))
  
  # Print the plot to PDF
  print(p)
}

# Create a visualization to compare all skewness values
# Sort by skewness value
skewness_df <- skewness_df %>%
  arrange(Skewness)

# Create a bar plot of all skewness values
skewness_plot <- ggplot(skewness_df, aes(x = reorder(Variable, Skewness), y = Skewness, 
                                         fill = WithinRange)) +
  geom_bar(stat = "identity") +
  geom_hline(yintercept = 1, linetype = "dashed", color = "darkred") +
  geom_hline(yintercept = -1, linetype = "dashed", color = "darkred") +
  scale_fill_manual(values = c("darkred", "darkgreen"), 
                    labels = c("Outside [-1, 1]", "Within [-1, 1]")) +
  labs(title = "Skewness Comparison Across All Variables",
       x = "Variable",
       y = "Skewness Value",
       fill = "Status") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 90, hjust = 1))

# Print the comparison plot to PDF
print(skewness_plot)

# For extremely skewed variables, create a plot
extreme_skewness <- skewness_df %>%
  filter(abs(Skewness) > 2)

if(nrow(extreme_skewness) > 0) {
  extreme_plot <- ggplot(extreme_skewness, aes(x = reorder(Variable, Skewness), y = Skewness)) +
    geom_bar(stat = "identity", fill = "darkred") +
    labs(title = "Variables with Extreme Skewness (|Skewness| > 2)",
         x = "Variable",
         y = "Skewness Value") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
  
  # Print extreme plot to PDF
  print(extreme_plot)
}

# Close the PDF device
dev.off()

# Create a summary table showing variables within and outside normal range
normal_vars <- skewness_df %>% 
  filter(WithinRange) %>% 
  arrange(desc(abs(Skewness))) %>%
  dplyr::select(Variable, Skewness)

skewed_vars <- skewness_df %>% 
  filter(!WithinRange) %>% 
  arrange(desc(abs(Skewness))) %>%
  dplyr::select(Variable, Skewness)

# Print summary
cat("\nSummary of Variable Skewness:\n")
cat("Variables within normal skewness range [-1, 1]:", nrow(normal_vars), "\n")
cat("Variables outside normal skewness range:", nrow(skewed_vars), "\n")
cat("\nMost skewed variables (top 5):\n")
if(nrow(skewed_vars) > 0) {
  print(head(skewed_vars %>% arrange(desc(abs(Skewness))), 5))
} else {
  cat("No skewed variables found.\n")
}

# Add additional information to skewness_df
skewness_df <- skewness_df %>%
  mutate(
    SkewnessStatus = ifelse(WithinRange, "Normal [-1, 1]", "Skewed"),
    SkewnessCategory = case_when(
      Skewness < -2 ~ "Highly Negatively Skewed",
      Skewness < -1 ~ "Moderately Negatively Skewed",
      Skewness < 0 ~ "Slightly Negatively Skewed",
      Skewness == 0 ~ "Symmetric",
      Skewness <= 1 ~ "Slightly Positively Skewed",
      Skewness <= 2 ~ "Moderately Positively Skewed",
      TRUE ~ "Highly Positively Skewed"
    ),
    TransformationSuggestion = case_when(
      Skewness > 1 ~ "Consider log, sqrt, or inverse transformation",
      Skewness < -1 ~ "Consider square or exponential transformation",
      TRUE ~ "No transformation needed"
    )
  ) %>%
  dplyr::select(Variable, Skewness, SkewnessStatus, SkewnessCategory, TransformationSuggestion)

# Create a workbook to save the results
wb <- createWorkbook()

# Add sheets for different results
addWorksheet(wb, "All Variables Skewness")
writeData(wb, "All Variables Skewness", skewness_df, rowNames = FALSE)

# Only create and write to worksheets if they have data
if(nrow(normal_vars) > 0) {
  addWorksheet(wb, "Normal Variables")
  writeData(wb, "Normal Variables", normal_vars, rowNames = FALSE)
}

if(nrow(skewed_vars) > 0) {
  addWorksheet(wb, "Skewed Variables")
  writeData(wb, "Skewed Variables", skewed_vars, rowNames = FALSE)
}

if(nrow(extreme_skewness) > 0) {
  addWorksheet(wb, "Extremely Skewed Variables")
  writeData(wb, "Extremely Skewed Variables", extreme_skewness %>% dplyr::select(Variable, Skewness), rowNames = FALSE)
}

# Create the header style
headerStyle <- createStyle(
  fgFill = "#4F81BD", 
  halign = "center", 
  textDecoration = "bold",
  fontColour = "white",
  border = "TopBottom"
)

# Apply styling to headers - fix the error by checking if data exists
for(sheet in names(wb)) {
  setColWidths(wb, sheet, cols = 1:5, widths = c(30, 15, 20, 30, 40))
  # Read the sheet data first to check if it exists
  sheet_data <- tryCatch({
    readWorkbook(wb, sheet)
  }, error = function(e) {
    NULL
  })
  
  # Only apply styles if the sheet has data and columns
  if(!is.null(sheet_data) && nrow(sheet_data) > 0 && ncol(sheet_data) > 0) {
    addStyle(wb, sheet, headerStyle, rows = 1, cols = 1:ncol(sheet_data))
  }
}

# Save the workbook
output_path <- "C:/Users/rabae1g1/Desktop/Skewness_Analysis_Results.xlsx"
saveWorkbook(wb, output_path, overwrite = TRUE)

cat("\nSkewness analysis results have been saved to:", output_path, "\n")
cat("\nAll histograms have been saved to:", pdf_path, "\n")