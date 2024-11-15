# Load necessary libraries
library(boot)      # For bootstrapping
library(readxl)    # For reading Excel files
library(writexl)   # For writing Excel files

# Function to calculate p-values based on different statistical tests
pmake <- function(t, hyp = 0, side = 2) {
  
  p_value <- matrix(data = NA, ncol = ncol(t))  # Initialize matrix to store p-values
  
  # One-sided test when a non-zero hypothesis is given
  if (side == 1 & mean(hyp) != 0) {
    for (ii in 1:ncol(t)) {
      p_value[ii] <- mean(t[, ii] >= hyp[ii])
    }
  } 
  
  # Two-sided test when no hypothesis is provided
  else if (side == 2 & mean(hyp) == 0) {
    med <- apply(X = t, MARGIN = 2, FUN = median)  # Calculate median for each column
    t2 <- matrix(rep(med, times = nrow(t)), nrow = nrow(t), byrow = TRUE)  # Replicate median across rows
    t3 <- abs(t - t2)  # Absolute deviation from the median
    
    for (ii in 1:ncol(t3)) {
      p_value[ii] <- mean(t3[, ii] >= abs(0 - med[ii]))  # Calculate p-value based on deviation from median
    }
  }
  
  # Two-sided test with a non-zero hypothesis
  else if (side == 2 & mean(hyp) != 0) {
    med <- apply(X = t, MARGIN = 2, FUN = median)  # Calculate median for each column
    t2 <- matrix(rep(med, times = nrow(t)), nrow = nrow(t), byrow = TRUE)  # Replicate median across rows
    t3 <- abs(t - t2)  # Absolute deviation from the median
    
    for (ii in 1:ncol(t3)) {
      p_value[ii] <- mean(t3[, ii] >= abs(hyp[ii] - med[ii]))  # Calculate p-value based on deviation from hypothesis
    }
  }
  
  return(p_value)  # Return p-values
}

# Read Excel data from the file
df <- read_excel("E:/长新冠共激活/相关.xlsx")

# Define columns for correlation analysis
target_columns <- c("gaohuanxing", "gaoqinxi", "huibi", "T2_SUM_IES_R")

# Select the relevant columns from the dataframe
selected_data <- df[, target_columns]  # Data with target columns
other_data <- df[, !(names(df) %in% target_columns)]  # Data with other columns

# Function to compute Pearson correlation using bootstrap
cor_func <- function(data, indices, x_col, y_col) {
  d <- data[indices, ]  # Extract bootstrap sample
  
  # Try to compute Pearson correlation, return NULL if error occurs
  model <- tryCatch({
    cor(d[[x_col]], d[[y_col]], method = "pearson")
  }, error = function(e) NULL)
  
  return(model)  # Return the correlation result
}

# Initialize an empty dataframe to store results
results_df <- data.frame(X_Column = character(), Y_Column = character(), Correlation = numeric(), P_Value = numeric())

# Loop through each column in selected_data and other_data to compute correlations and p-values
for (x_col in names(selected_data)) {
  for (y_col in names(other_data)) {
    
    # Extract the two columns for analysis
    temp_data <- df[, c(x_col, y_col)]
    
    # Perform bootstrap to compute the correlation with 1000 resamples
    boot_result <- boot(data = temp_data, statistic = cor_func, sim = "ordinary", R = 1000, x_col = x_col, y_col = y_col)
    
    # Compute p-value for the correlation
    t0 <- boot_result$t0  # Observed correlation
    t <- boot_result$t  # Bootstrapped correlations
    p_diff <- pmake(t, 0)  # Calculate p-value based on bootstrap results
    p_value <- p_diff[1]  # Extract p-value
    
    # Append the results to the dataframe
    results_df <- rbind(results_df, data.frame(X_Column = x_col, Y_Column = y_col, Correlation = mean(boot_result$t), P_Value = p_value))
  }
}

# Write the results to a new Excel file
write_xlsx(results_df, "E:/corr.xlsx")
