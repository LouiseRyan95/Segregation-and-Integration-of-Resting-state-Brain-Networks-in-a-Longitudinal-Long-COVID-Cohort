# Load necessary libraries
library(readxl)
library(writexl)
library(lmPerm)
library(boot)
library(emmeans)

# Function to compute p-values for permutation tests (one-sided or two-sided)
pmake <- function(t, hyp = 0, side = 2) {
  # Initialize the p_value matrix with NA values
  p_value <- matrix(data = NA, ncol = ncol(t))
  
  # One-sided test with non-zero hypothesis
  if (side == 1 & mean(hyp) != 0) {
    for (ii in 1:ncol(t)) {
      p_value[ii] <- mean(t[, ii] >= hyp[ii])
    }
  } 
  # Two-sided test with zero hypothesis
  else if (side == 2 & mean(hyp) == 0) {
    med <- apply(X = t, MARGIN = 2, FUN = median)
    t2 <- matrix(rep(med, times = nrow(t)), nrow = nrow(t), byrow = TRUE)
    t3 <- abs(t - t2)
    for (ii in 1:ncol(t3)) {
      p_value[ii] <- mean(t3[, ii] >= abs(0 - med[ii]))
    }
  } 
  # Two-sided test with non-zero hypothesis
  else if (side == 2 & mean(hyp) != 0) {
    med <- apply(X = t, MARGIN = 2, FUN = median)
    t2 <- matrix(rep(med, times = nrow(t)), nrow = nrow(t), byrow = TRUE)
    t3 <- abs(t - t2)
    for (ii in 1:ncol(t3)) {
      p_value[ii] <- mean(t3[, ii] >= abs(hyp[ii] - med[ii]))
    }
  }
  
  return(p_value)
}

# Function to perform bootstrapping and hypothesis testing (paired t-test)
hetert <- function(data, i, res, group, method) {
  
  # Prepare the results data frame
  results_df <- data.frame(Results = data[[res]], Group = data[[group]])
  results_dfu <- results_df
  
  # Modify results based on method (permutation or other)
  if (method == "permutation") {
    results_dfu[, "Results"] <- results_df[i, "Results"]
  } else {
    results_dfu <- results_df[i, ]
  }
  
  # Get unique group levels
  group_levels <- unique(results_dfu$Group)
  
  # Perform paired t-test if there are two groups
  if (length(group_levels) == 2) {
    group1 <- results_dfu %>% filter(Group == group_levels[1]) %>% pull(Results)
    group2 <- results_dfu %>% filter(Group == group_levels[2]) %>% pull(Results)
    
    t_test_result <- t.test(group1, group2, paired = TRUE)
    t_value <- t_test_result$statistic
  } else {
    t_value <- NA
  }
  
  # Return the results
  if (!is.null(t_test_result)) {
    t_value <- t_test_result$statistic
    p_value <- t_test_result$p.value
    conf_int <- t_test_result$conf.int
    mean_diff <- t_test_result$estimate
    
    return(c(t_value = t_value, p_value = p_value, 
             conf_int_low = conf_int[1], conf_int_high = conf_int[2],
             mean_diff = mean_diff))
  } else {
    return(rep(NA, 5))
  }
}

# Set seed for reproducibility
set.seed(725)

# Read the input data from Excel file
data <- read_excel("E:/nsp/T1T2.xlsx")

# Convert SEX to a factor (for use in group comparison)
data$sex_female <- as.factor(data$SEX)

# Convert Time to a factor (for analysis)
data$Time <- as.factor(data$Time)

# Check the unique groups
unique(data$Group)

# Set up sampling for bootstrap
l <- nrow(data)
i <- sample(c(1:l), replace = TRUE)

# List of results to analyze
res_all <- c("HIN_Default", "HIN_Cont", "HIN_Limbic", "HIN_SalVentAttn", "HIN_DorsAttn", 
             "HIN_SomMot", "HIN_Vis", "HSE_Default", "HSE_Cont", "HSE_Limbic", "HSE_SalVentAttn", 
             "HSE_DorsAttn", "HSE_SomMot", "HSE_Vis", "HB_Default", "HB_Cont", "HB_Limbic", 
             "HB_SalVentAttn", "HB_DorsAttn", "HB_SomMot", "HB_Vis")

# Define group and variable for analysis
group <- "Time"
var <- "Time"
sim1 <- "permutation"

# Initialize list to store results
output_list <- list()

# Loop through all results and perform bootstrapping and hypothesis testing
for (res in res_all) {
  # Perform bootstrapping
  boot_result <- boot(
    data = data,
    statistic = hetert,
    R = 1000,
    sim = sim1,
    res = res,
    group = group,
    method = sim1,
    parallel = "no"
  )
  
  # Extract t-values and p-values from the bootstrap results
  t0 <- boot_result$t0
  t <- boot_result$t
  p_diff <- pmake(t[, 5, drop = FALSE], t0[5])
  
  # Create a data frame with the results
  result_df <- data.frame(
    t_value = t0[1],
    p_value = p_diff[1],
    conf_int_low = t0[3],
    conf_int_high = t0[4],
    mean_diff = t0[5]
  )
  
  # Save the results to the output list
  output_list[[res]] <- result_df
}

# Combine all results into one final data frame
final_df <- do.call(rbind, output_list)

# Save the final results to an Excel file
write_xlsx(list(emmeans_df = final_df), "E:/nsp/tpair.xlsx")
