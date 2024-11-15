# Load necessary libraries
library(readxl)
library(writexl)
library(lmPerm)
library(boot)
library(emmeans)

# Function to calculate p-values based on various conditions
# t: Matrix of values to compute p-values from
# hyp: Hypothesis values (default is 0)
# side: The type of test: 1 = one-sided, 2 = two-sided
pmake <- function(t, hyp = 0, side = 2) {
  p_value <- matrix(data = NA, ncol = ncol(t))  # Initialize p-value matrix

  # For one-sided tests with non-zero hypotheses
  if (side == 1 & mean(hyp) != 0) {
    for (ii in 1:ncol(t)) {
      p_value[ii] <- mean(t[, ii] >= hyp[ii])  # Calculate p-value for one-sided test
    }
  } 
  # For two-sided tests with zero hypothesis
  else if (side == 2 & mean(hyp) == 0) {
    med <- apply(t, 2, median)  # Calculate median for each column
    t2 <- matrix(rep(med, times = nrow(t)), nrow = nrow(t), byrow = TRUE)  # Repeat median values
    t3 <- abs(t - t2)  # Compute absolute differences from the median
    for (ii in 1:ncol(t3)) {
      p_value[ii] <- mean(t3[, ii] >= abs(0 - med[ii]))  # Compute p-value for two-sided test
    }
  } 
  # For two-sided tests with non-zero hypothesis
  else if (side == 2 & mean(hyp) != 0) {
    med <- apply(t, 2, median)  # Median of each column
    t2 <- matrix(rep(med, times = nrow(t)), nrow = nrow(t), byrow = TRUE)  # Repeat median values
    t3 <- abs(t - t2)  # Compute absolute differences
    for (ii in 1:ncol(t3)) {
      p_value[ii] <- mean(t3[, ii] >= abs(hyp[ii] - med[ii]))  # Compute p-value with non-zero hypothesis
    }
  }

  return(p_value)  # Return the matrix of p-values
}

# Full heterogeneity model function
# data: Data frame containing the dataset
# formula: The model formula to use
# res: Response variable
# group: Grouping variable
# var: Variable to examine
# sim2: Simulation method (permutation or bootstrap)
heter_full <- function(data, formula, res, group, var, sim2) {
  datau <- data
  groups <- levels(datau[[group]])  # Get unique group levels
  results_df <- data.frame()  # Initialize empty data frame to store results

  # Loop over each group and perform regression analysis
  for (g in groups) {
    subset_data <- datau[datau[[group]] == g, ]  # Subset data by group
    
    model <- tryCatch({
      lm(formula = formula, data = subset_data)  # Fit linear model
    }, error = function(e) NULL)

    if (is.null(model)) next  # Skip if model fitting fails

    model_summary <- summary(model)  # Summary of the model
    intercept <- model_summary$coefficients["(Intercept)", 1]  # Extract intercept
    residuals <- subset_data[, res] - predict(model, subset_data)  # Calculate residuals
    results = intercept + residuals  # Adjust results based on intercept

    results_df <- rbind(results_df, data.frame(Results = results, Group = g))  # Add to results data frame
  }

  # Fit model for estimated marginal means
  formula_str <- paste(res, "~ Group")
  emmeans_model <- lm(as.formula(formula_str), data = results_df)

  # Calculate estimated marginal means
  emmeans_result <- emmeans(emmeans_model, pairwise ~ Group)
  emmeans_summary <- summary(emmeans_result)  # Summary of the emmeans results
  emmeans_df <- as.data.frame(emmeans_summary$emmeans)  # Data frame of emmeans
  
  # Extract contrasts (comparisons between group means)
  contrasts_summary <- summary(emmeans_result$contrasts)
  contrasts_df <- as.data.frame(contrasts_summary)

  # Return both emmeans and contrasts data frames
  return(list(emmeans_df = emmeans_df, contrasts_df = contrasts_df))
}

# Heterogeneity model function for bootstrap and permutation
# data: Data frame
# i: Index for bootstrap/permutation
# formula: Model formula
# res: Response variable
# group: Grouping variable
# var: Variable for analysis
# method: Method used (either "permutation" or "bootstrap")
heter <- function(data, i, formula, res, group, var, method) {
  datau <- data
  groups <- levels(datau[[group]])  # Get unique group levels
  results_df <- data.frame()  # Data frame to store results

  # Loop over groups and compute residuals for each group
  for (g in groups) {
    subset_data <- datau[datau[[group]] == g, ]  # Subset data by group

    model <- tryCatch({
      lm(formula = formula, data = subset_data)  # Fit linear model
    }, error = function(e) NULL)

    if (is.null(model)) next  # Skip if model fitting fails

    model_summary <- summary(model)
    coef <- model_summary$coefficients[var, 1]  # Extract coefficient
    residuals <- subset_data[, res] - predict(model, subset_data)  # Calculate residuals
    results = coef * subset_data[, var] + residuals  # Adjust results based on coefficient

    results_df <- rbind(results_df, data.frame(Results = results, Group = g))  # Add to results
  }

  # Adjust results based on method (permutation or bootstrap)
  results_dfu <- results_df
  if (method == "permutation") {
    results_dfu[, res] <- results_df[i, res]  # Use permutation index
  } else {
    results_dfu <- results_df[i, ]  # Use bootstrap index
  }

  # Fit model and compute contrasts
  formula_str <- paste(res, "~ Group")
  emmeans_model <- lm(as.formula(formula_str), data = results_dfu)
  emmeans_result <- emmeans(emmeans_model, pairwise ~ Group)
  contrasts_summary <- summary(emmeans_result$contrasts)
  contrasts_df <- as.data.frame(contrasts_summary)

  # Return contrasts as numeric vector
  contrasts_vec <- c(t(contrasts_df[, c("estimate")]))
  return(as.numeric(contrasts_vec))
}

# Set seed for reproducibility
set.seed(725)  

# Load dataset from Excel file
data <- read_excel("C:/Users/XJTU/Desktop/roche.xlsx")

# Get the number of rows and sample a random index for bootstrap
l <- nrow(data)
i <- sample(1:l, replace = TRUE)

# Preprocess data
data$Sex <- as.factor(data$Sex)
data$Group <- as.factor(data$Group)

# Set simulation method
sim1 <- "permutation"
all_contrasts_df <- list()
all_emmeans_df <- list()

# List of response variables to analyze
res_all <- c("T2SUM", "T6SUM", "T10SUM", "T12SUM", "T13SUM", "T14SUM")

# Loop over each response variable and perform analysis
for (res in res_all) {
  formula <- as.formula(paste(res, "~ Sex"))  # Define the model formula
  result_full <- heter_full(data, formula, res, group, var, sim1)  # Full heterogeneity analysis

  # Perform bootstrapping using the heter function
  boot_result <- boot(
    data = data,
    statistic = heter,
    R = 1000,
    sim = sim1,
    formula = formula,
    res = res,
    group = group,
    var = var,
    method = sim1,
    parallel = "no"
  )

  # Extract results from bootstrap
  t0 <- boot_result$t0
  t <- boot_result$t
  p_diff <- pmake(t, t0)  # Calculate p-values

  # Get the emmeans and contrasts data frames
  emmeans_df <- result_full$emmeans_df
  contrasts_df <- result_full$contrasts_df

  # Convert p-values to data frame and add to contrasts
  p_diff_df <- as.data.frame(p_diff)
  p_diff_long <- as.data.frame(as.vector(t(p_diff_df)))
  colnames(p_diff_long) <- "p_diff"
  contrasts_df$p_diff <- p_diff_long$p_diff

  # Add response variable label
  emmeans_df$res <- res
  contrasts_df$res <- res

  # Append results to the lists
  all_emmeans_df[[res]] <- emmeans_df
  all_contrasts_df[[res]] <- contrasts_df
}

# Combine all the results into final data frames
final_emmeans_df <- do.call(rbind, all_emmeans_df)
final_contrasts_df <- do.call(rbind, all_contrasts_df)

# Export results to Excel file
write_xlsx(final_contrasts_df, "final_contrasts.xlsx")
write_xlsx(final_emmeans_df, "final_emmeans.xlsx")
