# SETTING UP R

# Install necessary libraries
install.packages("readxl")         # For reading Excel files
install.packages("prophet")        # For time series forecasting
install.packages("ggplot2")        # For visualization
install.packages("dplyr")          # For data manipulation
install.packages("zoo")            # For handling missing data and interpolation
install.packages("reshape2")       # For reshaping data

# Import necessary libraries
library(readxl)
library(prophet)
library(ggplot2)
library(dplyr)
library(zoo)
library(reshape2)


# LOADING THE DATA

# Load the Excel file without headers
df <- read_excel("ARW_structured.xlsx", col_names = FALSE)

# Preview the data
print(head(df))


# DATA PROCESSING

# Extract dates from the first row (excluding the first column)
dates <- as.character(unlist(df[1, -1]))

# Extract line items and their values from the remaining rows
line_items <- list()
for (i in 2:14) {  # Loop through rows 2 to 14
  line_item <- df[[i, 1]]  # First column (line item name)
  values <- as.numeric(unlist(df[i, -1]))  # Remaining columns (values)
  line_items[[i - 1]] <- list(line_item = line_item, values = values)  # Append to list
}

# Print extracted line items
cat("Extracted Line Items:\n")
for (item in line_items) {
  cat(item$line_item, "\n")
}

# Print information about the data frame
cat("Structure of the DataFrame:\n")
str(df)

cat("\nSummary of the DataFrame:\n")
summary(df)

# Transpose the data frame
df_transposed <- as.data.frame(t(df[-1]))  # Exclude the first column (headers) and transpose
colnames(df_transposed) <- df[[1]]         # Set column names from the first column of the original dataset
rownames(df_transposed) <- NULL            # Remove row names for cleanliness

# Remove the first row (original column headers/dates) and reset as numeric
df_transposed <- df_transposed[-1, ]  # Select rows from index 2 onwards

# Convert all columns to numeric
df_transposed <- as.data.frame(lapply(df_transposed, function(col) as.numeric(as.character(col))))

# Calculate basic statistics for the whole dataset
basic_stats <- summary(df_transposed)

# Print basic statistics
print(basic_stats)


# CORRELATION ANALYSIS

# Convert the extracted line_items into a data frame
data_for_correlation <- data.frame(
  matrix(unlist(lapply(line_items, `[[`, "values")), 
         ncol = length(line_items), byrow = FALSE),
  stringsAsFactors = FALSE
)

# Assign column names based on line item names
colnames(data_for_correlation) <- sapply(line_items, `[[`, "line_item")

# Assign row names based on dates
rownames(data_for_correlation) <- dates

# Display the head of the data frame
print(head(data_for_correlation))

# Calculate the correlation matrix
correlation_matrix <- cor(data_for_correlation, use = "pairwise.complete.obs")

# Reshape the correlation matrix for ggplot2
correlation_melt <- melt(correlation_matrix)

# Generate the heatmap
ggplot(data = correlation_melt, aes(Var2, Var1, fill = value)) +  # Flip axes for alignment
  geom_tile(color = "white") +
  scale_fill_gradient2(
    low = "blue", mid = "white", high = "red", 
    midpoint = 0.5, limits = c(0, 1), 
    name = "Correlation"
  ) +
  geom_text(aes(label = sprintf("%.2f", value)), color = "black", size = 3) +  # Add annotations
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1, size = 8),
    axis.text.y = element_text(size = 8),
    plot.title = element_text(size = 14, hjust = 0.5),
    panel.grid = element_blank()  # Remove grid lines for clarity
  ) +
  labs(
    title = "Correlation Heatmap of Financial Line Items (0 to 1)",
    x = "Line Items",
    y = "Line Items"
  )


# DATA EXTRACTION

# Initialize an empty list to store data frames for each line item
data_frames <- list()

# Extract dates from the dataset (assuming the first row contains dates)
dates <- as.Date(unlist(df[1, -1]), format = "%d-%m-%Y")  # Ensure proper date format

# Loop through each line item (rows 2 to 14 in the original data)
for (i in 2:nrow(df)) {
  # Extract line item name and values
  line_item <- as.character(df[i, 1])  # First column contains the line item names
  values <- as.numeric(as.character(df[i, -1]))  # Exclude the first column and convert to numeric
  
  # Create a data frame for the line item
  data <- data.frame(
    ds = dates,
    y = values,
    line_item = line_item,
    stringsAsFactors = FALSE
  )
  
  # Handle missing values using interpolation
  data$y <- na.approx(data$y, na.rm = FALSE, rule = 2)  # Linear interpolation (forward/backward fill)
  
  # Custom imputation for "Interest Expense" using linear regression
  if (line_item == "Interest Expense") {
    # Assuming "Revenue" data is already created
    revenue_data <- data_frames[["Revenue"]]
    
    if (!is.null(revenue_data) && sum(!is.na(revenue_data$y)) > 0) {
      # Merge non-NA data from Interest Expense and Revenue
      valid_data <- na.omit(data) %>%
        inner_join(na.omit(revenue_data), by = "ds", suffix = c("_interest", "_revenue"))
      
      if (nrow(valid_data) > 0) {
        # Fit a linear regression model
        model <- lm(y_interest ~ y_revenue, data = valid_data)
        
        # Predict missing values for Interest Expense based on Revenue
        missing_indices <- which(is.na(data$y))
        if (length(missing_indices) > 0) {
          data$y[missing_indices] <- predict(
            model, 
            newdata = data.frame(y_revenue = revenue_data$y[missing_indices])
          )
        }
      }
    }
  }
  
  # Skip line items where all values are NA
  if (!all(is.na(data$y))) {
    # Add the processed data frame to the list
    data_frames[[line_item]] <- data
  }
}

# Check the resulting data frames
str(data_frames)  # Inspect the structure of the list


# TIME SERIES MODELING WITH PROPHET

# PROHPET MODEL

# Initialize a list to store Prophet models
models <- list()

# Fit a Prophet model for each line item
for (i in seq_along(data_frames)) {
  line_item_data <- data_frames[[i]]
  line_item_name <- names(data_frames)[i]
  
  # Initialize Prophet model with adjusted parameters
  model <- prophet(
    changepoint.prior.scale = 0.05,
    seasonality.prior.scale = 10.0,
    holidays.prior.scale = 20.0,
    weekly.seasonality = FALSE,
    daily.seasonality = FALSE
  )
  
  # Dynamically add country holidays
  model <- add_country_holidays(model, country_name = "US")
  
  # Handle errors during model fitting
  tryCatch(
    {
      model <- fit.prophet(model, line_item_data)
      models[[line_item_name]] <- model  # Store the model in the list
    },
    error = function(e) {
      cat("An error occurred while fitting the model for", line_item_name, ":\n", e$message, "\n")
    }
  )
}

# Generate future dates for prediction (next 5 years, quarterly)
future_dates <- data.frame(
  ds = seq.Date(from = as.Date("2024-01-01"), by = "quarter", length.out = 5 * 4)  # 5 years of quarterly dates
)

# Print the future dates for reference
print(future_dates)


# FORECASTING

# Initialize a list to store forecast results for each line item
forecast_results <- list()

# Generate future forecasts for each line item
for (i in seq_along(models)) {
  line_item_name <- names(models)[i]
  model <- models[[line_item_name]]
  
  # Generate forecast
  forecast <- predict(model, future_dates)
  
  # Store forecast data along with line item name
  forecast_results[[line_item_name]] <- list(
    line_item = line_item_name,
    forecast = forecast[, c("ds", "yhat", "yhat_lower", "yhat_upper")]
  )
}

# Print the forecasted values for the next 5 years (quarterly) for each line item
cat("\nForecasted Line Items for the Next 5 Years (Quarterly):\n\n")
for (result in forecast_results) {
  cat(result$line_item, "Forecast:\n")
  print(tail(result$forecast, 20))
  cat("\n")
}


# VISUALIZATION

# Loop through the forecast results to generate plots
for (line_item_name in names(forecast_results)) {
  forecast <- forecast_results[[line_item_name]]$forecast
  historical_data <- data_frames[[line_item_name]]
  
  # Combine historical and forecasted data for plotting
  forecast$Type <- "Forecast"
  historical_data$Type <- "Historical"
  
  combined_data <- rbind(
    data.frame(ds = historical_data$ds, y = historical_data$y, Type = historical_data$Type),
    data.frame(ds = forecast$ds, y = forecast$yhat, Type = forecast$Type)
  )
  
  # Plot historical and forecasted data
  p <- ggplot(data = combined_data, aes(x = ds, y = y, color = Type)) +
    geom_line(linewidth = 1) +  # Use linewidth instead of size
    geom_point(size = 2) +
    scale_color_manual(values = c("Historical" = "blue", "Forecast" = "orange")) +
    labs(
      title = paste(line_item_name, "Historical and Forecast for Next 5 Years (Quarterly)"),
      x = "Date",
      y = "Value",
      color = "Type"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14, face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1),
      legend.position = "top"
    )
  
  # Display the plot
  print(p)
}


# FREE CASH FLOW TO FIRM

# Tax Rate Adjustment
tax_rate <- 0.21  # Adjust the tax rate to 21%

# Calculate Historical FCFF
fcff_data <- data.frame(
  Net_Income = data_frames[["Operating Income"]]$y - data_frames[["Taxes"]]$y,
  D_A = data_frames[["D&A"]]$y,
  Interest_Expense = data_frames[["Interest Expense"]]$y,
  Change_in_Working_Capital = c(NA, diff(
    data_frames[["Total Current Assets"]]$y - data_frames[["Total Current Liabilities"]]$y
  )),
  Cap_Ex = data_frames[["Cap-Ex"]]$y
)

# Compute FCFF
fcff_data$FCFF <- with(fcff_data, 
                       Net_Income + D_A + Interest_Expense * (1 - tax_rate) - Change_in_Working_Capital - Cap_Ex
)

# Add dates to FCFF data
fcff_data$Date <- data_frames[["Operating Income"]]$ds

# Predicted FCFF
predicted_fcff <- data.frame(
  Date = future_dates$ds,  # Future dates for prediction
  Net_Income = forecast_results[["Operating Income"]]$forecast$yhat - 
    forecast_results[["Taxes"]]$forecast$yhat,
  D_A = forecast_results[["D&A"]]$forecast$yhat,
  Interest_Expense = forecast_results[["Interest Expense"]]$forecast$yhat,
  Change_in_Working_Capital = c(NA, diff(
    forecast_results[["Total Current Assets"]]$forecast$yhat - 
      forecast_results[["Total Current Liabilities"]]$forecast$yhat
  )),
  Cap_Ex = forecast_results[["Cap-Ex"]]$forecast$yhat
)

# Compute Predicted FCFF
predicted_fcff$FCFF <- with(predicted_fcff, 
                            Net_Income + D_A + Interest_Expense * (1 - tax_rate) - Change_in_Working_Capital - Cap_Ex
)

# Combine Historical and Forecasted FCFF
combined_fcff <- rbind(
  data.frame(Date = fcff_data$Date, FCFF = fcff_data$FCFF, Type = "Historical"),
  data.frame(Date = predicted_fcff$Date, FCFF = predicted_fcff$FCFF, Type = "Forecast")
)

# Filter out NA values before plotting
combined_fcff <- combined_fcff[!is.na(combined_fcff$FCFF), ]

# Plot Historical and Predicted FCFF
ggplot(combined_fcff, aes(x = Date, y = FCFF, color = Type)) +
  geom_line(linewidth = 1) +
  geom_point(size = 2) +
  scale_color_manual(values = c("Historical" = "blue", "Forecast" = "orange")) +
  labs(
    title = "Historic and Predicted Free Cash Flow to Firm (FCFF)",
    x = "Date",
    y = "FCFF",
    color = "Type"
  ) +
  theme_minimal() +
  theme(
    plot.title = element_text(size = 16, hjust = 0.5),
    axis.text.x = element_text(angle = 45, hjust = 1),
    legend.position = "top"
  )


# ADDITIONAL ANALYSIS

# Load required libraries
library(gridExtra)
library(grid)

# Directory to save images
output_dir <- "forecasted_images"
if (!dir.exists(output_dir)) dir.create(output_dir)

# Loop through forecast results and generate improved images
for (line_item_name in names(forecast_results)) {
  # Extract forecasted values for the current line item
  forecast <- forecast_results[[line_item_name]]$forecast
  
  # Create a data frame with relevant columns
  forecast_table <- data.frame(
    Date = as.character(forecast$ds), 
    yhat = forecast$yhat, 
    yhat_lower = forecast$yhat_lower, 
    yhat_upper = forecast$yhat_upper
  )
  
  # Create a table grob
  table_grob <- tableGrob(
    forecast_table, 
    rows = NULL, # Remove row names
    theme = ttheme_minimal(
      core = list(bg_params = list(fill = "white"), fg_params = list(hjust = 0.5, fontsize = 10)),
      colhead = list(fg_params = list(fontface = "bold", fontsize = 12))
    )
  )
  
  # Add a title (caption) below the table
  caption <- textGrob(
    paste(line_item_name, "Forecast"),
    gp = gpar(fontsize = 14, fontface = "italic", col = "black"),
    hjust = 0.5
  )
  
  # Combine the table and the caption
  full_table <- gtable::gtable_add_rows(
    table_grob, 
    heights = unit(1, "cm"), 
    pos = -1
  )
  full_table <- gtable::gtable_add_grob(
    full_table, 
    caption, 
    t = nrow(full_table), l = 1, r = ncol(full_table)
  )
  
  # Save the combined table as a high-quality image
  output_path <- file.path(output_dir, paste0(line_item_name, "_forecast.png"))
  png(output_path, width = 650, height = 950, res = 150) # Higher resolution and appropriate aspect ratio
  grid::grid.draw(full_table)
  dev.off()
}

cat("Improved images of forecasted values saved in:", output_dir, "\n")

