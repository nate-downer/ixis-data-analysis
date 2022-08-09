# Call Lobraries
library(tidyr)
library(dplyr)
library(lubridate)
library(tibble)
library(openxlsx)

# Import Data
getwd()
path <- "provided-reasources/DataAnalyst_Ecom_data_addsToCart.csv"
adds_data <- read.csv(path, header = TRUE, sep = ",")

path <- "provided-reasources/DataAnalyst_Ecom_data_sessionCounts.csv"
ses_data <- read.csv(path, header = TRUE, sep = ",")

# Extract datetime, format month names, and get the weekday number
for (i in seq_len(nrow(adds_data))) {
    month <- toString(adds_data[i, 2])
    year <- toString(adds_data[i, 1])
    dt <- mdy(paste(month, "/1/", year, sep = ""))
    adds_data[i, 4] <- dt
    adds_data[i, 5] <- format.Date(dt, "%Y-%m")
}

adds_data <- adds_data %>%
  rename(datetime = V4,
         month = V5)

for (i in seq_len(nrow(ses_data))) {
    dt <- mdy(ses_data[i, 3])
    ses_data[i, 7] <- dt
    ses_data[i, 8] <- format.Date(dt, "%Y-%m")
    ses_data[i, 9] <- wday(dt)
}

ses_data <- ses_data %>%
  rename(datetime = V7,
         month = V8,
         weekday_no = V9)

# Create a list of "user browsers" with browsers that were used in transactions
    # Note: The omitted browsers are likely bots or webscrappers, and should
    # not be counted as user sessions. Identifying and removing these sessions
    # will make the data more useful to the client

browsers <- ses_data %>%
    group_by(dim_browser) %>%
        summarise(all_sessions = sum(sessions),
                  transactions = sum(transactions),
                  QTY = sum(QTY),
                  .groups = "drop")

browsers <- browsers %>%
    filter(transactions > 0 & QTY > 0)

user_broswers <- browsers %>%
    pull(dim_browser)

clean <- ses_data %>%
    filter(dim_browser %in% user_broswers)

# Create Sheet 1:
# Group the data by month and device category
sheet1 <- ses_data %>%
    group_by(month, dim_deviceCategory) %>%
        summarise(all_sessions = sum(sessions),
                  transactions = sum(transactions),
                  QTY = sum(QTY),
                  .groups = "drop")

# Add a column containing the session data from user browsers
    # Note: Because of the method used to identify user browsers,
    # removing the non-user browsers only effects session data,
    # so alternate values for transactions and quantity
    # don't need to be included

user_sessions <- clean %>%
    group_by(month, dim_deviceCategory) %>%
        summarise(user_sessions = sum(sessions),
        .groups = "drop")

sheet1 <- merge(sheet1, user_sessions, by = c("month", "dim_deviceCategory"))

# Add an ECR column
sheet1 <- sheet1 %>%
    mutate(ECR = transactions / user_sessions)

# Re-Order columns for readability
sheet1 <- sheet1 %>%
    select(month,
           dim_deviceCategory,
           transactions,
           QTY,
           all_sessions,
           user_sessions,
           ECR)

# Rename columns for easier reading in Excel
sheet1 <- sheet1 %>%
    rename(qty = QTY,
           ecr = ECR,
           device_category = dim_deviceCategory)

# Create Sheet 2:
# Group the data by month
sheet2 <- ses_data %>%
    group_by(month) %>%
        summarise(all_sessions = sum(sessions),
                  transactions = sum(transactions),
                  QTY = sum(QTY),
                  .groups = "drop")

# Merge in the adds data
sheet2 <- merge(sheet2, select(adds_data, month, addsToCart), by = "month")

# Add a column containing the session data from user browsers
user_sessions <- clean %>%
    group_by(month) %>%
        summarise(user_sessions = sum(sessions),
        .groups = "drop")

sheet2 <- merge(sheet2, user_sessions, by = "month")

sheet2 <- sheet2 %>%
    relocate(all_sessions, .before = user_sessions)

# Add an ECR column
sheet2 <- sheet2 %>%
    mutate(ECR = transactions / user_sessions)

# Get the data concerning the two most recent months
sheet2 <- sheet2 %>%
    arrange(desc(month))

sheet2 <- sheet2[1:2, ]

# Format the sheet to display month names and differneces
sheet2 <- sheet2 %>%
    remove_rownames %>%
        column_to_rownames("month")

sheet2 <- sheet2 %>%
    select(transactions,
           QTY,
           all_sessions,
           user_sessions,
           ECR,
           addsToCart)

sheet2 <- sheet2 %>%
    rename(qty = QTY,
           ecr = ECR,
           adds_to_cart = addsToCart)

sheet2 <- as.data.frame(t(sheet2))

for (i in seq_len(nrow(sheet2))) {
    abs <- sheet2[i, 1] - sheet2[i, 2]
    rel <- abs / sheet2[i, 2]
    sheet2[i, 3] <- abs
    sheet2[i, 4] <- rel
}

sheet2 <- sheet2 %>%
  rename(absolute_dif = V3,
         relative_dif = V4)

sheet2 <- tibble::rownames_to_column(sheet2, "metric")

# Note: Sheets 3 and 4 are additional analysis
# beyond the original scope of the assignment

# Create Sheet 3:
# Group the session data by day of the week
sheet3 <- ses_data %>%
    group_by(weekday_no) %>%
        summarise(all_sessions = mean(sessions),
                  transactions = mean(transactions),
                  QTY = mean(QTY),
                  .groups = "drop")

# Add a column containing the session data from user browsers
user_sessions <- clean %>%
    group_by(weekday_no) %>%
        summarise(user_sessions = mean(sessions),
        .groups = "drop")

sheet3 <- merge(sheet3, user_sessions, by = "weekday_no")

# Add ECR column
sheet3 <- sheet3 %>%
    mutate(ECR = transactions / user_sessions)

# Reformat the days of the week for readability
weekdays <- c("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")

for (i in seq_len(nrow(sheet3))) {
    sheet3[i, 7] <- weekdays[i]
}

sheet3 <- sheet3 %>%
  rename(weekday = V7)

sheet3 <- sheet3 %>%
    select(weekday,
           transactions,
           QTY,
           all_sessions,
           user_sessions,
           ECR)

# Rename columns for easier reading in Excel
sheet3 <- sheet3 %>%
    rename(qty = QTY,
           ecr = ECR)

# Create Sheet 4:
# Group the session data by browser
    # Note: Because browsers were used to distinguish between
    # 'user sessions' and non-user sessions, that distinction
    # has no meaning when the data is grouped by browser

sheet4 <- ses_data %>%
    group_by(dim_browser) %>%
        summarise(all_sessions = sum(sessions),
                  transactions = sum(transactions),
                  QTY = sum(QTY),
                  .groups = "drop")

sheet4 <- sheet4 %>%
    arrange(desc(all_sessions))

# Add ECR column
sheet4 <- sheet4 %>%
    mutate(ECR = transactions / all_sessions)

sheet4 <- sheet4 %>%
    arrange(desc(all_sessions))

# Add a column indicating which browsers are included in user_sessions data
for (i in seq_len(nrow(sheet4))) {
    if (sheet4[i, 3] > 0 && sheet4[i, 4] > 0) {
        sheet4[i, 6] <- TRUE
    } else {
        sheet4[i, 6] <- FALSE
    }
}

sheet4 <- sheet4 %>%
    rename(browser_in_user_sessions = ...6)

sheet4 <- sheet4 %>%
    select(dim_browser,
           all_sessions,
           transactions,
           QTY,
           ECR,
           browser_in_user_sessions)

# Rename columns for easier reading in Excel
sheet4 <- sheet4 %>%
    rename(browser = dim_browser,
           qty = QTY,
           ecr = ECR)

# Export data to Excel
sheets_list <- list("Volume by Month + Device" = sheet1,
                    "Month Over Month Comparison" = sheet2,
                    "Ave Volume By Weekday" = sheet3,
                    "Volume By Browser" = sheet4)

write.xlsx(sheets_list, file = "r_worksheets.xlsx")