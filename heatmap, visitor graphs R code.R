library(readxl)
library(dplyr)
library(ggplot2)
library(tidyr)
library(scales)

# Read the Excel file
file_path <- "data/2025 Allianz Datathon Dataset.xlsx"

# Read the Visitation Data sheet
visitation_data <- read_excel(file_path, sheet = "Visitation Data")

# Reshape data from wide to long format
long_data <- visitation_data %>%
  select(Year, Week, `Mt. Baw Baw`, `Mt. Stirling`, `Mt. Hotham`, `Falls Creek`, 
         `Mt. Buller`, Selwyn, Thredbo, Perisher, `Charlotte Pass`) %>%
  pivot_longer(
    cols = -c(Year, Week),
    names_to = "Resort",
    values_to = "Visitors"
  ) %>%
  filter(Week <= 15)  # Only include weeks 1-15

# Calculate weekly averages for each resort
weekly_averages <- long_data %>%
  group_by(Resort, Week) %>%
  summarise(Average_Visitors = mean(Visitors, na.rm = TRUE),
            .groups = 'drop') %>%
  arrange(Resort, Week)

# Create a faceted bar plot for all resorts (NO numbers on bars)
ggplot(weekly_averages, aes(x = factor(Week), y = Average_Visitors, fill = Resort)) +
  geom_bar(stat = "identity", alpha = 0.8) +
  facet_wrap(~ Resort, scales = "free_y", ncol = 3) +
  labs(title = "Average Weekly Visitors by Resort (2014-2024)",
       subtitle = "Based on historical data across ski seasons (Weeks 1-15)",
       x = "Week of Ski Season",
       y = "Average Number of Visitors",
       caption = "Source: 2025 Allianz Datathon Dataset") +
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 16),
    plot.subtitle = element_text(hjust = 0.5, size = 12),
    axis.title = element_text(face = "bold"),
    axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5),
    axis.text.y = element_text(size = 8),
    strip.text = element_text(face = "bold", size = 10),
    legend.position = "none",
    panel.spacing = unit(1.5, "lines")
  ) +
  scale_y_continuous(labels = comma) +
  scale_fill_brewer(palette = "Set3")

# Clean line plot without clutter
ggplot(weekly_averages, aes(x = Week, y = Average_Visitors, color = Resort, group = Resort)) +
  geom_line(linewidth = 1.2) +
  geom_point(size = 2) +
  labs(title = "Comparison of Average Weekly Visitors Across Resorts",
       subtitle = "Seasonal patterns across ski weeks",
       x = "Week of Ski Season",
       y = "Average Number of Visitors",
       color = "Resort") +
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 16),
    plot.subtitle = element_text(hjust = 0.5, size = 12),
    axis.title = element_text(face = "bold"),
    legend.position = "bottom",
    legend.text = element_text(size = 9)
  ) +
  scale_y_continuous(labels = comma, trans = "log10") +
  scale_x_continuous(breaks = 1:15)

# Heatmap WITH numerical labels (kept as requested)
ggplot(weekly_averages, aes(x = factor(Week), y = Resort, fill = Average_Visitors)) +
  geom_tile(color = "white", linewidth = 0.5) +
  geom_text(aes(label = comma(round(Average_Visitors, 0))), 
            color = "white", size = 2.5, fontface = "bold") +
  scale_fill_gradient(low = "blue", high = "red", 
                      trans = "log10",
                      labels = comma,
                      name = "Average\nVisitors") +
  labs(title = "Heatmap of Average Weekly Visitors by Resort",
       subtitle = "Darker red colors indicate higher visitor numbers",
       x = "Week of Ski Season",
       y = "Resort") +
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold", size = 16),
    plot.subtitle = element_text(hjust = 0.5, size = 12),
    axis.title = element_text(face = "bold"),
    axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5),
    panel.grid = element_blank()
  )

# Display summary statistics in a clean table
summary_stats <- weekly_averages %>%
  group_by(Resort) %>%
  summarise(
    Total_Avg_Visitors = sum(Average_Visitors),
    Peak_Week = Week[which.max(Average_Visitors)],
    Peak_Visitors = max(Average_Visitors),
    Low_Week = Week[which.min(Average_Visitors)],
    Low_Visitors = min(Average_Visitors),
    .groups = 'drop'
  ) %>%
  arrange(desc(Total_Avg_Visitors))

# Format the summary table nicely
summary_stats_formatted <- summary_stats %>%
  mutate(across(where(is.numeric), ~round(., 0))) %>%
  mutate(Total_Avg_Visitors = comma(Total_Avg_Visitors),
         Peak_Visitors = comma(Peak_Visitors),
         Low_Visitors = comma(Low_Visitors))

print("Summary Statistics by Resort (Ordered by Total Average Visitors):")
print(summary_stats_formatted)

# Also show the raw weekly averages for reference
cat("\nWeekly Averages Table (first few rows):\n")
print(head(weekly_averages, 20))

# interpretation:
# number of visitors per week relative to a resort itself 
# is correlated with its distance from metro melbourne (less so)
# and its resort size (number of runs, hotels... etc)

# therefore the middle range of the average visitors per week is ideal,
# not too many ppl, whereas lower visitors are likely due to unfavourable skiing 
# conditions




