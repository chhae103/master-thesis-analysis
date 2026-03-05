# SCATTER-PLOT COMPARISON SCRIPT (All Treatments Only)

# --- Install and load required packages ---
install_and_load <- function(pkg) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    message(paste("Installing package:", pkg))
    install.packages(pkg, dependencies = TRUE)
  } else {
    message(paste("Package already installed:", pkg))
  }
  library(pkg, character.only = TRUE)
}

required_packages <- c("readxl", "dplyr", "ggplot2", "stringr", "openxlsx")
for (pkg in required_packages) install_and_load(pkg)

# Define paths

# Input-Datei (dynamisch generiert)
input_file <- file.path(
  data_path,
  glue("Experiment_{day}"),
  glue("Results_{experiment}"),
  glue("{run}_{experiment}.xlsx")
)

# Output-Ordner erstellen, falls nicht vorhanden
output_folder <- file.path(data_path, "Ergebnisse")
if (!dir_exists(output_folder)) dir_create(output_folder)

# Output-Datei (dynamisch generiert)
output_file <- file.path(
  output_folder,
  glue("Statistics_Results_{experiment}.xlsx")
)


if (!dir.exists(output_path)) dir.create(output_path, recursive = TRUE)

# Load and clean data

data <- read_excel(file_path, sheet = "Data")

# Clean column names
colnames(data) <- trimws(colnames(data))
colnames(data) <- gsub("\\.", " ", colnames(data))

# Required columns
required_columns <- c("Bends per 30s", "Speed", "Max speed", "Dist per bend")

missing_cols <- setdiff(required_columns, colnames(data))
if (length(missing_cols) > 0) stop("Missing columns: ", paste(missing_cols, collapse = ", "))

# Define groups, colors, and shapes

groups <- c(
  "H2O",
  "Aerosil200 1 µg/ml",
  "Aerosil200 10 µg/ml",
  "Aerosil200 100 µg/ml",
  "Aerosil200 200 µg/ml",
  "Aerosil200 500 µg/ml"
)

palette <- c(
  "H2O" = "blue",
  "Aerosil200 1 µg/ml" = "grey70",
  "Aerosil200 10 µg/ml" = "grey60",
  "Aerosil200 100 µg/ml" = "grey40",
  "Aerosil200 200 µg/ml" = "grey20",
  "Aerosil200 500 µg/ml" = "grey10"
)

# Gleiche Shapes wie im Spiderweb-Plot
shapes <- c(
  "H2O" = 16,               # ausgefüllter Kreis
  "Aerosil200 1 µg/ml" = 16, # ausgefüllter Kreis
  "Aerosil200 10 µg/ml" = 1, # offener Kreis
  "Aerosil200 100 µg/ml" = 17, # Dreieck
  "Aerosil200 200 µg/ml" = 15, # Quadrat
  "Aerosil200 500 µg/ml" = 16  # ausgefüllter Kreis
)

data <- data %>%
  filter(Group %in% groups) %>%
  mutate(Group = factor(Group, levels = groups))

# Create combined scatterplot for all treatments (excluding H2O)

# Filter all treatments except H2O
data_treatments <- data %>% filter(Group != "H2O")

# Parameters to plot
x_param <- "Bends per 30s"
y_param <- "Speed"

# Create scatterplot
p_combined <- ggplot(
  data_treatments,
  aes(
    x = .data[[x_param]],
    y = .data[[y_param]],
    color = Group,
    shape = Group
  )
) +
  geom_point(alpha = 0.9, size = 2.5, stroke = 0.8) +
  scale_color_manual(values = palette) +
  scale_shape_manual(values = shapes) +
  coord_cartesian(ylim = c(NA, 0.5)) +
  theme_minimal(base_size = 12) +
  theme(
    panel.grid = element_blank(),
    axis.line = element_line(color = "black"),
    legend.position = "top",
    legend.title = element_blank(),
    legend.text = element_text(size = 3),          # Schriftgröße der Legende
    legend.key.size = unit(0.4, "lines"),          # Symbolgröße
    legend.spacing.x = unit(0.25, "cm"),
    legend.margin = margin(t = 0, b = 0),
    axis.text = element_text(color = "black"),
    plot.title = element_text(hjust = 0.5, size = 13)
  ) +
  labs(x = x_param, y = y_param)
# + ggtitle("All Treatments (excluding H2O)")

# Save plot

ggsave(
  filename = file.path(output_path, "Scatterplots_Treatment_Triangle_Square.png"),
  plot = p_combined,
  width = 10,
  height = 8,
  units = "cm",
  dpi = 600
)

cat("\n✅ Combined scatterplot (all treatments) saved successfully as 'Scatterplots_Only_Treatment.png' in:\n", output_path, "\n")
