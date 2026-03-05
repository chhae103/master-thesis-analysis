# SPIDERWEB-PLOT SCRIPT – TREATMENT COMPARISON

# Funktion zur Installation und Laden von Paketen
install_and_load <- function(pkg) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    message(paste("Installing package:", pkg))
    install.packages(pkg, dependencies = TRUE)
  }
  library(pkg, character.only = TRUE)
}

# Pakete installieren und laden
required_packages <- c("readxl", "dplyr", "openxlsx", "tidyr", "fmsb")
lapply(required_packages, install_and_load)

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

# Daten laden und vorbereiten

data <- read.xlsx(file_path)
colnames(data) <- trimws(colnames(data))
colnames(data) <- gsub("\\.", " ", colnames(data))

required_columns <- c("Speed", "Area", "Dist per bend", "Bends per 30s", "Max speed")

# Nur Aerosil-Gruppen
concentrations <- c(
  "Aerosil200 1 µg/ml",
  "Aerosil200 10 µg/ml",
  "Aerosil200 100 µg/ml",
  "Aerosil200 200 µg/ml",
  "Aerosil200 500 µg/ml"
)

data_filtered <- data %>%
  filter(Group %in% concentrations)

# Mittelwerte berechnen
calculate_mean_values <- function(data) {
  data_clean <- data %>%
    select(Group, all_of(required_columns)) %>%
    drop_na()
  
  data_clean %>%
    group_by(Group) %>%
    summarise(across(all_of(required_columns), mean, na.rm = TRUE), .groups = 'drop')
}

mean_data <- calculate_mean_values(data_filtered)

# Ergebnisse speichern
write.xlsx(mean_data, new_file_path)
cat("✅ Datei gespeichert: Spider_Web_Treatment_Comparison_Data.xlsx\n")

# Spider-Web-Plot vorbereiten

data <- read_excel(new_file_path)

# Existierende Spalten robust auswählen
existing_params <- intersect(required_columns, colnames(data))

max_values <- c(0.3, 80, 0.3, 150, 0.5)[1:length(existing_params)]

# Farben definieren (Graustufen)
colors <- c(
  "Aerosil200 1 µg/ml" = "grey70",
  "Aerosil200 10 µg/ml" = "grey60",
  "Aerosil200 100 µg/ml" = "grey40",
  "Aerosil200 200 µg/ml" = "grey20",
  "Aerosil200 500 µg/ml" = "grey10"
)

# Linientypen definieren
line_types <- c(
  "Aerosil200 1 µg/ml" = 1,
  "Aerosil200 10 µg/ml" = 2,
  "Aerosil200 100 µg/ml" = 3,
  "Aerosil200 200 µg/ml" = 4,
  "Aerosil200 500 µg/ml" = 1
)

# Punkt-Symbole definieren
point_shapes <- c(
  "Aerosil200 1 µg/ml" = 16,  # Kreis ausgefüllt
  "Aerosil200 10 µg/ml" = 1,  # Kreis offen
  "Aerosil200 100 µg/ml" = 17, # Dreieck
  "Aerosil200 200 µg/ml" = 15, # Quadrat
  "Aerosil200 500 µg/ml" = 16  # Kreis ausgefüllt
)

# Daten für Radar vorbereiten
df_list <- lapply(concentrations, function(group) {
  data[data$Group == group, existing_params, drop = FALSE]
})

df <- do.call(rbind, df_list)
df <- rbind(max_values, rep(0, length(existing_params)), df)
colnames(df) <- existing_params
rownames(df) <- c("Max", "Min", concentrations)

colors_vector <- colors[concentrations]
lty_vector <- line_types[concentrations]
pch_vector <- point_shapes[concentrations]

# Spider-Web-Plot zeichnen

png(
  filename = paste0(output_path, "Spider_Web_Treatment_Comparison.png"),
  width = 3000,
  height = 3000,
  res = 600,
  units = "px"
)

op <- par(mar = c(7, 4, 4, 2))  # mehr Platz unten

# Linienplot
fmsb::radarchart(
  df,
  axistype = 1,
  pcol = colors_vector,
  plwd = 2,
  plty = lty_vector,
  cglcol = "grey60",
  cglty = 1,
  cglwd = 1.5,
  axislabcol = "grey30",
  calcex = 0.5,
  vlcex = 0.6,
  caxislabels = rep("", length(existing_params)),
  legend = FALSE,
  pty = pch_vector
)

# Legende – korrigierte Version mit offenen Symbolen

legend(
  "bottom",
  inset = -0.3,
  legend = names(colors_vector),
  col = colors_vector,
  lty = lty_vector,
  lwd = 2,
  pch = pch_vector,
  pt.bg = ifelse(pch_vector == 1, NA, colors_vector),  # offene Kreise bleiben ungefüllt
  pt.cex = 1.2,   # Symbole etwas größer
  cex = 0.4,
  horiz = TRUE,
  xpd = TRUE,
  bty = "n",
  text.col = "black"
)

par(op)
dev.off()

cat("✅ Spider-Web-Plot gespeichert: Spider_Web_Treatment_Comparison.png\n")
