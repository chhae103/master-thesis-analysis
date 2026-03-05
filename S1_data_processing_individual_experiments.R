suppressWarnings({
  # Pakete installieren und laden
  if (!require("pacman")) install.packages("pacman")
  pacman::p_load("readr", "dplyr", "purrr", "stringr", "openxlsx", "fs", "writexl", "ggplot2", "readxl", "future", "furrr")
  
  # Parallelisierung aktivieren
  plan(multisession)
  
  # Hauptordner und Präfixe festlegen
  all_folders <- c("D2", "D4", "D10")
  durchlaeufe <- c("Experiment 1")
  
  data_path <- "data/"
  
  # Zielordner definieren
  target_folder <- file.path(data_path, "Ergebnisse")
  if (!dir_exists(target_folder)) dir_create(target_folder)
  
  # Template für die zu durchsuchenden Ordner
  folder_pairs_template <- list( 
    list(source = file.path(data_path, "{durchlauf}", "{day}", "Output H2O"), prefix = "H2O"),
    list(source = file.path(data_path, "{durchlauf}", "{day}", "Aerosil200 1"), prefix = "Aerosil200 1"),
    list(source = file.path(data_path, "{durchlauf}", "{day}", "Aerosil200 10"), prefix = "Aerosil200 10"),
    list(source = file.path(data_path, "{durchlauf}", "{day}", "Aerosil200 100"), prefix = "Aerosil200 100"),
    list(source = file.path(data_path, "{durchlauf}", "{day}", "Aerosil200 200"), prefix = "Aerosil200 200"),
    list(source = file.path(data_path, "{durchlauf}", "{day}", "Aerosil200 500"), prefix = "Aerosil200 500")
  )
  
  
  
  neue_namen <- c("n", "BPM", "Bends per 30s", "Speed", "Max speed", "Dist per bend", "Area", "Appears in frames", "Moving (non-paralyzed)", "Region", "Round ratio", "Eccentricity")
  
  # Funktion zur Datenverarbeitung
  verarbeite_daten <- function(folder_info) {
    cat("Verarbeitung: ", folder_info$source, "\n")
    if (!dir_exists(folder_info$source)) {
      cat("Verzeichnis existiert nicht: ", folder_info$source, "\n")
      return(NULL)
    }
    
    daten_liste <- dir_ls(folder_info$source, recurse = TRUE, glob = "*.csv") %>%
      future_map_dfr(~ read_csv(., col_types = cols()) %>% setNames(neue_namen)) %>%
      mutate(
        Group = folder_info$prefix,
        n = row_number(),
        Untersuchungstag = str_extract(folder_info$source, "D[0-9]+")
      )
    return(daten_liste)
  }
  
  # Funktion für Statistik
  berechne_statistiken <- function(daten) {
    daten %>%
      group_by(Group) %>%
      summarise(
        n = n(),
        Mittelwert = mean(`Bends per 30s`, na.rm = TRUE),
        SD = sd(`Bends per 30s`, na.rm = TRUE),
        Median = median(`Bends per 30s`, na.rm = TRUE),
        Min = min(`Bends per 30s`, na.rm = TRUE),
        Max = max(`Bends per 30s`, na.rm = TRUE)
      ) %>%
      ungroup()
  }
  
  # Hauptverarbeitungsschleife
  future_walk(durchlaeufe, function(durchlauf) {
    future_walk(all_folders, function(day) {
      # Pfade ersetzen
      folder_pairs <- lapply(folder_pairs_template, function(fp) {
        fp$source <- str_replace_all(fp$source, c("\\{durchlauf\\}" = durchlauf, "\\{day\\}" = day))
        return(fp)
      })
      
      alle_daten_tag <- bind_rows(lapply(folder_pairs, verarbeite_daten))
      
      if (is.null(alle_daten_tag) || nrow(alle_daten_tag) == 0) {
        cat("Keine Daten für den Durchlauf ", durchlauf, " und Tag ", day, " gefunden.\n")
        return(NULL)
      }
      
      excel_filename <- paste0("Summary_", day, ".xlsx")
      excel_filepath <- file.path(target_folder, excel_filename)
      
      write_xlsx(
        list("Data" = alle_daten_tag, 
             "Statistics Bends" = berechne_statistiken(alle_daten_tag)), 
        path = excel_filepath
      )
      
      gesamte_daten <- read_excel(excel_filepath, sheet = "Data")
      
      # Gruppennamen mit Einheit versehen
      gesamte_daten <- gesamte_daten %>%
        mutate(
          Group = case_when(
            Group == "H2O" ~ "H2O",
            TRUE ~ paste0(Group, " µg/ml")
          )
        )
      
      gesamte_daten$Group <- factor(gesamte_daten$Group,
                                    levels = c(
                                      "H2O",
                                      "Aerosil200 1 µg/ml",
                                      "Aerosil200 10 µg/ml",
                                      "Aerosil200 100 µg/ml",
                                      "Aerosil200 200 µg/ml",
                                      "Aerosil200 500 µg/ml"
                                    )
      )
      
      # partikel <- list("Evonik Industries" = "Aerosil200") # deaktiviert
      
      # Plot-Daten definieren
      datensatz <- filter(gesamte_daten, Group %in% levels(gesamte_daten$Group))
      partikel_name <- "Aerosil200"
      
      plots <- list(
        boxplot = ggplot(datensatz, aes(x = Group, y = `Bends per 30s`, fill = Group)) +
          geom_boxplot(color = "black") +
          scale_fill_manual(values = c(
            "H2O" = "blue",
            "Aerosil200 1 µg/ml" = "grey90",
            "Aerosil200 10 µg/ml" = "grey70",
            "Aerosil200 100 µg/ml" = "grey50",
            "Aerosil200 200 µg/ml" = "grey30",
            "Aerosil200 500 µg/ml" = "grey10"
          )) +
          theme(
            panel.background = element_rect(fill = "white"),
            panel.grid.major = element_blank(),
            panel.grid.minor = element_blank(),
            axis.line = element_line(colour = "black"),
            legend.position = "none",
            axis.title.x = element_blank(),
            axis.text.x = element_text(angle = 45, hjust = 1, color = "black"),
            axis.text.y = element_text(color = "black"),
            plot.title = element_text(hjust = 0.5)
          ) +
          ggtitle(partikel_name) +
          ylim(0, 250),
        
        violinplot1 = ggplot(datensatz, aes(x = Group, y = `Bends per 30s`, fill = Group, color = Group)) +
          geom_violin(trim = FALSE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) +
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          scale_color_manual(values = c(
            "H2O" = "blue",
            "H2O" = "blue",
            "Aerosil200 1 µg/ml" = "grey90",
            "Aerosil200 10 µg/ml" = "grey70",
            "Aerosil200 100 µg/ml" = "grey50",
            "Aerosil200 200 µg/ml" = "grey30",
            "Aerosil200 500 µg/ml" = "grey10"
          )) +
          theme(
            panel.background = element_rect(fill = "white"),
            panel.grid.major = element_blank(),
            panel.grid.minor = element_blank(),
            axis.line = element_line(colour = "black"),
            legend.position = "none",
            axis.title.x = element_blank(),
            axis.text.x = element_text(angle = 45, hjust = 1, color = "black"),
            axis.text.y = element_text(color = "black"),
            plot.title = element_text(hjust = 0.5)
          ) +
          ggtitle(partikel_name) +
          ylim(0, 250),
        
        violinplot2 = ggplot(datensatz, aes(x = Group, y = `Bends per 30s`, fill = Group, color = Group)) +
          geom_violin(trim = FALSE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) +
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          geom_point(shape = 16, size = 0.4, color = "grey20", alpha = 0.85, position = position_jitter(width = 0.1, height = 0)) +
          scale_color_manual(values = c(
            "H2O" = "blue",
            "Aerosil200 1 µg/ml" = "grey90",
            "Aerosil200 10 µg/ml" = "grey70",
            "Aerosil200 100 µg/ml" = "grey50",
            "Aerosil200 200 µg/ml" = "grey30",
            "Aerosil200 500 µg/ml" = "grey10"
          )) +
          theme(
            panel.background = element_rect(fill = "white"),
            panel.grid.major = element_blank(),
            panel.grid.minor = element_blank(),
            axis.line = element_line(colour = "black"),
            legend.position = "none",
            axis.title.x = element_blank(),
            axis.text.x = element_text(angle = 45, hjust = 1, color = "black"),
            axis.text.y = element_text(color = "black"),
            plot.title = element_text(hjust = 0.5)
          ) +
          ggtitle(partikel_name) +
          ylim(0, 250)
      )
      
      lapply(names(plots), function(p) {
        ggsave(file.path(target_folder, paste0(partikel_name, "_", p, "_Bends per 30s_", day, ".tiff")),
               plot = plots[[p]], width = 10, height = 8, units = "cm", dpi = 600)
      })
    })
  })
})

cat("Fertig!\n")
