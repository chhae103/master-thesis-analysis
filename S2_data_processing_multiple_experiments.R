##An example of the required file structure is provided in the same folder as this R script
#Expected folder structure:
#Source folder should contain experiment folders named "Experiment 1" to "Experiment x".
#Each experiment folder should have subfolders for examination days (e.g. "D4"), 
#and these should contain folders for each treatment group (e.g. "BZ555_Control").
#The script will create an Excel file called 'Summarized_Data_", day, ".xlsx', summarizing the details from each 'particles' file (previously created by SingleWormTracker software)
#Additionally, it will create box plots, violin plots, and violin plots with individual data points for each particle type and each parameter, using the mean values of the summarized dataset
#All files will be saved in different specified target directories (see reference folder for an output example)

suppressWarnings({ #Suppresses unnecessary warning messages in the console
  # Install and load packages
  if (!require("pacman")) install.packages("pacman")
  pacman::p_load("readxl", "writexl", "dplyr", "ggplot2", "stringr", "fs") #Installs and loads required packages
  
  #Basisordner für alle Experimente
  data_path <- "data/" 
  
  # Zielordner definieren
  target_folder <- file.path(data_path, "Ergebnisse")
  if (!dir_exists(target_folder)) dir_create(target_folder)
  
   #Define the source directory (must contain data of all 20°C experiments, see folder as reference)
  experiment_day_folder <- file.path(main_folder, "D2")  #Beispiel für D2
  
  #Choose which experiments you want to include and name them here (experiment names in this script must be identical to your file names, see reference folder)
  experiments <- c("Experiment 1", "Experiment 2", "Experiment 3")
  
  # Specify which examination day to analyze (e.g., "D2", "D4", "D6")
  day <- "D2"
  
  
  #Function to load and label the data for each experiment
  load_data_and_label <- function(experiment) {
    #Constructs the file path to the 'Summarized_Data_day.xlsx' file for a specific experiment and day. These files should have been previously created by the R script 'SWT_Titration_example_with_several_parameters_for_all_experiments' and must exist for each experiment you want to include in the 'D4' subfolder
    file_path <- file.path(main_folder, experiment, "Ergebnisse", paste0("Summary_", day, ".xlsx"))
    
    #Checks for the presence of the file at the specified location. If missing, displays a message and returns NULL, allowing the script to continue processing other files without crashing, while notifying that a file is missing
    if (!file.exists(file_path)) {
      cat("File does not exist: ", file_path, "\n")
      return(NULL)
    }
    
    #TryCatch ensures the script continues executing even if a file is missing or unreadable. It logs a warning message for any missing files, allowing the script to skip over incomplete data (e.g., missing files for specific examination days) without interruption.
    data <- tryCatch({
      read_excel(file_path, sheet = "Data")
    }, error = function(e) {
      cat("Error reading the file: ", file_path, "\n")
      return(NULL)
    })
    
    if (is.null(data) || nrow(data) == 0) {
      cat("No data found in the file: ", file_path, "\n")
      return(NULL)
    }
    
    data$Experiment <- experiment #Adds an 'Experiment' column to the data in the Excel file for later identifications
    return(data)
  }
  
  #Reads and combines all data into one table
  all_data <- bind_rows(lapply(experiments, load_data_and_label))
  
  if (nrow(all_data) == 0) { #If no data is successfully loaded after processing all experiments, a message is displayed, and the script halts further execution to prevent subsequent errors due to missing data
    cat("No data available to save.\n")
  } else {
    
    
    # Gruppennamen anpassen: Einheit "µg/ml" anhängen, wenn noch nicht vorhanden
    all_data <- all_data %>%
      mutate(
        Group = case_when(
          Group == "H2O" ~ "H2O",
          str_detect(Group, "µg/ml") ~ Group,
          TRUE ~ paste0(Group, " µg/ml")
        )
      )
    
    # Ergebnisordner definieren (z. B. Results_D4)
    output_folder <- file.path(main_folder, paste0("Results_", day))
    
    # Falls der Ordner noch nicht existiert, anlegen
    if (!dir.exists(output_folder)) {
      dir.create(output_folder, recursive = TRUE)
      cat("Output folder created at:", output_folder, "\n")
    }
    
    # Kombinierte Excel-Datei nun dort speichern
    target_file_path <- file.path(output_folder, paste0("Triplicate_Experiments_", day, ".xlsx"))
    
    
    
    #Write data into a new Excel file
    write_xlsx(list("Data" = all_data), path = target_file_path)
    
    cat("Data successfully combined and saved in: ", target_file_path, "\n")
    
    #Path to the Excel file for plot creation
    excel_filepath <- target_file_path
    
    #Reads the newly saved Excel file 'Combined_Data_All_Experiments.xlsx'
    complete_data <- read_excel(excel_filepath, sheet = "Data")
    
    #Ensures that the required columns are present in the Excel file, include all parameters you would like to look at here
    required_columns <- c("Bends per 30s", "Speed", "Max speed", "Dist per bend")
    missing_columns <- setdiff(required_columns, colnames(complete_data))
    
    if (length(missing_columns) > 0) { #Stops execution and lists missing columns if required data is incomplete
      stop("The following columns are missing in the Excel file: ", paste(missing_columns, collapse = ", "))
    }
    
    #Sets the 'Group' column as a factor (categorical variable) and defines the order of its levels.
    desired_order <- c("H2O", "Aerosil200 1 µg/ml", "Aerosil200 10 µg/ml", "Aerosil200 100 µg/ml", "Aerosil200 200 µg/ml", "Aerosil200 500 µg/ml")
    complete_data$Group <- factor(complete_data$Group, levels = desired_order)
    
    # Define particles: assigns all levels to one particle type ("Aerosil200")
    particles <- list("Aerosil200" = levels(complete_data$Group))
    
    #Organizes treatment groups by particle type (e.g., Kisker, Ae90, Ae200) e.g., matches treatment group names containing 'Kisker', 'Ae90' or 'Ae200'. This allows the code to process the data of each particle type separately
    #particles <- list(
    # "Kisker Day 2" = grep("Kisker", levels(complete_data$Group), value = TRUE)
    #)
    
    #Set a consistent color palette for treatment groups to ensure uniformity across all plots
    get_palette <- function() {
      c("H2O" = "blue", 
        "Aerosil200 1 µg/ml" = "grey70",
        "Aerosil200 10 µg/ml" = "grey60",
        "Aerosil200 100 µg/ml" = "grey40",
        "Aerosil200 200 µg/ml" = "grey20",
        "Aerosil200 500 µg/ml" = "grey10"
      )
    }
    
    #Set a consistent theme for all plots to ensure uniformity
    theme_custom <- function() {
      theme(
        panel.background = element_rect(fill = "white"), #Sets plot background to a specific color
        panel.grid.major = element_blank(), #Removes major grid lines
        panel.grid.minor = element_blank(), #Removes minor grid lines
        axis.line = element_line(colour = "black"), #Draws an axis line to separate axes in a specific color
        legend.position = "none", #Hides legend
        axis.title.x = element_blank(), #Removes x-axis title
        axis.text.x = element_text(angle = 45, hjust = 1, color = "black"), #Rotates labels in a specific angle, aligns text with the axis and sets the text to a specific color
        axis.text.y = element_text(color = "black"), #Sets the color of the y-axis to a specific color
        plot.title = element_text(hjust = 0.5),  #Center-aligns the plot title
        clip = "off"
      )
    }
    
    #Function to create and save different types of plots (box plot, violin plot, violin plot with data points, scatter plot)
    create_plots <- function(dataset, particle_name, output_folder) {
      plots <- list(
        # Plots for Bends per 30s
        boxplot_bends_per_30s = ggplot(dataset, aes(x = Group, y = `Bends per 30s`, fill = Group)) + #Box plot = Displays distribution of the selected metric/parameter for each group
          geom_boxplot(color = "black") + #Outlines the boxes in a specific color
          stat_summary(fun = mean, geom = "point", size = 2, shape = 4, color = "black", stroke = 1) + #Adds a mean point in the box plot, with a specific shape and color
          scale_fill_manual(values = get_palette()) + #Customizes fill colors for each treatment group using the previously set color palette
          theme_custom() + #Customizes the appearance of the plot using the previously set custom theme
          ggtitle(particle_name) + #Adds the title of the plot using e.g. particle names
          coord_cartesian(ylim = c(0, 250)), #Sets the y-axis limits to specific values
        
        violinplot1_bends_per_30s = ggplot(dataset, aes(x = Group, y = `Bends per 30s`, fill = Group, color = Group)) + #Violinplot1 = Displays distribution with violins and a box plot inside
          geom_violin(trim = FALSE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) + 
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          scale_color_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 250)),
        
        violinplot2_bends_per_30s = ggplot(dataset, aes(x = Group, y = `Bends per 30s`, fill = Group, color = Group)) + #Violinplot2 = Adds individual data points over the violin plot
          geom_violin(trim = FALSE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) +
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          geom_point(shape = 16, size = 0.4, color = "grey20", alpha = 0.85, position = position_jitter(width = 0.1, height = 0)) +
          scale_color_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 250), clip = "on"),
        
        # Plots for Dist per Bend
        boxplot_dist_per_bend = ggplot(dataset, aes(x = Group, y = `Dist per bend`, fill = Group)) +
          geom_boxplot(color = "black") +
          stat_summary(fun = mean, geom = "point", size = 2, shape = 4, color = "black", stroke = 1) +
          scale_fill_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 0.3)),
        
        violinplot1_bends_per_30s = ggplot(dataset, aes(x = Group, y = `Dist per bend`, fill = Group, color = Group)) +
          geom_violin(trim = FALSE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) +
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          scale_color_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 250)),
        
        violinplot2_dist_per_bend = ggplot(dataset, aes(x = Group, y = `Dist per bend`, fill = Group, color = Group)) +
          geom_violin(trim = TRUE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) +
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          geom_point(shape = 16, size = 0.4, color = "grey20", alpha = 0.85, position = position_jitter(width = 0.1, height = 0)) +
          scale_color_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 0.3)),
        
        # Plots for Speed
        boxplot_speed = ggplot(dataset, aes(x = Group, y = `Speed`, fill = Group)) +
          geom_boxplot(color = "black") +
          stat_summary(fun = mean, geom = "point", size = 2, shape = 4, color = "black", stroke = 1) +
          scale_fill_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 0.5)),
        
        violinplot1_bends_per_30s = ggplot(dataset, aes(x = Group, y = `Speed`, fill = Group, color = Group)) +
          geom_violin(trim = FALSE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) +
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          scale_color_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 250)),
        
        violinplot2_speed = ggplot(dataset, aes(x = Group, y = `Speed`, fill = Group, color = Group)) +
          geom_violin(trim = TRUE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) +
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          geom_point(shape = 16, size = 0.4, color = "grey20", alpha = 0.85, position = position_jitter(width = 0.1, height = 0)) +
          scale_color_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(clip = "on"),
        
        # Plots for Max Speed
        boxplot_max_speed = ggplot(dataset, aes(x = Group, y = `Max speed`, fill = Group)) +
          geom_boxplot(color = "black") +
          stat_summary(fun = mean, geom = "point", size = 2, shape = 4, color = "black", stroke = 1) +
          scale_fill_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 0.5)),
        
        violinplot1_bends_per_30s = ggplot(dataset, aes(x = Group, y = `Max speed`, fill = Group, color = Group)) +
          geom_violin(trim = FALSE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) +
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          scale_color_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 250)),
        
        violinplot2_max_speed = ggplot(dataset, aes(x = Group, y = `Max speed`, fill = Group, color = Group)) +
          geom_violin(trim = TRUE, fill = "white") +
          geom_boxplot(width = 0.05, fill = "black", color = "black", outlier.shape = NA, fatten = 1, size = 0.25) +
          stat_summary(fun = median, geom = "point", aes(color = "white"), size = 0.9, shape = 21, fill = "white", stroke = 0.25) +
          geom_point(shape = 16, size = 0.4, color = "grey20", alpha = 0.85, position = position_jitter(width = 0.1, height = 0)) +
          scale_color_manual(values = get_palette()) +
          theme_custom() +
          ggtitle(particle_name) +
          coord_cartesian(ylim = c(0, 0.5)),
        
        scatterplot_bends_vs_speed = ggplot(dataset, aes(x = `Bends per 30s`, y = Speed, color = Group)) +
          geom_point(alpha = 0.8, shape = 16, size = 1.25) +
          labs(x = "Bends per 30s", y = "Speed [mm/s]") +
          scale_color_manual(values = get_palette()) +
          theme_minimal() +
          theme(
            panel.grid = element_blank(),
            legend.position = "none",
            axis.title.x = element_text(color = "black", size = 10),
            axis.title.y = element_text(color = "black", size = 10),
            axis.text.x = element_text(color = "black", size = 10),
            axis.text.y = element_text(color = "black", size = 10),
            plot.title = element_text(hjust = 0.5),
            axis.line = element_line(color = "black")
          ) +
          ggtitle(paste(particle_name)) #Shows the particle names corresponding to the displayed data
      )
      
      #Save plots
      lapply(names(plots), function(p) { #Loops through each plot type and saves each plot as a .tiff file
        ggsave(file.path(output_folder, paste0(particle_name, "_", p, ".tiff")), plot = plots[[p]], width = 10, height = 8, units = "cm", dpi = 600) #Saves each plot to the target folder with a specific naming format and size
      })
    }
    
    #Creates plots for each particle type
    output_folder <- file.path(main_folder, paste0("Results_", day))
    
    
    # Create output folder if it doesn't exist
    if (!dir.exists(output_folder)) {
      dir.create(output_folder, recursive = TRUE)
      cat("Output folder created at:", output_folder, "\n")
    }
    #Iterates over each particle type defined in the 'particles' object
    for (particle_name in names(particles)) {
      dataset <- filter(complete_data, Group %in% particles[[particle_name]])
      create_plots(dataset, particle_name, output_folder)
    }
  }
})