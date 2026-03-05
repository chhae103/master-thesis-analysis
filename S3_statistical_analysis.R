##An example of the required file structure is provided in the same folder as this R script
#Expected folder structure:
#Source folder should contain experiment folders named "Experiment 1" to "Experiment x".
#Each experiment folder should have subfolders for examination days (e.g. "D4"), 
#and these should contain folders for each treatment group (e.g. "BZ555_Control").
#The script will create...................

suppressWarnings({ #Suppresses unnecessary warning messages in the console
  
  if (!requireNamespace("pacman", quietly = TRUE)) install.packages("pacman") #Installs and loads required packages
  pacman::p_load(readxl, openxlsx, dplyr, car, onewaytests, rstatix, stringr)
  
  #Basisordner für alle Experimente
  data_path <- "data/" 
  
  # Parameter für Experiment und Tag
  experiment <- "D10"          # z.B. D10, D11, etc.
  day <- 4                      # Tag oder Run
  run <- "Triplicate_Experiments" # optional für Präfix/Dateinamen
  
  #Define input and output files
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
  
  data <- read_excel(input_file) #Loads the data from the input file
  
  parameters <- c("Bends per 30s", "Speed", "Max speed", "Dist per bend") #Parameters to analyze. Add more here if additional data was included in the input file by previous scripts
  
  #Creates a new Excel workbook and defines styles for headers and titles
  wb <- createWorkbook() #Creates a new workbook
  header_style <- createStyle(textDecoration = "bold", fontSize = 12) #Style for headers
  title_style <- createStyle(textDecoration = "bold", fontSize = 14) #Style for titles
  
  #Define the custom order of treatment groups for the Excel list for consistent comparison in analysis
  custom_order <- c("H2O", "Aerosil200 1 µg/ml", "Aerosil200 10 µg/ml", "Aerosil200 100 µg/ml", "Aerosil200 200 µg/ml", "Aerosil200 500 µg/ml")
  
  #Function to perform normality (Shapiro-Wilk) and homogeneity (Levene's) tests
  get_norm_homogeneity_results <- function(data, param) {
    #Perform the Shapiro-Wilk test for normality within each treatment group
    normality <- data %>%
      group_by(Group) %>%
      summarise(p_value = shapiro.test(!!sym(param))$p.value, .groups = "drop") %>%
      mutate(normal = p_value > 0.05) #Define 'normal' as TRUE if p-value > 0.05
    
    #Perform Levene's test for homogeneity of variances
    homogeneity <- leveneTest(as.formula(paste0("`", param, "` ~ Group")), data = data)$"Pr(>F)"[1] > 0.05
    list(normality = normality, homogeneity_check = homogeneity)
  }
  
  #Function to format p-values into significance categories
  format_p_values <- function(p_values) {
    ifelse(p_values <= 0.0001, "****", 
           ifelse(p_values <= 0.001, "***",
                  ifelse(p_values <= 0.01, "**",
                         ifelse(p_values <= 0.05, "*", "n.s."))))
  }
  
  #Main loop through parameters
  for (param in parameters) {
    param_data <- data %>%
      select(all_of(c(param, "Group", "Experiment"))) %>%
      filter(!is.na(!!sym(param)))
    
    param_data$Group <- factor(param_data$Group, levels = custom_order)
    
    test_results <- get_norm_homogeneity_results(param_data, param)
    
    norm_homogeneity_table <- data.frame(
      Group = test_results$normality$Group,
      Shapiro_p_value = test_results$normality$p_value,
      Normal_Distribution = test_results$normality$normal,
      Levene_p_value = if (is.logical(test_results$homogeneity_check)) {
        test_results$homogeneity_check
      } else { NA }
    )
    
    result_table <- data.frame(
      Groups = character(),
      Mean_Rank_Diff = numeric(),
      Z = numeric(),
      Test = character(),
      P_Value = numeric(),
      Significance = integer(),
      P_Value_Asterisks = character(),
      stringsAsFactors = FALSE
    )
    
    rank_summary <- data.frame(
      Group = character(),
      n = numeric(),
      Mean_Rank = numeric(),
      Sum_Rank = numeric(),
      stringsAsFactors = FALSE
    )
    
    test_statistics <- data.frame(
      Chi_Square = numeric(),
      DF = numeric(),
      P_Value_Greater_Chi_Square = numeric(),
      stringsAsFactors = FALSE
    )
    
    #Choose test based on normality & homogeneity
    if (all(test_results$normality$normal) & test_results$homogeneity_check) {
      anova_result <- aov(as.formula(paste0("`", param, "` ~ Group")), data = param_data)
      tukey_result <- TukeyHSD(anova_result)
      
      anova_row <- data.frame(
        Groups = "ANOVA",
        Mean_Rank_Diff = NA,
        Z = NA,
        Test = "ANOVA",
        P_Value = summary(anova_result)[[1]][, "Pr(>F)"][1],
        Significance = as.integer(summary(anova_result)[[1]][, "Pr(>F)"][1] < 0.05),
        P_Value_Asterisks = ""
      )
      
      tukey_rows <- data.frame(
        Groups = rownames(tukey_result$Group),
        Mean_Rank_Diff = tukey_result$Group[, "diff"],
        Z = tukey_result$Group[, "diff"] / sqrt(var(tukey_result$Group[, "diff"])),
        Test = "Tukey",
        P_Value = tukey_result$Group[, "p adj"],
        Significance = as.integer(tukey_result$Group[, "p adj"] < 0.05),
        P_Value_Asterisks = ""
      )
      
      result_table <- rbind(result_table, anova_row, tukey_rows)
      rank_summary <- param_data %>%
        group_by(Group) %>%
        summarise(n = n(), Mean_Rank = mean(rank(!!sym(param))), Sum_Rank = sum(rank(!!sym(param))), .groups = "drop")
      test_statistics <- data.frame(Chi_Square = NA, DF = df.residual(anova_result), P_Value_Greater_Chi_Square = summary(anova_result)[[1]][, "Pr(>F)"][1])
      
    } else {
      kruskal_result <- kruskal_test(param_data, formula = as.formula(paste0("`", param, "` ~ Group")))
      dunn_test <- dunn_test(param_data, formula = as.formula(paste0("`", param, "` ~ Group")), p.adjust.method = "bonferroni")
      
      kruskal_row <- data.frame(
        Groups = "Kruskal-Wallis",
        Mean_Rank_Diff = NA,
        Z = NA,
        Test = "Kruskal-Wallis",
        P_Value = NA,
        Significance = NA,
        P_Value_Asterisks = ""
      )
      
      dunn_rows <- data.frame(
        Groups = paste(dunn_test$group1, "vs", dunn_test$group2),
        Mean_Rank_Diff = dunn_test$statistic,
        Z = dunn_test$statistic,
        Test = "Dunn",
        P_Value = dunn_test$p.adj,
        Significance = as.integer(dunn_test$p.adj < 0.05),
        P_Value_Asterisks = ""
      )
      
      result_table <- rbind(result_table, kruskal_row, dunn_rows)
      rank_summary <- param_data %>%
        group_by(Group) %>%
        summarise(n = n(), Mean_Rank = mean(rank(!!sym(param))), Sum_Rank = sum(rank(!!sym(param))), .groups = "drop")
      test_statistics <- data.frame(Chi_Square = kruskal_result$statistic, DF = length(unique(param_data$Group)) - 1, P_Value_Greater_Chi_Square = kruskal_result$p)
    }
    
    result_table$P_Value_Asterisks <- format_p_values(result_table$P_Value)
    
    addWorksheet(wb, param)
    writeData(wb, param, "Shapiro-Wilk and Levene's Test", startRow = 1, startCol = 1, headerStyle = title_style)
    writeData(wb, param, norm_homogeneity_table, startRow = 2, startCol = 1, headerStyle = header_style)
    
    start_col_result <- ncol(norm_homogeneity_table) + 2
    
    writeData(wb, param,
              ifelse(all(test_results$normality$normal) & test_results$homogeneity_check, 
                     "1wayANOVA (Tukey's post hoc test)", 
                     "Kruskal-Wallis ANOVA (Dunn's post hoc test)"), 
              startRow = 1, startCol = start_col_result, headerStyle = title_style)
    writeData(wb, param, result_table, startRow = 2, startCol = start_col_result, headerStyle = header_style)
    
    start_col_rank <- start_col_result + ncol(result_table) + 1
    writeData(wb, param, "Ranks", startRow = 1, startCol = start_col_rank, headerStyle = title_style)
    writeData(wb, param, rank_summary, startRow = 2, startCol = start_col_rank, headerStyle = header_style)
    
    start_col_test_stats <- start_col_rank + ncol(rank_summary) + 1
    writeData(wb, param, "Test Statistics", startRow = 1, startCol = start_col_test_stats, headerStyle = title_style)
    writeData(wb, param, test_statistics, startRow = 2, startCol = start_col_test_stats, headerStyle = header_style)
    
    # --- Mittelwerte pro Gruppe berechnen ---
    group_means <- param_data %>%
      group_by(Group) %>%
      summarise(
        n = n(),
        Mean = mean(!!sym(param), na.rm = TRUE),
        SD = sd(!!sym(param), na.rm = TRUE),
        SE = SD / sqrt(n),
        .groups = "drop"
      )
    
    # --- Mittelwert-Tabelle in Excel hinzufügen ---
    start_col_means <- start_col_test_stats + ncol(test_statistics) + 1  # +1 für Abstand
    writeData(wb, param, "Group Means", startRow = 1, startCol = start_col_means, headerStyle = title_style)
    writeData(wb, param, group_means, startRow = 2, startCol = start_col_means, headerStyle = header_style)
    setColWidths(wb, param, cols = start_col_means:(start_col_means + ncol(group_means) - 1), widths = "auto")
    
    #Adjust column widths
    setColWidths(wb, param, cols = 1:ncol(norm_homogeneity_table), widths = "auto")
    setColWidths(wb, param, cols = start_col_result:(start_col_result + ncol(result_table) - 1), widths = "auto")
    setColWidths(wb, param, cols = start_col_rank:(start_col_rank + ncol(rank_summary) - 1), widths = "auto")
    setColWidths(wb, param, cols = start_col_test_stats:(start_col_test_stats + ncol(test_statistics) - 1), widths = "auto")
  }
  
  saveWorkbook(wb, output_file, overwrite = TRUE)
})
