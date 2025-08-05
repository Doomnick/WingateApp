#verze 5.8.25
library(shiny)
library(shinyFiles)
library(readxl)
library(fs)
library(shinyjs)
library(future.apply)
library(dplyr)
plan(multisession)



shinyServer(function(input, output, session) {
  
  result_val <- reactiveVal()
  spatne.wingate <- reactiveVal(c())
  spatne.compar <- reactiveVal(c())
  spatne.compar_2 <- reactiveVal(c())
  spatne.spiro <- reactiveVal(c())
  duplikaty_check <- reactiveVal(c())
  chybne.pocty <- reactiveVal(c())
  spiro_ids <- reactiveVal(NULL)
  i2 <- reactiveVal(NULL)
  df4_a <- reactiveVal(NULL)
  date_check <- reactiveVal(NULL)
  SJ_check <- reactiveVal(FALSE)
  wd1 <- reactiveVal(c())
  makeReactiveBinding("refreshFiles")
  refreshFiles <<- 0
  file.list <- reactiveVal(c())
  
  # Volumes and Main Folder
  volumes <- shinyFiles::getVolumes()()
  program_slozka <- normalizePath(getwd())
  volumes <- c("Program Directory" = dirname(program_slozka), volumes)
  names(volumes) <- gsub(".*\\((.*)\\).*", "\\1", names(volumes))
  
  shinyDirChoose(input, "main_folder", roots = volumes, session = session, allowDirCreate = TRUE)
  shinyDirChoose(input, "antropometrie", roots = volumes, session = session, allowDirCreate = TRUE)
  shinyDirChoose(input, "wingate_path", roots = volumes, session = session, allowDirCreate = TRUE)
  shinyDirChoose(input, "srovnani_path", roots = volumes, session = session, allowDirCreate = TRUE)
  shinyDirChoose(input, "srovnani2_path", roots = volumes, session = session, allowDirCreate = TRUE)
  shinyDirChoose(input, "spirometrie_path", roots = volumes, session = session, allowDirCreate = TRUE)
  
  autoPaths <- reactiveValues(
    antropometrie = NULL,
    wingate = NULL,
    spiro = NULL,
    srovnani = NULL,
    srovnani2 = NULL
  )
  
  
  observeEvent(input$main_folder, {
    selected_path <- parseDirPath(volumes, input$main_folder)
    
    if (length(selected_path) > 0 && dir.exists(selected_path)) {
      
      somato_path <- file.path(selected_path, "somato")
      wingate_path <- file.path(selected_path, "wingate")
      spiro_path <- file.path(selected_path, "spiro")
      srovnani_path <- file.path(wingate_path, "srovnani")
      srovnani2_path <- file.path(wingate_path, "srovnani2")
      
      autoPaths$antropometrie <- if (dir.exists(somato_path)) somato_path else NULL
      autoPaths$wingate <- if (dir.exists(wingate_path)) wingate_path else NULL
      autoPaths$spiro <- if (dir.exists(spiro_path)) spiro_path else NULL
      autoPaths$srovnani <- if (dir.exists(srovnani_path)) srovnani_path else NULL
      autoPaths$srovnani2 <- if (dir.exists(srovnani2_path)) srovnani2_path else NULL
      
      wd1(selected_path)
      
      updateCheckboxInput(session, "wingate", value = !is.null(autoPaths$wingate))
      updateCheckboxInput(session, "spirometrie", value = !is.null(autoPaths$spiro))
      updateCheckboxInput(session, "srovnani", value = !is.null(autoPaths$srovnani))
      updateCheckboxInput(session, "srovnani2", value = !is.null(autoPaths$srovnani2))
    }
  })


  # Detect subfolders after main folder selection
  observeEvent(input$antropometrie, {
    path <- parseDirPath(volumes, input$antropometrie)
    if (length(path) > 0 && dir.exists(path)) autoPaths$antropometrie <- path
  })
  
  observeEvent(input$wingate_path, {
    path <- parseDirPath(volumes, input$wingate_path)
    if (length(path) > 0 && dir.exists(path)) autoPaths$wingate <- path
  })
  
  observeEvent(input$srovnani_path, {
    path <- parseDirPath(volumes, input$srovnani_path)
    if (length(path) > 0 && dir.exists(path)) autoPaths$srovnani <- path
  })
  
  observeEvent(input$srovnani2_path, {
    path <- parseDirPath(volumes, input$srovnani2_path)
    if (length(path) > 0 && dir.exists(path)) autoPaths$srovnani2 <- path
  })
  
  observeEvent(input$spirometrie_path, {
    path <- parseDirPath(volumes, input$spirometrie_path)
    if (length(path) > 0 && dir.exists(path)) autoPaths$spiro <- path
  })

  # Reactive paths
  antropometrie_path <- reactive({ autoPaths$antropometrie })
  wingate_path <- reactive({ autoPaths$wingate })
  spirometrie_path <- reactive({ autoPaths$spiro })
  srovnani_path <- reactive({ autoPaths$srovnani })
  srovnani2_path <- reactive({ autoPaths$srovnani2 })

  # Output main path and detected subfolders
  output$main_folder_path <- renderPrint({
    selected_path <- parseDirPath(volumes, input$main_folder)
    if (length(selected_path) > 0 && !is.na(selected_path) && nzchar(selected_path)) {
      cat(selected_path)
    } else {
      cat("Should contain subfolders: 'somato', 'wingate', 'spiro',
and optionally 'wingate/srovnani', 'wingate/srovnani2'.")
    }
  })
  output$auto_detect_info <- renderPrint({
    req(parseDirPath(volumes, input$main_folder))
    cat("Auto-detected subfolders:\n")
    cat("Antropometrie: ", ifelse(is.null(autoPaths$antropometrie), "Not found", autoPaths$antropometrie), "\n")
    cat("Wingate: ", ifelse(is.null(autoPaths$wingate), "Not found", autoPaths$wingate), "\n")
    cat("Spiro: ", ifelse(is.null(autoPaths$spiro), "Not found", autoPaths$spiro), "\n")
    cat("Comparison: ", ifelse(is.null(autoPaths$srovnani), "Not found", autoPaths$srovnani), "\n")
    cat("Comparison 2: ", ifelse(is.null(autoPaths$srovnani2), "Not found", autoPaths$srovnani2), "\n")
  })
    
  
  file.lists <- reactive({
    refreshFiles
    input$refresh_files
    
    parseManualPath <- function(inputId, defaultPath) {
      manualPath <- parseDirPath(volumes, input[[inputId]])
      if (length(manualPath) > 0 && !is.na(manualPath) && dir.exists(manualPath)) {
        return(manualPath)
      }
      return(defaultPath)
    }
    
    antropometrie_path_val <- parseManualPath("antropometrie", antropometrie_path())
    wingate_path_val <- if (!is.null(input$wingate) && input$wingate) parseManualPath("wingate_path", wingate_path()) else NULL
    spirometrie_path_val <- if (!is.null(input$spirometrie) && input$spirometrie) parseManualPath("spirometrie_path", spirometrie_path()) else NULL
    srovnani_path_val <- if (!is.null(input$srovnani) && input$srovnani) parseManualPath("srovnani_path", srovnani_path()) else NULL
    srovnani2_path_val <- if (!is.null(input$srovnani2) && input$srovnani2) parseManualPath("srovnani2_path", srovnani2_path()) else NULL
    
    output$antropometrie_path <- renderPrint({
      cat(ifelse(!is.null(antropometrie_path_val) && dir.exists(antropometrie_path_val), antropometrie_path_val, "Not selected"))
    })
    output$wingate_path_display <- renderPrint({
      cat(ifelse(!is.null(wingate_path_val) && dir.exists(wingate_path_val), wingate_path_val, "Not selected"))
    })
    output$srovnani_path_display <- renderPrint({
      cat(ifelse(!is.null(srovnani_path_val) && dir.exists(srovnani_path_val), srovnani_path_val, "Not selected"))
    })
    output$srovnani2_path_display <- renderPrint({
      cat(ifelse(!is.null(srovnani2_path_val) && dir.exists(srovnani2_path_val), srovnani2_path_val, "Not selected"))
    })
    output$spirometrie_path_display <- renderPrint({
      cat(ifelse(!is.null(spirometrie_path_val) && dir.exists(spirometrie_path_val), spirometrie_path_val, "Not selected"))
    })
    
    
    if (length(antropometrie_path_val) > 0) {
      wd <- dirname(as.character(antropometrie_path_val))
      vymazat_folder <- file.path(wd, "vymazat")
      wd1(wd)
      
      if (!dir.exists(vymazat_folder)) {
        dir.create(vymazat_folder)
      }
      files_in_vymazat <- list.files(vymazat_folder)
      
      if (length(files_in_vymazat) > 0) {
        output$clear_button_ui <- renderUI({
          actionButton("clear_vymazat", "Empty 'vymazat' Folder")
        })
      } else {
        output$clear_button_ui <- renderUI({ NULL })
      }
    }
    
    antropometrie_files <- if (!is.null(antropometrie_path_val)) {
      files <- list.files(path = antropometrie_path_val, pattern = '\\.(xls[xm])$', full.names = TRUE)
      if (length(files) > 0) {
        file <- files[1]
        data <- readxl::read_excel(file, sheet = "Data_Sheet")
        
        check_dates <- function(data) {
          values_birth <- as.Date(data$Birth, format = "%d/%m/%Y")
          values_date_measurement <- as.Date(data$Date_measurement, format = "%d/%m/%Y")
          
          birth_check <- any(is.na(values_birth))
          date_measurement_check <- any(is.na(values_date_measurement))
          
          if (birth_check && date_measurement_check) {
            return("Invalid format of Birth and Date measurement columns! Please correct to continue.")
          } else if (birth_check) {
            return("Invalid format Birth column! Please correct to continue.")
          } else if (date_measurement_check) {
            return("Invalid format of Date measurement column! Please correct to continue.")
          } else {
            return("Birth and Date measurement column are in right format.")
          }
        }
        
        wrong_format <- check_dates(data)
        date_check(wrong_format)
        
        if ("ID" %in% colnames(data)) {
          id_values <- na.omit(data$ID)
          id_values
        } else {
          "ID column not found in 'Data_Sheet'."
        }
      }
    } else NULL
    
    squatjump_files <- NULL
    squat_jumps_count <- if (!is.null(antropometrie_path_val)) {
      files <- list.files(path = antropometrie_path_val, pattern = '\\.(xls[xm])$', full.names = TRUE)
      if (length(files) > 0) {
        file <- files[1]
        data <- readxl::read_excel(file, sheet = "Data_Sheet")
        
        if ("SJ" %in% colnames(data)) {
          SJ_check(TRUE)
          squatjump_files <- data$ID[!is.na(data$SJ)]
          length(na.omit(data$SJ))
        } else {
          "SJs not present"
        }
      } else {
        "SJs not present"
      }
    } else NULL
    
    wingate_files <- if (!is.null(wingate_path_val)) {
      files <- list.files(path = wingate_path_val, pattern = '\\.(txt)$', full.names = TRUE)
      basename(tools::file_path_sans_ext(files))
    } else NULL
    
    srovnani_files <- if (!is.null(srovnani_path_val)) {
      files <- list.files(path = srovnani_path_val, pattern = '\\.(txt)$', full.names = TRUE)
      basename(tools::file_path_sans_ext(files))
    } else NULL
    
    srovnani2_files <- if (!is.null(srovnani2_path_val)) {
      files <- list.files(path = srovnani2_path_val, pattern = '\\.(txt)$', full.names = TRUE)
      basename(tools::file_path_sans_ext(files))
    } else NULL
    
    spirometrie_files <- if (!is.null(spirometrie_path_val)) {
      files <- list.files(path = spirometrie_path_val, pattern = '\\.(xls[xm])$', full.names = TRUE)
      basename(tools::file_path_sans_ext(files))
    } else NULL
    
    list(
      Antropometrie = antropometrie_files,
      SquatJumps = squat_jumps_count,
      Wingate = wingate_files,
      Spirometrie = spirometrie_files,
      Srovnani = srovnani_files,
      Srovnani2 = srovnani2_files,
      SquatJumpsIDs = squatjump_files
    )
  })
  
  
  observeEvent(input$clear_vymazat, {
    antropometrie_dir <- antropometrie_path()
    
    if (!is.null(antropometrie_dir)) {
      wd <- dirname(as.character(antropometrie_dir))
      vymazat_folder <- file.path(wd, "vymazat")
      
      if (dir.exists(vymazat_folder)) {
        unlink(list.files(vymazat_folder, full.names = TRUE), recursive = TRUE, force = TRUE)
      }
    }
    
    refreshFiles <<- refreshFiles + 1  # Triger pro aktualizaci file.lists()
    
    output$clear_button_ui <- renderUI({
      NULL
    })
  })
  
  
  observe({
    file_data <- file.lists()
    disable_buttons <- is.null(file_data$Antropometrie) || 
      (length(file_data$Wingate) < 1 && length(file_data$Spirometrie) < 1)
    # Check based on Antropometrie files
    if (disable_buttons) {
      shinyjs::disable("continue")
      shinyjs::disable("refresh_files")
      shinyjs::disable("check")
    } else {
      shinyjs::enable("refresh_files")
      shinyjs::enable("check")
      shinyjs::runjs('$("#continue").show()')
      shinyjs::enable("continue")
    }
    if (!is.null(date_check()) && grepl("Invalid", date_check()) || disable_buttons) {
      shinyjs::disable("submit")
      shinyjs::disable("continue")
    } else {
      shinyjs::enable("submit")
    }
    
  })
  
  printFileList <- function(file.list, input) {
    output_text <- ""
    
    if (!is.null(file.list$Antropometrie)) {
      output_text <- paste0(output_text, "Anthropometry Files: ", length(file.list$Antropometrie), "\n")
    }
    
    if (input$wingate && !is.null(file.list$Wingate)) {
      output_text <- paste0(output_text, "Wingate Files: ", length(file.list$Wingate), "\n")
    }
    
    if (input$spirometrie && !is.null(file.list$Spirometrie)) {
      output_text <- paste0(output_text, "Spirometrie Files: ", length(file.list$Spirometrie), "\n")
    }
    
    if (!is.null(file.list$SquatJumps)) {
      output_text <- paste0(output_text, "Squat Jumps: ", file.list$SquatJumps, "\n")
    }
    
    if (input$srovnani && !is.null(file.list$Srovnani)) {
      output_text <- paste0(output_text, "Comparison Files: ", length(file.list$Srovnani), "\n")
    }
    
    if (input$srovnani2 && !is.null(file.list$Srovnani2)) {
      output_text <- paste0(output_text, "Comparison 2 Files: ", length(file.list$Srovnani2), "\n")
    }
    
    return(output_text)
  }
  output$invalid_ids <- renderPrint({
    # Check if `invalid_dates` is not NULL and not empty
    if (!is.null(date_check()) && nzchar(date_check())) {
      cat(date_check())
    } else {
      cat("No anthropometry file selected.")
    }
  })
  
  
  output$duplikaty.an <- renderPrint({
    if ((input$wingate && !input$srovnani) || (input$spirometrie && !input$wingate)) {
      ids_duplikaty_an <- duplikaty_check()
      if (length(ids_duplikaty_an) > 0) {
        print(ids_duplikaty_an)
        cat("\nCannot continue if there are duplicates in the Anthropometry file!")
        shinyjs::disable("submit")
      } else {
        cat("No duplicates found in the Anthropometry file.")
        shinyjs::enable("submit")
      }
    } else {
      shinyjs::enable("submit")
    }
  })  
  
  observe({
    file_data <- file.lists()
    output$file_list_output <- renderPrint({
      cat(printFileList(file_data, input))
    })
  })
  
  
  # Refresh button event
  observeEvent(input$refresh_files, {
    updateUIAndCheck()
    refreshFiles <<- refreshFiles + 1
    file.list <- file.lists()
    printFileList(file.list, input)
    shinyjs::runjs('$("#continue").show(); $("#submit").hide();')
    updateTabsetPanel(session, "main_tabs", selected = "Data")
  })
  
  # Render paths and file lists in the UI using autoPaths directly
  
  output$antropometrie_path <- renderPrint({
    req(!is.null(autoPaths$antropometrie))
    cat(autoPaths$antropometrie)
  })
  
  output$wingate_path_display <- renderPrint({
    if (input$wingate) {
      req(!is.null(autoPaths$wingate))
      cat(autoPaths$wingate)
    }
  })
  
  output$srovnani_path_display <- renderPrint({
    if (input$srovnani) {
      req(!is.null(autoPaths$srovnani))
      cat(autoPaths$srovnani)
    }
  })
  
  output$srovnani2_path_display <- renderPrint({
    if (input$srovnani2) {
      req(!is.null(autoPaths$srovnani2))
      cat(autoPaths$srovnani2)
    }
  })
  
  output$spirometrie_path_display <- renderPrint({
    if (input$spirometrie) {
      req(!is.null(autoPaths$spiro))
      cat(autoPaths$spiro)
    }
  })
  
  updateUIAndCheck <- function() {
    file_list <- file.lists()
    antropometrie_files <- file_list$Antropometrie
    wingate_files <- file_list$Wingate
    spirometrie_files <- file_list$Spirometrie
    srovnani_files <- file_list$Srovnani
    srovnani2_files <- file_list$Srovnani2
    if (SJ_check()) {
      SJ_files <- file_list$SquatJumpsIDs
    }
    
    all_files <- sort(unique(antropometrie_files))
    
    if (input$wingate) {
      all_files <- unique(c(all_files, wingate_files))
    }
    if (input$spirometrie) {
      all_files <- unique(c(all_files, spirometrie_files))
    }
    if (input$srovnani) {
      all_files <- unique(c(all_files, srovnani_files))
    }
    if (input$srovnani2) {
      all_files <- unique(c(all_files, srovnani2_files))
    }
    if(SJ_check()) {
      all_files <- unique(c(all_files, SJ_files))
    }
    
    all_files <- sort(unique(all_files))
    
    comparison_df <- data.frame(
      ID = all_files,
      Anthropometry = ifelse(all_files %in% antropometrie_files, "✓", "✗"),
      Wingate = if (input$wingate) ifelse(all_files %in% wingate_files, "✓", "✗") else NA,
      Comparison = if (input$srovnani) ifelse(all_files %in% srovnani_files, "✓", "✗") else NA,
      `Comparison 2` = if (input$srovnani2) ifelse(all_files %in% srovnani2_files, "✓", "✗") else NA,
      Spirometry = if (input$spirometrie) ifelse(all_files %in% spirometrie_files, "✓", "✗") else NA,
      SJ = if (SJ_check()) ifelse(all_files %in% SJ_files, "✓", "✗") else NA,
      stringsAsFactors = FALSE
    )
    
    
    if (input$srovnani || input$srovnani2) {
      comparison_df$`Anthropo compar.` <- sapply(comparison_df$ID, function(id) {
        # Count appearances of ID in antropometrie_files based on the input flags
        actual_count <- if (input$srovnani) sum(id == antropometrie_files) else 0
        
        # Determine expected count
        expected_count <- if (input$srovnani && input$srovnani2) 3 else 2
        
        # Determine the display value based on the count match
        if (actual_count >= expected_count) {
          "✓"
        } else if (input$srovnani && input$srovnani2 && actual_count == 2) {
          "2"
        } else {
          "✗"
        }
      })
    }
    
    
    
    if ("Anthropo compar." %in% names(comparison_df)) {
      # Reorder the columns
      comparison_df <- comparison_df[c("ID", "Anthropometry", "Anthropo compar.", setdiff(names(comparison_df), c("ID", "Anthropometry", "Anthropo compar.")))]
    }
    
    
    
    
    if (nrow(comparison_df) > 0 && ncol(comparison_df) > 1) {
      # Remove columns with only NA values
      comparison_df <- comparison_df[, colSums(!is.na(comparison_df)) > 0]
      
      # Update the Result column based on the conditions
      comparison_df$Report <- apply(comparison_df, 1, function(row) {
        # Extract column values
        ant_col <- row["Anthropometry"]
        wing_col <- row["Wingate"]
        comp_col <- row["Comparison"]
        comp2_col <- row["Comparison.2"]
        anthro_comp_col <- row["Anthropo compar."]
        spiro_col <- row["Spirometry"] 
        
        # Check if columns are present
        has_antropometrie <- !is.na(ant_col)
        has_wingate <- !is.na(wing_col)
        has_comparison <- !is.na(comp_col)
        has_comparison2 <- !is.na(comp2_col)
        has_anthropo_compar <- !is.na(anthro_comp_col)
        has_spiro <- !is.na(spiro_col) && spiro_col == "✓" 
        
        # Define the result
        report <- "FAILED"  # Default result
        
        
        # Check conditions
        if (has_spiro && has_antropometrie && 
            (!has_wingate || wing_col == "✗") && 
            (!has_comparison || comp_col == "✗") && 
            (!has_comparison2 || comp2_col == "✗") && 
            (!has_anthropo_compar || anthro_comp_col == "✗")) {
          report <- "Spiro"
        } else if (has_antropometrie && has_wingate) {
          if (has_comparison2 && all(c(ant_col, wing_col, comp_col, comp2_col, anthro_comp_col) == "✓", na.rm = TRUE)) {
            report <- "Triple"
          } else if (has_comparison) {
            if (all(c(ant_col, wing_col, comp_col) == "✓", na.rm = TRUE) && anthro_comp_col == "2") {
              report <- "Double"
            } else if (all(c(ant_col, wing_col, comp_col, anthro_comp_col) == "✓", na.rm = TRUE)) {
              report <- "Double"
            } else if (ant_col == "✓" && wing_col == "✓" && any(c(comp_col, comp2_col, anthro_comp_col) == "✗", na.rm = TRUE)) {
              report <- "Single"
            }
          } else if (!has_comparison && !has_comparison2 && ant_col == "✓" && wing_col == "✓") {
            report <- "Single"
          }
        } else if (has_comparison2 && all(c(ant_col, wing_col, comp_col, comp2_col) == "✓", na.rm = TRUE)) {
          report <- "Triple"
        } else if (has_comparison && all(c(ant_col, wing_col, comp_col) == "✓", na.rm = TRUE)) {
          report <- "Double"
        }
        
        if (report != "FAILED" && report != "Spiro" && has_spiro) {
          report <- paste(report, "+ Spiro")
        }
        
        return(report)
      })
    } else {
      comparison_df$Report <- character(nrow(comparison_df))  # Or any default value handling
    }
    
    colnames(comparison_df) <- gsub("Comparison\\.2", "Comparison 2", colnames(comparison_df))
    
    spiro_ids <- comparison_df$ID[comparison_df$Report == "Spiro"]
    spiro_ids(spiro_ids)
    
    output$comparison_table <- renderTable({
      comparison_df
    })
    
    if (input$wingate) {
      wingate_spatne <- comparison_df$ID[comparison_df$Wingate == "✗" | comparison_df$Anthropometry == "✗"]
      wingate_spatne <- paste0(wingate_spatne, ".txt")
      spatne.wingate(wingate_spatne)
    }
    
    if ((input$wingate && !input$srovnani) || (input$spirometrie && !input$wingate)){
      duplikaty_an <- antropometrie_files[duplicated(antropometrie_files)]
      duplikaty_check(duplikaty_an)
    }
    
    
    if (input$spirometrie) {
      spiro_spatne <- comparison_df$ID[comparison_df$Spirometry == "✗"]
      spatne.spiro(spiro_spatne)
    }
    
    if (input$srovnani) {
      srovnani_spatne <- comparison_df$ID[comparison_df$Comparison == "✗"]
      file.list.an.bez <- antropometrie_files 
      file.list.bez <- wingate_files
      file.list.compar.bez <- srovnani_files
      pocty.duplikatu.an <- dplyr::as_tibble(table(file.list.an.bez))
      names(pocty.duplikatu.an )[1] <- "ID"
      names(pocty.duplikatu.an )[2] <- "antropa"
      if (input$srovnani2) {
        srovnani2_spatne <- comparison_df$ID[comparison_df$`Comparison 2` == "✗"]
        file.list.compar_2.bez <- srovnani2_files
        merged.file.list <- c(file.list.bez, file.list.compar.bez, file.list.compar_2.bez)
        spatne.compar_2(srovnani2_spatne)
      } else {
        merged.file.list <- c(file.list.bez, file.list.compar.bez)
      }
      pocty.file.list <- dplyr::as_tibble(table(merged.file.list))
      pocty.file.list <- pocty.file.list[pocty.file.list$n != 1, ]
      names(pocty.file.list)[1] <- "ID"
      names(pocty.file.list)[2] <- "wingate"
      pocty.file.list <- pocty.file.list %>%  dplyr::left_join(pocty.duplikatu.an, by = "ID")
      pocty.file.list$diff <- pocty.file.list$wingate - pocty.file.list$antropa
      chybne_pocty <- as.vector(pocty.file.list$ID[pocty.file.list$diff > 0])
      chybne.pocty(chybne_pocty)
      spatne.compar(srovnani_spatne)
    }
    
    updateTabsetPanel(session, "main_tabs", selected = "Data")
    
    path_antropometrie <- !is.null(antropometrie_path())
    path_wingate <- !is.null(wingate_path())
    path_spirometrie <- !is.null(spirometrie_path())
    
    if (input$wingate && !input$srovnani) {
      ids_duplikaty_an <- duplikaty_check()
      duplicates_condition <- length(ids_duplikaty_an) > 0
    }
  }
  

  
  observeEvent(input$check, {
    refreshFiles <<- refreshFiles + 1
    file.list <- file.lists()
    printFileList(file.list, input)
    updateUIAndCheck()
    shinyjs::runjs('$("#continue").show(); $("#submit").hide();')
  })
  
  output$spatne_wingate_ids <- renderPrint({
    ids <- spatne.wingate()
    ids <- sub("\\.txt$", "", ids)
    if (length(ids) > 0) {
      print(ids)
      cat("\nUpon pressing Submit, no Wingate reports will be generated for these IDs.")
    } else {
      cat("No issues found in Anthropometry : Wingate.")
    }
  })
  
  output$spatne.spiro <- renderPrint({
    ids_spiro <- spatne.spiro()
    if (length(ids_spiro) > 0) {
      print(ids_spiro)
      cat("\nUpon pressing Submit, no Spirometry reports will be generated for these IDs.")
    } else {
      cat("No issues found in Spirometry.")
    }
  })
  
  output$spatne_srovnani <- renderPrint({
    ids_srovnani <- spatne.compar()
    if (length(ids_srovnani) > 0) {
      print(ids_srovnani)
      cat("\nUpon pressing Submit, no Comparison 1 reports will be generated for these IDs.")
    } else {
      cat("No issues found in Comparison 1.")
    }
  })
  
  output$spatne_srovnani2 <- renderPrint({
    ids_srovnani2 <- spatne.compar_2()
    if (length(ids_srovnani2) > 0) {
      print(ids_srovnani2)
      cat("\nUpon pressing Submit, no Comparison 2 reports will be generated for these IDs.")
    } else {
      cat("No issues found in Comparison 2.")
    }
  })
  
  output$spatne_srovnani.an <- renderPrint({
    ids_pocty <- chybne.pocty()
    if (length(ids_pocty) > 0) {
      print(ids_pocty)
      cat("\nUpon pressing Submit, no Comparison reports will be generated for these IDs.")
    } else {
      cat("No issues found in Comparison.")
    }
  })
  
  # output$console_output <- renderText({
  #   console_log()
  # })
  # 
  observeEvent(input$continue, {
    refreshFiles <<- refreshFiles + 1
    file.list <- file.lists()
    printFileList(file.list, input)
    updateUIAndCheck()
    updateTabsetPanel(session, "main_tabs", selected = "Summary")
    shinyjs::runjs('$("#continue").hide(); $("#submit").show();')
  })
  
  
  
  observeEvent(input$submit, {
    
    result_val(NULL)
    
    # Reactive paths and inputs
    sport <- input$sport 
    antropometrie_path <- antropometrie_path()
    wingate_path <- wingate_path()
    spirometrie_path <- spirometrie_path()
    comparison.path <- srovnani_path()
    comparison2_path <- srovnani2_path()
    srovnani <- input$srovnani
    srovnani2 <- input$srovnani2
    dotaz_spiro <- input$spirometrie
    units_switch <- input$toggle_switch
    team <- input$team
    tri_graf <- input$tri_graf
    program_slozka <- normalizePath(getwd())
    
    # ReactiveVal values
    spatne_wingate <- trimws(spatne.wingate())
    spatne_compar <- trimws(spatne.compar())
    chybne.pocty <- trimws(chybne.pocty())
    spatne.compar_2 <- trimws(spatne.compar_2())
    spatne.spiro <- trimws(spatne.spiro())
    spiro_ids <- trimws(spiro_ids())
    wd <- wd1()
    
    if (length(chybne.pocty) > 0) {
      spatne_compar <- unique(c(spatne_compar, chybne.pocty))
      spatne_compar2 <- unique(c(spatne.compar_2, chybne.pocty))
    }
    
    setwd(wd)
    
    # Convert logical to YES/NO
    dotaz_spiro <- ifelse(dotaz_spiro, "YES", "NO")
    srovnani <- ifelse(srovnani, "YES", "NO")
    srovnani2 <- ifelse(srovnani2, "YES", "NO")
    tri_graf <- ifelse(tri_graf, "YES", "NO")
    
    # Define paths
    file.path <- wingate_path
    file.path.an <- antropometrie_path
    report.path <- file.path(wd, "reporty")
    database.path <- file.path(wd, "databaze")
    spiro.path <- spirometrie_path
    
    # File lists
    safe_list_files <- function(path, pattern, full.names = FALSE, ignore.case = FALSE) {
      if (!is.null(path) && dir.exists(path)) {
        return(list.files(path = path, pattern = pattern, full.names = full.names, ignore.case = ignore.case))
      } else {
        return(character(0))  # prázdný character vector, aby to nespadlo
      }
    }
    
    # Použití:
    file.list <- basename(safe_list_files(file.path, '\\.txt$', full.names = TRUE))
    file.list.an <- basename(safe_list_files(file.path.an, '\\.xls[xm]$', ignore.case = TRUE))
    file.list.dat <- basename(safe_list_files(database.path, '\\.xls$', ignore.case = TRUE))
    file.list.bez <- sort(tools::file_path_sans_ext(file.list))
    
    file.list.compar_2.bez <- NA
    file.list.compar.bez <- NA
    
    if (dotaz_spiro == "YES") {
      file.list.spiro <- basename(list.files(path = spiro.path, pattern = '\\.xls[xm]$', ignore.case = TRUE))
      file.list.spiro.bez <- sort(tools::file_path_sans_ext(file.list.spiro))
    }
    
    # Load antropometry data
    antropo <- readxl::read_excel(file.path(file.path.an, file.list.an[1]), sheet = "Data_Sheet") 
    an.input <- "batch"
    
    if (an.input == "batch") {
      file.list.an.bez <- antropo$ID
      lactate.bez <- antropo$ID[!is.na(antropo$LA)]
    } else {
      file.list.an.bez <- sort(tools::file_path_sans_ext(file.list.an))
    }
    
    # Exclude problematic Wingate files
    file.list <- file.list[!file.list %in% spatne_wingate]
    
    # Process comparison files
    if (srovnani == "YES") {
      file.list.compar.bez <- basename(list.files(path = comparison.path, pattern = '\\.txt$', full.names = TRUE))
      file.list.compar.bez <- sub("\\.txt$", "", file.list.compar.bez)
      
      if (srovnani2 == "YES") {
        file.list.compar2_bez <- basename(list.files(path = comparison2_path, pattern = '\\.txt$', full.names = TRUE))
        file.list.compar2_bez <- sub("\\.txt$", "", file.list.compar2_bez)
        
        if (length(spatne_compar2) > 0) {
          file.list.compar2_bez <- file.list.compar2_bez[!file.list.compar2_bez %in% spatne_compar2]
          if (length(file.list.compar2_bez) < 1) {
            file.list.compar2_bez <- NA
          }
        }
      } else {
        file.list.compar2_bez <- NA
      }
      
      if (length(spatne_compar) > 0) {
        file.list.compar.bez <- file.list.compar.bez[!file.list.compar.bez %in% spatne_compar]
      }
    }
    
    #-----------------------------------------------------------------------------#
    
    options(error=traceback)
    ##### Reference values for GUI of sport selection ####
    sports_data <- data.frame(
      Sport = c("hokej-dospělí","hokej-junioři", "hokej-dorost", "gymnastika"),
      FVC = c(4.7, 4.7, 4.7, 3.8),    
      FEV1 = c(4.7, 4.7, 4.7, 3.1),   
      VO2max = c(58, 58, 58, 48),   
      PowerVO2 = c(4.5, 4.4, 4.4 ,3.3),
      Pmax = c(15.7, 15.3, 14.6, 13),
      SJ = c(45, 42, 42, 35),
      ANC = c(365, 330, 320 ,300)
    )
    
    #### Automatic check and instalation of missing packages ####
    
    
    process_and_generate_pdfs <- function(wingate_path, antropometrie_path, spirometrie_path, N, incProgress, units_switch, program_slozka) {  
      
      instalace <- function() {
        files <- list.files(pattern='[.](R|rmd)$', all.files=T, recursive=T, full.names = T, ignore.case=T) # find all source code files in (sub)folders
        code=unlist(sapply(files, scan, what = 'character', quiet = TRUE))   # read in source code
        
        code <- code[grepl('^library', code, ignore.case=T)]   # retain only source code starting with library
        code <- gsub('^library[(]', '', code)
        code <- gsub('[)]', '', code)
        code <- gsub('^library$', '', code)
        
        uniq_packages <- unique(code)   # retain unique packages
        uniq_packages <- uniq_packages[!uniq_packages == '']   # kick out "empty" package names
        uniq_packages <- uniq_packages[order(uniq_packages)]   # order alphabetically
        
        cat('Required packages: \n')
        cat(paste0(uniq_packages, collapse= ', '),fill=T)
        cat('\n\n\n')
        
        installed_packages <- installed.packages()[, 'Package']   # retrieve list of already installed packages
        
        to_be_installed <- setdiff(uniq_packages, installed_packages)   # identify missing packages
        
        if (length(to_be_installed)==length(uniq_packages)) cat('Vsechny balicky musi byt nainstalovany - instaluji.\n')
        if (length(to_be_installed)>0) cat('Instaluji chybejici balicky.\n')
        if (length(to_be_installed)==0) cat('Vsechny balicky jsou jiz nainstalovany!\n')
        
        if (length(to_be_installed)>0) install.packages(to_be_installed, repos = 'https://cloud.r-project.org')   # install missing packages
      }
      
      instalace()
      
      if (!dir.exists(paste(wd, "/vysledky", sep=""))) {
        dir.create(paste(wd, "/vysledky", sep=""))
      }
      
      
      
      if (!dir.exists(paste(wd, "/reporty", sep=""))) {
        dir.create(paste(wd, "/reporty", sep=""))
      }
      
      
      
      library(finalfit)
      library(ggplot2)
      library(lubridate)
      library(gt)
      library(cowplot)
      library(tidyverse)
      library(tinytex)
      library(rapportools)
      library(scales)
      library(webshot)
      library(patchwork)
      print(program_slozka)
      
      #### MAIN FOR LOOP ####
      for(i in 1:length(file.list)) {
        id <- tools::file_path_sans_ext(file.list[i])
        incProgress(1/N, detail = paste("Processing file", i, "of", N, ": ", id))
        # data Wingate
        if (exists("df2")) {
          rm(df2)
        }
        
        print(paste("Tisknu", i,"/", length(file.list), ": ", id))
        writeLines(iconv(readLines(paste(file.path, "/", file.list[i], sep = "")), from = "ANSI_X3.4-1986", to = "UTF8"), 
                   file(paste(file.path, "/", file.list[i], sep = ""), encoding="UTF-8"))
        df <- read.delim(paste(file.path, "/", file.list[i], sep = ""))
        if (an.input == "solo") {
          file.an <- file.list.an[which(file.list.an.bez == id)]
          antropo <- readxl::read_excel(paste(file.path.an, "/", file.an, sep=""), sheet = "Data_Sheet")
          if (srovnani == "YES") {
            antropo[order(as.Date(antropo$Date_measurement, format="%d/%m/%Y"), decreasing = TRUE),]
          }
        }
        
        if (an.input == "batch") {
          antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet") 
          antropo <- antropo[order(as.Date(antropo$Date_measurement, format="%d/%m/%Y"), decreasing = TRUE),]
          
        }
        
        
        # 3rd comparison (oldest)
        if (any(!is.na(file.list.compar_2.bez))) {
          if (srovnani == "YES" & id %in% file.list.compar_2.bez) {
            writeLines(iconv(readLines(paste(comparison_2.path, "/", id, ".txt", sep = "")), from = "ANSI_X3.4-1986", to = "UTF8"), 
                       file(paste(comparison_2.path, "/", id, ".txt", sep = ""), encoding="UTF-8"))
            compare.wingate_2 <- read.delim(paste(comparison_2.path, "/", id, ".txt", sep = ""))
            compare.wingate_2$Work.total..KJ. <- as.numeric(gsub(",",".", compare.wingate_2$Work.total..KJ.))
            compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. <- period_to_seconds(hms(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.))
            if (compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
              compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. <- compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.- compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.[1]
            } else if (tail(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.,1) > 31) {
              cas.zacatku <- tail(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.,1) - 30
              radek.zacatku <- which.min(abs(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. - cas.zacatku)) - 1
              compare.wingate_2 <- compare.wingate_2[-(1:radek.zacatku),]
            }
            if (compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
              compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. <- compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.- compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.[1] 
            }
            
            s5 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 5), n=1))
            s10 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 10), n=1))
            s15 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 15), n=1))
            s20 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 20), n=1))
            s25 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 25), n=1))
            s30 <- as.numeric(length(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.))
            radky5s <- round(base::mean(s5, (s10-s5), (s15-s10), (s20-s15), (s25-s20), (s30-s25)),0)
            compare.wingate_2$RM5_Power <- zoo::rollmean(compare.wingate_2$Power..W., k=radky5s, fill = NA, align = "right")
            
            for (j in 1:length(compare.wingate_2$Power..W.)) {
              compare.wingate_2$AvP_dopocet[j] <- mean(compare.wingate_2$Power..W.[1:j] )
            }
          }
        }
        
        # The latest comparison
        if (any(!is.na(file.list.compar.bez))) {
          if (srovnani == "YES" & id %in% file.list.compar.bez) {
            writeLines(iconv(readLines(paste(comparison.path, "/", id, ".txt", sep = "")), from = "ANSI_X3.4-1986", to = "UTF8"), 
                       file(paste(comparison.path, "/", id, ".txt", sep = ""), encoding="UTF-8"))
            compare.wingate <- read.delim(paste(comparison.path, "/", id, ".txt", sep = ""))
            compare.wingate$Work.total..KJ. <- as.numeric(gsub(",",".", compare.wingate$Work.total..KJ.))
            compare.wingate$Elapsed.time.total..h.mm.ss.hh. <- period_to_seconds(hms(compare.wingate$Elapsed.time.total..h.mm.ss.hh.))
            if (compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
              compare.wingate$Elapsed.time.total..h.mm.ss.hh. <- compare.wingate$Elapsed.time.total..h.mm.ss.hh.- compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1]
            } else if (tail(compare.wingate$Elapsed.time.total..h.mm.ss.hh.,1) > 31) {
              cas.zacatku <- tail(compare.wingate$Elapsed.time.total..h.mm.ss.hh.,1) - 30
              radek.zacatku <- which.min(abs(compare.wingate$Elapsed.time.total..h.mm.ss.hh. - cas.zacatku)) - 1
              compare.wingate <- compare.wingate[-(1:radek.zacatku),]
            }
            
            if (compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
              compare.wingate$Elapsed.time.total..h.mm.ss.hh. <- compare.wingate$Elapsed.time.total..h.mm.ss.hh.- compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1] 
            }
            
            s5 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 5), n=1))
            s10 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 10), n=1))
            s15 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 15), n=1))
            s20 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 20), n=1))
            s25 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 25), n=1))
            s30 <- as.numeric(length(compare.wingate$Elapsed.time.total..h.mm.ss.hh.))
            differences <- c(s5, s10, s15, s20, s25, s30) - c(0, s5, s10, s15, s20, s25)
            radky5s <- round(mean(differences), 0)
            compare.wingate$RM5_Power <- zoo::rollmean(compare.wingate$Power..W., k=radky5s, fill = NA, align = "right")
            
            for (j in 1:length(compare.wingate$Power..W.)) {
              compare.wingate$AvP_dopocet[j] <- mean(compare.wingate$Power..W.[1:j] )
            }
          }
        }
        
        
        df$Work.total..KJ. <- as.numeric(gsub(",",".", df$Work.total..KJ.))
        
        df <- df[!apply(is.na(df) | df == "", 1, all),]
        df$Elapsed.time.total..h.mm.ss.hh. <- period_to_seconds(hms(df$Elapsed.time.total..h.mm.ss.hh.))
        
        
        df$Elapsed.time.total..h.mm.ss.hh. <- as.numeric(df$Elapsed.time.total..h.mm.ss.hh.)  
        
        # cut the original Wingate to 30s
        if (df$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
          df$Elapsed.time.total..h.mm.ss.hh. <- df$Elapsed.time.total..h.mm.ss.hh.- df$Elapsed.time.total..h.mm.ss.hh.[1]
        } else if (tail(df$Elapsed.time.total..h.mm.ss.hh.,1) > 31) {
          cas.zacatku <- tail(df$Elapsed.time.total..h.mm.ss.hh.,1) - 30
          radek.zacatku <- which.min(abs(df$Elapsed.time.total..h.mm.ss.hh. - cas.zacatku)) - 1
          df <- df[-(1:radek.zacatku),]
          if (df$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
            df$Elapsed.time.total..h.mm.ss.hh. <- df$Elapsed.time.total..h.mm.ss.hh.- df$Elapsed.time.total..h.mm.ss.hh.[1]
          }
        }
        
        # find number of rows for 5 seconds of test
        s5 <- as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 5), n=1))
        s10 <- as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 10), n=1))
        s15 <-as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 15), n=1))
        s20 <- as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 20), n=1))
        s25 <- as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 25), n=1))
        s30 <- as.numeric(length(df$Elapsed.time.total..h.mm.ss.hh.))
        differences <- c(s5, s10, s15, s20, s25, s30) - c(0, s5, s10, s15, s20, s25)
        radky5s <- round(mean(differences), 0)
        df$RM5_Power <- zoo::rollmean(df$Power..W., k=radky5s, fill = NA, align = "right")
        
        for (j in 1:length(df$Power..W.)) {
          df$AvP_dopocet[j] <- mean(df$Power..W.[1:j] )
        }
        
        
        ####  Anthropometry values calculation ####
        
        if (an.input == "batch") {
          datum_mer <- head(as.Date(na.omit(antropo$Date_measurement[antropo$ID == id],"%d/%m/%Y")),1)
          sj <- NA
          if("SJ" %in% colnames(antropo)) {sj <- head(na.omit(antropo$SJ[antropo$ID == id]),1)}
          sj <- ifelse(is.empty(sj), NA, sj)
          vaha <- head(na.omit(antropo$Weight[antropo$ID == id]),1)
          vaha <- head(ifelse(length(vaha) == 0, NA, vaha),1)
          vyska <- head(na.omit(antropo$Height[antropo$ID == id]),1)
          vyska <- head(ifelse(length(vyska) == 0, NA, vyska),1)
          fat <- head(round(na.omit(antropo$Fat[antropo$ID == id]),1),1)
          fat <- ifelse(length(fat) == 0, NA, fat)
          ath <- head(round(na.omit(antropo$ATH[antropo$ID == id]),1),1)
          ath <- ifelse(length(ath) == 0, NA, ath)
          birth <- head(na.omit(antropo$Birth[antropo$ID == id]),1)
          fullname <- head(paste(na.omit(antropo$Name[antropo$ID == id]), na.omit(antropo$Surname[antropo$ID == id]), sep = " ", collapse = NULL),1)
          fullname.rev <- head(paste(na.omit(antropo$Surname[antropo$ID == id]), na.omit(antropo$Name[antropo$ID == id]), sep = " ", collapse = NULL),1)
          datum_nar <- head(format(as.Date(na.omit(antropo$Birth[antropo$ID == id])),"%d/%m/%Y"),1)
          datum_mer <- head(format(as.Date(na.omit(antropo$Date_measurement[antropo$ID == id])),"%d/%m/%Y"),1)
          la <- ifelse(!is.empty(head(antropo$LA[antropo$ID == id],1)), head(antropo$LA[antropo$ID == id],1), NA)
          age <- ifelse(!is.empty(antropo$Age[antropo$ID == id]), 
                        round(na.omit(antropo$Age[antropo$ID == id]),2), 
                        ifelse(!is.na(datum_nar) & datum_nar != "",
                               floor(as.integer(difftime(as.Date(datum_mer, format = "%d/%m/%Y"), as.Date(datum_nar, format = "%d/%m/%Y"), units = "days") / 365.25)),
                               NA))
          age <- head(age,1)
        } else {
          datum_mer <- head(format(as.Date(na.omit(antropo$Date_measurement[antropo$ID == id])),"%d/%m/%Y"),1)
          vaha <- na.omit(tail(antropo$Weight, n = 1))
          sj <- NA
          if("SJ" %in% colnames(antropo)) {sj <- na.omit(tail(antropo$SJ, n = 1))}
          sj <- ifelse(is.empty(sj), NA, sj)
          vaha <- ifelse(length(vaha) == 0, NA, vaha)
          vyska <- na.omit(tail(antropo$Height, n = 1))
          vyska <- ifelse(length(vyska) == 0, NA, vyska)
          fat <- round(na.omit(tail(antropo$Fat, n =1)),1)
          fat <- ifelse(length(fat) == 0, NA, fat)
          ath <- round(na.omit(tail(antropo$ATH, n =1)),1)
          ath <- ifelse(length(ath) == 0, NA, ath)
          birth <- na.omit(tail(antropo$Birth, n = 1))
          fullname <- paste(na.omit(tail(antropo$Name, n =1)), na.omit(tail(antropo$Surname, n = 1)), sep = " ", collapse = NULL)
          fullname.rev <- paste(tail(antropo$Surname, n = 1), tail(antropo$Name, n = 1), sep = " ", collapse = NULL)
          datum_nar <- format(as.Date(na.omit(tail(antropo$Birth, n =1))),"%d/%m/%Y")
          datum_mer <- head(format(as.Date(na.omit(antropo$Date_measurement[antropo$ID == id])),"%d/%m/%Y"),1)
          la <- tail(antropo$LA, n=1)
          age <- ifelse(!is.empty(antropo$Age[antropo$ID == id]), 
                        round(na.omit(tail(antropo$Age, n =1)),2), 
                        ifelse(!is.na(datum_nar) & datum_nar != "",
                               floor(as.integer(difftime(as.Date(datum_mer, format = "%d/%m/%Y"), as.Date(datum_nar, format = "%d/%m/%Y"), units = "days") / 365.25)),
                               NA))
        }
        
        speed <- NA
        pp <- max(df$Power..W.)
        minp <- min(df$Power..W.[round((length(df$Power..W.)/2),0):length(df$Power..W.)])
        pp5s <- round(max(df$RM5_Power, na.rm = T),1)
        minp5s <- round(base::min(df$RM5_Power, na.rm = T),1)
        drop <- pp - minp
        iu <- round(((pp5s-minp5s)/pp5s)*100,1)
        avrp <- round(mean(df$Power..W.),1)
        totalw <- round(mean(df$AvP_dopocet*30),0)
        pp5sradek <- which(df$RM5_Power == pp5s)
        an.cap <- round(totalw/vaha,2)
        
        
        # Line smoothing
        vyhlazeno <- smooth.spline(df$Elapsed.time.total..h.mm.ss.hh., df$Power..W., spar = 0.4)
        y <- predict(vyhlazeno, newdata = df)
        y <- y$y
        
        if(any(!is.na(file.list.compar.bez))) {
          if(srovnani == "YES" & id %in% file.list.compar.bez) {
            compare.wingate <- compare.wingate[rowSums(is.na(compare.wingate)) < 3, ]
            vyhlazeno.compare <- smooth.spline(compare.wingate$Elapsed.time.total..h.mm.ss.hh., compare.wingate$Power..W., spar = 0.4)
            y2 <- predict(vyhlazeno.compare, newdata = compare.wingate)
            y2 <- y2$y
          }
        }
        
        if(any(!is.na(file.list.compar_2.bez))) {
          if(srovnani == "YES" & id %in% file.list.compar_2.bez) {
            vyhlazeno.compare_2 <- smooth.spline(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh., compare.wingate_2$Power..W., spar = 0.4)
            y3 <- predict(vyhlazeno.compare_2, newdata = compare.wingate_2)
            y3 <- y3$y
          }
        }
        
        
        
        # Wingate reporting table
        columns <- c("Antropometrie", "V2", "Wingate test", "Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pětivteřinový průměr")
        df1 = data.frame(matrix(nrow = 8, ncol = length(columns))) 
        colnames(df1) <- columns
        df1$Antropometrie <-  `length<-`(c("Datum narození", "Věk", "Výška", "TH - Hmotnost", "Tuk", "ATH - Aktivní tělesná hmota"), nrow(df1))
        df1$`Wingate test` <- `length<-`(c("Maximální výkon", "Minimální výkon", "Pokles výkonu", "Celková práce", "Index Únavy", "Celkový počet otáček", "Laktát", "Maximální Tepová frekvence (BPM)"),nrow(df1))
        
        df1[1,2] <- ifelse(!is.empty(datum_nar), datum_nar, NA)
        df1[2,2] <- ifelse(!is.empty(age), 
                           floor(age), 
                           ifelse(!is.na(datum_nar) & datum_nar != "",
                                  floor(as.integer(difftime(as.Date(datum_mer, format = "%d/%m/%Y"), as.Date(datum_nar, format = "%d/%m/%Y"), units = "days") / 365.25)),
                                  NA))
        df1[3,2] <- paste(vyska, "cm", sep = " ")
        df1[4,2] <- paste(vaha, "kg", sep = " ")
        df1[5,2] <- paste(fat, "%", sep = " ")
        df1[6,2] <- paste(ath, "kg", sep = " ")
        
        df1[1,4] <- paste(pp, "W", sep = " ")
        df1[2,4] <- paste(minp, "W", sep = " ")
        df1[3,4] <- paste(drop, "W", sep = " ")
        df1[4,4] <- paste(round(totalw/1000, 1), "kJ", sep = " ")
        df1[5,4] <- paste(iu, "%", sep = " ")
        df1[6,4] <- round(tail(df$Turns.number..Nr.,1),0)
        df1[7,4] <- ifelse(!is.na(la), paste(la, "mmol/l", sep = " "), NA)
        df1[8,4] <- as.numeric(max(df$Heart.rate..bpm., na.rm =T))
        
        df1$Absolutní[df1$Absolutní == 0] <- NA
        df1$`Relativní (TH)`[1] <- paste(round(pp/vaha,1), "W/kg", sep = " ")
        df1$`Relativní (TH)`[2] <- paste(round(minp/vaha,1), "W/kg", sep = " ")
        df1$`Relativní (TH)`[3] <- paste(round(drop/vaha,1), "W/kg", sep = " ")
        df1$`Relativní (TH)`[4] <- paste(round(totalw/vaha,0), "J/kg", sep = " ")
        df1$`Relativní (ATH)`[1] <- paste(round(pp/ath,1), "W/kg", sep = " ")
        df1$`Relativní (ATH)`[2] <- paste(round(minp/ath,1), "W/kg", sep = " ")
        df1$`Relativní (ATH)`[3] <- paste(round(drop/ath,1), "W/kg", sep = " ")
        df1$`Relativní (ATH)`[4] <- paste(round(totalw/ath,0), "J/kg", sep = " ")
        
        df1[5,5:6] <- NA
        df1[6,5:6] <- NA
        df1[6,5:6] <- NA
        df1[7:8,5:6] <- NA
        
        df1[1,7] <- paste(pp5s, "W", sep = " ")
        df1[2,7] <- paste(minp5s, "W", sep = " ")
        
        
        dotaz_spiro_original <- dotaz_spiro
        
        
        if (length(spatne.spiro) > 0) {
          if (id %in% spatne.spiro) {
            dotaz_spiro <- "NO"  
          }
        }
        
        
        #### Spiroergomatery values  ####
        if (dotaz_spiro == "YES") {
          k <- which(file.list.spiro.bez==id)
          if (length(file.list.spiro) != 0) {
            # spiro <- readxl::read_excel("C:/Users/Dominik Kolinger/Desktop/VO2_Repre_Zeny/VO2_Max/Kosinova_Karolina.xlsx") 
            spiro <- readxl::read_excel(paste(spiro.path, "/", file.list.spiro[k], sep=""))
           
            spiro <- spiro[rowSums(!is.na(spiro)) > 0,]
            rows_with_bf <- which(spiro[[1]] == "BF")
            spiro_info <- spiro[1:rows_with_bf, ]
            spiro <- spiro[(rows_with_bf + 1):nrow(spiro), ]
            colnames(spiro) <- as.character(spiro[1, ])
            spiro <- spiro[-c(1,2), ]
            spiro <- subset(spiro, Fáze == "Zátěž")
            spiro[4:ncol(spiro)] <- lapply(spiro[4:ncol(spiro)], as.numeric)
            spiro$t <- gsub(",", ".", spiro$t)
            spiro$t <- as.POSIXct(spiro$t, format = "%H:%M:%OS", tz = "UTC")
            spiro$t <- spiro$t - min(spiro$t)
            time_intervals <- diff(spiro$t)
            median_interval <- as.numeric(median(time_intervals))
            closest_rows <- max(1, round(5 / median_interval))
            spiro$VT_5s <- zoo::rollmean(spiro$VT, k = closest_rows, align = "right", fill = NA)
          }
          
          if (units_switch) {
            speed <- as.numeric(spiro_info[which(spiro_info[,1] == "v"), 15])
          }
          s.datum.mer <- format(as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Počáteční čas")], format = "%d.%m.%Y %H:%M"), "%d/%m/%Y")
          s.datum.nar <- format(as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Datum narození")], format = "%d.%m.%Y"),"%d.%m.%Y")
          s.age <- floor(as.integer(difftime(as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Počáteční čas")], format = "%d.%m.%Y %H:%M"), as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Datum narození")], format = "%d.%m.%Y")), units = "days") / 365.25)
          s.height <- as.numeric(gsub("cm", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "Výška")][1])))  
          FVC <- as.numeric(gsub("L", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "VC")][1]))) 
          FEV1 <-  as.numeric(gsub("L", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "FEV1")][1])))
          s.pomer <- round(FEV1/FVC*100,1)
          VO2max <- round(max(spiro$`V'O2`), 1)
          VO2_kg <- as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "V'O2/kg"), 12]))
          vykon <- suppressWarnings(as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "WR"), 12])))
          if (is.na(vykon)) {
            vykon <- suppressWarnings(as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "WR"), 15])))
          }
          if (is.na(vykon)) {
            vykon <- as.numeric(round(max(spiro$WR, na.rm = TRUE), 0))
          }
          vykon_kg <- as.numeric(vykon/vaha)
          tep.kyslik <- as.numeric(spiro_info[which(spiro_info[,1] == "V'O2/HR"), 12])
          radky_spiro_info <- as.character(unlist(spiro_info[,1]))
          
          if ("TF" %in% radky_spiro_info) {
            index <- which(spiro_info[,1] == "TF")
          } else if ("SF" %in% radky_spiro_info) {
            index <- which(spiro_info[,1] == "SF")
            colnames(spiro)[colnames(spiro) == "SF"] <- "TF"
          } else {
            stop("Nebyl nalezen TF ani SF.")
          }
          anp <- as.numeric(ifelse(spiro_info[index, 9] != "-", 
                                   as.numeric(spiro_info[index, 9]), 
                                   as.numeric(spiro_info[index, 15]) * 0.85))
          
          hrmax <- round(max(spiro$TF),0)
          min.ventilace <- as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "V'E"), 12]))
          dech.frek <- spiro_info[which(spiro_info[,1] == "BF"), 12]
          dech.objem <- round(max(spiro$VT_5s[which.max(as.numeric(format(spiro$t, "%H%M%S")) > 100):nrow(spiro)]),1)
          dech.objem.per <- round((dech.objem/FVC)*100,0)
          rer <- round(max(spiro$RER),2)
          LaMax <- la
          
          
          #Spiroergometry table
          columns_s <- c("Antropometrie", "Hodnota", "V3", "Spiroergometrie", "Absolutní", "Relativní (TH)", "% z nál. hodnoty" ,"Relativní (ATH)")
          df3 = data.frame(matrix(nrow = 14, ncol = length(columns_s))) 
          colnames(df3) <- columns_s
          df3$Antropometrie <-  `length<-`(c("Datum narození", "Věk", "Výška", "TH - Hmotnost", "Tuk", "ATH - Aktivní tělesná hmota", NA, "Klidová ventilace", "FVC", "FEV1", "Poměr FVC/FEV1"), nrow(df3))
          df3$Spiroergometrie <- `length<-`(c("VO2max", "Dosažený výkon", "HrMax", "Tepový kyslík", "ANP", NA, NA, "Ventilace", "Minutová ventilace", "Dechová frekvence", "Dechový objem", "Dechový objem %", "RER", "LaMax"),nrow(df3))
          df3$V3[8] <- "% z nál. hodnoty"
          df3$`Relativní (TH)`[8] <- "Optimum"
          
          df3[1,2] <- datum_nar
          df3[2,2] <- age
          df3[3,2] <-  paste(s.height, "cm", sep=" ")
          df3[4,2] <- paste(vaha, "kg", sep=" ")
          df3[5,2] <- paste(fat, "%", sep=" ")
          df3[6,2] <- paste(ath, "kg", sep=" ")
          df3[9,2] <- paste(FVC,"l", sep=" ")
          df3[10,2] <- paste(FEV1,"l", sep=" ")
          df3[11,2] <- paste(s.pomer,"%", sep=" ")
          df3[9,3] <- round((FVC/sports_data$FVC[sports_data$Sport==sport])*100,0)
          df3[10,3] <- round((FEV1/sports_data$FEV1[sports_data$Sport==sport])*100,0)
          df3[1,7] <- round((VO2_kg/sports_data$VO2max[sports_data$Sport==sport])*100,0)
          df3[2,7] <- round(((vykon/vaha)/sports_data$PowerVO2[sports_data$Sport==sport])*100,0)
          df3[1,5] <- paste(VO2max,"l", sep=" ")
          if (units_switch) {
            df3[2,5] <- paste(speed,"km/h", sep=" ")
          } else {
            df3[2,5] <- paste(vykon,"W", sep=" ")
          }
          df3[3,5] <- paste(hrmax,"BPM", sep=" ")
          df3[4,5] <- ifelse(is.na(tep.kyslik), tep.kyslik, paste(tep.kyslik,"ml/tep", sep=" "))
          df3[5,5] <- paste(round(anp,0),"BPM", sep=" ")
          df3[9,5] <- paste(min.ventilace,"l/min", sep=" ")
          df3[10,5] <- paste(dech.frek,"d/min", sep=" ")
          df3[11,5] <- paste(dech.objem,"l", sep=" ")
          df3[12,5] <- paste(dech.objem.per,"%", sep=" ")
          df3[13,5] <- rer
          df3[14,5] <- ifelse(is.empty(la), NA, paste(la, "mmol/l", sep = " "))
          df3[1,6] <- paste(VO2_kg, "ml/min/kg", sep=" ")
          df3[2,6] <- paste(round(vykon/vaha, 1), "W/kg", sep=" ")
          df3[1,8] <- paste(round((VO2max*1000)/ath,1), "ml/min/kg", sep=" ")
          if (units_switch) {
            df3[2,8] <- "-"
          } else {
            df3[2,8] <- paste(round(vykon/ath, 1), "W/kg", sep=" ")
          }
          
          df3[9,6] <- round(30*FVC,0)
          df3[10,6] <- "50-60"
          df3[12,6] <- "50-60"
          df3[13,6] <- "1.08-1.18"
          df3 <- df3[-7,]
          
          
          # Training zones definition
          columns_tz <- c("Zóna", "od (BPM)", "do (BPM)")
          tz = data.frame(matrix(nrow = 3, ncol = length(columns_tz))) 
          colnames(tz) <- columns_tz
          tz$Zóna <- c("Aerobní", "Smíšená", "Anaerobní")
          tz$`od (BPM)`[2] <- round((anp / 0.85)*0.76,0)
          tz$`od (BPM)`[3] <- round(anp,0)
          tz$`do (BPM)`[1] <- round((anp / 0.85)*0.75,0)
          tz$`do (BPM)`[2] <- round((anp / 0.85)*0.84,0)
          
          
          vo2_range <- range(spiro$`V'O2`)
          coeff <- vo2_range * 10 
          
          
          
          #### Spiroergometry history table ####
          if (!exists("df4")) {
            df4 <- data.frame()
            df4 <- df4[nrow(df4) + 1,]
          }
          
          
          df4$`Date meas.` <- append(na.omit(df4$`Date meas.`), datum_mer)
          df4$Name <- append(na.omit(df4$Name), fullname)
          df4$Weight <- append(na.omit(df4$Weight), vaha)
          df4$Fat <- append(na.omit(df4$Fat), fat)
          df4$`VO2max (l)` <- append(na.omit(df4$`VO2max (l)`), VO2max)
          df4$`VO2max (ml/kg/min)` <- append(na.omit(df4$`VO2max (ml/kg/min)`), VO2_kg)
          
          if (units_switch) {
            df4$`Rychlost (km/h)` <- append(na.omit(df4$`Rychlost (km/h)`), speed)
          } else {
            df4$`Výkon (W)` <- append(na.omit(df4$`Výkon (W)`), vykon)
          }
          df4$`Výkon (l/kg)` <- append(na.omit( df4$`Výkon (l/kg)`), round(vykon_kg,1))
          df4$`HRmax (BPM)` <- append(na.omit(df4$`HRmax (BPM)`), hrmax)
          df4$`ANP (BPM)` <- append(na.omit(df4$`ANP (BPM)`), round(anp,0))
          df4$`Tep. kyslík (ml)` <- append(df4$`Tep. kyslík (ml)`[1:length(df4$`Tep. kyslík (ml)`)-1], tep.kyslik)
          df4$`VT (l)` <- append(na.omit(df4$`VT (l)`), dech.objem)
          df4$RER <- append(na.omit(df4$RER), rer)
          df4$`LaMax (mmol/l)` <- append(na.omit(df4$`LaMax (mmol/l)`), la)
          df4$`FEV1 (l)` <- append(df4$`FEV1 (l)`[1:length(df4$`FEV1 (l)`)-1], FEV1)
          df4$`FVC (l)` <- append(na.omit(df4$`FVC (l)`), FVC)
          df4$`Aerobní Z. do` <- append(na.omit(df4$`Aerobní Z. do`), round((anp / 0.85)*0.75,0))
          df4$`Smíšená Z. od` <- append(na.omit(df4$`Smíšená Z. od` ), round((anp / 0.85)*0.76,0))
          df4$`Smíšená Z. do` <- append(na.omit(df4$`Smíšená Z. do`), round((anp / 0.85)*0.84,0))
          df4$`Anaerobní Z. od` <- append(na.omit(df4$`Anaerobní Z. od`), anp)
          df4 <- add_row(df4)
          
          
          hrmin <- round(min(spiro$TF),0)
          
          # Y axis limits calculation
          ylim.prim <- c(0, min.ventilace*1.1)
          ylim.sec <- c(hrmin*0.9, hrmax*1.05)  
          
          b <- diff(ylim.prim)/diff(ylim.sec)
          a <- ylim.prim[1] - b*ylim.sec[1]
          
          
          #### Spiroergometry plots ####
          plot2 <- ggplot(spiro, aes(x = t)) +
            geom_line(aes(y = a + TF*b, color = "Tepová frekvence (BPM)", linetype = "Tepová frekvence (BPM)"), linewidth = 1) +
            geom_line(aes(y = `V'E`, color = "Minutová ventilace (l/min)", linetype = "Minutová ventilace (l/min)"), linewidth = 0.5, alpha = 0.2, linetype = "dashed") +
            geom_smooth(aes(y = `V'E`), color = "black", linetype = "solid", size = 1, method = "loess", se = FALSE, span = 0.2) +
            scale_y_continuous(
              name = "Minutová ventilace (l/min)", 
              labels = scales::comma,
              sec.axis = sec_axis(~ (. - a)/b, name = "Tepová frekvence (BPM)")
            ) +
            scale_x_datetime(labels = scales::date_format("%M:%S"), date_breaks = "1 min") +
            theme_classic() +
            theme(
              text = element_text(size = 30),
              panel.grid.major = element_blank(),
              panel.grid.minor = element_blank(),
              axis.line = element_line(colour = "black", linewidth = 1.5),
              axis.text.x = element_text(angle = -45, hjust = 0),
              axis.title.y = element_text(size = 25),
              legend.position = "bottom"
            ) +
            scale_linetype_manual(values = c('solid', 'solid'), 
                                  labels = c("Minutová ventilace (l/min)", "Tepová frekvence (BPM)")) +
            scale_color_manual(values = c("Minutová ventilace (l/min)" = "black", "Tepová frekvence (BPM)" = "orange")) +
            xlab(expression("Čas")) +
            ylab("Hodnota") +
            labs(colour="") +
            guides(
              color = guide_legend(override.aes = list(linetype = c("solid", "solid"))), 
              linetype = "none"
            )
          
          plot3 <- ggplot(spiro, aes(x = t)) +
            geom_line(aes(y = `V'O2`), color = "darkmagenta", linewidth = 0.5, alpha = 0.2, linetype = "dashed") +
            geom_smooth(aes(y = `V'O2`), color = "darkmagenta", linetype = "solid", size = 1, method = "loess", se = FALSE, span = 0.3) +
            scale_x_datetime(labels = scales::date_format("%M:%S"), date_breaks = "1 min") +
            scale_y_continuous(
              name = "Spotřeba kyslíku (l)", 
              labels = scales::comma
            ) +
            theme_classic() +
            theme(
              text = element_text(size = 30),
              panel.grid.major = element_blank(),
              panel.grid.minor = element_blank(),
              axis.line = element_line(colour = "black", linewidth = 1.5),
              axis.text.x = element_text(angle = -45, hjust = 0),
              axis.title.y = element_text(size = 25),
              legend.position = "none"
            ) +
            xlab(expression("Čas")) +
            ylab("Spotřeba kyslíku (l)") 
          
          
          
          combined_plot <- plot2 + plot3 +
            plot_annotation(
              title = 'Spiroergometrie',
              caption = '',
              theme = theme(plot.title = element_text(size = 30, hjust = 0.5))
            )
        }
        
        
        #### Wingate values calculation ####
        if (any(!is.na(file.list.compar_2.bez))) {
          if (srovnani == "YES" & id %in% file.list.compar_2.bez) {
            antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet") 
            antropo <- antropo %>% 
              filter(ID==id)
            antropo <- antropo[order(as.Date(antropo$Date_measurement, format="%d/%m/%Y")),]
            df2 <- data.frame()
            df2 <- df2[nrow(df2) + 1,]
            df2$Date_meas. <- format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-2]), "%d/%m/%Y")
            df2$Weight <- antropo$Weight[length(antropo$Weight)-2]
            df2$`Fat (%)` <- round(antropo$Fat[length(antropo$Fat)-2], 1)
            df2$ATH <- round(antropo$ATH[length(antropo$ATH)-2], 1)
            df2$`Pmax (W)` <- max(compare.wingate_2$Power..W.)
            df2$`Pmax_kg (W/kg)` <- round(max(compare.wingate_2$Power..W.)/antropo$Weight[length(antropo$Weight)-2],2)
            df2$`Pmin (W)` <- min(compare.wingate_2$Power..W.[round((length(compare.wingate_2$Power..W.)/2),0):length(compare.wingate_2$Power..W.)])
            df2$`Avg_power (W)` <- round(mean(compare.wingate_2$Power..W.),1)
            df2$`Pdrop (W)` <- df2$`Pmax (W)` - df2$`Pmin (W)`
            df2$`Pmax_5s (W)` <- round(max(compare.wingate_2$RM5_Power, na.rm = T),1)
            df2$`IU (%)` <- round((max(na.omit(compare.wingate_2$RM5_Power))-round(min(na.omit(compare.wingate_2$RM5_Power)),1))/max(na.omit(compare.wingate_2$RM5_Power))*100,1)
            df2$`Work (kJ)` <- round(mean(compare.wingate_2$AvP_dopocet*30),0)/1000
            df2$`HR_max (BPM)` <- max(compare.wingate_2$Heart.rate..bpm., na.rm =T)
            df2$`La_max (mmol/l)` <- antropo$LA[length(antropo$LA)-1]
            df2[nrow(df2) + 1,] <- NA
          }
        }
        
        if (any(!is.na(file.list.compar.bez))) {
          if (srovnani == "YES" & id %in% file.list.compar.bez) {
            if (exists("duplikaty_check")) {
              antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet") 
              antropo <- antropo %>% 
                filter(ID==id)
              antropo <- antropo[order(as.Date(antropo$Date_measurement, format = "%d/%m/%Y")),]
            }
            if (!exists("df2")) {
              df2 <- data.frame()
              df2 <- df2[nrow(df2) + 1,]
            }
            df2$Date_meas. <- append(na.omit(df2$Date_meas.), format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-1]), "%d/%m/%Y"))
            df2$Weight <- append(na.omit(df2$Weight), antropo$Weight[length(antropo$Weight)-1])
            df2$`Fat (%)` <- append(na.omit(df2$`Fat (%)`), round(antropo$Fat[length(antropo$Fat)-1], 1))
            df2$ATH <- append(na.omit(df2$ATH), round(antropo$ATH[length(antropo$ATH)-1], 1))
            df2$`Pmax (W)` <- append(na.omit(df2$`Pmax (W)`), max(compare.wingate$Power..W.))
            df2$`Pmax_kg (W/kg)` <- append(na.omit(df2$`Pmax_kg (W/kg)`), round(max(compare.wingate$Power..W.)/antropo$Weight[length(antropo$Weight)-1],2))
            df2$`Pmin (W)` <- append(na.omit(df2$`Pmin (W)`), min(compare.wingate$Power..W.[round((length(compare.wingate$Power..W.)/2),0):length(compare.wingate$Power..W.)]))
            df2$`Avg_power (W)` <- append(na.omit(df2$`Avg_power (W)`), round(mean(compare.wingate$Power..W.),1))
            df2$`Pdrop (W)` <- append(na.omit(df2$`Pdrop (W)`), (df2$`Pmax (W)`[nrow(df2)] - df2$`Pmin (W)`[nrow(df2)]))
            df2$`Pmax_5s (W)` <- append(na.omit(df2$`Pmax_5s (W)`), round(max(compare.wingate$RM5_Power, na.rm = T),1))
            df2$`IU (%)` <- append(na.omit(df2$`IU (%)`), round((max(na.omit(compare.wingate$RM5_Power))-round(min(na.omit(compare.wingate$RM5_Power)),1))/max(na.omit(compare.wingate$RM5_Power))*100,1))
            df2$`Work (kJ)` <- append(na.omit(df2$`Work (kJ)`), round(mean(compare.wingate$AvP_dopocet*30),0)/1000)
            df2$`HR_max (BPM)` <- append(na.omit(df2$`HR_max (BPM)`), max(compare.wingate$Heart.rate..bpm., na.rm =T))
            df2$`La_max (mmol/l)` <- append(na.omit(df2$`La_max (mmol/l)`), antropo$LA[length(antropo$LA)-1])
            df2[nrow(df2) + 1,] <- NA
          }
        }
        
        
        #Wingate history table
        if (!exists("df2")) {
          df2 <- data.frame()
          df2 <- df2[nrow(df2) + 1,]
        }
        
        
        df2$Date_meas. <- append(na.omit(df2$Date_meas.), datum_mer)
        df2$Weight <- append(na.omit(df2$Weight), vaha)
        df2$`Fat (%)` <- append(na.omit(df2$`Fat (%)`), fat)
        df2$ATH <- append(na.omit(df2$ATH), ath)
        df2$`Pmax (W)` <- append(na.omit(df2$`Pmax (W)`), pp)
        df2$`Pmax_kg (W/kg)` <- append(na.omit(df2$`Pmax_kg (W/kg)`), round(pp/vaha,2))
        df2$`Pmin (W)` <- append(na.omit(df2$`Pmin (W)`), minp)
        df2$`Avg_power (W)` <- append(na.omit(df2$`Avg_power (W)`), avrp)
        df2$`Pdrop (W)` <- append(na.omit(df2$`Pdrop (W)`), drop)
        df2$`Pmax_5s (W)` <- append(na.omit(df2$`Pmax_5s (W)`), pp5s)
        df2$`IU (%)` <- append(na.omit(df2$`IU (%)`), iu)
        df2$`Work (kJ)` <- append(na.omit(df2$`Work (kJ)`), totalw/1000)
        df2$`HR_max (BPM)` <- as.numeric(append(na.omit(df2$`HR_max (BPM)`), max(df$Heart.rate..bpm., na.rm =T)))
        df2$`La_max (mmol/l)` <- append(na.omit(df2$`La_max (mmol/l)`), la)
        
        
        
        
        #### Wingate plots ####
        my_breaks <- seq(0, tail(df$Elapsed.time.total..h.mm.ss.hh.,n=1) + 5, by = 5)
        
        if (any(!is.na(file.list.compar_2.bez)) && srovnani == "YES" && id %in% file.list.compar_2.bez && tri_graf == "YES") {
          plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
            theme_classic()  + 
            theme(text = element_text(size = 30),
                  panel.grid.major = element_line(color = "grey90",
                                                  linewidth = 0.1,
                                                  linetype = 1),
                  panel.grid.minor.y = element_line(color = "grey90",
                                                    linewidth = 0.1,
                                                    linetype = 1),
                  axis.line = element_line(colour = "black",
                                           linewidth = 1.5)) +
            geom_line(data = compare.wingate_2, aes(y = y3, color = "darkmagenta", linetype = "dashed"), lwd = 0.5) +
            geom_line(data = compare.wingate, aes(y = y2, color = "orange", linetype = "solid"), lwd = 0.9) + 
            geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
            geom_line(aes(color = "grey", linetype = "dashed")) +
            geom_line(data = df, aes(y=y, x = Elapsed.time.total..h.mm.ss.hh., color = "black", linetype = "solid"), lwd = 0.9) + 
            scale_linetype_manual(values = c('dashed', 'solid')) +
            scale_colour_identity(guide = "legend", breaks = c("black", "grey", "blue", "orange", "darkmagenta"), labels = c("black"=paste("vyhlazená data,", datum_mer, sep = " "), "blue"="průměr měření","grey"="hrubá data","orange"=format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-1]), "%d/%m/%Y"), "darkmagenta"=format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-2]), "%d/%m/%Y"))) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed", "solid", "dashed")))) + 
            guides(linetype = "none") + 
            xlab(expression("Čas (s)")) + 
            ylab("Výkon (W)") + 
            ggtitle("Anaerobní Wingate test - 30 s")  + 
            labs(colour="") +
            scale_x_continuous(breaks = my_breaks)  
        } else if ((tri_graf == "NO" & id %in% file.list.compar.bez) |
                   !(id %in% file.list.compar_2.bez) & id %in% file.list.compar.bez) {
          plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
            theme_classic()  + 
            theme(text = element_text(size = 30),
                  panel.grid.major = element_line(color = "grey90",
                                                  linewidth = 0.1,
                                                  linetype = 1),
                  panel.grid.minor.y = element_line(color = "grey90",
                                                    linewidth = 0.1,
                                                    linetype = 1),
                  axis.line = element_line(colour = "black",
                                           linewidth = 1.5)) +
            geom_line(data = compare.wingate, aes(y = y2, color = "orange", linetype = "solid"), lwd = 0.9) + 
            geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
            geom_line(aes(color = "grey", linetype = "dashed")) +
            geom_line(data = df, aes(y=y, x = Elapsed.time.total..h.mm.ss.hh., color = "black", linetype = "solid"), lwd = 0.9) + 
            scale_linetype_manual(values = c('dashed', 'solid')) +
            scale_colour_identity(guide = "legend", breaks = c("black", "grey", "blue", "orange"), labels = c("black"=paste("vyhlazená data,", datum_mer, sep = " "), "blue"="průměr měření","grey"="hrubá data","orange"=format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-1]), "%d/%m/%Y"))) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed", "solid")))) + 
            guides(linetype = "none") + 
            xlab(expression("Čas (s)")) + 
            ylab("Výkon (W)") + 
            ggtitle("Anaerobní Wingate test - 30 s")  + 
            labs(colour="") +
            scale_x_continuous(breaks = my_breaks)  
        } else if ((any(!is.na(file.list.compar.bez))) & is.na(file.list.compar_2.bez)) {
          if (srovnani == "YES" & id %in% file.list.compar.bez) {
            plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
              theme_classic()  + 
              theme(text = element_text(size = 30),
                    panel.grid.major = element_line(color = "grey90",
                                                    linewidth = 0.1,
                                                    linetype = 1),
                    panel.grid.minor.y = element_line(color = "grey90",
                                                      linewidth = 0.1,
                                                      linetype = 1),
                    axis.line = element_line(colour = "black",
                                             linewidth = 1.5)) +
              geom_line(aes(color = "grey", linetype = "dashed")) +
              geom_line(data = compare.wingate, aes(y = y2, color = "orange", linetype = "solid"), lwd = 0.9) + 
              geom_line(data = df, aes(y=y, x = Elapsed.time.total..h.mm.ss.hh., color = "black", linetype = "solid"), lwd = 0.9) + 
              geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
              scale_linetype_manual(values = c('dashed', 'solid')) +
              scale_colour_identity(guide = "legend", labels = c(paste("vyhlazená data,", datum_mer, sep = " "), "průměr měření", "hrubá data", format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-1]), "%d/%m/%Y"))) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "solid", "dashed", "solid")))) + 
              guides(linetype = "none") + 
              xlab(expression("Čas (s)")) + 
              ylab("Výkon (W)") + 
              ggtitle("Anaerobní Wingate test - 30 s")  + 
              labs(colour="") +
              scale_x_continuous(breaks = my_breaks)
          } else {
            plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
              theme_classic()  + 
              theme(text = element_text(size = 30),
                    panel.grid.major = element_line(color = "grey90",
                                                    linewidth = 0.1,
                                                    linetype = 1),
                    panel.grid.minor.y = element_line(color = "grey90",
                                                      linewidth = 0.1,
                                                      linetype = 1),
                    axis.line = element_line(colour = "black",
                                             linewidth = 1.5)) +
              geom_line(aes(color = "grey", linetype = "dashed")) +
              geom_line(aes(y=y, color = "black", linetype = "solid"), lwd = 0.9) + 
              geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
              scale_linetype_manual(values = c('dashed', 'solid')) +
              scale_colour_identity(guide = "legend", labels = c(paste("vyhlazená data,", datum_mer, sep = " "), "průměr měření", "hrubá data")) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed")))) + 
              guides(linetype = "none") + 
              xlab(expression("Čas (s)")) + 
              ylab("Výkon (W)") + 
              ggtitle("Anaerobní Wingate test - 30 s")  + 
              labs(colour="") +
              scale_x_continuous(breaks = my_breaks)  
          }
        } else { 
          plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
            theme_classic()  + 
            theme(text = element_text(size = 30),
                  panel.grid.major = element_line(color = "grey90",
                                                  linewidth = 0.1,
                                                  linetype = 1),
                  panel.grid.minor.y = element_line(color = "grey90",
                                                    linewidth = 0.1,
                                                    linetype = 1),
                  axis.line = element_line(colour = "black",
                                           linewidth = 1.5)) +
            geom_line(aes(color = "grey", linetype = "dashed")) +
            geom_line(aes(y=y, color = "black", linetype = "solid"), lwd = 0.9) + 
            geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
            scale_linetype_manual(values = c('dashed', 'solid')) +
            scale_colour_identity(guide = "legend", labels = c(paste("vyhlazená data,", datum_mer, sep = " "), "průměr měření", "hrubá data")) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed")))) + 
            guides(linetype = "none") + 
            xlab(expression("Čas (s)")) + 
            ylab("Výkon (W)") + 
            ggtitle("Anaerobní Wingate test - 30 s")  + 
            labs(colour="") +
            scale_x_continuous(breaks = my_breaks)  
        }
        
        
        
        
        #### Export of spiro report tables ####
        table1 <- df1 %>% gt() %>% 
          tab_header(title = fullname, subtitle = paste("Datum měření: ",datum_mer, sep = "")) %>% 
          cols_label(V2 = "") %>% sub_missing(columns = everything(),
                                              rows = everything(),
                                              missing_text = "")   %>%
          tab_style(
            style = cell_text(weight = "bold"),
            locations = cells_column_labels()) %>%
          tab_style(
            style = cell_text(weight = "bold"),
            locations = cells_body(
              columns = c("Antropometrie", "Wingate test") 
            )
          ) %>%
          tab_style(
            style = "padding-right:30px",
            locations = cells_column_labels()
          )   %>%
          tab_style(
            style = "padding-right:30px",
            locations = cells_body()) %>% 
          tab_options(
            table.width = pct(c(100)),
            container.width = 1500,
            table.font.size = px(18L),
            heading.padding = pct(0)) %>%
          gtExtras::gt_add_divider(columns = "V2", style = "solid", color = "#808080") %>% 
          opt_table_lines(extent = "none") %>%
          tab_style(
            style = cell_borders(
              sides = "top",
              color = "#808080",
              style = "solid"
            ),
            locations = 
              cells_body(
                rows = 1
              )
          ) %>% tab_style(
            style = cell_text(align = "right"),
            locations = cells_column_labels(columns = c("Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pětivteřinový průměr"))) %>% 
          cols_align(
            align = c("center"),
            columns = c("V2", "Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pětivteřinový průměr")
          ) %>% tab_style(
            style = cell_text(align = "center"),
            locations = cells_body(columns = c("V2", "Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pětivteřinový průměr")))
        
        
        #### History table - Wingate ####
        table2 <- df2 %>% gt() %>% tab_header(title = "Historie měření") %>% 
          cols_align(
            align = c("center")) %>% 
          tab_options(
            table.width = pct(c(100)),
            container.width = 1500,
            table.font.size = px(18L)) %>% 
          cols_label(
            Date_meas. = "Datum",
            Weight = "Váha (kg)",
            ATH = "ATH (kg)",
            `Fat (%)`  = "Tuk (%)",
            `Avg_power (W)` = "Průměr P (W)", 
            `Pdrop (W)` = "Pokles (W)",
            `Pmax_5s (W)` = "Pmax 5s (W)", 
            `Pmax_kg (W/kg)` = "Pmax (W/kg)",
            `IU (%)` = "IU (%)",
            `HR_max (BPM)` = "TF (BPM)",
            `La_max (mmol/l)` = "LA (mmol/l)",
            `Pmax (W)` = "Pmax (W)",
            `Pmin (W)` = "Pmin (W)",
            `Work (kJ)` = "Práce (kJ)"
          ) %>% 
          cols_width(
            c("Fat (%)", "ATH") ~ pct(6),
            c("Avg_power (W)", "La_max (mmol/l)") ~ pct(10.5),
            c("Pmax_5s (W)", "Pdrop (W)", "Pmax (W)", "Weight", "Work (kJ)", "IU (%)") ~ pct(8),
            c("Date_meas.") ~ pct(7.2),
            c("Pmin (W)", "Pmax (W)",  "HR_max (BPM)") ~ pct(7),
            c("Pmax_kg (W/kg)") ~ pct(9)  
          ) %>% 
          sub_missing(columns = everything(), rows = everything(), missing_text = "")
        
        #tabulka spiro
        if (dotaz_spiro == "YES") {
          table3 <- df3 %>% 
            gt() %>% 
            tab_header(
              title = fullname, 
              subtitle = paste("Datum měření: ", datum_mer, sep = "")
            ) %>% 
            cols_label(Hodnota = "") %>% 
            cols_label(V3 = "")   %>%
            tab_style(
              style = cell_text(weight = "bold"),
              locations = cells_column_labels()
            ) %>%
            tab_style(
              style = cell_text(weight = "bold"),
              locations = cells_body(
                columns = c("Antropometrie", "Spiroergometrie")
              )
            ) %>%
            tab_style(
              style = cell_text(weight = "bold"),
              locations = cells_body(
                rows = 7,
                columns = everything()  # Apply bold style to all columns in row 9
              )
            ) %>%
            tab_style(
              style = "padding-right:30px",
              locations = cells_column_labels()
            ) %>%
            tab_style(
              style = "padding-right:30px",
              locations = cells_body()
            ) %>% 
            tab_options(
              table.width = pct(c(100)),
              container.width = 1600,
              container.height = 670,
              table.font.size = px(18L),
              heading.padding = pct(0)
            ) %>%
            gtExtras::gt_add_divider(columns = "V3", style = "solid", color = "#808080") %>% 
            opt_table_lines(extent = "none") %>%
            tab_style(
              style = cell_borders(
                sides = "top",
                color = "#808080",
                style = "solid"
              ),
              locations = cells_body(
                rows = c(1,8)
              )
            )  %>% 
            tab_style(
              style = cell_text(align = "right"),
              locations = cells_column_labels(columns = c("Absolutní", "Relativní (TH)", "% z nál. hodnoty", "Relativní (ATH)"))
            ) %>% 
            cols_align(
              align = c("center"),
              columns = c("V3", "Absolutní", "Relativní (TH)", "% z nál. hodnoty", "Relativní (ATH)")
            ) %>% 
            tab_style(
              style = cell_text(align = "center"),
              locations = cells_body(columns = c("V3", "Absolutní", "Relativní (TH)", "% z nál. hodnoty", "Relativní (ATH)"))
            ) %>%
            tab_style(
              style = "padding-bottom:20px",
              locations = cells_body(rows = 6)
            ) %>% 
            sub_missing(columns = everything(), rows = everything(), missing_text = "") 
          
          table4  <- tz %>% gt() %>% 
            tab_header(title = "Tréninkové zóny") %>% 
            cols_label(Zóna = "")  %>% sub_missing(columns = everything(),
                                                   rows = everything(),
                                                   missing_text = "") %>% 
            tab_options(
              table.width = pct(c(100)),
              container.width = 300,
              container.height = 600,
              table.font.size = px(18L),
              heading.padding = pct(0))
          
          gtsave(
            table4,
            paste0(program_slozka, "/t4.html") # Ensure styles are applied
          )
          
          webshot(
            url = paste0(program_slozka, "/t4.html"), # Path to the saved HTML
            file = paste0(program_slozka, "/t4.png"), # Save as PNG
            vwidth = 350,  # Set viewport width
            vheight = 300,
            zoom = 2,
            cliprect = c(0, 0, 350, 300)# Set viewport height
          )
          
          png(paste0(program_slozka, "/p2.png"), width = 1500, height = 600)
          plot(combined_plot)
          dev.off()
          
          table3 %>%
            gtsave(paste0(program_slozka, "/t3.png"), vwidth = 1700, vheight = 1000) 
        }
        
        
        table2 %>%
          gt::gtsave(paste0(program_slozka, "/t2.png"), vwidth = 1500)  
        
        table1 %>%
          gtsave(paste0(program_slozka, "/t1.png"), vwidth = 1500, vheight = 1000)
        
        png(paste0(program_slozka, "/p1.png"), width = 2000, height = 600)
        plot(plot1)
        dev.off()
        #### Radarchart ####
        
        hg_anckg <- round(((an.cap/sports_data$ANC[sports_data$Sport==sport])*100),1)
        hg_ppkg <- round((((pp/vaha)/sports_data$Pmax[sports_data$Sport==sport])*100),1)
        
        
        if (!is.na(sj)) {
          hg_sj <- round(((sj/sports_data$SJ[sports_data$Sport==sport])*100),1)
          data_radar <- data.frame(
            `Anaerobní kapacita (J/kg) %` = hg_anckg,
            `Výkon Wingate  (W/kg) %` = hg_ppkg,
            `Squat Jump (cm) %`= hg_sj,
            check.names = FALSE
          )} else {
            data_radar <- data.frame(
              `Anaerobní kapacita (J/kg) %` = hg_anckg,
              `Výkon Wingate  (W/kg) %` = hg_ppkg,
              check.names = FALSE
            )
          }
        
        
        if (dotaz_spiro == "YES") {
          hg_Vo2max <- round(((VO2_kg / sports_data$VO2max[sports_data$Sport==sport]) * 100), 1)
          hg_VO2vykon <- round(((vykon_kg / sports_data$PowerVO2[sports_data$Sport==sport]) * 100), 1)
          data_radar <- cbind( `VO2Max (ml/min/kg) %` = hg_Vo2max, data_radar, `Výkon VO2Max (W/kg) %` = hg_VO2vykon)
        }
        
        
        max_values <- rep(100, ncol(data_radar))
        min_values <- rep(0, ncol(data_radar))
        
        data_radar <- rbind(max_values, min_values, data_radar)
        
        
        
        radarchart2 <- function(df, axistype=0, seg=4, pty=16, pcol=1:8, plty=1:6, plwd=1,
                                pdensity=NULL, pangle=45, pfcol=NA, cglty=3, cglwd=1,
                                cglcol="navy", axislabcol="blue", title="", maxmin=TRUE,
                                na.itp=TRUE, centerzero=FALSE, vlabels=NULL, vlcex=NULL,
                                caxislabels=NULL, calcex=NULL,
                                paxislabels=NULL, palcex=NULL, ...) {
          if (!is.data.frame(df)) { cat("The data must be given as dataframe.\n"); return() }
          if ((n <- length(df))<3) { cat("The number of variables must be 3 or more.\n"); return() }
          if (maxmin==FALSE) { # when the dataframe does not include max and min as the top 2 rows.
            dfmax <- apply(df, 2, max)
            dfmin <- apply(df, 2, min)
            df <- rbind(dfmax, dfmin, df)
          }
          plot(c(-1.2, 1.2), c(-1.2, 1.2), type="n", frame.plot=FALSE, axes=FALSE, 
               xlab="", ylab="", main=title, cex.main=2, asp=1, ...) # define x-y coordinates without any plot
          theta <- seq(90, 450, length=n+1)*pi/180
          theta <- theta[1:n]
          xx <- cos(theta)
          yy <- sin(theta)
          CGap <- ifelse(centerzero, 0, 1)
          for (i in 0:seg) { # complementary guide lines, dotted navy line by default
            polygon(xx*(i+CGap)/(seg+CGap), yy*(i+CGap)/(seg+CGap), lty=cglty, lwd=cglwd, border=cglcol)
            if (axistype==1|axistype==3) CAXISLABELS <- paste(i/seg*100,"(%)")
            if (axistype==4|axistype==5) CAXISLABELS <- sprintf("%3.2f",i/seg)
            if (!is.null(caxislabels)&(i<length(caxislabels))) CAXISLABELS <- caxislabels[i+1]
            if (axistype==1|axistype==3|axistype==4|axistype==5) {
              if (is.null(calcex)) text(-0.05, (i+CGap)/(seg+CGap), CAXISLABELS, col=axislabcol) else
                text(-0.05, (i+CGap)/(seg+CGap), CAXISLABELS, col=axislabcol, cex=calcex)
            }
          }
          if (centerzero) {
            arrows(0, 0, xx*1, yy*1, lwd=cglwd, lty=cglty, length=0, col=cglcol)
          }
          else {
            arrows(xx/(seg+CGap), yy/(seg+CGap), xx*1, yy*1, lwd=cglwd, lty=cglty, length=0, col=cglcol)
          }
          PAXISLABELS <- df[1,1:n]
          if (!is.null(paxislabels)) PAXISLABELS <- paxislabels
          if (axistype==2|axistype==3|axistype==5) {
            if (is.null(palcex)) text(xx[1:n], yy[1:n], PAXISLABELS, col=axislabcol) else
              text(xx[1:n], yy[1:n], PAXISLABELS, col=axislabcol, cex=palcex)
          }
          VLABELS <- colnames(df)
          if (!is.null(vlabels)) VLABELS <- vlabels
          for (i in 1:n) {
            if (i == 1) {
              x_offset <- xx[i] * 0.1  # Example offset value for the top value
              y_offset <- yy[i] * 0.01
            } else if (i == 2) {
              x_offset <- xx[i] * 0.4  # Example offset value for the 2nd and 4th columns
              y_offset <- yy[i] * 0.3
            } else if (i == 4) {
              x_offset <- xx[i] * 0.4  # Example offset value for the 2nd and 4th columns
              y_offset <- yy[i] * 0.3
            } else if (i == 3 || i == 5) {
              x_offset <- xx[i] * 0.1  # Example offset value for the 3rd and 5th columns
              y_offset <- yy[i] * 0.03
            } else {
              x_offset <- 0
              y_offset <- 0
            }
            if (is.null(vlcex)) text(xx[i] * 1.2 + x_offset, yy[i] * 1.2 + y_offset, VLABELS[i]) else
              text(xx[i] * 1.2 + x_offset, yy[i] * 1.2 + y_offset, VLABELS[i], cex=vlcex, font = 2)
          }
          series <- length(df[[1]])
          SX <- series-2
          if (length(pty) < SX) { ptys <- rep(pty, SX) } else { ptys <- pty }
          if (length(pcol) < SX) { pcols <- rep(pcol, SX) } else { pcols <- pcol }
          if (length(plty) < SX) { pltys <- rep(plty, SX) } else { pltys <- plty }
          if (length(plwd) < SX) { plwds <- rep(plwd, SX) } else { plwds <- plwd }
          if (length(pdensity) < SX) { pdensities <- rep(pdensity, SX) } else { pdensities <- pdensity }
          if (length(pangle) < SX) { pangles <- rep(pangle, SX)} else { pangles <- pangle }
          if (length(pfcol) < SX) { pfcols <- rep(pfcol, SX) } else { pfcols <- pfcol }
          for (i in 3:series) {
            xxs <- xx
            yys <- yy
            scale <- CGap/(seg+CGap)+(df[i,]-df[2,])/(df[1,]-df[2,])*seg/(seg+CGap)
            if (sum(!is.na(df[i,]))<3) { cat(sprintf("[DATA NOT ENOUGH] at %d\n%g\n",i,df[i,])) # for too many NA's (1.2.2012)
            } else {
              for (j in 1:n) {
                if (is.na(df[i, j])) { # how to treat NA
                  if (na.itp) { # treat NA using interpolation
                    left <- ifelse(j>1, j-1, n)
                    while (is.na(df[i, left])) {
                      left <- ifelse(left>1, left-1, n)
                    }
                    right <- ifelse(j<n, j+1, 1)
                    while (is.na(df[i, right])) {
                      right <- ifelse(right<n, right+1, 1)
                    }
                    xxleft <- xx[left]*CGap/(seg+CGap)+xx[left]*(df[i,left]-df[2,left])/(df[1,left]-df[2,left])*seg/(seg+CGap)
                    yyleft <- yy[left]*CGap/(seg+CGap)+yy[left]*(df[i,left]-df[2,left])/(df[1,left]-df[2,left])*seg/(seg+CGap)
                    xxright <- xx[right]*CGap/(seg+CGap)+xx[right]*(df[i,right]-df[2,right])/(df[1,right]-df[2,right])*seg/(seg+CGap)
                    yyright <- yy[right]*CGap/(seg+CGap)+yy[right]*(df[i,right]-df[2,right])/(df[1,right]-df[2,right])*seg/(seg+CGap)
                    if (xxleft > xxright) {
                      xxtmp <- xxleft; yytmp <- yyleft;
                      xxleft <- xxright; yyleft <- yyright;
                      xxright <- xxtmp; yyright <- yytmp;
                    }
                    xxs[j] <- xx[j]*(yyleft*xxright-yyright*xxleft)/(yy[j]*(xxright-xxleft)-xx[j]*(yyright-yyleft))
                    yys[j] <- (yy[j]/xx[j])*xxs[j]
                  } else { # treat NA as zero (origin)
                    xxs[j] <- 0
                    yys[j] <- 0
                  }
                }
                else {
                  xxs[j] <- xx[j]*CGap/(seg+CGap)+xx[j]*(df[i, j]-df[2, j])/(df[1, j]-df[2, j])*seg/(seg+CGap)
                  yys[j] <- yy[j]*CGap/(seg+CGap)+yy[j]*(df[i, j]-df[2, j])/(df[1, j]-df[2, j])*seg/(seg+CGap)
                }
              }
              if (is.null(pdensities)) {
                polygon(xxs, yys, lty=pltys[i-2], lwd=plwds[i-2], border=pcols[i-2], col=pfcols[i-2])
              } else {
                polygon(xxs, yys, lty=pltys[i-2], lwd=plwds[i-2], border=pcols[i-2], 
                        density=pdensities[i-2], angle=pangles[i-2], col=pfcols[i-2])
              }
              points(xx*scale, yy*scale, pch=ptys[i-2], col=pcols[i-2])
              
              ## Line added to add textvalues to points
              for (j in 1:n) {
                if (j == 1) {
                  if (df[i, j] >= 100) {
                    color <- "green4"
                  } else if (df[i, j] >= 90) {
                    color <- "orange"
                  } else {
                    color <- "red"
                  }
                } else if (j == 2) {
                  if (df[i, j] >= 100) {
                    color <- "green4"
                  } else if (df[i, j] >= 90) {
                    color <- "orange"
                  } else {
                    color <- "red"
                  }
                } else if (j == 3) {
                  if (df[i, j] >= 100) {
                    color <- "green4"
                  } else if (df[i, j] >= 90) {
                    color <- "orange"
                  } else {
                    color <- "red"
                  }
                } else if (j == 4) {
                  if (df[i, j] >= 105) {
                    color <- "green4"
                  } else if (df[i, j] >= 95) {
                    color <- "orange"
                  } else {
                    color <- "red"
                  }
                } else if (j == 5) {
                  if (df[i, j] >= 104) {
                    color <- "green4"
                  } else if (df[i, j] >= 95) {
                    color <- "orange"
                  } else {
                    color <- "red"
                  }
                } else {
                  color <- "black"  # Default color for any additional values
                }
                text(xx[j]*scale[j]*0.8, yy[j]*scale[j]*0.8, df[i, j], cex = 2, font = 2,col = color)
              }
            }
          }
        }
        
        
        if (!is.na(sj) | dotaz_spiro == "YES")  {
          png(paste0(program_slozka,"/radar.png"), width = 1124, height = 797)
          
          
          # Create radar chart
          
          radarchart2(
            data_radar,
            axistype = 1,
            pcol = rgb(0.2, 0.5, 0.5, 0.9),
            pfcol = NA, 
            cglcol = "grey",
            cglty = 1,
            axislabcol = "grey",
            caxislabels = seq(0, 100, 25),
            cglwd = 0.8,
            vlcex = 1,
            title = "Rozložení parametrů"
          )
          
          dev.off()
          
          
          
          
        }
        if (!is.na(sj) | dotaz_spiro == "YES") {
          Sys.setenv(RSTUDIO_PANDOC = "C:/Program Files/RStudio/resources/app/bin/quarto/bin/tools")
          rmarkdown::render(paste(program_slozka, "/export_NEOTVIRAT.Rmd", sep = ""), params = list(dotaz_spiro = dotaz_spiro), output_file = paste(id, ".pdf", sep= ""), clean = T, quiet = F)
        } else {
          Sys.setenv(RSTUDIO_PANDOC = "C:/Program Files/RStudio/resources/app/bin/quarto/bin/tools")
          rmarkdown::render(paste(program_slozka, "/export_NEOTVIRAT_wingate.Rmd", sep = ""), params = list(dotaz_spiro = dotaz_spiro), output_file = paste(id, ".pdf", sep= ""), clean = T, quiet = F)
        }
        
        
        
        fs::file_move(path = paste(program_slozka, "/", id, ".pdf", sep = ""), new_path = paste(wd, "/reporty/", id, ".pdf", sep = ""))
        # fs::file_move(path = paste(program_slozka, "/", id, ".tex", sep = ""), new_path = paste(wd, "/vymazat/", id, ".tex", sep = ""))
        # file.rename(from = paste(program_slozka, "/", id, "_files", sep = ""), to = paste(wd, "/vymazat/", id, "_files", sep = ""))
        
        
        
        df2$Name <- fullname.rev
        df2 <- df2 %>% relocate(Name, .before = Date_meas.)
        df2$Age <- ifelse(!is.na(age), floor(age), 
                          ifelse(!is.na(s.age), floor(s.age), NA))
        df2 <- df2 %>% relocate(Age, .after = Date_meas.)
        df2$Turns <- round(tail(df$Turns.number..Nr.,1),0)
        
        if (srovnani == "YES") {
          if (id %in% file.list.compar.bez) {
            df2$Turns[length(df2$Turns)-1] <- round(tail(compare.wingate$Turns.number..Nr.,1),0) 
          }
        }
        
        df2 <- df2 %>% relocate(Turns, .after = ATH)
        df2$Sport <- sport
        df2 <- df2 %>% relocate(Sport, .after = Date_meas.)
        df2$Team <- team
        df2 <- df2 %>% relocate(Team, .after = Sport)
        df2$`HR_max (BPM)` <- as.numeric(df2$`HR_max (BPM)`)
        df2$`Pmax/kgATH (W/kg)` <- df2$`Pmax (W)`/df2$ATH
        df2$id <- id
        df2 <- df2 %>% relocate(id, .before = Name)
        df2$`Anc/kg (J/kg)` <- (df2$`Work (kJ)`*1000)/df2$Weight
        
        
        
        if(!exists("databaze")) {
          databaze <- data.frame()
          databaze <- df2
        } else {
          databaze <- rbind(databaze, df2)
        }
        
        dotaz_spiro <- dotaz_spiro_original
      }
      
      # end of MAIN for loop
      
      if (!dir.exists(paste(wd, "/vysledky/spiro/", sep=""))) {
        dir.create(paste(wd, "/vysledky/spiro/", sep=""))
      }
      
      #### Report printing ####
      if (exists("df4")) {
        if (dotaz_spiro == "YES" && length(file.list.spiro) > 0) {
          df4_a(df4)
          df4 <- df4[rev(order(df4$`VO2max (ml/kg/min)`)), ]
          df4 <- df4[rowSums(is.na(df4)) != ncol(df4), ]
          if (!(input$spirometrie && input$wingate && length(spiro_ids) > 0)) {
            writexl::write_xlsx(df4, paste(wd, "/vysledky/spiro/spiro_vysledek", "_", team, "_", format(Sys.time(), "_%d-%m %H-%M"), ".xlsx", sep = ""), col_names = T, format_headers = T)
          }
        }
      }
      
      writexl::write_xlsx(databaze, paste(wd, "/vysledky/vysledek", "_", team, "_", format(Sys.time(), "_%d-%m %H-%M"), ".xlsx", sep = ""), col_names = T, format_headers = T)
      
      i2_ <- i
      i2(i2_)
      
      if (exists("chybne.pocty")) {
        if (srovnani == "YES" & length(chybne.pocty) > 0) {
          print("Srovnání nevyhodnoceno (chybí antropa):")
          print(chybne.pocty)
        }
      }
      
      # input <- readline("Nahrat vysledky do celkove databaze? (A/N): ")
      # if(input == "A") {
      #   zmeny.cas <- c()
      #   for (i in 1:length(file.list.dat)) {
      #     zmeny.cas <- append(zmeny.cas, file.info(paste(database.path, "/", file.list.dat[i], sep=""))$mtime)
      #   }
      #   
      #   
      #   datab_old <- readxl::read_excel(paste(database.path, "/", file.list.dat[which.min(abs(zmeny.cas-Sys.time()))], sep=""))
      #   datab_new <- rbind(datab_old, databaze)
      #   
      #   writexl::write_xlsx(datab_new, paste(database.path, "/databaze", "_", format(Sys.time(), "_%d-%m %H-%M"), ".xlsx", sep = ""), col_names = T, format_headers = T)  
      # }
      
    } 
    
    
    
    process_and_generate_spiro <- function(antropometrie_path, spirometrie_path, N, incProgress, units_switch, program_slozka) {
      
      if (!dir.exists(paste(wd, "/vysledky", sep=""))) {
        dir.create(paste(wd, "/vysledky", sep=""))
      }
      
      if (!dir.exists(paste(wd, "/vymazat", sep=""))) {
        dir.create(paste(wd, "/vymazat", sep=""))
      }
      
      if (!dir.exists(paste(wd, "/reporty", sep=""))) {
        dir.create(paste(wd, "/reporty", sep=""))
      }
      
      
      
      library(finalfit)
      library(dplyr)
      library(ggplot2)
      library(lubridate)
      library(gt)
      library(cowplot)
      library(tidyverse)
      library(tinytex)
      library(rapportools)
      library(scales)
      library(webshot)
      
      if(!is.null(df4_a())) {
        df4 <- df4_a()
      } else {
        df4 <- data.frame(
          `Date meas.` = character(),
          Name = character(),
          Weight = numeric(),
          Fat = numeric(),
          `VO2max (l)` = numeric(),
          `VO2max (ml/kg/min)` = numeric(),
          `Výkon (W)` = numeric(),
          `Rychlost (km/h)` = numeric(),
          `Výkon (l/kg)` = numeric(),
          `HRmax (BPM)` = numeric(),
          `ANP (BPM)` = numeric(),
          `Tep. kyslík (ml)` = numeric(),
          `VT (l)` = numeric(),
          RER = numeric(),
          `LaMax (mmol/l)` = numeric(),
          `FEV1 (l)` = numeric(),
          `FVC (l)` = numeric(),
          `Aerobní Z. do` = numeric(),
          `Smíšená Z. od` = numeric(),
          `Smíšená Z. do` = numeric(),
          `Anaerobní Z. od` = numeric(),
          stringsAsFactors = FALSE,
          check.names = FALSE)
        df4 <- df4[nrow(df4) + 1,]
        
      }
      
      wd <-  dirname(as.character(antropometrie_path))
      
      
      for(i in 1:length(file.list.spiro)) {
        id <- tools::file_path_sans_ext(file.list.spiro[i])
        if (!is.null(i2())) {
          i2 <- i2()
          incProgress(1/N, detail = paste("Processing file", i+i2, "of", N, ": ", id))
        } else {
          incProgress(1/N, detail = paste("Processing file", i, "of", N, ": ", id))
        }
        
        print(paste("Tisknu", i,"/", length(file.list.spiro), ": ", id))
        if (an.input == "solo") {
          file.an <- file.list.an[which(file.list.an.bez == id)]
          antropo <- readxl::read_excel(paste(file.path.an, "/", file.an, sep=""), sheet = "Data_Sheet")
        }
        
        if (an.input == "batch") {
          antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet") 
        }
        
        if (an.input == "batch") {
          datum_mer <- na.omit(antropo$Date_measurement[antropo$ID == id])
          vaha <- na.omit(antropo$Weight[antropo$ID == id])
          vaha <- ifelse(length(vaha) == 0, NA, vaha)
          vyska <- na.omit(antropo$Height[antropo$ID == id])
          vyska <- ifelse(length(vyska) == 0, NA, vyska)
          fat <- round(na.omit(antropo$Fat[antropo$ID == id]),1)
          fat <- ifelse(length(fat) == 0, NA, fat)
          ath <- round(na.omit(antropo$ATH[antropo$ID == id]),1)
          ath <- ifelse(length(ath) == 0, NA, ath)
          birth <- na.omit(antropo$Birth[antropo$ID == id])
          fullname <- paste(na.omit(antropo$Name[antropo$ID == id]), na.omit(antropo$Surname[antropo$ID == id]), sep = " ", collapse = NULL)
          fullname.rev <- paste(na.omit(antropo$Surname[antropo$ID == id]), na.omit(antropo$Name[antropo$ID == id]), sep = " ", collapse = NULL)
          datum_nar <- head(format(as.Date(na.omit(antropo$Birth[antropo$ID == id])),"%d/%m/%Y"),1)
          datum_mer <- format(as.Date(na.omit(antropo$Date_measurement[antropo$ID == id])),"%d/%m/%Y")
          la <- antropo$LA[antropo$ID == id][1]
          age <- ifelse(!is.empty(antropo$Age[antropo$ID == id]), 
                        round(na.omit(antropo$Age[antropo$ID == id]),2), 
                        ifelse(!is.na(datum_nar) & datum_nar != "",
                               floor(as.integer(difftime(as.Date(datum_mer, format = "%d/%m/%Y"), as.Date(datum_nar, format = "%d/%m/%Y"), units = "days") / 365.25)),
                               NA))
          
        } else {
          datum_mer <- na.omit(tail(antropo$Date_measurement, n=1))
          vaha <- na.omit(tail(antropo$Weight, n = 1))
          vaha <- ifelse(length(vaha) == 0, NA, vaha)
          vyska <- na.omit(tail(antropo$Height, n = 1))
          vyska <- ifelse(length(vyska) == 0, NA, vyska)
          fat <- round(na.omit(tail(antropo$Fat, n =1)),1)
          fat <- ifelse(length(fat) == 0, NA, fat)
          ath <- round(na.omit(tail(antropo$ATH, n =1)),1)
          ath <- ifelse(length(ath) == 0, NA, ath)
          birth <- na.omit(tail(antropo$Birth, n = 1))
          fullname <- paste(na.omit(tail(antropo$Name, n =1)), na.omit(tail(antropo$Surname, n = 1)), sep = " ", collapse = NULL)
          fullname.rev <- paste(tail(antropo$Surname, n = 1), tail(antropo$Name, n = 1), sep = " ", collapse = NULL)
          datum_nar <- format(as.Date(na.omit(tail(antropo$Birth, n =1))),"%d/%m/%Y")
          datum_mer <- format(as.Date(na.omit(tail(antropo$Date_measurement, n =1))),"%d/%m/%Y")
          la <- tail(antropo$LA, n=1)
          age <- ifelse(!is.empty(antropo$Age[antropo$ID == id]), 
                        round(na.omit(tail(antropo$Age, n =1)),2), 
                        ifelse(!is.na(datum_nar) & datum_nar != "",
                               floor(as.integer(difftime(as.Date(datum_mer, format = "%d/%m/%Y"), as.Date(datum_nar, format = "%d/%m/%Y"), units = "days") / 365.25)),
                               NA))
        }
        if (length(file.list.spiro) != 0) {
          spiro <- readxl::read_excel(paste(spiro.path, "/", file.list.spiro[i], sep=""))
          spiro <- spiro[rowSums(!is.na(spiro)) > 0,]
          rows_with_bf <- which(spiro[[1]] == "BF")
          spiro_info <- spiro[1:rows_with_bf, ]
          
          spiro <- spiro[(rows_with_bf + 1):nrow(spiro), ]
          colnames(spiro) <- as.character(spiro[1, ])
          spiro <- spiro[-c(1,2), ]
          spiro <- subset(spiro, Fáze == "Zátěž")
          spiro[4:ncol(spiro)] <- lapply(spiro[4:ncol(spiro)], as.numeric)
          spiro$t <- gsub(",", ".", spiro$t)
          spiro$t <- as.POSIXct(spiro$t, format = "%H:%M:%OS", tz = "UTC")
          spiro$t <- spiro$t - min(spiro$t)
          time_intervals <- diff(spiro$t)
          median_interval <- as.numeric(median(time_intervals))
          closest_rows <- max(1, round(5 / median_interval))
          spiro$VT_5s <- zoo::rollmean(spiro$VT, k = closest_rows, align = "right", fill = NA)
          
        }
        
        if (units_switch) {
          speed <- as.numeric(gsub(",", ".", spiro_info[which(spiro_info[, 1] == "v"), 15]))
        }
        s.datum.mer <- format(as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Počáteční čas")], format = "%d.%m.%Y %H:%M"), "%d/%m/%Y")
        s.datum.nar <- format(as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Datum narození")], format = "%d.%m.%Y"),"%d.%m.%Y")
        s.age <- floor(as.integer(difftime(as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Počáteční čas")], format = "%d.%m.%Y %H:%M"), as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Datum narození")], format = "%d.%m.%Y")), units = "days") / 365.25)
        s.height <- as.numeric(gsub("cm", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "Výška")][1])))  
        FVC <- as.numeric(gsub("L", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "VC")][1]))) 
        FEV1 <- as.numeric(gsub("L", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "FEV1")][1])))
        s.pomer <- round(FEV1/FVC*100,1)
        VO2max <- round(max(spiro$`V'O2`), 1)
        VO2_kg <- VO2_kg <- as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "V'O2/kg"), 12]))
        vykon <- suppressWarnings(as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "WR"), 12])))
        if (is.na(vykon)) {
          vykon <- suppressWarnings(as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "WR"), 15])))
        }
        if (is.na(vykon)) {
          vykon <- as.numeric(round(max(spiro$WR, na.rm = TRUE), 0))
        }
        vykon_kg <- as.numeric(vykon/vaha)
        vykon_kg <- as.numeric(vykon/vaha)
        hrmax <- round(max(spiro$TF),0)
        tep.kyslik <- as.numeric(spiro_info[which(spiro_info[,1] == "V'O2/HR"), 12])
        radky_spiro_info <- as.character(unlist(spiro_info[,1]))
        
        if ("TF" %in% radky_spiro_info) {
          index <- which(spiro_info[,1] == "TF")
        } else if ("SF" %in% radky_spiro_info) {
          index <- which(spiro_info[,1] == "SF")
          colnames(spiro)[colnames(spiro) == "SF"] <- "TF"
        } else {
          stop("Nebyl nalezen TF ani SF.")
        }
        anp <- as.numeric(ifelse(spiro_info[index, 9] != "-", 
                                 as.numeric(spiro_info[index, 9]), 
                                 as.numeric(spiro_info[index, 15]) * 0.85))
        
        min.ventilace <- as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "V'E"), 12]))
        dech.frek <- as.numeric(spiro_info[which(spiro_info[,1] == "BF"), 12])
        dech.objem <-   dech.objem <- round(max(spiro$VT_5s[which.max(as.numeric(format(spiro$t, "%H%M%S")) > 100):nrow(spiro)]),1)
        dech.objem.per <- round((dech.objem/FVC)*100,0)
        rer <- round(max(spiro$RER),2)
        LaMax <- la
        
        
        #spiro tabulka
        columns_s <- c("Antropometrie", "Hodnota", "V3", "Spiroergometrie", "Absolutní", "Relativní (TH)", "% z nál. hodnoty" ,"Relativní (ATH)")
        df3 = data.frame(matrix(nrow = 14, ncol = length(columns_s))) 
        colnames(df3) <- columns_s
        df3$Antropometrie <-  `length<-`(c("Datum narození", "Věk", "Výška", "TH - Hmotnost", "Tuk", "ATH - Aktivní tělesná hmota", NA, "Klidová ventilace", "FVC", "FEV1", "Poměr FVC/FEV1"), nrow(df3))
        df3$Spiroergometrie <- `length<-`(c("VO2max", "Dosažený výkon", "HrMax", "Tepový kyslík", "ANP", NA, NA, "Ventilace", "Minutová ventilace (l/min)", "Dechová frekvence", "Dechový objem", "Dechový objem %", "RER", "LaMax"),nrow(df3))
        df3$V3[8] <- "% z nál. hodnoty"
        df3$`Relativní (TH)`[8] <- "Optimum"
        
        df3[1,2] <- datum_nar
        df3[2,2] <- age
        df3[3,2] <-  paste(s.height, "cm", sep=" ")
        df3[4,2] <- paste(vaha, "kg", sep=" ")
        df3[5,2] <- paste(fat, "%", sep=" ")
        df3[6,2] <- paste(ath, "kg", sep=" ")
        df3[9,2] <- paste(FVC,"l", sep=" ")
        df3[10,2] <- paste(FEV1,"l", sep=" ")
        df3[11,2] <- paste(s.pomer,"%", sep=" ")
        df3[9,3] <- round((FVC/sports_data$FVC[sports_data$Sport==sport])*100,1)
        df3[10,3] <- round((FEV1/sports_data$FEV1[sports_data$Sport==sport])*100,1)
        df3[1,7] <- round((VO2_kg/sports_data$VO2max[sports_data$Sport==sport])*100,1)
        if (units_switch) {
          df3[2,7] <- "-"
        } else {
          df3[2,7] <- round(((vykon/vaha)/sports_data$Power[sports_data$Sport==sport])*100,1)
        }
        df3[1,5] <- paste(VO2max,"l", sep=" ")
        if (units_switch) {
          df3[2,5] <- paste(speed,"km/h", sep=" ")
        } else {
          df3[2,5] <- paste(vykon,"W", sep=" ")
        }
        
        df3[3,5] <- paste(hrmax,"BPM", sep=" ")
        df3[4,5] <- paste(tep.kyslik,"ml/tep", sep=" ")
        df3[5,5] <- paste(round(anp,0),"BPM", sep=" ")
        df3[9,5] <- paste(min.ventilace,"l/min", sep=" ")
        df3[10,5] <- paste(dech.frek,"d/min", sep=" ")
        df3[11,5] <- paste(dech.objem,"l", sep=" ")
        df3[12,5] <- paste(dech.objem.per,"%", sep=" ")
        df3[13,5] <- rer
        df3[14,5] <- ifelse(is.empty(la), NA, paste(la, "mmol/l", sep = " "))
        df3[1,6] <- paste(VO2_kg, "ml/min/kg", sep=" ")
        if (units_switch) {
          df3[2,6] <- "-"
        } else {
          df3[2,6] <- paste(round(vykon/vaha, 1), "W/kg", sep=" ")
        }
        df3[1,8] <- paste(round((VO2max*1000)/ath,1), "ml/min/kg", sep=" ")
        if (units_switch) {
          df3[2,8] <- "-"
        } else {
          df3[2,8] <- paste(round(vykon/ath, 1), "W/kg", sep=" ")
        }
        df3[9,6] <- round(30*FVC,0)
        df3[10,6] <- "50-60"
        df3[12,6] <- "50-60"
        df3[13,6] <- "1.08-1.18"
        df3 <- df3[-7,]
        
        
        #treninkove zony
        columns_tz <- c("Zóna", "od (BPM)", "do (BPM)")
        tz = data.frame(matrix(nrow = 3, ncol = length(columns_tz))) 
        colnames(tz) <- columns_tz
        tz$Zóna <- c("Aerobní", "Smíšená", "Anaerobní")
        tz$`od (BPM)`[2] <- round((anp / 0.85)*0.76,0)
        tz$`od (BPM)`[3] <- round(anp,0)
        tz$`do (BPM)`[1] <- round((anp / 0.85)*0.75,0)
        tz$`do (BPM)`[2] <- round((anp / 0.85)*0.84,0)
        
        tf_range <- min(spiro$`V'E`)
        ve_range <- max(spiro$TF)
        min(spiro$TF)
        
        if (!exists("df4")) {
          df4 <- data.frame()
          df4 <- df4[nrow(df4) + 1,]
        } 
        
        df4$`Date meas.` <- append(na.omit(df4$`Date meas.`), datum_mer)
        df4$Name <- append(na.omit(df4$Name), fullname)
        df4$Weight <- append(na.omit(df4$Weight), vaha)
        df4$Fat <- append(na.omit(df4$Fat), fat)
        df4$`VO2max (l)` <- append(na.omit(df4$`VO2max (l)`), VO2max)
        df4$`VO2max (ml/kg/min)` <- append(na.omit(df4$`VO2max (ml/kg/min)`), VO2_kg)
        df4$`Výkon (W)` <- append(na.omit(df4$`Výkon (W)`), vykon)
        if (units_switch) {
          df4$`Rychlost (km/h)` <- append(na.omit(df4$`Rychlost (km/h)`), speed)
        }
        df4$`Výkon (l/kg)` <- append(na.omit( df4$`Výkon (l/kg)`), round(vykon_kg,1))
        df4$`HRmax (BPM)` <- append(na.omit(df4$`HRmax (BPM)`), hrmax)
        df4$`ANP (BPM)` <- append(na.omit(df4$`ANP (BPM)`), round(anp,0))
        df4$`Tep. kyslík (ml)` <- append(df4$`Tep. kyslík (ml)`[1:length(df4$`Tep. kyslík (ml)`)-1], tep.kyslik)
        df4$`VT (l)` <- append(na.omit(df4$`VT (l)`), dech.objem)
        df4$RER <- append(na.omit(df4$RER), rer)
        df4$`LaMax (mmol/l)` <- append(na.omit(df4$`LaMax (mmol/l)`), la)
        df4$`FEV1 (l)` <- append(df4$`FEV1 (l)`[1:length(df4$`FEV1 (l)`)-1], FEV1)
        df4$`FVC (l)` <- append(na.omit(df4$`FVC (l)`), FVC)
        df4$`Aerobní Z. do` <- append(na.omit(df4$`Aerobní Z. do`), round((anp / 0.85)*0.75,0))
        df4$`Smíšená Z. od` <- append(na.omit(df4$`Smíšená Z. od` ), round((anp / 0.85)*0.76,0))
        df4$`Smíšená Z. do` <- append(na.omit(df4$`Smíšená Z. do`), round((anp / 0.85)*0.84,0))
        df4$`Anaerobní Z. od` <- append(na.omit(df4$`Anaerobní Z. od`), anp)
        df4 <- add_row(df4)
        
        
        
        coeff <- 50 
        
        hrmin <- min(spiro$TF)
        
        ylim.prim <- c(0, min.ventilace*1.1)# in this example, precipitation
        ylim.sec <- c(hrmin*0.9, hrmax*1.05)  
        
        b <- diff(ylim.prim)/diff(ylim.sec)
        a <- ylim.prim[1] - b*ylim.sec[1]
        
        # Plot with TF and VE
        plot1 <- ggplot(spiro, aes(x = t)) +
          geom_line(aes(y = a + TF*b, color = "Tepová frekvence (BPM)", linetype = "Tepová frekvence (BPM)"), size = 1) +
          geom_line(aes(y = `V'E`, color = "Minutová ventilace (l/min)", linetype = "Minutová ventilace (l/min)"), size = 0.5, alpha = 0.2, linetype = "dashed") +
          geom_smooth(aes(y = `V'E`), color = "black", linetype = "solid", size = 1, method = "loess", se = FALSE, span = 0.2) +
          scale_y_continuous(
            name = "Minutová ventilace (l/min)", 
            labels = scales::comma,
            sec.axis = sec_axis(~ (. - a)/b, name = "Tepová frekvence (BPM)")
          ) +
          scale_x_datetime(labels = scales::date_format("%M:%S"), date_breaks = "1 min") +
          theme_classic() +
          theme(
            text = element_text(size = 30),
            panel.grid.major = element_blank(),
            panel.grid.minor = element_blank(),
            axis.line = element_line(colour = "black", linewidth = 1.5),
            axis.text.x = element_text(angle = -45, hjust = 0),
            axis.title.y = element_text(size = 25),
            legend.position = "bottom"
          ) +
          scale_linetype_manual(values = c('solid', 'solid'), 
                                labels = c("Minutová ventilace (l/min)", "Tepová frekvence (BPM)")) +
          scale_color_manual(values = c("Minutová ventilace (l/min)" = "black", "Tepová frekvence (BPM)" = "orange")) +
          xlab(expression("Čas")) +
          ylab("Hodnota") +
          labs(colour="") +
          guides(
            color = guide_legend(override.aes = list(linetype = c("solid", "solid"))), 
            linetype = "none"
          )
        
        plot2 <- ggplot(spiro, aes(x = t)) +
          geom_line(aes(y = `V'O2`), color = "darkmagenta", size = 0.5, alpha = 0.2, linetype = "dashed") +
          geom_smooth(aes(y = `V'O2`), color = "darkmagenta", linetype = "solid", size = 1, method = "loess", se = FALSE, span = 0.3) +
          scale_x_datetime(labels = scales::date_format("%M:%S"), date_breaks = "1 min") +
          scale_y_continuous(
            name = "Spotřeba kyslíku (l)", 
            labels = scales::comma
          ) +
          theme_classic() +
          theme(
            text = element_text(size = 30),
            panel.grid.major = element_blank(),
            panel.grid.minor = element_blank(),
            axis.line = element_line(colour = "black", linewidth = 1.5),
            axis.text.x = element_text(angle = -45, hjust = 0),
            axis.title.y = element_text(size = 25),
            legend.position = "none"
          ) +
          xlab(expression("Čas")) +
          ylab("Spotřeba kyslíku (l)") 
        
        library(patchwork)
        combined_plot <- plot1 + plot2 +
          plot_annotation(
            title = 'Spiroergometrie',
            caption = '',
            theme = theme(plot.title = element_text(size = 30, hjust = 0.5))
          )
        
        
        table3 <- df3 %>% 
          gt() %>% 
          tab_header(
            title = fullname, 
            subtitle = paste("Datum měření: ", datum_mer, sep = "")
          ) %>% 
          cols_label(Hodnota = "") %>% 
          cols_label(V3 = "")  %>% 
          sub_missing(columns = everything(), rows = everything(), missing_text = "")  %>%
          tab_style(
            style = cell_text(weight = "bold"),
            locations = cells_column_labels()
          ) %>%
          tab_style(
            style = cell_text(weight = "bold"),
            locations = cells_body(
              columns = c("Antropometrie", "Spiroergometrie")
            )
          ) %>%
          tab_style(
            style = cell_text(weight = "bold"),
            locations = cells_body(
              rows = 7,
              columns = everything()  # Apply bold style to all columns in row 9
            )
          ) %>%
          tab_style(
            style = "padding-right:30px",
            locations = cells_column_labels()
          ) %>%
          tab_style(
            style = "padding-right:30px",
            locations = cells_body()
          ) %>% 
          tab_options(
            table.width = pct(c(100)),
            container.width = 1600,
            container.height = 650,
            table.font.size = px(18L),
            heading.padding = pct(0)
          ) %>%
          gtExtras::gt_add_divider(columns = "V3", style = "solid", color = "#808080") %>% 
          opt_table_lines(extent = "none") %>%
          tab_style(
            style = cell_borders(
              sides = "top",
              color = "#808080",
              style = "solid"
            ),
            locations = cells_body(
              rows = c(1,8)
            )
          )  %>% 
          tab_style(
            style = cell_text(align = "right"),
            locations = cells_column_labels(columns = c("Absolutní", "Relativní (TH)", "% z nál. hodnoty", "Relativní (ATH)"))
          ) %>% 
          cols_align(
            align = c("center"),
            columns = c("V3", "Absolutní", "Relativní (TH)", "% z nál. hodnoty", "Relativní (ATH)")
          ) %>% 
          tab_style(
            style = cell_text(align = "center"),
            locations = cells_body(columns = c("V3", "Absolutní", "Relativní (TH)", "% z nál. hodnoty", "Relativní (ATH)"))
          ) %>%
          tab_style(
            style = "padding-bottom:20px",
            locations = cells_body(rows = 6)
          )
        
        table4  <- tz %>% gt() %>% 
          tab_header(title = "Tréninkové zóny") %>% 
          cols_label(Zóna = "")  %>% sub_missing(columns = everything(),
                                                 rows = everything(),
                                                 missing_text = "") %>% 
          tab_options(
            table.width = pct(c(100)),
            container.width = 300,
            container.height = 600,
            table.font.size = px(18L),
            heading.padding = pct(0))
        
        gtsave(
          table4,
          paste0(program_slozka, "/t4.html") # Ensure styles are applied
        )
        
        webshot(
          url = paste0(program_slozka, "/t4.html"), # Path to the saved HTML
          file = paste0(program_slozka, "/t4.png"), # Save as PNG
          vwidth = 350,  # Set viewport width
          vheight = 300,
          zoom = 2,
          cliprect = c(0, 0, 350, 300)# Set viewport height
        )
        
        png(paste0(program_slozka, "/p2.png"), width = 1500, height = 600)
        plot(combined_plot)
        dev.off()
        
        # Save table3 as HTML, then convert to PNG
        table3 %>%
          gtsave(paste0(program_slozka, "/t3.png"), vwidth = 1700, vheight = 1000) 
      
      
  
        Sys.setenv(RSTUDIO_PANDOC = "C:/Program Files/RStudio/resources/app/bin/quarto/bin/tools")
        rmarkdown::render(paste(program_slozka, "/export_spiro_NEOTVIRAT.Rmd", sep = ""), output_file = paste(id, ".pdf", sep= ""), clean = T, quiet = F)
        
        fs::file_move(path = paste(program_slozka, "/", id, ".pdf", sep = ""), new_path = paste(wd, "/reporty/", id, ".pdf", sep = ""))
        # fs::file_move(path = paste(program_slozka, "/", id, ".tex", sep = ""), new_path = paste(wd, "/vymazat/", id, ".tex", sep = ""))
        # file.rename(from = paste(program_slozka, "/", id, "_files", sep = ""), to = paste(wd, "/vymazat/", id, "_files", sep = ""))
      }
      
      if (!dir.exists(paste(wd, "/vysledky/spiro/", sep=""))) {
        dir.create(paste(wd, "/vysledky/spiro/", sep=""))
      }
      
      
      df4 <- df4[rev(order(df4$`VO2max (ml/kg/min)`)), ]
      df4 <- df4[rowSums(is.na(df4)) != ncol(df4), ]
      writexl::write_xlsx(df4, paste(wd, "/vysledky/spiro/spiro_vysledek", "_", team, "_", format(Sys.time(), "_%d-%m %H-%M"), ".xlsx", sep = ""), col_names = T, format_headers = T)
    }
    
    if (input$spirometrie && !input$wingate && !is.null(antropometrie_path)) {
      print("spiro")
      process_and_generate_spiro(antropometrie_path, spirometrie_path, N, incProgress, units_switch, program_slozka)
    }
    
    else if (input$spirometrie && input$wingate && length(spiro_ids) > 0) {
      print("oboje")
      spiro_ids <- paste0(spiro_ids, ".xlsx")
      file.list.spiro.2 <- file.list.spiro[file.list.spiro %in% spiro_ids]
      all_files <- c(file.list.spiro.2, file.list)
      all_files <- tools::file_path_sans_ext(all_files)
      process_and_generate_pdfs(wingate_path, antropometrie_path,spirometrie_path, N, incProgress, units_switch, program_slozka)
      
      file.list.spiro<- file.list.spiro[file.list.spiro %in% spiro_ids]
      
      process_and_generate_spiro(antropometrie_path, spirometrie_path, N, incProgress, units_switch, program_slozka)
    } else {
      print("wingate")
      withProgress(message = 'Calculation in progress', {
        N <- length(file.list)
        process_and_generate_pdfs(wingate_path, antropometrie_path,spirometrie_path, N, incProgress, units_switch, program_slozka)
      })
    }
    
    
    output$result <- renderTable({
      result_val()
    })
    
  })
})
