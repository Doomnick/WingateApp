library(shinyjs)
library(shinyFiles)

shinyUI(fluidPage(
  useShinyjs(),
  tags$head(tags$style(HTML("
    .shiny-directory-button { margin-bottom: 5px !important; }
    .section-box {
      border: 1px solid #ddd;
      padding: 10px;
      margin-bottom: 15px;
      border-radius: 5px;
      background-color: #f9f9f9;
    }
    .action-btn {
      margin-right: 5px;
      margin-bottom: 5px;
    }
  "))),
  titlePanel("Performance Analysis"),
  sidebarLayout(
    sidebarPanel(
      style = "padding-top: 10px;",
      
      # Automatic Folders Section
      div(class = "section-box",
          tags$h4("Automatic folders selection:", style = "margin-top: 0px; margin-bottom: 10px;"),
          shinyDirButton("main_folder", "Select Main Folder", "Please select a folder"),
          verbatimTextOutput("main_folder_path"),
          verbatimTextOutput("auto_detect_info")
      ),
      
      # Manual Folders Section
      div(class = "section-box",
          tags$h4("Manual folders selection:", style = "margin-top: 0px; margin-bottom: 10px;"),
          
          shinyDirButton("antropometrie", "Select Anthropometry Folder", "Optional override"),
          verbatimTextOutput("antropometrie_path"),
          
          fluidRow(
            column(4, checkboxInput("wingate", "Wingate", FALSE)),
            column(4, conditionalPanel(
              condition = "input.wingate == true",
              checkboxInput("srovnani", "Comparison", FALSE)
            ))
          ),
          
          conditionalPanel(
            condition = "input.wingate == true",
            shinyDirButton("wingate_path", "Select Wingate Folder", "Optional override"),
            verbatimTextOutput("wingate_path_display")
          ),
          
          conditionalPanel(
            condition = "input.srovnani == true",
            shinyDirButton("srovnani_path", "Select Comparison Folder", "Optional override"),
            verbatimTextOutput("srovnani_path_display")
          ),
          
          conditionalPanel(
            condition = "input.srovnani == true",
            fluidRow(
              column(4, checkboxInput("srovnani2", HTML("Comparison&nbsp;2"), FALSE)),
              column(8, conditionalPanel(
                condition = "input.srovnani2 == true",
                shinyDirButton("srovnani2_path", "Select Comparison 2 Folder", "Optional override"),
                verbatimTextOutput("srovnani2_path_display")
              ))
            )
          ),
          
          conditionalPanel(
            condition = "input.wingate == true && input.srovnani == true && input.srovnani2 == true",
            checkboxInput("tri_graf", HTML("Tri&nbsp;graph"), FALSE)
          ),
          
          checkboxInput("spirometrie", "Spirometrie", FALSE),
          conditionalPanel(
            condition = "input.spirometrie == true",
            radioButtons(
              inputId = "toggle_switch",
              label = "Choose Metric",
              choices = list("Power" = FALSE, "Speed" = TRUE),
              selected = FALSE
            ),
            shinyDirButton("spirometrie_path", "Select Spirometrie Folder", "Optional override"),
            verbatimTextOutput("spirometrie_path_display")
          )
      ),
      
      # Other Inputs Section
      div(class = "section-box",
          selectInput("sport", "Select Sport", 
                      choices = c("hokej-dospělí", "hokej-junioři", "hokej-dorost", "gymnastika"),
                      selected = "hokej-dospělí"),
          textInput("team", "Enter Team Name", "")
      ),
      
      # Action Buttons
      div(style = "margin-top: 10px;",
          actionButton("refresh_files", "Refresh Files", class = "btn-primary action-btn"),
          actionButton("check", "Check", class = "btn-secondary action-btn"),
          actionButton("continue", "Continue", style = "display: none;", class = "btn-success action-btn"),
          actionButton("submit", "Submit", style = "display: none;", class = "btn-danger action-btn"),
          uiOutput("clear_button_ui")
      )
    ),
    
    mainPanel(
      tabsetPanel(
        id = "main_tabs",
        tabPanel("Data",
                 verbatimTextOutput("paths"),
                 verbatimTextOutput("console_output"),
                 h4("Date Format check:"),
                 verbatimTextOutput("invalid_ids"),
                 h4("File List:"),
                 verbatimTextOutput("file_list_output"),
                 h4("File Comparison Table:"),
                 tableOutput("comparison_table")
        ), 
        tabPanel("Summary",
                 h4("Summary Overview"),
                 h5("Errors in Anthropometry : Wingate"),
                 verbatimTextOutput("spatne_wingate_ids"),
                 conditionalPanel(
                   condition = "input.wingate && !input.srovnani || input.spirometrie && !input$wingate",
                   h5("Duplicates in Anthropometry"),
                   verbatimTextOutput("duplikaty.an")
                 ),
                 conditionalPanel(
                   condition = "input.srovnani",
                   h5("Missing Wingate file for Comparison 1 (newest)"),
                   verbatimTextOutput("spatne_srovnani")
                 ),
                 conditionalPanel(
                   condition = "input.srovnani && input.srovnani2",
                   h5("Missing Wingate file for Comparison 2 (oldest)"),
                   verbatimTextOutput("spatne_srovnani2")
                 ),
                 conditionalPanel(
                   condition = "input.srovnani",
                   h5("Missing Anthropometry input for Comparison"),
                   verbatimTextOutput("spatne_srovnani.an")
                 ),
                 conditionalPanel(
                   condition = "input.spirometrie",
                   h5("Errors in Spirometry"),
                   verbatimTextOutput("spatne.spiro")
                 ),
                 tableOutput("result")
        )
      )
    )
  )
))