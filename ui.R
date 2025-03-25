library(shiny)
library(shinyFiles)
library(readxl)
library(fs)
library(shinyjs)
library(shinyWidgets)

shinyUI(fluidPage(
  useShinyjs(),
  titlePanel("Performance Analysis"),
  sidebarLayout(
    sidebarPanel(
      # Path to Antropometrie folder
      shinyDirButton("antropometrie", "Select Anthropometry Folder", "Please select a folder"),
      verbatimTextOutput("antropometrie_path"),
      
      # Wingate checkbox and folder path input
      fluidRow(
        column(4, checkboxInput("wingate", "Wingate", FALSE)),
        column(4, conditionalPanel(
          condition = "input.wingate == true",
          checkboxInput("srovnani", "Comparison", FALSE)
        ))
      ),
      
      
      # Comparison 2 checkbox on the next row
      fluidRow(
        column(4, conditionalPanel(
          condition = "input.wingate == true && input.srovnani == true",
          checkboxInput("srovnani2", HTML("Comparison&nbsp;2"), FALSE)
        )),
        column(4, offset = 3, conditionalPanel(
          condition = "input.wingate == true && input.srovnani == true && input.srovnani2 == true",
          checkboxInput("tri_graf", HTML("Tri&nbsp;graph"), FALSE)
        ))
      ),
      
      conditionalPanel(
        condition = "input.wingate == true",
        shinyDirButton("wingate_path", "Select Wingate Folder", "Please select a folder"),
        verbatimTextOutput("wingate_path_display")
      ),
      
      # Comparison folder path input
      conditionalPanel(
        condition = "input.srovnani == true",
        shinyDirButton("srovnani_path", "Select Comparison Folder", "Please select a folder"),
        verbatimTextOutput("srovnani_path_display")
      ),
      
      # Comparison 2 folder path input
      conditionalPanel(
        condition = "input.srovnani2 == true",
        shinyDirButton("srovnani2_path", "Select Comparison 2 Folder", "Please select a folder"),
        verbatimTextOutput("srovnani2_path_display")
      ),
      
      # Spirometrie checkbox and folder path input
      checkboxInput("spirometrie", "Spirometrie", FALSE),
      conditionalPanel(
        condition = "input.spirometrie == true",
        radioButtons(
          inputId = "toggle_switch",    # Same inputId as the switchInput
          label = "Choose Metric",       # Label for the radio button group
          choices = list(
            "Power" = FALSE,  
            "Speed" = TRUE           
          ),
          selected = FALSE               # Default selected value (Power)
        ),
        shinyDirButton("spirometrie_path", "Select Spirometrie Folder", "Please select a folder"),
        verbatimTextOutput("spirometrie_path_display")
     
      ),
      
      selectInput("sport", "Select Sport", 
                  choices = c("hokej-dospělí", "hokej-junioři", "hokej-dorost", "gymnastika"),
                  selected = "hokej-dospělí"),
      
      textInput("team", "Enter Team Name", ""),
      
      # Submit button
      
      actionButton("refresh_files", "Refresh Files"),
   
      actionButton("check", "Check"),
      actionButton("continue", "Continue", style = "display: none;"),
      actionButton("submit", "Submit", style = "display: none;"),
      uiOutput("clear_button_ui"),
    ),
    
    
    
    mainPanel(
        tabsetPanel(
          id = "main_tabs",
          tabPanel("Data",
      verbatimTextOutput("paths"),
      verbatimTextOutput("console_output"),
      h4("Date Format check:"),
      verbatimTextOutput("invalid_ids"),
      # File list output
      h4("File List:"),
      verbatimTextOutput("file_list_output"),
      #comparison
      h4("File Comparison Table:"),
      tableOutput("comparison_table")
    ), 
    tabPanel("Summary",
                # Content for the Summary panel
      h4("Summary Overview"),
      h5("Errors in Anthropometry : Wingate"),
      verbatimTextOutput("spatne_wingate_ids"),
      conditionalPanel(
        condition = "input.wingate && !input.srovnani || input.spirometrie && !input.wingate",  # JavaScript condition to match input values
        h5("Duplicates in Anthropometry"),
        verbatimTextOutput("duplikaty.an")
      ),
      conditionalPanel(
        condition = "input.srovnani",  # JavaScript condition to match input values
        h5("Missing Wingate file for Comparison 1 (newest)"),
        verbatimTextOutput("spatne_srovnani")
      ),
      conditionalPanel(
        condition = "input.srovnani && input.srovnani2",  # JavaScript condition to match input values
        h5("Missing Wingate file for Comparison 2 (oldest)"),
        verbatimTextOutput("spatne_srovnani2")
      ),
      conditionalPanel(
        condition = "input.srovnani",  # JavaScript condition to match input values
        h5("Missing Anthropometry input for Comparison"),
        verbatimTextOutput("spatne_srovnani.an")
      ),
      conditionalPanel(
        condition = "input.spirometrie",  # JavaScript condition to match input values
        h5("Errors in Spirometry"),
        verbatimTextOutput("spatne_spiro")
      ),
      tableOutput("result")
    )
  )
    
))
))