# app.R
library(shiny)

# Source the UI and server code
source("ui.R")
source("server.R")

# Run the application
shinyApp(ui = ui, server = server, options = list(launch.browser = T, port = 8080))

