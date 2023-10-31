source("shiny_functions.R")
check_packages()

library(shiny)
library(tidyverse)
library(data.table) 
library(readxl)
library(purrr)
library(janitor)
library(knitr)
library(kableExtra)
library(stringr)
library(viridis)
library(geomtextpath)
library(ggrepel)
library(plotly)


schedules <- setNames(c("N1", "B1", "M1"), c("Non-Union", "Bargaining Unit", "Management"))

ui <- fluidPage(
  tabsetPanel(
    tabPanel("Data Import",
             fluidRow(
               sidebarPanel(
                 fileInput("upload", "Upload Reports", multiple = TRUE, placeholder = "No file uploaded"),
                 selectInput("files", label = "Select file", choices = "Upload file first"),
                 actionButton("run", "Run Revision!")
               ),
               mainPanel(
                 htmlOutput("contact")
               )
             ),
             fluidRow(
               column(12, htmlOutput("funding"))
             )),
    tabPanel("Report Revisions",
             h1("Revisions"),
             htmlOutput("home_rev"),
             htmlOutput("rev")),
    tabPanel("Report",
             sidebarLayout(
               sidebarPanel(
                 selectInput("vsched", "Select schedule", choices = "N1"),
                 checkboxGroupInput("vvar", "Variables to display:", c("All", "None")),
                 width = 2),
               mainPanel(htmlOutput("vsched"),
                         width = 10)
             )),
    tabPanel("Graphs",
             sidebarPanel(
               selectInput("plotData", "Select Plot", choices = c("Payroll vs. Hours Scatter Plot",
                                                                  "Funding Source Composition",
                                                                  "Compensation Costs by Employee Group")),
               conditionalPanel(
                 condition = "input.plotData == 'Payroll vs. Hours Scatter Plot'",
                 numericInput("min_wage", "BC Minimum Wage", value = 15.6)),
               conditionalPanel(
                 condition = "input.plotData == 'Funding Source Composition'",
                 radioButtons("funding", "Level of data aggregation", c("By Union/Non-Union Programs", "By Funder"))),
               conditionalPanel(
                 condition = "input.plotData == 'Compensation Costs by Employee Group'",
                 radioButtons("compensation", "Level of Data Aggregation", c("By PF/NPF", "By Wage Cost Drivers")),
                 conditionalPanel(
                   condition = "input.compensation == 'By Wage Cost Drivers'",
                   checkboxInput("fine", "Fine level of data aggregation?", value = F)
                 )
               )
             ),
             mainPanel(
               plotlyOutput("plot")
             ))
  )
)

server <- function(input, output, session) {
  observeEvent(input$upload, {
    updateSelectInput(session, "files", label = "Select file", choices = input$upload$name)
  })
  observeEvent(input$run, {
    updateSelectInput(session, "vsched", label = "Select Schedule", choices = names(data()))
  })
  
  data <- eventReactive(input$run, {
    clean_schedule(load_schedule(input$upload[input$upload$name == input$files,]$datapath))
  })
  output$vsched <- renderText(shiny_tableOutput(data(), "N1"))
  
  output$contact <- renderText(shiny_tableOutput(data(), "Contact"))
  output$funding <- renderText(shiny_tableOutput(data(), "Home"))
  
  sched_rev <- reactive(revisions(data()))
  output$home_rev <- renderText({
    if(sched_rev()[["Home"]][Notes != "No Issues", .N] != 0) {
      kable(sched_rev()[["Home"]][Notes != "No Issues",], 
                               caption = "<span style='font-size:200%'>Home Schedule</span>", 
                               align = "lccccc",
                               col.names = c("Funder", "Funding for Union Programs", "Funding for Non-Union Programs",
                                             "# of Union Contracts", "# of Non-Union Contracts", "Notes")) %>%
        kable_styling("striped")
    }
  })
  output$rev <- renderText(shiny_tableOutput(sched_rev()))
  
  output$plot <- renderPlotly({
    if(input$plotData == "Payroll vs. Hours Scatter Plot") {
      plot_data(data(), input$plotData, min_wage = input$min_wage)
    } else if(input$plotData == "Funding Source Composition") {
      plot_data(data(), input$plotData, funding = input$funding)
    } else if(input$plotData == "Compensation Costs by Employee Group") {
      if(input$compensation == "By Wage Cost Drivers") {
        plot_data(data(), input$plotData, compensation = input$compensation, fine = input$fine)
      } else {
        plot_data(data(), input$plotData, compensation = input$compensation)
      }
      
    }
    
  })
 
}
shinyApp(ui = ui, server = server)


