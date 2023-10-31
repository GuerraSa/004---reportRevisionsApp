source("Functions.R")

shiny_tableOutput <- function(file, schedule = NULL) {
  temp <- copy(file)
  if(!is.null(schedule)) {
    if(schedule == "Contact"){
      if(as.numeric(str_extract(file[[schedule]][[8,2]], "(?<=\\().+(?=%)")) >= 100) { # Total Funding Utilization > 100%
        kable(file[[schedule]], col.names = NULL,
              caption = "<span style='font-size:200%'>Agency Information</span>") %>%
          kable_minimal(full_width = F) %>%
          row_spec(8, bold = TRUE, color = "white", background = "red") %>%
          column_spec(1, bold = TRUE, width = "18em", color = "black", background = "white") 
        
      } else {
        kable(file[[schedule]], escape = F, col.names = NULL,
              caption = "<span style='font-size:200%'>Agency Information</span>") %>%
          kable_minimal() %>%
          column_spec(1, bold = TRUE, width = "18em") 
      }
      
    } else if(schedule == "Home" & temp[[schedule]][,.N] != 0) {
      convert_and_format(temp[[schedule]], "$", names(temp[[schedule]])[3:5])
      convert_and_format(temp[[schedule]], "%", names(temp[[schedule]])[6:8])
      cols <- names(temp[[schedule]])[sapply(temp[[schedule]], is.numeric)]
      temp[[schedule]][, (cols) := lapply(.SD, function(x) ifelse(is.na(x), "", as.character(x))), .SDcols = cols]
      
      temp[[schedule]] %>%
        select(-contains(c("Pct", "Fed_Prov"))) %>%
        kable(col.names = c("", "Funding for Union Programs", "Funding for Non-Union Programs", 
                            "Total Funding", "# of Union Contracts", "# of Non-Union Contracts", "Total Contracts"),
              align = "lcccccc",
              caption = "<span style='font-size:200%'>Funding Overview</span>") %>%
        pack_rows(index = table(fct_inorder(file[["Home"]]$Fed_Prov))) %>%
        kable_styling("striped")
    } else if(schedule == "N1") {
      cols <- names(temp[[schedule]])[sapply(temp[[schedule]], is.numeric)]
      temp[[schedule]][, (cols) := lapply(.SD, function(x) ifelse(is.na(x), "", as.character(x))), .SDcols = cols]
      temp[[schedule]] %>% 
        kable()
    }
  } 
  else {
    file[-1] %>%
      bind_rows() %>%
      mutate(Notes = transform_to_html(Notes)) %>%
      select(-Schedule) %>%
      kable("html",
            align = "lr",
            valign = "c",
            escape = FALSE) %>%
      pack_rows(index = table(fct_inorder(bind_rows(file[-1])$Schedule))) %>%
      kable_styling("striped", full_width = F) %>% 
      column_spec(1, "30em") 
  }
}

# Replace every pair of "**" with "<b>" and "</b>"
transform_to_html <- function(text) {
  text <- gsub("\\*\\*(.*?)\\*\\*", "<b>\\1</b>", text)
  return(text)
}

plot_data <- function(report, plotData, min_wage = NULL, funding = NULL, compensation = NULL, fine = NULL) {
  dt <- list(N1 = setnafill(copy(report$N1), fill = 0, cols = names(report$N1)[unlist(lapply(report$N1, is.numeric))]),
             B1 = setnafill(copy(report$B1), fill = 0, cols = names(report$B1)[unlist(lapply(report$B1, is.numeric))]),
             M1 = setnafill(copy(report$M1), fill = 0, cols = names(report$M1)[unlist(lapply(report$M1, is.numeric))]),
             S2 = setnafill(copy(report$S2), fill = 0, cols = names(report$S2)[unlist(lapply(report$S2, is.numeric))]))
  if(plotData == "Payroll vs. Hours Scatter Plot") {
    dt <- list(N1 = dt$N1[, .(row_id, Classification,
                              `Total Payroll Amount` = (NPF_HrlyWage*NPF_Hrs_StraightTime)+(PF_HrlyWage*PF_Hrs_StraightTime),
                              `Total Straight Time Hours` = NPF_Hrs_StraightTime + PF_Hrs_StraightTime,
                              `Employee Equivalent` = NPF_Active + PF_Active + Terminated,
                              `Employee Group` = "Non-Union")],
               B1 = dt$B1[, .(row_id, Classification,
                              `Total Payroll Amount` = (NPF_HrlyWage*NPF_Hrs_StraightTime)+(PF_HrlyWage*PF_Hrs_StraightTime),
                              `Total Straight Time Hours` = NPF_Hrs_StraightTime + PF_Hrs_StraightTime,
                              `Employee Equivalent` = NPF_Active + PF_Active + Terminated,
                              `Employee Group` = "Bargaining Unit")],
               M1 = dt$M1[, .(row_id, Classification,
                              `Total Payroll Amount` = NPF_Payroll + PF_Payroll,
                              `Total Straight Time Hours` = NPF_Hrs + PF_Hrs,
                              `Employee Equivalent` = NPF_Active + PF_Active + Terminated,
                              `Employee Group` = "Management & Excluded")]) %>%
      rbindlist()
    dt[, ":="(`Total Avg. Payroll Amount` = `Total Payroll Amount`/`Employee Equivalent`,
              `Total Avg. Straight Time Hours` = `Total Straight Time Hours`/`Employee Equivalent`)]
    
    p <- ggplot(dt[`Total Avg. Straight Time Hours` & `Total Avg. Payroll Amount` > 0], 
                aes(x = `Total Avg. Straight Time Hours`, y = `Total Avg. Payroll Amount`, color = `Employee Group`)) +
      geom_point(aes(text = map(paste("<b>Classification:</b>", Classification, "<br>",
                                      "<b>Employee Group:</b>", `Employee Group`, "<br>",
                                      "<b># of EE's:</b>", `Employee Equivalent`, "<br>",
                                      "<b>Avg. Hours per EE:</b>", format(round(`Total Avg. Straight Time Hours`, 2), nsmall = 2, big.mark = ","), "<br>",
                                      "<b>Avg. Payroll per EE:</b> $ ", format(round(`Total Avg. Payroll Amount`, 2), nsmall = 2, big.mark = ","), "<br>"),
                                HTML))) +
      geom_abline(slope = min_wage, intercept = 0, color = "red", linetype = "dashed") +
      scale_y_continuous(labels = scales::dollar_format()) +
      scale_x_continuous(labels = scales::comma_format()) +
      theme_light() +
      ggtitle("Total Avg. Payroll vs. Total Avg. Hours per EE") 
    
    
    p <- ggplotly(p, tooltip = c("text")) %>%
      layout(dragmode = "zoom")
  } else if(plotData == "Funding Source Composition") {
    if(funding == "By Union/Non-Union Programs") {
      dt <- copy(report$Home)[, .(`Union` = sum(Funding_Union, na.rm = T),
                                  `Non-Union` = sum(Funding_NonUnion, na.rm = T)), 
                              by = Fed_Prov] %>%
        pivot_longer(!Fed_Prov, names_to = "Program", values_to = "Funding") %>%
        pivot_wider(names_from = Fed_Prov, values_from = Funding) %>%
        as.data.table()
      if(!"Non-Provincial" %in% names(dt)) {
        dt <- dt[, `Non-Provincial` := c(0,0)]
      }
      dt <- dt[, RowSum := rowSums(.SD), .SDcols = c("Provincial", "Non-Provincial")
               ][, .(Program, Provincial, `Non-Provincial`, 
                     propProvincial = paste(ifelse(Provincial + `Non-Provincial` == 0, 0, round(Provincial/RowSum*100, 2)), "%"),
                     propNonProvincial = paste(ifelse(`Non-Provincial` + Provincial == 0, 0, round(`Non-Provincial`/RowSum*100, 2)), "%"))]
      
      
      p <- plot_ly(dt, x = ~`Program`, y = ~`Non-Provincial`, 
                   type = "bar", name = "Non-Provincially Funded",
                   color = I("royalblue"),
                   hovertemplate = paste(
                     "<b>Funding Amount: </b>%{y:$,.0f}"
                     )) %>%
        add_trace(y = ~`Provincial`, name = "Provincially Funded",
                  color = I("coral"),
                  hovertemplate = paste(
                    "<b>Funding Amount: </b>%{y:$,.0f}"
                    )) %>%
        layout(title = list(text = "Funding Source Composition"),
               yaxis = list(title = "Funding Amount"), 
               xaxis = list(title = "Program"), 
               barmode = "stack")
    } else if(funding == "By Funder") {
      dt <- copy(report$Home)[, .(Fed_Prov = factor(Fed_Prov, levels = c("Provincial", "Non-Provincial")), Funder, totFunding = `Total Funding`)] 
      
     p <- plot_ly(dt, x = ~Funder, y = ~totFunding,
                  type = "bar", color = ~`Fed_Prov`,
                  colors = c(`Non-Provincial` = "royalblue", `Provincial` = "coral"),
                  hovertemplate = paste(
                    "<b>%{y:$,.0f}</b>"
                  )) %>%
       layout(title = list(text = "Funding Source Composition"),
              yaxis = list(title = "Funding Amount"),
              xaxis = list(categoryorder = "total descending"),
              showlegend = T)
      
    }
    
  } else if(plotData == "Compensation Costs by Employee Group") {
    if(compensation == "By PF/NPF") {
      dt <- list(N1 = dt$N1[, .(`Funding Source` = c("Provincial", "Non-Provincial"),
                                `Payroll Amount` = c(sum(dt$S2$PF_NU, na.rm = T),
                                                     sum(dt$S2$NPF_NU, na.rm = T)),
                                `Employee Group` = "Non-Union")],
                 B1 = dt$B1[, .(`Funding Source` = c("Provincial", "Non-Provincial"),
                                `Payroll Amount` = c(sum(dt$S2$PF_BU, na.rm = T),
                                                     sum(dt$S2$NPF_BU, na.rm = T)),
                                `Employee Group` = "Bargainig Unit")],
                 M1 = dt$M1[, .(`Funding Source` = c("Provincial", "Non-Provincial"),
                                `Payroll Amount` = c(sum(dt$S2$PF_Mgmnt, na.rm = T),
                                                     sum(dt$S2$NPF_Mgmnt, na.rm = T)),
                                `Employee Group` = "Management & Excluded")])
      dt <- rbindlist(dt)
      dt <- as.data.table(pivot_wider(dt, names_from = `Funding Source`, values_from = `Payroll Amount`))
      
      p <- plot_ly(dt, x = ~`Employee Group`, y = ~`Non-Provincial`, 
                   type = "bar", name = "Non-Provincial",
                   color = I("royalblue"),
                   hovertemplate = paste("<b>%{y:$,.0f}</b>")) %>%
        add_trace(y = ~`Provincial`, color = I("coral"), name = "Provincial") %>%
        layout(title = list(text = "Funding Source Composition for Compensation Costs by Employee Group"),
               yaxis = list(title = "Compensation Costs"), 
               xaxis = list(title = "Employee Group", tickangle = 45), 
               barmode = "stack")
    } else if(compensation == "By Wage Cost Drivers") {
      dt <- list(N1 = copy(dt$S2)[, .(cat_1, cat_2, Cost, Total = PF_NU + NPF_NU,
                                      `Employee Group` = "Non-Union")],
                 B1 = copy(dt$S2)[, .(cat_1, cat_2, Cost, Total = PF_BU + NPF_BU,
                                      `Employee Group` = "Bargaining Unit")],
                 M1 = copy(dt$S2)[, .(cat_1, cat_2, Cost, Total = PF_Mgmnt + NPF_Mgmnt,
                                      `Employee Group` = "Management & Excluded")])
      dt <- rbindlist(dt)
      if(isTRUE(fine)) {
        dt <- copy(dt)[, .(Total = sum(Total)), by = .(`Employee Group`, cat_2)]
        
        p <- dt %>%
          plot_ly(x = ~`Employee Group`, y = ~`Total`, text = ~cat_2,
                  type = "bar", name = ~cat_2, color = ~cat_2,
                  hovertemplate = paste("<b>%{text}</b>",
                                        "%{y:$,.0f}")) %>%
          layout(barmode = "stack")
      } else {
        dt <- copy(dt)[, .(Total = sum(Total)), by = .(`Employee Group`, cat_1)]
        
        p <- dt %>%
          plot_ly(x = ~`Employee Group`, y = ~`Total`, text = ~cat_1,
                  type = "bar", name = ~cat_1, color = ~cat_1,
                  hovertemplate = paste("<b>%{text}</b>",
                                        "%{y:$,.0f}")) %>%
          layout(barmode = "stack")
      }
    }
  }
  
  return(p)
}
