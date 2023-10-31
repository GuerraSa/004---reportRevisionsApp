# Load relational data
source("data.R")

# Load schedule
load_schedule <- function(path) {
  sheet_names <- excel_sheets(path)[excel_sheets(path) %in% c("Home", 
                                                              paste0("Q", 1:3), 
                                                              paste0(rep(c("N", "M", "B", "S"), each = 2), 1:2), 
                                                              paste0("T", 1:4))]
  file <- vector(mode = "list", length = (1+length(sheet_names)))
  names(file) <- c("Contact", sheet_names)
  
  # Contact info
  file[["Contact"]] <- read_excel(path, sheet = "Home", col_names = F)[6:10, 1:2] %>%
    as.data.frame()
 
  # Home Schedule
  temp <-  read_excel(path, sheet = "Home", col_names = F)
  file[["Home"]] <- list("Provincial Funding" = read_excel(path, sheet = "Home", col_names = F)[20:58,],
                         "Non-Provincial Funding" = read_excel(path, sheet = "Home", col_names = F)[60:64,])
  
  # Schedules Q1-Q3
  ## Schedule Q1
  temp <- read_excel(path, sheet = "Q1", col_names = F)
  file[["Q1"]] <- list("Legal Status" = unlist(temp[[6,2]]),
                       "Employer Health Tax" = as.numeric(unlist(temp[[11,2]])),
                       "Service Subdivision" = temp[16:23, 1:2],
                       "EI Premium Reduction Program" = unlist(temp[[28,2]]),
                       "BC Housing Funding Supp." = unlist(temp[[32,2]]),
                       "CLBC Funding Supp." = data.frame(unlist(temp[c(35,39,42,46),1]), unlist(temp[c(36,40,44,47),2])),
                       "Live-In Home Support Workers" = unlist(temp[[52,2]]),
                       "Licensed Child Care" = unlist(temp[[57,2]]),
                       "Payroll Vendor/System" = data.frame(unlist(temp[62:64,1]), unlist(temp[62:64,2])),
                       "Group Benefit Provider" = data.frame(unlist(temp[69:71,1]), unlist(temp[69:71,2])),
                       "Pension or Retirement Plan" = data.frame(unlist(temp[76:78,1]), unlist(temp[76:78,2])),
                       "Short Term Illness and Injury Plan" = temp[92,2:7])
  ## Schedule Q2
  temp <- as.data.frame(read_excel(path, "Q2"))
  file[["Q2"]] <- list("SSO" = temp[str_which(temp[,1], "SSO")+2, 2],
                       "SSO Table" = temp[seq.int(str_which(temp[,1], "Facility Name")+1, length.out = 20),1:4],
                       "WorkSafeBC" = temp[str_which(temp[,1], "^WorkSafeBC")+2,2],
                       "WorkSafeBC Claims" = temp[str_which(temp[,1], "^WorkSafeBC")+4,2],
                       "Self Isolation" = temp[str_which(temp[,1], "^Self-Isolation")+2,2],
                       "Self Isolation Table" = temp[seq.int(str_which(temp[,1], "^Classification")+1, length.out = 20), 1:2],
                       "Mandatory Vaccination Status Order" = temp[str_which(temp[,1], "Mandatory")+2, 2])
  ## Schedule Q3
  temp <- read_excel(path, sheet = "Q3", col_names = F)
  if(ncol(temp) != 3){
    temp <- cbind(temp,
                  data.frame(rep.int(NA, nrow(temp)),
                             rep.int(NA, nrow(temp))))
  }
  file[["Q3"]] <- temp[-(1:5), c(1,3)]
  
  # Schedules N1, M1, and B1
  ## Schedule N1
  file[["N1"]] <- read_excel(path, sheet = "N1", skip = 15)[,1:23] %>% # Load Schedule N1 worksheet
    as.data.table()
  ## Schedule M1
  file[["M1"]] <- read_excel(path, sheet = "M1")[-(1:7),1:25] %>% # Load Schedule M1 worksheet
    as.data.table()
  ## Schedule B1
  file[["B1"]] <- read_excel(path, sheet = "B1", skip = 15)[,1:23] %>% # Load Schedule B1 worksheet
    as.data.table()
  
  # Schedules N2, M2, and B2
  ## Schedule N2
  file[["N2"]] <- read_excel(path, sheet = "N2")[-c(1:5, 58),] # Load Schedule N2
  ## Schedule M2
  file[["M2"]] <- read_excel(path, sheet = "M2")[-c(1:5, 58),] # Load Schedule M2
  ## Schedule B2
  file[["B2"]] <- read_excel(path, sheet = "B2")[-c(1:5, 58),] # Load Schedule B2
  
  # Schedule S1 & S2
  ## Schedule S1
  temp <- read_excel(path, sheet = "S1", col_names = F)
  file[["S1"]] <- list("# of Active Employees" = list("Regular vs Casual" = temp[11:13,1:7],
                                                      "By Region" = temp[19:23,1:7],
                                                      "By Union" = temp[30:nrow(temp),1:7]),
                       "Total Regular and Casual Hrs" = temp[10:13,9:15],
                       "Sick and Annual Leave Utilization" = temp[18:24,9:15],
                       "Total Sick Leave Wage Costs" = temp[29:30,9:15])
  ## Schedule S2
  file[["S2"]] <- read_excel(path, sheet = "S2", col_names = F)[-(1:10),]
  
  # Schedule T1-T4
  ## Schedule T1
  temp <- read_excel(path, sheet = "T1", col_names = F)
  file[["T1"]] <- list("Avg Time to Fill Vacancies" = list("Non-Union" = temp[8:16, c(1,2,4)],
                                                           "Management" = unlist(temp[[9,7]]),
                                                           "Bargaining Unit" = temp[8:16, c(9,10,12)]),
                       "Reasons for Termination" = temp[21:44,c(1,4:10)],
                       "Where do terminated employees go?" = temp[49:64,c(1,4:10)])
  ## Schedule T2
  file[["T2"]] <- read_excel(path, sheet = "T2", col_names = F)[-(1:3),1:21]
  ## Schedule T3
  file[["T3"]] <- read_excel(path, sheet = "T3", col_names = F)[-(1:3),1:21]
  ## Schedule T4
  file[["T4"]] <- read_excel(path, sheet = "T4", col_names = F)[-(1:3),1:21]
  
  return(file)
}

# Clean schedule formatting
clean_schedule <- function(file) {
  # Home Schedule
  cnames <- c("Funder", "Funding_Union", "Funding_NonUnion", "Total Funding", "Pct_Funding_Union", "Pct_Funding_NonUnion", "Pct_Funding_Tot", "N_UnionContracts", "N_NonUnionContracts", "N_TotalContracts")
  file[["Home"]][["Provincial Funding"]][[1,5]] <- NA
  
  names(file[["Home"]][["Provincial Funding"]]) <- cnames
  file[["Home"]][["Provincial Funding"]] <- file[["Home"]][["Provincial Funding"]] %>% 
    mutate_at(cnames[-1], as.numeric) %>% # Convert numeric columns 
    mutate_if(is.numeric, ~na_if(., 0)) %>% # Convert '0' to NA
    mutate(Fed_Prov = "Provincial") %>% # Create Prov/Non-Prov identifier column
    as.data.table() %>% # Convert to data.table
    setcolorder(c("Fed_Prov", cnames)) # Reorder columns
  
  names(file[["Home"]][["Non-Provincial Funding"]]) <- cnames
  file[["Home"]][["Non-Provincial Funding"]] <- file[["Home"]][["Non-Provincial Funding"]] %>% 
    mutate_at(cnames[-1], as.numeric) %>% # Convert numeric columns
    mutate_if(is.numeric, ~na_if(., 0)) %>% # Convert '0' to NA
    mutate(Fed_Prov = "Non-Provincial") %>% # Create Prov/Non-Prov identifier column
    as.data.table() %>% # Convert to data.table
    setcolorder(c("Fed_Prov", cnames)) # Reorder columns
  
  file[["Home"]] <- merge(file[["Home"]][["Provincial Funding"]], # Merge Provincial/Non-Provincial funding data
                          file[["Home"]][["Non-Provincial Funding"]], all = TRUE)
  file[["Home"]] <- file[["Home"]][rowSums(!is.na(file[["Home"]][,-(1:2)])) != 0] %>% # Keep only non-NA values
    setorder(cols = - "Fed_Prov")
  
  
  # Schedules Q1-Q3
  ## Schedule Q1
  names(file[["Q1"]][["CLBC Funding Supp."]]) <- c("Question", "Answer")
  names(file[["Q1"]][["Payroll Vendor/System"]]) <- c("Question", "Answer")
  names(file[["Q1"]][["Group Benefit Provider"]]) <- c("Question", "Answer")
  names(file[["Q1"]][["Pension or Retirement Plan"]]) <- c("Employee_Type", "Answer")
  ## Schedule Q2
  file[["Q2"]][["SSO"]] <- ifelse(file[["Q2"]][["SSO"]] == "Y", TRUE, FALSE)
  names(file[["Q2"]][["SSO Table"]]) <- c("Name", "Site_ID", "Classification", "FTE")
  file[["Q2"]][["SSO Table"]] <- file[["Q2"]][["SSO Table"]] %>%
    mutate_at(c("Site_ID", "FTE"), as.numeric) %>% # Convert numeric variables
    drop_na(Name) # Drop NA Rows
  names(file[["Q2"]][["Self Isolation Table"]]) <- c("Classification", "Count")
  file[["Q2"]][["Self Isolation Table"]] <- file[["Q2"]][["Self Isolation Table"]] %>%
    mutate_at("Count", as.numeric) %>%
    drop_na(Classification) # Drop NA rows
  ## Schedule Q3
  names(file[["Q3"]]) <- c("Question", "Answer")
  file[["Q3"]] <- file[["Q3"]] %>% drop_na(Question) # Drop NA rows
  
  # Schedules N1, M1, and B1
  ## Schedule N1
  cnames <- c("Classification", "RegularOrCasual", "YrlyStdHrs", "NPF_Hrs_StraightTime", "PF_Hrs_StraightTime", "NPF_HrlyWage", "PF_HrlyWage", "NPF_Active", "PF_Active", "PF_LTD", "PF_WCB", "PF_MPLeave", "PF_Other", "Vacant", "Terminated", "NewHiresExternal", "NewHiresInternal", "Backfill_StraightTimeHrs", "Bacfill_pct", "SSO_NPF_Hrs", "SSO_NPF_HrlyWage", "SSO_PF_Hrs", "SSO_PF_HrlyWage") # Column names
  names(file[["N1"]]) <- cnames # Rename columns
  file[["N1"]] <- as.data.table(file[["N1"]]) # Convert to data.table
  file[["N1"]][, row_id:=(1:nrow(file[["N1"]])+16)] # Create row id column
  file[["N1"]] <- file[["N1"]] %>% drop_na(Classification) # Drop NA rows
  file[["N1"]] <- file[["N1"]] %>% mutate_at(cnames[-(1:2)], as.numeric) %>% # Convert numeric columns
    mutate_if(is.numeric, ~na_if(., 0)) %>% # Replace 0 values with NA 
    mutate(YrlyStdHrs = round(YrlyStdHrs,10), # Round Values to 2 decimal places
           NPF_Hrs_StraightTime = round(NPF_Hrs_StraightTime,2),
           PF_Hrs_StraightTime = round(PF_Hrs_StraightTime,2)) 
  ## Schedule M1
  cnames <- c("Classification", "Gender", "AvgAnnualSalary", "NPF_Payroll", "NPF_Expenses", "PF_Payroll", "PF_Expenses", "NPF_Hrs", "PF_Hrs", "NPF_Active", "PF_Active", "PF_LTD", "PF_WCB", "PF_MPLeave", "PF_Other", "Vacant", "Terminated", "NewHiresExternal", "NewHiresInternal", "Backfill_StraightTimeHrs", "Bacfill_pct", "SSO_NPF_Hrs", "SSO_NPF_HrlyWage", "SSO_PF_Hrs", "SSO_PF_HrlyWage") # Column names
  names(file[["M1"]]) <- cnames # Rename columns
  file[["M1"]] <- as.data.table(file[["M1"]]) # Convert to data.table
  file[["M1"]][, row_id := (1:nrow(file[["M1"]])+16)] # Create row id column
  file[["M1"]] <- file[["M1"]] %>% drop_na(Classification) # Drop NA rows
  file[["M1"]] <- file[["M1"]] %>% mutate_at(cnames[-(1:2)], as.numeric) %>% # Convert numeric columns 
    mutate_if(is.numeric, ~na_if(., 0)) %>% # Replace 0 values with NA
    mutate(AvgAnnualSalary = round(AvgAnnualSalary,2), # Round values to 2 decimal places
           PF_Payroll = round(PF_Payroll,2), 
           NPF_Payroll = round(NPF_Payroll,2))
  ## Schedule B1
  cnames <- c("Classification", "RegularOrCasual", "YrlyStdHrs", "NPF_Hrs_StraightTime", "PF_Hrs_StraightTime", "NPF_HrlyWage", "PF_HrlyWage", "NPF_Active", "PF_Active", "PF_LTD", "PF_WCB", "PF_MPLeave", "PF_Other", "Vacant", "Terminated", "NewHiresExternal", "NewHiresInternal", "Backfill_StraightTimeHrs", "Bacfill_pct", "SSO_NPF_Hrs", "SSO_NPF_HrlyWage", "SSO_PF_Hrs", "SSO_PF_HrlyWage") # Column names
  names(file[["B1"]]) <- cnames # Rename columns
  file[["B1"]] <- as.data.table(file[["B1"]]) # Convert to data.table
  file[["B1"]][, row_id := (1:nrow(file[["B1"]])+16)] # Create row id column
  file[["B1"]] <- file[["B1"]] %>% drop_na(Classification) # Drop NA rows
  file[["B1"]] <- file[["B1"]] %>% mutate_at(cnames[-(1:2)], as.numeric) %>% # Convert numeric columns
    mutate_if(is.numeric, ~na_if(., 0)) %>% # Replace 0 values with NA
    mutate(YrlyStdHrs = round(YrlyStdHrs,2), # Round Values to 2 decimal places
           NPF_Hrs_StraightTime = round(NPF_Hrs_StraightTime,2),
           PF_Hrs_StraightTime = round(PF_Hrs_StraightTime,2)) 
  
  # Schedules N2, M2, and B2
  cnames <- c("LengthofService", "Seniority_Regular", "", "Seniority_Casual", "", "Age", "AgeGender_Regular_M", "AgeGender_Regular_F", "AgeGender_Regular_GD", "", "AgeGender_Casual_M", "AgeGender_Casual_F", "AgeGender_Casual_GD", "", "Benefit Type", "Part_Single", "Part_Couple", "Part_Family", "NonPart_Eligible", "NonPart_Ineligible", "Total") # Column names
  col.num <- cnames[which(!cnames %in% c("LengthofService", "Age", "Benefit Type", ""))]
  ## Schedule N2
  names(file[["N2"]]) <- cnames # Rename columns
  file[["N2"]][col.num] <- sapply(file[["N2"]][col.num], as.numeric) # Convert numeric columns
  file[["N2"]] <- file[["N2"]] %>% 
    as.data.table() # Convert to data.table
  ## Schedule B2
  names(file[["B2"]]) <- cnames # Rename columns
  file[["B2"]][col.num] <- sapply(file[["B2"]][col.num], as.numeric) # Convert numeric columns
  file[["B2"]] <- file[["B2"]] %>% 
    as.data.table() # Convert to data.table
  ## Schedule M2
  cnames <- c("LengthofService", "All", "", "ED_Only", "", "Age", "AgeGender_M", "AgeGender_F", "AgeGender_GD", "", "", "", "", "", "Benefit Type", "Part_Single", "Part_Couple", "Part_Family", "NonPart_Eligible", "NonPart_Ineligible", "Total") # Rename columns
  names(file[["M2"]]) <-  cnames
  col.num <- cnames[which(!cnames %in% c("LengthofService", "Age", "Benefit Type", ""))]
  file[["M2"]][col.num] <- sapply(file[["M2"]][col.num], as.numeric) # Convert numeric columns
  file[["M2"]] <- file[["M2"]] %>% 
    as.data.table() # Convert to data.table
  
  # Schedule S1 & S2
  cnames <- c("PF_NU", "PF_Mgmnt", "PF_BU", "NPF_NU", "NPF_Mgmnt", "NPF_BU")
  ## Schedule S1
  names(file[["S1"]][["# of Active Employees"]][["Regular vs Casual"]]) <- c("Classication", cnames) # Rename columns
  names(file[["S1"]][["# of Active Employees"]][["By Region"]]) <- c("Region", cnames) # Rename columns
  names(file[["S1"]][["# of Active Employees"]][["By Union"]]) <- c("Union", cnames) # Rename columns
  names(file[["S1"]][["Total Regular and Casual Hrs"]]) <- c("Classifcation", cnames) # Rename columns
  names(file[["S1"]][["Sick and Annual Leave Utilization"]]) <- c("Type of Leave", cnames) # Rename columns
  names(file[["S1"]][["Total Sick Leave Wage Costs"]]) <- c("Classifcation", cnames) # Rename columns
  for(i in seq_along(file[["S1"]])) { # Convert numeric columns
    if(inherits(file[["S1"]][[i]], "list")) {
      for(j in seq_along(file[["S1"]][[i]])) {
        file[["S1"]][[i]][[j]] <- file[["S1"]][[i]][[j]] %>%
          mutate_at(vars(-1), as.numeric)
      }
    }
    else {
      file[["S1"]][[i]] <- file[["S1"]][[i]] %>%
        mutate_at(vars(-1), as.numeric) %>%
        mutate_if(is.numeric, ~na_if(., 0)) # Convert '0' to NA
    }
  } 
  ## Schedule S2
  ind <- unique(which(file[["S2"]] == "$", arr.ind = T)[,1])
  file[["S2"]] <- cbind(file[["S2"]][,1:2], sapply(file[["S2"]][,3:8], as.numeric)) %>% # Convert numeric columns
    mutate_if(is.numeric, ~na_if(., 0)) %>% # Convert '0' to NA
    as.data.table() # Convert to data.table
  file[["S2"]][is.na(file[["S2"]][[2]]), 2] <- file[["S2"]][is.na(file[["S2"]][[2]]), 1]
  file[["S2"]] <- file[["S2"]][-ind,]
  names(file[["S2"]]) <- c("","Cost", cnames) # Rename columns
  file[["S2"]] <- file[["S2"]][, .(cat_1 = c(rep("Wage Costs", 6),
                                             rep("Expenses & Allowances", 3),
                                             rep("Benefit costs", 13)),
                                   cat_2 = c(rep("Straight Time Pay", 2),
                                             rep("Premium Pay", 2),
                                             "Vacation & Holiday Pay", "All Other Wage Costs",
                                             rep("Expenses & Alowances", 3),
                                             rep("Statuatory Benefits", 3),
                                             rep("Group Benefits", 7),
                                             rep("Super Annuation", 3)),
                                   Cost, PF_NU, PF_Mgmnt, PF_BU, NPF_NU, NPF_Mgmnt, NPF_BU)]
  
  
  # Schedule T1-T4
  ## Schedule T1
  file[["T1"]][[1]][["Non-Union"]] <- file[["T1"]][[1]][["Non-Union"]] %>%
    rename("Classification Type" = 1, "Classification" = 2, "Days to Fill" = 3) %>%
    mutate_at("Days to Fill", as.numeric) %>% # Rename variables and convert Days to fill to numeric
    as.data.table()
  file[["T1"]][[1]][["Management"]] <- as.numeric(file[["T1"]][[1]][["Management"]]) # Convert Mgmnt days to numeric
  file[["T1"]][[1]][["Bargaining Unit"]] <- file[["T1"]][[1]][["Bargaining Unit"]] %>%
    rename("Classification Type" = 1, "Classification" = 2, "Days to Fill" = 3) %>%
    mutate_at("Days to Fill", as.numeric) %>% # Rename variables and convert Days to fill to numeric
    as.data.table()
  file[["T1"]][[2]] <- file[["T1"]][[2]][-(1:2),] %>%
    rename("Termination Reason" = 1, "NU_Paraprofessional" = 2, "NU_Benchmark" = 3, "NU_Delegated" = 4,
           "Management" = 5, "BU_Paraprofessional" = 6, "BU_Benchmark" = 7, "BU_Delegated" = 8) %>%
    mutate_at(vars(2:8), as.numeric) %>% # Rename variables and convert counts to numeric
    as.data.table()
  file[["T1"]][[3]] <- file[["T1"]][[3]][-(1:2),] %>%
    rename("Where they go" = 1, "NU_Paraprofessional" = 2, "NU_Benchmark" = 3, "NU_Delegated" = 4,
           "Management" = 5, "BU_Paraprofessional" = 6, "BU_Benchmark" = 7, "BU_Delegated" = 8) %>%
    mutate_at(vars(2:8), as.numeric) %>% # Rename variables and convert counts to numeric
    as.data.table()
  ## Schedule T2
  file[["T2"]][3,1:3] <- t(c("Classification", "Regular Or Casual", "Total Terminated"))
  file[["T2"]] <- list("Age" = file[["T2"]][-c(1:2, 4:5),c(1:3, 4:9)],
                       "Gender" = file[["T2"]][-c(1:2, 4:5),c(1:3, 10:12)],
                       "Length of Service" = file[["T2"]][-c(1:2, 4:5),c(1:3, 13:16)],
                       "Region" = file[["T2"]][-c(1:2, 4:5),c(1:3, 17:21)]) # Remove unnecessary rows and split different demographic data
  for(i in seq_along(file[["T2"]])) {
    file[["T2"]][[i]] <- file[["T2"]][[i]] %>%
      row_to_names(row_number = 1) %>% # Convert 1st row to col names
      mutate_at(vars(-(1:3)), as.numeric) %>% # Convert numeric columns 
      drop_na(Classification) %>% # Drop empty rows
      as.data.table()
  }
  ## Schedule T3
  file[["T3"]][3,1:3] <- t(c("Classification", "Regular Or Casual", "Total Terminated"))
  file[["T3"]] <- list("Age" = file[["T3"]][-c(1:2, 4:5),c(1:3, 4:9)],
                       "Gender" = file[["T3"]][-c(1:2, 4:5),c(1:3, 10:12)],
                       "Length of Service" = file[["T3"]][-c(1:2, 4:5),c(1:3, 13:16)],
                       "Region" = file[["T3"]][-c(1:2, 4:5),c(1:3, 17:21)]) # Remove unnecessary rows and split different demographic data
  for(i in seq_along(file[["T3"]])) {
    file[["T3"]][[i]] <- file[["T3"]][[i]] %>%
      row_to_names(row_number = 1) %>% # Convert 1st row to col names
      mutate_at(vars(-(1:3)), as.numeric) %>% # Convert numeric columns 
      drop_na(Classification) %>% # Drop empty rows
      as.data.table()
  }
  ## Schedule T4
  file[["T4"]][3,1:3] <- t(c("Classification", "Regular Or Casual", "Total Terminated"))
  file[["T4"]] <- list("Age" = file[["T4"]][-c(1:2, 4:5),c(1:3, 4:9)],
                       "Gender" = file[["T4"]][-c(1:2, 4:5),c(1:3, 10:12)],
                       "Length of Service" = file[["T4"]][-c(1:2, 4:5),c(1:3, 13:16)],
                       "Region" = file[["T4"]][-c(1:2, 4:5),c(1:3, 17:21)]) # Remove unnecessary rows and split different demographic data
  for(i in seq_along(file[["T4"]])) {
    file[["T4"]][[i]] <- file[["T4"]][[i]] %>%
      row_to_names(row_number = 1) %>% # Convert 1st row to col names
      mutate_at(vars(-(1:3)), as.numeric) %>% # Convert numeric columns 
      drop_na(Classification) %>% # Drop empty rows
      as.data.table()
  }
  
  # Contact + Funding Info
  h <- c(setNames(c(sum(file[["N1"]][, PF_Active], na.rm = TRUE), sum(file[["M1"]][, PF_Active], na.rm = TRUE), sum(file[["B1"]][, PF_Active], na.rm = TRUE)),
                  c("PF_NU", "PF_Mgmnt", "PF_BU")), # PF headcounts
         setNames(c(sum(file[["N1"]][, NPF_Active], na.rm = TRUE), sum(file[["M1"]][, NPF_Active], na.rm = TRUE), sum(file[["B1"]][, NPF_Active], na.rm = TRUE)),
                  c("NPF_NU", "NPF_Mgmnt", "NPF_BU"))) # NPF headcounts
  names(file[["Contact"]]) <- NULL
  file[["Contact"]][nrow(file[["Contact"]])+1,] <- list("Total Provincial Funding:", 
                                                        paste0("$", format(round(sum(file[["Home"]][Fed_Prov=="Provincial"]$`Total Funding`, na.rm = TRUE), 2), big.mark = ","), " (",
                                                               format(round(sum(file[["Home"]][Fed_Prov=="Provincial"]$`Total Funding`, na.rm = TRUE)*100/
                                                                              sum(file[["Home"]][Fed_Prov=="Provincial"]$`Total Funding`, file[["Home"]][Fed_Prov=="Non-Provincial"]$`Total Funding`, na.rm = TRUE), 1), big.mark = ","), "%)"))
  file[["Contact"]][nrow(file[["Contact"]])+1,] <- list("Total Non-Provincial Funding:", 
                                                        paste0("$", format(round(sum(file[["Home"]][Fed_Prov=="Non-Provincial"]$`Total Funding`, na.rm = TRUE), 2), big.mark = ","), " (",
                                                               format(round(sum(file[["Home"]][Fed_Prov=="Non-Provincial"]$`Total Funding`, na.rm = TRUE)*100/
                                                                              sum(file[["Home"]][Fed_Prov=="Provincial"]$`Total Funding`, file[["Home"]][Fed_Prov=="Non-Provincial"]$`Total Funding`, na.rm = TRUE), 1), big.mark = ","), "%)"))
  file[["Contact"]][nrow(file[["Contact"]])+1,] <- list("Total Compensation (Utilization):",
                                                        paste0("$", format(round(sum(file[["S2"]][,-(1:3)], na.rm = TRUE), 2), big.mark = ","), " (",
                                                               round(sum(file[["S2"]][,-(1:3)], na.rm = TRUE)*100/sum(parse_number(unlist(file[["Contact"]][(6:7),2]))), 1), "%)"))

  return(file)
}

# Perform revisions
revisions <- function(file) {
  emp_rev <- function(file, s) { # Employee Group Revisions
    if(as.numeric(str_extract(s, "[:digit:]")) == 1) { # N1, B1, M1
      temp <- copy(file[[s]])
      cols <- str_which(names(temp), "(?<!SSO_(PF|NPF)_)HrlyWage$")
      setnames(temp, cols, str_replace_all(names(temp)[cols], "HrlyWage", "Payroll"))
      cols <- str_which(names(temp), "^(NPF|PF)_Hrs.+")
      setnames(temp, cols, str_remove_all(names(temp)[cols], "_StraightTime$"))
      cols <- names(temp)
      
      output <- data.table(Data = c("Status Check", "PF/NPF Wage + Hours Check", "PF/NPF Data Seggregation", "NPF Wage Distribution", "PF Wage Distribution", "Headcounts", "SSO Data"),
                           Notes = c(case_when(rbind(temp[PF_Active == (NPF_Active|PF_LTD|PF_WCB|PF_MPLeave|PF_Other|Terminated)], # Duplicate
                                                     temp[temp[, Reduce(`&`, lapply(.SD, is.na)), .SDcols = str_subset(cols, "PF_(Active|LTD|WCB|MPLeave|Other)|Terminated")]])[, .N] == 0 ~ paste0(icons::ionicons("checkmark")),
                                               TRUE ~ paste0(ifelse(temp[temp[, Reduce(`&`, lapply(.SD, is.na)), .SDcols = str_subset(cols, "PF_(Active|LTD|WCB|MPLeave|Other)|Terminated")]][,.N] == 0, "",
                                                                    paste0("**Missing status for:**<br>",
                                                                           paste0(temp[temp[, Reduce(`&`, lapply(.SD, is.na)), .SDcols = str_subset(cols, "PF_(Active|LTD|WCB|MPLeave|Other)|Terminated")]]$Classification, " (Line ",
                                                                                  paste(temp[temp[, Reduce(`&`, lapply(.SD, is.na)), .SDcols = str_subset(cols, "PF_(Active|LTD|WCB|MPLeave|Other)|Terminated")]]$row_id), ")",
                                                                                  collapse = "<br>"),
                                                                           ifelse(temp[PF_Active == (NPF_Active|PF_LTD|PF_WCB|PF_MPLeave|PF_Other|Terminated)][, .N] != 0, "<br><br>", ""))),
                                                             ifelse(temp[PF_Active == (NPF_Active|PF_LTD|PF_WCB|PF_MPLeave|PF_Other|Terminated)][, .N] == 0, "",
                                                                    paste0("**Possible Duplicate Headcounts for:**<br>",
                                                                           paste0(temp[PF_Active == (PF_LTD|PF_WCB|PF_MPLeave|PF_Other|Terminated)]$Classification, " (Line ",
                                                                                  paste(temp[PF_Active == (PF_LTD|PF_WCB|PF_MPLeave|PF_Other|Terminated)]$row_id), ")",
                                                                                  collapse = "<br>"))))),
                                     case_when(nrow(temp[(is.na(PF_Payroll) & !is.na(PF_Hrs)) | (is.na(PF_Hrs) & !is.na(PF_Payroll)) | (is.na(NPF_Payroll) & !is.na(NPF_Hrs)) | (is.na(NPF_Hrs) & !is.na(NPF_Payroll))]) == 0 ~ paste0(icons::ionicons("checkmark")) # PF/NPF Wage + Hours Check
                                               ,TRUE ~ paste0(case_when(nrow(temp[is.na(PF_Payroll) & !is.na(PF_Hrs)]) == 0 ~ "" # Missing PF Hrly Wage
                                                                        ,TRUE ~ paste0("**Missing PF Wage Data for:**<br>",
                                                                                       paste0(temp[is.na(PF_Payroll) & !is.na(PF_Hrs)]$Classification, " (Line ", 
                                                                                              temp[is.na(PF_Payroll) & !is.na(PF_Hrs)]$row_id, ")<br>",
                                                                                              collapse = ""))),
                                                              case_when(nrow(temp[is.na(PF_Hrs) & !is.na(PF_Payroll)]) == 0 ~ "" # Missing PF Hrs Straight Time
                                                                        ,TRUE ~ paste0("**Missing PF Hrs Data for:**<br>",
                                                                                       paste0(temp[is.na(PF_Hrs) & !is.na(PF_Payroll)]$Classification, " (Line ", 
                                                                                              temp[is.na(PF_Hrs) & !is.na(PF_Payroll)]$row_id, ")<br>",
                                                                                              collapse = ""))),
                                                              case_when(nrow(temp[is.na(NPF_Payroll) & !is.na(NPF_Hrs)]) == 0 ~ "" # Missing NPF Hrly Wage
                                                                        ,TRUE ~ paste0("**Missing NPF Wage Data for:**<br>",
                                                                                       paste0(temp[is.na(NPF_Payroll) & !is.na(NPF_Hrs)]$Classification, " (Line ", 
                                                                                              temp[is.na(NPF_Payroll) & !is.na(NPF_Hrs)]$row_id, ")<br>",
                                                                                              collapse = ""))),
                                                              case_when(nrow(temp[is.na(NPF_Hrs) & !is.na(NPF_Payroll)]) == 0 ~ "" # Missing NPF Hrs Straight Time
                                                                        ,TRUE ~ paste0("**Missing NPF Hrs Data for:**<br>",
                                                                                       paste0(temp[is.na(NPF_Hrs) & !is.na(NPF_Payroll)]$Classification, " (Line ", 
                                                                                              temp[is.na(NPF_Hrs) & !is.na(NPF_Payroll)]$row_id, ")<br>",
                                                                                              collapse = ""))))),
                                     case_when(temp[!is.na(NPF_Active) & !is.na(PF_Active), .N] == 0 ~ paste0(icons::ionicons("checkmark")), # PF/NPF Data Seggregation
                                               TRUE ~ paste0("**NPF/PF headcounts indicated on same row for:**<br>", 
                                                             paste0(temp[!is.na(NPF_Active) & !is.na(PF_Active),]$Classification, " (Line ",
                                                                    temp[!is.na(NPF_Active) & !is.na(PF_Active),]$row_id, ")",
                                                                    collapse = "<br>"))),
                                     case_when(sum(temp[, .(NPF_Payroll, PF_Payroll)], na.rm = TRUE) == 0 ~ "No Payroll Data" # Non-Provincial Salary Distribution
                                               ,any(c(is.na(range(temp$NPF_Payroll, na.rm = TRUE)), is.infinite(range(temp$NPF_Payroll, na.rm = TRUE)))) ~ "No NPF Payroll Data" 
                                               ,TRUE ~ paste0(c("**Min:** ", "**Max:** "), paste0("$", format(range(temp$NPF_Payroll, na.rm = TRUE), big.mark = ",")), collapse = "<br>")),
                                     case_when(sum(temp[, .(NPF_Payroll, PF_Payroll)], na.rm = TRUE) == 0 ~ "No Payroll Data" # Provincial Salary Distribution
                                               ,any(c(is.na(range(temp$PF_Payroll, na.rm = TRUE)), is.infinite(range(temp$PF_Payroll, na.rm = TRUE)))) ~ "No PF Payroll Data" 
                                               ,TRUE ~ paste0(c("**Min:** ", "**Max:** "), paste0("$", format(range(temp$PF_Payroll, na.rm = TRUE), big.mark = ",")), collapse = "<br>")),
                                     case_when(sum(temp[, .(NPF_Active, PF_Active)], na.rm = TRUE) == 0 ~ "No Headcount Data", # Headcounts
                                               TRUE ~ paste0("**NPF Headcount:** ", h[paste0("NPF_", ifelse(s=="M1", "Mgmnt", paste0(str_sub(s, 1, 1), "U")))],
                                                             "<br>**PF Headcount:** ", h[paste0("PF_", ifelse(s=="M1", "Mgmnt", paste0(str_sub(s, 1, 1), "U")))])),
                                     case_when(temp[PF_Hrs == SSO_PF_Hrs | PF_Payroll == SSO_PF_HrlyWage | NPF_Hrs == SSO_NPF_Hrs | NPF_Payroll == SSO_NPF_HrlyWage, .N] > 0 ~ paste0("**Evidence of copy/paste for:**<br>", # SSO data check
                                                                                                                                                                                      paste0(temp[PF_Hrs == SSO_PF_Hrs | PF_Payroll == SSO_PF_HrlyWage | NPF_Hrs == SSO_NPF_Hrs | NPF_Payroll == SSO_NPF_HrlyWage]$Classification, " (Line ", 
                                                                                                                                                                                             temp[PF_Hrs == SSO_PF_Hrs | PF_Payroll == SSO_PF_HrlyWage | NPF_Hrs == SSO_NPF_Hrs | NPF_Payroll == SSO_NPF_HrlyWage]$row_id, ")<br>",
                                                                                                                                                                                             collapse = "")),
                                               !file[["Q2"]][["SSO"]] & (temp[!is.na(SSO_PF_Hrs) | !is.na(SSO_PF_HrlyWage) | !is.na(SSO_NPF_Hrs) | !is.na(SSO_NPF_HrlyWage), .N] > 0) ~ paste0("**NO SSO Status, but data enetered for:**<br>",
                                                                                                                                                                                               paste0(temp[!is.na(SSO_PF_Hrs) | !is.na(SSO_PF_HrlyWage) | !is.na(SSO_NPF_Hrs) | !is.na(SSO_NPF_HrlyWage)]$Classification, " (Line ", 
                                                                                                                                                                                                      temp[!is.na(SSO_PF_Hrs) | !is.na(SSO_PF_HrlyWage) | !is.na(SSO_NPF_Hrs) | !is.na(SSO_NPF_HrlyWage)]$row_id, ")<br>",
                                                                                                                                                                                                      collapse = "")),
                                               TRUE ~ paste0(icons::ionicons("checkmark")))),
                           Schedule = paste(s, "Schedule"))
      if(s == "M1") {
        temp[, TotPayroll := sum(NPF_Payroll, PF_Payroll, na.rm = TRUE), by = row_id
        ][, PayrollCopyPaste := (round(AvgAnnualSalary/TotPayroll, 2)==1)]
        output <- rbind(data.table(Data = c("Avg Salary != Payroll Amount", "Avg Salary > Payroll Amount"),
                                   Notes = c(case_when(sum(temp$TotPayroll, na.rm = TRUE) == 0 ~ "No Payroll Data", # Avg Salary != Payroll Amount
                                                       temp[PayrollCopyPaste == TRUE, .N] != 0 ~ paste0("**Evidence of copy/paste for:**<br>",
                                                                                                        paste0(temp[PayrollCopyPaste == TRUE]$Classification, " (Line ",
                                                                                                               temp[PayrollCopyPaste == TRUE]$row_id, ")",
                                                                                                               collapse = "<br>")),
                                                       TRUE ~ paste0(icons::ionicons("checkmark"))),
                                             case_when(sum(temp$TotPayroll, na.rm = TRUE) == 0 ~ "No Payroll Data", # Avg Salary > Payroll Amount
                                                       temp[AvgAnnualSalary < TotPayroll, .N] != 0 ~ paste0("**Payroll Amount > Avg. Salary for:**<br>",
                                                                                                            paste0(temp[AvgAnnualSalary < TotPayroll]$Classification, " (Line ",
                                                                                                                   temp[AvgAnnualSalary < TotPayroll]$row_id, ")",
                                                                                                                   collapse = "<br>")),
                                                       TRUE ~ paste0(icons::ionicons("checkmark")))),
                                   Schedule = paste(s, "Schedule")),
                        output)
      } else {
        output <- rbind(data.table(Data = c("Std Hrs per Yr != Hrs Paid at Straight Time", "Director/Manager check"),
                                   Notes = c(case_when(sum(temp[, .(NPF_Payroll, PF_Payroll)], na.rm = TRUE) == 0 ~ "No Payroll Data" # Std Hrs per Yr != Hrs Paid at Straight Time
                                                       ,nrow(temp[YrlyStdHrs == PF_Hrs | YrlyStdHrs == NPF_Hrs]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                       ,TRUE ~ paste0("**Evidence of Copy/Paste for:**<br>",
                                                                      paste0(temp[YrlyStdHrs == PF_Hrs | YrlyStdHrs == NPF_Hrs]$Classification, " (Line ",
                                                                             paste(temp[YrlyStdHrs == PF_Hrs | YrlyStdHrs == NPF_Hrs]$row_id), ")",
                                                                             collapse = "<br>"))),
                                             case_when(nrow(temp[str_which(Classification, "Manager|Director")]) == 0 ~ paste0(icons::ionicons("checkmark")) # Management/Director Seggregation check
                                                       ,TRUE ~ paste0("**The following classifications may have ability to hire/fire:**<br>",
                                                                      paste0(temp[str_which(Classification, "Manager|Director")]$Classification, " (Line ", 
                                                                             temp[str_which(Classification, "Manager|Director")]$row_id, ")<br>",
                                                                             collapse = "")))),
                                   Schedule = paste(s, "Schedule")),
                        output)
      }
    }
    else if(as.numeric(str_extract(s, "[:digit:]")) == 2) { # N2, B2, M2
      ags <- setNames(c(sum(file[[s]][, .SD, .SDcols = patterns("Seniority|All")], na.rm = TRUE),
                        sum(file[[s]][, .SD, .SDcols = patterns("AgeGender")], na.rm = TRUE)),
                      c("Seniority", "AgeGender"))
      gb <- setNames(rowSums(file[[s]][1:5, 16:20], na.rm = TRUE),
                     unlist(file[[s]][1:5, 15])) 
      dat <- ifelse(s=="M2", "PF_Mgmnt", paste0("PF_", str_sub(s, 1, 1), "U"))
      gb <- gb[gb != h[dat]]
      
      output <- data.table(Data = c("Seniority", "AgeGender", "Group Benefit Participation"),
                           Notes = c(case_when(ags["Seniority"] == h[dat] ~ paste0(icons::ionicons("checkmark")) # Seniority headcounts
                                               ,TRUE ~ paste0(str_sub(s, 1, 1), "1 PF Active Headcount: ", h[dat], 
                                                              "<br>", s, " Headcount: ", ags["Seniority"])),
                                     case_when(ags["AgeGender"] == h[dat] ~ paste0(icons::ionicons("checkmark")) # AgeGender Headcounts
                                               ,TRUE ~ paste0(str_sub(s, 1, 1), "1 PF Active Headcount: ", h[dat], 
                                                              "<br>", s, " Headcount: ", ags["AgeGender"])),
                                     case_when(all(gb == h[dat]) ~ paste0(icons::ionicons("checkmark")) # GB Participation headcounts
                                               ,TRUE ~ paste0(str_sub(s, 1, 1), "1 PF Active Headcount: ", h[dat],
                                                              "<br>", paste0(str_c(names(gb), gb, sep = ": "), collapse = "<br>")))),
                           Schedule = paste(s, "Schedule"))
    }
    
    return(output)
  }
  h <- c(setNames(c(sum(file[["N1"]][, PF_Active], na.rm = TRUE), sum(file[["M1"]][, PF_Active], na.rm = TRUE), sum(file[["B1"]][, PF_Active], na.rm = TRUE)),
                  c("PF_NU", "PF_Mgmnt", "PF_BU")), # PF headcounts
         setNames(c(sum(file[["N1"]][, NPF_Active], na.rm = TRUE), sum(file[["M1"]][, NPF_Active], na.rm = TRUE), sum(file[["B1"]][, NPF_Active], na.rm = TRUE)),
                  c("NPF_NU", "NPF_Mgmnt", "NPF_BU"))) # NPF headcounts
  output <- vector(mode = "list", length = 16)
  names(output) <- names(file[-1])
  # Home Schedule
  output[["Home"]] <- copy(file[["Home"]])
  output[["Home"]] <- output[["Home"]] %>% 
    mutate(Notes = case_when(Funding_Union > 0 & (N_UnionContracts == 0 | is.na(N_UnionContracts)) ~ "Missing # of BU contracts"
                             ,N_UnionContracts > 0 & (Funding_Union == 0 | is.na(Funding_Union)) ~ "Missing $ value of BU contracts"
                             ,Funding_NonUnion > 0 & (N_NonUnionContracts == 0 | is.na(N_NonUnionContracts)) ~ "Missing # of NU contracts"
                             ,N_NonUnionContracts > 0 & (Funding_NonUnion == 0 | is.na(Funding_NonUnion)) ~ "Missing $ value of NU contracts"
                             ,TRUE ~ "No Issues")) %>% 
    select(Funder, Funding_Union, Funding_NonUnion, N_UnionContracts, N_NonUnionContracts, Notes) %>%
    as.data.table() # Convert to data.table
   convert_and_format(output[["Home"]], "$", names(output[["Home"]])[2:3])
   cols <- names(output[["Home"]])[sapply(output[["Home"]], is.numeric)]
   output[["Home"]][, (cols) := lapply(.SD, function(x) ifelse(is.na(x), "", as.character(x))), .SDcols = cols]
  # Q1
   output[["Q1"]] <- data.table(Data = c("Legal Status", "Employer Health Tax", "Service Subdivision", "EI Premium Reduction Program", 
                                         "BC Housing - Supplementary Q's","CLBC Funding - Supplementary Q's", "Live-In Home Support Workers", 
                                         "Licensed Child Care", "Payroll Vendor/System", "Group Benefit Provider", "Pension or Retirement Plan", 
                                         "Short Term Illness and Injury Plan"),
                                Notes = c(case_when(is.na(file[["Q1"]][["Legal Status"]]) ~ "Missing" # Legal status filled
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(sum(file[["Home"]]$`Total Funding`, na.rm = TRUE) > 1000000 & file[["Q1"]][[2]] < 0 ~ "EHT data might be missing" # EHT information for funding > $1,000,000
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(sum(is.na(file[["Q1"]][[3]][2])) == 8 ~ "Missing" # Service subdivision 
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(is.na(file[["Q1"]][[4]]) ~ "Missing Y/N Answer" # EI Premium Reduction Program
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(any(file[["Home"]][Funder=="BC Housing"]$`Total Funding` > 0) & is.na(file[["Q1"]][[5]]) ~ "Missing" # BC Housing Supp. Q's
                                                    ,nrow(file[["Home"]][Funder=="BC Housing"]) == 0 ~ "No BC Housing Funding"
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(any(file[["Home"]][Funder=="Community Living BC"]$`Total Funding` > 0) & sum(!is.na(file[["Q1"]][[6]][2])) != 4 ~ "Missing" # CLBC Supp. Q's
                                                    ,nrow(file[["Home"]][Funder=="Community Living BC"]) == 0 ~ "No CLBC Funding"
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(is.na(file[["Q1"]][[7]]) ~ "Missing Y/N Answer" # Live-in Home Support Workers
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(is.na(file[["Q1"]][[8]]) ~ "Missing Y/N Answer" # Licensed Chld Care
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(sum(is.na(file[["Q1"]][[9]][2])) == 3  ~ "Missing" # Payroll System
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(sum(is.na(file[["Q1"]][[10]][2])) == 3 & sum(file[["S2"]][-(1:14),-(1:3)], na.rm = TRUE) > 0 ~ "Missing" # GBP
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(is.na(file[["Q1"]][[11]][1,2]) & sum(file[["S2"]][22:24, 4], na.rm = TRUE) != 0 ~ "Missing NU Plan" # Pension
                                                    ,is.na(file[["Q1"]][[11]][2,2]) & sum(file[["S2"]][22:24, 5], na.rm = TRUE) != 0 ~ "Missing Managment Plan"
                                                    ,is.na(file[["Q1"]][[11]][3,2]) & sum(file[["S2"]][22:24, 6], na.rm = TRUE) != 0 ~ "Missing BU Plan"
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                          case_when(sum(is.na(file[["Q1"]][[12]])) == 6 ~ "Missing" # Short Term Illness and Injury Plan
                                                    ,TRUE ~ paste0(icons::ionicons("checkmark")))),
                                Schedule = "Q1 Schedule") 
  
  # Q2
  output[["Q2"]] <- data.table(Data = c("Single Site Order", "WorkSafe BC", "Self-Isolation", "Mandatory Vaccination Status Order"),
                               Notes = c(case_when(is.na(file[["Q2"]][1]) ~ "Missing Answer" # SS0
                                                   ,file[["Q2"]][["SSO"]] & nrow(file[["Q2"]][[2]]) == 0 ~ "Missing SSO Table"
                                                   ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                         case_when(is.na(file[["Q2"]][["WorkSafeBC"]]) ~ "Missing Y/N Answer" # WorkSafeBC
                                                   ,file[["Q2"]][["WorkSafeBC"]] == "Y" & is.na(file[["Q2"]][["WorkSafeBC Claims"]]) ~ "Missing # of Claims"
                                                   ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                         case_when(is.na(file[["Q2"]][["Self Isolation"]]) ~ "Missing Y/N Answer" # Self Isolation
                                                   ,file[["Q2"]][["Self Isolation"]] == "Y" & nrow(file[["Q2"]][["Self Isolation Table"]]) == 0 ~ "Missing employee isolation data"
                                                   ,TRUE ~ paste0(icons::ionicons("checkmark"))), 
                                         case_when(is.na(file[["Q2"]][["Mandatory Vaccination Status Order"]]) ~ "Missing Y/N Answer" # Mandatory Vaccination
                                                   ,TRUE ~ paste0(icons::ionicons("checkmark")))),
                               Schedule = "Q2 Schedule") 
  
  # Q3
  output[["Q3"]] <- data.table(Data = "Survey Answers",
                               Notes = case_when(sum(!is.na(file[["Q3"]][,2])) != 12 ~ paste0("Missing ", 12-sum(!is.na(file[["Q3"]][,2])), " Answers")
                                                 ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                               Schedule = "Q3 Schedule")

  
  # Non-Union
  ## N1
  output[["N1"]] <- emp_rev(file, "N1")
  ## N2
  output[["N2"]] <- emp_rev(file, "N2")
  
  # Management
  ## M1
  output[["M1"]] <- emp_rev(file, "M1")
  
  ## M2
  output[["M2"]] <- emp_rev(file, "M2")
  
  # Bargaining Unit
  ## B1
  output[["B1"]] <- emp_rev(file, "B1")
  ## B2
  output[["B2"]] <- emp_rev(file, "B2")

  
  # S1-S2
  ## S1
  ot_h <- as.data.table(!is.na(file[["S1"]]$`Sick and Annual Leave Utilization`[7,2:7])) # IF OT hours not empty -> TRUE
  ot_c <- as.data.table(!is.na(file[["S2"]][4,4:9])) # If OT costs not empty -> TRUE
  w_ot_h <- as.data.table(which(ot_c == TRUE & ot_h == FALSE, arr.ind = TRUE)) # Which OT hours missing
  w_ot_c <- as.data.table(which(ot_c == FALSE & ot_h == TRUE, arr.ind = TRUE)) # Which OT costs missing
  output[["S1"]] <- data.table(Data = c("Regular vs Casual Headcount", "Regional Headcount", "Sick Leave", "Overtime Hours"),
                               Notes = c(case_when(any(colSums(file[["S1"]][[1]][[1]][,-1], na.rm = TRUE) != h) ~ paste("**Headcount differs for:**<br>", # Regular vs. Casual
                                                                                                                        paste(unlist(names(h)[colSums(file[["S1"]][[1]][[1]][,-1], na.rm = TRUE) != h]), collapse = ", ")) 
                                                   ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                         case_when(any(colSums(file[["S1"]][[1]][[2]][,-1], na.rm = TRUE) != h) ~ paste("**Headcount differs for:**<br>", # By Region
                                                                                                                        paste(unlist(names(h)[colSums(file[["S1"]][[1]][[2]][,-1], na.rm = TRUE) != h]), collapse = ", ")) 
                                                   ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                         case_when(all(is.na(file[["S1"]][["Sick and Annual Leave Utilization"]][1:2,-1])==is.na(file[["S1"]][["Total Sick Leave Wage Costs"]][,-1])) ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste("Missing sick leave data for:", 
                                                                 paste(names(h)[which(is.na(file[["S1"]][["Sick and Annual Leave Utilization"]][1:2,-1])!=is.na(file[["S1"]][["Total Sick Leave Wage Costs"]][,-1]), arr.ind = TRUE)[,2]], collapse = ", "))),
                                         case_when(nrow(w_ot_h) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste0("**Missing OT hours for:**<br>",
                                                                  paste(names(ot_h)[w_ot_h$col], collapse = "<br>")))),
                               Schedule = "S1 Schedule")
  ## S2
  cpp_ei_wcb <- cbind(file[["S2"]][10:12,3],
                      setnafill(file[["S2"]][10:12, -(1:3)], fill = 0))
  cols <- names(cpp_ei_wcb)[-1]
  cpp_ei_wcb[, (cols) := lapply(.SD, function(x) x>0), .SDcols = cols] # Check CPP, EI, WCB > 0
  comp <- as.data.table(file[["S2"]][1:5, lapply(.SD, sum, na.rm = TRUE), .SDcols = -(1:3)]>0) # Compensation costs
  comp <- comp[rep(1,3),]
  comp <- data.frame(cbind(cpp_ei_wcb[, 1],
                           as.data.table(cpp_ei_wcb[,-1] == comp)))
  w_comp <- as.data.table(which(comp==FALSE, arr.ind = TRUE)) # Which comp costs missing
  gbp_h <- cbind(file[["N2"]][!is.na(`Benefit Type`), .(`Benefit Type`)], # Check GBP headcounts > 0
                 data.table(PF_NU = file[["N2"]][!is.na(`Benefit Type`)]$Total>0,
                            PF_Mgmnt = file[["M2"]][!is.na(`Benefit Type`)]$Total>0,
                            PF_BU = file[["B2"]][!is.na(`Benefit Type`)]$Total>0))
  gbp_c <- cbind(data.table("Benefit Type" = c("Dental", "Extended Health Care (EHC)", "Long Term Disability (LTD)", # GBP costs
                                               "Pension or Retirement Plan", "Employee & Family Assistance Program (EFAP)")),
                 rbind(setnafill(file[["S2"]][c(14,13,17), 4:6], fill = 0), # If cost is not empty -> TRUE
                       file[["S2"]][20:22, lapply(.SD, sum, na.rm = TRUE), .SDcols = 4:6], # Column Sums for Pension/Retirement Plan
                       setnafill(file[["S2"]][18, 4:6], fill = 0))>0) 
  w_gbp_h <- as.data.table(which(gbp_c==TRUE & gbp_h==FALSE, arr.ind = TRUE)) # Which gbp headcount missing
  w_gbp_c <- as.data.table(which(gbp_c==FALSE & gbp_h==TRUE, arr.ind = TRUE)) # Which gbp cost missing
  output[["S2"]] <- data.table(Data = c("CPP/EI/WCB", "Group Benefits", "Overtime Costs"),
                               Notes = c(case_when(sum(!comp[,-1]) == 0 ~ paste0(icons::ionicons("checkmark")) # CPP/EI/WCB
                                                   ,TRUE ~ paste(case_when(nrow(w_comp[row==1]) > 0 ~ paste0("**Missing CPP for:**<br>", paste0(names(comp)[w_comp[row==1]$col], collapse = "<br>")) ## CPP
                                                                           ,TRUE ~ ""),
                                                                 case_when(nrow(w_comp[row==2]) > 0 ~ paste0("**Missing EI for:**<br>", paste0(names(comp)[w_comp[row==2]$col], collapse = "<br>")) ## EI
                                                                           ,TRUE ~ ""),
                                                                 case_when(nrow(w_comp[row==3]) > 0 ~ paste0("**Missing WCB for:**<br>", paste0(names(comp)[w_comp[row==3]$col], collapse = "<br>")) # WCB
                                                                           ,TRUE ~ ""), 
                                                                 collapse = "<br>",
                                                                 sep = "<br>")),
                                         case_when(nrow(rbind(w_gbp_c, w_gbp_h)) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste(case_when(nrow(w_gbp_h) > 0 ~ paste0("**Missing headcount for:**<br>", 
                                                                                                      paste(names(gbp_h)[w_gbp_h$col], gbp_h[w_gbp_h$row]$`Benefit Type`, sep = " - ", collapse = "<br>"))
                                                                           ,TRUE ~ ""),
                                                                 case_when(nrow(w_gbp_c) >0 ~ paste0("**Missing costs for:**<br>",
                                                                                                     paste(names(gbp_c)[w_gbp_c$col], gbp_c[w_gbp_c$row]$`Benefit Type`, sep = " - ", collapse = "<br>"))
                                                                           ,TRUE ~ ""),
                                                                 collapse = "<br>",
                                                                 sep = "<br>")),
                                         case_when(nrow(w_ot_c) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste0("**Missing OT Costs for:**<br>",
                                                                  paste(names(ot_c)[w_ot_c$col], collapse = "<br>")))),
                               Schedule = "S2 Schedule")
  
  # T1-T4
  ## T1
  T1 <- vector("list", 3)
  names(T1) <- names(file[["T1"]])
  T1[["Avg Time to Fill Vacancies"]] <- c(ifelse((sum(file[["N1"]]$Terminated, na.rm = TRUE) > 0) != any(!is.na(file[["T1"]][[1]][["Non-Union"]]$`Days to Fill`)), "Non-Union", NA),
                                          ifelse((sum(file[["M1"]]$Terminated, na.rm = TRUE) > 0) != !is.na(file[["T1"]][[1]][["Management"]]), "Management", NA),
                                          ifelse((sum(file[["B1"]]$Terminated, na.rm = TRUE) > 0) != any(!is.na(file[["T1"]][[1]][["Bargaining Unit"]]$`Days to Fill`)), "Bargaining Unit", NA))
  T1[["Reasons for Termination"]] <- c(ifelse(sum(file[["N1"]]$Terminated, na.rm = TRUE) != sum(file[["T1"]][[2]][, .SD, .SDcols = patterns("NU")]), "Non-Union", NA),
                                       ifelse(sum(file[["M1"]]$Terminated, na.rm = TRUE) != sum(file[["T1"]][[2]]$Management, na.rm = TRUE), "Management", NA),
                                       ifelse(sum(file[["B1"]]$Terminated, na.rm = TRUE) != sum(file[["T1"]][[2]][, .SD, .SDcols = patterns("BU")]), "Bargaining Unit", NA))
  T1[["Where do terminated employees go?"]] <- c(ifelse(sum(file[["N1"]]$Terminated, na.rm = TRUE) != sum(file[["T1"]][[3]][, .SD, .SDcols = patterns("NU")]), "Non-Union", NA),
                                                 ifelse(sum(file[["M1"]]$Terminated, na.rm = TRUE) != sum(file[["T1"]][[3]]$Management, na.rm = TRUE), "Managament", NA),
                                                 ifelse(sum(file[["B1"]]$Terminated, na.rm = TRUE) != sum(file[["T1"]][[3]][, .SD, .SDcols = patterns("BU")]), "Bargaining Unit", NA))
  
  output[["T1"]] <- data.table(Data = c("Time to fill vacancies", "Reasons for termination", "Where do terminated employees go"),
                               Notes = c(case_when(sum(is.na(T1[["Avg Time to Fill Vacancies"]])) != 3 ~ paste0("**Missing Data for:**<br>", paste(T1[[1]][!is.na(T1[[1]])], collapse = ", "))
                                                   ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                         case_when(sum(is.na(T1[["Reasons for Termination"]])) != 3 ~ paste0("**Missing Data for:**<br>", paste(T1[[2]][!is.na(T1[[2]])], collapse = ", "))
                                                   ,TRUE ~ paste0(icons::ionicons("checkmark"))),
                                         case_when(sum(is.na(T1[["Where do terminated employees go?"]])) != 3 ~ paste0("**Missing Data for:**<br>", paste(T1[[3]][!is.na(T1[[3]])], collapse = ", "))
                                                   ,TRUE ~ paste0(icons::ionicons("checkmark")))),
                               Schedule = "T1 Schedule")
  ## T2
  T2 <- vector("list", 4)
  names(T2) <- names(file[["T2"]])
  T2[["Age"]] <- file[["N1"]][["Classification"]][all(rowSums(file[["T2"]][["Age"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["N1"]][, 15], fill = 0))]
  T2[["Gender"]] <- file[["N1"]][["Classification"]][all(rowSums(file[["T2"]][["Gender"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["N1"]][, 15], fill = 0))]
  T2[["Length of Service"]] <- file[["N1"]][["Classification"]][all(rowSums(file[["T2"]][["Length of Service"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["N1"]][, 15], fill = 0))]
  T2[["Region"]] <- file[["N1"]][["Classification"]][all(rowSums(file[["T2"]][["Region"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["N1"]][, 15], fill = 0))]
  
  output[["T2"]] <- data.table(Data = names(T2),
                               Notes = c(case_when(length(T2[["Age"]][!is.na(T2[["Age"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste0("**Missing Data for:**<br>", 
                                                                  paste(T2[["Age"]][!is.na(T2[["Age"]])], collapse = ", "))),
                                         case_when(length(T2[["Gender"]][!is.na(T2[["Gender"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste("**Missing Data for:**<br>", 
                                                                 paste(T2[["Gender"]][!is.na(T2[["Gender"]])], collapse = ", "))),
                                         case_when(length(T2[["Length of Service"]][!is.na(T2[["Length of Service"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste("**Missing Data for:**<br>", 
                                                                 paste(T2[["Length of Service"]][!is.na(T2[["Length of Service"]])], collapse = ", "))),
                                         case_when(length(T2[["Region"]][!is.na(T2[["Region"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste("**Missing Data for:**<br>", 
                                                                 paste(T2[["Region"]][!is.na(T2[["Region"]])], collapse = ", ")))),
                               Schedule = "T2 Schedule")
  if(file[["N1"]][, .N] == 0) { # Remove NU related schedules if no NU
    output <-  output[names(output) %in% c("N1", "N2", "T4") == FALSE]
  }
  ## T3
  T3 <- vector("list", 4)
  names(T3) <- names(file[["T3"]])
  T3[["Age"]] <- file[["M1"]][["Classification"]][all(rowSums(file[["T3"]][["Age"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["M1"]][, 15], fill = 0))]
  T3[["Gender"]] <- file[["M1"]][["Classification"]][all(rowSums(file[["T3"]][["Gender"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["M1"]][, 15], fill = 0))]
  T3[["Length of Service"]] <- file[["M1"]][["Classification"]][all(rowSums(file[["T3"]][["Length of Service"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["M1"]][, 15], fill = 0))]
  T3[["Region"]] <- file[["M1"]][["Classification"]][all(rowSums(file[["T3"]][["Region"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["M1"]][, 15], fill = 0))]
  
  output[["T3"]] <- data.table(Data = names(T3),
                               Notes = c(case_when(length(T3[["Age"]][!is.na(T3[["Age"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste0("**Missing Data for:**<br>",
                                                                  paste(T3[["Age"]][!is.na(T3[["Age"]])], collapse = ", "))),
                                         case_when(length(T3[["Gender"]][!is.na(T3[["Gender"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste0("**Missing Data for:**<br>", 
                                                                  paste(T3[["Gender"]][!is.na(T3[["Gender"]])], collapse = ", "))),
                                         case_when(length(T3[["Length of Service"]][!is.na(T3[["Length of Service"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste0("**Missing Data for:**<br>", 
                                                                  paste(T3[["Length of Service"]][!is.na(T3[["Length of Service"]])], collapse = ", "))),
                                         case_when(length(T3[["Region"]][!is.na(T3[["Region"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                   ,TRUE ~ paste0("**Missing Data for:**<br>", 
                                                                  paste(T3[["Region"]][!is.na(T3[["Region"]])], collapse = ", ")))),
                               Schedule = "T3 Schedule")
  if(file[["M1"]][, .N] == 0) { # Remove Management related schedules if no Management
    output <-  output[names(output) %in% c("M1", "M2", "T4") == FALSE]
  }
  ## T4
  T4 <- vector("list", 4)
  names(T4) <- names(file[["T4"]])
  T4[["Age"]] <- file[["B1"]][["Classification"]][all(rowSums(file[["T4"]][["Age"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["B1"]][, 15], fill = 0))]
  T4[["Gender"]] <- file[["B1"]][["Classification"]][all(rowSums(file[["T4"]][["Gender"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["B1"]][, 15], fill = 0))]
  T4[["Length of Service"]] <- file[["B1"]][["Classification"]][all(rowSums(file[["T4"]][["Length of Service"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["B1"]][, 15], fill = 0))]
  T4[["Region"]] <- file[["B1"]][["Classification"]][all(rowSums(file[["T4"]][["Region"]][, -(1:3)], na.rm = TRUE) != setnafill(file[["B1"]][, 15], fill = 0))]
  
  output[["T4"]] <- data.table(Data = names(T4),
                               Notes = c(case_when(length(T4[["Age"]][!is.na(T4[["Age"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                     ,TRUE ~ paste0("**Missing Data for:**<br>", 
                                                                    paste(T4[["Age"]][!is.na(T4[["Age"]])], collapse = ", "))),
                                           case_when(length(T4[["Gender"]][!is.na(T4[["Gender"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                     ,TRUE ~ paste0("**Missing Data for:**<br>", 
                                                                    paste(T4[["Gender"]][!is.na(T4[["Gender"]])], collapse = ", "))),
                                           case_when(length(T4[["Length of Service"]][!is.na(T4[["Length of Service"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                     ,TRUE ~ paste0("**Missing Data for:**<br>", 
                                                                    paste(T4[["Length of Service"]][!is.na(T4[["Length of Service"]])], collapse = ", "))),
                                           case_when(length(T4[["Region"]][!is.na(T4[["Region"]])]) == 0 ~ paste0(icons::ionicons("checkmark"))
                                                     ,TRUE ~ paste0("**Missing Data for:**<br>", 
                                                                    paste(T4[["Region"]][!is.na(T4[["Region"]])], collapse = ", ")))),
                               Schedule = "T4 Schedule")
  if(file[["B1"]][, .N] == 0) { # Remove BU related schedules if no BU
    output <-  output[names(output) %in% c("B1", "B2", "T4") == FALSE]
  }
  output[sapply(output, is.null)] <- NULL
  output <- lapply(output, replace_br, "Notes")
  ouptut <- lapply(output, replace_strings)
  return(output)
}

# Format data for Kable output
table_output <- function(file, schedule = NULL) {
  temp <- copy(file)
  if(!is.null(schedule)) {
    if(schedule == "Contact"){
      file[[schedule]][,1] <- paste0("**", file[[schedule]][,1], "**")
      if(as.numeric(str_extract(file[[schedule]][[8,2]], "(?<=\\().+(?=%)")) >= 100) { # Total Funding Utilization > 100%
        kable(file[[schedule]], col.names = NULL) %>%
          kable_minimal(full_width = F) %>%
          row_spec(8, bold = TRUE, color = "white", background = "red") %>%
          column_spec(1, width = "18em", color = "black", background = "white") 
          
      } else {
        kable(file[["Contact"]], col.names = NULL) %>%
          kable_minimal(full_width = F) %>%
          column_spec(1, width = "18em") 
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
              align = "lcccccc") %>%
        pack_rows(index = table(fct_inorder(file[["Home"]]$Fed_Prov))) %>%
        kable_styling("striped")
    } 
  }
  else {
    bind_rows(file[-1]) %>%
      select(-Schedule) %>%
      kable("html",
            align = "lr",
            valign = "c",
            format.args = list(big.mark = ",", digits = 15),
            escape = FALSE) %>%
      pack_rows(index = table(fct_inorder(bind_rows(file[-1])$Schedule))) %>%
      kable_styling("striped", full_width = F) %>% 
      column_spec(1, "30em") 
  }
}

# Convert and format $ values
convert_and_format <- function(dt, format, cols_to_convert) {
  if(format == "$") {
    dt[, (cols_to_convert) := lapply(.SD, function(x) {
      ifelse(is.na(x), "", paste0("$ ", format(round(x, 2), nsmall = 2, big.mark = ",")))
    }), .SDcols = cols_to_convert]
  } else if(format == "%") {
    dt[, (cols_to_convert) := lapply(.SD, function(x) {
      ifelse(is.na(x), "", paste0(format(round(x*100, 2), nsmall = 2, big.mark = ","), "%"))
    }), .SDcols = cols_to_convert]
  }
  
  return(dt)
}

# Remove first <br> in string
replace_br <- function(dt, column_name) {
  dt[, (column_name) := ifelse(grepl("^<br>", get(column_name)), sub("^<br>", "", get(column_name)), get(column_name))]
  return(dt)
}

# Replace Employee group abbrv. with full name
replace_strings <- function(dt) {
  dt[, Notes := gsub("NPF_NU", "NPF Non-Union", Notes)]
  dt[, Notes := gsub("NPF_Mgmnt", "NPF Management", Notes)]
  dt[, Notes := gsub("NPF_BU", "NPF Bargaining Unit", Notes)]
  dt[, Notes := gsub("PF_NU", "PF Non-Union", Notes)]
  dt[, Notes := gsub("PF_Mgmnt", "PF Management", Notes)]
  dt[, Notes := gsub("PF_BU", "PF Bargaining Unit", Notes)]
  
  return(dt)
}
