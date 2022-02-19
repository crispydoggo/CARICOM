



# CLEAR LAC
# Spring 2022
# CARICOM RBM Policy

# ==========================
# Packages 
# ==========================
library(haven)
library(readr)
library(tidyverse)
library(XLConnect) #Writes Excel Sheets for easy reading
library(readxl)

# ==========================
# Data
# ==========================

CARICOM <-  read_delim("CARICOM.txt", delim = "=")
RBM_ideal <- read_csv("RBM_ideal.csv")

# Correct some features of data

CARICOM <- CARICOM %>% 
  select(bullet_id = `bullet_id `, comment = ` comment` )%>% 
  mutate( bullet_id = as.integer(bullet_id))


# ==========================
# Functions 
# ==========================


sort_rbm <- function(CARICOM){
  # Sorts the comments by bullet id and exports them into an excel file
  # ----------------------------------
  
  book <- loadWorkbook("RBM_com.xlsx")
  RBM <- tibble()
  
  for( i in 1:19){
    RBM <- CARICOM %>% 
      filter( bullet_id == i)
    createSheet(book, name = paste("bullet_", i, sep = ""))
    writeWorksheet(book, RBM, sheet = paste("bullet_", i, sep = ""))
  }
  saveWorkbook(book, file = "RBM_com_clean2.xlsx")
}

add_rbm <- function(file_name){
  
  RBM_info <- read_excel(file_name, sheet = 1)
  RBM_info <- arrange(RBM_info, sub_id)
  for( i in 2:19){
    
    aux <- read_excel(file_name, sheet = i)
    aux <- arrange(aux, sub_id)
    RBM_info <- bind_rows(RBM_info, aux)
    
  }
  
  RBM_info <- RBM_info %>% 
    select( bullet_id, sub_id, comment) %>% 
    mutate(bullet_id = as.integer(bullet_id), sub_id = as.integer(sub_id)) 
  
  return(RBM_info)
  
}

# ==========================
# Data Cleaning
# ==========================


# Arranges the raw Comment section 
sort_rbm(CARICOM)

# Process the ordered Comment Section
RBM_info <- add_rbm("RBM_com_clean_Jamaica.xlsx")


# ==========================
# Export Data
# ==========================


book <- loadWorkbook("RBM_Dominica.xlsx")
writeWorksheet(book, RBM_info, sheet = 1)
saveWorkbook(book, file = "RBM_Jamaica.xlsx")
