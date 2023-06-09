# This program is to producing BOM Piping material from MTO File

rm(list = ls())
cat('\014')

# Step 1
# Loading all necessary packages
if(!require(readxl))
  install.packages('readxl')
library(readxl)

if(!require(tidyverse))
  install.packages('tidyverse')
library(tidyverse)

if(!require(plyr))
install.packages('plyr')
library(plyr)

if(!require(stringr))
  install.packages('stringr')
library(stringr)

if(!require(dplyr))
  install.packages('dplyr')
library(dplyr)

if(!require(writexl))
  install.packages('writexl')
library(writexl)

# Step 2
# set working directory and open MTO File
# please make sure you change this folder path otherwise it wont work
setwd("C:\\Users\\akhma\\OneDrive\\Documents\\01. Personal Stuff\\04 Studio Kunio\\00. R in MTO\\Test")
# Open excel file with specified Sheet and Colomns
Sheet1 <- read_xlsx("Test-MTO.xlsx","Test",range = cell_cols("D:G"))
# To remove QTY = NA
Sheet1 <- subset(Sheet1, !QTY=="NA")

# Step 3 
# the main code of ths program
# Processing the Piping MTO by adding up all the materials that have same description, material and size
MTO_piping_df <- tbl_df(Sheet1)
MTO_piping_df <- select(MTO_piping_df,DESC, MATERIAL, SIZE,QTY)
MTO_piping_df <- aggregate(MTO_piping_df$QTY, by=list(DESC=MTO_piping_df$DESC,MATERIAL=MTO_piping_df$MATERIAL,
                                                      SIZE=MTO_piping_df$SIZE), FUN=sum)
# Change the header name of X into the Quantity
names(MTO_piping_df)[names(MTO_piping_df) == "x"] <- "QTY"
# Arrange the table with the right order
MTO_piping_df<- arrange(MTO_piping_df, DESC, MATERIAL, SIZE, QTY)

# Step 4
# Filtering the piping material into different sheet
# Valve Sheet
Valve <- MTO_piping_df %>% filter(str_detect(DESC, 'Valve'))
Valve <- Valve %>%
  add_column(NO = seq(1, nrow(Valve), length.out = dim(Valve)[1]),
             .before = "DESC") 
Valve_df <- tbl_df(Valve)

# Pipe Sheet
Pipe <- MTO_piping_df %>% filter(str_detect(DESC, 'Pipe'))
Pipe <- Pipe %>%
  add_column(NO = seq(1, nrow(Pipe), length.out = dim(Pipe)[1]),
             .before = "DESC") 
Pipe_df <- tbl_df(Pipe)

# Flange Sheet
Flange <- MTO_piping_df %>% filter(str_detect(DESC, 'Flange'))
Flange <- Flange %>%
  add_column(NO = seq(1, nrow(Flange), length.out = dim(Flange)[1]),
             .before = "DESC") 
Flange_df <- tbl_df(Flange)

# Gasket Sheet
Gasket <- MTO_piping_df  %>% filter(str_detect(DESC, 'Gasket'))
Gasket <- Gasket %>%
  add_column(NO = seq(1, nrow(Gasket), length.out = dim(Gasket)[1]),
             .before = "DESC") 
Gasket_df <- tbl_df(Gasket)

# Bolt and Nut Sheet
Bolt <- MTO_piping_df %>% filter(str_detect(DESC, 'Bolt and Nut'))
Bolt <- Bolt %>%
  add_column(NO = seq(1, nrow(Bolt), length.out = dim(Bolt)[1]),
             .before = "DESC") 
Bolt_df <- tbl_df(Bolt)

# Fitting Sheet
Fitting <- MTO_piping_df %>% filter(str_detect(DESC, 'Valve|Pipe|Flange|Gasket|Bolt and Nut', negate=TRUE))
Fitting <- Fitting %>%
  add_column(NO = seq(1, nrow(Fitting), length.out = dim(Fitting)[1]),
             .before = "DESC") 
Fitting_df <- tbl_df(Fitting)

# Compiled sheet
MTO_piping_df <- MTO_piping_df %>%
  add_column(NO = seq(1, nrow(MTO_piping_df), length.out = dim(MTO_piping_df)[1]),
             .before = "DESC") 
MTO_piping_df <- tbl_df(MTO_piping_df)


# Step  5
# Write excel file with the all separated piping items
require(openxlsx)
list_of_datasets <- list("All Piping" = MTO_piping_df, "Valve" = Valve_df,"Pipe" = Pipe_df,
                         "Flange" = Flange_df,"Gasket" = Gasket_df,"Bolt and Nut" = Bolt_df,
                         "Fitting" = Fitting_df)
write.xlsx(list_of_datasets, file = "BOM_Piping_Summary.xlsx")


