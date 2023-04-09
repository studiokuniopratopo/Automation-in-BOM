# Automaton-in-BOM
This folder contains of the sample input MTO file, the R program for producing the Bill of Material (BOM) sheets and the result file in the excel file.
# Input File
The input MTO file in this sample is Test-MTO.xlsx. 
There are the informations of No, Line Number, Material, Description, Size and Quantity
The main aim of this code is to summarize the quantity of the all piping descriptions that have same material and size.

# The code
The code is written in R, an open souce program
The name of the R code is Processing_Piping 
There are 5 main steps in this code to make a BOM File:
Step 1: Install all the required packages
Step 2: Set Working directory, open & clean the data in the MTO File
Step 3: The main processing including filtering, mutate, groupping and summarizing
Step 4: Separating the piping material into different sheets
Step 5: Write an excel file with all the separated sheets that got from the previous step

# Output File
The output file BOM is BOM_Piping_Summary.xlsx
There are seven different sheets based on the type of the piping description namely; pipe, fitting, flange, gasket, valve, Bolt & Nut, and All Piping
Once the BOM file is ready then you can send it to the vendor
