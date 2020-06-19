#V 1.0

## AIM OF THE PROJECT
Script to automatize the reporting system of a small clothing brand. The program
takes input from the company's website in Big Cartel and outputs a pdf file
containing graphs, text and tables.

Note that the script DOES NOT WORK in its GitHub version because it would rely
on the website's credentials.

The main file is "Xenos_Report.py", everything else is aimed at its functioning.

## LOGIC OF THE PROGRAM
1. Log into the website through webbot and download a .csv containing the sales'
records
1. Extract from the external file "Report_path.txt" the automatic location where
the file gets downloaded.
1. Move the .csv into the resource folder
1. Run the analysis with pandas
    1. Open the .csv
    1. Clean up the table
        - Remove useless information
        - Extract data of interest
            - Gender based on name, from external file ("Names_list.csv")
            - Size from product name
    1. Match the columns with the names used throughout the program
    1. Perform the analysis
        - Create desired tables
        - Create desired graphs
        - Save the contents of interest
1. Open a template in -docx and create the desired styles for the final document
using python-docx
1. Write into the document and insert graphs and save a .docx output
1. Transform the .docx output into a .pdf
1. Move back the .pdf into the main folder

## TO COME, V 2.0
In section 4.2
- Uniformed city, province and region names, from external file
- Type of item, from external file
In section 4.4
- Extract and format tables through openpyxl
- Add them to the report
