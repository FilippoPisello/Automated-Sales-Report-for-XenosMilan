# AIM OF THE PROJECT
Script to automatize the reporting system of a small clothing brand. The program
takes input from the company's website in Big Cartel and outputs a pdf file
containing graphs, text and tables.

Note that the script DOES NOT WORK in its GitHub version because it would rely
on the website's credentials.

The main file is "Xenos_Report.py", everything else is aimed at its functioning.

#LOGIC OF THE PROGRAM
- Log into the website through WEBBOT and download a .csv containing the sales'
records
- Find the file into the folder where it gets automatically downloaded and move
it into the resource folder
- Open the .csv in pandas
    - Clean up the table
        - Remove useless information
        - Extract data of interest
            - Gender based on name, from external file
            - Uniformed city, province and region names, from external file
            - Size from product name
            - Time of item, from external file
    - Match the columns with the names used throughout the program
    - Perform the analysis
        - Create desired tables
        - Create desired graphs
- Save the contents of interest
- Open a template in -docx and create the desired styles for the final document
using python-docx
- Write into the document and insert graphs
- Save a .docx output
- Transform the .docx output into a .pdf
- Move back the .pdf into the main folder
