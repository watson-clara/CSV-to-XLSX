# CSV to XLSX

This python program takes CSV files from one directory and copies their content, creates new XLSX files in a new directory, and populates it with the data from the CSV file. 

## Description
First the program starts by looping through the directory that contains your CSV files. As it loops, it gets the path for each individual CSV file as a string. The program uses this path to create a new path to the directory in which you want to save the new XLSX files. Next, the program opens each CSV files and copies the data of each row into an array. It then takes the path you created for your new XLSX file and opens it as an excel workbook. It then populates the new excel workbook with the array that contains the data from the CSV file. Lastly, there is a function to delete all the origional CSV files from their directory.  
