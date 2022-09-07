# countexcelaverage
counts average from rows in excel

This script is written to calculate the average of 899 excel rows.

#Problem
Had around 100 excel files, where I have to open them individually and calculate the average of the cells and take the value to proceed.

#Solution
Wrote this python script, which will automatically iterate to all the excel files in the folder and automatically calculate the average of the cells and write the value in a new cell within the same sheet.
 ==> So, I can now just open those excel files and copy the value directly and close the sheet.
 
 #Further Improvement?
This script is very basic. There are a few improvements that can be made such as:
--> Obtain the avg value and copy it to a new excel sheet. So, that after the loop has ended, all the average values of each excel workbook are written in one single excel. So, have to open only one excel workbook to obtain the values.
