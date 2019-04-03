# Excel-ADC-value-Conversion-Makro
A Makro that converts a column of values into a row of values in the form of x,x,x, which can be easily implememted as array or LUT.

How to Use:

1. Define where the output shall be stored in the variable "strFileName".
2. Define the name of your worksheet.
3. Define the row in which your data is stored
    a = 1
    b = 2
    .
    .
    .
4. Define the number of deimal places you want to cut the value at.
    ATTENTION: Define a number Smaller then the smallest no. of decimal places!
5. Define the range of columns in which your data is stored
5. OPTIONAL
    Comment the replace section out if you have no need to transform , to .
    e.g. 1,205 to 1.205
