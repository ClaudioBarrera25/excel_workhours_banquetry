# Purpose of the work

In this code it was requested to take all the xls files and select the 'Registro asistencia' sheet and fill up the days that are empty or that were filled with just one hour instead of two or three. 

For this, it was generated a random time function, and the files available were read using string formatting by replacing each month's date. 

The workbook is loaded and then some key info is retrieved like employees names, cells locations of times and names. 

Then with a numpy array is iterated for each column to check if it has all values None, if that's the case, then the day wasn't worked. On the contrary, if there is at least one time on record, then it's selected as workday. 

Then, we write on the file the new times and the workbook is saved on the new location.
