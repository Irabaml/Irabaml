The challenge's purpose is to calculate the Quarterly change price, percent change, and total volume after the system has displayed the greatest percentage increase, decrease, and total volume. In addition, the system codes must work through each sheet.
My module was named stockdata1(), and after I declared and initialized the necessary variables that I used in this challenge.
For example, I defined the following variables such as:
•	Ws for worksheet
•	Count for the row 
•	Total volume for the sum of the volume
•	Last_row for the last row that contains the data
•	First_row to help me increment the row
•	Startcell to order where the system starts to count in a manner you can know the ranges that are needed (defined as A1)
•	Col1 to pull the information concerning the first open value for each ticker name in the column
•	Col2 to pull the information concerning the last close value for each ticker name in the column
I created the headers manually by using the indexes of cells 
I defined the last row that contains the data, and with it, I managed to create the loop that helped me to analyze the condition and pull the requirement information _ please find the details in the comments in the codes.
I used different functions and objects such as:
•	WorksheetFunction.Max for determining the maximum number
•	WorksheetFunction.Min for determining the minimum number
•	EntireColumn.Hidden hiding the entire columns that were not needed but the columns helped me to find the required calculation
•	WorksheetFunction.Match for finding the positions of the values within the ranges
•	Interior.Color for terminating the colors needed
•	offset for pointing the exact cells that I needed to use to operate

I Separated VBA script files using the Export File option and saved the file as a .bas extension.

