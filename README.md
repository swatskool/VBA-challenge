			***************************************************
			     		The VBA of Wall Street
			***************************************************

1. The macro in the module is named stock_analysis.
2. It can be attached to a button on the first sheet or run as it is. 
3. Ticker name column and date column have been sorted in the script
4 . Corresponding to each ticker name, using for loop, all the stocks are iterated. For each ticker name the opening value on the lowest date and the closing value on the last date are stored along with cumulative volume.
5. The yearly change and percentage change for the each ticker is calculated and outputed along with total volume.

BONUS

6. Another loop runs through this newly analyzed data and tries to calculate the lowsest percentage change, the highest percentage change and the highest volume and saves the corresponding ticker name too.
7. These values are then reflected as the summary on the corner of the table

**Warnings**
- Incase there is any Ticker with opening and closing values as zero, it will reflect in the percentage change cell as the string "Error", but would not hinder the execution of the script.



			***************************************************
					Swati Oberoi Dham
			***************************************************
