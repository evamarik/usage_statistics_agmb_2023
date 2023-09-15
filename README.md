# Usage statistics AGMB 2023
Background information for a poster presented at [AGMB 2023](https://agmb.de/de_DE/2023-bonn-startseite): "Bestandsnutzung an der neuen Universitätsmedizin Augsburg: Erste Einblicke".

All the files the code reads must be in the same ordner as the code.


### Information on the code for plotting 'Bestandsliste 172':
#### prepare_172_excel_file:
This code edits a Counter5 Report file that contains stock information on books with a pressmark starting with 172. 
The data is grouped by 'ISBN' and some columns are removed.
The new file contains the following columns: 
- Signatur
- ISBN
- Titel
- Titelzusatz
- Beilagen
- Verfasser / Urheber	
- Verlag	
- Ort	
- Jahr
- Ausleihzähler gesamt	
- Ausleihzähler lf. Jahr
- Ausleihzähler Vorjahr
- Ausleihzähler Vorvorjahr	
- Ausleihzähler aller Jahre vor Vorvorjahr

    
#### plot_172_for_poster:
This code contains multiple functions that can be used to plot the data from the previously updated file (prepare_172_excel_file):
First the code reads the updated file and sorts it based on the usage per year (the order for each year is saved separately and can be used as input for the functions).
##### The functions for plotting:
- plot_top_of_year plots the most used books of a certain year as horizontal bar plot.
- plot_top_10_2022_shorttitle is a very specific version of the function above. It's for plotting the top tens of 2022 with shorttitles instead of the book title 
(those shorttitles were manually added to the updated and soreted file). 
- plot_top_over_years creates a bar plot, that shows the percentage of the top books within of the books that are lend overall.
- plot_preclinic_over_years does the same as the function above just for preclinic books instead of the tops.
