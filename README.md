# Geo-App

VBA Excel program: generates a report in Word using information from Database. Utilizes userforms, arrays, user input, spreadsheet data reading, MS Access database reading (in SQL), error handling, custom ribbon buttons, and generating reports in MS Word.

## Instructions

Download geo-app.xlsm and geo.accdb and place them in the same folder (or filepath). Unblock macros of the xlsm file by going to properties and checking the box. Open the xlsm file and go to the geo-app tab at the far right in the ribbon. Select Open UI, from there select a country and generate the report.

## Other Notes

Many countries are missing from the list yet present in the database. This is due to 2 reasons.
- Some fields for longitude and latitude are empty causing problems
- The second is that some countries very clearly had inaccurate data. For example, there was this one small country in Asia with all its provinces reasonably together, but one was in the middle of the Atlantic. This was easily discoverable via the scatter plot (easier to visualize too). After searching to see the actual location of that province, it was very much with in its borders making that data incorrect, hence I decided to remove countries with obviously visible outliers apart from countries with colonies

The measurement of the region coordinates is not clear as to if they’ve been measured at the centre or edge of the regions. This can create some confusion in the furthest in each direction section of the report although it functions just fine 

The scatter plot is rather interesting as for smaller more dense countries, the shape of the points almost resembles the borders (e.g. turkey). Its also interesting to see the many offshore bases or colonies some countries have; it makes the scatter plot of the Unites States look like an archipelago 
