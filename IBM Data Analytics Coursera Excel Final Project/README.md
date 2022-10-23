# Microsoft Excel Final Project by IBM Data Analytics Certificate

### Overview 

After learning basics and advanced Excel, the IBM Data Analytics Certificate program had a graded assignment that was issued out to test a student's understanding of the Excel functionalities needed for Analysis.
### The Assignment

The Assignment had two sections:

1. The first section dealt with data cleaning and preparation
2. The other part dealt with analysis of the data


#### The first assignment

Data source : https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/IBMDeveloperSkillsNetwork-DA0130EN-SkillsNetwork/Hands-on%20Labs/Peer%20Graded%20Assignment%20-%20Part%201/Montgomery_Fleet_Equipment_Inventory_FA_PART_1_START.csv

1. **Save the CSV file as an XLSX file:** Click the 'Edit Workbook' button in the ToolTip to save the file as an XLSX file. The file is converted when you click 'OK' in the prompt.

2. **Column widths:** Sort out the widths of all columns so that the data is clearly visible in all cells.

3. **Empty rows:** Use the Filter feature to look for blanks and remove all empty rows from the data.

4. **Duplicate records:** Use either the Conditional Formatting or Remove Duplicates feature to look for and remove any duplicated records from the data.

5. **Spelling:** The original source file data has not been checked for errors in the spelling. Check for spelling mistakes in the data and fix them.

6. **Whitespace:** Use the Find and Replace feature to remove all double-spaces from the data.

7. **Department names:** When the data was converted from its data source, the department names (see correct list below) didn’t import correctly and they are now split over two columns in the data. Use Flash Fill to reduce the department names to just one column, and then remove any unnecessary columns.

![a1](https://user-images.githubusercontent.com/74729228/197416790-38dbae48-016d-49c8-8583-fa197da0afcc.png)

8. **Save your workbook:** Use 'Save As' to save your completed workbook as **Montgomery_Fleet_Equipment_Inventory_FA_PART_1_END.XLSX**


#### The second assignment

Data source : 

1. **Format the data as a table**: Use the Format as Table option to format the data as a table.

2. **Use AutoSum to calculate values:** Use AutoSum to find the following values for column 'C' and record each of the values:

   - SUM
   - AVERAGE
   - MIN
   - MAX
   - COUNT

3. **Create a Pivot Table:** Use the PivotTable feature to create a pivot table that displays the Department field in the Rows section, and the Equipment Count in the Values section, so that the pivot table displays the sum of equipment count by department.

4. **Sort the pivot table data:** Use the Sort By Value setting on the pivot table to sort it in descending order by the sum of equipment count.

5. **Make two more pivot tables exactly the same as task 3:** Follow the same steps you performed in Tasks 3 and 4 to create two more identical pivot tables so that you end up with 3 worksheets that contain identical pivot tables.

6. **Analyze data in the pivot table:** Use the PivotTable Fields pane to manipulate and analyze data in the two copied pivot table as follows:

     - In pivot table 2 add the Equipment Class field below the Department field so that the different vehicle types appear under each department with their respective counts.
     - Collapse all fields except the top one - **Transportation**
     - In pivot table 3 add the Equipment Class field above the Department field so that the different vehicle types appear first, with the different departments listed underneath each vehicle type with their respective counts.
     - Collapse all fields except the top one - **CUV**

7. **Save your workbook:** Use 'Save As' to save your workbook as **Montgomery_Fleet_Equipment_Inventory_FA_PART_2_END.XLSX**
