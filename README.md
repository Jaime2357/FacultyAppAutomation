***Developer Documentation***

**Overview**
This is script is used to automate the organization of data for Peer Connections Faculty Requests. 
It's main purpose (at the time of writing this) is to read a list of Faculty Peer Educator Requests 
and copy each line to a new sheet. 

The need for automation is due to two things:
- Each entry has up to four *course* requests
- Each course request (up to four) can have multiple sections

We want these requests divided into separate entries such that each entry is for a single section of a single course.
The purpose of the script is to copy each entry into a new sheet and duplicate each row based on courses and sections. 

For example: If Prof. Smith sends a request for Math 1 and Phys 50, which have 5 and 2 sections respectively, we would be inserting
7 new rows into the duplicates table. Ideally we'd also be filtering out data based on that section. So for the 5 Math 1 sections, the first entry would have the section specific info for just that section.

Unless a large scale change is requested, your only work will be to test the script to ensure it works as intended. The current version hardcodes certain column values. As the source sheet may change in structure, your main job in terms of maintenance is to ensure all the hardcoded values are accurate.

**Main Functions**
- AppendSingleRow(): Append the specified row to the section duplicates list, duplicating by sections as specified above (You must trigger this manually and specify the row in the Row Variable)
- AppendNewRows(): Append all rows that don't have the "Duplicated" column (usually the last column) checked. (Can be manually triggered, but is also linked to a trigger to automatically run every time the sheet is opened)
- AppendAllRows(): Run the duplication process as specified above on the entire table indiscriminately (Should be manually run in the event the duplicates table needs many or all its entries updated)

**Triggers**
As of now, applications are collected through Qualtrics and logged into the Google Sheet this script is attached to. Due to the varying levels of access required due to these scripts being maintained by Peer Educators, at the time of writing, a solution to trigger functions everytime a new entry comes in has not yet been develped. 

There is talk about changing the system, but for the time being, the code has a trigger set to run AppendAllRows() whenever someone opens the sheet. That way, every time the sheet is opened, any new entries in the main sheet will be automatically ended.

*NOTE*: For the time being, the code is run on my (Jaime Elepano) Google Account. Should ownership of the Google Apps Script Project be transferred to your account, every time the code executes (ie. every time someone opens the sheet), your Google Account will show up in the edit history, as the automated actions are run from your gmail account. Exercise caution as you are handling sensitive data.

**Underlying Logic**

*Notes*
Columns are indexed by number, not letter. I've written a helper function ```colNum(<column letter>)``` to simplify the process of specifying columns

*Duplicating a Single Row*
1. Checking if the specific row has already been duplicated. There should be a checkbox in the very last cell of each row in the sourcesheet to represent this. If it isn't already, please label this column for ease of comprehension. The variable ```dCheckColumn``` stores this column's number in the code. Make sure to hardcode this as needed with each iteration of the sheet.
2. Next we want to know how many section requests there are. The utility function ```setRequests(SpecificRow)``` does this. Just specify the row as its parameter. Once run, the request count is stored in a variable
3. Generate the ranges of included values As mentioned before, for each section's entry, there is no need to specify all the other sections' information. The helper function ```getRanges(i)``` does all the work. It takes the current i (indicating what course you're making duplicates of) as a parameter.
4. Next Build the row. The ```buildRow(row, range)``` function does all the work. You'll need to specify the current row and included range as the parameters. It iterates through each column index and checks if its in the range specified in its parameter, then appends that value to an array. The returned array is stored in the variable ```builtRow```.
5. Finally, use the built array to append a the new row to the duplicates table. The function ```appendTable(builtRow, sectionCount)``` does all thework. You'll need to push the built arraw and the number of sections as parameters respectively. If you're wondering about the getRange() call in the function, here's the call structure ```getRange(startRow, startColumn, optionalNumberOfRows, optionalNumberOfColumns)```

*Duplicating New Rows/All Rows*
The logic builds on that of duplicating a single row.

1. Since you'll be duplicating multiple rows, you'll want to use the function ```filledCount()``` to get a count of filled rows.
2. Essentially all you have to do now is steps 2 - 5 of duplicating a single row, but for each new entry. Simply put the same code in a loop. If you're only duplicating new row, within the loop, create a check for if that row has already been duplicated. (We could simply look for the first value without the duplication check mark since theoretically all those after would be new as well, but we're iterating over the whole list to be safe.). 

Note: The code originally had the code rewritten for each function due to the developer (Jaime) being lazy, plus the ```appendRow()``` function had to be runnable without parameters so that less tech savvy users could run it standalone, meaning it couldn't be used as a driver function. Now the code has been restructured and ```appendRow()``` is a driver function while ```appendSingleRow()``` is the one executed manually.

3. Rememeber after each loop to mark the row as duplicated. The code should already have a line to do this.

For any questions contact Jaime Elepano: (jaimeelepano2357@gmail.com)