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

**Functions**
- AppendRow(): Append the specified row to the section duplic ates list, duplicating by sections as specified above (You must trigger this manually and specify the row in the Row Variable)
- AppendNewestRow(): Append the newest row to the end of the section duplicates list, duplicating by sections as specified above. (Manually triggered in Google Apps Script Editor/Terminal)
- AppendAllRows(): Run the duplication process as specified above on the entire table (Can be manually triggered, but is also linked to a trigger to automatically run every time the sheet is opened)

**Triggers**
As of now, applications are collected through Qualtrics and logged into the Google Sheet this script is attached to. Due to the varying levels of access required due to these scripts being maintained by Peer Educators, at the time of writing, a solution to trigger functions everytime a new entry comes in has not yet been develped. 

There is talk about changing the system, but for the time being, the code has a trigger set to run AppendAllRows() whenever someone opens the sheet. That way, every time the sheet is opened, any new entries in the main sheet will be automatically ended.

*NOTE*: For the time being, the code is run on my (Jaime Elepano) Google Account. Should ownership of the Google Apps Script Project be transferred to your account, every time the code executes (ie. every time someone opens the sheet), your Google Account will show up in the edit history, as the automated actions are run from your gmail account. Exercise caution as you are handling sensitive data.