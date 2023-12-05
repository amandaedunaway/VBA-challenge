# VBA-challenge


Comments about my code:
The sub procedure scans stock data to pull designated values and perform calculations. I noticed after I took my screenshots that my values for the second set of calculations are not correct. This means there is a problem either with my min and max functions or the subsequent loop that pulls the associated ticker.

To improve the code, I would also need to get the LastRow function working properly (currently commented out in my code). It worked for my alphabetical file, but not my multiple year file. As an intermediate solution, I manually placed the final row number of the 2018 data. However, this is not a true solution since the other years do not have the same amount of rows as 2018, but when I attempt to use a different year’s row total, my excel crashes.

These two shortcomings highlight the importance of being detail-oriented and doing triple checks beyond the double checks to ensure everything is correct and working properly. 



Submission comments:
I received an error message that my file is too large to upload to Github. Zipping the file didn’t work, and I was not able to determine how to save a VBA script file as a separate entity from a macro-enabled excel workbook. To make sure I submit all relevant files, I will also share a Google Drive link on my BCS submission, and I can figure out how to submit appropriately sized files for my next challenge assignment in python.
Here is the Google Drive link to the excel file:
https://drive.google.com/drive/folders/1SIE6wpzSiBfYaAg4y10h-vIhrO95Zzjz?usp=drive_link


Sources:
My tutoring session with Justin Moore, who helped me take my script from partially working to up and running. He helped me by pointing out that I need to change the way I’m tracking Total Volume and a much simpler way to display the tickers with DisplayRow.

I acquired the code to loop through the sheets from the following website.
https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html

