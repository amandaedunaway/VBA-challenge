# VBA-challenge
## Module 2 Challenge

### Outline
The sub procedure scans stock data to pull designated values and perform calculations. The yearly change in stock price, percent change in stock price, and average stock volume are calculated for each ticker. Then, the script scans these calculations to determine the tickers with the greatest percent increase, greatest percent decrease, and the greatest total volume.


### Submission comments
This is my third submission for this challenge. For this submission, I updated the format for the values in the columns with percent changes to include percent symbols.


### Sources
First, I received help during my tutoring session with Justin Moore, who helped me take my script from partially working to up and running. He helped me by pointing out that I need to change the way I’m tracking Total Volume and a much simpler way to display the tickers with DisplayRow. Next, I acquired the code to loop through the sheets from [an article](https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html). Then, I determined how to implement the "LastRow" in the Range function in the search to find the greatest percent increases and decreases by consulting [ChatGPT](https://www.openai.com/). This resolved the issue from the first submission having a manually-entered last row, which produced incorrect results in the table. Finally, I consulted [stack overflow](https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage) to correctly change the format in the columns with percent changes to include a percent symbol.
