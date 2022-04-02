# STOCK ANALYSIS 

## OVERVIEW: VBA Stock Analysis 

### Purpose
In this project and analyisis, I’ll edit/refactor the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire information. Then, I’ll determine if refactoring my code successfully made the VBA script run faster. Finally, I'll present a written analysis that will explain my findings. 


## RESULTS: Code descriptions with screenshots 
 
### Deliverable Requirements
1. Created and set a TickerIndex variable to equal zero before iterating over all the rows. This TickerIndex is used to access the correct index across the four different arrays on VBA Code created in the next next.


2. Created Arrays are for Tickers, TickerVolumes, TickerStartingPrices, and TickerEndingPrices.


3. The TickerIndex is used to access the stock ticker index for the Tickers, TickerVolumes, TickerStartingPrices, and TickerEndingPrices arrays.

       **Below is the screnshot exhibiting the codes in step 1-3 above.**

![name-of-you-image](https://github.com/QIhunwoKingsley/stock-analysis/blob/1cb284076ca2599875baaab943732d805d35df38/Resources/Requirements%201-3.png)


4. Created a loop that will loop over all the rows in the spreadsheet.
 - Inside the loop, I created a script that increases the current TickerVolumes (stock ticker volume) variable and then, I added the ticker volume for the current stock ticker.


**Stored values from** TickerStartingPrices **and** TickerEndingPrices

 - Created an **if-then** statement to check if the current row is the first row with the selected TickerIndex. If yes, I assigned the current closing price to the TickerStartingPrices and TickerEndingPrices variable.


![name-of-you-image](https://github.com/QIhunwoKingsley/stock-analysis/blob/9ef4f64675db26cf5bb069ce8e40fd04aca1cd2f/Resources/Requirements%204.png)



5. To check if the code for formatting the cells in the spreadsheet is working - I made positive returns green and negative returns red to see which stocks did well and which ones didn't.

![name-of-you-image](https://github.com/QIhunwoKingsley/stock-analysis/blob/dc00e0945d88563c4e2f2cdb1f9aad31a5ddffee/Resources/Requirements%205.1.png)

![name-of-you-image](https://github.com/QIhunwoKingsley/stock-analysis/blob/85b143411fba9548f9b62d84f2195d0f519b642a/Resources/Requirements%205.2.png)




6. Finally, I ran the stock analysis to confirm that my stock analysis outputs for 2017 and 2018 are the same as the example provided.


 - Below our Final VBA Analysis PNGs,


***Final VBA Analysis 2017***

![name-of-you-image](https://github.com/QIhunwoKingsley/stock-analysis/blob/ac33e2da8512d79abf6e4c9dfebb2df3235904a9/Resources/All_Stock_2017.png)


***Final VBA Analysis 2018***

![name-of-you-image](https://github.com/QIhunwoKingsley/stock-analysis/blob/43cdc5ce3f55454b0942e213d13b8d0a2933bb4c/Resources/All_Stock_2018.png)


7. The pop-up messages showing the elapsed run time for my script.

 - I ran my 2017 and 2018 data stock analysis which gave me elapsed run times for each year.


***Time on VBA_Challenge_2017.PNG***

![name-of-you-image](https://github.com/QIhunwoKingsley/stock-analysis/blob/afcfbfeda001a4ee0b503a3f77d19a38da08c69b/Resources/VBA_Challenge_2017.png)


***Time on VBA_Challenge_2018.PNG***

![name-of-you-image](https://github.com/QIhunwoKingsley/stock-analysis/blob/ae99f442e2ee18e9d3ad51179f7c792783849e09/Resources/VBA_Challenge_2018.png)



## SUMMARY:

**Advantages:**
- Logical errors easily appear in well structure code that contains nested conditionals and loops. 
- VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.


**Disadvantages:**

- A complex unstructured code is usually best to split in several functions. 
- Refactoring process can affect the testing outcomes. 
