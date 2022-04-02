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




6. Finally, I ran the stock analysis, to confirm that my stock analysis outputs for 2017 and 2018 are the same as the example provided.


 - Below our Final VBA Analysis PNGs,


***Final VBA Analysis 2017***

![name-of-you-image](https://github.com/emmanuelmartinezs/stock-analysis/blob/master/data_files/resources/VBA_Challenge_2017.PNG?raw=true)


***Final VBA Analysis 2018***

![name-of-you-image](https://github.com/emmanuelmartinezs/stock-analysis/blob/master/data_files/resources/VBA_Challenge_2018.PNG?raw=true)


**8. The pop-up messages showing the elapsed run time for the script are saved as `VBA_Challenge_2017.png` and `VBA_Challenge_2018.png`**

> Running our fully 2017 and 2018 data stock analysis gave us an elapsed run time for each year, below our results.


***Time on VBA_Challenge_2017.PNG***

![name-of-you-image](https://github.com/emmanuelmartinezs/stock-analysis/blob/master/data_files/resources/Time%20for%202017%20analysis.PNG?raw=true)


***Time on VBA_Challenge_2018.PNG***

![name-of-you-image](https://github.com/emmanuelmartinezs/stock-analysis/blob/master/data_files/resources/Time%20for%202018%20analysis.PNG?raw=true)



## SUMMARY: Our Statement:

### Deliverable with detail analysis:
**1. What are the advantages or disadvantages of refactoring code?**

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

**Disadvantages:**

> - A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
> - A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
> - A complex unstructured code is usually best to split in several functions. 
> - Refactoring process can affect the testing outcomes. 


**Advantages:**
> - Logical errors easily appear in well structure code that contains nested conditionals and loops. 
> - In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
> - VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.

**2. How do these pros and cons apply to refactoring the original VBA script?**

> Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring.
Now, let's think about something, **What happens after a couple of days or months yo need to troubleshoot your code? Is it complicated? Is it hard to understand?** If yes then definitely you didn’t pay attention to improve your code or to restructure your code. 

***We need to consider the code refactoring process as cleaning up the orderly house.*** 
*Unnecessary clutter in a home can create a chaotic and stressful environment.* - The same goes for written code. 

A clean and well-organized code is always easy to change, easy to understand, and easy to maintain. You can avoid facing difficulty later if you pay attention to the code refactoring process earlier.



