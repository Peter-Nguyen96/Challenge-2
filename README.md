# Challenge-2
VBA Scripting Challenge: Green Stocks Analysis

Project overview
A newly appointed financial analyst is looking to review past performance data of 12 securities in order to diversify a green resources investment portfolio.  He is reviewing the annualized performance and trading volume for these securities in order to make an informed investment decision.  In order to perform the analysis raw tables containing trading data for the year 2017 and 2018 were scripted using VBA in order to determine the annualized return on investment, as well as the annualized trading volume.  


Results:

Data Pulling
Data was pulled from two sheets containing summarized trading data from 2017 and 2018.  To pull the data, script was made to query the user for a year to analyze, then pulled data on 12 ticker symbols using forloops to sort through where trading data for a given ticker starts and ends.  While looping through the data by checking the ticker, trading volume was summed for all rows containing the same ticker through each iterration of the loop. The data was stored in an array according to the ticker symbol.

![image](https://user-images.githubusercontent.com/108313294/177612180-d27c7126-e3ef-4423-846e-00cb1a8487ac.png)
![image](https://user-images.githubusercontent.com/108313294/177612384-4affc9ed-9b02-47c3-a695-82ddae1ef108.png)

Data Calculations and formatting
Data collected from the forloops was placed in a new worksheet.  The headings for this worksheet were placed once by the script since they did not need to be looped.

![image](https://user-images.githubusercontent.com/108313294/177612798-e08c5f35-aeb6-4713-bf17-da00225c0f01.png)

The ticker data was stored in the TickerIndex and was placed in cells using a simple forloop.  TickerVolume is a single value stored in the TickerVolume array and is the sum of all trading data under the ticker defined by the TickerIndex.  The sum calculation was already completed during data pulling in the forloop.  Annualized performance was calculated by dividing the ticker end price and start price then subtracting 1 to determine the annualized return as a percentage in raw decimal output form.  The output of this data from the array was done using a loop according to the TickerIndex.

![image](https://user-images.githubusercontent.com/108313294/177613448-80f2ce61-61b5-4091-8215-a56952651315.png)

Final formatting of the data was performed to make it easier to quickly draw conclusions from the analysis.  A table with lines drawn separated the headings from the table in an aesthetically pleasing way.  The fonts were bolded to make it easier to read.  Trading volume was formatted to be separated by commas to be easier to read, and the raw decimal output of the annualized return was reformatted to be an actual percentage value. Using a simple loop, the annualized returns were also conditionally colour coded to determine if the security had a positive or negative return on investment.

![image](https://user-images.githubusercontent.com/108313294/177613917-cfaf295c-80a6-4ea9-89bb-6d6b54b973ad.png)

Summary:

Benefits/Disadvantages of code refactoring
The code use for this analysis was refactored from an original code that did not index the data by ticker in order to utilize the forloops more efficiently.  In general, refactoring code has the benefit of optimizing code for faster execution which increases the scalability to accept larger potential datasets.  The code can be made easier to understand and modify later. Refactoring code can also often also uncover bugs and limitations of code. However, refactoring takes time and unless the end user will be appling the code to very large datasets, opimizations may not always be noticed. In this case the difference between 0.3 seconds and 0.04 seconds is neglegable to the end user. 

Original vs Refactored Stock Analysis Code
The refactored code is about 750% faster than the original code, which when scaled to massive sets of data can save a lot of time.  The refactored code also makes quality of life differences to the end user, allowing them to pick the dataset by year they want to analyze. However, a disadvantage of this code is that it creates three arrays of data to store all the ticker information.  While in the scale of this project there is negligable impact on modern hardware, if this project were scaled up, there would likely be a situation where RAM storage may become a problem.  The original code looped and wrote data directly to the spreadsheet instead of committing code to volatile memory.



