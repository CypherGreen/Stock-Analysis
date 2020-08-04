# **Stock Analysis**
---
## *Analysing Wallstreet stock data to determine the best investments*
---
### **Overview of Project**

The purpose of this project was assisting Steve in analyzing Wallstreet data from 2017 and 2018, then giving him the ability to analyze the results himself for any other year and  any other additional stocks he wants to add to the workbook. His first goal was to analyze DQ stock data for his parents. After getting those results he needed to analyze other stocks for a well-rounded comparison for them. Refactoring the code originally used to pull the required results helped to improve the overall efficiency of the workbook, while also leaving it adaptable for his future needs.
---

## **Results**

### **Brief Explanation of Coding Syntax**

The refactoring process of this project focused on reformatting the original code to pull data from any year of the workbook, not just 2017 and 2018. You can see this in the example of the code used by the activated worksheets calling for any year to be run and by the statements and code generalizing the ticker index instead of it being more specific like the original. The code for our original ticker index only included the current twelve tickers. Since Steve wants the ability to run a larger dataset of tickers, the tickers were the main focus throughout the blocks of code. This new code is more generalized but also more direct in the loop it utilizes, which allows for less run time.

![Example_Code](https://github.com/CypherGreen/Stock-Analysis/blob/master/Challenge_Resources/Example_Code.png)

Citation for the use of the RowEnd code line: Laviano, Joe(2013)Excel VBA[rowEnd = Cells(Rows.Count, "A").End(xlUp).Row]https://stackoverflow.com/questions/18088729/row-count-where-data-exists

### **Analysis of Original Code vs. Refactored Code**

The run through of our first code ran smoothly, however, it had more code lines put into it and was a little on the slow side for only analyzing up to twelve tickers. This original code was not bad, but the timeliness and adaptability could be improved. The original runtime for the first code looked like this:

![Stocks_2017](https://github.com/CypherGreen/Stock-Analysis/blob/master/Resources/Stocks_2017.png)
![Stocks_2018](https://github.com/CypherGreen/Stock-Analysis/blob/master/Resources/Stocks_2018.png)

Now, the refactored code produced a much quicker runtime as can be seen below. Steve will spend less time running the code for his larger dataset. 

![VBA_Challenge_2017](https://github.com/CypherGreen/Stock-Analysis/blob/master/Challenge_Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/CypherGreen/Stock-Analysis/blob/master/Challenge_Resources/VBA_Challenge_2018.png)

### **Analysis of 2017 and 2018 Stock Performance**

The yearly returns for the 2017 ticker dataset was relatively good in overall, with only one stock plummeting severely. The yearly return for 2018 was much worse, with only two stocks making it through the year in positive percentages. Eight out of the twelve tickers analysed show an increase in daily volume for 2018, despite the majority performing poorly. Steve will need to look at other stocks to recommend to his parents. Even though ENPH was still in the positive for 2018, it still decreased in percentage. As for RUN, it did have a positive increase in 2018, so it may be one to keep an eye on in the future while still taking a look at other stocks that might have done better over time.
---

## **Summary**

There are benefits in refactoring previously used code. The original code might need to be updated to encompass new worksheets or data to the workbook. Refactoring code can also improve analysis time lapse as shown above. If there is a way to condense code while still leaving it effective, then it would also be good to view it from that perspective. Looking for new ways to improve productivity is always a good idea. A downside of refactoring code is that if you do not understand the coding of the original then that would make it harder to improve upon. If a programmer’s code is unorganized and without any comments to explain the purpose of their code, then that leaves the next programmer after them with potentially unusable code as a basis. 

For example, since the original code for the All Stocks Analysis already had comments to explain the purpose of each code block, the ground work was already laid out and all that was needed was a few minor adjustments. It would have taken a lot longer to analyze the data and interpret how the code functioned without those statements. 
