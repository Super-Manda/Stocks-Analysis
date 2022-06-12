# Green Stocks Analysis Project Overview :sun_with_face:

## Purpose of This Analysis

###  
Steve has just obtained a finance degree, and his first clients are his parents, who foresee the days when fossil fuels will be depleted.  Steve's parents have invested their money in a green energy company known as DAQO Energy Corporation (DQ), and Steve is concerned that they’re not well diversified within their penchant for green stocks.  Steve has a 2017-2018 dataset of green energy companies and wants to see more research about DAQO and other comparable companies using macros.  From there, he wishes to have an Excel workbook that will accommodate future years of data over time.  This Excel file is found here: [VBA_Challenge](https://github.com/Super-Manda/Stocks-Analysis/blob/main/VBA_CHALLENGE.xlsm)

# Results 

##  Comparison of Stock Performance Between 2017 and 2018

### **2017** 

In 2017, most of these stocks saw substantial returns at an average of 67.3%.  Only TERP, the TerraForm Power Plant company, saw a loss of 7.2%.  DAQO performed the best out of these tickers in 2017 with a return of 199.4%, followed closely by SEDG, a solar photovoltaics company, and ENPH, a software company that specializes in monitoring green homes, in third place.  

What appeared most interesting in the Total Daily Volume column is that DAQO did not do as brisk of a business as any of the other companies that were examined in 2017; however, it produced the highest return.  Likewise, although a solar power company like SPWR had a large number of daily shares traded, it was on the lower end of returns in 2017.  There is a slightly negative, but negligible, correlation overall (-0.12).

![VBA_2017_ORIGINAL]( https://github.com/Super-Manda/Stocks-Analysis/blob/main/VBA_2017_Original.png)


### **2018**
In 2018, these stocks surveyed had an average loss of 8.5%.  The correlation between the Total Daily Volume and the Return is now moderate (0.5).  RUN, true to its ticker name, comes out of nowhere and produces an 84% return while significantly increasing its total daily volume.  ENPH remains a good stock in both years surveyed.  Based upon ENPH’s mission statement, one hypothesis could be that this stock is not relying upon the commodity, but rather, on the software that governs users of the commodity.  In other words, if there is no sun or wind gusts for a period of time, then companies like TERP (wind and solar) and SPWR (solar) may lose business.  If a trade or currency war occurs, and/or silicon becomes scarce, then DAQO could lose business.  It is harder for ENPH to theoretically lose business because, to lose business, it would be the equivalent of a person uninstalling their green thermostat or their solar panel.  Usually, once people have green software, they remain a customer, unless the software breaks down a lot, so perhaps the company can focus more on new business than on the vagaries of the commodities.  Therefore, it seems logical to advise Steve to look at some “safer” (and also more lucrative) selections, such as ENPH, to counterbalance his parents’ DAQO selection, because DAQO varied from a 199% return to a 63% loss from 2017 to 2018.  Although it may have averaged out to still be positive, it still appears to be volatile relative to ENPH.  

![VBA_2018_ORIGINAL](https://github.com/Super-Manda/Stocks-Analysis/blob/main/VBA_2018_Original.png)


##  Execution Times of the Original Script and the Refactored Script

### 
The original script ran in .42 seconds for the year 2017 and .44 seconds for 2018.  In comparison, the refactored script ran in .06 seconds for the year 2017 and .07 seconds for the year 2018.  This is not a huge difference in practicality based on this dataset of green stocks, but if the dataset were so enormous in the future that it took time to process (_e.g._: a dataset of every single stock presently available regardless of company type), then it would appear that the refactoring should allow the data to be processed much more efficiently.  

**Here are the end results of the refactored timers:**

![VBA_2017_CHALLENGE]( https://github.com/Super-Manda/Stocks-Analysis/blob/main/VBA_Challenge_2017.png)

![VBA_2018_CHALLENGE]( https://github.com/Super-Manda/Stocks-Analysis/blob/main/VBA_Challenge_2018.png)

### 
The tickerIndex is used more often in the refactored code, which may contribute something toward this efficiency.  For example, at outset, “tickerIndex = 0” was created.  Then, in the refactored code, it becomes a much more active reference in the code, especially in sections 3A through 3D.

![VBA_REFACTORED_CODE_SAMPLE](https://github.com/Super-Manda/Stocks-Analysis/blob/main/VBA%20code%20sample.png)

Another example of efficiency is that the original script has a separate macro for the formatting, so Steve would have to click on multiple macros to achieve a user-friendly formatting.  With the refactored version, all Steve has to do is tell the input box whether he needs to analyze 2017 or 2018, and then everything will come up with the proper colors and percentages that he needs to see.


# Summary

## Advantages and/or Disadvantages of Refactoring Code  

### 
As aforementioned, the primary advantage is that refactoring the VBA code will allow Steve to add to his dataset over time.  

In addition, refactoring can streamline or pare down things that were previously hard-coded (to no longer require as many separate updates).  Therefore, the next person to read the code will find it to be simpler and easy to follow at a quick glance.  It is also like doing an audit of which facets of the code are most important to a particular subroutine in terms of using less memory. 

The disadvantages of refactoring include that it’s something else that could potentially go wrong and provoke a bug.  The instructions on the module also allude to the fact that, sometimes, it’s a pre-existing “legacy” code that everyone is stuck with, so the best bet may be to refactor it.  Also, Steve may not have an appreciable improvement when he clicks the macros if there is ultimately no change in the goal of each subroutine—it’s more of a back-end type of thing.  Lastly, without refactoring, in theory, a team could potentially split up different tasks and operate independently up until they collaborate their contributions into one final end product if Steve wants to significantly increase his dataset and hire more people, but with refactoring, one has to look more at the whole picture.  In other words, it seems as though refactoring requires re-reading the entire code to understand the whole picture. 

##  Pros and Cons Applied to Refactoring the Original VBA Script 

### 
In this exercise, there were some examples where hard-coded values were replaced and concatenated, such as the three lines of code that depended upon the year.  For example, these were changed from Range("A1").Value = "All Stocks (2018)" to instead say, Range("A1").Value = "All Stocks (" + yearValue + ")".  This is also an example of how Steve can grow his dataset over time.  As aforementioned, the tickerIndex and streamlined formatting will also be helpful to Steve in this refactored code.

In this challenge, there are instances where the instructions state to leave something alone, such as Sheets("All Stocks Analysis").Activate, and this appears to be so that the creation of bugs can be avoided.  It also sets up the steps to refactoring so that a skipped step does not hamper the overall code.  In addition, it’s true that Steve generally gets quick results regardless of the refactoring.

## Outside References
### 
To do this analysis, the following links were used to research more about some of these companies: 
- https://en.wikipedia.org/wiki/Daqo_New_Energy
- https://en.wikipedia.org/wiki/Enphase_Energy 
- https://en.wikipedia.org/wiki/SolarEdge 
- https://en.wikipedia.org/wiki/SunPower 
- https://en.wikipedia.org/wiki/Sunrun 
- https://www.terraform.com/ 
