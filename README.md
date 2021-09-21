# VBA of Wall Street
## Overview of Project
The purpose of this analysis was to learn the basics of VBA (Visual Basic for Applications) and apply it to an investing problem which looked at stock prices for several companies in a two-year span (2017 and 2018).  In the given scenario, Steve is looking into his parent's investment into DAQO New Energy Corporation, a green energy company.  He decides to look not only into DAQO's stock (ticker "DAQ") but other green enrgy companies as well.  The data was initially given in daily ticker prices for each company and using VBA we were able to calculate total volume by year, as well as starting and ending prices for each year for each ticker.  This allowed us to see the total yearly volume for each company, as well as the yearly growth for each stock.

## Results
We were tasked to look at the stock "DQ" in particular.  While virtually every stock in the analysis did well in 2017, DQ did exceptionally well, posting a return of 199%, compared with an average return for the stocks of 67% (see results in link below).  
https://github.com/bmoazen/Vandy-Stock-Analysis/blob/main/2017%20Stock%20Analysis.PNG  
However, while every stock posted negative returns in 2018, DQ had a much worse return (-62.6%) than the average (-8.5%).  Furthermore, two stocks (tickers "ENPH" and "RUN") posted positive return in both years, with ENPH posting better than 80% return in 2017 and 2018 (link below).  This suggests that perhaps ENPH would be the better investment, or would at least be a good candidate for diversifying Steve's parent's portfolio.
https://github.com/bmoazen/Vandy-Stock-Analysis/blob/main/2018%20Stock%20Analysis.PNG  

Using the refactored code, the run times for the 2017 and 2018 stock analysis differed by less than one-hundreth of a second (0.168 and 0.172 seconds, respectively - see images below)
https://github.com/bmoazen/Vandy-Stock-Analysis/blob/main/VBA_Challenge_2017.PNG
https://github.com/bmoazen/Vandy-Stock-Analysis/blob/main/VBA_Challenge_2018.PNG  
Both of these times were much faster than the original code I wrote in the modules (0.97 seconds for the 2018 analysis).  While the difference in time between the original and refactored code may be almost inperceptible now, if the analysis involved millions of rows, this time difference would be much larger and would most likely become a factor in which code should be used.

## Summary
One major advantage of using refactored code is the time it have saves.  If you can use existing code and modify it as needed, this is much faster than writing your own code from scratch.  However, using refactored code can have its disadvantages, as it may be tempting to use pre-written code without knowing what the code is actually doing.  For someone learning to code, knowing what each line is doing is crucial to the learning process.  Also, without knowing what the code is actually doing, you are much more likely to have errors in the code.  For example, the code may be calculating something other than what you intended.  Without knowledge of what the code is doing, you may have less of a chance of catching that.  For these reasons, it may be more advantageous in some situations for beginners to write thier code, even if it is longer and perhaps more "clunky" than refactored code.

In this project, using the refactored was advantageous for me, since I understood how the refactored code was structured and why it was written that way.  
