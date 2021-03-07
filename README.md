# Using VBA to Conduct Stock Market Analysis

## Overview of Project
For this project, I helped Steve create a program to analyze a set of stocks for 2017 and 2018. Microsoft Visual Basic for Applications, or VBA, was used to complete this effort. We initially looked at a stock his parents had interest in "DQ" but we soon realized this stock was underperforming. DQ shares witnessed losses of 63% in 2018 making it the worst performing stock in an underwhelming market. We expanded the scope of our analysis to include 11 additional stocks which informed our ultimate recommendation to **Buy ENPH**!

## Results

### Stock Market Findings and Recommendation

We recommend that Steve's parents buy ENPH. If one is looking only at **Figure 1**, he or she may determine that buying RUN is the smarter move because its return of 84% leads the market. However, if one looks at **Figure 2**, he or she will see that RUN actually saw one of the smallest gains at 5.55%. In this same year, ENPH did very well by ending the year **+129.52%**. Over the two years, ENPH had the greatest returns. However, the case with DQ serves as a cautionary tale that past success is not indicative of future returns. 


#### Figure 1

![2018StockReturns](2018StockReturns.png)

#### Figure 2

![2017StockReturns](2017StockReturns.png) 

### Stock Analysis Program 

The stocks' performance was not the only significant finding from this project. I also gained experience in re-factoring my Stock Analysis program. The original code I used to analyze the stock returns in 2017 and 2018 was funcitonal, but it had an awkward lag. With the goal of improving the customer experience, I used a timer function to measure the time it took for the program to run from start to finish. This is a useful proxy in gauging the impact on user experience. **Figure 3** shows the time it took for the original program to analyze stock performance in 2017 and 2018. Both processes took greater than one second. This may not seem like much, but this delay is enough to render this program of little value on the commercial market. Faster speeds and an improved customer experience are critical if I want to bring this functionality to more customers beyond Steve's parents.

#### Figure 3
![VBA_Challenge_Slow2017](VBA_Challenge_Slow2017.png)
![VBA_Challenge_Slow2018](VBA_Challenge_Slow2018.png)

### Stock Analysis Program Re-factored
My Re-factored program analyzed the stock market performance of 2017 and 2018 with a lot more efficiency. As shown in **Figure 4** both processes took less than a third of a second. This is a much-needed improvement. 


#### Figure 4
![VBA_Challenge_2017](VBA_Challenge_2017.png)
![VBA_Challenge_2018](VBA_Challenge_2018.png)

## Summary

Re-factoring this code was not a simple task. It was a frustrating and challenging process. "Why am I spending so much time 'fixing' a functioning program?" crossed my mind more than once during the project. Yes, there was concrete evidence that the re-factoring process decreased the amount of time it took for the program to run, but I was interested in greater returns that I uncovered upong further reflection.

### Pros of Re-factoring
1) The code is *cleaner* and thus easier to read and understand. This makes working in teams a smoother, more effective process.
2) The re-factoring process uncovers bugs that would otherwise be less apparent.
3) The re-factoring process speeds up processes. Reducing lag and processing times is a omnipresent challenge and this applies to both simple tasks and complex programs.

### Cons of Re-factoring
Re-factoring does have some drawbacks, however, beyond being a frustrating process for novice programmers.
1) Re-factoring code does take time, and in the commercial world, time means resources and money. Most businesses seek to reduce operating costs, but this is especially critical for start-ups working with limited capital. Although re-factored code is more robust and adaptable, if re-factoring the program would lead into missed deadlines and budget overruns, re-factoring efforts may need to be limited or delayed until time and funding factors allow for the team to re-factor the existing program.

In the case of re-factoring the Stock Analysis program, I did not face time or budget constraints. Additionally, the clunky nature of the first code, without tickerIndex, provided ample incentive to re-factor this program, although it was functional in its unrefined form. The benefits go beyond a shorter processing time. Additional features and analysis can be applied to this set of stocks more easily and efficiently by utilizig the re-factored codes and indices.


