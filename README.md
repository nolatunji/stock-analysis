# Green Stock Analysis

## Project Overview
The goal of this project is to analyze stock data from the years of 2017 and 2018. The analysis will focus on several companies that operate in the green energy space. The goal is to allow our clients to make an informed investment decision regarding their upcoming investment into green energy companies. To facilitate this analysis an Excel workbook was employed to store the data and to perform the analysis. The Excel workbook consists of two wokrsheets that hold the 2017 and 2018 stock data as well as another worksheet named " All Stock Analysis" used to perform the analysis on the data.  

## Results
The analysis for the project began with raw 2017 and 2018 stock data for 12 green energy companies which included the total daily volume for each company. Using VBA code the daily volumes for each company were used to calculate the yearly return for each of the 12 companies. 
![image](https://user-images.githubusercontent.com/96552268/162619975-4724fc41-7827-42e6-b650-3fc1b001ff37.png)
![image](https://user-images.githubusercontent.com/96552268/162619981-7a6d66cf-9a01-4793-80c7-b59d2824d6a1.png)



In the VBA code FOR loops were used to loop through each row for the current company's ticker and increase the ticker volume by adding each subsequent row's daily volume to the previous total for that company until the yearly total volume was calculated. Through the FOR loop this process was repeated for each company's ticker name until the yearly total volume was calculated for all 12 companies. Next the yearly return for each of the companies was calculated in order to give our clients the best information to make their decision regarding their investment. The final calculations for each company ticker were captured in the variables "Total Daily Volume" and "Return". 


![image](https://user-images.githubusercontent.com/96552268/162620427-8e74f716-b42b-4e34-bc5f-374b3c9a31a9.png)
![image](https://user-images.githubusercontent.com/96552268/162620441-44d76b00-ba81-4199-8b51-93240ab570e1.png)

After running the original Green stocks analysis we wanted to refactor the code to be able to run the analysis faster than we did with the original code. After refactoring the code the speed with which the analysis could be run was increased for both data years. 
![image](https://user-images.githubusercontent.com/96552268/162619779-4e018d23-be11-4bb3-9f4b-31138dbb14fe.png)
![image](https://user-images.githubusercontent.com/96552268/162620600-cb2af243-070e-4c0c-91a0-2700584b0bed.png)
![image](https://user-images.githubusercontent.com/96552268/162620611-50bf71d8-0231-491e-ba17-95fac364df18.png)
After running the Green stocks analysis we were able to see that the DQ stock that our clients were orginally planning to invest in had returns that decreased over 60% between 2017 and 2018 and might not be the best investment option.

## Summary
Overall there are advantages and disadvantages to refactoring code. The greatest advantage to refactoring code is that you create code that is easier to read and understand and that will run faster than if you had not refactored the code. A disadvantage of refactoring code is the time it takes to re-write the code to improve the function. 

In reference to the orginal VBA code the advantages were that the code was easier to write due to the fact that we used less variables and appeared to be more straigtforward. The disadvantages were that the code took longer to complete the analysis. The advantage of the refactored code is that it was able to process the analysis faster than the original code. The disadvantages of the refactored code was the time it takes to test the changes that are being made to refactor the code without disrupting the functionality of the orginal code. There were specific challenges with getting "Runtime Overflow Errors" on different varibles. The errors were debugged by reviewing the code several times and making changes to the variable data types (which did not resolve the issue) and continuting to review the code to discover the (tickerIndex) was missing as an arguement in certain lines of the code. 
