# Green Stock Analysis

## Project Overview
The goal of this project is to analyze stock data from the years of 2017 and 2018. The analysis will focus on several companies that operate in the green energy space. The goal is to allow our clients to make an informed investment decision regarding their upcoming investment into green energy companies. To facilitate this analysis an Excel workbook was employed to store the data and to perform the analysis. The Excel workbook consists of two wokrsheets that hold the 2017 and 2018 stock data as well as another worksheet named " All Stock Analysis" used to perform the analysis on the data.  

## Results
The analysis for the project began with raw 2017 and 2018 stock data for 12 green energy companies which included the total daily volume for each company. Using VBA code the daily volumes for each company were used to calculate the yearly return for each of the 12 companies. 

![image](https://user-images.githubusercontent.com/96552268/162618732-64c78a2b-454d-4654-841d-a3c5333c6226.png)
![image](https://user-images.githubusercontent.com/96552268/162618752-a7d75318-2f13-4a9c-a2bc-c8fba0ed32f9.png)

In the VBA code FOR loops were used to loop through each row for the current company's ticker and increase the ticker volume by adding each subsequent row's daily volume to the previous total for that company until the yearly total volume was calculated. Through the FOR loop this process was repeated for each company's ticker name until the yearly total volume was calculated for all 12 companies. 


## Summary
