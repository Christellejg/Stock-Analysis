# Stock-Analysis

Deliverable 2: Written Analysis of Results

Overview of Project: 

This challenge was to help Steve do more research for his parents on the stock analysis to include the entire stock market. Therefore, the challenge was to refractor the code used in Module 2 to now loop through all the data in the datasheet for the year 2017 and 2018. We also wanted to determine, after refracting a successful code, whether the refracted code made the VBA run faster than our original code.

Results:

Based on the results that I get from a successful refactored code, all stocks, with the exception of one, had a positive return for the year 2017. However, when we run for the year 2018, we can see that only 2 had a positive return. Stock: ENPH and RUN where the only two stocks that was positive for 2017 and 2018.

2017 Results:
![image](https://user-images.githubusercontent.com/78320504/149605907-2573ccc7-24d3-4b71-b063-73d22b748e25.png)

2018 Results:
![image](https://user-images.githubusercontent.com/78320504/149605946-34eee1dc-b001-40a2-849d-08e59a3f9dd1.png)

Below, we can conclude also that the refactored code runs faster than the original code. The orignal code, for 2017, ran in 0.6328125 seconds (see below):
![image](https://user-images.githubusercontent.com/78320504/149606095-1930b5c7-25b2-46e9-937c-f938dbef0267.png)

And, for 2018, the original code ran for 0.609375 seconds (see below):
![image](https://user-images.githubusercontent.com/78320504/149606121-deb531a7-4319-4b2b-8d7d-35c5d42f4a85.png)

See the uploaded Resources folder to see the lapse of time for the refactored codes for 2017 and see how 2017 and 2018 ran faster.

Summary: 

One advantage of refactoring codes is that it makes the code more efficient. Another advantage is that the refactored code runs faster than the original code. These are great advantages especially when analyzing a large dataset. A disadvantage to refactoring codes the possibility of duplicate a line. The efficiency of refactoring the original VBA script is beneficial because it allows us to better analyze a larger dataset. This is a pro in the even that we want to come back and revisit this dataset in the future. The con of easily duplicating codes when refactoring can occur when you are writing a line of code over and over. Another example of this disadvantage is when we copy and paste multiple lines. By either typing codes over and over or copying and paste codes, we can end up with a duplicate code line in our code.
