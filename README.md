# Refactoring Steve's Stock Analysis VBA Code and Measuring Performance
## Overview of Project
The purpose of this project was to compare a dozen stocks for Steve between the years of 2017 and 2018, and find their total daily volume (total number of shares traded throughout the day) and yearly return for each stock. We then needed to refactor the intial VBA code that was designed to analyze a dozen of stocks for Steve, and make it executable for thousands of stocks without taking a long time to execute.
## Analysis and Results

![2017 Stock Analysis](https://user-images.githubusercontent.com/101950175/161590584-3644e495-f3e3-4f5f-9eab-75bddf60a840.png)

![2018 Stock Analysis](https://user-images.githubusercontent.com/101950175/161590591-a7d80720-f3f0-439d-bfa0-36378d649fa7.png)


![Initial 2017 Stock Analysis Output](https://user-images.githubusercontent.com/101950175/161590391-ad9c1238-0b86-4b40-8b7b-521f8e4cdee7.png)
![Intial 2018 Stock Analysis Output](https://user-images.githubusercontent.com/101950175/161590433-f667fed5-440c-4c89-8a41-a9acd5bf88b6.png)

After refactoring the code we were able to shorten the run time in 2017 from 0.287 seconds to 0.066 seconds, and in 2018 from 0.273 to 0.074 seconds - it now runs for both years in about a fourth or the original run time.
### Refactored 2017 Stock Execution Time
![VBA Challenge 2017](https://user-images.githubusercontent.com/101950175/161594236-5929a954-5fc2-487c-8543-a9ea737d12ea.png)
### Refactored 2018 Stock Execution Time
![VBA Challenge 2018](https://user-images.githubusercontent.com/101950175/161594253-2f900419-e3e0-4839-8d1d-a389a6df4935.png)
## Summary
### Advantages and Disadvantages of Refactoring Code
There are advantages to refactoring code:
- efficiency - taking fewer steps, using less memory, and the removal of redundant code.
- readability - improves design and makes the code easier to understand.
- 
### Pros and Cons of Applying Refactored Code to Original VBA Script
