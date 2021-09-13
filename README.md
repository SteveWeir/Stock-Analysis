# Stock Analysis with VBA

## Overview of Project

### Purpose

This project was intended to evaluate an analyst's ability to edit or refactor exisiting VBA modules and subroutines, for the purposes of making code run faster and more efficiently. In this specific case, we are refactoring and condensing nested "for" loops within a VBA module that can be used to aggregate volume and calculate returns for an index of 12 stocks, between 2017-2018 

## Results

The efforts to refactor proved a success.  Code can be found in the files (VBA_Challenge) attached to this repository, with Module1 being the original script and Module2 containing the refactored code. As with the original iteration of the VBA script, 2017's returns were far more promising than 2018's almost index-wide, which can be seen below:


![VBA_Challenge_2017 Data](https://user-images.githubusercontent.com/85244439/133020472-17b425e1-061b-43ae-ac18-67b9e5dcf55b.PNG)
![VBA_Challenge_2018 Data](https://user-images.githubusercontent.com/85244439/133020478-9c572f97-7514-4206-a3e2-36d995ecc914.PNG)


While the accuracy of the script was unaffected and the same, reliable results were achieved, the primary benefit of the refactoring is reflected in the run-times of the original script (top) vs the refactored code (bottom):


![VBA_Challenge_2017_ORIGINAL SCRIPT](https://user-images.githubusercontent.com/85244439/133020837-59afbd6a-038e-4b80-844a-7ccf92bd5ca6.PNG)
![VBA_Challenge_2018_ORIGINAL SCRIPT](https://user-images.githubusercontent.com/85244439/133021098-afce3594-6271-4833-8bb6-49feb9fdfd7f.PNG)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/85244439/133020852-97e152e3-e8d0-48b5-8afb-fa5f24ba89f5.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/85244439/133021166-99a7b41b-3847-4242-b99f-96c46b4bdbe2.PNG)





## Summary

Refactoring of existing code is a process complete with benefits and drawbacks. By condensing existing code, an analyst or developer may be successful in fine-tuning the logic of a module or sdubroutine to make it run faster and/or more reliably. However, as was shown to the case during my completion of this project, refactoring of existing code can, at times, seemingly create more problems than it solves. While I was eventually able to get my code to run smoothly and quicker than before, I spent multiple hours trying to debug overflow errors, complitation errors, and subscript-out-of-range errors that stemmed from my attempts to refactor the existing VBA script.
