# VBA-challenge
All the files are attached in the Word document 

Module2_Challenge_VBS.doc
Summary of the vbs files
Main Module.
  This is the main logic.  
1.	Call Turnoff : I had to use this to improvise little bit of performance on the code.  I have learnt this from youtube videos.  This module turns off screen updates , events and makes calculation manual until the last step.
2.	Call Turnon:  This module turns on the screen,events and calculations back on in the last step.
3.	Stock Analysis:  This is the main code which reads all the data of the worksheets one by one sheet into dictionary and process final result into  arrays 
4.	Write_data : After processing writes below columns on to the appropriate worksheet
a.	The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
b.	The total stock volume of the stock. The result should match the following image:

Class Module.
   This module consists of two subroutines.
1.	CalculatePercentageChange() :  Calculates the difference of the close price for last day and open price for first day of the year by ticker symbol 
2.	AccumulateVolume : Calculates the Total volume of each Ticker  

Thank you
