

# stocks-analysis module 2 challenge

## Overview of Project

#### the purpose of this analysis

######  The purpose of this project is using a Excel VBA coding to refactor the function that i made during the module 2, to make it more effective to gather huge amounts of data, in order to visualize if the stocks are worth to investing. however，the goal for this challenge is to use a better way to make the operation process faster and efficiency。

## Results

#### **part 1** (formatting and set up)
<img width="868" alt="Screen Shot 2022-06-06 at 10 47 46 PM" src="https://user-images.githubusercontent.com/106010498/172314136-a4600f1b-d62f-4bcf-8c2f-fb04d4056ad8.png">

###### First in this code, The following couple things are created, 
###### I create start-time and end-time for calculating the processing time; 
###### And set up year-value in the inputbox in order to let the button choose what worksheets we want to run on;
###### Format the output sheet by create the all stocks (“ + yearvalue + ”)and create header row of “ticker”,” total daily volume”, and”return”; 
###### It’s needed to initialize the array of all tickers, in this case we have 12 count of value, so the tickers(12) if from tickers(0) to tickers(11)
###### And at by following I activate the data worksheet that we going to choose from the button bar, and get the number of rows to loop over By using this function, “Rowcount=cells(rows.count,”A”).End (xlup.)row”. 
###### For this part , the coding is similar to the original one, most of part of coding Is to format the worksheets and set up the Marco. 


#### **part 2** (the core code of different from the original one)
<img width="998" alt="Screen Shot 2022-06-06 at 11 11 26 PM" src="https://user-images.githubusercontent.com/106010498/172314482-5036ab9e-2b9b-475f-a48c-63a40231ff80.png">

###### from here the Marco will be different from the original one, and the way that loop & if would be changed. 

###### at the first an index value was made to across the following arrays that i will make next. 

###### And set up three output arrays will be used to increase and Calculation by next couple “for loop” and “if then” function.

###### The first “for loop" function is to initialize the ticker volumes to zero.

###### And set an other for loop i from 2 to Rowcount 
	set 4 condition if then function 
		1， when the index value that we set up before “tickerindex = 0”, we want to continuing increase the volume for current ticker .
		2, finding the current row is the first row with the index value, get the value as the starting price
		3, and we also need finding the last row with the index value and get the value as the ending price
		4, if the row is the last row of the index value, the index value will be increased by +1.

###### According this loop, the index will increase when the last row show up, and stop increase the totalvolume of this index, and index + 1, then totalvolume, startingprice and ending price will initialize to 0

###### And the end we loop the through out our arrays to ourput cells . 

#### **part 3** (formatting and closing the sub)


<img width="948" alt="Screen Shot 2022-06-06 at 10 48 06 PM" src="https://user-images.githubusercontent.com/106010498/172315187-a9e4ec64-a621-4de5-ad08-647f09235cec.png">

###### finaly  we format the worksheets and get then the Marco get the end.

#### **compare the execution times of the original script and the refactored script**

<img width="998" alt="2017 new" src="https://user-images.githubusercontent.com/106010498/172316995-14b9e415-e7c4-4ea8-9d5a-f98339bc02ce.png">
<img width="988" alt="2017 old" src="https://user-images.githubusercontent.com/106010498/172316998-d81d6ed4-7fcd-4d95-9f67-5de23f6d212b.png">
<img width="1004" alt="2018 new" src="https://user-images.githubusercontent.com/106010498/172316999-9ce1479e-6fda-43a7-a3ab-ff7c1650ac57.png">
<img width="980" alt="2018 old" src="https://user-images.githubusercontent.com/106010498/172317006-0d621318-6070-4075-bf5b-f348339b4374.png">

###### the execution times of 2017 original is about 0.25 second and 2017 refactored is about 0.07 second
###### the execution times of 2018 original is about 0.25 second and 2018 refactored is about 0.04 second 
###### the refactored marco is 4 time faster then the original marco 

## Summary

#### advantages or disadvantages 

###### the advantages of refactoring the macro is running time, which means we can gather more huge amounts of data using less amounts of time.
###### the disadvantages of refactoring is that may cause more debug and there must be some random factor that we need to be take care.


#### pros and cons 

###### pros of refactoring a code make our thinking about in different way and getting better marco we did, they would be more organized, easy to read, faster and cleaner.  not only more efficiency, but more imagination

###### cons of refactoring that i think is sometime we don't really know if the new code will be whether or not efficiency. we may run into some debug, may be the new way we thought just not work, in the really world we dont know if the result is same, and we have to take more time to fix a debug, test the marco, which may couse more risk and the result could be slower then before. 

## However, we like to make things better and better, sometimes refacorting is necessary. we like the adventure, because that play an important role in the progress of our science and technology
