# VBA of Wall Street

## Stock Analysis Overview    
In this green stock analysis, we compared green investments in terms of profit and loss over the course of a year given two years: 2017 and 2018  

### Purpose  
The goal was to provide the client (Steve) with the tools (using VBA in excel) to effectively and efficiently analyze the given dataset made up of 12 different stock tickers 
but also incorporating into the code the room for growth in the event he needs to work with larger datasets in the future.  

#### Output Categories 
- Total Volume for given year  
- Rate of Return (whether loss or gain)  
	- formatted as red for loss and green for gain to easily discern between the two.  
	- this is extremely helpful in the event Steve does begin to compare/analyze larger datasets.  

#### Built-In Automations  
- Input button to request the year the client would like to analyze  
- Button that will clear the contents and allow for reformatting of the data once called again.  

## Analysis and Challenges  
Given the results of the analysis (most all stocks did not do well in either 2017 or 2018), Steve would benefit from two things:  
- Expanding the stocks that he chooses to analyzes...i.e. increasing his dataset  
- Expanding the years that he looks at those stocks to determine if there is a time-pattern for those type (green) stocks  

These two simple but important additions could help him make a more informed decision when advising which stocks to choose.

###Testing for expandability  
The original VBA code had to be refactored in order to improve efficiency. Changes made were as follows:  
- created a loop that would loop over only rows that showed a value omitting any empty rows  
- created a variable that eliminated the "magic number" therefore increasing the expandability of the code  
- used a tickerIndex variable for the array that allowed for greater expansion in the event more tickers were added in the future  

## Results  
See the runtime of both the 2017 and 2018 years:  
![VBA_Challenge_2017](https://user-images.githubusercontent.com/83401820/123564394-f4551580-d77e-11eb-8255-3f5e9827539f.png)  

![VBA_Challenge_2018](https://user-images.githubusercontent.com/83401820/123564400-fa4af680-d77e-11eb-954f-67c722e1bb64.png)  


## Summary  
- Advantages of refactoring code are:  
	- makes debugging easier  
	- makes the code fresher, easier to understand  
- Specific advantages of refactoring this code were:  
	- eliminated redundancy  
	- improved expandability  

- Disadvantages of refactoring code are:  
	- if the code has a bad design, it would not be beneficial  
	- might require a great deal of testing and retesting  
- Specific disadvantages of refactoring this code were:  
	- because I wrote the original code and checked it, new it worked, I was reluctant to see any ways that it could be improved (owner-bias)  
	- the desired outcome was to make it run faster by streamlining some code, but in the effort to make the code more expandable, the actual runtime increased  
In conclusion, while the code, refactored does allow Steve to run a larger dataset, it does not necessarily decrease the runtime.  
Obviously the more data the program has to analyze, the longer it will take to analyze that data, but with the enhanced loop structure and the buit in automations, Steve  
will be able to do so without making major changes to the code.



