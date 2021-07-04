# Challange 2

## Overview of Project: Analyse an entire dataset by refactoring code

### Results:

The refactoring consisted of going through the complete dataset only once and assigning the corresponding values to each ticker, thus reducing the time it took to run the code:

    ''Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then

                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then

                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
            End If
        
        Next i

Time with initial code:

![img](https://github.com/CarmenU18/Module-2-Challenge/blob/main/Resources/Time_Initial%20code_2017.PNG)

Time with refactoring code:

![img](https://github.com/CarmenU18/Module-2-Challenge/blob/main/Resources/Time_Refactoring_2017.PNG)

Now that we have the 2017 and 2018 summary, we identify that 2017 performed better than 2018, with an average return of 67% versus -8.5% respectively.

![img](https://github.com/CarmenU18/Module-2-Challenge/blob/main/Resources/Return.PNG)

The best performing stocks were RUN and ENPH with positive performance in both years.

![img](https://github.com/CarmenU18/Module-2-Challenge/blob/main/Resources/VBA_Challenge_2017.PNG)

![img](https://github.com/CarmenU18/Module-2-Challenge/blob/main/Resources/VBA_Challenge_2018.PNG)

The best performing stocks were RUN and ENPH with positive performance in both years.
Despite the negative performance in 2018, we have issuers that we would focus on and follow up on on a timely basis in order to be able to trim positions while still maintaining positive results: SEDG, VSLR, FSLR and

![img](https://github.com/CarmenU18/Module-2-Challenge/blob/main/Resources/2017.PNG)

![img](https://github.com/CarmenU18/Module-2-Challenge/blob/main/Resources/2018.PNG)

The stocks in which I would take the decision to close position and not continue adding losses to the portfolio would be SPWR and TERP.

### Sumary
Code refactoring is used to reorder or restructure existing code.  The goal is to execute the refactoring without changing the external behavior of the code.
But, as in everything we have advantages and disadvantages of refactoring. One of the advantages is that after refactoring, the code is easier to understand or read, less complex and faster to run. The disadvantages are that the change makes the code slower or not functional, so you will need to spend more time to correct or revert to the previous version.

In this case, it was necessary to refactor in order to create a code that would be optimal for reading a larger amount of data while optimizing the reading time.

