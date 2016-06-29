Attribute VB_Name = "Module1"
'***************************************************************
'Subroutine that supplies spreadsheet input as parameters to relevant functions
'Also responsible for printing the output in a clean manner
'***************************************************************
Sub forecastSavings()

    'Read spreadsheet inputs
    startingSavings = Range("B1").Value
    startingIncome = Range("B2").Value
    incomeGrowth = Range("B3").Value
    incomeSavingRate = Range("B4").Value
    returnOnSavings = Range("B5").Value
    inflationRate = Range("B6").Value
    yearsRemaining = Range("B7").Value
    proposedWithdrawal = 0 'Placeholder. Will update in later call to function
    lifeExpectancy = Range("B8").Value
    desiredInheritance = Range("B9").Value
    
    'Set active cell and clear everything below the Summary section
    ActiveSheet.Range("A12").Select 'Start at A12
    upperLimitRow = ActiveCell.Offset(7, 0).Row 'Clear contents below active row
    With Sheets("Retirement Savings")
        .Rows(upperLimitRow & ":" & .Rows.Count).Delete
    End With
    
    'Retrieve nominal & real savings, nominal & real savings after withdrawals
    nominalEndSavings = nominalSavings(startingSavings, startingIncome, incomeGrowth, incomeSavingRate, returnOnSavings, inflationRate, yearsRemaining)
    realEndSavings = nominalToReal(nominalEndSavings, inflationRate, yearsRemaining)
    'placeholder = nominalEndSavings 'Required for resolving a bug that I just cannot figure out
    'nominalSavingsAfterWithdrawal = nominalMonthlyWithdrawal(placeholder, proposedWithdrawal, lifeExpectancy, returnOnSavings)
    'realSavingsAfterWithdrawal = nominalSavingsAfterWithdrawal / (1 + inflationRate) ^ (yearsRemaining + lifeExpectancy)
    
    'Print values to spreadsheet and format
    'Organized by row offset
    Selection.Value = "SUMMARY OF RESULTS" 'Summary title
    Selection.Style = "Title"
    
    Selection.Offset(2, 0).Value = "Final nominal savings" 'Nominal savings
    Selection.Offset(2, 1).Value = nominalEndSavings
    Selection.Offset(2, 1).NumberFormat = "$#,##0.00"
    
    Selection.Offset(3, 0).Value = "Final real savings" 'Real savings
    Selection.Offset(3, 1).Value = realEndSavings
    Selection.Offset(3, 1).NumberFormat = "$#,##0.00"
    
    Selection.Offset(4, 0).Value = "Proposed nominal monthly withdrawal" 'Nominal monthly withdrawal
    'Selection.Offset(4, 1).Value = proposedWithdrawal
    Selection.Offset(4, 1).NumberFormat = "$#,##0.00"
    
    Selection.Offset(5, 0).Value = "Remaining inheritance in nominal terms" 'Remaining nominal inheritance
    'Selection.Offset(5, 1).Value = nominalSavingsAfterWithdrawal
    Selection.Offset(5, 1).NumberFormat = "$#,##0.00"
    ActiveCell.Offset(5, 1).Style = "Explanatory Text" 'Change style
    
    Selection.Offset(6, 0).Value = "Remaining inheritance in real terms" 'Remaining real inheritance
    'Selection.Offset(6, 1).Value = realSavingsAfterWithdrawal
    Selection.Offset(6, 1).NumberFormat = "$#,##0.00"
    ActiveCell.Offset(6, 1).Style = "Explanatory Text" 'Change style
    
    Selection.Offset(10, 0).Value = "Savings Projections Pre-Retirement" 'Savings projections while working
    Selection.Offset(10, 0).Style = "Heading 3"
    
    Range("A:A").Font.Bold = True 'Bold font
    
    'Call the Solver macro to solve for nominal monthly payment amount
    callSolver
        
    'Visual cue for difference between target and Solver solution
    solverError = ActiveCell.Offset(6, 1).Value - desiredInheritance
    ActiveCell.Offset(6, 2).Value = solverError
    ActiveCell.Offset(6, 2).Font.Bold = True 'Bold font
    ActiveCell.Offset(6, 2).NumberFormat = "$#,##0.00" 'Currency style
    If solverError < 0 Then
        ActiveCell.Offset(6, 2).Font.Color = RGB(255, 0, 0) 'Red if negative
    Else
        ActiveCell.Offset(6, 2).Font.Color = RGB(0, 255, 0) 'Green if positive
    End If
End Sub
'***************************************************************
'Subroutine that automates the Solver add-in
'Called in the main subroutine above
'NOTE: SOLVER REFERENCES ARE HARD-CODED, SO CHANGE CELL LOCATIONS FOR FORMATTING CAREFULLY
'https://msdn.microsoft.com/en-us/library/office/ff839427.aspx?f=255&MSPPError=-2147217396
'***************************************************************
Sub callSolver()
    Application.Calculation = xlAutomatic
    desiredInheritance = Range("B9").Value 'Extract desired real inheritance
    Worksheets("Retirement Savings").Activate
    SolverReset 'Reset Solver info

    SolverOptions precision:=0.0001 'Low error tolerance because I hate Solver and I want to make it work
    'SolverOk SetCell:=Range("B18") 'Set objective cell
    'SolverOk MaxMinVal:=3 'Indicate that Solver needs to match a specific value
    'SolverOK ValueOf:=desiredInheritance 'Need to match desired inheritance
    'SolverOk ValueOf:="0"
    'SolverOk ByChange:=Range("B16") 'Change proposed withdrawal value to match target
    'SolverSolve UserFinish:=False 'Don't show a dialog box, just show the result
    'SolverSolve 'Run Solver
    SolverOk SetCell:="$B$18", MaxMinVal:=3, ValueOf:=desiredInheritance, ByChange:="$B$16", _
        Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverSolve UserFinish:=False
End Sub
'***************************************************************
'Determines the nominal impact of a given monthly withdrawal on savings
'Depends on the initial savings, amount of money withdrawn, life expectancy (years), expected rate of return, and inheritance
'***************************************************************
Function nominalMonthlyWithdrawal(savings, withdrawal, lifeExpectancy, returnOnSavings)
    'Convert annual figures to monthly figures
    monthsRemaining = lifeExpectancy * 12
    monthlyReturn = ((1 + returnOnSavings) ^ (1 / 12)) - 1
    For currentMonth = 1 To monthsRemaining
        savings = savings * (1 + monthlyReturn) 'Add expected return
        savings = savings - withdrawal 'Remove withdrawal
    Next currentMonth
        
    nominalMonthlyWithdrawal = savings
End Function
'***************************************************************
'Computes year over year nominal savings
'Prints out evolution of savings account over time
'***************************************************************
Function nominalSavings(savings, income, incomeGrowth, incomeSavingRate, returnOnSavings, inflationRate, yearsRemaining)
    
    For currentYear = 1 To yearsRemaining 'Iterate through expected earnings years
        savings = savings * (1 + returnOnSavings) 'Add return on savings
        savings = savings + income * incomeSavingRate 'Add savings from annual income
        realSavings = nominalToReal(savings, inflationRate, currentYear) 'Compute savings in real terms
        income = income * (1 + incomeGrowth) 'Grow income for next year
        
        'Print intermediate results
        'Should use subroutine, but for such a simple program, my convenience > perfect design
        Selection.Offset(11, currentYear).Value = currentYear
        Selection.Offset(12, currentYear).Value = savings
        Selection.Offset(13, currentYear).Value = realSavings
        
        Selection.Offset(12, currentYear).NumberFormat = "$#,##0.00"
        Selection.Offset(13, currentYear).NumberFormat = "$#,##0.00"
    Next currentYear
    
    'Some more printing
    Selection.Offset(11, 0).Value = "Year"
    Selection.Offset(12, 0).Value = "Nominal savings"
    Selection.Offset(13, 0).Value = "Real savings"
    nominalSavings = savings 'Return nominal savings
End Function
'***************************************************************
'Discounts nominal savings to real savings by using expected inflation rate
'Result should be interpreted as the purchasing power available in given year stated in today's terms
'***************************************************************
Function nominalToReal(nominalSavings, inflationRate, years)
    discountFactor = (1 + inflationRate) ^ years
    realSavings = nominalSavings / discountFactor
    nominalToReal = realSavings 'Return real savings
    'Debug.Print "Discount factor: "; discountFactor
End Function
