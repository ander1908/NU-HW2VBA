Attribute VB_Name = "Module1"
Sub VBA_Wall_Street()

'Steps and Process breakdown
'-------------------------------------------------------------------------
'0) Defitions and Variables
Dim Ticker As String
Dim Percent_Change As Double
Dim ENDV As Double
Dim STARTV As Double
Dim VOLUME As LongLong
Dim TS As Integer



'--------------------------------------------------------------------------
'1) Ticker Symbol Pull
    '1a - Use Last Row to determine length of the conditional loop
    '1b - Define Location of Ticker List
    '1c - Create criteria for Extraction of Ticker
        '1c1 - If next value <> current value, then PRINT current_value to Ticker ListLocation
            '1c1a - SEE 3a BELOW for building Sum Values
        '1c2 - Ticker List Location + 1 to keep list moving
        
        
'2) Yearly Change from opening price at the beginning of a given year to close
    '2a Using Unix Code find max and min values that denote the earlier and latest dates in year
        '2a1 Create Conditional Loop to refresh Max and Min Values and Print the difference
        '2a2 Assign Max + Min to Variables ENDV and STARTV
        '2a3 Percent Change = (ENDV - STARTV) / STARTV
        '2a4 Print Percent Change to Corresponding Location
            'HOW TO FIND THAT LOCATION?
     '2b CONDITIONAL FORMATTING DONT FORGET
        
'3) Total Stock Volume of the Stock
    '3a In previous For Loop Create a Value to store all of the volume
        '3a1 SVolume = sum of the offset value
        '3a2 When the Value of the ticker Changes print the SVolume then reset it to 0
        
' In order to do this all in one swoop we will need to put all of these into the for loop
'--------------------------------------------------------------------------

'Summarized on each sheet
'SCRAP LINES --- USE LATER TO BUILD COMPILATION SHEET

    'Add Combined Data Sheet to Organize Data
    'Sheets.Add.Name = "Summary_Sheet"
    'Move Sheet to head of WorkBook
    'Sheets("Summary_Sheet").Move Before:=Sheets(1)

    'Cycle through all sheets
    'Set ALL_Sheets = Worksheets("Summary_Sheet")


    For Each ws In Worksheets


    'Start For Loop for Data Set

        'Setting All Initial Values (protect against runover data from last loop)
        VOLUME = 0
        ENDV = 0
        STARTV = 0
        Percent_Change = 0
        'Setting Ticker Symbol Space to start at 2nd Row
        TS = 2
        'Creating Label Row
        ws.Cells(1, 9).Value = "Ticker List"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Determine Last Row Using Class Formula
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'For Loop for pulling information from WS
        For i = 2 To LastRow
            
            'To identify the Start (value in A Col is =/= to prior)
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                'Set the Start Value using Open Value
                STARTV = ws.Cells(i, 3).Value
                'Add Volume to running Total as this will be first entry
                VOLUME = VOLUME + ws.Cells(i, 7).Value
                
                'Testing Script
                'MsgBox ("TEST BOX to show If working")
                'MsgBox (Str(Cells(i, 1).Value))
                
                
                
            'To identify the end (value in A Col is =/= next)
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Set the Value of A:A (TicketNAME) to TS Location
                ws.Cells(TS, 9).Value = ws.Cells(i, 1).Value
                
                '------------------------------------------
                
                'Set the Value of ENDV for the stock price
                ENDV = ws.Cells(i, 6).Value
                'Set Value for Yearly Change
                ws.Cells(TS, 10).Value = ENDV - STARTV
                    If ws.Cells(TS, 10).Value < 0 Then
                        ws.Cells(TS, 10).Interior.ColorIndex = 3
                    ElseIf ws.Cells(TS, 10).Value > 0 Then
                        ws.Cells(TS, 10).Interior.ColorIndex = 4
                    End If
                                    
                'Calculate Percent change
                If STARTV = 0 Then
                    Percent_Change = "0"
                Else
                
                Percent_Change = ((ENDV - STARTV) / STARTV)
                End If
                
                'push value into corresponding cell to match sticker
                ws.Cells(TS, 11).Value = Percent_Change
                'https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
                ws.Cells(TS, 11).NumberFormat = "0.00%"
                
                    
                '------------------------------------------
                
                'Add Corresponding Volume cell to Volume Total
                VOLUME = VOLUME + ws.Cells(i, 7).Value
                'Push Volume Value into corresponding Cell
                ws.Cells(TS, 12).Value = VOLUME
                '------------------------------------------
                
                'Time to clean this house for next cycle
                VOLUME = 0
                ENDV = 0
                STARTV = 0
                Percent_Change = 0
                
                'Set the TS Value + 1 to add next row
                TS = TS + 1
                                
            
            
            Else
            'If it is not the first or last cell, than only the volume is needed to be collected
            'Add volume to total
                VOLUME = VOLUME + ws.Cells(i, 7).Value
            
            End If
            
            Next i

        '-----------------------------------------------------------
        'Bonus Portion
        'Establish new limits for Last Row based on Unique Tickers
        LastRowAnalysis = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'Search Unique Ticker Analysis for Greatest Vlaue
        'Establish Offset Start
        Dim BStart As Integer
        BStart = 2
        'Print Out labels
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker Symbol"
        ws.Cells(1, 17).Value = "Values"
        
        'Dim out new Variables for Bonus
        Dim GI As Double
        Dim GD As Double
        Dim GTV As LongLong
       'Assign Greatest Increase/Decrease/TotalVolume
        GD = 0
        GTV = 0
        GI = 0
        'Search the TickerList using the new Last Row Analysis
       For j = 2 To LastRowAnalysis
            'Identify Greatest Value
            If ws.Cells(j, 11).Value > GI Then
                GI = ws.Cells(j, 11).Value
                ws.Cells(BStart, 17).Value = GI
                'https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
                'Making the output into percentage
                ws.Cells(BStart, 17).NumberFormat = "0.00%"
                ws.Cells(BStart, 16).Value = ws.Cells(j, 9).Value
            'Identify Lowest Value
            ElseIf Cells(j, 11).Value < GD Then
                GD = ws.Cells(j, 11).Value
                ws.Cells(BStart + 1, 17).Value = GD
                'https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage
                'Making the output into percentage
                ws.Cells(BStart + 1, 17).NumberFormat = "0.00%"
                
                
                ws.Cells(BStart + 1, 16).Value = ws.Cells(j, 9).Value
            End If
            'New If to search the Total Volumes
            If ws.Cells(j, 12).Value > GTV Then
                GTV = ws.Cells(j, 12).Value
                ws.Cells(BStart + 2, 17).Value = GTV
                ws.Cells(BStart + 2, 16).Value = ws.Cells(j, 9).Value
                End If
                
            Next j
            
            
    Next ws


End Sub


