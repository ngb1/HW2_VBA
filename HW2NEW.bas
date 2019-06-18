Attribute VB_Name = "HW2"
Sub MasterCode()
'entered elapsed time calculator code from internet
    'PURPOSE: Determine how many minutes it took for code to completely run
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim StartTime As Double
Dim MinutesElapsed As String
Dim WS_Count As Integer
Dim i As Long
'Remember time when macro starts
  StartTime = Timer

'*****************************
'used the code shown by Jeff here: https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For i = 1 To WS_Count
         
         ActiveWorkbook.Worksheets(i).Activate
         'Call NewTicker
         'Call CountVolumePerTicker
         'Call OpenClose
         Call Headers
         Call CountVolumePerTickerIfOrdered
         Call ConditionalYearChange
         Call GreatestPercentIncrease
         Call GreatestPercentDecrease
         Call GreatestTotalVolume

         Next i

'*****************************

'Determine how many seconds code took to run
  MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")

'Notify user in seconds
  MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub


Sub Headers()

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Columns("I").ColumnWidth = 16
Columns("J").ColumnWidth = 16
Columns("K").ColumnWidth = 16
Columns("L").ColumnWidth = 16

End Sub

Sub CountVolumePerTickerIfOrdered()
Dim Summary_Table_Row As Long
Dim CountRefTable As Long
Dim TickerName As String
Dim TickerVolume As Double
Dim close_new As Double
Dim open_old As Double


CountRefTable = Range("A2").End(xlDown).Row 'count number of tickers on ref table
Summary_Table_Row = 2
TickerVolume = 0

For a = 2 To CountRefTable + 1

If Cells(a - 1, 1).Value <> Cells(a, 1).Value And TickerVolume = 0 Then
open_old = Cells(a, 3).Value                          'saves first open value of the same ticker
'Range("M" & Summary_Table_Row).Value = open_old
End If
If Cells(a + 1, 1).Value = Cells(a, 1).Value Then
    TickerVolume = TickerVolume + Cells(a, 7).Value     'adds up stock volumes of the same ticker
ElseIf Cells(a + 1, 1).Value <> Cells(a, 1).Value Then
    TickerName = Cells(a, 1).Value
    TickerVolume = TickerVolume + Cells(a, 7).Value     'adds up the last stock volume of the same ticker
    'If TickerVolume <> 0 Then
        close_new = Cells(a, 6).Value                   'saves last close value of the same ticker
        'Range("N" & Summary_Table_Row).Value = close_new
    'End If
    Range("I" & Summary_Table_Row).Value = TickerName
    Range("L" & Summary_Table_Row).Value = TickerVolume
    Range("J" & Summary_Table_Row).Value = close_new - open_old
    If open_old > 0 Then
        PercentChange = Round(((close_new - open_old) * 100 / open_old), 3)
        Range("K" & Summary_Table_Row).Value = PercentChange
    Else
        Range("K" & Summary_Table_Row).Value = "-0.00000001"
    End If
    'Range("N" & Summary_Table_Row).Value = close_new
    Summary_Table_Row = Summary_Table_Row + 1
    TickerVolume = 0
End If
    
Next a

End Sub



Sub ConditionalYearChange()

Dim j As Integer
Dim CountNewTable As Long


CountNewTable = Range("I2").End(xlDown).Row 'count number of tickers on new table
'MsgBox (CountNewTable)

For j = 2 To CountNewTable

    If Cells(j, 10).Value > 0 Then
    Cells(j, 10).Interior.ColorIndex = 4
    ElseIf Cells(j, 10).Value <= 0 Then
    Cells(j, 10).Interior.ColorIndex = 3
    End If
    
Next j

End Sub

Sub GreatestPercentIncrease()

Dim j As Integer
Dim CountNewTable As Long
Dim Ref As Double
Dim PercentMax As Double
Dim TickerMax As String

CountNewTable = Range("I2").End(xlDown).Row 'count number of tickers on new table

Ref = Cells(2, 11).Value

For j = 2 To CountNewTable
    If Cells(j, 11).Value >= Ref Then
    PercentMax = Cells(j, 11).Value
    TickerMax = Cells(j, 9).Value
    Ref = PercentMax
    End If

Next j
'formating

Columns("N").ColumnWidth = 19
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

'results
Range("N2").Value = "Greatest % Increase"
Range("O2").Value = TickerMax
Range("P2").Value = PercentMax


End Sub
    
Sub GreatestPercentDecrease()

Dim j As Integer
Dim CountNewTable As Long
Dim Ref As Double
Dim PercentMin As Double
Dim TickerMin As String

CountNewTable = Range("I2").End(xlDown).Row 'count number of tickers on new table

Ref = Cells(2, 11).Value

For j = 2 To CountNewTable
    If Cells(j, 11).Value <= Ref Then
    PercentMin = Cells(j, 11).Value
    TickerMin = Cells(j, 9).Value
    Ref = PercentMin
    End If

Next j

Range("N3").Value = "Greatest % Decrease"
Range("O3").Value = TickerMin
Range("P3").Value = PercentMin
    
End Sub
        
Sub GreatestTotalVolume()

Dim j As Integer
Dim CountNewTable As Long
Dim Ref As Double
Dim VolumeMax As Double
Dim TickerMax As String

CountNewTable = Range("I2").End(xlDown).Row 'count number of tickers on new table

Ref = Cells(2, 12).Value

For j = 2 To CountNewTable
    If Cells(j, 12).Value >= Ref Then
    VolumeMax = Cells(j, 12).Value
    TickerMax = Cells(j, 9).Value
    Ref = VolumeMax
    End If

Next j

Range("N4").Value = "Greatest Total Volume"
Range("O4").Value = TickerMax
Range("P4").Value = VolumeMax
Columns("P").ColumnWidth = 16
    
End Sub

'### ATTEMPTED CODES ASSUMING THE DATA IS NOT SORTED ALPHABETICALLY,
'### THESE CODES WORK BUT NESTED LOOPS TAKE TOO LONG, KEEP FOR FUTURE REFERENCE ###

Sub OLDNewTicker()
Dim CountRefTable
Dim i As Long
Dim FirstI As String


CountRefTable = Range("A2").End(xlDown).Row 'count number of tickers on ref table
'MsgBox (CountA)

'Create table with tickers,
'Create reference for the first value to compare Tickers between new table with ref table
Range("I2").Value = Range("A2").Value
'MsgBox (Range("A2").Value)
'Enter header name
Range("I1").Value = "Ticker"

'create new table summarizing tickers
Dim a As Long
i = 2  'start the reference of i outside of the loop, i is not a loop, it's a counter
For a = 3 To CountRefTable + 1
    If Cells(i, 9) <> Cells(a, 1) Then 'if cell on new table <> cell on ref table
        i = i + 1 'go to next cell on new table
    Cells(i, 9).Value = Cells(a, 1) 'enter this name on the new table
    End If
Next a

End Sub
        
        
        
Sub OLDOpenClose()

Dim datemax As Long
Dim datemin As Long
Dim a As Long
Dim i As Long
Dim CountRefTable As Long
Dim CountNewTable As Long
Dim date_new As Long
Dim date_old As Long
Dim open_new As Double
Dim close_new As Double
Dim open_old As Double
Dim close_old As Double
Dim date_in As Double
Dim open_in As Double
Dim close_in As Double


CountRefTable = Range("A2").End(xlDown).Row 'count number of tickers on ref table
'MsgBox (CountA)

CountNewTable = Range("I2").End(xlDown).Row 'count number of tickers on new table
'MsgBox (CountNewTable)

'Enter header name



For i = 2 To CountNewTable  '289
    'assign incoming variables values if tickers match
    
date_old = Range("B2").Value
'MsgBox (date_old)
    For a = 2 To CountRefTable
        If Cells(a, 1).Value = Cells(i, 9).Value Then
            date_in = Cells(a, 2).Value
            'MsgBox ("date_in=" & date_in)
            open_in = Cells(a, 3).Value
            'MsgBox ("open_in=" & open_in)
            close_in = Cells(a, 6).Value
            'MsgBox (close_new)
        End If
        'determine if incoming values are for early date (open_new) or last date (close_old)
        If date_in > date_old Then
            close_new = close_in
            open_new = open_in
            date_new = date_in
           ' MsgBox (close_old)
            
        ElseIf date_in <= date_old Then
            open_old = open_in
            'MsgBox (open_old)
            date_old = date_in
            'MsgBox ("date_old=" & date_old)
                   
        End If
        'values assigned, go to next

    Next a
'date_old = Range("B2").Value
    Cells(i, 11).Value = close_new - open_old
    Cells(i, 12).Value = Round(((close_new - open_old) * 100 / open_old), 3)
Next i
        'MsgBox (open_old)
        'MsgBox (close_new)
    
    'Cells(i - 1, 11).Value = close_new - open_old
    'Cells(i - 1, 12).Value = Round(((close_new - open_old) * 100 / open_old), 3)
    'Cells(i - 1, 13).Value = open_old
    'Cells(i - 1, 14).Value = close_new
End Sub

Sub OLDCountVolumePerTicker()
Dim i As Long
Dim j As Long
Dim jold As Double
Dim jnew As Double
Dim CountNewTable
Dim CountRefTable

CountRefTable = Range("A2").End(xlDown).Row 'count number of tickers on ref table
'MsgBox (CountA)

CountNewTable = Range("I2").End(xlDown).Row 'count number of tickers on new table
'MsgBox (CountNewTable)

'Enter header name
Range("J1").Value = "Total Stock Volume"

For i = 2 To CountNewTable

jold = 0 'resets jold
    For a = 2 To CountRefTable + 1 'starts loop on ref table
        If Cells(i, 9).Value = Cells(a, 1).Value Then 'if tickers match then
            Cells(i, 10).Value = jold 'assigns the cell to jold
            'MsgBox ("start_jold=" & jold)
            jnew = Cells(a, 7).Value 'sets cell on ref table as jnew
            'MsgBox ("jnew=" & jnew)
            jold = jnew + jold 'sums up jold and jnew as the new jold
            'MsgBox ("end_jold=" & jold)
            Cells(i, 10).Value = jold 'sets jold as value on cell
        End If
        
    Next a
    'jold = 0
Next i


End Sub



Sub OLDMasterCode2()



         ' Declare Current as a worksheet object variable.
         Dim Current As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each Current In Worksheets

            ' Insert your code here.
            ' This line displays the worksheet name in a message box.
            MsgBox Current.Name
         Next

      
End Sub
