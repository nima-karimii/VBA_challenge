Attribute VB_Name = "Module1"
Sub Report()


Dim i, j, k As Integer           ' i,j,k are cunters

Dim Tabalecunt As Integer
tablecunt = 1
 
Dim StockTotal, PersentsgeChange, PriceChange, CountRows, FirsPrice, LastPrice As Double
' StockTotal: Total stock Volume calculater
' PersentsgeChange :The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' PriceChange : Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' CountRows : The Number of the rows in main tabale
' FirsPrice :opening price at the beginning of the year
' LastPrice :the closing price at the end of the year

CountRows = Cells(Rows.Count, 1).End(xlUp).Row


k = 2                        ' The result tabale starts from row 2


' Desiging the result tabale

Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"
Range("J1:M1").ShrinkToFit = True




StocktTotal = 0             ' Prevalue of StockTotal is 0


Firstprice = Cells(2, 3).Value


' Starting the loop to cover all main Data


For i = 2 To CountRows


If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then    ' Finding the chanaging poit
 
  Cells(k, 10).Value = Cells(i, 1).Value                ' Ticker in result table
  
  LastPrice = Cells(i, 6).Value                        ' The last price of the year
  
  StockTotal = StockTotal + Cells(i, 7).Value          ' Total of The Stock value
  
  PriceChange = LastPrice - Firstprice                 ' Yearly price change  calculation
  
  If Firstprice <> 0 Then                               ' Persent yearly change calculation , Using if to skip 0 diviosion
     PersentsgeChange = (LastPrice - Firstprice) / Firstprice
    Else: PersentsgeChange = 0
  End If

  
  Cells(k, 11).Value = PriceChange                      'Yearly Price change in result tabale
  
  Cells(k, 12).Value = PersentsgeChange               'Percent price change in result table
    Cells(k, 12).NumberFormat = "%00"                   'Formating to "%"
    
  Cells(k, 13).Value = StockTotal                      'Total stock Volume in result tabale
  
 ' Conditional Formating for Yearly Price change
  If Cells(k, 11).Value > 0 Then
    Cells(k, 11).Interior.ColorIndex = 3 'Green
        ElseIf Cells(k, 11).Value < 0 Then Cells(k, 11).Interior.ColorIndex = 4 ' RED
    End If
    
   
  StockTotal = 0                                       'Reset the Total Stocke for next Ticker
   
   Firstprice = Cells(i + 1, 3).Value                   ' Opening price for New Ticker
   
    k = k + 1                                          ' Next row in result table
   
   
Else                                                    ' if the Ticker does not change we need to sum up the volume and finding the Last price
   
   
   LastPrice = Cells(i, 6).Value
  StockTotal = StockTotal + Cells(i, 7).Value
  
  
 End If


Next i      ' Next row in main table


'Call DesigningTheResult


End Sub ' End of Bulding the result tabale


Sub DesigningTheResult()

    Dim Tabalecunt As Integer
    Tabalecunt = 1
    
    'Range("J1").Select
    ' ActiveSheet.ListObjects.Add(xlSrcRange, Range("J:M"), , xlYes).TableStyle = "TableStyleLight8"
    
    'Columns("J:M").Select
    'Range("J:M").TableStyle = "TableStyleMedium8"
    
   

    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$J:$M"), , xlYes).Name = _
        "Table" & "i"
    Columns("J:M").Select
    ActiveSheet.ListObjects("Table" & "i").TableStyle = "TableStyleMedium8"
    
    Tabalecunt = Tabalecunt + 1
    
    
    
    
    
End Sub

'###########################################################################################################

Sub Summery()                                           'Creating the summery table


Dim i, CountRows As Integer
Dim Maxvalue, Minvalue, MaxTotal, Mintotal As Double
Dim MaxTicker, MinTicker, MaxTotalTicker As String


CountRows = Cells(Rows.Count, 12).End(xlUp).Row

' prevalues for variables to start a search

MaxTicker = ""
MinTicker = ""
Maxvalue = Cells(2, 12).Value
Minvalue = Cells(2, 12).Value
MaxTotal = Cells(2, 13).Value

' Beaning of the loop for searching the minimum and Maximum and Total

For i = 3 To CountRows

If Cells(i, 12).Value > Maxvalue Then
Maxvalue = Cells(i, 12).Value
MaxTicker = Cells(i, 10).Value
' MsgBox ("MAX : " & MaxTicker)

 ElseIf Cells(i, 12).Value < Minvalue Then
Minvalue = Cells(i, 12).Value
MinTicker = Cells(i, 10).Value
' MsgBox ("Min:" & MinTicker)

End If

If Cells(i, 13).Value2 > MaxTotal Then
MaxTotal = Cells(i, 13).Value
MaxTotalTicker = Cells(i, 10).Value

' MsgBox (MaxTotalTicker)

End If


Next i

' showing the results in spearshet cells

Range("O1").Value = "..."
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Cells(2, 15).Value = " Greatest % Increase "
Cells(2, 16).Value = MaxTicker
Cells(2, 17).Value = Maxvalue
Cells(2, 17).NumberFormat = "%00"
Cells(3, 15).Value = " Greatest % Decrease "
Cells(3, 16).Value = MinTicker
Cells(3, 17).Value = Minvalue
Cells(3, 17).NumberFormat = "%00"
Cells(4, 15).Value = " Greatest Total Volume "
Cells(4, 16).Value = MaxTotalTicker
Cells(4, 17).Value = MaxTotal

Columns("O:Q").EntireColumn.AutoFit

End Sub

'###########################################################################################################
Sub All_Sheet()

Dim ws As Worksheet

Dim starting_ws As Worksheet

Dim Tabalecunt As Integer

Tabalecunt = 1
    
  
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
   Call Report
    Call Summery
    
   ' Designing the reports as a tabales
   
  ActiveSheet.ListObjects.Add(xlSrcRange, Range("$J:$M"), , xlYes).Name = _
        "Table" & "i"
    Columns("J:M").Select
    ActiveSheet.ListObjects("Table" & "i").TableStyle = "TableStyleMedium8"
    
  ' Designing the Summery as a tabales
    
  ActiveSheet.ListObjects.Add(xlSrcRange, Range("$o1:$q4"), , xlYes).Name = _
        "TableS" & "i"
    Columns("J:M").Select
    ActiveSheet.ListObjects("TableS" & "i").TableStyle = "TableStyleMedium10"
    
    
    
    Tabalecunt = Tabalecunt + 1
    
       
Next

starting_ws.Activate 'activate the worksheet that was originally active


End Sub

'###########################################################################################################

Sub ClearAllWorkshhet() ' Clearing all worksheets

Dim ws As Worksheet
Dim starting_ws As Worksheet

Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
   
   Call DeletingColums
   
   Range("j:R").Clear
   
    
Next

starting_ws.Activate 'activate the worksheet that was originally active


End Sub

'###########################################################################################################

Sub DeletingColums()             ' Deleting previous report tabales

   If (Cells(1, 10) <> "") Then        ' just cheking is there a report on that sheet
    Columns("J:M").Select
    
    Selection.ListObject.ListColumns(1).Delete ' Deleting the coulums
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    Selection.ListObject.ListColumns(1).Delete
    End If
End Sub



