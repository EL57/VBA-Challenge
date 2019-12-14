Attribute VB_Name = "HomeWork_Complete"
Sub HomeWork()

'Variable for the worksheet loop
Dim ws As Worksheet

'Loop through the worksheets
For Each ws In Worksheets

'Activate the worksheets for the loop to work
    ws.Activate

'Variable to hold the ticker symbol
Dim tick As String

'Variables to hold the Largest and the Smallest dates
Dim Ldate As Variant
Dim Sdate As Variant
Dim Cudate As Variant

'Variables to hold the Opening and the Closing price
Dim Oprice As Double
Dim Cprice As Double

'Variable to hold the Total Stock Volume
Dim TSV As Double

'Variable for row
Dim i As Long

'Variable to hold ticker symbol column
Dim Sumrow As Long

'Variables to hold the headers
Dim Header1 As String
Dim Header2 As String
Dim Header3 As String
Dim Header4 As String

'Creating the headers for the summary row
Header1 = "Ticker"
Header2 = "Yearly Change"
Header3 = "Percent Change"
Header4 = "Total Stock Volume"




'Formulas:
'Yearly Change = Closing Price - Opening Price
'Percent Change = (Closing Price / Opening Price) - 1

'Conditional format for the yearly change
'Range("J:J").Select
Range(Range("J2"), Range("J2").End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

'Formatting the percent change column to percent
ActiveSheet.Columns("K").NumberFormat = "0.00%"

'Formatting result area if the headers are blank
If Range("I1").Value = "" Then
    Range("I1").Value = Header1
End If

If Range("J1").Value = "" Then
    Range("J1").Value = Header2
End If

If Range("K1").Value = "" Then
    Range("K1").Value = Header3
End If

If Range("L1").Value = "" Then
    Range("L1").Value = Header4
End If
  
'Setting variable i to 2 so it can start after the header
i = 2
   
'Setting variable Sumrow to 2 for the summary table
Sumrow = 2



    'Using a Do Until loop that stops the loop when the first column is blank. Convient
    Do Until Cells(i, 1).Value = ""
     
     If Cells(i, 2).Value = Cells(2, 2).Value Then
    
            Sdate = Cells(i, 2).Value
            Oprice = Cells(i, 3).Value
        
      End If
    ' Check if we are still within the same ticker if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    

         
        
    ' Set the ticker name
    tick = Cells(i, 1).Value
    
      
      ' Add to the ticker Total
      TSV = TSV + Cells(i, 7).Value
      
      
      
      
        
        
            
      ' Print the ticker name in the Summary Table
      Range("I" & Sumrow).Value = tick
      
      ' Print the ticker total to the Summary Table
      Range("L" & Sumrow).Value = TSV
      
      'Print the Opening Price to the summary table
      Range("J" & Sumrow).Value = Cprice - Oprice


      'Print the Opening Price to the summary table
      
      If Oprice = 0 Then
      
        Range("K" & Sumrow).Value = "Can't compute"
      
            Else
      
                Range("K" & Sumrow).Value = (Cprice / Oprice) - 1
      
      End If
      
      ' Add one to the summary table row
      Sumrow = Sumrow + 1
      
      ' Reset the Ticker Total
      TSV = 0
      
            ' Next cell if the ticker is the same
            Else
    
                ' Add to the total of the ticker
                TSV = TSV + Cells(i, 7).Value
                

      
      If Cells(i + 1, 2).Value > Cells(i, 2).Value Then
            
            Ldate = Cells(i + 1, 2).Value
            Cprice = Cells(i + 1, 6).Value
        
        End If
        
          
    End If
    
   
   i = i + 1

    Loop
    
    
Next ws

    
End Sub
