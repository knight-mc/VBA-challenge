Attribute VB_Name = "Module1"
Sub StockCalc():

Dim ticker As String
Dim tickHold As String
Dim earliest As Long
Dim latest As Long
Dim openV As Double
Dim closeV As Double
Dim holdDate As Long
Dim volume As Double
Dim changeY As Double
Dim changeP As Double
Dim lastRow1 As Long
Dim count1 As Long
Dim count2 As Long


Dim greatINC As Double
Dim greatDEC As Double
Dim greatVOL As Double
Dim tickINC As String
Dim tickDEC As String
Dim tickVOL As String

Dim ws As Worksheet

For Each ws In Worksheets

    Range("A1", Range("G1").End(xlDown)).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
    lastRow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row 'works
    'MsgBox (lastRow1)

    ticker = "HOLD"
    earliest = 0
    latest = 0
    openV = 0
    closeV = 0
    count2 = 2

    greatINC = 0
    greatDEC = 0
    greatVOL = 0

    ws.Cells(1, 11).Value = "Ticker"                    'works
    ws.Cells(1, 12).Value = "Yearly Change"             'works
    ws.Cells(1, 13).Value = "Percent Change"            'works
    ws.Cells(1, 14).Value = "Total Stock Volume"        'works

    ws.Cells(1, 18).Value = "Ticker"                    'works
    ws.Cells(1, 19).Value = "Value"                     'works
    ws.Cells(2, 17).Value = "Greatest % Increase"       'works
    ws.Cells(3, 17).Value = "Greatest % Decrease"       'works
    ws.Cells(4, 17).Value = "Greatest Total Volume"     'works

    ws.Columns("A:S").AutoFit                           'works
    
    For count1 = 2 To (lastRow1 + 1)
        tickHold = ws.Cells(count1, 1).Value
        If tickHold = ticker Then
      ' Check for first and last dates to pull starting and ending values, sum volume value
      ' ================================================================================
            holdDate = ws.Cells(count1, 2).Value
            If holdDate < earliest Then
                openV = ws.Cells(count1, 3).Value
                earliest = ws.Cells(count1, 2).Value
            End If
        
            If holdDate > latest Then
                closeV = ws.Cells(count1, 6).Value
                latest = ws.Cells(count1, 2).Value
            End If
        
            volume = volume + ws.Cells(count1, 7).Value
        
        Else
      ' Ticker Close Out Processes
      ' ================================================================================
        
      ' Input Ticker Computation Information
      ' ==========================================================================
            If count2 > 2 Then
                changeY = Round((closeV - openV), 2)
                changeP = Round(((changeY / openV) * 100), 2)
                ws.Cells((count2 - 1), 12).Value = changeY
                ws.Cells((count2 - 1), 13).Value = (Str(changeP) + "%")
                ws.Cells((count2 - 1), 14).Value = volume
                If ws.Cells((count2 - 1), 12).Value < 0 Then
                    ws.Cells((count2 - 1), 12).Interior.ColorIndex = 3
                    ws.Cells((count2 - 1), 13).Interior.ColorIndex = 3
                ElseIf ws.Cells((count2 - 1), 12).Value > 0 Then
                    ws.Cells((count2 - 1), 12).Interior.ColorIndex = 4
                    ws.Cells((count2 - 1), 13).Interior.ColorIndex = 4
                Else
                    ws.Cells((count2 - 1), 12).Interior.ColorIndex = 6
                    ws.Cells((count2 - 1), 13).Interior.ColorIndex = 6
                End If
'        Else
            End If
        
        
      ' Greatest Changes Check
      ' ==========================================================================
            If changeP > greatINC Then
                greatINC = changeP
                tickINC = ticker
            End If
        
            If changeP < greatDEC Then
                greatDEC = changeP
                tickDEC = ticker
            End If
        
            If volume > greatVOL Then
                greatVOL = volume
                tickVOL = ticker
            End If
  
     ' New Ticker Start Up Processes (Resets)
     ' ===================================================================================
            
        ticker = ws.Cells(count1, 1).Value
        ws.Cells(count2, 11).Value = CStr(ticker)
        
        count2 = count2 + 1
        
        earliest = ws.Cells(count1, 2).Value
        latest = ws.Cells(count1, 2).Value
        openV = ws.Cells(count1, 3).Value
        closeV = ws.Cells(count1, 6).Value
                
        
        
        End If
        
    Next count1
    
' Input Greatest Values
'===============================================================================
    
    ws.Cells(2, 18).Value = tickINC
    ws.Cells(3, 18).Value = tickDEC
    ws.Cells(4, 18).Value = tickVOL
    
    ws.Cells(2, 19).Value = (Str(greatINC) + "%")          'works
    ws.Cells(3, 19).Value = (Str(greatDEC) + "%")          'works
    ws.Cells(4, 19).Value = greatVOL                       'works
    
    ws.Columns("A:S").AutoFit                              'should work, same statement works earlier in document

Next ws

End Sub

