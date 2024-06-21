Attribute VB_Name = "S_Behaivor"
Sub Stock_Stats()
  Dim WsName, Ticker, SGIn, SGDe, SGVl As String
  Dim Ws As Worksheet
  ' Set an initial variable for holding the top/bottom value of serie
  Dim OpenV, CloseV, QtCh, PcCh, GtIn, GtDe As Double


  ' Keep track of the location for each seir in the summary table
  Dim S_Table, SumTbl, BeginSerie As Integer
  Dim TotalRows, TotalRS, Counter As Long
  Dim TotalVS, GtVs As Variant
  
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    For Each Ws In Worksheets 'Recorre las hojas del libro
    WsName = Ws.Name
      Ws.Columns("H:T").Delete
      'Ws.Range("A1").Select
      If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
        ' Print the titles for summary
        Ws.Range("I" & 1).Value = "Ticker"
        Ws.Range("J" & 1).Value = "Quarterly Change"
        Ws.Range("K" & 1).Value = "Percent Change"
        Ws.Range("L" & 1).Value = "Total Stock Volume"
        Ws.Range("P" & 1).Value = "Ticker"
        Ws.Range("Q" & 1).Value = "Value"
        Ws.Range("O" & 2).Value = "Greatest % Increase"
        Ws.Range("O" & 3).Value = "Greatest % Decrease"
        Ws.Range("O" & 4).Value = "Greatest Total Volume"
    Next Ws
  
  For Each Ws In Worksheets
  WsName = Ws.Name
  S_Table = 2
  TotalRows = Ws.Cells(Rows.Count, 1).End(xlUp).Row
  BeginSerie = 0
  TotalVS = 0
  
  ' Loop through all tickers
  For Counter = 2 To TotalRows
    ' Check if is the begining of the serie, if it is not...
    If Ws.Cells(Counter + 1, 1).Value = Ws.Cells(Counter, 1).Value Then
        If BeginSerie = 0 Then
            BeginSerie = 1
          ' Set the ticker name
            Ticker = Ws.Cells(Counter, 1).Value
            OpenV = Ws.Cells(Counter, 3).Value
            'Ws.Cells(Counter, 8).Value = "Begin"
        End If
        TotalVS = TotalVS + Ws.Cells(Counter, 7).Value
    Else
     'Ws.Cells(Counter, 8).Value = "End"
      ' Get close value
      CloseV = Ws.Cells(Counter, 6).Value
      ' Get quarterly change value
      QtCh = CloseV - OpenV
      ' Get quarterly percent change value
      PcCh = QtCh / OpenV
        TotalVS = TotalVS + Ws.Cells(Counter, 7).Value
   
      ' Print the ticker serie in the Summary Table
      Ws.Range("I" & S_Table).Value = Ticker
      ' Print the Quarterly change to the Summary Table
      Ws.Range("J" & S_Table).Value = QtCh
      ' Print the Quarterly change to the Summary Table
      Ws.Range("K" & S_Table).Value = PcCh
      ' Print the Quarterly change to the Summary Table
      Ws.Range("L" & S_Table).Value = TotalVS

        Ws.Columns("I:Q").EntireColumn.AutoFit
      ' Reset the Brand Total
      S_Table = S_Table + 1
      TotalVS = 0
      BeginSerie = 0
    End If
  Next Counter
    
    
    'Get stats asked
    
    GtIn = Application.WorksheetFunction.Max(Ws.Range("K2:K" & S_Table), 0)
    GtDe = Application.WorksheetFunction.Min(Ws.Range("K2:K" & S_Table), 0)
    GtVs = Application.WorksheetFunction.Max(Ws.Range("L2:L" & S_Table), 0)
    
    'Get ticker's from stats found
    TotalRS = Ws.Cells(Rows.Count, 9).End(xlUp).Row
    For SumTbl = 2 To TotalRS
    If GtIn = Ws.Cells(SumTbl, 11).Value Then
    SGIn = Ws.Cells(SumTbl, 9).Value
    ElseIf GtDe = Ws.Cells(SumTbl, 11).Value Then
    SGDe = Ws.Cells(SumTbl, 9).Value
    ElseIf GtVs = Ws.Cells(SumTbl, 12).Value Then
    SGVl = Ws.Cells(SumTbl, 9).Value
    End If
    Next SumTbl
    
    'Print ticker names
    Ws.Range("P2").Value = SGIn
    Ws.Range("P3").Value = SGDe
    Ws.Range("P4").Value = SGVl
    
    'Print stat values
    Ws.Range("Q2").Value = GtIn
    Ws.Range("Q3").Value = GtDe
    Ws.Range("Q4").Value = GtVs
    Ws.Range("J2:J" & TotalRS).NumberFormat = "0.00"
    Ws.Range("K2:K" & TotalRS).NumberFormat = "0.00%"
    Ws.Range("Q2:Q3").NumberFormat = "0.00%"
    Ws.Range("Q4").NumberFormat = "0"
    Ws.Columns("I:Q").AutoFit
        With Ws.Range("J2:J" & TotalRS)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=J2>0"
            With .FormatConditions(.FormatConditions.Count).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 3407769
                .TintAndShade = 0
            End With
            .FormatConditions(.FormatConditions.Count).StopIfTrue = False
            .FormatConditions.Add Type:=xlExpression, Formula1:="=J2<0"
            With .FormatConditions(.FormatConditions.Count).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
            End With
            .FormatConditions(.FormatConditions.Count).StopIfTrue = False
        End With
    
        With Ws.Range("K2:K" & TotalRS)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlExpression, Formula1:="=K2>0"
            With .FormatConditions(.FormatConditions.Count).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 3407769
                .TintAndShade = 0
            End With
            .FormatConditions(.FormatConditions.Count).StopIfTrue = False
            .FormatConditions.Add Type:=xlExpression, Formula1:="=K2<0"
            With .FormatConditions(.FormatConditions.Count).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
            End With
            .FormatConditions(.FormatConditions.Count).StopIfTrue = False
        End With
    
    
    Next Ws
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub


