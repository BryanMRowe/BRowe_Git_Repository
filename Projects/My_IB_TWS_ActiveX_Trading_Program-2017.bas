Attribute VB_Name = "My_Code"
Option Explicit

Dim OldWorkbook As Workbook
Dim NewWorkbook As Workbook
Dim AbortTime As Date
Dim r As Integer
Dim c As Integer
Dim ApproxAccountValue As Double
Dim LastDataRow As Long
Dim LastDataColumn As Long
Dim ExtractedDataHeight As Long
Dim LastAnalysisRow As Long
Dim NewLine As Long
Dim LastPreviousDayAnalysisRow As Long
Dim TopCurrentDayAnalysisRow As Long
Dim CurrentAnalysisRow As Long
Dim ParametersRange As Range
Dim AnalysisRange As Range
Dim RequestedTimeRange As Range
Dim RequestedNUGTRange As Range
Dim RequestedDUSTRange As Range
Dim CurrentMinuteTime As Range
Dim MostRecentNUGTPrice As Double
Dim MostRecentDUSTPrice As Double
Dim TradeBetSize As Double
Dim PreviousOpenMarketDate As String
Sub AutoAll()

    Application.OnTime TimeValue("09:00:00"), "ConnectToTWS"
    Application.OnTime TimeValue("09:01:00"), "AccountBalanceRequest"
    Application.OnTime TimeValue("09:02:00"), "FillBalanceAndRequestPreviousDay"
    Application.OnTime TimeValue("09:03:00"), "FillPreviousDay"
    Application.OnTime TimeValue("09:30:00"), "MarketOpen"

End Sub
Sub ConnectToTWS()

    'Call Sheets("General").ConnectToTWS_Click
    Application.Run "Sheet1.ConnectToTWS_Click"

End Sub
Sub AccountBalanceRequest()

    'request account updates
    Application.Run "Sheet10.RequestAccountUpdates_Click"

End Sub
Sub FillBalanceAndRequestPreviousDay()

    'updates account balance ONLY if account worksheet has been sucessfully subscribed to
    If Worksheets("Account").Range("A15") <> "" Then
        Call FillAccountBalance
    End If

    Call RequestPreviousDay

End Sub
Sub FillAccountBalance()

    'fill formula
    Worksheets("Parameters").Range("I2").value = Application.WorksheetFunction.VLookup("AvailableFunds", Sheets("Account").Range("a8:e207"), 5, False)
    ApproxAccountValue = Worksheets("Parameters").Range("I2").value
    Worksheets("Parameters").Range("J2").value = Application.WorksheetFunction.RoundDown((ApproxAccountValue * 0.24), -3)

    'clear account subscription request
    'Call Sheets("Account").cancelAcctsSubscription
    Application.Run "Sheet10.CancelAccountUpdates_Click"
    'Call Sheets("Account").clearAccts
    Application.Run "Sheet10.ClearAccountData_Click"

End Sub
Sub ClearOutOldData()

    'clear all market data
    Worksheets("Technical Analysis").Range("D80:D1000").Clear
    Worksheets("Technical Analysis").Range("H80:J1000").Clear
    Worksheets("Technical Analysis").Range("X80:Z1000").Clear

End Sub
Sub RequestPreviousDay()

    'clear all market data
    Worksheets("Technical Analysis").Range("D80:D1000").Clear
    Worksheets("Technical Analysis").Range("H80:J1000").Clear
    Worksheets("Technical Analysis").Range("X80:Z1000").Clear
    
    'fills user name
    Worksheets("Historical Data").Range("D5").value = "my_username"

    'clears any data requests for nugt & dust, preps arguments for running data request
    Worksheets("Historical Data").Range("A12").value = "NUGT"
    Worksheets("Historical Data").Range("A13").value = "DUST"
    Worksheets("Historical Data").Range("B12:B13").value = "STK"
    Worksheets("Historical Data").Range("H12:H13").value = "SMART"
    Worksheets("Historical Data").Range("J12:J13").value = "USD"
    Worksheets("Historical Data").Range("L12:L13").value = ""
    'Worksheets("Historical Data").Range("M12:M13").Value = "=TEXT(NOW(),""yyyymmdd hh:mm:ss"")"
    Worksheets("Historical Data").Range("M12:M13").value = "=IF(NOW()-TODAY()>TIME(16,1,0),CONCATENATE(TEXT(NOW(),""yyyymmdd""),"" 16:01:00""),CONCATENATE(TEXT(NOW()-1,""yyyymmdd""),"" 16:01:00""))"
    Worksheets("Historical Data").Range("N12:N13").value = "1 D"
    Worksheets("Historical Data").Range("O12:O13").value = "1 min"
    Worksheets("Historical Data").Range("P12:P13").value = "TRADES"
    Worksheets("Historical Data").Range("Q12:Q13").value = "1"
    Worksheets("Historical Data").Range("R12:R13").value = "1"
    Worksheets("Historical Data").Range("S12").value = "NUGT data"
    Worksheets("Historical Data").Range("S13").value = "DUST data"

    'requests data
    Worksheets("Historical Data").Activate
    Range("A12").Select
    'Call Sheets("Historical Data").RequestHistoricalData
    Application.Run "Sheet13.RequestHistoricalData_Click"
    Worksheets("Historical Data").Activate
    Range("A13").Select
    'Call Sheets("Historical Data").RequestHistoricalData
    Application.Run "Sheet13.RequestHistoricalData_Click"

'    Application.OnTime Now + TimeValue("00:01:30"), "ContinuePre"

'PROCESS MUST FINISH BEFORE CONTINUING!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Sub
Sub FillPreviousDay()

    'finds last row and applies it to the variable "LastDataRow"
    LastDataRow = Worksheets("NUGT data").Cells.Find(What:="*", After:=Range("e1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row

    'sets height of data extracted
    ExtractedDataHeight = LastDataRow - 2

    'gives the name of "RequestedDataRange" to the entire requested price table
    Set RequestedTimeRange = Worksheets("NUGT data").Range("B3:B" & LastDataRow)
    Set RequestedNUGTRange = Worksheets("NUGT data").Range("D3:F" & LastDataRow)
    Set RequestedDUSTRange = Worksheets("DUST data").Range("D3:F" & LastDataRow)

    'copies RequestedDataRange to Technical Analysis sheet
    Worksheets("Technical Analysis").Range("D80").Resize(ExtractedDataHeight, 1) = RequestedTimeRange.value
    Worksheets("Technical Analysis").Range("H80").Resize(ExtractedDataHeight, 3) = RequestedNUGTRange.value
    Worksheets("Technical Analysis").Range("X80").Resize(ExtractedDataHeight, 3) = RequestedDUSTRange.value

    'defines PreviousOpenMarketDate
    PreviousOpenMarketDate = Worksheets("Technical Analysis").Range("D80").value

    'finds last previous day row and applies it to the variable "LastPreviousDayAnalysisRow"
    LastPreviousDayAnalysisRow = Cells.Find(What:=Left(PreviousOpenMarketDate, 8), After:=Range("e1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row

    'colors range light grey
    Worksheets("Technical Analysis").Range("A80:AQ" & LastPreviousDayAnalysisRow).Font.Color = RGB(160, 160, 160)
    
    Worksheets("Technical Analysis").Range("A" & LastPreviousDayAnalysisRow & ":AQ" & LastPreviousDayAnalysisRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
End Sub
Sub MarketOpen()

    AbortTime = TimeValue("16:00:00")
    If Range("ProgramActive") = "Yes" And Range("Weekday") = "Yes" Then
        Application.OnTime TimeValue("09:30:59"), "RequestCurrentMinute"
        Do While True
            Application.OnTime Now + TimeValue("00:00:06"), "FillAnalysisAndOrder"
            If Now > AbortTime Then Exit Do
            Application.OnTime Now + TimeValue("00:00:54"), "RequestCurrentMinute"
        Loop
    End If

End Sub
Sub RequestCurrentMinute()

    'fills user name & monitor rows
    Worksheets("Historical Data").Range("D5").value = "my_username"
    Worksheets("Historical Data").Range("J4").value = 12
    Worksheets("Historical Data").Range("J5").value = 13

    'clears any data requests for nugt & dust, preps arguments for running data request
    Worksheets("Historical Data").Range("A12").value = "NUGT"
    Worksheets("Historical Data").Range("A13").value = "DUST"
    Worksheets("Historical Data").Range("B12:B13").value = "STK"
    Worksheets("Historical Data").Range("H12:H13").value = "SMART"
    Worksheets("Historical Data").Range("J12:J13").value = "USD"
    Worksheets("Historical Data").Range("L12:L13").value = ""
    Worksheets("Historical Data").Range("M12:M13").value = "=TEXT(NOW(),""yyyymmdd hh:mm:ss"")"
    Worksheets("Historical Data").Range("N12:N13").value = "60 S"
    Worksheets("Historical Data").Range("O12:O13").value = "1 min"
    Worksheets("Historical Data").Range("P12:P13").value = "TRADES"
    Worksheets("Historical Data").Range("Q12:Q13").value = "1"
    Worksheets("Historical Data").Range("R12:R13").value = "1"
    Worksheets("Historical Data").Range("S12").value = "NUGT data"
    Worksheets("Historical Data").Range("S13").value = "DUST data"

    'requests data
    Worksheets("Historical Data").Activate
    Range("A12").Select
    Call Sheets("Historical Data").RequestHistoricalData
    Worksheets("Historical Data").Activate
    Range("A13").Select
    Call Sheets("Historical Data").RequestHistoricalData

'REQUESTED TABLES MUST BE POPULATED BEFORE CONTINUING!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

End Sub
Sub FillAnalysisAndOrder()

    Call FillCurrentMinute
    Call CallBuySellOrders

End Sub
Sub FillCurrentMinute()

    'defines LastDataRow, LastAnalysisRow & NewLine
    LastDataRow = Worksheets("NUGT data").Cells.Find(What:="*", After:=Range("e1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row
    LastAnalysisRow = Worksheets("Technical Analysis").Cells.Find(What:="*", After:=Range("i1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row
    NewLine = LastAnalysisRow + 1

    'names extracted ranges
    Set CurrentMinuteTime = Worksheets("NUGT data").Range("B" & LastDataRow)
    Set RequestedNUGTRange = Worksheets("NUGT data").Range("D" & LastDataRow & ":" & "F" & LastDataRow)
    Set RequestedDUSTRange = Worksheets("DUST data").Range("D" & LastDataRow & ":" & "F" & LastDataRow)

    'copies RequestedDataRange to Technical Analysis sheet
    Worksheets("Technical Analysis").Range("D" & NewLine).value = CurrentMinuteTime.value
    Worksheets("Technical Analysis").Range("H" & NewLine).Resize(1, 3) = RequestedNUGTRange.value
    Worksheets("Technical Analysis").Range("X" & NewLine).Resize(1, 3) = RequestedDUSTRange.value

    'fills in formula
    'If Range("nugt_trade_active") = 0 And Range("InTradeWindow") = "Yes" Then
    'if range("nugt_trade_active")=0
    
End Sub
Sub CallBuySellOrders()

    'tests if there are any order signals
    If Worksheets("Technical Analysis").Range("AO" & LastAnalysisRow) = "buy" Then
        Call BuyNUGT
    
    ElseIf Worksheets("Technical Analysis").Range("AR" & LastAnalysisRow) = "buy" Then
        Call BuyDUST
    
    ElseIf Worksheets("Technical Analysis").Range("AO" & LastAnalysisRow) = "sell" Then
        Call SellNUGT
    
    ElseIf Worksheets("Technical Analysis").Range("AR" & LastAnalysisRow) = "sell" Then
        Call SellDUST
    
    End If

End Sub
Sub BuyNUGT()

    'sets all variables
    LastAnalysisRow = Worksheets("Technical Analysis").Cells.Find(What:="*", After:=Range("i1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row
    MostRecentNUGTPrice = Worksheets("Technical Analysis").Range("J" & LastAnalysisRow)
    TradeBetSize = Worksheets("Parameters").Range("J2")

    'fills user name
    Worksheets("Basic Orders").Range("D5").value = "my_username"

    'fills order parameters
    Worksheets("Basic Orders").Range("A12").value = "NUGT"
    Worksheets("Basic Orders").Range("B12").value = "STK"
    Worksheets("Basic Orders").Range("H12").value = "SMART"
    Worksheets("Basic Orders").Range("J12").value = "USD"
    Worksheets("Basic Orders").Range("M12").value = "BUY"
    'Worksheets("Basic Orders").Range("N12").value = WorksheetFunction.RoundDown(500 / MostRecentNUGTPrice, 0) '-for testing order sub
    'Worksheets("Basic Orders").Range("N13").value = WorksheetFunction.RoundDown(500 / MostRecentNUGTPrice, 0) '-for testing order sub
    Worksheets("Basic Orders").Range("N12").value = WorksheetFunction.RoundDown(TradeBetSize / MostRecentNUGTPrice, 0)
    Worksheets("Basic Orders").Range("N13").value = WorksheetFunction.RoundDown(TradeBetSize / MostRecentNUGTPrice, 0)
    Worksheets("Basic Orders").Range("O12").value = "LMT"
    'Worksheets("Basic Orders").Range("P12").value = 10 * 1.01 '(for testing order sub)
    Worksheets("Basic Orders").Range("P12").value = MostRecentNUGTPrice * 1.01

    'places order
    Worksheets("Basic Orders").Activate
    Range("A12").Select
    Call Sheets("Basic Orders").placeOrder

End Sub
Sub BuyDUST()

    'sets all variables
    LastAnalysisRow = Worksheets("Technical Analysis").Cells.Find(What:="*", After:=Range("i1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row
    MostRecentDUSTPrice = Worksheets("Technical Analysis").Range("Z" & LastAnalysisRow)
    TradeBetSize = Worksheets("Parameters").Range("J2")

    'fills user name
    Worksheets("Basic Orders").Range("D5").value = "my_username"
    
    'fills order parameters
    Worksheets("Basic Orders").Range("A14").value = "DUST"
    Worksheets("Basic Orders").Range("B14").value = "STK"
    Worksheets("Basic Orders").Range("H14").value = "SMART"
    Worksheets("Basic Orders").Range("J14").value = "USD"
    Worksheets("Basic Orders").Range("M14").value = "BUY"
    Worksheets("Basic Orders").Range("N14").value = WorksheetFunction.RoundDown(TradeBetSize / MostRecentDUSTPrice, 0)
    Worksheets("Basic Orders").Range("N15").value = WorksheetFunction.RoundDown(TradeBetSize / MostRecentDUSTPrice, 0)
    Worksheets("Basic Orders").Range("O14").value = "LMT"
    Worksheets("Basic Orders").Range("P14").value = MostRecentDUSTPrice * 1.01
    
    'places order
    Worksheets("Basic Orders").Activate
    Range("A14").Select
    Call Sheets("Basic Orders").placeOrder

End Sub
Sub SellNUGT()

    'sets all variables
    LastAnalysisRow = Worksheets("Technical Analysis").Cells.Find(What:="*", After:=Range("i1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row
    MostRecentNUGTPrice = Worksheets("Technical Analysis").Range("J" & LastAnalysisRow)
    TradeBetSize = Worksheets("Parameters").Range("J2")

    'fills user name
    Worksheets("Basic Orders").Range("D5").value = "my_username"
    
    'fills order parameters
    Worksheets("Basic Orders").Range("A13").value = "NUGT"
    Worksheets("Basic Orders").Range("B13").value = "STK"
    Worksheets("Basic Orders").Range("H13").value = "SMART"
    Worksheets("Basic Orders").Range("J13").value = "USD"
    Worksheets("Basic Orders").Range("M13").value = "SELL"
    Worksheets("Basic Orders").Range("O13").value = "LMT"
    Worksheets("Basic Orders").Range("P13").value = MostRecentNUGTPrice * 0.99
    
    'places order
    Worksheets("Basic Orders").Activate
    Range("A13").Select
    Call Sheets("Basic Orders").placeOrder

End Sub
Sub SellDUST()

    'sets all variables
    LastAnalysisRow = Worksheets("Technical Analysis").Cells.Find(What:="*", After:=Range("i1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row
    MostRecentDUSTPrice = Worksheets("Technical Analysis").Range("Z" & LastAnalysisRow)
    TradeBetSize = Worksheets("Parameters").Range("J2")

    'fills user name
    Worksheets("Basic Orders").Range("D5").value = "my_username"
    
    'fills order parameters
    Worksheets("Basic Orders").Range("A15").value = "DUST"
    Worksheets("Basic Orders").Range("B15").value = "STK"
    Worksheets("Basic Orders").Range("H15").value = "SMART"
    Worksheets("Basic Orders").Range("J15").value = "USD"
    Worksheets("Basic Orders").Range("M15").value = "SELL"
    Worksheets("Basic Orders").Range("O15").value = "LMT"
    Worksheets("Basic Orders").Range("P15").value = MostRecentDUSTPrice * 0.99

    'places order
    Worksheets("Basic Orders").Activate
    Range("A15").Select
    Call Sheets("Basic Orders").placeOrder

End Sub
Sub MidDayUpdateAnalysis()

    Call ClearOutCurrentDayData
    Call MidDayRequest
    Call DefineLastDataRow
    Call UpdateCurrentData
    'then immediately jump into debug mode of MarketOpen and press play on next minute

End Sub
Sub ClearOutCurrentDayDataAndRunCurrentDayRequest()

    'defines PreviousOpenMarketDate
    PreviousOpenMarketDate = Worksheets("Technical Analysis").Range("D80").value

    'finds last previous day row and applies it to the variable "LastPreviousDayAnalysisRow"
    LastPreviousDayAnalysisRow = Cells.Find(What:=Left(PreviousOpenMarketDate, 8), After:=Range("e1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row
    TopCurrentDayAnalysisRow = LastPreviousDayAnalysisRow + 1

    'clears current day data
    Worksheets("Technical Analysis").Range("D" & TopCurrentDayAnalysisRow & ":D1000").Clear
    Worksheets("Technical Analysis").Range("H" & TopCurrentDayAnalysisRow & ":J1000").Clear
    Worksheets("Technical Analysis").Range("X" & TopCurrentDayAnalysisRow & ":Z1000").Clear
    
    'fills user name
    Worksheets("Historical Data").Range("D5").value = "my_username"

    'clears any data requests for nugt & dust, preps arguments for running data request
    Worksheets("Historical Data").Range("A12").value = "NUGT"
    Worksheets("Historical Data").Range("A13").value = "DUST"
    Worksheets("Historical Data").Range("B12:B13").value = "STK"
    Worksheets("Historical Data").Range("H12:H13").value = "SMART"
    Worksheets("Historical Data").Range("J12:J13").value = "USD"
    Worksheets("Historical Data").Range("L12:L13").value = ""
    Worksheets("Historical Data").Range("M12:M13").value = "=TEXT(NOW(),""yyyymmdd hh:mm:ss"")"
    Worksheets("Historical Data").Range("N12:N13").value = "1 D"
    Worksheets("Historical Data").Range("O12:O13").value = "5"
    Worksheets("Historical Data").Range("P12:P13").value = "TRADES"
    Worksheets("Historical Data").Range("Q12:Q13").value = "1"
    Worksheets("Historical Data").Range("R12:R13").value = "1"
    Worksheets("Historical Data").Range("S12").value = "NUGT data"
    Worksheets("Historical Data").Range("S13").value = "DUST data"

    'requests data
    Worksheets("Historical Data").Activate
    Range("A12").Select
    Call Sheets("Historical Data").RequestHistoricalData
    Worksheets("Historical Data").Activate
    Range("A13").Select
    Call Sheets("Historical Data").RequestHistoricalData

'    Application.OnTime Now + TimeValue("00:01:30"), "ContinuePre"

'PROCESS MUST FINISH BEFORE CONTINUING!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

End Sub
Sub UpdateCurrentData() 'evaluate this one step by step

    'finds last row and applies it to the variable "LastDataRow"
    LastDataRow = Worksheets("NUGT data").Cells.Find(What:="*", After:=Range("e1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row

    'sets height of data extracted
    ExtractedDataHeight = LastDataRow - 2

    'gives the name of "RequestedDataRange" to the entire requested price table
    Set RequestedTimeRange = Worksheets("NUGT data").Range("B3:B" & LastDataRow)
    Set RequestedNUGTRange = Worksheets("NUGT data").Range("D3:F" & LastDataRow)
    Set RequestedDUSTRange = Worksheets("DUST data").Range("D3:F" & LastDataRow)

    'copies RequestedDataRange to Technical Analysis sheet
    Worksheets("Technical Analysis").Range("D80").Resize(ExtractedDataHeight, 1) = RequestedTimeRange.value.Font.Color = RGB(160, 160, 160)
    Worksheets("Technical Analysis").Range("H80").Resize(ExtractedDataHeight, 3) = RequestedNUGTRange.value
    Worksheets("Technical Analysis").Range("X80").Resize(ExtractedDataHeight, 3) = RequestedDUSTRange.value

    'defines PreviousOpenMarketDate
    PreviousOpenMarketDate = Worksheets("Technical Analysis").Range("D80").value

    'finds last previous day row and applies it to the variable "LastPreviousDayAnalysisRow"
    LastPreviousDayAnalysisRow = Cells.Find(What:=Left(PreviousOpenMarketDate, 8), After:=Range("e1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).row

    'colors range light grey
    Worksheets("Technical Analysis").Range("A80:AQ" & LastPreviousDayAnalysisRow).Font.Color = RGB(160, 160, 160)
    
    Worksheets("Technical Analysis").Range("A" & LastPreviousDayAnalysisRow & ":AQ" & LastPreviousDayAnalysisRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .colorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    'gives the name of "RequestedDataRange" to the entire requested price table
    Set RequestedTimeRange = Worksheets("NUGT data").Range("B3:B" & LastDataRow)
    Set RequestedNUGTRange = Worksheets("NUGT data").Range("D3:F" & LastDataRow)
    Set RequestedDUSTRange = Worksheets("DUST data").Range("D3:F" & LastDataRow)
    
    'copies RequestedDataRange to Technical Analysis sheet
    Worksheets("Technical Analysis").Activate
    Range("D" & (LastPreviousDayAnalysisRow + 1)).Resize(LastDataRow - 2, 1) = RequestedTimeRange.value
    Range("H" & (LastPreviousDayAnalysisRow + 1)).Resize(LastDataRow - 2, 3) = RequestedNUGTRange.value
    Range("X" & (LastPreviousDayAnalysisRow + 1)).Resize(LastDataRow - 2, 3) = RequestedDUSTRange.value

End Sub
Sub CopyOverParameterPage()

    Set OldWorkbook = Workbooks("old DDE API for parameters & analysis pages.xlsm")
    Set NewWorkbook = Workbooks("My_TWS_ActiveX.xlsm")
    'NewWorkbook.Activate
    'Sheets.Add(Before:=Sheets("Tickers")).name = "Parameters"
    OldWorkbook.Activate
    LastDataRow = Worksheets("Parameters").Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).row
    LastDataColumn = Worksheets("Parameters").Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).column
    
    Set ParametersRange = Sheets("Parameters").Range(Cells(1, 1), Cells(LastDataRow, LastDataColumn))
    ParametersRange.Copy
    NewWorkbook.Sheets("General").Range("A10").PasteSpecial Paste:=xlPasteFormulas
    NewWorkbook.Sheets("General").Range("A10").PasteSpecial Paste:=xlPasteFormats

End Sub
Sub CopyOverTechnicalAnalysisPage()

    Set OldWorkbook = Workbooks("old DDE API for parameters & analysis pages.xlsm")
    Set NewWorkbook = Workbooks("My_TWS_ActiveX.xlsm")
    NewWorkbook.Activate
    Sheets.Add(After:=Sheets("Historical Data")).name = "Technical Analysis"
    OldWorkbook.Activate
    LastDataRow = Worksheets("Technical Analysis").Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).row
    LastDataColumn = Worksheets("Technical Analysis").Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, _
    LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).column
    
    Set AnalysisRange = Sheets("Technical Analysis").Range(Cells(1, 1), Cells(LastDataRow, LastDataColumn))
    AnalysisRange.Copy
    NewWorkbook.Activate
    Sheets("Technical Analysis").Range("A1").PasteSpecial Paste:=xlPasteFormulas
    Sheets("Technical Analysis").Range("A1").PasteSpecial Paste:=xlPasteFormats

    rows("1:76").Hidden = True
    columns("K:U").Hidden = True
    columns("AA:AK").Hidden = True

    rows(80).Select
    ActiveWindow.FreezePanes = True
    
    Range("A80:AQ469").Font.Color = RGB(160, 160, 160)

End Sub
Sub CopyHistoricalStockTickers()

    Set OldWorkbook = Workbooks("My_TWS_DDE20170206-1.xlsm")
    Set NewWorkbook = Workbooks("My_TWS_DDE.xlsm")
    
    NewWorkbook.Sheets("Historical Data").Range("D5").value = "my_username"
    NewWorkbook.Sheets("Historical Data").Range("J5").value = 13
    
    OldWorkbook.Sheets("Historical Data").Range("A12:V206").Copy
    
    NewWorkbook.Sheets("Historical Data").Range("A12").PasteSpecial Paste:=xlPasteValues
    NewWorkbook.Sheets("Historical Data").Range("A12").PasteSpecial Paste:=xlPasteFormats

End Sub
Sub CopyBasicOrderStockTickers()

    Set OldWorkbook = Workbooks("My_TWS_DDE20170206-1.xlsm")
    Set NewWorkbook = Workbooks("My_TWS_DDE.xlsm")
    
    NewWorkbook.Sheets("Basic Orders").Range("D5").value = "my_username"
    NewWorkbook.Sheets("Market Scanner").Range("G5").value = 13
    
    OldWorkbook.Sheets("Basic Orders").Range("A12:BZ95").Copy
    
    NewWorkbook.Sheets("Basic Orders").Range("A12").PasteSpecial Paste:=xlPasteValues
    NewWorkbook.Sheets("Basic Orders").Range("A12").PasteSpecial Paste:=xlPasteFormats

End Sub
Sub Copy_All_Defined_Names()
    
    Set OldWorkbook = Workbooks("My_TWS_DDE20170206-1.xlsm")
    Set NewWorkbook = Workbooks("My_TWS_DDE.xlsm")
    
    'Loops through all of the defined names in the active workbook.
    For Each x In OldWorkbook.Names
        'for each x, this adds defined names from active workbook to target workbook ("Book2.xlsm")
        '"x.value" refers to the cell references the defined name points to
        'Workbooks("Book2.xls").Names.Add name:=x.name, refersTo:=x.value
        NewWorkbook.Names.Add name:=x.name, refersTo:=x.value
   Next x
End Sub
Sub CopyButtons()

    Set OldWorkbook = Workbooks("My_TWS_DDE20170206-1.xlsm")
    Set NewWorkbook = Workbooks("My_TWS_DDE.xlsm")

    Application.ScreenUpdating = False
    Dim x As OLEObject, y As OLEObject
'    Set x = OldWorkbook.Sheets("Parameters").OLEObjects("HasCustomName")   "HasCustomName" needs to be changed to something else
    Set y = x.Duplicate
    Dim xName As String
    xName = x.name
    y.Cut
    With NewWorkbook.Sheets("Parameters")
        .Paste
        .OLEObjects(.OLEObjects.Count).name = xName
        .Activate
    End With
    Application.ScreenUpdating = True
End Sub
