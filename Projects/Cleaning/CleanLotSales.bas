Attribute VB_Name = "NewlandTableProcessing"
Dim TxtBoxSectionResult As String, r As Integer, LastColumn As Long, LastRow As Long, LotIDColumn As Long, LotIDRange As Range, _
BlockColumn As Long, SectionColumn As Long, SectionRange As Range, BlockRange As Range, ValueCheck As Range, _
LotNumColumn As Long, ConstructingBuilderColumn As Long, SettlementDate As Long
Sub PrepareLotTable()

Call CleanFieldNames
Call OrganizeLotIDkeyComponents
Call CreateLotID
Call AddBuilderLabel
Call AddYearSettled
Call AddLotClosed
Call AddBuilderClosed
Call SortFreezeBold
Call SaveSheetAs
'Call CloseSheet

End Sub
Sub CombineTwoColumns()

Dim ask As Range, target As Range, LotCount As Long, Column1 As String, Column2 As String, Character As String, Column1Index As Long, Column2Index As Long, Column1Range As Range, Column2Range As Range

    LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row

    'combs the lot numbers to ensure they meet the 4-digit standard
    Call FourDigitNums
        
    NoEntry = "You didn't enter a letter, exiting"
    NoEntryTitle = "Error"

    Column1 = Application.InputBox("Please indicate first column letter(s) for merge", "Select Column 1")
    If IsNumeric(Column1) Then
        MsgBox NoEntry, vbExclamation, NoEntryTitle
        Exit Sub
    End If
    Column1Index = Range(Column1 & 1).Column
    Set Column1Range = Range(Cells(1, Column1Index), Cells(LastRow, Column1Index))

    Column2 = Application.InputBox("Now indicate second column letter(s) for merge", "Select Column 2")
    If IsNumeric(Column2) Then
        MsgBox NoEntry, vbExclamation, NoEntryTitle
        Exit Sub
    End If
    Column2Index = Range(Column2 & 1).Column
    Set Column2Range = Range(Cells(1, Column2Index), Cells(LastRow, Column2Index))
    
    Character = Application.InputBox("Indicate dividing character, or hit 'enter' to leave blank (ex: -,|)", "Select Dividing Character (if any)")
    
    Column2Range.Offset(0, 1).Insert (xlShiftToRight)
    For r = 1 To LastRow
        Cells(r, Column2Index + 1) = Cells(r, Column1Index) & Character & Cells(r, Column2Index)
    Next r
    Column2Range.Offset(0, 1).Columns.AutoFit

End Sub
Sub CleanFieldNames()

    'defines LastColumn
    LastColumn = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column

    'cleans headers (field names) of trailing & leading spaces
    For c = 1 To LastColumn
        Cells(1, c).Formula = Trim(Cells(1, c))
    Next c

End Sub
Sub CalcLotCount()

    LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
    LotCount = LastRow - 1
    MsgBox "There are " & LotCount & " lots total in this table."

End Sub
Sub OrganizeLotIDkeyComponents()

'   defines LastRow & LastColumn
    LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
    LastColumn = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
        
'   checks for "Lot Identifier" field
    Set ValueCheck = Cells.Find(What:="Lot Identifier", After:=Range("Z1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

'   if "Lot Identifier" exists, defines "Lot Identifier" index and renames to "Lot ID"
    If Not ValueCheck Is Nothing Then

        LotIDIndex = Cells.Find(What:="Lot Identifier", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
            SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
        Cells(1, LotIDColumn).Value = "LOT_ID"
        
'       replaces trailing spaces from Lot ID
        For r = 2 To LastRow
            Cells(r, LotIDIndex).Formula = Trim(Cells(r, LotIDIndex))
        Next r

'   else, if "Lot Identifier" doesn't exist, checks for "Section" field
    Else
        Set ValueCheck = Cells.Find(What:="Section", After:=Range("Z1"), LookAt:=xlPart, LookIn:=xlFormulas, _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
        
'       if "Section" exists, moves it all the way to the left
        If Not ValueCheck Is Nothing Then

            SectionIndex = Cells.Find(What:="Section", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
            Columns(SectionIndex).Select
            Selection.Cut
            Columns(1).Select
            Selection.Insert Shift:=xlToRight

'           if "Block" field present
            Set ValueCheck = Cells.Find(What:="Block", After:=Range("Z1"), LookAt:=xlPart, LookIn:=xlFormulas, _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
            If Not ValueCheck Is Nothing Then
            
'               defines BlockColumn
                BlockIndex = Cells.Find(What:="Block", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
        
'               defines the Block records as BlockRange (which is the range data type)
                Set BlockRange = Range(Cells(2, BlockIndex), Cells(LastRow, BlockIndex))
            
'               capitalizes any text in BlockRange
                BlockRange.Value = BlockRange.Parent.Evaluate("INDEX(UPPER(" & BlockRange.Address & "),)")

'               if "Block" isn't in 2nd column, moves "Block" to 2nd column, after "Section"
                If BlockColumn <> 2 Then
                    Columns(BlockColumn).Select
                    Selection.Cut
                    Columns(2).Select
                    Selection.Insert Shift:=xlToRight
                End If
                
'               defines column # of "Lot Number"
                LotNumIndex = Cells.Find(What:="Lot Number", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
                
'               if "Lot Number" isn't in 3rd column, moves it to 3rd column, after "Block"
                If LotNumIndex <> 3 Then
                    Columns(LotNumIndex).Select
                    Selection.Cut
                    Columns(3).Select
                    Selection.Insert Shift:=xlToRight
                End If
            End If

'       ---IF NO SECTION IS PRESENT IN TABLE---
        Else
'           if block is present...
            Set ValueCheck = Cells.Find(What:="Block", After:=Range("Z1"), LookAt:=xlPart, LookIn:=xlFormulas, _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
            If Not ValueCheck Is Nothing Then

'               defines BlockColumn
                BlockColumn = Cells.Find(What:="Block", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
        
'               defines the Block records as BlockRange (which is the range data type)
                Set BlockRange = Range(Cells(2, BlockColumn), Cells(LastRow, BlockColumn))
            
'               capitalizes any text in BlockRange
                BlockRange.Value = BlockRange.Parent.Evaluate("INDEX(UPPER(" & BlockRange.Address & "),)")
    
'               if "Block" isn't in 1st column, moves "Block" to 1st column
                If BlockColumn <> 1 Then
                    Columns(BlockColumn).Select
                    Selection.Cut
                    Columns(1).Select
                    Selection.Insert Shift:=xlToRight
                End If

'               defines column # of "Lot Number"
                LotNumColumn = Cells.Find(What:="Lot Number", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
                
'               if "Lot Number" isn't in 2nd column, moves it to 2nd column, after "Block"
                If LotNumIndex <> 2 Then
                    Columns(BlockIndex).Select
                    Selection.Cut
                    Columns(2).Select
                    Selection.Insert Shift:=xlToRight
                End If
            End If
        End If

'       combs the lot numbers to ensure they meet the 4-digit standard
        Call FourDigitNums
        
    End If
End Sub
Sub FourDigitNums()
'   combs the lot numbers to ensure they meet the 4-digit standard

'   defines column # of "Lot Number"
    LotNumIndex = Cells.Find(What:="Lot Number", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column

'   defines LastRow
    LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row

'   loops through the lot numbers, and if they don't have preceeding 0s, adds them
    r = 2
    For r = 2 To LastRow
        LotNum = Cells(r, LotNumIndex)
        If InStr(LotNum, ".") > 0 And InStr(LotNum, ".") < 5 Then
            Cells(r, LotNumIndex).Formula = "'" & Application.Rept("0", 5 - InStr(LotNum, ".")) & LotNum
        ElseIf Len(LotNum) < 4 Then
            Cells(r, LotNumIndex).Formula = "'" & Application.Rept("0", 4 - Len(LotNum)) & LotNum
        Else
            Cells(r, LotNumIndex).Formula = "'" & LotNum
        End If
    Next r

End Sub
Sub CreateLotID()

'   Sub is only completed if there is no "LOT_ID" field
    Set ValueCheck = Cells.Find(What:="LOT_ID", After:=Range("Z1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If ValueCheck Is Nothing Then

    '   defines LastRow & LastColumn
        LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
        LastColumn = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
    
    '   defines column # of "Lot Number"
        LotNumColumn = Cells.Find(What:="Lot Number", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
    
    '   ---POPULATES LOT_ID VALUES---
    '   inserts new field after "Lot Number", gives name "LOT_ID", sets record number type to 'general'
        Columns(LotNumColumn + 1).Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Cells(1, LotNumColumn + 1).Value = "LOT_ID"
        Cells(2, LotNumColumn + 1).NumberFormat = "General"
        
    '   defines column # of "LOT_ID"
        LotIDColumn = Cells.Find(What:="LOT_ID", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
    
    '   defines the Lot ID's entire range as LotIDRange
        Set LotIDRange = Range(Cells(2, LotIDColumn), Cells(LastRow, LotIDColumn))
    
    '   creates composite LOT_ID value from lot#, blocks & sections (if they're present)
        If LotIDColumn = 2 Then
            LotIDRange.FormulaR1C1 = "=RC[-1]"
        ElseIf LotIDColumn = 3 Then
            LotIDRange.FormulaR1C1 = "=IF(RC[-2]<>"""",CONCATENATE(RC[-2],""|"",RC[-1]),RC[-1])"
        ElseIf LotIDColumn = 4 Then
    '       popup text box <Use Section in Unique Identifier? (Y/N)>
            TxtBoxSectionResult = InputBox("Use Section in Unique Identifier? (Y/N)", "Decision Must Be Made")
            If TxtBoxSectionResult <> "n" Then
                LotIDRange.FormulaR1C1 = "=IF(RC[-2]<>"""",CONCATENATE(RC[-3],""|"",RC[-2],""|"",RC[-1]),CONCATENATE(RC[-3],""|"",RC[-1]))"
            Else
                LotIDRange.FormulaR1C1 = "=IF(RC[-2]<>"""",CONCATENATE(RC[-2],""|"",RC[-1]),RC[-1])"
            End If
        End If
    
    '   auto sizes column width, copy & paste values (overwrites concatenate formulas with just the result value)
        Cells(2, LotIDColumn).AutoFill Destination:=Range(Cells(2, LotIDColumn), Cells(LastRow, LotIDColumn))
        Columns(LotIDColumn).EntireColumn.AutoFit
        Columns(LotIDColumn).Copy
        Columns(LotIDColumn).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If

End Sub
Sub AddBuilderLabel()

    'defines LastRow & LastColumn
    LastColumn = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
    LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row

    'find if "Constructing Builder" is in the top row
    Set rng = Cells.Find(What:="Constructing Builder", After:=Range("Z1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

    If Not rng Is Nothing Then 'if "Constructing Builder" was used:
        ConstructingBuilderColumn = Cells.Find(What:="Constructing Builder", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
            SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
        Columns(ConstructingBuilderColumn).Cut
        Columns(LastColumn + 1).Select
        Selection.Insert Shift:=xlToRight
        'finds & moves "Constructing Builder" to end
    End If

    'find if "Construction Name" is in the top row
    Set rng = Cells.Find(What:="Construction Name", After:=Range("Z1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

    'if "Construction Name" was used, then
    If Not rng Is Nothing Then 'when rng <> nothing means found something'

        'finds & moves "Construction Name" to end
        ConstructingBuilderColumn = Cells.Find(What:="Construction Name", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
            SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
        Columns(ConstructingBuilderColumn).Cut
        Columns(LastColumn + 1).Select
        Selection.Insert Shift:=xlToRight

    End If

    'add BLDR_LABEL from conversion table
    Cells(1, LastColumn + 1).Value = "BLDR_LABEL"
    Cells(2, LastColumn + 1).Select
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]<>"""",VLOOKUP(TRIM(RC[-1]),'B:\NEWLAND\Conversion_Table-Builder_Name.xlsx'!Builder_Label_Table,2,FALSE),"""")"
    Selection.AutoFill Destination:=Range(Cells(2, LastColumn + 1), Cells(LastRow, LastColumn + 1))
    Columns(LastColumn + 1).EntireColumn.AutoFit
    Columns(LastColumn + 1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, _
        Transpose:=False
    
End Sub
Sub AddYearSettled()

    'defines LastColumn & LastRow
    LastColumn = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
    LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row

    'finds & moves "Lot Settlement Date" to end
    SettlementDate = Cells.Find(What:="Lot Settlement Date", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False).Column
    Columns(SettlementDate).Select
    Selection.Cut
    Columns(LastColumn + 1).Select
    Selection.Insert Shift:=xlToRight
    
    'creates "YR_SETTLED"
    Cells(1, LastColumn + 1).Value = "YR_SETTLED"
    Cells(2, LastColumn + 1).Select
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",TEXT(RC[-1],""yyyy""),"""")"
    Selection.AutoFill Destination:=Range(Cells(2, LastColumn + 1), Cells(LastRow, LastColumn + 1))
    Columns(LastColumn + 1).EntireColumn.AutoFit
    Columns(LastColumn + 1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
End Sub
Sub AddLotClosed()

    're-defines LastColumn & LastRow
    LastColumn = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
    LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row

    'creates "LOT_CLOSED"
    Cells(1, LastColumn + 1).Value = "LOT_CLOSED"
    Cells(2, LastColumn + 1).Select
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<>"""",""Closed"",""Inventory"")"
    Selection.AutoFill Destination:=Range(Cells(2, LastColumn + 1), Cells(LastRow, LastColumn + 1))
    Columns(LastColumn + 1).EntireColumn.AutoFit
    Columns(LastColumn + 1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Sub
Sub AddBuilderClosed()

    're-defines LastColumn & LastRow
    LastColumn = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
    LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
    
    'creates "BLDR_CLOSED"
    Cells(1, LastColumn + 1).Value = "BLDR_CLOSED"
    Cells(2, LastColumn + 1).Select
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-4]<>"""",CONCATENATE(RC[-4],"" - "",RC[-1]),IF(RC[-3]<>"""",""? - Closed"",""Newland Inventory""))"
    Selection.AutoFill Destination:=Range(Cells(2, LastColumn + 1), Cells(LastRow, LastColumn + 1))
    Columns(LastColumn + 1).EntireColumn.AutoFit
    Columns(LastColumn + 1).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Sub
Sub SortFreezeBold()

    're-defines LastColumn & LastRow
    LastColumn = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
    LastRow = Cells.Find(What:="*", After:=Range("a1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
    LotID_Index = Cells.Find(What:="LOT_ID", After:=Range("Z1"), LookAt:=xlPart, LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Column

    'sorts on LotID
    Range(Cells(1, 1), Cells(LastRow, LastColumn)).Sort key1:=Range(Cells(2, LotID_Index), Cells(2, LotID_Index)), order1:=xlAscending, Header:=xlYes

    'freezes top row
    Rows(2).Select
    ActiveWindow.FreezePanes = True
    
    'makes headers bold
    Range(Cells(1, 1), Cells(1, LastColumn)).Font.Bold = True

End Sub
Sub SaveSheetAs()
    
    'defines path & base file name
    Path = ActiveWorkbook.Path
    BaseName = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5)
    
    'saves to current directory with 'processed' appended to the end of file name
    ActiveWorkbook.SaveAs (Path & "\" & BaseName & "processed.xlsx")

End Sub
Sub CloseSheet()

    'Closes the workbook & saves any changes
    ActiveWorkbook.Close True

End Sub
