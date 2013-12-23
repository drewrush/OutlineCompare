Attribute VB_Name = "Module1"
Public Sub ImportCSVs()

    Call TextFileLoop("_V11")
    Call TextFileLoop("_V9")
    Call Cleanup
    Call WorksheetLoop
    Srt = SortWorksheetsByName(0, 0)

    
End Sub

Function GetFolder(strPath As String) As String
Dim fldr As FileDialog
Dim sItem As String
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
    .Title = "Select the folder that contains your Outline Extracts."
    .AllowMultiSelect = False
    .InitialFileName = strPath
    If .Show <> -1 Then GoTo NextCode
    sItem = .SelectedItems(1)
End With
NextCode:
GetFolder = sItem
Set fldr = Nothing
End Function

Sub TextFileLoop(version As String)
    Dim StrFile As String
    Dim ConnStr As String
    StrFldr = GetFolder("C:\")
    StrFile = Dir(StrFldr & "\*csv")
    
    
    Do While Len(StrFile) > 0
    FldrFile = StrFldr & "\" & StrFile
    
    ConnStr = "TEXT;" & FldrFile
        
    ActiveWorkbook.Worksheets.Add
    With ActiveSheet.QueryTables.Add(Connection:= _
        ConnStr _
        , Destination:=Range("$A$1"))
        .Name = StrFile
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = "?"
        .TextFileColumnDataTypes = Array(xlTextFormat, xlTextFormat, xlTextFormat, xlTextFormat, xlTextFormat)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

        dimdiff = Left(Replace([B2], " ", ""), 27)
        shtnm = ScrubInvalidChars(dimdiff)
        ActiveSheet.Name = shtnm & version
        Call Formulas
        StrFile = Dir
    Loop
End Sub


Sub Formulas()
    Dim lastRow As Long
    Dim lastCol As Long
    With ActiveSheet
        
        lastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        .Cells(1, lastCol + 1).FormulaR1C1 = "Concatenated Value"
        .Cells(1, lastCol + 2).FormulaR1C1 = "Match?"
        .Cells(1, lastCol + 3).FormulaR1C1 = "Exact?"
        Range(Cells(1, lastCol + 1), Cells(lastRow, lastCol + 1)).NumberFormat = "@"
    
        For I = 2 To lastRow
            For J = 1 To lastCol
                .Cells(I, lastCol + 1) = .Cells(I, lastCol + 1).Text & .Cells(I, J).Text
            Next J
        Next I
    
        repname = Replace(Right(ActiveSheet.Name, 3), "_", "") & Replace([B2], " ", "")
        rname = ScrubInvalidChars(repname)
        Range(Cells(2, lastCol + 1), Cells(lastRow, lastCol + 1)).Name = rname
    
    End With
End Sub

Sub WorksheetLoop()

    Dim WS_Count As Integer
    Dim I As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count - 1

    ' Begin the loop.
    For I = 1 To WS_Count
    ActiveWorkbook.Sheets(I).Activate

    With ActiveSheet
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        lastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
    
            frmAddress = .Cells(2, lastCol - 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            
            If Replace(Right(ActiveSheet.Name, 3), "_", "") = "V9" Then
                namepref = "V11"
                dname = Replace(ActiveSheet.Name, "_V9", "")
                oShn = Replace(ActiveSheet.Name, "_V9", "_V11")
            Else
                namepref = "V9"
                dname = Replace(ActiveSheet.Name, "_V11", "")
                oShn = Replace(ActiveSheet.Name, "_V11", "_V9")
            End If
                
            rname = namepref & dname
            
            form1name = "=IF(ISNA(MATCH(" & frmAddress & "," & rname & ",0)),FALSE,TRUE)"
            
            form2name = "=EXACT(" & frmAddress & "," & oShn & "!" & frmAddress & ")"
                    
            .Cells(2, lastCol - 1).Formula = form1name
            .Cells(2, lastCol - 1).AutoFill Destination:=Range(Cells(2, lastCol - 1), Cells(lastRow, lastCol - 1))
            .Cells(2, lastCol).Formula = form2name
            .Cells(2, lastCol).AutoFill Destination:=Range(Cells(2, lastCol), Cells(lastRow, lastCol))
            
            Range(Cells(1, 1), Cells(1, lastCol)).AutoFilter Field:=lastCol - 1, Criteria1:="FALSE"
    
    End With
    Next I

End Sub

Public Function SortWorksheetsByName(ByVal FirstToSort As Long, _
                            ByVal LastToSort As Long, _
                            Optional ByVal SortDescending As Boolean = False, _
                            Optional ByVal Numeric As Boolean = False) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SortWorksheetsByName
' This sorts the worskheets from FirstToSort to LastToSort by name
' in either ascending (default) or descending order. If successful,
' ErrorText is vbNullString and the function returns True. If
' unsuccessful, ErrorText gets the reason why the function failed
' and the function returns False. If you include the Numeric
' parameter and it is True, (1) all sheet names to be sorted
' must be numeric, and (2) the sort compares names as numbers, not
' text.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim M As Long
Dim N As Long
Dim WB As Workbook
Dim B As Boolean

Set WB = Worksheets.Parent
ErrorText = vbNullString

If WB.ProtectStructure = True Then
    ErrorText = "Workbook is protected."
    SortWorksheetsByName = False
End If
    
'''''''''''''''''''''''''''''''''''''''''''''''
' If First and Last are both 0, sort all sheets.
''''''''''''''''''''''''''''''''''''''''''''''
If (FirstToSort = 0) And (LastToSort = 0) Then
    FirstToSort = 1
    LastToSort = WB.Worksheets.Count
Else
    '''''''''''''''''''''''''''''''''''''''
    ' More than one sheet selected. We
    ' can sort only if the selected
    ' sheet are adjacent.
    '''''''''''''''''''''''''''''''''''''''

End If

If Numeric = True Then
    For N = FirstToSort To LastToSort
        If IsNumeric(WB.Worksheets(N).Name) = False Then
            ' can't sort non-numeric names
            ErrorText = "Not all sheets to sort have numeric names."
            SortWorksheetsByName = False
            Exit Function
        End If
    Next N
End If

'''''''''''''''''''''''''''''''''''''''''''''
' Do the sort, essentially a Bubble Sort.
'''''''''''''''''''''''''''''''''''''''''''''
For M = FirstToSort To LastToSort
    For N = M To LastToSort
        If SortDescending = True Then
            If Numeric = False Then
                If StrComp(WB.Worksheets(N).Name, WB.Worksheets(M).Name, vbTextCompare) > 0 Then
                    WB.Worksheets(N).Move before:=WB.Worksheets(M)
                End If
            Else
                If CLng(WB.Worksheets(N).Name) > CLng(WB.Worksheets(M).Name) Then
                    WB.Worksheets(N).Move before:=WB.Worksheets(M)
                End If
            End If
        Else
            If Numeric = False Then
                If StrComp(WB.Worksheets(N).Name, WB.Worksheets(M).Name, vbTextCompare) < 0 Then
                    WB.Worksheets(N).Move before:=WB.Worksheets(M)
                End If
            Else
                If CLng(WB.Worksheets(N).Name) < CLng(WB.Worksheets(M).Name) Then
                    WB.Worksheets(N).Move before:=WB.Worksheets(M)
                End If
            End If
        End If
    Next N
Next M

SortWorksheetsByName = True

End Function

Sub Cleanup()
Dim vaNames As Variant

Application.DisplayAlerts = False

vaNames = Array("Sheet1", "Sheet2", "Sheet3")
Worksheets(vaNames).Delete

Application.DisplayAlerts = True


End Sub


Function ScrubInvalidChars(ByVal strIn As String) As String
    Dim objRegex As Object
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Global = True
        .Pattern = "[\<\>\*\\\/\?|]"
        ScrubInvalidChars = .Replace(strIn, "")
    End With
End Function




