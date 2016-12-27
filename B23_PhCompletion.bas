Attribute VB_Name = "B23_PhCompletion"


Sub GetPHARMINDEX(control As IRibbonControl)
    Call DefGlobal
    Dim FileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    'code to import PHARMINDEX
    FileName = SelectFile(Many:=False, Target:="PharmIndex")
    If FileName = "" Then Exit Sub
    Set wb = Workbooks.Open(FileName:=FileName, corruptload:=xlRepairFile)
    wb.Worksheets.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    wb.Close
    Set ws = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ws.Name = PHAUNI_SH.Name
    'Complete fields with new info from PHARMINDEX
    Call Completion_DB_To_Unique_Vals(Worksheets(PHAUNI_SH.Name), Worksheets(ws.Name))
End Sub

Sub CommitEdits(control As IRibbonControl)
    Call DefGlobal
    
    If Not CorrectlyFilled(Worksheets(PHAUNI_SH.Name)) Then If MsgBox("Un ou plusieurs champs ne sont pas renseignés (rouge ou blanc)." & vbNewLine & "Continuer?", vbYesNo) = vbNo Then Exit Sub
    'Call VerifyCorrectlyFilled(Worksheets(PHAUNI_SH.Name))
    
    Call Completion_DB_To_Unique_Vals(Worksheets(PHARMA_SH.Name), Worksheets(PHAUNI_SH.Name), True)
    Call MergeSheets
    ActiveWorkbook.Worksheets(DATA_SH.Name).visible = True
    ActiveWorkbook.Worksheets(DATA_SH.Name).Select
    Call CleanNewPharmacodes(Worksheets(PHAUNI_SH.Name))
    Call UpdateStage(5)
End Sub
    

Sub Extract_Unique_Vals(ws As Worksheet)
    'creates a list of unique rows with problematic pharmacode to process
    
    Dim ws_uniquevals As Worksheet
    Dim ColsToKeep As String
    Dim DelRange As Range
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ws.Copy After:=Worksheets(ThisWorkbook.Worksheets.Count)
    Set ws_uniquevals = ActiveSheet
    ws_uniquevals.Name = PHAUNI_SH.Name
    
    'Remove old events from sheet
    Call RemoveEventsProcedure(ws_uniquevals)
    'Add new events to sheet
    Call CreateEventsForPharmacodeCompletion(ws_uniquevals)
    
    'Remove unneeded columns
    ColsToKeep = Join(Array("YEAR_OF_ANALYSIS", "EMS_CODE", "PHARMACIST", "pharmacode", "designation"), "|")
    Set DelRange = Nothing
    For Each column In ws_uniquevals.UsedRange.Columns
        If InStr(ColsToKeep, column.Cells(1).value) = 0 Then
            If DelRange Is Nothing Then Set DelRange = column.EntireColumn Else Set DelRange = Union(DelRange, column.EntireColumn)
        End If
    Next
    DelRange.EntireColumn.Delete
    
    'Keep only unique values
    ws_uniquevals.UsedRange.RemoveDuplicates Columns:=Array(1, 3, 4, 5), Header:=xlYes
    
    'Add fields from pharmindex table
    ws_uniquevals.Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1).Resize(1, INTERNALS.ListObjects("PHARMINDEX_attributes").ListColumns(1).DataBodyRange.Count) = Application.Transpose(INTERNALS.ListObjects("PHARMINDEX_attributes").ListColumns(1).DataBodyRange)
    
    'sort values by designation
    ws_uniquevals.UsedRange.Sort Key1:=ws_uniquevals.Range("1:1").Find(What:="designation").Offset(1, 0), Order1:=xlAscending, Header:=xlYes
    
    ws_uniquevals.Range("A1").AutoFilter
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub



Sub Completion_DB_To_Unique_Vals(UV_ws As Worksheet, DB_ws As Worksheet, Optional OnlyPharmacode As Boolean = False)
    Dim UV_designations As Variant
    Dim DB_designations As Variant
    Dim All_DB_designations As String
    Dim Index As Long
    Dim MatchIndex As Long
    Dim MatchPos As String
    Dim Strlength As String

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    UV_designations = Application.Transpose(UV_ws.UsedRange.Rows(1).Find("designation").Offset(1, 0).Resize(UV_ws.UsedRange.Rows.Count - 1, 1))
    DB_designations = Application.Transpose(DB_ws.UsedRange.Rows(1).Find("designation").Offset(1, 0).Resize(DB_ws.UsedRange.Rows.Count - 1, 1))
    
    All_DB_designations = Join(DB_designations, "|")
    For i = LBound(UV_designations) To UBound(UV_designations)
        MatchPos = InStr(1, All_DB_designations, UV_designations(i), vbTextCompare)
        If MatchPos > 0 Then
            Index = i + 1
            Strlength = 0
            If OnlyPharmacode Then
                Dim DB_PhCol As String
                Dim UV_PhCol As String
                DB_PhCol = IncCol("A", DB_ws.Range("1:1").Find("PHCODE").column - 1)
                UV_PhCol = IncCol("A", UV_ws.Range("1:1").Find("pharmacode").column - 1)
                For j = LBound(DB_designations) To UBound(DB_designations)
                    Strlength = Strlength + Len(DB_designations(j))
                    If Strlength >= MatchPos Then
                        MatchIndex = j + 1
                        'report values
                        DB_ws.Range(DB_PhCol & MatchIndex).Copy UV_ws.Range(UV_PhCol & Index)
                        'UV_ws.Range(UV_PhCol & Index).Cells.Interior.ColorIndex = 4
                        Exit For
                    End If
                Next j
            Else
                For j = LBound(DB_designations) To UBound(DB_designations)
                    Strlength = Strlength + Len(DB_designations(j))
                    If Strlength >= MatchPos Then
                        MatchIndex = j + 1
                        'report values
                        UV_ws.Range(UV_ws.Cells(Index, PHAUNI_SH.HOffset + 1), UV_ws.Cells(Index, PHAUNI_SH.HOffset + DB_ws.UsedRange.Columns.Count)) = _
                                                DB_ws.Range("A" & MatchIndex & ":" & IncCol("A", DB_ws.UsedRange.Columns.Count) & MatchIndex).value
                        'mark as filled
                        UV_ws.Range(UV_ws.Cells(Index, PHAUNI_SH.HOffset + 1), UV_ws.Cells(Index, PHAUNI_SH.HOffset + DB_ws.UsedRange.Columns.Count)).Cells.Interior.ColorIndex = 4
                        UV_ws.Rows(Index).EntireRow.Hidden = True
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

Function CorrectlyFilled(ws As Worksheet) As Boolean
    CorrectlyFilled = True
    With ws
        For Each cell In .UsedRange.Offset(PHAUNI_SH.VOffset, PHAUNI_SH.HOffset).Resize(.UsedRange.Rows.Count - PHAUNI_SH.VOffset, .UsedRange.Columns.Count - PHAUNI_SH.HOffset)
            If cell.Interior.ColorIndex = 3 Or cell.Interior.ColorIndex = xlNone Or cell.Interior.ColorIndex = 45 Then
                CorrectlyFilled = False
                Exit Function
            End If
        Next cell
    End With
End Function


Sub CleanNewPharmacodes(ws As Worksheet)
    
    Dim lastRow As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Remove events from sheet
    Call RemoveEventsProcedure(ws)
    
    'Remove known values of PHARMINDEX
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row
    For iCntr = lastRow To 1 Step -1
        If Rows(iCntr).Hidden = True Then Rows(iCntr).EntireRow.Delete
    Next
    
    'Remove formats
    ws.Cells.ClearFormats
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
