Attribute VB_Name = "B23_PhCompletion"
Sub testeuv()
    Call DefGlobal
    Call Extract_Unique_Vals(Worksheets(InPh_colname))
    
End Sub

Sub testcdtuv()
    Call DefGlobal
    Call Completion_DB_To_Unique_Vals(Worksheets("EntriesToComplete"), Worksheets("DB_PHARMINDEX_Extract"))
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
    ws_uniquevals.Name = "EntriesToComplete"
    
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
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub



Sub Completion_DB_To_Unique_Vals(UV_ws As Worksheet, DB_ws As Worksheet)
    Dim UV_designations As Variant
    Dim DB_designations As Variant
    Dim All_DB_designations As String
    'Dim MatchIndex As Long
    Dim MatchPos As String
    Dim Strlength As String
    Dim hOffset As Integer
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    hOffset = 5
    UV_designations = Application.Transpose(UV_ws.UsedRange.Rows(1).Find("designation").Offset(1, 0).Resize(UV_ws.UsedRange.Rows.Count - 1, 1))
    DB_designations = Application.Transpose(DB_ws.UsedRange.Rows(1).Find("designation").Offset(1, 0).Resize(DB_ws.UsedRange.Rows.Count - 1, 1))
    
    All_DB_designations = Join(DB_designations, "|")
    For i = LBound(UV_designations) To UBound(UV_designations)
        MatchPos = InStr(All_DB_designations, UV_designations(i))
        If MatchPos > 0 Then
            Strlength = 0
            For j = LBound(DB_designations) To UBound(DB_designations)
                Strlength = Strlength + Len(DB_designations(j))
                If Strlength >= MatchPos Then
                    MatchIndex = j + 1
                    UV_ws.Range(UV_ws.Cells(i + 1, hOffset + 1), UV_ws.Cells(i + 1, hOffset + DB_ws.UsedRange.Columns.Count)) = _
                                            DB_ws.Range("A" & MatchIndex & ":" & IncCol("A", DB_ws.UsedRange.Columns.Count) & MatchIndex).value
                    UV_ws.Range(UV_ws.Cells(i + 1, hOffset + 1), UV_ws.Cells(i + 1, hOffset + DB_ws.UsedRange.Columns.Count)).Cells.Interior.ColorIndex = 4
                    UV_ws.Rows(i + 1).EntireRow.Hidden = True
                    Exit For
                End If
            Next j
        End If
    Next i
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub
