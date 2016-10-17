Attribute VB_Name = "B0_SheetUtilities"
Sub SaveFilesList(ByRef FilesList)
    Dim table As ListObject
    Dim Path As Range
    Dim r As Range
    
    Set table = INTERNALS.ListObjects("file_to_load")
    Set Path = INTERNALS.ListObjects("path").ListColumns("path").DataBodyRange
    
    Path(1).value = Left(FilesList(0), InStrRev(FilesList(0), "\"))
    
    table.DataBodyRange.ClearContents
    Set r = table.Range.Rows(2).Offset(-1, 0).Resize(3)
    table.Resize r
    For i = LBound(FilesList) To UBound(FilesList)
        table.ListColumns(1).DataBodyRange(i + 1) = i + 1
        table.ListColumns(2).DataBodyRange(i + 1) = Right(FilesList(i), Len(FilesList(i)) - InStrRev(FilesList(i), "\"))
    Next i
    
End Sub

'Function GetFilesList()
'    Dim table As ListObject
'    Dim path As Range
'    Dim FilesList As Variant
'    Set table = INTERNALS.ListObjects("file_to_load")
'    Set path = INTERNALS.ListObjects("path").ListColumns("path").DataBodyRange
'
'    For i = 1 To table.ListRows.Count '
'        FilesList(i - 1) = path(1).value & table.ListColumns(2).DataBodyRange(i).value
'    Next i
'    GetFilesList = FilesList
'End Function


Sub TransferColumns()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim input_wb As Workbook
    Dim output_wb As Workbook
    
    Dim InpRangetxt As String
    Dim OutRangetxt As String
    
    Dim COffset As Integer
    Dim ROffset As Integer
    
    Dim ColumnOrder As Variant
    Dim CurrentFileColumnOrder As Variant
    Dim OutputColumnOrder() As Integer
    Dim OutputColumnOrder_length As Integer
    Dim InputDataTable As Variant
    Dim FilePath As String
    Dim InputLastRow As Long
    Dim OutputLastRow As Long
    
    Dim DestinationRange As Range
    Dim DestinationColumn As String
    
    Set output_wb = ActiveWorkbook
    COffset = 2
    ROffset = 1
    On Error Resume Next
    Sheets("DATA").Delete
    On Error GoTo 0
    
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "DATA"
    
    ColumnOrders = Application.Transpose(INTERNALS.ListObjects("file_to_load").ListColumns("reordering").DataBodyRange)
    
    'Name the columns:
    For i = 1 To INTERNALS.ListObjects("attributes").ListColumns("DBB_col").DataBodyRange.Rows.Count
        Worksheets("DATA").Cells(1, INTERNALS.ListObjects("attributes").ListColumns("DBB_col").DataBodyRange(i) + COffset).value = INTERNALS.ListObjects("attributes").ListColumns("DBB_name").DataBodyRange(i)
    Next i
    
    FilePath = INTERNALS.ListObjects("path").ListColumns("path").DataBodyRange(1).value
    files = Application.Transpose(INTERNALS.ListObjects("file_to_load").ListColumns("file_to_load").DataBodyRange)
    
    
    
    For i = 0 To INTERNALS.ListObjects("file_to_load").ListColumns("file_to_load").DataBodyRange.Rows.Count - 1
        
        
        Set input_wb = Workbooks.Open(Filename:=FilePath & files(i + 1), corruptload:=xlRepairFile)
        
        'Last row of the output file
        OutputLastRow = Application.Max(output_wb.Worksheets("DATA").Cells(output_wb.Worksheets("DATA").Rows.Count, "A").End(xlUp).Row, _
                                  output_wb.Worksheets("DATA").Cells(output_wb.Worksheets("DATA").Rows.Count, "C").End(xlUp).Row, _
                                  output_wb.Worksheets("DATA").Cells(output_wb.Worksheets("DATA").Rows.Count, "E").End(xlUp).Row)
        'Last row of the input file
        InputLastRow = Application.Max(input_wb.Worksheets(1).Cells(input_wb.Worksheets(1).Rows.Count, "A").End(xlUp).Row, _
                                   input_wb.Worksheets(1).Cells(input_wb.Worksheets(1).Rows.Count, "C").End(xlUp).Row, _
                                   input_wb.Worksheets(1).Cells(input_wb.Worksheets(1).Rows.Count, "E").End(xlUp).Row)
                                   
        'Il faut maintenant intervertir les index et les valeurs de ColumnOrders
        CurrentFileColumnOrder = Split(ColumnOrders(i + 1), "|")
        
        OutputColumnOrder_length = 0
        
        'VBA is RETARDED! we cannot cast a string array to Int to find the max value (which is the size of the new array OutputColumnOrder)
        'so we have to iterate through it to find it... sigh!
        For n = LBound(CurrentFileColumnOrder) To UBound(CurrentFileColumnOrder)
            If CurrentFileColumnOrder(n) <> "" Then
                If OutputColumnOrder_length < CInt(CurrentFileColumnOrder(n)) Then
                    OutputColumnOrder_length = CInt(CurrentFileColumnOrder(n))
                End If
            End If
        Next n
        
        ReDim OutputColumnOrder(1 To OutputColumnOrder_length)
        'Now we can reorder:
        For n = LBound(CurrentFileColumnOrder) To UBound(CurrentFileColumnOrder)
            If CurrentFileColumnOrder(n) <> "" Then
                OutputColumnOrder(CInt(CurrentFileColumnOrder(n))) = n + 1
            End If
        Next n
        
        'copier l'intégralité de la table en mémoire
        InputDataTable = Application.Transpose(input_wb.Worksheets(1).Range("A2:" & IncCol("A", UBound(CurrentFileColumnOrder)) & CStr(InputLastRow)))
        output_wb.Activate
        
        For column = LBound(CurrentFileColumnOrder) To UBound(CurrentFileColumnOrder)
            If OutputColumnOrder(column + 1) <> 0 Then
                DestinationColumn = IncCol("A", column + COffset)
                Set DestinationRange = output_wb.Worksheets("DATA").Range(DestinationColumn & OutputLastRow + ROffset & ":" & DestinationColumn & OutputLastRow + ROffset + InputLastRow - 2)
                DestinationRange = Application.Transpose(Application.index(InputDataTable, OutputColumnOrder(column + 1)))
            End If
        Next column
        
        output_wb.Worksheets("DATA").Range("A1").value = "YEAR_OF_ANALYSIS"
        output_wb.Worksheets("DATA").Range("B1").value = "EMS_CODE"
        output_wb.Worksheets("DATA").Range("A" & ROffset + OutputLastRow & ":A" & ROffset + InputLastRow + OutputLastRow - 2) = Year
        output_wb.Worksheets("DATA").Range("B" & ROffset + OutputLastRow & ":B" & ROffset + InputLastRow + OutputLastRow - 2) = Left(input_wb.Name, InStr(input_wb.Name, "_") - 1)
        
        input_wb.Close SaveChanges:=False
        Set input_wb = Nothing
        
        
        
    Next
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub



Private Function StringArrayToIntArray(ByRef values()) As Integer()
    Dim lIndx As Long, lLwrBnd As Long, lUprBnd As Long
    Dim adRtnVals() As Integer
    lLwrBnd = LBound(values)
    lUprBnd = UBound(values)
    ReDim adRtnVals(lLwrBnd To lUprBnd) As Integer
    For lIndx = lLwrBnd To lUprBnd
        adRtnVals(lIndx) = CInt(values(lIndx))
    Next
    StringArrayToIntArray = adRtnVals
End Function



Function IncCol(column As String, IncrementStep As Integer) As String
    
    Dim reminder As Double
    Dim nloops As Long
    reminder = Asc(column) + IncrementStep
    nloops = 1
    If (Asc(column) + IncrementStep) < Asc("Z") Then
        IncCol = Chr(Asc(column) + IncrementStep)
    Else
        Do While (reminder > Asc("Z"))
            reminder = reminder - (Asc("Z") - Asc("A"))
            IncCol = IncrementColumn & "A"
            nloops = nloops + 1
            If nloops > 10000 Then
                Exit Do
            End If
        Loop
        IncCol = IncrementColumn & Chr(reminder)
    End If
End Function
