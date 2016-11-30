Attribute VB_Name = "B02_SheetUtilities"
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


Sub TransferColumns(ByVal InPh_colname As String)

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
    Dim OutputLastCol As Long
    
    Dim DestinationRange As Range
    Dim DestinationColumn As String
    
    Dim PharmacodeColumn As Long
    Dim PharmacodeDetectionEnabled As Boolean
    Dim IncorrectPharmacodes As Variant
    
    Set output_wb = ActiveWorkbook
    COffset = 3
    ROffset = 1
    On Error Resume Next
    Sheets(DataSheetName).Delete
    On Error GoTo 0
    
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = DataSheetName
    Call SetWsName(Worksheets(DataSheetName), DataSheetName)
    
    ColumnOrders = Application.Transpose(INTERNALS.ListObjects("file_to_load").ListColumns("reordering").DataBodyRange)
    
    'Name the columns:
    For i = 1 To INTERNALS.ListObjects("attributes").ListColumns("DBB_col").DataBodyRange.Rows.Count
        Worksheets(DataSheetName).Cells(1, INTERNALS.ListObjects("attributes").ListColumns("DBB_col").DataBodyRange(i) + COffset).value = INTERNALS.ListObjects("attributes").ListColumns("DBB_name").DataBodyRange(i)
    Next i
    
    FilePath = INTERNALS.ListObjects("path").ListColumns("path").DataBodyRange(1).value
    files = Application.Transpose(INTERNALS.ListObjects("file_to_load").ListColumns("file_to_load").DataBodyRange)
    
    
    
    For i = 0 To INTERNALS.ListObjects("file_to_load").ListColumns("file_to_load").DataBodyRange.Rows.Count - 1
        
        
        Set input_wb = Workbooks.Open(Filename:=FilePath & files(i + 1), corruptload:=xlRepairFile)
        
        
        'Last row of the output file
        With output_wb.Worksheets(DataSheetName)
        OutputLastRow = Application.Max(.Cells(.Rows.Count, "A").End(xlUp).row, _
                                  .Cells(.Rows.Count, "C").End(xlUp).row, _
                                  .Cells(.Rows.Count, "E").End(xlUp).row)
        End With
        'Last row of the input file
        With input_wb.Worksheets(1)
        InputLastRow = Application.Max(.Cells(.Rows.Count, "A").End(xlUp).row, _
                                   .Cells(.Rows.Count, "C").End(xlUp).row, _
                                   .Cells(.Rows.Count, "E").End(xlUp).row)
                                   
        End With
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
        'Trim la table
        For k = LBound(InputDataTable, 1) To UBound(InputDataTable, 1)
            For p = LBound(InputDataTable, 2) To UBound(InputDataTable, 2)
                On Error Resume Next
                InputDataTable(k, p) = Trim(InputDataTable(k, p))
            Next p
        Next k
        
        output_wb.Activate
        
        PharmacodeColumn = INTERNALS.ListObjects("AttributeTypeAndPlacement").ListColumns(1).DataBodyRange.Find("pharmacode").Offset(0, 1).value
        PharmacodeDetectionEnabled = PARAM_TABLE.Columns(1).Find("CheckPharmacodes").Offset(0, 1).value
                
        For column = LBound(CurrentFileColumnOrder) To UBound(CurrentFileColumnOrder)
            If OutputColumnOrder(column + 1) <> 0 Then
                DestinationColumn = IncCol("A", column + COffset)
                Set DestinationRange = output_wb.Worksheets(DataSheetName).Range(DestinationColumn & OutputLastRow + ROffset & ":" & DestinationColumn & OutputLastRow + ROffset + InputLastRow - 2)
                DestinationRange = Application.Transpose(Application.index(InputDataTable, OutputColumnOrder(column + 1)))
            
           'if column is a PHARMACODE column and pharmacode detection is enabled, flag rows with invalid pharmacodes
                
                If (column + 1) = PharmacodeColumn And PharmacodeDetectionEnabled Then
                    OutputLastCol = COffset + Application.Max(INTERNALS.ListObjects("AttributeTypeAndPlacement").ListColumns("DBB_col").DataBodyRange)
                    
                    output_wb.Worksheets(DataSheetName).Cells(1, OutputLastCol + 1).value = InPh_colname
                    IncorrectPharmacodes = Split(CheckElementsType(Application.index(InputDataTable, OutputColumnOrder(column + 1)), "PHARMACODE"), ",")
                    For k = LBound(IncorrectPharmacodes) + 1 To UBound(IncorrectPharmacodes) - 1
                        output_wb.Worksheets(DataSheetName).Cells(OutputLastRow + ROffset, OutputLastCol + 1).Offset(IncorrectPharmacodes(k) - 1, 0) = 1
                    Next k
                    output_wb.Worksheets(DataSheetName).Range(IncCol("A", OutputLastCol - 1) & OutputLastRow + ROffset & ":" & IncCol("A", OutputLastCol - 1) & OutputLastRow + ROffset + InputLastRow - 2).SpecialCells(xlCellTypeBlanks).value = 0
                    Set IncorrectPharmacodes = Nothing
                End If
            End If
        Next column
        
        With output_wb.Worksheets(DataSheetName)
            .Range("A1").value = "YEAR_OF_ANALYSIS"
            .Range("B1").value = "EMS_CODE"
            .Range("C1").value = "PHARMACIST"
            .Range("A" & ROffset + OutputLastRow & ":A" & ROffset + InputLastRow + OutputLastRow - 2) = Year
            .Range("B" & ROffset + OutputLastRow & ":B" & ROffset + InputLastRow + OutputLastRow - 2) = Left(input_wb.Name, InStr(input_wb.Name, "_") - 1)
            .Range("C" & ROffset + OutputLastRow & ":C" & ROffset + InputLastRow + OutputLastRow - 2) = Mid(input_wb.Name, _
                                                                    InStr(input_wb.Name, "_") + 1, _
                                                                        InStr(InStr(input_wb.Name, "_") + 1, _
                                                                                input_wb.Name, "_") - InStr(input_wb.Name, "_") - 1)
        End With
        
        input_wb.Close SaveChanges:=False
        Set input_wb = Nothing
        
        
        
    Next
    
    Call CreateEventsProcedure(Worksheets(DataSheetName))
    
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

Sub testIncCol()
    Debug.Print IncCol("A", 26)
End Sub

Function IncCol(ByVal column As String, ByVal IncrementStep As Integer) As String
    
    Dim reminder As Double
    Dim nloops As Long
    Dim IncrementColumn As String
    reminder = Asc(column) + IncrementStep
    nloops = 1
    If (Asc(column) + IncrementStep) < Asc("Z") Then
        IncCol = Chr(Asc(column) + IncrementStep)
    Else
        IncrementColumn = ""
        Do While (reminder > Asc("Z"))
            reminder = reminder - (Asc("Z") - Asc("A"))
            IncrementColumn = IncrementColumn & "A"
            nloops = nloops + 1
            If nloops > 10000 Then
                Exit Do
            End If
        Loop
        IncCol = IncrementColumn & Chr(reminder)
    End If
End Function

Function GetUniqueValues(ByRef Data)
    Dim temp As Variant
    Dim obj As Object
    Set obj = CreateObject("scripting.dictionary")
    For i = LBound(Data) To UBound(Data)
        obj(Data(i) & "") = ""
    Next
    GetUniqueValues = obj.keys
End Function

Sub testMoveRowsToSheet()
    Call MoveRowsToSheet("Move", 1, Worksheets(DataSheetName), Worksheets("Pharmacheck"))
End Sub


Sub MoveRowsToSheet(ByVal IndicatorCol As String, ByVal Criterion As Integer, ByRef InputSheet As Worksheet, ByRef OutputSheet As Worksheet)
    
    Dim LastRow As Long         'Last row with data in InputSheet
    Dim LastCol As Long         'Last column with data in InputSheet

    Dim IndCol As Long          'number of the column IndicatorCol
    
    Dim DataRange As Range      'Data Range from the input sheet
    Dim Atributes As Range      'Row of attributes range
    Dim RowsToMove As Range     'Data to move to the Outputsheet
    
    With InputSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).column
        Set Atributes = .Range(.Cells(1, 1), .Cells(1, LastCol))
        Set DataRange = .Range(.Cells(2, 1), .Cells(LastRow, LastCol))
        
        On Error GoTo Handler
        IndCol = .Cells.Rows(1).Find(IndicatorCol).column

        .Range("A:" & IncCol("A", LastCol)).AutoFilter field:=IndCol, Criteria1:=Criterion
        
        Set RowsToMove = DataRange.SpecialCells(xlCellTypeVisible)
        
        Atributes.Copy OutputSheet.Cells(1, 1).EntireRow
        RowsToMove.Copy OutputSheet.Cells(2, 1).EntireRow
        
        RowsToMove.EntireRow.Delete
        
        .Range("A:" & IncCol("A", LastCol)).AutoFilter
        
    End With
    
Exit Sub
Handler:
        MsgBox "Column " & IndicatorCol & " not found in sheet " & InputSheet.Name
End Sub


















Public Function NumericOnly(ByVal s As String) As String
    Dim s2 As String
    Dim replace_hyphen As String
 
    Static re As RegExp
    If re Is Nothing Then Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "[^0-9]" 'includes space, if you want to exclude space "[^0-9]"
    's2 = re.Replace(s, vbNullString)

    NumericOnly = re.Replace(s2, replace_hyphen)
End Function

Sub test()

    s = "1,2,3,4"
    
    Debug.Print NumericOnly(s)
    
End Sub
