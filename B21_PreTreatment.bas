Attribute VB_Name = "B21_PreTreatment"
Sub StartPreTreatment(control As IRibbonControl)
    Call DefGlobal
    
    Dim StatusColumn As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Check if RAPPORT sheet exists, if not, create it.
    If Worksheets(REPORT_SH.Name) Is Nothing Then Call Refresh(Nothing)
           
    'find the column "Status"
    StatusColumn = IncCol("A", Worksheets(REPORT_SH.Name).Range("1:1").Find("Status").column - 1)
    
retry:
    If Not Worksheets(REPORT_SH.Name).Range(StatusColumn & ":" & StatusColumn).Find("WARNING") Is Nothing Then
        Dim Choice1 As Variant
            Choice1 = MsgBox("les status des fichiers médicaments n'ont pas été résolus. Merci de les résoudres puis d'actualiser le rapport avant de réessayer.", vbAbortRetryIgnore, "Status invalides")
        If Choice1 = 3 Then  'abort
            Exit Sub
        ElseIf Choice1 = 4 Then
            Call Refresh(Nothing)
            GoTo retry 'YES it is a dreaded GoTo!
        Else 'ignore
            MsgBox "La conformité des données n'est pas garantie lorsque les status ne sont pas résolus.", vbExclamation
        End If
        
    End If
    
    
    Call TransferColumns(PHARMA_SH.Name)
    
'   DEPRECATED
'    If PARAM_TABLE.Columns(1).Find("DispatchFiles").Offset(0, 1).value Then
'
'        If SheetExists(PHARMA_SH.Name) Then GoTo Handler
'continue:
'        Sheets.Add(After:=Sheets(Sheets.Count)).Name = PHARMA_SH.Name
'        Worksheets(PHARMA_SH.Name).Tab.ColorIndex = EXPORTCOLOR
'        Call SetWsName(Worksheets(PHARMA_SH.Name), PHARMA_SH.Name)
'
'        PARAM_TABLE.Columns(1).Find("TbtnToggleSeparateByPhStatus").Offset(0, 1).value = True
'        Call SplitSheets
'        Call ShowOnlyCustomTabs
'
'
'    End If

    
    
    'Change the ribbon focus and ribbon configuration?
    'something like [back][filters][]
    
    
    'Call AddToCellMenu
    
    Call SplitSheets
    Call Extract_Unique_Vals(Worksheets(PHARMA_SH.Name))
    ActiveWorkbook.Worksheets(PHARMA_SH.Name).visible = False
    ActiveWorkbook.Worksheets(DATA_SH.Name).visible = False
    
    Call UpdateStage(3)
    
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
'DEPRECATED
'Exit Sub
'Handler:
'    Dim choice2 As Integer
'    Dim iter As Integer
'    choice2 = MsgBox("Il y a déjà une feuille InvalidPharmacodes en traitment." & Chr(10) & _
'           "Écraser la feuille existante?", vbYesNoCancel)
'
'    Select Case choice2
'        Case vbYes
'            Sheets(PHARMA_SH.Name).Delete
'            GoTo continue
'        Case vbNo
'            iter = 1
'            Do
'                iter = iter + 1
'            Loop While SheetExists(PHARMA_SH.Name & iter) And iter <= 10
'            InPh_colname = InPh_colname & iter
'
'            GoTo continue
'
'        Case vbCancel
'            Exit Sub
'    End Select
End Sub



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
    Dim filepath As String
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
    Sheets(DATA_SH.Name).Delete
    On Error GoTo 0
    
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = DATA_SH.Name
    Worksheets(DATA_SH.Name).Tab.ColorIndex = EXPORTCOLOR
    Call SetWsName(Worksheets(DATA_SH.Name), DATA_SH.Name)
    
    ColumnOrders = Application.Transpose(INTERNALS.ListObjects("file_to_load").ListColumns("reordering").DataBodyRange)
    
    'Name the columns:
    For i = 1 To INTERNALS.ListObjects("attributes").ListColumns("DBB_col").DataBodyRange.Rows.Count
        Worksheets(DATA_SH.Name).Cells(1, INTERNALS.ListObjects("attributes").ListColumns("DBB_col").DataBodyRange(i) + COffset).value = INTERNALS.ListObjects("attributes").ListColumns("DBB_name").DataBodyRange(i)
    Next i
    
    filepath = INTERNALS.ListObjects("path").ListColumns("path").DataBodyRange(1).value
    files = Application.Transpose(INTERNALS.ListObjects("file_to_load").ListColumns("file_to_load").DataBodyRange)
    
    
    
    For i = 0 To INTERNALS.ListObjects("file_to_load").ListColumns("file_to_load").DataBodyRange.Rows.Count - 1
        
        
        Set input_wb = Workbooks.Open(FileName:=filepath & files(i + 1), corruptload:=xlRepairFile)
        
        
        'Last row of the output file
        With output_wb.Worksheets(DATA_SH.Name)
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
                Set DestinationRange = output_wb.Worksheets(DATA_SH.Name).Range(DestinationColumn & OutputLastRow + ROffset & ":" & DestinationColumn & OutputLastRow + ROffset + InputLastRow - 2)
                DestinationRange = Application.Transpose(Application.Index(InputDataTable, OutputColumnOrder(column + 1)))
            
           'if column is a PHARMACODE column and pharmacode detection is enabled, flag rows with invalid pharmacodes
                
                If (column + 1) = PharmacodeColumn And PharmacodeDetectionEnabled Then
                    OutputLastCol = COffset + Application.Max(INTERNALS.ListObjects("AttributeTypeAndPlacement").ListColumns("DBB_col").DataBodyRange)
                    
                    output_wb.Worksheets(DATA_SH.Name).Cells(1, OutputLastCol + 1).value = InPh_colname
                    IncorrectPharmacodes = Split(CheckElementsType(Application.Index(InputDataTable, OutputColumnOrder(column + 1)), "PHARMACODE"), ",")
                    For k = LBound(IncorrectPharmacodes) + 1 To UBound(IncorrectPharmacodes) - 1
                        output_wb.Worksheets(DATA_SH.Name).Cells(OutputLastRow + ROffset, OutputLastCol + 1).Offset(IncorrectPharmacodes(k) - 1, 0) = 1
                    Next k
                    output_wb.Worksheets(DATA_SH.Name).Range(IncCol("A", OutputLastCol - 1) & OutputLastRow + ROffset & ":" & IncCol("A", OutputLastCol - 1) & OutputLastRow + ROffset + InputLastRow - 2).SpecialCells(xlCellTypeBlanks).value = 0
                    Set IncorrectPharmacodes = Nothing
                End If
            End If
        Next column
        
        With output_wb.Worksheets(DATA_SH.Name)
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
    
    Call CreateEventsForPreTreatment(Worksheets(DATA_SH.Name))
    AppActivate Application.Caption
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub




Sub MoveRowsToSheet(ByVal IndicatorCol As String, ByVal Criterion As Integer, ByRef InputSheet As Worksheet, ByRef OutputSheet As Worksheet)
    
    Dim lastRow As Long         'Last row with data in InputSheet
    Dim LastCol As Long         'Last column with data in InputSheet

    Dim IndCol As Long          'number of the column IndicatorCol
    
    Dim DataRange As Range      'Data Range from the input sheet
    Dim Atributes As Range      'Row of attributes range
    Dim RowsToMove As Range     'Data to move to the Outputsheet
    
    Application.EnableEvents = False
    
    With InputSheet
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).column
        Set Atributes = .Range(.Cells(1, 1), .Cells(1, LastCol))
        Set DataRange = .Range(.Cells(2, 1), .Cells(lastRow, LastCol))
        
        On Error GoTo Handler
        IndCol = .Cells.Rows(1).Find(IndicatorCol).column

        .Range("A:" & IncCol("A", LastCol)).AutoFilter field:=IndCol, Criteria1:=Criterion
        
        Set RowsToMove = DataRange.SpecialCells(xlCellTypeVisible)
        
        Atributes.Copy OutputSheet.Cells(1, 1).EntireRow
        RowsToMove.Copy OutputSheet.Cells(2, 1).EntireRow
        
        RowsToMove.EntireRow.Delete
        
        .Range("A:" & IncCol("A", LastCol)).AutoFilter
        
    End With
    
    Application.EnableEvents = True
    
Exit Sub
Handler:
        MsgBox "Column " & IndicatorCol & " not found in sheet " & InputSheet.Name
End Sub




Sub MergeSheets()

    Call DefGlobal
    Application.EnableEvents = False
    
    If Not SheetExists(PHARMA_SH.Name) Then Exit Sub
    
    Dim in_LastRow As Long
    Dim in_LastCol As Long
    Dim out_LastRow As Long
    Dim in_ws As Worksheet
    Dim out_ws As Worksheet
    
    Set in_ws = Worksheets(PHARMA_SH.Name)
    Set out_ws = Worksheets(DATA_SH.Name)
    
    in_LastRow = in_ws.Cells(in_ws.Rows.Count, "A").End(xlUp).row
    in_LastCol = in_ws.Cells(1, in_ws.Columns.Count).End(xlToLeft).column
    out_LastRow = out_ws.Cells(out_ws.Rows.Count, "A").End(xlUp).row
    
    in_ws.Range(in_ws.Cells(2, 1), in_ws.Cells(in_LastRow, in_LastCol)).Cut out_ws.Range("A" & out_LastRow + 1)
    
    With Application
        .DisplayAlerts = False
        in_ws.Delete
        .DisplayAlerts = True
        .EnableEvents = True
    End With
End Sub


Sub SplitSheets()

    Call DefGlobal
    Application.EnableEvents = False
    
    If SheetExists(PHARMA_SH.Name) Then Exit Sub
    
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = PHARMA_SH.Name
    Worksheets(PHARMA_SH.Name).Tab.ColorIndex = EXPORTCOLOR
    Call SetWsName(Worksheets(PHARMA_SH.Name), PHARMA_SH.Name)
        
    Call MoveRowsToSheet(PHARMA_SH.Name, 1, Worksheets(DATA_SH.Name), Worksheets(PHARMA_SH.Name))
    
    Call CreateEventsForPreTreatment(Worksheets(PHARMA_SH.Name))
    AppActivate Application.Caption
    Application.EnableEvents = True
End Sub

