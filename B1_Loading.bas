Attribute VB_Name = "B1_Loading"
Sub MainLoadingLoop(ByRef FilesList, ByRef nb_sheets)

    Dim wk As Workbook
    Dim pctCompl As Double
    Dim counter As Integer
    'Dim FilesList As Variant
    Dim table As ListObject
    Set table = INTERNALS.ListObjects("file_to_load")
    Application.DisplayAlerts = False
    Dim VnamesRange As Range
    Dim Vnames As Variant
    Dim curr_col_num As Double
    Dim curr_col_nrows As Long
    Dim Unique_ColumnData As Variant
    Dim Lookup_names As Variant
    Dim Lookup_colnum As Variant
    Dim Lookup_expectedtype As Variant
    Dim empty_column_count As Integer
    Dim ColumnOrder() As String
    'FilesList = GetFilesList
    
    Lookup_names = Application.Transpose(INTERNALS.ListObjects("attributes").ListColumns("trimmed_raw_name").DataBodyRange)
    Lookup_colnum = Application.Transpose(INTERNALS.ListObjects("attributes").ListColumns("DBB_col").DataBodyRange)
    Lookup_expectedtype = Application.Transpose(INTERNALS.ListObjects("attributes").ListColumns("type").DataBodyRange)
    
    counter = 1
    For Each FILE In FilesList
    
        Application.ScreenUpdating = False
        ' set ReadOnly=False
        SetAttr FILE, vbNormal
        
        Set wk = Workbooks.Open(Filename:=FILE, corruptload:=xlRepairFile)
        'wk.Windows(1).Visible = False
        
        'detect the type of the system (flawed)
        
        'table.ListColumns("info_sys").DataBodyRange(counter) = CheckColumnNames(wk.Worksheets(1).Range("A1:S1"))
        
        'determine the reordering of the columns (in variable "reordering")
        Set VnamesRange = wk.Worksheets(1).Range(wk.Worksheets(1).Cells(1, 1), wk.Worksheets(1).Cells(1, wk.Worksheets(1).Cells(1, wk.Worksheets(1).Columns.Count).End(xlToLeft).column))
        Vnames = Application.Transpose(Application.Transpose(VnamesRange))
        
        empty_column_count = 0
        
        If VarType(Vnames) > 8000 Then
            ReDim ColumnOrder(UBound(Vnames))
            For i = LBound(Vnames) To UBound(Vnames)
                If Vnames(i) = "" Or Vnames(i) = " " Then
                    empty_column_count = empty_column_count + 1
                    If empty_column_count > 1 Then
                        table.ListColumns("more_than_one_empty_column").DataBodyRange(counter).value = True
                        Exit For
                    End If
                Else
                    curr_col_num = Application.Match(Trim(Vnames(i)), Lookup_names, 0)
                    'Debug.Print Vnames(i)
                    If VarType(curr_col_num) = vbError Then
                        'Debug.Print wk.Name & ": " & VnamesRange(i).value
                        table.ListColumns("unidentified_fields").DataBodyRange(counter) = table.ListColumns("unidentified_fields").DataBodyRange(counter) & "," & VnamesRange(i).value
                    Else
                        ColumnOrder(i - 1) = CStr(Lookup_colnum(curr_col_num))
                        Debug.Print Vnames(i) & " " & ColumnOrder(i - 1)
                        
                        'On vérifie le type des données de la colonne

                        curr_col_nrows = wk.Worksheets(1).Cells(wk.Worksheets(1).Rows.Count, VnamesRange(i).column).End(xlUp).Row
                        Data = Application.Transpose(VnamesRange(i).Offset(1, 0).Resize(RowSize:=curr_col_nrows - 1))
                        
                        Call CheckType(Data, Lookup_expectedtype(curr_col_num))
                        
                        
                        
                    End If
                End If
            Next i
            reordering = Join(ColumnOrder, "|")
            
            table.ListColumns("reordering").DataBodyRange(counter).value = Left(reordering, Len(reordering) - 1)
            table.ListColumns("required_fields_ok").DataBodyRange(counter).value = ((InStr(reordering, "1|") > 0) And (InStr(reordering, "2|") > 0) And (InStr(reordering, "3|") > 0))
            
            'check columns data type:
            
            
            
            
        Else
            table.ListColumns("more_than_one_empty_column").DataBodyRange(counter).value = True
        End If
        
        
        
        
        
        
        wk.Close SaveChanges:=False
        Set wk = Nothing
        Erase ColumnOrder
        reordering = ""
        
        ' set ReadOnly=True
        'SetAttr FILE, vbReadOnly
        
        'calcul de la progression de la tâche:
        pctCompl = (counter + 1) / (UBound(FilesList) + 1)
        Application.StatusBar = "Progression chargement: étape (2/2) " & (counter + 1) & " of " & UBound(FilesList) + 1 & ": " & Format(pctCompl, "percent")
        
        Application.ScreenUpdating = True
        counter = counter + 1
    Next FILE
    
End Sub


Function ConformableFileName(Filename As String) As Boolean
    ConformableFileName = Filename Like "#_*" Or _
                          Filename Like "##_*" Or _
                          Filename Like "#&#_*" Or _
                          Filename Like "#&##_*" Or _
                          Filename Like "##&##_*" Or _
                          Filename Like "[A-Z]_*" Or _
                          Filename Like "[A-Z][A-Z]_*" Or _
                          Filename Like "[A-Z]#_*" Or _
                          Filename Like "[A-Z]##_*"
    'ConformableFileName = ConformableFileName * Not Filename Like "[!0-9,A-Z]_"
End Function


Sub test()
    Debug.Print ConformableFileName("1_Barbay_Baud_medicaments_2015_brut.xlsx")
End Sub

Function HowManySheets(ByRef FilesList) As Variant
    Dim wk As Workbook
    
    Dim table As ListObject
    Dim nbsheets() As Variant

    Dim nbusedcells As Double
        nbusedcells = 0

    ReDim nbsheets(UBound(FilesList) - LBound(FilesList))
    
    Dim pctCompl As Double
        
    Application.DisplayAlerts = False
    
    Set table = INTERNALS.ListObjects("have_several_tabs")
    table.DataBodyRange.ClearContents
    Set r = table.Range.Rows(2).Offset(-1, 0).Resize(3)
    table.Resize r
    
    Dim counter As Integer
    Dim counter_sev_sheets As Integer
    counter = LBound(FilesList)
    counter_sev_sheets = 0
    For Each FILE In FilesList
    
        Application.ScreenUpdating = False

        Set wk = Workbooks.Open(Filename:=FILE, corruptload:=xlRepairFile)
        wk.Windows(1).Visible = False
        
        'just to export the columns name in a sheet. mest creat sheet TITLES first to use.
        'wk.Worksheets(1).Range("A1:S1").Copy Destination:=Workbooks("Prétraitement_Données.xlsb").Worksheets("TITLES").Range("A" & counter + 1 & ":S" & counter + 1)
        
        nbsheets(counter) = wk.Worksheets.Count
        
        'substract empty sheet to the number of sheet count
        If nbsheets(counter) > 1 Then
            For i = 2 To wk.Worksheets.Count
                If Application.WorksheetFunction.CountA(wk.Worksheets(i).Cells) = 0 Then
                    nbsheets(counter) = nbsheets(counter) - 1
                End If
            Next i
        End If
        
        'report the sheets with count>1 to table "have_several_tabs"
        If nbsheets(counter) > 1 Then
           counter_sev_sheets = counter_sev_sheets + 1
           table.DataBodyRange(counter_sev_sheets, 1) = counter + 1
           table.DataBodyRange(counter_sev_sheets, 2) = wk.Name
        End If
        'report the sheets count to table "file_to_load"
        INTERNALS.ListObjects("file_to_load").ListColumns("n_sheets").DataBodyRange(counter + 1).value = nbsheets(counter)
        
        
        wk.Close SaveChanges:=False
        Set wk = Nothing
        
        pctCompl = (counter + 1) / (UBound(FilesList) + 1)
        Application.StatusBar = "Progression chargement:  étape (1/2) " & (counter + 1) & " of " & UBound(FilesList) + 1 & ": " & Format(pctCompl, "percent")
            
        Application.ScreenUpdating = True
        
        counter = counter + 1
    Next FILE
    Application.DisplayAlerts = True
    HowManySheets = nbsheets
End Function

Function GetStats(FilesList)
    counter = 1
    For Each FILE In FileList
        Application.ScreenUpdating = False

        Set wk = Workbooks.Open(Filename:=FILE, corruptload:=xlRepairFile)
        wk.Windows(1).Visible = False
    
    Next FILE
End Function
