Attribute VB_Name = "B31_Export"
Sub Export()
    Call DefGlobal
    Dim DispatchFiles As Boolean
    Dim SaveinSeparateSheets As Boolean
    Dim TrackChanges As Boolean
    Dim SheetsToExport As String
    
    
    DispatchFiles = PARAM_TABLE.Columns(1).Find("DispatchFiles").Offset(0, 1).value
    SaveinSeparateSheets = PARAM_TABLE.Columns(1).Find("SaveinSeparateSheets").Offset(0, 1).value
    
    If DispatchFiles And Not SaveinSeparateSheets Then
        Call MergeSheets
    ElseIf (Not DispatchFiles) And SaveinSeparateSheets Then
        If Evaluate("ISREF('" & InPh_colname & "'!A1)") Then GoTo Handler
continue:
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = InPh_colname
        Worksheets(InPh_colname).Tab.ColorIndex = EXPORTCOLOR
        Call MoveRowsToSheet("InvalidPharmacodes", 1, Worksheets(DataSheetName), Worksheets(InPh_colname))
    End If
    
    
    'If SaveinSeparateSheets Then
    '    SheetsToExport = Array(DataSheetName, InPh_colname)
    'Else
    '    SheetsToExport = DataSheetName
    'End If
    
    'If TrackChanges Then Call
    
    'List of sheets to export (those with tab colored EXPORTCOLOR)
    For Each sheet In Worksheets
        If sheet.Tab.ColorIndex = EXPORTCOLOR Then
            If SheetsToExport = "" Then
                SheetsToExport = sheet.Name
            Else
                SheetsToExport = Join(Array(SheetsToExport, sheet.Name), "|")
            End If
        End If
    Next sheet
    
    
    
    
    
    Call UpdateStage("PreTreatment")
    
Exit Sub
Handler:
    Dim choice2 As Integer
    Dim iter As Integer
    choice2 = MsgBox("Il y a déjà une feuille InvalidPharmacodes en traitment." & Chr(10) & _
           "Écraser la feuille existante?", vbYesNoCancel)
    Select Case choice2
        Case vbYes
            Sheets(InPh_colname).Delete
            GoTo continue
        Case vbNo
            iter = 1
            Do
                iter = iter + 1
            Loop While Evaluate("ISREF('" & InPh_colname & iter & "'!A1)") And iter <= 10
            InPh_colname = InPh_colname & iter
            
            GoTo continue
            
        Case vbCancel
            Exit Sub
    End Select
    
End Sub


