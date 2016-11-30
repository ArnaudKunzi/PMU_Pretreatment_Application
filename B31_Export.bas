Attribute VB_Name = "B31_Export"
Sub Export()
    Call DefGlobal
    Dim DispatchFiles As Boolean
    Dim SaveinSeparateSheets As Boolean
    Dim TraceChanges As Boolean
    Dim SheetsToExport As String
    
    
    DispatchFiles = PARAM_TABLE.Columns(1).Find("DispatchFiles").Offset(0, 1).value
    SaveinSeparateSheets = PARAM_TABLE.Columns(1).Find("SaveinSeparateSheets").Offset(0, 1).value
    
    If DispatchFiles & (Not SaveinSeparateSheets) Then
        Call MergeExport
    ElseIf (Not DispatchFiles) & SaveinSeparateSheets Then
        If Evaluate("ISREF('" & InPh_colname & "'!A1)") Then GoTo Handler
continue:
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = InPh_colname
        Call MoveRowsToSheet("InvalidPharmacodes", 1, Worksheets(DataSheetName), Worksheets(InPh_colname))
    End If
    
    
    If SaveinSeparateSheets Then
        SheetsToExport = Join(DataSheetName, InPh_colname)
    Else
        SheetsToExport = DataSheetName
    End If
    
    'If TraceChanges Then Call
    
    
    
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




