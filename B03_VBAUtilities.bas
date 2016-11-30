Attribute VB_Name = "B03_VBAUtilities"
Sub CreateEventsProcedure(WorksheetToInject As Worksheet)
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim LineNum As Long
        Const DQUOTE = """" ' one " character

        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents(WorksheetToInject.CodeName)
        Set CodeMod = VBComp.CodeModule
        
        With CodeMod
            'OnChange
            LineNum = .CreateEventProc("Change", "Worksheet")
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "Call RegisterChange(Target)"
            'OnSelectionChange
            LineNum = .CreateEventProc("SelectionChange", "Worksheet")
            LineNum = LineNum + 1
            '.InsertLines LineNum, vbTab & "If Target.Cells.Count > 10 Then rly = MsgBox" & DQUOTE & "Éditer " & Target.Cells.Count & " cellules à la fois? " & DQUOTE & ", vbyesno"
            'LineNum = LineNum + 1
            '.InsertLines LineNum, vbTab & "LastValueSelected = CStr(Target(1).value)"
            'LineNum = LineNum + 1
            '.InsertLines LineNum, vbTab & "LastValueSelected = CStr(Target(1).value)"
            ' LineNum = LineNum + 1
            '.InsertLines LineNum, vbTab & "Else"
            'LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "LastValueSelected = Application.Transpose(Target.value)"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "LastCommentsSelected = GetComments(Target)"
        End With
End Sub



Sub SetWsName(ws As Worksheet, NewName As String)

    ws.Parent.VBProject.VBComponents(ws.CodeName).Name = NewName

End Sub
