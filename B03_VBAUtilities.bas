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
             .InsertLines LineNum, vbTab & "If Target.Count>10000 Then Exit Sub"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "LastValueSelected = Application.Transpose(Target.value)"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "LastCommentsSelected = GetComments(Target)"
            
            'Activate
            LineNum = .CreateEventProc("Activate", "Worksheet")
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "Call AddToCellMenu"
            
            'Deactivate
            LineNum = .CreateEventProc("Deactivate", "Worksheet")
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "Call DeleteFromCellMenu"
        End With
End Sub

Sub RemoveEventsProcedure(WorksheetToClean As Worksheet)

Dim activeIDE As Object 'VBProject
Set activeIDE = WorksheetToClean.Parent.VBProject

Dim Element As VBComponent

Dim LineCount As Integer
For Each Element In activeIDE.VBComponents
    If Element.Name = WorksheetToClean.Name Then    'change name if necessary
        LineCount = Element.CodeModule.CountOfLines
        Element.CodeModule.DeleteLines 1, LineCount
    End If
Next

End Sub

Sub SetWsName(ws As Worksheet, NewName As String)

    ws.Parent.VBProject.VBComponents(ws.CodeName).Name = NewName

End Sub
