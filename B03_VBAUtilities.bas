Attribute VB_Name = "B03_VBAUtilities"
Sub CreateEventsForPreTreatment(WorksheetToInject As Worksheet)
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

Sub CreateEventsForPharmacodeCompletion(WorksheetToInject As Worksheet)
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
            .InsertLines LineNum, vbTab & "If LastEditedCell Is Nothing Then LastEditedCell = ActiveCell"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "Call ColorLabelling(LastEditedCell)"
            
            'OnSelectionChange
            LineNum = .CreateEventProc("SelectionChange", "Worksheet")
            LineNum = LineNum + 1
             .InsertLines LineNum, vbTab & "Set LastEditedCell = ActiveCell"
  
            'Deactivate
            LineNum = .CreateEventProc("Deactivate", "Worksheet")
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "Set LastEditedCell = Nothing"
        End With
End Sub



Sub RemoveEventsProcedure(ByVal WorksheetToClean As Worksheet)

    Dim strObjectName As String
    strObjectName = WorksheetToClean.CodeName
    
    ' Remove all lines from module...
    With ThisWorkbook.VBProject.VBComponents(strObjectName).CodeModule
        .DeleteLines 1, .CountOfLines
    End With

End Sub

Sub SetWsName(ByVal ws As Worksheet, NewName As String)

    ws.Name = ThisWorkbook.VBProject.VBComponents(ws.Name).Properties("Codename")
    ws.Parent.VBProject.VBComponents(ws.CodeName).Name = NewName
    
    
End Sub
