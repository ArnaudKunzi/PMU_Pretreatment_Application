Attribute VB_Name = "B03_VBAUtilities"
Private Declare Function apiPostMessage _
    Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) _
    As Long

Private Declare Function apiFindWindow _
    Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) _
    As Long

Private Declare Function apiIsWindow _
    Lib "user32" Alias "IsWindow" _
    (ByVal hwnd As Long) _
    As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub CreateEventsForPreTreatment(WorksheetToInject As Worksheet)
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim LineNum As Long
        Const DQUOTE = """" ' one " character

        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents(WorksheetToInject.CodeName)
        Set CodeMod = VBComp.CodeModule
        
        Application.ScreenUpdating = False
        
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
        
        Application.SendKeys ("%q")
        Application.ScreenUpdating = True
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
        
        Application.ScreenUpdating = False
        
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
        
        Application.SendKeys ("%q")
        Application.ScreenUpdating = True
        
End Sub

Sub CreateEventsForReport(WorksheetToInject As Worksheet)
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Dim LineNum As Long
        Const DQUOTE = """" ' one " character

        Set VBProj = ActiveWorkbook.VBProject
        Set VBComp = VBProj.VBComponents(WorksheetToInject.CodeName)
        Set CodeMod = VBComp.CodeModule
        
        Application.ScreenUpdating = False
        
        With CodeMod

            'FollowHyperlink
            LineNum = .CreateEventProc("FollowHyperlink", "Worksheet")
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "Dim WhatToColor As Range"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "Dim Typage As String"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "Dim ChampsInconnus As String"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "Typage = Target.Parent.Offset(0, Me.Range(" & DQUOTE & "1:1" & DQUOTE & ").Find(" & DQUOTE & "typage" & DQUOTE & ").column - Target.Parent.column).value"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "ChampsInconnus = Target.Parent.Offset(0, Me.Range(" & DQUOTE & "1:1" & DQUOTE & ").Find(" & DQUOTE & "Champs inconnus" & DQUOTE & ").column - Target.Parent.column).value"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "With ActiveWorkbook.ActiveSheet"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & "If Typage <> " & DQUOTE & DQUOTE & " Then"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & "For Each value In Split(Typage, Chr(10))"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & "value = Replace(value, " & DQUOTE & "Col. " & DQUOTE & ", " & DQUOTE & DQUOTE & ")"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & "col = IncCol(" & DQUOTE & "A" & DQUOTE & ", .Range(" & DQUOTE & "1:1" & DQUOTE & ").Find(Split(value, " & DQUOTE & ":" & DQUOTE & ")(0)).column - 1)"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & "rows = Replace(Replace(Split(value, " & DQUOTE & ":" & DQUOTE & ")(1), " & DQUOTE & "l." & DQUOTE & ", " & DQUOTE & DQUOTE & "), " & DQUOTE & " " & DQUOTE & ", " & DQUOTE & DQUOTE & ")"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & "For Each cellgroup In Split(rows, " & DQUOTE & ", " & DQUOTE & ")"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & vbTab & "Dim row As Variant"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & vbTab & "row = Split(cellgroup, " & DQUOTE & "-" & DQUOTE & ")"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & vbTab & "If WhatToColor Is Nothing Then"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Set WhatToColor = .Range(col & row(0) & " & DQUOTE & ":" & DQUOTE & " & col & row(UBound(row)))"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & vbTab & "Else"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Set WhatToColor = Union(WhatToColor, .Range(col & row(0) & " & DQUOTE & ":" & DQUOTE & " & col & row(UBound(row))))"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & vbTab & "End If"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & "Next cellgroup"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & "Next value"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & "WhatToColor.Cells.Interior.ColorIndex = 3"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & "WhatToColor.Select"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & "End If"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & "If ChampsInconnus <> " & DQUOTE & DQUOTE & " Then"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & "For Each value In Split(ChampsInconnus, " & DQUOTE & "," & DQUOTE & ")"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & "If WhatToColor Is Nothing Then"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & vbTab & "Set WhatToColor = .Range(" & DQUOTE & "1:1" & DQUOTE & ").Find(value)"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & "Else"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & vbTab & "Set WhatToColor = Union(WhatToColor, .Range(" & DQUOTE & "1:1" & DQUOTE & ").Find(value))"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & vbTab & "End If"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & "Next value"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & "WhatToColor.Cells.Interior.ColorIndex = 3"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & vbTab & "WhatToColor.Select"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & "End If"
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "End With"
            
            'SelectionChange
            LineNum = .CreateEventProc("SelectionChange", "Worksheet")
            LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "If Target.Column = Me.Range(" & DQUOTE & "1:1" & DQUOTE & ").Find(" & DQUOTE & "Champs inconnus" & DQUOTE & ").column And Len(Target.Value)>0 Then"
             LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & "INTERNALS.visible=xlSheetVisible"
             LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & vbTab & "INTERNALS.Activate"
             LineNum = LineNum + 1
             .InsertLines LineNum, vbTab & vbTab & "INTERNALS.ListObjects(" & DQUOTE & "attributes" & DQUOTE & ").ListColumns(1).DataBodyRange(INTERNALS.ListObjects(" & DQUOTE & "attributes" & DQUOTE & ").ListColumns(1).DataBodyRange.Count+1).Select"
             LineNum = LineNum + 1
            .InsertLines LineNum, vbTab & "End If"

        End With
        
        Application.SendKeys ("%q")
        Application.ScreenUpdating = True
End Sub

Sub RemoveEventsProcedure(ByVal WorksheetToClean As Worksheet)

    Application.ScreenUpdating = False
    
    Dim strObjectName As String
    strObjectName = WorksheetToClean.CodeName
    
    ' Remove all lines from module...
    With ThisWorkbook.VBProject.VBComponents(strObjectName).CodeModule
        .DeleteLines 1, .CountOfLines
    End With
    Application.SendKeys ("%q")
    Application.ScreenUpdating = True
    
End Sub

Sub SetWsName(ByVal ws As Worksheet, NewName As String)
    
    'ws.Parent.VBProject.VBComponents(ws.CodeName).Name = NewName
      
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function fCloseVBEWindow() As Boolean

    Const VBE_CLASS = "wndclass_desked_gsk"
    Dim hwnd As Long
    hwnd = apiFindWindow(VBE_CLASS, Application.VBE.MainWindow.Caption)
    If hwnd Then
        Call apiPostMessage(hwnd, WM_CLOSE, 0, 0&)
        fCloseVBEWindow = (apiIsWindow(hwnd) <> 0)
    End If
    Application.SendKeys "~"
End Function
