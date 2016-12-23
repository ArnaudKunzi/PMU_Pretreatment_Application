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
