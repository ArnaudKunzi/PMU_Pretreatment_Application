Attribute VB_Name = "Z2_Parameters_Ribbon"
Dim Rib As IRibbonUI
Public MyTag As String


Public Function Function_Clicked(control As IRibbonControl, ByRef pressed)
    pressed = GetKey(control.ID)
End Function

Public Function Function_Action(control As IRibbonControl, pressed As Boolean)
    Store control.ID, pressed
    'MsgBox control.ID & " " & pressed
End Function

'Callback for Instructions getLabel
Sub GetInstructionLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Instructions: lol"
End Sub

Public Sub Store(control_id As String, value As Boolean)
    DefGlobal
    PARAM_TABLE.Columns(1).Find(control_id).Offset(0, 1).value = value
    
    Select Case control_id
        Case Is = "VerifyNbSheets"
        Case Is = "VerifyColumnsTitle"
        Case Is = "VerifyColumnsContent"
        'Case Is = "MergeFiles"
        Case Is = "DispatchFiles"
        Case Is = "CheckPharmacodes"
        'Case Is = ""
        'Case Is = "AuthorizeChangesOnOpening"
        'Case Is = "SaveReadOnly"
        'Case Is = "SaveinSeparateSheets"
        Case Is = "ShowEveryTabs"
            If value Then
                Call ShowAllTabs
            Else
                Call ShowOnlyCustomTabs
            End If
        Case Else
            MsgBox "Feature not implemented yet"
    End Select

End Sub

Public Function GetKey(control_id As String) As Boolean
    '''write the code for getting the key back from the source which you might have used to store the value.
    '''return the correct value here
    'Select Case control_id
    'Case Is = "DispatchFiles"
    '    'PARAM_TABLE.Columns(1).Find("CheckPharmacodes").Offset(0, 1).value
    'Case Else
    'End Select
    
    GetKey = PARAM_TABLE.Columns(1).Find(control_id).Offset(0, 1).value ' True ' or whatever you have selected previously
End Function


'CALLBACKS ON VISIBILITY 1

'Callback for customUI.onLoad

Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set Rib = ribbon
End Sub

Sub GetVisible(control As IRibbonControl, ByRef visible)
    If MyTag = "show" Then
        visible = True
    Else
        If control.Tag Like MyTag Then
            visible = True
        Else
            visible = False
        End If
    End If
End Sub

Sub RefreshRibbon(Tag As String)
    MyTag = Tag
    If Rib Is Nothing Then
        MsgBox "Error, Save/Restart your workbook"
    Else
        Rib.Invalidate
    End If
End Sub




' Macros ON VISIBILITY

Sub ShowAllTabs()
'Show every Tab, Group or Control(we use the wildgard "*")
'You can also use "rib*" because all tags start with rib in this file
    Call RefreshRibbon(Tag:="*")
End Sub

Sub ShowOnlyCustomTabs()
'Show every Tab, Group or Control(we use the wildgard "*")
'You can also use "rib*" because all tags start with rib in this file
    Call RefreshRibbon(Tag:="Custom*")
End Sub


'Callback for customUI.onLoad
'Sub ribbonLoaded(ribbon As IRibbonUI)
'End Sub

'Callback for MyCustomHomeTab getVisible
'Sub GetVisible(control As IRibbonControl, ByRef returnedVal)
'End Sub

'Callback for VerifyNbSheets getPressed
'Sub Function_Clicked(control As IRibbonControl, ByRef returnedVal)
'End Sub

'Callback for VerifyNbSheets onAction
'Sub Function_Action(control As IRibbonControl, pressed As Boolean)
'End Sub



