Attribute VB_Name = "Z2_Parameters_Ribbon"
Public Function Function_Clicked(control As IRibbonControl, ByRef pressed)
    pressed = GetKey(control.ID)
End Function

Public Function Function_Action(control As IRibbonControl, pressed As Boolean)
    Store control.ID, pressed
    'MsgBox control.ID & " " & pressed
End Function

Public Sub Store(control_id As String, value As Boolean)
    DefGlobal
    PARAM_TABLE.Columns(1).Find(control_id).Offset(0, 1).value = value
    
    Select Case control_id
        Case Is = "VerifyNbSheets"
        Case Is = "VerifyColumnsTitle"
        Case Is = "VerifyColumnsContent"
        'Case Is = "MergeFiles"
        Case Is = "DispatchFiles"
        'Case Is = "CheckPharmacodes"
        'Case Is = ""
        'Case Is = "AuthorizeChangesOnOpening"
        'Case Is = "SaveReadOnly"
        'Case Is = "SaveinSeparateSheets"
        'Case Is = ""
        'Case Is = ""
        Case Else
            MsgBox "Feature not implemented yet"
    End Select

End Sub

Public Function GetKey(control_id As String) As Boolean
    '''write the code for getting the key back from the source which you might have used to store the value.
    '''return the correct value here
    GetKey = True ' or whatever you have selected previously
End Function




