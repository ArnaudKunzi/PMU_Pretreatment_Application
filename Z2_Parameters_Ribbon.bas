Attribute VB_Name = "Z2_Parameters_Ribbon"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)
Public Rib As IRibbonUI
Public MyTag As String


Public Sub Function_Clicked(control As IRibbonControl, ByRef pressed)
    
    pressed = GetKey(control.ID)
    'MsgBox control.ID & " " & pressed
End Sub


'Callback for TbtnToggleSeparateByPhStatus getPressed
'Sub Function_Clicked(control As IRibbonControl, ByRef returnedVal)

'End Sub



Public Function Function_Action(control As IRibbonControl, pressed As Boolean)
    Store control.ID, pressed
    
    'MsgBox control.ID & " " & pressed
    
    '    Select Case control.ID
    'Case Is = "TbtnToggleSeparateByPhStatus"
    '
    'Case Else
    'End Select
    
End Function

'Callback for Instructions getLabel
'Sub GetInstructionLabel(control As IRibbonControl, ByRef returnedVal)
    'NOT WORKING SINCE NO MULTILINE COMMENT POSSIBLE
    'returnedVal = INTERNALS.ListObjects("Instructions").ListColumns(1).DataBodyRange.Find(STAGE.value).Offset(0, 1).value '"Instructions: lol"
    'returnedVal = "uiopg" & vbCrLf & "srgsdths"
    '"Renseignez le canton et l'année d'analyse des données puis cliquez sur « Charger les fichiers » pour selectionner les fichiers médicaments." & Chr(10) & _
    '             "Le programme se chargera de produire un rapport sur la conformité des données." & Chr(10) & _
    '             "Il est possible de choisir quels sont les critères de conformité dans le ruban « Paramètres»"
    'If PARAM_TABLE.Columns(1).Find("ShowEveryTabs").Offset(0, 1).value Then
    '    Call RefreshRibbon(Tag:="*")
    'Else
    '    Call RefreshRibbon(Tag:="Custom*")
    'End If
'End Sub

Public Sub Store(control_id As String, value As Boolean)
    Call DefGlobal
    PARAM_TABLE.Columns(1).Find(control_id).Offset(0, 1).value = value
    
    Select Case control_id
        Case Is = "VerifyNbSheets"
        Case Is = "VerifyColumnsTitle"
        Case Is = "VerifyColumnsContent"
        'Case Is = "MergeFiles"
        Case Is = "AllowAllButtons"
            Debug.Print "AllowAllButtons"
            If value Then
                Call UpdateStage(-1)
            Else
                Call UpdateStage(STAGE.value)
            End If
        Case Is = "CheckPharmacodes"
        'Case Is = "TrackChanges"
        'Case Is = "AuthorizeChangesOnOpening"
        'Case Is = "SaveReadOnly"
        Case Is = "SaveInSameWB"
        Case Is = "TbtnToggleSeparateByPhStatus"
            'Call TbtnToggleSeparateByPhStatus(value)
        Case Is = "ShowEveryTabs"
            If value Then
                Call ShowAllTabs
                'Call UpdateStage(STAGE.value)
            Else
                Call ShowOnlyCustomTabs
                'Call UpdateStage(STAGE.value)
            End If
        Case Else
            MsgBox "Feature not implemented yet"
    End Select

End Sub

Public Function GetKey(control_id As String) As Boolean
    Call DefGlobal
    '''write the code for getting the key back from the source which you might have used to store the value.
    '''return the correct value here
    
    GetKey = PARAM_TABLE.Columns(1).Find(control_id).Offset(0, 1).value ' True ' or whatever you have selected previously
End Function


'CALLBACKS ON VISIBILITY 1

'Callback for customUI.onLoad
Sub RibbonOnLoad(ribbon As IRibbonUI)
    Call DefGlobal
    Set Rib = ribbon
    INTERNALS.ListObjects("IRibbonUI").DataBodyRange.value = ObjPtr(ribbon)
    
    Call UpdateStage(1)
End Sub

Sub GetVisible(control As IRibbonControl, ByRef visible)
    If control.Tag Like MyTag Then

            visible = True
     
    'ElseIf Len(control.Tag) > 0 Then
        
    '    If MyTag Like "*VG*" Or Len(MyTag) = 0 Then
    '        If PARAM_TABLE.Columns(1).Find("ShowEveryTabs").Offset(0, 1).value Then
    '            visible = True
    '        Else
    '            visible = False
    '        End If
    '    Else
    '        visible = False
    '    End If
    Else
        visible = False
    End If
    
End Sub

Sub GetEnabledMacro(control As IRibbonControl, ByRef Enabled)
    Call DefGlobal

    'If MyTag = "Enable" Then
    '    Enabled = True
    'Else
        'If control.Tag Like MyTag And control.Tag Like DisplayTag.value Then
        If control.Tag Like DisplayTag.value Then
            Enabled = True
        Else
            Enabled = False
        End If
    'End If
End Sub

Sub RefreshRibbon(Tag As String)
    MyTag = Tag
    If Rib Is Nothing Then
        Dim ribbonPointer As Long
        ribbonPointer = INTERNALS.ListObjects("IRibbonUI").DataBodyRange.value
        Call CopyMemory(Rib, ribbonPointer, 4)
    End If
    
    Rib.Invalidate

End Sub


'Sub RefreshButton(controlID As String)
'    Rib.InvalidateControl controlID
'End Sub



' Macros ON VISIBILITY

Sub ShowAllTabs()
'Show every Tab, Group or Control(wildgard "*")
    Call RefreshRibbon(Tag:="*")
End Sub

Sub ShowOnlyCustomTabs()
'Show  Tab, Group or Control(wildgard "*")
    Call RefreshRibbon(Tag:="*C_*")
End Sub


Sub TbtnToggleSeparateByPhStatus(control As IRibbonControl, pressed As Boolean)
    Call DefGlobal
    
    Store control.ID, pressed
    
    If Not PARAM_TABLE.Columns(1).Find("TbtnToggleSeparateByPhStatus").Offset(0, 1).value Then
        Call MergeSheets
    Else
        Call SplitSheets
    End If
End Sub




'Sub AllowEdit(control As IRibbonControl, ByRef CancelDefault)
'
'  MsgBox "Yes?!", vbOKOnly, "Command Repurposing Demo"
'
'  CancelDefault = False
'End Sub


'*************************************************************
'Sub rxMenu_onAction(control As IRibbonControl)
'  msSplitStyle = Mid$(control.ID, 7)
'  getRibbon().InvalidateControl "rxButton"
'  MsgBox "Control invalidated"
'End Sub


