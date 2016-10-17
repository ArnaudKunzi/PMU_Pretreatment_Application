VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ProtectionOfSheet()

    Dim UnprotectedRange As Range
    
    With ThisWorksheet
        Unprotect Password:="IKnowWhatIAmDoing"
        Set UnprotectedRange = Union(Range("E7"), Range("E9"))
        UnprotectedRange.Locked = False
        'Protect Password:="IKnowWhatIAmDoing"
    End With

End Sub

Sub SetValidation()
    Dim ListOfCantons As String
    Dim list As Variant
    list = INTERNALS.ListObjects("cantons").ListColumns("canton_code").DataBodyRange.value
    
    ListOfCantons = ""
    For Each entry In list
        ListOfCantons = ListOfCantons & "," & entry
    Next
    ListOfCantons = Right(ListOfCantons, Len(ListOfCantons) - 1)
    
    With Canton.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
         xlBetween, Formula1:=ListOfCantons  '"Vaud,Fribourg"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .InputMessage = ""
        .ShowInput = True
    End With
End Sub
