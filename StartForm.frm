VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartForm 
   Caption         =   "Plateforme de prétraitement des données médicaments"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5790
   OleObjectBlob   =   "StartForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonCANCEL_Click()
    Unload Me
End Sub

Private Sub CommandButtonOK_Click()
    
    If Not IsNumeric(Me.YearTextBox.value) And Me.YearTextBox.value > 2000 And Me.YearTextBox.value < Format(Now, "yyyy") Then
        MsgBox "La saisie du champ 'Année d'analyse' est invalide.", vbCritical
        Exit Sub
    End If
    
    Me.Hide
    
    Canton.value = Me.CantonCombobox.value
    Year.value = Me.YearTextBox.value
    
    
    Call UpdateStage(2)
End Sub

Private Sub UserForm_Initialize()
    YearTextBox.value = Format(Now, "yyyy") - 1
    
    CantonCombobox.Clear
    
    With CantonCombobox
        .AddItem "VAUD"
        .AddItem "FRIBOURG"
        .ListIndex = 0
        
    End With
    
    YearTextBox.SetFocus

End Sub
