VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progression 
   Caption         =   "Progression"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5010
   OleObjectBlob   =   "Progression.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Progression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    ' Set the width of the progress bar to 0.
    Progression.LabelProgress.Width = 0
    ' Call the main subroutine.
    'Call test
    'CallByName Progression, UFORM_PAR_SUBNAME, VbMethod
    
    'Application.Run UFORM_PAR_SUBNAME
    
End Sub


Sub UpdateProgressBar(PctDone As Single)

    With Progression
        ' Update the Caption property of the Frame control.
        .FrameProgress.Caption = Format(PctDone, "0%")

        ' Widen the Label control.
        .LabelProgress.Width = PctDone * _
            (.FrameProgress.Width - 10)
    End With

    ' The DoEvents allows the UserForm to update.
    DoEvents
    
End Sub

