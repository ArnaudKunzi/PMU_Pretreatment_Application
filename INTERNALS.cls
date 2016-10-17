VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INTERNALS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

''A change in the range ListRange causes the ribbon to be invalidated
'Private Sub Worksheet_Change(ByVal Target As Range)
    'If Not Intersect(Target, INTERNALS.ListObjects("have_several_tabs").ListColumns("have_several_tabs").DataBodyRange) Is Nothing Then
    '    RedoRib
    'End If
'End Sub
