Attribute VB_Name = "A0_GlobalVariables"
Global Year As Range
Global Canton As Range
Global Status As Range
Global PARAM_TABLE As Range

Sub DefGlobal()
    Set Year = A_0.Range("E7")
    Set Canton = A_0.Range("E9")
    Set Status = INTERNALS.ListObjects("status").ListColumns("style").DataBodyRange
    Set PARAM_TABLE = INTERNALS.ListObjects("Parameters").DataBodyRange
End Sub
