Attribute VB_Name = "A0_GlobalVariables"
Global Year As Range
Global Canton As Range
Global Status As Range
Global PARAM_TABLE As Range
Global STAGE As Range
Global InPh_colname As String
Global DataSheetName  As String
Global LogSheetName As String

Global LastValueSelected As Variant
Global LastCommentsSelected As Variant
Global LastEditedCell As Range
Global EDITCOLOR As Integer
Global EXPORTCOLOR As Integer

Sub DefGlobal()
    Set Year = A_0.Range("E7")
    Set Canton = A_0.Range("E9")
    Set Status = INTERNALS.ListObjects("status").ListColumns("style").DataBodyRange
    Set PARAM_TABLE = INTERNALS.ListObjects("Parameters").DataBodyRange
    Set STAGE = INTERNALS.ListObjects("stage").ListColumns(1).DataBodyRange
    InPh_colname = "InvalidPharmacodes"
    DataSheetName = "DATA"
    LogSheetName = "LOG_" & Year
    
    EDITCOLOR = 8
    EXPORTCOLOR = 23
End Sub



