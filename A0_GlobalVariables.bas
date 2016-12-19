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

Global DATA_SH As GlbSheet
Global PHARMA_SH As GlbSheet
Global PHAUNI_SH As GlbSheet
Global REPORT_SH As GlbSheet

Public Type GlbSheet
    Name As String
    VOffset As Integer
    HOffset As Integer
End Type


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
    
    REPORT_SH.Name = "RAPPORT"
    REPORT_SH.VOffset = 1
    REPORT_SH.HOffset = 3
    
    DATA_SH.Name = "DATA"
    DATA_SH.VOffset = 1
    DATA_SH.HOffset = 3
    
    PHARMA_SH.Name = "invalid_pharmacodes"
    PHARMA_SH.VOffset = 1
    PHARMA_SH.HOffset = 3
    
    PHAUNI_SH.Name = "Pharmacodes à compléter"
    PHAUNI_SH.VOffset = 1
    PHAUNI_SH.HOffset = 5
    
End Sub



