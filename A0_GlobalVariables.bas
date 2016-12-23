Attribute VB_Name = "A0_GlobalVariables"
Global DisplayTag As Range

Global Year As Range
Global Canton As Range
Global Status As Range
Global PARAM_TABLE As Range
Global STAGE As Range

Global LastValueSelected As Variant
Global LastCommentsSelected As Variant
Global LastEditedCell As Range
Global EDITCOLOR As Integer
Global EXPORTCOLOR As Integer

Global DATA_SH As GlbSheet
Global PHARMA_SH As GlbSheet
Global PHAUNI_SH As GlbSheet
Global REPORT_SH As GlbSheet
Global LOG_SH As GlbSheet

Public Type GlbSheet
    Name As String
    VOffset As Integer
    HOffset As Integer
End Type


Sub DefGlobal()
    Set Year = INTERNALS.ListObjects("Variables_danalyse").ListColumns(2).DataBodyRange(2)
    Set Canton = INTERNALS.ListObjects("Variables_danalyse").ListColumns(2).DataBodyRange(1)
    Set Status = INTERNALS.ListObjects("status").ListColumns("style").DataBodyRange
    Set PARAM_TABLE = INTERNALS.ListObjects("Parameters").DataBodyRange
    Set STAGE = INTERNALS.ListObjects("stage").ListColumns(1).DataBodyRange
    Set DisplayTag = INTERNALS.ListObjects("DisplayTag").ListColumns(1).DataBodyRange
    
    EDITCOLOR = 8
    EXPORTCOLOR = 23
    
    REPORT_SH.Name = "RAPPORT"
    REPORT_SH.VOffset = 1
    REPORT_SH.HOffset = 0
    
    DATA_SH.Name = "DATA"
    DATA_SH.VOffset = 1
    DATA_SH.HOffset = 3
    
    PHARMA_SH.Name = "invalid_pharmacodes"
    PHARMA_SH.VOffset = 1
    PHARMA_SH.HOffset = 3
    
    PHAUNI_SH.Name = "Pharmacodes_a_completer"
    PHAUNI_SH.VOffset = 1
    PHAUNI_SH.HOffset = 5
    
    LOG_SH.Name = "LOG_" & Year
    
    'Reload ribbon
    'If Rib Is Nothing Then
    '    Dim ribbonPointer As Long
    '    ribbonPointer = INTERNALS.ListObjects("IRibbonUI").DataBodyRange.value
    '    Call CopyMemory(Rib, ribbonPointer, 4)
    'End If
    
End Sub



