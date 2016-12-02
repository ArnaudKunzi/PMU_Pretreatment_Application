Attribute VB_Name = "B31_Export"
Sub ExportSeparately(control As IRibbonControl)
    Call Export(True)
End Sub

Sub ExportTogether(control As IRibbonControl)
    Call Export(False)
End Sub

Sub Export(separately As Boolean)
    Call DefGlobal
    
    Dim PrevSheet As Worksheet
    Dim ExportWorkbook As Workbook
    
    Dim SaveinSeparateSheets As Boolean
    Dim TrackChanges As Boolean
    Dim SheetsToExport As String
    
    SaveInSameWB = PARAM_TABLE.Columns(1).Find("SaveInSameWB").Offset(0, 1).value
    
    If Not separately Then Call MergeSheets
    
    'List of sheets to export (those with tab colored EXPORTCOLOR)
    For Each sheet In Worksheets
        If sheet.Tab.ColorIndex = EXPORTCOLOR Then
            If SheetsToExport = "" Then
                SheetsToExport = sheet.Name
            Else
                SheetsToExport = Join(Array(SheetsToExport, sheet.Name), "|")
            End If
        End If
    Next sheet
    
    
    'Ask for save path
    varResult = Application.GetSaveAsFilename(INTERNALS.ListObjects("SavePath").ListColumns(1).DataBodyRange.value)
    'checks to make sure the user hasn't canceled the dialog
    If varResult <> False Then
        INTERNALS.ListObjects("SavePath").ListColumns(1).DataBodyRange.value = varResult
    End If
    
    
    
    'Move the sheets
    If SaveInSameWB Then
        For Each sheet In Split(SheetsToExport, "|")
            Call RemoveEventsProcedure(Worksheets(sheet))
            If ExportWorkbook Is Nothing Then
                ThisWorkbook.Worksheets(sheet).Move
                Set ExportWorkbook = ActiveWorkbook
                Set PrevSheet = ActiveWorksheet
            Else
                ThisWorkbook.Worksheets(sheet).Move After:=PrevSheet
            End If
        Next
        ExportWorkbook.SaveAs "Données Médicament Pré-Traitées " & "" & Year
        ExportWorkbook.Close False
    Else
        For Each sheet In Split(SheetsToExport, "|")
            Call RemoveEventsProcedure(Worksheets(sheet))
            ThisWorkbook.Worksheets(sheet).Move
            
            ActiveWorkbook.SaveAs ActiveWorkbook.Worksheets(1).Name & "_" & Year
            ActiveWorkbook.Close False
        Next
    End If
    
    
    Call UpdateStage("PreTreatment")
    
End Sub



Sub test()
    Dim wb As Workbook
    Set wb = Workbooks.Add
End Sub

