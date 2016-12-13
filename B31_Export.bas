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
    
    Dim sPath As String
    Dim SaveinSeparateSheets As Boolean
    Dim SeparatedSheets As Boolean
    Dim TrackChanges As Boolean
    Dim SheetsToExport As String
    Dim FileName As String
    
    SaveInSameWB = PARAM_TABLE.Columns(1).Find("SaveInSameWB").Offset(0, 1).value
    SeparatedSheets = PARAM_TABLE.Columns(1).Find("TbtnToggleSeparateByPhStatus").Offset(0, 1).value
    
    If (Not separately) And SeparatedSheets Then Call MergeSheets
    If separately And (Not SeparatedSheets) Then Call SplitSheets
    
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
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = INTERNALS.ListObjects("SavePath").ListColumns(1).DataBodyRange.value
        .InitialView = msoFileDialogViewDetails
        
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count > 0 Then
             sPath = .SelectedItems(1)
        End If
    End With
    'varResult = Application.GetSaveAsFilename(INTERNALS.ListObjects("SavePath").ListColumns(1).DataBodyRange.value)
    'checks to make sure the user hasn't canceled the dialog
    'If varResult <> False Then
    '    INTERNALS.ListObjects("SavePath").ListColumns(1).DataBodyRange.value = varResult
    'Else
    '    varResult = INTERNALS.ListObjects("SavePath").ListColumns(1).DataBodyRange.value
    'End If
    
    
    
    'Move the sheets
    If SaveInSameWB Then
    FileName = "Données Médicament Pré-Traitées " & "" & Year
        For Each sheet In Split(SheetsToExport, "|")
            ThisWorkbook.Activate
            Call RemoveEventsProcedure(Worksheets(sheet))
            If ExportWorkbook Is Nothing Then
                Debug.Print "1 " & sheet
                ThisWorkbook.Worksheets(sheet).Move
                Set ExportWorkbook = ActiveWorkbook
                
            Else
                Debug.Print "2 " & sheet
                ThisWorkbook.Worksheets(sheet).Move After:=PrevSheet
                
            End If
            Set PrevSheet = ActiveSheet
        Next
        ExportWorkbook.SaveAs sPath & Application.PathSeparator & FileName & ".xlsx"
        ExportWorkbook.Close False
    Else
        For Each sheet In Split(SheetsToExport, "|")
            
            Call RemoveEventsProcedure(Worksheets(sheet))
            ThisWorkbook.Worksheets(sheet).Move
            FileName = "Données Médicament Pré-Traitées " & Year & " " & ActiveWorkbook.Worksheets(1).Name
            Set ExportWorkbook = ActiveWorkbook
            ExportWorkbook.SaveAs sPath & Application.PathSeparator & FileName & ".xlsx"
            ExportWorkbook.Close False
        Next
    End If
    
''''''''''''
                'ExportWorkbook.SaveAs FileName:=sPath & " ", _
                'FileFormat:=xlNormal, _
                'Password:="", _
                'WriteResPassword:="", _
                'ReadOnlyRecommended:=False, _
                'CreateBackup:=False
'''''''''''''
    
    OpenExplorerWithFileSelected (sPath & Application.PathSeparator & FileName & ".xlsx")
    
    Call UpdateStage("PreTreatment")
    
End Sub
