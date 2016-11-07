Attribute VB_Name = "B0_FileUtilities"
Public Function SelectFile(Many As Boolean)
    Dim FileDialog As Office.FileDialog
    Set FileDialog = Application.FileDialog(msoFileDialogOpen)
    Dim FILE As String
    
    With FileDialog
        .Title = "Select file"
        .InitialFileName = "L:\PMU\COMMUN_PHARMACIE\RECHERCHE\01 Travaux de recherche\ANI_EMS\EMS " & Canton.value & "\03 Donnees\033 Donnees brutes\" & Year.value & "\"
        If Many Then
            .AllowMultiSelect = True
        Else
            .AllowMultiSelect = False
        End If
        .Filters.Clear
        .Filters.Add "Tous les fichiers", "*.*"
        .Filters.Add "Document Excel", "*.xls; *.xlsx; *.xlsb; *.csv"
        .FilterIndex = 2
        If .Show Then
             If Many Then
                FILE = ""
                For Each F In .SelectedItems
                    FILE = FILE & "|" & F
                Next
                FILE = Right(FILE, Len(FILE) - 1)
            Else
                FILE = .SelectedItems(1)
            End If
        End If
    End With
    Set FileDialog = Nothing
    SelectFile = FILE
End Function




'Sub LoopThroughFiles(directorypath As String)
'    Dim MyObj As Object, MySource As Object, FILE As Variant
'   FILE = Dir(directorypath)
'   While (FILE <> "")
'      If ConformableFileName(FILE) Then
'         'MsgBox "found " & file
'         'Exit Sub
'      End If
'     FILE = Dir
'  Wend
'End Sub
