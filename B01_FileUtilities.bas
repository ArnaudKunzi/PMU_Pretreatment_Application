Attribute VB_Name = "B01_FileUtilities"
Public Function SelectFile(Many As Boolean)
    Dim FileDialog As Office.FileDialog
    Set FileDialog = Application.FileDialog(msoFileDialogOpen)
    Dim FILE As String
    Dim DirectoryPath As String
    
    DirectoryPath = INTERNALS.ListObjects("path").ListColumns(1).DataBodyRange(1).value
    DirectoryPath = Replace(DirectoryPath, "$", INTERNALS.ListObjects("cantons").ListColumns(1).DataBodyRange.Find(Canton.value).Offset(0, 1).value)
    DirectoryPath = Replace(DirectoryPath, "%", Year.value)
    If Dir(DirectoryPath, vbDirectory) = "" Then DirectoryPath = Replace(DirectoryPath, "MEDICAMENTS_" & Year.value & "\", "")
    
    With FileDialog
        .Title = "Select file"
        '.InitialFileName = "L:\PMU\COMMUN_PHARMACIE\RECHERCHE\01 Travaux de recherche\ANI_EMS\EMS " & Canton.value & "\03 Donnees\033 Donnees brutes\" & Year.value & "\"
        .InitialFileName = DirectoryPath
        
        
        
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

Sub testopenexp()
    OpenExplorerWithFileSelected ("L:\PMU\COMMUN_PHARMACIE\RECHERCHE\01 Travaux de recherche\ANI_EMS\EMS VD\03 Donnees\034 Donnees traitees et analyse\[Template] Données_EMS-VD_YYYY.xlsx")
End Sub


Sub OpenExplorerWithFileSelected(filepath$)

'Test for file exists
retval = Dir(filepath$)
If retval <> "" Then
    'If Exists, then open Windows Explorer and select
    shellparm = "/select," & filepath$
    Shell "explorer """"" & shellparm & """""", vbNormalFocus
End If


End Sub



