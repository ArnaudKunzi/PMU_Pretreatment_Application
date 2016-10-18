Attribute VB_Name = "A1_Main"
Sub LoadFiles(control As IRibbonControl)
    Dim FilesListString As String
    Dim FileList As Variant
    Call DefGlobal
    
    If (Len(Year.value) <> 0 And Len(Canton) <> 0) Then
        FilesListString = SelectFile(True)
    Else
        MsgBox "année d'analyse et/ou canton à analyser non renseigné.", vbCritical
        Exit Sub
    End If
    
    'Overwrite the statistics sheet
    If Len(FilesListString) <> 0 Then
        Call PrepareOverviewSheet(FilesListString)
        'FilesList = Split(FilesListString, "|")

    End If
    
End Sub

Sub Refresh(control As IRibbonControl)
    Dim FilesListString As String
    Dim FilesList As Variant
    Dim table As ListObject
    Dim Path As Range
    
    Call DefGlobal
    
    Set table = INTERNALS.ListObjects("file_to_load")
    Set Path = INTERNALS.ListObjects("path").ListColumns("path").DataBodyRange
    
    FilesListString = Path(1).value & table.ListColumns(2).DataBodyRange(1).value
    For i = 2 To table.ListRows.Count
        FilesListString = FilesListString & "|" & Path(1).value & table.ListColumns(2).DataBodyRange(i).value
    Next i
    
    If Len(FilesListString) <> 0 Then
        Call PrepareOverviewSheet(FilesListString)
        'FilesList = Split(FilesListString, "|")
        
    End If
End Sub

Sub StartPreTreatment(control As IRibbonControl)
    Call DefGlobal
    Call TransferColumns
End Sub



Sub PrepareOverviewSheet(FilesListSring As String)
    
    Dim FilesList As Variant
        FilesList = Split(FilesListSring, "|")
    Dim counter As Integer
    Dim nb_sheets As Variant
    
    Dim HOffset As Integer
        HOffset = 0
    Dim VOffset As Integer
        VOffset = 1
        
    Dim CondStatusOK As Boolean
        
    Call SaveFilesList(FilesList)
        
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("RAPPORT").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "RAPPORT"
    
    nb_sheets = HowManySheets(FilesList)
    
    Call MainLoadingLoop(FilesList, nb_sheets)
    
    Application.ScreenUpdating = False
    
    With Worksheets("RAPPORT")
        .Cells.Font.Size = "8"
        .Columns("A:B").Group
        .Outline.ShowLevels ColumnLevels:=1
        .Range(Chr(Asc("A") + HOffset) & VOffset).value = "n°"
        .Range(Chr(Asc("B") + HOffset) & VOffset).value = "Chemin"
        .Range(Chr(Asc("C") + HOffset) & VOffset).value = "Nom"
        .Range(Chr(Asc("D") + HOffset) & VOffset).value = "Status"
        .Range(Chr(Asc("E") + HOffset) & VOffset).value = "n° EMS"
        .Range(Chr(Asc("F") + HOffset) & VOffset).value = "EMS conforme"
        .Range(Chr(Asc("G") + HOffset) & VOffset).value = "Nombre d'onglets"
            .Range(Chr(Asc("G") + HOffset) & VOffset).AddComment ("Seules les données présentes dans le" & Chr(10) & "premier onglet sont pris en compte." & Chr(10) _
                                                                     & "assurez-vous que toutes les données" & Chr(10) & "pertinentes soient dans une" & Chr(10) _
                                                                     & "table dans le premier onglet.")
        '.Range(Chr(Asc("H") + HOffset) & VOffset).value = "Système informatique"
        .Range(Chr(Asc("H") + HOffset) & VOffset).value = "typage"
            .Range(Chr(Asc("H") + HOffset) & VOffset).AddComment ("Non implémenté pour le moment")
        .Range(Chr(Asc("I") + HOffset) & VOffset).value = "Champs requis présents"
            .Range(Chr(Asc("I") + HOffset) & VOffset).AddComment ("La feuille manque des attributs indispensables " & Chr(10) & " à son transfert dans la base de donnée" & Chr(10) & "(n°Client, Pharmacode, Désignation).")
        .Range(Chr(Asc("J") + HOffset) & VOffset).value = "attributs manquants"
            .Range(Chr(Asc("J") + HOffset) & VOffset).AddComment ("Les attributs des colonnes (titres) doivent se" & Chr(10) & "trouver sur la première ligne de la feuille" & Chr(10) & "de calcul de manière contiguë." & Chr(10) & "Assurez-vous que ce soit le cas.")
        .Range(Chr(Asc("K") + HOffset) & VOffset).value = "Champs inconnus"
            .Range(Chr(Asc("K") + HOffset) & VOffset).AddComment ("Les attributs reportés sont inconnus de l'application." & Chr(10) & "Enregistrez-les dans la table [attributes] de la" & Chr(10) & "feuille de calcul cachée [INTERNALS]. Si le type " & Chr(10) & "d'attribut n'existe pas (DBB_name)," & Chr(10) & "créez-en un nouveau dans la table suivante" & Chr(10) & "[AttributeTypeAndPlacement] et renseignez un n°" & Chr(10) & "de colonne (DBB_col) non-utilisé et un type.")
        
        Call FitComments
        
        counter = VOffset + 1
        
        For Each FILE In FilesList
'A
            .Range(Chr(Asc("A") + HOffset) & counter).value = counter - VOffset
'B
            .Range(Chr(Asc("B") + HOffset) & counter).value = Left(FILE, InStrRev(FILE, "\"))
'C
            .Hyperlinks.Add Anchor:=.Range(Chr(Asc("C") + HOffset) & counter), _
                            Address:=FILE, _
                            TextToDisplay:=Right(FILE, Len(FILE) - InStrRev(FILE, "\"))
            'Verify the Overall status
'NEEDS TO BE COMPLETED WITH OTHER CONDITIONS
'D
            CondStatusOK = nb_sheets(counter - VOffset - 1) = 1 _
                        And ConformableFileName(.Range(Chr(Asc("C") + HOffset) & counter).value) _
                        And InStr(INTERNALS.ListObjects("file_to_load").ListColumns("required_fields_ok").DataBodyRange(counter - VOffset).value, "FAUX") = 0 _
                        And InStr(INTERNALS.ListObjects("file_to_load").ListColumns("more_than_one_empty_column").DataBodyRange(counter - VOffset).value, "VRAI") = 0 _
                        And Len(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - VOffset).value) = 0
                        
            If CondStatusOK Then
                Status(1).Copy
                .Range(Chr(Asc("D") + HOffset) & counter).PasteSpecial Paste:=xlPasteAll
            Else
                Status(2).Copy
                .Range(Chr(Asc("D") + HOffset) & counter).PasteSpecial Paste:=xlPasteAll
            End If
'/NEEDS TO BE COMPLETED WITH OTHER CONDITIONS
'E
            'n° d'EMS
            .Range(Chr(Asc("E") + HOffset) & counter).value = Left(.Range("C" & counter).value, InStr(.Range("C" & counter).value, "_") - 1)
'F
            'n° d'EMS conforme?
            If ConformableFileName(.Range(Chr(Asc("C") + HOffset) & counter).value) Then
                Status(1).Copy
                .Range(Chr(Asc("F") + HOffset) & counter).PasteSpecial Paste:=xlPasteAll
            Else
                Status(2).Copy
                .Range(Chr(Asc("F") + HOffset) & counter).PasteSpecial Paste:=xlPasteAll
            End If
'G
            'nb of sheets
            If (PARAM_TABLE.Columns(1).Find("VerifyNbSheets").Offset(0, 1).value) Then
                .Range(Chr(Asc("G") + HOffset) & counter).value = nb_sheets(counter - VOffset - 1)
                Call ApplyStyle(.Range(Chr(Asc("G") + HOffset) & counter), "=1", "xlGreater")
            End If
'H
            'Type problems
            If (PARAM_TABLE.Columns(1).Find("VerifyColumnsContent").Offset(0, 1).value) Then
                .Range(Chr(Asc("H") + HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("typing").DataBodyRange(counter - VOffset).value
            End If
'H
            
'I
            .Range(Chr(Asc("I") + HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("required_fields_ok").DataBodyRange(counter - VOffset).value
            Call ApplyStyle(.Range(Chr(Asc("I") + HOffset) & counter), "FAUX", "xlEqual")
'J
            .Range(Chr(Asc("J") + HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("more_than_one_empty_column").DataBodyRange(counter - VOffset).value
            Call ApplyStyle(.Range(Chr(Asc("J") + HOffset) & counter), "VRAI", "xlEqual")
'K
            .Range(Chr(Asc("K") + HOffset) & counter).value = Right(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - VOffset).value, Application.Max(Len(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - VOffset).value) - 1, 0))
            Call ApplyStyle(.Range(Chr(Asc("K") + HOffset) & counter), "=""""", "xlNotEqual")

            
            counter = counter + 1
        Next
        
        .Range(Chr(Asc("A") + HOffset) & VOffset & ":" & Chr(Asc("L") + HOffset) & VOffset).Font.Bold = True
        With .Range(Chr(Asc("C") + HOffset) & VOffset & ":" & Chr(Asc("L") + HOffset) & counter).Cells
            .Columns.AutoFit 'ColumnWidth = 14.5
        End With
        
        
    End With
    
    'RedoRib
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = "Chargement terminé"

End Sub
