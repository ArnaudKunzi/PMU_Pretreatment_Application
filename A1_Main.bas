Attribute VB_Name = "A1_Main"

Sub PrepareOverviewSheet(FilesListSring As String)
    
    Dim FilesList As Variant
        FilesList = Split(FilesListSring, "|")
    Dim counter As Integer
    Dim nb_sheets As Variant
    
    Dim hOffset As Integer
        hOffset = 0
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
    Call SetWsName(Worksheets("RAPPORT"), "A_1")
    
    nb_sheets = HowManySheets(FilesList)
    
    Call MainLoadingLoop(FilesList, nb_sheets)
    
    Application.ScreenUpdating = False
    
    With Worksheets("RAPPORT")
        .Cells.Font.Size = "8"
        .Columns("A:B").Group
        .Outline.ShowLevels ColumnLevels:=1
        .Range(Chr(Asc("A") + hOffset) & VOffset).value = "n°"
        .Range(Chr(Asc("B") + hOffset) & VOffset).value = "Chemin"
        .Range(Chr(Asc("C") + hOffset) & VOffset).value = "Nom"
        .Range(Chr(Asc("D") + hOffset) & VOffset).value = "Status"
        .Range(Chr(Asc("E") + hOffset) & VOffset).value = "n° EMS"
        .Range(Chr(Asc("F") + hOffset) & VOffset).value = "Pharmacien"
        .Range(Chr(Asc("G") + hOffset) & VOffset).value = "# onglets"
            .Range(Chr(Asc("G") + hOffset) & VOffset).AddComment ("Seules les données présentes dans le" & Chr(10) & "premier onglet sont prises en compte." & Chr(10) _
                                                                     & "assurez-vous que toutes les données" & Chr(10) & "pertinentes soient dans une" & Chr(10) _
                                                                     & "table dans le premier onglet.")
        .Range(Chr(Asc("H") + hOffset) & VOffset).value = "typage"
            .Range(Chr(Asc("H") + hOffset) & VOffset).AddComment ("Cellules comportant une valeur inattendue (ex: 'chaise'" & Chr(10) & " pour la variable Pharmacode est une valeur de type incorrect).")
            
        .Range(Chr(Asc("I") + hOffset) & VOffset).value = "Champs requis"
            .Range(Chr(Asc("I") + hOffset) & VOffset).AddComment ("La feuille manque des attributs indispensables " & Chr(10) & " à son transfert dans la base de donnée" & Chr(10) & "(n°Client, Pharmacode, Désignation).")
        .Range(Chr(Asc("J") + hOffset) & VOffset).value = "attributs manquants"
            .Range(Chr(Asc("J") + hOffset) & VOffset).AddComment ("Les attributs des colonnes (titres) doivent se" & Chr(10) & "trouver sur la première ligne de la feuille" & Chr(10) & "de calcul de manière contiguë." & Chr(10) & "Assurez-vous que ce soit le cas.")
        .Range(Chr(Asc("K") + hOffset) & VOffset).value = "Champs inconnus"
            .Range(Chr(Asc("K") + hOffset) & VOffset).AddComment ("Les attributs reportés sont inconnus de l'application." & Chr(10) & "Enregistrez-les dans la table [attributes] de la" & Chr(10) & "feuille de calcul cachée [INTERNALS]. Si le type " & Chr(10) & "d'attribut n'existe pas (DBB_name)," & Chr(10) & "créez-en un nouveau dans la table suivante" & Chr(10) & "[AttributeTypeAndPlacement] et renseignez un n°" & Chr(10) & "de colonne (DBB_col) non-utilisé et un type.")
        .Range(Chr(Asc("L") + hOffset) & VOffset).value = "Pharmacode"
            .Range(Chr(Asc("L") + hOffset) & VOffset).AddComment ("Nombre de pharmacodes" & Chr(10) & "invalides détectés")
        
        Call FitComments
        
        counter = VOffset + 1
        
        For Each FILE In FilesList
'A
            .Range(Chr(Asc("A") + hOffset) & counter).value = counter - VOffset
'B
            .Range(Chr(Asc("B") + hOffset) & counter).value = Left(FILE, InStrRev(FILE, "\"))
'C
            .Hyperlinks.Add Anchor:=.Range(Chr(Asc("C") + hOffset) & counter), _
                            Address:=FILE, _
                            TextToDisplay:=Right(FILE, Len(FILE) - InStrRev(FILE, "\"))
            'Verify the Overall status
'NEEDS TO BE COMPLETED WITH OTHER CONDITIONS
'D
            CondStatusOK = nb_sheets(counter - VOffset - 1) = 1 _
                        And ConformableFileName(.Range(Chr(Asc("C") + hOffset) & counter).value) _
                        And InStr(INTERNALS.ListObjects("file_to_load").ListColumns("required_fields_ok").DataBodyRange(counter - VOffset).value, "FAUX") = 0 _
                        And InStr(INTERNALS.ListObjects("file_to_load").ListColumns("more_than_one_empty_column").DataBodyRange(counter - VOffset).value, "VRAI") = 0 _
                        And Len(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - VOffset).value) = 0
                        
            If CondStatusOK Then
                Status(1).Copy
                .Range(Chr(Asc("D") + hOffset) & counter).PasteSpecial Paste:=xlPasteAll
            Else
                Status(2).Copy
                .Range(Chr(Asc("D") + hOffset) & counter).PasteSpecial Paste:=xlPasteAll
            End If
'/NEEDS TO BE COMPLETED WITH OTHER CONDITIONS
'E
            'n° d'EMS
            .Range(Chr(Asc("E") + hOffset) & counter).value = Left(.Range("C" & counter).value, InStr(.Range("C" & counter).value, "_") - 1)
            If ConformableFileName(.Range(Chr(Asc("C") + hOffset) & counter).value) Then
                Status(1).Copy
                .Range(Chr(Asc("E") + hOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            Else
                Status(2).Copy
                .Range(Chr(Asc("E") + hOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            End If
'F
            'n° d'EMS conforme?
            .Range(Chr(Asc("F") + hOffset) & counter).value = Mid(.Range("C" & counter).value, _
                                                                    InStr(.Range("C" & counter).value, "_") + 1, _
                                                                        InStr(InStr(.Range("C" & counter).value, "_") + 1, _
                                                                                .Range("C" & counter).value, "_") - InStr(.Range("C" & counter).value, "_") - 1)
            'If ConformableFileName(.Range(Chr(Asc("C") + HOffset) & counter).value) Then
            '    Status(1).Copy
            '    .Range(Chr(Asc("F") + HOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            'Else
            '    Status(2).Copy
            '    .Range(Chr(Asc("F") + HOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            'End If
'G
            'nb of sheets
            If (PARAM_TABLE.Columns(1).Find("VerifyNbSheets").Offset(0, 1).value) Then
                .Range(Chr(Asc("G") + hOffset) & counter).value = nb_sheets(counter - VOffset - 1)
                Call ApplyStyle(.Range(Chr(Asc("G") + hOffset) & counter), "=1", "xlGreater", "bad")
                Call ApplyStyle(.Range(Chr(Asc("G") + hOffset) & counter), "=1", "xlEqual", "good")
            End If
'H
            'Type problems
            If (PARAM_TABLE.Columns(1).Find("VerifyColumnsContent").Offset(0, 1).value) Then
                .Range(Chr(Asc("H") + hOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("typing").DataBodyRange(counter - VOffset).value
                .Range(Chr(Asc("H") + hOffset) & counter).WrapText = False
                Call ApplyStyle(.Range(Chr(Asc("H") + hOffset) & counter), "=""""", "xlNotEqual", "bad")
            End If
            
            If (PARAM_TABLE.Columns(1).Find("VerifyColumnsTitle").Offset(0, 1).value) Then
'I
                .Range(Chr(Asc("I") + hOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("required_fields_ok").DataBodyRange(counter - VOffset).value
                Call ApplyStyle(.Range(Chr(Asc("I") + hOffset) & counter), "FAUX", "xlEqual", "bad")
                Call ApplyStyle(.Range(Chr(Asc("I") + hOffset) & counter), "VRAI", "xlEqual", "good")
'J
                .Range(Chr(Asc("J") + hOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("more_than_one_empty_column").DataBodyRange(counter - VOffset).value
                Call ApplyStyle(.Range(Chr(Asc("J") + hOffset) & counter), "VRAI", "xlEqual", "bad")
                Call ApplyStyle(.Range(Chr(Asc("J") + hOffset) & counter), "=""""", "xlEqual", "good")
'K
                .Range(Chr(Asc("K") + hOffset) & counter).value = Right(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - VOffset).value, Application.Max(Len(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - VOffset).value) - 1, 0))
                Call ApplyStyle(.Range(Chr(Asc("K") + hOffset) & counter), "=""""", "xlNotEqual", "bad")
                Call ApplyStyle(.Range(Chr(Asc("K") + hOffset) & counter), "=""""", "xlEqual", "good")
            End If
            
            If PARAM_TABLE.Columns(1).Find("CheckPharmacodes").Offset(0, 1).value Then
'L
                .Range(Chr(Asc("L") + hOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("invalid_pharmacodes").DataBodyRange(counter - VOffset).value
                Call ApplyStyle(.Range(Chr(Asc("L") + hOffset) & counter), "=0", "xlGreater", "bad")
                Call ApplyStyle(.Range(Chr(Asc("L") + hOffset) & counter), "=0", "xlEqual", "good")
            End If
            
            counter = counter + 1
        Next
        
        .Range(Chr(Asc("A") + hOffset) & VOffset & ":" & Chr(Asc("L") + hOffset) & VOffset).Font.Bold = True
        With .Range(Chr(Asc("C") + hOffset) & VOffset & ":" & Chr(Asc("L") + hOffset) & counter).Cells
            .Columns.AutoFit 'ColumnWidth = 14.5
        End With
        .Range("H:H").Columns.ColumnWidth = 10
        .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
        
        lcol = .Cells(1, .Columns.Count).End(xlToLeft).column
        lrow = .Cells(.Rows.Count, "A").End(xlUp).row
        
        .Range(Cells(1, lcol + 1), Cells(.Rows.Count, .Columns.Count)).EntireColumn.Hidden = True
        .Range(Cells(lrow + 1, 1), Cells(.Rows.Count, .Columns.Count)).EntireRow.Hidden = True
    End With
    
    'RedoRib
    
    Call UpdateStage("PreTreatment")
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = "Chargement terminé"

End Sub

Sub UpdateStage(NewStage As String)
    STAGE.value = NewStage
    'Call GetInstructionLabel(Nothing, NewStage)  '(control As IRibbonControl, ByRef returnedVal)
End Sub

