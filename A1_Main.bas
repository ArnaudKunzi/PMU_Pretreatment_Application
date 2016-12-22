Attribute VB_Name = "A1_Main"
Sub Start(control As IRibbonControl)
    Call DefGlobal
    StartForm.Show
End Sub


Sub PrepareOverviewSheet(FilesListSring As String)
    Call DefGlobal
    Dim FilesList As Variant
        FilesList = Split(FilesListSring, "|")
        
    Dim ws As Worksheet
        
    Dim counter As Integer
    Dim nb_sheets As Variant
            
    Dim CondStatusOK As Boolean
        
    Call SaveFilesList(FilesList)
        
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets(REPORT_SH.Name).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = REPORT_SH.Name
    Debug.Print ws.VBProject.VBComponents(ws.Name).Properties("Codename")
    
    Call SetWsName(Worksheets(REPORT_SH.Name), "A_1")
    
    nb_sheets = HowManySheets(FilesList)
    
    Call MainLoadingLoop(FilesList, nb_sheets)
    
    Application.ScreenUpdating = False
    
    With Worksheets(REPORT_SH.Name)
        .Cells.Font.Size = "8"
        .Columns("A:B").Group
        .Outline.ShowLevels ColumnLevels:=1
        .Range(Chr(Asc("A") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "n°"
        .Range(Chr(Asc("B") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "Chemin"
        .Range(Chr(Asc("C") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "Nom"
        .Range(Chr(Asc("D") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "Status"
        .Range(Chr(Asc("E") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "n° EMS"
        .Range(Chr(Asc("F") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "Pharmacien"
        .Range(Chr(Asc("G") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "# onglets"
            .Range(Chr(Asc("G") + REPORT_SH.HOffset) & REPORT_SH.VOffset).AddComment ("Seules les données présentes dans le" & Chr(10) & "premier onglet sont prises en compte." & Chr(10) _
                                                                     & "assurez-vous que toutes les données" & Chr(10) & "pertinentes soient dans une" & Chr(10) _
                                                                     & "table dans le premier onglet.")
        .Range(Chr(Asc("H") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "typage"
            .Range(Chr(Asc("H") + REPORT_SH.HOffset) & REPORT_SH.VOffset).AddComment ("Cellules comportant une valeur inattendue (ex: 'chaise'" & Chr(10) & " pour la variable Pharmacode est une valeur de type incorrect).")
            
        .Range(Chr(Asc("I") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "Champs requis"
            .Range(Chr(Asc("I") + REPORT_SH.HOffset) & REPORT_SH.VOffset).AddComment ("La feuille manque des attributs indispensables " & Chr(10) & " à son transfert dans la base de donnée" & Chr(10) & "(n°Client, Pharmacode, Désignation).")
        .Range(Chr(Asc("J") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "attributs manquants"
            .Range(Chr(Asc("J") + REPORT_SH.HOffset) & REPORT_SH.VOffset).AddComment ("Les attributs des colonnes (titres) doivent se" & Chr(10) & "trouver sur la première ligne de la feuille" & Chr(10) & "de calcul de manière contiguë." & Chr(10) & "Assurez-vous que ce soit le cas.")
        .Range(Chr(Asc("K") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "Champs inconnus"
            .Range(Chr(Asc("K") + REPORT_SH.HOffset) & REPORT_SH.VOffset).AddComment ("Les attributs reportés sont inconnus de l'application." & Chr(10) & "Enregistrez-les dans la table [attributes] de la" & Chr(10) & "feuille de calcul cachée [INTERNALS]. Si le type " & Chr(10) & "d'attribut n'existe pas (DBB_name)," & Chr(10) & "créez-en un nouveau dans la table suivante" & Chr(10) & "[AttributeTypeAndPlacement] et renseignez un n°" & Chr(10) & "de colonne (DBB_col) non-utilisé et un type.")
        .Range(Chr(Asc("L") + REPORT_SH.HOffset) & REPORT_SH.VOffset).value = "Pharmacode"
            .Range(Chr(Asc("L") + REPORT_SH.HOffset) & REPORT_SH.VOffset).AddComment ("Nombre de pharmacodes" & Chr(10) & "invalides détectés")
        
        Call FitComments
        
        counter = REPORT_SH.VOffset + 1
        
        For Each FILE In FilesList
'A
            .Range(Chr(Asc("A") + REPORT_SH.HOffset) & counter).value = counter - REPORT_SH.VOffset
'B
            .Range(Chr(Asc("B") + REPORT_SH.HOffset) & counter).value = Left(FILE, InStrRev(FILE, "\"))
'C
            .Hyperlinks.Add Anchor:=.Range(Chr(Asc("C") + REPORT_SH.HOffset) & counter), _
                            Address:=FILE, _
                            TextToDisplay:=Right(FILE, Len(FILE) - InStrRev(FILE, "\"))
            'Verify the Overall status
'NEEDS TO BE COMPLETED WITH OTHER CONDITIONS
'D
            CondStatusOK = nb_sheets(counter - REPORT_SH.VOffset - 1) = 1 _
                        And ConformableFileName(.Range(Chr(Asc("C") + REPORT_SH.HOffset) & counter).value) _
                        And InStr(INTERNALS.ListObjects("file_to_load").ListColumns("required_fields_ok").DataBodyRange(counter - REPORT_SH.VOffset).value, "FAUX") = 0 _
                        And InStr(INTERNALS.ListObjects("file_to_load").ListColumns("more_than_one_empty_column").DataBodyRange(counter - REPORT_SH.VOffset).value, "VRAI") = 0 _
                        And Len(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - REPORT_SH.VOffset).value) = 0
                        
            If CondStatusOK Then
                Status(1).Copy
                .Range(Chr(Asc("D") + REPORT_SH.HOffset) & counter).PasteSpecial Paste:=xlPasteAll
            Else
                Status(2).Copy
                .Range(Chr(Asc("D") + REPORT_SH.HOffset) & counter).PasteSpecial Paste:=xlPasteAll
            End If
'/NEEDS TO BE COMPLETED WITH OTHER CONDITIONS
'E
            'n° d'EMS
            .Range(Chr(Asc("E") + REPORT_SH.HOffset) & counter).value = Left(.Range("C" & counter).value, InStr(.Range("C" & counter).value, "_") - 1)
            If ConformableFileName(.Range(Chr(Asc("C") + REPORT_SH.HOffset) & counter).value) Then
                Status(1).Copy
                .Range(Chr(Asc("E") + REPORT_SH.HOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            Else
                Status(2).Copy
                .Range(Chr(Asc("E") + REPORT_SH.HOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            End If
'F
            'n° d'EMS conforme?
            .Range(Chr(Asc("F") + REPORT_SH.HOffset) & counter).value = Mid(.Range("C" & counter).value, _
                                                                    InStr(.Range("C" & counter).value, "_") + 1, _
                                                                        InStr(InStr(.Range("C" & counter).value, "_") + 1, _
                                                                                .Range("C" & counter).value, "_") - InStr(.Range("C" & counter).value, "_") - 1)
            'If ConformableFileName(.Range(Chr(Asc("C") + REPORT_SH.HOffset) & counter).value) Then
            '    Status(1).Copy
            '    .Range(Chr(Asc("F") + REPORT_SH.HOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            'Else
            '    Status(2).Copy
            '    .Range(Chr(Asc("F") + REPORT_SH.HOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            'End If
'G
            'nb of sheets
            If (PARAM_TABLE.Columns(1).Find("VerifyNbSheets").Offset(0, 1).value) Then
                .Range(Chr(Asc("G") + REPORT_SH.HOffset) & counter).value = nb_sheets(counter - REPORT_SH.VOffset - 1)
                Call ApplyStyle(.Range(Chr(Asc("G") + REPORT_SH.HOffset) & counter), "=1", "xlGreater", "bad")
                Call ApplyStyle(.Range(Chr(Asc("G") + REPORT_SH.HOffset) & counter), "=1", "xlEqual", "good")
            End If
'H
            'Type problems
            If (PARAM_TABLE.Columns(1).Find("VerifyColumnsContent").Offset(0, 1).value) Then
                .Range(Chr(Asc("H") + REPORT_SH.HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("typing").DataBodyRange(counter - REPORT_SH.VOffset).value
                .Range(Chr(Asc("H") + REPORT_SH.HOffset) & counter).WrapText = False
                Call ApplyStyle(.Range(Chr(Asc("H") + REPORT_SH.HOffset) & counter), "=""""", "xlNotEqual", "bad")
            End If
            
            If (PARAM_TABLE.Columns(1).Find("VerifyColumnsTitle").Offset(0, 1).value) Then
'I
                .Range(Chr(Asc("I") + REPORT_SH.HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("required_fields_ok").DataBodyRange(counter - REPORT_SH.VOffset).value
                Call ApplyStyle(.Range(Chr(Asc("I") + REPORT_SH.HOffset) & counter), "FAUX", "xlEqual", "bad")
                Call ApplyStyle(.Range(Chr(Asc("I") + REPORT_SH.HOffset) & counter), "VRAI", "xlEqual", "good")
'J
                .Range(Chr(Asc("J") + REPORT_SH.HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("more_than_one_empty_column").DataBodyRange(counter - REPORT_SH.VOffset).value
                Call ApplyStyle(.Range(Chr(Asc("J") + REPORT_SH.HOffset) & counter), "VRAI", "xlEqual", "bad")
                Call ApplyStyle(.Range(Chr(Asc("J") + REPORT_SH.HOffset) & counter), "=""""", "xlEqual", "good")
'K
                .Range(Chr(Asc("K") + REPORT_SH.HOffset) & counter).value = Right(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - REPORT_SH.VOffset).value, Application.Max(Len(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - REPORT_SH.VOffset).value) - 1, 0))
                Call ApplyStyle(.Range(Chr(Asc("K") + REPORT_SH.HOffset) & counter), "=""""", "xlNotEqual", "bad")
                Call ApplyStyle(.Range(Chr(Asc("K") + REPORT_SH.HOffset) & counter), "=""""", "xlEqual", "good")
            End If
            
            If PARAM_TABLE.Columns(1).Find("CheckPharmacodes").Offset(0, 1).value Then
'L
                .Range(Chr(Asc("L") + REPORT_SH.HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("invalid_pharmacodes").DataBodyRange(counter - REPORT_SH.VOffset).value
                Call ApplyStyle(.Range(Chr(Asc("L") + REPORT_SH.HOffset) & counter), "=0", "xlGreater", "bad")
                Call ApplyStyle(.Range(Chr(Asc("L") + REPORT_SH.HOffset) & counter), "=0", "xlEqual", "good")
            End If
            
            counter = counter + 1
        Next
        
        .Range(Chr(Asc("A") + REPORT_SH.HOffset) & REPORT_SH.VOffset & ":" & Chr(Asc("L") + REPORT_SH.HOffset) & REPORT_SH.VOffset).Font.Bold = True
        With .Range(Chr(Asc("C") + REPORT_SH.HOffset) & REPORT_SH.VOffset & ":" & Chr(Asc("L") + REPORT_SH.HOffset) & counter).Cells
            .Columns.AutoFit 'ColumnWidth = 14.5
        End With
        .Range("H:H").Columns.ColumnWidth = 10
        .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
        
        lcol = .Cells(1, .Columns.Count).End(xlToLeft).column
        lrow = .Cells(.Rows.Count, "A").End(xlUp).row
        
        
        'Debug.Print .Range(Cells(1, lcol + 1), Cells(.Rows.Count, .Columns.Count)).Address
        '.Range(Cells(1, lcol + 1), Cells(.Rows.Count, .Columns.Count)).EntireColumn.Hidden = True
        
        '.Range(Cells(lrow + 1, 1), Cells(.Rows.Count, .Columns.Count)).EntireRow.Hidden = True
        
           
        ActiveWindow.SplitColumn = 0
        ActiveWindow.SplitRow = 1
        ActiveWindow.FreezePanes = True
        
    End With
    
    Application.ScreenUpdating = True
    
    Application.StatusBar = "Chargement terminé"

End Sub

Sub UpdateStage(NewStage As Integer)
    Call DefGlobal
    If NewStage > 0 And NewStage < 6 Then
        STAGE.value = NewStage
    End If
    'Call GetInstructionLabel(Nothing, NewStage)  '(control As IRibbonControl, ByRef returnedVal)
    
    Select Case NewStage
        Case 1:
            DisplayTag = "*VG_1*"
        Case 2:
            DisplayTag = "*VG_*2*"
        Case 3:
            DisplayTag = "*VG_*3*"
        Case 4:
            DisplayTag = "*VG_*4*"
        Case 5:
            DisplayTag = "*VG_*5*"
        Case Else
            DisplayTag = "*VG_*"
    End Select
        
    Call RefreshRibbon(DisplayTag)
    'Call RefreshButton(DisplayTag)
End Sub

