Attribute VB_Name = "A1_Main"

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
    Call SetWsName(Worksheets("RAPPORT"), "A_1")
    
    nb_sheets = HowManySheets(FilesList)
    
    Call MainLoadingLoop(FilesList, nb_sheets)
    
    Application.ScreenUpdating = False
    
    With Worksheets("RAPPORT")
        .Cells.Font.Size = "8"
        .Columns("A:B").Group
        .Outline.ShowLevels ColumnLevels:=1
        .Range(Chr(Asc("A") + HOffset) & VOffset).value = "n�"
        .Range(Chr(Asc("B") + HOffset) & VOffset).value = "Chemin"
        .Range(Chr(Asc("C") + HOffset) & VOffset).value = "Nom"
        .Range(Chr(Asc("D") + HOffset) & VOffset).value = "Status"
        .Range(Chr(Asc("E") + HOffset) & VOffset).value = "n� EMS"
        .Range(Chr(Asc("F") + HOffset) & VOffset).value = "Pharmacien"
        .Range(Chr(Asc("G") + HOffset) & VOffset).value = "# onglets"
            .Range(Chr(Asc("G") + HOffset) & VOffset).AddComment ("Seules les donn�es pr�sentes dans le" & Chr(10) & "premier onglet sont prises en compte." & Chr(10) _
                                                                     & "assurez-vous que toutes les donn�es" & Chr(10) & "pertinentes soient dans une" & Chr(10) _
                                                                     & "table dans le premier onglet.")
        .Range(Chr(Asc("H") + HOffset) & VOffset).value = "typage"
            .Range(Chr(Asc("H") + HOffset) & VOffset).AddComment ("Cellules comportant une valeur inattendue (ex: 'chaise'" & Chr(10) & " pour la variable Pharmacode est une valeur de type incorrect).")
            
        .Range(Chr(Asc("I") + HOffset) & VOffset).value = "Champs requis"
            .Range(Chr(Asc("I") + HOffset) & VOffset).AddComment ("La feuille manque des attributs indispensables " & Chr(10) & " � son transfert dans la base de donn�e" & Chr(10) & "(n�Client, Pharmacode, D�signation).")
        .Range(Chr(Asc("J") + HOffset) & VOffset).value = "attributs manquants"
            .Range(Chr(Asc("J") + HOffset) & VOffset).AddComment ("Les attributs des colonnes (titres) doivent se" & Chr(10) & "trouver sur la premi�re ligne de la feuille" & Chr(10) & "de calcul de mani�re contigu�." & Chr(10) & "Assurez-vous que ce soit le cas.")
        .Range(Chr(Asc("K") + HOffset) & VOffset).value = "Champs inconnus"
            .Range(Chr(Asc("K") + HOffset) & VOffset).AddComment ("Les attributs report�s sont inconnus de l'application." & Chr(10) & "Enregistrez-les dans la table [attributes] de la" & Chr(10) & "feuille de calcul cach�e [INTERNALS]. Si le type " & Chr(10) & "d'attribut n'existe pas (DBB_name)," & Chr(10) & "cr�ez-en un nouveau dans la table suivante" & Chr(10) & "[AttributeTypeAndPlacement] et renseignez un n�" & Chr(10) & "de colonne (DBB_col) non-utilis� et un type.")
        .Range(Chr(Asc("L") + HOffset) & VOffset).value = "Pharmacode"
            .Range(Chr(Asc("L") + HOffset) & VOffset).AddComment ("Nombre de pharmacodes" & Chr(10) & "invalides d�tect�s")
        
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
            'n� d'EMS
            .Range(Chr(Asc("E") + HOffset) & counter).value = Left(.Range("C" & counter).value, InStr(.Range("C" & counter).value, "_") - 1)
            If ConformableFileName(.Range(Chr(Asc("C") + HOffset) & counter).value) Then
                Status(1).Copy
                .Range(Chr(Asc("E") + HOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            Else
                Status(2).Copy
                .Range(Chr(Asc("E") + HOffset) & counter).PasteSpecial Paste:=xlPasteFormats
            End If
'F
            'n� d'EMS conforme?
            .Range(Chr(Asc("F") + HOffset) & counter).value = Mid(.Range("C" & counter).value, _
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
                .Range(Chr(Asc("G") + HOffset) & counter).value = nb_sheets(counter - VOffset - 1)
                Call ApplyStyle(.Range(Chr(Asc("G") + HOffset) & counter), "=1", "xlGreater", "bad")
                Call ApplyStyle(.Range(Chr(Asc("G") + HOffset) & counter), "=1", "xlEqual", "good")
            End If
'H
            'Type problems
            If (PARAM_TABLE.Columns(1).Find("VerifyColumnsContent").Offset(0, 1).value) Then
                .Range(Chr(Asc("H") + HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("typing").DataBodyRange(counter - VOffset).value
                .Range(Chr(Asc("H") + HOffset) & counter).WrapText = False
                Call ApplyStyle(.Range(Chr(Asc("H") + HOffset) & counter), "=""""", "xlNotEqual", "bad")
            End If
            
            If (PARAM_TABLE.Columns(1).Find("VerifyColumnsTitle").Offset(0, 1).value) Then
'I
                .Range(Chr(Asc("I") + HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("required_fields_ok").DataBodyRange(counter - VOffset).value
                Call ApplyStyle(.Range(Chr(Asc("I") + HOffset) & counter), "FAUX", "xlEqual", "bad")
                Call ApplyStyle(.Range(Chr(Asc("I") + HOffset) & counter), "VRAI", "xlEqual", "good")
'J
                .Range(Chr(Asc("J") + HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("more_than_one_empty_column").DataBodyRange(counter - VOffset).value
                Call ApplyStyle(.Range(Chr(Asc("J") + HOffset) & counter), "VRAI", "xlEqual", "bad")
                Call ApplyStyle(.Range(Chr(Asc("J") + HOffset) & counter), "=""""", "xlEqual", "good")
'K
                .Range(Chr(Asc("K") + HOffset) & counter).value = Right(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - VOffset).value, Application.Max(Len(INTERNALS.ListObjects("file_to_load").ListColumns("unidentified_fields").DataBodyRange(counter - VOffset).value) - 1, 0))
                Call ApplyStyle(.Range(Chr(Asc("K") + HOffset) & counter), "=""""", "xlNotEqual", "bad")
                Call ApplyStyle(.Range(Chr(Asc("K") + HOffset) & counter), "=""""", "xlEqual", "good")
            End If
            
            If PARAM_TABLE.Columns(1).Find("CheckPharmacodes").Offset(0, 1).value Then
'L
                .Range(Chr(Asc("L") + HOffset) & counter).value = INTERNALS.ListObjects("file_to_load").ListColumns("invalid_pharmacodes").DataBodyRange(counter - VOffset).value
                Call ApplyStyle(.Range(Chr(Asc("L") + HOffset) & counter), "=0", "xlGreater", "bad")
                Call ApplyStyle(.Range(Chr(Asc("L") + HOffset) & counter), "=0", "xlEqual", "good")
            End If
            
            counter = counter + 1
        Next
        
        .Range(Chr(Asc("A") + HOffset) & VOffset & ":" & Chr(Asc("L") + HOffset) & VOffset).Font.Bold = True
        With .Range(Chr(Asc("C") + HOffset) & VOffset & ":" & Chr(Asc("L") + HOffset) & counter).Cells
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
    
    Application.StatusBar = "Chargement termin�"

End Sub

Sub UpdateStage(NewStage As String)
    STAGE.value = NewStage
    'Call GetInstructionLabel(Nothing, NewStage)  '(control As IRibbonControl, ByRef returnedVal)
End Sub

