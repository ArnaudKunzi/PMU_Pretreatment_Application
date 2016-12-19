Attribute VB_Name = "B22_ManualEdit"

Sub RegisterChange(ChangeRange As Range)
    Call DefGlobal
    Dim NewValues As Variant
    Dim OldComments As Variant
    Dim rly As VbMsgBoxResult
    
    If ChangeRange.Count > 10000 Then GoTo HandleTooManyPaste
    
    OldComments = Split(LastCommentsSelected, "||")
    
    With ChangeRange
        
        NewValues = Application.Transpose(.value)
                
            If .Count > 1 Then
                'if we replace several cells
                On Error GoTo HandleInequalRanges
                If UBound(NewValues) = UBound(LastValueSelected) Then
                    If Not Join(NewValues, "|") = Join(LastValueSelected, "|") Then
                        On Error GoTo 0
                        .Cells.ClearComments
                        For i = 1 To ChangeRange.Count
                            'if there already is a comment
                           If Len(OldComments(i)) < 1 Then
                                ChangeRange(i).AddComment "Original value: " & CStr(LastValueSelected(i)) & _
                                                          vbNewLine & Format(Now(), "yyyy.mm.dd hh:mm") & "|" & Application.UserName & ": " & NewValues(i)
                            Else
                                ChangeRange(i).AddComment OldComments(i) & vbNewLine & Format(Now(), "yyyy.mm.dd hh:mm") & "|" & Application.UserName & ": " & NewValues(i)
                            End If
                            ChangeRange(i).Comment.Shape.TextFrame.AutoSize = True
                        Next i
                    End If
                    .Interior.ColorIndex = EDITCOLOR
                    Call CommentStyle(ChangeRange)
                Else
                    GoTo HandleInequalRanges
                End If
                
            'if we replace a single cell
            Else
                If Not NewValues = LastValueSelected Then
                    .Cells.ClearComments
                     'if there already is a comment
                    If Len(OldComments(1)) < 1 Then
                        .AddComment "Original value: " & CStr(LastValueSelected) & _
                                    vbNewLine & Format(Now(), "yyyy.mm.dd hh:mm") & "|" & Application.UserName & ": " & .value
                    Else
                        .AddComment OldComments(1) & vbNewLine & Format(Now(), "yyyy.mm.dd hh:mm") & "|" & Application.UserName & ": " & .value
                    End If
                    'already in: Call CommentStyle(ChangeRange)
                    '.Comment.Shape.TextFrame.AutoSize = True
                    
                    .Interior.ColorIndex = EDITCOLOR
                    Call CommentStyle(ChangeRange)
                End If
                
            End If
            
    End With
        
Exit Sub
HandleInequalRanges:
    Dim choice
    rly = MsgBox("Pour raison de sécurité, l'application n'autorise pas les collages sans" & _
           " sélection explicite de la plage de destination. Nous recommendons d'annuler le collage.", vbYesNo)
    If rly = vbYes Then
        With Application
            .EnableEvents = False
            .Undo
            .EnableEvents = True
         End With
    End If
    Exit Sub
HandleTooManyPaste:
    With Application
        .EnableEvents = False
        .Undo
        .EnableEvents = True
     End With
    MsgBox "Pour raison de sécurité, l'application n'autorise pas les collages de plus" & _
    " de 10'000 cellules. Le collage a été annulé.", vbCritical
    Exit Sub
End Sub

Sub test()
    Call ProduceLog(Worksheets("InvalidPharmacodes"))
End Sub

Sub ProduceLog(control As IRibbonControl)
    Call DefGlobal
    Dim ws As Worksheet
    Dim CommentsTexts() As String
    Dim i As Long
    Dim LastRow As Long
    

    Set ws = ActiveSheet
    
    If Not (InStr(ws.CodeName, DataSheetName) <> 0 Or ws.CodeName = InPh_colname) Then Exit Sub
    
    If ws.comments.Count = 0 Then
        MsgBox "Rien à journaliser", vbExclamation
        Exit Sub
    End If
    
    If Not SheetExists("LOG_" & Year) Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = LogSheetName
        Worksheets(LogSheetName).Tab.ColorIndex = EXPORTCOLOR
        Call SetWsName(Worksheets(LogSheetName), "LOG_EDITS")
    Else
        Worksheets(LogSheetName).Cells.Clear
    End If
    
    Worksheets(LogSheetName).Range("A1").value = "LOG FEUILLE " & ws.Name
    
    ReDim CommentsTexts(ws.comments.Count - 1)
    i = 0
    For Each CellComment In ws.comments
        CommentsTexts(i) = Worksheets(ws.Name).Range(Mid(CellComment.Parent.Address, 2, 1) & "1").value _
                        & " l." & CellComment.Parent.row & ": " & CellComment.Text
        i = i + 1
    Next CellComment
    
    'Sort the array by columns:
    Dim a As Long, b As Long
    Dim Temp1 As String, Temp2 As String
    
    'Alphabetize Sheet Names in Array List
      For a = LBound(CommentsTexts) To UBound(CommentsTexts)
        For b = a To UBound(CommentsTexts)
          If UCase(CommentsTexts(b)) < UCase(CommentsTexts(a)) Then
            Temp1 = CommentsTexts(a)
            Temp2 = CommentsTexts(b)
            CommentsTexts(a) = Temp2
            CommentsTexts(b) = Temp1
          End If
         Next b
      Next a
    
    With Worksheets(LogSheetName)
    
        'LastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        .Range("A" & 2).Resize(ws.comments.Count, 1) = Application.Transpose(CommentsTexts)
        '.Range("B" & 2).Resize(ws.Comments.Count, 1) = 1
        '.Range("C" & 2).Resize(ws.Comments.Count, 1) = 2
        .Range("A:D").WrapText = False
        .Columns.AutoFit
        .Rows.AutoFit
        .Range("A:D").WrapText = True
        
    End With
    
End Sub

Sub testc()
    Call CommentStyle(Worksheets("InvalidPharmacodes"))
End Sub

Sub CommentStyle(ByRef rangewithcomment As Range) 'ByRef ws As Worksheet)
    
    Dim CommentsToStyle As Comment
    For Each cellwithcomment In rangewithcomment 'ws.comments
        Set CommentsToStyle = cellwithcomment.Cells.Comment
        If CommentsToStyle Is Nothing Then GoTo NextIteration
            With CommentsToStyle
                .Shape.TextFrame.AutoSize = True
                .Shape.AutoShapeType = msoShapeRoundedRectangle
                .Shape.TextFrame.Characters.Font.Name = "Tahoma"
                .Shape.TextFrame.Characters.Font.Size = 8
                .Shape.TextFrame.Characters.Font.ColorIndex = 1
                .Shape.Line.ForeColor.RGB = RGB(0, 0, 0)
                .Shape.Line.BackColor.RGB = RGB(255, 255, 255)
                .Shape.Fill.visible = msoTrue
                .Shape.Fill.ForeColor.RGB = RGB(153, 255, 255)
                '.Shape.Fill.OneColorGradient msoGradientDiagonalUp, 1, 0.5
            End With
NextIteration:
    Next cellwithcomment

End Sub

Sub test3()
    Debug.Print GetComments(Worksheets("InvalidPharmacodes").Range("C1:C1"))
End Sub

Function GetComments(Target As Range)
    Call DefGlobal
    Dim comments() As String
    ReDim comments(Target.Count)
    For i = LBound(comments) + 1 To UBound(comments)
        If Target(i).Cells.Interior.ColorIndex = EDITCOLOR Then
            On Error Resume Next
            comments(i) = Target(i).Comment.Text
        End If
    Next i
    GetComments = Join(comments, "||")
End Function

Sub RestoreFirstValue()
    Call DefGlobal
    Dim commentcontent As String
    Dim FirstValue As Variant
    On Error GoTo Abort
    Application.EnableEvents = False
     For Each c In Selection
        With c
            If .Interior.ColorIndex = EDITCOLOR Then
                commentcontent = .Comment.Text
                commentcontent = Left(commentcontent, InStr(commentcontent, vbNewLine) - 1)
                FirstValue = Right(commentcontent, Len(commentcontent) - InStrRev(commentcontent, ": ") - 1)
                .value = FirstValue
                .ClearComments
                .Interior.ColorIndex = xlNone
            End If
        End With
    Next c
Abort:
     Application.EnableEvents = True
End Sub

Sub RevertToLastValue()
    Call DefGlobal
    Dim LastValue As Variant
    Dim commentcontent As String
    Dim temparray As Variant
    'On Error GoTo Abort
    Application.EnableEvents = False
    For Each c In Selection
        With c
            If .Interior.ColorIndex = EDITCOLOR Then
                commentcontent = .Comment.Text
                temparray = Split(commentcontent, vbNewLine)
                commentcontent = temparray(UBound(temparray) - 1)
                LastValue = Right(commentcontent, Len(commentcontent) - InStrRev(commentcontent, ": ") - 1)
                .value = LastValue
                .ClearComments
                If UBound(temparray) > 1 Then
                    ReDim Preserve temparray(UBound(temparray) - 1)
                    commentcontent = Join(temparray, vbNewLine)
                    .AddComment commentcontent
                    Call CommentStyle(Range(c.Address))
                Else
                    .Interior.ColorIndex = xlNone
                End If
            End If
        End With
    Next c
Abort:
     Application.EnableEvents = True
End Sub


Sub ColorLabelling(ByRef Target As Range)

    Call DefGlobal
    
    Dim RangeToCheck As Range
    Dim ws As Worksheet
    Dim hOffset As Integer
    
    hOffset = 5
    
    Set ws = Target.Parent
    Set RangeToCheck = ws.Range(ws.Cells(Target.row, hOffset + 1), ws.Cells(Target.row, ws.Cells(1, 1).End(xlToRight).column))
    
    If WorksheetFunction.CountA(RangeToCheck) = 0 Then
        RangeToCheck.Cells.Interior.ColorIndex = 3 'red
    ElseIf Not RangeToCheck.Find("") Is Nothing Then
        'MsgBox "Finissez de compléter l'entrée."
        'RangeToCheck.Select
        RangeToCheck.Cells.Interior.ColorIndex = 45 'orange
    Else
        RangeToCheck.Cells.Interior.ColorIndex = EDITCOLOR
    End If
    
End Sub
