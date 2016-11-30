Attribute VB_Name = "B22_ManualEdit"

Sub RegisterChange(ChangeRange As Range)
    Call DefGlobal
    Dim NewValues As Variant
    Dim OldComments As Variant
    Dim rly As Integer
    
    OldComments = Split(LastCommentsSelected, "||")
    
    With ChangeRange
        
        NewValues = Application.Transpose(.value)
                
            If .Count > 1 Then
                'if we replace several cells
                On Error GoTo Handle
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
                Else
                    GoTo Handle
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
                    .Comment.Shape.TextFrame.AutoSize = True
                End If
            End If
            .Interior.ColorIndex = EDITCOLOR
            Call CommentStyle(ChangeRange)
    End With
        
Exit Sub
Handle:
    With Application
        .EnableEvents = False
        .Undo
        .EnableEvents = True
     End With
    MsgBox "Pour raison de sécurité, l'application n'autorise pas les collages sans" & _
    " sélection explicite de la plage de destination. Le collage a été annulé.", vbCritical
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
    
    If Not Evaluate("ISREF('" & "LOG_" & Year & "'!A1)") Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = LogSheetName
        'Worksheets(LogSheetName).Activate
        Call SetWsName(Worksheets(LogSheetName), "LOG_EDITS")
        'LOG_EDITS.Range("A1").value = "HAHA" 'Array("Activity", "2")
    Else
        Worksheets(LogSheetName).Cells.Clear
    End If

    Worksheets(LogSheetName).Range("A1:C1") = Split("Date|Éditeur|Édition", "|")

    ReDim CommentsTexts(ws.comments.Count - 1)
    i = 0
    For Each CellComment In ws.comments
        CommentsTexts(i) = Worksheets(ws.Name).Range(Mid(CellComment.Parent.Address, 2, 1) & "1").value _
                        & " l." & CellComment.Parent.Rows.Count & ": " & CellComment.Text
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
    'CommentsToStyle = rangewithcomment.Comment
    
    For Each cellwithcomment In rangewithcomment 'ws.comments
        Set CommentsToStyle = cellwithcomment.Cells.Comment
        With CommentsToStyle
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

