Attribute VB_Name = "C1_Style"
Sub ApplyStyle(srange As Range, cond As String, oper As String, style As String)

    If InStr(oper, "xlGreater") <> 0 Then
        op = xlGreater
    ElseIf InStr(oper, "xlEqual") <> 0 Then
        op = xlEqual
    ElseIf InStr(oper, "xlNotEqual") <> 0 Then
        op = xlNotEqual
    'ElseIf InStr(oper, "xlExpression") <> 0 Then
    '    op = xlExpression
    End If
    
    'If op = xlExpression Then
    '    srange.FormatConditions.Add Type:=xlExpression, Formula1:= _
    '    "=NBCAR(SUPPRESPACE(" & srange(1).Address & "))=0"
    'Else
        srange.FormatConditions.Add Type:=xlCellValue, Operator:=op, _
        Formula1:=cond
    'End If
    
    srange.FormatConditions(srange.FormatConditions.Count).SetFirstPriority
    
    If InStr(style, "bad") <> 0 Then
        With srange.FormatConditions(1).Font
            .Color = -16383844
            .TintAndShade = 0
            .Bold = True
        End With
        With srange.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        
    ElseIf InStr(style, "good") <> 0 Then
        With srange.FormatConditions(1).Font
            .Color = -16752384
            .TintAndShade = 0
        End With
        With srange.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13561798
            .TintAndShade = 0
        End With
    ElseIf InStr(style, "warning") <> 0 Then
        With srange.FormatConditions(1).Font
            .Color = -16751204
            .TintAndShade = 0
        End With
        With srange.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 10284031
            .TintAndShade = 0
        End With
    End If
    srange.FormatConditions(1).StopIfTrue = False
    
End Sub

Sub FitComments()
'source: https://www.extendoffice.com/documents/excel/1572-excel-autosize-comments.html#a1
'Updateby20140325
Dim xComment As Comment
For Each xComment In Application.ActiveSheet.comments
    xComment.Shape.TextFrame.AutoSize = True
Next
End Sub

Sub test()
    Call ResizeColumnWidths(8, 4, 3)
End Sub

Sub ResizeColumnWidths(n, startcolumn As Integer, ignorecolumn As Integer)
  ignorewidth = Range(IncCol("A", ignorecolumn - 1) & "1").ColumnWidth
  w = (ActiveWindow.Width - ignorewidth) / n
  Columns(startcolumn).Resize(, n).ColumnWidth = pointsToChars(w, startcolumn)
End Sub

Function pointsToChars(x, startcolumn) As Integer
'source: http://stackoverflow.com/questions/28439773/excel-vba-fit-columns-to-page

  p = Range(IncCol("A", startcolumn - 1 + 1) & "1").Width
  c = Range(IncCol("A", startcolumn - 1 + 1) & "1").ColumnWidth
  pointsToChars = x * c / p
End Function
