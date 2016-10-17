Attribute VB_Name = "C1_Style"
Sub ApplyStyle(srange As Range, cond As String, oper As String)

    If InStr(oper, "xlGreater") <> 0 Then
        op = xlGreater
    ElseIf InStr(oper, "xlEqual") <> 0 Then
        op = xlEqual
    ElseIf InStr(oper, "xlNotEqual") <> 0 Then
        op = xlNotEqual
    End If
    
        srange.FormatConditions.Add Type:=xlCellValue, Operator:=op, _
        Formula1:=cond
    srange.FormatConditions(srange.FormatConditions.Count).SetFirstPriority
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
End Sub

Sub FitComments()
'source: https://www.extendoffice.com/documents/excel/1572-excel-autosize-comments.html#a1
'Updateby20140325
Dim xComment As Comment
For Each xComment In Application.ActiveSheet.Comments
    xComment.Shape.TextFrame.AutoSize = True
Next
End Sub
