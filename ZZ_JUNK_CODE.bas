Attribute VB_Name = "ZZ_JUNK_CODE"



Sub notasub()
    Dim canton_code As String
    Dim list_canton_name As Variant
    
    list_canton_name = INTERNALS.ListObjects("cantons").ListColumns("canton_name").DataBodyRange.value
    canton_code = Application.Match(Canton.value, list_canton_name, 0)
    Debug.Print canton_code

End Sub


Sub chart()
    Debug.Print Asc("Þ")
End Sub


Public Sub ShowTable()
'max 27 lines!!!
Dim myData
Dim myStr As String
Dim x As Integer
Dim myRange As Range

Set myRange = Range("D1:E28")

myData = myRange.value

For x = 1 To UBound(myData, 1)
    myStr = myStr & myData(x, 1) & vbTab & myData(x, 2) & vbCrLf
Next x

MsgBox myStr, vbYesNoCancel

End Sub
