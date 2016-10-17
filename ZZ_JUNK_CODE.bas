Attribute VB_Name = "ZZ_JUNK_CODE"



Sub notasub()
    Dim canton_code As String
    Dim list_canton_name As Variant
    
    list_canton_name = INTERNALS.ListObjects("cantons").ListColumns("canton_name").DataBodyRange.value
    canton_code = Application.Match(Canton.value, list_canton_name, 0)
    Debug.Print canton_code

End Sub



