Attribute VB_Name = "B02_SheetUtilities"
Sub SaveFilesList(ByRef FilesList)
    Dim table As ListObject
    Dim Path As Range
    Dim r As Range
    
    'INTERNALS.Unprotect "mdp"
    
    Set table = INTERNALS.ListObjects("file_to_load")
    Set Path = INTERNALS.ListObjects("path").ListColumns("path").DataBodyRange
    
    Path(1).value = Left(FilesList(0), InStrRev(FilesList(0), "\"))
    
    table.DataBodyRange.ClearContents
    Set r = table.Range.Rows(2).Offset(-1, 0).Resize(3)
    
    table.Resize r
    For i = LBound(FilesList) To UBound(FilesList)
        table.ListColumns(1).DataBodyRange(i + 1) = i + 1
        table.ListColumns(2).DataBodyRange(i + 1) = Right(FilesList(i), Len(FilesList(i)) - InStrRev(FilesList(i), "\"))
    Next i
    'INTERNALS.Protect "mdp"
End Sub

'Function GetFilesList()
'    Dim table As ListObject
'    Dim path As Range
'    Dim FilesList As Variant
'    Set table = INTERNALS.ListObjects("file_to_load")
'    Set path = INTERNALS.ListObjects("path").ListColumns("path").DataBodyRange
'
'    For i = 1 To table.ListRows.Count '
'        FilesList(i - 1) = path(1).value & table.ListColumns(2).DataBodyRange(i).value
'    Next i
'    GetFilesList = FilesList
'End Function

Private Function StringArrayToIntArray(ByRef values()) As Integer()
    Dim lIndx As Long, lLwrBnd As Long, lUprBnd As Long
    Dim adRtnVals() As Integer
    lLwrBnd = LBound(values)
    lUprBnd = UBound(values)
    ReDim adRtnVals(lLwrBnd To lUprBnd) As Integer
    For lIndx = lLwrBnd To lUprBnd
        adRtnVals(lIndx) = CInt(values(lIndx))
    Next
    StringArrayToIntArray = adRtnVals
End Function

Sub testIncCol()
    Debug.Print IncCol("A", 26)
End Sub

Function IncCol(ByVal column As String, ByVal IncrementStep As Integer) As String
    
    Dim reminder As Double
    Dim nloops As Long
    Dim IncrementColumn As String
    reminder = Asc(column) + IncrementStep
    nloops = 1
    If (Asc(column) + IncrementStep) < Asc("Z") Then
        IncCol = Chr(Asc(column) + IncrementStep)
    Else
        IncrementColumn = ""
        Do While (reminder > Asc("Z"))
            reminder = reminder - (Asc("Z") - Asc("A"))
            IncrementColumn = IncrementColumn & "A"
            nloops = nloops + 1
            If nloops > 10000 Then
                Exit Do
            End If
        Loop
        IncCol = IncrementColumn & Chr(reminder)
    End If
End Function

Function GetUniqueValues(ByRef DATA)
    Dim temp As Variant
    Dim obj As Object
    Set obj = CreateObject("scripting.dictionary")
    For i = LBound(DATA) To UBound(DATA)
        obj(DATA(i) & "") = ""
    Next
    GetUniqueValues = obj.keys
End Function

Public Function NumericOnly(ByVal s As String) As String
    Dim s2 As String
    Dim replace_hyphen As String
 
    Static re As RegExp
    If re Is Nothing Then Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "[^0-9]" 'includes space, if you want to exclude space "[^0-9]"
    's2 = re.Replace(s, vbNullString)

    NumericOnly = re.Replace(s2, replace_hyphen)
End Function



Public Function SheetExists(worksheetName As String)
    SheetExists = Evaluate("ISREF('" & worksheetName & "'!A1)")
End Function
