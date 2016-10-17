Attribute VB_Name = "B2_Validation"
Sub MainValidationLoop(ByRef FilesList)
'unused yet
    Dim wk As Workbook
    Dim pctCompl As Double
    Dim counter As Integer
    
    Application.DisplayAlerts = False
    
    counter = 1
    For Each FILE In FilesList
    
        Application.ScreenUpdating = False

        Set wk = Workbooks.Open(Filename:=FILE, corruptload:=xlRepairFile)
        wk.Windows(1).Visible = False
        
        wk.Close SaveChanges:=False
        Set wk = Nothing
        
        pctCompl = (counter + 1) / (UBound(FilesList) + 1)
        Application.StatusBar = "Progression Validation: étape (1/1) " & (counter) & " of " & UBound(FilesList) + 1 & ": " & Format(pctCompl, "percent")
        Application.ScreenUpdating = True
        counter = counter + 1
    Next FILE
End Sub

Function CheckColumnNames(Range1 As Range)
    Dim a As Application
    Set a = Application
    
    Dim RefRange As Range
    
    For Each system In INTERNALS.ListObjects("sys_info_attributes").ListRows
        Set RefRange = system.Range.Offset(0, 1).Resize(1, system.Range.Columns.Count - 1)
        CheckColumnNames = Join(a.Transpose(a.Transpose(Range1.value)), Chr(0)) = _
                           Join(a.Transpose(a.Transpose(RefRange.value)), Chr(0))
        If CheckColumnNames Then
            CheckColumnNames = system.Range(1).value
            Exit Function
        End If
    Next system
    CheckColumnNames = "Error"
End Function

Function CheckType(ByRef ColumnData, ByVal ExpectedType As String)
    Dim Unique_ColumnData As Variant
    Dim TypeViolation As String
    Dim TypeViolationLoc

    Unique_ColumnData = GetUniqueValues(Data)
    TypeViolation = CheckElementsType(Unique_ColumnData, Lookup_expectedtype(curr_col_num))
    
    
    
    TypeViolationLoc = "line:"
    
    If Len(TypeViolation) > 0 Then
            'substitute a list of violation indexes for a list of violation values
            For Each Violation In Split(Right(TypeViolation, Len(TypeViolation) - 1), ",")
                TypeViolation = Replace(TypeViolation, Violation & ",", Unique_ColumnData(Violation) & ",")
            Next Violation
    
            For i = LBound(ColumnData) To UBound(ColumnData)
                If "*" & ColumnData(i) & "*" Like TypeViolation Then
                    
                    'check if i-1 also present to not congest the string
                    
                    'TypeViolationLoc = TypeViolationLoc & i
                End If
            Next i
        
    Else
        TypeViolationLoc = "FAUX"
    End If
    
    CheckType = TypeViolationLoc
End Function



Function CheckElementsType(ByRef ColumnData, ByVal ExpectedType As String) As String
    'CheckType returns a string of the indexes of the list that are not of the expected type
    'we do not return the values because the separator (here we chose ",") could be hidden in one of the values
    
    Select Case ExpectedType
        Case "NUM"
            CheckType = ""
            For j = LBound(ColumnData) To UBound(ColumnData)
                If Not IsNumeric(ColumnData(j)) Then CheckType = CheckType & "," & j
            Next j
        Case "CHR", "CHR_NON_NUM"
            CheckType = ""
            If IsNumeric(ColumnData(j)) Then CheckType = CheckType & "," & j
        Case "DAT"
            CheckType = ""
            For j = LBound(ColumnData) To UBound(ColumnData)
                If Not IsDate(ColumnData(j)) Then CheckType = CheckType & "," & j
            Next j
        Case "NONE", ""
            CheckType = ""
    End Select
    If Len(CheckType) > 0 Then CheckType = Left(CheckType, Len(CheckType) - 1) & ","
End Function






Function CheckForSpecialCharacters(sh As Worksheet)
    'Ne fonctionne pas encore, il faut changer la condition du like
    Dim r As Range
    Dim rangetoscoop As Range
    Set rangetoscoop = sh.Range(sh.Cells(2, 1), _
                        sh.Cells(sh.Cells(sh.Rows.Count, 1).End(xlUp).Row, _
                        sh.Cells(1, sh.Columns.Count).End(xlToLeft).column))
    For Each r In rangetoscoop
        If r.value Like "*[!0-9,a-z,A-Z,.,/]*" Then
            'r.Font.Color = vbRed
            r.Interior.ColorIndex = 46
            r.Font.Bold = True
            Debug.Print r.Address & ": " & r.value
        End If
    Next
End Function


'Sub test()
'    Call CheckForSpecialCharacters(Workbooks(2).Worksheets(1))
'End Sub
