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
        Application.StatusBar = "Progression Validation: �tape (1/1) " & (counter) & " of " & UBound(FilesList) + 1 & ": " & Format(pctCompl, "percent")
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

Function CheckType(ByRef ColumnData, ByRef Lookup_expectedtype, ByVal ExpectedType As String)
    Dim Unique_ColumnData As Variant
    Dim TypeViolation As String
    Dim TypeViolationLoc

    Unique_ColumnData = GetUniqueValues(ColumnData)
    TypeViolation = CheckElementsType(Unique_ColumnData, ExpectedType)
    
    TypeViolationLoc = ""
      
    If Len(TypeViolation) > 2 Then
        'substitute a list of violation indexes for a list of violation values for next step
        For Each Violation In Split(Right(Left(TypeViolation, Len(TypeViolation) - 1), Len(TypeViolation) - 2), ",")
            If Len(Trim(Unique_ColumnData(CDbl(Violation)))) > 0 Then
                TypeViolation = Replace(TypeViolation, Violation & ",", Unique_ColumnData(Violation) & ",")
            Else
                'empty space strings can be ignored because the paste from TransferColumns() trims values by default
                TypeViolation = Replace(TypeViolation, Violation & ",", "")
            End If
        Next Violation
        
        'Fill a string with all the lines containing a type violation:
        If Len(TypeViolation) > 2 Then
            For i = LBound(ColumnData) - 1 To UBound(ColumnData) - 1
                If InStr(1, TypeViolation, "," & CStr(ColumnData(i + 1)) & ",", vbTextCompare) <> 0 Then
                    If StrComp(Right(TypeViolationLoc, 1 + Len(CStr(i))), i & ",") = 0 And Len(TypeViolationLoc) > 2 Then
                        TypeViolationLoc = Left(TypeViolationLoc, Len(TypeViolationLoc) - (2 + Len(CStr(i)))) & "-" & i + 1 & ","
                    'ElseIf StrComp(Right(TypeViolationLoc, 3), "," & i & ",") = 0 Then
                    '    TypeViolationLoc = Left(TypeViolationLoc, Len(TypeViolationLoc) - 2) & "-" & i + 1 & ","
                    Else
                        TypeViolationLoc = TypeViolationLoc & i + 1 & ","
                    End If
                End If
            Next i
            
            TypeViolationLoc = Left(TypeViolationLoc, Len(TypeViolationLoc) - 1)
            'Debug.Print TypeViolationLoc
        End If
            
    Else
        TypeViolationLoc = ""
    End If
    
    CheckType = TypeViolationLoc
End Function



Function CheckElementsType(ByRef ColumnData, ByVal ExpectedType As String) As String
    'CheckType returns a string of the indexes of the list that are not of the expected type
    'we do not return the values because the separator (here we chose ",") could be hidden in one of the values
    
    Select Case ExpectedType
        Case "NUM"
            CheckElementsType = ""
            For j = LBound(ColumnData) To UBound(ColumnData)
                If Not IsNumeric(ColumnData(j)) Then CheckElementsType = CheckElementsType & "," & j
            Next j
        Case "CHR", "CHR_NON_NUM"
            CheckElementsType = ""
            If IsNumeric(ColumnData(j)) Then CheckElementsType = CheckElementsType & "," & j
        Case "DAT"
            CheckElementsType = ""
            For j = LBound(ColumnData) To UBound(ColumnData)
                If Not IsDate(ColumnData(j)) Then CheckElementsType = CheckElementsType & "," & j
            Next j
        Case "NONE", ""
            CheckElementsType = ""
    End Select
    If Len(CheckElementsType) > 0 Then CheckElementsType = CheckElementsType & ","
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
