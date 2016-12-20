Attribute VB_Name = "B12_Validation"
Sub MainValidationLoop(ByRef FilesList)
'unused yet
    Dim wk As Workbook
    Dim pctCompl As Double
    Dim counter As Integer
    
    Application.DisplayAlerts = False
    
    counter = 1
    For Each FILE In FilesList
    
        Application.ScreenUpdating = False

        Set wk = Workbooks.Open(FileName:=FILE, corruptload:=xlRepairFile)
        wk.Windows(1).visible = False
        
        wk.Close SaveChanges:=False
        Set wk = Nothing
        
        pctCompl = (counter + 1) / (UBound(FilesList) + 1)
        Application.StatusBar = "Progression Validation: étape (1/1) " & (counter) & " of " & UBound(FilesList) + 1 & ": " & Format(pctCompl, "percent")
        Application.ScreenUpdating = True
        counter = counter + 1
    Next FILE
    
    
    
End Sub


Function ConformableFileName(FileName As String) As Boolean
                          'Filename Like "#&#_*" Or _
                          'Filename Like "#&##_*" Or _
                          'Filename Like "##&##_*" Or _

    ConformableFileName = FileName Like "#_*" Or _
                          FileName Like "##_*" Or _
                          FileName Like "[A-Z]_*" Or _
                          FileName Like "[A-Z][A-Z]_*" Or _
                          FileName Like "[A-Z]#_*" Or _
                          FileName Like "[A-Z]##_*"
    'ConformableFileName = ConformableFileName * Not Filename Like "[!0-9,A-Z]_"
End Function


Sub testConformableFileName()
    Debug.Print ConformableFileName("1_Barbay_Baud_medicaments_2015_brut.xlsx")
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

Function CheckType(ByRef ColumnData, ByRef Lookup_expectedtype, ByVal ExpectedType As String, ByVal FileNumber As Long)
    Dim Unique_ColumnData As Variant
    Dim TypeViolation As String
    Dim TypeViolationLoc As String


    If ExpectedType = "PHARMACODE" And PARAM_TABLE.Columns(1).Find("CheckPharmacodes").Offset(0, 1).value Then
        Dim n_violations As Long
        n_violations = CheckPharmacodes(ColumnData, ExpectedType, FileNumber)
        
        CheckType = ""
        Exit Function
    End If
    
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
        If Len(TypeViolation) > 2 Then ' because we previously removed the empty space strings we have to test again for TypeViolation length
            For i = LBound(ColumnData) To UBound(ColumnData)
                If InStr(1, TypeViolation, "," & CStr(ColumnData(i)) & ",", vbTextCompare) <> 0 Then
                    str_search1 = "-" & i - 1 & ","
                    str_search2 = i - 2 & "," & i - 1 & ","
                    
                    If StrComp(Right(TypeViolationLoc, Len(str_search1)), str_search1) = 0 And Len(TypeViolationLoc) > 2 Then
                        TypeViolationLoc = Left(TypeViolationLoc, Len(TypeViolationLoc) - Len(str_search1) + 1) & i & ","
                    ElseIf StrComp(Right(TypeViolationLoc, Len(str_search2)), str_search2) = 0 And Len(TypeViolationLoc) > 2 Then
                        TypeViolationLoc = Left(TypeViolationLoc, Len(TypeViolationLoc) - Len(str_search2) + Len(CStr(i))) & "-" & i & ","
                    Else
                        TypeViolationLoc = TypeViolationLoc & i & ","
                    End If
                    
                End If
            Next i
            
            'increment numbers to take into account the first row which contains the variable's name on the sheet
            TypeViolationLoc = Left(TypeViolationLoc, Len(TypeViolationLoc) - 1)
            'Debug.Print TypeViolationLoc
            s = Split(TypeViolationLoc, ",")
            For n = LBound(s) To UBound(s)
                t = Split(s(n), "-")
                For m = LBound(t) To UBound(t)
                    t(m) = t(m) + 1
                Next m
                s(n) = Join(t, "-")
            Next n
            TypeViolationLoc = Join(s, ",")
        End If
            
    Else
        TypeViolationLoc = ""
    End If
    
    CheckType = TypeViolationLoc
End Function

Function CheckPharmacodes(ByRef ColumnData, ByVal ExpectedType As String, row As Long) As Long
    CheckPharmacodes = CLng(CheckElementsType(ColumnData, ExpectedType, True))
    INTERNALS.ListObjects("file_to_load").ListColumns("invalid_pharmacodes").DataBodyRange(row).value = CheckPharmacodes
End Function

Function CheckElementsType(ByRef ColumnData, ByVal ExpectedType As String, Optional n_mode As Boolean = False) As String
    'CheckElementsType returns a string of the indexes of the list that are not of the expected type
    'we do not return the values because the separator (here we chose ",") could be hidden in one of the values
    
    'if n_mode == True , CheckElementsType returns instead the number of times it encountered a type violation,
    'in a string. it is a bad practice for a function to return different types of value, I have no excuse...
    
    Select Case ExpectedType
        Case "NUM"
            CheckElementsType = ""
            For j = LBound(ColumnData) To UBound(ColumnData)
                If Not IsNumeric(ColumnData(j)) Then CheckElementsType = CheckElementsType & "," & j
            Next j
        Case "CHR_NON_NUM"
            CheckElementsType = ""
            If IsNumeric(ColumnData(j)) Then CheckElementsType = CheckElementsType & "," & j
        Case "DAT"
            CheckElementsType = ""
            For j = LBound(ColumnData) To UBound(ColumnData)
                If Not IsDate(ColumnData(j)) Then CheckElementsType = CheckElementsType & "," & j
            Next j
        Case "PHARMACODE"
            Dim vx As MSScriptControl.ScriptControl
            Set vx = New MSScriptControl.ScriptControl

            Dim restriction As String
            Dim restrictedvalues As Variant
            
            restrictedvalues = INTERNALS.ListObjects("PharmacodeRestrictedValues").ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeConstants).Resize(INTERNALS.ListObjects("PharmacodeRestrictedValues").ListColumns(1).DataBodyRange.SpecialCells(xlCellTypeConstants).Cells.Count, 2).value
            restriction = Join2D(restrictedvalues, " ", " Or ")
            restriction = "val " & Replace(restriction, "Or", "Or val")

            With vx
                .Language = "VBScript"
                .AddCode "function stub(val): stub=" & restriction & ": end function"
            End With
            
            For j = LBound(ColumnData) To UBound(ColumnData)
                'If Not (ColumnData(j) > 0 And ColumnData(j) <= 8999999 And ColumnData(j) <> 8888887 And ColumnData(j) <> "") Then CheckElementsType = CheckElementsType & "," & j
                If ColumnData(j) = "" Then
                    CheckElementsType = CheckElementsType & "," & j
                Else
                    If vx.Run("stub", CLng(ColumnData(j))) Then CheckElementsType = CheckElementsType & "," & j
                End If
            Next j
        Case "CHR", "NONE", ""
            CheckElementsType = ""
    End Select
    
    If n_mode Then
        CheckElementsType = UBound(Split(CheckElementsType, ","))
        If CheckElementsType = "-1" Then CheckElementsType = "0"
    Else
        If Len(CheckElementsType) > 0 Then CheckElementsType = CheckElementsType & ","
    End If
        
End Function

Function AssertStatus()
    
End Function

Function CheckForSpecialCharacters(sh As Worksheet)
    'Ne fonctionne pas encore, il faut changer la condition du like
    Dim r As Range
    Dim rangetoscoop As Range
    Set rangetoscoop = sh.Range(sh.Cells(2, 1), _
                        sh.Cells(sh.Cells(sh.Rows.Count, 1).End(xlUp).row, _
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


Public Function Join2D(ByVal vArray As Variant, Optional ByVal sWordDelim As String = " ", Optional ByVal sLineDelim As String = vbNewLine) As String
    
    Dim i As Long, j As Long
    Dim aReturn() As String
    Dim aLine() As String
    
    ReDim aReturn(LBound(vArray, 1) To UBound(vArray, 1))
    ReDim aLine(LBound(vArray, 2) To UBound(vArray, 2))
    
    For i = LBound(vArray, 1) To UBound(vArray, 1)
        For j = LBound(vArray, 2) To UBound(vArray, 2)
            'Put the current line into a 1d array
            aLine(j) = vArray(i, j)
        Next j
        'Join the current line into a 1d array
        aReturn(i) = Join(aLine, sWordDelim)
    Next i
    
    Join2D = Join(aReturn, sLineDelim)
    
End Function
