VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Call DefGlobal
    Year.value = Format(Now, "yyyy") - 1
    Call A_0.SetValidation
    'MsgBox ""
    Call UpdateStage("Start")
End Sub



Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Call ExportAllComponents
End Sub


Public Sub ExportAllComponents()
    Dim VBComp As VBIDE.VBComponent
    Dim destDir As String, fName As String, ext As String
    'Create the directory where code will be created.
    'Alternatively, you could change this so that the user is prompted
    If ActiveWorkbook.Path = "" Then
        MsgBox "You must first save this workbook somewhere so that it has a path.", , "Error"
        Exit Sub
    End If
    destDir = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name & " Modules"
    If Dir(destDir, vbDirectory) = vbNullString Then MkDir destDir
    
    'Export all non-blank components to the directory
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
        If VBComp.CodeModule.CountOfLines > 0 Then
            'Determine the standard extention of the exported file.
            'These can be anything, but for re-importing, should be the following:
            Select Case VBComp.Type
                Case vbext_ct_ClassModule: ext = ".cls"
                Case vbext_ct_Document: ext = ".cls"
                Case vbext_ct_StdModule: ext = ".bas"
                Case vbext_ct_MSForm: ext = ".frm"
                Case Else: ext = vbNullString
            End Select
            If ext <> vbNullString Then
                fName = destDir & "\" & VBComp.Name & ext
                'Overwrite the existing file
                'Alternatively, you can prompt the user before killing the file.
                If Dir(fName, vbNormal) <> vbNullString Then Kill (fName)
                VBComp.Export (fName)
            End If
        End If
    Next VBComp
End Sub

Sub OpenGit()
'execute with ctrl+g
'opening git:
    Dim x As Variant
    Dim Path As String

    Path = "C:\Users\akunzi\Desktop\PortableGit\git-bash.exe"

    x = Shell(Path, vbNormalFocus)
    
    'Application.Wait Now + TimeValue("00:00:03") ' waits 3 seconds
    'AppActivate x
    'SendKeys "setdir_p"
    
End Sub
