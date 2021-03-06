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
    
    'INTERNALS.Protect "mdp", UserInterFaceOnly:=True
    ACCUEIL.Activate
    'DisplayTag = "*"
    'Call OpenGit
    
End Sub

'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Source As Range)
' runs when a sheet is changed
'End Sub

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
'opening git:
    Dim x As Variant
    Dim Path As String
    Path = "C:\Users\akunzi\Desktop\PortableGit\git-bash.exe"
    
    Application.Wait Now + TimeValue("00:00:05")
    x = Shell(Path, vbNormalFocus)
    Application.Wait Now + TimeValue("00:00:03")
    AppActivate Title:="MINGW64" ', Wait:=True
    SendKeys String:="set_path~", Wait:=True
    SendKeys String:="git status~", Wait:=True
End Sub
