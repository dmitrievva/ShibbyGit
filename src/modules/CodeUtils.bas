Attribute VB_Name = "CodeUtils"
'***********************************************************************
' Original Author:   Eric Addison
' Link:     https://github.com/ericaddison/ShibbyGit
'
' Changed by: Vladimir Dmitriev, https://github.com/dmitrievva/ShibbyGit
'***********************************************************************

' Any functions to help with the actual coding process
Option Explicit

' component type constants
Public Const Module As Integer = 1
Public Const ClassModule As Integer = 2
Public Const Form As Integer = 3
Public Const Document As Integer = 100
Public Const padding As Integer = 24

Private pFolder As String

'****************************************************
' Public functions
'****************************************************

' public interface for export all, no msg box
' input: folder - the folder to export code modules to
' output: String with list of modules exported
Public Function ExportAllString(ByVal folder As String) As String
    pFolder = folder
    If Not FileUtils.FileOrDirExists(folder) Then
        ExportAllString = "Invalid folder: " & folder
        Exit Function
    End If
    ExportAllString = ExportAll
End Function


' public interface for import, no msg box
' input: folder - the folder to import code modules from
' output: String with list of modules imported
Public Function ImportSelectedString(ByVal files As FileDialogSelectedItems) As String
    ImportSelectedString = ImportSelected(files)
End Function


' Find the index of the VBProject corresponding to
' the active presentation
' output: the index of the VBProject corresponding to the active filename
'           -1 if not found
Public Function FindFileVBProject(Optional ByVal fileName As String = "") As Integer
    If fileName = "" Then
        fileName = ShibbySettings.GetProjectFileName
    End If

    With Application
        Dim ind As Integer
        For ind = 1 To .VBE.VBProjects.count
            Dim VBFileName As String
            On Error Resume Next
                VBFileName = .VBE.VBProjects.Item(ind).fileName
            On Error GoTo 0
            If VBFileName = fileName Then
                FindFileVBProject = ind
                Exit Function
            End If
        Next ind
    End With
    
    MsgBox "Could not find VB Project associated with open file: " & fileName _
            & vbCrLf & "Is this a new, unsaved document?"
    FindFileVBProject = -1
End Function


' remove a code module and import
' input: projectInd - the index of the desired VBProject in Application.VBE.VBProjects
' input: file - the full path to the file for import
' output: the name of the imported module, or "" if none
Public Function RemoveAndImportModule(ByVal projectInd As Integer, ByVal file As String) As String
    With Application.VBE.VBProjects.Item(projectInd).VBComponents
        If CheckCodeType(file) <> -1 Then
            Dim ModuleName As String
            ModuleName = FileBaseName(file)
            
            ' don't import modules with running code!
            If ModuleName = "NonModalMsgBoxForm" Or ModuleName = "CodeUtils" Then
                Exit Function
            End If
            
            ' check if module already exists in project
            Dim moduleExists As Boolean
            On Error Resume Next
                .Item (ModuleName)
                moduleExists = (Err = 0)
                Err.Clear
            On Error GoTo 0
            
            ' rename and remove
            If moduleExists Then
                If CheckCodeType(file) = 3 Then
                    .Remove .Item(ModuleName)
                Else
                    .Item(ModuleName).name = ModuleName & "R"
                    .Remove .Item(ModuleName & "R")
                End If
                DoEventsAndWait 10, 2
            End If
            
            ' import new module
            Dim newModule As Variant
            Set newModule = .Import(file)
            RemoveAndImportModule = newModule.name
        End If
    End With
End Function


' return a Module type based on the file extension
' input: file - filename of a code module
' output: integer corresponding to module type
Public Function CheckCodeType(ByVal file As String) As Integer

    If file Like "*.bas" Then
        CheckCodeType = Module
    ElseIf file Like "*.frm" Then
        CheckCodeType = Form
    ElseIf file Like "*.cls" Then
        CheckCodeType = ClassModule
    Else
        CheckCodeType = -1
    End If

End Function

' import all code modules from the given folder to the VBProject with the given index
Public Function ImportCodeFromFolder(ByVal folder As String, ByVal projectInd As Integer) As String
    Dim file As String
    Dim ModuleName As String
    Dim filesRead As String
    file = Dir(folder & "\")
    While file <> ""
        ModuleName = RemoveAndImportModule(projectInd, folder & "\" & file)
        If ModuleName <> "" Then
            filesRead = filesRead & vbCrLf & ModuleName
        End If
        file = Dir
    Wend
    ImportCodeFromFolder = filesRead
End Function

' open selected file in Visual Basic Editor
Public Sub EditSelectedFile(ByVal file As String)
    Dim oVBComponent    As Object
    Dim fileName        As String
    Dim found           As Boolean
    
    fileName = Split(file, ".")(0)
    
    On Error Resume Next
    found = False
    For Each oVBComponent In ActiveWorkbook.VBProject.VBComponents
        If UCase(oVBComponent.name) = UCase(fileName) Then
            oVBComponent.Activate
            found = True
            Exit For
        End If
    Next
    
    If Not found Then
        MsgBox "File '" & file & "' not found!"
    End If
    
    Set oVBComponent = Nothing
End Sub


'****************************************************
' Private functions
'****************************************************

Private Function ExportAll() As String

    ' write files
    Dim projectInd As Integer
    projectInd = FindFileVBProject
    If projectInd = -1 Then
        ExportAll = "Uh oh! Could not find VBProject associated with " & ShibbySettings.GetProjectName
        Exit Function
    End If
    
    With Application.VBE.VBProjects.Item(projectInd).VBComponents
        Dim ind As Integer
        Dim filesWritten As String
        Dim extension As String
        For ind = 1 To .count
            extension = ""
            Select Case .Item(ind).Type
               Case ClassModule
                   extension = ".cls"
               Case Form
                   extension = ".frm"
               Case Module
                   extension = ".bas"
            End Select

            If (extension <> "") Then
                .Item(ind).Export (pFolder & "\" & .Item(ind).name & extension)
                filesWritten = filesWritten & vbCrLf & .Item(ind).name & extension
            End If
        Next ind
    
    End With
     
    ' clean up frx forms if requested
    If ShibbySettings.FrxCleanup Then
        GitProject.RemoveUnusedFrx
    End If
    
    ' return list of exported files
    ExportAll = "ShibbyGit: " & vbCrLf & "Code Exported to " & pFolder & vbCrLf & filesWritten

End Function


Private Function ImportSelected(ByVal files As FileDialogSelectedItems) As String

    ' get project index from active file name
    Dim projectInd As Integer
    projectInd = FindFileVBProject
    If projectInd = -1 Then
        ImportSelected = "Uh oh! Could not find VBProject associated with " & ShibbySettings.GetProjectName
        Exit Function
    End If

    ' import files
    Dim file As Variant
    Dim ModuleName As String
    Dim filesRead As String
    For Each file In files
        ModuleName = RemoveAndImportModule(projectInd, file)
        If ModuleName <> "" Then
            filesRead = filesRead & vbCrLf & ModuleName
        End If
    Next file


    ImportSelected = "ShibbyGit Modules Loaded: " & filesRead

End Function


Private Sub test()
'    Dim proc As New Process
'    proc.StartInfo.fileName = "cmd.exe"
'    proc.StartInfo.Arguments = "/k ipconfig"
'    proc.StartInfo.CreateNoWindow = True
'    proc.StartInfo.UseShellExecute = False
'    proc.StartInfo.RedirectStandardOutput = True
'    proc.Start()
'    proc.WaitForExit()
'
'    Dim output() As String = proc.StandardOutput.ReadToEnd.Split(CChar(vbLf))
'    For Each ln As String In output
'        RichTextBox1.AppendText (ln & vbNewLine)
'    Next
End Sub
