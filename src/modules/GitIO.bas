Attribute VB_Name = "GitIO"
'***********************************************************************
' Original Author:   Eric Addison
' Link:     https://github.com/ericaddison/ShibbyGit
'
' Changed by: Vladimir Dmitriev, https://github.com/dmitrievva/ShibbyGit
'***********************************************************************

Option Explicit
Private Const MODULEFOLDER As String = "modules"
Private Const CLASSFOLDER As String = "classModules"
Private Const FORMFOLDER As String = "forms"
Private Const SOURCEFOLDER As String = "src"
Private pFileStructure As ShibbyFileStructure
Private pGitDir As String
Private pProjectInd As Integer


'****************************************************
' Public functions
'****************************************************

' Public entry point for Git Import
Public Function GitImport(ByVal gitDir As String, ByVal fileStructure As ShibbyFileStructure) As String
    pFileStructure = fileStructure
    pGitDir = gitDir
    GitImport = GitImportAll
End Function


' Public entry point for Git Export
Public Function GitExport(ByVal gitDir As String, ByVal fileStructure As ShibbyFileStructure) As String
    pFileStructure = fileStructure
    pGitDir = gitDir
    GitExport = GitExportAll
End Function


'****************************************************
' Private functions
'****************************************************

' check that the incoming folder is valid
Private Function CheckGitFolder() As Boolean
    CheckGitFolder = True
    ' check the incoming folder
    Dim folderCheck As String
    folderCheck = FileUtils.VerifyFolder(pGitDir, "Please select Git project folder")
    If folderCheck = FileUtils.BAD_FOLDER Then
        CheckGitFolder = False
        Exit Function
    ElseIf folderCheck <> FileUtils.GOOD_FOLDER Then
        pGitDir = folderCheck
        ShibbySettings.GitProjectPath = pGitDir
    End If
End Function


' Import all code modules from git directory
' based on the selected file structure
Private Function GitImportAll() As String

    If Not CheckGitFolder Then
        Exit Function
    End If
    
    ' project ind
    pProjectInd = CodeUtils.FindFileVBProject
    If pProjectInd = -1 Then
        GitImportAll = "Uh oh! Could not find VBProject associated with " & ShibbySettings.GetProjectName
        Exit Function
    End If
    
    ' import files
    Dim filesRead As String
    If pFileStructure = flat Then
        filesRead = CodeUtils.ImportCodeFromFolder(pGitDir, pProjectInd)
    ElseIf pFileStructure = SimpleSrc Then
        filesRead = CodeUtils.ImportCodeFromFolder(pGitDir & "\" & SOURCEFOLDER, pProjectInd)
    ElseIf pFileStructure = SeparatedSrc Then
        filesRead = CodeUtils.ImportCodeFromFolder(pGitDir & "\" & SOURCEFOLDER & _
            "\" & MODULEFOLDER, pProjectInd)
        filesRead = filesRead & CodeUtils.ImportCodeFromFolder(pGitDir & "\" & SOURCEFOLDER & _
            "\" & FORMFOLDER, pProjectInd)
        filesRead = filesRead & CodeUtils.ImportCodeFromFolder(pGitDir & "\" & SOURCEFOLDER & _
            "\" & CLASSFOLDER, pProjectInd)
    End If
    
    If filesRead = "" Then
        filesRead = "<No files imported>"
    End If
    GitImportAll = "ShibbyGit Modules Loaded: " & filesRead

End Function



' Export all code modules to git directory
' based on the selected file structure
Private Function GitExportAll() As String
    
    If Not CheckGitFolder Then
        Exit Function
    End If
    
    ' remove files before export if needed
    RemoveFilesBeforeExport
    
    ' create folders if needed
    CheckCodeFolders
    
    ' write files
    pProjectInd = CodeUtils.FindFileVBProject
    If pProjectInd = -1 Then
        GitExportAll = "Uh oh! Could not find VBProject associated with " & ShibbySettings.GetProjectName
        Exit Function
    End If
    
    Dim compInd As Integer
    Dim filesWritten As String
    Dim nextFile As String
    Dim nComps As Integer
    nComps = Application.VBE.VBProjects.Item(pProjectInd).VBComponents.count
    
    For compInd = 1 To nComps
        nextFile = ExportToProperFolder(compInd)
        filesWritten = filesWritten & vbCrLf & nextFile
    Next compInd
     
    ' clean up frx forms if requested
    If ShibbySettings.FrxCleanup Then
        GitProject.RemoveUnusedFrx
    End If
    
    ' return list of exported files
    If filesWritten = "" Then
        filesWritten = "<No files exported>"
    End If
    GitExportAll = "ShibbyGit: " & vbCrLf & "Code Exported to " & pGitDir & vbCrLf & filesWritten

End Function


' return the correct file extension based on the type of module
' module type constants defined in CodeUtils
Private Function GetExtensionFromModuleType(ByVal codeType As Integer)
    Dim extension As String
    Select Case codeType
       Case CodeUtils.ClassModule
           extension = ".cls"
       Case CodeUtils.Form
           extension = ".frm"
       Case CodeUtils.Module
           extension = ".bas"
    End Select
    GetExtensionFromModuleType = extension
End Function


' export one module to the proper directory
' input: compInd - the index of the desired component in project.VBComponents.Item(pProjectInd)
' output: the path of the output file, relative to pGitDir
Private Function ExportToProperFolder(ByVal compInd As Integer)
    With Application.VBE.VBProjects.Item(pProjectInd).VBComponents.Item(compInd)
        
        Dim extension As String
        extension = GetExtensionFromModuleType(.Type)

        If (extension <> "") Then
            Dim file As String
            file = SOURCEFOLDER & "\"
            
            ' flat file structure
            If pFileStructure = flat Then
                file = .name & extension
                .Export (pGitDir & "\" & file)
                
            ' simple source folder structure
            ElseIf pFileStructure = SimpleSrc Then
                file = file & .name & extension
                .Export (pGitDir & "\" & file)
                
            ' separated source folder structure
            Else
                Select Case .Type
                    Case ClassModule
                        file = file & CLASSFOLDER
                    Case Form
                        file = file & FORMFOLDER
                    Case Module
                        file = file & MODULEFOLDER
                End Select
                
                file = file & "\" & .name & extension

                .Export (pGitDir & "\" & file)
            End If
         End If
    End With
    ExportToProperFolder = file
End Function


' Check for existence of required code folders based
' on the file structure type. Create if necessary
Private Sub CheckCodeFolders()
    ' create folders if needed
    If pFileStructure <> flat Then
        If Not FileUtils.FileOrDirExists(pGitDir & "\" & SOURCEFOLDER & "\") Then
            MkDir pGitDir & "\" & SOURCEFOLDER & "\"
        End If
        If pFileStructure = SeparatedSrc Then
            If Not FileUtils.FileOrDirExists(pGitDir & "\" & SOURCEFOLDER & "\" & MODULEFOLDER & "\") Then
                MkDir pGitDir & "\" & SOURCEFOLDER & "\" & MODULEFOLDER & "\"
            End If
            If Not FileUtils.FileOrDirExists(pGitDir & "\" & SOURCEFOLDER & "\" & FORMFOLDER & "\") Then
                MkDir pGitDir & "\" & SOURCEFOLDER & "\" & FORMFOLDER & "\"
            End If
            If Not FileUtils.FileOrDirExists(pGitDir & "\" & SOURCEFOLDER & "\" & CLASSFOLDER & "\") Then
                MkDir pGitDir & "\" & SOURCEFOLDER & "\" & CLASSFOLDER & "\"
            End If
        End If
    End If
End Sub

' Remove files before export
Sub RemoveFilesBeforeExport()
    Dim folderPath  As String
    Dim strFile     As String
    
    If Not ShibbySettings.RemoveFiles Then Exit Sub
    
    folderPath = pGitDir & "\"
    
    strFile = Dir(folderPath)
    Do While Len(strFile) > 0
        If Not Strings.left(strFile, 4) = ".git" Then
            Kill folderPath & strFile
        End If
        strFile = Dir
    Loop
    
End Sub
