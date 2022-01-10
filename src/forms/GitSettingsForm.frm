VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitSettingsForm 
   Caption         =   "ShibbyGit Settings"
   ClientHeight    =   7845
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   8580
   OleObjectBlob   =   "GitSettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
' Original Author:   Eric Addison
' Link:     https://github.com/ericaddison/ShibbyGit
'
' Changed by: Vladimir Dmitriev, https://github.com/dmitrievva/ShibbyGit
'***********************************************************************

Option Explicit


Private needGitUserNameUpdate As Boolean
Private needGitUserEmailUpdate As Boolean


'****************************************************************
' initialize

Public Sub ResetForm()
    ' set the gitExe path text
    GitExeTextBox.text = ShibbySettings.GitExePath
    
    ' set the project path text
    ProjectPathTextBox.text = ShibbySettings.GitProjectPath
    
    If GitExeTextBox.text <> "" Then
        ' set the username field
        
        Dim UserName As String
        
        UserName = ShibbySettings.UserName
        If UserName = "" Then
            If ProjectPathTextBox.text = "" Then
                UserName = GitCommands.RunGitAsProcess("config user.name", UseProjectPath:=False)
            Else
                UserName = GitCommands.RunGitAsProcess("config user.name")
            End If
            If Len(UserName) > 0 Then
                UserName = left(UserName, Len(UserName) - 1)
            End If
        End If
        UserNameBox.value = UserName
        
        ' set the email field
        Dim UserEmail As String
        
        UserEmail = ShibbySettings.UserEmail
        If UserEmail = "" Then
            If ProjectPathTextBox.text = "" Then
                UserEmail = GitCommands.RunGitAsProcess("config user.email", UseProjectPath:=False)
            Else
                UserEmail = GitCommands.RunGitAsProcess("config user.email")
            End If
            If Len(UserEmail) > 0 Then
                UserEmail = left(UserEmail, Len(UserEmail) - 1)
            End If
        End If

        UserEmailBox.value = UserEmail
    End If
    
    ' set the frx box value
    FrxCleanupBox.value = ShibbySettings.FrxCleanup
    
    ' set the export on git box value
    ExportOnGitBox.value = ShibbySettings.ExportOnGit
    
    ' set the remove files before export value
    RemoveFilesBox.value = ShibbySettings.RemoveFiles
    
    ' Add items to the file structure box
    FileStructureBox.AddItem "Flat File Stucture"
    FileStructureBox.AddItem "Simple Src Structure"
    FileStructureBox.AddItem "Separated Src Structure"
    
    Dim fsIndex As ShibbyFileStructure
    fsIndex = ShibbySettings.fileStructure
    FileStructureBox.ListIndex = fsIndex
    
    needGitUserNameUpdate = False
    needGitUserEmailUpdate = False
    
End Sub

'****************************************************************
' component callbacks

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SaveAndHide
    End If
End Sub

Private Sub CancelButton_Click()
    GitSettingsForm.Hide
End Sub

Private Sub OKButton_Click()
    SaveAndHide
End Sub

Private Sub SaveAndHide()
    SaveGitExe
    SaveProjectPath
    SaveUserName
    SaveUserEmail
    SaveFrxCleanup
    SaveExportOnGit
    SaveFileStructure
    SaveRemoveFilesBeforeExport
    
    GitSettingsForm.Hide
End Sub

Private Sub UserEmailBox_Change()
    needGitUserEmailUpdate = True
End Sub

Private Sub UserNameBox_Change()
    needGitUserNameUpdate = True
End Sub


Private Sub GitExeBrowseButton_Click()
    GitExeTextBox.text = FileUtils.FileBrowser("Browser for git.exe")
End Sub


Private Sub ProjectPathBrowseButton_Click()
    ProjectPathTextBox.text = FileUtils.FolderBrowser("Browse for Git project folder")
End Sub


'****************************************************************
' save methods

' Save the project path as a document property
Private Sub SaveProjectPath()
    Dim newPath As String
    newPath = ProjectPathTextBox.text
    
    If newPath <> "" And FileUtils.FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find file: " & newPath
        Exit Sub
    End If

    'save this one in the registry
    ShibbySettings.GitProjectPath = newPath
End Sub


' save the gitExe path as a registry property
Private Sub SaveGitExe()
    Dim newPath As String
    newPath = GitExeTextBox.text
    
    If newPath <> "" And FileUtils.FileOrDirExists(newPath) = False Then
        MsgBox "Cannot find file: " & newPath
        Exit Sub
    End If

    'save this one in the registry
    ShibbySettings.GitExePath = newPath
End Sub

' save the user email to the git repo
Private Sub SaveUserEmail()
    Dim email       As String
    
    email = UserEmailBox.value
    
    If needGitUserEmailUpdate Then
        GitCommands.RunGitAsProcess ("config --local user.email """ & email & """")
    End If
    needGitUserEmailUpdate = False
    
    ShibbySettings.UserEmail = email
End Sub


' save the user name to the git repo
Private Sub SaveUserName()
    Dim name        As String
    
    name = UserNameBox.value
    
    If needGitUserNameUpdate Then
        GitCommands.RunGitAsProcess ("config --local user.name """ & name & """")
    End If
    needGitUserNameUpdate = False
    
    ShibbySettings.UserName = name
End Sub

' save the frx setting
Private Sub SaveFrxCleanup()
    ShibbySettings.FrxCleanup = FrxCleanupBox.value
End Sub

' save the export on git setting
Private Sub SaveExportOnGit()
    ShibbySettings.ExportOnGit = ExportOnGitBox.value
End Sub

' save the File structure
Private Sub SaveFileStructure()
    ShibbySettings.fileStructure = FileStructureBox.ListIndex
End Sub

' save remove files before export
Private Sub SaveRemoveFilesBeforeExport()
    ShibbySettings.RemoveFiles = RemoveFilesBox.value
End Sub
