Attribute VB_Name = "UI"
'***********************************************************************
' Original Author:   Eric Addison
' Link:     https://github.com/ericaddison/ShibbyGit
'
' Changed by: Vladimir Dmitriev, https://github.com/dmitrievva/ShibbyGit
'***********************************************************************

Option Explicit

Public Sub DoGitAddAll(Control As IRibbonControl)
    GitCommands.GitAddAll
End Sub

Public Sub DoGitStatus(Control As IRibbonControl)
    GitCommands.GitStatus
End Sub

Public Sub DoGitLog(Control As IRibbonControl)
    GitCommands.GitLog
End Sub

Public Sub DoGitInit(Control As IRibbonControl)
    GitCommands.GitInit
End Sub

Public Sub ShowGitSettingsForm(Control As IRibbonControl)
    Load GitSettingsForm
    GitSettingsForm.ResetForm
    MoveFormOnApplication GitSettingsForm
    GitSettingsForm.Show
    Unload GitSettingsForm
End Sub

Public Sub ShowGitRemoteForm(Control As IRibbonControl)
    Load GitRemoteForm
    GitRemoteForm.ResetForm
    MoveFormOnApplication GitRemoteForm
    GitRemoteForm.Show False
End Sub


Public Sub ShowGitCommitForm(Control As IRibbonControl)
    Load GitCommitMessageForm
    GitCommitMessageForm.ResetForm
    
    With GitCommitMessageForm
        .caption = "Git Commit Message"
        .TitleLabel.caption = "Enter Commit Message"
        .OKButton.caption = "commit -am"
        
        .Callback = "GitCommitAm"
    End With
    
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show False
End Sub

Public Sub ShowGitOtherForm(Control As IRibbonControl)
    If ShibbySettings.ExportOnGit Then
        GitIO.GitExport ShibbySettings.GitProjectPath, ShibbySettings.fileStructure
    End If
    Load GitConsoleForm
    MoveFormOnApplication GitConsoleForm
    GitConsoleForm.OutputBox.scrollBars = fmScrollBarsVertical
    GitConsoleForm.Show False
End Sub

Public Sub ShowGitTreeForm(Control As IRibbonControl)
    Load GitTreeForm
    GitTreeForm.ResetForm
    MoveFormOnApplication GitTreeForm
    GitTreeForm.Show False
End Sub

Public Sub NonModalMsgBox(ByVal message As String)
    Load NonModalMsgBoxForm
    MoveFormOnApplication NonModalMsgBoxForm
    NonModalMsgBoxForm.Show False
    NonModalMsgBoxForm.Label1.caption = message
End Sub

Public Sub HideNonModalMsgBox()
    NonModalMsgBoxForm.Hide
End Sub


Public Sub MoveFormOnApplication(ByVal Form As Variant)
    Form.left = Application.ActiveWindow.left
    Form.top = Application.ActiveWindow.top
End Sub


' public interface for export all
Public Sub ExportAllMsgBox(Control As IRibbonControl)
    Dim folder As String
    folder = FileUtils.FolderBrowser("Browse for folder for export")
    If folder = "" Then
        Exit Sub
    End If
    NonModalMsgBox "Exporting files" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = CodeUtils.ExportAllString(folder)
    
    HideNonModalMsgBox
    MsgBox output
End Sub


' public interface for import from
Public Sub ImportSelectedMsgBox(Control As IRibbonControl)
    Dim files As FileDialogSelectedItems
    Set files = FileUtils.FileBrowserMultiSelect("Browse for code files to import", _
            "VBA Code Module", "*.bas; *.frm; *.cls")
    
    If files Is Nothing Then
        Exit Sub
    End If
    
    NonModalMsgBox "Importing files" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = CodeUtils.ImportSelectedString(files)
    
    HideNonModalMsgBox
    MsgBox output
End Sub


' public interface for GitExport
Public Sub GitExportMsgBox(Control As IRibbonControl)
    NonModalMsgBox "Exporting files to Git Folder" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = GitIO.GitExport(ShibbySettings.GitProjectPath, ShibbySettings.fileStructure)
    If output = "" Then
        output = "No files exported"
    End If
    
    HideNonModalMsgBox
    MsgBox output
End Sub

' public interface for GitImport
Public Sub GitImportMsgBox(Control As IRibbonControl)
    NonModalMsgBox "Importing files from Git Folder" & vbCrLf & vbCrLf & "This could take a second or two . . ."
    FileUtils.DoEventsAndWait 10, 2
    
    Dim output As String
    output = GitIO.GitImport(ShibbySettings.GitProjectPath, ShibbySettings.fileStructure)
    If output = "" Then
        output = "No files imported"
    End If
    
    HideNonModalMsgBox
    MsgBox output
End Sub

' public interface for Export Code and Commit
Public Sub GitExportAndCommit(Control As IRibbonControl)
    Dim output      As String
    
    ' Export code
    output = GitIO.GitExport(ShibbySettings.GitProjectPath, ShibbySettings.fileStructure)
    
    ' Add All
    GitCommands.GitAddAll
    
    Load GitCommitMessageForm
    GitCommitMessageForm.ResetForm
    
    ' Commit
    With GitCommitMessageForm
        .caption = "Git Commit Message"
        .TitleLabel.caption = "Enter Commit Message"
        .OKButton.caption = "commit -am"
        
        .Callback = "GitCommitAm"
    End With
    
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show False
End Sub
