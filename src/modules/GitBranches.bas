Attribute VB_Name = "GitBranches"
'*********************************************************
' Author:   Vladimir Dmitriev
' Link:     https://github.com/dmitrievva/ShibbyGit
'*********************************************************

Option Explicit

Sub CreateNewBranch()
    
    Load GitCommitMessageForm
    
    With GitCommitMessageForm
        .caption = "Create New Branch"
        .TitleLabel.caption = "Write branch name"
        .OKButton.caption = "checkout -b"
        
        .Callback = "RunGitNewBranch"
        
        .Placeholder = ""
    End With
    
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show False
   
End Sub

Sub RunGitNewBranch(ByVal branchName As String)
    Dim output      As String
    
    output = GitCommands.RunGitAsProcess("checkout -b " & branchName)
    
    MsgBox output
    
    GitTreeForm.ResetForm forceRefresh:=True
End Sub

Sub CheckoutSelectedBranch(branch As String)
    Dim output          As String
    
    If branch = "" Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("checkout " & branch)
    
    MsgBox output
    
    GitTreeForm.ResetForm forceRefresh:=True
End Sub

Sub NewBranchFromSelectedBranch(branch As String)
    Load GitCommitMessageForm
    
    With GitCommitMessageForm
        .caption = "New Branch From Branch " & branch
        .TitleLabel.caption = "Write branch name"
        .OKButton.caption = "checkout -b"
        
        .Callback = "RunGitNewBranchFromSelected"
        .CallbackArguments = branch
        
        .Placeholder = branch
    End With
    
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show False
    
End Sub

Sub RunGitNewBranchFromSelected(ByVal newBranchName As String, ByVal branchName As String)
    Dim output      As String
    
    output = GitCommands.RunGitAsProcess("checkout -b " & newBranchName & " " & branchName)
    
    MsgBox output
    
    GitTreeForm.ResetForm forceRefresh:=True
End Sub

Sub MergeSelectedIntoCurrentBranch(currentBranch As String, selectedBranch As String)
    Dim output          As String
    
    If UCase(selectedBranch) = UCase(currentBranch) Then Exit Sub
    If selectedBranch = "" Or currentBranch = "" Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("merge " & currentBranch & " " & selectedBranch)
    
    MsgBox output
    
    GitTreeForm.ResetForm forceRefresh:=True
End Sub

Sub RebaseCurrentOntoSelectedBranch(currentBranch As String, selectedBranch As String)
    Dim output          As String

    If UCase(selectedBranch) = UCase(currentBranch) Then Exit Sub
    If selectedBranch = "" Or currentBranch = "" Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("rebase --onto " & currentBranch & " " & selectedBranch)
    
    MsgBox output
    
    GitTreeForm.ResetForm forceRefresh:=True
End Sub

Sub RenameSelectedBranch(ByVal branch As String)
    Load GitCommitMessageForm
    
    With GitCommitMessageForm
        .caption = "Rename Selected Branch "
        .TitleLabel.caption = "Write branch name"
        .OKButton.caption = "branch -m"
        
        .Placeholder = branch
        
        .Callback = "RunRenameSelectedBranch"
        .CallbackArguments = branch
    End With
    
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show False
   
End Sub

Sub RunRenameSelectedBranch(ByVal newBranchName As String, ByVal selectedBranch As String)
    Dim output      As String
    
    output = GitCommands.RunGitAsProcess("branch -m " & selectedBranch & " " & newBranchName)
    
    If output = "" Then
        MsgBox "No output from git"
    Else
        MsgBox output
    End If
    
    GitTreeForm.ResetForm forceRefresh:=False
    
End Sub

Sub DeleteSelectedBranch(selectedBranch As String)
    Dim answer          As VbMsgBoxResult
    Dim output          As String

    If selectedBranch = "" Then Exit Sub
    
    answer = MsgBox("Are you sure to delete branch '" & selectedBranch & "'?", vbYesNo, "Delete branch")
    If answer = vbNo Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("branch -D " & selectedBranch)
    
    MsgBox output
    
    GitTreeForm.ResetForm forceRefresh:=True
    
End Sub

