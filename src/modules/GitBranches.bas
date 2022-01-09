Attribute VB_Name = "GitBranches"
'*********************************************************
' Author:   Vladimir Dmitriev
' Link:     https://github.com/dmitrievva/ShibbyGit
'*********************************************************

Option Explicit

Sub CreateNewBranch()

    With GitCommitMessageForm
        .caption = "Create New Branch"
        .TitleLabel.caption = "Write branch name"
        .OKButton.caption = "checkout -b"
        
        .Callback = "RunGitNewBranch"
    End With
    
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show False

End Sub

Sub RunGitNewBranch(ByVal branchName As String)
    Dim output      As String
    
    output = GitCommands.RunGitAsProcess("checkout -b " & branchName)
    
    MsgBox output
    
    GitTreeForm.ResetForm
End Sub

Sub CheckoutSelectedBranch(branch As String)
    Dim output          As String
    
    If branch = "" Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("checkout " & branch)
    
    MsgBox output
    
    GitTreeForm.ResetForm
End Sub

Sub NewBranchFromSelectedBranch(branch As String)
    With GitCommitMessageForm
        .caption = "New Branch From Branch " & branch
        .TitleLabel.caption = "Write branch name"
        .OKButton.caption = "checkout -b"
        
        .Callback = "RunGitNewBranchFromSelected"
        .CallbackArguments = branch
    End With
    
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show False
End Sub

Sub RunGitNewBranchFromSelected(ByVal newBranchName As String, ByVal branchName As String)
    Dim output      As String
    
    output = GitCommands.RunGitAsProcess("checkout -b " & newBranchName & " " & branchName)
    
    MsgBox output
    
    GitTreeForm.ResetForm
End Sub

Sub MergeSelectedIntoCurrentBranch(currentBranch As String, selectedBranch As String)
    Dim output          As String
    
    If UCase(selectedBranch) = UCase(currentBranch) Then Exit Sub
    If selectedBranch = "" Or currentBranch = "" Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("merge " & currentBranch & " " & selectedBranch)
    
    MsgBox output
    
    GitTreeForm.ResetForm
End Sub

Sub RebaseCurrentOntoSelectedBranch(currentBranch As String, selectedBranch As String)
    Dim output          As String

    If UCase(selectedBranch) = UCase(currentBranch) Then Exit Sub
    If selectedBranch = "" Or currentBranch = "" Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("rebase --onto " & currentBranch & " " & selectedBranch)
    
    MsgBox output
    
    GitTreeForm.ResetForm
End Sub

Sub RenameSelectedBranch(ByVal branch As String)

    With GitCommitMessageForm
        .caption = "Rename Selected Branch "
        .TitleLabel.caption = "Write branch name"
        .OKButton.caption = "branch -m"
        
        .Callback = "RunRenameSelectedBranch"
        .CallbackArguments = branch
    End With
    
    MoveFormOnApplication GitCommitMessageForm
    GitCommitMessageForm.Show False
    
End Sub

Sub RunRenameSelectedBranch(ByVal newBranchName As String, ByVal selectedBranch As String)
    Dim output      As String
    
    output = GitCommands.RunGitAsProcess("branch -m " & selectedBranch & " " & newBranchName)
    
    MsgBox output
    
    GitTreeForm.ResetForm
    
End Sub

Sub DeleteSelectedBranch(selectedBranch As String)
    Dim answer          As VbMsgBoxResult
    Dim output          As String

    If selectedBranch = "" Then Exit Sub
    
    answer = MsgBox("Are you sure to delete branch '" & selectedBranch & "'?", vbYesNo, "Delete branch")
    If answer = vbNo Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("branch -D " & selectedBranch)
    
    MsgBox output
    
    GitTreeForm.ResetForm
    
End Sub

