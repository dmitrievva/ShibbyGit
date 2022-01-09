Attribute VB_Name = "GitCommits"

'*********************************************************
' Author:   Vladimir Dmitriev
' Link:     https://github.com/dmitrievva/ShibbyGit
'*********************************************************

Option Explicit

Public Const COMMIT_RAZD As String = "|"

' commitDetails = {
'   hash1: {
'       file1: changes1,
'       file2: changes2,
'       ...
'   },
'}
Private commitDetails   As Object
Private commitsArr      As Variant


Public Sub ClearCacheChanges()
    commitDetails.RemoveAll
End Sub

Public Sub ClearDiffCache()
    If Not commitDetails Is Nothing Then commitDetails.Item("") = Empty
End Sub

Public Static Function GetCommits() As Variant
    Dim commit      As Variant
    Dim Commits     As String
    Dim commitNames As Variant
    Dim result      As Variant
    Dim parts       As Variant
    Dim i           As Long
    Dim count       As Long
    Dim cmd         As String
    
    ' %h - hash
    ' %an - author name
    ' %ad - commit date
    ' %s - comment
    Commits = GitCommands.RunGitAsProcess("log --pretty=format:" & Chr(34) & _
                                            "%h " & COMMIT_RAZD & _
                                            " %an " & COMMIT_RAZD & _
                                            " %ad " & COMMIT_RAZD & _
                                            " %s" & Chr(34) & " --date=human")
                                            
    commitNames = Split(Commits, vbLf)
    
    ReDim result(UBound(commitNames) + 1, 3)
    
    result(0, 0) = "Hash"
    result(0, 1) = "Author"
    result(0, 2) = "Date"
    result(0, 3) = "Comment"
    
    count = 1
    For Each commit In commitNames
        If commit <> "" Then
            parts = Split(commit, COMMIT_RAZD)
            
            For i = LBound(parts) To UBound(parts)
                result(count, i) = parts(i)
            Next i
            count = count + 1
        End If
    Next commit
    
    commitsArr = result
    
    GetCommits = result
End Function


' result = {
'   file1: changes1,
'   file2: changes2,
'   ...,
'}
Public Static Function GetCommitDetail(hash As String) As Object
    Dim files       As String
    Dim file        As Variant
    Dim fileChanges As String
    Dim cmd         As String
    Dim res         As Object
    Dim needAdd     As Boolean
    
    If commitDetails Is Nothing Then Set commitDetails = CreateObject("Scripting.Dictionary")
    
    Set res = CreateObject("Scripting.Dictionary")
    
    If commitDetails.exists(hash) Then
        If TypeName(commitDetails.Item(hash)) = "Dictionary" Then
            Set GetCommitDetail = commitDetails.Item(hash)
            Exit Function
        End If
    End If
    
    If hash = "" Then
        cmd = "diff --name-only"
    Else
        cmd = "diff-tree --no-commit-id --name-only -r " & hash
    End If
    
    files = GitCommands.RunGitAsProcess(cmd)
    
    For Each file In Split(files, vbLf)
        needAdd = file <> Empty And Not (ShibbySettings.FrxCleanup And file Like "*.frx")
        
        If needAdd Then
            res.Item(file) = Empty
        End If
    Next file
    
    Set commitDetails.Item(hash) = res
    
    
    Set GetCommitDetail = commitDetails.Item(hash)
End Function

Function GetFileChanges(file As String, hash As String, Optional prevHash As String) As String
    Dim needAdd     As Boolean
    Dim result      As String
    Dim fileChanges As String
    Dim files       As Object
    Dim cmd         As String
    Dim key         As String
    
    key = IIf(prevHash = "", hash, hash & " " & prevHash)
    
    If IsEmpty(commitDetails.Item(key)) Then
        GetCommitDetail (key)
    End If
    
    needAdd = file <> Empty And Not (ShibbySettings.FrxCleanup And file Like "*.frx") And _
                            Not IsEmpty(commitDetails.Item(key))
    
    If needAdd Then
        Set files = commitDetails.Item(key)
        
        If files.Item(file) = Empty Then
            cmd = "diff " & key & " -- " & file
            
            files.Item(file) = GitCommands.RunGitAsProcess(cmd)

            Set commitDetails.Item(key) = files
        End If
        
        result = files.Item(file)
    End If
    
    GetFileChanges = result
End Function

Sub CheckoutSelectedCommit(commit As String)
    Dim output As String
    
    If commit = "" Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("checkout " & commit)
    
    MsgBox output
    
    GitTreeForm.ResetForm
    
End Sub

Sub RevertCommit(commit As String)
    Dim output As String
    
    If commit = "" Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("revert --no-commit " & commit)
    
    MsgBox output
    
    GitTreeForm.ResetForm
End Sub

Sub CherryPickCommit(commit As String)
    Dim output As String
    
    If commit = "" Then Exit Sub
    
    output = GitCommands.RunGitAsProcess("cherry-pick " & commit)
    
    MsgBox output
    
    GitTreeForm.ResetForm
End Sub

Sub AddSelectedFilesAndCommit(files As Collection)
    Dim cmd         As String
    Dim i           As Long
    Dim output      As String
    
    ' Add Selected files
    cmd = "add"
    For i = 1 To files.count
        cmd = cmd & " " & files(i)
    Next i
    
    output = GitCommands.RunGitAsProcess(cmd)
    
    If output = "" Then
        MsgBox "Staged all files for commit"
    Else
        MsgBox "Git response: " & vbCrLf & output
    End If
    
    ' Commit
    Load GitCommitMessageForm
    MoveFormOnApplication GitCommitMessageForm
    
    With GitCommitMessageForm
        .caption = "Git Commit Message"
        .TitleLabel.caption = "Enter Commit Message"
        .OKButton.caption = "commit -m"
    
        .Callback = "Commit"
        
        .Show False
    End With
End Sub

Sub SquashCommits(index As String, message As String)
    Dim output      As String
    
    ' interactive rebase
    GitCommands.RunGitInShell ("rebase -i HEAD~" & index)
    
    ' Commit
    Load GitCommitMessageForm
    MoveFormOnApplication GitCommitMessageForm
    
    With GitCommitMessageForm
        .caption = "Git Commit Message"
        .TitleLabel.caption = "Enter Commit Message"
        .OKButton.caption = "commit -m"
        
        .Placeholder = message
    
        .Callback = "Commit"
        
        .Show False
    End With
    
    GitTreeForm.ResetForm
End Sub

Public Property Get Commits() As Variant
    Commits = commitsArr
End Property

Public Property Let Commits(arr As Variant)
    commitsArr = arr
End Property

Public Sub ClearCommits()
    commitsArr = Empty
End Sub

