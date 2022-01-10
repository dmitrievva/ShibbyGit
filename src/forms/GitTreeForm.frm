VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitTreeForm 
   Caption         =   "Git Tree"
   ClientHeight    =   9900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "GitTreeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitTreeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
' Author:   Vladimir Dmitriev
' Link:     https://github.com/dmitrievva/ShibbyGit
'*********************************************************

Option Explicit

Const MENU_NAME As String = "tempMenu"

Const SELECTED_COLOR As Long = 16501423

Const LOCAL_ID = "local_id"
Const REMOTE_ID = "remote_id"
Const DIFF = "diff"

Const ID_POSTFIX = "_id"

Public selectedBranch   As String
Public currentBranch    As String

Public selectedCommit   As String

Public selectedFile     As String

Dim buttonNumber        As Integer

Public WithEvents commitsListBox        As clsDesignListBox
Attribute commitsListBox.VB_VarHelpID = -1
Public WithEvents fileChangesListBox    As clsDesignListBox
Attribute fileChangesListBox.VB_VarHelpID = -1


'****************************************************************
' Initialize
'****************************************************************

Private Sub UserForm_initialize()

    Me.RefreshButton.caption = ChrW(8635)
    
    Me.AddSelectedFilesAndCommitButton.Visible = False
    Me.AddSelectedFilesAndCommitButton.Enabled = False
    
    Me.LabelLoading.Visible = False
    
    With Me.TV_files
        .CheckBoxes = False
        .LabelEdit = tvwManual
        
        .Indentation = 18
    End With
    
    
    With Me.TV_branches
        .LabelEdit = tvwManual
        .CheckBoxes = False
        
        .Indentation = 18
    End With

    
    Me.TV_branches.SetFocus
End Sub


Public Sub ResetForm(Optional forceRefresh As Boolean = False)
    Dim TV          As TreeView
    Dim localBr     As String
    Dim remote      As String
    Dim Commits     As Variant
    Dim DiffDict    As Object
    Dim data        As Variant
        
    Set TV = Me.TV_branches
    TV.Nodes.Clear
    
    Me.selectedBranch = ""
    Me.currentBranch = ""
    Me.selectedCommit = ""
    Me.selectedFile = ""
    
    Me.TV_files.Nodes.Clear
    
    If forceRefresh Then RefreshForm
    
    ' local branches
    localBr = GitCommands.RunGitAsProcess("branch")
    If left(localBr, 5) = "fatal" Then
        MsgBox "Some Errors occured. Please check Git Settings"
        Exit Sub
    End If
    
    Call AddParentNode(TV, "local", LOCAL_ID)
    Call AddBranchesToTree(TV, localBr, LOCAL_ID)
    
    TV.Nodes(LOCAL_ID).Expanded = True
    
    ' remote branches
    remote = GitCommands.RunGitAsProcess("branch -r")
    If left(remote, 5) = "fatal" Then
        MsgBox "Some Errors occured. Please check Git Settings"
        Exit Sub
    End If

    If remote <> "" Then
        Call AddParentNode(TV, "remote", REMOTE_ID)
        Call AddBranchesToTree(TV, remote, REMOTE_ID)
        
        TV.Nodes(REMOTE_ID).Expanded = True
    End If

    ' select active branch
    Call SelectActiveBranch(TV, LOCAL_ID)
    
    ' commits
    If IsEmpty(GitCommits.Commits) Then
        Commits = GetCommits()
    Else
        Commits = GitCommits.Commits
    End If
    
    ' uncommited diff
    Set DiffDict = GetCommitDetail("")
    
    If DiffDict.count > 0 Then
        data = FormDataWithUncommited(Commits)
    Else
        data = Commits
    End If
    
    If Not commitsListBox Is Nothing Then
        commitsListBox.Clear
        commitsListBox.Fill data
    Else
        Call SetupCommitsListBox(data)
    End If
    
    If DiffDict.count > 0 Then
        Call SelectDiffColor
    End If

End Sub

Private Function FormDataWithUncommited(commitsArr As Variant) As Variant
    Dim data        As Variant
    Dim i           As Long
    Dim j           As Integer
        
    ReDim data(UBound(commitsArr) + 1, 3)
    
    For j = 0 To 3
        data(0, j) = commitsArr(0, j)
    Next j
    
    data(1, 0) = DIFF
    data(1, 1) = "---"
    data(1, 2) = "Now"
    data(1, 3) = "Uncommited changes"
    
    For i = 1 To UBound(commitsArr, 1)
        For j = 0 To 3
            data(i + 1, j) = commitsArr(i, j)
        Next j
    Next i
    
    FormDataWithUncommited = data
End Function

'****************************************************
' Branches Tree View
'****************************************************

Private Sub AddBranchesToTree(TV As TreeView, branches As String, parentId As String)
    Dim branchNames     As Variant
    Dim branch          As Variant
    
    branchNames = Split(branches, vbLf)
    For Each branch In branchNames
        If branch <> "" Then
            branch = Replace(branch, "*", "")
            Call AddNodes(TV, TV.Nodes(parentId), CStr(branch))
        End If
    Next branch
End Sub

Private Sub SelectActiveBranch(TV As TreeView, localId As String)
    Dim id          As String
    Dim i           As Long
    Dim branchName  As String
    
    branchName = GitCommands.RunGitAsProcess("rev-parse --abbrev-ref HEAD")
    
    id = Trim(localId & "/" & branchName)
    id = left(id, Len(id) - 1)
    
    Me.currentBranch = id
    
    For i = 1 To TV.Nodes.count
        If TV.Nodes(i).key = id Then Exit For
    Next i
    
    If i > TV.Nodes.count Then Exit Sub

    TV.Nodes.Item(i).BackColor = SELECTED_COLOR
    
    With TV.Nodes(i)
        .Bold = True
        .Selected = True
        .Parent.Expanded = True
    End With
    
End Sub

Private Sub TV_branches_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    buttonNumber = Button
End Sub

Private Sub TV_branches_NodeClick(ByVal NodeItem As MSComctlLib.Node)
    Me.selectedBranch = NodeItem.key
    
    selectedBranch = Replace(Me.selectedBranch, LOCAL_ID & "/", "")
    selectedBranch = Replace(selectedBranch, REMOTE_ID & "/", "")
    
    If Not NodeItem.Child Is Nothing Then Exit Sub
    
    If buttonNumber = MouseButton.RIGHT_BUTTON Then
        BranchPopupClickMenu
    End If
End Sub

Sub BranchPopupClickMenu()
    Dim currentBranch       As String
    Dim selectedBranch      As String
    Dim action              As String
    
    DeleteMenu
    
    selectedBranch = Replace(Me.selectedBranch, LOCAL_ID, "")
    selectedBranch = Replace(selectedBranch, REMOTE_ID, "")
    
    currentBranch = Replace(Me.currentBranch, LOCAL_ID & "/", "")
    currentBranch = Replace(currentBranch, REMOTE_ID & "/", "")
    
    With CommandBars.Add(MENU_NAME, msoBarPopup, , False)
    
        ' create new branch
        action = "CreateNewBranch"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&Create new branch"
        End With
        
        ' checkout selected branch
        action = "'CheckoutSelectedBranch """ & selectedBranch & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "C&heckout"
            .BeginGroup = True
        End With
        
        ' new branch from selected
        action = "'NewBranchFromSelectedBranch """ & selectedBranch & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&New Branch from Selected"
        End With
        
        ' merge selected branch into current branch
        action = "'MergeSelectedIntoCurrentBranch """ & currentBranch & """, " & """" & selectedBranch & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&Merge Branch into Current Branch"
        End With
        
        ' rebase current onto selected branch
        action = "'RebaseCurrentOntoSelectedBranch """ & currentBranch & """, " & """" & selectedBranch & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "Re&base Current Branch onto Selected"
        End With
        
        ' rename selected branch
        action = "'RenameSelectedBranch """ & selectedBranch & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&Rename Branch"
            .BeginGroup = True
        End With
        
        ' delete selected branch
        action = "'DeleteSelectedBranch """ & selectedBranch & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&Delete Branch"
        End With
        
        
        .showPopup
    End With
    
End Sub

'****************************************************
' Commits DesignListBox
'****************************************************

Private Sub SetupCommitsListBox(data As Variant)
    Dim top         As Long
    Dim left        As Long
    Dim width       As Long
    Dim height      As Long
    Dim scroll      As fmScrollBars
    
    top = Me.TV_branches.top
    
    height = Me.TV_branches.height
    
    left = Me.TV_branches.left + Me.TV_branches.width + 10
    
    width = Me.width - left - 20
        
    Dim DesignListBox As New clsDesignListBox
    
    DesignListBox.ColumnWidths = "50;70;70;" & width - (50 + 70 + 70)
    DesignListBox.Headers = True
    DesignListBox.multiSelect = True
    
    scroll = IIf(UBound(data) > 12, fmScrollBarsVertical, fmScrollBarsNone)
    
    DesignListBox.Create Me, top, left, height, width, data, scroll, 1
    
    
    Set commitsListBox = DesignListBox
End Sub

Private Sub commitsListBox_Click()
    Dim Commits     As Variant
    Dim commit      As String
    Dim isDiff      As Boolean
    Dim showPopup   As Boolean
    Dim index       As Long
    Dim prevCommit  As String
    Dim rowCollect  As Collection

    Commits = CheckTwoCommits()
    
    If IsEmpty(Commits) Then Exit Sub
    
    commit = Commits(0) & " " & Commits(1)

    buttonNumber = Me.commitsListBox.Events("mouseUp")(1)
    
    isDiff = (Commits(0) = DIFF Or Commits(0) = "")
    
    ' show popup menu if clicked right button and selected commit is not uncommited diff
    showPopup = (buttonNumber = MouseButton.RIGHT_BUTTON) And Not isDiff
                    
    If showPopup Then
        CommitPopupClickMenu
        Exit Sub
    End If
    
    ' load commit data
    Call Me.StartLoading
    
    Call ClearListBox(Me.fileChangesListBox)
    
    Me.TV_files.CheckBoxes = isDiff
    Me.AddSelectedFilesAndCommitButton.Visible = isDiff
    
    If isDiff Then commit = ""
    
    Me.selectedCommit = commit
    Me.selectedFile = ""
    
    Call WriteCommitDetails(commit)
    
    Call Me.StopLoading
    
End Sub

Private Function CheckTwoCommits() As Variant
    Dim arr         As Variant
    Dim index       As Long
    Dim commit1     As String
    Dim commit2     As String
    Dim rowCollect  As Collection
    
    index = Me.commitsListBox.ListIndex(0)
    If index = -1 Then
        CheckTwoCommits = arr
        Exit Function
    End If
    
    Set rowCollect = Me.commitsListBox.RowLabels(index, False)
    Me.selectedCommit = rowCollect(1)
    
    commit1 = Me.selectedCommit
        
    index = Me.commitsListBox.ListIndex(0)
    If index < Me.commitsListBox.RowsCount - 1 Then
        Set rowCollect = Me.commitsListBox.RowLabels(index + 1, False)
        If rowCollect.count > 0 Then
            commit2 = rowCollect(1).caption
        End If
    End If
    
    If UBound(Me.commitsListBox.ListIndex) = 1 Then
        index = Me.commitsListBox.ListIndex(1)
        Set rowCollect = Me.commitsListBox.RowLabels(index, False)
        commit2 = rowCollect(1).caption
    End If
    
    commit1 = Replace(commit1, DIFF, "")
    commit2 = Replace(commit2, DIFF, "")
    
    If commit1 = commit2 Then commit2 = ""
    
    arr = Array(commit1, commit2)
    
    CheckTwoCommits = arr
End Function

Sub WriteCommitDetails(commit As String)
    Dim details     As Object
    
    Me.selectedCommit = commit
    
    Set details = GetCommitDetail(commit)
    
    Me.TV_files.Nodes.Clear
    
    Call AddFilesToTree(Me.TV_files, details)
End Sub

Sub CommitPopupClickMenu()
    Dim action              As String
    Dim selectedCommits     As Variant
    Dim indexes             As Variant
    Dim i                   As Long
    Dim rowCollect          As Collection
    
    selectedCommits = Me.commitsListBox.SelectedValue
    
    Me.selectedCommit = Join(selectedCommits, " ")
    
    indexes = Me.commitsListBox.ListIndex
    
    DeleteMenu
    
    
    With CommandBars.Add(MENU_NAME, msoBarPopup, , False)
    
        ' create new branch
        action = "CreateNewBranch"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "Create &new branch"
        End With
        
        ' copy commit hash
        action = "'Clipboard """ & Me.selectedCommit & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&Copy commit hash"
        End With
        
        ' copy commit message
        Dim message As String
        
        message = MessagesFromSelectedCommits()
        
        action = "'Clipboard """ & message & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "Copy commit &message"
        End With
        
        ' checkout selected commit
        action = "'CheckoutSelectedCommit """ & Me.selectedCommit & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "Ch&eckout"
            .BeginGroup = True
        End With
        
        ' cherry pick
        action = "'CherryPickCommit """ & Me.selectedCommit & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "Cherr&y-pick"
        End With
        
        ' revert
        action = "'RevertCommit """ & Me.selectedCommit & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&Revert"
        End With
        
        ' squash commits
        If CheckCanSquash Then
            message = MessagesFromSelectedCommits
            
            indexes = Me.commitsListBox.ListIndex
            i = indexes(UBound(indexes))
            
            action = "'SquashCommits """ & CStr(i) & """, " & """" & Replace(message, "'", "") & """'"

            With .Controls.Add(msoControlButton, , , , False)
                .OnAction = action
                .caption = "&Squash commits"
            End With
        End If
        
        .showPopup
    End With
End Sub

Private Function MessagesFromSelectedCommits() As String
    Dim message     As String
    Dim indexes     As Variant
    Dim rowCollect  As Collection
    Dim i           As Long
    
    indexes = Me.commitsListBox.ListIndex
    
    For i = LBound(indexes) To UBound(indexes)
        Set rowCollect = Me.commitsListBox.RowLabels(CLng(indexes(i)), False)
        message = message & vbLf & rowCollect(rowCollect.count)
    Next i
    
    If Strings.left(message, Len(vbLf)) = vbLf Then
        message = Right(message, Len(message) - Len(vbLf))
    End If
    
    MessagesFromSelectedCommits = message
End Function

Private Function CheckCanSquash() As Boolean
    Dim indexes     As Variant
    Dim i           As Long
    Dim canSquash   As Boolean
    Dim commit      As String
    Dim rowCollect  As Collection
    
    indexes = Me.commitsListBox.ListIndex
    Set rowCollect = Me.commitsListBox.RowLabels(1, False)
    commit = rowCollect(1)
    
    canSquash = UBound(indexes) >= 1 And commit <> "" And commit <> DIFF
    
    If canSquash Then
        For i = LBound(indexes) To UBound(indexes) - 1
            If indexes(i) + 1 <> indexes(i + 1) Then
                canSquash = False
                Exit For
            End If
        Next i
    End If
    
    CheckCanSquash = canSquash
End Function

Private Sub SelectDiffColor()
    Dim Labl        As Object
    Dim rowNumber   As String
    
    For Each Labl In Me.commitsListBox.AllLabels
        rowNumber = Split(Labl.name, ";")(0)
        If rowNumber = 1 Then
            Labl.ForeColor = vbRed
        End If
    Next
End Sub


'****************************************************
' Files Tree View
'****************************************************

Private Sub AddFilesToTree(TV As TreeView, files As Object)
    Dim file    As Variant
    Dim curr    As String
    Dim prev    As String
    Dim arr     As Variant
    Dim i       As Integer
        
    TV.Nodes.Clear
    
    For Each file In files
        If file = "" Then GoTo next_file
        arr = Split(file, "/")
        prev = ""
        
        For i = LBound(arr) To UBound(arr)
            curr = arr(i)

            If prev <> "" Then curr = prev & "/" & arr(i)

            If Not IsNodeExists(TV, curr) Then
                If prev <> "" And IsNodeExists(TV, prev) Then
                    Call AddChildNode(TV, prev, CStr(arr(i)), curr)
                    TV.Nodes(prev).Expanded = True
                Else
                    Call AddParentNode(TV, CStr(arr(i)), curr)
                End If
            End If
            
            prev = curr
        Next i
next_file:
    Next file
    
    
End Sub


Private Sub TV_files_NodeClick(ByVal NodeItem As MSComctlLib.Node)
    Dim key             As String
    
    key = Replace(NodeItem.key, ID_POSTFIX, "")
    
    If Not NodeItem.Child Is Nothing Then
        Call ClearListBox(Me.fileChangesListBox)
        Exit Sub
    End If
    
    If Me.selectedFile = key And buttonNumber = MouseButton.RIGHT_BUTTON Then
        Call FilesPopupClickMenu
        Exit Sub
    End If

    
    Me.selectedFile = key
    
    Call Me.StartLoading
    
    Call ClearListBox(Me.fileChangesListBox)
    
    Call LoadFileChangesDetails(NodeItem)
    
    Call Me.StopLoading
    
End Sub

Private Sub TV_files_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim checkedCol  As New Collection
    Dim i           As Long
    
    With Me.TV_files
        Call CheckNodeChilds(Me.TV_files, Node)
        
        For i = 1 To .Nodes.count
            If .Nodes(i).Checked Then
                checkedCol.Add .Nodes(i).text
            End If
        Next i
    End With
    
    Me.AddSelectedFilesAndCommitButton.Enabled = (checkedCol.count > 0)
    
End Sub

Private Sub TV_files_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    buttonNumber = Button
End Sub

Sub FilesPopupClickMenu()
    Dim action              As String
    Dim file                As String
    
    DeleteMenu
    
    With CommandBars.Add(MENU_NAME, msoBarPopup, , False)
        ' copy file name
        action = "'Clipboard """ & Me.selectedFile & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&Copy File name"
        End With
        
        ' copy file path
        file = ShibbySettings.GitProjectPath & "\" & Me.selectedFile
        action = "'Clipboard """ & file & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "Co&py File path"
        End With
        
        ' view file in editor
        action = "'EditSelectedFile """ & Me.selectedFile & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&View in Editor"
            .BeginGroup = True
        End With
        
        ' Reveal file in explorer
        file = ShibbySettings.GitProjectPath & "\" & Me.selectedFile
        action = "'OpenFileFolder """ & file & """'"
        With .Controls.Add(msoControlButton, , , , False)
            .OnAction = action
            .caption = "&Reveal in explorer"
        End With
        
        
        .showPopup
    End With
End Sub

Private Sub AddSelectedFilesAndCommitButton_Click()
    Dim filesCol    As New Collection
    Dim i           As Long
    
    With Me.TV_files
        For i = 1 To .Nodes.count
            If .Nodes(i).Checked Then
                filesCol.Add .Nodes(i).text
            End If
        Next i
    End With
    
    If filesCol.count = 0 Then Exit Sub
    
    Call AddSelectedFilesAndCommit(filesCol)
    
    Me.ResetForm
    
End Sub

'****************************************************
' File Changes DesignListBox
'****************************************************
Private Sub LoadFileChangesDetails(ByVal NodeItem As MSComctlLib.Node)
    Dim details         As String
    Dim arr             As Variant
    Dim data            As Variant
    Dim prevCommit      As String
    Dim index           As Long
    Dim rowCollect      As Collection
    Dim i               As Long
    Dim Commits         As Variant
    
    Call ClearNodesFormat(Me.TV_files)
    
    NodeItem.BackColor = SELECTED_COLOR
    NodeItem.Bold = True

    Commits = CheckTwoCommits()
    If IsEmpty(Commits) Then Exit Sub
    
    details = GetFileChanges(Me.selectedFile, CStr(Commits(1)), CStr(Commits(0)))
    If details = "" Then Exit Sub
    
    arr = Split(details, vbLf)
    
    ReDim data(UBound(arr), 1)
    For i = LBound(arr) To UBound(arr)
        data(i, 0) = arr(i)
    Next i
    
    If fileChangesListBox Is Nothing Then
        Call SetupFileChangesListBox(data)
    Else
        fileChangesListBox.Clear
        fileChangesListBox.Fill data, "2"
    End If
    
    Call SetColorsToFileChanges

    
End Sub

Private Sub SetupFileChangesListBox(data As Variant)
    Dim top         As Long
    Dim left        As Long
    Dim width       As Long
    Dim height      As Long
    
    top = Me.TV_files.top
    
    height = Me.TV_files.height
    
    left = Me.TV_files.left + Me.TV_files.width + 10
    
    width = Me.width - left - 20
    
    
    Dim DesignListBox As New clsDesignListBox
        
    DesignListBox.ColumnWidths = width
    DesignListBox.Headers = False
    DesignListBox.multiSelect = True
    
    DesignListBox.Create Me, top, left, height, width, data, fmScrollBarsVertical, listBoxNumber:=2
    
    Set fileChangesListBox = DesignListBox
End Sub

Private Sub SetColorsToFileChanges()
    Dim Labl    As MSForms.Label
    Dim Number  As Double
    
    Dim green   As Long
    Dim red     As Long
    
    green = RGB(219, 230, 194)
    red = RGB(255, 204, 204)
    
    
    For Each Labl In Me.fileChangesListBox.AllLabels
        Labl.Tag = Replace(Labl.Tag, "color", "")
        
        If Strings.left(Labl, 1) = "+" Then
            Labl.BackColor = green
            Labl.Tag = Labl.Tag & "color:" & green
        End If
        
        If Strings.left(Labl, 1) = "-" Then
            Labl.BackColor = red
            Labl.Tag = Labl.Tag & "color:" & red
        End If
    Next Labl
End Sub


'****************************************************
' Tree View Common Functions
'****************************************************

Private Sub AddNodes(TV As TreeView, rootNode As MSComctlLib.Node, path As String, Optional expandParent As Boolean = False)
    Dim parentNode  As MSComctlLib.Node
    Dim nodeKey     As String
    Dim pathNodes() As String
    Dim i           As Long
    
    On Error GoTo errH
    
    pathNodes = Split(path, "/")
    nodeKey = rootNode.key
    For i = LBound(pathNodes) To UBound(pathNodes)
        Set parentNode = TV.Nodes(nodeKey)
        If Right(nodeKey, 1) <> "/" Then nodeKey = nodeKey & "/" & Trim(pathNodes(i))
        
        Call AddChildNode(TV, parentNode.key, pathNodes(i), nodeKey)
        
        parentNode.Expanded = expandParent
    Next i
    
errH:
    If Err.Number = 35601 Then
        Set parentNode = rootNode
        Resume
    End If
    Resume Next
End Sub

Private Sub AddParentNode(TV As TreeView, text As String, Optional id As String)
    If id = "" Then id = text & ID_POSTFIX
    If Not IsNodeExists(TV, id) Then
        TV.Nodes.Add key:=id, text:=text
    End If
End Sub

Private Sub AddChildNode(TV As TreeView, parentId As String, text As String, Optional id As String)
    If id = "" Then id = text & ID_POSTFIX
    If Not IsNodeExists(TV, id) Then
        TV.Nodes.Add Relative:=parentId, Relationship:=tvwChild, key:=id, text:=text
    End If
End Sub

Private Sub ClearNodesFormat(TV As TreeView)
    Dim i           As Long
    
    For i = 1 To TV.Nodes.count
        TV.Nodes(i).BackColor = vbWhite
        TV.Nodes(i).Bold = False
    Next i
End Sub

Private Function IsNodeExists(TV As TreeView, key As String)
    Dim NodeItem    As MSComctlLib.Node
    
    For Each NodeItem In TV.Nodes
        If Trim(UCase(NodeItem.key)) = Trim(UCase(key)) Then
            IsNodeExists = True
            Exit Function
        End If
    Next
    
    IsNodeExists = False
End Function

Private Sub CheckNodeChilds(TV As TreeView, Node As MSComctlLib.Node)
    Dim i           As Long
    
    If Node.Children = 0 Then Exit Sub
    
    For i = 1 To TV.Nodes.count
        If TV.Nodes(i).Parent Is Nothing Then GoTo next_node
        
        If TV.Nodes(i).Parent.key = Node.key Then
            TV.Nodes(i).Checked = Node.Checked
            
            If TV.Nodes(i).Children > 0 Then
                Call CheckNodeChilds(TV, TV.Nodes(i))
            End If
        End If
        
next_node:
    Next i
End Sub


'****************************************************
' Userform Functions
'****************************************************

Private Sub DeleteMenu()
    On Error Resume Next
    CommandBars(MENU_NAME).Delete
    On Error GoTo 0
End Sub

Private Sub ClearListBox(LB As clsDesignListBox)
    If Not LB Is Nothing Then LB.Clear
End Sub

Sub StartLoading()

    With Me.LabelLoading
        .left = Me.width - 150
        .width = 130
        .height = 20
        
        .caption = "Loading, please wait"
        .TextAlign = fmTextAlignCenter
        .Font.Bold = True
        
        .BackColor = RGB(165, 214, 167)
        .ForeColor = vbWhite

        .Visible = True
    End With
    
    Me.Repaint
End Sub

Sub StopLoading()
    Me.LabelLoading.Visible = False
End Sub

'****************************************************
' Userform Refresh
'****************************************************

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RefreshForm
End Sub

Private Sub RefreshButton_Click()
    RefreshForm
End Sub

Private Sub RefreshForm()
    

    GitCommits.ClearCacheChanges
    GitCommits.ClearCommits
    
    Me.TV_branches.Nodes.Clear
    Me.TV_files.Nodes.Clear
    
    Call ClearListBox(Me.fileChangesListBox)
    Call ClearListBox(Me.commitsListBox)
    
    Me.ResetForm
End Sub

