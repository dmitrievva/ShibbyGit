Attribute VB_Name = "FileUtils"
'***********************************************************************
' Original Author:   Eric Addison
' Link:     https://github.com/ericaddison/ShibbyGit
'
' Changed by: Vladimir Dmitriev, https://github.com/dmitrievva/ShibbyGit
'***********************************************************************

Option Explicit

Public Const GOOD_FOLDER As String = "goodFolder"
Public Const BAD_FOLDER As String = "badFolder"

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Check an incoming folder path to make sure it is valid
' broswe for a new one if invalid
' return true if found, false if not found
' input: folder - the folder to check
' output: if folder is good, outputs string GOOD_FOLDER
'           if folder is bad and none is chosen, outputs string BAD_FOLDER
'           if folder is bad but a good one is chosen, output the new folder
Public Function VerifyFolder(ByVal folder As String, Optional sTitle As String = "Browse") As String
    
    ' bad directory, browse for new one
    If FileOrDirExists(folder) = False Then
        folder = ""
    End If
    
    ' if nothing, browse for new folder
    If folder = "" Then
        folder = FolderBrowser(sTitle)
        ' if browse is cancelled, exit
        If (folder = "") Then
            VerifyFolder = BAD_FOLDER
            Exit Function
        End If
        VerifyFolder = folder
    Else
        VerifyFolder = GOOD_FOLDER
    End If
End Function


' get the base name of a file from its full path
' input: file - full path of a file
' output: the base name of the file (path removed)
Public Function FileBaseName(ByVal file As String) As String
    FileBaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(file)
End Function


'used to test filepaths of commmand button   links to see if they work - change their color if not working
' from http://superuser.com/questions/649745/check-if-path-to-file-is-correct-on-excel-column
Public Function FileOrDirExists(PathName As String) As Boolean
    'Macro Purpose: Function returns TRUE if the specified file
    Dim iTemp As Integer
    
    'Ignore errors to allow for error evaluation
    On Error Resume Next
        iTemp = GetAttr(PathName)
        
        'Check if error exists and set response appropriately
        Select Case Err.Number
            Case Is = 0
                FileOrDirExists = True
            Case Else
                FileOrDirExists = False
        End Select
    
    'Resume error checking
    On Error GoTo 0
End Function


' do events and wait loop
' input: nLoops - number of times to loop through DoEvents
' input: sleepTimeMs - number of ms to sleep in each loop
Public Sub DoEventsAndWait(ByVal nLoops As Integer, ByVal sleepTimeMs As Integer)
    Dim ind As Integer
    For ind = 1 To nLoops
        DoEvents
        Sleep sleepTimeMs
    Next ind
End Sub


' Launch a folder browser dialog
' optional input: sTitle - title of the browser dialog
' output: path to selected folder
Public Function FolderBrowser(Optional ByVal sTitle As String = "Browse") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = sTitle
    With fd
        If .Show = -1 Then
            FolderBrowser = .SelectedItems(1)
            Exit Function
        End If
    End With
    FolderBrowser = ""
End Function


' Launch a file browser dialog
' optional input: sTitle - title of the browser dialog
' output: path to selected file
Public Function FileBrowser(Optional ByVal sTitle As String = "Browse", _
        Optional ByVal filterDescription As String = "", _
        Optional ByVal filterExtenstion As String = "") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = sTitle
    If filterExtenstion <> "" Then
        fd.Filters.Add filterDescription, filterExtenstion, 1
    End If
    With fd
        If .Show = -1 Then
            FileBrowser = .SelectedItems(1)
            Exit Function
        End If
    End With
    FileBrowser = ""
End Function

' Launch a file browser dialog with multiselect on
' optional input: sTitle - title of the browser dialog
' output: selected items collection
Public Function FileBrowserMultiSelect(Optional ByVal sTitle As String = "Browse", _
        Optional ByVal filterDescription As String = "", _
        Optional ByVal filterExtenstion As String = "") As FileDialogSelectedItems
        
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = sTitle
    If filterExtenstion <> "" Then
        fd.Filters.Add filterDescription, filterExtenstion, 1
    End If
    fd.InitialView = msoFileDialogViewList
    With fd
        If .Show = -1 Then
            Set FileBrowserMultiSelect = .SelectedItems
            Exit Function
        End If
    End With
    Set FileBrowserMultiSelect = Nothing
End Function

' Open file folder in explorer
Public Sub OpenFileFolder(file As String)
    Dim folderPath  As String
    
    If file = "" Then Exit Sub
    If Dir(file, vbDirectory) = "" Then Exit Sub
    
    folderPath = left(file, InStrRev(file, "\"))
    Shell "explorer """ & folderPath & "", vbNormalFocus
    
End Sub

' Remove vba files
Public Sub RemoveVBAfiles(ByRef folderPath As Object)
    Dim fileObj         As Object
    Dim fileName        As String
    Dim objSubFolder    As Object
    Dim needDel         As Boolean
    
    For Each fileObj In folderPath.files
        fileName = UCase(fileObj.name)
        needDel = Right(fileName, 4) = ".BAS" Or Right(fileName, 4) = ".CLS" Or _
                Right(fileName, 4) = ".FRM"
        
        If needDel Then
            Kill fileObj
        End If
    Next fileObj
    
    For Each objSubFolder In folderPath.subfolders
        Call RemoveVBAfiles(objSubFolder)
    Next objSubFolder
    
    Set objSubFolder = Nothing
End Sub
