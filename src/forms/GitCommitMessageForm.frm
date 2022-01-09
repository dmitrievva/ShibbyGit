VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitCommitMessageForm 
   Caption         =   "Git Commit Message"
   ClientHeight    =   1950
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6375
   OleObjectBlob   =   "GitCommitMessageForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitCommitMessageForm"
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

Private callbackFn      As String
Private callbackArgs    As Variant
Private pText           As String

Public Property Get Callback() As String
    Callback = callbackFn
End Property

Public Property Let Callback(cb As String)
    callbackFn = cb
End Property

Public Property Get CallbackArguments() As Variant
    CallbackArguments = callbackArgs
End Property

Public Property Let CallbackArguments(args As Variant)
    callbackArgs = args
End Property

Public Property Get Placeholder() As String
    Placeholder = pText
End Property

Public Property Let Placeholder(text As String)
    pText = text
    Me.MessageTextBox.text = pText
End Property

Private Sub UserForm_initialize()
    Me.ResetForm
End Sub

Private Sub CancelButton_Click()
    GitCommitMessageForm.Hide
End Sub

Public Sub ResetForm()
    callbackFn = ""
    callbackArgs = ""
    pText = ""
    
    Me.MessageTextBox.text = pText
End Sub

Private Sub MessageTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OKButton.setFocus
        DoEvents
        OKButton_Click
    End If
End Sub

Private Sub OKButton_Click()
    Dim commitMessage As String
    commitMessage = Me.MessageTextBox.text
    
    If commitMessage = "" Then
        MsgBox "Please enter a commit message"
        Exit Sub
    End If
    
    If callbackFn <> "" Then
        On Error Resume Next
        
        If IsEmpty(callbackArgs) Or callbackArgs = "" Then
            Call Application.Run(callbackFn, commitMessage)
        Else
            Call Application.Run(callbackFn, commitMessage, callbackArgs)
        End If
    End If

    GitCommitMessageForm.Hide
End Sub
