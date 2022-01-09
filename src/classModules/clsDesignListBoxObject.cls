VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDesignListBoxObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************
' Author:   balin77
' Link:     https://github.com/balin77/ExcelDesignListBox
'*********************************************************

Option Explicit

Private WithEvents pLabl        As MSForms.Label
Attribute pLabl.VB_VarHelpID = -1
Private pParent                 As clsDesignListBox
Private multiSelect             As Boolean

' --------------------------
' --------------------------
' Events
' --------------------------
' --------------------------

Private Sub pLabl_Click()

    Call pParent.RaiseEventBeforeClick
    
    If pLabl.Tag = "Header" Then Exit Sub
    
    Dim rowNumber As Long
    rowNumber = Split(pLabl.name, ";")(0)
    
    Call pParent.SelectRow(rowNumber, multiSelect)
    Call pParent.RaiseEventClick

End Sub

Private Sub pLabl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    multiSelect = (Shift = 1 Or Shift = 2)

    Set pParent.Events = Array("mouseUp", Button, Shift, X, Y)

    Call pParent.RaiseEventClick
End Sub

' --------------------------
' --------------------------
' Public Functions
' --------------------------
' --------------------------

Public Sub Create(InpLabel As MSForms.Label, InpParent As clsDesignListBox)

    Set pLabl = InpLabel
    Set pParent = InpParent

End Sub

' --------------------------
' --------------------------
' Private Functions
' --------------------------
' --------------------------

