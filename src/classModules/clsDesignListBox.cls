VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDesignListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************
' Author:   balin77
' Link:     https://github.com/balin77/ExcelDesignListBox
'*********************************************************

Option Explicit

Private pHeadersCol                 As Collection
Private pDesignListBoxObjectsCol    As New Collection
Private pFreezeColumnsCol           As Collection
Private pFreezeRowsCol              As Collection
Private pInpArr()                   As Variant
Private pParent                     As UserForm
Private WithEvents pFrame           As MSForms.frame
Attribute pFrame.VB_VarHelpID = -1
Private pParamDict                  As Object
Private pIsEmpty                    As Boolean

' Public Events
Public Event BeforeClick()
Public Event Change()
Public Event Click()

' --------------------------
' Properties
' --------------------------
Private Sub InitParamDict()
    If pParamDict Is Nothing Then Set pParamDict = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get IsEmpty() As Boolean

    IsEmpty = pIsEmpty

End Property

Public Property Get frame() As MSForms.frame

    Set frame = pFrame

End Property

Public Property Get ColumnsCount() As Long
    'has to return 0 if empty. No Error Handler required
    ColumnsCount = pParamDict("ColumnsCount")

End Property

Public Property Get RowsCount() As Long

    'has to return 0 if empty. No Error Handler required
    RowsCount = pParamDict("RowsCount")

End Property

Public Property Get RowHeight() As Long

    'has to return 0 if empty. No Error Handler required
    RowHeight = pParamDict("RowHeight")

End Property

Public Property Let RowHeight(InpHeight As Long)

    If InpHeight < 1 Then InpHeight = 1
    
    pParamDict("RowHeight") = InpHeight
    
    Dim AllLabelsCol As Collection
    Set AllLabelsCol = Me.AllLabels
    
    Dim Labl
    For Each Labl In AllLabelsCol
        Labl.height = InpHeight
    Next Labl

End Property

Public Property Get ColumnWidths() As String
    
    'has to return vbnullstring if empty. No Error Handler required
    ColumnWidths = pParamDict("ColumnWidths")

End Property

Public Property Let ColumnWidths(strWidths As String)

    Dim k As Long
    
    If Me.ColumnsCount = 0 Then
        pParamDict("ColumnsCount") = UBound(Split(strWidths, ";")) + 1
    End If
    
    'Adapt Widths to ColumnsCount
    If Me.ColumnsCount > UBound(Split(strWidths, ";")) Then
        Do While Me.ColumnsCount > UBound(Split(strWidths, ";"))
            If Right(strWidths, 1) = ";" Then
                strWidths = strWidths + "0;"
            Else
                strWidths = strWidths + ";0;"
            End If
        Loop
    ElseIf Me.ColumnsCount < UBound(Split(strWidths, ";")) Then
        Dim pivotString As String, i As Long
        i = 0
        Do Until Me.ColumnsCount < UBound(Split(pivotString, ";"))
            pivotString = pivotString & Split(strWidths, ";")(i) & ";"
            i = i + 1
        Loop
        strWidths = pivotString
    End If
    
    'Eliminate last ;
    If Right(strWidths, 1) = ";" Then
        strWidths = left(strWidths, Len(strWidths) - 1)
    End If
    
    For k = 0 To Me.ColumnsCount - 1
        If Not IsNumeric(Split(strWidths, ";")(k)) Then
            Err.Raise 123, "clsDesignListBox", "Input for ColumnWidths is not correct.", "Correct notation example: '10; 10; 10;'"
            Exit Property
        End If
    Next k
    
    pParamDict("ColumnWidths") = strWidths
    
    On Error GoTo ErrHandler
    If Not UBound(pInpArr, 1) = 0 Then
        Call Me.Fill(pInpArr)
    End If
    
    Exit Property
ErrHandler:

End Property

Public Property Get Headers() As Boolean

    Headers = pParamDict("Headers")

End Property

Public Property Let Headers(isOn As Boolean)

    pParamDict("Headers") = isOn
    
    If Not Me.IsEmpty Then
        Call ActivateDeactivateHeaders(isOn)
        If Me.FreezeRows = 0 Then
            Me.FreezeRows = 1
        End If
    End If

End Property

Public Property Get multiSelect() As Boolean

    multiSelect = pParamDict("Multiselect")

End Property

Public Property Let multiSelect(isOn As Boolean)
    
    pParamDict("Multiselect") = isOn

End Property

Public Property Get Events() As Object
    Set Events = pParamDict("events")
End Property

Public Property Set Events(eventParams As Variant)
    Dim dict        As Object
    Dim eventName   As String
    
    eventName = eventParams(0)
    
    If pParamDict.exists("events") And TypeName(pParamDict("events")) = "Dictionary" Then
        Set dict = pParamDict("events")
    Else
        Set dict = CreateObject("Scripting.Dictionary")
    End If
    
    dict(eventName) = eventParams
    
    Set pParamDict("events") = dict
End Property

Public Property Get AllLabels() As Collection

    Dim ResultCol As New Collection
    
    Dim Labl
    For Each Labl In pFrame.Controls
        ResultCol.Add Labl
    Next Labl
    
    Set AllLabels = ResultCol

End Property

Public Property Get RowLabels(rowNumber As Long, Optional InludingHeaders As Boolean) As Collection

    Dim ResultCol As New Collection
    
    If Not CheckInputRowNumber(rowNumber) Then Exit Property
    
    Dim Labl
    For Each Labl In pFrame.Controls
        If Not Labl.Tag = "Header" Or InludingHeaders Then
            If Split(Labl.name, ";")(0) = rowNumber Then
                ResultCol.Add Labl
            End If
        End If
    Next Labl
    
    Set RowLabels = ResultCol

End Property

Public Property Get ColumnLabels(ColumnNumber As Long, Optional InludingHeaders As Boolean) As Collection

    If Not CheckInputColumnNumber(ColumnNumber) Then Exit Property
    
    Dim ResultCol As New Collection
    
    Dim Labl
    For Each Labl In pFrame.Controls
        If Not Labl.Tag = "Header" Or InludingHeaders Then
            If Split(Labl.name, ";")(1) = ColumnNumber Then
                ResultCol.Add Labl
            End If
        End If
    Next Labl
    
    Set ColumnLabels = ResultCol

End Property

Public Property Get ExactLabel(rowNumber As Long, ColumnNumber As Long) As Object

    If Not CheckInputRowNumber(rowNumber) Then Exit Property
    If Not CheckInputColumnNumber(ColumnNumber) Then Exit Property
    
    Dim Labl
    For Each Labl In pFrame.Controls
        If Split(Labl.name, ";")(0) = rowNumber And Split(Labl.name, ";")(1) = ColumnNumber Then
            Set ExactLabel = Labl
            Exit Property
        End If
    Next Labl

End Property

Public Property Get HeadersLabels() As Collection

    Set HeadersLabels = pHeadersCol

End Property

Public Property Get ColumnSource() As Long
    
    Dim TestNumber As Long
    TestNumber = Round(pParamDict("ColumnSource"))
    
    If TestNumber < 0 Then TestNumber = 0
    If TestNumber > Me.ColumnsCount Then TestNumber = Me.ColumnsCount
    
    ColumnSource = TestNumber

End Property

Public Property Let ColumnSource(ColumnNumber As Long)

    pParamDict("ColumnSource") = Round(ColumnNumber)

End Property

Public Property Get TrueSelectedValue() As Variant

    Dim ColumnSourceNumber As Long
    ColumnSourceNumber = Me.ColumnSource
    
    Dim Labl
    For Each Labl In pFrame.Controls
        If Labl.BackColor = Me.SelectionColor And Split(Labl.name, ";")(1) = ColumnSourceNumber Then
            SelectedValue = Labl.caption
        End If
    Next Labl

End Property

Public Property Get SelectedValue() As Variant
    Dim index       As Long
    Dim result      As Variant
    Dim rows        As Variant
    Dim rowNumber As Long
    Dim ColumnSourceNumber As Long
    
    ColumnSourceNumber = Me.ColumnSource + GetDeviation("Columns")
    
    rows = Me.ListIndex
    
    rowNumber = rows(0) + GetDeviation("Rows")
    
    If rowNumber = -1 Then
        SelectedValue = Empty
        Exit Property
    End If
    
    If Not Me.multiSelect Then
        SelectedValue = Array(pInpArr(rowNumber, ColumnSourceNumber))
        Exit Property
    Else
        ReDim result(UBound(rows))
        For index = LBound(rows) To UBound(rows)
            result(index) = pInpArr(rows(index), ColumnSourceNumber)
        Next index
        
        SelectedValue = result
    End If

End Property

Public Property Get SelectionColor() As Long

    SelectionColor = pParamDict("SelectionColor")

End Property

Public Property Let SelectionColor(ColorNumber As Long)
    If Not Me.IsEmpty Then Call DeselectAllLabels
    
    pParamDict("SelectionColor") = ColorNumber

End Property


Public Property Get ListIndex() As Variant
    Dim index   As Long
    Dim result  As Variant
    
    result = Array(-1)
    
    Dim Labl
    For Each Labl In pFrame.Controls
        If Labl.BackColor = Me.SelectionColor Then
            index = Split(Labl.name, ";")(0)
            
            If Not Me.multiSelect Then
                result = Array(index)
                Exit For
            End If
            
            If result(0) = -1 Then
                result(0) = index
            Else
                If result(UBound(result)) <> index Then
                    ReDim Preserve result(UBound(result) + 1)
                    result(UBound(result)) = index
                End If
            End If
        End If
    Next Labl
    
    ListIndex = result

End Property

Public Property Let ListIndex(rowNumbers As Variant)
    Dim i           As Long
    Dim rowNumber   As Long
    
    If rowNumber = -1 Then
        Call DeselectAllLabels
    Else
        For i = LBound(rowNumbers) To UBound(rowNumbers)
            rowNumber = rowNumbers(i)
            Call Me.SelectRow(rowNumber)
        Next i
        
    End If

End Property

Public Property Get FreezeRows() As Long

    FreezeRows = pParamDict("FreezeRows")

End Property

Public Property Let FreezeRows(RowsFromTop As Long)

    RowsFromTop = Round(RowsFromTop)
    If RowsFromTop < 0 Then RowsFromTop = 0
    If RowsFromTop > Me.RowsCount Then RowsFromTop = Me.RowsCount
    
    pParamDict("FreezeRows") = RowsFromTop
    Set pFreezeRowsCol = New Collection
    
    pFrame.scroll fmScrollActionBegin, fmScrollActionBegin
    
    If RowsFromTop = 0 Then Exit Property
    
    Dim i As Long
    For i = 0 To RowsFromTop - 1
        pFreezeRowsCol.Add Me.RowLabels(i, True)
    Next i
    
    Dim Col, Labl
    
    For Each Col In pFreezeColumnsCol
        For Each Labl In Col
            Labl.ZOrder msoBringToFront
        Next Labl
    Next Col
    
    For Each Col In pFreezeRowsCol
        For Each Labl In Col
            Labl.ZOrder msoBringToFront
        Next Labl
    Next Col

End Property

Public Property Get FreezeColumns() As Long

    FreezeColumns = pParamDict("FreezeColumns")

End Property

Public Property Let FreezeColumns(ColumnsFromLeft As Long)

    ColumnsFromLeft = Round(ColumnsFromLeft)
    If ColumnsFromLeft < 0 Then ColumnsFromLeft = 0
    If ColumnsFromLeft > Me.ColumnsCount Then ColumnsFromLeft = Me.ColumnsCount
    
    pParamDict("FreezeColumns") = ColumnsFromLeft
    Set pFreezeColumnsCol = New Collection
    
    pFrame.scroll fmScrollActionBegin, fmScrollActionBegin
    
    If ColumnsFromLeft = 0 Then Exit Property
    
    Dim i As Long
    For i = 0 To ColumnsFromLeft - 1
        pFreezeColumnsCol.Add Me.ColumnLabels(i, True)
    Next i
    
    Dim Col, Labl
    
    For Each Col In pFreezeColumnsCol
        For Each Labl In Col
            Labl.ZOrder msoBringToFront
        Next Labl
    Next Col
    
    For Each Col In pFreezeRowsCol
        For Each Labl In Col
            Labl.ZOrder msoBringToFront
        Next Labl
    Next Col

End Property


' --------------------------
' Events
' --------------------------

Private Sub Class_Initialize()
    Call InitParamDict
    
    pIsEmpty = True
    Me.SelectionColor = RGB(175, 202, 251)
    Me.ColumnSource = 0
    Set pFreezeColumnsCol = New Collection
    Set pFreezeRowsCol = New Collection

End Sub

Private Sub pFrame_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, _
                          ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    Dim Col, Labl

    For Each Col In pFreezeColumnsCol
        For Each Labl In Col
            Labl.left = Labl.left + ActualDx
            If Labl.left < 0 Then Labl.left = 0
        Next Labl
    Next Col

    For Each Col In pFreezeRowsCol
        For Each Labl In Col
            Labl.top = Labl.top + ActualDy
            If Labl.top < 0 Then Labl.top = 0
        Next Labl
    Next Col

End Sub

' --------------------------
'Public Procedures
' --------------------------

Public Sub Create(Parent As UserForm, inpTop As Long, inpLeft As Long, InpHeight As Long, inpWidth As Long, InpArr As Variant, _
                  Optional scrollBars As fmScrollBars = fmScrollBarsNone, Optional listBoxNumber As String = 1)
    Dim frame       As String
    
    frame = "DesignListBox"
    If listBoxNumber <> 1 Then frame = frame & listBoxNumber
    
'    Call AddControlsToUserform(Parent, Parent, TypeName(Parent), "Frame", "DesignListBox", InpHeight, inpLeft, inpTop, inpWidth, InpParent:=True)
    Call AddControlsToUserform(Parent, Parent, TypeName(Parent), "Frame", frame, InpHeight, inpLeft, inpTop, inpWidth, InpParent:=True)
    
    If listBoxNumber = 1 Or listBoxNumber = "" Then
        Set pFrame = Parent.DesignListBox
    ElseIf listBoxNumber = 2 Then
        Set pFrame = Parent.DesignListBox2
    Else
        ' ...
    End If
    
    pFrame.scrollBars = scrollBars

    Set pParent = Parent

    Call Fill(InpArr, CStr(listBoxNumber))

End Sub

Public Sub Fill(InpArr As Variant, Optional listBoxNumber As String = "")

    If Not Right(TypeName(InpArr), 2) = "()" Then
        Err.Raise 123, "clsDesignListBox", "Input is not an Array", "Use Arrays only."
        Exit Sub
    End If
        
    Call Me.Clear
    
    'Store inpArr
    pInpArr = InpArr
    
    
    'Store Deviation
    pParamDict("DeviationRows") = LBound(InpArr, 1)
    pParamDict("DeviationColumns") = LBound(InpArr, 2)
    
    Dim ArrayDimensions As Integer
    ArrayDimensions = NumberOfArrayDimensions(InpArr)
    
    'ColumnsCount
    If ArrayDimensions = 1 Then
        pParamDict("ColumnsCount") = 1
    ElseIf ArrayDimensions = 2 Then
        pParamDict("ColumnsCount") = UBound(InpArr, 2) - LBound(InpArr, 2) + 1
    Else
        Err.Raise 123, "clsDesignListBox", "Input Array has more than 2 dimensions", "Use 1 or 2 dimension arrays."
    End If
    
    'RowsCount
    pParamDict("RowsCount") = UBound(InpArr, 1) - LBound(InpArr, 1) + 1
    
    'ColumnWidths
    If Me.ColumnWidths = vbNullString Then
        Dim strWidths As String
        Dim strWidth As String
        strWidth = CStr(Round((pFrame.width - 20) / pParamDict("ColumnsCount"), 2)) 'subtract 20 because of Scroll Bar
        Dim i As Long
        For i = 1 To pParamDict("ColumnsCount")
            strWidths = strWidths + strWidth & ";"
        Next i
        'Eliminate last ;
        If Right(strWidths, 1) = ";" Then
            strWidths = left(strWidths, Len(strWidths) - 1)
        End If
        pParamDict("ColumnWidths") = strWidths
    End If
    
    'RowHeight
    pParamDict("RowHeight") = 12
    
    Call ExecuteFill(InpArr, listBoxNumber)
 
End Sub

Public Sub Sort(ColumnNumber As Long, Optional Descending As Boolean)

    If Not CheckInputColumnNumber(ColumnNumber) Then Exit Sub
    
    'If Headers then remove Arrow Char
    If Me.Headers Then
        Call SortArray(pInpArr, LBound(pInpArr, 1) + 1, , ColumnNumber + GetDeviation("Columns"), Descending)
    Else
        Call SortArray(pInpArr, , , ColumnNumber + GetDeviation("Columns"), Descending)
    End If
    
    Dim FoundElementsDict As Object
    Dim PivotCol As Collection
    
    Set FoundElementsDict = CreateObject("Scripting.Dictionary")
    
    'Reorganize Labels
    Dim Labl, i As Long
    For Each Labl In pFrame.Controls
        If Not Labl.Tag = "Header" Then
            Labl.Tag = GetPositionInArr(Labl.caption, Split(Labl.name, ";")(1), FoundElementsDict(Labl.caption), pInpArr)
            If TypeName(FoundElementsDict(Labl.caption)) = "Collection" Then
                FoundElementsDict(Labl.caption).Add Labl.Tag
            Else
                Set PivotCol = New Collection
                PivotCol.Add Labl.Tag
                Set FoundElementsDict(Labl.caption) = PivotCol
            End If
            Labl.name = i
        Else
            'Remove or add Arrow Char
            If Right(Labl.caption, 1) = ChrW(&H2191) Or Right(Labl.caption, 1) = ChrW(&H2193) Then
                Labl.caption = left(Labl.caption, Len(Labl.caption) - 2)
            End If
            If CLng(Split(Labl.name, ";")(1)) = ColumnNumber Then
                If Descending Then
                    Labl.caption = Labl.caption & Space(1) & ChrW(&H2191)
                Else
                    Labl.caption = Labl.caption & Space(1) & ChrW(&H2193)
                End If
            End If
        End If
        i = i + 1
    Next Labl
    
    For Each Labl In pFrame.Controls
        If Not Labl.Tag = "Header" Then
            Labl.name = Labl.Tag
            Labl.top = (Split(Labl.name, ";")(0)) * Me.RowHeight
            'Debug.Assert Labl.Top <> 0
        End If
    Next Labl
    
    pFrame.Repaint

End Sub

Public Sub SelectRow(rowNumber As Long, Optional withMultiSelect As Boolean = False)

    If Not CheckInputRowNumber(rowNumber) Then Exit Sub
    
    If Not (withMultiSelect And Me.multiSelect) Then
        Call DeselectAllLabels
    End If
    
    Call SelectCorrectLabels(rowNumber)
    
    RaiseEvent Change

End Sub

Public Sub Clear()

    Dim Ctrl As Control
    For Each Ctrl In pFrame.Controls
        pFrame.Controls.Remove Ctrl.name
    Next Ctrl
    
    pIsEmpty = True

End Sub

Public Sub RaiseEventBeforeClick()

    RaiseEvent BeforeClick

End Sub

Public Sub RaiseEventClick()

    RaiseEvent Click

End Sub

' --------------------------
' Private Procedures
' --------------------------

Private Sub ExecuteFill(ByVal InpArr As Variant, Optional ByVal listBoxNumber As String = "")

    Dim i As Long, j As Long, k As Long
    Dim setTop, setLeft, setHeight, setWidth
    Dim CreatedLabel As MSForms.Label
    Dim InpBorderStyle As Long, InpIsBold As Boolean
    Dim DLBObject As clsDesignListBoxObject
    Dim outpName    As String
    
    
    If Me.Headers = True Then
        Set pHeadersCol = New Collection
    End If
    
    For i = LBound(InpArr, 1) To UBound(InpArr, 1)
        For j = LBound(InpArr, 2) To UBound(InpArr, 2)
            setTop = (i - GetDeviation("Rows")) * Me.RowHeight
            setHeight = Me.RowHeight
            setLeft = 0
            For k = LBound(InpArr, 2) To j
                If k = LBound(InpArr, 2) Then
                    setLeft = 0
                ElseIf k - GetDeviation("Columns") <= UBound(Split(Me.ColumnWidths, ";")) Then
                    setLeft = setLeft + CLng(Split(Me.ColumnWidths, ";")(k - GetDeviation("Columns") - 1))
                End If
            Next k
            If j - GetDeviation("Columns") <= UBound(Split(Me.ColumnWidths, ";")) Then
                setWidth = CLng(Split(Me.ColumnWidths, ";")(j - GetDeviation("Columns")))
            Else
                setWidth = 0
            End If
            If setWidth = 0 Then GoTo NextIteration
            If Me.Headers = True And i = LBound(InpArr, 1) Then
                InpBorderStyle = 1
                InpIsBold = True
            Else
                InpBorderStyle = 0
                InpIsBold = False
            End If
            
            If listBoxNumber = "1" Then listBoxNumber = ""
            outpName = "DesignListBox" & listBoxNumber
            
            Set CreatedLabel = AddControlsToUserform(pParent, pFrame, outpName, "Label", i - GetDeviation("Rows") & ";" & _
                                                    j - GetDeviation("Columns"), setHeight, setLeft, setTop, setWidth, InpArr(i, j), InpArr(i, j), _
                                                    InpBorderStyle:=InpBorderStyle, InpIsBold:=InpIsBold)

            CreatedLabel.Tag = outpName
            
            If Me.Headers = True And i = LBound(InpArr, 1) Then
                pHeadersCol.Add CreatedLabel
                CreatedLabel.Tag = "Header"
            End If
            
            Set DLBObject = New clsDesignListBoxObject
            DLBObject.Create CreatedLabel, Me
            pDesignListBoxObjectsCol.Add DLBObject

NextIteration:
        Next j
    Next i
    
    pIsEmpty = False
    
    If Me.Headers Then
        Me.FreezeRows = 1
    End If
    
    pFrame.ScrollWidth = setLeft + 10
    pFrame.ScrollHeight = setTop + Me.RowHeight

End Sub

Private Function GetDeviation(Mode As String) As Long
    'Mode can be "Rows" or "Columns"
    
    If Mode = "Rows" Then
        GetDeviation = pParamDict("DeviationRows")
    ElseIf Mode = "Columns" Then
        GetDeviation = pParamDict("DeviationColumns")
    End If

End Function

Private Sub ActivateDeactivateHeaders(isOn As Boolean)

    Dim Labl
    
    If isOn Then
        Set pHeadersCol = New Collection
        Set pHeadersCol = Me.RowLabels(0)
        For Each Labl In pHeadersCol
            Labl.BorderStyle = 1
            Labl.Font.Bold = True
            Labl.Tag = "Header"
        Next Labl
    Else
        For Each Labl In pHeadersCol
            Labl.BorderStyle = 0
            Labl.Font.Bold = False
            Labl.Tag = ""
        Next Labl
        Set pHeadersCol = New Collection
    End If

End Sub

Private Sub SelectCorrectLabels(rowNumber As Long)
    
    Dim Labl
    For Each Labl In pParent.Controls
        If Split(Labl.name, ";")(0) = rowNumber And InStr(Labl.Tag, pParent.ActiveControl.name) > 0 Then
            Labl.BackColor = Me.SelectionColor
        End If
    Next Labl

End Sub

Private Sub DeselectAllLabels()
    
    Dim Labl
    Dim color       As String
    
    For Each Labl In pFrame.Controls
        If Labl.BackColor = Me.SelectionColor Then
            If InStr(Labl.Tag, "color") = 0 Then
                Labl.BackColor = vbWhite
            Else
                color = Trim(Split(Labl.Tag, "color:")(1))
                Labl.BackColor = color
            End If
        End If
    Next Labl

End Sub

Private Function CheckInputRowNumber(ByRef rowNumber As Long) As Boolean

    Err.Clear
    
    rowNumber = Round(rowNumber)
    
    If Me.IsEmpty Then
        Call IsEmptyErr
    ElseIf rowNumber < 0 Then
        Err.Raise 123, "clsDesignListBox", "RowNumber must be bigger than 0"
    ElseIf rowNumber > Me.RowsCount - 1 Then
        Err.Raise 123, "clsDesignListBox", "RowNumber must be smaller than " & Me.RowsCount - 1, "Rows are zero based"
    End If
    
    If Err.Number = 0 Then
        CheckInputRowNumber = True
    End If

End Function

Private Function CheckInputColumnNumber(ByRef ColumnNumber As Long) As Boolean

    Err.Clear
    
    ColumnNumber = Round(ColumnNumber)
    
    If Me.IsEmpty Then
        Call IsEmptyErr
    ElseIf ColumnNumber < 0 Then
        Err.Raise 123, "clsDesignListBox", "ColumnNumber must be bigger than 0"
    ElseIf ColumnNumber > Me.ColumnsCount - 1 Then
        Err.Raise 123, "clsDesignListBox", "ColumnNumber must be smaller than " & Me.ColumnsCount - 1, "Columns are zero based"
    End If
    
    If Err.Number = 0 Then
        CheckInputColumnNumber = True
    End If

End Function

Private Sub IsEmptyErr()

    Err.Raise 123, "clsDesignListBox", "DesignListBox is empty", "Use .create or .fill to instantiate DesignListBox"

End Sub


' --------------------------
' Special Procedures
' --------------------------


Public Function AddControlsToUserform(UForm As UserForm, OutpLocation As Object, strOutputLocation As String, InpType, InpName, InpHeight, inpLeft, inpTop, _
                                        inpWidth, Optional InpCaption, Optional InpValue, Optional InpTextAlign, Optional InpTextSize, _
                                        Optional InpForeColor, Optional InpItems As Collection, Optional InpListIndex, Optional InpTipText, _
                                        Optional InpSpecialEffect, Optional InpBorderStyle, Optional InpVisibility As Boolean = True, _
                                        Optional InpParent As Boolean, Optional InpIsBold _
                                    ) As MSForms.Control
                                        
    'InpType can be: CheckBox, ComboBox, CommandButton,Frame,Image,Label,ListBox,MultiPage,OptionButton,ScrollBar
    'SpinButton,TabStrip,TextBox,ToggleButton
    '''''''''''''''''''''''
    'If a problem with adding objects to a frame arises (automation error), then its probably because the frame is not the last
    'created frame on the userform. The frame to fill has to be the last created (on runtime or on designtime)
    '''''''''''''''''''''''
    'ControlClass defines the WithEvents Class assigned to the control. 0 means, that no class is assigned
    
    Dim ProgId As String
    ProgId = "Forms." & InpType & ".1"
    
    Dim NewForm, ControlToEdit
    
    If Not InpParent Then
        Set NewForm = OutpLocation.Controls.Add(ProgId, InpName)
    Else
        Set NewForm = OutpLocation.Add(ProgId, InpName)
    End If
    Set ControlToEdit = NewForm
    
    With ControlToEdit
        On Error Resume Next
        .height = InpHeight
        .left = inpLeft
        .top = inpTop
        .width = inpWidth
        If Not IsMissing(InpCaption) Then
            .caption = InpCaption
        End If
        If Not IsMissing(InpValue) Then
            .value = InpValue
        End If
        If Not IsMissing(InpTextAlign) Then
            .TextAlign = InpTextAlign
        End If
        If Not IsMissing(InpTextSize) Then
            .Font.Size = InpTextSize
        End If
        If Not IsMissing(InpForeColor) Then
            .ForeColor = InpForeColor
        End If
        If Not IsMissing(InpItems) Then
            Dim Item
            For Each Item In InpItems
                .AddItem Item
            Next Item
        End If
        If Not IsMissing(InpListIndex) Then
            .ListIndex = InpListIndex
        End If
        If Not IsMissing(InpTipText) Then
            .ControlTipText = InpTipText
            .MousePointer = fmMousePointerHelp
        End If
        If Not IsMissing(InpSpecialEffect) Then
            .SpecialEffect = InpSpecialEffect
        End If
        If Not IsMissing(InpBorderStyle) Then
            .BorderStyle = InpBorderStyle
        End If
        .Font.Bold = InpIsBold
        .Visible = InpVisibility
    End With
    
    OutpLocation.Repaint
    
    Set AddControlsToUserform = ControlToEdit

End Function

Private Function GetPositionInArr(Content As Variant, ColumnNumber, LastRowInput As Variant, pInpArr As Variant) As String
    
    Dim PivotCol As Collection
    
    Dim i, Element
    For i = LBound(pInpArr, 1) To UBound(pInpArr, 1)
        If CStr(Content) = CStr(pInpArr(i, ColumnNumber + GetDeviation("Columns"))) Then
            If TypeName(LastRowInput) = "Collection" Then
                Set PivotCol = LastRowInput
                If Not IsInCollection(PivotCol, , i - GetDeviation("Rows") & ";" & ColumnNumber) Then
                    GetPositionInArr = i - GetDeviation("Rows") & ";" & ColumnNumber
                    'Debug.Assert i - GetDeviation("Rows") <> 0
                    Exit Function
                End If
            Else
                GetPositionInArr = i - GetDeviation("Rows") & ";" & ColumnNumber
                'Debug.Assert i - GetDeviation("Rows") <> 0
                Exit Function
            End If
        End If
    Next i
    
    Debug.Print "Result: None"

End Function

