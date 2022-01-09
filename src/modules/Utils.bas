Attribute VB_Name = "Utils"
'*********************************************************
' Author:   Vladimir Dmitriev
' Link:     https://github.com/dmitrievva/ShibbyGit
'*********************************************************

Option Explicit

Public Function gTranslit(iValue As String) As String
    Dim iRussian$, iCount%, iTranslit As Variant
    iRussian$ = "¿¡¬√ƒ≈®∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰Â∏ÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ"
    iTranslit = Array("", "A", "B", "V", "G", "D", "E", "Jo", "Zh", "Z", "I", "Jj", "K", "L", "M", "N", _
                        "O", "P", "R", "S", "T", "U", "F", "H", "C", "Ch", "Sh", "Zch", "''", "'Y", "'", _
                        "Eh", "Ju", "Ja", "a", "b", "v", "g", "d", "e", "jo", "zh", "z", "i", "jj", "k", _
                        "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "h", "c", "ch", "sh", "zch", _
                        "''", "'y", "'", "eh", "ju", "ja")
    For iCount% = 1 To 65
        iValue = Replace(iValue, Mid(iRussian$, iCount%, 1), iTranslit(iCount%), , , vbBinaryCompare)  'MS Excel 2000
    Next
    gTranslit = iValue
End Function


Public Function NumberOfArrayDimensions(arr As Variant) As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NumberOfArrayDimensions
    ' This function returns the number of dimensions of an array. An unallocated dynamic array
    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ndx As Integer
    Dim res As Integer
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0
    Err.Clear
    
    NumberOfArrayDimensions = Ndx - 1

End Function

Public Sub SortArray(ByRef InpArr As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0, Optional Descending As Boolean)

    If NumberOfArrayDimensions(InpArr) = 1 Then
        Call QuickSortSingleDimArray(InpArr, lngMin, lngMax)
    Else
        Call QuickSortMultiDimArray(InpArr, lngMin, lngMax, lngColumn)
    End If
    
    If Descending Then
        Call ReverseArrayInPlace(InpArr, lngMin, lngMax)
    End If

End Sub

Public Sub QuickSortSingleDimArray(ByRef InpArr As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    On Error Resume Next

    'Sort a 1-Dimensional array

    ' SampleUsage: sort arrData
    '
    '   QuickSortSingleDimArray arrData

    '
    ' Originally posted by Jim Rech 10/20/98 Excel.Programming


    ' Modifications, Nigel Heffernan:
    '       ' Escape failed comparison with an empty variant in the array
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim varX As Variant

    If TypeName(InpArr) = "Empty" Then
        Exit Sub
    End If
    If InStr(TypeName(InpArr), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(InpArr)
    End If
    If lngMax = -1 Then
        lngMax = UBound(InpArr)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = InpArr((lngMin + lngMax) \ 2)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(InpArr(n)) - varMid *might* pick up a default member or property
        i = lngMax
        j = lngMin
    ElseIf TypeName(varMid) = "Empty" Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j

        While InpArr(i) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < InpArr(j) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the item
            varX = InpArr(i)
            InpArr(i) = InpArr(j)
            InpArr(j) = varX

            i = i + 1
            j = j - 1
        End If

    Wend

    If (lngMin < j) Then Call QuickSortSingleDimArray(InpArr, lngMin, j)
    If (i < lngMax) Then Call QuickSortSingleDimArray(InpArr, i, lngMax)

End Sub

Public Sub QuickSortMultiDimArray(ByRef InpArr As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If TypeName(InpArr) = "Empty" Then
        Exit Sub
    End If
    If InStr(TypeName(InpArr), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(InpArr, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(InpArr, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = InpArr((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(InpArr(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf TypeName(varMid) = "Empty" Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While InpArr(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < InpArr(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(InpArr, 2) To UBound(InpArr, 2))
            For lngColTemp = LBound(InpArr, 2) To UBound(InpArr, 2)
                arrRowTemp(lngColTemp) = InpArr(i, lngColTemp)
                InpArr(i, lngColTemp) = InpArr(j, lngColTemp)
                InpArr(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortMultiDimArray(InpArr, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortMultiDimArray(InpArr, i, lngMax, lngColumn)
    
End Sub

Public Function ReverseArrayInPlace(InputArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, _
    Optional Reversed As Boolean = False, Optional DoubleReversed As Boolean = False) As Boolean
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ReverseArrayInPlace
    ' This procedure reverses the order of an array in place -- this is, the array variable
    ' If Reversed is true then it does reorder it in the other dimension
    ' in the calling procedure is reversed. This works only on arrays
    ' of simple data types (String, Single, Double, Integer, Long). It will not work
    ' on arrays of objects. Use ReverseArrayOfObjectsInPlace to reverse an array of objects.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim Temp As Variant
    Dim Ndx As Long
    Dim Ndx2 As Long
    Dim i
    
    If lngMin = -1 Then lngMin = LBound(InputArray, 1)
    If lngMax = -1 Then lngMax = UBound(InputArray, 1)
    
    '''''''''''''''''''''''''''''''''
    ' Set the default return value.
    '''''''''''''''''''''''''''''''''
    ReverseArrayInPlace = False
    
    Ndx2 = lngMax
    ''''''''''''''''''''''''''''''''''''''
    ' loop from the LBound of InputArray to
    ' the midpoint of InputArray
    ''''''''''''''''''''''''''''''''''''''
    If Reversed = False Then
        For Ndx = lngMin To ((lngMax - lngMin + 1) \ 2)
            For i = LBound(InputArray, 2) To UBound(InputArray, 2)
                'swap the elements
                Temp = InputArray(Ndx, i)
                InputArray(Ndx, i) = InputArray(Ndx2, i)
                InputArray(Ndx2, i) = Temp
            Next i
            ' decrement the upper index
            Ndx2 = Ndx2 - 1
        Next Ndx
        If DoubleReversed = True Then GoTo Rev
    ElseIf Reversed = True Or DoubleReversed = True Then
Rev:
        Ndx2 = UBound(InputArray, 2)
        For Ndx = LBound(InputArray, 2) To Int(((UBound(InputArray, 2) - LBound(InputArray, 2) + 1) \ 2))
            For i = LBound(InputArray, 1) To UBound(InputArray, 1)
                'swap the elements
                Temp = InputArray(i, Ndx)
                InputArray(i, Ndx) = InputArray(i, Ndx2)
                InputArray(i, Ndx2) = Temp
            Next i
            ' decrement the upper index
            Ndx2 = Ndx2 - 1
        Next Ndx
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' OK - Return True
    ''''''''''''''''''''''''''''''''''''''
    ReverseArrayInPlace = True

End Function

Public Function IsInCollection(Kollection As Collection, Optional key As Variant, Optional Item As Variant) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'INPUT       : Kollection, the collection we would like to examine
    '            : (Optional) Key, the Key we want to find in the collection
    '            : (Optional) Item, the Item we want to find in the collection
    'OUTPUT      : True if Key or Item is found, False if not
    'SPECIAL CASE: If both Key and Item are missing, return False
    
    Dim strKey As String
    Dim var As Variant

    'First, investigate assuming a Key was provided
    If Not IsMissing(key) Then
    
        strKey = CStr(key)
        
        'Handling errors is the strategy here
        On Error Resume Next
            IsInCollection = True
            var = Kollection(strKey) '<~ this is where our (potential) error will occur
            If Err.Number = 91 Then GoTo CheckForObject
            If Err.Number = 5 Then GoTo NotFound
        On Error GoTo 0
        Exit Function

CheckForObject:
        If IsObject(Kollection(strKey)) Then
            IsInCollection = True
            On Error GoTo 0
            Exit Function
        End If

NotFound:
        IsInCollection = False
        On Error GoTo 0
        Exit Function
        
    'If the Item was provided but the Key was not, then...
    ElseIf Not IsMissing(Item) Then
    
        IsInCollection = False '<~ assume that we will not find the item
    
        'We have to loop through the collection and check each item against the passed-in Item
        For Each var In Kollection
            If TypeName(var) = "clsSymbol" Then var = var.id
            If TypeName(var) = "clsDividedSymbol" Then var = var.CombinedId
            If var = Item Then
                IsInCollection = True
                Exit Function
            End If
        Next var
    
    'Otherwise, no Key OR Item was provided, so we default to False
    Else
        IsInCollection = False
    End If
    
End Function


Public Function Clipboard(Optional text) As String
    Dim v: v = text  'Cast to variant for 64-bit VBA support
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(text):
                    .SetData "text", v
                Case Else:
                    Clipboard = .GetData("text")
            End Select
        End With
    End With
End Function
