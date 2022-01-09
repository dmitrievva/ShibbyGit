Attribute VB_Name = "DocPropIO"
'***********************************************************************
' Original Author:   Eric Addison
' Link:     https://github.com/ericaddison/ShibbyGit
'
' Changed by: Vladimir Dmitriev, https://github.com/dmitrievva/ShibbyGit
'***********************************************************************

' intentionally NOT option explicit so GetDocProps works

Private docProps As Office.DocumentProperties

Public Function GetItemFromDocProperties(ByVal name As String, Optional defaultValue As Variant = "") As Variant
  Set docProps = GetDocProps
    
  On Error Resume Next
    
    Dim val As String
    val = docProps.Item(name).value
    If Err.Number <> 0 Then
        Err.Clear
        val = defaultValue
    End If

  On Error GoTo 0
  
  GetItemFromDocProperties = val
End Function


Public Function GetBooleanFromDocProperties(ByVal name As String, Optional defaultValue As Boolean = False) As Boolean
  Set docProps = GetDocProps
    
  On Error Resume Next
    
    Dim val As Boolean
    val = docProps.Item(name).value
    If Err.Number <> 0 Then
        Err.Clear
        val = defaultValue
    End If

  On Error GoTo 0
  
  GetBooleanFromDocProperties = val
End Function


Public Sub AddStringToDocProperties(ByVal name As String, ByVal value As Variant)
  Set docProps = GetDocProps
    
  On Error Resume Next
    docProps.Item(name).Delete
  On Error GoTo 0
  docProps.Add name:=name, LinkToContent:=False, value:=value, Type:=msoPropertyTypeString
  
End Sub

Public Sub AddBooleanToDocProperties(ByVal name As String, ByVal value As Boolean)
  Set docProps = GetDocProps
    
  On Error Resume Next
    docProps.Item(name).Delete
  On Error GoTo 0
  docProps.Add name:=name, LinkToContent:=False, value:=value, Type:=msoPropertyTypeBoolean
  
End Sub


Public Sub AddNumberToDocProperties(ByVal name As String, ByVal value As Variant)
  Set docProps = GetDocProps
    
  On Error Resume Next
    docProps.Item(name).Delete
  On Error GoTo 0
  docProps.Add name:=name, LinkToContent:=False, value:=value, Type:=msoPropertyTypeNumber
  
End Sub

Public Sub RemoveDocProp(ByVal name As String)
  Set docProps = GetDocProps
  On Error Resume Next
    docProps.Item(name).Delete
  On Error GoTo 0
End Sub


Public Function GetDocProps() As DocumentProperties
    Dim app As Object
    Set app = Application
    
    Select Case Application.name
        Case "Microsoft PowerPoint"
            Set GetDocProps = app.ActivePresentation.CustomDocumentProperties
        Case "Microsoft Excel"
            Set GetDocProps = app.ActiveWorkbook.CustomDocumentProperties
        Case "Microsoft Word"
            Set GetDocProps = app.ActiveDocument.CustomDocumentProperties
    End Select
    
End Function

