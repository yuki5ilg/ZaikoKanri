Attribute VB_Name = "mSetup"
Option Explicit

Public Sub CreateForms()
    Dim vbp As Object
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    On Error GoTo 0
    
    If vbp Is Nothing Then
        MsgBox "VBA Project access failed. Enable 'Trust VBA project object model' first.", vbExclamation
        Exit Sub
    End If
    
    ' Delete existing forms
    Dim comp As Object, killMe As New Collection
    For Each comp In vbp.VBComponents
        If comp.Name = "frmSettings" Or comp.Name = "frmInput" Then
            killMe.Add comp
        End If
    Next comp
    Dim v As Variant
    For Each v In killMe
        vbp.VBComponents.Remove v
    Next v
    
    ' Import forms from TEMP
    Dim tempFolder As String
    tempFolder = Environ("TEMP")
    
    On Error Resume Next
    vbp.VBComponents.Import tempFolder & "\frmSettings.frm"
    vbp.VBComponents.Import tempFolder & "\frmInput.frm"
    On Error GoTo 0
    
    MsgBox "Forms generated successfully." & vbCrLf & _
           "Run Main via Alt+F8", vbInformation
End Sub
