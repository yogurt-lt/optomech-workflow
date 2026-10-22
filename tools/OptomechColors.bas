Attribute VB_Name = "OptomechColors"
Option Explicit

Public Sub main()
    On Error GoTo Failed
    Dim app As SldWorks.SldWorks, model As SldWorks.ModelDoc2
    Dim assembly As SldWorks.AssemblyDoc, component As SldWorks.Component2
    Dim components As Variant, i As Long, count As Long
    Set app = Application.SldWorks
    Set model = app.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 601, , "Open the generated assembly first."
    If model.GetType <> 2 Then Err.Raise vbObjectError + 602, , "Activate the assembly document before running."
    Set assembly = model
    components = assembly.GetComponents(False)
    If IsArray(components) = False Then Err.Raise vbObjectError + 603, , "No assembly components found."
    For i = LBound(components) To UBound(components)
        Set component = components(i)
        If component.IsSuppressed = False Then
            ApplyComponentColor component, component.Name2
            count = count + 1
        End If
    Next i
    model.GraphicsRedraw2
    MsgBox "Colored " & CStr(count) & " components." & vbCrLf & "Lens: blue / Barrel: slate / Spacer: green / Retainer: orange / Adapter: violet" & vbCrLf & "Save the assembly to keep these appearance overrides.", vbInformation
    Exit Sub
Failed:
    MsgBox "Color update stopped: " & CStr(Err.Number) & " - " & Err.Description, vbExclamation
End Sub

' Shared palette. Returns a Variant containing a SAFEARRAY of Double.
Private Function ConceptPalette(ByVal label As String) As Variant
    Dim values(8) As Double, key As String
    key = LCase$(label)
    values(3) = 0.3: values(4) = 0.75: values(5) = 0.25: values(6) = 0.2
    If InStr(key, "lens") > 0 Then
        ' Lens: blue, 45 percent transparent.
        values(0) = 0.25: values(1) = 0.65: values(2) = 0.95: values(7) = 0.45
    ElseIf InStr(key, "retainer") > 0 Then
        ' Retainer: warm orange.
        values(0) = 0.95: values(1) = 0.55: values(2) = 0.20
    ElseIf InStr(key, "spacer") > 0 Then
        ' Spacer: green.
        values(0) = 0.35: values(1) = 0.72: values(2) = 0.45
    ElseIf InStr(key, "barrel") > 0 Then
        ' Barrel: dark slate gray.
        values(0) = 0.28: values(1) = 0.34: values(2) = 0.42
    ElseIf InStr(key, "adapter") > 0 Then
        ' Seat adapter: violet.
        values(0) = 0.65: values(1) = 0.48: values(2) = 0.82
    Else
        values(0) = 0.65: values(1) = 0.68: values(2) = 0.72
    End If
    ConceptPalette = values
End Function

Private Sub ApplyComponentColor(ByVal component As SldWorks.Component2, ByVal label As String)
    component.MaterialPropertyValues = ConceptPalette(label)
End Sub
