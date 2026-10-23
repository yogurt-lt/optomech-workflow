Attribute VB_Name = "DoubleGauss50_Manual_v02"
Option Explicit

' Data-driven editable meridional profiles; review the accompanying plan and checks.
' SaveAs options: Silent=1; Copy=2 would leave the current document unnamed.
Private Const SAVE_SILENT As Long = 1
Private app As SldWorks.SldWorks
Private buildLog As Object
Private currentStage As String

Public Sub main()
    On Error GoTo Failed
    Dim model As SldWorks.ModelDoc2, assembly As SldWorks.AssemblyDoc
    Dim component As SldWorks.Component2, active As SldWorks.ModelDoc2
    Dim partModel As SldWorks.ModelDoc2, math As SldWorks.MathUtility
    Dim move As SldWorks.MathTransform, readback As SldWorks.MathTransform
    Dim fs As Object, log As Object, profiles As Variant, names As Variant, row As Variant
    Dim paths() As String, zpos() As Double
    Dim folder As String, output As String, template As String, path As String, runID As String
    Dim matrix(15) As Double, actual As Variant, i As Long, j As Long, status As Long
    Dim loadErrors As Long, loadWarnings As Long, selected As Boolean, stage As String
    Dim loaded(24) As SldWorks.ModelDoc2
    Set app = Application.SldWorks
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = fs.GetParentFolderName(app.GetCurrentMacroPathName)
    If Len(folder) = 0 Then Err.Raise vbObjectError + 400, , "Save the macro before running."
    output = folder & "\Optomech_" & Format$(Now, "yyyymmdd_hhnnss")
    runID = "DG50_" & Format$(Now, "yyyymmdd_hhnnss")
    fs.CreateFolder output
    Set log = fs.CreateTextFile(output & "\build_log.txt", False, True)
    Set buildLog = log
    log.WriteLine "Optomech workflow concept; SOLIDWORKS " & app.RevisionNumber
    names = Array("01_OpticalCartridge", "02_FrontSpacer", "03_FrontRetainer", "04_RearSpacer34", "05_RearSpacer45", "06_RearRetainer", "07_IntegratedHousing", "08_ShortFocusRing", "09_CouplingKeeper", "10_FrontThrustWasher", "11_RearThrustWasher", "12_EFMountInterfaceEnvelope", "Lens1_RevisedBlank_Reference", "Lens2_RevisedBlank_Reference", "Lens3_RevisedBlank_Reference", "Lens4_RevisedBlank_Reference", "Lens5_RevisedBlank_Reference", "Lens6_RevisedBlank_Reference", "Lens7_RevisedBlank_Reference", "21_GuideDogScrew", "22_GuideDogScrew", "31_MountScrew", "32_MountScrew", "33_MountScrew", "40_KeeperLockScrew")
    ReDim paths(UBound(names)), zpos(UBound(names))
    For i = 0 To UBound(names)
        profiles = ProfileFor(i)
        paths(i) = output & "\" & runID & "_" & names(i) & ".SLDPRT"
        zpos(i) = 0#
        stage = "Build " & names(i)
        Set loaded(i) = MakeProfile(profiles, CStr(names(i)), paths(i))
        Set partModel = loaded(i)
    Next i
    stage = "Create assembly"
    template = app.GetDocumentTemplate(2, "", 0, 0#, 0#)
    If Len(template) = 0 Then Err.Raise vbObjectError + 302, , "Configure a default ASSEMBLY template. Native parts are already saved."
    Set model = app.NewDocument(template, 0, 0#, 0#)
    If model Is Nothing Then Err.Raise vbObjectError + 303, , "Assembly creation failed. Parts remain saved."
    Set assembly = model
    Set math = app.GetMathUtility
    For i = 0 To UBound(names)
        stage = "Preload " & names(i)
        Set partModel = loaded(i)
        If partModel Is Nothing Then Set partModel = app.OpenDoc6(paths(i), 1, 1, "", loadErrors, loadWarnings)
        log.WriteLine "Preload " & names(i) & "; errors=" & CStr(loadErrors) & "; warnings=" & CStr(loadWarnings)
        If partModel Is Nothing Then Err.Raise vbObjectError + 314, , "Part preload failed: " & names(i) & "; errors=" & CStr(loadErrors)
        If StrComp(partModel.GetPathName, paths(i), vbTextCompare) <> 0 Then Err.Raise vbObjectError + 421, , "Preloaded part path mismatch."
        stage = "Insert " & names(i)
        log.WriteLine stage & ": activate assembly"
1000    stage = "Activate assembly for " & names(i)
        Set active = app.ActivateDoc3(model.GetTitle, False, 0, status)
        If active Is Nothing Then Err.Raise vbObjectError + 304, , "Assembly activation failed."
        ' swAddComponentConfigOptions_CurrentSelectedConfig = 0 (2023 API).
        stage = "AddComponent5 " & names(i)
1010    Set component = assembly.AddComponent5(paths(i), 0, "", False, "", 0#, 0#, 0#)
        If component Is Nothing Then Err.Raise vbObjectError + 305, , "Component insertion failed: " & names(i)
        log.WriteLine "Inserted " & names(i)
        ApplyComponentColor component, CStr(names(i))
        stage = "Select component " & names(i)
        model.ClearSelection2 True
1020    selected = component.Select4(False, Nothing, False)
        If selected = False Then Err.Raise vbObjectError + 306, , "Component selection failed: " & names(i)
        stage = "Unfix component " & names(i)
1030    assembly.UnfixComponent
        For j = 0 To 15: matrix(j) = 0#: Next j
        matrix(0) = 1#: matrix(4) = 1#: matrix(8) = 1#: matrix(12) = 1#
        actual = FocusPlacement(CStr(names(i)), 0#)
        For j = 0 To 15: matrix(j) = CDbl(actual(j)): Next j
        stage = "Create transform " & names(i)
1040    Set move = math.CreateTransform(matrix)
        If move Is Nothing Then Err.Raise vbObjectError + 315, , "Transform creation failed."
        stage = "Assign Transform2 " & names(i)
        ' Transform2's COM setter uses Property Let in SOLIDWORKS VBA.
        ' Do not use Set component.Transform2 = move (Property Set).
1050    component.Transform2 = move
        stage = "Read back transform " & names(i)
1060    Set readback = component.Transform2
        If readback Is Nothing Then Err.Raise vbObjectError + 316, , "Transform readback failed."
1070    actual = readback.ArrayData
        For j = 0 To 12
            If Abs(actual(j) - matrix(j)) > 0.00000001 Then Err.Raise vbObjectError + 307, , "Placement verification failed: " & names(i)
        Next j
        stage = "Fix component " & names(i)
1080    assembly.FixComponent
        log.WriteLine names(i) & " saved and positioned at Z_mm=" & CStr(zpos(i))
    Next i
    model.ClearSelection2 True
    model.EditRebuild3
    path = output & "\Optomech_Assembly_" & Format$(Now, "yyyymmdd_hhnnss") & ".SLDASM"
    stage = "Save assembly"
    log.WriteLine "Save requested: " & path
    log.WriteLine "Path before save: " & model.GetPathName
    status = model.SaveAs3(path, 0, SAVE_SILENT)
    log.WriteLine "Save status: " & CStr(status)
    log.WriteLine "Path after save: " & model.GetPathName
    If status <> 0 Then Err.Raise vbObjectError + 308, , "Assembly save status=" & CStr(status)
    If StrComp(model.GetPathName, path, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 422, , "Saved assembly path differs from requested output."
    If fs.FileExists(path) = False Then Err.Raise vbObjectError + 423, , "Saved assembly file was not found."
    log.WriteLine "SUCCESS: all planned parts assembled."
    log.WriteLine "Components fixed at verified coordinates; no assembly mates added."
    log.WriteLine "Review source design plan; nominal thread envelopes only; optical and manufacturing verification pending."
    log.Close
    Set log = Nothing
    model.ShowNamedView2 "", 7
    model.ViewZoomtofit2
    MsgBox "Optomech models generated and saved:" & vbCrLf & path, vbInformation
    Exit Sub

Failed:
    Dim message As String
    message = "Optomech build stopped: " & Err.Number & " - " & Err.Description _
        & vbCrLf & "Stage: " & stage & vbCrLf & "Detail: " & currentStage & vbCrLf & "Line: " & CStr(Erl)
    On Error Resume Next
    If Not log Is Nothing Then
        log.WriteLine message
        log.Close
    End If
    MsgBox message, vbExclamation
End Sub

Private Function MakeProfile(ByVal points As Variant, ByVal label As String, ByVal target As String) As SldWorks.ModelDoc2
    Dim model As SldWorks.ModelDoc2, part As SldWorks.PartDoc
    Dim feat As SldWorks.Feature, plane As SldWorks.Feature, profile As SldWorks.Feature
    Dim ref As SldWorks.RefPlane, data As Variant, bodies As Variant
    Dim planeTransform As SldWorks.MathTransform, props As SldWorks.CustomPropertyManager
    Dim sketch As SldWorks.Sketch, mapping As SldWorks.MathTransform
    Dim math As SldWorks.MathUtility, seg As SldWorks.SketchSegment
    Dim axis As SldWorks.SketchSegment, revolve As SldWorks.Feature
    Dim selectData As SldWorks.SelectData, template As String
    Dim a As Variant, b As Variant, p As Variant, q As Variant
    Dim i As Long, j As Long, count As Long, status As Long
    Dim x As Double, y As Double, z As Double
    Dim zmin As Double, zmax As Double, expectedMin As Double, expectedMax As Double
    currentStage = label & ": new part"
    template = app.GetDocumentTemplate(1, "", 0, 0#, 0#)
    If Len(template) = 0 Then Err.Raise vbObjectError + 401, , "Default part template missing."
    Set model = app.NewDocument(template, 0, 0#, 0#)
    If model Is Nothing Then Err.Raise vbObjectError + 402, , "Part creation failed."
    Set part = model
    Set math = app.GetMathUtility
    currentStage = label & ": find XZ plane"
    Set feat = model.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = "RefPlane" Then
            Set ref = feat.GetSpecificFeature2
            Set planeTransform = ref.Transform
            data = planeTransform.ArrayData
            If Abs(data(6)) < 0.0000001 And Abs(Abs(data(7)) - 1#) < 0.0000001 _
                And Abs(data(8)) < 0.0000001 And Abs(data(9)) < 0.0000001 _
                And Abs(data(10)) < 0.0000001 And Abs(data(11)) < 0.0000001 Then
                Set plane = feat
                Exit Do
            End If
        End If
        Set feat = feat.GetNextFeature
    Loop
    If plane Is Nothing Then Err.Raise vbObjectError + 403, , "No origin XZ plane found."
    model.ClearSelection2 True
    If model.Extension.SelectByID2(plane.Name, "PLANE", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then Err.Raise vbObjectError + 404, , "XZ plane selection failed."
    model.SketchManager.InsertSketch True
    Set sketch = model.SketchManager.ActiveSketch
    If sketch Is Nothing Then Err.Raise vbObjectError + 405, , "Profile sketch creation failed."
    Set mapping = sketch.ModelToSketchTransform
    model.SketchManager.AddToDB = True
    currentStage = label & ": draw meridional profile"
    expectedMin = 100000#: expectedMax = -100000#
    For i = LBound(points) To UBound(points)
        p = points(i)
        j = i + 1
        If j > UBound(points) Then j = LBound(points)
        q = points(j)
        a = SketchPoint(math, mapping, CDbl(p(0)), CDbl(p(1)))
        b = SketchPoint(math, mapping, CDbl(q(0)), CDbl(q(1)))
        Set seg = model.SketchManager.CreateLine(a(0), a(1), 0#, b(0), b(1), 0#)
        If seg Is Nothing Then Err.Raise vbObjectError + 406, , "Profile line failed."
        If p(0) < expectedMin Then expectedMin = p(0)
        If p(0) > expectedMax Then expectedMax = p(0)
    Next i
    a = SketchPoint(math, mapping, -60.0, 0#)
    b = SketchPoint(math, mapping, 15.0, 0#)
    Set axis = model.SketchManager.CreateCenterLine(a(0), a(1), 0#, b(0), b(1), 0#)
    If axis Is Nothing Then Err.Raise vbObjectError + 407, , "Revolution centerline failed."
    model.SketchManager.AddToDB = False
    model.SketchManager.InsertSketch True
    Set feat = model.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = "ProfileFeature" Then
            Set profile = feat
            count = count + 1
        End If
        Set feat = feat.GetNextFeature
    Loop
    If count <> 1 Then Err.Raise vbObjectError + 408, , "Expected one unabsorbed profile sketch."
    profile.Name = "Meridional_Profile"
    currentStage = label & ": select profile and axis"
    model.ClearSelection2 True
    If profile.Select2(False, 0) = False Then Err.Raise vbObjectError + 409, , "Profile selection failed."
    Set selectData = model.SelectionManager.CreateSelectData
    selectData.Mark = 4
    If axis.Select4(True, selectData) = False Then Err.Raise vbObjectError + 410, , "Axis selection failed."
    currentStage = label & ": revolve"
2000 Set revolve = model.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, _
        0, 0, 6.283185307179586, 0#, False, False, 0#, 0#, 0, 0#, 0#, True, True, True)
    If revolve Is Nothing Then Err.Raise vbObjectError + 411, , "Profile revolution failed."
    revolve.Name = "Editable_Revolve"
    AddFocusDetails model, label
    currentStage = label & ": check geometry"
    model.ClearSelection2 True
    model.EditRebuild3
    bodies = part.GetBodies2(0, False)
    If IsArray(bodies) = False Then Err.Raise vbObjectError + 412, , "No solid generated."
    If UBound(bodies) - LBound(bodies) <> 0 Then Err.Raise vbObjectError + 413, , "Expected one solid."
    Dim body As SldWorks.Body2
    Set body = bodies(LBound(bodies))
    If body.GetExtremePoint(0#, 0#, -1#, x, y, z) = False Then Err.Raise vbObjectError + 414, , "Z-min check failed."
    zmin = z * 1000#
    If body.GetExtremePoint(0#, 0#, 1#, x, y, z) = False Then Err.Raise vbObjectError + 415, , "Z-max check failed."
    zmax = z * 1000#
    If Abs(zmin - expectedMin) > 0.001 Or Abs(zmax - expectedMax) > 0.001 Then Err.Raise vbObjectError + 416, , "Revolve orientation / Z range differs from design."
    currentStage = label & ": properties and stop check"
    Set props = model.Extension.CustomPropertyManager("")
    props.Add3 "DESIGN_STATUS", 30, "Engineering concept; review design plan and verification notes", 2
    props.Add3 "THREADS", 30, "Envelopes only; thread geometry and fits require design", 2
    ApplyConceptColor model, label
    currentStage = label & ": save"
    buildLog.WriteLine "Save requested: " & target
    buildLog.WriteLine "Path before save: " & model.GetPathName
    status = model.SaveAs3(target, 0, SAVE_SILENT)
    buildLog.WriteLine "Save status: " & CStr(status)
    buildLog.WriteLine "Path after save: " & model.GetPathName
    If status <> 0 Then Err.Raise vbObjectError + 417, , "Part save failed: " & CStr(status)
    If StrComp(model.GetPathName, target, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 420, , "Saved document path differs from requested output."
    buildLog.WriteLine label & "; solid=1; z_mm=" & CStr(zmin) & ".." & CStr(zmax)
    Set MakeProfile = model
End Function



Private Function SketchPoint(ByVal math As SldWorks.MathUtility, ByVal mapping As SldWorks.MathTransform, ByVal axialMM As Double, ByVal radiusMM As Double) As Variant
    Dim xyz(2) As Double, point As SldWorks.MathPoint
    xyz(0) = radiusMM / 1000#: xyz(1) = 0#: xyz(2) = axialMM / 1000#
    Set point = math.CreatePoint(xyz)
    Set point = point.MultiplyTransform(mapping)
    SketchPoint = point.ArrayData
End Function

' Shared palette. Returns a Variant containing a SAFEARRAY of Double.
Private Function ConceptPalette(ByVal label As String) As Variant
    Dim values(8) As Double, key As String
    key = LCase$(label)
    values(3) = 0.3: values(4) = 0.75: values(5) = 0.25: values(6) = 0.2
    If InStr(key, "01_opticalcartridge") > 0 Then
        values(0) = 0.2: values(1) = 0.34: values(2) = 0.68
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "02_frontspacer") > 0 Then
        values(0) = 0.27: values(1) = 0.59: values(2) = 0.35
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "03_frontretainer") > 0 Then
        values(0) = 0.25: values(1) = 0.28: values(2) = 0.33
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "04_rearspacer34") > 0 Then
        values(0) = 0.27: values(1) = 0.59: values(2) = 0.35
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "05_rearspacer45") > 0 Then
        values(0) = 0.27: values(1) = 0.59: values(2) = 0.35
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "06_rearretainer") > 0 Then
        values(0) = 0.25: values(1) = 0.28: values(2) = 0.33
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "07_integratedhousing") > 0 Then
        values(0) = 0.06: values(1) = 0.52: values(2) = 0.55
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "08_shortfocusring") > 0 Then
        values(0) = 0.82: values(1) = 0.51: values(2) = 0.16
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "09_couplingkeeper") > 0 Then
        values(0) = 0.25: values(1) = 0.28: values(2) = 0.33
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "10_frontthrustwasher") > 0 Then
        values(0) = 0.8: values(1) = 0.81: values(2) = 0.75
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "11_rearthrustwasher") > 0 Then
        values(0) = 0.8: values(1) = 0.81: values(2) = 0.75
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "12_efmountinterfaceenvelope") > 0 Then
        values(0) = 0.62: values(1) = 0.65: values(2) = 0.7
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "lens1_revisedblank_reference") > 0 Then
        values(0) = 0.24: values(1) = 0.63: values(2) = 0.86
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "lens2_revisedblank_reference") > 0 Then
        values(0) = 0.24: values(1) = 0.63: values(2) = 0.86
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "lens3_revisedblank_reference") > 0 Then
        values(0) = 0.24: values(1) = 0.63: values(2) = 0.86
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "lens4_revisedblank_reference") > 0 Then
        values(0) = 0.24: values(1) = 0.63: values(2) = 0.86
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "lens5_revisedblank_reference") > 0 Then
        values(0) = 0.24: values(1) = 0.63: values(2) = 0.86
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "lens6_revisedblank_reference") > 0 Then
        values(0) = 0.24: values(1) = 0.63: values(2) = 0.86
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "lens7_revisedblank_reference") > 0 Then
        values(0) = 0.24: values(1) = 0.63: values(2) = 0.86
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "21_guidedogscrew") > 0 Then
        values(0) = 0.62: values(1) = 0.65: values(2) = 0.7
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "22_guidedogscrew") > 0 Then
        values(0) = 0.62: values(1) = 0.65: values(2) = 0.7
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "31_mountscrew") > 0 Then
        values(0) = 0.62: values(1) = 0.65: values(2) = 0.7
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "32_mountscrew") > 0 Then
        values(0) = 0.62: values(1) = 0.65: values(2) = 0.7
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "33_mountscrew") > 0 Then
        values(0) = 0.62: values(1) = 0.65: values(2) = 0.7
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "40_keeperlockscrew") > 0 Then
        values(0) = 0.62: values(1) = 0.65: values(2) = 0.7
        ConceptPalette = values
        Exit Function
    End If
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

Private Sub ApplyConceptColor(ByVal model As SldWorks.ModelDoc2, ByVal label As String)
    On Error Resume Next
    model.MaterialPropertyValues = ConceptPalette(label)
    On Error GoTo 0
End Sub

Private Function ProfileFor(ByVal index As Long) As Variant
    Select Case index
        Case 0: ProfileFor = Profile_0()
        Case 1: ProfileFor = Profile_1()
        Case 2: ProfileFor = Profile_2()
        Case 3: ProfileFor = Profile_3()
        Case 4: ProfileFor = Profile_4()
        Case 5: ProfileFor = Profile_5()
        Case 6: ProfileFor = Profile_6()
        Case 7: ProfileFor = Profile_7()
        Case 8: ProfileFor = Profile_8()
        Case 9: ProfileFor = Profile_9()
        Case 10: ProfileFor = Profile_10()
        Case 11: ProfileFor = Profile_11()
        Case 12: ProfileFor = Profile_12()
        Case 13: ProfileFor = Profile_13()
        Case 14: ProfileFor = Profile_14()
        Case 15: ProfileFor = Profile_15()
        Case 16: ProfileFor = Profile_16()
        Case 17: ProfileFor = Profile_17()
        Case 18: ProfileFor = Profile_18()
        Case 19: ProfileFor = Profile_19()
        Case 20: ProfileFor = Profile_20()
        Case 21: ProfileFor = Profile_21()
        Case 22: ProfileFor = Profile_22()
        Case 23: ProfileFor = Profile_23()
        Case 24: ProfileFor = Profile_24()
        Case Else: Err.Raise vbObjectError + 418, , "Unknown profile."
    End Select
End Function

Private Function Profile_0() As Variant
    Dim pts(33) As Variant
    pts(0) = Array(-47.000000000000, 25.000000000000)
    pts(1) = Array(-38.900000000000, 25.000000000000)
    pts(2) = Array(-38.900000000000, 29.000000000000)
    pts(3) = Array(-37.500000000000, 29.000000000000)
    pts(4) = Array(-37.500000000000, 27.000000000000)
    pts(5) = Array(-25.000000000000, 27.000000000000)
    pts(6) = Array(-25.000000000000, 26.600000000000)
    pts(7) = Array(-12.000000000000, 26.600000000000)
    pts(8) = Array(-12.000000000000, 27.000000000000)
    pts(9) = Array(-5.000000000000, 27.000000000000)
    pts(10) = Array(-5.000000000000, 25.800000000000)
    pts(11) = Array(12.000000000000, 25.800000000000)
    pts(12) = Array(12.000000000000, 24.080000000000)
    pts(13) = Array(8.789581280691, 24.080000000000)
    pts(14) = Array(8.789581280691, 23.080000000000)
    pts(15) = Array(-2.032163854282, 23.080000000000)
    pts(16) = Array(-2.032163854282, 21.080000000000)
    pts(17) = Array(-8.177115975542, 21.080000000000)
    pts(18) = Array(-8.177115975542, 16.580000000000)
    pts(19) = Array(-15.633209860706, 16.580000000000)
    pts(20) = Array(-15.633209860706, 15.900000000000)
    pts(21) = Array(-18.080000000000, 15.900000000000)
    pts(22) = Array(-18.080000000000, 12.550000000000)
    pts(23) = Array(-21.080000000000, 9.550000000000)
    pts(24) = Array(-22.080000000000, 9.550000000000)
    pts(25) = Array(-25.080000000000, 12.550000000000)
    pts(26) = Array(-25.080000000000, 15.900000000000)
    pts(27) = Array(-26.588086491670, 15.900000000000)
    pts(28) = Array(-26.588086491670, 18.080000000000)
    pts(29) = Array(-33.826479267413, 18.080000000000)
    pts(30) = Array(-33.826479267413, 21.580000000000)
    pts(31) = Array(-38.700444021751, 21.580000000000)
    pts(32) = Array(-38.700444021751, 23.080000000000)
    pts(33) = Array(-47.000000000000, 23.080000000000)
    Profile_0 = pts
End Function

Private Function Profile_1() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(-37.522729699984, 21.450000000000)
    pts(1) = Array(-34.126479267413, 21.450000000000)
    pts(2) = Array(-34.126479267413, 17.100000000000)
    pts(3) = Array(-35.126479267413, 17.100000000000)
    pts(4) = Array(-36.522729699984, 20.500000000000)
    pts(5) = Array(-37.522729699984, 20.500000000000)
    Profile_1 = pts
End Function

Private Function Profile_2() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-41.700444021751, 23.000000000000)
    pts(1) = Array(-38.700444021751, 23.000000000000)
    pts(2) = Array(-38.700444021751, 20.500000000000)
    pts(3) = Array(-41.700444021751, 20.500000000000)
    Profile_2 = pts
End Function

Private Function Profile_3() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(-7.877115975542, 21.000000000000)
    pts(1) = Array(-2.558814296742, 21.000000000000)
    pts(2) = Array(-2.558814296742, 20.250000000000)
    pts(3) = Array(-4.058814296742, 20.250000000000)
    pts(4) = Array(-6.377115975542, 16.000000000000)
    pts(5) = Array(-7.877115975542, 16.000000000000)
    Profile_3 = pts
End Function

Private Function Profile_4() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(-1.732163854282, 23.000000000000)
    pts(1) = Array(7.762586533956, 23.000000000000)
    pts(2) = Array(7.762586533956, 22.200000000000)
    pts(3) = Array(6.262586533956, 22.200000000000)
    pts(4) = Array(-0.232163854282, 20.250000000000)
    pts(5) = Array(-1.732163854282, 20.250000000000)
    Profile_4 = pts
End Function

Private Function Profile_5() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(8.789581280691, 24.000000000000)
    pts(1) = Array(11.289581280691, 24.000000000000)
    pts(2) = Array(11.289581280691, 22.200000000000)
    pts(3) = Array(8.789581280691, 22.200000000000)
    Profile_5 = pts
End Function

Private Function Profile_6() As Variant
    Dim pts(9) As Variant
    pts(0) = Array(-35.000000000000, 30.300000000000)
    pts(1) = Array(-28.500000000000, 30.300000000000)
    pts(2) = Array(-28.500000000000, 31.000000000000)
    pts(3) = Array(-12.000000000000, 31.000000000000)
    pts(4) = Array(-12.000000000000, 30.300000000000)
    pts(5) = Array(-9.000000000000, 30.300000000000)
    pts(6) = Array(-9.000000000000, 32.000000000000)
    pts(7) = Array(-2.000000000000, 32.000000000000)
    pts(8) = Array(-2.000000000000, 27.060000000000)
    pts(9) = Array(-35.000000000000, 27.060000000000)
    Profile_6 = pts
End Function

Private Function Profile_7() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(-41.200000000000, 35.500000000000)
    pts(1) = Array(-16.500000000000, 35.500000000000)
    pts(2) = Array(-16.500000000000, 30.500000000000)
    pts(3) = Array(-39.200000000000, 30.500000000000)
    pts(4) = Array(-39.200000000000, 25.200000000000)
    pts(5) = Array(-41.200000000000, 25.200000000000)
    Profile_7 = pts
End Function

Private Function Profile_8() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-44.000000000000, 29.000000000000)
    pts(1) = Array(-41.500000000000, 29.000000000000)
    pts(2) = Array(-41.500000000000, 25.040000000000)
    pts(3) = Array(-44.000000000000, 25.040000000000)
    Profile_8 = pts
End Function

Private Function Profile_9() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-41.480000000000, 29.200000000000)
    pts(1) = Array(-41.220000000000, 29.200000000000)
    pts(2) = Array(-41.220000000000, 25.100000000000)
    pts(3) = Array(-41.480000000000, 25.100000000000)
    Profile_9 = pts
End Function

Private Function Profile_10() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-39.180000000000, 29.200000000000)
    pts(1) = Array(-38.920000000000, 29.200000000000)
    pts(2) = Array(-38.920000000000, 25.100000000000)
    pts(3) = Array(-39.180000000000, 25.100000000000)
    Profile_10 = pts
End Function

Private Function Profile_11() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.000000000000, 32.000000000000)
    pts(1) = Array(0.000000000000, 32.000000000000)
    pts(2) = Array(0.000000000000, 26.200000000000)
    pts(3) = Array(-2.000000000000, 26.200000000000)
    Profile_11 = pts
End Function

Private Function Profile_12() As Variant
    Dim pts(153) As Variant
    pts(0) = Array(-45.710000000000, 0.000000000000)
    pts(1) = Array(-45.709299698667, 0.215000000000)
    pts(2) = Array(-45.707646150441, 0.394166666667)
    pts(3) = Array(-45.701420284488, 0.752500000000)
    pts(4) = Array(-45.695816000548, 0.967500000000)
    pts(5) = Array(-45.690074514158, 1.146666666667)
    pts(6) = Array(-45.675667741729, 1.505000000000)
    pts(7) = Array(-45.665150717214, 1.720000000000)
    pts(8) = Array(-45.655312314102, 1.899166666667)
    pts(9) = Array(-45.632702078253, 2.257500000000)
    pts(10) = Array(-45.617255821527, 2.472500000000)
    pts(11) = Array(-45.603305055109, 2.651666666667)
    pts(12) = Array(-45.572455787018, 3.010000000000)
    pts(13) = Array(-45.552055942837, 3.225000000000)
    pts(14) = Array(-45.533970776543, 3.404166666667)
    pts(15) = Array(-45.494833613747, 3.762500000000)
    pts(16) = Array(-45.469447769580, 3.977500000000)
    pts(17) = Array(-45.447199401094, 4.156666666667)
    pts(18) = Array(-45.399711796001, 4.515000000000)
    pts(19) = Array(-45.369299220991, 4.730000000000)
    pts(20) = Array(-45.342851849388, 4.909166666667)
    pts(21) = Array(-45.286937060649, 5.267500000000)
    pts(22) = Array(-45.226978080395, 5.625833333333)
    pts(23) = Array(-45.195475080401, 5.805000000000)
    pts(24) = Array(-45.162952495660, 5.984166666667)
    pts(25) = Array(-45.122575214524, 6.199166666667)
    pts(26) = Array(-45.087798589468, 6.378333333333)
    pts(27) = Array(-45.051992174456, 6.557500000000)
    pts(28) = Array(-45.015152545714, 6.736666666667)
    pts(29) = Array(-44.977276164812, 6.915833333333)
    pts(30) = Array(-44.938359376875, 7.095000000000)
    pts(31) = Array(-44.898398408717, 7.274166666667)
    pts(32) = Array(-44.857389366903, 7.453333333333)
    pts(33) = Array(-44.806789396663, 7.668333333333)
    pts(34) = Array(-44.763460282281, 7.847500000000)
    pts(35) = Array(-44.719069803163, 8.026666666667)
    pts(36) = Array(-44.673613537853, 8.205833333333)
    pts(37) = Array(-44.627086933248, 8.385000000000)
    pts(38) = Array(-44.579485302137, 8.564166666667)
    pts(39) = Array(-44.530803820640, 8.743333333333)
    pts(40) = Array(-44.481037525548, 8.922500000000)
    pts(41) = Array(-44.430181311562, 9.101666666667)
    pts(42) = Array(-44.378229928423, 9.280833333333)
    pts(43) = Array(-44.325177977930, 9.460000000000)
    pts(44) = Array(-44.271019910848, 9.639166666667)
    pts(45) = Array(-44.215750023692, 9.818333333333)
    pts(46) = Array(-44.159362455388, 9.997500000000)
    pts(47) = Array(-44.101851183804, 10.176666666667)
    pts(48) = Array(-44.043210022148, 10.355833333333)
    pts(49) = Array(-43.983432615226, 10.535000000000)
    pts(50) = Array(-43.922512435551, 10.714166666667)
    pts(51) = Array(-43.860442779303, 10.893333333333)
    pts(52) = Array(-43.797216762134, 11.072500000000)
    pts(53) = Array(-43.732827314797, 11.251666666667)
    pts(54) = Array(-43.667267178614, 11.430833333333)
    pts(55) = Array(-43.546284881232, 11.753333333333)
    pts(56) = Array(-43.435503998600, 12.040000000000)
    pts(57) = Array(-43.321638240757, 12.326666666667)
    pts(58) = Array(-43.204653055868, 12.613333333333)
    pts(59) = Array(-43.084512405975, 12.900000000000)
    pts(60) = Array(-42.961178700880, 13.186666666667)
    pts(61) = Array(-42.834612727586, 13.473333333333)
    pts(62) = Array(-42.704773574931, 13.760000000000)
    pts(63) = Array(-42.571618553038, 14.046666666667)
    pts(64) = Array(-42.435103107150, 14.333333333333)
    pts(65) = Array(-42.295180725393, 14.620000000000)
    pts(66) = Array(-42.151802839944, 14.906666666667)
    pts(67) = Array(-42.023472772222, 15.157500000000)
    pts(68) = Array(-41.873477298735, 15.444166666667)
    pts(69) = Array(-41.739273859487, 15.695000000000)
    pts(70) = Array(-41.582467002657, 15.981666666667)
    pts(71) = Array(-41.442213385726, 16.232500000000)
    pts(72) = Array(-41.278384370907, 16.519166666667)
    pts(73) = Array(-41.131888130364, 16.770000000000)
    pts(74) = Array(-40.960807347108, 17.056666666667)
    pts(75) = Array(-40.807858652943, 17.307500000000)
    pts(76) = Array(-40.629275511758, 17.594166666667)
    pts(77) = Array(-40.469645131718, 17.845000000000)
    pts(78) = Array(-40.306812161460, 18.095833333333)
    pts(79) = Array(-40.140719576154, 18.346666666667)
    pts(80) = Array(-39.971307805891, 18.597500000000)
    pts(81) = Array(-39.798514585315, 18.848333333333)
    pts(82) = Array(-39.622274791264, 19.099166666667)
    pts(83) = Array(-39.442520267229, 19.350000000000)
    pts(84) = Array(-39.259179633311, 19.600833333333)
    pts(85) = Array(-39.072178080175, 19.851666666667)
    pts(86) = Array(-38.881437145346, 20.102500000000)
    pts(87) = Array(-38.700444021751, 20.336000000000)
    pts(88) = Array(-38.700444021751, 21.500000000000)
    pts(89) = Array(-37.522729699984, 21.500000000000)
    pts(90) = Array(-37.522729699984, 20.336000000000)
    pts(91) = Array(-37.607387697801, 19.995000000000)
    pts(92) = Array(-37.686056540173, 19.672500000000)
    pts(93) = Array(-37.763372051617, 19.350000000000)
    pts(94) = Array(-37.839338083655, 19.027500000000)
    pts(95) = Array(-37.913958406364, 18.705000000000)
    pts(96) = Array(-37.987236709375, 18.382500000000)
    pts(97) = Array(-38.059176602844, 18.060000000000)
    pts(98) = Array(-38.137544371048, 17.701666666667)
    pts(99) = Array(-38.206670236622, 17.379166666667)
    pts(100) = Array(-38.274468426209, 17.056666666667)
    pts(101) = Array(-38.340942232927, 16.734166666667)
    pts(102) = Array(-38.413252641367, 16.375833333333)
    pts(103) = Array(-38.476941007119, 16.053333333333)
    pts(104) = Array(-38.539314759975, 15.730833333333)
    pts(105) = Array(-38.600376890306, 15.408333333333)
    pts(106) = Array(-38.666688926631, 15.050000000000)
    pts(107) = Array(-38.724991574721, 14.727500000000)
    pts(108) = Array(-38.781991448637, 14.405000000000)
    pts(109) = Array(-38.837691248710, 14.082500000000)
    pts(110) = Array(-38.898058335870, 13.724166666667)
    pts(111) = Array(-38.951022094581, 13.401666666667)
    pts(112) = Array(-39.002693743418, 13.079166666667)
    pts(113) = Array(-39.053075703931, 12.756666666667)
    pts(114) = Array(-39.107545925101, 12.398333333333)
    pts(115) = Array(-39.155212865028, 12.075833333333)
    pts(116) = Array(-39.201597229121, 11.753333333333)
    pts(117) = Array(-39.246701170070, 11.430833333333)
    pts(118) = Array(-39.295317467388, 11.072500000000)
    pts(119) = Array(-39.337725074546, 10.750000000000)
    pts(120) = Array(-39.383350367071, 10.391666666667)
    pts(121) = Array(-39.423070354304, 10.069166666667)
    pts(122) = Array(-39.465713997200, 9.710833333333)
    pts(123) = Array(-39.530688032820, 9.137500000000)
    pts(124) = Array(-39.561674874733, 8.850833333333)
    pts(125) = Array(-39.595341467516, 8.528333333333)
    pts(126) = Array(-39.624207537447, 8.241666666667)
    pts(127) = Array(-39.655490944504, 7.919166666667)
    pts(128) = Array(-39.682240924205, 7.632500000000)
    pts(129) = Array(-39.711146212029, 7.310000000000)
    pts(130) = Array(-39.735784427752, 7.023333333333)
    pts(131) = Array(-39.762316264932, 6.700833333333)
    pts(132) = Array(-39.784846690966, 6.414166666667)
    pts(133) = Array(-39.809009351996, 6.091666666667)
    pts(134) = Array(-39.829435613907, 5.805000000000)
    pts(135) = Array(-39.851232982660, 5.482500000000)
    pts(136) = Array(-39.869558360233, 5.195833333333)
    pts(137) = Array(-39.888993933107, 4.873333333333)
    pts(138) = Array(-39.905221362994, 4.586666666667)
    pts(139) = Array(-39.922298251737, 4.264166666667)
    pts(140) = Array(-39.936430329825, 3.977500000000)
    pts(141) = Array(-39.951151264036, 3.655000000000)
    pts(142) = Array(-39.963190247533, 3.368333333333)
    pts(143) = Array(-39.975557576849, 3.045833333333)
    pts(144) = Array(-39.985505386092, 2.759166666667)
    pts(145) = Array(-39.995521082074, 2.436666666667)
    pts(146) = Array(-40.003379302066, 2.150000000000)
    pts(147) = Array(-40.011044959771, 1.827500000000)
    pts(148) = Array(-40.016814841443, 1.540833333333)
    pts(149) = Array(-40.022131680697, 1.218333333333)
    pts(150) = Array(-40.025814141908, 0.931666666667)
    pts(151) = Array(-40.028783008275, 0.609166666667)
    pts(152) = Array(-40.030378634540, 0.322500000000)
    pts(153) = Array(-40.031000000000, 0.000000000000)
    Profile_12 = pts
End Function

Private Function Profile_13() As Variant
    Dim pts(89) As Variant
    pts(0) = Array(-39.758000000000, 0.000000000000)
    pts(1) = Array(-39.756059863106, 0.330000000000)
    pts(2) = Array(-39.750238647499, 0.660000000000)
    pts(3) = Array(-39.740533936732, 0.990000000000)
    pts(4) = Array(-39.726941697817, 1.320000000000)
    pts(5) = Array(-39.709456272831, 1.650000000000)
    pts(6) = Array(-39.688070367107, 1.980000000000)
    pts(7) = Array(-39.662775033953, 2.310000000000)
    pts(8) = Array(-39.636377896099, 2.610000000000)
    pts(9) = Array(-39.603588202752, 2.940000000000)
    pts(10) = Array(-39.566853470467, 3.270000000000)
    pts(11) = Array(-39.526158079597, 3.600000000000)
    pts(12) = Array(-39.481484629982, 3.930000000000)
    pts(13) = Array(-39.432813902984, 4.260000000000)
    pts(14) = Array(-39.380124819135, 4.590000000000)
    pts(15) = Array(-39.323394391163, 4.920000000000)
    pts(16) = Array(-39.268293369487, 5.220000000000)
    pts(17) = Array(-39.203776678361, 5.550000000000)
    pts(18) = Array(-39.166851337177, 5.730000000000)
    pts(19) = Array(-39.135140417962, 5.880000000000)
    pts(20) = Array(-39.095955116005, 6.060000000000)
    pts(21) = Array(-39.062353596547, 6.210000000000)
    pts(22) = Array(-38.992554133570, 6.510000000000)
    pts(23) = Array(-38.949003797189, 6.690000000000)
    pts(24) = Array(-38.911749374305, 6.840000000000)
    pts(25) = Array(-38.834600391267, 7.140000000000)
    pts(26) = Array(-38.786611198539, 7.320000000000)
    pts(27) = Array(-38.745640346541, 7.470000000000)
    pts(28) = Array(-38.661009177202, 7.770000000000)
    pts(29) = Array(-38.608497840544, 7.950000000000)
    pts(30) = Array(-38.563738983549, 8.100000000000)
    pts(31) = Array(-38.471476395180, 8.400000000000)
    pts(32) = Array(-38.375522181301, 8.700000000000)
    pts(33) = Array(-38.275836367236, 9.000000000000)
    pts(34) = Array(-38.214216272401, 9.180000000000)
    pts(35) = Array(-38.161821608768, 9.330000000000)
    pts(36) = Array(-38.054159547351, 9.630000000000)
    pts(37) = Array(-37.942627725828, 9.930000000000)
    pts(38) = Array(-37.827176601661, 10.230000000000)
    pts(39) = Array(-37.731958832533, 10.470000000000)
    pts(40) = Array(-37.646536327971, 10.680000000000)
    pts(41) = Array(-37.559126937485, 10.890000000000)
    pts(42) = Array(-37.469710449146, 11.100000000000)
    pts(43) = Array(-37.365035582104, 11.340000000000)
    pts(44) = Array(-37.271246587810, 11.550000000000)
    pts(45) = Array(-37.175381957660, 11.760000000000)
    pts(46) = Array(-37.077418176583, 11.970000000000)
    pts(47) = Array(-36.962857840603, 12.210000000000)
    pts(48) = Array(-36.860312537792, 12.420000000000)
    pts(49) = Array(-36.755588250687, 12.630000000000)
    pts(50) = Array(-36.648657548638, 12.840000000000)
    pts(51) = Array(-36.539491945868, 13.050000000000)
    pts(52) = Array(-36.428061853092, 13.260000000000)
    pts(53) = Array(-36.314336526008, 13.470000000000)
    pts(54) = Array(-36.198284010432, 13.680000000000)
    pts(55) = Array(-36.079871083799, 13.890000000000)
    pts(56) = Array(-35.959063192731, 14.100000000000)
    pts(57) = Array(-35.835824386372, 14.310000000000)
    pts(58) = Array(-35.728227990265, 14.490000000000)
    pts(59) = Array(-35.600374181445, 14.700000000000)
    pts(60) = Array(-35.469978383369, 14.910000000000)
    pts(61) = Array(-35.336998540917, 15.120000000000)
    pts(62) = Array(-35.220925942337, 15.300000000000)
    pts(63) = Array(-35.083029391627, 15.510000000000)
    pts(64) = Array(-34.962673733263, 15.690000000000)
    pts(65) = Array(-34.819696729246, 15.900000000000)
    pts(66) = Array(-34.694911902626, 16.080000000000)
    pts(67) = Array(-34.546676895550, 16.290000000000)
    pts(68) = Array(-34.417304310393, 16.470000000000)
    pts(69) = Array(-34.263618373523, 16.680000000000)
    pts(70) = Array(-34.126479267413, 16.864000000000)
    pts(71) = Array(-34.126479267413, 18.000000000000)
    pts(72) = Array(-32.710467384946, 18.000000000000)
    pts(73) = Array(-32.710467384946, 16.864000000000)
    pts(74) = Array(-32.761789398983, 15.750000000000)
    pts(75) = Array(-32.809433309812, 14.640000000000)
    pts(76) = Array(-32.853591091018, 13.530000000000)
    pts(77) = Array(-32.895314962302, 12.390000000000)
    pts(78) = Array(-32.932410156840, 11.280000000000)
    pts(79) = Array(-32.966882980116, 10.140000000000)
    pts(80) = Array(-32.996920085418, 9.030000000000)
    pts(81) = Array(-33.024146038491, 7.890000000000)
    pts(82) = Array(-33.044226251731, 6.930000000000)
    pts(83) = Array(-33.062208305634, 5.940000000000)
    pts(84) = Array(-33.077423176867, 4.950000000000)
    pts(85) = Array(-33.089871221912, 3.960000000000)
    pts(86) = Array(-33.099552732378, 2.970000000000)
    pts(87) = Array(-33.106467935040, 1.980000000000)
    pts(88) = Array(-33.110616991857, 0.990000000000)
    pts(89) = Array(-33.112000000000, 0.000000000000)
    Profile_13 = pts
End Function

Private Function Profile_14() As Variant
    Dim pts(83) As Variant
    pts(0) = Array(-33.112000000000, 0.000000000000)
    pts(1) = Array(-33.110616991857, 0.990000000000)
    pts(2) = Array(-33.106467935040, 1.980000000000)
    pts(3) = Array(-33.099552732378, 2.970000000000)
    pts(4) = Array(-33.089871221912, 3.960000000000)
    pts(5) = Array(-33.077423176867, 4.950000000000)
    pts(6) = Array(-33.062208305634, 5.940000000000)
    pts(7) = Array(-33.044226251731, 6.930000000000)
    pts(8) = Array(-33.024146038491, 7.890000000000)
    pts(9) = Array(-32.996920085418, 9.030000000000)
    pts(10) = Array(-32.966882980116, 10.140000000000)
    pts(11) = Array(-32.932410156840, 11.280000000000)
    pts(12) = Array(-32.895314962302, 12.390000000000)
    pts(13) = Array(-32.853591091018, 13.530000000000)
    pts(14) = Array(-32.809433309812, 14.640000000000)
    pts(15) = Array(-32.761789398983, 15.750000000000)
    pts(16) = Array(-32.710467384946, 16.864000000000)
    pts(17) = Array(-32.710467384946, 18.000000000000)
    pts(18) = Array(-26.588086491670, 18.000000000000)
    pts(19) = Array(-26.588086491670, 11.300000000000)
    pts(20) = Array(-26.720656753404, 11.160000000000)
    pts(21) = Array(-26.859277392368, 11.010000000000)
    pts(22) = Array(-26.994479046255, 10.860000000000)
    pts(23) = Array(-27.126374173222, 10.710000000000)
    pts(24) = Array(-27.255067776702, 10.560000000000)
    pts(25) = Array(-27.380658067817, 10.410000000000)
    pts(26) = Array(-27.503237053187, 10.260000000000)
    pts(27) = Array(-27.622891058159, 10.110000000000)
    pts(28) = Array(-27.739701193925, 9.960000000000)
    pts(29) = Array(-27.853743775724, 9.810000000000)
    pts(30) = Array(-27.965090698245, 9.660000000000)
    pts(31) = Array(-28.095243945131, 9.480000000000)
    pts(32) = Array(-28.200893815781, 9.330000000000)
    pts(33) = Array(-28.304052113305, 9.180000000000)
    pts(34) = Array(-28.404775494107, 9.030000000000)
    pts(35) = Array(-28.522504835184, 8.850000000000)
    pts(36) = Array(-28.618056427578, 8.700000000000)
    pts(37) = Array(-28.729722422783, 8.520000000000)
    pts(38) = Array(-28.820333828469, 8.370000000000)
    pts(39) = Array(-28.926199135247, 8.190000000000)
    pts(40) = Array(-29.012077371950, 8.040000000000)
    pts(41) = Array(-29.112377169021, 7.860000000000)
    pts(42) = Array(-29.209731665904, 7.680000000000)
    pts(43) = Array(-29.304200150250, 7.500000000000)
    pts(44) = Array(-29.380759801944, 7.350000000000)
    pts(45) = Array(-29.470080544924, 7.170000000000)
    pts(46) = Array(-29.556666199993, 6.990000000000)
    pts(47) = Array(-29.640564623820, 6.810000000000)
    pts(48) = Array(-29.721821128425, 6.630000000000)
    pts(49) = Array(-29.800478632819, 6.450000000000)
    pts(50) = Array(-29.876577802074, 6.270000000000)
    pts(51) = Array(-29.950157175035, 6.090000000000)
    pts(52) = Array(-30.021253281750, 5.910000000000)
    pts(53) = Array(-30.089900751558, 5.730000000000)
    pts(54) = Array(-30.156132412684, 5.550000000000)
    pts(55) = Array(-30.219979384103, 5.370000000000)
    pts(56) = Array(-30.281471160339, 5.190000000000)
    pts(57) = Array(-30.340635689801, 5.010000000000)
    pts(58) = Array(-30.397499447204, 4.830000000000)
    pts(59) = Array(-30.460965945322, 4.620000000000)
    pts(60) = Array(-30.512928903008, 4.440000000000)
    pts(61) = Array(-30.562665933803, 4.260000000000)
    pts(62) = Array(-30.610198038596, 4.080000000000)
    pts(63) = Array(-30.662891762928, 3.870000000000)
    pts(64) = Array(-30.705713026134, 3.690000000000)
    pts(65) = Array(-30.752960051335, 3.480000000000)
    pts(66) = Array(-30.791152318605, 3.300000000000)
    pts(67) = Array(-30.833041931608, 3.090000000000)
    pts(68) = Array(-30.866676608245, 2.910000000000)
    pts(69) = Array(-30.898229148217, 2.730000000000)
    pts(70) = Array(-30.927711768330, 2.550000000000)
    pts(71) = Array(-30.959507075689, 2.340000000000)
    pts(72) = Array(-30.984542581556, 2.160000000000)
    pts(73) = Array(-31.011176768893, 1.950000000000)
    pts(74) = Array(-31.031809692794, 1.770000000000)
    pts(75) = Array(-31.053329834273, 1.560000000000)
    pts(76) = Array(-31.069596431248, 1.380000000000)
    pts(77) = Array(-31.086040144816, 1.170000000000)
    pts(78) = Array(-31.097968717384, 0.990000000000)
    pts(79) = Array(-31.109364498809, 0.780000000000)
    pts(80) = Array(-31.116975646349, 0.600000000000)
    pts(81) = Array(-31.123343118639, 0.390000000000)
    pts(82) = Array(-31.126649916650, 0.210000000000)
    pts(83) = Array(-31.128000000000, 0.000000000000)
    Profile_14 = pts
End Function

Private Function Profile_15() As Variant
    Dim pts(99) As Variant
    pts(0) = Array(-11.610000000000, 0.000000000000)
    pts(1) = Array(-11.611195794770, 0.192500000000)
    pts(2) = Array(-11.614124661595, 0.357500000000)
    pts(3) = Array(-11.619764289815, 0.550000000000)
    pts(4) = Array(-11.626505241142, 0.715000000000)
    pts(5) = Array(-11.636597731181, 0.907500000000)
    pts(6) = Array(-11.647161575757, 1.072500000000)
    pts(7) = Array(-11.661723145653, 1.265000000000)
    pts(8) = Array(-11.676126904404, 1.430000000000)
    pts(9) = Array(-11.695181092888, 1.622500000000)
    pts(10) = Array(-11.713448137046, 1.787500000000)
    pts(11) = Array(-11.737025995597, 1.980000000000)
    pts(12) = Array(-11.759186238586, 2.145000000000)
    pts(13) = Array(-11.787326588218, 2.337500000000)
    pts(14) = Array(-11.813416735014, 2.502500000000)
    pts(15) = Array(-11.841335015797, 2.667500000000)
    pts(16) = Array(-11.871091481501, 2.832500000000)
    pts(17) = Array(-11.908145062218, 3.025000000000)
    pts(18) = Array(-11.941922311087, 3.190000000000)
    pts(19) = Array(-11.983699832686, 3.382500000000)
    pts(20) = Array(-12.021556469095, 3.547500000000)
    pts(21) = Array(-12.068131143752, 3.740000000000)
    pts(22) = Array(-12.110134212004, 3.905000000000)
    pts(23) = Array(-12.154076115504, 4.070000000000)
    pts(24) = Array(-12.199974002035, 4.235000000000)
    pts(25) = Array(-12.256017922093, 4.427500000000)
    pts(26) = Array(-12.306217322023, 4.592500000000)
    pts(27) = Array(-12.358433662374, 4.757500000000)
    pts(28) = Array(-12.412688447695, 4.922500000000)
    pts(29) = Array(-12.469004347158, 5.087500000000)
    pts(30) = Array(-12.527405248121, 5.252500000000)
    pts(31) = Array(-12.587916313874, 5.417500000000)
    pts(32) = Array(-12.650564045941, 5.582500000000)
    pts(33) = Array(-12.715376351343, 5.747500000000)
    pts(34) = Array(-12.782382615248, 5.912500000000)
    pts(35) = Array(-12.851613779526, 6.077500000000)
    pts(36) = Array(-12.923102427743, 6.242500000000)
    pts(37) = Array(-12.996882877223, 6.407500000000)
    pts(38) = Array(-13.072991278865, 6.572500000000)
    pts(39) = Array(-13.151465725487, 6.737500000000)
    pts(40) = Array(-13.232346369566, 6.902500000000)
    pts(41) = Array(-13.315675551355, 7.067500000000)
    pts(42) = Array(-13.401497938483, 7.232500000000)
    pts(43) = Array(-13.489860678274, 7.397500000000)
    pts(44) = Array(-13.565472681811, 7.535000000000)
    pts(45) = Array(-13.658624094203, 7.700000000000)
    pts(46) = Array(-13.738301875557, 7.837500000000)
    pts(47) = Array(-13.836427403447, 8.002500000000)
    pts(48) = Array(-13.920333716775, 8.140000000000)
    pts(49) = Array(-14.023639214516, 8.305000000000)
    pts(50) = Array(-14.111954985455, 8.442500000000)
    pts(51) = Array(-14.220669642935, 8.607500000000)
    pts(52) = Array(-14.313596636803, 8.745000000000)
    pts(53) = Array(-14.427976492883, 8.910000000000)
    pts(54) = Array(-14.525740532527, 9.047500000000)
    pts(55) = Array(-14.625785281117, 9.185000000000)
    pts(56) = Array(-14.728166044986, 9.322500000000)
    pts(57) = Array(-14.832941289254, 9.460000000000)
    pts(58) = Array(-14.940172884500, 9.597500000000)
    pts(59) = Array(-15.049926379321, 9.735000000000)
    pts(60) = Array(-15.162271302169, 9.872500000000)
    pts(61) = Array(-15.277281496417, 10.010000000000)
    pts(62) = Array(-15.395035493222, 10.147500000000)
    pts(63) = Array(-15.515616927550, 10.285000000000)
    pts(64) = Array(-15.633209860706, 10.416000000000)
    pts(65) = Array(-15.633209860706, 16.500000000000)
    pts(66) = Array(-8.708198123012, 16.500000000000)
    pts(67) = Array(-8.708198123012, 15.750000000000)
    pts(68) = Array(-8.765803118206, 15.235000000000)
    pts(69) = Array(-8.822266837958, 14.712500000000)
    pts(70) = Array(-8.873921181537, 14.217500000000)
    pts(71) = Array(-8.926507767916, 13.695000000000)
    pts(72) = Array(-8.974493180761, 13.200000000000)
    pts(73) = Array(-9.023211214283, 12.677500000000)
    pts(74) = Array(-9.067535514792, 12.182500000000)
    pts(75) = Array(-9.112392922162, 11.660000000000)
    pts(76) = Array(-9.153063312492, 11.165000000000)
    pts(77) = Array(-9.194067373133, 10.642500000000)
    pts(78) = Array(-9.233095307949, 10.120000000000)
    pts(79) = Array(-9.270148781018, 9.597500000000)
    pts(80) = Array(-9.303432184635, 9.102500000000)
    pts(81) = Array(-9.336645104719, 8.580000000000)
    pts(82) = Array(-9.367887969187, 8.057500000000)
    pts(83) = Array(-9.397162104874, 7.535000000000)
    pts(84) = Array(-9.421686986414, 7.067500000000)
    pts(85) = Array(-9.444637622181, 6.600000000000)
    pts(86) = Array(-9.466014790641, 6.132500000000)
    pts(87) = Array(-9.485819216428, 5.665000000000)
    pts(88) = Array(-9.504051570467, 5.197500000000)
    pts(89) = Array(-9.520712470089, 4.730000000000)
    pts(90) = Array(-9.535802479132, 4.262500000000)
    pts(91) = Array(-9.550068485796, 3.767500000000)
    pts(92) = Array(-9.561925856394, 3.300000000000)
    pts(93) = Array(-9.572213730002, 2.832500000000)
    pts(94) = Array(-9.580932454236, 2.365000000000)
    pts(95) = Array(-9.588082323612, 1.897500000000)
    pts(96) = Array(-9.593663579597, 1.430000000000)
    pts(97) = Array(-9.597676410647, 0.962500000000)
    pts(98) = Array(-9.600120952240, 0.495000000000)
    pts(99) = Array(-9.601000000000, 0.000000000000)
    Profile_15 = pts
End Function

Private Function Profile_16() As Variant
    Dim pts(100) As Variant
    pts(0) = Array(-9.601000000000, 0.000000000000)
    pts(1) = Array(-9.600120952240, 0.495000000000)
    pts(2) = Array(-9.597676410647, 0.962500000000)
    pts(3) = Array(-9.593663579597, 1.430000000000)
    pts(4) = Array(-9.588082323612, 1.897500000000)
    pts(5) = Array(-9.580932454236, 2.365000000000)
    pts(6) = Array(-9.572213730002, 2.832500000000)
    pts(7) = Array(-9.561925856394, 3.300000000000)
    pts(8) = Array(-9.550068485796, 3.767500000000)
    pts(9) = Array(-9.535802479132, 4.262500000000)
    pts(10) = Array(-9.520712470089, 4.730000000000)
    pts(11) = Array(-9.504051570467, 5.197500000000)
    pts(12) = Array(-9.485819216428, 5.665000000000)
    pts(13) = Array(-9.466014790641, 6.132500000000)
    pts(14) = Array(-9.444637622181, 6.600000000000)
    pts(15) = Array(-9.421686986414, 7.067500000000)
    pts(16) = Array(-9.397162104874, 7.535000000000)
    pts(17) = Array(-9.367887969187, 8.057500000000)
    pts(18) = Array(-9.336645104719, 8.580000000000)
    pts(19) = Array(-9.303432184635, 9.102500000000)
    pts(20) = Array(-9.270148781018, 9.597500000000)
    pts(21) = Array(-9.233095307949, 10.120000000000)
    pts(22) = Array(-9.194067373133, 10.642500000000)
    pts(23) = Array(-9.153063312492, 11.165000000000)
    pts(24) = Array(-9.112392922162, 11.660000000000)
    pts(25) = Array(-9.067535514792, 12.182500000000)
    pts(26) = Array(-9.023211214283, 12.677500000000)
    pts(27) = Array(-8.974493180761, 13.200000000000)
    pts(28) = Array(-8.926507767916, 13.695000000000)
    pts(29) = Array(-8.873921181537, 14.217500000000)
    pts(30) = Array(-8.822266837958, 14.712500000000)
    pts(31) = Array(-8.765803118206, 15.235000000000)
    pts(32) = Array(-8.708198123012, 15.750000000000)
    pts(33) = Array(-8.708198123012, 16.500000000000)
    pts(34) = Array(-7.877115975542, 16.500000000000)
    pts(35) = Array(-7.877115975542, 15.750000000000)
    pts(36) = Array(-7.714349399320, 15.565000000000)
    pts(37) = Array(-7.548556322309, 15.372500000000)
    pts(38) = Array(-7.386308142781, 15.180000000000)
    pts(39) = Array(-7.227510904193, 14.987500000000)
    pts(40) = Array(-7.072075837088, 14.795000000000)
    pts(41) = Array(-6.919918976450, 14.602500000000)
    pts(42) = Array(-6.770960815088, 14.410000000000)
    pts(43) = Array(-6.625125988977, 14.217500000000)
    pts(44) = Array(-6.482342991042, 14.025000000000)
    pts(45) = Array(-6.322812393876, 13.805000000000)
    pts(46) = Array(-6.186344645181, 13.612500000000)
    pts(47) = Array(-6.033867171174, 13.392500000000)
    pts(48) = Array(-5.903434384394, 13.200000000000)
    pts(49) = Array(-5.757706062478, 12.980000000000)
    pts(50) = Array(-5.633055645400, 12.787500000000)
    pts(51) = Array(-5.493802077052, 12.567500000000)
    pts(52) = Array(-5.374705645908, 12.375000000000)
    pts(53) = Array(-5.241678356350, 12.155000000000)
    pts(54) = Array(-5.111878149776, 11.935000000000)
    pts(55) = Array(-4.985244044731, 11.715000000000)
    pts(56) = Array(-4.861718019284, 11.495000000000)
    pts(57) = Array(-4.741244833819, 11.275000000000)
    pts(58) = Array(-4.623771867895, 11.055000000000)
    pts(59) = Array(-4.509248969833, 10.835000000000)
    pts(60) = Array(-4.397628317876, 10.615000000000)
    pts(61) = Array(-4.288864291847, 10.395000000000)
    pts(62) = Array(-4.182913354372, 10.175000000000)
    pts(63) = Array(-4.079733940842, 9.955000000000)
    pts(64) = Array(-3.979286357331, 9.735000000000)
    pts(65) = Array(-3.881532685822, 9.515000000000)
    pts(66) = Array(-3.786436696109, 9.295000000000)
    pts(67) = Array(-3.682587416592, 9.047500000000)
    pts(68) = Array(-3.582010995005, 8.800000000000)
    pts(69) = Array(-3.474043697484, 8.525000000000)
    pts(70) = Array(-3.370003780038, 8.250000000000)
    pts(71) = Array(-3.269836327990, 7.975000000000)
    pts(72) = Array(-3.173489433952, 7.700000000000)
    pts(73) = Array(-3.080914042188, 7.425000000000)
    pts(74) = Array(-2.992063805400, 7.150000000000)
    pts(75) = Array(-2.906894952855, 6.875000000000)
    pts(76) = Array(-2.825366168887, 6.600000000000)
    pts(77) = Array(-2.747438480899, 6.325000000000)
    pts(78) = Array(-2.673075156107, 6.050000000000)
    pts(79) = Array(-2.637219239775, 5.912500000000)
    pts(80) = Array(-2.595351119130, 5.747500000000)
    pts(81) = Array(-2.528362824784, 5.472500000000)
    pts(82) = Array(-2.464838268607, 5.197500000000)
    pts(83) = Array(-2.404748982469, 4.922500000000)
    pts(84) = Array(-2.342586760092, 4.620000000000)
    pts(85) = Array(-2.289626923423, 4.345000000000)
    pts(86) = Array(-2.235249161110, 4.042500000000)
    pts(87) = Array(-2.189317253144, 3.767500000000)
    pts(88) = Array(-2.142620669195, 3.465000000000)
    pts(89) = Array(-2.103630197053, 3.190000000000)
    pts(90) = Array(-2.064527552880, 2.887500000000)
    pts(91) = Array(-2.032406205751, 2.612500000000)
    pts(92) = Array(-2.000825463746, 2.310000000000)
    pts(93) = Array(-1.975514423183, 2.035000000000)
    pts(94) = Array(-1.951398067097, 1.732500000000)
    pts(95) = Array(-1.932851457628, 1.457500000000)
    pts(96) = Array(-1.916155960135, 1.155000000000)
    pts(97) = Array(-1.904340427978, 0.880000000000)
    pts(98) = Array(-1.895035853722, 0.577500000000)
    pts(99) = Array(-1.889930264426, 0.302500000000)
    pts(100) = Array(-1.888000000000, 0.000000000000)
    Profile_16 = pts
End Function

Private Function Profile_17() As Variant
    Dim pts(101) As Variant
    pts(0) = Array(-1.615000000000, 0.000000000000)
    pts(1) = Array(-1.615833480150, 0.595000000000)
    pts(2) = Array(-1.618333940227, 1.190000000000)
    pts(3) = Array(-1.622501439112, 1.785000000000)
    pts(4) = Array(-1.628336074949, 2.380000000000)
    pts(5) = Array(-1.635837985153, 2.975000000000)
    pts(6) = Array(-1.645007346431, 3.570000000000)
    pts(7) = Array(-1.655844374801, 4.165000000000)
    pts(8) = Array(-1.668349325614, 4.760000000000)
    pts(9) = Array(-1.682522493590, 5.355000000000)
    pts(10) = Array(-1.698364212847, 5.950000000000)
    pts(11) = Array(-1.715874856946, 6.545000000000)
    pts(12) = Array(-1.735054838930, 7.140000000000)
    pts(13) = Array(-1.755904611376, 7.735000000000)
    pts(14) = Array(-1.778424666448, 8.330000000000)
    pts(15) = Array(-1.802615535953, 8.925000000000)
    pts(16) = Array(-1.828477791408, 9.520000000000)
    pts(17) = Array(-1.859361321454, 10.185000000000)
    pts(18) = Array(-1.892334315047, 10.850000000000)
    pts(19) = Array(-1.927397746829, 11.515000000000)
    pts(20) = Array(-1.962544972898, 12.145000000000)
    pts(21) = Array(-2.001682290996, 12.810000000000)
    pts(22) = Array(-2.042913283811, 13.475000000000)
    pts(23) = Array(-2.086239173452, 14.140000000000)
    pts(24) = Array(-2.129218328852, 14.770000000000)
    pts(25) = Array(-2.176627500139, 15.435000000000)
    pts(26) = Array(-2.226135538158, 16.100000000000)
    pts(27) = Array(-2.277743915324, 16.765000000000)
    pts(28) = Array(-2.328574885622, 17.395000000000)
    pts(29) = Array(-2.384277864073, 18.060000000000)
    pts(30) = Array(-2.438990858126, 18.690000000000)
    pts(31) = Array(-2.498794738957, 19.355000000000)
    pts(32) = Array(-2.558814296742, 20.000000000000)
    pts(33) = Array(-2.558814296742, 21.000000000000)
    pts(34) = Array(-1.732163854282, 21.000000000000)
    pts(35) = Array(-1.732163854282, 20.000000000000)
    pts(36) = Array(-1.563329668698, 19.740000000000)
    pts(37) = Array(-1.407087548762, 19.495000000000)
    pts(38) = Array(-1.253569647746, 19.250000000000)
    pts(39) = Array(-1.081404786688, 18.970000000000)
    pts(40) = Array(-0.933590731765, 18.725000000000)
    pts(41) = Array(-0.767844992281, 18.445000000000)
    pts(42) = Array(-0.625563837169, 18.200000000000)
    pts(43) = Array(-0.466049337693, 17.920000000000)
    pts(44) = Array(-0.329142929075, 17.675000000000)
    pts(45) = Array(-0.175685714489, 17.395000000000)
    pts(46) = Array(-0.025392343025, 17.115000000000)
    pts(47) = Array(0.121781256019, 16.835000000000)
    pts(48) = Array(0.248032312647, 16.590000000000)
    pts(49) = Array(0.389469248373, 16.310000000000)
    pts(50) = Array(0.527903890019, 16.030000000000)
    pts(51) = Array(0.663374124279, 15.750000000000)
    pts(52) = Array(0.795916411842, 15.470000000000)
    pts(53) = Array(0.925565848906, 15.190000000000)
    pts(54) = Array(1.052356225008, 14.910000000000)
    pts(55) = Array(1.176320077455, 14.630000000000)
    pts(56) = Array(1.297488742564, 14.350000000000)
    pts(57) = Array(1.415892403951, 14.070000000000)
    pts(58) = Array(1.531560138052, 13.790000000000)
    pts(59) = Array(1.644519957063, 13.510000000000)
    pts(60) = Array(1.754798849473, 13.230000000000)
    pts(61) = Array(1.862422818336, 12.950000000000)
    pts(62) = Array(1.967416917434, 12.670000000000)
    pts(63) = Array(2.069805285457, 12.390000000000)
    pts(64) = Array(2.169611178325, 12.110000000000)
    pts(65) = Array(2.266856999762, 11.830000000000)
    pts(66) = Array(2.361564330235, 11.550000000000)
    pts(67) = Array(2.465101636173, 11.235000000000)
    pts(68) = Array(2.565480879707, 10.920000000000)
    pts(69) = Array(2.673342567048, 10.570000000000)
    pts(70) = Array(2.777374726753, 10.220000000000)
    pts(71) = Array(2.877611385751, 9.870000000000)
    pts(72) = Array(2.974084949218, 9.520000000000)
    pts(73) = Array(3.066826258280, 9.170000000000)
    pts(74) = Array(3.155864643980, 8.820000000000)
    pts(75) = Array(3.241227977739, 8.470000000000)
    pts(76) = Array(3.322942718537, 8.120000000000)
    pts(77) = Array(3.401033957008, 7.770000000000)
    pts(78) = Array(3.475525456636, 7.420000000000)
    pts(79) = Array(3.546439692212, 7.070000000000)
    pts(80) = Array(3.613797885714, 6.720000000000)
    pts(81) = Array(3.677620039744, 6.370000000000)
    pts(82) = Array(3.737924968659, 6.020000000000)
    pts(83) = Array(3.794730327497, 5.670000000000)
    pts(84) = Array(3.848052638831, 5.320000000000)
    pts(85) = Array(3.897907317629, 4.970000000000)
    pts(86) = Array(3.944308694216, 4.620000000000)
    pts(87) = Array(3.987270035424, 4.270000000000)
    pts(88) = Array(4.026803564005, 3.920000000000)
    pts(89) = Array(4.062920476366, 3.570000000000)
    pts(90) = Array(4.095630958696, 3.220000000000)
    pts(91) = Array(4.110711699171, 3.045000000000)
    pts(92) = Array(4.127689002107, 2.835000000000)
    pts(93) = Array(4.153274724238, 2.485000000000)
    pts(94) = Array(4.175479322528, 2.135000000000)
    pts(95) = Array(4.194308977588, 1.785000000000)
    pts(96) = Array(4.209768918817, 1.435000000000)
    pts(97) = Array(4.221863431700, 1.085000000000)
    pts(98) = Array(4.230595863763, 0.735000000000)
    pts(99) = Array(4.235968629211, 0.385000000000)
    pts(100) = Array(4.237395637039, 0.210000000000)
    pts(101) = Array(4.238000000000, 0.000000000000)
    Profile_17 = pts
End Function

Private Function Profile_18() As Variant
    Dim pts(83) As Variant
    pts(0) = Array(4.511000000000, 0.000000000000)
    pts(1) = Array(4.511782538115, 0.345000000000)
    pts(2) = Array(4.514130200774, 0.690000000000)
    pts(3) = Array(4.518043132935, 1.035000000000)
    pts(4) = Array(4.522835541442, 1.341666666667)
    pts(5) = Array(4.529705831301, 1.686666666667)
    pts(6) = Array(4.538142352841, 2.031666666667)
    pts(7) = Array(4.548145627553, 2.376666666667)
    pts(8) = Array(4.558353224189, 2.683333333333)
    pts(9) = Array(4.571317688139, 3.028333333333)
    pts(10) = Array(4.585850958197, 3.373333333333)
    pts(11) = Array(4.601953934833, 3.718333333333)
    pts(12) = Array(4.617586275224, 4.025000000000)
    pts(13) = Array(4.636657057080, 4.370000000000)
    pts(14) = Array(4.654929249714, 4.676666666667)
    pts(15) = Array(4.676972063727, 5.021666666667)
    pts(16) = Array(4.697888162525, 5.328333333333)
    pts(17) = Array(4.722907904374, 5.673333333333)
    pts(18) = Array(4.749505878721, 6.018333333333)
    pts(19) = Array(4.777683747302, 6.363333333333)
    pts(20) = Array(4.804058495495, 6.670000000000)
    pts(21) = Array(4.835225501148, 7.015000000000)
    pts(22) = Array(4.867977772350, 7.360000000000)
    pts(23) = Array(4.902317367834, 7.705000000000)
    pts(24) = Array(4.934175766969, 8.011666666667)
    pts(25) = Array(4.971519625668, 8.356666666667)
    pts(26) = Array(5.010457340553, 8.701666666667)
    pts(27) = Array(5.050991377268, 9.046666666667)
    pts(28) = Array(5.088363831099, 9.353333333333)
    pts(29) = Array(5.131920246923, 9.698333333333)
    pts(30) = Array(5.171983578726, 10.005000000000)
    pts(31) = Array(5.218572277338, 10.350000000000)
    pts(32) = Array(5.261335616772, 10.656666666667)
    pts(33) = Array(5.316581528175, 11.040000000000)
    pts(34) = Array(5.368008378328, 11.385000000000)
    pts(35) = Array(5.427048535614, 11.768333333333)
    pts(36) = Array(5.481897732118, 12.113333333333)
    pts(37) = Array(5.544749162461, 12.496666666667)
    pts(38) = Array(5.603036571853, 12.841666666667)
    pts(39) = Array(5.662958462593, 13.186666666667)
    pts(40) = Array(5.724518760836, 13.531666666667)
    pts(41) = Array(5.794845607178, 13.915000000000)
    pts(42) = Array(5.859878202825, 14.260000000000)
    pts(43) = Array(5.926562197304, 14.605000000000)
    pts(44) = Array(5.994902021065, 14.950000000000)
    pts(45) = Array(6.072782727371, 15.333333333333)
    pts(46) = Array(6.144633311243, 15.678333333333)
    pts(47) = Array(6.218154324054, 16.023333333333)
    pts(48) = Array(6.293350726643, 16.368333333333)
    pts(49) = Array(6.378873448951, 16.751666666667)
    pts(50) = Array(6.457623679239, 17.096666666667)
    pts(51) = Array(6.538065594370, 17.441666666667)
    pts(52) = Array(6.620204715540, 17.786666666667)
    pts(53) = Array(6.713467848033, 18.170000000000)
    pts(54) = Array(6.799208729604, 18.515000000000)
    pts(55) = Array(6.886664907446, 18.860000000000)
    pts(56) = Array(6.975842497382, 19.205000000000)
    pts(57) = Array(7.076955275088, 19.588333333333)
    pts(58) = Array(7.169787720888, 19.933333333333)
    pts(59) = Array(7.264361578136, 20.278333333333)
    pts(60) = Array(7.360683596384, 20.623333333333)
    pts(61) = Array(7.458760685847, 20.968333333333)
    pts(62) = Array(7.558599920125, 21.313333333333)
    pts(63) = Array(7.660208539006, 21.658333333333)
    pts(64) = Array(7.762586533956, 22.000000000000)
    pts(65) = Array(7.762586533956, 23.000000000000)
    pts(66) = Array(8.789581280691, 23.000000000000)
    pts(67) = Array(8.789581280691, 22.000000000000)
    pts(68) = Array(8.816465934037, 20.546666666667)
    pts(69) = Array(8.840931722403, 19.128333333333)
    pts(70) = Array(8.863647800012, 17.710000000000)
    pts(71) = Array(8.885156650475, 16.253333333333)
    pts(72) = Array(8.904326326777, 14.835000000000)
    pts(73) = Array(8.922193120334, 13.378333333333)
    pts(74) = Array(8.937816754772, 11.960000000000)
    pts(75) = Array(8.952041826416, 10.503333333333)
    pts(76) = Array(8.963205586441, 9.200000000000)
    pts(77) = Array(8.972892340994, 7.896666666667)
    pts(78) = Array(8.981102127393, 6.593333333333)
    pts(79) = Array(8.988010644073, 5.251666666667)
    pts(80) = Array(8.993223145143, 3.948333333333)
    pts(81) = Array(8.996958756379, 2.645000000000)
    pts(82) = Array(8.999217492170, 1.341666666667)
    pts(83) = Array(9.000000000000, 0.000000000000)
    Profile_18 = pts
End Function

Private Function Profile_19() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(25.900000000000, 0.000000000000)
    pts(1) = Array(25.900000000000, 1.400000000000)
    pts(2) = Array(27.060000000000, 1.400000000000)
    pts(3) = Array(27.060000000000, 1.500000000000)
    pts(4) = Array(30.300000000000, 1.500000000000)
    pts(5) = Array(30.300000000000, 0.000000000000)
    Profile_19 = pts
End Function

Private Function Profile_20() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(25.900000000000, 0.000000000000)
    pts(1) = Array(25.900000000000, 1.400000000000)
    pts(2) = Array(27.060000000000, 1.400000000000)
    pts(3) = Array(27.060000000000, 1.500000000000)
    pts(4) = Array(30.300000000000, 1.500000000000)
    pts(5) = Array(30.300000000000, 0.000000000000)
    Profile_20 = pts
End Function

Private Function Profile_21() As Variant
    Dim pts(4) As Variant
    pts(0) = Array(-7.000000000000, 0.000000000000)
    pts(1) = Array(-7.000000000000, 1.500000000000)
    pts(2) = Array(-1.200000000000, 1.500000000000)
    pts(3) = Array(0.000000000000, 2.750000000000)
    pts(4) = Array(0.000000000000, 0.000000000000)
    Profile_21 = pts
End Function

Private Function Profile_22() As Variant
    Dim pts(4) As Variant
    pts(0) = Array(-7.000000000000, 0.000000000000)
    pts(1) = Array(-7.000000000000, 1.500000000000)
    pts(2) = Array(-1.200000000000, 1.500000000000)
    pts(3) = Array(0.000000000000, 2.750000000000)
    pts(4) = Array(0.000000000000, 0.000000000000)
    Profile_22 = pts
End Function

Private Function Profile_23() As Variant
    Dim pts(4) As Variant
    pts(0) = Array(-7.000000000000, 0.000000000000)
    pts(1) = Array(-7.000000000000, 1.500000000000)
    pts(2) = Array(-1.200000000000, 1.500000000000)
    pts(3) = Array(0.000000000000, 2.750000000000)
    pts(4) = Array(0.000000000000, 0.000000000000)
    Profile_23 = pts
End Function

Private Function Profile_24() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.800000000000, 0.000000000000)
    pts(1) = Array(24.800000000000, 1.250000000000)
    pts(2) = Array(28.900000000000, 1.250000000000)
    pts(3) = Array(28.900000000000, 0.000000000000)
    Profile_24 = pts
End Function

Private Sub AddFocusDetails(ByVal model As SldWorks.ModelDoc2, ByVal label As String)
    Select Case label
        Case "01_OpticalCartridge"
            FocusLoftCut model, "ClosedGuideSlot_1", Array(0, 1)
            FocusLoftCut model, "ClosedGuideSlot_2", Array(2, 3)
            FocusLoftCut model, "M2p5_KeeperLockSeat", Array(4, 5)
        Case "07_IntegratedHousing"
            FocusLoftCut model, "M3_GuideThreadEnvelope_1", Array(6, 7)
            FocusLoftCut model, "M3_GuideThreadEnvelope_2", Array(8, 9)
            FocusLoftCut model, "M3_CustomAttachmentSeat_1", Array(10, 11)
            FocusLoftCut model, "M3_CustomAttachmentSeat_2", Array(12, 13)
            FocusLoftCut model, "M3_CustomAttachmentSeat_3", Array(14, 15)
            FocusLoftCut model, "Male_HelicoidGroove_1", Array(16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, _
                36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, _
                56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, _
                76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, _
                96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, _
                116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, _
                136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, _
                156, 157, 158, 159, 160, 161, 162, 163)
        Case "12_EFMountInterfaceEnvelope"
            FocusLoftCut model, "M3_CustomAttachment_1", Array(164, 165)
            FocusLoftCut model, "M3_HeadRelief_1", Array(166, 167)
            FocusLoftCut model, "M3_CustomAttachment_2", Array(168, 169)
            FocusLoftCut model, "M3_HeadRelief_2", Array(170, 171)
            FocusLoftCut model, "M3_CustomAttachment_3", Array(172, 173)
            FocusLoftCut model, "M3_HeadRelief_3", Array(174, 175)
        Case "09_CouplingKeeper"
            FocusLoftCut model, "M2p5_KeeperLockSeat", Array(176, 177)
            FocusLoftCut model, "WrenchSlot_M1", Array(178, 179)
            FocusLoftCut model, "WrenchSlot_1", Array(180, 181)
        Case "03_FrontRetainer"
            FocusLoftCut model, "WrenchSlot_M1", Array(182, 183)
            FocusLoftCut model, "WrenchSlot_1", Array(184, 185)
        Case "06_RearRetainer"
            FocusLoftCut model, "WrenchSlot_M1", Array(186, 187)
            FocusLoftCut model, "WrenchSlot_1", Array(188, 189)
        Case "08_ShortFocusRing"
            FocusLoftCut model, "Female_HelicoidGroove_1", Array(190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, _
                210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227, 228, 229, _
                230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240, 241, 242, 243, 244, 245, 246, 247, 248, 249, _
                250, 251, 252, 253, 254, 255, 256, 257, 258, 259, 260, 261, 262, 263, 264, 265, 266, 267, 268, 269, _
                270, 271, 272, 273, 274, 275, 276, 277, 278, 279, 280, 281, 282, 283, 284, 285, 286, 287, 288, 289, _
                290, 291, 292, 293, 294, 295, 296, 297, 298, 299, 300, 301, 302, 303, 304, 305, 306, 307, 308, 309, _
                310, 311, 312, 313, 314, 315, 316, 317, 318, 319, 320, 321, 322, 323, 324, 325, 326, 327, 328, 329, _
                330, 331, 332, 333, 334, 335, 336, 337, 338, 339, 340, 341, 342, 343, 344, 345, 346, 347, 348, 349, _
                350, 351, 352, 353, 354, 355, 356, 357, 358, 359, 360, 361, 362, 363, 364, 365, 366, 367, 368, 369, _
                370, 371, 372)
    End Select
End Sub

Private Function FocusSectionData(ByVal index As Long) As Variant
    Select Case index
        Case 0: FocusSectionData = FSData_0()
        Case 1: FocusSectionData = FSData_1()
        Case 2: FocusSectionData = FSData_2()
        Case 3: FocusSectionData = FSData_3()
        Case 4: FocusSectionData = FSData_4()
        Case 5: FocusSectionData = FSData_5()
        Case 6: FocusSectionData = FSData_6()
        Case 7: FocusSectionData = FSData_7()
        Case 8: FocusSectionData = FSData_8()
        Case 9: FocusSectionData = FSData_9()
        Case 10: FocusSectionData = FSData_10()
        Case 11: FocusSectionData = FSData_11()
        Case 12: FocusSectionData = FSData_12()
        Case 13: FocusSectionData = FSData_13()
        Case 14: FocusSectionData = FSData_14()
        Case 15: FocusSectionData = FSData_15()
        Case 16: FocusSectionData = FSData_16()
        Case 17: FocusSectionData = FSData_17()
        Case 18: FocusSectionData = FSData_18()
        Case 19: FocusSectionData = FSData_19()
        Case 20: FocusSectionData = FSData_20()
        Case 21: FocusSectionData = FSData_21()
        Case 22: FocusSectionData = FSData_22()
        Case 23: FocusSectionData = FSData_23()
        Case 24: FocusSectionData = FSData_24()
        Case 25: FocusSectionData = FSData_25()
        Case 26: FocusSectionData = FSData_26()
        Case 27: FocusSectionData = FSData_27()
        Case 28: FocusSectionData = FSData_28()
        Case 29: FocusSectionData = FSData_29()
        Case 30: FocusSectionData = FSData_30()
        Case 31: FocusSectionData = FSData_31()
        Case 32: FocusSectionData = FSData_32()
        Case 33: FocusSectionData = FSData_33()
        Case 34: FocusSectionData = FSData_34()
        Case 35: FocusSectionData = FSData_35()
        Case 36: FocusSectionData = FSData_36()
        Case 37: FocusSectionData = FSData_37()
        Case 38: FocusSectionData = FSData_38()
        Case 39: FocusSectionData = FSData_39()
        Case 40: FocusSectionData = FSData_40()
        Case 41: FocusSectionData = FSData_41()
        Case 42: FocusSectionData = FSData_42()
        Case 43: FocusSectionData = FSData_43()
        Case 44: FocusSectionData = FSData_44()
        Case 45: FocusSectionData = FSData_45()
        Case 46: FocusSectionData = FSData_46()
        Case 47: FocusSectionData = FSData_47()
        Case 48: FocusSectionData = FSData_48()
        Case 49: FocusSectionData = FSData_49()
        Case 50: FocusSectionData = FSData_50()
        Case 51: FocusSectionData = FSData_51()
        Case 52: FocusSectionData = FSData_52()
        Case 53: FocusSectionData = FSData_53()
        Case 54: FocusSectionData = FSData_54()
        Case 55: FocusSectionData = FSData_55()
        Case 56: FocusSectionData = FSData_56()
        Case 57: FocusSectionData = FSData_57()
        Case 58: FocusSectionData = FSData_58()
        Case 59: FocusSectionData = FSData_59()
        Case 60: FocusSectionData = FSData_60()
        Case 61: FocusSectionData = FSData_61()
        Case 62: FocusSectionData = FSData_62()
        Case 63: FocusSectionData = FSData_63()
        Case 64: FocusSectionData = FSData_64()
        Case 65: FocusSectionData = FSData_65()
        Case 66: FocusSectionData = FSData_66()
        Case 67: FocusSectionData = FSData_67()
        Case 68: FocusSectionData = FSData_68()
        Case 69: FocusSectionData = FSData_69()
        Case 70: FocusSectionData = FSData_70()
        Case 71: FocusSectionData = FSData_71()
        Case 72: FocusSectionData = FSData_72()
        Case 73: FocusSectionData = FSData_73()
        Case 74: FocusSectionData = FSData_74()
        Case 75: FocusSectionData = FSData_75()
        Case 76: FocusSectionData = FSData_76()
        Case 77: FocusSectionData = FSData_77()
        Case 78: FocusSectionData = FSData_78()
        Case 79: FocusSectionData = FSData_79()
        Case 80: FocusSectionData = FSData_80()
        Case 81: FocusSectionData = FSData_81()
        Case 82: FocusSectionData = FSData_82()
        Case 83: FocusSectionData = FSData_83()
        Case 84: FocusSectionData = FSData_84()
        Case 85: FocusSectionData = FSData_85()
        Case 86: FocusSectionData = FSData_86()
        Case 87: FocusSectionData = FSData_87()
        Case 88: FocusSectionData = FSData_88()
        Case 89: FocusSectionData = FSData_89()
        Case 90: FocusSectionData = FSData_90()
        Case 91: FocusSectionData = FSData_91()
        Case 92: FocusSectionData = FSData_92()
        Case 93: FocusSectionData = FSData_93()
        Case 94: FocusSectionData = FSData_94()
        Case 95: FocusSectionData = FSData_95()
        Case 96: FocusSectionData = FSData_96()
        Case 97: FocusSectionData = FSData_97()
        Case 98: FocusSectionData = FSData_98()
        Case 99: FocusSectionData = FSData_99()
        Case 100: FocusSectionData = FSData_100()
        Case 101: FocusSectionData = FSData_101()
        Case 102: FocusSectionData = FSData_102()
        Case 103: FocusSectionData = FSData_103()
        Case 104: FocusSectionData = FSData_104()
        Case 105: FocusSectionData = FSData_105()
        Case 106: FocusSectionData = FSData_106()
        Case 107: FocusSectionData = FSData_107()
        Case 108: FocusSectionData = FSData_108()
        Case 109: FocusSectionData = FSData_109()
        Case 110: FocusSectionData = FSData_110()
        Case 111: FocusSectionData = FSData_111()
        Case 112: FocusSectionData = FSData_112()
        Case 113: FocusSectionData = FSData_113()
        Case 114: FocusSectionData = FSData_114()
        Case 115: FocusSectionData = FSData_115()
        Case 116: FocusSectionData = FSData_116()
        Case 117: FocusSectionData = FSData_117()
        Case 118: FocusSectionData = FSData_118()
        Case 119: FocusSectionData = FSData_119()
        Case 120: FocusSectionData = FSData_120()
        Case 121: FocusSectionData = FSData_121()
        Case 122: FocusSectionData = FSData_122()
        Case 123: FocusSectionData = FSData_123()
        Case 124: FocusSectionData = FSData_124()
        Case 125: FocusSectionData = FSData_125()
        Case 126: FocusSectionData = FSData_126()
        Case 127: FocusSectionData = FSData_127()
        Case 128: FocusSectionData = FSData_128()
        Case 129: FocusSectionData = FSData_129()
        Case 130: FocusSectionData = FSData_130()
        Case 131: FocusSectionData = FSData_131()
        Case 132: FocusSectionData = FSData_132()
        Case 133: FocusSectionData = FSData_133()
        Case 134: FocusSectionData = FSData_134()
        Case 135: FocusSectionData = FSData_135()
        Case 136: FocusSectionData = FSData_136()
        Case 137: FocusSectionData = FSData_137()
        Case 138: FocusSectionData = FSData_138()
        Case 139: FocusSectionData = FSData_139()
        Case 140: FocusSectionData = FSData_140()
        Case 141: FocusSectionData = FSData_141()
        Case 142: FocusSectionData = FSData_142()
        Case 143: FocusSectionData = FSData_143()
        Case 144: FocusSectionData = FSData_144()
        Case 145: FocusSectionData = FSData_145()
        Case 146: FocusSectionData = FSData_146()
        Case 147: FocusSectionData = FSData_147()
        Case 148: FocusSectionData = FSData_148()
        Case 149: FocusSectionData = FSData_149()
        Case 150: FocusSectionData = FSData_150()
        Case 151: FocusSectionData = FSData_151()
        Case 152: FocusSectionData = FSData_152()
        Case 153: FocusSectionData = FSData_153()
        Case 154: FocusSectionData = FSData_154()
        Case 155: FocusSectionData = FSData_155()
        Case 156: FocusSectionData = FSData_156()
        Case 157: FocusSectionData = FSData_157()
        Case 158: FocusSectionData = FSData_158()
        Case 159: FocusSectionData = FSData_159()
        Case 160: FocusSectionData = FSData_160()
        Case 161: FocusSectionData = FSData_161()
        Case 162: FocusSectionData = FSData_162()
        Case 163: FocusSectionData = FSData_163()
        Case 164: FocusSectionData = FSData_164()
        Case 165: FocusSectionData = FSData_165()
        Case 166: FocusSectionData = FSData_166()
        Case 167: FocusSectionData = FSData_167()
        Case 168: FocusSectionData = FSData_168()
        Case 169: FocusSectionData = FSData_169()
        Case 170: FocusSectionData = FSData_170()
        Case 171: FocusSectionData = FSData_171()
        Case 172: FocusSectionData = FSData_172()
        Case 173: FocusSectionData = FSData_173()
        Case 174: FocusSectionData = FSData_174()
        Case 175: FocusSectionData = FSData_175()
        Case 176: FocusSectionData = FSData_176()
        Case 177: FocusSectionData = FSData_177()
        Case 178: FocusSectionData = FSData_178()
        Case 179: FocusSectionData = FSData_179()
        Case 180: FocusSectionData = FSData_180()
        Case 181: FocusSectionData = FSData_181()
        Case 182: FocusSectionData = FSData_182()
        Case 183: FocusSectionData = FSData_183()
        Case 184: FocusSectionData = FSData_184()
        Case 185: FocusSectionData = FSData_185()
        Case 186: FocusSectionData = FSData_186()
        Case 187: FocusSectionData = FSData_187()
        Case 188: FocusSectionData = FSData_188()
        Case 189: FocusSectionData = FSData_189()
        Case 190: FocusSectionData = FSData_190()
        Case 191: FocusSectionData = FSData_191()
        Case 192: FocusSectionData = FSData_192()
        Case 193: FocusSectionData = FSData_193()
        Case 194: FocusSectionData = FSData_194()
        Case 195: FocusSectionData = FSData_195()
        Case 196: FocusSectionData = FSData_196()
        Case 197: FocusSectionData = FSData_197()
        Case 198: FocusSectionData = FSData_198()
        Case 199: FocusSectionData = FSData_199()
        Case 200: FocusSectionData = FSData_200()
        Case 201: FocusSectionData = FSData_201()
        Case 202: FocusSectionData = FSData_202()
        Case 203: FocusSectionData = FSData_203()
        Case 204: FocusSectionData = FSData_204()
        Case 205: FocusSectionData = FSData_205()
        Case 206: FocusSectionData = FSData_206()
        Case 207: FocusSectionData = FSData_207()
        Case 208: FocusSectionData = FSData_208()
        Case 209: FocusSectionData = FSData_209()
        Case 210: FocusSectionData = FSData_210()
        Case 211: FocusSectionData = FSData_211()
        Case 212: FocusSectionData = FSData_212()
        Case 213: FocusSectionData = FSData_213()
        Case 214: FocusSectionData = FSData_214()
        Case 215: FocusSectionData = FSData_215()
        Case 216: FocusSectionData = FSData_216()
        Case 217: FocusSectionData = FSData_217()
        Case 218: FocusSectionData = FSData_218()
        Case 219: FocusSectionData = FSData_219()
        Case 220: FocusSectionData = FSData_220()
        Case 221: FocusSectionData = FSData_221()
        Case 222: FocusSectionData = FSData_222()
        Case 223: FocusSectionData = FSData_223()
        Case 224: FocusSectionData = FSData_224()
        Case 225: FocusSectionData = FSData_225()
        Case 226: FocusSectionData = FSData_226()
        Case 227: FocusSectionData = FSData_227()
        Case 228: FocusSectionData = FSData_228()
        Case 229: FocusSectionData = FSData_229()
        Case 230: FocusSectionData = FSData_230()
        Case 231: FocusSectionData = FSData_231()
        Case 232: FocusSectionData = FSData_232()
        Case 233: FocusSectionData = FSData_233()
        Case 234: FocusSectionData = FSData_234()
        Case 235: FocusSectionData = FSData_235()
        Case 236: FocusSectionData = FSData_236()
        Case 237: FocusSectionData = FSData_237()
        Case 238: FocusSectionData = FSData_238()
        Case 239: FocusSectionData = FSData_239()
        Case 240: FocusSectionData = FSData_240()
        Case 241: FocusSectionData = FSData_241()
        Case 242: FocusSectionData = FSData_242()
        Case 243: FocusSectionData = FSData_243()
        Case 244: FocusSectionData = FSData_244()
        Case 245: FocusSectionData = FSData_245()
        Case 246: FocusSectionData = FSData_246()
        Case 247: FocusSectionData = FSData_247()
        Case 248: FocusSectionData = FSData_248()
        Case 249: FocusSectionData = FSData_249()
        Case 250: FocusSectionData = FSData_250()
        Case 251: FocusSectionData = FSData_251()
        Case 252: FocusSectionData = FSData_252()
        Case 253: FocusSectionData = FSData_253()
        Case 254: FocusSectionData = FSData_254()
        Case 255: FocusSectionData = FSData_255()
        Case 256: FocusSectionData = FSData_256()
        Case 257: FocusSectionData = FSData_257()
        Case 258: FocusSectionData = FSData_258()
        Case 259: FocusSectionData = FSData_259()
        Case 260: FocusSectionData = FSData_260()
        Case 261: FocusSectionData = FSData_261()
        Case 262: FocusSectionData = FSData_262()
        Case 263: FocusSectionData = FSData_263()
        Case 264: FocusSectionData = FSData_264()
        Case 265: FocusSectionData = FSData_265()
        Case 266: FocusSectionData = FSData_266()
        Case 267: FocusSectionData = FSData_267()
        Case 268: FocusSectionData = FSData_268()
        Case 269: FocusSectionData = FSData_269()
        Case 270: FocusSectionData = FSData_270()
        Case 271: FocusSectionData = FSData_271()
        Case 272: FocusSectionData = FSData_272()
        Case 273: FocusSectionData = FSData_273()
        Case 274: FocusSectionData = FSData_274()
        Case 275: FocusSectionData = FSData_275()
        Case 276: FocusSectionData = FSData_276()
        Case 277: FocusSectionData = FSData_277()
        Case 278: FocusSectionData = FSData_278()
        Case 279: FocusSectionData = FSData_279()
        Case 280: FocusSectionData = FSData_280()
        Case 281: FocusSectionData = FSData_281()
        Case 282: FocusSectionData = FSData_282()
        Case 283: FocusSectionData = FSData_283()
        Case 284: FocusSectionData = FSData_284()
        Case 285: FocusSectionData = FSData_285()
        Case 286: FocusSectionData = FSData_286()
        Case 287: FocusSectionData = FSData_287()
        Case 288: FocusSectionData = FSData_288()
        Case 289: FocusSectionData = FSData_289()
        Case 290: FocusSectionData = FSData_290()
        Case 291: FocusSectionData = FSData_291()
        Case 292: FocusSectionData = FSData_292()
        Case 293: FocusSectionData = FSData_293()
        Case 294: FocusSectionData = FSData_294()
        Case 295: FocusSectionData = FSData_295()
        Case 296: FocusSectionData = FSData_296()
        Case 297: FocusSectionData = FSData_297()
        Case 298: FocusSectionData = FSData_298()
        Case 299: FocusSectionData = FSData_299()
        Case 300: FocusSectionData = FSData_300()
        Case 301: FocusSectionData = FSData_301()
        Case 302: FocusSectionData = FSData_302()
        Case 303: FocusSectionData = FSData_303()
        Case 304: FocusSectionData = FSData_304()
        Case 305: FocusSectionData = FSData_305()
        Case 306: FocusSectionData = FSData_306()
        Case 307: FocusSectionData = FSData_307()
        Case 308: FocusSectionData = FSData_308()
        Case 309: FocusSectionData = FSData_309()
        Case 310: FocusSectionData = FSData_310()
        Case 311: FocusSectionData = FSData_311()
        Case 312: FocusSectionData = FSData_312()
        Case 313: FocusSectionData = FSData_313()
        Case 314: FocusSectionData = FSData_314()
        Case 315: FocusSectionData = FSData_315()
        Case 316: FocusSectionData = FSData_316()
        Case 317: FocusSectionData = FSData_317()
        Case 318: FocusSectionData = FSData_318()
        Case 319: FocusSectionData = FSData_319()
        Case 320: FocusSectionData = FSData_320()
        Case 321: FocusSectionData = FSData_321()
        Case 322: FocusSectionData = FSData_322()
        Case 323: FocusSectionData = FSData_323()
        Case 324: FocusSectionData = FSData_324()
        Case 325: FocusSectionData = FSData_325()
        Case 326: FocusSectionData = FSData_326()
        Case 327: FocusSectionData = FSData_327()
        Case 328: FocusSectionData = FSData_328()
        Case 329: FocusSectionData = FSData_329()
        Case 330: FocusSectionData = FSData_330()
        Case 331: FocusSectionData = FSData_331()
        Case 332: FocusSectionData = FSData_332()
        Case 333: FocusSectionData = FSData_333()
        Case 334: FocusSectionData = FSData_334()
        Case 335: FocusSectionData = FSData_335()
        Case 336: FocusSectionData = FSData_336()
        Case 337: FocusSectionData = FSData_337()
        Case 338: FocusSectionData = FSData_338()
        Case 339: FocusSectionData = FSData_339()
        Case 340: FocusSectionData = FSData_340()
        Case 341: FocusSectionData = FSData_341()
        Case 342: FocusSectionData = FSData_342()
        Case 343: FocusSectionData = FSData_343()
        Case 344: FocusSectionData = FSData_344()
        Case 345: FocusSectionData = FSData_345()
        Case 346: FocusSectionData = FSData_346()
        Case 347: FocusSectionData = FSData_347()
        Case 348: FocusSectionData = FSData_348()
        Case 349: FocusSectionData = FSData_349()
        Case 350: FocusSectionData = FSData_350()
        Case 351: FocusSectionData = FSData_351()
        Case 352: FocusSectionData = FSData_352()
        Case 353: FocusSectionData = FSData_353()
        Case 354: FocusSectionData = FSData_354()
        Case 355: FocusSectionData = FSData_355()
        Case 356: FocusSectionData = FSData_356()
        Case 357: FocusSectionData = FSData_357()
        Case 358: FocusSectionData = FSData_358()
        Case 359: FocusSectionData = FSData_359()
        Case 360: FocusSectionData = FSData_360()
        Case 361: FocusSectionData = FSData_361()
        Case 362: FocusSectionData = FSData_362()
        Case 363: FocusSectionData = FSData_363()
        Case 364: FocusSectionData = FSData_364()
        Case 365: FocusSectionData = FSData_365()
        Case 366: FocusSectionData = FSData_366()
        Case 367: FocusSectionData = FSData_367()
        Case 368: FocusSectionData = FSData_368()
        Case 369: FocusSectionData = FSData_369()
        Case 370: FocusSectionData = FSData_370()
        Case 371: FocusSectionData = FSData_371()
        Case 372: FocusSectionData = FSData_372()
        Case Else: Err.Raise vbObjectError + 809, , "Unknown focus section."
    End Select
End Function

Private Function FSData_0() As Variant
    Dim pts(33) As Variant
    pts(0) = Array(25.400000000000, 1.500000000000, -15.500000000000)
    pts(1) = Array(25.400000000000, 1.471177920605, -15.207364516976)
    pts(2) = Array(25.400000000000, 1.385819298767, -14.925974851452)
    pts(3) = Array(25.400000000000, 1.247204418454, -14.666644650471)
    pts(4) = Array(25.400000000000, 1.060660171780, -14.439339828220)
    pts(5) = Array(25.400000000000, 0.833355349529, -14.252795581546)
    pts(6) = Array(25.400000000000, 0.574025148548, -14.114180701233)
    pts(7) = Array(25.400000000000, 0.292635483024, -14.028822079395)
    pts(8) = Array(25.400000000000, 0.000000000000, -14.000000000000)
    pts(9) = Array(25.400000000000, -0.292635483024, -14.028822079395)
    pts(10) = Array(25.400000000000, -0.574025148548, -14.114180701233)
    pts(11) = Array(25.400000000000, -0.833355349529, -14.252795581546)
    pts(12) = Array(25.400000000000, -1.060660171780, -14.439339828220)
    pts(13) = Array(25.400000000000, -1.247204418454, -14.666644650471)
    pts(14) = Array(25.400000000000, -1.385819298767, -14.925974851452)
    pts(15) = Array(25.400000000000, -1.471177920605, -15.207364516976)
    pts(16) = Array(25.400000000000, -1.500000000000, -15.500000000000)
    pts(17) = Array(25.400000000000, -1.500000000000, -20.000000000000)
    pts(18) = Array(25.400000000000, -1.471177920605, -20.292635483024)
    pts(19) = Array(25.400000000000, -1.385819298767, -20.574025148548)
    pts(20) = Array(25.400000000000, -1.247204418454, -20.833355349529)
    pts(21) = Array(25.400000000000, -1.060660171780, -21.060660171780)
    pts(22) = Array(25.400000000000, -0.833355349529, -21.247204418454)
    pts(23) = Array(25.400000000000, -0.574025148548, -21.385819298767)
    pts(24) = Array(25.400000000000, -0.292635483024, -21.471177920605)
    pts(25) = Array(25.400000000000, -0.000000000000, -21.500000000000)
    pts(26) = Array(25.400000000000, 0.292635483024, -21.471177920605)
    pts(27) = Array(25.400000000000, 0.574025148548, -21.385819298767)
    pts(28) = Array(25.400000000000, 0.833355349529, -21.247204418454)
    pts(29) = Array(25.400000000000, 1.060660171780, -21.060660171780)
    pts(30) = Array(25.400000000000, 1.247204418454, -20.833355349529)
    pts(31) = Array(25.400000000000, 1.385819298767, -20.574025148548)
    pts(32) = Array(25.400000000000, 1.471177920605, -20.292635483024)
    pts(33) = Array(25.400000000000, 1.500000000000, -20.000000000000)
    FSData_0 = pts
End Function

Private Function FSData_1() As Variant
    Dim pts(33) As Variant
    pts(0) = Array(27.200000000000, 1.500000000000, -15.500000000000)
    pts(1) = Array(27.200000000000, 1.471177920605, -15.207364516976)
    pts(2) = Array(27.200000000000, 1.385819298767, -14.925974851452)
    pts(3) = Array(27.200000000000, 1.247204418454, -14.666644650471)
    pts(4) = Array(27.200000000000, 1.060660171780, -14.439339828220)
    pts(5) = Array(27.200000000000, 0.833355349529, -14.252795581546)
    pts(6) = Array(27.200000000000, 0.574025148548, -14.114180701233)
    pts(7) = Array(27.200000000000, 0.292635483024, -14.028822079395)
    pts(8) = Array(27.200000000000, 0.000000000000, -14.000000000000)
    pts(9) = Array(27.200000000000, -0.292635483024, -14.028822079395)
    pts(10) = Array(27.200000000000, -0.574025148548, -14.114180701233)
    pts(11) = Array(27.200000000000, -0.833355349529, -14.252795581546)
    pts(12) = Array(27.200000000000, -1.060660171780, -14.439339828220)
    pts(13) = Array(27.200000000000, -1.247204418454, -14.666644650471)
    pts(14) = Array(27.200000000000, -1.385819298767, -14.925974851452)
    pts(15) = Array(27.200000000000, -1.471177920605, -15.207364516976)
    pts(16) = Array(27.200000000000, -1.500000000000, -15.500000000000)
    pts(17) = Array(27.200000000000, -1.500000000000, -20.000000000000)
    pts(18) = Array(27.200000000000, -1.471177920605, -20.292635483024)
    pts(19) = Array(27.200000000000, -1.385819298767, -20.574025148548)
    pts(20) = Array(27.200000000000, -1.247204418454, -20.833355349529)
    pts(21) = Array(27.200000000000, -1.060660171780, -21.060660171780)
    pts(22) = Array(27.200000000000, -0.833355349529, -21.247204418454)
    pts(23) = Array(27.200000000000, -0.574025148548, -21.385819298767)
    pts(24) = Array(27.200000000000, -0.292635483024, -21.471177920605)
    pts(25) = Array(27.200000000000, -0.000000000000, -21.500000000000)
    pts(26) = Array(27.200000000000, 0.292635483024, -21.471177920605)
    pts(27) = Array(27.200000000000, 0.574025148548, -21.385819298767)
    pts(28) = Array(27.200000000000, 0.833355349529, -21.247204418454)
    pts(29) = Array(27.200000000000, 1.060660171780, -21.060660171780)
    pts(30) = Array(27.200000000000, 1.247204418454, -20.833355349529)
    pts(31) = Array(27.200000000000, 1.385819298767, -20.574025148548)
    pts(32) = Array(27.200000000000, 1.471177920605, -20.292635483024)
    pts(33) = Array(27.200000000000, 1.500000000000, -20.000000000000)
    FSData_1 = pts
End Function

Private Function FSData_2() As Variant
    Dim pts(33) As Variant
    pts(0) = Array(-25.400000000000, -1.500000000000, -15.500000000000)
    pts(1) = Array(-25.400000000000, -1.471177920605, -15.207364516976)
    pts(2) = Array(-25.400000000000, -1.385819298767, -14.925974851452)
    pts(3) = Array(-25.400000000000, -1.247204418454, -14.666644650471)
    pts(4) = Array(-25.400000000000, -1.060660171780, -14.439339828220)
    pts(5) = Array(-25.400000000000, -0.833355349529, -14.252795581546)
    pts(6) = Array(-25.400000000000, -0.574025148548, -14.114180701233)
    pts(7) = Array(-25.400000000000, -0.292635483024, -14.028822079395)
    pts(8) = Array(-25.400000000000, 0.000000000000, -14.000000000000)
    pts(9) = Array(-25.400000000000, 0.292635483024, -14.028822079395)
    pts(10) = Array(-25.400000000000, 0.574025148548, -14.114180701233)
    pts(11) = Array(-25.400000000000, 0.833355349529, -14.252795581546)
    pts(12) = Array(-25.400000000000, 1.060660171780, -14.439339828220)
    pts(13) = Array(-25.400000000000, 1.247204418454, -14.666644650471)
    pts(14) = Array(-25.400000000000, 1.385819298767, -14.925974851452)
    pts(15) = Array(-25.400000000000, 1.471177920605, -15.207364516976)
    pts(16) = Array(-25.400000000000, 1.500000000000, -15.500000000000)
    pts(17) = Array(-25.400000000000, 1.500000000000, -20.000000000000)
    pts(18) = Array(-25.400000000000, 1.471177920605, -20.292635483024)
    pts(19) = Array(-25.400000000000, 1.385819298767, -20.574025148548)
    pts(20) = Array(-25.400000000000, 1.247204418454, -20.833355349529)
    pts(21) = Array(-25.400000000000, 1.060660171780, -21.060660171780)
    pts(22) = Array(-25.400000000000, 0.833355349529, -21.247204418454)
    pts(23) = Array(-25.400000000000, 0.574025148548, -21.385819298767)
    pts(24) = Array(-25.400000000000, 0.292635483024, -21.471177920605)
    pts(25) = Array(-25.400000000000, 0.000000000000, -21.500000000000)
    pts(26) = Array(-25.400000000000, -0.292635483024, -21.471177920605)
    pts(27) = Array(-25.400000000000, -0.574025148548, -21.385819298767)
    pts(28) = Array(-25.400000000000, -0.833355349529, -21.247204418454)
    pts(29) = Array(-25.400000000000, -1.060660171780, -21.060660171780)
    pts(30) = Array(-25.400000000000, -1.247204418454, -20.833355349529)
    pts(31) = Array(-25.400000000000, -1.385819298767, -20.574025148548)
    pts(32) = Array(-25.400000000000, -1.471177920605, -20.292635483024)
    pts(33) = Array(-25.400000000000, -1.500000000000, -20.000000000000)
    FSData_2 = pts
End Function

Private Function FSData_3() As Variant
    Dim pts(33) As Variant
    pts(0) = Array(-27.200000000000, -1.500000000000, -15.500000000000)
    pts(1) = Array(-27.200000000000, -1.471177920605, -15.207364516976)
    pts(2) = Array(-27.200000000000, -1.385819298767, -14.925974851452)
    pts(3) = Array(-27.200000000000, -1.247204418454, -14.666644650471)
    pts(4) = Array(-27.200000000000, -1.060660171780, -14.439339828220)
    pts(5) = Array(-27.200000000000, -0.833355349529, -14.252795581546)
    pts(6) = Array(-27.200000000000, -0.574025148548, -14.114180701233)
    pts(7) = Array(-27.200000000000, -0.292635483024, -14.028822079395)
    pts(8) = Array(-27.200000000000, 0.000000000000, -14.000000000000)
    pts(9) = Array(-27.200000000000, 0.292635483024, -14.028822079395)
    pts(10) = Array(-27.200000000000, 0.574025148548, -14.114180701233)
    pts(11) = Array(-27.200000000000, 0.833355349529, -14.252795581546)
    pts(12) = Array(-27.200000000000, 1.060660171780, -14.439339828220)
    pts(13) = Array(-27.200000000000, 1.247204418454, -14.666644650471)
    pts(14) = Array(-27.200000000000, 1.385819298767, -14.925974851452)
    pts(15) = Array(-27.200000000000, 1.471177920605, -15.207364516976)
    pts(16) = Array(-27.200000000000, 1.500000000000, -15.500000000000)
    pts(17) = Array(-27.200000000000, 1.500000000000, -20.000000000000)
    pts(18) = Array(-27.200000000000, 1.471177920605, -20.292635483024)
    pts(19) = Array(-27.200000000000, 1.385819298767, -20.574025148548)
    pts(20) = Array(-27.200000000000, 1.247204418454, -20.833355349529)
    pts(21) = Array(-27.200000000000, 1.060660171780, -21.060660171780)
    pts(22) = Array(-27.200000000000, 0.833355349529, -21.247204418454)
    pts(23) = Array(-27.200000000000, 0.574025148548, -21.385819298767)
    pts(24) = Array(-27.200000000000, 0.292635483024, -21.471177920605)
    pts(25) = Array(-27.200000000000, 0.000000000000, -21.500000000000)
    pts(26) = Array(-27.200000000000, -0.292635483024, -21.471177920605)
    pts(27) = Array(-27.200000000000, -0.574025148548, -21.385819298767)
    pts(28) = Array(-27.200000000000, -0.833355349529, -21.247204418454)
    pts(29) = Array(-27.200000000000, -1.060660171780, -21.060660171780)
    pts(30) = Array(-27.200000000000, -1.247204418454, -20.833355349529)
    pts(31) = Array(-27.200000000000, -1.385819298767, -20.574025148548)
    pts(32) = Array(-27.200000000000, -1.471177920605, -20.292635483024)
    pts(33) = Array(-27.200000000000, -1.500000000000, -20.000000000000)
    FSData_3 = pts
End Function

Private Function FSData_4() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-1.275000000000, 24.600000000000, -42.800000000000)
    pts(1) = Array(-1.264092198252, 24.600000000000, -42.633579104919)
    pts(2) = Array(-1.231555428519, 24.600000000000, -42.470005717494)
    pts(3) = Array(-1.177946403952, 24.600000000000, -42.312078623735)
    pts(4) = Array(-1.104182389825, 24.600000000000, -42.162500000000)
    pts(5) = Array(-1.011525508871, 24.600000000000, -42.023829178014)
    pts(6) = Array(-0.901561146013, 24.600000000000, -41.898438853987)
    pts(7) = Array(-0.776170821986, 24.600000000000, -41.788474491129)
    pts(8) = Array(-0.637500000000, 24.600000000000, -41.695817610175)
    pts(9) = Array(-0.487921376265, 24.600000000000, -41.622053596048)
    pts(10) = Array(-0.329994282506, 24.600000000000, -41.568444571481)
    pts(11) = Array(-0.166420895081, 24.600000000000, -41.535907801748)
    pts(12) = Array(0.000000000000, 24.600000000000, -41.525000000000)
    pts(13) = Array(0.166420895081, 24.600000000000, -41.535907801748)
    pts(14) = Array(0.329994282506, 24.600000000000, -41.568444571481)
    pts(15) = Array(0.487921376265, 24.600000000000, -41.622053596048)
    pts(16) = Array(0.637500000000, 24.600000000000, -41.695817610175)
    pts(17) = Array(0.776170821986, 24.600000000000, -41.788474491129)
    pts(18) = Array(0.901561146013, 24.600000000000, -41.898438853987)
    pts(19) = Array(1.011525508871, 24.600000000000, -42.023829178014)
    pts(20) = Array(1.104182389825, 24.600000000000, -42.162500000000)
    pts(21) = Array(1.177946403952, 24.600000000000, -42.312078623735)
    pts(22) = Array(1.231555428519, 24.600000000000, -42.470005717494)
    pts(23) = Array(1.264092198252, 24.600000000000, -42.633579104919)
    pts(24) = Array(1.275000000000, 24.600000000000, -42.800000000000)
    pts(25) = Array(1.264092198252, 24.600000000000, -42.966420895081)
    pts(26) = Array(1.231555428519, 24.600000000000, -43.129994282506)
    pts(27) = Array(1.177946403952, 24.600000000000, -43.287921376265)
    pts(28) = Array(1.104182389825, 24.600000000000, -43.437500000000)
    pts(29) = Array(1.011525508871, 24.600000000000, -43.576170821986)
    pts(30) = Array(0.901561146013, 24.600000000000, -43.701561146013)
    pts(31) = Array(0.776170821986, 24.600000000000, -43.811525508871)
    pts(32) = Array(0.637500000000, 24.600000000000, -43.904182389825)
    pts(33) = Array(0.487921376265, 24.600000000000, -43.977946403952)
    pts(34) = Array(0.329994282506, 24.600000000000, -44.031555428519)
    pts(35) = Array(0.166420895081, 24.600000000000, -44.064092198252)
    pts(36) = Array(0.000000000000, 24.600000000000, -44.075000000000)
    pts(37) = Array(-0.166420895081, 24.600000000000, -44.064092198252)
    pts(38) = Array(-0.329994282506, 24.600000000000, -44.031555428519)
    pts(39) = Array(-0.487921376265, 24.600000000000, -43.977946403952)
    pts(40) = Array(-0.637500000000, 24.600000000000, -43.904182389825)
    pts(41) = Array(-0.776170821986, 24.600000000000, -43.811525508871)
    pts(42) = Array(-0.901561146013, 24.600000000000, -43.701561146013)
    pts(43) = Array(-1.011525508871, 24.600000000000, -43.576170821986)
    pts(44) = Array(-1.104182389825, 24.600000000000, -43.437500000000)
    pts(45) = Array(-1.177946403952, 24.600000000000, -43.287921376265)
    pts(46) = Array(-1.231555428519, 24.600000000000, -43.129994282506)
    pts(47) = Array(-1.264092198252, 24.600000000000, -42.966420895081)
    FSData_4 = pts
End Function

Private Function FSData_5() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-1.275000000000, 25.200000000000, -42.800000000000)
    pts(1) = Array(-1.264092198252, 25.200000000000, -42.633579104919)
    pts(2) = Array(-1.231555428519, 25.200000000000, -42.470005717494)
    pts(3) = Array(-1.177946403952, 25.200000000000, -42.312078623735)
    pts(4) = Array(-1.104182389825, 25.200000000000, -42.162500000000)
    pts(5) = Array(-1.011525508871, 25.200000000000, -42.023829178014)
    pts(6) = Array(-0.901561146013, 25.200000000000, -41.898438853987)
    pts(7) = Array(-0.776170821986, 25.200000000000, -41.788474491129)
    pts(8) = Array(-0.637500000000, 25.200000000000, -41.695817610175)
    pts(9) = Array(-0.487921376265, 25.200000000000, -41.622053596048)
    pts(10) = Array(-0.329994282506, 25.200000000000, -41.568444571481)
    pts(11) = Array(-0.166420895081, 25.200000000000, -41.535907801748)
    pts(12) = Array(0.000000000000, 25.200000000000, -41.525000000000)
    pts(13) = Array(0.166420895081, 25.200000000000, -41.535907801748)
    pts(14) = Array(0.329994282506, 25.200000000000, -41.568444571481)
    pts(15) = Array(0.487921376265, 25.200000000000, -41.622053596048)
    pts(16) = Array(0.637500000000, 25.200000000000, -41.695817610175)
    pts(17) = Array(0.776170821986, 25.200000000000, -41.788474491129)
    pts(18) = Array(0.901561146013, 25.200000000000, -41.898438853987)
    pts(19) = Array(1.011525508871, 25.200000000000, -42.023829178014)
    pts(20) = Array(1.104182389825, 25.200000000000, -42.162500000000)
    pts(21) = Array(1.177946403952, 25.200000000000, -42.312078623735)
    pts(22) = Array(1.231555428519, 25.200000000000, -42.470005717494)
    pts(23) = Array(1.264092198252, 25.200000000000, -42.633579104919)
    pts(24) = Array(1.275000000000, 25.200000000000, -42.800000000000)
    pts(25) = Array(1.264092198252, 25.200000000000, -42.966420895081)
    pts(26) = Array(1.231555428519, 25.200000000000, -43.129994282506)
    pts(27) = Array(1.177946403952, 25.200000000000, -43.287921376265)
    pts(28) = Array(1.104182389825, 25.200000000000, -43.437500000000)
    pts(29) = Array(1.011525508871, 25.200000000000, -43.576170821986)
    pts(30) = Array(0.901561146013, 25.200000000000, -43.701561146013)
    pts(31) = Array(0.776170821986, 25.200000000000, -43.811525508871)
    pts(32) = Array(0.637500000000, 25.200000000000, -43.904182389825)
    pts(33) = Array(0.487921376265, 25.200000000000, -43.977946403952)
    pts(34) = Array(0.329994282506, 25.200000000000, -44.031555428519)
    pts(35) = Array(0.166420895081, 25.200000000000, -44.064092198252)
    pts(36) = Array(0.000000000000, 25.200000000000, -44.075000000000)
    pts(37) = Array(-0.166420895081, 25.200000000000, -44.064092198252)
    pts(38) = Array(-0.329994282506, 25.200000000000, -44.031555428519)
    pts(39) = Array(-0.487921376265, 25.200000000000, -43.977946403952)
    pts(40) = Array(-0.637500000000, 25.200000000000, -43.904182389825)
    pts(41) = Array(-0.776170821986, 25.200000000000, -43.811525508871)
    pts(42) = Array(-0.901561146013, 25.200000000000, -43.701561146013)
    pts(43) = Array(-1.011525508871, 25.200000000000, -43.576170821986)
    pts(44) = Array(-1.104182389825, 25.200000000000, -43.437500000000)
    pts(45) = Array(-1.177946403952, 25.200000000000, -43.287921376265)
    pts(46) = Array(-1.231555428519, 25.200000000000, -43.129994282506)
    pts(47) = Array(-1.264092198252, 25.200000000000, -42.966420895081)
    FSData_5 = pts
End Function

Private Function FSData_6() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(26.900000000000, 1.525000000000, -20.000000000000)
    pts(1) = Array(26.900000000000, 1.511953413595, -19.800947556864)
    pts(2) = Array(26.900000000000, 1.473036885091, -19.605300956219)
    pts(3) = Array(26.900000000000, 1.408916287080, -19.416407765643)
    pts(4) = Array(26.900000000000, 1.320688740771, -19.237500000000)
    pts(5) = Array(26.900000000000, 1.209863843944, -19.071638820762)
    pts(6) = Array(26.900000000000, 1.078337841309, -18.921662158691)
    pts(7) = Array(26.900000000000, 0.928361179238, -18.790136156056)
    pts(8) = Array(26.900000000000, 0.762500000000, -18.679311259229)
    pts(9) = Array(26.900000000000, 0.583592234357, -18.591083712920)
    pts(10) = Array(26.900000000000, 0.394699043781, -18.526963114909)
    pts(11) = Array(26.900000000000, 0.199052443136, -18.488046586405)
    pts(12) = Array(26.900000000000, 0.000000000000, -18.475000000000)
    pts(13) = Array(26.900000000000, -0.199052443136, -18.488046586405)
    pts(14) = Array(26.900000000000, -0.394699043781, -18.526963114909)
    pts(15) = Array(26.900000000000, -0.583592234357, -18.591083712920)
    pts(16) = Array(26.900000000000, -0.762500000000, -18.679311259229)
    pts(17) = Array(26.900000000000, -0.928361179238, -18.790136156056)
    pts(18) = Array(26.900000000000, -1.078337841309, -18.921662158691)
    pts(19) = Array(26.900000000000, -1.209863843944, -19.071638820762)
    pts(20) = Array(26.900000000000, -1.320688740771, -19.237500000000)
    pts(21) = Array(26.900000000000, -1.408916287080, -19.416407765643)
    pts(22) = Array(26.900000000000, -1.473036885091, -19.605300956219)
    pts(23) = Array(26.900000000000, -1.511953413595, -19.800947556864)
    pts(24) = Array(26.900000000000, -1.525000000000, -20.000000000000)
    pts(25) = Array(26.900000000000, -1.511953413595, -20.199052443136)
    pts(26) = Array(26.900000000000, -1.473036885091, -20.394699043781)
    pts(27) = Array(26.900000000000, -1.408916287080, -20.583592234357)
    pts(28) = Array(26.900000000000, -1.320688740771, -20.762500000000)
    pts(29) = Array(26.900000000000, -1.209863843944, -20.928361179238)
    pts(30) = Array(26.900000000000, -1.078337841309, -21.078337841309)
    pts(31) = Array(26.900000000000, -0.928361179238, -21.209863843944)
    pts(32) = Array(26.900000000000, -0.762500000000, -21.320688740771)
    pts(33) = Array(26.900000000000, -0.583592234357, -21.408916287080)
    pts(34) = Array(26.900000000000, -0.394699043781, -21.473036885091)
    pts(35) = Array(26.900000000000, -0.199052443136, -21.511953413595)
    pts(36) = Array(26.900000000000, -0.000000000000, -21.525000000000)
    pts(37) = Array(26.900000000000, 0.199052443136, -21.511953413595)
    pts(38) = Array(26.900000000000, 0.394699043781, -21.473036885091)
    pts(39) = Array(26.900000000000, 0.583592234357, -21.408916287080)
    pts(40) = Array(26.900000000000, 0.762500000000, -21.320688740771)
    pts(41) = Array(26.900000000000, 0.928361179238, -21.209863843944)
    pts(42) = Array(26.900000000000, 1.078337841309, -21.078337841309)
    pts(43) = Array(26.900000000000, 1.209863843944, -20.928361179238)
    pts(44) = Array(26.900000000000, 1.320688740771, -20.762500000000)
    pts(45) = Array(26.900000000000, 1.408916287080, -20.583592234357)
    pts(46) = Array(26.900000000000, 1.473036885091, -20.394699043781)
    pts(47) = Array(26.900000000000, 1.511953413595, -20.199052443136)
    FSData_6 = pts
End Function

Private Function FSData_7() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(31.300000000000, 1.525000000000, -20.000000000000)
    pts(1) = Array(31.300000000000, 1.511953413595, -19.800947556864)
    pts(2) = Array(31.300000000000, 1.473036885091, -19.605300956219)
    pts(3) = Array(31.300000000000, 1.408916287080, -19.416407765643)
    pts(4) = Array(31.300000000000, 1.320688740771, -19.237500000000)
    pts(5) = Array(31.300000000000, 1.209863843944, -19.071638820762)
    pts(6) = Array(31.300000000000, 1.078337841309, -18.921662158691)
    pts(7) = Array(31.300000000000, 0.928361179238, -18.790136156056)
    pts(8) = Array(31.300000000000, 0.762500000000, -18.679311259229)
    pts(9) = Array(31.300000000000, 0.583592234357, -18.591083712920)
    pts(10) = Array(31.300000000000, 0.394699043781, -18.526963114909)
    pts(11) = Array(31.300000000000, 0.199052443136, -18.488046586405)
    pts(12) = Array(31.300000000000, 0.000000000000, -18.475000000000)
    pts(13) = Array(31.300000000000, -0.199052443136, -18.488046586405)
    pts(14) = Array(31.300000000000, -0.394699043781, -18.526963114909)
    pts(15) = Array(31.300000000000, -0.583592234357, -18.591083712920)
    pts(16) = Array(31.300000000000, -0.762500000000, -18.679311259229)
    pts(17) = Array(31.300000000000, -0.928361179238, -18.790136156056)
    pts(18) = Array(31.300000000000, -1.078337841309, -18.921662158691)
    pts(19) = Array(31.300000000000, -1.209863843944, -19.071638820762)
    pts(20) = Array(31.300000000000, -1.320688740771, -19.237500000000)
    pts(21) = Array(31.300000000000, -1.408916287080, -19.416407765643)
    pts(22) = Array(31.300000000000, -1.473036885091, -19.605300956219)
    pts(23) = Array(31.300000000000, -1.511953413595, -19.800947556864)
    pts(24) = Array(31.300000000000, -1.525000000000, -20.000000000000)
    pts(25) = Array(31.300000000000, -1.511953413595, -20.199052443136)
    pts(26) = Array(31.300000000000, -1.473036885091, -20.394699043781)
    pts(27) = Array(31.300000000000, -1.408916287080, -20.583592234357)
    pts(28) = Array(31.300000000000, -1.320688740771, -20.762500000000)
    pts(29) = Array(31.300000000000, -1.209863843944, -20.928361179238)
    pts(30) = Array(31.300000000000, -1.078337841309, -21.078337841309)
    pts(31) = Array(31.300000000000, -0.928361179238, -21.209863843944)
    pts(32) = Array(31.300000000000, -0.762500000000, -21.320688740771)
    pts(33) = Array(31.300000000000, -0.583592234357, -21.408916287080)
    pts(34) = Array(31.300000000000, -0.394699043781, -21.473036885091)
    pts(35) = Array(31.300000000000, -0.199052443136, -21.511953413595)
    pts(36) = Array(31.300000000000, -0.000000000000, -21.525000000000)
    pts(37) = Array(31.300000000000, 0.199052443136, -21.511953413595)
    pts(38) = Array(31.300000000000, 0.394699043781, -21.473036885091)
    pts(39) = Array(31.300000000000, 0.583592234357, -21.408916287080)
    pts(40) = Array(31.300000000000, 0.762500000000, -21.320688740771)
    pts(41) = Array(31.300000000000, 0.928361179238, -21.209863843944)
    pts(42) = Array(31.300000000000, 1.078337841309, -21.078337841309)
    pts(43) = Array(31.300000000000, 1.209863843944, -20.928361179238)
    pts(44) = Array(31.300000000000, 1.320688740771, -20.762500000000)
    pts(45) = Array(31.300000000000, 1.408916287080, -20.583592234357)
    pts(46) = Array(31.300000000000, 1.473036885091, -20.394699043781)
    pts(47) = Array(31.300000000000, 1.511953413595, -20.199052443136)
    FSData_7 = pts
End Function

Private Function FSData_8() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-26.900000000000, -1.525000000000, -20.000000000000)
    pts(1) = Array(-26.900000000000, -1.511953413595, -19.800947556864)
    pts(2) = Array(-26.900000000000, -1.473036885091, -19.605300956219)
    pts(3) = Array(-26.900000000000, -1.408916287080, -19.416407765643)
    pts(4) = Array(-26.900000000000, -1.320688740771, -19.237500000000)
    pts(5) = Array(-26.900000000000, -1.209863843944, -19.071638820762)
    pts(6) = Array(-26.900000000000, -1.078337841309, -18.921662158691)
    pts(7) = Array(-26.900000000000, -0.928361179238, -18.790136156056)
    pts(8) = Array(-26.900000000000, -0.762500000000, -18.679311259229)
    pts(9) = Array(-26.900000000000, -0.583592234357, -18.591083712920)
    pts(10) = Array(-26.900000000000, -0.394699043781, -18.526963114909)
    pts(11) = Array(-26.900000000000, -0.199052443136, -18.488046586405)
    pts(12) = Array(-26.900000000000, 0.000000000000, -18.475000000000)
    pts(13) = Array(-26.900000000000, 0.199052443136, -18.488046586405)
    pts(14) = Array(-26.900000000000, 0.394699043781, -18.526963114909)
    pts(15) = Array(-26.900000000000, 0.583592234357, -18.591083712920)
    pts(16) = Array(-26.900000000000, 0.762500000000, -18.679311259229)
    pts(17) = Array(-26.900000000000, 0.928361179238, -18.790136156056)
    pts(18) = Array(-26.900000000000, 1.078337841309, -18.921662158691)
    pts(19) = Array(-26.900000000000, 1.209863843944, -19.071638820762)
    pts(20) = Array(-26.900000000000, 1.320688740771, -19.237500000000)
    pts(21) = Array(-26.900000000000, 1.408916287080, -19.416407765643)
    pts(22) = Array(-26.900000000000, 1.473036885091, -19.605300956219)
    pts(23) = Array(-26.900000000000, 1.511953413595, -19.800947556864)
    pts(24) = Array(-26.900000000000, 1.525000000000, -20.000000000000)
    pts(25) = Array(-26.900000000000, 1.511953413595, -20.199052443136)
    pts(26) = Array(-26.900000000000, 1.473036885091, -20.394699043781)
    pts(27) = Array(-26.900000000000, 1.408916287080, -20.583592234357)
    pts(28) = Array(-26.900000000000, 1.320688740771, -20.762500000000)
    pts(29) = Array(-26.900000000000, 1.209863843944, -20.928361179238)
    pts(30) = Array(-26.900000000000, 1.078337841309, -21.078337841309)
    pts(31) = Array(-26.900000000000, 0.928361179238, -21.209863843944)
    pts(32) = Array(-26.900000000000, 0.762500000000, -21.320688740771)
    pts(33) = Array(-26.900000000000, 0.583592234357, -21.408916287080)
    pts(34) = Array(-26.900000000000, 0.394699043781, -21.473036885091)
    pts(35) = Array(-26.900000000000, 0.199052443136, -21.511953413595)
    pts(36) = Array(-26.900000000000, 0.000000000000, -21.525000000000)
    pts(37) = Array(-26.900000000000, -0.199052443136, -21.511953413595)
    pts(38) = Array(-26.900000000000, -0.394699043781, -21.473036885091)
    pts(39) = Array(-26.900000000000, -0.583592234357, -21.408916287080)
    pts(40) = Array(-26.900000000000, -0.762500000000, -21.320688740771)
    pts(41) = Array(-26.900000000000, -0.928361179238, -21.209863843944)
    pts(42) = Array(-26.900000000000, -1.078337841309, -21.078337841309)
    pts(43) = Array(-26.900000000000, -1.209863843944, -20.928361179238)
    pts(44) = Array(-26.900000000000, -1.320688740771, -20.762500000000)
    pts(45) = Array(-26.900000000000, -1.408916287080, -20.583592234357)
    pts(46) = Array(-26.900000000000, -1.473036885091, -20.394699043781)
    pts(47) = Array(-26.900000000000, -1.511953413595, -20.199052443136)
    FSData_8 = pts
End Function

Private Function FSData_9() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-31.300000000000, -1.525000000000, -20.000000000000)
    pts(1) = Array(-31.300000000000, -1.511953413595, -19.800947556864)
    pts(2) = Array(-31.300000000000, -1.473036885091, -19.605300956219)
    pts(3) = Array(-31.300000000000, -1.408916287080, -19.416407765643)
    pts(4) = Array(-31.300000000000, -1.320688740771, -19.237500000000)
    pts(5) = Array(-31.300000000000, -1.209863843944, -19.071638820762)
    pts(6) = Array(-31.300000000000, -1.078337841309, -18.921662158691)
    pts(7) = Array(-31.300000000000, -0.928361179238, -18.790136156056)
    pts(8) = Array(-31.300000000000, -0.762500000000, -18.679311259229)
    pts(9) = Array(-31.300000000000, -0.583592234357, -18.591083712920)
    pts(10) = Array(-31.300000000000, -0.394699043781, -18.526963114909)
    pts(11) = Array(-31.300000000000, -0.199052443136, -18.488046586405)
    pts(12) = Array(-31.300000000000, 0.000000000000, -18.475000000000)
    pts(13) = Array(-31.300000000000, 0.199052443136, -18.488046586405)
    pts(14) = Array(-31.300000000000, 0.394699043781, -18.526963114909)
    pts(15) = Array(-31.300000000000, 0.583592234357, -18.591083712920)
    pts(16) = Array(-31.300000000000, 0.762500000000, -18.679311259229)
    pts(17) = Array(-31.300000000000, 0.928361179238, -18.790136156056)
    pts(18) = Array(-31.300000000000, 1.078337841309, -18.921662158691)
    pts(19) = Array(-31.300000000000, 1.209863843944, -19.071638820762)
    pts(20) = Array(-31.300000000000, 1.320688740771, -19.237500000000)
    pts(21) = Array(-31.300000000000, 1.408916287080, -19.416407765643)
    pts(22) = Array(-31.300000000000, 1.473036885091, -19.605300956219)
    pts(23) = Array(-31.300000000000, 1.511953413595, -19.800947556864)
    pts(24) = Array(-31.300000000000, 1.525000000000, -20.000000000000)
    pts(25) = Array(-31.300000000000, 1.511953413595, -20.199052443136)
    pts(26) = Array(-31.300000000000, 1.473036885091, -20.394699043781)
    pts(27) = Array(-31.300000000000, 1.408916287080, -20.583592234357)
    pts(28) = Array(-31.300000000000, 1.320688740771, -20.762500000000)
    pts(29) = Array(-31.300000000000, 1.209863843944, -20.928361179238)
    pts(30) = Array(-31.300000000000, 1.078337841309, -21.078337841309)
    pts(31) = Array(-31.300000000000, 0.928361179238, -21.209863843944)
    pts(32) = Array(-31.300000000000, 0.762500000000, -21.320688740771)
    pts(33) = Array(-31.300000000000, 0.583592234357, -21.408916287080)
    pts(34) = Array(-31.300000000000, 0.394699043781, -21.473036885091)
    pts(35) = Array(-31.300000000000, 0.199052443136, -21.511953413595)
    pts(36) = Array(-31.300000000000, 0.000000000000, -21.525000000000)
    pts(37) = Array(-31.300000000000, -0.199052443136, -21.511953413595)
    pts(38) = Array(-31.300000000000, -0.394699043781, -21.473036885091)
    pts(39) = Array(-31.300000000000, -0.583592234357, -21.408916287080)
    pts(40) = Array(-31.300000000000, -0.762500000000, -21.320688740771)
    pts(41) = Array(-31.300000000000, -0.928361179238, -21.209863843944)
    pts(42) = Array(-31.300000000000, -1.078337841309, -21.078337841309)
    pts(43) = Array(-31.300000000000, -1.209863843944, -20.928361179238)
    pts(44) = Array(-31.300000000000, -1.320688740771, -20.762500000000)
    pts(45) = Array(-31.300000000000, -1.408916287080, -20.583592234357)
    pts(46) = Array(-31.300000000000, -1.473036885091, -20.394699043781)
    pts(47) = Array(-31.300000000000, -1.511953413595, -20.199052443136)
    FSData_9 = pts
End Function

Private Function FSData_10() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(31.025000000000, 0.000000000000, -7.200000000000)
    pts(1) = Array(31.011953413595, 0.199052443136, -7.200000000000)
    pts(2) = Array(30.973036885091, 0.394699043781, -7.200000000000)
    pts(3) = Array(30.908916287080, 0.583592234357, -7.200000000000)
    pts(4) = Array(30.820688740771, 0.762500000000, -7.200000000000)
    pts(5) = Array(30.709863843944, 0.928361179238, -7.200000000000)
    pts(6) = Array(30.578337841309, 1.078337841309, -7.200000000000)
    pts(7) = Array(30.428361179238, 1.209863843944, -7.200000000000)
    pts(8) = Array(30.262500000000, 1.320688740771, -7.200000000000)
    pts(9) = Array(30.083592234357, 1.408916287080, -7.200000000000)
    pts(10) = Array(29.894699043781, 1.473036885091, -7.200000000000)
    pts(11) = Array(29.699052443136, 1.511953413595, -7.200000000000)
    pts(12) = Array(29.500000000000, 1.525000000000, -7.200000000000)
    pts(13) = Array(29.300947556864, 1.511953413595, -7.200000000000)
    pts(14) = Array(29.105300956219, 1.473036885091, -7.200000000000)
    pts(15) = Array(28.916407765643, 1.408916287080, -7.200000000000)
    pts(16) = Array(28.737500000000, 1.320688740771, -7.200000000000)
    pts(17) = Array(28.571638820762, 1.209863843944, -7.200000000000)
    pts(18) = Array(28.421662158691, 1.078337841309, -7.200000000000)
    pts(19) = Array(28.290136156056, 0.928361179238, -7.200000000000)
    pts(20) = Array(28.179311259229, 0.762500000000, -7.200000000000)
    pts(21) = Array(28.091083712920, 0.583592234357, -7.200000000000)
    pts(22) = Array(28.026963114909, 0.394699043781, -7.200000000000)
    pts(23) = Array(27.988046586405, 0.199052443136, -7.200000000000)
    pts(24) = Array(27.975000000000, 0.000000000000, -7.200000000000)
    pts(25) = Array(27.988046586405, -0.199052443136, -7.200000000000)
    pts(26) = Array(28.026963114909, -0.394699043781, -7.200000000000)
    pts(27) = Array(28.091083712920, -0.583592234357, -7.200000000000)
    pts(28) = Array(28.179311259229, -0.762500000000, -7.200000000000)
    pts(29) = Array(28.290136156056, -0.928361179238, -7.200000000000)
    pts(30) = Array(28.421662158691, -1.078337841309, -7.200000000000)
    pts(31) = Array(28.571638820762, -1.209863843944, -7.200000000000)
    pts(32) = Array(28.737500000000, -1.320688740771, -7.200000000000)
    pts(33) = Array(28.916407765643, -1.408916287080, -7.200000000000)
    pts(34) = Array(29.105300956219, -1.473036885091, -7.200000000000)
    pts(35) = Array(29.300947556864, -1.511953413595, -7.200000000000)
    pts(36) = Array(29.500000000000, -1.525000000000, -7.200000000000)
    pts(37) = Array(29.699052443136, -1.511953413595, -7.200000000000)
    pts(38) = Array(29.894699043781, -1.473036885091, -7.200000000000)
    pts(39) = Array(30.083592234357, -1.408916287080, -7.200000000000)
    pts(40) = Array(30.262500000000, -1.320688740771, -7.200000000000)
    pts(41) = Array(30.428361179238, -1.209863843944, -7.200000000000)
    pts(42) = Array(30.578337841309, -1.078337841309, -7.200000000000)
    pts(43) = Array(30.709863843944, -0.928361179238, -7.200000000000)
    pts(44) = Array(30.820688740771, -0.762500000000, -7.200000000000)
    pts(45) = Array(30.908916287080, -0.583592234357, -7.200000000000)
    pts(46) = Array(30.973036885091, -0.394699043781, -7.200000000000)
    pts(47) = Array(31.011953413595, -0.199052443136, -7.200000000000)
    FSData_10 = pts
End Function

Private Function FSData_11() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(31.025000000000, 0.000000000000, -1.900000000000)
    pts(1) = Array(31.011953413595, 0.199052443136, -1.900000000000)
    pts(2) = Array(30.973036885091, 0.394699043781, -1.900000000000)
    pts(3) = Array(30.908916287080, 0.583592234357, -1.900000000000)
    pts(4) = Array(30.820688740771, 0.762500000000, -1.900000000000)
    pts(5) = Array(30.709863843944, 0.928361179238, -1.900000000000)
    pts(6) = Array(30.578337841309, 1.078337841309, -1.900000000000)
    pts(7) = Array(30.428361179238, 1.209863843944, -1.900000000000)
    pts(8) = Array(30.262500000000, 1.320688740771, -1.900000000000)
    pts(9) = Array(30.083592234357, 1.408916287080, -1.900000000000)
    pts(10) = Array(29.894699043781, 1.473036885091, -1.900000000000)
    pts(11) = Array(29.699052443136, 1.511953413595, -1.900000000000)
    pts(12) = Array(29.500000000000, 1.525000000000, -1.900000000000)
    pts(13) = Array(29.300947556864, 1.511953413595, -1.900000000000)
    pts(14) = Array(29.105300956219, 1.473036885091, -1.900000000000)
    pts(15) = Array(28.916407765643, 1.408916287080, -1.900000000000)
    pts(16) = Array(28.737500000000, 1.320688740771, -1.900000000000)
    pts(17) = Array(28.571638820762, 1.209863843944, -1.900000000000)
    pts(18) = Array(28.421662158691, 1.078337841309, -1.900000000000)
    pts(19) = Array(28.290136156056, 0.928361179238, -1.900000000000)
    pts(20) = Array(28.179311259229, 0.762500000000, -1.900000000000)
    pts(21) = Array(28.091083712920, 0.583592234357, -1.900000000000)
    pts(22) = Array(28.026963114909, 0.394699043781, -1.900000000000)
    pts(23) = Array(27.988046586405, 0.199052443136, -1.900000000000)
    pts(24) = Array(27.975000000000, 0.000000000000, -1.900000000000)
    pts(25) = Array(27.988046586405, -0.199052443136, -1.900000000000)
    pts(26) = Array(28.026963114909, -0.394699043781, -1.900000000000)
    pts(27) = Array(28.091083712920, -0.583592234357, -1.900000000000)
    pts(28) = Array(28.179311259229, -0.762500000000, -1.900000000000)
    pts(29) = Array(28.290136156056, -0.928361179238, -1.900000000000)
    pts(30) = Array(28.421662158691, -1.078337841309, -1.900000000000)
    pts(31) = Array(28.571638820762, -1.209863843944, -1.900000000000)
    pts(32) = Array(28.737500000000, -1.320688740771, -1.900000000000)
    pts(33) = Array(28.916407765643, -1.408916287080, -1.900000000000)
    pts(34) = Array(29.105300956219, -1.473036885091, -1.900000000000)
    pts(35) = Array(29.300947556864, -1.511953413595, -1.900000000000)
    pts(36) = Array(29.500000000000, -1.525000000000, -1.900000000000)
    pts(37) = Array(29.699052443136, -1.511953413595, -1.900000000000)
    pts(38) = Array(29.894699043781, -1.473036885091, -1.900000000000)
    pts(39) = Array(30.083592234357, -1.408916287080, -1.900000000000)
    pts(40) = Array(30.262500000000, -1.320688740771, -1.900000000000)
    pts(41) = Array(30.428361179238, -1.209863843944, -1.900000000000)
    pts(42) = Array(30.578337841309, -1.078337841309, -1.900000000000)
    pts(43) = Array(30.709863843944, -0.928361179238, -1.900000000000)
    pts(44) = Array(30.820688740771, -0.762500000000, -1.900000000000)
    pts(45) = Array(30.908916287080, -0.583592234357, -1.900000000000)
    pts(46) = Array(30.973036885091, -0.394699043781, -1.900000000000)
    pts(47) = Array(31.011953413595, -0.199052443136, -1.900000000000)
    FSData_11 = pts
End Function

Private Function FSData_12() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-13.225000000000, 25.547749411641, -7.200000000000)
    pts(1) = Array(-13.238046586405, 25.746801854777, -7.200000000000)
    pts(2) = Array(-13.276963114909, 25.942448455422, -7.200000000000)
    pts(3) = Array(-13.341083712920, 26.131341645998, -7.200000000000)
    pts(4) = Array(-13.429311259229, 26.310249411641, -7.200000000000)
    pts(5) = Array(-13.540136156056, 26.476110590879, -7.200000000000)
    pts(6) = Array(-13.671662158691, 26.626087252950, -7.200000000000)
    pts(7) = Array(-13.821638820762, 26.757613255585, -7.200000000000)
    pts(8) = Array(-13.987500000000, 26.868438152412, -7.200000000000)
    pts(9) = Array(-14.166407765643, 26.956665698721, -7.200000000000)
    pts(10) = Array(-14.355300956219, 27.020786296732, -7.200000000000)
    pts(11) = Array(-14.550947556864, 27.059702825236, -7.200000000000)
    pts(12) = Array(-14.750000000000, 27.072749411641, -7.200000000000)
    pts(13) = Array(-14.949052443136, 27.059702825236, -7.200000000000)
    pts(14) = Array(-15.144699043781, 27.020786296732, -7.200000000000)
    pts(15) = Array(-15.333592234357, 26.956665698721, -7.200000000000)
    pts(16) = Array(-15.512500000000, 26.868438152412, -7.200000000000)
    pts(17) = Array(-15.678361179238, 26.757613255585, -7.200000000000)
    pts(18) = Array(-15.828337841309, 26.626087252950, -7.200000000000)
    pts(19) = Array(-15.959863843944, 26.476110590879, -7.200000000000)
    pts(20) = Array(-16.070688740771, 26.310249411641, -7.200000000000)
    pts(21) = Array(-16.158916287080, 26.131341645998, -7.200000000000)
    pts(22) = Array(-16.223036885091, 25.942448455422, -7.200000000000)
    pts(23) = Array(-16.261953413595, 25.746801854777, -7.200000000000)
    pts(24) = Array(-16.275000000000, 25.547749411641, -7.200000000000)
    pts(25) = Array(-16.261953413595, 25.348696968505, -7.200000000000)
    pts(26) = Array(-16.223036885091, 25.153050367860, -7.200000000000)
    pts(27) = Array(-16.158916287080, 24.964157177284, -7.200000000000)
    pts(28) = Array(-16.070688740771, 24.785249411641, -7.200000000000)
    pts(29) = Array(-15.959863843944, 24.619388232403, -7.200000000000)
    pts(30) = Array(-15.828337841309, 24.469411570331, -7.200000000000)
    pts(31) = Array(-15.678361179238, 24.337885567697, -7.200000000000)
    pts(32) = Array(-15.512500000000, 24.227060670870, -7.200000000000)
    pts(33) = Array(-15.333592234357, 24.138833124561, -7.200000000000)
    pts(34) = Array(-15.144699043781, 24.074712526550, -7.200000000000)
    pts(35) = Array(-14.949052443136, 24.035795998046, -7.200000000000)
    pts(36) = Array(-14.750000000000, 24.022749411641, -7.200000000000)
    pts(37) = Array(-14.550947556864, 24.035795998046, -7.200000000000)
    pts(38) = Array(-14.355300956219, 24.074712526550, -7.200000000000)
    pts(39) = Array(-14.166407765643, 24.138833124561, -7.200000000000)
    pts(40) = Array(-13.987500000000, 24.227060670870, -7.200000000000)
    pts(41) = Array(-13.821638820762, 24.337885567697, -7.200000000000)
    pts(42) = Array(-13.671662158691, 24.469411570331, -7.200000000000)
    pts(43) = Array(-13.540136156056, 24.619388232403, -7.200000000000)
    pts(44) = Array(-13.429311259229, 24.785249411641, -7.200000000000)
    pts(45) = Array(-13.341083712920, 24.964157177284, -7.200000000000)
    pts(46) = Array(-13.276963114909, 25.153050367860, -7.200000000000)
    pts(47) = Array(-13.238046586405, 25.348696968505, -7.200000000000)
    FSData_12 = pts
End Function

Private Function FSData_13() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-13.225000000000, 25.547749411641, -1.900000000000)
    pts(1) = Array(-13.238046586405, 25.746801854777, -1.900000000000)
    pts(2) = Array(-13.276963114909, 25.942448455422, -1.900000000000)
    pts(3) = Array(-13.341083712920, 26.131341645998, -1.900000000000)
    pts(4) = Array(-13.429311259229, 26.310249411641, -1.900000000000)
    pts(5) = Array(-13.540136156056, 26.476110590879, -1.900000000000)
    pts(6) = Array(-13.671662158691, 26.626087252950, -1.900000000000)
    pts(7) = Array(-13.821638820762, 26.757613255585, -1.900000000000)
    pts(8) = Array(-13.987500000000, 26.868438152412, -1.900000000000)
    pts(9) = Array(-14.166407765643, 26.956665698721, -1.900000000000)
    pts(10) = Array(-14.355300956219, 27.020786296732, -1.900000000000)
    pts(11) = Array(-14.550947556864, 27.059702825236, -1.900000000000)
    pts(12) = Array(-14.750000000000, 27.072749411641, -1.900000000000)
    pts(13) = Array(-14.949052443136, 27.059702825236, -1.900000000000)
    pts(14) = Array(-15.144699043781, 27.020786296732, -1.900000000000)
    pts(15) = Array(-15.333592234357, 26.956665698721, -1.900000000000)
    pts(16) = Array(-15.512500000000, 26.868438152412, -1.900000000000)
    pts(17) = Array(-15.678361179238, 26.757613255585, -1.900000000000)
    pts(18) = Array(-15.828337841309, 26.626087252950, -1.900000000000)
    pts(19) = Array(-15.959863843944, 26.476110590879, -1.900000000000)
    pts(20) = Array(-16.070688740771, 26.310249411641, -1.900000000000)
    pts(21) = Array(-16.158916287080, 26.131341645998, -1.900000000000)
    pts(22) = Array(-16.223036885091, 25.942448455422, -1.900000000000)
    pts(23) = Array(-16.261953413595, 25.746801854777, -1.900000000000)
    pts(24) = Array(-16.275000000000, 25.547749411641, -1.900000000000)
    pts(25) = Array(-16.261953413595, 25.348696968505, -1.900000000000)
    pts(26) = Array(-16.223036885091, 25.153050367860, -1.900000000000)
    pts(27) = Array(-16.158916287080, 24.964157177284, -1.900000000000)
    pts(28) = Array(-16.070688740771, 24.785249411641, -1.900000000000)
    pts(29) = Array(-15.959863843944, 24.619388232403, -1.900000000000)
    pts(30) = Array(-15.828337841309, 24.469411570331, -1.900000000000)
    pts(31) = Array(-15.678361179238, 24.337885567697, -1.900000000000)
    pts(32) = Array(-15.512500000000, 24.227060670870, -1.900000000000)
    pts(33) = Array(-15.333592234357, 24.138833124561, -1.900000000000)
    pts(34) = Array(-15.144699043781, 24.074712526550, -1.900000000000)
    pts(35) = Array(-14.949052443136, 24.035795998046, -1.900000000000)
    pts(36) = Array(-14.750000000000, 24.022749411641, -1.900000000000)
    pts(37) = Array(-14.550947556864, 24.035795998046, -1.900000000000)
    pts(38) = Array(-14.355300956219, 24.074712526550, -1.900000000000)
    pts(39) = Array(-14.166407765643, 24.138833124561, -1.900000000000)
    pts(40) = Array(-13.987500000000, 24.227060670870, -1.900000000000)
    pts(41) = Array(-13.821638820762, 24.337885567697, -1.900000000000)
    pts(42) = Array(-13.671662158691, 24.469411570331, -1.900000000000)
    pts(43) = Array(-13.540136156056, 24.619388232403, -1.900000000000)
    pts(44) = Array(-13.429311259229, 24.785249411641, -1.900000000000)
    pts(45) = Array(-13.341083712920, 24.964157177284, -1.900000000000)
    pts(46) = Array(-13.276963114909, 25.153050367860, -1.900000000000)
    pts(47) = Array(-13.238046586405, 25.348696968505, -1.900000000000)
    FSData_13 = pts
End Function

Private Function FSData_14() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-13.225000000000, -25.547749411641, -7.200000000000)
    pts(1) = Array(-13.238046586405, -25.348696968505, -7.200000000000)
    pts(2) = Array(-13.276963114909, -25.153050367860, -7.200000000000)
    pts(3) = Array(-13.341083712920, -24.964157177284, -7.200000000000)
    pts(4) = Array(-13.429311259229, -24.785249411641, -7.200000000000)
    pts(5) = Array(-13.540136156056, -24.619388232403, -7.200000000000)
    pts(6) = Array(-13.671662158691, -24.469411570331, -7.200000000000)
    pts(7) = Array(-13.821638820762, -24.337885567697, -7.200000000000)
    pts(8) = Array(-13.987500000000, -24.227060670870, -7.200000000000)
    pts(9) = Array(-14.166407765643, -24.138833124561, -7.200000000000)
    pts(10) = Array(-14.355300956219, -24.074712526550, -7.200000000000)
    pts(11) = Array(-14.550947556864, -24.035795998046, -7.200000000000)
    pts(12) = Array(-14.750000000000, -24.022749411641, -7.200000000000)
    pts(13) = Array(-14.949052443136, -24.035795998046, -7.200000000000)
    pts(14) = Array(-15.144699043781, -24.074712526550, -7.200000000000)
    pts(15) = Array(-15.333592234357, -24.138833124561, -7.200000000000)
    pts(16) = Array(-15.512500000000, -24.227060670870, -7.200000000000)
    pts(17) = Array(-15.678361179238, -24.337885567697, -7.200000000000)
    pts(18) = Array(-15.828337841309, -24.469411570331, -7.200000000000)
    pts(19) = Array(-15.959863843944, -24.619388232403, -7.200000000000)
    pts(20) = Array(-16.070688740771, -24.785249411641, -7.200000000000)
    pts(21) = Array(-16.158916287080, -24.964157177284, -7.200000000000)
    pts(22) = Array(-16.223036885091, -25.153050367860, -7.200000000000)
    pts(23) = Array(-16.261953413595, -25.348696968505, -7.200000000000)
    pts(24) = Array(-16.275000000000, -25.547749411641, -7.200000000000)
    pts(25) = Array(-16.261953413595, -25.746801854777, -7.200000000000)
    pts(26) = Array(-16.223036885091, -25.942448455422, -7.200000000000)
    pts(27) = Array(-16.158916287080, -26.131341645998, -7.200000000000)
    pts(28) = Array(-16.070688740771, -26.310249411641, -7.200000000000)
    pts(29) = Array(-15.959863843944, -26.476110590879, -7.200000000000)
    pts(30) = Array(-15.828337841309, -26.626087252950, -7.200000000000)
    pts(31) = Array(-15.678361179238, -26.757613255585, -7.200000000000)
    pts(32) = Array(-15.512500000000, -26.868438152412, -7.200000000000)
    pts(33) = Array(-15.333592234357, -26.956665698721, -7.200000000000)
    pts(34) = Array(-15.144699043781, -27.020786296732, -7.200000000000)
    pts(35) = Array(-14.949052443136, -27.059702825236, -7.200000000000)
    pts(36) = Array(-14.750000000000, -27.072749411641, -7.200000000000)
    pts(37) = Array(-14.550947556864, -27.059702825236, -7.200000000000)
    pts(38) = Array(-14.355300956219, -27.020786296732, -7.200000000000)
    pts(39) = Array(-14.166407765643, -26.956665698721, -7.200000000000)
    pts(40) = Array(-13.987500000000, -26.868438152412, -7.200000000000)
    pts(41) = Array(-13.821638820762, -26.757613255585, -7.200000000000)
    pts(42) = Array(-13.671662158691, -26.626087252950, -7.200000000000)
    pts(43) = Array(-13.540136156056, -26.476110590879, -7.200000000000)
    pts(44) = Array(-13.429311259229, -26.310249411641, -7.200000000000)
    pts(45) = Array(-13.341083712920, -26.131341645998, -7.200000000000)
    pts(46) = Array(-13.276963114909, -25.942448455422, -7.200000000000)
    pts(47) = Array(-13.238046586405, -25.746801854777, -7.200000000000)
    FSData_14 = pts
End Function

Private Function FSData_15() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-13.225000000000, -25.547749411641, -1.900000000000)
    pts(1) = Array(-13.238046586405, -25.348696968505, -1.900000000000)
    pts(2) = Array(-13.276963114909, -25.153050367860, -1.900000000000)
    pts(3) = Array(-13.341083712920, -24.964157177284, -1.900000000000)
    pts(4) = Array(-13.429311259229, -24.785249411641, -1.900000000000)
    pts(5) = Array(-13.540136156056, -24.619388232403, -1.900000000000)
    pts(6) = Array(-13.671662158691, -24.469411570331, -1.900000000000)
    pts(7) = Array(-13.821638820762, -24.337885567697, -1.900000000000)
    pts(8) = Array(-13.987500000000, -24.227060670870, -1.900000000000)
    pts(9) = Array(-14.166407765643, -24.138833124561, -1.900000000000)
    pts(10) = Array(-14.355300956219, -24.074712526550, -1.900000000000)
    pts(11) = Array(-14.550947556864, -24.035795998046, -1.900000000000)
    pts(12) = Array(-14.750000000000, -24.022749411641, -1.900000000000)
    pts(13) = Array(-14.949052443136, -24.035795998046, -1.900000000000)
    pts(14) = Array(-15.144699043781, -24.074712526550, -1.900000000000)
    pts(15) = Array(-15.333592234357, -24.138833124561, -1.900000000000)
    pts(16) = Array(-15.512500000000, -24.227060670870, -1.900000000000)
    pts(17) = Array(-15.678361179238, -24.337885567697, -1.900000000000)
    pts(18) = Array(-15.828337841309, -24.469411570331, -1.900000000000)
    pts(19) = Array(-15.959863843944, -24.619388232403, -1.900000000000)
    pts(20) = Array(-16.070688740771, -24.785249411641, -1.900000000000)
    pts(21) = Array(-16.158916287080, -24.964157177284, -1.900000000000)
    pts(22) = Array(-16.223036885091, -25.153050367860, -1.900000000000)
    pts(23) = Array(-16.261953413595, -25.348696968505, -1.900000000000)
    pts(24) = Array(-16.275000000000, -25.547749411641, -1.900000000000)
    pts(25) = Array(-16.261953413595, -25.746801854777, -1.900000000000)
    pts(26) = Array(-16.223036885091, -25.942448455422, -1.900000000000)
    pts(27) = Array(-16.158916287080, -26.131341645998, -1.900000000000)
    pts(28) = Array(-16.070688740771, -26.310249411641, -1.900000000000)
    pts(29) = Array(-15.959863843944, -26.476110590879, -1.900000000000)
    pts(30) = Array(-15.828337841309, -26.626087252950, -1.900000000000)
    pts(31) = Array(-15.678361179238, -26.757613255585, -1.900000000000)
    pts(32) = Array(-15.512500000000, -26.868438152412, -1.900000000000)
    pts(33) = Array(-15.333592234357, -26.956665698721, -1.900000000000)
    pts(34) = Array(-15.144699043781, -27.020786296732, -1.900000000000)
    pts(35) = Array(-14.949052443136, -27.059702825236, -1.900000000000)
    pts(36) = Array(-14.750000000000, -27.072749411641, -1.900000000000)
    pts(37) = Array(-14.550947556864, -27.059702825236, -1.900000000000)
    pts(38) = Array(-14.355300956219, -27.020786296732, -1.900000000000)
    pts(39) = Array(-14.166407765643, -26.956665698721, -1.900000000000)
    pts(40) = Array(-13.987500000000, -26.868438152412, -1.900000000000)
    pts(41) = Array(-13.821638820762, -26.757613255585, -1.900000000000)
    pts(42) = Array(-13.671662158691, -26.626087252950, -1.900000000000)
    pts(43) = Array(-13.540136156056, -26.476110590879, -1.900000000000)
    pts(44) = Array(-13.429311259229, -26.310249411641, -1.900000000000)
    pts(45) = Array(-13.341083712920, -26.131341645998, -1.900000000000)
    pts(46) = Array(-13.276963114909, -25.942448455422, -1.900000000000)
    pts(47) = Array(-13.238046586405, -25.746801854777, -1.900000000000)
    FSData_15 = pts
End Function

Private Function FSData_16() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.594116628998, -17.868671669691, -30.940000000000)
    pts(1) = Array(24.594116628998, -17.868671669691, -27.860000000000)
    pts(2) = Array(25.322231923936, -18.397678396754, -27.620000000000)
    pts(3) = Array(25.322231923936, -18.397678396754, -31.180000000000)
    FSData_16 = pts
End Function

Private Function FSData_17() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.062558156396, -20.914194404654, -30.815510204082)
    pts(1) = Array(22.062558156396, -20.914194404654, -27.735510204082)
    pts(2) = Array(22.715725996552, -21.533364633739, -27.495510204082)
    pts(3) = Array(22.715725996552, -21.533364633739, -31.055510204082)
    FSData_17 = pts
End Function

Private Function FSData_18() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.156574155112, -23.604780588679, -30.691020408163)
    pts(1) = Array(19.156574155112, -23.604780588679, -27.611020408163)
    pts(2) = Array(19.723709574178, -24.303606329791, -27.371020408163)
    pts(3) = Array(19.723709574178, -24.303606329791, -30.931020408163)
    FSData_18 = pts
End Function

Private Function FSData_19() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(15.925482323863, -25.894768057512, -30.566530612245)
    pts(1) = Array(15.925482323863, -25.894768057512, -27.486530612245)
    pts(2) = Array(16.396960418977, -26.661389480267, -27.246530612245)
    pts(3) = Array(16.396960418977, -26.661389480267, -30.806530612245)
    FSData_19 = pts
End Function

Private Function FSData_20() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(12.424117793666, -27.745293241361, -30.442040816327)
    pts(1) = Array(12.424117793666, -27.745293241361, -27.362040816327)
    pts(2) = Array(12.791937070452, -28.566699949164, -27.122040816327)
    pts(3) = Array(12.791937070452, -28.566699949164, -30.682040816327)
    FSData_20 = pts
End Function

Private Function FSData_21() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(8.711902516230, -29.124950721807, -30.317551020408)
    pts(1) = Array(8.711902516230, -29.124950721807, -27.237551020408)
    pts(2) = Array(8.969820682829, -29.987202552387, -26.997551020408)
    pts(3) = Array(8.969820682829, -29.987202552387, -30.557551020408)
    FSData_21 = pts
End Function

Private Function FSData_22() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(4.851836809015, -30.010326215799, -30.193061224490)
    pts(1) = Array(4.851836809015, -30.010326215799, -27.113061224490)
    pts(2) = Array(4.995476714545, -30.898789820872, -26.873061224490)
    pts(3) = Array(4.995476714545, -30.898789820872, -30.433061224490)
    FSData_22 = pts
End Function

Private Function FSData_23() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(0.909430171550, -30.386393941419, -30.068571428571)
    pts(1) = Array(0.909430171550, -30.386393941419, -26.988571428571)
    pts(2) = Array(0.936354091102, -31.285991130474, -26.748571428571)
    pts(3) = Array(0.936354091102, -31.285991130474, -30.308571428571)
    FSData_23 = pts
End Function

Private Function FSData_24() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-3.048410481788, -30.246771621688, -29.944081632653)
    pts(1) = Array(-3.048410481788, -30.246771621688, -26.864081632653)
    pts(2) = Array(-3.138659476315, -31.142235255224, -26.624081632653)
    pts(3) = Array(-3.138659476315, -31.142235255224, -30.184081632653)
    FSData_24 = pts
End Function

Private Function FSData_25() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.954516304641, -29.593828798729, -29.819591836735)
    pts(1) = Array(-6.954516304641, -29.593828798729, -26.739591836735)
    pts(2) = Array(-7.160406589976, -30.469961888165, -26.499591836735)
    pts(3) = Array(-7.160406589976, -30.469961888165, -30.059591836735)
    FSData_25 = pts
End Function

Private Function FSData_26() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.742596446808, -28.438646620066, -29.695102040816)
    pts(1) = Array(-10.742596446808, -28.438646620066, -26.615102040816)
    pts(2) = Array(-11.060633841615, -29.280580237108, -26.375102040816)
    pts(3) = Array(-11.060633841615, -29.280580237108, -29.935102040816)
    FSData_26 = pts
End Function

Private Function FSData_27() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.348363081834, -26.800829779539, -29.570612244898)
    pts(1) = Array(-14.348363081834, -26.800829779539, -26.490612244898)
    pts(2) = Array(-14.773150146757, -27.594275398013, -26.250612244898)
    pts(3) = Array(-14.773150146757, -27.594275398013, -29.810612244898)
    FSData_27 = pts
End Function

Private Function FSData_28() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.710622441116, -24.708173804396, -29.446122448980)
    pts(1) = Array(-17.710622441116, -24.708173804396, -26.366122448980)
    pts(2) = Array(-18.234950079175, -25.439665792026, -26.126122448980)
    pts(3) = Array(-18.234950079175, -25.439665792026, -29.686122448980)
    FSData_28 = pts
End Function

Private Function FSData_29() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.772313338492, -22.196193335063, -29.321632653061)
    pts(1) = Array(-20.772313338492, -22.196193335063, -26.241632653061)
    pts(2) = Array(-21.387283141276, -22.853317479852, -26.001632653061)
    pts(3) = Array(-21.387283141276, -22.853317479852, -29.561632653061)
    FSData_29 = pts
End Function

Private Function FSData_30() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.481475560428, -19.307519403202, -29.197142857143)
    pts(1) = Array(-23.481475560428, -19.307519403202, -26.117142857143)
    pts(2) = Array(-24.176650823730, -19.879123596060, -25.877142857143)
    pts(3) = Array(-24.176650823730, -19.879123596060, -29.437142857143)
    FSData_30 = pts
End Function

Private Function FSData_31() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.792131687169, -16.091175936883, -29.072653061224)
    pts(1) = Array(-25.792131687169, -16.091175936883, -25.992653061224)
    pts(2) = Array(-26.555714533171, -16.567559435015, -25.752653061224)
    pts(3) = Array(-26.555714533171, -16.567559435015, -29.312653061224)
    FSData_31 = pts
End Function

Private Function FSData_32() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.665067379394, -12.601747771385, -28.948163265306)
    pts(1) = Array(-27.665067379394, -12.601747771385, -25.868163265306)
    pts(2) = Array(-28.484098979442, -12.974825830406, -25.628163265306)
    pts(3) = Array(-28.484098979442, -12.974825830406, -29.188163265306)
    FSData_32 = pts
End Function

Private Function FSData_33() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.068496888061, -8.898454285367, -28.823673469388)
    pts(1) = Array(-29.068496888061, -8.898454285367, -25.743673469388)
    pts(2) = Array(-29.929077388037, -9.161895366184, -25.503673469388)
    pts(3) = Array(-29.929077388037, -9.161895366184, -29.063673469388)
    FSData_33 = pts
End Function

Private Function FSData_34() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.978602493032, -5.044144383818, -28.699183673469)
    pts(1) = Array(-29.978602493032, -5.044144383818, -25.619183673469)
    pts(2) = Array(-30.866126908944, -5.193477605707, -25.379183673469)
    pts(3) = Array(-30.866126908944, -5.193477605707, -28.939183673469)
    FSData_34 = pts
End Function

Private Function FSData_35() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.379938715594, -1.104229884004, -28.574693877551)
    pts(1) = Array(-30.379938715594, -1.104229884004, -25.494693877551)
    pts(2) = Array(-31.279344795990, -1.136920900307, -25.254693877551)
    pts(3) = Array(-31.279344795990, -1.136920900307, -28.814693877551)
    FSData_35 = pts
End Function

Private Function FSData_36() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.265694444987, 2.854424593973, -28.450204081633)
    pts(1) = Array(-30.265694444987, 2.854424593973, -25.370204081633)
    pts(2) = Array(-31.161718293687, 2.938930585242, -25.130204081633)
    pts(3) = Array(-31.161718293687, 2.938930585242, -28.690204081633)
    FSData_36 = pts
End Function

Private Function FSData_37() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.637808530327, 6.764636392272, -28.325714285714)
    pts(1) = Array(-29.637808530327, 6.764636392272, -25.245714285714)
    pts(2) = Array(-30.515243651291, 6.964905232833, -25.005714285714)
    pts(3) = Array(-30.515243651291, 6.964905232833, -28.565714285714)
    FSData_37 = pts
End Function

Private Function FSData_38() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.506936876243, 10.560044977835, -28.201224489796)
    pts(1) = Array(-28.506936876243, 10.560044977835, -25.121224489796)
    pts(2) = Array(-29.350892244290, 10.872677888363, -24.881224489796)
    pts(3) = Array(-29.350892244290, 10.872677888363, -28.441224489796)
    FSData_38 = pts
End Function

Private Function FSData_39() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-26.892271600617, 14.176238152580, -28.076734693878)
    pts(1) = Array(-26.892271600617, 14.176238152580, -24.996734693878)
    pts(2) = Array(-27.688424378267, 14.595929413676, -24.756734693878)
    pts(3) = Array(-27.688424378267, 14.595929413676, -28.316734693878)
    FSData_39 = pts
End Function

Private Function FSData_40() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.821215323544, 17.551845198220, -27.952244897959)
    pts(1) = Array(-24.821215323544, 17.551845198220, -24.872244897959)
    pts(2) = Array(-25.556053935096, 18.071472194220, -24.632244897959)
    pts(3) = Array(-25.556053935096, 18.071472194220, -28.192244897959)
    FSData_40 = pts
End Function

Private Function FSData_41() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.328916115160, 20.629578403888, -27.827755102041)
    pts(1) = Array(-22.328916115160, 20.629578403888, -24.747755102041)
    pts(2) = Array(-22.989969552780, 21.240322501371, -24.507755102041)
    pts(3) = Array(-22.989969552780, 21.240322501371, -28.067755102041)
    FSData_41 = pts
End Function

Private Function FSData_42() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.457670994792, 23.357205300687, -27.703265306122)
    pts(1) = Array(-19.457670994792, 23.357205300687, -24.623265306122)
    pts(2) = Array(-20.033720465032, 24.048701510247, -24.383265306122)
    pts(3) = Array(-20.033720465032, 24.048701510247, -27.943265306122)
    FSData_42 = pts
End Function

Private Function FSData_43() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.256208104686, 25.688435103314, -27.578775510204)
    pts(1) = Array(-16.256208104686, 25.688435103314, -24.498775510204)
    pts(2) = Array(-16.737477423574, 26.448947984662, -24.258775510204)
    pts(3) = Array(-16.737477423574, 26.448947984662, -27.818775510204)
    FSData_43 = pts
End Function

Private Function FSData_44() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.778859740611, 27.583704314863, -27.454285714286)
    pts(1) = Array(-12.778859740611, 27.583704314863, -24.374285714286)
    pts(2) = Array(-13.157181246089, 28.400327139974, -24.134285714286)
    pts(3) = Array(-13.157181246089, 28.400327139974, -27.694285714286)
    FSData_44 = pts
End Function

Private Function FSData_45() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.084640273902, 29.010848162262, -27.329795918367)
    pts(1) = Array(-9.084640273902, 29.010848162262, -24.249795918367)
    pts(2) = Array(-9.353593439905, 29.869721956539, -24.009795918367)
    pts(3) = Array(-9.353593439905, 29.869721956539, -27.569795918367)
    FSData_45 = pts
End Function

Private Function FSData_46() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.236244613602, 29.945646467334, -27.205306122449)
    pts(1) = Array(-5.236244613602, 29.945646467334, -24.125306122449)
    pts(2) = Array(-5.391265013347, 30.832195211432, -23.885306122449)
    pts(3) = Array(-5.391265013347, 30.832195211432, -27.445306122449)
    FSData_46 = pts
End Function

Private Function FSData_47() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.298984205893, 30.372234689480, -27.080816326531)
    pts(1) = Array(-1.298984205893, 30.372234689480, -24.000816326531)
    pts(2) = Array(-1.337440975147, 31.271412690155, -23.760816326531)
    pts(3) = Array(-1.337440975147, 31.271412690155, -27.320816326531)
    FSData_47 = pts
End Function

Private Function FSData_48() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.660321371944, 30.283373164131, -26.956326530612)
    pts(1) = Array(2.660321371944, 30.283373164131, -23.876326530612)
    pts(2) = Array(2.739080886245, 31.179920395963, -23.636326530612)
    pts(3) = Array(2.739080886245, 31.179920395963, -27.196326530612)
    FSData_48 = pts
End Function

Private Function FSData_49() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(6.574478412196, 29.680569967701, -26.831836734694)
    pts(1) = Array(6.574478412196, 29.680569967701, -23.751836734694)
    pts(2) = Array(6.769117575715, 30.559271052271, -23.511836734694)
    pts(3) = Array(6.769117575715, 30.559271052271, -27.071836734694)
    FSData_49 = pts
End Function

Private Function FSData_50() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(10.377059426774, 28.574055323898, -26.707346938776)
    pts(1) = Array(10.377059426774, 28.574055323898, -23.627346938776)
    pts(2) = Array(10.684275001909, 29.419997751250, -23.387346938776)
    pts(3) = Array(10.684275001909, 29.419997751250, -26.947346938776)
    FSData_50 = pts
End Function

Private Function FSData_51() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(14.003530493699, 26.982607985738, -26.582857142857)
    pts(1) = Array(14.003530493699, 26.982607985738, -23.502857142857)
    pts(2) = Array(14.418108699105, 27.781435195843, -23.262857142857)
    pts(3) = Array(14.418108699105, 27.781435195843, -26.822857142857)
    FSData_51 = pts
End Function

Private Function FSData_52() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(17.392346467714, 24.933236539747, -26.458367346939)
    pts(1) = Array(17.392346467714, 24.933236539747, -23.378367346939)
    pts(2) = Array(17.907251461824, 25.671391568885, -23.138367346939)
    pts(3) = Array(17.907251461824, 25.671391568885, -26.698367346939)
    FSData_52 = pts
End Function

Private Function FSData_53() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.485995468116, 22.460721040971, -26.333877551020)
    pts(1) = Array(20.485995468116, 22.460721040971, -23.253877551020)
    pts(2) = Array(21.092488755001, 23.125676598105, -23.013877551020)
    pts(3) = Array(21.092488755001, 23.125676598105, -26.573877551020)
    FSData_53 = pts
End Function

Private Function FSData_54() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.231974917720, 19.607022757737, -26.209387755102)
    pts(1) = Array(23.231974917720, 19.607022757737, -23.129387755102)
    pts(2) = Array(23.919763648837, 20.187493826223, -22.889387755102)
    pts(3) = Array(23.919763648837, 20.187493826223, -26.449387755102)
    FSData_54 = pts
End Function

Private Function FSData_55() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.583682568506, 16.420572043444, -26.084897959184)
    pts(1) = Array(25.583682568506, 16.420572043444, -23.004897959184)
    pts(2) = Array(26.341094223494, 16.906707399993, -22.764897959184)
    pts(3) = Array(26.341094223494, 16.906707399993, -26.324897959184)
    FSData_55 = pts
End Function

Private Function FSData_56() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.501207392286, 12.955446420964, -25.960408163265)
    pts(1) = Array(27.501207392286, 12.955446420964, -22.880408163265)
    pts(2) = Array(28.315387874295, 13.338995821585, -22.640408163265)
    pts(3) = Array(28.315387874295, 13.338995821585, -26.200408163265)
    FSData_56 = pts
End Function

Private Function FSData_57() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.952006914122, 9.270452828457, -25.835918367347)
    pts(1) = Array(28.952006914122, 9.270452828457, -22.755918367347)
    pts(2) = Array(29.809138697764, 9.544907024036, -22.515918367347)
    pts(3) = Array(29.809138697764, 9.544907024036, -26.075918367347)
    FSData_57 = pts
End Function

Private Function FSData_58() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.911459493398, 5.428129601879, -25.711428571429)
    pts(1) = Array(29.911459493398, 5.428129601879, -22.631428571429)
    pts(2) = Array(30.796996123137, 5.588830807197, -22.391428571429)
    pts(3) = Array(30.796996123137, 5.588830807197, -25.951428571429)
    FSData_58 = pts
End Function

Private Function FSData_59() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.363282179757, 1.493685131630, -25.586938775510)
    pts(1) = Array(30.363282179757, 1.493685131630, -22.506938775510)
    pts(2) = Array(31.262195139026, 1.537906073027, -22.266938775510)
    pts(3) = Array(31.262195139026, 1.537906073027, -25.826938775510)
    FSData_59 = pts
End Function

Private Function FSData_60() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.299807052415, -2.466108794522, -25.462448979592)
    pts(1) = Array(30.299807052415, -2.466108794522, -22.382448979592)
    pts(2) = Array(31.196840813835, -2.539118594360, -22.142448979592)
    pts(3) = Array(31.196840813835, -2.539118594360, -25.702448979592)
    FSData_60 = pts
End Function

Private Function FSData_61() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.722111353093, -6.384050181063, -25.337959183673)
    pts(1) = Array(29.722111353093, -6.384050181063, -22.257959183673)
    pts(2) = Array(30.602042281310, -6.573051666686, -22.017959183673)
    pts(3) = Array(30.602042281310, -6.573051666686, -25.577959183673)
    FSData_61 = pts
End Function

Private Function FSData_62() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.639999204056, -10.193647315444, -25.213469387755)
    pts(1) = Array(28.639999204056, -10.193647315444, -22.133469387755)
    pts(2) = Array(29.487893917334, -10.495432926756, -21.893469387755)
    pts(3) = Array(29.487893917334, -10.495432926756, -25.453469387755)
    FSData_62 = pts
End Function

Private Function FSData_63() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.071835221529, -13.830247204528, -25.088979591837)
    pts(1) = Array(27.071835221529, -13.830247204528, -22.008979591837)
    pts(2) = Array(27.873304027430, -14.239695312557, -21.768979591837)
    pts(3) = Array(27.873304027430, -14.239695312557, -25.328979591837)
    FSData_63 = pts
End Function

Private Function FSData_64() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.044232848254, -17.232132805966, -24.964489795918)
    pts(1) = Array(25.044232848254, -17.232132805966, -21.884489795918)
    pts(2) = Array(25.785673952314, -17.742294632458, -21.644489795918)
    pts(3) = Array(25.785673952314, -17.742294632458, -25.204489795918)
    FSData_64 = pts
End Function

Private Function FSData_65() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.591602694513, -20.341570433309, -24.840000000000)
    pts(1) = Array(22.591602694513, -20.341570433309, -21.760000000000)
    pts(2) = Array(23.260433037443, -20.943787979032, -21.520000000000)
    pts(3) = Array(23.260433037443, -20.943787979032, -25.080000000000)
    FSData_65 = pts
End Function

Private Function FSData_66() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.755568552772, -23.105789559258, -24.715510204082)
    pts(1) = Array(19.755568552772, -23.105789559258, -21.635510204082)
    pts(2) = Array(20.340437358611, -23.789842539631, -21.395510204082)
    pts(3) = Array(20.340437358611, -23.789842539631, -24.955510204082)
    FSData_66 = pts
End Function

Private Function FSData_67() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(16.584260996801, -25.477878388712, -24.591020408163)
    pts(1) = Array(16.584260996801, -25.477878388712, -21.511020408163)
    pts(2) = Array(17.075242407890, -26.232157683115, -21.271020408163)
    pts(3) = Array(17.075242407890, -26.232157683115, -24.831020408163)
    FSData_67 = pts
End Function

Private Function FSData_68() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.131500553658, -27.417580002788, -24.466530612245)
    pts(1) = Array(13.131500553658, -27.417580002788, -21.386530612245)
    pts(2) = Array(13.520262083207, -28.229284673923, -21.146530612245)
    pts(3) = Array(13.520262083207, -28.229284673923, -24.706530612245)
    FSData_68 = pts
End Function

Private Function FSData_69() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(9.455884311006, -28.891975562375, -24.342040816327)
    pts(1) = Array(9.455884311006, -28.891975562375, -21.262040816327)
    pts(2) = Array(9.735828254423, -29.747330102051, -21.022040816327)
    pts(3) = Array(9.735828254423, -29.747330102051, -24.582040816327)
    FSData_69 = pts
End Function

Private Function FSData_70() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.619791461008, -29.876042976518, -24.217551020408)
    pts(1) = Array(5.619791461008, -29.876042976518, -21.137551020408)
    pts(2) = Array(5.786166866104, -30.760531090955, -20.897551020408)
    pts(3) = Array(5.786166866104, -30.760531090955, -24.457551020408)
    FSData_70 = pts
End Function

Private Function FSData_71() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.688324657820, -30.353081554429, -24.093061224490)
    pts(1) = Array(1.688324657820, -30.353081554429, -21.013061224490)
    pts(2) = Array(1.738307953611, -31.251692521501, -20.773061224490)
    pts(3) = Array(1.738307953611, -31.251692521501, -24.333061224490)
    FSData_71 = pts
End Function

Private Function FSData_72() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.271794845027, -30.314995434308, -23.968571428571)
    pts(1) = Array(-2.271794845027, -30.314995434308, -20.888571428571)
    pts(2) = Array(-2.339051929255, -31.212478851771, -20.648571428571)
    pts(3) = Array(-2.339051929255, -31.212478851771, -24.208571428571)
    FSData_72 = pts
End Function

Private Function FSData_73() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.193359526631, -29.762430978902, -23.844081632653)
    pts(1) = Array(-6.193359526631, -29.762430978902, -20.764081632653)
    pts(2) = Array(-6.376715565249, -30.643555580251, -20.524081632653)
    pts(3) = Array(-6.376715565249, -30.643555580251, -24.084081632653)
    FSData_73 = pts
End Function

Private Function FSData_74() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.009816183199, -28.704765806022, -23.719591836735)
    pts(1) = Array(-10.009816183199, -28.704765806022, -20.639591836735)
    pts(2) = Array(-10.306159425464, -29.554577951595, -20.399591836735)
    pts(3) = Array(-10.306159425464, -29.554577951595, -23.959591836735)
    FSData_74 = pts
End Function

Private Function FSData_75() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.656395408064, -27.159949640208, -23.595102040816)
    pts(1) = Array(-13.656395408064, -27.159949640208, -20.515102040816)
    pts(2) = Array(-14.060696587908, -27.964027096661, -20.275102040816)
    pts(3) = Array(-14.060696587908, -27.964027096661, -23.835102040816)
    FSData_75 = pts
End Function

Private Function FSData_76() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.071210798731, -25.154199686440, -23.470612244898)
    pts(1) = Array(-17.071210798731, -25.154199686440, -20.390612244898)
    pts(2) = Array(-17.576608486852, -25.898896387683, -20.150612244898)
    pts(3) = Array(-17.576608486852, -25.898896387683, -23.710612244898)
    FSData_76 = pts
End Function

Private Function FSData_77() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.196309236214, -22.721555695753, -23.346122448980)
    pts(1) = Array(-20.196309236214, -22.721555695753, -20.266122448980)
    pts(2) = Array(-20.794226285971, -23.394233331483, -20.026122448980)
    pts(3) = Array(-20.794226285971, -23.394233331483, -23.586122448980)
    FSData_77 = pts
End Function

Private Function FSData_78() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.978654412286, -19.903302273761, -23.221632653061)
    pts(1) = Array(-22.978654412286, -19.903302273761, -20.141632653061)
    pts(2) = Array(-23.658943523176, -20.492544775287, -19.901632653061)
    pts(3) = Array(-23.658943523176, -20.492544775287, -23.461632653061)
    FSData_78 = pts
End Function

Private Function FSData_79() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.371026913128, -16.747268236144, -23.097142857143)
    pts(1) = Array(-25.371026913128, -16.747268236144, -20.017142857143)
    pts(2) = Array(-26.122142841477, -17.243075519451, -19.777142857143)
    pts(3) = Array(-26.122142841477, -17.243075519451, -23.337142857143)
    FSData_79 = pts
End Function

Private Function FSData_80() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.332825583963, -13.307014901796, -22.972653061224)
    pts(1) = Array(-27.332825583963, -13.307014901796, -19.892653061224)
    pts(2) = Array(-28.142021078226, -13.700972579810, -19.652653061224)
    pts(3) = Array(-28.142021078226, -13.700972579810, -23.212653061224)
    FSData_80 = pts
End Function

Private Function FSData_81() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.830756574675, -9.640927099187, -22.848163265306)
    pts(1) = Array(-28.830756574675, -9.640927099187, -19.768163265306)
    pts(2) = Array(-29.684298710110, -9.926349283044, -19.528163265306)
    pts(3) = Array(-29.684298710110, -9.926349283044, -23.088163265306)
    FSData_81 = pts
End Function

Private Function FSData_82() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.839398372528, -5.811222312522, -22.723673469388)
    pts(1) = Array(-29.839398372528, -5.811222312522, -19.643673469388)
    pts(2) = Array(-30.722801613820, -5.983265078353, -19.403673469388)
    pts(3) = Array(-30.722801613820, -5.983265078353, -22.963673469388)
    FSData_82 = pts
End Function

Private Function FSData_83() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.341633232802, -1.882894783597, -22.599183673469)
    pts(1) = Array(-30.341633232802, -1.882894783597, -19.519183673469)
    pts(2) = Array(-31.239905269300, -1.938638379164, -19.279183673469)
    pts(3) = Array(-31.239905269300, -1.938638379164, -22.839183673469)
    FSData_83 = pts
End Function

Private Function FSData_84() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.328937685474, 2.077387510944, -22.474693877551)
    pts(1) = Array(-30.328937685474, 2.077387510944, -19.394693877551)
    pts(2) = Array(-31.226833866952, 2.138889114887, -19.154693877551)
    pts(3) = Array(-31.226833866952, 2.138889114887, -22.714693877551)
    FSData_84 = pts
End Function

Private Function FSData_85() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.801527187744, 6.002414287447, -22.350204081633)
    pts(1) = Array(-29.801527187744, 6.002414287447, -19.270204081633)
    pts(2) = Array(-30.683809242644, 6.180117342010, -19.030204081633)
    pts(3) = Array(-30.683809242644, 6.180117342010, -22.590204081633)
    FSData_85 = pts
End Function

Private Function FSData_86() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.768352467495, 9.825573586616, -22.225714285714)
    pts(1) = Array(-28.768352467495, 9.825573586616, -19.145714285714)
    pts(2) = Array(-29.620047112914, 10.116462278325, -18.905714285714)
    pts(3) = Array(-29.620047112914, 10.116462278325, -22.465714285714)
    FSData_86 = pts
End Function

Private Function FSData_87() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.246947619734, 13.481982250673, -22.101224489796)
    pts(1) = Array(-27.246947619734, 13.481982250673, -19.021224489796)
    pts(2) = Array(-28.053600674266, 13.881119883094, -18.781224489796)
    pts(3) = Array(-28.053600674266, 13.881119883094, -22.341224489796)
    FSData_87 = pts
End Function

Private Function FSData_88() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.263132533998, 16.909587060885, -21.976734693878)
    pts(1) = Array(-25.263132533998, 16.909587060885, -18.896734693878)
    pts(2) = Array(-26.011054220860, 17.410199835713, -18.656734693878)
    pts(3) = Array(-26.011054220860, 17.410199835713, -22.216734693878)
    FSData_88 = pts
End Function

Private Function FSData_89() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.850574702833, 20.050217847950, -21.852244897959)
    pts(1) = Array(-22.850574702833, 20.050217847950, -18.772244897959)
    pts(2) = Array(-23.527071980219, 20.643809823712, -18.532244897959)
    pts(3) = Array(-23.527071980219, 20.643809823712, -22.092244897959)
    FSData_89 = pts
End Function

Private Function FSData_90() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.050217847950, 22.850574702833, -21.727755102041)
    pts(1) = Array(-20.050217847950, 22.850574702833, -18.647755102041)
    pts(2) = Array(-20.643809823712, 23.527071980219, -18.407755102041)
    pts(3) = Array(-20.643809823712, 23.527071980219, -21.967755102041)
    FSData_90 = pts
End Function

Private Function FSData_91() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.909587060885, 25.263132533998, -21.603265306122)
    pts(1) = Array(-16.909587060885, 25.263132533998, -18.523265306122)
    pts(2) = Array(-17.410199835713, 26.011054220860, -18.283265306122)
    pts(3) = Array(-17.410199835713, 26.011054220860, -21.843265306122)
    FSData_91 = pts
End Function

Private Function FSData_92() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.481982250673, 27.246947619734, -21.478775510204)
    pts(1) = Array(-13.481982250673, 27.246947619734, -18.398775510204)
    pts(2) = Array(-13.881119883094, 28.053600674266, -18.158775510204)
    pts(3) = Array(-13.881119883094, 28.053600674266, -21.718775510204)
    FSData_92 = pts
End Function

Private Function FSData_93() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.825573586616, 28.768352467495, -21.354285714286)
    pts(1) = Array(-9.825573586616, 28.768352467495, -18.274285714286)
    pts(2) = Array(-10.116462278325, 29.620047112914, -18.034285714286)
    pts(3) = Array(-10.116462278325, 29.620047112914, -21.594285714286)
    FSData_93 = pts
End Function

Private Function FSData_94() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.002414287447, 29.801527187744, -21.229795918367)
    pts(1) = Array(-6.002414287447, 29.801527187744, -18.149795918367)
    pts(2) = Array(-6.180117342010, 30.683809242644, -17.909795918367)
    pts(3) = Array(-6.180117342010, 30.683809242644, -21.469795918367)
    FSData_94 = pts
End Function

Private Function FSData_95() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.077387510944, 30.328937685474, -21.105306122449)
    pts(1) = Array(-2.077387510944, 30.328937685474, -18.025306122449)
    pts(2) = Array(-2.138889114887, 31.226833866952, -17.785306122449)
    pts(3) = Array(-2.138889114887, 31.226833866952, -21.345306122449)
    FSData_95 = pts
End Function

Private Function FSData_96() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.882894783597, 30.341633232802, -20.980816326531)
    pts(1) = Array(1.882894783597, 30.341633232802, -17.900816326531)
    pts(2) = Array(1.938638379164, 31.239905269300, -17.660816326531)
    pts(3) = Array(1.938638379164, 31.239905269300, -21.220816326531)
    FSData_96 = pts
End Function

Private Function FSData_97() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.811222312522, 29.839398372528, -20.856326530612)
    pts(1) = Array(5.811222312522, 29.839398372528, -17.776326530612)
    pts(2) = Array(5.983265078353, 30.722801613820, -17.536326530612)
    pts(3) = Array(5.983265078353, 30.722801613820, -21.096326530612)
    FSData_97 = pts
End Function

Private Function FSData_98() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(9.640927099187, 28.830756574675, -20.731836734694)
    pts(1) = Array(9.640927099187, 28.830756574675, -17.651836734694)
    pts(2) = Array(9.926349283044, 29.684298710110, -17.411836734694)
    pts(3) = Array(9.926349283044, 29.684298710110, -20.971836734694)
    FSData_98 = pts
End Function

Private Function FSData_99() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.307014901797, 27.332825583963, -20.607346938776)
    pts(1) = Array(13.307014901797, 27.332825583963, -17.527346938776)
    pts(2) = Array(13.700972579810, 28.142021078225, -17.287346938775)
    pts(3) = Array(13.700972579810, 28.142021078225, -20.847346938776)
    FSData_99 = pts
End Function

Private Function FSData_100() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(16.747268236144, 25.371026913128, -20.482857142857)
    pts(1) = Array(16.747268236144, 25.371026913128, -17.402857142857)
    pts(2) = Array(17.243075519451, 26.122142841477, -17.162857142857)
    pts(3) = Array(17.243075519451, 26.122142841477, -20.722857142857)
    FSData_100 = pts
End Function

Private Function FSData_101() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.903302273761, 22.978654412286, -20.358367346939)
    pts(1) = Array(19.903302273761, 22.978654412286, -17.278367346939)
    pts(2) = Array(20.492544775287, 23.658943523176, -17.038367346939)
    pts(3) = Array(20.492544775287, 23.658943523176, -20.598367346939)
    FSData_101 = pts
End Function

Private Function FSData_102() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.721555695753, 20.196309236214, -20.233877551020)
    pts(1) = Array(22.721555695753, 20.196309236214, -17.153877551020)
    pts(2) = Array(23.394233331483, 20.794226285971, -16.913877551020)
    pts(3) = Array(23.394233331483, 20.794226285971, -20.473877551020)
    FSData_102 = pts
End Function

Private Function FSData_103() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.154199686440, 17.071210798731, -20.109387755102)
    pts(1) = Array(25.154199686440, 17.071210798731, -17.029387755102)
    pts(2) = Array(25.898896387683, 17.576608486852, -16.789387755102)
    pts(3) = Array(25.898896387683, 17.576608486852, -20.349387755102)
    FSData_103 = pts
End Function

Private Function FSData_104() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.159949640208, 13.656395408064, -19.984897959184)
    pts(1) = Array(27.159949640208, 13.656395408064, -16.904897959184)
    pts(2) = Array(27.964027096661, 14.060696587908, -16.664897959184)
    pts(3) = Array(27.964027096661, 14.060696587908, -20.224897959184)
    FSData_104 = pts
End Function

Private Function FSData_105() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.704765806022, 10.009816183199, -19.860408163265)
    pts(1) = Array(28.704765806022, 10.009816183199, -16.780408163265)
    pts(2) = Array(29.554577951595, 10.306159425464, -16.540408163265)
    pts(3) = Array(29.554577951595, 10.306159425464, -20.100408163265)
    FSData_105 = pts
End Function

Private Function FSData_106() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.762430978902, 6.193359526631, -19.735918367347)
    pts(1) = Array(29.762430978902, 6.193359526631, -16.655918367347)
    pts(2) = Array(30.643555580251, 6.376715565249, -16.415918367347)
    pts(3) = Array(30.643555580251, 6.376715565249, -19.975918367347)
    FSData_106 = pts
End Function

Private Function FSData_107() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.314995434308, 2.271794845027, -19.611428571429)
    pts(1) = Array(30.314995434308, 2.271794845027, -16.531428571429)
    pts(2) = Array(31.212478851771, 2.339051929255, -16.291428571429)
    pts(3) = Array(31.212478851771, 2.339051929255, -19.851428571429)
    FSData_107 = pts
End Function

Private Function FSData_108() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.353081554429, -1.688324657820, -19.486938775510)
    pts(1) = Array(30.353081554429, -1.688324657820, -16.406938775510)
    pts(2) = Array(31.251692521501, -1.738307953611, -16.166938775510)
    pts(3) = Array(31.251692521501, -1.738307953611, -19.726938775510)
    FSData_108 = pts
End Function

Private Function FSData_109() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.876042976519, -5.619791461008, -19.362448979592)
    pts(1) = Array(29.876042976519, -5.619791461008, -16.282448979592)
    pts(2) = Array(30.760531090955, -5.786166866104, -16.042448979592)
    pts(3) = Array(30.760531090955, -5.786166866104, -19.602448979592)
    FSData_109 = pts
End Function

Private Function FSData_110() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.891975562375, -9.455884311005, -19.237959183673)
    pts(1) = Array(28.891975562375, -9.455884311005, -16.157959183673)
    pts(2) = Array(29.747330102051, -9.735828254423, -15.917959183673)
    pts(3) = Array(29.747330102051, -9.735828254423, -19.477959183673)
    FSData_110 = pts
End Function

Private Function FSData_111() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.417580002788, -13.131500553658, -19.113469387755)
    pts(1) = Array(27.417580002788, -13.131500553658, -16.033469387755)
    pts(2) = Array(28.229284673923, -13.520262083207, -15.793469387755)
    pts(3) = Array(28.229284673923, -13.520262083207, -19.353469387755)
    FSData_111 = pts
End Function

Private Function FSData_112() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.477878388712, -16.584260996801, -18.988979591837)
    pts(1) = Array(25.477878388712, -16.584260996801, -15.908979591837)
    pts(2) = Array(26.232157683115, -17.075242407890, -15.668979591837)
    pts(3) = Array(26.232157683115, -17.075242407890, -19.228979591837)
    FSData_112 = pts
End Function

Private Function FSData_113() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.105789559258, -19.755568552772, -18.864489795918)
    pts(1) = Array(23.105789559258, -19.755568552772, -15.784489795918)
    pts(2) = Array(23.789842539631, -20.340437358611, -15.544489795918)
    pts(3) = Array(23.789842539631, -20.340437358611, -19.104489795918)
    FSData_113 = pts
End Function

Private Function FSData_114() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.341570433309, -22.591602694513, -18.740000000000)
    pts(1) = Array(20.341570433309, -22.591602694513, -15.660000000000)
    pts(2) = Array(20.943787979032, -23.260433037442, -15.420000000000)
    pts(3) = Array(20.943787979032, -23.260433037442, -18.980000000000)
    FSData_114 = pts
End Function

Private Function FSData_115() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(17.232132805966, -25.044232848254, -18.615510204082)
    pts(1) = Array(17.232132805966, -25.044232848254, -15.535510204082)
    pts(2) = Array(17.742294632458, -25.785673952314, -15.295510204082)
    pts(3) = Array(17.742294632458, -25.785673952314, -18.855510204082)
    FSData_115 = pts
End Function

Private Function FSData_116() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.830247204528, -27.071835221529, -18.491020408163)
    pts(1) = Array(13.830247204528, -27.071835221529, -15.411020408163)
    pts(2) = Array(14.239695312557, -27.873304027430, -15.171020408163)
    pts(3) = Array(14.239695312557, -27.873304027430, -18.731020408163)
    FSData_116 = pts
End Function

Private Function FSData_117() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(10.193647315444, -28.639999204056, -18.366530612245)
    pts(1) = Array(10.193647315444, -28.639999204056, -15.286530612245)
    pts(2) = Array(10.495432926756, -29.487893917334, -15.046530612245)
    pts(3) = Array(10.495432926756, -29.487893917334, -18.606530612245)
    FSData_117 = pts
End Function

Private Function FSData_118() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(6.384050181063, -29.722111353093, -18.242040816327)
    pts(1) = Array(6.384050181063, -29.722111353093, -15.162040816327)
    pts(2) = Array(6.573051666686, -30.602042281310, -14.922040816327)
    pts(3) = Array(6.573051666686, -30.602042281310, -18.482040816327)
    FSData_118 = pts
End Function

Private Function FSData_119() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.466108794522, -30.299807052415, -18.117551020408)
    pts(1) = Array(2.466108794522, -30.299807052415, -15.037551020408)
    pts(2) = Array(2.539118594360, -31.196840813835, -14.797551020408)
    pts(3) = Array(2.539118594360, -31.196840813835, -18.357551020408)
    FSData_119 = pts
End Function

Private Function FSData_120() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.493685131630, -30.363282179757, -17.993061224490)
    pts(1) = Array(-1.493685131630, -30.363282179757, -14.913061224490)
    pts(2) = Array(-1.537906073026, -31.262195139026, -14.673061224490)
    pts(3) = Array(-1.537906073026, -31.262195139026, -18.233061224490)
    FSData_120 = pts
End Function

Private Function FSData_121() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.428129601878, -29.911459493398, -17.868571428571)
    pts(1) = Array(-5.428129601878, -29.911459493398, -14.788571428571)
    pts(2) = Array(-5.588830807197, -30.796996123137, -14.548571428571)
    pts(3) = Array(-5.588830807197, -30.796996123137, -18.108571428571)
    FSData_121 = pts
End Function

Private Function FSData_122() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.270452828457, -28.952006914122, -17.744081632653)
    pts(1) = Array(-9.270452828457, -28.952006914122, -14.664081632653)
    pts(2) = Array(-9.544907024036, -29.809138697764, -14.424081632653)
    pts(3) = Array(-9.544907024036, -29.809138697764, -17.984081632653)
    FSData_122 = pts
End Function

Private Function FSData_123() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.955446420964, -27.501207392286, -17.619591836735)
    pts(1) = Array(-12.955446420964, -27.501207392286, -14.539591836735)
    pts(2) = Array(-13.338995821585, -28.315387874295, -14.299591836735)
    pts(3) = Array(-13.338995821585, -28.315387874295, -17.859591836735)
    FSData_123 = pts
End Function

Private Function FSData_124() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.420572043444, -25.583682568506, -17.495102040816)
    pts(1) = Array(-16.420572043444, -25.583682568506, -14.415102040816)
    pts(2) = Array(-16.906707399993, -26.341094223495, -14.175102040816)
    pts(3) = Array(-16.906707399993, -26.341094223495, -17.735102040816)
    FSData_124 = pts
End Function

Private Function FSData_125() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.607022757737, -23.231974917721, -17.370612244898)
    pts(1) = Array(-19.607022757737, -23.231974917721, -14.290612244898)
    pts(2) = Array(-20.187493826222, -23.919763648837, -14.050612244898)
    pts(3) = Array(-20.187493826222, -23.919763648837, -17.610612244898)
    FSData_125 = pts
End Function

Private Function FSData_126() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.460721040971, -20.485995468116, -17.246122448980)
    pts(1) = Array(-22.460721040971, -20.485995468116, -14.166122448980)
    pts(2) = Array(-23.125676598105, -21.092488755001, -13.926122448980)
    pts(3) = Array(-23.125676598105, -21.092488755001, -17.486122448980)
    FSData_126 = pts
End Function

Private Function FSData_127() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.933236539747, -17.392346467714, -17.121632653061)
    pts(1) = Array(-24.933236539747, -17.392346467714, -14.041632653061)
    pts(2) = Array(-25.671391568885, -17.907251461824, -13.801632653061)
    pts(3) = Array(-25.671391568885, -17.907251461824, -17.361632653061)
    FSData_127 = pts
End Function

Private Function FSData_128() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-26.982607985739, -14.003530493699, -16.997142857143)
    pts(1) = Array(-26.982607985739, -14.003530493699, -13.917142857143)
    pts(2) = Array(-27.781435195843, -14.418108699105, -13.677142857143)
    pts(3) = Array(-27.781435195843, -14.418108699105, -17.237142857143)
    FSData_128 = pts
End Function

Private Function FSData_129() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.574055323898, -10.377059426774, -16.872653061224)
    pts(1) = Array(-28.574055323898, -10.377059426774, -13.792653061224)
    pts(2) = Array(-29.419997751250, -10.684275001909, -13.552653061224)
    pts(3) = Array(-29.419997751250, -10.684275001909, -17.112653061224)
    FSData_129 = pts
End Function

Private Function FSData_130() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.680569967701, -6.574478412196, -16.748163265306)
    pts(1) = Array(-29.680569967701, -6.574478412196, -13.668163265306)
    pts(2) = Array(-30.559271052271, -6.769117575715, -13.428163265306)
    pts(3) = Array(-30.559271052271, -6.769117575715, -16.988163265306)
    FSData_130 = pts
End Function

Private Function FSData_131() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.283373164130, -2.660321371944, -16.623673469388)
    pts(1) = Array(-30.283373164130, -2.660321371944, -13.543673469388)
    pts(2) = Array(-31.179920395963, -2.739080886245, -13.303673469388)
    pts(3) = Array(-31.179920395963, -2.739080886245, -16.863673469388)
    FSData_131 = pts
End Function

Private Function FSData_132() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.372234689480, 1.298984205893, -16.499183673469)
    pts(1) = Array(-30.372234689480, 1.298984205893, -13.419183673469)
    pts(2) = Array(-31.271412690155, 1.337440975147, -13.179183673469)
    pts(3) = Array(-31.271412690155, 1.337440975147, -16.739183673469)
    FSData_132 = pts
End Function

Private Function FSData_133() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.945646467334, 5.236244613602, -16.374693877551)
    pts(1) = Array(-29.945646467334, 5.236244613602, -13.294693877551)
    pts(2) = Array(-30.832195211432, 5.391265013346, -13.054693877551)
    pts(3) = Array(-30.832195211432, 5.391265013346, -16.614693877551)
    FSData_133 = pts
End Function

Private Function FSData_134() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.010848162262, 9.084640273902, -16.250204081633)
    pts(1) = Array(-29.010848162262, 9.084640273902, -13.170204081633)
    pts(2) = Array(-29.869721956539, 9.353593439905, -12.930204081633)
    pts(3) = Array(-29.869721956539, 9.353593439905, -16.490204081633)
    FSData_134 = pts
End Function

Private Function FSData_135() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.583704314863, 12.778859740611, -16.125714285714)
    pts(1) = Array(-27.583704314863, 12.778859740611, -13.045714285714)
    pts(2) = Array(-28.400327139974, 13.157181246089, -12.805714285714)
    pts(3) = Array(-28.400327139974, 13.157181246089, -16.365714285714)
    FSData_135 = pts
End Function

Private Function FSData_136() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.688435103314, 16.256208104685, -16.001224489796)
    pts(1) = Array(-25.688435103314, 16.256208104685, -12.921224489796)
    pts(2) = Array(-26.448947984662, 16.737477423574, -12.681224489796)
    pts(3) = Array(-26.448947984662, 16.737477423574, -16.241224489796)
    FSData_136 = pts
End Function

Private Function FSData_137() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.357205300687, 19.457670994791, -15.876734693878)
    pts(1) = Array(-23.357205300687, 19.457670994791, -12.796734693878)
    pts(2) = Array(-24.048701510247, 20.033720465032, -12.556734693878)
    pts(3) = Array(-24.048701510247, 20.033720465032, -16.116734693878)
    FSData_137 = pts
End Function

Private Function FSData_138() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.629578403888, 22.328916115160, -15.752244897959)
    pts(1) = Array(-20.629578403888, 22.328916115160, -12.672244897959)
    pts(2) = Array(-21.240322501372, 22.989969552780, -12.432244897959)
    pts(3) = Array(-21.240322501372, 22.989969552780, -15.992244897959)
    FSData_138 = pts
End Function

Private Function FSData_139() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.551845198220, 24.821215323544, -15.627755102041)
    pts(1) = Array(-17.551845198220, 24.821215323544, -12.547755102041)
    pts(2) = Array(-18.071472194220, 25.556053935096, -12.307755102041)
    pts(3) = Array(-18.071472194220, 25.556053935096, -15.867755102041)
    FSData_139 = pts
End Function

Private Function FSData_140() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.176238152580, 26.892271600617, -15.503265306122)
    pts(1) = Array(-14.176238152580, 26.892271600617, -12.423265306122)
    pts(2) = Array(-14.595929413676, 27.688424378267, -12.183265306122)
    pts(3) = Array(-14.595929413676, 27.688424378267, -15.743265306122)
    FSData_140 = pts
End Function

Private Function FSData_141() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.560044977836, 28.506936876243, -15.378775510204)
    pts(1) = Array(-10.560044977836, 28.506936876243, -12.298775510204)
    pts(2) = Array(-10.872677888364, 29.350892244289, -12.058775510204)
    pts(3) = Array(-10.872677888364, 29.350892244289, -15.618775510204)
    FSData_141 = pts
End Function

Private Function FSData_142() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.764636392272, 29.637808530327, -15.254285714286)
    pts(1) = Array(-6.764636392272, 29.637808530327, -12.174285714286)
    pts(2) = Array(-6.964905232833, 30.515243651291, -11.934285714286)
    pts(3) = Array(-6.964905232833, 30.515243651291, -15.494285714286)
    FSData_142 = pts
End Function

Private Function FSData_143() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.854424593974, 30.265694444987, -15.129795918367)
    pts(1) = Array(-2.854424593974, 30.265694444987, -12.049795918367)
    pts(2) = Array(-2.938930585243, 31.161718293687, -11.809795918367)
    pts(3) = Array(-2.938930585243, 31.161718293687, -15.369795918367)
    FSData_143 = pts
End Function

Private Function FSData_144() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.104229884004, 30.379938715594, -15.005306122449)
    pts(1) = Array(1.104229884004, 30.379938715594, -11.925306122449)
    pts(2) = Array(1.136920900307, 31.279344795990, -11.685306122449)
    pts(3) = Array(1.136920900307, 31.279344795990, -15.245306122449)
    FSData_144 = pts
End Function

Private Function FSData_145() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.044144383818, 29.978602493032, -14.880816326531)
    pts(1) = Array(5.044144383818, 29.978602493032, -11.800816326531)
    pts(2) = Array(5.193477605707, 30.866126908944, -11.560816326531)
    pts(3) = Array(5.193477605707, 30.866126908944, -15.120816326531)
    FSData_145 = pts
End Function

Private Function FSData_146() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(8.898454285367, 29.068496888061, -14.756326530612)
    pts(1) = Array(8.898454285367, 29.068496888061, -11.676326530612)
    pts(2) = Array(9.161895366184, 29.929077388037, -11.436326530612)
    pts(3) = Array(9.161895366184, 29.929077388037, -14.996326530612)
    FSData_146 = pts
End Function

Private Function FSData_147() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(12.601747771385, 27.665067379394, -14.631836734694)
    pts(1) = Array(12.601747771385, 27.665067379394, -11.551836734694)
    pts(2) = Array(12.974825830406, 28.484098979442, -11.311836734694)
    pts(3) = Array(12.974825830406, 28.484098979442, -14.871836734694)
    FSData_147 = pts
End Function

Private Function FSData_148() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(16.091175936883, 25.792131687169, -14.507346938776)
    pts(1) = Array(16.091175936883, 25.792131687169, -11.427346938776)
    pts(2) = Array(16.567559435015, 26.555714533171, -11.187346938776)
    pts(3) = Array(16.567559435015, 26.555714533171, -14.747346938776)
    FSData_148 = pts
End Function

Private Function FSData_149() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.307519403202, 23.481475560428, -14.382857142857)
    pts(1) = Array(19.307519403202, 23.481475560428, -11.302857142857)
    pts(2) = Array(19.879123596060, 24.176650823730, -11.062857142857)
    pts(3) = Array(19.879123596060, 24.176650823730, -14.622857142857)
    FSData_149 = pts
End Function

Private Function FSData_150() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.196193335063, 20.772313338492, -14.258367346939)
    pts(1) = Array(22.196193335063, 20.772313338492, -11.178367346939)
    pts(2) = Array(22.853317479852, 21.387283141276, -10.938367346939)
    pts(3) = Array(22.853317479852, 21.387283141276, -14.498367346939)
    FSData_150 = pts
End Function

Private Function FSData_151() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.708173804396, 17.710622441116, -14.133877551020)
    pts(1) = Array(24.708173804396, 17.710622441116, -11.053877551020)
    pts(2) = Array(25.439665792026, 18.234950079175, -10.813877551020)
    pts(3) = Array(25.439665792026, 18.234950079175, -14.373877551020)
    FSData_151 = pts
End Function

Private Function FSData_152() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.800829779539, 14.348363081835, -14.009387755102)
    pts(1) = Array(26.800829779539, 14.348363081835, -10.929387755102)
    pts(2) = Array(27.594275398013, 14.773150146757, -10.689387755102)
    pts(3) = Array(27.594275398013, 14.773150146757, -14.249387755102)
    FSData_152 = pts
End Function

Private Function FSData_153() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.438646620066, 10.742596446808, -13.884897959184)
    pts(1) = Array(28.438646620066, 10.742596446808, -10.804897959184)
    pts(2) = Array(29.280580237108, 11.060633841615, -10.564897959184)
    pts(3) = Array(29.280580237108, 11.060633841615, -14.124897959184)
    FSData_153 = pts
End Function

Private Function FSData_154() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.593828798729, 6.954516304641, -13.760408163265)
    pts(1) = Array(29.593828798729, 6.954516304641, -10.680408163265)
    pts(2) = Array(30.469961888165, 7.160406589976, -10.440408163265)
    pts(3) = Array(30.469961888165, 7.160406589976, -14.000408163265)
    FSData_154 = pts
End Function

Private Function FSData_155() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.246771621688, 3.048410481788, -13.635918367347)
    pts(1) = Array(30.246771621688, 3.048410481788, -10.555918367347)
    pts(2) = Array(31.142235255224, 3.138659476315, -10.315918367347)
    pts(3) = Array(31.142235255224, 3.138659476315, -13.875918367347)
    FSData_155 = pts
End Function

Private Function FSData_156() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.386393941419, -0.909430171549, -13.511428571429)
    pts(1) = Array(30.386393941419, -0.909430171549, -10.431428571429)
    pts(2) = Array(31.285991130474, -0.936354091102, -10.191428571429)
    pts(3) = Array(31.285991130474, -0.936354091102, -13.751428571429)
    FSData_156 = pts
End Function

Private Function FSData_157() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.010326215799, -4.851836809015, -13.386938775510)
    pts(1) = Array(30.010326215799, -4.851836809015, -10.306938775510)
    pts(2) = Array(30.898789820872, -4.995476714545, -10.066938775510)
    pts(3) = Array(30.898789820872, -4.995476714545, -13.626938775510)
    FSData_157 = pts
End Function

Private Function FSData_158() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.124950721807, -8.711902516230, -13.262448979592)
    pts(1) = Array(29.124950721807, -8.711902516230, -10.182448979592)
    pts(2) = Array(29.987202552387, -8.969820682829, -9.942448979592)
    pts(3) = Array(29.987202552387, -8.969820682829, -13.502448979592)
    FSData_158 = pts
End Function

Private Function FSData_159() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.745293241361, -12.424117793666, -13.137959183673)
    pts(1) = Array(27.745293241361, -12.424117793666, -10.057959183673)
    pts(2) = Array(28.566699949164, -12.791937070452, -9.817959183673)
    pts(3) = Array(28.566699949164, -12.791937070452, -13.377959183673)
    FSData_159 = pts
End Function

Private Function FSData_160() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.894768057512, -15.925482323863, -13.013469387755)
    pts(1) = Array(25.894768057512, -15.925482323863, -9.933469387755)
    pts(2) = Array(26.661389480267, -16.396960418977, -9.693469387755)
    pts(3) = Array(26.661389480267, -16.396960418977, -13.253469387755)
    FSData_160 = pts
End Function

Private Function FSData_161() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.604780588679, -19.156574155112, -12.888979591837)
    pts(1) = Array(23.604780588679, -19.156574155112, -9.808979591837)
    pts(2) = Array(24.303606329791, -19.723709574178, -9.568979591837)
    pts(3) = Array(24.303606329791, -19.723709574178, -13.128979591837)
    FSData_161 = pts
End Function

Private Function FSData_162() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.914194404654, -22.062558156396, -12.764489795918)
    pts(1) = Array(20.914194404654, -22.062558156396, -9.684489795918)
    pts(2) = Array(21.533364633739, -22.715725996552, -9.444489795918)
    pts(3) = Array(21.533364633739, -22.715725996552, -13.004489795918)
    FSData_162 = pts
End Function

Private Function FSData_163() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(17.868671669691, -24.594116628998, -12.640000000000)
    pts(1) = Array(17.868671669691, -24.594116628998, -9.560000000000)
    pts(2) = Array(18.397678396754, -25.322231923936, -9.320000000000)
    pts(3) = Array(18.397678396754, -25.322231923936, -12.880000000000)
    FSData_163 = pts
End Function

Private Function FSData_164() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(31.150000000000, 0.000000000000, -2.100000000000)
    pts(1) = Array(31.135884021267, 0.215368217163, -2.100000000000)
    pts(2) = Array(31.093777613377, 0.427051424419, -2.100000000000)
    pts(3) = Array(31.024401228644, 0.631427663402, -2.100000000000)
    pts(4) = Array(30.928941916244, 0.825000000000, -2.100000000000)
    pts(5) = Array(30.809033011481, 1.004456357864, -2.100000000000)
    pts(6) = Array(30.666726188958, 1.166726188958, -2.100000000000)
    pts(7) = Array(30.504456357864, 1.309033011481, -2.100000000000)
    pts(8) = Array(30.325000000000, 1.428941916244, -2.100000000000)
    pts(9) = Array(30.131427663402, 1.524401228644, -2.100000000000)
    pts(10) = Array(29.927051424419, 1.593777613377, -2.100000000000)
    pts(11) = Array(29.715368217163, 1.635884021267, -2.100000000000)
    pts(12) = Array(29.500000000000, 1.650000000000, -2.100000000000)
    pts(13) = Array(29.284631782837, 1.635884021267, -2.100000000000)
    pts(14) = Array(29.072948575581, 1.593777613377, -2.100000000000)
    pts(15) = Array(28.868572336598, 1.524401228644, -2.100000000000)
    pts(16) = Array(28.675000000000, 1.428941916244, -2.100000000000)
    pts(17) = Array(28.495543642136, 1.309033011481, -2.100000000000)
    pts(18) = Array(28.333273811042, 1.166726188958, -2.100000000000)
    pts(19) = Array(28.190966988519, 1.004456357864, -2.100000000000)
    pts(20) = Array(28.071058083756, 0.825000000000, -2.100000000000)
    pts(21) = Array(27.975598771356, 0.631427663402, -2.100000000000)
    pts(22) = Array(27.906222386623, 0.427051424419, -2.100000000000)
    pts(23) = Array(27.864115978733, 0.215368217163, -2.100000000000)
    pts(24) = Array(27.850000000000, 0.000000000000, -2.100000000000)
    pts(25) = Array(27.864115978733, -0.215368217163, -2.100000000000)
    pts(26) = Array(27.906222386623, -0.427051424419, -2.100000000000)
    pts(27) = Array(27.975598771356, -0.631427663402, -2.100000000000)
    pts(28) = Array(28.071058083756, -0.825000000000, -2.100000000000)
    pts(29) = Array(28.190966988519, -1.004456357864, -2.100000000000)
    pts(30) = Array(28.333273811042, -1.166726188958, -2.100000000000)
    pts(31) = Array(28.495543642136, -1.309033011481, -2.100000000000)
    pts(32) = Array(28.675000000000, -1.428941916244, -2.100000000000)
    pts(33) = Array(28.868572336598, -1.524401228644, -2.100000000000)
    pts(34) = Array(29.072948575581, -1.593777613377, -2.100000000000)
    pts(35) = Array(29.284631782837, -1.635884021267, -2.100000000000)
    pts(36) = Array(29.500000000000, -1.650000000000, -2.100000000000)
    pts(37) = Array(29.715368217163, -1.635884021267, -2.100000000000)
    pts(38) = Array(29.927051424419, -1.593777613377, -2.100000000000)
    pts(39) = Array(30.131427663402, -1.524401228644, -2.100000000000)
    pts(40) = Array(30.325000000000, -1.428941916244, -2.100000000000)
    pts(41) = Array(30.504456357864, -1.309033011481, -2.100000000000)
    pts(42) = Array(30.666726188958, -1.166726188958, -2.100000000000)
    pts(43) = Array(30.809033011481, -1.004456357864, -2.100000000000)
    pts(44) = Array(30.928941916244, -0.825000000000, -2.100000000000)
    pts(45) = Array(31.024401228644, -0.631427663402, -2.100000000000)
    pts(46) = Array(31.093777613377, -0.427051424419, -2.100000000000)
    pts(47) = Array(31.135884021267, -0.215368217163, -2.100000000000)
    FSData_164 = pts
End Function

Private Function FSData_165() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(31.150000000000, 0.000000000000, 0.100000000000)
    pts(1) = Array(31.135884021267, 0.215368217163, 0.100000000000)
    pts(2) = Array(31.093777613377, 0.427051424419, 0.100000000000)
    pts(3) = Array(31.024401228644, 0.631427663402, 0.100000000000)
    pts(4) = Array(30.928941916244, 0.825000000000, 0.100000000000)
    pts(5) = Array(30.809033011481, 1.004456357864, 0.100000000000)
    pts(6) = Array(30.666726188958, 1.166726188958, 0.100000000000)
    pts(7) = Array(30.504456357864, 1.309033011481, 0.100000000000)
    pts(8) = Array(30.325000000000, 1.428941916244, 0.100000000000)
    pts(9) = Array(30.131427663402, 1.524401228644, 0.100000000000)
    pts(10) = Array(29.927051424419, 1.593777613377, 0.100000000000)
    pts(11) = Array(29.715368217163, 1.635884021267, 0.100000000000)
    pts(12) = Array(29.500000000000, 1.650000000000, 0.100000000000)
    pts(13) = Array(29.284631782837, 1.635884021267, 0.100000000000)
    pts(14) = Array(29.072948575581, 1.593777613377, 0.100000000000)
    pts(15) = Array(28.868572336598, 1.524401228644, 0.100000000000)
    pts(16) = Array(28.675000000000, 1.428941916244, 0.100000000000)
    pts(17) = Array(28.495543642136, 1.309033011481, 0.100000000000)
    pts(18) = Array(28.333273811042, 1.166726188958, 0.100000000000)
    pts(19) = Array(28.190966988519, 1.004456357864, 0.100000000000)
    pts(20) = Array(28.071058083756, 0.825000000000, 0.100000000000)
    pts(21) = Array(27.975598771356, 0.631427663402, 0.100000000000)
    pts(22) = Array(27.906222386623, 0.427051424419, 0.100000000000)
    pts(23) = Array(27.864115978733, 0.215368217163, 0.100000000000)
    pts(24) = Array(27.850000000000, 0.000000000000, 0.100000000000)
    pts(25) = Array(27.864115978733, -0.215368217163, 0.100000000000)
    pts(26) = Array(27.906222386623, -0.427051424419, 0.100000000000)
    pts(27) = Array(27.975598771356, -0.631427663402, 0.100000000000)
    pts(28) = Array(28.071058083756, -0.825000000000, 0.100000000000)
    pts(29) = Array(28.190966988519, -1.004456357864, 0.100000000000)
    pts(30) = Array(28.333273811042, -1.166726188958, 0.100000000000)
    pts(31) = Array(28.495543642136, -1.309033011481, 0.100000000000)
    pts(32) = Array(28.675000000000, -1.428941916244, 0.100000000000)
    pts(33) = Array(28.868572336598, -1.524401228644, 0.100000000000)
    pts(34) = Array(29.072948575581, -1.593777613377, 0.100000000000)
    pts(35) = Array(29.284631782837, -1.635884021267, 0.100000000000)
    pts(36) = Array(29.500000000000, -1.650000000000, 0.100000000000)
    pts(37) = Array(29.715368217163, -1.635884021267, 0.100000000000)
    pts(38) = Array(29.927051424419, -1.593777613377, 0.100000000000)
    pts(39) = Array(30.131427663402, -1.524401228644, 0.100000000000)
    pts(40) = Array(30.325000000000, -1.428941916244, 0.100000000000)
    pts(41) = Array(30.504456357864, -1.309033011481, 0.100000000000)
    pts(42) = Array(30.666726188958, -1.166726188958, 0.100000000000)
    pts(43) = Array(30.809033011481, -1.004456357864, 0.100000000000)
    pts(44) = Array(30.928941916244, -0.825000000000, 0.100000000000)
    pts(45) = Array(31.024401228644, -0.631427663402, 0.100000000000)
    pts(46) = Array(31.093777613377, -0.427051424419, 0.100000000000)
    pts(47) = Array(31.135884021267, -0.215368217163, 0.100000000000)
    FSData_165 = pts
End Function

Private Function FSData_166() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(32.400000000000, 0.000000000000, -1.300000000000)
    pts(1) = Array(32.375190097984, 0.378525957438, -1.300000000000)
    pts(2) = Array(32.301184896238, 0.750575230797, -1.300000000000)
    pts(3) = Array(32.179250644283, 1.109781953859, -1.300000000000)
    pts(4) = Array(32.011473670975, 1.450000000000, -1.300000000000)
    pts(5) = Array(31.800724686845, 1.765408144125, -1.300000000000)
    pts(6) = Array(31.550609665441, 2.050609665441, -1.300000000000)
    pts(7) = Array(31.265408144125, 2.300724686845, -1.300000000000)
    pts(8) = Array(30.950000000000, 2.511473670975, -1.300000000000)
    pts(9) = Array(30.609781953859, 2.679250644283, -1.300000000000)
    pts(10) = Array(30.250575230797, 2.801184896238, -1.300000000000)
    pts(11) = Array(29.878525957438, 2.875190097984, -1.300000000000)
    pts(12) = Array(29.500000000000, 2.900000000000, -1.300000000000)
    pts(13) = Array(29.121474042562, 2.875190097984, -1.300000000000)
    pts(14) = Array(28.749424769203, 2.801184896238, -1.300000000000)
    pts(15) = Array(28.390218046141, 2.679250644283, -1.300000000000)
    pts(16) = Array(28.050000000000, 2.511473670975, -1.300000000000)
    pts(17) = Array(27.734591855875, 2.300724686845, -1.300000000000)
    pts(18) = Array(27.449390334559, 2.050609665441, -1.300000000000)
    pts(19) = Array(27.199275313155, 1.765408144125, -1.300000000000)
    pts(20) = Array(26.988526329025, 1.450000000000, -1.300000000000)
    pts(21) = Array(26.820749355717, 1.109781953859, -1.300000000000)
    pts(22) = Array(26.698815103762, 0.750575230797, -1.300000000000)
    pts(23) = Array(26.624809902016, 0.378525957438, -1.300000000000)
    pts(24) = Array(26.600000000000, 0.000000000000, -1.300000000000)
    pts(25) = Array(26.624809902016, -0.378525957438, -1.300000000000)
    pts(26) = Array(26.698815103762, -0.750575230797, -1.300000000000)
    pts(27) = Array(26.820749355717, -1.109781953859, -1.300000000000)
    pts(28) = Array(26.988526329025, -1.450000000000, -1.300000000000)
    pts(29) = Array(27.199275313155, -1.765408144125, -1.300000000000)
    pts(30) = Array(27.449390334559, -2.050609665441, -1.300000000000)
    pts(31) = Array(27.734591855875, -2.300724686845, -1.300000000000)
    pts(32) = Array(28.050000000000, -2.511473670975, -1.300000000000)
    pts(33) = Array(28.390218046141, -2.679250644283, -1.300000000000)
    pts(34) = Array(28.749424769203, -2.801184896238, -1.300000000000)
    pts(35) = Array(29.121474042562, -2.875190097984, -1.300000000000)
    pts(36) = Array(29.500000000000, -2.900000000000, -1.300000000000)
    pts(37) = Array(29.878525957438, -2.875190097984, -1.300000000000)
    pts(38) = Array(30.250575230797, -2.801184896238, -1.300000000000)
    pts(39) = Array(30.609781953859, -2.679250644283, -1.300000000000)
    pts(40) = Array(30.950000000000, -2.511473670975, -1.300000000000)
    pts(41) = Array(31.265408144125, -2.300724686845, -1.300000000000)
    pts(42) = Array(31.550609665441, -2.050609665441, -1.300000000000)
    pts(43) = Array(31.800724686845, -1.765408144125, -1.300000000000)
    pts(44) = Array(32.011473670975, -1.450000000000, -1.300000000000)
    pts(45) = Array(32.179250644283, -1.109781953859, -1.300000000000)
    pts(46) = Array(32.301184896238, -0.750575230797, -1.300000000000)
    pts(47) = Array(32.375190097984, -0.378525957438, -1.300000000000)
    FSData_166 = pts
End Function

Private Function FSData_167() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(32.400000000000, 0.000000000000, 0.100000000000)
    pts(1) = Array(32.375190097984, 0.378525957438, 0.100000000000)
    pts(2) = Array(32.301184896238, 0.750575230797, 0.100000000000)
    pts(3) = Array(32.179250644283, 1.109781953859, 0.100000000000)
    pts(4) = Array(32.011473670975, 1.450000000000, 0.100000000000)
    pts(5) = Array(31.800724686845, 1.765408144125, 0.100000000000)
    pts(6) = Array(31.550609665441, 2.050609665441, 0.100000000000)
    pts(7) = Array(31.265408144125, 2.300724686845, 0.100000000000)
    pts(8) = Array(30.950000000000, 2.511473670975, 0.100000000000)
    pts(9) = Array(30.609781953859, 2.679250644283, 0.100000000000)
    pts(10) = Array(30.250575230797, 2.801184896238, 0.100000000000)
    pts(11) = Array(29.878525957438, 2.875190097984, 0.100000000000)
    pts(12) = Array(29.500000000000, 2.900000000000, 0.100000000000)
    pts(13) = Array(29.121474042562, 2.875190097984, 0.100000000000)
    pts(14) = Array(28.749424769203, 2.801184896238, 0.100000000000)
    pts(15) = Array(28.390218046141, 2.679250644283, 0.100000000000)
    pts(16) = Array(28.050000000000, 2.511473670975, 0.100000000000)
    pts(17) = Array(27.734591855875, 2.300724686845, 0.100000000000)
    pts(18) = Array(27.449390334559, 2.050609665441, 0.100000000000)
    pts(19) = Array(27.199275313155, 1.765408144125, 0.100000000000)
    pts(20) = Array(26.988526329025, 1.450000000000, 0.100000000000)
    pts(21) = Array(26.820749355717, 1.109781953859, 0.100000000000)
    pts(22) = Array(26.698815103762, 0.750575230797, 0.100000000000)
    pts(23) = Array(26.624809902016, 0.378525957438, 0.100000000000)
    pts(24) = Array(26.600000000000, 0.000000000000, 0.100000000000)
    pts(25) = Array(26.624809902016, -0.378525957438, 0.100000000000)
    pts(26) = Array(26.698815103762, -0.750575230797, 0.100000000000)
    pts(27) = Array(26.820749355717, -1.109781953859, 0.100000000000)
    pts(28) = Array(26.988526329025, -1.450000000000, 0.100000000000)
    pts(29) = Array(27.199275313155, -1.765408144125, 0.100000000000)
    pts(30) = Array(27.449390334559, -2.050609665441, 0.100000000000)
    pts(31) = Array(27.734591855875, -2.300724686845, 0.100000000000)
    pts(32) = Array(28.050000000000, -2.511473670975, 0.100000000000)
    pts(33) = Array(28.390218046141, -2.679250644283, 0.100000000000)
    pts(34) = Array(28.749424769203, -2.801184896238, 0.100000000000)
    pts(35) = Array(29.121474042562, -2.875190097984, 0.100000000000)
    pts(36) = Array(29.500000000000, -2.900000000000, 0.100000000000)
    pts(37) = Array(29.878525957438, -2.875190097984, 0.100000000000)
    pts(38) = Array(30.250575230797, -2.801184896238, 0.100000000000)
    pts(39) = Array(30.609781953859, -2.679250644283, 0.100000000000)
    pts(40) = Array(30.950000000000, -2.511473670975, 0.100000000000)
    pts(41) = Array(31.265408144125, -2.300724686845, 0.100000000000)
    pts(42) = Array(31.550609665441, -2.050609665441, 0.100000000000)
    pts(43) = Array(31.800724686845, -1.765408144125, 0.100000000000)
    pts(44) = Array(32.011473670975, -1.450000000000, 0.100000000000)
    pts(45) = Array(32.179250644283, -1.109781953859, 0.100000000000)
    pts(46) = Array(32.301184896238, -0.750575230797, 0.100000000000)
    pts(47) = Array(32.375190097984, -0.378525957438, 0.100000000000)
    FSData_167 = pts
End Function

Private Function FSData_168() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-13.100000000000, 25.547749411641, -2.100000000000)
    pts(1) = Array(-13.114115978733, 25.763117628804, -2.100000000000)
    pts(2) = Array(-13.156222386623, 25.974800836060, -2.100000000000)
    pts(3) = Array(-13.225598771356, 26.179177075043, -2.100000000000)
    pts(4) = Array(-13.321058083756, 26.372749411641, -2.100000000000)
    pts(5) = Array(-13.440966988519, 26.552205769505, -2.100000000000)
    pts(6) = Array(-13.583273811042, 26.714475600599, -2.100000000000)
    pts(7) = Array(-13.745543642136, 26.856782423121, -2.100000000000)
    pts(8) = Array(-13.925000000000, 26.976691327885, -2.100000000000)
    pts(9) = Array(-14.118572336598, 27.072150640285, -2.100000000000)
    pts(10) = Array(-14.322948575581, 27.141527025018, -2.100000000000)
    pts(11) = Array(-14.534631782837, 27.183633432908, -2.100000000000)
    pts(12) = Array(-14.750000000000, 27.197749411641, -2.100000000000)
    pts(13) = Array(-14.965368217163, 27.183633432908, -2.100000000000)
    pts(14) = Array(-15.177051424419, 27.141527025018, -2.100000000000)
    pts(15) = Array(-15.381427663402, 27.072150640285, -2.100000000000)
    pts(16) = Array(-15.575000000000, 26.976691327885, -2.100000000000)
    pts(17) = Array(-15.754456357864, 26.856782423121, -2.100000000000)
    pts(18) = Array(-15.916726188958, 26.714475600599, -2.100000000000)
    pts(19) = Array(-16.059033011481, 26.552205769505, -2.100000000000)
    pts(20) = Array(-16.178941916244, 26.372749411641, -2.100000000000)
    pts(21) = Array(-16.274401228644, 26.179177075043, -2.100000000000)
    pts(22) = Array(-16.343777613377, 25.974800836060, -2.100000000000)
    pts(23) = Array(-16.385884021267, 25.763117628804, -2.100000000000)
    pts(24) = Array(-16.400000000000, 25.547749411641, -2.100000000000)
    pts(25) = Array(-16.385884021267, 25.332381194478, -2.100000000000)
    pts(26) = Array(-16.343777613377, 25.120697987222, -2.100000000000)
    pts(27) = Array(-16.274401228644, 24.916321748239, -2.100000000000)
    pts(28) = Array(-16.178941916244, 24.722749411641, -2.100000000000)
    pts(29) = Array(-16.059033011481, 24.543293053777, -2.100000000000)
    pts(30) = Array(-15.916726188958, 24.381023222683, -2.100000000000)
    pts(31) = Array(-15.754456357864, 24.238716400160, -2.100000000000)
    pts(32) = Array(-15.575000000000, 24.118807495397, -2.100000000000)
    pts(33) = Array(-15.381427663402, 24.023348182997, -2.100000000000)
    pts(34) = Array(-15.177051424419, 23.953971798264, -2.100000000000)
    pts(35) = Array(-14.965368217163, 23.911865390374, -2.100000000000)
    pts(36) = Array(-14.750000000000, 23.897749411641, -2.100000000000)
    pts(37) = Array(-14.534631782837, 23.911865390374, -2.100000000000)
    pts(38) = Array(-14.322948575581, 23.953971798264, -2.100000000000)
    pts(39) = Array(-14.118572336598, 24.023348182997, -2.100000000000)
    pts(40) = Array(-13.925000000000, 24.118807495397, -2.100000000000)
    pts(41) = Array(-13.745543642136, 24.238716400160, -2.100000000000)
    pts(42) = Array(-13.583273811042, 24.381023222683, -2.100000000000)
    pts(43) = Array(-13.440966988519, 24.543293053777, -2.100000000000)
    pts(44) = Array(-13.321058083756, 24.722749411641, -2.100000000000)
    pts(45) = Array(-13.225598771356, 24.916321748239, -2.100000000000)
    pts(46) = Array(-13.156222386623, 25.120697987222, -2.100000000000)
    pts(47) = Array(-13.114115978733, 25.332381194478, -2.100000000000)
    FSData_168 = pts
End Function

Private Function FSData_169() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-13.100000000000, 25.547749411641, 0.100000000000)
    pts(1) = Array(-13.114115978733, 25.763117628804, 0.100000000000)
    pts(2) = Array(-13.156222386623, 25.974800836060, 0.100000000000)
    pts(3) = Array(-13.225598771356, 26.179177075043, 0.100000000000)
    pts(4) = Array(-13.321058083756, 26.372749411641, 0.100000000000)
    pts(5) = Array(-13.440966988519, 26.552205769505, 0.100000000000)
    pts(6) = Array(-13.583273811042, 26.714475600599, 0.100000000000)
    pts(7) = Array(-13.745543642136, 26.856782423121, 0.100000000000)
    pts(8) = Array(-13.925000000000, 26.976691327885, 0.100000000000)
    pts(9) = Array(-14.118572336598, 27.072150640285, 0.100000000000)
    pts(10) = Array(-14.322948575581, 27.141527025018, 0.100000000000)
    pts(11) = Array(-14.534631782837, 27.183633432908, 0.100000000000)
    pts(12) = Array(-14.750000000000, 27.197749411641, 0.100000000000)
    pts(13) = Array(-14.965368217163, 27.183633432908, 0.100000000000)
    pts(14) = Array(-15.177051424419, 27.141527025018, 0.100000000000)
    pts(15) = Array(-15.381427663402, 27.072150640285, 0.100000000000)
    pts(16) = Array(-15.575000000000, 26.976691327885, 0.100000000000)
    pts(17) = Array(-15.754456357864, 26.856782423121, 0.100000000000)
    pts(18) = Array(-15.916726188958, 26.714475600599, 0.100000000000)
    pts(19) = Array(-16.059033011481, 26.552205769505, 0.100000000000)
    pts(20) = Array(-16.178941916244, 26.372749411641, 0.100000000000)
    pts(21) = Array(-16.274401228644, 26.179177075043, 0.100000000000)
    pts(22) = Array(-16.343777613377, 25.974800836060, 0.100000000000)
    pts(23) = Array(-16.385884021267, 25.763117628804, 0.100000000000)
    pts(24) = Array(-16.400000000000, 25.547749411641, 0.100000000000)
    pts(25) = Array(-16.385884021267, 25.332381194478, 0.100000000000)
    pts(26) = Array(-16.343777613377, 25.120697987222, 0.100000000000)
    pts(27) = Array(-16.274401228644, 24.916321748239, 0.100000000000)
    pts(28) = Array(-16.178941916244, 24.722749411641, 0.100000000000)
    pts(29) = Array(-16.059033011481, 24.543293053777, 0.100000000000)
    pts(30) = Array(-15.916726188958, 24.381023222683, 0.100000000000)
    pts(31) = Array(-15.754456357864, 24.238716400160, 0.100000000000)
    pts(32) = Array(-15.575000000000, 24.118807495397, 0.100000000000)
    pts(33) = Array(-15.381427663402, 24.023348182997, 0.100000000000)
    pts(34) = Array(-15.177051424419, 23.953971798264, 0.100000000000)
    pts(35) = Array(-14.965368217163, 23.911865390374, 0.100000000000)
    pts(36) = Array(-14.750000000000, 23.897749411641, 0.100000000000)
    pts(37) = Array(-14.534631782837, 23.911865390374, 0.100000000000)
    pts(38) = Array(-14.322948575581, 23.953971798264, 0.100000000000)
    pts(39) = Array(-14.118572336598, 24.023348182997, 0.100000000000)
    pts(40) = Array(-13.925000000000, 24.118807495397, 0.100000000000)
    pts(41) = Array(-13.745543642136, 24.238716400160, 0.100000000000)
    pts(42) = Array(-13.583273811042, 24.381023222683, 0.100000000000)
    pts(43) = Array(-13.440966988519, 24.543293053777, 0.100000000000)
    pts(44) = Array(-13.321058083756, 24.722749411641, 0.100000000000)
    pts(45) = Array(-13.225598771356, 24.916321748239, 0.100000000000)
    pts(46) = Array(-13.156222386623, 25.120697987222, 0.100000000000)
    pts(47) = Array(-13.114115978733, 25.332381194478, 0.100000000000)
    FSData_169 = pts
End Function

Private Function FSData_170() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-11.850000000000, 25.547749411641, -1.300000000000)
    pts(1) = Array(-11.874809902016, 25.926275369079, -1.300000000000)
    pts(2) = Array(-11.948815103762, 26.298324642438, -1.300000000000)
    pts(3) = Array(-12.070749355717, 26.657531365500, -1.300000000000)
    pts(4) = Array(-12.238526329025, 26.997749411641, -1.300000000000)
    pts(5) = Array(-12.449275313155, 27.313157555766, -1.300000000000)
    pts(6) = Array(-12.699390334559, 27.598359077082, -1.300000000000)
    pts(7) = Array(-12.984591855875, 27.848474098486, -1.300000000000)
    pts(8) = Array(-13.300000000000, 28.059223082616, -1.300000000000)
    pts(9) = Array(-13.640218046141, 28.227000055924, -1.300000000000)
    pts(10) = Array(-13.999424769203, 28.348934307879, -1.300000000000)
    pts(11) = Array(-14.371474042562, 28.422939509625, -1.300000000000)
    pts(12) = Array(-14.750000000000, 28.447749411641, -1.300000000000)
    pts(13) = Array(-15.128525957438, 28.422939509625, -1.300000000000)
    pts(14) = Array(-15.500575230797, 28.348934307879, -1.300000000000)
    pts(15) = Array(-15.859781953859, 28.227000055924, -1.300000000000)
    pts(16) = Array(-16.200000000000, 28.059223082616, -1.300000000000)
    pts(17) = Array(-16.515408144125, 27.848474098486, -1.300000000000)
    pts(18) = Array(-16.800609665441, 27.598359077082, -1.300000000000)
    pts(19) = Array(-17.050724686845, 27.313157555766, -1.300000000000)
    pts(20) = Array(-17.261473670975, 26.997749411641, -1.300000000000)
    pts(21) = Array(-17.429250644283, 26.657531365500, -1.300000000000)
    pts(22) = Array(-17.551184896238, 26.298324642438, -1.300000000000)
    pts(23) = Array(-17.625190097984, 25.926275369079, -1.300000000000)
    pts(24) = Array(-17.650000000000, 25.547749411641, -1.300000000000)
    pts(25) = Array(-17.625190097984, 25.169223454203, -1.300000000000)
    pts(26) = Array(-17.551184896238, 24.797174180844, -1.300000000000)
    pts(27) = Array(-17.429250644283, 24.437967457782, -1.300000000000)
    pts(28) = Array(-17.261473670975, 24.097749411641, -1.300000000000)
    pts(29) = Array(-17.050724686845, 23.782341267516, -1.300000000000)
    pts(30) = Array(-16.800609665441, 23.497139746200, -1.300000000000)
    pts(31) = Array(-16.515408144125, 23.247024724796, -1.300000000000)
    pts(32) = Array(-16.200000000000, 23.036275740666, -1.300000000000)
    pts(33) = Array(-15.859781953859, 22.868498767358, -1.300000000000)
    pts(34) = Array(-15.500575230797, 22.746564515403, -1.300000000000)
    pts(35) = Array(-15.128525957438, 22.672559313657, -1.300000000000)
    pts(36) = Array(-14.750000000000, 22.647749411641, -1.300000000000)
    pts(37) = Array(-14.371474042562, 22.672559313657, -1.300000000000)
    pts(38) = Array(-13.999424769203, 22.746564515403, -1.300000000000)
    pts(39) = Array(-13.640218046141, 22.868498767358, -1.300000000000)
    pts(40) = Array(-13.300000000000, 23.036275740666, -1.300000000000)
    pts(41) = Array(-12.984591855875, 23.247024724796, -1.300000000000)
    pts(42) = Array(-12.699390334559, 23.497139746200, -1.300000000000)
    pts(43) = Array(-12.449275313155, 23.782341267516, -1.300000000000)
    pts(44) = Array(-12.238526329025, 24.097749411641, -1.300000000000)
    pts(45) = Array(-12.070749355717, 24.437967457782, -1.300000000000)
    pts(46) = Array(-11.948815103762, 24.797174180844, -1.300000000000)
    pts(47) = Array(-11.874809902016, 25.169223454203, -1.300000000000)
    FSData_170 = pts
End Function

Private Function FSData_171() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-11.850000000000, 25.547749411641, 0.100000000000)
    pts(1) = Array(-11.874809902016, 25.926275369079, 0.100000000000)
    pts(2) = Array(-11.948815103762, 26.298324642438, 0.100000000000)
    pts(3) = Array(-12.070749355717, 26.657531365500, 0.100000000000)
    pts(4) = Array(-12.238526329025, 26.997749411641, 0.100000000000)
    pts(5) = Array(-12.449275313155, 27.313157555766, 0.100000000000)
    pts(6) = Array(-12.699390334559, 27.598359077082, 0.100000000000)
    pts(7) = Array(-12.984591855875, 27.848474098486, 0.100000000000)
    pts(8) = Array(-13.300000000000, 28.059223082616, 0.100000000000)
    pts(9) = Array(-13.640218046141, 28.227000055924, 0.100000000000)
    pts(10) = Array(-13.999424769203, 28.348934307879, 0.100000000000)
    pts(11) = Array(-14.371474042562, 28.422939509625, 0.100000000000)
    pts(12) = Array(-14.750000000000, 28.447749411641, 0.100000000000)
    pts(13) = Array(-15.128525957438, 28.422939509625, 0.100000000000)
    pts(14) = Array(-15.500575230797, 28.348934307879, 0.100000000000)
    pts(15) = Array(-15.859781953859, 28.227000055924, 0.100000000000)
    pts(16) = Array(-16.200000000000, 28.059223082616, 0.100000000000)
    pts(17) = Array(-16.515408144125, 27.848474098486, 0.100000000000)
    pts(18) = Array(-16.800609665441, 27.598359077082, 0.100000000000)
    pts(19) = Array(-17.050724686845, 27.313157555766, 0.100000000000)
    pts(20) = Array(-17.261473670975, 26.997749411641, 0.100000000000)
    pts(21) = Array(-17.429250644283, 26.657531365500, 0.100000000000)
    pts(22) = Array(-17.551184896238, 26.298324642438, 0.100000000000)
    pts(23) = Array(-17.625190097984, 25.926275369079, 0.100000000000)
    pts(24) = Array(-17.650000000000, 25.547749411641, 0.100000000000)
    pts(25) = Array(-17.625190097984, 25.169223454203, 0.100000000000)
    pts(26) = Array(-17.551184896238, 24.797174180844, 0.100000000000)
    pts(27) = Array(-17.429250644283, 24.437967457782, 0.100000000000)
    pts(28) = Array(-17.261473670975, 24.097749411641, 0.100000000000)
    pts(29) = Array(-17.050724686845, 23.782341267516, 0.100000000000)
    pts(30) = Array(-16.800609665441, 23.497139746200, 0.100000000000)
    pts(31) = Array(-16.515408144125, 23.247024724796, 0.100000000000)
    pts(32) = Array(-16.200000000000, 23.036275740666, 0.100000000000)
    pts(33) = Array(-15.859781953859, 22.868498767358, 0.100000000000)
    pts(34) = Array(-15.500575230797, 22.746564515403, 0.100000000000)
    pts(35) = Array(-15.128525957438, 22.672559313657, 0.100000000000)
    pts(36) = Array(-14.750000000000, 22.647749411641, 0.100000000000)
    pts(37) = Array(-14.371474042562, 22.672559313657, 0.100000000000)
    pts(38) = Array(-13.999424769203, 22.746564515403, 0.100000000000)
    pts(39) = Array(-13.640218046141, 22.868498767358, 0.100000000000)
    pts(40) = Array(-13.300000000000, 23.036275740666, 0.100000000000)
    pts(41) = Array(-12.984591855875, 23.247024724796, 0.100000000000)
    pts(42) = Array(-12.699390334559, 23.497139746200, 0.100000000000)
    pts(43) = Array(-12.449275313155, 23.782341267516, 0.100000000000)
    pts(44) = Array(-12.238526329025, 24.097749411641, 0.100000000000)
    pts(45) = Array(-12.070749355717, 24.437967457782, 0.100000000000)
    pts(46) = Array(-11.948815103762, 24.797174180844, 0.100000000000)
    pts(47) = Array(-11.874809902016, 25.169223454203, 0.100000000000)
    FSData_171 = pts
End Function

Private Function FSData_172() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-13.100000000000, -25.547749411641, -2.100000000000)
    pts(1) = Array(-13.114115978733, -25.332381194478, -2.100000000000)
    pts(2) = Array(-13.156222386623, -25.120697987222, -2.100000000000)
    pts(3) = Array(-13.225598771356, -24.916321748239, -2.100000000000)
    pts(4) = Array(-13.321058083756, -24.722749411641, -2.100000000000)
    pts(5) = Array(-13.440966988519, -24.543293053777, -2.100000000000)
    pts(6) = Array(-13.583273811042, -24.381023222683, -2.100000000000)
    pts(7) = Array(-13.745543642136, -24.238716400160, -2.100000000000)
    pts(8) = Array(-13.925000000000, -24.118807495397, -2.100000000000)
    pts(9) = Array(-14.118572336598, -24.023348182997, -2.100000000000)
    pts(10) = Array(-14.322948575581, -23.953971798264, -2.100000000000)
    pts(11) = Array(-14.534631782837, -23.911865390374, -2.100000000000)
    pts(12) = Array(-14.750000000000, -23.897749411641, -2.100000000000)
    pts(13) = Array(-14.965368217163, -23.911865390374, -2.100000000000)
    pts(14) = Array(-15.177051424419, -23.953971798264, -2.100000000000)
    pts(15) = Array(-15.381427663402, -24.023348182997, -2.100000000000)
    pts(16) = Array(-15.575000000000, -24.118807495397, -2.100000000000)
    pts(17) = Array(-15.754456357864, -24.238716400160, -2.100000000000)
    pts(18) = Array(-15.916726188958, -24.381023222683, -2.100000000000)
    pts(19) = Array(-16.059033011481, -24.543293053777, -2.100000000000)
    pts(20) = Array(-16.178941916244, -24.722749411641, -2.100000000000)
    pts(21) = Array(-16.274401228644, -24.916321748239, -2.100000000000)
    pts(22) = Array(-16.343777613377, -25.120697987222, -2.100000000000)
    pts(23) = Array(-16.385884021267, -25.332381194478, -2.100000000000)
    pts(24) = Array(-16.400000000000, -25.547749411641, -2.100000000000)
    pts(25) = Array(-16.385884021267, -25.763117628804, -2.100000000000)
    pts(26) = Array(-16.343777613377, -25.974800836060, -2.100000000000)
    pts(27) = Array(-16.274401228644, -26.179177075043, -2.100000000000)
    pts(28) = Array(-16.178941916244, -26.372749411641, -2.100000000000)
    pts(29) = Array(-16.059033011481, -26.552205769505, -2.100000000000)
    pts(30) = Array(-15.916726188958, -26.714475600599, -2.100000000000)
    pts(31) = Array(-15.754456357864, -26.856782423121, -2.100000000000)
    pts(32) = Array(-15.575000000000, -26.976691327885, -2.100000000000)
    pts(33) = Array(-15.381427663402, -27.072150640285, -2.100000000000)
    pts(34) = Array(-15.177051424419, -27.141527025018, -2.100000000000)
    pts(35) = Array(-14.965368217163, -27.183633432908, -2.100000000000)
    pts(36) = Array(-14.750000000000, -27.197749411641, -2.100000000000)
    pts(37) = Array(-14.534631782837, -27.183633432908, -2.100000000000)
    pts(38) = Array(-14.322948575581, -27.141527025018, -2.100000000000)
    pts(39) = Array(-14.118572336598, -27.072150640285, -2.100000000000)
    pts(40) = Array(-13.925000000000, -26.976691327885, -2.100000000000)
    pts(41) = Array(-13.745543642136, -26.856782423121, -2.100000000000)
    pts(42) = Array(-13.583273811042, -26.714475600599, -2.100000000000)
    pts(43) = Array(-13.440966988519, -26.552205769505, -2.100000000000)
    pts(44) = Array(-13.321058083756, -26.372749411641, -2.100000000000)
    pts(45) = Array(-13.225598771356, -26.179177075043, -2.100000000000)
    pts(46) = Array(-13.156222386623, -25.974800836060, -2.100000000000)
    pts(47) = Array(-13.114115978733, -25.763117628804, -2.100000000000)
    FSData_172 = pts
End Function

Private Function FSData_173() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-13.100000000000, -25.547749411641, 0.100000000000)
    pts(1) = Array(-13.114115978733, -25.332381194478, 0.100000000000)
    pts(2) = Array(-13.156222386623, -25.120697987222, 0.100000000000)
    pts(3) = Array(-13.225598771356, -24.916321748239, 0.100000000000)
    pts(4) = Array(-13.321058083756, -24.722749411641, 0.100000000000)
    pts(5) = Array(-13.440966988519, -24.543293053777, 0.100000000000)
    pts(6) = Array(-13.583273811042, -24.381023222683, 0.100000000000)
    pts(7) = Array(-13.745543642136, -24.238716400160, 0.100000000000)
    pts(8) = Array(-13.925000000000, -24.118807495397, 0.100000000000)
    pts(9) = Array(-14.118572336598, -24.023348182997, 0.100000000000)
    pts(10) = Array(-14.322948575581, -23.953971798264, 0.100000000000)
    pts(11) = Array(-14.534631782837, -23.911865390374, 0.100000000000)
    pts(12) = Array(-14.750000000000, -23.897749411641, 0.100000000000)
    pts(13) = Array(-14.965368217163, -23.911865390374, 0.100000000000)
    pts(14) = Array(-15.177051424419, -23.953971798264, 0.100000000000)
    pts(15) = Array(-15.381427663402, -24.023348182997, 0.100000000000)
    pts(16) = Array(-15.575000000000, -24.118807495397, 0.100000000000)
    pts(17) = Array(-15.754456357864, -24.238716400160, 0.100000000000)
    pts(18) = Array(-15.916726188958, -24.381023222683, 0.100000000000)
    pts(19) = Array(-16.059033011481, -24.543293053777, 0.100000000000)
    pts(20) = Array(-16.178941916244, -24.722749411641, 0.100000000000)
    pts(21) = Array(-16.274401228644, -24.916321748239, 0.100000000000)
    pts(22) = Array(-16.343777613377, -25.120697987222, 0.100000000000)
    pts(23) = Array(-16.385884021267, -25.332381194478, 0.100000000000)
    pts(24) = Array(-16.400000000000, -25.547749411641, 0.100000000000)
    pts(25) = Array(-16.385884021267, -25.763117628804, 0.100000000000)
    pts(26) = Array(-16.343777613377, -25.974800836060, 0.100000000000)
    pts(27) = Array(-16.274401228644, -26.179177075043, 0.100000000000)
    pts(28) = Array(-16.178941916244, -26.372749411641, 0.100000000000)
    pts(29) = Array(-16.059033011481, -26.552205769505, 0.100000000000)
    pts(30) = Array(-15.916726188958, -26.714475600599, 0.100000000000)
    pts(31) = Array(-15.754456357864, -26.856782423121, 0.100000000000)
    pts(32) = Array(-15.575000000000, -26.976691327885, 0.100000000000)
    pts(33) = Array(-15.381427663402, -27.072150640285, 0.100000000000)
    pts(34) = Array(-15.177051424419, -27.141527025018, 0.100000000000)
    pts(35) = Array(-14.965368217163, -27.183633432908, 0.100000000000)
    pts(36) = Array(-14.750000000000, -27.197749411641, 0.100000000000)
    pts(37) = Array(-14.534631782837, -27.183633432908, 0.100000000000)
    pts(38) = Array(-14.322948575581, -27.141527025018, 0.100000000000)
    pts(39) = Array(-14.118572336598, -27.072150640285, 0.100000000000)
    pts(40) = Array(-13.925000000000, -26.976691327885, 0.100000000000)
    pts(41) = Array(-13.745543642136, -26.856782423121, 0.100000000000)
    pts(42) = Array(-13.583273811042, -26.714475600599, 0.100000000000)
    pts(43) = Array(-13.440966988519, -26.552205769505, 0.100000000000)
    pts(44) = Array(-13.321058083756, -26.372749411641, 0.100000000000)
    pts(45) = Array(-13.225598771356, -26.179177075043, 0.100000000000)
    pts(46) = Array(-13.156222386623, -25.974800836060, 0.100000000000)
    pts(47) = Array(-13.114115978733, -25.763117628804, 0.100000000000)
    FSData_173 = pts
End Function

Private Function FSData_174() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-11.850000000000, -25.547749411641, -1.300000000000)
    pts(1) = Array(-11.874809902016, -25.169223454203, -1.300000000000)
    pts(2) = Array(-11.948815103762, -24.797174180844, -1.300000000000)
    pts(3) = Array(-12.070749355717, -24.437967457782, -1.300000000000)
    pts(4) = Array(-12.238526329025, -24.097749411641, -1.300000000000)
    pts(5) = Array(-12.449275313155, -23.782341267516, -1.300000000000)
    pts(6) = Array(-12.699390334559, -23.497139746200, -1.300000000000)
    pts(7) = Array(-12.984591855875, -23.247024724796, -1.300000000000)
    pts(8) = Array(-13.300000000000, -23.036275740666, -1.300000000000)
    pts(9) = Array(-13.640218046141, -22.868498767358, -1.300000000000)
    pts(10) = Array(-13.999424769203, -22.746564515403, -1.300000000000)
    pts(11) = Array(-14.371474042562, -22.672559313657, -1.300000000000)
    pts(12) = Array(-14.750000000000, -22.647749411641, -1.300000000000)
    pts(13) = Array(-15.128525957438, -22.672559313657, -1.300000000000)
    pts(14) = Array(-15.500575230797, -22.746564515403, -1.300000000000)
    pts(15) = Array(-15.859781953859, -22.868498767358, -1.300000000000)
    pts(16) = Array(-16.200000000000, -23.036275740666, -1.300000000000)
    pts(17) = Array(-16.515408144125, -23.247024724796, -1.300000000000)
    pts(18) = Array(-16.800609665441, -23.497139746200, -1.300000000000)
    pts(19) = Array(-17.050724686845, -23.782341267516, -1.300000000000)
    pts(20) = Array(-17.261473670975, -24.097749411641, -1.300000000000)
    pts(21) = Array(-17.429250644283, -24.437967457782, -1.300000000000)
    pts(22) = Array(-17.551184896238, -24.797174180844, -1.300000000000)
    pts(23) = Array(-17.625190097984, -25.169223454203, -1.300000000000)
    pts(24) = Array(-17.650000000000, -25.547749411641, -1.300000000000)
    pts(25) = Array(-17.625190097984, -25.926275369079, -1.300000000000)
    pts(26) = Array(-17.551184896238, -26.298324642438, -1.300000000000)
    pts(27) = Array(-17.429250644283, -26.657531365500, -1.300000000000)
    pts(28) = Array(-17.261473670975, -26.997749411641, -1.300000000000)
    pts(29) = Array(-17.050724686845, -27.313157555766, -1.300000000000)
    pts(30) = Array(-16.800609665441, -27.598359077082, -1.300000000000)
    pts(31) = Array(-16.515408144125, -27.848474098486, -1.300000000000)
    pts(32) = Array(-16.200000000000, -28.059223082616, -1.300000000000)
    pts(33) = Array(-15.859781953859, -28.227000055924, -1.300000000000)
    pts(34) = Array(-15.500575230797, -28.348934307879, -1.300000000000)
    pts(35) = Array(-15.128525957438, -28.422939509625, -1.300000000000)
    pts(36) = Array(-14.750000000000, -28.447749411641, -1.300000000000)
    pts(37) = Array(-14.371474042562, -28.422939509625, -1.300000000000)
    pts(38) = Array(-13.999424769203, -28.348934307879, -1.300000000000)
    pts(39) = Array(-13.640218046141, -28.227000055924, -1.300000000000)
    pts(40) = Array(-13.300000000000, -28.059223082616, -1.300000000000)
    pts(41) = Array(-12.984591855875, -27.848474098486, -1.300000000000)
    pts(42) = Array(-12.699390334559, -27.598359077082, -1.300000000000)
    pts(43) = Array(-12.449275313155, -27.313157555766, -1.300000000000)
    pts(44) = Array(-12.238526329025, -26.997749411641, -1.300000000000)
    pts(45) = Array(-12.070749355717, -26.657531365500, -1.300000000000)
    pts(46) = Array(-11.948815103762, -26.298324642438, -1.300000000000)
    pts(47) = Array(-11.874809902016, -25.926275369079, -1.300000000000)
    FSData_174 = pts
End Function

Private Function FSData_175() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-11.850000000000, -25.547749411641, 0.100000000000)
    pts(1) = Array(-11.874809902016, -25.169223454203, 0.100000000000)
    pts(2) = Array(-11.948815103762, -24.797174180844, 0.100000000000)
    pts(3) = Array(-12.070749355717, -24.437967457782, 0.100000000000)
    pts(4) = Array(-12.238526329025, -24.097749411641, 0.100000000000)
    pts(5) = Array(-12.449275313155, -23.782341267516, 0.100000000000)
    pts(6) = Array(-12.699390334559, -23.497139746200, 0.100000000000)
    pts(7) = Array(-12.984591855875, -23.247024724796, 0.100000000000)
    pts(8) = Array(-13.300000000000, -23.036275740666, 0.100000000000)
    pts(9) = Array(-13.640218046141, -22.868498767358, 0.100000000000)
    pts(10) = Array(-13.999424769203, -22.746564515403, 0.100000000000)
    pts(11) = Array(-14.371474042562, -22.672559313657, 0.100000000000)
    pts(12) = Array(-14.750000000000, -22.647749411641, 0.100000000000)
    pts(13) = Array(-15.128525957438, -22.672559313657, 0.100000000000)
    pts(14) = Array(-15.500575230797, -22.746564515403, 0.100000000000)
    pts(15) = Array(-15.859781953859, -22.868498767358, 0.100000000000)
    pts(16) = Array(-16.200000000000, -23.036275740666, 0.100000000000)
    pts(17) = Array(-16.515408144125, -23.247024724796, 0.100000000000)
    pts(18) = Array(-16.800609665441, -23.497139746200, 0.100000000000)
    pts(19) = Array(-17.050724686845, -23.782341267516, 0.100000000000)
    pts(20) = Array(-17.261473670975, -24.097749411641, 0.100000000000)
    pts(21) = Array(-17.429250644283, -24.437967457782, 0.100000000000)
    pts(22) = Array(-17.551184896238, -24.797174180844, 0.100000000000)
    pts(23) = Array(-17.625190097984, -25.169223454203, 0.100000000000)
    pts(24) = Array(-17.650000000000, -25.547749411641, 0.100000000000)
    pts(25) = Array(-17.625190097984, -25.926275369079, 0.100000000000)
    pts(26) = Array(-17.551184896238, -26.298324642438, 0.100000000000)
    pts(27) = Array(-17.429250644283, -26.657531365500, 0.100000000000)
    pts(28) = Array(-17.261473670975, -26.997749411641, 0.100000000000)
    pts(29) = Array(-17.050724686845, -27.313157555766, 0.100000000000)
    pts(30) = Array(-16.800609665441, -27.598359077082, 0.100000000000)
    pts(31) = Array(-16.515408144125, -27.848474098486, 0.100000000000)
    pts(32) = Array(-16.200000000000, -28.059223082616, 0.100000000000)
    pts(33) = Array(-15.859781953859, -28.227000055924, 0.100000000000)
    pts(34) = Array(-15.500575230797, -28.348934307879, 0.100000000000)
    pts(35) = Array(-15.128525957438, -28.422939509625, 0.100000000000)
    pts(36) = Array(-14.750000000000, -28.447749411641, 0.100000000000)
    pts(37) = Array(-14.371474042562, -28.422939509625, 0.100000000000)
    pts(38) = Array(-13.999424769203, -28.348934307879, 0.100000000000)
    pts(39) = Array(-13.640218046141, -28.227000055924, 0.100000000000)
    pts(40) = Array(-13.300000000000, -28.059223082616, 0.100000000000)
    pts(41) = Array(-12.984591855875, -27.848474098486, 0.100000000000)
    pts(42) = Array(-12.699390334559, -27.598359077082, 0.100000000000)
    pts(43) = Array(-12.449275313155, -27.313157555766, 0.100000000000)
    pts(44) = Array(-12.238526329025, -26.997749411641, 0.100000000000)
    pts(45) = Array(-12.070749355717, -26.657531365500, 0.100000000000)
    pts(46) = Array(-11.948815103762, -26.298324642438, 0.100000000000)
    pts(47) = Array(-11.874809902016, -25.926275369079, 0.100000000000)
    FSData_175 = pts
End Function

Private Function FSData_176() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-1.275000000000, 24.900000000000, -42.800000000000)
    pts(1) = Array(-1.264092198252, 24.900000000000, -42.633579104919)
    pts(2) = Array(-1.231555428519, 24.900000000000, -42.470005717494)
    pts(3) = Array(-1.177946403952, 24.900000000000, -42.312078623735)
    pts(4) = Array(-1.104182389825, 24.900000000000, -42.162500000000)
    pts(5) = Array(-1.011525508871, 24.900000000000, -42.023829178014)
    pts(6) = Array(-0.901561146013, 24.900000000000, -41.898438853987)
    pts(7) = Array(-0.776170821986, 24.900000000000, -41.788474491129)
    pts(8) = Array(-0.637500000000, 24.900000000000, -41.695817610175)
    pts(9) = Array(-0.487921376265, 24.900000000000, -41.622053596048)
    pts(10) = Array(-0.329994282506, 24.900000000000, -41.568444571481)
    pts(11) = Array(-0.166420895081, 24.900000000000, -41.535907801748)
    pts(12) = Array(0.000000000000, 24.900000000000, -41.525000000000)
    pts(13) = Array(0.166420895081, 24.900000000000, -41.535907801748)
    pts(14) = Array(0.329994282506, 24.900000000000, -41.568444571481)
    pts(15) = Array(0.487921376265, 24.900000000000, -41.622053596048)
    pts(16) = Array(0.637500000000, 24.900000000000, -41.695817610175)
    pts(17) = Array(0.776170821986, 24.900000000000, -41.788474491129)
    pts(18) = Array(0.901561146013, 24.900000000000, -41.898438853987)
    pts(19) = Array(1.011525508871, 24.900000000000, -42.023829178014)
    pts(20) = Array(1.104182389825, 24.900000000000, -42.162500000000)
    pts(21) = Array(1.177946403952, 24.900000000000, -42.312078623735)
    pts(22) = Array(1.231555428519, 24.900000000000, -42.470005717494)
    pts(23) = Array(1.264092198252, 24.900000000000, -42.633579104919)
    pts(24) = Array(1.275000000000, 24.900000000000, -42.800000000000)
    pts(25) = Array(1.264092198252, 24.900000000000, -42.966420895081)
    pts(26) = Array(1.231555428519, 24.900000000000, -43.129994282506)
    pts(27) = Array(1.177946403952, 24.900000000000, -43.287921376265)
    pts(28) = Array(1.104182389825, 24.900000000000, -43.437500000000)
    pts(29) = Array(1.011525508871, 24.900000000000, -43.576170821986)
    pts(30) = Array(0.901561146013, 24.900000000000, -43.701561146013)
    pts(31) = Array(0.776170821986, 24.900000000000, -43.811525508871)
    pts(32) = Array(0.637500000000, 24.900000000000, -43.904182389825)
    pts(33) = Array(0.487921376265, 24.900000000000, -43.977946403952)
    pts(34) = Array(0.329994282506, 24.900000000000, -44.031555428519)
    pts(35) = Array(0.166420895081, 24.900000000000, -44.064092198252)
    pts(36) = Array(0.000000000000, 24.900000000000, -44.075000000000)
    pts(37) = Array(-0.166420895081, 24.900000000000, -44.064092198252)
    pts(38) = Array(-0.329994282506, 24.900000000000, -44.031555428519)
    pts(39) = Array(-0.487921376265, 24.900000000000, -43.977946403952)
    pts(40) = Array(-0.637500000000, 24.900000000000, -43.904182389825)
    pts(41) = Array(-0.776170821986, 24.900000000000, -43.811525508871)
    pts(42) = Array(-0.901561146013, 24.900000000000, -43.701561146013)
    pts(43) = Array(-1.011525508871, 24.900000000000, -43.576170821986)
    pts(44) = Array(-1.104182389825, 24.900000000000, -43.437500000000)
    pts(45) = Array(-1.177946403952, 24.900000000000, -43.287921376265)
    pts(46) = Array(-1.231555428519, 24.900000000000, -43.129994282506)
    pts(47) = Array(-1.264092198252, 24.900000000000, -42.966420895081)
    FSData_176 = pts
End Function

Private Function FSData_177() As Variant
    Dim pts(47) As Variant
    pts(0) = Array(-1.275000000000, 29.200000000000, -42.800000000000)
    pts(1) = Array(-1.264092198252, 29.200000000000, -42.633579104919)
    pts(2) = Array(-1.231555428519, 29.200000000000, -42.470005717494)
    pts(3) = Array(-1.177946403952, 29.200000000000, -42.312078623735)
    pts(4) = Array(-1.104182389825, 29.200000000000, -42.162500000000)
    pts(5) = Array(-1.011525508871, 29.200000000000, -42.023829178014)
    pts(6) = Array(-0.901561146013, 29.200000000000, -41.898438853987)
    pts(7) = Array(-0.776170821986, 29.200000000000, -41.788474491129)
    pts(8) = Array(-0.637500000000, 29.200000000000, -41.695817610175)
    pts(9) = Array(-0.487921376265, 29.200000000000, -41.622053596048)
    pts(10) = Array(-0.329994282506, 29.200000000000, -41.568444571481)
    pts(11) = Array(-0.166420895081, 29.200000000000, -41.535907801748)
    pts(12) = Array(0.000000000000, 29.200000000000, -41.525000000000)
    pts(13) = Array(0.166420895081, 29.200000000000, -41.535907801748)
    pts(14) = Array(0.329994282506, 29.200000000000, -41.568444571481)
    pts(15) = Array(0.487921376265, 29.200000000000, -41.622053596048)
    pts(16) = Array(0.637500000000, 29.200000000000, -41.695817610175)
    pts(17) = Array(0.776170821986, 29.200000000000, -41.788474491129)
    pts(18) = Array(0.901561146013, 29.200000000000, -41.898438853987)
    pts(19) = Array(1.011525508871, 29.200000000000, -42.023829178014)
    pts(20) = Array(1.104182389825, 29.200000000000, -42.162500000000)
    pts(21) = Array(1.177946403952, 29.200000000000, -42.312078623735)
    pts(22) = Array(1.231555428519, 29.200000000000, -42.470005717494)
    pts(23) = Array(1.264092198252, 29.200000000000, -42.633579104919)
    pts(24) = Array(1.275000000000, 29.200000000000, -42.800000000000)
    pts(25) = Array(1.264092198252, 29.200000000000, -42.966420895081)
    pts(26) = Array(1.231555428519, 29.200000000000, -43.129994282506)
    pts(27) = Array(1.177946403952, 29.200000000000, -43.287921376265)
    pts(28) = Array(1.104182389825, 29.200000000000, -43.437500000000)
    pts(29) = Array(1.011525508871, 29.200000000000, -43.576170821986)
    pts(30) = Array(0.901561146013, 29.200000000000, -43.701561146013)
    pts(31) = Array(0.776170821986, 29.200000000000, -43.811525508871)
    pts(32) = Array(0.637500000000, 29.200000000000, -43.904182389825)
    pts(33) = Array(0.487921376265, 29.200000000000, -43.977946403952)
    pts(34) = Array(0.329994282506, 29.200000000000, -44.031555428519)
    pts(35) = Array(0.166420895081, 29.200000000000, -44.064092198252)
    pts(36) = Array(0.000000000000, 29.200000000000, -44.075000000000)
    pts(37) = Array(-0.166420895081, 29.200000000000, -44.064092198252)
    pts(38) = Array(-0.329994282506, 29.200000000000, -44.031555428519)
    pts(39) = Array(-0.487921376265, 29.200000000000, -43.977946403952)
    pts(40) = Array(-0.637500000000, 29.200000000000, -43.904182389825)
    pts(41) = Array(-0.776170821986, 29.200000000000, -43.811525508871)
    pts(42) = Array(-0.901561146013, 29.200000000000, -43.701561146013)
    pts(43) = Array(-1.011525508871, 29.200000000000, -43.576170821986)
    pts(44) = Array(-1.104182389825, 29.200000000000, -43.437500000000)
    pts(45) = Array(-1.177946403952, 29.200000000000, -43.287921376265)
    pts(46) = Array(-1.231555428519, 29.200000000000, -43.129994282506)
    pts(47) = Array(-1.264092198252, 29.200000000000, -42.966420895081)
    FSData_177 = pts
End Function

Private Function FSData_178() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.800000000000, -1.500000000000, -44.050000000000)
    pts(1) = Array(-29.100000000000, -1.500000000000, -44.050000000000)
    pts(2) = Array(-29.100000000000, 1.500000000000, -44.050000000000)
    pts(3) = Array(-27.800000000000, 1.500000000000, -44.050000000000)
    FSData_178 = pts
End Function

Private Function FSData_179() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.800000000000, -1.500000000000, -42.800000000000)
    pts(1) = Array(-29.100000000000, -1.500000000000, -42.800000000000)
    pts(2) = Array(-29.100000000000, 1.500000000000, -42.800000000000)
    pts(3) = Array(-27.800000000000, 1.500000000000, -42.800000000000)
    FSData_179 = pts
End Function

Private Function FSData_180() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.800000000000, -1.500000000000, -44.050000000000)
    pts(1) = Array(29.100000000000, -1.500000000000, -44.050000000000)
    pts(2) = Array(29.100000000000, 1.500000000000, -44.050000000000)
    pts(3) = Array(27.800000000000, 1.500000000000, -44.050000000000)
    FSData_180 = pts
End Function

Private Function FSData_181() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.800000000000, -1.500000000000, -42.800000000000)
    pts(1) = Array(29.100000000000, -1.500000000000, -42.800000000000)
    pts(2) = Array(29.100000000000, 1.500000000000, -42.800000000000)
    pts(3) = Array(27.800000000000, 1.500000000000, -42.800000000000)
    FSData_181 = pts
End Function

Private Function FSData_182() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.200000000000, -1.500000000000, -41.750444021751)
    pts(1) = Array(-23.100000000000, -1.500000000000, -41.750444021751)
    pts(2) = Array(-23.100000000000, 1.500000000000, -41.750444021751)
    pts(3) = Array(-22.200000000000, 1.500000000000, -41.750444021751)
    FSData_182 = pts
End Function

Private Function FSData_183() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.200000000000, -1.500000000000, -40.500444021751)
    pts(1) = Array(-23.100000000000, -1.500000000000, -40.500444021751)
    pts(2) = Array(-23.100000000000, 1.500000000000, -40.500444021751)
    pts(3) = Array(-22.200000000000, 1.500000000000, -40.500444021751)
    FSData_183 = pts
End Function

Private Function FSData_184() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.200000000000, -1.500000000000, -41.750444021751)
    pts(1) = Array(23.100000000000, -1.500000000000, -41.750444021751)
    pts(2) = Array(23.100000000000, 1.500000000000, -41.750444021751)
    pts(3) = Array(22.200000000000, 1.500000000000, -41.750444021751)
    FSData_184 = pts
End Function

Private Function FSData_185() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.200000000000, -1.500000000000, -40.500444021751)
    pts(1) = Array(23.100000000000, -1.500000000000, -40.500444021751)
    pts(2) = Array(23.100000000000, 1.500000000000, -40.500444021751)
    pts(3) = Array(22.200000000000, 1.500000000000, -40.500444021751)
    FSData_185 = pts
End Function

Private Function FSData_186() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.200000000000, -1.500000000000, 10.089581280691)
    pts(1) = Array(-24.100000000000, -1.500000000000, 10.089581280691)
    pts(2) = Array(-24.100000000000, 1.500000000000, 10.089581280691)
    pts(3) = Array(-23.200000000000, 1.500000000000, 10.089581280691)
    FSData_186 = pts
End Function

Private Function FSData_187() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.200000000000, -1.500000000000, 11.339581280691)
    pts(1) = Array(-24.100000000000, -1.500000000000, 11.339581280691)
    pts(2) = Array(-24.100000000000, 1.500000000000, 11.339581280691)
    pts(3) = Array(-23.200000000000, 1.500000000000, 11.339581280691)
    FSData_187 = pts
End Function

Private Function FSData_188() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.200000000000, -1.500000000000, 10.089581280691)
    pts(1) = Array(24.100000000000, -1.500000000000, 10.089581280691)
    pts(2) = Array(24.100000000000, 1.500000000000, 10.089581280691)
    pts(3) = Array(23.200000000000, 1.500000000000, 10.089581280691)
    FSData_188 = pts
End Function

Private Function FSData_189() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.200000000000, -1.500000000000, 11.339581280691)
    pts(1) = Array(24.100000000000, -1.500000000000, 11.339581280691)
    pts(2) = Array(24.100000000000, 1.500000000000, 11.339581280691)
    pts(3) = Array(23.200000000000, 1.500000000000, 11.339581280691)
    FSData_189 = pts
End Function

Private Function FSData_190() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-3.156759590683, -30.034561240122, -38.986513257037)
    pts(1) = Array(-3.156759590683, -30.034561240122, -35.813486742963)
    pts(2) = Array(-3.252925776889, -30.949521383861, -36.060000000000)
    pts(3) = Array(-3.252925776889, -30.949521383861, -38.740000000000)
    FSData_190 = pts
End Function

Private Function FSData_191() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-7.041601467507, -29.367598621147, -38.861787982311)
    pts(1) = Array(-7.041601467507, -29.367598621147, -35.688761468238)
    pts(2) = Array(-7.256113830094, -30.262240698348, -35.935274725275)
    pts(3) = Array(-7.256113830094, -30.262240698348, -38.615274725275)
    FSData_191 = pts
End Function

Private Function FSData_192() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.806487855581, -28.200351420278, -38.737062707586)
    pts(1) = Array(-10.806487855581, -28.200351420278, -35.564036193513)
    pts(2) = Array(-11.135692121380, -29.059434973479, -35.810549450549)
    pts(3) = Array(-11.135692121380, -29.059434973479, -38.490549450549)
    FSData_192 = pts
End Function

Private Function FSData_193() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.387282948744, -26.552703993243, -38.612337432861)
    pts(1) = Array(-14.387282948744, -26.552703993243, -35.439310918788)
    pts(2) = Array(-14.825571038573, -27.361594313567, -35.685824175824)
    pts(3) = Array(-14.825571038573, -27.361594313567, -38.365824175824)
    FSData_193 = pts
End Function

Private Function FSData_194() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.722986983307, -24.452724436952, -38.487612158136)
    pts(1) = Array(-17.722986983307, -24.452724436952, -35.314585644062)
    pts(2) = Array(-18.262892547037, -25.197641870131, -35.561098901099)
    pts(3) = Array(-18.262892547037, -25.197641870131, -38.241098901099)
    FSData_194 = pts
End Function

Private Function FSData_195() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.756775384684, -21.936186442264, -38.362886883410)
    pts(1) = Array(-20.756775384684, -21.936186442264, -35.189860369337)
    pts(2) = Array(-21.389100992429, -22.604441128585, -35.436373626374)
    pts(3) = Array(-21.389100992429, -22.604441128585, -38.116373626374)
    FSData_195 = pts
End Function

Private Function FSData_196() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.436966788609, -19.045959879975, -38.238161608685)
    pts(1) = Array(-23.436966788609, -19.045959879975, -35.065135094612)
    pts(2) = Array(-24.150940611309, -19.626167929299, -35.311648351648)
    pts(3) = Array(-24.150940611309, -19.626167929299, -37.991648351648)
    FSData_196 = pts
End Function

Private Function FSData_197() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.717903446447, -15.831280501565, -38.113436333960)
    pts(1) = Array(-25.717903446447, -15.831280501565, -34.940409819886)
    pts(2) = Array(-26.501362756736, -16.313557920818, -35.186923076923)
    pts(3) = Array(-26.501362756736, -16.313557920818, -37.866923076923)
    FSData_197 = pts
End Function

Private Function FSData_198() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.560729016657, -12.346911195534, -37.988711059234)
    pts(1) = Array(-27.560729016657, -12.346911195534, -34.815684545161)
    pts(2) = Array(-28.400327384052, -12.723042265066, -35.062197802198)
    pts(3) = Array(-28.400327384052, -12.723042265066, -37.742197802198)
    FSData_198 = pts
End Function

Private Function FSData_199() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.934050492554, -8.652209087531, -37.863985784509)
    pts(1) = Array(-28.934050492554, -8.652209087531, -34.690959270436)
    pts(2) = Array(-29.815485143320, -8.915786318012, -34.937472527473)
    pts(3) = Array(-29.815485143320, -8.915786318012, -37.617472527473)
    FSData_199 = pts
End Function

Private Function FSData_200() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.814472990229, -4.810114376488, -37.739260509784)
    pts(1) = Array(-29.814472990229, -4.810114376488, -34.566233995711)
    pts(2) = Array(-30.722728458806, -4.956647662129, -34.812747252747)
    pts(3) = Array(-30.722728458806, -4.956647662129, -37.492747252747)
    FSData_200 = pts
End Function

Private Function FSData_201() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.186998286410, -0.886078132167, -37.614535235059)
    pts(1) = Array(-30.186998286410, -0.886078132167, -34.441508720985)
    pts(2) = Array(-31.106602207717, -0.913071240829, -34.688021978022)
    pts(3) = Array(-31.106602207717, -0.913071240829, -37.368021978022)
    FSData_201 = pts
End Function

Private Function FSData_202() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.045280317038, 3.053052680613, -37.489809960333)
    pts(1) = Array(-30.045280317038, 3.053052680613, -34.316783446260)
    pts(2) = Array(-30.960567002193, 3.146059583466, -34.563296703297)
    pts(3) = Array(-30.960567002193, 3.146059583466, -37.243296703297)
    FSData_202 = pts
End Function

Private Function FSData_203() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.391733284106, 6.940173957186, -37.365084685608)
    pts(1) = Array(-29.391733284106, 6.940173957186, -34.192058171535)
    pts(2) = Array(-30.287110589449, 7.151596475087, -34.438571428571)
    pts(3) = Array(-30.287110589449, 7.151596475087, -37.118571428571)
    FSData_203 = pts
End Function

Private Function FSData_204() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.237490529100, 10.709067588685, -37.240359410883)
    pts(1) = Array(-28.237490529100, 10.709067588685, -34.067332896810)
    pts(2) = Array(-29.097705472370, 11.035304084764, -34.313846153846)
    pts(3) = Array(-29.097705472370, 11.035304084764, -36.993846153846)
    FSData_204 = pts
End Function

Private Function FSData_205() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-26.602214873674, 14.295529504530, -37.115634136158)
    pts(1) = Array(-26.602214873674, 14.295529504530, -33.942607622084)
    pts(2) = Array(-27.412613472474, 14.731022456324, -34.189120879121)
    pts(3) = Array(-27.412613472474, 14.731022456324, -36.869120879121)
    FSData_205 = pts
End Function

Private Function FSData_206() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.513763658434, 17.638463405253, -36.990908861432)
    pts(1) = Array(-24.513763658434, 17.638463405253, -33.817882347359)
    pts(2) = Array(-25.260540564585, 18.175794078525, -34.064395604396)
    pts(3) = Array(-25.260540564585, 18.175794078525, -36.744395604396)
    FSData_206 = pts
End Function

Private Function FSData_207() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.007714185988, 20.680921553641, -36.866183586707)
    pts(1) = Array(-22.007714185988, 20.680921553641, -33.693157072634)
    pts(2) = Array(-22.678147863178, 21.310936382427, -33.939670329670)
    pts(3) = Array(-22.678147863178, 21.310936382427, -36.619670329670)
    FSData_207 = pts
End Function

Private Function FSData_208() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.126757652464, 23.371074894063, -36.741458311982)
    pts(1) = Array(-19.126757652464, 23.371074894063, -33.568431797908)
    pts(2) = Array(-19.709427090884, 24.083041414014, -33.814945054945)
    pts(3) = Array(-19.709427090884, 24.083041414014, -36.494945054945)
    FSData_208 = pts
End Function

Private Function FSData_209() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-15.919971892023, 25.663095973736, -36.616733037256)
    pts(1) = Array(-15.919971892023, 25.663095973736, -33.443706523183)
    pts(2) = Array(-16.404951168204, 26.444885652406, -33.690219780220)
    pts(3) = Array(-16.404951168204, 26.444885652406, -36.370219780220)
    FSData_209 = pts
End Function

Private Function FSData_210() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.441985323354, 27.517939625151, -36.492007762531)
    pts(1) = Array(-12.441985323354, 27.517939625151, -33.318981248458)
    pts(2) = Array(-12.821012690820, 28.356234474659, -33.565494505494)
    pts(3) = Array(-12.821012690820, 28.356234474659, -36.245494505495)
    FSData_210 = pts
End Function

Private Function FSData_211() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-8.752046340538, 28.904008110521, -36.367282487806)
    pts(1) = Array(-8.752046340538, 28.904008110521, -33.194255973733)
    pts(2) = Array(-9.018664970780, 29.784527562895, -33.440769230769)
    pts(3) = Array(-9.018664970780, 29.784527562895, -36.120769230769)
    FSData_211 = pts
End Function

Private Function FSData_212() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-4.913014001447, 29.797689397361, -36.242557213081)
    pts(1) = Array(-4.913014001447, 29.797689397361, -33.069530699007)
    pts(2) = Array(-5.062681977650, 30.705433577678, -33.316043956044)
    pts(3) = Array(-5.062681977650, 30.705433577678, -35.996043956044)
    FSData_212 = pts
End Function

Private Function FSData_213() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-0.990287207549, 30.183759395519, -36.117831938355)
    pts(1) = Array(-0.990287207549, 30.183759395519, -32.944805424282)
    pts(2) = Array(-1.020454897316, 31.103264648627, -33.191318681319)
    pts(3) = Array(-1.020454897316, 31.103264648627, -35.871318681319)
    FSData_213 = pts
End Function

Private Function FSData_214() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.949309383159, 30.055641303463, -35.993106663630)
    pts(1) = Array(2.949309383159, 30.055641303463, -32.820080149557)
    pts(2) = Array(3.039155894170, 30.971243621316, -33.066593406593)
    pts(3) = Array(3.039155894170, 30.971243621316, -35.746593406593)
    FSData_214 = pts
End Function

Private Function FSData_215() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(6.838663731366, 29.415517645782, -35.868381388905)
    pts(1) = Array(6.838663731366, 29.415517645782, -32.695354874831)
    pts(2) = Array(7.046993884772, 30.311619507839, -32.941868131868)
    pts(3) = Array(7.046993884772, 30.311619507839, -35.621868131868)
    FSData_215 = pts
End Function

Private Function FSData_216() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(10.611519687256, 28.274293093320, -35.743656114179)
    pts(1) = Array(10.611519687256, 28.274293093320, -32.570629600106)
    pts(2) = Array(10.934784525411, 29.135629174309, -32.817142857143)
    pts(3) = Array(10.934784525411, 29.135629174309, -35.497142857143)
    FSData_216 = pts
End Function

Private Function FSData_217() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(14.203605681037, 26.651408699309, -35.618930839454)
    pts(1) = Array(14.203605681037, 26.651408699309, -32.445904325381)
    pts(2) = Array(14.636298304433, 27.463305917964, -32.692417582418)
    pts(3) = Array(14.636298304433, 27.463305917964, -35.372417582418)
    FSData_217 = pts
End Function

Private Function FSData_218() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(17.553729605627, 24.574510716035, -35.494205564729)
    pts(1) = Array(17.553729605627, 24.574510716035, -32.321179050656)
    pts(2) = Array(18.088478984342, 25.323138194802, -32.567692307692)
    pts(3) = Array(18.088478984342, 25.323138194802, -35.247692307692)
    FSData_218 = pts
End Function

Private Function FSData_219() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.604821239912, 22.078979633834, -35.369480290004)
    pts(1) = Array(20.604821239912, 22.078979633834, -32.196453775930)
    pts(2) = Array(21.232517780996, 22.751584311421, -32.442967032967)
    pts(3) = Array(21.232517780996, 22.751584311421, -35.122967032967)
    FSData_219 = pts
End Function

Private Function FSData_220() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.304904454613, 19.207327465355, -35.244755015278)
    pts(1) = Array(23.304904454613, 19.207327465355, -32.071728501205)
    pts(2) = Array(24.014855186343, 19.792451348405, -32.318241758242)
    pts(3) = Array(24.014855186343, 19.792451348405, -34.998241758242)
    FSData_220 = pts
End Function

Private Function FSData_221() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.607982638988, 16.008473542511, -35.120029740553)
    pts(1) = Array(25.607982638988, 16.008473542511, -31.947003226480)
    pts(2) = Array(26.388093368388, 16.496148895462, -32.193516483516)
    pts(3) = Array(26.388093368388, 16.496148895462, -34.873516483516)
    FSData_221 = pts
End Function

Private Function FSData_222() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.474822264911, 12.536911163103, -34.995304465828)
    pts(1) = Array(27.474822264911, 12.536911163103, -31.822277951755)
    pts(2) = Array(28.311803605432, 12.918830311118, -32.068791208791)
    pts(3) = Array(28.311803605432, 12.918830311118, -34.748791208791)
    FSData_222 = pts
End Function

Private Function FSData_223() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.873621240091, 8.851779283499, -34.870579191103)
    pts(1) = Array(28.873621240091, 8.851779283499, -31.697552677029)
    pts(2) = Array(29.753214999723, 9.121436135844, -31.944065934066)
    pts(3) = Array(29.753214999723, 9.121436135844, -34.624065934066)
    FSData_223 = pts
End Function

Private Function FSData_224() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.780550664878, 5.015855071332, -34.745853916377)
    pts(1) = Array(29.780550664878, 5.015855071332, -31.572827402304)
    pts(2) = Array(30.687772738113, 5.168655954300, -31.819340659341)
    pts(3) = Array(30.687772738113, 5.168655954300, -34.499340659341)
    FSData_224 = pts
End Function

Private Function FSData_225() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.180160763692, 1.094484480331, -34.621128641652)
    pts(1) = Array(30.180160763692, 1.094484480331, -31.448102127579)
    pts(2) = Array(31.099556389606, 1.127826391653, -31.694615384615)
    pts(3) = Array(31.099556389606, 1.127826391653, -34.374615384615)
    FSData_225 = pts
End Function

Private Function FSData_226() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.065644075909, -2.845530934773, -34.496403366927)
    pts(1) = Array(30.065644075909, -2.845530934773, -31.323376852853)
    pts(2) = Array(30.981551113983, -2.932215983117, -31.569890109890)
    pts(3) = Array(30.981551113983, -2.932215983117, -34.249890109890)
    FSData_226 = pts
End Function

Private Function FSData_227() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.438951422705, -6.737071999885, -34.371678092201)
    pts(1) = Array(29.438951422705, -6.737071999885, -31.198651578128)
    pts(2) = Array(30.335767161410, -6.942307305842, -31.445164835165)
    pts(3) = Array(30.335767161410, -6.942307305842, -34.125164835165)
    FSData_227 = pts
End Function

Private Function FSData_228() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.310758674313, -10.513845313908, -34.246952817476)
    pts(1) = Array(28.310758674313, -10.513845313908, -31.073926303403)
    pts(2) = Array(29.173205627306, -10.834134641352, -31.320439560440)
    pts(3) = Array(29.173205627306, -10.834134641352, -34.000439560440)
    FSData_228 = pts
End Function

Private Function FSData_229() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.700284883839, -14.111512573847, -34.122227542751)
    pts(1) = Array(26.700284883839, -14.111512573847, -30.949201028678)
    pts(2) = Array(27.513671045863, -14.541399711858, -31.195714285714)
    pts(3) = Array(27.513671045863, -14.541399711858, -33.875714285714)
    FSData_229 = pts
End Function

Private Function FSData_230() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.634964885751, -17.468786594318, -33.997502268026)
    pts(1) = Array(24.634964885751, -17.468786594318, -30.824475753952)
    pts(2) = Array(25.385434014721, -18.000948305138, -31.070989010989)
    pts(3) = Array(25.385434014721, -18.000948305138, -33.750989010989)
    FSData_230 = pts
End Function

Private Function FSData_231() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.149981936434, -20.528475350490, -33.872776993300)
    pts(1) = Array(22.149981936434, -20.528475350490, -30.699750479227)
    pts(2) = Array(22.824749598073, -21.153846122756, -30.946263736264)
    pts(3) = Array(22.824749598073, -21.153846122756, -33.626263736264)
    FSData_231 = pts
End Function

Private Function FSData_232() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.287668358386, -23.238456258902, -33.748051718575)
    pts(1) = Array(19.287668358386, -23.238456258902, -30.575025204502)
    pts(2) = Array(19.875239712350, -23.946382740961, -30.821538461538)
    pts(3) = Array(19.875239712350, -23.946382740961, -33.501538461538)
    FSData_232 = pts
End Function

Private Function FSData_233() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(16.096784398236, -25.552564099063, -33.623326443850)
    pts(1) = Array(16.096784398236, -25.552564099063, -30.450299929777)
    pts(2) = Array(16.587150015665, -26.330986581551, -30.696813186813)
    pts(3) = Array(16.587150015665, -26.330986581551, -33.376813186813)
    FSData_233 = pts
End Function

Private Function FSData_234() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(12.631687583422, -27.431377449826, -33.498601169125)
    pts(1) = Array(12.631687583422, -27.431377449826, -30.325574655051)
    pts(2) = Array(13.016493960136, -28.267035305913, -30.572087912088)
    pts(3) = Array(13.016493960136, -28.267035305913, -33.252087912088)
    FSData_234 = pts
End Function

Private Function FSData_235() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(8.951406727761, -28.842890243424, -33.373875894399)
    pts(1) = Array(8.951406727761, -28.842890243424, -30.200849380326)
    pts(2) = Array(9.224098588342, -29.721547826998, -30.447362637363)
    pts(3) = Array(9.224098588342, -29.721547826998, -33.127362637363)
    FSData_235 = pts
End Function

Private Function FSData_236() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.118636360446, -29.763056997048, -33.249150619674)
    pts(1) = Array(5.118636360446, -29.763056997048, -30.076124105601)
    pts(2) = Array(5.274568329043, -30.669746150600, -30.322637362637)
    pts(3) = Array(5.274568329043, -30.669746150600, -33.002637362637)
    FSData_236 = pts
End Function

Private Function FSData_237() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.198668708652, -30.176202433820, -33.124425344949)
    pts(1) = Array(1.198668708652, -30.176202433820, -29.951398830875)
    pts(2) = Array(1.235184444148, -31.095477474851, -30.197912087912)
    pts(3) = Array(1.235184444148, -31.095477474851, -32.877912087912)
    FSData_237 = pts
End Function

Private Function FSData_238() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.741718572321, -30.075288515161, -32.999700070223)
    pts(1) = Array(-2.741718572321, -30.075288515161, -29.826673556150)
    pts(2) = Array(-2.825241124856, -30.991489357345, -30.073186813187)
    pts(3) = Array(-2.825241124856, -30.991489357345, -32.753186813187)
    FSData_238 = pts
End Function

Private Function FSData_239() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.635399973549, -29.462034335582, -32.874974795498)
    pts(1) = Array(-6.635399973549, -29.462034335582, -29.701948281425)
    pts(2) = Array(-6.837537985988, -30.359553262362, -29.948461538462)
    pts(3) = Array(-6.837537985988, -30.359553262362, -32.628461538462)
    FSData_239 = pts
End Function

Private Function FSData_240() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.416045632757, -28.346886837470, -32.750249520773)
    pts(1) = Array(-10.416045632757, -28.346886837470, -29.577223006700)
    pts(2) = Array(-10.733355632166, -29.210434383512, -29.823736263736)
    pts(3) = Array(-10.733355632166, -29.210434383512, -32.503736263736)
    FSData_240 = pts
End Function

Private Function FSData_241() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.019251280557, -26.748842844740, -32.625524246048)
    pts(1) = Array(-14.019251280557, -26.748842844740, -29.452497731974)
    pts(2) = Array(-14.446327809634, -27.563708255904, -29.699010989011)
    pts(3) = Array(-14.446327809634, -27.563708255904, -32.379010989011)
    FSData_241 = pts
End Function

Private Function FSData_242() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.383635383707, -24.695125447066, -32.500798971322)
    pts(1) = Array(-17.383635383707, -24.695125447066, -29.327772457249)
    pts(2) = Array(-17.913203084138, -25.447427281877, -29.574285714286)
    pts(3) = Array(-17.913203084138, -25.447427281877, -32.254285714286)
    FSData_242 = pts
End Function

Private Function FSData_243() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.451884795290, -22.220720247557, -32.376073696597)
    pts(1) = Array(-20.451884795290, -22.220720247557, -29.203047182524)
    pts(2) = Array(-21.074922345346, -22.897642851125, -29.449560439560)
    pts(3) = Array(-21.074922345346, -22.897642851125, -32.129560439560)
    FSData_243 = pts
End Function

Private Function FSData_244() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.171731098885, -19.367779374026, -32.251348421872)
    pts(1) = Array(-23.171731098885, -19.367779374026, -29.078321907799)
    pts(2) = Array(-23.877624893950, -19.957791196017, -29.324835164835)
    pts(3) = Array(-23.877624893950, -19.957791196017, -32.004835164835)
    FSData_244 = pts
End Function

Private Function FSData_245() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.496841014459, -16.184903406676, -32.126623147147)
    pts(1) = Array(-25.496841014459, -16.184903406676, -28.953596633073)
    pts(2) = Array(-26.273565972515, -16.677953444230, -29.200109890110)
    pts(3) = Array(-26.273565972515, -16.677953444230, -31.880109890110)
    FSData_245 = pts
End Function

Private Function FSData_246() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.387605697686, -12.726313454730, -32.001897872421)
    pts(1) = Array(-27.387605697686, -12.726313454730, -28.828871358348)
    pts(2) = Array(-28.221930109669, -13.114002473881, -29.075384615385)
    pts(3) = Array(-28.221930109669, -13.114002473881, -31.755384615385)
    FSData_246 = pts
End Function

Private Function FSData_247() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.811815486784, -9.050927485927, -31.877172597696)
    pts(1) = Array(-28.811815486784, -9.050927485927, -28.704146083623)
    pts(2) = Array(-29.689526422143, -9.326651104704, -28.950659340659)
    pts(3) = Array(-29.689526422143, -9.326651104704, -31.630659340659)
    FSData_247 = pts
End Function

Private Function FSData_248() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.745208602365, -5.221356643806, -31.752447322971)
    pts(1) = Array(-29.745208602365, -5.221356643806, -28.579420808897)
    pts(2) = Array(-30.651354029986, -5.380417839578, -28.825934065934)
    pts(3) = Array(-30.651354029986, -5.380417839578, -31.505934065934)
    FSData_248 = pts
End Function

Private Function FSData_249() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.171884453079, -1.302838650807, -31.627722048245)
    pts(1) = Array(-30.171884453079, -1.302838650807, -28.454695534172)
    pts(2) = Array(-31.091027952974, -1.342527775269, -28.701208791209)
    pts(3) = Array(-31.091027952974, -1.342527775269, -31.381208791209)
    FSData_249 = pts
End Function

Private Function FSData_250() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.084574506273, 2.637873533079, -31.502996773520)
    pts(1) = Array(-30.084574506273, 2.637873533079, -28.329970259447)
    pts(2) = Array(-31.001058232954, 2.718232594351, -28.576483516484)
    pts(3) = Array(-31.001058232954, 2.718232594351, -31.256483516484)
    FSData_250 = pts
End Function

Private Function FSData_251() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.484766109304, 6.533648864122, -31.378271498795)
    pts(1) = Array(-29.484766109304, 6.533648864122, -28.205244984722)
    pts(2) = Array(-30.382977527203, 6.732687173890, -28.451758241758)
    pts(3) = Array(-30.382977527203, 6.732687173890, -31.131758241758)
    FSData_251 = pts
End Function

Private Function FSData_252() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.382677152200, 10.318121809417, -31.253546224070)
    pts(1) = Array(-28.382677152200, 10.318121809417, -28.080519709996)
    pts(2) = Array(-29.247314999221, 10.632448698975, -28.327032967033)
    pts(3) = Array(-29.247314999221, 10.632448698975, -31.007032967033)
    FSData_252 = pts
End Function

Private Function FSData_253() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-26.797082003280, 13.926822900773, -31.128820949344)
    pts(1) = Array(-26.797082003280, 13.926822900773, -27.955794435271)
    pts(2) = Array(-27.613416951724, 14.351083730863, -28.202307692308)
    pts(3) = Array(-27.613416951724, 14.351083730863, -30.882307692308)
    FSData_253 = pts
End Function

Private Function FSData_254() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.754991682965, 17.298276988658, -31.004095674619)
    pts(1) = Array(-24.754991682965, 17.298276988658, -27.831069160546)
    pts(2) = Array(-25.509117257413, 17.825244367120, -28.077582417582)
    pts(3) = Array(-25.509117257413, 17.825244367120, -30.757582417582)
    FSData_254 = pts
End Function

Private Function FSData_255() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.291193724118, 20.375050487149, -30.879370399894)
    pts(1) = Array(-22.291193724118, 20.375050487149, -27.706343885821)
    pts(2) = Array(-22.970263201806, 20.995747389406, -27.952857142857)
    pts(3) = Array(-22.970263201806, 20.995747389406, -30.632857142857)
    FSData_255 = pts
End Function

Private Function FSData_256() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.447659557484, 23.104729769816, -30.754645125169)
    pts(1) = Array(-19.447659557484, 23.104729769816, -27.581618611095)
    pts(2) = Array(-20.040104815526, 23.808582464790, -27.828131868132)
    pts(3) = Array(-20.040104815526, 23.808582464790, -30.508131868132)
    FSData_256 = pts
End Function

Private Function FSData_257() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.272829517597, 25.440814049303, -30.629919850443)
    pts(1) = Array(-16.272829517597, 25.440814049303, -27.456893336370)
    pts(2) = Array(-16.768558098927, 26.215832225640, -27.703406593407)
    pts(3) = Array(-16.768558098927, 26.215832225640, -30.383406593407)
    FSData_257 = pts
End Function

Private Function FSData_258() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.820787649243, 27.343507530180, -30.505194575718)
    pts(1) = Array(-12.820787649243, 27.343507530180, -27.332168061645)
    pts(2) = Array(-13.211354690213, 28.176488554278, -27.578681318681)
    pts(3) = Array(-13.211354690213, 28.176488554278, -30.258681318681)
    FSData_258 = pts
End Function

Private Function FSData_259() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.150340371874, 28.780397340531, -30.380469300993)
    pts(1) = Array(-9.150340371874, 28.780397340531, -27.207442786919)
    pts(2) = Array(-9.429092462673, 29.657151166799, -27.453956043956)
    pts(3) = Array(-9.429092462673, 29.657151166799, -30.133956043956)
    FSData_259 = pts
End Function

Private Function FSData_260() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.324014697153, 29.727005693552, -30.255744026267)
    pts(1) = Array(-5.324014697153, 29.727005693552, -27.082717512194)
    pts(2) = Array(-5.486203224351, 30.632596595475, -27.329230769231)
    pts(3) = Array(-5.486203224351, 30.632596595475, -30.009230769231)
    FSData_260 = pts
End Function

Private Function FSData_261() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.406993065261, 30.167206872933, -30.131018751542)
    pts(1) = Array(-1.406993065261, 30.167206872933, -26.957992237469)
    pts(2) = Array(-1.449855105660, 31.086207877009, -27.204505494505)
    pts(3) = Array(-1.449855105660, 31.086207877009, -29.884505494506)
    FSData_261 = pts
End Function

Private Function FSData_262() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.533997054708, 30.093501938570, -30.006293476817)
    pts(1) = Array(2.533997054708, 30.093501938570, -26.833266962744)
    pts(2) = Array(2.611191666970, 31.010257626765, -27.079780219780)
    pts(3) = Array(2.611191666970, 31.010257626765, -29.759780219780)
    FSData_262 = pts
End Function

Private Function FSData_263() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(6.431819884310, 29.507146472944, -29.881568202092)
    pts(1) = Array(6.431819884310, 29.507146472944, -26.708541688018)
    pts(2) = Array(6.627756119196, 30.406039676756, -26.955054945055)
    pts(3) = Array(6.627756119196, 30.406039676756, -29.635054945055)
    FSData_263 = pts
End Function

Private Function FSData_264() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(10.220075010977, 28.418129191944, -29.756842927366)
    pts(1) = Array(10.220075010977, 28.418129191944, -26.583816413293)
    pts(2) = Array(10.531415044424, 29.283847034877, -26.830329670330)
    pts(3) = Array(10.531415044424, 29.283847034877, -29.510329670330)
    FSData_264 = pts
End Function

Private Function FSData_265() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.834228536088, 26.845001784528, -29.632117652641)
    pts(1) = Array(13.834228536088, 26.845001784528, -26.459091138568)
    pts(2) = Array(14.255668610697, 27.662796540877, -26.705604395604)
    pts(3) = Array(14.255668610697, 27.662796540877, -29.385604395604)
    FSData_265 = pts
End Function

Private Function FSData_266() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(17.212712426501, 24.814562879942, -29.507392377916)
    pts(1) = Array(17.212712426501, 24.814562879942, -26.334365863842)
    pts(2) = Array(17.737073202408, 25.570503206086, -26.580879120879)
    pts(3) = Array(17.737073202408, 25.570503206086, -29.260879120879)
    FSData_266 = pts
End Function

Private Function FSData_267() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.297973341803, 22.361401526189, -29.382667103190)
    pts(1) = Array(20.297973341803, 22.361401526189, -26.209640589117)
    pts(2) = Array(20.916322198574, 23.042609784603, -26.456153846154)
    pts(3) = Array(20.916322198574, 23.042609784603, -29.136153846154)
    FSData_267 = pts
End Function

Private Function FSData_268() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.037453070240, 19.527307956717, -29.257941828465)
    pts(1) = Array(23.037453070240, 19.527307956717, -26.084915314392)
    pts(2) = Array(23.739256276354, 20.122179589836, -26.331428571429)
    pts(3) = Array(23.739256276354, 20.122179589836, -29.011428571429)
    FSData_268 = pts
End Function

Private Function FSData_269() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.384483871347, 16.360561683064, -29.133216553740)
    pts(1) = Array(25.384483871347, 16.360561683064, -25.960190039667)
    pts(2) = Array(26.157786029017, 16.858962899899, -26.206703296703)
    pts(3) = Array(26.157786029017, 16.858962899899, -28.886703296703)
    FSData_269 = pts
End Function

Private Function FSData_270() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.299083472885, 12.915109040982, -29.008491279015)
    pts(1) = Array(27.299083472885, 12.915109040982, -25.835464764941)
    pts(2) = Array(28.130711181330, 13.308549448853, -26.081978021978)
    pts(3) = Array(28.130711181330, 13.308549448853, -28.761978021978)
    FSData_270 = pts
End Function

Private Function FSData_271() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.748636179118, 9.249644200762, -28.883766004289)
    pts(1) = Array(28.748636179118, 9.249644200762, -25.710739490216)
    pts(2) = Array(29.624422446826, 9.531421441315, -25.957252747253)
    pts(3) = Array(29.624422446826, 9.531421441315, -28.637252747253)
    FSData_271 = pts
End Function

Private Function FSData_272() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.708448487560, 5.426609296972, -28.759040729564)
    pts(1) = Array(29.708448487560, 5.426609296972, -25.586014215491)
    pts(2) = Array(30.613474070625, 5.591923222575, -25.832527472527)
    pts(3) = Array(30.613474070625, 5.591923222575, -28.512527472527)
    FSData_272 = pts
End Function

Private Function FSData_273() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.162169749129, 1.511130710664, -28.634315454839)
    pts(1) = Array(30.162169749129, 1.511130710664, -25.461288940766)
    pts(2) = Array(31.081017304401, 1.557165156154, -25.707802197802)
    pts(3) = Array(31.081017304401, 1.557165156154, -28.387802197802)
    FSData_273 = pts
End Function

Private Function FSData_274() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.102070705653, -2.430090375245, -28.509590180114)
    pts(1) = Array(30.102070705653, -2.430090375245, -25.336563666040)
    pts(2) = Array(31.019087429137, -2.504119618465, -25.583076923077)
    pts(3) = Array(31.019087429137, -2.504119618465, -28.263076923077)
    FSData_274 = pts
End Function

Private Function FSData_275() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.529175159766, -6.329914247748, -28.384864905388)
    pts(1) = Array(29.529175159766, -6.329914247748, -25.211838391315)
    pts(2) = Array(30.428739436157, -6.522746072514, -25.458351648352)
    pts(3) = Array(30.428739436157, -6.522746072514, -28.138351648352)
    FSData_275 = pts
End Function

Private Function FSData_276() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.453242534169, -10.121906405996, -28.260139630663)
    pts(1) = Array(28.453242534169, -10.121906405996, -25.087113116590)
    pts(2) = Array(29.320030055078, -10.430255872668, -25.333626373626)
    pts(3) = Array(29.320030055078, -10.430255872668, -28.013626373626)
    FSData_276 = pts
End Function

Private Function FSData_277() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.892601617358, -13.741469290076, -28.135414355938)
    pts(1) = Array(26.892601617358, -13.741469290076, -24.962387841864)
    pts(2) = Array(27.711846434840, -14.160083586330, -25.208901098901)
    pts(3) = Array(27.711846434840, -14.160083586330, -27.888901098901)
    FSData_277 = pts
End Function

Private Function FSData_278() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.873838328004, -17.126942717027, -28.010689081212)
    pts(1) = Array(24.873838328004, -17.126942717027, -24.837662567139)
    pts(2) = Array(25.631584396274, -17.648690640857, -25.084175824176)
    pts(3) = Array(25.631584396274, -17.648690640857, -27.764175824176)
    FSData_278 = pts
End Function

Private Function FSData_279() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.431342817008, -20.220654277888, -27.885963806487)
    pts(1) = Array(22.431342817008, -20.220654277888, -24.712937292414)
    pts(2) = Array(23.114681737261, -20.836647719466, -24.959450549451)
    pts(3) = Array(23.114681737261, -20.836647719466, -27.639450549451)
    FSData_279 = pts
End Function

Private Function FSData_280() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.606723622449, -22.969901801987, -27.761238531762)
    pts(1) = Array(19.606723622449, -22.969901801987, -24.588212017689)
    pts(2) = Array(20.204014540748, -23.669647154895, -24.834725274725)
    pts(3) = Array(20.204014540748, -23.669647154895, -27.514725274725)
    FSData_280 = pts
End Function

Private Function FSData_281() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(16.448098857454, -25.327851151952, -27.636513257037)
    pts(1) = Array(16.448098857454, -25.327851151952, -24.463486742963)
    pts(2) = Array(16.949166769668, -26.099428074462, -24.710000000000)
    pts(3) = Array(16.949166769668, -26.099428074462, -27.390000000000)
    FSData_281 = pts
End Function

Private Function FSData_282() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.009276505792, -27.254334055263, -27.511787982311)
    pts(1) = Array(13.009276505792, -27.254334055263, -24.338761468238)
    pts(2) = Array(13.405585591399, -28.084598536416, -24.585274725275)
    pts(3) = Array(13.405585591399, -28.084598536416, -27.265274725275)
    FSData_282 = pts
End Function

Private Function FSData_283() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(9.348837789054, -28.716532381086, -27.387062707586)
    pts(1) = Array(9.348837789054, -28.716532381086, -24.214036193513)
    pts(2) = Array(9.633636821038, -29.591340652298, -24.460549450549)
    pts(3) = Array(9.633636821038, -29.591340652298, -27.140549450549)
    FSData_283 = pts
End Function

Private Function FSData_284() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.529139220503, -29.689537205559, -27.262337432861)
    pts(1) = Array(5.529139220503, -29.689537205559, -24.089310918788)
    pts(2) = Array(5.697576574240, -30.593986683345, -24.335824175824)
    pts(3) = Array(5.697576574240, -30.593986683345, -27.015824175824)
    FSData_284 = pts
End Function

Private Function FSData_285() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.615250345865, -30.156773141704, -27.137612158136)
    pts(1) = Array(1.615250345865, -30.156773141704, -23.964585644062)
    pts(2) = Array(1.664456647792, -31.075456297014, -24.211098901099)
    pts(3) = Array(1.664456647792, -31.075456297014, -26.891098901099)
    FSData_285 = pts
End Function

Private Function FSData_286() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.326154733089, -30.110280705396, -27.012886883410)
    pts(1) = Array(-2.326154733089, -30.110280705396, -23.839860369337)
    pts(2) = Array(-2.397017724958, -31.027547534832, -24.086373626374)
    pts(3) = Array(-2.397017724958, -31.027547534832, -26.766373626374)
    FSData_286 = pts
End Function

Private Function FSData_287() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.227933168984, -29.550851907224, -26.888161608685)
    pts(1) = Array(-6.227933168984, -29.550851907224, -23.715135094612)
    pts(2) = Array(-6.417658285390, -30.451076534861, -23.961648351648)
    pts(3) = Array(-6.417658285390, -30.451076534861, -26.641648351648)
    FSData_287 = pts
End Function

Private Function FSData_288() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.023617164482, -28.488016760384, -26.763436333960)
    pts(1) = Array(-10.023617164482, -28.488016760384, -23.590409819886)
    pts(2) = Array(-10.328972389360, -29.355863628581, -23.836923076923)
    pts(3) = Array(-10.328972389360, -29.355863628581, -26.516923076923)
    FSData_288 = pts
End Function

Private Function FSData_289() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.648546268274, -26.939880934458, -26.638711059234)
    pts(1) = Array(-13.648546268274, -26.939880934458, -23.465684545161)
    pts(2) = Array(-14.064329796976, -27.760566049018, -23.712197802198)
    pts(3) = Array(-14.064329796976, -27.760566049018, -26.392197802198)
    FSData_289 = pts
End Function

Private Function FSData_290() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.040968882470, -24.932817320686, -26.513985784509)
    pts(1) = Array(-17.040968882470, -24.932817320686, -23.340959270436)
    pts(2) = Array(-17.560097735844, -25.692360099991, -23.587472527473)
    pts(3) = Array(-17.560097735844, -25.692360099991, -26.267472527473)
    FSData_290 = pts
End Function

Private Function FSData_291() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.143094216919, -22.501016762990, -26.389260509784)
    pts(1) = Array(-20.143094216919, -22.501016762990, -23.216233995711)
    pts(2) = Array(-20.756724901672, -23.186478200803, -23.462747252747)
    pts(3) = Array(-20.756724901672, -23.186478200803, -26.142747252747)
    FSData_291 = pts
End Function

Private Function FSData_292() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.902076770156, -19.685905608173, -26.264535235059)
    pts(1) = Array(-22.902076770156, -19.685905608173, -23.091508720985)
    pts(2) = Array(-23.599755930041, -20.285608692925, -23.338021978022)
    pts(3) = Array(-23.599755930041, -20.285608692925, -26.018021978022)
    FSData_292 = pts
End Function

Private Function FSData_293() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.270916566088, -16.535439997467, -26.139809960333)
    pts(1) = Array(-25.270916566088, -16.535439997467, -22.966783446260)
    pts(2) = Array(-26.040759057505, -17.039168633151, -23.213296703297)
    pts(3) = Array(-26.040759057505, -17.039168633151, -25.893296703297)
    FSData_293 = pts
End Function

Private Function FSData_294() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.209259810653, -13.103288921350, -26.015084685608)
    pts(1) = Array(-27.209259810653, -13.103288921350, -22.842058171535)
    pts(2) = Array(-28.038151169123, -13.502461961339, -23.088571428571)
    pts(3) = Array(-28.038151169123, -13.502461961339, -25.768571428571)
    FSData_294 = pts
End Function

Private Function FSData_295() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.684086329059, -9.447919954523, -25.890359410883)
    pts(1) = Array(-28.684086329059, -9.447919954523, -22.717332896810)
    pts(2) = Array(-29.557906177494, -9.735737383601, -22.963846153846)
    pts(3) = Array(-29.557906177494, -9.735737383601, -25.643846153846)
    FSData_295 = pts
End Function

Private Function FSData_296() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.670272072941, -5.631603245758, -25.765634136158)
    pts(1) = Array(-29.670272072941, -5.631603245758, -22.592607622084)
    pts(2) = Array(-30.574134665892, -5.803162020132, -22.839120879121)
    pts(3) = Array(-30.574134665892, -5.803162020132, -25.519120879121)
    FSData_296 = pts
End Function

Private Function FSData_297() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.151017114975, -1.719350729930, -25.640908861432)
    pts(1) = Array(-30.151017114975, -1.719350729930, -22.467882347359)
    pts(2) = Array(-31.069524921126, -1.771728301835, -22.714395604396)
    pts(3) = Array(-31.069524921126, -1.771728301835, -25.394395604396)
    FSData_297 = pts
End Function

Private Function FSData_298() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.118131839948, 2.222191366981, -25.516183586707)
    pts(1) = Array(-30.118131839948, 2.222191366981, -22.343157072634)
    pts(2) = Array(-31.035637843020, 2.289887262929, -22.589670329670)
    pts(3) = Array(-31.035637843020, 2.289887262929, -25.269670329670)
    FSData_298 = pts
End Function

Private Function FSData_299() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.572176456966, 6.125877863465, -25.391458311982)
    pts(1) = Array(-29.572176456966, 6.125877863465, -22.218431797908)
    pts(2) = Array(-30.473050706648, 6.312494010299, -22.464945054945)
    pts(3) = Array(-30.473050706648, 6.312494010299, -25.144945054945)
    FSData_299 = pts
End Function

Private Function FSData_300() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.522451456135, 9.925208457881, -25.266733037256)
    pts(1) = Array(-28.522451456135, 9.925208457881, -22.093706523183)
    pts(2) = Array(-29.391347328309, 10.227565801631, -22.340219780220)
    pts(3) = Array(-29.391347328309, 10.227565801631, -25.020219780220)
    FSData_300 = pts
End Function

Private Function FSData_301() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-26.986839172337, 13.555460578173, -25.142007762531)
    pts(1) = Array(-26.986839172337, 13.555460578173, -21.968981248458)
    pts(2) = Array(-27.808954802753, 13.968408383866, -22.215494505495)
    pts(3) = Array(-27.808954802753, 13.968408383866, -24.895494505495)
    FSData_301 = pts
End Function

Private Function FSData_302() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.991499155054, 16.954791947497, -25.017282487806)
    pts(1) = Array(-24.991499155054, 16.954791947497, -21.844255973733)
    pts(2) = Array(-25.752829592890, 17.471295543249, -22.090769230769)
    pts(3) = Array(-25.752829592890, 17.471295543249, -24.770769230769)
    FSData_302 = pts
End Function

Private Function FSData_303() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.570422533735, 20.065294083284, -24.892557213081)
    pts(1) = Array(-22.570422533735, 20.065294083284, -21.719530699007)
    pts(2) = Array(-23.257998319531, 20.676554697742, -21.966043956044)
    pts(3) = Array(-23.257998319531, 20.676554697742, -24.646043956044)
    FSData_303 = pts
End Function

Private Function FSData_304() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.764852970172, 22.833978783110, -24.767831938355)
    pts(1) = Array(-19.764852970172, 22.833978783110, -21.594805424282)
    pts(2) = Array(-20.366961073899, 23.529583434781, -21.841318681319)
    pts(3) = Array(-20.366961073899, 23.529583434781, -24.521318681319)
    FSData_304 = pts
End Function

Private Function FSData_305() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.622584062139, 25.213680792322, -24.643106663630)
    pts(1) = Array(-16.622584062139, 25.213680792322, -21.470080149557)
    pts(2) = Array(-17.128967417675, 25.981779677386, -21.716593406593)
    pts(3) = Array(-17.128967417675, 25.981779677386, -24.396593406593)
    FSData_305 = pts
End Function

Private Function FSData_306() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.197145167183, 27.163861276267, -24.518381388905)
    pts(1) = Array(-13.197145167183, 27.163861276267, -21.345354874832)
    pts(2) = Array(-13.599177404064, 27.991369633028, -21.591868131868)
    pts(3) = Array(-13.599177404064, 27.991369633028, -24.271868131868)
    FSData_306 = pts
End Function

Private Function FSData_307() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.546889516273, 28.651298409742, -24.393656114179)
    pts(1) = Array(-9.546889516273, 28.651298409742, -21.220629600106)
    pts(2) = Array(-9.837721912133, 29.524119420899, -21.467142857143)
    pts(3) = Array(-9.837721912133, 29.524119420899, -24.147142857143)
    FSData_307 = pts
End Function

Private Function FSData_308() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.734000151533, 29.650653319315, -24.268930839454)
    pts(1) = Array(-5.734000151533, 29.650653319315, -21.095904325381)
    pts(2) = Array(-5.908678301845, 30.553918254871, -21.342417582418)
    pts(3) = Array(-5.908678301845, 30.553918254871, -24.022417582418)
    FSData_308 = pts
End Function

Private Function FSData_309() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.823430622153, 30.144901737544, -24.144205564729)
    pts(1) = Array(-1.823430622153, 30.144901737544, -20.971179050656)
    pts(2) = Array(-1.878978839782, 31.063223247430, -21.217692307692)
    pts(3) = Array(-1.878978839782, 31.063223247430, -23.897692307692)
    FSData_309 = pts
End Function

Private Function FSData_310() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.118201515995, 30.125624015738, -24.019480290004)
    pts(1) = Array(2.118201515995, 30.125624015738, -20.846453775930)
    pts(2) = Array(2.182729509197, 31.043358257278, -21.092967032967)
    pts(3) = Array(2.182729509197, 31.043358257278, -23.772967032967)
    FSData_310 = pts
End Function

Private Function FSData_311() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(6.023749547523, 29.593148554838, -23.894755015278)
    pts(1) = Array(6.023749547523, 29.593148554838, -20.721728501205)
    pts(2) = Array(6.207254500627, 30.494661689621, -20.968241758242)
    pts(3) = Array(6.207254500627, 30.494661689621, -23.648241758242)
    FSData_311 = pts
End Function

Private Function FSData_312() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(9.826681459066, 28.556546211019, -23.770029740553)
    pts(1) = Array(9.826681459066, 28.556546211019, -20.597003226480)
    pts(2) = Array(10.126037318083, 29.426480731355, -20.843516483516)
    pts(3) = Array(10.126037318083, 29.426480731355, -23.523516483516)
    FSData_312 = pts
End Function

Private Function FSData_313() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.462213329203, 27.033475771329, -23.645304465828)
    pts(1) = Array(13.462213329203, 27.033475771329, -20.472277951755)
    pts(2) = Array(13.872320490225, 27.857012119330, -20.718791208791)
    pts(3) = Array(13.872320490225, 27.857012119330, -23.398791208791)
    FSData_313 = pts
End Function

Private Function FSData_314() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(16.868412939196, 25.049883131719, -23.520579191103)
    pts(1) = Array(16.868412939196, 25.049883131719, -20.347552677029)
    pts(2) = Array(17.382285121450, 25.812992154274, -20.594065934066)
    pts(3) = Array(17.382285121450, 25.812992154274, -23.274065934066)
    FSData_314 = pts
End Function

Private Function FSData_315() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.987254804235, 22.639559302040, -23.395853916377)
    pts(1) = Array(19.987254804235, 22.639559302040, -20.222827402304)
    pts(2) = Array(20.596138063172, 23.329241241042, -20.469340659341)
    pts(3) = Array(20.596138063172, 23.329241241042, -23.149340659341)
    FSData_315 = pts
End Function

Private Function FSData_316() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.765608652465, 19.843564767521, -23.271128641652)
    pts(1) = Array(22.765608652465, 19.843564767521, -20.098102127579)
    pts(2) = Array(23.459130505454, 20.448070714081, -20.344615384615)
    pts(3) = Array(23.459130505454, 20.448070714081, -23.024615384615)
    FSData_316 = pts
End Function

Private Function FSData_317() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.156144512811, 16.709530012857, -23.146403366927)
    pts(1) = Array(25.156144512811, 16.709530012857, -19.973376852853)
    pts(2) = Array(25.922490637042, 17.218562052984, -20.219890109890)
    pts(3) = Array(25.922490637042, 17.218562052984, -22.899890109890)
    FSData_317 = pts
End Function

Private Function FSData_318() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.118138993182, 13.290844124678, -23.021678092201)
    pts(1) = Array(27.118138993182, 13.290844124678, -19.848651578128)
    pts(2) = Array(27.944254485689, 13.695730766887, -20.095164835165)
    pts(3) = Array(27.944254485689, 13.695730766887, -22.775164835165)
    FSData_318 = pts
End Function

Private Function FSData_319() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.618169013912, 9.645745294750, -22.896952817476)
    pts(1) = Array(28.618169013912, 9.645745294750, -19.723926303403)
    pts(2) = Array(29.489980785196, 9.939589191146, -19.970439560440)
    pts(3) = Array(29.489980785196, 9.939589191146, -22.650439560440)
    FSData_319 = pts
End Function

Private Function FSData_320() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.630681178504, 5.836328717426, -22.772227542751)
    pts(1) = Array(29.630681178504, 5.836328717426, -19.599201028678)
    pts(2) = Array(30.533337691227, 6.014124161798, -19.845714285714)
    pts(3) = Array(30.533337691227, 6.014124161798, -22.525714285714)
    FSData_320 = pts
End Function

Private Function FSData_321() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.138427082298, 1.927488782072, -22.647502268026)
    pts(1) = Array(30.138427082298, 1.927488782072, -19.474475753952)
    pts(2) = Array(31.056551351030, 1.986206983380, -19.720989010989)
    pts(3) = Array(31.056551351030, 1.986206983380, -22.400989010989)
    FSData_321 = pts
End Function

Private Function FSData_322() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.132757143471, -2.014186419519, -22.522776993300)
    pts(1) = Array(30.132757143471, -2.014186419519, -19.349750479227)
    pts(2) = Array(31.050708685590, -2.075545740908, -19.596263736264)
    pts(3) = Array(31.050708685590, -2.075545740908, -22.276263736264)
    FSData_322 = pts
End Function

Private Function FSData_323() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.613767950888, -5.921549438361, -22.398051718575)
    pts(1) = Array(29.613767950888, -5.921549438361, -19.225025204502)
    pts(2) = Array(30.515909226213, -6.101941010656, -19.471538461538)
    pts(3) = Array(30.515909226213, -6.101941010656, -22.151538461538)
    FSData_323 = pts
End Function

Private Function FSData_324() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.590300618681, -9.728037342314, -22.273326443850)
    pts(1) = Array(28.590300618681, -9.728037342314, -19.100299929777)
    pts(2) = Array(29.461263418985, -10.024388148769, -19.346813186813)
    pts(3) = Array(29.461263418985, -10.024388148769, -22.026813186813)
    FSData_324 = pts
End Function

Private Function FSData_325() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.079790175603, -13.368805632716, -22.148601169125)
    pts(1) = Array(27.079790175603, -13.368805632716, -18.975574655051)
    pts(2) = Array(27.904737425985, -13.776067261262, -19.222087912088)
    pts(3) = Array(27.904737425985, -13.776067261262, -21.902087912088)
    FSData_325 = pts
End Function

Private Function FSData_326() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.107968554837, -16.781832887063, -22.023875894399)
    pts(1) = Array(25.107968554837, -16.781832887063, -18.850849380326)
    pts(2) = Array(25.872847067104, -17.293067531305, -19.097362637363)
    pts(3) = Array(25.872847067104, -17.293067531305, -21.777362637363)
    FSData_326 = pts
End Function

Private Function FSData_327() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.708426243908, -19.908977309871, -21.899150619674)
    pts(1) = Array(22.708426243908, -19.908977309871, -18.726124105601)
    pts(2) = Array(23.400206116239, -20.515475956397, -18.972637362637)
    pts(3) = Array(23.400206116239, -20.515475956397, -21.652637362637)
    FSData_327 = pts
End Function

Private Function FSData_328() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.922040062105, -22.696967193083, -21.774425344949)
    pts(1) = Array(19.922040062105, -22.696967193083, -18.601398830875)
    pts(2) = Array(20.528936646778, -23.388397981747, -18.847912087912)
    pts(3) = Array(20.528936646778, -23.388397981747, -21.527912087912)
    FSData_328 = pts
End Function

Private Function FSData_329() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(16.796276813367, -25.098308413293, -21.649700070223)
    pts(1) = Array(16.796276813367, -25.098308413293, -18.476673556150)
    pts(2) = Array(17.307951471258, -25.862892643101, -18.723186813187)
    pts(3) = Array(17.307951471258, -25.862892643101, -21.403186813187)
    FSData_329 = pts
End Function

Private Function FSData_330() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.384384677097, -27.072093506330, -21.524974795498)
    pts(1) = Array(13.384384677097, -27.072093506330, -18.351948281425)
    pts(2) = Array(13.792120899048, -27.896806288642, -18.598461538462)
    pts(3) = Array(13.792120899048, -27.896806288642, -21.278461538462)
    FSData_330 = pts
End Function

Private Function FSData_331() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(9.744486111753, -28.584698536417, -21.400249520773)
    pts(1) = Array(9.744486111753, -28.584698536417, -18.227223006700)
    pts(2) = Array(10.041338006548, -29.455490677262, -18.473736263736)
    pts(3) = Array(10.041338006548, -29.455490677262, -21.153736263736)
    FSData_331 = pts
End Function

Private Function FSData_332() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.938587723847, -29.610355888543, -21.275524246048)
    pts(1) = Array(5.938587723847, -29.610355888543, -18.102497731974)
    pts(2) = Array(6.119498343249, -30.512393220247, -18.349010989011)
    pts(3) = Array(6.119498343249, -30.512393220247, -21.029010989011)
    FSData_332 = pts
End Function

Private Function FSData_333() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.031523969484, -30.131593226403, -21.150798971322)
    pts(1) = Array(2.031523969484, -30.131593226403, -17.977772457249)
    pts(2) = Array(2.093411454647, -31.049509311446, -18.224285714286)
    pts(3) = Array(2.093411454647, -31.049509311446, -20.904285714286)
    FSData_333 = pts
End Function

Private Function FSData_334() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.910147317242, -30.139531138132, -21.026073696597)
    pts(1) = Array(-1.910147317242, -30.139531138132, -17.853047182524)
    pts(2) = Array(-1.968337235515, -31.057689040353, -18.099560439560)
    pts(3) = Array(-1.968337235515, -31.057689040353, -20.779560439560)
    FSData_334 = pts
End Function

Private Function FSData_335() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.819278754036, -29.634034399366, -20.901348421872)
    pts(1) = Array(-5.819278754036, -29.634034399366, -17.728321907799)
    pts(2) = Array(-5.996554795550, -30.536793063188, -17.974835164835)
    pts(3) = Array(-5.996554795550, -30.536793063188, -20.654835164835)
    FSData_335 = pts
End Function

Private Function FSData_336() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.629277283304, -28.623714276824, -20.776623147147)
    pts(1) = Array(-9.629277283304, -28.623714276824, -17.603596633073)
    pts(2) = Array(-9.922619505180, -29.495694976648, -17.850109890110)
    pts(3) = Array(-9.922619505180, -29.495694976648, -20.530109890110)
    FSData_336 = pts
End Function

Private Function FSData_337() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.275238601979, -27.125781833166, -20.651897872421)
    pts(1) = Array(-13.275238601979, -27.125781833166, -17.478871358348)
    pts(2) = Array(-13.679649844159, -27.952130153911, -17.725384615385)
    pts(3) = Array(-13.679649844159, -27.952130153911, -20.405384615385)
    FSData_337 = pts
End Function

Private Function FSData_338() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.695052822990, -25.165754732127, -20.527172597696)
    pts(1) = Array(-16.695052822990, -25.165754732127, -17.354146083623)
    pts(2) = Array(-17.203643836141, -25.932393618006, -17.600659340659)
    pts(3) = Array(-17.203643836141, -25.932393618006, -20.280659340659)
    FSData_338 = pts
End Function

Private Function FSData_339() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.830462533132, -22.777022538559, -20.402447322971)
    pts(1) = Array(-19.830462533132, -22.777022538559, -17.229420808897)
    pts(2) = Array(-20.434569338777, -23.470892099337, -17.475934065934)
    pts(3) = Array(-20.434569338777, -23.470892099337, -20.155934065934)
    FSData_339 = pts
End Function

Private Function FSData_340() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.628055223056, -20.000277918627, -20.277722048245)
    pts(1) = Array(-22.628055223056, -20.000277918627, -17.104695534172)
    pts(2) = Array(-23.317386706672, -20.609557908201, -17.351208791209)
    pts(3) = Array(-23.317386706672, -20.609557908201, -20.031208791209)
    FSData_340 = pts
End Function

Private Function FSData_341() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.040173183078, -16.882823429789, -20.152996773520)
    pts(1) = Array(-25.040173183078, -16.882823429789, -16.979970259447)
    pts(2) = Array(-25.802986405874, -17.397134607121, -17.226483516484)
    pts(3) = Array(-25.802986405874, -17.397134607121, -19.906483516484)
    FSData_341 = pts
End Function

Private Function FSData_342() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.025725364500, -13.477765709590, -20.028271498795)
    pts(1) = Array(-27.025725364500, -13.477765709590, -16.855244984722)
    pts(2) = Array(-27.849025607392, -13.888346651736, -17.101758241758)
    pts(3) = Array(-27.849025607392, -13.888346651736, -19.781758241758)
    FSData_342 = pts
End Function

Private Function FSData_343() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.550887376172, -9.843110790455, -19.903546224070)
    pts(1) = Array(-28.550887376172, -9.843110790455, -16.730519709996)
    pts(2) = Array(-29.420649508161, -10.142967145661, -16.977032967033)
    pts(3) = Array(-29.420649508161, -10.142967145661, -19.657032967033)
    FSData_343 = pts
End Function

Private Function FSData_344() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.589677691676, -6.040775952037, -19.778820949344)
    pts(1) = Array(-29.589677691676, -6.040775952037, -16.605794435271)
    pts(2) = Array(-30.491085091555, -6.224799590311, -16.852307692308)
    pts(3) = Array(-30.491085091555, -6.224799590311, -19.532307692308)
    FSData_344 = pts
End Function

Private Function FSData_345() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.124400251308, -2.135534944460, -19.654095674619)
    pts(1) = Array(-30.124400251308, -2.135534944460, -16.481069160546)
    pts(2) = Array(-31.042097212606, -2.200590975880, -16.727582417582)
    pts(3) = Array(-31.042097212606, -2.200590975880, -19.407582417582)
    FSData_345 = pts
End Function

Private Function FSData_346() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-30.145945918986, 1.806085449140, -19.529370399894)
    pts(1) = Array(-30.145945918986, 1.806085449140, -16.356343885821)
    pts(2) = Array(-31.064299238372, 1.861105270769, -16.602857142857)
    pts(3) = Array(-31.064299238372, 1.861105270769, -19.282857142857)
    FSData_346 = pts
End Function

Private Function FSData_347() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.653947658730, 5.716938713447, -19.404645125168)
    pts(1) = Array(-29.653947658730, 5.716938713447, -16.231618611095)
    pts(2) = Array(-30.557312951645, 5.891097111340, -16.478131868132)
    pts(3) = Array(-30.557312951645, 5.891097111340, -19.158131868132)
    FSData_347 = pts
End Function

Private Function FSData_348() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-28.656786787212, 9.530402459093, -19.279919850443)
    pts(1) = Array(-28.656786787212, 9.530402459093, -16.106893336370)
    pts(2) = Array(-29.529774993975, 9.820732600230, -16.353406593407)
    pts(3) = Array(-29.529774993975, 9.820732600230, -19.033406593407)
    FSData_348 = pts
End Function

Private Function FSData_349() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-27.171450195874, 13.181513352159, -19.155194575718)
    pts(1) = Array(-27.171450195874, 13.181513352159, -15.982168061645)
    pts(2) = Array(-27.999189738264, 13.583069388053, -16.228681318681)
    pts(3) = Array(-27.999189738264, 13.583069388053, -18.908681318681)
    FSData_349 = pts
End Function

Private Function FSData_350() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.223240974871, 16.608073781254, -19.030469300993)
    pts(1) = Array(-25.223240974871, 16.608073781254, -15.857442786919)
    pts(2) = Array(-25.991631097284, 17.114015101743, -16.103956043956)
    pts(3) = Array(-25.991631097284, 17.114015101743, -18.783956043956)
    FSData_350 = pts
End Function

Private Function FSData_351() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.845347368437, 19.751711409785, -18.905744026267)
    pts(1) = Array(-22.845347368437, 19.751711409785, -15.732717512194)
    pts(2) = Array(-23.541298347873, 20.353419174587, -15.979230769231)
    pts(3) = Array(-23.541298347873, 20.353419174587, -18.659230769231)
    FSData_351 = pts
End Function

Private Function FSData_352() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.078277404621, 22.558873563702, -18.781018751542)
    pts(1) = Array(-20.078277404621, 22.558873563702, -15.607992237469)
    pts(2) = Array(-20.689933537477, 23.246097526570, -15.854505494505)
    pts(3) = Array(-20.689933537477, 23.246097526570, -18.534505494505)
    FSData_352 = pts
End Function

Private Function FSData_353() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.969168830630, 24.981739515045, -18.656293476817)
    pts(1) = Array(-16.969168830630, 24.981739515045, -15.483266962744)
    pts(2) = Array(-17.486110397656, 25.742772639344, -15.729780219780)
    pts(3) = Array(-17.486110397656, 25.742772639344, -18.409780219780)
    FSData_353 = pts
End Function

Private Function FSData_354() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.570986109208, 26.979035120324, -18.531568202092)
    pts(1) = Array(-13.570986109208, 26.979035120324, -15.358541688018)
    pts(2) = Array(-13.984406878098, 27.800913011407, -15.605054945055)
    pts(3) = Array(-13.984406878098, 27.800913011407, -18.285054945055)
    FSData_354 = pts
End Function

Private Function FSData_355() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.941618155411, 28.516735936148, -18.406842927366)
    pts(1) = Array(-9.941618155411, 28.516735936148, -15.233816413293)
    pts(2) = Array(-10.244475397231, 29.385457693143, -15.480329670330)
    pts(3) = Array(-10.244475397231, 29.385457693143, -18.160329670330)
    FSData_355 = pts
End Function

Private Function FSData_356() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.142892184079, 29.568646834355, -18.282117652641)
    pts(1) = Array(-6.142892184079, 29.568646834355, -15.109091138568)
    pts(2) = Array(-6.330026647965, 30.469413559110, -15.355604395604)
    pts(3) = Array(-6.330026647965, 30.469413559110, -18.035604395604)
    FSData_356 = pts
End Function

Private Function FSData_357() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.239520467360, 30.116848242741, -18.157392377916)
    pts(1) = Array(-2.239520467360, 30.116848242741, -14.984365863842)
    pts(2) = Array(-2.307744269677, 31.034315142851, -15.230879120879)
    pts(3) = Array(-2.307744269677, 31.034315142851, -17.910879120879)
    FSData_357 = pts
End Function

Private Function FSData_358() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.702002055460, 30.152001409578, -18.032667103190)
    pts(1) = Array(1.702002055460, 30.152001409578, -14.859640589117)
    pts(2) = Array(1.753851124699, 31.070539200863, -15.106153846154)
    pts(3) = Array(1.753851124699, 31.070539200863, -17.786153846154)
    FSData_358 = pts
End Function

Private Function FSData_359() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.614530536320, 29.673507491645, -17.907941828465)
    pts(1) = Array(5.614530536320, 29.673507491645, -14.734915314392)
    pts(2) = Array(5.785569214910, 30.577468647020, -14.981428571429)
    pts(3) = Array(5.785569214910, 30.577468647020, -17.661428571429)
    FSData_359 = pts
End Function

Private Function FSData_360() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(9.431414048106, 28.689517755675, -17.783216553740)
    pts(1) = Array(9.431414048106, 28.689517755675, -14.610190039667)
    pts(2) = Array(9.718728648247, 29.563503064788, -14.856703296703)
    pts(3) = Array(9.718728648247, 29.563503064788, -17.536703296703)
    FSData_360 = pts
End Function

Private Function FSData_361() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.087631000307, 27.216794719434, -17.658491279015)
    pts(1) = Array(13.087631000307, 27.216794719434, -14.485464764941)
    pts(2) = Array(13.486327044025, 28.045915618172, -14.731978021978)
    pts(3) = Array(13.486327044025, 28.045915618172, -17.411978021978)
    FSData_361 = pts
End Function

Private Function FSData_362() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(16.520896798502, 25.280426597928, -17.533766004289)
    pts(1) = Array(16.520896798502, 25.280426597928, -14.360739490216)
    pts(2) = Array(17.024182396337, 26.050558798925, -14.607252747253)
    pts(3) = Array(17.024182396337, 26.050558798925, -17.287252747253)
    FSData_362 = pts
End Function

Private Function FSData_363() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.672724878415, 22.913399919222, -17.409040729564)
    pts(1) = Array(19.672724878415, 22.913399919222, -14.236014215491)
    pts(2) = Array(20.272026431002, 23.611424022721, -14.482527472527)
    pts(3) = Array(20.272026431002, 23.611424022721, -17.162527472527)
    FSData_363 = pts
End Function

Private Function FSData_364() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.489423039553, 20.156037590459, -17.284315454839)
    pts(1) = Array(22.489423039553, 20.156037590459, -14.111288940766)
    pts(2) = Array(23.174531291089, 20.770062576659, -14.357802197802)
    pts(3) = Array(23.174531291089, 20.770062576659, -17.037802197802)
    FSData_364 = pts
End Function

Private Function FSData_365() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.923008105626, 17.055311986794, -17.159590180114)
    pts(1) = Array(24.923008105626, 17.055311986794, -13.986563666040)
    pts(2) = Array(25.682252061162, 17.574877782419, -14.233076923077)
    pts(3) = Array(25.682252061162, 17.574877782419, -16.913076923077)
    FSData_365 = pts
End Function

Private Function FSData_366() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.932023330272, 13.664044764917, -17.034864905388)
    pts(1) = Array(26.932023330272, 13.664044764917, -13.861838391315)
    pts(2) = Array(27.752469074109, 14.080300433252, -14.108351648352)
    pts(3) = Array(27.752469074109, 14.080300433252, -16.788351648352)
    FSData_366 = pts
End Function

Private Function FSData_367() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.482244623376, 10.040007032574, -16.910139630663)
    pts(1) = Array(28.482244623376, 10.040007032574, -13.737113116590)
    pts(2) = Array(29.349915651638, 10.345861551447, -13.983626373626)
    pts(3) = Array(29.349915651638, 10.345861551447, -16.663626373626)
    FSData_367 = pts
End Function

Private Function FSData_368() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.547263567230, 6.244935202916, -16.785414355938)
    pts(1) = Array(29.547263567230, 6.244935202916, -13.612387841864)
    pts(2) = Array(30.447378881199, 6.435178262077, -13.858901098901)
    pts(3) = Array(30.447378881199, 6.435178262077, -16.538901098901)
    FSData_368 = pts
End Function

Private Function FSData_369() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.108937290709, 2.343479298847, -16.660689081212)
    pts(1) = Array(30.108937290709, 2.343479298847, -13.487662567139)
    pts(2) = Array(31.026163194930, 2.414870058945, -13.734175824176)
    pts(3) = Array(31.026163194930, 2.414870058945, -16.414175824176)
    FSData_369 = pts
End Function

Private Function FSData_370() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(30.157697537739, -1.597898376706, -16.535963806487)
    pts(1) = Array(30.157697537739, -1.597898376706, -13.362937292414)
    pts(2) = Array(31.076408853458, -1.646576075599, -13.609450549451)
    pts(3) = Array(31.076408853458, -1.646576075599, -16.289450549451)
    FSData_370 = pts
End Function

Private Function FSData_371() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.692713664992, -5.512055443192, -16.411238531762)
    pts(1) = Array(29.692713664992, -5.512055443192, -13.238212017689)
    pts(2) = Array(30.597259909091, -5.679972363978, -13.484725274725)
    pts(3) = Array(30.597259909091, -5.679972363978, -16.164725274725)
    FSData_371 = pts
End Function

Private Function FSData_372() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(28.721906792114, -9.332313230123, -16.286513257037)
    pts(1) = Array(28.721906792114, -9.332313230123, -13.113486742963)
    pts(2) = Array(29.596878787105, -9.616608864948, -13.360000000000)
    pts(3) = Array(29.596878787105, -9.616608864948, -16.040000000000)
    FSData_372 = pts
End Function

Private Function FocusPlacement(ByVal label As String, ByVal angleDeg As Double) As Variant
    Dim m(15) As Double, known As Boolean, moving As Boolean, ring As Boolean
    Dim theta As Double, dz As Double
    m(0) = 1#: m(4) = 1#: m(8) = 1#: m(12) = 1#
    dz = -angleDeg / 60000#
    If label = "01_OpticalCartridge" Or InStr(label, "_01_OpticalCartridge-") > 0 Then
        known = True
        moving = True
    End If
    If label = "02_FrontSpacer" Or InStr(label, "_02_FrontSpacer-") > 0 Then
        known = True
        moving = True
    End If
    If label = "03_FrontRetainer" Or InStr(label, "_03_FrontRetainer-") > 0 Then
        known = True
        moving = True
    End If
    If label = "04_RearSpacer34" Or InStr(label, "_04_RearSpacer34-") > 0 Then
        known = True
        moving = True
    End If
    If label = "05_RearSpacer45" Or InStr(label, "_05_RearSpacer45-") > 0 Then
        known = True
        moving = True
    End If
    If label = "06_RearRetainer" Or InStr(label, "_06_RearRetainer-") > 0 Then
        known = True
        moving = True
    End If
    If label = "07_IntegratedHousing" Or InStr(label, "_07_IntegratedHousing-") > 0 Then
        known = True
    End If
    If label = "08_ShortFocusRing" Or InStr(label, "_08_ShortFocusRing-") > 0 Then
        known = True
        moving = True
        ring = True
    End If
    If label = "09_CouplingKeeper" Or InStr(label, "_09_CouplingKeeper-") > 0 Then
        known = True
        moving = True
    End If
    If label = "10_FrontThrustWasher" Or InStr(label, "_10_FrontThrustWasher-") > 0 Then
        known = True
        moving = True
    End If
    If label = "11_RearThrustWasher" Or InStr(label, "_11_RearThrustWasher-") > 0 Then
        known = True
        moving = True
    End If
    If label = "12_EFMountInterfaceEnvelope" Or InStr(label, "_12_EFMountInterfaceEnvelope-") > 0 Then
        known = True
    End If
    If label = "Lens1_RevisedBlank_Reference" Or InStr(label, "_Lens1_RevisedBlank_Reference-") > 0 Then
        known = True
        moving = True
    End If
    If label = "Lens2_RevisedBlank_Reference" Or InStr(label, "_Lens2_RevisedBlank_Reference-") > 0 Then
        known = True
        moving = True
    End If
    If label = "Lens3_RevisedBlank_Reference" Or InStr(label, "_Lens3_RevisedBlank_Reference-") > 0 Then
        known = True
        moving = True
    End If
    If label = "Lens4_RevisedBlank_Reference" Or InStr(label, "_Lens4_RevisedBlank_Reference-") > 0 Then
        known = True
        moving = True
    End If
    If label = "Lens5_RevisedBlank_Reference" Or InStr(label, "_Lens5_RevisedBlank_Reference-") > 0 Then
        known = True
        moving = True
    End If
    If label = "Lens6_RevisedBlank_Reference" Or InStr(label, "_Lens6_RevisedBlank_Reference-") > 0 Then
        known = True
        moving = True
    End If
    If label = "Lens7_RevisedBlank_Reference" Or InStr(label, "_Lens7_RevisedBlank_Reference-") > 0 Then
        known = True
        moving = True
    End If
    If label = "21_GuideDogScrew" Or InStr(label, "_21_GuideDogScrew-") > 0 Then
        known = True
        m(0) = -0.000000000000000
        m(1) = 1.000000000000000
        m(2) = 0.000000000000000
        m(3) = 0.000000000000000
        m(4) = 0.000000000000000
        m(5) = 1.000000000000000
        m(6) = 1.000000000000000
        m(7) = 0.000000000000000
        m(8) = 0.000000000000000
        m(11) = -0.020000000000
    End If
    If label = "22_GuideDogScrew" Or InStr(label, "_22_GuideDogScrew-") > 0 Then
        known = True
        m(0) = -0.000000000000000
        m(1) = -1.000000000000000
        m(2) = 0.000000000000000
        m(3) = 0.000000000000000
        m(4) = 0.000000000000000
        m(5) = 1.000000000000000
        m(6) = -1.000000000000000
        m(7) = 0.000000000000000
        m(8) = 0.000000000000000
        m(11) = -0.020000000000
    End If
    If label = "31_MountScrew" Or InStr(label, "_31_MountScrew-") > 0 Then
        known = True
        m(9) = 0.029500000000: m(10) = 0.000000000000
    End If
    If label = "32_MountScrew" Or InStr(label, "_32_MountScrew-") > 0 Then
        known = True
        m(9) = -0.014750000000: m(10) = 0.025547749412
    End If
    If label = "33_MountScrew" Or InStr(label, "_33_MountScrew-") > 0 Then
        known = True
        m(9) = -0.014750000000: m(10) = -0.025547749412
    End If
    If label = "40_KeeperLockScrew" Or InStr(label, "_40_KeeperLockScrew-") > 0 Then
        known = True
        moving = True
        m(0) = -1.000000000000000
        m(1) = 0.000000000000000
        m(2) = 0.000000000000000
        m(3) = 0.000000000000000
        m(4) = 0.000000000000000
        m(5) = 1.000000000000000
        m(6) = 0.000000000000000
        m(7) = 1.000000000000000
        m(8) = 0.000000000000000
        m(11) = -0.042800000000
    End If
    If Not known Then Err.Raise vbObjectError + 810, , "Unrecognized focus component."
    If moving Then m(11) = m(11) + dz
    If ring Then
        theta = angleDeg * 0.0174532925199433
        m(0) = Cos(theta): m(1) = Sin(theta): m(3) = -Sin(theta): m(4) = Cos(theta)
    End If
    FocusPlacement = m
End Function

Private Function FocusSection(ByVal model As SldWorks.ModelDoc2, ByVal vertices As Variant, ByVal sectionName As String) As SldWorks.Feature
    Dim feat As SldWorks.Feature, found As SldWorks.Feature, seg As SldWorks.SketchSegment
    Dim p As Variant, q As Variant, i As Long, j As Long, count As Long
    model.ClearSelection2 True
    model.SketchManager.Insert3DSketch True
    model.SketchManager.AddToDB = True
    For i = LBound(vertices) To UBound(vertices)
        j = i + 1: If j > UBound(vertices) Then j = LBound(vertices)
        p = vertices(i): q = vertices(j)
        Set seg = model.SketchManager.CreateLine(CDbl(p(0)) / 1000#, CDbl(p(1)) / 1000#, CDbl(p(2)) / 1000#, CDbl(q(0)) / 1000#, CDbl(q(1)) / 1000#, CDbl(q(2)) / 1000#)
        If seg Is Nothing Then Err.Raise vbObjectError + 801, , "Focus section line failed: " & sectionName
    Next i
    model.SketchManager.AddToDB = False
    model.SketchManager.Insert3DSketch True
    Set feat = model.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = "3DProfileFeature" Then
            If Left$(feat.Name, 3) <> "FS_" Then
                Set found = feat: count = count + 1
            End If
        End If
        Set feat = feat.GetNextFeature
    Loop
    If count <> 1 Then Err.Raise vbObjectError + 802, , "Expected one new focus section: " & sectionName
    found.Name = "FS_" & sectionName
    model.ClearSelection2 True
    Set FocusSection = found
End Function

Private Sub FocusLoftCut(ByVal model As SldWorks.ModelDoc2, ByVal cutName As String, ByVal sectionIDs As Variant)
    Dim sections() As SldWorks.Feature, section As SldWorks.Feature, cutFeat As SldWorks.Feature
    Dim vertices As Variant, i As Long
    currentStage = "Focus cut " & cutName
    ReDim sections(UBound(sectionIDs))
    For i = 0 To UBound(sectionIDs)
        vertices = FocusSectionData(CLng(sectionIDs(i)))
        Set sections(i) = FocusSection(model, vertices, cutName & "_" & CStr(i))
    Next i
    model.ClearSelection2 True
    For i = 0 To UBound(sections)
        Set section = sections(i)
        If section.Select2(True, 1) = False Then Err.Raise vbObjectError + 803, , "Focus loft profile selection failed: " & cutName
    Next i
    Set cutFeat = model.FeatureManager.InsertCutBlend(False, False, False, 1#, 0, 0, False, 0#, 0#, 0, False, True)
    If cutFeat Is Nothing Then Err.Raise vbObjectError + 804, , "Focus loft cut failed: " & cutName
    cutFeat.Name = cutName
    model.ClearSelection2 True
    model.EditRebuild3
    buildLog.WriteLine "Created focus loft cut: " & cutName
End Sub

Public Sub FocusNear()
    FocusSetAngle 270#
End Sub

Public Sub FocusMiddle()
    FocusSetAngle 135#
End Sub

Public Sub FocusReference()
    FocusSetAngle 0#
End Sub

Public Sub FocusSetAngle(ByVal angleDeg As Double)
    On Error GoTo FocusFailed
    Dim model As SldWorks.ModelDoc2, assembly As SldWorks.AssemblyDoc
    Dim component As SldWorks.Component2, components As Variant, matrix As Variant
    Dim math As SldWorks.MathUtility, move As SldWorks.MathTransform, i As Long
    If angleDeg < 0# Or angleDeg > 270# Then Err.Raise vbObjectError + 805, , "Angle must be 0 to 270 degrees."
    Set app = Application.SldWorks
    Set model = app.ActiveDoc
    If model Is Nothing Then Err.Raise vbObjectError + 806, , "Open the focus assembly."
    If model.GetType <> 2 Then Err.Raise vbObjectError + 807, , "Activate the generated focus assembly."
    Set assembly = model: Set math = app.GetMathUtility
    components = assembly.GetComponents(False)
    For i = 0 To UBound(components)
        Set component = components(i)
        If FocusKnownComponent(component.Name2) Then
            model.ClearSelection2 True
            If component.Select4(False, Nothing, False) = False Then Err.Raise vbObjectError + 808, , "Cannot select focus component."
            assembly.UnfixComponent
            matrix = FocusPlacement(component.Name2, angleDeg)
            Set move = math.CreateTransform(matrix)
            component.Transform2 = move
            assembly.FixComponent
        End If
    Next i
    model.ClearSelection2 True: model.EditRebuild3: model.ViewZoomtofit2
    MsgBox "Focus angle=" & CStr(angleDeg) & " deg; whole optical shift=" & CStr(-angleDeg / 60#) & " mm." & vbCrLf & "Coordinate demonstration; no assembly mates. EF interface is an envelope.", vbInformation
    Exit Sub
FocusFailed:
    MsgBox "Focus pose stopped: " & Err.Description, vbExclamation
End Sub

Private Function FocusKnownComponent(ByVal label As String) As Boolean
    FocusKnownComponent = (InStr(label, "DG50_") > 0)
End Function
