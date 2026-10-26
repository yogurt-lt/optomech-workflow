Attribute VB_Name = "F35_Electric_v02"
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
    Dim loaded(23) As SldWorks.ModelDoc2
    Set app = Application.SldWorks
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = fs.GetParentFolderName(app.GetCurrentMacroPathName)
    If Len(folder) = 0 Then Err.Raise vbObjectError + 400, , "Save the macro before running."
    output = folder & "\Optomech_" & Format$(Now, "yyyymmdd_hhnnss")
    runID = "F35_" & Format$(Now, "yyyymmdd_hhnnss")
    fs.CreateFolder output
    Set log = fs.CreateTextFile(output & "\build_log.txt", False, True)
    Set buildLog = log
    log.WriteLine "Optomech workflow concept; SOLIDWORKS " & app.RevisionNumber
    names = Array("01_FixedGuideBarrel", "02_L1Retainer", "03_L2MovingCell", "04_L2Retainer", "09_L7Retainer", "10_GearedCamRing", "11_CamKeeper", "12_FrontThrustWasher", "13_RearThrustWasher", "14_InputPinion", "05_RearRectangularSpacer", "Lens1_Optical_Reference", "Lens2_Optical_Reference", "Lens3_RevisedBlank_Reference", "Lens4_RevisedBlank_Reference", "Lens5_RevisedBlank_Reference", "Lens6_RevisedBlank_Reference", "Lens7_Optical_Reference", "21_GuideAxle", "31_CamRoller", "22_GuideAxle", "32_CamRoller", "23_GuideAxle", "33_CamRoller")
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
    If InStr(key, "01_fixedguidebarrel") > 0 Then
        values(0) = 0.06: values(1) = 0.51: values(2) = 0.54
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "02_l1retainer") > 0 Then
        values(0) = 0.26: values(1) = 0.29: values(2) = 0.33
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "03_l2movingcell") > 0 Then
        values(0) = 0.2: values(1) = 0.36: values(2) = 0.69
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "04_l2retainer") > 0 Then
        values(0) = 0.26: values(1) = 0.29: values(2) = 0.33
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "09_l7retainer") > 0 Then
        values(0) = 0.26: values(1) = 0.29: values(2) = 0.33
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "10_gearedcamring") > 0 Then
        values(0) = 0.83: values(1) = 0.55: values(2) = 0.19
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "11_camkeeper") > 0 Then
        values(0) = 0.26: values(1) = 0.29: values(2) = 0.33
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "12_frontthrustwasher") > 0 Then
        values(0) = 0.72: values(1) = 0.76: values(2) = 0.8
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "13_rearthrustwasher") > 0 Then
        values(0) = 0.72: values(1) = 0.76: values(2) = 0.8
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "14_inputpinion") > 0 Then
        values(0) = 0.28: values(1) = 0.62: values(2) = 0.38
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "05_rearrectangularspacer") > 0 Then
        values(0) = 0.28: values(1) = 0.62: values(2) = 0.38
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "lens1_optical_reference") > 0 Then
        values(0) = 0.24: values(1) = 0.63: values(2) = 0.86
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "lens2_optical_reference") > 0 Then
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
    If InStr(key, "lens7_optical_reference") > 0 Then
        values(0) = 0.24: values(1) = 0.63: values(2) = 0.86
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "21_guideaxle") > 0 Then
        values(0) = 0.72: values(1) = 0.76: values(2) = 0.8
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "31_camroller") > 0 Then
        values(0) = 0.91: values(1) = 0.35: values(2) = 0.13
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "22_guideaxle") > 0 Then
        values(0) = 0.72: values(1) = 0.76: values(2) = 0.8
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "32_camroller") > 0 Then
        values(0) = 0.91: values(1) = 0.35: values(2) = 0.13
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "23_guideaxle") > 0 Then
        values(0) = 0.72: values(1) = 0.76: values(2) = 0.8
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "33_camroller") > 0 Then
        values(0) = 0.91: values(1) = 0.35: values(2) = 0.13
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
        Case Else: Err.Raise vbObjectError + 418, , "Unknown profile."
    End Select
End Function

Private Function Profile_0() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-2.000000000000, 28.000000000000)
    pts(1) = Array(14.800000000000, 28.000000000000)
    pts(2) = Array(14.800000000000, 26.800000000000)
    pts(3) = Array(38.500000000000, 26.800000000000)
    pts(4) = Array(38.500000000000, 17.000000000000)
    pts(5) = Array(69.000000000000, 17.000000000000)
    pts(6) = Array(69.000000000000, 14.500000000000)
    pts(7) = Array(62.480320557798, 14.500000000000)
    pts(8) = Array(62.480320557798, 13.420000000000)
    pts(9) = Array(58.013404737590, 13.420000000000)
    pts(10) = Array(58.013404737590, 11.280000000000)
    pts(11) = Array(42.128372085100, 11.280000000000)
    pts(12) = Array(42.128372085100, 10.700000000000)
    pts(13) = Array(35.000000000000, 10.700000000000)
    pts(14) = Array(35.000000000000, 22.580000000000)
    pts(15) = Array(10.943587831425, 22.580000000000)
    pts(16) = Array(10.943587831425, 23.789000000000)
    pts(17) = Array(3.484452102183, 23.789000000000)
    pts(18) = Array(3.484452102183, 25.500000000000)
    pts(19) = Array(-2.000000000000, 25.500000000000)
    Profile_0 = pts
End Function

Private Function Profile_1() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(0.484452102183, 25.400000000000)
    pts(1) = Array(3.484452102183, 25.400000000000)
    pts(2) = Array(3.484452102183, 23.300000000000)
    pts(3) = Array(0.484452102183, 23.300000000000)
    Profile_1 = pts
End Function

Private Function Profile_2() As Variant
    Dim pts(7) As Variant
    pts(0) = Array(18.000000000000, 22.500000000000)
    pts(1) = Array(28.400000000000, 22.500000000000)
    pts(2) = Array(28.400000000000, 18.600000000000)
    pts(3) = Array(25.666404746767, 18.600000000000)
    pts(4) = Array(25.666404746767, 19.070000000000)
    pts(5) = Array(22.137311495393, 19.070000000000)
    pts(6) = Array(22.137311495393, 20.100000000000)
    pts(7) = Array(18.000000000000, 20.100000000000)
    Profile_2 = pts
End Function

Private Function Profile_3() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.637311495393, 20.000000000000)
    pts(1) = Array(22.137311495393, 20.000000000000)
    pts(2) = Array(22.137311495393, 18.600000000000)
    pts(3) = Array(19.637311495393, 18.600000000000)
    Profile_3 = pts
End Function

Private Function Profile_4() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(62.480320557798, 14.400000000000)
    pts(1) = Array(65.480320557798, 14.400000000000)
    pts(2) = Array(65.480320557798, 12.950000000000)
    pts(3) = Array(62.480320557798, 12.950000000000)
    Profile_4 = pts
End Function

Private Function Profile_5() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(15.200000000000, 32.800000000000)
    pts(1) = Array(18.200000000000, 32.800000000000)
    pts(2) = Array(18.200000000000, 31.000000000000)
    pts(3) = Array(35.000000000000, 31.000000000000)
    pts(4) = Array(35.000000000000, 26.900000000000)
    pts(5) = Array(15.200000000000, 26.900000000000)
    Profile_5 = pts
End Function

Private Function Profile_6() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(35.400000000000, 31.100000000000)
    pts(1) = Array(38.400000000000, 31.100000000000)
    pts(2) = Array(38.400000000000, 26.850000000000)
    pts(3) = Array(35.400000000000, 26.850000000000)
    Profile_6 = pts
End Function

Private Function Profile_7() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(14.850000000000, 30.800000000000)
    pts(1) = Array(15.150000000000, 30.800000000000)
    pts(2) = Array(15.150000000000, 26.900000000000)
    pts(3) = Array(14.850000000000, 26.900000000000)
    Profile_7 = pts
End Function

Private Function Profile_8() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(35.050000000000, 30.800000000000)
    pts(1) = Array(35.350000000000, 30.800000000000)
    pts(2) = Array(35.350000000000, 26.900000000000)
    pts(3) = Array(35.050000000000, 26.900000000000)
    Profile_8 = pts
End Function

Private Function Profile_9() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(15.200000000000, 7.200000000000)
    pts(1) = Array(18.200000000000, 7.200000000000)
    pts(2) = Array(18.200000000000, 2.000000000000)
    pts(3) = Array(15.200000000000, 2.000000000000)
    Profile_9 = pts
End Function

Private Function Profile_10() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(58.013404737590, 13.300000000000)
    pts(1) = Array(60.320544535343, 13.300000000000)
    pts(2) = Array(60.320544535343, 10.700000000000)
    pts(3) = Array(58.013404737590, 10.700000000000)
    Profile_10 = pts
End Function

Private Function Profile_11() As Variant
    Dim pts(169) As Variant
    pts(0) = Array(0.000000000000, 0.000000000000)
    pts(1) = Array(0.000901888441, 0.379024000000)
    pts(2) = Array(0.003607615044, 0.758048000000)
    pts(3) = Array(0.008117363673, 1.137072000000)
    pts(4) = Array(0.014431440835, 1.516096000000)
    pts(5) = Array(0.022550275786, 1.895120000000)
    pts(6) = Array(0.032474420677, 2.274144000000)
    pts(7) = Array(0.044204550742, 2.653168000000)
    pts(8) = Array(0.057741464527, 3.032192000000)
    pts(9) = Array(0.073086084164, 3.411216000000)
    pts(10) = Array(0.090239455684, 3.790240000000)
    pts(11) = Array(0.109202749373, 4.169264000000)
    pts(12) = Array(0.129977260177, 4.548288000000)
    pts(13) = Array(0.152564408141, 4.927312000000)
    pts(14) = Array(0.176965738898, 5.306336000000)
    pts(15) = Array(0.203182924202, 5.685360000000)
    pts(16) = Array(0.227613940220, 6.017006000000)
    pts(17) = Array(0.257240800250, 6.396030000000)
    pts(18) = Array(0.288689029044, 6.775054000000)
    pts(19) = Array(0.321960792918, 7.154078000000)
    pts(20) = Array(0.357058387540, 7.533102000000)
    pts(21) = Array(0.393984238739, 7.912126000000)
    pts(22) = Array(0.432740903345, 8.291150000000)
    pts(23) = Array(0.473331070088, 8.670174000000)
    pts(24) = Array(0.510353725564, 9.001820000000)
    pts(25) = Array(0.554389419042, 9.380844000000)
    pts(26) = Array(0.600267088260, 9.759868000000)
    pts(27) = Array(0.647989942451, 10.138892000000)
    pts(28) = Array(0.697561327841, 10.517916000000)
    pts(29) = Array(0.748984728841, 10.896940000000)
    pts(30) = Array(0.802263769304, 11.275964000000)
    pts(31) = Array(0.857402213829, 11.654988000000)
    pts(32) = Array(0.907176712862, 11.986634000000)
    pts(33) = Array(0.965812182619, 12.365658000000)
    pts(34) = Array(1.026318676275, 12.744682000000)
    pts(35) = Array(1.088700517244, 13.123706000000)
    pts(36) = Array(1.152962176902, 13.502730000000)
    pts(37) = Array(1.219108276240, 13.881754000000)
    pts(38) = Array(1.287143587592, 14.260778000000)
    pts(39) = Array(1.357073036421, 14.639802000000)
    pts(40) = Array(1.419819078860, 14.971448000000)
    pts(41) = Array(1.493313852167, 15.350472000000)
    pts(42) = Array(1.568717793748, 15.729496000000)
    pts(43) = Array(1.646036440625, 16.108520000000)
    pts(44) = Array(1.715265390150, 16.440166000000)
    pts(45) = Array(1.796189602123, 16.819190000000)
    pts(46) = Array(1.879045353256, 17.198214000000)
    pts(47) = Array(1.963838824245, 17.577238000000)
    pts(48) = Array(2.039627636875, 17.908884000000)
    pts(49) = Array(2.128071591811, 18.287908000000)
    pts(50) = Array(2.218472002399, 18.666932000000)
    pts(51) = Array(2.310835729772, 19.045956000000)
    pts(52) = Array(2.393270056979, 19.377602000000)
    pts(53) = Array(2.489334132691, 19.756626000000)
    pts(54) = Array(2.575017496960, 20.088272000000)
    pts(55) = Array(2.674807913501, 20.467296000000)
    pts(56) = Array(2.763763608060, 20.798942000000)
    pts(57) = Array(2.867307624626, 21.177966000000)
    pts(58) = Array(2.972865156828, 21.556990000000)
    pts(59) = Array(3.080444532717, 21.936014000000)
    pts(60) = Array(3.176241734674, 22.267660000000)
    pts(61) = Array(3.287635237106, 22.646684000000)
    pts(62) = Array(3.386783381433, 22.978330000000)
    pts(63) = Array(3.502022417863, 23.357354000000)
    pts(64) = Array(3.604549614033, 23.689000000000)
    pts(65) = Array(10.943587831425, 23.689000000000)
    pts(66) = Array(10.943587831425, 20.209000000000)
    pts(67) = Array(10.811576036280, 20.088272000000)
    pts(68) = Array(10.658357214488, 19.946138000000)
    pts(69) = Array(10.507461979185, 19.804004000000)
    pts(70) = Array(10.358834809296, 19.661870000000)
    pts(71) = Array(10.212422734351, 19.519736000000)
    pts(72) = Array(10.068175174182, 19.377602000000)
    pts(73) = Array(9.926043791438, 19.235468000000)
    pts(74) = Array(9.785982355669, 19.093334000000)
    pts(75) = Array(9.647946617900, 18.951200000000)
    pts(76) = Array(9.511894194715, 18.809066000000)
    pts(77) = Array(9.377784460984, 18.666932000000)
    pts(78) = Array(9.245578450457, 18.524798000000)
    pts(79) = Array(9.115238763533, 18.382664000000)
    pts(80) = Array(8.986729481570, 18.240530000000)
    pts(81) = Array(8.860016087192, 18.098396000000)
    pts(82) = Array(8.693801314046, 17.908884000000)
    pts(83) = Array(8.571151415267, 17.766750000000)
    pts(84) = Array(8.450191550402, 17.624616000000)
    pts(85) = Array(8.330892478232, 17.482482000000)
    pts(86) = Array(8.174361939232, 17.292970000000)
    pts(87) = Array(8.058830184226, 17.150836000000)
    pts(88) = Array(7.944869112364, 17.008702000000)
    pts(89) = Array(7.832453714244, 16.866568000000)
    pts(90) = Array(7.684929271293, 16.677056000000)
    pts(91) = Array(7.576027925877, 16.534922000000)
    pts(92) = Array(7.433106375553, 16.345410000000)
    pts(93) = Array(7.327597590420, 16.203276000000)
    pts(94) = Array(7.189123139563, 16.013764000000)
    pts(95) = Array(7.086894492237, 15.871630000000)
    pts(96) = Array(6.952722832150, 15.682118000000)
    pts(97) = Array(6.853670085281, 15.539984000000)
    pts(98) = Array(6.723667288164, 15.350472000000)
    pts(99) = Array(6.627693614750, 15.208338000000)
    pts(100) = Array(6.501735165521, 15.018826000000)
    pts(101) = Array(6.408750467900, 14.876692000000)
    pts(102) = Array(6.286720411665, 14.687180000000)
    pts(103) = Array(6.196640720804, 14.545046000000)
    pts(104) = Array(6.078430909560, 14.355534000000)
    pts(105) = Array(5.991177853379, 14.213400000000)
    pts(106) = Array(5.876687278234, 14.023888000000)
    pts(107) = Array(5.792187609555, 13.881754000000)
    pts(108) = Array(5.681321806928, 13.692242000000)
    pts(109) = Array(5.572487225904, 13.502730000000)
    pts(110) = Array(5.465656045680, 13.313218000000)
    pts(111) = Array(5.386831223890, 13.171084000000)
    pts(112) = Array(5.283442046264, 12.981572000000)
    pts(113) = Array(5.181984807310, 12.792060000000)
    pts(114) = Array(5.082435229992, 12.602548000000)
    pts(115) = Array(5.009010826333, 12.460414000000)
    pts(116) = Array(4.912743777608, 12.270902000000)
    pts(117) = Array(4.818321907125, 12.081390000000)
    pts(118) = Array(4.725723949123, 11.891878000000)
    pts(119) = Array(4.634929364850, 11.702366000000)
    pts(120) = Array(4.545918313919, 11.512854000000)
    pts(121) = Array(4.458671627225, 11.323342000000)
    pts(122) = Array(4.373170781318, 11.133830000000)
    pts(123) = Array(4.310180062992, 10.991696000000)
    pts(124) = Array(4.227691722462, 10.802184000000)
    pts(125) = Array(4.146901418876, 10.612672000000)
    pts(126) = Array(4.067792844122, 10.423160000000)
    pts(127) = Array(3.990350231306, 10.233648000000)
    pts(128) = Array(3.914558335924, 10.044136000000)
    pts(129) = Array(3.840402417989, 9.854624000000)
    pts(130) = Array(3.767868225053, 9.665112000000)
    pts(131) = Array(3.662077649138, 9.380844000000)
    pts(132) = Array(3.559860451819, 9.096576000000)
    pts(133) = Array(3.461174231266, 8.812308000000)
    pts(134) = Array(3.365978684595, 8.528040000000)
    pts(135) = Array(3.274235513705, 8.243772000000)
    pts(136) = Array(3.185908337821, 7.959504000000)
    pts(137) = Array(3.100962612241, 7.675236000000)
    pts(138) = Array(3.059747481447, 7.533102000000)
    pts(139) = Array(3.006089400260, 7.343590000000)
    pts(140) = Array(2.928359933827, 7.059322000000)
    pts(141) = Array(2.853913813050, 6.775054000000)
    pts(142) = Array(2.782723388528, 6.490786000000)
    pts(143) = Array(2.748340833207, 6.348652000000)
    pts(144) = Array(2.703747924871, 6.159140000000)
    pts(145) = Array(2.639523725683, 5.874872000000)
    pts(146) = Array(2.608604685014, 5.732738000000)
    pts(147) = Array(2.568610316793, 5.543226000000)
    pts(148) = Array(2.511243950011, 5.258958000000)
    pts(149) = Array(2.483736612907, 5.116824000000)
    pts(150) = Array(2.448274121621, 4.927312000000)
    pts(151) = Array(2.397670662573, 4.643044000000)
    pts(152) = Array(2.373529800393, 4.500910000000)
    pts(153) = Array(2.342541150643, 4.311398000000)
    pts(154) = Array(2.298618314804, 4.027130000000)
    pts(155) = Array(2.277804890810, 3.884996000000)
    pts(156) = Array(2.251240177650, 3.695484000000)
    pts(157) = Array(2.213927622811, 3.411216000000)
    pts(158) = Array(2.196408464099, 3.269082000000)
    pts(159) = Array(2.174225496420, 3.079570000000)
    pts(160) = Array(2.143464240763, 2.795302000000)
    pts(161) = Array(2.111375722199, 2.463656000000)
    pts(162) = Array(2.087117666686, 2.179388000000)
    pts(163) = Array(2.062592813486, 1.847742000000)
    pts(164) = Array(2.044800359304, 1.563474000000)
    pts(165) = Array(2.027801295297, 1.231828000000)
    pts(166) = Array(2.016447049132, 0.947560000000)
    pts(167) = Array(2.006947669307, 0.615914000000)
    pts(168) = Array(2.002014231000, 0.331646000000)
    pts(169) = Array(2.000000000000, 0.000000000000)
    Profile_11 = pts
End Function

Private Function Profile_12() As Variant
    Dim pts(81) As Variant
    pts(0) = Array(22.014517220987, 0.000000000000)
    pts(1) = Array(22.015042029090, 1.216000000000)
    pts(2) = Array(22.016551365982, 2.394000000000)
    pts(3) = Array(22.022525178391, 4.750000000000)
    pts(4) = Array(22.027150082814, 5.966000000000)
    pts(5) = Array(22.032631413736, 7.144000000000)
    pts(6) = Array(22.046549323730, 9.500000000000)
    pts(7) = Array(22.055274482220, 10.716000000000)
    pts(8) = Array(22.064727994129, 11.894000000000)
    pts(9) = Array(22.086590476439, 14.250000000000)
    pts(10) = Array(22.099416186614, 15.466000000000)
    pts(11) = Array(22.112842201982, 16.644000000000)
    pts(12) = Array(22.142650002396, 19.000000000000)
    pts(13) = Array(25.503713564385, 19.000000000000)
    pts(14) = Array(25.643064279120, 18.658000000000)
    pts(15) = Array(25.794453601027, 18.278000000000)
    pts(16) = Array(25.942250968616, 17.898000000000)
    pts(17) = Array(26.086490468133, 17.518000000000)
    pts(18) = Array(26.185358826341, 17.252000000000)
    pts(19) = Array(26.282510869751, 16.986000000000)
    pts(20) = Array(26.377957238071, 16.720000000000)
    pts(21) = Array(26.471708298738, 16.454000000000)
    pts(22) = Array(26.563774153404, 16.188000000000)
    pts(23) = Array(26.654164644185, 15.922000000000)
    pts(24) = Array(26.742889359670, 15.656000000000)
    pts(25) = Array(26.829957640711, 15.390000000000)
    pts(26) = Array(26.915378585999, 15.124000000000)
    pts(27) = Array(26.999161057429, 14.858000000000)
    pts(28) = Array(27.081313685274, 14.592000000000)
    pts(29) = Array(27.161844873163, 14.326000000000)
    pts(30) = Array(27.240762802880, 14.060000000000)
    pts(31) = Array(27.318075438988, 13.794000000000)
    pts(32) = Array(27.393790533278, 13.528000000000)
    pts(33) = Array(27.467915629066, 13.262000000000)
    pts(34) = Array(27.540458065320, 12.996000000000)
    pts(35) = Array(27.611424980646, 12.730000000000)
    pts(36) = Array(27.680823317125, 12.464000000000)
    pts(37) = Array(27.748659824010, 12.198000000000)
    pts(38) = Array(27.814941061287, 11.932000000000)
    pts(39) = Array(27.879673403105, 11.666000000000)
    pts(40) = Array(27.942863041082, 11.400000000000)
    pts(41) = Array(28.013198449517, 11.096000000000)
    pts(42) = Array(28.073102312936, 10.830000000000)
    pts(43) = Array(28.131481778378, 10.564000000000)
    pts(44) = Array(28.188342313862, 10.298000000000)
    pts(45) = Array(28.251472660841, 9.994000000000)
    pts(46) = Array(28.305095984269, 9.728000000000)
    pts(47) = Array(28.357216485296, 9.462000000000)
    pts(48) = Array(28.407838955073, 9.196000000000)
    pts(49) = Array(28.463864819602, 8.892000000000)
    pts(50) = Array(28.511292627909, 8.626000000000)
    pts(51) = Array(28.557236449783, 8.360000000000)
    pts(52) = Array(28.601700439233, 8.094000000000)
    pts(53) = Array(28.650709518940, 7.790000000000)
    pts(54) = Array(28.692015738609, 7.524000000000)
    pts(55) = Array(28.731854222815, 7.258000000000)
    pts(56) = Array(28.770228521957, 6.992000000000)
    pts(57) = Array(28.812296350929, 6.688000000000)
    pts(58) = Array(28.847544411038, 6.422000000000)
    pts(59) = Array(28.881338531677, 6.156000000000)
    pts(60) = Array(28.913681687627, 5.890000000000)
    pts(61) = Array(28.948872239278, 5.586000000000)
    pts(62) = Array(28.978115574731, 5.320000000000)
    pts(63) = Array(29.009770357490, 5.016000000000)
    pts(64) = Array(29.035925600119, 4.750000000000)
    pts(65) = Array(29.064057277609, 4.446000000000)
    pts(66) = Array(29.087134866576, 4.180000000000)
    pts(67) = Array(29.108779432333, 3.914000000000)
    pts(68) = Array(29.128992844644, 3.648000000000)
    pts(69) = Array(29.150343684893, 3.344000000000)
    pts(70) = Array(29.167496048410, 3.078000000000)
    pts(71) = Array(29.183222314352, 2.812000000000)
    pts(72) = Array(29.197523832806, 2.546000000000)
    pts(73) = Array(29.212125404001, 2.242000000000)
    pts(74) = Array(29.223377863379, 1.976000000000)
    pts(75) = Array(29.233209014866, 1.710000000000)
    pts(76) = Array(29.241619698879, 1.444000000000)
    pts(77) = Array(29.249493477321, 1.140000000000)
    pts(78) = Array(29.254862565739, 0.874000000000)
    pts(79) = Array(29.259261603762, 0.570000000000)
    pts(80) = Array(29.261591200371, 0.304000000000)
    pts(81) = Array(29.262517220987, 0.000000000000)
    Profile_12 = pts
End Function

Private Function Profile_13() As Variant
    Dim pts(1013) As Variant
    pts(0) = Array(38.703517220987, 0.000000000000)
    pts(1) = Array(38.703531309043, 0.022400000000)
    pts(2) = Array(38.703573573278, 0.044800000000)
    pts(3) = Array(38.703644013892, 0.067200000000)
    pts(4) = Array(38.703742631219, 0.089600000000)
    pts(5) = Array(38.703869425728, 0.112000000000)
    pts(6) = Array(38.704024398021, 0.134400000000)
    pts(7) = Array(38.704207548833, 0.156800000000)
    pts(8) = Array(38.704418879034, 0.179200000000)
    pts(9) = Array(38.704658389627, 0.201600000000)
    pts(10) = Array(38.704926081749, 0.224000000000)
    pts(11) = Array(38.705221956671, 0.246400000000)
    pts(12) = Array(38.705546015799, 0.268800000000)
    pts(13) = Array(38.705898260670, 0.291200000000)
    pts(14) = Array(38.706278692959, 0.313600000000)
    pts(15) = Array(38.706687314470, 0.336000000000)
    pts(16) = Array(38.707124127147, 0.358400000000)
    pts(17) = Array(38.707589133063, 0.380800000000)
    pts(18) = Array(38.708082334427, 0.403200000000)
    pts(19) = Array(38.708603733583, 0.425600000000)
    pts(20) = Array(38.709153333009, 0.448000000000)
    pts(21) = Array(38.709731135316, 0.470400000000)
    pts(22) = Array(38.710337143251, 0.492800000000)
    pts(23) = Array(38.710971359694, 0.515200000000)
    pts(24) = Array(38.711633787661, 0.537600000000)
    pts(25) = Array(38.712324430302, 0.560000000000)
    pts(26) = Array(38.713043290900, 0.582400000000)
    pts(27) = Array(38.713790372876, 0.604800000000)
    pts(28) = Array(38.714565679782, 0.627200000000)
    pts(29) = Array(38.715369215309, 0.649600000000)
    pts(30) = Array(38.716200983280, 0.672000000000)
    pts(31) = Array(38.717060987654, 0.694400000000)
    pts(32) = Array(38.717949232526, 0.716800000000)
    pts(33) = Array(38.718865722124, 0.739200000000)
    pts(34) = Array(38.719810460814, 0.761600000000)
    pts(35) = Array(38.720783453096, 0.784000000000)
    pts(36) = Array(38.721784703605, 0.806400000000)
    pts(37) = Array(38.722814217114, 0.828800000000)
    pts(38) = Array(38.723871998530, 0.851200000000)
    pts(39) = Array(38.724958052896, 0.873600000000)
    pts(40) = Array(38.726072385392, 0.896000000000)
    pts(41) = Array(38.727215001333, 0.918400000000)
    pts(42) = Array(38.728385906171, 0.940800000000)
    pts(43) = Array(38.729585105494, 0.963200000000)
    pts(44) = Array(38.730812605028, 0.985600000000)
    pts(45) = Array(38.732068410633, 1.008000000000)
    pts(46) = Array(38.733352528309, 1.030400000000)
    pts(47) = Array(38.734664964190, 1.052800000000)
    pts(48) = Array(38.736005724549, 1.075200000000)
    pts(49) = Array(38.737374815797, 1.097600000000)
    pts(50) = Array(38.738772244479, 1.120000000000)
    pts(51) = Array(38.740198017282, 1.142400000000)
    pts(52) = Array(38.741652141027, 1.164800000000)
    pts(53) = Array(38.743134622675, 1.187200000000)
    pts(54) = Array(38.744645469326, 1.209600000000)
    pts(55) = Array(38.746184688214, 1.232000000000)
    pts(56) = Array(38.747752286717, 1.254400000000)
    pts(57) = Array(38.749348272348, 1.276800000000)
    pts(58) = Array(38.750972652760, 1.299200000000)
    pts(59) = Array(38.752625435744, 1.321600000000)
    pts(60) = Array(38.754306629231, 1.344000000000)
    pts(61) = Array(38.756016241291, 1.366400000000)
    pts(62) = Array(38.757754280135, 1.388800000000)
    pts(63) = Array(38.759520754111, 1.411200000000)
    pts(64) = Array(38.761315671709, 1.433600000000)
    pts(65) = Array(38.763139041559, 1.456000000000)
    pts(66) = Array(38.764990872431, 1.478400000000)
    pts(67) = Array(38.766871173235, 1.500800000000)
    pts(68) = Array(38.768779953023, 1.523200000000)
    pts(69) = Array(38.770717220987, 1.545600000000)
    pts(70) = Array(38.772682986461, 1.568000000000)
    pts(71) = Array(38.774677258920, 1.590400000000)
    pts(72) = Array(38.776700047980, 1.612800000000)
    pts(73) = Array(38.778751363400, 1.635200000000)
    pts(74) = Array(38.780831215082, 1.657600000000)
    pts(75) = Array(38.782939613069, 1.680000000000)
    pts(76) = Array(38.785076567545, 1.702400000000)
    pts(77) = Array(38.787242088842, 1.724800000000)
    pts(78) = Array(38.789436187429, 1.747200000000)
    pts(79) = Array(38.791658873924, 1.769600000000)
    pts(80) = Array(38.793910159084, 1.792000000000)
    pts(81) = Array(38.796190053814, 1.814400000000)
    pts(82) = Array(38.798498569160, 1.836800000000)
    pts(83) = Array(38.800835716313, 1.859200000000)
    pts(84) = Array(38.803201506612, 1.881600000000)
    pts(85) = Array(38.805595951536, 1.904000000000)
    pts(86) = Array(38.808019062714, 1.926400000000)
    pts(87) = Array(38.810470851916, 1.948800000000)
    pts(88) = Array(38.812951331063, 1.971200000000)
    pts(89) = Array(38.815460512217, 1.993600000000)
    pts(90) = Array(38.817998407591, 2.016000000000)
    pts(91) = Array(38.820565029542, 2.038400000000)
    pts(92) = Array(38.823160390575, 2.060800000000)
    pts(93) = Array(38.825784503342, 2.083200000000)
    pts(94) = Array(38.828437380643, 2.105600000000)
    pts(95) = Array(38.831119035426, 2.128000000000)
    pts(96) = Array(38.833829480788, 2.150400000000)
    pts(97) = Array(38.836568729973, 2.172800000000)
    pts(98) = Array(38.839336796376, 2.195200000000)
    pts(99) = Array(38.842133693541, 2.217600000000)
    pts(100) = Array(38.844959435159, 2.240000000000)
    pts(101) = Array(38.847814035075, 2.262400000000)
    pts(102) = Array(38.850697507282, 2.284800000000)
    pts(103) = Array(38.853609865924, 2.307200000000)
    pts(104) = Array(38.856551125296, 2.329600000000)
    pts(105) = Array(38.859521299846, 2.352000000000)
    pts(106) = Array(38.862520404171, 2.374400000000)
    pts(107) = Array(38.865548453022, 2.396800000000)
    pts(108) = Array(38.868605461303, 2.419200000000)
    pts(109) = Array(38.871691444070, 2.441600000000)
    pts(110) = Array(38.874806416532, 2.464000000000)
    pts(111) = Array(38.877950394052, 2.486400000000)
    pts(112) = Array(38.881123392147, 2.508800000000)
    pts(113) = Array(38.884325426490, 2.531200000000)
    pts(114) = Array(38.887556512906, 2.553600000000)
    pts(115) = Array(38.890816667377, 2.576000000000)
    pts(116) = Array(38.894105906041, 2.598400000000)
    pts(117) = Array(38.897424245191, 2.620800000000)
    pts(118) = Array(38.900771701277, 2.643200000000)
    pts(119) = Array(38.904148290906, 2.665600000000)
    pts(120) = Array(38.907554030842, 2.688000000000)
    pts(121) = Array(38.910988938007, 2.710400000000)
    pts(122) = Array(38.914453029481, 2.732800000000)
    pts(123) = Array(38.917946322503, 2.755200000000)
    pts(124) = Array(38.921468834471, 2.777600000000)
    pts(125) = Array(38.925020582941, 2.800000000000)
    pts(126) = Array(38.928601585633, 2.822400000000)
    pts(127) = Array(38.932211860422, 2.844800000000)
    pts(128) = Array(38.935851425348, 2.867200000000)
    pts(129) = Array(38.939520298611, 2.889600000000)
    pts(130) = Array(38.943218498573, 2.912000000000)
    pts(131) = Array(38.946946043757, 2.934400000000)
    pts(132) = Array(38.950702952850, 2.956800000000)
    pts(133) = Array(38.954489244704, 2.979200000000)
    pts(134) = Array(38.958304938331, 3.001600000000)
    pts(135) = Array(38.962150052910, 3.024000000000)
    pts(136) = Array(38.966024607783, 3.046400000000)
    pts(137) = Array(38.969928622460, 3.068800000000)
    pts(138) = Array(38.973862116613, 3.091200000000)
    pts(139) = Array(38.977825110083, 3.113600000000)
    pts(140) = Array(38.981817622877, 3.136000000000)
    pts(141) = Array(38.985839675169, 3.158400000000)
    pts(142) = Array(38.989891287302, 3.180800000000)
    pts(143) = Array(38.993972479786, 3.203200000000)
    pts(144) = Array(38.998083273300, 3.225600000000)
    pts(145) = Array(39.002223688694, 3.248000000000)
    pts(146) = Array(39.006393746987, 3.270400000000)
    pts(147) = Array(39.010593469368, 3.292800000000)
    pts(148) = Array(39.014822877198, 3.315200000000)
    pts(149) = Array(39.019081992009, 3.337600000000)
    pts(150) = Array(39.023370835507, 3.360000000000)
    pts(151) = Array(39.027689429568, 3.382400000000)
    pts(152) = Array(39.032037796245, 3.404800000000)
    pts(153) = Array(39.036415957761, 3.427200000000)
    pts(154) = Array(39.040823936518, 3.449600000000)
    pts(155) = Array(39.045261755090, 3.472000000000)
    pts(156) = Array(39.049729436227, 3.494400000000)
    pts(157) = Array(39.054227002858, 3.516800000000)
    pts(158) = Array(39.058754478086, 3.539200000000)
    pts(159) = Array(39.063311885194, 3.561600000000)
    pts(160) = Array(39.067899247642, 3.584000000000)
    pts(161) = Array(39.072516589070, 3.606400000000)
    pts(162) = Array(39.077163933297, 3.628800000000)
    pts(163) = Array(39.081841304322, 3.651200000000)
    pts(164) = Array(39.086548726324, 3.673600000000)
    pts(165) = Array(39.091286223667, 3.696000000000)
    pts(166) = Array(39.096053820894, 3.718400000000)
    pts(167) = Array(39.100851542731, 3.740800000000)
    pts(168) = Array(39.105679414091, 3.763200000000)
    pts(169) = Array(39.110537460066, 3.785600000000)
    pts(170) = Array(39.115425705937, 3.808000000000)
    pts(171) = Array(39.120344177170, 3.830400000000)
    pts(172) = Array(39.125292899416, 3.852800000000)
    pts(173) = Array(39.130271898514, 3.875200000000)
    pts(174) = Array(39.135281200491, 3.897600000000)
    pts(175) = Array(39.140320831562, 3.920000000000)
    pts(176) = Array(39.145390818131, 3.942400000000)
    pts(177) = Array(39.150491186792, 3.964800000000)
    pts(178) = Array(39.155621964332, 3.987200000000)
    pts(179) = Array(39.160783177726, 4.009600000000)
    pts(180) = Array(39.165974854143, 4.032000000000)
    pts(181) = Array(39.171197020945, 4.054400000000)
    pts(182) = Array(39.176449705688, 4.076800000000)
    pts(183) = Array(39.181732936121, 4.099200000000)
    pts(184) = Array(39.187046740191, 4.121600000000)
    pts(185) = Array(39.192391146038, 4.144000000000)
    pts(186) = Array(39.197766182001, 4.166400000000)
    pts(187) = Array(39.203171876617, 4.188800000000)
    pts(188) = Array(39.208608258619, 4.211200000000)
    pts(189) = Array(39.214075356943, 4.233600000000)
    pts(190) = Array(39.219573200722, 4.256000000000)
    pts(191) = Array(39.225101819291, 4.278400000000)
    pts(192) = Array(39.230661242188, 4.300800000000)
    pts(193) = Array(39.236251499153, 4.323200000000)
    pts(194) = Array(39.241872620128, 4.345600000000)
    pts(195) = Array(39.247524635261, 4.368000000000)
    pts(196) = Array(39.253207574907, 4.390400000000)
    pts(197) = Array(39.258921469623, 4.412800000000)
    pts(198) = Array(39.264666350176, 4.435200000000)
    pts(199) = Array(39.270442247541, 4.457600000000)
    pts(200) = Array(39.276249192900, 4.480000000000)
    pts(201) = Array(39.282087217647, 4.502400000000)
    pts(202) = Array(39.287956353385, 4.524800000000)
    pts(203) = Array(39.293856631929, 4.547200000000)
    pts(204) = Array(39.299788085308, 4.569600000000)
    pts(205) = Array(39.305750745762, 4.592000000000)
    pts(206) = Array(39.311744645747, 4.614400000000)
    pts(207) = Array(39.317769817934, 4.636800000000)
    pts(208) = Array(39.323826295211, 4.659200000000)
    pts(209) = Array(39.329914110684, 4.681600000000)
    pts(210) = Array(39.336033297674, 4.704000000000)
    pts(211) = Array(39.342183889725, 4.726400000000)
    pts(212) = Array(39.348365920600, 4.748800000000)
    pts(213) = Array(39.354579424282, 4.771200000000)
    pts(214) = Array(39.360824434980, 4.793600000000)
    pts(215) = Array(39.367100987123, 4.816000000000)
    pts(216) = Array(39.373409115365, 4.838400000000)
    pts(217) = Array(39.379748854587, 4.860800000000)
    pts(218) = Array(39.386120239895, 4.883200000000)
    pts(219) = Array(39.392523306624, 4.905600000000)
    pts(220) = Array(39.398958090337, 4.928000000000)
    pts(221) = Array(39.405424626827, 4.950400000000)
    pts(222) = Array(39.411922952117, 4.972800000000)
    pts(223) = Array(39.418453102463, 4.995200000000)
    pts(224) = Array(39.425015114354, 5.017600000000)
    pts(225) = Array(39.431609024513, 5.040000000000)
    pts(226) = Array(39.438234869899, 5.062400000000)
    pts(227) = Array(39.444892687706, 5.084800000000)
    pts(228) = Array(39.451582515368, 5.107200000000)
    pts(229) = Array(39.458304390556, 5.129600000000)
    pts(230) = Array(39.465058351181, 5.152000000000)
    pts(231) = Array(39.471844435397, 5.174400000000)
    pts(232) = Array(39.478662681598, 5.196800000000)
    pts(233) = Array(39.485513128424, 5.219200000000)
    pts(234) = Array(39.492395814757, 5.241600000000)
    pts(235) = Array(39.499310779729, 5.264000000000)
    pts(236) = Array(39.506258062715, 5.286400000000)
    pts(237) = Array(39.513237703342, 5.308800000000)
    pts(238) = Array(39.520249741485, 5.331200000000)
    pts(239) = Array(39.527294217270, 5.353600000000)
    pts(240) = Array(39.534371171077, 5.376000000000)
    pts(241) = Array(39.541480643537, 5.398400000000)
    pts(242) = Array(39.548622675539, 5.420800000000)
    pts(243) = Array(39.555797308226, 5.443200000000)
    pts(244) = Array(39.563004583000, 5.465600000000)
    pts(245) = Array(39.570244541522, 5.488000000000)
    pts(246) = Array(39.577517225712, 5.510400000000)
    pts(247) = Array(39.584822677752, 5.532800000000)
    pts(248) = Array(39.592160940088, 5.555200000000)
    pts(249) = Array(39.599532055431, 5.577600000000)
    pts(250) = Array(39.606936066756, 5.600000000000)
    pts(251) = Array(39.614373017306, 5.622400000000)
    pts(252) = Array(39.621842950594, 5.644800000000)
    pts(253) = Array(39.629345910401, 5.667200000000)
    pts(254) = Array(39.636881940781, 5.689600000000)
    pts(255) = Array(39.644451086060, 5.712000000000)
    pts(256) = Array(39.652053390841, 5.734400000000)
    pts(257) = Array(39.659688900001, 5.756800000000)
    pts(258) = Array(39.667357658694, 5.779200000000)
    pts(259) = Array(39.675059712354, 5.801600000000)
    pts(260) = Array(39.682795106697, 5.824000000000)
    pts(261) = Array(39.690563887719, 5.846400000000)
    pts(262) = Array(39.698366101700, 5.868800000000)
    pts(263) = Array(39.706201795206, 5.891200000000)
    pts(264) = Array(39.714071015090, 5.913600000000)
    pts(265) = Array(39.721973808494, 5.936000000000)
    pts(266) = Array(39.729910222848, 5.958400000000)
    pts(267) = Array(39.737880305876, 5.980800000000)
    pts(268) = Array(39.745884105593, 6.003200000000)
    pts(269) = Array(39.753921670312, 6.025600000000)
    pts(270) = Array(39.761993048641, 6.048000000000)
    pts(271) = Array(39.770098289487, 6.070400000000)
    pts(272) = Array(39.778237442056, 6.092800000000)
    pts(273) = Array(39.786410555857, 6.115200000000)
    pts(274) = Array(39.794617680702, 6.137600000000)
    pts(275) = Array(39.802858866709, 6.160000000000)
    pts(276) = Array(39.811134164303, 6.182400000000)
    pts(277) = Array(39.819443624217, 6.204800000000)
    pts(278) = Array(39.827787297496, 6.227200000000)
    pts(279) = Array(39.836165235497, 6.249600000000)
    pts(280) = Array(39.844577489891, 6.272000000000)
    pts(281) = Array(39.853024112666, 6.294400000000)
    pts(282) = Array(39.861505156127, 6.316800000000)
    pts(283) = Array(39.870020672900, 6.339200000000)
    pts(284) = Array(39.878570715933, 6.361600000000)
    pts(285) = Array(39.887155338497, 6.384000000000)
    pts(286) = Array(39.895774594189, 6.406400000000)
    pts(287) = Array(39.904428536934, 6.428800000000)
    pts(288) = Array(39.913117220987, 6.451200000000)
    pts(289) = Array(39.921840700934, 6.473600000000)
    pts(290) = Array(39.930599031695, 6.496000000000)
    pts(291) = Array(39.939392268526, 6.518400000000)
    pts(292) = Array(39.948220467020, 6.540800000000)
    pts(293) = Array(39.957083683110, 6.563200000000)
    pts(294) = Array(39.965981973072, 6.585600000000)
    pts(295) = Array(39.974915393526, 6.608000000000)
    pts(296) = Array(39.983884001436, 6.630400000000)
    pts(297) = Array(39.992887854118, 6.652800000000)
    pts(298) = Array(40.001927009235, 6.675200000000)
    pts(299) = Array(40.011001524804, 6.697600000000)
    pts(300) = Array(40.020111459197, 6.720000000000)
    pts(301) = Array(40.029256871145, 6.742400000000)
    pts(302) = Array(40.038437819734, 6.764800000000)
    pts(303) = Array(40.047654364416, 6.787200000000)
    pts(304) = Array(40.056906565004, 6.809600000000)
    pts(305) = Array(40.066194481679, 6.832000000000)
    pts(306) = Array(40.075518174991, 6.854400000000)
    pts(307) = Array(40.084877705858, 6.876800000000)
    pts(308) = Array(40.094273135575, 6.899200000000)
    pts(309) = Array(40.103704525811, 6.921600000000)
    pts(310) = Array(40.113171938612, 6.944000000000)
    pts(311) = Array(40.122675436407, 6.966400000000)
    pts(312) = Array(40.132215082007, 6.988800000000)
    pts(313) = Array(40.141790938609, 7.011200000000)
    pts(314) = Array(40.151403069796, 7.033600000000)
    pts(315) = Array(40.161051539546, 7.056000000000)
    pts(316) = Array(40.170736412225, 7.078400000000)
    pts(317) = Array(40.180457752600, 7.100800000000)
    pts(318) = Array(40.190215625833, 7.123200000000)
    pts(319) = Array(40.200010097487, 7.145600000000)
    pts(320) = Array(40.209841233531, 7.168000000000)
    pts(321) = Array(40.219709100340, 7.190400000000)
    pts(322) = Array(40.229613764696, 7.212800000000)
    pts(323) = Array(40.239555293795, 7.235200000000)
    pts(324) = Array(40.249533755247, 7.257600000000)
    pts(325) = Array(40.259549217081, 7.280000000000)
    pts(326) = Array(40.269601747743, 7.302400000000)
    pts(327) = Array(40.279691416106, 7.324800000000)
    pts(328) = Array(40.289818291467, 7.347200000000)
    pts(329) = Array(40.299982443552, 7.369600000000)
    pts(330) = Array(40.310183942519, 7.392000000000)
    pts(331) = Array(40.320422858962, 7.414400000000)
    pts(332) = Array(40.330699263912, 7.436800000000)
    pts(333) = Array(40.341013228842, 7.459200000000)
    pts(334) = Array(40.351364825666, 7.481600000000)
    pts(335) = Array(40.358698205578, 7.497419679825)
    pts(336) = Array(40.359270104667, 7.498651685025)
    pts(337) = Array(40.359782651745, 7.499755623002)
    pts(338) = Array(40.361754126750, 7.504000000000)
    pts(339) = Array(40.372181204905, 7.526400000000)
    pts(340) = Array(40.382646133400, 7.548800000000)
    pts(341) = Array(40.393148985958, 7.571200000000)
    pts(342) = Array(40.403689836762, 7.593600000000)
    pts(343) = Array(40.414268760458, 7.616000000000)
    pts(344) = Array(40.424885832160, 7.638400000000)
    pts(345) = Array(40.435541127449, 7.660800000000)
    pts(346) = Array(40.446234722382, 7.683200000000)
    pts(347) = Array(40.456966693489, 7.705600000000)
    pts(348) = Array(40.467737117782, 7.728000000000)
    pts(349) = Array(40.478546072757, 7.750400000000)
    pts(350) = Array(40.489393636393, 7.772800000000)
    pts(351) = Array(40.500279887165, 7.795200000000)
    pts(352) = Array(40.511204904035, 7.817600000000)
    pts(353) = Array(40.522168766469, 7.840000000000)
    pts(354) = Array(40.533171554430, 7.862400000000)
    pts(355) = Array(40.544213348386, 7.884800000000)
    pts(356) = Array(40.555294229314, 7.907200000000)
    pts(357) = Array(40.566414278704, 7.929600000000)
    pts(358) = Array(40.577573578562, 7.952000000000)
    pts(359) = Array(40.588772211411, 7.974400000000)
    pts(360) = Array(40.600010260300, 7.996800000000)
    pts(361) = Array(40.611287808806, 8.019200000000)
    pts(362) = Array(40.622604941036, 8.041600000000)
    pts(363) = Array(40.633961741632, 8.064000000000)
    pts(364) = Array(40.645358295778, 8.086400000000)
    pts(365) = Array(40.656794689199, 8.108800000000)
    pts(366) = Array(40.668271008168, 8.131200000000)
    pts(367) = Array(40.679787339512, 8.153600000000)
    pts(368) = Array(40.691343770611, 8.176000000000)
    pts(369) = Array(40.702940389406, 8.198400000000)
    pts(370) = Array(40.714577284405, 8.220800000000)
    pts(371) = Array(40.726254544681, 8.243200000000)
    pts(372) = Array(40.737972259882, 8.265600000000)
    pts(373) = Array(40.749730520235, 8.288000000000)
    pts(374) = Array(40.761529416548, 8.310400000000)
    pts(375) = Array(40.773369040213, 8.332800000000)
    pts(376) = Array(40.785249483218, 8.355200000000)
    pts(377) = Array(40.797170838143, 8.377600000000)
    pts(378) = Array(40.809133198171, 8.400000000000)
    pts(379) = Array(40.821136657088, 8.422400000000)
    pts(380) = Array(40.833181309291, 8.444800000000)
    pts(381) = Array(40.845267249791, 8.467200000000)
    pts(382) = Array(40.857394574221, 8.489600000000)
    pts(383) = Array(40.869563378836, 8.512000000000)
    pts(384) = Array(40.881773760521, 8.534400000000)
    pts(385) = Array(40.894025816796, 8.556800000000)
    pts(386) = Array(40.906319645820, 8.579200000000)
    pts(387) = Array(40.918655346399, 8.601600000000)
    pts(388) = Array(40.931033017985, 8.624000000000)
    pts(389) = Array(40.943452760689, 8.646400000000)
    pts(390) = Array(40.955914675280, 8.668800000000)
    pts(391) = Array(40.968418863195, 8.691200000000)
    pts(392) = Array(40.980965426540, 8.713600000000)
    pts(393) = Array(40.993554468100, 8.736000000000)
    pts(394) = Array(41.006186091340, 8.758400000000)
    pts(395) = Array(41.018860400416, 8.780800000000)
    pts(396) = Array(41.031577500174, 8.803200000000)
    pts(397) = Array(41.044337496163, 8.825600000000)
    pts(398) = Array(41.057140494636, 8.848000000000)
    pts(399) = Array(41.069986602554, 8.870400000000)
    pts(400) = Array(41.082875927600, 8.892800000000)
    pts(401) = Array(41.095808578177, 8.915200000000)
    pts(402) = Array(41.108784663418, 8.937600000000)
    pts(403) = Array(41.121804293190, 8.960000000000)
    pts(404) = Array(41.134867578103, 8.982400000000)
    pts(405) = Array(41.147974629514, 9.004800000000)
    pts(406) = Array(41.161125559534, 9.027200000000)
    pts(407) = Array(41.174320481035, 9.049600000000)
    pts(408) = Array(41.187559507655, 9.072000000000)
    pts(409) = Array(41.200842753806, 9.094400000000)
    pts(410) = Array(41.214170334680, 9.116800000000)
    pts(411) = Array(41.227542366258, 9.139200000000)
    pts(412) = Array(41.240958965311, 9.161600000000)
    pts(413) = Array(41.254420249414, 9.184000000000)
    pts(414) = Array(41.267926336947, 9.206400000000)
    pts(415) = Array(41.281477347106, 9.228800000000)
    pts(416) = Array(41.295073399910, 9.251200000000)
    pts(417) = Array(41.308714616204, 9.273600000000)
    pts(418) = Array(41.322401117672, 9.296000000000)
    pts(419) = Array(41.336133026839, 9.318400000000)
    pts(420) = Array(41.349910467084, 9.340800000000)
    pts(421) = Array(41.363733562643, 9.363200000000)
    pts(422) = Array(41.377602438618, 9.385600000000)
    pts(423) = Array(41.391517220987, 9.408000000000)
    pts(424) = Array(41.405478036608, 9.430400000000)
    pts(425) = Array(41.419485013229, 9.452800000000)
    pts(426) = Array(41.433538279498, 9.475200000000)
    pts(427) = Array(41.447637964966, 9.497600000000)
    pts(428) = Array(41.461784200099, 9.520000000000)
    pts(429) = Array(41.475977116289, 9.542400000000)
    pts(430) = Array(41.490216845853, 9.564800000000)
    pts(431) = Array(41.504503522053, 9.587200000000)
    pts(432) = Array(41.518837279096, 9.609600000000)
    pts(433) = Array(41.533218252146, 9.632000000000)
    pts(434) = Array(41.547646577334, 9.654400000000)
    pts(435) = Array(41.562122391764, 9.676800000000)
    pts(436) = Array(41.576645833526, 9.699200000000)
    pts(437) = Array(41.591217041701, 9.721600000000)
    pts(438) = Array(41.605836156373, 9.744000000000)
    pts(439) = Array(41.620503318637, 9.766400000000)
    pts(440) = Array(41.635218670608, 9.788800000000)
    pts(441) = Array(41.649982355436, 9.811200000000)
    pts(442) = Array(41.664794517306, 9.833600000000)
    pts(443) = Array(41.679655301457, 9.856000000000)
    pts(444) = Array(41.694564854188, 9.878400000000)
    pts(445) = Array(41.709523322867, 9.900800000000)
    pts(446) = Array(41.724530855943, 9.923200000000)
    pts(447) = Array(41.739587602958, 9.945600000000)
    pts(448) = Array(41.754693714555, 9.968000000000)
    pts(449) = Array(41.769849342487, 9.990400000000)
    pts(450) = Array(41.785054639634, 10.012800000000)
    pts(451) = Array(41.800309760007, 10.035200000000)
    pts(452) = Array(41.815614858764, 10.057600000000)
    pts(453) = Array(41.830970092218, 10.080000000000)
    pts(454) = Array(41.846375617852, 10.102400000000)
    pts(455) = Array(41.861831594326, 10.124800000000)
    pts(456) = Array(41.877338181492, 10.147200000000)
    pts(457) = Array(41.892895540405, 10.169600000000)
    pts(458) = Array(41.908503833333, 10.192000000000)
    pts(459) = Array(41.924163223774, 10.214400000000)
    pts(460) = Array(41.939873876464, 10.236800000000)
    pts(461) = Array(41.955635957388, 10.259200000000)
    pts(462) = Array(41.971449633799, 10.281600000000)
    pts(463) = Array(41.987315074225, 10.304000000000)
    pts(464) = Array(42.003232448486, 10.326400000000)
    pts(465) = Array(42.019201927702, 10.348800000000)
    pts(466) = Array(42.035223684311, 10.371200000000)
    pts(467) = Array(42.051297892083, 10.393600000000)
    pts(468) = Array(42.067424726130, 10.416000000000)
    pts(469) = Array(42.083604362919, 10.438400000000)
    pts(470) = Array(42.099836980294, 10.460800000000)
    pts(471) = Array(42.116122757479, 10.483200000000)
    pts(472) = Array(42.128372085100, 10.500000000000)
    pts(473) = Array(42.128372085100, 10.505600000000)
    pts(474) = Array(42.128372085100, 10.528000000000)
    pts(475) = Array(42.128372085100, 10.550000000000)
    pts(476) = Array(42.128372085100, 10.550400000000)
    pts(477) = Array(42.128372085100, 10.572800000000)
    pts(478) = Array(42.128372085100, 10.595200000000)
    pts(479) = Array(42.128372085100, 10.617600000000)
    pts(480) = Array(42.128372085100, 10.640000000000)
    pts(481) = Array(42.128372085100, 10.662400000000)
    pts(482) = Array(42.128372085100, 10.684800000000)
    pts(483) = Array(42.128372085100, 10.700000000000)
    pts(484) = Array(42.128372085100, 10.707200000000)
    pts(485) = Array(42.128372085100, 10.729600000000)
    pts(486) = Array(42.128372085100, 10.752000000000)
    pts(487) = Array(42.128372085100, 10.774400000000)
    pts(488) = Array(42.128372085100, 10.796800000000)
    pts(489) = Array(42.128372085100, 10.819200000000)
    pts(490) = Array(42.128372085100, 10.841600000000)
    pts(491) = Array(42.128372085100, 10.864000000000)
    pts(492) = Array(42.128372085100, 10.886400000000)
    pts(493) = Array(42.128372085100, 10.908800000000)
    pts(494) = Array(42.128372085100, 10.931200000000)
    pts(495) = Array(42.128372085100, 10.953600000000)
    pts(496) = Array(42.128372085100, 10.976000000000)
    pts(497) = Array(42.128372085100, 10.998400000000)
    pts(498) = Array(42.128372085100, 11.020800000000)
    pts(499) = Array(42.128372085100, 11.043200000000)
    pts(500) = Array(42.128372085100, 11.065600000000)
    pts(501) = Array(42.128372085100, 11.088000000000)
    pts(502) = Array(42.128372085100, 11.110400000000)
    pts(503) = Array(42.128372085100, 11.132800000000)
    pts(504) = Array(42.128372085100, 11.155200000000)
    pts(505) = Array(42.128372085100, 11.177600000000)
    pts(506) = Array(42.128372085100, 11.200000000000)
    pts(507) = Array(43.127834759587, 11.200000000000)
    pts(508) = Array(43.127834759587, 11.177600000000)
    pts(509) = Array(43.127834759587, 11.155200000000)
    pts(510) = Array(43.127834759587, 11.132800000000)
    pts(511) = Array(43.127834759587, 11.110400000000)
    pts(512) = Array(43.127834759587, 11.088000000000)
    pts(513) = Array(43.127834759587, 11.065600000000)
    pts(514) = Array(43.127834759587, 11.043200000000)
    pts(515) = Array(43.127834759587, 11.020800000000)
    pts(516) = Array(43.127834759587, 10.998400000000)
    pts(517) = Array(43.127834759587, 10.976000000000)
    pts(518) = Array(43.127834759587, 10.953600000000)
    pts(519) = Array(43.127834759587, 10.931200000000)
    pts(520) = Array(43.127834759587, 10.908800000000)
    pts(521) = Array(43.127834759587, 10.886400000000)
    pts(522) = Array(43.127834759587, 10.864000000000)
    pts(523) = Array(43.127834759587, 10.841600000000)
    pts(524) = Array(43.127834759587, 10.819200000000)
    pts(525) = Array(43.127834759587, 10.796800000000)
    pts(526) = Array(43.127834759587, 10.774400000000)
    pts(527) = Array(43.127834759587, 10.752000000000)
    pts(528) = Array(43.127834759587, 10.729600000000)
    pts(529) = Array(43.127834759587, 10.707200000000)
    pts(530) = Array(43.127834759587, 10.700000000000)
    pts(531) = Array(43.127834759587, 10.684800000000)
    pts(532) = Array(43.127834759587, 10.662400000000)
    pts(533) = Array(43.127834759587, 10.640000000000)
    pts(534) = Array(43.127834759587, 10.617600000000)
    pts(535) = Array(43.127834759587, 10.595200000000)
    pts(536) = Array(43.127834759587, 10.572800000000)
    pts(537) = Array(43.127834759587, 10.550400000000)
    pts(538) = Array(43.127834759587, 10.550000000000)
    pts(539) = Array(43.127834759587, 10.528000000000)
    pts(540) = Array(43.127834759587, 10.505600000000)
    pts(541) = Array(43.127834759587, 10.500000000000)
    pts(542) = Array(43.133857136335, 10.483200000000)
    pts(543) = Array(43.141870049656, 10.460800000000)
    pts(544) = Array(43.149863637705, 10.438400000000)
    pts(545) = Array(43.157837916303, 10.416000000000)
    pts(546) = Array(43.165792901214, 10.393600000000)
    pts(547) = Array(43.173728608147, 10.371200000000)
    pts(548) = Array(43.181645052757, 10.348800000000)
    pts(549) = Array(43.189542250643, 10.326400000000)
    pts(550) = Array(43.197420217350, 10.304000000000)
    pts(551) = Array(43.205278968367, 10.281600000000)
    pts(552) = Array(43.213118519131, 10.259200000000)
    pts(553) = Array(43.220938885023, 10.236800000000)
    pts(554) = Array(43.228740081371, 10.214400000000)
    pts(555) = Array(43.236522123449, 10.192000000000)
    pts(556) = Array(43.244285026476, 10.169600000000)
    pts(557) = Array(43.252028805619, 10.147200000000)
    pts(558) = Array(43.259753475991, 10.124800000000)
    pts(559) = Array(43.267459052654, 10.102400000000)
    pts(560) = Array(43.275145550612, 10.080000000000)
    pts(561) = Array(43.282812984822, 10.057600000000)
    pts(562) = Array(43.290461370185, 10.035200000000)
    pts(563) = Array(43.298090721550, 10.012800000000)
    pts(564) = Array(43.305701053713, 9.990400000000)
    pts(565) = Array(43.313292381420, 9.968000000000)
    pts(566) = Array(43.320864719362, 9.945600000000)
    pts(567) = Array(43.328418082182, 9.923200000000)
    pts(568) = Array(43.335952484467, 9.900800000000)
    pts(569) = Array(43.343467940756, 9.878400000000)
    pts(570) = Array(43.350964465535, 9.856000000000)
    pts(571) = Array(43.358442073240, 9.833600000000)
    pts(572) = Array(43.365900778253, 9.811200000000)
    pts(573) = Array(43.373340594908, 9.788800000000)
    pts(574) = Array(43.380761537488, 9.766400000000)
    pts(575) = Array(43.388163620225, 9.744000000000)
    pts(576) = Array(43.395546857299, 9.721600000000)
    pts(577) = Array(43.402911262843, 9.699200000000)
    pts(578) = Array(43.410256850937, 9.676800000000)
    pts(579) = Array(43.417583635611, 9.654400000000)
    pts(580) = Array(43.424891630849, 9.632000000000)
    pts(581) = Array(43.432180850580, 9.609600000000)
    pts(582) = Array(43.439451308688, 9.587200000000)
    pts(583) = Array(43.446703019005, 9.564800000000)
    pts(584) = Array(43.453935995314, 9.542400000000)
    pts(585) = Array(43.461150251350, 9.520000000000)
    pts(586) = Array(43.468345800799, 9.497600000000)
    pts(587) = Array(43.475522657297, 9.475200000000)
    pts(588) = Array(43.482680834432, 9.452800000000)
    pts(589) = Array(43.489820345745, 9.430400000000)
    pts(590) = Array(43.496941204726, 9.408000000000)
    pts(591) = Array(43.504043424819, 9.385600000000)
    pts(592) = Array(43.511127019418, 9.363200000000)
    pts(593) = Array(43.518192001872, 9.340800000000)
    pts(594) = Array(43.525238385478, 9.318400000000)
    pts(595) = Array(43.532266183490, 9.296000000000)
    pts(596) = Array(43.539275409111, 9.273600000000)
    pts(597) = Array(43.546266075499, 9.251200000000)
    pts(598) = Array(43.553238195763, 9.228800000000)
    pts(599) = Array(43.560191782965, 9.206400000000)
    pts(600) = Array(43.567126850122, 9.184000000000)
    pts(601) = Array(43.574043410202, 9.161600000000)
    pts(602) = Array(43.580941476128, 9.139200000000)
    pts(603) = Array(43.587821060775, 9.116800000000)
    pts(604) = Array(43.594682176973, 9.094400000000)
    pts(605) = Array(43.601524837504, 9.072000000000)
    pts(606) = Array(43.608349055106, 9.049600000000)
    pts(607) = Array(43.615154842470, 9.027200000000)
    pts(608) = Array(43.621942212240, 9.004800000000)
    pts(609) = Array(43.628711177016, 8.982400000000)
    pts(610) = Array(43.635461749352, 8.960000000000)
    pts(611) = Array(43.642193941755, 8.937600000000)
    pts(612) = Array(43.648907766690, 8.915200000000)
    pts(613) = Array(43.655603236572, 8.892800000000)
    pts(614) = Array(43.662280363776, 8.870400000000)
    pts(615) = Array(43.668939160628, 8.848000000000)
    pts(616) = Array(43.675579639411, 8.825600000000)
    pts(617) = Array(43.682201812364, 8.803200000000)
    pts(618) = Array(43.688805691680, 8.780800000000)
    pts(619) = Array(43.695391289507, 8.758400000000)
    pts(620) = Array(43.701958617951, 8.736000000000)
    pts(621) = Array(43.708507689071, 8.713600000000)
    pts(622) = Array(43.715038514885, 8.691200000000)
    pts(623) = Array(43.721551107364, 8.668800000000)
    pts(624) = Array(43.728045478437, 8.646400000000)
    pts(625) = Array(43.734521639989, 8.624000000000)
    pts(626) = Array(43.740979603861, 8.601600000000)
    pts(627) = Array(43.747419381850, 8.579200000000)
    pts(628) = Array(43.753840985711, 8.556800000000)
    pts(629) = Array(43.760244427155, 8.534400000000)
    pts(630) = Array(43.766629717850, 8.512000000000)
    pts(631) = Array(43.772996869420, 8.489600000000)
    pts(632) = Array(43.779345893448, 8.467200000000)
    pts(633) = Array(43.785676801472, 8.444800000000)
    pts(634) = Array(43.791989604990, 8.422400000000)
    pts(635) = Array(43.798284315455, 8.400000000000)
    pts(636) = Array(43.804560944278, 8.377600000000)
    pts(637) = Array(43.810819502829, 8.355200000000)
    pts(638) = Array(43.817060002435, 8.332800000000)
    pts(639) = Array(43.823282454380, 8.310400000000)
    pts(640) = Array(43.829486869908, 8.288000000000)
    pts(641) = Array(43.835673260218, 8.265600000000)
    pts(642) = Array(43.841841636471, 8.243200000000)
    pts(643) = Array(43.847992009784, 8.220800000000)
    pts(644) = Array(43.854124391233, 8.198400000000)
    pts(645) = Array(43.860238791852, 8.176000000000)
    pts(646) = Array(43.866335222635, 8.153600000000)
    pts(647) = Array(43.872413694533, 8.131200000000)
    pts(648) = Array(43.878474218458, 8.108800000000)
    pts(649) = Array(43.884516805278, 8.086400000000)
    pts(650) = Array(43.890541465824, 8.064000000000)
    pts(651) = Array(43.896548210883, 8.041600000000)
    pts(652) = Array(43.902537051202, 8.019200000000)
    pts(653) = Array(43.908507997488, 7.996800000000)
    pts(654) = Array(43.914461060408, 7.974400000000)
    pts(655) = Array(43.920396250587, 7.952000000000)
    pts(656) = Array(43.926313578612, 7.929600000000)
    pts(657) = Array(43.932213055027, 7.907200000000)
    pts(658) = Array(43.938094690339, 7.884800000000)
    pts(659) = Array(43.943958495011, 7.862400000000)
    pts(660) = Array(43.949804479471, 7.840000000000)
    pts(661) = Array(43.955632654104, 7.817600000000)
    pts(662) = Array(43.961443029255, 7.795200000000)
    pts(663) = Array(43.967235615232, 7.772800000000)
    pts(664) = Array(43.973010422302, 7.750400000000)
    pts(665) = Array(43.978767460692, 7.728000000000)
    pts(666) = Array(43.984506740590, 7.705600000000)
    pts(667) = Array(43.990228272146, 7.683200000000)
    pts(668) = Array(43.995932065470, 7.660800000000)
    pts(669) = Array(44.001618130633, 7.638400000000)
    pts(670) = Array(44.007286477668, 7.616000000000)
    pts(671) = Array(44.012937116568, 7.593600000000)
    pts(672) = Array(44.018570057287, 7.571200000000)
    pts(673) = Array(44.024185309742, 7.548800000000)
    pts(674) = Array(44.029782883810, 7.526400000000)
    pts(675) = Array(44.035362789331, 7.504000000000)
    pts(676) = Array(44.036418085758, 7.499755623002)
    pts(677) = Array(44.036692458384, 7.498651685025)
    pts(678) = Array(44.036998610190, 7.497419679825)
    pts(679) = Array(44.040925036106, 7.481600000000)
    pts(680) = Array(44.046469633898, 7.459200000000)
    pts(681) = Array(44.051996592430, 7.436800000000)
    pts(682) = Array(44.057505921389, 7.414400000000)
    pts(683) = Array(44.062997630425, 7.392000000000)
    pts(684) = Array(44.068471729147, 7.369600000000)
    pts(685) = Array(44.073928227129, 7.347200000000)
    pts(686) = Array(44.079367133906, 7.324800000000)
    pts(687) = Array(44.084788458975, 7.302400000000)
    pts(688) = Array(44.090192211796, 7.280000000000)
    pts(689) = Array(44.095578401793, 7.257600000000)
    pts(690) = Array(44.100947038351, 7.235200000000)
    pts(691) = Array(44.106298130818, 7.212800000000)
    pts(692) = Array(44.111631688504, 7.190400000000)
    pts(693) = Array(44.116947720684, 7.168000000000)
    pts(694) = Array(44.122246236595, 7.145600000000)
    pts(695) = Array(44.127527245437, 7.123200000000)
    pts(696) = Array(44.132790756373, 7.100800000000)
    pts(697) = Array(44.138036778530, 7.078400000000)
    pts(698) = Array(44.143265320996, 7.056000000000)
    pts(699) = Array(44.148476392827, 7.033600000000)
    pts(700) = Array(44.153670003039, 7.011200000000)
    pts(701) = Array(44.158846160611, 6.988800000000)
    pts(702) = Array(44.164004874489, 6.966400000000)
    pts(703) = Array(44.169146153580, 6.944000000000)
    pts(704) = Array(44.174270006756, 6.921600000000)
    pts(705) = Array(44.179376442853, 6.899200000000)
    pts(706) = Array(44.184465470671, 6.876800000000)
    pts(707) = Array(44.189537098973, 6.854400000000)
    pts(708) = Array(44.194591336488, 6.832000000000)
    pts(709) = Array(44.199628191908, 6.809600000000)
    pts(710) = Array(44.204647673889, 6.787200000000)
    pts(711) = Array(44.209649791053, 6.764800000000)
    pts(712) = Array(44.214634551985, 6.742400000000)
    pts(713) = Array(44.219601965236, 6.720000000000)
    pts(714) = Array(44.224552039320, 6.697600000000)
    pts(715) = Array(44.229484782717, 6.675200000000)
    pts(716) = Array(44.234400203871, 6.652800000000)
    pts(717) = Array(44.239298311192, 6.630400000000)
    pts(718) = Array(44.244179113053, 6.608000000000)
    pts(719) = Array(44.249042617794, 6.585600000000)
    pts(720) = Array(44.253888833720, 6.563200000000)
    pts(721) = Array(44.258717769100, 6.540800000000)
    pts(722) = Array(44.263529432169, 6.518400000000)
    pts(723) = Array(44.268323831127, 6.496000000000)
    pts(724) = Array(44.273100974140, 6.473600000000)
    pts(725) = Array(44.277860869339, 6.451200000000)
    pts(726) = Array(44.282603524820, 6.428800000000)
    pts(727) = Array(44.287328948646, 6.406400000000)
    pts(728) = Array(44.292037148845, 6.384000000000)
    pts(729) = Array(44.296728133410, 6.361600000000)
    pts(730) = Array(44.301401910302, 6.339200000000)
    pts(731) = Array(44.306058487445, 6.316800000000)
    pts(732) = Array(44.310697872733, 6.294400000000)
    pts(733) = Array(44.315320074021, 6.272000000000)
    pts(734) = Array(44.319925099134, 6.249600000000)
    pts(735) = Array(44.324512955861, 6.227200000000)
    pts(736) = Array(44.329083651960, 6.204800000000)
    pts(737) = Array(44.333637195152, 6.182400000000)
    pts(738) = Array(44.338173593127, 6.160000000000)
    pts(739) = Array(44.342692853540, 6.137600000000)
    pts(740) = Array(44.347194984013, 6.115200000000)
    pts(741) = Array(44.351679992135, 6.092800000000)
    pts(742) = Array(44.356147885461, 6.070400000000)
    pts(743) = Array(44.360598671513, 6.048000000000)
    pts(744) = Array(44.365032357780, 6.025600000000)
    pts(745) = Array(44.369448951718, 6.003200000000)
    pts(746) = Array(44.373848460749, 5.980800000000)
    pts(747) = Array(44.378230892264, 5.958400000000)
    pts(748) = Array(44.382596253619, 5.936000000000)
    pts(749) = Array(44.386944552137, 5.913600000000)
    pts(750) = Array(44.391275795111, 5.891200000000)
    pts(751) = Array(44.395589989799, 5.868800000000)
    pts(752) = Array(44.399887143426, 5.846400000000)
    pts(753) = Array(44.404167263185, 5.824000000000)
    pts(754) = Array(44.408430356238, 5.801600000000)
    pts(755) = Array(44.412676429712, 5.779200000000)
    pts(756) = Array(44.416905490703, 5.756800000000)
    pts(757) = Array(44.421117546275, 5.734400000000)
    pts(758) = Array(44.425312603458, 5.712000000000)
    pts(759) = Array(44.429490669252, 5.689600000000)
    pts(760) = Array(44.433651750623, 5.667200000000)
    pts(761) = Array(44.437795854505, 5.644800000000)
    pts(762) = Array(44.441922987803, 5.622400000000)
    pts(763) = Array(44.446033157385, 5.600000000000)
    pts(764) = Array(44.450126370091, 5.577600000000)
    pts(765) = Array(44.454202632727, 5.555200000000)
    pts(766) = Array(44.458261952068, 5.532800000000)
    pts(767) = Array(44.462304334858, 5.510400000000)
    pts(768) = Array(44.466329787809, 5.488000000000)
    pts(769) = Array(44.470338317599, 5.465600000000)
    pts(770) = Array(44.474329930878, 5.443200000000)
    pts(771) = Array(44.478304634262, 5.420800000000)
    pts(772) = Array(44.482262434337, 5.398400000000)
    pts(773) = Array(44.486203337656, 5.376000000000)
    pts(774) = Array(44.490127350742, 5.353600000000)
    pts(775) = Array(44.494034480087, 5.331200000000)
    pts(776) = Array(44.497924732150, 5.308800000000)
    pts(777) = Array(44.501798113361, 5.286400000000)
    pts(778) = Array(44.505654630117, 5.264000000000)
    pts(779) = Array(44.509494288785, 5.241600000000)
    pts(780) = Array(44.513317095701, 5.219200000000)
    pts(781) = Array(44.517123057169, 5.196800000000)
    pts(782) = Array(44.520912179464, 5.174400000000)
    pts(783) = Array(44.524684468827, 5.152000000000)
    pts(784) = Array(44.528439931472, 5.129600000000)
    pts(785) = Array(44.532178573580, 5.107200000000)
    pts(786) = Array(44.535900401301, 5.084800000000)
    pts(787) = Array(44.539605420756, 5.062400000000)
    pts(788) = Array(44.543293638035, 5.040000000000)
    pts(789) = Array(44.546965059195, 5.017600000000)
    pts(790) = Array(44.550619690266, 4.995200000000)
    pts(791) = Array(44.554257537246, 4.972800000000)
    pts(792) = Array(44.557878606102, 4.950400000000)
    pts(793) = Array(44.561482902772, 4.928000000000)
    pts(794) = Array(44.565070433162, 4.905600000000)
    pts(795) = Array(44.568641203150, 4.883200000000)
    pts(796) = Array(44.572195218583, 4.860800000000)
    pts(797) = Array(44.575732485276, 4.838400000000)
    pts(798) = Array(44.579253009016, 4.816000000000)
    pts(799) = Array(44.582756795560, 4.793600000000)
    pts(800) = Array(44.586243850634, 4.771200000000)
    pts(801) = Array(44.589714179934, 4.748800000000)
    pts(802) = Array(44.593167789128, 4.726400000000)
    pts(803) = Array(44.596604683851, 4.704000000000)
    pts(804) = Array(44.600024869710, 4.681600000000)
    pts(805) = Array(44.603428352284, 4.659200000000)
    pts(806) = Array(44.606815137119, 4.636800000000)
    pts(807) = Array(44.610185229733, 4.614400000000)
    pts(808) = Array(44.613538635614, 4.592000000000)
    pts(809) = Array(44.616875360222, 4.569600000000)
    pts(810) = Array(44.620195408984, 4.547200000000)
    pts(811) = Array(44.623498787302, 4.524800000000)
    pts(812) = Array(44.626785500544, 4.502400000000)
    pts(813) = Array(44.630055554053, 4.480000000000)
    pts(814) = Array(44.633308953139, 4.457600000000)
    pts(815) = Array(44.636545703085, 4.435200000000)
    pts(816) = Array(44.639765809144, 4.412800000000)
    pts(817) = Array(44.642969276539, 4.390400000000)
    pts(818) = Array(44.646156110465, 4.368000000000)
    pts(819) = Array(44.649326316088, 4.345600000000)
    pts(820) = Array(44.652479898544, 4.323200000000)
    pts(821) = Array(44.655616862941, 4.300800000000)
    pts(822) = Array(44.658737214356, 4.278400000000)
    pts(823) = Array(44.661840957840, 4.256000000000)
    pts(824) = Array(44.664928098413, 4.233600000000)
    pts(825) = Array(44.667998641067, 4.211200000000)
    pts(826) = Array(44.671052590763, 4.188800000000)
    pts(827) = Array(44.674089952438, 4.166400000000)
    pts(828) = Array(44.677110730995, 4.144000000000)
    pts(829) = Array(44.680114931312, 4.121600000000)
    pts(830) = Array(44.683102558237, 4.099200000000)
    pts(831) = Array(44.686073616588, 4.076800000000)
    pts(832) = Array(44.689028111158, 4.054400000000)
    pts(833) = Array(44.691966046707, 4.032000000000)
    pts(834) = Array(44.694887427970, 4.009600000000)
    pts(835) = Array(44.697792259652, 3.987200000000)
    pts(836) = Array(44.700680546430, 3.964800000000)
    pts(837) = Array(44.703552292952, 3.942400000000)
    pts(838) = Array(44.706407503839, 3.920000000000)
    pts(839) = Array(44.709246183682, 3.897600000000)
    pts(840) = Array(44.712068337045, 3.875200000000)
    pts(841) = Array(44.714873968463, 3.852800000000)
    pts(842) = Array(44.717663082443, 3.830400000000)
    pts(843) = Array(44.720435683465, 3.808000000000)
    pts(844) = Array(44.723191775979, 3.785600000000)
    pts(845) = Array(44.725931364408, 3.763200000000)
    pts(846) = Array(44.728654453146, 3.740800000000)
    pts(847) = Array(44.731361046561, 3.718400000000)
    pts(848) = Array(44.734051148991, 3.696000000000)
    pts(849) = Array(44.736724764746, 3.673600000000)
    pts(850) = Array(44.739381898110, 3.651200000000)
    pts(851) = Array(44.742022553338, 3.628800000000)
    pts(852) = Array(44.744646734656, 3.606400000000)
    pts(853) = Array(44.747254446265, 3.584000000000)
    pts(854) = Array(44.749845692335, 3.561600000000)
    pts(855) = Array(44.752420477011, 3.539200000000)
    pts(856) = Array(44.754978804409, 3.516800000000)
    pts(857) = Array(44.757520678617, 3.494400000000)
    pts(858) = Array(44.760046103695, 3.472000000000)
    pts(859) = Array(44.762555083678, 3.449600000000)
    pts(860) = Array(44.765047622571, 3.427200000000)
    pts(861) = Array(44.767523724352, 3.404800000000)
    pts(862) = Array(44.769983392972, 3.382400000000)
    pts(863) = Array(44.772426632353, 3.360000000000)
    pts(864) = Array(44.774853446392, 3.337600000000)
    pts(865) = Array(44.777263838957, 3.315200000000)
    pts(866) = Array(44.779657813888, 3.292800000000)
    pts(867) = Array(44.782035375000, 3.270400000000)
    pts(868) = Array(44.784396526078, 3.248000000000)
    pts(869) = Array(44.786741270882, 3.225600000000)
    pts(870) = Array(44.789069613144, 3.203200000000)
    pts(871) = Array(44.791381556567, 3.180800000000)
    pts(872) = Array(44.793677104829, 3.158400000000)
    pts(873) = Array(44.795956261580, 3.136000000000)
    pts(874) = Array(44.798219030444, 3.113600000000)
    pts(875) = Array(44.800465415016, 3.091200000000)
    pts(876) = Array(44.802695418865, 3.068800000000)
    pts(877) = Array(44.804909045532, 3.046400000000)
    pts(878) = Array(44.807106298533, 3.024000000000)
    pts(879) = Array(44.809287181355, 3.001600000000)
    pts(880) = Array(44.811451697459, 2.979200000000)
    pts(881) = Array(44.813599850279, 2.956800000000)
    pts(882) = Array(44.815731643222, 2.934400000000)
    pts(883) = Array(44.817847079668, 2.912000000000)
    pts(884) = Array(44.819946162970, 2.889600000000)
    pts(885) = Array(44.822028896454, 2.867200000000)
    pts(886) = Array(44.824095283420, 2.844800000000)
    pts(887) = Array(44.826145327141, 2.822400000000)
    pts(888) = Array(44.828179030863, 2.800000000000)
    pts(889) = Array(44.830196397806, 2.777600000000)
    pts(890) = Array(44.832197431161, 2.755200000000)
    pts(891) = Array(44.834182134095, 2.732800000000)
    pts(892) = Array(44.836150509747, 2.710400000000)
    pts(893) = Array(44.838102561230, 2.688000000000)
    pts(894) = Array(44.840038291631, 2.665600000000)
    pts(895) = Array(44.841957704007, 2.643200000000)
    pts(896) = Array(44.843860801393, 2.620800000000)
    pts(897) = Array(44.845747586795, 2.598400000000)
    pts(898) = Array(44.847618063194, 2.576000000000)
    pts(899) = Array(44.849472233541, 2.553600000000)
    pts(900) = Array(44.851310100765, 2.531200000000)
    pts(901) = Array(44.853131667766, 2.508800000000)
    pts(902) = Array(44.854936937419, 2.486400000000)
    pts(903) = Array(44.856725912570, 2.464000000000)
    pts(904) = Array(44.858498596042, 2.441600000000)
    pts(905) = Array(44.860254990630, 2.419200000000)
    pts(906) = Array(44.861995099103, 2.396800000000)
    pts(907) = Array(44.863718924202, 2.374400000000)
    pts(908) = Array(44.865426468645, 2.352000000000)
    pts(909) = Array(44.867117735121, 2.329600000000)
    pts(910) = Array(44.868792726295, 2.307200000000)
    pts(911) = Array(44.870451444803, 2.284800000000)
    pts(912) = Array(44.872093893257, 2.262400000000)
    pts(913) = Array(44.873720074242, 2.240000000000)
    pts(914) = Array(44.875329990318, 2.217600000000)
    pts(915) = Array(44.876923644017, 2.195200000000)
    pts(916) = Array(44.878501037846, 2.172800000000)
    pts(917) = Array(44.880062174286, 2.150400000000)
    pts(918) = Array(44.881607055792, 2.128000000000)
    pts(919) = Array(44.883135684792, 2.105600000000)
    pts(920) = Array(44.884648063689, 2.083200000000)
    pts(921) = Array(44.886144194859, 2.060800000000)
    pts(922) = Array(44.887624080654, 2.038400000000)
    pts(923) = Array(44.889087723397, 2.016000000000)
    pts(924) = Array(44.890535125388, 1.993600000000)
    pts(925) = Array(44.891966288899, 1.971200000000)
    pts(926) = Array(44.893381216178, 1.948800000000)
    pts(927) = Array(44.894779909444, 1.926400000000)
    pts(928) = Array(44.896162370894, 1.904000000000)
    pts(929) = Array(44.897528602696, 1.881600000000)
    pts(930) = Array(44.898878606994, 1.859200000000)
    pts(931) = Array(44.900212385905, 1.836800000000)
    pts(932) = Array(44.901529941521, 1.814400000000)
    pts(933) = Array(44.902831275908, 1.792000000000)
    pts(934) = Array(44.904116391107, 1.769600000000)
    pts(935) = Array(44.905385289130, 1.747200000000)
    pts(936) = Array(44.906637971969, 1.724800000000)
    pts(937) = Array(44.907874441584, 1.702400000000)
    pts(938) = Array(44.909094699913, 1.680000000000)
    pts(939) = Array(44.910298748869, 1.657600000000)
    pts(940) = Array(44.911486590336, 1.635200000000)
    pts(941) = Array(44.912658226175, 1.612800000000)
    pts(942) = Array(44.913813658220, 1.590400000000)
    pts(943) = Array(44.914952888280, 1.568000000000)
    pts(944) = Array(44.916075918138, 1.545600000000)
    pts(945) = Array(44.917182749553, 1.523200000000)
    pts(946) = Array(44.918273384255, 1.500800000000)
    pts(947) = Array(44.919347823951, 1.478400000000)
    pts(948) = Array(44.920406070323, 1.456000000000)
    pts(949) = Array(44.921448125025, 1.433600000000)
    pts(950) = Array(44.922473989688, 1.411200000000)
    pts(951) = Array(44.923483665915, 1.388800000000)
    pts(952) = Array(44.924477155284, 1.366400000000)
    pts(953) = Array(44.925454459350, 1.344000000000)
    pts(954) = Array(44.926415579640, 1.321600000000)
    pts(955) = Array(44.927360517656, 1.299200000000)
    pts(956) = Array(44.928289274875, 1.276800000000)
    pts(957) = Array(44.929201852748, 1.254400000000)
    pts(958) = Array(44.930098252701, 1.232000000000)
    pts(959) = Array(44.930978476133, 1.209600000000)
    pts(960) = Array(44.931842524421, 1.187200000000)
    pts(961) = Array(44.932690398914, 1.164800000000)
    pts(962) = Array(44.933522100935, 1.142400000000)
    pts(963) = Array(44.934337631784, 1.120000000000)
    pts(964) = Array(44.935136992734, 1.097600000000)
    pts(965) = Array(44.935920185033, 1.075200000000)
    pts(966) = Array(44.936687209902, 1.052800000000)
    pts(967) = Array(44.937438068541, 1.030400000000)
    pts(968) = Array(44.938172762119, 1.008000000000)
    pts(969) = Array(44.938891291785, 0.985600000000)
    pts(970) = Array(44.939593658659, 0.963200000000)
    pts(971) = Array(44.940279863837, 0.940800000000)
    pts(972) = Array(44.940949908389, 0.918400000000)
    pts(973) = Array(44.941603793362, 0.896000000000)
    pts(974) = Array(44.942241519774, 0.873600000000)
    pts(975) = Array(44.942863088621, 0.851200000000)
    pts(976) = Array(44.943468500872, 0.828800000000)
    pts(977) = Array(44.944057757471, 0.806400000000)
    pts(978) = Array(44.944630859337, 0.784000000000)
    pts(979) = Array(44.945187807364, 0.761600000000)
    pts(980) = Array(44.945728602419, 0.739200000000)
    pts(981) = Array(44.946253245347, 0.716800000000)
    pts(982) = Array(44.946761736964, 0.694400000000)
    pts(983) = Array(44.947254078064, 0.672000000000)
    pts(984) = Array(44.947730269414, 0.649600000000)
    pts(985) = Array(44.948190311757, 0.627200000000)
    pts(986) = Array(44.948634205808, 0.604800000000)
    pts(987) = Array(44.949061952261, 0.582400000000)
    pts(988) = Array(44.949473551781, 0.560000000000)
    pts(989) = Array(44.949869005010, 0.537600000000)
    pts(990) = Array(44.950248312564, 0.515200000000)
    pts(991) = Array(44.950611475034, 0.492800000000)
    pts(992) = Array(44.950958492987, 0.470400000000)
    pts(993) = Array(44.951289366961, 0.448000000000)
    pts(994) = Array(44.951604097474, 0.425600000000)
    pts(995) = Array(44.951902685016, 0.403200000000)
    pts(996) = Array(44.952185130050, 0.380800000000)
    pts(997) = Array(44.952451433019, 0.358400000000)
    pts(998) = Array(44.952701594335, 0.336000000000)
    pts(999) = Array(44.952935614389, 0.313600000000)
    pts(1000) = Array(44.953153493546, 0.291200000000)
    pts(1001) = Array(44.953355232145, 0.268800000000)
    pts(1002) = Array(44.953540830499, 0.246400000000)
    pts(1003) = Array(44.953710288899, 0.224000000000)
    pts(1004) = Array(44.953863607607, 0.201600000000)
    pts(1005) = Array(44.954000786863, 0.179200000000)
    pts(1006) = Array(44.954121826881, 0.156800000000)
    pts(1007) = Array(44.954226727848, 0.134400000000)
    pts(1008) = Array(44.954315489929, 0.112000000000)
    pts(1009) = Array(44.954388113261, 0.089600000000)
    pts(1010) = Array(44.954444597957, 0.067200000000)
    pts(1011) = Array(44.954484944106, 0.044800000000)
    pts(1012) = Array(44.954509151770, 0.022400000000)
    pts(1013) = Array(44.954517220987, 0.000000000000)
    Profile_13 = pts
End Function

Private Function Profile_14() As Variant
    Dim pts(1013) As Variant
    pts(0) = Array(44.954517220987, 0.000000000000)
    pts(1) = Array(44.954509151770, 0.022400000000)
    pts(2) = Array(44.954484944106, 0.044800000000)
    pts(3) = Array(44.954444597957, 0.067200000000)
    pts(4) = Array(44.954388113261, 0.089600000000)
    pts(5) = Array(44.954315489929, 0.112000000000)
    pts(6) = Array(44.954226727848, 0.134400000000)
    pts(7) = Array(44.954121826881, 0.156800000000)
    pts(8) = Array(44.954000786863, 0.179200000000)
    pts(9) = Array(44.953863607607, 0.201600000000)
    pts(10) = Array(44.953710288899, 0.224000000000)
    pts(11) = Array(44.953540830499, 0.246400000000)
    pts(12) = Array(44.953355232145, 0.268800000000)
    pts(13) = Array(44.953153493546, 0.291200000000)
    pts(14) = Array(44.952935614389, 0.313600000000)
    pts(15) = Array(44.952701594335, 0.336000000000)
    pts(16) = Array(44.952451433019, 0.358400000000)
    pts(17) = Array(44.952185130050, 0.380800000000)
    pts(18) = Array(44.951902685016, 0.403200000000)
    pts(19) = Array(44.951604097474, 0.425600000000)
    pts(20) = Array(44.951289366961, 0.448000000000)
    pts(21) = Array(44.950958492987, 0.470400000000)
    pts(22) = Array(44.950611475034, 0.492800000000)
    pts(23) = Array(44.950248312564, 0.515200000000)
    pts(24) = Array(44.949869005010, 0.537600000000)
    pts(25) = Array(44.949473551781, 0.560000000000)
    pts(26) = Array(44.949061952261, 0.582400000000)
    pts(27) = Array(44.948634205808, 0.604800000000)
    pts(28) = Array(44.948190311757, 0.627200000000)
    pts(29) = Array(44.947730269414, 0.649600000000)
    pts(30) = Array(44.947254078064, 0.672000000000)
    pts(31) = Array(44.946761736964, 0.694400000000)
    pts(32) = Array(44.946253245347, 0.716800000000)
    pts(33) = Array(44.945728602419, 0.739200000000)
    pts(34) = Array(44.945187807364, 0.761600000000)
    pts(35) = Array(44.944630859337, 0.784000000000)
    pts(36) = Array(44.944057757471, 0.806400000000)
    pts(37) = Array(44.943468500872, 0.828800000000)
    pts(38) = Array(44.942863088621, 0.851200000000)
    pts(39) = Array(44.942241519774, 0.873600000000)
    pts(40) = Array(44.941603793362, 0.896000000000)
    pts(41) = Array(44.940949908389, 0.918400000000)
    pts(42) = Array(44.940279863837, 0.940800000000)
    pts(43) = Array(44.939593658659, 0.963200000000)
    pts(44) = Array(44.938891291785, 0.985600000000)
    pts(45) = Array(44.938172762119, 1.008000000000)
    pts(46) = Array(44.937438068541, 1.030400000000)
    pts(47) = Array(44.936687209902, 1.052800000000)
    pts(48) = Array(44.935920185033, 1.075200000000)
    pts(49) = Array(44.935136992734, 1.097600000000)
    pts(50) = Array(44.934337631784, 1.120000000000)
    pts(51) = Array(44.933522100935, 1.142400000000)
    pts(52) = Array(44.932690398914, 1.164800000000)
    pts(53) = Array(44.931842524421, 1.187200000000)
    pts(54) = Array(44.930978476133, 1.209600000000)
    pts(55) = Array(44.930098252701, 1.232000000000)
    pts(56) = Array(44.929201852748, 1.254400000000)
    pts(57) = Array(44.928289274875, 1.276800000000)
    pts(58) = Array(44.927360517656, 1.299200000000)
    pts(59) = Array(44.926415579640, 1.321600000000)
    pts(60) = Array(44.925454459350, 1.344000000000)
    pts(61) = Array(44.924477155284, 1.366400000000)
    pts(62) = Array(44.923483665915, 1.388800000000)
    pts(63) = Array(44.922473989688, 1.411200000000)
    pts(64) = Array(44.921448125025, 1.433600000000)
    pts(65) = Array(44.920406070323, 1.456000000000)
    pts(66) = Array(44.919347823951, 1.478400000000)
    pts(67) = Array(44.918273384255, 1.500800000000)
    pts(68) = Array(44.917182749553, 1.523200000000)
    pts(69) = Array(44.916075918138, 1.545600000000)
    pts(70) = Array(44.914952888280, 1.568000000000)
    pts(71) = Array(44.913813658220, 1.590400000000)
    pts(72) = Array(44.912658226175, 1.612800000000)
    pts(73) = Array(44.911486590336, 1.635200000000)
    pts(74) = Array(44.910298748869, 1.657600000000)
    pts(75) = Array(44.909094699913, 1.680000000000)
    pts(76) = Array(44.907874441584, 1.702400000000)
    pts(77) = Array(44.906637971969, 1.724800000000)
    pts(78) = Array(44.905385289130, 1.747200000000)
    pts(79) = Array(44.904116391107, 1.769600000000)
    pts(80) = Array(44.902831275908, 1.792000000000)
    pts(81) = Array(44.901529941521, 1.814400000000)
    pts(82) = Array(44.900212385905, 1.836800000000)
    pts(83) = Array(44.898878606994, 1.859200000000)
    pts(84) = Array(44.897528602696, 1.881600000000)
    pts(85) = Array(44.896162370894, 1.904000000000)
    pts(86) = Array(44.894779909444, 1.926400000000)
    pts(87) = Array(44.893381216178, 1.948800000000)
    pts(88) = Array(44.891966288899, 1.971200000000)
    pts(89) = Array(44.890535125388, 1.993600000000)
    pts(90) = Array(44.889087723397, 2.016000000000)
    pts(91) = Array(44.887624080654, 2.038400000000)
    pts(92) = Array(44.886144194859, 2.060800000000)
    pts(93) = Array(44.884648063689, 2.083200000000)
    pts(94) = Array(44.883135684792, 2.105600000000)
    pts(95) = Array(44.881607055792, 2.128000000000)
    pts(96) = Array(44.880062174286, 2.150400000000)
    pts(97) = Array(44.878501037846, 2.172800000000)
    pts(98) = Array(44.876923644017, 2.195200000000)
    pts(99) = Array(44.875329990318, 2.217600000000)
    pts(100) = Array(44.873720074242, 2.240000000000)
    pts(101) = Array(44.872093893257, 2.262400000000)
    pts(102) = Array(44.870451444803, 2.284800000000)
    pts(103) = Array(44.868792726295, 2.307200000000)
    pts(104) = Array(44.867117735121, 2.329600000000)
    pts(105) = Array(44.865426468645, 2.352000000000)
    pts(106) = Array(44.863718924202, 2.374400000000)
    pts(107) = Array(44.861995099103, 2.396800000000)
    pts(108) = Array(44.860254990630, 2.419200000000)
    pts(109) = Array(44.858498596042, 2.441600000000)
    pts(110) = Array(44.856725912570, 2.464000000000)
    pts(111) = Array(44.854936937419, 2.486400000000)
    pts(112) = Array(44.853131667766, 2.508800000000)
    pts(113) = Array(44.851310100765, 2.531200000000)
    pts(114) = Array(44.849472233541, 2.553600000000)
    pts(115) = Array(44.847618063194, 2.576000000000)
    pts(116) = Array(44.845747586795, 2.598400000000)
    pts(117) = Array(44.843860801393, 2.620800000000)
    pts(118) = Array(44.841957704007, 2.643200000000)
    pts(119) = Array(44.840038291631, 2.665600000000)
    pts(120) = Array(44.838102561230, 2.688000000000)
    pts(121) = Array(44.836150509747, 2.710400000000)
    pts(122) = Array(44.834182134095, 2.732800000000)
    pts(123) = Array(44.832197431161, 2.755200000000)
    pts(124) = Array(44.830196397806, 2.777600000000)
    pts(125) = Array(44.828179030863, 2.800000000000)
    pts(126) = Array(44.826145327141, 2.822400000000)
    pts(127) = Array(44.824095283420, 2.844800000000)
    pts(128) = Array(44.822028896454, 2.867200000000)
    pts(129) = Array(44.819946162970, 2.889600000000)
    pts(130) = Array(44.817847079668, 2.912000000000)
    pts(131) = Array(44.815731643222, 2.934400000000)
    pts(132) = Array(44.813599850279, 2.956800000000)
    pts(133) = Array(44.811451697459, 2.979200000000)
    pts(134) = Array(44.809287181355, 3.001600000000)
    pts(135) = Array(44.807106298533, 3.024000000000)
    pts(136) = Array(44.804909045532, 3.046400000000)
    pts(137) = Array(44.802695418865, 3.068800000000)
    pts(138) = Array(44.800465415016, 3.091200000000)
    pts(139) = Array(44.798219030444, 3.113600000000)
    pts(140) = Array(44.795956261580, 3.136000000000)
    pts(141) = Array(44.793677104829, 3.158400000000)
    pts(142) = Array(44.791381556567, 3.180800000000)
    pts(143) = Array(44.789069613144, 3.203200000000)
    pts(144) = Array(44.786741270882, 3.225600000000)
    pts(145) = Array(44.784396526078, 3.248000000000)
    pts(146) = Array(44.782035375000, 3.270400000000)
    pts(147) = Array(44.779657813888, 3.292800000000)
    pts(148) = Array(44.777263838957, 3.315200000000)
    pts(149) = Array(44.774853446392, 3.337600000000)
    pts(150) = Array(44.772426632353, 3.360000000000)
    pts(151) = Array(44.769983392972, 3.382400000000)
    pts(152) = Array(44.767523724352, 3.404800000000)
    pts(153) = Array(44.765047622571, 3.427200000000)
    pts(154) = Array(44.762555083678, 3.449600000000)
    pts(155) = Array(44.760046103695, 3.472000000000)
    pts(156) = Array(44.757520678617, 3.494400000000)
    pts(157) = Array(44.754978804409, 3.516800000000)
    pts(158) = Array(44.752420477011, 3.539200000000)
    pts(159) = Array(44.749845692335, 3.561600000000)
    pts(160) = Array(44.747254446265, 3.584000000000)
    pts(161) = Array(44.744646734656, 3.606400000000)
    pts(162) = Array(44.742022553338, 3.628800000000)
    pts(163) = Array(44.739381898110, 3.651200000000)
    pts(164) = Array(44.736724764746, 3.673600000000)
    pts(165) = Array(44.734051148991, 3.696000000000)
    pts(166) = Array(44.731361046561, 3.718400000000)
    pts(167) = Array(44.728654453146, 3.740800000000)
    pts(168) = Array(44.725931364408, 3.763200000000)
    pts(169) = Array(44.723191775979, 3.785600000000)
    pts(170) = Array(44.720435683465, 3.808000000000)
    pts(171) = Array(44.717663082443, 3.830400000000)
    pts(172) = Array(44.714873968463, 3.852800000000)
    pts(173) = Array(44.712068337045, 3.875200000000)
    pts(174) = Array(44.709246183682, 3.897600000000)
    pts(175) = Array(44.706407503839, 3.920000000000)
    pts(176) = Array(44.703552292952, 3.942400000000)
    pts(177) = Array(44.700680546430, 3.964800000000)
    pts(178) = Array(44.697792259652, 3.987200000000)
    pts(179) = Array(44.694887427970, 4.009600000000)
    pts(180) = Array(44.691966046707, 4.032000000000)
    pts(181) = Array(44.689028111158, 4.054400000000)
    pts(182) = Array(44.686073616588, 4.076800000000)
    pts(183) = Array(44.683102558237, 4.099200000000)
    pts(184) = Array(44.680114931312, 4.121600000000)
    pts(185) = Array(44.677110730995, 4.144000000000)
    pts(186) = Array(44.674089952438, 4.166400000000)
    pts(187) = Array(44.671052590763, 4.188800000000)
    pts(188) = Array(44.667998641067, 4.211200000000)
    pts(189) = Array(44.664928098413, 4.233600000000)
    pts(190) = Array(44.661840957840, 4.256000000000)
    pts(191) = Array(44.658737214356, 4.278400000000)
    pts(192) = Array(44.655616862941, 4.300800000000)
    pts(193) = Array(44.652479898544, 4.323200000000)
    pts(194) = Array(44.649326316088, 4.345600000000)
    pts(195) = Array(44.646156110465, 4.368000000000)
    pts(196) = Array(44.642969276539, 4.390400000000)
    pts(197) = Array(44.639765809144, 4.412800000000)
    pts(198) = Array(44.636545703085, 4.435200000000)
    pts(199) = Array(44.633308953139, 4.457600000000)
    pts(200) = Array(44.630055554053, 4.480000000000)
    pts(201) = Array(44.626785500544, 4.502400000000)
    pts(202) = Array(44.623498787302, 4.524800000000)
    pts(203) = Array(44.620195408984, 4.547200000000)
    pts(204) = Array(44.616875360222, 4.569600000000)
    pts(205) = Array(44.613538635614, 4.592000000000)
    pts(206) = Array(44.610185229733, 4.614400000000)
    pts(207) = Array(44.606815137119, 4.636800000000)
    pts(208) = Array(44.603428352284, 4.659200000000)
    pts(209) = Array(44.600024869710, 4.681600000000)
    pts(210) = Array(44.596604683851, 4.704000000000)
    pts(211) = Array(44.593167789128, 4.726400000000)
    pts(212) = Array(44.589714179934, 4.748800000000)
    pts(213) = Array(44.586243850634, 4.771200000000)
    pts(214) = Array(44.582756795560, 4.793600000000)
    pts(215) = Array(44.579253009016, 4.816000000000)
    pts(216) = Array(44.575732485276, 4.838400000000)
    pts(217) = Array(44.572195218583, 4.860800000000)
    pts(218) = Array(44.568641203150, 4.883200000000)
    pts(219) = Array(44.565070433162, 4.905600000000)
    pts(220) = Array(44.561482902772, 4.928000000000)
    pts(221) = Array(44.557878606102, 4.950400000000)
    pts(222) = Array(44.554257537246, 4.972800000000)
    pts(223) = Array(44.550619690266, 4.995200000000)
    pts(224) = Array(44.546965059195, 5.017600000000)
    pts(225) = Array(44.543293638035, 5.040000000000)
    pts(226) = Array(44.539605420756, 5.062400000000)
    pts(227) = Array(44.535900401301, 5.084800000000)
    pts(228) = Array(44.532178573580, 5.107200000000)
    pts(229) = Array(44.528439931472, 5.129600000000)
    pts(230) = Array(44.524684468827, 5.152000000000)
    pts(231) = Array(44.520912179464, 5.174400000000)
    pts(232) = Array(44.517123057169, 5.196800000000)
    pts(233) = Array(44.513317095701, 5.219200000000)
    pts(234) = Array(44.509494288785, 5.241600000000)
    pts(235) = Array(44.505654630117, 5.264000000000)
    pts(236) = Array(44.501798113361, 5.286400000000)
    pts(237) = Array(44.497924732150, 5.308800000000)
    pts(238) = Array(44.494034480087, 5.331200000000)
    pts(239) = Array(44.490127350742, 5.353600000000)
    pts(240) = Array(44.486203337656, 5.376000000000)
    pts(241) = Array(44.482262434337, 5.398400000000)
    pts(242) = Array(44.478304634262, 5.420800000000)
    pts(243) = Array(44.474329930878, 5.443200000000)
    pts(244) = Array(44.470338317599, 5.465600000000)
    pts(245) = Array(44.466329787809, 5.488000000000)
    pts(246) = Array(44.462304334858, 5.510400000000)
    pts(247) = Array(44.458261952068, 5.532800000000)
    pts(248) = Array(44.454202632727, 5.555200000000)
    pts(249) = Array(44.450126370091, 5.577600000000)
    pts(250) = Array(44.446033157385, 5.600000000000)
    pts(251) = Array(44.441922987803, 5.622400000000)
    pts(252) = Array(44.437795854505, 5.644800000000)
    pts(253) = Array(44.433651750623, 5.667200000000)
    pts(254) = Array(44.429490669252, 5.689600000000)
    pts(255) = Array(44.425312603458, 5.712000000000)
    pts(256) = Array(44.421117546275, 5.734400000000)
    pts(257) = Array(44.416905490703, 5.756800000000)
    pts(258) = Array(44.412676429712, 5.779200000000)
    pts(259) = Array(44.408430356238, 5.801600000000)
    pts(260) = Array(44.404167263185, 5.824000000000)
    pts(261) = Array(44.399887143426, 5.846400000000)
    pts(262) = Array(44.395589989799, 5.868800000000)
    pts(263) = Array(44.391275795111, 5.891200000000)
    pts(264) = Array(44.386944552137, 5.913600000000)
    pts(265) = Array(44.382596253619, 5.936000000000)
    pts(266) = Array(44.378230892264, 5.958400000000)
    pts(267) = Array(44.373848460749, 5.980800000000)
    pts(268) = Array(44.369448951718, 6.003200000000)
    pts(269) = Array(44.365032357780, 6.025600000000)
    pts(270) = Array(44.360598671513, 6.048000000000)
    pts(271) = Array(44.356147885461, 6.070400000000)
    pts(272) = Array(44.351679992135, 6.092800000000)
    pts(273) = Array(44.347194984013, 6.115200000000)
    pts(274) = Array(44.342692853540, 6.137600000000)
    pts(275) = Array(44.338173593127, 6.160000000000)
    pts(276) = Array(44.333637195152, 6.182400000000)
    pts(277) = Array(44.329083651960, 6.204800000000)
    pts(278) = Array(44.324512955861, 6.227200000000)
    pts(279) = Array(44.319925099134, 6.249600000000)
    pts(280) = Array(44.315320074021, 6.272000000000)
    pts(281) = Array(44.310697872733, 6.294400000000)
    pts(282) = Array(44.306058487445, 6.316800000000)
    pts(283) = Array(44.301401910302, 6.339200000000)
    pts(284) = Array(44.296728133410, 6.361600000000)
    pts(285) = Array(44.292037148845, 6.384000000000)
    pts(286) = Array(44.287328948646, 6.406400000000)
    pts(287) = Array(44.282603524820, 6.428800000000)
    pts(288) = Array(44.277860869339, 6.451200000000)
    pts(289) = Array(44.273100974140, 6.473600000000)
    pts(290) = Array(44.268323831127, 6.496000000000)
    pts(291) = Array(44.263529432169, 6.518400000000)
    pts(292) = Array(44.258717769100, 6.540800000000)
    pts(293) = Array(44.253888833720, 6.563200000000)
    pts(294) = Array(44.249042617794, 6.585600000000)
    pts(295) = Array(44.244179113053, 6.608000000000)
    pts(296) = Array(44.239298311192, 6.630400000000)
    pts(297) = Array(44.234400203871, 6.652800000000)
    pts(298) = Array(44.229484782717, 6.675200000000)
    pts(299) = Array(44.224552039320, 6.697600000000)
    pts(300) = Array(44.219601965236, 6.720000000000)
    pts(301) = Array(44.214634551985, 6.742400000000)
    pts(302) = Array(44.209649791053, 6.764800000000)
    pts(303) = Array(44.204647673889, 6.787200000000)
    pts(304) = Array(44.199628191908, 6.809600000000)
    pts(305) = Array(44.194591336488, 6.832000000000)
    pts(306) = Array(44.189537098973, 6.854400000000)
    pts(307) = Array(44.184465470671, 6.876800000000)
    pts(308) = Array(44.179376442853, 6.899200000000)
    pts(309) = Array(44.174270006756, 6.921600000000)
    pts(310) = Array(44.169146153580, 6.944000000000)
    pts(311) = Array(44.164004874489, 6.966400000000)
    pts(312) = Array(44.158846160611, 6.988800000000)
    pts(313) = Array(44.153670003039, 7.011200000000)
    pts(314) = Array(44.148476392827, 7.033600000000)
    pts(315) = Array(44.143265320996, 7.056000000000)
    pts(316) = Array(44.138036778530, 7.078400000000)
    pts(317) = Array(44.132790756373, 7.100800000000)
    pts(318) = Array(44.127527245437, 7.123200000000)
    pts(319) = Array(44.122246236595, 7.145600000000)
    pts(320) = Array(44.116947720684, 7.168000000000)
    pts(321) = Array(44.111631688504, 7.190400000000)
    pts(322) = Array(44.106298130818, 7.212800000000)
    pts(323) = Array(44.100947038351, 7.235200000000)
    pts(324) = Array(44.095578401793, 7.257600000000)
    pts(325) = Array(44.090192211796, 7.280000000000)
    pts(326) = Array(44.084788458975, 7.302400000000)
    pts(327) = Array(44.079367133906, 7.324800000000)
    pts(328) = Array(44.073928227129, 7.347200000000)
    pts(329) = Array(44.068471729147, 7.369600000000)
    pts(330) = Array(44.062997630425, 7.392000000000)
    pts(331) = Array(44.057505921389, 7.414400000000)
    pts(332) = Array(44.051996592430, 7.436800000000)
    pts(333) = Array(44.046469633898, 7.459200000000)
    pts(334) = Array(44.040925036106, 7.481600000000)
    pts(335) = Array(44.036998610190, 7.497419679825)
    pts(336) = Array(44.036692458384, 7.498651685025)
    pts(337) = Array(44.036418085758, 7.499755623002)
    pts(338) = Array(44.035362789331, 7.504000000000)
    pts(339) = Array(44.029782883810, 7.526400000000)
    pts(340) = Array(44.024185309742, 7.548800000000)
    pts(341) = Array(44.018570057287, 7.571200000000)
    pts(342) = Array(44.012937116568, 7.593600000000)
    pts(343) = Array(44.007286477668, 7.616000000000)
    pts(344) = Array(44.001618130633, 7.638400000000)
    pts(345) = Array(43.995932065470, 7.660800000000)
    pts(346) = Array(43.990228272146, 7.683200000000)
    pts(347) = Array(43.984506740590, 7.705600000000)
    pts(348) = Array(43.978767460692, 7.728000000000)
    pts(349) = Array(43.973010422302, 7.750400000000)
    pts(350) = Array(43.967235615232, 7.772800000000)
    pts(351) = Array(43.961443029255, 7.795200000000)
    pts(352) = Array(43.955632654104, 7.817600000000)
    pts(353) = Array(43.949804479471, 7.840000000000)
    pts(354) = Array(43.943958495011, 7.862400000000)
    pts(355) = Array(43.938094690339, 7.884800000000)
    pts(356) = Array(43.932213055027, 7.907200000000)
    pts(357) = Array(43.926313578612, 7.929600000000)
    pts(358) = Array(43.920396250587, 7.952000000000)
    pts(359) = Array(43.914461060408, 7.974400000000)
    pts(360) = Array(43.908507997488, 7.996800000000)
    pts(361) = Array(43.902537051202, 8.019200000000)
    pts(362) = Array(43.896548210883, 8.041600000000)
    pts(363) = Array(43.890541465824, 8.064000000000)
    pts(364) = Array(43.884516805278, 8.086400000000)
    pts(365) = Array(43.878474218458, 8.108800000000)
    pts(366) = Array(43.872413694533, 8.131200000000)
    pts(367) = Array(43.866335222635, 8.153600000000)
    pts(368) = Array(43.860238791852, 8.176000000000)
    pts(369) = Array(43.854124391233, 8.198400000000)
    pts(370) = Array(43.847992009784, 8.220800000000)
    pts(371) = Array(43.841841636471, 8.243200000000)
    pts(372) = Array(43.835673260218, 8.265600000000)
    pts(373) = Array(43.829486869908, 8.288000000000)
    pts(374) = Array(43.823282454380, 8.310400000000)
    pts(375) = Array(43.817060002435, 8.332800000000)
    pts(376) = Array(43.810819502829, 8.355200000000)
    pts(377) = Array(43.804560944278, 8.377600000000)
    pts(378) = Array(43.798284315455, 8.400000000000)
    pts(379) = Array(43.791989604990, 8.422400000000)
    pts(380) = Array(43.785676801472, 8.444800000000)
    pts(381) = Array(43.779345893448, 8.467200000000)
    pts(382) = Array(43.772996869420, 8.489600000000)
    pts(383) = Array(43.766629717850, 8.512000000000)
    pts(384) = Array(43.760244427155, 8.534400000000)
    pts(385) = Array(43.753840985711, 8.556800000000)
    pts(386) = Array(43.747419381850, 8.579200000000)
    pts(387) = Array(43.740979603861, 8.601600000000)
    pts(388) = Array(43.734521639989, 8.624000000000)
    pts(389) = Array(43.728045478437, 8.646400000000)
    pts(390) = Array(43.721551107364, 8.668800000000)
    pts(391) = Array(43.715038514885, 8.691200000000)
    pts(392) = Array(43.708507689071, 8.713600000000)
    pts(393) = Array(43.701958617951, 8.736000000000)
    pts(394) = Array(43.695391289507, 8.758400000000)
    pts(395) = Array(43.688805691680, 8.780800000000)
    pts(396) = Array(43.682201812364, 8.803200000000)
    pts(397) = Array(43.675579639411, 8.825600000000)
    pts(398) = Array(43.668939160628, 8.848000000000)
    pts(399) = Array(43.662280363776, 8.870400000000)
    pts(400) = Array(43.655603236572, 8.892800000000)
    pts(401) = Array(43.648907766690, 8.915200000000)
    pts(402) = Array(43.642193941755, 8.937600000000)
    pts(403) = Array(43.635461749352, 8.960000000000)
    pts(404) = Array(43.628711177016, 8.982400000000)
    pts(405) = Array(43.621942212240, 9.004800000000)
    pts(406) = Array(43.615154842470, 9.027200000000)
    pts(407) = Array(43.608349055106, 9.049600000000)
    pts(408) = Array(43.601524837504, 9.072000000000)
    pts(409) = Array(43.594682176973, 9.094400000000)
    pts(410) = Array(43.587821060775, 9.116800000000)
    pts(411) = Array(43.580941476128, 9.139200000000)
    pts(412) = Array(43.574043410202, 9.161600000000)
    pts(413) = Array(43.567126850122, 9.184000000000)
    pts(414) = Array(43.560191782965, 9.206400000000)
    pts(415) = Array(43.553238195763, 9.228800000000)
    pts(416) = Array(43.546266075499, 9.251200000000)
    pts(417) = Array(43.539275409111, 9.273600000000)
    pts(418) = Array(43.532266183490, 9.296000000000)
    pts(419) = Array(43.525238385478, 9.318400000000)
    pts(420) = Array(43.518192001872, 9.340800000000)
    pts(421) = Array(43.511127019418, 9.363200000000)
    pts(422) = Array(43.504043424819, 9.385600000000)
    pts(423) = Array(43.496941204726, 9.408000000000)
    pts(424) = Array(43.489820345745, 9.430400000000)
    pts(425) = Array(43.482680834432, 9.452800000000)
    pts(426) = Array(43.475522657297, 9.475200000000)
    pts(427) = Array(43.468345800799, 9.497600000000)
    pts(428) = Array(43.461150251350, 9.520000000000)
    pts(429) = Array(43.453935995314, 9.542400000000)
    pts(430) = Array(43.446703019005, 9.564800000000)
    pts(431) = Array(43.439451308688, 9.587200000000)
    pts(432) = Array(43.432180850580, 9.609600000000)
    pts(433) = Array(43.424891630849, 9.632000000000)
    pts(434) = Array(43.417583635611, 9.654400000000)
    pts(435) = Array(43.410256850937, 9.676800000000)
    pts(436) = Array(43.402911262843, 9.699200000000)
    pts(437) = Array(43.395546857299, 9.721600000000)
    pts(438) = Array(43.388163620225, 9.744000000000)
    pts(439) = Array(43.380761537488, 9.766400000000)
    pts(440) = Array(43.373340594908, 9.788800000000)
    pts(441) = Array(43.365900778253, 9.811200000000)
    pts(442) = Array(43.358442073240, 9.833600000000)
    pts(443) = Array(43.350964465535, 9.856000000000)
    pts(444) = Array(43.343467940756, 9.878400000000)
    pts(445) = Array(43.335952484467, 9.900800000000)
    pts(446) = Array(43.328418082182, 9.923200000000)
    pts(447) = Array(43.320864719362, 9.945600000000)
    pts(448) = Array(43.313292381420, 9.968000000000)
    pts(449) = Array(43.305701053713, 9.990400000000)
    pts(450) = Array(43.298090721550, 10.012800000000)
    pts(451) = Array(43.290461370185, 10.035200000000)
    pts(452) = Array(43.282812984822, 10.057600000000)
    pts(453) = Array(43.275145550612, 10.080000000000)
    pts(454) = Array(43.267459052654, 10.102400000000)
    pts(455) = Array(43.259753475991, 10.124800000000)
    pts(456) = Array(43.252028805619, 10.147200000000)
    pts(457) = Array(43.244285026476, 10.169600000000)
    pts(458) = Array(43.236522123449, 10.192000000000)
    pts(459) = Array(43.228740081371, 10.214400000000)
    pts(460) = Array(43.220938885023, 10.236800000000)
    pts(461) = Array(43.213118519131, 10.259200000000)
    pts(462) = Array(43.205278968367, 10.281600000000)
    pts(463) = Array(43.197420217350, 10.304000000000)
    pts(464) = Array(43.189542250643, 10.326400000000)
    pts(465) = Array(43.181645052757, 10.348800000000)
    pts(466) = Array(43.173728608147, 10.371200000000)
    pts(467) = Array(43.165792901214, 10.393600000000)
    pts(468) = Array(43.157837916303, 10.416000000000)
    pts(469) = Array(43.149863637705, 10.438400000000)
    pts(470) = Array(43.141870049656, 10.460800000000)
    pts(471) = Array(43.133857136335, 10.483200000000)
    pts(472) = Array(43.127834759587, 10.500000000000)
    pts(473) = Array(43.127834759587, 10.505600000000)
    pts(474) = Array(43.127834759587, 10.528000000000)
    pts(475) = Array(43.127834759587, 10.550000000000)
    pts(476) = Array(43.127834759587, 10.550400000000)
    pts(477) = Array(43.127834759587, 10.572800000000)
    pts(478) = Array(43.127834759587, 10.595200000000)
    pts(479) = Array(43.127834759587, 10.617600000000)
    pts(480) = Array(43.127834759587, 10.640000000000)
    pts(481) = Array(43.127834759587, 10.662400000000)
    pts(482) = Array(43.127834759587, 10.684800000000)
    pts(483) = Array(43.127834759587, 10.700000000000)
    pts(484) = Array(43.127834759587, 10.707200000000)
    pts(485) = Array(43.127834759587, 10.729600000000)
    pts(486) = Array(43.127834759587, 10.752000000000)
    pts(487) = Array(43.127834759587, 10.774400000000)
    pts(488) = Array(43.127834759587, 10.796800000000)
    pts(489) = Array(43.127834759587, 10.819200000000)
    pts(490) = Array(43.127834759587, 10.841600000000)
    pts(491) = Array(43.127834759587, 10.864000000000)
    pts(492) = Array(43.127834759587, 10.886400000000)
    pts(493) = Array(43.127834759587, 10.908800000000)
    pts(494) = Array(43.127834759587, 10.931200000000)
    pts(495) = Array(43.127834759587, 10.953600000000)
    pts(496) = Array(43.127834759587, 10.976000000000)
    pts(497) = Array(43.127834759587, 10.998400000000)
    pts(498) = Array(43.127834759587, 11.020800000000)
    pts(499) = Array(43.127834759587, 11.043200000000)
    pts(500) = Array(43.127834759587, 11.065600000000)
    pts(501) = Array(43.127834759587, 11.088000000000)
    pts(502) = Array(43.127834759587, 11.110400000000)
    pts(503) = Array(43.127834759587, 11.132800000000)
    pts(504) = Array(43.127834759587, 11.155200000000)
    pts(505) = Array(43.127834759587, 11.177600000000)
    pts(506) = Array(43.127834759587, 11.200000000000)
    pts(507) = Array(47.922517220987, 11.200000000000)
    pts(508) = Array(47.922517220987, 11.177600000000)
    pts(509) = Array(47.922517220987, 11.155200000000)
    pts(510) = Array(47.922517220987, 11.132800000000)
    pts(511) = Array(47.922517220987, 11.110400000000)
    pts(512) = Array(47.922517220987, 11.088000000000)
    pts(513) = Array(47.922517220987, 11.065600000000)
    pts(514) = Array(47.922517220987, 11.043200000000)
    pts(515) = Array(47.922517220987, 11.020800000000)
    pts(516) = Array(47.922517220987, 10.998400000000)
    pts(517) = Array(47.922517220987, 10.976000000000)
    pts(518) = Array(47.922517220987, 10.953600000000)
    pts(519) = Array(47.922517220987, 10.931200000000)
    pts(520) = Array(47.922517220987, 10.908800000000)
    pts(521) = Array(47.922517220987, 10.886400000000)
    pts(522) = Array(47.922517220987, 10.864000000000)
    pts(523) = Array(47.922517220987, 10.841600000000)
    pts(524) = Array(47.922517220987, 10.819200000000)
    pts(525) = Array(47.922517220987, 10.796800000000)
    pts(526) = Array(47.922517220987, 10.774400000000)
    pts(527) = Array(47.922517220987, 10.752000000000)
    pts(528) = Array(47.922517220987, 10.729600000000)
    pts(529) = Array(47.922517220987, 10.707200000000)
    pts(530) = Array(47.922517220987, 10.700000000000)
    pts(531) = Array(47.918759446695, 10.684800000000)
    pts(532) = Array(47.913221674053, 10.662400000000)
    pts(533) = Array(47.907683901412, 10.640000000000)
    pts(534) = Array(47.902146128770, 10.617600000000)
    pts(535) = Array(47.896608356128, 10.595200000000)
    pts(536) = Array(47.891070583487, 10.572800000000)
    pts(537) = Array(47.885532810845, 10.550400000000)
    pts(538) = Array(47.885433922048, 10.550000000000)
    pts(539) = Array(47.885433922048, 10.528000000000)
    pts(540) = Array(47.885433922048, 10.505600000000)
    pts(541) = Array(47.885433922048, 10.500000000000)
    pts(542) = Array(47.885433922048, 10.483200000000)
    pts(543) = Array(47.885433922048, 10.460800000000)
    pts(544) = Array(47.885433922048, 10.438400000000)
    pts(545) = Array(47.885433922048, 10.416000000000)
    pts(546) = Array(47.885433922048, 10.393600000000)
    pts(547) = Array(47.885433922048, 10.371200000000)
    pts(548) = Array(47.885433922048, 10.348800000000)
    pts(549) = Array(47.885433922048, 10.326400000000)
    pts(550) = Array(47.885433922048, 10.304000000000)
    pts(551) = Array(47.885433922048, 10.281600000000)
    pts(552) = Array(47.885433922048, 10.259200000000)
    pts(553) = Array(47.885433922048, 10.236800000000)
    pts(554) = Array(47.885433922048, 10.214400000000)
    pts(555) = Array(47.885433922048, 10.192000000000)
    pts(556) = Array(47.885433922048, 10.169600000000)
    pts(557) = Array(47.885433922048, 10.147200000000)
    pts(558) = Array(47.885433922048, 10.124800000000)
    pts(559) = Array(47.885433922048, 10.102400000000)
    pts(560) = Array(47.885433922048, 10.080000000000)
    pts(561) = Array(47.885433922048, 10.057600000000)
    pts(562) = Array(47.885433922048, 10.035200000000)
    pts(563) = Array(47.885433922048, 10.012800000000)
    pts(564) = Array(47.885433922048, 9.990400000000)
    pts(565) = Array(47.885433922048, 9.968000000000)
    pts(566) = Array(47.885433922048, 9.945600000000)
    pts(567) = Array(47.885433922048, 9.923200000000)
    pts(568) = Array(47.885433922048, 9.900800000000)
    pts(569) = Array(47.885433922048, 9.878400000000)
    pts(570) = Array(47.885433922048, 9.856000000000)
    pts(571) = Array(47.885433922048, 9.833600000000)
    pts(572) = Array(47.885433922048, 9.811200000000)
    pts(573) = Array(47.885433922048, 9.788800000000)
    pts(574) = Array(47.885433922048, 9.766400000000)
    pts(575) = Array(47.885433922048, 9.744000000000)
    pts(576) = Array(47.885433922048, 9.721600000000)
    pts(577) = Array(47.885433922048, 9.699200000000)
    pts(578) = Array(47.885433922048, 9.676800000000)
    pts(579) = Array(47.885433922048, 9.654400000000)
    pts(580) = Array(47.885433922048, 9.632000000000)
    pts(581) = Array(47.885433922048, 9.609600000000)
    pts(582) = Array(47.885433922048, 9.587200000000)
    pts(583) = Array(47.885433922048, 9.564800000000)
    pts(584) = Array(47.885433922048, 9.542400000000)
    pts(585) = Array(47.885433922048, 9.520000000000)
    pts(586) = Array(47.885433922048, 9.497600000000)
    pts(587) = Array(47.885433922048, 9.475200000000)
    pts(588) = Array(47.885433922048, 9.452800000000)
    pts(589) = Array(47.885433922048, 9.430400000000)
    pts(590) = Array(47.885433922048, 9.408000000000)
    pts(591) = Array(47.885433922048, 9.385600000000)
    pts(592) = Array(47.885433922048, 9.363200000000)
    pts(593) = Array(47.885433922048, 9.340800000000)
    pts(594) = Array(47.885433922048, 9.318400000000)
    pts(595) = Array(47.885433922048, 9.296000000000)
    pts(596) = Array(47.885433922048, 9.273600000000)
    pts(597) = Array(47.885433922048, 9.251200000000)
    pts(598) = Array(47.885433922048, 9.228800000000)
    pts(599) = Array(47.885433922048, 9.206400000000)
    pts(600) = Array(47.885433922048, 9.184000000000)
    pts(601) = Array(47.885433922048, 9.161600000000)
    pts(602) = Array(47.885433922048, 9.139200000000)
    pts(603) = Array(47.885433922048, 9.116800000000)
    pts(604) = Array(47.885433922048, 9.094400000000)
    pts(605) = Array(47.885433922048, 9.072000000000)
    pts(606) = Array(47.885433922048, 9.049600000000)
    pts(607) = Array(47.885433922048, 9.027200000000)
    pts(608) = Array(47.885433922048, 9.004800000000)
    pts(609) = Array(47.885433922048, 8.982400000000)
    pts(610) = Array(47.885433922048, 8.960000000000)
    pts(611) = Array(47.885433922048, 8.937600000000)
    pts(612) = Array(47.885433922048, 8.915200000000)
    pts(613) = Array(47.885433922048, 8.892800000000)
    pts(614) = Array(47.885433922048, 8.870400000000)
    pts(615) = Array(47.885433922048, 8.848000000000)
    pts(616) = Array(47.885433922048, 8.825600000000)
    pts(617) = Array(47.885433922048, 8.803200000000)
    pts(618) = Array(47.885433922048, 8.780800000000)
    pts(619) = Array(47.885433922048, 8.758400000000)
    pts(620) = Array(47.885433922048, 8.736000000000)
    pts(621) = Array(47.885433922048, 8.713600000000)
    pts(622) = Array(47.885433922048, 8.691200000000)
    pts(623) = Array(47.885433922048, 8.668800000000)
    pts(624) = Array(47.885433922048, 8.646400000000)
    pts(625) = Array(47.885433922048, 8.624000000000)
    pts(626) = Array(47.885433922048, 8.601600000000)
    pts(627) = Array(47.885433922048, 8.579200000000)
    pts(628) = Array(47.885433922048, 8.556800000000)
    pts(629) = Array(47.885433922048, 8.534400000000)
    pts(630) = Array(47.885433922048, 8.512000000000)
    pts(631) = Array(47.885433922048, 8.489600000000)
    pts(632) = Array(47.885433922048, 8.467200000000)
    pts(633) = Array(47.885433922048, 8.444800000000)
    pts(634) = Array(47.885433922048, 8.422400000000)
    pts(635) = Array(47.885433922048, 8.400000000000)
    pts(636) = Array(47.885433922048, 8.377600000000)
    pts(637) = Array(47.885433922048, 8.355200000000)
    pts(638) = Array(47.885433922048, 8.332800000000)
    pts(639) = Array(47.885433922048, 8.310400000000)
    pts(640) = Array(47.885433922048, 8.288000000000)
    pts(641) = Array(47.885433922048, 8.265600000000)
    pts(642) = Array(47.885433922048, 8.243200000000)
    pts(643) = Array(47.885433922048, 8.220800000000)
    pts(644) = Array(47.885433922048, 8.198400000000)
    pts(645) = Array(47.885433922048, 8.176000000000)
    pts(646) = Array(47.885433922048, 8.153600000000)
    pts(647) = Array(47.885433922048, 8.131200000000)
    pts(648) = Array(47.885433922048, 8.108800000000)
    pts(649) = Array(47.885433922048, 8.086400000000)
    pts(650) = Array(47.885433922048, 8.064000000000)
    pts(651) = Array(47.885433922048, 8.041600000000)
    pts(652) = Array(47.885433922048, 8.019200000000)
    pts(653) = Array(47.885433922048, 7.996800000000)
    pts(654) = Array(47.885433922048, 7.974400000000)
    pts(655) = Array(47.885433922048, 7.952000000000)
    pts(656) = Array(47.885433922048, 7.929600000000)
    pts(657) = Array(47.885433922048, 7.907200000000)
    pts(658) = Array(47.885433922048, 7.884800000000)
    pts(659) = Array(47.885433922048, 7.862400000000)
    pts(660) = Array(47.885433922048, 7.840000000000)
    pts(661) = Array(47.885433922048, 7.817600000000)
    pts(662) = Array(47.885433922048, 7.795200000000)
    pts(663) = Array(47.885433922048, 7.772800000000)
    pts(664) = Array(47.885433922048, 7.750400000000)
    pts(665) = Array(47.885433922048, 7.728000000000)
    pts(666) = Array(47.885433922048, 7.705600000000)
    pts(667) = Array(47.885433922048, 7.683200000000)
    pts(668) = Array(47.885433922048, 7.660800000000)
    pts(669) = Array(47.885433922048, 7.638400000000)
    pts(670) = Array(47.885433922048, 7.616000000000)
    pts(671) = Array(47.885433922048, 7.593600000000)
    pts(672) = Array(47.885433922048, 7.571200000000)
    pts(673) = Array(47.885433922048, 7.548800000000)
    pts(674) = Array(47.885433922048, 7.526400000000)
    pts(675) = Array(47.885433922048, 7.504000000000)
    pts(676) = Array(47.885433922048, 7.499755623002)
    pts(677) = Array(47.885433922048, 7.498651685025)
    pts(678) = Array(47.885433922048, 7.497419679825)
    pts(679) = Array(47.881417471317, 7.481600000000)
    pts(680) = Array(47.875745788653, 7.459200000000)
    pts(681) = Array(47.870092197618, 7.436800000000)
    pts(682) = Array(47.864456687815, 7.414400000000)
    pts(683) = Array(47.858839248892, 7.392000000000)
    pts(684) = Array(47.853239870534, 7.369600000000)
    pts(685) = Array(47.847658542470, 7.347200000000)
    pts(686) = Array(47.842095254468, 7.324800000000)
    pts(687) = Array(47.836549996339, 7.302400000000)
    pts(688) = Array(47.831022757932, 7.280000000000)
    pts(689) = Array(47.825513529138, 7.257600000000)
    pts(690) = Array(47.820022299888, 7.235200000000)
    pts(691) = Array(47.814549060155, 7.212800000000)
    pts(692) = Array(47.809093799949, 7.190400000000)
    pts(693) = Array(47.803656509324, 7.168000000000)
    pts(694) = Array(47.798237178370, 7.145600000000)
    pts(695) = Array(47.792835797221, 7.123200000000)
    pts(696) = Array(47.787452356047, 7.100800000000)
    pts(697) = Array(47.782086845062, 7.078400000000)
    pts(698) = Array(47.776739254516, 7.056000000000)
    pts(699) = Array(47.771409574700, 7.033600000000)
    pts(700) = Array(47.766097795945, 7.011200000000)
    pts(701) = Array(47.760803908621, 6.988800000000)
    pts(702) = Array(47.755527903137, 6.966400000000)
    pts(703) = Array(47.750269769942, 6.944000000000)
    pts(704) = Array(47.745029499522, 6.921600000000)
    pts(705) = Array(47.739807082405, 6.899200000000)
    pts(706) = Array(47.734602509156, 6.876800000000)
    pts(707) = Array(47.729415770380, 6.854400000000)
    pts(708) = Array(47.724246856720, 6.832000000000)
    pts(709) = Array(47.719095758857, 6.809600000000)
    pts(710) = Array(47.713962467511, 6.787200000000)
    pts(711) = Array(47.708846973443, 6.764800000000)
    pts(712) = Array(47.703749267449, 6.742400000000)
    pts(713) = Array(47.698669340364, 6.720000000000)
    pts(714) = Array(47.693607183062, 6.697600000000)
    pts(715) = Array(47.688562786456, 6.675200000000)
    pts(716) = Array(47.683536141496, 6.652800000000)
    pts(717) = Array(47.678527239168, 6.630400000000)
    pts(718) = Array(47.673536070500, 6.608000000000)
    pts(719) = Array(47.668562626555, 6.585600000000)
    pts(720) = Array(47.663606898433, 6.563200000000)
    pts(721) = Array(47.658668877275, 6.540800000000)
    pts(722) = Array(47.653748554256, 6.518400000000)
    pts(723) = Array(47.648845920591, 6.496000000000)
    pts(724) = Array(47.643960967531, 6.473600000000)
    pts(725) = Array(47.639093686364, 6.451200000000)
    pts(726) = Array(47.634244068416, 6.428800000000)
    pts(727) = Array(47.629412105050, 6.406400000000)
    pts(728) = Array(47.624597787667, 6.384000000000)
    pts(729) = Array(47.619801107702, 6.361600000000)
    pts(730) = Array(47.615022056631, 6.339200000000)
    pts(731) = Array(47.610260625963, 6.316800000000)
    pts(732) = Array(47.605516807245, 6.294400000000)
    pts(733) = Array(47.600790592063, 6.272000000000)
    pts(734) = Array(47.596081972036, 6.249600000000)
    pts(735) = Array(47.591390938821, 6.227200000000)
    pts(736) = Array(47.586717484112, 6.204800000000)
    pts(737) = Array(47.582061599638, 6.182400000000)
    pts(738) = Array(47.577423277166, 6.160000000000)
    pts(739) = Array(47.572802508498, 6.137600000000)
    pts(740) = Array(47.568199285472, 6.115200000000)
    pts(741) = Array(47.563613599962, 6.092800000000)
    pts(742) = Array(47.559045443879, 6.070400000000)
    pts(743) = Array(47.554494809168, 6.048000000000)
    pts(744) = Array(47.549961687812, 6.025600000000)
    pts(745) = Array(47.545446071828, 6.003200000000)
    pts(746) = Array(47.540947953270, 5.980800000000)
    pts(747) = Array(47.536467324225, 5.958400000000)
    pts(748) = Array(47.532004176820, 5.936000000000)
    pts(749) = Array(47.527558503213, 5.913600000000)
    pts(750) = Array(47.523130295600, 5.891200000000)
    pts(751) = Array(47.518719546211, 5.868800000000)
    pts(752) = Array(47.514326247311, 5.846400000000)
    pts(753) = Array(47.509950391202, 5.824000000000)
    pts(754) = Array(47.505591970219, 5.801600000000)
    pts(755) = Array(47.501250976733, 5.779200000000)
    pts(756) = Array(47.496927403150, 5.756800000000)
    pts(757) = Array(47.492621241911, 5.734400000000)
    pts(758) = Array(47.488332485491, 5.712000000000)
    pts(759) = Array(47.484061126399, 5.689600000000)
    pts(760) = Array(47.479807157182, 5.667200000000)
    pts(761) = Array(47.475570570418, 5.644800000000)
    pts(762) = Array(47.471351358721, 5.622400000000)
    pts(763) = Array(47.467149514740, 5.600000000000)
    pts(764) = Array(47.462965031157, 5.577600000000)
    pts(765) = Array(47.458797900689, 5.555200000000)
    pts(766) = Array(47.454648116088, 5.532800000000)
    pts(767) = Array(47.450515670139, 5.510400000000)
    pts(768) = Array(47.446400555662, 5.488000000000)
    pts(769) = Array(47.442302765510, 5.465600000000)
    pts(770) = Array(47.438222292571, 5.443200000000)
    pts(771) = Array(47.434159129766, 5.420800000000)
    pts(772) = Array(47.430113270052, 5.398400000000)
    pts(773) = Array(47.426084706417, 5.376000000000)
    pts(774) = Array(47.422073431884, 5.353600000000)
    pts(775) = Array(47.418079439510, 5.331200000000)
    pts(776) = Array(47.414102722386, 5.308800000000)
    pts(777) = Array(47.410143273634, 5.286400000000)
    pts(778) = Array(47.406201086412, 5.264000000000)
    pts(779) = Array(47.402276153911, 5.241600000000)
    pts(780) = Array(47.398368469355, 5.219200000000)
    pts(781) = Array(47.394478026002, 5.196800000000)
    pts(782) = Array(47.390604817141, 5.174400000000)
    pts(783) = Array(47.386748836097, 5.152000000000)
    pts(784) = Array(47.382910076226, 5.129600000000)
    pts(785) = Array(47.379088530919, 5.107200000000)
    pts(786) = Array(47.375284193599, 5.084800000000)
    pts(787) = Array(47.371497057721, 5.062400000000)
    pts(788) = Array(47.367727116774, 5.040000000000)
    pts(789) = Array(47.363974364280, 5.017600000000)
    pts(790) = Array(47.360238793793, 4.995200000000)
    pts(791) = Array(47.356520398900, 4.972800000000)
    pts(792) = Array(47.352819173221, 4.950400000000)
    pts(793) = Array(47.349135110409, 4.928000000000)
    pts(794) = Array(47.345468204148, 4.905600000000)
    pts(795) = Array(47.341818448156, 4.883200000000)
    pts(796) = Array(47.338185836182, 4.860800000000)
    pts(797) = Array(47.334570362009, 4.838400000000)
    pts(798) = Array(47.330972019451, 4.816000000000)
    pts(799) = Array(47.327390802355, 4.793600000000)
    pts(800) = Array(47.323826704599, 4.771200000000)
    pts(801) = Array(47.320279720096, 4.748800000000)
    pts(802) = Array(47.316749842788, 4.726400000000)
    pts(803) = Array(47.313237066650, 4.704000000000)
    pts(804) = Array(47.309741385691, 4.681600000000)
    pts(805) = Array(47.306262793948, 4.659200000000)
    pts(806) = Array(47.302801285494, 4.636800000000)
    pts(807) = Array(47.299356854430, 4.614400000000)
    pts(808) = Array(47.295929494893, 4.592000000000)
    pts(809) = Array(47.292519201047, 4.569600000000)
    pts(810) = Array(47.289125967092, 4.547200000000)
    pts(811) = Array(47.285749787258, 4.524800000000)
    pts(812) = Array(47.282390655805, 4.502400000000)
    pts(813) = Array(47.279048567027, 4.480000000000)
    pts(814) = Array(47.275723515247, 4.457600000000)
    pts(815) = Array(47.272415494822, 4.435200000000)
    pts(816) = Array(47.269124500140, 4.412800000000)
    pts(817) = Array(47.265850525618, 4.390400000000)
    pts(818) = Array(47.262593565706, 4.368000000000)
    pts(819) = Array(47.259353614885, 4.345600000000)
    pts(820) = Array(47.256130667668, 4.323200000000)
    pts(821) = Array(47.252924718599, 4.300800000000)
    pts(822) = Array(47.249735762250, 4.278400000000)
    pts(823) = Array(47.246563793229, 4.256000000000)
    pts(824) = Array(47.243408806172, 4.233600000000)
    pts(825) = Array(47.240270795746, 4.211200000000)
    pts(826) = Array(47.237149756649, 4.188800000000)
    pts(827) = Array(47.234045683611, 4.166400000000)
    pts(828) = Array(47.230958571392, 4.144000000000)
    pts(829) = Array(47.227888414783, 4.121600000000)
    pts(830) = Array(47.224835208605, 4.099200000000)
    pts(831) = Array(47.221798947710, 4.076800000000)
    pts(832) = Array(47.218779626982, 4.054400000000)
    pts(833) = Array(47.215777241334, 4.032000000000)
    pts(834) = Array(47.212791785710, 4.009600000000)
    pts(835) = Array(47.209823255085, 3.987200000000)
    pts(836) = Array(47.206871644463, 3.964800000000)
    pts(837) = Array(47.203936948880, 3.942400000000)
    pts(838) = Array(47.201019163402, 3.920000000000)
    pts(839) = Array(47.198118283125, 3.897600000000)
    pts(840) = Array(47.195234303176, 3.875200000000)
    pts(841) = Array(47.192367218710, 3.852800000000)
    pts(842) = Array(47.189517024915, 3.830400000000)
    pts(843) = Array(47.186683717009, 3.808000000000)
    pts(844) = Array(47.183867290237, 3.785600000000)
    pts(845) = Array(47.181067739878, 3.763200000000)
    pts(846) = Array(47.178285061239, 3.740800000000)
    pts(847) = Array(47.175519249656, 3.718400000000)
    pts(848) = Array(47.172770300498, 3.696000000000)
    pts(849) = Array(47.170038209162, 3.673600000000)
    pts(850) = Array(47.167322971074, 3.651200000000)
    pts(851) = Array(47.164624581692, 3.628800000000)
    pts(852) = Array(47.161943036502, 3.606400000000)
    pts(853) = Array(47.159278331021, 3.584000000000)
    pts(854) = Array(47.156630460795, 3.561600000000)
    pts(855) = Array(47.153999421400, 3.539200000000)
    pts(856) = Array(47.151385208441, 3.516800000000)
    pts(857) = Array(47.148787817555, 3.494400000000)
    pts(858) = Array(47.146207244405, 3.472000000000)
    pts(859) = Array(47.143643484686, 3.449600000000)
    pts(860) = Array(47.141096534121, 3.427200000000)
    pts(861) = Array(47.138566388466, 3.404800000000)
    pts(862) = Array(47.136053043500, 3.382400000000)
    pts(863) = Array(47.133556495038, 3.360000000000)
    pts(864) = Array(47.131076738921, 3.337600000000)
    pts(865) = Array(47.128613771019, 3.315200000000)
    pts(866) = Array(47.126167587233, 3.292800000000)
    pts(867) = Array(47.123738183491, 3.270400000000)
    pts(868) = Array(47.121325555754, 3.248000000000)
    pts(869) = Array(47.118929700008, 3.225600000000)
    pts(870) = Array(47.116550612270, 3.203200000000)
    pts(871) = Array(47.114188288587, 3.180800000000)
    pts(872) = Array(47.111842725033, 3.158400000000)
    pts(873) = Array(47.109513917714, 3.136000000000)
    pts(874) = Array(47.107201862762, 3.113600000000)
    pts(875) = Array(47.104906556339, 3.091200000000)
    pts(876) = Array(47.102627994638, 3.068800000000)
    pts(877) = Array(47.100366173877, 3.046400000000)
    pts(878) = Array(47.098121090306, 3.024000000000)
    pts(879) = Array(47.095892740204, 3.001600000000)
    pts(880) = Array(47.093681119876, 2.979200000000)
    pts(881) = Array(47.091486225658, 2.956800000000)
    pts(882) = Array(47.089308053915, 2.934400000000)
    pts(883) = Array(47.087146601040, 2.912000000000)
    pts(884) = Array(47.085001863454, 2.889600000000)
    pts(885) = Array(47.082873837607, 2.867200000000)
    pts(886) = Array(47.080762519980, 2.844800000000)
    pts(887) = Array(47.078667907080, 2.822400000000)
    pts(888) = Array(47.076589995443, 2.800000000000)
    pts(889) = Array(47.074528781634, 2.777600000000)
    pts(890) = Array(47.072484262246, 2.755200000000)
    pts(891) = Array(47.070456433901, 2.732800000000)
    pts(892) = Array(47.068445293251, 2.710400000000)
    pts(893) = Array(47.066450836972, 2.688000000000)
    pts(894) = Array(47.064473061774, 2.665600000000)
    pts(895) = Array(47.062511964392, 2.643200000000)
    pts(896) = Array(47.060567541588, 2.620800000000)
    pts(897) = Array(47.058639790157, 2.598400000000)
    pts(898) = Array(47.056728706919, 2.576000000000)
    pts(899) = Array(47.054834288722, 2.553600000000)
    pts(900) = Array(47.052956532445, 2.531200000000)
    pts(901) = Array(47.051095434991, 2.508800000000)
    pts(902) = Array(47.049250993296, 2.486400000000)
    pts(903) = Array(47.047423204321, 2.464000000000)
    pts(904) = Array(47.045612065056, 2.441600000000)
    pts(905) = Array(47.043817572519, 2.419200000000)
    pts(906) = Array(47.042039723756, 2.396800000000)
    pts(907) = Array(47.040278515842, 2.374400000000)
    pts(908) = Array(47.038533945880, 2.352000000000)
    pts(909) = Array(47.036806010999, 2.329600000000)
    pts(910) = Array(47.035094708359, 2.307200000000)
    pts(911) = Array(47.033400035145, 2.284800000000)
    pts(912) = Array(47.031721988572, 2.262400000000)
    pts(913) = Array(47.030060565883, 2.240000000000)
    pts(914) = Array(47.028415764347, 2.217600000000)
    pts(915) = Array(47.026787581264, 2.195200000000)
    pts(916) = Array(47.025176013958, 2.172800000000)
    pts(917) = Array(47.023581059784, 2.150400000000)
    pts(918) = Array(47.022002716124, 2.128000000000)
    pts(919) = Array(47.020440980387, 2.105600000000)
    pts(920) = Array(47.018895850011, 2.083200000000)
    pts(921) = Array(47.017367322461, 2.060800000000)
    pts(922) = Array(47.015855395229, 2.038400000000)
    pts(923) = Array(47.014360065836, 2.016000000000)
    pts(924) = Array(47.012881331830, 1.993600000000)
    pts(925) = Array(47.011419190788, 1.971200000000)
    pts(926) = Array(47.009973640312, 1.948800000000)
    pts(927) = Array(47.008544678035, 1.926400000000)
    pts(928) = Array(47.007132301616, 1.904000000000)
    pts(929) = Array(47.005736508739, 1.881600000000)
    pts(930) = Array(47.004357297121, 1.859200000000)
    pts(931) = Array(47.002994664502, 1.836800000000)
    pts(932) = Array(47.001648608652, 1.814400000000)
    pts(933) = Array(47.000319127367, 1.792000000000)
    pts(934) = Array(46.999006218472, 1.769600000000)
    pts(935) = Array(46.997709879818, 1.747200000000)
    pts(936) = Array(46.996430109284, 1.724800000000)
    pts(937) = Array(46.995166904778, 1.702400000000)
    pts(938) = Array(46.993920264234, 1.680000000000)
    pts(939) = Array(46.992690185613, 1.657600000000)
    pts(940) = Array(46.991476666903, 1.635200000000)
    pts(941) = Array(46.990279706122, 1.612800000000)
    pts(942) = Array(46.989099301313, 1.590400000000)
    pts(943) = Array(46.987935450547, 1.568000000000)
    pts(944) = Array(46.986788151922, 1.545600000000)
    pts(945) = Array(46.985657403565, 1.523200000000)
    pts(946) = Array(46.984543203629, 1.500800000000)
    pts(947) = Array(46.983445550293, 1.478400000000)
    pts(948) = Array(46.982364441766, 1.456000000000)
    pts(949) = Array(46.981299876282, 1.433600000000)
    pts(950) = Array(46.980251852104, 1.411200000000)
    pts(951) = Array(46.979220367521, 1.388800000000)
    pts(952) = Array(46.978205420850, 1.366400000000)
    pts(953) = Array(46.977207010434, 1.344000000000)
    pts(954) = Array(46.976225134644, 1.321600000000)
    pts(955) = Array(46.975259791880, 1.299200000000)
    pts(956) = Array(46.974310980566, 1.276800000000)
    pts(957) = Array(46.973378699155, 1.254400000000)
    pts(958) = Array(46.972462946126, 1.232000000000)
    pts(959) = Array(46.971563719987, 1.209600000000)
    pts(960) = Array(46.970681019271, 1.187200000000)
    pts(961) = Array(46.969814842540, 1.164800000000)
    pts(962) = Array(46.968965188382, 1.142400000000)
    pts(963) = Array(46.968132055412, 1.120000000000)
    pts(964) = Array(46.967315442272, 1.097600000000)
    pts(965) = Array(46.966515347631, 1.075200000000)
    pts(966) = Array(46.965731770187, 1.052800000000)
    pts(967) = Array(46.964964708663, 1.030400000000)
    pts(968) = Array(46.964214161809, 1.008000000000)
    pts(969) = Array(46.963480128403, 0.985600000000)
    pts(970) = Array(46.962762607250, 0.963200000000)
    pts(971) = Array(46.962061597181, 0.940800000000)
    pts(972) = Array(46.961377097054, 0.918400000000)
    pts(973) = Array(46.960709105755, 0.896000000000)
    pts(974) = Array(46.960057622198, 0.873600000000)
    pts(975) = Array(46.959422645320, 0.851200000000)
    pts(976) = Array(46.958804174089, 0.828800000000)
    pts(977) = Array(46.958202207499, 0.806400000000)
    pts(978) = Array(46.957616744568, 0.784000000000)
    pts(979) = Array(46.957047784346, 0.761600000000)
    pts(980) = Array(46.956495325906, 0.739200000000)
    pts(981) = Array(46.955959368348, 0.716800000000)
    pts(982) = Array(46.955439910802, 0.694400000000)
    pts(983) = Array(46.954936952422, 0.672000000000)
    pts(984) = Array(46.954450492390, 0.649600000000)
    pts(985) = Array(46.953980529914, 0.627200000000)
    pts(986) = Array(46.953527064231, 0.604800000000)
    pts(987) = Array(46.953090094602, 0.582400000000)
    pts(988) = Array(46.952669620317, 0.560000000000)
    pts(989) = Array(46.952265640693, 0.537600000000)
    pts(990) = Array(46.951878155072, 0.515200000000)
    pts(991) = Array(46.951507162825, 0.492800000000)
    pts(992) = Array(46.951152663347, 0.470400000000)
    pts(993) = Array(46.950814656063, 0.448000000000)
    pts(994) = Array(46.950493140424, 0.425600000000)
    pts(995) = Array(46.950188115906, 0.403200000000)
    pts(996) = Array(46.949899582013, 0.380800000000)
    pts(997) = Array(46.949627538277, 0.358400000000)
    pts(998) = Array(46.949371984256, 0.336000000000)
    pts(999) = Array(46.949132919533, 0.313600000000)
    pts(1000) = Array(46.948910343721, 0.291200000000)
    pts(1001) = Array(46.948704256457, 0.268800000000)
    pts(1002) = Array(46.948514657407, 0.246400000000)
    pts(1003) = Array(46.948341546263, 0.224000000000)
    pts(1004) = Array(46.948184922742, 0.201600000000)
    pts(1005) = Array(46.948044786592, 0.179200000000)
    pts(1006) = Array(46.947921137583, 0.156800000000)
    pts(1007) = Array(46.947813975514, 0.134400000000)
    pts(1008) = Array(46.947723300213, 0.112000000000)
    pts(1009) = Array(46.947649111531, 0.089600000000)
    pts(1010) = Array(46.947591409348, 0.067200000000)
    pts(1011) = Array(46.947550193570, 0.044800000000)
    pts(1012) = Array(46.947525464130, 0.022400000000)
    pts(1013) = Array(46.947517220987, 0.000000000000)
    Profile_14 = pts
End Function

Private Function Profile_15() As Variant
    Dim pts(1013) As Variant
    pts(0) = Array(48.960517220987, 0.000000000000)
    pts(1) = Array(48.960508415707, 0.022400000000)
    pts(2) = Array(48.960481999851, 0.044800000000)
    pts(3) = Array(48.960437973369, 0.067200000000)
    pts(4) = Array(48.960376336180, 0.089600000000)
    pts(5) = Array(48.960297088170, 0.112000000000)
    pts(6) = Array(48.960200229192, 0.134400000000)
    pts(7) = Array(48.960085759065, 0.156800000000)
    pts(8) = Array(48.959953677579, 0.179200000000)
    pts(9) = Array(48.959803984488, 0.201600000000)
    pts(10) = Array(48.959636679513, 0.224000000000)
    pts(11) = Array(48.959451762346, 0.246400000000)
    pts(12) = Array(48.959249232643, 0.268800000000)
    pts(13) = Array(48.959029090029, 0.291200000000)
    pts(14) = Array(48.958791334095, 0.313600000000)
    pts(15) = Array(48.958535964400, 0.336000000000)
    pts(16) = Array(48.958262980471, 0.358400000000)
    pts(17) = Array(48.957972381801, 0.380800000000)
    pts(18) = Array(48.957664167851, 0.403200000000)
    pts(19) = Array(48.957338338050, 0.425600000000)
    pts(20) = Array(48.956994891793, 0.448000000000)
    pts(21) = Array(48.956633828443, 0.470400000000)
    pts(22) = Array(48.956255147330, 0.492800000000)
    pts(23) = Array(48.955858847752, 0.515200000000)
    pts(24) = Array(48.955444928973, 0.537600000000)
    pts(25) = Array(48.955013390225, 0.560000000000)
    pts(26) = Array(48.954564230708, 0.582400000000)
    pts(27) = Array(48.954097449587, 0.604800000000)
    pts(28) = Array(48.953613045997, 0.627200000000)
    pts(29) = Array(48.953111019039, 0.649600000000)
    pts(30) = Array(48.952591367780, 0.672000000000)
    pts(31) = Array(48.952054091257, 0.694400000000)
    pts(32) = Array(48.951499188471, 0.716800000000)
    pts(33) = Array(48.950926658393, 0.739200000000)
    pts(34) = Array(48.950336499959, 0.761600000000)
    pts(35) = Array(48.949728712074, 0.784000000000)
    pts(36) = Array(48.949103293609, 0.806400000000)
    pts(37) = Array(48.948460243402, 0.828800000000)
    pts(38) = Array(48.947799560260, 0.851200000000)
    pts(39) = Array(48.947121242954, 0.873600000000)
    pts(40) = Array(48.946425290226, 0.896000000000)
    pts(41) = Array(48.945711700781, 0.918400000000)
    pts(42) = Array(48.944980473295, 0.940800000000)
    pts(43) = Array(48.944231606407, 0.963200000000)
    pts(44) = Array(48.943465098727, 0.985600000000)
    pts(45) = Array(48.942680948830, 1.008000000000)
    pts(46) = Array(48.941879155259, 1.030400000000)
    pts(47) = Array(48.941059716522, 1.052800000000)
    pts(48) = Array(48.940222631096, 1.075200000000)
    pts(49) = Array(48.939367897425, 1.097600000000)
    pts(50) = Array(48.938495513919, 1.120000000000)
    pts(51) = Array(48.937605478956, 1.142400000000)
    pts(52) = Array(48.936697790880, 1.164800000000)
    pts(53) = Array(48.935772448002, 1.187200000000)
    pts(54) = Array(48.934829448601, 1.209600000000)
    pts(55) = Array(48.933868790922, 1.232000000000)
    pts(56) = Array(48.932890473178, 1.254400000000)
    pts(57) = Array(48.931894493546, 1.276800000000)
    pts(58) = Array(48.930880850173, 1.299200000000)
    pts(59) = Array(48.929849541172, 1.321600000000)
    pts(60) = Array(48.928800564623, 1.344000000000)
    pts(61) = Array(48.927733918571, 1.366400000000)
    pts(62) = Array(48.926649601029, 1.388800000000)
    pts(63) = Array(48.925547609979, 1.411200000000)
    pts(64) = Array(48.924427943366, 1.433600000000)
    pts(65) = Array(48.923290599104, 1.456000000000)
    pts(66) = Array(48.922135575074, 1.478400000000)
    pts(67) = Array(48.920962869121, 1.500800000000)
    pts(68) = Array(48.919772479060, 1.523200000000)
    pts(69) = Array(48.918564402671, 1.545600000000)
    pts(70) = Array(48.917338637700, 1.568000000000)
    pts(71) = Array(48.916095181861, 1.590400000000)
    pts(72) = Array(48.914834032835, 1.612800000000)
    pts(73) = Array(48.913555188267, 1.635200000000)
    pts(74) = Array(48.912258645772, 1.657600000000)
    pts(75) = Array(48.910944402928, 1.680000000000)
    pts(76) = Array(48.909612457282, 1.702400000000)
    pts(77) = Array(48.908262806347, 1.724800000000)
    pts(78) = Array(48.906895447602, 1.747200000000)
    pts(79) = Array(48.905510378493, 1.769600000000)
    pts(80) = Array(48.904107596431, 1.792000000000)
    pts(81) = Array(48.902687098795, 1.814400000000)
    pts(82) = Array(48.901248882929, 1.836800000000)
    pts(83) = Array(48.899792946146, 1.859200000000)
    pts(84) = Array(48.898319285722, 1.881600000000)
    pts(85) = Array(48.896827898901, 1.904000000000)
    pts(86) = Array(48.895318782894, 1.926400000000)
    pts(87) = Array(48.893791934875, 1.948800000000)
    pts(88) = Array(48.892247351989, 1.971200000000)
    pts(89) = Array(48.890685031343, 1.993600000000)
    pts(90) = Array(48.889104970012, 2.016000000000)
    pts(91) = Array(48.887507165037, 2.038400000000)
    pts(92) = Array(48.885891613426, 2.060800000000)
    pts(93) = Array(48.884258312150, 2.083200000000)
    pts(94) = Array(48.882607258150, 2.105600000000)
    pts(95) = Array(48.880938448331, 2.128000000000)
    pts(96) = Array(48.879251879563, 2.150400000000)
    pts(97) = Array(48.877547548684, 2.172800000000)
    pts(98) = Array(48.875825452497, 2.195200000000)
    pts(99) = Array(48.874085587771, 2.217600000000)
    pts(100) = Array(48.872327951240, 2.240000000000)
    pts(101) = Array(48.870552539605, 2.262400000000)
    pts(102) = Array(48.868759349533, 2.284800000000)
    pts(103) = Array(48.866948377655, 2.307200000000)
    pts(104) = Array(48.865119620570, 2.329600000000)
    pts(105) = Array(48.863273074842, 2.352000000000)
    pts(106) = Array(48.861408736999, 2.374400000000)
    pts(107) = Array(48.859526603537, 2.396800000000)
    pts(108) = Array(48.857626670916, 2.419200000000)
    pts(109) = Array(48.855708935562, 2.441600000000)
    pts(110) = Array(48.853773393867, 2.464000000000)
    pts(111) = Array(48.851820042188, 2.486400000000)
    pts(112) = Array(48.849848876849, 2.508800000000)
    pts(113) = Array(48.847859894136, 2.531200000000)
    pts(114) = Array(48.845853090304, 2.553600000000)
    pts(115) = Array(48.843828461571, 2.576000000000)
    pts(116) = Array(48.841786004122, 2.598400000000)
    pts(117) = Array(48.839725714106, 2.620800000000)
    pts(118) = Array(48.837647587638, 2.643200000000)
    pts(119) = Array(48.835551620797, 2.665600000000)
    pts(120) = Array(48.833437809630, 2.688000000000)
    pts(121) = Array(48.831306150146, 2.710400000000)
    pts(122) = Array(48.829156638321, 2.732800000000)
    pts(123) = Array(48.826989270095, 2.755200000000)
    pts(124) = Array(48.824804041373, 2.777600000000)
    pts(125) = Array(48.822600948027, 2.800000000000)
    pts(126) = Array(48.820379985892, 2.822400000000)
    pts(127) = Array(48.818141150768, 2.844800000000)
    pts(128) = Array(48.815884438420, 2.867200000000)
    pts(129) = Array(48.813609844579, 2.889600000000)
    pts(130) = Array(48.811317364939, 2.912000000000)
    pts(131) = Array(48.809006995160, 2.934400000000)
    pts(132) = Array(48.806678730867, 2.956800000000)
    pts(133) = Array(48.804332567647, 2.979200000000)
    pts(134) = Array(48.801968501056, 3.001600000000)
    pts(135) = Array(48.799586526610, 3.024000000000)
    pts(136) = Array(48.797186639793, 3.046400000000)
    pts(137) = Array(48.794768836053, 3.068800000000)
    pts(138) = Array(48.792333110800, 3.091200000000)
    pts(139) = Array(48.789879459411, 3.113600000000)
    pts(140) = Array(48.787407877226, 3.136000000000)
    pts(141) = Array(48.784918359550, 3.158400000000)
    pts(142) = Array(48.782410901652, 3.180800000000)
    pts(143) = Array(48.779885498766, 3.203200000000)
    pts(144) = Array(48.777342146088, 3.225600000000)
    pts(145) = Array(48.774780838781, 3.248000000000)
    pts(146) = Array(48.772201571969, 3.270400000000)
    pts(147) = Array(48.769604340743, 3.292800000000)
    pts(148) = Array(48.766989140156, 3.315200000000)
    pts(149) = Array(48.764355965227, 3.337600000000)
    pts(150) = Array(48.761704810936, 3.360000000000)
    pts(151) = Array(48.759035672228, 3.382400000000)
    pts(152) = Array(48.756348544014, 3.404800000000)
    pts(153) = Array(48.753643421165, 3.427200000000)
    pts(154) = Array(48.750920298520, 3.449600000000)
    pts(155) = Array(48.748179170876, 3.472000000000)
    pts(156) = Array(48.745420033000, 3.494400000000)
    pts(157) = Array(48.742642879617, 3.516800000000)
    pts(158) = Array(48.739847705420, 3.539200000000)
    pts(159) = Array(48.737034505062, 3.561600000000)
    pts(160) = Array(48.734203273160, 3.584000000000)
    pts(161) = Array(48.731354004297, 3.606400000000)
    pts(162) = Array(48.728486693016, 3.628800000000)
    pts(163) = Array(48.725601333825, 3.651200000000)
    pts(164) = Array(48.722697921195, 3.673600000000)
    pts(165) = Array(48.719776449560, 3.696000000000)
    pts(166) = Array(48.716836913316, 3.718400000000)
    pts(167) = Array(48.713879306824, 3.740800000000)
    pts(168) = Array(48.710903624406, 3.763200000000)
    pts(169) = Array(48.707909860348, 3.785600000000)
    pts(170) = Array(48.704898008900, 3.808000000000)
    pts(171) = Array(48.701868064272, 3.830400000000)
    pts(172) = Array(48.698820020639, 3.852800000000)
    pts(173) = Array(48.695753872137, 3.875200000000)
    pts(174) = Array(48.692669612866, 3.897600000000)
    pts(175) = Array(48.689567236889, 3.920000000000)
    pts(176) = Array(48.686446738229, 3.942400000000)
    pts(177) = Array(48.683308110873, 3.964800000000)
    pts(178) = Array(48.680151348772, 3.987200000000)
    pts(179) = Array(48.676976445837, 4.009600000000)
    pts(180) = Array(48.673783395941, 4.032000000000)
    pts(181) = Array(48.670572192922, 4.054400000000)
    pts(182) = Array(48.667342830576, 4.076800000000)
    pts(183) = Array(48.664095302665, 4.099200000000)
    pts(184) = Array(48.660829602912, 4.121600000000)
    pts(185) = Array(48.657545724999, 4.144000000000)
    pts(186) = Array(48.654243662575, 4.166400000000)
    pts(187) = Array(48.650923409246, 4.188800000000)
    pts(188) = Array(48.647584958583, 4.211200000000)
    pts(189) = Array(48.644228304117, 4.233600000000)
    pts(190) = Array(48.640853439342, 4.256000000000)
    pts(191) = Array(48.637460357712, 4.278400000000)
    pts(192) = Array(48.634049052643, 4.300800000000)
    pts(193) = Array(48.630619517513, 4.323200000000)
    pts(194) = Array(48.627171745662, 4.345600000000)
    pts(195) = Array(48.623705730390, 4.368000000000)
    pts(196) = Array(48.620221464958, 4.390400000000)
    pts(197) = Array(48.616718942589, 4.412800000000)
    pts(198) = Array(48.613198156468, 4.435200000000)
    pts(199) = Array(48.609659099738, 4.457600000000)
    pts(200) = Array(48.606101765507, 4.480000000000)
    pts(201) = Array(48.602526146841, 4.502400000000)
    pts(202) = Array(48.598932236767, 4.524800000000)
    pts(203) = Array(48.595320028274, 4.547200000000)
    pts(204) = Array(48.591689514312, 4.569600000000)
    pts(205) = Array(48.588040687789, 4.592000000000)
    pts(206) = Array(48.584373541576, 4.614400000000)
    pts(207) = Array(48.580688068505, 4.636800000000)
    pts(208) = Array(48.576984261365, 4.659200000000)
    pts(209) = Array(48.573262112909, 4.681600000000)
    pts(210) = Array(48.569521615848, 4.704000000000)
    pts(211) = Array(48.565762762855, 4.726400000000)
    pts(212) = Array(48.561985546561, 4.748800000000)
    pts(213) = Array(48.558189959558, 4.771200000000)
    pts(214) = Array(48.554375994399, 4.793600000000)
    pts(215) = Array(48.550543643594, 4.816000000000)
    pts(216) = Array(48.546692899617, 4.838400000000)
    pts(217) = Array(48.542823754898, 4.860800000000)
    pts(218) = Array(48.538936201829, 4.883200000000)
    pts(219) = Array(48.535030232760, 4.905600000000)
    pts(220) = Array(48.531105840001, 4.928000000000)
    pts(221) = Array(48.527163015822, 4.950400000000)
    pts(222) = Array(48.523201752452, 4.972800000000)
    pts(223) = Array(48.519222042079, 4.995200000000)
    pts(224) = Array(48.515223876850, 5.017600000000)
    pts(225) = Array(48.511207248872, 5.040000000000)
    pts(226) = Array(48.507172150210, 5.062400000000)
    pts(227) = Array(48.503118572888, 5.084800000000)
    pts(228) = Array(48.499046508891, 5.107200000000)
    pts(229) = Array(48.494955950159, 5.129600000000)
    pts(230) = Array(48.490846888594, 5.152000000000)
    pts(231) = Array(48.486719316054, 5.174400000000)
    pts(232) = Array(48.482573224358, 5.196800000000)
    pts(233) = Array(48.478408605282, 5.219200000000)
    pts(234) = Array(48.474225450561, 5.241600000000)
    pts(235) = Array(48.470023751887, 5.264000000000)
    pts(236) = Array(48.465803500911, 5.286400000000)
    pts(237) = Array(48.461564689243, 5.308800000000)
    pts(238) = Array(48.457307308449, 5.331200000000)
    pts(239) = Array(48.453031350055, 5.353600000000)
    pts(240) = Array(48.448736805544, 5.376000000000)
    pts(241) = Array(48.444423666356, 5.398400000000)
    pts(242) = Array(48.440091923889, 5.420800000000)
    pts(243) = Array(48.435741569499, 5.443200000000)
    pts(244) = Array(48.431372594500, 5.465600000000)
    pts(245) = Array(48.426984990162, 5.488000000000)
    pts(246) = Array(48.422578747714, 5.510400000000)
    pts(247) = Array(48.418153858339, 5.532800000000)
    pts(248) = Array(48.413710313181, 5.555200000000)
    pts(249) = Array(48.409248103339, 5.577600000000)
    pts(250) = Array(48.404767219869, 5.600000000000)
    pts(251) = Array(48.400267653783, 5.622400000000)
    pts(252) = Array(48.395749396053, 5.644800000000)
    pts(253) = Array(48.391212437603, 5.667200000000)
    pts(254) = Array(48.386656769317, 5.689600000000)
    pts(255) = Array(48.382082382034, 5.712000000000)
    pts(256) = Array(48.377489266550, 5.734400000000)
    pts(257) = Array(48.372877413617, 5.756800000000)
    pts(258) = Array(48.368246813942, 5.779200000000)
    pts(259) = Array(48.363597458190, 5.801600000000)
    pts(260) = Array(48.358929336982, 5.824000000000)
    pts(261) = Array(48.354242440892, 5.846400000000)
    pts(262) = Array(48.349536760452, 5.868800000000)
    pts(263) = Array(48.344812286150, 5.891200000000)
    pts(264) = Array(48.340069008428, 5.913600000000)
    pts(265) = Array(48.335306917685, 5.936000000000)
    pts(266) = Array(48.330526004275, 5.958400000000)
    pts(267) = Array(48.325726258506, 5.980800000000)
    pts(268) = Array(48.320907670642, 6.003200000000)
    pts(269) = Array(48.316070230902, 6.025600000000)
    pts(270) = Array(48.311213929460, 6.048000000000)
    pts(271) = Array(48.306338756445, 6.070400000000)
    pts(272) = Array(48.301444701941, 6.092800000000)
    pts(273) = Array(48.296531755984, 6.115200000000)
    pts(274) = Array(48.291599908569, 6.137600000000)
    pts(275) = Array(48.286649149641, 6.160000000000)
    pts(276) = Array(48.281679469103, 6.182400000000)
    pts(277) = Array(48.276690856809, 6.204800000000)
    pts(278) = Array(48.271683302569, 6.227200000000)
    pts(279) = Array(48.266656796146, 6.249600000000)
    pts(280) = Array(48.261611327258, 6.272000000000)
    pts(281) = Array(48.256546885575, 6.294400000000)
    pts(282) = Array(48.251463460722, 6.316800000000)
    pts(283) = Array(48.246361042278, 6.339200000000)
    pts(284) = Array(48.241239619773, 6.361600000000)
    pts(285) = Array(48.236099182692, 6.384000000000)
    pts(286) = Array(48.230939720474, 6.406400000000)
    pts(287) = Array(48.225761222509, 6.428800000000)
    pts(288) = Array(48.220563678141, 6.451200000000)
    pts(289) = Array(48.215347076666, 6.473600000000)
    pts(290) = Array(48.210111407335, 6.496000000000)
    pts(291) = Array(48.204856659349, 6.518400000000)
    pts(292) = Array(48.199582821862, 6.540800000000)
    pts(293) = Array(48.194289883981, 6.563200000000)
    pts(294) = Array(48.188977834765, 6.585600000000)
    pts(295) = Array(48.183646663225, 6.608000000000)
    pts(296) = Array(48.178296358324, 6.630400000000)
    pts(297) = Array(48.172926908977, 6.652800000000)
    pts(298) = Array(48.167538304051, 6.675200000000)
    pts(299) = Array(48.162130532363, 6.697600000000)
    pts(300) = Array(48.156703582683, 6.720000000000)
    pts(301) = Array(48.151257443733, 6.742400000000)
    pts(302) = Array(48.145792104184, 6.764800000000)
    pts(303) = Array(48.140307552659, 6.787200000000)
    pts(304) = Array(48.134803777734, 6.809600000000)
    pts(305) = Array(48.129280767932, 6.832000000000)
    pts(306) = Array(48.123738511731, 6.854400000000)
    pts(307) = Array(48.118176997556, 6.876800000000)
    pts(308) = Array(48.112596213784, 6.899200000000)
    pts(309) = Array(48.106996148743, 6.921600000000)
    pts(310) = Array(48.101376790710, 6.944000000000)
    pts(311) = Array(48.095738127912, 6.966400000000)
    pts(312) = Array(48.090080148526, 6.988800000000)
    pts(313) = Array(48.084402840681, 7.011200000000)
    pts(314) = Array(48.078706192452, 7.033600000000)
    pts(315) = Array(48.072990191867, 7.056000000000)
    pts(316) = Array(48.067254826900, 7.078400000000)
    pts(317) = Array(48.061500085477, 7.100800000000)
    pts(318) = Array(48.055725955472, 7.123200000000)
    pts(319) = Array(48.049932424707, 7.145600000000)
    pts(320) = Array(48.044119480956, 7.168000000000)
    pts(321) = Array(48.038287111939, 7.190400000000)
    pts(322) = Array(48.032435305324, 7.212800000000)
    pts(323) = Array(48.026564048729, 7.235200000000)
    pts(324) = Array(48.020673329720, 7.257600000000)
    pts(325) = Array(48.014763135812, 7.280000000000)
    pts(326) = Array(48.008833454466, 7.302400000000)
    pts(327) = Array(48.002884273092, 7.324800000000)
    pts(328) = Array(47.996915579047, 7.347200000000)
    pts(329) = Array(47.990927359636, 7.369600000000)
    pts(330) = Array(47.984919602113, 7.392000000000)
    pts(331) = Array(47.978892293676, 7.414400000000)
    pts(332) = Array(47.972845421472, 7.436800000000)
    pts(333) = Array(47.966778972596, 7.459200000000)
    pts(334) = Array(47.960692934088, 7.481600000000)
    pts(335) = Array(47.956382946514, 7.497419679825)
    pts(336) = Array(47.956046882757, 7.498651685025)
    pts(337) = Array(47.955745702507, 7.499755623002)
    pts(338) = Array(47.955745702507, 7.504000000000)
    pts(339) = Array(47.955745702507, 7.526400000000)
    pts(340) = Array(47.955745702507, 7.548800000000)
    pts(341) = Array(47.955745702507, 7.571200000000)
    pts(342) = Array(47.955745702507, 7.593600000000)
    pts(343) = Array(47.955745702507, 7.616000000000)
    pts(344) = Array(47.955745702507, 7.638400000000)
    pts(345) = Array(47.955745702507, 7.660800000000)
    pts(346) = Array(47.955745702507, 7.683200000000)
    pts(347) = Array(47.955745702507, 7.705600000000)
    pts(348) = Array(47.955745702507, 7.728000000000)
    pts(349) = Array(47.955745702507, 7.750400000000)
    pts(350) = Array(47.955745702507, 7.772800000000)
    pts(351) = Array(47.955745702507, 7.795200000000)
    pts(352) = Array(47.955745702507, 7.817600000000)
    pts(353) = Array(47.955745702507, 7.840000000000)
    pts(354) = Array(47.955745702507, 7.862400000000)
    pts(355) = Array(47.955745702507, 7.884800000000)
    pts(356) = Array(47.955745702507, 7.907200000000)
    pts(357) = Array(47.955745702507, 7.929600000000)
    pts(358) = Array(47.955745702507, 7.952000000000)
    pts(359) = Array(47.955745702507, 7.974400000000)
    pts(360) = Array(47.955745702507, 7.996800000000)
    pts(361) = Array(47.955745702507, 8.019200000000)
    pts(362) = Array(47.955745702507, 8.041600000000)
    pts(363) = Array(47.955745702507, 8.064000000000)
    pts(364) = Array(47.955745702507, 8.086400000000)
    pts(365) = Array(47.955745702507, 8.108800000000)
    pts(366) = Array(47.955745702507, 8.131200000000)
    pts(367) = Array(47.955745702507, 8.153600000000)
    pts(368) = Array(47.955745702507, 8.176000000000)
    pts(369) = Array(47.955745702507, 8.198400000000)
    pts(370) = Array(47.955745702507, 8.220800000000)
    pts(371) = Array(47.955745702507, 8.243200000000)
    pts(372) = Array(47.955745702507, 8.265600000000)
    pts(373) = Array(47.955745702507, 8.288000000000)
    pts(374) = Array(47.955745702507, 8.310400000000)
    pts(375) = Array(47.955745702507, 8.332800000000)
    pts(376) = Array(47.955745702507, 8.355200000000)
    pts(377) = Array(47.955745702507, 8.377600000000)
    pts(378) = Array(47.955745702507, 8.400000000000)
    pts(379) = Array(47.955745702507, 8.422400000000)
    pts(380) = Array(47.955745702507, 8.444800000000)
    pts(381) = Array(47.955745702507, 8.467200000000)
    pts(382) = Array(47.955745702507, 8.489600000000)
    pts(383) = Array(47.955745702507, 8.512000000000)
    pts(384) = Array(47.955745702507, 8.534400000000)
    pts(385) = Array(47.955745702507, 8.556800000000)
    pts(386) = Array(47.955745702507, 8.579200000000)
    pts(387) = Array(47.955745702507, 8.601600000000)
    pts(388) = Array(47.955745702507, 8.624000000000)
    pts(389) = Array(47.955745702507, 8.646400000000)
    pts(390) = Array(47.955745702507, 8.668800000000)
    pts(391) = Array(47.955745702507, 8.691200000000)
    pts(392) = Array(47.955745702507, 8.713600000000)
    pts(393) = Array(47.955745702507, 8.736000000000)
    pts(394) = Array(47.955745702507, 8.758400000000)
    pts(395) = Array(47.955745702507, 8.780800000000)
    pts(396) = Array(47.955745702507, 8.803200000000)
    pts(397) = Array(47.955745702507, 8.825600000000)
    pts(398) = Array(47.955745702507, 8.848000000000)
    pts(399) = Array(47.955745702507, 8.870400000000)
    pts(400) = Array(47.955745702507, 8.892800000000)
    pts(401) = Array(47.955745702507, 8.915200000000)
    pts(402) = Array(47.955745702507, 8.937600000000)
    pts(403) = Array(47.955745702507, 8.960000000000)
    pts(404) = Array(47.955745702507, 8.982400000000)
    pts(405) = Array(47.955745702507, 9.004800000000)
    pts(406) = Array(47.955745702507, 9.027200000000)
    pts(407) = Array(47.955745702507, 9.049600000000)
    pts(408) = Array(47.955745702507, 9.072000000000)
    pts(409) = Array(47.955745702507, 9.094400000000)
    pts(410) = Array(47.955745702507, 9.116800000000)
    pts(411) = Array(47.955745702507, 9.139200000000)
    pts(412) = Array(47.955745702507, 9.161600000000)
    pts(413) = Array(47.955745702507, 9.184000000000)
    pts(414) = Array(47.955745702507, 9.206400000000)
    pts(415) = Array(47.955745702507, 9.228800000000)
    pts(416) = Array(47.955745702507, 9.251200000000)
    pts(417) = Array(47.955745702507, 9.273600000000)
    pts(418) = Array(47.955745702507, 9.296000000000)
    pts(419) = Array(47.955745702507, 9.318400000000)
    pts(420) = Array(47.955745702507, 9.340800000000)
    pts(421) = Array(47.955745702507, 9.363200000000)
    pts(422) = Array(47.955745702507, 9.385600000000)
    pts(423) = Array(47.955745702507, 9.408000000000)
    pts(424) = Array(47.955745702507, 9.430400000000)
    pts(425) = Array(47.955745702507, 9.452800000000)
    pts(426) = Array(47.955745702507, 9.475200000000)
    pts(427) = Array(47.955745702507, 9.497600000000)
    pts(428) = Array(47.955745702507, 9.520000000000)
    pts(429) = Array(47.955745702507, 9.542400000000)
    pts(430) = Array(47.955745702507, 9.564800000000)
    pts(431) = Array(47.955745702507, 9.587200000000)
    pts(432) = Array(47.955745702507, 9.609600000000)
    pts(433) = Array(47.955745702507, 9.632000000000)
    pts(434) = Array(47.955745702507, 9.654400000000)
    pts(435) = Array(47.955745702507, 9.676800000000)
    pts(436) = Array(47.955745702507, 9.699200000000)
    pts(437) = Array(47.955745702507, 9.721600000000)
    pts(438) = Array(47.955745702507, 9.744000000000)
    pts(439) = Array(47.955745702507, 9.766400000000)
    pts(440) = Array(47.955745702507, 9.788800000000)
    pts(441) = Array(47.955745702507, 9.811200000000)
    pts(442) = Array(47.955745702507, 9.833600000000)
    pts(443) = Array(47.955745702507, 9.856000000000)
    pts(444) = Array(47.955745702507, 9.878400000000)
    pts(445) = Array(47.955745702507, 9.900800000000)
    pts(446) = Array(47.955745702507, 9.923200000000)
    pts(447) = Array(47.955745702507, 9.945600000000)
    pts(448) = Array(47.955745702507, 9.968000000000)
    pts(449) = Array(47.955745702507, 9.990400000000)
    pts(450) = Array(47.955745702507, 10.012800000000)
    pts(451) = Array(47.955745702507, 10.035200000000)
    pts(452) = Array(47.955745702507, 10.057600000000)
    pts(453) = Array(47.955745702507, 10.080000000000)
    pts(454) = Array(47.955745702507, 10.102400000000)
    pts(455) = Array(47.955745702507, 10.124800000000)
    pts(456) = Array(47.955745702507, 10.147200000000)
    pts(457) = Array(47.955745702507, 10.169600000000)
    pts(458) = Array(47.955745702507, 10.192000000000)
    pts(459) = Array(47.955745702507, 10.214400000000)
    pts(460) = Array(47.955745702507, 10.236800000000)
    pts(461) = Array(47.955745702507, 10.259200000000)
    pts(462) = Array(47.955745702507, 10.281600000000)
    pts(463) = Array(47.955745702507, 10.304000000000)
    pts(464) = Array(47.955745702507, 10.326400000000)
    pts(465) = Array(47.955745702507, 10.348800000000)
    pts(466) = Array(47.955745702507, 10.371200000000)
    pts(467) = Array(47.955745702507, 10.393600000000)
    pts(468) = Array(47.955745702507, 10.416000000000)
    pts(469) = Array(47.955745702507, 10.438400000000)
    pts(470) = Array(47.955745702507, 10.460800000000)
    pts(471) = Array(47.955745702507, 10.483200000000)
    pts(472) = Array(47.955745702507, 10.500000000000)
    pts(473) = Array(47.955745702507, 10.505600000000)
    pts(474) = Array(47.955745702507, 10.528000000000)
    pts(475) = Array(47.955745702507, 10.550000000000)
    pts(476) = Array(47.955657093223, 10.550400000000)
    pts(477) = Array(47.950694973316, 10.572800000000)
    pts(478) = Array(47.945732853409, 10.595200000000)
    pts(479) = Array(47.940770733502, 10.617600000000)
    pts(480) = Array(47.935808613595, 10.640000000000)
    pts(481) = Array(47.930846493688, 10.662400000000)
    pts(482) = Array(47.925884373781, 10.684800000000)
    pts(483) = Array(47.922517220987, 10.700000000000)
    pts(484) = Array(47.922517220987, 10.707200000000)
    pts(485) = Array(47.922517220987, 10.729600000000)
    pts(486) = Array(47.922517220987, 10.752000000000)
    pts(487) = Array(47.922517220987, 10.774400000000)
    pts(488) = Array(47.922517220987, 10.796800000000)
    pts(489) = Array(47.922517220987, 10.819200000000)
    pts(490) = Array(47.922517220987, 10.841600000000)
    pts(491) = Array(47.922517220987, 10.864000000000)
    pts(492) = Array(47.922517220987, 10.886400000000)
    pts(493) = Array(47.922517220987, 10.908800000000)
    pts(494) = Array(47.922517220987, 10.931200000000)
    pts(495) = Array(47.922517220987, 10.953600000000)
    pts(496) = Array(47.922517220987, 10.976000000000)
    pts(497) = Array(47.922517220987, 10.998400000000)
    pts(498) = Array(47.922517220987, 11.020800000000)
    pts(499) = Array(47.922517220987, 11.043200000000)
    pts(500) = Array(47.922517220987, 11.065600000000)
    pts(501) = Array(47.922517220987, 11.088000000000)
    pts(502) = Array(47.922517220987, 11.110400000000)
    pts(503) = Array(47.922517220987, 11.132800000000)
    pts(504) = Array(47.922517220987, 11.155200000000)
    pts(505) = Array(47.922517220987, 11.177600000000)
    pts(506) = Array(47.922517220987, 11.200000000000)
    pts(507) = Array(55.681899312027, 11.200000000000)
    pts(508) = Array(55.681899312027, 11.177600000000)
    pts(509) = Array(55.681899312027, 11.155200000000)
    pts(510) = Array(55.681899312027, 11.132800000000)
    pts(511) = Array(55.681899312027, 11.110400000000)
    pts(512) = Array(55.681899312027, 11.088000000000)
    pts(513) = Array(55.681899312027, 11.065600000000)
    pts(514) = Array(55.681899312027, 11.043200000000)
    pts(515) = Array(55.681899312027, 11.020800000000)
    pts(516) = Array(55.681899312027, 10.998400000000)
    pts(517) = Array(55.681899312027, 10.976000000000)
    pts(518) = Array(55.681899312027, 10.953600000000)
    pts(519) = Array(55.681899312027, 10.931200000000)
    pts(520) = Array(55.681899312027, 10.908800000000)
    pts(521) = Array(55.681899312027, 10.886400000000)
    pts(522) = Array(55.681899312027, 10.864000000000)
    pts(523) = Array(55.681899312027, 10.841600000000)
    pts(524) = Array(55.681899312027, 10.819200000000)
    pts(525) = Array(55.681899312027, 10.796800000000)
    pts(526) = Array(55.681899312027, 10.774400000000)
    pts(527) = Array(55.681899312027, 10.752000000000)
    pts(528) = Array(55.681899312027, 10.729600000000)
    pts(529) = Array(55.681899312027, 10.707200000000)
    pts(530) = Array(55.681899312027, 10.700000000000)
    pts(531) = Array(55.681899312027, 10.684800000000)
    pts(532) = Array(55.681899312027, 10.662400000000)
    pts(533) = Array(55.681899312027, 10.640000000000)
    pts(534) = Array(55.681899312027, 10.617600000000)
    pts(535) = Array(55.681899312027, 10.595200000000)
    pts(536) = Array(55.681899312027, 10.572800000000)
    pts(537) = Array(55.681899312027, 10.550400000000)
    pts(538) = Array(55.681899312027, 10.550000000000)
    pts(539) = Array(55.681899312027, 10.528000000000)
    pts(540) = Array(55.681899312027, 10.505600000000)
    pts(541) = Array(55.681899312027, 10.500000000000)
    pts(542) = Array(55.674846740682, 10.483200000000)
    pts(543) = Array(55.665463964496, 10.460800000000)
    pts(544) = Array(55.656104766564, 10.438400000000)
    pts(545) = Array(55.646769120433, 10.416000000000)
    pts(546) = Array(55.637456999752, 10.393600000000)
    pts(547) = Array(55.628168378277, 10.371200000000)
    pts(548) = Array(55.618903229869, 10.348800000000)
    pts(549) = Array(55.609661528494, 10.326400000000)
    pts(550) = Array(55.600443248221, 10.304000000000)
    pts(551) = Array(55.591248363224, 10.281600000000)
    pts(552) = Array(55.582076847780, 10.259200000000)
    pts(553) = Array(55.572928676269, 10.236800000000)
    pts(554) = Array(55.563803823173, 10.214400000000)
    pts(555) = Array(55.554702263077, 10.192000000000)
    pts(556) = Array(55.545623970667, 10.169600000000)
    pts(557) = Array(55.536568920730, 10.147200000000)
    pts(558) = Array(55.527537088155, 10.124800000000)
    pts(559) = Array(55.518528447930, 10.102400000000)
    pts(560) = Array(55.509542975144, 10.080000000000)
    pts(561) = Array(55.500580644985, 10.057600000000)
    pts(562) = Array(55.491641432740, 10.035200000000)
    pts(563) = Array(55.482725313796, 10.012800000000)
    pts(564) = Array(55.473832263636, 9.990400000000)
    pts(565) = Array(55.464962257845, 9.968000000000)
    pts(566) = Array(55.456115272101, 9.945600000000)
    pts(567) = Array(55.447291282183, 9.923200000000)
    pts(568) = Array(55.438490263963, 9.900800000000)
    pts(569) = Array(55.429712193414, 9.878400000000)
    pts(570) = Array(55.420957046600, 9.856000000000)
    pts(571) = Array(55.412224799686, 9.833600000000)
    pts(572) = Array(55.403515428927, 9.811200000000)
    pts(573) = Array(55.394828910676, 9.788800000000)
    pts(574) = Array(55.386165221381, 9.766400000000)
    pts(575) = Array(55.377524337582, 9.744000000000)
    pts(576) = Array(55.368906235914, 9.721600000000)
    pts(577) = Array(55.360310893105, 9.699200000000)
    pts(578) = Array(55.351738285976, 9.676800000000)
    pts(579) = Array(55.343188391441, 9.654400000000)
    pts(580) = Array(55.334661186505, 9.632000000000)
    pts(581) = Array(55.326156648268, 9.609600000000)
    pts(582) = Array(55.317674753919, 9.587200000000)
    pts(583) = Array(55.309215480738, 9.564800000000)
    pts(584) = Array(55.300778806096, 9.542400000000)
    pts(585) = Array(55.292364707457, 9.520000000000)
    pts(586) = Array(55.283973162372, 9.497600000000)
    pts(587) = Array(55.275604148484, 9.475200000000)
    pts(588) = Array(55.267257643524, 9.452800000000)
    pts(589) = Array(55.258933625312, 9.430400000000)
    pts(590) = Array(55.250632071758, 9.408000000000)
    pts(591) = Array(55.242352960860, 9.385600000000)
    pts(592) = Array(55.234096270704, 9.363200000000)
    pts(593) = Array(55.225861979464, 9.340800000000)
    pts(594) = Array(55.217650065400, 9.318400000000)
    pts(595) = Array(55.209460506860, 9.296000000000)
    pts(596) = Array(55.201293282281, 9.273600000000)
    pts(597) = Array(55.193148370182, 9.251200000000)
    pts(598) = Array(55.185025749172, 9.228800000000)
    pts(599) = Array(55.176925397943, 9.206400000000)
    pts(600) = Array(55.168847295275, 9.184000000000)
    pts(601) = Array(55.160791420031, 9.161600000000)
    pts(602) = Array(55.152757751159, 9.139200000000)
    pts(603) = Array(55.144746267694, 9.116800000000)
    pts(604) = Array(55.136756948752, 9.094400000000)
    pts(605) = Array(55.128789773534, 9.072000000000)
    pts(606) = Array(55.120844721325, 9.049600000000)
    pts(607) = Array(55.112921771493, 9.027200000000)
    pts(608) = Array(55.105020903490, 9.004800000000)
    pts(609) = Array(55.097142096847, 8.982400000000)
    pts(610) = Array(55.089285331182, 8.960000000000)
    pts(611) = Array(55.081450586192, 8.937600000000)
    pts(612) = Array(55.073637841657, 8.915200000000)
    pts(613) = Array(55.065847077437, 8.892800000000)
    pts(614) = Array(55.058078273476, 8.870400000000)
    pts(615) = Array(55.050331409796, 8.848000000000)
    pts(616) = Array(55.042606466499, 8.825600000000)
    pts(617) = Array(55.034903423772, 8.803200000000)
    pts(618) = Array(55.027222261876, 8.780800000000)
    pts(619) = Array(55.019562961155, 8.758400000000)
    pts(620) = Array(55.011925502033, 8.736000000000)
    pts(621) = Array(55.004309865011, 8.713600000000)
    pts(622) = Array(54.996716030669, 8.691200000000)
    pts(623) = Array(54.989143979667, 8.668800000000)
    pts(624) = Array(54.981593692743, 8.646400000000)
    pts(625) = Array(54.974065150711, 8.624000000000)
    pts(626) = Array(54.966558334466, 8.601600000000)
    pts(627) = Array(54.959073224977, 8.579200000000)
    pts(628) = Array(54.951609803292, 8.556800000000)
    pts(629) = Array(54.944168050536, 8.534400000000)
    pts(630) = Array(54.936747947910, 8.512000000000)
    pts(631) = Array(54.929349476691, 8.489600000000)
    pts(632) = Array(54.921972618234, 8.467200000000)
    pts(633) = Array(54.914617353967, 8.444800000000)
    pts(634) = Array(54.907283665395, 8.422400000000)
    pts(635) = Array(54.899971534099, 8.400000000000)
    pts(636) = Array(54.892680941734, 8.377600000000)
    pts(637) = Array(54.885411870030, 8.355200000000)
    pts(638) = Array(54.878164300790, 8.332800000000)
    pts(639) = Array(54.870938215895, 8.310400000000)
    pts(640) = Array(54.863733597297, 8.288000000000)
    pts(641) = Array(54.856550427021, 8.265600000000)
    pts(642) = Array(54.849388687168, 8.243200000000)
    pts(643) = Array(54.842248359912, 8.220800000000)
    pts(644) = Array(54.835129427498, 8.198400000000)
    pts(645) = Array(54.828031872244, 8.176000000000)
    pts(646) = Array(54.820955676543, 8.153600000000)
    pts(647) = Array(54.813900822858, 8.131200000000)
    pts(648) = Array(54.806867293725, 8.108800000000)
    pts(649) = Array(54.799855071750, 8.086400000000)
    pts(650) = Array(54.792864139612, 8.064000000000)
    pts(651) = Array(54.785894480063, 8.041600000000)
    pts(652) = Array(54.778946075922, 8.019200000000)
    pts(653) = Array(54.772018910081, 7.996800000000)
    pts(654) = Array(54.765112965504, 7.974400000000)
    pts(655) = Array(54.758228225223, 7.952000000000)
    pts(656) = Array(54.751364672341, 7.929600000000)
    pts(657) = Array(54.744522290031, 7.907200000000)
    pts(658) = Array(54.737701061534, 7.884800000000)
    pts(659) = Array(54.730900970164, 7.862400000000)
    pts(660) = Array(54.724121999300, 7.840000000000)
    pts(661) = Array(54.717364132393, 7.817600000000)
    pts(662) = Array(54.710627352962, 7.795200000000)
    pts(663) = Array(54.703911644594, 7.772800000000)
    pts(664) = Array(54.697216990944, 7.750400000000)
    pts(665) = Array(54.690543375735, 7.728000000000)
    pts(666) = Array(54.683890782760, 7.705600000000)
    pts(667) = Array(54.677259195876, 7.683200000000)
    pts(668) = Array(54.670648599012, 7.660800000000)
    pts(669) = Array(54.664058976159, 7.638400000000)
    pts(670) = Array(54.657490311379, 7.616000000000)
    pts(671) = Array(54.650942588799, 7.593600000000)
    pts(672) = Array(54.644415792614, 7.571200000000)
    pts(673) = Array(54.637909907082, 7.548800000000)
    pts(674) = Array(54.631424916531, 7.526400000000)
    pts(675) = Array(54.624960805353, 7.504000000000)
    pts(676) = Array(54.623738330014, 7.499755623002)
    pts(677) = Array(54.623420493969, 7.498651685025)
    pts(678) = Array(54.623065845765, 7.497419679825)
    pts(679) = Array(54.618517558006, 7.481600000000)
    pts(680) = Array(54.612095159013, 7.459200000000)
    pts(681) = Array(54.605693592963, 7.436800000000)
    pts(682) = Array(54.599312844511, 7.414400000000)
    pts(683) = Array(54.592952898375, 7.392000000000)
    pts(684) = Array(54.586613739338, 7.369600000000)
    pts(685) = Array(54.580295352249, 7.347200000000)
    pts(686) = Array(54.573997722020, 7.324800000000)
    pts(687) = Array(54.567720833627, 7.302400000000)
    pts(688) = Array(54.561464672111, 7.280000000000)
    pts(689) = Array(54.555229222575, 7.257600000000)
    pts(690) = Array(54.549014470188, 7.235200000000)
    pts(691) = Array(54.542820400181, 7.212800000000)
    pts(692) = Array(54.536646997846, 7.190400000000)
    pts(693) = Array(54.530494248542, 7.168000000000)
    pts(694) = Array(54.524362137687, 7.145600000000)
    pts(695) = Array(54.518250650764, 7.123200000000)
    pts(696) = Array(54.512159773319, 7.100800000000)
    pts(697) = Array(54.506089490957, 7.078400000000)
    pts(698) = Array(54.500039789347, 7.056000000000)
    pts(699) = Array(54.494010654220, 7.033600000000)
    pts(700) = Array(54.488002071369, 7.011200000000)
    pts(701) = Array(54.482014026647, 6.988800000000)
    pts(702) = Array(54.476046505968, 6.966400000000)
    pts(703) = Array(54.470099495310, 6.944000000000)
    pts(704) = Array(54.464172980709, 6.921600000000)
    pts(705) = Array(54.458266948262, 6.899200000000)
    pts(706) = Array(54.452381384128, 6.876800000000)
    pts(707) = Array(54.446516274526, 6.854400000000)
    pts(708) = Array(54.440671605735, 6.832000000000)
    pts(709) = Array(54.434847364092, 6.809600000000)
    pts(710) = Array(54.429043535998, 6.787200000000)
    pts(711) = Array(54.423260107910, 6.764800000000)
    pts(712) = Array(54.417497066346, 6.742400000000)
    pts(713) = Array(54.411754397884, 6.720000000000)
    pts(714) = Array(54.406032089160, 6.697600000000)
    pts(715) = Array(54.400330126869, 6.675200000000)
    pts(716) = Array(54.394648497765, 6.652800000000)
    pts(717) = Array(54.388987188662, 6.630400000000)
    pts(718) = Array(54.383346186431, 6.608000000000)
    pts(719) = Array(54.377725478002, 6.585600000000)
    pts(720) = Array(54.372125050362, 6.563200000000)
    pts(721) = Array(54.366544890558, 6.540800000000)
    pts(722) = Array(54.360984985693, 6.518400000000)
    pts(723) = Array(54.355445322929, 6.496000000000)
    pts(724) = Array(54.349925889484, 6.473600000000)
    pts(725) = Array(54.344426672636, 6.451200000000)
    pts(726) = Array(54.338947659717, 6.428800000000)
    pts(727) = Array(54.333488838119, 6.406400000000)
    pts(728) = Array(54.328050195289, 6.384000000000)
    pts(729) = Array(54.322631718731, 6.361600000000)
    pts(730) = Array(54.317233396007, 6.339200000000)
    pts(731) = Array(54.311855214732, 6.316800000000)
    pts(732) = Array(54.306497162582, 6.294400000000)
    pts(733) = Array(54.301159227286, 6.272000000000)
    pts(734) = Array(54.295841396629, 6.249600000000)
    pts(735) = Array(54.290543658453, 6.227200000000)
    pts(736) = Array(54.285266000656, 6.204800000000)
    pts(737) = Array(54.280008411190, 6.182400000000)
    pts(738) = Array(54.274770878063, 6.160000000000)
    pts(739) = Array(54.269553389339, 6.137600000000)
    pts(740) = Array(54.264355933136, 6.115200000000)
    pts(741) = Array(54.259178497628, 6.092800000000)
    pts(742) = Array(54.254021071044, 6.070400000000)
    pts(743) = Array(54.248883641665, 6.048000000000)
    pts(744) = Array(54.243766197829, 6.025600000000)
    pts(745) = Array(54.238668727929, 6.003200000000)
    pts(746) = Array(54.233591220410, 5.980800000000)
    pts(747) = Array(54.228533663773, 5.958400000000)
    pts(748) = Array(54.223496046572, 5.936000000000)
    pts(749) = Array(54.218478357414, 5.913600000000)
    pts(750) = Array(54.213480584962, 5.891200000000)
    pts(751) = Array(54.208502717930, 5.868800000000)
    pts(752) = Array(54.203544745088, 5.846400000000)
    pts(753) = Array(54.198606655258, 5.824000000000)
    pts(754) = Array(54.193688437313, 5.801600000000)
    pts(755) = Array(54.188790080183, 5.779200000000)
    pts(756) = Array(54.183911572849, 5.756800000000)
    pts(757) = Array(54.179052904344, 5.734400000000)
    pts(758) = Array(54.174214063754, 5.712000000000)
    pts(759) = Array(54.169395040219, 5.689600000000)
    pts(760) = Array(54.164595822930, 5.667200000000)
    pts(761) = Array(54.159816401130, 5.644800000000)
    pts(762) = Array(54.155056764114, 5.622400000000)
    pts(763) = Array(54.150316901231, 5.600000000000)
    pts(764) = Array(54.145596801880, 5.577600000000)
    pts(765) = Array(54.140896455511, 5.555200000000)
    pts(766) = Array(54.136215851628, 5.532800000000)
    pts(767) = Array(54.131554979784, 5.510400000000)
    pts(768) = Array(54.126913829586, 5.488000000000)
    pts(769) = Array(54.122292390690, 5.465600000000)
    pts(770) = Array(54.117690652804, 5.443200000000)
    pts(771) = Array(54.113108605688, 5.420800000000)
    pts(772) = Array(54.108546239149, 5.398400000000)
    pts(773) = Array(54.104003543051, 5.376000000000)
    pts(774) = Array(54.099480507302, 5.353600000000)
    pts(775) = Array(54.094977121866, 5.331200000000)
    pts(776) = Array(54.090493376754, 5.308800000000)
    pts(777) = Array(54.086029262029, 5.286400000000)
    pts(778) = Array(54.081584767803, 5.264000000000)
    pts(779) = Array(54.077159884239, 5.241600000000)
    pts(780) = Array(54.072754601549, 5.219200000000)
    pts(781) = Array(54.068368909996, 5.196800000000)
    pts(782) = Array(54.064002799891, 5.174400000000)
    pts(783) = Array(54.059656261597, 5.152000000000)
    pts(784) = Array(54.055329285525, 5.129600000000)
    pts(785) = Array(54.051021862134, 5.107200000000)
    pts(786) = Array(54.046733981936, 5.084800000000)
    pts(787) = Array(54.042465635488, 5.062400000000)
    pts(788) = Array(54.038216813400, 5.040000000000)
    pts(789) = Array(54.033987506327, 5.017600000000)
    pts(790) = Array(54.029777704975, 4.995200000000)
    pts(791) = Array(54.025587400100, 4.972800000000)
    pts(792) = Array(54.021416582504, 4.950400000000)
    pts(793) = Array(54.017265243039, 4.928000000000)
    pts(794) = Array(54.013133372605, 4.905600000000)
    pts(795) = Array(54.009020962150, 4.883200000000)
    pts(796) = Array(54.004928002671, 4.860800000000)
    pts(797) = Array(54.000854485212, 4.838400000000)
    pts(798) = Array(53.996800400867, 4.816000000000)
    pts(799) = Array(53.992765740775, 4.793600000000)
    pts(800) = Array(53.988750496125, 4.771200000000)
    pts(801) = Array(53.984754658153, 4.748800000000)
    pts(802) = Array(53.980778218142, 4.726400000000)
    pts(803) = Array(53.976821167424, 4.704000000000)
    pts(804) = Array(53.972883497378, 4.681600000000)
    pts(805) = Array(53.968965199429, 4.659200000000)
    pts(806) = Array(53.965066265050, 4.636800000000)
    pts(807) = Array(53.961186685761, 4.614400000000)
    pts(808) = Array(53.957326453129, 4.592000000000)
    pts(809) = Array(53.953485558769, 4.569600000000)
    pts(810) = Array(53.949663994341, 4.547200000000)
    pts(811) = Array(53.945861751552, 4.524800000000)
    pts(812) = Array(53.942078822158, 4.502400000000)
    pts(813) = Array(53.938315197959, 4.480000000000)
    pts(814) = Array(53.934570870802, 4.457600000000)
    pts(815) = Array(53.930845832580, 4.435200000000)
    pts(816) = Array(53.927140075234, 4.412800000000)
    pts(817) = Array(53.923453590749, 4.390400000000)
    pts(818) = Array(53.919786371158, 4.368000000000)
    pts(819) = Array(53.916138408539, 4.345600000000)
    pts(820) = Array(53.912509695015, 4.323200000000)
    pts(821) = Array(53.908900222757, 4.300800000000)
    pts(822) = Array(53.905309983979, 4.278400000000)
    pts(823) = Array(53.901738970944, 4.256000000000)
    pts(824) = Array(53.898187175958, 4.233600000000)
    pts(825) = Array(53.894654591373, 4.211200000000)
    pts(826) = Array(53.891141209586, 4.188800000000)
    pts(827) = Array(53.887647023042, 4.166400000000)
    pts(828) = Array(53.884172024227, 4.144000000000)
    pts(829) = Array(53.880716205676, 4.121600000000)
    pts(830) = Array(53.877279559966, 4.099200000000)
    pts(831) = Array(53.873862079720, 4.076800000000)
    pts(832) = Array(53.870463757608, 4.054400000000)
    pts(833) = Array(53.867084586342, 4.032000000000)
    pts(834) = Array(53.863724558679, 4.009600000000)
    pts(835) = Array(53.860383667422, 3.987200000000)
    pts(836) = Array(53.857061905419, 3.964800000000)
    pts(837) = Array(53.853759265560, 3.942400000000)
    pts(838) = Array(53.850475740781, 3.920000000000)
    pts(839) = Array(53.847211324063, 3.897600000000)
    pts(840) = Array(53.843966008429, 3.875200000000)
    pts(841) = Array(53.840739786950, 3.852800000000)
    pts(842) = Array(53.837532652736, 3.830400000000)
    pts(843) = Array(53.834344598946, 3.808000000000)
    pts(844) = Array(53.831175618780, 3.785600000000)
    pts(845) = Array(53.828025705482, 3.763200000000)
    pts(846) = Array(53.824894852340, 3.740800000000)
    pts(847) = Array(53.821783052688, 3.718400000000)
    pts(848) = Array(53.818690299901, 3.696000000000)
    pts(849) = Array(53.815616587397, 3.673600000000)
    pts(850) = Array(53.812561908641, 3.651200000000)
    pts(851) = Array(53.809526257138, 3.628800000000)
    pts(852) = Array(53.806509626438, 3.606400000000)
    pts(853) = Array(53.803512010134, 3.584000000000)
    pts(854) = Array(53.800533401863, 3.561600000000)
    pts(855) = Array(53.797573795304, 3.539200000000)
    pts(856) = Array(53.794633184178, 3.516800000000)
    pts(857) = Array(53.791711562253, 3.494400000000)
    pts(858) = Array(53.788808923336, 3.472000000000)
    pts(859) = Array(53.785925261279, 3.449600000000)
    pts(860) = Array(53.783060569976, 3.427200000000)
    pts(861) = Array(53.780214843364, 3.404800000000)
    pts(862) = Array(53.777388075423, 3.382400000000)
    pts(863) = Array(53.774580260175, 3.360000000000)
    pts(864) = Array(53.771791391685, 3.337600000000)
    pts(865) = Array(53.769021464060, 3.315200000000)
    pts(866) = Array(53.766270471450, 3.292800000000)
    pts(867) = Array(53.763538408047, 3.270400000000)
    pts(868) = Array(53.760825268087, 3.248000000000)
    pts(869) = Array(53.758131045845, 3.225600000000)
    pts(870) = Array(53.755455735641, 3.203200000000)
    pts(871) = Array(53.752799331835, 3.180800000000)
    pts(872) = Array(53.750161828831, 3.158400000000)
    pts(873) = Array(53.747543221074, 3.136000000000)
    pts(874) = Array(53.744943503051, 3.113600000000)
    pts(875) = Array(53.742362669292, 3.091200000000)
    pts(876) = Array(53.739800714366, 3.068800000000)
    pts(877) = Array(53.737257632887, 3.046400000000)
    pts(878) = Array(53.734733419509, 3.024000000000)
    pts(879) = Array(53.732228068927, 3.001600000000)
    pts(880) = Array(53.729741575880, 2.979200000000)
    pts(881) = Array(53.727273935147, 2.956800000000)
    pts(882) = Array(53.724825141547, 2.934400000000)
    pts(883) = Array(53.722395189944, 2.912000000000)
    pts(884) = Array(53.719984075241, 2.889600000000)
    pts(885) = Array(53.717591792381, 2.867200000000)
    pts(886) = Array(53.715218336352, 2.844800000000)
    pts(887) = Array(53.712863702181, 2.822400000000)
    pts(888) = Array(53.710527884936, 2.800000000000)
    pts(889) = Array(53.708210879726, 2.777600000000)
    pts(890) = Array(53.705912681703, 2.755200000000)
    pts(891) = Array(53.703633286057, 2.732800000000)
    pts(892) = Array(53.701372688022, 2.710400000000)
    pts(893) = Array(53.699130882870, 2.688000000000)
    pts(894) = Array(53.696907865916, 2.665600000000)
    pts(895) = Array(53.694703632515, 2.643200000000)
    pts(896) = Array(53.692518178064, 2.620800000000)
    pts(897) = Array(53.690351497997, 2.598400000000)
    pts(898) = Array(53.688203587794, 2.576000000000)
    pts(899) = Array(53.686074442972, 2.553600000000)
    pts(900) = Array(53.683964059088, 2.531200000000)
    pts(901) = Array(53.681872431743, 2.508800000000)
    pts(902) = Array(53.679799556575, 2.486400000000)
    pts(903) = Array(53.677745429265, 2.464000000000)
    pts(904) = Array(53.675710045532, 2.441600000000)
    pts(905) = Array(53.673693401138, 2.419200000000)
    pts(906) = Array(53.671695491882, 2.396800000000)
    pts(907) = Array(53.669716313607, 2.374400000000)
    pts(908) = Array(53.667755862193, 2.352000000000)
    pts(909) = Array(53.665814133562, 2.329600000000)
    pts(910) = Array(53.663891123676, 2.307200000000)
    pts(911) = Array(53.661986828535, 2.284800000000)
    pts(912) = Array(53.660101244183, 2.262400000000)
    pts(913) = Array(53.658234366701, 2.240000000000)
    pts(914) = Array(53.656386192209, 2.217600000000)
    pts(915) = Array(53.654556716871, 2.195200000000)
    pts(916) = Array(53.652745936886, 2.172800000000)
    pts(917) = Array(53.650953848497, 2.150400000000)
    pts(918) = Array(53.649180447985, 2.128000000000)
    pts(919) = Array(53.647425731670, 2.105600000000)
    pts(920) = Array(53.645689695913, 2.083200000000)
    pts(921) = Array(53.643972337114, 2.060800000000)
    pts(922) = Array(53.642273651713, 2.038400000000)
    pts(923) = Array(53.640593636190, 2.016000000000)
    pts(924) = Array(53.638932287062, 1.993600000000)
    pts(925) = Array(53.637289600890, 1.971200000000)
    pts(926) = Array(53.635665574270, 1.948800000000)
    pts(927) = Array(53.634060203841, 1.926400000000)
    pts(928) = Array(53.632473486279, 1.904000000000)
    pts(929) = Array(53.630905418301, 1.881600000000)
    pts(930) = Array(53.629355996661, 1.859200000000)
    pts(931) = Array(53.627825218156, 1.836800000000)
    pts(932) = Array(53.626313079619, 1.814400000000)
    pts(933) = Array(53.624819577924, 1.792000000000)
    pts(934) = Array(53.623344709984, 1.769600000000)
    pts(935) = Array(53.621888472750, 1.747200000000)
    pts(936) = Array(53.620450863213, 1.724800000000)
    pts(937) = Array(53.619031878404, 1.702400000000)
    pts(938) = Array(53.617631515392, 1.680000000000)
    pts(939) = Array(53.616249771285, 1.657600000000)
    pts(940) = Array(53.614886643231, 1.635200000000)
    pts(941) = Array(53.613542128417, 1.612800000000)
    pts(942) = Array(53.612216224066, 1.590400000000)
    pts(943) = Array(53.610908927444, 1.568000000000)
    pts(944) = Array(53.609620235854, 1.545600000000)
    pts(945) = Array(53.608350146638, 1.523200000000)
    pts(946) = Array(53.607098657177, 1.500800000000)
    pts(947) = Array(53.605865764891, 1.478400000000)
    pts(948) = Array(53.604651467238, 1.456000000000)
    pts(949) = Array(53.603455761716, 1.433600000000)
    pts(950) = Array(53.602278645860, 1.411200000000)
    pts(951) = Array(53.601120117247, 1.388800000000)
    pts(952) = Array(53.599980173489, 1.366400000000)
    pts(953) = Array(53.598858812238, 1.344000000000)
    pts(954) = Array(53.597756031186, 1.321600000000)
    pts(955) = Array(53.596671828062, 1.299200000000)
    pts(956) = Array(53.595606200635, 1.276800000000)
    pts(957) = Array(53.594559146710, 1.254400000000)
    pts(958) = Array(53.593530664134, 1.232000000000)
    pts(959) = Array(53.592520750790, 1.209600000000)
    pts(960) = Array(53.591529404601, 1.187200000000)
    pts(961) = Array(53.590556623527, 1.164800000000)
    pts(962) = Array(53.589602405568, 1.142400000000)
    pts(963) = Array(53.588666748762, 1.120000000000)
    pts(964) = Array(53.587749651184, 1.097600000000)
    pts(965) = Array(53.586851110950, 1.075200000000)
    pts(966) = Array(53.585971126212, 1.052800000000)
    pts(967) = Array(53.585109695163, 1.030400000000)
    pts(968) = Array(53.584266816031, 1.008000000000)
    pts(969) = Array(53.583442487085, 0.985600000000)
    pts(970) = Array(53.582636706632, 0.963200000000)
    pts(971) = Array(53.581849473016, 0.940800000000)
    pts(972) = Array(53.581080784620, 0.918400000000)
    pts(973) = Array(53.580330639865, 0.896000000000)
    pts(974) = Array(53.579599037212, 0.873600000000)
    pts(975) = Array(53.578885975158, 0.851200000000)
    pts(976) = Array(53.578191452239, 0.828800000000)
    pts(977) = Array(53.577515467029, 0.806400000000)
    pts(978) = Array(53.576858018141, 0.784000000000)
    pts(979) = Array(53.576219104225, 0.761600000000)
    pts(980) = Array(53.575598723970, 0.739200000000)
    pts(981) = Array(53.574996876104, 0.716800000000)
    pts(982) = Array(53.574413559391, 0.694400000000)
    pts(983) = Array(53.573848772634, 0.672000000000)
    pts(984) = Array(53.573302514675, 0.649600000000)
    pts(985) = Array(53.572774784393, 0.627200000000)
    pts(986) = Array(53.572265580706, 0.604800000000)
    pts(987) = Array(53.571774902570, 0.582400000000)
    pts(988) = Array(53.571302748977, 0.560000000000)
    pts(989) = Array(53.570849118961, 0.537600000000)
    pts(990) = Array(53.570414011590, 0.515200000000)
    pts(991) = Array(53.569997425972, 0.492800000000)
    pts(992) = Array(53.569599361254, 0.470400000000)
    pts(993) = Array(53.569219816619, 0.448000000000)
    pts(994) = Array(53.568858791290, 0.425600000000)
    pts(995) = Array(53.568516284525, 0.403200000000)
    pts(996) = Array(53.568192295623, 0.380800000000)
    pts(997) = Array(53.567886823921, 0.358400000000)
    pts(998) = Array(53.567599868791, 0.336000000000)
    pts(999) = Array(53.567331429645, 0.313600000000)
    pts(1000) = Array(53.567081505935, 0.291200000000)
    pts(1001) = Array(53.566850097146, 0.268800000000)
    pts(1002) = Array(53.566637202805, 0.246400000000)
    pts(1003) = Array(53.566442822477, 0.224000000000)
    pts(1004) = Array(53.566266955761, 0.201600000000)
    pts(1005) = Array(53.566109602299, 0.179200000000)
    pts(1006) = Array(53.565970761768, 0.156800000000)
    pts(1007) = Array(53.565850433882, 0.134400000000)
    pts(1008) = Array(53.565748618396, 0.112000000000)
    pts(1009) = Array(53.565665315102, 0.089600000000)
    pts(1010) = Array(53.565600523827, 0.067200000000)
    pts(1011) = Array(53.565554244440, 0.044800000000)
    pts(1012) = Array(53.565526476846, 0.022400000000)
    pts(1013) = Array(53.565517220987, 0.000000000000)
    Profile_15 = pts
End Function

Private Function Profile_16() As Variant
    Dim pts(1013) As Variant
    pts(0) = Array(53.565517220987, 0.000000000000)
    pts(1) = Array(53.565526476846, 0.022400000000)
    pts(2) = Array(53.565554244440, 0.044800000000)
    pts(3) = Array(53.565600523827, 0.067200000000)
    pts(4) = Array(53.565665315102, 0.089600000000)
    pts(5) = Array(53.565748618396, 0.112000000000)
    pts(6) = Array(53.565850433882, 0.134400000000)
    pts(7) = Array(53.565970761768, 0.156800000000)
    pts(8) = Array(53.566109602299, 0.179200000000)
    pts(9) = Array(53.566266955761, 0.201600000000)
    pts(10) = Array(53.566442822477, 0.224000000000)
    pts(11) = Array(53.566637202805, 0.246400000000)
    pts(12) = Array(53.566850097146, 0.268800000000)
    pts(13) = Array(53.567081505935, 0.291200000000)
    pts(14) = Array(53.567331429645, 0.313600000000)
    pts(15) = Array(53.567599868791, 0.336000000000)
    pts(16) = Array(53.567886823921, 0.358400000000)
    pts(17) = Array(53.568192295623, 0.380800000000)
    pts(18) = Array(53.568516284525, 0.403200000000)
    pts(19) = Array(53.568858791290, 0.425600000000)
    pts(20) = Array(53.569219816619, 0.448000000000)
    pts(21) = Array(53.569599361254, 0.470400000000)
    pts(22) = Array(53.569997425972, 0.492800000000)
    pts(23) = Array(53.570414011590, 0.515200000000)
    pts(24) = Array(53.570849118961, 0.537600000000)
    pts(25) = Array(53.571302748977, 0.560000000000)
    pts(26) = Array(53.571774902570, 0.582400000000)
    pts(27) = Array(53.572265580706, 0.604800000000)
    pts(28) = Array(53.572774784393, 0.627200000000)
    pts(29) = Array(53.573302514675, 0.649600000000)
    pts(30) = Array(53.573848772634, 0.672000000000)
    pts(31) = Array(53.574413559391, 0.694400000000)
    pts(32) = Array(53.574996876104, 0.716800000000)
    pts(33) = Array(53.575598723970, 0.739200000000)
    pts(34) = Array(53.576219104225, 0.761600000000)
    pts(35) = Array(53.576858018141, 0.784000000000)
    pts(36) = Array(53.577515467029, 0.806400000000)
    pts(37) = Array(53.578191452239, 0.828800000000)
    pts(38) = Array(53.578885975158, 0.851200000000)
    pts(39) = Array(53.579599037212, 0.873600000000)
    pts(40) = Array(53.580330639865, 0.896000000000)
    pts(41) = Array(53.581080784620, 0.918400000000)
    pts(42) = Array(53.581849473016, 0.940800000000)
    pts(43) = Array(53.582636706632, 0.963200000000)
    pts(44) = Array(53.583442487085, 0.985600000000)
    pts(45) = Array(53.584266816031, 1.008000000000)
    pts(46) = Array(53.585109695163, 1.030400000000)
    pts(47) = Array(53.585971126212, 1.052800000000)
    pts(48) = Array(53.586851110950, 1.075200000000)
    pts(49) = Array(53.587749651184, 1.097600000000)
    pts(50) = Array(53.588666748762, 1.120000000000)
    pts(51) = Array(53.589602405568, 1.142400000000)
    pts(52) = Array(53.590556623527, 1.164800000000)
    pts(53) = Array(53.591529404601, 1.187200000000)
    pts(54) = Array(53.592520750790, 1.209600000000)
    pts(55) = Array(53.593530664134, 1.232000000000)
    pts(56) = Array(53.594559146710, 1.254400000000)
    pts(57) = Array(53.595606200635, 1.276800000000)
    pts(58) = Array(53.596671828062, 1.299200000000)
    pts(59) = Array(53.597756031186, 1.321600000000)
    pts(60) = Array(53.598858812238, 1.344000000000)
    pts(61) = Array(53.599980173489, 1.366400000000)
    pts(62) = Array(53.601120117247, 1.388800000000)
    pts(63) = Array(53.602278645860, 1.411200000000)
    pts(64) = Array(53.603455761716, 1.433600000000)
    pts(65) = Array(53.604651467238, 1.456000000000)
    pts(66) = Array(53.605865764891, 1.478400000000)
    pts(67) = Array(53.607098657177, 1.500800000000)
    pts(68) = Array(53.608350146638, 1.523200000000)
    pts(69) = Array(53.609620235854, 1.545600000000)
    pts(70) = Array(53.610908927444, 1.568000000000)
    pts(71) = Array(53.612216224066, 1.590400000000)
    pts(72) = Array(53.613542128417, 1.612800000000)
    pts(73) = Array(53.614886643231, 1.635200000000)
    pts(74) = Array(53.616249771285, 1.657600000000)
    pts(75) = Array(53.617631515392, 1.680000000000)
    pts(76) = Array(53.619031878404, 1.702400000000)
    pts(77) = Array(53.620450863213, 1.724800000000)
    pts(78) = Array(53.621888472750, 1.747200000000)
    pts(79) = Array(53.623344709984, 1.769600000000)
    pts(80) = Array(53.624819577924, 1.792000000000)
    pts(81) = Array(53.626313079619, 1.814400000000)
    pts(82) = Array(53.627825218156, 1.836800000000)
    pts(83) = Array(53.629355996661, 1.859200000000)
    pts(84) = Array(53.630905418301, 1.881600000000)
    pts(85) = Array(53.632473486279, 1.904000000000)
    pts(86) = Array(53.634060203841, 1.926400000000)
    pts(87) = Array(53.635665574270, 1.948800000000)
    pts(88) = Array(53.637289600890, 1.971200000000)
    pts(89) = Array(53.638932287062, 1.993600000000)
    pts(90) = Array(53.640593636190, 2.016000000000)
    pts(91) = Array(53.642273651713, 2.038400000000)
    pts(92) = Array(53.643972337114, 2.060800000000)
    pts(93) = Array(53.645689695913, 2.083200000000)
    pts(94) = Array(53.647425731670, 2.105600000000)
    pts(95) = Array(53.649180447985, 2.128000000000)
    pts(96) = Array(53.650953848497, 2.150400000000)
    pts(97) = Array(53.652745936886, 2.172800000000)
    pts(98) = Array(53.654556716871, 2.195200000000)
    pts(99) = Array(53.656386192209, 2.217600000000)
    pts(100) = Array(53.658234366701, 2.240000000000)
    pts(101) = Array(53.660101244183, 2.262400000000)
    pts(102) = Array(53.661986828535, 2.284800000000)
    pts(103) = Array(53.663891123676, 2.307200000000)
    pts(104) = Array(53.665814133562, 2.329600000000)
    pts(105) = Array(53.667755862193, 2.352000000000)
    pts(106) = Array(53.669716313607, 2.374400000000)
    pts(107) = Array(53.671695491882, 2.396800000000)
    pts(108) = Array(53.673693401138, 2.419200000000)
    pts(109) = Array(53.675710045532, 2.441600000000)
    pts(110) = Array(53.677745429265, 2.464000000000)
    pts(111) = Array(53.679799556575, 2.486400000000)
    pts(112) = Array(53.681872431743, 2.508800000000)
    pts(113) = Array(53.683964059088, 2.531200000000)
    pts(114) = Array(53.686074442972, 2.553600000000)
    pts(115) = Array(53.688203587794, 2.576000000000)
    pts(116) = Array(53.690351497997, 2.598400000000)
    pts(117) = Array(53.692518178064, 2.620800000000)
    pts(118) = Array(53.694703632515, 2.643200000000)
    pts(119) = Array(53.696907865916, 2.665600000000)
    pts(120) = Array(53.699130882870, 2.688000000000)
    pts(121) = Array(53.701372688022, 2.710400000000)
    pts(122) = Array(53.703633286057, 2.732800000000)
    pts(123) = Array(53.705912681703, 2.755200000000)
    pts(124) = Array(53.708210879726, 2.777600000000)
    pts(125) = Array(53.710527884936, 2.800000000000)
    pts(126) = Array(53.712863702181, 2.822400000000)
    pts(127) = Array(53.715218336352, 2.844800000000)
    pts(128) = Array(53.717591792381, 2.867200000000)
    pts(129) = Array(53.719984075241, 2.889600000000)
    pts(130) = Array(53.722395189944, 2.912000000000)
    pts(131) = Array(53.724825141547, 2.934400000000)
    pts(132) = Array(53.727273935147, 2.956800000000)
    pts(133) = Array(53.729741575880, 2.979200000000)
    pts(134) = Array(53.732228068927, 3.001600000000)
    pts(135) = Array(53.734733419509, 3.024000000000)
    pts(136) = Array(53.737257632887, 3.046400000000)
    pts(137) = Array(53.739800714366, 3.068800000000)
    pts(138) = Array(53.742362669292, 3.091200000000)
    pts(139) = Array(53.744943503051, 3.113600000000)
    pts(140) = Array(53.747543221074, 3.136000000000)
    pts(141) = Array(53.750161828831, 3.158400000000)
    pts(142) = Array(53.752799331835, 3.180800000000)
    pts(143) = Array(53.755455735641, 3.203200000000)
    pts(144) = Array(53.758131045845, 3.225600000000)
    pts(145) = Array(53.760825268087, 3.248000000000)
    pts(146) = Array(53.763538408047, 3.270400000000)
    pts(147) = Array(53.766270471450, 3.292800000000)
    pts(148) = Array(53.769021464060, 3.315200000000)
    pts(149) = Array(53.771791391685, 3.337600000000)
    pts(150) = Array(53.774580260175, 3.360000000000)
    pts(151) = Array(53.777388075423, 3.382400000000)
    pts(152) = Array(53.780214843364, 3.404800000000)
    pts(153) = Array(53.783060569976, 3.427200000000)
    pts(154) = Array(53.785925261279, 3.449600000000)
    pts(155) = Array(53.788808923336, 3.472000000000)
    pts(156) = Array(53.791711562253, 3.494400000000)
    pts(157) = Array(53.794633184178, 3.516800000000)
    pts(158) = Array(53.797573795304, 3.539200000000)
    pts(159) = Array(53.800533401863, 3.561600000000)
    pts(160) = Array(53.803512010134, 3.584000000000)
    pts(161) = Array(53.806509626438, 3.606400000000)
    pts(162) = Array(53.809526257138, 3.628800000000)
    pts(163) = Array(53.812561908641, 3.651200000000)
    pts(164) = Array(53.815616587397, 3.673600000000)
    pts(165) = Array(53.818690299901, 3.696000000000)
    pts(166) = Array(53.821783052688, 3.718400000000)
    pts(167) = Array(53.824894852340, 3.740800000000)
    pts(168) = Array(53.828025705482, 3.763200000000)
    pts(169) = Array(53.831175618780, 3.785600000000)
    pts(170) = Array(53.834344598946, 3.808000000000)
    pts(171) = Array(53.837532652736, 3.830400000000)
    pts(172) = Array(53.840739786950, 3.852800000000)
    pts(173) = Array(53.843966008429, 3.875200000000)
    pts(174) = Array(53.847211324063, 3.897600000000)
    pts(175) = Array(53.850475740781, 3.920000000000)
    pts(176) = Array(53.853759265560, 3.942400000000)
    pts(177) = Array(53.857061905419, 3.964800000000)
    pts(178) = Array(53.860383667422, 3.987200000000)
    pts(179) = Array(53.863724558679, 4.009600000000)
    pts(180) = Array(53.867084586342, 4.032000000000)
    pts(181) = Array(53.870463757608, 4.054400000000)
    pts(182) = Array(53.873862079720, 4.076800000000)
    pts(183) = Array(53.877279559966, 4.099200000000)
    pts(184) = Array(53.880716205676, 4.121600000000)
    pts(185) = Array(53.884172024227, 4.144000000000)
    pts(186) = Array(53.887647023042, 4.166400000000)
    pts(187) = Array(53.891141209586, 4.188800000000)
    pts(188) = Array(53.894654591373, 4.211200000000)
    pts(189) = Array(53.898187175958, 4.233600000000)
    pts(190) = Array(53.901738970944, 4.256000000000)
    pts(191) = Array(53.905309983979, 4.278400000000)
    pts(192) = Array(53.908900222757, 4.300800000000)
    pts(193) = Array(53.912509695015, 4.323200000000)
    pts(194) = Array(53.916138408539, 4.345600000000)
    pts(195) = Array(53.919786371158, 4.368000000000)
    pts(196) = Array(53.923453590749, 4.390400000000)
    pts(197) = Array(53.927140075234, 4.412800000000)
    pts(198) = Array(53.930845832580, 4.435200000000)
    pts(199) = Array(53.934570870802, 4.457600000000)
    pts(200) = Array(53.938315197959, 4.480000000000)
    pts(201) = Array(53.942078822158, 4.502400000000)
    pts(202) = Array(53.945861751552, 4.524800000000)
    pts(203) = Array(53.949663994341, 4.547200000000)
    pts(204) = Array(53.953485558769, 4.569600000000)
    pts(205) = Array(53.957326453129, 4.592000000000)
    pts(206) = Array(53.961186685761, 4.614400000000)
    pts(207) = Array(53.965066265050, 4.636800000000)
    pts(208) = Array(53.968965199429, 4.659200000000)
    pts(209) = Array(53.972883497378, 4.681600000000)
    pts(210) = Array(53.976821167424, 4.704000000000)
    pts(211) = Array(53.980778218142, 4.726400000000)
    pts(212) = Array(53.984754658153, 4.748800000000)
    pts(213) = Array(53.988750496125, 4.771200000000)
    pts(214) = Array(53.992765740775, 4.793600000000)
    pts(215) = Array(53.996800400867, 4.816000000000)
    pts(216) = Array(54.000854485212, 4.838400000000)
    pts(217) = Array(54.004928002671, 4.860800000000)
    pts(218) = Array(54.009020962150, 4.883200000000)
    pts(219) = Array(54.013133372605, 4.905600000000)
    pts(220) = Array(54.017265243039, 4.928000000000)
    pts(221) = Array(54.021416582504, 4.950400000000)
    pts(222) = Array(54.025587400100, 4.972800000000)
    pts(223) = Array(54.029777704975, 4.995200000000)
    pts(224) = Array(54.033987506327, 5.017600000000)
    pts(225) = Array(54.038216813400, 5.040000000000)
    pts(226) = Array(54.042465635488, 5.062400000000)
    pts(227) = Array(54.046733981936, 5.084800000000)
    pts(228) = Array(54.051021862134, 5.107200000000)
    pts(229) = Array(54.055329285525, 5.129600000000)
    pts(230) = Array(54.059656261597, 5.152000000000)
    pts(231) = Array(54.064002799891, 5.174400000000)
    pts(232) = Array(54.068368909996, 5.196800000000)
    pts(233) = Array(54.072754601549, 5.219200000000)
    pts(234) = Array(54.077159884239, 5.241600000000)
    pts(235) = Array(54.081584767803, 5.264000000000)
    pts(236) = Array(54.086029262029, 5.286400000000)
    pts(237) = Array(54.090493376754, 5.308800000000)
    pts(238) = Array(54.094977121866, 5.331200000000)
    pts(239) = Array(54.099480507302, 5.353600000000)
    pts(240) = Array(54.104003543051, 5.376000000000)
    pts(241) = Array(54.108546239149, 5.398400000000)
    pts(242) = Array(54.113108605688, 5.420800000000)
    pts(243) = Array(54.117690652804, 5.443200000000)
    pts(244) = Array(54.122292390690, 5.465600000000)
    pts(245) = Array(54.126913829586, 5.488000000000)
    pts(246) = Array(54.131554979784, 5.510400000000)
    pts(247) = Array(54.136215851628, 5.532800000000)
    pts(248) = Array(54.140896455511, 5.555200000000)
    pts(249) = Array(54.145596801880, 5.577600000000)
    pts(250) = Array(54.150316901231, 5.600000000000)
    pts(251) = Array(54.155056764114, 5.622400000000)
    pts(252) = Array(54.159816401130, 5.644800000000)
    pts(253) = Array(54.164595822930, 5.667200000000)
    pts(254) = Array(54.169395040219, 5.689600000000)
    pts(255) = Array(54.174214063754, 5.712000000000)
    pts(256) = Array(54.179052904344, 5.734400000000)
    pts(257) = Array(54.183911572849, 5.756800000000)
    pts(258) = Array(54.188790080183, 5.779200000000)
    pts(259) = Array(54.193688437313, 5.801600000000)
    pts(260) = Array(54.198606655258, 5.824000000000)
    pts(261) = Array(54.203544745088, 5.846400000000)
    pts(262) = Array(54.208502717930, 5.868800000000)
    pts(263) = Array(54.213480584962, 5.891200000000)
    pts(264) = Array(54.218478357414, 5.913600000000)
    pts(265) = Array(54.223496046572, 5.936000000000)
    pts(266) = Array(54.228533663773, 5.958400000000)
    pts(267) = Array(54.233591220410, 5.980800000000)
    pts(268) = Array(54.238668727929, 6.003200000000)
    pts(269) = Array(54.243766197829, 6.025600000000)
    pts(270) = Array(54.248883641665, 6.048000000000)
    pts(271) = Array(54.254021071044, 6.070400000000)
    pts(272) = Array(54.259178497628, 6.092800000000)
    pts(273) = Array(54.264355933136, 6.115200000000)
    pts(274) = Array(54.269553389339, 6.137600000000)
    pts(275) = Array(54.274770878063, 6.160000000000)
    pts(276) = Array(54.280008411190, 6.182400000000)
    pts(277) = Array(54.285266000656, 6.204800000000)
    pts(278) = Array(54.290543658453, 6.227200000000)
    pts(279) = Array(54.295841396629, 6.249600000000)
    pts(280) = Array(54.301159227286, 6.272000000000)
    pts(281) = Array(54.306497162582, 6.294400000000)
    pts(282) = Array(54.311855214732, 6.316800000000)
    pts(283) = Array(54.317233396007, 6.339200000000)
    pts(284) = Array(54.322631718731, 6.361600000000)
    pts(285) = Array(54.328050195289, 6.384000000000)
    pts(286) = Array(54.333488838119, 6.406400000000)
    pts(287) = Array(54.338947659717, 6.428800000000)
    pts(288) = Array(54.344426672636, 6.451200000000)
    pts(289) = Array(54.349925889484, 6.473600000000)
    pts(290) = Array(54.355445322929, 6.496000000000)
    pts(291) = Array(54.360984985693, 6.518400000000)
    pts(292) = Array(54.366544890558, 6.540800000000)
    pts(293) = Array(54.372125050362, 6.563200000000)
    pts(294) = Array(54.377725478002, 6.585600000000)
    pts(295) = Array(54.383346186431, 6.608000000000)
    pts(296) = Array(54.388987188662, 6.630400000000)
    pts(297) = Array(54.394648497765, 6.652800000000)
    pts(298) = Array(54.400330126869, 6.675200000000)
    pts(299) = Array(54.406032089160, 6.697600000000)
    pts(300) = Array(54.411754397884, 6.720000000000)
    pts(301) = Array(54.417497066346, 6.742400000000)
    pts(302) = Array(54.423260107910, 6.764800000000)
    pts(303) = Array(54.429043535998, 6.787200000000)
    pts(304) = Array(54.434847364092, 6.809600000000)
    pts(305) = Array(54.440671605735, 6.832000000000)
    pts(306) = Array(54.446516274526, 6.854400000000)
    pts(307) = Array(54.452381384128, 6.876800000000)
    pts(308) = Array(54.458266948262, 6.899200000000)
    pts(309) = Array(54.464172980709, 6.921600000000)
    pts(310) = Array(54.470099495310, 6.944000000000)
    pts(311) = Array(54.476046505968, 6.966400000000)
    pts(312) = Array(54.482014026647, 6.988800000000)
    pts(313) = Array(54.488002071369, 7.011200000000)
    pts(314) = Array(54.494010654220, 7.033600000000)
    pts(315) = Array(54.500039789347, 7.056000000000)
    pts(316) = Array(54.506089490957, 7.078400000000)
    pts(317) = Array(54.512159773319, 7.100800000000)
    pts(318) = Array(54.518250650764, 7.123200000000)
    pts(319) = Array(54.524362137687, 7.145600000000)
    pts(320) = Array(54.530494248542, 7.168000000000)
    pts(321) = Array(54.536646997846, 7.190400000000)
    pts(322) = Array(54.542820400181, 7.212800000000)
    pts(323) = Array(54.549014470188, 7.235200000000)
    pts(324) = Array(54.555229222575, 7.257600000000)
    pts(325) = Array(54.561464672111, 7.280000000000)
    pts(326) = Array(54.567720833627, 7.302400000000)
    pts(327) = Array(54.573997722020, 7.324800000000)
    pts(328) = Array(54.580295352249, 7.347200000000)
    pts(329) = Array(54.586613739338, 7.369600000000)
    pts(330) = Array(54.592952898375, 7.392000000000)
    pts(331) = Array(54.599312844511, 7.414400000000)
    pts(332) = Array(54.605693592963, 7.436800000000)
    pts(333) = Array(54.612095159013, 7.459200000000)
    pts(334) = Array(54.618517558006, 7.481600000000)
    pts(335) = Array(54.623065845765, 7.497419679825)
    pts(336) = Array(54.623420493969, 7.498651685025)
    pts(337) = Array(54.623738330014, 7.499755623002)
    pts(338) = Array(54.624960805353, 7.504000000000)
    pts(339) = Array(54.631424916531, 7.526400000000)
    pts(340) = Array(54.637909907082, 7.548800000000)
    pts(341) = Array(54.644415792614, 7.571200000000)
    pts(342) = Array(54.650942588799, 7.593600000000)
    pts(343) = Array(54.657490311379, 7.616000000000)
    pts(344) = Array(54.664058976159, 7.638400000000)
    pts(345) = Array(54.670648599012, 7.660800000000)
    pts(346) = Array(54.677259195876, 7.683200000000)
    pts(347) = Array(54.683890782760, 7.705600000000)
    pts(348) = Array(54.690543375735, 7.728000000000)
    pts(349) = Array(54.697216990944, 7.750400000000)
    pts(350) = Array(54.703911644594, 7.772800000000)
    pts(351) = Array(54.710627352962, 7.795200000000)
    pts(352) = Array(54.717364132393, 7.817600000000)
    pts(353) = Array(54.724121999300, 7.840000000000)
    pts(354) = Array(54.730900970164, 7.862400000000)
    pts(355) = Array(54.737701061534, 7.884800000000)
    pts(356) = Array(54.744522290031, 7.907200000000)
    pts(357) = Array(54.751364672341, 7.929600000000)
    pts(358) = Array(54.758228225223, 7.952000000000)
    pts(359) = Array(54.765112965504, 7.974400000000)
    pts(360) = Array(54.772018910081, 7.996800000000)
    pts(361) = Array(54.778946075922, 8.019200000000)
    pts(362) = Array(54.785894480063, 8.041600000000)
    pts(363) = Array(54.792864139612, 8.064000000000)
    pts(364) = Array(54.799855071750, 8.086400000000)
    pts(365) = Array(54.806867293725, 8.108800000000)
    pts(366) = Array(54.813900822858, 8.131200000000)
    pts(367) = Array(54.820955676543, 8.153600000000)
    pts(368) = Array(54.828031872244, 8.176000000000)
    pts(369) = Array(54.835129427498, 8.198400000000)
    pts(370) = Array(54.842248359912, 8.220800000000)
    pts(371) = Array(54.849388687168, 8.243200000000)
    pts(372) = Array(54.856550427021, 8.265600000000)
    pts(373) = Array(54.863733597297, 8.288000000000)
    pts(374) = Array(54.870938215895, 8.310400000000)
    pts(375) = Array(54.878164300790, 8.332800000000)
    pts(376) = Array(54.885411870030, 8.355200000000)
    pts(377) = Array(54.892680941734, 8.377600000000)
    pts(378) = Array(54.899971534099, 8.400000000000)
    pts(379) = Array(54.907283665395, 8.422400000000)
    pts(380) = Array(54.914617353967, 8.444800000000)
    pts(381) = Array(54.921972618234, 8.467200000000)
    pts(382) = Array(54.929349476691, 8.489600000000)
    pts(383) = Array(54.936747947910, 8.512000000000)
    pts(384) = Array(54.944168050536, 8.534400000000)
    pts(385) = Array(54.951609803292, 8.556800000000)
    pts(386) = Array(54.959073224977, 8.579200000000)
    pts(387) = Array(54.966558334466, 8.601600000000)
    pts(388) = Array(54.974065150711, 8.624000000000)
    pts(389) = Array(54.981593692743, 8.646400000000)
    pts(390) = Array(54.989143979667, 8.668800000000)
    pts(391) = Array(54.996716030669, 8.691200000000)
    pts(392) = Array(55.004309865011, 8.713600000000)
    pts(393) = Array(55.011925502033, 8.736000000000)
    pts(394) = Array(55.019562961155, 8.758400000000)
    pts(395) = Array(55.027222261876, 8.780800000000)
    pts(396) = Array(55.034903423772, 8.803200000000)
    pts(397) = Array(55.042606466499, 8.825600000000)
    pts(398) = Array(55.050331409796, 8.848000000000)
    pts(399) = Array(55.058078273476, 8.870400000000)
    pts(400) = Array(55.065847077437, 8.892800000000)
    pts(401) = Array(55.073637841657, 8.915200000000)
    pts(402) = Array(55.081450586192, 8.937600000000)
    pts(403) = Array(55.089285331182, 8.960000000000)
    pts(404) = Array(55.097142096847, 8.982400000000)
    pts(405) = Array(55.105020903490, 9.004800000000)
    pts(406) = Array(55.112921771493, 9.027200000000)
    pts(407) = Array(55.120844721325, 9.049600000000)
    pts(408) = Array(55.128789773534, 9.072000000000)
    pts(409) = Array(55.136756948752, 9.094400000000)
    pts(410) = Array(55.144746267694, 9.116800000000)
    pts(411) = Array(55.152757751159, 9.139200000000)
    pts(412) = Array(55.160791420031, 9.161600000000)
    pts(413) = Array(55.168847295275, 9.184000000000)
    pts(414) = Array(55.176925397943, 9.206400000000)
    pts(415) = Array(55.185025749172, 9.228800000000)
    pts(416) = Array(55.193148370182, 9.251200000000)
    pts(417) = Array(55.201293282281, 9.273600000000)
    pts(418) = Array(55.209460506860, 9.296000000000)
    pts(419) = Array(55.217650065400, 9.318400000000)
    pts(420) = Array(55.225861979464, 9.340800000000)
    pts(421) = Array(55.234096270704, 9.363200000000)
    pts(422) = Array(55.242352960860, 9.385600000000)
    pts(423) = Array(55.250632071758, 9.408000000000)
    pts(424) = Array(55.258933625312, 9.430400000000)
    pts(425) = Array(55.267257643524, 9.452800000000)
    pts(426) = Array(55.275604148484, 9.475200000000)
    pts(427) = Array(55.283973162372, 9.497600000000)
    pts(428) = Array(55.292364707457, 9.520000000000)
    pts(429) = Array(55.300778806096, 9.542400000000)
    pts(430) = Array(55.309215480738, 9.564800000000)
    pts(431) = Array(55.317674753919, 9.587200000000)
    pts(432) = Array(55.326156648268, 9.609600000000)
    pts(433) = Array(55.334661186505, 9.632000000000)
    pts(434) = Array(55.343188391441, 9.654400000000)
    pts(435) = Array(55.351738285976, 9.676800000000)
    pts(436) = Array(55.360310893105, 9.699200000000)
    pts(437) = Array(55.368906235914, 9.721600000000)
    pts(438) = Array(55.377524337582, 9.744000000000)
    pts(439) = Array(55.386165221381, 9.766400000000)
    pts(440) = Array(55.394828910676, 9.788800000000)
    pts(441) = Array(55.403515428927, 9.811200000000)
    pts(442) = Array(55.412224799686, 9.833600000000)
    pts(443) = Array(55.420957046600, 9.856000000000)
    pts(444) = Array(55.429712193414, 9.878400000000)
    pts(445) = Array(55.438490263963, 9.900800000000)
    pts(446) = Array(55.447291282183, 9.923200000000)
    pts(447) = Array(55.456115272101, 9.945600000000)
    pts(448) = Array(55.464962257845, 9.968000000000)
    pts(449) = Array(55.473832263636, 9.990400000000)
    pts(450) = Array(55.482725313796, 10.012800000000)
    pts(451) = Array(55.491641432740, 10.035200000000)
    pts(452) = Array(55.500580644985, 10.057600000000)
    pts(453) = Array(55.509542975144, 10.080000000000)
    pts(454) = Array(55.518528447930, 10.102400000000)
    pts(455) = Array(55.527537088155, 10.124800000000)
    pts(456) = Array(55.536568920730, 10.147200000000)
    pts(457) = Array(55.545623970667, 10.169600000000)
    pts(458) = Array(55.554702263077, 10.192000000000)
    pts(459) = Array(55.563803823173, 10.214400000000)
    pts(460) = Array(55.572928676269, 10.236800000000)
    pts(461) = Array(55.582076847780, 10.259200000000)
    pts(462) = Array(55.591248363224, 10.281600000000)
    pts(463) = Array(55.600443248221, 10.304000000000)
    pts(464) = Array(55.609661528494, 10.326400000000)
    pts(465) = Array(55.618903229869, 10.348800000000)
    pts(466) = Array(55.628168378277, 10.371200000000)
    pts(467) = Array(55.637456999752, 10.393600000000)
    pts(468) = Array(55.646769120433, 10.416000000000)
    pts(469) = Array(55.656104766564, 10.438400000000)
    pts(470) = Array(55.665463964496, 10.460800000000)
    pts(471) = Array(55.674846740682, 10.483200000000)
    pts(472) = Array(55.681899312027, 10.500000000000)
    pts(473) = Array(55.681899312027, 10.505600000000)
    pts(474) = Array(55.681899312027, 10.528000000000)
    pts(475) = Array(55.681899312027, 10.550000000000)
    pts(476) = Array(55.681899312027, 10.550400000000)
    pts(477) = Array(55.681899312027, 10.572800000000)
    pts(478) = Array(55.681899312027, 10.595200000000)
    pts(479) = Array(55.681899312027, 10.617600000000)
    pts(480) = Array(55.681899312027, 10.640000000000)
    pts(481) = Array(55.681899312027, 10.662400000000)
    pts(482) = Array(55.681899312027, 10.684800000000)
    pts(483) = Array(55.681899312027, 10.700000000000)
    pts(484) = Array(55.681899312027, 10.707200000000)
    pts(485) = Array(55.681899312027, 10.729600000000)
    pts(486) = Array(55.681899312027, 10.752000000000)
    pts(487) = Array(55.681899312027, 10.774400000000)
    pts(488) = Array(55.681899312027, 10.796800000000)
    pts(489) = Array(55.681899312027, 10.819200000000)
    pts(490) = Array(55.681899312027, 10.841600000000)
    pts(491) = Array(55.681899312027, 10.864000000000)
    pts(492) = Array(55.681899312027, 10.886400000000)
    pts(493) = Array(55.681899312027, 10.908800000000)
    pts(494) = Array(55.681899312027, 10.931200000000)
    pts(495) = Array(55.681899312027, 10.953600000000)
    pts(496) = Array(55.681899312027, 10.976000000000)
    pts(497) = Array(55.681899312027, 10.998400000000)
    pts(498) = Array(55.681899312027, 11.020800000000)
    pts(499) = Array(55.681899312027, 11.043200000000)
    pts(500) = Array(55.681899312027, 11.065600000000)
    pts(501) = Array(55.681899312027, 11.088000000000)
    pts(502) = Array(55.681899312027, 11.110400000000)
    pts(503) = Array(55.681899312027, 11.132800000000)
    pts(504) = Array(55.681899312027, 11.155200000000)
    pts(505) = Array(55.681899312027, 11.177600000000)
    pts(506) = Array(55.681899312027, 11.200000000000)
    pts(507) = Array(58.013404737590, 11.200000000000)
    pts(508) = Array(58.013404737590, 11.177600000000)
    pts(509) = Array(58.013404737590, 11.155200000000)
    pts(510) = Array(58.013404737590, 11.132800000000)
    pts(511) = Array(58.013404737590, 11.110400000000)
    pts(512) = Array(58.013404737590, 11.088000000000)
    pts(513) = Array(58.013404737590, 11.065600000000)
    pts(514) = Array(58.013404737590, 11.043200000000)
    pts(515) = Array(58.013404737590, 11.020800000000)
    pts(516) = Array(58.013404737590, 10.998400000000)
    pts(517) = Array(58.013404737590, 10.976000000000)
    pts(518) = Array(58.013404737590, 10.953600000000)
    pts(519) = Array(58.013404737590, 10.931200000000)
    pts(520) = Array(58.013404737590, 10.908800000000)
    pts(521) = Array(58.013404737590, 10.886400000000)
    pts(522) = Array(58.013404737590, 10.864000000000)
    pts(523) = Array(58.013404737590, 10.841600000000)
    pts(524) = Array(58.013404737590, 10.819200000000)
    pts(525) = Array(58.013404737590, 10.796800000000)
    pts(526) = Array(58.013404737590, 10.774400000000)
    pts(527) = Array(58.013404737590, 10.752000000000)
    pts(528) = Array(58.013404737590, 10.729600000000)
    pts(529) = Array(58.013404737590, 10.707200000000)
    pts(530) = Array(58.013404737590, 10.700000000000)
    pts(531) = Array(58.013404737590, 10.684800000000)
    pts(532) = Array(58.013404737590, 10.662400000000)
    pts(533) = Array(58.013404737590, 10.640000000000)
    pts(534) = Array(58.013404737590, 10.617600000000)
    pts(535) = Array(58.013404737590, 10.595200000000)
    pts(536) = Array(58.013404737590, 10.572800000000)
    pts(537) = Array(58.013404737590, 10.550400000000)
    pts(538) = Array(58.013404737590, 10.550000000000)
    pts(539) = Array(58.013404737590, 10.528000000000)
    pts(540) = Array(58.013404737590, 10.505600000000)
    pts(541) = Array(58.013404737590, 10.500000000000)
    pts(542) = Array(58.021080606919, 10.483200000000)
    pts(543) = Array(58.031292002492, 10.460800000000)
    pts(544) = Array(58.041477033890, 10.438400000000)
    pts(545) = Array(58.051635736140, 10.416000000000)
    pts(546) = Array(58.061768144120, 10.393600000000)
    pts(547) = Array(58.071874292554, 10.371200000000)
    pts(548) = Array(58.081954216016, 10.348800000000)
    pts(549) = Array(58.092007948929, 10.326400000000)
    pts(550) = Array(58.102035525568, 10.304000000000)
    pts(551) = Array(58.112036980059, 10.281600000000)
    pts(552) = Array(58.122012346380, 10.259200000000)
    pts(553) = Array(58.131961658359, 10.236800000000)
    pts(554) = Array(58.141884949683, 10.214400000000)
    pts(555) = Array(58.151782253887, 10.192000000000)
    pts(556) = Array(58.161653604365, 10.169600000000)
    pts(557) = Array(58.171499034364, 10.147200000000)
    pts(558) = Array(58.181318576990, 10.124800000000)
    pts(559) = Array(58.191112265202, 10.102400000000)
    pts(560) = Array(58.200880131821, 10.080000000000)
    pts(561) = Array(58.210622209521, 10.057600000000)
    pts(562) = Array(58.220338530839, 10.035200000000)
    pts(563) = Array(58.230029128171, 10.012800000000)
    pts(564) = Array(58.239694033770, 9.990400000000)
    pts(565) = Array(58.249333279753, 9.968000000000)
    pts(566) = Array(58.258946898098, 9.945600000000)
    pts(567) = Array(58.268534920644, 9.923200000000)
    pts(568) = Array(58.278097379094, 9.900800000000)
    pts(569) = Array(58.287634305012, 9.878400000000)
    pts(570) = Array(58.297145729828, 9.856000000000)
    pts(571) = Array(58.306631684838, 9.833600000000)
    pts(572) = Array(58.316092201199, 9.811200000000)
    pts(573) = Array(58.325527309937, 9.788800000000)
    pts(574) = Array(58.334937041944, 9.766400000000)
    pts(575) = Array(58.344321427979, 9.744000000000)
    pts(576) = Array(58.353680498667, 9.721600000000)
    pts(577) = Array(58.363014284505, 9.699200000000)
    pts(578) = Array(58.372322815854, 9.676800000000)
    pts(579) = Array(58.381606122948, 9.654400000000)
    pts(580) = Array(58.390864235891, 9.632000000000)
    pts(581) = Array(58.400097184655, 9.609600000000)
    pts(582) = Array(58.409304999085, 9.587200000000)
    pts(583) = Array(58.418487708899, 9.564800000000)
    pts(584) = Array(58.427645343684, 9.542400000000)
    pts(585) = Array(58.436777932903, 9.520000000000)
    pts(586) = Array(58.445885505891, 9.497600000000)
    pts(587) = Array(58.454968091857, 9.475200000000)
    pts(588) = Array(58.464025719886, 9.452800000000)
    pts(589) = Array(58.473058418936, 9.430400000000)
    pts(590) = Array(58.482066217841, 9.408000000000)
    pts(591) = Array(58.491049145314, 9.385600000000)
    pts(592) = Array(58.500007229942, 9.363200000000)
    pts(593) = Array(58.508940500190, 9.340800000000)
    pts(594) = Array(58.517848984401, 9.318400000000)
    pts(595) = Array(58.526732710796, 9.296000000000)
    pts(596) = Array(58.535591707475, 9.273600000000)
    pts(597) = Array(58.544426002419, 9.251200000000)
    pts(598) = Array(58.553235623487, 9.228800000000)
    pts(599) = Array(58.562020598420, 9.206400000000)
    pts(600) = Array(58.570780954838, 9.184000000000)
    pts(601) = Array(58.579516720243, 9.161600000000)
    pts(602) = Array(58.588227922022, 9.139200000000)
    pts(603) = Array(58.596914587441, 9.116800000000)
    pts(604) = Array(58.605576743650, 9.094400000000)
    pts(605) = Array(58.614214417683, 9.072000000000)
    pts(606) = Array(58.622827636458, 9.049600000000)
    pts(607) = Array(58.631416426778, 9.027200000000)
    pts(608) = Array(58.639980815329, 9.004800000000)
    pts(609) = Array(58.648520828684, 8.982400000000)
    pts(610) = Array(58.657036493302, 8.960000000000)
    pts(611) = Array(58.665527835529, 8.937600000000)
    pts(612) = Array(58.673994881596, 8.915200000000)
    pts(613) = Array(58.682437657623, 8.892800000000)
    pts(614) = Array(58.690856189618, 8.870400000000)
    pts(615) = Array(58.699250503475, 8.848000000000)
    pts(616) = Array(58.707620624981, 8.825600000000)
    pts(617) = Array(58.715966579807, 8.803200000000)
    pts(618) = Array(58.724288393519, 8.780800000000)
    pts(619) = Array(58.732586091569, 8.758400000000)
    pts(620) = Array(58.740859699302, 8.736000000000)
    pts(621) = Array(58.749109241953, 8.713600000000)
    pts(622) = Array(58.757334744649, 8.691200000000)
    pts(623) = Array(58.765536232409, 8.668800000000)
    pts(624) = Array(58.773713730144, 8.646400000000)
    pts(625) = Array(58.781867262659, 8.624000000000)
    pts(626) = Array(58.789996854650, 8.601600000000)
    pts(627) = Array(58.798102530708, 8.579200000000)
    pts(628) = Array(58.806184315320, 8.556800000000)
    pts(629) = Array(58.814242232864, 8.534400000000)
    pts(630) = Array(58.822276307616, 8.512000000000)
    pts(631) = Array(58.830286563745, 8.489600000000)
    pts(632) = Array(58.838273025317, 8.467200000000)
    pts(633) = Array(58.846235716296, 8.444800000000)
    pts(634) = Array(58.854174660539, 8.422400000000)
    pts(635) = Array(58.862089881803, 8.400000000000)
    pts(636) = Array(58.869981403740, 8.377600000000)
    pts(637) = Array(58.877849249904, 8.355200000000)
    pts(638) = Array(58.885693443742, 8.332800000000)
    pts(639) = Array(58.893514008604, 8.310400000000)
    pts(640) = Array(58.901310967737, 8.288000000000)
    pts(641) = Array(58.909084344287, 8.265600000000)
    pts(642) = Array(58.916834161301, 8.243200000000)
    pts(643) = Array(58.924560441727, 8.220800000000)
    pts(644) = Array(58.932263208412, 8.198400000000)
    pts(645) = Array(58.939942484105, 8.176000000000)
    pts(646) = Array(58.947598291456, 8.153600000000)
    pts(647) = Array(58.955230653017, 8.131200000000)
    pts(648) = Array(58.962839591242, 8.108800000000)
    pts(649) = Array(58.970425128487, 8.086400000000)
    pts(650) = Array(58.977987287012, 8.064000000000)
    pts(651) = Array(58.985526088981, 8.041600000000)
    pts(652) = Array(58.993041556460, 8.019200000000)
    pts(653) = Array(59.000533711418, 7.996800000000)
    pts(654) = Array(59.008002575733, 7.974400000000)
    pts(655) = Array(59.015448171182, 7.952000000000)
    pts(656) = Array(59.022870519452, 7.929600000000)
    pts(657) = Array(59.030269642131, 7.907200000000)
    pts(658) = Array(59.037645560718, 7.884800000000)
    pts(659) = Array(59.044998296614, 7.862400000000)
    pts(660) = Array(59.052327871128, 7.840000000000)
    pts(661) = Array(59.059634305476, 7.817600000000)
    pts(662) = Array(59.066917620781, 7.795200000000)
    pts(663) = Array(59.074177838073, 7.772800000000)
    pts(664) = Array(59.081414978293, 7.750400000000)
    pts(665) = Array(59.088629062285, 7.728000000000)
    pts(666) = Array(59.095820110806, 7.705600000000)
    pts(667) = Array(59.102988144520, 7.683200000000)
    pts(668) = Array(59.110133184001, 7.660800000000)
    pts(669) = Array(59.117255249732, 7.638400000000)
    pts(670) = Array(59.124354362106, 7.616000000000)
    pts(671) = Array(59.131430541428, 7.593600000000)
    pts(672) = Array(59.138483807910, 7.571200000000)
    pts(673) = Array(59.145514181678, 7.548800000000)
    pts(674) = Array(59.152521682768, 7.526400000000)
    pts(675) = Array(59.159506331128, 7.504000000000)
    pts(676) = Array(59.160827216406, 7.499755623002)
    pts(677) = Array(59.161170636729, 7.498651685025)
    pts(678) = Array(59.161553831582, 7.497419679825)
    pts(679) = Array(59.166468146618, 7.481600000000)
    pts(680) = Array(59.173407149008, 7.459200000000)
    pts(681) = Array(59.180323357985, 7.436800000000)
    pts(682) = Array(59.187216793145, 7.414400000000)
    pts(683) = Array(59.194087473997, 7.392000000000)
    pts(684) = Array(59.200935419967, 7.369600000000)
    pts(685) = Array(59.207760650390, 7.347200000000)
    pts(686) = Array(59.214563184520, 7.324800000000)
    pts(687) = Array(59.221343041521, 7.302400000000)
    pts(688) = Array(59.228100240474, 7.280000000000)
    pts(689) = Array(59.234834800374, 7.257600000000)
    pts(690) = Array(59.241546740132, 7.235200000000)
    pts(691) = Array(59.248236078575, 7.212800000000)
    pts(692) = Array(59.254902834444, 7.190400000000)
    pts(693) = Array(59.261547026397, 7.168000000000)
    pts(694) = Array(59.268168673009, 7.145600000000)
    pts(695) = Array(59.274767792771, 7.123200000000)
    pts(696) = Array(59.281344404092, 7.100800000000)
    pts(697) = Array(59.287898525296, 7.078400000000)
    pts(698) = Array(59.294430174626, 7.056000000000)
    pts(699) = Array(59.300939370245, 7.033600000000)
    pts(700) = Array(59.307426130230, 7.011200000000)
    pts(701) = Array(59.313890472578, 6.988800000000)
    pts(702) = Array(59.320332415207, 6.966400000000)
    pts(703) = Array(59.326751975950, 6.944000000000)
    pts(704) = Array(59.333149172562, 6.921600000000)
    pts(705) = Array(59.339524022716, 6.899200000000)
    pts(706) = Array(59.345876544006, 6.876800000000)
    pts(707) = Array(59.352206753945, 6.854400000000)
    pts(708) = Array(59.358514669967, 6.832000000000)
    pts(709) = Array(59.364800309425, 6.809600000000)
    pts(710) = Array(59.371063689596, 6.787200000000)
    pts(711) = Array(59.377304827674, 6.764800000000)
    pts(712) = Array(59.383523740778, 6.742400000000)
    pts(713) = Array(59.389720445945, 6.720000000000)
    pts(714) = Array(59.395894960138, 6.697600000000)
    pts(715) = Array(59.402047300238, 6.675200000000)
    pts(716) = Array(59.408177483051, 6.652800000000)
    pts(717) = Array(59.414285525304, 6.630400000000)
    pts(718) = Array(59.420371443649, 6.608000000000)
    pts(719) = Array(59.426435254659, 6.585600000000)
    pts(720) = Array(59.432476974831, 6.563200000000)
    pts(721) = Array(59.438496620585, 6.540800000000)
    pts(722) = Array(59.444494208266, 6.518400000000)
    pts(723) = Array(59.450469754143, 6.496000000000)
    pts(724) = Array(59.456423274409, 6.473600000000)
    pts(725) = Array(59.462354785180, 6.451200000000)
    pts(726) = Array(59.468264302499, 6.428800000000)
    pts(727) = Array(59.474151842332, 6.406400000000)
    pts(728) = Array(59.480017420574, 6.384000000000)
    pts(729) = Array(59.485861053040, 6.361600000000)
    pts(730) = Array(59.491682755476, 6.339200000000)
    pts(731) = Array(59.497482543550, 6.316800000000)
    pts(732) = Array(59.503260432858, 6.294400000000)
    pts(733) = Array(59.509016438923, 6.272000000000)
    pts(734) = Array(59.514750577192, 6.249600000000)
    pts(735) = Array(59.520462863043, 6.227200000000)
    pts(736) = Array(59.526153311776, 6.204800000000)
    pts(737) = Array(59.531821938623, 6.182400000000)
    pts(738) = Array(59.537468758741, 6.160000000000)
    pts(739) = Array(59.543093787215, 6.137600000000)
    pts(740) = Array(59.548697039058, 6.115200000000)
    pts(741) = Array(59.554278529212, 6.092800000000)
    pts(742) = Array(59.559838272547, 6.070400000000)
    pts(743) = Array(59.565376283860, 6.048000000000)
    pts(744) = Array(59.570892577879, 6.025600000000)
    pts(745) = Array(59.576387169260, 6.003200000000)
    pts(746) = Array(59.581860072589, 5.980800000000)
    pts(747) = Array(59.587311302380, 5.958400000000)
    pts(748) = Array(59.592740873078, 5.936000000000)
    pts(749) = Array(59.598148799057, 5.913600000000)
    pts(750) = Array(59.603535094621, 5.891200000000)
    pts(751) = Array(59.608899774006, 5.868800000000)
    pts(752) = Array(59.614242851377, 5.846400000000)
    pts(753) = Array(59.619564340828, 5.824000000000)
    pts(754) = Array(59.624864256388, 5.801600000000)
    pts(755) = Array(59.630142612012, 5.779200000000)
    pts(756) = Array(59.635399421592, 5.756800000000)
    pts(757) = Array(59.640634698946, 5.734400000000)
    pts(758) = Array(59.645848457828, 5.712000000000)
    pts(759) = Array(59.651040711920, 5.689600000000)
    pts(760) = Array(59.656211474839, 5.667200000000)
    pts(761) = Array(59.661360760133, 5.644800000000)
    pts(762) = Array(59.666488581283, 5.622400000000)
    pts(763) = Array(59.671594951701, 5.600000000000)
    pts(764) = Array(59.676679884734, 5.577600000000)
    pts(765) = Array(59.681743393661, 5.555200000000)
    pts(766) = Array(59.686785491693, 5.532800000000)
    pts(767) = Array(59.691806191976, 5.510400000000)
    pts(768) = Array(59.696805507589, 5.488000000000)
    pts(769) = Array(59.701783451545, 5.465600000000)
    pts(770) = Array(59.706740036790, 5.443200000000)
    pts(771) = Array(59.711675276205, 5.420800000000)
    pts(772) = Array(59.716589182605, 5.398400000000)
    pts(773) = Array(59.721481768738, 5.376000000000)
    pts(774) = Array(59.726353047289, 5.353600000000)
    pts(775) = Array(59.731203030877, 5.331200000000)
    pts(776) = Array(59.736031732054, 5.308800000000)
    pts(777) = Array(59.740839163310, 5.286400000000)
    pts(778) = Array(59.745625337068, 5.264000000000)
    pts(779) = Array(59.750390265689, 5.241600000000)
    pts(780) = Array(59.755133961466, 5.219200000000)
    pts(781) = Array(59.759856436631, 5.196800000000)
    pts(782) = Array(59.764557703350, 5.174400000000)
    pts(783) = Array(59.769237773727, 5.152000000000)
    pts(784) = Array(59.773896659800, 5.129600000000)
    pts(785) = Array(59.778534373546, 5.107200000000)
    pts(786) = Array(59.783150926876, 5.084800000000)
    pts(787) = Array(59.787746331639, 5.062400000000)
    pts(788) = Array(59.792320599622, 5.040000000000)
    pts(789) = Array(59.796873742548, 5.017600000000)
    pts(790) = Array(59.801405772076, 4.995200000000)
    pts(791) = Array(59.805916699805, 4.972800000000)
    pts(792) = Array(59.810406537269, 4.950400000000)
    pts(793) = Array(59.814875295943, 4.928000000000)
    pts(794) = Array(59.819322987236, 4.905600000000)
    pts(795) = Array(59.823749622497, 4.883200000000)
    pts(796) = Array(59.828155213014, 4.860800000000)
    pts(797) = Array(59.832539770013, 4.838400000000)
    pts(798) = Array(59.836903304656, 4.816000000000)
    pts(799) = Array(59.841245828046, 4.793600000000)
    pts(800) = Array(59.845567351226, 4.771200000000)
    pts(801) = Array(59.849867885174, 4.748800000000)
    pts(802) = Array(59.854147440811, 4.726400000000)
    pts(803) = Array(59.858406028995, 4.704000000000)
    pts(804) = Array(59.862643660523, 4.681600000000)
    pts(805) = Array(59.866860346134, 4.659200000000)
    pts(806) = Array(59.871056096503, 4.636800000000)
    pts(807) = Array(59.875230922249, 4.614400000000)
    pts(808) = Array(59.879384833927, 4.592000000000)
    pts(809) = Array(59.883517842035, 4.569600000000)
    pts(810) = Array(59.887629957010, 4.547200000000)
    pts(811) = Array(59.891721189228, 4.524800000000)
    pts(812) = Array(59.895791549008, 4.502400000000)
    pts(813) = Array(59.899841046609, 4.480000000000)
    pts(814) = Array(59.903869692229, 4.457600000000)
    pts(815) = Array(59.907877496009, 4.435200000000)
    pts(816) = Array(59.911864468029, 4.412800000000)
    pts(817) = Array(59.915830618312, 4.390400000000)
    pts(818) = Array(59.919775956822, 4.368000000000)
    pts(819) = Array(59.923700493463, 4.345600000000)
    pts(820) = Array(59.927604238083, 4.323200000000)
    pts(821) = Array(59.931487200468, 4.300800000000)
    pts(822) = Array(59.935349390350, 4.278400000000)
    pts(823) = Array(59.939190817400, 4.256000000000)
    pts(824) = Array(59.943011491232, 4.233600000000)
    pts(825) = Array(59.946811421402, 4.211200000000)
    pts(826) = Array(59.950590617409, 4.188800000000)
    pts(827) = Array(59.954349088693, 4.166400000000)
    pts(828) = Array(59.958086844638, 4.144000000000)
    pts(829) = Array(59.961803894569, 4.121600000000)
    pts(830) = Array(59.965500247757, 4.099200000000)
    pts(831) = Array(59.969175913412, 4.076800000000)
    pts(832) = Array(59.972830900689, 4.054400000000)
    pts(833) = Array(59.976465218686, 4.032000000000)
    pts(834) = Array(59.980078876445, 4.009600000000)
    pts(835) = Array(59.983671882949, 3.987200000000)
    pts(836) = Array(59.987244247127, 3.964800000000)
    pts(837) = Array(59.990795977851, 3.942400000000)
    pts(838) = Array(59.994327083935, 3.920000000000)
    pts(839) = Array(59.997837574140, 3.897600000000)
    pts(840) = Array(60.001327457168, 3.875200000000)
    pts(841) = Array(60.004796741666, 3.852800000000)
    pts(842) = Array(60.008245436226, 3.830400000000)
    pts(843) = Array(60.011673549383, 3.808000000000)
    pts(844) = Array(60.015081089618, 3.785600000000)
    pts(845) = Array(60.018468065355, 3.763200000000)
    pts(846) = Array(60.021834484963, 3.740800000000)
    pts(847) = Array(60.025180356755, 3.718400000000)
    pts(848) = Array(60.028505688991, 3.696000000000)
    pts(849) = Array(60.031810489875, 3.673600000000)
    pts(850) = Array(60.035094767553, 3.651200000000)
    pts(851) = Array(60.038358530122, 3.628800000000)
    pts(852) = Array(60.041601785618, 3.606400000000)
    pts(853) = Array(60.044824542027, 3.584000000000)
    pts(854) = Array(60.048026807278, 3.561600000000)
    pts(855) = Array(60.051208589246, 3.539200000000)
    pts(856) = Array(60.054369895753, 3.516800000000)
    pts(857) = Array(60.057510734565, 3.494400000000)
    pts(858) = Array(60.060631113394, 3.472000000000)
    pts(859) = Array(60.063731039898, 3.449600000000)
    pts(860) = Array(60.066810521683, 3.427200000000)
    pts(861) = Array(60.069869566297, 3.404800000000)
    pts(862) = Array(60.072908181239, 3.382400000000)
    pts(863) = Array(60.075926373951, 3.360000000000)
    pts(864) = Array(60.078924151822, 3.337600000000)
    pts(865) = Array(60.081901522188, 3.315200000000)
    pts(866) = Array(60.084858492332, 3.292800000000)
    pts(867) = Array(60.087795069483, 3.270400000000)
    pts(868) = Array(60.090711260817, 3.248000000000)
    pts(869) = Array(60.093607073456, 3.225600000000)
    pts(870) = Array(60.096482514470, 3.203200000000)
    pts(871) = Array(60.099337590876, 3.180800000000)
    pts(872) = Array(60.102172309638, 3.158400000000)
    pts(873) = Array(60.104986677666, 3.136000000000)
    pts(874) = Array(60.107780701820, 3.113600000000)
    pts(875) = Array(60.110554388904, 3.091200000000)
    pts(876) = Array(60.113307745672, 3.068800000000)
    pts(877) = Array(60.116040778824, 3.046400000000)
    pts(878) = Array(60.118753495009, 3.024000000000)
    pts(879) = Array(60.121445900824, 3.001600000000)
    pts(880) = Array(60.124118002811, 2.979200000000)
    pts(881) = Array(60.126769807462, 2.956800000000)
    pts(882) = Array(60.129401321218, 2.934400000000)
    pts(883) = Array(60.132012550466, 2.912000000000)
    pts(884) = Array(60.134603501541, 2.889600000000)
    pts(885) = Array(60.137174180727, 2.867200000000)
    pts(886) = Array(60.139724594258, 2.844800000000)
    pts(887) = Array(60.142254748313, 2.822400000000)
    pts(888) = Array(60.144764649022, 2.800000000000)
    pts(889) = Array(60.147254302462, 2.777600000000)
    pts(890) = Array(60.149723714659, 2.755200000000)
    pts(891) = Array(60.152172891587, 2.732800000000)
    pts(892) = Array(60.154601839171, 2.710400000000)
    pts(893) = Array(60.157010563283, 2.688000000000)
    pts(894) = Array(60.159399069743, 2.665600000000)
    pts(895) = Array(60.161767364322, 2.643200000000)
    pts(896) = Array(60.164115452739, 2.620800000000)
    pts(897) = Array(60.166443340662, 2.598400000000)
    pts(898) = Array(60.168751033708, 2.576000000000)
    pts(899) = Array(60.171038537444, 2.553600000000)
    pts(900) = Array(60.173305857385, 2.531200000000)
    pts(901) = Array(60.175552998996, 2.508800000000)
    pts(902) = Array(60.177779967693, 2.486400000000)
    pts(903) = Array(60.179986768839, 2.464000000000)
    pts(904) = Array(60.182173407747, 2.441600000000)
    pts(905) = Array(60.184339889681, 2.419200000000)
    pts(906) = Array(60.186486219854, 2.396800000000)
    pts(907) = Array(60.188612403428, 2.374400000000)
    pts(908) = Array(60.190718445516, 2.352000000000)
    pts(909) = Array(60.192804351180, 2.329600000000)
    pts(910) = Array(60.194870125432, 2.307200000000)
    pts(911) = Array(60.196915773234, 2.284800000000)
    pts(912) = Array(60.198941299498, 2.262400000000)
    pts(913) = Array(60.200946709088, 2.240000000000)
    pts(914) = Array(60.202932006814, 2.217600000000)
    pts(915) = Array(60.204897197441, 2.195200000000)
    pts(916) = Array(60.206842285681, 2.172800000000)
    pts(917) = Array(60.208767276197, 2.150400000000)
    pts(918) = Array(60.210672173603, 2.128000000000)
    pts(919) = Array(60.212556982463, 2.105600000000)
    pts(920) = Array(60.214421707293, 2.083200000000)
    pts(921) = Array(60.216266352556, 2.060800000000)
    pts(922) = Array(60.218090922670, 2.038400000000)
    pts(923) = Array(60.219895422000, 2.016000000000)
    pts(924) = Array(60.221679854864, 1.993600000000)
    pts(925) = Array(60.223444225530, 1.971200000000)
    pts(926) = Array(60.225188538217, 1.948800000000)
    pts(927) = Array(60.226912797094, 1.926400000000)
    pts(928) = Array(60.228617006283, 1.904000000000)
    pts(929) = Array(60.230301169854, 1.881600000000)
    pts(930) = Array(60.231965291830, 1.859200000000)
    pts(931) = Array(60.233609376186, 1.836800000000)
    pts(932) = Array(60.235233426846, 1.814400000000)
    pts(933) = Array(60.236837447686, 1.792000000000)
    pts(934) = Array(60.238421442533, 1.769600000000)
    pts(935) = Array(60.239985415166, 1.747200000000)
    pts(936) = Array(60.241529369315, 1.724800000000)
    pts(937) = Array(60.243053308660, 1.702400000000)
    pts(938) = Array(60.244557236835, 1.680000000000)
    pts(939) = Array(60.246041157423, 1.657600000000)
    pts(940) = Array(60.247505073959, 1.635200000000)
    pts(941) = Array(60.248948989932, 1.612800000000)
    pts(942) = Array(60.250372908779, 1.590400000000)
    pts(943) = Array(60.251776833890, 1.568000000000)
    pts(944) = Array(60.253160768607, 1.545600000000)
    pts(945) = Array(60.254524716224, 1.523200000000)
    pts(946) = Array(60.255868679987, 1.500800000000)
    pts(947) = Array(60.257192663091, 1.478400000000)
    pts(948) = Array(60.258496668686, 1.456000000000)
    pts(949) = Array(60.259780699872, 1.433600000000)
    pts(950) = Array(60.261044759702, 1.411200000000)
    pts(951) = Array(60.262288851181, 1.388800000000)
    pts(952) = Array(60.263512977265, 1.366400000000)
    pts(953) = Array(60.264717140861, 1.344000000000)
    pts(954) = Array(60.265901344831, 1.321600000000)
    pts(955) = Array(60.267065591987, 1.299200000000)
    pts(956) = Array(60.268209885094, 1.276800000000)
    pts(957) = Array(60.269334226867, 1.254400000000)
    pts(958) = Array(60.270438619977, 1.232000000000)
    pts(959) = Array(60.271523067043, 1.209600000000)
    pts(960) = Array(60.272587570639, 1.187200000000)
    pts(961) = Array(60.273632133291, 1.164800000000)
    pts(962) = Array(60.274656757477, 1.142400000000)
    pts(963) = Array(60.275661445626, 1.120000000000)
    pts(964) = Array(60.276646200120, 1.097600000000)
    pts(965) = Array(60.277611023295, 1.075200000000)
    pts(966) = Array(60.278555917437, 1.052800000000)
    pts(967) = Array(60.279480884787, 1.030400000000)
    pts(968) = Array(60.280385927535, 1.008000000000)
    pts(969) = Array(60.281271047828, 0.985600000000)
    pts(970) = Array(60.282136247760, 0.963200000000)
    pts(971) = Array(60.282981529383, 0.940800000000)
    pts(972) = Array(60.283806894698, 0.918400000000)
    pts(973) = Array(60.284612345659, 0.896000000000)
    pts(974) = Array(60.285397884175, 0.873600000000)
    pts(975) = Array(60.286163512104, 0.851200000000)
    pts(976) = Array(60.286909231259, 0.828800000000)
    pts(977) = Array(60.287635043405, 0.806400000000)
    pts(978) = Array(60.288340950260, 0.784000000000)
    pts(979) = Array(60.289026953495, 0.761600000000)
    pts(980) = Array(60.289693054732, 0.739200000000)
    pts(981) = Array(60.290339255548, 0.716800000000)
    pts(982) = Array(60.290965557471, 0.694400000000)
    pts(983) = Array(60.291571961983, 0.672000000000)
    pts(984) = Array(60.292158470518, 0.649600000000)
    pts(985) = Array(60.292725084463, 0.627200000000)
    pts(986) = Array(60.293271805158, 0.604800000000)
    pts(987) = Array(60.293798633896, 0.582400000000)
    pts(988) = Array(60.294305571922, 0.560000000000)
    pts(989) = Array(60.294792620436, 0.537600000000)
    pts(990) = Array(60.295259780587, 0.515200000000)
    pts(991) = Array(60.295707053481, 0.492800000000)
    pts(992) = Array(60.296134440175, 0.470400000000)
    pts(993) = Array(60.296541941679, 0.448000000000)
    pts(994) = Array(60.296929558955, 0.425600000000)
    pts(995) = Array(60.297297292921, 0.403200000000)
    pts(996) = Array(60.297645144444, 0.380800000000)
    pts(997) = Array(60.297973114347, 0.358400000000)
    pts(998) = Array(60.298281203405, 0.336000000000)
    pts(999) = Array(60.298569412345, 0.313600000000)
    pts(1000) = Array(60.298837741849, 0.291200000000)
    pts(1001) = Array(60.299086192550, 0.268800000000)
    pts(1002) = Array(60.299314765036, 0.246400000000)
    pts(1003) = Array(60.299523459846, 0.224000000000)
    pts(1004) = Array(60.299712277473, 0.201600000000)
    pts(1005) = Array(60.299881218363, 0.179200000000)
    pts(1006) = Array(60.300030282916, 0.156800000000)
    pts(1007) = Array(60.300159471483, 0.134400000000)
    pts(1008) = Array(60.300268784369, 0.112000000000)
    pts(1009) = Array(60.300358221833, 0.089600000000)
    pts(1010) = Array(60.300427784087, 0.067200000000)
    pts(1011) = Array(60.300477471293, 0.044800000000)
    pts(1012) = Array(60.300507283570, 0.022400000000)
    pts(1013) = Array(60.300517220987, 0.000000000000)
    Profile_16 = pts
End Function

Private Function Profile_17() As Variant
    Dim pts(61) As Variant
    pts(0) = Array(60.389517220987, 0.000000000000)
    pts(1) = Array(60.388510206901, 1.548716000000)
    pts(2) = Array(60.385558309633, 3.070730000000)
    pts(3) = Array(60.380557938820, 4.619446000000)
    pts(4) = Array(60.373681496607, 6.141460000000)
    pts(5) = Array(60.363818038903, 7.823686000000)
    pts(6) = Array(60.351791029926, 9.479210000000)
    pts(7) = Array(60.337212429445, 11.161436000000)
    pts(8) = Array(60.320544535343, 12.817000000000)
    pts(9) = Array(60.320544535343, 13.351000000000)
    pts(10) = Array(62.314166547745, 13.351000000000)
    pts(11) = Array(62.414454514184, 13.110682000000)
    pts(12) = Array(62.512613703237, 12.870364000000)
    pts(13) = Array(62.608663736002, 12.630046000000)
    pts(14) = Array(62.702623579236, 12.389728000000)
    pts(15) = Array(62.794511567523, 12.149410000000)
    pts(16) = Array(62.884345424339, 11.909092000000)
    pts(17) = Array(62.972142282052, 11.668774000000)
    pts(18) = Array(63.067325396232, 11.401754000000)
    pts(19) = Array(63.187370282350, 11.054628000000)
    pts(20) = Array(63.303282719471, 10.707502000000)
    pts(21) = Array(63.415106578204, 10.360376000000)
    pts(22) = Array(63.522883595075, 10.013250000000)
    pts(23) = Array(63.626653460228, 9.666124000000)
    pts(24) = Array(63.726453899125, 9.318998000000)
    pts(25) = Array(63.822320748669, 8.971872000000)
    pts(26) = Array(63.914288028128, 8.624746000000)
    pts(27) = Array(64.002388005243, 8.277620000000)
    pts(28) = Array(64.086651257814, 7.930494000000)
    pts(29) = Array(64.167106731076, 7.583368000000)
    pts(30) = Array(64.208861375917, 7.396454000000)
    pts(31) = Array(64.249524038319, 7.209540000000)
    pts(32) = Array(64.322156735627, 6.862414000000)
    pts(33) = Array(64.391061023148, 6.515288000000)
    pts(34) = Array(64.456259705164, 6.168162000000)
    pts(35) = Array(64.489839379703, 5.981248000000)
    pts(36) = Array(64.522354055731, 5.794334000000)
    pts(37) = Array(64.579923321221, 5.447208000000)
    pts(38) = Array(64.609411654307, 5.260294000000)
    pts(39) = Array(64.637846202496, 5.073380000000)
    pts(40) = Array(64.687866404095, 4.726254000000)
    pts(41) = Array(64.713304343192, 4.539340000000)
    pts(42) = Array(64.737698113085, 4.352426000000)
    pts(43) = Array(64.780238152849, 4.005300000000)
    pts(44) = Array(64.801660583041, 3.818386000000)
    pts(45) = Array(64.822046929913, 3.631472000000)
    pts(46) = Array(64.857164747569, 3.284346000000)
    pts(47) = Array(64.874600736966, 3.097432000000)
    pts(48) = Array(64.891007255336, 2.910518000000)
    pts(49) = Array(64.918750231292, 2.563392000000)
    pts(50) = Array(64.932223231974, 2.376478000000)
    pts(51) = Array(64.944671945589, 2.189564000000)
    pts(52) = Array(64.965077225397, 1.842438000000)
    pts(53) = Array(64.974605233193, 1.655524000000)
    pts(54) = Array(64.983112744590, 1.468610000000)
    pts(55) = Array(64.996207490111, 1.121484000000)
    pts(56) = Array(65.001803165797, 0.934570000000)
    pts(57) = Array(65.006380767833, 0.747656000000)
    pts(58) = Array(65.009940703230, 0.560742000000)
    pts(59) = Array(65.012483288365, 0.373828000000)
    pts(60) = Array(65.014008749121, 0.186914000000)
    pts(61) = Array(65.014517220987, 0.000000000000)
    Profile_17 = pts
End Function

Private Function Profile_18() As Variant
    Dim pts(9) As Variant
    pts(0) = Array(21.000000000000, 0.000000000000)
    pts(1) = Array(21.000000000000, 1.500000000000)
    pts(2) = Array(22.600000000000, 1.500000000000)
    pts(3) = Array(22.600000000000, 2.000000000000)
    pts(4) = Array(26.850000000000, 2.000000000000)
    pts(5) = Array(26.850000000000, 1.500000000000)
    pts(6) = Array(30.200000000000, 1.500000000000)
    pts(7) = Array(30.200000000000, 2.400000000000)
    pts(8) = Array(31.600000000000, 2.400000000000)
    pts(9) = Array(31.600000000000, 0.000000000000)
    Profile_18 = pts
End Function

Private Function Profile_19() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.050000000000, 2.500000000000)
    pts(1) = Array(30.000000000000, 2.500000000000)
    pts(2) = Array(30.000000000000, 1.550000000000)
    pts(3) = Array(27.050000000000, 1.550000000000)
    Profile_19 = pts
End Function

Private Function Profile_20() As Variant
    Dim pts(9) As Variant
    pts(0) = Array(21.000000000000, 0.000000000000)
    pts(1) = Array(21.000000000000, 1.500000000000)
    pts(2) = Array(22.600000000000, 1.500000000000)
    pts(3) = Array(22.600000000000, 2.000000000000)
    pts(4) = Array(26.850000000000, 2.000000000000)
    pts(5) = Array(26.850000000000, 1.500000000000)
    pts(6) = Array(30.200000000000, 1.500000000000)
    pts(7) = Array(30.200000000000, 2.400000000000)
    pts(8) = Array(31.600000000000, 2.400000000000)
    pts(9) = Array(31.600000000000, 0.000000000000)
    Profile_20 = pts
End Function

Private Function Profile_21() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.050000000000, 2.500000000000)
    pts(1) = Array(30.000000000000, 2.500000000000)
    pts(2) = Array(30.000000000000, 1.550000000000)
    pts(3) = Array(27.050000000000, 1.550000000000)
    Profile_21 = pts
End Function

Private Function Profile_22() As Variant
    Dim pts(9) As Variant
    pts(0) = Array(21.000000000000, 0.000000000000)
    pts(1) = Array(21.000000000000, 1.500000000000)
    pts(2) = Array(22.600000000000, 1.500000000000)
    pts(3) = Array(22.600000000000, 2.000000000000)
    pts(4) = Array(26.850000000000, 2.000000000000)
    pts(5) = Array(26.850000000000, 1.500000000000)
    pts(6) = Array(30.200000000000, 1.500000000000)
    pts(7) = Array(30.200000000000, 2.400000000000)
    pts(8) = Array(31.600000000000, 2.400000000000)
    pts(9) = Array(31.600000000000, 0.000000000000)
    Profile_22 = pts
End Function

Private Function Profile_23() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(27.050000000000, 2.500000000000)
    pts(1) = Array(30.000000000000, 2.500000000000)
    pts(2) = Array(30.000000000000, 1.550000000000)
    pts(3) = Array(27.050000000000, 1.550000000000)
    Profile_23 = pts
End Function

Private Sub AddFocusDetails(ByVal model As SldWorks.ModelDoc2, ByVal label As String)
    Select Case label
        Case "01_FixedGuideBarrel"
            FocusLoftCut model, "AxialGuideSlot_1", Array(0, 1)
            FocusLoftCut model, "AxialGuideSlot_2", Array(2, 3)
            FocusLoftCut model, "AxialGuideSlot_3", Array(4, 5)
        Case "03_L2MovingCell"
            FocusLoftCut model, "M3_AxleSeat_1", Array(6, 7)
            FocusLoftCut model, "M3_AxleSeat_2", Array(8, 9)
            FocusLoftCut model, "M3_AxleSeat_3", Array(10, 11)
        Case "10_GearedCamRing"
            FocusLoftCut model, "CamSlot_1", Array(12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, _
                32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, _
                52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, _
                72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, _
                92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104)
            FocusLoftCut model, "CamSlot_2", Array(105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, _
                125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, _
                145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, _
                165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, 183, 184, _
                185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197)
            FocusLoftCut model, "CamSlot_3", Array(198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, _
                218, 219, 220, 221, 222, 223, 224, 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, _
                238, 239, 240, 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256, 257, _
                258, 259, 260, 261, 262, 263, 264, 265, 266, 267, 268, 269, 270, 271, 272, 273, 274, 275, 276, 277, _
                278, 279, 280, 281, 282, 283, 284, 285, 286, 287, 288, 289, 290)
            FocusLoftCut model, "GearSpace_1", Array(291, 292)
            FocusLoftCut model, "GearSpace_2", Array(293, 294)
            FocusLoftCut model, "GearSpace_3", Array(295, 296)
            FocusLoftCut model, "GearSpace_4", Array(297, 298)
            FocusLoftCut model, "GearSpace_5", Array(299, 300)
            FocusLoftCut model, "GearSpace_6", Array(301, 302)
            FocusLoftCut model, "GearSpace_7", Array(303, 304)
            FocusLoftCut model, "GearSpace_8", Array(305, 306)
            FocusLoftCut model, "GearSpace_9", Array(307, 308)
            FocusLoftCut model, "GearSpace_10", Array(309, 310)
            FocusLoftCut model, "GearSpace_11", Array(311, 312)
            FocusLoftCut model, "GearSpace_12", Array(313, 314)
            FocusLoftCut model, "GearSpace_13", Array(315, 316)
            FocusLoftCut model, "GearSpace_14", Array(317, 318)
            FocusLoftCut model, "GearSpace_15", Array(319, 320)
            FocusLoftCut model, "GearSpace_16", Array(321, 322)
            FocusLoftCut model, "GearSpace_17", Array(323, 324)
            FocusLoftCut model, "GearSpace_18", Array(325, 326)
            FocusLoftCut model, "GearSpace_19", Array(327, 328)
            FocusLoftCut model, "GearSpace_20", Array(329, 330)
            FocusLoftCut model, "GearSpace_21", Array(331, 332)
            FocusLoftCut model, "GearSpace_22", Array(333, 334)
            FocusLoftCut model, "GearSpace_23", Array(335, 336)
            FocusLoftCut model, "GearSpace_24", Array(337, 338)
            FocusLoftCut model, "GearSpace_25", Array(339, 340)
            FocusLoftCut model, "GearSpace_26", Array(341, 342)
            FocusLoftCut model, "GearSpace_27", Array(343, 344)
            FocusLoftCut model, "GearSpace_28", Array(345, 346)
            FocusLoftCut model, "GearSpace_29", Array(347, 348)
            FocusLoftCut model, "GearSpace_30", Array(349, 350)
            FocusLoftCut model, "GearSpace_31", Array(351, 352)
            FocusLoftCut model, "GearSpace_32", Array(353, 354)
            FocusLoftCut model, "GearSpace_33", Array(355, 356)
            FocusLoftCut model, "GearSpace_34", Array(357, 358)
            FocusLoftCut model, "GearSpace_35", Array(359, 360)
            FocusLoftCut model, "GearSpace_36", Array(361, 362)
            FocusLoftCut model, "GearSpace_37", Array(363, 364)
            FocusLoftCut model, "GearSpace_38", Array(365, 366)
            FocusLoftCut model, "GearSpace_39", Array(367, 368)
            FocusLoftCut model, "GearSpace_40", Array(369, 370)
            FocusLoftCut model, "GearSpace_41", Array(371, 372)
            FocusLoftCut model, "GearSpace_42", Array(373, 374)
            FocusLoftCut model, "GearSpace_43", Array(375, 376)
            FocusLoftCut model, "GearSpace_44", Array(377, 378)
            FocusLoftCut model, "GearSpace_45", Array(379, 380)
            FocusLoftCut model, "GearSpace_46", Array(381, 382)
            FocusLoftCut model, "GearSpace_47", Array(383, 384)
            FocusLoftCut model, "GearSpace_48", Array(385, 386)
            FocusLoftCut model, "GearSpace_49", Array(387, 388)
            FocusLoftCut model, "GearSpace_50", Array(389, 390)
            FocusLoftCut model, "GearSpace_51", Array(391, 392)
            FocusLoftCut model, "GearSpace_52", Array(393, 394)
            FocusLoftCut model, "GearSpace_53", Array(395, 396)
            FocusLoftCut model, "GearSpace_54", Array(397, 398)
            FocusLoftCut model, "GearSpace_55", Array(399, 400)
            FocusLoftCut model, "GearSpace_56", Array(401, 402)
            FocusLoftCut model, "GearSpace_57", Array(403, 404)
            FocusLoftCut model, "GearSpace_58", Array(405, 406)
            FocusLoftCut model, "GearSpace_59", Array(407, 408)
            FocusLoftCut model, "GearSpace_60", Array(409, 410)
            FocusLoftCut model, "GearSpace_61", Array(411, 412)
            FocusLoftCut model, "GearSpace_62", Array(413, 414)
            FocusLoftCut model, "GearSpace_63", Array(415, 416)
            FocusLoftCut model, "GearSpace_64", Array(417, 418)
            FocusLoftCut model, "GearSpace_65", Array(419, 420)
            FocusLoftCut model, "GearSpace_66", Array(421, 422)
            FocusLoftCut model, "GearSpace_67", Array(423, 424)
            FocusLoftCut model, "GearSpace_68", Array(425, 426)
            FocusLoftCut model, "GearSpace_69", Array(427, 428)
            FocusLoftCut model, "GearSpace_70", Array(429, 430)
            FocusLoftCut model, "GearSpace_71", Array(431, 432)
            FocusLoftCut model, "GearSpace_72", Array(433, 434)
            FocusLoftCut model, "GearSpace_73", Array(435, 436)
            FocusLoftCut model, "GearSpace_74", Array(437, 438)
            FocusLoftCut model, "GearSpace_75", Array(439, 440)
            FocusLoftCut model, "GearSpace_76", Array(441, 442)
            FocusLoftCut model, "GearSpace_77", Array(443, 444)
            FocusLoftCut model, "GearSpace_78", Array(445, 446)
            FocusLoftCut model, "GearSpace_79", Array(447, 448)
            FocusLoftCut model, "GearSpace_80", Array(449, 450)
        Case "02_L1Retainer"
            FocusLoftCut model, "WrenchSlot_M1", Array(451, 452)
            FocusLoftCut model, "WrenchSlot_1", Array(453, 454)
        Case "04_L2Retainer"
            FocusLoftCut model, "WrenchSlot_M1", Array(455, 456)
            FocusLoftCut model, "WrenchSlot_1", Array(457, 458)
        Case "09_L7Retainer"
            FocusLoftCut model, "WrenchSlot_M1", Array(459, 460)
            FocusLoftCut model, "WrenchSlot_1", Array(461, 462)
        Case "11_CamKeeper"
            FocusLoftCut model, "WrenchSlot_M1", Array(463, 464)
            FocusLoftCut model, "WrenchSlot_1", Array(465, 466)
        Case "14_InputPinion"
            FocusLoftCut model, "GearSpace_1", Array(467, 468)
            FocusLoftCut model, "GearSpace_2", Array(469, 470)
            FocusLoftCut model, "GearSpace_3", Array(471, 472)
            FocusLoftCut model, "GearSpace_4", Array(473, 474)
            FocusLoftCut model, "GearSpace_5", Array(475, 476)
            FocusLoftCut model, "GearSpace_6", Array(477, 478)
            FocusLoftCut model, "GearSpace_7", Array(479, 480)
            FocusLoftCut model, "GearSpace_8", Array(481, 482)
            FocusLoftCut model, "GearSpace_9", Array(483, 484)
            FocusLoftCut model, "GearSpace_10", Array(485, 486)
            FocusLoftCut model, "GearSpace_11", Array(487, 488)
            FocusLoftCut model, "GearSpace_12", Array(489, 490)
            FocusLoftCut model, "GearSpace_13", Array(491, 492)
            FocusLoftCut model, "GearSpace_14", Array(493, 494)
            FocusLoftCut model, "GearSpace_15", Array(495, 496)
            FocusLoftCut model, "GearSpace_16", Array(497, 498)
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
        Case 373: FocusSectionData = FSData_373()
        Case 374: FocusSectionData = FSData_374()
        Case 375: FocusSectionData = FSData_375()
        Case 376: FocusSectionData = FSData_376()
        Case 377: FocusSectionData = FSData_377()
        Case 378: FocusSectionData = FSData_378()
        Case 379: FocusSectionData = FSData_379()
        Case 380: FocusSectionData = FSData_380()
        Case 381: FocusSectionData = FSData_381()
        Case 382: FocusSectionData = FSData_382()
        Case 383: FocusSectionData = FSData_383()
        Case 384: FocusSectionData = FSData_384()
        Case 385: FocusSectionData = FSData_385()
        Case 386: FocusSectionData = FSData_386()
        Case 387: FocusSectionData = FSData_387()
        Case 388: FocusSectionData = FSData_388()
        Case 389: FocusSectionData = FSData_389()
        Case 390: FocusSectionData = FSData_390()
        Case 391: FocusSectionData = FSData_391()
        Case 392: FocusSectionData = FSData_392()
        Case 393: FocusSectionData = FSData_393()
        Case 394: FocusSectionData = FSData_394()
        Case 395: FocusSectionData = FSData_395()
        Case 396: FocusSectionData = FSData_396()
        Case 397: FocusSectionData = FSData_397()
        Case 398: FocusSectionData = FSData_398()
        Case 399: FocusSectionData = FSData_399()
        Case 400: FocusSectionData = FSData_400()
        Case 401: FocusSectionData = FSData_401()
        Case 402: FocusSectionData = FSData_402()
        Case 403: FocusSectionData = FSData_403()
        Case 404: FocusSectionData = FSData_404()
        Case 405: FocusSectionData = FSData_405()
        Case 406: FocusSectionData = FSData_406()
        Case 407: FocusSectionData = FSData_407()
        Case 408: FocusSectionData = FSData_408()
        Case 409: FocusSectionData = FSData_409()
        Case 410: FocusSectionData = FSData_410()
        Case 411: FocusSectionData = FSData_411()
        Case 412: FocusSectionData = FSData_412()
        Case 413: FocusSectionData = FSData_413()
        Case 414: FocusSectionData = FSData_414()
        Case 415: FocusSectionData = FSData_415()
        Case 416: FocusSectionData = FSData_416()
        Case 417: FocusSectionData = FSData_417()
        Case 418: FocusSectionData = FSData_418()
        Case 419: FocusSectionData = FSData_419()
        Case 420: FocusSectionData = FSData_420()
        Case 421: FocusSectionData = FSData_421()
        Case 422: FocusSectionData = FSData_422()
        Case 423: FocusSectionData = FSData_423()
        Case 424: FocusSectionData = FSData_424()
        Case 425: FocusSectionData = FSData_425()
        Case 426: FocusSectionData = FSData_426()
        Case 427: FocusSectionData = FSData_427()
        Case 428: FocusSectionData = FSData_428()
        Case 429: FocusSectionData = FSData_429()
        Case 430: FocusSectionData = FSData_430()
        Case 431: FocusSectionData = FSData_431()
        Case 432: FocusSectionData = FSData_432()
        Case 433: FocusSectionData = FSData_433()
        Case 434: FocusSectionData = FSData_434()
        Case 435: FocusSectionData = FSData_435()
        Case 436: FocusSectionData = FSData_436()
        Case 437: FocusSectionData = FSData_437()
        Case 438: FocusSectionData = FSData_438()
        Case 439: FocusSectionData = FSData_439()
        Case 440: FocusSectionData = FSData_440()
        Case 441: FocusSectionData = FSData_441()
        Case 442: FocusSectionData = FSData_442()
        Case 443: FocusSectionData = FSData_443()
        Case 444: FocusSectionData = FSData_444()
        Case 445: FocusSectionData = FSData_445()
        Case 446: FocusSectionData = FSData_446()
        Case 447: FocusSectionData = FSData_447()
        Case 448: FocusSectionData = FSData_448()
        Case 449: FocusSectionData = FSData_449()
        Case 450: FocusSectionData = FSData_450()
        Case 451: FocusSectionData = FSData_451()
        Case 452: FocusSectionData = FSData_452()
        Case 453: FocusSectionData = FSData_453()
        Case 454: FocusSectionData = FSData_454()
        Case 455: FocusSectionData = FSData_455()
        Case 456: FocusSectionData = FSData_456()
        Case 457: FocusSectionData = FSData_457()
        Case 458: FocusSectionData = FSData_458()
        Case 459: FocusSectionData = FSData_459()
        Case 460: FocusSectionData = FSData_460()
        Case 461: FocusSectionData = FSData_461()
        Case 462: FocusSectionData = FSData_462()
        Case 463: FocusSectionData = FSData_463()
        Case 464: FocusSectionData = FSData_464()
        Case 465: FocusSectionData = FSData_465()
        Case 466: FocusSectionData = FSData_466()
        Case 467: FocusSectionData = FSData_467()
        Case 468: FocusSectionData = FSData_468()
        Case 469: FocusSectionData = FSData_469()
        Case 470: FocusSectionData = FSData_470()
        Case 471: FocusSectionData = FSData_471()
        Case 472: FocusSectionData = FSData_472()
        Case 473: FocusSectionData = FSData_473()
        Case 474: FocusSectionData = FSData_474()
        Case 475: FocusSectionData = FSData_475()
        Case 476: FocusSectionData = FSData_476()
        Case 477: FocusSectionData = FSData_477()
        Case 478: FocusSectionData = FSData_478()
        Case 479: FocusSectionData = FSData_479()
        Case 480: FocusSectionData = FSData_480()
        Case 481: FocusSectionData = FSData_481()
        Case 482: FocusSectionData = FSData_482()
        Case 483: FocusSectionData = FSData_483()
        Case 484: FocusSectionData = FSData_484()
        Case 485: FocusSectionData = FSData_485()
        Case 486: FocusSectionData = FSData_486()
        Case 487: FocusSectionData = FSData_487()
        Case 488: FocusSectionData = FSData_488()
        Case 489: FocusSectionData = FSData_489()
        Case 490: FocusSectionData = FSData_490()
        Case 491: FocusSectionData = FSData_491()
        Case 492: FocusSectionData = FSData_492()
        Case 493: FocusSectionData = FSData_493()
        Case 494: FocusSectionData = FSData_494()
        Case 495: FocusSectionData = FSData_495()
        Case 496: FocusSectionData = FSData_496()
        Case 497: FocusSectionData = FSData_497()
        Case 498: FocusSectionData = FSData_498()
        Case Else: Err.Raise vbObjectError + 809, , "Unknown focus section."
    End Select
End Function

Private Function FSData_0() As Variant
    Dim pts(25) As Variant
    pts(0) = Array(22.400000000000, 2.100000000000, 26.200000000000)
    pts(1) = Array(22.400000000000, 2.028444235207, 26.743519994715)
    pts(2) = Array(22.400000000000, 1.818653347947, 27.250000000000)
    pts(3) = Array(22.400000000000, 1.484924240492, 27.684924240492)
    pts(4) = Array(22.400000000000, 1.050000000000, 28.018653347947)
    pts(5) = Array(22.400000000000, 0.543519994715, 28.228444235207)
    pts(6) = Array(22.400000000000, 0.000000000000, 28.300000000000)
    pts(7) = Array(22.400000000000, -0.543519994715, 28.228444235207)
    pts(8) = Array(22.400000000000, -1.050000000000, 28.018653347947)
    pts(9) = Array(22.400000000000, -1.484924240492, 27.684924240492)
    pts(10) = Array(22.400000000000, -1.818653347947, 27.250000000000)
    pts(11) = Array(22.400000000000, -2.028444235207, 26.743519994715)
    pts(12) = Array(22.400000000000, -2.100000000000, 26.200000000000)
    pts(13) = Array(22.400000000000, -2.100000000000, 21.800000000000)
    pts(14) = Array(22.400000000000, -2.028444235207, 21.256480005285)
    pts(15) = Array(22.400000000000, -1.818653347947, 20.750000000000)
    pts(16) = Array(22.400000000000, -1.484924240492, 20.315075759508)
    pts(17) = Array(22.400000000000, -1.050000000000, 19.981346652053)
    pts(18) = Array(22.400000000000, -0.543519994715, 19.771555764793)
    pts(19) = Array(22.400000000000, -0.000000000000, 19.700000000000)
    pts(20) = Array(22.400000000000, 0.543519994715, 19.771555764793)
    pts(21) = Array(22.400000000000, 1.050000000000, 19.981346652053)
    pts(22) = Array(22.400000000000, 1.484924240492, 20.315075759508)
    pts(23) = Array(22.400000000000, 1.818653347947, 20.750000000000)
    pts(24) = Array(22.400000000000, 2.028444235207, 21.256480005285)
    pts(25) = Array(22.400000000000, 2.100000000000, 21.800000000000)
    FSData_0 = pts
End Function

Private Function FSData_1() As Variant
    Dim pts(25) As Variant
    pts(0) = Array(26.950000000000, 2.100000000000, 26.200000000000)
    pts(1) = Array(26.950000000000, 2.028444235207, 26.743519994715)
    pts(2) = Array(26.950000000000, 1.818653347947, 27.250000000000)
    pts(3) = Array(26.950000000000, 1.484924240492, 27.684924240492)
    pts(4) = Array(26.950000000000, 1.050000000000, 28.018653347947)
    pts(5) = Array(26.950000000000, 0.543519994715, 28.228444235207)
    pts(6) = Array(26.950000000000, 0.000000000000, 28.300000000000)
    pts(7) = Array(26.950000000000, -0.543519994715, 28.228444235207)
    pts(8) = Array(26.950000000000, -1.050000000000, 28.018653347947)
    pts(9) = Array(26.950000000000, -1.484924240492, 27.684924240492)
    pts(10) = Array(26.950000000000, -1.818653347947, 27.250000000000)
    pts(11) = Array(26.950000000000, -2.028444235207, 26.743519994715)
    pts(12) = Array(26.950000000000, -2.100000000000, 26.200000000000)
    pts(13) = Array(26.950000000000, -2.100000000000, 21.800000000000)
    pts(14) = Array(26.950000000000, -2.028444235207, 21.256480005285)
    pts(15) = Array(26.950000000000, -1.818653347947, 20.750000000000)
    pts(16) = Array(26.950000000000, -1.484924240492, 20.315075759508)
    pts(17) = Array(26.950000000000, -1.050000000000, 19.981346652053)
    pts(18) = Array(26.950000000000, -0.543519994715, 19.771555764793)
    pts(19) = Array(26.950000000000, -0.000000000000, 19.700000000000)
    pts(20) = Array(26.950000000000, 0.543519994715, 19.771555764793)
    pts(21) = Array(26.950000000000, 1.050000000000, 19.981346652053)
    pts(22) = Array(26.950000000000, 1.484924240492, 20.315075759508)
    pts(23) = Array(26.950000000000, 1.818653347947, 20.750000000000)
    pts(24) = Array(26.950000000000, 2.028444235207, 21.256480005285)
    pts(25) = Array(26.950000000000, 2.100000000000, 21.800000000000)
    FSData_1 = pts
End Function

Private Function FSData_2() As Variant
    Dim pts(25) As Variant
    pts(0) = Array(-13.018653347947, 18.348969044771, 26.200000000000)
    pts(1) = Array(-12.956684237849, 18.384746927168, 26.743519994715)
    pts(2) = Array(-12.775000000000, 18.489642370798, 27.250000000000)
    pts(3) = Array(-12.485982114961, 18.656506924526, 27.684924240492)
    pts(4) = Array(-12.109326673974, 18.873969044771, 28.018653347947)
    pts(5) = Array(-11.670702122888, 19.127209047414, 28.228444235207)
    pts(6) = Array(-11.200000000000, 19.398969044771, 28.300000000000)
    pts(7) = Array(-10.729297877112, 19.670729042129, 28.228444235207)
    pts(8) = Array(-10.290673326026, 19.923969044771, 28.018653347947)
    pts(9) = Array(-9.914017885039, 20.141431165017, 27.684924240492)
    pts(10) = Array(-9.625000000000, 20.308295718745, 27.250000000000)
    pts(11) = Array(-9.443315762151, 20.413191162375, 26.743519994715)
    pts(12) = Array(-9.381346652053, 20.448969044771, 26.200000000000)
    pts(13) = Array(-9.381346652053, 20.448969044771, 21.800000000000)
    pts(14) = Array(-9.443315762151, 20.413191162375, 21.256480005285)
    pts(15) = Array(-9.625000000000, 20.308295718745, 20.750000000000)
    pts(16) = Array(-9.914017885039, 20.141431165017, 20.315075759508)
    pts(17) = Array(-10.290673326026, 19.923969044771, 19.981346652053)
    pts(18) = Array(-10.729297877112, 19.670729042129, 19.771555764793)
    pts(19) = Array(-11.200000000000, 19.398969044771, 19.700000000000)
    pts(20) = Array(-11.670702122888, 19.127209047414, 19.771555764793)
    pts(21) = Array(-12.109326673974, 18.873969044771, 19.981346652053)
    pts(22) = Array(-12.485982114961, 18.656506924526, 20.315075759508)
    pts(23) = Array(-12.775000000000, 18.489642370798, 20.750000000000)
    pts(24) = Array(-12.956684237849, 18.384746927168, 21.256480005285)
    pts(25) = Array(-13.018653347947, 18.348969044771, 21.800000000000)
    FSData_2 = pts
End Function

Private Function FSData_3() As Variant
    Dim pts(25) As Variant
    pts(0) = Array(-15.293653347947, 22.289384631991, 26.200000000000)
    pts(1) = Array(-15.231684237849, 22.325162514387, 26.743519994715)
    pts(2) = Array(-15.050000000000, 22.430057958017, 27.250000000000)
    pts(3) = Array(-14.760982114961, 22.596922511745, 27.684924240492)
    pts(4) = Array(-14.384326673974, 22.814384631991, 28.018653347947)
    pts(5) = Array(-13.945702122888, 23.067624634633, 28.228444235207)
    pts(6) = Array(-13.475000000000, 23.339384631991, 28.300000000000)
    pts(7) = Array(-13.004297877112, 23.611144629348, 28.228444235207)
    pts(8) = Array(-12.565673326026, 23.864384631991, 28.018653347947)
    pts(9) = Array(-12.189017885039, 24.081846752236, 27.684924240492)
    pts(10) = Array(-11.900000000000, 24.248711305964, 27.250000000000)
    pts(11) = Array(-11.718315762151, 24.353606749594, 26.743519994715)
    pts(12) = Array(-11.656346652053, 24.389384631991, 26.200000000000)
    pts(13) = Array(-11.656346652053, 24.389384631991, 21.800000000000)
    pts(14) = Array(-11.718315762151, 24.353606749594, 21.256480005285)
    pts(15) = Array(-11.900000000000, 24.248711305964, 20.750000000000)
    pts(16) = Array(-12.189017885039, 24.081846752236, 20.315075759508)
    pts(17) = Array(-12.565673326026, 23.864384631991, 19.981346652053)
    pts(18) = Array(-13.004297877112, 23.611144629348, 19.771555764793)
    pts(19) = Array(-13.475000000000, 23.339384631991, 19.700000000000)
    pts(20) = Array(-13.945702122888, 23.067624634633, 19.771555764793)
    pts(21) = Array(-14.384326673974, 22.814384631991, 19.981346652053)
    pts(22) = Array(-14.760982114961, 22.596922511745, 20.315075759508)
    pts(23) = Array(-15.050000000000, 22.430057958017, 20.750000000000)
    pts(24) = Array(-15.231684237849, 22.325162514387, 21.256480005285)
    pts(25) = Array(-15.293653347947, 22.289384631991, 21.800000000000)
    FSData_3 = pts
End Function

Private Function FSData_4() As Variant
    Dim pts(25) As Variant
    pts(0) = Array(-9.381346652053, -20.448969044771, 26.200000000000)
    pts(1) = Array(-9.443315762151, -20.413191162375, 26.743519994715)
    pts(2) = Array(-9.625000000000, -20.308295718745, 27.250000000000)
    pts(3) = Array(-9.914017885039, -20.141431165017, 27.684924240492)
    pts(4) = Array(-10.290673326026, -19.923969044771, 28.018653347947)
    pts(5) = Array(-10.729297877112, -19.670729042129, 28.228444235207)
    pts(6) = Array(-11.200000000000, -19.398969044771, 28.300000000000)
    pts(7) = Array(-11.670702122888, -19.127209047414, 28.228444235207)
    pts(8) = Array(-12.109326673974, -18.873969044771, 28.018653347947)
    pts(9) = Array(-12.485982114961, -18.656506924526, 27.684924240492)
    pts(10) = Array(-12.775000000000, -18.489642370798, 27.250000000000)
    pts(11) = Array(-12.956684237849, -18.384746927168, 26.743519994715)
    pts(12) = Array(-13.018653347947, -18.348969044771, 26.200000000000)
    pts(13) = Array(-13.018653347947, -18.348969044771, 21.800000000000)
    pts(14) = Array(-12.956684237849, -18.384746927168, 21.256480005285)
    pts(15) = Array(-12.775000000000, -18.489642370798, 20.750000000000)
    pts(16) = Array(-12.485982114961, -18.656506924526, 20.315075759508)
    pts(17) = Array(-12.109326673974, -18.873969044771, 19.981346652053)
    pts(18) = Array(-11.670702122888, -19.127209047414, 19.771555764793)
    pts(19) = Array(-11.200000000000, -19.398969044771, 19.700000000000)
    pts(20) = Array(-10.729297877112, -19.670729042129, 19.771555764793)
    pts(21) = Array(-10.290673326026, -19.923969044771, 19.981346652053)
    pts(22) = Array(-9.914017885039, -20.141431165017, 20.315075759508)
    pts(23) = Array(-9.625000000000, -20.308295718745, 20.750000000000)
    pts(24) = Array(-9.443315762151, -20.413191162375, 21.256480005285)
    pts(25) = Array(-9.381346652053, -20.448969044771, 21.800000000000)
    FSData_4 = pts
End Function

Private Function FSData_5() As Variant
    Dim pts(25) As Variant
    pts(0) = Array(-11.656346652053, -24.389384631991, 26.200000000000)
    pts(1) = Array(-11.718315762151, -24.353606749594, 26.743519994715)
    pts(2) = Array(-11.900000000000, -24.248711305964, 27.250000000000)
    pts(3) = Array(-12.189017885039, -24.081846752236, 27.684924240492)
    pts(4) = Array(-12.565673326026, -23.864384631991, 28.018653347947)
    pts(5) = Array(-13.004297877112, -23.611144629348, 28.228444235207)
    pts(6) = Array(-13.475000000000, -23.339384631991, 28.300000000000)
    pts(7) = Array(-13.945702122888, -23.067624634633, 28.228444235207)
    pts(8) = Array(-14.384326673974, -22.814384631991, 28.018653347947)
    pts(9) = Array(-14.760982114961, -22.596922511745, 27.684924240492)
    pts(10) = Array(-15.050000000000, -22.430057958017, 27.250000000000)
    pts(11) = Array(-15.231684237849, -22.325162514387, 26.743519994715)
    pts(12) = Array(-15.293653347947, -22.289384631991, 26.200000000000)
    pts(13) = Array(-15.293653347947, -22.289384631991, 21.800000000000)
    pts(14) = Array(-15.231684237849, -22.325162514387, 21.256480005285)
    pts(15) = Array(-15.050000000000, -22.430057958017, 20.750000000000)
    pts(16) = Array(-14.760982114961, -22.596922511745, 20.315075759508)
    pts(17) = Array(-14.384326673974, -22.814384631991, 19.981346652053)
    pts(18) = Array(-13.945702122888, -23.067624634633, 19.771555764793)
    pts(19) = Array(-13.475000000000, -23.339384631991, 19.700000000000)
    pts(20) = Array(-13.004297877112, -23.611144629348, 19.771555764793)
    pts(21) = Array(-12.565673326026, -23.864384631991, 19.981346652053)
    pts(22) = Array(-12.189017885039, -24.081846752236, 20.315075759508)
    pts(23) = Array(-11.900000000000, -24.248711305964, 20.750000000000)
    pts(24) = Array(-11.718315762151, -24.353606749594, 21.256480005285)
    pts(25) = Array(-11.656346652053, -24.389384631991, 21.800000000000)
    FSData_5 = pts
End Function

Private Function FSData_6() As Variant
    Dim pts(31) As Variant
    pts(0) = Array(20.900000000000, 1.525000000000, 24.000000000000)
    pts(1) = Array(20.900000000000, 1.495697552615, 24.297512741075)
    pts(2) = Array(20.900000000000, 1.408916287080, 24.583592234357)
    pts(3) = Array(20.900000000000, 1.267991158761, 24.847244605355)
    pts(4) = Array(20.900000000000, 1.078337841309, 25.078337841309)
    pts(5) = Array(20.900000000000, 0.847244605355, 25.267991158761)
    pts(6) = Array(20.900000000000, 0.583592234357, 25.408916287080)
    pts(7) = Array(20.900000000000, 0.297512741075, 25.495697552615)
    pts(8) = Array(20.900000000000, 0.000000000000, 25.525000000000)
    pts(9) = Array(20.900000000000, -0.297512741075, 25.495697552615)
    pts(10) = Array(20.900000000000, -0.583592234357, 25.408916287080)
    pts(11) = Array(20.900000000000, -0.847244605355, 25.267991158761)
    pts(12) = Array(20.900000000000, -1.078337841309, 25.078337841309)
    pts(13) = Array(20.900000000000, -1.267991158761, 24.847244605355)
    pts(14) = Array(20.900000000000, -1.408916287080, 24.583592234357)
    pts(15) = Array(20.900000000000, -1.495697552615, 24.297512741075)
    pts(16) = Array(20.900000000000, -1.525000000000, 24.000000000000)
    pts(17) = Array(20.900000000000, -1.495697552615, 23.702487258925)
    pts(18) = Array(20.900000000000, -1.408916287080, 23.416407765643)
    pts(19) = Array(20.900000000000, -1.267991158761, 23.152755394645)
    pts(20) = Array(20.900000000000, -1.078337841309, 22.921662158691)
    pts(21) = Array(20.900000000000, -0.847244605355, 22.732008841239)
    pts(22) = Array(20.900000000000, -0.583592234357, 22.591083712920)
    pts(23) = Array(20.900000000000, -0.297512741075, 22.504302447385)
    pts(24) = Array(20.900000000000, -0.000000000000, 22.475000000000)
    pts(25) = Array(20.900000000000, 0.297512741075, 22.504302447385)
    pts(26) = Array(20.900000000000, 0.583592234357, 22.591083712920)
    pts(27) = Array(20.900000000000, 0.847244605355, 22.732008841239)
    pts(28) = Array(20.900000000000, 1.078337841309, 22.921662158691)
    pts(29) = Array(20.900000000000, 1.267991158761, 23.152755394645)
    pts(30) = Array(20.900000000000, 1.408916287080, 23.416407765643)
    pts(31) = Array(20.900000000000, 1.495697552615, 23.702487258925)
    FSData_6 = pts
End Function

Private Function FSData_7() As Variant
    Dim pts(31) As Variant
    pts(0) = Array(22.700000000000, 1.525000000000, 24.000000000000)
    pts(1) = Array(22.700000000000, 1.495697552615, 24.297512741075)
    pts(2) = Array(22.700000000000, 1.408916287080, 24.583592234357)
    pts(3) = Array(22.700000000000, 1.267991158761, 24.847244605355)
    pts(4) = Array(22.700000000000, 1.078337841309, 25.078337841309)
    pts(5) = Array(22.700000000000, 0.847244605355, 25.267991158761)
    pts(6) = Array(22.700000000000, 0.583592234357, 25.408916287080)
    pts(7) = Array(22.700000000000, 0.297512741075, 25.495697552615)
    pts(8) = Array(22.700000000000, 0.000000000000, 25.525000000000)
    pts(9) = Array(22.700000000000, -0.297512741075, 25.495697552615)
    pts(10) = Array(22.700000000000, -0.583592234357, 25.408916287080)
    pts(11) = Array(22.700000000000, -0.847244605355, 25.267991158761)
    pts(12) = Array(22.700000000000, -1.078337841309, 25.078337841309)
    pts(13) = Array(22.700000000000, -1.267991158761, 24.847244605355)
    pts(14) = Array(22.700000000000, -1.408916287080, 24.583592234357)
    pts(15) = Array(22.700000000000, -1.495697552615, 24.297512741075)
    pts(16) = Array(22.700000000000, -1.525000000000, 24.000000000000)
    pts(17) = Array(22.700000000000, -1.495697552615, 23.702487258925)
    pts(18) = Array(22.700000000000, -1.408916287080, 23.416407765643)
    pts(19) = Array(22.700000000000, -1.267991158761, 23.152755394645)
    pts(20) = Array(22.700000000000, -1.078337841309, 22.921662158691)
    pts(21) = Array(22.700000000000, -0.847244605355, 22.732008841239)
    pts(22) = Array(22.700000000000, -0.583592234357, 22.591083712920)
    pts(23) = Array(22.700000000000, -0.297512741075, 22.504302447385)
    pts(24) = Array(22.700000000000, -0.000000000000, 22.475000000000)
    pts(25) = Array(22.700000000000, 0.297512741075, 22.504302447385)
    pts(26) = Array(22.700000000000, 0.583592234357, 22.591083712920)
    pts(27) = Array(22.700000000000, 0.847244605355, 22.732008841239)
    pts(28) = Array(22.700000000000, 1.078337841309, 22.921662158691)
    pts(29) = Array(22.700000000000, 1.267991158761, 23.152755394645)
    pts(30) = Array(22.700000000000, 1.408916287080, 23.416407765643)
    pts(31) = Array(22.700000000000, 1.495697552615, 23.702487258925)
    FSData_7 = pts
End Function

Private Function FSData_8() As Variant
    Dim pts(31) As Variant
    pts(0) = Array(-11.770688740771, 17.337430939095, 24.000000000000)
    pts(1) = Array(-11.745312076943, 17.352082162787, 24.297512741075)
    pts(2) = Array(-11.670157296417, 17.395472795555, 24.583592234357)
    pts(3) = Array(-11.548112555261, 17.465935359714, 24.847244605355)
    pts(4) = Array(-11.383867964436, 17.560762018440, 25.078337841309)
    pts(5) = Array(-11.183735351457, 17.676308636417, 25.267991158761)
    pts(6) = Array(-10.955405700404, 17.808134821916, 25.408916287080)
    pts(7) = Array(-10.707653591720, 17.951174568557, 25.495697552615)
    pts(8) = Array(-10.450000000000, 18.099930939095, 25.525000000000)
    pts(9) = Array(-10.192346408280, 18.248687309632, 25.495697552615)
    pts(10) = Array(-9.944594299596, 18.391727056273, 25.408916287080)
    pts(11) = Array(-9.716264648543, 18.523553241772, 25.267991158761)
    pts(12) = Array(-9.516132035564, 18.639099859750, 25.078337841309)
    pts(13) = Array(-9.351887444739, 18.733926518475, 24.847244605355)
    pts(14) = Array(-9.229842703583, 18.804389082635, 24.583592234357)
    pts(15) = Array(-9.154687923057, 18.847779715402, 24.297512741075)
    pts(16) = Array(-9.129311259229, 18.862430939095, 24.000000000000)
    pts(17) = Array(-9.154687923057, 18.847779715402, 23.702487258925)
    pts(18) = Array(-9.229842703583, 18.804389082635, 23.416407765643)
    pts(19) = Array(-9.351887444739, 18.733926518475, 23.152755394645)
    pts(20) = Array(-9.516132035564, 18.639099859750, 22.921662158691)
    pts(21) = Array(-9.716264648543, 18.523553241772, 22.732008841239)
    pts(22) = Array(-9.944594299596, 18.391727056273, 22.591083712920)
    pts(23) = Array(-10.192346408280, 18.248687309632, 22.504302447385)
    pts(24) = Array(-10.450000000000, 18.099930939095, 22.475000000000)
    pts(25) = Array(-10.707653591720, 17.951174568557, 22.504302447385)
    pts(26) = Array(-10.955405700404, 17.808134821916, 22.591083712920)
    pts(27) = Array(-11.183735351457, 17.676308636417, 22.732008841239)
    pts(28) = Array(-11.383867964436, 17.560762018440, 22.921662158691)
    pts(29) = Array(-11.548112555261, 17.465935359714, 23.152755394645)
    pts(30) = Array(-11.670157296417, 17.395472795555, 23.416407765643)
    pts(31) = Array(-11.745312076943, 17.352082162787, 23.702487258925)
    FSData_8 = pts
End Function

Private Function FSData_9() As Variant
    Dim pts(31) As Variant
    pts(0) = Array(-12.670688740771, 18.896276665907, 24.000000000000)
    pts(1) = Array(-12.645312076943, 18.910927889599, 24.297512741075)
    pts(2) = Array(-12.570157296417, 18.954318522367, 24.583592234357)
    pts(3) = Array(-12.448112555261, 19.024781086526, 24.847244605355)
    pts(4) = Array(-12.283867964436, 19.119607745252, 25.078337841309)
    pts(5) = Array(-12.083735351457, 19.235154363229, 25.267991158761)
    pts(6) = Array(-11.855405700404, 19.366980548728, 25.408916287080)
    pts(7) = Array(-11.607653591720, 19.510020295369, 25.495697552615)
    pts(8) = Array(-11.350000000000, 19.658776665907, 25.525000000000)
    pts(9) = Array(-11.092346408280, 19.807533036444, 25.495697552615)
    pts(10) = Array(-10.844594299596, 19.950572783085, 25.408916287080)
    pts(11) = Array(-10.616264648543, 20.082398968584, 25.267991158761)
    pts(12) = Array(-10.416132035564, 20.197945586561, 25.078337841309)
    pts(13) = Array(-10.251887444739, 20.292772245287, 24.847244605355)
    pts(14) = Array(-10.129842703583, 20.363234809447, 24.583592234357)
    pts(15) = Array(-10.054687923057, 20.406625442214, 24.297512741075)
    pts(16) = Array(-10.029311259229, 20.421276665907, 24.000000000000)
    pts(17) = Array(-10.054687923057, 20.406625442214, 23.702487258925)
    pts(18) = Array(-10.129842703583, 20.363234809447, 23.416407765643)
    pts(19) = Array(-10.251887444739, 20.292772245287, 23.152755394645)
    pts(20) = Array(-10.416132035564, 20.197945586561, 22.921662158691)
    pts(21) = Array(-10.616264648543, 20.082398968584, 22.732008841239)
    pts(22) = Array(-10.844594299596, 19.950572783085, 22.591083712920)
    pts(23) = Array(-11.092346408280, 19.807533036444, 22.504302447385)
    pts(24) = Array(-11.350000000000, 19.658776665907, 22.475000000000)
    pts(25) = Array(-11.607653591720, 19.510020295369, 22.504302447385)
    pts(26) = Array(-11.855405700404, 19.366980548728, 22.591083712920)
    pts(27) = Array(-12.083735351457, 19.235154363229, 22.732008841239)
    pts(28) = Array(-12.283867964436, 19.119607745252, 22.921662158691)
    pts(29) = Array(-12.448112555261, 19.024781086526, 23.152755394645)
    pts(30) = Array(-12.570157296417, 18.954318522367, 23.416407765643)
    pts(31) = Array(-12.645312076943, 18.910927889599, 23.702487258925)
    FSData_9 = pts
End Function

Private Function FSData_10() As Variant
    Dim pts(31) As Variant
    pts(0) = Array(-9.129311259229, -18.862430939095, 24.000000000000)
    pts(1) = Array(-9.154687923057, -18.847779715402, 24.297512741075)
    pts(2) = Array(-9.229842703583, -18.804389082635, 24.583592234357)
    pts(3) = Array(-9.351887444739, -18.733926518475, 24.847244605355)
    pts(4) = Array(-9.516132035564, -18.639099859750, 25.078337841309)
    pts(5) = Array(-9.716264648543, -18.523553241772, 25.267991158761)
    pts(6) = Array(-9.944594299596, -18.391727056273, 25.408916287080)
    pts(7) = Array(-10.192346408280, -18.248687309632, 25.495697552615)
    pts(8) = Array(-10.450000000000, -18.099930939095, 25.525000000000)
    pts(9) = Array(-10.707653591720, -17.951174568557, 25.495697552615)
    pts(10) = Array(-10.955405700404, -17.808134821916, 25.408916287080)
    pts(11) = Array(-11.183735351457, -17.676308636417, 25.267991158761)
    pts(12) = Array(-11.383867964436, -17.560762018440, 25.078337841309)
    pts(13) = Array(-11.548112555261, -17.465935359714, 24.847244605355)
    pts(14) = Array(-11.670157296417, -17.395472795555, 24.583592234357)
    pts(15) = Array(-11.745312076943, -17.352082162787, 24.297512741075)
    pts(16) = Array(-11.770688740771, -17.337430939095, 24.000000000000)
    pts(17) = Array(-11.745312076943, -17.352082162787, 23.702487258925)
    pts(18) = Array(-11.670157296417, -17.395472795555, 23.416407765643)
    pts(19) = Array(-11.548112555261, -17.465935359714, 23.152755394645)
    pts(20) = Array(-11.383867964436, -17.560762018440, 22.921662158691)
    pts(21) = Array(-11.183735351457, -17.676308636417, 22.732008841239)
    pts(22) = Array(-10.955405700404, -17.808134821916, 22.591083712920)
    pts(23) = Array(-10.707653591720, -17.951174568557, 22.504302447385)
    pts(24) = Array(-10.450000000000, -18.099930939095, 22.475000000000)
    pts(25) = Array(-10.192346408280, -18.248687309632, 22.504302447385)
    pts(26) = Array(-9.944594299596, -18.391727056273, 22.591083712920)
    pts(27) = Array(-9.716264648543, -18.523553241772, 22.732008841239)
    pts(28) = Array(-9.516132035564, -18.639099859750, 22.921662158691)
    pts(29) = Array(-9.351887444739, -18.733926518475, 23.152755394645)
    pts(30) = Array(-9.229842703583, -18.804389082635, 23.416407765643)
    pts(31) = Array(-9.154687923057, -18.847779715402, 23.702487258925)
    FSData_10 = pts
End Function

Private Function FSData_11() As Variant
    Dim pts(31) As Variant
    pts(0) = Array(-10.029311259229, -20.421276665907, 24.000000000000)
    pts(1) = Array(-10.054687923057, -20.406625442214, 24.297512741075)
    pts(2) = Array(-10.129842703583, -20.363234809447, 24.583592234357)
    pts(3) = Array(-10.251887444739, -20.292772245287, 24.847244605355)
    pts(4) = Array(-10.416132035564, -20.197945586561, 25.078337841309)
    pts(5) = Array(-10.616264648543, -20.082398968584, 25.267991158761)
    pts(6) = Array(-10.844594299596, -19.950572783085, 25.408916287080)
    pts(7) = Array(-11.092346408280, -19.807533036444, 25.495697552615)
    pts(8) = Array(-11.350000000000, -19.658776665907, 25.525000000000)
    pts(9) = Array(-11.607653591720, -19.510020295369, 25.495697552615)
    pts(10) = Array(-11.855405700404, -19.366980548728, 25.408916287080)
    pts(11) = Array(-12.083735351457, -19.235154363229, 25.267991158761)
    pts(12) = Array(-12.283867964436, -19.119607745252, 25.078337841309)
    pts(13) = Array(-12.448112555261, -19.024781086526, 24.847244605355)
    pts(14) = Array(-12.570157296417, -18.954318522367, 24.583592234357)
    pts(15) = Array(-12.645312076943, -18.910927889599, 24.297512741075)
    pts(16) = Array(-12.670688740771, -18.896276665907, 24.000000000000)
    pts(17) = Array(-12.645312076943, -18.910927889599, 23.702487258925)
    pts(18) = Array(-12.570157296417, -18.954318522367, 23.416407765643)
    pts(19) = Array(-12.448112555261, -19.024781086526, 23.152755394645)
    pts(20) = Array(-12.283867964436, -19.119607745252, 22.921662158691)
    pts(21) = Array(-12.083735351457, -19.235154363229, 22.732008841239)
    pts(22) = Array(-11.855405700404, -19.366980548728, 22.591083712920)
    pts(23) = Array(-11.607653591720, -19.510020295369, 22.504302447385)
    pts(24) = Array(-11.350000000000, -19.658776665907, 22.475000000000)
    pts(25) = Array(-11.092346408280, -19.807533036444, 22.504302447385)
    pts(26) = Array(-10.844594299596, -19.950572783085, 22.591083712920)
    pts(27) = Array(-10.616264648543, -20.082398968584, 22.732008841239)
    pts(28) = Array(-10.416132035564, -20.197945586561, 22.921662158691)
    pts(29) = Array(-10.251887444739, -20.292772245287, 23.152755394645)
    pts(30) = Array(-10.129842703583, -20.363234809447, 23.416407765643)
    pts(31) = Array(-10.054687923057, -20.406625442214, 23.702487258925)
    FSData_11 = pts
End Function

Private Function FSData_12() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(18.512645572732, 19.170405679025, 19.300000000000)
    pts(1) = Array(21.708074076844, 22.479368760583, 19.300000000000)
    pts(2) = Array(21.708074076844, 22.479368760583, 24.700000000000)
    pts(3) = Array(18.512645572732, 19.170405679025, 24.700000000000)
    FSData_12 = pts
End Function

Private Function FSData_13() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(18.844395718621, 18.844395718621, 19.300000000000)
    pts(1) = Array(22.097086912080, 22.097086912080, 19.300000000000)
    pts(2) = Array(22.097086912080, 22.097086912080, 24.700000000000)
    pts(3) = Array(18.844395718621, 18.844395718621, 24.700000000000)
    FSData_13 = pts
End Function

Private Function FSData_14() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.170405679025, 18.512645572732, 19.300000000000)
    pts(1) = Array(22.479368760583, 21.708074076844, 19.300000000000)
    pts(2) = Array(22.479368760583, 21.708074076844, 24.700000000000)
    pts(3) = Array(19.170405679025, 18.512645572732, 24.700000000000)
    FSData_14 = pts
End Function

Private Function FSData_15() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.490576148151, 18.175256295666, 19.300000000000)
    pts(1) = Array(22.854803175599, 21.312448751953, 19.300000000000)
    pts(2) = Array(22.854803175599, 21.312448751953, 24.700000000000)
    pts(3) = Array(19.490576148151, 18.175256295666, 24.700000000000)
    FSData_15 = pts
End Function

Private Function FSData_16() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.804809598973, 17.832330659464, 19.300000000000)
    pts(1) = Array(23.223275796169, 20.910331448714, 19.300000000000)
    pts(2) = Array(23.223275796169, 20.910331448714, 24.700000000000)
    pts(3) = Array(19.804809598973, 17.832330659464, 24.700000000000)
    FSData_16 = pts
End Function

Private Function FSData_17() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.113010312937, 17.483973122597, 19.300000000000)
    pts(1) = Array(23.584674381962, 20.501844655953, 19.300000000000)
    pts(2) = Array(23.584674381962, 20.501844655953, 24.700000000000)
    pts(3) = Array(20.113010312937, 17.483973122597, 24.700000000000)
    FSData_17 = pts
End Function

Private Function FSData_18() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.415084409121, 17.130289798146, 19.300000000000)
    pts(1) = Array(23.938888847468, 20.087112802704, 19.300000000000)
    pts(2) = Array(23.938888847468, 20.087112802704, 24.700000000000)
    pts(3) = Array(20.415084409121, 17.130289798146, 24.700000000000)
    FSData_18 = pts
End Function

Private Function FSData_19() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.710939872828, 16.771388421478, 19.301859375000)
    pts(1) = Array(24.285811295530, 19.666262220307, 19.301859375000)
    pts(2) = Array(24.285811295530, 19.666262220307, 24.701859375000)
    pts(3) = Array(20.710939872828, 16.771388421478, 24.701859375000)
    FSData_19 = pts
End Function

Private Function FSData_20() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(21.000486583619, 16.407378317429, 19.307375000000)
    pts(1) = Array(24.625336050210, 19.239421103927, 19.307375000000)
    pts(2) = Array(24.625336050210, 19.239421103927, 24.707375000000)
    pts(3) = Array(21.000486583619, 16.407378317429, 24.707375000000)
    FSData_20 = pts
End Function

Private Function FSData_21() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(21.283636342760, 16.038370367002, 19.316453125000)
    pts(1) = Array(24.957359688978, 18.806719473502, 19.316453125000)
    pts(2) = Array(24.957359688978, 18.806719473502, 24.716453125000)
    pts(3) = Array(21.283636342760, 16.038370367002, 24.716453125000)
    FSData_21 = pts
End Function

Private Function FSData_22() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(21.560302900092, 15.664476973594, 19.329000000000)
    pts(1) = Array(25.281781074217, 18.368289134140, 19.329000000000)
    pts(2) = Array(25.281781074217, 18.368289134140, 24.729000000000)
    pts(3) = Array(21.560302900092, 15.664476973594, 24.729000000000)
    FSData_22 = pts
End Function

Private Function FSData_23() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(21.830401980302, 15.285812028755, 19.344921875000)
    pts(1) = Array(25.598501384031, 17.924263635970, 19.344921875000)
    pts(2) = Array(25.598501384031, 17.924263635970, 24.744921875000)
    pts(3) = Array(21.830401980302, 15.285812028755, 24.744921875000)
    FSData_23 = pts
End Function

Private Function FSData_24() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.093851308592, 14.902490877495, 19.364125000000)
    pts(1) = Array(25.907424142345, 17.474778233461, 19.364125000000)
    pts(2) = Array(25.907424142345, 17.474778233461, 24.764125000000)
    pts(3) = Array(22.093851308592, 14.902490877495, 24.764125000000)
    FSData_24 = pts
End Function

Private Function FSData_25() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.350570635746, 14.514630283150, 19.386515625000)
    pts(1) = Array(26.208455248295, 17.019969844220, 19.386515625000)
    pts(2) = Array(26.208455248295, 17.019969844220, 24.786515625000)
    pts(3) = Array(22.350570635746, 14.514630283150, 24.786515625000)
    FSData_25 = pts
End Function

Private Function FSData_26() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.600481762569, 14.122348391815, 19.412000000000)
    pts(1) = Array(26.501503004888, 16.559977007288, 19.412000000000)
    pts(2) = Array(26.501503004888, 16.559977007288, 24.812000000000)
    pts(3) = Array(22.600481762569, 14.122348391815, 24.812000000000)
    FSData_26 = pts
End Function

Private Function FSData_27() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.843508563711, 13.725764696353, 19.440484375000)
    pts(1) = Array(26.786478146941, 16.094939840939, 19.440484375000)
    pts(2) = Array(26.786478146941, 16.094939840939, 24.840484375000)
    pts(3) = Array(22.843508563711, 13.725764696353, 24.840484375000)
    FSData_27 = pts
End Function

Private Function FSData_28() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.079577010855, 13.325000000000, 19.471875000000)
    pts(1) = Array(27.063293868264, 15.625000000000, 19.471875000000)
    pts(2) = Array(27.063293868264, 15.625000000000, 24.871875000000)
    pts(3) = Array(23.079577010855, 13.325000000000, 24.871875000000)
    FSData_28 = pts
End Function

Private Function FSData_29() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.308615195265, 12.920176379565, 19.506078125000)
    pts(1) = Array(27.331865848106, 15.150300632698, 19.506078125000)
    pts(2) = Array(27.331865848106, 15.150300632698, 24.906078125000)
    pts(3) = Array(23.308615195265, 12.920176379565, 24.906078125000)
    FSData_29 = pts
End Function

Private Function FSData_30() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.530553349690, 12.511417148244, 19.543000000000)
    pts(1) = Array(27.592112276841, 14.670986337059, 19.543000000000)
    pts(2) = Array(27.592112276841, 14.670986337059, 24.943000000000)
    pts(3) = Array(23.530553349690, 12.511417148244, 24.943000000000)
    FSData_30 = pts
End Function

Private Function FSData_31() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.745323869620, 12.098846818059, 19.582546875000)
    pts(1) = Array(27.843953880886, 14.187203116861, 19.582546875000)
    pts(2) = Array(27.843953880886, 14.187203116861, 24.982546875000)
    pts(3) = Array(23.745323869620, 12.098846818059, 24.982546875000)
    FSData_31 = pts
End Function

Private Function FSData_32() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.952861333873, 11.682591061929, 19.624625000000)
    pts(1) = Array(28.087313946849, 13.699098337159, 19.624625000000)
    pts(2) = Array(28.087313946849, 13.699098337159, 25.024625000000)
    pts(3) = Array(23.952861333873, 11.682591061929, 25.024625000000)
    FSData_32 = pts
End Function

Private Function FSData_33() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.153102524527, 11.262776675390, 19.669140625000)
    pts(1) = Array(28.322118344895, 13.206820679397, 19.669140625000)
    pts(2) = Array(28.322118344895, 13.206820679397, 25.069140625000)
    pts(3) = Array(24.153102524527, 11.262776675390, 25.069140625000)
    FSData_33 = pts
End Function

Private Function FSData_34() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.345986446175, 10.839531537970, 19.716000000000)
    pts(1) = Array(28.548295551331, 12.710520096119, 19.716000000000)
    pts(2) = Array(28.548295551331, 12.710520096119, 25.116000000000)
    pts(3) = Array(24.345986446175, 10.839531537970, 25.116000000000)
    FSData_34 = pts
End Function

Private Function FSData_35() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.531454344508, 10.412984574239, 19.765109375000)
    pts(1) = Array(28.765776670389, 12.210347765290, 19.765109375000)
    pts(2) = Array(28.765776670389, 12.210347765290, 25.165109375000)
    pts(3) = Array(24.531454344508, 10.412984574239, 25.165109375000)
    FSData_35 = pts
End Function

Private Function FSData_36() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.709449724205, 9.983265714534, 19.816375000000)
    pts(1) = Array(28.974495455212, 11.706456044247, 19.816375000000)
    pts(2) = Array(28.974495455212, 11.706456044247, 25.216375000000)
    pts(3) = Array(24.709449724205, 9.983265714534, 25.216375000000)
    FSData_36 = pts
End Function

Private Function FSData_37() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.879918366150, 9.550505855382, 19.869703125000)
    pts(1) = Array(29.174388328038, 11.198998423291, 19.869703125000)
    pts(2) = Array(29.174388328038, 11.198998423291, 25.269703125000)
    pts(3) = Array(24.879918366150, 9.550505855382, 25.269703125000)
    FSData_37 = pts
End Function

Private Function FSData_38() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.042808343944, 9.114836819629, 19.925000000000)
    pts(1) = Array(29.365394399560, 10.688129478927, 19.925000000000)
    pts(2) = Array(29.365394399560, 10.688129478927, 25.325000000000)
    pts(3) = Array(25.042808343944, 9.114836819629, 25.325000000000)
    FSData_38 = pts
End Function

Private Function FSData_39() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.198070039722, 8.676391316283, 19.982171875000)
    pts(1) = Array(29.547455487479, 10.174004826786, 19.982171875000)
    pts(2) = Array(29.547455487479, 10.174004826786, 25.382171875000)
    pts(3) = Array(25.198070039722, 8.676391316283, 25.382171875000)
    FSData_39 = pts
End Function

Private Function FSData_40() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.345656159266, 8.235302900092, 20.041125000000)
    pts(1) = Array(29.720516134224, 9.656781074217, 20.041125000000)
    pts(2) = Array(29.720516134224, 9.656781074217, 25.441125000000)
    pts(3) = Array(25.345656159266, 8.235302900092, 25.441125000000)
    FSData_40 = pts
End Function

Private Function FSData_41() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.485521746415, 7.791705930861, 20.101765625000)
    pts(1) = Array(29.884523623845, 9.136615772586, 20.101765625000)
    pts(2) = Array(29.884523623845, 9.136615772586, 25.501765625000)
    pts(3) = Array(25.485521746415, 7.791705930861, 25.501765625000)
    FSData_41 = pts
End Function

Private Function FSData_42() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.617624196756, 7.345735532523, 20.164000000000)
    pts(1) = Array(30.039427998072, 8.613667369281, 20.164000000000)
    pts(2) = Array(30.039427998072, 8.613667369281, 25.564000000000)
    pts(3) = Array(25.617624196756, 7.345735532523, 25.564000000000)
    FSData_42 = pts
End Function

Private Function FSData_43() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.741923270604, 6.897527551982, 20.227734375000)
    pts(1) = Array(30.185182071533, 8.088095159454, 20.227734375000)
    pts(2) = Array(30.185182071533, 8.088095159454, 25.627734375000)
    pts(3) = Array(25.741923270604, 6.897527551982, 25.627734375000)
    FSData_43 = pts
End Function

Private Function FSData_44() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.858381105255, 6.447218517731, 20.292875000000)
    pts(1) = Array(30.321741446125, 7.560059237490, 20.292875000000)
    pts(2) = Array(30.321741446125, 7.560059237490, 25.692875000000)
    pts(3) = Array(25.858381105255, 6.447218517731, 25.692875000000)
    FSData_44 = pts
End Function

Private Function FSData_45() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.966962226527, 5.994945598264, 20.359328125000)
    pts(1) = Array(30.449064524539, 7.029720448246, 20.359328125000)
    pts(2) = Array(30.449064524539, 7.029720448246, 25.759328125000)
    pts(3) = Array(25.966962226527, 5.994945598264, 25.759328125000)
    FSData_45 = pts
End Function

Private Function FSData_46() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.067633559556, 5.540846560293, 20.427000000000)
    pts(1) = Array(30.567112522931, 6.497240338055, 20.427000000000)
    pts(2) = Array(30.567112522931, 6.497240338055, 25.827000000000)
    pts(3) = Array(26.067633559556, 5.540846560293, 25.827000000000)
    FSData_46 = pts
End Function

Private Function FSData_47() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.160364438880, 5.085059726785, 20.495796875000)
    pts(1) = Array(30.675849482739, 5.962781105517, 20.495796875000)
    pts(2) = Array(30.675849482739, 5.962781105517, 25.895796875000)
    pts(3) = Array(26.160364438880, 5.085059726785, 25.895796875000)
    FSData_47 = pts
End Function

Private Function FSData_48() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.245126617775, 4.627723934824, 20.565625000000)
    pts(1) = Array(30.775242281631, 5.426505552092, 20.565625000000)
    pts(2) = Array(30.775242281631, 5.426505552092, 25.965625000000)
    pts(3) = Array(26.245126617775, 4.627723934824, 25.965625000000)
    FSData_48 = pts
End Function

Private Function FSData_49() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.321894276860, 4.168978493322, 20.636390625000)
    pts(1) = Array(30.865260643598, 4.888577032507, 20.636390625000)
    pts(2) = Array(30.865260643598, 4.888577032507, 26.036390625000)
    pts(3) = Array(26.321894276860, 4.168978493322, 26.036390625000)
    FSData_49 = pts
End Function

Private Function FSData_50() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.390644031963, 3.708963140586, 20.708000000000)
    pts(1) = Array(30.945877148174, 4.349159405002, 20.708000000000)
    pts(2) = Array(30.945877148174, 4.349159405002, 26.108000000000)
    pts(3) = Array(26.390644031963, 3.708963140586, 26.108000000000)
    FSData_50 = pts
End Function

Private Function FSData_51() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.451354941241, 3.247818001747, 20.780359375000)
    pts(1) = Array(31.017067238791, 3.808416981411, 20.780359375000)
    pts(2) = Array(31.017067238791, 3.808416981411, 26.180359375000)
    pts(3) = Array(26.451354941241, 3.247818001747, 26.180359375000)
    FSData_51 = pts
End Function

Private Function FSData_52() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.504008511564, 2.785683546083, 20.853375000000)
    pts(1) = Array(31.078809230259, 3.266514477114, 20.853375000000)
    pts(2) = Array(31.078809230259, 3.266514477114, 26.253375000000)
    pts(3) = Array(26.504008511564, 2.785683546083, 26.253375000000)
    FSData_52 = pts
End Function

Private Function FSData_53() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.548588704145, 2.322700544225, 20.926953125000)
    pts(1) = Array(31.131084315367, 2.723616960864, 20.926953125000)
    pts(2) = Array(31.131084315367, 2.723616960864, 26.326953125000)
    pts(3) = Array(26.548588704145, 2.322700544225, 26.326953125000)
    FSData_53 = pts
End Function

Private Function FSData_54() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.585081939424, 1.859010025281, 21.001000000000)
    pts(1) = Array(31.173876570620, 2.179889804504, 21.001000000000)
    pts(2) = Array(31.173876570620, 2.179889804504, 26.401000000000)
    pts(3) = Array(26.585081939424, 1.859010025281, 26.401000000000)
    FSData_54 = pts
End Function

Private Function FSData_55() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.613477101209, 1.394753233874, 21.075421875000)
    pts(1) = Array(31.207172961080, 1.635498632592, 21.075421875000)
    pts(2) = Array(31.207172961080, 1.635498632592, 26.475421875000)
    pts(3) = Array(26.613477101209, 1.394753233874, 26.475421875000)
    FSData_55 = pts
End Function

Private Function FSData_56() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.633765540059, 0.930071587122, 21.150125000000)
    pts(1) = Array(31.230963344347, 1.090609271953, 21.150125000000)
    pts(2) = Array(31.230963344347, 1.090609271953, 26.550125000000)
    pts(3) = Array(26.633765540059, 0.930071587122, 26.550125000000)
    FSData_56 = pts
End Function

Private Function FSData_57() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.645941075918, 0.465106631554, 21.225015625000)
    pts(1) = Array(31.245240473637, 0.545387701165, 21.225015625000)
    pts(2) = Array(31.245240473637, 0.545387701165, 26.625015625000)
    pts(3) = Array(26.645941075918, 0.465106631554, 26.625015625000)
    FSData_57 = pts
End Function

Private Function FSData_58() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.650000000000, 0.000000000000, 21.300000000000)
    pts(1) = Array(31.250000000000, 0.000000000000, 21.300000000000)
    pts(2) = Array(31.250000000000, 0.000000000000, 26.700000000000)
    pts(3) = Array(26.650000000000, 0.000000000000, 26.700000000000)
    FSData_58 = pts
End Function

Private Function FSData_59() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.645941075918, -0.465106631554, 21.374984375000)
    pts(1) = Array(31.245240473637, -0.545387701165, 21.374984375000)
    pts(2) = Array(31.245240473637, -0.545387701165, 26.774984375000)
    pts(3) = Array(26.645941075918, -0.465106631554, 26.774984375000)
    FSData_59 = pts
End Function

Private Function FSData_60() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.633765540059, -0.930071587122, 21.449875000000)
    pts(1) = Array(31.230963344347, -1.090609271953, 21.449875000000)
    pts(2) = Array(31.230963344347, -1.090609271953, 26.849875000000)
    pts(3) = Array(26.633765540059, -0.930071587122, 26.849875000000)
    FSData_60 = pts
End Function

Private Function FSData_61() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.613477101209, -1.394753233874, 21.524578125000)
    pts(1) = Array(31.207172961080, -1.635498632592, 21.524578125000)
    pts(2) = Array(31.207172961080, -1.635498632592, 26.924578125000)
    pts(3) = Array(26.613477101209, -1.394753233874, 26.924578125000)
    FSData_61 = pts
End Function

Private Function FSData_62() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.585081939424, -1.859010025281, 21.599000000000)
    pts(1) = Array(31.173876570620, -2.179889804504, 21.599000000000)
    pts(2) = Array(31.173876570620, -2.179889804504, 26.999000000000)
    pts(3) = Array(26.585081939424, -1.859010025281, 26.999000000000)
    FSData_62 = pts
End Function

Private Function FSData_63() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.548588704145, -2.322700544225, 21.673046875000)
    pts(1) = Array(31.131084315367, -2.723616960864, 21.673046875000)
    pts(2) = Array(31.131084315367, -2.723616960864, 27.073046875000)
    pts(3) = Array(26.548588704145, -2.322700544225, 27.073046875000)
    FSData_63 = pts
End Function

Private Function FSData_64() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.504008511564, -2.785683546083, 21.746625000000)
    pts(1) = Array(31.078809230259, -3.266514477114, 21.746625000000)
    pts(2) = Array(31.078809230259, -3.266514477114, 27.146625000000)
    pts(3) = Array(26.504008511564, -2.785683546083, 27.146625000000)
    FSData_64 = pts
End Function

Private Function FSData_65() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.451354941241, -3.247818001747, 21.819640625000)
    pts(1) = Array(31.017067238791, -3.808416981411, 21.819640625000)
    pts(2) = Array(31.017067238791, -3.808416981411, 27.219640625000)
    pts(3) = Array(26.451354941241, -3.247818001747, 27.219640625000)
    FSData_65 = pts
End Function

Private Function FSData_66() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.390644031963, -3.708963140586, 21.892000000000)
    pts(1) = Array(30.945877148174, -4.349159405002, 21.892000000000)
    pts(2) = Array(30.945877148174, -4.349159405002, 27.292000000000)
    pts(3) = Array(26.390644031963, -3.708963140586, 27.292000000000)
    FSData_66 = pts
End Function

Private Function FSData_67() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.321894276860, -4.168978493322, 21.963609375000)
    pts(1) = Array(30.865260643598, -4.888577032507, 21.963609375000)
    pts(2) = Array(30.865260643598, -4.888577032507, 27.363609375000)
    pts(3) = Array(26.321894276860, -4.168978493322, 27.363609375000)
    FSData_67 = pts
End Function

Private Function FSData_68() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.245126617775, -4.627723934824, 22.034375000000)
    pts(1) = Array(30.775242281631, -5.426505552092, 22.034375000000)
    pts(2) = Array(30.775242281631, -5.426505552092, 27.434375000000)
    pts(3) = Array(26.245126617775, -4.627723934824, 27.434375000000)
    FSData_68 = pts
End Function

Private Function FSData_69() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.160364438880, -5.085059726785, 22.104203125000)
    pts(1) = Array(30.675849482739, -5.962781105517, 22.104203125000)
    pts(2) = Array(30.675849482739, -5.962781105517, 27.504203125000)
    pts(3) = Array(26.160364438880, -5.085059726785, 27.504203125000)
    FSData_69 = pts
End Function

Private Function FSData_70() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(26.067633559556, -5.540846560293, 22.173000000000)
    pts(1) = Array(30.567112522931, -6.497240338055, 22.173000000000)
    pts(2) = Array(30.567112522931, -6.497240338055, 27.573000000000)
    pts(3) = Array(26.067633559556, -5.540846560293, 27.573000000000)
    FSData_70 = pts
End Function

Private Function FSData_71() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.966962226527, -5.994945598264, 22.240671875000)
    pts(1) = Array(30.449064524539, -7.029720448246, 22.240671875000)
    pts(2) = Array(30.449064524539, -7.029720448246, 27.640671875000)
    pts(3) = Array(25.966962226527, -5.994945598264, 27.640671875000)
    FSData_71 = pts
End Function

Private Function FSData_72() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.858381105255, -6.447218517731, 22.307125000000)
    pts(1) = Array(30.321741446125, -7.560059237490, 22.307125000000)
    pts(2) = Array(30.321741446125, -7.560059237490, 27.707125000000)
    pts(3) = Array(25.858381105255, -6.447218517731, 27.707125000000)
    FSData_72 = pts
End Function

Private Function FSData_73() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.741923270604, -6.897527551982, 22.372265625000)
    pts(1) = Array(30.185182071533, -8.088095159454, 22.372265625000)
    pts(2) = Array(30.185182071533, -8.088095159454, 27.772265625000)
    pts(3) = Array(25.741923270604, -6.897527551982, 27.772265625000)
    FSData_73 = pts
End Function

Private Function FSData_74() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.617624196756, -7.345735532523, 22.436000000000)
    pts(1) = Array(30.039427998072, -8.613667369281, 22.436000000000)
    pts(2) = Array(30.039427998072, -8.613667369281, 27.836000000000)
    pts(3) = Array(25.617624196756, -7.345735532523, 27.836000000000)
    FSData_74 = pts
End Function

Private Function FSData_75() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.485521746415, -7.791705930861, 22.498234375000)
    pts(1) = Array(29.884523623845, -9.136615772586, 22.498234375000)
    pts(2) = Array(29.884523623845, -9.136615772586, 27.898234375000)
    pts(3) = Array(25.485521746415, -7.791705930861, 27.898234375000)
    FSData_75 = pts
End Function

Private Function FSData_76() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.345656159266, -8.235302900092, 22.558875000000)
    pts(1) = Array(29.720516134224, -9.656781074217, 22.558875000000)
    pts(2) = Array(29.720516134224, -9.656781074217, 27.958875000000)
    pts(3) = Array(25.345656159266, -8.235302900092, 27.958875000000)
    FSData_76 = pts
End Function

Private Function FSData_77() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.198070039722, -8.676391316283, 22.617828125000)
    pts(1) = Array(29.547455487479, -10.174004826786, 22.617828125000)
    pts(2) = Array(29.547455487479, -10.174004826786, 28.017828125000)
    pts(3) = Array(25.198070039722, -8.676391316283, 28.017828125000)
    FSData_77 = pts
End Function

Private Function FSData_78() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(25.042808343944, -9.114836819629, 22.675000000000)
    pts(1) = Array(29.365394399560, -10.688129478927, 22.675000000000)
    pts(2) = Array(29.365394399560, -10.688129478927, 28.075000000000)
    pts(3) = Array(25.042808343944, -9.114836819629, 28.075000000000)
    FSData_78 = pts
End Function

Private Function FSData_79() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.879918366150, -9.550505855382, 22.730296875000)
    pts(1) = Array(29.174388328038, -11.198998423291, 22.730296875000)
    pts(2) = Array(29.174388328038, -11.198998423291, 28.130296875000)
    pts(3) = Array(24.879918366150, -9.550505855382, 28.130296875000)
    FSData_79 = pts
End Function

Private Function FSData_80() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.709449724205, -9.983265714534, 22.783625000000)
    pts(1) = Array(28.974495455212, -11.706456044247, 22.783625000000)
    pts(2) = Array(28.974495455212, -11.706456044247, 28.183625000000)
    pts(3) = Array(24.709449724205, -9.983265714534, 28.183625000000)
    FSData_80 = pts
End Function

Private Function FSData_81() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.531454344508, -10.412984574239, 22.834890625000)
    pts(1) = Array(28.765776670389, -12.210347765290, 22.834890625000)
    pts(2) = Array(28.765776670389, -12.210347765290, 28.234890625000)
    pts(3) = Array(24.531454344508, -10.412984574239, 28.234890625000)
    FSData_81 = pts
End Function

Private Function FSData_82() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.345986446175, -10.839531537970, 22.884000000000)
    pts(1) = Array(28.548295551331, -12.710520096119, 22.884000000000)
    pts(2) = Array(28.548295551331, -12.710520096119, 28.284000000000)
    pts(3) = Array(24.345986446175, -10.839531537970, 28.284000000000)
    FSData_82 = pts
End Function

Private Function FSData_83() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.153102524527, -11.262776675390, 22.930859375000)
    pts(1) = Array(28.322118344895, -13.206820679397, 22.930859375000)
    pts(2) = Array(28.322118344895, -13.206820679397, 28.330859375000)
    pts(3) = Array(24.153102524527, -11.262776675390, 28.330859375000)
    FSData_83 = pts
End Function

Private Function FSData_84() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.952861333873, -11.682591061929, 22.975375000000)
    pts(1) = Array(28.087313946849, -13.699098337159, 22.975375000000)
    pts(2) = Array(28.087313946849, -13.699098337159, 28.375375000000)
    pts(3) = Array(23.952861333873, -11.682591061929, 28.375375000000)
    FSData_84 = pts
End Function

Private Function FSData_85() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.745323869620, -12.098846818059, 23.017453125000)
    pts(1) = Array(27.843953880886, -14.187203116861, 23.017453125000)
    pts(2) = Array(27.843953880886, -14.187203116861, 28.417453125000)
    pts(3) = Array(23.745323869620, -12.098846818059, 28.417453125000)
    FSData_85 = pts
End Function

Private Function FSData_86() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.530553349690, -12.511417148244, 23.057000000000)
    pts(1) = Array(27.592112276841, -14.670986337059, 23.057000000000)
    pts(2) = Array(27.592112276841, -14.670986337059, 28.457000000000)
    pts(3) = Array(23.530553349690, -12.511417148244, 28.457000000000)
    FSData_86 = pts
End Function

Private Function FSData_87() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.308615195265, -12.920176379565, 23.093921875000)
    pts(1) = Array(27.331865848106, -15.150300632698, 23.093921875000)
    pts(2) = Array(27.331865848106, -15.150300632698, 28.493921875000)
    pts(3) = Array(23.308615195265, -12.920176379565, 28.493921875000)
    FSData_87 = pts
End Function

Private Function FSData_88() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(23.079577010855, -13.325000000000, 23.128125000000)
    pts(1) = Array(27.063293868264, -15.625000000000, 23.128125000000)
    pts(2) = Array(27.063293868264, -15.625000000000, 28.528125000000)
    pts(3) = Array(23.079577010855, -13.325000000000, 28.528125000000)
    FSData_88 = pts
End Function

Private Function FSData_89() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.843508563711, -13.725764696353, 23.159515625000)
    pts(1) = Array(26.786478146941, -16.094939840939, 23.159515625000)
    pts(2) = Array(26.786478146941, -16.094939840939, 28.559515625000)
    pts(3) = Array(22.843508563711, -13.725764696353, 28.559515625000)
    FSData_89 = pts
End Function

Private Function FSData_90() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.600481762569, -14.122348391815, 23.188000000000)
    pts(1) = Array(26.501503004888, -16.559977007288, 23.188000000000)
    pts(2) = Array(26.501503004888, -16.559977007288, 28.588000000000)
    pts(3) = Array(22.600481762569, -14.122348391815, 28.588000000000)
    FSData_90 = pts
End Function

Private Function FSData_91() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.350570635746, -14.514630283150, 23.213484375000)
    pts(1) = Array(26.208455248295, -17.019969844220, 23.213484375000)
    pts(2) = Array(26.208455248295, -17.019969844220, 28.613484375000)
    pts(3) = Array(22.350570635746, -14.514630283150, 28.613484375000)
    FSData_91 = pts
End Function

Private Function FSData_92() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(22.093851308592, -14.902490877495, 23.235875000000)
    pts(1) = Array(25.907424142345, -17.474778233461, 23.235875000000)
    pts(2) = Array(25.907424142345, -17.474778233461, 28.635875000000)
    pts(3) = Array(22.093851308592, -14.902490877495, 28.635875000000)
    FSData_92 = pts
End Function

Private Function FSData_93() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(21.830401980302, -15.285812028755, 23.255078125000)
    pts(1) = Array(25.598501384031, -17.924263635970, 23.255078125000)
    pts(2) = Array(25.598501384031, -17.924263635970, 28.655078125000)
    pts(3) = Array(21.830401980302, -15.285812028755, 28.655078125000)
    FSData_93 = pts
End Function

Private Function FSData_94() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(21.560302900092, -15.664476973594, 23.271000000000)
    pts(1) = Array(25.281781074217, -18.368289134140, 23.271000000000)
    pts(2) = Array(25.281781074217, -18.368289134140, 28.671000000000)
    pts(3) = Array(21.560302900092, -15.664476973594, 28.671000000000)
    FSData_94 = pts
End Function

Private Function FSData_95() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(21.283636342760, -16.038370367002, 23.283546875000)
    pts(1) = Array(24.957359688978, -18.806719473502, 23.283546875000)
    pts(2) = Array(24.957359688978, -18.806719473502, 28.683546875000)
    pts(3) = Array(21.283636342760, -16.038370367002, 28.683546875000)
    FSData_95 = pts
End Function

Private Function FSData_96() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(21.000486583619, -16.407378317429, 23.292625000000)
    pts(1) = Array(24.625336050210, -19.239421103927, 23.292625000000)
    pts(2) = Array(24.625336050210, -19.239421103927, 28.692625000000)
    pts(3) = Array(21.000486583619, -16.407378317429, 28.692625000000)
    FSData_96 = pts
End Function

Private Function FSData_97() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.710939872828, -16.771388421478, 23.298140625000)
    pts(1) = Array(24.285811295530, -19.666262220307, 23.298140625000)
    pts(2) = Array(24.285811295530, -19.666262220307, 28.698140625000)
    pts(3) = Array(20.710939872828, -16.771388421478, 28.698140625000)
    FSData_97 = pts
End Function

Private Function FSData_98() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.415084409121, -17.130289798146, 23.300000000000)
    pts(1) = Array(23.938888847468, -20.087112802704, 23.300000000000)
    pts(2) = Array(23.938888847468, -20.087112802704, 28.700000000000)
    pts(3) = Array(20.415084409121, -17.130289798146, 28.700000000000)
    FSData_98 = pts
End Function

Private Function FSData_99() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(20.113010312937, -17.483973122597, 23.300000000000)
    pts(1) = Array(23.584674381962, -20.501844655953, 23.300000000000)
    pts(2) = Array(23.584674381962, -20.501844655953, 28.700000000000)
    pts(3) = Array(20.113010312937, -17.483973122597, 28.700000000000)
    FSData_99 = pts
End Function

Private Function FSData_100() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.804809598973, -17.832330659464, 23.300000000000)
    pts(1) = Array(23.223275796169, -20.910331448714, 23.300000000000)
    pts(2) = Array(23.223275796169, -20.910331448714, 28.700000000000)
    pts(3) = Array(19.804809598973, -17.832330659464, 28.700000000000)
    FSData_100 = pts
End Function

Private Function FSData_101() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.490576148151, -18.175256295666, 23.300000000000)
    pts(1) = Array(22.854803175599, -21.312448751953, 23.300000000000)
    pts(2) = Array(22.854803175599, -21.312448751953, 28.700000000000)
    pts(3) = Array(19.490576148151, -18.175256295666, 28.700000000000)
    FSData_101 = pts
End Function

Private Function FSData_102() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(19.170405679025, -18.512645572732, 23.300000000000)
    pts(1) = Array(22.479368760583, -21.708074076844, 23.300000000000)
    pts(2) = Array(22.479368760583, -21.708074076844, 28.700000000000)
    pts(3) = Array(19.170405679025, -18.512645572732, 28.700000000000)
    FSData_102 = pts
End Function

Private Function FSData_103() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(18.844395718621, -18.844395718621, 23.300000000000)
    pts(1) = Array(22.097086912080, -22.097086912080, 23.300000000000)
    pts(2) = Array(22.097086912080, -22.097086912080, 28.700000000000)
    pts(3) = Array(18.844395718621, -18.844395718621, 28.700000000000)
    FSData_103 = pts
End Function

Private Function FSData_104() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(18.512645572732, -19.170405679025, 23.300000000000)
    pts(1) = Array(21.708074076844, -22.479368760583, 23.300000000000)
    pts(2) = Array(21.708074076844, -22.479368760583, 28.700000000000)
    pts(3) = Array(18.512645572732, -19.170405679025, 28.700000000000)
    FSData_104 = pts
End Function

Private Function FSData_105() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.858381105255, 6.447218517731, 19.300000000000)
    pts(1) = Array(-30.321741446125, 7.560059237490, 19.300000000000)
    pts(2) = Array(-30.321741446125, 7.560059237490, 24.700000000000)
    pts(3) = Array(-25.858381105255, 6.447218517731, 24.700000000000)
    FSData_105 = pts
End Function

Private Function FSData_106() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.741923270604, 6.897527551982, 19.300000000000)
    pts(1) = Array(-30.185182071533, 8.088095159454, 19.300000000000)
    pts(2) = Array(-30.185182071533, 8.088095159454, 24.700000000000)
    pts(3) = Array(-25.741923270604, 6.897527551982, 24.700000000000)
    FSData_106 = pts
End Function

Private Function FSData_107() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.617624196756, 7.345735532523, 19.300000000000)
    pts(1) = Array(-30.039427998072, 8.613667369281, 19.300000000000)
    pts(2) = Array(-30.039427998072, 8.613667369281, 24.700000000000)
    pts(3) = Array(-25.617624196756, 7.345735532523, 24.700000000000)
    FSData_107 = pts
End Function

Private Function FSData_108() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.485521746415, 7.791705930861, 19.300000000000)
    pts(1) = Array(-29.884523623845, 9.136615772586, 19.300000000000)
    pts(2) = Array(-29.884523623845, 9.136615772586, 24.700000000000)
    pts(3) = Array(-25.485521746415, 7.791705930861, 24.700000000000)
    FSData_108 = pts
End Function

Private Function FSData_109() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.345656159266, 8.235302900092, 19.300000000000)
    pts(1) = Array(-29.720516134224, 9.656781074217, 19.300000000000)
    pts(2) = Array(-29.720516134224, 9.656781074217, 24.700000000000)
    pts(3) = Array(-25.345656159266, 8.235302900092, 24.700000000000)
    FSData_109 = pts
End Function

Private Function FSData_110() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.198070039722, 8.676391316283, 19.300000000000)
    pts(1) = Array(-29.547455487479, 10.174004826786, 19.300000000000)
    pts(2) = Array(-29.547455487479, 10.174004826786, 24.700000000000)
    pts(3) = Array(-25.198070039722, 8.676391316283, 24.700000000000)
    FSData_110 = pts
End Function

Private Function FSData_111() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.042808343944, 9.114836819629, 19.300000000000)
    pts(1) = Array(-29.365394399560, 10.688129478927, 19.300000000000)
    pts(2) = Array(-29.365394399560, 10.688129478927, 24.700000000000)
    pts(3) = Array(-25.042808343944, 9.114836819629, 24.700000000000)
    FSData_111 = pts
End Function

Private Function FSData_112() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.879918366150, 9.550505855382, 19.301859375000)
    pts(1) = Array(-29.174388328038, 11.198998423291, 19.301859375000)
    pts(2) = Array(-29.174388328038, 11.198998423291, 24.701859375000)
    pts(3) = Array(-24.879918366150, 9.550505855382, 24.701859375000)
    FSData_112 = pts
End Function

Private Function FSData_113() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.709449724205, 9.983265714534, 19.307375000000)
    pts(1) = Array(-28.974495455212, 11.706456044247, 19.307375000000)
    pts(2) = Array(-28.974495455212, 11.706456044247, 24.707375000000)
    pts(3) = Array(-24.709449724205, 9.983265714534, 24.707375000000)
    FSData_113 = pts
End Function

Private Function FSData_114() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.531454344508, 10.412984574239, 19.316453125000)
    pts(1) = Array(-28.765776670389, 12.210347765290, 19.316453125000)
    pts(2) = Array(-28.765776670389, 12.210347765290, 24.716453125000)
    pts(3) = Array(-24.531454344508, 10.412984574239, 24.716453125000)
    FSData_114 = pts
End Function

Private Function FSData_115() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.345986446175, 10.839531537970, 19.329000000000)
    pts(1) = Array(-28.548295551331, 12.710520096119, 19.329000000000)
    pts(2) = Array(-28.548295551331, 12.710520096119, 24.729000000000)
    pts(3) = Array(-24.345986446175, 10.839531537970, 24.729000000000)
    FSData_115 = pts
End Function

Private Function FSData_116() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.153102524527, 11.262776675390, 19.344921875000)
    pts(1) = Array(-28.322118344895, 13.206820679397, 19.344921875000)
    pts(2) = Array(-28.322118344895, 13.206820679397, 24.744921875000)
    pts(3) = Array(-24.153102524527, 11.262776675390, 24.744921875000)
    FSData_116 = pts
End Function

Private Function FSData_117() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.952861333873, 11.682591061929, 19.364125000000)
    pts(1) = Array(-28.087313946849, 13.699098337159, 19.364125000000)
    pts(2) = Array(-28.087313946849, 13.699098337159, 24.764125000000)
    pts(3) = Array(-23.952861333873, 11.682591061929, 24.764125000000)
    FSData_117 = pts
End Function

Private Function FSData_118() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.745323869620, 12.098846818059, 19.386515625000)
    pts(1) = Array(-27.843953880886, 14.187203116861, 19.386515625000)
    pts(2) = Array(-27.843953880886, 14.187203116861, 24.786515625000)
    pts(3) = Array(-23.745323869620, 12.098846818059, 24.786515625000)
    FSData_118 = pts
End Function

Private Function FSData_119() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.530553349690, 12.511417148244, 19.412000000000)
    pts(1) = Array(-27.592112276841, 14.670986337059, 19.412000000000)
    pts(2) = Array(-27.592112276841, 14.670986337059, 24.812000000000)
    pts(3) = Array(-23.530553349690, 12.511417148244, 24.812000000000)
    FSData_119 = pts
End Function

Private Function FSData_120() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.308615195265, 12.920176379565, 19.440484375000)
    pts(1) = Array(-27.331865848106, 15.150300632698, 19.440484375000)
    pts(2) = Array(-27.331865848106, 15.150300632698, 24.840484375000)
    pts(3) = Array(-23.308615195265, 12.920176379565, 24.840484375000)
    FSData_120 = pts
End Function

Private Function FSData_121() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.079577010855, 13.325000000000, 19.471875000000)
    pts(1) = Array(-27.063293868264, 15.625000000000, 19.471875000000)
    pts(2) = Array(-27.063293868264, 15.625000000000, 24.871875000000)
    pts(3) = Array(-23.079577010855, 13.325000000000, 24.871875000000)
    FSData_121 = pts
End Function

Private Function FSData_122() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.843508563711, 13.725764696353, 19.506078125000)
    pts(1) = Array(-26.786478146941, 16.094939840939, 19.506078125000)
    pts(2) = Array(-26.786478146941, 16.094939840939, 24.906078125000)
    pts(3) = Array(-22.843508563711, 13.725764696353, 24.906078125000)
    FSData_122 = pts
End Function

Private Function FSData_123() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.600481762569, 14.122348391815, 19.543000000000)
    pts(1) = Array(-26.501503004888, 16.559977007288, 19.543000000000)
    pts(2) = Array(-26.501503004888, 16.559977007288, 24.943000000000)
    pts(3) = Array(-22.600481762569, 14.122348391815, 24.943000000000)
    FSData_123 = pts
End Function

Private Function FSData_124() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.350570635746, 14.514630283150, 19.582546875000)
    pts(1) = Array(-26.208455248294, 17.019969844220, 19.582546875000)
    pts(2) = Array(-26.208455248294, 17.019969844220, 24.982546875000)
    pts(3) = Array(-22.350570635746, 14.514630283150, 24.982546875000)
    FSData_124 = pts
End Function

Private Function FSData_125() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.093851308592, 14.902490877495, 19.624625000000)
    pts(1) = Array(-25.907424142345, 17.474778233461, 19.624625000000)
    pts(2) = Array(-25.907424142345, 17.474778233461, 25.024625000000)
    pts(3) = Array(-22.093851308592, 14.902490877495, 25.024625000000)
    FSData_125 = pts
End Function

Private Function FSData_126() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-21.830401980302, 15.285812028755, 19.669140625000)
    pts(1) = Array(-25.598501384031, 17.924263635970, 19.669140625000)
    pts(2) = Array(-25.598501384031, 17.924263635970, 25.069140625000)
    pts(3) = Array(-21.830401980302, 15.285812028755, 25.069140625000)
    FSData_126 = pts
End Function

Private Function FSData_127() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-21.560302900092, 15.664476973594, 19.716000000000)
    pts(1) = Array(-25.281781074217, 18.368289134140, 19.716000000000)
    pts(2) = Array(-25.281781074217, 18.368289134140, 25.116000000000)
    pts(3) = Array(-21.560302900092, 15.664476973594, 25.116000000000)
    FSData_127 = pts
End Function

Private Function FSData_128() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-21.283636342760, 16.038370367002, 19.765109375000)
    pts(1) = Array(-24.957359688978, 18.806719473502, 19.765109375000)
    pts(2) = Array(-24.957359688978, 18.806719473502, 25.165109375000)
    pts(3) = Array(-21.283636342760, 16.038370367002, 25.165109375000)
    FSData_128 = pts
End Function

Private Function FSData_129() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-21.000486583619, 16.407378317429, 19.816375000000)
    pts(1) = Array(-24.625336050210, 19.239421103927, 19.816375000000)
    pts(2) = Array(-24.625336050210, 19.239421103927, 25.216375000000)
    pts(3) = Array(-21.000486583619, 16.407378317429, 25.216375000000)
    FSData_129 = pts
End Function

Private Function FSData_130() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.710939872828, 16.771388421478, 19.869703125000)
    pts(1) = Array(-24.285811295530, 19.666262220307, 19.869703125000)
    pts(2) = Array(-24.285811295530, 19.666262220307, 25.269703125000)
    pts(3) = Array(-20.710939872828, 16.771388421478, 25.269703125000)
    FSData_130 = pts
End Function

Private Function FSData_131() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.415084409121, 17.130289798146, 19.925000000000)
    pts(1) = Array(-23.938888847468, 20.087112802704, 19.925000000000)
    pts(2) = Array(-23.938888847468, 20.087112802704, 25.325000000000)
    pts(3) = Array(-20.415084409121, 17.130289798146, 25.325000000000)
    FSData_131 = pts
End Function

Private Function FSData_132() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.113010312937, 17.483973122597, 19.982171875000)
    pts(1) = Array(-23.584674381962, 20.501844655953, 19.982171875000)
    pts(2) = Array(-23.584674381962, 20.501844655953, 25.382171875000)
    pts(3) = Array(-20.113010312937, 17.483973122597, 25.382171875000)
    FSData_132 = pts
End Function

Private Function FSData_133() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.804809598973, 17.832330659464, 20.041125000000)
    pts(1) = Array(-23.223275796169, 20.910331448714, 20.041125000000)
    pts(2) = Array(-23.223275796169, 20.910331448714, 25.441125000000)
    pts(3) = Array(-19.804809598973, 17.832330659464, 25.441125000000)
    FSData_133 = pts
End Function

Private Function FSData_134() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.490576148151, 18.175256295666, 20.101765625000)
    pts(1) = Array(-22.854803175599, 21.312448751953, 20.101765625000)
    pts(2) = Array(-22.854803175599, 21.312448751953, 25.501765625000)
    pts(3) = Array(-19.490576148151, 18.175256295666, 25.501765625000)
    FSData_134 = pts
End Function

Private Function FSData_135() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.170405679025, 18.512645572732, 20.164000000000)
    pts(1) = Array(-22.479368760583, 21.708074076844, 20.164000000000)
    pts(2) = Array(-22.479368760583, 21.708074076844, 25.564000000000)
    pts(3) = Array(-19.170405679025, 18.512645572732, 25.564000000000)
    FSData_135 = pts
End Function

Private Function FSData_136() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-18.844395718621, 18.844395718621, 20.227734375000)
    pts(1) = Array(-22.097086912080, 22.097086912080, 20.227734375000)
    pts(2) = Array(-22.097086912080, 22.097086912080, 25.627734375000)
    pts(3) = Array(-18.844395718621, 18.844395718621, 25.627734375000)
    FSData_136 = pts
End Function

Private Function FSData_137() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-18.512645572732, 19.170405679025, 20.292875000000)
    pts(1) = Array(-21.708074076844, 22.479368760583, 20.292875000000)
    pts(2) = Array(-21.708074076844, 22.479368760583, 25.692875000000)
    pts(3) = Array(-18.512645572732, 19.170405679025, 25.692875000000)
    FSData_137 = pts
End Function

Private Function FSData_138() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-18.175256295666, 19.490576148151, 20.359328125000)
    pts(1) = Array(-21.312448751953, 22.854803175599, 20.359328125000)
    pts(2) = Array(-21.312448751953, 22.854803175599, 25.759328125000)
    pts(3) = Array(-18.175256295666, 19.490576148151, 25.759328125000)
    FSData_138 = pts
End Function

Private Function FSData_139() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.832330659464, 19.804809598973, 20.427000000000)
    pts(1) = Array(-20.910331448714, 23.223275796169, 20.427000000000)
    pts(2) = Array(-20.910331448714, 23.223275796169, 25.827000000000)
    pts(3) = Array(-17.832330659464, 19.804809598973, 25.827000000000)
    FSData_139 = pts
End Function

Private Function FSData_140() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.483973122597, 20.113010312937, 20.495796875000)
    pts(1) = Array(-20.501844655953, 23.584674381962, 20.495796875000)
    pts(2) = Array(-20.501844655953, 23.584674381962, 25.895796875000)
    pts(3) = Array(-17.483973122597, 20.113010312937, 25.895796875000)
    FSData_140 = pts
End Function

Private Function FSData_141() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.130289798146, 20.415084409121, 20.565625000000)
    pts(1) = Array(-20.087112802704, 23.938888847468, 20.565625000000)
    pts(2) = Array(-20.087112802704, 23.938888847468, 25.965625000000)
    pts(3) = Array(-17.130289798146, 20.415084409121, 25.965625000000)
    FSData_141 = pts
End Function

Private Function FSData_142() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.771388421478, 20.710939872828, 20.636390625000)
    pts(1) = Array(-19.666262220307, 24.285811295530, 20.636390625000)
    pts(2) = Array(-19.666262220307, 24.285811295530, 26.036390625000)
    pts(3) = Array(-16.771388421478, 20.710939872828, 26.036390625000)
    FSData_142 = pts
End Function

Private Function FSData_143() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.407378317429, 21.000486583619, 20.708000000000)
    pts(1) = Array(-19.239421103927, 24.625336050210, 20.708000000000)
    pts(2) = Array(-19.239421103927, 24.625336050210, 26.108000000000)
    pts(3) = Array(-16.407378317429, 21.000486583619, 26.108000000000)
    FSData_143 = pts
End Function

Private Function FSData_144() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.038370367002, 21.283636342760, 20.780359375000)
    pts(1) = Array(-18.806719473502, 24.957359688978, 20.780359375000)
    pts(2) = Array(-18.806719473502, 24.957359688978, 26.180359375000)
    pts(3) = Array(-16.038370367002, 21.283636342760, 26.180359375000)
    FSData_144 = pts
End Function

Private Function FSData_145() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-15.664476973594, 21.560302900092, 20.853375000000)
    pts(1) = Array(-18.368289134140, 25.281781074217, 20.853375000000)
    pts(2) = Array(-18.368289134140, 25.281781074217, 26.253375000000)
    pts(3) = Array(-15.664476973594, 21.560302900092, 26.253375000000)
    FSData_145 = pts
End Function

Private Function FSData_146() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-15.285812028755, 21.830401980302, 20.926953125000)
    pts(1) = Array(-17.924263635970, 25.598501384031, 20.926953125000)
    pts(2) = Array(-17.924263635970, 25.598501384031, 26.326953125000)
    pts(3) = Array(-15.285812028755, 21.830401980302, 26.326953125000)
    FSData_146 = pts
End Function

Private Function FSData_147() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.902490877495, 22.093851308592, 21.001000000000)
    pts(1) = Array(-17.474778233461, 25.907424142345, 21.001000000000)
    pts(2) = Array(-17.474778233461, 25.907424142345, 26.401000000000)
    pts(3) = Array(-14.902490877495, 22.093851308592, 26.401000000000)
    FSData_147 = pts
End Function

Private Function FSData_148() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.514630283150, 22.350570635746, 21.075421875000)
    pts(1) = Array(-17.019969844220, 26.208455248295, 21.075421875000)
    pts(2) = Array(-17.019969844220, 26.208455248295, 26.475421875000)
    pts(3) = Array(-14.514630283150, 22.350570635746, 26.475421875000)
    FSData_148 = pts
End Function

Private Function FSData_149() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.122348391815, 22.600481762569, 21.150125000000)
    pts(1) = Array(-16.559977007288, 26.501503004888, 21.150125000000)
    pts(2) = Array(-16.559977007288, 26.501503004888, 26.550125000000)
    pts(3) = Array(-14.122348391815, 22.600481762569, 26.550125000000)
    FSData_149 = pts
End Function

Private Function FSData_150() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.725764696353, 22.843508563711, 21.225015625000)
    pts(1) = Array(-16.094939840939, 26.786478146941, 21.225015625000)
    pts(2) = Array(-16.094939840939, 26.786478146941, 26.625015625000)
    pts(3) = Array(-13.725764696353, 22.843508563711, 26.625015625000)
    FSData_150 = pts
End Function

Private Function FSData_151() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.325000000000, 23.079577010855, 21.300000000000)
    pts(1) = Array(-15.625000000000, 27.063293868264, 21.300000000000)
    pts(2) = Array(-15.625000000000, 27.063293868264, 26.700000000000)
    pts(3) = Array(-13.325000000000, 23.079577010855, 26.700000000000)
    FSData_151 = pts
End Function

Private Function FSData_152() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.920176379565, 23.308615195265, 21.374984375000)
    pts(1) = Array(-15.150300632698, 27.331865848106, 21.374984375000)
    pts(2) = Array(-15.150300632698, 27.331865848106, 26.774984375000)
    pts(3) = Array(-12.920176379565, 23.308615195265, 26.774984375000)
    FSData_152 = pts
End Function

Private Function FSData_153() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.511417148244, 23.530553349690, 21.449875000000)
    pts(1) = Array(-14.670986337059, 27.592112276841, 21.449875000000)
    pts(2) = Array(-14.670986337059, 27.592112276841, 26.849875000000)
    pts(3) = Array(-12.511417148244, 23.530553349690, 26.849875000000)
    FSData_153 = pts
End Function

Private Function FSData_154() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.098846818059, 23.745323869620, 21.524578125000)
    pts(1) = Array(-14.187203116861, 27.843953880886, 21.524578125000)
    pts(2) = Array(-14.187203116861, 27.843953880886, 26.924578125000)
    pts(3) = Array(-12.098846818059, 23.745323869620, 26.924578125000)
    FSData_154 = pts
End Function

Private Function FSData_155() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-11.682591061929, 23.952861333873, 21.599000000000)
    pts(1) = Array(-13.699098337159, 28.087313946849, 21.599000000000)
    pts(2) = Array(-13.699098337159, 28.087313946849, 26.999000000000)
    pts(3) = Array(-11.682591061929, 23.952861333873, 26.999000000000)
    FSData_155 = pts
End Function

Private Function FSData_156() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-11.262776675390, 24.153102524527, 21.673046875000)
    pts(1) = Array(-13.206820679397, 28.322118344895, 21.673046875000)
    pts(2) = Array(-13.206820679397, 28.322118344895, 27.073046875000)
    pts(3) = Array(-11.262776675390, 24.153102524527, 27.073046875000)
    FSData_156 = pts
End Function

Private Function FSData_157() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.839531537970, 24.345986446175, 21.746625000000)
    pts(1) = Array(-12.710520096119, 28.548295551331, 21.746625000000)
    pts(2) = Array(-12.710520096119, 28.548295551331, 27.146625000000)
    pts(3) = Array(-10.839531537970, 24.345986446175, 27.146625000000)
    FSData_157 = pts
End Function

Private Function FSData_158() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.412984574239, 24.531454344508, 21.819640625000)
    pts(1) = Array(-12.210347765290, 28.765776670389, 21.819640625000)
    pts(2) = Array(-12.210347765290, 28.765776670389, 27.219640625000)
    pts(3) = Array(-10.412984574239, 24.531454344508, 27.219640625000)
    FSData_158 = pts
End Function

Private Function FSData_159() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.983265714534, 24.709449724205, 21.892000000000)
    pts(1) = Array(-11.706456044247, 28.974495455212, 21.892000000000)
    pts(2) = Array(-11.706456044247, 28.974495455212, 27.292000000000)
    pts(3) = Array(-9.983265714534, 24.709449724205, 27.292000000000)
    FSData_159 = pts
End Function

Private Function FSData_160() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.550505855382, 24.879918366150, 21.963609375000)
    pts(1) = Array(-11.198998423291, 29.174388328038, 21.963609375000)
    pts(2) = Array(-11.198998423291, 29.174388328038, 27.363609375000)
    pts(3) = Array(-9.550505855382, 24.879918366150, 27.363609375000)
    FSData_160 = pts
End Function

Private Function FSData_161() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.114836819629, 25.042808343944, 22.034375000000)
    pts(1) = Array(-10.688129478927, 29.365394399560, 22.034375000000)
    pts(2) = Array(-10.688129478927, 29.365394399560, 27.434375000000)
    pts(3) = Array(-9.114836819629, 25.042808343944, 27.434375000000)
    FSData_161 = pts
End Function

Private Function FSData_162() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-8.676391316283, 25.198070039722, 22.104203125000)
    pts(1) = Array(-10.174004826786, 29.547455487479, 22.104203125000)
    pts(2) = Array(-10.174004826786, 29.547455487479, 27.504203125000)
    pts(3) = Array(-8.676391316283, 25.198070039722, 27.504203125000)
    FSData_162 = pts
End Function

Private Function FSData_163() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-8.235302900092, 25.345656159266, 22.173000000000)
    pts(1) = Array(-9.656781074217, 29.720516134224, 22.173000000000)
    pts(2) = Array(-9.656781074217, 29.720516134224, 27.573000000000)
    pts(3) = Array(-8.235302900092, 25.345656159266, 27.573000000000)
    FSData_163 = pts
End Function

Private Function FSData_164() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-7.791705930861, 25.485521746415, 22.240671875000)
    pts(1) = Array(-9.136615772586, 29.884523623845, 22.240671875000)
    pts(2) = Array(-9.136615772586, 29.884523623845, 27.640671875000)
    pts(3) = Array(-7.791705930861, 25.485521746415, 27.640671875000)
    FSData_164 = pts
End Function

Private Function FSData_165() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-7.345735532523, 25.617624196756, 22.307125000000)
    pts(1) = Array(-8.613667369281, 30.039427998072, 22.307125000000)
    pts(2) = Array(-8.613667369281, 30.039427998072, 27.707125000000)
    pts(3) = Array(-7.345735532523, 25.617624196756, 27.707125000000)
    FSData_165 = pts
End Function

Private Function FSData_166() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.897527551982, 25.741923270604, 22.372265625000)
    pts(1) = Array(-8.088095159454, 30.185182071533, 22.372265625000)
    pts(2) = Array(-8.088095159454, 30.185182071533, 27.772265625000)
    pts(3) = Array(-6.897527551982, 25.741923270604, 27.772265625000)
    FSData_166 = pts
End Function

Private Function FSData_167() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.447218517731, 25.858381105255, 22.436000000000)
    pts(1) = Array(-7.560059237490, 30.321741446125, 22.436000000000)
    pts(2) = Array(-7.560059237490, 30.321741446125, 27.836000000000)
    pts(3) = Array(-6.447218517731, 25.858381105255, 27.836000000000)
    FSData_167 = pts
End Function

Private Function FSData_168() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.994945598264, 25.966962226527, 22.498234375000)
    pts(1) = Array(-7.029720448246, 30.449064524539, 22.498234375000)
    pts(2) = Array(-7.029720448246, 30.449064524539, 27.898234375000)
    pts(3) = Array(-5.994945598264, 25.966962226527, 27.898234375000)
    FSData_168 = pts
End Function

Private Function FSData_169() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.540846560293, 26.067633559556, 22.558875000000)
    pts(1) = Array(-6.497240338055, 30.567112522931, 22.558875000000)
    pts(2) = Array(-6.497240338055, 30.567112522931, 27.958875000000)
    pts(3) = Array(-5.540846560293, 26.067633559556, 27.958875000000)
    FSData_169 = pts
End Function

Private Function FSData_170() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.085059726785, 26.160364438880, 22.617828125000)
    pts(1) = Array(-5.962781105517, 30.675849482739, 22.617828125000)
    pts(2) = Array(-5.962781105517, 30.675849482739, 28.017828125000)
    pts(3) = Array(-5.085059726785, 26.160364438880, 28.017828125000)
    FSData_170 = pts
End Function

Private Function FSData_171() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-4.627723934824, 26.245126617775, 22.675000000000)
    pts(1) = Array(-5.426505552092, 30.775242281631, 22.675000000000)
    pts(2) = Array(-5.426505552092, 30.775242281631, 28.075000000000)
    pts(3) = Array(-4.627723934824, 26.245126617775, 28.075000000000)
    FSData_171 = pts
End Function

Private Function FSData_172() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-4.168978493322, 26.321894276860, 22.730296875000)
    pts(1) = Array(-4.888577032507, 30.865260643598, 22.730296875000)
    pts(2) = Array(-4.888577032507, 30.865260643598, 28.130296875000)
    pts(3) = Array(-4.168978493322, 26.321894276860, 28.130296875000)
    FSData_172 = pts
End Function

Private Function FSData_173() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-3.708963140586, 26.390644031963, 22.783625000000)
    pts(1) = Array(-4.349159405002, 30.945877148174, 22.783625000000)
    pts(2) = Array(-4.349159405002, 30.945877148174, 28.183625000000)
    pts(3) = Array(-3.708963140586, 26.390644031963, 28.183625000000)
    FSData_173 = pts
End Function

Private Function FSData_174() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-3.247818001747, 26.451354941241, 22.834890625000)
    pts(1) = Array(-3.808416981411, 31.017067238791, 22.834890625000)
    pts(2) = Array(-3.808416981411, 31.017067238791, 28.234890625000)
    pts(3) = Array(-3.247818001747, 26.451354941241, 28.234890625000)
    FSData_174 = pts
End Function

Private Function FSData_175() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.785683546083, 26.504008511564, 22.884000000000)
    pts(1) = Array(-3.266514477114, 31.078809230259, 22.884000000000)
    pts(2) = Array(-3.266514477114, 31.078809230259, 28.284000000000)
    pts(3) = Array(-2.785683546083, 26.504008511564, 28.284000000000)
    FSData_175 = pts
End Function

Private Function FSData_176() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.322700544225, 26.548588704145, 22.930859375000)
    pts(1) = Array(-2.723616960864, 31.131084315367, 22.930859375000)
    pts(2) = Array(-2.723616960864, 31.131084315367, 28.330859375000)
    pts(3) = Array(-2.322700544225, 26.548588704145, 28.330859375000)
    FSData_176 = pts
End Function

Private Function FSData_177() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.859010025281, 26.585081939424, 22.975375000000)
    pts(1) = Array(-2.179889804504, 31.173876570620, 22.975375000000)
    pts(2) = Array(-2.179889804504, 31.173876570620, 28.375375000000)
    pts(3) = Array(-1.859010025281, 26.585081939424, 28.375375000000)
    FSData_177 = pts
End Function

Private Function FSData_178() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.394753233874, 26.613477101209, 23.017453125000)
    pts(1) = Array(-1.635498632592, 31.207172961080, 23.017453125000)
    pts(2) = Array(-1.635498632592, 31.207172961080, 28.417453125000)
    pts(3) = Array(-1.394753233874, 26.613477101209, 28.417453125000)
    FSData_178 = pts
End Function

Private Function FSData_179() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-0.930071587122, 26.633765540059, 23.057000000000)
    pts(1) = Array(-1.090609271953, 31.230963344347, 23.057000000000)
    pts(2) = Array(-1.090609271953, 31.230963344347, 28.457000000000)
    pts(3) = Array(-0.930071587122, 26.633765540059, 28.457000000000)
    FSData_179 = pts
End Function

Private Function FSData_180() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-0.465106631554, 26.645941075918, 23.093921875000)
    pts(1) = Array(-0.545387701165, 31.245240473637, 23.093921875000)
    pts(2) = Array(-0.545387701165, 31.245240473637, 28.493921875000)
    pts(3) = Array(-0.465106631554, 26.645941075918, 28.493921875000)
    FSData_180 = pts
End Function

Private Function FSData_181() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(0.000000000000, 26.650000000000, 23.128125000000)
    pts(1) = Array(0.000000000000, 31.250000000000, 23.128125000000)
    pts(2) = Array(0.000000000000, 31.250000000000, 28.528125000000)
    pts(3) = Array(0.000000000000, 26.650000000000, 28.528125000000)
    FSData_181 = pts
End Function

Private Function FSData_182() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(0.465106631554, 26.645941075918, 23.159515625000)
    pts(1) = Array(0.545387701165, 31.245240473637, 23.159515625000)
    pts(2) = Array(0.545387701165, 31.245240473637, 28.559515625000)
    pts(3) = Array(0.465106631554, 26.645941075918, 28.559515625000)
    FSData_182 = pts
End Function

Private Function FSData_183() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(0.930071587122, 26.633765540059, 23.188000000000)
    pts(1) = Array(1.090609271953, 31.230963344347, 23.188000000000)
    pts(2) = Array(1.090609271953, 31.230963344347, 28.588000000000)
    pts(3) = Array(0.930071587122, 26.633765540059, 28.588000000000)
    FSData_183 = pts
End Function

Private Function FSData_184() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.394753233874, 26.613477101209, 23.213484375000)
    pts(1) = Array(1.635498632592, 31.207172961080, 23.213484375000)
    pts(2) = Array(1.635498632592, 31.207172961080, 28.613484375000)
    pts(3) = Array(1.394753233874, 26.613477101209, 28.613484375000)
    FSData_184 = pts
End Function

Private Function FSData_185() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.859010025281, 26.585081939424, 23.235875000000)
    pts(1) = Array(2.179889804504, 31.173876570620, 23.235875000000)
    pts(2) = Array(2.179889804504, 31.173876570620, 28.635875000000)
    pts(3) = Array(1.859010025281, 26.585081939424, 28.635875000000)
    FSData_185 = pts
End Function

Private Function FSData_186() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.322700544225, 26.548588704145, 23.255078125000)
    pts(1) = Array(2.723616960864, 31.131084315367, 23.255078125000)
    pts(2) = Array(2.723616960864, 31.131084315367, 28.655078125000)
    pts(3) = Array(2.322700544225, 26.548588704145, 28.655078125000)
    FSData_186 = pts
End Function

Private Function FSData_187() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.785683546083, 26.504008511564, 23.271000000000)
    pts(1) = Array(3.266514477114, 31.078809230259, 23.271000000000)
    pts(2) = Array(3.266514477114, 31.078809230259, 28.671000000000)
    pts(3) = Array(2.785683546083, 26.504008511564, 28.671000000000)
    FSData_187 = pts
End Function

Private Function FSData_188() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(3.247818001747, 26.451354941241, 23.283546875000)
    pts(1) = Array(3.808416981411, 31.017067238791, 23.283546875000)
    pts(2) = Array(3.808416981411, 31.017067238791, 28.683546875000)
    pts(3) = Array(3.247818001747, 26.451354941241, 28.683546875000)
    FSData_188 = pts
End Function

Private Function FSData_189() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(3.708963140586, 26.390644031963, 23.292625000000)
    pts(1) = Array(4.349159405002, 30.945877148174, 23.292625000000)
    pts(2) = Array(4.349159405002, 30.945877148174, 28.692625000000)
    pts(3) = Array(3.708963140586, 26.390644031963, 28.692625000000)
    FSData_189 = pts
End Function

Private Function FSData_190() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(4.168978493322, 26.321894276860, 23.298140625000)
    pts(1) = Array(4.888577032507, 30.865260643598, 23.298140625000)
    pts(2) = Array(4.888577032507, 30.865260643598, 28.698140625000)
    pts(3) = Array(4.168978493322, 26.321894276860, 28.698140625000)
    FSData_190 = pts
End Function

Private Function FSData_191() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(4.627723934824, 26.245126617775, 23.300000000000)
    pts(1) = Array(5.426505552092, 30.775242281631, 23.300000000000)
    pts(2) = Array(5.426505552092, 30.775242281631, 28.700000000000)
    pts(3) = Array(4.627723934824, 26.245126617775, 28.700000000000)
    FSData_191 = pts
End Function

Private Function FSData_192() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.085059726785, 26.160364438880, 23.300000000000)
    pts(1) = Array(5.962781105517, 30.675849482739, 23.300000000000)
    pts(2) = Array(5.962781105517, 30.675849482739, 28.700000000000)
    pts(3) = Array(5.085059726785, 26.160364438880, 28.700000000000)
    FSData_192 = pts
End Function

Private Function FSData_193() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.540846560293, 26.067633559556, 23.300000000000)
    pts(1) = Array(6.497240338055, 30.567112522931, 23.300000000000)
    pts(2) = Array(6.497240338055, 30.567112522931, 28.700000000000)
    pts(3) = Array(5.540846560293, 26.067633559556, 28.700000000000)
    FSData_193 = pts
End Function

Private Function FSData_194() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.994945598264, 25.966962226527, 23.300000000000)
    pts(1) = Array(7.029720448246, 30.449064524539, 23.300000000000)
    pts(2) = Array(7.029720448246, 30.449064524539, 28.700000000000)
    pts(3) = Array(5.994945598264, 25.966962226527, 28.700000000000)
    FSData_194 = pts
End Function

Private Function FSData_195() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(6.447218517731, 25.858381105255, 23.300000000000)
    pts(1) = Array(7.560059237490, 30.321741446125, 23.300000000000)
    pts(2) = Array(7.560059237490, 30.321741446125, 28.700000000000)
    pts(3) = Array(6.447218517731, 25.858381105255, 28.700000000000)
    FSData_195 = pts
End Function

Private Function FSData_196() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(6.897527551982, 25.741923270604, 23.300000000000)
    pts(1) = Array(8.088095159454, 30.185182071533, 23.300000000000)
    pts(2) = Array(8.088095159454, 30.185182071533, 28.700000000000)
    pts(3) = Array(6.897527551982, 25.741923270604, 28.700000000000)
    FSData_196 = pts
End Function

Private Function FSData_197() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(7.345735532523, 25.617624196756, 23.300000000000)
    pts(1) = Array(8.613667369281, 30.039427998072, 23.300000000000)
    pts(2) = Array(8.613667369281, 30.039427998072, 28.700000000000)
    pts(3) = Array(7.345735532523, 25.617624196756, 28.700000000000)
    FSData_197 = pts
End Function

Private Function FSData_198() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(7.345735532523, -25.617624196756, 19.300000000000)
    pts(1) = Array(8.613667369281, -30.039427998072, 19.300000000000)
    pts(2) = Array(8.613667369281, -30.039427998072, 24.700000000000)
    pts(3) = Array(7.345735532523, -25.617624196756, 24.700000000000)
    FSData_198 = pts
End Function

Private Function FSData_199() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(6.897527551982, -25.741923270604, 19.300000000000)
    pts(1) = Array(8.088095159454, -30.185182071533, 19.300000000000)
    pts(2) = Array(8.088095159454, -30.185182071533, 24.700000000000)
    pts(3) = Array(6.897527551982, -25.741923270604, 24.700000000000)
    FSData_199 = pts
End Function

Private Function FSData_200() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(6.447218517731, -25.858381105255, 19.300000000000)
    pts(1) = Array(7.560059237490, -30.321741446125, 19.300000000000)
    pts(2) = Array(7.560059237490, -30.321741446125, 24.700000000000)
    pts(3) = Array(6.447218517731, -25.858381105255, 24.700000000000)
    FSData_200 = pts
End Function

Private Function FSData_201() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.994945598264, -25.966962226527, 19.300000000000)
    pts(1) = Array(7.029720448246, -30.449064524539, 19.300000000000)
    pts(2) = Array(7.029720448246, -30.449064524539, 24.700000000000)
    pts(3) = Array(5.994945598264, -25.966962226527, 24.700000000000)
    FSData_201 = pts
End Function

Private Function FSData_202() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.540846560293, -26.067633559556, 19.300000000000)
    pts(1) = Array(6.497240338055, -30.567112522931, 19.300000000000)
    pts(2) = Array(6.497240338055, -30.567112522931, 24.700000000000)
    pts(3) = Array(5.540846560293, -26.067633559556, 24.700000000000)
    FSData_202 = pts
End Function

Private Function FSData_203() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(5.085059726785, -26.160364438880, 19.300000000000)
    pts(1) = Array(5.962781105517, -30.675849482739, 19.300000000000)
    pts(2) = Array(5.962781105517, -30.675849482739, 24.700000000000)
    pts(3) = Array(5.085059726785, -26.160364438880, 24.700000000000)
    FSData_203 = pts
End Function

Private Function FSData_204() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(4.627723934824, -26.245126617775, 19.300000000000)
    pts(1) = Array(5.426505552092, -30.775242281632, 19.300000000000)
    pts(2) = Array(5.426505552092, -30.775242281632, 24.700000000000)
    pts(3) = Array(4.627723934824, -26.245126617775, 24.700000000000)
    FSData_204 = pts
End Function

Private Function FSData_205() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(4.168978493322, -26.321894276860, 19.301859375000)
    pts(1) = Array(4.888577032507, -30.865260643598, 19.301859375000)
    pts(2) = Array(4.888577032507, -30.865260643598, 24.701859375000)
    pts(3) = Array(4.168978493322, -26.321894276860, 24.701859375000)
    FSData_205 = pts
End Function

Private Function FSData_206() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(3.708963140586, -26.390644031963, 19.307375000000)
    pts(1) = Array(4.349159405002, -30.945877148174, 19.307375000000)
    pts(2) = Array(4.349159405002, -30.945877148174, 24.707375000000)
    pts(3) = Array(3.708963140586, -26.390644031963, 24.707375000000)
    FSData_206 = pts
End Function

Private Function FSData_207() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(3.247818001747, -26.451354941241, 19.316453125000)
    pts(1) = Array(3.808416981411, -31.017067238791, 19.316453125000)
    pts(2) = Array(3.808416981411, -31.017067238791, 24.716453125000)
    pts(3) = Array(3.247818001747, -26.451354941241, 24.716453125000)
    FSData_207 = pts
End Function

Private Function FSData_208() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.785683546083, -26.504008511564, 19.329000000000)
    pts(1) = Array(3.266514477114, -31.078809230259, 19.329000000000)
    pts(2) = Array(3.266514477114, -31.078809230259, 24.729000000000)
    pts(3) = Array(2.785683546083, -26.504008511564, 24.729000000000)
    FSData_208 = pts
End Function

Private Function FSData_209() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(2.322700544225, -26.548588704145, 19.344921875000)
    pts(1) = Array(2.723616960864, -31.131084315367, 19.344921875000)
    pts(2) = Array(2.723616960864, -31.131084315367, 24.744921875000)
    pts(3) = Array(2.322700544225, -26.548588704145, 24.744921875000)
    FSData_209 = pts
End Function

Private Function FSData_210() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.859010025281, -26.585081939424, 19.364125000000)
    pts(1) = Array(2.179889804504, -31.173876570620, 19.364125000000)
    pts(2) = Array(2.179889804504, -31.173876570620, 24.764125000000)
    pts(3) = Array(1.859010025281, -26.585081939424, 24.764125000000)
    FSData_210 = pts
End Function

Private Function FSData_211() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.394753233874, -26.613477101209, 19.386515625000)
    pts(1) = Array(1.635498632592, -31.207172961080, 19.386515625000)
    pts(2) = Array(1.635498632592, -31.207172961080, 24.786515625000)
    pts(3) = Array(1.394753233874, -26.613477101209, 24.786515625000)
    FSData_211 = pts
End Function

Private Function FSData_212() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(0.930071587122, -26.633765540059, 19.412000000000)
    pts(1) = Array(1.090609271953, -31.230963344347, 19.412000000000)
    pts(2) = Array(1.090609271953, -31.230963344347, 24.812000000000)
    pts(3) = Array(0.930071587122, -26.633765540059, 24.812000000000)
    FSData_212 = pts
End Function

Private Function FSData_213() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(0.465106631554, -26.645941075918, 19.440484375000)
    pts(1) = Array(0.545387701165, -31.245240473637, 19.440484375000)
    pts(2) = Array(0.545387701165, -31.245240473637, 24.840484375000)
    pts(3) = Array(0.465106631554, -26.645941075918, 24.840484375000)
    FSData_213 = pts
End Function

Private Function FSData_214() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-0.000000000000, -26.650000000000, 19.471875000000)
    pts(1) = Array(-0.000000000000, -31.250000000000, 19.471875000000)
    pts(2) = Array(-0.000000000000, -31.250000000000, 24.871875000000)
    pts(3) = Array(-0.000000000000, -26.650000000000, 24.871875000000)
    FSData_214 = pts
End Function

Private Function FSData_215() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-0.465106631554, -26.645941075918, 19.506078125000)
    pts(1) = Array(-0.545387701165, -31.245240473637, 19.506078125000)
    pts(2) = Array(-0.545387701165, -31.245240473637, 24.906078125000)
    pts(3) = Array(-0.465106631554, -26.645941075918, 24.906078125000)
    FSData_215 = pts
End Function

Private Function FSData_216() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-0.930071587122, -26.633765540059, 19.543000000000)
    pts(1) = Array(-1.090609271953, -31.230963344347, 19.543000000000)
    pts(2) = Array(-1.090609271953, -31.230963344347, 24.943000000000)
    pts(3) = Array(-0.930071587122, -26.633765540059, 24.943000000000)
    FSData_216 = pts
End Function

Private Function FSData_217() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.394753233874, -26.613477101209, 19.582546875000)
    pts(1) = Array(-1.635498632592, -31.207172961080, 19.582546875000)
    pts(2) = Array(-1.635498632592, -31.207172961080, 24.982546875000)
    pts(3) = Array(-1.394753233874, -26.613477101209, 24.982546875000)
    FSData_217 = pts
End Function

Private Function FSData_218() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.859010025281, -26.585081939424, 19.624625000000)
    pts(1) = Array(-2.179889804504, -31.173876570620, 19.624625000000)
    pts(2) = Array(-2.179889804504, -31.173876570620, 25.024625000000)
    pts(3) = Array(-1.859010025281, -26.585081939424, 25.024625000000)
    FSData_218 = pts
End Function

Private Function FSData_219() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.322700544225, -26.548588704145, 19.669140625000)
    pts(1) = Array(-2.723616960864, -31.131084315367, 19.669140625000)
    pts(2) = Array(-2.723616960864, -31.131084315367, 25.069140625000)
    pts(3) = Array(-2.322700544225, -26.548588704145, 25.069140625000)
    FSData_219 = pts
End Function

Private Function FSData_220() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-2.785683546083, -26.504008511564, 19.716000000000)
    pts(1) = Array(-3.266514477114, -31.078809230259, 19.716000000000)
    pts(2) = Array(-3.266514477114, -31.078809230259, 25.116000000000)
    pts(3) = Array(-2.785683546083, -26.504008511564, 25.116000000000)
    FSData_220 = pts
End Function

Private Function FSData_221() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-3.247818001747, -26.451354941241, 19.765109375000)
    pts(1) = Array(-3.808416981411, -31.017067238791, 19.765109375000)
    pts(2) = Array(-3.808416981411, -31.017067238791, 25.165109375000)
    pts(3) = Array(-3.247818001747, -26.451354941241, 25.165109375000)
    FSData_221 = pts
End Function

Private Function FSData_222() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-3.708963140586, -26.390644031963, 19.816375000000)
    pts(1) = Array(-4.349159405002, -30.945877148174, 19.816375000000)
    pts(2) = Array(-4.349159405002, -30.945877148174, 25.216375000000)
    pts(3) = Array(-3.708963140586, -26.390644031963, 25.216375000000)
    FSData_222 = pts
End Function

Private Function FSData_223() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-4.168978493322, -26.321894276860, 19.869703125000)
    pts(1) = Array(-4.888577032507, -30.865260643598, 19.869703125000)
    pts(2) = Array(-4.888577032507, -30.865260643598, 25.269703125000)
    pts(3) = Array(-4.168978493322, -26.321894276860, 25.269703125000)
    FSData_223 = pts
End Function

Private Function FSData_224() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-4.627723934824, -26.245126617775, 19.925000000000)
    pts(1) = Array(-5.426505552092, -30.775242281631, 19.925000000000)
    pts(2) = Array(-5.426505552092, -30.775242281631, 25.325000000000)
    pts(3) = Array(-4.627723934824, -26.245126617775, 25.325000000000)
    FSData_224 = pts
End Function

Private Function FSData_225() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.085059726785, -26.160364438880, 19.982171875000)
    pts(1) = Array(-5.962781105517, -30.675849482739, 19.982171875000)
    pts(2) = Array(-5.962781105517, -30.675849482739, 25.382171875000)
    pts(3) = Array(-5.085059726785, -26.160364438880, 25.382171875000)
    FSData_225 = pts
End Function

Private Function FSData_226() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.540846560293, -26.067633559556, 20.041125000000)
    pts(1) = Array(-6.497240338055, -30.567112522931, 20.041125000000)
    pts(2) = Array(-6.497240338055, -30.567112522931, 25.441125000000)
    pts(3) = Array(-5.540846560293, -26.067633559556, 25.441125000000)
    FSData_226 = pts
End Function

Private Function FSData_227() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-5.994945598264, -25.966962226527, 20.101765625000)
    pts(1) = Array(-7.029720448246, -30.449064524539, 20.101765625000)
    pts(2) = Array(-7.029720448246, -30.449064524539, 25.501765625000)
    pts(3) = Array(-5.994945598264, -25.966962226527, 25.501765625000)
    FSData_227 = pts
End Function

Private Function FSData_228() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.447218517731, -25.858381105255, 20.164000000000)
    pts(1) = Array(-7.560059237490, -30.321741446125, 20.164000000000)
    pts(2) = Array(-7.560059237490, -30.321741446125, 25.564000000000)
    pts(3) = Array(-6.447218517731, -25.858381105255, 25.564000000000)
    FSData_228 = pts
End Function

Private Function FSData_229() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-6.897527551982, -25.741923270604, 20.227734375000)
    pts(1) = Array(-8.088095159454, -30.185182071533, 20.227734375000)
    pts(2) = Array(-8.088095159454, -30.185182071533, 25.627734375000)
    pts(3) = Array(-6.897527551982, -25.741923270604, 25.627734375000)
    FSData_229 = pts
End Function

Private Function FSData_230() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-7.345735532523, -25.617624196756, 20.292875000000)
    pts(1) = Array(-8.613667369281, -30.039427998072, 20.292875000000)
    pts(2) = Array(-8.613667369281, -30.039427998072, 25.692875000000)
    pts(3) = Array(-7.345735532523, -25.617624196756, 25.692875000000)
    FSData_230 = pts
End Function

Private Function FSData_231() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-7.791705930861, -25.485521746415, 20.359328125000)
    pts(1) = Array(-9.136615772586, -29.884523623845, 20.359328125000)
    pts(2) = Array(-9.136615772586, -29.884523623845, 25.759328125000)
    pts(3) = Array(-7.791705930861, -25.485521746415, 25.759328125000)
    FSData_231 = pts
End Function

Private Function FSData_232() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-8.235302900092, -25.345656159266, 20.427000000000)
    pts(1) = Array(-9.656781074217, -29.720516134224, 20.427000000000)
    pts(2) = Array(-9.656781074217, -29.720516134224, 25.827000000000)
    pts(3) = Array(-8.235302900092, -25.345656159266, 25.827000000000)
    FSData_232 = pts
End Function

Private Function FSData_233() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-8.676391316283, -25.198070039722, 20.495796875000)
    pts(1) = Array(-10.174004826786, -29.547455487479, 20.495796875000)
    pts(2) = Array(-10.174004826786, -29.547455487479, 25.895796875000)
    pts(3) = Array(-8.676391316283, -25.198070039722, 25.895796875000)
    FSData_233 = pts
End Function

Private Function FSData_234() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.114836819629, -25.042808343944, 20.565625000000)
    pts(1) = Array(-10.688129478927, -29.365394399560, 20.565625000000)
    pts(2) = Array(-10.688129478927, -29.365394399560, 25.965625000000)
    pts(3) = Array(-9.114836819629, -25.042808343944, 25.965625000000)
    FSData_234 = pts
End Function

Private Function FSData_235() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.550505855382, -24.879918366150, 20.636390625000)
    pts(1) = Array(-11.198998423291, -29.174388328038, 20.636390625000)
    pts(2) = Array(-11.198998423291, -29.174388328038, 26.036390625000)
    pts(3) = Array(-9.550505855382, -24.879918366150, 26.036390625000)
    FSData_235 = pts
End Function

Private Function FSData_236() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-9.983265714534, -24.709449724205, 20.708000000000)
    pts(1) = Array(-11.706456044247, -28.974495455212, 20.708000000000)
    pts(2) = Array(-11.706456044247, -28.974495455212, 26.108000000000)
    pts(3) = Array(-9.983265714534, -24.709449724205, 26.108000000000)
    FSData_236 = pts
End Function

Private Function FSData_237() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.412984574239, -24.531454344508, 20.780359375000)
    pts(1) = Array(-12.210347765290, -28.765776670389, 20.780359375000)
    pts(2) = Array(-12.210347765290, -28.765776670389, 26.180359375000)
    pts(3) = Array(-10.412984574239, -24.531454344508, 26.180359375000)
    FSData_237 = pts
End Function

Private Function FSData_238() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-10.839531537970, -24.345986446175, 20.853375000000)
    pts(1) = Array(-12.710520096119, -28.548295551331, 20.853375000000)
    pts(2) = Array(-12.710520096119, -28.548295551331, 26.253375000000)
    pts(3) = Array(-10.839531537970, -24.345986446175, 26.253375000000)
    FSData_238 = pts
End Function

Private Function FSData_239() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-11.262776675390, -24.153102524527, 20.926953125000)
    pts(1) = Array(-13.206820679397, -28.322118344895, 20.926953125000)
    pts(2) = Array(-13.206820679397, -28.322118344895, 26.326953125000)
    pts(3) = Array(-11.262776675390, -24.153102524527, 26.326953125000)
    FSData_239 = pts
End Function

Private Function FSData_240() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-11.682591061929, -23.952861333873, 21.001000000000)
    pts(1) = Array(-13.699098337159, -28.087313946849, 21.001000000000)
    pts(2) = Array(-13.699098337159, -28.087313946849, 26.401000000000)
    pts(3) = Array(-11.682591061929, -23.952861333873, 26.401000000000)
    FSData_240 = pts
End Function

Private Function FSData_241() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.098846818059, -23.745323869620, 21.075421875000)
    pts(1) = Array(-14.187203116861, -27.843953880886, 21.075421875000)
    pts(2) = Array(-14.187203116861, -27.843953880886, 26.475421875000)
    pts(3) = Array(-12.098846818059, -23.745323869620, 26.475421875000)
    FSData_241 = pts
End Function

Private Function FSData_242() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.511417148244, -23.530553349690, 21.150125000000)
    pts(1) = Array(-14.670986337059, -27.592112276841, 21.150125000000)
    pts(2) = Array(-14.670986337059, -27.592112276841, 26.550125000000)
    pts(3) = Array(-12.511417148244, -23.530553349690, 26.550125000000)
    FSData_242 = pts
End Function

Private Function FSData_243() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-12.920176379565, -23.308615195265, 21.225015625000)
    pts(1) = Array(-15.150300632698, -27.331865848106, 21.225015625000)
    pts(2) = Array(-15.150300632698, -27.331865848106, 26.625015625000)
    pts(3) = Array(-12.920176379565, -23.308615195265, 26.625015625000)
    FSData_243 = pts
End Function

Private Function FSData_244() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.325000000000, -23.079577010855, 21.300000000000)
    pts(1) = Array(-15.625000000000, -27.063293868264, 21.300000000000)
    pts(2) = Array(-15.625000000000, -27.063293868264, 26.700000000000)
    pts(3) = Array(-13.325000000000, -23.079577010855, 26.700000000000)
    FSData_244 = pts
End Function

Private Function FSData_245() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.725764696353, -22.843508563711, 21.374984375000)
    pts(1) = Array(-16.094939840939, -26.786478146941, 21.374984375000)
    pts(2) = Array(-16.094939840939, -26.786478146941, 26.774984375000)
    pts(3) = Array(-13.725764696353, -22.843508563711, 26.774984375000)
    FSData_245 = pts
End Function

Private Function FSData_246() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.122348391815, -22.600481762569, 21.449875000000)
    pts(1) = Array(-16.559977007288, -26.501503004888, 21.449875000000)
    pts(2) = Array(-16.559977007288, -26.501503004888, 26.849875000000)
    pts(3) = Array(-14.122348391815, -22.600481762569, 26.849875000000)
    FSData_246 = pts
End Function

Private Function FSData_247() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.514630283150, -22.350570635746, 21.524578125000)
    pts(1) = Array(-17.019969844220, -26.208455248295, 21.524578125000)
    pts(2) = Array(-17.019969844220, -26.208455248295, 26.924578125000)
    pts(3) = Array(-14.514630283150, -22.350570635746, 26.924578125000)
    FSData_247 = pts
End Function

Private Function FSData_248() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-14.902490877495, -22.093851308592, 21.599000000000)
    pts(1) = Array(-17.474778233461, -25.907424142345, 21.599000000000)
    pts(2) = Array(-17.474778233461, -25.907424142345, 26.999000000000)
    pts(3) = Array(-14.902490877495, -22.093851308592, 26.999000000000)
    FSData_248 = pts
End Function

Private Function FSData_249() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-15.285812028755, -21.830401980302, 21.673046875000)
    pts(1) = Array(-17.924263635970, -25.598501384031, 21.673046875000)
    pts(2) = Array(-17.924263635970, -25.598501384031, 27.073046875000)
    pts(3) = Array(-15.285812028755, -21.830401980302, 27.073046875000)
    FSData_249 = pts
End Function

Private Function FSData_250() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-15.664476973594, -21.560302900092, 21.746625000000)
    pts(1) = Array(-18.368289134140, -25.281781074217, 21.746625000000)
    pts(2) = Array(-18.368289134140, -25.281781074217, 27.146625000000)
    pts(3) = Array(-15.664476973594, -21.560302900092, 27.146625000000)
    FSData_250 = pts
End Function

Private Function FSData_251() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.038370367002, -21.283636342760, 21.819640625000)
    pts(1) = Array(-18.806719473502, -24.957359688978, 21.819640625000)
    pts(2) = Array(-18.806719473502, -24.957359688978, 27.219640625000)
    pts(3) = Array(-16.038370367002, -21.283636342760, 27.219640625000)
    FSData_251 = pts
End Function

Private Function FSData_252() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.407378317429, -21.000486583619, 21.892000000000)
    pts(1) = Array(-19.239421103927, -24.625336050210, 21.892000000000)
    pts(2) = Array(-19.239421103927, -24.625336050210, 27.292000000000)
    pts(3) = Array(-16.407378317429, -21.000486583619, 27.292000000000)
    FSData_252 = pts
End Function

Private Function FSData_253() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-16.771388421478, -20.710939872828, 21.963609375000)
    pts(1) = Array(-19.666262220307, -24.285811295530, 21.963609375000)
    pts(2) = Array(-19.666262220307, -24.285811295530, 27.363609375000)
    pts(3) = Array(-16.771388421478, -20.710939872828, 27.363609375000)
    FSData_253 = pts
End Function

Private Function FSData_254() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.130289798146, -20.415084409121, 22.034375000000)
    pts(1) = Array(-20.087112802704, -23.938888847468, 22.034375000000)
    pts(2) = Array(-20.087112802704, -23.938888847468, 27.434375000000)
    pts(3) = Array(-17.130289798146, -20.415084409121, 27.434375000000)
    FSData_254 = pts
End Function

Private Function FSData_255() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.483973122597, -20.113010312937, 22.104203125000)
    pts(1) = Array(-20.501844655953, -23.584674381962, 22.104203125000)
    pts(2) = Array(-20.501844655953, -23.584674381962, 27.504203125000)
    pts(3) = Array(-17.483973122597, -20.113010312937, 27.504203125000)
    FSData_255 = pts
End Function

Private Function FSData_256() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-17.832330659464, -19.804809598973, 22.173000000000)
    pts(1) = Array(-20.910331448714, -23.223275796169, 22.173000000000)
    pts(2) = Array(-20.910331448714, -23.223275796169, 27.573000000000)
    pts(3) = Array(-17.832330659464, -19.804809598973, 27.573000000000)
    FSData_256 = pts
End Function

Private Function FSData_257() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-18.175256295666, -19.490576148151, 22.240671875000)
    pts(1) = Array(-21.312448751953, -22.854803175599, 22.240671875000)
    pts(2) = Array(-21.312448751953, -22.854803175599, 27.640671875000)
    pts(3) = Array(-18.175256295666, -19.490576148151, 27.640671875000)
    FSData_257 = pts
End Function

Private Function FSData_258() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-18.512645572732, -19.170405679025, 22.307125000000)
    pts(1) = Array(-21.708074076844, -22.479368760583, 22.307125000000)
    pts(2) = Array(-21.708074076844, -22.479368760583, 27.707125000000)
    pts(3) = Array(-18.512645572732, -19.170405679025, 27.707125000000)
    FSData_258 = pts
End Function

Private Function FSData_259() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-18.844395718621, -18.844395718621, 22.372265625000)
    pts(1) = Array(-22.097086912080, -22.097086912080, 22.372265625000)
    pts(2) = Array(-22.097086912080, -22.097086912080, 27.772265625000)
    pts(3) = Array(-18.844395718621, -18.844395718621, 27.772265625000)
    FSData_259 = pts
End Function

Private Function FSData_260() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.170405679025, -18.512645572732, 22.436000000000)
    pts(1) = Array(-22.479368760583, -21.708074076844, 22.436000000000)
    pts(2) = Array(-22.479368760583, -21.708074076844, 27.836000000000)
    pts(3) = Array(-19.170405679025, -18.512645572732, 27.836000000000)
    FSData_260 = pts
End Function

Private Function FSData_261() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.490576148151, -18.175256295666, 22.498234375000)
    pts(1) = Array(-22.854803175599, -21.312448751953, 22.498234375000)
    pts(2) = Array(-22.854803175599, -21.312448751953, 27.898234375000)
    pts(3) = Array(-19.490576148151, -18.175256295666, 27.898234375000)
    FSData_261 = pts
End Function

Private Function FSData_262() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-19.804809598973, -17.832330659464, 22.558875000000)
    pts(1) = Array(-23.223275796169, -20.910331448714, 22.558875000000)
    pts(2) = Array(-23.223275796169, -20.910331448714, 27.958875000000)
    pts(3) = Array(-19.804809598973, -17.832330659464, 27.958875000000)
    FSData_262 = pts
End Function

Private Function FSData_263() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.113010312937, -17.483973122597, 22.617828125000)
    pts(1) = Array(-23.584674381962, -20.501844655953, 22.617828125000)
    pts(2) = Array(-23.584674381962, -20.501844655953, 28.017828125000)
    pts(3) = Array(-20.113010312937, -17.483973122597, 28.017828125000)
    FSData_263 = pts
End Function

Private Function FSData_264() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.415084409121, -17.130289798146, 22.675000000000)
    pts(1) = Array(-23.938888847468, -20.087112802704, 22.675000000000)
    pts(2) = Array(-23.938888847468, -20.087112802704, 28.075000000000)
    pts(3) = Array(-20.415084409121, -17.130289798146, 28.075000000000)
    FSData_264 = pts
End Function

Private Function FSData_265() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-20.710939872828, -16.771388421478, 22.730296875000)
    pts(1) = Array(-24.285811295530, -19.666262220307, 22.730296875000)
    pts(2) = Array(-24.285811295530, -19.666262220307, 28.130296875000)
    pts(3) = Array(-20.710939872828, -16.771388421478, 28.130296875000)
    FSData_265 = pts
End Function

Private Function FSData_266() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-21.000486583619, -16.407378317429, 22.783625000000)
    pts(1) = Array(-24.625336050210, -19.239421103927, 22.783625000000)
    pts(2) = Array(-24.625336050210, -19.239421103927, 28.183625000000)
    pts(3) = Array(-21.000486583619, -16.407378317429, 28.183625000000)
    FSData_266 = pts
End Function

Private Function FSData_267() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-21.283636342760, -16.038370367002, 22.834890625000)
    pts(1) = Array(-24.957359688978, -18.806719473502, 22.834890625000)
    pts(2) = Array(-24.957359688978, -18.806719473502, 28.234890625000)
    pts(3) = Array(-21.283636342760, -16.038370367002, 28.234890625000)
    FSData_267 = pts
End Function

Private Function FSData_268() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-21.560302900092, -15.664476973594, 22.884000000000)
    pts(1) = Array(-25.281781074217, -18.368289134140, 22.884000000000)
    pts(2) = Array(-25.281781074217, -18.368289134140, 28.284000000000)
    pts(3) = Array(-21.560302900092, -15.664476973594, 28.284000000000)
    FSData_268 = pts
End Function

Private Function FSData_269() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-21.830401980302, -15.285812028755, 22.930859375000)
    pts(1) = Array(-25.598501384031, -17.924263635970, 22.930859375000)
    pts(2) = Array(-25.598501384031, -17.924263635970, 28.330859375000)
    pts(3) = Array(-21.830401980302, -15.285812028755, 28.330859375000)
    FSData_269 = pts
End Function

Private Function FSData_270() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.093851308592, -14.902490877495, 22.975375000000)
    pts(1) = Array(-25.907424142345, -17.474778233461, 22.975375000000)
    pts(2) = Array(-25.907424142345, -17.474778233461, 28.375375000000)
    pts(3) = Array(-22.093851308592, -14.902490877495, 28.375375000000)
    FSData_270 = pts
End Function

Private Function FSData_271() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.350570635746, -14.514630283150, 23.017453125000)
    pts(1) = Array(-26.208455248295, -17.019969844220, 23.017453125000)
    pts(2) = Array(-26.208455248295, -17.019969844220, 28.417453125000)
    pts(3) = Array(-22.350570635746, -14.514630283150, 28.417453125000)
    FSData_271 = pts
End Function

Private Function FSData_272() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.600481762569, -14.122348391815, 23.057000000000)
    pts(1) = Array(-26.501503004888, -16.559977007288, 23.057000000000)
    pts(2) = Array(-26.501503004888, -16.559977007288, 28.457000000000)
    pts(3) = Array(-22.600481762569, -14.122348391815, 28.457000000000)
    FSData_272 = pts
End Function

Private Function FSData_273() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-22.843508563711, -13.725764696353, 23.093921875000)
    pts(1) = Array(-26.786478146941, -16.094939840939, 23.093921875000)
    pts(2) = Array(-26.786478146941, -16.094939840939, 28.493921875000)
    pts(3) = Array(-22.843508563711, -13.725764696353, 28.493921875000)
    FSData_273 = pts
End Function

Private Function FSData_274() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.079577010855, -13.325000000000, 23.128125000000)
    pts(1) = Array(-27.063293868264, -15.625000000000, 23.128125000000)
    pts(2) = Array(-27.063293868264, -15.625000000000, 28.528125000000)
    pts(3) = Array(-23.079577010855, -13.325000000000, 28.528125000000)
    FSData_274 = pts
End Function

Private Function FSData_275() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.308615195265, -12.920176379565, 23.159515625000)
    pts(1) = Array(-27.331865848106, -15.150300632698, 23.159515625000)
    pts(2) = Array(-27.331865848106, -15.150300632698, 28.559515625000)
    pts(3) = Array(-23.308615195265, -12.920176379565, 28.559515625000)
    FSData_275 = pts
End Function

Private Function FSData_276() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.530553349690, -12.511417148244, 23.188000000000)
    pts(1) = Array(-27.592112276841, -14.670986337059, 23.188000000000)
    pts(2) = Array(-27.592112276841, -14.670986337059, 28.588000000000)
    pts(3) = Array(-23.530553349690, -12.511417148244, 28.588000000000)
    FSData_276 = pts
End Function

Private Function FSData_277() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.745323869620, -12.098846818059, 23.213484375000)
    pts(1) = Array(-27.843953880886, -14.187203116861, 23.213484375000)
    pts(2) = Array(-27.843953880886, -14.187203116861, 28.613484375000)
    pts(3) = Array(-23.745323869620, -12.098846818059, 28.613484375000)
    FSData_277 = pts
End Function

Private Function FSData_278() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-23.952861333873, -11.682591061929, 23.235875000000)
    pts(1) = Array(-28.087313946849, -13.699098337159, 23.235875000000)
    pts(2) = Array(-28.087313946849, -13.699098337159, 28.635875000000)
    pts(3) = Array(-23.952861333873, -11.682591061929, 28.635875000000)
    FSData_278 = pts
End Function

Private Function FSData_279() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.153102524527, -11.262776675390, 23.255078125000)
    pts(1) = Array(-28.322118344895, -13.206820679397, 23.255078125000)
    pts(2) = Array(-28.322118344895, -13.206820679397, 28.655078125000)
    pts(3) = Array(-24.153102524527, -11.262776675390, 28.655078125000)
    FSData_279 = pts
End Function

Private Function FSData_280() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.345986446175, -10.839531537970, 23.271000000000)
    pts(1) = Array(-28.548295551331, -12.710520096119, 23.271000000000)
    pts(2) = Array(-28.548295551331, -12.710520096119, 28.671000000000)
    pts(3) = Array(-24.345986446175, -10.839531537970, 28.671000000000)
    FSData_280 = pts
End Function

Private Function FSData_281() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.531454344508, -10.412984574239, 23.283546875000)
    pts(1) = Array(-28.765776670389, -12.210347765290, 23.283546875000)
    pts(2) = Array(-28.765776670389, -12.210347765290, 28.683546875000)
    pts(3) = Array(-24.531454344508, -10.412984574239, 28.683546875000)
    FSData_281 = pts
End Function

Private Function FSData_282() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.709449724205, -9.983265714534, 23.292625000000)
    pts(1) = Array(-28.974495455212, -11.706456044247, 23.292625000000)
    pts(2) = Array(-28.974495455212, -11.706456044247, 28.692625000000)
    pts(3) = Array(-24.709449724205, -9.983265714534, 28.692625000000)
    FSData_282 = pts
End Function

Private Function FSData_283() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.879918366150, -9.550505855382, 23.298140625000)
    pts(1) = Array(-29.174388328038, -11.198998423291, 23.298140625000)
    pts(2) = Array(-29.174388328038, -11.198998423291, 28.698140625000)
    pts(3) = Array(-24.879918366150, -9.550505855382, 28.698140625000)
    FSData_283 = pts
End Function

Private Function FSData_284() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.042808343944, -9.114836819629, 23.300000000000)
    pts(1) = Array(-29.365394399560, -10.688129478927, 23.300000000000)
    pts(2) = Array(-29.365394399560, -10.688129478927, 28.700000000000)
    pts(3) = Array(-25.042808343944, -9.114836819629, 28.700000000000)
    FSData_284 = pts
End Function

Private Function FSData_285() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.198070039722, -8.676391316283, 23.300000000000)
    pts(1) = Array(-29.547455487479, -10.174004826786, 23.300000000000)
    pts(2) = Array(-29.547455487479, -10.174004826786, 28.700000000000)
    pts(3) = Array(-25.198070039722, -8.676391316283, 28.700000000000)
    FSData_285 = pts
End Function

Private Function FSData_286() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.345656159266, -8.235302900092, 23.300000000000)
    pts(1) = Array(-29.720516134224, -9.656781074217, 23.300000000000)
    pts(2) = Array(-29.720516134224, -9.656781074217, 28.700000000000)
    pts(3) = Array(-25.345656159266, -8.235302900092, 28.700000000000)
    FSData_286 = pts
End Function

Private Function FSData_287() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.485521746415, -7.791705930861, 23.300000000000)
    pts(1) = Array(-29.884523623845, -9.136615772586, 23.300000000000)
    pts(2) = Array(-29.884523623845, -9.136615772586, 28.700000000000)
    pts(3) = Array(-25.485521746415, -7.791705930861, 28.700000000000)
    FSData_287 = pts
End Function

Private Function FSData_288() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.617624196756, -7.345735532523, 23.300000000000)
    pts(1) = Array(-30.039427998072, -8.613667369281, 23.300000000000)
    pts(2) = Array(-30.039427998072, -8.613667369281, 28.700000000000)
    pts(3) = Array(-25.617624196756, -7.345735532523, 28.700000000000)
    FSData_288 = pts
End Function

Private Function FSData_289() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.741923270604, -6.897527551982, 23.300000000000)
    pts(1) = Array(-30.185182071533, -8.088095159454, 23.300000000000)
    pts(2) = Array(-30.185182071533, -8.088095159454, 28.700000000000)
    pts(3) = Array(-25.741923270604, -6.897527551982, 28.700000000000)
    FSData_289 = pts
End Function

Private Function FSData_290() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-25.858381105255, -6.447218517731, 23.300000000000)
    pts(1) = Array(-30.321741446125, -7.560059237490, 23.300000000000)
    pts(2) = Array(-30.321741446125, -7.560059237490, 28.700000000000)
    pts(3) = Array(-25.858381105255, -6.447218517731, 28.700000000000)
    FSData_290 = pts
End Function

Private Function FSData_291() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.987111129080, 0.893836603677, 15.150000000000)
    pts(1) = Array(30.961717818547, 1.540139514675, 15.150000000000)
    pts(2) = Array(31.183412171574, 1.611033934388, 15.150000000000)
    pts(3) = Array(31.404633938134, 1.688628796329, 15.150000000000)
    pts(4) = Array(31.625369983547, 1.772455473001, 15.150000000000)
    pts(5) = Array(31.845602167721, 1.862155357441, 15.150000000000)
    pts(6) = Array(32.065308962267, 1.957443780632, 15.150000000000)
    pts(7) = Array(32.284466417400, 2.058088468403, 15.150000000000)
    pts(8) = Array(32.503048780631, 2.163895783978, 15.150000000000)
    pts(9) = Array(32.721028913739, 2.274701480689, 15.150000000000)
    pts(10) = Array(33.319584320698, 2.316311873629, 15.150000000000)
    pts(11) = Array(33.398606897115, 0.305052998677, 15.150000000000)
    pts(12) = Array(32.798631922915, 0.299573004689, 15.150000000000)
    pts(13) = Array(32.572630035045, 0.392934600286, 15.150000000000)
    pts(14) = Array(32.346419941990, 0.481265972658, 15.150000000000)
    pts(15) = Array(32.120041584959, 0.564405510693, 15.150000000000)
    pts(16) = Array(31.893535829620, 0.642162195051, 15.150000000000)
    pts(17) = Array(31.666944775951, 0.714306346663, 15.150000000000)
    pts(18) = Array(31.440312220850, 0.780555863149, 15.150000000000)
    pts(19) = Array(31.213684384465, 0.840554666267, 15.150000000000)
    FSData_291 = pts
End Function

Private Function FSData_292() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.987111129080, 0.893836603677, 18.250000000000)
    pts(1) = Array(30.961717818547, 1.540139514675, 18.250000000000)
    pts(2) = Array(31.183412171574, 1.611033934388, 18.250000000000)
    pts(3) = Array(31.404633938134, 1.688628796329, 18.250000000000)
    pts(4) = Array(31.625369983547, 1.772455473001, 18.250000000000)
    pts(5) = Array(31.845602167721, 1.862155357441, 18.250000000000)
    pts(6) = Array(32.065308962267, 1.957443780632, 18.250000000000)
    pts(7) = Array(32.284466417400, 2.058088468403, 18.250000000000)
    pts(8) = Array(32.503048780631, 2.163895783978, 18.250000000000)
    pts(9) = Array(32.721028913739, 2.274701480689, 18.250000000000)
    pts(10) = Array(33.319584320698, 2.316311873629, 18.250000000000)
    pts(11) = Array(33.398606897115, 0.305052998677, 18.250000000000)
    pts(12) = Array(32.798631922915, 0.299573004689, 18.250000000000)
    pts(13) = Array(32.572630035045, 0.392934600286, 18.250000000000)
    pts(14) = Array(32.346419941990, 0.481265972658, 18.250000000000)
    pts(15) = Array(32.120041584959, 0.564405510693, 18.250000000000)
    pts(16) = Array(31.893535829620, 0.642162195051, 18.250000000000)
    pts(17) = Array(31.666944775951, 0.714306346663, 18.250000000000)
    pts(18) = Array(31.440312220850, 0.780555863149, 18.250000000000)
    pts(19) = Array(31.213684384465, 0.840554666267, 18.250000000000)
    FSData_292 = pts
End Function

Private Function FSData_293() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.821458595241, 3.322301922137, 15.150000000000)
    pts(1) = Array(30.745435221847, 3.964620160771, 15.150000000000)
    pts(2) = Array(30.960883853107, 4.052689975114, 15.150000000000)
    pts(3) = Array(31.175335644090, 4.147402497751, 15.150000000000)
    pts(4) = Array(31.388814268692, 4.248289515272, 15.150000000000)
    pts(5) = Array(31.601329778721, 4.354992102925, 15.150000000000)
    pts(6) = Array(31.812883047026, 4.467224780134, 15.150000000000)
    pts(7) = Array(32.023468421672, 4.584754109672, 15.150000000000)
    pts(8) = Array(32.233075422124, 4.707385031166, 15.150000000000)
    pts(9) = Array(32.441689880464, 4.834951675024, 15.150000000000)
    pts(10) = Array(33.035135427058, 4.923395912982, 15.150000000000)
    pts(11) = Array(33.271715955847, 2.924537117806, 15.150000000000)
    pts(12) = Array(32.674020459634, 2.872000522875, 15.150000000000)
    pts(13) = Array(32.441390193794, 2.947342412076, 15.150000000000)
    pts(14) = Array(32.208947031361, 3.017653248961, 15.150000000000)
    pts(15) = Array(31.976743470282, 3.082775054361, 15.150000000000)
    pts(16) = Array(31.744835237452, 3.142520604070, 15.150000000000)
    pts(17) = Array(31.513282323483, 3.196664230169, 15.150000000000)
    pts(18) = Array(31.282150523759, 3.244928136169, 15.150000000000)
    pts(19) = Array(31.051513853523, 3.286960967891, 15.150000000000)
    FSData_293 = pts
End Function

Private Function FSData_294() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.821458595241, 3.322301922137, 18.250000000000)
    pts(1) = Array(30.745435221847, 3.964620160771, 18.250000000000)
    pts(2) = Array(30.960883853107, 4.052689975114, 18.250000000000)
    pts(3) = Array(31.175335644090, 4.147402497751, 18.250000000000)
    pts(4) = Array(31.388814268692, 4.248289515272, 18.250000000000)
    pts(5) = Array(31.601329778721, 4.354992102925, 18.250000000000)
    pts(6) = Array(31.812883047026, 4.467224780134, 18.250000000000)
    pts(7) = Array(32.023468421672, 4.584754109672, 18.250000000000)
    pts(8) = Array(32.233075422124, 4.707385031166, 18.250000000000)
    pts(9) = Array(32.441689880464, 4.834951675024, 18.250000000000)
    pts(10) = Array(33.035135427058, 4.923395912982, 18.250000000000)
    pts(11) = Array(33.271715955847, 2.924537117806, 18.250000000000)
    pts(12) = Array(32.674020459634, 2.872000522875, 18.250000000000)
    pts(13) = Array(32.441390193794, 2.947342412076, 18.250000000000)
    pts(14) = Array(32.208947031361, 3.017653248961, 18.250000000000)
    pts(15) = Array(31.976743470282, 3.082775054361, 18.250000000000)
    pts(16) = Array(31.744835237452, 3.142520604070, 18.250000000000)
    pts(17) = Array(31.513282323483, 3.196664230169, 18.250000000000)
    pts(18) = Array(31.282150523759, 3.244928136169, 18.250000000000)
    pts(19) = Array(31.051513853523, 3.286960967891, 18.250000000000)
    FSData_294 = pts
End Function

Private Function FSData_295() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.465781519988, 5.730284144469, 15.150000000000)
    pts(1) = Array(30.339596793110, 6.364657605205, 15.150000000000)
    pts(2) = Array(30.547471390148, 6.469359834487, 15.150000000000)
    pts(3) = Array(30.753831038948, 6.580606083624, 15.150000000000)
    pts(4) = Array(30.958736076030, 6.697931439982, 15.150000000000)
    pts(5) = Array(31.162224683126, 6.820978873913, 15.150000000000)
    pts(6) = Array(31.364320128943, 6.949463853363, 15.150000000000)
    pts(7) = Array(31.565035094241, 7.083153217270, 15.150000000000)
    pts(8) = Array(31.764374435055, 7.221851684273, 15.150000000000)
    pts(9) = Array(31.962337041119, 7.365392784498, 15.150000000000)
    pts(10) = Array(32.547013938212, 7.500125579336, 15.150000000000)
    pts(11) = Array(32.939693821744, 5.525990493096, 15.150000000000)
    pts(12) = Array(32.347962795006, 5.426721202801, 15.150000000000)
    pts(13) = Array(32.110138394141, 5.483578877805, 15.150000000000)
    pts(14) = Array(31.872895251723, 5.535435689512, 15.150000000000)
    pts(15) = Array(31.636298098764, 5.582138264692, 15.150000000000)
    pts(16) = Array(31.400417179817, 5.623504328571, 15.150000000000)
    pts(17) = Array(31.165330006262, 5.659313615697, 15.150000000000)
    pts(18) = Array(30.931123966318, 5.689294348182, 15.150000000000)
    pts(19) = Array(30.697900413998, 5.713102062123, 15.150000000000)
    FSData_295 = pts
End Function

Private Function FSData_296() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.465781519988, 5.730284144469, 18.250000000000)
    pts(1) = Array(30.339596793110, 6.364657605205, 18.250000000000)
    pts(2) = Array(30.547471390148, 6.469359834487, 18.250000000000)
    pts(3) = Array(30.753831038948, 6.580606083624, 18.250000000000)
    pts(4) = Array(30.958736076030, 6.697931439982, 18.250000000000)
    pts(5) = Array(31.162224683126, 6.820978873913, 18.250000000000)
    pts(6) = Array(31.364320128943, 6.949463853363, 18.250000000000)
    pts(7) = Array(31.565035094241, 7.083153217270, 18.250000000000)
    pts(8) = Array(31.764374435055, 7.221851684273, 18.250000000000)
    pts(9) = Array(31.962337041119, 7.365392784498, 18.250000000000)
    pts(10) = Array(32.547013938212, 7.500125579336, 18.250000000000)
    pts(11) = Array(32.939693821744, 5.525990493096, 18.250000000000)
    pts(12) = Array(32.347962795006, 5.426721202801, 18.250000000000)
    pts(13) = Array(32.110138394141, 5.483578877805, 18.250000000000)
    pts(14) = Array(31.872895251723, 5.535435689512, 18.250000000000)
    pts(15) = Array(31.636298098764, 5.582138264692, 18.250000000000)
    pts(16) = Array(31.400417179817, 5.623504328571, 18.250000000000)
    pts(17) = Array(31.165330006262, 5.659313615697, 18.250000000000)
    pts(18) = Array(30.931123966318, 5.689294348182, 18.250000000000)
    pts(19) = Array(30.697900413998, 5.713102062123, 18.250000000000)
    FSData_296 = pts
End Function

Private Function FSData_297() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(29.922272770764, 8.102937259537, 15.150000000000)
    pts(1) = Array(29.746704661204, 8.725454819040, 15.150000000000)
    pts(2) = Array(29.945723608004, 8.846143939200, 15.150000000000)
    pts(3) = Array(30.142718838765, 8.973238044717, 15.150000000000)
    pts(4) = Array(30.337786980636, 9.106278390077, 15.150000000000)
    pts(5) = Array(30.530994109869, 9.244912041937, 15.150000000000)
    pts(6) = Array(30.722385747570, 9.388857171005, 15.150000000000)
    pts(7) = Array(30.911992829015, 9.537882329896, 15.150000000000)
    pts(8) = Array(31.099835516867, 9.691793220237, 15.150000000000)
    pts(9) = Array(31.285925765359, 9.850423798214, 15.150000000000)
    pts(10) = Array(31.858229285457, 10.030614477449, 15.150000000000)
    pts(11) = Array(32.404587521669, 8.093374299418, 15.150000000000)
    pts(12) = Array(31.822469182957, 7.947984341944, 15.150000000000)
    pts(13) = Array(31.580916913584, 7.986007256277, 15.150000000000)
    pts(14) = Array(31.340336474044, 8.019090328318, 15.150000000000)
    pts(15) = Array(31.100804429331, 8.047085736372, 15.150000000000)
    pts(16) = Array(30.862405108570, 8.069817278881, 15.150000000000)
    pts(17) = Array(30.625233066028, 8.087071450871, 15.150000000000)
    pts(18) = Array(30.389396743983, 8.098584168655, 15.150000000000)
    pts(19) = Array(30.155024210333, 8.104019982343, 15.150000000000)
    FSData_297 = pts
End Function

Private Function FSData_298() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(29.922272770764, 8.102937259537, 18.250000000000)
    pts(1) = Array(29.746704661204, 8.725454819040, 18.250000000000)
    pts(2) = Array(29.945723608004, 8.846143939200, 18.250000000000)
    pts(3) = Array(30.142718838765, 8.973238044717, 18.250000000000)
    pts(4) = Array(30.337786980636, 9.106278390077, 18.250000000000)
    pts(5) = Array(30.530994109869, 9.244912041937, 18.250000000000)
    pts(6) = Array(30.722385747570, 9.388857171005, 18.250000000000)
    pts(7) = Array(30.911992829015, 9.537882329896, 18.250000000000)
    pts(8) = Array(31.099835516867, 9.691793220237, 18.250000000000)
    pts(9) = Array(31.285925765359, 9.850423798214, 18.250000000000)
    pts(10) = Array(31.858229285457, 10.030614477449, 18.250000000000)
    pts(11) = Array(32.404587521669, 8.093374299418, 18.250000000000)
    pts(12) = Array(31.822469182957, 7.947984341944, 18.250000000000)
    pts(13) = Array(31.580916913584, 7.986007256277, 18.250000000000)
    pts(14) = Array(31.340336474044, 8.019090328318, 18.250000000000)
    pts(15) = Array(31.100804429331, 8.047085736372, 18.250000000000)
    pts(16) = Array(30.862405108570, 8.069817278881, 18.250000000000)
    pts(17) = Array(30.625233066028, 8.087071450871, 18.250000000000)
    pts(18) = Array(30.389396743983, 8.098584168655, 18.250000000000)
    pts(19) = Array(30.155024210333, 8.104019982343, 18.250000000000)
    FSData_298 = pts
End Function

Private Function FSData_299() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(29.194283259743, 10.425633071901, 15.150000000000)
    pts(1) = Array(28.970414203278, 11.032456702408, 15.150000000000)
    pts(2) = Array(29.159350481852, 11.168388624887, 15.150000000000)
    pts(3) = Array(29.345766753470, 11.310547009360, 15.150000000000)
    pts(4) = Array(29.529795340167, 11.458482105752, 15.150000000000)
    pts(5) = Array(29.711529805339, 11.611847252978, 15.150000000000)
    pts(6) = Array(29.891037641833, 11.770365062076, 15.150000000000)
    pts(7) = Array(30.068367848717, 11.933807226290, 15.150000000000)
    pts(8) = Array(30.243555770973, 12.101981628149, 15.150000000000)
    pts(9) = Array(30.416626353631, 12.274723673615, 15.150000000000)
    pts(10) = Array(30.973028055222, 12.499261301791, 15.150000000000)
    pts(11) = Array(31.669696163905, 10.610859761864, 15.150000000000)
    pts(12) = Array(31.100779466350, 10.420245514645, 15.150000000000)
    pts(13) = Array(30.856988578534, 10.439199244397, 15.150000000000)
    pts(14) = Array(30.614554100283, 10.453304608632, 15.150000000000)
    pts(15) = Array(30.373563958524, 10.462420248560, 15.150000000000)
    pts(16) = Array(30.134116047037, 10.466377122181, 15.150000000000)
    pts(17) = Array(29.896321380018, 10.464969801327, 15.150000000000)
    pts(18) = Array(29.660308785221, 10.457943524677, 15.150000000000)
    pts(19) = Array(29.426232254848, 10.444973924511, 15.150000000000)
    FSData_299 = pts
End Function

Private Function FSData_300() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(29.194283259743, 10.425633071901, 18.250000000000)
    pts(1) = Array(28.970414203278, 11.032456702408, 18.250000000000)
    pts(2) = Array(29.159350481852, 11.168388624887, 18.250000000000)
    pts(3) = Array(29.345766753470, 11.310547009360, 18.250000000000)
    pts(4) = Array(29.529795340167, 11.458482105752, 18.250000000000)
    pts(5) = Array(29.711529805339, 11.611847252978, 18.250000000000)
    pts(6) = Array(29.891037641833, 11.770365062076, 18.250000000000)
    pts(7) = Array(30.068367848717, 11.933807226290, 18.250000000000)
    pts(8) = Array(30.243555770973, 12.101981628149, 18.250000000000)
    pts(9) = Array(30.416626353631, 12.274723673615, 18.250000000000)
    pts(10) = Array(30.973028055222, 12.499261301791, 18.250000000000)
    pts(11) = Array(31.669696163905, 10.610859761864, 18.250000000000)
    pts(12) = Array(31.100779466350, 10.420245514645, 18.250000000000)
    pts(13) = Array(30.856988578534, 10.439199244397, 18.250000000000)
    pts(14) = Array(30.614554100283, 10.453304608632, 18.250000000000)
    pts(15) = Array(30.373563958524, 10.462420248560, 18.250000000000)
    pts(16) = Array(30.134116047037, 10.466377122181, 18.250000000000)
    pts(17) = Array(29.896321380018, 10.464969801327, 18.250000000000)
    pts(18) = Array(29.660308785221, 10.457943524677, 18.250000000000)
    pts(19) = Array(29.426232254848, 10.444973924511, 18.250000000000)
    FSData_300 = pts
End Function

Private Function FSData_301() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(28.286301284341, 12.684051389501, 15.150000000000)
    pts(1) = Array(28.015511508149, 13.271439821541, 15.150000000000)
    pts(2) = Array(28.193200263512, 13.421776480835, 15.150000000000)
    pts(3) = Array(28.367888257682, 13.578122690552, 15.150000000000)
    pts(4) = Array(28.539742691774, 13.740040468913, 15.150000000000)
    pts(5) = Array(28.708884039472, 13.907191564372, 15.150000000000)
    pts(6) = Array(28.875401349256, 14.079304738495, 15.150000000000)
    pts(7) = Array(29.039361381885, 14.256156232741, 15.150000000000)
    pts(8) = Array(29.200814446748, 14.437557295007, 15.150000000000)
    pts(9) = Array(29.359798325877, 14.623345795809, 15.150000000000)
    pts(10) = Array(29.896867807448, 14.890846023781, 15.150000000000)
    pts(11) = Array(30.739550598048, 13.062925745409, 15.150000000000)
    pts(12) = Array(30.187343102274, 12.828262408665, 15.150000000000)
    pts(13) = Array(29.942816647908, 12.828030097789, 15.150000000000)
    pts(14) = Array(29.700022820122, 12.823070789957, 15.150000000000)
    pts(15) = Array(29.459060365677, 12.813250460808, 15.150000000000)
    pts(16) = Array(29.220040139464, 12.798408270098, 15.150000000000)
    pts(17) = Array(28.983088931165, 12.778348133001, 15.150000000000)
    pts(18) = Array(28.748355159745, 12.752826181250, 15.150000000000)
    pts(19) = Array(28.516017792297, 12.721531129129, 15.150000000000)
    FSData_301 = pts
End Function

Private Function FSData_302() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(28.286301284341, 12.684051389501, 18.250000000000)
    pts(1) = Array(28.015511508149, 13.271439821541, 18.250000000000)
    pts(2) = Array(28.193200263512, 13.421776480835, 18.250000000000)
    pts(3) = Array(28.367888257682, 13.578122690552, 18.250000000000)
    pts(4) = Array(28.539742691774, 13.740040468913, 18.250000000000)
    pts(5) = Array(28.708884039472, 13.907191564372, 18.250000000000)
    pts(6) = Array(28.875401349256, 14.079304738495, 18.250000000000)
    pts(7) = Array(29.039361381885, 14.256156232741, 18.250000000000)
    pts(8) = Array(29.200814446748, 14.437557295007, 18.250000000000)
    pts(9) = Array(29.359798325877, 14.623345795809, 18.250000000000)
    pts(10) = Array(29.896867807448, 14.890846023781, 18.250000000000)
    pts(11) = Array(30.739550598048, 13.062925745409, 18.250000000000)
    pts(12) = Array(30.187343102274, 12.828262408665, 18.250000000000)
    pts(13) = Array(29.942816647908, 12.828030097789, 18.250000000000)
    pts(14) = Array(29.700022820122, 12.823070789957, 18.250000000000)
    pts(15) = Array(29.459060365677, 12.813250460808, 18.250000000000)
    pts(16) = Array(29.220040139464, 12.798408270098, 18.250000000000)
    pts(17) = Array(28.983088931165, 12.778348133001, 18.250000000000)
    pts(18) = Array(28.748355159745, 12.752826181250, 18.250000000000)
    pts(19) = Array(28.516017792297, 12.721531129129, 18.250000000000)
    FSData_302 = pts
End Function

Private Function FSData_303() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(27.203924855371, 14.864268312410, 15.150000000000)
    pts(1) = Array(26.887883868469, 15.428600100973, 15.150000000000)
    pts(2) = Array(27.053229590357, 15.592414621585, 15.150000000000)
    pts(3) = Array(27.215112297505, 15.761984730172, 15.150000000000)
    pts(4) = Array(27.373733039258, 15.936886913556, 15.150000000000)
    pts(5) = Array(27.529238456829, 16.116793415163, 15.150000000000)
    pts(6) = Array(27.681738605315, 16.301440819358, 15.150000000000)
    pts(7) = Array(27.831317595565, 16.490611295364, 15.150000000000)
    pts(8) = Array(27.978040391202, 16.684120620166, 15.150000000000)
    pts(9) = Array(28.121957378321, 16.881810128416, 15.150000000000)
    pts(10) = Array(28.636383427925, 17.190623728326, 15.150000000000)
    pts(11) = Array(29.619885480815, 15.434454447870, 15.150000000000)
    pts(12) = Array(29.087791729663, 15.157188799106, 15.150000000000)
    pts(13) = Array(28.844037295650, 15.137771879876, 15.150000000000)
    pts(14) = Array(28.602381023015, 15.113778475758, 15.150000000000)
    pts(15) = Array(28.362931869545, 15.085082723126, 15.150000000000)
    pts(16) = Array(28.125812967782, 15.051532975128, 15.150000000000)
    pts(17) = Array(27.891166101196, 15.012943699204, 15.150000000000)
    pts(18) = Array(27.659158364911, 14.969083423670, 15.150000000000)
    pts(19) = Array(27.429992597518, 14.919655863997, 15.150000000000)
    FSData_303 = pts
End Function

Private Function FSData_304() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(27.203924855371, 14.864268312410, 18.250000000000)
    pts(1) = Array(26.887883868469, 15.428600100973, 18.250000000000)
    pts(2) = Array(27.053229590357, 15.592414621585, 18.250000000000)
    pts(3) = Array(27.215112297505, 15.761984730172, 18.250000000000)
    pts(4) = Array(27.373733039258, 15.936886913556, 18.250000000000)
    pts(5) = Array(27.529238456829, 16.116793415163, 18.250000000000)
    pts(6) = Array(27.681738605315, 16.301440819358, 18.250000000000)
    pts(7) = Array(27.831317595565, 16.490611295364, 18.250000000000)
    pts(8) = Array(27.978040391202, 16.684120620166, 18.250000000000)
    pts(9) = Array(28.121957378321, 16.881810128416, 18.250000000000)
    pts(10) = Array(28.636383427925, 17.190623728326, 18.250000000000)
    pts(11) = Array(29.619885480815, 15.434454447870, 18.250000000000)
    pts(12) = Array(29.087791729663, 15.157188799106, 18.250000000000)
    pts(13) = Array(28.844037295650, 15.137771879876, 18.250000000000)
    pts(14) = Array(28.602381023015, 15.113778475758, 18.250000000000)
    pts(15) = Array(28.362931869545, 15.085082723126, 18.250000000000)
    pts(16) = Array(28.125812967782, 15.051532975128, 18.250000000000)
    pts(17) = Array(27.891166101196, 15.012943699204, 18.250000000000)
    pts(18) = Array(27.659158364911, 14.969083423670, 18.250000000000)
    pts(19) = Array(27.429992597518, 14.919655863997, 18.250000000000)
    FSData_304 = pts
End Function

Private Function FSData_305() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(25.953827183445, 16.952842078303, 15.150000000000)
    pts(1) = Array(25.594483483611, 17.490637930252, 15.150000000000)
    pts(2) = Array(25.746466760666, 17.666920341190, 15.150000000000)
    pts(3) = Array(25.894546120071, 17.848668892540, 15.150000000000)
    pts(4) = Array(26.038955219865, 18.035477150824, 15.150000000000)
    pts(5) = Array(26.179865964700, 18.227029875171, 15.150000000000)
    pts(6) = Array(26.317408737760, 18.423073096790, 15.150000000000)
    pts(7) = Array(26.451684481417, 18.623396255666, 15.150000000000)
    pts(8) = Array(26.582772413003, 18.827820793669, 15.150000000000)
    pts(9) = Array(26.710735212028, 19.036192487809, 15.150000000000)
    pts(10) = Array(27.199346222004, 19.384415521123, 15.150000000000)
    pts(11) = Array(28.317603919982, 17.710824606183, 15.150000000000)
    pts(12) = Array(27.808904448366, 17.392666080324, 15.150000000000)
    pts(13) = Array(27.567424861849, 17.354184264503, 15.150000000000)
    pts(14) = Array(27.328396035644, 17.311304691415, 15.150000000000)
    pts(15) = Array(27.091936466805, 17.263910434158, 15.150000000000)
    pts(16) = Array(26.858180806371, 17.211859974224, 15.150000000000)
    pts(17) = Array(26.627284957460, 17.154979475191, 15.150000000000)
    pts(18) = Array(26.399433661153, 17.093051289058, 15.150000000000)
    pts(19) = Array(26.174852376978, 17.025795959174, 15.150000000000)
    FSData_305 = pts
End Function

Private Function FSData_306() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(25.953827183445, 16.952842078303, 18.250000000000)
    pts(1) = Array(25.594483483611, 17.490637930252, 18.250000000000)
    pts(2) = Array(25.746466760666, 17.666920341190, 18.250000000000)
    pts(3) = Array(25.894546120071, 17.848668892540, 18.250000000000)
    pts(4) = Array(26.038955219865, 18.035477150824, 18.250000000000)
    pts(5) = Array(26.179865964700, 18.227029875171, 18.250000000000)
    pts(6) = Array(26.317408737760, 18.423073096790, 18.250000000000)
    pts(7) = Array(26.451684481417, 18.623396255666, 18.250000000000)
    pts(8) = Array(26.582772413003, 18.827820793669, 18.250000000000)
    pts(9) = Array(26.710735212028, 19.036192487809, 18.250000000000)
    pts(10) = Array(27.199346222004, 19.384415521123, 18.250000000000)
    pts(11) = Array(28.317603919982, 17.710824606183, 18.250000000000)
    pts(12) = Array(27.808904448366, 17.392666080324, 18.250000000000)
    pts(13) = Array(27.567424861849, 17.354184264503, 18.250000000000)
    pts(14) = Array(27.328396035644, 17.311304691415, 18.250000000000)
    pts(15) = Array(27.091936466805, 17.263910434158, 18.250000000000)
    pts(16) = Array(26.858180806371, 17.211859974224, 18.250000000000)
    pts(17) = Array(26.627284957460, 17.154979475191, 18.250000000000)
    pts(18) = Array(26.399433661153, 17.093051289058, 18.250000000000)
    pts(19) = Array(26.174852376978, 17.025795959174, 18.250000000000)
    FSData_306 = pts
End Function

Private Function FSData_307() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(24.543715536410, 18.936895935390, 15.150000000000)
    pts(1) = Array(24.143284597047, 19.444840160463, 15.150000000000)
    pts(2) = Array(24.280968401826, 19.632503622043, 15.150000000000)
    pts(3) = Array(24.414331454996, 19.825310075900, 15.150000000000)
    pts(4) = Array(24.543638582710, 20.022872674052, 15.150000000000)
    pts(5) = Array(24.669085893209, 20.224890634896, 15.150000000000)
    pts(6) = Array(24.790823293910, 20.431121002287, 15.150000000000)
    pts(7) = Array(24.908967936364, 20.641361785144, 15.150000000000)
    pts(8) = Array(25.023612803208, 20.855441191093, 15.150000000000)
    pts(9) = Array(25.134832480932, 21.073210390339, 15.150000000000)
    pts(10) = Array(25.594616001925, 21.458695946260, 15.150000000000)
    pts(11) = Array(26.840734914424, 19.878001641353, 15.150000000000)
    pts(12) = Array(26.358566023746, 19.520911791508, 15.150000000000)
    pts(13) = Array(26.120850086675, 19.463602332288, 15.150000000000)
    pts(14) = Array(25.885922399100, 19.402100957056, 15.150000000000)
    pts(15) = Array(25.653910266764, 19.336300396530, 15.150000000000)
    pts(16) = Array(25.424959029038, 19.266070133054, 15.150000000000)
    pts(17) = Array(25.199237747490, 19.191249098104, 15.150000000000)
    pts(18) = Array(24.976947670172, 19.111634809233, 15.150000000000)
    pts(19) = Array(24.758335487511, 19.026966360612, 15.150000000000)
    FSData_307 = pts
End Function

Private Function FSData_308() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(24.543715536410, 18.936895935390, 18.250000000000)
    pts(1) = Array(24.143284597047, 19.444840160463, 18.250000000000)
    pts(2) = Array(24.280968401826, 19.632503622043, 18.250000000000)
    pts(3) = Array(24.414331454996, 19.825310075900, 18.250000000000)
    pts(4) = Array(24.543638582710, 20.022872674052, 18.250000000000)
    pts(5) = Array(24.669085893209, 20.224890634896, 18.250000000000)
    pts(6) = Array(24.790823293910, 20.431121002287, 18.250000000000)
    pts(7) = Array(24.908967936364, 20.641361785144, 18.250000000000)
    pts(8) = Array(25.023612803208, 20.855441191093, 18.250000000000)
    pts(9) = Array(25.134832480932, 21.073210390339, 18.250000000000)
    pts(10) = Array(25.594616001925, 21.458695946260, 18.250000000000)
    pts(11) = Array(26.840734914424, 19.878001641353, 18.250000000000)
    pts(12) = Array(26.358566023746, 19.520911791508, 18.250000000000)
    pts(13) = Array(26.120850086675, 19.463602332288, 18.250000000000)
    pts(14) = Array(25.885922399100, 19.402100957056, 18.250000000000)
    pts(15) = Array(25.653910266764, 19.336300396530, 18.250000000000)
    pts(16) = Array(25.424959029038, 19.266070133054, 18.250000000000)
    pts(17) = Array(25.199237747490, 19.191249098104, 18.250000000000)
    pts(18) = Array(24.976947670172, 19.111634809233, 18.250000000000)
    pts(19) = Array(24.758335487511, 19.026966360612, 18.250000000000)
    FSData_308 = pts
End Function

Private Function FSData_309() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(22.982283721479, 20.804197531879, 15.150000000000)
    pts(1) = Array(22.543234332486, 21.279158485020, 15.150000000000)
    pts(2) = Array(22.665769798547, 21.477045989597, 15.150000000000)
    pts(3) = Array(22.783594317912, 21.679721630058, 15.150000000000)
    pts(4) = Array(22.897002252105, 21.886820528965, 15.150000000000)
    pts(5) = Array(23.006212703883, 22.098058218398, 15.150000000000)
    pts(6) = Array(23.111394180669, 22.313204252768, 15.150000000000)
    pts(7) = Array(23.212679320910, 22.532066455268, 15.150000000000)
    pts(8) = Array(23.310174299286, 22.754480858436, 15.150000000000)
    pts(9) = Array(23.403965149411, 22.980304943259, 15.150000000000)
    pts(10) = Array(23.832086463120, 23.400676375148, 15.150000000000)
    pts(11) = Array(25.198383852668, 21.922624186296, 15.150000000000)
    pts(12) = Array(24.745718274477, 21.528804590135, 15.150000000000)
    pts(13) = Array(24.513231584653, 21.453020819389, 15.150000000000)
    pts(14) = Array(24.283853443022, 21.373276818444, 15.150000000000)
    pts(15) = Array(24.057719179137, 21.289475636985, 15.150000000000)
    pts(16) = Array(23.834983924634, 21.201498562895, 15.150000000000)
    pts(17) = Array(23.615828857210, 21.109198288589, 15.150000000000)
    pts(18) = Array(23.400470491127, 21.012388745545, 15.150000000000)
    pts(19) = Array(23.189175226782, 20.910829187328, 15.150000000000)
    FSData_309 = pts
End Function

Private Function FSData_310() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(22.982283721479, 20.804197531879, 18.250000000000)
    pts(1) = Array(22.543234332486, 21.279158485020, 18.250000000000)
    pts(2) = Array(22.665769798547, 21.477045989597, 18.250000000000)
    pts(3) = Array(22.783594317912, 21.679721630058, 18.250000000000)
    pts(4) = Array(22.897002252105, 21.886820528965, 18.250000000000)
    pts(5) = Array(23.006212703883, 22.098058218398, 18.250000000000)
    pts(6) = Array(23.111394180669, 22.313204252768, 18.250000000000)
    pts(7) = Array(23.212679320910, 22.532066455268, 18.250000000000)
    pts(8) = Array(23.310174299286, 22.754480858436, 18.250000000000)
    pts(9) = Array(23.403965149411, 22.980304943259, 18.250000000000)
    pts(10) = Array(23.832086463120, 23.400676375148, 18.250000000000)
    pts(11) = Array(25.198383852668, 21.922624186296, 18.250000000000)
    pts(12) = Array(24.745718274477, 21.528804590135, 18.250000000000)
    pts(13) = Array(24.513231584653, 21.453020819389, 18.250000000000)
    pts(14) = Array(24.283853443022, 21.373276818444, 18.250000000000)
    pts(15) = Array(24.057719179137, 21.289475636985, 18.250000000000)
    pts(16) = Array(23.834983924634, 21.201498562895, 18.250000000000)
    pts(17) = Array(23.615828857210, 21.109198288589, 18.250000000000)
    pts(18) = Array(23.400470491127, 21.012388745545, 18.250000000000)
    pts(19) = Array(23.189175226782, 20.910829187328, 18.250000000000)
    FSData_310 = pts
End Function

Private Function FSData_311() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(21.279158485020, 22.543234332486, 15.150000000000)
    pts(1) = Array(20.804197531879, 22.982283721479, 15.150000000000)
    pts(2) = Array(20.910829187328, 23.189175226782, 15.150000000000)
    pts(3) = Array(21.012388745545, 23.400470491127, 15.150000000000)
    pts(4) = Array(21.109198288589, 23.615828857210, 15.150000000000)
    pts(5) = Array(21.201498562895, 23.834983924634, 15.150000000000)
    pts(6) = Array(21.289475636985, 24.057719179137, 15.150000000000)
    pts(7) = Array(21.373276818444, 24.283853443022, 15.150000000000)
    pts(8) = Array(21.453020819389, 24.513231584653, 15.150000000000)
    pts(9) = Array(21.528804590135, 24.745718274477, 15.150000000000)
    pts(10) = Array(21.922624186296, 25.198383852668, 15.150000000000)
    pts(11) = Array(23.400676375148, 23.832086463120, 15.150000000000)
    pts(12) = Array(22.980304943259, 23.403965149411, 15.150000000000)
    pts(13) = Array(22.754480858436, 23.310174299286, 15.150000000000)
    pts(14) = Array(22.532066455268, 23.212679320910, 15.150000000000)
    pts(15) = Array(22.313204252768, 23.111394180669, 15.150000000000)
    pts(16) = Array(22.098058218398, 23.006212703883, 15.150000000000)
    pts(17) = Array(21.886820528965, 22.897002252105, 15.150000000000)
    pts(18) = Array(21.679721630058, 22.783594317912, 15.150000000000)
    pts(19) = Array(21.477045989597, 22.665769798547, 15.150000000000)
    FSData_311 = pts
End Function

Private Function FSData_312() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(21.279158485020, 22.543234332486, 18.250000000000)
    pts(1) = Array(20.804197531879, 22.982283721479, 18.250000000000)
    pts(2) = Array(20.910829187328, 23.189175226782, 18.250000000000)
    pts(3) = Array(21.012388745545, 23.400470491127, 18.250000000000)
    pts(4) = Array(21.109198288589, 23.615828857210, 18.250000000000)
    pts(5) = Array(21.201498562895, 23.834983924634, 18.250000000000)
    pts(6) = Array(21.289475636985, 24.057719179137, 18.250000000000)
    pts(7) = Array(21.373276818444, 24.283853443022, 18.250000000000)
    pts(8) = Array(21.453020819389, 24.513231584653, 18.250000000000)
    pts(9) = Array(21.528804590135, 24.745718274477, 18.250000000000)
    pts(10) = Array(21.922624186296, 25.198383852668, 18.250000000000)
    pts(11) = Array(23.400676375148, 23.832086463120, 18.250000000000)
    pts(12) = Array(22.980304943259, 23.403965149411, 18.250000000000)
    pts(13) = Array(22.754480858436, 23.310174299286, 18.250000000000)
    pts(14) = Array(22.532066455268, 23.212679320910, 18.250000000000)
    pts(15) = Array(22.313204252768, 23.111394180669, 18.250000000000)
    pts(16) = Array(22.098058218398, 23.006212703883, 18.250000000000)
    pts(17) = Array(21.886820528965, 22.897002252105, 18.250000000000)
    pts(18) = Array(21.679721630058, 22.783594317912, 18.250000000000)
    pts(19) = Array(21.477045989597, 22.665769798547, 18.250000000000)
    FSData_312 = pts
End Function

Private Function FSData_313() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(19.444840160463, 24.143284597047, 15.150000000000)
    pts(1) = Array(18.936895935390, 24.543715536410, 15.150000000000)
    pts(2) = Array(19.026966360612, 24.758335487511, 15.150000000000)
    pts(3) = Array(19.111634809233, 24.976947670172, 15.150000000000)
    pts(4) = Array(19.191249098104, 25.199237747490, 15.150000000000)
    pts(5) = Array(19.266070133054, 25.424959029038, 15.150000000000)
    pts(6) = Array(19.336300396530, 25.653910266764, 15.150000000000)
    pts(7) = Array(19.402100957056, 25.885922399100, 15.150000000000)
    pts(8) = Array(19.463602332288, 26.120850086675, 15.150000000000)
    pts(9) = Array(19.520911791508, 26.358566023746, 15.150000000000)
    pts(10) = Array(19.878001641353, 26.840734914424, 15.150000000000)
    pts(11) = Array(21.458695946260, 25.594616001925, 15.150000000000)
    pts(12) = Array(21.073210390339, 25.134832480932, 15.150000000000)
    pts(13) = Array(20.855441191093, 25.023612803208, 15.150000000000)
    pts(14) = Array(20.641361785144, 24.908967936364, 15.150000000000)
    pts(15) = Array(20.431121002287, 24.790823293910, 15.150000000000)
    pts(16) = Array(20.224890634896, 24.669085893209, 15.150000000000)
    pts(17) = Array(20.022872674052, 24.543638582710, 15.150000000000)
    pts(18) = Array(19.825310075900, 24.414331454996, 15.150000000000)
    pts(19) = Array(19.632503622043, 24.280968401826, 15.150000000000)
    FSData_313 = pts
End Function

Private Function FSData_314() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(19.444840160463, 24.143284597047, 18.250000000000)
    pts(1) = Array(18.936895935390, 24.543715536410, 18.250000000000)
    pts(2) = Array(19.026966360612, 24.758335487511, 18.250000000000)
    pts(3) = Array(19.111634809233, 24.976947670172, 18.250000000000)
    pts(4) = Array(19.191249098104, 25.199237747490, 18.250000000000)
    pts(5) = Array(19.266070133054, 25.424959029038, 18.250000000000)
    pts(6) = Array(19.336300396530, 25.653910266764, 18.250000000000)
    pts(7) = Array(19.402100957056, 25.885922399100, 18.250000000000)
    pts(8) = Array(19.463602332288, 26.120850086675, 18.250000000000)
    pts(9) = Array(19.520911791508, 26.358566023746, 18.250000000000)
    pts(10) = Array(19.878001641353, 26.840734914424, 18.250000000000)
    pts(11) = Array(21.458695946260, 25.594616001925, 18.250000000000)
    pts(12) = Array(21.073210390339, 25.134832480932, 18.250000000000)
    pts(13) = Array(20.855441191093, 25.023612803208, 18.250000000000)
    pts(14) = Array(20.641361785144, 24.908967936364, 18.250000000000)
    pts(15) = Array(20.431121002287, 24.790823293910, 18.250000000000)
    pts(16) = Array(20.224890634896, 24.669085893209, 18.250000000000)
    pts(17) = Array(20.022872674052, 24.543638582710, 18.250000000000)
    pts(18) = Array(19.825310075900, 24.414331454996, 18.250000000000)
    pts(19) = Array(19.632503622043, 24.280968401826, 18.250000000000)
    FSData_314 = pts
End Function

Private Function FSData_315() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(17.490637930252, 25.594483483611, 15.150000000000)
    pts(1) = Array(16.952842078303, 25.953827183445, 15.150000000000)
    pts(2) = Array(17.025795959174, 26.174852376978, 15.150000000000)
    pts(3) = Array(17.093051289058, 26.399433661153, 15.150000000000)
    pts(4) = Array(17.154979475191, 26.627284957460, 15.150000000000)
    pts(5) = Array(17.211859974224, 26.858180806371, 15.150000000000)
    pts(6) = Array(17.263910434158, 27.091936466805, 15.150000000000)
    pts(7) = Array(17.311304691415, 27.328396035644, 15.150000000000)
    pts(8) = Array(17.354184264503, 27.567424861849, 15.150000000000)
    pts(9) = Array(17.392666080324, 27.808904448366, 15.150000000000)
    pts(10) = Array(17.710824606183, 28.317603919982, 15.150000000000)
    pts(11) = Array(19.384415521123, 27.199346222004, 15.150000000000)
    pts(12) = Array(19.036192487809, 26.710735212028, 15.150000000000)
    pts(13) = Array(18.827820793669, 26.582772413003, 15.150000000000)
    pts(14) = Array(18.623396255666, 26.451684481417, 15.150000000000)
    pts(15) = Array(18.423073096790, 26.317408737760, 15.150000000000)
    pts(16) = Array(18.227029875171, 26.179865964700, 15.150000000000)
    pts(17) = Array(18.035477150824, 26.038955219865, 15.150000000000)
    pts(18) = Array(17.848668892540, 25.894546120071, 15.150000000000)
    pts(19) = Array(17.666920341190, 25.746466760666, 15.150000000000)
    FSData_315 = pts
End Function

Private Function FSData_316() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(17.490637930252, 25.594483483611, 18.250000000000)
    pts(1) = Array(16.952842078303, 25.953827183445, 18.250000000000)
    pts(2) = Array(17.025795959174, 26.174852376978, 18.250000000000)
    pts(3) = Array(17.093051289058, 26.399433661153, 18.250000000000)
    pts(4) = Array(17.154979475191, 26.627284957460, 18.250000000000)
    pts(5) = Array(17.211859974224, 26.858180806371, 18.250000000000)
    pts(6) = Array(17.263910434158, 27.091936466805, 18.250000000000)
    pts(7) = Array(17.311304691415, 27.328396035644, 18.250000000000)
    pts(8) = Array(17.354184264503, 27.567424861849, 18.250000000000)
    pts(9) = Array(17.392666080324, 27.808904448366, 18.250000000000)
    pts(10) = Array(17.710824606183, 28.317603919982, 18.250000000000)
    pts(11) = Array(19.384415521123, 27.199346222004, 18.250000000000)
    pts(12) = Array(19.036192487809, 26.710735212028, 18.250000000000)
    pts(13) = Array(18.827820793669, 26.582772413003, 18.250000000000)
    pts(14) = Array(18.623396255666, 26.451684481417, 18.250000000000)
    pts(15) = Array(18.423073096790, 26.317408737760, 18.250000000000)
    pts(16) = Array(18.227029875171, 26.179865964700, 18.250000000000)
    pts(17) = Array(18.035477150824, 26.038955219865, 18.250000000000)
    pts(18) = Array(17.848668892540, 25.894546120071, 18.250000000000)
    pts(19) = Array(17.666920341190, 25.746466760666, 18.250000000000)
    FSData_316 = pts
End Function

Private Function FSData_317() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(15.428600100973, 26.887883868469, 15.150000000000)
    pts(1) = Array(14.864268312410, 27.203924855371, 15.150000000000)
    pts(2) = Array(14.919655863997, 27.429992597518, 15.150000000000)
    pts(3) = Array(14.969083423670, 27.659158364911, 15.150000000000)
    pts(4) = Array(15.012943699204, 27.891166101196, 15.150000000000)
    pts(5) = Array(15.051532975128, 28.125812967782, 15.150000000000)
    pts(6) = Array(15.085082723126, 28.362931869545, 15.150000000000)
    pts(7) = Array(15.113778475758, 28.602381023015, 15.150000000000)
    pts(8) = Array(15.137771879876, 28.844037295650, 15.150000000000)
    pts(9) = Array(15.157188799106, 29.087791729663, 15.150000000000)
    pts(10) = Array(15.434454447870, 29.619885480815, 15.150000000000)
    pts(11) = Array(17.190623728326, 28.636383427925, 15.150000000000)
    pts(12) = Array(16.881810128416, 28.121957378321, 15.150000000000)
    pts(13) = Array(16.684120620166, 27.978040391202, 15.150000000000)
    pts(14) = Array(16.490611295364, 27.831317595565, 15.150000000000)
    pts(15) = Array(16.301440819358, 27.681738605315, 15.150000000000)
    pts(16) = Array(16.116793415163, 27.529238456829, 15.150000000000)
    pts(17) = Array(15.936886913556, 27.373733039258, 15.150000000000)
    pts(18) = Array(15.761984730172, 27.215112297505, 15.150000000000)
    pts(19) = Array(15.592414621585, 27.053229590357, 15.150000000000)
    FSData_317 = pts
End Function

Private Function FSData_318() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(15.428600100973, 26.887883868469, 18.250000000000)
    pts(1) = Array(14.864268312410, 27.203924855371, 18.250000000000)
    pts(2) = Array(14.919655863997, 27.429992597518, 18.250000000000)
    pts(3) = Array(14.969083423670, 27.659158364911, 18.250000000000)
    pts(4) = Array(15.012943699204, 27.891166101196, 18.250000000000)
    pts(5) = Array(15.051532975128, 28.125812967782, 18.250000000000)
    pts(6) = Array(15.085082723126, 28.362931869545, 18.250000000000)
    pts(7) = Array(15.113778475758, 28.602381023015, 18.250000000000)
    pts(8) = Array(15.137771879876, 28.844037295650, 18.250000000000)
    pts(9) = Array(15.157188799106, 29.087791729663, 18.250000000000)
    pts(10) = Array(15.434454447870, 29.619885480815, 18.250000000000)
    pts(11) = Array(17.190623728326, 28.636383427925, 18.250000000000)
    pts(12) = Array(16.881810128416, 28.121957378321, 18.250000000000)
    pts(13) = Array(16.684120620166, 27.978040391202, 18.250000000000)
    pts(14) = Array(16.490611295364, 27.831317595565, 18.250000000000)
    pts(15) = Array(16.301440819358, 27.681738605315, 18.250000000000)
    pts(16) = Array(16.116793415163, 27.529238456829, 18.250000000000)
    pts(17) = Array(15.936886913556, 27.373733039258, 18.250000000000)
    pts(18) = Array(15.761984730172, 27.215112297505, 18.250000000000)
    pts(19) = Array(15.592414621585, 27.053229590357, 18.250000000000)
    FSData_318 = pts
End Function

Private Function FSData_319() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(13.271439821541, 28.015511508149, 15.150000000000)
    pts(1) = Array(12.684051389501, 28.286301284341, 15.150000000000)
    pts(2) = Array(12.721531129129, 28.516017792297, 15.150000000000)
    pts(3) = Array(12.752826181250, 28.748355159745, 15.150000000000)
    pts(4) = Array(12.778348133001, 28.983088931165, 15.150000000000)
    pts(5) = Array(12.798408270098, 29.220040139464, 15.150000000000)
    pts(6) = Array(12.813250460808, 29.459060365677, 15.150000000000)
    pts(7) = Array(12.823070789957, 29.700022820122, 15.150000000000)
    pts(8) = Array(12.828030097789, 29.942816647908, 15.150000000000)
    pts(9) = Array(12.828262408665, 30.187343102274, 15.150000000000)
    pts(10) = Array(13.062925745409, 30.739550598048, 15.150000000000)
    pts(11) = Array(14.890846023781, 29.896867807448, 15.150000000000)
    pts(12) = Array(14.623345795809, 29.359798325877, 15.150000000000)
    pts(13) = Array(14.437557295007, 29.200814446748, 15.150000000000)
    pts(14) = Array(14.256156232741, 29.039361381885, 15.150000000000)
    pts(15) = Array(14.079304738495, 28.875401349256, 15.150000000000)
    pts(16) = Array(13.907191564372, 28.708884039472, 15.150000000000)
    pts(17) = Array(13.740040468913, 28.539742691774, 15.150000000000)
    pts(18) = Array(13.578122690552, 28.367888257682, 15.150000000000)
    pts(19) = Array(13.421776480835, 28.193200263512, 15.150000000000)
    FSData_319 = pts
End Function

Private Function FSData_320() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(13.271439821541, 28.015511508149, 18.250000000000)
    pts(1) = Array(12.684051389501, 28.286301284341, 18.250000000000)
    pts(2) = Array(12.721531129129, 28.516017792297, 18.250000000000)
    pts(3) = Array(12.752826181250, 28.748355159745, 18.250000000000)
    pts(4) = Array(12.778348133001, 28.983088931165, 18.250000000000)
    pts(5) = Array(12.798408270098, 29.220040139464, 18.250000000000)
    pts(6) = Array(12.813250460808, 29.459060365677, 18.250000000000)
    pts(7) = Array(12.823070789957, 29.700022820122, 18.250000000000)
    pts(8) = Array(12.828030097789, 29.942816647908, 18.250000000000)
    pts(9) = Array(12.828262408665, 30.187343102274, 18.250000000000)
    pts(10) = Array(13.062925745409, 30.739550598048, 18.250000000000)
    pts(11) = Array(14.890846023781, 29.896867807448, 18.250000000000)
    pts(12) = Array(14.623345795809, 29.359798325877, 18.250000000000)
    pts(13) = Array(14.437557295007, 29.200814446748, 18.250000000000)
    pts(14) = Array(14.256156232741, 29.039361381885, 18.250000000000)
    pts(15) = Array(14.079304738495, 28.875401349256, 18.250000000000)
    pts(16) = Array(13.907191564372, 28.708884039472, 18.250000000000)
    pts(17) = Array(13.740040468913, 28.539742691774, 18.250000000000)
    pts(18) = Array(13.578122690552, 28.367888257682, 18.250000000000)
    pts(19) = Array(13.421776480835, 28.193200263512, 18.250000000000)
    FSData_320 = pts
End Function

Private Function FSData_321() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(11.032456702408, 28.970414203278, 15.150000000000)
    pts(1) = Array(10.425633071901, 29.194283259743, 15.150000000000)
    pts(2) = Array(10.444973924511, 29.426232254848, 15.150000000000)
    pts(3) = Array(10.457943524677, 29.660308785221, 15.150000000000)
    pts(4) = Array(10.464969801327, 29.896321380018, 15.150000000000)
    pts(5) = Array(10.466377122181, 30.134116047037, 15.150000000000)
    pts(6) = Array(10.462420248560, 30.373563958524, 15.150000000000)
    pts(7) = Array(10.453304608632, 30.614554100283, 15.150000000000)
    pts(8) = Array(10.439199244397, 30.856988578534, 15.150000000000)
    pts(9) = Array(10.420245514645, 31.100779466350, 15.150000000000)
    pts(10) = Array(10.610859761864, 31.669696163905, 15.150000000000)
    pts(11) = Array(12.499261301791, 30.973028055222, 15.150000000000)
    pts(12) = Array(12.274723673615, 30.416626353631, 15.150000000000)
    pts(13) = Array(12.101981628149, 30.243555770973, 15.150000000000)
    pts(14) = Array(11.933807226290, 30.068367848717, 15.150000000000)
    pts(15) = Array(11.770365062076, 29.891037641833, 15.150000000000)
    pts(16) = Array(11.611847252978, 29.711529805339, 15.150000000000)
    pts(17) = Array(11.458482105752, 29.529795340167, 15.150000000000)
    pts(18) = Array(11.310547009360, 29.345766753470, 15.150000000000)
    pts(19) = Array(11.168388624887, 29.159350481852, 15.150000000000)
    FSData_321 = pts
End Function

Private Function FSData_322() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(11.032456702408, 28.970414203278, 18.250000000000)
    pts(1) = Array(10.425633071901, 29.194283259743, 18.250000000000)
    pts(2) = Array(10.444973924511, 29.426232254848, 18.250000000000)
    pts(3) = Array(10.457943524677, 29.660308785221, 18.250000000000)
    pts(4) = Array(10.464969801327, 29.896321380018, 18.250000000000)
    pts(5) = Array(10.466377122181, 30.134116047037, 18.250000000000)
    pts(6) = Array(10.462420248560, 30.373563958524, 18.250000000000)
    pts(7) = Array(10.453304608632, 30.614554100283, 18.250000000000)
    pts(8) = Array(10.439199244397, 30.856988578534, 18.250000000000)
    pts(9) = Array(10.420245514645, 31.100779466350, 18.250000000000)
    pts(10) = Array(10.610859761864, 31.669696163905, 18.250000000000)
    pts(11) = Array(12.499261301791, 30.973028055222, 18.250000000000)
    pts(12) = Array(12.274723673615, 30.416626353631, 18.250000000000)
    pts(13) = Array(12.101981628149, 30.243555770973, 18.250000000000)
    pts(14) = Array(11.933807226290, 30.068367848717, 18.250000000000)
    pts(15) = Array(11.770365062076, 29.891037641833, 18.250000000000)
    pts(16) = Array(11.611847252978, 29.711529805339, 18.250000000000)
    pts(17) = Array(11.458482105752, 29.529795340167, 18.250000000000)
    pts(18) = Array(11.310547009360, 29.345766753470, 18.250000000000)
    pts(19) = Array(11.168388624887, 29.159350481852, 18.250000000000)
    FSData_322 = pts
End Function

Private Function FSData_323() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(8.725454819040, 29.746704661204, 15.150000000000)
    pts(1) = Array(8.102937259537, 29.922272770764, 15.150000000000)
    pts(2) = Array(8.104019982343, 30.155024210333, 15.150000000000)
    pts(3) = Array(8.098584168655, 30.389396743983, 15.150000000000)
    pts(4) = Array(8.087071450871, 30.625233066028, 15.150000000000)
    pts(5) = Array(8.069817278881, 30.862405108570, 15.150000000000)
    pts(6) = Array(8.047085736372, 31.100804429331, 15.150000000000)
    pts(7) = Array(8.019090328318, 31.340336474044, 15.150000000000)
    pts(8) = Array(7.986007256277, 31.580916913584, 15.150000000000)
    pts(9) = Array(7.947984341944, 31.822469182957, 15.150000000000)
    pts(10) = Array(8.093374299418, 32.404587521669, 15.150000000000)
    pts(11) = Array(10.030614477449, 31.858229285457, 15.150000000000)
    pts(12) = Array(9.850423798214, 31.285925765359, 15.150000000000)
    pts(13) = Array(9.691793220237, 31.099835516867, 15.150000000000)
    pts(14) = Array(9.537882329896, 30.911992829015, 15.150000000000)
    pts(15) = Array(9.388857171005, 30.722385747570, 15.150000000000)
    pts(16) = Array(9.244912041937, 30.530994109869, 15.150000000000)
    pts(17) = Array(9.106278390077, 30.337786980636, 15.150000000000)
    pts(18) = Array(8.973238044717, 30.142718838765, 15.150000000000)
    pts(19) = Array(8.846143939200, 29.945723608004, 15.150000000000)
    FSData_323 = pts
End Function

Private Function FSData_324() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(8.725454819040, 29.746704661204, 18.250000000000)
    pts(1) = Array(8.102937259537, 29.922272770764, 18.250000000000)
    pts(2) = Array(8.104019982343, 30.155024210333, 18.250000000000)
    pts(3) = Array(8.098584168655, 30.389396743983, 18.250000000000)
    pts(4) = Array(8.087071450871, 30.625233066028, 18.250000000000)
    pts(5) = Array(8.069817278881, 30.862405108570, 18.250000000000)
    pts(6) = Array(8.047085736372, 31.100804429331, 18.250000000000)
    pts(7) = Array(8.019090328318, 31.340336474044, 18.250000000000)
    pts(8) = Array(7.986007256277, 31.580916913584, 18.250000000000)
    pts(9) = Array(7.947984341944, 31.822469182957, 18.250000000000)
    pts(10) = Array(8.093374299418, 32.404587521669, 18.250000000000)
    pts(11) = Array(10.030614477449, 31.858229285457, 18.250000000000)
    pts(12) = Array(9.850423798214, 31.285925765359, 18.250000000000)
    pts(13) = Array(9.691793220237, 31.099835516867, 18.250000000000)
    pts(14) = Array(9.537882329896, 30.911992829015, 18.250000000000)
    pts(15) = Array(9.388857171005, 30.722385747570, 18.250000000000)
    pts(16) = Array(9.244912041937, 30.530994109869, 18.250000000000)
    pts(17) = Array(9.106278390077, 30.337786980636, 18.250000000000)
    pts(18) = Array(8.973238044717, 30.142718838765, 18.250000000000)
    pts(19) = Array(8.846143939200, 29.945723608004, 18.250000000000)
    FSData_324 = pts
End Function

Private Function FSData_325() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(6.364657605205, 30.339596793110, 15.150000000000)
    pts(1) = Array(5.730284144469, 30.465781519988, 15.150000000000)
    pts(2) = Array(5.713102062123, 30.697900413998, 15.150000000000)
    pts(3) = Array(5.689294348182, 30.931123966318, 15.150000000000)
    pts(4) = Array(5.659313615697, 31.165330006262, 15.150000000000)
    pts(5) = Array(5.623504328571, 31.400417179817, 15.150000000000)
    pts(6) = Array(5.582138264692, 31.636298098764, 15.150000000000)
    pts(7) = Array(5.535435689512, 31.872895251723, 15.150000000000)
    pts(8) = Array(5.483578877805, 32.110138394141, 15.150000000000)
    pts(9) = Array(5.426721202801, 32.347962795006, 15.150000000000)
    pts(10) = Array(5.525990493096, 32.939693821744, 15.150000000000)
    pts(11) = Array(7.500125579336, 32.547013938212, 15.150000000000)
    pts(12) = Array(7.365392784498, 31.962337041119, 15.150000000000)
    pts(13) = Array(7.221851684273, 31.764374435055, 15.150000000000)
    pts(14) = Array(7.083153217270, 31.565035094241, 15.150000000000)
    pts(15) = Array(6.949463853363, 31.364320128943, 15.150000000000)
    pts(16) = Array(6.820978873913, 31.162224683126, 15.150000000000)
    pts(17) = Array(6.697931439982, 30.958736076030, 15.150000000000)
    pts(18) = Array(6.580606083624, 30.753831038948, 15.150000000000)
    pts(19) = Array(6.469359834487, 30.547471390148, 15.150000000000)
    FSData_325 = pts
End Function

Private Function FSData_326() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(6.364657605205, 30.339596793110, 18.250000000000)
    pts(1) = Array(5.730284144469, 30.465781519988, 18.250000000000)
    pts(2) = Array(5.713102062123, 30.697900413998, 18.250000000000)
    pts(3) = Array(5.689294348182, 30.931123966318, 18.250000000000)
    pts(4) = Array(5.659313615697, 31.165330006262, 18.250000000000)
    pts(5) = Array(5.623504328571, 31.400417179817, 18.250000000000)
    pts(6) = Array(5.582138264692, 31.636298098764, 18.250000000000)
    pts(7) = Array(5.535435689512, 31.872895251723, 18.250000000000)
    pts(8) = Array(5.483578877805, 32.110138394141, 18.250000000000)
    pts(9) = Array(5.426721202801, 32.347962795006, 18.250000000000)
    pts(10) = Array(5.525990493096, 32.939693821744, 18.250000000000)
    pts(11) = Array(7.500125579336, 32.547013938212, 18.250000000000)
    pts(12) = Array(7.365392784498, 31.962337041119, 18.250000000000)
    pts(13) = Array(7.221851684273, 31.764374435055, 18.250000000000)
    pts(14) = Array(7.083153217270, 31.565035094241, 18.250000000000)
    pts(15) = Array(6.949463853363, 31.364320128943, 18.250000000000)
    pts(16) = Array(6.820978873913, 31.162224683126, 18.250000000000)
    pts(17) = Array(6.697931439982, 30.958736076030, 18.250000000000)
    pts(18) = Array(6.580606083624, 30.753831038948, 18.250000000000)
    pts(19) = Array(6.469359834487, 30.547471390148, 18.250000000000)
    FSData_326 = pts
End Function

Private Function FSData_327() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(3.964620160771, 30.745435221847, 15.150000000000)
    pts(1) = Array(3.322301922137, 30.821458595241, 15.150000000000)
    pts(2) = Array(3.286960967891, 31.051513853523, 15.150000000000)
    pts(3) = Array(3.244928136169, 31.282150523759, 15.150000000000)
    pts(4) = Array(3.196664230169, 31.513282323483, 15.150000000000)
    pts(5) = Array(3.142520604070, 31.744835237452, 15.150000000000)
    pts(6) = Array(3.082775054361, 31.976743470282, 15.150000000000)
    pts(7) = Array(3.017653248961, 32.208947031361, 15.150000000000)
    pts(8) = Array(2.947342412076, 32.441390193794, 15.150000000000)
    pts(9) = Array(2.872000522875, 32.674020459634, 15.150000000000)
    pts(10) = Array(2.924537117806, 33.271715955847, 15.150000000000)
    pts(11) = Array(4.923395912982, 33.035135427058, 15.150000000000)
    pts(12) = Array(4.834951675024, 32.441689880464, 15.150000000000)
    pts(13) = Array(4.707385031166, 32.233075422124, 15.150000000000)
    pts(14) = Array(4.584754109672, 32.023468421672, 15.150000000000)
    pts(15) = Array(4.467224780134, 31.812883047026, 15.150000000000)
    pts(16) = Array(4.354992102925, 31.601329778721, 15.150000000000)
    pts(17) = Array(4.248289515272, 31.388814268692, 15.150000000000)
    pts(18) = Array(4.147402497751, 31.175335644090, 15.150000000000)
    pts(19) = Array(4.052689975114, 30.960883853107, 15.150000000000)
    FSData_327 = pts
End Function

Private Function FSData_328() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(3.964620160771, 30.745435221847, 18.250000000000)
    pts(1) = Array(3.322301922137, 30.821458595241, 18.250000000000)
    pts(2) = Array(3.286960967891, 31.051513853523, 18.250000000000)
    pts(3) = Array(3.244928136169, 31.282150523759, 18.250000000000)
    pts(4) = Array(3.196664230169, 31.513282323483, 18.250000000000)
    pts(5) = Array(3.142520604070, 31.744835237452, 18.250000000000)
    pts(6) = Array(3.082775054361, 31.976743470282, 18.250000000000)
    pts(7) = Array(3.017653248961, 32.208947031361, 18.250000000000)
    pts(8) = Array(2.947342412076, 32.441390193794, 18.250000000000)
    pts(9) = Array(2.872000522875, 32.674020459634, 18.250000000000)
    pts(10) = Array(2.924537117806, 33.271715955847, 18.250000000000)
    pts(11) = Array(4.923395912982, 33.035135427058, 18.250000000000)
    pts(12) = Array(4.834951675024, 32.441689880464, 18.250000000000)
    pts(13) = Array(4.707385031166, 32.233075422124, 18.250000000000)
    pts(14) = Array(4.584754109672, 32.023468421672, 18.250000000000)
    pts(15) = Array(4.467224780134, 31.812883047026, 18.250000000000)
    pts(16) = Array(4.354992102925, 31.601329778721, 18.250000000000)
    pts(17) = Array(4.248289515272, 31.388814268692, 18.250000000000)
    pts(18) = Array(4.147402497751, 31.175335644090, 18.250000000000)
    pts(19) = Array(4.052689975114, 30.960883853107, 18.250000000000)
    FSData_328 = pts
End Function

Private Function FSData_329() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(1.540139514675, 30.961717818547, 15.150000000000)
    pts(1) = Array(0.893836603677, 30.987111129080, 15.150000000000)
    pts(2) = Array(0.840554666267, 31.213684384465, 15.150000000000)
    pts(3) = Array(0.780555863149, 31.440312220850, 15.150000000000)
    pts(4) = Array(0.714306346663, 31.666944775951, 15.150000000000)
    pts(5) = Array(0.642162195051, 31.893535829620, 15.150000000000)
    pts(6) = Array(0.564405510693, 32.120041584959, 15.150000000000)
    pts(7) = Array(0.481265972658, 32.346419941990, 15.150000000000)
    pts(8) = Array(0.392934600286, 32.572630035045, 15.150000000000)
    pts(9) = Array(0.299573004689, 32.798631922915, 15.150000000000)
    pts(10) = Array(0.305052998677, 33.398606897115, 15.150000000000)
    pts(11) = Array(2.316311873629, 33.319584320698, 15.150000000000)
    pts(12) = Array(2.274701480689, 32.721028913739, 15.150000000000)
    pts(13) = Array(2.163895783978, 32.503048780631, 15.150000000000)
    pts(14) = Array(2.058088468403, 32.284466417400, 15.150000000000)
    pts(15) = Array(1.957443780632, 32.065308962267, 15.150000000000)
    pts(16) = Array(1.862155357441, 31.845602167721, 15.150000000000)
    pts(17) = Array(1.772455473001, 31.625369983547, 15.150000000000)
    pts(18) = Array(1.688628796329, 31.404633938134, 15.150000000000)
    pts(19) = Array(1.611033934388, 31.183412171574, 15.150000000000)
    FSData_329 = pts
End Function

Private Function FSData_330() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(1.540139514675, 30.961717818547, 18.250000000000)
    pts(1) = Array(0.893836603677, 30.987111129080, 18.250000000000)
    pts(2) = Array(0.840554666267, 31.213684384465, 18.250000000000)
    pts(3) = Array(0.780555863149, 31.440312220850, 18.250000000000)
    pts(4) = Array(0.714306346663, 31.666944775951, 18.250000000000)
    pts(5) = Array(0.642162195051, 31.893535829620, 18.250000000000)
    pts(6) = Array(0.564405510693, 32.120041584959, 18.250000000000)
    pts(7) = Array(0.481265972658, 32.346419941990, 18.250000000000)
    pts(8) = Array(0.392934600286, 32.572630035045, 18.250000000000)
    pts(9) = Array(0.299573004689, 32.798631922915, 18.250000000000)
    pts(10) = Array(0.305052998677, 33.398606897115, 18.250000000000)
    pts(11) = Array(2.316311873629, 33.319584320698, 18.250000000000)
    pts(12) = Array(2.274701480689, 32.721028913739, 18.250000000000)
    pts(13) = Array(2.163895783978, 32.503048780631, 18.250000000000)
    pts(14) = Array(2.058088468403, 32.284466417400, 18.250000000000)
    pts(15) = Array(1.957443780632, 32.065308962267, 18.250000000000)
    pts(16) = Array(1.862155357441, 31.845602167721, 18.250000000000)
    pts(17) = Array(1.772455473001, 31.625369983547, 18.250000000000)
    pts(18) = Array(1.688628796329, 31.404633938134, 18.250000000000)
    pts(19) = Array(1.611033934388, 31.183412171574, 18.250000000000)
    FSData_330 = pts
End Function

Private Function FSData_331() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-0.893836603677, 30.987111129080, 15.150000000000)
    pts(1) = Array(-1.540139514675, 30.961717818547, 15.150000000000)
    pts(2) = Array(-1.611033934388, 31.183412171574, 15.150000000000)
    pts(3) = Array(-1.688628796329, 31.404633938134, 15.150000000000)
    pts(4) = Array(-1.772455473001, 31.625369983547, 15.150000000000)
    pts(5) = Array(-1.862155357441, 31.845602167721, 15.150000000000)
    pts(6) = Array(-1.957443780632, 32.065308962267, 15.150000000000)
    pts(7) = Array(-2.058088468403, 32.284466417400, 15.150000000000)
    pts(8) = Array(-2.163895783978, 32.503048780631, 15.150000000000)
    pts(9) = Array(-2.274701480689, 32.721028913739, 15.150000000000)
    pts(10) = Array(-2.316311873629, 33.319584320698, 15.150000000000)
    pts(11) = Array(-0.305052998677, 33.398606897115, 15.150000000000)
    pts(12) = Array(-0.299573004689, 32.798631922915, 15.150000000000)
    pts(13) = Array(-0.392934600286, 32.572630035045, 15.150000000000)
    pts(14) = Array(-0.481265972658, 32.346419941990, 15.150000000000)
    pts(15) = Array(-0.564405510693, 32.120041584959, 15.150000000000)
    pts(16) = Array(-0.642162195051, 31.893535829620, 15.150000000000)
    pts(17) = Array(-0.714306346663, 31.666944775951, 15.150000000000)
    pts(18) = Array(-0.780555863149, 31.440312220850, 15.150000000000)
    pts(19) = Array(-0.840554666267, 31.213684384465, 15.150000000000)
    FSData_331 = pts
End Function

Private Function FSData_332() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-0.893836603677, 30.987111129080, 18.250000000000)
    pts(1) = Array(-1.540139514675, 30.961717818547, 18.250000000000)
    pts(2) = Array(-1.611033934388, 31.183412171574, 18.250000000000)
    pts(3) = Array(-1.688628796329, 31.404633938134, 18.250000000000)
    pts(4) = Array(-1.772455473001, 31.625369983547, 18.250000000000)
    pts(5) = Array(-1.862155357441, 31.845602167721, 18.250000000000)
    pts(6) = Array(-1.957443780632, 32.065308962267, 18.250000000000)
    pts(7) = Array(-2.058088468403, 32.284466417400, 18.250000000000)
    pts(8) = Array(-2.163895783978, 32.503048780631, 18.250000000000)
    pts(9) = Array(-2.274701480689, 32.721028913739, 18.250000000000)
    pts(10) = Array(-2.316311873629, 33.319584320698, 18.250000000000)
    pts(11) = Array(-0.305052998677, 33.398606897115, 18.250000000000)
    pts(12) = Array(-0.299573004689, 32.798631922915, 18.250000000000)
    pts(13) = Array(-0.392934600286, 32.572630035045, 18.250000000000)
    pts(14) = Array(-0.481265972658, 32.346419941990, 18.250000000000)
    pts(15) = Array(-0.564405510693, 32.120041584959, 18.250000000000)
    pts(16) = Array(-0.642162195051, 31.893535829620, 18.250000000000)
    pts(17) = Array(-0.714306346663, 31.666944775951, 18.250000000000)
    pts(18) = Array(-0.780555863149, 31.440312220850, 18.250000000000)
    pts(19) = Array(-0.840554666267, 31.213684384465, 18.250000000000)
    FSData_332 = pts
End Function

Private Function FSData_333() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-3.322301922137, 30.821458595241, 15.150000000000)
    pts(1) = Array(-3.964620160771, 30.745435221847, 15.150000000000)
    pts(2) = Array(-4.052689975114, 30.960883853107, 15.150000000000)
    pts(3) = Array(-4.147402497751, 31.175335644090, 15.150000000000)
    pts(4) = Array(-4.248289515272, 31.388814268692, 15.150000000000)
    pts(5) = Array(-4.354992102925, 31.601329778721, 15.150000000000)
    pts(6) = Array(-4.467224780134, 31.812883047026, 15.150000000000)
    pts(7) = Array(-4.584754109672, 32.023468421672, 15.150000000000)
    pts(8) = Array(-4.707385031166, 32.233075422124, 15.150000000000)
    pts(9) = Array(-4.834951675024, 32.441689880464, 15.150000000000)
    pts(10) = Array(-4.923395912982, 33.035135427058, 15.150000000000)
    pts(11) = Array(-2.924537117806, 33.271715955847, 15.150000000000)
    pts(12) = Array(-2.872000522875, 32.674020459634, 15.150000000000)
    pts(13) = Array(-2.947342412076, 32.441390193794, 15.150000000000)
    pts(14) = Array(-3.017653248961, 32.208947031361, 15.150000000000)
    pts(15) = Array(-3.082775054361, 31.976743470282, 15.150000000000)
    pts(16) = Array(-3.142520604070, 31.744835237452, 15.150000000000)
    pts(17) = Array(-3.196664230169, 31.513282323483, 15.150000000000)
    pts(18) = Array(-3.244928136169, 31.282150523759, 15.150000000000)
    pts(19) = Array(-3.286960967891, 31.051513853523, 15.150000000000)
    FSData_333 = pts
End Function

Private Function FSData_334() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-3.322301922137, 30.821458595241, 18.250000000000)
    pts(1) = Array(-3.964620160771, 30.745435221847, 18.250000000000)
    pts(2) = Array(-4.052689975114, 30.960883853107, 18.250000000000)
    pts(3) = Array(-4.147402497751, 31.175335644090, 18.250000000000)
    pts(4) = Array(-4.248289515272, 31.388814268692, 18.250000000000)
    pts(5) = Array(-4.354992102925, 31.601329778721, 18.250000000000)
    pts(6) = Array(-4.467224780134, 31.812883047026, 18.250000000000)
    pts(7) = Array(-4.584754109672, 32.023468421672, 18.250000000000)
    pts(8) = Array(-4.707385031166, 32.233075422124, 18.250000000000)
    pts(9) = Array(-4.834951675024, 32.441689880464, 18.250000000000)
    pts(10) = Array(-4.923395912982, 33.035135427058, 18.250000000000)
    pts(11) = Array(-2.924537117806, 33.271715955847, 18.250000000000)
    pts(12) = Array(-2.872000522875, 32.674020459634, 18.250000000000)
    pts(13) = Array(-2.947342412076, 32.441390193794, 18.250000000000)
    pts(14) = Array(-3.017653248961, 32.208947031361, 18.250000000000)
    pts(15) = Array(-3.082775054361, 31.976743470282, 18.250000000000)
    pts(16) = Array(-3.142520604070, 31.744835237452, 18.250000000000)
    pts(17) = Array(-3.196664230169, 31.513282323483, 18.250000000000)
    pts(18) = Array(-3.244928136169, 31.282150523759, 18.250000000000)
    pts(19) = Array(-3.286960967891, 31.051513853523, 18.250000000000)
    FSData_334 = pts
End Function

Private Function FSData_335() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-5.730284144469, 30.465781519988, 15.150000000000)
    pts(1) = Array(-6.364657605205, 30.339596793110, 15.150000000000)
    pts(2) = Array(-6.469359834487, 30.547471390148, 15.150000000000)
    pts(3) = Array(-6.580606083624, 30.753831038948, 15.150000000000)
    pts(4) = Array(-6.697931439982, 30.958736076030, 15.150000000000)
    pts(5) = Array(-6.820978873913, 31.162224683126, 15.150000000000)
    pts(6) = Array(-6.949463853363, 31.364320128943, 15.150000000000)
    pts(7) = Array(-7.083153217270, 31.565035094241, 15.150000000000)
    pts(8) = Array(-7.221851684273, 31.764374435055, 15.150000000000)
    pts(9) = Array(-7.365392784498, 31.962337041119, 15.150000000000)
    pts(10) = Array(-7.500125579336, 32.547013938212, 15.150000000000)
    pts(11) = Array(-5.525990493096, 32.939693821744, 15.150000000000)
    pts(12) = Array(-5.426721202801, 32.347962795006, 15.150000000000)
    pts(13) = Array(-5.483578877805, 32.110138394141, 15.150000000000)
    pts(14) = Array(-5.535435689512, 31.872895251723, 15.150000000000)
    pts(15) = Array(-5.582138264692, 31.636298098764, 15.150000000000)
    pts(16) = Array(-5.623504328571, 31.400417179817, 15.150000000000)
    pts(17) = Array(-5.659313615697, 31.165330006262, 15.150000000000)
    pts(18) = Array(-5.689294348182, 30.931123966318, 15.150000000000)
    pts(19) = Array(-5.713102062123, 30.697900413998, 15.150000000000)
    FSData_335 = pts
End Function

Private Function FSData_336() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-5.730284144469, 30.465781519988, 18.250000000000)
    pts(1) = Array(-6.364657605205, 30.339596793110, 18.250000000000)
    pts(2) = Array(-6.469359834487, 30.547471390148, 18.250000000000)
    pts(3) = Array(-6.580606083624, 30.753831038948, 18.250000000000)
    pts(4) = Array(-6.697931439982, 30.958736076030, 18.250000000000)
    pts(5) = Array(-6.820978873913, 31.162224683126, 18.250000000000)
    pts(6) = Array(-6.949463853363, 31.364320128943, 18.250000000000)
    pts(7) = Array(-7.083153217270, 31.565035094241, 18.250000000000)
    pts(8) = Array(-7.221851684273, 31.764374435055, 18.250000000000)
    pts(9) = Array(-7.365392784498, 31.962337041119, 18.250000000000)
    pts(10) = Array(-7.500125579336, 32.547013938212, 18.250000000000)
    pts(11) = Array(-5.525990493096, 32.939693821744, 18.250000000000)
    pts(12) = Array(-5.426721202801, 32.347962795006, 18.250000000000)
    pts(13) = Array(-5.483578877805, 32.110138394141, 18.250000000000)
    pts(14) = Array(-5.535435689512, 31.872895251723, 18.250000000000)
    pts(15) = Array(-5.582138264692, 31.636298098764, 18.250000000000)
    pts(16) = Array(-5.623504328571, 31.400417179817, 18.250000000000)
    pts(17) = Array(-5.659313615697, 31.165330006262, 18.250000000000)
    pts(18) = Array(-5.689294348182, 30.931123966318, 18.250000000000)
    pts(19) = Array(-5.713102062123, 30.697900413998, 18.250000000000)
    FSData_336 = pts
End Function

Private Function FSData_337() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-8.102937259537, 29.922272770764, 15.150000000000)
    pts(1) = Array(-8.725454819040, 29.746704661204, 15.150000000000)
    pts(2) = Array(-8.846143939200, 29.945723608004, 15.150000000000)
    pts(3) = Array(-8.973238044717, 30.142718838765, 15.150000000000)
    pts(4) = Array(-9.106278390077, 30.337786980636, 15.150000000000)
    pts(5) = Array(-9.244912041937, 30.530994109869, 15.150000000000)
    pts(6) = Array(-9.388857171005, 30.722385747570, 15.150000000000)
    pts(7) = Array(-9.537882329896, 30.911992829015, 15.150000000000)
    pts(8) = Array(-9.691793220237, 31.099835516867, 15.150000000000)
    pts(9) = Array(-9.850423798214, 31.285925765359, 15.150000000000)
    pts(10) = Array(-10.030614477449, 31.858229285457, 15.150000000000)
    pts(11) = Array(-8.093374299418, 32.404587521669, 15.150000000000)
    pts(12) = Array(-7.947984341944, 31.822469182957, 15.150000000000)
    pts(13) = Array(-7.986007256277, 31.580916913584, 15.150000000000)
    pts(14) = Array(-8.019090328318, 31.340336474044, 15.150000000000)
    pts(15) = Array(-8.047085736372, 31.100804429331, 15.150000000000)
    pts(16) = Array(-8.069817278881, 30.862405108570, 15.150000000000)
    pts(17) = Array(-8.087071450871, 30.625233066028, 15.150000000000)
    pts(18) = Array(-8.098584168655, 30.389396743983, 15.150000000000)
    pts(19) = Array(-8.104019982343, 30.155024210333, 15.150000000000)
    FSData_337 = pts
End Function

Private Function FSData_338() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-8.102937259537, 29.922272770764, 18.250000000000)
    pts(1) = Array(-8.725454819040, 29.746704661204, 18.250000000000)
    pts(2) = Array(-8.846143939200, 29.945723608004, 18.250000000000)
    pts(3) = Array(-8.973238044717, 30.142718838765, 18.250000000000)
    pts(4) = Array(-9.106278390077, 30.337786980636, 18.250000000000)
    pts(5) = Array(-9.244912041937, 30.530994109869, 18.250000000000)
    pts(6) = Array(-9.388857171005, 30.722385747570, 18.250000000000)
    pts(7) = Array(-9.537882329896, 30.911992829015, 18.250000000000)
    pts(8) = Array(-9.691793220237, 31.099835516867, 18.250000000000)
    pts(9) = Array(-9.850423798214, 31.285925765359, 18.250000000000)
    pts(10) = Array(-10.030614477449, 31.858229285457, 18.250000000000)
    pts(11) = Array(-8.093374299418, 32.404587521669, 18.250000000000)
    pts(12) = Array(-7.947984341944, 31.822469182957, 18.250000000000)
    pts(13) = Array(-7.986007256277, 31.580916913584, 18.250000000000)
    pts(14) = Array(-8.019090328318, 31.340336474044, 18.250000000000)
    pts(15) = Array(-8.047085736372, 31.100804429331, 18.250000000000)
    pts(16) = Array(-8.069817278881, 30.862405108570, 18.250000000000)
    pts(17) = Array(-8.087071450871, 30.625233066028, 18.250000000000)
    pts(18) = Array(-8.098584168655, 30.389396743983, 18.250000000000)
    pts(19) = Array(-8.104019982343, 30.155024210333, 18.250000000000)
    FSData_338 = pts
End Function

Private Function FSData_339() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-10.425633071901, 29.194283259743, 15.150000000000)
    pts(1) = Array(-11.032456702408, 28.970414203278, 15.150000000000)
    pts(2) = Array(-11.168388624887, 29.159350481852, 15.150000000000)
    pts(3) = Array(-11.310547009360, 29.345766753470, 15.150000000000)
    pts(4) = Array(-11.458482105752, 29.529795340167, 15.150000000000)
    pts(5) = Array(-11.611847252978, 29.711529805339, 15.150000000000)
    pts(6) = Array(-11.770365062076, 29.891037641833, 15.150000000000)
    pts(7) = Array(-11.933807226290, 30.068367848717, 15.150000000000)
    pts(8) = Array(-12.101981628149, 30.243555770973, 15.150000000000)
    pts(9) = Array(-12.274723673615, 30.416626353631, 15.150000000000)
    pts(10) = Array(-12.499261301791, 30.973028055222, 15.150000000000)
    pts(11) = Array(-10.610859761864, 31.669696163905, 15.150000000000)
    pts(12) = Array(-10.420245514645, 31.100779466350, 15.150000000000)
    pts(13) = Array(-10.439199244397, 30.856988578534, 15.150000000000)
    pts(14) = Array(-10.453304608632, 30.614554100283, 15.150000000000)
    pts(15) = Array(-10.462420248560, 30.373563958524, 15.150000000000)
    pts(16) = Array(-10.466377122181, 30.134116047037, 15.150000000000)
    pts(17) = Array(-10.464969801327, 29.896321380018, 15.150000000000)
    pts(18) = Array(-10.457943524677, 29.660308785221, 15.150000000000)
    pts(19) = Array(-10.444973924511, 29.426232254848, 15.150000000000)
    FSData_339 = pts
End Function

Private Function FSData_340() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-10.425633071901, 29.194283259743, 18.250000000000)
    pts(1) = Array(-11.032456702408, 28.970414203278, 18.250000000000)
    pts(2) = Array(-11.168388624887, 29.159350481852, 18.250000000000)
    pts(3) = Array(-11.310547009360, 29.345766753470, 18.250000000000)
    pts(4) = Array(-11.458482105752, 29.529795340167, 18.250000000000)
    pts(5) = Array(-11.611847252978, 29.711529805339, 18.250000000000)
    pts(6) = Array(-11.770365062076, 29.891037641833, 18.250000000000)
    pts(7) = Array(-11.933807226290, 30.068367848717, 18.250000000000)
    pts(8) = Array(-12.101981628149, 30.243555770973, 18.250000000000)
    pts(9) = Array(-12.274723673615, 30.416626353631, 18.250000000000)
    pts(10) = Array(-12.499261301791, 30.973028055222, 18.250000000000)
    pts(11) = Array(-10.610859761864, 31.669696163905, 18.250000000000)
    pts(12) = Array(-10.420245514645, 31.100779466350, 18.250000000000)
    pts(13) = Array(-10.439199244397, 30.856988578534, 18.250000000000)
    pts(14) = Array(-10.453304608632, 30.614554100283, 18.250000000000)
    pts(15) = Array(-10.462420248560, 30.373563958524, 18.250000000000)
    pts(16) = Array(-10.466377122181, 30.134116047037, 18.250000000000)
    pts(17) = Array(-10.464969801327, 29.896321380018, 18.250000000000)
    pts(18) = Array(-10.457943524677, 29.660308785221, 18.250000000000)
    pts(19) = Array(-10.444973924511, 29.426232254848, 18.250000000000)
    FSData_340 = pts
End Function

Private Function FSData_341() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-12.684051389501, 28.286301284341, 15.150000000000)
    pts(1) = Array(-13.271439821541, 28.015511508149, 15.150000000000)
    pts(2) = Array(-13.421776480835, 28.193200263512, 15.150000000000)
    pts(3) = Array(-13.578122690552, 28.367888257682, 15.150000000000)
    pts(4) = Array(-13.740040468913, 28.539742691774, 15.150000000000)
    pts(5) = Array(-13.907191564372, 28.708884039472, 15.150000000000)
    pts(6) = Array(-14.079304738495, 28.875401349256, 15.150000000000)
    pts(7) = Array(-14.256156232741, 29.039361381885, 15.150000000000)
    pts(8) = Array(-14.437557295007, 29.200814446748, 15.150000000000)
    pts(9) = Array(-14.623345795809, 29.359798325877, 15.150000000000)
    pts(10) = Array(-14.890846023781, 29.896867807448, 15.150000000000)
    pts(11) = Array(-13.062925745409, 30.739550598048, 15.150000000000)
    pts(12) = Array(-12.828262408665, 30.187343102274, 15.150000000000)
    pts(13) = Array(-12.828030097789, 29.942816647908, 15.150000000000)
    pts(14) = Array(-12.823070789957, 29.700022820122, 15.150000000000)
    pts(15) = Array(-12.813250460808, 29.459060365677, 15.150000000000)
    pts(16) = Array(-12.798408270098, 29.220040139464, 15.150000000000)
    pts(17) = Array(-12.778348133001, 28.983088931165, 15.150000000000)
    pts(18) = Array(-12.752826181250, 28.748355159745, 15.150000000000)
    pts(19) = Array(-12.721531129129, 28.516017792297, 15.150000000000)
    FSData_341 = pts
End Function

Private Function FSData_342() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-12.684051389501, 28.286301284341, 18.250000000000)
    pts(1) = Array(-13.271439821541, 28.015511508149, 18.250000000000)
    pts(2) = Array(-13.421776480835, 28.193200263512, 18.250000000000)
    pts(3) = Array(-13.578122690552, 28.367888257682, 18.250000000000)
    pts(4) = Array(-13.740040468913, 28.539742691774, 18.250000000000)
    pts(5) = Array(-13.907191564372, 28.708884039472, 18.250000000000)
    pts(6) = Array(-14.079304738495, 28.875401349256, 18.250000000000)
    pts(7) = Array(-14.256156232741, 29.039361381885, 18.250000000000)
    pts(8) = Array(-14.437557295007, 29.200814446748, 18.250000000000)
    pts(9) = Array(-14.623345795809, 29.359798325877, 18.250000000000)
    pts(10) = Array(-14.890846023781, 29.896867807448, 18.250000000000)
    pts(11) = Array(-13.062925745409, 30.739550598048, 18.250000000000)
    pts(12) = Array(-12.828262408665, 30.187343102274, 18.250000000000)
    pts(13) = Array(-12.828030097789, 29.942816647908, 18.250000000000)
    pts(14) = Array(-12.823070789957, 29.700022820122, 18.250000000000)
    pts(15) = Array(-12.813250460808, 29.459060365677, 18.250000000000)
    pts(16) = Array(-12.798408270098, 29.220040139464, 18.250000000000)
    pts(17) = Array(-12.778348133001, 28.983088931165, 18.250000000000)
    pts(18) = Array(-12.752826181250, 28.748355159745, 18.250000000000)
    pts(19) = Array(-12.721531129129, 28.516017792297, 18.250000000000)
    FSData_342 = pts
End Function

Private Function FSData_343() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-14.864268312410, 27.203924855371, 15.150000000000)
    pts(1) = Array(-15.428600100973, 26.887883868469, 15.150000000000)
    pts(2) = Array(-15.592414621585, 27.053229590357, 15.150000000000)
    pts(3) = Array(-15.761984730172, 27.215112297505, 15.150000000000)
    pts(4) = Array(-15.936886913556, 27.373733039258, 15.150000000000)
    pts(5) = Array(-16.116793415163, 27.529238456829, 15.150000000000)
    pts(6) = Array(-16.301440819358, 27.681738605315, 15.150000000000)
    pts(7) = Array(-16.490611295364, 27.831317595565, 15.150000000000)
    pts(8) = Array(-16.684120620166, 27.978040391202, 15.150000000000)
    pts(9) = Array(-16.881810128416, 28.121957378321, 15.150000000000)
    pts(10) = Array(-17.190623728326, 28.636383427925, 15.150000000000)
    pts(11) = Array(-15.434454447870, 29.619885480815, 15.150000000000)
    pts(12) = Array(-15.157188799106, 29.087791729663, 15.150000000000)
    pts(13) = Array(-15.137771879876, 28.844037295650, 15.150000000000)
    pts(14) = Array(-15.113778475758, 28.602381023015, 15.150000000000)
    pts(15) = Array(-15.085082723126, 28.362931869545, 15.150000000000)
    pts(16) = Array(-15.051532975128, 28.125812967782, 15.150000000000)
    pts(17) = Array(-15.012943699204, 27.891166101196, 15.150000000000)
    pts(18) = Array(-14.969083423670, 27.659158364911, 15.150000000000)
    pts(19) = Array(-14.919655863997, 27.429992597518, 15.150000000000)
    FSData_343 = pts
End Function

Private Function FSData_344() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-14.864268312410, 27.203924855371, 18.250000000000)
    pts(1) = Array(-15.428600100973, 26.887883868469, 18.250000000000)
    pts(2) = Array(-15.592414621585, 27.053229590357, 18.250000000000)
    pts(3) = Array(-15.761984730172, 27.215112297505, 18.250000000000)
    pts(4) = Array(-15.936886913556, 27.373733039258, 18.250000000000)
    pts(5) = Array(-16.116793415163, 27.529238456829, 18.250000000000)
    pts(6) = Array(-16.301440819358, 27.681738605315, 18.250000000000)
    pts(7) = Array(-16.490611295364, 27.831317595565, 18.250000000000)
    pts(8) = Array(-16.684120620166, 27.978040391202, 18.250000000000)
    pts(9) = Array(-16.881810128416, 28.121957378321, 18.250000000000)
    pts(10) = Array(-17.190623728326, 28.636383427925, 18.250000000000)
    pts(11) = Array(-15.434454447870, 29.619885480815, 18.250000000000)
    pts(12) = Array(-15.157188799106, 29.087791729663, 18.250000000000)
    pts(13) = Array(-15.137771879876, 28.844037295650, 18.250000000000)
    pts(14) = Array(-15.113778475758, 28.602381023015, 18.250000000000)
    pts(15) = Array(-15.085082723126, 28.362931869545, 18.250000000000)
    pts(16) = Array(-15.051532975128, 28.125812967782, 18.250000000000)
    pts(17) = Array(-15.012943699204, 27.891166101196, 18.250000000000)
    pts(18) = Array(-14.969083423670, 27.659158364911, 18.250000000000)
    pts(19) = Array(-14.919655863997, 27.429992597518, 18.250000000000)
    FSData_344 = pts
End Function

Private Function FSData_345() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-16.952842078303, 25.953827183445, 15.150000000000)
    pts(1) = Array(-17.490637930252, 25.594483483611, 15.150000000000)
    pts(2) = Array(-17.666920341190, 25.746466760666, 15.150000000000)
    pts(3) = Array(-17.848668892540, 25.894546120071, 15.150000000000)
    pts(4) = Array(-18.035477150824, 26.038955219865, 15.150000000000)
    pts(5) = Array(-18.227029875171, 26.179865964700, 15.150000000000)
    pts(6) = Array(-18.423073096790, 26.317408737760, 15.150000000000)
    pts(7) = Array(-18.623396255666, 26.451684481417, 15.150000000000)
    pts(8) = Array(-18.827820793669, 26.582772413003, 15.150000000000)
    pts(9) = Array(-19.036192487809, 26.710735212028, 15.150000000000)
    pts(10) = Array(-19.384415521123, 27.199346222004, 15.150000000000)
    pts(11) = Array(-17.710824606183, 28.317603919982, 15.150000000000)
    pts(12) = Array(-17.392666080324, 27.808904448366, 15.150000000000)
    pts(13) = Array(-17.354184264503, 27.567424861849, 15.150000000000)
    pts(14) = Array(-17.311304691415, 27.328396035644, 15.150000000000)
    pts(15) = Array(-17.263910434158, 27.091936466805, 15.150000000000)
    pts(16) = Array(-17.211859974224, 26.858180806371, 15.150000000000)
    pts(17) = Array(-17.154979475191, 26.627284957460, 15.150000000000)
    pts(18) = Array(-17.093051289058, 26.399433661153, 15.150000000000)
    pts(19) = Array(-17.025795959174, 26.174852376978, 15.150000000000)
    FSData_345 = pts
End Function

Private Function FSData_346() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-16.952842078303, 25.953827183445, 18.250000000000)
    pts(1) = Array(-17.490637930252, 25.594483483611, 18.250000000000)
    pts(2) = Array(-17.666920341190, 25.746466760666, 18.250000000000)
    pts(3) = Array(-17.848668892540, 25.894546120071, 18.250000000000)
    pts(4) = Array(-18.035477150824, 26.038955219865, 18.250000000000)
    pts(5) = Array(-18.227029875171, 26.179865964700, 18.250000000000)
    pts(6) = Array(-18.423073096790, 26.317408737760, 18.250000000000)
    pts(7) = Array(-18.623396255666, 26.451684481417, 18.250000000000)
    pts(8) = Array(-18.827820793669, 26.582772413003, 18.250000000000)
    pts(9) = Array(-19.036192487809, 26.710735212028, 18.250000000000)
    pts(10) = Array(-19.384415521123, 27.199346222004, 18.250000000000)
    pts(11) = Array(-17.710824606183, 28.317603919982, 18.250000000000)
    pts(12) = Array(-17.392666080324, 27.808904448366, 18.250000000000)
    pts(13) = Array(-17.354184264503, 27.567424861849, 18.250000000000)
    pts(14) = Array(-17.311304691415, 27.328396035644, 18.250000000000)
    pts(15) = Array(-17.263910434158, 27.091936466805, 18.250000000000)
    pts(16) = Array(-17.211859974224, 26.858180806371, 18.250000000000)
    pts(17) = Array(-17.154979475191, 26.627284957460, 18.250000000000)
    pts(18) = Array(-17.093051289058, 26.399433661153, 18.250000000000)
    pts(19) = Array(-17.025795959174, 26.174852376978, 18.250000000000)
    FSData_346 = pts
End Function

Private Function FSData_347() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-18.936895935390, 24.543715536410, 15.150000000000)
    pts(1) = Array(-19.444840160463, 24.143284597047, 15.150000000000)
    pts(2) = Array(-19.632503622043, 24.280968401826, 15.150000000000)
    pts(3) = Array(-19.825310075900, 24.414331454996, 15.150000000000)
    pts(4) = Array(-20.022872674052, 24.543638582710, 15.150000000000)
    pts(5) = Array(-20.224890634896, 24.669085893209, 15.150000000000)
    pts(6) = Array(-20.431121002287, 24.790823293910, 15.150000000000)
    pts(7) = Array(-20.641361785144, 24.908967936364, 15.150000000000)
    pts(8) = Array(-20.855441191093, 25.023612803208, 15.150000000000)
    pts(9) = Array(-21.073210390339, 25.134832480932, 15.150000000000)
    pts(10) = Array(-21.458695946260, 25.594616001925, 15.150000000000)
    pts(11) = Array(-19.878001641353, 26.840734914424, 15.150000000000)
    pts(12) = Array(-19.520911791508, 26.358566023746, 15.150000000000)
    pts(13) = Array(-19.463602332288, 26.120850086675, 15.150000000000)
    pts(14) = Array(-19.402100957056, 25.885922399100, 15.150000000000)
    pts(15) = Array(-19.336300396530, 25.653910266764, 15.150000000000)
    pts(16) = Array(-19.266070133054, 25.424959029038, 15.150000000000)
    pts(17) = Array(-19.191249098104, 25.199237747490, 15.150000000000)
    pts(18) = Array(-19.111634809233, 24.976947670172, 15.150000000000)
    pts(19) = Array(-19.026966360612, 24.758335487511, 15.150000000000)
    FSData_347 = pts
End Function

Private Function FSData_348() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-18.936895935390, 24.543715536410, 18.250000000000)
    pts(1) = Array(-19.444840160463, 24.143284597047, 18.250000000000)
    pts(2) = Array(-19.632503622043, 24.280968401826, 18.250000000000)
    pts(3) = Array(-19.825310075900, 24.414331454996, 18.250000000000)
    pts(4) = Array(-20.022872674052, 24.543638582710, 18.250000000000)
    pts(5) = Array(-20.224890634896, 24.669085893209, 18.250000000000)
    pts(6) = Array(-20.431121002287, 24.790823293910, 18.250000000000)
    pts(7) = Array(-20.641361785144, 24.908967936364, 18.250000000000)
    pts(8) = Array(-20.855441191093, 25.023612803208, 18.250000000000)
    pts(9) = Array(-21.073210390339, 25.134832480932, 18.250000000000)
    pts(10) = Array(-21.458695946260, 25.594616001925, 18.250000000000)
    pts(11) = Array(-19.878001641353, 26.840734914424, 18.250000000000)
    pts(12) = Array(-19.520911791508, 26.358566023746, 18.250000000000)
    pts(13) = Array(-19.463602332288, 26.120850086675, 18.250000000000)
    pts(14) = Array(-19.402100957056, 25.885922399100, 18.250000000000)
    pts(15) = Array(-19.336300396530, 25.653910266764, 18.250000000000)
    pts(16) = Array(-19.266070133054, 25.424959029038, 18.250000000000)
    pts(17) = Array(-19.191249098104, 25.199237747490, 18.250000000000)
    pts(18) = Array(-19.111634809233, 24.976947670172, 18.250000000000)
    pts(19) = Array(-19.026966360612, 24.758335487511, 18.250000000000)
    FSData_348 = pts
End Function

Private Function FSData_349() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-20.804197531879, 22.982283721479, 15.150000000000)
    pts(1) = Array(-21.279158485020, 22.543234332486, 15.150000000000)
    pts(2) = Array(-21.477045989597, 22.665769798547, 15.150000000000)
    pts(3) = Array(-21.679721630058, 22.783594317913, 15.150000000000)
    pts(4) = Array(-21.886820528965, 22.897002252105, 15.150000000000)
    pts(5) = Array(-22.098058218398, 23.006212703883, 15.150000000000)
    pts(6) = Array(-22.313204252768, 23.111394180669, 15.150000000000)
    pts(7) = Array(-22.532066455268, 23.212679320910, 15.150000000000)
    pts(8) = Array(-22.754480858436, 23.310174299286, 15.150000000000)
    pts(9) = Array(-22.980304943259, 23.403965149411, 15.150000000000)
    pts(10) = Array(-23.400676375148, 23.832086463120, 15.150000000000)
    pts(11) = Array(-21.922624186296, 25.198383852668, 15.150000000000)
    pts(12) = Array(-21.528804590135, 24.745718274477, 15.150000000000)
    pts(13) = Array(-21.453020819389, 24.513231584653, 15.150000000000)
    pts(14) = Array(-21.373276818444, 24.283853443022, 15.150000000000)
    pts(15) = Array(-21.289475636985, 24.057719179137, 15.150000000000)
    pts(16) = Array(-21.201498562895, 23.834983924634, 15.150000000000)
    pts(17) = Array(-21.109198288589, 23.615828857210, 15.150000000000)
    pts(18) = Array(-21.012388745545, 23.400470491127, 15.150000000000)
    pts(19) = Array(-20.910829187328, 23.189175226782, 15.150000000000)
    FSData_349 = pts
End Function

Private Function FSData_350() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-20.804197531879, 22.982283721479, 18.250000000000)
    pts(1) = Array(-21.279158485020, 22.543234332486, 18.250000000000)
    pts(2) = Array(-21.477045989597, 22.665769798547, 18.250000000000)
    pts(3) = Array(-21.679721630058, 22.783594317913, 18.250000000000)
    pts(4) = Array(-21.886820528965, 22.897002252105, 18.250000000000)
    pts(5) = Array(-22.098058218398, 23.006212703883, 18.250000000000)
    pts(6) = Array(-22.313204252768, 23.111394180669, 18.250000000000)
    pts(7) = Array(-22.532066455268, 23.212679320910, 18.250000000000)
    pts(8) = Array(-22.754480858436, 23.310174299286, 18.250000000000)
    pts(9) = Array(-22.980304943259, 23.403965149411, 18.250000000000)
    pts(10) = Array(-23.400676375148, 23.832086463120, 18.250000000000)
    pts(11) = Array(-21.922624186296, 25.198383852668, 18.250000000000)
    pts(12) = Array(-21.528804590135, 24.745718274477, 18.250000000000)
    pts(13) = Array(-21.453020819389, 24.513231584653, 18.250000000000)
    pts(14) = Array(-21.373276818444, 24.283853443022, 18.250000000000)
    pts(15) = Array(-21.289475636985, 24.057719179137, 18.250000000000)
    pts(16) = Array(-21.201498562895, 23.834983924634, 18.250000000000)
    pts(17) = Array(-21.109198288589, 23.615828857210, 18.250000000000)
    pts(18) = Array(-21.012388745545, 23.400470491127, 18.250000000000)
    pts(19) = Array(-20.910829187328, 23.189175226782, 18.250000000000)
    FSData_350 = pts
End Function

Private Function FSData_351() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-22.543234332486, 21.279158485020, 15.150000000000)
    pts(1) = Array(-22.982283721479, 20.804197531879, 15.150000000000)
    pts(2) = Array(-23.189175226782, 20.910829187328, 15.150000000000)
    pts(3) = Array(-23.400470491127, 21.012388745545, 15.150000000000)
    pts(4) = Array(-23.615828857210, 21.109198288589, 15.150000000000)
    pts(5) = Array(-23.834983924634, 21.201498562895, 15.150000000000)
    pts(6) = Array(-24.057719179137, 21.289475636985, 15.150000000000)
    pts(7) = Array(-24.283853443022, 21.373276818444, 15.150000000000)
    pts(8) = Array(-24.513231584653, 21.453020819389, 15.150000000000)
    pts(9) = Array(-24.745718274477, 21.528804590135, 15.150000000000)
    pts(10) = Array(-25.198383852668, 21.922624186296, 15.150000000000)
    pts(11) = Array(-23.832086463120, 23.400676375148, 15.150000000000)
    pts(12) = Array(-23.403965149411, 22.980304943259, 15.150000000000)
    pts(13) = Array(-23.310174299286, 22.754480858436, 15.150000000000)
    pts(14) = Array(-23.212679320910, 22.532066455268, 15.150000000000)
    pts(15) = Array(-23.111394180669, 22.313204252768, 15.150000000000)
    pts(16) = Array(-23.006212703883, 22.098058218398, 15.150000000000)
    pts(17) = Array(-22.897002252105, 21.886820528965, 15.150000000000)
    pts(18) = Array(-22.783594317913, 21.679721630058, 15.150000000000)
    pts(19) = Array(-22.665769798547, 21.477045989597, 15.150000000000)
    FSData_351 = pts
End Function

Private Function FSData_352() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-22.543234332486, 21.279158485020, 18.250000000000)
    pts(1) = Array(-22.982283721479, 20.804197531879, 18.250000000000)
    pts(2) = Array(-23.189175226782, 20.910829187328, 18.250000000000)
    pts(3) = Array(-23.400470491127, 21.012388745545, 18.250000000000)
    pts(4) = Array(-23.615828857210, 21.109198288589, 18.250000000000)
    pts(5) = Array(-23.834983924634, 21.201498562895, 18.250000000000)
    pts(6) = Array(-24.057719179137, 21.289475636985, 18.250000000000)
    pts(7) = Array(-24.283853443022, 21.373276818444, 18.250000000000)
    pts(8) = Array(-24.513231584653, 21.453020819389, 18.250000000000)
    pts(9) = Array(-24.745718274477, 21.528804590135, 18.250000000000)
    pts(10) = Array(-25.198383852668, 21.922624186296, 18.250000000000)
    pts(11) = Array(-23.832086463120, 23.400676375148, 18.250000000000)
    pts(12) = Array(-23.403965149411, 22.980304943259, 18.250000000000)
    pts(13) = Array(-23.310174299286, 22.754480858436, 18.250000000000)
    pts(14) = Array(-23.212679320910, 22.532066455268, 18.250000000000)
    pts(15) = Array(-23.111394180669, 22.313204252768, 18.250000000000)
    pts(16) = Array(-23.006212703883, 22.098058218398, 18.250000000000)
    pts(17) = Array(-22.897002252105, 21.886820528965, 18.250000000000)
    pts(18) = Array(-22.783594317913, 21.679721630058, 18.250000000000)
    pts(19) = Array(-22.665769798547, 21.477045989597, 18.250000000000)
    FSData_352 = pts
End Function

Private Function FSData_353() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-24.143284597047, 19.444840160463, 15.150000000000)
    pts(1) = Array(-24.543715536410, 18.936895935390, 15.150000000000)
    pts(2) = Array(-24.758335487511, 19.026966360612, 15.150000000000)
    pts(3) = Array(-24.976947670172, 19.111634809233, 15.150000000000)
    pts(4) = Array(-25.199237747490, 19.191249098104, 15.150000000000)
    pts(5) = Array(-25.424959029038, 19.266070133054, 15.150000000000)
    pts(6) = Array(-25.653910266764, 19.336300396530, 15.150000000000)
    pts(7) = Array(-25.885922399100, 19.402100957056, 15.150000000000)
    pts(8) = Array(-26.120850086675, 19.463602332288, 15.150000000000)
    pts(9) = Array(-26.358566023746, 19.520911791508, 15.150000000000)
    pts(10) = Array(-26.840734914424, 19.878001641353, 15.150000000000)
    pts(11) = Array(-25.594616001925, 21.458695946260, 15.150000000000)
    pts(12) = Array(-25.134832480932, 21.073210390339, 15.150000000000)
    pts(13) = Array(-25.023612803208, 20.855441191093, 15.150000000000)
    pts(14) = Array(-24.908967936364, 20.641361785144, 15.150000000000)
    pts(15) = Array(-24.790823293910, 20.431121002287, 15.150000000000)
    pts(16) = Array(-24.669085893209, 20.224890634896, 15.150000000000)
    pts(17) = Array(-24.543638582710, 20.022872674052, 15.150000000000)
    pts(18) = Array(-24.414331454996, 19.825310075900, 15.150000000000)
    pts(19) = Array(-24.280968401826, 19.632503622043, 15.150000000000)
    FSData_353 = pts
End Function

Private Function FSData_354() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-24.143284597047, 19.444840160463, 18.250000000000)
    pts(1) = Array(-24.543715536410, 18.936895935390, 18.250000000000)
    pts(2) = Array(-24.758335487511, 19.026966360612, 18.250000000000)
    pts(3) = Array(-24.976947670172, 19.111634809233, 18.250000000000)
    pts(4) = Array(-25.199237747490, 19.191249098104, 18.250000000000)
    pts(5) = Array(-25.424959029038, 19.266070133054, 18.250000000000)
    pts(6) = Array(-25.653910266764, 19.336300396530, 18.250000000000)
    pts(7) = Array(-25.885922399100, 19.402100957056, 18.250000000000)
    pts(8) = Array(-26.120850086675, 19.463602332288, 18.250000000000)
    pts(9) = Array(-26.358566023746, 19.520911791508, 18.250000000000)
    pts(10) = Array(-26.840734914424, 19.878001641353, 18.250000000000)
    pts(11) = Array(-25.594616001925, 21.458695946260, 18.250000000000)
    pts(12) = Array(-25.134832480932, 21.073210390339, 18.250000000000)
    pts(13) = Array(-25.023612803208, 20.855441191093, 18.250000000000)
    pts(14) = Array(-24.908967936364, 20.641361785144, 18.250000000000)
    pts(15) = Array(-24.790823293910, 20.431121002287, 18.250000000000)
    pts(16) = Array(-24.669085893209, 20.224890634896, 18.250000000000)
    pts(17) = Array(-24.543638582710, 20.022872674052, 18.250000000000)
    pts(18) = Array(-24.414331454996, 19.825310075900, 18.250000000000)
    pts(19) = Array(-24.280968401826, 19.632503622043, 18.250000000000)
    FSData_354 = pts
End Function

Private Function FSData_355() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-25.594483483611, 17.490637930252, 15.150000000000)
    pts(1) = Array(-25.953827183445, 16.952842078303, 15.150000000000)
    pts(2) = Array(-26.174852376978, 17.025795959174, 15.150000000000)
    pts(3) = Array(-26.399433661153, 17.093051289058, 15.150000000000)
    pts(4) = Array(-26.627284957460, 17.154979475191, 15.150000000000)
    pts(5) = Array(-26.858180806371, 17.211859974224, 15.150000000000)
    pts(6) = Array(-27.091936466805, 17.263910434158, 15.150000000000)
    pts(7) = Array(-27.328396035644, 17.311304691415, 15.150000000000)
    pts(8) = Array(-27.567424861849, 17.354184264503, 15.150000000000)
    pts(9) = Array(-27.808904448366, 17.392666080324, 15.150000000000)
    pts(10) = Array(-28.317603919982, 17.710824606183, 15.150000000000)
    pts(11) = Array(-27.199346222004, 19.384415521123, 15.150000000000)
    pts(12) = Array(-26.710735212028, 19.036192487809, 15.150000000000)
    pts(13) = Array(-26.582772413003, 18.827820793669, 15.150000000000)
    pts(14) = Array(-26.451684481417, 18.623396255666, 15.150000000000)
    pts(15) = Array(-26.317408737760, 18.423073096790, 15.150000000000)
    pts(16) = Array(-26.179865964700, 18.227029875171, 15.150000000000)
    pts(17) = Array(-26.038955219865, 18.035477150824, 15.150000000000)
    pts(18) = Array(-25.894546120071, 17.848668892540, 15.150000000000)
    pts(19) = Array(-25.746466760666, 17.666920341190, 15.150000000000)
    FSData_355 = pts
End Function

Private Function FSData_356() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-25.594483483611, 17.490637930252, 18.250000000000)
    pts(1) = Array(-25.953827183445, 16.952842078303, 18.250000000000)
    pts(2) = Array(-26.174852376978, 17.025795959174, 18.250000000000)
    pts(3) = Array(-26.399433661153, 17.093051289058, 18.250000000000)
    pts(4) = Array(-26.627284957460, 17.154979475191, 18.250000000000)
    pts(5) = Array(-26.858180806371, 17.211859974224, 18.250000000000)
    pts(6) = Array(-27.091936466805, 17.263910434158, 18.250000000000)
    pts(7) = Array(-27.328396035644, 17.311304691415, 18.250000000000)
    pts(8) = Array(-27.567424861849, 17.354184264503, 18.250000000000)
    pts(9) = Array(-27.808904448366, 17.392666080324, 18.250000000000)
    pts(10) = Array(-28.317603919982, 17.710824606183, 18.250000000000)
    pts(11) = Array(-27.199346222004, 19.384415521123, 18.250000000000)
    pts(12) = Array(-26.710735212028, 19.036192487809, 18.250000000000)
    pts(13) = Array(-26.582772413003, 18.827820793669, 18.250000000000)
    pts(14) = Array(-26.451684481417, 18.623396255666, 18.250000000000)
    pts(15) = Array(-26.317408737760, 18.423073096790, 18.250000000000)
    pts(16) = Array(-26.179865964700, 18.227029875171, 18.250000000000)
    pts(17) = Array(-26.038955219865, 18.035477150824, 18.250000000000)
    pts(18) = Array(-25.894546120071, 17.848668892540, 18.250000000000)
    pts(19) = Array(-25.746466760666, 17.666920341190, 18.250000000000)
    FSData_356 = pts
End Function

Private Function FSData_357() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-26.887883868469, 15.428600100973, 15.150000000000)
    pts(1) = Array(-27.203924855371, 14.864268312410, 15.150000000000)
    pts(2) = Array(-27.429992597518, 14.919655863997, 15.150000000000)
    pts(3) = Array(-27.659158364911, 14.969083423670, 15.150000000000)
    pts(4) = Array(-27.891166101196, 15.012943699204, 15.150000000000)
    pts(5) = Array(-28.125812967782, 15.051532975128, 15.150000000000)
    pts(6) = Array(-28.362931869545, 15.085082723126, 15.150000000000)
    pts(7) = Array(-28.602381023015, 15.113778475758, 15.150000000000)
    pts(8) = Array(-28.844037295650, 15.137771879876, 15.150000000000)
    pts(9) = Array(-29.087791729663, 15.157188799106, 15.150000000000)
    pts(10) = Array(-29.619885480815, 15.434454447870, 15.150000000000)
    pts(11) = Array(-28.636383427925, 17.190623728326, 15.150000000000)
    pts(12) = Array(-28.121957378321, 16.881810128416, 15.150000000000)
    pts(13) = Array(-27.978040391202, 16.684120620166, 15.150000000000)
    pts(14) = Array(-27.831317595565, 16.490611295364, 15.150000000000)
    pts(15) = Array(-27.681738605315, 16.301440819358, 15.150000000000)
    pts(16) = Array(-27.529238456829, 16.116793415163, 15.150000000000)
    pts(17) = Array(-27.373733039258, 15.936886913556, 15.150000000000)
    pts(18) = Array(-27.215112297505, 15.761984730172, 15.150000000000)
    pts(19) = Array(-27.053229590357, 15.592414621585, 15.150000000000)
    FSData_357 = pts
End Function

Private Function FSData_358() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-26.887883868469, 15.428600100973, 18.250000000000)
    pts(1) = Array(-27.203924855371, 14.864268312410, 18.250000000000)
    pts(2) = Array(-27.429992597518, 14.919655863997, 18.250000000000)
    pts(3) = Array(-27.659158364911, 14.969083423670, 18.250000000000)
    pts(4) = Array(-27.891166101196, 15.012943699204, 18.250000000000)
    pts(5) = Array(-28.125812967782, 15.051532975128, 18.250000000000)
    pts(6) = Array(-28.362931869545, 15.085082723126, 18.250000000000)
    pts(7) = Array(-28.602381023015, 15.113778475758, 18.250000000000)
    pts(8) = Array(-28.844037295650, 15.137771879876, 18.250000000000)
    pts(9) = Array(-29.087791729663, 15.157188799106, 18.250000000000)
    pts(10) = Array(-29.619885480815, 15.434454447870, 18.250000000000)
    pts(11) = Array(-28.636383427925, 17.190623728326, 18.250000000000)
    pts(12) = Array(-28.121957378321, 16.881810128416, 18.250000000000)
    pts(13) = Array(-27.978040391202, 16.684120620166, 18.250000000000)
    pts(14) = Array(-27.831317595565, 16.490611295364, 18.250000000000)
    pts(15) = Array(-27.681738605315, 16.301440819358, 18.250000000000)
    pts(16) = Array(-27.529238456829, 16.116793415163, 18.250000000000)
    pts(17) = Array(-27.373733039258, 15.936886913556, 18.250000000000)
    pts(18) = Array(-27.215112297505, 15.761984730172, 18.250000000000)
    pts(19) = Array(-27.053229590357, 15.592414621585, 18.250000000000)
    FSData_358 = pts
End Function

Private Function FSData_359() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-28.015511508149, 13.271439821541, 15.150000000000)
    pts(1) = Array(-28.286301284341, 12.684051389501, 15.150000000000)
    pts(2) = Array(-28.516017792297, 12.721531129129, 15.150000000000)
    pts(3) = Array(-28.748355159745, 12.752826181250, 15.150000000000)
    pts(4) = Array(-28.983088931165, 12.778348133001, 15.150000000000)
    pts(5) = Array(-29.220040139464, 12.798408270098, 15.150000000000)
    pts(6) = Array(-29.459060365677, 12.813250460808, 15.150000000000)
    pts(7) = Array(-29.700022820122, 12.823070789957, 15.150000000000)
    pts(8) = Array(-29.942816647908, 12.828030097789, 15.150000000000)
    pts(9) = Array(-30.187343102274, 12.828262408665, 15.150000000000)
    pts(10) = Array(-30.739550598048, 13.062925745409, 15.150000000000)
    pts(11) = Array(-29.896867807448, 14.890846023781, 15.150000000000)
    pts(12) = Array(-29.359798325877, 14.623345795809, 15.150000000000)
    pts(13) = Array(-29.200814446748, 14.437557295007, 15.150000000000)
    pts(14) = Array(-29.039361381885, 14.256156232741, 15.150000000000)
    pts(15) = Array(-28.875401349256, 14.079304738495, 15.150000000000)
    pts(16) = Array(-28.708884039472, 13.907191564372, 15.150000000000)
    pts(17) = Array(-28.539742691774, 13.740040468913, 15.150000000000)
    pts(18) = Array(-28.367888257682, 13.578122690552, 15.150000000000)
    pts(19) = Array(-28.193200263512, 13.421776480835, 15.150000000000)
    FSData_359 = pts
End Function

Private Function FSData_360() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-28.015511508149, 13.271439821541, 18.250000000000)
    pts(1) = Array(-28.286301284341, 12.684051389501, 18.250000000000)
    pts(2) = Array(-28.516017792297, 12.721531129129, 18.250000000000)
    pts(3) = Array(-28.748355159745, 12.752826181250, 18.250000000000)
    pts(4) = Array(-28.983088931165, 12.778348133001, 18.250000000000)
    pts(5) = Array(-29.220040139464, 12.798408270098, 18.250000000000)
    pts(6) = Array(-29.459060365677, 12.813250460808, 18.250000000000)
    pts(7) = Array(-29.700022820122, 12.823070789957, 18.250000000000)
    pts(8) = Array(-29.942816647908, 12.828030097789, 18.250000000000)
    pts(9) = Array(-30.187343102274, 12.828262408665, 18.250000000000)
    pts(10) = Array(-30.739550598048, 13.062925745409, 18.250000000000)
    pts(11) = Array(-29.896867807448, 14.890846023781, 18.250000000000)
    pts(12) = Array(-29.359798325877, 14.623345795809, 18.250000000000)
    pts(13) = Array(-29.200814446748, 14.437557295007, 18.250000000000)
    pts(14) = Array(-29.039361381885, 14.256156232741, 18.250000000000)
    pts(15) = Array(-28.875401349256, 14.079304738495, 18.250000000000)
    pts(16) = Array(-28.708884039472, 13.907191564372, 18.250000000000)
    pts(17) = Array(-28.539742691774, 13.740040468913, 18.250000000000)
    pts(18) = Array(-28.367888257682, 13.578122690552, 18.250000000000)
    pts(19) = Array(-28.193200263512, 13.421776480835, 18.250000000000)
    FSData_360 = pts
End Function

Private Function FSData_361() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-28.970414203278, 11.032456702408, 15.150000000000)
    pts(1) = Array(-29.194283259743, 10.425633071901, 15.150000000000)
    pts(2) = Array(-29.426232254848, 10.444973924511, 15.150000000000)
    pts(3) = Array(-29.660308785221, 10.457943524677, 15.150000000000)
    pts(4) = Array(-29.896321380018, 10.464969801327, 15.150000000000)
    pts(5) = Array(-30.134116047037, 10.466377122181, 15.150000000000)
    pts(6) = Array(-30.373563958524, 10.462420248560, 15.150000000000)
    pts(7) = Array(-30.614554100283, 10.453304608632, 15.150000000000)
    pts(8) = Array(-30.856988578534, 10.439199244397, 15.150000000000)
    pts(9) = Array(-31.100779466350, 10.420245514645, 15.150000000000)
    pts(10) = Array(-31.669696163905, 10.610859761864, 15.150000000000)
    pts(11) = Array(-30.973028055222, 12.499261301791, 15.150000000000)
    pts(12) = Array(-30.416626353631, 12.274723673615, 15.150000000000)
    pts(13) = Array(-30.243555770973, 12.101981628149, 15.150000000000)
    pts(14) = Array(-30.068367848717, 11.933807226290, 15.150000000000)
    pts(15) = Array(-29.891037641833, 11.770365062076, 15.150000000000)
    pts(16) = Array(-29.711529805339, 11.611847252978, 15.150000000000)
    pts(17) = Array(-29.529795340167, 11.458482105752, 15.150000000000)
    pts(18) = Array(-29.345766753470, 11.310547009360, 15.150000000000)
    pts(19) = Array(-29.159350481852, 11.168388624887, 15.150000000000)
    FSData_361 = pts
End Function

Private Function FSData_362() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-28.970414203278, 11.032456702408, 18.250000000000)
    pts(1) = Array(-29.194283259743, 10.425633071901, 18.250000000000)
    pts(2) = Array(-29.426232254848, 10.444973924511, 18.250000000000)
    pts(3) = Array(-29.660308785221, 10.457943524677, 18.250000000000)
    pts(4) = Array(-29.896321380018, 10.464969801327, 18.250000000000)
    pts(5) = Array(-30.134116047037, 10.466377122181, 18.250000000000)
    pts(6) = Array(-30.373563958524, 10.462420248560, 18.250000000000)
    pts(7) = Array(-30.614554100283, 10.453304608632, 18.250000000000)
    pts(8) = Array(-30.856988578534, 10.439199244397, 18.250000000000)
    pts(9) = Array(-31.100779466350, 10.420245514645, 18.250000000000)
    pts(10) = Array(-31.669696163905, 10.610859761864, 18.250000000000)
    pts(11) = Array(-30.973028055222, 12.499261301791, 18.250000000000)
    pts(12) = Array(-30.416626353631, 12.274723673615, 18.250000000000)
    pts(13) = Array(-30.243555770973, 12.101981628149, 18.250000000000)
    pts(14) = Array(-30.068367848717, 11.933807226290, 18.250000000000)
    pts(15) = Array(-29.891037641833, 11.770365062076, 18.250000000000)
    pts(16) = Array(-29.711529805339, 11.611847252978, 18.250000000000)
    pts(17) = Array(-29.529795340167, 11.458482105752, 18.250000000000)
    pts(18) = Array(-29.345766753470, 11.310547009360, 18.250000000000)
    pts(19) = Array(-29.159350481852, 11.168388624887, 18.250000000000)
    FSData_362 = pts
End Function

Private Function FSData_363() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-29.746704661204, 8.725454819040, 15.150000000000)
    pts(1) = Array(-29.922272770764, 8.102937259537, 15.150000000000)
    pts(2) = Array(-30.155024210333, 8.104019982343, 15.150000000000)
    pts(3) = Array(-30.389396743983, 8.098584168656, 15.150000000000)
    pts(4) = Array(-30.625233066028, 8.087071450871, 15.150000000000)
    pts(5) = Array(-30.862405108570, 8.069817278881, 15.150000000000)
    pts(6) = Array(-31.100804429331, 8.047085736372, 15.150000000000)
    pts(7) = Array(-31.340336474044, 8.019090328318, 15.150000000000)
    pts(8) = Array(-31.580916913584, 7.986007256277, 15.150000000000)
    pts(9) = Array(-31.822469182957, 7.947984341944, 15.150000000000)
    pts(10) = Array(-32.404587521669, 8.093374299418, 15.150000000000)
    pts(11) = Array(-31.858229285457, 10.030614477449, 15.150000000000)
    pts(12) = Array(-31.285925765359, 9.850423798214, 15.150000000000)
    pts(13) = Array(-31.099835516867, 9.691793220237, 15.150000000000)
    pts(14) = Array(-30.911992829015, 9.537882329896, 15.150000000000)
    pts(15) = Array(-30.722385747570, 9.388857171005, 15.150000000000)
    pts(16) = Array(-30.530994109869, 9.244912041937, 15.150000000000)
    pts(17) = Array(-30.337786980636, 9.106278390077, 15.150000000000)
    pts(18) = Array(-30.142718838765, 8.973238044717, 15.150000000000)
    pts(19) = Array(-29.945723608004, 8.846143939200, 15.150000000000)
    FSData_363 = pts
End Function

Private Function FSData_364() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-29.746704661204, 8.725454819040, 18.250000000000)
    pts(1) = Array(-29.922272770764, 8.102937259537, 18.250000000000)
    pts(2) = Array(-30.155024210333, 8.104019982343, 18.250000000000)
    pts(3) = Array(-30.389396743983, 8.098584168656, 18.250000000000)
    pts(4) = Array(-30.625233066028, 8.087071450871, 18.250000000000)
    pts(5) = Array(-30.862405108570, 8.069817278881, 18.250000000000)
    pts(6) = Array(-31.100804429331, 8.047085736372, 18.250000000000)
    pts(7) = Array(-31.340336474044, 8.019090328318, 18.250000000000)
    pts(8) = Array(-31.580916913584, 7.986007256277, 18.250000000000)
    pts(9) = Array(-31.822469182957, 7.947984341944, 18.250000000000)
    pts(10) = Array(-32.404587521669, 8.093374299418, 18.250000000000)
    pts(11) = Array(-31.858229285457, 10.030614477449, 18.250000000000)
    pts(12) = Array(-31.285925765359, 9.850423798214, 18.250000000000)
    pts(13) = Array(-31.099835516867, 9.691793220237, 18.250000000000)
    pts(14) = Array(-30.911992829015, 9.537882329896, 18.250000000000)
    pts(15) = Array(-30.722385747570, 9.388857171005, 18.250000000000)
    pts(16) = Array(-30.530994109869, 9.244912041937, 18.250000000000)
    pts(17) = Array(-30.337786980636, 9.106278390077, 18.250000000000)
    pts(18) = Array(-30.142718838765, 8.973238044717, 18.250000000000)
    pts(19) = Array(-29.945723608004, 8.846143939200, 18.250000000000)
    FSData_364 = pts
End Function

Private Function FSData_365() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.339596793110, 6.364657605205, 15.150000000000)
    pts(1) = Array(-30.465781519988, 5.730284144469, 15.150000000000)
    pts(2) = Array(-30.697900413998, 5.713102062123, 15.150000000000)
    pts(3) = Array(-30.931123966318, 5.689294348182, 15.150000000000)
    pts(4) = Array(-31.165330006262, 5.659313615697, 15.150000000000)
    pts(5) = Array(-31.400417179817, 5.623504328571, 15.150000000000)
    pts(6) = Array(-31.636298098764, 5.582138264692, 15.150000000000)
    pts(7) = Array(-31.872895251723, 5.535435689512, 15.150000000000)
    pts(8) = Array(-32.110138394141, 5.483578877805, 15.150000000000)
    pts(9) = Array(-32.347962795006, 5.426721202801, 15.150000000000)
    pts(10) = Array(-32.939693821744, 5.525990493096, 15.150000000000)
    pts(11) = Array(-32.547013938212, 7.500125579336, 15.150000000000)
    pts(12) = Array(-31.962337041119, 7.365392784498, 15.150000000000)
    pts(13) = Array(-31.764374435055, 7.221851684273, 15.150000000000)
    pts(14) = Array(-31.565035094241, 7.083153217270, 15.150000000000)
    pts(15) = Array(-31.364320128943, 6.949463853363, 15.150000000000)
    pts(16) = Array(-31.162224683126, 6.820978873913, 15.150000000000)
    pts(17) = Array(-30.958736076030, 6.697931439982, 15.150000000000)
    pts(18) = Array(-30.753831038948, 6.580606083624, 15.150000000000)
    pts(19) = Array(-30.547471390148, 6.469359834487, 15.150000000000)
    FSData_365 = pts
End Function

Private Function FSData_366() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.339596793110, 6.364657605205, 18.250000000000)
    pts(1) = Array(-30.465781519988, 5.730284144469, 18.250000000000)
    pts(2) = Array(-30.697900413998, 5.713102062123, 18.250000000000)
    pts(3) = Array(-30.931123966318, 5.689294348182, 18.250000000000)
    pts(4) = Array(-31.165330006262, 5.659313615697, 18.250000000000)
    pts(5) = Array(-31.400417179817, 5.623504328571, 18.250000000000)
    pts(6) = Array(-31.636298098764, 5.582138264692, 18.250000000000)
    pts(7) = Array(-31.872895251723, 5.535435689512, 18.250000000000)
    pts(8) = Array(-32.110138394141, 5.483578877805, 18.250000000000)
    pts(9) = Array(-32.347962795006, 5.426721202801, 18.250000000000)
    pts(10) = Array(-32.939693821744, 5.525990493096, 18.250000000000)
    pts(11) = Array(-32.547013938212, 7.500125579336, 18.250000000000)
    pts(12) = Array(-31.962337041119, 7.365392784498, 18.250000000000)
    pts(13) = Array(-31.764374435055, 7.221851684273, 18.250000000000)
    pts(14) = Array(-31.565035094241, 7.083153217270, 18.250000000000)
    pts(15) = Array(-31.364320128943, 6.949463853363, 18.250000000000)
    pts(16) = Array(-31.162224683126, 6.820978873913, 18.250000000000)
    pts(17) = Array(-30.958736076030, 6.697931439982, 18.250000000000)
    pts(18) = Array(-30.753831038948, 6.580606083624, 18.250000000000)
    pts(19) = Array(-30.547471390148, 6.469359834487, 18.250000000000)
    FSData_366 = pts
End Function

Private Function FSData_367() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.745435221847, 3.964620160771, 15.150000000000)
    pts(1) = Array(-30.821458595241, 3.322301922137, 15.150000000000)
    pts(2) = Array(-31.051513853523, 3.286960967891, 15.150000000000)
    pts(3) = Array(-31.282150523759, 3.244928136169, 15.150000000000)
    pts(4) = Array(-31.513282323483, 3.196664230169, 15.150000000000)
    pts(5) = Array(-31.744835237452, 3.142520604070, 15.150000000000)
    pts(6) = Array(-31.976743470282, 3.082775054361, 15.150000000000)
    pts(7) = Array(-32.208947031361, 3.017653248961, 15.150000000000)
    pts(8) = Array(-32.441390193794, 2.947342412076, 15.150000000000)
    pts(9) = Array(-32.674020459634, 2.872000522875, 15.150000000000)
    pts(10) = Array(-33.271715955847, 2.924537117806, 15.150000000000)
    pts(11) = Array(-33.035135427058, 4.923395912982, 15.150000000000)
    pts(12) = Array(-32.441689880464, 4.834951675024, 15.150000000000)
    pts(13) = Array(-32.233075422124, 4.707385031166, 15.150000000000)
    pts(14) = Array(-32.023468421672, 4.584754109672, 15.150000000000)
    pts(15) = Array(-31.812883047026, 4.467224780134, 15.150000000000)
    pts(16) = Array(-31.601329778721, 4.354992102925, 15.150000000000)
    pts(17) = Array(-31.388814268692, 4.248289515272, 15.150000000000)
    pts(18) = Array(-31.175335644090, 4.147402497751, 15.150000000000)
    pts(19) = Array(-30.960883853107, 4.052689975114, 15.150000000000)
    FSData_367 = pts
End Function

Private Function FSData_368() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.745435221847, 3.964620160771, 18.250000000000)
    pts(1) = Array(-30.821458595241, 3.322301922137, 18.250000000000)
    pts(2) = Array(-31.051513853523, 3.286960967891, 18.250000000000)
    pts(3) = Array(-31.282150523759, 3.244928136169, 18.250000000000)
    pts(4) = Array(-31.513282323483, 3.196664230169, 18.250000000000)
    pts(5) = Array(-31.744835237452, 3.142520604070, 18.250000000000)
    pts(6) = Array(-31.976743470282, 3.082775054361, 18.250000000000)
    pts(7) = Array(-32.208947031361, 3.017653248961, 18.250000000000)
    pts(8) = Array(-32.441390193794, 2.947342412076, 18.250000000000)
    pts(9) = Array(-32.674020459634, 2.872000522875, 18.250000000000)
    pts(10) = Array(-33.271715955847, 2.924537117806, 18.250000000000)
    pts(11) = Array(-33.035135427058, 4.923395912982, 18.250000000000)
    pts(12) = Array(-32.441689880464, 4.834951675024, 18.250000000000)
    pts(13) = Array(-32.233075422124, 4.707385031166, 18.250000000000)
    pts(14) = Array(-32.023468421672, 4.584754109672, 18.250000000000)
    pts(15) = Array(-31.812883047026, 4.467224780134, 18.250000000000)
    pts(16) = Array(-31.601329778721, 4.354992102925, 18.250000000000)
    pts(17) = Array(-31.388814268692, 4.248289515272, 18.250000000000)
    pts(18) = Array(-31.175335644090, 4.147402497751, 18.250000000000)
    pts(19) = Array(-30.960883853107, 4.052689975114, 18.250000000000)
    FSData_368 = pts
End Function

Private Function FSData_369() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.961717818547, 1.540139514675, 15.150000000000)
    pts(1) = Array(-30.987111129080, 0.893836603677, 15.150000000000)
    pts(2) = Array(-31.213684384465, 0.840554666267, 15.150000000000)
    pts(3) = Array(-31.440312220850, 0.780555863149, 15.150000000000)
    pts(4) = Array(-31.666944775951, 0.714306346663, 15.150000000000)
    pts(5) = Array(-31.893535829620, 0.642162195051, 15.150000000000)
    pts(6) = Array(-32.120041584959, 0.564405510693, 15.150000000000)
    pts(7) = Array(-32.346419941990, 0.481265972658, 15.150000000000)
    pts(8) = Array(-32.572630035045, 0.392934600286, 15.150000000000)
    pts(9) = Array(-32.798631922915, 0.299573004689, 15.150000000000)
    pts(10) = Array(-33.398606897115, 0.305052998677, 15.150000000000)
    pts(11) = Array(-33.319584320698, 2.316311873629, 15.150000000000)
    pts(12) = Array(-32.721028913739, 2.274701480689, 15.150000000000)
    pts(13) = Array(-32.503048780631, 2.163895783978, 15.150000000000)
    pts(14) = Array(-32.284466417400, 2.058088468403, 15.150000000000)
    pts(15) = Array(-32.065308962267, 1.957443780632, 15.150000000000)
    pts(16) = Array(-31.845602167721, 1.862155357441, 15.150000000000)
    pts(17) = Array(-31.625369983547, 1.772455473001, 15.150000000000)
    pts(18) = Array(-31.404633938134, 1.688628796329, 15.150000000000)
    pts(19) = Array(-31.183412171574, 1.611033934388, 15.150000000000)
    FSData_369 = pts
End Function

Private Function FSData_370() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.961717818547, 1.540139514675, 18.250000000000)
    pts(1) = Array(-30.987111129080, 0.893836603677, 18.250000000000)
    pts(2) = Array(-31.213684384465, 0.840554666267, 18.250000000000)
    pts(3) = Array(-31.440312220850, 0.780555863149, 18.250000000000)
    pts(4) = Array(-31.666944775951, 0.714306346663, 18.250000000000)
    pts(5) = Array(-31.893535829620, 0.642162195051, 18.250000000000)
    pts(6) = Array(-32.120041584959, 0.564405510693, 18.250000000000)
    pts(7) = Array(-32.346419941990, 0.481265972658, 18.250000000000)
    pts(8) = Array(-32.572630035045, 0.392934600286, 18.250000000000)
    pts(9) = Array(-32.798631922915, 0.299573004689, 18.250000000000)
    pts(10) = Array(-33.398606897115, 0.305052998677, 18.250000000000)
    pts(11) = Array(-33.319584320698, 2.316311873629, 18.250000000000)
    pts(12) = Array(-32.721028913739, 2.274701480689, 18.250000000000)
    pts(13) = Array(-32.503048780631, 2.163895783978, 18.250000000000)
    pts(14) = Array(-32.284466417400, 2.058088468403, 18.250000000000)
    pts(15) = Array(-32.065308962267, 1.957443780632, 18.250000000000)
    pts(16) = Array(-31.845602167721, 1.862155357441, 18.250000000000)
    pts(17) = Array(-31.625369983547, 1.772455473001, 18.250000000000)
    pts(18) = Array(-31.404633938134, 1.688628796329, 18.250000000000)
    pts(19) = Array(-31.183412171574, 1.611033934388, 18.250000000000)
    FSData_370 = pts
End Function

Private Function FSData_371() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.987111129080, -0.893836603677, 15.150000000000)
    pts(1) = Array(-30.961717818547, -1.540139514675, 15.150000000000)
    pts(2) = Array(-31.183412171574, -1.611033934388, 15.150000000000)
    pts(3) = Array(-31.404633938134, -1.688628796329, 15.150000000000)
    pts(4) = Array(-31.625369983547, -1.772455473001, 15.150000000000)
    pts(5) = Array(-31.845602167721, -1.862155357441, 15.150000000000)
    pts(6) = Array(-32.065308962267, -1.957443780632, 15.150000000000)
    pts(7) = Array(-32.284466417400, -2.058088468403, 15.150000000000)
    pts(8) = Array(-32.503048780631, -2.163895783978, 15.150000000000)
    pts(9) = Array(-32.721028913739, -2.274701480689, 15.150000000000)
    pts(10) = Array(-33.319584320698, -2.316311873629, 15.150000000000)
    pts(11) = Array(-33.398606897115, -0.305052998677, 15.150000000000)
    pts(12) = Array(-32.798631922915, -0.299573004689, 15.150000000000)
    pts(13) = Array(-32.572630035045, -0.392934600286, 15.150000000000)
    pts(14) = Array(-32.346419941990, -0.481265972658, 15.150000000000)
    pts(15) = Array(-32.120041584959, -0.564405510693, 15.150000000000)
    pts(16) = Array(-31.893535829620, -0.642162195051, 15.150000000000)
    pts(17) = Array(-31.666944775951, -0.714306346663, 15.150000000000)
    pts(18) = Array(-31.440312220850, -0.780555863149, 15.150000000000)
    pts(19) = Array(-31.213684384465, -0.840554666267, 15.150000000000)
    FSData_371 = pts
End Function

Private Function FSData_372() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.987111129080, -0.893836603677, 18.250000000000)
    pts(1) = Array(-30.961717818547, -1.540139514675, 18.250000000000)
    pts(2) = Array(-31.183412171574, -1.611033934388, 18.250000000000)
    pts(3) = Array(-31.404633938134, -1.688628796329, 18.250000000000)
    pts(4) = Array(-31.625369983547, -1.772455473001, 18.250000000000)
    pts(5) = Array(-31.845602167721, -1.862155357441, 18.250000000000)
    pts(6) = Array(-32.065308962267, -1.957443780632, 18.250000000000)
    pts(7) = Array(-32.284466417400, -2.058088468403, 18.250000000000)
    pts(8) = Array(-32.503048780631, -2.163895783978, 18.250000000000)
    pts(9) = Array(-32.721028913739, -2.274701480689, 18.250000000000)
    pts(10) = Array(-33.319584320698, -2.316311873629, 18.250000000000)
    pts(11) = Array(-33.398606897115, -0.305052998677, 18.250000000000)
    pts(12) = Array(-32.798631922915, -0.299573004689, 18.250000000000)
    pts(13) = Array(-32.572630035045, -0.392934600286, 18.250000000000)
    pts(14) = Array(-32.346419941990, -0.481265972658, 18.250000000000)
    pts(15) = Array(-32.120041584959, -0.564405510693, 18.250000000000)
    pts(16) = Array(-31.893535829620, -0.642162195051, 18.250000000000)
    pts(17) = Array(-31.666944775951, -0.714306346663, 18.250000000000)
    pts(18) = Array(-31.440312220850, -0.780555863149, 18.250000000000)
    pts(19) = Array(-31.213684384465, -0.840554666267, 18.250000000000)
    FSData_372 = pts
End Function

Private Function FSData_373() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.821458595241, -3.322301922137, 15.150000000000)
    pts(1) = Array(-30.745435221847, -3.964620160771, 15.150000000000)
    pts(2) = Array(-30.960883853107, -4.052689975114, 15.150000000000)
    pts(3) = Array(-31.175335644090, -4.147402497751, 15.150000000000)
    pts(4) = Array(-31.388814268692, -4.248289515272, 15.150000000000)
    pts(5) = Array(-31.601329778721, -4.354992102925, 15.150000000000)
    pts(6) = Array(-31.812883047026, -4.467224780134, 15.150000000000)
    pts(7) = Array(-32.023468421672, -4.584754109672, 15.150000000000)
    pts(8) = Array(-32.233075422124, -4.707385031166, 15.150000000000)
    pts(9) = Array(-32.441689880464, -4.834951675024, 15.150000000000)
    pts(10) = Array(-33.035135427058, -4.923395912982, 15.150000000000)
    pts(11) = Array(-33.271715955847, -2.924537117806, 15.150000000000)
    pts(12) = Array(-32.674020459634, -2.872000522875, 15.150000000000)
    pts(13) = Array(-32.441390193794, -2.947342412076, 15.150000000000)
    pts(14) = Array(-32.208947031361, -3.017653248961, 15.150000000000)
    pts(15) = Array(-31.976743470282, -3.082775054361, 15.150000000000)
    pts(16) = Array(-31.744835237452, -3.142520604070, 15.150000000000)
    pts(17) = Array(-31.513282323483, -3.196664230169, 15.150000000000)
    pts(18) = Array(-31.282150523759, -3.244928136169, 15.150000000000)
    pts(19) = Array(-31.051513853523, -3.286960967891, 15.150000000000)
    FSData_373 = pts
End Function

Private Function FSData_374() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.821458595241, -3.322301922137, 18.250000000000)
    pts(1) = Array(-30.745435221847, -3.964620160771, 18.250000000000)
    pts(2) = Array(-30.960883853107, -4.052689975114, 18.250000000000)
    pts(3) = Array(-31.175335644090, -4.147402497751, 18.250000000000)
    pts(4) = Array(-31.388814268692, -4.248289515272, 18.250000000000)
    pts(5) = Array(-31.601329778721, -4.354992102925, 18.250000000000)
    pts(6) = Array(-31.812883047026, -4.467224780134, 18.250000000000)
    pts(7) = Array(-32.023468421672, -4.584754109672, 18.250000000000)
    pts(8) = Array(-32.233075422124, -4.707385031166, 18.250000000000)
    pts(9) = Array(-32.441689880464, -4.834951675024, 18.250000000000)
    pts(10) = Array(-33.035135427058, -4.923395912982, 18.250000000000)
    pts(11) = Array(-33.271715955847, -2.924537117806, 18.250000000000)
    pts(12) = Array(-32.674020459634, -2.872000522875, 18.250000000000)
    pts(13) = Array(-32.441390193794, -2.947342412076, 18.250000000000)
    pts(14) = Array(-32.208947031361, -3.017653248961, 18.250000000000)
    pts(15) = Array(-31.976743470282, -3.082775054361, 18.250000000000)
    pts(16) = Array(-31.744835237452, -3.142520604070, 18.250000000000)
    pts(17) = Array(-31.513282323483, -3.196664230169, 18.250000000000)
    pts(18) = Array(-31.282150523759, -3.244928136169, 18.250000000000)
    pts(19) = Array(-31.051513853523, -3.286960967891, 18.250000000000)
    FSData_374 = pts
End Function

Private Function FSData_375() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.465781519988, -5.730284144469, 15.150000000000)
    pts(1) = Array(-30.339596793110, -6.364657605205, 15.150000000000)
    pts(2) = Array(-30.547471390148, -6.469359834487, 15.150000000000)
    pts(3) = Array(-30.753831038948, -6.580606083623, 15.150000000000)
    pts(4) = Array(-30.958736076030, -6.697931439982, 15.150000000000)
    pts(5) = Array(-31.162224683126, -6.820978873913, 15.150000000000)
    pts(6) = Array(-31.364320128943, -6.949463853363, 15.150000000000)
    pts(7) = Array(-31.565035094241, -7.083153217270, 15.150000000000)
    pts(8) = Array(-31.764374435055, -7.221851684273, 15.150000000000)
    pts(9) = Array(-31.962337041119, -7.365392784498, 15.150000000000)
    pts(10) = Array(-32.547013938212, -7.500125579336, 15.150000000000)
    pts(11) = Array(-32.939693821744, -5.525990493096, 15.150000000000)
    pts(12) = Array(-32.347962795006, -5.426721202801, 15.150000000000)
    pts(13) = Array(-32.110138394141, -5.483578877805, 15.150000000000)
    pts(14) = Array(-31.872895251723, -5.535435689512, 15.150000000000)
    pts(15) = Array(-31.636298098764, -5.582138264692, 15.150000000000)
    pts(16) = Array(-31.400417179817, -5.623504328571, 15.150000000000)
    pts(17) = Array(-31.165330006262, -5.659313615697, 15.150000000000)
    pts(18) = Array(-30.931123966318, -5.689294348182, 15.150000000000)
    pts(19) = Array(-30.697900413998, -5.713102062123, 15.150000000000)
    FSData_375 = pts
End Function

Private Function FSData_376() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-30.465781519988, -5.730284144469, 18.250000000000)
    pts(1) = Array(-30.339596793110, -6.364657605205, 18.250000000000)
    pts(2) = Array(-30.547471390148, -6.469359834487, 18.250000000000)
    pts(3) = Array(-30.753831038948, -6.580606083623, 18.250000000000)
    pts(4) = Array(-30.958736076030, -6.697931439982, 18.250000000000)
    pts(5) = Array(-31.162224683126, -6.820978873913, 18.250000000000)
    pts(6) = Array(-31.364320128943, -6.949463853363, 18.250000000000)
    pts(7) = Array(-31.565035094241, -7.083153217270, 18.250000000000)
    pts(8) = Array(-31.764374435055, -7.221851684273, 18.250000000000)
    pts(9) = Array(-31.962337041119, -7.365392784498, 18.250000000000)
    pts(10) = Array(-32.547013938212, -7.500125579336, 18.250000000000)
    pts(11) = Array(-32.939693821744, -5.525990493096, 18.250000000000)
    pts(12) = Array(-32.347962795006, -5.426721202801, 18.250000000000)
    pts(13) = Array(-32.110138394141, -5.483578877805, 18.250000000000)
    pts(14) = Array(-31.872895251723, -5.535435689512, 18.250000000000)
    pts(15) = Array(-31.636298098764, -5.582138264692, 18.250000000000)
    pts(16) = Array(-31.400417179817, -5.623504328571, 18.250000000000)
    pts(17) = Array(-31.165330006262, -5.659313615697, 18.250000000000)
    pts(18) = Array(-30.931123966318, -5.689294348182, 18.250000000000)
    pts(19) = Array(-30.697900413998, -5.713102062123, 18.250000000000)
    FSData_376 = pts
End Function

Private Function FSData_377() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-29.922272770764, -8.102937259537, 15.150000000000)
    pts(1) = Array(-29.746704661204, -8.725454819040, 15.150000000000)
    pts(2) = Array(-29.945723608004, -8.846143939200, 15.150000000000)
    pts(3) = Array(-30.142718838765, -8.973238044717, 15.150000000000)
    pts(4) = Array(-30.337786980636, -9.106278390077, 15.150000000000)
    pts(5) = Array(-30.530994109869, -9.244912041937, 15.150000000000)
    pts(6) = Array(-30.722385747570, -9.388857171005, 15.150000000000)
    pts(7) = Array(-30.911992829015, -9.537882329896, 15.150000000000)
    pts(8) = Array(-31.099835516867, -9.691793220237, 15.150000000000)
    pts(9) = Array(-31.285925765359, -9.850423798214, 15.150000000000)
    pts(10) = Array(-31.858229285457, -10.030614477449, 15.150000000000)
    pts(11) = Array(-32.404587521669, -8.093374299418, 15.150000000000)
    pts(12) = Array(-31.822469182957, -7.947984341944, 15.150000000000)
    pts(13) = Array(-31.580916913584, -7.986007256277, 15.150000000000)
    pts(14) = Array(-31.340336474044, -8.019090328318, 15.150000000000)
    pts(15) = Array(-31.100804429331, -8.047085736372, 15.150000000000)
    pts(16) = Array(-30.862405108570, -8.069817278881, 15.150000000000)
    pts(17) = Array(-30.625233066028, -8.087071450871, 15.150000000000)
    pts(18) = Array(-30.389396743983, -8.098584168655, 15.150000000000)
    pts(19) = Array(-30.155024210333, -8.104019982343, 15.150000000000)
    FSData_377 = pts
End Function

Private Function FSData_378() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-29.922272770764, -8.102937259537, 18.250000000000)
    pts(1) = Array(-29.746704661204, -8.725454819040, 18.250000000000)
    pts(2) = Array(-29.945723608004, -8.846143939200, 18.250000000000)
    pts(3) = Array(-30.142718838765, -8.973238044717, 18.250000000000)
    pts(4) = Array(-30.337786980636, -9.106278390077, 18.250000000000)
    pts(5) = Array(-30.530994109869, -9.244912041937, 18.250000000000)
    pts(6) = Array(-30.722385747570, -9.388857171005, 18.250000000000)
    pts(7) = Array(-30.911992829015, -9.537882329896, 18.250000000000)
    pts(8) = Array(-31.099835516867, -9.691793220237, 18.250000000000)
    pts(9) = Array(-31.285925765359, -9.850423798214, 18.250000000000)
    pts(10) = Array(-31.858229285457, -10.030614477449, 18.250000000000)
    pts(11) = Array(-32.404587521669, -8.093374299418, 18.250000000000)
    pts(12) = Array(-31.822469182957, -7.947984341944, 18.250000000000)
    pts(13) = Array(-31.580916913584, -7.986007256277, 18.250000000000)
    pts(14) = Array(-31.340336474044, -8.019090328318, 18.250000000000)
    pts(15) = Array(-31.100804429331, -8.047085736372, 18.250000000000)
    pts(16) = Array(-30.862405108570, -8.069817278881, 18.250000000000)
    pts(17) = Array(-30.625233066028, -8.087071450871, 18.250000000000)
    pts(18) = Array(-30.389396743983, -8.098584168655, 18.250000000000)
    pts(19) = Array(-30.155024210333, -8.104019982343, 18.250000000000)
    FSData_378 = pts
End Function

Private Function FSData_379() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-29.194283259743, -10.425633071901, 15.150000000000)
    pts(1) = Array(-28.970414203278, -11.032456702408, 15.150000000000)
    pts(2) = Array(-29.159350481852, -11.168388624887, 15.150000000000)
    pts(3) = Array(-29.345766753470, -11.310547009360, 15.150000000000)
    pts(4) = Array(-29.529795340167, -11.458482105752, 15.150000000000)
    pts(5) = Array(-29.711529805339, -11.611847252978, 15.150000000000)
    pts(6) = Array(-29.891037641833, -11.770365062076, 15.150000000000)
    pts(7) = Array(-30.068367848717, -11.933807226290, 15.150000000000)
    pts(8) = Array(-30.243555770973, -12.101981628149, 15.150000000000)
    pts(9) = Array(-30.416626353631, -12.274723673615, 15.150000000000)
    pts(10) = Array(-30.973028055222, -12.499261301791, 15.150000000000)
    pts(11) = Array(-31.669696163905, -10.610859761864, 15.150000000000)
    pts(12) = Array(-31.100779466350, -10.420245514645, 15.150000000000)
    pts(13) = Array(-30.856988578534, -10.439199244397, 15.150000000000)
    pts(14) = Array(-30.614554100283, -10.453304608631, 15.150000000000)
    pts(15) = Array(-30.373563958524, -10.462420248560, 15.150000000000)
    pts(16) = Array(-30.134116047037, -10.466377122181, 15.150000000000)
    pts(17) = Array(-29.896321380018, -10.464969801327, 15.150000000000)
    pts(18) = Array(-29.660308785221, -10.457943524677, 15.150000000000)
    pts(19) = Array(-29.426232254848, -10.444973924511, 15.150000000000)
    FSData_379 = pts
End Function

Private Function FSData_380() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-29.194283259743, -10.425633071901, 18.250000000000)
    pts(1) = Array(-28.970414203278, -11.032456702408, 18.250000000000)
    pts(2) = Array(-29.159350481852, -11.168388624887, 18.250000000000)
    pts(3) = Array(-29.345766753470, -11.310547009360, 18.250000000000)
    pts(4) = Array(-29.529795340167, -11.458482105752, 18.250000000000)
    pts(5) = Array(-29.711529805339, -11.611847252978, 18.250000000000)
    pts(6) = Array(-29.891037641833, -11.770365062076, 18.250000000000)
    pts(7) = Array(-30.068367848717, -11.933807226290, 18.250000000000)
    pts(8) = Array(-30.243555770973, -12.101981628149, 18.250000000000)
    pts(9) = Array(-30.416626353631, -12.274723673615, 18.250000000000)
    pts(10) = Array(-30.973028055222, -12.499261301791, 18.250000000000)
    pts(11) = Array(-31.669696163905, -10.610859761864, 18.250000000000)
    pts(12) = Array(-31.100779466350, -10.420245514645, 18.250000000000)
    pts(13) = Array(-30.856988578534, -10.439199244397, 18.250000000000)
    pts(14) = Array(-30.614554100283, -10.453304608631, 18.250000000000)
    pts(15) = Array(-30.373563958524, -10.462420248560, 18.250000000000)
    pts(16) = Array(-30.134116047037, -10.466377122181, 18.250000000000)
    pts(17) = Array(-29.896321380018, -10.464969801327, 18.250000000000)
    pts(18) = Array(-29.660308785221, -10.457943524677, 18.250000000000)
    pts(19) = Array(-29.426232254848, -10.444973924511, 18.250000000000)
    FSData_380 = pts
End Function

Private Function FSData_381() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-28.286301284341, -12.684051389501, 15.150000000000)
    pts(1) = Array(-28.015511508149, -13.271439821541, 15.150000000000)
    pts(2) = Array(-28.193200263512, -13.421776480835, 15.150000000000)
    pts(3) = Array(-28.367888257682, -13.578122690552, 15.150000000000)
    pts(4) = Array(-28.539742691774, -13.740040468913, 15.150000000000)
    pts(5) = Array(-28.708884039472, -13.907191564372, 15.150000000000)
    pts(6) = Array(-28.875401349256, -14.079304738495, 15.150000000000)
    pts(7) = Array(-29.039361381885, -14.256156232741, 15.150000000000)
    pts(8) = Array(-29.200814446748, -14.437557295007, 15.150000000000)
    pts(9) = Array(-29.359798325877, -14.623345795809, 15.150000000000)
    pts(10) = Array(-29.896867807448, -14.890846023781, 15.150000000000)
    pts(11) = Array(-30.739550598048, -13.062925745409, 15.150000000000)
    pts(12) = Array(-30.187343102274, -12.828262408665, 15.150000000000)
    pts(13) = Array(-29.942816647908, -12.828030097789, 15.150000000000)
    pts(14) = Array(-29.700022820122, -12.823070789957, 15.150000000000)
    pts(15) = Array(-29.459060365677, -12.813250460808, 15.150000000000)
    pts(16) = Array(-29.220040139464, -12.798408270098, 15.150000000000)
    pts(17) = Array(-28.983088931165, -12.778348133001, 15.150000000000)
    pts(18) = Array(-28.748355159745, -12.752826181250, 15.150000000000)
    pts(19) = Array(-28.516017792297, -12.721531129129, 15.150000000000)
    FSData_381 = pts
End Function

Private Function FSData_382() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-28.286301284341, -12.684051389501, 18.250000000000)
    pts(1) = Array(-28.015511508149, -13.271439821541, 18.250000000000)
    pts(2) = Array(-28.193200263512, -13.421776480835, 18.250000000000)
    pts(3) = Array(-28.367888257682, -13.578122690552, 18.250000000000)
    pts(4) = Array(-28.539742691774, -13.740040468913, 18.250000000000)
    pts(5) = Array(-28.708884039472, -13.907191564372, 18.250000000000)
    pts(6) = Array(-28.875401349256, -14.079304738495, 18.250000000000)
    pts(7) = Array(-29.039361381885, -14.256156232741, 18.250000000000)
    pts(8) = Array(-29.200814446748, -14.437557295007, 18.250000000000)
    pts(9) = Array(-29.359798325877, -14.623345795809, 18.250000000000)
    pts(10) = Array(-29.896867807448, -14.890846023781, 18.250000000000)
    pts(11) = Array(-30.739550598048, -13.062925745409, 18.250000000000)
    pts(12) = Array(-30.187343102274, -12.828262408665, 18.250000000000)
    pts(13) = Array(-29.942816647908, -12.828030097789, 18.250000000000)
    pts(14) = Array(-29.700022820122, -12.823070789957, 18.250000000000)
    pts(15) = Array(-29.459060365677, -12.813250460808, 18.250000000000)
    pts(16) = Array(-29.220040139464, -12.798408270098, 18.250000000000)
    pts(17) = Array(-28.983088931165, -12.778348133001, 18.250000000000)
    pts(18) = Array(-28.748355159745, -12.752826181250, 18.250000000000)
    pts(19) = Array(-28.516017792297, -12.721531129129, 18.250000000000)
    FSData_382 = pts
End Function

Private Function FSData_383() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-27.203924855371, -14.864268312410, 15.150000000000)
    pts(1) = Array(-26.887883868469, -15.428600100973, 15.150000000000)
    pts(2) = Array(-27.053229590357, -15.592414621585, 15.150000000000)
    pts(3) = Array(-27.215112297505, -15.761984730172, 15.150000000000)
    pts(4) = Array(-27.373733039258, -15.936886913556, 15.150000000000)
    pts(5) = Array(-27.529238456829, -16.116793415163, 15.150000000000)
    pts(6) = Array(-27.681738605315, -16.301440819358, 15.150000000000)
    pts(7) = Array(-27.831317595565, -16.490611295364, 15.150000000000)
    pts(8) = Array(-27.978040391202, -16.684120620166, 15.150000000000)
    pts(9) = Array(-28.121957378321, -16.881810128416, 15.150000000000)
    pts(10) = Array(-28.636383427925, -17.190623728326, 15.150000000000)
    pts(11) = Array(-29.619885480815, -15.434454447870, 15.150000000000)
    pts(12) = Array(-29.087791729663, -15.157188799106, 15.150000000000)
    pts(13) = Array(-28.844037295650, -15.137771879876, 15.150000000000)
    pts(14) = Array(-28.602381023015, -15.113778475758, 15.150000000000)
    pts(15) = Array(-28.362931869545, -15.085082723126, 15.150000000000)
    pts(16) = Array(-28.125812967782, -15.051532975128, 15.150000000000)
    pts(17) = Array(-27.891166101196, -15.012943699204, 15.150000000000)
    pts(18) = Array(-27.659158364911, -14.969083423670, 15.150000000000)
    pts(19) = Array(-27.429992597518, -14.919655863997, 15.150000000000)
    FSData_383 = pts
End Function

Private Function FSData_384() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-27.203924855371, -14.864268312410, 18.250000000000)
    pts(1) = Array(-26.887883868469, -15.428600100973, 18.250000000000)
    pts(2) = Array(-27.053229590357, -15.592414621585, 18.250000000000)
    pts(3) = Array(-27.215112297505, -15.761984730172, 18.250000000000)
    pts(4) = Array(-27.373733039258, -15.936886913556, 18.250000000000)
    pts(5) = Array(-27.529238456829, -16.116793415163, 18.250000000000)
    pts(6) = Array(-27.681738605315, -16.301440819358, 18.250000000000)
    pts(7) = Array(-27.831317595565, -16.490611295364, 18.250000000000)
    pts(8) = Array(-27.978040391202, -16.684120620166, 18.250000000000)
    pts(9) = Array(-28.121957378321, -16.881810128416, 18.250000000000)
    pts(10) = Array(-28.636383427925, -17.190623728326, 18.250000000000)
    pts(11) = Array(-29.619885480815, -15.434454447870, 18.250000000000)
    pts(12) = Array(-29.087791729663, -15.157188799106, 18.250000000000)
    pts(13) = Array(-28.844037295650, -15.137771879876, 18.250000000000)
    pts(14) = Array(-28.602381023015, -15.113778475758, 18.250000000000)
    pts(15) = Array(-28.362931869545, -15.085082723126, 18.250000000000)
    pts(16) = Array(-28.125812967782, -15.051532975128, 18.250000000000)
    pts(17) = Array(-27.891166101196, -15.012943699204, 18.250000000000)
    pts(18) = Array(-27.659158364911, -14.969083423670, 18.250000000000)
    pts(19) = Array(-27.429992597518, -14.919655863997, 18.250000000000)
    FSData_384 = pts
End Function

Private Function FSData_385() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-25.953827183445, -16.952842078303, 15.150000000000)
    pts(1) = Array(-25.594483483611, -17.490637930252, 15.150000000000)
    pts(2) = Array(-25.746466760666, -17.666920341190, 15.150000000000)
    pts(3) = Array(-25.894546120071, -17.848668892540, 15.150000000000)
    pts(4) = Array(-26.038955219865, -18.035477150824, 15.150000000000)
    pts(5) = Array(-26.179865964700, -18.227029875171, 15.150000000000)
    pts(6) = Array(-26.317408737760, -18.423073096790, 15.150000000000)
    pts(7) = Array(-26.451684481417, -18.623396255666, 15.150000000000)
    pts(8) = Array(-26.582772413003, -18.827820793669, 15.150000000000)
    pts(9) = Array(-26.710735212028, -19.036192487809, 15.150000000000)
    pts(10) = Array(-27.199346222004, -19.384415521123, 15.150000000000)
    pts(11) = Array(-28.317603919982, -17.710824606183, 15.150000000000)
    pts(12) = Array(-27.808904448366, -17.392666080324, 15.150000000000)
    pts(13) = Array(-27.567424861849, -17.354184264503, 15.150000000000)
    pts(14) = Array(-27.328396035644, -17.311304691415, 15.150000000000)
    pts(15) = Array(-27.091936466805, -17.263910434158, 15.150000000000)
    pts(16) = Array(-26.858180806371, -17.211859974224, 15.150000000000)
    pts(17) = Array(-26.627284957460, -17.154979475191, 15.150000000000)
    pts(18) = Array(-26.399433661153, -17.093051289058, 15.150000000000)
    pts(19) = Array(-26.174852376978, -17.025795959174, 15.150000000000)
    FSData_385 = pts
End Function

Private Function FSData_386() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-25.953827183445, -16.952842078303, 18.250000000000)
    pts(1) = Array(-25.594483483611, -17.490637930252, 18.250000000000)
    pts(2) = Array(-25.746466760666, -17.666920341190, 18.250000000000)
    pts(3) = Array(-25.894546120071, -17.848668892540, 18.250000000000)
    pts(4) = Array(-26.038955219865, -18.035477150824, 18.250000000000)
    pts(5) = Array(-26.179865964700, -18.227029875171, 18.250000000000)
    pts(6) = Array(-26.317408737760, -18.423073096790, 18.250000000000)
    pts(7) = Array(-26.451684481417, -18.623396255666, 18.250000000000)
    pts(8) = Array(-26.582772413003, -18.827820793669, 18.250000000000)
    pts(9) = Array(-26.710735212028, -19.036192487809, 18.250000000000)
    pts(10) = Array(-27.199346222004, -19.384415521123, 18.250000000000)
    pts(11) = Array(-28.317603919982, -17.710824606183, 18.250000000000)
    pts(12) = Array(-27.808904448366, -17.392666080324, 18.250000000000)
    pts(13) = Array(-27.567424861849, -17.354184264503, 18.250000000000)
    pts(14) = Array(-27.328396035644, -17.311304691415, 18.250000000000)
    pts(15) = Array(-27.091936466805, -17.263910434158, 18.250000000000)
    pts(16) = Array(-26.858180806371, -17.211859974224, 18.250000000000)
    pts(17) = Array(-26.627284957460, -17.154979475191, 18.250000000000)
    pts(18) = Array(-26.399433661153, -17.093051289058, 18.250000000000)
    pts(19) = Array(-26.174852376978, -17.025795959174, 18.250000000000)
    FSData_386 = pts
End Function

Private Function FSData_387() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-24.543715536410, -18.936895935390, 15.150000000000)
    pts(1) = Array(-24.143284597047, -19.444840160463, 15.150000000000)
    pts(2) = Array(-24.280968401826, -19.632503622043, 15.150000000000)
    pts(3) = Array(-24.414331454996, -19.825310075900, 15.150000000000)
    pts(4) = Array(-24.543638582710, -20.022872674052, 15.150000000000)
    pts(5) = Array(-24.669085893209, -20.224890634896, 15.150000000000)
    pts(6) = Array(-24.790823293910, -20.431121002287, 15.150000000000)
    pts(7) = Array(-24.908967936364, -20.641361785144, 15.150000000000)
    pts(8) = Array(-25.023612803208, -20.855441191093, 15.150000000000)
    pts(9) = Array(-25.134832480932, -21.073210390339, 15.150000000000)
    pts(10) = Array(-25.594616001925, -21.458695946260, 15.150000000000)
    pts(11) = Array(-26.840734914424, -19.878001641353, 15.150000000000)
    pts(12) = Array(-26.358566023746, -19.520911791508, 15.150000000000)
    pts(13) = Array(-26.120850086675, -19.463602332288, 15.150000000000)
    pts(14) = Array(-25.885922399100, -19.402100957056, 15.150000000000)
    pts(15) = Array(-25.653910266764, -19.336300396530, 15.150000000000)
    pts(16) = Array(-25.424959029038, -19.266070133054, 15.150000000000)
    pts(17) = Array(-25.199237747490, -19.191249098104, 15.150000000000)
    pts(18) = Array(-24.976947670172, -19.111634809233, 15.150000000000)
    pts(19) = Array(-24.758335487511, -19.026966360612, 15.150000000000)
    FSData_387 = pts
End Function

Private Function FSData_388() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-24.543715536410, -18.936895935390, 18.250000000000)
    pts(1) = Array(-24.143284597047, -19.444840160463, 18.250000000000)
    pts(2) = Array(-24.280968401826, -19.632503622043, 18.250000000000)
    pts(3) = Array(-24.414331454996, -19.825310075900, 18.250000000000)
    pts(4) = Array(-24.543638582710, -20.022872674052, 18.250000000000)
    pts(5) = Array(-24.669085893209, -20.224890634896, 18.250000000000)
    pts(6) = Array(-24.790823293910, -20.431121002287, 18.250000000000)
    pts(7) = Array(-24.908967936364, -20.641361785144, 18.250000000000)
    pts(8) = Array(-25.023612803208, -20.855441191093, 18.250000000000)
    pts(9) = Array(-25.134832480932, -21.073210390339, 18.250000000000)
    pts(10) = Array(-25.594616001925, -21.458695946260, 18.250000000000)
    pts(11) = Array(-26.840734914424, -19.878001641353, 18.250000000000)
    pts(12) = Array(-26.358566023746, -19.520911791508, 18.250000000000)
    pts(13) = Array(-26.120850086675, -19.463602332288, 18.250000000000)
    pts(14) = Array(-25.885922399100, -19.402100957056, 18.250000000000)
    pts(15) = Array(-25.653910266764, -19.336300396530, 18.250000000000)
    pts(16) = Array(-25.424959029038, -19.266070133054, 18.250000000000)
    pts(17) = Array(-25.199237747490, -19.191249098104, 18.250000000000)
    pts(18) = Array(-24.976947670172, -19.111634809233, 18.250000000000)
    pts(19) = Array(-24.758335487511, -19.026966360612, 18.250000000000)
    FSData_388 = pts
End Function

Private Function FSData_389() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-22.982283721479, -20.804197531879, 15.150000000000)
    pts(1) = Array(-22.543234332486, -21.279158485020, 15.150000000000)
    pts(2) = Array(-22.665769798547, -21.477045989597, 15.150000000000)
    pts(3) = Array(-22.783594317913, -21.679721630058, 15.150000000000)
    pts(4) = Array(-22.897002252105, -21.886820528965, 15.150000000000)
    pts(5) = Array(-23.006212703883, -22.098058218398, 15.150000000000)
    pts(6) = Array(-23.111394180669, -22.313204252768, 15.150000000000)
    pts(7) = Array(-23.212679320910, -22.532066455268, 15.150000000000)
    pts(8) = Array(-23.310174299286, -22.754480858436, 15.150000000000)
    pts(9) = Array(-23.403965149411, -22.980304943259, 15.150000000000)
    pts(10) = Array(-23.832086463120, -23.400676375148, 15.150000000000)
    pts(11) = Array(-25.198383852668, -21.922624186296, 15.150000000000)
    pts(12) = Array(-24.745718274477, -21.528804590135, 15.150000000000)
    pts(13) = Array(-24.513231584653, -21.453020819389, 15.150000000000)
    pts(14) = Array(-24.283853443022, -21.373276818444, 15.150000000000)
    pts(15) = Array(-24.057719179137, -21.289475636985, 15.150000000000)
    pts(16) = Array(-23.834983924634, -21.201498562895, 15.150000000000)
    pts(17) = Array(-23.615828857210, -21.109198288589, 15.150000000000)
    pts(18) = Array(-23.400470491127, -21.012388745545, 15.150000000000)
    pts(19) = Array(-23.189175226782, -20.910829187328, 15.150000000000)
    FSData_389 = pts
End Function

Private Function FSData_390() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-22.982283721479, -20.804197531879, 18.250000000000)
    pts(1) = Array(-22.543234332486, -21.279158485020, 18.250000000000)
    pts(2) = Array(-22.665769798547, -21.477045989597, 18.250000000000)
    pts(3) = Array(-22.783594317913, -21.679721630058, 18.250000000000)
    pts(4) = Array(-22.897002252105, -21.886820528965, 18.250000000000)
    pts(5) = Array(-23.006212703883, -22.098058218398, 18.250000000000)
    pts(6) = Array(-23.111394180669, -22.313204252768, 18.250000000000)
    pts(7) = Array(-23.212679320910, -22.532066455268, 18.250000000000)
    pts(8) = Array(-23.310174299286, -22.754480858436, 18.250000000000)
    pts(9) = Array(-23.403965149411, -22.980304943259, 18.250000000000)
    pts(10) = Array(-23.832086463120, -23.400676375148, 18.250000000000)
    pts(11) = Array(-25.198383852668, -21.922624186296, 18.250000000000)
    pts(12) = Array(-24.745718274477, -21.528804590135, 18.250000000000)
    pts(13) = Array(-24.513231584653, -21.453020819389, 18.250000000000)
    pts(14) = Array(-24.283853443022, -21.373276818444, 18.250000000000)
    pts(15) = Array(-24.057719179137, -21.289475636985, 18.250000000000)
    pts(16) = Array(-23.834983924634, -21.201498562895, 18.250000000000)
    pts(17) = Array(-23.615828857210, -21.109198288589, 18.250000000000)
    pts(18) = Array(-23.400470491127, -21.012388745545, 18.250000000000)
    pts(19) = Array(-23.189175226782, -20.910829187328, 18.250000000000)
    FSData_390 = pts
End Function

Private Function FSData_391() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-21.279158485020, -22.543234332486, 15.150000000000)
    pts(1) = Array(-20.804197531879, -22.982283721479, 15.150000000000)
    pts(2) = Array(-20.910829187328, -23.189175226782, 15.150000000000)
    pts(3) = Array(-21.012388745545, -23.400470491127, 15.150000000000)
    pts(4) = Array(-21.109198288589, -23.615828857210, 15.150000000000)
    pts(5) = Array(-21.201498562895, -23.834983924634, 15.150000000000)
    pts(6) = Array(-21.289475636985, -24.057719179137, 15.150000000000)
    pts(7) = Array(-21.373276818444, -24.283853443022, 15.150000000000)
    pts(8) = Array(-21.453020819389, -24.513231584653, 15.150000000000)
    pts(9) = Array(-21.528804590135, -24.745718274477, 15.150000000000)
    pts(10) = Array(-21.922624186296, -25.198383852668, 15.150000000000)
    pts(11) = Array(-23.400676375148, -23.832086463120, 15.150000000000)
    pts(12) = Array(-22.980304943259, -23.403965149411, 15.150000000000)
    pts(13) = Array(-22.754480858436, -23.310174299286, 15.150000000000)
    pts(14) = Array(-22.532066455268, -23.212679320910, 15.150000000000)
    pts(15) = Array(-22.313204252768, -23.111394180669, 15.150000000000)
    pts(16) = Array(-22.098058218398, -23.006212703883, 15.150000000000)
    pts(17) = Array(-21.886820528965, -22.897002252105, 15.150000000000)
    pts(18) = Array(-21.679721630058, -22.783594317912, 15.150000000000)
    pts(19) = Array(-21.477045989597, -22.665769798547, 15.150000000000)
    FSData_391 = pts
End Function

Private Function FSData_392() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-21.279158485020, -22.543234332486, 18.250000000000)
    pts(1) = Array(-20.804197531879, -22.982283721479, 18.250000000000)
    pts(2) = Array(-20.910829187328, -23.189175226782, 18.250000000000)
    pts(3) = Array(-21.012388745545, -23.400470491127, 18.250000000000)
    pts(4) = Array(-21.109198288589, -23.615828857210, 18.250000000000)
    pts(5) = Array(-21.201498562895, -23.834983924634, 18.250000000000)
    pts(6) = Array(-21.289475636985, -24.057719179137, 18.250000000000)
    pts(7) = Array(-21.373276818444, -24.283853443022, 18.250000000000)
    pts(8) = Array(-21.453020819389, -24.513231584653, 18.250000000000)
    pts(9) = Array(-21.528804590135, -24.745718274477, 18.250000000000)
    pts(10) = Array(-21.922624186296, -25.198383852668, 18.250000000000)
    pts(11) = Array(-23.400676375148, -23.832086463120, 18.250000000000)
    pts(12) = Array(-22.980304943259, -23.403965149411, 18.250000000000)
    pts(13) = Array(-22.754480858436, -23.310174299286, 18.250000000000)
    pts(14) = Array(-22.532066455268, -23.212679320910, 18.250000000000)
    pts(15) = Array(-22.313204252768, -23.111394180669, 18.250000000000)
    pts(16) = Array(-22.098058218398, -23.006212703883, 18.250000000000)
    pts(17) = Array(-21.886820528965, -22.897002252105, 18.250000000000)
    pts(18) = Array(-21.679721630058, -22.783594317912, 18.250000000000)
    pts(19) = Array(-21.477045989597, -22.665769798547, 18.250000000000)
    FSData_392 = pts
End Function

Private Function FSData_393() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-19.444840160463, -24.143284597047, 15.150000000000)
    pts(1) = Array(-18.936895935390, -24.543715536410, 15.150000000000)
    pts(2) = Array(-19.026966360612, -24.758335487511, 15.150000000000)
    pts(3) = Array(-19.111634809233, -24.976947670172, 15.150000000000)
    pts(4) = Array(-19.191249098104, -25.199237747490, 15.150000000000)
    pts(5) = Array(-19.266070133054, -25.424959029038, 15.150000000000)
    pts(6) = Array(-19.336300396530, -25.653910266764, 15.150000000000)
    pts(7) = Array(-19.402100957056, -25.885922399100, 15.150000000000)
    pts(8) = Array(-19.463602332288, -26.120850086675, 15.150000000000)
    pts(9) = Array(-19.520911791508, -26.358566023746, 15.150000000000)
    pts(10) = Array(-19.878001641353, -26.840734914424, 15.150000000000)
    pts(11) = Array(-21.458695946260, -25.594616001925, 15.150000000000)
    pts(12) = Array(-21.073210390339, -25.134832480932, 15.150000000000)
    pts(13) = Array(-20.855441191093, -25.023612803208, 15.150000000000)
    pts(14) = Array(-20.641361785144, -24.908967936364, 15.150000000000)
    pts(15) = Array(-20.431121002287, -24.790823293910, 15.150000000000)
    pts(16) = Array(-20.224890634896, -24.669085893209, 15.150000000000)
    pts(17) = Array(-20.022872674052, -24.543638582710, 15.150000000000)
    pts(18) = Array(-19.825310075900, -24.414331454996, 15.150000000000)
    pts(19) = Array(-19.632503622043, -24.280968401826, 15.150000000000)
    FSData_393 = pts
End Function

Private Function FSData_394() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-19.444840160463, -24.143284597047, 18.250000000000)
    pts(1) = Array(-18.936895935390, -24.543715536410, 18.250000000000)
    pts(2) = Array(-19.026966360612, -24.758335487511, 18.250000000000)
    pts(3) = Array(-19.111634809233, -24.976947670172, 18.250000000000)
    pts(4) = Array(-19.191249098104, -25.199237747490, 18.250000000000)
    pts(5) = Array(-19.266070133054, -25.424959029038, 18.250000000000)
    pts(6) = Array(-19.336300396530, -25.653910266764, 18.250000000000)
    pts(7) = Array(-19.402100957056, -25.885922399100, 18.250000000000)
    pts(8) = Array(-19.463602332288, -26.120850086675, 18.250000000000)
    pts(9) = Array(-19.520911791508, -26.358566023746, 18.250000000000)
    pts(10) = Array(-19.878001641353, -26.840734914424, 18.250000000000)
    pts(11) = Array(-21.458695946260, -25.594616001925, 18.250000000000)
    pts(12) = Array(-21.073210390339, -25.134832480932, 18.250000000000)
    pts(13) = Array(-20.855441191093, -25.023612803208, 18.250000000000)
    pts(14) = Array(-20.641361785144, -24.908967936364, 18.250000000000)
    pts(15) = Array(-20.431121002287, -24.790823293910, 18.250000000000)
    pts(16) = Array(-20.224890634896, -24.669085893209, 18.250000000000)
    pts(17) = Array(-20.022872674052, -24.543638582710, 18.250000000000)
    pts(18) = Array(-19.825310075900, -24.414331454996, 18.250000000000)
    pts(19) = Array(-19.632503622043, -24.280968401826, 18.250000000000)
    FSData_394 = pts
End Function

Private Function FSData_395() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-17.490637930252, -25.594483483611, 15.150000000000)
    pts(1) = Array(-16.952842078303, -25.953827183445, 15.150000000000)
    pts(2) = Array(-17.025795959174, -26.174852376978, 15.150000000000)
    pts(3) = Array(-17.093051289058, -26.399433661153, 15.150000000000)
    pts(4) = Array(-17.154979475191, -26.627284957460, 15.150000000000)
    pts(5) = Array(-17.211859974224, -26.858180806371, 15.150000000000)
    pts(6) = Array(-17.263910434158, -27.091936466805, 15.150000000000)
    pts(7) = Array(-17.311304691415, -27.328396035644, 15.150000000000)
    pts(8) = Array(-17.354184264503, -27.567424861849, 15.150000000000)
    pts(9) = Array(-17.392666080324, -27.808904448366, 15.150000000000)
    pts(10) = Array(-17.710824606183, -28.317603919982, 15.150000000000)
    pts(11) = Array(-19.384415521123, -27.199346222004, 15.150000000000)
    pts(12) = Array(-19.036192487809, -26.710735212028, 15.150000000000)
    pts(13) = Array(-18.827820793669, -26.582772413003, 15.150000000000)
    pts(14) = Array(-18.623396255666, -26.451684481417, 15.150000000000)
    pts(15) = Array(-18.423073096790, -26.317408737760, 15.150000000000)
    pts(16) = Array(-18.227029875171, -26.179865964700, 15.150000000000)
    pts(17) = Array(-18.035477150824, -26.038955219865, 15.150000000000)
    pts(18) = Array(-17.848668892540, -25.894546120071, 15.150000000000)
    pts(19) = Array(-17.666920341190, -25.746466760666, 15.150000000000)
    FSData_395 = pts
End Function

Private Function FSData_396() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-17.490637930252, -25.594483483611, 18.250000000000)
    pts(1) = Array(-16.952842078303, -25.953827183445, 18.250000000000)
    pts(2) = Array(-17.025795959174, -26.174852376978, 18.250000000000)
    pts(3) = Array(-17.093051289058, -26.399433661153, 18.250000000000)
    pts(4) = Array(-17.154979475191, -26.627284957460, 18.250000000000)
    pts(5) = Array(-17.211859974224, -26.858180806371, 18.250000000000)
    pts(6) = Array(-17.263910434158, -27.091936466805, 18.250000000000)
    pts(7) = Array(-17.311304691415, -27.328396035644, 18.250000000000)
    pts(8) = Array(-17.354184264503, -27.567424861849, 18.250000000000)
    pts(9) = Array(-17.392666080324, -27.808904448366, 18.250000000000)
    pts(10) = Array(-17.710824606183, -28.317603919982, 18.250000000000)
    pts(11) = Array(-19.384415521123, -27.199346222004, 18.250000000000)
    pts(12) = Array(-19.036192487809, -26.710735212028, 18.250000000000)
    pts(13) = Array(-18.827820793669, -26.582772413003, 18.250000000000)
    pts(14) = Array(-18.623396255666, -26.451684481417, 18.250000000000)
    pts(15) = Array(-18.423073096790, -26.317408737760, 18.250000000000)
    pts(16) = Array(-18.227029875171, -26.179865964700, 18.250000000000)
    pts(17) = Array(-18.035477150824, -26.038955219865, 18.250000000000)
    pts(18) = Array(-17.848668892540, -25.894546120071, 18.250000000000)
    pts(19) = Array(-17.666920341190, -25.746466760666, 18.250000000000)
    FSData_396 = pts
End Function

Private Function FSData_397() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-15.428600100973, -26.887883868469, 15.150000000000)
    pts(1) = Array(-14.864268312410, -27.203924855371, 15.150000000000)
    pts(2) = Array(-14.919655863997, -27.429992597518, 15.150000000000)
    pts(3) = Array(-14.969083423670, -27.659158364911, 15.150000000000)
    pts(4) = Array(-15.012943699204, -27.891166101196, 15.150000000000)
    pts(5) = Array(-15.051532975128, -28.125812967782, 15.150000000000)
    pts(6) = Array(-15.085082723126, -28.362931869545, 15.150000000000)
    pts(7) = Array(-15.113778475758, -28.602381023015, 15.150000000000)
    pts(8) = Array(-15.137771879876, -28.844037295650, 15.150000000000)
    pts(9) = Array(-15.157188799106, -29.087791729663, 15.150000000000)
    pts(10) = Array(-15.434454447870, -29.619885480815, 15.150000000000)
    pts(11) = Array(-17.190623728326, -28.636383427925, 15.150000000000)
    pts(12) = Array(-16.881810128416, -28.121957378321, 15.150000000000)
    pts(13) = Array(-16.684120620166, -27.978040391202, 15.150000000000)
    pts(14) = Array(-16.490611295364, -27.831317595565, 15.150000000000)
    pts(15) = Array(-16.301440819358, -27.681738605315, 15.150000000000)
    pts(16) = Array(-16.116793415163, -27.529238456829, 15.150000000000)
    pts(17) = Array(-15.936886913556, -27.373733039258, 15.150000000000)
    pts(18) = Array(-15.761984730172, -27.215112297505, 15.150000000000)
    pts(19) = Array(-15.592414621585, -27.053229590357, 15.150000000000)
    FSData_397 = pts
End Function

Private Function FSData_398() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-15.428600100973, -26.887883868469, 18.250000000000)
    pts(1) = Array(-14.864268312410, -27.203924855371, 18.250000000000)
    pts(2) = Array(-14.919655863997, -27.429992597518, 18.250000000000)
    pts(3) = Array(-14.969083423670, -27.659158364911, 18.250000000000)
    pts(4) = Array(-15.012943699204, -27.891166101196, 18.250000000000)
    pts(5) = Array(-15.051532975128, -28.125812967782, 18.250000000000)
    pts(6) = Array(-15.085082723126, -28.362931869545, 18.250000000000)
    pts(7) = Array(-15.113778475758, -28.602381023015, 18.250000000000)
    pts(8) = Array(-15.137771879876, -28.844037295650, 18.250000000000)
    pts(9) = Array(-15.157188799106, -29.087791729663, 18.250000000000)
    pts(10) = Array(-15.434454447870, -29.619885480815, 18.250000000000)
    pts(11) = Array(-17.190623728326, -28.636383427925, 18.250000000000)
    pts(12) = Array(-16.881810128416, -28.121957378321, 18.250000000000)
    pts(13) = Array(-16.684120620166, -27.978040391202, 18.250000000000)
    pts(14) = Array(-16.490611295364, -27.831317595565, 18.250000000000)
    pts(15) = Array(-16.301440819358, -27.681738605315, 18.250000000000)
    pts(16) = Array(-16.116793415163, -27.529238456829, 18.250000000000)
    pts(17) = Array(-15.936886913556, -27.373733039258, 18.250000000000)
    pts(18) = Array(-15.761984730172, -27.215112297505, 18.250000000000)
    pts(19) = Array(-15.592414621585, -27.053229590357, 18.250000000000)
    FSData_398 = pts
End Function

Private Function FSData_399() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-13.271439821541, -28.015511508149, 15.150000000000)
    pts(1) = Array(-12.684051389501, -28.286301284341, 15.150000000000)
    pts(2) = Array(-12.721531129129, -28.516017792297, 15.150000000000)
    pts(3) = Array(-12.752826181250, -28.748355159745, 15.150000000000)
    pts(4) = Array(-12.778348133001, -28.983088931165, 15.150000000000)
    pts(5) = Array(-12.798408270098, -29.220040139464, 15.150000000000)
    pts(6) = Array(-12.813250460808, -29.459060365677, 15.150000000000)
    pts(7) = Array(-12.823070789957, -29.700022820122, 15.150000000000)
    pts(8) = Array(-12.828030097789, -29.942816647908, 15.150000000000)
    pts(9) = Array(-12.828262408665, -30.187343102274, 15.150000000000)
    pts(10) = Array(-13.062925745409, -30.739550598048, 15.150000000000)
    pts(11) = Array(-14.890846023781, -29.896867807448, 15.150000000000)
    pts(12) = Array(-14.623345795809, -29.359798325877, 15.150000000000)
    pts(13) = Array(-14.437557295007, -29.200814446748, 15.150000000000)
    pts(14) = Array(-14.256156232741, -29.039361381885, 15.150000000000)
    pts(15) = Array(-14.079304738495, -28.875401349256, 15.150000000000)
    pts(16) = Array(-13.907191564372, -28.708884039472, 15.150000000000)
    pts(17) = Array(-13.740040468913, -28.539742691774, 15.150000000000)
    pts(18) = Array(-13.578122690552, -28.367888257682, 15.150000000000)
    pts(19) = Array(-13.421776480835, -28.193200263512, 15.150000000000)
    FSData_399 = pts
End Function

Private Function FSData_400() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-13.271439821541, -28.015511508149, 18.250000000000)
    pts(1) = Array(-12.684051389501, -28.286301284341, 18.250000000000)
    pts(2) = Array(-12.721531129129, -28.516017792297, 18.250000000000)
    pts(3) = Array(-12.752826181250, -28.748355159745, 18.250000000000)
    pts(4) = Array(-12.778348133001, -28.983088931165, 18.250000000000)
    pts(5) = Array(-12.798408270098, -29.220040139464, 18.250000000000)
    pts(6) = Array(-12.813250460808, -29.459060365677, 18.250000000000)
    pts(7) = Array(-12.823070789957, -29.700022820122, 18.250000000000)
    pts(8) = Array(-12.828030097789, -29.942816647908, 18.250000000000)
    pts(9) = Array(-12.828262408665, -30.187343102274, 18.250000000000)
    pts(10) = Array(-13.062925745409, -30.739550598048, 18.250000000000)
    pts(11) = Array(-14.890846023781, -29.896867807448, 18.250000000000)
    pts(12) = Array(-14.623345795809, -29.359798325877, 18.250000000000)
    pts(13) = Array(-14.437557295007, -29.200814446748, 18.250000000000)
    pts(14) = Array(-14.256156232741, -29.039361381885, 18.250000000000)
    pts(15) = Array(-14.079304738495, -28.875401349256, 18.250000000000)
    pts(16) = Array(-13.907191564372, -28.708884039472, 18.250000000000)
    pts(17) = Array(-13.740040468913, -28.539742691774, 18.250000000000)
    pts(18) = Array(-13.578122690552, -28.367888257682, 18.250000000000)
    pts(19) = Array(-13.421776480835, -28.193200263512, 18.250000000000)
    FSData_400 = pts
End Function

Private Function FSData_401() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-11.032456702408, -28.970414203278, 15.150000000000)
    pts(1) = Array(-10.425633071901, -29.194283259743, 15.150000000000)
    pts(2) = Array(-10.444973924511, -29.426232254848, 15.150000000000)
    pts(3) = Array(-10.457943524677, -29.660308785221, 15.150000000000)
    pts(4) = Array(-10.464969801327, -29.896321380018, 15.150000000000)
    pts(5) = Array(-10.466377122181, -30.134116047037, 15.150000000000)
    pts(6) = Array(-10.462420248560, -30.373563958524, 15.150000000000)
    pts(7) = Array(-10.453304608632, -30.614554100283, 15.150000000000)
    pts(8) = Array(-10.439199244397, -30.856988578534, 15.150000000000)
    pts(9) = Array(-10.420245514645, -31.100779466350, 15.150000000000)
    pts(10) = Array(-10.610859761864, -31.669696163905, 15.150000000000)
    pts(11) = Array(-12.499261301791, -30.973028055222, 15.150000000000)
    pts(12) = Array(-12.274723673615, -30.416626353631, 15.150000000000)
    pts(13) = Array(-12.101981628149, -30.243555770973, 15.150000000000)
    pts(14) = Array(-11.933807226290, -30.068367848717, 15.150000000000)
    pts(15) = Array(-11.770365062076, -29.891037641833, 15.150000000000)
    pts(16) = Array(-11.611847252978, -29.711529805339, 15.150000000000)
    pts(17) = Array(-11.458482105752, -29.529795340167, 15.150000000000)
    pts(18) = Array(-11.310547009360, -29.345766753470, 15.150000000000)
    pts(19) = Array(-11.168388624887, -29.159350481852, 15.150000000000)
    FSData_401 = pts
End Function

Private Function FSData_402() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-11.032456702408, -28.970414203278, 18.250000000000)
    pts(1) = Array(-10.425633071901, -29.194283259743, 18.250000000000)
    pts(2) = Array(-10.444973924511, -29.426232254848, 18.250000000000)
    pts(3) = Array(-10.457943524677, -29.660308785221, 18.250000000000)
    pts(4) = Array(-10.464969801327, -29.896321380018, 18.250000000000)
    pts(5) = Array(-10.466377122181, -30.134116047037, 18.250000000000)
    pts(6) = Array(-10.462420248560, -30.373563958524, 18.250000000000)
    pts(7) = Array(-10.453304608632, -30.614554100283, 18.250000000000)
    pts(8) = Array(-10.439199244397, -30.856988578534, 18.250000000000)
    pts(9) = Array(-10.420245514645, -31.100779466350, 18.250000000000)
    pts(10) = Array(-10.610859761864, -31.669696163905, 18.250000000000)
    pts(11) = Array(-12.499261301791, -30.973028055222, 18.250000000000)
    pts(12) = Array(-12.274723673615, -30.416626353631, 18.250000000000)
    pts(13) = Array(-12.101981628149, -30.243555770973, 18.250000000000)
    pts(14) = Array(-11.933807226290, -30.068367848717, 18.250000000000)
    pts(15) = Array(-11.770365062076, -29.891037641833, 18.250000000000)
    pts(16) = Array(-11.611847252978, -29.711529805339, 18.250000000000)
    pts(17) = Array(-11.458482105752, -29.529795340167, 18.250000000000)
    pts(18) = Array(-11.310547009360, -29.345766753470, 18.250000000000)
    pts(19) = Array(-11.168388624887, -29.159350481852, 18.250000000000)
    FSData_402 = pts
End Function

Private Function FSData_403() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-8.725454819040, -29.746704661204, 15.150000000000)
    pts(1) = Array(-8.102937259537, -29.922272770764, 15.150000000000)
    pts(2) = Array(-8.104019982343, -30.155024210333, 15.150000000000)
    pts(3) = Array(-8.098584168655, -30.389396743983, 15.150000000000)
    pts(4) = Array(-8.087071450871, -30.625233066028, 15.150000000000)
    pts(5) = Array(-8.069817278881, -30.862405108570, 15.150000000000)
    pts(6) = Array(-8.047085736372, -31.100804429331, 15.150000000000)
    pts(7) = Array(-8.019090328318, -31.340336474044, 15.150000000000)
    pts(8) = Array(-7.986007256277, -31.580916913584, 15.150000000000)
    pts(9) = Array(-7.947984341944, -31.822469182957, 15.150000000000)
    pts(10) = Array(-8.093374299418, -32.404587521669, 15.150000000000)
    pts(11) = Array(-10.030614477449, -31.858229285457, 15.150000000000)
    pts(12) = Array(-9.850423798214, -31.285925765359, 15.150000000000)
    pts(13) = Array(-9.691793220237, -31.099835516867, 15.150000000000)
    pts(14) = Array(-9.537882329896, -30.911992829015, 15.150000000000)
    pts(15) = Array(-9.388857171005, -30.722385747570, 15.150000000000)
    pts(16) = Array(-9.244912041937, -30.530994109869, 15.150000000000)
    pts(17) = Array(-9.106278390077, -30.337786980636, 15.150000000000)
    pts(18) = Array(-8.973238044717, -30.142718838765, 15.150000000000)
    pts(19) = Array(-8.846143939200, -29.945723608004, 15.150000000000)
    FSData_403 = pts
End Function

Private Function FSData_404() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-8.725454819040, -29.746704661204, 18.250000000000)
    pts(1) = Array(-8.102937259537, -29.922272770764, 18.250000000000)
    pts(2) = Array(-8.104019982343, -30.155024210333, 18.250000000000)
    pts(3) = Array(-8.098584168655, -30.389396743983, 18.250000000000)
    pts(4) = Array(-8.087071450871, -30.625233066028, 18.250000000000)
    pts(5) = Array(-8.069817278881, -30.862405108570, 18.250000000000)
    pts(6) = Array(-8.047085736372, -31.100804429331, 18.250000000000)
    pts(7) = Array(-8.019090328318, -31.340336474044, 18.250000000000)
    pts(8) = Array(-7.986007256277, -31.580916913584, 18.250000000000)
    pts(9) = Array(-7.947984341944, -31.822469182957, 18.250000000000)
    pts(10) = Array(-8.093374299418, -32.404587521669, 18.250000000000)
    pts(11) = Array(-10.030614477449, -31.858229285457, 18.250000000000)
    pts(12) = Array(-9.850423798214, -31.285925765359, 18.250000000000)
    pts(13) = Array(-9.691793220237, -31.099835516867, 18.250000000000)
    pts(14) = Array(-9.537882329896, -30.911992829015, 18.250000000000)
    pts(15) = Array(-9.388857171005, -30.722385747570, 18.250000000000)
    pts(16) = Array(-9.244912041937, -30.530994109869, 18.250000000000)
    pts(17) = Array(-9.106278390077, -30.337786980636, 18.250000000000)
    pts(18) = Array(-8.973238044717, -30.142718838765, 18.250000000000)
    pts(19) = Array(-8.846143939200, -29.945723608004, 18.250000000000)
    FSData_404 = pts
End Function

Private Function FSData_405() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-6.364657605205, -30.339596793110, 15.150000000000)
    pts(1) = Array(-5.730284144469, -30.465781519988, 15.150000000000)
    pts(2) = Array(-5.713102062123, -30.697900413998, 15.150000000000)
    pts(3) = Array(-5.689294348182, -30.931123966318, 15.150000000000)
    pts(4) = Array(-5.659313615697, -31.165330006262, 15.150000000000)
    pts(5) = Array(-5.623504328571, -31.400417179817, 15.150000000000)
    pts(6) = Array(-5.582138264692, -31.636298098764, 15.150000000000)
    pts(7) = Array(-5.535435689512, -31.872895251723, 15.150000000000)
    pts(8) = Array(-5.483578877805, -32.110138394141, 15.150000000000)
    pts(9) = Array(-5.426721202801, -32.347962795006, 15.150000000000)
    pts(10) = Array(-5.525990493096, -32.939693821744, 15.150000000000)
    pts(11) = Array(-7.500125579336, -32.547013938212, 15.150000000000)
    pts(12) = Array(-7.365392784498, -31.962337041119, 15.150000000000)
    pts(13) = Array(-7.221851684273, -31.764374435055, 15.150000000000)
    pts(14) = Array(-7.083153217270, -31.565035094241, 15.150000000000)
    pts(15) = Array(-6.949463853363, -31.364320128943, 15.150000000000)
    pts(16) = Array(-6.820978873913, -31.162224683126, 15.150000000000)
    pts(17) = Array(-6.697931439982, -30.958736076030, 15.150000000000)
    pts(18) = Array(-6.580606083623, -30.753831038948, 15.150000000000)
    pts(19) = Array(-6.469359834487, -30.547471390148, 15.150000000000)
    FSData_405 = pts
End Function

Private Function FSData_406() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-6.364657605205, -30.339596793110, 18.250000000000)
    pts(1) = Array(-5.730284144469, -30.465781519988, 18.250000000000)
    pts(2) = Array(-5.713102062123, -30.697900413998, 18.250000000000)
    pts(3) = Array(-5.689294348182, -30.931123966318, 18.250000000000)
    pts(4) = Array(-5.659313615697, -31.165330006262, 18.250000000000)
    pts(5) = Array(-5.623504328571, -31.400417179817, 18.250000000000)
    pts(6) = Array(-5.582138264692, -31.636298098764, 18.250000000000)
    pts(7) = Array(-5.535435689512, -31.872895251723, 18.250000000000)
    pts(8) = Array(-5.483578877805, -32.110138394141, 18.250000000000)
    pts(9) = Array(-5.426721202801, -32.347962795006, 18.250000000000)
    pts(10) = Array(-5.525990493096, -32.939693821744, 18.250000000000)
    pts(11) = Array(-7.500125579336, -32.547013938212, 18.250000000000)
    pts(12) = Array(-7.365392784498, -31.962337041119, 18.250000000000)
    pts(13) = Array(-7.221851684273, -31.764374435055, 18.250000000000)
    pts(14) = Array(-7.083153217270, -31.565035094241, 18.250000000000)
    pts(15) = Array(-6.949463853363, -31.364320128943, 18.250000000000)
    pts(16) = Array(-6.820978873913, -31.162224683126, 18.250000000000)
    pts(17) = Array(-6.697931439982, -30.958736076030, 18.250000000000)
    pts(18) = Array(-6.580606083623, -30.753831038948, 18.250000000000)
    pts(19) = Array(-6.469359834487, -30.547471390148, 18.250000000000)
    FSData_406 = pts
End Function

Private Function FSData_407() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-3.964620160771, -30.745435221847, 15.150000000000)
    pts(1) = Array(-3.322301922137, -30.821458595241, 15.150000000000)
    pts(2) = Array(-3.286960967891, -31.051513853523, 15.150000000000)
    pts(3) = Array(-3.244928136169, -31.282150523759, 15.150000000000)
    pts(4) = Array(-3.196664230169, -31.513282323483, 15.150000000000)
    pts(5) = Array(-3.142520604070, -31.744835237452, 15.150000000000)
    pts(6) = Array(-3.082775054361, -31.976743470282, 15.150000000000)
    pts(7) = Array(-3.017653248961, -32.208947031361, 15.150000000000)
    pts(8) = Array(-2.947342412076, -32.441390193794, 15.150000000000)
    pts(9) = Array(-2.872000522875, -32.674020459634, 15.150000000000)
    pts(10) = Array(-2.924537117806, -33.271715955847, 15.150000000000)
    pts(11) = Array(-4.923395912982, -33.035135427058, 15.150000000000)
    pts(12) = Array(-4.834951675024, -32.441689880464, 15.150000000000)
    pts(13) = Array(-4.707385031166, -32.233075422124, 15.150000000000)
    pts(14) = Array(-4.584754109672, -32.023468421672, 15.150000000000)
    pts(15) = Array(-4.467224780134, -31.812883047026, 15.150000000000)
    pts(16) = Array(-4.354992102925, -31.601329778721, 15.150000000000)
    pts(17) = Array(-4.248289515272, -31.388814268692, 15.150000000000)
    pts(18) = Array(-4.147402497751, -31.175335644090, 15.150000000000)
    pts(19) = Array(-4.052689975114, -30.960883853107, 15.150000000000)
    FSData_407 = pts
End Function

Private Function FSData_408() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-3.964620160771, -30.745435221847, 18.250000000000)
    pts(1) = Array(-3.322301922137, -30.821458595241, 18.250000000000)
    pts(2) = Array(-3.286960967891, -31.051513853523, 18.250000000000)
    pts(3) = Array(-3.244928136169, -31.282150523759, 18.250000000000)
    pts(4) = Array(-3.196664230169, -31.513282323483, 18.250000000000)
    pts(5) = Array(-3.142520604070, -31.744835237452, 18.250000000000)
    pts(6) = Array(-3.082775054361, -31.976743470282, 18.250000000000)
    pts(7) = Array(-3.017653248961, -32.208947031361, 18.250000000000)
    pts(8) = Array(-2.947342412076, -32.441390193794, 18.250000000000)
    pts(9) = Array(-2.872000522875, -32.674020459634, 18.250000000000)
    pts(10) = Array(-2.924537117806, -33.271715955847, 18.250000000000)
    pts(11) = Array(-4.923395912982, -33.035135427058, 18.250000000000)
    pts(12) = Array(-4.834951675024, -32.441689880464, 18.250000000000)
    pts(13) = Array(-4.707385031166, -32.233075422124, 18.250000000000)
    pts(14) = Array(-4.584754109672, -32.023468421672, 18.250000000000)
    pts(15) = Array(-4.467224780134, -31.812883047026, 18.250000000000)
    pts(16) = Array(-4.354992102925, -31.601329778721, 18.250000000000)
    pts(17) = Array(-4.248289515272, -31.388814268692, 18.250000000000)
    pts(18) = Array(-4.147402497751, -31.175335644090, 18.250000000000)
    pts(19) = Array(-4.052689975114, -30.960883853107, 18.250000000000)
    FSData_408 = pts
End Function

Private Function FSData_409() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-1.540139514675, -30.961717818547, 15.150000000000)
    pts(1) = Array(-0.893836603677, -30.987111129080, 15.150000000000)
    pts(2) = Array(-0.840554666267, -31.213684384465, 15.150000000000)
    pts(3) = Array(-0.780555863149, -31.440312220850, 15.150000000000)
    pts(4) = Array(-0.714306346663, -31.666944775951, 15.150000000000)
    pts(5) = Array(-0.642162195051, -31.893535829620, 15.150000000000)
    pts(6) = Array(-0.564405510693, -32.120041584959, 15.150000000000)
    pts(7) = Array(-0.481265972658, -32.346419941990, 15.150000000000)
    pts(8) = Array(-0.392934600286, -32.572630035045, 15.150000000000)
    pts(9) = Array(-0.299573004689, -32.798631922915, 15.150000000000)
    pts(10) = Array(-0.305052998677, -33.398606897115, 15.150000000000)
    pts(11) = Array(-2.316311873629, -33.319584320698, 15.150000000000)
    pts(12) = Array(-2.274701480689, -32.721028913739, 15.150000000000)
    pts(13) = Array(-2.163895783978, -32.503048780631, 15.150000000000)
    pts(14) = Array(-2.058088468403, -32.284466417400, 15.150000000000)
    pts(15) = Array(-1.957443780632, -32.065308962267, 15.150000000000)
    pts(16) = Array(-1.862155357441, -31.845602167721, 15.150000000000)
    pts(17) = Array(-1.772455473001, -31.625369983547, 15.150000000000)
    pts(18) = Array(-1.688628796329, -31.404633938134, 15.150000000000)
    pts(19) = Array(-1.611033934388, -31.183412171574, 15.150000000000)
    FSData_409 = pts
End Function

Private Function FSData_410() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-1.540139514675, -30.961717818547, 18.250000000000)
    pts(1) = Array(-0.893836603677, -30.987111129080, 18.250000000000)
    pts(2) = Array(-0.840554666267, -31.213684384465, 18.250000000000)
    pts(3) = Array(-0.780555863149, -31.440312220850, 18.250000000000)
    pts(4) = Array(-0.714306346663, -31.666944775951, 18.250000000000)
    pts(5) = Array(-0.642162195051, -31.893535829620, 18.250000000000)
    pts(6) = Array(-0.564405510693, -32.120041584959, 18.250000000000)
    pts(7) = Array(-0.481265972658, -32.346419941990, 18.250000000000)
    pts(8) = Array(-0.392934600286, -32.572630035045, 18.250000000000)
    pts(9) = Array(-0.299573004689, -32.798631922915, 18.250000000000)
    pts(10) = Array(-0.305052998677, -33.398606897115, 18.250000000000)
    pts(11) = Array(-2.316311873629, -33.319584320698, 18.250000000000)
    pts(12) = Array(-2.274701480689, -32.721028913739, 18.250000000000)
    pts(13) = Array(-2.163895783978, -32.503048780631, 18.250000000000)
    pts(14) = Array(-2.058088468403, -32.284466417400, 18.250000000000)
    pts(15) = Array(-1.957443780632, -32.065308962267, 18.250000000000)
    pts(16) = Array(-1.862155357441, -31.845602167721, 18.250000000000)
    pts(17) = Array(-1.772455473001, -31.625369983547, 18.250000000000)
    pts(18) = Array(-1.688628796329, -31.404633938134, 18.250000000000)
    pts(19) = Array(-1.611033934388, -31.183412171574, 18.250000000000)
    FSData_410 = pts
End Function

Private Function FSData_411() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(0.893836603677, -30.987111129080, 15.150000000000)
    pts(1) = Array(1.540139514675, -30.961717818547, 15.150000000000)
    pts(2) = Array(1.611033934388, -31.183412171574, 15.150000000000)
    pts(3) = Array(1.688628796329, -31.404633938134, 15.150000000000)
    pts(4) = Array(1.772455473001, -31.625369983547, 15.150000000000)
    pts(5) = Array(1.862155357441, -31.845602167721, 15.150000000000)
    pts(6) = Array(1.957443780632, -32.065308962267, 15.150000000000)
    pts(7) = Array(2.058088468403, -32.284466417400, 15.150000000000)
    pts(8) = Array(2.163895783978, -32.503048780631, 15.150000000000)
    pts(9) = Array(2.274701480689, -32.721028913739, 15.150000000000)
    pts(10) = Array(2.316311873629, -33.319584320698, 15.150000000000)
    pts(11) = Array(0.305052998677, -33.398606897115, 15.150000000000)
    pts(12) = Array(0.299573004689, -32.798631922915, 15.150000000000)
    pts(13) = Array(0.392934600286, -32.572630035045, 15.150000000000)
    pts(14) = Array(0.481265972658, -32.346419941990, 15.150000000000)
    pts(15) = Array(0.564405510693, -32.120041584959, 15.150000000000)
    pts(16) = Array(0.642162195051, -31.893535829620, 15.150000000000)
    pts(17) = Array(0.714306346663, -31.666944775951, 15.150000000000)
    pts(18) = Array(0.780555863149, -31.440312220850, 15.150000000000)
    pts(19) = Array(0.840554666267, -31.213684384465, 15.150000000000)
    FSData_411 = pts
End Function

Private Function FSData_412() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(0.893836603677, -30.987111129080, 18.250000000000)
    pts(1) = Array(1.540139514675, -30.961717818547, 18.250000000000)
    pts(2) = Array(1.611033934388, -31.183412171574, 18.250000000000)
    pts(3) = Array(1.688628796329, -31.404633938134, 18.250000000000)
    pts(4) = Array(1.772455473001, -31.625369983547, 18.250000000000)
    pts(5) = Array(1.862155357441, -31.845602167721, 18.250000000000)
    pts(6) = Array(1.957443780632, -32.065308962267, 18.250000000000)
    pts(7) = Array(2.058088468403, -32.284466417400, 18.250000000000)
    pts(8) = Array(2.163895783978, -32.503048780631, 18.250000000000)
    pts(9) = Array(2.274701480689, -32.721028913739, 18.250000000000)
    pts(10) = Array(2.316311873629, -33.319584320698, 18.250000000000)
    pts(11) = Array(0.305052998677, -33.398606897115, 18.250000000000)
    pts(12) = Array(0.299573004689, -32.798631922915, 18.250000000000)
    pts(13) = Array(0.392934600286, -32.572630035045, 18.250000000000)
    pts(14) = Array(0.481265972658, -32.346419941990, 18.250000000000)
    pts(15) = Array(0.564405510693, -32.120041584959, 18.250000000000)
    pts(16) = Array(0.642162195051, -31.893535829620, 18.250000000000)
    pts(17) = Array(0.714306346663, -31.666944775951, 18.250000000000)
    pts(18) = Array(0.780555863149, -31.440312220850, 18.250000000000)
    pts(19) = Array(0.840554666267, -31.213684384465, 18.250000000000)
    FSData_412 = pts
End Function

Private Function FSData_413() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(3.322301922137, -30.821458595241, 15.150000000000)
    pts(1) = Array(3.964620160771, -30.745435221847, 15.150000000000)
    pts(2) = Array(4.052689975114, -30.960883853107, 15.150000000000)
    pts(3) = Array(4.147402497751, -31.175335644090, 15.150000000000)
    pts(4) = Array(4.248289515272, -31.388814268692, 15.150000000000)
    pts(5) = Array(4.354992102925, -31.601329778721, 15.150000000000)
    pts(6) = Array(4.467224780134, -31.812883047026, 15.150000000000)
    pts(7) = Array(4.584754109672, -32.023468421672, 15.150000000000)
    pts(8) = Array(4.707385031166, -32.233075422124, 15.150000000000)
    pts(9) = Array(4.834951675024, -32.441689880464, 15.150000000000)
    pts(10) = Array(4.923395912982, -33.035135427058, 15.150000000000)
    pts(11) = Array(2.924537117806, -33.271715955847, 15.150000000000)
    pts(12) = Array(2.872000522875, -32.674020459634, 15.150000000000)
    pts(13) = Array(2.947342412076, -32.441390193794, 15.150000000000)
    pts(14) = Array(3.017653248961, -32.208947031361, 15.150000000000)
    pts(15) = Array(3.082775054361, -31.976743470282, 15.150000000000)
    pts(16) = Array(3.142520604070, -31.744835237452, 15.150000000000)
    pts(17) = Array(3.196664230169, -31.513282323483, 15.150000000000)
    pts(18) = Array(3.244928136169, -31.282150523759, 15.150000000000)
    pts(19) = Array(3.286960967891, -31.051513853523, 15.150000000000)
    FSData_413 = pts
End Function

Private Function FSData_414() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(3.322301922137, -30.821458595241, 18.250000000000)
    pts(1) = Array(3.964620160771, -30.745435221847, 18.250000000000)
    pts(2) = Array(4.052689975114, -30.960883853107, 18.250000000000)
    pts(3) = Array(4.147402497751, -31.175335644090, 18.250000000000)
    pts(4) = Array(4.248289515272, -31.388814268692, 18.250000000000)
    pts(5) = Array(4.354992102925, -31.601329778721, 18.250000000000)
    pts(6) = Array(4.467224780134, -31.812883047026, 18.250000000000)
    pts(7) = Array(4.584754109672, -32.023468421672, 18.250000000000)
    pts(8) = Array(4.707385031166, -32.233075422124, 18.250000000000)
    pts(9) = Array(4.834951675024, -32.441689880464, 18.250000000000)
    pts(10) = Array(4.923395912982, -33.035135427058, 18.250000000000)
    pts(11) = Array(2.924537117806, -33.271715955847, 18.250000000000)
    pts(12) = Array(2.872000522875, -32.674020459634, 18.250000000000)
    pts(13) = Array(2.947342412076, -32.441390193794, 18.250000000000)
    pts(14) = Array(3.017653248961, -32.208947031361, 18.250000000000)
    pts(15) = Array(3.082775054361, -31.976743470282, 18.250000000000)
    pts(16) = Array(3.142520604070, -31.744835237452, 18.250000000000)
    pts(17) = Array(3.196664230169, -31.513282323483, 18.250000000000)
    pts(18) = Array(3.244928136169, -31.282150523759, 18.250000000000)
    pts(19) = Array(3.286960967891, -31.051513853523, 18.250000000000)
    FSData_414 = pts
End Function

Private Function FSData_415() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(5.730284144469, -30.465781519988, 15.150000000000)
    pts(1) = Array(6.364657605205, -30.339596793110, 15.150000000000)
    pts(2) = Array(6.469359834487, -30.547471390148, 15.150000000000)
    pts(3) = Array(6.580606083624, -30.753831038948, 15.150000000000)
    pts(4) = Array(6.697931439982, -30.958736076030, 15.150000000000)
    pts(5) = Array(6.820978873913, -31.162224683126, 15.150000000000)
    pts(6) = Array(6.949463853363, -31.364320128943, 15.150000000000)
    pts(7) = Array(7.083153217270, -31.565035094241, 15.150000000000)
    pts(8) = Array(7.221851684273, -31.764374435055, 15.150000000000)
    pts(9) = Array(7.365392784498, -31.962337041119, 15.150000000000)
    pts(10) = Array(7.500125579336, -32.547013938212, 15.150000000000)
    pts(11) = Array(5.525990493096, -32.939693821744, 15.150000000000)
    pts(12) = Array(5.426721202801, -32.347962795006, 15.150000000000)
    pts(13) = Array(5.483578877805, -32.110138394141, 15.150000000000)
    pts(14) = Array(5.535435689512, -31.872895251723, 15.150000000000)
    pts(15) = Array(5.582138264692, -31.636298098764, 15.150000000000)
    pts(16) = Array(5.623504328571, -31.400417179817, 15.150000000000)
    pts(17) = Array(5.659313615697, -31.165330006262, 15.150000000000)
    pts(18) = Array(5.689294348182, -30.931123966318, 15.150000000000)
    pts(19) = Array(5.713102062123, -30.697900413998, 15.150000000000)
    FSData_415 = pts
End Function

Private Function FSData_416() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(5.730284144469, -30.465781519988, 18.250000000000)
    pts(1) = Array(6.364657605205, -30.339596793110, 18.250000000000)
    pts(2) = Array(6.469359834487, -30.547471390148, 18.250000000000)
    pts(3) = Array(6.580606083624, -30.753831038948, 18.250000000000)
    pts(4) = Array(6.697931439982, -30.958736076030, 18.250000000000)
    pts(5) = Array(6.820978873913, -31.162224683126, 18.250000000000)
    pts(6) = Array(6.949463853363, -31.364320128943, 18.250000000000)
    pts(7) = Array(7.083153217270, -31.565035094241, 18.250000000000)
    pts(8) = Array(7.221851684273, -31.764374435055, 18.250000000000)
    pts(9) = Array(7.365392784498, -31.962337041119, 18.250000000000)
    pts(10) = Array(7.500125579336, -32.547013938212, 18.250000000000)
    pts(11) = Array(5.525990493096, -32.939693821744, 18.250000000000)
    pts(12) = Array(5.426721202801, -32.347962795006, 18.250000000000)
    pts(13) = Array(5.483578877805, -32.110138394141, 18.250000000000)
    pts(14) = Array(5.535435689512, -31.872895251723, 18.250000000000)
    pts(15) = Array(5.582138264692, -31.636298098764, 18.250000000000)
    pts(16) = Array(5.623504328571, -31.400417179817, 18.250000000000)
    pts(17) = Array(5.659313615697, -31.165330006262, 18.250000000000)
    pts(18) = Array(5.689294348182, -30.931123966318, 18.250000000000)
    pts(19) = Array(5.713102062123, -30.697900413998, 18.250000000000)
    FSData_416 = pts
End Function

Private Function FSData_417() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(8.102937259537, -29.922272770764, 15.150000000000)
    pts(1) = Array(8.725454819040, -29.746704661204, 15.150000000000)
    pts(2) = Array(8.846143939200, -29.945723608004, 15.150000000000)
    pts(3) = Array(8.973238044717, -30.142718838765, 15.150000000000)
    pts(4) = Array(9.106278390077, -30.337786980636, 15.150000000000)
    pts(5) = Array(9.244912041937, -30.530994109869, 15.150000000000)
    pts(6) = Array(9.388857171005, -30.722385747570, 15.150000000000)
    pts(7) = Array(9.537882329896, -30.911992829015, 15.150000000000)
    pts(8) = Array(9.691793220237, -31.099835516867, 15.150000000000)
    pts(9) = Array(9.850423798214, -31.285925765359, 15.150000000000)
    pts(10) = Array(10.030614477449, -31.858229285457, 15.150000000000)
    pts(11) = Array(8.093374299418, -32.404587521669, 15.150000000000)
    pts(12) = Array(7.947984341943, -31.822469182957, 15.150000000000)
    pts(13) = Array(7.986007256277, -31.580916913584, 15.150000000000)
    pts(14) = Array(8.019090328318, -31.340336474044, 15.150000000000)
    pts(15) = Array(8.047085736372, -31.100804429331, 15.150000000000)
    pts(16) = Array(8.069817278881, -30.862405108570, 15.150000000000)
    pts(17) = Array(8.087071450871, -30.625233066028, 15.150000000000)
    pts(18) = Array(8.098584168655, -30.389396743983, 15.150000000000)
    pts(19) = Array(8.104019982343, -30.155024210333, 15.150000000000)
    FSData_417 = pts
End Function

Private Function FSData_418() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(8.102937259537, -29.922272770764, 18.250000000000)
    pts(1) = Array(8.725454819040, -29.746704661204, 18.250000000000)
    pts(2) = Array(8.846143939200, -29.945723608004, 18.250000000000)
    pts(3) = Array(8.973238044717, -30.142718838765, 18.250000000000)
    pts(4) = Array(9.106278390077, -30.337786980636, 18.250000000000)
    pts(5) = Array(9.244912041937, -30.530994109869, 18.250000000000)
    pts(6) = Array(9.388857171005, -30.722385747570, 18.250000000000)
    pts(7) = Array(9.537882329896, -30.911992829015, 18.250000000000)
    pts(8) = Array(9.691793220237, -31.099835516867, 18.250000000000)
    pts(9) = Array(9.850423798214, -31.285925765359, 18.250000000000)
    pts(10) = Array(10.030614477449, -31.858229285457, 18.250000000000)
    pts(11) = Array(8.093374299418, -32.404587521669, 18.250000000000)
    pts(12) = Array(7.947984341943, -31.822469182957, 18.250000000000)
    pts(13) = Array(7.986007256277, -31.580916913584, 18.250000000000)
    pts(14) = Array(8.019090328318, -31.340336474044, 18.250000000000)
    pts(15) = Array(8.047085736372, -31.100804429331, 18.250000000000)
    pts(16) = Array(8.069817278881, -30.862405108570, 18.250000000000)
    pts(17) = Array(8.087071450871, -30.625233066028, 18.250000000000)
    pts(18) = Array(8.098584168655, -30.389396743983, 18.250000000000)
    pts(19) = Array(8.104019982343, -30.155024210333, 18.250000000000)
    FSData_418 = pts
End Function

Private Function FSData_419() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(10.425633071901, -29.194283259743, 15.150000000000)
    pts(1) = Array(11.032456702408, -28.970414203278, 15.150000000000)
    pts(2) = Array(11.168388624887, -29.159350481852, 15.150000000000)
    pts(3) = Array(11.310547009360, -29.345766753470, 15.150000000000)
    pts(4) = Array(11.458482105752, -29.529795340167, 15.150000000000)
    pts(5) = Array(11.611847252977, -29.711529805339, 15.150000000000)
    pts(6) = Array(11.770365062076, -29.891037641833, 15.150000000000)
    pts(7) = Array(11.933807226290, -30.068367848717, 15.150000000000)
    pts(8) = Array(12.101981628149, -30.243555770973, 15.150000000000)
    pts(9) = Array(12.274723673615, -30.416626353631, 15.150000000000)
    pts(10) = Array(12.499261301791, -30.973028055222, 15.150000000000)
    pts(11) = Array(10.610859761864, -31.669696163905, 15.150000000000)
    pts(12) = Array(10.420245514645, -31.100779466350, 15.150000000000)
    pts(13) = Array(10.439199244397, -30.856988578534, 15.150000000000)
    pts(14) = Array(10.453304608631, -30.614554100283, 15.150000000000)
    pts(15) = Array(10.462420248560, -30.373563958524, 15.150000000000)
    pts(16) = Array(10.466377122181, -30.134116047037, 15.150000000000)
    pts(17) = Array(10.464969801327, -29.896321380018, 15.150000000000)
    pts(18) = Array(10.457943524677, -29.660308785221, 15.150000000000)
    pts(19) = Array(10.444973924511, -29.426232254848, 15.150000000000)
    FSData_419 = pts
End Function

Private Function FSData_420() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(10.425633071901, -29.194283259743, 18.250000000000)
    pts(1) = Array(11.032456702408, -28.970414203278, 18.250000000000)
    pts(2) = Array(11.168388624887, -29.159350481852, 18.250000000000)
    pts(3) = Array(11.310547009360, -29.345766753470, 18.250000000000)
    pts(4) = Array(11.458482105752, -29.529795340167, 18.250000000000)
    pts(5) = Array(11.611847252977, -29.711529805339, 18.250000000000)
    pts(6) = Array(11.770365062076, -29.891037641833, 18.250000000000)
    pts(7) = Array(11.933807226290, -30.068367848717, 18.250000000000)
    pts(8) = Array(12.101981628149, -30.243555770973, 18.250000000000)
    pts(9) = Array(12.274723673615, -30.416626353631, 18.250000000000)
    pts(10) = Array(12.499261301791, -30.973028055222, 18.250000000000)
    pts(11) = Array(10.610859761864, -31.669696163905, 18.250000000000)
    pts(12) = Array(10.420245514645, -31.100779466350, 18.250000000000)
    pts(13) = Array(10.439199244397, -30.856988578534, 18.250000000000)
    pts(14) = Array(10.453304608631, -30.614554100283, 18.250000000000)
    pts(15) = Array(10.462420248560, -30.373563958524, 18.250000000000)
    pts(16) = Array(10.466377122181, -30.134116047037, 18.250000000000)
    pts(17) = Array(10.464969801327, -29.896321380018, 18.250000000000)
    pts(18) = Array(10.457943524677, -29.660308785221, 18.250000000000)
    pts(19) = Array(10.444973924511, -29.426232254848, 18.250000000000)
    FSData_420 = pts
End Function

Private Function FSData_421() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(12.684051389501, -28.286301284341, 15.150000000000)
    pts(1) = Array(13.271439821541, -28.015511508149, 15.150000000000)
    pts(2) = Array(13.421776480835, -28.193200263512, 15.150000000000)
    pts(3) = Array(13.578122690552, -28.367888257682, 15.150000000000)
    pts(4) = Array(13.740040468913, -28.539742691774, 15.150000000000)
    pts(5) = Array(13.907191564372, -28.708884039472, 15.150000000000)
    pts(6) = Array(14.079304738495, -28.875401349256, 15.150000000000)
    pts(7) = Array(14.256156232741, -29.039361381885, 15.150000000000)
    pts(8) = Array(14.437557295007, -29.200814446748, 15.150000000000)
    pts(9) = Array(14.623345795809, -29.359798325877, 15.150000000000)
    pts(10) = Array(14.890846023781, -29.896867807448, 15.150000000000)
    pts(11) = Array(13.062925745409, -30.739550598048, 15.150000000000)
    pts(12) = Array(12.828262408665, -30.187343102274, 15.150000000000)
    pts(13) = Array(12.828030097789, -29.942816647908, 15.150000000000)
    pts(14) = Array(12.823070789957, -29.700022820122, 15.150000000000)
    pts(15) = Array(12.813250460808, -29.459060365677, 15.150000000000)
    pts(16) = Array(12.798408270098, -29.220040139464, 15.150000000000)
    pts(17) = Array(12.778348133001, -28.983088931165, 15.150000000000)
    pts(18) = Array(12.752826181250, -28.748355159745, 15.150000000000)
    pts(19) = Array(12.721531129129, -28.516017792297, 15.150000000000)
    FSData_421 = pts
End Function

Private Function FSData_422() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(12.684051389501, -28.286301284341, 18.250000000000)
    pts(1) = Array(13.271439821541, -28.015511508149, 18.250000000000)
    pts(2) = Array(13.421776480835, -28.193200263512, 18.250000000000)
    pts(3) = Array(13.578122690552, -28.367888257682, 18.250000000000)
    pts(4) = Array(13.740040468913, -28.539742691774, 18.250000000000)
    pts(5) = Array(13.907191564372, -28.708884039472, 18.250000000000)
    pts(6) = Array(14.079304738495, -28.875401349256, 18.250000000000)
    pts(7) = Array(14.256156232741, -29.039361381885, 18.250000000000)
    pts(8) = Array(14.437557295007, -29.200814446748, 18.250000000000)
    pts(9) = Array(14.623345795809, -29.359798325877, 18.250000000000)
    pts(10) = Array(14.890846023781, -29.896867807448, 18.250000000000)
    pts(11) = Array(13.062925745409, -30.739550598048, 18.250000000000)
    pts(12) = Array(12.828262408665, -30.187343102274, 18.250000000000)
    pts(13) = Array(12.828030097789, -29.942816647908, 18.250000000000)
    pts(14) = Array(12.823070789957, -29.700022820122, 18.250000000000)
    pts(15) = Array(12.813250460808, -29.459060365677, 18.250000000000)
    pts(16) = Array(12.798408270098, -29.220040139464, 18.250000000000)
    pts(17) = Array(12.778348133001, -28.983088931165, 18.250000000000)
    pts(18) = Array(12.752826181250, -28.748355159745, 18.250000000000)
    pts(19) = Array(12.721531129129, -28.516017792297, 18.250000000000)
    FSData_422 = pts
End Function

Private Function FSData_423() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(14.864268312410, -27.203924855371, 15.150000000000)
    pts(1) = Array(15.428600100973, -26.887883868469, 15.150000000000)
    pts(2) = Array(15.592414621585, -27.053229590357, 15.150000000000)
    pts(3) = Array(15.761984730172, -27.215112297505, 15.150000000000)
    pts(4) = Array(15.936886913556, -27.373733039258, 15.150000000000)
    pts(5) = Array(16.116793415163, -27.529238456829, 15.150000000000)
    pts(6) = Array(16.301440819358, -27.681738605315, 15.150000000000)
    pts(7) = Array(16.490611295364, -27.831317595565, 15.150000000000)
    pts(8) = Array(16.684120620166, -27.978040391202, 15.150000000000)
    pts(9) = Array(16.881810128416, -28.121957378321, 15.150000000000)
    pts(10) = Array(17.190623728326, -28.636383427925, 15.150000000000)
    pts(11) = Array(15.434454447870, -29.619885480815, 15.150000000000)
    pts(12) = Array(15.157188799106, -29.087791729663, 15.150000000000)
    pts(13) = Array(15.137771879876, -28.844037295650, 15.150000000000)
    pts(14) = Array(15.113778475758, -28.602381023015, 15.150000000000)
    pts(15) = Array(15.085082723126, -28.362931869545, 15.150000000000)
    pts(16) = Array(15.051532975128, -28.125812967782, 15.150000000000)
    pts(17) = Array(15.012943699204, -27.891166101196, 15.150000000000)
    pts(18) = Array(14.969083423670, -27.659158364911, 15.150000000000)
    pts(19) = Array(14.919655863997, -27.429992597518, 15.150000000000)
    FSData_423 = pts
End Function

Private Function FSData_424() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(14.864268312410, -27.203924855371, 18.250000000000)
    pts(1) = Array(15.428600100973, -26.887883868469, 18.250000000000)
    pts(2) = Array(15.592414621585, -27.053229590357, 18.250000000000)
    pts(3) = Array(15.761984730172, -27.215112297505, 18.250000000000)
    pts(4) = Array(15.936886913556, -27.373733039258, 18.250000000000)
    pts(5) = Array(16.116793415163, -27.529238456829, 18.250000000000)
    pts(6) = Array(16.301440819358, -27.681738605315, 18.250000000000)
    pts(7) = Array(16.490611295364, -27.831317595565, 18.250000000000)
    pts(8) = Array(16.684120620166, -27.978040391202, 18.250000000000)
    pts(9) = Array(16.881810128416, -28.121957378321, 18.250000000000)
    pts(10) = Array(17.190623728326, -28.636383427925, 18.250000000000)
    pts(11) = Array(15.434454447870, -29.619885480815, 18.250000000000)
    pts(12) = Array(15.157188799106, -29.087791729663, 18.250000000000)
    pts(13) = Array(15.137771879876, -28.844037295650, 18.250000000000)
    pts(14) = Array(15.113778475758, -28.602381023015, 18.250000000000)
    pts(15) = Array(15.085082723126, -28.362931869545, 18.250000000000)
    pts(16) = Array(15.051532975128, -28.125812967782, 18.250000000000)
    pts(17) = Array(15.012943699204, -27.891166101196, 18.250000000000)
    pts(18) = Array(14.969083423670, -27.659158364911, 18.250000000000)
    pts(19) = Array(14.919655863997, -27.429992597518, 18.250000000000)
    FSData_424 = pts
End Function

Private Function FSData_425() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(16.952842078303, -25.953827183445, 15.150000000000)
    pts(1) = Array(17.490637930252, -25.594483483611, 15.150000000000)
    pts(2) = Array(17.666920341189, -25.746466760666, 15.150000000000)
    pts(3) = Array(17.848668892540, -25.894546120071, 15.150000000000)
    pts(4) = Array(18.035477150824, -26.038955219865, 15.150000000000)
    pts(5) = Array(18.227029875171, -26.179865964700, 15.150000000000)
    pts(6) = Array(18.423073096790, -26.317408737760, 15.150000000000)
    pts(7) = Array(18.623396255666, -26.451684481417, 15.150000000000)
    pts(8) = Array(18.827820793669, -26.582772413003, 15.150000000000)
    pts(9) = Array(19.036192487809, -26.710735212028, 15.150000000000)
    pts(10) = Array(19.384415521123, -27.199346222004, 15.150000000000)
    pts(11) = Array(17.710824606183, -28.317603919982, 15.150000000000)
    pts(12) = Array(17.392666080324, -27.808904448366, 15.150000000000)
    pts(13) = Array(17.354184264503, -27.567424861849, 15.150000000000)
    pts(14) = Array(17.311304691415, -27.328396035644, 15.150000000000)
    pts(15) = Array(17.263910434157, -27.091936466805, 15.150000000000)
    pts(16) = Array(17.211859974224, -26.858180806371, 15.150000000000)
    pts(17) = Array(17.154979475191, -26.627284957460, 15.150000000000)
    pts(18) = Array(17.093051289058, -26.399433661153, 15.150000000000)
    pts(19) = Array(17.025795959174, -26.174852376978, 15.150000000000)
    FSData_425 = pts
End Function

Private Function FSData_426() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(16.952842078303, -25.953827183445, 18.250000000000)
    pts(1) = Array(17.490637930252, -25.594483483611, 18.250000000000)
    pts(2) = Array(17.666920341189, -25.746466760666, 18.250000000000)
    pts(3) = Array(17.848668892540, -25.894546120071, 18.250000000000)
    pts(4) = Array(18.035477150824, -26.038955219865, 18.250000000000)
    pts(5) = Array(18.227029875171, -26.179865964700, 18.250000000000)
    pts(6) = Array(18.423073096790, -26.317408737760, 18.250000000000)
    pts(7) = Array(18.623396255666, -26.451684481417, 18.250000000000)
    pts(8) = Array(18.827820793669, -26.582772413003, 18.250000000000)
    pts(9) = Array(19.036192487809, -26.710735212028, 18.250000000000)
    pts(10) = Array(19.384415521123, -27.199346222004, 18.250000000000)
    pts(11) = Array(17.710824606183, -28.317603919982, 18.250000000000)
    pts(12) = Array(17.392666080324, -27.808904448366, 18.250000000000)
    pts(13) = Array(17.354184264503, -27.567424861849, 18.250000000000)
    pts(14) = Array(17.311304691415, -27.328396035644, 18.250000000000)
    pts(15) = Array(17.263910434157, -27.091936466805, 18.250000000000)
    pts(16) = Array(17.211859974224, -26.858180806371, 18.250000000000)
    pts(17) = Array(17.154979475191, -26.627284957460, 18.250000000000)
    pts(18) = Array(17.093051289058, -26.399433661153, 18.250000000000)
    pts(19) = Array(17.025795959174, -26.174852376978, 18.250000000000)
    FSData_426 = pts
End Function

Private Function FSData_427() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(18.936895935390, -24.543715536410, 15.150000000000)
    pts(1) = Array(19.444840160463, -24.143284597047, 15.150000000000)
    pts(2) = Array(19.632503622043, -24.280968401826, 15.150000000000)
    pts(3) = Array(19.825310075900, -24.414331454996, 15.150000000000)
    pts(4) = Array(20.022872674052, -24.543638582710, 15.150000000000)
    pts(5) = Array(20.224890634896, -24.669085893209, 15.150000000000)
    pts(6) = Array(20.431121002287, -24.790823293910, 15.150000000000)
    pts(7) = Array(20.641361785144, -24.908967936364, 15.150000000000)
    pts(8) = Array(20.855441191093, -25.023612803208, 15.150000000000)
    pts(9) = Array(21.073210390339, -25.134832480932, 15.150000000000)
    pts(10) = Array(21.458695946260, -25.594616001925, 15.150000000000)
    pts(11) = Array(19.878001641353, -26.840734914424, 15.150000000000)
    pts(12) = Array(19.520911791508, -26.358566023746, 15.150000000000)
    pts(13) = Array(19.463602332288, -26.120850086675, 15.150000000000)
    pts(14) = Array(19.402100957056, -25.885922399100, 15.150000000000)
    pts(15) = Array(19.336300396530, -25.653910266764, 15.150000000000)
    pts(16) = Array(19.266070133054, -25.424959029038, 15.150000000000)
    pts(17) = Array(19.191249098104, -25.199237747490, 15.150000000000)
    pts(18) = Array(19.111634809233, -24.976947670172, 15.150000000000)
    pts(19) = Array(19.026966360612, -24.758335487511, 15.150000000000)
    FSData_427 = pts
End Function

Private Function FSData_428() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(18.936895935390, -24.543715536410, 18.250000000000)
    pts(1) = Array(19.444840160463, -24.143284597047, 18.250000000000)
    pts(2) = Array(19.632503622043, -24.280968401826, 18.250000000000)
    pts(3) = Array(19.825310075900, -24.414331454996, 18.250000000000)
    pts(4) = Array(20.022872674052, -24.543638582710, 18.250000000000)
    pts(5) = Array(20.224890634896, -24.669085893209, 18.250000000000)
    pts(6) = Array(20.431121002287, -24.790823293910, 18.250000000000)
    pts(7) = Array(20.641361785144, -24.908967936364, 18.250000000000)
    pts(8) = Array(20.855441191093, -25.023612803208, 18.250000000000)
    pts(9) = Array(21.073210390339, -25.134832480932, 18.250000000000)
    pts(10) = Array(21.458695946260, -25.594616001925, 18.250000000000)
    pts(11) = Array(19.878001641353, -26.840734914424, 18.250000000000)
    pts(12) = Array(19.520911791508, -26.358566023746, 18.250000000000)
    pts(13) = Array(19.463602332288, -26.120850086675, 18.250000000000)
    pts(14) = Array(19.402100957056, -25.885922399100, 18.250000000000)
    pts(15) = Array(19.336300396530, -25.653910266764, 18.250000000000)
    pts(16) = Array(19.266070133054, -25.424959029038, 18.250000000000)
    pts(17) = Array(19.191249098104, -25.199237747490, 18.250000000000)
    pts(18) = Array(19.111634809233, -24.976947670172, 18.250000000000)
    pts(19) = Array(19.026966360612, -24.758335487511, 18.250000000000)
    FSData_428 = pts
End Function

Private Function FSData_429() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(20.804197531879, -22.982283721479, 15.150000000000)
    pts(1) = Array(21.279158485020, -22.543234332486, 15.150000000000)
    pts(2) = Array(21.477045989597, -22.665769798547, 15.150000000000)
    pts(3) = Array(21.679721630058, -22.783594317912, 15.150000000000)
    pts(4) = Array(21.886820528965, -22.897002252105, 15.150000000000)
    pts(5) = Array(22.098058218398, -23.006212703883, 15.150000000000)
    pts(6) = Array(22.313204252768, -23.111394180669, 15.150000000000)
    pts(7) = Array(22.532066455268, -23.212679320910, 15.150000000000)
    pts(8) = Array(22.754480858436, -23.310174299286, 15.150000000000)
    pts(9) = Array(22.980304943259, -23.403965149411, 15.150000000000)
    pts(10) = Array(23.400676375148, -23.832086463120, 15.150000000000)
    pts(11) = Array(21.922624186296, -25.198383852668, 15.150000000000)
    pts(12) = Array(21.528804590135, -24.745718274477, 15.150000000000)
    pts(13) = Array(21.453020819389, -24.513231584653, 15.150000000000)
    pts(14) = Array(21.373276818444, -24.283853443022, 15.150000000000)
    pts(15) = Array(21.289475636985, -24.057719179137, 15.150000000000)
    pts(16) = Array(21.201498562895, -23.834983924634, 15.150000000000)
    pts(17) = Array(21.109198288589, -23.615828857210, 15.150000000000)
    pts(18) = Array(21.012388745545, -23.400470491127, 15.150000000000)
    pts(19) = Array(20.910829187328, -23.189175226782, 15.150000000000)
    FSData_429 = pts
End Function

Private Function FSData_430() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(20.804197531879, -22.982283721479, 18.250000000000)
    pts(1) = Array(21.279158485020, -22.543234332486, 18.250000000000)
    pts(2) = Array(21.477045989597, -22.665769798547, 18.250000000000)
    pts(3) = Array(21.679721630058, -22.783594317912, 18.250000000000)
    pts(4) = Array(21.886820528965, -22.897002252105, 18.250000000000)
    pts(5) = Array(22.098058218398, -23.006212703883, 18.250000000000)
    pts(6) = Array(22.313204252768, -23.111394180669, 18.250000000000)
    pts(7) = Array(22.532066455268, -23.212679320910, 18.250000000000)
    pts(8) = Array(22.754480858436, -23.310174299286, 18.250000000000)
    pts(9) = Array(22.980304943259, -23.403965149411, 18.250000000000)
    pts(10) = Array(23.400676375148, -23.832086463120, 18.250000000000)
    pts(11) = Array(21.922624186296, -25.198383852668, 18.250000000000)
    pts(12) = Array(21.528804590135, -24.745718274477, 18.250000000000)
    pts(13) = Array(21.453020819389, -24.513231584653, 18.250000000000)
    pts(14) = Array(21.373276818444, -24.283853443022, 18.250000000000)
    pts(15) = Array(21.289475636985, -24.057719179137, 18.250000000000)
    pts(16) = Array(21.201498562895, -23.834983924634, 18.250000000000)
    pts(17) = Array(21.109198288589, -23.615828857210, 18.250000000000)
    pts(18) = Array(21.012388745545, -23.400470491127, 18.250000000000)
    pts(19) = Array(20.910829187328, -23.189175226782, 18.250000000000)
    FSData_430 = pts
End Function

Private Function FSData_431() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(22.543234332486, -21.279158485020, 15.150000000000)
    pts(1) = Array(22.982283721479, -20.804197531879, 15.150000000000)
    pts(2) = Array(23.189175226782, -20.910829187328, 15.150000000000)
    pts(3) = Array(23.400470491127, -21.012388745545, 15.150000000000)
    pts(4) = Array(23.615828857210, -21.109198288589, 15.150000000000)
    pts(5) = Array(23.834983924634, -21.201498562895, 15.150000000000)
    pts(6) = Array(24.057719179137, -21.289475636985, 15.150000000000)
    pts(7) = Array(24.283853443022, -21.373276818444, 15.150000000000)
    pts(8) = Array(24.513231584653, -21.453020819389, 15.150000000000)
    pts(9) = Array(24.745718274477, -21.528804590135, 15.150000000000)
    pts(10) = Array(25.198383852668, -21.922624186296, 15.150000000000)
    pts(11) = Array(23.832086463120, -23.400676375148, 15.150000000000)
    pts(12) = Array(23.403965149411, -22.980304943259, 15.150000000000)
    pts(13) = Array(23.310174299286, -22.754480858436, 15.150000000000)
    pts(14) = Array(23.212679320910, -22.532066455268, 15.150000000000)
    pts(15) = Array(23.111394180669, -22.313204252768, 15.150000000000)
    pts(16) = Array(23.006212703883, -22.098058218398, 15.150000000000)
    pts(17) = Array(22.897002252105, -21.886820528965, 15.150000000000)
    pts(18) = Array(22.783594317912, -21.679721630058, 15.150000000000)
    pts(19) = Array(22.665769798547, -21.477045989597, 15.150000000000)
    FSData_431 = pts
End Function

Private Function FSData_432() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(22.543234332486, -21.279158485020, 18.250000000000)
    pts(1) = Array(22.982283721479, -20.804197531879, 18.250000000000)
    pts(2) = Array(23.189175226782, -20.910829187328, 18.250000000000)
    pts(3) = Array(23.400470491127, -21.012388745545, 18.250000000000)
    pts(4) = Array(23.615828857210, -21.109198288589, 18.250000000000)
    pts(5) = Array(23.834983924634, -21.201498562895, 18.250000000000)
    pts(6) = Array(24.057719179137, -21.289475636985, 18.250000000000)
    pts(7) = Array(24.283853443022, -21.373276818444, 18.250000000000)
    pts(8) = Array(24.513231584653, -21.453020819389, 18.250000000000)
    pts(9) = Array(24.745718274477, -21.528804590135, 18.250000000000)
    pts(10) = Array(25.198383852668, -21.922624186296, 18.250000000000)
    pts(11) = Array(23.832086463120, -23.400676375148, 18.250000000000)
    pts(12) = Array(23.403965149411, -22.980304943259, 18.250000000000)
    pts(13) = Array(23.310174299286, -22.754480858436, 18.250000000000)
    pts(14) = Array(23.212679320910, -22.532066455268, 18.250000000000)
    pts(15) = Array(23.111394180669, -22.313204252768, 18.250000000000)
    pts(16) = Array(23.006212703883, -22.098058218398, 18.250000000000)
    pts(17) = Array(22.897002252105, -21.886820528965, 18.250000000000)
    pts(18) = Array(22.783594317912, -21.679721630058, 18.250000000000)
    pts(19) = Array(22.665769798547, -21.477045989597, 18.250000000000)
    FSData_432 = pts
End Function

Private Function FSData_433() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(24.143284597047, -19.444840160463, 15.150000000000)
    pts(1) = Array(24.543715536410, -18.936895935390, 15.150000000000)
    pts(2) = Array(24.758335487511, -19.026966360612, 15.150000000000)
    pts(3) = Array(24.976947670172, -19.111634809233, 15.150000000000)
    pts(4) = Array(25.199237747490, -19.191249098104, 15.150000000000)
    pts(5) = Array(25.424959029038, -19.266070133054, 15.150000000000)
    pts(6) = Array(25.653910266764, -19.336300396530, 15.150000000000)
    pts(7) = Array(25.885922399100, -19.402100957056, 15.150000000000)
    pts(8) = Array(26.120850086675, -19.463602332288, 15.150000000000)
    pts(9) = Array(26.358566023746, -19.520911791508, 15.150000000000)
    pts(10) = Array(26.840734914424, -19.878001641353, 15.150000000000)
    pts(11) = Array(25.594616001925, -21.458695946260, 15.150000000000)
    pts(12) = Array(25.134832480932, -21.073210390339, 15.150000000000)
    pts(13) = Array(25.023612803208, -20.855441191093, 15.150000000000)
    pts(14) = Array(24.908967936364, -20.641361785144, 15.150000000000)
    pts(15) = Array(24.790823293910, -20.431121002287, 15.150000000000)
    pts(16) = Array(24.669085893209, -20.224890634896, 15.150000000000)
    pts(17) = Array(24.543638582710, -20.022872674052, 15.150000000000)
    pts(18) = Array(24.414331454996, -19.825310075900, 15.150000000000)
    pts(19) = Array(24.280968401826, -19.632503622043, 15.150000000000)
    FSData_433 = pts
End Function

Private Function FSData_434() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(24.143284597047, -19.444840160463, 18.250000000000)
    pts(1) = Array(24.543715536410, -18.936895935390, 18.250000000000)
    pts(2) = Array(24.758335487511, -19.026966360612, 18.250000000000)
    pts(3) = Array(24.976947670172, -19.111634809233, 18.250000000000)
    pts(4) = Array(25.199237747490, -19.191249098104, 18.250000000000)
    pts(5) = Array(25.424959029038, -19.266070133054, 18.250000000000)
    pts(6) = Array(25.653910266764, -19.336300396530, 18.250000000000)
    pts(7) = Array(25.885922399100, -19.402100957056, 18.250000000000)
    pts(8) = Array(26.120850086675, -19.463602332288, 18.250000000000)
    pts(9) = Array(26.358566023746, -19.520911791508, 18.250000000000)
    pts(10) = Array(26.840734914424, -19.878001641353, 18.250000000000)
    pts(11) = Array(25.594616001925, -21.458695946260, 18.250000000000)
    pts(12) = Array(25.134832480932, -21.073210390339, 18.250000000000)
    pts(13) = Array(25.023612803208, -20.855441191093, 18.250000000000)
    pts(14) = Array(24.908967936364, -20.641361785144, 18.250000000000)
    pts(15) = Array(24.790823293910, -20.431121002287, 18.250000000000)
    pts(16) = Array(24.669085893209, -20.224890634896, 18.250000000000)
    pts(17) = Array(24.543638582710, -20.022872674052, 18.250000000000)
    pts(18) = Array(24.414331454996, -19.825310075900, 18.250000000000)
    pts(19) = Array(24.280968401826, -19.632503622043, 18.250000000000)
    FSData_434 = pts
End Function

Private Function FSData_435() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(25.594483483611, -17.490637930252, 15.150000000000)
    pts(1) = Array(25.953827183445, -16.952842078303, 15.150000000000)
    pts(2) = Array(26.174852376978, -17.025795959174, 15.150000000000)
    pts(3) = Array(26.399433661153, -17.093051289058, 15.150000000000)
    pts(4) = Array(26.627284957460, -17.154979475191, 15.150000000000)
    pts(5) = Array(26.858180806371, -17.211859974224, 15.150000000000)
    pts(6) = Array(27.091936466805, -17.263910434158, 15.150000000000)
    pts(7) = Array(27.328396035644, -17.311304691415, 15.150000000000)
    pts(8) = Array(27.567424861849, -17.354184264503, 15.150000000000)
    pts(9) = Array(27.808904448366, -17.392666080324, 15.150000000000)
    pts(10) = Array(28.317603919982, -17.710824606183, 15.150000000000)
    pts(11) = Array(27.199346222004, -19.384415521123, 15.150000000000)
    pts(12) = Array(26.710735212028, -19.036192487809, 15.150000000000)
    pts(13) = Array(26.582772413003, -18.827820793669, 15.150000000000)
    pts(14) = Array(26.451684481417, -18.623396255666, 15.150000000000)
    pts(15) = Array(26.317408737760, -18.423073096790, 15.150000000000)
    pts(16) = Array(26.179865964700, -18.227029875171, 15.150000000000)
    pts(17) = Array(26.038955219865, -18.035477150824, 15.150000000000)
    pts(18) = Array(25.894546120071, -17.848668892540, 15.150000000000)
    pts(19) = Array(25.746466760666, -17.666920341190, 15.150000000000)
    FSData_435 = pts
End Function

Private Function FSData_436() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(25.594483483611, -17.490637930252, 18.250000000000)
    pts(1) = Array(25.953827183445, -16.952842078303, 18.250000000000)
    pts(2) = Array(26.174852376978, -17.025795959174, 18.250000000000)
    pts(3) = Array(26.399433661153, -17.093051289058, 18.250000000000)
    pts(4) = Array(26.627284957460, -17.154979475191, 18.250000000000)
    pts(5) = Array(26.858180806371, -17.211859974224, 18.250000000000)
    pts(6) = Array(27.091936466805, -17.263910434158, 18.250000000000)
    pts(7) = Array(27.328396035644, -17.311304691415, 18.250000000000)
    pts(8) = Array(27.567424861849, -17.354184264503, 18.250000000000)
    pts(9) = Array(27.808904448366, -17.392666080324, 18.250000000000)
    pts(10) = Array(28.317603919982, -17.710824606183, 18.250000000000)
    pts(11) = Array(27.199346222004, -19.384415521123, 18.250000000000)
    pts(12) = Array(26.710735212028, -19.036192487809, 18.250000000000)
    pts(13) = Array(26.582772413003, -18.827820793669, 18.250000000000)
    pts(14) = Array(26.451684481417, -18.623396255666, 18.250000000000)
    pts(15) = Array(26.317408737760, -18.423073096790, 18.250000000000)
    pts(16) = Array(26.179865964700, -18.227029875171, 18.250000000000)
    pts(17) = Array(26.038955219865, -18.035477150824, 18.250000000000)
    pts(18) = Array(25.894546120071, -17.848668892540, 18.250000000000)
    pts(19) = Array(25.746466760666, -17.666920341190, 18.250000000000)
    FSData_436 = pts
End Function

Private Function FSData_437() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(26.887883868469, -15.428600100973, 15.150000000000)
    pts(1) = Array(27.203924855371, -14.864268312410, 15.150000000000)
    pts(2) = Array(27.429992597518, -14.919655863997, 15.150000000000)
    pts(3) = Array(27.659158364911, -14.969083423670, 15.150000000000)
    pts(4) = Array(27.891166101196, -15.012943699204, 15.150000000000)
    pts(5) = Array(28.125812967782, -15.051532975128, 15.150000000000)
    pts(6) = Array(28.362931869545, -15.085082723126, 15.150000000000)
    pts(7) = Array(28.602381023015, -15.113778475758, 15.150000000000)
    pts(8) = Array(28.844037295650, -15.137771879876, 15.150000000000)
    pts(9) = Array(29.087791729663, -15.157188799106, 15.150000000000)
    pts(10) = Array(29.619885480815, -15.434454447870, 15.150000000000)
    pts(11) = Array(28.636383427925, -17.190623728326, 15.150000000000)
    pts(12) = Array(28.121957378321, -16.881810128416, 15.150000000000)
    pts(13) = Array(27.978040391202, -16.684120620166, 15.150000000000)
    pts(14) = Array(27.831317595565, -16.490611295364, 15.150000000000)
    pts(15) = Array(27.681738605315, -16.301440819358, 15.150000000000)
    pts(16) = Array(27.529238456829, -16.116793415163, 15.150000000000)
    pts(17) = Array(27.373733039258, -15.936886913556, 15.150000000000)
    pts(18) = Array(27.215112297505, -15.761984730172, 15.150000000000)
    pts(19) = Array(27.053229590357, -15.592414621585, 15.150000000000)
    FSData_437 = pts
End Function

Private Function FSData_438() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(26.887883868469, -15.428600100973, 18.250000000000)
    pts(1) = Array(27.203924855371, -14.864268312410, 18.250000000000)
    pts(2) = Array(27.429992597518, -14.919655863997, 18.250000000000)
    pts(3) = Array(27.659158364911, -14.969083423670, 18.250000000000)
    pts(4) = Array(27.891166101196, -15.012943699204, 18.250000000000)
    pts(5) = Array(28.125812967782, -15.051532975128, 18.250000000000)
    pts(6) = Array(28.362931869545, -15.085082723126, 18.250000000000)
    pts(7) = Array(28.602381023015, -15.113778475758, 18.250000000000)
    pts(8) = Array(28.844037295650, -15.137771879876, 18.250000000000)
    pts(9) = Array(29.087791729663, -15.157188799106, 18.250000000000)
    pts(10) = Array(29.619885480815, -15.434454447870, 18.250000000000)
    pts(11) = Array(28.636383427925, -17.190623728326, 18.250000000000)
    pts(12) = Array(28.121957378321, -16.881810128416, 18.250000000000)
    pts(13) = Array(27.978040391202, -16.684120620166, 18.250000000000)
    pts(14) = Array(27.831317595565, -16.490611295364, 18.250000000000)
    pts(15) = Array(27.681738605315, -16.301440819358, 18.250000000000)
    pts(16) = Array(27.529238456829, -16.116793415163, 18.250000000000)
    pts(17) = Array(27.373733039258, -15.936886913556, 18.250000000000)
    pts(18) = Array(27.215112297505, -15.761984730172, 18.250000000000)
    pts(19) = Array(27.053229590357, -15.592414621585, 18.250000000000)
    FSData_438 = pts
End Function

Private Function FSData_439() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(28.015511508149, -13.271439821541, 15.150000000000)
    pts(1) = Array(28.286301284341, -12.684051389501, 15.150000000000)
    pts(2) = Array(28.516017792297, -12.721531129129, 15.150000000000)
    pts(3) = Array(28.748355159745, -12.752826181250, 15.150000000000)
    pts(4) = Array(28.983088931165, -12.778348133001, 15.150000000000)
    pts(5) = Array(29.220040139464, -12.798408270098, 15.150000000000)
    pts(6) = Array(29.459060365677, -12.813250460808, 15.150000000000)
    pts(7) = Array(29.700022820122, -12.823070789957, 15.150000000000)
    pts(8) = Array(29.942816647908, -12.828030097789, 15.150000000000)
    pts(9) = Array(30.187343102274, -12.828262408665, 15.150000000000)
    pts(10) = Array(30.739550598048, -13.062925745409, 15.150000000000)
    pts(11) = Array(29.896867807448, -14.890846023781, 15.150000000000)
    pts(12) = Array(29.359798325877, -14.623345795809, 15.150000000000)
    pts(13) = Array(29.200814446748, -14.437557295007, 15.150000000000)
    pts(14) = Array(29.039361381885, -14.256156232741, 15.150000000000)
    pts(15) = Array(28.875401349256, -14.079304738495, 15.150000000000)
    pts(16) = Array(28.708884039472, -13.907191564372, 15.150000000000)
    pts(17) = Array(28.539742691774, -13.740040468913, 15.150000000000)
    pts(18) = Array(28.367888257682, -13.578122690552, 15.150000000000)
    pts(19) = Array(28.193200263512, -13.421776480835, 15.150000000000)
    FSData_439 = pts
End Function

Private Function FSData_440() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(28.015511508149, -13.271439821541, 18.250000000000)
    pts(1) = Array(28.286301284341, -12.684051389501, 18.250000000000)
    pts(2) = Array(28.516017792297, -12.721531129129, 18.250000000000)
    pts(3) = Array(28.748355159745, -12.752826181250, 18.250000000000)
    pts(4) = Array(28.983088931165, -12.778348133001, 18.250000000000)
    pts(5) = Array(29.220040139464, -12.798408270098, 18.250000000000)
    pts(6) = Array(29.459060365677, -12.813250460808, 18.250000000000)
    pts(7) = Array(29.700022820122, -12.823070789957, 18.250000000000)
    pts(8) = Array(29.942816647908, -12.828030097789, 18.250000000000)
    pts(9) = Array(30.187343102274, -12.828262408665, 18.250000000000)
    pts(10) = Array(30.739550598048, -13.062925745409, 18.250000000000)
    pts(11) = Array(29.896867807448, -14.890846023781, 18.250000000000)
    pts(12) = Array(29.359798325877, -14.623345795809, 18.250000000000)
    pts(13) = Array(29.200814446748, -14.437557295007, 18.250000000000)
    pts(14) = Array(29.039361381885, -14.256156232741, 18.250000000000)
    pts(15) = Array(28.875401349256, -14.079304738495, 18.250000000000)
    pts(16) = Array(28.708884039472, -13.907191564372, 18.250000000000)
    pts(17) = Array(28.539742691774, -13.740040468913, 18.250000000000)
    pts(18) = Array(28.367888257682, -13.578122690552, 18.250000000000)
    pts(19) = Array(28.193200263512, -13.421776480835, 18.250000000000)
    FSData_440 = pts
End Function

Private Function FSData_441() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(28.970414203278, -11.032456702408, 15.150000000000)
    pts(1) = Array(29.194283259743, -10.425633071901, 15.150000000000)
    pts(2) = Array(29.426232254848, -10.444973924511, 15.150000000000)
    pts(3) = Array(29.660308785221, -10.457943524677, 15.150000000000)
    pts(4) = Array(29.896321380018, -10.464969801327, 15.150000000000)
    pts(5) = Array(30.134116047037, -10.466377122181, 15.150000000000)
    pts(6) = Array(30.373563958524, -10.462420248560, 15.150000000000)
    pts(7) = Array(30.614554100283, -10.453304608632, 15.150000000000)
    pts(8) = Array(30.856988578534, -10.439199244397, 15.150000000000)
    pts(9) = Array(31.100779466350, -10.420245514645, 15.150000000000)
    pts(10) = Array(31.669696163905, -10.610859761864, 15.150000000000)
    pts(11) = Array(30.973028055222, -12.499261301791, 15.150000000000)
    pts(12) = Array(30.416626353631, -12.274723673615, 15.150000000000)
    pts(13) = Array(30.243555770973, -12.101981628149, 15.150000000000)
    pts(14) = Array(30.068367848717, -11.933807226290, 15.150000000000)
    pts(15) = Array(29.891037641833, -11.770365062076, 15.150000000000)
    pts(16) = Array(29.711529805339, -11.611847252978, 15.150000000000)
    pts(17) = Array(29.529795340167, -11.458482105752, 15.150000000000)
    pts(18) = Array(29.345766753470, -11.310547009360, 15.150000000000)
    pts(19) = Array(29.159350481852, -11.168388624887, 15.150000000000)
    FSData_441 = pts
End Function

Private Function FSData_442() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(28.970414203278, -11.032456702408, 18.250000000000)
    pts(1) = Array(29.194283259743, -10.425633071901, 18.250000000000)
    pts(2) = Array(29.426232254848, -10.444973924511, 18.250000000000)
    pts(3) = Array(29.660308785221, -10.457943524677, 18.250000000000)
    pts(4) = Array(29.896321380018, -10.464969801327, 18.250000000000)
    pts(5) = Array(30.134116047037, -10.466377122181, 18.250000000000)
    pts(6) = Array(30.373563958524, -10.462420248560, 18.250000000000)
    pts(7) = Array(30.614554100283, -10.453304608632, 18.250000000000)
    pts(8) = Array(30.856988578534, -10.439199244397, 18.250000000000)
    pts(9) = Array(31.100779466350, -10.420245514645, 18.250000000000)
    pts(10) = Array(31.669696163905, -10.610859761864, 18.250000000000)
    pts(11) = Array(30.973028055222, -12.499261301791, 18.250000000000)
    pts(12) = Array(30.416626353631, -12.274723673615, 18.250000000000)
    pts(13) = Array(30.243555770973, -12.101981628149, 18.250000000000)
    pts(14) = Array(30.068367848717, -11.933807226290, 18.250000000000)
    pts(15) = Array(29.891037641833, -11.770365062076, 18.250000000000)
    pts(16) = Array(29.711529805339, -11.611847252978, 18.250000000000)
    pts(17) = Array(29.529795340167, -11.458482105752, 18.250000000000)
    pts(18) = Array(29.345766753470, -11.310547009360, 18.250000000000)
    pts(19) = Array(29.159350481852, -11.168388624887, 18.250000000000)
    FSData_442 = pts
End Function

Private Function FSData_443() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(29.746704661204, -8.725454819040, 15.150000000000)
    pts(1) = Array(29.922272770764, -8.102937259537, 15.150000000000)
    pts(2) = Array(30.155024210333, -8.104019982343, 15.150000000000)
    pts(3) = Array(30.389396743983, -8.098584168655, 15.150000000000)
    pts(4) = Array(30.625233066028, -8.087071450871, 15.150000000000)
    pts(5) = Array(30.862405108570, -8.069817278881, 15.150000000000)
    pts(6) = Array(31.100804429331, -8.047085736372, 15.150000000000)
    pts(7) = Array(31.340336474044, -8.019090328318, 15.150000000000)
    pts(8) = Array(31.580916913584, -7.986007256277, 15.150000000000)
    pts(9) = Array(31.822469182957, -7.947984341944, 15.150000000000)
    pts(10) = Array(32.404587521669, -8.093374299418, 15.150000000000)
    pts(11) = Array(31.858229285457, -10.030614477449, 15.150000000000)
    pts(12) = Array(31.285925765359, -9.850423798214, 15.150000000000)
    pts(13) = Array(31.099835516867, -9.691793220237, 15.150000000000)
    pts(14) = Array(30.911992829015, -9.537882329896, 15.150000000000)
    pts(15) = Array(30.722385747570, -9.388857171005, 15.150000000000)
    pts(16) = Array(30.530994109869, -9.244912041937, 15.150000000000)
    pts(17) = Array(30.337786980636, -9.106278390077, 15.150000000000)
    pts(18) = Array(30.142718838765, -8.973238044717, 15.150000000000)
    pts(19) = Array(29.945723608004, -8.846143939200, 15.150000000000)
    FSData_443 = pts
End Function

Private Function FSData_444() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(29.746704661204, -8.725454819040, 18.250000000000)
    pts(1) = Array(29.922272770764, -8.102937259537, 18.250000000000)
    pts(2) = Array(30.155024210333, -8.104019982343, 18.250000000000)
    pts(3) = Array(30.389396743983, -8.098584168655, 18.250000000000)
    pts(4) = Array(30.625233066028, -8.087071450871, 18.250000000000)
    pts(5) = Array(30.862405108570, -8.069817278881, 18.250000000000)
    pts(6) = Array(31.100804429331, -8.047085736372, 18.250000000000)
    pts(7) = Array(31.340336474044, -8.019090328318, 18.250000000000)
    pts(8) = Array(31.580916913584, -7.986007256277, 18.250000000000)
    pts(9) = Array(31.822469182957, -7.947984341944, 18.250000000000)
    pts(10) = Array(32.404587521669, -8.093374299418, 18.250000000000)
    pts(11) = Array(31.858229285457, -10.030614477449, 18.250000000000)
    pts(12) = Array(31.285925765359, -9.850423798214, 18.250000000000)
    pts(13) = Array(31.099835516867, -9.691793220237, 18.250000000000)
    pts(14) = Array(30.911992829015, -9.537882329896, 18.250000000000)
    pts(15) = Array(30.722385747570, -9.388857171005, 18.250000000000)
    pts(16) = Array(30.530994109869, -9.244912041937, 18.250000000000)
    pts(17) = Array(30.337786980636, -9.106278390077, 18.250000000000)
    pts(18) = Array(30.142718838765, -8.973238044717, 18.250000000000)
    pts(19) = Array(29.945723608004, -8.846143939200, 18.250000000000)
    FSData_444 = pts
End Function

Private Function FSData_445() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.339596793110, -6.364657605205, 15.150000000000)
    pts(1) = Array(30.465781519988, -5.730284144469, 15.150000000000)
    pts(2) = Array(30.697900413998, -5.713102062123, 15.150000000000)
    pts(3) = Array(30.931123966318, -5.689294348182, 15.150000000000)
    pts(4) = Array(31.165330006262, -5.659313615697, 15.150000000000)
    pts(5) = Array(31.400417179817, -5.623504328571, 15.150000000000)
    pts(6) = Array(31.636298098764, -5.582138264692, 15.150000000000)
    pts(7) = Array(31.872895251723, -5.535435689512, 15.150000000000)
    pts(8) = Array(32.110138394141, -5.483578877805, 15.150000000000)
    pts(9) = Array(32.347962795006, -5.426721202801, 15.150000000000)
    pts(10) = Array(32.939693821744, -5.525990493096, 15.150000000000)
    pts(11) = Array(32.547013938212, -7.500125579336, 15.150000000000)
    pts(12) = Array(31.962337041119, -7.365392784498, 15.150000000000)
    pts(13) = Array(31.764374435055, -7.221851684273, 15.150000000000)
    pts(14) = Array(31.565035094241, -7.083153217270, 15.150000000000)
    pts(15) = Array(31.364320128943, -6.949463853363, 15.150000000000)
    pts(16) = Array(31.162224683126, -6.820978873913, 15.150000000000)
    pts(17) = Array(30.958736076030, -6.697931439982, 15.150000000000)
    pts(18) = Array(30.753831038948, -6.580606083623, 15.150000000000)
    pts(19) = Array(30.547471390148, -6.469359834487, 15.150000000000)
    FSData_445 = pts
End Function

Private Function FSData_446() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.339596793110, -6.364657605205, 18.250000000000)
    pts(1) = Array(30.465781519988, -5.730284144469, 18.250000000000)
    pts(2) = Array(30.697900413998, -5.713102062123, 18.250000000000)
    pts(3) = Array(30.931123966318, -5.689294348182, 18.250000000000)
    pts(4) = Array(31.165330006262, -5.659313615697, 18.250000000000)
    pts(5) = Array(31.400417179817, -5.623504328571, 18.250000000000)
    pts(6) = Array(31.636298098764, -5.582138264692, 18.250000000000)
    pts(7) = Array(31.872895251723, -5.535435689512, 18.250000000000)
    pts(8) = Array(32.110138394141, -5.483578877805, 18.250000000000)
    pts(9) = Array(32.347962795006, -5.426721202801, 18.250000000000)
    pts(10) = Array(32.939693821744, -5.525990493096, 18.250000000000)
    pts(11) = Array(32.547013938212, -7.500125579336, 18.250000000000)
    pts(12) = Array(31.962337041119, -7.365392784498, 18.250000000000)
    pts(13) = Array(31.764374435055, -7.221851684273, 18.250000000000)
    pts(14) = Array(31.565035094241, -7.083153217270, 18.250000000000)
    pts(15) = Array(31.364320128943, -6.949463853363, 18.250000000000)
    pts(16) = Array(31.162224683126, -6.820978873913, 18.250000000000)
    pts(17) = Array(30.958736076030, -6.697931439982, 18.250000000000)
    pts(18) = Array(30.753831038948, -6.580606083623, 18.250000000000)
    pts(19) = Array(30.547471390148, -6.469359834487, 18.250000000000)
    FSData_446 = pts
End Function

Private Function FSData_447() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.745435221847, -3.964620160771, 15.150000000000)
    pts(1) = Array(30.821458595241, -3.322301922137, 15.150000000000)
    pts(2) = Array(31.051513853523, -3.286960967891, 15.150000000000)
    pts(3) = Array(31.282150523759, -3.244928136169, 15.150000000000)
    pts(4) = Array(31.513282323483, -3.196664230169, 15.150000000000)
    pts(5) = Array(31.744835237452, -3.142520604070, 15.150000000000)
    pts(6) = Array(31.976743470282, -3.082775054361, 15.150000000000)
    pts(7) = Array(32.208947031361, -3.017653248961, 15.150000000000)
    pts(8) = Array(32.441390193794, -2.947342412076, 15.150000000000)
    pts(9) = Array(32.674020459634, -2.872000522875, 15.150000000000)
    pts(10) = Array(33.271715955847, -2.924537117806, 15.150000000000)
    pts(11) = Array(33.035135427058, -4.923395912982, 15.150000000000)
    pts(12) = Array(32.441689880464, -4.834951675024, 15.150000000000)
    pts(13) = Array(32.233075422124, -4.707385031166, 15.150000000000)
    pts(14) = Array(32.023468421672, -4.584754109672, 15.150000000000)
    pts(15) = Array(31.812883047026, -4.467224780134, 15.150000000000)
    pts(16) = Array(31.601329778721, -4.354992102925, 15.150000000000)
    pts(17) = Array(31.388814268692, -4.248289515272, 15.150000000000)
    pts(18) = Array(31.175335644090, -4.147402497751, 15.150000000000)
    pts(19) = Array(30.960883853107, -4.052689975114, 15.150000000000)
    FSData_447 = pts
End Function

Private Function FSData_448() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.745435221847, -3.964620160771, 18.250000000000)
    pts(1) = Array(30.821458595241, -3.322301922137, 18.250000000000)
    pts(2) = Array(31.051513853523, -3.286960967891, 18.250000000000)
    pts(3) = Array(31.282150523759, -3.244928136169, 18.250000000000)
    pts(4) = Array(31.513282323483, -3.196664230169, 18.250000000000)
    pts(5) = Array(31.744835237452, -3.142520604070, 18.250000000000)
    pts(6) = Array(31.976743470282, -3.082775054361, 18.250000000000)
    pts(7) = Array(32.208947031361, -3.017653248961, 18.250000000000)
    pts(8) = Array(32.441390193794, -2.947342412076, 18.250000000000)
    pts(9) = Array(32.674020459634, -2.872000522875, 18.250000000000)
    pts(10) = Array(33.271715955847, -2.924537117806, 18.250000000000)
    pts(11) = Array(33.035135427058, -4.923395912982, 18.250000000000)
    pts(12) = Array(32.441689880464, -4.834951675024, 18.250000000000)
    pts(13) = Array(32.233075422124, -4.707385031166, 18.250000000000)
    pts(14) = Array(32.023468421672, -4.584754109672, 18.250000000000)
    pts(15) = Array(31.812883047026, -4.467224780134, 18.250000000000)
    pts(16) = Array(31.601329778721, -4.354992102925, 18.250000000000)
    pts(17) = Array(31.388814268692, -4.248289515272, 18.250000000000)
    pts(18) = Array(31.175335644090, -4.147402497751, 18.250000000000)
    pts(19) = Array(30.960883853107, -4.052689975114, 18.250000000000)
    FSData_448 = pts
End Function

Private Function FSData_449() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.961717818547, -1.540139514675, 15.150000000000)
    pts(1) = Array(30.987111129080, -0.893836603677, 15.150000000000)
    pts(2) = Array(31.213684384465, -0.840554666267, 15.150000000000)
    pts(3) = Array(31.440312220850, -0.780555863149, 15.150000000000)
    pts(4) = Array(31.666944775951, -0.714306346663, 15.150000000000)
    pts(5) = Array(31.893535829620, -0.642162195051, 15.150000000000)
    pts(6) = Array(32.120041584959, -0.564405510693, 15.150000000000)
    pts(7) = Array(32.346419941990, -0.481265972658, 15.150000000000)
    pts(8) = Array(32.572630035045, -0.392934600286, 15.150000000000)
    pts(9) = Array(32.798631922915, -0.299573004689, 15.150000000000)
    pts(10) = Array(33.398606897115, -0.305052998677, 15.150000000000)
    pts(11) = Array(33.319584320698, -2.316311873629, 15.150000000000)
    pts(12) = Array(32.721028913739, -2.274701480689, 15.150000000000)
    pts(13) = Array(32.503048780631, -2.163895783978, 15.150000000000)
    pts(14) = Array(32.284466417400, -2.058088468403, 15.150000000000)
    pts(15) = Array(32.065308962267, -1.957443780632, 15.150000000000)
    pts(16) = Array(31.845602167721, -1.862155357441, 15.150000000000)
    pts(17) = Array(31.625369983547, -1.772455473001, 15.150000000000)
    pts(18) = Array(31.404633938134, -1.688628796329, 15.150000000000)
    pts(19) = Array(31.183412171574, -1.611033934388, 15.150000000000)
    FSData_449 = pts
End Function

Private Function FSData_450() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(30.961717818547, -1.540139514675, 18.250000000000)
    pts(1) = Array(30.987111129080, -0.893836603677, 18.250000000000)
    pts(2) = Array(31.213684384465, -0.840554666267, 18.250000000000)
    pts(3) = Array(31.440312220850, -0.780555863149, 18.250000000000)
    pts(4) = Array(31.666944775951, -0.714306346663, 18.250000000000)
    pts(5) = Array(31.893535829620, -0.642162195051, 18.250000000000)
    pts(6) = Array(32.120041584959, -0.564405510693, 18.250000000000)
    pts(7) = Array(32.346419941990, -0.481265972658, 18.250000000000)
    pts(8) = Array(32.572630035045, -0.392934600286, 18.250000000000)
    pts(9) = Array(32.798631922915, -0.299573004689, 18.250000000000)
    pts(10) = Array(33.398606897115, -0.305052998677, 18.250000000000)
    pts(11) = Array(33.319584320698, -2.316311873629, 18.250000000000)
    pts(12) = Array(32.721028913739, -2.274701480689, 18.250000000000)
    pts(13) = Array(32.503048780631, -2.163895783978, 18.250000000000)
    pts(14) = Array(32.284466417400, -2.058088468403, 18.250000000000)
    pts(15) = Array(32.065308962267, -1.957443780632, 18.250000000000)
    pts(16) = Array(31.845602167721, -1.862155357441, 18.250000000000)
    pts(17) = Array(31.625369983547, -1.772455473001, 18.250000000000)
    pts(18) = Array(31.404633938134, -1.688628796329, 18.250000000000)
    pts(19) = Array(31.183412171574, -1.611033934388, 18.250000000000)
    FSData_450 = pts
End Function

Private Function FSData_451() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.200000000000, -1.250000000000, 0.434452102183)
    pts(1) = Array(-25.500000000000, -1.250000000000, 0.434452102183)
    pts(2) = Array(-25.500000000000, 1.250000000000, 0.434452102183)
    pts(3) = Array(-24.200000000000, 1.250000000000, 0.434452102183)
    FSData_451 = pts
End Function

Private Function FSData_452() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-24.200000000000, -1.250000000000, 1.284452102183)
    pts(1) = Array(-25.500000000000, -1.250000000000, 1.284452102183)
    pts(2) = Array(-25.500000000000, 1.250000000000, 1.284452102183)
    pts(3) = Array(-24.200000000000, 1.250000000000, 1.284452102183)
    FSData_452 = pts
End Function

Private Function FSData_453() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.200000000000, -1.250000000000, 0.434452102183)
    pts(1) = Array(25.500000000000, -1.250000000000, 0.434452102183)
    pts(2) = Array(25.500000000000, 1.250000000000, 0.434452102183)
    pts(3) = Array(24.200000000000, 1.250000000000, 0.434452102183)
    FSData_453 = pts
End Function

Private Function FSData_454() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(24.200000000000, -1.250000000000, 1.284452102183)
    pts(1) = Array(25.500000000000, -1.250000000000, 1.284452102183)
    pts(2) = Array(25.500000000000, 1.250000000000, 1.284452102183)
    pts(3) = Array(24.200000000000, 1.250000000000, 1.284452102183)
    FSData_454 = pts
End Function

Private Function FSData_455() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-18.800000000000, -1.250000000000, 19.587311495393)
    pts(1) = Array(-20.100000000000, -1.250000000000, 19.587311495393)
    pts(2) = Array(-20.100000000000, 1.250000000000, 19.587311495393)
    pts(3) = Array(-18.800000000000, 1.250000000000, 19.587311495393)
    FSData_455 = pts
End Function

Private Function FSData_456() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-18.800000000000, -1.250000000000, 20.437311495393)
    pts(1) = Array(-20.100000000000, -1.250000000000, 20.437311495393)
    pts(2) = Array(-20.100000000000, 1.250000000000, 20.437311495393)
    pts(3) = Array(-18.800000000000, 1.250000000000, 20.437311495393)
    FSData_456 = pts
End Function

Private Function FSData_457() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(18.800000000000, -1.250000000000, 19.587311495393)
    pts(1) = Array(20.100000000000, -1.250000000000, 19.587311495393)
    pts(2) = Array(20.100000000000, 1.250000000000, 19.587311495393)
    pts(3) = Array(18.800000000000, 1.250000000000, 19.587311495393)
    FSData_457 = pts
End Function

Private Function FSData_458() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(18.800000000000, -1.250000000000, 20.437311495393)
    pts(1) = Array(20.100000000000, -1.250000000000, 20.437311495393)
    pts(2) = Array(20.100000000000, 1.250000000000, 20.437311495393)
    pts(3) = Array(18.800000000000, 1.250000000000, 20.437311495393)
    FSData_458 = pts
End Function

Private Function FSData_459() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.200000000000, -1.250000000000, 64.680320557798)
    pts(1) = Array(-14.500000000000, -1.250000000000, 64.680320557798)
    pts(2) = Array(-14.500000000000, 1.250000000000, 64.680320557798)
    pts(3) = Array(-13.200000000000, 1.250000000000, 64.680320557798)
    FSData_459 = pts
End Function

Private Function FSData_460() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-13.200000000000, -1.250000000000, 65.530320557798)
    pts(1) = Array(-14.500000000000, -1.250000000000, 65.530320557798)
    pts(2) = Array(-14.500000000000, 1.250000000000, 65.530320557798)
    pts(3) = Array(-13.200000000000, 1.250000000000, 65.530320557798)
    FSData_460 = pts
End Function

Private Function FSData_461() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.200000000000, -1.250000000000, 64.680320557798)
    pts(1) = Array(14.500000000000, -1.250000000000, 64.680320557798)
    pts(2) = Array(14.500000000000, 1.250000000000, 64.680320557798)
    pts(3) = Array(13.200000000000, 1.250000000000, 64.680320557798)
    FSData_461 = pts
End Function

Private Function FSData_462() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(13.200000000000, -1.250000000000, 65.530320557798)
    pts(1) = Array(14.500000000000, -1.250000000000, 65.530320557798)
    pts(2) = Array(14.500000000000, 1.250000000000, 65.530320557798)
    pts(3) = Array(13.200000000000, 1.250000000000, 65.530320557798)
    FSData_462 = pts
End Function

Private Function FSData_463() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.900000000000, -1.250000000000, 37.600000000000)
    pts(1) = Array(-31.200000000000, -1.250000000000, 37.600000000000)
    pts(2) = Array(-31.200000000000, 1.250000000000, 37.600000000000)
    pts(3) = Array(-29.900000000000, 1.250000000000, 37.600000000000)
    FSData_463 = pts
End Function

Private Function FSData_464() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-29.900000000000, -1.250000000000, 38.450000000000)
    pts(1) = Array(-31.200000000000, -1.250000000000, 38.450000000000)
    pts(2) = Array(-31.200000000000, 1.250000000000, 38.450000000000)
    pts(3) = Array(-29.900000000000, 1.250000000000, 38.450000000000)
    FSData_464 = pts
End Function

Private Function FSData_465() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.900000000000, -1.250000000000, 37.600000000000)
    pts(1) = Array(31.200000000000, -1.250000000000, 37.600000000000)
    pts(2) = Array(31.200000000000, 1.250000000000, 37.600000000000)
    pts(3) = Array(29.900000000000, 1.250000000000, 37.600000000000)
    FSData_465 = pts
End Function

Private Function FSData_466() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(29.900000000000, -1.250000000000, 38.450000000000)
    pts(1) = Array(31.200000000000, -1.250000000000, 38.450000000000)
    pts(2) = Array(31.200000000000, 1.250000000000, 38.450000000000)
    pts(3) = Array(29.900000000000, 1.250000000000, 38.450000000000)
    FSData_466 = pts
End Function

Private Function FSData_467() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-5.149596556417, -1.625316986357, 15.150000000000)
    pts(1) = Array(-4.790587308048, -2.492041982786, 15.150000000000)
    pts(2) = Array(-4.990579632167, -2.595137748753, 15.150000000000)
    pts(3) = Array(-5.190571855178, -2.698233462143, 15.150000000000)
    pts(4) = Array(-5.387877952086, -2.806491791084, 15.150000000000)
    pts(5) = Array(-5.562495864990, -2.957809958732, 15.150000000000)
    pts(6) = Array(-5.721403640235, -3.137063178437, 15.150000000000)
    pts(7) = Array(-5.867010279853, -3.337767274107, 15.150000000000)
    pts(8) = Array(-5.999845590694, -3.557032174135, 15.150000000000)
    pts(9) = Array(-6.119878675531, -3.793031109387, 15.150000000000)
    pts(10) = Array(-6.629868565159, -4.109117035169, 15.150000000000)
    pts(11) = Array(-7.593609541057, -1.782440500542, 15.150000000000)
    pts(12) = Array(-7.009485730206, -1.645329692808, 15.150000000000)
    pts(13) = Array(-6.757733074481, -1.727329932022, 15.150000000000)
    pts(14) = Array(-6.508760627719, -1.788444880632, 15.150000000000)
    pts(15) = Array(-6.263881958399, -1.827404665432, 15.150000000000)
    pts(16) = Array(-6.024766025737, -1.841791067176, 15.150000000000)
    pts(17) = Array(-5.794294412946, -1.825315659305, 15.150000000000)
    pts(18) = Array(-5.578227735339, -1.762349378727, 15.150000000000)
    pts(19) = Array(-5.363912200213, -1.693833199700, 15.150000000000)
    FSData_467 = pts
End Function

Private Function FSData_468() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-5.149596556417, -1.625316986357, 18.250000000000)
    pts(1) = Array(-4.790587308048, -2.492041982786, 18.250000000000)
    pts(2) = Array(-4.990579632167, -2.595137748753, 18.250000000000)
    pts(3) = Array(-5.190571855178, -2.698233462143, 18.250000000000)
    pts(4) = Array(-5.387877952086, -2.806491791084, 18.250000000000)
    pts(5) = Array(-5.562495864990, -2.957809958732, 18.250000000000)
    pts(6) = Array(-5.721403640235, -3.137063178437, 18.250000000000)
    pts(7) = Array(-5.867010279853, -3.337767274107, 18.250000000000)
    pts(8) = Array(-5.999845590694, -3.557032174135, 18.250000000000)
    pts(9) = Array(-6.119878675531, -3.793031109387, 18.250000000000)
    pts(10) = Array(-6.629868565159, -4.109117035169, 18.250000000000)
    pts(11) = Array(-7.593609541057, -1.782440500542, 18.250000000000)
    pts(12) = Array(-7.009485730206, -1.645329692808, 18.250000000000)
    pts(13) = Array(-6.757733074481, -1.727329932022, 18.250000000000)
    pts(14) = Array(-6.508760627719, -1.788444880632, 18.250000000000)
    pts(15) = Array(-6.263881958399, -1.827404665432, 18.250000000000)
    pts(16) = Array(-6.024766025737, -1.841791067176, 18.250000000000)
    pts(17) = Array(-5.794294412946, -1.825315659305, 18.250000000000)
    pts(18) = Array(-5.578227735339, -1.762349378727, 18.250000000000)
    pts(19) = Array(-5.363912200213, -1.693833199700, 18.250000000000)
    FSData_468 = pts
End Function

Private Function FSData_469() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-4.135624976143, -3.472262383043, 15.150000000000)
    pts(1) = Array(-3.472262383043, -4.135624976143, 15.150000000000)
    pts(2) = Array(-3.617578156373, -4.307406793249, 15.150000000000)
    pts(3) = Array(-3.762893856413, -4.479188523089, 15.150000000000)
    pts(4) = Array(-3.903752252084, -4.654711951812, 15.150000000000)
    pts(5) = Array(-4.007171212051, -4.861335092061, 15.150000000000)
    pts(6) = Array(-4.085385615778, -5.087754845744, 15.150000000000)
    pts(7) = Array(-4.143102477698, -5.328902500449, 15.150000000000)
    pts(8) = Array(-4.181917258039, -5.582310726475, 15.150000000000)
    pts(9) = Array(-4.202500485768, -5.846279985351, 15.150000000000)
    pts(10) = Array(-4.552708859582, -6.333469984131, 15.150000000000)
    pts(11) = Array(-6.333469984131, -4.552708859582, 15.150000000000)
    pts(12) = Array(-5.846279985351, -4.202500485768, 15.150000000000)
    pts(13) = Array(-5.582310726475, -4.181917258039, 15.150000000000)
    pts(14) = Array(-5.328902500449, -4.143102477698, 15.150000000000)
    pts(15) = Array(-5.087754845744, -4.085385615778, 15.150000000000)
    pts(16) = Array(-4.861335092061, -4.007171212051, 15.150000000000)
    pts(17) = Array(-4.654711951812, -3.903752252084, 15.150000000000)
    pts(18) = Array(-4.479188523089, -3.762893856413, 15.150000000000)
    pts(19) = Array(-4.307406793249, -3.617578156373, 15.150000000000)
    FSData_469 = pts
End Function

Private Function FSData_470() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-4.135624976143, -3.472262383043, 18.250000000000)
    pts(1) = Array(-3.472262383043, -4.135624976143, 18.250000000000)
    pts(2) = Array(-3.617578156373, -4.307406793249, 18.250000000000)
    pts(3) = Array(-3.762893856413, -4.479188523089, 18.250000000000)
    pts(4) = Array(-3.903752252084, -4.654711951812, 18.250000000000)
    pts(5) = Array(-4.007171212051, -4.861335092061, 18.250000000000)
    pts(6) = Array(-4.085385615778, -5.087754845744, 18.250000000000)
    pts(7) = Array(-4.143102477698, -5.328902500449, 18.250000000000)
    pts(8) = Array(-4.181917258039, -5.582310726475, 18.250000000000)
    pts(9) = Array(-4.202500485768, -5.846279985351, 18.250000000000)
    pts(10) = Array(-4.552708859582, -6.333469984131, 18.250000000000)
    pts(11) = Array(-6.333469984131, -4.552708859582, 18.250000000000)
    pts(12) = Array(-5.846279985351, -4.202500485768, 18.250000000000)
    pts(13) = Array(-5.582310726475, -4.181917258039, 18.250000000000)
    pts(14) = Array(-5.328902500449, -4.143102477698, 18.250000000000)
    pts(15) = Array(-5.087754845744, -4.085385615778, 18.250000000000)
    pts(16) = Array(-4.861335092061, -4.007171212051, 18.250000000000)
    pts(17) = Array(-4.654711951812, -3.903752252084, 18.250000000000)
    pts(18) = Array(-4.479188523089, -3.762893856413, 18.250000000000)
    pts(19) = Array(-4.307406793249, -3.617578156373, 18.250000000000)
    FSData_470 = pts
End Function

Private Function FSData_471() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-2.492041982786, -4.790587308048, 15.150000000000)
    pts(1) = Array(-1.625316986357, -5.149596556417, 15.150000000000)
    pts(2) = Array(-1.693833199700, -5.363912200213, 15.150000000000)
    pts(3) = Array(-1.762349378727, -5.578227735339, 15.150000000000)
    pts(4) = Array(-1.825315659305, -5.794294412946, 15.150000000000)
    pts(5) = Array(-1.841791067176, -6.024766025737, 15.150000000000)
    pts(6) = Array(-1.827404665432, -6.263881958399, 15.150000000000)
    pts(7) = Array(-1.788444880632, -6.508760627719, 15.150000000000)
    pts(8) = Array(-1.727329932022, -6.757733074481, 15.150000000000)
    pts(9) = Array(-1.645329692808, -7.009485730206, 15.150000000000)
    pts(10) = Array(-1.782440500542, -7.593609541057, 15.150000000000)
    pts(11) = Array(-4.109117035169, -6.629868565159, 15.150000000000)
    pts(12) = Array(-3.793031109387, -6.119878675531, 15.150000000000)
    pts(13) = Array(-3.557032174135, -5.999845590694, 15.150000000000)
    pts(14) = Array(-3.337767274107, -5.867010279853, 15.150000000000)
    pts(15) = Array(-3.137063178437, -5.721403640235, 15.150000000000)
    pts(16) = Array(-2.957809958732, -5.562495864990, 15.150000000000)
    pts(17) = Array(-2.806491791084, -5.387877952086, 15.150000000000)
    pts(18) = Array(-2.698233462143, -5.190571855178, 15.150000000000)
    pts(19) = Array(-2.595137748753, -4.990579632167, 15.150000000000)
    FSData_471 = pts
End Function

Private Function FSData_472() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-2.492041982786, -4.790587308048, 18.250000000000)
    pts(1) = Array(-1.625316986357, -5.149596556417, 18.250000000000)
    pts(2) = Array(-1.693833199700, -5.363912200213, 18.250000000000)
    pts(3) = Array(-1.762349378727, -5.578227735339, 18.250000000000)
    pts(4) = Array(-1.825315659305, -5.794294412946, 18.250000000000)
    pts(5) = Array(-1.841791067176, -6.024766025737, 18.250000000000)
    pts(6) = Array(-1.827404665432, -6.263881958399, 18.250000000000)
    pts(7) = Array(-1.788444880632, -6.508760627719, 18.250000000000)
    pts(8) = Array(-1.727329932022, -6.757733074481, 18.250000000000)
    pts(9) = Array(-1.645329692808, -7.009485730206, 18.250000000000)
    pts(10) = Array(-1.782440500542, -7.593609541057, 18.250000000000)
    pts(11) = Array(-4.109117035169, -6.629868565159, 18.250000000000)
    pts(12) = Array(-3.793031109387, -6.119878675531, 18.250000000000)
    pts(13) = Array(-3.557032174135, -5.999845590694, 18.250000000000)
    pts(14) = Array(-3.337767274107, -5.867010279853, 18.250000000000)
    pts(15) = Array(-3.137063178437, -5.721403640235, 18.250000000000)
    pts(16) = Array(-2.957809958732, -5.562495864990, 18.250000000000)
    pts(17) = Array(-2.806491791084, -5.387877952086, 18.250000000000)
    pts(18) = Array(-2.698233462143, -5.190571855178, 18.250000000000)
    pts(19) = Array(-2.595137748753, -4.990579632167, 18.250000000000)
    FSData_472 = pts
End Function

Private Function FSData_473() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-0.469068187966, -5.379588742184, 15.150000000000)
    pts(1) = Array(0.469068187966, -5.379588742184, 15.150000000000)
    pts(2) = Array(0.487782506992, -5.603810598680, 15.150000000000)
    pts(3) = Array(0.506496816134, -5.828032341644, 15.150000000000)
    pts(4) = Array(0.531008696075, -6.051748075118, 15.150000000000)
    pts(5) = Array(0.603985071800, -6.270980946634, 15.150000000000)
    pts(6) = Array(0.708782079762, -6.486389825119, 15.150000000000)
    pts(7) = Array(0.838487237218, -6.697718951481, 15.150000000000)
    pts(8) = Array(0.990227717860, -6.904351820901, 15.150000000000)
    pts(9) = Array(1.162327630931, -7.105560813783, 15.150000000000)
    pts(10) = Array(1.259188266842, -7.697690881599, 15.150000000000)
    pts(11) = Array(-1.259188266842, -7.697690881599, 15.150000000000)
    pts(12) = Array(-1.162327630931, -7.105560813783, 15.150000000000)
    pts(13) = Array(-0.990227717860, -6.904351820901, 15.150000000000)
    pts(14) = Array(-0.838487237218, -6.697718951481, 15.150000000000)
    pts(15) = Array(-0.708782079762, -6.486389825119, 15.150000000000)
    pts(16) = Array(-0.603985071800, -6.270980946634, 15.150000000000)
    pts(17) = Array(-0.531008696075, -6.051748075118, 15.150000000000)
    pts(18) = Array(-0.506496816134, -5.828032341644, 15.150000000000)
    pts(19) = Array(-0.487782506992, -5.603810598680, 15.150000000000)
    FSData_473 = pts
End Function

Private Function FSData_474() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-0.469068187966, -5.379588742184, 18.250000000000)
    pts(1) = Array(0.469068187966, -5.379588742184, 18.250000000000)
    pts(2) = Array(0.487782506992, -5.603810598680, 18.250000000000)
    pts(3) = Array(0.506496816134, -5.828032341644, 18.250000000000)
    pts(4) = Array(0.531008696075, -6.051748075118, 18.250000000000)
    pts(5) = Array(0.603985071800, -6.270980946634, 18.250000000000)
    pts(6) = Array(0.708782079762, -6.486389825119, 18.250000000000)
    pts(7) = Array(0.838487237218, -6.697718951481, 18.250000000000)
    pts(8) = Array(0.990227717860, -6.904351820901, 18.250000000000)
    pts(9) = Array(1.162327630931, -7.105560813783, 18.250000000000)
    pts(10) = Array(1.259188266842, -7.697690881599, 18.250000000000)
    pts(11) = Array(-1.259188266842, -7.697690881599, 18.250000000000)
    pts(12) = Array(-1.162327630931, -7.105560813783, 18.250000000000)
    pts(13) = Array(-0.990227717860, -6.904351820901, 18.250000000000)
    pts(14) = Array(-0.838487237218, -6.697718951481, 18.250000000000)
    pts(15) = Array(-0.708782079762, -6.486389825119, 18.250000000000)
    pts(16) = Array(-0.603985071800, -6.270980946634, 18.250000000000)
    pts(17) = Array(-0.531008696075, -6.051748075118, 18.250000000000)
    pts(18) = Array(-0.506496816134, -5.828032341644, 18.250000000000)
    pts(19) = Array(-0.487782506992, -5.603810598680, 18.250000000000)
    FSData_474 = pts
End Function

Private Function FSData_475() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(1.625316986357, -5.149596556417, 15.150000000000)
    pts(1) = Array(2.492041982786, -4.790587308048, 15.150000000000)
    pts(2) = Array(2.595137748753, -4.990579632167, 15.150000000000)
    pts(3) = Array(2.698233462143, -5.190571855178, 15.150000000000)
    pts(4) = Array(2.806491791084, -5.387877952086, 15.150000000000)
    pts(5) = Array(2.957809958732, -5.562495864990, 15.150000000000)
    pts(6) = Array(3.137063178437, -5.721403640235, 15.150000000000)
    pts(7) = Array(3.337767274107, -5.867010279853, 15.150000000000)
    pts(8) = Array(3.557032174135, -5.999845590694, 15.150000000000)
    pts(9) = Array(3.793031109387, -6.119878675531, 15.150000000000)
    pts(10) = Array(4.109117035169, -6.629868565159, 15.150000000000)
    pts(11) = Array(1.782440500542, -7.593609541057, 15.150000000000)
    pts(12) = Array(1.645329692808, -7.009485730206, 15.150000000000)
    pts(13) = Array(1.727329932022, -6.757733074481, 15.150000000000)
    pts(14) = Array(1.788444880632, -6.508760627719, 15.150000000000)
    pts(15) = Array(1.827404665432, -6.263881958399, 15.150000000000)
    pts(16) = Array(1.841791067176, -6.024766025737, 15.150000000000)
    pts(17) = Array(1.825315659305, -5.794294412946, 15.150000000000)
    pts(18) = Array(1.762349378727, -5.578227735339, 15.150000000000)
    pts(19) = Array(1.693833199700, -5.363912200213, 15.150000000000)
    FSData_475 = pts
End Function

Private Function FSData_476() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(1.625316986357, -5.149596556417, 18.250000000000)
    pts(1) = Array(2.492041982786, -4.790587308048, 18.250000000000)
    pts(2) = Array(2.595137748753, -4.990579632167, 18.250000000000)
    pts(3) = Array(2.698233462143, -5.190571855178, 18.250000000000)
    pts(4) = Array(2.806491791084, -5.387877952086, 18.250000000000)
    pts(5) = Array(2.957809958732, -5.562495864990, 18.250000000000)
    pts(6) = Array(3.137063178437, -5.721403640235, 18.250000000000)
    pts(7) = Array(3.337767274107, -5.867010279853, 18.250000000000)
    pts(8) = Array(3.557032174135, -5.999845590694, 18.250000000000)
    pts(9) = Array(3.793031109387, -6.119878675531, 18.250000000000)
    pts(10) = Array(4.109117035169, -6.629868565159, 18.250000000000)
    pts(11) = Array(1.782440500542, -7.593609541057, 18.250000000000)
    pts(12) = Array(1.645329692808, -7.009485730206, 18.250000000000)
    pts(13) = Array(1.727329932022, -6.757733074481, 18.250000000000)
    pts(14) = Array(1.788444880632, -6.508760627719, 18.250000000000)
    pts(15) = Array(1.827404665432, -6.263881958399, 18.250000000000)
    pts(16) = Array(1.841791067176, -6.024766025737, 18.250000000000)
    pts(17) = Array(1.825315659305, -5.794294412946, 18.250000000000)
    pts(18) = Array(1.762349378727, -5.578227735339, 18.250000000000)
    pts(19) = Array(1.693833199700, -5.363912200213, 18.250000000000)
    FSData_476 = pts
End Function

Private Function FSData_477() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(3.472262383043, -4.135624976143, 15.150000000000)
    pts(1) = Array(4.135624976143, -3.472262383043, 15.150000000000)
    pts(2) = Array(4.307406793249, -3.617578156373, 15.150000000000)
    pts(3) = Array(4.479188523089, -3.762893856413, 15.150000000000)
    pts(4) = Array(4.654711951812, -3.903752252084, 15.150000000000)
    pts(5) = Array(4.861335092061, -4.007171212051, 15.150000000000)
    pts(6) = Array(5.087754845744, -4.085385615778, 15.150000000000)
    pts(7) = Array(5.328902500449, -4.143102477698, 15.150000000000)
    pts(8) = Array(5.582310726475, -4.181917258039, 15.150000000000)
    pts(9) = Array(5.846279985351, -4.202500485768, 15.150000000000)
    pts(10) = Array(6.333469984131, -4.552708859582, 15.150000000000)
    pts(11) = Array(4.552708859582, -6.333469984131, 15.150000000000)
    pts(12) = Array(4.202500485768, -5.846279985351, 15.150000000000)
    pts(13) = Array(4.181917258039, -5.582310726475, 15.150000000000)
    pts(14) = Array(4.143102477698, -5.328902500449, 15.150000000000)
    pts(15) = Array(4.085385615778, -5.087754845744, 15.150000000000)
    pts(16) = Array(4.007171212051, -4.861335092061, 15.150000000000)
    pts(17) = Array(3.903752252084, -4.654711951812, 15.150000000000)
    pts(18) = Array(3.762893856413, -4.479188523089, 15.150000000000)
    pts(19) = Array(3.617578156373, -4.307406793249, 15.150000000000)
    FSData_477 = pts
End Function

Private Function FSData_478() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(3.472262383043, -4.135624976143, 18.250000000000)
    pts(1) = Array(4.135624976143, -3.472262383043, 18.250000000000)
    pts(2) = Array(4.307406793249, -3.617578156373, 18.250000000000)
    pts(3) = Array(4.479188523089, -3.762893856413, 18.250000000000)
    pts(4) = Array(4.654711951812, -3.903752252084, 18.250000000000)
    pts(5) = Array(4.861335092061, -4.007171212051, 18.250000000000)
    pts(6) = Array(5.087754845744, -4.085385615778, 18.250000000000)
    pts(7) = Array(5.328902500449, -4.143102477698, 18.250000000000)
    pts(8) = Array(5.582310726475, -4.181917258039, 18.250000000000)
    pts(9) = Array(5.846279985351, -4.202500485768, 18.250000000000)
    pts(10) = Array(6.333469984131, -4.552708859582, 18.250000000000)
    pts(11) = Array(4.552708859582, -6.333469984131, 18.250000000000)
    pts(12) = Array(4.202500485768, -5.846279985351, 18.250000000000)
    pts(13) = Array(4.181917258039, -5.582310726475, 18.250000000000)
    pts(14) = Array(4.143102477698, -5.328902500449, 18.250000000000)
    pts(15) = Array(4.085385615778, -5.087754845744, 18.250000000000)
    pts(16) = Array(4.007171212051, -4.861335092061, 18.250000000000)
    pts(17) = Array(3.903752252084, -4.654711951812, 18.250000000000)
    pts(18) = Array(3.762893856413, -4.479188523089, 18.250000000000)
    pts(19) = Array(3.617578156373, -4.307406793249, 18.250000000000)
    FSData_478 = pts
End Function

Private Function FSData_479() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(4.790587308048, -2.492041982786, 15.150000000000)
    pts(1) = Array(5.149596556417, -1.625316986357, 15.150000000000)
    pts(2) = Array(5.363912200213, -1.693833199700, 15.150000000000)
    pts(3) = Array(5.578227735339, -1.762349378727, 15.150000000000)
    pts(4) = Array(5.794294412946, -1.825315659305, 15.150000000000)
    pts(5) = Array(6.024766025737, -1.841791067176, 15.150000000000)
    pts(6) = Array(6.263881958399, -1.827404665432, 15.150000000000)
    pts(7) = Array(6.508760627719, -1.788444880632, 15.150000000000)
    pts(8) = Array(6.757733074481, -1.727329932022, 15.150000000000)
    pts(9) = Array(7.009485730206, -1.645329692808, 15.150000000000)
    pts(10) = Array(7.593609541057, -1.782440500542, 15.150000000000)
    pts(11) = Array(6.629868565159, -4.109117035169, 15.150000000000)
    pts(12) = Array(6.119878675531, -3.793031109387, 15.150000000000)
    pts(13) = Array(5.999845590694, -3.557032174135, 15.150000000000)
    pts(14) = Array(5.867010279853, -3.337767274107, 15.150000000000)
    pts(15) = Array(5.721403640235, -3.137063178437, 15.150000000000)
    pts(16) = Array(5.562495864990, -2.957809958732, 15.150000000000)
    pts(17) = Array(5.387877952086, -2.806491791084, 15.150000000000)
    pts(18) = Array(5.190571855178, -2.698233462143, 15.150000000000)
    pts(19) = Array(4.990579632167, -2.595137748753, 15.150000000000)
    FSData_479 = pts
End Function

Private Function FSData_480() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(4.790587308048, -2.492041982786, 18.250000000000)
    pts(1) = Array(5.149596556417, -1.625316986357, 18.250000000000)
    pts(2) = Array(5.363912200213, -1.693833199700, 18.250000000000)
    pts(3) = Array(5.578227735339, -1.762349378727, 18.250000000000)
    pts(4) = Array(5.794294412946, -1.825315659305, 18.250000000000)
    pts(5) = Array(6.024766025737, -1.841791067176, 18.250000000000)
    pts(6) = Array(6.263881958399, -1.827404665432, 18.250000000000)
    pts(7) = Array(6.508760627719, -1.788444880632, 18.250000000000)
    pts(8) = Array(6.757733074481, -1.727329932022, 18.250000000000)
    pts(9) = Array(7.009485730206, -1.645329692808, 18.250000000000)
    pts(10) = Array(7.593609541057, -1.782440500542, 18.250000000000)
    pts(11) = Array(6.629868565159, -4.109117035169, 18.250000000000)
    pts(12) = Array(6.119878675531, -3.793031109387, 18.250000000000)
    pts(13) = Array(5.999845590694, -3.557032174135, 18.250000000000)
    pts(14) = Array(5.867010279853, -3.337767274107, 18.250000000000)
    pts(15) = Array(5.721403640235, -3.137063178437, 18.250000000000)
    pts(16) = Array(5.562495864990, -2.957809958732, 18.250000000000)
    pts(17) = Array(5.387877952086, -2.806491791084, 18.250000000000)
    pts(18) = Array(5.190571855178, -2.698233462143, 18.250000000000)
    pts(19) = Array(4.990579632167, -2.595137748753, 18.250000000000)
    FSData_480 = pts
End Function

Private Function FSData_481() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(5.379588742184, -0.469068187966, 15.150000000000)
    pts(1) = Array(5.379588742184, 0.469068187966, 15.150000000000)
    pts(2) = Array(5.603810598680, 0.487782506992, 15.150000000000)
    pts(3) = Array(5.828032341644, 0.506496816134, 15.150000000000)
    pts(4) = Array(6.051748075118, 0.531008696075, 15.150000000000)
    pts(5) = Array(6.270980946634, 0.603985071800, 15.150000000000)
    pts(6) = Array(6.486389825119, 0.708782079762, 15.150000000000)
    pts(7) = Array(6.697718951481, 0.838487237218, 15.150000000000)
    pts(8) = Array(6.904351820901, 0.990227717860, 15.150000000000)
    pts(9) = Array(7.105560813783, 1.162327630931, 15.150000000000)
    pts(10) = Array(7.697690881599, 1.259188266842, 15.150000000000)
    pts(11) = Array(7.697690881599, -1.259188266842, 15.150000000000)
    pts(12) = Array(7.105560813783, -1.162327630931, 15.150000000000)
    pts(13) = Array(6.904351820901, -0.990227717860, 15.150000000000)
    pts(14) = Array(6.697718951481, -0.838487237218, 15.150000000000)
    pts(15) = Array(6.486389825119, -0.708782079762, 15.150000000000)
    pts(16) = Array(6.270980946634, -0.603985071800, 15.150000000000)
    pts(17) = Array(6.051748075118, -0.531008696075, 15.150000000000)
    pts(18) = Array(5.828032341644, -0.506496816134, 15.150000000000)
    pts(19) = Array(5.603810598680, -0.487782506992, 15.150000000000)
    FSData_481 = pts
End Function

Private Function FSData_482() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(5.379588742184, -0.469068187966, 18.250000000000)
    pts(1) = Array(5.379588742184, 0.469068187966, 18.250000000000)
    pts(2) = Array(5.603810598680, 0.487782506992, 18.250000000000)
    pts(3) = Array(5.828032341644, 0.506496816134, 18.250000000000)
    pts(4) = Array(6.051748075118, 0.531008696075, 18.250000000000)
    pts(5) = Array(6.270980946634, 0.603985071800, 18.250000000000)
    pts(6) = Array(6.486389825119, 0.708782079762, 18.250000000000)
    pts(7) = Array(6.697718951481, 0.838487237218, 18.250000000000)
    pts(8) = Array(6.904351820901, 0.990227717860, 18.250000000000)
    pts(9) = Array(7.105560813783, 1.162327630931, 18.250000000000)
    pts(10) = Array(7.697690881599, 1.259188266842, 18.250000000000)
    pts(11) = Array(7.697690881599, -1.259188266842, 18.250000000000)
    pts(12) = Array(7.105560813783, -1.162327630931, 18.250000000000)
    pts(13) = Array(6.904351820901, -0.990227717860, 18.250000000000)
    pts(14) = Array(6.697718951481, -0.838487237218, 18.250000000000)
    pts(15) = Array(6.486389825119, -0.708782079762, 18.250000000000)
    pts(16) = Array(6.270980946634, -0.603985071800, 18.250000000000)
    pts(17) = Array(6.051748075118, -0.531008696075, 18.250000000000)
    pts(18) = Array(5.828032341644, -0.506496816134, 18.250000000000)
    pts(19) = Array(5.603810598680, -0.487782506992, 18.250000000000)
    FSData_482 = pts
End Function

Private Function FSData_483() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(5.149596556417, 1.625316986357, 15.150000000000)
    pts(1) = Array(4.790587308048, 2.492041982786, 15.150000000000)
    pts(2) = Array(4.990579632167, 2.595137748753, 15.150000000000)
    pts(3) = Array(5.190571855178, 2.698233462143, 15.150000000000)
    pts(4) = Array(5.387877952086, 2.806491791084, 15.150000000000)
    pts(5) = Array(5.562495864990, 2.957809958732, 15.150000000000)
    pts(6) = Array(5.721403640235, 3.137063178437, 15.150000000000)
    pts(7) = Array(5.867010279853, 3.337767274107, 15.150000000000)
    pts(8) = Array(5.999845590694, 3.557032174135, 15.150000000000)
    pts(9) = Array(6.119878675531, 3.793031109387, 15.150000000000)
    pts(10) = Array(6.629868565159, 4.109117035169, 15.150000000000)
    pts(11) = Array(7.593609541057, 1.782440500542, 15.150000000000)
    pts(12) = Array(7.009485730206, 1.645329692808, 15.150000000000)
    pts(13) = Array(6.757733074481, 1.727329932022, 15.150000000000)
    pts(14) = Array(6.508760627719, 1.788444880632, 15.150000000000)
    pts(15) = Array(6.263881958399, 1.827404665432, 15.150000000000)
    pts(16) = Array(6.024766025737, 1.841791067176, 15.150000000000)
    pts(17) = Array(5.794294412946, 1.825315659305, 15.150000000000)
    pts(18) = Array(5.578227735339, 1.762349378727, 15.150000000000)
    pts(19) = Array(5.363912200213, 1.693833199700, 15.150000000000)
    FSData_483 = pts
End Function

Private Function FSData_484() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(5.149596556417, 1.625316986357, 18.250000000000)
    pts(1) = Array(4.790587308048, 2.492041982786, 18.250000000000)
    pts(2) = Array(4.990579632167, 2.595137748753, 18.250000000000)
    pts(3) = Array(5.190571855178, 2.698233462143, 18.250000000000)
    pts(4) = Array(5.387877952086, 2.806491791084, 18.250000000000)
    pts(5) = Array(5.562495864990, 2.957809958732, 18.250000000000)
    pts(6) = Array(5.721403640235, 3.137063178437, 18.250000000000)
    pts(7) = Array(5.867010279853, 3.337767274107, 18.250000000000)
    pts(8) = Array(5.999845590694, 3.557032174135, 18.250000000000)
    pts(9) = Array(6.119878675531, 3.793031109387, 18.250000000000)
    pts(10) = Array(6.629868565159, 4.109117035169, 18.250000000000)
    pts(11) = Array(7.593609541057, 1.782440500542, 18.250000000000)
    pts(12) = Array(7.009485730206, 1.645329692808, 18.250000000000)
    pts(13) = Array(6.757733074481, 1.727329932022, 18.250000000000)
    pts(14) = Array(6.508760627719, 1.788444880632, 18.250000000000)
    pts(15) = Array(6.263881958399, 1.827404665432, 18.250000000000)
    pts(16) = Array(6.024766025737, 1.841791067176, 18.250000000000)
    pts(17) = Array(5.794294412946, 1.825315659305, 18.250000000000)
    pts(18) = Array(5.578227735339, 1.762349378727, 18.250000000000)
    pts(19) = Array(5.363912200213, 1.693833199700, 18.250000000000)
    FSData_484 = pts
End Function

Private Function FSData_485() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(4.135624976143, 3.472262383043, 15.150000000000)
    pts(1) = Array(3.472262383043, 4.135624976143, 15.150000000000)
    pts(2) = Array(3.617578156373, 4.307406793249, 15.150000000000)
    pts(3) = Array(3.762893856413, 4.479188523089, 15.150000000000)
    pts(4) = Array(3.903752252084, 4.654711951812, 15.150000000000)
    pts(5) = Array(4.007171212051, 4.861335092061, 15.150000000000)
    pts(6) = Array(4.085385615778, 5.087754845744, 15.150000000000)
    pts(7) = Array(4.143102477698, 5.328902500449, 15.150000000000)
    pts(8) = Array(4.181917258039, 5.582310726475, 15.150000000000)
    pts(9) = Array(4.202500485768, 5.846279985351, 15.150000000000)
    pts(10) = Array(4.552708859582, 6.333469984131, 15.150000000000)
    pts(11) = Array(6.333469984131, 4.552708859582, 15.150000000000)
    pts(12) = Array(5.846279985351, 4.202500485768, 15.150000000000)
    pts(13) = Array(5.582310726475, 4.181917258039, 15.150000000000)
    pts(14) = Array(5.328902500449, 4.143102477698, 15.150000000000)
    pts(15) = Array(5.087754845744, 4.085385615778, 15.150000000000)
    pts(16) = Array(4.861335092061, 4.007171212051, 15.150000000000)
    pts(17) = Array(4.654711951812, 3.903752252084, 15.150000000000)
    pts(18) = Array(4.479188523089, 3.762893856413, 15.150000000000)
    pts(19) = Array(4.307406793249, 3.617578156373, 15.150000000000)
    FSData_485 = pts
End Function

Private Function FSData_486() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(4.135624976143, 3.472262383043, 18.250000000000)
    pts(1) = Array(3.472262383043, 4.135624976143, 18.250000000000)
    pts(2) = Array(3.617578156373, 4.307406793249, 18.250000000000)
    pts(3) = Array(3.762893856413, 4.479188523089, 18.250000000000)
    pts(4) = Array(3.903752252084, 4.654711951812, 18.250000000000)
    pts(5) = Array(4.007171212051, 4.861335092061, 18.250000000000)
    pts(6) = Array(4.085385615778, 5.087754845744, 18.250000000000)
    pts(7) = Array(4.143102477698, 5.328902500449, 18.250000000000)
    pts(8) = Array(4.181917258039, 5.582310726475, 18.250000000000)
    pts(9) = Array(4.202500485768, 5.846279985351, 18.250000000000)
    pts(10) = Array(4.552708859582, 6.333469984131, 18.250000000000)
    pts(11) = Array(6.333469984131, 4.552708859582, 18.250000000000)
    pts(12) = Array(5.846279985351, 4.202500485768, 18.250000000000)
    pts(13) = Array(5.582310726475, 4.181917258039, 18.250000000000)
    pts(14) = Array(5.328902500449, 4.143102477698, 18.250000000000)
    pts(15) = Array(5.087754845744, 4.085385615778, 18.250000000000)
    pts(16) = Array(4.861335092061, 4.007171212051, 18.250000000000)
    pts(17) = Array(4.654711951812, 3.903752252084, 18.250000000000)
    pts(18) = Array(4.479188523089, 3.762893856413, 18.250000000000)
    pts(19) = Array(4.307406793249, 3.617578156373, 18.250000000000)
    FSData_486 = pts
End Function

Private Function FSData_487() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(2.492041982786, 4.790587308048, 15.150000000000)
    pts(1) = Array(1.625316986357, 5.149596556417, 15.150000000000)
    pts(2) = Array(1.693833199700, 5.363912200213, 15.150000000000)
    pts(3) = Array(1.762349378727, 5.578227735339, 15.150000000000)
    pts(4) = Array(1.825315659305, 5.794294412946, 15.150000000000)
    pts(5) = Array(1.841791067176, 6.024766025737, 15.150000000000)
    pts(6) = Array(1.827404665432, 6.263881958399, 15.150000000000)
    pts(7) = Array(1.788444880632, 6.508760627719, 15.150000000000)
    pts(8) = Array(1.727329932022, 6.757733074481, 15.150000000000)
    pts(9) = Array(1.645329692808, 7.009485730206, 15.150000000000)
    pts(10) = Array(1.782440500542, 7.593609541057, 15.150000000000)
    pts(11) = Array(4.109117035169, 6.629868565159, 15.150000000000)
    pts(12) = Array(3.793031109387, 6.119878675531, 15.150000000000)
    pts(13) = Array(3.557032174135, 5.999845590694, 15.150000000000)
    pts(14) = Array(3.337767274107, 5.867010279853, 15.150000000000)
    pts(15) = Array(3.137063178437, 5.721403640235, 15.150000000000)
    pts(16) = Array(2.957809958732, 5.562495864990, 15.150000000000)
    pts(17) = Array(2.806491791084, 5.387877952086, 15.150000000000)
    pts(18) = Array(2.698233462143, 5.190571855178, 15.150000000000)
    pts(19) = Array(2.595137748753, 4.990579632167, 15.150000000000)
    FSData_487 = pts
End Function

Private Function FSData_488() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(2.492041982786, 4.790587308048, 18.250000000000)
    pts(1) = Array(1.625316986357, 5.149596556417, 18.250000000000)
    pts(2) = Array(1.693833199700, 5.363912200213, 18.250000000000)
    pts(3) = Array(1.762349378727, 5.578227735339, 18.250000000000)
    pts(4) = Array(1.825315659305, 5.794294412946, 18.250000000000)
    pts(5) = Array(1.841791067176, 6.024766025737, 18.250000000000)
    pts(6) = Array(1.827404665432, 6.263881958399, 18.250000000000)
    pts(7) = Array(1.788444880632, 6.508760627719, 18.250000000000)
    pts(8) = Array(1.727329932022, 6.757733074481, 18.250000000000)
    pts(9) = Array(1.645329692808, 7.009485730206, 18.250000000000)
    pts(10) = Array(1.782440500542, 7.593609541057, 18.250000000000)
    pts(11) = Array(4.109117035169, 6.629868565159, 18.250000000000)
    pts(12) = Array(3.793031109387, 6.119878675531, 18.250000000000)
    pts(13) = Array(3.557032174135, 5.999845590694, 18.250000000000)
    pts(14) = Array(3.337767274107, 5.867010279853, 18.250000000000)
    pts(15) = Array(3.137063178437, 5.721403640235, 18.250000000000)
    pts(16) = Array(2.957809958732, 5.562495864990, 18.250000000000)
    pts(17) = Array(2.806491791084, 5.387877952086, 18.250000000000)
    pts(18) = Array(2.698233462143, 5.190571855178, 18.250000000000)
    pts(19) = Array(2.595137748753, 4.990579632167, 18.250000000000)
    FSData_488 = pts
End Function

Private Function FSData_489() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(0.469068187966, 5.379588742184, 15.150000000000)
    pts(1) = Array(-0.469068187966, 5.379588742184, 15.150000000000)
    pts(2) = Array(-0.487782506992, 5.603810598680, 15.150000000000)
    pts(3) = Array(-0.506496816134, 5.828032341644, 15.150000000000)
    pts(4) = Array(-0.531008696075, 6.051748075118, 15.150000000000)
    pts(5) = Array(-0.603985071800, 6.270980946634, 15.150000000000)
    pts(6) = Array(-0.708782079762, 6.486389825119, 15.150000000000)
    pts(7) = Array(-0.838487237218, 6.697718951481, 15.150000000000)
    pts(8) = Array(-0.990227717860, 6.904351820901, 15.150000000000)
    pts(9) = Array(-1.162327630931, 7.105560813783, 15.150000000000)
    pts(10) = Array(-1.259188266842, 7.697690881599, 15.150000000000)
    pts(11) = Array(1.259188266842, 7.697690881599, 15.150000000000)
    pts(12) = Array(1.162327630931, 7.105560813783, 15.150000000000)
    pts(13) = Array(0.990227717860, 6.904351820901, 15.150000000000)
    pts(14) = Array(0.838487237218, 6.697718951481, 15.150000000000)
    pts(15) = Array(0.708782079762, 6.486389825119, 15.150000000000)
    pts(16) = Array(0.603985071800, 6.270980946634, 15.150000000000)
    pts(17) = Array(0.531008696075, 6.051748075118, 15.150000000000)
    pts(18) = Array(0.506496816134, 5.828032341644, 15.150000000000)
    pts(19) = Array(0.487782506992, 5.603810598680, 15.150000000000)
    FSData_489 = pts
End Function

Private Function FSData_490() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(0.469068187966, 5.379588742184, 18.250000000000)
    pts(1) = Array(-0.469068187966, 5.379588742184, 18.250000000000)
    pts(2) = Array(-0.487782506992, 5.603810598680, 18.250000000000)
    pts(3) = Array(-0.506496816134, 5.828032341644, 18.250000000000)
    pts(4) = Array(-0.531008696075, 6.051748075118, 18.250000000000)
    pts(5) = Array(-0.603985071800, 6.270980946634, 18.250000000000)
    pts(6) = Array(-0.708782079762, 6.486389825119, 18.250000000000)
    pts(7) = Array(-0.838487237218, 6.697718951481, 18.250000000000)
    pts(8) = Array(-0.990227717860, 6.904351820901, 18.250000000000)
    pts(9) = Array(-1.162327630931, 7.105560813783, 18.250000000000)
    pts(10) = Array(-1.259188266842, 7.697690881599, 18.250000000000)
    pts(11) = Array(1.259188266842, 7.697690881599, 18.250000000000)
    pts(12) = Array(1.162327630931, 7.105560813783, 18.250000000000)
    pts(13) = Array(0.990227717860, 6.904351820901, 18.250000000000)
    pts(14) = Array(0.838487237218, 6.697718951481, 18.250000000000)
    pts(15) = Array(0.708782079762, 6.486389825119, 18.250000000000)
    pts(16) = Array(0.603985071800, 6.270980946634, 18.250000000000)
    pts(17) = Array(0.531008696075, 6.051748075118, 18.250000000000)
    pts(18) = Array(0.506496816134, 5.828032341644, 18.250000000000)
    pts(19) = Array(0.487782506992, 5.603810598680, 18.250000000000)
    FSData_490 = pts
End Function

Private Function FSData_491() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-1.625316986357, 5.149596556417, 15.150000000000)
    pts(1) = Array(-2.492041982786, 4.790587308048, 15.150000000000)
    pts(2) = Array(-2.595137748753, 4.990579632167, 15.150000000000)
    pts(3) = Array(-2.698233462143, 5.190571855178, 15.150000000000)
    pts(4) = Array(-2.806491791085, 5.387877952086, 15.150000000000)
    pts(5) = Array(-2.957809958732, 5.562495864990, 15.150000000000)
    pts(6) = Array(-3.137063178437, 5.721403640235, 15.150000000000)
    pts(7) = Array(-3.337767274107, 5.867010279853, 15.150000000000)
    pts(8) = Array(-3.557032174135, 5.999845590694, 15.150000000000)
    pts(9) = Array(-3.793031109387, 6.119878675531, 15.150000000000)
    pts(10) = Array(-4.109117035169, 6.629868565159, 15.150000000000)
    pts(11) = Array(-1.782440500542, 7.593609541057, 15.150000000000)
    pts(12) = Array(-1.645329692808, 7.009485730206, 15.150000000000)
    pts(13) = Array(-1.727329932022, 6.757733074481, 15.150000000000)
    pts(14) = Array(-1.788444880632, 6.508760627719, 15.150000000000)
    pts(15) = Array(-1.827404665432, 6.263881958399, 15.150000000000)
    pts(16) = Array(-1.841791067176, 6.024766025737, 15.150000000000)
    pts(17) = Array(-1.825315659305, 5.794294412946, 15.150000000000)
    pts(18) = Array(-1.762349378727, 5.578227735339, 15.150000000000)
    pts(19) = Array(-1.693833199700, 5.363912200213, 15.150000000000)
    FSData_491 = pts
End Function

Private Function FSData_492() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-1.625316986357, 5.149596556417, 18.250000000000)
    pts(1) = Array(-2.492041982786, 4.790587308048, 18.250000000000)
    pts(2) = Array(-2.595137748753, 4.990579632167, 18.250000000000)
    pts(3) = Array(-2.698233462143, 5.190571855178, 18.250000000000)
    pts(4) = Array(-2.806491791085, 5.387877952086, 18.250000000000)
    pts(5) = Array(-2.957809958732, 5.562495864990, 18.250000000000)
    pts(6) = Array(-3.137063178437, 5.721403640235, 18.250000000000)
    pts(7) = Array(-3.337767274107, 5.867010279853, 18.250000000000)
    pts(8) = Array(-3.557032174135, 5.999845590694, 18.250000000000)
    pts(9) = Array(-3.793031109387, 6.119878675531, 18.250000000000)
    pts(10) = Array(-4.109117035169, 6.629868565159, 18.250000000000)
    pts(11) = Array(-1.782440500542, 7.593609541057, 18.250000000000)
    pts(12) = Array(-1.645329692808, 7.009485730206, 18.250000000000)
    pts(13) = Array(-1.727329932022, 6.757733074481, 18.250000000000)
    pts(14) = Array(-1.788444880632, 6.508760627719, 18.250000000000)
    pts(15) = Array(-1.827404665432, 6.263881958399, 18.250000000000)
    pts(16) = Array(-1.841791067176, 6.024766025737, 18.250000000000)
    pts(17) = Array(-1.825315659305, 5.794294412946, 18.250000000000)
    pts(18) = Array(-1.762349378727, 5.578227735339, 18.250000000000)
    pts(19) = Array(-1.693833199700, 5.363912200213, 18.250000000000)
    FSData_492 = pts
End Function

Private Function FSData_493() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-3.472262383043, 4.135624976143, 15.150000000000)
    pts(1) = Array(-4.135624976143, 3.472262383043, 15.150000000000)
    pts(2) = Array(-4.307406793249, 3.617578156373, 15.150000000000)
    pts(3) = Array(-4.479188523089, 3.762893856413, 15.150000000000)
    pts(4) = Array(-4.654711951812, 3.903752252084, 15.150000000000)
    pts(5) = Array(-4.861335092061, 4.007171212051, 15.150000000000)
    pts(6) = Array(-5.087754845744, 4.085385615778, 15.150000000000)
    pts(7) = Array(-5.328902500449, 4.143102477698, 15.150000000000)
    pts(8) = Array(-5.582310726475, 4.181917258039, 15.150000000000)
    pts(9) = Array(-5.846279985351, 4.202500485768, 15.150000000000)
    pts(10) = Array(-6.333469984131, 4.552708859582, 15.150000000000)
    pts(11) = Array(-4.552708859582, 6.333469984131, 15.150000000000)
    pts(12) = Array(-4.202500485768, 5.846279985351, 15.150000000000)
    pts(13) = Array(-4.181917258039, 5.582310726475, 15.150000000000)
    pts(14) = Array(-4.143102477698, 5.328902500449, 15.150000000000)
    pts(15) = Array(-4.085385615778, 5.087754845744, 15.150000000000)
    pts(16) = Array(-4.007171212051, 4.861335092061, 15.150000000000)
    pts(17) = Array(-3.903752252084, 4.654711951812, 15.150000000000)
    pts(18) = Array(-3.762893856413, 4.479188523089, 15.150000000000)
    pts(19) = Array(-3.617578156373, 4.307406793249, 15.150000000000)
    FSData_493 = pts
End Function

Private Function FSData_494() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-3.472262383043, 4.135624976143, 18.250000000000)
    pts(1) = Array(-4.135624976143, 3.472262383043, 18.250000000000)
    pts(2) = Array(-4.307406793249, 3.617578156373, 18.250000000000)
    pts(3) = Array(-4.479188523089, 3.762893856413, 18.250000000000)
    pts(4) = Array(-4.654711951812, 3.903752252084, 18.250000000000)
    pts(5) = Array(-4.861335092061, 4.007171212051, 18.250000000000)
    pts(6) = Array(-5.087754845744, 4.085385615778, 18.250000000000)
    pts(7) = Array(-5.328902500449, 4.143102477698, 18.250000000000)
    pts(8) = Array(-5.582310726475, 4.181917258039, 18.250000000000)
    pts(9) = Array(-5.846279985351, 4.202500485768, 18.250000000000)
    pts(10) = Array(-6.333469984131, 4.552708859582, 18.250000000000)
    pts(11) = Array(-4.552708859582, 6.333469984131, 18.250000000000)
    pts(12) = Array(-4.202500485768, 5.846279985351, 18.250000000000)
    pts(13) = Array(-4.181917258039, 5.582310726475, 18.250000000000)
    pts(14) = Array(-4.143102477698, 5.328902500449, 18.250000000000)
    pts(15) = Array(-4.085385615778, 5.087754845744, 18.250000000000)
    pts(16) = Array(-4.007171212051, 4.861335092061, 18.250000000000)
    pts(17) = Array(-3.903752252084, 4.654711951812, 18.250000000000)
    pts(18) = Array(-3.762893856413, 4.479188523089, 18.250000000000)
    pts(19) = Array(-3.617578156373, 4.307406793249, 18.250000000000)
    FSData_494 = pts
End Function

Private Function FSData_495() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-4.790587308048, 2.492041982786, 15.150000000000)
    pts(1) = Array(-5.149596556417, 1.625316986357, 15.150000000000)
    pts(2) = Array(-5.363912200213, 1.693833199700, 15.150000000000)
    pts(3) = Array(-5.578227735339, 1.762349378727, 15.150000000000)
    pts(4) = Array(-5.794294412946, 1.825315659305, 15.150000000000)
    pts(5) = Array(-6.024766025737, 1.841791067176, 15.150000000000)
    pts(6) = Array(-6.263881958399, 1.827404665432, 15.150000000000)
    pts(7) = Array(-6.508760627719, 1.788444880632, 15.150000000000)
    pts(8) = Array(-6.757733074481, 1.727329932022, 15.150000000000)
    pts(9) = Array(-7.009485730206, 1.645329692808, 15.150000000000)
    pts(10) = Array(-7.593609541057, 1.782440500542, 15.150000000000)
    pts(11) = Array(-6.629868565159, 4.109117035169, 15.150000000000)
    pts(12) = Array(-6.119878675531, 3.793031109387, 15.150000000000)
    pts(13) = Array(-5.999845590694, 3.557032174135, 15.150000000000)
    pts(14) = Array(-5.867010279853, 3.337767274107, 15.150000000000)
    pts(15) = Array(-5.721403640235, 3.137063178437, 15.150000000000)
    pts(16) = Array(-5.562495864990, 2.957809958732, 15.150000000000)
    pts(17) = Array(-5.387877952086, 2.806491791084, 15.150000000000)
    pts(18) = Array(-5.190571855178, 2.698233462143, 15.150000000000)
    pts(19) = Array(-4.990579632167, 2.595137748753, 15.150000000000)
    FSData_495 = pts
End Function

Private Function FSData_496() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-4.790587308048, 2.492041982786, 18.250000000000)
    pts(1) = Array(-5.149596556417, 1.625316986357, 18.250000000000)
    pts(2) = Array(-5.363912200213, 1.693833199700, 18.250000000000)
    pts(3) = Array(-5.578227735339, 1.762349378727, 18.250000000000)
    pts(4) = Array(-5.794294412946, 1.825315659305, 18.250000000000)
    pts(5) = Array(-6.024766025737, 1.841791067176, 18.250000000000)
    pts(6) = Array(-6.263881958399, 1.827404665432, 18.250000000000)
    pts(7) = Array(-6.508760627719, 1.788444880632, 18.250000000000)
    pts(8) = Array(-6.757733074481, 1.727329932022, 18.250000000000)
    pts(9) = Array(-7.009485730206, 1.645329692808, 18.250000000000)
    pts(10) = Array(-7.593609541057, 1.782440500542, 18.250000000000)
    pts(11) = Array(-6.629868565159, 4.109117035169, 18.250000000000)
    pts(12) = Array(-6.119878675531, 3.793031109387, 18.250000000000)
    pts(13) = Array(-5.999845590694, 3.557032174135, 18.250000000000)
    pts(14) = Array(-5.867010279853, 3.337767274107, 18.250000000000)
    pts(15) = Array(-5.721403640235, 3.137063178437, 18.250000000000)
    pts(16) = Array(-5.562495864990, 2.957809958732, 18.250000000000)
    pts(17) = Array(-5.387877952086, 2.806491791084, 18.250000000000)
    pts(18) = Array(-5.190571855178, 2.698233462143, 18.250000000000)
    pts(19) = Array(-4.990579632167, 2.595137748753, 18.250000000000)
    FSData_496 = pts
End Function

Private Function FSData_497() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-5.379588742184, 0.469068187966, 15.150000000000)
    pts(1) = Array(-5.379588742184, -0.469068187966, 15.150000000000)
    pts(2) = Array(-5.603810598680, -0.487782506992, 15.150000000000)
    pts(3) = Array(-5.828032341644, -0.506496816134, 15.150000000000)
    pts(4) = Array(-6.051748075118, -0.531008696075, 15.150000000000)
    pts(5) = Array(-6.270980946634, -0.603985071800, 15.150000000000)
    pts(6) = Array(-6.486389825119, -0.708782079762, 15.150000000000)
    pts(7) = Array(-6.697718951481, -0.838487237218, 15.150000000000)
    pts(8) = Array(-6.904351820901, -0.990227717860, 15.150000000000)
    pts(9) = Array(-7.105560813783, -1.162327630931, 15.150000000000)
    pts(10) = Array(-7.697690881599, -1.259188266842, 15.150000000000)
    pts(11) = Array(-7.697690881599, 1.259188266842, 15.150000000000)
    pts(12) = Array(-7.105560813783, 1.162327630931, 15.150000000000)
    pts(13) = Array(-6.904351820901, 0.990227717860, 15.150000000000)
    pts(14) = Array(-6.697718951481, 0.838487237218, 15.150000000000)
    pts(15) = Array(-6.486389825119, 0.708782079762, 15.150000000000)
    pts(16) = Array(-6.270980946634, 0.603985071800, 15.150000000000)
    pts(17) = Array(-6.051748075118, 0.531008696075, 15.150000000000)
    pts(18) = Array(-5.828032341644, 0.506496816134, 15.150000000000)
    pts(19) = Array(-5.603810598680, 0.487782506992, 15.150000000000)
    FSData_497 = pts
End Function

Private Function FSData_498() As Variant
    Dim pts(19) As Variant
    pts(0) = Array(-5.379588742184, 0.469068187966, 18.250000000000)
    pts(1) = Array(-5.379588742184, -0.469068187966, 18.250000000000)
    pts(2) = Array(-5.603810598680, -0.487782506992, 18.250000000000)
    pts(3) = Array(-5.828032341644, -0.506496816134, 18.250000000000)
    pts(4) = Array(-6.051748075118, -0.531008696075, 18.250000000000)
    pts(5) = Array(-6.270980946634, -0.603985071800, 18.250000000000)
    pts(6) = Array(-6.486389825119, -0.708782079762, 18.250000000000)
    pts(7) = Array(-6.697718951481, -0.838487237218, 18.250000000000)
    pts(8) = Array(-6.904351820901, -0.990227717860, 18.250000000000)
    pts(9) = Array(-7.105560813783, -1.162327630931, 18.250000000000)
    pts(10) = Array(-7.697690881599, -1.259188266842, 18.250000000000)
    pts(11) = Array(-7.697690881599, 1.259188266842, 18.250000000000)
    pts(12) = Array(-7.105560813783, 1.162327630931, 18.250000000000)
    pts(13) = Array(-6.904351820901, 0.990227717860, 18.250000000000)
    pts(14) = Array(-6.697718951481, 0.838487237218, 18.250000000000)
    pts(15) = Array(-6.486389825119, 0.708782079762, 18.250000000000)
    pts(16) = Array(-6.270980946634, 0.603985071800, 18.250000000000)
    pts(17) = Array(-6.051748075118, 0.531008696075, 18.250000000000)
    pts(18) = Array(-5.828032341644, 0.506496816134, 18.250000000000)
    pts(19) = Array(-5.603810598680, 0.487782506992, 18.250000000000)
    FSData_498 = pts
End Function

Private Function FocusPlacement(ByVal label As String, ByVal angleDeg As Double) As Variant
    Dim m(15) As Double, known As Boolean, moving As Boolean, ring As Boolean
    Dim theta As Double, dz As Double
    m(0) = 1#: m(4) = 1#: m(8) = 1#: m(12) = 1#
    dz = FocusTravel(angleDeg) / 1000#
    If label = "01_FixedGuideBarrel" Or InStr(label, "_01_FixedGuideBarrel-") > 0 Then
        known = True
    End If
    If label = "02_L1Retainer" Or InStr(label, "_02_L1Retainer-") > 0 Then
        known = True
    End If
    If label = "03_L2MovingCell" Or InStr(label, "_03_L2MovingCell-") > 0 Then
        known = True
        moving = True
    End If
    If label = "04_L2Retainer" Or InStr(label, "_04_L2Retainer-") > 0 Then
        known = True
        moving = True
    End If
    If label = "09_L7Retainer" Or InStr(label, "_09_L7Retainer-") > 0 Then
        known = True
    End If
    If label = "10_GearedCamRing" Or InStr(label, "_10_GearedCamRing-") > 0 Then
        known = True
        ring = True
        theta = angleDeg * 0.0174532925199433 * 1#
    End If
    If label = "11_CamKeeper" Or InStr(label, "_11_CamKeeper-") > 0 Then
        known = True
    End If
    If label = "12_FrontThrustWasher" Or InStr(label, "_12_FrontThrustWasher-") > 0 Then
        known = True
    End If
    If label = "13_RearThrustWasher" Or InStr(label, "_13_RearThrustWasher-") > 0 Then
        known = True
    End If
    If label = "14_InputPinion" Or InStr(label, "_14_InputPinion-") > 0 Then
        known = True
        ring = True
        m(9) = 0.038400000000: m(10) = 0.000000000000
        theta = angleDeg * 0.0174532925199433 * -5#
    End If
    If label = "05_RearRectangularSpacer" Or InStr(label, "_05_RearRectangularSpacer-") > 0 Then
        known = True
    End If
    If label = "Lens1_Optical_Reference" Or InStr(label, "_Lens1_Optical_Reference-") > 0 Then
        known = True
    End If
    If label = "Lens2_Optical_Reference" Or InStr(label, "_Lens2_Optical_Reference-") > 0 Then
        known = True
        moving = True
    End If
    If label = "Lens3_RevisedBlank_Reference" Or InStr(label, "_Lens3_RevisedBlank_Reference-") > 0 Then
        known = True
    End If
    If label = "Lens4_RevisedBlank_Reference" Or InStr(label, "_Lens4_RevisedBlank_Reference-") > 0 Then
        known = True
    End If
    If label = "Lens5_RevisedBlank_Reference" Or InStr(label, "_Lens5_RevisedBlank_Reference-") > 0 Then
        known = True
    End If
    If label = "Lens6_RevisedBlank_Reference" Or InStr(label, "_Lens6_RevisedBlank_Reference-") > 0 Then
        known = True
    End If
    If label = "Lens7_Optical_Reference" Or InStr(label, "_Lens7_Optical_Reference-") > 0 Then
        known = True
    End If
    If label = "21_GuideAxle" Or InStr(label, "_21_GuideAxle-") > 0 Then
        known = True
        moving = True
        m(0) = -0.000000000000000
        m(1) = 1.000000000000000
        m(2) = 0.000000000000000
        m(3) = 0.000000000000000
        m(4) = 0.000000000000000
        m(5) = 1.000000000000000
        m(6) = 1.000000000000000
        m(7) = 0.000000000000000
        m(8) = 0.000000000000000
        m(11) = 0.024000000000
    End If
    If label = "31_CamRoller" Or InStr(label, "_31_CamRoller-") > 0 Then
        known = True
        moving = True
        m(0) = -0.000000000000000
        m(1) = 1.000000000000000
        m(2) = 0.000000000000000
        m(3) = 0.000000000000000
        m(4) = 0.000000000000000
        m(5) = 1.000000000000000
        m(6) = 1.000000000000000
        m(7) = 0.000000000000000
        m(8) = 0.000000000000000
        m(11) = 0.024000000000
    End If
    If label = "22_GuideAxle" Or InStr(label, "_22_GuideAxle-") > 0 Then
        known = True
        moving = True
        m(0) = -0.866025403784439
        m(1) = -0.500000000000000
        m(2) = 0.000000000000000
        m(3) = 0.000000000000000
        m(4) = 0.000000000000000
        m(5) = 1.000000000000000
        m(6) = -0.500000000000000
        m(7) = 0.866025403784439
        m(8) = 0.000000000000000
        m(11) = 0.024000000000
    End If
    If label = "32_CamRoller" Or InStr(label, "_32_CamRoller-") > 0 Then
        known = True
        moving = True
        m(0) = -0.866025403784439
        m(1) = -0.500000000000000
        m(2) = 0.000000000000000
        m(3) = 0.000000000000000
        m(4) = 0.000000000000000
        m(5) = 1.000000000000000
        m(6) = -0.500000000000000
        m(7) = 0.866025403784439
        m(8) = 0.000000000000000
        m(11) = 0.024000000000
    End If
    If label = "23_GuideAxle" Or InStr(label, "_23_GuideAxle-") > 0 Then
        known = True
        moving = True
        m(0) = 0.866025403784438
        m(1) = -0.500000000000000
        m(2) = 0.000000000000000
        m(3) = 0.000000000000000
        m(4) = 0.000000000000000
        m(5) = 1.000000000000000
        m(6) = -0.500000000000000
        m(7) = -0.866025403784438
        m(8) = 0.000000000000000
        m(11) = 0.024000000000
    End If
    If label = "33_CamRoller" Or InStr(label, "_33_CamRoller-") > 0 Then
        known = True
        moving = True
        m(0) = 0.866025403784438
        m(1) = -0.500000000000000
        m(2) = 0.000000000000000
        m(3) = 0.000000000000000
        m(4) = 0.000000000000000
        m(5) = 1.000000000000000
        m(6) = -0.500000000000000
        m(7) = -0.866025403784438
        m(8) = 0.000000000000000
        m(11) = 0.024000000000
    End If
    If Not known Then Err.Raise vbObjectError + 810, , "Unrecognized focus component."
    If moving Then m(11) = m(11) + dz
    If ring Then
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

Public Sub FocusMinus2mm()
    FocusSetAngle -40#
End Sub

Public Sub FocusMiddle()
    FocusSetAngle 0#
End Sub

Public Sub FocusPlus2mm()
    FocusSetAngle 40#
End Sub

Public Sub FocusSetAngle(ByVal angleDeg As Double)
    On Error GoTo FocusFailed
    Dim model As SldWorks.ModelDoc2, assembly As SldWorks.AssemblyDoc
    Dim component As SldWorks.Component2, components As Variant, matrix As Variant
    Dim math As SldWorks.MathUtility, move As SldWorks.MathTransform, i As Long
    If Abs(angleDeg) > 40# Then Err.Raise vbObjectError + 805, , "Angle must be -40 to +40 degrees."
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
    MsgBox "Focus angle=" & CStr(angleDeg) & " deg; L2 axial shift=" & CStr(FocusTravel(angleDeg)) & " mm." & vbCrLf & "Only L2 moves; cam/pinion rotate. No assembly mates or optical distance calibration.", vbInformation
    Exit Sub
FocusFailed:
    MsgBox "Focus pose stopped: " & Err.Description, vbExclamation
End Sub

Private Function FocusKnownComponent(ByVal label As String) As Boolean
    FocusKnownComponent = (InStr(label, "F35_") > 0)
End Function

Private Function FocusTravel(ByVal angleDeg As Double) As Double
    Dim u As Double
    u = (angleDeg + 40#) / 80#
    FocusTravel = -2# + 4# * (3# * u * u - 2# * u * u * u)
End Function
