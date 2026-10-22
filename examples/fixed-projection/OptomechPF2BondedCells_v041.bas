Attribute VB_Name = "OptomechPF2BondedCells_v041"
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
    Dim loaded(18) As SldWorks.ModelDoc2
    Set app = Application.SldWorks
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = fs.GetParentFolderName(app.GetCurrentMacroPathName)
    If Len(folder) = 0 Then Err.Raise vbObjectError + 400, , "Save the macro before running."
    output = folder & "\Optomech_" & Format$(Now, "yyyymmdd_hhnnss")
    runID = "OM_" & Format$(Now, "yyyymmdd_hhnnss")
    fs.CreateFolder output
    Set log = fs.CreateTextFile(output & "\build_log.txt", False, True)
    Set buildLog = log
    log.WriteLine "Optomech workflow concept; SOLIDWORKS " & app.RevisionNumber
    names = Array("01_CommonBarrel", "02_L1_BondedCell", "03_L2_BondedCell", "04_L3_BondedCell", "05_G4G5_Retainer", "06_Spacer_G4_G5", "07_RearSpacer_1", "08_RearSpacer_2", "09_RearRetainer", "Lens1_Optical_Reference", "Lens2_Optical_Reference", "Lens3_RevisedBlank_Reference", "Lens4_Optical_Reference", "Lens5_Optical_Reference", "Lens6_Optical_Reference", "Lens7_Optical_Reference", "Lens8_Optical_Reference", "Lens9_Optical_Reference", "Lens10_Optical_Reference")
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
        matrix(11) = zpos(i) / 1000#
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
    a = SketchPoint(math, mapping, -1.000000000000, 0#)
    b = SketchPoint(math, mapping, 98.268860245388, 0#)
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
    AddBondedCellDetails model, plane, axis, label
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
    If InStr(label, "BondedCell") > 0 Then
        props.Add3 "CELL_LOCKING", 30, "External thread to common barrel; rear face against own shoulder. No master retaining ring.", 2
        props.Add3 "GLASS_FIXING", 30, "Lens seated and optically centered on bench; three local bonds. Replace lens-cell assembly for service.", 2
    End If
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
    If InStr(key, "01_commonbarrel") > 0 Then
        values(0) = 0.18: values(1) = 0.31: values(2) = 0.65
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "02_l1_bondedcell") > 0 Then
        values(0) = 0.52: values(1) = 0.55: values(2) = 0.6
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "03_l2_bondedcell") > 0 Then
        values(0) = 0.52: values(1) = 0.55: values(2) = 0.6
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "04_l3_bondedcell") > 0 Then
        values(0) = 0.52: values(1) = 0.55: values(2) = 0.6
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "05_g4g5_retainer") > 0 Then
        values(0) = 0.24: values(1) = 0.27: values(2) = 0.31
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "06_spacer_g4_g5") > 0 Then
        values(0) = 0.23: values(1) = 0.61: values(2) = 0.32
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "07_rearspacer_1") > 0 Then
        values(0) = 0.23: values(1) = 0.61: values(2) = 0.32
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "08_rearspacer_2") > 0 Then
        values(0) = 0.23: values(1) = 0.61: values(2) = 0.32
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "09_rearretainer") > 0 Then
        values(0) = 0.24: values(1) = 0.27: values(2) = 0.31
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
        Case Else: Err.Raise vbObjectError + 418, , "Unknown profile index."
    End Select
End Function

Private Function Profile_0() As Variant
    Dim pts(33) As Variant
    pts(0) = Array(2.619523040734, 32.500000000000)
    pts(1) = Array(15.524273656847, 32.500000000000)
    pts(2) = Array(16.800000000000, 29.000000000000)
    pts(3) = Array(23.398871715914, 29.000000000000)
    pts(4) = Array(24.400000000000, 25.500000000000)
    pts(5) = Array(39.200000000000, 25.500000000000)
    pts(6) = Array(40.200000000000, 21.600000000000)
    pts(7) = Array(97.268860245388, 21.600000000000)
    pts(8) = Array(97.268860245388, 19.000152279125)
    pts(9) = Array(94.168860245388, 19.000152279125)
    pts(10) = Array(93.468860245388, 18.530152279125)
    pts(11) = Array(86.637064283183, 18.530152279125)
    pts(12) = Array(85.887064283183, 17.578123947777)
    pts(13) = Array(79.700736699199, 17.578123947777)
    pts(14) = Array(78.950736699199, 14.525475919527)
    pts(15) = Array(72.605226962642, 14.525475919527)
    pts(16) = Array(72.605226962642, 13.645475919527)
    pts(17) = Array(71.825739509650, 13.645475919527)
    pts(18) = Array(63.317968956453, 5.137705366330)
    pts(19) = Array(62.317968956453, 5.137705366330)
    pts(20) = Array(54.664713965508, 12.790960357275)
    pts(21) = Array(53.743220456623, 12.790960357275)
    pts(22) = Array(53.743220456623, 13.670960357275)
    pts(23) = Array(47.533458635636, 13.670960357275)
    pts(24) = Array(46.783458635636, 18.211991759500)
    pts(25) = Array(38.405122642205, 18.211991759500)
    pts(26) = Array(38.130122642205, 19.050000000000)
    pts(27) = Array(23.398871715914, 19.050000000000)
    pts(28) = Array(23.398871715914, 24.580000000000)
    pts(29) = Array(19.134889960676, 24.580000000000)
    pts(30) = Array(19.134889960676, 26.080000000000)
    pts(31) = Array(15.524273656847, 26.080000000000)
    pts(32) = Array(15.524273656847, 30.080000000000)
    pts(33) = Array(2.619523040734, 30.080000000000)
    Profile_0 = pts
End Function

Private Function Profile_1() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(5.800000000000, 30.000000000000)
    pts(1) = Array(15.524273656847, 30.000000000000)
    pts(2) = Array(15.524273656847, 26.500000000000)
    pts(3) = Array(13.524273656847, 26.500000000000)
    pts(4) = Array(13.524273656847, 27.874249726437)
    pts(5) = Array(5.800000000000, 27.874249726437)
    Profile_1 = pts
End Function

Private Function Profile_2() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(13.750000000000, 26.000000000000)
    pts(1) = Array(19.134889960676, 26.000000000000)
    pts(2) = Array(19.134889960676, 22.500000000000)
    pts(3) = Array(18.134889960676, 22.500000000000)
    pts(4) = Array(18.134889960676, 23.174930206996)
    pts(5) = Array(13.750000000000, 23.174930206996)
    Profile_2 = pts
End Function

Private Function Profile_3() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(19.350000000000, 24.500000000000)
    pts(1) = Array(23.398871715914, 24.500000000000)
    pts(2) = Array(23.398871715914, 21.800000000000)
    pts(3) = Array(21.398871715914, 21.800000000000)
    pts(4) = Array(21.398871715914, 22.380000000000)
    pts(5) = Array(19.350000000000, 22.380000000000)
    Profile_3 = pts
End Function

Private Function Profile_4() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(36.505122642205, 18.980000000000)
    pts(1) = Array(38.105122642205, 18.980000000000)
    pts(2) = Array(38.105122642205, 17.831991759500)
    pts(3) = Array(36.505122642205, 17.831991759500)
    Profile_4 = pts
End Function

Private Function Profile_5() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(38.962257061193, 18.181991759500)
    pts(1) = Array(46.733458635636, 18.181991759500)
    pts(2) = Array(46.733458635636, 13.290960357275)
    pts(3) = Array(38.962257061193, 17.831991759500)
    Profile_5 = pts
End Function

Private Function Profile_6() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(79.750736699199, 17.548123947777)
    pts(1) = Array(84.623573372765, 17.548123947777)
    pts(2) = Array(84.623573372765, 16.917212754320)
    pts(3) = Array(79.750736699199, 14.145475919527)
    Profile_6 = pts
End Function

Private Function Profile_7() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(86.687064283183, 18.500152279125)
    pts(1) = Array(93.410974640273, 18.500152279125)
    pts(2) = Array(93.410974640273, 18.150152279125)
    pts(3) = Array(86.687064283183, 17.198123947777)
    Profile_7 = pts
End Function

Private Function Profile_8() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(94.268860245388, 18.930152279125)
    pts(1) = Array(96.968860245388, 18.930152279125)
    pts(2) = Array(96.968860245388, 18.150152279125)
    pts(3) = Array(94.268860245388, 18.150152279125)
    Profile_8 = pts
End Function

Private Function Profile_9() As Variant
    Dim pts(175) As Variant
    pts(0) = Array(0.000000000000, 0.000000000000)
    pts(1) = Array(0.001531277799, 0.463237495441)
    pts(2) = Array(0.006125311999, 0.926474990881)
    pts(3) = Array(0.013782705139, 1.389712486322)
    pts(4) = Array(0.023294352900, 1.806626232218)
    pts(5) = Array(0.036775228783, 2.269863727659)
    pts(6) = Array(0.051530570663, 2.686777473556)
    pts(7) = Array(0.070841322637, 3.150014968996)
    pts(8) = Array(0.090847362049, 3.566928714893)
    pts(9) = Array(0.115997164616, 4.030166210333)
    pts(10) = Array(0.144224783784, 4.493403705774)
    pts(11) = Array(0.175533948612, 4.956641201215)
    pts(12) = Array(0.206350334255, 5.373554947111)
    pts(13) = Array(0.243526194872, 5.836792442552)
    pts(14) = Array(0.279630298757, 6.253706188448)
    pts(15) = Array(0.322690683624, 6.716943683889)
    pts(16) = Array(0.364099856621, 7.133857429785)
    pts(17) = Array(0.413065506670, 7.597094925226)
    pts(18) = Array(0.459799742756, 8.014008671123)
    pts(19) = Array(0.514694361939, 8.477246166563)
    pts(20) = Array(0.566776347122, 8.894159912460)
    pts(21) = Array(0.627626660665, 9.357397407900)
    pts(22) = Array(0.685081828994, 9.774311153797)
    pts(23) = Array(0.751917649691, 10.237548649238)
    pts(24) = Array(0.814774246663, 10.654462395134)
    pts(25) = Array(0.887628549609, 11.117699890575)
    pts(26) = Array(0.955917703368, 11.534613636471)
    pts(27) = Array(1.034826709560, 11.997851131912)
    pts(28) = Array(1.108582510307, 12.414764877808)
    pts(29) = Array(1.193585779645, 12.878002373249)
    pts(30) = Array(1.272845367746, 13.294916119146)
    pts(31) = Array(1.363985901986, 13.758153614586)
    pts(32) = Array(1.448789565330, 14.175067360483)
    pts(33) = Array(1.546113921652, 14.638304855923)
    pts(34) = Array(1.636505202876, 15.055218601820)
    pts(35) = Array(1.740063618791, 15.518456097261)
    pts(36) = Array(1.836089433062, 15.935369843157)
    pts(37) = Array(1.945935963472, 16.398607338598)
    pts(38) = Array(2.047646727616, 16.815521084494)
    pts(39) = Array(2.152069002516, 17.232434830391)
    pts(40) = Array(2.259215314706, 17.649348576287)
    pts(41) = Array(2.381477376149, 18.112586071728)
    pts(42) = Array(2.494417412940, 18.529499817625)
    pts(43) = Array(2.610123100477, 18.946413563521)
    pts(44) = Array(2.728608694682, 19.363327309418)
    pts(45) = Array(2.863537634029, 19.826564804858)
    pts(46) = Array(2.987940747522, 20.243478550755)
    pts(47) = Array(3.115171005995, 20.660392296651)
    pts(48) = Array(3.245244557738, 21.077306042548)
    pts(49) = Array(3.393125704514, 21.540543537989)
    pts(50) = Array(3.529257006655, 21.957457283885)
    pts(51) = Array(3.668285017875, 22.374371029782)
    pts(52) = Array(3.810227972204, 22.791284775678)
    pts(53) = Array(3.955104644211, 23.208198521575)
    pts(54) = Array(4.102934363145, 23.625112267471)
    pts(55) = Array(4.253737027685, 24.042026013368)
    pts(56) = Array(4.407533121334, 24.458939759264)
    pts(57) = Array(4.564343728472, 24.875853505161)
    pts(58) = Array(4.724190551130, 25.292767251058)
    pts(59) = Array(4.887095926485, 25.709680996954)
    pts(60) = Array(5.053082845145, 26.126594742851)
    pts(61) = Array(5.222174970238, 26.543508488747)
    pts(62) = Array(5.394396657364, 26.960422234644)
    pts(63) = Array(5.569772975441, 27.377335980540)
    pts(64) = Array(5.748329728506, 27.794249726437)
    pts(65) = Array(13.524273656847, 27.794249726437)
    pts(66) = Array(13.524273656847, 23.271534595474)
    pts(67) = Array(13.348268642548, 23.115551022487)
    pts(68) = Array(13.193699579191, 22.976579773855)
    pts(69) = Array(13.041192662852, 22.837608525222)
    pts(70) = Array(12.840980878828, 22.652313527046)
    pts(71) = Array(12.693114937128, 22.513342278414)
    pts(72) = Array(12.498940421564, 22.328047280238)
    pts(73) = Array(12.355494442888, 22.189076031605)
    pts(74) = Array(12.167076139951, 22.003781033429)
    pts(75) = Array(12.027848934479, 21.864809784797)
    pts(76) = Array(11.844930560760, 21.679514786621)
    pts(77) = Array(11.665048449185, 21.494219788445)
    pts(78) = Array(11.488129787763, 21.308924790268)
    pts(79) = Array(11.357343708788, 21.169953541636)
    pts(80) = Array(11.185446141868, 20.984658543460)
    pts(81) = Array(11.016328938534, 20.799363545284)
    pts(82) = Array(10.849931142337, 20.614068547107)
    pts(83) = Array(10.726882268465, 20.475097298475)
    pts(84) = Array(10.565104809002, 20.289802300299)
    pts(85) = Array(10.405892603066, 20.104507302123)
    pts(86) = Array(10.249194043886, 19.919212303946)
    pts(87) = Array(10.094959682504, 19.733917305770)
    pts(88) = Array(9.943142106279, 19.548622307594)
    pts(89) = Array(9.793695826148, 19.363327309418)
    pts(90) = Array(9.646577171896, 19.178032311241)
    pts(91) = Array(9.501744194750, 18.992737313065)
    pts(92) = Array(9.359156576672, 18.807442314889)
    pts(93) = Array(9.218775545806, 18.622147316713)
    pts(94) = Array(9.080563797566, 18.436852318536)
    pts(95) = Array(8.944485420925, 18.251557320360)
    pts(96) = Array(8.810505829474, 18.066262322184)
    pts(97) = Array(8.678591696900, 17.880967324008)
    pts(98) = Array(8.516554749342, 17.649348576287)
    pts(99) = Array(8.389172172152, 17.464053578111)
    pts(100) = Array(8.263754740225, 17.278758579935)
    pts(101) = Array(8.140273851718, 17.093463581759)
    pts(102) = Array(7.988604017120, 16.861844834038)
    pts(103) = Array(7.869380788763, 16.676549835862)
    pts(104) = Array(7.752007703891, 16.491254837686)
    pts(105) = Array(7.636460043393, 16.305959839510)
    pts(106) = Array(7.494556125565, 16.074341091789)
    pts(107) = Array(7.383029445590, 15.889046093613)
    pts(108) = Array(7.246080241115, 15.657427345893)
    pts(109) = Array(7.138462107481, 15.472132347717)
    pts(110) = Array(7.006332107756, 15.240513599996)
    pts(111) = Array(6.902517967029, 15.055218601820)
    pts(112) = Array(6.775081076925, 14.823599854100)
    pts(113) = Array(6.674973576400, 14.638304855923)
    pts(114) = Array(6.552112314071, 14.406686108203)
    pts(115) = Array(6.455620698794, 14.221391110027)
    pts(116) = Array(6.337225485084, 13.989772362307)
    pts(117) = Array(6.244265061608, 13.804477364130)
    pts(118) = Array(6.130233587219, 13.572858616410)
    pts(119) = Array(6.040725245397, 13.387563618234)
    pts(120) = Array(5.930961905824, 13.155944870513)
    pts(121) = Array(5.844831691267, 12.970649872337)
    pts(122) = Array(5.739247080979, 12.739031124617)
    pts(123) = Array(5.656425811901, 12.553736126441)
    pts(124) = Array(5.554936270515, 12.322117378720)
    pts(125) = Array(5.455687597839, 12.090498631000)
    pts(126) = Array(5.358656491558, 11.858879883280)
    pts(127) = Array(5.282613127125, 11.673584885104)
    pts(128) = Array(5.189517554242, 11.441966137383)
    pts(129) = Array(5.098578628886, 11.210347389663)
    pts(130) = Array(5.009776049601, 10.978728641943)
    pts(131) = Array(4.889004354146, 10.654462395134)
    pts(132) = Array(4.772330183666, 10.330196148326)
    pts(133) = Array(4.659704641380, 10.005929901517)
    pts(134) = Array(4.597146961615, 9.820634903341)
    pts(135) = Array(4.535887742137, 9.635339905165)
    pts(136) = Array(4.431784343256, 9.311073658356)
    pts(137) = Array(4.331591230528, 8.986807411548)
    pts(138) = Array(4.235268904508, 8.662541164740)
    pts(139) = Array(4.181950977226, 8.477246166563)
    pts(140) = Array(4.129878003929, 8.291951168387)
    pts(141) = Array(4.041726526808, 7.967684921579)
    pts(142) = Array(3.957334517027, 7.643418674770)
    pts(143) = Array(3.876670368576, 7.319152427962)
    pts(144) = Array(3.832238734892, 7.133857429785)
    pts(145) = Array(3.789009245886, 6.948562431609)
    pts(146) = Array(3.716234892572, 6.624296184801)
    pts(147) = Array(3.676285234384, 6.439001186625)
    pts(148) = Array(3.637519262440, 6.253706188448)
    pts(149) = Array(3.572513522693, 5.929439941640)
    pts(150) = Array(3.536979811549, 5.744144943464)
    pts(151) = Array(3.502613616870, 5.558849945287)
    pts(152) = Array(3.445270335676, 5.234583698479)
    pts(153) = Array(3.414094927672, 5.049288700303)
    pts(154) = Array(3.384073015963, 4.863993702126)
    pts(155) = Array(3.334300151524, 4.539727455318)
    pts(156) = Array(3.307433296628, 4.354432457142)
    pts(157) = Array(3.281707950223, 4.169137458966)
    pts(158) = Array(3.239426803853, 3.844871212157)
    pts(159) = Array(3.216826233566, 3.659576213981)
    pts(160) = Array(3.195357118514, 3.474281215805)
    pts(161) = Array(3.160501692116, 3.150014968996)
    pts(162) = Array(3.142132277205, 2.964719970820)
    pts(163) = Array(3.124886118492, 2.779424972644)
    pts(164) = Array(3.097402586995, 2.455158725835)
    pts(165) = Array(3.083236058623, 2.269863727659)
    pts(166) = Array(3.070186377281, 2.084568729483)
    pts(167) = Array(3.050032665942, 1.760302482674)
    pts(168) = Array(3.040047394944, 1.575007484498)
    pts(169) = Array(3.031174302080, 1.389712486322)
    pts(170) = Array(3.018319760779, 1.065446239513)
    pts(171) = Array(3.012500590970, 0.880151241337)
    pts(172) = Array(3.007790634945, 0.694856243161)
    pts(173) = Array(3.002215803511, 0.370589996352)
    pts(174) = Array(3.000553936025, 0.185294998176)
    pts(175) = Array(3.000000000000, 0.000000000000)
    Profile_9 = pts
End Function

Private Function Profile_10() As Variant
    Dim pts(122) As Variant
    pts(0) = Array(11.393697678420, 0.000000000000)
    pts(1) = Array(11.394263884011, 0.384915503450)
    pts(2) = Array(11.395741692149, 0.731339456555)
    pts(3) = Array(11.398459543796, 1.116254960005)
    pts(4) = Array(11.401873924945, 1.462678913110)
    pts(5) = Array(11.406743677443, 1.847594416560)
    pts(6) = Array(11.412094951728, 2.194018369665)
    pts(7) = Array(11.419117061713, 2.578933873115)
    pts(8) = Array(11.426405730999, 2.925357826219)
    pts(9) = Array(11.435580857157, 3.310273329669)
    pts(10) = Array(11.444807605287, 3.656697282774)
    pts(11) = Array(11.456136608668, 4.041612786224)
    pts(12) = Array(11.467302301778, 4.388036739329)
    pts(13) = Array(11.480786246200, 4.772952242779)
    pts(14) = Array(11.493891933133, 5.119376195884)
    pts(15) = Array(11.522861692616, 5.812224102094)
    pts(16) = Array(11.540545922196, 6.197139605544)
    pts(17) = Array(11.557433194094, 6.543563558649)
    pts(18) = Array(11.577276684951, 6.928479062099)
    pts(19) = Array(11.596108109956, 7.274903015204)
    pts(20) = Array(11.618112732111, 7.659818518654)
    pts(21) = Array(11.638890085921, 8.006242471759)
    pts(22) = Array(11.683212540581, 8.699090377968)
    pts(23) = Array(11.709431827633, 9.084005881418)
    pts(24) = Array(11.734004420832, 9.430429834523)
    pts(25) = Array(11.762391554058, 9.815345337973)
    pts(26) = Array(11.788916395284, 10.161769291078)
    pts(27) = Array(11.819474064708, 10.546684794528)
    pts(28) = Array(11.847953669629, 10.893108747633)
    pts(29) = Array(11.907694127206, 11.585956653843)
    pts(30) = Array(11.942487146836, 11.970872157293)
    pts(31) = Array(11.974781335714, 12.317296110398)
    pts(32) = Array(12.011754050824, 12.702211613848)
    pts(33) = Array(12.046011529447, 13.048635566953)
    pts(34) = Array(12.085167470096, 13.433551070403)
    pts(35) = Array(12.121391512860, 13.779975023507)
    pts(36) = Array(12.196638586650, 14.472822929717)
    pts(37) = Array(12.240057022498, 14.857738433167)
    pts(38) = Array(12.280120828263, 15.204162386272)
    pts(39) = Array(12.325734161960, 15.589077889722)
    pts(40) = Array(12.367775328725, 15.935501842827)
    pts(41) = Array(12.415587956984, 16.320417346277)
    pts(42) = Array(12.459610542609, 16.666841299382)
    pts(43) = Array(12.550476816265, 17.359689205592)
    pts(44) = Array(12.602585749187, 17.744604709042)
    pts(45) = Array(12.650479301627, 18.091028662147)
    pts(46) = Array(12.704801803141, 18.475944165597)
    pts(47) = Array(12.754689924314, 18.822368118702)
    pts(48) = Array(12.811231291467, 19.207283622151)
    pts(49) = Array(12.863118853855, 19.553707575256)
    pts(50) = Array(12.969741738202, 20.246555481466)
    pts(51) = Array(13.030620143059, 20.631470984916)
    pts(52) = Array(13.086416140520, 20.977894938021)
    pts(53) = Array(13.200870781444, 21.670742844231)
    pts(54) = Array(13.266108964272, 22.055658347681)
    pts(55) = Array(13.325834276298, 22.402082300786)
    pts(56) = Array(13.448163617157, 23.094930206996)
    pts(57) = Array(18.134889960676, 23.094930206996)
    pts(58) = Array(18.134889960676, 21.846075837929)
    pts(59) = Array(17.986251987077, 21.516776642851)
    pts(60) = Array(17.832803494130, 21.170352689746)
    pts(61) = Array(17.682320192866, 20.823928736641)
    pts(62) = Array(17.534774516532, 20.477504783536)
    pts(63) = Array(17.406067356626, 20.169572380776)
    pts(64) = Array(17.263998364748, 19.823148427671)
    pts(65) = Array(17.124792139733, 19.476724474566)
    pts(66) = Array(16.988424335439, 19.130300521461)
    pts(67) = Array(16.854871369109, 18.783876568357)
    pts(68) = Array(16.724110397321, 18.437452615252)
    pts(69) = Array(16.596119293081, 18.091028662147)
    pts(70) = Array(16.470876623976, 17.744604709042)
    pts(71) = Array(16.348361631352, 17.398180755937)
    pts(72) = Array(16.228554210455, 17.051756802832)
    pts(73) = Array(16.111434891484, 16.705332849727)
    pts(74) = Array(15.984432116940, 16.320417346277)
    pts(75) = Array(15.872926496469, 15.973993393172)
    pts(76) = Array(15.764052273482, 15.627569440067)
    pts(77) = Array(15.657792407357, 15.281145486962)
    pts(78) = Array(15.542772072428, 14.896229983512)
    pts(79) = Array(15.441977904390, 14.549806030407)
    pts(80) = Array(15.343748536842, 14.203382077302)
    pts(81) = Array(15.248069105115, 13.856958124197)
    pts(82) = Array(15.144731831156, 13.472042620748)
    pts(83) = Array(15.054389000288, 13.125618667643)
    pts(84) = Array(14.966552979756, 12.779194714538)
    pts(85) = Array(14.881210871269, 12.432770761433)
    pts(86) = Array(14.789296140249, 12.047855257983)
    pts(87) = Array(14.709178616515, 11.701431304878)
    pts(88) = Array(14.623039884664, 11.316515801428)
    pts(89) = Array(14.548095559862, 10.970091848323)
    pts(90) = Array(14.467677935959, 10.585176344873)
    pts(91) = Array(14.405497659156, 10.277243942113)
    pts(92) = Array(14.337823996099, 9.930819989008)
    pts(93) = Array(14.279688188655, 9.622887586248)
    pts(94) = Array(14.216548014350, 9.276463633143)
    pts(95) = Array(14.162427830615, 8.968531230383)
    pts(96) = Array(14.103789962710, 8.622107277278)
    pts(97) = Array(14.053658795545, 8.314174874518)
    pts(98) = Array(13.999494530822, 7.967750921414)
    pts(99) = Array(13.953327943545, 7.659818518654)
    pts(100) = Array(13.903610985014, 7.313394565549)
    pts(101) = Array(13.861386649750, 7.005462162789)
    pts(102) = Array(13.816093040742, 6.659038209684)
    pts(103) = Array(13.777790682815, 6.351105806924)
    pts(104) = Array(13.736898747673, 6.004681853819)
    pts(105) = Array(13.698329097645, 5.658257900714)
    pts(106) = Array(13.662076607870, 5.311833947609)
    pts(107) = Array(13.631793552687, 5.003901544849)
    pts(108) = Array(13.599905073517, 4.657477591744)
    pts(109) = Array(13.573494230980, 4.349545188984)
    pts(110) = Array(13.545954759606, 4.003121235879)
    pts(111) = Array(13.523403731589, 3.695188833119)
    pts(112) = Array(13.500200366611, 3.348764880014)
    pts(113) = Array(13.479288072280, 3.002340926909)
    pts(114) = Array(13.460664117758, 2.655916973805)
    pts(115) = Array(13.446028597029, 2.347984571045)
    pts(116) = Array(13.431720679892, 2.001560617940)
    pts(117) = Array(13.420918456761, 1.693628215180)
    pts(118) = Array(13.410919935530, 1.347204262075)
    pts(119) = Array(13.403946000580, 1.039271859315)
    pts(120) = Array(13.398252242209, 0.692847906210)
    pts(121) = Array(13.394836282465, 0.346423953105)
    pts(122) = Array(13.393697678420, 0.000000000000)
    Profile_10 = pts
End Function

Private Function Profile_11() As Variant
    Dim pts(78) As Variant
    pts(0) = Array(19.659197820326, 0.000000000000)
    pts(1) = Array(19.658337338352, 0.669000000000)
    pts(2) = Array(19.655755875345, 1.338000000000)
    pts(3) = Array(19.651453380057, 2.007000000000)
    pts(4) = Array(19.645429767065, 2.676000000000)
    pts(5) = Array(19.637684916764, 3.345000000000)
    pts(6) = Array(19.628218675360, 4.014000000000)
    pts(7) = Array(19.617030854847, 4.683000000000)
    pts(8) = Array(19.604883609130, 5.314833333333)
    pts(9) = Array(19.590347606650, 5.983833333333)
    pts(10) = Array(19.574089272648, 6.652833333333)
    pts(11) = Array(19.556108283976, 7.321833333333)
    pts(12) = Array(19.537544158849, 7.953666666667)
    pts(13) = Array(19.516212509545, 8.622666666667)
    pts(14) = Array(19.493157054555, 9.291666666667)
    pts(15) = Array(19.468377335081, 9.960666666667)
    pts(16) = Array(19.443390585887, 10.592500000000)
    pts(17) = Array(19.418467890393, 11.187166666667)
    pts(18) = Array(19.392181638325, 11.781833333333)
    pts(19) = Array(19.364531415745, 12.376500000000)
    pts(20) = Array(19.335516787057, 12.971166666667)
    pts(21) = Array(19.305137294969, 13.565833333333)
    pts(22) = Array(19.273392460461, 14.160500000000)
    pts(23) = Array(19.240281782743, 14.755166666667)
    pts(24) = Array(19.205804739220, 15.349833333333)
    pts(25) = Array(19.156166781337, 16.167500000000)
    pts(26) = Array(19.106372948637, 16.948000000000)
    pts(27) = Array(19.051679470232, 17.765666666667)
    pts(28) = Array(18.997057000119, 18.546166666667)
    pts(29) = Array(18.937301748902, 19.363833333333)
    pts(30) = Array(18.877844391940, 20.144333333333)
    pts(31) = Array(18.816023369568, 20.924833333333)
    pts(32) = Array(18.751313851654, 21.711578595572)
    pts(33) = Array(18.751313851654, 22.300000000000)
    pts(34) = Array(21.398871715914, 22.300000000000)
    pts(35) = Array(21.398871715914, 21.574226291719)
    pts(36) = Array(21.462307749995, 21.222166666667)
    pts(37) = Array(21.521582258341, 20.887666666667)
    pts(38) = Array(21.579886223155, 20.553166666667)
    pts(39) = Array(21.643531849211, 20.181500000000)
    pts(40) = Array(21.755085045773, 19.512500000000)
    pts(41) = Array(21.809413703407, 19.178000000000)
    pts(42) = Array(21.868648873705, 18.806333333333)
    pts(43) = Array(21.972280175677, 18.137333333333)
    pts(44) = Array(22.022655829147, 17.802833333333)
    pts(45) = Array(22.077504814599, 17.431166666667)
    pts(46) = Array(22.173256567801, 16.762166666667)
    pts(47) = Array(22.219699748496, 16.427666666667)
    pts(48) = Array(22.270184864273, 16.056000000000)
    pts(49) = Array(22.358095926912, 15.387000000000)
    pts(50) = Array(22.400625437691, 15.052500000000)
    pts(51) = Array(22.446767089932, 14.680833333333)
    pts(52) = Array(22.487293633687, 14.346333333333)
    pts(53) = Array(22.531212179276, 13.974666666667)
    pts(54) = Array(22.569740043090, 13.640166666667)
    pts(55) = Array(22.611440121387, 13.268500000000)
    pts(56) = Array(22.683563154840, 12.599500000000)
    pts(57) = Array(22.718210316317, 12.265000000000)
    pts(58) = Array(22.755602607627, 11.893333333333)
    pts(59) = Array(22.788262437022, 11.558833333333)
    pts(60) = Array(22.823448510991, 11.187166666667)
    pts(61) = Array(22.854124428255, 10.852666666667)
    pts(62) = Array(22.887107965206, 10.481000000000)
    pts(63) = Array(22.940523246520, 9.849166666667)
    pts(64) = Array(22.993438362592, 9.180166666667)
    pts(65) = Array(23.039978224781, 8.548333333333)
    pts(66) = Array(23.085622494287, 7.879333333333)
    pts(67) = Array(23.125303342041, 7.247500000000)
    pts(68) = Array(23.163692829354, 6.578500000000)
    pts(69) = Array(23.198356694893, 5.909500000000)
    pts(70) = Array(23.229298152401, 5.240500000000)
    pts(71) = Array(23.255105274988, 4.608666666667)
    pts(72) = Array(23.278816607200, 3.939666666667)
    pts(73) = Array(23.297799494781, 3.307833333333)
    pts(74) = Array(23.314288990451, 2.638833333333)
    pts(75) = Array(23.326454274163, 2.007000000000)
    pts(76) = Array(23.335727690043, 1.338000000000)
    pts(77) = Array(23.341291397489, 0.669000000000)
    pts(78) = Array(23.343145909631, 0.000000000000)
    Profile_11 = pts
End Function

Private Function Profile_12() As Variant
    Dim pts(82) As Variant
    pts(0) = Array(36.741147574361, 0.000000000000)
    pts(1) = Array(36.742553529618, 0.574179739051)
    pts(2) = Array(36.746771496552, 1.148359478102)
    pts(3) = Array(36.753801778685, 1.722539217152)
    pts(4) = Array(36.763644882010, 2.296718956203)
    pts(5) = Array(36.776301515173, 2.870898695254)
    pts(6) = Array(36.791772589726, 3.445078434305)
    pts(7) = Array(36.810059220457, 4.019258173356)
    pts(8) = Array(36.829981764427, 4.563217926141)
    pts(9) = Array(36.853755288949, 5.137397665192)
    pts(10) = Array(36.880348841802, 5.711577404242)
    pts(11) = Array(36.909764344660, 6.285757143293)
    pts(12) = Array(36.940236661856, 6.829716896078)
    pts(13) = Array(36.975153834235, 7.403896635129)
    pts(14) = Array(37.012899821530, 7.978076374180)
    pts(15) = Array(37.053477363157, 8.552256113231)
    pts(16) = Array(37.094533847002, 9.096215866016)
    pts(17) = Array(37.140634121563, 9.670395605067)
    pts(18) = Array(37.189575247062, 10.244575344117)
    pts(19) = Array(37.241360796258, 10.818755083168)
    pts(20) = Array(37.293048023165, 11.362714835953)
    pts(21) = Array(37.350383776347, 11.936894575004)
    pts(22) = Array(37.410575734701, 12.511074314055)
    pts(23) = Array(37.473628324865, 13.085254053106)
    pts(24) = Array(37.536005320772, 13.629213805891)
    pts(25) = Array(37.604642151593, 14.203393544942)
    pts(26) = Array(37.676153948704, 14.777573283992)
    pts(27) = Array(37.750546018381, 15.351753023043)
    pts(28) = Array(37.823684600253, 15.895712775828)
    pts(29) = Array(37.903701711025, 16.469892514879)
    pts(30) = Array(37.982179822381, 17.013852267664)
    pts(31) = Array(38.067844665979, 17.588032006715)
    pts(32) = Array(38.151684622980, 18.131991759500)
    pts(33) = Array(38.938331521198, 18.131991759500)
    pts(34) = Array(38.938331521198, 17.951778482436)
    pts(35) = Array(38.998589377732, 17.648471979246)
    pts(36) = Array(39.063406604421, 17.316052130322)
    pts(37) = Array(39.121222309514, 17.013852267664)
    pts(38) = Array(39.183602147104, 16.681432418740)
    pts(39) = Array(39.239206493577, 16.379232556081)
    pts(40) = Array(39.299158622394, 16.046812707157)
    pts(41) = Array(39.357842814595, 15.714392858233)
    pts(42) = Array(39.415261549247, 15.381973009309)
    pts(43) = Array(39.466364311570, 15.079773146651)
    pts(44) = Array(39.521373841131, 14.747353297726)
    pts(45) = Array(39.570290285748, 14.445153435068)
    pts(46) = Array(39.622899018139, 14.112733586144)
    pts(47) = Array(39.669636604452, 13.810533723486)
    pts(48) = Array(39.719852544083, 13.478113874562)
    pts(49) = Array(39.768818176975, 13.145694025637)
    pts(50) = Array(39.816535533831, 12.813274176713)
    pts(51) = Array(39.871322122877, 12.420414355257)
    pts(52) = Array(39.924371167016, 12.027554533802)
    pts(53) = Array(39.975685718278, 11.634694712346)
    pts(54) = Array(40.025268722107, 11.241834890890)
    pts(55) = Array(40.073123018222, 10.848975069434)
    pts(56) = Array(40.119251341451, 10.456115247978)
    pts(57) = Array(40.163656322528, 10.063255426522)
    pts(58) = Array(40.209552680120, 9.640175618801)
    pts(59) = Array(40.250386370161, 9.247315797345)
    pts(60) = Array(40.289504167633, 8.854455975889)
    pts(61) = Array(40.326908286444, 8.461596154433)
    pts(62) = Array(40.365275586150, 8.038516346712)
    pts(63) = Array(40.399127164325, 7.645656525256)
    pts(64) = Array(40.431271247794, 7.252796703800)
    pts(65) = Array(40.461709644197, 6.859936882344)
    pts(66) = Array(40.492583870704, 6.436857074622)
    pts(67) = Array(40.519485036014, 6.043997253167)
    pts(68) = Array(40.544685461919, 5.651137431711)
    pts(69) = Array(40.568186558325, 5.258277610255)
    pts(70) = Array(40.591596502200, 4.835197802533)
    pts(71) = Array(40.611572308475, 4.442337981077)
    pts(72) = Array(40.629852519864, 4.049478159622)
    pts(73) = Array(40.646438155136, 3.656618338166)
    pts(74) = Array(40.662405548153, 3.233538530444)
    pts(75) = Array(40.675474521566, 2.840678708988)
    pts(76) = Array(40.686851457847, 2.447818887532)
    pts(77) = Array(40.696536989361, 2.054959066077)
    pts(78) = Array(40.705076605891, 1.631879258355)
    pts(79) = Array(40.711250831256, 1.239019436899)
    pts(80) = Array(40.716009949285, 0.815939629177)
    pts(81) = Array(40.718674347222, 0.423079807722)
    pts(82) = Array(40.719654106282, 0.000000000000)
    Profile_12 = pts
End Function

Private Function Profile_13() As Variant
    Dim pts(66) As Variant
    pts(0) = Array(43.040957210658, 0.000000000000)
    pts(1) = Array(43.041952900642, 0.226516005955)
    pts(2) = Array(43.044940201490, 0.453032011909)
    pts(3) = Array(43.049919806155, 0.679548017864)
    pts(4) = Array(43.056892870451, 0.906064023818)
    pts(5) = Array(43.065861014398, 1.132580029773)
    pts(6) = Array(43.076826324105, 1.359096035727)
    pts(7) = Array(43.089791354203, 1.585612041682)
    pts(8) = Array(43.103172133530, 1.789476447041)
    pts(9) = Array(43.119945361931, 2.015992452996)
    pts(10) = Array(43.138728398692, 2.242508458950)
    pts(11) = Array(43.159525659446, 2.469024464905)
    pts(12) = Array(43.182342048442, 2.695540470859)
    pts(13) = Array(43.207182964401, 2.922056476814)
    pts(14) = Array(43.234054306986, 3.148572482769)
    pts(15) = Array(43.262962483924, 3.375088488723)
    pts(16) = Array(43.290727048080, 3.578952894082)
    pts(17) = Array(43.323524717894, 3.805468900037)
    pts(18) = Array(43.358380752489, 4.031984905992)
    pts(19) = Array(43.395303621868, 4.258500911946)
    pts(20) = Array(43.430308793956, 4.462365317305)
    pts(21) = Array(43.471183974051, 4.688881323260)
    pts(22) = Array(43.514153720718, 4.915397329214)
    pts(23) = Array(43.559228728217, 5.141913335169)
    pts(24) = Array(43.601605556937, 5.345777740528)
    pts(25) = Array(43.650712191975, 5.572293746483)
    pts(26) = Array(43.696737267466, 5.776158151842)
    pts(27) = Array(43.749920606448, 6.002674157796)
    pts(28) = Array(43.799636497385, 6.206538563155)
    pts(29) = Array(43.856946120232, 6.433054569110)
    pts(30) = Array(43.910399556468, 6.636918974469)
    pts(31) = Array(43.971889794131, 6.863434980424)
    pts(32) = Array(44.029131903051, 7.067299385783)
    pts(33) = Array(44.094862116069, 7.293815391737)
    pts(34) = Array(44.155948688425, 7.497679797097)
    pts(35) = Array(44.225983580097, 7.724195803051)
    pts(36) = Array(44.290975368803, 7.928060208410)
    pts(37) = Array(44.365385337130, 8.154576214365)
    pts(38) = Array(44.434348393036, 8.358440619724)
    pts(39) = Array(44.505217247145, 8.562305025083)
    pts(40) = Array(44.578008622430, 8.766169430442)
    pts(41) = Array(44.661163900060, 8.992685436397)
    pts(42) = Array(44.738071839580, 9.196549841756)
    pts(43) = Array(44.816958644440, 9.400414247115)
    pts(44) = Array(44.897843901858, 9.604278652474)
    pts(45) = Array(44.980747964619, 9.808143057833)
    pts(46) = Array(45.065691981343, 10.012007463192)
    pts(47) = Array(45.152697928572, 10.215871868552)
    pts(48) = Array(45.241788644823, 10.419736273911)
    pts(49) = Array(45.332987866745, 10.623600679270)
    pts(50) = Array(45.426320267511, 10.827465084629)
    pts(51) = Array(45.521811497642, 11.031329489988)
    pts(52) = Array(45.619488228417, 11.235193895347)
    pts(53) = Array(45.719378198085, 11.439058300706)
    pts(54) = Array(45.821510261079, 11.642922706065)
    pts(55) = Array(45.925914440480, 11.846787111425)
    pts(56) = Array(46.020650843776, 12.027999916188)
    pts(57) = Array(46.129433091560, 12.231864321547)
    pts(58) = Array(46.240581282654, 12.435728726906)
    pts(59) = Array(46.354130552938, 12.639593132266)
    pts(60) = Array(46.457108525602, 12.820805937029)
    pts(61) = Array(46.571029354160, 13.017387667291)
    pts(62) = Array(46.682455265300, 13.205883147152)
    pts(63) = Array(46.805419077654, 13.409747552511)
    pts(64) = Array(46.916893872631, 13.590960357275)
    pts(65) = Array(47.540065789245, 13.590960357275)
    pts(66) = Array(47.540957210658, 0.000000000000)
    Profile_13 = pts
End Function

Private Function Profile_14() As Variant
    Dim pts(67) As Variant
    pts(0) = Array(47.540957210658, 0.000000000000)
    pts(1) = Array(47.540065789245, 13.590960357275)
    pts(2) = Array(53.743220456623, 13.590960357275)
    pts(3) = Array(53.743220456623, 9.596625723305)
    pts(4) = Array(53.606153546424, 9.468369048901)
    pts(5) = Array(53.488332311260, 9.355111045924)
    pts(6) = Array(53.373407167350, 9.241853042947)
    pts(7) = Array(53.261271130221, 9.128595039970)
    pts(8) = Array(53.151824833270, 9.015337036992)
    pts(9) = Array(53.023909966103, 8.879427433420)
    pts(10) = Array(52.920064309850, 8.766169430442)
    pts(11) = Array(52.798630655180, 8.630259826869)
    pts(12) = Array(52.700000127300, 8.517001823892)
    pts(13) = Array(52.584617430045, 8.381092220319)
    pts(14) = Array(52.472374366645, 8.245182616747)
    pts(15) = Array(52.363167605952, 8.109273013174)
    pts(16) = Array(52.256900920367, 7.973363409601)
    pts(17) = Array(52.153484538614, 7.837453806028)
    pts(18) = Array(52.052834573494, 7.701544202456)
    pts(19) = Array(51.954872514267, 7.565634598883)
    pts(20) = Array(51.859524774952, 7.429724995310)
    pts(21) = Array(51.766722291213, 7.293815391737)
    pts(22) = Array(51.676400159586, 7.157905788165)
    pts(23) = Array(51.574077959571, 6.999344583997)
    pts(24) = Array(51.488925148583, 6.863434980424)
    pts(25) = Array(51.406071128608, 6.727525376851)
    pts(26) = Array(51.325465194745, 6.591615773278)
    pts(27) = Array(51.234202502177, 6.433054569110)
    pts(28) = Array(51.158306537141, 6.297144965537)
    pts(29) = Array(51.072418979229, 6.138583761369)
    pts(30) = Array(51.001033322895, 6.002674157796)
    pts(31) = Array(50.920301266464, 5.844112953628)
    pts(32) = Array(50.853248299714, 5.708203350055)
    pts(33) = Array(50.777476079921, 5.549642145887)
    pts(34) = Array(50.704300940548, 5.391080941719)
    pts(35) = Array(50.633675215320, 5.232519737551)
    pts(36) = Array(50.565553998786, 5.073958533383)
    pts(37) = Array(50.499894971838, 4.915397329214)
    pts(38) = Array(50.436658242763, 4.756836125046)
    pts(39) = Array(50.375806202256, 4.598274920878)
    pts(40) = Array(50.317303390969, 4.439713716710)
    pts(41) = Array(50.261116378396, 4.281152512542)
    pts(42) = Array(50.207213651982, 4.122591308373)
    pts(43) = Array(50.155565515516, 3.964030104205)
    pts(44) = Array(50.106143995941, 3.805468900037)
    pts(45) = Array(50.058922757843, 3.646907695869)
    pts(46) = Array(50.013877024926, 3.488346491701)
    pts(47) = Array(49.970983507899, 3.329785287532)
    pts(48) = Array(49.930220338219, 3.171224083364)
    pts(49) = Array(49.891567007233, 3.012662879196)
    pts(50) = Array(49.855004310276, 2.854101675028)
    pts(51) = Array(49.815755357724, 2.672888870264)
    pts(52) = Array(49.783613684799, 2.514327666096)
    pts(53) = Array(49.753510229891, 2.355766461928)
    pts(54) = Array(49.725430689496, 2.197205257759)
    pts(55) = Array(49.695801136733, 2.015992452996)
    pts(56) = Array(49.672015379950, 1.857431248828)
    pts(57) = Array(49.647262724862, 1.676218444064)
    pts(58) = Array(49.627720168442, 1.517657239896)
    pts(59) = Array(49.607792166079, 1.336444435132)
    pts(60) = Array(49.592451843688, 1.177883230964)
    pts(61) = Array(49.579061140431, 1.019322026796)
    pts(62) = Array(49.567614044085, 0.860760822627)
    pts(63) = Array(49.556905047607, 0.679548017864)
    pts(64) = Array(49.549606711755, 0.520986813696)
    pts(65) = Array(49.543629446558, 0.339774008932)
    pts(66) = Array(49.540464964020, 0.181212804764)
    pts(67) = Array(49.539207246133, 0.000000000000)
    Profile_14 = pts
End Function

Private Function Profile_15() As Variant
    Dim pts(94) As Variant
    pts(0) = Array(74.785052642083, 0.000000000000)
    pts(1) = Array(74.784413218444, 0.192606345594)
    pts(2) = Array(74.782804606023, 0.361136897988)
    pts(3) = Array(74.779766982770, 0.553743243582)
    pts(4) = Array(74.776059452242, 0.722273795976)
    pts(5) = Array(74.770622216011, 0.914880141570)
    pts(6) = Array(74.764814041511, 1.083410693965)
    pts(7) = Array(74.750253501285, 1.420471798754)
    pts(8) = Array(74.740168839343, 1.613078144347)
    pts(9) = Array(74.730290871031, 1.781608696742)
    pts(10) = Array(74.717795968969, 1.974215042335)
    pts(11) = Array(74.705806605508, 2.142745594730)
    pts(12) = Array(74.690895603481, 2.335351940324)
    pts(13) = Array(74.676789212590, 2.503882492718)
    pts(14) = Array(74.645604397066, 2.840943597507)
    pts(15) = Array(74.626001333024, 3.033549943101)
    pts(16) = Array(74.607782245225, 3.202080495495)
    pts(17) = Array(74.568350211314, 3.539141600284)
    pts(18) = Array(74.544020280835, 3.731747945878)
    pts(19) = Array(74.521656063870, 3.900278498272)
    pts(20) = Array(74.473906335837, 4.237339603061)
    pts(21) = Array(74.444805861541, 4.429945948655)
    pts(22) = Array(74.418256236197, 4.598476501050)
    pts(23) = Array(74.390689394230, 4.767007053444)
    pts(24) = Array(74.362102392758, 4.935537605839)
    pts(25) = Array(74.328178423193, 5.128143951432)
    pts(26) = Array(74.297394891414, 5.296674503827)
    pts(27) = Array(74.265581155895, 5.465205056221)
    pts(28) = Array(74.232733761380, 5.633735608616)
    pts(29) = Array(74.193923573839, 5.826341954209)
    pts(30) = Array(74.158848960451, 5.994872506604)
    pts(31) = Array(74.122728982487, 6.163403058998)
    pts(32) = Array(74.085559638567, 6.331933611393)
    pts(33) = Array(74.047336789711, 6.500464163787)
    pts(34) = Array(74.008056156897, 6.668994716182)
    pts(35) = Array(73.967713318501, 6.837525268576)
    pts(36) = Array(73.926303707644, 7.006055820971)
    pts(37) = Array(73.877666130927, 7.198662166564)
    pts(38) = Array(73.833954508607, 7.367192718959)
    pts(39) = Array(73.789160783476, 7.535723271353)
    pts(40) = Array(73.743279756014, 7.704253823748)
    pts(41) = Array(73.696306065899, 7.872784376142)
    pts(42) = Array(73.648234188582, 8.041314928537)
    pts(43) = Array(73.599058431730, 8.209845480931)
    pts(44) = Array(73.548772931514, 8.378376033326)
    pts(45) = Array(73.497371648750, 8.546906585720)
    pts(46) = Array(73.444848364872, 8.715437138115)
    pts(47) = Array(73.391196677745, 8.883967690509)
    pts(48) = Array(73.336409997302, 9.052498242904)
    pts(49) = Array(73.280481540998, 9.221028795298)
    pts(50) = Array(73.223404329080, 9.389559347693)
    pts(51) = Array(73.165171179645, 9.558089900087)
    pts(52) = Array(73.105774703513, 9.726620452482)
    pts(53) = Array(73.045207298859, 9.895151004876)
    pts(54) = Array(72.983461145640, 10.063681557271)
    pts(55) = Array(72.865634819049, 10.376666868861)
    pts(56) = Array(72.800475698696, 10.545197421255)
    pts(57) = Array(72.734105512857, 10.713727973649)
    pts(58) = Array(72.666515294072, 10.882258526044)
    pts(59) = Array(72.605226962642, 11.032492589555)
    pts(60) = Array(72.605226962642, 14.445475919527)
    pts(61) = Array(79.109023162909, 14.445475919527)
    pts(62) = Array(79.109023162909, 13.856712874950)
    pts(63) = Array(79.006691605088, 13.410216811961)
    pts(64) = Array(78.910734585160, 12.976852534375)
    pts(65) = Array(78.818076092559, 12.543488256790)
    pts(66) = Array(78.723831472483, 12.086048186005)
    pts(67) = Array(78.637905845421, 11.652683908419)
    pts(68) = Array(78.550736121698, 11.195243837634)
    pts(69) = Array(78.471484016159, 10.761879560048)
    pts(70) = Array(78.391329332612, 10.304439489263)
    pts(71) = Array(78.318696630406, 9.871075211677)
    pts(72) = Array(78.245502542362, 9.413635140892)
    pts(73) = Array(78.175863362998, 8.956195070107)
    pts(74) = Array(78.109766803163, 8.498754999322)
    pts(75) = Array(78.050406324266, 8.065390721736)
    pts(76) = Array(77.991175831273, 7.607950650951)
    pts(77) = Array(77.935455639349, 7.150510580166)
    pts(78) = Array(77.883236075255, 6.693070509381)
    pts(79) = Array(77.839473196573, 6.283782024994)
    pts(80) = Array(77.798499478265, 5.874493540608)
    pts(81) = Array(77.760309279869, 5.465205056221)
    pts(82) = Array(77.722900718640, 5.031840778635)
    pts(83) = Array(77.690425222516, 4.622552294249)
    pts(84) = Array(77.659057118283, 4.189188016663)
    pts(85) = Array(77.632277510772, 3.779899532276)
    pts(86) = Array(77.606931729513, 3.346535254691)
    pts(87) = Array(77.585832558776, 2.937246770304)
    pts(88) = Array(77.567487875159, 2.527958285917)
    pts(89) = Array(77.551895193240, 2.118669801531)
    pts(90) = Array(77.538382542313, 1.685305523945)
    pts(91) = Array(77.528449516981, 1.276017039558)
    pts(92) = Array(77.520926005972, 0.842652761972)
    pts(93) = Array(77.516646947953, 0.433364277586)
    pts(94) = Array(77.515108270133, 0.000000000000)
    Profile_15 = pts
End Function

Private Function Profile_16() As Variant
    Dim pts(97) As Variant
    pts(0) = Array(77.515108270133, 0.000000000000)
    pts(1) = Array(77.516646947953, 0.433364277586)
    pts(2) = Array(77.520926005972, 0.842652761972)
    pts(3) = Array(77.528449516981, 1.276017039558)
    pts(4) = Array(77.538382542313, 1.685305523945)
    pts(5) = Array(77.551895193240, 2.118669801531)
    pts(6) = Array(77.567487875159, 2.527958285917)
    pts(7) = Array(77.585832558776, 2.937246770304)
    pts(8) = Array(77.606931729513, 3.346535254691)
    pts(9) = Array(77.632277510772, 3.779899532276)
    pts(10) = Array(77.659057118283, 4.189188016663)
    pts(11) = Array(77.690425222516, 4.622552294249)
    pts(12) = Array(77.722900718640, 5.031840778635)
    pts(13) = Array(77.760309279869, 5.465205056221)
    pts(14) = Array(77.798499478265, 5.874493540608)
    pts(15) = Array(77.839473196573, 6.283782024994)
    pts(16) = Array(77.883236075255, 6.693070509381)
    pts(17) = Array(77.935455639349, 7.150510580166)
    pts(18) = Array(77.991175831273, 7.607950650951)
    pts(19) = Array(78.050406324266, 8.065390721736)
    pts(20) = Array(78.109766803163, 8.498754999322)
    pts(21) = Array(78.175863362998, 8.956195070107)
    pts(22) = Array(78.245502542362, 9.413635140892)
    pts(23) = Array(78.318696630406, 9.871075211677)
    pts(24) = Array(78.391329332612, 10.304439489263)
    pts(25) = Array(78.471484016159, 10.761879560048)
    pts(26) = Array(78.550736121698, 11.195243837634)
    pts(27) = Array(78.637905845421, 11.652683908419)
    pts(28) = Array(78.723831472483, 12.086048186005)
    pts(29) = Array(78.818076092559, 12.543488256790)
    pts(30) = Array(78.910734585160, 12.976852534375)
    pts(31) = Array(79.006691605088, 13.410216811961)
    pts(32) = Array(79.109023162909, 13.856712874950)
    pts(33) = Array(79.109023162909, 14.445475919527)
    pts(34) = Array(79.518859940962, 14.445475919527)
    pts(35) = Array(79.705560767098, 14.204717987535)
    pts(36) = Array(79.887311199322, 13.963960055543)
    pts(37) = Array(80.064253866219, 13.723202123551)
    pts(38) = Array(80.150964380031, 13.602823157555)
    pts(39) = Array(80.253497078648, 13.458368398360)
    pts(40) = Array(80.420768607938, 13.217610466368)
    pts(41) = Array(80.502740003563, 13.097231500372)
    pts(42) = Array(80.599667015454, 12.952776741176)
    pts(43) = Array(80.679256498475, 12.832397775180)
    pts(44) = Array(80.773363811910, 12.687943015985)
    pts(45) = Array(80.850634705393, 12.567564049989)
    pts(46) = Array(80.941995874561, 12.423109290794)
    pts(47) = Array(81.091010870805, 12.182351358801)
    pts(48) = Array(81.164017085119, 12.061972392805)
    pts(49) = Array(81.250323610755, 11.917517633610)
    pts(50) = Array(81.391060927255, 11.676759701618)
    pts(51) = Array(81.541469655320, 11.411925976427)
    pts(52) = Array(81.687364568937, 11.147092251235)
    pts(53) = Array(81.828841158191, 10.882258526044)
    pts(54) = Array(81.965989465497, 10.617424800853)
    pts(55) = Array(82.098894457454, 10.352591075661)
    pts(56) = Array(82.227636362685, 10.087757350470)
    pts(57) = Array(82.352290979314, 9.822923625279)
    pts(58) = Array(82.472929955303, 9.558089900087)
    pts(59) = Array(82.537064872237, 9.413635140892)
    pts(60) = Array(82.600035805066, 9.269180381697)
    pts(61) = Array(82.712493086467, 9.004346656505)
    pts(62) = Array(82.821132572289, 8.739512931314)
    pts(63) = Array(82.926011375322, 8.474679206123)
    pts(64) = Array(82.981652366431, 8.330224446927)
    pts(65) = Array(83.036199017366, 8.185769687732)
    pts(66) = Array(83.124700284437, 7.945011755740)
    pts(67) = Array(83.210220577409, 7.704253823748)
    pts(68) = Array(83.292794419364, 7.463495891756)
    pts(69) = Array(83.380261627977, 7.198662166564)
    pts(70) = Array(83.456752852753, 6.957904234572)
    pts(71) = Array(83.530394210589, 6.717146302580)
    pts(72) = Array(83.601213731540, 6.476388370588)
    pts(73) = Array(83.675887701339, 6.211554645397)
    pts(74) = Array(83.740866354320, 5.970796713405)
    pts(75) = Array(83.809174748815, 5.705962988213)
    pts(76) = Array(83.868417570172, 5.465205056221)
    pts(77) = Array(83.930469512109, 5.200371331030)
    pts(78) = Array(83.984070081826, 4.959613399038)
    pts(79) = Array(84.039962546817, 4.694779673846)
    pts(80) = Array(84.092663486665, 4.429945948655)
    pts(81) = Array(84.142194837415, 4.165112223464)
    pts(82) = Array(84.184490090795, 3.924354291472)
    pts(83) = Array(84.228025921953, 3.659520566280)
    pts(84) = Array(84.268447592932, 3.394686841089)
    pts(85) = Array(84.305771404928, 3.129853115898)
    pts(86) = Array(84.337026524583, 2.889095183905)
    pts(87) = Array(84.368476648174, 2.624261458714)
    pts(88) = Array(84.396868766091, 2.359427733523)
    pts(89) = Array(84.422214065568, 2.094594008331)
    pts(90) = Array(84.442619701877, 1.853836076339)
    pts(91) = Array(84.462174919062, 1.589002351148)
    pts(92) = Array(84.478708843793, 1.324168625957)
    pts(93) = Array(84.492227893355, 1.059334900765)
    pts(94) = Array(84.502737301376, 0.794501175574)
    pts(95) = Array(84.510241127993, 0.529667450383)
    pts(96) = Array(84.514742267719, 0.264833725191)
    pts(97) = Array(84.516242455032, 0.000000000000)
    Profile_16 = pts
End Function

Private Function Profile_17() As Variant
    Dim pts(76) As Variant
    pts(0) = Array(84.616242455032, 0.000000000000)
    pts(1) = Array(84.617948382334, 8.136627635717)
    pts(2) = Array(84.623573372765, 17.498123947777)
    pts(3) = Array(86.532120450268, 17.498123947777)
    pts(4) = Array(86.682790037672, 17.206488548648)
    pts(5) = Array(86.830272760860, 16.914853149518)
    pts(6) = Array(86.974610147036, 16.623217750389)
    pts(7) = Array(87.129796037449, 16.302418811346)
    pts(8) = Array(87.267656437444, 16.010783412216)
    pts(9) = Array(87.402490490434, 15.719148013087)
    pts(10) = Array(87.534333665132, 15.427512613957)
    pts(11) = Array(87.675947410388, 15.106713674915)
    pts(12) = Array(87.801619382238, 14.815078275785)
    pts(13) = Array(87.936522842137, 14.494279336742)
    pts(14) = Array(88.056162794376, 14.202643937613)
    pts(15) = Array(88.184503152745, 13.881844998570)
    pts(16) = Array(88.298239724819, 13.590209599441)
    pts(17) = Array(88.420153035102, 13.269410660398)
    pts(18) = Array(88.538750613169, 12.948611721355)
    pts(19) = Array(88.654065565750, 12.627812782313)
    pts(20) = Array(88.715592459168, 12.452831542835)
    pts(21) = Array(88.786158088378, 12.248686763444)
    pts(22) = Array(88.845606416332, 12.073705523966)
    pts(23) = Array(88.913760228394, 11.869560744576)
    pts(24) = Array(88.971152310808, 11.694579505098)
    pts(25) = Array(89.036919712472, 11.490434725707)
    pts(26) = Array(89.092276795066, 11.315453486229)
    pts(27) = Array(89.155681966217, 11.111308706839)
    pts(28) = Array(89.209024264946, 10.936327467361)
    pts(29) = Array(89.270090208651, 10.732182687970)
    pts(30) = Array(89.321436951122, 10.557201448492)
    pts(31) = Array(89.380185540276, 10.353056669102)
    pts(32) = Array(89.429555004351, 10.178075429624)
    pts(33) = Array(89.486007025162, 9.973930650233)
    pts(34) = Array(89.541230673714, 9.769785870842)
    pts(35) = Array(89.595231514097, 9.565641091452)
    pts(36) = Array(89.640548797068, 9.390659851974)
    pts(37) = Array(89.692292941066, 9.186515072583)
    pts(38) = Array(89.742829326568, 8.982370293192)
    pts(39) = Array(89.792162933134, 8.778225513802)
    pts(40) = Array(89.880607829403, 8.399099494933)
    pts(41) = Array(89.926531081036, 8.194954715542)
    pts(42) = Array(89.971269425724, 7.990809936152)
    pts(43) = Array(90.051225858390, 7.611683917283)
    pts(44) = Array(90.092601988543, 7.407539137892)
    pts(45) = Array(90.132809064986, 7.203394358502)
    pts(46) = Array(90.204390633969, 6.824268339633)
    pts(47) = Array(90.241278044101, 6.620123560242)
    pts(48) = Array(90.277010354182, 6.415978780852)
    pts(49) = Array(90.340317046637, 6.036852761983)
    pts(50) = Array(90.372766961908, 5.832707982592)
    pts(51) = Array(90.404073933272, 5.628563203202)
    pts(52) = Array(90.434240829891, 5.424418423811)
    pts(53) = Array(90.463270407363, 5.220273644420)
    pts(54) = Array(90.491165309023, 5.016128865030)
    pts(55) = Array(90.517928067171, 4.811984085639)
    pts(56) = Array(90.564634869909, 4.432858066770)
    pts(57) = Array(90.588175910370, 4.228713287380)
    pts(58) = Array(90.610593555499, 4.024568507989)
    pts(59) = Array(90.631889814034, 3.820423728598)
    pts(60) = Array(90.652066589800, 3.616278949207)
    pts(61) = Array(90.671125682565, 3.412134169817)
    pts(62) = Array(90.689068788863, 3.207989390426)
    pts(63) = Array(90.719436283188, 2.828863371557)
    pts(64) = Array(90.734199293673, 2.624718592167)
    pts(65) = Array(90.747851907718, 2.420573812776)
    pts(66) = Array(90.760395330395, 2.216429033385)
    pts(67) = Array(90.771830667337, 2.012284253994)
    pts(68) = Array(90.782158925224, 1.808139474604)
    pts(69) = Array(90.791381012229, 1.603994695213)
    pts(70) = Array(90.805575697911, 1.224868676344)
    pts(71) = Array(90.811641427119, 1.020723896954)
    pts(72) = Array(90.816603572410, 0.816579117563)
    pts(73) = Array(90.820462568398, 0.612434338172)
    pts(74) = Array(90.823218752919, 0.408289558781)
    pts(75) = Array(90.824872367179, 0.204144779391)
    pts(76) = Array(90.825423555859, 0.000000000000)
    Profile_17 = pts
End Function

Private Function Profile_18() As Variant
    Dim pts(114) As Variant
    pts(0) = Array(90.925423555859, 0.000000000000)
    pts(1) = Array(90.926123869602, 0.307502537985)
    pts(2) = Array(90.927951722696, 0.584254822172)
    pts(3) = Array(90.931313420813, 0.891757360158)
    pts(4) = Array(90.935536791340, 1.168509644345)
    pts(5) = Array(90.941560629409, 1.476012182330)
    pts(6) = Array(90.948180466832, 1.752764466517)
    pts(7) = Array(90.956867799492, 2.060267004502)
    pts(8) = Array(90.965885593035, 2.337019288689)
    pts(9) = Array(90.977238375512, 2.644521826675)
    pts(10) = Array(90.988656155841, 2.921274110861)
    pts(11) = Array(91.002676946145, 3.228776648847)
    pts(12) = Array(91.016497287655, 3.505528933034)
    pts(13) = Array(91.033189249475, 3.813031471019)
    pts(14) = Array(91.049415273201, 4.089783755206)
    pts(15) = Array(91.068782179482, 4.397286293192)
    pts(16) = Array(91.087417556635, 4.674038577378)
    pts(17) = Array(91.109463793859, 4.981541115364)
    pts(18) = Array(91.130512750001, 5.258293399551)
    pts(19) = Array(91.155243323184, 5.565795937536)
    pts(20) = Array(91.178710643045, 5.842548221723)
    pts(21) = Array(91.206131181457, 6.150050759708)
    pts(22) = Array(91.232022214422, 6.426803043895)
    pts(23) = Array(91.262138978051, 6.734305581881)
    pts(24) = Array(91.290459644310, 7.011057866068)
    pts(25) = Array(91.323279531085, 7.318560404053)
    pts(26) = Array(91.354036328488, 7.595312688240)
    pts(27) = Array(91.389566882281, 7.902815226225)
    pts(28) = Array(91.422766893893, 8.179567510412)
    pts(29) = Array(91.461016313319, 8.487070048398)
    pts(30) = Array(91.496667215713, 8.763822332584)
    pts(31) = Array(91.533481918779, 9.040574616771)
    pts(32) = Array(91.571462344812, 9.317326900958)
    pts(33) = Array(91.615032443832, 9.624829438944)
    pts(34) = Array(91.655480447865, 9.901581723131)
    pts(35) = Array(91.701797481997, 10.209084261116)
    pts(36) = Array(91.744722545474, 10.485836545303)
    pts(37) = Array(91.793797400181, 10.793339083288)
    pts(38) = Array(91.839209638404, 11.070091367475)
    pts(39) = Array(91.885803728052, 11.346843651662)
    pts(40) = Array(91.933582153243, 11.623595935849)
    pts(41) = Array(91.988061437230, 11.931098473834)
    pts(42) = Array(92.038348601665, 12.207850758021)
    pts(43) = Array(92.095621980063, 12.515353296007)
    pts(44) = Array(92.148429942227, 12.792105580194)
    pts(45) = Array(92.208511193339, 13.099608118179)
    pts(46) = Array(92.263852699525, 13.376360402366)
    pts(47) = Array(92.320398956301, 13.653112686553)
    pts(48) = Array(92.378153057470, 13.929864970740)
    pts(49) = Array(92.443744747852, 14.237367508725)
    pts(50) = Array(92.504059250432, 14.514119792912)
    pts(51) = Array(92.565591719344, 14.790872077099)
    pts(52) = Array(92.628345571071, 15.067624361286)
    pts(53) = Array(92.699508815669, 15.375126899271)
    pts(54) = Array(92.764852727553, 15.651879183458)
    pts(55) = Array(92.831429173947, 15.928631467645)
    pts(56) = Array(92.899241911069, 16.205383751832)
    pts(57) = Array(92.976044042830, 16.512886289817)
    pts(58) = Array(93.046479448319, 16.789638574004)
    pts(59) = Array(93.118163382366, 17.066390858191)
    pts(60) = Array(93.191099960158, 17.343143142378)
    pts(61) = Array(93.265293387376, 17.619895426565)
    pts(62) = Array(93.340747961514, 17.896647710751)
    pts(63) = Array(93.417468073234, 18.173399994938)
    pts(64) = Array(93.495458207759, 18.450152279125)
    pts(65) = Array(94.216474860807, 18.450152279125)
    pts(66) = Array(94.216474860807, 18.420023289926)
    pts(67) = Array(94.293905004267, 18.019648725946)
    pts(68) = Array(94.369456115900, 17.619895426565)
    pts(69) = Array(94.443253932465, 17.220142127184)
    pts(70) = Array(94.515302562822, 16.820388827803)
    pts(71) = Array(94.569536827430, 16.512886289817)
    pts(72) = Array(94.622740215596, 16.205383751832)
    pts(73) = Array(94.674914466065, 15.897881213846)
    pts(74) = Array(94.726061280152, 15.590378675861)
    pts(75) = Array(94.776182322032, 15.282876137875)
    pts(76) = Array(94.830132642306, 14.944623346091)
    pts(77) = Array(94.882845861326, 14.606370554307)
    pts(78) = Array(94.934324040791, 14.268117762524)
    pts(79) = Array(94.984569189680, 13.929864970740)
    pts(80) = Array(95.033583264650, 13.591612178956)
    pts(81) = Array(95.081368170431, 13.253359387172)
    pts(82) = Array(95.127925760205, 12.915106595388)
    pts(83) = Array(95.181368502419, 12.515353296007)
    pts(84) = Array(95.233102456967, 12.115599996626)
    pts(85) = Array(95.283130408503, 11.715846697245)
    pts(86) = Array(95.331455044102, 11.316093397863)
    pts(87) = Array(95.378078953995, 10.916340098482)
    pts(88) = Array(95.423004632276, 10.516586799101)
    pts(89) = Array(95.466234477585, 10.116833499720)
    pts(90) = Array(95.510895812804, 9.686329946541)
    pts(91) = Array(95.550610796823, 9.286576647160)
    pts(92) = Array(95.588636736332, 8.886823347779)
    pts(93) = Array(95.624975647248, 8.487070048398)
    pts(94) = Array(95.659629453055, 8.087316749017)
    pts(95) = Array(95.692599985315, 7.687563449636)
    pts(96) = Array(95.723888984164, 7.287810150254)
    pts(97) = Array(95.753498098769, 6.888056850873)
    pts(98) = Array(95.783507936263, 6.457553297694)
    pts(99) = Array(95.809632937434, 6.057799998313)
    pts(100) = Array(95.834082562681, 5.658046698932)
    pts(101) = Array(95.856858094582, 5.258293399551)
    pts(102) = Array(95.877960726718, 4.858540100170)
    pts(103) = Array(95.897391563985, 4.458786800789)
    pts(104) = Array(95.915151622888, 4.059033501408)
    pts(105) = Array(95.931241831807, 3.659280202027)
    pts(106) = Array(95.946703245867, 3.228776648847)
    pts(107) = Array(95.959327889881, 2.829023349466)
    pts(108) = Array(95.970284990173, 2.429270050085)
    pts(109) = Array(95.979575118112, 2.029516750704)
    pts(110) = Array(95.987198757934, 1.629763451323)
    pts(111) = Array(95.993156306874, 1.230010151942)
    pts(112) = Array(95.997448075266, 0.830256852561)
    pts(113) = Array(96.000074286623, 0.430503553180)
    pts(114) = Array(96.001040004807, 0.000000000000)
    Profile_18 = pts
End Function

Private Sub AddBondedCellDetails(ByVal model As SldWorks.ModelDoc2, ByVal plane As SldWorks.Feature, ByVal mainAxis As SldWorks.SketchSegment, ByVal label As String)
    Select Case label
        Case "01_CommonBarrel"
        Case "02_L1_BondedCell"
            CellBondAccess model, plane, mainAxis, label, 9.500000000000, 0#, 2.000000000000, 0#
            CapToolSlots model, plane, 5.750000000000, 6.800000000000, 29.200000000000, 30.050000000000
        Case "03_L2_BondedCell"
            CellBondAccess model, plane, mainAxis, label, 16.200000000000, 0#, 2.000000000000, 0#
            CapToolSlots model, plane, 13.700000000000, 14.750000000000, 25.200000000000, 26.050000000000
        Case "04_L3_BondedCell"
            CellBondAccess model, plane, mainAxis, label, 19.300000000000, 20.850000000000, 0#, 2.000000000000
            CapToolSlots model, plane, 19.300000000000, 20.350000000000, 23.700000000000, 24.550000000000
        Case "05_G4G5_Retainer"
        Case "06_Spacer_G4_G5"
        Case "07_RearSpacer_1"
        Case "08_RearSpacer_2"
        Case "09_RearRetainer"
    End Select
End Sub
Private Function NewModularProfile(ByVal model As SldWorks.ModelDoc2) As SldWorks.Feature
    Dim feat As SldWorks.Feature, profile As SldWorks.Feature, found As Long
    Set feat = model.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = "ProfileFeature" Then
            If feat.Name <> "Meridional_Profile" And InStr(feat.Name, "_Profile") = 0 Then
                Set profile = feat: found = found + 1
            End If
        End If
        Set feat = feat.GetNextFeature
    Loop
    If found <> 1 Then Err.Raise vbObjectError + 607, , "Expected one new modular detail profile."
    Set NewModularProfile = profile
End Function

Private Sub CapToolSlots(ByVal model As SldWorks.ModelDoc2, ByVal plane As SldWorks.Feature, ByVal zStartMM As Double, ByVal zEndMM As Double, ByVal rInnerMM As Double, ByVal rOuterMM As Double)
    Dim math As SldWorks.MathUtility, sketch As SldWorks.Sketch, mapping As SldWorks.MathTransform
    Dim seg As SldWorks.SketchSegment, profile As SldWorks.Feature, cut As SldWorks.Feature
    Dim coords(3, 1) As Double, a As Variant, b As Variant, side As Long, i As Long, j As Long
    Set math = app.GetMathUtility
    model.ClearSelection2 True
    If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 608, , "Cap slot plane selection failed."
    model.SketchManager.InsertSketch True
    Set sketch = model.SketchManager.ActiveSketch
    If sketch Is Nothing Then Err.Raise vbObjectError + 609, , "Cap slot sketch missing."
    Set mapping = sketch.ModelToSketchTransform
    model.SketchManager.AddToDB = True
    For side = -1 To 1 Step 2
        coords(0, 0) = zStartMM: coords(0, 1) = side * rInnerMM
        coords(1, 0) = zEndMM: coords(1, 1) = side * rInnerMM
        coords(2, 0) = zEndMM: coords(2, 1) = side * rOuterMM
        coords(3, 0) = zStartMM: coords(3, 1) = side * rOuterMM
        For i = 0 To 3
            j = (i + 1) Mod 4
            a = SketchPoint(math, mapping, coords(i, 0), coords(i, 1))
            b = SketchPoint(math, mapping, coords(j, 0), coords(j, 1))
            Set seg = model.SketchManager.CreateLine(a(0), a(1), 0#, b(0), b(1), 0#)
            If seg Is Nothing Then Err.Raise vbObjectError + 610, , "Cap slot line failed."
            seg.ConstructionGeometry = False
        Next i
    Next side
    model.SketchManager.AddToDB = False
    model.SketchManager.InsertSketch True
    Set profile = NewModularProfile(model)
    profile.Name = "CapWrench_Profile"
    model.ClearSelection2 True
    If profile.Select2(False, 0) = False Then Err.Raise vbObjectError + 611, , "Cap slot profile selection failed."
    Set cut = model.FeatureManager.FeatureCut3(False, False, False, 0, 0, 0.001, 0.001, False, False, False, False, 0#, 0#, _
        False, False, False, False, False, False, True, False, False, False, 0, 0#, False)
    If cut Is Nothing Then Err.Raise vbObjectError + 612, , "Cap wrench slot cut failed."
    cut.Name = "Two_Wrench_Slots_W2_A1"
    model.ClearSelection2 True
    model.EditRebuild3
End Sub

' Native, editable details for the PF-2 bonded-cell candidate.
Private Sub CellBondAccess(ByVal model As SldWorks.ModelDoc2, ByVal basePlane As SldWorks.Feature, ByVal mainAxis As SldWorks.SketchSegment, ByVal label As String, ByVal axialLowMM As Double, ByVal axialHighMM As Double, ByVal diameterMM As Double, ByVal windowWidthMM As Double)
    Dim plane As SldWorks.Feature, feat As SldWorks.Feature, profile As SldWorks.Feature
    Dim ref As SldWorks.RefPlane, created As SldWorks.RefPlane, tf As SldWorks.MathTransform
    Dim sd As SldWorks.SelectData, math As SldWorks.MathUtility
    Dim sketch As SldWorks.Sketch, mapping As SldWorks.MathTransform
    Dim portCircle As SldWorks.SketchSegment, segment As SldWorks.SketchSegment, cut As SldWorks.Feature
    Dim data As Variant, targets As Variant, a As Variant, b As Variant
    Dim seen(2) As Boolean, coords(3, 1) As Double
    Dim direction As Long, j As Long, k As Long, matched As Long
    Dim angle As Double, dot As Double, normalDot As Double, reverseCut As Boolean
    targets = Array(30#, 150#, 270#)
    Set math = app.GetMathUtility
    For direction = 0 To 2
        currentStage = label & ": access plane " & CStr(direction)
        If direction = 0 Then
            Set plane = basePlane
        Else
            model.ClearSelection2 True
            If basePlane.Select2(False, 0) = False Then Err.Raise vbObjectError + 701, , "Cell access base plane selection failed."
            Set sd = model.SelectionManager.CreateSelectData: sd.Mark = 1
            If mainAxis.Select4(True, sd) = False Then Err.Raise vbObjectError + 702, , "Cell access axis selection failed."
            Set created = model.FeatureManager.InsertRefPlane(16, direction * 1.047197551196598, 4, 0#, 0, 0#)
            If created Is Nothing Then Err.Raise vbObjectError + 703, , "Cell access plane creation failed."
            Set plane = Nothing: Set feat = model.FirstFeature
            Do While Not feat Is Nothing
                If feat.GetTypeName2 = "RefPlane" Then Set plane = feat
                Set feat = feat.GetNextFeature
            Loop
            If plane Is Nothing Then Err.Raise vbObjectError + 704, , "Cell access plane missing."
            plane.Name = label & "_AccessPlane_" & CStr(direction)
        End If
        Set ref = plane.GetSpecificFeature2
        Set tf = ref.Transform: data = tf.ArrayData
        If Abs(data(8)) > 0.000001 Then Err.Raise vbObjectError + 705, , "Cell access plane is not parallel to optical axis."
        matched = -1
        For j = 0 To 2
            angle = CDbl(targets(j)) * 0.0174532925199433
            dot = data(6) * Cos(angle) + data(7) * Sin(angle)
            If Abs(Abs(dot) - 1#) < 0.000001 Then
                matched = j: normalDot = dot: Exit For
            End If
        Next j
        If matched < 0 Then Err.Raise vbObjectError + 706, , "Access azimuth does not match 30/150/270 degrees."
        If seen(matched) Then Err.Raise vbObjectError + 707, , "Duplicate access azimuth."
        seen(matched) = True
        angle = CDbl(targets(matched)) * 0.0174532925199433
        ' Default cut direction is minus sketch normal; choose the required outward ray.
        reverseCut = (normalDot > 0#)
        model.ClearSelection2 True
        If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 708, , "Cell access sketch plane selection failed."
        model.SketchManager.InsertSketch True
        Set sketch = model.SketchManager.ActiveSketch
        If sketch Is Nothing Then Err.Raise vbObjectError + 709, , "Cell access sketch missing."
        Set mapping = sketch.ModelToSketchTransform
        model.SketchManager.AddToDB = True
        If windowWidthMM > 0# Then
            coords(0, 0) = -windowWidthMM / 2#: coords(0, 1) = axialLowMM
            coords(1, 0) = windowWidthMM / 2#: coords(1, 1) = axialLowMM
            coords(2, 0) = windowWidthMM / 2#: coords(2, 1) = axialHighMM
            coords(3, 0) = -windowWidthMM / 2#: coords(3, 1) = axialHighMM
            For j = 0 To 3
                k = (j + 1) Mod 4
                a = CellMappedPoint(math, mapping, -Sin(angle) * coords(j, 0), Cos(angle) * coords(j, 0), coords(j, 1))
                b = CellMappedPoint(math, mapping, -Sin(angle) * coords(k, 0), Cos(angle) * coords(k, 0), coords(k, 1))
                Set segment = model.SketchManager.CreateLine(a(0), a(1), 0#, b(0), b(1), 0#)
                If segment Is Nothing Then Err.Raise vbObjectError + 710, , "Cell glue window line failed."
            Next j
        Else
            a = CellMappedPoint(math, mapping, 0#, 0#, axialLowMM)
            Set portCircle = model.SketchManager.CreateCircleByRadius(a(0), a(1), 0#, diameterMM / 2000#)
            If portCircle Is Nothing Then Err.Raise vbObjectError + 711, , "Cell glue hole circle failed."
            portCircle.ConstructionGeometry = False
        End If
        model.SketchManager.AddToDB = False
        model.SketchManager.InsertSketch True
        model.ClearSelection2 True: model.EditRebuild3
        Set profile = NewModularProfile(model)
        profile.Name = label & "_Access_Profile_" & CStr(matched + 1)
        If profile.Select2(False, 0) = False Then Err.Raise vbObjectError + 712, , "Cell access profile selection failed."
        currentStage = label & ": cut access azimuth " & CStr(targets(matched))
        Set cut = model.FeatureManager.FeatureCut3(True, False, reverseCut, 1, 0, 0.08, 0#, False, False, False, False, 0#, 0#, _
            False, False, False, False, False, False, True, False, False, False, 0, 0#, False)
        If cut Is Nothing Then Err.Raise vbObjectError + 713, , "Cell adhesive access cut failed: " & CStr(targets(matched))
        If windowWidthMM > 0# Then
            cut.Name = label & "_OpenGlueWindow_" & CStr(matched + 1)
        Else
            cut.Name = label & "_GlueHole_" & CStr(matched + 1)
        End If
        model.ClearSelection2 True: model.EditRebuild3
        buildLog.WriteLine "Created " & cut.Name & "; azimuth_deg=" & CStr(targets(matched)) & "; reverse=" & CStr(reverseCut)
    Next direction
End Sub

Private Function CellMappedPoint(ByVal math As SldWorks.MathUtility, ByVal mapping As SldWorks.MathTransform, ByVal xMM As Double, ByVal yMM As Double, ByVal zMM As Double) As Variant
    Dim xyz(2) As Double, point As SldWorks.MathPoint, data As Variant
    xyz(0) = xMM / 1000#: xyz(1) = yMM / 1000#: xyz(2) = zMM / 1000#
    Set point = math.CreatePoint(xyz)
    If point Is Nothing Then Err.Raise vbObjectError + 714, , "Cell access point creation failed."
    Set point = point.MultiplyTransform(mapping): data = point.ArrayData
    If Abs(data(2)) > 0.0000001 Then Err.Raise vbObjectError + 715, , "Cell access point does not lie on sketch plane."
    CellMappedPoint = data
End Function
