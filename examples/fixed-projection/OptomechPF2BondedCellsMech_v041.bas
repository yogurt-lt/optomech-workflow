Attribute VB_Name = "OptomechPF2BondedCellsMech_v041"
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
    Dim loaded(8) As SldWorks.ModelDoc2
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
    names = Array("01_CommonBarrel", "02_L1_BondedCell", "03_L2_BondedCell", "04_L3_BondedCell", "05_G4G5_Retainer", "06_Spacer_G4_G5", "07_RearSpacer_1", "08_RearSpacer_2", "09_RearRetainer")
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
    a = SketchPoint(math, mapping, 1.619523040734, 0#)
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
