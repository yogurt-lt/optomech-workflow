Attribute VB_Name = "OptomechMicroCells_v10"
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
    Dim loaded(16) As SldWorks.ModelDoc2
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
    names = Array("01_ObjectiveBarrel", "02_G1_CellAdapter", "03_G2_CellAdapter", "04_G3_CellAdapter", "05_G4_CellAdapter", "06_G5_CellAdapter", "07_NoseRetainer", "Lens1_Optical_Reference", "Lens2_Optical_Reference", "Lens3_Optical_Reference", "Lens4_Optical_Reference", "Lens5_Optical_Reference", "Lens6_Optical_Reference", "Lens7_Optical_Reference", "Lens8_Optical_Reference", "Lens9_Optical_Reference", "Lens10_Optical_Reference")
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
    a = SketchPoint(math, mapping, -5.345842111572, 0#)
    b = SketchPoint(math, mapping, 39.346291438329, 0#)
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
    If label = "07_NoseRetainer" Then AddRetainerSlots model, plane
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
        Case Else: Err.Raise vbObjectError + 418, , "Unknown profile index."
    End Select
End Function

Private Function Profile_0() As Variant
    Dim pts(7) As Variant
    pts(0) = Array(-4.345842111572, 10.800000000000)
    pts(1) = Array(38.346291438329, 10.800000000000)
    pts(2) = Array(38.346291438329, 9.500000000000)
    pts(3) = Array(35.866291438329, 9.500000000000)
    pts(4) = Array(35.866291438329, 9.020000000000)
    pts(5) = Array(-3.345842111572, 9.020000000000)
    pts(6) = Array(-3.345842111572, 8.000000000000)
    pts(7) = Array(-4.345842111572, 8.000000000000)
    Profile_0 = pts
End Function

Private Function Profile_1() As Variant
    Dim pts(8) As Variant
    pts(0) = Array(-3.345842111572, 8.900000000000)
    pts(1) = Array(-3.245842111572, 9.000000000000)
    pts(2) = Array(11.209808444006, 9.000000000000)
    pts(3) = Array(11.309808444006, 8.900000000000)
    pts(4) = Array(11.309808444006, 6.100000000000)
    pts(5) = Array(9.806832236696, 6.100000000000)
    pts(6) = Array(9.806832236696, 6.640000000000)
    pts(7) = Array(-3.195842111572, 6.640000000000)
    pts(8) = Array(-3.345842111572, 6.790000000000)
    Profile_1 = pts
End Function

Private Function Profile_2() As Variant
    Dim pts(8) As Variant
    pts(0) = Array(11.309808444006, 8.900000000000)
    pts(1) = Array(11.409808444006, 9.000000000000)
    pts(2) = Array(18.320832507803, 9.000000000000)
    pts(3) = Array(18.420832507803, 8.900000000000)
    pts(4) = Array(18.420832507803, 6.550000000000)
    pts(5) = Array(16.843500492553, 6.550000000000)
    pts(6) = Array(16.843500492553, 7.090000000000)
    pts(7) = Array(11.459808444006, 7.090000000000)
    pts(8) = Array(11.309808444006, 7.240000000000)
    Profile_2 = pts
End Function

Private Function Profile_3() As Variant
    Dim pts(8) As Variant
    pts(0) = Array(18.420832507803, 8.900000000000)
    pts(1) = Array(18.520832507803, 9.000000000000)
    pts(2) = Array(24.982978927004, 9.000000000000)
    pts(3) = Array(25.082978927004, 8.900000000000)
    pts(4) = Array(25.082978927004, 7.050000000000)
    pts(5) = Array(24.368071624474, 7.050000000000)
    pts(6) = Array(24.368071624474, 7.590000000000)
    pts(7) = Array(18.570832507803, 7.590000000000)
    pts(8) = Array(18.420832507803, 7.740000000000)
    Profile_3 = pts
End Function

Private Function Profile_4() As Variant
    Dim pts(8) As Variant
    pts(0) = Array(25.082978927004, 8.900000000000)
    pts(1) = Array(25.182978927004, 9.000000000000)
    pts(2) = Array(31.631352640362, 9.000000000000)
    pts(3) = Array(31.731352640362, 8.900000000000)
    pts(4) = Array(31.731352640362, 6.950000000000)
    pts(5) = Array(29.107888663194, 6.950000000000)
    pts(6) = Array(29.107888663194, 7.490000000000)
    pts(7) = Array(25.232978927004, 7.490000000000)
    pts(8) = Array(25.082978927004, 7.640000000000)
    Profile_4 = pts
End Function

Private Function Profile_5() As Variant
    Dim pts(8) As Variant
    pts(0) = Array(31.731352640362, 8.900000000000)
    pts(1) = Array(31.831352640362, 9.000000000000)
    pts(2) = Array(35.846291438329, 9.000000000000)
    pts(3) = Array(35.946291438329, 8.900000000000)
    pts(4) = Array(35.946291438329, 5.900000000000)
    pts(5) = Array(34.946291438329, 5.900000000000)
    pts(6) = Array(34.946291438329, 6.440000000000)
    pts(7) = Array(31.881352640362, 6.440000000000)
    pts(8) = Array(31.731352640362, 6.590000000000)
    Profile_5 = pts
End Function

Private Function Profile_6() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(35.946291438329, 9.460000000000)
    pts(1) = Array(37.746291438329, 9.460000000000)
    pts(2) = Array(37.746291438329, 7.900000000000)
    pts(3) = Array(35.946291438329, 7.900000000000)
    Profile_6 = pts
End Function

Private Function Profile_7() As Variant
    Dim pts(82) As Variant
    pts(0) = Array(0.000000000000, 0.000000000000)
    pts(1) = Array(-0.000862443631, 0.088000000000)
    pts(2) = Array(-0.003032761430, 0.165000000000)
    pts(3) = Array(-0.006526346682, 0.242000000000)
    pts(4) = Array(-0.011346296039, 0.319000000000)
    pts(5) = Array(-0.017496897709, 0.396000000000)
    pts(6) = Array(-0.024983650646, 0.473000000000)
    pts(7) = Array(-0.033813289369, 0.550000000000)
    pts(8) = Array(-0.043993814669, 0.627000000000)
    pts(9) = Array(-0.055534530521, 0.704000000000)
    pts(10) = Array(-0.068446087629, 0.781000000000)
    pts(11) = Array(-0.082740534073, 0.858000000000)
    pts(12) = Array(-0.098431373643, 0.935000000000)
    pts(13) = Array(-0.115533632544, 1.012000000000)
    pts(14) = Array(-0.134063935272, 1.089000000000)
    pts(15) = Array(-0.154040590596, 1.166000000000)
    pts(16) = Array(-0.175483688755, 1.243000000000)
    pts(17) = Array(-0.198415211137, 1.320000000000)
    pts(18) = Array(-0.222859153953, 1.397000000000)
    pts(19) = Array(-0.248841667657, 1.474000000000)
    pts(20) = Array(-0.272358478960, 1.540000000000)
    pts(21) = Array(-0.301275731204, 1.617000000000)
    pts(22) = Array(-0.331819628732, 1.694000000000)
    pts(23) = Array(-0.364026296739, 1.771000000000)
    pts(24) = Array(-0.392985111084, 1.837000000000)
    pts(25) = Array(-0.428386527499, 1.914000000000)
    pts(26) = Array(-0.460148886125, 1.980000000000)
    pts(27) = Array(-0.493253323014, 2.046000000000)
    pts(28) = Array(-0.527733476910, 2.112000000000)
    pts(29) = Array(-0.569747584657, 2.189000000000)
    pts(30) = Array(-0.607336610006, 2.255000000000)
    pts(31) = Array(-0.646426272335, 2.321000000000)
    pts(32) = Array(-0.687062845642, 2.387000000000)
    pts(33) = Array(-0.729296475392, 2.453000000000)
    pts(34) = Array(-0.773181602499, 2.519000000000)
    pts(35) = Array(-0.818777451584, 2.585000000000)
    pts(36) = Array(-0.866148595762, 2.651000000000)
    pts(37) = Array(-0.915365613101, 2.717000000000)
    pts(38) = Array(-0.957845416746, 2.772000000000)
    pts(39) = Array(-1.010652934817, 2.838000000000)
    pts(40) = Array(-1.056248261741, 2.893000000000)
    pts(41) = Array(-1.112957062755, 2.959000000000)
    pts(42) = Array(-1.161951923424, 3.014000000000)
    pts(43) = Array(-1.212601794106, 3.069000000000)
    pts(44) = Array(-1.264984651199, 3.124000000000)
    pts(45) = Array(-1.319186382015, 3.179000000000)
    pts(46) = Array(-1.375301940798, 3.234000000000)
    pts(47) = Array(-1.433436733846, 3.289000000000)
    pts(48) = Array(-1.493708291905, 3.344000000000)
    pts(49) = Array(-1.556248306349, 3.399000000000)
    pts(50) = Array(-1.608012664844, 3.443000000000)
    pts(51) = Array(-1.675022913060, 3.498000000000)
    pts(52) = Array(-1.730591367702, 3.542000000000)
    pts(53) = Array(-1.802679587396, 3.597000000000)
    pts(54) = Array(-1.862602237955, 3.641000000000)
    pts(55) = Array(-1.924678382737, 3.685000000000)
    pts(56) = Array(-1.989068373586, 3.729000000000)
    pts(57) = Array(-2.055955834419, 3.773000000000)
    pts(58) = Array(-2.125552707291, 3.817000000000)
    pts(59) = Array(-2.198105805235, 3.861000000000)
    pts(60) = Array(-2.254633363406, 3.894000000000)
    pts(61) = Array(-2.333091100672, 3.938000000000)
    pts(62) = Array(-2.394466893605, 3.971000000000)
    pts(63) = Array(-2.480055970929, 4.015000000000)
    pts(64) = Array(-2.545842111572, 4.047264521237)
    pts(65) = Array(-2.545842111572, 6.600000000000)
    pts(66) = Array(5.320955105176, 6.600000000000)
    pts(67) = Array(5.261741255449, 6.193000000000)
    pts(68) = Array(5.206367257666, 5.786000000000)
    pts(69) = Array(5.154818648053, 5.379000000000)
    pts(70) = Array(5.105844635729, 4.961000000000)
    pts(71) = Array(5.062010182148, 4.554000000000)
    pts(72) = Array(5.020933888737, 4.136000000000)
    pts(73) = Array(4.984767167219, 3.729000000000)
    pts(74) = Array(4.951545256820, 3.311000000000)
    pts(75) = Array(4.923008244238, 2.904000000000)
    pts(76) = Array(4.897605899088, 2.486000000000)
    pts(77) = Array(4.876668770952, 2.079000000000)
    pts(78) = Array(4.859059507225, 1.661000000000)
    pts(79) = Array(4.845700480621, 1.254000000000)
    pts(80) = Array(4.835866009522, 0.836000000000)
    pts(81) = Array(4.829966371814, 0.418000000000)
    pts(82) = Array(4.828000000000, 0.000000000000)
    Profile_7 = pts
End Function

Private Function Profile_8() As Variant
    Dim pts(85) As Variant
    pts(0) = Array(4.828000000000, 0.000000000000)
    pts(1) = Array(4.829966371814, 0.418000000000)
    pts(2) = Array(4.835866009522, 0.836000000000)
    pts(3) = Array(4.845700480621, 1.254000000000)
    pts(4) = Array(4.859059507225, 1.661000000000)
    pts(5) = Array(4.876668770952, 2.079000000000)
    pts(6) = Array(4.897605899088, 2.486000000000)
    pts(7) = Array(4.923008244238, 2.904000000000)
    pts(8) = Array(4.951545256820, 3.311000000000)
    pts(9) = Array(4.984767167219, 3.729000000000)
    pts(10) = Array(5.020933888737, 4.136000000000)
    pts(11) = Array(5.062010182148, 4.554000000000)
    pts(12) = Array(5.105844635729, 4.961000000000)
    pts(13) = Array(5.154818648053, 5.379000000000)
    pts(14) = Array(5.206367257666, 5.786000000000)
    pts(15) = Array(5.261741255449, 6.193000000000)
    pts(16) = Array(5.320955105176, 6.600000000000)
    pts(17) = Array(9.275743736595, 6.600000000000)
    pts(18) = Array(9.351115841964, 6.534000000000)
    pts(19) = Array(9.424775925434, 6.468000000000)
    pts(20) = Array(9.496787744771, 6.402000000000)
    pts(21) = Array(9.567210605900, 6.336000000000)
    pts(22) = Array(9.647435881822, 6.259000000000)
    pts(23) = Array(9.725655028476, 6.182000000000)
    pts(24) = Array(9.801943500305, 6.105000000000)
    pts(25) = Array(9.876371286405, 6.028000000000)
    pts(26) = Array(9.949003445035, 5.951000000000)
    pts(27) = Array(10.029890442864, 5.863000000000)
    pts(28) = Array(10.098873698290, 5.786000000000)
    pts(29) = Array(10.175732944857, 5.698000000000)
    pts(30) = Array(10.241308168990, 5.621000000000)
    pts(31) = Array(10.314396300026, 5.533000000000)
    pts(32) = Array(10.385566672448, 5.445000000000)
    pts(33) = Array(10.454878698875, 5.357000000000)
    pts(34) = Array(10.522387904119, 5.269000000000)
    pts(35) = Array(10.588146261823, 5.181000000000)
    pts(36) = Array(10.652202493934, 5.093000000000)
    pts(37) = Array(10.722288023260, 4.994000000000)
    pts(38) = Array(10.782875691451, 4.906000000000)
    pts(39) = Array(10.841895314538, 4.818000000000)
    pts(40) = Array(10.899384736367, 4.730000000000)
    pts(41) = Array(10.962275640223, 4.631000000000)
    pts(42) = Array(11.016629451482, 4.543000000000)
    pts(43) = Array(11.076075601909, 4.444000000000)
    pts(44) = Array(11.133760595871, 4.345000000000)
    pts(45) = Array(11.189724395178, 4.246000000000)
    pts(46) = Array(11.244004621793, 4.147000000000)
    pts(47) = Array(11.296636725371, 4.048000000000)
    pts(48) = Array(11.347654134899, 3.949000000000)
    pts(49) = Array(11.397088396233, 3.850000000000)
    pts(50) = Array(11.444969297073, 3.751000000000)
    pts(51) = Array(11.491324980726, 3.652000000000)
    pts(52) = Array(11.536182049847, 3.553000000000)
    pts(53) = Array(11.579565661179, 3.454000000000)
    pts(54) = Array(11.621499612211, 3.355000000000)
    pts(55) = Array(11.666420018733, 3.245000000000)
    pts(56) = Array(11.705366056620, 3.146000000000)
    pts(57) = Array(11.747017656384, 3.036000000000)
    pts(58) = Array(11.783065206554, 2.937000000000)
    pts(59) = Array(11.821541759899, 2.827000000000)
    pts(60) = Array(11.854770438212, 2.728000000000)
    pts(61) = Array(11.890155396653, 2.618000000000)
    pts(62) = Array(11.920636029404, 2.519000000000)
    pts(63) = Array(11.953003575685, 2.409000000000)
    pts(64) = Array(11.983809937292, 2.299000000000)
    pts(65) = Array(12.013072189041, 2.189000000000)
    pts(66) = Array(12.038101238333, 2.090000000000)
    pts(67) = Array(12.064472872391, 1.980000000000)
    pts(68) = Array(12.089343940987, 1.870000000000)
    pts(69) = Array(12.112727539518, 1.760000000000)
    pts(70) = Array(12.134635841542, 1.650000000000)
    pts(71) = Array(12.155080132308, 1.540000000000)
    pts(72) = Array(12.174070839332, 1.430000000000)
    pts(73) = Array(12.191617560189, 1.320000000000)
    pts(74) = Array(12.207729087716, 1.210000000000)
    pts(75) = Array(12.222413432791, 1.100000000000)
    pts(76) = Array(12.235677844818, 0.990000000000)
    pts(77) = Array(12.247528830052, 0.880000000000)
    pts(78) = Array(12.257972167868, 0.770000000000)
    pts(79) = Array(12.267012925082, 0.660000000000)
    pts(80) = Array(12.274655468394, 0.550000000000)
    pts(81) = Array(12.280903475033, 0.440000000000)
    pts(82) = Array(12.285759941670, 0.330000000000)
    pts(83) = Array(12.289227191633, 0.220000000000)
    pts(84) = Array(12.291306880489, 0.110000000000)
    pts(85) = Array(12.292000000000, 0.000000000000)
    Profile_8 = pts
End Function

Private Function Profile_9() As Variant
    Dim pts(81) As Variant
    pts(0) = Array(12.564000000000, 0.000000000000)
    pts(1) = Array(12.564996666717, 0.446500000000)
    pts(2) = Array(12.567882500806, 0.881250000000)
    pts(3) = Array(12.572813666666, 1.327750000000)
    pts(4) = Array(12.579530907659, 1.762500000000)
    pts(5) = Array(12.588397721586, 2.209000000000)
    pts(6) = Array(12.598947934913, 2.643750000000)
    pts(7) = Array(12.611752463932, 3.090250000000)
    pts(8) = Array(12.626138110007, 3.525000000000)
    pts(9) = Array(12.642883341657, 3.971500000000)
    pts(10) = Array(12.661107778749, 4.406250000000)
    pts(11) = Array(12.681797624571, 4.852750000000)
    pts(12) = Array(12.703865112732, 5.287500000000)
    pts(13) = Array(12.728504412847, 5.734000000000)
    pts(14) = Array(12.754420118903, 6.168750000000)
    pts(15) = Array(12.783014647695, 6.615250000000)
    pts(16) = Array(12.812784651316, 7.050000000000)
    pts(17) = Array(15.649597580331, 7.050000000000)
    pts(18) = Array(15.578508404608, 6.956000000000)
    pts(19) = Array(15.508902242902, 6.862000000000)
    pts(20) = Array(15.440746266531, 6.768000000000)
    pts(21) = Array(15.365765414258, 6.662250000000)
    pts(22) = Array(15.300589325768, 6.568250000000)
    pts(23) = Array(15.236770785257, 6.474250000000)
    pts(24) = Array(15.174282997041, 6.380250000000)
    pts(25) = Array(15.105543123400, 6.274500000000)
    pts(26) = Array(15.045799770558, 6.180500000000)
    pts(27) = Array(14.980087101783, 6.074750000000)
    pts(28) = Array(14.922983497275, 5.980750000000)
    pts(29) = Array(14.860185791618, 5.875000000000)
    pts(30) = Array(14.805626794760, 5.781000000000)
    pts(31) = Array(14.745641917986, 5.675250000000)
    pts(32) = Array(14.687107944520, 5.569500000000)
    pts(33) = Array(14.630000003051, 5.463750000000)
    pts(34) = Array(14.574294373139, 5.358000000000)
    pts(35) = Array(14.519968421069, 5.252250000000)
    pts(36) = Array(14.467000540502, 5.146500000000)
    pts(37) = Array(14.415370097498, 5.040750000000)
    pts(38) = Array(14.365057379546, 4.935000000000)
    pts(39) = Array(14.316043548229, 4.829250000000)
    pts(40) = Array(14.268310595251, 4.723500000000)
    pts(41) = Array(14.216755327947, 4.606000000000)
    pts(42) = Array(14.171670817177, 4.500250000000)
    pts(43) = Array(14.127816037774, 4.394500000000)
    pts(44) = Array(14.085175993766, 4.288750000000)
    pts(45) = Array(14.039205427074, 4.171250000000)
    pts(46) = Array(13.999083460090, 4.065500000000)
    pts(47) = Array(13.955877633952, 3.948000000000)
    pts(48) = Array(13.918215457691, 3.842250000000)
    pts(49) = Array(13.877712561164, 3.724750000000)
    pts(50) = Array(13.842457088875, 3.619000000000)
    pts(51) = Array(13.804600840606, 3.501500000000)
    pts(52) = Array(13.771703753840, 3.395750000000)
    pts(53) = Array(13.736442954692, 3.278250000000)
    pts(54) = Array(13.705860327252, 3.172500000000)
    pts(55) = Array(13.673148469871, 3.055000000000)
    pts(56) = Array(13.641760964527, 2.937500000000)
    pts(57) = Array(13.611686738507, 2.820000000000)
    pts(58) = Array(13.585734104057, 2.714250000000)
    pts(59) = Array(13.558126660416, 2.596750000000)
    pts(60) = Array(13.531803588480, 2.479250000000)
    pts(61) = Array(13.506755965605, 2.361750000000)
    pts(62) = Array(13.482975378066, 2.244250000000)
    pts(63) = Array(13.460453906104, 2.126750000000)
    pts(64) = Array(13.439184110030, 2.009250000000)
    pts(65) = Array(13.419159017363, 1.891750000000)
    pts(66) = Array(13.400372110901, 1.774250000000)
    pts(67) = Array(13.382817317721, 1.656750000000)
    pts(68) = Array(13.366488999028, 1.539250000000)
    pts(69) = Array(13.351381940830, 1.421750000000)
    pts(70) = Array(13.337491345407, 1.304250000000)
    pts(71) = Array(13.324812823514, 1.186750000000)
    pts(72) = Array(13.313342387325, 1.069250000000)
    pts(73) = Array(13.303076444056, 0.951750000000)
    pts(74) = Array(13.294011790269, 0.834250000000)
    pts(75) = Array(13.286145606822, 0.716750000000)
    pts(76) = Array(13.279475454444, 0.599250000000)
    pts(77) = Array(13.273999269935, 0.481750000000)
    pts(78) = Array(13.269715362958, 0.364250000000)
    pts(79) = Array(13.266622413417, 0.246750000000)
    pts(80) = Array(13.264719469429, 0.129250000000)
    pts(81) = Array(13.264000000000, 0.000000000000)
    Profile_9 = pts
End Function

Private Function Profile_10() As Variant
    Dim pts(101) As Variant
    pts(0) = Array(13.264000000000, 0.000000000000)
    pts(1) = Array(13.264719469429, 0.129250000000)
    pts(2) = Array(13.266622413417, 0.246750000000)
    pts(3) = Array(13.269715362958, 0.364250000000)
    pts(4) = Array(13.273999269935, 0.481750000000)
    pts(5) = Array(13.279475454444, 0.599250000000)
    pts(6) = Array(13.286145606822, 0.716750000000)
    pts(7) = Array(13.294011790269, 0.834250000000)
    pts(8) = Array(13.303076444056, 0.951750000000)
    pts(9) = Array(13.313342387325, 1.069250000000)
    pts(10) = Array(13.324812823514, 1.186750000000)
    pts(11) = Array(13.337491345407, 1.304250000000)
    pts(12) = Array(13.351381940830, 1.421750000000)
    pts(13) = Array(13.366488999028, 1.539250000000)
    pts(14) = Array(13.382817317721, 1.656750000000)
    pts(15) = Array(13.400372110901, 1.774250000000)
    pts(16) = Array(13.419159017363, 1.891750000000)
    pts(17) = Array(13.439184110030, 2.009250000000)
    pts(18) = Array(13.460453906104, 2.126750000000)
    pts(19) = Array(13.482975378066, 2.244250000000)
    pts(20) = Array(13.506755965605, 2.361750000000)
    pts(21) = Array(13.531803588480, 2.479250000000)
    pts(22) = Array(13.558126660416, 2.596750000000)
    pts(23) = Array(13.585734104057, 2.714250000000)
    pts(24) = Array(13.611686738507, 2.820000000000)
    pts(25) = Array(13.641760964527, 2.937500000000)
    pts(26) = Array(13.673148469871, 3.055000000000)
    pts(27) = Array(13.705860327252, 3.172500000000)
    pts(28) = Array(13.736442954692, 3.278250000000)
    pts(29) = Array(13.771703753840, 3.395750000000)
    pts(30) = Array(13.804600840606, 3.501500000000)
    pts(31) = Array(13.842457088875, 3.619000000000)
    pts(32) = Array(13.877712561164, 3.724750000000)
    pts(33) = Array(13.918215457691, 3.842250000000)
    pts(34) = Array(13.955877633952, 3.948000000000)
    pts(35) = Array(13.999083460090, 4.065500000000)
    pts(36) = Array(14.039205427074, 4.171250000000)
    pts(37) = Array(14.085175993766, 4.288750000000)
    pts(38) = Array(14.127816037774, 4.394500000000)
    pts(39) = Array(14.171670817177, 4.500250000000)
    pts(40) = Array(14.216755327947, 4.606000000000)
    pts(41) = Array(14.268310595251, 4.723500000000)
    pts(42) = Array(14.316043548229, 4.829250000000)
    pts(43) = Array(14.365057379546, 4.935000000000)
    pts(44) = Array(14.415370097498, 5.040750000000)
    pts(45) = Array(14.467000540502, 5.146500000000)
    pts(46) = Array(14.519968421069, 5.252250000000)
    pts(47) = Array(14.574294373139, 5.358000000000)
    pts(48) = Array(14.630000003051, 5.463750000000)
    pts(49) = Array(14.687107944520, 5.569500000000)
    pts(50) = Array(14.745641917986, 5.675250000000)
    pts(51) = Array(14.805626794760, 5.781000000000)
    pts(52) = Array(14.860185791618, 5.875000000000)
    pts(53) = Array(14.922983497275, 5.980750000000)
    pts(54) = Array(14.980087101783, 6.074750000000)
    pts(55) = Array(15.045799770558, 6.180500000000)
    pts(56) = Array(15.105543123400, 6.274500000000)
    pts(57) = Array(15.174282997041, 6.380250000000)
    pts(58) = Array(15.236770785257, 6.474250000000)
    pts(59) = Array(15.300589325768, 6.568250000000)
    pts(60) = Array(15.365765414258, 6.662250000000)
    pts(61) = Array(15.440746266531, 6.768000000000)
    pts(62) = Array(15.508902242902, 6.862000000000)
    pts(63) = Array(15.578508404608, 6.956000000000)
    pts(64) = Array(15.649597580331, 7.050000000000)
    pts(65) = Array(16.664021878634, 7.050000000000)
    pts(66) = Array(16.737512143372, 6.850250000000)
    pts(67) = Array(16.812727157898, 6.638750000000)
    pts(68) = Array(16.885300743042, 6.427250000000)
    pts(69) = Array(16.955262937278, 6.215750000000)
    pts(70) = Array(17.015282370679, 6.027750000000)
    pts(71) = Array(17.073280118856, 5.839750000000)
    pts(72) = Array(17.129274389518, 5.651750000000)
    pts(73) = Array(17.183282554526, 5.463750000000)
    pts(74) = Array(17.235321179777, 5.275750000000)
    pts(75) = Array(17.285406053212, 5.087750000000)
    pts(76) = Array(17.333552211048, 4.899750000000)
    pts(77) = Array(17.382599239753, 4.700000000000)
    pts(78) = Array(17.426791200398, 4.512000000000)
    pts(79) = Array(17.471666789403, 4.312250000000)
    pts(80) = Array(17.511959367979, 4.124250000000)
    pts(81) = Array(17.552718676768, 3.924500000000)
    pts(82) = Array(17.589160971841, 3.736500000000)
    pts(83) = Array(17.625853260395, 3.536750000000)
    pts(84) = Array(17.658488971141, 3.348750000000)
    pts(85) = Array(17.691157913965, 3.149000000000)
    pts(86) = Array(17.720025615252, 2.961000000000)
    pts(87) = Array(17.748709573027, 2.761250000000)
    pts(88) = Array(17.775353960048, 2.561500000000)
    pts(89) = Array(17.799966931876, 2.361750000000)
    pts(90) = Array(17.821283115720, 2.173750000000)
    pts(91) = Array(17.841973567275, 1.974000000000)
    pts(92) = Array(17.860652805769, 1.774250000000)
    pts(93) = Array(17.877326447487, 1.574500000000)
    pts(94) = Array(17.891191666582, 1.386500000000)
    pts(95) = Array(17.903985802660, 1.186750000000)
    pts(96) = Array(17.914787290593, 0.987000000000)
    pts(97) = Array(17.923599340682, 0.787250000000)
    pts(98) = Array(17.930078042530, 0.599250000000)
    pts(99) = Array(17.935035173882, 0.399500000000)
    pts(100) = Array(17.938008866689, 0.199750000000)
    pts(101) = Array(17.939000000000, 0.000000000000)
    Profile_10 = pts
End Function

Private Function Profile_11() As Variant
    Dim pts(76) As Variant
    pts(0) = Array(18.043000000000, 0.000000000000)
    pts(1) = Array(18.043508977397, 0.125833333333)
    pts(2) = Array(18.045036009529, 0.251666666667)
    pts(3) = Array(18.050743346896, 0.490750000000)
    pts(4) = Array(18.055225151439, 0.616583333333)
    pts(5) = Array(18.060727314839, 0.742416666667)
    pts(6) = Array(18.073996561776, 0.981500000000)
    pts(7) = Array(18.082464627707, 1.107333333333)
    pts(8) = Array(18.091958436396, 1.233166666667)
    pts(9) = Array(18.112829515388, 1.472250000000)
    pts(10) = Array(18.125309391715, 1.598083333333)
    pts(11) = Array(18.138823556011, 1.723916666667)
    pts(12) = Array(18.167359831815, 1.963000000000)
    pts(13) = Array(18.183889571384, 2.088833333333)
    pts(14) = Array(18.201465436808, 2.214666666667)
    pts(15) = Array(18.237754691493, 2.453750000000)
    pts(16) = Array(18.277860521559, 2.692833333333)
    pts(17) = Array(18.321812986883, 2.931916666667)
    pts(18) = Array(18.346502042832, 3.057750000000)
    pts(19) = Array(18.372271046688, 3.183583333333)
    pts(20) = Array(18.424227302055, 3.422666666667)
    pts(21) = Array(18.480142614196, 3.661750000000)
    pts(22) = Array(18.540061087069, 3.900833333333)
    pts(23) = Array(18.604030712548, 4.139916666667)
    pts(24) = Array(18.672103578009, 4.379000000000)
    pts(25) = Array(18.744336097569, 4.618083333333)
    pts(26) = Array(18.816659149915, 4.844583333333)
    pts(27) = Array(18.858495830601, 4.970416666667)
    pts(28) = Array(18.897171512387, 5.083666666667)
    pts(29) = Array(18.977461194189, 5.310166666667)
    pts(30) = Array(19.023778614279, 5.436000000000)
    pts(31) = Array(19.066522117916, 5.549250000000)
    pts(32) = Array(19.155052345954, 5.775750000000)
    pts(33) = Array(19.247703141137, 6.002250000000)
    pts(34) = Array(19.344554857246, 6.228750000000)
    pts(35) = Array(19.445693741390, 6.455250000000)
    pts(36) = Array(19.497899548230, 6.568500000000)
    pts(37) = Array(19.551212397936, 6.681750000000)
    pts(38) = Array(19.605644944471, 6.795000000000)
    pts(39) = Array(19.661210307338, 6.908250000000)
    pts(40) = Array(19.769306389531, 7.122166666667)
    pts(41) = Array(19.881589889246, 7.336083333333)
    pts(42) = Array(19.998164523054, 7.550000000000)
    pts(43) = Array(20.799917655384, 7.550000000000)
    pts(44) = Array(20.838315100031, 7.373833333333)
    pts(45) = Array(20.886290699343, 7.147333333333)
    pts(46) = Array(20.932705753961, 6.920833333333)
    pts(47) = Array(20.977566581880, 6.694333333333)
    pts(48) = Array(21.020879257160, 6.467833333333)
    pts(49) = Array(21.062649614250, 6.241333333333)
    pts(50) = Array(21.102883252102, 6.014833333333)
    pts(51) = Array(21.141585538087, 5.788333333333)
    pts(52) = Array(21.180782298900, 5.549250000000)
    pts(53) = Array(21.216352698561, 5.322750000000)
    pts(54) = Array(21.250406749727, 5.096250000000)
    pts(55) = Array(21.282948916898, 4.869750000000)
    pts(56) = Array(21.315663467537, 4.630666666667)
    pts(57) = Array(21.345110985874, 4.404166666667)
    pts(58) = Array(21.373058944188, 4.177666666667)
    pts(59) = Array(21.399510953772, 3.951166666667)
    pts(60) = Array(21.425813361626, 3.712083333333)
    pts(61) = Array(21.449200843052, 3.485583333333)
    pts(62) = Array(21.471102155759, 3.259083333333)
    pts(63) = Array(21.491520097056, 3.032583333333)
    pts(64) = Array(21.511465965197, 2.793500000000)
    pts(65) = Array(21.528842709619, 2.567000000000)
    pts(66) = Array(21.545583559574, 2.327916666667)
    pts(67) = Array(21.559928422174, 2.101416666667)
    pts(68) = Array(21.573473210604, 1.862333333333)
    pts(69) = Array(21.584793852505, 1.635833333333)
    pts(70) = Array(21.595149801706, 1.396750000000)
    pts(71) = Array(21.603452250678, 1.170250000000)
    pts(72) = Array(21.610624867933, 0.931166666667)
    pts(73) = Array(21.615913534503, 0.704666666667)
    pts(74) = Array(21.619906626770, 0.465583333333)
    pts(75) = Array(21.622184315968, 0.239083333333)
    pts(76) = Array(21.623000000000, 0.000000000000)
    Profile_11 = pts
End Function

Private Function Profile_12() As Variant
    Dim pts(65) As Variant
    pts(0) = Array(21.623000000000, 0.000000000000)
    pts(1) = Array(21.622184315968, 0.239083333333)
    pts(2) = Array(21.619737149929, 0.478166666667)
    pts(3) = Array(21.615658159977, 0.717250000000)
    pts(4) = Array(21.609946776005, 0.956333333333)
    pts(5) = Array(21.602602199303, 1.195416666667)
    pts(6) = Array(21.593623402002, 1.434500000000)
    pts(7) = Array(21.583009126351, 1.673583333333)
    pts(8) = Array(21.571443523733, 1.900083333333)
    pts(9) = Array(21.557639887381, 2.139166666667)
    pts(10) = Array(21.542195719137, 2.378250000000)
    pts(11) = Array(21.525108843627, 2.617333333333)
    pts(12) = Array(21.506376849613, 2.856416666667)
    pts(13) = Array(21.485997088276, 3.095500000000)
    pts(14) = Array(21.463966671329, 3.334583333333)
    pts(15) = Array(21.440282468955, 3.573666666667)
    pts(16) = Array(21.416316235151, 3.800166666667)
    pts(17) = Array(21.389401598567, 4.039250000000)
    pts(18) = Array(21.360822524332, 4.278333333333)
    pts(19) = Array(21.330574905445, 4.517416666667)
    pts(20) = Array(21.300376188475, 4.743916666667)
    pts(21) = Array(21.266866548196, 4.983000000000)
    pts(22) = Array(21.231674765556, 5.222083333333)
    pts(23) = Array(21.194795709415, 5.461166666667)
    pts(24) = Array(21.158296361443, 5.687666666667)
    pts(25) = Array(21.118115819541, 5.926750000000)
    pts(26) = Array(21.078478332590, 6.153250000000)
    pts(27) = Array(21.034973796913, 6.392333333333)
    pts(28) = Array(20.992175864256, 6.618833333333)
    pts(29) = Array(20.945322733461, 6.857916666667)
    pts(30) = Array(20.899340031727, 7.084416666667)
    pts(31) = Array(20.849111538561, 7.323500000000)
    pts(32) = Array(20.799917655384, 7.550000000000)
    pts(33) = Array(23.647865108350, 7.550000000000)
    pts(34) = Array(23.567233823149, 7.323500000000)
    pts(35) = Array(23.489353779931, 7.097000000000)
    pts(36) = Array(23.414194490795, 6.870500000000)
    pts(37) = Array(23.337779349141, 6.631416666667)
    pts(38) = Array(23.268123234153, 6.404916666667)
    pts(39) = Array(23.201103688728, 6.178416666667)
    pts(40) = Array(23.136695896714, 5.951916666667)
    pts(41) = Array(23.071517313718, 5.712833333333)
    pts(42) = Array(23.012405487500, 5.486333333333)
    pts(43) = Array(22.952768684521, 5.247250000000)
    pts(44) = Array(22.898863571950, 5.020750000000)
    pts(45) = Array(22.844679495078, 4.781666666667)
    pts(46) = Array(22.795901446199, 4.555166666667)
    pts(47) = Array(22.747090831422, 4.316083333333)
    pts(48) = Array(22.701011767678, 4.077000000000)
    pts(49) = Array(22.657646965772, 3.837916666667)
    pts(50) = Array(22.616980287955, 3.598833333333)
    pts(51) = Array(22.578996716605, 3.359750000000)
    pts(52) = Array(22.543682325439, 3.120666666667)
    pts(53) = Array(22.511024253139, 2.881583333333)
    pts(54) = Array(22.481010679250, 2.642500000000)
    pts(55) = Array(22.453630802257, 2.403416666667)
    pts(56) = Array(22.428874819735, 2.164333333333)
    pts(57) = Array(22.406733910491, 1.925250000000)
    pts(58) = Array(22.387200218621, 1.686166666667)
    pts(59) = Array(22.370266839407, 1.447083333333)
    pts(60) = Array(22.355927807001, 1.208000000000)
    pts(61) = Array(22.344178083837, 0.968916666667)
    pts(62) = Array(22.335013551745, 0.729833333333)
    pts(63) = Array(22.328431004697, 0.490750000000)
    pts(64) = Array(22.324428143195, 0.251666666667)
    pts(65) = Array(22.323000000000, 0.000000000000)
    Profile_12 = pts
End Function

Private Function Profile_13() As Variant
    Dim pts(49) As Variant
    pts(0) = Array(22.323000000000, 0.000000000000)
    pts(1) = Array(22.324428143195, 0.251666666667)
    pts(2) = Array(22.328431004697, 0.490750000000)
    pts(3) = Array(22.335013551745, 0.729833333333)
    pts(4) = Array(22.344178083837, 0.968916666667)
    pts(5) = Array(22.355927807001, 1.208000000000)
    pts(6) = Array(22.370266839407, 1.447083333333)
    pts(7) = Array(22.387200218621, 1.686166666667)
    pts(8) = Array(22.406733910491, 1.925250000000)
    pts(9) = Array(22.428874819735, 2.164333333333)
    pts(10) = Array(22.453630802257, 2.403416666667)
    pts(11) = Array(22.481010679250, 2.642500000000)
    pts(12) = Array(22.511024253139, 2.881583333333)
    pts(13) = Array(22.543682325439, 3.120666666667)
    pts(14) = Array(22.578996716605, 3.359750000000)
    pts(15) = Array(22.616980287955, 3.598833333333)
    pts(16) = Array(22.657646965772, 3.837916666667)
    pts(17) = Array(22.701011767678, 4.077000000000)
    pts(18) = Array(22.747090831422, 4.316083333333)
    pts(19) = Array(22.795901446199, 4.555166666667)
    pts(20) = Array(22.844679495078, 4.781666666667)
    pts(21) = Array(22.898863571950, 5.020750000000)
    pts(22) = Array(22.952768684521, 5.247250000000)
    pts(23) = Array(23.012405487500, 5.486333333333)
    pts(24) = Array(23.071517313718, 5.712833333333)
    pts(25) = Array(23.136695896714, 5.951916666667)
    pts(26) = Array(23.201103688728, 6.178416666667)
    pts(27) = Array(23.268123234153, 6.404916666667)
    pts(28) = Array(23.337779349141, 6.631416666667)
    pts(29) = Array(23.414194490795, 6.870500000000)
    pts(30) = Array(23.489353779931, 7.097000000000)
    pts(31) = Array(23.567233823149, 7.323500000000)
    pts(32) = Array(23.647865108350, 7.550000000000)
    pts(33) = Array(24.347804648784, 7.550000000000)
    pts(34) = Array(24.366721138823, 7.084416666667)
    pts(35) = Array(24.384894149904, 6.606250000000)
    pts(36) = Array(24.401367571025, 6.140666666667)
    pts(37) = Array(24.417032207679, 5.662500000000)
    pts(38) = Array(24.431063918198, 5.196916666667)
    pts(39) = Array(24.444221471591, 4.718750000000)
    pts(40) = Array(24.455812627697, 4.253166666667)
    pts(41) = Array(24.466464181662, 3.775000000000)
    pts(42) = Array(24.475615737897, 3.309416666667)
    pts(43) = Array(24.483762169383, 2.831250000000)
    pts(44) = Array(24.490474879033, 2.365666666667)
    pts(45) = Array(24.496116858463, 1.887500000000)
    pts(46) = Array(24.500391273871, 1.421916666667)
    pts(47) = Array(24.503529265421, 0.943750000000)
    pts(48) = Array(24.505365738212, 0.478166666667)
    pts(49) = Array(24.506000000000, 0.000000000000)
    Profile_13 = pts
End Function

Private Function Profile_14() As Variant
    Dim pts(82) As Variant
    pts(0) = Array(24.582000000000, 0.000000000000)
    pts(1) = Array(24.583235640716, 0.235916666667)
    pts(2) = Array(24.586686209799, 0.459416666667)
    pts(3) = Array(24.592356150063, 0.682916666667)
    pts(4) = Array(24.600247139057, 0.906416666667)
    pts(5) = Array(24.610988673022, 1.142333333333)
    pts(6) = Array(24.623453224942, 1.365833333333)
    pts(7) = Array(24.638148056315, 1.589333333333)
    pts(8) = Array(24.655077546894, 1.812833333333)
    pts(9) = Array(24.675377495764, 2.048750000000)
    pts(10) = Array(24.696917103347, 2.272250000000)
    pts(11) = Array(24.720709001188, 2.495750000000)
    pts(12) = Array(24.746760379745, 2.719250000000)
    pts(13) = Array(24.776719064512, 2.955166666667)
    pts(14) = Array(24.807440574335, 3.178666666667)
    pts(15) = Array(24.840448037830, 3.402166666667)
    pts(16) = Array(24.875751638462, 3.625666666667)
    pts(17) = Array(24.911212130120, 3.836750000000)
    pts(18) = Array(24.948740369259, 4.047833333333)
    pts(19) = Array(24.988346882211, 4.258916666667)
    pts(20) = Array(25.027532102567, 4.457583333333)
    pts(21) = Array(25.071205452881, 4.668666666667)
    pts(22) = Array(25.114239878020, 4.867333333333)
    pts(23) = Array(25.162027250046, 5.078416666667)
    pts(24) = Array(25.208957741693, 5.277083333333)
    pts(25) = Array(25.273450625964, 5.537833333333)
    pts(26) = Array(25.337953980901, 5.786166666667)
    pts(27) = Array(25.408947964423, 6.046916666667)
    pts(28) = Array(25.479700902977, 6.295250000000)
    pts(29) = Array(25.557321648106, 6.556000000000)
    pts(30) = Array(25.634451055300, 6.804333333333)
    pts(31) = Array(25.714741553553, 7.052666666667)
    pts(32) = Array(25.797886229535, 7.300000000000)
    pts(33) = Array(25.797886229535, 7.450000000000)
    pts(34) = Array(27.093569760623, 7.450000000000)
    pts(35) = Array(27.165655239237, 7.350666666667)
    pts(36) = Array(27.245005647873, 7.238916666667)
    pts(37) = Array(27.322548364045, 7.127166666667)
    pts(38) = Array(27.398323566232, 7.015416666667)
    pts(39) = Array(27.472369374721, 6.903666666667)
    pts(40) = Array(27.544721987928, 6.791916666667)
    pts(41) = Array(27.615415807014, 6.680166666667)
    pts(42) = Array(27.684483550002, 6.568416666667)
    pts(43) = Array(27.751956356453, 6.456666666667)
    pts(44) = Array(27.817863883636, 6.344916666667)
    pts(45) = Array(27.882234395005, 6.233166666667)
    pts(46) = Array(27.951987250446, 6.109000000000)
    pts(47) = Array(28.013199954766, 5.997250000000)
    pts(48) = Array(28.072955513330, 5.885500000000)
    pts(49) = Array(28.131277250676, 5.773750000000)
    pts(50) = Array(28.194424612832, 5.649583333333)
    pts(51) = Array(28.273958677358, 5.488166666667)
    pts(52) = Array(28.350657295738, 5.326750000000)
    pts(53) = Array(28.424577237311, 5.165333333333)
    pts(54) = Array(28.501136862791, 4.991500000000)
    pts(55) = Array(28.569452255243, 4.830083333333)
    pts(56) = Array(28.640088483501, 4.656250000000)
    pts(57) = Array(28.707735256297, 4.482416666667)
    pts(58) = Array(28.772443811772, 4.308583333333)
    pts(59) = Array(28.834262072291, 4.134750000000)
    pts(60) = Array(28.893234850956, 3.960916666667)
    pts(61) = Array(28.949404038576, 3.787083333333)
    pts(62) = Array(29.002808773131, 3.613250000000)
    pts(63) = Array(29.053485593485, 3.439416666667)
    pts(64) = Array(29.101468578909, 3.265583333333)
    pts(65) = Array(29.146789475775, 3.091750000000)
    pts(66) = Array(29.192426980738, 2.905500000000)
    pts(67) = Array(29.232325086038, 2.731666666667)
    pts(68) = Array(29.272212882809, 2.545416666667)
    pts(69) = Array(29.306796664716, 2.371583333333)
    pts(70) = Array(29.341041793073, 2.185333333333)
    pts(71) = Array(29.370402942098, 2.011500000000)
    pts(72) = Array(29.399095188435, 1.825250000000)
    pts(73) = Array(29.423309899008, 1.651416666667)
    pts(74) = Array(29.446523028821, 1.465166666667)
    pts(75) = Array(29.465653027838, 1.291333333333)
    pts(76) = Array(29.483445766380, 1.105083333333)
    pts(77) = Array(29.498453171263, 0.918833333333)
    pts(78) = Array(29.510685211886, 0.732583333333)
    pts(79) = Array(29.519604952393, 0.558750000000)
    pts(80) = Array(29.526492595972, 0.372500000000)
    pts(81) = Array(29.530623374699, 0.186250000000)
    pts(82) = Array(29.532000000000, 0.000000000000)
    Profile_14 = pts
End Function

Private Function Profile_15() As Variant
    Dim pts(97) As Variant
    pts(0) = Array(29.532000000000, 0.000000000000)
    pts(1) = Array(29.531388185102, 0.124166666667)
    pts(2) = Array(29.529552562127, 0.248333333333)
    pts(3) = Array(29.526492595972, 0.372500000000)
    pts(4) = Array(29.522207393932, 0.496666666667)
    pts(5) = Array(29.516695704395, 0.620833333333)
    pts(6) = Array(29.509955915013, 0.745000000000)
    pts(7) = Array(29.501986050333, 0.869166666667)
    pts(8) = Array(29.492783768893, 0.993333333333)
    pts(9) = Array(29.482346359762, 1.117500000000)
    pts(10) = Array(29.470670738515, 1.241666666667)
    pts(11) = Array(29.457753442630, 1.365833333333)
    pts(12) = Array(29.443590626295, 1.490000000000)
    pts(13) = Array(29.428178054598, 1.614166666667)
    pts(14) = Array(29.411511097083, 1.738333333333)
    pts(15) = Array(29.393584720653, 1.862500000000)
    pts(16) = Array(29.374393481785, 1.986666666667)
    pts(17) = Array(29.353931518041, 2.110833333333)
    pts(18) = Array(29.332192538828, 2.235000000000)
    pts(19) = Array(29.309169815385, 2.359166666667)
    pts(20) = Array(29.284856169953, 2.483333333333)
    pts(21) = Array(29.259243964082, 2.607500000000)
    pts(22) = Array(29.232325086038, 2.731666666667)
    pts(23) = Array(29.204090937254, 2.855833333333)
    pts(24) = Array(29.177548136352, 2.967583333333)
    pts(25) = Array(29.146789475775, 3.091750000000)
    pts(26) = Array(29.114687716307, 3.215916666667)
    pts(27) = Array(29.081232211395, 3.340083333333)
    pts(28) = Array(29.049955546993, 3.451833333333)
    pts(29) = Array(29.013896539865, 3.576000000000)
    pts(30) = Array(28.980257132777, 3.687750000000)
    pts(31) = Array(28.941550005158, 3.811916666667)
    pts(32) = Array(28.905505747645, 3.923666666667)
    pts(33) = Array(28.864101462255, 4.047833333333)
    pts(34) = Array(28.825606118123, 4.159583333333)
    pts(35) = Array(28.781450870715, 4.283750000000)
    pts(36) = Array(28.740453740859, 4.395500000000)
    pts(37) = Array(28.693488563197, 4.519666666667)
    pts(38) = Array(28.649934103746, 4.631416666667)
    pts(39) = Array(28.605147389296, 4.743166666667)
    pts(40) = Array(28.559114180214, 4.854916666667)
    pts(41) = Array(28.506486129402, 4.979083333333)
    pts(42) = Array(28.457771816769, 5.090833333333)
    pts(43) = Array(28.407762574214, 5.202583333333)
    pts(44) = Array(28.356441396507, 5.314333333333)
    pts(45) = Array(28.303790521273, 5.426083333333)
    pts(46) = Array(28.249791391090, 5.537833333333)
    pts(47) = Array(28.194424612832, 5.649583333333)
    pts(48) = Array(28.137669914031, 5.761333333333)
    pts(49) = Array(28.079506095972, 5.873083333333)
    pts(50) = Array(28.019910983239, 5.984833333333)
    pts(51) = Array(27.958861369389, 6.096583333333)
    pts(52) = Array(27.903354367612, 6.195916666667)
    pts(53) = Array(27.839490143118, 6.307666666667)
    pts(54) = Array(27.774098028505, 6.419416666667)
    pts(55) = Array(27.707150206481, 6.531166666667)
    pts(56) = Array(27.646311445445, 6.630500000000)
    pts(57) = Array(27.576344269963, 6.742250000000)
    pts(58) = Array(27.512772065074, 6.841583333333)
    pts(59) = Array(27.447877320813, 6.940916666667)
    pts(60) = Array(27.381635397348, 7.040250000000)
    pts(61) = Array(27.305470778867, 7.152000000000)
    pts(62) = Array(27.236278960164, 7.251333333333)
    pts(63) = Array(27.165655239237, 7.350666666667)
    pts(64) = Array(27.093569760623, 7.450000000000)
    pts(65) = Array(28.912140900348, 7.450000000000)
    pts(66) = Array(29.001578756801, 7.226500000000)
    pts(67) = Array(29.087879094303, 7.003000000000)
    pts(68) = Array(29.171085698109, 6.779500000000)
    pts(69) = Array(29.251240015629, 6.556000000000)
    pts(70) = Array(29.328381276515, 6.332500000000)
    pts(71) = Array(29.402546603344, 6.109000000000)
    pts(72) = Array(29.473771113700, 5.885500000000)
    pts(73) = Array(29.545798782671, 5.649583333333)
    pts(74) = Array(29.611080551107, 5.426083333333)
    pts(75) = Array(29.676903189287, 5.190166666667)
    pts(76) = Array(29.736367136251, 4.966666666667)
    pts(77) = Array(29.796108969746, 4.730750000000)
    pts(78) = Array(29.849865852517, 4.507250000000)
    pts(79) = Array(29.903636775308, 4.271333333333)
    pts(80) = Array(29.954379911613, 4.035416666667)
    pts(81) = Array(30.002119040805, 3.799500000000)
    pts(82) = Array(30.046876307677, 3.563583333333)
    pts(83) = Array(30.088672276532, 3.327666666667)
    pts(84) = Array(30.127525980645, 3.091750000000)
    pts(85) = Array(30.163454967387, 2.855833333333)
    pts(86) = Array(30.196475339270, 2.619916666667)
    pts(87) = Array(30.226601791142, 2.384000000000)
    pts(88) = Array(30.253847643744, 2.148083333333)
    pts(89) = Array(30.278224873806, 1.912166666667)
    pts(90) = Array(30.299744140847, 1.676250000000)
    pts(91) = Array(30.318414810823, 1.440333333333)
    pts(92) = Array(30.334244976745, 1.204416666667)
    pts(93) = Array(30.347847119376, 0.956083333333)
    pts(94) = Array(30.357866836073, 0.720166666667)
    pts(95) = Array(30.365063049061, 0.484250000000)
    pts(96) = Array(30.369438844632, 0.248333333333)
    pts(97) = Array(30.371000000000, 0.000000000000)
    Profile_15 = pts
End Function

Private Function Profile_16() As Variant
    Dim pts(96) As Variant
    pts(0) = Array(30.502000000000, 0.000000000000)
    pts(1) = Array(30.503327686219, 0.138666666667)
    pts(2) = Array(30.507312206387, 0.277333333333)
    pts(3) = Array(30.513957953096, 0.416000000000)
    pts(4) = Array(30.523272274286, 0.554666666667)
    pts(5) = Array(30.535265514052, 0.693333333333)
    pts(6) = Array(30.549951070503, 0.832000000000)
    pts(7) = Array(30.567345471445, 0.970666666667)
    pts(8) = Array(30.585823168183, 1.098666666667)
    pts(9) = Array(30.608485286805, 1.237333333333)
    pts(10) = Array(30.631867486549, 1.365333333333)
    pts(11) = Array(30.659894693047, 1.504000000000)
    pts(12) = Array(30.688283168947, 1.632000000000)
    pts(13) = Array(30.721799237023, 1.770666666667)
    pts(14) = Array(30.755321377478, 1.898666666667)
    pts(15) = Array(30.794479377044, 2.037333333333)
    pts(16) = Array(30.833291240722, 2.165333333333)
    pts(17) = Array(30.874707007106, 2.293333333333)
    pts(18) = Array(30.918774077813, 2.421333333333)
    pts(19) = Array(30.965544161422, 2.549333333333)
    pts(20) = Array(31.015073620389, 2.677333333333)
    pts(21) = Array(31.067423864817, 2.805333333333)
    pts(22) = Array(31.122661800243, 2.933333333333)
    pts(23) = Array(31.180860338053, 3.061333333333)
    pts(24) = Array(31.236877577437, 3.178666666667)
    pts(25) = Array(31.300979003949, 3.306666666667)
    pts(26) = Array(31.362559096881, 3.424000000000)
    pts(27) = Array(31.426918219445, 3.541333333333)
    pts(28) = Array(31.494142223840, 3.658666666667)
    pts(29) = Array(31.564324603994, 3.776000000000)
    pts(30) = Array(31.637567362563, 3.893333333333)
    pts(31) = Array(31.713982015873, 4.010666666667)
    pts(32) = Array(31.793690764338, 4.128000000000)
    pts(33) = Array(31.884560777685, 4.256000000000)
    pts(34) = Array(31.971607120478, 4.373333333333)
    pts(35) = Array(32.070856054864, 4.501333333333)
    pts(36) = Array(32.165964304325, 4.618666666667)
    pts(37) = Array(32.265234810652, 4.736000000000)
    pts(38) = Array(32.368898181842, 4.853333333333)
    pts(39) = Array(32.477213798824, 4.970666666667)
    pts(40) = Array(32.579965702841, 5.077333333333)
    pts(41) = Array(32.687053490494, 5.184000000000)
    pts(42) = Array(32.798758830731, 5.290666666667)
    pts(43) = Array(32.915402181906, 5.397333333333)
    pts(44) = Array(33.024905501169, 5.493333333333)
    pts(45) = Array(33.139011738462, 5.589333333333)
    pts(46) = Array(33.258072126204, 5.685333333333)
    pts(47) = Array(33.382492360559, 5.781333333333)
    pts(48) = Array(33.512745234697, 5.877333333333)
    pts(49) = Array(33.618438220071, 5.952000000000)
    pts(50) = Array(33.712336819622, 6.016000000000)
    pts(51) = Array(33.826048551147, 6.090666666667)
    pts(52) = Array(33.927372925968, 6.154666666667)
    pts(53) = Array(34.032565117249, 6.218666666667)
    pts(54) = Array(34.141963970812, 6.282666666667)
    pts(55) = Array(34.255965277949, 6.346666666667)
    pts(56) = Array(34.354816617530, 6.400000000000)
    pts(57) = Array(34.946291438329, 6.400000000000)
    pts(58) = Array(34.946291438329, 6.255707447925)
    pts(59) = Array(34.892637333901, 6.154666666667)
    pts(60) = Array(34.837202542532, 6.048000000000)
    pts(61) = Array(34.735230026567, 5.845333333333)
    pts(62) = Array(34.683294338104, 5.738666666667)
    pts(63) = Array(34.632534213872, 5.632000000000)
    pts(64) = Array(34.582934763296, 5.525333333333)
    pts(65) = Array(34.534481678256, 5.418666666667)
    pts(66) = Array(34.487161207542, 5.312000000000)
    pts(67) = Array(34.440960132880, 5.205333333333)
    pts(68) = Array(34.395865746408, 5.098666666667)
    pts(69) = Array(34.351865829511, 4.992000000000)
    pts(70) = Array(34.267102857906, 4.778666666667)
    pts(71) = Array(34.226317638779, 4.672000000000)
    pts(72) = Array(34.182666380444, 4.554666666667)
    pts(73) = Array(34.121414577534, 4.384000000000)
    pts(74) = Array(34.062783958876, 4.213333333333)
    pts(75) = Array(34.006737928499, 4.042666666667)
    pts(76) = Array(33.949982580459, 3.861333333333)
    pts(77) = Array(33.899161162853, 3.690666666667)
    pts(78) = Array(33.847886668389, 3.509333333333)
    pts(79) = Array(33.802161803378, 3.338666666667)
    pts(80) = Array(33.756241024044, 3.157333333333)
    pts(81) = Array(33.710576429238, 2.965333333333)
    pts(82) = Array(33.667922152889, 2.773333333333)
    pts(83) = Array(33.628248510424, 2.581333333333)
    pts(84) = Array(33.589574310116, 2.378666666667)
    pts(85) = Array(33.555944482880, 2.186666666667)
    pts(86) = Array(33.523596978917, 1.984000000000)
    pts(87) = Array(33.495915478188, 1.792000000000)
    pts(88) = Array(33.469803748875, 1.589333333333)
    pts(89) = Array(33.447993480447, 1.397333333333)
    pts(90) = Array(33.428045399048, 1.194666666667)
    pts(91) = Array(33.412046556758, 1.002666666667)
    pts(92) = Array(33.398207770242, 0.800000000000)
    pts(93) = Array(33.387976999339, 0.608000000000)
    pts(94) = Array(33.380210176003, 0.405333333333)
    pts(95) = Array(33.375552270702, 0.202666666667)
    pts(96) = Array(33.374000000000, 0.000000000000)
    Profile_16 = pts
End Function

Private Sub AddRetainerSlots(ByVal model As SldWorks.ModelDoc2, ByVal plane As SldWorks.Feature)
    Dim math As SldWorks.MathUtility, sketch As SldWorks.Sketch
    Dim mapping As SldWorks.MathTransform, seg As SldWorks.SketchSegment
    Dim feat As SldWorks.Feature, profile As SldWorks.Feature, cut As SldWorks.Feature
    Dim coords(3, 1) As Double, a As Variant, b As Variant
    Dim side As Long, i As Long, j As Long, found As Long
    Set math = app.GetMathUtility
    currentStage = "Nose retainer: draw two tool slots"
    model.ClearSelection2 True
    If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 520, , "Tool slot plane selection failed."
    model.SketchManager.InsertSketch True
    Set sketch = model.SketchManager.ActiveSketch
    If sketch Is Nothing Then Err.Raise vbObjectError + 521, , "Tool slot sketch missing."
    Set mapping = sketch.ModelToSketchTransform
    model.SketchManager.AddToDB = True
    For side = -1 To 1 Step 2
        coords(0, 0) = 37.046291438329: coords(0, 1) = side * 8.2
        coords(1, 0) = 37.846291438329: coords(1, 1) = side * 8.2
        coords(2, 0) = 37.846291438329: coords(2, 1) = side * 9.7
        coords(3, 0) = 37.046291438329: coords(3, 1) = side * 9.7
        For i = 0 To 3
            j = (i + 1) Mod 4
            a = SketchPoint(math, mapping, coords(i, 0), coords(i, 1))
            b = SketchPoint(math, mapping, coords(j, 0), coords(j, 1))
            Set seg = model.SketchManager.CreateLine(a(0), a(1), 0#, b(0), b(1), 0#)
            If seg Is Nothing Then Err.Raise vbObjectError + 522, , "Tool slot line failed."
            seg.ConstructionGeometry = False
        Next i
    Next side
    model.SketchManager.AddToDB = False
    model.SketchManager.InsertSketch True
    Set feat = model.FirstFeature
    Do While Not feat Is Nothing
        If feat.GetTypeName2 = "ProfileFeature" And feat.Name <> "Meridional_Profile" Then
            Set profile = feat: found = found + 1
        End If
        Set feat = feat.GetNextFeature
    Loop
    If found <> 1 Then Err.Raise vbObjectError + 523, , "Expected one tool slot sketch."
    profile.Name = "NoseRetainer_ToolSlot_Profile"
    model.ClearSelection2 True
    If model.Extension.SelectByID2(profile.Name, "SKETCH", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then Err.Raise vbObjectError + 524, , "Tool slot sketch selection failed."
    currentStage = "Nose retainer: cut two tool slots"
    Set cut = model.FeatureManager.FeatureCut3(False, False, False, 0, 0, 0.0006, 0.0006, False, False, False, False, 0#, 0#, _
        False, False, False, False, False, False, True, False, False, False, 0, 0#, False)
    If cut Is Nothing Then Err.Raise vbObjectError + 525, , "Retainer tool slot cut failed."
    cut.Name = "Two_Spanner_Slots_W1p2_D0p7"
    model.ClearSelection2 True
    model.EditRebuild3
    buildLog.WriteLine "Nose retainer: two opposite spanner slots, width 1.2 mm, axial depth 0.7 mm."
End Sub
