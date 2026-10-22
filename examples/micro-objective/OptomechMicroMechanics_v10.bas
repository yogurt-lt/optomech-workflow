Attribute VB_Name = "OptomechMicroMechanics_v10"
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
    Dim loaded(6) As SldWorks.ModelDoc2
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
    names = Array("01_ObjectiveBarrel", "02_G1_CellAdapter", "03_G2_CellAdapter", "04_G3_CellAdapter", "05_G4_CellAdapter", "06_G5_CellAdapter", "07_NoseRetainer")
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
