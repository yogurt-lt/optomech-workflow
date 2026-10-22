Attribute VB_Name = "OptomechWideAngleMech_v10"
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
    Dim loaded(7) As SldWorks.ModelDoc2
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
    names = Array("01_FrontBarrel", "02_RearBarrel", "03_G2G3_Spacer", "04_G4G5_Spacer", "05_G6G7_Spacer", "06_FrontRetainer", "07_MiddleRetainer", "08_RearRetainer")
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
    a = SketchPoint(math, mapping, -3.790498622485, 0#)
    b = SketchPoint(math, mapping, 76.559401477824, 0#)
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
    AddEnclosedDetails model, plane, axis, label
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
    If InStr(key, "01_frontbarrel") > 0 Then
        values(0) = 0.18: values(1) = 0.31: values(2) = 0.55
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "02_rearbarrel") > 0 Then
        values(0) = 0.57: values(1) = 0.36: values(2) = 0.17
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "03_g2g3_spacer") > 0 Then
        values(0) = 0.35: values(1) = 0.65: values(2) = 0.36
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "04_g4g5_spacer") > 0 Then
        values(0) = 0.35: values(1) = 0.65: values(2) = 0.36
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "05_g6g7_spacer") > 0 Then
        values(0) = 0.35: values(1) = 0.65: values(2) = 0.36
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "06_frontretainer") > 0 Then
        values(0) = 0.3: values(1) = 0.32: values(2) = 0.35
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "07_middleretainer") > 0 Then
        values(0) = 0.3: values(1) = 0.32: values(2) = 0.35
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "08_rearretainer") > 0 Then
        values(0) = 0.3: values(1) = 0.32: values(2) = 0.35
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
        Case Else: Err.Raise vbObjectError + 418, , "Unknown profile index."
    End Select
End Function

Private Function Profile_0() As Variant
    Dim pts(15) As Variant
    pts(0) = Array(-2.790498622485, 28.000000000000)
    pts(1) = Array(10.682529709373, 28.000000000000)
    pts(2) = Array(22.000000000000, 20.000000000000)
    pts(3) = Array(28.000000000000, 20.000000000000)
    pts(4) = Array(28.000000000000, 17.000000000000)
    pts(5) = Array(38.000000000000, 17.000000000000)
    pts(6) = Array(38.000000000000, 10.500000000000)
    pts(7) = Array(23.853734603604, 10.500000000000)
    pts(8) = Array(23.853734603604, 12.150000000000)
    pts(9) = Array(19.556737253133, 12.150000000000)
    pts(10) = Array(19.556737253133, 15.150000000000)
    pts(11) = Array(10.682529709373, 15.150000000000)
    pts(12) = Array(10.682529709373, 20.150000000000)
    pts(13) = Array(1.409501377515, 20.150000000000)
    pts(14) = Array(1.409501377515, 24.000000000000)
    pts(15) = Array(-2.790498622485, 24.000000000000)
    Profile_0 = pts
End Function

Private Function Profile_1() As Variant
    Dim pts(23) As Variant
    pts(0) = Array(28.000000000000, 21.500000000000)
    pts(1) = Array(39.000000000000, 21.500000000000)
    pts(2) = Array(46.000000000000, 18.000000000000)
    pts(3) = Array(75.559401477824, 18.000000000000)
    pts(4) = Array(75.559401477824, 11.900000000000)
    pts(5) = Array(71.359401477824, 11.900000000000)
    pts(6) = Array(71.359401477824, 8.650000000000)
    pts(7) = Array(66.143526146502, 8.650000000000)
    pts(8) = Array(66.143526146502, 6.924190460673)
    pts(9) = Array(61.486575676992, 6.924190460673)
    pts(10) = Array(61.486575676992, 5.300000000000)
    pts(11) = Array(55.951438715753, 5.300000000000)
    pts(12) = Array(54.760138015529, 4.108699299776)
    pts(13) = Array(53.760138015529, 4.108699299776)
    pts(14) = Array(53.760138015529, 6.000000000000)
    pts(15) = Array(53.505938791537, 6.000000000000)
    pts(16) = Array(53.505938791537, 7.939938124400)
    pts(17) = Array(48.500000000000, 7.939938124400)
    pts(18) = Array(48.500000000000, 9.150000000000)
    pts(19) = Array(41.715324397031, 9.150000000000)
    pts(20) = Array(41.715324397031, 11.900000000000)
    pts(21) = Array(38.000000000000, 11.900000000000)
    pts(22) = Array(38.000000000000, 17.020000000000)
    pts(23) = Array(28.000000000000, 17.020000000000)
    Profile_1 = pts
End Function

Private Function Profile_2() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(15.191517298849, 15.000000000000)
    pts(1) = Array(17.556737253133, 15.000000000000)
    pts(2) = Array(17.556737253133, 10.500000000000)
    pts(3) = Array(15.191517298849, 10.500000000000)
    Profile_2 = pts
End Function

Private Function Profile_3() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(44.178059594266, 9.000000000000)
    pts(1) = Array(47.449380900746, 9.000000000000)
    pts(2) = Array(47.449380900746, 6.000000000000)
    pts(3) = Array(44.178059594266, 6.000000000000)
    Profile_3 = pts
End Function

Private Function Profile_4() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(67.643526146502, 8.500000000000)
    pts(1) = Array(69.983683370336, 8.500000000000)
    pts(2) = Array(69.983683370336, 6.100000000000)
    pts(3) = Array(67.643526146502, 6.100000000000)
    Profile_4 = pts
End Function

Private Function Profile_5() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-1.790498622485, 23.900000000000)
    pts(1) = Array(1.209501377515, 23.900000000000)
    pts(2) = Array(1.209501377515, 18.000000000000)
    pts(3) = Array(-1.790498622485, 18.000000000000)
    Profile_5 = pts
End Function

Private Function Profile_6() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(38.515324397031, 11.800000000000)
    pts(1) = Array(41.515324397031, 11.800000000000)
    pts(2) = Array(41.515324397031, 7.000000000000)
    pts(3) = Array(38.515324397031, 7.000000000000)
    Profile_6 = pts
End Function

Private Function Profile_7() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(71.559401477824, 11.800000000000)
    pts(1) = Array(74.559401477824, 11.800000000000)
    pts(2) = Array(74.559401477824, 7.000000000000)
    pts(3) = Array(71.559401477824, 7.000000000000)
    Profile_7 = pts
End Function

Private Sub AddEnclosedDetails(ByVal model As SldWorks.ModelDoc2, ByVal plane As SldWorks.Feature, ByVal mainAxis As SldWorks.SketchSegment, ByVal label As String)
    Select Case label
        Case "01_FrontBarrel"
            RadialFeatures_0 model, plane, mainAxis
        Case "02_RearBarrel"
            RadialFeatures_1 model, plane, mainAxis
        Case "03_G2G3_Spacer"
        Case "04_G4G5_Spacer"
        Case "05_G6G7_Spacer"
        Case "06_FrontRetainer"
            CapToolSlots model, plane, -1.840498622485, -0.290498622485, 21.900000000000, 23.950000000000
        Case "07_MiddleRetainer"
            CapToolSlots model, plane, 38.465324397031, 40.015324397031, 9.800000000000, 11.850000000000
        Case "08_RearRetainer"
            CapToolSlots model, plane, 73.059401477824, 74.609401477824, 9.800000000000, 11.850000000000
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
    cut.Name = "Two_Wrench_Slots_W2_R2_A1p5"
    model.ClearSelection2 True
    model.EditRebuild3
End Sub


Private Sub RadialFeatures_0(ByVal model As SldWorks.ModelDoc2, ByVal basePlane As SldWorks.Feature, ByVal mainAxis As SldWorks.SketchSegment)
    Dim plane As SldWorks.Feature, feature As SldWorks.Feature, profile As SldWorks.Feature
    Dim ref As SldWorks.RefPlane, created As SldWorks.RefPlane
    Dim tf As SldWorks.MathTransform, mapping As SldWorks.MathTransform
    Dim sketch As SldWorks.Sketch, portCircle As SldWorks.SketchSegment, cut As SldWorks.Feature
    Dim mp As SldWorks.MathPoint
    Dim math As SldWorks.MathUtility, sd As SldWorks.SelectData
    Dim normal As Variant, baseNormal As Variant, zs As Variant, inputCheck As Variant
    Dim position(2) As Double
    Dim row As Long, direction As Long, found As Long, reverseCut As Boolean
    Dim dot As Double, p As Variant
    Set math = app.GetMathUtility
    Set ref = basePlane.GetSpecificFeature2
    Set tf = ref.Transform: baseNormal = tf.ArrayData
    zs = Array(0.033000000000, 0.033000000000, 0.033000000000)
    For direction = 0 To 2
        currentStage = "Radial holes: plane " & CStr(direction)
        If direction = 0 Then
            Set plane = basePlane
            Set ref = plane.GetSpecificFeature2
        Else
            model.ClearSelection2 True
            If basePlane.Select2(False, 0) = False Then Err.Raise vbObjectError + 501, , "Adhesive plane base selection failed."
            Set sd = model.SelectionManager.CreateSelectData: sd.Mark = 1
            If mainAxis.Select4(True, sd) = False Then Err.Raise vbObjectError + 502, , "Adhesive plane axis selection failed."
            Set created = model.FeatureManager.InsertRefPlane(16, direction * 1.047197551196598, 4, 0#, 0, 0#)
            If created Is Nothing Then Err.Raise vbObjectError + 503, , "Adhesive reference plane failed."
            Set feature = model.FirstFeature: Set plane = Nothing
            Do While Not feature Is Nothing
                If feature.GetTypeName2 = "RefPlane" Then Set plane = feature
                Set feature = feature.GetNextFeature
            Loop
            If plane Is Nothing Then Err.Raise vbObjectError + 504, , "Adhesive plane feature missing."
            plane.Name = "RadialPlane_0_" & CStr(direction * 60)
            Set ref = plane.GetSpecificFeature2
        End If
        Set tf = ref.Transform: normal = tf.ArrayData
        dot = normal(6) * baseNormal(6) + normal(7) * baseNormal(7) + normal(8) * baseNormal(8)
        If direction = 0 Then
            reverseCut = True
        Else
            If Abs(Abs(dot) - 0.5) > 0.0001 Then Err.Raise vbObjectError + 505, , "Adhesive plane angle verification failed."
            reverseCut = (dot < 0#)
        End If
        buildLog.WriteLine "Radial direction=" & CStr(direction) & "; normal dot base=" & CStr(dot) & "; reverse=" & CStr(reverseCut)
        For row = 0 To 0
            currentStage = "Radial holes: sketch direction=" & CStr(direction) & "; row=" & CStr(row)
            model.ClearSelection2 True
            If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 506, , "Adhesive sketch plane selection failed."
            model.SketchManager.InsertSketch True
            Set sketch = model.SketchManager.ActiveSketch
            If sketch Is Nothing Then Err.Raise vbObjectError + 507, , "Adhesive sketch missing."
            Set mapping = sketch.ModelToSketchTransform
            model.SketchManager.AddToDB = True
            ' CreatePoint requires SAFEARRAY(Double), not VBA Array() of Variants.
            position(0) = 0#: position(1) = 0#: position(2) = CDbl(zs(row))
            Set mp = math.CreatePoint(position)
            If mp Is Nothing Then Err.Raise vbObjectError + 513, , "Port math point creation failed."
            inputCheck = mp.ArrayData
            If Abs(inputCheck(0) - position(0)) > 0.00000001 Or Abs(inputCheck(1) - position(1)) > 0.00000001 Or Abs(inputCheck(2) - position(2)) > 0.00000001 Then Err.Raise vbObjectError + 514, , "Port input coordinate readback mismatch."
            Set mp = mp.MultiplyTransform(mapping)
            p = mp.ArrayData
            If Abs(p(2)) > 0.00000001 Then Err.Raise vbObjectError + 512, , "Port center is not on the selected plane."
            If Abs(Sqr(p(0) * p(0) + p(1) * p(1)) - Abs(position(2))) > 0.00000001 Then Err.Raise vbObjectError + 515, , "Port transformed coordinate radius mismatch."
            Set portCircle = model.SketchManager.CreateCircleByRadius(p(0), p(1), 0#, 0.001250000000)
            If portCircle Is Nothing Then Err.Raise vbObjectError + 508, , "Adhesive circle creation failed."
            portCircle.ConstructionGeometry = False
            model.SketchManager.AddToDB = False
            model.SketchManager.InsertSketch True
            model.ClearSelection2 True
            model.EditRebuild3
            Set feature = model.FirstFeature: Set profile = Nothing: found = 0
            Do While Not feature Is Nothing
                If feature.GetTypeName2 = "ProfileFeature" Then
                    If feature.Name <> "Meridional_Profile" And Left$(feature.Name, 9) <> "GluePort_" Then
                        Set profile = feature: found = found + 1
                    End If
                End If
                Set feature = feature.GetNextFeature
            Loop
            If found <> 1 Then Err.Raise vbObjectError + 509, , "Expected one new adhesive profile."
            profile.Name = "GluePort_Group0_Profile_" & CStr(direction) & "_" & CStr(row)
            model.ClearSelection2 True
            If model.Extension.SelectByID2(profile.Name, "SKETCH", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then Err.Raise vbObjectError + 510, , "Adhesive cut sketch selection failed."
            currentStage = "Radial holes: cut direction=" & CStr(direction) & "; row=" & CStr(row)
            buildLog.WriteLine currentStage & "; Z_m=" & CStr(zs(row)) & "; local=" & CStr(p(0)) & "," & CStr(p(1)) & "; construction=" & CStr(portCircle.ConstructionGeometry)
            ' One circular contour, through all in ONE direction: only the selected tube wall.
            ' All bodies are in scope; no assembly scope or sheet-metal options.
            Set cut = model.FeatureManager.FeatureCut3(True, False, reverseCut, 1, 0, 0.04, 0#, False, False, False, False, 0#, 0#, _
                False, False, False, False, False, False, True, False, False, False, 0, 0#, False)
            If cut Is Nothing Then Err.Raise vbObjectError + 511, , "Adhesive port cut failed: direction=" & CStr(direction) & "; row=" & CStr(row)
            cut.Name = "RadialGroup0_" & CStr(direction + 1) & "_" & CStr(row + 1)
            model.ClearSelection2 True
            model.EditRebuild3
            buildLog.WriteLine "Created " & cut.Name
        Next row
    Next direction
    buildLog.WriteLine "3 native radial holes created; diameter_mm=2.5."
End Sub


Private Sub RadialFeatures_1(ByVal model As SldWorks.ModelDoc2, ByVal basePlane As SldWorks.Feature, ByVal mainAxis As SldWorks.SketchSegment)
    Dim plane As SldWorks.Feature, feature As SldWorks.Feature, profile As SldWorks.Feature
    Dim ref As SldWorks.RefPlane, created As SldWorks.RefPlane
    Dim tf As SldWorks.MathTransform, mapping As SldWorks.MathTransform
    Dim sketch As SldWorks.Sketch, portCircle As SldWorks.SketchSegment, cut As SldWorks.Feature
    Dim mp As SldWorks.MathPoint
    Dim math As SldWorks.MathUtility, sd As SldWorks.SelectData
    Dim normal As Variant, baseNormal As Variant, zs As Variant, inputCheck As Variant
    Dim position(2) As Double
    Dim row As Long, direction As Long, found As Long, reverseCut As Boolean
    Dim dot As Double, p As Variant
    Set math = app.GetMathUtility
    Set ref = basePlane.GetSpecificFeature2
    Set tf = ref.Transform: baseNormal = tf.ArrayData
    zs = Array(0.033000000000, 0.033000000000, 0.033000000000)
    For direction = 0 To 2
        currentStage = "Radial holes: plane " & CStr(direction)
        If direction = 0 Then
            Set plane = basePlane
            Set ref = plane.GetSpecificFeature2
        Else
            model.ClearSelection2 True
            If basePlane.Select2(False, 0) = False Then Err.Raise vbObjectError + 501, , "Adhesive plane base selection failed."
            Set sd = model.SelectionManager.CreateSelectData: sd.Mark = 1
            If mainAxis.Select4(True, sd) = False Then Err.Raise vbObjectError + 502, , "Adhesive plane axis selection failed."
            Set created = model.FeatureManager.InsertRefPlane(16, direction * 1.047197551196598, 4, 0#, 0, 0#)
            If created Is Nothing Then Err.Raise vbObjectError + 503, , "Adhesive reference plane failed."
            Set feature = model.FirstFeature: Set plane = Nothing
            Do While Not feature Is Nothing
                If feature.GetTypeName2 = "RefPlane" Then Set plane = feature
                Set feature = feature.GetNextFeature
            Loop
            If plane Is Nothing Then Err.Raise vbObjectError + 504, , "Adhesive plane feature missing."
            plane.Name = "RadialPlane_1_" & CStr(direction * 60)
            Set ref = plane.GetSpecificFeature2
        End If
        Set tf = ref.Transform: normal = tf.ArrayData
        dot = normal(6) * baseNormal(6) + normal(7) * baseNormal(7) + normal(8) * baseNormal(8)
        If direction = 0 Then
            reverseCut = True
        Else
            If Abs(Abs(dot) - 0.5) > 0.0001 Then Err.Raise vbObjectError + 505, , "Adhesive plane angle verification failed."
            reverseCut = (dot < 0#)
        End If
        buildLog.WriteLine "Radial direction=" & CStr(direction) & "; normal dot base=" & CStr(dot) & "; reverse=" & CStr(reverseCut)
        For row = 0 To 0
            currentStage = "Radial holes: sketch direction=" & CStr(direction) & "; row=" & CStr(row)
            model.ClearSelection2 True
            If plane.Select2(False, 0) = False Then Err.Raise vbObjectError + 506, , "Adhesive sketch plane selection failed."
            model.SketchManager.InsertSketch True
            Set sketch = model.SketchManager.ActiveSketch
            If sketch Is Nothing Then Err.Raise vbObjectError + 507, , "Adhesive sketch missing."
            Set mapping = sketch.ModelToSketchTransform
            model.SketchManager.AddToDB = True
            ' CreatePoint requires SAFEARRAY(Double), not VBA Array() of Variants.
            position(0) = 0#: position(1) = 0#: position(2) = CDbl(zs(row))
            Set mp = math.CreatePoint(position)
            If mp Is Nothing Then Err.Raise vbObjectError + 513, , "Port math point creation failed."
            inputCheck = mp.ArrayData
            If Abs(inputCheck(0) - position(0)) > 0.00000001 Or Abs(inputCheck(1) - position(1)) > 0.00000001 Or Abs(inputCheck(2) - position(2)) > 0.00000001 Then Err.Raise vbObjectError + 514, , "Port input coordinate readback mismatch."
            Set mp = mp.MultiplyTransform(mapping)
            p = mp.ArrayData
            If Abs(p(2)) > 0.00000001 Then Err.Raise vbObjectError + 512, , "Port center is not on the selected plane."
            If Abs(Sqr(p(0) * p(0) + p(1) * p(1)) - Abs(position(2))) > 0.00000001 Then Err.Raise vbObjectError + 515, , "Port transformed coordinate radius mismatch."
            Set portCircle = model.SketchManager.CreateCircleByRadius(p(0), p(1), 0#, 0.001650000000)
            If portCircle Is Nothing Then Err.Raise vbObjectError + 508, , "Adhesive circle creation failed."
            portCircle.ConstructionGeometry = False
            model.SketchManager.AddToDB = False
            model.SketchManager.InsertSketch True
            model.ClearSelection2 True
            model.EditRebuild3
            Set feature = model.FirstFeature: Set profile = Nothing: found = 0
            Do While Not feature Is Nothing
                If feature.GetTypeName2 = "ProfileFeature" Then
                    If feature.Name <> "Meridional_Profile" And Left$(feature.Name, 9) <> "GluePort_" Then
                        Set profile = feature: found = found + 1
                    End If
                End If
                Set feature = feature.GetNextFeature
            Loop
            If found <> 1 Then Err.Raise vbObjectError + 509, , "Expected one new adhesive profile."
            profile.Name = "GluePort_Group1_Profile_" & CStr(direction) & "_" & CStr(row)
            model.ClearSelection2 True
            If model.Extension.SelectByID2(profile.Name, "SKETCH", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then Err.Raise vbObjectError + 510, , "Adhesive cut sketch selection failed."
            currentStage = "Radial holes: cut direction=" & CStr(direction) & "; row=" & CStr(row)
            buildLog.WriteLine currentStage & "; Z_m=" & CStr(zs(row)) & "; local=" & CStr(p(0)) & "," & CStr(p(1)) & "; construction=" & CStr(portCircle.ConstructionGeometry)
            ' One circular contour, through all in ONE direction: only the selected tube wall.
            ' All bodies are in scope; no assembly scope or sheet-metal options.
            Set cut = model.FeatureManager.FeatureCut3(True, False, reverseCut, 1, 0, 0.04, 0#, False, False, False, False, 0#, 0#, _
                False, False, False, False, False, False, True, False, False, False, 0, 0#, False)
            If cut Is Nothing Then Err.Raise vbObjectError + 511, , "Adhesive port cut failed: direction=" & CStr(direction) & "; row=" & CStr(row)
            cut.Name = "RadialGroup1_" & CStr(direction + 1) & "_" & CStr(row + 1)
            model.ClearSelection2 True
            model.EditRebuild3
            buildLog.WriteLine "Created " & cut.Name
        Next row
    Next direction
    buildLog.WriteLine "3 native radial holes created; diameter_mm=3.3."
End Sub
