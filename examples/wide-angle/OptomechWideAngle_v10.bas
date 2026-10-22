Attribute VB_Name = "OptomechWideAngle_v10"
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
    Dim loaded(17) As SldWorks.ModelDoc2
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
    names = Array("01_FrontBarrel", "02_RearBarrel", "03_G2G3_Spacer", "04_G4G5_Spacer", "05_G6G7_Spacer", "06_FrontRetainer", "07_MiddleRetainer", "08_RearRetainer", "Lens1_Optical_Reference", "Lens2_RevisedBlank_Reference", "Lens3_RevisedBlank_Reference", "Lens4_RevisedBlank_Reference", "Lens5_RevisedBlank_Reference", "Lens6_Optical_Reference", "Lens7_Optical_Reference", "Lens8_Optical_Reference", "Lens9_Optical_Reference", "Lens10_RevisedBlank_Reference")
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

Private Function Profile_8() As Variant
    Dim pts(124) As Variant
    pts(0) = Array(0.000000000000, 0.000000000000)
    pts(1) = Array(0.001551598167, 0.646153846154)
    pts(2) = Array(0.005914464416, 1.261538461538)
    pts(3) = Array(0.013525189932, 1.907692307692)
    pts(4) = Array(0.023659417912, 2.523076923077)
    pts(5) = Array(0.037331278905, 3.169230769231)
    pts(6) = Array(0.053239543287, 3.784615384615)
    pts(7) = Array(0.071965840769, 4.400000000000)
    pts(8) = Array(0.093511348839, 5.015384615385)
    pts(9) = Array(0.119169793391, 5.661538461538)
    pts(10) = Array(0.146499116334, 6.276923076923)
    pts(11) = Array(0.178234165078, 6.923076923077)
    pts(12) = Array(0.211354540333, 7.538461538462)
    pts(13) = Array(0.249174604836, 8.184615384615)
    pts(14) = Array(0.288094819640, 8.800000000000)
    pts(15) = Array(0.329847882562, 9.415384615385)
    pts(16) = Array(0.374436438312, 10.030769230769)
    pts(17) = Array(0.424309217829, 10.676923076923)
    pts(18) = Array(0.474719576260, 11.292307692308)
    pts(19) = Array(0.527974629507, 11.907692307692)
    pts(20) = Array(0.584077770246, 12.523076923077)
    pts(21) = Array(0.646055241895, 13.169230769231)
    pts(22) = Array(0.708008360013, 13.784615384615)
    pts(23) = Array(0.772821074902, 14.400000000000)
    pts(24) = Array(0.840497545118, 15.015384615385)
    pts(25) = Array(0.914644718878, 15.661538461538)
    pts(26) = Array(0.988205701610, 16.276923076923)
    pts(27) = Array(1.064644317385, 16.892307692308)
    pts(28) = Array(1.143965512871, 17.507692307692)
    pts(29) = Array(1.230360777099, 18.153846153846)
    pts(30) = Array(1.315607565322, 18.769230769231)
    pts(31) = Array(1.403753254095, 19.384615384615)
    pts(32) = Array(1.494803605012, 20.000000000000)
    pts(33) = Array(10.682529709373, 20.000000000000)
    pts(34) = Array(10.682529709373, 11.800000000000)
    pts(35) = Array(10.591026373579, 11.753846153846)
    pts(36) = Array(10.471697382920, 11.692307692308)
    pts(37) = Array(10.355265042419, 11.630769230769)
    pts(38) = Array(10.241569646473, 11.569230769231)
    pts(39) = Array(10.075845147051, 11.476923076923)
    pts(40) = Array(9.968386573861, 11.415384615385)
    pts(41) = Array(9.863211616806, 11.353846153846)
    pts(42) = Array(9.760214574585, 11.292307692308)
    pts(43) = Array(9.659298166207, 11.230769230769)
    pts(44) = Array(9.560372617179, 11.169230769231)
    pts(45) = Array(9.463354869684, 11.107692307692)
    pts(46) = Array(9.274740104028, 10.984615384615)
    pts(47) = Array(9.092899766575, 10.861538461538)
    pts(48) = Array(8.917351366730, 10.738461538462)
    pts(49) = Array(8.831802346787, 10.676923076923)
    pts(50) = Array(8.706123951739, 10.584615384615)
    pts(51) = Array(8.543259587508, 10.461538461538)
    pts(52) = Array(8.463753319424, 10.400000000000)
    pts(53) = Array(8.346796445292, 10.307692307692)
    pts(54) = Array(8.194972789294, 10.184615384615)
    pts(55) = Array(8.120754732439, 10.123076923077)
    pts(56) = Array(8.011462150832, 10.030769230769)
    pts(57) = Array(7.939916718681, 9.969230769231)
    pts(58) = Array(7.834508371379, 9.876923076923)
    pts(59) = Array(7.663720925919, 9.723076923077)
    pts(60) = Array(7.498751789465, 9.569230769231)
    pts(61) = Array(7.339289242771, 9.415384615385)
    pts(62) = Array(7.185052798391, 9.261538461538)
    pts(63) = Array(7.094913010551, 9.169230769231)
    pts(64) = Array(7.006512358321, 9.076923076923)
    pts(65) = Array(6.862914148459, 8.923076923077)
    pts(66) = Array(6.778924669370, 8.830769230769)
    pts(67) = Array(6.696511397503, 8.738461538462)
    pts(68) = Array(6.562553271402, 8.584615384615)
    pts(69) = Array(6.484157271246, 8.492307692308)
    pts(70) = Array(6.407203345196, 8.400000000000)
    pts(71) = Array(6.282063852695, 8.246153846154)
    pts(72) = Array(6.208800727386, 8.153846153846)
    pts(73) = Array(6.136867787185, 8.061538461538)
    pts(74) = Array(6.066237721734, 7.969230769231)
    pts(75) = Array(5.996884438058, 7.876923076923)
    pts(76) = Array(5.861909504567, 7.692307692308)
    pts(77) = Array(5.731755987799, 7.507692307692)
    pts(78) = Array(5.606252287106, 7.323076923077)
    pts(79) = Array(5.485240562828, 7.138461538462)
    pts(80) = Array(5.426373277518, 7.046153846154)
    pts(81) = Array(5.349544101298, 6.923076923077)
    pts(82) = Array(5.237780639023, 6.738461538462)
    pts(83) = Array(5.183431705938, 6.646153846154)
    pts(84) = Array(5.112523273680, 6.523076923077)
    pts(85) = Array(5.009430501952, 6.338461538462)
    pts(86) = Array(4.959327224899, 6.246153846154)
    pts(87) = Array(4.893991512089, 6.123076923077)
    pts(88) = Array(4.799078598692, 5.938461538462)
    pts(89) = Array(4.752988708830, 5.846153846154)
    pts(90) = Array(4.692928807253, 5.723076923077)
    pts(91) = Array(4.605776125959, 5.538461538462)
    pts(92) = Array(4.563500702078, 5.446153846154)
    pts(93) = Array(4.508462020522, 5.323076923077)
    pts(94) = Array(4.468168099496, 5.230769230769)
    pts(95) = Array(4.415740368125, 5.107692307692)
    pts(96) = Array(4.377382237132, 5.015384615385)
    pts(97) = Array(4.327506881615, 4.892307692308)
    pts(98) = Array(4.243665234730, 4.676923076923)
    pts(99) = Array(4.164126434299, 4.461538461538)
    pts(100) = Array(4.088808264017, 4.246153846154)
    pts(101) = Array(4.017634790656, 4.030769230769)
    pts(102) = Array(3.950535924936, 3.815384615385)
    pts(103) = Array(3.887447030491, 3.600000000000)
    pts(104) = Array(3.828308575278, 3.384615384615)
    pts(105) = Array(3.773065820652, 3.169230769231)
    pts(106) = Array(3.721668543958, 2.953846153846)
    pts(107) = Array(3.694007358387, 2.830769230769)
    pts(108) = Array(3.667578784870, 2.707692307692)
    pts(109) = Array(3.624272258420, 2.492307692308)
    pts(110) = Array(3.584680163490, 2.276923076923)
    pts(111) = Array(3.548768636168, 2.061538461538)
    pts(112) = Array(3.529888438771, 1.938461538462)
    pts(113) = Array(3.512194916906, 1.815384615385)
    pts(114) = Array(3.484072403174, 1.600000000000)
    pts(115) = Array(3.459546237910, 1.384615384615)
    pts(116) = Array(3.438596280608, 1.169230769231)
    pts(117) = Array(3.428223822089, 1.046153846154)
    pts(118) = Array(3.419010723283, 0.923076923077)
    pts(119) = Array(3.405670364773, 0.707692307692)
    pts(120) = Array(3.399633890202, 0.584615384615)
    pts(121) = Array(3.394749244920, 0.461538461538)
    pts(122) = Array(3.388969025526, 0.246153846154)
    pts(123) = Array(3.387246220486, 0.123076923077)
    pts(124) = Array(3.386672002143, 0.000000000000)
    Profile_8 = pts
End Function

Private Function Profile_9() As Variant
    Dim pts(100) As Variant
    pts(0) = Array(10.160862462816, 0.000000000000)
    pts(1) = Array(10.161627711532, 0.415384615385)
    pts(2) = Array(10.163755791598, 0.807692307692)
    pts(3) = Array(10.167249142198, 1.200000000000)
    pts(4) = Array(10.172107890259, 1.592307692308)
    pts(5) = Array(10.178740880051, 2.007692307692)
    pts(6) = Array(10.186411351695, 2.400000000000)
    pts(7) = Array(10.195447917322, 2.792307692308)
    pts(8) = Array(10.205850905561, 3.184615384615)
    pts(9) = Array(10.218355609057, 3.600000000000)
    pts(10) = Array(10.231573067444, 3.992307692308)
    pts(11) = Array(10.246158263255, 4.384615384615)
    pts(12) = Array(10.262111727810, 4.776923076923)
    pts(13) = Array(10.280495644086, 5.192307692308)
    pts(14) = Array(10.299268019350, 5.584615384615)
    pts(15) = Array(10.319410600409, 5.976923076923)
    pts(16) = Array(10.340924122872, 6.369230769231)
    pts(17) = Array(10.365198299974, 6.784615384615)
    pts(18) = Array(10.389536881063, 7.176923076923)
    pts(19) = Array(10.415248968684, 7.569230769231)
    pts(20) = Array(10.442335504954, 7.961538461538)
    pts(21) = Array(10.470797483285, 8.353846153846)
    pts(22) = Array(10.500635948565, 8.746153846154)
    pts(23) = Array(10.531851997352, 9.138461538462)
    pts(24) = Array(10.564446778077, 9.530769230769)
    pts(25) = Array(10.600463003984, 9.946153846154)
    pts(26) = Array(10.635900188777, 10.338461538462)
    pts(27) = Array(10.672719942647, 10.730769230769)
    pts(28) = Array(10.710923626862, 11.123076923077)
    pts(29) = Array(10.750512655746, 11.515384615385)
    pts(30) = Array(10.791488496940, 11.907692307692)
    pts(31) = Array(10.833852671684, 12.300000000000)
    pts(32) = Array(10.878600812516, 12.701078078520)
    pts(33) = Array(10.682529709373, 14.000000000000)
    pts(34) = Array(10.682529709373, 15.000000000000)
    pts(35) = Array(15.191517298849, 15.000000000000)
    pts(36) = Array(15.191517298849, 10.668736294680)
    pts(37) = Array(15.069722589258, 10.523076923077)
    pts(38) = Array(14.956547909945, 10.384615384615)
    pts(39) = Array(14.845841193335, 10.246153846154)
    pts(40) = Array(14.737539702984, 10.107692307692)
    pts(41) = Array(14.631584106779, 9.969230769231)
    pts(42) = Array(14.527918231813, 9.830769230769)
    pts(43) = Array(14.426488841867, 9.692307692308)
    pts(44) = Array(14.310913990585, 9.530769230769)
    pts(45) = Array(14.214160460548, 9.392307692308)
    pts(46) = Array(14.103913609077, 9.230769230769)
    pts(47) = Array(14.011623580437, 9.092307692308)
    pts(48) = Array(13.906470023201, 8.930769230769)
    pts(49) = Array(13.818452757965, 8.792307692308)
    pts(50) = Array(13.718181312076, 8.630769230769)
    pts(51) = Array(13.620458530527, 8.469230769231)
    pts(52) = Array(13.525232079896, 8.307692307692)
    pts(53) = Array(13.432452333265, 8.146153846154)
    pts(54) = Array(13.342072196237, 7.984615384615)
    pts(55) = Array(13.254046947709, 7.823076923077)
    pts(56) = Array(13.168334093896, 7.661538461538)
    pts(57) = Array(13.084893234323, 7.500000000000)
    pts(58) = Array(13.003685938594, 7.338461538462)
    pts(59) = Array(12.924675632939, 7.176923076923)
    pts(60) = Array(12.847827495599, 7.015384615385)
    pts(61) = Array(12.773108360259, 6.853846153846)
    pts(62) = Array(12.700486626794, 6.692307692308)
    pts(63) = Array(12.629932178670, 6.530769230769)
    pts(64) = Array(12.551793126783, 6.346153846154)
    pts(65) = Array(12.485573612122, 6.184615384615)
    pts(66) = Array(12.421335568148, 6.023076923077)
    pts(67) = Array(12.359054268993, 5.861538461538)
    pts(68) = Array(12.290241455677, 5.676923076923)
    pts(69) = Array(12.232075173023, 5.515384615385)
    pts(70) = Array(12.167907915160, 5.330769230769)
    pts(71) = Array(12.113758777542, 5.169230769231)
    pts(72) = Array(12.054130791234, 4.984615384615)
    pts(73) = Array(12.003910457754, 4.823076923077)
    pts(74) = Array(11.948725865786, 4.638461538462)
    pts(75) = Array(11.902354706407, 4.476923076923)
    pts(76) = Array(11.851527168226, 4.292307692308)
    pts(77) = Array(11.808933556613, 4.130769230769)
    pts(78) = Array(11.762385517305, 3.946153846154)
    pts(79) = Array(11.723505220137, 3.784615384615)
    pts(80) = Array(11.681167261750, 3.600000000000)
    pts(81) = Array(11.645942911300, 3.438461538462)
    pts(82) = Array(11.607753192809, 3.253846153846)
    pts(83) = Array(11.571753049804, 3.069230769231)
    pts(84) = Array(11.537927943187, 2.884615384615)
    pts(85) = Array(11.510104513453, 2.723076923077)
    pts(86) = Array(11.480321856798, 2.538461538462)
    pts(87) = Array(11.452677885082, 2.353846153846)
    pts(88) = Array(11.427161767198, 2.169230769231)
    pts(89) = Array(11.406572839131, 2.007692307692)
    pts(90) = Array(11.385020367011, 1.823076923077)
    pts(91) = Array(11.365569546838, 1.638461538462)
    pts(92) = Array(11.348212925537, 1.453846153846)
    pts(93) = Array(11.334738554309, 1.292307692308)
    pts(94) = Array(11.321291354927, 1.107692307692)
    pts(95) = Array(11.309921519059, 0.923076923077)
    pts(96) = Array(11.300624752831, 0.738461538462)
    pts(97) = Array(11.293397554064, 0.553846153846)
    pts(98) = Array(11.288237205655, 0.369230769231)
    pts(99) = Array(11.285141770453, 0.184615384615)
    pts(100) = Array(11.284110087621, 0.000000000000)
    Profile_9 = pts
End Function

Private Function Profile_10() As Variant
    Dim pts(86) As Variant
    pts(0) = Array(14.974659932496, 0.000000000000)
    pts(1) = Array(14.975172383126, 0.147692307692)
    pts(2) = Array(14.976709809055, 0.295384615385)
    pts(3) = Array(14.982355997806, 0.572307692308)
    pts(4) = Array(14.991609376522, 0.849230769231)
    pts(5) = Array(15.004474656715, 1.126153846154)
    pts(6) = Array(15.012815197308, 1.273846153846)
    pts(7) = Array(15.022186171376, 1.421538461538)
    pts(8) = Array(15.042538956039, 1.698461538462)
    pts(9) = Array(15.066529721264, 1.975384615385)
    pts(10) = Array(15.094170838831, 2.252307692308)
    pts(11) = Array(15.110410222853, 2.400000000000)
    pts(12) = Array(15.127694373215, 2.547692307692)
    pts(13) = Array(15.162927185762, 2.824615384615)
    pts(14) = Array(15.201860631340, 3.101538461538)
    pts(15) = Array(15.244515279397, 3.378461538462)
    pts(16) = Array(15.290913840787, 3.655384615385)
    pts(17) = Array(15.341081230139, 3.932307692308)
    pts(18) = Array(15.395044635237, 4.209230769231)
    pts(19) = Array(15.452833593867, 4.486153846154)
    pts(20) = Array(15.493500559804, 4.670769230769)
    pts(21) = Array(15.535892023345, 4.855384615385)
    pts(22) = Array(15.580018590570, 5.040000000000)
    pts(23) = Array(15.625891394268, 5.224615384615)
    pts(24) = Array(15.673522108997, 5.409230769231)
    pts(25) = Array(15.722922967100, 5.593846153846)
    pts(26) = Array(15.774106775724, 5.778461538462)
    pts(27) = Array(15.827086934909, 5.963076923077)
    pts(28) = Array(15.881877456806, 6.147692307692)
    pts(29) = Array(15.938492986092, 6.332307692308)
    pts(30) = Array(15.996948821668, 6.516923076923)
    pts(31) = Array(16.051145725483, 6.683076923077)
    pts(32) = Array(16.113142736836, 6.867692307692)
    pts(33) = Array(16.177028340222, 7.052307692308)
    pts(34) = Array(16.242820620021, 7.236923076923)
    pts(35) = Array(16.303679456501, 7.403076923077)
    pts(36) = Array(16.373147069373, 7.587692307692)
    pts(37) = Array(16.437346246220, 7.753846153846)
    pts(38) = Array(16.510563005565, 7.938461538462)
    pts(39) = Array(16.578171662011, 8.104615384615)
    pts(40) = Array(16.655217520715, 8.289230769231)
    pts(41) = Array(16.726310546792, 8.455384615385)
    pts(42) = Array(16.807272113813, 8.640000000000)
    pts(43) = Array(16.881930644604, 8.806153846154)
    pts(44) = Array(16.966901769161, 8.990769230769)
    pts(45) = Array(17.045213744394, 9.156923076923)
    pts(46) = Array(17.125289107117, 9.323076923077)
    pts(47) = Array(17.207150093206, 9.489230769231)
    pts(48) = Array(17.290819822562, 9.655384615385)
    pts(49) = Array(17.376322340288, 9.821538461538)
    pts(50) = Array(17.463682660582, 9.987692307692)
    pts(51) = Array(17.556737253133, 10.160862462820)
    pts(52) = Array(17.556737253133, 12.000000000000)
    pts(53) = Array(18.674638214572, 12.000000000000)
    pts(54) = Array(18.674638214572, 10.160862462820)
    pts(55) = Array(18.735244758896, 9.803076923077)
    pts(56) = Array(18.795450655116, 9.433846153846)
    pts(57) = Array(18.850451828057, 9.083076923077)
    pts(58) = Array(18.906044547073, 8.713846153846)
    pts(59) = Array(18.956675501140, 8.363076923077)
    pts(60) = Array(19.007680361513, 7.993846153846)
    pts(61) = Array(19.056341561063, 7.624615384615)
    pts(62) = Array(19.102664866896, 7.255384615385)
    pts(63) = Array(19.144511521096, 6.904615384615)
    pts(64) = Array(19.186291408957, 6.535384615385)
    pts(65) = Array(19.225748711721, 6.166153846154)
    pts(66) = Array(19.262888051156, 5.796923076923)
    pts(67) = Array(19.297713766320, 5.427692307692)
    pts(68) = Array(19.330229916135, 5.058461538462)
    pts(69) = Array(19.360440281776, 4.689230769231)
    pts(70) = Array(19.388348368894, 4.320000000000)
    pts(71) = Array(19.406515915913, 4.061538461538)
    pts(72) = Array(19.424732212865, 3.784615384615)
    pts(73) = Array(19.440569440402, 3.526153846154)
    pts(74) = Array(19.456291100338, 3.249230769231)
    pts(75) = Array(19.469801849940, 2.990769230769)
    pts(76) = Array(19.483032682341, 2.713846153846)
    pts(77) = Array(19.494976449619, 2.436923076923)
    pts(78) = Array(19.505633923887, 2.160000000000)
    pts(79) = Array(19.514420986286, 1.901538461538)
    pts(80) = Array(19.522593505894, 1.624615384615)
    pts(81) = Array(19.529062268007, 1.366153846154)
    pts(82) = Array(19.534751803945, 1.089230769231)
    pts(83) = Array(19.538903825097, 0.830769230769)
    pts(84) = Array(19.542111748294, 0.553846153846)
    pts(85) = Array(19.544036419564, 0.276923076923)
    pts(86) = Array(19.544677962881, 0.000000000000)
    Profile_10 = pts
End Function

Private Function Profile_11() As Variant
    Dim pts(99) As Variant
    pts(0) = Array(19.544677962881, 0.000000000000)
    pts(1) = Array(19.544036419564, 0.276923076923)
    pts(2) = Array(19.542111748294, 0.553846153846)
    pts(3) = Array(19.538903825097, 0.830769230769)
    pts(4) = Array(19.534751803945, 1.089230769231)
    pts(5) = Array(19.529062268007, 1.366153846154)
    pts(6) = Array(19.522593505894, 1.624615384615)
    pts(7) = Array(19.514420986286, 1.901538461538)
    pts(8) = Array(19.505633923887, 2.160000000000)
    pts(9) = Array(19.494976449619, 2.436923076923)
    pts(10) = Array(19.483032682341, 2.713846153846)
    pts(11) = Array(19.469801849940, 2.990769230769)
    pts(12) = Array(19.456291100338, 3.249230769231)
    pts(13) = Array(19.440569440402, 3.526153846154)
    pts(14) = Array(19.424732212865, 3.784615384615)
    pts(15) = Array(19.406515915913, 4.061538461538)
    pts(16) = Array(19.388348368894, 4.320000000000)
    pts(17) = Array(19.360440281776, 4.689230769231)
    pts(18) = Array(19.330229916135, 5.058461538462)
    pts(19) = Array(19.297713766320, 5.427692307692)
    pts(20) = Array(19.262888051156, 5.796923076923)
    pts(21) = Array(19.225748711721, 6.166153846154)
    pts(22) = Array(19.186291408957, 6.535384615385)
    pts(23) = Array(19.144511521096, 6.904615384615)
    pts(24) = Array(19.102664866896, 7.255384615385)
    pts(25) = Array(19.056341561063, 7.624615384615)
    pts(26) = Array(19.007680361513, 7.993846153846)
    pts(27) = Array(18.956675501140, 8.363076923077)
    pts(28) = Array(18.906044547073, 8.713846153846)
    pts(29) = Array(18.850451828057, 9.083076923077)
    pts(30) = Array(18.795450655116, 9.433846153846)
    pts(31) = Array(18.735244758896, 9.803076923077)
    pts(32) = Array(18.674638214572, 10.160862462820)
    pts(33) = Array(18.674638214572, 12.000000000000)
    pts(34) = Array(23.853734603604, 12.000000000000)
    pts(35) = Array(23.853734603604, 7.451355569824)
    pts(36) = Array(23.765402380894, 7.366153846154)
    pts(37) = Array(23.672036504562, 7.273846153846)
    pts(38) = Array(23.581007679486, 7.181538461538)
    pts(39) = Array(23.492230393707, 7.089230769231)
    pts(40) = Array(23.405625223467, 6.996923076923)
    pts(41) = Array(23.321118248822, 6.904615384615)
    pts(42) = Array(23.238640540022, 6.812307692308)
    pts(43) = Array(23.142255550917, 6.701538461538)
    pts(44) = Array(23.064021363129, 6.609230769231)
    pts(45) = Array(22.972560870238, 6.498461538462)
    pts(46) = Array(22.898299846187, 6.406153846154)
    pts(47) = Array(22.811461111719, 6.295384615385)
    pts(48) = Array(22.740937481154, 6.203076923077)
    pts(49) = Array(22.658455625207, 6.092307692308)
    pts(50) = Array(22.578250161215, 5.981538461538)
    pts(51) = Array(22.500256310125, 5.870769230769)
    pts(52) = Array(22.424413415782, 5.760000000000)
    pts(53) = Array(22.350664604188, 5.649230769231)
    pts(54) = Array(22.278956478958, 5.538461538462)
    pts(55) = Array(22.209238848385, 5.427692307692)
    pts(56) = Array(22.141464480222, 5.316923076923)
    pts(57) = Array(22.075588880810, 5.206153846154)
    pts(58) = Array(22.011570095695, 5.095384615385)
    pts(59) = Array(21.949368529229, 4.984615384615)
    pts(60) = Array(21.888946781019, 4.873846153846)
    pts(61) = Array(21.830269497356, 4.763076923077)
    pts(62) = Array(21.773303235991, 4.652307692308)
    pts(63) = Array(21.708963004066, 4.523076923077)
    pts(64) = Array(21.655597596195, 4.412307692308)
    pts(65) = Array(21.603848583930, 4.301538461538)
    pts(66) = Array(21.553689277292, 4.190769230769)
    pts(67) = Array(21.497145601439, 4.061538461538)
    pts(68) = Array(21.450345552053, 3.950769230769)
    pts(69) = Array(21.397657390332, 3.821538461538)
    pts(70) = Array(21.354110468278, 3.710769230769)
    pts(71) = Array(21.305160816302, 3.581538461538)
    pts(72) = Array(21.264772249278, 3.470769230769)
    pts(73) = Array(21.219456611937, 3.341538461538)
    pts(74) = Array(21.176058442829, 3.212307692308)
    pts(75) = Array(21.134552323058, 3.083076923077)
    pts(76) = Array(21.100463415547, 2.972307692308)
    pts(77) = Array(21.062408815468, 2.843076923077)
    pts(78) = Array(21.026181793588, 2.713846153846)
    pts(79) = Array(20.991762415799, 2.584615384615)
    pts(80) = Array(20.959132010386, 2.455384615385)
    pts(81) = Array(20.928273112862, 2.326153846154)
    pts(82) = Array(20.899169415462, 2.196923076923)
    pts(83) = Array(20.871805720913, 2.067692307692)
    pts(84) = Array(20.846167900170, 1.938461538462)
    pts(85) = Array(20.822242853827, 1.809230769231)
    pts(86) = Array(20.800018476935, 1.680000000000)
    pts(87) = Array(20.779483627004, 1.550769230769)
    pts(88) = Array(20.760628094981, 1.421538461538)
    pts(89) = Array(20.743442579023, 1.292307692308)
    pts(90) = Array(20.727918660901, 1.163076923077)
    pts(91) = Array(20.714048784885, 1.033846153846)
    pts(92) = Array(20.701826239003, 0.904615384615)
    pts(93) = Array(20.691245138541, 0.775384615385)
    pts(94) = Array(20.682300411700, 0.646153846154)
    pts(95) = Array(20.674987787323, 0.516923076923)
    pts(96) = Array(20.669303784633, 0.387692307692)
    pts(97) = Array(20.665245704904, 0.258461538462)
    pts(98) = Array(20.662811625034, 0.129230769231)
    pts(99) = Array(20.662000392981, 0.000000000000)
    Profile_11 = pts
End Function

Private Function Profile_12() As Variant
    Dim pts(57) As Variant
    pts(0) = Array(41.557622748621, 0.000000000000)
    pts(1) = Array(41.555633810868, 1.689230769231)
    pts(2) = Array(41.549732042595, 3.364615384615)
    pts(3) = Array(41.543681226479, 4.472307692308)
    pts(4) = Array(41.535919857586, 5.580000000000)
    pts(5) = Array(41.526447880392, 6.687692307692)
    pts(6) = Array(41.515324397031, 7.789938124400)
    pts(7) = Array(41.515324397031, 9.000000000000)
    pts(8) = Array(43.389334334808, 9.000000000000)
    pts(9) = Array(43.389334334808, 7.789938124400)
    pts(10) = Array(43.456538553511, 7.656923076923)
    pts(11) = Array(43.518189029274, 7.532307692308)
    pts(12) = Array(43.578587820065, 7.407692307692)
    pts(13) = Array(43.637749556746, 7.283076923077)
    pts(14) = Array(43.695688350261, 7.158461538462)
    pts(15) = Array(43.758647011100, 7.020000000000)
    pts(16) = Array(43.814048154647, 6.895384615385)
    pts(17) = Array(43.874219047026, 6.756923076923)
    pts(18) = Array(43.927138818042, 6.632307692308)
    pts(19) = Array(43.984582169559, 6.493846153846)
    pts(20) = Array(44.035072781085, 6.369230769231)
    pts(21) = Array(44.089844480821, 6.230769230769)
    pts(22) = Array(44.137954392074, 6.106153846154)
    pts(23) = Array(44.190106316242, 5.967692307692)
    pts(24) = Array(44.240898978057, 5.829230769231)
    pts(25) = Array(44.290345199114, 5.690769230769)
    pts(26) = Array(44.333705789621, 5.566153846154)
    pts(27) = Array(44.380627340540, 5.427692307692)
    pts(28) = Array(44.426236900422, 5.289230769231)
    pts(29) = Array(44.470545459494, 5.150769230769)
    pts(30) = Array(44.513563577245, 5.012307692308)
    pts(31) = Array(44.555301396570, 4.873846153846)
    pts(32) = Array(44.595768657137, 4.735384615385)
    pts(33) = Array(44.634974708019, 4.596923076923)
    pts(34) = Array(44.672928519629, 4.458461538462)
    pts(35) = Array(44.709638695002, 4.320000000000)
    pts(36) = Array(44.745113480460, 4.181538461538)
    pts(37) = Array(44.779360775685, 4.043076923077)
    pts(38) = Array(44.812388143243, 3.904615384615)
    pts(39) = Array(44.844202817577, 3.766153846154)
    pts(40) = Array(44.874811713502, 3.627692307692)
    pts(41) = Array(44.904221434227, 3.489230769231)
    pts(42) = Array(44.946101259870, 3.281538461538)
    pts(43) = Array(44.987837187516, 3.060000000000)
    pts(44) = Array(45.024231782204, 2.852307692308)
    pts(45) = Array(45.060157298509, 2.630769230769)
    pts(46) = Array(45.091139582597, 2.423076923077)
    pts(47) = Array(45.121326053344, 2.201538461538)
    pts(48) = Array(45.148574296556, 1.980000000000)
    pts(49) = Array(45.172898410321, 1.758461538462)
    pts(50) = Array(45.193057709263, 1.550769230769)
    pts(51) = Array(45.211750534838, 1.329230769231)
    pts(52) = Array(45.227551572716, 1.107692307692)
    pts(53) = Array(45.240468837263, 0.886153846154)
    pts(54) = Array(45.250508858189, 0.664615384615)
    pts(55) = Array(45.257676697138, 0.443076923077)
    pts(56) = Array(45.261975960490, 0.221538461538)
    pts(57) = Array(45.263408808452, 0.000000000000)
    Profile_12 = pts
End Function

Private Function Profile_13() As Variant
    Dim pts(81) As Variant
    pts(0) = Array(46.496695763494, 0.000000000000)
    pts(1) = Array(46.498330815992, 0.251674924019)
    pts(2) = Array(46.503236801855, 0.503349848038)
    pts(3) = Array(46.511416208297, 0.755024772057)
    pts(4) = Array(46.522873187683, 1.006699696076)
    pts(5) = Array(46.537613568093, 1.258374620095)
    pts(6) = Array(46.555644868197, 1.510049544114)
    pts(7) = Array(46.576976316526, 1.761724468134)
    pts(8) = Array(46.600370157329, 2.001414871961)
    pts(9) = Array(46.628177958049, 2.253089795980)
    pts(10) = Array(46.659323380502, 2.504764719999)
    pts(11) = Array(46.693822689844, 2.756439644018)
    pts(12) = Array(46.729813877655, 2.996130047846)
    pts(13) = Array(46.770915293583, 3.247804971865)
    pts(14) = Array(46.813232273681, 3.487495375693)
    pts(15) = Array(46.861019509944, 3.739170299712)
    pts(16) = Array(46.909749229329, 3.978860703540)
    pts(17) = Array(46.964321407026, 4.230535627559)
    pts(18) = Array(47.019565975082, 4.470226031386)
    pts(19) = Array(47.081038701230, 4.721900955406)
    pts(20) = Array(47.142916519233, 4.961591359233)
    pts(21) = Array(47.208078286364, 5.201281763061)
    pts(22) = Array(47.276558534512, 5.440972166889)
    pts(23) = Array(47.348394078026, 5.680662570716)
    pts(24) = Array(47.423624122714, 5.920352974544)
    pts(25) = Array(47.502290384696, 6.160043378372)
    pts(26) = Array(47.584437220004, 6.399733782199)
    pts(27) = Array(47.670111765982, 6.639424186027)
    pts(28) = Array(47.754815698864, 6.867130069663)
    pts(29) = Array(47.847516128437, 7.106820473491)
    pts(30) = Array(47.938993417129, 7.334526357127)
    pts(31) = Array(48.033847133510, 7.562232240764)
    pts(32) = Array(48.132131453256, 7.789938124400)
    pts(33) = Array(49.249802891773, 7.789938124400)
    pts(34) = Array(49.367933994398, 7.658108402295)
    pts(35) = Array(49.492745931242, 7.514294159998)
    pts(36) = Array(49.603584292214, 7.382464437893)
    pts(37) = Array(49.720753692796, 7.238650195596)
    pts(38) = Array(49.834161024638, 7.094835953300)
    pts(39) = Array(49.943944588948, 6.951021711003)
    pts(40) = Array(50.050231881626, 6.807207468706)
    pts(41) = Array(50.153140700302, 6.663393226410)
    pts(42) = Array(50.252780107244, 6.519578984113)
    pts(43) = Array(50.357150518339, 6.363780221625)
    pts(44) = Array(50.457921706161, 6.207981459137)
    pts(45) = Array(50.555203958068, 6.052182696649)
    pts(46) = Array(50.649099507583, 5.896383934161)
    pts(47) = Array(50.739703274271, 5.740585171673)
    pts(48) = Array(50.827103515434, 5.584786409185)
    pts(49) = Array(50.911382402045, 5.428987646697)
    pts(50) = Array(50.992616529339, 5.273188884209)
    pts(51) = Array(51.070877370804, 5.117390121721)
    pts(52) = Array(51.146231682956, 4.961591359233)
    pts(53) = Array(51.224203468783, 4.793808076554)
    pts(54) = Array(51.293715915150, 4.638009314066)
    pts(55) = Array(51.365527152137, 4.470226031386)
    pts(56) = Array(51.429432505300, 4.314427268898)
    pts(57) = Array(51.495319336235, 4.146643986219)
    pts(58) = Array(51.558217146836, 3.978860703540)
    pts(59) = Array(51.618177543733, 3.811077420860)
    pts(60) = Array(51.675248622893, 3.643294138181)
    pts(61) = Array(51.729475191948, 3.475510855502)
    pts(62) = Array(51.780898970675, 3.307727572822)
    pts(63) = Array(51.829558771949, 3.139944290143)
    pts(64) = Array(51.875490665182, 2.972161007463)
    pts(65) = Array(51.921714244561, 2.792393204593)
    pts(66) = Array(51.962099159881, 2.624609921913)
    pts(67) = Array(52.002447826100, 2.444842119042)
    pts(68) = Array(52.037408207271, 2.277058836363)
    pts(69) = Array(52.072002550737, 2.097291033492)
    pts(70) = Array(52.101641498975, 1.929507750813)
    pts(71) = Array(52.130582295869, 1.749739947942)
    pts(72) = Array(52.154985180669, 1.581956665263)
    pts(73) = Array(52.178354936476, 1.402188862392)
    pts(74) = Array(52.197590679335, 1.234405579713)
    pts(75) = Array(52.215454848833, 1.054637776842)
    pts(76) = Array(52.229576926882, 0.886854494162)
    pts(77) = Array(52.241984854284, 0.707086691292)
    pts(78) = Array(52.251032059014, 0.539303408612)
    pts(79) = Array(52.258017670607, 0.359535605742)
    pts(80) = Array(52.262207013600, 0.179767802871)
    pts(81) = Array(52.263603124304, 0.000000000000)
    Profile_13 = pts
End Function

Private Function Profile_14() As Variant
    Dim pts(74) As Variant
    pts(0) = Array(52.263603124304, 0.000000000000)
    pts(1) = Array(52.262709632855, 0.143814242297)
    pts(2) = Array(52.260320445371, 0.275643964402)
    pts(3) = Array(52.256428426648, 0.407473686507)
    pts(4) = Array(52.251032059014, 0.539303408612)
    pts(5) = Array(52.244129234800, 0.671133130718)
    pts(6) = Array(52.235717252212, 0.802962852823)
    pts(7) = Array(52.225792810024, 0.934792574928)
    pts(8) = Array(52.214352001062, 1.066622297033)
    pts(9) = Array(52.201390304450, 1.198452019138)
    pts(10) = Array(52.186902576582, 1.330281741244)
    pts(11) = Array(52.170883040788, 1.462111463349)
    pts(12) = Array(52.153325275621, 1.593941185454)
    pts(13) = Array(52.134222201732, 1.725770907559)
    pts(14) = Array(52.113566067262, 1.857600629665)
    pts(15) = Array(52.091348431670, 1.989430351770)
    pts(16) = Array(52.067560147928, 2.121260073875)
    pts(17) = Array(52.042191342991, 2.253089795980)
    pts(18) = Array(52.015231396430, 2.384919518086)
    pts(19) = Array(51.986668917133, 2.516749240191)
    pts(20) = Array(51.956491717933, 2.648578962296)
    pts(21) = Array(51.924686788043, 2.780408684401)
    pts(22) = Array(51.891240263120, 2.912238406506)
    pts(23) = Array(51.856137392816, 3.044068128612)
    pts(24) = Array(51.822775199052, 3.163913330526)
    pts(25) = Array(51.784465890777, 3.295743052631)
    pts(26) = Array(51.748160500202, 3.415588254545)
    pts(27) = Array(51.706580796723, 3.547417976650)
    pts(28) = Array(51.667270776017, 3.667263178564)
    pts(29) = Array(51.622349319616, 3.799092900669)
    pts(30) = Array(51.579966207167, 3.918938102583)
    pts(31) = Array(51.531623551694, 4.050767824688)
    pts(32) = Array(51.486091182453, 4.170613026602)
    pts(33) = Array(51.434238991164, 4.302442748707)
    pts(34) = Array(51.385472708626, 4.422287950621)
    pts(35) = Array(51.335134020961, 4.542133152535)
    pts(36) = Array(51.283200513854, 4.661978354449)
    pts(37) = Array(51.224203468783, 4.793808076554)
    pts(38) = Array(51.168842647366, 4.913653278468)
    pts(39) = Array(51.111809862201, 5.033498480382)
    pts(40) = Array(51.053077510808, 5.153343682295)
    pts(41) = Array(50.992616529339, 5.273188884209)
    pts(42) = Array(50.930396302009, 5.393034086123)
    pts(43) = Array(50.866384562509, 5.512879288037)
    pts(44) = Array(50.800547286547, 5.632724489951)
    pts(45) = Array(50.732848574527, 5.752569691865)
    pts(46) = Array(50.670296922532, 5.860430373587)
    pts(47) = Array(50.598955289794, 5.980275575501)
    pts(48) = Array(50.533058402373, 6.088136257223)
    pts(49) = Array(50.457921706161, 6.207981459137)
    pts(50) = Array(50.388535099932, 6.315842140860)
    pts(51) = Array(50.317440300159, 6.423702822582)
    pts(52) = Array(50.244599178793, 6.531563504305)
    pts(53) = Array(50.169971484933, 6.639424186027)
    pts(54) = Array(50.093514686964, 6.747284867750)
    pts(55) = Array(50.015183798809, 6.855145549472)
    pts(56) = Array(49.934931188285, 6.963006231194)
    pts(57) = Array(49.852706365294, 7.070866912917)
    pts(58) = Array(49.768455747163, 7.178727594639)
    pts(59) = Array(49.682122398061, 7.286588276362)
    pts(60) = Array(49.603584292214, 7.382464437893)
    pts(61) = Array(49.513148413004, 7.490325119615)
    pts(62) = Array(49.430858477974, 7.586201281146)
    pts(63) = Array(49.336075805910, 7.694061962869)
    pts(64) = Array(49.249802891773, 7.789938124400)
    pts(65) = Array(53.505938791537, 7.789938124400)
    pts(66) = Array(53.505938791537, 4.227203193877)
    pts(67) = Array(53.492328574531, 3.691232218946)
    pts(68) = Array(53.480737083330, 3.163913330526)
    pts(69) = Array(53.470929571447, 2.636594442105)
    pts(70) = Array(53.462905702306, 2.109275553684)
    pts(71) = Array(53.456665200605, 1.581956665263)
    pts(72) = Array(53.452207852269, 1.054637776842)
    pts(73) = Array(53.449533504414, 0.527318888421)
    pts(74) = Array(53.448642065319, 0.000000000000)
    Profile_14 = pts
End Function

Private Function Profile_15() As Variant
    Dim pts(52) As Variant
    pts(0) = Array(61.484898998173, 0.000000000000)
    pts(1) = Array(61.485583781106, 3.387095230337)
    pts(2) = Array(61.487638130241, 6.774190460673)
    pts(3) = Array(64.666192770344, 6.774190460673)
    pts(4) = Array(64.666192770344, 6.097025351521)
    pts(5) = Array(64.605497147996, 6.013396762782)
    pts(6) = Array(64.538926987622, 5.919600279480)
    pts(7) = Array(64.473912162797, 5.825803796179)
    pts(8) = Array(64.410417269450, 5.732007312877)
    pts(9) = Array(64.348408761032, 5.638210829576)
    pts(10) = Array(64.287854822601, 5.544414346274)
    pts(11) = Array(64.228725255976, 5.450617862972)
    pts(12) = Array(64.170991374779, 5.356821379671)
    pts(13) = Array(64.114625908355, 5.263024896369)
    pts(14) = Array(64.059602913677, 5.169228413067)
    pts(15) = Array(64.005897694439, 5.075431929766)
    pts(16) = Array(63.947742193674, 4.971213614986)
    pts(17) = Array(63.896743033470, 4.877417131685)
    pts(18) = Array(63.841540274483, 4.773198816905)
    pts(19) = Array(63.793152312952, 4.679402333603)
    pts(20) = Array(63.740801992591, 4.575184018824)
    pts(21) = Array(63.694939283767, 4.481387535522)
    pts(22) = Array(63.645350437534, 4.377169220743)
    pts(23) = Array(63.597181316244, 4.272950905963)
    pts(24) = Array(63.550410160285, 4.168732591183)
    pts(25) = Array(63.509494247046, 4.074936107882)
    pts(26) = Array(63.465323010726, 3.970717793102)
    pts(27) = Array(63.422492405381, 3.866499478323)
    pts(28) = Array(63.380984483392, 3.762281163543)
    pts(29) = Array(63.340782155826, 3.658062848763)
    pts(30) = Array(63.301869150809, 3.553844533984)
    pts(31) = Array(63.264229974876, 3.449626219204)
    pts(32) = Array(63.227849877072, 3.345407904425)
    pts(33) = Array(63.192714815579, 3.241189589645)
    pts(34) = Array(63.158811426673, 3.136971274865)
    pts(35) = Array(63.126126995830, 3.032752960086)
    pts(36) = Array(63.091567604344, 2.918112813828)
    pts(37) = Array(63.040994790715, 2.740941678703)
    pts(38) = Array(62.993821183145, 2.563770543578)
    pts(39) = Array(62.949998814450, 2.386599408452)
    pts(40) = Array(62.907202833990, 2.199006441849)
    pts(41) = Array(62.870146719588, 2.021835306724)
    pts(42) = Array(62.834430385120, 1.834242340121)
    pts(43) = Array(62.803988604407, 1.657071204995)
    pts(44) = Array(62.775207587392, 1.469478238392)
    pts(45) = Array(62.751258725238, 1.292307103267)
    pts(46) = Array(62.729299566752, 1.104714136664)
    pts(47) = Array(62.711750186987, 0.927543001538)
    pts(48) = Array(62.696527981501, 0.739950034935)
    pts(49) = Array(62.685310751742, 0.562778899810)
    pts(50) = Array(62.676767468322, 0.375185933207)
    pts(51) = Array(62.671644897234, 0.187592966603)
    pts(52) = Array(62.669937939188, 0.000000000000)
    Profile_15 = pts
End Function

Private Function Profile_16() As Variant
    Dim pts(83) As Variant
    pts(0) = Array(62.669937939188, 0.000000000000)
    pts(1) = Array(62.671644897234, 0.187592966603)
    pts(2) = Array(62.676767468322, 0.375185933207)
    pts(3) = Array(62.685310751742, 0.562778899810)
    pts(4) = Array(62.696527981501, 0.739950034935)
    pts(5) = Array(62.711750186987, 0.927543001538)
    pts(6) = Array(62.729299566752, 1.104714136664)
    pts(7) = Array(62.751258725238, 1.292307103267)
    pts(8) = Array(62.775207587392, 1.469478238392)
    pts(9) = Array(62.803988604407, 1.657071204995)
    pts(10) = Array(62.834430385120, 1.834242340121)
    pts(11) = Array(62.870146719588, 2.021835306724)
    pts(12) = Array(62.907202833990, 2.199006441849)
    pts(13) = Array(62.949998814450, 2.386599408452)
    pts(14) = Array(62.993821183145, 2.563770543578)
    pts(15) = Array(63.040994790715, 2.740941678703)
    pts(16) = Array(63.091567604344, 2.918112813828)
    pts(17) = Array(63.126126995830, 3.032752960086)
    pts(18) = Array(63.158811426673, 3.136971274865)
    pts(19) = Array(63.192714815579, 3.241189589645)
    pts(20) = Array(63.227849877072, 3.345407904425)
    pts(21) = Array(63.264229974876, 3.449626219204)
    pts(22) = Array(63.301869150809, 3.553844533984)
    pts(23) = Array(63.340782155826, 3.658062848763)
    pts(24) = Array(63.380984483392, 3.762281163543)
    pts(25) = Array(63.422492405381, 3.866499478323)
    pts(26) = Array(63.465323010726, 3.970717793102)
    pts(27) = Array(63.509494247046, 4.074936107882)
    pts(28) = Array(63.550410160285, 4.168732591183)
    pts(29) = Array(63.597181316244, 4.272950905963)
    pts(30) = Array(63.645350437534, 4.377169220743)
    pts(31) = Array(63.694939283767, 4.481387535522)
    pts(32) = Array(63.740801992591, 4.575184018824)
    pts(33) = Array(63.793152312952, 4.679402333603)
    pts(34) = Array(63.841540274483, 4.773198816905)
    pts(35) = Array(63.896743033470, 4.877417131685)
    pts(36) = Array(63.947742193674, 4.971213614986)
    pts(37) = Array(64.005897694439, 5.075431929766)
    pts(38) = Array(64.059602913677, 5.169228413067)
    pts(39) = Array(64.114625908355, 5.263024896369)
    pts(40) = Array(64.170991374779, 5.356821379671)
    pts(41) = Array(64.228725255976, 5.450617862972)
    pts(42) = Array(64.287854822601, 5.544414346274)
    pts(43) = Array(64.348408761032, 5.638210829576)
    pts(44) = Array(64.410417269450, 5.732007312877)
    pts(45) = Array(64.473912162797, 5.825803796179)
    pts(46) = Array(64.538926987622, 5.919600279480)
    pts(47) = Array(64.605497147996, 6.013396762782)
    pts(48) = Array(64.666192770344, 6.097025351521)
    pts(49) = Array(64.666192770344, 6.774190460673)
    pts(50) = Array(67.643526146502, 6.774190460673)
    pts(51) = Array(67.643526146502, 6.097025351521)
    pts(52) = Array(67.687874462136, 5.898756616524)
    pts(53) = Array(67.732841525915, 5.690319986965)
    pts(54) = Array(67.776119974919, 5.481883357406)
    pts(55) = Array(67.817717939190, 5.273446727847)
    pts(56) = Array(67.857643183964, 5.065010098288)
    pts(57) = Array(67.895903117291, 4.856573468729)
    pts(58) = Array(67.932504797233, 4.648136839169)
    pts(59) = Array(67.967454938660, 4.439700209610)
    pts(60) = Array(68.000759919657, 4.231263580051)
    pts(61) = Array(68.032425787556, 4.022826950492)
    pts(62) = Array(68.062458264610, 3.814390320933)
    pts(63) = Array(68.090862753322, 3.605953691374)
    pts(64) = Array(68.117644341435, 3.397517061814)
    pts(65) = Array(68.142807806602, 3.189080432255)
    pts(66) = Array(68.166357620739, 2.980643802696)
    pts(67) = Array(68.189352792593, 2.761785341659)
    pts(68) = Array(68.205684683774, 2.595036038012)
    pts(69) = Array(68.221913576914, 2.417864902886)
    pts(70) = Array(68.236131997043, 2.251115599239)
    pts(71) = Array(68.250119065900, 2.073944464114)
    pts(72) = Array(68.262230796334, 1.907195160466)
    pts(73) = Array(68.273982698190, 1.730024025341)
    pts(74) = Array(68.283993492648, 1.563274721694)
    pts(75) = Array(68.293515800572, 1.386103586568)
    pts(76) = Array(68.301430398911, 1.219354282921)
    pts(77) = Array(68.308727615053, 1.042183147796)
    pts(78) = Array(68.314549754545, 0.875433844149)
    pts(79) = Array(68.319625320832, 0.698262709023)
    pts(80) = Array(68.323357745021, 0.531513405376)
    pts(81) = Array(68.326214051293, 0.354342270251)
    pts(82) = Array(68.327927692535, 0.177171135125)
    pts(83) = Array(68.328498882534, 0.000000000000)
    Profile_16 = pts
End Function

Private Function Profile_17() As Variant
    Dim pts(51) As Variant
    pts(0) = Array(68.531648415279, 0.000000000000)
    pts(1) = Array(68.533473652977, 0.222307692308)
    pts(2) = Array(68.538950843257, 0.444615384615)
    pts(3) = Array(68.548084423666, 0.666923076923)
    pts(4) = Array(68.560881810117, 0.889230769231)
    pts(5) = Array(68.577353427094, 1.111538461538)
    pts(6) = Array(68.597512750340, 1.333846153846)
    pts(7) = Array(68.621376362440, 1.556153846154)
    pts(8) = Array(68.647237733958, 1.765384615385)
    pts(9) = Array(68.678351345720, 1.987692307692)
    pts(10) = Array(68.713236750348, 2.210000000000)
    pts(11) = Array(68.751923585743, 2.432307692308)
    pts(12) = Array(68.791837025281, 2.641538461538)
    pts(13) = Array(68.838001324243, 2.863846153846)
    pts(14) = Array(68.885021064746, 3.073076923077)
    pts(15) = Array(68.938816168421, 3.295384615385)
    pts(16) = Array(68.993100613113, 3.504615384615)
    pts(17) = Array(69.054710550341, 3.726923076923)
    pts(18) = Array(69.116448491867, 3.936153846154)
    pts(19) = Array(69.181878175490, 4.145384615385)
    pts(20) = Array(69.251056127423, 4.354615384615)
    pts(21) = Array(69.324043374619, 4.563846153846)
    pts(22) = Array(69.400905749760, 4.773076923077)
    pts(23) = Array(69.481714231622, 4.982307692308)
    pts(24) = Array(69.561124003151, 5.178461538462)
    pts(25) = Array(69.649801064240, 5.387692307692)
    pts(26) = Array(69.736737778165, 5.583846153846)
    pts(27) = Array(69.833615981669, 5.793076923077)
    pts(28) = Array(69.928415051357, 5.989230769231)
    pts(29) = Array(70.027154435757, 6.185384615385)
    pts(30) = Array(70.129931882866, 6.381538461538)
    pts(31) = Array(70.236852604272, 6.577692307692)
    pts(32) = Array(70.348228931040, 6.774190460673)
    pts(33) = Array(70.348228931040, 8.500000000000)
    pts(34) = Array(71.559401477824, 8.500000000000)
    pts(35) = Array(71.559401477824, 6.197753661508)
    pts(36) = Array(71.505929797533, 5.806153846154)
    pts(37) = Array(71.457528385429, 5.426923076923)
    pts(38) = Array(71.412442944455, 5.047692307692)
    pts(39) = Array(71.369281477557, 4.655384615385)
    pts(40) = Array(71.330911774368, 4.276153846154)
    pts(41) = Array(71.294678406011, 3.883846153846)
    pts(42) = Array(71.262988893569, 3.504615384615)
    pts(43) = Array(71.233650169199, 3.112307692308)
    pts(44) = Array(71.208611678191, 2.733076923077)
    pts(45) = Array(71.186140666810, 2.340769230769)
    pts(46) = Array(71.167153440599, 1.948461538462)
    pts(47) = Array(71.151645523718, 1.556153846154)
    pts(48) = Array(71.139958376815, 1.176923076923)
    pts(49) = Array(71.131283226105, 0.784615384615)
    pts(50) = Array(71.126078949981, 0.392307692308)
    pts(51) = Array(71.124344326942, 0.000000000000)
    Profile_17 = pts
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
