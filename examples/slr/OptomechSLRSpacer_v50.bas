Attribute VB_Name = "OptomechSLRSpacer_v50"
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
    runID = "OM_" & Format$(Now, "yyyymmdd_hhnnss")
    fs.CreateFolder output
    Set log = fs.CreateTextFile(output & "\build_log.txt", False, True)
    Set buildLog = log
    log.WriteLine "Optomech workflow concept; SOLIDWORKS " & app.RevisionNumber
    names = Array("01_FrontBarrel", "02_MiddleBarrel", "03_RearBarrel", "04_G3G4_InternalCell", "05_G5G6_InternalCell", "06_G8_InternalCell", "07_FrontCupSpacer", "08_G7G8_Spacer", "09_G8G9_Spacer", "10_G9G10_CupSpacer", "11_FrontRetainer", "12_MiddleCellRetainer", "13_RearRetainer", "Lens1_Optical_Reference", "Lens2_Optical_Reference", "Lens3_RevisedBlank_Reference", "Lens4_Optical_Reference", "Lens5_Optical_Reference", "Lens6_Optical_Reference", "Lens7_Optical_Reference", "Lens8_Optical_Reference", "Lens9_RevisedBlank_Reference", "Lens10_RevisedBlank_Reference", "Lens11_Optical_Reference")
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
    a = SketchPoint(math, mapping, -387.074000000000, 0#)
    b = SketchPoint(math, mapping, -156.794000000000, 0#)
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
    If InStr(key, "02_middlebarrel") > 0 Then
        values(0) = 0.06: values(1) = 0.51: values(2) = 0.51
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "03_rearbarrel") > 0 Then
        values(0) = 0.57: values(1) = 0.36: values(2) = 0.17
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "04_g3g4_internalcell") > 0 Then
        values(0) = 0.52: values(1) = 0.54: values(2) = 0.57
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "05_g5g6_internalcell") > 0 Then
        values(0) = 0.52: values(1) = 0.54: values(2) = 0.57
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "06_g8_internalcell") > 0 Then
        values(0) = 0.52: values(1) = 0.54: values(2) = 0.57
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "07_frontcupspacer") > 0 Then
        values(0) = 0.35: values(1) = 0.65: values(2) = 0.36
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "08_g7g8_spacer") > 0 Then
        values(0) = 0.35: values(1) = 0.65: values(2) = 0.36
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "09_g8g9_spacer") > 0 Then
        values(0) = 0.35: values(1) = 0.65: values(2) = 0.36
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "10_g9g10_cupspacer") > 0 Then
        values(0) = 0.35: values(1) = 0.65: values(2) = 0.36
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "11_frontretainer") > 0 Then
        values(0) = 0.3: values(1) = 0.32: values(2) = 0.35
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "12_middlecellretainer") > 0 Then
        values(0) = 0.3: values(1) = 0.32: values(2) = 0.35
        ConceptPalette = values
        Exit Function
    End If
    If InStr(key, "13_rearretainer") > 0 Then
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
        Case 18: ProfileFor = Profile_18()
        Case 19: ProfileFor = Profile_19()
        Case 20: ProfileFor = Profile_20()
        Case 21: ProfileFor = Profile_21()
        Case 22: ProfileFor = Profile_22()
        Case 23: ProfileFor = Profile_23()
        Case Else: Err.Raise vbObjectError + 418, , "Unknown profile index."
    End Select
End Function

Private Function Profile_0() As Variant
    Dim pts(13) As Variant
    pts(0) = Array(-382.246097289273, 104.000000000000)
    pts(1) = Array(-367.352878765768, 104.000000000000)
    pts(2) = Array(-348.837733361397, 81.000000000000)
    pts(3) = Array(-310.074000000000, 81.000000000000)
    pts(4) = Array(-310.074000000000, 72.020000000000)
    pts(5) = Array(-325.451622484856, 72.020000000000)
    pts(6) = Array(-325.451622484856, 67.000000000000)
    pts(7) = Array(-329.451622484856, 67.000000000000)
    pts(8) = Array(-329.451622484856, 69.200000000000)
    pts(9) = Array(-358.837733361397, 69.200000000000)
    pts(10) = Array(-367.352878765768, 88.200000000000)
    pts(11) = Array(-374.246097289273, 88.200000000000)
    pts(12) = Array(-374.246097289273, 93.900000000000)
    pts(13) = Array(-382.246097289273, 93.900000000000)
    Profile_0 = pts
End Function

Private Function Profile_1() As Variant
    Dim pts(13) As Variant
    pts(0) = Array(-321.574000000000, 72.000000000000)
    pts(1) = Array(-310.074000000000, 72.000000000000)
    pts(2) = Array(-310.074000000000, 76.000000000000)
    pts(3) = Array(-262.074000000000, 76.000000000000)
    pts(4) = Array(-256.074000000000, 70.000000000000)
    pts(5) = Array(-242.074000000000, 70.000000000000)
    pts(6) = Array(-242.074000000000, 46.000000000000)
    pts(7) = Array(-244.414000000000, 43.600000000000)
    pts(8) = Array(-245.914000000000, 43.600000000000)
    pts(9) = Array(-250.186391992003, 51.000000000000)
    pts(10) = Array(-250.186391992003, 60.100000000000)
    pts(11) = Array(-313.074000000000, 60.100000000000)
    pts(12) = Array(-313.074000000000, 63.900000000000)
    pts(13) = Array(-321.574000000000, 63.900000000000)
    Profile_1 = pts
End Function

Private Function Profile_2() As Variant
    Dim pts(17) As Variant
    pts(0) = Array(-256.074000000000, 80.000000000000)
    pts(1) = Array(-240.074000000000, 80.000000000000)
    pts(2) = Array(-222.274000000000, 62.200000000000)
    pts(3) = Array(-196.074000000000, 62.200000000000)
    pts(4) = Array(-178.074000000000, 73.000000000000)
    pts(5) = Array(-158.300531091410, 73.000000000000)
    pts(6) = Array(-158.300531091410, 62.900000000000)
    pts(7) = Array(-166.300531091410, 62.900000000000)
    pts(8) = Array(-166.300531091410, 55.200000000000)
    pts(9) = Array(-185.948064643883, 55.200000000000)
    pts(10) = Array(-185.948064643883, 50.200000000000)
    pts(11) = Array(-236.222866362856, 50.200000000000)
    pts(12) = Array(-236.222866362856, 49.100000000000)
    pts(13) = Array(-239.681515390175, 49.100000000000)
    pts(14) = Array(-239.681515390175, 46.900000000000)
    pts(15) = Array(-242.074000000000, 46.900000000000)
    pts(16) = Array(-242.074000000000, 70.020000000000)
    pts(17) = Array(-256.074000000000, 70.020000000000)
    Profile_2 = pts
End Function

Private Function Profile_3() As Variant
    Dim pts(11) As Variant
    pts(0) = Array(-329.074000000000, 54.000000000000)
    pts(1) = Array(-314.074000000000, 54.000000000000)
    pts(2) = Array(-314.074000000000, 60.000000000000)
    pts(3) = Array(-288.074000000000, 60.000000000000)
    pts(4) = Array(-288.074000000000, 42.000000000000)
    pts(5) = Array(-292.428548958027, 42.000000000000)
    pts(6) = Array(-292.428548958027, 44.200000000000)
    pts(7) = Array(-313.405590774832, 44.200000000000)
    pts(8) = Array(-313.405590774832, 46.000000000000)
    pts(9) = Array(-317.405590774832, 46.000000000000)
    pts(10) = Array(-317.405590774832, 48.200000000000)
    pts(11) = Array(-329.074000000000, 48.200000000000)
    Profile_3 = pts
End Function

Private Function Profile_4() As Variant
    Dim pts(7) As Variant
    pts(0) = Array(-288.074000000000, 60.000000000000)
    pts(1) = Array(-250.186391992003, 60.000000000000)
    pts(2) = Array(-250.186391992003, 45.500000000000)
    pts(3) = Array(-271.322595210054, 45.500000000000)
    pts(4) = Array(-271.322595210054, 42.000000000000)
    pts(5) = Array(-275.959894706412, 42.000000000000)
    pts(6) = Array(-275.959894706412, 44.200000000000)
    pts(7) = Array(-288.074000000000, 44.200000000000)
    Profile_4 = pts
End Function

Private Function Profile_5() As Variant
    Dim pts(5) As Variant
    pts(0) = Array(-223.074000000000, 50.000000000000)
    pts(1) = Array(-189.574000000000, 50.000000000000)
    pts(2) = Array(-189.574000000000, 45.900000000000)
    pts(3) = Array(-214.306381059205, 45.900000000000)
    pts(4) = Array(-214.306381059205, 43.700000000000)
    pts(5) = Array(-223.074000000000, 43.700000000000)
    Profile_5 = pts
End Function

Private Function Profile_6() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-367.352878765768, 88.000000000000)
    pts(1) = Array(-358.837733361397, 69.000000000000)
    pts(2) = Array(-358.837733361397, 61.000000000000)
    pts(3) = Array(-367.352878765768, 80.000000000000)
    Profile_6 = pts
End Function

Private Function Profile_7() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-236.222866362856, 50.000000000000)
    pts(1) = Array(-223.074000000000, 50.000000000000)
    pts(2) = Array(-223.074000000000, 46.900000000000)
    pts(3) = Array(-236.222866362856, 46.900000000000)
    Profile_7 = pts
End Function

Private Function Profile_8() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-189.574000000000, 50.000000000000)
    pts(1) = Array(-187.274000000000, 50.000000000000)
    pts(2) = Array(-187.274000000000, 47.000000000000)
    pts(3) = Array(-189.574000000000, 47.000000000000)
    Profile_8 = pts
End Function

Private Function Profile_9() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-185.948064643883, 55.000000000000)
    pts(1) = Array(-170.895799373902, 55.000000000000)
    pts(2) = Array(-170.895799373902, 50.000000000000)
    pts(3) = Array(-185.948064643883, 47.000000000000)
    Profile_9 = pts
End Function

Private Function Profile_10() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-381.246097289273, 93.800000000000)
    pts(1) = Array(-375.246097289273, 93.800000000000)
    pts(2) = Array(-375.246097289273, 86.000000000000)
    pts(3) = Array(-381.246097289273, 86.000000000000)
    Profile_10 = pts
End Function

Private Function Profile_11() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-320.074000000000, 63.800000000000)
    pts(1) = Array(-314.074000000000, 63.800000000000)
    pts(2) = Array(-314.074000000000, 54.200000000000)
    pts(3) = Array(-320.074000000000, 54.200000000000)
    Profile_11 = pts
End Function

Private Function Profile_12() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-165.300531091410, 62.800000000000)
    pts(1) = Array(-159.300531091410, 62.800000000000)
    pts(2) = Array(-159.300531091410, 53.000000000000)
    pts(3) = Array(-165.300531091410, 53.000000000000)
    Profile_12 = pts
End Function

Private Function Profile_13() As Variant
    Dim pts(146) As Variant
    pts(0) = Array(-386.074000000000, 0.000000000000)
    pts(1) = Array(-386.073372223368, 0.660000000000)
    pts(2) = Array(-386.071889968478, 1.210000000000)
    pts(3) = Array(-386.066309657978, 2.310000000000)
    pts(4) = Array(-386.061287301784, 2.970000000000)
    pts(5) = Array(-386.056142799052, 3.520000000000)
    pts(6) = Array(-386.043237609025, 4.620000000000)
    pts(7) = Array(-386.033820005200, 5.280000000000)
    pts(8) = Array(-386.025012570267, 5.830000000000)
    pts(9) = Array(-386.004780783798, 6.930000000000)
    pts(10) = Array(-385.981059255399, 8.030000000000)
    pts(11) = Array(-385.953847269003, 9.130000000000)
    pts(12) = Array(-385.935844319906, 9.790000000000)
    pts(13) = Array(-385.919881621157, 10.340000000000)
    pts(14) = Array(-385.885336872223, 11.440000000000)
    pts(15) = Array(-385.847298773782, 12.540000000000)
    pts(16) = Array(-385.805766175921, 13.640000000000)
    pts(17) = Array(-385.779168739357, 14.300000000000)
    pts(18) = Array(-385.756042669080, 14.850000000000)
    pts(19) = Array(-385.707167406732, 15.950000000000)
    pts(20) = Array(-385.654793404360, 17.050000000000)
    pts(21) = Array(-385.598919075277, 18.150000000000)
    pts(22) = Array(-385.563713620854, 18.810000000000)
    pts(23) = Array(-385.533412411200, 19.360000000000)
    pts(24) = Array(-385.470181754612, 20.460000000000)
    pts(25) = Array(-385.430560642928, 21.120000000000)
    pts(26) = Array(-385.396578609294, 21.670000000000)
    pts(27) = Array(-385.325983159585, 22.770000000000)
    pts(28) = Array(-385.281941082830, 23.430000000000)
    pts(29) = Array(-385.244273670903, 23.980000000000)
    pts(30) = Array(-385.166303961072, 25.080000000000)
    pts(31) = Array(-385.084819079299, 26.180000000000)
    pts(32) = Array(-384.999816541087, 27.280000000000)
    pts(33) = Array(-384.947125468440, 27.940000000000)
    pts(34) = Array(-384.902247756348, 28.490000000000)
    pts(35) = Array(-384.809849563864, 29.590000000000)
    pts(36) = Array(-384.713925313741, 30.690000000000)
    pts(37) = Array(-384.614472068891, 31.790000000000)
    pts(38) = Array(-384.553104940954, 32.450000000000)
    pts(39) = Array(-384.500993868823, 33.000000000000)
    pts(40) = Array(-384.394119678250, 34.100000000000)
    pts(41) = Array(-384.283706677281, 35.200000000000)
    pts(42) = Array(-384.169751468602, 36.300000000000)
    pts(43) = Array(-384.099676627912, 36.960000000000)
    pts(44) = Array(-384.040305295526, 37.510000000000)
    pts(45) = Array(-383.918899886581, 38.610000000000)
    pts(46) = Array(-383.793941014157, 39.710000000000)
    pts(47) = Array(-383.665424811801, 40.810000000000)
    pts(48) = Array(-383.586605920193, 41.470000000000)
    pts(49) = Array(-383.519943516699, 42.020000000000)
    pts(50) = Array(-383.383943819730, 43.120000000000)
    pts(51) = Array(-383.244374067316, 44.220000000000)
    pts(52) = Array(-383.101229913635, 45.320000000000)
    pts(53) = Array(-383.013625856566, 45.980000000000)
    pts(54) = Array(-382.939637579251, 46.530000000000)
    pts(55) = Array(-382.788972508685, 47.630000000000)
    pts(56) = Array(-382.634718807745, 48.730000000000)
    pts(57) = Array(-382.476871639624, 49.830000000000)
    pts(58) = Array(-382.380436416859, 50.490000000000)
    pts(59) = Array(-382.299083378491, 51.040000000000)
    pts(60) = Array(-382.133673643208, 52.140000000000)
    pts(61) = Array(-381.964654669779, 53.240000000000)
    pts(62) = Array(-381.792021117326, 54.340000000000)
    pts(63) = Array(-381.686703719307, 55.000000000000)
    pts(64) = Array(-381.597942843455, 55.550000000000)
    pts(65) = Array(-381.417700732888, 56.650000000000)
    pts(66) = Array(-381.233826687552, 57.750000000000)
    pts(67) = Array(-381.046314847871, 58.850000000000)
    pts(68) = Array(-380.855159225352, 59.950000000000)
    pts(69) = Array(-380.660353701595, 61.050000000000)
    pts(70) = Array(-380.461892027276, 62.150000000000)
    pts(71) = Array(-380.259767821100, 63.250000000000)
    pts(72) = Array(-380.136732577709, 63.910000000000)
    pts(73) = Array(-380.033193191163, 64.460000000000)
    pts(74) = Array(-379.823356303822, 65.560000000000)
    pts(75) = Array(-379.609836252255, 66.660000000000)
    pts(76) = Array(-379.392626102634, 67.760000000000)
    pts(77) = Array(-379.171718783427, 68.860000000000)
    pts(78) = Array(-378.947107084194, 69.960000000000)
    pts(79) = Array(-378.718783654343, 71.060000000000)
    pts(80) = Array(-378.486741001869, 72.160000000000)
    pts(81) = Array(-378.274716372175, 73.150000000000)
    pts(82) = Array(-378.059667324465, 74.140000000000)
    pts(83) = Array(-377.841588098017, 75.130000000000)
    pts(84) = Array(-377.620472835984, 76.120000000000)
    pts(85) = Array(-377.396315584563, 77.110000000000)
    pts(86) = Array(-377.169110292141, 78.100000000000)
    pts(87) = Array(-376.938850808423, 79.090000000000)
    pts(88) = Array(-376.705530883547, 80.080000000000)
    pts(89) = Array(-376.469144167174, 81.070000000000)
    pts(90) = Array(-376.229684207567, 82.060000000000)
    pts(91) = Array(-375.987144450647, 83.050000000000)
    pts(92) = Array(-375.741518239026, 84.040000000000)
    pts(93) = Array(-375.492798811027, 85.030000000000)
    pts(94) = Array(-375.240979299684, 86.020000000000)
    pts(95) = Array(-374.986052731712, 87.010000000000)
    pts(96) = Array(-374.728012026471, 88.000000000000)
    pts(97) = Array(-367.352878765768, 88.000000000000)
    pts(98) = Array(-367.352878765768, 83.000000000000)
    pts(99) = Array(-367.422539885880, 81.620000000000)
    pts(100) = Array(-367.488076679807, 80.300000000000)
    pts(101) = Array(-367.552542369299, 78.980000000000)
    pts(102) = Array(-367.621171635252, 77.550000000000)
    pts(103) = Array(-367.683406262673, 76.230000000000)
    pts(104) = Array(-367.749618835485, 74.800000000000)
    pts(105) = Array(-367.809622932895, 73.480000000000)
    pts(106) = Array(-367.873419378547, 72.050000000000)
    pts(107) = Array(-367.931193458777, 70.730000000000)
    pts(108) = Array(-367.992574322430, 69.300000000000)
    pts(109) = Array(-368.048118879104, 67.980000000000)
    pts(110) = Array(-368.107084685119, 66.550000000000)
    pts(111) = Array(-368.160400192675, 65.230000000000)
    pts(112) = Array(-368.216951444643, 63.800000000000)
    pts(113) = Array(-368.268038358350, 62.480000000000)
    pts(114) = Array(-368.322175539107, 61.050000000000)
    pts(115) = Array(-368.371034295091, 59.730000000000)
    pts(116) = Array(-368.418823684662, 58.410000000000)
    pts(117) = Array(-368.465543801782, 57.090000000000)
    pts(118) = Array(-368.514950724211, 55.660000000000)
    pts(119) = Array(-368.559443482941, 54.340000000000)
    pts(120) = Array(-368.606437641924, 52.910000000000)
    pts(121) = Array(-368.648703421820, 51.590000000000)
    pts(122) = Array(-368.693285218208, 50.160000000000)
    pts(123) = Array(-368.733324379748, 48.840000000000)
    pts(124) = Array(-368.775494193731, 47.410000000000)
    pts(125) = Array(-368.815846319091, 46.000000000000)
    pts(126) = Array(-368.853065269459, 44.660000000000)
    pts(127) = Array(-368.888652199490, 43.340000000000)
    pts(128) = Array(-368.925999106681, 41.910000000000)
    pts(129) = Array(-368.959360385477, 40.590000000000)
    pts(130) = Array(-368.994296327039, 39.160000000000)
    pts(131) = Array(-369.050562934620, 36.740000000000)
    pts(132) = Array(-369.103239768663, 34.320000000000)
    pts(133) = Array(-369.152327176703, 31.900000000000)
    pts(134) = Array(-369.199808314061, 29.370000000000)
    pts(135) = Array(-369.241554697198, 26.950000000000)
    pts(136) = Array(-369.279712566645, 24.530000000000)
    pts(137) = Array(-369.314282174036, 22.110000000000)
    pts(138) = Array(-369.346586751174, 19.580000000000)
    pts(139) = Array(-369.373817415918, 17.160000000000)
    pts(140) = Array(-369.397460439079, 14.740000000000)
    pts(141) = Array(-369.417515976523, 12.320000000000)
    pts(142) = Array(-369.434647479418, 9.790000000000)
    pts(143) = Array(-369.447365364008, 7.370000000000)
    pts(144) = Array(-369.456496091849, 4.950000000000)
    pts(145) = Array(-369.462039723125, 2.530000000000)
    pts(146) = Array(-369.464000000000, 0.000000000000)
    Profile_13 = pts
End Function

Private Function Profile_14() As Variant
    Dim pts(264) As Variant
    pts(0) = Array(-368.684000000000, 0.000000000000)
    pts(1) = Array(-368.683059937316, 0.603750000000)
    pts(2) = Array(-368.680757723736, 1.121250000000)
    pts(3) = Array(-368.676325885629, 1.725000000000)
    pts(4) = Array(-368.671030569595, 2.242500000000)
    pts(5) = Array(-368.663106511310, 2.846250000000)
    pts(6) = Array(-368.654817561360, 3.363750000000)
    pts(7) = Array(-368.645146671764, 3.881250000000)
    pts(8) = Array(-368.634093635669, 4.398750000000)
    pts(9) = Array(-368.619451230171, 5.002500000000)
    pts(10) = Array(-368.605402692578, 5.520000000000)
    pts(11) = Array(-368.587264754716, 6.123750000000)
    pts(12) = Array(-368.570219303868, 6.641250000000)
    pts(13) = Array(-368.548584010331, 7.245000000000)
    pts(14) = Array(-368.528539932115, 7.762500000000)
    pts(15) = Array(-368.507111210029, 8.280000000000)
    pts(16) = Array(-368.484297384477, 8.797500000000)
    pts(17) = Array(-368.455929983043, 9.401250000000)
    pts(18) = Array(-368.430113380636, 9.918750000000)
    pts(19) = Array(-368.398241264339, 10.522500000000)
    pts(20) = Array(-368.369419281640, 11.040000000000)
    pts(21) = Array(-368.334039234467, 11.643750000000)
    pts(22) = Array(-368.302208961410, 12.161250000000)
    pts(23) = Array(-368.268989179940, 12.678750000000)
    pts(24) = Array(-368.234379174147, 13.196250000000)
    pts(25) = Array(-368.192242747513, 13.800000000000)
    pts(26) = Array(-368.154617987691, 14.317500000000)
    pts(27) = Array(-368.108962166744, 14.921250000000)
    pts(28) = Array(-368.068318837474, 15.438750000000)
    pts(29) = Array(-368.019138988398, 16.042500000000)
    pts(30) = Array(-367.975472960921, 16.560000000000)
    pts(31) = Array(-367.930410362207, 17.077500000000)
    pts(32) = Array(-367.883950214415, 17.595000000000)
    pts(33) = Array(-367.827979013887, 18.198750000000)
    pts(34) = Array(-367.778487339355, 18.716250000000)
    pts(35) = Array(-367.718976453115, 19.320000000000)
    pts(36) = Array(-367.666448194413, 19.837500000000)
    pts(37) = Array(-367.603391534319, 20.441250000000)
    pts(38) = Array(-367.547821311405, 20.958750000000)
    pts(39) = Array(-367.490845200492, 21.476250000000)
    pts(40) = Array(-367.432461954056, 21.993750000000)
    pts(41) = Array(-367.362567994317, 22.597500000000)
    pts(42) = Array(-367.301131520254, 23.115000000000)
    pts(43) = Array(-367.227671810733, 23.718750000000)
    pts(44) = Array(-367.163175769290, 24.236250000000)
    pts(45) = Array(-367.086142718899, 24.840000000000)
    pts(46) = Array(-367.018580435699, 25.357500000000)
    pts(47) = Array(-366.949600619433, 25.875000000000)
    pts(48) = Array(-366.879201742851, 26.392500000000)
    pts(49) = Array(-366.795274101283, 26.996250000000)
    pts(50) = Array(-366.721795193980, 27.513750000000)
    pts(51) = Array(-366.634269756441, 28.117500000000)
    pts(52) = Array(-366.557703144979, 28.635000000000)
    pts(53) = Array(-366.466570755796, 29.238750000000)
    pts(54) = Array(-366.386908417143, 29.756250000000)
    pts(55) = Array(-366.305814477554, 30.273750000000)
    pts(56) = Array(-366.223287117537, 30.791250000000)
    pts(57) = Array(-366.125191036864, 31.395000000000)
    pts(58) = Array(-366.039551522085, 31.912500000000)
    pts(59) = Array(-365.952472591048, 32.430000000000)
    pts(60) = Array(-365.863952275612, 32.947500000000)
    pts(61) = Array(-365.758854148446, 33.551250000000)
    pts(62) = Array(-365.667203908593, 34.068750000000)
    pts(63) = Array(-365.574105811578, 34.586250000000)
    pts(64) = Array(-365.479557736613, 35.103750000000)
    pts(65) = Array(-365.367416156971, 35.707500000000)
    pts(66) = Array(-365.269719009685, 36.225000000000)
    pts(67) = Array(-365.170564921495, 36.742500000000)
    pts(68) = Array(-365.069951614593, 37.260000000000)
    pts(69) = Array(-364.950722040707, 37.863750000000)
    pts(70) = Array(-364.846939095863, 38.381250000000)
    pts(71) = Array(-364.741689464718, 38.898750000000)
    pts(72) = Array(-364.634970707681, 39.416250000000)
    pts(73) = Array(-364.508605369563, 40.020000000000)
    pts(74) = Array(-364.398694948276, 40.537500000000)
    pts(75) = Array(-364.287307413159, 41.055000000000)
    pts(76) = Array(-364.174440157673, 41.572500000000)
    pts(77) = Array(-364.040887956575, 42.176250000000)
    pts(78) = Array(-363.924805503148, 42.693750000000)
    pts(79) = Array(-363.807234804002, 43.211250000000)
    pts(80) = Array(-363.688173080039, 43.728750000000)
    pts(81) = Array(-363.547379479303, 44.332500000000)
    pts(82) = Array(-363.425077465669, 44.850000000000)
    pts(83) = Array(-363.301275345962, 45.367500000000)
    pts(84) = Array(-363.175970162444, 45.885000000000)
    pts(85) = Array(-363.049158913545, 46.402500000000)
    pts(86) = Array(-362.920838553476, 46.920000000000)
    pts(87) = Array(-362.791005991835, 47.437500000000)
    pts(88) = Array(-362.648164819298, 48.000000000000)
    pts(89) = Array(-362.504499414454, 48.558750000000)
    pts(90) = Array(-362.369857312484, 49.076250000000)
    pts(91) = Array(-362.233689643067, 49.593750000000)
    pts(92) = Array(-362.095993078749, 50.111250000000)
    pts(93) = Array(-361.933410228967, 50.715000000000)
    pts(94) = Array(-361.792389418455, 51.232500000000)
    pts(95) = Array(-361.649828862503, 51.750000000000)
    pts(96) = Array(-361.505725033944, 52.267500000000)
    pts(97) = Array(-361.335648601604, 52.871250000000)
    pts(98) = Array(-361.188188680780, 53.388750000000)
    pts(99) = Array(-361.039173991472, 53.906250000000)
    pts(100) = Array(-360.888600798510, 54.423750000000)
    pts(101) = Array(-360.736465315498, 54.941250000000)
    pts(102) = Array(-360.582763704285, 55.458750000000)
    pts(103) = Array(-360.427492074430, 55.976250000000)
    pts(104) = Array(-360.270646482659, 56.493750000000)
    pts(105) = Array(-360.085665310984, 57.097500000000)
    pts(106) = Array(-359.925395685596, 57.615000000000)
    pts(107) = Array(-359.763539256312, 58.132500000000)
    pts(108) = Array(-359.600091853610, 58.650000000000)
    pts(109) = Array(-359.435049252143, 59.167500000000)
    pts(110) = Array(-359.268407170134, 59.685000000000)
    pts(111) = Array(-359.100161268745, 60.202500000000)
    pts(112) = Array(-358.930307151447, 60.720000000000)
    pts(113) = Array(-358.758840363367, 61.237500000000)
    pts(114) = Array(-358.585756390633, 61.755000000000)
    pts(115) = Array(-358.411050659698, 62.272500000000)
    pts(116) = Array(-358.234718536658, 62.790000000000)
    pts(117) = Array(-358.056755326548, 63.307500000000)
    pts(118) = Array(-357.877156272637, 63.825000000000)
    pts(119) = Array(-357.695916555698, 64.342500000000)
    pts(120) = Array(-357.513031293271, 64.860000000000)
    pts(121) = Array(-357.328495538907, 65.377500000000)
    pts(122) = Array(-357.142304281403, 65.895000000000)
    pts(123) = Array(-356.954452444017, 66.412500000000)
    pts(124) = Array(-356.764934883672, 66.930000000000)
    pts(125) = Array(-356.573746390140, 67.447500000000)
    pts(126) = Array(-356.380881685217, 67.965000000000)
    pts(127) = Array(-356.186335421871, 68.482500000000)
    pts(128) = Array(-355.990102183388, 69.000000000000)
    pts(129) = Array(-329.451622484856, 69.000000000000)
    pts(130) = Array(-329.451622484856, 48.000000000000)
    pts(131) = Array(-329.836695392831, 47.696250000000)
    pts(132) = Array(-330.159761646846, 47.437500000000)
    pts(133) = Array(-330.478387125590, 47.178750000000)
    pts(134) = Array(-330.792680974215, 46.920000000000)
    pts(135) = Array(-331.102747377351, 46.661250000000)
    pts(136) = Array(-331.408685869336, 46.402500000000)
    pts(137) = Array(-331.710591619830, 46.143750000000)
    pts(138) = Array(-331.909666841544, 45.971250000000)
    pts(139) = Array(-332.107016308976, 45.798750000000)
    pts(140) = Array(-332.399859848151, 45.540000000000)
    pts(141) = Array(-332.688959589046, 45.281250000000)
    pts(142) = Array(-332.974394585183, 45.022500000000)
    pts(143) = Array(-333.162686211304, 44.850000000000)
    pts(144) = Array(-333.349404586306, 44.677500000000)
    pts(145) = Array(-333.626578455399, 44.418750000000)
    pts(146) = Array(-333.809455112753, 44.246250000000)
    pts(147) = Array(-333.990830243630, 44.073750000000)
    pts(148) = Array(-334.260119726455, 43.815000000000)
    pts(149) = Array(-334.437824549443, 43.642500000000)
    pts(150) = Array(-334.614093414082, 43.470000000000)
    pts(151) = Array(-334.875842901838, 43.211250000000)
    pts(152) = Array(-335.219978057401, 42.866250000000)
    pts(153) = Array(-335.474503333449, 42.607500000000)
    pts(154) = Array(-335.809203592838, 42.262500000000)
    pts(155) = Array(-336.056795032335, 42.003750000000)
    pts(156) = Array(-336.382432052460, 41.658750000000)
    pts(157) = Array(-336.703045823127, 41.313750000000)
    pts(158) = Array(-337.018743432464, 40.968750000000)
    pts(159) = Array(-337.252352229014, 40.710000000000)
    pts(160) = Array(-337.559690563792, 40.365000000000)
    pts(161) = Array(-337.862384324236, 40.020000000000)
    pts(162) = Array(-338.160524684113, 39.675000000000)
    pts(163) = Array(-338.381194544963, 39.416250000000)
    pts(164) = Array(-338.671574856822, 39.071250000000)
    pts(165) = Array(-338.957634069359, 38.726250000000)
    pts(166) = Array(-339.239450517920, 38.381250000000)
    pts(167) = Array(-339.517099535111, 38.036250000000)
    pts(168) = Array(-339.790653605321, 37.691250000000)
    pts(169) = Array(-340.060182509025, 37.346250000000)
    pts(170) = Array(-340.325753457685, 37.001250000000)
    pts(171) = Array(-340.587431219977, 36.656250000000)
    pts(172) = Array(-340.845278240031, 36.311250000000)
    pts(173) = Array(-341.099354748280, 35.966250000000)
    pts(174) = Array(-341.349718865483, 35.621250000000)
    pts(175) = Array(-341.596426700428, 35.276250000000)
    pts(176) = Array(-341.839532441769, 34.931250000000)
    pts(177) = Array(-342.079088444435, 34.586250000000)
    pts(178) = Array(-342.315145310990, 34.241250000000)
    pts(179) = Array(-342.547751968291, 33.896250000000)
    pts(180) = Array(-342.776955739795, 33.551250000000)
    pts(181) = Array(-343.002802413785, 33.206250000000)
    pts(182) = Array(-343.114480744210, 33.033750000000)
    pts(183) = Array(-343.280457217842, 32.775000000000)
    pts(184) = Array(-343.498910316214, 32.430000000000)
    pts(185) = Array(-343.714145281869, 32.085000000000)
    pts(186) = Array(-343.926202056312, 31.740000000000)
    pts(187) = Array(-344.031050764890, 31.567500000000)
    pts(188) = Array(-344.186862548471, 31.308750000000)
    pts(189) = Array(-344.391908116907, 30.963750000000)
    pts(190) = Array(-344.593897030222, 30.618750000000)
    pts(191) = Array(-344.792864307638, 30.273750000000)
    pts(192) = Array(-345.037375890251, 29.842500000000)
    pts(193) = Array(-345.229667175826, 29.497500000000)
    pts(194) = Array(-345.465935918356, 29.066250000000)
    pts(195) = Array(-345.651712064470, 28.721250000000)
    pts(196) = Array(-345.879932455172, 28.290000000000)
    pts(197) = Array(-346.059343992408, 27.945000000000)
    pts(198) = Array(-346.279698126542, 27.513750000000)
    pts(199) = Array(-346.452886057565, 27.168750000000)
    pts(200) = Array(-346.665544577217, 26.737500000000)
    pts(201) = Array(-346.832641089176, 26.392500000000)
    pts(202) = Array(-347.037764027621, 25.961250000000)
    pts(203) = Array(-347.198893132899, 25.616250000000)
    pts(204) = Array(-347.396630674856, 25.185000000000)
    pts(205) = Array(-347.551908788627, 24.840000000000)
    pts(206) = Array(-347.742401956233, 24.408750000000)
    pts(207) = Array(-347.928928238960, 23.977500000000)
    pts(208) = Array(-348.111527127377, 23.546250000000)
    pts(209) = Array(-348.254804196678, 23.201250000000)
    pts(210) = Array(-348.430429052486, 22.770000000000)
    pts(211) = Array(-348.568174131061, 22.425000000000)
    pts(212) = Array(-348.736941181065, 21.993750000000)
    pts(213) = Array(-348.869244421235, 21.648750000000)
    pts(214) = Array(-349.031262932618, 21.217500000000)
    pts(215) = Array(-349.189577045843, 20.786250000000)
    pts(216) = Array(-349.344217389001, 20.355000000000)
    pts(217) = Array(-349.465304469003, 20.010000000000)
    pts(218) = Array(-349.613405923142, 19.578750000000)
    pts(219) = Array(-349.757913958057, 19.147500000000)
    pts(220) = Array(-349.898855433951, 18.716250000000)
    pts(221) = Array(-350.009058154076, 18.371250000000)
    pts(222) = Array(-350.143644591399, 17.940000000000)
    pts(223) = Array(-350.274734867983, 17.508750000000)
    pts(224) = Array(-350.402352486503, 17.077500000000)
    pts(225) = Array(-350.501961534347, 16.732500000000)
    pts(226) = Array(-350.623384977430, 16.301250000000)
    pts(227) = Array(-350.741397253928, 15.870000000000)
    pts(228) = Array(-350.856018848589, 15.438750000000)
    pts(229) = Array(-350.967269500875, 15.007500000000)
    pts(230) = Array(-351.075168223453, 14.576250000000)
    pts(231) = Array(-351.179733319800, 14.145000000000)
    pts(232) = Array(-351.280982400953, 13.713750000000)
    pts(233) = Array(-351.359605537277, 13.368750000000)
    pts(234) = Array(-351.454928021164, 12.937500000000)
    pts(235) = Array(-351.546980249615, 12.506250000000)
    pts(236) = Array(-351.635777341177, 12.075000000000)
    pts(237) = Array(-351.721333791592, 11.643750000000)
    pts(238) = Array(-351.803663486276, 11.212500000000)
    pts(239) = Array(-351.882779712155, 10.781250000000)
    pts(240) = Array(-351.958695168896, 10.350000000000)
    pts(241) = Array(-352.031421979540, 9.918750000000)
    pts(242) = Array(-352.100971700575, 9.487500000000)
    pts(243) = Array(-352.167355331458, 9.056250000000)
    pts(244) = Array(-352.230583323607, 8.625000000000)
    pts(245) = Array(-352.290665588886, 8.193750000000)
    pts(246) = Array(-352.347611507594, 7.762500000000)
    pts(247) = Array(-352.401429935972, 7.331250000000)
    pts(248) = Array(-352.452129213253, 6.900000000000)
    pts(249) = Array(-352.499717168250, 6.468750000000)
    pts(250) = Array(-352.544201125518, 6.037500000000)
    pts(251) = Array(-352.585587911079, 5.606250000000)
    pts(252) = Array(-352.623883857735, 5.175000000000)
    pts(253) = Array(-352.659094809972, 4.743750000000)
    pts(254) = Array(-352.691226128466, 4.312500000000)
    pts(255) = Array(-352.720282694204, 3.881250000000)
    pts(256) = Array(-352.746268912217, 3.450000000000)
    pts(257) = Array(-352.769188714940, 3.018750000000)
    pts(258) = Array(-352.789045565201, 2.587500000000)
    pts(259) = Array(-352.805842458851, 2.156250000000)
    pts(260) = Array(-352.819581927028, 1.725000000000)
    pts(261) = Array(-352.830266038067, 1.293750000000)
    pts(262) = Array(-352.837896399062, 0.862500000000)
    pts(263) = Array(-352.842474157076, 0.431250000000)
    pts(264) = Array(-352.844000000000, 0.000000000000)
    Profile_14 = pts
End Function

Private Function Profile_15() As Variant
    Dim pts(259) As Variant
    pts(0) = Array(-339.574000000000, 0.000000000000)
    pts(1) = Array(-339.573267488304, 0.420000000000)
    pts(2) = Array(-339.571473563289, 0.780000000000)
    pts(3) = Array(-339.568020182387, 1.200000000000)
    pts(4) = Array(-339.563893935069, 1.560000000000)
    pts(5) = Array(-339.557719250121, 1.980000000000)
    pts(6) = Array(-339.551260160869, 2.340000000000)
    pts(7) = Array(-339.543724118709, 2.700000000000)
    pts(8) = Array(-339.535110921366, 3.060000000000)
    pts(9) = Array(-339.523700476538, 3.480000000000)
    pts(10) = Array(-339.512752610945, 3.840000000000)
    pts(11) = Array(-339.500726758322, 4.200000000000)
    pts(12) = Array(-339.487622595442, 4.560000000000)
    pts(13) = Array(-339.470971069362, 4.980000000000)
    pts(14) = Array(-339.455529319887, 5.340000000000)
    pts(15) = Array(-339.439008044199, 5.700000000000)
    pts(16) = Array(-339.421406797366, 6.060000000000)
    pts(17) = Array(-339.399506411482, 6.480000000000)
    pts(18) = Array(-339.379563560856, 6.840000000000)
    pts(19) = Array(-339.358539136171, 7.200000000000)
    pts(20) = Array(-339.336432569747, 7.560000000000)
    pts(21) = Array(-339.309273070114, 7.980000000000)
    pts(22) = Array(-339.284819774422, 8.340000000000)
    pts(23) = Array(-339.259282343352, 8.700000000000)
    pts(24) = Array(-339.232660085138, 9.060000000000)
    pts(25) = Array(-339.200228718312, 9.480000000000)
    pts(26) = Array(-339.171253485004, 9.840000000000)
    pts(27) = Array(-339.141191035480, 10.200000000000)
    pts(28) = Array(-339.110040552251, 10.560000000000)
    pts(29) = Array(-339.072322035162, 10.980000000000)
    pts(30) = Array(-339.038811195867, 11.340000000000)
    pts(31) = Array(-339.004209532686, 11.700000000000)
    pts(32) = Array(-338.968516100456, 12.060000000000)
    pts(33) = Array(-338.925492584095, 12.480000000000)
    pts(34) = Array(-338.887430262088, 12.840000000000)
    pts(35) = Array(-338.848272973068, 13.200000000000)
    pts(36) = Array(-338.808019641930, 13.560000000000)
    pts(37) = Array(-338.759670668644, 13.980000000000)
    pts(38) = Array(-338.717038740782, 14.340000000000)
    pts(39) = Array(-338.673307157345, 14.700000000000)
    pts(40) = Array(-338.628474710681, 15.060000000000)
    pts(41) = Array(-338.574777164830, 15.480000000000)
    pts(42) = Array(-338.527555217750, 15.840000000000)
    pts(43) = Array(-338.479228369699, 16.200000000000)
    pts(44) = Array(-338.429795277528, 16.560000000000)
    pts(45) = Array(-338.370723329237, 16.980000000000)
    pts(46) = Array(-338.318888609551, 17.340000000000)
    pts(47) = Array(-338.265943173790, 17.700000000000)
    pts(48) = Array(-338.211885539988, 18.060000000000)
    pts(49) = Array(-338.147410581653, 18.480000000000)
    pts(50) = Array(-338.090937939862, 18.840000000000)
    pts(51) = Array(-338.033348182754, 19.200000000000)
    pts(52) = Array(-337.974639685841, 19.560000000000)
    pts(53) = Array(-337.904730261023, 19.980000000000)
    pts(54) = Array(-337.843592088818, 20.340000000000)
    pts(55) = Array(-337.781329801843, 20.700000000000)
    pts(56) = Array(-337.717941628975, 21.060000000000)
    pts(57) = Array(-337.642563353264, 21.480000000000)
    pts(58) = Array(-337.576729513841, 21.840000000000)
    pts(59) = Array(-337.509763942170, 22.200000000000)
    pts(60) = Array(-337.441664715951, 22.560000000000)
    pts(61) = Array(-337.372429875874, 22.920000000000)
    pts(62) = Array(-337.302057425326, 23.280000000000)
    pts(63) = Array(-337.230545330086, 23.640000000000)
    pts(64) = Array(-337.157891518019, 24.000000000000)
    pts(65) = Array(-337.071682916473, 24.420000000000)
    pts(66) = Array(-336.996548093438, 24.780000000000)
    pts(67) = Array(-336.920264738055, 25.140000000000)
    pts(68) = Array(-336.842830616736, 25.500000000000)
    pts(69) = Array(-336.764243456269, 25.860000000000)
    pts(70) = Array(-336.684500943468, 26.220000000000)
    pts(71) = Array(-336.603600724810, 26.580000000000)
    pts(72) = Array(-336.521540406067, 26.940000000000)
    pts(73) = Array(-336.424333878663, 27.360000000000)
    pts(74) = Array(-336.339751599921, 27.720000000000)
    pts(75) = Array(-336.254001366494, 28.080000000000)
    pts(76) = Array(-336.167080610394, 28.440000000000)
    pts(77) = Array(-336.078986720894, 28.800000000000)
    pts(78) = Array(-335.989717044100, 29.160000000000)
    pts(79) = Array(-335.899268882525, 29.520000000000)
    pts(80) = Array(-335.807639494650, 29.880000000000)
    pts(81) = Array(-335.699241884022, 30.300000000000)
    pts(82) = Array(-335.605043554986, 30.660000000000)
    pts(83) = Array(-335.509655022524, 31.020000000000)
    pts(84) = Array(-335.413073356709, 31.380000000000)
    pts(85) = Array(-335.315295581161, 31.740000000000)
    pts(86) = Array(-335.216318672548, 32.100000000000)
    pts(87) = Array(-335.116139560071, 32.460000000000)
    pts(88) = Array(-335.014755124947, 32.820000000000)
    pts(89) = Array(-334.912162199876, 33.180000000000)
    pts(90) = Array(-334.808357568496, 33.540000000000)
    pts(91) = Array(-334.703337964833, 33.900000000000)
    pts(92) = Array(-334.597100072727, 34.260000000000)
    pts(93) = Array(-334.489640525263, 34.620000000000)
    pts(94) = Array(-334.380955904171, 34.980000000000)
    pts(95) = Array(-334.271042739230, 35.340000000000)
    pts(96) = Array(-334.159897507650, 35.700000000000)
    pts(97) = Array(-334.047516633443, 36.060000000000)
    pts(98) = Array(-333.933896486779, 36.420000000000)
    pts(99) = Array(-333.819033383334, 36.780000000000)
    pts(100) = Array(-333.702923583620, 37.140000000000)
    pts(101) = Array(-333.585563292297, 37.500000000000)
    pts(102) = Array(-333.466948657480, 37.860000000000)
    pts(103) = Array(-333.347075770024, 38.220000000000)
    pts(104) = Array(-333.225940662794, 38.580000000000)
    pts(105) = Array(-333.103539309922, 38.940000000000)
    pts(106) = Array(-332.979867626050, 39.300000000000)
    pts(107) = Array(-332.854921465550, 39.660000000000)
    pts(108) = Array(-332.728696621734, 40.020000000000)
    pts(109) = Array(-332.601188826043, 40.380000000000)
    pts(110) = Array(-332.472393747220, 40.740000000000)
    pts(111) = Array(-332.342306990465, 41.100000000000)
    pts(112) = Array(-332.232911446276, 41.400000000000)
    pts(113) = Array(-332.100444984690, 41.760000000000)
    pts(114) = Array(-331.966674040539, 42.120000000000)
    pts(115) = Array(-331.831593967280, 42.480000000000)
    pts(116) = Array(-331.695200050476, 42.840000000000)
    pts(117) = Array(-331.557487506839, 43.200000000000)
    pts(118) = Array(-331.418451483253, 43.560000000000)
    pts(119) = Array(-331.278087055767, 43.920000000000)
    pts(120) = Array(-331.160098344511, 44.220000000000)
    pts(121) = Array(-331.017285482108, 44.580000000000)
    pts(122) = Array(-330.873129870926, 44.940000000000)
    pts(123) = Array(-330.727626305822, 45.300000000000)
    pts(124) = Array(-330.605339840112, 45.600000000000)
    pts(125) = Array(-330.457351254264, 45.960000000000)
    pts(126) = Array(-330.307999550125, 46.320000000000)
    pts(127) = Array(-330.157279224991, 46.680000000000)
    pts(128) = Array(-330.022152125175, 47.000000000000)
    pts(129) = Array(-328.874000000000, 48.000000000000)
    pts(130) = Array(-317.405590774832, 48.000000000000)
    pts(131) = Array(-317.405590774832, 42.000000000000)
    pts(132) = Array(-317.694702312449, 41.700000000000)
    pts(133) = Array(-317.979855937378, 41.400000000000)
    pts(134) = Array(-318.261128361506, 41.100000000000)
    pts(135) = Array(-318.538593322466, 40.800000000000)
    pts(136) = Array(-318.812321740697, 40.500000000000)
    pts(137) = Array(-319.082381865899, 40.200000000000)
    pts(138) = Array(-319.348839413756, 39.900000000000)
    pts(139) = Array(-319.611757693692, 39.600000000000)
    pts(140) = Array(-319.871197728386, 39.300000000000)
    pts(141) = Array(-320.127218365676, 39.000000000000)
    pts(142) = Array(-320.379876383439, 38.700000000000)
    pts(143) = Array(-320.629226587987, 38.400000000000)
    pts(144) = Array(-320.875321906445, 38.100000000000)
    pts(145) = Array(-321.118213473576, 37.800000000000)
    pts(146) = Array(-321.357950713444, 37.500000000000)
    pts(147) = Array(-321.641538867244, 37.140000000000)
    pts(148) = Array(-321.874502615156, 36.840000000000)
    pts(149) = Array(-322.150093949019, 36.480000000000)
    pts(150) = Array(-322.376499417129, 36.180000000000)
    pts(151) = Array(-322.644343156248, 35.820000000000)
    pts(152) = Array(-322.864390847838, 35.520000000000)
    pts(153) = Array(-323.124719382475, 35.160000000000)
    pts(154) = Array(-323.338596412040, 34.860000000000)
    pts(155) = Array(-323.591626787628, 34.500000000000)
    pts(156) = Array(-323.799508049742, 34.200000000000)
    pts(157) = Array(-324.045443288431, 33.840000000000)
    pts(158) = Array(-324.247492496034, 33.540000000000)
    pts(159) = Array(-324.486522772395, 33.180000000000)
    pts(160) = Array(-324.682893381877, 32.880000000000)
    pts(161) = Array(-324.915197071934, 32.520000000000)
    pts(162) = Array(-325.143902979132, 32.160000000000)
    pts(163) = Array(-325.369059274041, 31.800000000000)
    pts(164) = Array(-325.554011666341, 31.500000000000)
    pts(165) = Array(-325.772780237291, 31.140000000000)
    pts(166) = Array(-325.988127294306, 30.780000000000)
    pts(167) = Array(-326.200095283729, 30.420000000000)
    pts(168) = Array(-326.374183731734, 30.120000000000)
    pts(169) = Array(-326.580062566685, 29.760000000000)
    pts(170) = Array(-326.782675448591, 29.400000000000)
    pts(171) = Array(-326.982059934129, 29.040000000000)
    pts(172) = Array(-327.178252421460, 28.680000000000)
    pts(173) = Array(-327.371288193860, 28.320000000000)
    pts(174) = Array(-327.561201461154, 27.960000000000)
    pts(175) = Array(-327.748025399069, 27.600000000000)
    pts(176) = Array(-327.931792186634, 27.240000000000)
    pts(177) = Array(-328.112533041749, 26.880000000000)
    pts(178) = Array(-328.290278255018, 26.520000000000)
    pts(179) = Array(-328.465057221957, 26.160000000000)
    pts(180) = Array(-328.636898473664, 25.800000000000)
    pts(181) = Array(-328.805829706037, 25.440000000000)
    pts(182) = Array(-328.971877807624, 25.080000000000)
    pts(183) = Array(-329.135068886180, 24.720000000000)
    pts(184) = Array(-329.295428294000, 24.360000000000)
    pts(185) = Array(-329.452980652096, 24.000000000000)
    pts(186) = Array(-329.607749873274, 23.640000000000)
    pts(187) = Array(-329.784827338443, 23.220000000000)
    pts(188) = Array(-329.933645223743, 22.860000000000)
    pts(189) = Array(-330.079751268301, 22.500000000000)
    pts(190) = Array(-330.223166660477, 22.140000000000)
    pts(191) = Array(-330.387111360820, 21.720000000000)
    pts(192) = Array(-330.524766838759, 21.360000000000)
    pts(193) = Array(-330.682045404699, 20.940000000000)
    pts(194) = Array(-330.814031751623, 20.580000000000)
    pts(195) = Array(-330.964746828310, 20.160000000000)
    pts(196) = Array(-331.070259754263, 19.860000000000)
    pts(197) = Array(-331.174000000000, 19.560000000000)
    pts(198) = Array(-331.275977218641, 19.260000000000)
    pts(199) = Array(-331.376200831060, 18.960000000000)
    pts(200) = Array(-331.474680030842, 18.660000000000)
    pts(201) = Array(-331.571423789061, 18.360000000000)
    pts(202) = Array(-331.666440858900, 18.060000000000)
    pts(203) = Array(-331.778194123238, 17.700000000000)
    pts(204) = Array(-331.869442244388, 17.400000000000)
    pts(205) = Array(-331.958990274055, 17.100000000000)
    pts(206) = Array(-332.046846097201, 16.800000000000)
    pts(207) = Array(-332.133017402469, 16.500000000000)
    pts(208) = Array(-332.217511685902, 16.200000000000)
    pts(209) = Array(-332.300336254529, 15.900000000000)
    pts(210) = Array(-332.381498229842, 15.600000000000)
    pts(211) = Array(-332.476707734946, 15.240000000000)
    pts(212) = Array(-332.554236179567, 14.940000000000)
    pts(213) = Array(-332.630123611123, 14.640000000000)
    pts(214) = Array(-332.704376407121, 14.340000000000)
    pts(215) = Array(-332.791330772756, 13.980000000000)
    pts(216) = Array(-332.862008989069, 13.680000000000)
    pts(217) = Array(-332.931071853647, 13.380000000000)
    pts(218) = Array(-332.998525051163, 13.080000000000)
    pts(219) = Array(-333.077351897920, 12.720000000000)
    pts(220) = Array(-333.141283061829, 12.420000000000)
    pts(221) = Array(-333.203621698180, 12.120000000000)
    pts(222) = Array(-333.264372842923, 11.820000000000)
    pts(223) = Array(-333.335185598669, 11.460000000000)
    pts(224) = Array(-333.392461258919, 11.160000000000)
    pts(225) = Array(-333.448164553747, 10.860000000000)
    pts(226) = Array(-333.502299908300, 10.560000000000)
    pts(227) = Array(-333.565198675666, 10.200000000000)
    pts(228) = Array(-333.615899448580, 9.900000000000)
    pts(229) = Array(-333.665045506440, 9.600000000000)
    pts(230) = Array(-333.712640696025, 9.300000000000)
    pts(231) = Array(-333.767712996426, 8.940000000000)
    pts(232) = Array(-333.811909160901, 8.640000000000)
    pts(233) = Array(-333.854565877112, 8.340000000000)
    pts(234) = Array(-333.895686440990, 8.040000000000)
    pts(235) = Array(-333.943007848676, 7.680000000000)
    pts(236) = Array(-333.980759852167, 7.380000000000)
    pts(237) = Array(-334.016985398960, 7.080000000000)
    pts(238) = Array(-334.051687257227, 6.780000000000)
    pts(239) = Array(-334.091321934534, 6.420000000000)
    pts(240) = Array(-334.122680818831, 6.120000000000)
    pts(241) = Array(-334.158311022948, 5.760000000000)
    pts(242) = Array(-334.186338449314, 5.460000000000)
    pts(243) = Array(-334.217977235660, 5.100000000000)
    pts(244) = Array(-334.242683451985, 4.800000000000)
    pts(245) = Array(-334.265883076605, 4.500000000000)
    pts(246) = Array(-334.287577852385, 4.200000000000)
    pts(247) = Array(-334.311627466116, 3.840000000000)
    pts(248) = Array(-334.330017141146, 3.540000000000)
    pts(249) = Array(-334.346906771427, 3.240000000000)
    pts(250) = Array(-334.362297618708, 2.940000000000)
    pts(251) = Array(-334.378789851098, 2.580000000000)
    pts(252) = Array(-334.390887259921, 2.280000000000)
    pts(253) = Array(-334.401489164696, 1.980000000000)
    pts(254) = Array(-334.410596354414, 1.680000000000)
    pts(255) = Array(-334.419552904008, 1.320000000000)
    pts(256) = Array(-334.425373947863, 1.020000000000)
    pts(257) = Array(-334.430388550468, 0.660000000000)
    pts(258) = Array(-334.432925541747, 0.360000000000)
    pts(259) = Array(-334.434000000000, 0.000000000000)
    Profile_15 = pts
End Function

Private Function Profile_16() As Variant
    Dim pts(82) As Variant
    pts(0) = Array(-319.784000000000, 0.000000000000)
    pts(1) = Array(-319.783236042930, 0.605000000000)
    pts(2) = Array(-319.781215648854, 1.155000000000)
    pts(3) = Array(-319.777534699594, 1.760000000000)
    pts(4) = Array(-319.772862401233, 2.310000000000)
    pts(5) = Array(-319.766264203060, 2.915000000000)
    pts(6) = Array(-319.758939674527, 3.465000000000)
    pts(7) = Array(-319.750352011014, 4.015000000000)
    pts(8) = Array(-319.740501076653, 4.565000000000)
    pts(9) = Array(-319.728205784390, 5.170000000000)
    pts(10) = Array(-319.715701449973, 5.720000000000)
    pts(11) = Array(-319.700486962653, 6.325000000000)
    pts(12) = Array(-319.685328355241, 6.875000000000)
    pts(13) = Array(-319.667193610483, 7.480000000000)
    pts(14) = Array(-319.649379671328, 8.030000000000)
    pts(15) = Array(-319.630300862366, 8.580000000000)
    pts(16) = Array(-319.609956881171, 9.130000000000)
    pts(17) = Array(-319.586116842314, 9.735000000000)
    pts(18) = Array(-319.563114925308, 10.285000000000)
    pts(19) = Array(-319.536350297055, 10.890000000000)
    pts(20) = Array(-319.510688834498, 11.440000000000)
    pts(21) = Array(-319.480997743003, 12.045000000000)
    pts(22) = Array(-319.452674937495, 12.595000000000)
    pts(23) = Array(-319.423084139485, 13.145000000000)
    pts(24) = Array(-319.392224878327, 13.695000000000)
    pts(25) = Array(-319.356814029418, 14.300000000000)
    pts(26) = Array(-319.323289372802, 14.850000000000)
    pts(27) = Array(-319.284945319160, 15.455000000000)
    pts(28) = Array(-319.248752911714, 16.005000000000)
    pts(29) = Array(-319.207472957746, 16.610000000000)
    pts(30) = Array(-319.168610253549, 17.160000000000)
    pts(31) = Array(-319.128475029142, 17.710000000000)
    pts(32) = Array(-319.087066643073, 18.260000000000)
    pts(33) = Array(-319.031331272854, 18.975000000000)
    pts(34) = Array(-318.977971198266, 19.635000000000)
    pts(35) = Array(-318.918091607193, 20.350000000000)
    pts(36) = Array(-318.860903377928, 21.010000000000)
    pts(37) = Array(-318.796873564341, 21.725000000000)
    pts(38) = Array(-318.735851442864, 22.385000000000)
    pts(39) = Array(-318.672986898611, 23.045000000000)
    pts(40) = Array(-318.608278473974, 23.705000000000)
    pts(41) = Array(-318.536095169738, 24.420000000000)
    pts(42) = Array(-318.467540451952, 25.080000000000)
    pts(43) = Array(-318.391186634459, 25.795000000000)
    pts(44) = Array(-318.318778700747, 26.455000000000)
    pts(45) = Array(-318.238246650265, 27.170000000000)
    pts(46) = Array(-318.161978170691, 27.830000000000)
    pts(47) = Array(-318.083853905000, 28.490000000000)
    pts(48) = Array(-318.003872023988, 29.150000000000)
    pts(49) = Array(-317.915126542943, 29.865000000000)
    pts(50) = Array(-317.831268553673, 30.525000000000)
    pts(51) = Array(-317.738319375988, 31.240000000000)
    pts(52) = Array(-317.650576720369, 31.900000000000)
    pts(53) = Array(-317.553414346028, 32.615000000000)
    pts(54) = Array(-317.461778042529, 33.275000000000)
    pts(55) = Array(-317.368269472275, 33.935000000000)
    pts(56) = Array(-317.272886419601, 34.595000000000)
    pts(57) = Array(-317.167436851232, 35.310000000000)
    pts(58) = Array(-317.068141300313, 35.970000000000)
    pts(59) = Array(-316.966964131837, 36.630000000000)
    pts(60) = Array(-316.863902932295, 37.290000000000)
    pts(61) = Array(-316.750124365257, 38.005000000000)
    pts(62) = Array(-316.643130133880, 38.665000000000)
    pts(63) = Array(-316.534244118803, 39.325000000000)
    pts(64) = Array(-316.420923929690, 40.000000000000)
    pts(65) = Array(-316.420923929690, 44.000000000000)
    pts(66) = Array(-292.463122195330, 44.000000000000)
    pts(67) = Array(-292.415998690253, 41.250000000000)
    pts(68) = Array(-292.371916219474, 38.500000000000)
    pts(69) = Array(-292.330874621323, 35.750000000000)
    pts(70) = Array(-292.292873745293, 33.000000000000)
    pts(71) = Array(-292.257913452035, 30.250000000000)
    pts(72) = Array(-292.225993613360, 27.500000000000)
    pts(73) = Array(-292.197114112231, 24.750000000000)
    pts(74) = Array(-292.171274842765, 22.000000000000)
    pts(75) = Array(-292.148475710229, 19.250000000000)
    pts(76) = Array(-292.128716631042, 16.500000000000)
    pts(77) = Array(-292.111997532767, 13.750000000000)
    pts(78) = Array(-292.098318354114, 11.000000000000)
    pts(79) = Array(-292.087679044940, 8.250000000000)
    pts(80) = Array(-292.080079566243, 5.500000000000)
    pts(81) = Array(-292.075519890168, 2.750000000000)
    pts(82) = Array(-292.074000000000, 0.000000000000)
    Profile_16 = pts
End Function

Private Function Profile_17() As Variant
    Dim pts(193) As Variant
    pts(0) = Array(-291.664000000000, 0.000000000000)
    pts(1) = Array(-291.662328873121, 0.715000000000)
    pts(2) = Array(-291.657315382932, 1.430000000000)
    pts(3) = Array(-291.648959200748, 2.145000000000)
    pts(4) = Array(-291.638278445210, 2.805000000000)
    pts(5) = Array(-291.623492279384, 3.520000000000)
    pts(6) = Array(-291.605361202974, 4.235000000000)
    pts(7) = Array(-291.583884025885, 4.950000000000)
    pts(8) = Array(-291.561087811379, 5.610000000000)
    pts(9) = Array(-291.533171670016, 6.325000000000)
    pts(10) = Array(-291.501904683963, 7.040000000000)
    pts(11) = Array(-291.467284795692, 7.755000000000)
    pts(12) = Array(-291.432350058855, 8.415000000000)
    pts(13) = Array(-291.391275679905, 9.130000000000)
    pts(14) = Array(-291.350378511765, 9.790000000000)
    pts(15) = Array(-291.302839617735, 10.505000000000)
    pts(16) = Array(-291.255970026414, 11.165000000000)
    pts(17) = Array(-291.201954989272, 11.880000000000)
    pts(18) = Array(-291.144567225835, 12.595000000000)
    pts(19) = Array(-291.083802925473, 13.310000000000)
    pts(20) = Array(-291.024712384839, 13.970000000000)
    pts(21) = Array(-290.957443185651, 14.685000000000)
    pts(22) = Array(-290.892340656179, 15.345000000000)
    pts(23) = Array(-290.818549958575, 16.060000000000)
    pts(24) = Array(-290.747419347365, 16.720000000000)
    pts(25) = Array(-290.667088877457, 17.435000000000)
    pts(26) = Array(-290.589912535775, 18.095000000000)
    pts(27) = Array(-290.503022322427, 18.810000000000)
    pts(28) = Array(-290.419781023472, 19.470000000000)
    pts(29) = Array(-290.326309373056, 20.185000000000)
    pts(30) = Array(-290.236982287569, 20.845000000000)
    pts(31) = Array(-290.136905756350, 21.560000000000)
    pts(32) = Array(-290.041470426001, 22.220000000000)
    pts(33) = Array(-289.934763790069, 22.935000000000)
    pts(34) = Array(-289.833196098544, 23.595000000000)
    pts(35) = Array(-289.719832321199, 24.310000000000)
    pts(36) = Array(-289.612106462883, 24.970000000000)
    pts(37) = Array(-289.492056659383, 25.685000000000)
    pts(38) = Array(-289.378145105556, 26.345000000000)
    pts(39) = Array(-289.251378505097, 27.060000000000)
    pts(40) = Array(-289.131251967534, 27.720000000000)
    pts(41) = Array(-288.997735872363, 28.435000000000)
    pts(42) = Array(-288.871363264214, 29.095000000000)
    pts(43) = Array(-288.731063005732, 29.810000000000)
    pts(44) = Array(-288.598411399568, 30.470000000000)
    pts(45) = Array(-288.451290291279, 31.185000000000)
    pts(46) = Array(-288.312324874152, 31.845000000000)
    pts(47) = Array(-288.158344161313, 32.560000000000)
    pts(48) = Array(-288.013028186699, 33.220000000000)
    pts(49) = Array(-287.852146992492, 33.935000000000)
    pts(50) = Array(-287.700441728933, 34.595000000000)
    pts(51) = Array(-287.545655252522, 35.255000000000)
    pts(52) = Array(-287.387777940150, 35.915000000000)
    pts(53) = Array(-287.213244824993, 36.630000000000)
    pts(54) = Array(-287.048896338879, 37.290000000000)
    pts(55) = Array(-286.881425861201, 37.950000000000)
    pts(56) = Array(-286.710822794240, 38.610000000000)
    pts(57) = Array(-286.522455180368, 39.325000000000)
    pts(58) = Array(-286.345290761025, 39.985000000000)
    pts(59) = Array(-286.164959500876, 40.645000000000)
    pts(60) = Array(-285.981449767096, 41.305000000000)
    pts(61) = Array(-285.779046932149, 42.020000000000)
    pts(62) = Array(-285.588876858558, 42.680000000000)
    pts(63) = Array(-285.395490765421, 43.340000000000)
    pts(64) = Array(-285.198875916367, 44.000000000000)
    pts(65) = Array(-275.959894706412, 44.000000000000)
    pts(66) = Array(-276.111022556013, 43.725000000000)
    pts(67) = Array(-276.290752740972, 43.395000000000)
    pts(68) = Array(-276.468723357117, 43.065000000000)
    pts(69) = Array(-276.644946085729, 42.735000000000)
    pts(70) = Array(-276.790471445457, 42.460000000000)
    pts(71) = Array(-276.963519337709, 42.130000000000)
    pts(72) = Array(-277.134851280237, 41.800000000000)
    pts(73) = Array(-277.304478150609, 41.470000000000)
    pts(74) = Array(-277.444539100437, 41.195000000000)
    pts(75) = Array(-277.611067600416, 40.865000000000)
    pts(76) = Array(-277.775920783484, 40.535000000000)
    pts(77) = Array(-277.939108788691, 40.205000000000)
    pts(78) = Array(-278.100641572251, 39.875000000000)
    pts(79) = Array(-278.260528911375, 39.545000000000)
    pts(80) = Array(-278.418780407991, 39.215000000000)
    pts(81) = Array(-278.575405492361, 38.885000000000)
    pts(82) = Array(-278.730413426592, 38.555000000000)
    pts(83) = Array(-278.883813308042, 38.225000000000)
    pts(84) = Array(-279.035614072640, 37.895000000000)
    pts(85) = Array(-279.185824498098, 37.565000000000)
    pts(86) = Array(-279.334453207048, 37.235000000000)
    pts(87) = Array(-279.481508670079, 36.905000000000)
    pts(88) = Array(-279.626999208691, 36.575000000000)
    pts(89) = Array(-279.770932998179, 36.245000000000)
    pts(90) = Array(-279.913318070419, 35.915000000000)
    pts(91) = Array(-280.054162316593, 35.585000000000)
    pts(92) = Array(-280.193473489830, 35.255000000000)
    pts(93) = Array(-280.331259207782, 34.925000000000)
    pts(94) = Array(-280.467526955123, 34.595000000000)
    pts(95) = Array(-280.602284085988, 34.265000000000)
    pts(96) = Array(-280.735537826342, 33.935000000000)
    pts(97) = Array(-280.867295276286, 33.605000000000)
    pts(98) = Array(-280.997563412303, 33.275000000000)
    pts(99) = Array(-281.126349089443, 32.945000000000)
    pts(100) = Array(-281.253659043453, 32.615000000000)
    pts(101) = Array(-281.379499892848, 32.285000000000)
    pts(102) = Array(-281.503878140927, 31.955000000000)
    pts(103) = Array(-281.626800177741, 31.625000000000)
    pts(104) = Array(-281.748272282005, 31.295000000000)
    pts(105) = Array(-281.888165407758, 30.910000000000)
    pts(106) = Array(-282.006517011907, 30.580000000000)
    pts(107) = Array(-282.123437852486, 30.250000000000)
    pts(108) = Array(-282.238933763839, 29.920000000000)
    pts(109) = Array(-282.353010478104, 29.590000000000)
    pts(110) = Array(-282.465673626840, 29.260000000000)
    pts(111) = Array(-282.576928742618, 28.930000000000)
    pts(112) = Array(-282.686781260566, 28.600000000000)
    pts(113) = Array(-282.813176921476, 28.215000000000)
    pts(114) = Array(-282.920008667920, 27.885000000000)
    pts(115) = Array(-283.025454401834, 27.555000000000)
    pts(116) = Array(-283.129519168147, 27.225000000000)
    pts(117) = Array(-283.249189277958, 26.840000000000)
    pts(118) = Array(-283.350278827499, 26.510000000000)
    pts(119) = Array(-283.450002797235, 26.180000000000)
    pts(120) = Array(-283.548365865461, 25.850000000000)
    pts(121) = Array(-283.661408873702, 25.465000000000)
    pts(122) = Array(-283.756838968962, 25.135000000000)
    pts(123) = Array(-283.850922423794, 24.805000000000)
    pts(124) = Array(-283.943663571053, 24.475000000000)
    pts(125) = Array(-284.050170717935, 24.090000000000)
    pts(126) = Array(-284.140018016582, 23.760000000000)
    pts(127) = Array(-284.228536205714, 23.430000000000)
    pts(128) = Array(-284.315729291826, 23.100000000000)
    pts(129) = Array(-284.415785015896, 22.715000000000)
    pts(130) = Array(-284.500120430598, 22.385000000000)
    pts(131) = Array(-284.583142934631, 22.055000000000)
    pts(132) = Array(-284.664856225292, 21.725000000000)
    pts(133) = Array(-284.758538530276, 21.340000000000)
    pts(134) = Array(-284.837427534629, 21.010000000000)
    pts(135) = Array(-284.927824449983, 20.625000000000)
    pts(136) = Array(-285.003905471718, 20.295000000000)
    pts(137) = Array(-285.091035608185, 19.910000000000)
    pts(138) = Array(-285.164324252021, 19.580000000000)
    pts(139) = Array(-285.236329179867, 19.250000000000)
    pts(140) = Array(-285.307053504463, 18.920000000000)
    pts(141) = Array(-285.387950738489, 18.535000000000)
    pts(142) = Array(-285.455910791685, 18.205000000000)
    pts(143) = Array(-285.533591100024, 17.820000000000)
    pts(144) = Array(-285.598800537941, 17.490000000000)
    pts(145) = Array(-285.673279479061, 17.105000000000)
    pts(146) = Array(-285.735751320662, 16.775000000000)
    pts(147) = Array(-285.796963874245, 16.445000000000)
    pts(148) = Array(-285.856919719720, 16.115000000000)
    pts(149) = Array(-285.925283225496, 15.730000000000)
    pts(150) = Array(-285.982524769803, 15.400000000000)
    pts(151) = Array(-286.047728227182, 15.015000000000)
    pts(152) = Array(-286.102266702245, 14.685000000000)
    pts(153) = Array(-286.164322870572, 14.300000000000)
    pts(154) = Array(-286.216168920337, 13.970000000000)
    pts(155) = Array(-286.275089879361, 13.585000000000)
    pts(156) = Array(-286.324253570557, 13.255000000000)
    pts(157) = Array(-286.380050732581, 12.870000000000)
    pts(158) = Array(-286.426541564806, 12.540000000000)
    pts(159) = Array(-286.479225686084, 12.155000000000)
    pts(160) = Array(-286.523052601256, 11.825000000000)
    pts(161) = Array(-286.572633792644, 11.440000000000)
    pts(162) = Array(-286.613805183820, 11.110000000000)
    pts(163) = Array(-286.660292920721, 10.725000000000)
    pts(164) = Array(-286.698816640324, 10.395000000000)
    pts(165) = Array(-286.742219771941, 10.010000000000)
    pts(166) = Array(-286.778103139401, 9.680000000000)
    pts(167) = Array(-286.818429897321, 9.295000000000)
    pts(168) = Array(-286.851679706157, 8.965000000000)
    pts(169) = Array(-286.888937712292, 8.580000000000)
    pts(170) = Array(-286.919560236643, 8.250000000000)
    pts(171) = Array(-286.953756510535, 7.865000000000)
    pts(172) = Array(-286.981757511162, 7.535000000000)
    pts(173) = Array(-287.012898476681, 7.150000000000)
    pts(174) = Array(-287.038283206447, 6.820000000000)
    pts(175) = Array(-287.066374697890, 6.435000000000)
    pts(176) = Array(-287.089147906734, 6.105000000000)
    pts(177) = Array(-287.114195174336, 5.720000000000)
    pts(178) = Array(-287.134361113755, 5.390000000000)
    pts(179) = Array(-287.156368828636, 5.005000000000)
    pts(180) = Array(-287.173931255678, 4.675000000000)
    pts(181) = Array(-287.192903514217, 4.290000000000)
    pts(182) = Array(-287.207865695000, 3.960000000000)
    pts(183) = Array(-287.223806022672, 3.575000000000)
    pts(184) = Array(-287.236170735433, 3.245000000000)
    pts(185) = Array(-287.249082090089, 2.860000000000)
    pts(186) = Array(-287.258851627771, 2.530000000000)
    pts(187) = Array(-287.268736402395, 2.145000000000)
    pts(188) = Array(-287.275912574775, 1.815000000000)
    pts(189) = Array(-287.282772599703, 1.430000000000)
    pts(190) = Array(-287.287356735075, 1.100000000000)
    pts(191) = Array(-287.291193279689, 0.715000000000)
    pts(192) = Array(-287.293186226092, 0.385000000000)
    pts(193) = Array(-287.294000000000, 0.000000000000)
    Profile_17 = pts
End Function

Private Function Profile_18() As Variant
    Dim pts(118) As Variant
    pts(0) = Array(-278.804000000000, 0.000000000000)
    pts(1) = Array(-278.803354149130, 0.396375000000)
    pts(2) = Array(-278.801772459047, 0.736125000000)
    pts(3) = Array(-278.795761849137, 1.415625000000)
    pts(4) = Array(-278.785954426807, 2.095125000000)
    pts(5) = Array(-278.772349273392, 2.774625000000)
    pts(6) = Array(-278.762658628378, 3.171000000000)
    pts(7) = Array(-278.753323235894, 3.510750000000)
    pts(8) = Array(-278.731801634151, 4.190250000000)
    pts(9) = Array(-278.706477223474, 4.869750000000)
    pts(10) = Array(-278.677347626112, 5.549250000000)
    pts(11) = Array(-278.658597030970, 5.945625000000)
    pts(12) = Array(-278.641493331135, 6.285375000000)
    pts(13) = Array(-278.604427050522, 6.964875000000)
    pts(14) = Array(-278.563545981454, 7.644375000000)
    pts(15) = Array(-278.518846268421, 8.323875000000)
    pts(16) = Array(-278.470323689827, 9.003375000000)
    pts(17) = Array(-278.417973655982, 9.682875000000)
    pts(18) = Array(-278.361791206920, 10.362375000000)
    pts(19) = Array(-278.301771010030, 11.041875000000)
    pts(20) = Array(-278.264984681063, 11.438250000000)
    pts(21) = Array(-278.232411709099, 11.778000000000)
    pts(22) = Array(-278.164377433436, 12.457500000000)
    pts(23) = Array(-278.092486593574, 13.137000000000)
    pts(24) = Array(-278.016732306731, 13.816500000000)
    pts(25) = Array(-277.937107302504, 14.496000000000)
    pts(26) = Array(-277.853603919330, 15.175500000000)
    pts(27) = Array(-277.766214100726, 15.855000000000)
    pts(28) = Array(-277.674929391327, 16.534500000000)
    pts(29) = Array(-277.579740932708, 17.214000000000)
    pts(30) = Array(-277.480639458981, 17.893500000000)
    pts(31) = Array(-277.377615292171, 18.573000000000)
    pts(32) = Array(-277.270658337357, 19.252500000000)
    pts(33) = Array(-277.159758077580, 19.932000000000)
    pts(34) = Array(-277.044903568500, 20.611500000000)
    pts(35) = Array(-276.926083432813, 21.291000000000)
    pts(36) = Array(-276.803285854405, 21.970500000000)
    pts(37) = Array(-276.687216870805, 22.593375000000)
    pts(38) = Array(-276.567785655802, 23.216250000000)
    pts(39) = Array(-276.444982109404, 23.839125000000)
    pts(40) = Array(-276.318795804203, 24.462000000000)
    pts(41) = Array(-276.189215980724, 25.084875000000)
    pts(42) = Array(-276.056231542599, 25.707750000000)
    pts(43) = Array(-275.919831051549, 26.330625000000)
    pts(44) = Array(-275.792856351325, 26.896875000000)
    pts(45) = Array(-275.649901281779, 27.519750000000)
    pts(46) = Array(-275.503494890919, 28.142625000000)
    pts(47) = Array(-275.353624352102, 28.765500000000)
    pts(48) = Array(-275.214361231954, 29.331750000000)
    pts(49) = Array(-275.057840360192, 29.954625000000)
    pts(50) = Array(-274.897815903387, 30.577500000000)
    pts(51) = Array(-274.734273554539, 31.200375000000)
    pts(52) = Array(-274.582533611947, 31.766625000000)
    pts(53) = Array(-274.474572494617, 32.163000000000)
    pts(54) = Array(-274.365171308173, 32.559375000000)
    pts(55) = Array(-274.254326015866, 32.955750000000)
    pts(56) = Array(-274.142032512261, 33.352125000000)
    pts(57) = Array(-274.028286622401, 33.748500000000)
    pts(58) = Array(-273.913084100955, 34.144875000000)
    pts(59) = Array(-273.796420631342, 34.541250000000)
    pts(60) = Array(-273.678291824844, 34.937625000000)
    pts(61) = Array(-273.558693219687, 35.334000000000)
    pts(62) = Array(-273.437620280110, 35.730375000000)
    pts(63) = Array(-273.315068395409, 36.126750000000)
    pts(64) = Array(-273.191032878957, 36.523125000000)
    pts(65) = Array(-273.065508967207, 36.919500000000)
    pts(66) = Array(-272.938491818664, 37.315875000000)
    pts(67) = Array(-272.809976512841, 37.712250000000)
    pts(68) = Array(-272.698624337212, 38.052000000000)
    pts(69) = Array(-272.567313411042, 38.448375000000)
    pts(70) = Array(-272.434489824131, 38.844750000000)
    pts(71) = Array(-272.300148341045, 39.241125000000)
    pts(72) = Array(-272.164283641915, 39.637500000000)
    pts(73) = Array(-272.026890321238, 40.033875000000)
    pts(74) = Array(-271.887962886652, 40.430250000000)
    pts(75) = Array(-271.747495757679, 40.826625000000)
    pts(76) = Array(-271.625865594955, 41.166375000000)
    pts(77) = Array(-271.482523922295, 41.562750000000)
    pts(78) = Array(-271.337626113914, 41.959125000000)
    pts(79) = Array(-271.191166237161, 42.355500000000)
    pts(80) = Array(-271.064381354833, 42.695250000000)
    pts(81) = Array(-270.915004430384, 43.091625000000)
    pts(82) = Array(-270.785710115413, 43.431375000000)
    pts(83) = Array(-270.633394734720, 43.827750000000)
    pts(84) = Array(-270.488901008622, 44.200000000000)
    pts(85) = Array(-270.488901008622, 45.300000000000)
    pts(86) = Array(-250.546258963164, 45.300000000000)
    pts(87) = Array(-250.476557635419, 43.884375000000)
    pts(88) = Array(-250.409074159256, 42.468750000000)
    pts(89) = Array(-250.343808038991, 41.053125000000)
    pts(90) = Array(-250.280758795369, 39.637500000000)
    pts(91) = Array(-250.219925965542, 38.221875000000)
    pts(92) = Array(-250.161309103057, 36.806250000000)
    pts(93) = Array(-250.104907777837, 35.390625000000)
    pts(94) = Array(-250.050721576166, 33.975000000000)
    pts(95) = Array(-249.998750100672, 32.559375000000)
    pts(96) = Array(-249.948992970316, 31.143750000000)
    pts(97) = Array(-249.901449820373, 29.728125000000)
    pts(98) = Array(-249.856120302424, 28.312500000000)
    pts(99) = Array(-249.813004084340, 26.896875000000)
    pts(100) = Array(-249.772100850271, 25.481250000000)
    pts(101) = Array(-249.733410300632, 24.065625000000)
    pts(102) = Array(-249.696932152095, 22.650000000000)
    pts(103) = Array(-249.662666137578, 21.234375000000)
    pts(104) = Array(-249.630612006235, 19.818750000000)
    pts(105) = Array(-249.600769523446, 18.403125000000)
    pts(106) = Array(-249.573138470808, 16.987500000000)
    pts(107) = Array(-249.547718646130, 15.571875000000)
    pts(108) = Array(-249.524509863422, 14.156250000000)
    pts(109) = Array(-249.503511952890, 12.740625000000)
    pts(110) = Array(-249.484724760930, 11.325000000000)
    pts(111) = Array(-249.468148150122, 9.909375000000)
    pts(112) = Array(-249.453781999225, 8.493750000000)
    pts(113) = Array(-249.441626203171, 7.078125000000)
    pts(114) = Array(-249.431680673065, 5.662500000000)
    pts(115) = Array(-249.423945336178, 4.246875000000)
    pts(116) = Array(-249.418420135946, 2.831250000000)
    pts(117) = Array(-249.415105031966, 1.415625000000)
    pts(118) = Array(-249.414000000000, 0.000000000000)
    Profile_18 = pts
End Function

Private Function Profile_19() As Variant
    Dim pts(152) As Variant
    pts(0) = Array(-239.014000000000, 0.000000000000)
    pts(1) = Array(-239.014651749170, 1.528125000000)
    pts(2) = Array(-239.016606998103, 3.056250000000)
    pts(3) = Array(-239.019865751068, 4.584375000000)
    pts(4) = Array(-239.024428015177, 6.112500000000)
    pts(5) = Array(-239.030293800389, 7.640625000000)
    pts(6) = Array(-239.037463119510, 9.168750000000)
    pts(7) = Array(-239.045935988189, 10.696875000000)
    pts(8) = Array(-239.055712424922, 12.225000000000)
    pts(9) = Array(-239.066792451052, 13.753125000000)
    pts(10) = Array(-239.079176090768, 15.281250000000)
    pts(11) = Array(-239.092863371104, 16.809375000000)
    pts(12) = Array(-239.107854321943, 18.337500000000)
    pts(13) = Array(-239.124148976014, 19.865625000000)
    pts(14) = Array(-239.141747368894, 21.393750000000)
    pts(15) = Array(-239.160649539008, 22.921875000000)
    pts(16) = Array(-239.180855527630, 24.450000000000)
    pts(17) = Array(-239.202365378881, 25.978125000000)
    pts(18) = Array(-239.225179139733, 27.506250000000)
    pts(19) = Array(-239.249296860008, 29.034375000000)
    pts(20) = Array(-239.274718592376, 30.562500000000)
    pts(21) = Array(-239.301444392362, 32.090625000000)
    pts(22) = Array(-239.329474318338, 33.618750000000)
    pts(23) = Array(-239.358808431532, 35.146875000000)
    pts(24) = Array(-239.389446796022, 36.675000000000)
    pts(25) = Array(-239.421389478742, 38.203125000000)
    pts(26) = Array(-239.454636549479, 39.731250000000)
    pts(27) = Array(-239.489188080874, 41.259375000000)
    pts(28) = Array(-239.525044148427, 42.787500000000)
    pts(29) = Array(-239.562204830492, 44.315625000000)
    pts(30) = Array(-239.600670208281, 45.843750000000)
    pts(31) = Array(-239.640440365865, 47.371875000000)
    pts(32) = Array(-239.681515390175, 48.900000000000)
    pts(33) = Array(-236.283268581790, 48.900000000000)
    pts(34) = Array(-236.283268581790, 47.000000000000)
    pts(35) = Array(-236.102287855250, 46.699500000000)
    pts(36) = Array(-235.956197607463, 46.455000000000)
    pts(37) = Array(-235.667124605863, 45.966000000000)
    pts(38) = Array(-235.524130402685, 45.721500000000)
    pts(39) = Array(-235.346822363120, 45.415875000000)
    pts(40) = Array(-235.206116986170, 45.171375000000)
    pts(41) = Array(-235.031652686408, 44.865750000000)
    pts(42) = Array(-234.893208680137, 44.621250000000)
    pts(43) = Array(-234.721554348599, 44.315625000000)
    pts(44) = Array(-234.585345125669, 44.071125000000)
    pts(45) = Array(-234.416468054632, 43.765500000000)
    pts(46) = Array(-234.282467862105, 43.521000000000)
    pts(47) = Array(-234.116336366567, 43.215375000000)
    pts(48) = Array(-233.984520253862, 42.970875000000)
    pts(49) = Array(-233.821103632526, 42.665250000000)
    pts(50) = Array(-233.691447421027, 42.420750000000)
    pts(51) = Array(-233.530715919299, 42.115125000000)
    pts(52) = Array(-233.403196173505, 41.870625000000)
    pts(53) = Array(-233.245120948359, 41.565000000000)
    pts(54) = Array(-233.119714948504, 41.320500000000)
    pts(55) = Array(-232.964268035116, 41.014875000000)
    pts(56) = Array(-232.840953751145, 40.770375000000)
    pts(57) = Array(-232.688108031188, 40.464750000000)
    pts(58) = Array(-232.566864098026, 40.220250000000)
    pts(59) = Array(-232.416593269519, 39.914625000000)
    pts(60) = Array(-232.297398963551, 39.670125000000)
    pts(61) = Array(-232.149677512187, 39.364500000000)
    pts(62) = Array(-232.032512728880, 39.120000000000)
    pts(63) = Array(-231.887315900731, 38.814375000000)
    pts(64) = Array(-231.772161133330, 38.569875000000)
    pts(65) = Array(-231.629464908863, 38.264250000000)
    pts(66) = Array(-231.376082297429, 37.714125000000)
    pts(67) = Array(-231.127127071490, 37.164000000000)
    pts(68) = Array(-230.882559439405, 36.613875000000)
    pts(69) = Array(-230.748570141937, 36.308250000000)
    pts(70) = Array(-230.615916626071, 36.002625000000)
    pts(71) = Array(-230.380486211376, 35.452500000000)
    pts(72) = Array(-230.251539343594, 35.146875000000)
    pts(73) = Array(-230.123904610008, 34.841250000000)
    pts(74) = Array(-229.897449634110, 34.291125000000)
    pts(75) = Array(-229.773457358831, 33.985500000000)
    pts(76) = Array(-229.650754995648, 33.679875000000)
    pts(77) = Array(-229.433123431845, 33.129750000000)
    pts(78) = Array(-229.314003200832, 32.824125000000)
    pts(79) = Array(-229.196151996046, 32.518500000000)
    pts(80) = Array(-228.987200959947, 31.968375000000)
    pts(81) = Array(-228.872875186986, 31.662750000000)
    pts(82) = Array(-228.759798809618, 31.357125000000)
    pts(83) = Array(-228.559394004867, 30.807000000000)
    pts(84) = Array(-228.449789768362, 30.501375000000)
    pts(85) = Array(-228.341416478301, 30.195750000000)
    pts(86) = Array(-228.149431691933, 29.645625000000)
    pts(87) = Array(-228.044480464004, 29.340000000000)
    pts(88) = Array(-227.940742847637, 29.034375000000)
    pts(89) = Array(-227.797540717890, 28.606500000000)
    pts(90) = Array(-227.656696889561, 28.178625000000)
    pts(91) = Array(-227.518199838099, 27.750750000000)
    pts(92) = Array(-227.362776666278, 27.261750000000)
    pts(93) = Array(-227.229270841292, 26.833875000000)
    pts(94) = Array(-227.098077275483, 26.406000000000)
    pts(95) = Array(-226.969185509925, 25.978125000000)
    pts(96) = Array(-226.824686074351, 25.489125000000)
    pts(97) = Array(-226.700692651649, 25.061250000000)
    pts(98) = Array(-226.578969731541, 24.633375000000)
    pts(99) = Array(-226.459507840842, 24.205500000000)
    pts(100) = Array(-226.325736722640, 23.716500000000)
    pts(101) = Array(-226.211089046489, 23.288625000000)
    pts(102) = Array(-226.098674087531, 22.860750000000)
    pts(103) = Array(-225.988483288630, 22.432875000000)
    pts(104) = Array(-225.865263712367, 21.943875000000)
    pts(105) = Array(-225.759811135747, 21.516000000000)
    pts(106) = Array(-225.656557187726, 21.088125000000)
    pts(107) = Array(-225.555494166413, 20.660250000000)
    pts(108) = Array(-225.442666721296, 20.171250000000)
    pts(109) = Array(-225.346273525800, 19.743375000000)
    pts(110) = Array(-225.238764159970, 19.254375000000)
    pts(111) = Array(-225.147008235775, 18.826500000000)
    pts(112) = Array(-225.044780807190, 18.337500000000)
    pts(113) = Array(-224.957631446445, 17.909625000000)
    pts(114) = Array(-224.872621812643, 17.481750000000)
    pts(115) = Array(-224.789745787629, 17.053875000000)
    pts(116) = Array(-224.697635298031, 16.564875000000)
    pts(117) = Array(-224.619311505305, 16.137000000000)
    pts(118) = Array(-224.532388709621, 15.648000000000)
    pts(119) = Array(-224.458591558496, 15.220125000000)
    pts(120) = Array(-224.376828161348, 14.731125000000)
    pts(121) = Array(-224.307533715347, 14.303250000000)
    pts(122) = Array(-224.230903288386, 13.814250000000)
    pts(123) = Array(-224.166089224908, 13.386375000000)
    pts(124) = Array(-224.094567162538, 12.897375000000)
    pts(125) = Array(-224.034212735666, 12.469500000000)
    pts(126) = Array(-223.967776214256, 11.980500000000)
    pts(127) = Array(-223.911862220549, 11.552625000000)
    pts(128) = Array(-223.850490160991, 11.063625000000)
    pts(129) = Array(-223.798998908123, 10.635750000000)
    pts(130) = Array(-223.742671941567, 10.146750000000)
    pts(131) = Array(-223.695587219702, 9.718875000000)
    pts(132) = Array(-223.644287656348, 9.229875000000)
    pts(133) = Array(-223.601594712128, 8.802000000000)
    pts(134) = Array(-223.555306512955, 8.313000000000)
    pts(135) = Array(-223.516992026008, 7.885125000000)
    pts(136) = Array(-223.475700777325, 7.396125000000)
    pts(137) = Array(-223.441752839207, 6.968250000000)
    pts(138) = Array(-223.405445729938, 6.479250000000)
    pts(139) = Array(-223.375853825420, 6.051375000000)
    pts(140) = Array(-223.344519627036, 5.562375000000)
    pts(141) = Array(-223.319274617680, 5.134500000000)
    pts(142) = Array(-223.292903666703, 4.645500000000)
    pts(143) = Array(-223.271997776654, 4.217625000000)
    pts(144) = Array(-223.250581959669, 3.728625000000)
    pts(145) = Array(-223.234008763623, 3.300750000000)
    pts(146) = Array(-223.217541504742, 2.811750000000)
    pts(147) = Array(-223.205295918036, 2.383875000000)
    pts(148) = Array(-223.193772168775, 1.894875000000)
    pts(149) = Array(-223.185850439570, 1.467000000000)
    pts(150) = Array(-223.179266671096, 0.978000000000)
    pts(151) = Array(-223.175316639137, 0.489000000000)
    pts(152) = Array(-223.174000000000, 0.000000000000)
    Profile_19 = pts
End Function

Private Function Profile_20() As Variant
    Dim pts(258) As Variant
    pts(0) = Array(-204.564000000000, 0.000000000000)
    pts(1) = Array(-204.564713478518, 0.399875000000)
    pts(2) = Array(-204.566460792035, 0.742625000000)
    pts(3) = Array(-204.569824447264, 1.142500000000)
    pts(4) = Array(-204.573843492405, 1.485250000000)
    pts(5) = Array(-204.579857767629, 1.885125000000)
    pts(6) = Array(-204.586149074119, 2.227875000000)
    pts(7) = Array(-204.594814762184, 2.627750000000)
    pts(8) = Array(-204.603379159572, 2.970500000000)
    pts(9) = Array(-204.614697403408, 3.370375000000)
    pts(10) = Array(-204.625536021612, 3.713125000000)
    pts(11) = Array(-204.639508314990, 4.113000000000)
    pts(12) = Array(-204.652622585046, 4.455750000000)
    pts(13) = Array(-204.669250773562, 4.855625000000)
    pts(14) = Array(-204.684642428569, 5.198375000000)
    pts(15) = Array(-204.701085852721, 5.541125000000)
    pts(16) = Array(-204.718581509673, 5.883875000000)
    pts(17) = Array(-204.740323675009, 6.283750000000)
    pts(18) = Array(-204.760100903175, 6.626500000000)
    pts(19) = Array(-204.784506388302, 7.026375000000)
    pts(20) = Array(-204.806567811063, 7.369125000000)
    pts(21) = Array(-204.833639855841, 7.769000000000)
    pts(22) = Array(-204.857988404314, 8.111750000000)
    pts(23) = Array(-204.887730608742, 8.511625000000)
    pts(24) = Array(-204.914369523741, 8.854375000000)
    pts(25) = Array(-204.946785850370, 9.254250000000)
    pts(26) = Array(-204.975718684577, 9.597000000000)
    pts(27) = Array(-205.010813461161, 9.996875000000)
    pts(28) = Array(-205.042044081514, 10.339625000000)
    pts(29) = Array(-205.079822003934, 10.739500000000)
    pts(30) = Array(-205.113354594236, 11.082250000000)
    pts(31) = Array(-205.147951131675, 11.425000000000)
    pts(32) = Array(-205.183612607180, 11.767750000000)
    pts(33) = Array(-205.226564991516, 12.167625000000)
    pts(34) = Array(-205.264537383952, 12.510375000000)
    pts(35) = Array(-205.310188782322, 12.910250000000)
    pts(36) = Array(-205.350477218807, 13.253000000000)
    pts(37) = Array(-205.398833804433, 13.652875000000)
    pts(38) = Array(-205.441443741399, 13.995625000000)
    pts(39) = Array(-205.485127072491, 14.338375000000)
    pts(40) = Array(-205.529885063951, 14.681125000000)
    pts(41) = Array(-205.583462707425, 15.081000000000)
    pts(42) = Array(-205.630553632137, 15.423750000000)
    pts(43) = Array(-205.686856739869, 15.823625000000)
    pts(44) = Array(-205.736287039379, 16.166375000000)
    pts(45) = Array(-205.795323327410, 16.566250000000)
    pts(46) = Array(-205.847099788804, 16.909000000000)
    pts(47) = Array(-205.899961480044, 17.251750000000)
    pts(48) = Array(-205.953909956257, 17.594500000000)
    pts(49) = Array(-206.018225544025, 17.994375000000)
    pts(50) = Array(-206.074534223231, 18.337125000000)
    pts(51) = Array(-206.141607894504, 18.737000000000)
    pts(52) = Array(-206.200284602311, 19.079750000000)
    pts(53) = Array(-206.270125699615, 19.479625000000)
    pts(54) = Array(-206.331178627584, 19.822375000000)
    pts(55) = Array(-206.393331138074, 20.165125000000)
    pts(56) = Array(-206.456585092171, 20.507875000000)
    pts(57) = Array(-206.531776009014, 20.907750000000)
    pts(58) = Array(-206.597422986934, 21.250500000000)
    pts(59) = Array(-206.664177549379, 21.593250000000)
    pts(60) = Array(-206.732041714829, 21.936000000000)
    pts(61) = Array(-206.812621735933, 22.335875000000)
    pts(62) = Array(-206.882897150036, 22.678625000000)
    pts(63) = Array(-206.954288815113, 23.021375000000)
    pts(64) = Array(-207.026798912605, 23.364125000000)
    pts(65) = Array(-207.112810566515, 23.764000000000)
    pts(66) = Array(-207.187751608857, 24.106750000000)
    pts(67) = Array(-207.263818257098, 24.449500000000)
    pts(68) = Array(-207.341012861744, 24.792250000000)
    pts(69) = Array(-207.432502038977, 25.192125000000)
    pts(70) = Array(-207.512148811224, 25.534875000000)
    pts(71) = Array(-207.592931259151, 25.877625000000)
    pts(72) = Array(-207.674851909031, 26.220375000000)
    pts(73) = Array(-207.771867991805, 26.620250000000)
    pts(74) = Array(-207.856263619131, 26.963000000000)
    pts(75) = Array(-207.941805735700, 27.305750000000)
    pts(76) = Array(-208.028497050901, 27.648500000000)
    pts(77) = Array(-208.131093055347, 28.048375000000)
    pts(78) = Array(-208.220283811248, 28.391125000000)
    pts(79) = Array(-208.310632645419, 28.733875000000)
    pts(80) = Array(-208.402142458399, 29.076625000000)
    pts(81) = Array(-208.510375190271, 29.476500000000)
    pts(82) = Array(-208.604410633193, 29.819250000000)
    pts(83) = Array(-208.699616553530, 30.162000000000)
    pts(84) = Array(-208.795996051762, 30.504750000000)
    pts(85) = Array(-208.909926276778, 30.904625000000)
    pts(86) = Array(-209.008859399558, 31.247375000000)
    pts(87) = Array(-209.108976246913, 31.590125000000)
    pts(88) = Array(-209.210280128870, 31.932875000000)
    pts(89) = Array(-209.329972758960, 32.332750000000)
    pts(90) = Array(-209.433860152209, 32.675500000000)
    pts(91) = Array(-209.538945406630, 33.018250000000)
    pts(92) = Array(-209.645232052301, 33.361000000000)
    pts(93) = Array(-209.752723673779, 33.703750000000)
    pts(94) = Array(-209.861423910780, 34.046500000000)
    pts(95) = Array(-209.971336458864, 34.389250000000)
    pts(96) = Array(-210.082465070145, 34.732000000000)
    pts(97) = Array(-210.213657170737, 35.131875000000)
    pts(98) = Array(-210.327433731343, 35.474625000000)
    pts(99) = Array(-210.442438618197, 35.817375000000)
    pts(100) = Array(-210.558675826972, 36.160125000000)
    pts(101) = Array(-210.676149413790, 36.502875000000)
    pts(102) = Array(-210.794863496018, 36.845625000000)
    pts(103) = Array(-210.914822253099, 37.188375000000)
    pts(104) = Array(-211.036029927388, 37.531125000000)
    pts(105) = Array(-211.158490825012, 37.873875000000)
    pts(106) = Array(-211.282209316751, 38.216625000000)
    pts(107) = Array(-211.407189838931, 38.559375000000)
    pts(108) = Array(-211.533436894348, 38.902125000000)
    pts(109) = Array(-211.660955053199, 39.244875000000)
    pts(110) = Array(-211.789748954048, 39.587625000000)
    pts(111) = Array(-211.919823304801, 39.930375000000)
    pts(112) = Array(-212.051182883710, 40.273125000000)
    pts(113) = Array(-212.183832540401, 40.615875000000)
    pts(114) = Array(-212.317777196918, 40.958625000000)
    pts(115) = Array(-212.453021848798, 41.301375000000)
    pts(116) = Array(-212.589571566164, 41.644125000000)
    pts(117) = Array(-212.727431494850, 41.986875000000)
    pts(118) = Array(-212.866606857544, 42.329625000000)
    pts(119) = Array(-213.007102954964, 42.672375000000)
    pts(120) = Array(-213.148925167056, 43.015125000000)
    pts(121) = Array(-213.292078954223, 43.357875000000)
    pts(122) = Array(-213.436569858581, 43.700625000000)
    pts(123) = Array(-213.582403505243, 44.043375000000)
    pts(124) = Array(-213.729585603636, 44.386125000000)
    pts(125) = Array(-213.878121948846, 44.728875000000)
    pts(126) = Array(-214.028018422996, 45.071625000000)
    pts(127) = Array(-214.179280996659, 45.414375000000)
    pts(128) = Array(-214.306381059205, 45.700000000000)
    pts(129) = Array(-205.578621189497, 45.700000000000)
    pts(130) = Array(-205.578621189497, 42.500000000000)
    pts(131) = Array(-205.399227294829, 42.158250000000)
    pts(132) = Array(-205.250708018915, 41.872625000000)
    pts(133) = Array(-205.074172711344, 41.529875000000)
    pts(134) = Array(-204.899467332593, 41.187125000000)
    pts(135) = Array(-204.755268872106, 40.901500000000)
    pts(136) = Array(-204.583887921555, 40.558750000000)
    pts(137) = Array(-204.414304089265, 40.216000000000)
    pts(138) = Array(-204.246506209124, 39.873250000000)
    pts(139) = Array(-204.108030999776, 39.587625000000)
    pts(140) = Array(-203.943479053069, 39.244875000000)
    pts(141) = Array(-203.780682538412, 38.902125000000)
    pts(142) = Array(-203.619631064798, 38.559375000000)
    pts(143) = Array(-203.460314433143, 38.216625000000)
    pts(144) = Array(-203.302722632282, 37.873875000000)
    pts(145) = Array(-203.146845835076, 37.531125000000)
    pts(146) = Array(-202.992674394630, 37.188375000000)
    pts(147) = Array(-202.865494059753, 36.902750000000)
    pts(148) = Array(-202.714424635129, 36.560000000000)
    pts(149) = Array(-202.565034181586, 36.217250000000)
    pts(150) = Array(-202.417313710120, 35.874500000000)
    pts(151) = Array(-202.271254395195, 35.531750000000)
    pts(152) = Array(-202.126847571566, 35.189000000000)
    pts(153) = Array(-201.984084731169, 34.846250000000)
    pts(154) = Array(-201.842957520120, 34.503500000000)
    pts(155) = Array(-201.703457735785, 34.160750000000)
    pts(156) = Array(-201.565577323935, 33.818000000000)
    pts(157) = Array(-201.429308375979, 33.475250000000)
    pts(158) = Array(-201.294643126280, 33.132500000000)
    pts(159) = Array(-201.161573949531, 32.789750000000)
    pts(160) = Array(-201.030093358218, 32.447000000000)
    pts(161) = Array(-200.900194000141, 32.104250000000)
    pts(162) = Array(-200.771868656008, 31.761500000000)
    pts(163) = Array(-200.645110237094, 31.418750000000)
    pts(164) = Array(-200.519911782960, 31.076000000000)
    pts(165) = Array(-200.396266459235, 30.733250000000)
    pts(166) = Array(-200.274167555461, 30.390500000000)
    pts(167) = Array(-200.153608482987, 30.047750000000)
    pts(168) = Array(-200.034582772929, 29.705000000000)
    pts(169) = Array(-199.917084074176, 29.362250000000)
    pts(170) = Array(-199.781923919430, 28.962375000000)
    pts(171) = Array(-199.667712506640, 28.619625000000)
    pts(172) = Array(-199.555008748588, 28.276875000000)
    pts(173) = Array(-199.443806764607, 27.934125000000)
    pts(174) = Array(-199.334100782325, 27.591375000000)
    pts(175) = Array(-199.225885135981, 27.248625000000)
    pts(176) = Array(-199.119154264776, 26.905875000000)
    pts(177) = Array(-199.013902711277, 26.563125000000)
    pts(178) = Array(-198.892971786019, 26.163250000000)
    pts(179) = Array(-198.790907181795, 25.820500000000)
    pts(180) = Array(-198.690305277832, 25.477750000000)
    pts(181) = Array(-198.591161031038, 25.135000000000)
    pts(182) = Array(-198.477328464766, 24.735125000000)
    pts(183) = Array(-198.381325628784, 24.392375000000)
    pts(184) = Array(-198.286765105281, 24.049625000000)
    pts(185) = Array(-198.193642241215, 23.706875000000)
    pts(186) = Array(-198.086809859360, 23.307000000000)
    pts(187) = Array(-197.996786390038, 22.964250000000)
    pts(188) = Array(-197.908186435375, 22.621500000000)
    pts(189) = Array(-197.821005710611, 22.278750000000)
    pts(190) = Array(-197.721083016268, 21.878875000000)
    pts(191) = Array(-197.636962995914, 21.536125000000)
    pts(192) = Array(-197.554249195312, 21.193375000000)
    pts(193) = Array(-197.472937678472, 20.850625000000)
    pts(194) = Array(-197.379841434570, 20.450750000000)
    pts(195) = Array(-197.314506156316, 20.165125000000)
    pts(196) = Array(-197.237378680301, 19.822375000000)
    pts(197) = Array(-197.174165708801, 19.536750000000)
    pts(198) = Array(-197.099578982652, 19.194000000000)
    pts(199) = Array(-197.038478383243, 18.908375000000)
    pts(200) = Array(-196.966420721042, 18.565625000000)
    pts(201) = Array(-196.907422949893, 18.280000000000)
    pts(202) = Array(-196.837883129927, 17.937250000000)
    pts(203) = Array(-196.780979026151, 17.651625000000)
    pts(204) = Array(-196.713946281684, 17.308875000000)
    pts(205) = Array(-196.659127059835, 17.023250000000)
    pts(206) = Array(-196.594591070328, 16.680500000000)
    pts(207) = Array(-196.541848313270, 16.394875000000)
    pts(208) = Array(-196.479799196046, 16.052125000000)
    pts(209) = Array(-196.419102543258, 15.709375000000)
    pts(210) = Array(-196.359755611474, 15.366625000000)
    pts(211) = Array(-196.311328939767, 15.081000000000)
    pts(212) = Array(-196.254449597012, 14.738250000000)
    pts(213) = Array(-196.208075631381, 14.452625000000)
    pts(214) = Array(-196.153655300115, 14.109875000000)
    pts(215) = Array(-196.109327063306, 13.824250000000)
    pts(216) = Array(-196.057357573986, 13.481500000000)
    pts(217) = Array(-196.015068426071, 13.195875000000)
    pts(218) = Array(-195.965542010822, 12.853125000000)
    pts(219) = Array(-195.925285644063, 12.567500000000)
    pts(220) = Array(-195.878194930694, 12.224750000000)
    pts(221) = Array(-195.839965364708, 11.939125000000)
    pts(222) = Array(-195.795303371079, 11.596375000000)
    pts(223) = Array(-195.759094948275, 11.310750000000)
    pts(224) = Array(-195.716855076988, 10.968000000000)
    pts(225) = Array(-195.675933718626, 10.625250000000)
    pts(226) = Array(-195.636329084430, 10.282500000000)
    pts(227) = Array(-195.604329804991, 9.996875000000)
    pts(228) = Array(-195.567134720262, 9.654125000000)
    pts(229) = Array(-195.537141093222, 9.368500000000)
    pts(230) = Array(-195.502350119642, 9.025750000000)
    pts(231) = Array(-195.474357766117, 8.740125000000)
    pts(232) = Array(-195.441965831680, 8.397375000000)
    pts(233) = Array(-195.410880456066, 8.054625000000)
    pts(234) = Array(-195.381100296573, 7.711875000000)
    pts(235) = Array(-195.357279608575, 7.426250000000)
    pts(236) = Array(-195.329889050642, 7.083500000000)
    pts(237) = Array(-195.308058008341, 6.797875000000)
    pts(238) = Array(-195.283053089739, 6.455125000000)
    pts(239) = Array(-195.263208538549, 6.169500000000)
    pts(240) = Array(-195.240585649793, 5.826750000000)
    pts(241) = Array(-195.219260654002, 5.484000000000)
    pts(242) = Array(-195.199232638003, 5.141250000000)
    pts(243) = Array(-195.183532753257, 4.855625000000)
    pts(244) = Array(-195.165880348611, 4.512875000000)
    pts(245) = Array(-195.152159046924, 4.227250000000)
    pts(246) = Array(-195.136879719005, 3.884500000000)
    pts(247) = Array(-195.125135031559, 3.598875000000)
    pts(248) = Array(-195.112226589392, 3.256125000000)
    pts(249) = Array(-195.100610556557, 2.913375000000)
    pts(250) = Array(-195.090286438331, 2.570625000000)
    pts(251) = Array(-195.082669566963, 2.285000000000)
    pts(252) = Array(-195.074712857186, 1.942250000000)
    pts(253) = Array(-195.069068315773, 1.656625000000)
    pts(254) = Array(-195.063477876790, 1.313875000000)
    pts(255) = Array(-195.059804858514, 1.028250000000)
    pts(256) = Array(-195.056579891438, 0.685500000000)
    pts(257) = Array(-195.054644966008, 0.342750000000)
    pts(258) = Array(-195.054000000000, 0.000000000000)
    Profile_20 = pts
End Function

Private Function Profile_21() As Variant
    Dim pts(195) As Variant
    pts(0) = Array(-195.054000000000, 0.000000000000)
    pts(1) = Array(-195.054644966008, 0.342750000000)
    pts(2) = Array(-195.056579891438, 0.685500000000)
    pts(3) = Array(-195.059804858514, 1.028250000000)
    pts(4) = Array(-195.063477876790, 1.313875000000)
    pts(5) = Array(-195.069068315773, 1.656625000000)
    pts(6) = Array(-195.074712857186, 1.942250000000)
    pts(7) = Array(-195.082669566963, 2.285000000000)
    pts(8) = Array(-195.090286438331, 2.570625000000)
    pts(9) = Array(-195.100610556557, 2.913375000000)
    pts(10) = Array(-195.112226589392, 3.256125000000)
    pts(11) = Array(-195.125135031559, 3.598875000000)
    pts(12) = Array(-195.136879719005, 3.884500000000)
    pts(13) = Array(-195.152159046924, 4.227250000000)
    pts(14) = Array(-195.165880348611, 4.512875000000)
    pts(15) = Array(-195.183532753257, 4.855625000000)
    pts(16) = Array(-195.199232638003, 5.141250000000)
    pts(17) = Array(-195.219260654002, 5.484000000000)
    pts(18) = Array(-195.240585649793, 5.826750000000)
    pts(19) = Array(-195.263208538549, 6.169500000000)
    pts(20) = Array(-195.283053089739, 6.455125000000)
    pts(21) = Array(-195.308058008341, 6.797875000000)
    pts(22) = Array(-195.329889050642, 7.083500000000)
    pts(23) = Array(-195.357279608575, 7.426250000000)
    pts(24) = Array(-195.381100296573, 7.711875000000)
    pts(25) = Array(-195.410880456066, 8.054625000000)
    pts(26) = Array(-195.441965831680, 8.397375000000)
    pts(27) = Array(-195.474357766117, 8.740125000000)
    pts(28) = Array(-195.502350119642, 9.025750000000)
    pts(29) = Array(-195.537141093222, 9.368500000000)
    pts(30) = Array(-195.567134720262, 9.654125000000)
    pts(31) = Array(-195.604329804991, 9.996875000000)
    pts(32) = Array(-195.636329084430, 10.282500000000)
    pts(33) = Array(-195.675933718626, 10.625250000000)
    pts(34) = Array(-195.716855076988, 10.968000000000)
    pts(35) = Array(-195.759094948275, 11.310750000000)
    pts(36) = Array(-195.795303371079, 11.596375000000)
    pts(37) = Array(-195.839965364708, 11.939125000000)
    pts(38) = Array(-195.878194930694, 12.224750000000)
    pts(39) = Array(-195.925285644063, 12.567500000000)
    pts(40) = Array(-195.965542010822, 12.853125000000)
    pts(41) = Array(-196.015068426071, 13.195875000000)
    pts(42) = Array(-196.057357573986, 13.481500000000)
    pts(43) = Array(-196.109327063306, 13.824250000000)
    pts(44) = Array(-196.153655300115, 14.109875000000)
    pts(45) = Array(-196.208075631381, 14.452625000000)
    pts(46) = Array(-196.254449597012, 14.738250000000)
    pts(47) = Array(-196.311328939767, 15.081000000000)
    pts(48) = Array(-196.359755611474, 15.366625000000)
    pts(49) = Array(-196.419102543258, 15.709375000000)
    pts(50) = Array(-196.479799196046, 16.052125000000)
    pts(51) = Array(-196.541848313270, 16.394875000000)
    pts(52) = Array(-196.594591070328, 16.680500000000)
    pts(53) = Array(-196.659127059835, 17.023250000000)
    pts(54) = Array(-196.713946281684, 17.308875000000)
    pts(55) = Array(-196.780979026151, 17.651625000000)
    pts(56) = Array(-196.837883129927, 17.937250000000)
    pts(57) = Array(-196.907422949893, 18.280000000000)
    pts(58) = Array(-196.966420721042, 18.565625000000)
    pts(59) = Array(-197.038478383243, 18.908375000000)
    pts(60) = Array(-197.099578982652, 19.194000000000)
    pts(61) = Array(-197.174165708801, 19.536750000000)
    pts(62) = Array(-197.237378680301, 19.822375000000)
    pts(63) = Array(-197.314506156316, 20.165125000000)
    pts(64) = Array(-197.379841434570, 20.450750000000)
    pts(65) = Array(-197.472937678472, 20.850625000000)
    pts(66) = Array(-197.554249195312, 21.193375000000)
    pts(67) = Array(-197.636962995914, 21.536125000000)
    pts(68) = Array(-197.721083016268, 21.878875000000)
    pts(69) = Array(-197.821005710611, 22.278750000000)
    pts(70) = Array(-197.908186435375, 22.621500000000)
    pts(71) = Array(-197.996786390038, 22.964250000000)
    pts(72) = Array(-198.086809859360, 23.307000000000)
    pts(73) = Array(-198.193642241215, 23.706875000000)
    pts(74) = Array(-198.286765105281, 24.049625000000)
    pts(75) = Array(-198.381325628784, 24.392375000000)
    pts(76) = Array(-198.477328464766, 24.735125000000)
    pts(77) = Array(-198.591161031038, 25.135000000000)
    pts(78) = Array(-198.690305277832, 25.477750000000)
    pts(79) = Array(-198.790907181795, 25.820500000000)
    pts(80) = Array(-198.892971786019, 26.163250000000)
    pts(81) = Array(-199.013902711277, 26.563125000000)
    pts(82) = Array(-199.119154264776, 26.905875000000)
    pts(83) = Array(-199.225885135981, 27.248625000000)
    pts(84) = Array(-199.334100782325, 27.591375000000)
    pts(85) = Array(-199.443806764607, 27.934125000000)
    pts(86) = Array(-199.555008748588, 28.276875000000)
    pts(87) = Array(-199.667712506640, 28.619625000000)
    pts(88) = Array(-199.781923919430, 28.962375000000)
    pts(89) = Array(-199.917084074176, 29.362250000000)
    pts(90) = Array(-200.034582772929, 29.705000000000)
    pts(91) = Array(-200.153608482987, 30.047750000000)
    pts(92) = Array(-200.274167555461, 30.390500000000)
    pts(93) = Array(-200.396266459235, 30.733250000000)
    pts(94) = Array(-200.519911782960, 31.076000000000)
    pts(95) = Array(-200.645110237094, 31.418750000000)
    pts(96) = Array(-200.771868656008, 31.761500000000)
    pts(97) = Array(-200.900194000141, 32.104250000000)
    pts(98) = Array(-201.030093358218, 32.447000000000)
    pts(99) = Array(-201.161573949531, 32.789750000000)
    pts(100) = Array(-201.294643126280, 33.132500000000)
    pts(101) = Array(-201.429308375979, 33.475250000000)
    pts(102) = Array(-201.565577323935, 33.818000000000)
    pts(103) = Array(-201.703457735785, 34.160750000000)
    pts(104) = Array(-201.842957520120, 34.503500000000)
    pts(105) = Array(-201.984084731169, 34.846250000000)
    pts(106) = Array(-202.126847571566, 35.189000000000)
    pts(107) = Array(-202.271254395195, 35.531750000000)
    pts(108) = Array(-202.417313710120, 35.874500000000)
    pts(109) = Array(-202.565034181586, 36.217250000000)
    pts(110) = Array(-202.714424635129, 36.560000000000)
    pts(111) = Array(-202.865494059753, 36.902750000000)
    pts(112) = Array(-202.992674394630, 37.188375000000)
    pts(113) = Array(-203.146845835076, 37.531125000000)
    pts(114) = Array(-203.302722632282, 37.873875000000)
    pts(115) = Array(-203.460314433143, 38.216625000000)
    pts(116) = Array(-203.619631064798, 38.559375000000)
    pts(117) = Array(-203.780682538412, 38.902125000000)
    pts(118) = Array(-203.943479053069, 39.244875000000)
    pts(119) = Array(-204.108030999776, 39.587625000000)
    pts(120) = Array(-204.246506209124, 39.873250000000)
    pts(121) = Array(-204.414304089265, 40.216000000000)
    pts(122) = Array(-204.583887921555, 40.558750000000)
    pts(123) = Array(-204.755268872106, 40.901500000000)
    pts(124) = Array(-204.899467332593, 41.187125000000)
    pts(125) = Array(-205.074172711344, 41.529875000000)
    pts(126) = Array(-205.250708018915, 41.872625000000)
    pts(127) = Array(-205.399227294829, 42.158250000000)
    pts(128) = Array(-205.578621189497, 42.500000000000)
    pts(129) = Array(-205.578621189497, 45.700000000000)
    pts(130) = Array(-189.074000000000, 45.700000000000)
    pts(131) = Array(-187.939202172987, 41.500000000000)
    pts(132) = Array(-188.063628698565, 40.844375000000)
    pts(133) = Array(-188.180958712339, 40.216000000000)
    pts(134) = Array(-188.296407752541, 39.587625000000)
    pts(135) = Array(-188.420210437979, 38.902125000000)
    pts(136) = Array(-188.531736154363, 38.273750000000)
    pts(137) = Array(-188.641389984413, 37.645375000000)
    pts(138) = Array(-188.749174756810, 37.017000000000)
    pts(139) = Array(-188.864629751599, 36.331500000000)
    pts(140) = Array(-188.968515398487, 35.703125000000)
    pts(141) = Array(-189.070540392605, 35.074750000000)
    pts(142) = Array(-189.170707344986, 34.446375000000)
    pts(143) = Array(-189.277864300284, 33.760875000000)
    pts(144) = Array(-189.374154463544, 33.132500000000)
    pts(145) = Array(-189.468594326338, 32.504125000000)
    pts(146) = Array(-189.561186287838, 31.875750000000)
    pts(147) = Array(-189.660090934505, 31.190250000000)
    pts(148) = Array(-189.748826630850, 30.561875000000)
    pts(149) = Array(-189.835721520611, 29.933500000000)
    pts(150) = Array(-189.920777796350, 29.305125000000)
    pts(151) = Array(-190.011472045682, 28.619625000000)
    pts(152) = Array(-190.092690812751, 27.991250000000)
    pts(153) = Array(-190.172077429394, 27.362875000000)
    pts(154) = Array(-190.249633886310, 26.734500000000)
    pts(155) = Array(-190.332155921303, 26.049000000000)
    pts(156) = Array(-190.405891899169, 25.420625000000)
    pts(157) = Array(-190.477803563711, 24.792250000000)
    pts(158) = Array(-190.547892708034, 24.163875000000)
    pts(159) = Array(-190.622277065645, 23.478375000000)
    pts(160) = Array(-190.688561070175, 22.850000000000)
    pts(161) = Array(-190.758798350279, 22.164500000000)
    pts(162) = Array(-190.821284412989, 21.536125000000)
    pts(163) = Array(-190.887382233380, 20.850625000000)
    pts(164) = Array(-190.946077123198, 20.222250000000)
    pts(165) = Array(-191.002961111830, 19.593875000000)
    pts(166) = Array(-191.058035603974, 18.965500000000)
    pts(167) = Array(-191.116054740107, 18.280000000000)
    pts(168) = Array(-191.167350075955, 17.651625000000)
    pts(169) = Array(-191.221249563782, 16.966125000000)
    pts(170) = Array(-191.268771264572, 16.337750000000)
    pts(171) = Array(-191.318556900089, 15.652250000000)
    pts(172) = Array(-191.362310071284, 15.023875000000)
    pts(173) = Array(-191.404262586298, 14.395500000000)
    pts(174) = Array(-191.444415473446, 13.767125000000)
    pts(175) = Array(-191.486167309351, 13.081625000000)
    pts(176) = Array(-191.522560463666, 12.453250000000)
    pts(177) = Array(-191.560212950205, 11.767750000000)
    pts(178) = Array(-191.592850263194, 11.139375000000)
    pts(179) = Array(-191.626407423829, 10.453875000000)
    pts(180) = Array(-191.655292380822, 9.825500000000)
    pts(181) = Array(-191.682383846696, 9.197125000000)
    pts(182) = Array(-191.707682482068, 8.568750000000)
    pts(183) = Array(-191.733237000977, 7.883250000000)
    pts(184) = Array(-191.754788931634, 7.254875000000)
    pts(185) = Array(-191.776257441720, 6.569375000000)
    pts(186) = Array(-191.794064962208, 5.941000000000)
    pts(187) = Array(-191.811449748243, 5.255500000000)
    pts(188) = Array(-191.825514753041, 4.627125000000)
    pts(189) = Array(-191.838817663943, 3.941625000000)
    pts(190) = Array(-191.849141648476, 3.313250000000)
    pts(191) = Array(-191.858364098290, 2.627750000000)
    pts(192) = Array(-191.864948159720, 1.999375000000)
    pts(193) = Array(-191.870091128372, 1.313875000000)
    pts(194) = Array(-191.872935966184, 0.685500000000)
    pts(195) = Array(-191.874000000000, 0.000000000000)
    Profile_21 = pts
End Function

Private Function Profile_22() As Variant
    Dim pts(164) As Variant
    pts(0) = Array(-186.324000000000, 0.000000000000)
    pts(1) = Array(-186.325528306981, 1.437500000000)
    pts(2) = Array(-186.330113248653, 2.875000000000)
    pts(3) = Array(-186.337754887209, 4.312500000000)
    pts(4) = Array(-186.348453326306, 5.750000000000)
    pts(5) = Array(-186.362208711074, 7.187500000000)
    pts(6) = Array(-186.379021228125, 8.625000000000)
    pts(7) = Array(-186.398891105568, 10.062500000000)
    pts(8) = Array(-186.421818613022, 11.500000000000)
    pts(9) = Array(-186.447804061635, 12.937500000000)
    pts(10) = Array(-186.476847804104, 14.375000000000)
    pts(11) = Array(-186.508950234703, 15.812500000000)
    pts(12) = Array(-186.544111789304, 17.250000000000)
    pts(13) = Array(-186.582332945411, 18.687500000000)
    pts(14) = Array(-186.623614222191, 20.125000000000)
    pts(15) = Array(-186.667956180509, 21.562500000000)
    pts(16) = Array(-186.715359422966, 23.000000000000)
    pts(17) = Array(-186.765824593941, 24.437500000000)
    pts(18) = Array(-186.819352379634, 25.875000000000)
    pts(19) = Array(-186.875943508114, 27.312500000000)
    pts(20) = Array(-186.935598749366, 28.750000000000)
    pts(21) = Array(-186.998318915345, 30.187500000000)
    pts(22) = Array(-187.064104860034, 31.625000000000)
    pts(23) = Array(-187.132957479497, 33.062500000000)
    pts(24) = Array(-187.204877711945, 34.500000000000)
    pts(25) = Array(-187.279866537795, 35.937500000000)
    pts(26) = Array(-187.357924979745, 37.375000000000)
    pts(27) = Array(-187.439054102835, 38.812500000000)
    pts(28) = Array(-187.523255014525, 40.250000000000)
    pts(29) = Array(-187.610528864772, 41.687500000000)
    pts(30) = Array(-187.700876846104, 43.125000000000)
    pts(31) = Array(-187.794300193707, 44.562500000000)
    pts(32) = Array(-187.890800185506, 46.000000000000)
    pts(33) = Array(-187.274000000000, 47.000000000000)
    pts(34) = Array(-187.274000000000, 50.000000000000)
    pts(35) = Array(-185.948064643883, 50.000000000000)
    pts(36) = Array(-185.948064643883, 46.000000000000)
    pts(37) = Array(-185.755377208934, 45.687500000000)
    pts(38) = Array(-185.564495633841, 45.375000000000)
    pts(39) = Array(-185.375406098249, 45.062500000000)
    pts(40) = Array(-185.188095050254, 44.750000000000)
    pts(41) = Array(-185.002549199803, 44.437500000000)
    pts(42) = Array(-184.818755512324, 44.125000000000)
    pts(43) = Array(-184.636701202554, 43.812500000000)
    pts(44) = Array(-184.420514367156, 43.437500000000)
    pts(45) = Array(-184.242242874836, 43.125000000000)
    pts(46) = Array(-184.065671454379, 42.812500000000)
    pts(47) = Array(-183.890788304581, 42.500000000000)
    pts(48) = Array(-183.683140751327, 42.125000000000)
    pts(49) = Array(-183.511931334709, 41.812500000000)
    pts(50) = Array(-183.342373880353, 41.500000000000)
    pts(51) = Array(-183.174457478933, 41.187500000000)
    pts(52) = Array(-182.975108922268, 40.812500000000)
    pts(53) = Array(-182.810765415785, 40.500000000000)
    pts(54) = Array(-182.648029375080, 40.187500000000)
    pts(55) = Array(-182.486890701882, 39.875000000000)
    pts(56) = Array(-182.295618870069, 39.500000000000)
    pts(57) = Array(-182.137959752071, 39.187500000000)
    pts(58) = Array(-181.950835445861, 38.812500000000)
    pts(59) = Array(-181.796609759403, 38.500000000000)
    pts(60) = Array(-181.613578716440, 38.125000000000)
    pts(61) = Array(-181.462742127795, 37.812500000000)
    pts(62) = Array(-181.313431847598, 37.500000000000)
    pts(63) = Array(-181.165639225535, 37.187500000000)
    pts(64) = Array(-180.990279432959, 36.812500000000)
    pts(65) = Array(-180.845795955150, 36.500000000000)
    pts(66) = Array(-180.674383172677, 36.125000000000)
    pts(67) = Array(-180.533169272976, 35.812500000000)
    pts(68) = Array(-180.365656932222, 35.437500000000)
    pts(69) = Array(-180.227674547073, 35.125000000000)
    pts(70) = Array(-180.064017844444, 34.750000000000)
    pts(71) = Array(-179.929230349955, 34.437500000000)
    pts(72) = Array(-179.769386173409, 34.062500000000)
    pts(73) = Array(-179.637758326146, 33.750000000000)
    pts(74) = Array(-179.481685186259, 33.375000000000)
    pts(75) = Array(-179.353183067507, 33.062500000000)
    pts(76) = Array(-179.200841032564, 32.687500000000)
    pts(77) = Array(-179.075431995826, 32.375000000000)
    pts(78) = Array(-178.926782630653, 32.000000000000)
    pts(79) = Array(-178.804435252157, 31.687500000000)
    pts(80) = Array(-178.659441560454, 31.312500000000)
    pts(81) = Array(-178.540125592485, 31.000000000000)
    pts(82) = Array(-178.398751962432, 30.625000000000)
    pts(83) = Array(-178.259338380285, 30.250000000000)
    pts(84) = Array(-178.121874858464, 29.875000000000)
    pts(85) = Array(-178.008804490025, 29.562500000000)
    pts(86) = Array(-177.874890819802, 29.187500000000)
    pts(87) = Array(-177.764765345572, 28.875000000000)
    pts(88) = Array(-177.634369882259, 28.500000000000)
    pts(89) = Array(-177.527163432741, 28.187500000000)
    pts(90) = Array(-177.400255723454, 27.812500000000)
    pts(91) = Array(-177.275238003200, 27.437500000000)
    pts(92) = Array(-177.152101736688, 27.062500000000)
    pts(93) = Array(-177.050919368651, 26.750000000000)
    pts(94) = Array(-176.931210856518, 26.375000000000)
    pts(95) = Array(-176.813360574244, 26.000000000000)
    pts(96) = Array(-176.697360654161, 25.625000000000)
    pts(97) = Array(-176.602101964500, 25.312500000000)
    pts(98) = Array(-176.489474492756, 24.937500000000)
    pts(99) = Array(-176.378675988136, 24.562500000000)
    pts(100) = Array(-176.269699207328, 24.187500000000)
    pts(101) = Array(-176.180271700372, 23.875000000000)
    pts(102) = Array(-176.074616451199, 23.500000000000)
    pts(103) = Array(-175.970763243833, 23.125000000000)
    pts(104) = Array(-175.868705420759, 22.750000000000)
    pts(105) = Array(-175.768436471466, 22.375000000000)
    pts(106) = Array(-175.669950030182, 22.000000000000)
    pts(107) = Array(-175.573239873668, 21.625000000000)
    pts(108) = Array(-175.478299919089, 21.250000000000)
    pts(109) = Array(-175.385124221936, 20.875000000000)
    pts(110) = Array(-175.293706974018, 20.500000000000)
    pts(111) = Array(-175.204042501511, 20.125000000000)
    pts(112) = Array(-175.116125263065, 19.750000000000)
    pts(113) = Array(-175.029949847968, 19.375000000000)
    pts(114) = Array(-174.945510974366, 19.000000000000)
    pts(115) = Array(-174.862803487536, 18.625000000000)
    pts(116) = Array(-174.781822358208, 18.250000000000)
    pts(117) = Array(-174.702562680947, 17.875000000000)
    pts(118) = Array(-174.625019672574, 17.500000000000)
    pts(119) = Array(-174.549188670642, 17.125000000000)
    pts(120) = Array(-174.475065131957, 16.750000000000)
    pts(121) = Array(-174.402644631147, 16.375000000000)
    pts(122) = Array(-174.331922859274, 16.000000000000)
    pts(123) = Array(-174.262895622489, 15.625000000000)
    pts(124) = Array(-174.195558840734, 15.250000000000)
    pts(125) = Array(-174.129908546481, 14.875000000000)
    pts(126) = Array(-174.065940883517, 14.500000000000)
    pts(127) = Array(-174.003652105764, 14.125000000000)
    pts(128) = Array(-173.943038576141, 13.750000000000)
    pts(129) = Array(-173.884096765466, 13.375000000000)
    pts(130) = Array(-173.826823251391, 13.000000000000)
    pts(131) = Array(-173.771214717379, 12.625000000000)
    pts(132) = Array(-173.708438165181, 12.187500000000)
    pts(133) = Array(-173.656426208514, 11.812500000000)
    pts(134) = Array(-173.606069416565, 11.437500000000)
    pts(135) = Array(-173.557364903416, 11.062500000000)
    pts(136) = Array(-173.510309884221, 10.687500000000)
    pts(137) = Array(-173.464901674389, 10.312500000000)
    pts(138) = Array(-173.421137688805, 9.937500000000)
    pts(139) = Array(-173.379015441076, 9.562500000000)
    pts(140) = Array(-173.338532542818, 9.187500000000)
    pts(141) = Array(-173.299686702962, 8.812500000000)
    pts(142) = Array(-173.262475727098, 8.437500000000)
    pts(143) = Array(-173.226897516847, 8.062500000000)
    pts(144) = Array(-173.192950069261, 7.687500000000)
    pts(145) = Array(-173.160631476251, 7.312500000000)
    pts(146) = Array(-173.129939924048, 6.937500000000)
    pts(147) = Array(-173.100873692686, 6.562500000000)
    pts(148) = Array(-173.069015149430, 6.125000000000)
    pts(149) = Array(-173.043464990111, 5.750000000000)
    pts(150) = Array(-173.019535318567, 5.375000000000)
    pts(151) = Array(-172.997224800758, 5.000000000000)
    pts(152) = Array(-172.976532194229, 4.625000000000)
    pts(153) = Array(-172.957456347762, 4.250000000000)
    pts(154) = Array(-172.939996201052, 3.875000000000)
    pts(155) = Array(-172.924150784415, 3.500000000000)
    pts(156) = Array(-172.907704136286, 3.062500000000)
    pts(157) = Array(-172.895354405056, 2.687500000000)
    pts(158) = Array(-172.884616928150, 2.312500000000)
    pts(159) = Array(-172.875491110844, 1.937500000000)
    pts(160) = Array(-172.867976447909, 1.562500000000)
    pts(161) = Array(-172.862072523473, 1.187500000000)
    pts(162) = Array(-172.857779010904, 0.812500000000)
    pts(163) = Array(-172.855095672722, 0.437500000000)
    pts(164) = Array(-172.854000000000, 0.000000000000)
    Profile_22 = pts
End Function

Private Function Profile_23() As Variant
    Dim pts(113) As Variant
    pts(0) = Array(-172.444000000000, 0.000000000000)
    pts(1) = Array(-172.442172333488, 1.718750000000)
    pts(2) = Array(-172.436689309152, 3.437500000000)
    pts(3) = Array(-172.427550852590, 5.156250000000)
    pts(4) = Array(-172.414756839794, 6.875000000000)
    pts(5) = Array(-172.398307097143, 8.593750000000)
    pts(6) = Array(-172.378201401385, 10.312500000000)
    pts(7) = Array(-172.354439479629, 12.031250000000)
    pts(8) = Array(-172.327021009324, 13.750000000000)
    pts(9) = Array(-172.295945618234, 15.468750000000)
    pts(10) = Array(-172.261212884418, 17.187500000000)
    pts(11) = Array(-172.222822336199, 18.906250000000)
    pts(12) = Array(-172.180773452128, 20.625000000000)
    pts(13) = Array(-172.135065660957, 22.343750000000)
    pts(14) = Array(-172.085698341592, 24.062500000000)
    pts(15) = Array(-172.032670823055, 25.781250000000)
    pts(16) = Array(-171.975982384440, 27.500000000000)
    pts(17) = Array(-171.915632254857, 29.218750000000)
    pts(18) = Array(-171.851619613387, 30.937500000000)
    pts(19) = Array(-171.783943589024, 32.656250000000)
    pts(20) = Array(-171.712603260611, 34.375000000000)
    pts(21) = Array(-171.637597656784, 36.093750000000)
    pts(22) = Array(-171.558925755901, 37.812500000000)
    pts(23) = Array(-171.476586485976, 39.531250000000)
    pts(24) = Array(-171.390578724600, 41.250000000000)
    pts(25) = Array(-171.300901298870, 42.968750000000)
    pts(26) = Array(-171.207552985308, 44.687500000000)
    pts(27) = Array(-171.110532509773, 46.406250000000)
    pts(28) = Array(-171.009838547378, 48.125000000000)
    pts(29) = Array(-170.905469722397, 49.843750000000)
    pts(30) = Array(-170.797424608173, 51.562500000000)
    pts(31) = Array(-170.685701727017, 53.281250000000)
    pts(32) = Array(-170.570299550106, 55.000000000000)
    pts(33) = Array(-165.890516172122, 55.000000000000)
    pts(34) = Array(-165.685028320155, 54.312500000000)
    pts(35) = Array(-165.462237967671, 53.556250000000)
    pts(36) = Array(-165.242836471219, 52.800000000000)
    pts(37) = Array(-165.026811683380, 52.043750000000)
    pts(38) = Array(-164.833345692974, 51.356250000000)
    pts(39) = Array(-164.623734495673, 50.600000000000)
    pts(40) = Array(-164.417466041229, 49.843750000000)
    pts(41) = Array(-164.214529114359, 49.087500000000)
    pts(42) = Array(-164.032922745902, 48.400000000000)
    pts(43) = Array(-163.836315694511, 47.643750000000)
    pts(44) = Array(-163.643008840834, 46.887500000000)
    pts(45) = Array(-163.452991850783, 46.131250000000)
    pts(46) = Array(-163.283095456259, 45.443750000000)
    pts(47) = Array(-163.099331204621, 44.687500000000)
    pts(48) = Array(-162.918827982233, 43.931250000000)
    pts(49) = Array(-162.741576290496, 43.175000000000)
    pts(50) = Array(-162.598964174166, 42.556250000000)
    pts(51) = Array(-162.443045759486, 41.868750000000)
    pts(52) = Array(-162.304999374819, 41.250000000000)
    pts(53) = Array(-162.154142003009, 40.562500000000)
    pts(54) = Array(-162.020639969973, 39.943750000000)
    pts(55) = Array(-161.874820349228, 39.256250000000)
    pts(56) = Array(-161.731643130687, 38.568750000000)
    pts(57) = Array(-161.591102246800, 37.881250000000)
    pts(58) = Array(-161.466864607200, 37.262500000000)
    pts(59) = Array(-161.331316506780, 36.575000000000)
    pts(60) = Array(-161.198387874295, 35.887500000000)
    pts(61) = Array(-161.068073140971, 35.200000000000)
    pts(62) = Array(-160.953020262770, 34.581250000000)
    pts(63) = Array(-160.827657041949, 33.893750000000)
    pts(64) = Array(-160.704892272446, 33.206250000000)
    pts(65) = Array(-160.584720865655, 32.518750000000)
    pts(66) = Array(-160.478779815936, 31.900000000000)
    pts(67) = Array(-160.363522209389, 31.212500000000)
    pts(68) = Array(-160.250843885415, 30.525000000000)
    pts(69) = Array(-160.140740218868, 29.837500000000)
    pts(70) = Array(-160.043844522172, 29.218750000000)
    pts(71) = Array(-159.938620378425, 28.531250000000)
    pts(72) = Array(-159.835958133793, 27.843750000000)
    pts(73) = Array(-159.735853611635, 27.156250000000)
    pts(74) = Array(-159.647943033320, 26.537500000000)
    pts(75) = Array(-159.552687080503, 25.850000000000)
    pts(76) = Array(-159.459977372736, 25.162500000000)
    pts(77) = Array(-159.369810168575, 24.475000000000)
    pts(78) = Array(-159.282181836212, 23.787500000000)
    pts(79) = Array(-159.197088852724, 23.100000000000)
    pts(80) = Array(-159.114527803354, 22.412500000000)
    pts(81) = Array(-159.034495380808, 21.725000000000)
    pts(82) = Array(-158.964625529039, 21.106250000000)
    pts(83) = Array(-158.889388767808, 20.418750000000)
    pts(84) = Array(-158.816671643832, 19.731250000000)
    pts(85) = Array(-158.746471262696, 19.043750000000)
    pts(86) = Array(-158.678784834367, 18.356250000000)
    pts(87) = Array(-158.613609672629, 17.668750000000)
    pts(88) = Array(-158.550943194544, 16.981250000000)
    pts(89) = Array(-158.490782919935, 16.293750000000)
    pts(90) = Array(-158.438779509307, 15.675000000000)
    pts(91) = Array(-158.383374554287, 14.987500000000)
    pts(92) = Array(-158.330469186848, 14.300000000000)
    pts(93) = Array(-158.280061323032, 13.612500000000)
    pts(94) = Array(-158.232148979414, 12.925000000000)
    pts(95) = Array(-158.186730272710, 12.237500000000)
    pts(96) = Array(-158.143803419398, 11.550000000000)
    pts(97) = Array(-158.103366735375, 10.862500000000)
    pts(98) = Array(-158.069101502432, 10.243750000000)
    pts(99) = Array(-158.033391854956, 9.556250000000)
    pts(100) = Array(-158.000168052071, 8.868750000000)
    pts(101) = Array(-157.969428794397, 8.181250000000)
    pts(102) = Array(-157.941172880532, 7.493750000000)
    pts(103) = Array(-157.915399206814, 6.806250000000)
    pts(104) = Array(-157.892106767109, 6.118750000000)
    pts(105) = Array(-157.871294652609, 5.431250000000)
    pts(106) = Array(-157.854683754770, 4.812500000000)
    pts(107) = Array(-157.838582102212, 4.125000000000)
    pts(108) = Array(-157.824958688155, 3.437500000000)
    pts(109) = Array(-157.813812981851, 2.750000000000)
    pts(110) = Array(-157.805144549194, 2.062500000000)
    pts(111) = Array(-157.798953052633, 1.375000000000)
    pts(112) = Array(-157.795238251108, 0.687500000000)
    pts(113) = Array(-157.794000000000, 0.000000000000)
    Profile_23 = pts
End Function

Private Sub AddEnclosedDetails(ByVal model As SldWorks.ModelDoc2, ByVal plane As SldWorks.Feature, ByVal mainAxis As SldWorks.SketchSegment, ByVal label As String)
    Select Case label
        Case "01_FrontBarrel"
            RadialFeatures_0 model, plane, mainAxis
        Case "02_MiddleBarrel"
            RadialFeatures_1 model, plane, mainAxis
        Case "03_RearBarrel"
            RadialFeatures_2 model, plane, mainAxis
        Case "04_G3G4_InternalCell"
        Case "05_G5G6_InternalCell"
        Case "06_G8_InternalCell"
            RadialFeatures_3 model, plane, mainAxis
        Case "07_FrontCupSpacer"
        Case "08_G7G8_Spacer"
        Case "09_G8G9_Spacer"
        Case "10_G9G10_CupSpacer"
        Case "11_FrontRetainer"
            CapToolSlots model, plane, -381.296097289273, -378.246097289273, 89.800000000000, 93.850000000000
        Case "12_MiddleCellRetainer"
            CapToolSlots model, plane, -320.124000000000, -317.074000000000, 59.800000000000, 63.850000000000
        Case "13_RearRetainer"
            CapToolSlots model, plane, -162.300531091410, -159.250531091410, 58.800000000000, 62.850000000000
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
    Set cut = model.FeatureManager.FeatureCut3(False, False, False, 0, 0, 0.002, 0.002, False, False, False, False, 0#, 0#, _
        False, False, False, False, False, False, True, False, False, False, 0, 0#, False)
    If cut Is Nothing Then Err.Raise vbObjectError + 612, , "Cap wrench slot cut failed."
    cut.Name = "Two_Wrench_Slots_W4_D3"
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
    zs = Array(-0.316074000000, -0.316074000000, -0.316074000000)
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
            Set portCircle = model.SketchManager.CreateCircleByRadius(p(0), p(1), 0#, 0.002250000000)
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
    buildLog.WriteLine "3 native radial holes created; diameter_mm=4.5."
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
    zs = Array(-0.316074000000, -0.252074000000, -0.252074000000)
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
        For row = 0 To 1
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
    buildLog.WriteLine "6 native radial holes created; diameter_mm=3.3."
End Sub


Private Sub RadialFeatures_2(ByVal model As SldWorks.ModelDoc2, ByVal basePlane As SldWorks.Feature, ByVal mainAxis As SldWorks.SketchSegment)
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
    zs = Array(-0.252074000000, -0.252074000000, -0.252074000000)
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
            plane.Name = "RadialPlane_2_" & CStr(direction * 60)
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
            Set portCircle = model.SketchManager.CreateCircleByRadius(p(0), p(1), 0#, 0.002250000000)
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
            profile.Name = "GluePort_Group2_Profile_" & CStr(direction) & "_" & CStr(row)
            model.ClearSelection2 True
            If model.Extension.SelectByID2(profile.Name, "SKETCH", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then Err.Raise vbObjectError + 510, , "Adhesive cut sketch selection failed."
            currentStage = "Radial holes: cut direction=" & CStr(direction) & "; row=" & CStr(row)
            buildLog.WriteLine currentStage & "; Z_m=" & CStr(zs(row)) & "; local=" & CStr(p(0)) & "," & CStr(p(1)) & "; construction=" & CStr(portCircle.ConstructionGeometry)
            ' One circular contour, through all in ONE direction: only the selected tube wall.
            ' All bodies are in scope; no assembly scope or sheet-metal options.
            Set cut = model.FeatureManager.FeatureCut3(True, False, reverseCut, 1, 0, 0.04, 0#, False, False, False, False, 0#, 0#, _
                False, False, False, False, False, False, True, False, False, False, 0, 0#, False)
            If cut Is Nothing Then Err.Raise vbObjectError + 511, , "Adhesive port cut failed: direction=" & CStr(direction) & "; row=" & CStr(row)
            cut.Name = "RadialGroup2_" & CStr(direction + 1) & "_" & CStr(row + 1)
            model.ClearSelection2 True
            model.EditRebuild3
            buildLog.WriteLine "Created " & cut.Name
        Next row
    Next direction
    buildLog.WriteLine "3 native radial holes created; diameter_mm=4.5."
End Sub


Private Sub RadialFeatures_3(ByVal model As SldWorks.ModelDoc2, ByVal basePlane As SldWorks.Feature, ByVal mainAxis As SldWorks.SketchSegment)
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
    zs = Array(-0.201074000000, -0.201074000000, -0.201074000000)
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
            plane.Name = "RadialPlane_3_" & CStr(direction * 60)
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
            Set portCircle = model.SketchManager.CreateCircleByRadius(p(0), p(1), 0#, 0.001500000000)
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
            profile.Name = "GluePort_Group3_Profile_" & CStr(direction) & "_" & CStr(row)
            model.ClearSelection2 True
            If model.Extension.SelectByID2(profile.Name, "SKETCH", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then Err.Raise vbObjectError + 510, , "Adhesive cut sketch selection failed."
            currentStage = "Radial holes: cut direction=" & CStr(direction) & "; row=" & CStr(row)
            buildLog.WriteLine currentStage & "; Z_m=" & CStr(zs(row)) & "; local=" & CStr(p(0)) & "," & CStr(p(1)) & "; construction=" & CStr(portCircle.ConstructionGeometry)
            ' One circular contour, through all in ONE direction: only the selected tube wall.
            ' All bodies are in scope; no assembly scope or sheet-metal options.
            Set cut = model.FeatureManager.FeatureCut3(True, False, reverseCut, 1, 0, 0.04, 0#, False, False, False, False, 0#, 0#, _
                False, False, False, False, False, False, True, False, False, False, 0, 0#, False)
            If cut Is Nothing Then Err.Raise vbObjectError + 511, , "Adhesive port cut failed: direction=" & CStr(direction) & "; row=" & CStr(row)
            cut.Name = "RadialGroup3_" & CStr(direction + 1) & "_" & CStr(row + 1)
            model.ClearSelection2 True
            model.EditRebuild3
            buildLog.WriteLine "Created " & cut.Name
        Next row
    Next direction
    buildLog.WriteLine "3 native radial holes created; diameter_mm=3.0."
End Sub
