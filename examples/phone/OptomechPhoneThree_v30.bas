Attribute VB_Name = "OptomechPhoneThree_v30"
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
    Dim loaded(10) As SldWorks.ModelDoc2
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
    names = Array("01_SteppedPhoneBarrel", "02_Spacer_L1_L2", "03_Spacer_L2_L3", "04_Spacer_L3_L4", "05_BondedRearRetainer", "Lens1_RevisedBlank_Reference", "Lens2_Optical_Reference", "Lens3_RevisedBlank_Reference", "Lens4_RevisedBlank_Reference", "Lens5_RevisedBlank_Reference", "Lens6_Optical_Reference")
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
    a = SketchPoint(math, mapping, -2.047142607112, 0#)
    b = SketchPoint(math, mapping, 4.911186840894, 0#)
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
        Case Else: Err.Raise vbObjectError + 418, , "Unknown profile index."
    End Select
End Function

Private Function Profile_0() As Variant
    Dim pts(15) As Variant
    pts(0) = Array(-1.047142607112, 1.995149444839)
    pts(1) = Array(0.404068664528, 1.995149444839)
    pts(2) = Array(0.404068664528, 2.570949831590)
    pts(3) = Array(1.662287461361, 2.570949831590)
    pts(4) = Array(1.662287461361, 3.049853976933)
    pts(5) = Array(3.911186840894, 3.049853976933)
    pts(6) = Array(3.911186840894, 2.549853976933)
    pts(7) = Array(1.962287461361, 2.549853976933)
    pts(8) = Array(1.962287461361, 2.070949831590)
    pts(9) = Array(0.704068664528, 2.070949831590)
    pts(10) = Array(0.704068664528, 1.495149444839)
    pts(11) = Array(-0.546568733320, 1.495149444839)
    pts(12) = Array(-0.546568733320, 1.305149444839)
    pts(13) = Array(-0.747142607112, 1.305149444839)
    pts(14) = Array(-0.747142607112, 1.000000000000)
    pts(15) = Array(-1.047142607112, 1.000000000000)
    Profile_0 = pts
End Function

Private Function Profile_1() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(-0.291591378405, 1.485149444839)
    pts(1) = Array(0.037873685489, 1.485149444839)
    pts(2) = Array(0.037873685489, 1.175149444839)
    pts(3) = Array(-0.291591378405, 1.175149444839)
    Profile_1 = pts
End Function

Private Function Profile_2() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(0.495096793915, 1.485149444839)
    pts(1) = Array(0.704068664528, 1.485149444839)
    pts(2) = Array(0.704068664528, 1.235149444839)
    pts(3) = Array(0.495096793915, 1.235149444839)
    Profile_2 = pts
End Function

Private Function Profile_3() As Variant
    Dim pts(3) As Variant
    pts(0) = Array(1.150385814017, 2.060949831590)
    pts(1) = Array(1.580664806168, 2.060949831590)
    pts(2) = Array(1.580664806168, 1.750949831590)
    pts(3) = Array(1.150385814017, 1.750949831590)
    Profile_3 = pts
End Function

Private Function Profile_4() As Variant
    Dim pts(4) As Variant
    pts(0) = Array(3.591186840894, 2.509853976933)
    pts(1) = Array(3.871186840894, 2.509853976933)
    pts(2) = Array(3.891186840894, 2.489853976933)
    pts(3) = Array(3.891186840894, 2.269853976933)
    pts(4) = Array(3.591186840894, 2.269853976933)
    Profile_4 = pts
End Function

Private Function Profile_5() As Variant
    Dim pts(36) As Variant
    pts(0) = Array(-0.708850530635, 0.000000000000)
    pts(1) = Array(-0.706875563981, 0.098853097180)
    pts(2) = Array(-0.700975764039, 0.197706194360)
    pts(3) = Array(-0.691019225223, 0.298356620579)
    pts(4) = Array(-0.676958073764, 0.400804375838)
    pts(5) = Array(-0.661251612158, 0.492468156859)
    pts(6) = Array(-0.642174338401, 0.587726595960)
    pts(7) = Array(-0.622137936160, 0.677593047942)
    pts(8) = Array(-0.580671929240, 0.853731293826)
    pts(9) = Array(-0.566187730984, 0.920232468292)
    pts(10) = Array(-0.554823669788, 0.981341655640)
    pts(11) = Array(-0.547551065207, 1.033464197789)
    pts(12) = Array(-0.544274625651, 1.071208107621)
    pts(13) = Array(-0.543198230337, 1.105357359374)
    pts(14) = Array(-0.544006306410, 1.134114624008)
    pts(15) = Array(-0.546568733320, 1.161868697612)
    pts(16) = Array(-0.546568733320, 1.455149444839)
    pts(17) = Array(-0.337338441758, 1.455149444839)
    pts(18) = Array(-0.337338441758, 1.258130327743)
    pts(19) = Array(-0.319874153587, 1.229373063109)
    pts(20) = Array(-0.304137712941, 1.200615798475)
    pts(21) = Array(-0.290062389996, 1.171858533841)
    pts(22) = Array(-0.276839436322, 1.141303940167)
    pts(23) = Array(-0.265298042484, 1.110749346493)
    pts(24) = Array(-0.255325775496, 1.080194752819)
    pts(25) = Array(-0.246347310456, 1.047842830106)
    pts(26) = Array(-0.238856201213, 1.015490907393)
    pts(27) = Array(-0.228777206599, 0.957976378124)
    pts(28) = Array(-0.221888121244, 0.896867190777)
    pts(29) = Array(-0.217906355978, 0.830366016310)
    pts(30) = Array(-0.216667006888, 0.754878196646)
    pts(31) = Array(-0.217458069607, 0.693769009298)
    pts(32) = Array(-0.219651501340, 0.623673176753)
    pts(33) = Array(-0.232012176096, 0.355871149847)
    pts(34) = Array(-0.236638433210, 0.233652775152)
    pts(35) = Array(-0.239478155558, 0.115029058537)
    pts(36) = Array(-0.240412348006, 0.000000000000)
    Profile_5 = pts
End Function

Private Function Profile_6() As Variant
    Dim pts(27) As Variant
    pts(0) = Array(0.000000000000, 0.000000000000)
    pts(1) = Array(0.000590833821, 0.122648310351)
    pts(2) = Array(0.002334217110, 0.243217835780)
    pts(3) = Array(0.005244729683, 0.363787361210)
    pts(4) = Array(0.009370994733, 0.486435671560)
    pts(5) = Array(0.029950012809, 0.939610784382)
    pts(6) = Array(0.040499981677, 1.243113382877)
    pts(7) = Array(0.043134813583, 1.288846651143)
    pts(8) = Array(0.046454973104, 1.328343564646)
    pts(9) = Array(0.050672639805, 1.363682908306)
    pts(10) = Array(0.055576036851, 1.393894689879)
    pts(11) = Array(0.055576036851, 1.455149444839)
    pts(12) = Array(0.495096793915, 1.455149444839)
    pts(13) = Array(0.475869463889, 1.359525338464)
    pts(14) = Array(0.460057265286, 1.290925436064)
    pts(15) = Array(0.442947777697, 1.224404318586)
    pts(16) = Array(0.388683774201, 1.022762181230)
    pts(17) = Array(0.356428062678, 0.891798731194)
    pts(18) = Array(0.328590046181, 0.762914066080)
    pts(19) = Array(0.304804165962, 0.634029400966)
    pts(20) = Array(0.291916955659, 0.552956789039)
    pts(21) = Array(0.280647875763, 0.471884177112)
    pts(22) = Array(0.271032167318, 0.390811565185)
    pts(23) = Array(0.263300725803, 0.311817738180)
    pts(24) = Array(0.257237900764, 0.232823911174)
    pts(25) = Array(0.252985125528, 0.155908869090)
    pts(26) = Array(0.250352111237, 0.076915042084)
    pts(27) = Array(0.249502305877, 0.000000000000)
    Profile_6 = pts
End Function

Private Function Profile_7() As Variant
    Dim pts(50) As Variant
    pts(0) = Array(0.445601345385, 0.000000000000)
    pts(1) = Array(0.446124299028, 0.050925636720)
    pts(2) = Array(0.447785904928, 0.104166075108)
    pts(3) = Array(0.450576636787, 0.157406513497)
    pts(4) = Array(0.454478970013, 0.210646951886)
    pts(5) = Array(0.459469010890, 0.263887390274)
    pts(6) = Array(0.465517243938, 0.317127828663)
    pts(7) = Array(0.472589455679, 0.370368267051)
    pts(8) = Array(0.481019924548, 0.425923507109)
    pts(9) = Array(0.490478345735, 0.481478747167)
    pts(10) = Array(0.501373201695, 0.539348788894)
    pts(11) = Array(0.526147421734, 0.655088872347)
    pts(12) = Array(0.556408730340, 0.777773360808)
    pts(13) = Array(0.590484780132, 0.900457849269)
    pts(14) = Array(0.629072088641, 1.025457139399)
    pts(15) = Array(0.669458113585, 1.143512024521)
    pts(16) = Array(0.690045471801, 1.199067264579)
    pts(17) = Array(0.710004652890, 1.249992901299)
    pts(18) = Array(0.730214267241, 1.298603736349)
    pts(19) = Array(0.750731015048, 1.344899769731)
    pts(20) = Array(0.770481000753, 1.386566199774)
    pts(21) = Array(0.790488908603, 1.425917828148)
    pts(22) = Array(0.810770274744, 1.462954654853)
    pts(23) = Array(0.831316858167, 1.497676679889)
    pts(24) = Array(0.852087851634, 1.530083903256)
    pts(25) = Array(0.874683367731, 1.562491126623)
    pts(26) = Array(0.897589494235, 1.592583548321)
    pts(27) = Array(0.920657645420, 1.620361168350)
    pts(28) = Array(0.920657645420, 2.030949831590)
    pts(29) = Array(1.150385814017, 2.030949831590)
    pts(30) = Array(1.150385814017, 1.600320065968)
    pts(31) = Array(1.152636755763, 1.574065134969)
    pts(32) = Array(1.158759022893, 1.539343109933)
    pts(33) = Array(1.168685943741, 1.499991481559)
    pts(34) = Array(1.180526722405, 1.460639853184)
    pts(35) = Array(1.213778267325, 1.358788579745)
    pts(36) = Array(1.229565166931, 1.307862943026)
    pts(37) = Array(1.245132433058, 1.252307702968)
    pts(38) = Array(1.258622538270, 1.196752462910)
    pts(39) = Array(1.271180393341, 1.134252817845)
    pts(40) = Array(1.281040271600, 1.071753172780)
    pts(41) = Array(1.288643411573, 1.006938726046)
    pts(42) = Array(1.294179566227, 0.937494675974)
    pts(43) = Array(1.296849163468, 0.884254237585)
    pts(44) = Array(1.298439445752, 0.826384195859)
    pts(45) = Array(1.298882746512, 0.766199352463)
    pts(46) = Array(1.298210623410, 0.699070104060)
    pts(47) = Array(1.294764840316, 0.576385615599)
    pts(48) = Array(1.280896259357, 0.222220960231)
    pts(49) = Array(1.278283178404, 0.108795678446)
    pts(50) = Array(1.277420370897, 0.000000000000)
    Profile_7 = pts
End Function

Private Function Profile_8() As Variant
    Dim pts(66) As Variant
    pts(0) = Array(1.592776390031, 0.000000000000)
    pts(1) = Array(1.593302037800, 0.044526140860)
    pts(2) = Array(1.594875168087, 0.089052281719)
    pts(3) = Array(1.597484409474, 0.133578422579)
    pts(4) = Array(1.601111027742, 0.178104563438)
    pts(5) = Array(1.605729227967, 0.222630704298)
    pts(6) = Array(1.611625931356, 0.269500326255)
    pts(7) = Array(1.626836693084, 0.365583051268)
    pts(8) = Array(1.645830419257, 0.461665776280)
    pts(9) = Array(1.669728788032, 0.564778944586)
    pts(10) = Array(1.695098615402, 0.663205150697)
    pts(11) = Array(1.753653250539, 0.878805411701)
    pts(12) = Array(1.766713780809, 0.930361995854)
    pts(13) = Array(1.777597802804, 0.977231617811)
    pts(14) = Array(1.788381588892, 1.031131683062)
    pts(15) = Array(1.795952314723, 1.080344786118)
    pts(16) = Array(1.800270890717, 1.124870926977)
    pts(17) = Array(1.801508637905, 1.164710105641)
    pts(18) = Array(1.800556117373, 1.192831878815)
    pts(19) = Array(1.798123921655, 1.218610170892)
    pts(20) = Array(1.793989659062, 1.244388462968)
    pts(21) = Array(1.787928071820, 1.270166755045)
    pts(22) = Array(1.780538190111, 1.293601566024)
    pts(23) = Array(1.771151525133, 1.317036377002)
    pts(24) = Array(1.760818373006, 1.338127706883)
    pts(25) = Array(1.748525823811, 1.359219036764)
    pts(26) = Array(1.734096753729, 1.380310366645)
    pts(27) = Array(1.719326884346, 1.399058215428)
    pts(28) = Array(1.700346978765, 1.420149545308)
    pts(29) = Array(1.681218073475, 1.438897394091)
    pts(30) = Array(1.659812226348, 1.457645242874)
    pts(31) = Array(1.635980357960, 1.476393091657)
    pts(32) = Array(1.609571342283, 1.495140940440)
    pts(33) = Array(1.580664806168, 1.513746720214)
    pts(34) = Array(1.580664806168, 2.010949831590)
    pts(35) = Array(1.600664806168, 2.030949831590)
    pts(36) = Array(2.042287461361, 2.030949831590)
    pts(37) = Array(2.062287461361, 2.010949831590)
    pts(38) = Array(2.062287461361, 1.790436768508)
    pts(39) = Array(2.026542179045, 1.640436768508)
    pts(40) = Array(2.021945452264, 1.628719363019)
    pts(41) = Array(2.018731353632, 1.617001957529)
    pts(42) = Array(2.016043726902, 1.598254108747)
    pts(43) = Array(2.015965506092, 1.577162778866)
    pts(44) = Array(2.019065821988, 1.551384486789)
    pts(45) = Array(2.025847973007, 1.520919232517)
    pts(46) = Array(2.034658875595, 1.490453978245)
    pts(47) = Array(2.058827669582, 1.415462583113)
    pts(48) = Array(2.070049740084, 1.377966885547)
    pts(49) = Array(2.080385054471, 1.338127706883)
    pts(50) = Array(2.088746333981, 1.298288528219)
    pts(51) = Array(2.095615015714, 1.253762387360)
    pts(52) = Array(2.100086493680, 1.206892765402)
    pts(53) = Array(2.101955160845, 1.160023143445)
    pts(54) = Array(2.101381773026, 1.108466559292)
    pts(55) = Array(2.099334163304, 1.068627380628)
    pts(56) = Array(2.095672905650, 1.024101239769)
    pts(57) = Array(2.090543948851, 0.977231617811)
    pts(58) = Array(2.083395042643, 0.923331552560)
    pts(59) = Array(2.068049446366, 0.824905346450)
    pts(60) = Array(2.017671080006, 0.527283247020)
    pts(61) = Array(2.003734064243, 0.433544003106)
    pts(62) = Array(1.992592513406, 0.346835202485)
    pts(63) = Array(1.983084814785, 0.255439439668)
    pts(64) = Array(1.976500728901, 0.168730639047)
    pts(65) = Array(1.972590423792, 0.084365319523)
    pts(66) = Array(1.971272356616, 0.000000000000)
    Profile_8 = pts
End Function

Private Function Profile_9() As Variant
    Dim pts(62) As Variant
    pts(0) = Array(2.583324676005, 0.000000000000)
    pts(1) = Array(2.584344685857, 0.103047491579)
    pts(2) = Array(2.587408251870, 0.209039197204)
    pts(3) = Array(2.592306022551, 0.317975116874)
    pts(4) = Array(2.606109217999, 0.574121738800)
    pts(5) = Array(2.608741024971, 0.644782875883)
    pts(6) = Array(2.609828164952, 0.706611370831)
    pts(7) = Array(2.608941438423, 0.780216721959)
    pts(8) = Array(2.605243825516, 0.847933644997)
    pts(9) = Array(2.598688423534, 0.909762139945)
    pts(10) = Array(2.588903609642, 0.968646420847)
    pts(11) = Array(2.579603402496, 1.009865417479)
    pts(12) = Array(2.567982610069, 1.051084414111)
    pts(13) = Array(2.554875266733, 1.089359196697)
    pts(14) = Array(2.539312267684, 1.127633979284)
    pts(15) = Array(2.521081053019, 1.165908761871)
    pts(16) = Array(2.501713670423, 1.201239330412)
    pts(17) = Array(2.479771800763, 1.236569898954)
    pts(18) = Array(2.455153904943, 1.271900467495)
    pts(19) = Array(2.432552770862, 1.301342607947)
    pts(20) = Array(2.405500642893, 1.333728962443)
    pts(21) = Array(2.378961498872, 1.363171102894)
    pts(22) = Array(2.347744449591, 1.395557457391)
    pts(23) = Array(2.286323851495, 1.454441738293)
    pts(24) = Array(2.162588365875, 1.566321872008)
    pts(25) = Array(2.117751480497, 1.610485082685)
    pts(26) = Array(2.099259720081, 1.631094581001)
    pts(27) = Array(2.083038233072, 1.651704079317)
    pts(28) = Array(2.071415428022, 1.669369363588)
    pts(29) = Array(2.062287461361, 1.687232954626)
    pts(30) = Array(2.062287461361, 2.489853976933)
    pts(31) = Array(2.082287461361, 2.509853976933)
    pts(32) = Array(2.670791142131, 2.509853976933)
    pts(33) = Array(2.690791142131, 2.489853976933)
    pts(34) = Array(2.690791142131, 2.240949831590)
    pts(35) = Array(2.538177405900, 2.060949831590)
    pts(36) = Array(2.569777328557, 2.022675049003)
    pts(37) = Array(2.599570267693, 1.984400266416)
    pts(38) = Array(2.627487623551, 1.946125483830)
    pts(39) = Array(2.653517346059, 1.907850701243)
    pts(40) = Array(2.677688971618, 1.869575918656)
    pts(41) = Array(2.700061757206, 1.831301136070)
    pts(42) = Array(2.720715363682, 1.793026353483)
    pts(43) = Array(2.741141743956, 1.751807356851)
    pts(44) = Array(2.759807687648, 1.710588360219)
    pts(45) = Array(2.778001002932, 1.666425149542)
    pts(46) = Array(2.794482214910, 1.622261938865)
    pts(47) = Array(2.809406735936, 1.578098728189)
    pts(48) = Array(2.823777448967, 1.530991303467)
    pts(49) = Array(2.836714854930, 1.483883878744)
    pts(50) = Array(2.849065331644, 1.433832239977)
    pts(51) = Array(2.860761895308, 1.380836387165)
    pts(52) = Array(2.872311484474, 1.321952106262)
    pts(53) = Array(2.883035061845, 1.260123611315)
    pts(54) = Array(2.893391067261, 1.192406688277)
    pts(55) = Array(2.902911980619, 1.121745551194)
    pts(56) = Array(2.912014493869, 1.045195986020)
    pts(57) = Array(2.920676013109, 0.962757992757)
    pts(58) = Array(2.936447750957, 0.783160936004)
    pts(59) = Array(2.950412664747, 0.580010166890)
    pts(60) = Array(2.960705280118, 0.376859397776)
    pts(61) = Array(2.966713711272, 0.185485484843)
    pts(62) = Array(2.968694011872, 0.000000000000)
    Profile_9 = pts
End Function

Private Function Profile_10() As Variant
    Dim pts(63) As Variant
    pts(0) = Array(3.413153570364, 0.000000000000)
    pts(1) = Array(3.412610528756, 0.046611573857)
    pts(2) = Array(3.410813207621, 0.096808653396)
    pts(3) = Array(3.407765225358, 0.147005732935)
    pts(4) = Array(3.403478319682, 0.197202812473)
    pts(5) = Array(3.397969315286, 0.247399892012)
    pts(6) = Array(3.391260362663, 0.297596971551)
    pts(7) = Array(3.383379215465, 0.347794051089)
    pts(8) = Array(3.373672705067, 0.401576636309)
    pts(9) = Array(3.363477497575, 0.451773715848)
    pts(10) = Array(3.351391404991, 0.505556301068)
    pts(11) = Array(3.338166987715, 0.559338886288)
    pts(12) = Array(3.323878566057, 0.613121471508)
    pts(13) = Array(3.290230470608, 0.727857653311)
    pts(14) = Array(3.248180229081, 0.856935857839)
    pts(15) = Array(3.218343682730, 0.942987994191)
    pts(16) = Array(3.183634072674, 1.039796647587)
    pts(17) = Array(3.091936796274, 1.287196539598)
    pts(18) = Array(3.030745413711, 1.444958789577)
    pts(19) = Array(3.005749762200, 1.505912386160)
    pts(20) = Array(2.981096189714, 1.563280477061)
    pts(21) = Array(2.958379048987, 1.613477556600)
    pts(22) = Array(2.936144735393, 1.660089130457)
    pts(23) = Array(2.912678771734, 1.706700704314)
    pts(24) = Array(2.889859373505, 1.749726772490)
    pts(25) = Array(2.851511648839, 1.817851380436)
    pts(26) = Array(2.765875970587, 1.961271607689)
    pts(27) = Array(2.735720839747, 2.015054192909)
    pts(28) = Array(2.723077982533, 2.040152732678)
    pts(29) = Array(2.713340196283, 2.061665766766)
    pts(30) = Array(2.704905315625, 2.083178800854)
    pts(31) = Array(2.699098393294, 2.101106329261)
    pts(32) = Array(2.693255561004, 2.126204869030)
    pts(33) = Array(2.690674455337, 2.151303408800)
    pts(34) = Array(2.690791142131, 2.509853976933)
    pts(35) = Array(3.429673883690, 2.509853976933)
    pts(36) = Array(3.565461303217, 2.309065658778)
    pts(37) = Array(3.600554159544, 2.255283073558)
    pts(38) = Array(3.629696362384, 2.208671499701)
    pts(39) = Array(3.661546391247, 2.154888914481)
    pts(40) = Array(3.691255355831, 2.101106329261)
    pts(41) = Array(3.716902788942, 2.050909249722)
    pts(42) = Array(3.740458109956, 2.000712170184)
    pts(43) = Array(3.769024014094, 1.932587562238)
    pts(44) = Array(3.782490232016, 1.896732505425)
    pts(45) = Array(3.794891598520, 1.860877448612)
    pts(46) = Array(3.816588911684, 1.789167334985)
    pts(47) = Array(3.825941846438, 1.753312278172)
    pts(48) = Array(3.835131248599, 1.713871715677)
    pts(49) = Array(3.849656539562, 1.638576096369)
    pts(50) = Array(3.860972756196, 1.559694971380)
    pts(51) = Array(3.869022184450, 1.477228340709)
    pts(52) = Array(3.873991914314, 1.387590698676)
    pts(53) = Array(3.875491633687, 1.323051596412)
    pts(54) = Array(3.875517508970, 1.258512494148)
    pts(55) = Array(3.874191904432, 1.190387886202)
    pts(56) = Array(3.871400477017, 1.115092266894)
    pts(57) = Array(3.862328884426, 0.957330016916)
    pts(58) = Array(3.834572882193, 0.580851920376)
    pts(59) = Array(3.822918600229, 0.394405624947)
    pts(60) = Array(3.817985196291, 0.290425960188)
    pts(61) = Array(3.814626610486, 0.193617306792)
    pts(62) = Array(3.812568977290, 0.096808653396)
    pts(63) = Array(3.811876040815, 0.000000000000)
    Profile_10 = pts
End Function
