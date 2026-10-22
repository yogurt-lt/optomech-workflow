"""Native editable radial access holes for the bonded-seat concept.

FeatureCut3 default direction is opposite the sketch normal. Angled reference
planes are 0/60/120 degrees as planes; cut sense gives three 120-degree rays.
"""
def glue_macro_code(ports,offset=0.):
    if len(ports)!=3 or any(p['count']!=3 or p['diameter_mm']!=ports[0]['diameter_mm'] for p in ports):
        raise ValueError('This port pattern requires three axial rows and three equal holes per row.')
    zs=', '.join(f"{(p['z_mm']+offset)/1000:.12f}" for p in ports)
    code=r'''
Private Sub AddAdhesivePorts(ByVal model As SldWorks.ModelDoc2, ByVal basePlane As SldWorks.Feature, ByVal mainAxis As SldWorks.SketchSegment)
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
    zs = Array(__Z_ROWS__)
    For direction = 0 To 2
        currentStage = "Adhesive ports: plane " & CStr(direction)
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
            plane.Name = "GluePlane_" & CStr(direction * 60)
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
        buildLog.WriteLine "Adhesive direction=" & CStr(direction) & "; normal dot base=" & CStr(dot) & "; reverse=" & CStr(reverseCut)
        For row = 0 To 2
            currentStage = "Adhesive ports: sketch direction=" & CStr(direction) & "; row=" & CStr(row)
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
            Set portCircle = model.SketchManager.CreateCircleByRadius(p(0), p(1), 0#, __PORT_RADIUS_M__)
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
            profile.Name = "GluePort_Profile_" & CStr(direction) & "_" & CStr(row)
            model.ClearSelection2 True
            If model.Extension.SelectByID2(profile.Name, "SKETCH", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then Err.Raise vbObjectError + 510, , "Adhesive cut sketch selection failed."
            currentStage = "Adhesive ports: cut direction=" & CStr(direction) & "; row=" & CStr(row)
            buildLog.WriteLine currentStage & "; Z_m=" & CStr(zs(row)) & "; local=" & CStr(p(0)) & "," & CStr(p(1)) & "; construction=" & CStr(portCircle.ConstructionGeometry)
            ' One circular contour, through all in ONE direction: only the selected tube wall.
            ' All bodies are in scope; no assembly scope or sheet-metal options.
            Set cut = model.FeatureManager.FeatureCut3(True, False, reverseCut, 1, 0, 0.04, 0#, False, False, False, False, 0#, 0#, _
                False, False, False, False, False, False, True, False, False, False, 0, 0#, False)
            If cut Is Nothing Then Err.Raise vbObjectError + 511, , "Adhesive port cut failed: direction=" & CStr(direction) & "; row=" & CStr(row)
            cut.Name = "Adhesive_Access_" & CStr(direction + 1) & "_" & CStr(row + 1)
            model.ClearSelection2 True
            model.EditRebuild3
            buildLog.WriteLine "Created " & cut.Name
        Next row
    Next direction
    buildLog.WriteLine "Nine native adhesive access holes created; diameter_mm=__PORT_DIAMETER_MM__."
End Sub
'''
    return code.replace('__Z_ROWS__',zs).replace('__PORT_RADIUS_M__',f"{ports[0]['diameter_mm']/2000:.12f}").replace('__PORT_DIAMETER_MM__',str(ports[0]['diameter_mm']))
