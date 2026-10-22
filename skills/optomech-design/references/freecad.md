# FreeCAD adapter v0.3

## Generate and run

```sh
python scripts/workflow.py macro design_plan.json --cad freecad --output Optomech.FCMacro --include-glass
```

The generator uses ordinary Python; execute the resulting `.FCMacro` inside FreeCAD, not with an external Python interpreter. The Macro menu prompts for an output directory. Each run creates a timestamped folder containing `Optomech_Baseline.FCStd` and `build_report.json`.

Six approved demos were executed in FreeCAD 1.1.4 and reviewed by the user. Their reports belong to those specific versions; newly generated or edited macros need their own native execution. A status on the SolidWorks version does not transfer automatically.

## Modeling and units

Each mechanical or glass-reference component has its own `PartDesign::Body`, `Sketcher::SketchObject` and `PartDesign::Revolution`, with supported cuts added as `PartDesign::Pocket`. An `App::Part` groups the coordinate assembly; it has no Assembly-workbench joints.

Plan points `[Z,r]` in mm become sketch points `(r,Z+source_origin_z_mm,0)`. Rotate the sketch +90 degrees about X and revolve around its `V_Axis`. Do not divide FreeCAD coordinates by 1000; that conversion belongs to the SolidWorks API.

Coincident constraints close the profile; baseline profile sketches remain editable rather than fully dimension-constrained. Glass references are simplified sampled outlines. STEP exchange does not preserve these native sketch/revolve histories.

## Supported cut data

Use explicit ownership, dimensions and angles. Example fields:

```json
{
  "radial_features": [{"part":"01_Barrel","zs_mm":[20],"diameter_mm":2,"angles_degrees":[30,150,270]}],
  "parts": [{
    "name":"02_Cell",
    "profile_zr_mm":[[0,10],[4,10],[4,8],[0,8]],
    "radial_windows":[{"width_mm":2,"z_start_mm":-0.05,"z_end_mm":1.5,"angles_degrees":[30,150,270]}],
    "cap_slot":"front",
    "wrench_slots":{"width_mm":2,"radial_depth_mm":1,"axial_depth_mm":1}
  }]
}
```

- Radial circular cuts begin outside the owner and end on-axis, cutting only one tube wall. Supply all required azimuths explicitly; paired housing holes use the same list.
- `radial_windows` are rectangular radial cuts. Extending an axial bound past a cell mouth creates an open glue window without leaving a thin bridge.
- `cap_slot` is `front` or `rear`; `wrench_slots` creates two opposed end-face notches. Check the remaining radial/axial section, not just the nominal slot dimensions.
- Legacy `glue_ports` lacks explicit per-part ownership and remains unsupported by this adapter. Normalize only after understanding which part receives the glue access. A `cap_slot` without dimensions or radial holes without angles are also incomplete.

Unimplemented/incomplete features are listed by the capability report and block export by default. `--allow-base-only` is available only for an explicitly accepted baseline with documented omissions. Unknown legacy VBA operations still require manual audit: absence from the metadata is not proof of parity. Real threads, screw solids, adhesive material/cure and assembly joints are not generated.

The general SolidWorks generator has a different capability boundary: use the reviewed case macros for complete dedicated cuts, or implement the corresponding VBA operations. FreeCAD support does not silently add them to that backend.

## Demonstrated runtime pitfalls

- **Exact bounds:** a displayed shape's `BoundBox` can follow its tessellation and underestimate circular extrema. Use `shape.optimalBoundingBox(False, False)` for analytic geometric checks. The initial adapter stopped after the first barrel because it compared tessellated extrema to exact radii.
- **Visibility and color:** after all features are recomputed, show the container, each Body and its final Tip; hide old sketches/revolutions/pockets. Apply color and per-face diffuse colors to Body and Tip. An all-gray first component is not proof that the rest was built.
- **Diagnostics:** inspect `build_report.json` status, part count, failure stage and the actual saved file. A partial model can remain visible after a failure. Do not hide an exception and report success.
- **Cuts:** require a valid positive-volume single solid and actual material removal for every Pocket. Check hole direction and the opposite wall. Probe slot depth and the uncut ring section, then save/reopen.

For a new installation, first run a sleeve of known diameter/length/volume, then one full demo. Validate source offsets, part counts, colors, native feature history and save/reopen. Native Boolean contact checks use the sampled reference glass; keep small contact intersections distinct from solid mechanical interference and from exact optical-surface qualification. Sampled insertion paths do not simulate screw engagement or actual tools.

Official references: [macro execution](https://github.com/FreeCAD/FreeCAD-documentation/blob/main/wiki/How_to_install_macros.md), [Pocket](https://github.com/FreeCAD/FreeCAD-documentation/blob/main/wiki/PartDesign_Pocket.md), [clipping view](https://blog.freecad.org/2026/05/28/tutorial-using-the-clipping-view-function/).
