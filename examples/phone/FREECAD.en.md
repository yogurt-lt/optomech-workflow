# Phone 6P, three bore stages: FreeCAD v0.2

Open `FreeCAD_Model.FCStd`, or execute `FreeCAD.FCMacro` inside FreeCAD and choose an output directory. Original input files and accepted SolidWorks models are not overwritten.

This is the accepted three-stage design, not the earlier six-equal-flange concept. It contains five mechanical bodies (one barrel, three rectangular spacers, one bonded tail ring) and six sampled lens reference bodies. Every mechanical profile vertex matches the accepted SolidWorks plan. The approved peripheral revisions to L1/L3/L4/L5 are retained. Maximum barrel OD is approximately 6.0997 mm.

Load from the +Z image-side opening toward -Z: L1, spacer L1/L2, L2, spacer L2/L3, L3, spacer L3/L4, L4, L5, L6, tail ring. Keep the stack seated in a fixture, locally bond the tail ring to the barrel, cure and recheck. Individual glass elements are not directly glued to the barrel. The integral stop is Ø2 mm at local Z0. Original STEP offset -0.747142607111703 mm is preserved.

The original design includes no physical threads, spanner slots or radial adhesive holes; no such existing details have been omitted. The filter is a reserved space, not a generated body. Colors: dark gray barrel, green spacers, orange tail ring, translucent blue glass. All bodies have native editable sketches and revolutions, sharing coordinates without Assembly workbench joints.

FreeCAD 1.1.4 verification passed for eleven valid single-solid bodies, save/reopen, coordinate offset, native mechanical/reference-glass intersection checks, and ten sampled insertion stages. Sampling uses 0.05 mm intervals plus near-final positions at 0.001/0.005/0.01 mm. See `freecad_report.json` for actual values. These checks do not simulate tooling or adhesive curing and do not certify exact optical contacts, ray clearance, molding draft, gates, production tolerances or stress.

Hide the barrel to inspect all six lenses, three spacers and tail ring; restore it and use X=0 clipping to inspect the three bore stages. Do not overlay duplicate STEP glass.


User-run and reviewed. See `freecad_report.json` for native FreeCAD 1.1.4 checks.
