# 63X microscope objective: FreeCAD v0.2

Open `FreeCAD_Model.FCStd`, or execute `FreeCAD.FCMacro` inside FreeCAD and select an output directory. Source optical files and accepted SolidWorks files are not overwritten.

The document contains seven mechanical bodies and ten sampled glass reference bodies. Mechanical sections and glass outlines match the accepted SolidWorks design. The two nose-retainer spanner slots are native sketch/Pocket features, with 1.2 mm tangential width and 0.7 mm axial material removal. The common barrel is dark gray, all five cells violet, the nose retainer orange, and the glass translucent blue.

The prescription runs from infinity toward the specimen: +Z is the specimen/nose end and -Z is the tube-lens side. G1 contains L1–L2; G2 L3–L4; G3 L5–L7; G4 L8–L9; G5 L10.

On the bench, insert each glass group into its own cell from -Z, seat against the +Z shoulder, optically center, locally bond and cure under fixture control, then recheck. Insert the preassembled G1, G2, G3, G4 and G5 cells through the +Z nose opening, then install the nose retainer. Its preload passes through metal cell end faces to the barrel shoulder, rather than through the glass stack.

There are no radial adhesive holes in the original design: local fillets are applied through exposed cell mouths before barrel insertion. The curved S7 optical STOP is retained; no separate flat diaphragm is inserted there. Actual thread teeth and adhesive solids are not modeled. Bodies share coordinates without Assembly workbench joints.

Native FreeCAD verification covers solid validity, save/reopen, five bench insertion paths, five cell-to-barrel insertion paths, retainer insertion, slot position/depth, and editing/recomputing the Pocket depth. Sampled-glass contact intersections, if present, are recorded separately in `freecad_report.json`; they do not certify exact optical contacts. Thread engagement, actual tools, adhesive stress, optical clipping and production fits remain outside this CAD migration.

Hide the barrel to inspect the cells, then restore it and use X=0 clipping. Check cell-to-cell metal contacts, clearance for protruding glass centers, and the orange nose-retainer slots. Do not overlay duplicate STEP glass.


User-run and reviewed. See `freecad_report.json` for native FreeCAD 1.1.4 checks.
