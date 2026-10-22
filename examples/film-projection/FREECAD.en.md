# Film projection: FreeCAD v0.2

Open `FreeCAD_Model.FCStd`, or execute `FreeCAD.FCMacro` inside FreeCAD and choose an output directory. The macro creates a new document without overwriting source files.

The model contains seven mechanical bodies and five sampled optical reference bodies. Every mechanical profile vertex matches the accepted SolidWorks design. The previous FreeCAD baseline contained only the mechanics; this candidate explicitly adds all five reference lenses. Colors follow the SolidWorks palette: dark gray barrel, violet L1 adapter, green spacers, orange retainers, and translucent blue glass.

Preassemble L1 in its seat adapter with the small retainer. Load the main barrel from the L1 end in this order: L5, L4–L5 spacer, L4, L3–L4 spacer, L3, L2–L3 spacer, L2, preassembled L1 module, and main retainer. Each body has an editable native sketch and revolution. Bodies share coordinates; Assembly workbench joints are not created.

FreeCAD 1.1.4 checks passed for twelve valid single-solid bodies, native save/reopen, source coordinate offset, mechanical intersections, L1 preassembly, and nine sampled main insertion stages. The original Z offset is -1017.479856504 mm. Barrel OD is 8.8 mm and length is 20.5 mm.

Two sampled-glass contact interfaces retain tiny intersections of approximately 0.00004485 and 0.00004162 mm³, recorded in `freecad_report.json`. These are not certified exact-surface contacts. Original ray-clearance issues and thin spacers remain unchanged; this is a CAD migration of the existing demo. The accepted source design has no wrench slots, physical thread teeth or defined aperture stop; none of these existing features have been silently omitted.

Hide the barrel to inspect all five glass bodies and the L1 adapter. Restore it and use X=0 clipping for an axial section. The macro includes reference glass, so do not overlay duplicate lenses from STEP. Thread envelopes and sampled optical references are not production drawings.


User-run and reviewed. See `freecad_report.json` for native FreeCAD 1.1.4 checks.
