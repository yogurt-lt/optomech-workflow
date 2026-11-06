# 30–39 mm F3.5 projection zoom and whole-unit focus · v0.5

Three removable cells run inside a straight-slot guide sleeve. An outer cam drives zoom; a short rear ring drives whole-unit focus. No guide rods. Prefer simple stepped cells, substantial spacers and accessible end retainers.

|Mechanical module|Bench assembly|Location and preload|
|---|---|---|
|G1 front|Front-load L2/L3 cemented doublet → L1 → retainer|New L2 OD35 peripheral flat bears against L1; stepped pilot bores and rear seat|
|G2 middle shared cell|Front-load L6 → rectangular spacer → L4/L5 doublet → retainer|L6 ID24.04 pilot and rear seat; doublet ID30.3 pilot; spacer clears the metal step by about0.14 mm|
|G3 rear|Rear-load L7 → L8 → tapered-bore spacer → L9 → long spacer → L10 → retainer|New L7 OD28 rim bears against L8; stepped pilots and front seat; independent stop foil carries no glass preload|

These are mechanical module numbers: G2 combines the original optical G2/G3; G3 corresponds to original optical G4. Internal CAD IDs remain stable. L6 pilot engagement is about1.94 mm; diametral clearance0.04 mm is a demo assumption, not a released fit.

Zoom rotates the outer cam0–90° to move three units. The axially retained short focus ring rotates±45° and shifts the complete inner guide/cells/cam assembly±0.6 mm relative to the fixed rear base. Bench-assemble cells; install guide, cam and external followers; finish with rear thrust and focus components. Disengage drive, remove thrust retention and external axles to withdraw a module.

The main sleeves are shortened17.9 mm at the front. The L8–L9 spacer has outer-diameter location, finite end contacts and a tapered bore. Only L2 radius grows0.5 mm and L7 radius1 mm, requiring newly manufactured glass rims and optical review of edge relief. Source prescriptions remain unchanged. No new adhesive joints or filled optical air gaps.

## Running and loading

Recommended SolidWorks macro: Projection_Zoom_v05_Fast.bas. Import into a new VBA macro, save as swp and run main. The user confirmed execution of this revision. SolidWorks macro generation is slow because many native cam cuts are built; please allow time. The optimized macro reduces intermediate display/rebuild overhead; no speed multiplier has been measured. mainReuse can reuse matching prior Fast outputs; retain .reuse sidecars. The baseline Projection_Zoom_v05.bas is also provided.

FreeCAD: open FreeCAD_Model.FCStd and edit MotionControl.ZoomDeg / FocusDeg, or run FreeCAD.FCMacro.

Mobile 3D previews load full native meshes and edge lines and may take time to render. Please wait. The preview retains transparent mechanics, prominent lenses, solid edge lines and optics-only/mechanics-only controls.

## Evidence and lessons

57 valid native single solids,21 coaxial motion samples and662 insertion samples passed. Each lens has a nominal radial pilot; axial seats and preload paths are recorded separately. The previous L6 had axial closure but lacked radial location in an oversized bore. This version adds a pilot. Clear insertion and no overlap do not establish a complete assembly.

Combine synchronous constant-spacing groups first, then restore each small lens pilot. Compare local peripheral bearing flats or serviceable bonded cells before forcing tiny spacers into close gaps. Spacers need location, finite contacts and adequate section; they must not bottom on unrelated metal before glass closes. Use a separately retained blackened stop foil instead of a sharp nose; check guidance, groove-end stock and aperture envelope across travel when shortening sleeves.

Optical clearance, contact stress, preload, tolerances and complete3D cam contact remain open. The source has30/34.5/39 mm focal nodes; intermediate positions use mechanical interpolation.
