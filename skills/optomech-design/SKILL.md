---
name: optomech-design
description: Plan reviewable lens barrels, spacers, retainers and optical cells from Zemax/STEP inputs, then generate editable CAD concepts through SolidWorks or FreeCAD adapters. Use for optical-to-mechanical assembly design, manual/electric focusing, CAD migration and iterative model repair.
---

# Optical-to-mechanical design

Convert an optical layout into an explained, editable mechanical concept. Work in the user's language. Keep original optical inputs unchanged and distinguish newly manufactured mechanical rims from the supplied glass. This skill supports both English and Chinese requests.

## Select the CAD adapter

Keep assembly reasoning and the reviewed plan independent of CAD. Identify the target product/version before emitting code. Reuse the same coordinates, profiles, component IDs, optical offset and engineering assumptions; generate each application's own feature operations instead of translating VBA syntax.

Read [CAD portability](references/cad-portability.md) when switching software. SolidWorks uses VBA; FreeCAD creates native `Body → Sketch → Revolution/Pocket` trees in an `App::Part` coordinate container. Explicit radial holes, open glue windows and wrench slots are supported by the FreeCAD adapter; assembly joints and real threads are separate capabilities. Unsupported cuts block export by default; use the explicit baseline option only for a clearly labeled concept within the task scope, listing every omission. A language match, STEP transfer or script syntax check is not proof of native feature parity or target-CAD runtime success.

## Establish the model before choosing hardware

- Inventory the intended input files, units, optical axis, vertex order, glass intervals, cemented groups and explicit STOP definition. Match STEP placements to the prescription; do not assume that STEP Z=0 or that the first surface faces the specimen.
- The bundled parser handles sequential MM STANDARD/EVENASPH prescriptions and the demonstrated surface-numbered STEP export. It does not execute Zemax solves, trace rays or import arbitrary trimmed STEP geometry. Verify unsupported inputs with a suitable CAD/optical tool before using them.
- Read [assembly reasoning](references/assembly-reasoning.md) before designing seats or deciding to extend glass rims. Consult primary manufacturer, patent or first-hand assembly sources appropriate to the lens type; explain which principle applies and which dimensions are this design's assumptions.

## Explain and then implement the assembly

Write a short plan with the barrel splits, insertion directions, bench subassemblies, seats, radial centering, fastening and preload path. Explain it before generating geometry; proceed within the user's existing authorization. Do not impose a particular number of barrels, an adhesive process or a rim extension merely because another demo used it.

Prefer substantial, accessible rings and simple turned sections when the space permits. A short vertex air gap does not directly determine spacer length: calculate the actual contact sag. Close edges may justify a revised peripheral bearing rim, a locally bonded seat or a preassembled cell. These are alternatives requiring geometric and manufacturing reasons.

Validate incoming components against the housing **and all already-installed components** at shared radii. Include the glass inside a cell when checking cell insertion. Check tool access and the order in which a closure hides other rings. Distinguish how glass enters a cell from how the completed cell enters the barrel.

Keep the optical stop's aperture and axial definition explicit. Give its metal aperture land finite thickness and a practical relief bevel; do not run a wedge to zero thickness. A stop integral with a barrel is one part and uses that barrel's color. A STOP on a curved optical surface is not automatically a separate flat aperture plate.

For a focusing mechanism or lens-integrated aperture, read [focusing and apertures](references/focusing-and-apertures.md). Define fixed, rotating and translating groups before drawing slots or gears; distinguish a geometric motion law from optical focus calibration.

## Deliver geometry with evidence

Read [plan data and bundled tools](references/plan-and-tools.md) for the input schema and commands. The standalone bundle includes a small Python standard-library engine and separate SolidWorks/FreeCAD templates; it does not require this repository or any user CAD files to be installed elsewhere.

Check self-intersections, mechanical/glass overlap, adjacent glass gaps, continuous seating, centering, insertion, and the preload chain. Check minimum distance between inner and outer meridional wall chains, including axial and diagonal step roots; radial wall scans alone can miss a thin bridge. Numerical geometry checks are not stress or tolerance qualification. Select dimensions by scale and process; demo dimensions are not universal manufacturing limits.

Deliver the plan JSON, editable macro, a readable section, assembly sequence and a check report. The section and macro must come from the same final geometry. If requested, add a revolved 3D illustration with consistent part colors; identify cuts it does not represent. Keep real threads, screw solids and production optical surfaces separate from nominal envelopes and sampled references.

If a nominal thread retains a cell but its teeth are omitted, label the mating parts, locking method and axial shoulder in the section. Do not leave a visibly smooth cylinder unexplained as if its retention were obvious.

## Target CAD and the review loop

Read [SolidWorks 2023 implementation](references/solidworks-2023.md) when emitting or repairing VBA. Preserve working plane selection, save and component insertion code; avoid replacing it with an untested API guess. Use the bundled generator for the revolved baseline. Implement and verify any additional cuts explicitly; the bundled generator does not infer every feature from prose or from arbitrary metadata.

For FreeCAD, read [the FreeCAD adapter](references/freecad.md). Check its capability report and native save/shape/axis results. Never carry the SolidWorks meter conversion into a millimeter FreeCAD sketch. Keep each part a separate Body and retain the editable feature tree in `.FCStd`; STEP is a geometry handoff, not an equivalent feature-history transfer.

Tell the user which checks ran and whether this exact macro was run in CAD. A screenshot of success confirms execution, not production fitness. When reviewing feedback, update the mechanical scheme if necessary, rather than repeatedly shrinking rings to make a section pass. Regenerate the section and macro together and retain the last user-accepted version.

For design choices and known failure modes from the public fixed and focusing demos, read [case lessons](references/case-lessons.md). They are examples, not a catalog of mandatory architectures.

Do not publish input optics, conversation transcripts or a whole working directory as part of using this skill. Package only the specifically approved examples when the user asks for release preparation.
