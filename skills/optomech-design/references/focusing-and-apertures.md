# Focusing mechanisms and closely spaced rear groups

## Choose the motion before its hardware

Identify which optical elements move, their coordinate system, the travel and the fixed installation reference. Distinguish whole-unit focus, single-group focus and zoom. A geometric travel law is not a calibrated optical focus curve. Obtain optical configurations before claiming subject distances, maintained image quality or F-number over travel.

Select an architecture for the actual scale and access: a helicoid with a thrust capture can translate a complete optical unit; an axially captured cam ring with followers can translate an internal cell; a lead screw is another motor-driven option. A geared input demonstrates actuation without specifying a motor, bearings or controller. State those omissions where relevant.

Explain the constraints explicitly: what rotates, what translates, what stays fixed, and what transfers force in both directions. Straight guides prevent rotation; they do not create axial travel. Do not mix a rotating/translating helicoid ring with an axially fixed cam ring in either the drawing or the explanation.

For a housing-fixed follower azimuth alpha and cam angle theta, draw the cam path in ring coordinates using beta=alpha−theta, with axial ordinate tied to the follower's actual axial location plus cell travel. Verify sign conventions against assembled native CAD at endpoints and intermediate poses. Define follower diameter, slot clearance, end retention, axial cam capture and an accessible installation sequence. Three followers can introduce overconstraint; nominal clearance is not a bearing/preload solution.

Use `scripts/motion_law.py` to export an explicit geometric linear or smoothstep law. All motion endpoints are inputs, not design defaults. Smoothstep has zero endpoint slope but is not automatically an optimum optical focus law. Check pressure angle, slot thickness, remaining walls, closed-slot insertion, collision with installed parts and full travel. Static screenshots do not prove a mechanism works.

## When a spacer cannot fit

Measure the peripheral sag and space at the intended bearing radii. Consider a wider practical spacer, a removable lens cell, a short sub-barrel or a newly manufactured peripheral rim. Avoid a stack of inaccessible tiny retaining rings simply to eliminate spacers. Do not adopt bonding by default when servicing and assembly access favor mechanical retention.

With authorization to change the mechanical glass blank, peripheral flat annuli can meet directly, retaining the original optical faces and central air gap. The contact must have finite width, a defined axial datum, a centering method and a continuous preload path. Distinguish this new blank from the original STEP: material cannot be added to an existing glass edge. Glass bearing stress, chips, thermal expansion, preload and tolerances need separate qualification.

For touching annuli, check zero nominal solid overlap and zero separation. Compute contact area using the coplanar faces; a solid/solid common operation can return no area for a zero-volume contact. Sample shared cemented interfaces on the same radial grid to avoid false overlap from inconsistent approximations.

## A lens edge is not automatically an opaque stop

A lens boundary may limit a beam, but extending transparent glass does not create an opaque aperture. Keep the original STOP location and radius in the plan. If a lens-surface opaque annulus replaces a separate metal stop, identify the surface, manufacturing process and candidate opening; keep it marked pending optical checks until traced. Moving the limiting surface can change pupils, F-number and field clipping. Preserve the effective optical surface and account for focus travel in the check.

Metal stops need a finite platform and practical relief, not a zero-thickness tip. An integrated barrel stop is the same part and color as its barrel.

## Save and show the actual mechanism

Keep accepted versions and their input hashes, plan, native CAD, macros, section, assembly explanation and verification reports together. Record exact CAD runtime status separately from static script checks.

For native CAD animation, transform each Body Tip from local to global coordinates exactly once. Re-cut rotating solids at every pose against a fixed world section plane; rotating an already cut half-model produces a false section. Hold camera bounds fixed. Only the designated moving group translates. Provide labelled 3D/2D views, an angle/travel readout and manual playback controls. Describe these as geometric motion, not optical or stress simulation. Save GUI colors after final recompute; a headless save can remove GUI state from an existing FCStd.

## Examples, not defaults

- Double Gauss 50 mm F1.8: whole-unit motion, 6 mm helicoid lead, 270° rotation, 4.5 mm travel, captured thrust joint and straight guides.
- F35 F2.8 v0.2: L2-only motion; 80:16 module-0.8 gearing; axially fixed cam ±40°; smoothstep L2 travel ±2 mm. Rear L4/L5 new-blank annuli contact at radii 10.7–11.2 mm; one rear cap closes a spacer stack. The proposed lens-surface opaque annulus remains optically unverified. These dimensions belong only to that demo.

## Primary references

- [Sony geared cam example, US20070091475A1](https://patents.google.com/patent/US20070091475A1/en): an actuation principle, not validation of the F35 dimensions or exact constraints.
- [Ricoh/Pentax lens drive technologies](https://www.ricoh-imaging.co.jp/english/products/star_lens/technology/): examples of motor-driven optical group actuation.
- [Edmund Optics: aperture stop](https://www.edmundoptics.com/knowledge-center/glossary/aperture-stop/): the system's limiting aperture can be a lens boundary.
- [Ansys: realistic edge apertures](https://optics.ansys.com/hc/en-us/articles/42661783548563-Modeling-optics-with-realistic-edge-apertures): distinguish optical apertures and mechanical edge geometry.
