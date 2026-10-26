# F35 F2.8 — electric internal-focus concept v0.2

Seven elements in five optical groups; only L2 moves. The fixed housing also guides the moving cell. A 16-tooth input pinion drives an axially captured 80-tooth cam ring (module 0.8). Three followers pass through straight housing slots into closed cam slots. The ring rotates; L2 translates without rotating. No motor selection is included.

## Motion and assembly

Cam rotation −40…+40° produces L2 travel −2…+2 mm using u=(theta+40)/80 and dz=−2+4(3u²−2u³). Input rotation is −5 theta. These are geometric demo positions, not calibrated subject distances or an optical focus curve.

1. Bench-assemble L2 and its accessible front retaining cap in the moving cell.
2. Insert that cell into the fixed guide housing; install the shoulder axles and rollers through the straight guides and cam slots. Capture the cam ring axially, then install the input pinion. Check both travel endpoints.
3. From the rear, install G3 (L3/L4), G4 (L5/L6), the rectangular spacer, L7 and one rear retaining cap. Fit L1 and its separate front retainer.

## Rear support and aperture

L3–L6 require newly manufactured mechanical rims, increasing OD from 21 to 22.4 mm. The original optical faces and central air gap remain unchanged. L4/L5 meet on a finite annular flat at Z=47.922517 mm, radii 10.7…11.2 mm. Their original peripheral gap was only about 0.0703 mm. Material cannot simply be added to the existing glass. Rear preload runs cap → L7 → rectangular spacer → G4 → rim contact → G3 → housing shoulder. No main-housing glass bonding is used in this rear stack.

Glass-rim contact does not itself make an opaque aperture. A lens-surface opaque annulus near original STOP surface 8 is an indicated candidate only. Its process, exact placement, F-number and all-field transmission require optical validation. There is no extra knife-edge stop plate squeezed between these groups.

## Evidence and scope

Native FreeCAD checks passed for 24 bodies (17 mechanical + 7 glass), nine motion poses and 35 insertion positions. Coplanar rim contact area is 34.40044 mm², with zero nominal solid overlap. Fourteen optical faces retain the original sampled geometry. These checks do not qualify preload, tolerances or stress. The approved concept is saved as v0.2; this exact SolidWorks revision has static generation checks, with its separate runtime status recorded in solidworks_checks.json.

The animation is generated from native CAD shapes with a stationary section plane. It shows geometry, not ray tracing. Camera interface, motor mount, bearings, control electronics and production thread specifications are outside this demo.

## Primary references

- [Edmund Optics: aperture stop](https://www.edmundoptics.com/knowledge-center/glossary/aperture-stop/) — a lens boundary can limit a beam; a transparent mechanical rim is not an opaque stop.
- [Ansys: realistic edge apertures](https://optics.ansys.com/hc/en-us/articles/42661783548563-Modeling-optics-with-realistic-edge-apertures) — distinguish optical and mechanical apertures.
- [Sony geared cam-ring example, US20070091475A1](https://patents.google.com/patent/US20070091475A1/en) — an example of geared cam actuation; it does not establish this demo's dimensions or exact constraint arrangement.
