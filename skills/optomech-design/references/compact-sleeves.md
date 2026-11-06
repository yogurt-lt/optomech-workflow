# Compact sleeves: assembly and CAD lessons

Start with the optical grouping and source configurations, not a copied barrel count. Four cells can contain only three independent motion units; rigid internal gaps should share drive. Distinguish group zoom motion, whole-unit focusing and the fixed rear datum; explicit pickups/solves and STEP configuration can differ from base DISZ.

Prefer a compact candidate: short cells or spacer stacks, a straight-slot inner guide and an outer cam, with a short axial-neighbor focus ring if useful. No guide rods does not mean no load guidance: define spaced cylindrical bearing lands, anti-rotation ears and shaft/roller contacts. A local fixed rear pilot need not become a full fourth sleeve. Treat cell count, stroke and dimensions as case decisions.

Preserve original glass diameters first. Work out stepped bores and rear/front loading before enlarging expensive glass blanks. Small edge bevels, new rim extensions and altered optical zones are distinct operations; label each and retain original inputs. Do not call an enlarged STEP rim an existing lens feature. A tiny negative edge gap can be a source/peripheral modeling issue: investigate it before changing optics.

Calculate spacer ends from contact sag. A short vertex gap can allow a substantial peripheral spacer. Reject zero-thickness hooks and sharp tips: truncate, flatten or use a finite-thickness simple taper. Check minimum end section and centering/installation, not just nominal OD. A thin functional stop foil is not a structural retainer.

Consider a separate blackened aperture disk on a short seat with a removable cap instead of a long tapered nose. Keep the optical stop position explicit and clamp only its outer rim; route lens preload through seats and metal frames. Material, thickness and blackening are design assumptions that need process/thermal review, not universal recipe values.

Check assemblies in their actual order including enclosed glass and already installed rings. Provide outside access to shoulder axles, end spanner slots and thrust parts. Separate cam drive, precision guidance, optical preload and overtravel stops. A dedicated arc slot may be replaceable by an end lug and stop pin; justify the retained solution.

SolidWorks pitfalls: many closed part windows raise resource pressure; save/close parts and insert one at a time. Long multi-profile cam lofts may fail although another kernel succeeds. Adjacent-section cuts are a tested recovery option but can create hundreds of features; check body count, stop on failure and retain diagnostics. They are not automatically the most efficient production implementation. Verify profile order and finite radial penetration.

VBA procedure size matters independently of line length: split large cut dispatchers into small helper procedures and large section lookups into small range dispatchers. The demonstrated repair used24 cut calls/helper and64 lookup cases/helper; these are practical examples, not compiler guarantees. Check procedure names, boundaries, exact call ordering and all section entries. Static checks do not prove VBA compilation.

Local resume macros may reuse verified successful parts from the exact geometry version, checking labels, file identity and final body status. Keep workstation paths out of distributed macros. Preserve a portable fresh-build variant. Separate user runtime confirmation, native solid validity, coaxial envelope checks, full3D contact and optical certification in every case report.

For teaching previews, show zoom and focus as separate inputs with correct nested transforms, a paused/scrub mode and optics-only/mechanics-only modes. Explain changing group spacing versus whole-unit translation. Display source focal targets as targets; interpolate mechanically without claiming refocused imaging.

## Close pairs and synchronous optical groups

In user-facing mechanical drawings, number the final moving modules once. For example, after combining original optical G2/G3, label the three mechanical cells G1/G2/G3 and give the optical-to-mechanical mapping once. Keep stable internal CAD IDs when renaming them would break references or validated caches.

The reviewed projection revision used a short L6 pilot in the shared middle cell, a substantial externally located tapered-bore rear spacer with finite end contacts, shortened main sleeves, and an independently rim-clamped black stop foil. These details matter as much as simplifying barrel count. Read [assembly completeness](assembly-completeness.md) before carrying the architecture to another lens.

Do not make one cell per named optical group by default. Groups with constant separation should first be tested as a single stepped cell. Check the smallest/deepest lens can enter first, followed by a substantial spacer, the larger group and an accessible end retainer. Separate cells only for a stated alignment, insertion or service need.

Where peripheral separation is too small for a centered, seated spacer, compare local flat bearing rims with individually located bonded seats. A positive minimum ring thickness alone does not prove centering, insertion, contact or acceptable aperture. With permission to manufacture revised glass, a modest local rim can be simpler than a fragile spacer. Show exact before/after diameters, flat contact plane and band, edge relief and preload path. Do not add a large common glass flange. Existing optical-zone material changes must be marked for optical review even in a demo.

Direct contact must occur at designed peripheral flats, not by translating the optical vertices until curved faces collide. Peripheral mounting adhesive and cementing an optical air gap are different operations. Retain the prescription air gap. For bonded seats, fixtures or datums establish position before cure; include cure access and service as a module. A source patent or adhesive datasheet supports a principle, not the chosen dimensions or optical safety.

For transparent teaching previews, pair shaded faces with feature edges and view-dependent silhouettes, excluding internal tessellation diagonals. Use line contrast to keep low-opacity sleeves readable, and lighter line weights for minor hardware than for optical rims or main slot boundaries. Transform edges by the same zoom/focus pose as their owning solids. Check canvas alpha compositing before compensating with opacity: an opaque cleared background should remain opaque through transparent draws; incorrect accumulated alpha can wash colors white. Validate the displayed result in the browser with motion and layer controls.
