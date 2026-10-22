# SolidWorks 2023: preserve the demonstrated invariants

Use the bundled engine's `templates/solidworks.bas.in` rather than reconstructing the fragile COM lifecycle from memory. It creates separate native parts and a fixed-position assembly from explicit meridional profiles.

- **Planes:** identify an origin `RefPlane` and its transform/orientation. Do not assume that localized names such as `Front Plane` or `前视基准面` exist or are selectable in the active document. Check selection and entry into a new sketch before drawing.
- **Coordinates:** convert millimeters to meters exactly once. Keep profiles in a documented local Z and restore `source_origin_z_mm` consistently for glass and mechanics. Check SAFEARRAY types and math-point readback when constructing transformed sketch points.
- **Save:** save each new document to its requested full path and read back `GetPathName`. Silent SaveAs option 1 is different from Copy option 2; a copied save can leave the current document unnamed. Use unique run output paths to avoid reuse of stale parts.
- **Insertion:** keep/open saved part documents, activate the assembly for each insertion, and check `AddComponent5` return values. Preserve the template's validated placement method; do not replace it with an unsupported `Transform2` assignment. Errors 438 and component insertion failures occurred in earlier unverified variants.
- **Cuts:** every hole profile must be closed, non-construction and in the intended plane. Confirm cut direction, start plane and result. The project repaired adhesive-port failures by checking those conditions, not by ignoring a failed cut.
- **Tool slots:** size width, radial depth and axial depth for the actual ring; a slot that consumes a thin retainer is not a useful tool interface.
- **Appearance:** apply explicit `color_rgb` to both the part and assembly component. Use a common cell color and distinguish structural barrels; an integral stop does not get a fictitious separate part or color.
- **Diagnostics:** keep staged error messages, output paths and build logs. Run syntax/invariant checks on the emitted source. State that CAD runtime is pending until this version has actually run.
- **VBA identifiers:** avoid statement names such as `Circle` and `Line` for variables; use `portCircle` and `segment`. A declaration using `circle` caused a real syntax failure in the PF-2 detail routine. Procedure-balance and line-length checks are not a substitute for VBA compilation.

The general bundled generator emits revolved bodies and optional built-in adhesive ports. It does not automatically recreate per-case `radial_features` or `cap_slot` additions. Its CLI helper reports that limitation. Implement extra native features with a version-checked API or retain a suitable reviewed case macro as the starting point.
