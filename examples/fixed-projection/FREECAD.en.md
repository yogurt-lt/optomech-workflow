# PF-2 v0.4: one barrel with three removable bonded lens cells

Open `FreeCAD_Model.FCStd`, or execute `FreeCAD.FCMacro` inside FreeCAD and select an output directory. This candidate creates a new document; source Zemax/STEP files and the published examples are unchanged.

The model contains 9 mechanical bodies and 10 optical reference bodies. A single blue barrel encloses three gray cells. The v0.3 peripheral lens flanges and the three individual lens retaining rings have been removed. L1 and L2 use the earlier demo outlines; L3 keeps the previously approved Ø44.6 mm rim. Maximum barrel diameter is 65 mm instead of 74.2 mm.

The L1 cell screws directly into the barrel using its external thread, until its rear face seats against the barrel shoulder. No additional master retaining ring is used. The section figures now explicitly label this interface; actual thread teeth are omitted. The L1 glass itself remains bonded to its cell at three locations.

Each lens is seated, optically centered, locally bonded at three locations, cured and rechecked on the bench. L1 and L2 each have three Ø2 mm adhesive access holes. L3 has three front-open windows, 2 mm wide, extending from Z19.30 to Z20.85 mm. Locations are spaced by 120 degrees and offset from the spanner slots. The provisional radial bond gap is 0.08 mm; adhesive solids and cure stresses are not modeled.

Complete the deeper front-side optical stack first. Install the preassembled L3, L2 and L1 cells in that order. Each cell locates against its own metal shoulder in the common barrel, with a nominal external locking thread. Remove in reverse order. The service unit is the lens plus its cell; replacing the glass alone still requires debonding. Nominal threads, guide fits and manufacturing reliefs need production detailing.

The integral stop retains the 1 mm axial land, 45-degree transitions, Ø10.2754 mm opening and nominal center at Z62.817969 mm. Its finite-length optical acceptance has not been ray traced.

FreeCAD 1.1.4 checks passed for native construction, saving/reopening, new-part interference, window dimension editing/recomputation, and sampled lens-to-cell and cell-to-barrel insertion paths. Paths are sampled at 0.5 mm intervals plus near-seat positions; actual threads and tools are not simulated. The inherited rear-spacer contact with Lens9 retains a recorded approximate-reference overlap of 0.010903 mm³; this contact is not marked as passed. See `freecad_report.json` and the detailed Chinese guide.


User-run and reviewed. See `freecad_report.json` for native FreeCAD 1.1.4 checks.
