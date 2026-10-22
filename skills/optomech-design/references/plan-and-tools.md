# Plan and standalone tools

Python 3.10+ and the standard library are sufficient. Resolve paths relative to this skill directory; no installed repository, network service or CAD license is needed for these numerical checks. SolidWorks is required to execute the generated macro.

```sh
python scripts/workflow.py analyze --step lens.STP --zmx lens.zmx
python scripts/workflow.py check design_plan.json --report checks.json
python scripts/workflow.py macro design_plan.json --output Optomech_Concept.bas --include-glass
```

`analyze` accepts the demonstrated surface-numbered STEP export and optional sequential Zemax text. It identifies group candidates, not actual trimmed surfaces or a manufacturing plan. `check` adds inner/outer meridional chain distances when each mechanical part records `outer_chain_vertex_count`. This detects counterbore and pilot roots, but a ring's axial thickness and local holes/slots need separate checks. It also uses the engine's overlap samples. Loading and preload paths still require explicit design/checking; the tool does not infer them from prose.

Minimal plan shape:

```json
{
  "title": "Reviewed concept",
  "source_origin_z_mm": 0,
  "parts": [{
    "name": "01_Barrel",
    "profile_zr_mm": [[0,10],[20,10],[20,7],[0,7]],
    "outer_chain_vertex_count": 2,
    "role": "Describe seats, insertion and support",
    "color_rgb": [0.18,0.31,0.55]
  }],
  "lenses": [],
  "assembly_scheme": "Describe the selected scheme",
  "assembly_order": {"front": ["First component", "Next component"]},
  "warnings": []
}
```

Profiles are finite `[Z_mm, radius_mm]` points around a closed, simple positive-radius polygon, without repeating its first point. Order the outer chain and then the reversed inner chain when using the wall-distance check. Optical polygons may end on the axis. Lens objects use `id`, `group_id`, `radius_mm`, `front`/`rear` `[radius,Z]` samples and `polygon_zr_mm`. Each part name is a unique ASCII identifier.

Use `reference_lens_ids` to explicitly include optical reference bodies. `blank_revisions` records new rim geometry by lens ID and reason; it is not itself a geometry modifier. Describe STOP and cemented groups separately. Do not derive the exact optical geometry by extending arbitrary untrimmed STEP surfaces or by bounding boxes.

`macro` preserves explicit part colors. It emits the native revolved baseline and the generator's supported glue-port option. It reports per-case extra cut metadata that it does **not** implement; do not claim those features exist until added and checked in the VBA/CAD output. The emitted macro is ASCII, saves beside the `.swp` under a new timestamped folder, and should be run from its `main` entrypoint.
