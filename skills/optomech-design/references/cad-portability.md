# CAD-independent reasoning, product-specific feature creation

Reuse the reviewed **design plan**, not a textual translation of recorded macros.

```
Zemax / STEP → optical and coordinate checks → assembly decision
            → shared plan JSON → target adapter → native CAD verification
                               ├ SolidWorks: VBA / .bas → .SLDPRT / .SLDASM
                               └ FreeCAD: Python / .FCMacro → .FCStd
```

## What transfers

Group order, central sag, new rim definitions, axis/units/origin, closed mechanical profiles, part IDs/colors, contact locations, insertion sequence and preload path transfer as explicit data. Native tool planes, sketch constraints, cut directions, hole/thread implementations, document saves and assembly constraints require an adapter.

Two products using VBA still expose different object models. For example, Inventor's documentation separates learning VBA from learning its own API. FreeCAD macros are Python scripts, not SolidWorks VBA. Renaming an extension or translating `Dim` to Python cannot preserve feature behavior.

STEP is useful for a solid geometry exchange and comparison, but it does not by itself transfer the original editable sketch/revolve history or assembly mate definitions. When native editability matters, rebuild those operations from the plan and retain the target native document.

## Capability contract

Each adapter must state which operations it implements. Do not silently drop dedicated hole/slot/adhesive-port metadata. If only an agreed baseline is available, mark both the generated file and its report accordingly. Separate these statuses:

1. Shared-plan numerical checks passed.
2. Script emitted and syntax/invariant checks passed.
3. Native target-CAD execution passed, including shape validity and saving.
4. Cross-CAD equivalence checked: axis, bounds, solid count and volume; then cuts, contacts and interfaces.

A mocked API test only checks orchestration, not the CAD kernel. User confirmation in one backend does not validate another. Adding a backend should not modify the source optics or overwrite previously reviewed native macros.

## Existing adapter status

| Backend | Implemented | Validation |
|---|---|---|
| SolidWorks 2023 | Reviewed per-demo macros; general baseline generator | Six example macros were user-run; modified/new outputs need their own run |
| FreeCAD | Separate Bodies, editable sketch/revolve/Pocket features, explicit holes/windows/slots, colors and native save/report | Six demos executed in FreeCAD 1.1.4 and user-reviewed; each new output still needs native checks |
| Other CAD | Same reasoning and plan can be reused | No backend or native-runtime claim until implemented and checked |

Official background:

- [FreeCAD macro documentation](https://github.com/FreeCAD/FreeCAD-documentation/blob/main/wiki/How_to_install_macros.md)
- [FreeCAD native history vs STEP](https://github.com/FreeCAD/FreeCAD-documentation/blob/main/wiki/Manual_Modeling_for_product_design.md)
- [Inventor VBA overview](https://help.autodesk.com/cloudhelp/2026/ENU/Inventor-API/files/VBA_Overview.htm)
