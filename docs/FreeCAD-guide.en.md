# FreeCAD use and migration v0.7

All six demos have been run and reviewed by the user. Native verification used FreeCAD 1.1.4. Each `examples/<case>/` contains `FreeCAD.FCMacro`, `FreeCAD_Model.FCStd`, bilingual guides and `freecad_report.json`.

Mechanical/glass counts: PF-2 9/10, film 7/5, microscope 7/10, phone 5/6, reinforced SLR v5 13/11, wide angle 8/10.

Install a stable release from the [official FreeCAD download page](https://www.freecad.org/downloads.php?lang=en). In FreeCAD, use Macro → Macros, set the user macro directory to the demo folder, select `FreeCAD.FCMacro` and execute. No external Python installation is needed to run it. Select an output directory; each run saves a new native document and `build_report.json`. Alternatively open the supplied `.FCStd` directly.

Each component is an independent Body with Sketch/Revolution and supported Pocket features. Hide outer barrels to inspect the internals. View → Clipping View, X enabled and offset zero, gives an axial display cut; missing filled cut faces are a display limitation, not an invalid solid.

Explicit radial holes, open adhesive windows and wrench slots are supported. Real thread teeth, screw bodies, adhesive solids/cure and Assembly workbench joints are not generated. Actual reports separate native validity, save/reopen, sampled insertion and approximate optical contact checks. Small sampled-reference contact intersections remain recorded rather than certified as exact surfaces. User review is not manufacturing qualification.

```sh
python skills/optomech-design/scripts/workflow.py macro examples/phone/plan.json --cad freecad --output Phone.FCMacro --include-glass
```

Generate native operations from the same reviewed plan; do not translate VBA syntax or assume STEP preserves feature history. Newly generated or edited macros need target-CAD execution. A visible first barrel can be a partial failed build: inspect report status and expected component counts. Read each case guide for its insertion directions.
