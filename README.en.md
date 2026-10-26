# Optomech Design Skill

[中文](README.md)

**An open-source Skill for AI assistants: reason through lens assembly from Zemax / STEP inputs, then generate editable SolidWorks and FreeCAD mechanical concepts.**

**[Download the Skill ZIP](https://yogurt-lt.github.io/optomech-workflow/downloads/optomech-design-skill.zip)** · [Setup & usage](docs/SKILL_USAGE.en.md) · [Live demos & 3D focus animation](https://yogurt-lt.github.io/optomech-workflow/)

Extract the ZIP and add the complete `optomech-design/` folder to an assistant that supports `SKILL.md`, or ask the assistant to read that entry file directly. Keep its scripts, references and templates together. The standalone Skill does not require the demo inputs; executing CAD macros requires the chosen CAD application.

Try this prompt:

> Use optomech-design to inspect my Zemax and STEP files. Research the assembly approach, explain barrel segmentation, loading direction, seats, retention and load paths, then generate editable CAD macros and section drawings. Target CAD: FreeCAD or SolidWorks 2023.

The Skill supplies input checks, assembly reasoning, numerical review, CAD adapters and an iterative repair process. Eight demos show the resulting designs, including a 50 mm F1.8 whole-unit manual-focus lens with a native CAD 3D cutaway animation. Manufacturing tolerances, preload and clear aperture require design-specific verification.

## Manual-focus demo

[Double Gauss 50 mm F1.8](examples/double-gauss-focus/DESIGN.en.md): seven elements, five optical groups and 18 mechanical parts; 4.5 mm whole-unit travel driven by a short helicoid ring. CAD execution, native motion and assembly paths have been checked at sampled poses.

## Electric internal-focus demo

[F35 F2.8](examples/f35-electric/DESIGN.en.md): only L2 moves; 17 mechanical parts, geared cam and 4 mm travel. The rear stack uses new peripheral contact rims, a spacer and one cap. [View the 3D animation](https://yogurt-lt.github.io/optomech-workflow/?demo=f35-electric&lang=en). Native motion and insertion checks passed; the proposed lens-surface opaque aperture remains optically unverified.

## Try the local workbench

Use Python 3.10+; no third-party Python packages are required.

```sh
python server.py
```

Open `http://127.0.0.1:8765`. Use **English / 中文** in the header. The choice is remembered locally. It switches controls, case descriptions, assembly explanations, section drawings and design notes. It does not change geometry or CAD macro identifiers.

| Demo | Assembly concept |
|---|---|
| Fixed projection PF-2 | Three removable bonded cells in one common barrel; separate rear spacer stacks |
| Film projection | One-sided loading with a preassembled small first-lens adapter |
| 63X microscope objective | Five short preassembled cells inside a common structural barrel |
| Phone 6P | Three bore stages, local paired rim extensions and rectangular spacers |
| SLR v5 | Three enclosing barrels, selected internal cells and substantial spacer stacks |
| Wide angle | Two barrels, three rectangular spacers and an integral stop |

The main download uses the reviewed macro for the selected case, preserving its dedicated cuts. Import the `.bas` into the SolidWorks VBA editor, save as `.swp`, and run `main`. Output is saved in a new timestamped folder beside the macro. SolidWorks is needed for CAD execution, but not for gallery viewing.

## GitHub Pages

`python tools/build_pages.py` exports a static gallery to `dist/`. The included `.github/workflows/pages.yml` builds and deploys it when pushed to `main`. For the suggested repository `yogurt-lt/optomech-workflow`, enable **Settings → Pages → GitHub Actions**. The expected URL after a successful deployment is `https://yogurt-lt.github.io/optomech-workflow/`. See [deployment instructions](docs/GITHUB_PAGES.md). The public gallery is deployed on GitHub Pages.

The static version retains all six demos, bilingual explanations, 3D views, section images and downloads. File analysis and modified-body generation remain available in the local Python workbench.

Advanced tools let you edit profiles, inspect overlap, save/import plans and analyze another STEP/Zemax system. Edited profiles update the 3D preview; section drawings remain the reviewed baseline. The general modified-body generator does not automatically reproduce all dedicated per-case cuts.

## Reuse the Skill

Download the standalone Skill from the gallery, or use `skills/optomech-design/` in this repository. It contains the assembly reasoning workflow, six-case lessons, SolidWorks implementation guidance and a small standalone Python engine. See [setup and usage](docs/SKILL_USAGE.en.md).

## FreeCAD and CAD portability

The reviewed plan is shared; each backend creates its own native features. FreeCAD creates `Body → Sketch → Revolution/Pocket` history, including explicit holes, open windows and wrench slots, and saves `.FCStd`. It does not translate VBA or recover native history from STEP. See [the migration guide](docs/FreeCAD-guide.en.md).

```sh
python skills/optomech-design/scripts/workflow.py macro examples/wide-angle/plan.json --cad freecad --output WideAngle.FCMacro --include-glass
```

All six demo versions were executed in FreeCAD 1.1.4 and user-reviewed. Incomplete/unsupported features block export unless an explicitly accepted baseline lists its omissions. New or edited macros need their own target-CAD execution. Download per-case macros and native models from the gallery.

## Verification and scope

All six concepts have been run and reviewed by the user in SolidWorks 2023. Numerical geometry checks, CAD execution and production qualification remain separate. This workflow does not claim automatic arbitrary-STEP design, production optical surfaces, ray tracing, qualified tolerances, preload or stress analysis. Known aperture issues are documented per case.

The six fixed-structure examples and the approved manual-focus demo are packaged. The earlier private company design, unrelated CAD outputs, caches and private conversations are excluded. Only approved demo CAD models are included. Sanitized input hashes are recorded in `docs/source_manifest.json`. Code, documentation and provided examples use the included MIT license. Nothing is automatically published to GitHub.

## Six FreeCAD demos run and reviewed

Each demo includes `FreeCAD.FCMacro`, `FreeCAD_Model.FCStd`, bilingual guides and `freecad_report.json`. The gallery exposes separate SolidWorks and FreeCAD downloads. See the [FreeCAD guide](docs/FreeCAD-guide.en.md). Native verification used 1.1.4; new edits require their own checks. PF-2 now uses the accepted three removable bonded cells; its old variant is archived in `examples/fixed-projection/legacy-v02/`.
