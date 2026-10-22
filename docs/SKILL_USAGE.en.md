# Optomechanical design Skill

[中文](SKILL_USAGE.md)

Download `optomech-design-skill.zip` and extract the `optomech-design/` folder. Add the whole folder to the skill directory of an AI assistant that supports `SKILL.md`, then load or invoke `optomech-design` using that tool's normal mechanism. Keep its `references/`, `scripts/` and `assets/` directories. If automatic discovery is unavailable, ask the assistant to read `SKILL.md` and follow it directly.

For a first task, provide authorized Zemax/STEP files, units and axis direction, the SolidWorks version, whether mechanical lens rims may be revised, whether aperture checks are required, and the intended manufacturing process.

Example prompt:

> Use optomech-design. Check that the Zemax and STEP files in this folder agree, then consult assembly sources appropriate to this lens type. Explain the barrel splits, insertion directions, seats, retention and load paths before generating editable barrel/spacer/retainer macros and a matching section. Do not choose an architecture from diameter ordering alone. State which checks actually ran. Respond in English.

The bundle includes lessons from six reviewed demos, SolidWorks 2023 implementation guidance, and standalone Python tools for file analysis, profile checks and baseline macro generation. Python 3.10+ is sufficient for those tools. Executing the CAD macro still requires SolidWorks.

```sh
python scripts/workflow.py analyze --step lens.STP --zmx lens.zmx
python scripts/workflow.py check design_plan.json --report checks.json
python scripts/workflow.py macro design_plan.json --output Concept.bas --include-glass
```

The tools check an explicit plan and generate revolved bodies. Dedicated holes/slots must be implemented and verified separately; assembly design remains a reasoning task. See `references/plan-and-tools.md`. Track numerical checks, CAD execution and manufacturing qualification separately. Dimensions from one demo are not universal process limits.

The standalone Skill contains reusable code and distilled guidance, without private conversations, company optics or credentials. The full project package provides the six public examples, but the standalone Skill does not depend on them.
