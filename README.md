# QuickField Automation (starter)

This project scaffolds an external automation tool that talks to QuickField 6.2 using:
- QLMCall.exe (LabelMover command line) for batch sweeps
- (Optional) COM automation via ActiveField/LabelMover later

## Why external exe/script instead of paste-in?
QuickField does not host a Python plugin inside the UI. The reliable way is an external script or exe
that drives QuickField via COM or QLMCall. This keeps your workflow:
- Users can still open the model in QuickField and edit
- Your tool runs batch sweeps (e.g., -3 to +3 mm)

## Setup
1. Install Python 3.9+ (if not already)
2. Optional: `pip install -r requirements.txt`
3. Edit `config/settings.json` to set QuickField paths
4. Run:
   `python src/app.py probe`

## Sweep mover position (QLMCall)
Prereq (inside QuickField/LabelMover):
- Open the problem and ensure mesh is ready
- Define a variation for mover displacement
  - If displacement "in any direction", QLMCall expects two values: X and Y
  - If displacement "fixed along X", QLMCall expects only X

Example sweep:
`python src/app.py sweep --start -3 --end 3 --step 0.5 --mode any --y 0 --clear-results`

Output:
- `outputs/force_vs_pos.csv`

## Per-position multi-current (CSV table)
If each position needs multiple current cases (e.g., 4 currents), define a CSV with columns
that match your LabelMover variation order. Example `cases.csv`:

```
x_offset,i1,i2,i3,i4
-3,600,0,0,0
-3,0,600,0,0
-3,0,0,600,0
-3,0,0,0,600
-2.5,600,0,0,0
...
```

Run:
`python src/app.py table --table cases.csv --vars x_offset,i1,i2,i3,i4 --clear-results`

If you don't know the LabelMover parameter names, you can ignore header names
and use the CSV column order:
`python src/app.py table --table cases.csv --ignore-header --clear-results`

Output:
- `outputs/table_results.csv`

## Generate current x position cases
Generate a cases table for 8 current values (four magnitudes with +/-) and a position sweep:

```
python src/app.py gen-cases --start -3 --end 3 --step 1 --currents 600,-600,400,-400,300,-300,200,-200
```

Then run:
```
python src/app.py table --table inputs/cases.csv --vars I,x_offset --clear-results
```

## COM probe (no LabelMover)
If you want to avoid LabelMover, we can drive QuickField via ActiveField COM.
First run a probe to confirm COM access:

```
python src/app.py com-probe --pbm "C:/Users/wangjingjun/Desktop/stage/Décision/PKM2507/PKM2507/AM34.pbm" --model "C:/Users/wangjingjun/Desktop/stage/Décision/PKM2507/PKM2507/Am34r1-L2.mod"
```

## ActiveField COM probe (in-app)
This probe connects to the running QuickField instance and lists label names
so we can confirm COM access and exact label names before automation:

```
python src/activefield_probe.py
```

## Auto-modeling (ActiveField COM plan)
Use a JSON plan to move labeled shapes or vertices and optionally save a new model file.

1. Edit `config/modeling.json` (paths + actions).
2. Run:

```
python src/app.py model --plan config/modeling.json
```

Notes:
- Use `--pbm`, `--model`, or `--save-as` to override plan values.
- Use `--use-active` if you already opened the problem in QuickField.
- Labels must be geometry labels (shape labels), not just data labels.
- For testing, you can also move block labels (data labels) with `type: "move_block_labels"`.
- If your model has no geometry labels, try `type: "move_vertices_by_block_label"` to move
  the vertices of shapes that use a given block label.
- If you know the bounding rectangle for a region, use `type: "move_vertices_in_rect"`
  with `"rect": [x1, y1, x2, y2]` to move vertices inside that rectangle.
- You can also try `type: "move_blocks_in_rect"` to move blocks inside a rectangle using
  the selection Move() method.
- If you want to modify the currently opened model, set `save_model_as` to an empty string
  or omit it, and the script will avoid creating a new file.
- To create a new rectangle and assign an existing block label, use
  `type: "add_rect_with_block_label"` with `"rect": [x1, y1, x2, y2]`, optional `"inset"`,
  and `"label": "steel s300"`.

## Print block label coordinates
To print the position of a block label (data label), run:

```
python src/app.py label-pos --name "steel mover"
```

If no problem is open in QuickField, pass a PBM file:

```
python src/app.py label-pos --name "steel mover" --pbm "C:/path/to/problem.pbm"
```

## Mesh + solve + force (block label)
To build mesh, solve, and print force-related results for a block label:

```
python src/app.py solve-force --pbm "C:/path/to/problem.pbm" --label "steel s300" --mesh --remesh
```

If no force values show up, dump all result properties for that block:

```
python src/app.py result-dump --pbm "C:/path/to/problem.pbm" --label "steel s300" --solve
```

## Mechanical force via contour integral
QuickField’s sample uses an integral on a contour. Use:

```
python src/app.py solve-integral --pbm "C:/path/to/problem.pbm" --labels "steel mover" --mesh --remesh
```

`--integral-id 15` is Maxwell force (default). The output prints X/Y if available.

## Move block by label
Move a block (geometry) by its label:

```
python src/app.py move-block --label "steel mover" --dx 2 --dy 0
```

Move multiple blocks together:

```
python src/app.py move-blocks --labels "steel mover,magnet centre" --dx 3 --dy 0
```

Move multiple blocks as one rigid set (dedup shared edges/vertices):

```
python src/app.py move-blocks-once --labels "steel mover,magnet centre" --dx 3 --dy 0
```

With diagnostics:

```
python src/app.py move-blocks-once --labels "steel mover,magnet centre" --dx 3 --dy 0 --debug
```

List all block labels in the current model:

```
python src/app.py list-blocks --pbm "C:/path/to/problem.pbm" --model "C:/path/to/model.mod"
```

Get block bounds (left/bottom/right/top) by label:

```
python src/app.py block-bounds --labels "steel mover,magnet centre" --pbm "C:/path/to/problem.pbm" --model "C:/path/to/model.mod"
```

## Set coil current (Total Ampere-Turns)
Set the TotalCurrent for a block label (e.g., bobine):

```
python src/app.py set-current --label "bobine" --amps 400
```

## Next steps we can add
- Load parameter table (CSV/JSON) for other variations
- Export named results (Fx/Fy) with a header map
