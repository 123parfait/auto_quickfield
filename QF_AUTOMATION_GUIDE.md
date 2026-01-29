# QuickField Automation Notes

## 1) Environment

- QuickField version: 6.2 (ActiveField COM automation)
- Python: **32-bit** (required for 32-bit COM)
- pywin32: installed in 32-bit Python
- Launcher: `run_py32.bat` (forces 32-bit Python)

Why 32-bit: QuickField 6.2 COM server is 32-bit, so 64-bit Python cannot attach.

---
## 2) Issues encountered and how we fixed them

### 2.1 Labels vs geometry

- **Block Labels** (materials) are not the same as **Geometry labels**.
- COM movement of geometry requires block-based selection or geometry labels.
- We used block labels and `Blocks.LabeledAs` / `Blocks.InRectangle` to move geometry.

### 2.2 Why some moves failed

- Direct vertex-based movement often failed because `Vertices` is not exposed for some block objects.
- `InRectangle` in this COM version only accepts **PointXY,PointXY** (not 4 numeric args).
- Fix: use `PointXY` signature and selection `Move(0, PointXY)` on blocks.

### 2.3 Overlapping / shared edges moved multiple times

Problem:
- Moving multiple blocks independently can move the same shared edge **more than once**.
- Result looks like distortion instead of a rigid translation.

Fix:
- Move the set **as one** using `move-blocks-once`, which:
  - Combines the block bounds into one selection
  - Calls a single `Move(0, PointXY)` on the selection

### 2.4 Coil current not applying (critical)

Symptoms:
- `Loading/LoadingEx` seemed to change, but results stayed the same.
- `label-dump` kept showing old values after a run.

Root cause:
- QuickField COM **does not apply changes unless you reassign `Content` back to the label**.
- `TotalCurrent` is a **boolean flag**, not the numeric current value.

Fix (in code):
- Set `Loading/LoadingEx = amps`
- Set `TotalCurrent = True`
- Then **write back**: `label.Content = content`

Example:

```
.\run_py32.bat src/app.py move-blocks-once --labels "steel mover,magnet centre" --dx 3 --dy 0 --pbm "C:\Users\wangjingjun\Desktop\stage\Décision\PKM2507\test1\AM34.pbm" --model "C:\Users\wangjingjun\Desktop\stage\Décision\PKM2507\test1\Am34r1-L4.mod"
```

---
## 3) Function-level COM snippets (for future edits)

Below are minimal **Python/COM** snippets showing the exact calls used, so you can edit logic later without relying on CLI wrappers.

### 4.1 Connect QuickField + problem/model

```python
import win32com.client

QF = win32com.client.Dispatch("QuickField.Application")
pbm = QF.Problems.Item(1)          # first opened problem
pbm.LoadModel()
mdl = pbm.Model
```

### 4.2 Move a block by label (single)

```python
theBlock = mdl.Shapes.Blocks.LabeledAs("", "", "steel mover").Item(1)
theVector = QF.PointXY(3.0, 0.0)
theBlock.Move(0, theVector)        # 0 = qfShift
```

### 4.3 Move multiple blocks as one rigid set (avoid double-moving shared edges)

```python
# Get bounds for each block, then union into one rectangle.
blk1 = mdl.Shapes.Blocks.LabeledAs("", "", "steel mover").Item(1)
blk2 = mdl.Shapes.Blocks.LabeledAs("", "", "magnet centre").Item(1)

left   = min(blk1.Left,   blk2.Left)
right  = max(blk1.Right,  blk2.Right)
bottom = min(blk1.Bottom, blk2.Bottom)
top    = max(blk1.Top,    blk2.Top)

sel = mdl.Shapes.Blocks.InRectangle(QF.PointXY(left, bottom), QF.PointXY(right, top))
sel.Move(0, QF.PointXY(3.0, 0.0))
```

### 4.4 Build mesh + solve + analyze

```python
mdl.Shapes.BuildMesh(True, False)
mdl.Save()
if not pbm.Solved:
    pbm.SolveProblem()
pbm.AnalyzeResults()
```

### 4.5 Mechanical force (Maxwell force integral)

```python
res = pbm.Result
field = res.GetFieldWindow(1)
contour = field.Contour

contour.AddBlock1("steel mover")
contour.AddBlock1("magnet centre")

force = res.GetIntegral(15, contour).Value   # 15 = qfInt_MaxwellForce
print(force.X, force.Y)
```

### 4.6 Set coil current (bobine)

```python
labels = pbm.Labels(3)         # block labels
lbl = labels.Item(1)           # or search by Name
if lbl.Name.strip().lower() == "bobine":
    lbl.Content.TotalCurrent = 400
```
