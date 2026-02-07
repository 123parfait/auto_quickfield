from __future__ import annotations
import argparse
import json
from pathlib import Path
from typing import Any, Optional, Sequence

from .connection import (
    win32com,
    dispatch_qf_app,
    open_problem,
    ensure_model_loaded,
    normalize_labels,
    find_shapes_by_label,
    iter_collection,
)

def try_move(target: Any, qf: Any, dx: float, dy: float) -> Optional[str]:
    attempts: Sequence[tuple[str, tuple[Any, ...]]] = []
    try:
        point = qf.PointXY(dx, dy)
        origin = qf.PointXY(0.0, 0.0)
        attempts = [
            ("Move(dx,dy)", (dx, dy)),
            ("Move(PointXY)", (point,)),
            ("Move(PointXY,PointXY)", (origin, point)),
        ]
    except Exception:
        attempts = [
            ("Move(dx,dy)", (dx, dy)),
        ]

    attempts = [("Move()", tuple())] + list(attempts)

    for name, params in attempts:
        try:
            target.Move(*params)
            return name
        except Exception:
            continue
    return None

def try_point_assign(target: Any, qf: Any, dx: float, dy: float) -> Optional[str]:
    try:
        pt = target.Point
    except Exception as exc:
        return f"Point read failed: {exc}"

    try:
        x0 = float(getattr(pt, "X"))
        y0 = float(getattr(pt, "Y"))
    except Exception as exc:
        return f"Point.X/Y read failed: {exc}"

    try:
        new_pt = qf.PointXY(x0 + dx, y0 + dy)
    except Exception as exc:
        return f"PointXY failed: {exc}"

    try:
        target.Point = new_pt
        return "Point=PointXY"
    except Exception:
        pass

    try:
        setattr(pt, "X", x0 + dx)
        setattr(pt, "Y", y0 + dy)
        return "Point.X/Y set"
    except Exception as exc:
        return f"Point set failed: {exc}"

def move_vertices(selection: Any, qf: Any, dx: float, dy: float) -> tuple[int, int]:
    try:
        vertices = selection.Vertices
    except Exception:
        return 0, 0

    moved = 0
    total = 0
    for vtx in iter_collection(vertices):
        total += 1
        if move_vertex(vtx, qf, dx, dy):
            moved += 1
    return moved, total

def _shape_block_label_name(shape: Any) -> str:
    try:
        val = getattr(shape, "BlockLabel")
        if hasattr(val, "Name"):
            return str(val.Name)
        if val is not None:
            return str(val)
    except Exception:
        return ""
    return ""

def move_vertices_by_block_label(model: Any, label: str, qf: Any, dx: float, dy: float) -> tuple[int, int]:
    # Try selection by label first.
    try:
        sel = model.Selection
        sel = sel.LabeledAs(label)
        moved, total = move_vertices(sel, qf, dx, dy)
        if total > 0:
            return moved, total
    except Exception:
        pass

    # Fallback: scan shapes and move their vertices by BlockLabel name.
    moved = 0
    total = 0
    seen: set[tuple[float, float]] = set()
    for shp in iter_all_shapes(model.Shapes):
        if _shape_block_label_name(shp) != label:
            continue
        try:
            vertices = shp.Vertices
        except Exception:
            continue
        for vtx in iter_collection(vertices):
            key = None
            try:
                pt = vtx.Point
                key = (round(float(getattr(pt, "X")), 9), round(float(getattr(pt, "Y")), 9))
            except Exception:
                pass
            if key is not None and key in seen:
                continue
            total += 1
            if move_vertex(vtx, qf, dx, dy):
                moved += 1
                if key is not None:
                    seen.add(key)
    return moved, total

def _selection_in_rectangle(selection: Any, qf: Any, x1: float, y1: float, x2: float, y2: float) -> Optional[Any]:
    # Try common signatures for InRectangle.
    attempts: Sequence[tuple[str, tuple[Any, ...]]] = []
    try:
        p1 = qf.PointXY(x1, y1)
        p2 = qf.PointXY(x2, y2)
        attempts = [
            ("InRectangle(PointXY,PointXY)", (p1, p2)),
            ("InRectangle(x1,y1,x2,y2)", (x1, y1, x2, y2)),
        ]
    except Exception:
        attempts = [("InRectangle(x1,y1,x2,y2)", (x1, y1, x2, y2))]

    for _, params in attempts:
        try:
            return selection.InRectangle(*params)
        except Exception:
            continue
    return None

def _rect_with_epsilon(rect: Sequence[float], epsilon: float) -> Optional[tuple[float, float, float, float]]:
    if len(rect) != 4:
        return None
    x1, y1, x2, y2 = [float(v) for v in rect]
    x_min, x_max = (x1, x2) if x1 <= x2 else (x2, x1)
    y_min, y_max = (y1, y2) if y1 <= y2 else (y2, y1)
    return x_min - epsilon, y_min - epsilon, x_max + epsilon, y_max + epsilon

def _move_vertices_in_collection(col: Any, qf: Any, dx: float, dy: float) -> tuple[int, int]:
    moved = 0
    total = 0
    for vtx in iter_collection(col):
        total += 1
        if move_vertex(vtx, qf, dx, dy):
            moved += 1
    return moved, total

def _move_vertices_in_shapes(shapes: Any, qf: Any, dx: float, dy: float) -> tuple[int, int]:
    moved = 0
    total = 0
    for shp in iter_collection(shapes):
        try:
            vertices = shp.Vertices
        except Exception:
            continue
        moved_i, total_i = _move_vertices_in_collection(vertices, qf, dx, dy)
        moved += moved_i
        total += total_i
    return moved, total

def _try_selection_in_rect(obj: Any, qf: Any, x1: float, y1: float, x2: float, y2: float) -> Optional[Any]:
    try:
        sel = _selection_in_rectangle(obj, qf, x1, y1, x2, y2)
        if sel is not None:
            return sel
    except Exception:
        pass
    return None

def move_vertices_in_rect(model: Any, qf: Any, rect: Sequence[float], dx: float, dy: float, epsilon: float = 0.0) -> tuple[int, int]:
    rect_eps = _rect_with_epsilon(rect, epsilon)
    if rect_eps is None:
        return 0, 0
    x1, y1, x2, y2 = rect_eps

    # Try selection on Shapes first (it knows geometry).
    for selector in (getattr(model, "Shapes", None), getattr(model, "Selection", None)):
        if selector is None:
            continue
        sel = _try_selection_in_rect(selector, qf, x1, y1, x2, y2)
        if sel is None:
            continue
        # If selection supports Move, try that first.
        if try_move(sel, qf, dx, dy) is not None:
            return 1, 1
        try:
            vertices = sel.Vertices
            return _move_vertices_in_collection(vertices, qf, dx, dy)
        except Exception:
            pass
        try:
            blocks = sel.Blocks
            return _move_vertices_in_shapes(blocks, qf, dx, dy)
        except Exception:
            pass

    # Fallback: attempt on model.Shapes.Blocks / Vertices directly.
    try:
        blocks = model.Shapes.Blocks
        sel = _try_selection_in_rect(blocks, qf, x1, y1, x2, y2)
        if sel is not None:
            try:
                vertices = sel.Vertices
                return _move_vertices_in_collection(vertices, qf, dx, dy)
            except Exception:
                return _move_vertices_in_shapes(sel, qf, dx, dy)
    except Exception:
        pass
    return 0, 0

def move_blocks_in_rect(model: Any, qf: Any, rect: Sequence[float], dx: float, dy: float, epsilon: float = 0.0) -> tuple[int, int]:
    rect_eps = _rect_with_epsilon(rect, epsilon)
    if rect_eps is None:
        return 0, 0
    x1, y1, x2, y2 = rect_eps

    def _selection_count(sel: Any) -> Optional[int]:
        try:
            return int(getattr(sel, "Count"))
        except Exception:
            return None

    # Try Blocks.InRectangle to get a shape range, then move it.
    try:
        blocks = model.Shapes.Blocks
        sel = _try_selection_in_rect(blocks, qf, x1, y1, x2, y2)
        if sel is not None:
            count = _selection_count(sel)
            if count == 0:
                sel = None
            else:
                try:
                    vec = qf.PointXY(dx, dy)
                    sel.Move(0, vec)
                    moved = count if count is not None else 1
                    return moved, moved
                except Exception:
                    if try_move(sel, qf, dx, dy) is not None:
                        moved = count if count is not None else 1
                        return moved, moved
    except Exception:
        pass

    # Try Selection.InRectangle then Move.
    try:
        sel = _try_selection_in_rect(model.Selection, qf, x1, y1, x2, y2)
        if sel is not None:
            count = _selection_count(sel)
            if count == 0:
                sel = None
            else:
                try:
                    vec = qf.PointXY(dx, dy)
                    sel.Move(0, vec)
                    moved = count if count is not None else 1
                    return moved, moved
                except Exception:
                    if try_move(sel, qf, dx, dy) is not None:
                        moved = count if count is not None else 1
                        return moved, moved
    except Exception:
        pass

    return 0, 0

def _add_edge(shapes: Any, qf: Any, x1: float, y1: float, x2: float, y2: float, v1: Any = None, v2: Any = None) -> bool:
    attempts: Sequence[tuple[str, tuple[Any, ...]]] = []
    if v1 is not None and v2 is not None:
        attempts.append(("AddEdge(v1,v2)", (v1, v2)))
    try:
        p1 = qf.PointXY(x1, y1)
        p2 = qf.PointXY(x2, y2)
        attempts.append(("AddEdge(PointXY,PointXY)", (p1, p2)))
    except Exception:
        pass
    attempts.append(("AddEdge(x1,y1,x2,y2)", (x1, y1, x2, y2)))

    for _, params in attempts:
        try:
            shapes.AddEdge(*params)
            return True
        except Exception:
            continue
    return False

def _set_selection_label(selection: Any, label: str) -> bool:
    try:
        setattr(selection, "Label", label)
        return True
    except Exception:
        pass
    try:
        selection.Label = label
        return True
    except Exception:
        pass
    try:
        selection.Label(label)
        return True
    except Exception:
        return False

def _try_add_block_label(problem: Any, qf: Any, x: float, y: float, label: str) -> bool:
    # Prefer problem.Labels(3) which is block labels collection.
    candidates: list[Any] = []
    try:
        labels = problem.Labels(3)
        if labels is not None:
            candidates.append(labels)
    except Exception:
        pass

    data_doc = getattr(problem, "DataDoc", None)
    if data_doc is not None:
        try:
            labels = data_doc.Labels(3)
            if labels is not None:
                candidates.append(labels)
        except Exception:
            pass

    point = None
    try:
        point = qf.PointXY(x, y)
    except Exception:
        point = None

    method_names = ("Add", "Insert")
    arg_sets: list[tuple[Any, ...]] = []
    if point is not None:
        arg_sets.append((point,))
        arg_sets.append((point, label))
        arg_sets.append((label, point))
    arg_sets.append((x, y))
    arg_sets.append((x, y, label))
    arg_sets.append((label, x, y))

    for obj in candidates:
        for m in method_names:
            if not hasattr(obj, m):
                continue
            fn = getattr(obj, m)
            for args in arg_sets:
                try:
                    ret = fn(*args)
                    lbl = ret
                    # If Add returns None, try last item.
                    if lbl is None:
                        try:
                            count = int(obj.Count)
                            lbl = obj.Item(count)
                        except Exception:
                            lbl = None
                    if lbl is not None and label:
                        try:
                            setattr(lbl, "Name", label)
                        except Exception:
                            pass
                    return True
                except Exception:
                    continue
    return False

def add_rect_with_block_label(model: Any, qf: Any, problem: Any, rect: Sequence[float], inset: float, label: str) -> tuple[bool, str]:
    if len(rect) != 4:
        return False, "rect must have 4 numbers"
    x1, y1, x2, y2 = [float(v) for v in rect]
    x_min, x_max = (x1, x2) if x1 <= x2 else (x2, x1)
    y_min, y_max = (y1, y2) if y1 <= y2 else (y2, y1)
    x_min += inset
    y_min += inset
    x_max -= inset
    y_max -= inset
    if x_min >= x_max or y_min >= y_max:
        return False, "inset too large; rectangle collapsed"

    shapes = model.Shapes
    try:
        v1 = shapes.AddVertexXY(x_min, y_min)
        v2 = shapes.AddVertexXY(x_min, y_max)
        v3 = shapes.AddVertexXY(x_max, y_max)
        v4 = shapes.AddVertexXY(x_max, y_min)
    except Exception as exc:
        return False, f"AddVertexXY failed: {exc}"

    edges_ok = True
    edges_ok &= _add_edge(shapes, qf, x_min, y_min, x_min, y_max, v1, v2)
    edges_ok &= _add_edge(shapes, qf, x_min, y_max, x_max, y_max, v2, v3)
    edges_ok &= _add_edge(shapes, qf, x_max, y_max, x_max, y_min, v3, v4)
    edges_ok &= _add_edge(shapes, qf, x_max, y_min, x_min, y_min, v4, v1)
    if not edges_ok:
        return False, "AddEdge failed for one or more edges"

    rebuild_model(qf, problem)

    if label:
        # Try to assign label to a block selection inside the rectangle.
        rect_sel = None
        try:
            rect_sel = _selection_in_rectangle(model.Shapes.Blocks, qf, x_min, y_min, x_max, y_max)
        except Exception:
            rect_sel = None
        if rect_sel is None:
            rect_sel = _selection_in_rectangle(model.Selection, qf, x_min, y_min, x_max, y_max)
        if rect_sel is not None:
            _set_selection_label(rect_sel, label)

        # Also try to drop a block label point at the center.
        cx = (x_min + x_max) / 2.0
        cy = (y_min + y_max) / 2.0
        _try_add_block_label(problem, qf, cx, cy, label)

    return True, f"rect=({x_min},{y_min})-({x_max},{y_max}) label={label}"

def move_vertex(vtx: Any, qf: Any, dx: float, dy: float) -> bool:
    # Try Move with explicit delta first (some versions require it).
    if try_move(vtx, qf, dx, dy) is not None:
        return True

    try:
        pt = vtx.Point
    except Exception:
        pt = None

    if pt is not None:
        try:
            x0 = float(getattr(pt, "X"))
            y0 = float(getattr(pt, "Y"))
            try:
                new_pt = qf.PointXY(x0 + dx, y0 + dy)
                vtx.Point = new_pt
                return True
            except Exception:
                setattr(pt, "X", x0 + dx)
                setattr(pt, "Y", y0 + dy)
                return True
        except Exception:
            pass

    try:
        vtx.Move()
        return True
    except Exception:
        return False

def move_block_labels(problem: Any, qf: Any, names: list[str], dx: float, dy: float, debug: bool = False) -> int:
    moved = 0
    target_names = {n.lower(): n for n in names}
    for lbl in iter_labels_by_type(problem, 3):
        try:
            name = str(getattr(lbl, "Name"))
        except Exception:
            continue
        if name.lower() not in target_names:
            continue
        before = _label_point(lbl)
        moved_name = try_point_assign(lbl, qf, dx, dy)
        if moved_name is None:
            moved_name = try_move(lbl, qf, dx, dy)
        if moved_name is not None:
            moved += 1
            if debug:
                after = _label_point(lbl)
                if before and after:
                    print(f"  label '{name}': ({before[0]:.4f},{before[1]:.4f}) -> ({after[0]:.4f},{after[1]:.4f})")
                else:
                    print(f"  label '{name}': moved (point unavailable)")
    return moved

def _find_block_by_label(model: Any, label: str) -> Optional[Any]:
    # Try Blocks.LabeledAs("", "", label) first (as in official sample).
    try:
        blocks = model.Shapes.Blocks
        for method_name in ("LabeledAs", "GetLabeledAs"):
            if not hasattr(blocks, method_name):
                continue
            method = getattr(blocks, method_name)
            for args in ((label,), ("", "", label), ("", label, "")):
                try:
                    sel = method(*args)
                    if sel is None:
                        continue
                    try:
                        return sel.Item(1)
                    except Exception:
                        return sel(1)
                except Exception:
                    continue
    except Exception:
        pass

    # Fallback: scan blocks and compare Label name.
    try:
        blocks = model.Shapes.Blocks
        for blk in iter_collection(blocks):
            try:
                if str(getattr(blk, "Label")).strip().lower() == label.lower():
                    return blk
            except Exception:
                continue
    except Exception:
        pass
    return None

def cmd_move_block(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = open_problem(qf, args.pbm)
    if problem is None:
        return 2

    model = ensure_model_loaded(problem, Path(args.model) if args.model else None)
    if model is None:
        print("Failed to load model.")
        return 3

    label = args.label.strip()
    if not label:
        print("Missing --label.")
        return 4

    blk = _find_block_by_label(model, label)
    if blk is None:
        print(f"Block not found for label: {label}")
        return 5

    dx = float(args.dx)
    dy = float(args.dy)
    try:
        vec = qf.PointXY(dx, dy)
        try:
            blk.Move(0, vec)
            print(f"Moved block '{label}' by dx={dx}, dy={dy} using Move(0, PointXY).")
            return 0
        except Exception:
            pass
    except Exception:
        vec = None

    moved = try_move(blk, qf, dx, dy)
    if moved is None:
        print(f"Move failed for block '{label}'.")
        return 6
    print(f"Moved block '{label}' by dx={dx}, dy={dy} using {moved}.")
    return 0

def _collect_vertices_for_labels(model: Any, labels: list[str]) -> list[Any]:
    vertices: list[Any] = []
    seen: set[tuple[float, float]] = set()

    def _add_vertex(vtx: Any) -> None:
        key = None
        try:
            pt = vtx.Point
            key = (round(float(getattr(pt, "X")), 9), round(float(getattr(pt, "Y")), 9))
        except Exception:
            key = None
        if key is not None and key in seen:
            return
        if key is not None:
            seen.add(key)
        vertices.append(vtx)

    def _add_point_obj(pt: Any) -> None:
        try:
            key = (round(float(getattr(pt, "X")), 9), round(float(getattr(pt, "Y")), 9))
        except Exception:
            return
        if key in seen:
            return
        seen.add(key)
        # Create a temporary PointXY vertex-like object if possible.
        vertices.append(pt)

    for label in labels:
        blk = _find_block_by_label(model, label)
        if blk is None:
            continue

        # Try block vertices first.
        verts = None
        try:
            verts = blk.Vertices
        except Exception:
            verts = None

        # Fallback: use selection from blocks and take its vertices.
        if verts is None:
            try:
                blocks = model.Shapes.Blocks
                sel = None
                for method_name in ("LabeledAs", "GetLabeledAs"):
                    if hasattr(blocks, method_name):
                        try:
                            sel = getattr(blocks, method_name)("", "", label)
                            break
                        except Exception:
                            sel = None
                if sel is not None:
                    verts = sel.Vertices
            except Exception:
                verts = None

        if verts is not None:
            for vtx in iter_collection(verts):
                _add_vertex(vtx)
            continue

        # Last resort: use edges and move their endpoints.
        try:
            edges = blk.Edges
        except Exception:
            edges = None
        if edges is None:
            continue
        for edge in iter_collection(edges):
            for attr in ("Start", "End"):
                try:
                    pt = getattr(edge, attr)
                except Exception:
                    pt = None
                if pt is not None:
                    _add_point_obj(pt)

    return vertices

def cmd_move_blocks_once(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = open_problem(qf, args.pbm)
    if problem is None:
        return 2

    model = ensure_model_loaded(problem, Path(args.model) if args.model else None)
    if model is None:
        print("Failed to load model.")
        return 3

    labels = [v.strip() for v in args.labels.split(",") if v.strip()]
    if not labels:
        print("Missing --labels.")
        return 4

    dx = float(args.dx)
    dy = float(args.dy)
    verts = _collect_vertices_for_labels(model, labels)
    if not verts:
        # Fallback: compute union bounds and move selection once.
        bounds = []
        for label in labels:
            blk = _find_block_by_label(model, label)
            if blk is None:
                if args.debug:
                    print(f"Debug: block not found for label '{label}'")
                continue
            b = _block_bounds(blk)
            if b is not None:
                bounds.append(b)
                if args.debug:
                    print(f"Debug: bounds for '{label}' = {b}")
            elif args.debug:
                print(f"Debug: bounds unavailable for '{label}'")
        if bounds:
            left = min(b[0] for b in bounds)
            bottom = min(b[1] for b in bounds)
            right = max(b[2] for b in bounds)
            top = max(b[3] for b in bounds)
            if args.debug:
                print(f"Debug: union rect=({left},{bottom})-({right},{top})")

            sel = None
            try:
                sel = _selection_in_rectangle(model.Shapes.Blocks, qf, left, bottom, right, top)
                if args.debug and sel is not None:
                    try:
                        print(f"Debug: Blocks.InRectangle Count={int(sel.Count)}")
                    except Exception:
                        print("Debug: Blocks.InRectangle Count unavailable")
            except Exception as exc:
                if args.debug:
                    print(f"Debug: Blocks.InRectangle failed: {exc}")
                sel = None

            if sel is None:
                try:
                    sel = _selection_in_rectangle(model.Selection, qf, left, bottom, right, top)
                    if args.debug and sel is not None:
                        try:
                            print(f"Debug: Selection.InRectangle Count={int(sel.Count)}")
                        except Exception:
                            print("Debug: Selection.InRectangle Count unavailable")
                except Exception as exc:
                    if args.debug:
                        print(f"Debug: Selection.InRectangle failed: {exc}")
                    sel = None

            if sel is not None:
                try:
                    vec = qf.PointXY(dx, dy)
                    sel.Move(0, vec)
                    print(f"Moved selection in rect ({left},{bottom})-({right},{top}) by dx={dx}, dy={dy}.")
                    return 0
                except Exception as exc:
                    if args.debug:
                        print(f"Debug: Selection.Move failed: {exc}")

        print("No vertices found for the given labels.")
        return 5

    moved = 0
    for vtx in verts:
        # Handle Vertex objects first.
        if hasattr(vtx, "Point"):
            if move_vertex(vtx, qf, dx, dy):
                moved += 1
                continue
        # Fallback: move Point objects by setting X/Y.
        try:
            x0 = float(getattr(vtx, "X"))
            y0 = float(getattr(vtx, "Y"))
            setattr(vtx, "X", x0 + dx)
            setattr(vtx, "Y", y0 + dy)
            moved += 1
        except Exception:
            continue

    if moved == 0:
        print("No vertices moved.")
        return 6

    print(f"Moved {moved}/{len(verts)} unique vertices by dx={dx}, dy={dy}.")
    return 0

def cmd_list_blocks(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = open_problem(qf, args.pbm)
    if problem is None:
        return 2

    model = ensure_model_loaded(problem, Path(args.model) if args.model else None)
    if model is None:
        print("Failed to load model.")
        return 3

    try:
        blocks = model.Shapes.Blocks
    except Exception as exc:
        print(f"Failed to access Shapes.Blocks: {exc}")
        return 4

    names = []
    for blk in iter_collection(blocks):
        name = _block_label(blk)
        if name:
            names.append(name)
    if not names:
        print("No block labels found in model.")
        return 5

    print("Block labels in model:")
    for n in sorted(set(names)):
        print(f"- {n}")
    return 0

def _block_bounds(block: Any) -> Optional[tuple[float, float, float, float]]:
    # Try direct Left/Right/Top/Bottom
    try:
        left = float(getattr(block, "Left"))
        right = float(getattr(block, "Right"))
        top = float(getattr(block, "Top"))
        bottom = float(getattr(block, "Bottom"))
        return left, bottom, right, top
    except Exception:
        pass

    # Try Dimensions object
    try:
        dims = block.Dimensions
        left = float(getattr(dims, "Left"))
        right = float(getattr(dims, "Right"))
        top = float(getattr(dims, "Top"))
        bottom = float(getattr(dims, "Bottom"))
        return left, bottom, right, top
    except Exception:
        pass

    # Fallback: compute from vertices
    try:
        verts = block.Vertices
        xs = []
        ys = []
        for vtx in iter_collection(verts):
            pt = vtx.Point
            xs.append(float(getattr(pt, "X")))
            ys.append(float(getattr(pt, "Y")))
        if xs and ys:
            return min(xs), min(ys), max(xs), max(ys)
    except Exception:
        pass
    return None


def list_block_labels(model: Any) -> list[str]:
    labels: list[str] = []
    try:
        blocks = model.Shapes.Blocks
    except Exception:
        return labels
    for blk in iter_collection(blocks):
        try:
            name = str(getattr(blk, "Label"))
        except Exception:
            name = ""
        if name:
            labels.append(name)
    # de-dup while preserving order
    seen: set[str] = set()
    out: list[str] = []
    for name in labels:
        key = name.strip()
        if not key or key.lower() in seen:
            continue
        seen.add(key.lower())
        out.append(key)
    return out

def cmd_block_bounds(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = open_problem(qf, args.pbm)
    if problem is None:
        return 2

    model = ensure_model_loaded(problem, Path(args.model) if args.model else None)
    if model is None:
        print("Failed to load model.")
        return 3

    labels = [v.strip() for v in args.labels.split(",") if v.strip()]
    if not labels:
        print("Missing --labels.")
        return 4

    for label in labels:
        blk = _find_block_by_label(model, label)
        if blk is None:
            print(f"{label}: block not found")
            continue
        bounds = _block_bounds(blk)
        if bounds is None:
            print(f"{label}: bounds unavailable")
        else:
            left, bottom, right, top = bounds
            print(f"{label}: left={left}, bottom={bottom}, right={right}, top={top}")
    return 0

def save_model(problem: Any, model: Any, save_as: Path) -> bool:
    saved = False
    for obj in (model, problem):
        try:
            if hasattr(obj, "SaveAs"):
                obj.SaveAs(str(save_as))
                saved = True
                break
        except Exception:
            continue
    if not saved:
        try:
            problem.Save()
            saved = True
        except Exception:
            saved = False
    return saved

def cmd_model(args: argparse.Namespace) -> int:
    plan_path = Path(args.plan)
    if not plan_path.exists():
        print(f"Plan not found: {plan_path}")
        return 1

    try:
        # Allow UTF-8 with BOM (common on Windows).
        plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        print(f"Invalid JSON in {plan_path}: {exc}")
        return 1

    use_active = args.use_active or bool(plan.get("use_active_problem", False))
    pbm_path = Path(args.pbm) if args.pbm else None
    if pbm_path is None and not use_active and plan.get("pbm"):
        pbm_path = Path(plan["pbm"])
    model_path = Path(args.model) if args.model else None
    if model_path is None and plan.get("model"):
        model_path = Path(plan["model"])

    save_as = Path(args.save_as) if args.save_as else None
    if save_as is None:
        plan_save = plan.get("save_model_as")
        if isinstance(plan_save, str) and plan_save.strip():
            save_as = Path(plan_save)
    actions = plan.get("actions", [])
    rebuild_each = bool(plan.get("rebuild_each_action", False))
    rebuild_end = bool(plan.get("rebuild_at_end", True))

    if not actions:
        print("No actions in modeling plan.")
        return 1

    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if pbm_path is not None:
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    elif use_active:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
            return 2
    else:
        print("No PBM provided. Use --pbm or set use_active_problem=true in plan.")
        return 1

    model = ensure_model_loaded(problem, model_path)
    if model is None:
        print("Failed to load model.")
        return 2

    if args.dry_run:
        print(f"Plan: {plan_path}")
        print(f"Actions: {len(actions)}")
        return 0

    for idx, action in enumerate(actions, start=1):
        action_type = str(action.get("type", "")).strip()
        dx = float(action.get("dx", 0.0))
        dy = float(action.get("dy", 0.0))
        labels = normalize_labels(action.get("labels") or action.get("label"))

        if action_type in ("move_shape", "move_shapes"):
            if not labels:
                print(f"[{idx}] Missing label for move_shape.")
                return 3
            moved_total = 0
            for name in labels:
                shapes = find_shapes_by_label(model, name)
                if not shapes:
                    print(f"[{idx}] No shapes found for label '{name}'.")
                    return 4
                for shp in shapes:
                    moved = try_move(shp, qf, dx, dy)
                    if moved is None:
                        moved = try_point_assign(shp, qf, dx, dy)
                    if moved is None:
                        print(f"[{idx}] Move failed for '{name}'.")
                        return 5
                    moved_total += 1
            print(f"[{idx}] Moved shapes ({moved_total}) by dx={dx}, dy={dy}.")

        elif action_type == "move_vertices":
            if not labels:
                print(f"[{idx}] Missing labels for move_vertices.")
                return 3
            sel = model.Selection
            for name in labels:
                try:
                    ret = sel.LabeledAs(name)
                    if hasattr(ret, "Vertices"):
                        sel = ret
                except Exception as exc:
                    print(f"[{idx}] Selection.LabeledAs failed for '{name}': {exc}")
                    return 5
            moved, total = move_vertices(sel, qf, dx, dy)
            print(f"[{idx}] Vertices moved: {moved}/{total} (dx={dx}, dy={dy}).")
            if total == 0:
                print(f"[{idx}] No vertices found for labels: {', '.join(labels)}")
                return 6

        elif action_type == "move_block_labels":
            if not labels:
                print(f"[{idx}] Missing labels for move_block_labels.")
                return 3
            debug = bool(action.get("debug", False))
            moved = move_block_labels(problem, qf, labels, dx, dy, debug=debug)
            print(f"[{idx}] Block labels moved: {moved}/{len(labels)} (dx={dx}, dy={dy}).")
            if moved == 0:
                print(f"[{idx}] No block labels found for: {', '.join(labels)}")
                return 6

        elif action_type == "move_vertices_by_block_label":
            if not labels:
                print(f"[{idx}] Missing labels for move_vertices_by_block_label.")
                return 3
            total_all = 0
            moved_all = 0
            for name in labels:
                moved, total = move_vertices_by_block_label(model, name, qf, dx, dy)
                moved_all += moved
                total_all += total
                print(f"[{idx}] Vertices for '{name}': {moved}/{total} (dx={dx}, dy={dy}).")
            if total_all == 0:
                print(f"[{idx}] No vertices found for block labels: {', '.join(labels)}")
                return 6

        elif action_type == "move_vertices_in_rect":
            rect = action.get("rect")
            if not rect:
                print(f"[{idx}] Missing rect for move_vertices_in_rect.")
                return 3
            epsilon = float(action.get("epsilon", 0.0))
            moved, total = move_vertices_in_rect(model, qf, rect, dx, dy, epsilon=epsilon)
            print(f"[{idx}] Vertices in rect moved: {moved}/{total} (dx={dx}, dy={dy}).")
            if total == 0:
                print(f"[{idx}] No vertices found in rect: {rect}")
                return 6

        elif action_type == "move_blocks_in_rect":
            rect = action.get("rect")
            if not rect:
                print(f"[{idx}] Missing rect for move_blocks_in_rect.")
                return 3
            epsilon = float(action.get("epsilon", 0.0))
            moved, total = move_blocks_in_rect(model, qf, rect, dx, dy, epsilon=epsilon)
            print(f"[{idx}] Blocks in rect moved: {moved}/{total} (dx={dx}, dy={dy}).")
            if total == 0:
                print(f"[{idx}] No blocks found in rect: {rect}")
                return 6

        elif action_type == "add_rect_with_block_label":
            rect = action.get("rect")
            label = str(action.get("label", "")).strip()
            inset = float(action.get("inset", 0.0))
            ok, msg = add_rect_with_block_label(model, qf, problem, rect, inset, label)
            if ok:
                print(f"[{idx}] Added rectangle: {msg}")
            else:
                print(f"[{idx}] Add rectangle failed: {msg}")
                return 6

        else:
            print(f"[{idx}] Unsupported action type: {action_type}")
            return 7

        if rebuild_each:
            rebuild_model(qf, problem)

    if rebuild_end:
        rebuild_model(qf, problem)

    if save_as is not None:
        save_as.parent.mkdir(parents=True, exist_ok=True)
        if save_model(problem, model, save_as):
            print(f"Saved model to {save_as}")
        else:
            print("Save failed. Use SaveAs in QuickField to persist the model.")
            return 8

    print("Modeling done.")
    return 0
