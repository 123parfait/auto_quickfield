import argparse
import json
from pathlib import Path
import sys
import subprocess
from decimal import Decimal, InvalidOperation
import csv
import re
import time
from typing import Any, Iterable, Optional, Sequence

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except Exception:
    win32com = None
    pythoncom = None

DEFAULT_INSTALL = Path(r"C:\Program Files (x86)\Tera Analysis\QuickField 6.2")
DEFAULT_QLMCALL = DEFAULT_INSTALL / "Tools" / "QLMCall.exe"


def load_settings(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"Invalid JSON in {path}: {exc}")
        return {}


def resolve_paths(settings: dict) -> dict:
    qf = settings.get("quickfield", {})
    install_dir = Path(qf.get("install_dir", str(DEFAULT_INSTALL)))
    qlmcall = Path(qf.get("qlmcall", str(DEFAULT_QLMCALL)))
    return {
        "install_dir": install_dir,
        "qlmcall": qlmcall,
    }


def cmd_probe(args: argparse.Namespace) -> int:
    settings_path = Path(args.config)
    settings = load_settings(settings_path)
    paths = resolve_paths(settings)

    print("QuickField automation probe")
    print(f"- config: {settings_path}")
    print(f"- install_dir: {paths['install_dir']}")
    print(f"- qlmcall: {paths['qlmcall']}")

    ok = True
    if not paths["install_dir"].exists():
        ok = False
        print("  ! install_dir not found")
    if not paths["qlmcall"].exists():
        ok = False
        print("  ! QLMCall.exe not found")

    if ok:
        print("OK: QLMCall.exe found. We can drive batch solves via command line.")
        print("Next: add a run command that calls QLMCall with a template + parameters.")
    else:
        print("Fix paths in config/settings.json, then re-run probe.")

    return 0 if ok else 1


def parse_decimal(value: str) -> Decimal:
    try:
        return Decimal(value)
    except InvalidOperation as exc:
        raise argparse.ArgumentTypeError(f"Invalid number: {value}") from exc


def format_decimal(value: Decimal) -> str:
    # Normalize but keep at least one digit.
    text = format(value.normalize(), "f")
    if text == "-0":
        text = "0"
    return text


def generate_positions(start: Decimal, end: Decimal, step: Decimal) -> list[Decimal]:
    if step == 0:
        raise ValueError("step must be non-zero")
    direction = 1 if step > 0 else -1
    if direction * (end - start) < 0:
        raise ValueError("step sign does not move from start to end")

    positions: list[Decimal] = []
    current = start
    # Use a loop with tolerance on comparison in Decimal space.
    while (current - end) * direction <= 0:
        positions.append(current)
        current += step
    return positions


def run_qlmcall(qlmcall: Path, args_list: list[str]) -> subprocess.CompletedProcess:
    return subprocess.run(
        [str(qlmcall), *args_list],
        text=True,
        capture_output=True,
        check=False,
    )


def parse_results(output: str) -> list[float]:
    tokens = re.split(r"\s+", output.strip())
    results: list[float] = []
    for token in tokens:
        if not token:
            continue
        try:
            results.append(float(token))
        except ValueError:
            continue
    return results


def cmd_sweep(args: argparse.Namespace) -> int:
    settings_path = Path(args.config)
    settings = load_settings(settings_path)
    paths = resolve_paths(settings)

    if not paths["qlmcall"].exists():
        print("QLMCall.exe not found. Fix config/settings.json and re-run.")
        return 1

    start = parse_decimal(args.start)
    end = parse_decimal(args.end)
    step = parse_decimal(args.step)
    y_offset = parse_decimal(args.y)

    try:
        positions = generate_positions(start, end, step)
    except ValueError as exc:
        print(f"Invalid sweep range: {exc}")
        return 1

    fixed_values = [format_decimal(parse_decimal(v)) for v in args.fixed]
    mode = args.mode

    if args.clear_results:
        clear = run_qlmcall(paths["qlmcall"], ["ClearResults"])
        if clear.returncode != 0:
            print("Failed to clear results in LabelMover.")
            print(clear.stderr.strip())
            return clear.returncode

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    rows: list[list[str]] = []
    for pos in positions:
        params: list[str] = []
        params.extend(fixed_values)
        if mode == "any":
            params.append(format_decimal(pos))
            params.append(format_decimal(y_offset))
        elif mode == "x":
            params.append(format_decimal(pos))
        else:
            params.append(format_decimal(y_offset))

        result = run_qlmcall(paths["qlmcall"], params)
        if result.returncode != 0:
            print(f"QLMCall failed at position {pos}:")
            print(result.stderr.strip())
            return result.returncode

        values = parse_results(result.stdout)
        rows.append([format_decimal(pos)] + [str(v) for v in values])

        if args.verbose:
            print(f"{format_decimal(pos)} -> {values}")

    if rows:
        header = ["x_offset"] + [f"result_{i}" for i in range(len(rows[0]) - 1)]
    else:
        header = ["x_offset"]

    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)

    print(f"Wrote {len(rows)} rows to {output_path}")
    return 0


def read_table(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError("CSV has no header")
        rows = [row for row in reader]
    return list(reader.fieldnames), rows


def cmd_table(args: argparse.Namespace) -> int:
    settings_path = Path(args.config)
    settings = load_settings(settings_path)
    paths = resolve_paths(settings)

    if not paths["qlmcall"].exists():
        print("QLMCall.exe not found. Fix config/settings.json and re-run.")
        return 1

    table_path = Path(args.table)
    if not table_path.exists():
        print(f"Table not found: {table_path}")
        return 1

    try:
        header, rows = read_table(table_path)
    except ValueError as exc:
        print(f"Invalid table: {exc}")
        return 1

    if args.ignore_header:
        var_order = header
    else:
        var_order = [v.strip() for v in args.vars.split(",") if v.strip()]
        if not var_order:
            print("No vars specified. Use --vars \"x_offset,i1,i2\".")
            return 1

        missing = [v for v in var_order if v not in header]
        if missing:
            print(f"Table missing columns: {', '.join(missing)}")
            return 1

    if args.clear_results:
        clear = run_qlmcall(paths["qlmcall"], ["ClearResults"])
        if clear.returncode != 0:
            print("Failed to clear results in LabelMover.")
            print(clear.stderr.strip())
            return clear.returncode

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    out_rows: list[list[str]] = []
    for row in rows:
        params: list[str] = []
        for v in var_order:
            params.append(row.get(v, "").strip())

        result = run_qlmcall(paths["qlmcall"], params)
        if result.returncode != 0:
            print(f"QLMCall failed for row {row}:")
            if result.stdout.strip():
                print("stdout:", result.stdout.strip())
            if result.stderr.strip():
                print("stderr:", result.stderr.strip())
            print(f"returncode: {result.returncode}")
            return result.returncode

        values = parse_results(result.stdout)
        out_rows.append([row.get(v, "") for v in var_order] + [str(v) for v in values])

        if args.verbose:
            print(f"{[row.get(v, '') for v in var_order]} -> {values}")

    if out_rows:
        header_out = var_order + [f"result_{i}" for i in range(len(out_rows[0]) - len(var_order))]
    else:
        header_out = var_order

    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header_out)
        writer.writerows(out_rows)

    print(f"Wrote {len(out_rows)} rows to {output_path}")
    return 0


def cmd_gen_cases(args: argparse.Namespace) -> int:
    start = parse_decimal(args.start)
    end = parse_decimal(args.end)
    step = parse_decimal(args.step)
    try:
        positions = generate_positions(start, end, step)
    except ValueError as exc:
        print(f"Invalid sweep range: {exc}")
        return 1

    currents = [c.strip() for c in args.currents.split(",") if c.strip()]
    if len(currents) != 8:
        print("Expected 8 current values (four magnitudes with +/-).")
        print("Example: --currents 600,-600,400,-400,300,-300,200,-200")
        return 1

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    header = [args.current_name, "x_offset"]
    rows: list[list[str]] = []
    for i_val in currents:
        for pos in positions:
            rows.append([i_val, format_decimal(pos)])

    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)

    print(f"Wrote {len(rows)} rows to {output_path}")
    return 0


def com_open_problem(pbm_path: Path):
    if win32com is None:
        raise RuntimeError("pywin32 is not available. Install with pip install pywin32.")
    app = win32com.client.Dispatch("QuickField.Application")
    app.Problems.Open(str(pbm_path))
    return app, app.ActiveProblem


def dispatch_qf_app() -> Any:
    if win32com is None:
        raise RuntimeError("pywin32 is not available. Install with pip install pywin32.")
    try:
        return win32com.client.gencache.EnsureDispatch("QuickField.Application")
    except Exception:
        return win32com.client.Dispatch("QuickField.Application")


def get_active_problem(qf: Any) -> Optional[Any]:
    try:
        prb = qf.ActiveProblem
        if prb is not None:
            return prb
    except Exception:
        pass

    try:
        problems = qf.Problems
        count = int(problems.Count)
        if count > 0:
            try:
                return problems.Item(count)
            except Exception:
                return problems(count)
    except Exception:
        pass
    return None


def ensure_model_loaded(problem: Any, model_path: Optional[Path]) -> Optional[Any]:
    model = None
    if model_path:
        if not model_path.exists():
            raise FileNotFoundError(f"Model not found: {model_path}")
        try:
            problem.App.Models.Open(str(model_path))  # type: ignore[attr-defined]
            time.sleep(0.2)
            model = problem.Model
        except Exception:
            try:
                problem.Models.Open(str(model_path))
                time.sleep(0.2)
                model = problem.Model
            except Exception:
                model = None

    if model is None:
        try:
            problem.LoadModel()
            time.sleep(0.2)
            model = problem.Model
        except Exception:
            model = None
    return model


def iter_collection(col: Any) -> Iterable[Any]:
    try:
        count = int(col.Count)
    except Exception:
        count = -1
    if count and count > 0:
        for i in range(1, count + 1):
            try:
                yield col.Item(i)
            except Exception:
                try:
                    yield col(i)
                except Exception:
                    break
        return

    i = 1
    while True:
        try:
            yield col.Item(i)
        except Exception:
            try:
                yield col(i)
            except Exception:
                break
        i += 1


def iter_all_shapes(shapes: Any) -> Iterable[Any]:
    return iter_collection(shapes)


def shape_label_name(shape: Any) -> str:
    for attr in ("LabelName", "Label", "BlockLabel"):
        try:
            val = getattr(shape, attr)
            if hasattr(val, "Name"):
                return str(val.Name)
            if val is not None:
                return str(val)
        except Exception:
            continue
    return ""


def normalize_labels(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return [str(v).strip() for v in value if str(v).strip()]
    text = str(value)
    return [v.strip() for v in text.split(",") if v.strip()]


def find_shapes_by_label(model: Any, label_name: str) -> list[Any]:
    shapes: list[Any] = []
    try:
        shape_range = model.Shapes.LabeledAs(label_name)
    except Exception:
        shape_range = None

    if shape_range is not None:
        shapes = list(iter_collection(shape_range))

    if not shapes:
        for shp in iter_all_shapes(model.Shapes):
            if shape_label_name(shp) == label_name:
                shapes.append(shp)
    return shapes


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

    # Try Blocks.InRectangle to get a shape range, then move it.
    try:
        blocks = model.Shapes.Blocks
        sel = _try_selection_in_rect(blocks, qf, x1, y1, x2, y2)
        if sel is not None and try_move(sel, qf, dx, dy) is not None:
            return 1, 1
    except Exception:
        pass

    # Try Selection.InRectangle then Move.
    try:
        sel = _try_selection_in_rect(model.Selection, qf, x1, y1, x2, y2)
        if sel is not None and try_move(sel, qf, dx, dy) is not None:
            return 1, 1
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


def iter_labels_by_type(problem: Any, label_type: int) -> Iterable[Any]:
    try:
        data_doc = problem.DataDoc
    except Exception:
        data_doc = problem

    try:
        labels = data_doc.Labels(label_type)
    except Exception:
        labels = None

    if labels is None:
        return []
    return iter_collection(labels)


def _label_point(label: Any) -> Optional[tuple[float, float]]:
    try:
        pt = label.Point
    except Exception:
        pt = None

    if pt is not None:
        try:
            x0 = float(getattr(pt, "X"))
            y0 = float(getattr(pt, "Y"))
            return x0, y0
        except Exception:
            pass

    for attr in ("X", "Y"):
        if not hasattr(label, attr):
            return None
    try:
        return float(getattr(label, "X")), float(getattr(label, "Y"))
    except Exception:
        return None


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


def rebuild_model(qf: Any, problem: Any) -> None:
    for obj in (problem, qf):
        try:
            obj.Rebuild()
        except Exception:
            pass


def build_mesh(model: Any) -> bool:
    try:
        model.Shapes.BuildMesh()
        return True
    except Exception:
        pass
    try:
        model.BuildMesh()
        return True
    except Exception:
        return False


def remove_mesh(model: Any) -> bool:
    try:
        model.Shapes.RemoveMesh()
        return True
    except Exception:
        pass
    try:
        model.RemoveMesh()
        return True
    except Exception:
        return False


def solve_problem(problem: Any) -> bool:
    try:
        if hasattr(problem, "SolveProblem"):
            problem.SolveProblem()
        elif hasattr(problem, "Solve"):
            problem.Solve()
        else:
            return False
    except Exception:
        return False

    # Wait for solver if busy.
    for _ in range(600):
        try:
            if hasattr(problem, "IsBusy") and not bool(problem.IsBusy):
                break
        except Exception:
            break
        time.sleep(0.5)

    # Analyze results if needed.
    try:
        if hasattr(problem, "AnalyzeResults"):
            problem.AnalyzeResults()
    except Exception:
        pass
    return True


def _com_method_names(obj: Any) -> list[str]:
    if pythoncom is None:
        return []
    try:
        ti = obj._oleobj_.GetTypeInfo()
        attr = ti.GetTypeAttr()
    except Exception:
        return []
    names: set[str] = set()
    for i in range(attr.cFuncs):
        try:
            fd = ti.GetFuncDesc(i)
            n = ti.GetNames(fd.memid)
            if n:
                names.add(n[0])
        except Exception:
            continue
    return sorted(names)


def _numeric_prop(obj: Any, name: str) -> Optional[float]:
    try:
        val = getattr(obj, name)
    except Exception:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    return None


def _string_prop(obj: Any, name: str) -> Optional[str]:
    try:
        val = getattr(obj, name)
    except Exception:
        return None
    if isinstance(val, str):
        return val
    return None


def _iter_circuit_items(circuit: Any) -> Iterable[Any]:
    try:
        count = int(circuit.Count)
    except Exception:
        count = -1
    if count and count > 0:
        for i in range(1, count + 1):
            try:
                yield circuit.Item(i)
            except Exception:
                try:
                    yield circuit(i)
                except Exception:
                    break
        return
    i = 1
    while True:
        try:
            yield circuit.Item(i)
        except Exception:
            try:
                yield circuit(i)
            except Exception:
                break
        i += 1


def dump_force_candidates(obj: Any) -> list[tuple[str, float]]:
    candidates = []
    # Common names to try first.
    for key in ("Fx", "Fy", "Fz", "ForceX", "ForceY", "ForceZ", "Force", "MagneticForce"):
        val = _numeric_prop(obj, key)
        if val is not None:
            candidates.append((key, val))

    # Also scan COM property names containing "Force".
    for name in _com_method_names(obj):
        if "force" not in name.lower():
            continue
        val = _numeric_prop(obj, name)
        if val is not None and (name, val) not in candidates:
            candidates.append((name, val))
    return candidates


def _result_blocks(problem: Any) -> Optional[Any]:
    try:
        res = problem.Result
        if res is None:
            return None
        return res.Blocks
    except Exception:
        return None


def _block_label(block: Any) -> str:
    for attr in ("Label", "Name"):
        try:
            val = getattr(block, attr)
            if val is not None:
                return str(val)
        except Exception:
            continue
    return ""


def list_result_block_labels(problem: Any, limit: int = 50) -> list[str]:
    labels: list[str] = []
    blocks = _result_blocks(problem)
    if blocks is None:
        return labels
    for blk in iter_collection(blocks):
        name = _block_label(blk)
        if name:
            labels.append(name)
        if len(labels) >= limit:
            break
    return labels


def find_result_block(problem: Any, label: str, contains_ok: bool = True) -> Optional[Any]:
    blocks = _result_blocks(problem)
    if blocks is None:
        return None
    label_l = label.lower()
    for blk in iter_collection(blocks):
        name = _block_label(blk)
        name_l = name.lower()
        if name_l == label_l:
            return blk
    if contains_ok:
        for blk in iter_collection(blocks):
            name = _block_label(blk)
            if label_l in name.lower():
                return blk
    return None


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
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
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


def cmd_move_blocks(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
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
    moved = 0
    for label in labels:
        blk = _find_block_by_label(model, label)
        if blk is None:
            print(f"Block not found for label: {label}")
            continue
        try:
            vec = qf.PointXY(dx, dy)
            try:
                blk.Move(0, vec)
                print(f"Moved block '{label}' by dx={dx}, dy={dy} using Move(0, PointXY).")
                moved += 1
                continue
            except Exception:
                pass
        except Exception:
            vec = None

        how = try_move(blk, qf, dx, dy)
        if how is None:
            print(f"Move failed for block '{label}'.")
            continue
        print(f"Moved block '{label}' by dx={dx}, dy={dy} using {how}.")
        moved += 1

    if moved == 0:
        print("No blocks moved.")
        return 5
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
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
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
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
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


def _label_collection(problem: Any) -> Optional[Any]:
    try:
        return problem.DataDoc.Labels(3)
    except Exception:
        pass
    try:
        return problem.Labels(3)
    except Exception:
        return None


def _copy_label_content(src: Any, dst: Any) -> None:
    # Copy numeric properties from src.Content to dst.Content.
    try:
        src_c = src.Content
        dst_c = dst.Content
    except Exception:
        return

    # Explicit common fields first (more reliable than blind setattr).
    common = [
        "Kxx",
        "Kyy",
        "NonLinear",
        "Anisotropic",
        "Polar",
        "Radial",
        "Serial",
        "Coercive",
        "Conductivity",
        "ConductivityEx",
        "TemperatureEx",
        "Loading",
        "LoadingEx",
        "TotalCurrent",
    ]
    for name in common:
        val = _numeric_prop(src_c, name)
        if val is None:
            continue
        try:
            if callable(getattr(dst_c, name, None)):
                getattr(dst_c, name)(val)
            else:
                setattr(dst_c, name, val)
        except Exception:
            continue

    # Fallback: try to copy any other numeric props.
    for name in _com_method_names(src_c):
        val = _numeric_prop(src_c, name)
        if val is None:
            continue
        try:
            if callable(getattr(dst_c, name, None)):
                getattr(dst_c, name)(val)
            else:
                setattr(dst_c, name, val)
        except Exception:
            continue


def cmd_clone_label(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
            return 2

    src_name = args.src.strip()
    dst_name = args.dst.strip()
    amps = float(args.amps) if args.amps else None
    if not src_name or not dst_name:
        print("Missing --src or --dst.")
        return 3

    labels = _label_collection(problem)
    if labels is None:
        print("Failed to access block labels.")
        return 4

    src_lbl = None
    for lbl in iter_collection(labels):
        try:
            if str(lbl.Name).strip().lower() == src_name.lower():
                src_lbl = lbl
                break
        except Exception:
            continue
    if src_lbl is None:
        print(f"Source label not found: {src_name}")
        return 5

    # Create new label.
    new_lbl = None
    try:
        new_lbl = labels.Add()
    except Exception:
        try:
            new_lbl = labels.Insert()
        except Exception:
            new_lbl = None
    if new_lbl is None:
        try:
            new_lbl = labels.Item(int(labels.Count))
        except Exception:
            print("Failed to create new label.")
            return 6

    try:
        new_lbl.Name = dst_name
    except Exception:
        pass

    _copy_label_content(src_lbl, new_lbl)

    # Optionally override current in the cloned label.
    if amps is not None:
        try:
            content = new_lbl.Content
            try:
                if callable(getattr(content, "Loading", None)):
                    content.Loading(amps)
                else:
                    content.Loading = amps
            except Exception:
                pass
            try:
                if callable(getattr(content, "LoadingEx", None)):
                    content.LoadingEx(amps)
                else:
                    content.LoadingEx = amps
            except Exception:
                pass
            try:
                if callable(getattr(content, "TotalCurrent", None)):
                    content.TotalCurrent(amps)
                else:
                    content.TotalCurrent = amps
            except Exception:
                pass
        except Exception:
            pass

    try:
        problem.DataDoc.Save()
    except Exception:
        pass

    if amps is not None:
        print(f"Cloned label '{src_name}' -> '{dst_name}' with current {amps}.")
    else:
        print(f"Cloned label '{src_name}' -> '{dst_name}'.")
    return 0


def _set_coil_label_values(label_obj: Any, amps: float, mu_r: float = 1.0) -> None:
    try:
        content = label_obj.Content
    except Exception:
        return

    # Permeability (relative)
    for name, val in (("Kxx", mu_r), ("Kyy", mu_r)):
        try:
            if callable(getattr(content, name, None)):
                getattr(content, name)(val)
            else:
                setattr(content, name, val)
        except Exception:
            pass

    # Basic flags
    for name in ("NonLinear", "Anisotropic", "Polar", "Radial", "Serial"):
        try:
            if callable(getattr(content, name, None)):
                getattr(content, name)(0.0)
            else:
                setattr(content, name, 0.0)
        except Exception:
            pass

    # Zero magnet/coercive and conductivity
    for name in ("Coercive", "Conductivity", "ConductivityEx"):
        try:
            if callable(getattr(content, name, None)):
                getattr(content, name)(0.0)
            else:
                setattr(content, name, 0.0)
        except Exception:
            pass

    # Total Ampere-Turns
    for name in ("Loading", "LoadingEx", "TotalCurrent"):
        try:
            if callable(getattr(content, name, None)):
                getattr(content, name)(amps)
            else:
                setattr(content, name, amps)
        except Exception:
            pass

    # Ensure type if writable
    try:
        if callable(getattr(content, "Type", None)):
            content.Type(3)
        else:
            content.Type = 3
    except Exception:
        pass


def cmd_create_coil_label(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
            return 2

    name = args.name.strip()
    if not name:
        print("Missing --name.")
        return 3

    labels = _label_collection(problem)
    if labels is None:
        print("Failed to access block labels.")
        return 4

    # Create new label
    new_lbl = None
    try:
        new_lbl = labels.Add()
    except Exception:
        try:
            new_lbl = labels.Insert()
        except Exception:
            new_lbl = None
    if new_lbl is None:
        print("Failed to create new label.")
        return 5

    try:
        new_lbl.Name = name
    except Exception:
        pass

    _set_coil_label_values(new_lbl, float(args.amps), float(args.mu))

    try:
        problem.DataDoc.Save()
    except Exception:
        pass

    print(f"Created coil label '{name}' with mu_r={args.mu} and amps={args.amps}.")
    return 0


def cmd_assign_label(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
            return 2

    model = ensure_model_loaded(problem, Path(args.model) if args.model else None)
    if model is None:
        print("Failed to load model.")
        return 3

    src_name = args.src.strip()
    dst_name = args.dst.strip()
    if not src_name or not dst_name:
        print("Missing --src or --dst.")
        return 4

    try:
        blocks = model.Shapes.Blocks
    except Exception:
        print("Failed to access Shapes.Blocks.")
        return 5

    sel = None
    for method_name in ("LabeledAs", "GetLabeledAs"):
        if hasattr(blocks, method_name):
            try:
                sel = getattr(blocks, method_name)("", "", src_name)
                break
            except Exception:
                sel = None

    if sel is None:
        print(f"No blocks found for label: {src_name}")
        return 6

    ok = False
    try:
        setattr(sel, "Label", dst_name)
        ok = True
    except Exception:
        try:
            sel.Label = dst_name
            ok = True
        except Exception:
            ok = False

    if not ok:
        print("Failed to assign new label to selection.")
        return 7

    print(f"Assigned label '{dst_name}' to blocks previously labeled '{src_name}'.")
    return 0


def cmd_block_bounds(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
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


def cmd_set_current(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
            return 2

    label = args.label.strip()
    if not label:
        print("Missing --label.")
        return 3

    labels = None
    # Prefer DataDoc (AM34.dms) labels
    try:
        labels = problem.DataDoc.Labels(3)
    except Exception:
        labels = None
    # Fallback: problem-level labels
    if labels is None:
        try:
            labels = problem.Labels(3)
        except Exception:
            labels = None
    if labels is None:
        print("Failed to access block labels.")
        return 4

    targets = []
    for lbl in iter_collection(labels):
        try:
            if str(lbl.Name).strip().lower() == label.lower():
                targets.append(lbl)
        except Exception:
            continue

    if not targets:
        print(f"Block label not found: {label}")
        return 5

    amps = float(args.amps)
    changed = 0
    for target in targets:
        try:
            content = target.Content
            before_loading = _numeric_prop(content, "Loading")
            before_loading_ex = _numeric_prop(content, "LoadingEx")
            before_total_flag = _numeric_prop(content, "TotalCurrent")

            # QuickField 6.2 uses Loading/LoadingEx for Total Ampere-Turns in block label.
            # Loading / LoadingEx may be methods in some COM schemas.
            try:
                if callable(getattr(content, "Loading", None)):
                    content.Loading(amps)
                else:
                    content.Loading = amps
            except Exception:
                pass
            try:
                if callable(getattr(content, "LoadingEx", None)):
                    content.LoadingEx(amps)
                else:
                    content.LoadingEx = amps
            except Exception:
                pass

            # Ensure TotalCurrent flag is ON for ampere-turns mode.
            try:
                if callable(getattr(content, "TotalCurrent", None)):
                    content.TotalCurrent(True)
                else:
                    content.TotalCurrent = True
            except Exception:
                pass

            # IMPORTANT: reassign content back to label (per official COM samples).
            try:
                target.Content = content
            except Exception:
                pass

            after = {}
            for key in ("Loading", "LoadingEx", "TotalCurrent"):
                try:
                    after[key] = float(getattr(content, key))
                except Exception:
                    pass

            if before_total_flag is not None:
                print(f"Label '{label}': TotalCurrent {before_total_flag} -> {after.get('TotalCurrent', 'n/a')}")
            if before_loading is not None or before_loading_ex is not None:
                print(f"Label '{label}': Loading {before_loading} -> {after.get('Loading','n/a')}, LoadingEx {before_loading_ex} -> {after.get('LoadingEx','n/a')}")
            print(f"Label '{label}': Loading={after.get('Loading','n/a')}, LoadingEx={after.get('LoadingEx','n/a')}")
            changed += 1
        except Exception as exc:
            print(f"Failed to set current for one label '{label}': {exc}")

    # Persist changes (DataDoc is where block labels live).
    try:
        problem.DataDoc.Save()
    except Exception:
        pass
    try:
        problem.Save()
    except Exception:
        pass

    # Optional: save DataDoc to a new .dms file.
    if getattr(args, "save_dms", ""):
        dms_path = str(args.save_dms)
        try:
            problem.DataDoc.SaveAs(dms_path)
            print(f"DataDoc saved as: {dms_path}")
        except Exception as exc:
            print(f"DataDoc.SaveAs failed: {exc}")

    # Optional: reopen problem to confirm persistence.
    if getattr(args, "reopen", False):
        pbm_path = None
        for attr in ("FullName", "Path"):
            try:
                pbm_path = str(getattr(problem, attr))
                if pbm_path:
                    break
            except Exception:
                pbm_path = None
        try:
            problem.Close()
        except Exception:
            pass
        if pbm_path:
            try:
                qf.Problems.Open(pbm_path)
                problem = qf.ActiveProblem
            except Exception:
                problem = None
        if problem is not None:
            try:
                labels = problem.DataDoc.Labels(3)
            except Exception:
                labels = None
            if labels is not None:
                for lbl in iter_collection(labels):
                    try:
                        if str(lbl.Name).strip().lower() == label.lower():
                            content = lbl.Content
                            try:
                                print(
                                    "Reopen check:",
                                    "Loading",
                                    getattr(content, "Loading", "n/a"),
                                    "LoadingEx",
                                    getattr(content, "LoadingEx", "n/a"),
                                )
                            except Exception:
                                pass
                            break
                    except Exception:
                        continue

    if changed == 0:
        return 6
    print(f"Updated {changed} label(s) named '{label}'.")
    return 0


def cmd_label_dump(args: argparse.Namespace) -> int:
    if win32com is None or pythoncom is None:
        print("pywin32/pythoncom not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
            return 2

    label = args.label.strip()
    if not label:
        print("Missing --label.")
        return 3

    labels = None
    try:
        labels = problem.DataDoc.Labels(3)
    except Exception:
        labels = None
    if labels is None:
        try:
            labels = problem.Labels(3)
        except Exception:
            labels = None
    if labels is None:
        print("Failed to access block labels.")
        return 4

    target = None
    for lbl in iter_collection(labels):
        try:
            if str(lbl.Name).strip().lower() == label.lower():
                target = lbl
                break
        except Exception:
            continue

    if target is None:
        print(f"Block label not found: {label}")
        return 5

    content = None
    try:
        content = target.Content
    except Exception:
        content = None

    if content is None:
        print("Label Content is None.")
        return 6

    # Dump COM properties/methods and numeric values.
    names = _com_method_names(content)
    print(f"Label.Content properties ({len(names)}):")
    for name in names:
        val = _numeric_prop(content, name)
        if val is not None:
            print(f"- {name}: {val}")
        else:
            print(f"- {name}")
    return 0


def cmd_circuit_dump(args: argparse.Namespace) -> int:
    if win32com is None or pythoncom is None:
        print("pywin32/pythoncom not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
            return 2

    try:
        circuit = problem.Circuit
    except Exception as exc:
        print(f"Failed to access Circuit: {exc}")
        return 3

    if circuit is None:
        print("Circuit is None.")
        return 4

    print("Circuit properties:")
    for name in _com_method_names(circuit):
        val = _numeric_prop(circuit, name)
        if val is not None:
            print(f"- {name}: {val}")
        else:
            s = _string_prop(circuit, name)
            if s is not None:
                print(f"- {name}: {s}")
            else:
                print(f"- {name}")

    # Try list items if this is a collection.
    try:
        items = list(_iter_circuit_items(circuit))
    except Exception:
        items = []
    if items:
        print(f"Circuit items: {len(items)}")
        for i, item in enumerate(items[:20], start=1):
            name = _string_prop(item, "Name") or _string_prop(item, "Label") or _string_prop(item, "ID")
            print(f"- [{i}] {name or '<unnamed>'}")
    return 0


def cmd_set_circuit_current(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
            return 2

    try:
        circuit = problem.Circuit
    except Exception as exc:
        print(f"Failed to access Circuit: {exc}")
        return 3

    if circuit is None:
        print("Circuit is None.")
        return 4

    target = args.name.strip().lower()
    amps = float(args.amps)

    # Try set directly on circuit object first.
    for prop in ("Current", "I", "Value", "Amplitude"):
        try:
            if callable(getattr(circuit, prop, None)):
                getattr(circuit, prop)(amps)
            else:
                setattr(circuit, prop, amps)
            print(f"Set Circuit.{prop} = {amps}")
            return 0
        except Exception:
            continue

    # Try to set on circuit items by name.
    changed = 0
    for item in _iter_circuit_items(circuit):
        name = (_string_prop(item, "Name") or _string_prop(item, "Label") or _string_prop(item, "ID") or "").lower()
        if target and target != name:
            continue
        for prop in ("Current", "I", "Value", "Amplitude"):
            try:
                if callable(getattr(item, prop, None)):
                    getattr(item, prop)(amps)
                else:
                    setattr(item, prop, amps)
                print(f"Set {name or '<unnamed>'}.{prop} = {amps}")
                changed += 1
                break
            except Exception:
                continue
    if changed == 0:
        print("No circuit element updated. Use circuit-dump to inspect names/fields.")
        return 5
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


def cmd_com_probe(args: argparse.Namespace) -> int:
    pbm_path = Path(args.pbm)
    if not pbm_path.exists():
        print(f"PBM not found: {pbm_path}")
        return 1
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    app, problem = com_open_problem(pbm_path)
    print("COM probe OK")
    print(f"- Problem: {pbm_path.name}")
    print(f"- ActiveProblem: {problem}")

    # Try to access model after opening referenced model file (if provided).
    model = None
    if args.model:
        model_path = Path(args.model)
        if model_path.exists():
            try:
                app.Models.Open(str(model_path))
                time.sleep(0.2)
                model = problem.Model
            except Exception as exc:
                print(f"Model open error: {exc}")
        else:
            print(f"Model not found: {model_path}")
    else:
        try:
            problem.LoadModel()
            time.sleep(0.2)
            model = problem.Model
        except Exception as exc:
            print(f"LoadModel error: {exc}")

    print(f"- Model loaded: {model is not None}")

    # Try to enumerate block labels (type=3 seems to be block labels in QuickField)
    try:
        labels = problem.Labels(3)
        count = labels.Count if labels is not None else 0
        print(f"- Block labels count: {count}")
    except Exception as exc:
        print(f"- Block labels not accessible: {exc}")

    return 0


def cmd_label_pos(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = None
    if args.pbm:
        pbm_path = Path(args.pbm)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return 1
        qf.Problems.Open(str(pbm_path))
        problem = qf.ActiveProblem
    else:
        problem = get_active_problem(qf)
        if problem is None:
            print("No active problem found. Open a problem in QuickField and retry.")
            return 2

    target = args.name.strip().lower() if args.name else ""
    found = 0
    for lbl in iter_labels_by_type(problem, 3):
        try:
            name = str(getattr(lbl, "Name"))
        except Exception:
            continue
        if target and name.lower() != target:
            continue
        pt = _label_point(lbl)
        if pt is None:
            print(f"{name}: (point unavailable)")
        else:
            print(f"{name}: ({pt[0]}, {pt[1]})")
        found += 1

    if found == 0:
        if target:
            print(f"Block label not found: {args.name}")
        else:
            print("No block labels found.")
        return 3
    return 0


def cmd_solve_force(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    pbm_path = Path(args.pbm) if args.pbm else None
    if pbm_path is None or not pbm_path.exists():
        print("PBM not found. Provide --pbm with a valid .pbm file.")
        return 1

    qf.Problems.Open(str(pbm_path))
    problem = qf.ActiveProblem
    if problem is None:
        print("Failed to open problem.")
        return 2

    model = ensure_model_loaded(problem, Path(args.model) if args.model else None)
    if model is None:
        print("Failed to load model.")
        return 3

    if args.remesh:
        remove_mesh(model)
    if args.mesh:
        if not build_mesh(model):
            print("BuildMesh failed.")
            return 4

    if not solve_problem(problem):
        print("SolveProblem failed.")
        return 5

    # Force extraction
    if problem.Result is None:
        print("Result is None after solve. Try opening results in UI and re-run.")
        return 6

    label = args.label.strip()
    target = find_result_block(problem, label) if label else None
    if target is None:
        print(f"Result block not found for label: {label}")
        names = list_result_block_labels(problem, limit=20)
        if names:
            print("Available result block labels (first 20):", ", ".join(names))
        else:
            print("No result block labels available. Check if results are generated.")
        return 6

    candidates = dump_force_candidates(target)
    if not candidates:
        print("No force-like properties found on result block.")
        print("Try inspecting available properties via 'result-dump' if needed.")
        return 7

    print(f"Force candidates for '{label}':")
    for name, val in candidates:
        print(f"- {name}: {val}")
    return 0


def cmd_result_dump(args: argparse.Namespace) -> int:
    if win32com is None or pythoncom is None:
        print("pywin32/pythoncom not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    pbm_path = Path(args.pbm) if args.pbm else None
    if pbm_path is None or not pbm_path.exists():
        print("PBM not found. Provide --pbm with a valid .pbm file.")
        return 1

    qf.Problems.Open(str(pbm_path))
    problem = qf.ActiveProblem
    if problem is None:
        print("Failed to open problem.")
        return 2

    if args.solve:
        if not solve_problem(problem):
            print("SolveProblem failed.")
            return 3

    res = problem.Result
    if res is None:
        print("Result is None. Run solve first.")
        return 4

    label = args.label.strip()
    target = find_result_block(problem, label) if label else None
    if target is None:
        print(f"Result block not found for label: {label}")
        return 5

    names = _com_method_names(target)
    print(f"Result block properties ({len(names)}):")
    for name in names:
        val = _numeric_prop(target, name)
        if val is not None:
            print(f"- {name}: {val}")
        else:
            print(f"- {name}")
    return 0


def cmd_solve_integral(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    pbm_path = Path(args.pbm) if args.pbm else None
    if pbm_path is None or not pbm_path.exists():
        print("PBM not found. Provide --pbm with a valid .pbm file.")
        return 1

    qf.Problems.Open(str(pbm_path))
    problem = qf.ActiveProblem
    if problem is None:
        print("Failed to open problem.")
        return 2

    model = ensure_model_loaded(problem, Path(args.model) if args.model else None)
    if model is None:
        print("Failed to load model.")
        return 3

    # Optional: set coil current in memory before solving.
    if getattr(args, "current_label", "") and getattr(args, "amps", ""):
        label = args.current_label.strip()
        amps = float(args.amps)
        labels = _label_collection(problem)
        if labels is not None:
            for lbl in iter_collection(labels):
                try:
                    if str(lbl.Name).strip().lower() == label.lower():
                        content = lbl.Content
                        try:
                            if callable(getattr(content, "Loading", None)):
                                content.Loading(amps)
                            else:
                                content.Loading = amps
                        except Exception:
                            pass
                        try:
                            if callable(getattr(content, "LoadingEx", None)):
                                content.LoadingEx(amps)
                            else:
                                content.LoadingEx = amps
                        except Exception:
                            pass
                        try:
                            if callable(getattr(content, "TotalCurrent", None)):
                                content.TotalCurrent(True)
                            else:
                                content.TotalCurrent = True
                        except Exception:
                            pass
                        try:
                            lbl.Content = content
                        except Exception:
                            pass
                        # Debug: show current values after setting
                        if args.debug_current:
                            try:
                                print(
                                    f"Current set for '{label}': "
                                    f"Loading={getattr(content,'Loading','n/a')}, "
                                    f"LoadingEx={getattr(content,'LoadingEx','n/a')}, "
                                    f"TotalCurrent={getattr(content,'TotalCurrent','n/a')}"
                                )
                            except Exception:
                                pass
                        # Mark data as changed
                        try:
                            problem.DataDoc.Save()
                        except Exception:
                            pass
                        break
                except Exception:
                    continue

    if args.remesh:
        remove_mesh(model)
    if args.mesh:
        if not build_mesh(model):
            print("BuildMesh failed.")
            return 4

    if not solve_problem(problem):
        print("SolveProblem failed.")
        return 5

    try:
        res = problem.Result
    except Exception:
        res = None
    if res is None:
        print("Result is None after solve. Open results in UI and retry.")
        return 6

    try:
        field = res.GetFieldWindow(1)
        contour = field.Contour
    except Exception as exc:
        print(f"Failed to access FieldWindow/Contour: {exc}")
        return 7

    labels = [v.strip() for v in args.labels.split(",") if v.strip()]
    if not labels:
        print("No labels provided. Use --labels \"steel mover\"")
        return 8

    # Add blocks to contour.
    added = 0
    for name in labels:
        ok = False
        for meth in ("AddBlock1", "AddBlock", "AddBlock2"):
            if hasattr(contour, meth):
                try:
                    getattr(contour, meth)(name)
                    ok = True
                    break
                except Exception:
                    pass
        if ok:
            added += 1
        else:
            print(f"Failed to add block to contour: {name}")

    if added == 0:
        print("No blocks added to contour. Check label names.")
        return 9

    # Maxwell force integral: 15 in QuickField API.
    integral_id = int(args.integral_id)
    try:
        val = res.GetIntegral(integral_id, contour)
        if hasattr(val, "Value"):
            val = val.Value
    except Exception as exc:
        print(f"GetIntegral failed: {exc}")
        return 10

    # Output vector components if present.
    if hasattr(val, "X") and hasattr(val, "Y"):
        print(f"Integral {integral_id} result: X={val.X}, Y={val.Y}")
    else:
        print(f"Integral {integral_id} result: {val}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="QuickField automation starter")
    parser.add_argument(
        "--config",
        default=str(Path(__file__).resolve().parents[1] / "config" / "settings.json"),
        help="Path to settings.json",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_probe = sub.add_parser("probe", help="Check QuickField paths")
    p_probe.set_defaults(func=cmd_probe)

    p_sweep = sub.add_parser("sweep", help="Sweep mover position via QLMCall")
    p_sweep.add_argument("--start", required=True, help="Start position (mm)")
    p_sweep.add_argument("--end", required=True, help="End position (mm)")
    p_sweep.add_argument("--step", required=True, help="Step (mm)")
    p_sweep.add_argument("--y", default="0", help="Y offset for 'any' mode (mm)")
    p_sweep.add_argument(
        "--mode",
        choices=["any", "x", "y"],
        default="any",
        help="Displacement mode configured in LabelMover",
    )
    p_sweep.add_argument(
        "--fixed",
        action="append",
        default=[],
        help="Fixed variation values before displacement (repeatable)",
    )
    p_sweep.add_argument(
        "--output",
        default=str(Path(__file__).resolve().parents[1] / "outputs" / "force_vs_pos.csv"),
        help="CSV output path",
    )
    p_sweep.add_argument(
        "--clear-results",
        action="store_true",
        help="Clear LabelMover result table before sweep",
    )
    p_sweep.add_argument("--verbose", action="store_true", help="Print each step")
    p_sweep.set_defaults(func=cmd_sweep)

    p_table = sub.add_parser("table", help="Run QLMCall for each row in a CSV table")
    p_table.add_argument("--table", required=True, help="CSV with variation values")
    p_table.add_argument(
        "--vars",
        default="",
        help="Comma-separated variation order matching LabelMover (e.g., x_offset,i1,i2,i3,i4)",
    )
    p_table.add_argument(
        "--ignore-header",
        action="store_true",
        help="Ignore column names; use CSV column order as QLMCall parameter order",
    )
    p_table.add_argument(
        "--output",
        default=str(Path(__file__).resolve().parents[1] / "outputs" / "table_results.csv"),
        help="CSV output path",
    )
    p_table.add_argument(
        "--clear-results",
        action="store_true",
        help="Clear LabelMover result table before run",
    )
    p_table.add_argument("--verbose", action="store_true", help="Print each row")
    p_table.set_defaults(func=cmd_table)

    p_gen = sub.add_parser("gen-cases", help="Generate current x position cases CSV")
    p_gen.add_argument("--start", required=True, help="Start position (mm)")
    p_gen.add_argument("--end", required=True, help="End position (mm)")
    p_gen.add_argument("--step", required=True, help="Step (mm)")
    p_gen.add_argument(
        "--currents",
        required=True,
        help="Comma-separated 8 current values (A-turns), e.g. 600,-600,400,-400,300,-300,200,-200",
    )
    p_gen.add_argument(
        "--current-name",
        default="I",
        help="Column name for current parameter (matches LabelMover step name)",
    )
    p_gen.add_argument(
        "--output",
        default=str(Path(__file__).resolve().parents[1] / "inputs" / "cases.csv"),
        help="CSV output path",
    )
    p_gen.set_defaults(func=cmd_gen_cases)

    p_com_probe = sub.add_parser("com-probe", help="Probe ActiveField COM access")
    p_com_probe.add_argument("--pbm", required=True, help="Path to .pbm file")
    p_com_probe.add_argument("--model", default="", help="Optional path to .mod file")
    p_com_probe.set_defaults(func=cmd_com_probe)

    p_label = sub.add_parser("label-pos", help="Print block label coordinates")
    p_label.add_argument("--name", default="", help="Block label name (exact match)")
    p_label.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_label.set_defaults(func=cmd_label_pos)

    p_solve = sub.add_parser("solve-force", help="Build mesh, solve, and dump force candidates")
    p_solve.add_argument("--pbm", required=True, help="Path to .pbm file")
    p_solve.add_argument("--model", default="", help="Optional path to .mod file")
    p_solve.add_argument("--label", required=True, help="Block label name for force extraction")
    p_solve.add_argument("--mesh", action="store_true", help="Build mesh before solving")
    p_solve.add_argument("--remesh", action="store_true", help="Remove mesh before building")
    p_solve.set_defaults(func=cmd_solve_force)

    p_dump = sub.add_parser("result-dump", help="Dump result block properties")
    p_dump.add_argument("--pbm", required=True, help="Path to .pbm file")
    p_dump.add_argument("--label", required=True, help="Block label name for result block")
    p_dump.add_argument("--solve", action="store_true", help="Solve before dumping")
    p_dump.set_defaults(func=cmd_result_dump)

    p_int = sub.add_parser("solve-integral", help="Solve and compute integral on contour")
    p_int.add_argument("--pbm", required=True, help="Path to .pbm file")
    p_int.add_argument("--model", default="", help="Optional path to .mod file")
    p_int.add_argument("--labels", required=True, help="Comma-separated block labels for contour")
    p_int.add_argument("--integral-id", default="15", help="Integral ID (15=Maxwell force)")
    p_int.add_argument("--mesh", action="store_true", help="Build mesh before solving")
    p_int.add_argument("--remesh", action="store_true", help="Remove mesh before building")
    p_int.add_argument("--current-label", default="", help="Block label to set current for")
    p_int.add_argument("--amps", default="", help="Total Ampere-Turns to set before solve")
    p_int.add_argument("--debug-current", action="store_true", help="Print current values after set")
    p_int.set_defaults(func=cmd_solve_integral)

    p_move = sub.add_parser("move-block", help="Move a block by its label")
    p_move.add_argument("--label", required=True, help="Block label name")
    p_move.add_argument("--dx", default="0", help="Delta X")
    p_move.add_argument("--dy", default="0", help="Delta Y")
    p_move.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_move.add_argument("--model", default="", help="Optional path to .mod file")
    p_move.set_defaults(func=cmd_move_block)

    p_moves = sub.add_parser("move-blocks", help="Move multiple blocks by labels")
    p_moves.add_argument("--labels", required=True, help="Comma-separated block labels")
    p_moves.add_argument("--dx", default="0", help="Delta X")
    p_moves.add_argument("--dy", default="0", help="Delta Y")
    p_moves.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_moves.add_argument("--model", default="", help="Optional path to .mod file")
    p_moves.set_defaults(func=cmd_move_blocks)

    p_moves_once = sub.add_parser("move-blocks-once", help="Move multiple blocks as one rigid set")
    p_moves_once.add_argument("--labels", required=True, help="Comma-separated block labels")
    p_moves_once.add_argument("--dx", default="0", help="Delta X")
    p_moves_once.add_argument("--dy", default="0", help="Delta Y")
    p_moves_once.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_moves_once.add_argument("--model", default="", help="Optional path to .mod file")
    p_moves_once.add_argument("--debug", action="store_true", help="Print selection diagnostics")
    p_moves_once.set_defaults(func=cmd_move_blocks_once)

    p_list = sub.add_parser("list-blocks", help="List block labels in model")
    p_list.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_list.add_argument("--model", default="", help="Optional path to .mod file")
    p_list.set_defaults(func=cmd_list_blocks)

    p_bounds = sub.add_parser("block-bounds", help="Print block bounds by label")
    p_bounds.add_argument("--labels", required=True, help="Comma-separated block labels")
    p_bounds.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_bounds.add_argument("--model", default="", help="Optional path to .mod file")
    p_bounds.set_defaults(func=cmd_block_bounds)

    p_cur = sub.add_parser("set-current", help="Set TotalCurrent on a block label")
    p_cur.add_argument("--label", required=True, help="Block label name (e.g., bobine)")
    p_cur.add_argument("--amps", required=True, help="Total current (A-turns)")
    p_cur.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_cur.add_argument("--reopen", action="store_true", help="Reopen problem to confirm persistence")
    p_cur.add_argument("--save-dms", default="", help="Save DataDoc to a new .dms file")
    p_cur.set_defaults(func=cmd_set_current)

    p_clone = sub.add_parser("clone-label", help="Clone a block label")
    p_clone.add_argument("--src", required=True, help="Source label name (e.g., bobine)")
    p_clone.add_argument("--dst", required=True, help="New label name (e.g., bobine_100)")
    p_clone.add_argument("--amps", default="", help="Override Total Ampere-Turns on new label")
    p_clone.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_clone.set_defaults(func=cmd_clone_label)

    p_assign = sub.add_parser("assign-label", help="Assign a label to blocks with another label")
    p_assign.add_argument("--src", required=True, help="Existing label to replace")
    p_assign.add_argument("--dst", required=True, help="New label to assign")
    p_assign.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_assign.add_argument("--model", default="", help="Optional path to .mod file")
    p_assign.set_defaults(func=cmd_assign_label)

    p_coil = sub.add_parser("create-coil-label", help="Create a coil label with explicit values")
    p_coil.add_argument("--name", required=True, help="New label name (e.g., bobine_100)")
    p_coil.add_argument("--amps", required=True, help="Total Ampere-Turns (e.g., 100)")
    p_coil.add_argument("--mu", default="1", help="Relative permeability (default 1)")
    p_coil.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_coil.set_defaults(func=cmd_create_coil_label)

    p_ld = sub.add_parser("label-dump", help="Dump label Content properties")
    p_ld.add_argument("--label", required=True, help="Block label name (e.g., bobine)")
    p_ld.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_ld.set_defaults(func=cmd_label_dump)

    p_cd = sub.add_parser("circuit-dump", help="Dump circuit properties/items")
    p_cd.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_cd.set_defaults(func=cmd_circuit_dump)

    p_sc = sub.add_parser("set-circuit-current", help="Set current on circuit/element")
    p_sc.add_argument("--name", default="", help="Circuit element name (if applicable)")
    p_sc.add_argument("--amps", required=True, help="Current value")
    p_sc.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_sc.set_defaults(func=cmd_set_circuit_current)

    p_model = sub.add_parser("model", help="Auto-model via ActiveField COM plan")
    p_model.add_argument(
        "--plan",
        default=str(Path(__file__).resolve().parents[1] / "config" / "modeling.json"),
        help="Path to modeling plan JSON",
    )
    p_model.add_argument("--pbm", default="", help="Override PBM path in plan")
    p_model.add_argument("--model", default="", help="Override model path in plan")
    p_model.add_argument("--save-as", default="", help="Override save_model_as in plan")
    p_model.add_argument(
        "--use-active",
        action="store_true",
        help="Use active problem if no PBM is provided",
    )
    p_model.add_argument("--dry-run", action="store_true", help="Validate plan only")
    p_model.set_defaults(func=cmd_model)

    return parser


def main(argv: list[str]) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
