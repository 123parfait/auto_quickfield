from __future__ import annotations
import argparse
import time
from pathlib import Path
from typing import Any, Iterable, Optional

from .connection import (
    win32com,
    pythoncom,
    dispatch_qf_app,
    open_problem,
    ensure_model_loaded,
    normalize_labels,
    close_data_windows,
    _com_method_names,
    _numeric_prop,
    _string_prop,
)
from .geometry import _block_bounds
from .labels import _label_point

def rebuild_model(qf: Any, problem: Any) -> None:
    for obj in (problem, qf):
        try:
            obj.Rebuild()
        except Exception:
            pass

def build_mesh(model: Any) -> bool:
    def _try_build(target: Any) -> bool:
        if target is None:
            return False
        for args in ((True, False), (True,), ()):
            try:
                target.BuildMesh(*args)
                return True
            except Exception:
                continue
        return False

    shapes = None
    try:
        shapes = model.Shapes
    except Exception:
        shapes = None

    if _try_build(shapes):
        return True
    if _try_build(model):
        return True
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

def solve_problem(problem: Any, model: Optional[Any] = None) -> bool:
    def _try_solve() -> Optional[Exception]:
        try:
            if hasattr(problem, "SolveProblem"):
                problem.SolveProblem()
            elif hasattr(problem, "Solve"):
                problem.Solve()
            else:
                return Exception("No SolveProblem/Solve method on problem.")
            return None
        except Exception as exc:
            return exc

    exc = _try_solve()
    if exc is not None:
        msg = str(exc)
        msg_l = msg.lower()
        if "has no mesh" in msg_l and "air" in msg_l:
            if model is not None:
                try:
                    build_mesh(model)
                except Exception:
                    pass
                exc_retry = _try_solve()
                if exc_retry is None:
                    return True
                exc = exc_retry
                msg = str(exc)
                msg_l = msg.lower()
                if "has no mesh" in msg_l and "air" in msg_l:
                    print(f"SolveProblem warning (ignored): {exc}")
                    return True
            else:
                print(f"SolveProblem warning (ignored): {exc}")
                return True
        print(f"SolveProblem failed: {exc}")
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

def cmd_circuit_dump(args: argparse.Namespace) -> int:
    if win32com is None or pythoncom is None:
        print("pywin32/pythoncom not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = open_problem(qf, args.pbm)
    if problem is None:
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
    problem = open_problem(qf, args.pbm)
    if problem is None:
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

    if not solve_problem(problem, model):
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
                        close_data_windows(qf)
                        break
                except Exception:
                    continue

    if args.remesh:
        remove_mesh(model)
    if args.mesh:
        if not build_mesh(model):
            print("BuildMesh failed.")
            return 4

    if not solve_problem(problem, model):
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
