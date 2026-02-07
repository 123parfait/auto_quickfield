from __future__ import annotations

import argparse
import csv
import re
import threading
import time
from pathlib import Path
from typing import Any, Callable, Optional, Sequence

from .connection import dispatch_qf_app, open_problem, ensure_model_loaded
from .geometry import (
    _block_bounds,
    _find_block_by_label,
    _collect_vertices_for_labels,
    list_block_labels,
    move_blocks_in_rect,
    move_block_labels,
    move_vertex,
)
from .labels import set_label_field
from .solve import build_mesh, remove_mesh, solve_problem, rebuild_model


def _prompt(text: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default is not None else ""
    while True:
        val = input(f"{text}{suffix}: ").strip()
        if val:
            return val
        if default is not None:
            return default


def _prompt_float(text: str, default: float | None = None) -> float:
    while True:
        raw = _prompt(text, f"{default}" if default is not None else None)
        try:
            return float(raw)
        except ValueError:
            print("Please enter a number.")


def _prompt_start_end() -> tuple[float, float, float, float]:
    while True:
        raw = _prompt("3) Start position x0,y0 (e.g., 0,0)", "0,0")
        raw = raw.replace("，", ",").replace("；", ",").replace(";", ",")
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if len(parts) != 2:
            print("Start position must be x,y. Try again.")
            continue
        try:
            x0 = float(parts[0])
            y0 = float(parts[1])
        except ValueError:
            print("Start position must be numbers. Try again.")
            continue
        break

    while True:
        raw = _prompt("4) End position x1,y1 (e.g., 3,0)", "0,0")
        raw = raw.replace("，", ",").replace("；", ",").replace(";", ",")
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if len(parts) != 2:
            print("End position must be x,y. Try again.")
            continue
        try:
            x1 = float(parts[0])
            y1 = float(parts[1])
        except ValueError:
            print("End position must be numbers. Try again.")
            continue
        break
    return x0, y0, x1, y1


def _parse_indices(raw: str, max_n: int) -> list[int]:
    s = raw.strip().lower()
    s = s.replace("，", ",").replace("；", ",").replace(";", ",")
    if s in ("all", "*"):
        return list(range(1, max_n + 1))

    parts = re.split(r"[,\s]+", s)
    out: list[int] = []
    for part in parts:
        if not part:
            continue
        if "-" in part:
            a_str, b_str = part.split("-", 1)
            a = int(a_str)
            b = int(b_str)
            if a > b:
                a, b = b, a
            out.extend(range(a, b + 1))
        else:
            out.append(int(part))

    seen: set[int] = set()
    cleaned: list[int] = []
    for idx in out:
        if idx < 1 or idx > max_n:
            continue
        if idx in seen:
            continue
        seen.add(idx)
        cleaned.append(idx)
    return cleaned


def _parse_label_values(raw: str) -> dict[str, list[float]]:
    s = raw.strip()
    if not s:
        return {}
    s = s.replace("，", ",").replace("；", ";")
    s = s.replace(";", ",")

    # Find label= positions, labels may contain spaces.
    matches = list(re.finditer(r"([^=,]+?)=", s))
    if not matches:
        return {}

    mapping: dict[str, list[float]] = {}
    for i, match in enumerate(matches):
        label = match.group(1).strip()
        if not label:
            continue
        start_val = match.end()
        end_val = matches[i + 1].start() if i + 1 < len(matches) else len(s)
        value_str = s[start_val:end_val].strip()
        if value_str.startswith(","):
            value_str = value_str[1:].strip()
        if not value_str:
            raise ValueError("Each label must have at least one value")
        parts = [p.strip() for p in value_str.split(",") if p.strip()]
        if not parts:
            raise ValueError("Each label must have at least one value")
        vals = [float(p) for p in parts]
        mapping.setdefault(label, []).extend(vals)
    return mapping


def _cases_from_mapping_all(mapping: dict[str, list[float]]) -> list[dict[str, float]]:
    if not mapping:
        return []
    if len(mapping) == 1:
        name = next(iter(mapping))
        return [{name: v} for v in mapping[name]]
    names = list(mapping.keys())
    combos: list[dict[str, float]] = [{}]
    for name in names:
        new_combos: list[dict[str, float]] = []
        for combo in combos:
            for v in mapping[name]:
                item = dict(combo)
                item[name] = v
                new_combos.append(item)
        combos = new_combos
    return combos


def _cases_from_mapping_pair(mapping: dict[str, list[float]]) -> list[dict[str, float]]:
    if not mapping:
        return []
    if len(mapping) == 1:
        name = next(iter(mapping))
        return [{name: v} for v in mapping[name]]
    counts = {name: len(vals) for name, vals in mapping.items()}
    if len(set(counts.values())) != 1:
        raise ValueError("Pair mode requires equal counts for each label.")
    n = next(iter(counts.values()))
    names = list(mapping.keys())
    cases: list[dict[str, float]] = []
    for i in range(n):
        case = {name: mapping[name][i] for name in names}
        cases.append(case)
    return cases


def _positions_line(x0: float, y0: float, x1: float, y1: float, step: float) -> list[tuple[float, float]]:
    if step <= 0:
        raise ValueError("step must be > 0")
    dx = x1 - x0
    dy = y1 - y0
    length = (dx * dx + dy * dy) ** 0.5
    if length <= 1e-9:
        return [(round(x0, 6), round(y0, 6))]
    steps = max(1, int(length / step))
    out: list[tuple[float, float]] = []
    for i in range(steps + 1):
        t = i / steps
        out.append((round(x0 + dx * t, 6), round(y0 + dy * t, 6)))
    return out


def _union_bounds(bounds: Sequence[tuple[float, float, float, float]]) -> tuple[float, float, float, float]:
    left = min(b[0] for b in bounds)
    bottom = min(b[1] for b in bounds)
    right = max(b[2] for b in bounds)
    top = max(b[3] for b in bounds)
    return left, bottom, right, top


def _wait_idle(obj: Any, timeout_s: float = 120.0, interval_s: float = 0.2) -> None:
    if obj is None:
        return
    loops = int(timeout_s / interval_s)
    for _ in range(max(1, loops)):
        try:
            busy_attr = getattr(obj, "IsBusy", None)
        except Exception:
            return
        if busy_attr is None:
            return
        try:
            busy = busy_attr() if callable(busy_attr) else bool(busy_attr)
        except Exception:
            return
        if not busy:
            return
        time.sleep(interval_s)


def _clear_selection(model: Any) -> None:
    if model is None:
        return
    selections = []
    try:
        selections.append(model.Selection)
    except Exception:
        pass
    try:
        shapes = model.Shapes
        sel = getattr(shapes, "Selection", None)
        if sel is not None:
            selections.append(sel)
    except Exception:
        pass

    for sel in selections:
        for name in ("Clear", "ClearSelection", "ClearAll"):
            try:
                method = getattr(sel, name, None)
                if callable(method):
                    method()
                    break
            except Exception:
                continue


def _ensure_full_mesh(model: Any, qf: Any, problem: Any, force_remesh: bool) -> bool:
    _clear_selection(model)
    rebuild_model(qf, problem)

    if force_remesh:
        remove_mesh(model)
        if not build_mesh(model):
            return False
        _wait_idle(model)
        _wait_idle(problem)
        return True

    if build_mesh(model):
        _wait_idle(model)
        _wait_idle(problem)
        return True

    remove_mesh(model)
    if not build_mesh(model):
        return False
    _wait_idle(model)
    _wait_idle(problem)
    return True


def run_batch_force_plan(
    pbm: str,
    model_path: str,
    move_labels: Sequence[str],
    cases: Sequence[dict[str, float]],
    field_name: str,
    positions: Sequence[tuple[float, float]],
    integrals: Sequence[tuple[str, int]],
    mesh: bool,
    remesh: bool,
    mesh_once: bool,
    sleep_s: float,
    out_path: str,
    log: Optional[Callable[[str], None]] = None,
    cancel: Optional[threading.Event] = None,
) -> int:
    def emit(msg: str) -> None:
        if log is not None:
            log(msg)
        else:
            print(msg)

    def is_cancelled() -> bool:
        return cancel is not None and cancel.is_set()

    if not move_labels:
        emit("No moving labels provided.")
        return 3
    if not cases:
        emit("No cases provided.")
        return 4
    if not positions:
        emit("No positions provided.")
        return 5
    if not integrals:
        emit("No output values selected.")
        return 6

    qf = dispatch_qf_app()
    problem = open_problem(qf, pbm)
    if problem is None:
        return 1

    model = ensure_model_loaded(problem, Path(model_path) if model_path else None)
    if model is None:
        emit("Failed to load model.")
        return 2

    def _integral_components(val: Any) -> dict[str, float]:
        comps: dict[str, float] = {}
        for axis in ("X", "Y", "Z"):
            if hasattr(val, axis):
                try:
                    comps[axis] = float(getattr(val, axis))
                except Exception:
                    pass
        if comps:
            return comps
        try:
            return {"": float(val)}
        except Exception:
            return {"": float("nan")}

    x0, y0 = positions[0]
    bounds = []
    for name in move_labels:
        blk = _find_block_by_label(model, name)
        if blk is None:
            emit(f"Block not found: {name}")
            return 6
        b = _block_bounds(blk)
        if b is None:
            emit(f"Bounds unavailable for: {name}")
            return 7
        bounds.append(b)

    base_rect = _union_bounds(bounds)
    verts = _collect_vertices_for_labels(model, move_labels)
    mesh_once_effective = bool(mesh_once)
    if mesh and mesh_once:
        emit("Note: --mesh-once is ignored because geometry moves each step.")
        mesh_once_effective = False

    def _move_by_xy(
        dx: float,
        dy: float,
        rect_now: tuple[float, float, float, float],
        cur_x: float,
        cur_y: float,
    ) -> tuple[bool, tuple[float, float, float, float], float, float]:
        if abs(dx) <= 1e-9 and abs(dy) <= 1e-9:
            return True, rect_now, cur_x, cur_y
        moved = 0
        moved, _ = move_blocks_in_rect(model, qf, rect_now, dx, dy, epsilon=1e-6)
        if moved == 0 and verts:
            for vtx in verts:
                if move_vertex(vtx, qf, dx, dy):
                    moved += 1
            if moved > 0:
                try:
                    move_block_labels(problem, qf, move_labels, dx, dy)
                except Exception:
                    pass
        if moved == 0:
            return False, rect_now, cur_x, cur_y
        rect_now = (rect_now[0] + dx, rect_now[1] + dy, rect_now[2] + dx, rect_now[3] + dy)
        cur_x = cur_x + dx
        cur_y = cur_y + dy
        return True, rect_now, cur_x, cur_y

    table: dict[tuple[float, float], dict[int, dict[str, float]]] = {}
    seen_components: dict[str, list[str]] = {name: [] for name, _ in integrals}

    step_sleep = max(0.0, float(sleep_s or 0))
    cancelled = False

    for idx, case in enumerate(cases, start=1):
        if is_cancelled():
            cancelled = True
            break
        for name, val in case.items():
            set_label_field(problem, [name], field_name, val, qf=qf, log=log)

        rect = base_rect
        cur_x = 0.0
        cur_y = 0.0
        try:
            ok, rect, cur_x, cur_y = _move_by_xy(x0 - cur_x, y0 - cur_y, rect, cur_x, cur_y)
            if not ok:
                emit("Move failed at start position.")
                return 8

            if mesh_once_effective and mesh:
                if not _ensure_full_mesh(model, qf, problem, force_remesh=remesh):
                    emit("BuildMesh failed.")
                    return 9

            for dx, dy in positions:
                if is_cancelled():
                    cancelled = True
                    break
                ok, rect, cur_x, cur_y = _move_by_xy(dx - cur_x, dy - cur_y, rect, cur_x, cur_y)
                if not ok:
                    emit(f"Move failed at position dx={dx}, dy={dy}.")
                    return 8

                if not mesh_once_effective and mesh:
                    if not _ensure_full_mesh(model, qf, problem, force_remesh=remesh):
                        emit("BuildMesh failed.")
                        return 9

                if not solve_problem(problem, model):
                    emit("SolveProblem failed.")
                    return 10

                try:
                    res = problem.Result
                except Exception:
                    res = None
                if res is None:
                    emit("Result is None after solve.")
                    return 11

                try:
                    field = res.GetFieldWindow(1)
                    contour = field.Contour
                except Exception as exc:
                    emit(f"Failed to access FieldWindow/Contour: {exc}")
                    return 12

                try:
                    contour.Clear()
                except Exception:
                    pass
                for name in move_labels:
                    try:
                        contour.AddBlock1(name)
                    except Exception:
                        pass

                key = (float(dx), float(dy))
                table.setdefault(key, {}).setdefault(idx, {})
                outputs: list[str] = []
                for integral_name, integral_id in integrals:
                    try:
                        val = res.GetIntegral(integral_id, contour)
                        if hasattr(val, "Value"):
                            val = val.Value
                        comps = _integral_components(val)
                    except Exception as exc:
                        emit(f"Integral failed ({integral_name}) at dx={dx}, dy={dy}: {exc}")
                        return 13

                    for comp_name, comp_val in comps.items():
                        if comp_name not in seen_components[integral_name]:
                            seen_components[integral_name].append(comp_name)
                        col_name = integral_name if comp_name == "" else f"{integral_name}.{comp_name}"
                        table[key][idx][col_name] = comp_val
                        outputs.append(f"{col_name}={comp_val}")

                case_label = ",".join(f"{k}={v}" for k, v in case.items())
                emit(f"{case_label} dx={dx} dy={dy}: " + ", ".join(outputs))
                if step_sleep > 0:
                    time.sleep(step_sleep)
        finally:
            if abs(cur_x) > 1e-9 or abs(cur_y) > 1e-9:
                ok, rect, cur_x, cur_y = _move_by_xy(-cur_x, -cur_y, rect, cur_x, cur_y)
                if not ok:
                    emit("Warning: failed to return to start position.")
                elif mesh:
                    if not _ensure_full_mesh(model, qf, problem, force_remesh=False):
                        emit("Warning: BuildMesh failed after return to start position.")

        if cancelled:
            break

    if cancelled:
        emit("Canceled.")
        return 99

    if out_path:
        out_path_obj = Path(out_path)
        out_path_obj.parent.mkdir(parents=True, exist_ok=True)
        with out_path_obj.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            case_labels = [",".join(f"{k}={v}" for k, v in case.items()) for case in cases]
            output_cols: list[str] = []
            for integral_name, _ in integrals:
                comps = seen_components.get(integral_name, [])
                if not comps:
                    comps = [""]
                for comp in comps:
                    output_cols.append(integral_name if comp == "" else f"{integral_name}.{comp}")
            header = ["dx", "dy"] + [f"{case_label}:{col}" for case_label in case_labels for col in output_cols]
            writer.writerow(header)
            for dx, dy in positions:
                row = [dx, dy]
                key = (float(dx), float(dy))
                for idx in range(1, len(cases) + 1):
                    cell_map = table.get(key, {}).get(idx, {})
                    for col in output_cols:
                        row.append(cell_map.get(col, ""))
                writer.writerow(row)
        emit(f"Saved table to {out_path_obj}")
    else:
        emit("Results:")
        case_labels = [",".join(f"{k}={v}" for k, v in case.items()) for case in cases]
        output_cols = []
        for integral_name, _ in integrals:
            comps = seen_components.get(integral_name, [])
            if not comps:
                comps = [""]
            for comp in comps:
                output_cols.append(integral_name if comp == "" else f"{integral_name}.{comp}")
        header = ["dx", "dy"] + [f"{case_label}:{col}" for case_label in case_labels for col in output_cols]
        emit("\t".join(map(str, header)))
        for dx, dy in positions:
            key = (float(dx), float(dy))
            row = [dx, dy]
            for idx in range(1, len(cases) + 1):
                cell_map = table.get(key, {}).get(idx, {})
                for col in output_cols:
                    row.append(cell_map.get(col, ""))
            emit("\t".join(map(str, row)))
    return 0


def cmd_batch_force(args: argparse.Namespace) -> int:
    qf = dispatch_qf_app()
    problem = open_problem(qf, args.pbm)
    if problem is None:
        return 1

    model = ensure_model_loaded(problem, Path(args.model) if args.model else None)
    if model is None:
        print("Failed to load model.")
        return 2

    labels = list_block_labels(model)
    if not labels:
        print("No block labels found.")
        return 3

    print("Select moving blocks by index:")
    for i, name in enumerate(labels, start=1):
        print(f"{i}. {name}")

    while True:
        mode = _prompt("1.0) Case mode [all/pair]", "all").strip().lower()
        if mode in ("all", "pair"):
            break
        print("Invalid mode. Use 'all' or 'pair'. Try again.")

    while True:
        raw_cases = _prompt(
            "1) Values per label (same format for all/pair). Example: bobine1=100,200; bobine2=300,400"
        )
        try:
            mapping = _parse_label_values(raw_cases)
            cases = _cases_from_mapping_pair(mapping) if mode == "pair" else _cases_from_mapping_all(mapping)
        except ValueError:
            print("Invalid format. Use label=value, separate labels with ';'. Try again.")
            continue
        if not cases:
            print("No cases provided. Try again.")
            continue
        break

    while True:
        field_name = _prompt("1.2) Which field to modify? (e.g., Loading, Kxx, Kyy)", "Loading")
        if field_name.strip():
            break
        print("Field name cannot be empty. Try again.")

    while True:
        raw_indices = _prompt("2) Move which blocks? (e.g., 2,5 or 2-4 or all)")
        indices = _parse_indices(raw_indices, len(labels))
        if indices:
            break
        print("No valid indices selected. Try again.")
    move_labels = [labels[i - 1] for i in indices]
    print(f"Moving labels: {', '.join(move_labels)}")

    x0, y0, x1, y1 = _prompt_start_end()
    step = _prompt_float("5) Step size", 1.0)
    while step <= 0:
        print("Step size must be > 0. Try again.")
        step = _prompt_float("5) Step size", 1.0)
    try:
        positions = _positions_line(x0, y0, x1, y1, step)
    except ValueError as exc:
        print(str(exc))
        return 5

    return run_batch_force_plan(
        pbm=args.pbm,
        model_path=args.model,
        move_labels=move_labels,
        cases=cases,
        field_name=field_name,
        positions=positions,
        integrals=[(f"Integral{int(args.integral_id)}", int(args.integral_id))],
        mesh=bool(args.mesh),
        remesh=bool(args.remesh),
        mesh_once=bool(args.mesh_once),
        sleep_s=float(getattr(args, "sleep", 0) or 0),
        out_path=str(getattr(args, "out", "") or ""),
    )
