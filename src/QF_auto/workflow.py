from __future__ import annotations

import argparse
import csv
import re
from pathlib import Path
from typing import Any, Sequence

from .connection import dispatch_qf_app, open_problem, ensure_model_loaded
from .geometry import (
    _block_bounds,
    _find_block_by_label,
    _collect_vertices_for_labels,
    list_block_labels,
    move_blocks_in_rect,
    move_vertex,
)
from .labels import set_label_current
from .solve import build_mesh, remove_mesh, solve_problem


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


def _prompt_left_right() -> tuple[float, float]:
    raw = _prompt("3) Move left distance", "0.0")
    raw = raw.replace("，", ",").replace("；", ",").replace(";", ",")
    if "," in raw:
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if len(parts) >= 2:
            try:
                left = float(parts[0])
                right = float(parts[1])
                return left, right
            except ValueError:
                pass
    left = _prompt_float("3) Move left distance", 0.0) if raw == "0.0" else float(raw)
    right = _prompt_float("4) Move right distance", 0.0)
    return left, right


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

    # de-dup while preserving order
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


def _parse_float_list(raw: str) -> list[float]:
    s = raw.strip()
    if not s:
        return []
    s = s.replace("，", ",").replace("；", ",").replace(";", ",")
    parts = re.split(r"[,\s]+", s)
    out: list[float] = []
    for part in parts:
        if not part:
            continue
        out.append(float(part))
    return out


def _positions(left: float, right: float, step: float) -> list[float]:
    if step <= 0:
        raise ValueError("step must be > 0")
    start = -abs(left)
    end = abs(right)
    pos = start
    out: list[float] = []
    while pos <= end + 1e-9:
        out.append(round(pos, 6))
        pos += step
    return out


def _union_bounds(bounds: Sequence[tuple[float, float, float, float]]) -> tuple[float, float, float, float]:
    left = min(b[0] for b in bounds)
    bottom = min(b[1] for b in bounds)
    right = max(b[2] for b in bounds)
    top = max(b[3] for b in bounds)
    return left, bottom, right, top


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

    raw_currents = _prompt("1) Bobine current list (e.g., 100,200,300)")
    try:
        currents = _parse_float_list(raw_currents)
    except ValueError:
        print("Invalid current list.")
        return 3
    if not currents:
        print("No currents provided.")
        return 3
    current_label = args.current_label.strip() if args.current_label else "bobine"

    raw_indices = _prompt("2) Move which blocks? (e.g., 2,5 or 2-4 or all)")
    indices = _parse_indices(raw_indices, len(labels))
    if not indices:
        print("No valid indices selected.")
        return 4
    move_labels = [labels[i - 1] for i in indices]
    print(f"Moving labels: {', '.join(move_labels)}")

    try:
        left, right = _prompt_left_right()
    except ValueError:
        print("Please enter numbers like -3,3 or separate values.")
        return 5
    step = _prompt_float("5) Step size", 1.0)

    try:
        positions = _positions(left, right, step)
    except ValueError as exc:
        print(str(exc))
        return 5

    bounds = []
    for name in move_labels:
        blk = _find_block_by_label(model, name)
        if blk is None:
            print(f"Block not found: {name}")
            return 6
        b = _block_bounds(blk)
        if b is None:
            print(f"Bounds unavailable for: {name}")
            return 7
        bounds.append(b)

    base_rect = _union_bounds(bounds)
    verts = _collect_vertices_for_labels(model, move_labels)
    integral_id = int(args.integral_id)

    table: dict[float, dict[float, float]] = {}
    def _move_by(dx: float, rect_now: tuple[float, float, float, float], cur_pos: float) -> tuple[bool, tuple[float, float, float, float], float]:
        if abs(dx) <= 1e-9:
            return True, rect_now, cur_pos
        moved, _ = move_blocks_in_rect(model, qf, rect_now, dx, 0.0, epsilon=1e-6)
        if moved == 0 and verts:
            moved = 0
            for vtx in verts:
                if move_vertex(vtx, qf, dx, 0.0):
                    moved += 1
        if moved == 0:
            return False, rect_now, cur_pos
        rect_now = (rect_now[0] + dx, rect_now[1], rect_now[2] + dx, rect_now[3])
        cur_pos = cur_pos + dx
        return True, rect_now, cur_pos

    for amps in currents:
        set_label_current(problem, current_label, amps, qf=qf)
        rect = base_rect
        current_pos = 0.0
        try:
            for pos in positions:
                dx = pos - current_pos
                ok, rect, current_pos = _move_by(dx, rect, current_pos)
                if not ok:
                    print(f"Move failed at position {pos}.")
                    return 8

                if args.remesh:
                    remove_mesh(model)
                if args.mesh:
                    if not build_mesh(model):
                        print("BuildMesh failed.")
                        return 9

                if not solve_problem(problem):
                    print("SolveProblem failed.")
                    return 10

                try:
                    res = problem.Result
                except Exception:
                    res = None
                if res is None:
                    print("Result is None after solve.")
                    return 11

                try:
                    field = res.GetFieldWindow(1)
                    contour = field.Contour
                except Exception as exc:
                    print(f"Failed to access FieldWindow/Contour: {exc}")
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

                try:
                    val = res.GetIntegral(integral_id, contour).Value
                    fx = float(getattr(val, "X"))
                except Exception as exc:
                    print(f"Integral failed at position {pos}: {exc}")
                    return 13

                table.setdefault(pos, {})[amps] = fx
                print(f"I={amps} pos={pos}: Fx={fx}")
        finally:
            if abs(current_pos) > 1e-9:
                ok, rect, current_pos = _move_by(-current_pos, rect, current_pos)
                if not ok:
                    print("Warning: failed to return to start position.")

    if args.out:
        out_path = Path(args.out)
        with out_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            header = ["position"] + [f"I={amps}" for amps in currents]
            writer.writerow(header)
            for pos in positions:
                row = [pos]
                for amps in currents:
                    row.append(table.get(pos, {}).get(amps, ""))
                writer.writerow(row)
        print(f"Saved table to {out_path}")
    else:
        print("Results:")
        header = ["position"] + [f"I={amps}" for amps in currents]
        print("\t".join(map(str, header)))
        for pos in positions:
            row = [pos] + [table.get(pos, {}).get(amps, "") for amps in currents]
            print("\t".join(map(str, row)))
    return 0
