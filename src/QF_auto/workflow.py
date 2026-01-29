from __future__ import annotations

import argparse
import csv
import re
from pathlib import Path
from typing import Any, Sequence

from .connection import dispatch_qf_app, open_problem, ensure_model_loaded
from .geometry import _block_bounds, _find_block_by_label, list_block_labels, move_blocks_in_rect
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


def _parse_indices(raw: str, max_n: int) -> list[int]:
    s = raw.strip().lower()
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

    amps = _prompt_float("1) Bobine current (Total Ampere-Turns)", None)
    current_label = args.current_label.strip() if args.current_label else "bobine"
    set_label_current(problem, current_label, amps, qf=qf)

    raw_indices = _prompt("2) Move which blocks? (e.g., 2,5 or 2-4 or all)")
    indices = _parse_indices(raw_indices, len(labels))
    if not indices:
        print("No valid indices selected.")
        return 4
    move_labels = [labels[i - 1] for i in indices]
    print(f"Moving labels: {', '.join(move_labels)}")

    left = _prompt_float("3) Move left distance", 0.0)
    right = _prompt_float("4) Move right distance", 0.0)
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

    rect = _union_bounds(bounds)
    current_pos = 0.0

    results: list[tuple[float, float, float]] = []
    integral_id = int(args.integral_id)

    for pos in positions:
        dx = pos - current_pos
        if abs(dx) > 1e-9:
            moved, _ = move_blocks_in_rect(model, qf, rect, dx, 0.0)
            if moved == 0:
                print(f"Move failed at position {pos}.")
                return 8
            rect = (rect[0] + dx, rect[1], rect[2] + dx, rect[3])
            current_pos = pos

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
            fy = float(getattr(val, "Y"))
        except Exception as exc:
            print(f"Integral failed at position {pos}: {exc}")
            return 13

        results.append((pos, fx, fy))
        print(f"pos={pos}: Fx={fx}, Fy={fy}")

    if abs(current_pos) > 1e-9:
        move_blocks_in_rect(model, qf, rect, -current_pos, 0.0)

    if args.out:
        out_path = Path(args.out)
        with out_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["position", "force_x", "force_y"])
            for row in results:
                writer.writerow(row)
        print(f"Saved table to {out_path}")
    else:
        print("Results:")
        for pos, fx, fy in results:
            print(f"{pos}\t{fx}\t{fy}")
    return 0
