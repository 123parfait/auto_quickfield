from __future__ import annotations

import argparse
from pathlib import Path
import sys

from .connection import cmd_probe, cmd_sweep, cmd_table, cmd_gen_cases
from .geometry import (
    cmd_move_block,
    cmd_move_blocks_once,
    cmd_list_blocks,
    cmd_block_bounds,
    cmd_model,
)
from .labels import (
    cmd_clone_label,
    cmd_create_coil_label,
    cmd_assign_label,
    cmd_set_current,
    cmd_label_dump,
    cmd_label_pos,
)
from .solve import (
    cmd_circuit_dump,
    cmd_set_circuit_current,
    cmd_com_probe,
    cmd_solve_force,
    cmd_result_dump,
    cmd_solve_integral,
)
from .workflow import cmd_batch_force


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="QuickField automation helper")
    sub = parser.add_subparsers(dest="command", required=True)

    p_probe = sub.add_parser("probe", help="Print QuickField COM probe")
    p_probe.set_defaults(func=cmd_probe)

    p_sweep = sub.add_parser("sweep", help="Sweep QLMCall over a range of parameters")
    p_sweep.add_argument("--qlm", default="", help="Path to QLM file")
    p_sweep.add_argument("--start", required=True, help="Start value")
    p_sweep.add_argument("--end", required=True, help="End value")
    p_sweep.add_argument("--step", required=True, help="Step value")
    p_sweep.add_argument("--out", default="", help="Output CSV path")
    p_sweep.set_defaults(func=cmd_sweep)

    p_table = sub.add_parser("table", help="Run QLMCall from a table of inputs")
    p_table.add_argument("--qlm", default="", help="Path to QLM file")
    p_table.add_argument("--table", required=True, help="CSV table path")
    p_table.add_argument("--out", default="", help="Output CSV path")
    p_table.set_defaults(func=cmd_table)

    p_cases = sub.add_parser("gen-cases", help="Generate cases CSV from template")
    p_cases.add_argument("--start", required=True, help="Start value")
    p_cases.add_argument("--end", required=True, help="End value")
    p_cases.add_argument("--step", required=True, help="Step value")
    p_cases.add_argument("--out", default="", help="Output CSV path")
    p_cases.set_defaults(func=cmd_gen_cases)

    p_move = sub.add_parser("move-block", help="Move block with a given label")
    p_move.add_argument("--label", required=True, help="Block label name")
    p_move.add_argument("--dx", required=True, help="Delta X")
    p_move.add_argument("--dy", required=True, help="Delta Y")
    p_move.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_move.add_argument("--model", default="", help="Optional path to .mod file")
    p_move.set_defaults(func=cmd_move_block)

    p_move_once = sub.add_parser("move-blocks-once", help="Move blocks as a group (union bounds)")
    p_move_once.add_argument("--labels", required=True, help="Comma-separated labels")
    p_move_once.add_argument("--dx", required=True, help="Delta X")
    p_move_once.add_argument("--dy", required=True, help="Delta Y")
    p_move_once.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_move_once.add_argument("--model", default="", help="Optional path to .mod file")
    p_move_once.add_argument("--debug", action="store_true", help="Verbose selection debug")
    p_move_once.set_defaults(func=cmd_move_blocks_once)

    p_list = sub.add_parser("list-blocks", help="List block labels in the model")
    p_list.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_list.add_argument("--model", default="", help="Optional path to .mod file")
    p_list.set_defaults(func=cmd_list_blocks)

    p_bounds = sub.add_parser("block-bounds", help="Get bounds for block labels")
    p_bounds.add_argument("--labels", required=True, help="Comma-separated labels")
    p_bounds.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_bounds.add_argument("--model", default="", help="Optional path to .mod file")
    p_bounds.set_defaults(func=cmd_block_bounds)

    p_clone = sub.add_parser("clone-label", help="Clone a label to a new name")
    p_clone.add_argument("--src", required=True, help="Source label name")
    p_clone.add_argument("--dst", required=True, help="Destination label name")
    p_clone.add_argument("--amps", default="", help="Optional coil amps override")
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

    p_sc = sub.add_parser("set-current", help="Set coil current on a label")
    p_sc.add_argument("--label", required=True, help="Block label name (e.g., bobine)")
    p_sc.add_argument("--amps", required=True, help="Current value")
    p_sc.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_sc.add_argument("--reopen", action="store_true", help="Re-open data doc to verify")
    p_sc.add_argument("--save-dms", default="", help="Save DataDoc to .dms path")
    p_sc.set_defaults(func=cmd_set_current)

    p_ld = sub.add_parser("label-dump", help="Dump label Content properties")
    p_ld.add_argument("--label", required=True, help="Block label name (e.g., bobine)")
    p_ld.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_ld.set_defaults(func=cmd_label_dump)

    p_lp = sub.add_parser("label-pos", help="Print block label position")
    p_lp.add_argument("--label", required=True, help="Block label name")
    p_lp.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_lp.set_defaults(func=cmd_label_pos)

    p_cd = sub.add_parser("circuit-dump", help="Dump circuit properties/items")
    p_cd.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_cd.set_defaults(func=cmd_circuit_dump)

    p_scirc = sub.add_parser("set-circuit-current", help="Set current on circuit/element")
    p_scirc.add_argument("--name", default="", help="Circuit element name (if applicable)")
    p_scirc.add_argument("--amps", required=True, help="Current value")
    p_scirc.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_scirc.set_defaults(func=cmd_set_circuit_current)

    p_model = sub.add_parser("model", help="Auto-model via ActiveField COM plan")
    p_model.add_argument(
        "--plan",
        default=str(Path(__file__).resolve().parents[2] / "config" / "modeling.json"),
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

    p_com = sub.add_parser("com-probe", help="Dump COM method names")
    p_com.add_argument("--pbm", default="", help="Open PBM if no active problem")
    p_com.add_argument("--model", default="", help="Optional path to .mod file")
    p_com.set_defaults(func=cmd_com_probe)

    p_force = sub.add_parser("solve-force", help="Solve and dump mechanical force")
    p_force.add_argument("--pbm", required=True, help="PBM path")
    p_force.add_argument("--label", required=True, help="Block label")
    p_force.add_argument("--mesh", action="store_true", help="Build mesh before solving")
    p_force.add_argument("--remesh", action="store_true", help="Remove mesh before rebuild")
    p_force.add_argument("--solve", action="store_true", help="Force solve before result")
    p_force.set_defaults(func=cmd_solve_force)

    p_dump = sub.add_parser("result-dump", help="Dump result blocks/force candidates")
    p_dump.add_argument("--pbm", required=True, help="PBM path")
    p_dump.add_argument("--label", default="", help="Block label to inspect")
    p_dump.add_argument("--solve", action="store_true", help="Force solve before result")
    p_dump.set_defaults(func=cmd_result_dump)

    p_int = sub.add_parser("solve-integral", help="Solve and evaluate integral")
    p_int.add_argument("--pbm", required=True, help="PBM path")
    p_int.add_argument("--labels", required=True, help="Comma-separated block labels")
    p_int.add_argument("--model", default="", help="Optional path to .mod file")
    p_int.add_argument("--integral-id", default="15", help="Integral ID (default 15 = Maxwell force)")
    p_int.add_argument("--mesh", action="store_true", help="Build mesh before solving")
    p_int.add_argument("--remesh", action="store_true", help="Remove mesh before rebuild")
    p_int.add_argument("--solve", action="store_true", help="Force solve before result")
    p_int.add_argument("--current-label", default="", help="Label to update current")
    p_int.add_argument("--amps", default="", help="Total Ampere-Turns value")
    p_int.add_argument("--debug-current", action="store_true", help="Print current update info")
    p_int.set_defaults(func=cmd_solve_integral)

    p_batch = sub.add_parser("batch-force", help="Interactive sweep: move blocks and compute force table")
    p_batch.add_argument("--pbm", default="", help="PBM path (optional if a problem is already open)")
    p_batch.add_argument("--model", default="", help="Optional path to .mod file")
    p_batch.add_argument("--current-label", default="bobine", help="Current label name (default bobine)")
    p_batch.add_argument("--integral-id", default="15", help="Integral ID (default 15)")
    p_batch.add_argument("--mesh", action="store_true", help="Build mesh before solving")
    p_batch.add_argument("--remesh", action="store_true", help="Remove mesh before rebuild")
    p_batch.add_argument("--mesh-once", action="store_true", help="Build mesh once per case (faster, less accurate)")
    p_batch.add_argument("--sleep", default="0", help="Sleep seconds between steps (e.g., 0.5)")
    p_batch.add_argument("--out", default="", help="Optional CSV output path")
    p_batch.set_defaults(func=cmd_batch_force)

    return parser


def main(argv: list[str]) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
