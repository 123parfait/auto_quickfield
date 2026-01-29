from __future__ import annotations
import argparse
import csv
import json
import re
import subprocess
import time
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable, Optional

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


def open_problem(qf: Any, pbm_arg: str) -> Optional[Any]:
    if pbm_arg:
        pbm_path = Path(pbm_arg)
        if not pbm_path.exists():
            print(f"PBM not found: {pbm_path}")
            return None
        qf.Problems.Open(str(pbm_path))
        return qf.ActiveProblem

    problem = get_active_problem(qf)
    if problem is None:
        print("No active problem found. Open a problem in QuickField and retry.")
        return None
    return problem


def _com_method_names(obj: Any) -> list[str]:
    names = []
    try:
        names = [n for n in dir(obj) if not n.startswith("_")]
    except Exception:
        names = []
    return sorted(names)


def _numeric_prop(obj: Any, name: str) -> Optional[float]:
    try:
        val = getattr(obj, name)
    except Exception:
        return None
    try:
        if callable(val):
            return float(val())
        return float(val)
    except Exception:
        return None


def _string_prop(obj: Any, name: str) -> Optional[str]:
    try:
        val = getattr(obj, name)
    except Exception:
        return None
    try:
        if callable(val):
            val = val()
        if isinstance(val, str):
            return val
        return str(val)
    except Exception:
        return None


def close_data_windows(qf: Any) -> int:
    closed = 0
    try:
        windows = qf.Windows
    except Exception:
        return 0

    try:
        count = int(windows.Count)
    except Exception:
        count = -1

    def _close(win: Any) -> bool:
        for name in ("Close", "CloseWindow"):
            try:
                meth = getattr(win, name, None)
                if callable(meth):
                    meth()
                    return True
            except Exception:
                continue
        return False

    def _title(win: Any) -> str:
        for attr in ("Caption", "Title", "Name"):
            try:
                val = getattr(win, attr)
                if isinstance(val, str):
                    return val
            except Exception:
                continue
        return ""

    targets = []
    if count and count > 0:
        for i in range(1, count + 1):
            try:
                targets.append(windows.Item(i))
            except Exception:
                try:
                    targets.append(windows(i))
                except Exception:
                    break
    else:
        i = 1
        while True:
            try:
                targets.append(windows.Item(i))
            except Exception:
                try:
                    targets.append(windows(i))
                except Exception:
                    break
            i += 1

    for win in targets:
        title = _title(win).lower()
        if ".dms" in title or "data" in title:
            if _close(win):
                closed += 1
    return closed

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
