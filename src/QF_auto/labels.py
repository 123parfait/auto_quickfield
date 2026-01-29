from __future__ import annotations
import argparse
from pathlib import Path
from typing import Any, Iterable, Optional

from .connection import (
    win32com,
    dispatch_qf_app,
    open_problem,
    ensure_model_loaded,
    iter_collection,
    close_data_windows,
    _com_method_names,
    _numeric_prop,
)
from .geometry import _find_block_by_label

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
    problem = open_problem(qf, args.pbm)
    if problem is None:
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
    close_data_windows(qf)

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
    problem = open_problem(qf, args.pbm)
    if problem is None:
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
    close_data_windows(qf)

    print(f"Created coil label '{name}' with mu_r={args.mu} and amps={args.amps}.")
    return 0

def cmd_assign_label(args: argparse.Namespace) -> int:
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

def cmd_set_current(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = open_problem(qf, args.pbm)
    if problem is None:
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
    close_data_windows(qf)
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


def set_label_current(problem: Any, label: str, amps: float, qf: Optional[Any] = None) -> int:
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
        return 0

    targets = []
    for lbl in iter_collection(labels):
        try:
            if str(lbl.Name).strip().lower() == label.lower():
                targets.append(lbl)
        except Exception:
            continue

    if not targets:
        print(f"Block label not found: {label}")
        return 0

    changed = 0
    for target in targets:
        try:
            content = target.Content
            before_loading = _numeric_prop(content, "Loading")
            before_loading_ex = _numeric_prop(content, "LoadingEx")
            before_total_flag = _numeric_prop(content, "TotalCurrent")

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

    try:
        problem.DataDoc.Save()
    except Exception:
        pass
    try:
        problem.Save()
    except Exception:
        pass
    if qf is not None:
        close_data_windows(qf)
    print(f"Updated {changed} label(s) named '{label}'.")
    return changed

def cmd_label_dump(args: argparse.Namespace) -> int:
    if win32com is None or pythoncom is None:
        print("pywin32/pythoncom not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = open_problem(qf, args.pbm)
    if problem is None:
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

def cmd_label_pos(args: argparse.Namespace) -> int:
    if win32com is None:
        print("pywin32 is not available. Install with: pip install pywin32")
        return 1

    qf = dispatch_qf_app()
    problem = open_problem(qf, args.pbm)
    if problem is None:
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
