import argparse
from typing import Any, Optional, Sequence

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover - runtime dependency
    win32com = None


def _dispatch_app() -> Any:
    if win32com is None:
        raise RuntimeError("pywin32 is not installed. Run: pip install pywin32")
    try:
        return win32com.client.gencache.EnsureDispatch("QuickField.Application")
    except Exception:
        return win32com.client.Dispatch("QuickField.Application")


def _get_active_problem(qf: Any) -> Optional[Any]:
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


def _get_shapes_by_label(model: Any, label_name: str) -> Any:
    # Shapes.LabeledAs may accept a label name string.
    shapes = model.Shapes
    return shapes.LabeledAs(label_name)


def _rebuild(qf: Any, prb: Any) -> None:
    # Try common rebuild hooks; ignore failures.
    for obj in (prb, qf):
        try:
            obj.Rebuild()
        except Exception:
            pass


def _try_moves(target: Any, qf: Any, dx: float, dy: float) -> Optional[str]:
    # ActiveField Move signature is not explicit in some docs; try common variants.
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

    for name, params in attempts:
        try:
            target.Move(*params)
            return name
        except Exception:
            continue
    return None


def _iter_shapes(shape_range: Any):
    try:
        count = int(shape_range.Count)
    except Exception:
        count = -1
    if count and count > 0:
        for i in range(1, count + 1):
            try:
                yield shape_range.Item(i)
            except Exception:
                try:
                    yield shape_range(i)
                except Exception:
                    break
        return

    i = 1
    while True:
        try:
            yield shape_range.Item(i)
        except Exception:
            try:
                yield shape_range(i)
            except Exception:
                break
        i += 1


def _iter_all_shapes(shapes: Any):
    try:
        count = int(shapes.Count)
    except Exception:
        count = -1
    if count and count > 0:
        for i in range(1, count + 1):
            try:
                yield shapes.Item(i)
            except Exception:
                try:
                    yield shapes(i)
                except Exception:
                    break
        return

    i = 1
    while True:
        try:
            yield shapes.Item(i)
        except Exception:
            try:
                yield shapes(i)
            except Exception:
                break
        i += 1


def _shape_label_name(shape: Any) -> str:
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


def _parse_labels(value: str) -> list[str]:
    return [v.strip() for v in value.split(",") if v.strip()]


def main() -> int:
    parser = argparse.ArgumentParser(description="ActiveField label move test")
    parser.add_argument(
        "--label",
        required=True,
        help="Exact label name(s), comma-separated (block labels)",
    )
    parser.add_argument("--delta", type=float, default=0.5, help="Delta X (mm)")
    parser.add_argument("--dy", type=float, default=0.0, help="Delta Y (mm)")
    parser.add_argument("--no-restore", action="store_true", help="Do not restore original position")
    args = parser.parse_args()

    qf = _dispatch_app()
    prb = _get_active_problem(qf)
    if prb is None:
        print("No active problem found. Open a problem in QuickField and retry.")
        return 2

    model = prb.Model
    if model is None:
        try:
            prb.LoadModel()
            model = prb.Model
        except Exception as exc:
            print(f"Model not loaded: {exc}")
            return 3

    labels = _parse_labels(args.label)
    if not labels:
        print("No label names provided.")
        return 4

    shape_ranges = []
    for name in labels:
        try:
            shape_range = _get_shapes_by_label(model, name)
        except Exception as exc:
            print(f"Shapes.LabeledAs failed for '{name}': {exc}")
            shape_range = None

        shapes: list[Any] = []
        if shape_range is not None:
            shapes = list(_iter_shapes(shape_range))

        if not shapes:
            # Fallback: scan all shapes and match by label name
            for shp in _iter_all_shapes(model.Shapes):
                if _shape_label_name(shp) == name:
                    shapes.append(shp)

        print(f"Shapes labeled '{name}': {len(shapes)}")
        if not shapes:
            print(f"Warning: no shapes found for '{name}'.")
        shape_ranges.append((name, shapes))

    if not any(shapes for _, shapes in shape_ranges):
        print("No shapes found for any label. Movement aborted.")
        print("Tip: ensure geometry shapes are labeled in the model, not just data labels.")
        return 6

    move_used = None
    for name, shapes in shape_ranges:
        for shp in shapes:
            move_used = _try_moves(shp, qf, args.delta, args.dy)
            if move_used is None:
                print(f"Move failed for '{name}'. Could not find a compatible Move signature.")
                return 7
    print(f"Moved shapes using {move_used} (dx={args.delta}, dy={args.dy})")
    _rebuild(qf, prb)

    if not args.no_restore:
        for name, shapes in shape_ranges:
            for shp in shapes:
                move_back = _try_moves(shp, qf, -args.delta, -args.dy)
                if move_back is None:
                    print(f"Restore move failed for '{name}'.")
                    return 6
        _rebuild(qf, prb)
        print("Restored shapes to original position.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
