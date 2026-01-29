import argparse
from typing import Any, Optional, Sequence, Tuple

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


def _find_shape_by_label(model: Any, label_name: str) -> Any:
    for shp in _iter_all_shapes(model.Shapes):
        if _shape_label_name(shp) == label_name:
            return shp
    raise ValueError(f"No shape found for label '{label_name}'.")


def _get_point_xy(qf: Any, x: float, y: float):
    return qf.PointXY(x, y)


def _try_move_variants(shape: Any, qf: Any, dx: float, dy: float) -> Tuple[bool, str]:
    attempts: Sequence[Tuple[str, Tuple[Any, ...]]] = []
    try:
        p_delta = _get_point_xy(qf, dx, dy)
        p_origin = _get_point_xy(qf, 0.0, 0.0)
        attempts = [
            ("Move(dx,dy)", (dx, dy)),
            ("Move(PointXY)", (p_delta,)),
            ("Move(PointXY,PointXY)", (p_origin, p_delta)),
        ]
    except Exception:
        attempts = [("Move(dx,dy)", (dx, dy))]

    # Some versions expose Move() with no params (interactive or implicit).
    attempts = [("Move()", tuple())] + list(attempts)

    for name, params in attempts:
        try:
            shape.Move(*params)
            return True, name
        except Exception:
            continue
    return False, "Move(...) failed"


def _try_point_assign(shape: Any, qf: Any, dx: float, dy: float) -> Tuple[bool, str]:
    try:
        pt = shape.Point
    except Exception as exc:
        return False, f"Point read failed: {exc}"

    try:
        x0 = float(getattr(pt, "X"))
        y0 = float(getattr(pt, "Y"))
    except Exception as exc:
        return False, f"Point.X/Y read failed: {exc}"

    try:
        new_pt = _get_point_xy(qf, x0 + dx, y0 + dy)
    except Exception as exc:
        return False, f"PointXY failed: {exc}"

    try:
        shape.Point = new_pt
        return True, "Point=PointXY"
    except Exception:
        pass

    try:
        setattr(pt, "X", x0 + dx)
        setattr(pt, "Y", y0 + dy)
        return True, "Point.X/Y set"
    except Exception as exc:
        return False, f"Point set failed: {exc}"


def _rebuild(qf: Any, prb: Any) -> None:
    for obj in (prb, qf):
        try:
            obj.Rebuild()
        except Exception:
            pass


def main() -> int:
    parser = argparse.ArgumentParser(description="Try move variants on a labeled shape")
    parser.add_argument("--label", required=True, help="Exact label name")
    parser.add_argument("--delta", type=float, default=1.0, help="Delta X (mm)")
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

    shape = _find_shape_by_label(model, args.label)
    print(f"Found shape for '{args.label}'.")

    ok, how = _try_move_variants(shape, qf, args.delta, args.dy)
    if not ok:
        ok, how = _try_point_assign(shape, qf, args.delta, args.dy)

    if not ok:
        print(f"Move failed: {how}")
        return 4

    print(f"Moved using {how} (dx={args.delta}, dy={args.dy})")
    _rebuild(qf, prb)

    if not args.no_restore:
        ok2, how2 = _try_move_variants(shape, qf, -args.delta, -args.dy)
        if not ok2:
            ok2, how2 = _try_point_assign(shape, qf, -args.delta, -args.dy)
        if ok2:
            _rebuild(qf, prb)
            print("Restored to original position.")
        else:
            print(f"Restore failed: {how2}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
