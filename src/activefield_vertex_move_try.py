import argparse
from typing import Any, Optional

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


def _iter_collection(col: Any):
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


def _parse_labels(value: str) -> list[str]:
    return [v.strip() for v in value.split(",") if v.strip()]


def _rebuild(qf: Any, prb: Any) -> None:
    for obj in (prb, qf):
        try:
            obj.Rebuild()
        except Exception:
            pass


def _move_vertex(vtx: Any, qf: Any, dx: float, dy: float) -> bool:
    # Try Point property/field first
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

    # Fallback: try Move() with no args
    try:
        vtx.Move()
        return True
    except Exception:
        return False


def main() -> int:
    parser = argparse.ArgumentParser(description="Try moving vertices for labeled blocks")
    parser.add_argument("--labels", required=True, help="Comma-separated block labels")
    parser.add_argument("--delta", type=float, default=1.0, help="Delta X (mm)")
    parser.add_argument("--dy", type=float, default=0.0, help="Delta Y (mm)")
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

    labels = _parse_labels(args.labels)
    if not labels:
        print("No labels provided.")
        return 4

    sel = model.Selection
    for name in labels:
        try:
            ret = sel.LabeledAs(name)
            if hasattr(ret, "Vertices"):
                sel = ret
        except Exception as exc:
            print(f"Selection.LabeledAs failed for '{name}': {exc}")
            return 5

    try:
        vertices = sel.Vertices
    except Exception as exc:
        print(f"Selection.Vertices failed: {exc}")
        return 6

    moved = 0
    total = 0
    for vtx in _iter_collection(vertices):
        total += 1
        if _move_vertex(vtx, qf, args.delta, args.dy):
            moved += 1

    print(f"Vertices moved: {moved}/{total}")
    if moved > 0:
        _rebuild(qf, prb)
        print("Rebuild done.")
    else:
        print("No vertices moved. This COM path might be read-only.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
