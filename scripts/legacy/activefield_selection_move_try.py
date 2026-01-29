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


def _rebuild(qf: Any, prb: Any) -> None:
    for obj in (prb, qf):
        try:
            obj.Rebuild()
        except Exception:
            pass


def _parse_labels(value: str) -> list[str]:
    return [v.strip() for v in value.split(",") if v.strip()]


def main() -> int:
    parser = argparse.ArgumentParser(description="Try moving Selection by adjusting Left/Right")
    parser.add_argument("--labels", required=True, help="Comma-separated label names")
    parser.add_argument("--delta", type=float, default=1.0, help="Delta X (mm)")
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

    sel = model.Selection
    labels = _parse_labels(args.labels)
    if not labels:
        print("No labels provided.")
        return 4

    # Build selection by label
    for name in labels:
        try:
            ret = sel.LabeledAs(name)
            if hasattr(ret, "Move"):
                sel = ret
        except Exception as exc:
            print(f"Selection.LabeledAs failed for '{name}': {exc}")
            return 5

    # Try Left/Right shift
    try:
        left = float(getattr(sel, "Left"))
        right = float(getattr(sel, "Right"))
        print(f"Selection bounds: Left={left}, Right={right}")
    except Exception as exc:
        print(f"Selection.Left/Right read failed: {exc}")
        left = right = None

    moved = False
    if left is not None and right is not None:
        try:
            setattr(sel, "Left", left + args.delta)
            setattr(sel, "Right", right + args.delta)
            moved = True
            print("Shifted Selection.Left/Right.")
        except Exception as exc:
            print(f"Setting Left/Right failed: {exc}")

    # Try Move() with no args
    try:
        sel.Move()
        print("Called Selection.Move().")
        moved = True
    except Exception as exc:
        print(f"Selection.Move() failed: {exc}")

    if moved:
        _rebuild(qf, prb)
        print("Rebuild done.")
    else:
        print("No movement method succeeded.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
