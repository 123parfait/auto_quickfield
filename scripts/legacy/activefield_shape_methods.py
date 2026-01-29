import sys
from typing import Any, Optional

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except Exception:  # pragma: no cover - runtime dependency
    win32com = None
    pythoncom = None


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


def _get_first_shape(model: Any) -> Any:
    shapes = model.Shapes
    try:
        count = int(shapes.Count)
    except Exception:
        count = 0
    if count <= 0:
        raise RuntimeError("No shapes found in model.")
    try:
        return shapes.Item(1)
    except Exception:
        return shapes(1)


def _dump_methods(obj: Any) -> None:
    if pythoncom is None:
        raise RuntimeError("pythoncom not available; install pywin32")
    ti = obj._oleobj_.GetTypeInfo()
    attr = ti.GetTypeAttr()
    method_names = []
    for i in range(attr.cFuncs):
        fd = ti.GetFuncDesc(i)
        names = ti.GetNames(fd.memid)
        if names:
            method_names.append(names[0])
    method_names = sorted(set(method_names))
    print(f"Methods on Shape ({len(method_names)}):")
    for name in method_names:
        print(f"- {name}")


def main() -> int:
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

    shp = _get_first_shape(model)
    _dump_methods(shp)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
