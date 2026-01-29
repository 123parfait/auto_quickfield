from typing import Any, Optional

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except Exception:  # pragma: no cover - runtime dependency
    win32com = None
    pythoncom = None


VT_MAP = {
    pythoncom.VT_EMPTY: "VT_EMPTY",
    pythoncom.VT_NULL: "VT_NULL",
    pythoncom.VT_I2: "VT_I2",
    pythoncom.VT_I4: "VT_I4",
    pythoncom.VT_R4: "VT_R4",
    pythoncom.VT_R8: "VT_R8",
    pythoncom.VT_CY: "VT_CY",
    pythoncom.VT_DATE: "VT_DATE",
    pythoncom.VT_BSTR: "VT_BSTR",
    pythoncom.VT_DISPATCH: "VT_DISPATCH",
    pythoncom.VT_ERROR: "VT_ERROR",
    pythoncom.VT_BOOL: "VT_BOOL",
    pythoncom.VT_VARIANT: "VT_VARIANT",
    pythoncom.VT_UNKNOWN: "VT_UNKNOWN",
    pythoncom.VT_DECIMAL: "VT_DECIMAL",
    pythoncom.VT_I1: "VT_I1",
    pythoncom.VT_UI1: "VT_UI1",
    pythoncom.VT_UI2: "VT_UI2",
    pythoncom.VT_UI4: "VT_UI4",
    pythoncom.VT_I8: "VT_I8",
    pythoncom.VT_UI8: "VT_UI8",
    pythoncom.VT_INT: "VT_INT",
    pythoncom.VT_UINT: "VT_UINT",
    pythoncom.VT_ARRAY: "VT_ARRAY",
    pythoncom.VT_BYREF: "VT_BYREF",
}


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


def _get_first_in_collection(collection: Any, label: str) -> Optional[Any]:
    try:
        count = int(collection.Count)
    except Exception:
        count = 0
    if count <= 0:
        print(f"No items found in {label}.")
        return None
    try:
        return collection.Item(1)
    except Exception:
        try:
            return collection(1)
        except Exception:
            print(f"Failed to access first item in {label}.")
            return None


def _vt_name(vt: int) -> str:
    name = VT_MAP.get(vt, f"VT_{vt}")
    return name


def _param_count(fd: Any) -> int:
    try:
        return int(getattr(fd, "cParams"))
    except Exception:
        pass
    try:
        params = getattr(fd, "lprgelemdescParam")
        return len(params)
    except Exception:
        pass
    try:
        return int(getattr(fd, "cParamsOpt"))
    except Exception:
        return 0


def _dump_move_signature(shape: Any) -> None:
    if pythoncom is None:
        raise RuntimeError("pythoncom not available; install pywin32")
    ti = shape._oleobj_.GetTypeInfo()
    attr = ti.GetTypeAttr()
    found = False
    for i in range(attr.cFuncs):
        fd = ti.GetFuncDesc(i)
        names = ti.GetNames(fd.memid)
        if not names:
            continue
        if names[0] != "Move":
            continue
        found = True
        params = []
        pcount = _param_count(fd)
        for p in range(pcount):
            try:
                td = fd.lprgelemdescParam[p].tdesc
                vt = td.vt
                params.append(_vt_name(vt))
            except Exception:
                params.append("UNKNOWN")
        print(f"Move method: params={pcount} types={params}")
    if not found:
        print("Move method signature not found in type info.")


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

    shapes = model.Shapes

    # Collections (may have Move/Left/Right)
    for title, obj in [
        ("Shapes", shapes),
        ("Selection", getattr(model, "Selection", None)),
        ("Blocks", getattr(shapes, "Blocks", None)),
        ("Edges", getattr(shapes, "Edges", None)),
        ("Vertices", getattr(shapes, "Vertices", None)),
    ]:
        if obj is None:
            continue
        try:
            print(f"{title}.Move signature:")
            _dump_move_signature(obj)
        except Exception as exc:
            print(f"{title} signature lookup failed: {exc}")

    # Item types
    for title, collection in [
        ("Shape", shapes),
        ("Block", getattr(shapes, "Blocks", None)),
        ("Edge", getattr(shapes, "Edges", None)),
        ("Vertex", getattr(shapes, "Vertices", None)),
    ]:
        if collection is None:
            continue
        item = _get_first_in_collection(collection, title)
        if item is None:
            continue
        try:
            print(f"{title}.Move signature:")
            _dump_move_signature(item)
        except Exception as exc:
            print(f"{title} signature lookup failed: {exc}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
