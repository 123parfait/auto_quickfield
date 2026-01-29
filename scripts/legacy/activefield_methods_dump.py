import argparse
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


def _dump_methods(title: str, obj: Any) -> None:
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
    print(f"{title} methods ({len(method_names)}):")
    for name in method_names:
        print(f"- {name}")


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


def main() -> int:
    parser = argparse.ArgumentParser(description="Dump ActiveField methods")
    parser.add_argument("--label", default="", help="Label name to get a ShapeRange via LabeledAs")
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

    _dump_methods("Model", model)
    _dump_methods("Shapes", model.Shapes)
    try:
        selection = model.Selection
        _dump_methods("Selection", selection)
    except Exception as exc:
        print(f"Selection lookup failed: {exc}")

    # Collections: Blocks / Edges / Vertices
    try:
        blocks = model.Shapes.Blocks
        _dump_methods("Blocks", blocks)
        blk = blocks.Item(1)
        _dump_methods("Block", blk)
    except Exception as exc:
        print(f"Blocks lookup failed: {exc}")

    try:
        edges = model.Shapes.Edges
        _dump_methods("Edges", edges)
        edge = edges.Item(1)
        _dump_methods("Edge", edge)
    except Exception as exc:
        print(f"Edges lookup failed: {exc}")

    try:
        vertices = model.Shapes.Vertices
        _dump_methods("Vertices", vertices)
        vtx = vertices.Item(1)
        _dump_methods("Vertex", vtx)
    except Exception as exc:
        print(f"Vertices lookup failed: {exc}")

    try:
        shp = _find_shape_by_label(model, args.label) if args.label else _iter_all_shapes(model.Shapes).__next__()
        _dump_methods("Shape", shp)
    except Exception as exc:
        print(f"Shape lookup failed: {exc}")

    if args.label:
        try:
            shaperange = model.Shapes.LabeledAs(args.label)
            _dump_methods("ShapeRange", shaperange)
        except Exception as exc:
            print(f"ShapeRange lookup failed: {exc}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
