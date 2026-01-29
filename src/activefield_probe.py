import sys
from typing import Any, Optional

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover - runtime dependency
    win32com = None


def _dispatch_app() -> Any:
    if win32com is None:
        raise RuntimeError("pywin32 is not installed. Run: pip install pywin32")

    # EnsureDispatch can expose constants when available; fall back to Dispatch.
    try:
        return win32com.client.gencache.EnsureDispatch("QuickField.Application")
    except Exception:
        return win32com.client.Dispatch("QuickField.Application")


def _get_const(qf: Any, name: str) -> Optional[int]:
    try:
        consts = qf.Constants
    except Exception:
        return None

    for attr in (name, name.lower(), name.upper()):
        try:
            value = getattr(consts, attr)
            if isinstance(value, int):
                return value
        except Exception:
            continue
    return None


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


def _safe_get(obj: Any, attr: str) -> str:
    try:
        value = getattr(obj, attr)
        return str(value)
    except Exception:
        return "<unavailable>"


def _dump_labels(doc: Any, label_type: int, label_type_name: str) -> None:
    try:
        labels = doc.Labels(label_type)
        count = int(labels.Count)
    except Exception as exc:
        print(f"- Labels({label_type_name}) not accessible: {exc}")
        return

    print(f"- {label_type_name} labels: {count}")
    for i in range(1, count + 1):
        try:
            label = labels.Item(i)
        except Exception:
            try:
                label = labels(i)
            except Exception:
                break
        name = _safe_get(label, "Name")
        ltype = _safe_get(label, "Type")
        print(f"  [{i}] {name} (Type={ltype})")


def _dump_labels_by_guess(doc: Any) -> None:
    # Some versions don't expose qfBlock/qfEdge/qfVertex constants.
    # Try a small range of numeric IDs and show label.Type to infer mapping.
    for guess in range(0, 6):
        _dump_labels(doc, guess, f"type {guess}")


def main() -> int:
    print("QuickField ActiveField COM probe")
    try:
        qf = _dispatch_app()
    except Exception as exc:
        print(f"Failed to create QuickField.Application: {exc}")
        return 1

    print(f"- Version: {_safe_get(qf, 'Version')}")
    prb = _get_active_problem(qf)
    if prb is None:
        print("No active problem found.")
        print("Open a problem in QuickField, then re-run this script.")
        return 2

    print(f"- Problem Name: {_safe_get(prb, 'Name')}")
    print(f"- Problem FilePath: {_safe_get(prb, 'FilePath')}")
    print(f"- Problem Type: {_safe_get(prb, 'ProblemType')}")

    try:
        data_doc = prb.DataDoc
    except Exception as exc:
        print(f"Failed to access DataDoc: {exc}")
        return 3

    qf_block = _get_const(qf, "qfBlock")
    qf_edge = _get_const(qf, "qfEdge")
    qf_vertex = _get_const(qf, "qfVertex")

    if qf_block is None or qf_edge is None or qf_vertex is None:
        print("Could not read qfBlock/qfEdge/qfVertex constants. Trying numeric guesses 0..5.")
        _dump_labels_by_guess(data_doc)
    else:
        _dump_labels(data_doc, qf_block, "block")
        _dump_labels(data_doc, qf_edge, "edge")
        _dump_labels(data_doc, qf_vertex, "vertex")

    print("Probe done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
