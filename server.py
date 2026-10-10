"""SOLIDWORKS MCP server.

Lets AI agents (Claude Code etc.) work *together with a user* in a live SOLIDWORKS session
through the COM API. Typical workflow: 3D model first, then a drawing linked to the model.

Design rules for co-working with a human:
  * never interrupt the user: changes wait until no modal dialog / PropertyManager command /
    sketch editing is active (reads are allowed any time);
  * every agent change is one undo step (Undo in SOLIDWORKS reverts it in one click);
  * the user's selection is readable, and the agent can hand a selection back to the user;
  * all COM calls run on one dedicated thread, strictly one at a time;
  * nothing is saved, closed or overwritten unless the user asked.
"""

from __future__ import annotations

import asyncio
import base64
import contextlib
import ctypes
import io
import math
import os
import sys
import tempfile
import time
import traceback
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any, Literal

import pythoncom
import pywintypes
import win32com.client
from mcp.server.fastmcp import FastMCP, Image

import swcom
from swcom import SW

PROGID = os.environ.get("SOLIDWORKS_MCP_PROGID", "SldWorks.Application")
IDLE_TIMEOUT = float(os.environ.get("SOLIDWORKS_MCP_IDLE_TIMEOUT", "30"))
WORK_DIR = Path(os.environ.get("SOLIDWORKS_MCP_WORK_DIR", Path(tempfile.gettempdir()) / "solidworks-mcp"))
WORK_DIR.mkdir(parents=True, exist_ok=True)
# type library major version (33 = SOLIDWORKS 2025); default: the newest registered one
TLB_VERSION = int(os.environ["SOLIDWORKS_MCP_TLB_VERSION"]) if os.environ.get("SOLIDWORKS_MCP_TLB_VERSION") else None

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
BUSY_HRESULTS = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}
DEAD_HRESULTS = {-2147023174, -2147023170, -2147417848, -2147220995}

DOC_TYPES = {1: "part", 2: "assembly", 3: "drawing"}
LENGTH_UNITS = {0: "mm", 1: "cm", 2: "m", 3: "in", 4: "ft", 5: "ft-in", 6: "angstrom",
                7: "nm", 8: "micron", 9: "mil", 10: "uin"}
# swStandardViews_e == index in ModelDoc2.GetModelViewNames() (names are localized)
STD_VIEWS = {"normal": 0, "front": 1, "back": 2, "left": 3, "right": 4, "top": 5,
             "bottom": 6, "iso": 7, "trimetric": 8, "dimetric": 9}
PAPER = {"A4": 7, "A4L": 6, "A3": 8, "A2": 9, "A1": 10, "A0": 11}  # A4 = portrait (ГОСТ)
PAPER_MM = {"A4": (210, 297), "A3": (420, 297), "A2": (594, 420), "A1": (841, 594), "A0": (1189, 841)}

with contextlib.suppress(Exception):
    ctypes.windll.shcore.SetProcessDpiAwareness(2)


def log(*args: Any) -> None:
    print("[solidworks-mcp]", *args, file=sys.stderr, flush=True)


class SolidWorksBusy(Exception):
    pass


# ---------------------------------------------------------------------------
# COM worker thread
# ---------------------------------------------------------------------------

def _init_worker() -> None:
    pythoncom.CoInitialize()
    with contextlib.redirect_stdout(sys.stderr):  # makepy prints progress; stdout is the MCP channel
        swcom.init(WORK_DIR, TLB_VERSION)


_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="sw-com", initializer=_init_worker)
_app: SW | None = None
_hwnd: int | None = None
c: Any = None  # swconst constants, set on the worker


def _hresult(e: Exception) -> int | None:
    return e.hresult if isinstance(e, pywintypes.com_error) else None


def _connect(launch: bool = False) -> SW:
    global _app, _hwnd, c
    c = swcom.constants
    if _app is not None:
        try:
            _app.RevisionNumber()
            return _app
        except pywintypes.com_error as e:
            if _hresult(e) in BUSY_HRESULTS:
                return _app
            _app, _hwnd = None, None
    try:
        _app = swcom.get_active(PROGID)
    except pywintypes.com_error:
        if not launch:
            raise RuntimeError("SOLIDWORKS is not running. Ask the user to start it, "
                               "or call start_solidworks.") from None
        log("launching SOLIDWORKS...")
        _app = swcom.launch(PROGID)
    _hwnd = None
    return _app


def _main_hwnd() -> int:
    """SOLIDWORKS main frame window (found by process id - COM may be busy)."""
    global _hwnd
    import win32gui
    import win32process

    if _hwnd and win32gui.IsWindow(_hwnd):
        return _hwnd
    pid = int(_connect().GetProcessID())
    found: list[int] = []

    def cb(h, _):
        if win32process.GetWindowThreadProcessId(h)[1] == pid and win32gui.IsWindowVisible(h) \
                and win32gui.GetWindowText(h).startswith("SOLIDWORKS") and not win32gui.GetParent(h):
            found.append(h)
    win32gui.EnumWindows(cb, None)
    if not found:
        raise RuntimeError("SOLIDWORKS main window not found")
    _hwnd = found[0]
    return _hwnd


def ui_state() -> dict:
    """What the user is doing right now in SOLIDWORKS."""
    import win32gui

    app = _connect()
    st: dict[str, Any] = {"modal_dialog": False, "command": None, "editing_sketch": None}
    with contextlib.suppress(Exception):
        st["modal_dialog"] = not win32gui.IsWindowEnabled(_main_hwnd())
    if st["modal_dialog"]:
        return st
    with contextlib.suppress(Exception):
        cmd_id, title, ui_active = app.GetRunningCommandInfo()
        if cmd_id not in (-3, -1, 0) or ui_active:
            st["command"] = {"id": cmd_id, "title": title, "property_manager_open": bool(ui_active)}
    with contextlib.suppress(Exception):
        m = app.ActiveDoc
        # a drawing always reports its sheet sketch as active - only parts/assemblies count
        if m is not None and m.GetType() != 3:
            sk = m.SketchManager.ActiveSketch
            if sk is not None:
                st["editing_sketch"] = sk.Name
    return st


def _busy_reason(st: dict) -> str | None:
    if st["modal_dialog"]:
        return "a modal dialog is open"
    if st["command"]:
        return f"a command is running ({st['command']['title'] or st['command']['id']})"
    if st["editing_sketch"]:
        return f"sketch '{st['editing_sketch']}' is being edited"
    return None


def wait_idle(timeout: float = IDLE_TIMEOUT) -> None:
    """Block until the user has no dialog / command / sketch edit active."""
    deadline = time.monotonic() + timeout
    while True:
        reason = _busy_reason(ui_state())
        if not reason:
            return
        if time.monotonic() > deadline:
            raise SolidWorksBusy(
                f"SOLIDWORKS stayed busy for {timeout:.0f}s: {reason} (probably the user is working). "
                "Do NOT interrupt - tell the user and retry when they finish.")
        time.sleep(0.3)


def _format_com_error(e: pywintypes.com_error) -> str:
    hr, msg, exc, _ = e.args
    detail = f": {exc[2]}" if exc and len(exc) > 2 and exc[2] else ""
    if hr in BUSY_HRESULTS:
        return "SOLIDWORKS rejected the call because it is busy (dialog or long operation). Retry later."
    return f"COM error {hr} ({msg}){detail}"


def _call_on_worker(fn, *args, need_idle: bool = False, retries_for: float = 15.0, **kwargs):
    global _app
    deadline = time.monotonic() + retries_for
    reconnected = False
    while True:
        try:
            _connect()
            if need_idle:
                wait_idle()
            return fn(*args, **kwargs)
        except pywintypes.com_error as e:
            hr = _hresult(e)
            if hr in BUSY_HRESULTS and time.monotonic() < deadline:
                time.sleep(0.3)
                continue
            if hr in DEAD_HRESULTS and not reconnected:
                _app, reconnected = None, True
                continue
            raise RuntimeError(_format_com_error(e)) from None


def _worker_call(fn, args, kwargs):
    res = _call_on_worker(fn, *args, **kwargs)
    # Lazy COM values (swcom._PropOrMethod) must be resolved here, on the COM thread,
    # before the result is serialized on the event-loop thread.
    return _clean(res) if isinstance(res, (dict, list, tuple)) else res


async def run(fn, *args, **kwargs):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(_executor, lambda: _worker_call(fn, args, kwargs))


# ---------------------------------------------------------------------------
# Helpers (also exposed inside execute_python)
# ---------------------------------------------------------------------------

def model() -> SW:
    m = _connect().ActiveDoc
    if m is None:
        raise RuntimeError("No document is open in SOLIDWORKS.")
    return m


def find_doc(name: str) -> SW | None:
    """Open document by title (with or without extension) or full path."""
    key = name.lower()
    for d in _connect().GetDocuments() or ():
        title, path = d.GetTitle(), d.GetPathName() or ""
        if key in (title.lower(), path.lower(), Path(path).name.lower(), Path(path).stem.lower()):
            return d
    return None


def doc_type(m: SW) -> str:
    return DOC_TYPES.get(m.GetType(), "unknown")


def need(m: SW, *kinds: str) -> None:
    if doc_type(m) not in kinds:
        raise RuntimeError(f"The active document is a {doc_type(m)}; this needs a {' or '.join(kinds)}. "
                           "Activate the right document first (documents action=activate).")


def mm(v: float) -> float:
    """millimetres -> metres (the API always works in metres)."""
    return float(v) / 1000.0


def to_mm(v: float) -> float:
    return round(float(v) * 1000.0, 6)


def deg(v: float) -> float:
    """degrees -> radians."""
    return math.radians(float(v))


def doubles(values) -> win32com.client.VARIANT:
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(v) for v in values])


def dispatch_array(objs) -> win32com.client.VARIANT:
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH,
                                   [swcom.unwrap(o) for o in objs])


def _clean(v: Any, nd: int = 6) -> Any:
    if isinstance(v, swcom._PropOrMethod):
        v = v()
    if isinstance(v, SW):
        return "<COM object>"
    if isinstance(v, float):
        return round(v, nd)
    if isinstance(v, (list, tuple)):
        return [_clean(i, nd) for i in v]
    if isinstance(v, dict):
        return {k: _clean(i, nd) for k, i in v.items()}
    if isinstance(v, (str, int, bool)) or v is None:
        return v
    return str(v)


def win_path(p: str) -> str:
    """SOLIDWORKS only understands backslash paths."""
    return str(Path(p))


def features(m: SW | None = None):
    """Iterate top-level features in tree order."""
    f = (m or model()).FirstFeature()
    while f is not None:
        yield f
        f = f.GetNextFeature()


def subfeatures(f: SW):
    s = f.GetFirstSubFeature()
    while s is not None:
        yield s
        s = s.GetNextSubFeature()


def std_planes(m: SW | None = None) -> dict[str, SW]:
    """Front/Top/Right planes - the first three planes in the tree (names are localized:
    Спереди / Сверху / Справа)."""
    planes = [f for f in features(m) if f.GetTypeName2() == "RefPlane"][:3]
    return dict(zip(("front", "top", "right"), planes))


def model_view_name(m: SW, view: str) -> str:
    """Localized named view ('*Изометрия') for a standard view key (front, top, iso...)."""
    names = list(m.GetModelViewNames() or ())
    i = STD_VIEWS[view]
    if i >= len(names):
        raise ValueError(f"Model has no standard view #{i}: {names}")
    return names[i]


def persist_id(m: SW, obj: SW) -> str | None:
    with contextlib.suppress(Exception):
        ref = m.Extension.GetPersistReference3(obj)
        if ref:
            return base64.b64encode(bytes(ref) if not isinstance(ref, bytes) else ref).decode()
    return None


def from_persist_id(m: SW, pid: str) -> SW:
    res = m.Extension.GetObjectByPersistReference3(base64.b64decode(pid), None)
    obj = res[0] if isinstance(res, tuple) else res
    if obj is None:
        raise RuntimeError(f"Object id {pid[:12]}... not found in the active document (deleted or rebuilt?)")
    return obj


@contextlib.contextmanager
def undo_step(m: SW, name: str):
    """Everything inside becomes one Undo step for the user."""
    ext = m.Extension
    started = False
    with contextlib.suppress(Exception):
        ext.StartRecordingUndoObject()
        started = True
    try:
        yield
    finally:
        if started:
            with contextlib.suppress(Exception):
                ext.FinishRecordingUndoObject2(f"AI: {name}", False)


def feature_errors(m: SW) -> list[dict]:
    out = []
    for f in features(m):
        with contextlib.suppress(Exception):
            code, warning = f.GetErrorCode2()
            if code:
                out.append({"feature": f.Name, "type": f.GetTypeName2(), "code": code, "warning": bool(warning)})
    return out


def doc_summary(m: SW) -> dict:
    out = {"title": m.GetTitle(), "path": m.GetPathName() or None, "type": doc_type(m),
           "modified": bool(m.GetSaveFlag())}
    with contextlib.suppress(Exception):
        out["read_only"] = bool(m.IsOpenedReadOnly())
    with contextlib.suppress(Exception):
        if out["type"] != "drawing":
            out["configuration"] = m.ConfigurationManager.ActiveConfiguration.Name
    return out


def _sel_type_names() -> dict[int, str]:
    out = {}
    for k, v in vars(c).items():
        if k.startswith("swSel") and isinstance(v, int) and "Option" not in k and k[5:6].isupper():
            out.setdefault(v, k[5:])
    return out


def _entity_info(m: SW, obj: SW, t: int, detailed: bool) -> dict:
    info: dict[str, Any] = {}

    def attempt(key, fn):
        with contextlib.suppress(Exception):
            v = fn()
            if v is not None:
                info[key] = _clean(v)

    if t == c.swSelFACES:
        attempt("area_mm2", lambda: obj.GetArea() * 1e6)
        attempt("feature", lambda: obj.GetFeature().Name)
        s = obj.GetSurface()
        if s.IsPlane():
            attempt("surface", lambda: "plane")
            attempt("normal", lambda: list(obj.Normal))
        elif s.IsCylinder():
            info["surface"] = "cylinder"
            attempt("radius_mm", lambda: to_mm(s.CylinderParams[6]))
        if detailed:
            attempt("box_mm", lambda: [to_mm(v) for v in obj.GetBox()])
    elif t == c.swSelEDGES:
        crv = obj.GetCurve()
        p = obj.GetCurveParams3()
        attempt("length_mm", lambda: to_mm(crv.GetLength3(p.UMinValue, p.UMaxValue)))
        if crv.IsCircle():
            info["curve"] = "circle"
            attempt("radius_mm", lambda: to_mm(crv.CircleParams[6]))
        elif crv.IsLine():
            info["curve"] = "line"
    elif t == c.swSelVERTICES:
        attempt("point_mm", lambda: [to_mm(v) for v in obj.GetPoint()])
    elif t == c.swSelCOMPONENTS:
        attempt("name", lambda: obj.Name2)
        attempt("path", lambda: obj.GetPathName())
        attempt("configuration", lambda: obj.ReferencedConfiguration)
        attempt("suppressed", lambda: bool(obj.IsSuppressed()))
    elif t == c.swSelDIMENSIONS:
        d = obj.GetDimension2(0)
        attempt("name", lambda: d.FullName)
        attempt("value", lambda: _dim_value(d))
    elif t == c.swSelDRAWINGVIEWS:
        attempt("name", lambda: obj.Name)
        attempt("model", lambda: obj.GetReferencedModelName())
        attempt("configuration", lambda: obj.ReferencedConfiguration)
        attempt("scale", lambda: obj.ScaleDecimal)
    elif t == c.swSelSKETCHSEGS:
        attempt("length_mm", lambda: to_mm(obj.GetLength()))
        attempt("sketch", lambda: obj.GetSketch().Name)
    elif t == c.swSelSKETCHPOINTS:
        attempt("point_mm", lambda: [to_mm(obj.X), to_mm(obj.Y), to_mm(obj.Z)])
    if "name" not in info and obj.has("GetTypeName2"):  # any feature (sketch, plane, extrude...)
        attempt("name", lambda: obj.Name)
        attempt("feature_type", lambda: obj.GetTypeName2())
    return _clean(info)


def _dim_value(d: SW) -> float:
    v = d.SystemValue
    if d.GetType() == c.swDimensionParamTypeDoubleAngular:
        return round(math.degrees(v), 6)
    return to_mm(v)


def _is_angular(d: SW) -> bool:
    with contextlib.suppress(Exception):
        return d.GetType() == c.swDimensionParamTypeDoubleAngular
    return False


def iter_dimensions(m: SW, feature_names: list[str] | None = None):
    """(feature, display_dimension, dimension) for features' dimensions (sketch dims included)."""
    wanted = set(feature_names or [])

    def from_feature(f):
        dd = f.GetFirstDisplayDimension()
        while dd is not None:
            yield f, dd, dd.GetDimension2(0)
            dd = f.GetNextDisplayDimension(dd)

    for f in features(m):
        if wanted and f.Name not in wanted:
            # sketches of a wanted feature are its sub-features
            continue
        yield from from_feature(f)
        if wanted:
            for s in subfeatures(f):
                yield from from_feature(s)


# ---------------------------------------------------------------------------
# MCP server
# ---------------------------------------------------------------------------

mcp = FastMCP(
    "solidworks",
    instructions="""Controls the user's live SOLIDWORKS (Russian UI) via COM.
You share it with a human who may be modeling at the same time.
- Start with `status` (open documents, active one, units, what the user is doing, selection).
- Changes wait until the user has no dialog / PropertyManager command / sketch edit open.
  If a tool reports SOLIDWORKS is busy: tell the user and wait, never interrupt them.
- "This"/"selected" -> `get_selection`. To point at things for the user -> `select_for_user`.
- Each tool call that changes the model is one Undo step for the user.
- Tool parameters use millimetres and degrees. The raw COM API (execute_python) uses
  metres and radians - convert with mm()/deg().
- Feature, plane and view names are localized (Спереди/Сверху/Справа, *Изометрия...):
  use std_planes() / model_view_name() helpers instead of English names.
- Workflow is usually: model (part/assembly) -> drawing linked to the model. Drawings show
  model dimensions via insert_model_annotations. GOST title block fields come from the
  model's configuration-specific custom_properties (configuration="active"): Обозначение,
  Наименование, п_Разраб, п_Пров, Лит...; Масса/Материал come from the model itself.
- Verify visually with `screenshot`. Don't save, close or overwrite files unless the user asked.""",
)


@mcp.tool()
async def status() -> dict:
    """SOLIDWORKS state: version, open documents, the active document (type, units, configuration,
    material), what the user is doing (dialog/command/sketch edit) and the selection size."""
    def _status():
        app = _connect()
        out: dict[str, Any] = {"connected": True, "version": app.RevisionNumber(),
                               "api_typelib_version": swcom.tlb_version}
        st = ui_state()
        out["user_activity"] = st
        out["busy"] = _busy_reason(st)
        out["documents"] = [doc_summary(d) for d in (app.GetDocuments() or ())
                            if d.Visible] if not st["modal_dialog"] else "unavailable while a dialog is open"
        m = app.ActiveDoc
        if m is not None and not st["modal_dialog"]:
            a = doc_summary(m)
            a["length_unit"] = LENGTH_UNITS.get(m.GetUserPreferenceIntegerValue(c.swUnitsLinear), "?")
            with contextlib.suppress(Exception):
                a["selection_count"] = m.SelectionManager.GetSelectedObjectCount2(-1)
            if a["type"] == "part":
                with contextlib.suppress(Exception):
                    a["material"] = m.GetMaterialPropertyName2("")[0] or None
            if a["type"] == "drawing":
                with contextlib.suppress(Exception):
                    a["sheets"] = list(m.GetSheetNames())
                    a["current_sheet"] = m.GetCurrentSheet().GetName()
            out["active"] = a
        return out
    try:
        return await run(_status)
    except RuntimeError as e:
        return {"connected": False, "error": str(e)}


@mcp.tool()
async def start_solidworks(visible: bool = True) -> dict:
    """Start SOLIDWORKS if it is not running (can take 1-2 minutes) and connect to it."""
    def _start():
        app = _connect(launch=True)
        deadline = time.monotonic() + 240
        while True:
            try:
                app.Visible = visible
                return {"version": app.RevisionNumber()}
            except pywintypes.com_error as e:
                if _hresult(e) in BUSY_HRESULTS and time.monotonic() < deadline:
                    time.sleep(1)
                    continue
                raise
    return await run(_start, retries_for=240)


def _default_template(kind: str) -> str:
    app = _connect()
    pref = {"part": c.swDefaultTemplatePart, "assembly": c.swDefaultTemplateAssembly,
            "drawing": c.swDefaultTemplateDrawing}[kind]
    t = app.GetUserPreferenceStringValue(pref)
    if t and Path(t).is_file():
        return t
    ext = {"part": ".prtdot", "assembly": ".asmdot", "drawing": ".drwdot"}[kind]
    found = []
    for d in (app.GetUserPreferenceStringValue(c.swFileLocationsDocumentTemplates) or "").split(";"):
        if d and Path(d).is_dir():
            found += sorted(p for p in Path(d).glob(f"*{ext}"))
    if not found:
        raise RuntimeError(f"No {kind} template found; pass template=<path>")
    gost = [p for p in found if "gost" in p.name.lower()]
    if kind == "drawing":  # prefer the part-drawing template ("gost-part drw")
        gost = sorted(gost, key=lambda p: "part" not in p.name.lower())
    return str((gost or found)[0])


@mcp.tool()
async def documents(action: Literal["list", "new", "open", "activate", "save", "save_as", "close"],
                    kind: Literal["part", "assembly", "drawing"] = "part",
                    path: str | None = None, name: str | None = None,
                    template: str | None = None, paper: str | None = None,
                    read_only: bool = False, configuration: str | None = None) -> dict:
    """Documents. list; new (kind, optional template; drawing: paper A4/A4L/A3/A2/A1/A0 -
    default templates are the GOST ones); open (path, read_only, configuration); activate
    (name = title of an open document); save (active document); save_as (path - renames the
    active document; for exports to other formats use `export`); close (name; only documents
    the agent created or the user asked to close - unsaved changes are discarded!).
    Only save/close when the user asked to."""
    def _docs():
        app = _connect()
        if action == "list":
            return {"documents": [doc_summary(d) for d in (app.GetDocuments() or ()) if d.Visible]}
        if action in ("new", "open", "save", "save_as", "close"):
            wait_idle()
        if action == "new":
            tpl = win_path(template) if template else _default_template(kind)
            psize = PAPER.get((paper or "A3").upper(), 8) if kind == "drawing" else 0
            m = app.NewDocument(tpl, psize, 0.0, 0.0)
            if m is None:
                raise RuntimeError(f"NewDocument failed for template {tpl}")
            return {"created": doc_summary(m), "template": tpl}
        if action == "open":
            p = win_path(path)
            t = {".sldprt": 1, ".sldasm": 2, ".slddrw": 3}.get(Path(p).suffix.lower())
            if t is None:
                raise ValueError("Use open for .SLDPRT/.SLDASM/.SLDDRW (import other formats via execute_python LoadFile4)")
            opts = c.swOpenDocOptions_Silent | (c.swOpenDocOptions_ReadOnly if read_only else 0)
            m, err, warn = app.OpenDoc6(p, t, opts, configuration or "", 0, 0)
            if m is None:
                raise RuntimeError(f"OpenDoc6 failed: errors={err} warnings={warn}")
            app.ActivateDoc3(m.GetTitle(), False, 0, 0)
            return {"opened": doc_summary(m), "warnings": warn}
        if action == "activate":
            target = find_doc(name or "")
            if target is None:
                raise RuntimeError(f"No open document titled {name!r}")
            app.ActivateDoc3(target.GetTitle(), False, 0, 0)
            return {"active": doc_summary(target)}
        if action == "save":
            m = model()
            ok, err, warn = m.Save3(c.swSaveAsOptions_Silent, 0, 0)
            return {"ok": bool(ok), "errors": err, "warnings": warn, "path": m.GetPathName()}
        if action == "save_as":
            m = model()
            ok, err, warn = m.Extension.SaveAs3(win_path(path), 0, c.swSaveAsOptions_Silent, None, None, 0, 0)
            return {"ok": bool(ok), "errors": err, "warnings": warn, "document": doc_summary(m)}
        # close
        target = find_doc(name) if name else model()
        if target is None:
            raise RuntimeError(f"No open document {name!r}")
        title = target.GetTitle()
        app.CloseDoc(title)
        return {"closed": title}
    return await run(_docs)


def _component_tree(comp: SW, depth: int, max_depth: int) -> dict:
    node: dict[str, Any] = {"name": comp.Name2}
    with contextlib.suppress(Exception):
        node["path"] = comp.GetPathName()
        node["configuration"] = comp.ReferencedConfiguration
        node["suppressed"] = bool(comp.IsSuppressed())
        if comp.IsFixed():
            node["fixed"] = True
    if depth < max_depth:
        kids = comp.GetChildren() or ()
        if kids:
            node["children"] = [_component_tree(k, depth + 1, max_depth) for k in kids]
    return node


_HIDDEN_TYPES = {"FavoriteFolder", "HistoryFolder", "SelectionSetFolder", "DocsFolder",
                 "DetailCabinet", "InkMarkupFolder", "EnvFolder", "SensorFolder", "CommentsFolder",
                 "ConfigTableFolder", "SurfaceBodyFolder", "SolidBodyFolder", "MaterialFolder",
                 "EqnFolder", "BlockFolder", "LiveSectionFolder", "HistoryFolder", "MateReferenceGroupFolder",
                 "AnnotationsFolder", "DrawingViewFolder"}


@mcp.tool()
async def feature_tree(include_folders: bool = False, include_children: bool = True,
                       components: bool = True, max_depth: int = 3) -> dict:
    """FeatureManager tree of the active part/assembly: features in order with type,
    suppression, errors and their sketches; for assemblies also the component tree.
    Names are what the user sees (Russian)."""
    def _tree():
        m = model()
        need(m, "part", "assembly")
        out_feats = []
        for f in features(m):
            t = f.GetTypeName2()
            if not include_folders and t in _HIDDEN_TYPES:
                continue
            item: dict[str, Any] = {"name": f.Name, "type": t}
            with contextlib.suppress(Exception):
                if f.IsSuppressed2(c.swThisConfiguration, None)[0]:
                    item["suppressed"] = True
            with contextlib.suppress(Exception):
                code, warn = f.GetErrorCode2()
                if code:
                    item["error"] = {"code": code, "warning": bool(warn)}
            if include_children:
                kids = [{"name": s.Name, "type": s.GetTypeName2()} for s in subfeatures(f)]
                if kids:
                    item["children"] = kids
            out_feats.append(item)
        out: dict[str, Any] = {"document": doc_summary(m), "features": out_feats}
        if components and doc_type(m) == "assembly":
            root = m.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
            out["components"] = [_component_tree(k, 1, max_depth) for k in (root.GetChildren() or ())]
        return out
    return await run(_tree)


@mcp.tool()
async def get_selection(detailed: bool = False) -> dict:
    """What the user has selected in SOLIDWORKS (faces, edges, features, components, dimensions,
    drawing views...), with useful measurements in mm. Each item has an `id` usable with
    select_for_user / from_persist_id() in execute_python."""
    def _sel():
        m = model()
        sm = m.SelectionManager
        names = _sel_type_names()
        items = []
        for i in range(1, sm.GetSelectedObjectCount2(-1) + 1):
            t = sm.GetSelectedObjectType3(i, -1)
            item: dict[str, Any] = {"index": i, "type": names.get(t, str(t))}
            try:
                obj = sm.GetSelectedObject6(i, -1)
                if obj is not None:
                    item.update(_entity_info(m, obj, t, detailed))
                    pid = persist_id(m, obj)
                    if pid:
                        item["id"] = pid
                with contextlib.suppress(Exception):
                    comp = sm.GetSelectedObjectsComponent4(i, -1)
                    if comp is not None and t != c.swSelCOMPONENTS:
                        item["component"] = comp.Name2
                if detailed:
                    with contextlib.suppress(Exception):
                        item["pick_point_mm"] = [to_mm(v) for v in sm.GetSelectionPoint2(i, -1)]
            except Exception as ex:
                item["error"] = str(ex)
            items.append(item)
        return {"document": m.GetTitle(), "count": len(items), "items": items}
    return await run(_sel)


@mcp.tool()
async def select_for_user(items: list[dict] | None = None, clear_first: bool = True,
                          zoom: bool = False) -> dict:
    """Select objects for the user to see (highlight). items: [{"feature": "Бобышка-Вытянуть1"},
    {"component": "Вал-1"}, {"dimension": "D1@Эскиз1"}, {"view": "Чертежный вид1"},
    {"sheet": "Лист1"}, {"id": "<id from get_selection>"}]. Empty items + clear_first clears
    the selection. zoom=True zooms the 3D view to the selection."""
    def _select():
        m = model()
        ext = m.Extension
        if clear_first:
            m.ClearSelection2(True)
        done, errors = [], []
        for it in items or []:
            try:
                ok = False
                if "feature" in it:
                    f = m.FeatureByName(it["feature"])
                    ok = bool(f and f.Select2(True, 0))
                elif "component" in it:
                    comp = m.GetComponentByName(it["component"])
                    ok = bool(comp and comp.Select4(True, None, False))
                elif "id" in it:
                    obj = from_persist_id(m, it["id"])
                    for meth, args in (("Select4", (True, None)), ("Select2", (True, 0)), ("Select", (True,))):
                        if obj.has(meth):
                            ok = bool(getattr(obj, meth)(*args))
                            break
                else:
                    kind, sw_type = next(((k, t) for k, t in (("dimension", "DIMENSION"), ("view", "DRAWINGVIEW"),
                                                              ("sheet", "SHEET"), ("sketch", "SKETCH"))
                                          if k in it), (None, None))
                    if kind is None:
                        raise ValueError(f"Unknown item {it}")
                    ok = bool(ext.SelectByID2(it[kind], sw_type, 0, 0, 0, True, 0, None, 0))
                (done if ok else errors).append(it if ok else {"item": it, "error": "not found / not selectable"})
            except Exception as ex:
                errors.append({"item": it, "error": str(ex)})
        if zoom and done and doc_type(m) != "drawing":
            m.ViewZoomToSelection()
        m.GraphicsRedraw2()
        return {"selected": len(done), "errors": errors}
    return await run(_select)


@mcp.tool()
async def dimensions(action: Literal["list", "set"] = "list", features_filter: list[str] | None = None,
                     values: dict[str, float] | None = None,
                     configurations: Literal["this", "all"] = "this", rebuild: bool = True) -> dict:
    """Model dimensions (the driving parameters). list: all dimensions (or only of the given
    feature names, incl. their sketches) with full name, value (mm or degrees), whether it is
    marked for drawing (those go to the drawing via insert_model_annotations).
    set: values {"D1@Эскиз1": 120, "D2@Бобышка-Вытянуть1": 15} in mm/deg, then rebuild.
    One undo step."""
    def _dims():
        m = model()
        need(m, "part", "assembly")
        if action == "list":
            out = []
            for f, dd, d in iter_dimensions(m, features_filter):
                item = {"name": d.FullName.rsplit("@", 1)[0] if d.FullName.count("@") > 1 else d.FullName,
                        "feature": f.Name, "value": _dim_value(d),
                        "unit": "deg" if _is_angular(d) else "mm"}
                with contextlib.suppress(Exception):
                    item["marked_for_drawing"] = bool(dd.MarkedForDrawing)
                with contextlib.suppress(Exception):
                    if d.DrivenState == c.swDimensionDriven:
                        item["driven"] = True
                out.append(item)
            return {"document": m.GetTitle(), "count": len(out), "dimensions": out}
        wait_idle()
        which = c.swSetValue_InThisConfiguration if configurations == "this" else c.swSetValue_InAllConfigurations
        results, errors = {}, []
        with undo_step(m, "размеры"):
            for name, val in (values or {}).items():
                d = m.Parameter(name)
                if d is None:
                    errors.append({"name": name, "error": "dimension not found (use the full name from list)"})
                    continue
                si = deg(val) if _is_angular(d) else mm(val)
                rc = d.SetSystemValue3(si, which, None)
                if rc != 0:
                    errors.append({"name": name, "error": f"SetSystemValue3 returned {rc}"})
                else:
                    results[name] = val
            if rebuild and results:
                m.EditRebuild3()
        return {"set": results, "errors": errors, "rebuild_errors": feature_errors(m) if rebuild else None}
    return await run(_dims)


@mcp.tool()
async def equations(action: Literal["list", "set", "add", "delete"] = "list",
                    index: int | None = None, equation: str | None = None) -> dict:
    """Global variables and equations (Уравнения). list: index, text, value, is global variable.
    add: equation like '"Длина" = 120' or '"D1@Эскиз1" = "Длина" / 2'. set: replace equation
    at index. delete: index. Rebuilds after a change. One undo step."""
    def _eq():
        m = model()
        em = m.GetEquationMgr()
        if action == "list":
            out = []
            for i in range(em.GetCount()):
                out.append({"index": i, "equation": em.Equation(i), "value": _clean(em.Value(i)),
                            "global_variable": bool(em.GlobalVariable(i))})
            return {"equations": out}
        wait_idle()
        with undo_step(m, "уравнения"):
            if action == "add":
                if not equation:
                    raise ValueError("equation is required")
                rc = em.Add2(-1, equation, True)  # Add3 returns -1 on SW2025 - Add2 works
                if rc < 0:
                    raise RuntimeError(f"SOLIDWORKS rejected the equation {equation!r} (syntax? unknown dimension?)")
            elif action == "set":
                if index is None or not equation:
                    raise ValueError("index and equation are required")
                em.put("Equation", index, equation)
                rc = index
            else:
                if index is None:
                    raise ValueError("index is required")
                rc = em.Delete(index)
                if rc < 0:
                    raise RuntimeError(f"No equation with index {index}")
            em.EvaluateAll()
            m.EditRebuild3()
        return {"index": rc, "equations": [{"index": i, "equation": em.Equation(i), "value": em.Value(i)}
                                           for i in range(em.GetCount())],
                "rebuild_errors": feature_errors(m)}
    return await run(_eq)


@mcp.tool()
async def custom_properties(action: Literal["get", "set", "delete"] = "get",
                            properties: dict[str, str] | None = None, names: list[str] | None = None,
                            configuration: str = "") -> dict:
    """Custom properties of the active document (Свойства): Обозначение, Наименование,
    п_Разраб, п_Пров, Лит, Материал... configuration: "" = file-level (Общие), "active" = the
    active configuration, or a configuration name.
    IMPORTANT: the GOST drawing title block reads the MODEL's configuration-specific
    properties ($PRPSHEET) - so for a part/assembly set Обозначение/Наименование/п_Разраб...
    with configuration="active" (file-level ones do not show up in the title block).
    get: all (value + evaluated value). set: {name: value} (values may be links like
    '"SW-Масса@Деталь1.SLDPRT"'). delete: names. One undo step."""
    def _props():
        m = model()
        cfg = configuration
        if cfg == "active":
            cfg = "" if doc_type(m) == "drawing" else m.ConfigurationManager.ActiveConfiguration.Name
        cpm = m.Extension.CustomPropertyManager(cfg)
        if action == "get":
            out = {}
            for n in cpm.GetNames() or ():
                _, val, resolved, _, linked = cpm.Get6(n, False)
                out[n] = {"value": val, "evaluated": resolved} if val != resolved else {"value": val}
            return {"document": m.GetTitle(), "configuration": cfg or None, "properties": out}
        wait_idle()
        res = {}
        with undo_step(m, "свойства"):
            if action == "set":
                for n, v in (properties or {}).items():
                    res[n] = cpm.Add3(n, c.swCustomInfoText, str(v), c.swCustomPropertyReplaceValue)
            else:
                for n in names or []:
                    res[n] = cpm.Delete2(n)
        return {"results": res, "note": "0 = OK"}
    return await run(_props)


@mcp.tool()
async def mass_properties(component: str | None = None) -> dict:
    """Mass, volume, surface area, centre of mass (mm) of the active part/assembly
    (or of one assembly component by name). Uses the assigned material density."""
    def _mass():
        m = model()
        need(m, "part", "assembly")
        mp = m.Extension.CreateMassProperty()
        if component:
            comp = m.GetComponentByName(component)
            if comp is None:
                raise RuntimeError(f"Component {component!r} not found")
            mp.AddBodies(dispatch_array(comp.GetBodies3(c.swSolidBody, None)[0] or ()))
        mp.UseSystemUnits = True
        out = {"mass_kg": mp.Mass, "volume_mm3": mp.Volume * 1e9, "surface_area_mm2": mp.SurfaceArea * 1e6,
               "center_of_mass_mm": [to_mm(v) for v in mp.CenterOfMass], "density_kg_m3": mp.Density}
        if doc_type(m) == "part":
            with contextlib.suppress(Exception):
                out["material"] = m.GetMaterialPropertyName2("")[0] or None
        return _clean(out)
    return await run(_mass)


@mcp.tool()
async def rebuild(force: bool = False, all_configurations: bool = False) -> dict:
    """Rebuild the active document (Ctrl+B; force=True -> Ctrl+Q) and report feature errors.
    For a drawing this updates the views after the model changed."""
    def _rebuild():
        m = model()
        wait_idle()
        if force:
            ok = m.ForceRebuild3(not all_configurations)
        else:
            ok = m.EditRebuild3()
        return {"ok": bool(ok), "errors": feature_errors(m) if doc_type(m) != "drawing" else []}
    return await run(_rebuild)


@mcp.tool()
async def view(orientation: Literal["front", "back", "left", "right", "top", "bottom", "iso",
                                    "trimetric", "dimetric", "normal"] | None = None,
               zoom_fit: bool = True,
               display: Literal["shaded_edges", "shaded", "hidden_removed", "hidden_gray", "wireframe"] | None = None) -> dict:
    """Change the 3D view of the active document (orientation, zoom to fit, display style).
    Does not change the model."""
    def _view():
        m = model()
        if orientation and doc_type(m) != "drawing":
            m.ShowNamedView2("", STD_VIEWS[orientation])
        if display and doc_type(m) != "drawing":
            m.ActiveView.DisplayMode = {
                "shaded_edges": c.swViewDisplayMode_ShadedWithEdges, "shaded": c.swViewDisplayMode_Shaded,
                "hidden_removed": c.swViewDisplayMode_HiddenLinesRemoved,
                "hidden_gray": c.swViewDisplayMode_HiddenLinesGrayed,
                "wireframe": c.swViewDisplayMode_Wireframe}[display]
        if zoom_fit:
            m.ViewZoomtofit2()
        return {"ok": True}
    return await run(_view)


def _capture(hwnd: int):
    import win32gui
    import win32ui
    from PIL import Image as PILImage

    if win32gui.IsIconic(hwnd):
        raise RuntimeError("SOLIDWORKS window is minimized - ask the user to restore it for a screenshot")
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r - l, b - t
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc_dc, w, h)
    save_dc.SelectObject(bmp)
    try:
        ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)  # PW_RENDERFULLCONTENT
        info = bmp.GetInfo()
        img = PILImage.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]),
                                  bmp.GetBitmapBits(True), "raw", "BGRX", 0, 1)
    finally:
        win32gui.DeleteObject(bmp.GetHandle())
        save_dc.DeleteDC(); mfc_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
    if img.getextrema() == ((0, 0), (0, 0), (0, 0)):  # PrintWindow failed -> grab the window area
        from PIL import ImageGrab
        img = ImageGrab.grab(bbox=(l, t, r, b), all_screens=True)
    return img


@mcp.tool()
async def screenshot(max_size: int = 1400, zoom_fit: bool = False,
                     orientation: Literal["front", "top", "right", "left", "iso"] | None = None) -> Image:
    """Picture of the SOLIDWORKS window (only that window), to check results visually.
    zoom_fit / orientation change the view first (not the model)."""
    def _shot():
        m = _connect().ActiveDoc
        if m is not None and (zoom_fit or orientation):
            if orientation and doc_type(m) != "drawing":
                m.ShowNamedView2("", STD_VIEWS[orientation])
            m.ViewZoomtofit2()
        return _main_hwnd()
    hwnd = await run(_shot)
    img = _capture(hwnd)
    img.thumbnail((max_size, max_size))
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return Image(data=buf.getvalue(), format="png")


@mcp.tool()
async def export(path: str, sheets: Literal["current", "all"] = "all") -> dict:
    """Export a copy of the active document; format by extension: .step/.stp, .igs/.iges,
    .x_t, .stl, .3mf, .pdf, .dxf, .dwg, .png, .jpg, .tif, .eprt/.edrw. The open document
    keeps its name. Drawings: sheets=all exports every sheet to PDF/DXF/DWG."""
    def _export():
        m = model()
        wait_idle()
        p = win_path(path)
        Path(p).parent.mkdir(parents=True, exist_ok=True)
        data = None
        ext = Path(p).suffix.lower()
        if doc_type(m) == "drawing" and ext == ".pdf":
            data = _connect().GetExportFileData(c.swExportPdfData)
            names = list(m.GetSheetNames()) if sheets == "all" else [m.GetCurrentSheet().GetName()]
            data.SetSheets(c.swExportData_ExportSpecifiedSheets, names)
        ok, err, warn = m.Extension.SaveAs3(p, 0, c.swSaveAsOptions_Silent | c.swSaveAsOptions_Copy,
                                            data, None, 0, 0)
        return {"ok": bool(ok) and Path(p).exists(), "path": p, "errors": err, "warnings": warn}
    return await run(_export)


# --- drawings ---------------------------------------------------------------

def _view_info(v: SW) -> dict:
    info: dict[str, Any] = {"name": v.Name}
    with contextlib.suppress(Exception):
        info["model"] = v.GetReferencedModelName() or None
        info["configuration"] = v.ReferencedConfiguration or None
        info["scale"] = _clean(v.ScaleDecimal)
        info["position_mm"] = [to_mm(x) for x in v.Position]
        info["outline_mm"] = [to_mm(x) for x in v.GetOutline()]
        info["type"] = v.Type
        info["orientation"] = v.GetOrientationName()
    with contextlib.suppress(Exception):
        info["dimensions"] = v.GetDimensionCount5() if v.has("GetDimensionCount5") else v.GetDimensionCount4()
    return _clean(info)


def sheet_views(drw: SW):
    """{sheet name: [views]} - the first 'view' of each sheet is the sheet itself."""
    out = {}
    for sv in drw.GetViews() or ():
        sv = list(sv)
        out[sv[0].Name] = sv[1:]
    return out


@mcp.tool()
async def drawing_info() -> dict:
    """Active drawing: sheets (size, scale, format), views with the referenced model and
    configuration (the model<->drawing link), scale, position/outline in mm."""
    def _info():
        d = model()
        need(d, "drawing")
        out = {"document": doc_summary(d), "sheets": []}
        sv = sheet_views(d)
        for name in d.GetSheetNames():
            sh = d.Sheet(name)
            props = list(sh.GetProperties2())  # paper, templ, scale1, scale2, first_angle, w, h, same_custom
            out["sheets"].append({
                "name": name, "format": sh.GetTemplateName(),
                "scale": f"{props[2]:g}:{props[3]:g}", "first_angle_projection": bool(props[4]),
                "size_mm": [to_mm(props[5]), to_mm(props[6])],
                "views": [_view_info(v) for v in sv.get(name, [])],
            })
        out["current_sheet"] = d.GetCurrentSheet().GetName()
        return out
    return await run(_info)


def _model_path(app: SW, model_path: str | None) -> str:
    if model_path:
        return win_path(model_path)
    m = app.ActiveDoc
    if m is None or doc_type(m) == "drawing":
        raise RuntimeError("Pass model_path, or activate the part/assembly to draw")
    p = m.GetPathName()
    if not p:
        raise RuntimeError("The model is not saved yet - a drawing needs a saved model file. "
                           "Ask the user to save it (or save_as with their consent).")
    return p


def _sheet_format_path(spec: str, paper: str, orientation: str) -> str:
    if spec.lower() not in ("gost", "gost2"):
        return win_path(spec)
    form = "sh2" if spec.lower() == "gost2" else "sh1"
    orient = "port" if paper == "A4" or orientation == "portrait" else "land"
    name = f"{paper.lower()} - gost_{form}_{orient}.slddrt"
    for d in (_connect().GetUserPreferenceStringValue(c.swFileLocationsSheetFormat) or "").split(";"):
        if d and (Path(d) / name).is_file():
            return str(Path(d) / name)
    raise RuntimeError(f"Sheet format {name} not found in the sheet format folders")


@mcp.tool()
async def create_drawing(model_path: str | None = None, template: str | None = None,
                         paper: Literal["A4", "A3", "A2", "A1", "A0"] = "A3",
                         orientation: Literal["landscape", "portrait"] = "landscape",
                         sheet_format: str | None = "gost", views: list[str] | None = None,
                         scale: list[float] | None = None, insert_annotations: bool = True) -> dict:
    """New drawing linked to a model (default: the active part/assembly - it must be saved).
    paper + orientation (A4 is always portrait). sheet_format: "gost" = ГОСТ 2.104 frame with
    the main title block, first sheet (форма 1, "<paper> - gost_sh1_<land|port>.slddrt");
    "gost2" = subsequent-sheet form; a .slddrt path; or null to keep the template's format.
    views: "standard" = front+top+left in the sheet's projection (ГОСТ: first angle), plus any
    of "iso", "front", "top", "left", "right", "bottom", "back" (default ["standard"]).
    scale [num, den], e.g. [1, 2]. insert_annotations brings the model dimensions marked for
    drawing. One undo step."""
    def _create():
        app = _connect()
        wait_idle()
        mp = _model_path(app, model_path)
        tpl = win_path(template) if template else _default_template("drawing")
        d = app.NewDocument(tpl, PAPER.get(paper.upper(), 8), 0.0, 0.0)
        if d is None:
            raise RuntimeError(f"NewDocument failed for template {tpl}")
        sh = d.GetCurrentSheet()
        if sheet_format:
            fmt = _sheet_format_path(sheet_format, paper, orientation)
            pw, ph = PAPER_MM[paper]
            if paper != "A4" and orientation == "portrait":
                pw, ph = ph, pw
            props = list(sh.GetProperties2())
            prop_view = sh.CustomPropertyView  # localized "Default" - keep it
            if not d.SetupSheet5(sh.GetName(), c.swDwgPapersUserDefined, c.swDwgTemplateCustom,
                                 props[2], props[3], True, fmt, mm(pw), mm(ph), prop_view, True):
                raise RuntimeError(f"Could not apply sheet format {fmt}")
            sh = d.GetCurrentSheet()
        props = list(sh.GetProperties2())
        first_angle = bool(props[4])
        w, h = props[5], props[6]
        created, errors = [], []
        wanted = views or ["standard"]
        with undo_step(d, "чертёж"):
            if scale:
                sh.SetScale(float(scale[0]), float(scale[1]), True, False)
            if "standard" in wanted:
                ok = d.Create1stAngleViews2(mp) if first_angle else d.Create3rdAngleViews2(mp)
                if not ok:
                    errors.append("standard views were not created")
            extra = [v for v in wanted if v != "standard"]
            if extra:
                src = find_doc(mp)
                names = list(src.GetModelViewNames()) if src else []
                for i, key in enumerate(extra):
                    if key not in STD_VIEWS or STD_VIEWS[key] >= len(names):
                        errors.append(f"unknown view {key!r}")
                        continue
                    x, y = w * (0.78 - 0.18 * (i % 3)), h * (0.72 - 0.3 * (i // 3))
                    v = d.CreateDrawViewFromModelView3(mp, names[STD_VIEWS[key]], x, y, 0)
                    (created.append(v.Name) if v else errors.append(f"view {key} failed"))
            annotations = None
            if insert_annotations:
                res = d.InsertModelAnnotations3(c.swImportModelItemsFromEntireModel,
                                                c.swInsertDimensionsMarkedForDrawing, True, False, False, True)
                annotations = len(res) if res else 0
        d.ForceRebuild3(True)
        d.ViewZoomtofit2()
        info = {"drawing": doc_summary(d), "model": mp, "first_angle_projection": first_angle,
                "views": [_view_info(v) for v in sheet_views(d).get(sh.GetName(), [])],
                "annotations_inserted": annotations, "errors": errors}
        return _clean(info)
    return await run(_create)


@mcp.tool()
async def add_drawing_view(view: Literal["front", "back", "left", "right", "top", "bottom", "iso",
                                         "trimetric", "dimetric"],
                           x_mm: float, y_mm: float, model_path: str | None = None,
                           scale: list[float] | None = None) -> dict:
    """Add a model view to the current sheet of the active drawing at (x_mm, y_mm) from the
    sheet's lower-left corner. model_path defaults to the model of the drawing's first view.
    One undo step. (Projected/section/detail views: execute_python with the DrawingDoc API.)"""
    def _add():
        app = _connect()
        d = model()
        need(d, "drawing")
        wait_idle()
        mp = model_path
        if not mp:
            for vs in sheet_views(d).values():
                for v in vs:
                    if v.GetReferencedModelName():
                        mp = v.GetReferencedModelName()
                        break
                if mp:
                    break
        if not mp:
            raise RuntimeError("Drawing has no model views yet - pass model_path")
        mp = win_path(mp)
        src = find_doc(mp)
        if src is None:
            raise RuntimeError(f"Model {mp} is not loaded")
        with undo_step(d, "вид"):
            v = d.CreateDrawViewFromModelView3(mp, model_view_name(src, view), mm(x_mm), mm(y_mm), 0)
            if v is None:
                raise RuntimeError("CreateDrawViewFromModelView3 failed")
            if scale:
                v.UseSheetScale = False
                v.ScaleDecimal = float(scale[0]) / float(scale[1])
        return _view_info(v)
    return await run(_add)


_ANNOT_TYPES = {
    "dimensions": "swInsertDimensionsMarkedForDrawing",
    "dimensions_not_marked": "swInsertDimensionsNotMarkedForDrawing",
    "hole_wizard_profile": "swInsertHoleWizardProfileDimensions",
    "hole_wizard_location": "swInsertHoleWizardLocationDimensions",
    "hole_callouts": "swInsertholeCallout",
    "notes": "swInsertNotes",
    "gtols": "swInsertGTols",
    "datums": "swInsertDatums",
    "surface_finish": "swInsertSFSymbols",
    "cosmetic_threads": "swInsertCThreads",
    "center_marks": "swInsertAxes",
}


@mcp.tool()
async def insert_model_annotations(types: list[str] | None = None, views: list[str] | None = None,
                                   duplicate_dims: bool = False) -> dict:
    """Bring model items into the active drawing (Элементы модели) - this keeps the drawing
    associative: changing these dimensions in the drawing changes the model.
    types (default ["dimensions"]): dimensions (marked for drawing), dimensions_not_marked,
    hole_wizard_profile, hole_wizard_location, hole_callouts, notes, gtols, datums,
    surface_finish, cosmetic_threads, center_marks. views: names (default: all views).
    One undo step."""
    def _ins():
        d = model()
        need(d, "drawing")
        wait_idle()
        mask = 0
        for t in types or ["dimensions"]:
            if t not in _ANNOT_TYPES:
                raise ValueError(f"Unknown type {t!r}; use {list(_ANNOT_TYPES)}")
            mask |= getattr(c, _ANNOT_TYPES[t])
        with undo_step(d, "элементы модели"):
            if views:
                d.ClearSelection2(True)
                for vn in views:
                    d.Extension.SelectByID2(vn, "DRAWINGVIEW", 0, 0, 0, True, 0, None, 0)
                res = d.InsertModelAnnotations3(c.swImportModelItemsFromEntireModel, mask, False,
                                                duplicate_dims, False, True)
            else:
                res = d.InsertModelAnnotations3(c.swImportModelItemsFromEntireModel, mask, True,
                                                duplicate_dims, False, True)
        return {"inserted": len(res) if res else 0}
    return await run(_ins)


@mcp.tool()
async def execute_python(code: str, undo_name: str | None = "скрипт") -> dict:
    """Run Python with the full SOLIDWORKS API on the COM thread (VBA-like syntax works:
    model.FirstFeature(), feat.Name, ext.SelectByID2(...)). Pre-defined:
      swApp, model (active doc or None), ext (model.Extension), selmgr, c (all sw* constants),
      mm(x) mm->m, to_mm(x), deg(x) deg->rad, doubles([...]), dispatch_array([...]),
      features(m), subfeatures(f), std_planes(m) -> {"front","top","right"} plane features
      (localized names!), model_view_name(m, "iso"), iter_dimensions(m), persist_id(m, obj),
      from_persist_id(m, id), feature_errors(m), wait_idle(), math, win32com, pythoncom.
    The API uses metres/radians. print() output is returned; assign `result = ...` to return data.
    The whole script is one Undo step (undo_name=None to skip). Example:
      std_planes(model)["front"].Select2(False, 0)
      sk = model.SketchManager; sk.InsertSketch(True)
      sk.CreateCenterRectangle(0, 0, 0, mm(50), mm(30), 0); sk.InsertSketch(True)"""
    def _exec():
        app = _connect()
        wait_idle()
        m = app.ActiveDoc
        ns: dict[str, Any] = {
            "swApp": app, "model": m, "ext": m.Extension if m else None,
            "selmgr": m.SelectionManager if m else None, "c": c,
            "mm": mm, "to_mm": to_mm, "deg": deg, "doubles": doubles, "dispatch_array": dispatch_array,
            "features": features, "subfeatures": subfeatures, "std_planes": std_planes,
            "model_view_name": model_view_name, "iter_dimensions": iter_dimensions,
            "persist_id": persist_id, "from_persist_id": from_persist_id,
            "feature_errors": feature_errors, "wait_idle": wait_idle, "SW": SW,
            "math": math, "win32com": win32com, "pythoncom": pythoncom, "result": None,
        }
        out = io.StringIO()
        try:
            ctx = undo_step(m, undo_name) if (m is not None and undo_name) else contextlib.nullcontext()
            with contextlib.redirect_stdout(out), ctx:
                exec(compile(code, "<execute_python>", "exec"), ns)
            return {"ok": True, "stdout": out.getvalue()[-20000:], "result": _clean(ns.get("result"))}
        except Exception as ex:
            tb = traceback.format_exc()
            if isinstance(ex, pywintypes.com_error):
                tb += "\n" + _format_com_error(ex)
            return {"ok": False, "stdout": out.getvalue()[-8000:], "error": tb[-4000:]}
    return await run(_exec, retries_for=0)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
