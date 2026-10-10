"""Late-bound SOLIDWORKS COM objects with correct method/property semantics.

SOLIDWORKS objects expose no type info (GetTypeInfo fails), so pywin32 hands back plain
dynamic dispatch objects. Dynamic dispatch can't tell a no-argument method from a property
(`model.GetTitle()` -> "'str' object is not callable") and can't do [out] parameters.

`SW` fixes that: member names are resolved at runtime with GetIDsOfNames (always the right
DISPID for the real interface), while the call signature (method vs property, argument
types, by-ref outputs) is taken from the makepy module generated from sldworks.tlb.
Returned COM objects are wrapped in `SW` again, so agent code reads like VBA:

    model = app.ActiveDoc
    feat = model.FirstFeature()
    while feat:
        print(feat.Name, feat.GetTypeName2())
        feat = feat.GetNextFeature()
"""

from __future__ import annotations

import ast
import pickle
import winreg
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
from win32com.client import gencache

SLDWORKS_TLB_GUID = "{83A33D31-27C5-11CE-BFD4-00400513BB57}"
SWCONST_TLB_GUID = "{4687F359-55D0-4CD3-B6CF-2EB42C11F989}"
# major version of the type library == API revision: 33 = SOLIDWORKS 2025, 34 = 2026...
tlb_version: int | None = None


def registered_tlb_versions(guid: str) -> list[int]:
    """Major versions of a type library registered on this machine (newest first)."""
    out = []
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"TypeLib\{guid}") as k:
            i = 0
            while True:
                try:
                    out.append(int(winreg.EnumKey(k, i).split(".")[0], 16))
                except OSError:
                    break
                except ValueError:
                    pass
                i += 1
    except OSError:
        pass
    return sorted(set(out), reverse=True)

DISPATCH_METHOD = pythoncom.DISPATCH_METHOD
DISPATCH_PROPERTYGET = pythoncom.DISPATCH_PROPERTYGET
DISPATCH_PROPERTYPUT = pythoncom.DISPATCH_PROPERTYPUT
VT_BYREF = pythoncom.VT_BYREF

# name -> list of (flags, rettype, argtypes)
_methods: dict[str, list[tuple]] = {}
_getters: dict[str, list[tuple]] = {}
_putters: dict[str, list[tuple]] = {}
constants: Any = None


def _literal(node):
    return ast.literal_eval(node)


def _index_module(path: Path) -> tuple[dict, dict, dict]:
    tree = ast.parse(path.read_text(encoding="utf-8", errors="replace"))
    methods: dict[str, set] = {}
    getters: dict[str, set] = {}
    putters: dict[str, set] = {}
    for cls in tree.body:
        if not isinstance(cls, ast.ClassDef):
            continue
        for item in cls.body:
            if isinstance(item, ast.FunctionDef) and not item.name.startswith("_"):
                for node in ast.walk(item):
                    if not (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)):
                        continue
                    try:
                        if node.func.attr == "InvokeTypes":  # (dispid, LCID, flags, ret, args, ...)
                            flags, ret, args = (_literal(a) for a in node.args[2:5])
                        elif node.func.attr == "_ApplyTypes_":  # (dispid, flags, ret, args, ...)
                            flags, ret, args = (_literal(a) for a in node.args[1:4])
                        else:
                            continue
                    except (ValueError, IndexError):
                        continue
                    methods.setdefault(item.name, set()).add((flags, ret, args))
                    break
            elif isinstance(item, ast.Assign) and len(item.targets) == 1 \
                    and isinstance(item.targets[0], ast.Name) \
                    and item.targets[0].id in ("_prop_map_get_", "_prop_map_put_") \
                    and isinstance(item.value, ast.Dict):
                target = getters if item.targets[0].id == "_prop_map_get_" else putters
                for k, v in zip(item.value.keys, item.value.values):
                    try:
                        name = _literal(k)
                        spec = _literal(v)
                    except ValueError:
                        continue
                    # get: (dispid, flags, ret, args, name, clsid); put: ((dispid, 0), ())
                    if target is getters:
                        target.setdefault(name, set()).add((spec[1], spec[2], spec[3]))
                    else:
                        target.setdefault(name, set()).add((DISPATCH_PROPERTYPUT, (24, 0), ((12, 1),)))
    as_lists = lambda d: {k: sorted(v, key=repr) for k, v in d.items()}
    return as_lists(methods), as_lists(getters), as_lists(putters)


def init(cache_dir: Path, version: int | None = None) -> None:
    """Generate (once) the makepy modules and the member index.

    `version` is the type library major version (33 = SOLIDWORKS 2025); by default the
    newest registered one is used."""
    global _methods, _getters, _putters, constants, tlb_version
    if _methods:
        return
    if version is None:
        versions = registered_tlb_versions(SLDWORKS_TLB_GUID)
        if not versions:
            raise RuntimeError("SOLIDWORKS type library (sldworks.tlb) is not registered - "
                               "is SOLIDWORKS installed on this machine?")
        version = versions[0]
    tlb_version = version
    mod = gencache.EnsureModule(SLDWORKS_TLB_GUID, 0, version, 0)
    constants = gencache.EnsureModule(SWCONST_TLB_GUID, 0, version, 0).constants
    src = Path(mod.__file__)
    cache = cache_dir / f"swindex_{src.stem}_{int(src.stat().st_mtime)}.pickle"
    if cache.exists():
        _methods, _getters, _putters = pickle.loads(cache.read_bytes())
    else:
        _methods, _getters, _putters = _index_module(src)
        cache_dir.mkdir(parents=True, exist_ok=True)
        cache.write_bytes(pickle.dumps((_methods, _getters, _putters)))


def unwrap(v: Any) -> Any:
    if isinstance(v, SW):
        return v._obj
    if isinstance(v, (list, tuple)):
        return type(v)(unwrap(i) for i in v)
    return v


def wrap(v: Any) -> Any:
    if isinstance(v, pythoncom.TypeIIDs[pythoncom.IID_IDispatch]):
        return SW(v)
    if isinstance(v, tuple):
        return tuple(wrap(i) for i in v)
    return v


_VT_I4, _VT_R8, _VT_BSTR, _VT_DISPATCH, _VT_BOOL, _VT_VARIANT = 3, 5, 8, 9, 11, 12


def _type_score(vt: int, value: Any) -> int:
    """How well a Python value fits a COM argument type (higher is better)."""
    vt &= ~VT_BYREF
    if value is None:
        return 2 if vt in (_VT_DISPATCH, _VT_VARIANT) else 0
    if isinstance(value, bool):
        return 2 if vt == _VT_BOOL else (1 if vt in (_VT_I4, _VT_VARIANT) else 0)
    if isinstance(value, int):
        return 2 if vt == _VT_I4 else (1 if vt in (_VT_R8, _VT_VARIANT, _VT_BOOL) else 0)
    if isinstance(value, float):
        return 2 if vt == _VT_R8 else (1 if vt in (_VT_I4, _VT_VARIANT) else 0)
    if isinstance(value, str):
        return 2 if vt == _VT_BSTR else (1 if vt == _VT_VARIANT else 0)
    if isinstance(value, SW):
        return 2 if vt == _VT_DISPATCH else (1 if vt == _VT_VARIANT else 0)
    return 1 if vt == _VT_VARIANT else 0


def _pick(specs: list[tuple], args: tuple) -> tuple:
    """Signature that fits the call. The same member name exists in many interfaces
    (e.g. Delete() / Delete(int) / Delete(str)), so among signatures with the exact argument
    count the one whose types match the Python values best wins; otherwise the shortest
    signature taking at least `len(args)` (the rest are omitted optional/[out] args)."""
    nargs = len(args)
    exact = [s for s in specs if len(s[2]) == nargs]
    if exact:
        return max(exact, key=lambda s: sum(_type_score(t[0], a) for t, a in zip(s[2], args)))
    longer = sorted((s for s in specs if len(s[2]) > nargs), key=lambda s: len(s[2]))
    return (longer or specs)[0]


class _Method:
    __slots__ = ("_obj", "_name", "_dispid", "_specs")

    def __init__(self, obj, name, dispid, specs):
        self._obj, self._name, self._dispid, self._specs = obj, name, dispid, specs

    def __call__(self, *args):
        flags, ret, argtypes = _pick(self._specs, args)
        args = list(unwrap(args))
        # fill omitted trailing args: [out] params need Missing placeholders
        while len(args) < len(argtypes):
            args.append(pythoncom.Missing)
        if len(args) > len(argtypes):
            raise TypeError(f"{self._name}() takes at most {len(argtypes)} arguments ({len(args)} given)")
        res = self._obj.InvokeTypes(self._dispid, 0, flags, ret, tuple(argtypes), *args)
        return wrap(res)

    def __repr__(self):
        return f"<SW method {self._name}>"


class _PropOrMethod(_Method):
    """Name that is a property in some interfaces and a method in others (e.g. LengthUnit,
    CustomPropertyManager). SOLIDWORKS accepts either invoke kind, so this object can be
    called with arguments or used directly as the property value."""
    __slots__ = ()

    def _value(self):
        return self()

    def __getattr__(self, name):
        return getattr(self._value(), name)

    def __bool__(self): return bool(self._value())
    def __str__(self): return str(self._value())
    def __int__(self): return int(self._value())
    def __float__(self): return float(self._value())
    def __index__(self): return int(self._value())
    def __iter__(self): return iter(self._value())
    def __len__(self): return len(self._value())
    def __getitem__(self, i): return self._value()[i]
    def __eq__(self, other): return self._value() == other
    def __hash__(self): return hash(self._value())
    def __repr__(self): return repr(self._value())


class SW:
    """Wrapper around a SOLIDWORKS IDispatch (see module docstring)."""

    __slots__ = ("_obj", "_ids")

    def __init__(self, obj):
        if isinstance(obj, SW):
            obj = obj._obj
        obj = getattr(obj, "_oleobj_", obj)
        object.__setattr__(self, "_obj", obj)
        object.__setattr__(self, "_ids", {})

    def _dispid(self, name: str) -> int:
        try:
            return self._ids[name]
        except KeyError:
            pass
        try:
            dispid = self._obj.GetIDsOfNames(name)
        except pywintypes.com_error:
            raise AttributeError(f"SOLIDWORKS object has no member {name!r}") from None
        self._ids[name] = dispid
        return dispid

    def __getattr__(self, name: str):
        if name.startswith("__"):
            raise AttributeError(name)
        dispid = self._dispid(name)
        if name in _methods:
            if name in _getters:
                return _PropOrMethod(self._obj, name, dispid, _methods[name] + _getters[name])
            return _Method(self._obj, name, dispid, _methods[name])
        if name in _getters:
            flags, ret, argtypes = _getters[name][0]
            return wrap(self._obj.InvokeTypes(dispid, 0, DISPATCH_PROPERTYGET, ret, tuple(argtypes)))
        # unknown to the type library: plain late-bound property get
        return wrap(self._obj.Invoke(dispid, 0, DISPATCH_PROPERTYGET, 1))

    def __setattr__(self, name: str, value: Any) -> None:
        dispid = self._dispid(name)
        self._obj.Invoke(dispid, 0, DISPATCH_PROPERTYPUT, 0, unwrap(value))

    def __bool__(self) -> bool:
        return True

    def __eq__(self, other) -> bool:
        return isinstance(other, SW) and self._obj == other._obj

    def __hash__(self) -> int:
        return hash(self._obj)

    def __repr__(self) -> str:
        return "<SW object>"

    def put(self, name: str, *args: Any) -> None:
        """Indexed property put, e.g. em.put("Equation", 0, '"L" = 10') (VBA: em.Equation(0) = ...)."""
        self._obj.Invoke(self._dispid(name), 0, DISPATCH_PROPERTYPUT, 0, *unwrap(args))

    @property
    def raw(self):
        """The underlying PyIDispatch (for win32com interop)."""
        return self._obj

    def has(self, name: str) -> bool:
        try:
            self._dispid(name)
            return True
        except AttributeError:
            return False


def get_active(progid: str = "SldWorks.Application") -> SW:
    unk = pythoncom.GetActiveObject(pywintypes.IID(progid))
    return SW(unk.QueryInterface(pythoncom.IID_IDispatch))


def launch(progid: str = "SldWorks.Application") -> SW:
    import win32com.client.dynamic
    return SW(win32com.client.dynamic.Dispatch(progid))
