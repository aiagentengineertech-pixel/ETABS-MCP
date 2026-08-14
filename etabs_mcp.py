"""
etabs-mcp — MCP server for CSI ETABS via the OAPI (ETABSv1).

Attaches to an ALREADY RUNNING ETABS instance with a model open.
Runs on the machine where ETABS is installed (Windows only).

Design notes
------------
1. All READS go through DatabaseTables.GetTableForDisplayArray(). That is a
   single field-name-keyed code path, so it does not depend on the fragile
   [out]-parameter ordering of the Results.* functions, which shifts between
   ETABS versions. Tables are returned as list[dict[str, str]].

2. Every read calls SelectObj.ClearSelection() first. ETABS scopes many
   display tables to the current selection; that is what silently emptied
   half of an .e2k export and produced a one-member results set.

3. Design preferences are written by editing the preferences TABLE, not via
   DesignSteel.<code>.SetPreference(). SetPreference uses integer item
   indices that are undocumented per-version; table edits are field-name
   based and version-stable.

4. Signatures marked VERIFY are version-sensitive. They are wrapped so the
   raw COM error surfaces instead of failing silently.

Units: every tool takes an explicit `units` argument. Default kip_in_F.
"""

from __future__ import annotations

import concurrent.futures
import functools
import sys
import threading
from typing import Any, Iterable

# The SDK renamed FastMCP -> MCPServer in mcp 2.0 and moved it. Support both,
# so this file works whether the machine has mcp 1.x or 2.x installed.
try:                                                    # mcp >= 2.0
    from mcp.server.mcpserver import MCPServer as _Server
except ImportError:                                     # mcp 1.x
    from mcp.server.fastmcp import FastMCP as _Server

server = _Server("etabs-mcp")

# ETABS eUnits enum
UNITS = {
    "lb_in_F": 1, "lb_ft_F": 2, "kip_in_F": 3, "kip_ft_F": 4,
    "kN_mm_C": 5, "kN_m_C": 6, "kgf_mm_C": 7, "kgf_m_C": 8,
    "N_mm_C": 9, "N_m_C": 10, "Ton_mm_C": 11, "Ton_m_C": 12,
    "kN_cm_C": 13, "kgf_cm_C": 14, "N_cm_C": 15, "Ton_cm_C": 16,
}

# ItemType enum used by all assignment functions
ITEM_TYPE = {"object": 0, "group": 1, "selected": 2}

# LoadPatterns.Add MyType enum
PATTERN_TYPE = {
    "dead": 1, "super_dead": 2, "live": 3, "reduce_live": 4, "quake": 5,
    "wind": 6, "snow": 7, "other": 8, "move": 9, "temperature": 10,
    "roof_live": 11, "notional": 12,
}

# Direction codes for distributed / uniform loads
#  1-3 local axes, 4-6 global X/Y/Z, 7-9 projected global, 10 gravity,
#  11 projected gravity
DIRECTION = {
    "local_1": 1, "local_2": 2, "local_3": 3,
    "global_x": 4, "global_y": 5, "global_z": 6,
    "proj_x": 7, "proj_y": 8, "proj_z": 9,
    "gravity": 10, "proj_gravity": 11,
}

_state: dict[str, Any] = {"etabs": None, "sap": None, "attached_via": None}

# COM is per-thread. MCP runs each sync tool call on a worker thread from the
# anyio pool, and that thread is not guaranteed to be the same one next call.
# Two consequences drive the design below:
#   1. CoInitialize must run on whichever thread is about to touch COM.
#   2. A COM proxy obtained on thread A must not be reused from thread B
#      (RPC_E_WRONG_THREAD), so the attached objects are cached per-thread
#      rather than globally.
_tls = threading.local()

# --------------------------------------------------------------------------
# Single-threaded COM executor
# --------------------------------------------------------------------------
# MCP dispatches each sync tool call to an arbitrary anyio worker thread. COM
# apartment rules mean an interface pointer obtained on one thread cannot be
# used from another, and an instance created with CreateObjectProgID is NOT
# registered in the running-object table, so another thread cannot re-attach
# to it either. Both problems disappear if every COM operation runs on one
# dedicated thread for the life of the process.
_COM_POOL = concurrent.futures.ThreadPoolExecutor(
    max_workers=1, thread_name_prefix="etabs-com"
)


_COM_BUSY: dict[str, Any] = {"current": None, "queued": 0}


def com_thread(fn):
    """Run fn on the single dedicated COM thread and return its result.

    Long ETABS operations (run_analysis, start_design, check_model) can hold
    the COM thread for minutes. Calls made meanwhile queue behind them, and
    the MCP client may time out waiting -- the queued call still runs to
    completion. Use com_status to see what the thread is doing before
    stacking more calls.
    """
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        _COM_BUSY["queued"] += 1

        def tracked():
            _COM_BUSY["queued"] -= 1
            _COM_BUSY["current"] = fn.__name__
            try:
                return fn(*args, **kwargs)
            finally:
                _COM_BUSY["current"] = None

        return _COM_POOL.submit(tracked).result()
    return wrapper


@server.tool()
def com_status() -> dict:
    """Report what the ETABS COM thread is doing RIGHT NOW, without queueing.

    Deliberately does NOT run on the COM thread, so it answers immediately
    even while an analysis or design run is in progress. If `current` names a
    long operation, wait and poll this instead of calling more ETABS tools --
    they would only queue up behind it and time out at the client.
    """
    return {
        "current": _COM_BUSY["current"] or "idle",
        "queued_calls": _COM_BUSY["queued"],
    }




# --------------------------------------------------------------------------
# COM attach
# --------------------------------------------------------------------------

def _co_initialize() -> str:
    """Initialize COM on the current thread. Safe to call repeatedly."""
    if getattr(_tls, "co_init", False):
        return "already initialized on this thread"

    errors = []
    try:
        import comtypes
        try:
            comtypes.CoInitializeEx(comtypes.COINIT_APARTMENTTHREADED)
            _tls.co_init = True
            return "comtypes.CoInitializeEx(APARTMENTTHREADED)"
        except Exception as exc:
            errors.append(f"CoInitializeEx: {exc}")
            try:
                comtypes.CoInitialize()
                _tls.co_init = True
                return "comtypes.CoInitialize()"
            except Exception as exc2:
                errors.append(f"CoInitialize: {exc2}")
    except ImportError as exc:
        errors.append(f"import comtypes: {exc}")

    try:
        import pythoncom
        pythoncom.CoInitialize()
        _tls.co_init = True
        return "pythoncom.CoInitialize()"
    except Exception as exc:
        errors.append(f"pythoncom: {exc}")

    # S_FALSE means COM was already initialized on this thread — not an error.
    _tls.co_init = True
    return "assumed already initialized (" + "; ".join(errors) + ")"

def _attach_strategies():
    """Yield (label, callable) attach attempts, cheapest/most-documented first.

    ETABS registers itself in the COM running-object table (ROT). Different
    hosts see the ROT differently — notably a child process of an MSIX-packaged
    app may not see a normal desktop process's registration — so we try several
    routes rather than assuming one works.
    """

    def via_helper():
        import comtypes.client
        helper = comtypes.client.CreateObject("ETABSv1.Helper")
        helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
        return helper.GetObject("CSI.ETABS.API.ETABSObject")

    def via_comtypes_active():
        import comtypes.client
        return comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")

    def via_win32com():
        import win32com.client
        return win32com.client.GetActiveObject("CSI.ETABS.API.ETABSObject")

    # DELIBERATELY NOT INCLUDED: win32com.client.Dispatch(...) or
    # comtypes.client.CreateObject(...) on the ETABS class. Those are
    # CoCreateInstance calls — if no instance is running they LAUNCH a new
    # blank ETABS ("Untitled", no model, no analysis results) and happily
    # attach to it. That looks like a model-loading failure to the user and
    # silently discards the real model's results. Attach only, never create.
    # GetActiveObject first: on some machines helper.GetObject() returns None
    # even with ETABS plainly running, while GetActiveObject binds correctly.
    return [
        ("comtypes.GetActiveObject", via_comtypes_active),
        ("win32com.GetActiveObject", via_win32com),
        ("helper.GetObject", via_helper),
    ]


def _start_owned_etabs(model_path: str, visible: bool = True) -> Any:
    """Launch a NEW ETABS instance owned by this server and open a model.

    This is the documented CSI OAPI startup path and it does not depend on the
    running-object table, so it works where attaching to a GUI instance fails
    (ROT lookup returning None, stale registrations, broken .EDB file
    associations). The trade-off is a second ETABS window that belongs to the
    server rather than to you.
    """
    import comtypes.client

    _co_initialize()

    # Reuse an instance this server already owns. ApplicationStart() launches a
    # NEW ETABS every time it is called, so without this guard each open_model
    # call leaves another orphaned window behind, and it stops being obvious
    # which one the server is actually driving.
    existing = getattr(_tls, "etabs", None)
    if existing is not None and _state.get("owned"):
        try:
            sap_existing = existing.SapModel
            sap_existing.GetModelFilename(True)          # liveness probe
        except Exception:
            _tls.etabs = None
            _tls.sap = None
            _state["owned"] = False
        else:
            if model_path:
                ret = sap_existing.File.OpenFile(model_path)
                code = ret[-1] if isinstance(ret, (list, tuple)) else ret
                if code != 0:
                    raise RuntimeError(
                        f"Reused the existing ETABS instance but could not open "
                        f"{model_path!r} (return code {code})."
                    )
            _state["model_file"] = model_path
            _state["reused"] = True
            return sap_existing

    helper = comtypes.client.CreateObject("ETABSv1.Helper")
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)

    try:
        etabs = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
    except Exception as exc:
        raise RuntimeError(
            "Could not create an ETABS API object via CreateObjectProgID. "
            f"Underlying error: {exc}"
        ) from exc

    etabs.ApplicationStart()
    sap = etabs.SapModel
    if sap is None:
        raise RuntimeError("ETABS started but SapModel was None.")

    sap.SetPresentUnits(UNITS["kip_in_F"])
    if model_path:
        ret = sap.File.OpenFile(model_path)
        code = ret[-1] if isinstance(ret, (list, tuple)) else ret
        if code != 0:
            raise RuntimeError(
                f"ETABS started but could not open {model_path!r} "
                f"(return code {code}). Check the path exists and is a .EDB."
            )

    _tls.etabs, _tls.sap = etabs, sap
    _state.update({
        "etabs": etabs, "sap": sap, "attached_via": "owned instance (ApplicationStart)",
        "attached_thread": threading.current_thread().name,
        "model_file": model_path, "owned": True, "reused": False,
    })
    return sap


def _sap():
    """Return SapModel for the CURRENT thread, attaching on first use.

    Cached per-thread, not globally: COM apartment rules mean a proxy created
    on one worker thread cannot be reused from another.
    """
    # Validate any cached handle before handing it back. If ETABS was closed
    # and reopened, the old proxy raises "RPC server is unavailable"
    # (0x800706BA) and a stale entry can linger in the running-object table.
    # Probing here means a restarted ETABS recovers on the next tool call
    # instead of requiring the whole MCP server to be restarted.
    cached = getattr(_tls, "sap", None)
    if cached is not None:
        try:
            cached.GetModelFilename(True)
            return cached
        except Exception:
            _tls.sap = None
            _tls.etabs = None

    if not sys.platform.startswith("win"):
        raise RuntimeError(
            "etabs-mcp requires Windows with ETABS installed. "
            f"Current platform: {sys.platform}"
        )

    co_status = _co_initialize()
    attempts: list[str] = [f"CoInitialize: {co_status}"]
    fallback: tuple[str, Any, Any] | None = None

    for label, fn in _attach_strategies():
        try:
            etabs = fn()
        except Exception as exc:
            attempts.append(f"{label}: raised {type(exc).__name__}: {exc}")
            continue
        if etabs is None:
            # GetObject can return None rather than raising when the ROT
            # lookup finds nothing — guard explicitly.
            attempts.append(f"{label}: returned None (no ETABS in this process's ROT)")
            continue
        try:
            sap = etabs.SapModel
        except Exception as exc:
            hint = ""
            if "800706BA" in str(exc) or "RPC server is unavailable" in str(exc):
                hint = (" — this is a STALE running-object-table entry left by "
                        "an ETABS process that has exited; close all ETABS "
                        "windows and reopen the model")
            attempts.append(f"{label}: attached but .SapModel failed: {exc}{hint}")
            continue
        if sap is None:
            attempts.append(f"{label}: attached but SapModel was None")
            continue

        # Prefer an instance with a saved model open. If several ETABS
        # instances are running, the ROT can hand back a blank one, and
        # reading results from it would silently return nothing.
        try:
            fname = sap.GetModelFilename(True) or ""
        except Exception:
            fname = ""
        if not fname:
            attempts.append(
                f"{label}: attached, but that instance has no saved model open "
                "(blank/Untitled) — held as last resort"
            )
            if fallback is None:
                fallback = (label, etabs, sap)
            continue

        _tls.etabs, _tls.sap = etabs, sap
        _state.update({
            "etabs": etabs, "sap": sap, "attached_via": label,
            "attached_thread": threading.current_thread().name,
            "model_file": fname,
        })
        return sap

    if fallback is not None:
        label, etabs, sap = fallback
        _tls.etabs, _tls.sap = etabs, sap
        _state.update({
            "etabs": etabs, "sap": sap, "attached_via": label + " (BLANK MODEL)",
            "attached_thread": threading.current_thread().name,
            "model_file": "",
        })
        raise RuntimeError(
            "Attached to ETABS, but the instance has no saved model open — it "
            "reports an empty filename (an Untitled/blank model). Results "
            "tables would come back empty.\n"
            "Check for more than one ETABS window: close any blank ones, keep "
            "only the instance with your model, then retry.\nAttempts:\n  "
            + "\n  ".join(attempts)
        )

    raise RuntimeError(
        "Could not attach to a running ETABS instance. Confirm ETABS is open "
        "with a model loaded. If it is, this is a COM visibility problem "
        "rather than a configuration one — run the debug_attach tool for "
        "detail.\nAttempts:\n  " + "\n  ".join(attempts)
    )



_TYPELIB: dict[str, Any] = {}


def _etabs_module() -> Any:
    """Return the generated comtypes wrapper module for the ETABS type library.

    `import comtypes.gen.ETABSv1` only succeeds if comtypes has already
    generated the wrapper in THIS interpreter's site-packages. Attaching to a
    running ETABS through GetActiveObject does not generate it, so the import
    fails and every typed interface silently degrades to a bare IDispatch whose
    methods raise AttributeError. Generate it explicitly instead, preferring
    the type library carried by the live COM object so the wrapper always
    matches the ETABS build we are actually driving.
    """
    if "mod" in _TYPELIB:
        return _TYPELIB["mod"]

    mod = None
    how = "not resolved"

    try:
        import comtypes.gen.ETABSv1 as _E
        mod, how = _E, "already generated (comtypes.gen.ETABSv1)"
    except Exception:
        pass

    if mod is None:
        try:
            import comtypes.client
            src = _state.get("etabs") or _state.get("sap")
            tinfo = src.GetTypeInfo(0)
            tlib, _index = tinfo.GetContainingTypeLib()
            mod = comtypes.client.GetModule(tlib)
            how = "generated from the live ETABS object's type library"
        except Exception as exc:
            how = f"live-object generation failed: {type(exc).__name__}: {exc}"

    if mod is None:
        import glob
        import comtypes.client
        patterns = (
            r"C:\Program Files\Computers and Structures\ETABS*\ETABSv1.dll",
            r"C:\Program Files (x86)\Computers and Structures\ETABS*\ETABSv1.dll",
        )
        for pattern in patterns:
            for dll in sorted(glob.glob(pattern), reverse=True):
                try:
                    mod = comtypes.client.GetModule(dll)
                    how = f"generated from {dll}"
                    break
                except Exception:
                    continue
            if mod is not None:
                break

    _TYPELIB["mod"] = mod
    _TYPELIB["how"] = how
    return mod


def _typed(obj: Any, iface_name: str) -> Any:
    """Return an object whose named ETABS methods can actually be called.

    SapModel's direct children (AreaObj, PropArea, DatabaseTables) come back
    already typed, but nested ones such as LoadPatterns.AutoWind arrive as bare
    IDispatch pointers whose method names will not resolve. First try casting
    to the real generated interface. If the type library is unavailable, fall
    back to comtypes' dynamic dispatcher, which resolves names at call time
    through GetIDsOfNames/Invoke -- slower and untyped, but it works.
    """
    mod = _etabs_module()
    if mod is not None:
        iface = getattr(mod, iface_name, None)
        if iface is not None:
            try:
                return obj.QueryInterface(iface)
            except Exception:
                pass
    try:
        from comtypes.client.dynamic import Dispatch
        return Dispatch(obj)
    except Exception:
        return obj


def _set_units(units: str) -> None:
    code = UNITS.get(units)
    if code is None:
        raise ValueError(f"Unknown units {units!r}. Options: {sorted(UNITS)}")
    _sap().SetPresentUnits(code)


def _check(ret: Any, what: str) -> None:
    """ETABS returns 0 on success for most setters."""
    code = ret[-1] if isinstance(ret, (list, tuple)) else ret
    if code != 0:
        raise RuntimeError(f"{what} failed with ETABS return code {code}")


# --------------------------------------------------------------------------
# Table read layer
# --------------------------------------------------------------------------

def _get_table(table_key: str, *, clear_selection: bool = True) -> list[dict[str, str]]:
    """Read any ETABS display table as a list of field-name-keyed dicts."""
    sap = _sap()

    if clear_selection:
        try:
            sap.SelectObj.ClearSelection()
        except Exception:
            pass  # non-fatal; some model states disallow it

    result = sap.DatabaseTables.GetTableForDisplayArray(
        table_key, [], "", 0, [], 0, []
    )
    # [FieldKeyList, TableVersion, FieldsKeysIncluded, NumberRecords, TableData, ret]
    fields = list(result[2])
    n_records = int(result[3])
    data = list(result[4])
    ret = result[-1]

    if ret != 0:
        raise RuntimeError(
            f"Could not read table {table_key!r} (return code {ret}). "
            "Use list_tables to get exact available table keys."
        )
    if not fields or n_records == 0:
        return []

    width = len(fields)
    rows: list[dict[str, str]] = []
    for i in range(n_records):
        chunk = data[i * width:(i + 1) * width]
        if len(chunk) < width:
            break
        rows.append({fields[j]: chunk[j] for j in range(width)})
    return rows


def _num(row: dict[str, str], *keys: str, default: float = 0.0) -> float:
    """Pull the first present key from a row as a float."""
    for k in keys:
        if k in row and str(row[k]).strip() not in ("", "N/A"):
            try:
                return float(row[k])
            except ValueError:
                continue
    return default


def _txt(row: dict[str, str], *keys: str, default: str = "") -> str:
    for k in keys:
        if k in row and str(row[k]).strip():
            return str(row[k]).strip()
    return default


def _first_table(candidates: Iterable[str]) -> tuple[str, list[dict[str, str]]]:
    """Try several table keys, return the first that reads non-empty."""
    last_err = None
    for key in candidates:
        try:
            rows = _get_table(key)
            if rows:
                return key, rows
        except Exception as exc:
            last_err = exc
    if last_err:
        raise RuntimeError(f"None of {list(candidates)} could be read: {last_err}")
    return "", []


# --------------------------------------------------------------------------
# Read tools
# --------------------------------------------------------------------------

@server.tool()
@com_thread
def open_model(model_path: str, visible: bool = True) -> dict:
    """Launch an ETABS instance owned by this server and open a model by path.

    Use this when attaching to your own ETABS window fails — if the COM
    running-object table lookup returns nothing, or ETABS was started via a
    broken .EDB file association. The server drives its own instance, so it
    always knows which model it is reading.

    model_path: full path to the .EDB, e.g.
      C:\\ETABS_Work\\Grenada\\...\\27.2025 Grenada portal frame.EDB

    Close any ETABS windows you do not need first, to avoid confusion about
    which instance is which.
    """
    sap = _start_owned_etabs(model_path, visible=visible)
    try:
        opened = sap.GetModelFilename(True)
    except Exception as exc:
        opened = f"could not confirm: {exc}"
    return {
        "opened": opened,
        "requested": model_path,
        "instance": ("reused the ETABS instance this server already owns"
                     if _state.get("reused") else "launched a new ETABS instance"),
        "attached_via": _state.get("attached_via"),
        "note": "This ETABS instance belongs to the MCP server. Analysis and "
                "design results must exist or be re-run before reading them.",
    }



@server.tool()
@com_thread
def run_api_snippet(code: str) -> dict:
    """Run a Python snippet against the live ETABS API on the COM thread.

    This exists so a wrong COM call can be diagnosed and corrected in place
    rather than by editing this file and restarting the whole MCP host. It is
    the ETABS equivalent of revit-mcp's send_code_to_revit.

    Names available to the snippet:
      sap        - SapModel
      etabs      - the ETABS application object
      typed(o,n) - cast o to generated interface n, with dynamic fallback
      module     - the generated ETABSv1 wrapper module (may be None)
      introspect(o) - interface name, IID and method list from o's ITypeInfo
      comtypes   - the comtypes package

    Return a value by assigning to `result`. Anything printed is captured.
    The snippet runs with full access to the model, so it can modify it --
    prefer the purpose-built tools for anything routine.
    """
    import io
    import contextlib

    ns: dict[str, Any] = {
        "sap": _sap(),
        "etabs": _state.get("etabs"),
        "typed": _typed,
        "module": _etabs_module(),
        "introspect": _introspect,
        "result": None,
    }
    try:
        import comtypes
        import comtypes.client
        ns["comtypes"] = comtypes
    except Exception:
        pass

    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            exec(code, ns)  # noqa: S102 - deliberate; see docstring
    except Exception as exc:
        import traceback
        return {
            "ok": False,
            "error": f"{type(exc).__name__}: {exc}",
            "traceback": traceback.format_exc()[-3000:],
            "stdout": buf.getvalue()[-4000:],
        }

    value = ns.get("result")
    try:
        import json
        json.dumps(value)
    except Exception:
        value = repr(value)
    return {"ok": True, "result": value, "stdout": buf.getvalue()[-8000:]}


def _introspect(obj: Any) -> dict:
    """Ask a COM object what it actually is, via its own ITypeInfo.

    This is ground truth. dir() on a bare IDispatch pointer only shows the
    IDispatch plumbing, which tells you nothing about the real interface, and
    guessing method names from a type library that may not match the installed
    build is how this server kept failing.
    """
    out: dict[str, Any] = {}
    try:
        tinfo = obj.GetTypeInfo(0)
    except Exception as exc:
        return {"error": f"GetTypeInfo failed: {type(exc).__name__}: {exc}"}
    try:
        out["interface_name"] = tinfo.GetDocumentation(-1)[0]
    except Exception as exc:
        out["interface_name"] = f"<{type(exc).__name__}>"
    try:
        attr = tinfo.GetTypeAttr()
        out["iid"] = str(attr.guid)
        out["func_count"] = attr.cFuncs
        names = []
        for i in range(attr.cFuncs):
            try:
                fdesc = tinfo.GetFuncDesc(i)
                names.append(tinfo.GetDocumentation(fdesc.memid)[0])
            except Exception:
                continue
        out["functions"] = sorted(set(names))
    except Exception as exc:
        out["typeattr_error"] = f"{type(exc).__name__}: {exc}"
    return out


@server.tool()
@com_thread
def debug_com(path: str, iface_name: str = "", name_filter: str = "") -> dict:
    """Report what a COM sub-object REALLY is: its interface name, IID and
    every method it exposes, read from the object's own type information.

    path:        dotted sub-path from SapModel, e.g. "LoadPatterns.AutoWind"
    iface_name:  optional generated interface to test a QueryInterface against
    name_filter: optional case-insensitive substring to filter method names
    """
    obj = _sap()
    for part in [p for p in path.split(".") if p]:
        obj = getattr(obj, part, None)
        if obj is None:
            return {"path": path, "error": f"{part!r} does not exist on this object"}

    raw = _introspect(obj)
    result: dict[str, Any] = {"path": path, "raw_object": raw}

    if name_filter and "functions" in raw:
        needle = name_filter.lower()
        result["raw_object"] = dict(raw)
        result["raw_object"]["functions"] = [
            f for f in raw["functions"] if needle in f.lower()
        ]
        result["raw_object"]["func_count_unfiltered"] = len(raw["functions"])

    if iface_name:
        mod = _etabs_module()
        iface = getattr(mod, iface_name, None) if mod else None
        qi: dict[str, Any] = {"iface_found_in_module": iface is not None}
        if iface is not None:
            try:
                typed = obj.QueryInterface(iface)
                qi["queryinterface"] = "ok"
                methods = [m for m in dir(typed) if not m.startswith("_")]
                if name_filter:
                    needle = name_filter.lower()
                    qi["methods"] = sorted(m for m in methods if needle in m.lower())
                    qi["method_count_unfiltered"] = len(methods)
                else:
                    qi["methods"] = sorted(methods)
            except Exception as exc:
                qi["queryinterface"] = f"failed: {type(exc).__name__}: {exc}"
        result["typed"] = qi

    result["typelib"] = {"resolved": _etabs_module() is not None,
                         "how": _TYPELIB.get("how")}
    return result


@server.tool()
@com_thread
def debug_api(path: str = "") -> dict:
    """List the callable members of SapModel or one of its sub-objects.

    Use this instead of guessing API attribute names, which differ between
    ETABS versions. path is a dotted sub-path, e.g. "" for SapModel itself,
    "AreaObj", "PropArea", "LoadPatterns.AutoWind".
    """
    obj = _sap()
    for part in [p for p in path.split(".") if p]:
        obj = getattr(obj, part, None)
        if obj is None:
            return {"path": path, "error": f"{part!r} does not exist on this object"}
    members = [m for m in dir(obj) if not m.startswith("_")]
    mod = _etabs_module()
    return {
        "path": path or "SapModel",
        "count": len(members),
        "members": sorted(members),
        "typelib": {
            "resolved": mod is not None,
            "how": _TYPELIB.get("how"),
            "module": getattr(mod, "__name__", None),
            "has_cAutoWind": hasattr(mod, "cAutoWind") if mod else False,
            "has_cAutoSeismic": hasattr(mod, "cAutoSeismic") if mod else False,
            "has_cMassSource": hasattr(mod, "cMassSource") if mod else False,
            "has_cSapModel": hasattr(mod, "cSapModel") if mod else False,
        },
    }


@server.tool()
@com_thread
def add_membrane_property(
    name: str, thickness: float = 0.5, material: str = "4000Psi",
    weightless: bool = True, shell_type: int = 3, slab_type: int = 0,
    units: str = "kip_in_F",
) -> str:
    """Define a thin membrane area property, optionally with zero self weight.

    A weightless membrane is the correct way to model light cladding or roof
    sheeting in ETABS: the area object stays in the model so it can distribute
    wind and gravity area loads onto the supporting frame, but it contributes
    no phantom self weight or seismic mass. The real sheeting weight is then
    applied as an explicit area load instead.
    """
    _set_units(units)
    sap = _sap()
    # Confirmed signature: SetSlab(Name, SlabType, ShellType, MatProp,
    # Thickness [, Color, Notes, GUID]). eShellType: 1 shell-thin,
    # 2 shell-thick, 3 membrane, 4 layered. eSlabType: 0 slab.
    _check(sap.PropArea.SetSlab(name, slab_type, shell_type, material, thickness),
           f"PropArea.SetSlab({name}, slab_type={slab_type}, shell_type={shell_type})")
    if weightless:
        # modifiers: f11 f22 f12 m11 m22 m12 v13 v23 mass weight
        mods = [1.0] * 8 + [0.0, 0.0]
        _check(sap.PropArea.SetModifiers(name, mods), f"PropArea.SetModifiers({name})")
    return (f"Defined membrane area property {name!r}: {thickness} thick, {material}"
            + (", mass and weight modifiers set to ZERO (carries load, not weight)."
               if weightless else "."))


@server.tool()
@com_thread
def set_area_property(target: str, property_name: str, item_type: str = "object") -> str:
    """Reassign an area object (or a group of them) to a different area property.

    This is how phantom concrete is removed from a model: assign the roof and
    cladding areas to a weightless membrane property instead of the slab
    property a Revit import gave them.
    """
    it = ITEM_TYPE.get(item_type)
    if it is None:
        raise ValueError(f"item_type must be one of {sorted(ITEM_TYPE)}")
    _check(_sap().AreaObj.SetProperty(target, property_name, it),
           f"AreaObj.SetProperty({target})")
    return f"Assigned area property {property_name!r} to {item_type} {target!r}."


@server.tool()
@com_thread
def list_area_objects(units: str = "kip_in_F") -> dict:
    """List every area object with its story and assigned section property."""
    _set_units(units)
    rows = _get_table("Area Assignments - Section Properties")
    out = []
    for r in rows:
        out.append({
            "story": _txt(r, "Story"),
            "label": _txt(r, "Label"),
            "unique_name": _txt(r, "UniqueName"),
            "property": _txt(r, "SectProp"),
            "type": _txt(r, "PropType"),
        })
    by_prop: dict[str, int] = {}
    for o in out:
        by_prop[o["property"]] = by_prop.get(o["property"], 0) + 1
    return {"count": len(out), "by_property": by_prop, "areas": out}


def _edit_table_rows(table_key: str, edits: dict[str, str], match: dict[str, str]) -> list:
    """Read a full table, change fields on matching rows, submit ALL rows back.

    Submitting every row is what makes this safe against the interactive
    import's delete-unlisted-items semantics. Returns [(matched_key, {field:
    (old, new)}), ...]. Raises if ApplyEditedTables reports errors OTHER than
    the benign 'default value assumed' kind, which the caller must check for
    by reading the table back.

    Two hard-won rules are enforced here (2026-08-13 incident):
      1. The model is UNLOCKED before the read. Editing a definitions table
         while analysis results exist makes the display table include
         auto-generated child rows (e.g. 'EQ X(1/3)'), and applying on a
         locked model can silently no-op.
      2. Rows with IsAuto == 'Yes' are NEVER submitted. They are generated
         artifacts, not definitions; importing them once converted them into
         real load patterns and corrupted the pattern set.
    """
    sap = _sap()
    sap.SetModelIsLocked(False)
    rows = _get_table(table_key)
    if not rows:
        raise RuntimeError(f"{table_key!r} read back empty.")
    rows = [r for r in rows if r.get("IsAuto") != "Yes"]
    if not rows:
        raise RuntimeError(f"{table_key!r} contains only auto-generated rows.")
    fields = list(rows[0].keys())
    for f in list(edits) + list(match):
        if f not in fields:
            raise ValueError(f"Field {f!r} not in {table_key!r}. Available: {sorted(fields)}")
    changed = []
    for r in rows:
        if all(r.get(k) == v for k, v in match.items()):
            delta = {f: (r.get(f), v) for f, v in edits.items()}
            r.update(edits)
            changed.append((tuple(r.get(k) for k in match) or ("all",), delta))
    if not changed:
        raise RuntimeError(f"No row of {table_key!r} matched {match!r}.")
    flat: list[str] = []
    for r in rows:
        flat.extend("" if r.get(f) is None else str(r.get(f)) for f in fields)
    _check(sap.DatabaseTables.SetTableForEditingArray(table_key, 1, fields, len(rows), flat),
           "SetTableForEditingArray")
    res = sap.DatabaseTables.ApplyEditedTables(True, 0, 0, 0, 0, "")
    fatal = int(res[0]) if res else 0
    if fatal:
        raise RuntimeError(f"ApplyEditedTables fatal errors. Log: {res[4] if len(res) > 4 else ''}")
    return changed


@server.tool()
@com_thread
def set_auto_wind_asce710(
    pattern: str, wind_speed_mph: float, exposure: str = "C",
    exposure_from: str = "shells", kzt: float = 1.0,
    gust: float = 0.85, kd: float = 0.85,
) -> str:
    """Set ASCE 7-10 auto wind parameters on an existing auto-wind pattern.

    The ETABS API (unlike SAP2000's) exposes NO wind setters at all --
    cAutoWind is an empty interface in the ETABSv1 type library. The only
    programmatic route is the database table 'Load Pattern Definitions -
    Auto Wind - ASCE 7-10', which this tool edits with all rows submitted so
    nothing is deleted. The pattern must already have ASCE 7-10 auto wind
    assigned (from the UI); this tool cannot attach auto wind to a bare
    pattern.

    exposure_from:
      "shells"     - wind pressure applied to area objects. Correct for a
                     portal frame. REQUIRES per-shell Cp assignments; use
                     get_area_normals then set_wind_pressure_coefficient for
                     every wall and roof shell, per wind pattern. Shells with
                     no Cp receive NO wind.
      "diaphragms" - exposure from rigid diaphragm extents. Stories without a
                     diaphragm (portal roofs!) receive NO wind load.

    Verify afterwards by reading the table back and by checking base
    reactions: the along-wind base shear should be roughly q*G*(Cpw+|Cpl|)
    times the projected elevation.
    """
    src = {"shells": "Shells", "diaphragms": "Diaphragms"}.get(exposure_from.lower())
    if src is None:
        raise ValueError("exposure_from must be 'shells' or 'diaphragms'")
    if exposure.upper() not in ("A", "B", "C", "D"):
        raise ValueError("exposure must be A, B, C or D")
    key = "Load Pattern Definitions - Auto Wind - ASCE 7-10"
    changed = _edit_table_rows(
        key,
        edits={"WindSpeed": str(wind_speed_mph), "Exposure": src,
               "ExpType": exposure.upper(), "kzt": str(kzt),
               "GustFact": str(gust), "Kd": str(kd)},
        match={"Name": pattern},
    )
    rows = _get_table(key)
    back = next((r for r in rows if r.get("Name") == pattern), {})
    warn = ""
    if str(back.get("WindSpeed")) != str(wind_speed_mph):
        warn = (f"  WARNING: read-back shows WindSpeed={back.get('WindSpeed')!r} — "
                "ETABS may have rejected a value (check the enum strings).")
    if src == "Shells":
        warn += ("  Shells exposure: assign Cp to every exposed shell with "
                 "set_wind_pressure_coefficient or the pattern produces no load.")
    return (f"{pattern}: V={wind_speed_mph} mph, Exposure {exposure.upper()} "
            f"from {src}, Kzt={kzt}, G={gust}, Kd={kd}. "
            f"Changed: {changed}.{warn}")


@server.tool()
@com_thread
def get_area_normals() -> dict:
    """Each area object's winding-order normal (= its local-3 direction).

    Needed before set_wind_pressure_coefficient: ETABS applies an entered Cp
    along the shell's +local-3 axis, so the SIGN of the value you enter must
    be chosen per shell:  entered_Cp = physical_Cp x (+1 if local-3 points
    INWARD into the building, -1 if it points OUTWARD). Revit imports wind
    shells inconsistently, so never assume.
    """
    import math
    ao = _typed(_sap().AreaObj, "cAreaObj")
    po = _typed(_sap().PointObj, "cPointObj")
    rows = _get_table("Area Assignments - Section Properties")
    out = []
    for row in rows:
        uid = _txt(row, "UniqueName")
        r = ao.GetPoints(uid, 0, [])
        npts, pts = r[0], list(r[1])
        pcs = []
        for p in pts:
            c = po.GetCoordCartesian(p, 0.0, 0.0, 0.0)
            pcs.append((c[0], c[1], c[2]))
        n = [0.0, 0.0, 0.0]
        for k in range(1, npts - 1):
            ax, ay, az = (pcs[k][i] - pcs[0][i] for i in range(3))
            bx, by, bz = (pcs[k + 1][i] - pcs[0][i] for i in range(3))
            n[0] += ay * bz - az * by
            n[1] += az * bx - ax * bz
            n[2] += ax * by - ay * bx
        mag = math.sqrt(sum(x * x for x in n)) or 1.0
        out.append({
            "unique_name": uid, "story": _txt(row, "Story"),
            "label": _txt(row, "Label"), "property": _txt(row, "SectProp"),
            "local3_normal": [round(x / mag, 3) for x in n],
            "centroid": [round(sum(p[i] for p in pcs) / npts, 1) for i in range(3)],
        })
    return {"count": len(out), "areas": out,
            "note": "entered_Cp = physical_Cp x (+1 if local-3 inward, -1 if outward)"}


@server.tool()
@com_thread
def set_wind_pressure_coefficient(
    target: str, pattern: str, cp: float, windward: bool = False,
    item_type: str = "object",
) -> str:
    """Assign a wind pressure coefficient to a shell for ONE wind pattern.

    Verified ETABS API: cAreaObj.SetLoadWindPressure(Name, LoadPat, MyType,
    Cp, ItemType). MyType 1 = windward (pressure varies with height, qz);
    MyType 2 = other (leeward/side/roof, uniform qh).

    SIGN: ETABS applies the entered Cp along the shell's +local-3 axis.
    Run get_area_normals first and flip the physical Cp for any shell whose
    local-3 points OUT of the building. Getting one sign wrong makes side
    suctions add instead of cancel and can flip roof uplift into downforce --
    check base reactions for spurious cross-wind forces afterwards.
    """
    it = ITEM_TYPE.get(item_type)
    if it is None:
        raise ValueError(f"item_type must be one of {sorted(ITEM_TYPE)}")
    ao = _typed(_sap().AreaObj, "cAreaObj")
    _check(ao.SetLoadWindPressure(target, pattern, 1 if windward else 2, float(cp), it),
           f"AreaObj.SetLoadWindPressure({target}, {pattern})")
    kind = "windward (qz varies with height)" if windward else "other (uniform qh)"
    return f"{target}: Cp={cp} as {kind} in pattern {pattern!r}."



# ETABS applies table edits through its INTERACTIVE DATABASE IMPORT, whose
# default options include "Other Items Deleted from DB: Delete item from
# model". Submitting a table therefore asserts that its rows are the complete
# set — anything absent is DELETED. That is safe for a single-row preferences
# table and catastrophic for a definitions table: editing one field of the auto
# seismic table once deleted six load patterns and every load applied to them.
_TABLE_EDIT_DENYLIST = (
    "load pattern definitions",
    "load case definitions",
    "load combination definitions",
    "story definitions",
    "grid definitions",
    "frame section property",
    "area section property",
    "material properties",
    "mass source",
    "diaphragm definitions",
    "group definitions",
)

@server.tool()
@com_thread
def set_table_value(
    table_key: str, field: str, value: str,
    match_field: str = "", match_value: str = "",
    i_accept_unlisted_rows_are_deleted: bool = False,
) -> str:
    """Set a field on one or more rows of any editable ETABS table.

    The universal fallback when a direct API setter has an awkward or
    version-specific signature. Table edits are field-name based, so they do
    not depend on argument order. Leave match_field empty to set the field on
    every row.

    Example — retarget an auto seismic pattern without touching the API:
      table_key="Load Pattern Definitions - Auto Seismic - ASCE 7-10"
      match_field="Name", match_value="EQ X", field="Ss", value="1.441"
    """
    sap = _sap()
    low = table_key.lower()
    if any(m in low for m in _TABLE_EDIT_DENYLIST) and not i_accept_unlisted_rows_are_deleted:
        raise RuntimeError(
            f"REFUSED: {table_key!r} is a definitions table. ETABS applies table "
            "edits as an interactive database import that DELETES any item not "
            "present in the submitted table. Editing one field here has "
            "previously destroyed unrelated load patterns and every load applied "
            "to them. Use the dedicated API tool for this object instead. If you "
            "genuinely intend the submitted rows to become the complete set, "
            "pass i_accept_unlisted_rows_are_deleted=True."
        )
    # Unlock first and never submit auto-generated rows -- see _edit_table_rows.
    sap.SetModelIsLocked(False)
    rows = _get_table(table_key)
    if not rows:
        raise RuntimeError(f"{table_key!r} read back empty.")
    rows = [r for r in rows if r.get("IsAuto") != "Yes"]
    if not rows:
        raise RuntimeError(f"{table_key!r} contains only auto-generated rows.")
    fields = list(rows[0].keys())
    if field not in fields:
        raise ValueError(f"Field {field!r} not in {table_key!r}. Available: {sorted(fields)}")
    if match_field and match_field not in fields:
        raise ValueError(f"match_field {match_field!r} not in {table_key!r}.")

    changed = []
    for r in rows:
        if match_field and r.get(match_field) != match_value:
            continue
        changed.append((r.get(match_field, "?"), r.get(field)))
        r[field] = value
    if not changed:
        raise RuntimeError(
            f"No row matched {match_field}={match_value!r}. Values present: "
            f"{sorted({r.get(match_field, '') for r in rows})}"
        )

    flat: list[str] = []
    for r in rows:
        flat.extend("" if r.get(f) is None else str(r.get(f)) for f in fields)
    _check(sap.DatabaseTables.SetTableForEditingArray(table_key, 1, fields, len(rows), flat),
           "SetTableForEditingArray")
    res = sap.DatabaseTables.ApplyEditedTables(True, 0, 0, 0, 0, "")
    fatal = int(res[0]) if res else 0
    errors = int(res[1]) if len(res) > 1 else 0
    if fatal or errors:
        raise RuntimeError(f"ApplyEditedTables: {fatal} fatal, {errors} errors. "
                           f"Log: {res[4] if len(res) > 4 else ''}")
    return (f"{table_key}: set {field} = {value!r} on {len(changed)} row(s). "
            f"Previous: {changed[:6]}")


@server.tool()
@com_thread
def debug_attach() -> dict:
    """Diagnose the COM connection to ETABS without raising.

    Reports the host process context, whether an ETABS process is visible, and
    the exact outcome of every attach strategy. Run this when model_info fails.
    """
    import os

    report: dict[str, Any] = {
        "platform": sys.platform,
        "python": sys.version.split()[0],
        "python_exe": sys.executable,
        "bits": 64 if sys.maxsize > 2**32 else 32,
        "pid": os.getpid(),
        "thread": threading.current_thread().name,
        "co_initialize": _co_initialize(),
    }

    # Is the host process inside an MSIX package? A packaged child process can
    # be blocked from seeing a normal desktop app's ROT registration.
    report["package_family"] = os.environ.get("MSIX_PACKAGE_FAMILY_NAME") or None
    report["appdata"] = os.environ.get("APPDATA")
    report["packaged_context"] = "Packages" in (os.environ.get("LOCALAPPDATA") or "")

    # Can we see an ETABS process at all?
    try:
        import subprocess
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq ETABS.exe"],
            capture_output=True, text=True, timeout=15,
        ).stdout
        report["etabs_process_visible"] = "ETABS.exe" in out
        report["tasklist_raw"] = out.strip().splitlines()[-3:]
    except Exception as exc:
        report["etabs_process_visible"] = f"could not check: {exc}"

    # Are the COM libraries importable?
    for mod in ("comtypes", "comtypes.client", "win32com.client"):
        try:
            __import__(mod)
            report[f"import_{mod.replace('.', '_')}"] = "ok"
        except Exception as exc:
            report[f"import_{mod.replace('.', '_')}"] = f"FAILED: {exc}"

    # Try every strategy and record the outcome of each.
    results = []
    for label, fn in _attach_strategies():
        entry: dict[str, Any] = {"strategy": label}
        try:
            obj = fn()
            if obj is None:
                entry["result"] = "returned None"
            else:
                entry["result"] = "got object"
                try:
                    sap = obj.SapModel
                    entry["sap_model"] = "ok" if sap is not None else "None"
                    if sap is not None:
                        try:
                            entry["model_file"] = sap.GetModelFilename(True)
                        except Exception as exc:
                            entry["model_file"] = f"failed: {exc}"
                except Exception as exc:
                    entry["sap_model"] = f"failed: {exc}"
        except Exception as exc:
            entry["result"] = f"{type(exc).__name__}: {exc}"
        results.append(entry)
    report["attach_attempts"] = results
    report["attached_via"] = _state.get("attached_via")
    return report


@server.tool()
@com_thread
def model_info() -> dict:
    """Model file path, present units, and the design codes currently set."""
    sap = _sap()
    rows = _get_table("Program Control")
    info = rows[0] if rows else {}
    return {
        "file": sap.GetModelFilename(True),
        "program_control": info,
    }


@server.tool()
@com_thread
def list_tables(filter_text: str = "") -> list[str]:
    """List available table keys, optionally filtered by a case-insensitive substring.

    Call this first when a table key is rejected — keys differ between versions.
    """
    sap = _sap()
    result = sap.DatabaseTables.GetAllTables(0, [], [], [], [])
    keys = list(result[1]) if len(result) > 1 else []
    if filter_text:
        needle = filter_text.lower()
        keys = [k for k in keys if needle in str(k).lower()]
    return [str(k) for k in keys]


@server.tool()
@com_thread
def get_table(table_key: str, units: str = "kip_in_F", max_rows: int = 400) -> dict:
    """Read any ETABS table by key. Selection is cleared first so the whole
    model is reported, not just selected objects."""
    _set_units(units)
    rows = _get_table(table_key)
    return {
        "table": table_key,
        "units": units,
        "row_count": len(rows),
        "truncated": len(rows) > max_rows,
        "rows": rows[:max_rows],
    }


@server.tool()
@com_thread
def get_base_reactions(units: str = "kip_in_F") -> dict:
    """Base reactions per output case. Use this to verify seismic base shear
    against Cs*W and to total gravity loads."""
    _set_units(units)
    _, rows = _first_table(["Base Reactions"])
    out = []
    for r in rows:
        _case = _txt(r, "OutputCase", "Output Case")
        _step = _txt(r, "StepType", "Step Type")
        out.append({
            "case": _case,
            "step": _step,
            "is_modal": _is_modal(_case, _step),
            "FX": _num(r, "FX", "GlobalFX"),
            "FY": _num(r, "FY", "GlobalFY"),
            "FZ": _num(r, "FZ", "GlobalFZ"),
            "MX": _num(r, "MX", "GlobalMX"),
            "MY": _num(r, "MY", "GlobalMY"),
            "MZ": _num(r, "MZ", "GlobalMZ"),
        })
    return {"units": units, "reactions": out}


@server.tool()
@com_thread
def get_modal(units: str = "kip_in_F") -> dict:
    """Modal periods plus cumulative participating mass ratios.

    Reports the first mode reaching 90% cumulative mass in each direction so
    you can check ASCE 7 §12.9.1 mode-count adequacy.
    """
    _set_units(units)
    _, periods = _first_table(["Modal Periods And Frequencies"])
    _, ratios = _first_table(["Modal Participating Mass Ratios"])

    modes = []
    for r in periods:
        modes.append({
            "mode": int(_num(r, "Mode", "StepNum")),
            "period": _num(r, "Period"),
            "frequency": _num(r, "Frequency"),
        })

    mass = []
    ninety = {"UX": None, "UY": None}
    for r in ratios:
        m = int(_num(r, "Mode", "StepNum"))
        sum_ux, sum_uy = _num(r, "SumUX"), _num(r, "SumUY")
        mass.append({
            "mode": m, "period": _num(r, "Period"),
            "UX": _num(r, "UX"), "UY": _num(r, "UY"), "RZ": _num(r, "RZ"),
            "SumUX": sum_ux, "SumUY": sum_uy, "SumRZ": _num(r, "SumRZ"),
        })
        if ninety["UX"] is None and sum_ux >= 0.90:
            ninety["UX"] = m
        if ninety["UY"] is None and sum_uy >= 0.90:
            ninety["UY"] = m

    t1 = modes[0]["period"] if modes else None
    return {
        "units": units,
        "T1": t1,
        "modes": modes,
        "participating_mass": mass,
        "first_mode_reaching_90pct": ninety,
    }


@server.tool()
@com_thread
def get_drifts(units: str = "kip_in_F", cd: float = 1.0, importance: float = 1.0) -> dict:
    """Story drifts, worst per story and direction.

    Pass `cd` (deflection amplification factor) and `importance` to also get
    the amplified drift Cd*delta/I that ASCE 7 §12.8.6 actually limits.
    Also returns max/avg drift ratios for the torsional irregularity check.
    """
    _set_units(units)
    _, rows = _first_table(["Story Drifts"])

    worst: dict[tuple[str, str], dict] = {}
    for r in rows:
        case = _txt(r, "OutputCase", "Output Case")
        if _is_modal(case, _txt(r, "StepType", "Step Type")):
            continue
        story = _txt(r, "Story")
        direction = _txt(r, "Direction")
        drift = _num(r, "Drift")
        key = (story, direction)
        if key not in worst or drift > worst[key]["drift"]:
            worst[key] = {
                "story": story,
                "direction": direction,
                "case": _txt(r, "OutputCase", "Output Case"),
                "drift": drift,
                "amplified_drift": drift * cd / importance if importance else None,
            }

    torsion = []
    try:
        _, trows = _first_table(["Story Max Over Avg Drifts"])
        tworst: dict[tuple[str, str], float] = {}
        for r in trows:
            if _is_modal(_txt(r, "OutputCase", "Output Case"),
                         _txt(r, "StepType", "Step Type")):
                continue
            key = (_txt(r, "Story"), _txt(r, "Direction"))
            ratio = _num(r, "Ratio", "MaxOverAvg")
            tworst[key] = max(tworst.get(key, 0.0), ratio)
        for (story, direction), ratio in sorted(tworst.items()):
            flag = "extreme (>1.4)" if ratio > 1.4 else ("irregular (>1.2)" if ratio > 1.2 else "ok")
            torsion.append({
                "story": story, "direction": direction,
                "max_over_avg": ratio, "status": flag,
            })
    except Exception:
        pass

    return {
        "units": units, "cd": cd, "importance": importance,
        "worst_drifts": sorted(worst.values(), key=lambda d: (d["story"], d["direction"])),
        "torsion_ratios": torsion,
    }


@server.tool()
@com_thread
def get_story_forces(units: str = "kip_in_F") -> dict:
    """Story shears and overturning moments. Story shear must increase
    monotonically downward; if it does not, lateral load is leaking into
    restraints above the base."""
    _set_units(units)
    _, rows = _first_table(["Story Forces"])
    out = []
    for r in rows:
        out.append({
            "story": _txt(r, "Story"),
            "case": _txt(r, "OutputCase", "Output Case"),
            "location": _txt(r, "Location"),
            "P": _num(r, "P"), "VX": _num(r, "VX"), "VY": _num(r, "VY"),
            "MX": _num(r, "MX"), "MY": _num(r, "MY"), "T": _num(r, "T"),
        })
    return {"units": units, "story_forces": out}


@server.tool()
@com_thread
def get_design_summary(material: str = "steel", units: str = "kip_in_F") -> dict:
    """Design summary: steel PMM/shear ratios, or concrete beam/column results.

    material: "steel" | "concrete_beam" | "concrete_column"
    """
    _set_units(units)
    candidates = {
        "steel": [
            "Steel Frame Design Summary - AISC 360-22",
            "Steel Frame Design Summary - AISC 360-16",
            "Steel Design 1 - Summary Data - AISC 360-16",
        ],
        "concrete_beam": [
            "Concrete Beam Flexure Envelope - ACI 318-19",
            "Concrete Design 1 - Summary Data - ACI 318-19",
        ],
        "concrete_column": [
            "Concrete Column PMM Envelope - ACI 318-19",
            "Concrete Design 2 - Column Summary Data - ACI 318-19",
        ],
    }.get(material)
    if candidates is None:
        raise ValueError("material must be steel, concrete_beam or concrete_column")

    key, rows = _first_table(candidates)
    overstressed = [
        r for r in rows
        if _num(r, "PMMRatio", "PMM Ratio", "Ratio") > 1.0
    ]
    return {
        "table": key, "units": units, "row_count": len(rows),
        "overstressed_count": len(overstressed),
        "overstressed": overstressed[:50],
        "rows": rows[:200],
    }


# --------------------------------------------------------------------------
# Load application
# --------------------------------------------------------------------------

@server.tool()
@com_thread
def add_load_pattern(name: str, pattern_type: str, self_weight_multiplier: float = 0.0) -> str:
    """Create a load pattern. pattern_type: dead, super_dead, live, reduce_live,
    quake, wind, snow, roof_live, temperature, other."""
    code = PATTERN_TYPE.get(pattern_type)
    if code is None:
        raise ValueError(f"pattern_type must be one of {sorted(PATTERN_TYPE)}")
    _check(_sap().LoadPatterns.Add(name, code, self_weight_multiplier, True),
           f"LoadPatterns.Add({name})")
    return f"Added load pattern {name!r} ({pattern_type}, self-weight x{self_weight_multiplier})."


@server.tool()
@com_thread
def apply_frame_load(
    target: str,
    pattern: str,
    value: float,
    direction: str = "gravity",
    units: str = "kip_in_F",
    item_type: str = "object",
    value_end: float | None = None,
    dist_start: float = 0.0,
    dist_end: float = 1.0,
    load_type: str = "force",
    replace: bool = True,
) -> str:
    """Apply a distributed load to beams or columns.

    target: frame object name, or group name when item_type="group".
    value: intensity at start (force per unit length in the chosen units).
    value_end: intensity at end for a trapezoid; defaults to `value` (uniform).
    direction: gravity, global_x/y/z, local_1/2/3, proj_x/y/z, proj_gravity.
    dist_start/dist_end: relative distances 0..1 along the member.
    load_type: "force" or "moment".
    """
    _set_units(units)
    d = DIRECTION.get(direction)
    if d is None:
        raise ValueError(f"direction must be one of {sorted(DIRECTION)}")
    it = ITEM_TYPE.get(item_type)
    if it is None:
        raise ValueError(f"item_type must be one of {sorted(ITEM_TYPE)}")
    my_type = 1 if load_type == "force" else 2
    v_end = value if value_end is None else value_end

    _check(
        _sap().FrameObj.SetLoadDistributed(
            target, pattern, my_type, d, dist_start, dist_end,
            value, v_end, "Global", True, replace, it
        ),
        f"FrameObj.SetLoadDistributed({target})",
    )
    shape = "uniform" if v_end == value else f"trapezoid {value}->{v_end}"
    return (f"Applied {shape} {load_type} load to {item_type} {target!r} in pattern "
            f"{pattern!r}, direction {direction}, units {units}.")


@server.tool()
@com_thread
def apply_point_load_on_frame(
    target: str, pattern: str, value: float, distance: float,
    direction: str = "gravity", units: str = "kip_in_F",
    item_type: str = "object", relative: bool = True, replace: bool = True,
) -> str:
    """Apply a concentrated load at a point along a frame member."""
    _set_units(units)
    d = DIRECTION.get(direction)
    it = ITEM_TYPE.get(item_type)
    if d is None or it is None:
        raise ValueError("bad direction or item_type")
    _check(
        _sap().FrameObj.SetLoadPoint(
            target, pattern, 1, d, distance, value, "Global", relative, replace, it
        ),
        f"FrameObj.SetLoadPoint({target})",
    )
    where = f"{distance:.3f} relative" if relative else f"{distance} absolute"
    return f"Applied {value} point load at {where} on {target!r} in {pattern!r}."


@server.tool()
@com_thread
def apply_area_load(
    target: str, pattern: str, value: float, direction: str = "gravity",
    units: str = "kip_in_F", item_type: str = "object", replace: bool = True,
) -> str:
    """Apply a uniform pressure to a slab, deck or wall area object."""
    _set_units(units)
    d = DIRECTION.get(direction)
    it = ITEM_TYPE.get(item_type)
    if d is None or it is None:
        raise ValueError("bad direction or item_type")
    _check(
        _sap().AreaObj.SetLoadUniform(target, pattern, value, d, replace, "Global", it),
        f"AreaObj.SetLoadUniform({target})",
    )
    return f"Applied {value} pressure to {item_type} {target!r} in {pattern!r} ({units})."


@server.tool()
@com_thread
def apply_area_load_to_frames(
    target: str, pattern: str, value: float, direction: str = "gravity",
    distribution: str = "two_way", units: str = "kip_in_F",
    item_type: str = "object", replace: bool = True,
) -> str:
    """Apply an area pressure that distributes straight onto the supporting beams
    without meshing the slab.

    This is the right tool when the floor is modelled as a membrane or is only
    there to carry load — common in portal frames. distribution: one_way | two_way.
    """
    _set_units(units)
    d = DIRECTION.get(direction)
    it = ITEM_TYPE.get(item_type)
    if d is None or it is None:
        raise ValueError("bad direction or item_type")
    dist_type = 1 if distribution == "one_way" else 2
    _check(
        _sap().AreaObj.SetLoadUniformToFrame(
            target, pattern, value, d, dist_type, replace, "Global", it
        ),
        f"AreaObj.SetLoadUniformToFrame({target})",
    )
    return (f"Applied {value} ({units}) as {distribution} load to frames from "
            f"{item_type} {target!r} in pattern {pattern!r}.")


@server.tool()
@com_thread
def apply_joint_load(
    target: str, pattern: str,
    fx: float = 0.0, fy: float = 0.0, fz: float = 0.0,
    mx: float = 0.0, my: float = 0.0, mz: float = 0.0,
    units: str = "kip_in_F", item_type: str = "object", replace: bool = True,
) -> str:
    """Apply a 6-DOF force/moment vector to a joint."""
    _set_units(units)
    it = ITEM_TYPE.get(item_type)
    if it is None:
        raise ValueError("bad item_type")
    _check(
        _sap().PointObj.SetLoadForce(
            target, pattern, [fx, fy, fz, mx, my, mz], replace, "Global", it
        ),
        f"PointObj.SetLoadForce({target})",
    )
    return f"Applied [{fx},{fy},{fz},{mx},{my},{mz}] to {target!r} in {pattern!r} ({units})."


# --------------------------------------------------------------------------
# Model configuration writes
# --------------------------------------------------------------------------

@server.tool()
@com_thread
def set_mass_source(
    name: str = "MsSrc1",
    from_elements: bool = True,
    from_added_masses: bool = True,
    load_patterns: dict[str, float] | None = None,
    is_default: bool = True,
) -> str:
    """Set the seismic mass source.

    load_patterns maps pattern name -> multiplier, e.g. {"SD": 1.0, "Live": 0.25}.
    Leaving it empty means self-weight only, which understates W whenever
    superimposed dead load is carried by load patterns rather than elements.
    """
    load_patterns = load_patterns or {}
    names = list(load_patterns.keys())
    factors = [float(v) for v in load_patterns.values()]

    # Verified live against ETABS 23.2.0 (2026-08-12): in the ETABS API the
    # mass source lives on cPropMaterial, NOT on a SourceMass/cMassSource
    # object (that is SAP2000's CSiAPIv1 layout -- do not copy signatures
    # from the 4 MB CSiAPIv1 typelib, it is a different product).
    #   cPropMaterial.SetMassSource_1(IncludeElements, IncludeAddedMass,
    #       IncludeLoads, NumberLoads, LoadPat[], SF[]) -> long
    # Read back with GetMassSource_1 to confirm.
    pm = _typed(_sap().PropMaterial, "cPropMaterial")
    _check(
        pm.SetMassSource_1(
            bool(from_elements), bool(from_added_masses), bool(names),
            len(names), names, factors,
        ),
        "PropMaterial.SetMassSource_1",
    )
    back = pm.GetMassSource_1(False, False, False, 0, [], [])
    got_pats = list(back[4]) if len(back) > 4 else []

    included = ", ".join(f"{k} x{v}" for k, v in load_patterns.items()) or "none"
    return (f"Mass source: elements={from_elements}, added={from_added_masses}, "
            f"patterns={included}. Read-back patterns: {got_pats}. "
            f"(name/is_default are ignored: the ETABS API has a single mass "
            f"source.) Re-run analysis.")


@server.tool()
@com_thread
def add_diaphragm(name: str, semi_rigid: bool = False) -> str:
    """Define a diaphragm. Assign it with assign_diaphragm."""
    _check(_sap().Diaphragm.SetDiaphragm(name, semi_rigid),
           f"Diaphragm.SetDiaphragm({name})")
    kind = "semi-rigid" if semi_rigid else "rigid"
    return f"Defined {kind} diaphragm {name!r}."


@server.tool()
@com_thread
def assign_diaphragm(group: str, diaphragm: str) -> str:
    """Assign a diaphragm to every joint in a group.

    Create one group per story first. Stories without a diaphragm receive no
    auto-wind load when Exposure Source is set to Diaphragms, and their
    max/avg drift ratios are meaningless.
    """
    _check(
        _sap().PointObj.SetDiaphragm(group, 3, diaphragm, ITEM_TYPE["group"]),
        f"PointObj.SetDiaphragm({group})",
    )
    return f"Assigned diaphragm {diaphragm!r} to all joints in group {group!r}."


@server.tool()
@com_thread
def set_design_pref(table_key: str, field: str, value: str) -> str:
    """Set one design preference by editing the preferences table.

    Field-name based, so it does not depend on undocumented per-version item
    indices. Example:
      table_key="Steel Frame Design Preferences - AISC 360-22"
      field="Framing Type", value="SMF"
      field="Design System Sds", value="1.527"
      field="Design System Rho", value="1.3"
      field="Seismic Design Category", value="E"

    Use list_tables("Preferences") to find the exact key for your version, and
    get_table(table_key) to see exact field names before writing.
    """
    sap = _sap()
    rows = _get_table(table_key)
    if not rows:
        raise RuntimeError(f"{table_key!r} read back empty; check the key with list_tables.")
    if field not in rows[0]:
        raise ValueError(
            f"Field {field!r} not in {table_key!r}. Available: {sorted(rows[0])}"
        )

    fields = list(rows[0].keys())
    old = rows[0][field]
    rows[0][field] = value
    flat: list[str] = []
    for r in rows:
        flat.extend("" if r.get(f) is None else str(r.get(f)) for f in fields)

    _check(
        sap.DatabaseTables.SetTableForEditingArray(
            table_key, 1, fields, len(rows), flat
        ),
        "SetTableForEditingArray",
    )
    result = sap.DatabaseTables.ApplyEditedTables(True, 0, 0, 0, 0, "")
    fatal = int(result[0]) if result else 0
    errors = int(result[1]) if len(result) > 1 else 0
    log = result[4] if len(result) > 4 else ""
    if fatal or errors:
        raise RuntimeError(f"ApplyEditedTables reported {fatal} fatal, {errors} errors. Log: {log}")
    return f"{table_key}: {field} {old!r} -> {value!r}. Re-run design."


@server.tool()
@com_thread
def set_auto_seismic_asce7_10(
    pattern: str, ss: float, s1: float, site_class: str,
    r: float = 0.0, omega0: float = 0.0, cd: float = 0.0,
    importance: float = 0.0, ecc_ratio: float = 0.0, tl: float = 0.0,
) -> str:
    """Set ASCE 7-10 auto seismic hazard parameters on an existing pattern.

    The ETABS API has NO ASCE 7-10 seismic setter: cAutoSeismic in the
    ETABSv1 typelib offers only SetIBC2006 / SetASCE716 / SetASCE716_1
    (verified live on ETABS 23.2.0 -- the SetIBC2012/SetASCE710 methods seen
    elsewhere belong to SAP2000's CSiAPIv1 typelib, a different product).
    The only programmatic route is the database table 'Load Pattern
    Definitions - Auto Seismic - ASCE 7-10', edited here with all rows
    submitted so nothing is deleted. The pattern must already carry ASCE 7-10
    auto seismic from the UI; this tool cannot attach it to a bare pattern.

    Pass 0 for r/omega0/cd/importance/ecc_ratio/tl to leave the current
    table value unchanged. ETABS derives Fa/Fv/SDS/SD1 itself from Ss, S1
    and site class -- ALWAYS read the table back and check them against your
    hand calculation; that is the strongest single verification available.
    """
    if site_class.upper() not in "ABCDEF" or len(site_class) != 1:
        raise ValueError("site_class must be A-F")
    key = "Load Pattern Definitions - Auto Seismic - ASCE 7-10"
    edits = {"Ss": str(ss), "S1": str(s1), "SiteClass": site_class.upper()}
    if r: edits["R"] = str(r)
    if omega0: edits["Omega"] = str(omega0)
    if cd: edits["Cd"] = str(cd)
    if importance: edits["I"] = str(importance)
    if ecc_ratio: edits["EccRatio"] = str(ecc_ratio)
    if tl: edits["TL"] = str(tl)

    _edit_table_rows(key, edits=edits, match={"Name": pattern})

    rows = _get_table(key)
    back = next((x for x in rows if x.get("Name") == pattern), {})
    fa, fv = back.get("Fa"), back.get("Fv")
    sds, sd1 = back.get("SDS"), back.get("SD1")
    sdc_note = " NOTE: S1 >= 0.75g means SDC E for Risk Category I-III." if s1 >= 0.75 else ""
    return (f"{pattern}: Ss={ss}, S1={s1}, Site {site_class.upper()}"
            + (f", R={r}" if r else "") + (f", Omega0={omega0}" if omega0 else "")
            + (f", Cd={cd}" if cd else "") + (f", I={importance}" if importance else "")
            + f". ETABS computed Fa={fa}, Fv={fv}, SDS={sds}, SD1={sd1} -- "
            f"verify these against ASCE 7-10 Tables 11.4-1/11.4-2 by hand."
            f"{sdc_note}")


@server.tool()
@com_thread
def add_combo(name: str, cases: dict[str, float], combo_type: str = "linear_add") -> str:
    """Create a load combination. cases maps load case name -> scale factor,
    e.g. {"Dead": 1.505, "Live": 1.0, "EQ X": 1.3}."""
    sap = _sap()
    type_code = {"linear_add": 0, "envelope": 1, "absolute_add": 2, "srss": 3}.get(combo_type)
    if type_code is None:
        raise ValueError("combo_type must be linear_add, envelope, absolute_add or srss")
    _check(sap.RespCombo.Add(name, type_code), f"RespCombo.Add({name})")
    for case, sf in cases.items():
        _check(sap.RespCombo.SetCaseList(name, 0, case, float(sf)),
               f"RespCombo.SetCaseList({name}, {case})")
    terms = " + ".join(f"{sf}*{c}" for c, sf in cases.items())
    return f"Created combo {name!r} = {terms}"


@server.tool()
@com_thread
def run_analysis(save_first: bool = True, save_path: str = "") -> str:
    """Run the analysis.

    ETABS requires the model to be saved first. Pass save_path to save to a
    NEW file rather than overwriting the current one — worth doing whenever the
    model may not be in a state you want to keep, since a failed run otherwise
    leaves the overwritten file behind.
    """
    sap = _sap()
    try:
        sap.SetModelIsLocked(False)          # a locked model rejects RunAnalysis
    except Exception:
        pass
    if save_first:
        _check(sap.File.Save(save_path), f"File.Save({save_path or 'in place'})")
    ret = sap.Analyze.RunAnalysis()
    code = ret[-1] if isinstance(ret, (list, tuple)) else ret
    if code != 0:
        raise RuntimeError(
            f"Analyze.RunAnalysis returned {code}. Common causes: a load "
            "combination referencing a deleted load pattern, no load cases set "
            "to run, or an unstable model. Read the 'Analysis Messages' table "
            "for detail."
        )
    return "Analysis complete. Results tables are now current."


@server.tool()
@com_thread
def save_model(path: str = "") -> str:
    """Save the model. Empty path saves in place."""
    _check(_sap().File.Save(path), "File.Save")
    return f"Saved{(' to ' + path) if path else ' in place'}."


# --------------------------------------------------------------------------
# check_model — the audit
# --------------------------------------------------------------------------


# Output cases whose results are not physically meaningful to compare against
# static demands. Mode shapes are normalised arbitrarily, so modal "reactions"
# and "drifts" are scale-free numbers — including them in a max() produces
# spectacular false positives.
_MODAL_MARKERS = ("modal", "mode", "eigen", "buckling")


def _is_modal(case: str, step: str = "") -> bool:
    c = (case or "").strip().lower()
    s = (step or "").strip().lower()
    return any(m in c for m in _MODAL_MARKERS) or s in ("mode", "buckling")


def _finding(severity: str, item: str, detail: str, fix: str = "") -> dict:
    return {"severity": severity, "item": item, "detail": detail, "fix": fix}


@server.tool()
@com_thread
def check_model(units: str = "kip_in_F", drift_limit_coefficient: float = 0.020) -> dict:
    """Audit the model against ASCE 7 and the usual silent-failure modes.

    Checks: SDC from S1; combination D-factor against the real SDS; R against
    the design frame type; mass source completeness; diaphragm coverage against
    loaded stories; T1 against Cu*Ta; story-shear monotonicity against base
    reaction; soft story; torsional irregularity; drift with Cd applied.

    Every finding is advisory. It reports what to look at, not what is true.
    """
    _set_units(units)
    findings: list[dict] = []
    skipped: list[dict] = []
    context: dict[str, Any] = {}

    # --- seismic parameters -------------------------------------------------
    sds = sd1 = None
    r_used = None
    try:
        key, seis = _first_table([
            "Load Pattern Definitions - Auto Seismic - ASCE 7-10",
            "Load Pattern Definitions - Auto Seismic - ASCE 7-16",
            "Load Pattern Definitions - Auto Seismic - ASCE 7-05",
        ])
        if seis:
            row = seis[0]
            ss = _num(row, "Ss", "SS")
            s1 = _num(row, "S1")
            sds = _num(row, "SDS")
            sd1 = _num(row, "SD1")
            r_used = _num(row, "R")
            site = _txt(row, "SiteClass", "Site Class")
            context["seismic"] = {
                "table": key, "Ss": ss, "S1": s1, "SDS": sds, "SD1": sd1,
                "R": r_used, "site_class": site,
                "Omega0": _num(row, "Omega"), "Cd": _num(row, "Cd"),
                "I": _num(row, "I"),
            }
            if s1 >= 0.75:
                findings.append(_finding(
                    "high", "Seismic Design Category",
                    f"S1 = {s1:.3f}g >= 0.75g. ASCE 7 §11.6 puts Risk Category "
                    "I-III at SDC E regardless of the SDC tables.",
                    "Set Seismic Design Category to E in the design preferences.",
                ))
            if site.upper().startswith(("A", "B")):
                findings.append(_finding(
                    "medium", "Site Class",
                    f"Site Class {site} is favourable and needs geotechnical "
                    "justification. The default when unknown is D.",
                    "Confirm against a borehole log or revert to D.",
                ))
            if ss > 1.6:
                findings.append(_finding(
                    "high", "Seismic hazard",
                    f"Ss = {ss:.2f}g is very high — near-fault territory. "
                    "Confirm against the governing hazard map for the site.",
                    "Cross-check Ss and S1 against the project hazard source.",
                ))
    except Exception as exc:
        findings.append(_finding("info", "Auto seismic", f"Could not read: {exc}"))

    # --- design preferences vs. actual SDS and R ----------------------------
    try:
        key, prefs = _first_table([
            "Steel Frame Design Preferences - AISC 360-22",
            "Steel Frame Design Preferences - AISC 360-16",
        ])
        if prefs:
            p = prefs[0]
            # NOTE: DatabaseTables returns field *keys* (FrameType, Sds, SDC),
            # which are NOT the display names shown in the .e2k export
            # ("Frame Type", "Design System Sds"). Keys first, display names
            # kept as fallbacks.
            pref_sds = _num(p, "Sds", "Design System Sds", "SDS")
            pref_r = _num(p, "R", "Design System R")
            frame_type = _txt(p, "FrameType", "Framing Type", "Frame Type")
            rho = _num(p, "Rho", "Design System Rho")
            sdc = _txt(p, "SDC", "Seismic Design Category")
            notional = _txt(p, "AddNotional", "Add Notional Load Case")
            method = _txt(p, "AnalMethod", "Analysis Method")
            omega0 = _num(p, "Omega0", "Design System Omega0")
            cd_pref = _num(p, "Cd", "Design System Cd")
            context["steel_prefs"] = {
                "table": key, "frame_type": frame_type, "R": pref_r,
                "Sds": pref_sds, "Rho": rho, "SDC": sdc,
                "Omega0": omega0, "Cd": cd_pref,
                "analysis_method": method, "notional_loads": notional,
            }

            if sds and pref_sds and abs(pref_sds - sds) > 0.01:
                d_should = 1.2 + 0.2 * sds
                d_is = 1.2 + 0.2 * pref_sds
                findings.append(_finding(
                    "high", "Design System Sds",
                    f"Preference Sds = {pref_sds:.3f} but the auto seismic load "
                    f"gives SDS = {sds:.3f}. Auto combos therefore use "
                    f"{d_is:.3f}D where they should use {d_should:.3f}D, and "
                    f"{0.9 - 0.2 * pref_sds:.3f}D where they should use "
                    f"{0.9 - 0.2 * sds:.3f}D for uplift.",
                    f"Set Design System Sds to {sds:.3f} and regenerate combos.",
                ))

            ft = frame_type.upper()
            if r_used and r_used >= 8 and "SMF" not in ft and "SPECIAL" not in ft:
                findings.append(_finding(
                    "high", "R vs framing type",
                    f"Seismic load uses R = {r_used} (special moment frame) but "
                    f"the design framing type is {frame_type}. An OMF is R = 3.5.",
                    "Either set framing type to SMF, or reduce R to match.",
                ))
            if rho and rho < 1.3 and sdc.upper() in ("D", "E", "F"):
                findings.append(_finding(
                    "medium", "Redundancy factor",
                    f"Rho = {rho} in SDC {sdc}. ASCE 7 §12.3.4.2 requires "
                    "rho = 1.3 unless the redundancy exception is satisfied.",
                    "Confirm the exception or set rho = 1.3.",
                ))
            if "DIRECT" in method.upper() and notional.upper().startswith("N"):
                findings.append(_finding(
                    "medium", "Notional loads",
                    "Direct Analysis Method is selected but notional loads are "
                    "off. AISC 360 App. 7 / C2.2 requires them.",
                    "Enable the notional load case.",
                ))
    except Exception as exc:
        findings.append(_finding("info", "Steel preferences", f"Could not read: {exc}"))

    # --- mass source --------------------------------------------------------
    try:
        _, ms = _first_table(["Mass Source Definition"])
        if ms:
            m = ms[0]
            from_loads = _txt(m, "SourceLoads", "Source Load Patterns?", "MassFromLoads")
            context["mass_source"] = m
            if from_loads.upper().startswith("N"):
                findings.append(_finding(
                    "medium", "Seismic mass",
                    "Mass source excludes load patterns, so any superimposed "
                    "dead load carried as a pattern is not in W.",
                    "Add SD (and a live fraction where required) to the mass source.",
                ))
    except Exception as exc:
        skipped.append({"check": "mass source", "reason": f"{type(exc).__name__}: {exc}"})

    # --- diaphragm coverage -------------------------------------------------
    try:
        stories = [_txt(r, "Name", "Story") for r in _get_table("Story Definitions")]
        diaph_rows = _get_table("Mass Summary by Diaphragm")
        with_diaph = {_txt(r, "Story") for r in diaph_rows}
        missing = [s for s in stories if s and s not in with_diaph]
        context["diaphragms"] = {"stories": stories, "with_diaphragm": sorted(with_diaph)}
        if missing:
            findings.append(_finding(
                "high", "Diaphragm coverage",
                f"Stories without a diaphragm: {', '.join(missing)}. If auto wind "
                "Exposure Source is Diaphragms, these stories receive no wind "
                "load at all, and their max/avg drift ratios are meaningless.",
                "Add diaphragms at every story, or set Exposure Source to area/frame objects.",
            ))
    except Exception as exc:
        skipped.append({"check": "diaphragm coverage", "reason": f"{type(exc).__name__}: {exc}"})

    # --- T1 vs Cu*Ta --------------------------------------------------------
    try:
        modal = get_modal(units=units)
        t1 = modal["T1"]
        context["T1"] = t1
        hn = 0.0
        for r in _get_table("Story Definitions"):
            hn += _num(r, "Height")
        if t1 and hn:
            hn_ft = hn / 12.0 if units.endswith("in_F") else hn
            ta = 0.028 * (hn_ft ** 0.8)  # steel MRF; Ct=0.028, x=0.8
            context["Ta_steel_mrf"] = ta
            if t1 < 0.4 * ta:
                findings.append(_finding(
                    "high", "Model stiffness",
                    f"T1 = {t1:.4f}s against an empirical Ta of {ta:.3f}s for a "
                    f"steel moment frame of height {hn_ft:.1f}ft. The model is "
                    "far stiffer than the structure it represents, so drifts "
                    "and member forces are not trustworthy.",
                    "Check for over-restrained joints, unintended rigid links "
                    "(very large link stiffnesses), or spurious constraints.",
                ))
            if t1 > 2.5 * ta:
                findings.append(_finding(
                    "medium", "Model stiffness",
                    f"T1 = {t1:.4f}s is much longer than Ta = {ta:.3f}s. Check "
                    "for missing lateral system or released members.",
                ))
    except Exception as exc:
        skipped.append({"check": "T1 vs Cu*Ta", "reason": f"{type(exc).__name__}: {exc}"})

    # --- story shear monotonicity ------------------------------------------
    try:
        base = get_base_reactions(units=units)["reactions"]
        forces = get_story_forces(units=units)["story_forces"]
        stories = [_txt(r, "Name", "Story") for r in _get_table("Story Definitions")]
        order = {s: i for i, s in enumerate(stories)}  # index 0 = topmost

        static_base = [b for b in base if not b.get("is_modal")]
        for comp, axis in (("VX", "FX"), ("VY", "FY")):
            worst_base = max((abs(b[axis]) for b in static_base), default=0.0)
            per_story: dict[str, float] = {}
            for f in forces:
                if f["location"].lower() != "bottom":
                    continue
                if _is_modal(f.get("case", "")):
                    continue
                per_story[f["story"]] = max(per_story.get(f["story"], 0.0), abs(f[comp]))
            if not per_story or worst_base < 1e-6:
                continue
            lowest = sorted(per_story, key=lambda s: order.get(s, 99))[-1]
            v_lowest = per_story[lowest]
            if v_lowest < 0.5 * worst_base:
                findings.append(_finding(
                    "high", f"Story shear accumulation ({comp})",
                    f"Shear at the lowest story ({lowest}) is {v_lowest:.2f} but "
                    f"the largest base reaction {axis} is {worst_base:.2f}. Story "
                    "shear should build monotonically to the base reaction; "
                    "lateral load is being absorbed above the base.",
                    "Check for restrained joints above the base and for links "
                    "grounding upper levels.",
                ))
    except Exception as exc:
        skipped.append({"check": "story shear", "reason": f"{type(exc).__name__}: {exc}"})

    # --- soft story ---------------------------------------------------------
    try:
        rows = _get_table("Story Stiffness")
        stories = [_txt(r, "Name", "Story") for r in _get_table("Story Definitions")]
        order = {s: i for i, s in enumerate(stories)}
        for comp, keys in (("Stiff X", ("StiffX", "Stiff X")),
                           ("Stiff Y", ("StiffY", "Stiff Y"))):
            per: dict[str, float] = {}
            for r in rows:
                st = _txt(r, "Story")
                v = _num(r, *keys)
                if v > 0:
                    per[st] = max(per.get(st, 0.0), v)
            ordered = sorted(per, key=lambda s: order.get(s, 99))
            for above, below in zip(ordered, ordered[1:]):
                k_above, k_below = per[above], per[below]
                if k_above <= 0:
                    continue
                ratio = k_below / k_above
                if ratio < 0.60:
                    findings.append(_finding(
                        "high", f"Extreme soft story ({comp})",
                        f"{below} stiffness is {ratio * 100:.1f}% of {above} "
                        "(< 60%). Type 1b irregularity, prohibited in SDC E.",
                    ))
                elif ratio < 0.70:
                    findings.append(_finding(
                        "medium", f"Soft story ({comp})",
                        f"{below} stiffness is {ratio * 100:.1f}% of {above} "
                        "(< 70%). Type 1a irregularity.",
                    ))
    except Exception as exc:
        skipped.append({"check": "soft story", "reason": f"{type(exc).__name__}: {exc}"})

    # --- drift --------------------------------------------------------------
    try:
        cd = context.get("seismic", {}).get("Cd") or 1.0
        imp = context.get("seismic", {}).get("I") or 1.0
        dr = get_drifts(units=units, cd=cd, importance=imp)
        context["worst_amplified_drift"] = max(
            (d["amplified_drift"] or 0.0 for d in dr["worst_drifts"]), default=0.0
        )
        worst = context["worst_amplified_drift"]
        if worst and worst > drift_limit_coefficient:
            findings.append(_finding(
                "high", "Drift",
                f"Worst amplified drift Cd*delta/I = {worst:.5f} exceeds the "
                f"{drift_limit_coefficient} h_sx limit.",
            ))
        elif worst and worst < drift_limit_coefficient / 50:
            findings.append(_finding(
                "medium", "Drift plausibility",
                f"Worst amplified drift is {worst:.6f} — roughly 1/"
                f"{1 / worst:.0f} — which is implausibly small for a code "
                "seismic demand and points at the stiffness problem above.",
            ))
        for t in dr["torsion_ratios"]:
            if t["max_over_avg"] > 1.4:
                findings.append(_finding(
                    "medium", "Torsional irregularity",
                    f"{t['story']} {t['direction']}: max/avg = "
                    f"{t['max_over_avg']:.2f} (extreme, > 1.4). Meaningless if "
                    "that story has no rigid diaphragm — check coverage first.",
                ))
    except Exception as exc:
        skipped.append({"check": "drift and torsion", "reason": f"{type(exc).__name__}: {exc}"})

    order_key = {"high": 0, "medium": 1, "info": 2}
    findings.sort(key=lambda f: order_key.get(f["severity"], 3))
    return {
        "units": units,
        "finding_count": len(findings),
        "high": sum(1 for f in findings if f["severity"] == "high"),
        "findings": findings,
        # A skipped check is NOT a pass. Anything listed here was not
        # evaluated — usually a table with no data, or a field key this
        # ETABS version names differently.
        "checks_skipped": skipped,
        "checks_skipped_count": len(skipped),
        "context": context,
    }


def _main() -> None:
    """stdio by default; --http runs a standalone HTTP server.

    The HTTP mode exists because a COM client running inside the Microsoft
    Store (MSIX) build of the desktop app can be blocked from reaching a
    non-packaged out-of-process COM server such as ETABS, surfacing as
    "RPC server is unavailable" even though the identical code succeeds from
    an ordinary PowerShell process. Running this file yourself in PowerShell
    puts the COM client outside that restriction; the desktop app then talks
    to it over localhost instead of spawning it.

        python etabs_mcp.py --http            # 127.0.0.1:8765
        python etabs_mcp.py --http --port 9000
    """
    if "--http" in sys.argv:
        port = 8765
        if "--port" in sys.argv:
            try:
                port = int(sys.argv[sys.argv.index("--port") + 1])
            except (IndexError, ValueError):
                pass
        try:
            import uvicorn
        except ImportError:
            raise SystemExit("uvicorn is required for --http. pip install uvicorn")
        print(f"etabs-mcp listening on http://127.0.0.1:{port}/mcp", flush=True)
        uvicorn.run(server.streamable_http_app(), host="127.0.0.1", port=port)
    else:
        server.run()


if __name__ == "__main__":
    _main()
