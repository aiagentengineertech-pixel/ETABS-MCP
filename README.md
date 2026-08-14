# etabs-mcp

MCP server for CSI ETABS via the OAPI. Attaches to an **already running** ETABS
instance with a model open, so there is no Excel export step and no selection
scoping to trip over.

Windows only — ETABS must be installed on the same machine.

## Install

```powershell
pip install mcp comtypes
```

Put `etabs_mcp.py` somewhere stable, e.g. `C:\mcp\etabs-mcp\etabs_mcp.py`.

## Configure

Add to your Claude desktop MCP config, alongside `revit-mcp`:

```json
{
  "mcpServers": {
    "etabs-mcp": {
      "command": "python",
      "args": ["C:\\mcp\\etabs-mcp\\etabs_mcp.py"]
    }
  }
}
```

Restart the desktop app. Open ETABS and load a model **before** calling any tool.

## Tools

### Read

| Tool | Purpose |
|---|---|
| `model_info` | File path, units, design codes in effect |
| `list_tables` | Discover exact table keys for your ETABS version |
| `get_table` | Read any table by key, selection cleared first |
| `get_base_reactions` | Per-case base reactions — verify V against Cs·W |
| `get_modal` | Periods, participating mass, first mode reaching 90% |
| `get_drifts` | Worst drift per story/direction, amplified by Cd/I, plus max/avg ratios |
| `get_story_forces` | Story shears and overturning moments |
| `get_design_summary` | Steel PMM/shear ratios or concrete results, with overstressed members called out |

### Apply loads

| Tool | Purpose |
|---|---|
| `add_load_pattern` | New pattern with self-weight multiplier |
| `apply_frame_load` | UDL or trapezoid on beams/columns; force or moment; any direction |
| `apply_point_load_on_frame` | Concentrated load at a relative or absolute distance |
| `apply_area_load` | Uniform pressure on slab, deck or wall |
| `apply_area_load_to_frames` | One-way or two-way distribution onto supporting beams, no slab meshing |
| `apply_joint_load` | Full 6-DOF force/moment vector |

### Configure

| Tool | Purpose |
|---|---|
| `set_mass_source` | Include load patterns in seismic mass |
| `add_diaphragm` / `assign_diaphragm` | Define and assign, by group |
| `set_design_pref` | Any preference, by field name |
| `set_auto_seismic_asce7_10` | Ss, S1, site class, R, Ω₀, Cd, I |
| `add_combo` | User combinations with explicit factors |
| `run_analysis` / `save_model` | |

### Audit

`check_model` runs the whole review in one call:

- SDC from S1 (S1 ≥ 0.75g → SDC E for Risk Category I–III)
- Combination D-factor against the *real* SDS, not the preference default
- R against the design framing type (R = 8 with OMF detailing)
- Redundancy factor ρ in SDC D and above
- Direct Analysis Method with notional loads off
- Mass source excluding load patterns
- Diaphragm coverage against stories that need wind load
- T₁ against empirical Ta — catches over-restrained and over-stiff models
- Story shear building monotonically to the base reaction
- Soft story at the 70% and 60% thresholds
- Torsional irregularity at 1.2 and 1.4
- Amplified drift Cd·δ/I, in both directions: over the limit, and implausibly under it

Findings come back sorted by severity with a suggested fix. They are advisory —
the tool reports what to look at, not what is true.

## Design decisions worth knowing

**Reads go through `DatabaseTables.GetTableForDisplayArray`, not `Results.*`.**
The `Results` functions rely on `[out]` parameter ordering that shifts between
ETABS versions. Table reads are keyed by field name and survive upgrades.

**Every read calls `SelectObj.ClearSelection()` first.** ETABS scopes display
tables to the current selection. That is what silently emptied most of an `.e2k`
export and reduced a whole-model results set to one member.

**Design preferences are written by editing the preferences table**, not via
`DesignSteel.<code>.SetPreference()`. `SetPreference` takes integer item indices
that are undocumented and version-specific; a wrong index writes the wrong
preference with no error. Table edits are field-name based, so a bad field name
fails loudly.

## Known rough edges

1. **`set_auto_seismic_asce7_10` argument order is unverified.** The
   `cAutoSeismic.SetASCE7_10` signature varies between versions. The call is
   wrapped so the raw COM error surfaces with the ETABS version — correct the
   order against your OAPI documentation on first use. Everything else uses
   signatures that have been stable across v18–v23.

2. **`assign_diaphragm` works by group.** ETABS has no "assign to all joints on
   story X" call, so create one group per story first (or select the joints and
   pass `item_type="selected"` in a direct `PointObj.SetDiaphragm` call).

3. **Table keys differ between versions.** `get_design_summary` tries several
   candidates. If a read fails, `list_tables("Design")` gives the real key.

4. **`check_model` degrades quietly.** Each check is wrapped, so an unreadable
   table skips that check rather than failing the audit. If a finding you expect
   is missing, read the relevant table directly to confirm it is being seen.

## First run

```
model_info                          # confirm attach and units
list_tables "Preferences"           # confirm your version's key names
check_model                         # the audit
```

Then fix in place, `run_analysis`, and `check_model` again.

## Disclaimer

This project is not affiliated with, endorsed by, or supported by Computers
and Structures, Inc. (CSI). It drives a licensed ETABS installation on your
own machine through CSI's published Open API; it contains no CSI software or
type libraries. ETABS is a trademark of Computers and Structures, Inc.

This software can modify structural analysis models, including operations
that ETABS applies through its interactive database import, which deletes
items absent from a submitted table. Safeguards are built in (definitions-
table denylist, unlock-before-edit, auto-row filtering, read-back
verification), but you are responsible for your models: keep backups, verify
results, and treat every write as an engineering action. Use at your own
risk. Nothing produced by this tool is a substitute for review by a
qualified structural engineer.

## License

MIT — see [LICENSE](LICENSE).
