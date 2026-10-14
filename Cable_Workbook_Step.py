"""Final step of a run: fill the instrument list in a copy of the Cable Calculations workbook.

Done entirely in code, with no model:
  1. Each row is built from the combined instrument list by fixed rules (notes and voltage copied, signal from the tag's
     function letters, connection and cable from the wire count/type and the template's own stock-cable table).
  2. The rows are written into a COPY of the workbook through Excel itself (the template is opened read-only and never
     modified). Excel keeps the images, tables and data validation that Python spreadsheet libraries drop.
Afterwards the copy is compared with the template cell by cell, and anything that changed outside the list is logged.
"""
import json
import re
import subprocess
import sys
import tempfile
import threading
import time
import warnings
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

DEFAULT_TEMPLATE = Path(__file__).with_name("Instrument Cable Calculations.xlsx")
LIST_COLUMNS = ["TAG NO", "Controls Notes", "P&IDs Notes", "Voltage", "INTRINSICALLY SAFE", "SIGNAL", "CONNECTION",
                "Cable Designation", "Cable Qty", "Conduit", "Conduit Branch"]
LIST_SHEET_HINT = "Instrumentation Cables"
EXAMPLE_ROWS = 8
MAX_REPORTED_DIFFERENCES = 15
EXCEL_TIMEOUT_SECONDS = 600

_active: List[Tuple[subprocess.Popen, Path]] = []  # (PowerShell process, file holding the PID of the Excel it started)
_active_lock = threading.Lock()


def _taskkill(pid: int, tree: bool = False) -> None:
    subprocess.run(["taskkill", "/F"] + (["/T"] if tree else []) + ["/PID", str(pid)], capture_output=True,
                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) if sys.platform == "win32" else 0)


def _stop(process: subprocess.Popen, pid_file: Path) -> None:
    """Stop the PowerShell run and the one Excel instance it started (identified by PID, so a user's own Excel is safe)."""
    if process.poll() is None:
        _taskkill(process.pid, tree=True)
    try:
        excel_pid = int(pid_file.read_text().strip())
    except (OSError, ValueError):
        return
    _taskkill(excel_pid)


def kill_all() -> None:
    """Called when the window closes mid-run so no Excel instance is left behind."""
    with _active_lock:
        for process, pid_file in list(_active):
            _stop(process, pid_file)

# Runs inside Excel (COM). Copies the template (opened read-only), finds the list by its header row, rewrites its body.
EXCEL_SCRIPT = r"""
param([string]$Template, [string]$Output, [string]$RowsFile, [string]$HeadersFile, [string]$PidFile)
$ErrorActionPreference = 'Stop'
function Read-JsonArray([string]$Path) {
    # Windows PowerShell 5.1 returns a JSON array as ONE object when piped; unwrap it so it behaves as a list.
    $v = ConvertFrom-Json -InputObject (Get-Content -Raw -Encoding UTF8 $Path)
    if ($v.Count -eq 1 -and $v[0] -is [System.Array]) { return ,@($v[0]) }
    return ,@($v)
}
$rows = Read-JsonArray $RowsFile
$headers = Read-JsonArray $HeadersFile
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
# Record this Excel's process id so the caller can stop exactly this instance if it has to.
Add-Type -Namespace Win -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int pid);'
$excelPid = 0
[void][Win.User32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
Set-Content -Path $PidFile -Value $excelPid
$wb = $null
try {
    $wb = $excel.Workbooks.Open($Template, 0, $true)
    $wb.SaveAs($Output)   # from here on only the copy is open
    $target = $null
    foreach ($ws in $wb.Worksheets) {
        foreach ($lo in $ws.ListObjects) {
            $names = @(); foreach ($c in $lo.HeaderRowRange.Cells) { $names += ([string]$c.Value2).Trim() }
            $match = ($names.Count -ge $headers.Count)
            for ($i = 0; $match -and $i -lt $headers.Count; $i++) { if ($names[$i] -ne $headers[$i]) { $match = $false } }
            if ($match) { $target = $lo; break }
        }
        if ($target) { break }
    }
    if (-not $target) { throw "No table with the columns: $($headers -join ', ')" }
    $ws = $target.Parent
    $oldCount = $target.ListRows.Count
    $newCount = $rows.Count
    if ($newCount -lt 1) { throw "No rows to write" }
    if ($newCount -gt $oldCount) {
        $top = $target.Range.Row; $left = $target.Range.Column
        $target.Resize($ws.Range($ws.Cells.Item($top, $left), $ws.Cells.Item($top + $newCount, $left + $target.Range.Columns.Count - 1)))
    } elseif ($newCount -lt $oldCount) {
        for ($i = $oldCount; $i -gt $newCount; $i--) { $target.ListRows.Item($i).Delete() }
    }
    $body = $target.DataBodyRange
    for ($r = 0; $r -lt $newCount; $r++) {
        $vals = @($rows[$r])
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $cell = $body.Cells.Item($r + 1, $c + 1)
            $v = if ($c -lt $vals.Count) { $vals[$c] } else { $null }
            if ($null -eq $v -or "$v" -eq '') { $cell.ClearContents() | Out-Null }
            elseif ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) { $cell.Value2 = [double]$v }
            else { $cell.Value2 = [string]$v }
        }
    }
    $wb.Save()
    Write-Output "OK sheet=$($ws.Name) table=$($target.Name) rows=$newCount (was $oldCount)"
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}
"""


# ------------------------------------------------------------------------------------------------- reading
def _read_list_examples(template: Path) -> Tuple[str, List[List[object]]]:
    """(sheet name, up to EXAMPLE_ROWS existing list rows) read from the template, so Claude can follow its conventions."""
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(template), read_only=True, data_only=True)
        try:
            names = [wb[LIST_SHEET_HINT]] if LIST_SHEET_HINT in wb.sheetnames else []
            names += [ws for ws in wb.worksheets if ws.title != LIST_SHEET_HINT]
            for ws in names:
                rows = list(ws.iter_rows(min_row=1, max_row=60, max_col=len(LIST_COLUMNS), values_only=True))
                for i, row in enumerate(rows):
                    if [str(v or "").strip() for v in row] == LIST_COLUMNS:
                        examples = [[v for v in r] for r in rows[i + 1:i + 1 + EXAMPLE_ROWS] if any(v not in (None, "") for v in r)]
                        return ws.title, examples
        finally:
            wb.close()
    raise RuntimeError(f"Could not find a list with the columns {', '.join(LIST_COLUMNS)} in {template.name}")


def _load_cells(path: Path) -> Dict[str, Dict[Tuple[int, int], str]]:
    """Every non-empty cell (formulas as text) per sheet, read without images so a large workbook loads quickly."""
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=False)
        try:
            sheets: Dict[str, Dict[Tuple[int, int], str]] = {}
            for ws in wb.worksheets:
                cells: Dict[Tuple[int, int], str] = {}
                for row in ws.iter_rows():
                    for cell in row:
                        value = getattr(cell, "value", None)
                        if value is not None and value != "":
                            cells[(cell.row, cell.column)] = str(getattr(value, "text", value)).strip()
                sheets[ws.title] = cells
            return sheets
        finally:
            wb.close()


def compare_workbooks(template: Path, result: Path, list_sheet: str) -> Tuple[List[str], List[str]]:
    """(changes inside the instrument list, changes anywhere else) as readable strings."""
    before, after = _load_cells(template), _load_cells(result)
    inside: List[str] = []
    outside: List[str] = []
    for sheet in sorted(set(before) | set(after)):
        if sheet not in before or sheet not in after:
            outside.append(f"sheet '{sheet}' was {'removed' if sheet not in after else 'added'}")
            continue
        old, new = before[sheet], after[sheet]
        for key in sorted(set(old) | set(new)):
            if old.get(key) == new.get(key):
                continue
            row, col = key
            in_list = sheet == list_sheet and col <= len(LIST_COLUMNS) and row >= 3  # the list sits in columns A:K
            (inside if in_list else outside).append(
                f"{sheet}!R{row}C{col}: {old.get(key, '(empty)')!r} -> {new.get(key, '(empty)')!r}")
    return inside, outside


# ------------------------------------------------------------------------------------------------- rows (code only)
# Every column is derived from the combined list by fixed rules; no model is involved. Conduit and Conduit Branch are
# routing decisions that the drawings do not contain, so they are left empty for the user.
_ANALOG_TAG = re.compile(r"^(?:[A-Z]{0,2}T|[A-Z]?IT|II|[A-Z]E|[A-Z]CV|[PTLF]Y)$")      # transmitters, elements, control valves, I/P converters
_DIGITAL_TAG = re.compile(r"^(?:[A-Z]?S[HL]{0,2}|ZS[OC]?|S[VW]|X[A-Z]|HS|ES[DW]?|[A-Z]S[HL]{1,2})$")  # switches, solenoids, X-devices


def _is_blank(value: object) -> bool:
    return value is None or str(value).strip().lower() in {"", "[unclear]", "unclear", "none", "null"}


def _instrument_notes(inst: dict, which: str) -> str:
    keys = ("controls_notes", "Controls Notes") if which == "controls" else ("pids_notes", "P&ID Notes")
    value = next((inst[k] for k in keys if inst.get(k)), "")
    return "" if _is_blank(value) else str(value)


def signal_for(tag: object) -> Optional[str]:
    """Signal type from the tag's function letters (PT -> 4-20 mA, ZSO -> DIGITAL CONTACT); None when not recognised."""
    match = re.match(r"[A-Za-z]+", str(tag or "").strip())
    letters = match.group(0).upper() if match else ""
    if _ANALOG_TAG.match(letters):
        return "4-20 mA"
    if _DIGITAL_TAG.match(letters):
        return "DIGITAL CONTACT"
    return None


def read_stock_cables(template: Path) -> List[dict]:
    """The template's 'IN STOCK CABLES' table: designation, conductor count and shield, in table order."""
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(template), read_only=True, data_only=True)
        try:
            ws = wb[LIST_SHEET_HINT] if LIST_SHEET_HINT in wb.sheetnames else wb.worksheets[0]
            rows = list(ws.iter_rows(min_row=1, max_row=60, values_only=True))
        finally:
            wb.close()
    for i, row in enumerate(rows):
        cells = [str(v or "").strip().lower() for v in row]
        if "designation" in cells and "conductors" in cells and "shield" in cells:
            d, c, s = cells.index("designation"), cells.index("conductors"), cells.index("shield")
            stock = []
            for r in rows[i + 1:]:
                conductors = re.match(r"\s*(\d+)", str(r[c] or ""))
                if r[d] and conductors:
                    stock.append({"designation": str(r[d]).strip(), "conductors": int(conductors.group(1)),
                                  "shield": str(r[s] or "").strip().lower() == "yes"})
                elif stock:
                    break
            return stock
    return []


def cable_for(wire_count: int, shielded: bool, stock: List[dict]) -> Tuple[Optional[str], Optional[int]]:
    """Stock cable and quantity, following the template's convention: a shielded run uses one multi-conductor shielded
    cable with that many conductors; an unshielded run uses single-conductor wire, one per wire."""
    if shielded:
        match = next((c for c in stock if c["shield"] and c["conductors"] == wire_count), None)
        return (match["designation"], 1) if match else (None, None)
    match = next((c for c in stock if not c["shield"] and c["conductors"] == 1), None)
    return (match["designation"], wire_count) if match else (None, None)


def build_row(inst: dict, stock: List[dict]) -> List[object]:
    """One workbook row (in LIST_COLUMNS order) for a combined-list instrument."""
    voltage, safe = inst.get("voltage"), inst.get("intrinsic_safety")
    cells: Dict[str, object] = {
        "TAG NO": inst.get("name"),
        "Controls Notes": _instrument_notes(inst, "controls") or None,
        "P&IDs Notes": _instrument_notes(inst, "pids") or None,
        "Voltage": None if _is_blank(voltage) else voltage,
        "INTRINSICALLY SAFE": "YES" if safe is True or str(safe).strip().lower() in {"true", "yes"} else "NO",
        "SIGNAL": signal_for(inst.get("name")),
    }
    wire_type = str(inst.get("wire_type") or "").strip().lower()
    count = re.match(r"\s*(\d+)", str(inst.get("wire_count") if inst.get("wire_count") is not None else ""))
    if inst.get("source") != "pids" and count and wire_type in {"shielded", "unshielded"}:
        wires = int(count.group(1))
        shielded = wire_type == "shielded"
        cells["CONNECTION"] = f"{wires} wire shielded" if shielded else f"{wires} wire"
        cells["Cable Designation"], cells["Cable Qty"] = cable_for(wires, shielded, stock)
    return [cells.get(column) for column in LIST_COLUMNS]


# ------------------------------------------------------------------------------------------------- the step
def fill_cable_workbook(
    combined_json: Path,
    output: Path,
    template: Path = DEFAULT_TEMPLATE,
    log: Callable[[str], None] = print,
) -> Path:
    template, combined_json, output = Path(template).resolve(), Path(combined_json).resolve(), Path(output).resolve()
    if not template.is_file():
        raise FileNotFoundError(f"Cable Calculations template not found: {template}")
    if not combined_json.is_file():
        raise FileNotFoundError(f"Combined instrument list not found: {combined_json}")

    combined = json.loads(combined_json.read_text(encoding="utf-8"))
    instruments = [r for r in combined.get("combined_instruments", []) if isinstance(r, dict)]
    if not instruments:
        raise RuntimeError(f"No instruments in {combined_json.name}; nothing to write to the workbook.")
    log(f"Workbook step 1/3: building {len(instruments)} row(s) from the combined list (no model involved)")
    list_sheet, _examples = _read_list_examples(template)  # also confirms the list's column headings are present
    stock = read_stock_cables(template)
    if not stock:
        log("WARNING: no 'IN STOCK CABLES' table found in the template; Cable Designation and Cable Qty stay empty")
    rows = [build_row(inst, stock) for inst in instruments]
    filled = {column: sum(1 for r in rows if r[i] not in (None, "")) for i, column in enumerate(LIST_COLUMNS)}
    log("Workbook step: cells filled per column: " + ", ".join(f"{c} {n}" for c, n in filled.items()))

    log("Workbook step 2/3: writing the rows into a copy of the workbook through Excel (the original stays untouched)")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.unlink(missing_ok=True)
    started = time.time()
    with tempfile.TemporaryDirectory(prefix="cable_step_") as tmp:
        tmp_dir = Path(tmp)
        (tmp_dir / "rows.json").write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
        (tmp_dir / "headers.json").write_text(json.dumps(LIST_COLUMNS), encoding="utf-8")
        # utf-8 with BOM so Windows PowerShell 5.1 reads any non-ASCII text correctly
        (tmp_dir / "fill.ps1").write_text(EXCEL_SCRIPT, encoding="utf-8-sig")
        pid_file = tmp_dir / "excel.pid"
        process = subprocess.Popen(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(tmp_dir / "fill.ps1"),
             "-Template", str(template), "-Output", str(output),
             "-RowsFile", str(tmp_dir / "rows.json"), "-HeadersFile", str(tmp_dir / "headers.json"),
             "-PidFile", str(pid_file)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding="utf-8", errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        with _active_lock:
            _active.append((process, pid_file))
        try:
            stdout, stderr = process.communicate(timeout=EXCEL_TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired:
            _stop(process, pid_file)
            raise RuntimeError(f"Excel did not finish within {EXCEL_TIMEOUT_SECONDS // 60} minutes and was stopped.")
        finally:
            with _active_lock:
                _active[:] = [item for item in _active if item[0] is not process]
    if process.returncode != 0 or not output.is_file():
        raise RuntimeError(f"Excel could not write the workbook: {(stderr or stdout).strip()[-600:]}")
    log(f"Excel: {stdout.strip()}")
    if template.stat().st_mtime > started:
        log(f"WARNING: the original template '{template.name}' changed during this step; check it")

    log("Workbook step 3/3: comparing the copy with the template (takes a moment for a large file)...")
    try:
        inside, outside = compare_workbooks(template, output, list_sheet)
    except Exception as exc:  # noqa: BLE001 - the comparison is a safety net, not a reason to lose the result
        log(f"Could not compare the workbooks: {exc}")
        return output
    log(f"Workbook check: {len(inside)} cell(s) changed inside the instrument list (expected)")
    if outside:
        log(f"WARNING: {len(outside)} cell(s) changed OUTSIDE the instrument list:")
        for line in outside[:MAX_REPORTED_DIFFERENCES]:
            log(f"  {line}")
        if len(outside) > MAX_REPORTED_DIFFERENCES:
            log(f"  ... and {len(outside) - MAX_REPORTED_DIFFERENCES} more")
    else:
        log("Workbook check: nothing outside the instrument list changed")
    return output
