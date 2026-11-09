"""Run the optional SlideFill VBA comparison in installed Windows Excel."""

from decimal import Decimal
from fractions import Fraction
import ctypes
import json
import math
import os
from pathlib import Path
import re
import struct
import subprocess
import sys
import tempfile


MAX_CASES = 128
MAX_CASE_OPERANDS = 4096
MAX_BATCH_OPERANDS = 16384
MAX_RESPONSE_BYTES = 128 * 1024
ENGINE_SHA256 = "176b60ab11d57d627c88d369c2b6ec5ee3e51ad31102997e28eea96d1bd6d8f3"
OPERATIONS = frozenset({"sum", "mean", "count", "growth", "points", "round"})
ERROR_CODES = frozenset({2007, 2015, 2036})
STATUSES = frozenset({"executed", "invalid", "unavailable", "blocked", "failed"})


def _resource_base():
    return Path(getattr(sys, "_MEIPASS", Path(__file__).parent))


def bundled_engine_path():
    return _resource_base() / "assets" / "slidefill-vba.xlsm"


def _encode_operand(value):
    if value is None:
        return None
    if type(value) not in (int, float, Decimal, Fraction):
        raise ValueError("VBA comparison accepts only numeric operands and blanks.")
    try:
        number = float(value)
    except (ValueError, OverflowError) as exc:
        raise ValueError("The operand is outside the Excel Double range.") from exc
    if not math.isfinite(number) or (number == 0 and value != 0):
        raise ValueError("The operand is outside the Excel Double range.")
    if type(value) is int:
        exact_integer = value == int(number)
    elif type(value) is Fraction and value.denominator == 1:
        exact_integer = value.numerator == int(number)
    elif type(value) is Decimal and value == value.to_integral_value():
        exact_integer = value == Decimal.from_float(number)
    else:
        exact_integer = True
    if not exact_integer:
        raise ValueError("The integer cannot be represented exactly as an Excel Double.")
    return struct.pack("<d", number).hex()


def _encode_cases(cases):
    if not isinstance(cases, (list, tuple)) or not 1 <= len(cases) <= MAX_CASES:
        raise ValueError(f"Expected 1 to {MAX_CASES} comparison cases.")
    ids = []
    payload = []
    operand_count = 0
    for case in cases:
        if not isinstance(case, dict):
            raise ValueError("Expected a comparison case object.")
        case_id = case.get("id")
        if not isinstance(case_id, str) or not 1 <= len(case_id) <= 120 or case_id in ids:
            raise ValueError("Expected distinct, bounded comparison identifiers.")
        operation = case.get("operation")
        if not isinstance(operation, str) or operation not in OPERATIONS:
            raise ValueError("Unsupported VBA comparison operation.")
        digits = case.get("digits", 2)
        if type(digits) is not int or not -15 <= digits <= 15:
            raise ValueError("Excel rounding places must be an integer from -15 to 15.")
        operands = case.get("operands")
        if not isinstance(operands, (list, tuple)) or len(operands) > MAX_CASE_OPERANDS:
            raise ValueError("Too many operands in the VBA comparison case.")
        operand_count += max(1, len(operands))
        if operand_count > MAX_BATCH_OPERANDS:
            raise ValueError("Too many operands in the VBA comparison batch.")
        ids.append(case_id)
        # Identifiers stay in the Python process. The host gets only typed operands.
        payload.append({"operation": operation, "digits": digits,
                        "operands": [_encode_operand(value) for value in operands] or [None]})
    return ids, payload


def _response(response_path, ids, exit_code):
    if not 0 < response_path.stat().st_size <= MAX_RESPONSE_BYTES:
        raise ValueError("Invalid size for the VBA host response.")
    response = json.loads(response_path.read_text(encoding="utf-8"))
    if (not isinstance(response, dict) or not isinstance(response.get("status"), str)
            or response["status"] not in STATUSES):
        raise ValueError("Invalid VBA host response.")
    status = response["status"]
    if status != "executed":
        detail = response.get("detail")
        if exit_code == 0 or not isinstance(detail, str) or not detail:
            raise ValueError("Invalid VBA host failure response.")
        return {"status": status, "cases": [], "detail": detail[:1000]}
    results = response.get("cases")
    if exit_code != 0 or not isinstance(results, list) or len(results) != len(ids):
        raise ValueError("The VBA host returned an incomplete batch.")
    cases = []
    for case_id, row in zip(ids, results):
        if not isinstance(row, dict):
            raise ValueError("Invalid VBA case response.")
        if row.get("kind") == "error":
            code = row.get("error_code")
            if type(code) is not int or code not in ERROR_CODES:
                raise ValueError("Unknown VBA error result.")
            cases.append({"id": case_id, "kind": "error", "error_code": code})
        elif row.get("kind") == "number":
            bits = row.get("double_hex")
            if not isinstance(bits, str) or not re.fullmatch(r"[0-9a-fA-F]{16}", bits):
                raise ValueError("Invalid VBA Double representation.")
            number = struct.unpack("<d", bytes.fromhex(bits))[0]
            if not math.isfinite(number):
                raise ValueError("The VBA result is not finite.")
            cases.append({"id": case_id, "kind": "number", "value": number,
                          "decimal": repr(number), "double_hex": bits.upper()})
        else:
            raise ValueError("Invalid VBA case kind.")
    return {"status": "executed", "cases": cases}


def _powershell_path():
    directory = ctypes.create_unicode_buffer(32768)
    length = ctypes.windll.kernel32.GetSystemDirectoryW(directory, len(directory))
    if not 0 < length < len(directory):
        raise OSError("Windows PowerShell could not be located.")
    return Path(directory.value) / "WindowsPowerShell" / "v1.0" / "powershell.exe"


def run(engine_path, expected_sha256, cases, *, timeout=30):
    """Compare bounded numeric cases with the exact pinned, read-only engine file."""
    try:
        if not isinstance(expected_sha256, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", expected_sha256):
            raise ValueError("Expected the reviewed VBA engine's SHA-256 digest.")
        ids, payload = _encode_cases(cases)
        engine = os.path.abspath(os.fspath(engine_path))
        if Path(engine).suffix.lower() != ".xlsm":
            raise ValueError("Select the SlideFill VBA engine workbook.")
        if type(timeout) not in (int, float) or not math.isfinite(timeout) or not 0 < timeout <= 120:
            raise ValueError("Invalid VBA host timeout.")
    except (ValueError, TypeError, OSError, ArithmeticError) as exc:
        return {"status": "invalid", "cases": [], "detail": str(exc)}
    if os.name != "nt":
        return {"status": "unavailable", "cases": [], "detail": "The VBA comparison requires Windows Excel."}

    host = _resource_base() / "native_vba_host.ps1"
    try:
        with tempfile.TemporaryDirectory(prefix="SlideFill-VBA-") as folder:
            request = Path(folder) / "request.json"
            response = Path(folder) / "response.json"
            request.write_text(json.dumps({"version": 1, "cases": payload}, separators=(",", ":")),
                               encoding="ascii")
            command = [str(_powershell_path()), "-NoLogo", "-NoProfile", "-NonInteractive", "-STA",
                       "-File", str(host), "-EnginePath", engine, "-ExpectedSha256", expected_sha256,
                       "-RequestPath", str(request), "-ResponsePath", str(response)]
            completed = subprocess.run(command, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                       stderr=subprocess.DEVNULL, timeout=timeout,
                                       creationflags=subprocess.CREATE_NO_WINDOW)
            return _response(response, ids, completed.returncode)
    except subprocess.TimeoutExpired:
        return {"status": "timeout", "cases": [], "detail": "Excel did not complete the VBA comparison in time."}
    except (OSError, ValueError) as exc:
        return {"status": "failed", "cases": [], "detail": str(exc)}
