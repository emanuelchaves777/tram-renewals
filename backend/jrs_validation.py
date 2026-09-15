"""
jrs_validation.py — Validates a JRS value against the IBM Consulting Taxonomy.

Logic (confirmed from M389 file inspection):

  Tab "Consulting Short List"  →  Column A = "JR/S"
    If value found here → ACTIVE

  Tab "Change Log"  →  Column A = "Previous Value (as of prior publication)"
                        Column B = "Current Value (effective this publication)"
                        Column C = "Action taken"  (Add / Delete / Rename)
    If value found in Column A:
      - One row found, Column C = "Delete" → DELETED (no replacement)
      - One row found, Column C = "Rename" → ONE_REPLACEMENT (use Column B)
      - Multiple rows found               → MULTI_REPLACEMENT (user must choose)

  If not found in either tab → NOT_FOUND (escalate to CSP)

Outcomes:
  "active"           — JRS is current and valid
  "one_replacement"  — deleted/renamed; exactly one replacement available
  "multi_replacement"— deleted; multiple replacements; user must choose
  "deleted"          — deleted with no replacement
  "not_found"        — not in taxonomy at all; escalate
"""
import io
from typing import Any

import openpyxl

from config import (
    TAXONOMY_ACTIVE_TAB,
    TAXONOMY_ACTIVE_COL,
    TAXONOMY_CHANGELOG_TAB,
    TAXONOMY_PREV_COL,
    TAXONOMY_NEW_COL,
    TAXONOMY_ACTION_COL,
)


# ── In-memory taxonomy cache ──────────────────────────────────────────────────
_taxonomy_cache: dict | None = None
_taxonomy_publication: str = "not loaded"


def load_taxonomy_from_bytes(file_bytes: bytes) -> dict:
    """
    Parse a taxonomy .xlsx from raw bytes and cache it.
    Replaces the previous Box-based load_taxonomy(client).
    """
    global _taxonomy_cache, _taxonomy_publication

    active_set, changelog = _parse_taxonomy(file_bytes)

    _taxonomy_cache = {
        "active":    active_set,   # set of lowercase JRS strings
        "changelog": changelog,    # list of {prev, new, action} dicts
    }
    _taxonomy_publication = "loaded"
    return _taxonomy_cache


def get_taxonomy_cache() -> dict | None:
    return _taxonomy_cache


def validate_jrs(jrs_value: str) -> dict:
    """
    Validate a single JRS value.
    Returns a result dict:
      {
        "outcome":      str,    # active | one_replacement | multi_replacement | deleted | not_found
        "input":        str,    # original value
        "replacements": [str],  # candidate replacements (empty for active/not_found)
        "message":      str,    # human-readable explanation
        "publication":  str,    # taxonomy publication label
      }
    """
    if _taxonomy_cache is None:
        return {
            "outcome":      "not_loaded",
            "input":        jrs_value,
            "replacements": [],
            "message":      "Taxonomy not loaded. Click 'Refresh Taxonomy' first.",
            "publication":  _taxonomy_publication,
        }

    key = jrs_value.strip().lower()

    # ── Step 1: Check active list ─────────────────────────────────────────────
    if key in _taxonomy_cache["active"]:
        return {
            "outcome":      "active",
            "input":        jrs_value,
            "replacements": [],
            "message":      f"✓ '{jrs_value}' is active in the current taxonomy.",
            "publication":  _taxonomy_publication,
        }

    # ── Step 2: Check change log ──────────────────────────────────────────────
    matches = [
        entry for entry in _taxonomy_cache["changelog"]
        if entry["prev"].strip().lower() == key
    ]

    if not matches:
        return {
            "outcome":      "not_found",
            "input":        jrs_value,
            "replacements": [],
            "message":      (
                f"'{jrs_value}' was not found in the active list or the change log. "
                "Escalate to the CSP team for this contractor's sector."
            ),
            "publication":  _taxonomy_publication,
        }

    # Collect unique non-empty replacements
    replacements = list({
        entry["new"] for entry in matches
        if entry["new"] and entry["new"].strip()
    })
    actions = {entry["action"].strip().lower() for entry in matches}

    if not replacements:
        return {
            "outcome":      "deleted",
            "input":        jrs_value,
            "replacements": [],
            "message":      (
                f"'{jrs_value}' has been deleted from the taxonomy with no replacement. "
                "Escalate to the CSP team."
            ),
            "publication":  _taxonomy_publication,
        }

    if len(replacements) == 1:
        return {
            "outcome":      "one_replacement",
            "input":        jrs_value,
            "replacements": replacements,
            "message":      (
                f"'{jrs_value}' has been {'renamed' if 'rename' in actions else 'replaced'}. "
                f"Proposed replacement: '{replacements[0]}'. Review and confirm."
            ),
            "publication":  _taxonomy_publication,
        }

    return {
        "outcome":      "multi_replacement",
        "input":        jrs_value,
        "replacements": replacements,
        "message":      (
            f"'{jrs_value}' has multiple possible replacements. "
            "Select the correct one and confirm with the CSP team."
        ),
        "publication":  _taxonomy_publication,
    }


# ── .xlsx parsing ─────────────────────────────────────────────────────────────

def _parse_taxonomy(file_bytes: bytes) -> tuple[set, list]:
    """
    Parse active set and changelog from taxonomy .xlsx bytes.
    Returns (active_set, changelog_list).
    """
    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes), read_only=True, data_only=True
    )

    active_set = _parse_active_sheet(wb)
    changelog  = _parse_changelog_sheet(wb)

    wb.close()
    return active_set, changelog


def _parse_active_sheet(wb: openpyxl.Workbook) -> set:
    """
    Read 'Consulting Short List' tab.
    Column A = JR/S. Returns a set of lowercase JRS strings.
    """
    if TAXONOMY_ACTIVE_TAB not in wb.sheetnames:
        raise ValueError(
            f"Tab '{TAXONOMY_ACTIVE_TAB}' not found in taxonomy file. "
            f"Available tabs: {wb.sheetnames}"
        )

    ws = wb[TAXONOMY_ACTIVE_TAB]
    active = set()
    jrs_col_idx = None

    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            # Find the column index for "JR/S"
            for col_idx, cell in enumerate(row):
                if cell and str(cell).strip().upper() == TAXONOMY_ACTIVE_COL.upper():
                    jrs_col_idx = col_idx
                    break
            if jrs_col_idx is None:
                raise ValueError(
                    f"Column '{TAXONOMY_ACTIVE_COL}' not found in tab '{TAXONOMY_ACTIVE_TAB}'."
                )
            continue

        if jrs_col_idx < len(row) and row[jrs_col_idx]:
            val = str(row[jrs_col_idx]).strip()
            if val:
                active.add(val.lower())

    return active


def _parse_changelog_sheet(wb: openpyxl.Workbook) -> list:
    """
    Read 'Change Log' tab.
    Returns list of {prev, new, action} dicts.
    """
    if TAXONOMY_CHANGELOG_TAB not in wb.sheetnames:
        raise ValueError(
            f"Tab '{TAXONOMY_CHANGELOG_TAB}' not found in taxonomy file."
        )

    ws = wb[TAXONOMY_CHANGELOG_TAB]
    changelog = []

    # Default to positional columns A=0, B=1, C=2 — the Change Log structure
    # is always Previous Value | Current Value | Action regardless of header text.
    prev_idx = 0
    new_idx  = 1
    action_idx = 2

    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        if row_idx == 0:
            # Try to find columns by header text; fall back to positional if not found
            found_prev = found_new = found_action = None
            for col_idx, cell in enumerate(row):
                if not cell:
                    continue
                label = str(cell).strip().lower()
                if TAXONOMY_PREV_COL.lower() in label:
                    found_prev = col_idx
                elif TAXONOMY_NEW_COL.lower() in label:
                    found_new = col_idx
                elif TAXONOMY_ACTION_COL.lower() in label:
                    found_action = col_idx
            # Only override positional defaults if all three were found by name
            if None not in (found_prev, found_new, found_action):
                prev_idx   = found_prev
                new_idx    = found_new
                action_idx = found_action
            continue

        def _get(idx):
            if idx is not None and idx < len(row) and row[idx]:
                return str(row[idx]).strip()
            return ""

        prev   = _get(prev_idx)
        new    = _get(new_idx)
        action = _get(action_idx)

        if prev:
            changelog.append({"prev": prev, "new": new, "action": action})

    return changelog
