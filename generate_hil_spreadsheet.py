#!/usr/bin/env python3
"""Generate HIL testing spreadsheet for all I2C Wippersnapper components.

Creates an Excel workbook with:
  Sheet 1 - Component Matrix: all components, addresses, vendors, conflicts
  Sheet 2 - HIL Mux Layout: conflict-free channel assignments for testing
  Sheet 3 - Address Conflicts: per-address overlap summary

Hardware config:
  - 8-ch TCA9548A mux @ 0x77  (all 3 addr pads bridged)
  - 4-ch TCA9544A mux @ 0x71  (A0 bridged)
  - Reserved addresses on every channel: {0x77, 0x71}
"""

import json
import os
from pathlib import Path
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Hardware config ──────────────────────────────────────────────────────────
MUXES = [
    {"type": "TCA9548A", "address": 0x77, "channels": 8, "label": "8ch"},
    {"type": "TCA9544A", "address": 0x71, "channels": 4, "label": "4ch"},
]
MUX_RESERVED = {m["address"] for m in MUXES}
TOTAL_MUX_CHANNELS = sum(m["channels"] for m in MUXES)  # 12


# ── Colours ──────────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
CONFLICT_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
UNIQUE_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
UNPUBLISHED_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
MUX_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
MUX2_HEADER_FILL = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
CHANNEL_FILLS_MUX1 = [
    PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),  # Ch0
    PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),  # Ch1
    PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),  # Ch2
    PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),  # Ch3
    PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),  # Ch4
    PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),  # Ch5
    PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),  # Ch6
    PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),  # Ch7
]
CHANNEL_FILLS_MUX2 = [
    PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),  # Ch0
    PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),  # Ch1
    PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),  # Ch2
    PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),  # Ch3
]
NON_DEFAULT_FILL = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")  # orange — assigned != default
NON_DEFAULT_FONT = Font(bold=True, color="833C0B")
NOMUX_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
NOMUX_HEADER_FILL = PatternFill(start_color="C55A11", end_color="C55A11", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def load_components(base_dir):
    """Load all I2C component definitions."""
    components = []
    i2c_dir = Path(base_dir) / "components" / "i2c"
    for defn in sorted(i2c_dir.glob("*/definition.json")):
        with open(defn, "r", encoding="utf-8") as f:
            data = json.load(f)
        name = defn.parent.name
        addrs = [int(a, 16) for a in data.get("i2cAddresses", [])]  # preserve definition order (default first)
        usable = [a for a in addrs if a not in MUX_RESERVED]
        components.append({
            "dir": name,
            "displayName": data.get("displayName", name),
            "vendor": data.get("vendorName", ""),
            "all_addresses": addrs,
            "usable_addresses": usable,
            "published": data.get("published", True),
            "sensors": [
                s["sensorType"] if isinstance(s, dict) else s
                for s in data.get("subcomponents", [])
            ],
        })
    components.sort(key=lambda c: (c["all_addresses"][0] if c["all_addresses"] else 0xFF, c["dir"]))
    return components


def build_address_map(components):
    """Map address -> list of component dicts (using all_addresses)."""
    addr_map = defaultdict(list)
    for comp in components:
        for a in comp["all_addresses"]:
            addr_map[a].append(comp)
    return addr_map


def find_conflicts(components, addr_map):
    """For each component, find which other components share any address."""
    conflicts = {}
    for comp in components:
        peers = set()
        for a in comp["all_addresses"]:
            for other in addr_map[a]:
                if other["dir"] != comp["dir"]:
                    peers.add(other["dir"])
        conflicts[comp["dir"]] = sorted(peers)
    return conflicts


# ── Channel assignment with single-address picking ───────────────────────────

def assign_channels(components):
    """
    Assign each component to a channel AND pick ONE specific address.

    Strategy:
      1. Put components with truly unique addresses on the direct bus (ch 0).
         "Truly unique" = the component has a usable address that NO other
         component lists in its usable set.  These are free on the direct bus
         because they block an address nobody else needs.
      2. Everything else goes onto mux channels (1-12) via greedy colouring.
         Most-constrained-first (fewest usable addresses).

    Channel layout:
      0           = direct bus (no mux) — always visible
      1  .. 8     = 8ch TCA9548A @ 0x77, channels 0-7
      9  .. 12    = 4ch TCA9544A @ 0x71, channels 0-3

    Constraints:
      - Picked address must not be in MUX_RESERVED
      - No two components on the same channel share a picked address
      - No muxed component's picked address clashes with any direct-bus
        component's picked address (direct bus is always visible)

    Returns (assignment, picked_addr, channel_addrs).
    """
    n_channels = 1 + TOTAL_MUX_CHANNELS  # 13

    channel_addrs = defaultdict(set)  # ch -> set of picked addresses
    direct_addrs = set()              # mirror of channel_addrs[0]
    assignment = {}                   # comp_dir -> channel
    picked_addr = {}                  # comp_dir -> int address

    # ── Build "who else wants this address?" map ──
    addr_users = defaultdict(set)  # addr -> set of comp indices
    for i, comp in enumerate(components):
        for a in comp["usable_addresses"]:
            addr_users[a].add(i)

    # ── Phase 1: direct bus — truly-unique addresses ──
    # An address is "unique" if exactly one component lists it as usable.
    placed_indices = set()
    for i, comp in enumerate(components):
        usable = comp["usable_addresses"]
        if not usable:
            continue
        # Find addresses where this is the ONLY user
        unique_addrs = [a for a in usable if len(addr_users[a]) == 1]
        if unique_addrs:
            # Pick the first unique address
            addr = unique_addrs[0]
            channel_addrs[0].add(addr)
            direct_addrs.add(addr)
            assignment[comp["dir"]] = 0
            picked_addr[comp["dir"]] = addr
            placed_indices.add(i)

    # ── Phase 2: mux channels for everything else ──
    remaining = [i for i in range(len(components)) if i not in placed_indices]
    # Sort: fewest usable addresses first (most constrained)
    remaining.sort(key=lambda i: (len(components[i]["usable_addresses"]),
                                   components[i]["dir"]))

    for i in remaining:
        comp = components[i]
        usable = comp["usable_addresses"]
        placed = False

        if not usable:
            assignment[comp["dir"]] = -1
            picked_addr[comp["dir"]] = None
            continue

        # Try mux channels 1..12, then direct bus (0) as last resort
        channel_order = list(range(1, n_channels)) + [0]
        for ch in channel_order:
            for addr in usable:
                # Already taken on this channel?
                if addr in channel_addrs[ch]:
                    continue
                # Mux channel: can't clash with direct bus
                if ch > 0 and addr in direct_addrs:
                    continue
                # Direct bus: can't clash with any mux channel
                if ch == 0:
                    if any(addr in channel_addrs[mch]
                           for mch in range(1, n_channels)):
                        continue

                channel_addrs[ch].add(addr)
                if ch == 0:
                    direct_addrs.add(addr)
                assignment[comp["dir"]] = ch
                picked_addr[comp["dir"]] = addr
                placed = True
                break
            if placed:
                break

        if not placed:
            assignment[comp["dir"]] = -1
            picked_addr[comp["dir"]] = None

    return assignment, picked_addr, channel_addrs


def channel_label(ch):
    """Human-readable label for a channel number."""
    if ch == 0:
        return "Direct Bus (No Mux)"
    elif ch <= 8:
        m = MUXES[0]
        return f"{m['type']} (0x{m['address']:02X}) Ch{ch - 1}"
    else:
        m = MUXES[1]
        return f"{m['type']} (0x{m['address']:02X}) Ch{ch - 9}"


def channel_short_label(ch):
    """Short label for JSON export."""
    if ch == 0:
        return "direct"
    elif ch <= 8:
        return f"8ch_mux_ch{ch - 1}"
    else:
        return f"4ch_mux_ch{ch - 9}"


def channel_fill(ch):
    """Background colour for a channel."""
    if ch == 0:
        return NOMUX_FILL
    elif ch <= 8:
        return CHANNEL_FILLS_MUX1[(ch - 1) % len(CHANNEL_FILLS_MUX1)]
    else:
        return CHANNEL_FILLS_MUX2[(ch - 9) % len(CHANNEL_FILLS_MUX2)]


def channel_header_fill(ch):
    """Header colour for a channel."""
    if ch == 0:
        return NOMUX_HEADER_FILL
    elif ch <= 8:
        return MUX_HEADER_FILL
    else:
        return MUX2_HEADER_FILL


# ── Spreadsheet generation ───────────────────────────────────────────────────

def write_sheet1(ws, components, addr_map, conflicts):
    """Component Matrix + Addresses sheet."""
    ws.title = "Component Matrix"

    all_addrs = sorted(set(a for c in components for a in c["all_addresses"]))

    headers = [
        "Component", "Display Name", "Vendor", "Published",
        "I2C Addresses", "# Addrs", "Usable (excl mux)",
        "Conflicts With", "# Conflicts", "Sensor Types",
    ]
    for a in all_addrs:
        h = f"0x{a:02X}"
        if a in MUX_RESERVED:
            h += " MUX"
        headers.append(h)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER

    for row_idx, comp in enumerate(components, 2):
        conflict_list = conflicts.get(comp["dir"], [])
        is_conflicted = len(conflict_list) > 0

        vals = [
            comp["dir"],
            comp["displayName"],
            comp["vendor"],
            "Yes" if comp["published"] else "NO",
            ", ".join(f"0x{a:02X}" for a in comp["all_addresses"]),
            len(comp["all_addresses"]),
            ", ".join(f"0x{a:02X}" for a in comp["usable_addresses"]),
            ", ".join(conflict_list) if conflict_list else "None",
            len(conflict_list),
            ", ".join(comp["sensors"]) if comp["sensors"] else "",
        ]

        for col, v in enumerate(vals, 1):
            cell = ws.cell(row=row_idx, column=col, value=v)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if not comp["published"]:
                cell.fill = UNPUBLISHED_FILL
            elif is_conflicted and col in (5, 8):
                cell.fill = CONFLICT_FILL
            elif not is_conflicted and col == 5:
                cell.fill = UNIQUE_FILL

        # Address heatmap columns
        addr_col_start = len(vals) + 1
        for ai, a in enumerate(all_addrs):
            col = addr_col_start + ai
            cell = ws.cell(row=row_idx, column=col)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
            if a in comp["all_addresses"]:
                n_sharing = len(addr_map[a])
                cell.value = n_sharing
                if a in MUX_RESERVED:
                    cell.fill = PatternFill(start_color="BF8F00", end_color="BF8F00", fill_type="solid")
                    cell.font = Font(bold=True, color="FFFFFF")
                elif n_sharing > 1:
                    cell.fill = CONFLICT_FILL
                    cell.font = Font(bold=True, color="9C0006")
                else:
                    cell.fill = UNIQUE_FILL
                    cell.font = Font(color="006100")

    col_widths = {1: 18, 2: 28, 3: 26, 4: 10, 5: 40, 6: 8, 7: 36, 8: 50, 9: 10, 10: 30}
    for c, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    for ai in range(len(all_addrs)):
        ws.column_dimensions[get_column_letter(len(col_widths) + 1 + ai)].width = 6.5

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(components) + 1}"


def write_sheet2(ws, components, assignment, picked_addr, channel_addrs):
    """HIL Mux Layout sheet."""
    ws.title = "HIL Mux Layout"

    n_channels = 1 + TOTAL_MUX_CHANNELS

    # Organise by channel
    channels = defaultdict(list)
    unplaceable = []
    for comp in components:
        ch = assignment.get(comp["dir"], -1)
        if ch < 0:
            unplaceable.append(comp)
        else:
            channels[ch].append(comp)

    for ch in channels:
        channels[ch].sort(key=lambda c: (picked_addr.get(c["dir"], 0) or 0, c["dir"]))

    # ── Summary section ──
    row = 1
    ws.cell(row=row, column=1, value="HIL I2C Mux Testing Layout").font = Font(
        name="Calibri", bold=True, size=14
    )
    row += 1
    ws.cell(row=row, column=1, value=f"Total components: {len(components)}")
    row += 1
    placed = sum(len(v) for v in channels.values())
    ws.cell(row=row, column=1, value=f"Placed: {placed}  |  Unplaceable: {len(unplaceable)}")
    row += 1

    for m in MUXES:
        ws.cell(row=row, column=1,
                value=f"{m['label']}: {m['type']} @ 0x{m['address']:02X} — {m['channels']} channels")
        row += 1

    ws.cell(row=row, column=1,
            value=f"Reserved addresses (mux chips): {', '.join(f'0x{a:02X}' for a in sorted(MUX_RESERVED))}")
    row += 1

    # Channel utilisation summary
    row += 1
    ws.cell(row=row, column=1, value="Channel Summary").font = Font(bold=True, size=11)
    row += 1
    sum_headers = ["Channel", "Label", "Components", "Addresses Used"]
    for hi, h in enumerate(sum_headers):
        cell = ws.cell(row=row, column=hi + 1, value=h)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        cell.border = THIN_BORDER
    row += 1
    for ch in range(n_channels):
        comps_in = channels.get(ch, [])
        addrs_in = channel_addrs.get(ch, set())
        vals = [ch, channel_label(ch), len(comps_in),
                ", ".join(f"0x{a:02X}" for a in sorted(addrs_in))]
        for hi, v in enumerate(vals):
            cell = ws.cell(row=row, column=hi + 1, value=v)
            cell.border = THIN_BORDER
            cell.fill = channel_fill(ch)
        row += 1

    # Unplaceable warnings
    if unplaceable:
        row += 1
        cell = ws.cell(row=row, column=1,
                       value="UNPLACEABLE — all addresses blocked by mux reserved addresses:")
        cell.font = Font(bold=True, color="CC0000")
        row += 1
        for comp in unplaceable:
            ws.cell(row=row, column=1,
                    value=f"  {comp['dir']} ({comp['displayName']}) — "
                          f"addresses: {', '.join(f'0x{a:02X}' for a in comp['all_addresses'])}")
            row += 1

    row += 2

    # ── Channel blocks side-by-side ──
    BLOCK_WIDTH = 6
    block_headers = ["#", "Component", "Display Name", "Assigned Addr", "All Addrs", "Vendor"]
    layout_start_row = row

    active_channels = sorted(ch for ch in range(n_channels) if channels.get(ch))

    for block_idx, ch in enumerate(active_channels):
        col_offset = block_idx * BLOCK_WIDTH
        comps_in = channels.get(ch, [])

        # Header
        r = layout_start_row
        lbl = channel_label(ch)
        for hi in range(BLOCK_WIDTH):
            c = col_offset + hi + 1
            cell = ws.cell(row=r, column=c, value=lbl if hi == 0 else "")
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = channel_header_fill(ch)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=r, start_column=col_offset + 1,
                       end_row=r, end_column=col_offset + BLOCK_WIDTH)

        # Info
        r += 1
        info = f"{len(comps_in)} components"
        for hi in range(BLOCK_WIDTH):
            c = col_offset + hi + 1
            cell = ws.cell(row=r, column=c, value=info if hi == 0 else "")
            cell.font = Font(italic=True, size=9)
            cell.fill = channel_fill(ch)
            cell.border = THIN_BORDER
        ws.merge_cells(start_row=r, start_column=col_offset + 1,
                       end_row=r, end_column=col_offset + BLOCK_WIDTH)

        # Column headers
        r += 1
        for hi, hdr in enumerate(block_headers):
            c = col_offset + hi + 1
            cell = ws.cell(row=r, column=c, value=hdr)
            cell.font = Font(bold=True, size=10)
            cell.fill = channel_fill(ch)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")

        # Component rows
        for ci, comp in enumerate(comps_in):
            r += 1
            pa = picked_addr.get(comp["dir"])
            default_addr = comp["all_addresses"][0] if comp["all_addresses"] else None
            is_non_default = pa is not None and pa != default_addr
            vals = [
                ci + 1,
                comp["dir"],
                comp["displayName"],
                f"0x{pa:02X}" if pa is not None else "?",
                ", ".join(f"0x{a:02X}" for a in comp["all_addresses"]),
                comp["vendor"],
            ]
            for hi, v in enumerate(vals):
                c = col_offset + hi + 1
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if not comp["published"]:
                    cell.fill = UNPUBLISHED_FILL
                elif is_non_default and hi == 3:  # Assigned Addr column
                    cell.fill = NON_DEFAULT_FILL
                    cell.font = NON_DEFAULT_FONT

        # Column widths
        ws.column_dimensions[get_column_letter(col_offset + 1)].width = 4
        ws.column_dimensions[get_column_letter(col_offset + 2)].width = 18
        ws.column_dimensions[get_column_letter(col_offset + 3)].width = 24
        ws.column_dimensions[get_column_letter(col_offset + 4)].width = 12
        ws.column_dimensions[get_column_letter(col_offset + 5)].width = 30
        ws.column_dimensions[get_column_letter(col_offset + 6)].width = 22

    # ── Linear ordered list for JSON export ──
    row2 = layout_start_row
    linear_col = len(active_channels) * BLOCK_WIDTH + 2

    ws.cell(row=row2, column=linear_col,
            value="Ordered Layout (for JSON / pytest export)").font = Font(bold=True, size=12)
    row2 += 1

    linear_headers = ["Order", "Channel#", "Channel Label", "Component", "Display Name",
                       "Assigned Address", "Default Address", "All Addresses", "Vendor",
                       "Published", "Non-Default?"]
    for hi, hdr in enumerate(linear_headers):
        cell = ws.cell(row=row2, column=linear_col + hi, value=hdr)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    row2 += 1

    order_num = 1
    for ch in range(n_channels):
        for comp in channels.get(ch, []):
            pa = picked_addr.get(comp["dir"])
            default_addr = comp["all_addresses"][0] if comp["all_addresses"] else None
            is_non_default = pa is not None and pa != default_addr
            vals = [
                order_num, ch, channel_short_label(ch),
                comp["dir"], comp["displayName"],
                f"0x{pa:02X}" if pa is not None else "?",
                f"0x{default_addr:02X}" if default_addr is not None else "?",
                ", ".join(f"0x{a:02X}" for a in comp["all_addresses"]),
                comp["vendor"],
                "yes" if comp["published"] else "no",
                "NON-DEFAULT" if is_non_default else "",
            ]
            for hi, v in enumerate(vals):
                cell = ws.cell(row=row2, column=linear_col + hi, value=v)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.fill = channel_fill(ch)
                if is_non_default and hi in (5, 10):  # Assigned Address & Non-Default columns
                    cell.fill = NON_DEFAULT_FILL
                    cell.font = NON_DEFAULT_FONT
            order_num += 1
            row2 += 1

    widths = [6, 9, 16, 18, 28, 14, 14, 36, 26, 9, 14]
    for wi, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(linear_col + wi)].width = w

    ws.freeze_panes = "A1"


def write_sheet3(ws, components, addr_map):
    """Address conflict summary sheet."""
    ws.title = "Address Conflicts"

    all_addrs = sorted(set(a for c in components for a in c["all_addresses"]))

    headers = ["Address", "Mux?", "# Components", "Components"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = THIN_BORDER

    row = 2
    for a in all_addrs:
        comps_at = addr_map[a]
        n = len(comps_at)
        is_mux = a in MUX_RESERVED
        ws.cell(row=row, column=1, value=f"0x{a:02X}").border = THIN_BORDER
        mux_cell = ws.cell(row=row, column=2, value="MUX" if is_mux else "")
        mux_cell.border = THIN_BORDER
        cell_n = ws.cell(row=row, column=3, value=n)
        cell_n.border = THIN_BORDER
        cell_comps = ws.cell(
            row=row, column=4,
            value=", ".join(f"{c['dir']} ({c['displayName']})" for c in comps_at),
        )
        cell_comps.border = THIN_BORDER
        cell_comps.alignment = Alignment(wrap_text=True, vertical="top")

        if is_mux:
            for c in range(1, 5):
                ws.cell(row=row, column=c).fill = PatternFill(
                    start_color="BF8F00", end_color="BF8F00", fill_type="solid")
            mux_cell.font = Font(bold=True, color="FFFFFF")
        elif n > 1:
            ws.cell(row=row, column=1).fill = CONFLICT_FILL
            cell_n.fill = CONFLICT_FILL
            cell_n.font = Font(bold=True, color="9C0006")
        else:
            ws.cell(row=row, column=1).fill = UNIQUE_FILL
        row += 1

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 120
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{row - 1}"


def main():
    base_dir = Path(__file__).parent
    components = load_components(base_dir)
    addr_map = build_address_map(components)
    conflicts = find_conflicts(components, addr_map)

    print(f"Loaded {len(components)} I2C components")
    print(f"Unique I2C addresses: {len(set(a for c in components for a in c['all_addresses']))}")
    print(f"Mux reserved addresses: {', '.join(f'0x{a:02X}' for a in sorted(MUX_RESERVED))}")
    print()

    for m in MUXES:
        print(f"  {m['label']}: {m['type']} @ 0x{m['address']:02X} ({m['channels']} channels)")
    print(f"  Total mux channels: {TOTAL_MUX_CHANNELS}")
    print()

    assignment, picked_addr, channel_addrs = assign_channels(components)

    n_channels = 1 + TOTAL_MUX_CHANNELS
    placed = sum(1 for v in assignment.values() if v >= 0)
    unplaced = sum(1 for v in assignment.values() if v < 0)
    print(f"Placed: {placed}  |  Unplaceable: {unplaced}")

    for ch in range(n_channels):
        comps_in = [c for c in components if assignment.get(c["dir"]) == ch]
        if not comps_in:
            continue
        addrs_in = channel_addrs.get(ch, set())
        print(f"  {channel_label(ch)}: {len(comps_in)} components, "
              f"addrs: {', '.join(f'0x{a:02X}' for a in sorted(addrs_in))}")

    unplaced_comps = [c for c in components if assignment.get(c["dir"], -1) < 0]
    if unplaced_comps:
        print("\nUNPLACEABLE:")
        for c in unplaced_comps:
            print(f"  {c['dir']} — all addresses blocked by mux: "
                  f"{', '.join(f'0x{a:02X}' for a in c['all_addresses'])}")

    # Verify no conflicts
    errors = 0
    direct_addrs = channel_addrs.get(0, set())
    for ch in range(n_channels):
        seen = {}
        comps_in = [c for c in components if assignment.get(c["dir"]) == ch]
        for comp in comps_in:
            pa = picked_addr[comp["dir"]]
            if pa in seen:
                print(f"BUG: channel {ch} addr 0x{pa:02X} used by {seen[pa]} AND {comp['dir']}")
                errors += 1
            seen[pa] = comp["dir"]
            if ch > 0 and pa in direct_addrs:
                print(f"BUG: mux ch{ch} {comp['dir']} @ 0x{pa:02X} conflicts with direct bus")
                errors += 1
    if errors == 0:
        print("\nVERIFIED: zero address conflicts across all channels")
    else:
        print(f"\n{errors} BUGS FOUND")

    # Create workbook
    wb = Workbook()
    ws1 = wb.active
    write_sheet1(ws1, components, addr_map, conflicts)

    ws2 = wb.create_sheet()
    write_sheet2(ws2, components, assignment, picked_addr, channel_addrs)

    ws3 = wb.create_sheet()
    write_sheet3(ws3, components, addr_map)

    out_path = base_dir / "hil_i2c_components.xlsx"
    wb.save(out_path)
    print(f"\nSpreadsheet saved to: {out_path}")


if __name__ == "__main__":
    main()
