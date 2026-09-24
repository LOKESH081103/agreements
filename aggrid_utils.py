"""
aggrid_utils.py
================
High-performance AgGrid rendering engine for large-scale virtualized pivot tables.

Features:
 - Virtualized DOM scrolling configured for massive datasets (30k+ groups).
 - Pinned left column(s) (`pinned: 'left'`) for hierarchical row labels.
 - Native indentation & group badge/styling support based on `indent` and `is_group` flags.
 - Recursive multi-level nested column header structure support (pandas MultiIndex -> AgGrid column groups).
 - Drop-in replacement for `_render_styled_pivot_matrix` in Streamlit apps.
"""

from __future__ import annotations

import re
from typing import Any

import numpy as np
import pandas as pd
import streamlit as st

try:
    from st_aggrid import AgGrid, ColumnsAutoSizeMode, GridOptionsBuilder, GridUpdateMode, JsCode
    HAS_AGGRID = True
except ImportError:
    HAS_AGGRID = False
    AgGrid = None
    GridOptionsBuilder = None
    JsCode = lambda x: x
    ColumnsAutoSizeMode = None
    GridUpdateMode = None


# --------------------------------------------------------------------------- #
# JavaScript Renderers & Formatters for AgGrid
# --------------------------------------------------------------------------- #
JS_NUMBER_FORMATTER = (
    """
function(params) {
    if (params.value === null || params.value === undefined || isNaN(params.value) || params.value === '') {
        return '-';
    }
    return Number(params.value).toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}
"""
)

JS_COUNT_FORMATTER = (
    """
function(params) {
    if (params.value === null || params.value === undefined || isNaN(params.value) || params.value === '') {
        return '-';
    }
    return Number(params.value).toLocaleString('en-US', {
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    });
}
"""
)

JS_PERCENT_FORMATTER = (
    """
function(params) {
    if (params.value === null || params.value === undefined || isNaN(params.value) || params.value === '') {
        return '-';
    }
    return Number(params.value).toFixed(2) + '%';
}
"""
)

JS_ROW_LABEL_RENDERER = (
    """
function(params) {
    if (!params || !params.data) return params.value || '';
    
    var indent = params.data._indent;
    if (indent === undefined || indent === null) {
        indent = params.data.indent || 0;
    }
    
    var isGroup = params.data._is_group;
    if (isGroup === undefined || isGroup === null) {
        isGroup = params.data.is_group || false;
    }
    
    var isGrandTotal = params.data._is_grand_total || (params.value === 'Grand Total');
    var rawText = params.value || '';
    
    // Clean leading spaces or arrows if still present
    var cleanText = String(rawText).replace(/^\\s+/, '').replace(/^↳\\s*/, '');
    var paddingPx = (indent * 20 + 8) + 'px';
    
    var icon = '';
    if (isGrandTotal) {
        icon = '<span style="margin-right:6px; color:#1e293b; font-weight:700;">∑</span>';
    } else if (isGroup) {
        icon = '<span style="margin-right:6px; color:#3b82f6; font-size:11px;">▼</span>';
    } else if (indent > 0) {
        icon = '<span style="margin-right:6px; color:#64748b; font-size:12px;">↳</span>';
    }
    
    var fontWeight = (isGroup || isGrandTotal) ? '600' : '400';
    var fontColor = isGrandTotal ? '#0f172a' : (isGroup ? '#1e3a8a' : '#334155');
    
    return '<div style="padding-left:' + paddingPx + '; font-weight:' + fontWeight + '; color:' + fontColor + '; display:flex; align-items:center; height:100%;">' +
           icon + '<span>' + cleanText + '</span></div>';
}
"""
)

JS_ROW_STYLE = (
    """
function(params) {
    if (!params || !params.data) return null;
    
    var isGrandTotal = params.data._is_grand_total || (params.data.row_label === 'Grand Total');
    if (isGrandTotal) {
        return {
            'backgroundColor': '#e2e8f0',
            'fontWeight': '700',
            'borderTop': '2px solid #94a3b8',
            'borderBottom': '2px solid #94a3b8'
        };
    }
    
    var isGroup = params.data._is_group || params.data.is_group;
    var indent = params.data._indent || params.data.indent || 0;
    
    if (isGroup) {
        if (indent === 0) {
            return {
                'backgroundColor': '#f1f5f9',
                'fontWeight': '600',
                'borderBottom': '1px solid #cbd5e1'
            };
        } else if (indent === 1) {
            return {
                'backgroundColor': '#f8fafc',
                'fontWeight': '600'
            };
        }
        return {
            'backgroundColor': '#fafafa',
            'fontWeight': '600'
        };
    }
    
    return null;
}
"""
)

JS_TOTAL_COL_STYLE = (
    """
function(params) {
    return {
        'backgroundColor': '#f8fafc',
        'fontWeight': '600',
        'borderLeft': '1px solid #e2e8f0'
    };
}
"""
)


# --------------------------------------------------------------------------- #
# Helper Functions
# --------------------------------------------------------------------------- #
def _parse_row_meta(raw_label: str) -> tuple[int, bool, bool, str]:
    """
    Extract (indent_level, is_group, is_grand_total, clean_label) from a
    hierarchical label string (such as '        ↳ EAST_1' or 'EAST').
    """
    s = str(raw_label)
    lstripped = s.lstrip(" ")
    indent = (len(s) - len(lstripped)) // 4
    is_grand_total = lstripped.strip() in ("Grand Total", "Total")
    is_child = lstripped.startswith("↳ ")
    clean = lstripped[2:].strip() if is_child else lstripped.strip()
    is_group = (not is_child) and (not is_grand_total)
    return indent, is_group, is_grand_total, clean


def _build_nested_col_defs(
    col_tuples: list[tuple[str, ...]],
    field_keys: list[str],
    value_formatter_js: Any,
    total_col_flags: list[bool] | None = None,
) -> list[dict[str, Any]]:
    """
    Recursively transform a list of MultiIndex column tuples into an AgGrid
    nested column definition tree (`children` property).
    """
    n_levels = max(len(t) for t in col_tuples) if col_tuples else 1

    def _group_recursive(
        sub_indices: list[int],
        level: int,
    ) -> list[dict[str, Any]]:
        groups: list[dict[str, Any]] = []
        i = 0
        n = len(sub_indices)
        while i < n:
            curr_idx = sub_indices[i]
            curr_tup = col_tuples[curr_idx]
            curr_label = curr_tup[level] if level < len(curr_tup) else ""

            # Find matching contiguous span sharing the same label at this level
            span = [curr_idx]
            j = i + 1
            while j < n:
                next_idx = sub_indices[j]
                next_tup = col_tuples[next_idx]
                next_label = next_tup[level] if level < len(next_tup) else ""
                if next_label == curr_label:
                    span.append(next_idx)
                    j += 1
                else:
                    break

            # Check if this node is effectively a leaf:
            # - level is the last level
            # - or all deeper levels in this span are empty strings
            is_deepest = (level == n_levels - 1)
            all_remaining_empty = True
            for idx in span:
                for rem_lvl in range(level + 1, len(col_tuples[idx])):
                    if col_tuples[idx][rem_lvl] != "":
                        all_remaining_empty = False
                        break
                if not all_remaining_empty:
                    break

            if is_deepest or (all_remaining_empty and len(span) == 1):
                idx = span[0]
                is_total = bool(total_col_flags[idx]) if total_col_flags and idx < len(total_col_flags) else False
                header_title = curr_label if curr_label != "" else col_tuples[idx][0]
                col_def: dict[str, Any] = {
                    "headerName": header_title,
                    "field": field_keys[idx],
                    "type": "numericColumn",
                    "sortable": True,
                    "filter": True,
                    "resizable": True,
                    "minWidth": 115,
                    "valueFormatter": value_formatter_js,
                }
                if is_total:
                    col_def["cellStyle"] = JsCode(JS_TOTAL_COL_STYLE) if HAS_AGGRID else None
                groups.append(col_def)
            else:
                # Nested parent group
                children = _group_recursive(span, level + 1)
                groups.append({
                    "headerName": curr_label if curr_label != "" else "Group",
                    "children": children,
                    "marryChildren": True,
                })
            i = j
        return groups

    return _group_recursive(list(range(len(col_tuples))), 0)


# --------------------------------------------------------------------------- #
# Main Render Function
# --------------------------------------------------------------------------- #
def display_virtualized_pivot(
    df: pd.DataFrame,
    hierarchy_order: list[str] | str | None = None,
    group_row_flags: list[bool] | None = None,
    total_col_flags: list[bool] | None = None,
    agg_func: str = "sum",
    height: int = 580,
    pinned_row_title: str = "Row Labels",
    **kwargs: Any,
) -> Any:
    """
    Render a high-performance virtualized AgGrid pivot table in Streamlit.

    Drop-in replacement for `_render_styled_pivot_matrix(display_df, group_row_flags, total_col_flags, agg_func)`.

    Parameters
    ----------
    df : pd.DataFrame
        Pivot matrix with hierarchical row labels (index) and optionally MultiIndex columns.
        Can also contain explicit `indent` and `is_group` columns.
    hierarchy_order : list[str] | str | None, optional
        Order or names of hierarchical dimensions. If passed as a list of booleans,
        it is automatically treated as `group_row_flags` for backwards-compatibility.
    group_row_flags : list[bool] | None, optional
        Positional boolean flags indicating whether each row is a group/subtotal row.
    total_col_flags : list[bool] | None, optional
        Positional boolean flags indicating whether each column is a Total/subtotal column.
    agg_func : str, default 'sum'
        Aggregation function ('sum', 'count', 'mean', 'min', 'max', 'percent').
    height : int, default 580
        Fixed height in pixels (essential for DOM virtualization to activate).
    pinned_row_title : str, default 'Row Labels'
        Title for the pinned left row-header column.
    **kwargs
        Additional arguments passed to AgGrid.

    Returns
    -------
    AgGrid return object or None.
    """
    # Defensive check: if hierarchy_order was passed positionally where group_row_flags was expected
    if isinstance(hierarchy_order, list) and hierarchy_order and isinstance(hierarchy_order[0], bool):
        group_row_flags = hierarchy_order
        hierarchy_order = None

    if not HAS_AGGRID:
        st.warning(
            "⚠️ `streamlit-aggrid` is not installed in your Python environment. "
            "Install it via: `pip install streamlit-aggrid`. Falling back to standard `st.dataframe`."
        )
        num_fmt = "{:,.0f}" if agg_func == "count" else ("{:,.2f}%" if agg_func == "percent" else "{:,.2f}")
        try:
            return st.dataframe(df.style.format(num_fmt), use_container_width=True, height=height)
        except Exception:
            return st.dataframe(df, use_container_width=True, height=height)

    # 1. Prepare Data & Hierarchy Metadata
    flat_data: dict[str, Any] = {}
    n_rows = len(df)

    # Extract or resolve row labels and hierarchy attributes
    raw_index_labels = list(df.index)
    has_indent_col = "indent" in df.columns
    has_is_group_col = "is_group" in df.columns

    row_labels: list[str] = []
    indents: list[int] = []
    is_groups: list[bool] = []
    is_grand_totals: list[bool] = []

    for i in range(n_rows):
        raw_val = raw_index_labels[i]
        parsed_indent, parsed_is_group, parsed_gt, clean_val = _parse_row_meta(raw_val)

        # Use explicit columns if provided in df, else parsed metadata
        row_indent = int(df["indent"].iloc[i]) if has_indent_col else parsed_indent
        row_is_group = (
            bool(group_row_flags[i]) if group_row_flags and i < len(group_row_flags)
            else (bool(df["is_group"].iloc[i]) if has_is_group_col else parsed_is_group)
        )

        row_labels.append(clean_val)
        indents.append(row_indent)
        is_groups.append(row_is_group)
        is_grand_totals.append(parsed_gt)

    flat_data["row_label"] = row_labels
    flat_data["_indent"] = indents
    flat_data["_is_group"] = is_groups
    flat_data["_is_grand_total"] = is_grand_totals

    # 2. Prepare Columns & Column Definitions
    is_multi_cols = isinstance(df.columns, pd.MultiIndex)
    raw_cols = list(df.columns)
    col_tuples: list[tuple[str, ...]] = (
        raw_cols if is_multi_cols else [(str(c),) for c in raw_cols]
    )

    # Filter out helper columns if they were part of df.columns
    data_indices: list[int] = []
    field_keys: list[str] = []
    filtered_tuples: list[tuple[str, ...]] = []

    for i, tup in enumerate(col_tuples):
        col_first_name = str(tup[0]).strip()
        if col_first_name in ("indent", "is_group", "_indent", "_is_group", "_is_grand_total"):
            continue
        data_indices.append(i)
        fkey = f"val_col_{i}"
        field_keys.append(fkey)
        filtered_tuples.append(tup)
        # Add values to flat_data with sanitized keys
        flat_data[fkey] = df.iloc[:, i].replace([np.inf, -np.inf], np.nan).fillna("").tolist()

    grid_df = pd.DataFrame(flat_data)

    # Choose value formatter
    if agg_func == "count":
        formatter_js = JsCode(JS_COUNT_FORMATTER)
    elif agg_func == "percent":
        formatter_js = JsCode(JS_PERCENT_FORMATTER)
    else:
        formatter_js = JsCode(JS_NUMBER_FORMATTER)

    # 3. Build Column Definitions Tree
    column_defs: list[dict[str, Any]] = []

    # Pin Left Row Labels Column
    header_name = pinned_row_title
    if df.index.name:
        header_name = str(df.index.name)
    elif hierarchy_order and isinstance(hierarchy_order, (list, tuple)):
        header_name = " / ".join(str(h) for h in hierarchy_order)

    pinned_col_def: dict[str, Any] = {
        "headerName": header_name,
        "field": "row_label",
        "pinned": "left",
        "lockPinned": True,
        "suppressMovable": True,
        "cellRenderer": JsCode(JS_ROW_LABEL_RENDERER),
        "minWidth": 260,
        "resizable": True,
        "sortable": False,  # Hierarchical pre-order walk must remain intact
        "filter": True,
    }
    column_defs.append(pinned_col_def)

    # Data Columns (Flat or Multi-level Nested)
    filtered_total_flags = (
        [total_col_flags[i] for i in data_indices] if total_col_flags and len(total_col_flags) == len(col_tuples)
        else None
    )

    if is_multi_cols and df.columns.nlevels > 1:
        nested_defs = _build_nested_col_defs(
            filtered_tuples, field_keys, formatter_js, filtered_total_flags
        )
        column_defs.extend(nested_defs)
    else:
        for idx, (tup, fkey) in enumerate(zip(filtered_tuples, field_keys)):
            is_total = bool(filtered_total_flags[idx]) if filtered_total_flags else False
            c_def: dict[str, Any] = {
                "headerName": tup[0] if tup else fkey,
                "field": fkey,
                "type": "numericColumn",
                "sortable": True,
                "filter": True,
                "resizable": True,
                "minWidth": 110,
                "valueFormatter": formatter_js,
            }
            if is_total:
                c_def["cellStyle"] = JsCode(JS_TOTAL_COL_STYLE)
            column_defs.append(c_def)

    # Hidden helper fields for row styling & indent lookup
    for helper_key in ("_indent", "_is_group", "_is_grand_total"):
        column_defs.append({
            "field": helper_key,
            "hide": True,
            "suppressColumnsToolPanel": True,
        })

    # 4. Virtualized DOM Scrolling Grid Options (Optimized for 30k+ rows)
    grid_options: dict[str, Any] = {
        # Virtualization controls
        "rowBuffer": 25,                       # Keep small DOM buffer around viewport
        "suppressRowVirtualisation": False,    # CRITICAL: Virtualize DOM rows
        "suppressColumnVirtualisation": False, # Virtualize horizontal columns
        "domLayout": "normal",                 # 'normal' layout with fixed height activates virtual scroll
        "animateRows": False,                  # Disable animations for instant 30k+ row rendering
        "suppressAnimationFrame": False,
        "fastWatch": True,
        "debounceVerticalScrollbar": True,
        "rowModelType": "clientSide",
        # Visual metrics
        "headerHeight": 32,
        "rowHeight": 28,
        "columnDefs": column_defs,
        "defaultColDef": {
            "resizable": True,
            "sortable": True,
            "filter": True,
            "minWidth": 90,
        },
        "getRowStyle": JsCode(JS_ROW_STYLE),
    }

    # 5. Render via AgGrid
    custom_css = {
        ".ag-header-cell-label": {"justify-content": "center"},
        ".ag-header-group-cell-label": {"justify-content": "center", "font-weight": "600"},
        ".ag-theme-alpine .ag-pinned-left-header": {"border-right": "2px solid #cbd5e1 !important"},
        ".ag-theme-alpine .ag-cell-pinned-left": {"border-right": "2px solid #cbd5e1 !important"},
    }

    return AgGrid(
        grid_df,
        gridOptions=grid_options,
        height=height,
        theme="alpine",
        custom_css=custom_css,
        allow_unsafe_jscode=True,
        update_mode=GridUpdateMode.NO_UPDATE if hasattr(GridUpdateMode, "NO_UPDATE") else GridUpdateMode.VALUE_CHANGED,
        columns_auto_size_mode=ColumnsAutoSizeMode.NO_AUTOSIZE if hasattr(ColumnsAutoSizeMode, "NO_AUTOSIZE") else None,
        fit_columns_on_grid_load=False,
        **kwargs,
    )
