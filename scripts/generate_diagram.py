#!/usr/bin/env python3
"""Generate architecture diagram from Excel inventory.

Reads `Input/Application_Inventory.xlsx` and uses the columns that best
match `From_App`, `To-App`, and `Interface_Name` to build a Graphviz
diagram. Outputs a DOT file plus attempts to render `svg` and `png` into
`Output/Diagrams`.
"""
import os
import sys
from collections import defaultdict
import argparse

try:
    import pandas as pd
except Exception:
    print("pandas is required. Install with: pip install -r requirements.txt")
    raise

try:
    import graphviz
except Exception:
    print("graphviz (python package) is required. Install with: pip install -r requirements.txt")
    raise


EXCEL_PATH = os.path.join("Input", "Application_Inventory.xlsx")
OUT_DIR = os.path.join("Output", "Diagrams")


def find_column(df, candidates):
    cols = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        for k, orig in cols.items():
            if cand in k:
                return orig
    return None


def normalize(value):
    if pd.isna(value):
        return None
    return str(value).strip()


def main():
    parser = argparse.ArgumentParser(description="Generate architecture diagram from Excel")
    parser.add_argument("--focus", help="Only show edges where this node is source or destination (case-insensitive)")
    args = parser.parse_args()
    focus = args.focus.strip() if args.focus else None
    if not os.path.exists(EXCEL_PATH):
        print(f"Excel file not found at {EXCEL_PATH}")
        sys.exit(2)

    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")

    # heuristics for column names
    from_col = find_column(df, ["from_app", "from-app", "from", "from app"])
    to_col = find_column(df, ["to_app", "to-app", "to", "to app", "to_app"])
    iface_col = find_column(df, ["interface_name", "interface-name", "interface", "interface name"])
    # optional columns for styling/direction
    inttype_col = find_column(df, ["int_type", "int-type", "interface_type", "interface-type", "int type", "interface type", "int_type"])
    dir_col = find_column(df, ["direction", "dir", "Direction", "Direction"])

    if not from_col or not to_col or not iface_col:
        print("Could not locate required columns. Found:")
        print("  from_col:", from_col)
        print("  to_col:", to_col)
        print("  iface_col:", iface_col)
        print("Columns present in the sheet:")
        for c in df.columns:
            print(" -", c)
        sys.exit(3)

    # collect unique interfaces and optional styling
    edges = []
    for _, row in df.iterrows():
        src = normalize(row[from_col])
        dst = normalize(row[to_col])
        iface = normalize(row[iface_col])
        if not src or not dst or not iface:
            continue

        inttype = normalize(row[inttype_col]) if inttype_col else None
        raw_direction = row[dir_col] if dir_col else None

        # normalize direction to integer when possible (handle numeric types and '2.0')
        dir_val = None
        if raw_direction is not None and not pd.isna(raw_direction):
            try:
                # handle numeric types and strings like '2.0'
                dir_val = int(float(raw_direction))
            except Exception:
                s = str(raw_direction).strip()
                # extract leading digit if present
                digits = "".join(ch for ch in s if ch.isdigit())
                if digits:
                    try:
                        dir_val = int(digits)
                    except Exception:
                        dir_val = None
        edges.append((src, dst, iface, inttype, dir_val))

    # If a focus node was requested, filter edges to only those connected to focus
    if focus:
        focus_low = focus.lower()
        filtered = []
        for src, dst, iface, inttype, dir_val in edges:
            if src.lower() == focus_low or dst.lower() == focus_low:
                filtered.append((src, dst, iface, inttype, dir_val))
        edges = filtered

    if not edges:
        print("No valid edges found in the spreadsheet.")
        sys.exit(0)

    os.makedirs(OUT_DIR, exist_ok=True)

    dot = graphviz.Digraph(name="Architecture", format="svg")
    dot.attr(rankdir="LR")
    dot.attr(splines="true")
    dot.attr(nodesep="0.4")
    dot.attr(ranksep="0.6")

    # add nodes set
    nodes = set()
    for src, dst, *_ in edges:
        nodes.add(src)
        nodes.add(dst)

    for n in sorted(nodes):
        dot.node(n, shape="box")

    # Add legend: HTML table node, try to anchor at bottom-left using a sink rank
    # smaller-font legend placed in a sink-ranked subgraph to bias bottom placement
    legend_label = (
        '<<TABLE BORDER="0" CELLBORDER="0" CELLSPACING="2">'
        '<TR><TD><FONT POINT-SIZE="8" COLOR="red">&#9632;</FONT></TD>'
        '<TD><FONT POINT-SIZE="8">File-based interfaces</FONT></TD></TR>'
        '<TR><TD><FONT POINT-SIZE="8" COLOR="blue">&#9632;</FONT></TD>'
        '<TD><FONT POINT-SIZE="8">Webservice interfaces</FONT></TD></TR>'
        '</TABLE>>'
    )
    # create a subgraph for the legend with rank=sink to push it toward the bottom
    try:
        sg = graphviz.Digraph(name="cluster_legend")
        sg.attr(rank="sink")
        sg.node("legend", label=legend_label, shape="none")
        dot.subgraph(sg)
        # anchor legend to the leftmost node with an invisible, weighted edge so it sits left-bottom
        leftmost = sorted(nodes)[0] if nodes else None
        if leftmost:
            dot.edge("legend", leftmost, style="invis", weight="100", constraint="true")
    except Exception:
        # if any Graphviz API call fails, fall back to a simple legend node
        dot.node("legend", label=legend_label, shape="none")

    # For each unique interface create an edge (Graphviz will attempt to draw
    # parallel curved edges when splines are enabled).
    for idx, (src, dst, iface, inttype, dir_val) in enumerate(edges, start=1):
        edge_attrs = {}
        # color mapping
        if inttype:
            t = inttype.lower()
            if "file" in t:
                edge_attrs["color"] = "red"
                edge_attrs["fontcolor"] = "red"
            elif "webservice" in t or "web service" in t or "webservice" in t:
                edge_attrs["color"] = "blue"
                edge_attrs["fontcolor"] = "blue"

        # direction handling: 1 = one-way, 2 = two-way
        if dir_val == 2:
            edge_attrs["dir"] = "both"
        else:
            edge_attrs["dir"] = "forward"

        # ensure label/font size and place the interface name along the arrow
        edge_attrs.setdefault("fontsize", "10")
        # attempt to center the label on the edge (angle 0 and reasonable distance)
        edge_attrs.setdefault("labelangle", "0")
        edge_attrs.setdefault("labeldistance", "1.0")
        # Use `label` only so the interface name appears on the arrow itself
        # Graphviz python API will quote labels containing spaces
        dot.edge(src, dst, label=iface, **edge_attrs)

    # choose output base name (include focus if provided)
    base_name = f"architecture_{focus.replace(' ', '_')}" if focus else "architecture"
    dot_path = os.path.join(OUT_DIR, f"{base_name}.dot")
    with open(dot_path, "w", encoding="utf-8") as f:
        f.write(dot.source)

    print("Wrote DOT file to:", dot_path)

    # try to render SVG and PNG
    svg_path = os.path.join(OUT_DIR, f"{base_name}.svg")
    png_path = os.path.join(OUT_DIR, f"{base_name}.png")
    try:
        # render to svg
        dot.render(filename=os.path.join(OUT_DIR, base_name), cleanup=True)
        if os.path.exists(svg_path):
            print("Rendered SVG to:", svg_path)
        else:
            print("Render finished but SVG not found at expected path.")
    except Exception as e:
        print("Rendering failed (Graphviz executable may be missing):", e)
        print("You still have the DOT file and can render it with the 'dot' tool:")
        print("  dot -Tsvg -o Output/Diagrams/architecture.svg Output/Diagrams/architecture.dot")

    # try to create a PNG using dot command if available
    try:
        dot.format = "png"
        dot.render(filename=os.path.join(OUT_DIR, base_name), cleanup=True)
        if os.path.exists(png_path):
            print("Rendered PNG to:", png_path)
    except Exception:
        # not fatal
        pass


if __name__ == "__main__":
    main()
