"""
Streamlit packing-list viewer/editor — replicates the factory packing-list
sheet (BUYER / LOT NO / P.O NO header, per-carton size grid, TOTAL row,
CUT/ORDER/SHIP/EXS-SHT summary block, weight & CBM footer).

Run:
    streamlit run packing_list_app.py

Feed it a po_extracted.json produced by extract_po.py, or just open the
page as-is to see the demo line (VN000C9VEMQ) pre-filled to match the
reference sheet exactly.
"""

import json
import streamlit as st

st.set_page_config(page_title="Packing List", layout="wide")

# ============================================================
# 1. CARTON-SPLIT LOGIC
#    <-- Swap this for your own customized carton algorithm.
#    This placeholder just needs to return a list of dicts:
#    {"sizes": {size: qty, ...}, "total_ctn": int,
#     "ctn_pcs": int, "color": str|None, "upc": str|None, "mixed": bool}
# ============================================================
def compute_carton_rows(sublines, pack):
    full_rows, remainder_pool = [], []
    for s in sublines:
        qty, size = s["qty"], s["size"]
        full_ctn, rem = divmod(qty, pack)
        if full_ctn > 0:
            full_rows.append({"sizes": {size: pack}, "total_ctn": full_ctn,
                               "ctn_pcs": pack, "color": s["color"],
                               "upc": s["upc"], "mixed": False})
        if rem > 0:
            remainder_pool.append([size, rem, s["color"], s["upc"]])

    mixed_rows, current, current_total, idx = [], {}, 0, 0
    while idx < len(remainder_pool):
        size, rem, color, upc = remainder_pool[idx]
        take = min(rem, pack - current_total)
        if take > 0:
            current[size] = current.get(size, 0) + take
            current_total += take
            remainder_pool[idx][1] -= take
        if remainder_pool[idx][1] == 0:
            idx += 1
        if current_total == pack:
            mixed_rows.append({"sizes": dict(current), "total_ctn": 1,
                                "ctn_pcs": pack, "color": None, "upc": None,
                                "mixed": True})
            current, current_total = {}, 0
    if current:
        mixed_rows.append({"sizes": dict(current), "total_ctn": 1,
                            "ctn_pcs": current_total, "color": None,
                            "upc": None, "mixed": True})
    return full_rows + mixed_rows


# ============================================================
# 2. DEMO DATA — matches the reference sheet exactly (VN000C9VEMQ)
# ============================================================
DEMO_LINE = {
    "po_line_no": "600090369200002", "style": "VN000C9VEMQ",
    "description": "CHECK-5 BAGGY DENIM SHORT WASH", "order_qty": 180,
    "crd": "2026-01-10", "destination_country": "AUSTRALIA",
}
DEMO_SUBLINES = [
    {"style": "VN000C9VEMQ", "qty": 42, "color": "WASHED BLACK", "size": "30 REG",
     "upc": "196573552661", "items_per_outer_pack": 24},
    {"style": "VN000C9VEMQ", "qty": 66, "color": "WASHED BLACK", "size": "32 REG",
     "upc": "196573552876", "items_per_outer_pack": 24},
    {"style": "VN000C9VEMQ", "qty": 47, "color": "WASHED BLACK", "size": "34 REG",
     "upc": "196573552944", "items_per_outer_pack": 24},
    {"style": "VN000C9VEMQ", "qty": 25, "color": "WASHED BLACK", "size": "36 REG",
     "upc": "196573552968", "items_per_outer_pack": 24},
]
# Reference carton breakdown exactly as shown on the target sheet
DEMO_CARTON_ROWS = [
    {"sizes": {"30 REG": 22}, "total_ctn": 1, "ctn_pcs": 22, "color": "WASHED BLACK", "upc": "196573552661", "mixed": False},
    {"sizes": {"32 REG": 22}, "total_ctn": 3, "ctn_pcs": 22, "color": "WASHED BLACK", "upc": "196573552876", "mixed": False},
    {"sizes": {"34 REG": 23}, "total_ctn": 2, "ctn_pcs": 23, "color": "WASHED BLACK", "upc": "196573552944", "mixed": False},
    {"sizes": {"36 REG": 22}, "total_ctn": 1, "ctn_pcs": 22, "color": "WASHED BLACK", "upc": "196573552968", "mixed": False},
    {"sizes": {"30 REG": 20, "34 REG": 1, "36 REG": 3}, "total_ctn": 1, "ctn_pcs": 24, "color": None, "upc": None, "mixed": True},
]
DEMO_GRS_WT = [14.81, 15.47, 16.81, 16.18, 13.57]
DEMO_NET_WT = [13.81, 14.47, 15.81, 15.18, 12.57]
DEMO_CUT_QTY = {"30 REG": 43, "32 REG": 68, "34 REG": 48, "36 REG": 26}
DEMO_CTN_MEAS = ['L 24" X W 16" X H 10"', 'L 24" X W 16" X H 6"']
DEMO_CBM_PER_CTN = 0.06295081967

# ============================================================
# 3. SIDEBAR — load real data, or fall back to the demo line
# ============================================================
st.sidebar.header("Data source")
uploaded = st.sidebar.file_uploader("po_extracted.json", type="json")

use_demo = True
line, sublines = DEMO_LINE, DEMO_SUBLINES

if uploaded is not None:
    data = json.load(uploaded)
    styles = [l["style"] for l in data["lines"]]
    picked = st.sidebar.selectbox("PO line / style", styles)
    line = next(l for l in data["lines"] if l["style"] == picked)
    sublines = [s for s in data["sublines"] if s["style"] == picked]
    use_demo = (picked == "VN000C9VEMQ")

full_size_text = st.sidebar.text_input(
    "Full size range (comma-separated — shows 0-qty sizes too)",
    "28 REG, 30 REG, 32 REG, 34 REG, 36 REG" if use_demo else "",
)
full_sizes = [s.strip() for s in full_size_text.split(",") if s.strip()] or \
             sorted({s["size"] for s in sublines})
ordered_sizes = [s["size"] for s in sublines]  # sizes actually on the PO line
pack = sublines[0]["items_per_outer_pack"] if sublines else 1

if use_demo:
    carton_rows = DEMO_CARTON_ROWS
else:
    carton_rows = compute_carton_rows(sublines, pack)

# ============================================================
# 4. EDITABLE (YELLOW) INPUTS
# ============================================================
with st.sidebar.expander("Carton weights (yellow cells)", expanded=False):
    grs_wts, net_wts = [], []
    for i, row in enumerate(carton_rows):
        default_g = DEMO_GRS_WT[i] if use_demo and i < len(DEMO_GRS_WT) else 0.0
        default_n = DEMO_NET_WT[i] if use_demo and i < len(DEMO_NET_WT) else 0.0
        g = st.number_input(f"Ctn row {i+1} — Grs.wt/ctn (kg)", value=float(default_g), key=f"g{i}")
        n = st.number_input(f"Ctn row {i+1} — Net.wt/ctn (kg)", value=float(default_n), key=f"n{i}")
        grs_wts.append(g)
        net_wts.append(n)

with st.sidebar.expander("Cut quantities (yellow cells)", expanded=False):
    cut_qty = {}
    for size in full_sizes:
        default = DEMO_CUT_QTY.get(size, 0) if use_demo else 0
        cut_qty[size] = st.number_input(f"Cut qty — {size}", value=int(default), key=f"cut_{size}")

with st.sidebar.expander("Carton measurements & CBM (yellow cells)", expanded=False):
    ctn_meas_1 = st.text_input("Ctn meas — line 1", DEMO_CTN_MEAS[0] if use_demo else "")
    ctn_meas_2 = st.text_input("Ctn meas — line 2", DEMO_CTN_MEAS[1] if use_demo else "")
    cbm_per_ctn = st.number_input("CBM per carton (m³)", value=float(DEMO_CBM_PER_CTN if use_demo else 0.0), format="%.8f")

# ============================================================
# 5. DERIVED / COMPUTED VALUES
# ============================================================
total_ctn = sum(r["total_ctn"] for r in carton_rows)
total_pcs = sum(r["total_ctn"] * r["ctn_pcs"] for r in carton_rows)
total_grs = sum(r["total_ctn"] * g for r, g in zip(carton_rows, grs_wts))
total_net = sum(r["total_ctn"] * n for r, n in zip(carton_rows, net_wts))

order_qty_by_size = {s["size"]: s["qty"] for s in sublines}
ship_qty_by_size = {size: 0 for size in full_sizes}
for r in carton_rows:
    for size, qty in r["sizes"].items():
        ship_qty_by_size[size] = ship_qty_by_size.get(size, 0) + qty * r["total_ctn"]

order_total = sum(order_qty_by_size.get(sz, 0) for sz in full_sizes)
ship_total = sum(ship_qty_by_size.get(sz, 0) for sz in full_sizes)
cut_total = sum(cut_qty.get(sz, 0) for sz in full_sizes)
exs_by_size = {sz: ship_qty_by_size.get(sz, 0) - order_qty_by_size.get(sz, 0) for sz in full_sizes}
exs_total = ship_total - order_total
pct_by_size = {sz: (exs_by_size[sz] / order_qty_by_size[sz]) if order_qty_by_size.get(sz) else None for sz in full_sizes}
pct_total = (exs_total / order_total) if order_total else None
cbm_total = cbm_per_ctn * total_ctn

# cumulative carton numbers
cum, carton_no_list, case_label_list = 0, [], []
for i, r in enumerate(carton_rows):
    case_label_list.append(i + 1)
    cum += r["total_ctn"]
    carton_no_list.append(cum)

# ============================================================
# 6. RENDER — HTML/CSS table matching the reference sheet
# ============================================================
st.markdown("""
<style>
.plist-wrap { font-family: Arial, sans-serif; font-size: 13px; color: #000; }
table.plist { border-collapse: collapse; width: 100%; margin-bottom: 6px; }
table.plist td, table.plist th { border: 1px solid #000; padding: 3px 8px; text-align: center; }
.hdr { font-weight: bold; background: #fff; }
.yellow { background: #FFFF66; }
.gray { background: #D9D9D9; font-weight: bold; }
.bold { font-weight: bold; }
.noborder td { border: none; padding: 1px 4px; }
.desc-cell { text-align: left; font-weight: bold; }
.title-row td { border: none; font-weight: bold; padding: 2px 4px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="plist-wrap">', unsafe_allow_html=True)

# ---- Header info block ----
st.markdown(f"""
<table class="plist noborder">
<tr class="title-row">
  <td style="width:16%"><b>BUYER</b></td><td style="width:14%">VANS</td>
  <td style="width:14%"></td>
  <td style="width:16%"><b>ORDER QTY</b></td><td style="width:8%">{line['order_qty']}</td><td style="width:6%">PCS</td>
  <td style="width:8%"></td>
  <td style="width:6%"><u><b>CRD</b></u></td><td>{line.get('crd','')}</td>
</tr>
<tr class="title-row">
  <td><b>LOT NO</b></td><td>{line['style']}</td><td></td>
  <td><b>PACK QTY</b></td><td>{line['order_qty']}</td><td>PCS</td><td></td><td></td><td></td>
</tr>
<tr class="title-row">
  <td><b>P.O NO</b></td><td>{line['po_line_no']}</td><td></td>
  <td><b>EXCESS/SHORT QTY</b></td><td>{exs_total}</td><td>PCS</td><td></td><td></td><td></td>
</tr>
<tr class="title-row">
  <td colspan="2" class="desc-cell">DESCRIPTION:<br>{line['description']}</td><td></td>
  <td><b>PERCENTAGE</b></td>
  <td>{'-' if pct_total is None else f'{pct_total*100:.2f}'}</td><td>%</td><td></td>
  <td colspan="2">{line.get('destination_country','')}</td>
</tr>
</table>
""", unsafe_allow_html=True)

# ---- Main carton table ----
size_headers = "".join(f"<th>{sz}</th>" for sz in full_sizes)
rows_html = ""
for i, r in enumerate(carton_rows):
    color = "MIXED" if r["mixed"] else (r["color"] or "")
    upc = r["upc"] or ""
    size_cells = "".join(
        f"<td>{r['sizes'].get(sz, '')}</td>" for sz in full_sizes
    )
    row_class = "yellow" if r["mixed"] else ""
    rows_html += (
        "<tr>"
        f"<td>{case_label_list[i]}</td><td>{carton_no_list[i]}</td>"
        f'<td class="{row_class}">{color}</td><td class="{row_class}">{upc}</td>'
        f"{size_cells}"
        f"<td>{r['ctn_pcs']}</td><td>{r['total_ctn']}</td>"
        f"<td>{r['total_ctn'] * r['ctn_pcs']}</td>"
        f"<td>{grs_wts[i]:.2f}</td><td>{net_wts[i]:.2f}</td>"
        f"<td>{r['total_ctn'] * grs_wts[i]:.2f}</td><td>{r['total_ctn'] * net_wts[i]:.2f}</td>"
        "</tr>"
    )

size_total_cells = "".join(
    f"<td>{sum(r['sizes'].get(sz,0)*r['total_ctn'] for r in carton_rows)}</td>" for sz in full_sizes
)

st.markdown(f"""
<table class="plist">
<tr class="hdr">
  <th colspan="2">CASE LABEL NO.</th><th>COLOR</th><th>UPC Number</th>
  <th colspan="{len(full_sizes)}">SIZE</th>
  <th>CTN PCS</th><th>TOTAL CTN</th><th>TOTAL PCS</th>
  <th>Grs.wt/pr ctn</th><th>Net.wt/pr ctn</th><th>total Grs.wt</th><th>total net wt</th>
</tr>
<tr class="hdr">
  <td></td><td></td><td></td><td></td>{size_headers}<td></td><td></td><td></td><td></td><td></td><td></td><td></td>
</tr>
{rows_html}
<tr class="gray">
  <td colspan="2"></td><td>TOTAL</td><td></td>
  {size_total_cells}
  <td></td><td>{total_ctn}</td><td>{total_pcs}</td>
  <td colspan="2"></td><td>{total_grs:.2f}</td><td>{total_net:.2f}</td>
</tr>
</table>
""", unsafe_allow_html=True)

# ---- Summary block ----
def fmt_pct(v):
    return "" if v is None else f"{v*100:.2f}%"

summary_rows = [
    ("CUT QTY", [cut_qty.get(sz, "") for sz in full_sizes], cut_total, "yellow"),
    ("ORDER QTY", [order_qty_by_size.get(sz, "") for sz in full_sizes], order_total, "gray"),
    ("SHIP QTY", [ship_qty_by_size.get(sz, "") for sz in full_sizes], ship_total, ""),
    ("EXS/SHT QTY", [exs_by_size.get(sz, "") for sz in full_sizes], exs_total, ""),
    ("PERCENTAGE", [fmt_pct(pct_by_size.get(sz)) for sz in full_sizes], fmt_pct(pct_total), ""),
]
summary_body = ""
for label, vals, gt, cls in summary_rows:
    cells = "".join(f'<td class="{cls}">{v}</td>' for v in vals)
    summary_body += f'<tr><td class="{cls} bold">{label}</td>{cells}<td class="bold">{gt}</td></tr>'

st.markdown(f"""
<table class="plist">
<tr class="hdr"><td rowspan="2"></td>{"".join(f"<th>{sz}</th>" for sz in full_sizes)}<th rowspan="2">G Total</th></tr>
{summary_body}
</table>
""", unsafe_allow_html=True)

# ---- Weight / CBM footer ----
st.markdown(f"""
<table class="plist noborder">
<tr class="title-row"><td style="width:15%"><b>GROSS WEIGHT :</b></td><td>{total_grs:.2f}</td></tr>
<tr class="title-row"><td><b>NET WEIGHT :</b></td><td>{total_net:.2f}</td></tr>
<tr class="title-row"><td><b>CTN MEAS :</b></td><td>{ctn_meas_1}<br>{ctn_meas_2}</td></tr>
<tr class="title-row"><td><b>CBM :</b></td><td>{cbm_total:.8f}</td></tr>
</table>
<br><br>
<table class="plist noborder">
<tr class="title-row"><td style="width:30%; border-top:1px solid #000; text-align:center;">INCHARGE</td>
<td style="width:30%"></td>
<td style="width:30%; border-top:1px solid #000; text-align:center;">FM</td></tr>
</table>
""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)