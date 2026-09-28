
# # # # # # """
# # # # # # Build a packing-list-style sheet for EVERY PO line found in
# # # # # # po_extracted.json -- one worksheet per line, in a single workbook.

# # # # # # HOW TO RUN (after extract_po.py has created po_extracted.json):
# # # # # #     python3 build_all_packing_lists.py

# # # # # # Requires:
# # # # # #     pip install openpyxl
# # # # # # """
# # # # # # import json
# # # # # # import re
# # # # # # from openpyxl import Workbook
# # # # # # from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
# # # # # # from openpyxl.utils import get_column_letter

# # # # # # with open("po_extracted.json") as f:
# # # # # #     data = json.load(f)

# # # # # # FONT = "Arial"
# # # # # # bold = Font(name=FONT, bold=True, size=10)
# # # # # # normal = Font(name=FONT, size=10)
# # # # # # title_font = Font(name=FONT, bold=True, size=13)
# # # # # # thin = Side(style="thin")
# # # # # # box = Border(left=thin, right=thin, top=thin, bottom=thin)
# # # # # # center = Alignment(horizontal="center", vertical="center", wrap_text=True)
# # # # # # yellow = PatternFill("solid", fgColor="FFFF00")
# # # # # # grey = PatternFill("solid", fgColor="D9D9D9")


# # # # # # def sheet_name_for(line):
# # # # # #     raw = f"{line['style']}"
# # # # # #     return re.sub(r'[\[\]:*?/\\]', '-', raw)[:31]


# # # # # # def build_sheet(ws, line, sublines):
# # # # # #     sizes = [s["size"] for s in sublines]
# # # # # #     n_sizes = len(sizes)

# # # # # #     ws.merge_cells("A1:H1")
# # # # # #     ws["A1"] = "CREATIVE COLLECTIONS LTD. U-1-A"
# # # # # #     ws["A1"].font = title_font
# # # # # #     ws["A1"].alignment = center

# # # # # #     ws.merge_cells("A2:H2")
# # # # # #     ws["A2"] = "PACKING LIST DETAILS"
# # # # # #     ws["A2"].font = Font(name=FONT, bold=True, size=10)
# # # # # #     ws["A2"].alignment = center

# # # # # #     # ---- Small size-spec table: SIZE / N.W. / N.N.W. / EMPTY CTN ----
# # # # # #     spec_header_row = 4
# # # # # #     ws.cell(row=spec_header_row, column=1, value="SIZE").font = bold
# # # # # #     ws.cell(row=spec_header_row, column=1).border = box
# # # # # #     for sc_idx, size_name in enumerate(sizes):
# # # # # #         c = ws.cell(row=spec_header_row, column=2 + sc_idx, value=size_name)
# # # # # #         c.font = bold
# # # # # #         c.alignment = center
# # # # # #         c.border = box
# # # # # #         c.fill = grey

# # # # # #     spec_rows = ["N.W.", "N.N.W.", "EMPTY CTN"]
# # # # # #     for i, label in enumerate(spec_rows):
# # # # # #         rr = spec_header_row + 1 + i
# # # # # #         lc = ws.cell(row=rr, column=1, value=label)
# # # # # #         lc.font = bold
# # # # # #         lc.border = box
# # # # # #         for sc_idx, size_name in enumerate(sizes):
# # # # # #             c = ws.cell(row=rr, column=2 + sc_idx, value=1)
# # # # # #             c.alignment = center
# # # # # #             c.border = box

# # # # # #     r = spec_header_row + len(spec_rows) + 2
# # # # # #     ws.cell(row=r, column=1, value="BUYER").font = bold
# # # # # #     ws.cell(row=r, column=2, value="VANS").font = normal
# # # # # #     r += 1
# # # # # #     ws.cell(row=r, column=1, value="LOT NO").font = bold
# # # # # #     ws.cell(row=r, column=2, value=line["style"]).font = normal
# # # # # #     r += 1
# # # # # #     ws.cell(row=r, column=1, value="P.O NO").font = bold
# # # # # #     ws.cell(row=r, column=2, value=line["po_line_no"]).font = normal
# # # # # #     r += 1
# # # # # #     ws.cell(row=r, column=1, value="ORDER QTY").font = bold
# # # # # #     ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
# # # # # #     ws.cell(row=r, column=3, value="PCS").font = normal
# # # # # #     r += 1
# # # # # #     ws.cell(row=r, column=1, value="PACK QTY").font = bold
# # # # # #     ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
# # # # # #     ws.cell(row=r, column=3, value="PCS").font = normal
# # # # # #     r += 1
# # # # # #     excess_short_row = r  # filled in with a formula once the summary block below is built
# # # # # #     ws.cell(row=r, column=1, value="EXCESS/SHORT QTY").font = bold
# # # # # #     ws.cell(row=r, column=3, value="PCS").font = normal
# # # # # #     r += 1
# # # # # #     percentage_row = r  # filled in with a formula once the summary block below is built
# # # # # #     ws.cell(row=r, column=1, value="PERCENTAGE").font = bold
# # # # # #     r += 1
# # # # # #     ws.cell(row=r, column=1, value="CRD").font = bold
# # # # # #     ws.cell(row=r, column=2, value=line.get("crd")).font = normal
# # # # # #     r += 1
# # # # # #     ws.cell(row=r, column=1, value="COUNTRY").font = bold
# # # # # #     ws.cell(row=r, column=2, value=line.get("destination_country")).font = normal
# # # # # #     r += 1
# # # # # #     ws.cell(row=r, column=1, value="DESCRIPTION").font = bold
# # # # # #     ws.cell(row=r, column=2, value=line["description"]).font = normal

# # # # # #     header_row = r + 2
# # # # # #     headers = ["CASE LABEL NO.", "CARTON NO.", "COLOR", "UPC Number"] + sizes + \
# # # # # #               ["CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs.wt.pr ctn", "Net.wt.pr ctn",
# # # # # #                "total.Grs.wt", "total.net.wt"]
# # # # # #     for i, h in enumerate(headers, start=1):
# # # # # #         c = ws.cell(row=header_row, column=i, value=h)
# # # # # #         c.font = bold
# # # # # #         c.alignment = center
# # # # # #         c.border = box
# # # # # #         c.fill = grey
# # # # # #     n_cols = len(headers)

# # # # # #     data_row_start = header_row + 1
# # # # # #     for idx, s in enumerate(sublines):
# # # # # #         rr = data_row_start + idx
# # # # # #         ws.cell(row=rr, column=1, value=idx + 1).border = box
# # # # # #         cB = ws.cell(row=rr, column=2, value=None)
# # # # # #         cB.fill = yellow
# # # # # #         cB.border = box
# # # # # #         ws.cell(row=rr, column=3, value=s["color"]).border = box
# # # # # #         ws.cell(row=rr, column=4, value=s["upc"]).border = box
# # # # # #         for sc_idx, size_name in enumerate(sizes):
# # # # # #             col = 5 + sc_idx
# # # # # #             val = s["qty"] if size_name == s["size"] else None
# # # # # #             cell = ws.cell(row=rr, column=col, value=val)
# # # # # #             cell.border = box
# # # # # #             cell.alignment = center

# # # # # #         # CTN PCS = sum of all size columns in THIS row (handles mixed cartons)
# # # # # #         ctn_pcs_col = 5 + n_sizes
# # # # # #         first_size_letter = get_column_letter(5)
# # # # # #         last_size_letter = get_column_letter(4 + n_sizes)
# # # # # #         ctn_cell = ws.cell(
# # # # # #             row=rr, column=ctn_pcs_col,
# # # # # #             value=f"=SUM({first_size_letter}{rr}:{last_size_letter}{rr})"
# # # # # #         )
# # # # # #         ctn_cell.border = box
# # # # # #         ctn_cell.alignment = center
# # # # # #         ctn_cell.font = bold

# # # # # #         # Remaining columns after CTN PCS: TOTAL CTN, TOTAL PCS, then weights
# # # # # #         total_ctn_col = ctn_pcs_col + 1
# # # # # #         total_pcs_col = ctn_pcs_col + 2
# # # # # #         total_ctn_letter = get_column_letter(total_ctn_col)

# # # # # #         # TOTAL CTN stays manual/yellow (user enters actual carton count)
# # # # # #         tc_cell = ws.cell(row=rr, column=total_ctn_col, value=None)
# # # # # #         tc_cell.fill = yellow
# # # # # #         tc_cell.border = box

# # # # # #         # TOTAL PCS = TOTAL CTN * CTN PCS (blank until TOTAL CTN is filled in)
# # # # # #         ctn_pcs_letter = get_column_letter(ctn_pcs_col)
# # # # # #         tp_cell = ws.cell(
# # # # # #             row=rr, column=total_pcs_col,
# # # # # #             value=f'=IF({total_ctn_letter}{rr}="","",{total_ctn_letter}{rr}*{ctn_pcs_letter}{rr})'
# # # # # #         )
# # # # # #         tp_cell.border = box
# # # # # #         tp_cell.alignment = center

# # # # # #         # Grs.wt.pr ctn = SUMPRODUCT(this row's size qtys, N.W. row) + EMPTY CTN
# # # # # #         # (EMPTY CTN taken from the first size column, since it's normally
# # # # # #         #  the same physical carton regardless of size)
# # # # # #         grs_wt_col = total_pcs_col + 1
# # # # # #         nw_row = spec_header_row + 1       # N.W. row in the spec table
# # # # # #         empty_row = spec_header_row + 3    # EMPTY CTN row in the spec table
# # # # # #         first_spec_letter = get_column_letter(2)
# # # # # #         last_spec_letter = get_column_letter(1 + n_sizes)
# # # # # #         first_main_size_letter = get_column_letter(5)
# # # # # #         last_main_size_letter = get_column_letter(4 + n_sizes)

# # # # # #         grs_formula = (
# # # # # #             f"=SUMPRODUCT({first_main_size_letter}{rr}:{last_main_size_letter}{rr},"
# # # # # #             f"{first_spec_letter}{nw_row}:{last_spec_letter}{nw_row})"
# # # # # #             f"+{first_spec_letter}{empty_row}"
# # # # # #         )
# # # # # #         grs_cell = ws.cell(row=rr, column=grs_wt_col, value=grs_formula)
# # # # # #         grs_cell.border = box
# # # # # #         grs_cell.alignment = center

# # # # # #         # Net.wt.pr ctn stays manual/yellow (right after Grs.wt.pr ctn)
# # # # # #         net_wt_col = grs_wt_col + 1
# # # # # #         net_cell = ws.cell(row=rr, column=net_wt_col, value=None)
# # # # # #         net_cell.fill = yellow
# # # # # #         net_cell.border = box

# # # # # #         # total.Grs.wt = TOTAL CTN * Grs.wt.pr ctn (blank until TOTAL CTN filled in)
# # # # # #         total_grs_col = net_wt_col + 1
# # # # # #         grs_wt_letter = get_column_letter(grs_wt_col)
# # # # # #         tg_cell = ws.cell(
# # # # # #             row=rr, column=total_grs_col,
# # # # # #             value=f'=IF({total_ctn_letter}{rr}="","",{total_ctn_letter}{rr}*{grs_wt_letter}{rr})'
# # # # # #         )
# # # # # #         tg_cell.border = box
# # # # # #         tg_cell.alignment = center

# # # # # #         # total.net.wt = TOTAL CTN * Net.wt.pr ctn (blank until TOTAL CTN filled in)
# # # # # #         total_net_col = total_grs_col + 1
# # # # # #         net_wt_letter = get_column_letter(net_wt_col)
# # # # # #         tn_cell = ws.cell(
# # # # # #             row=rr, column=total_net_col,
# # # # # #             value=f'=IF({total_ctn_letter}{rr}="","",{total_ctn_letter}{rr}*{net_wt_letter}{rr})'
# # # # # #         )
# # # # # #         tn_cell.border = box
# # # # # #         tn_cell.alignment = center

# # # # # #     total_row = data_row_start + len(sublines)
# # # # # #     ws.cell(row=total_row, column=3, value="TOTAL").font = bold
# # # # # #     for sc_idx, size_name in enumerate(sizes):
# # # # # #         col = 5 + sc_idx
# # # # # #         col_letter = get_column_letter(col)
# # # # # #         cell = ws.cell(row=total_row, column=col,
# # # # # #                         value=f"=SUM({col_letter}{data_row_start}:{col_letter}{total_row-1})")
# # # # # #         cell.font = bold
# # # # # #         cell.border = box
# # # # # #     # Grand-total TOTAL CTN = sum of each row's TOTAL CTN
# # # # # #     # Grand-total TOTAL PCS = sum of each row's TOTAL PCS (not the size columns)
# # # # # #     total_ctn_col = ctn_pcs_col + 1
# # # # # #     total_pcs_col = ctn_pcs_col + 2
# # # # # #     total_ctn_letter = get_column_letter(total_ctn_col)
# # # # # #     total_pcs_letter = get_column_letter(total_pcs_col)

# # # # # #     ws.cell(row=total_row, column=total_ctn_col,
# # # # # #             value=f'=IF(COUNT({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})=0,"",'
# # # # # #                   f'SUM({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1}))').font = bold

# # # # # #     ws.cell(row=total_row, column=total_pcs_col,
# # # # # #             value=f'=IF(COUNT({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1})=0,"",'
# # # # # #                   f'SUM({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1}))').font = bold

# # # # # #     # Grand-total total.Grs.wt = sum of each row's total.Grs.wt
# # # # # #     total_grs_col = total_pcs_col + 3  # skip Grs.wt.pr ctn, Net.wt.pr ctn -> land on total.Grs.wt
# # # # # #     total_grs_letter = get_column_letter(total_grs_col)
# # # # # #     ws.cell(row=total_row, column=total_grs_col,
# # # # # #             value=f'=IF(COUNT({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1})=0,"",'
# # # # # #                   f'SUM({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1}))').font = bold

# # # # # #     # Grand-total total.net.wt = sum of each row's total.net.wt
# # # # # #     total_net_col = total_grs_col + 1
# # # # # #     total_net_letter = get_column_letter(total_net_col)
# # # # # #     ws.cell(row=total_row, column=total_net_col,
# # # # # #             value=f'=IF(COUNT({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1})=0,"",'
# # # # # #                   f'SUM({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1}))').font = bold

# # # # # #     sum_header_row = total_row + 3
# # # # # #     ws.cell(row=sum_header_row, column=1, value="SIZE").font = bold
# # # # # #     for sc_idx, size_name in enumerate(sizes):
# # # # # #         ws.cell(row=sum_header_row, column=2 + sc_idx, value=size_name).font = bold
# # # # # #     ws.cell(row=sum_header_row, column=2 + n_sizes, value="G TOTAL").font = bold

# # # # # #     # Two extra header columns: Grand Total gross/net weight (kg)
# # # # # #     gw_col = 3 + n_sizes
# # # # # #     nw_col = 4 + n_sizes
# # # # # #     ws.merge_cells(start_row=sum_header_row, start_column=gw_col,
# # # # # #                     end_row=sum_header_row, end_column=gw_col)
# # # # # #     ws.cell(row=sum_header_row, column=gw_col, value="Grand Total gross weight (kg)").font = bold
# # # # # #     ws.cell(row=sum_header_row, column=gw_col).alignment = center
# # # # # #     ws.merge_cells(start_row=sum_header_row, start_column=nw_col,
# # # # # #                     end_row=sum_header_row, end_column=nw_col)
# # # # # #     ws.cell(row=sum_header_row, column=nw_col, value="Grand Total net weight (kg)").font = bold
# # # # # #     ws.cell(row=sum_header_row, column=nw_col).alignment = center

# # # # # #     rows_needed = ["CUT QTY", "ORDER QTY", "SHIP QTY", "EXS/SHT QTY", "PERCENTAGE"]
# # # # # #     for i, label in enumerate(rows_needed):
# # # # # #         rr = sum_header_row + 1 + i
# # # # # #         ws.cell(row=rr, column=1, value=label).font = bold
# # # # # #         for sc_idx, size_name in enumerate(sizes):
# # # # # #             col = 2 + sc_idx
# # # # # #             col_letter = get_column_letter(col)
# # # # # #             cell = ws.cell(row=rr, column=col)
# # # # # #             if label == "ORDER QTY":
# # # # # #                 qty = next(s["qty"] for s in sublines if s["size"] == size_name)
# # # # # #                 cell.value = qty
# # # # # #             elif label in ("CUT QTY", "SHIP QTY"):
# # # # # #                 cell.fill = yellow
# # # # # #             elif label == "EXS/SHT QTY":
# # # # # #                 ship_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
# # # # # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # # # #                 cell.value = f'=IF({col_letter}{ship_row}="","",{col_letter}{ship_row}-{col_letter}{order_row})'
# # # # # #             elif label == "PERCENTAGE":
# # # # # #                 exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # # # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # # # #                 cell.value = f'=IF({col_letter}{exs_row}="","",IFERROR({col_letter}{exs_row}/{col_letter}{order_row},0))'
# # # # # #                 cell.number_format = "0.00%"
# # # # # #             cell.border = box
# # # # # #             cell.alignment = center

# # # # # #         gt_col = 2 + n_sizes
# # # # # #         gt_letter = get_column_letter(gt_col)
# # # # # #         first_letter = get_column_letter(2)
# # # # # #         last_letter = get_column_letter(1 + n_sizes)
# # # # # #         if label in ("CUT QTY", "SHIP QTY", "EXS/SHT QTY"):
# # # # # #             ws.cell(row=rr, column=gt_col,
# # # # # #                     value=f'=IF(COUNT({first_letter}{rr}:{last_letter}{rr})=0,"",SUM({first_letter}{rr}:{last_letter}{rr}))').font = bold
# # # # # #         elif label == "ORDER QTY":
# # # # # #             ws.cell(row=rr, column=gt_col,
# # # # # #                     value=f"=SUM({first_letter}{rr}:{last_letter}{rr})").font = bold
# # # # # #         else:  # PERCENTAGE
# # # # # #             exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # # # #             order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # # # #             ws.cell(row=rr, column=gt_col,
# # # # # #                     value=f'=IF({gt_letter}{exs_row}="","",IFERROR({gt_letter}{exs_row}/{gt_letter}{order_row},0))').number_format = "0.00%"
# # # # # #         ws.cell(row=rr, column=gt_col).border = box

# # # # # #     # Link the top header's EXCESS/SHORT QTY and PERCENTAGE to the
# # # # # #     # summary block's G-Total values (calculated above) - same numbers,
# # # # # #     # no separate manual entry needed.
# # # # # #     exs_gt_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # # # #     order_gt_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # # # #     gt_col_letter = get_column_letter(2 + n_sizes)

# # # # # #     ws.cell(row=excess_short_row, column=2,
# # # # # #             value=f'={gt_col_letter}{exs_gt_row}').font = normal
# # # # # #     ws.cell(row=percentage_row, column=2,
# # # # # #             value=f'=IF({gt_col_letter}{exs_gt_row}="","-",'
# # # # # #                   f'TEXT(IFERROR({gt_col_letter}{exs_gt_row}/{gt_col_letter}{order_gt_row},0),"0.00%"))'
# # # # # #             ).font = normal

# # # # # #     wt_row = sum_header_row + len(rows_needed) + 2
# # # # # #     labels = ["GROSS WEIGHT", "NET WEIGHT", "CTN MEAS", "CBM"]
# # # # # #     for i, lbl in enumerate(labels):
# # # # # #         ws.cell(row=wt_row + i, column=1, value=lbl + " :").font = bold
# # # # # #         ws.cell(row=wt_row + i, column=2, value=None).fill = yellow

# # # # # #     # Grand Total gross/net weight (kg) - merged, pulled from the main
# # # # # #     # table's total.Grs.wt / total.net.wt grand totals (computed above)
# # # # # #     first_sum_row = sum_header_row + 1
# # # # # #     last_sum_row = sum_header_row + len(rows_needed)
# # # # # #     total_grs_letter = get_column_letter(total_grs_col)
# # # # # #     total_net_letter = get_column_letter(total_net_col)

# # # # # #     ws.merge_cells(start_row=first_sum_row, start_column=gw_col,
# # # # # #                     end_row=last_sum_row, end_column=gw_col)
# # # # # #     gw_cell = ws.cell(row=first_sum_row, column=gw_col,
# # # # # #                        value=f'={total_grs_letter}{total_row}')
# # # # # #     gw_cell.alignment = center
# # # # # #     gw_cell.font = bold
# # # # # #     gw_cell.border = box

# # # # # #     ws.merge_cells(start_row=first_sum_row, start_column=nw_col,
# # # # # #                     end_row=last_sum_row, end_column=nw_col)
# # # # # #     nw_cell = ws.cell(row=first_sum_row, column=nw_col,
# # # # # #                        value=f'={total_net_letter}{total_row}')
# # # # # #     nw_cell.alignment = center
# # # # # #     nw_cell.font = bold
# # # # # #     nw_cell.border = box

# # # # # #     ws.column_dimensions["A"].width = 16
# # # # # #     ws.column_dimensions["B"].width = 14
# # # # # #     ws.column_dimensions["C"].width = 16
# # # # # #     ws.column_dimensions["D"].width = 16
# # # # # #     for i in range(n_sizes):
# # # # # #         ws.column_dimensions[get_column_letter(5 + i)].width = 10
# # # # # #     for col in range(5 + n_sizes, n_cols + 1):
# # # # # #         ws.column_dimensions[get_column_letter(col)].width = 13
# # # # # #     ws.column_dimensions[get_column_letter(gw_col)].width = 22
# # # # # #     ws.column_dimensions[get_column_letter(nw_col)].width = 22

# # # # # #     legend_row = wt_row + len(labels) + 2
# # # # # #     ws.cell(row=legend_row, column=1, value="Legend:").font = bold
# # # # # #     ws.cell(row=legend_row + 1, column=1,
# # # # # #             value="Yellow cells = data NOT available in the PO PDF (fill in from factory cutting/packing records)"
# # # # # #             ).font = Font(name=FONT, size=9, italic=True)


# # # # # # wb = Workbook()
# # # # # # wb.remove(wb.active)

# # # # # # used_names = set()
# # # # # # for line in data["lines"]:
# # # # # #     sublines = [s for s in data["sublines"] if s["style"] == line["style"]]
# # # # # #     if not sublines:
# # # # # #         continue
# # # # # #     name = sheet_name_for(line)
# # # # # #     base_name, i = name, 1
# # # # # #     while name in used_names:
# # # # # #         i += 1
# # # # # #         name = f"{base_name[:28]}_{i}"
# # # # # #     used_names.add(name)

# # # # # #     ws = wb.create_sheet(title=name)
# # # # # #     build_sheet(ws, line, sublines)

# # # # # # out_path = "packing_list_all_PO_lines.xlsx"
# # # # # # wb.save(out_path)
# # # # # # print("Saved:", out_path)
# # # # # # print("Sheets created:", wb.sheetnames)





# # # # # """
# # # # # Build a packing-list-style sheet for EVERY PO line found in
# # # # # po_extracted.json -- one worksheet per line, in a single workbook.

# # # # # HOW TO RUN (after extract_po.py has created po_extracted.json):
# # # # #     python3 build_all_packing_lists.py

# # # # # Requires:
# # # # #     pip install openpyxl
# # # # # """
# # # # # import json
# # # # # import re
# # # # # from openpyxl import Workbook
# # # # # from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
# # # # # from openpyxl.utils import get_column_letter

# # # # # with open("po_extracted.json") as f:
# # # # #     data = json.load(f)

# # # # # FONT = "Arial"
# # # # # bold = Font(name=FONT, bold=True, size=10)
# # # # # normal = Font(name=FONT, size=10)
# # # # # title_font = Font(name=FONT, bold=True, size=13)
# # # # # thin = Side(style="thin")
# # # # # box = Border(left=thin, right=thin, top=thin, bottom=thin)
# # # # # center = Alignment(horizontal="center", vertical="center", wrap_text=True)
# # # # # yellow = PatternFill("solid", fgColor="FFFF00")
# # # # # grey = PatternFill("solid", fgColor="D9D9D9")


# # # # # def sheet_name_for(line):
# # # # #     raw = f"{line['style']}"
# # # # #     return re.sub(r'[\[\]:*?/\\]', '-', raw)[:31]


# # # # # def build_sheet(ws, line, sublines):
# # # # #     sizes = [s["size"] for s in sublines]
# # # # #     n_sizes = len(sizes)

# # # # #     ws.merge_cells("A1:H1")
# # # # #     ws["A1"] = "CREATIVE COLLECTIONS LTD. U-1-A"
# # # # #     ws["A1"].font = title_font
# # # # #     ws["A1"].alignment = center

# # # # #     ws.merge_cells("A2:H2")
# # # # #     ws["A2"] = "PACKING LIST DETAILS"
# # # # #     ws["A2"].font = Font(name=FONT, bold=True, size=10)
# # # # #     ws["A2"].alignment = center

# # # # #     # ---- Small size-spec table: SIZE / N.W. / N.N.W. / EMPTY CTN ----
# # # # #     spec_header_row = 4
# # # # #     ws.cell(row=spec_header_row, column=1, value="SIZE").font = bold
# # # # #     ws.cell(row=spec_header_row, column=1).border = box
# # # # #     for sc_idx, size_name in enumerate(sizes):
# # # # #         c = ws.cell(row=spec_header_row, column=2 + sc_idx, value=size_name)
# # # # #         c.font = bold
# # # # #         c.alignment = center
# # # # #         c.border = box
# # # # #         c.fill = grey

# # # # #     spec_rows = ["N.W.", "N.N.W.", "EMPTY CTN"]
# # # # #     for i, label in enumerate(spec_rows):
# # # # #         rr = spec_header_row + 1 + i
# # # # #         lc = ws.cell(row=rr, column=1, value=label)
# # # # #         lc.font = bold
# # # # #         lc.border = box
# # # # #         for sc_idx, size_name in enumerate(sizes):
# # # # #             c = ws.cell(row=rr, column=2 + sc_idx, value=1)
# # # # #             c.alignment = center
# # # # #             c.border = box

# # # # #     r = spec_header_row + len(spec_rows) + 2
# # # # #     ws.cell(row=r, column=1, value="BUYER").font = bold
# # # # #     ws.cell(row=r, column=2, value="VANS").font = normal
# # # # #     r += 1
# # # # #     ws.cell(row=r, column=1, value="LOT NO").font = bold
# # # # #     ws.cell(row=r, column=2, value=line["style"]).font = normal
# # # # #     r += 1
# # # # #     ws.cell(row=r, column=1, value="P.O NO").font = bold
# # # # #     ws.cell(row=r, column=2, value=line["po_line_no"]).font = normal
# # # # #     r += 1
# # # # #     ws.cell(row=r, column=1, value="ORDER QTY").font = bold
# # # # #     ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
# # # # #     ws.cell(row=r, column=3, value="PCS").font = normal
# # # # #     r += 1
# # # # #     ws.cell(row=r, column=1, value="PACK QTY").font = bold
# # # # #     ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
# # # # #     ws.cell(row=r, column=3, value="PCS").font = normal
# # # # #     r += 1
# # # # #     excess_short_row = r  # filled in with a formula once the summary block below is built
# # # # #     ws.cell(row=r, column=1, value="EXCESS/SHORT QTY").font = bold
# # # # #     ws.cell(row=r, column=3, value="PCS").font = normal
# # # # #     r += 1
# # # # #     percentage_row = r  # filled in with a formula once the summary block below is built
# # # # #     ws.cell(row=r, column=1, value="PERCENTAGE").font = bold
# # # # #     r += 1
# # # # #     ws.cell(row=r, column=1, value="CRD").font = bold
# # # # #     ws.cell(row=r, column=2, value=line.get("crd")).font = normal
# # # # #     r += 1
# # # # #     ws.cell(row=r, column=1, value="COUNTRY").font = bold
# # # # #     ws.cell(row=r, column=2, value=line.get("destination_country")).font = normal
# # # # #     r += 1
# # # # #     ws.cell(row=r, column=1, value="DESCRIPTION").font = bold
# # # # #     ws.cell(row=r, column=2, value=line["description"]).font = normal

# # # # #     header_row = r + 2
# # # # #     headers = ["CASE LABEL NO.", "CARTON NO.", "COLOR", "UPC Number"] + sizes + \
# # # # #               ["CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs.wt.pr ctn", "Net.wt.pr ctn",
# # # # #                "total.Grs.wt", "total.net.wt"]
# # # # #     for i, h in enumerate(headers, start=1):
# # # # #         c = ws.cell(row=header_row, column=i, value=h)
# # # # #         c.font = bold
# # # # #         c.alignment = center
# # # # #         c.border = box
# # # # #         c.fill = grey
# # # # #     n_cols = len(headers)

# # # # #     data_row_start = header_row + 1
# # # # #     for idx, s in enumerate(sublines):
# # # # #         rr = data_row_start + idx
# # # # #         ws.cell(row=rr, column=1, value=idx + 1).border = box
# # # # #         cB = ws.cell(row=rr, column=2, value=None)
# # # # #         cB.fill = yellow
# # # # #         cB.border = box
# # # # #         ws.cell(row=rr, column=3, value=s["color"]).border = box
# # # # #         ws.cell(row=rr, column=4, value=s["upc"]).border = box
# # # # #         for sc_idx, size_name in enumerate(sizes):
# # # # #             col = 5 + sc_idx
# # # # #             val = s["qty"] if size_name == s["size"] else None
# # # # #             cell = ws.cell(row=rr, column=col, value=val)
# # # # #             cell.border = box
# # # # #             cell.alignment = center

# # # # #         # CTN PCS = sum of all size columns in THIS row (handles mixed cartons)
# # # # #         ctn_pcs_col = 5 + n_sizes
# # # # #         first_size_letter = get_column_letter(5)
# # # # #         last_size_letter = get_column_letter(4 + n_sizes)
# # # # #         ctn_cell = ws.cell(
# # # # #             row=rr, column=ctn_pcs_col,
# # # # #             value=f"=SUM({first_size_letter}{rr}:{last_size_letter}{rr})"
# # # # #         )
# # # # #         ctn_cell.border = box
# # # # #         ctn_cell.alignment = center
# # # # #         ctn_cell.font = bold

# # # # #         # Remaining columns after CTN PCS: TOTAL CTN, TOTAL PCS, then weights
# # # # #         total_ctn_col = ctn_pcs_col + 1
# # # # #         total_pcs_col = ctn_pcs_col + 2
# # # # #         total_ctn_letter = get_column_letter(total_ctn_col)

# # # # #         # TOTAL CTN stays manual/yellow (user enters actual carton count)
# # # # #         tc_cell = ws.cell(row=rr, column=total_ctn_col, value=None)
# # # # #         tc_cell.fill = yellow
# # # # #         tc_cell.border = box

# # # # #         # TOTAL PCS = TOTAL CTN * CTN PCS (blank until TOTAL CTN is filled in)
# # # # #         ctn_pcs_letter = get_column_letter(ctn_pcs_col)
# # # # #         tp_cell = ws.cell(
# # # # #             row=rr, column=total_pcs_col,
# # # # #             value=f'=IF({total_ctn_letter}{rr}="","",{total_ctn_letter}{rr}*{ctn_pcs_letter}{rr})'
# # # # #         )
# # # # #         tp_cell.border = box
# # # # #         tp_cell.alignment = center

# # # # #         # Grs.wt.pr ctn = SUMPRODUCT(this row's size qtys, N.W. row) + EMPTY CTN
# # # # #         # (EMPTY CTN taken from the first size column, since it's normally
# # # # #         #  the same physical carton regardless of size)
# # # # #         grs_wt_col = total_pcs_col + 1
# # # # #         nw_row = spec_header_row + 1       # N.W. row in the spec table
# # # # #         empty_row = spec_header_row + 3    # EMPTY CTN row in the spec table
# # # # #         first_spec_letter = get_column_letter(2)
# # # # #         last_spec_letter = get_column_letter(1 + n_sizes)
# # # # #         first_main_size_letter = get_column_letter(5)
# # # # #         last_main_size_letter = get_column_letter(4 + n_sizes)

# # # # #         grs_formula = (
# # # # #             f"=SUMPRODUCT({first_main_size_letter}{rr}:{last_main_size_letter}{rr},"
# # # # #             f"{first_spec_letter}{nw_row}:{last_spec_letter}{nw_row})"
# # # # #             f"+{first_spec_letter}{empty_row}"
# # # # #         )
# # # # #         grs_cell = ws.cell(row=rr, column=grs_wt_col, value=grs_formula)
# # # # #         grs_cell.border = box
# # # # #         grs_cell.alignment = center

# # # # #         # Net.wt.pr ctn = Grs.wt.pr ctn - EMPTY CTN (pulled from spec table,
# # # # #         # same single value used above - not a visible column here)
# # # # #         net_wt_col = grs_wt_col + 1
# # # # #         grs_wt_letter_for_net = get_column_letter(grs_wt_col)
# # # # #         net_cell = ws.cell(
# # # # #             row=rr, column=net_wt_col,
# # # # #             value=f"={grs_wt_letter_for_net}{rr}-{first_spec_letter}{empty_row}"
# # # # #         )
# # # # #         net_cell.border = box
# # # # #         net_cell.alignment = center

# # # # #         # total.Grs.wt = TOTAL CTN * Grs.wt.pr ctn (blank until TOTAL CTN filled in)
# # # # #         total_grs_col = net_wt_col + 1
# # # # #         grs_wt_letter = get_column_letter(grs_wt_col)
# # # # #         tg_cell = ws.cell(
# # # # #             row=rr, column=total_grs_col,
# # # # #             value=f'=IF({total_ctn_letter}{rr}="","",{total_ctn_letter}{rr}*{grs_wt_letter}{rr})'
# # # # #         )
# # # # #         tg_cell.border = box
# # # # #         tg_cell.alignment = center

# # # # #         # total.net.wt = TOTAL CTN * Net.wt.pr ctn (blank until TOTAL CTN filled in)
# # # # #         total_net_col = total_grs_col + 1
# # # # #         net_wt_letter = get_column_letter(net_wt_col)
# # # # #         tn_cell = ws.cell(
# # # # #             row=rr, column=total_net_col,
# # # # #             value=f'=IF({total_ctn_letter}{rr}="","",{total_ctn_letter}{rr}*{net_wt_letter}{rr})'
# # # # #         )
# # # # #         tn_cell.border = box
# # # # #         tn_cell.alignment = center

# # # # #     total_row = data_row_start + len(sublines)
# # # # #     ws.cell(row=total_row, column=3, value="TOTAL").font = bold
# # # # #     for sc_idx, size_name in enumerate(sizes):
# # # # #         col = 5 + sc_idx
# # # # #         col_letter = get_column_letter(col)
# # # # #         cell = ws.cell(row=total_row, column=col,
# # # # #                         value=f"=SUM({col_letter}{data_row_start}:{col_letter}{total_row-1})")
# # # # #         cell.font = bold
# # # # #         cell.border = box
# # # # #     # Grand-total TOTAL CTN = sum of each row's TOTAL CTN
# # # # #     # Grand-total TOTAL PCS = sum of each row's TOTAL PCS (not the size columns)
# # # # #     total_ctn_col = ctn_pcs_col + 1
# # # # #     total_pcs_col = ctn_pcs_col + 2
# # # # #     total_ctn_letter = get_column_letter(total_ctn_col)
# # # # #     total_pcs_letter = get_column_letter(total_pcs_col)

# # # # #     ws.cell(row=total_row, column=total_ctn_col,
# # # # #             value=f'=IF(COUNT({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})=0,"",'
# # # # #                   f'SUM({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1}))').font = bold

# # # # #     ws.cell(row=total_row, column=total_pcs_col,
# # # # #             value=f'=IF(COUNT({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1})=0,"",'
# # # # #                   f'SUM({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1}))').font = bold

# # # # #     # Grand-total Grs.wt.pr ctn / Net.wt.pr ctn = sum of each row's value
# # # # #     # (matches the reference image, which sums these two columns directly)
# # # # #     grs_wt_col = total_pcs_col + 1
# # # # #     net_wt_col = total_pcs_col + 2
# # # # #     grs_wt_letter = get_column_letter(grs_wt_col)
# # # # #     net_wt_letter = get_column_letter(net_wt_col)

# # # # #     ws.cell(row=total_row, column=grs_wt_col,
# # # # #             value=f"=SUM({grs_wt_letter}{data_row_start}:{grs_wt_letter}{total_row-1})").font = bold
# # # # #     ws.cell(row=total_row, column=net_wt_col,
# # # # #             value=f"=SUM({net_wt_letter}{data_row_start}:{net_wt_letter}{total_row-1})").font = bold

# # # # #     # Grand-total total.Grs.wt = sum of each row's total.Grs.wt
# # # # #     total_grs_col = total_pcs_col + 3  # skip Grs.wt.pr ctn, Net.wt.pr ctn -> land on total.Grs.wt
# # # # #     total_grs_letter = get_column_letter(total_grs_col)
# # # # #     ws.cell(row=total_row, column=total_grs_col,
# # # # #             value=f'=IF(COUNT({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1})=0,"",'
# # # # #                   f'SUM({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1}))').font = bold

# # # # #     # Grand-total total.net.wt = sum of each row's total.net.wt
# # # # #     total_net_col = total_grs_col + 1
# # # # #     total_net_letter = get_column_letter(total_net_col)
# # # # #     ws.cell(row=total_row, column=total_net_col,
# # # # #             value=f'=IF(COUNT({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1})=0,"",'
# # # # #                   f'SUM({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1}))').font = bold

# # # # #     sum_header_row = total_row + 3
# # # # #     ws.cell(row=sum_header_row, column=1, value="SIZE").font = bold
# # # # #     for sc_idx, size_name in enumerate(sizes):
# # # # #         ws.cell(row=sum_header_row, column=2 + sc_idx, value=size_name).font = bold
# # # # #     ws.cell(row=sum_header_row, column=2 + n_sizes, value="G TOTAL").font = bold

# # # # #     # Two extra header columns: Grand Total gross/net weight (kg)
# # # # #     gw_col = 3 + n_sizes
# # # # #     nw_col = 4 + n_sizes
# # # # #     ws.merge_cells(start_row=sum_header_row, start_column=gw_col,
# # # # #                     end_row=sum_header_row, end_column=gw_col)
# # # # #     ws.cell(row=sum_header_row, column=gw_col, value="Grand Total gross weight (kg)").font = bold
# # # # #     ws.cell(row=sum_header_row, column=gw_col).alignment = center
# # # # #     ws.merge_cells(start_row=sum_header_row, start_column=nw_col,
# # # # #                     end_row=sum_header_row, end_column=nw_col)
# # # # #     ws.cell(row=sum_header_row, column=nw_col, value="Grand Total net weight (kg)").font = bold
# # # # #     ws.cell(row=sum_header_row, column=nw_col).alignment = center

# # # # #     rows_needed = ["CUT QTY", "ORDER QTY", "SHIP QTY", "EXS/SHT QTY", "PERCENTAGE"]
# # # # #     for i, label in enumerate(rows_needed):
# # # # #         rr = sum_header_row + 1 + i
# # # # #         ws.cell(row=rr, column=1, value=label).font = bold
# # # # #         for sc_idx, size_name in enumerate(sizes):
# # # # #             col = 2 + sc_idx
# # # # #             col_letter = get_column_letter(col)
# # # # #             cell = ws.cell(row=rr, column=col)
# # # # #             if label == "ORDER QTY":
# # # # #                 qty = next(s["qty"] for s in sublines if s["size"] == size_name)
# # # # #                 cell.value = qty
# # # # #             elif label in ("CUT QTY", "SHIP QTY"):
# # # # #                 cell.fill = yellow
# # # # #             elif label == "EXS/SHT QTY":
# # # # #                 ship_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
# # # # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # # #                 cell.value = f'=IF({col_letter}{ship_row}="","",{col_letter}{ship_row}-{col_letter}{order_row})'
# # # # #             elif label == "PERCENTAGE":
# # # # #                 exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # # #                 cell.value = f'=IF({col_letter}{exs_row}="","",IFERROR({col_letter}{exs_row}/{col_letter}{order_row},0))'
# # # # #                 cell.number_format = "0.00%"
# # # # #             cell.border = box
# # # # #             cell.alignment = center

# # # # #         gt_col = 2 + n_sizes
# # # # #         gt_letter = get_column_letter(gt_col)
# # # # #         first_letter = get_column_letter(2)
# # # # #         last_letter = get_column_letter(1 + n_sizes)
# # # # #         if label in ("CUT QTY", "SHIP QTY", "EXS/SHT QTY"):
# # # # #             ws.cell(row=rr, column=gt_col,
# # # # #                     value=f'=IF(COUNT({first_letter}{rr}:{last_letter}{rr})=0,"",SUM({first_letter}{rr}:{last_letter}{rr}))').font = bold
# # # # #         elif label == "ORDER QTY":
# # # # #             ws.cell(row=rr, column=gt_col,
# # # # #                     value=f"=SUM({first_letter}{rr}:{last_letter}{rr})").font = bold
# # # # #         else:  # PERCENTAGE
# # # # #             exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # # #             order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # # #             ws.cell(row=rr, column=gt_col,
# # # # #                     value=f'=IF({gt_letter}{exs_row}="","",IFERROR({gt_letter}{exs_row}/{gt_letter}{order_row},0))').number_format = "0.00%"
# # # # #         ws.cell(row=rr, column=gt_col).border = box

# # # # #     # Link the top header's EXCESS/SHORT QTY and PERCENTAGE to the
# # # # #     # summary block's G-Total values (calculated above) - same numbers,
# # # # #     # no separate manual entry needed.
# # # # #     exs_gt_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # # #     order_gt_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # # #     gt_col_letter = get_column_letter(2 + n_sizes)

# # # # #     ws.cell(row=excess_short_row, column=2,
# # # # #             value=f'={gt_col_letter}{exs_gt_row}').font = normal
# # # # #     ws.cell(row=percentage_row, column=2,
# # # # #             value=f'=IF({gt_col_letter}{exs_gt_row}="","-",'
# # # # #                   f'TEXT(IFERROR({gt_col_letter}{exs_gt_row}/{gt_col_letter}{order_gt_row},0),"0.00%"))'
# # # # #             ).font = normal

# # # # #     wt_row = sum_header_row + len(rows_needed) + 2
# # # # #     labels = ["GROSS WEIGHT", "NET WEIGHT", "CTN MEAS", "CBM"]
# # # # #     for i, lbl in enumerate(labels):
# # # # #         ws.cell(row=wt_row + i, column=1, value=lbl + " :").font = bold
# # # # #         ws.cell(row=wt_row + i, column=2, value=None).fill = yellow

# # # # #     # Grand Total gross/net weight (kg) - merged, pulled from the main
# # # # #     # table's total.Grs.wt / total.net.wt grand totals (computed above)
# # # # #     first_sum_row = sum_header_row + 1
# # # # #     last_sum_row = sum_header_row + len(rows_needed)
# # # # #     total_grs_letter = get_column_letter(total_grs_col)
# # # # #     total_net_letter = get_column_letter(total_net_col)

# # # # #     ws.merge_cells(start_row=first_sum_row, start_column=gw_col,
# # # # #                     end_row=last_sum_row, end_column=gw_col)
# # # # #     gw_cell = ws.cell(row=first_sum_row, column=gw_col,
# # # # #                        value=f'={total_grs_letter}{total_row}')
# # # # #     gw_cell.alignment = center
# # # # #     gw_cell.font = bold
# # # # #     gw_cell.border = box

# # # # #     ws.merge_cells(start_row=first_sum_row, start_column=nw_col,
# # # # #                     end_row=last_sum_row, end_column=nw_col)
# # # # #     nw_cell = ws.cell(row=first_sum_row, column=nw_col,
# # # # #                        value=f'={total_net_letter}{total_row}')
# # # # #     nw_cell.alignment = center
# # # # #     nw_cell.font = bold
# # # # #     nw_cell.border = box

# # # # #     ws.column_dimensions["A"].width = 16
# # # # #     ws.column_dimensions["B"].width = 14
# # # # #     ws.column_dimensions["C"].width = 16
# # # # #     ws.column_dimensions["D"].width = 16
# # # # #     for i in range(n_sizes):
# # # # #         ws.column_dimensions[get_column_letter(5 + i)].width = 10
# # # # #     for col in range(5 + n_sizes, n_cols + 1):
# # # # #         ws.column_dimensions[get_column_letter(col)].width = 13
# # # # #     ws.column_dimensions[get_column_letter(gw_col)].width = 22
# # # # #     ws.column_dimensions[get_column_letter(nw_col)].width = 22

# # # # #     legend_row = wt_row + len(labels) + 2
# # # # #     ws.cell(row=legend_row, column=1, value="Legend:").font = bold
# # # # #     ws.cell(row=legend_row + 1, column=1,
# # # # #             value="Yellow cells = data NOT available in the PO PDF (fill in from factory cutting/packing records)"
# # # # #             ).font = Font(name=FONT, size=9, italic=True)


# # # # # wb = Workbook()
# # # # # wb.remove(wb.active)

# # # # # used_names = set()
# # # # # for line in data["lines"]:
# # # # #     sublines = [s for s in data["sublines"] if s["style"] == line["style"]]
# # # # #     if not sublines:
# # # # #         continue
# # # # #     name = sheet_name_for(line)
# # # # #     base_name, i = name, 1
# # # # #     while name in used_names:
# # # # #         i += 1
# # # # #         name = f"{base_name[:28]}_{i}"
# # # # #     used_names.add(name)

# # # # #     ws = wb.create_sheet(title=name)
# # # # #     build_sheet(ws, line, sublines)

# # # # # out_path = "packing_list_all_PO_lines.xlsx"
# # # # # wb.save(out_path)
# # # # # print("Saved:", out_path)
# # # # # print("Sheets created:", wb.sheetnames)


# # # # """
# # # # Build a packing-list-style sheet for EVERY PO line found in
# # # # po_extracted.json -- one worksheet per line, in a single workbook.

# # # # HOW TO RUN (after extract_po.py has created po_extracted.json):
# # # #     python3 build_all_packing_lists.py

# # # # Requires:
# # # #     pip install openpyxl
# # # # """
# # # # import json
# # # # import re
# # # # from openpyxl import Workbook
# # # # from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
# # # # from openpyxl.utils import get_column_letter

# # # # with open("po_extracted.json") as f:
# # # #     data = json.load(f)

# # # # FONT = "Arial"
# # # # bold = Font(name=FONT, bold=True, size=10)
# # # # normal = Font(name=FONT, size=10)
# # # # title_font = Font(name=FONT, bold=True, size=13)
# # # # thin = Side(style="thin")
# # # # box = Border(left=thin, right=thin, top=thin, bottom=thin)
# # # # center = Alignment(horizontal="center", vertical="center", wrap_text=True)
# # # # yellow = PatternFill("solid", fgColor="FFFF00")
# # # # grey = PatternFill("solid", fgColor="D9D9D9")


# # # # def sheet_name_for(line):
# # # #     raw = f"{line['style']}"
# # # #     return re.sub(r'[\[\]:*?/\\]', '-', raw)[:31]


# # # # def compute_carton_rows(sublines):
# # # #     """
# # # #     Splits each subline's qty into full-size cartons (qty // pack) plus a
# # # #     shared remainder pool that gets packed into mixed cartons, filled in
# # # #     size order until each hits pack capacity; the final leftover carton
# # # #     may be under capacity.
# # # #     Returns a list of dicts: {"sizes": {size: qty, ...}, "total_ctn": int,
# # # #     "ctn_pcs": int, "color": str or None, "upc": str or None, "mixed": bool}
# # # #     """
# # # #     if not sublines:
# # # #         return []
# # # #     pack = sublines[0].get("items_per_outer_pack") or 1

# # # #     full_rows = []
# # # #     remainder_pool = []  # [size, qty_left, color, upc]
# # # #     for s in sublines:
# # # #         qty = s["qty"]
# # # #         size = s["size"]
# # # #         full_ctn = qty // pack
# # # #         rem = qty % pack
# # # #         if full_ctn > 0:
# # # #             full_rows.append({
# # # #                 "sizes": {size: pack},
# # # #                 "total_ctn": full_ctn,
# # # #                 "ctn_pcs": pack,
# # # #                 "color": s["color"],
# # # #                 "upc": s["upc"],
# # # #                 "mixed": False,
# # # #             })
# # # #         if rem > 0:
# # # #             remainder_pool.append([size, rem, s["color"], s["upc"]])

# # # #     mixed_rows = []
# # # #     current = {}
# # # #     current_total = 0
# # # #     idx = 0
# # # #     while idx < len(remainder_pool):
# # # #         size, rem, color, upc = remainder_pool[idx]
# # # #         space_left = pack - current_total
# # # #         take = min(rem, space_left)
# # # #         if take > 0:
# # # #             current[size] = current.get(size, 0) + take
# # # #             current_total += take
# # # #             remainder_pool[idx][1] -= take
# # # #         if remainder_pool[idx][1] == 0:
# # # #             idx += 1
# # # #         if current_total == pack:
# # # #             mixed_rows.append({
# # # #                 "sizes": dict(current), "total_ctn": 1, "ctn_pcs": pack,
# # # #                 "color": None, "upc": None, "mixed": True,
# # # #             })
# # # #             current = {}
# # # #             current_total = 0
# # # #     if current:
# # # #         mixed_rows.append({
# # # #             "sizes": dict(current), "total_ctn": 1, "ctn_pcs": current_total,
# # # #             "color": None, "upc": None, "mixed": True,
# # # #         })

# # # #     return full_rows + mixed_rows


# # # # def build_sheet(ws, line, sublines):
# # # #     sizes = [s["size"] for s in sublines]
# # # #     n_sizes = len(sizes)

# # # #     ws.merge_cells("A1:H1")
# # # #     ws["A1"] = "CREATIVE COLLECTIONS LTD. U-1-A"
# # # #     ws["A1"].font = title_font
# # # #     ws["A1"].alignment = center

# # # #     ws.merge_cells("A2:H2")
# # # #     ws["A2"] = "PACKING LIST DETAILS"
# # # #     ws["A2"].font = Font(name=FONT, bold=True, size=10)
# # # #     ws["A2"].alignment = center

# # # #     # ---- Small size-spec table: SIZE / N.W. / N.N.W. / EMPTY CTN ----
# # # #     spec_header_row = 4
# # # #     ws.cell(row=spec_header_row, column=1, value="SIZE").font = bold
# # # #     ws.cell(row=spec_header_row, column=1).border = box
# # # #     for sc_idx, size_name in enumerate(sizes):
# # # #         c = ws.cell(row=spec_header_row, column=2 + sc_idx, value=size_name)
# # # #         c.font = bold
# # # #         c.alignment = center
# # # #         c.border = box
# # # #         c.fill = grey

# # # #     spec_rows = ["N.W.", "N.N.W.", "EMPTY CTN"]
# # # #     for i, label in enumerate(spec_rows):
# # # #         rr = spec_header_row + 1 + i
# # # #         lc = ws.cell(row=rr, column=1, value=label)
# # # #         lc.font = bold
# # # #         lc.border = box
# # # #         for sc_idx, size_name in enumerate(sizes):
# # # #             c = ws.cell(row=rr, column=2 + sc_idx, value=1)
# # # #             c.alignment = center
# # # #             c.border = box

# # # #     r = spec_header_row + len(spec_rows) + 2
# # # #     ws.cell(row=r, column=1, value="BUYER").font = bold
# # # #     ws.cell(row=r, column=2, value="VANS").font = normal
# # # #     r += 1
# # # #     ws.cell(row=r, column=1, value="LOT NO").font = bold
# # # #     ws.cell(row=r, column=2, value=line["style"]).font = normal
# # # #     r += 1
# # # #     ws.cell(row=r, column=1, value="P.O NO").font = bold
# # # #     ws.cell(row=r, column=2, value=line["po_line_no"]).font = normal
# # # #     r += 1
# # # #     ws.cell(row=r, column=1, value="ORDER QTY").font = bold
# # # #     ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
# # # #     ws.cell(row=r, column=3, value="PCS").font = normal
# # # #     r += 1
# # # #     ws.cell(row=r, column=1, value="PACK QTY").font = bold
# # # #     ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
# # # #     ws.cell(row=r, column=3, value="PCS").font = normal
# # # #     r += 1
# # # #     excess_short_row = r  # filled in with a formula once the summary block below is built
# # # #     ws.cell(row=r, column=1, value="EXCESS/SHORT QTY").font = bold
# # # #     ws.cell(row=r, column=3, value="PCS").font = normal
# # # #     r += 1
# # # #     percentage_row = r  # filled in with a formula once the summary block below is built
# # # #     ws.cell(row=r, column=1, value="PERCENTAGE").font = bold
# # # #     r += 1
# # # #     ws.cell(row=r, column=1, value="CRD").font = bold
# # # #     ws.cell(row=r, column=2, value=line.get("crd")).font = normal
# # # #     r += 1
# # # #     ws.cell(row=r, column=1, value="COUNTRY").font = bold
# # # #     ws.cell(row=r, column=2, value=line.get("destination_country")).font = normal
# # # #     r += 1
# # # #     ws.cell(row=r, column=1, value="DESCRIPTION").font = bold
# # # #     ws.cell(row=r, column=2, value=line["description"]).font = normal

# # # #     header_row = r + 2
# # # #     headers = ["CASE LABEL NO.", "CARTON NO.", "COLOR", "UPC Number"] + sizes + \
# # # #               ["CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs.wt.pr ctn", "Net.wt.pr ctn",
# # # #                "total.Grs.wt", "total.net.wt"]
# # # #     for i, h in enumerate(headers, start=1):
# # # #         c = ws.cell(row=header_row, column=i, value=h)
# # # #         c.font = bold
# # # #         c.alignment = center
# # # #         c.border = box
# # # #         c.fill = grey
# # # #     n_cols = len(headers)

# # # #     data_row_start = header_row + 1
# # # #     carton_rows = compute_carton_rows(sublines)
# # # #     for idx, row_data in enumerate(carton_rows):
# # # #         rr = data_row_start + idx
# # # #         ws.cell(row=rr, column=1, value=idx + 1).border = box
# # # #         cB = ws.cell(row=rr, column=2, value=None)
# # # #         cB.fill = yellow
# # # #         cB.border = box
# # # #         color_val = row_data["color"] if row_data["color"] is not None else "MIXED"
# # # #         color_cell = ws.cell(row=rr, column=3, value=color_val)
# # # #         color_cell.border = box
# # # #         if row_data["mixed"]:
# # # #             color_cell.fill = yellow
# # # #         ws.cell(row=rr, column=4, value=row_data["upc"]).border = box
# # # #         for sc_idx, size_name in enumerate(sizes):
# # # #             col = 5 + sc_idx
# # # #             val = row_data["sizes"].get(size_name)
# # # #             cell = ws.cell(row=rr, column=col, value=val)
# # # #             cell.border = box
# # # #             cell.alignment = center

# # # #         # CTN PCS = sum of all size columns in THIS row (handles mixed cartons)
# # # #         ctn_pcs_col = 5 + n_sizes
# # # #         first_size_letter = get_column_letter(5)
# # # #         last_size_letter = get_column_letter(4 + n_sizes)
# # # #         ctn_cell = ws.cell(
# # # #             row=rr, column=ctn_pcs_col,
# # # #             value=f"=SUM({first_size_letter}{rr}:{last_size_letter}{rr})"
# # # #         )
# # # #         ctn_cell.border = box
# # # #         ctn_cell.alignment = center
# # # #         ctn_cell.font = bold

# # # #         # Remaining columns after CTN PCS: TOTAL CTN, TOTAL PCS, then weights
# # # #         total_ctn_col = ctn_pcs_col + 1
# # # #         total_pcs_col = ctn_pcs_col + 2
# # # #         total_ctn_letter = get_column_letter(total_ctn_col)

# # # #         # TOTAL CTN = calculated value (qty // pack, or 1 per mixed carton)
# # # #         tc_cell = ws.cell(row=rr, column=total_ctn_col, value=row_data["total_ctn"])
# # # #         tc_cell.border = box
# # # #         tc_cell.alignment = center

# # # #         # TOTAL PCS = TOTAL CTN * CTN PCS
# # # #         ctn_pcs_letter = get_column_letter(ctn_pcs_col)
# # # #         tp_cell = ws.cell(
# # # #             row=rr, column=total_pcs_col,
# # # #             value=f"={total_ctn_letter}{rr}*{ctn_pcs_letter}{rr}"
# # # #         )
# # # #         tp_cell.border = box
# # # #         tp_cell.alignment = center

# # # #         # Grs.wt.pr ctn = SUMPRODUCT(this row's size qtys, N.W. row) + EMPTY CTN
# # # #         # (EMPTY CTN taken from the first size column, since it's normally
# # # #         #  the same physical carton regardless of size)
# # # #         grs_wt_col = total_pcs_col + 1
# # # #         nw_row = spec_header_row + 1       # N.W. row in the spec table
# # # #         empty_row = spec_header_row + 3    # EMPTY CTN row in the spec table
# # # #         first_spec_letter = get_column_letter(2)
# # # #         last_spec_letter = get_column_letter(1 + n_sizes)
# # # #         first_main_size_letter = get_column_letter(5)
# # # #         last_main_size_letter = get_column_letter(4 + n_sizes)

# # # #         grs_formula = (
# # # #             f"=SUMPRODUCT({first_main_size_letter}{rr}:{last_main_size_letter}{rr},"
# # # #             f"{first_spec_letter}{nw_row}:{last_spec_letter}{nw_row})"
# # # #             f"+{first_spec_letter}{empty_row}"
# # # #         )
# # # #         grs_cell = ws.cell(row=rr, column=grs_wt_col, value=grs_formula)
# # # #         grs_cell.border = box
# # # #         grs_cell.alignment = center

# # # #         # Net.wt.pr ctn = Grs.wt.pr ctn - EMPTY CTN (pulled from spec table,
# # # #         # same single value used above - not a visible column here)
# # # #         net_wt_col = grs_wt_col + 1
# # # #         grs_wt_letter_for_net = get_column_letter(grs_wt_col)
# # # #         net_cell = ws.cell(
# # # #             row=rr, column=net_wt_col,
# # # #             value=f"={grs_wt_letter_for_net}{rr}-{first_spec_letter}{empty_row}"
# # # #         )
# # # #         net_cell.border = box
# # # #         net_cell.alignment = center

# # # #         # total.Grs.wt = TOTAL CTN * Grs.wt.pr ctn
# # # #         total_grs_col = net_wt_col + 1
# # # #         grs_wt_letter = get_column_letter(grs_wt_col)
# # # #         tg_cell = ws.cell(
# # # #             row=rr, column=total_grs_col,
# # # #             value=f"={total_ctn_letter}{rr}*{grs_wt_letter}{rr}"
# # # #         )
# # # #         tg_cell.border = box
# # # #         tg_cell.alignment = center

# # # #         # total.net.wt = TOTAL CTN * Net.wt.pr ctn
# # # #         total_net_col = total_grs_col + 1
# # # #         net_wt_letter = get_column_letter(net_wt_col)
# # # #         tn_cell = ws.cell(
# # # #             row=rr, column=total_net_col,
# # # #             value=f"={total_ctn_letter}{rr}*{net_wt_letter}{rr}"
# # # #         )
# # # #         tn_cell.border = box
# # # #         tn_cell.alignment = center

# # # #     total_row = data_row_start + len(carton_rows)
# # # #     ws.cell(row=total_row, column=3, value="TOTAL").font = bold
# # # #     total_ctn_letter_for_sum = get_column_letter(total_ctn_col)
# # # #     for sc_idx, size_name in enumerate(sizes):
# # # #         col = 5 + sc_idx
# # # #         col_letter = get_column_letter(col)
# # # #         # SUMPRODUCT (not SUM): each row's size qty is a PER-CARTON amount,
# # # #         # so it must be multiplied by that row's TOTAL CTN before summing
# # # #         # (a row with TOTAL CTN=2 represents 2 identical cartons).
# # # #         cell = ws.cell(
# # # #             row=total_row, column=col,
# # # #             value=f"=SUMPRODUCT({col_letter}{data_row_start}:{col_letter}{total_row-1},"
# # # #                   f"{total_ctn_letter_for_sum}{data_row_start}:{total_ctn_letter_for_sum}{total_row-1})"
# # # #         )
# # # #         cell.font = bold
# # # #         cell.border = box
# # # #     # Grand-total TOTAL CTN = sum of each row's TOTAL CTN
# # # #     # Grand-total TOTAL PCS = sum of each row's TOTAL PCS (not the size columns)
# # # #     total_ctn_col = ctn_pcs_col + 1
# # # #     total_pcs_col = ctn_pcs_col + 2
# # # #     total_ctn_letter = get_column_letter(total_ctn_col)
# # # #     total_pcs_letter = get_column_letter(total_pcs_col)

# # # #     ws.cell(row=total_row, column=total_ctn_col,
# # # #             value=f'=IF(COUNT({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})=0,"",'
# # # #                   f'SUM({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1}))').font = bold

# # # #     ws.cell(row=total_row, column=total_pcs_col,
# # # #             value=f'=IF(COUNT({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1})=0,"",'
# # # #                   f'SUM({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1}))').font = bold

# # # #     # Grand-total Grs.wt.pr ctn / Net.wt.pr ctn = sum of each row's value
# # # #     # (matches the reference image, which sums these two columns directly)
# # # #     grs_wt_col = total_pcs_col + 1
# # # #     net_wt_col = total_pcs_col + 2
# # # #     grs_wt_letter = get_column_letter(grs_wt_col)
# # # #     net_wt_letter = get_column_letter(net_wt_col)

# # # #     ws.cell(row=total_row, column=grs_wt_col,
# # # #             value=f"=SUM({grs_wt_letter}{data_row_start}:{grs_wt_letter}{total_row-1})").font = bold
# # # #     ws.cell(row=total_row, column=net_wt_col,
# # # #             value=f"=SUM({net_wt_letter}{data_row_start}:{net_wt_letter}{total_row-1})").font = bold

# # # #     # Grand-total total.Grs.wt = sum of each row's total.Grs.wt
# # # #     total_grs_col = total_pcs_col + 3  # skip Grs.wt.pr ctn, Net.wt.pr ctn -> land on total.Grs.wt
# # # #     total_grs_letter = get_column_letter(total_grs_col)
# # # #     ws.cell(row=total_row, column=total_grs_col,
# # # #             value=f'=IF(COUNT({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1})=0,"",'
# # # #                   f'SUM({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1}))').font = bold

# # # #     # Grand-total total.net.wt = sum of each row's total.net.wt
# # # #     total_net_col = total_grs_col + 1
# # # #     total_net_letter = get_column_letter(total_net_col)
# # # #     ws.cell(row=total_row, column=total_net_col,
# # # #             value=f'=IF(COUNT({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1})=0,"",'
# # # #                   f'SUM({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1}))').font = bold

# # # #     sum_header_row = total_row + 3
# # # #     ws.cell(row=sum_header_row, column=1, value="SIZE").font = bold
# # # #     for sc_idx, size_name in enumerate(sizes):
# # # #         ws.cell(row=sum_header_row, column=2 + sc_idx, value=size_name).font = bold
# # # #     ws.cell(row=sum_header_row, column=2 + n_sizes, value="G TOTAL").font = bold

# # # #     # Two extra header columns: Grand Total gross/net weight (kg)
# # # #     gw_col = 3 + n_sizes
# # # #     nw_col = 4 + n_sizes
# # # #     ws.merge_cells(start_row=sum_header_row, start_column=gw_col,
# # # #                     end_row=sum_header_row, end_column=gw_col)
# # # #     ws.cell(row=sum_header_row, column=gw_col, value="Grand Total gross weight (kg)").font = bold
# # # #     ws.cell(row=sum_header_row, column=gw_col).alignment = center
# # # #     ws.merge_cells(start_row=sum_header_row, start_column=nw_col,
# # # #                     end_row=sum_header_row, end_column=nw_col)
# # # #     ws.cell(row=sum_header_row, column=nw_col, value="Grand Total net weight (kg)").font = bold
# # # #     ws.cell(row=sum_header_row, column=nw_col).alignment = center

# # # #     rows_needed = ["CUT QTY", "ORDER QTY", "SHIP QTY", "EXS/SHT QTY", "PERCENTAGE"]
# # # #     for i, label in enumerate(rows_needed):
# # # #         rr = sum_header_row + 1 + i
# # # #         ws.cell(row=rr, column=1, value=label).font = bold
# # # #         for sc_idx, size_name in enumerate(sizes):
# # # #             col = 2 + sc_idx
# # # #             col_letter = get_column_letter(col)
# # # #             cell = ws.cell(row=rr, column=col)
# # # #             if label == "ORDER QTY":
# # # #                 qty = next(s["qty"] for s in sublines if s["size"] == size_name)
# # # #                 cell.value = qty
# # # #             elif label in ("CUT QTY", "SHIP QTY"):
# # # #                 cell.fill = yellow
# # # #             elif label == "EXS/SHT QTY":
# # # #                 ship_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
# # # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # #                 cell.value = f'=IF({col_letter}{ship_row}="","",{col_letter}{ship_row}-{col_letter}{order_row})'
# # # #             elif label == "PERCENTAGE":
# # # #                 exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # #                 cell.value = f'=IF({col_letter}{exs_row}="","",IFERROR({col_letter}{exs_row}/{col_letter}{order_row},0))'
# # # #                 cell.number_format = "0.00%"
# # # #             cell.border = box
# # # #             cell.alignment = center

# # # #         gt_col = 2 + n_sizes
# # # #         gt_letter = get_column_letter(gt_col)
# # # #         first_letter = get_column_letter(2)
# # # #         last_letter = get_column_letter(1 + n_sizes)
# # # #         if label in ("CUT QTY", "SHIP QTY", "EXS/SHT QTY"):
# # # #             ws.cell(row=rr, column=gt_col,
# # # #                     value=f'=IF(COUNT({first_letter}{rr}:{last_letter}{rr})=0,"",SUM({first_letter}{rr}:{last_letter}{rr}))').font = bold
# # # #         elif label == "ORDER QTY":
# # # #             ws.cell(row=rr, column=gt_col,
# # # #                     value=f"=SUM({first_letter}{rr}:{last_letter}{rr})").font = bold
# # # #         else:  # PERCENTAGE
# # # #             exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # #             order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # #             ws.cell(row=rr, column=gt_col,
# # # #                     value=f'=IF({gt_letter}{exs_row}="","",IFERROR({gt_letter}{exs_row}/{gt_letter}{order_row},0))').number_format = "0.00%"
# # # #         ws.cell(row=rr, column=gt_col).border = box

# # # #     # Link the top header's EXCESS/SHORT QTY and PERCENTAGE to the
# # # #     # summary block's G-Total values (calculated above) - same numbers,
# # # #     # no separate manual entry needed.
# # # #     exs_gt_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # #     order_gt_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # #     gt_col_letter = get_column_letter(2 + n_sizes)

# # # #     ws.cell(row=excess_short_row, column=2,
# # # #             value=f'={gt_col_letter}{exs_gt_row}').font = normal
# # # #     ws.cell(row=percentage_row, column=2,
# # # #             value=f'=IF({gt_col_letter}{exs_gt_row}="","-",'
# # # #                   f'TEXT(IFERROR({gt_col_letter}{exs_gt_row}/{gt_col_letter}{order_gt_row},0),"0.00%"))'
# # # #             ).font = normal

# # # #     wt_row = sum_header_row + len(rows_needed) + 2
# # # #     labels = ["GROSS WEIGHT", "NET WEIGHT", "CTN MEAS", "CBM"]
# # # #     for i, lbl in enumerate(labels):
# # # #         ws.cell(row=wt_row + i, column=1, value=lbl + " :").font = bold
# # # #         ws.cell(row=wt_row + i, column=2, value=None).fill = yellow

# # # #     # Grand Total gross/net weight (kg) - merged, pulled from the main
# # # #     # table's total.Grs.wt / total.net.wt grand totals (computed above)
# # # #     first_sum_row = sum_header_row + 1
# # # #     last_sum_row = sum_header_row + len(rows_needed)
# # # #     total_grs_letter = get_column_letter(total_grs_col)
# # # #     total_net_letter = get_column_letter(total_net_col)

# # # #     ws.merge_cells(start_row=first_sum_row, start_column=gw_col,
# # # #                     end_row=last_sum_row, end_column=gw_col)
# # # #     gw_cell = ws.cell(row=first_sum_row, column=gw_col,
# # # #                        value=f'={total_grs_letter}{total_row}')
# # # #     gw_cell.alignment = center
# # # #     gw_cell.font = bold
# # # #     gw_cell.border = box

# # # #     ws.merge_cells(start_row=first_sum_row, start_column=nw_col,
# # # #                     end_row=last_sum_row, end_column=nw_col)
# # # #     nw_cell = ws.cell(row=first_sum_row, column=nw_col,
# # # #                        value=f'={total_net_letter}{total_row}')
# # # #     nw_cell.alignment = center
# # # #     nw_cell.font = bold
# # # #     nw_cell.border = box

# # # #     ws.column_dimensions["A"].width = 16
# # # #     ws.column_dimensions["B"].width = 14
# # # #     ws.column_dimensions["C"].width = 16
# # # #     ws.column_dimensions["D"].width = 16
# # # #     for i in range(n_sizes):
# # # #         ws.column_dimensions[get_column_letter(5 + i)].width = 10
# # # #     for col in range(5 + n_sizes, n_cols + 1):
# # # #         ws.column_dimensions[get_column_letter(col)].width = 13
# # # #     ws.column_dimensions[get_column_letter(gw_col)].width = 22
# # # #     ws.column_dimensions[get_column_letter(nw_col)].width = 22

# # # #     legend_row = wt_row + len(labels) + 2
# # # #     ws.cell(row=legend_row, column=1, value="Legend:").font = bold
# # # #     ws.cell(row=legend_row + 1, column=1,
# # # #             value="Yellow cells = data NOT available in the PO PDF (fill in from factory cutting/packing records)"
# # # #             ).font = Font(name=FONT, size=9, italic=True)


# # # # wb = Workbook()
# # # # wb.remove(wb.active)

# # # # used_names = set()
# # # # for line in data["lines"]:
# # # #     sublines = [s for s in data["sublines"] if s["style"] == line["style"]]
# # # #     if not sublines:
# # # #         continue
# # # #     name = sheet_name_for(line)
# # # #     base_name, i = name, 1
# # # #     while name in used_names:
# # # #         i += 1
# # # #         name = f"{base_name[:28]}_{i}"
# # # #     used_names.add(name)

# # # #     ws = wb.create_sheet(title=name)
# # # #     build_sheet(ws, line, sublines)

# # # # out_path = "packing_list_all_PO_lines.xlsx"
# # # # wb.save(out_path)
# # # # print("Saved:", out_path)
# # # # print("Sheets created:", wb.sheetnames)


# # # """
# # # Build a packing-list-style sheet for EVERY PO line found in
# # # po_extracted.json -- one worksheet per line, in a single workbook.

# # # HOW TO RUN (after extract_po.py has created po_extracted.json):
# # #     python3 build_all_packing_lists.py

# # # Requires:
# # #     pip install openpyxl
# # # """


# # # # import json
# # # # import re
# # # # from openpyxl import Workbook
# # # # from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
# # # # from openpyxl.utils import get_column_letter

# # # # with open("po_extracted.json") as f:
# # # #     data = json.load(f)

# # # # FONT = "Arial"
# # # # bold = Font(name=FONT, bold=True, size=10)
# # # # normal = Font(name=FONT, size=10)
# # # # title_font = Font(name=FONT, bold=True, size=14, color="1F4E79")
# # # # subtitle_font = Font(name=FONT, bold=True, size=11, color="2E75B6")
# # # # header_font = Font(name=FONT, bold=True, size=9, color="1F4E79")
# # # # label_font = Font(name=FONT, bold=True, size=10, color="333333")
# # # # thin = Side(style="thin", color="B4C6E7")
# # # # medium = Side(style="medium", color="5B9BD5")
# # # # box = Border(left=thin, right=thin, top=thin, bottom=thin)
# # # # header_box = Border(left=thin, right=thin, top=medium, bottom=medium)
# # # # center = Alignment(horizontal="center", vertical="center", wrap_text=True)
# # # # left_center = Alignment(horizontal="left", vertical="center", wrap_text=True)

# # # # # Light color palette
# # # # header_fill = PatternFill("solid", fgColor="D6E3F0")      # soft blue header
# # # # alt_row_fill = PatternFill("solid", fgColor="F2F7FB")     # very light blue alt rows
# # # # spec_fill = PatternFill("solid", fgColor="E2EFDA")        # soft green for size-spec
# # # # yellow = PatternFill("solid", fgColor="FFF2CC")           # soft yellow for inputs
# # # # title_fill = PatternFill("solid", fgColor="D6E3F0")
# # # # total_fill = PatternFill("solid", fgColor="DDEBF7")       # light blue for totals
# # # # summary_header_fill = PatternFill("solid", fgColor="C6EFCE")  # soft green summary


# # # # def sheet_name_for(line):
# # # #     raw = f"{line['style']}"
# # # #     return re.sub(r'[\[\]:*?/\\]', '-', raw)[:31]


# # # # def compute_carton_rows(sublines):
# # # #     """
# # # #     Splits each subline's qty into full-size cartons (qty // pack) plus a
# # # #     shared remainder pool that gets packed into mixed cartons, filled in
# # # #     size order until each hits pack capacity; the final leftover carton
# # # #     may be under capacity.
# # # #     Returns a list of dicts: {"sizes": {size: qty, ...}, "total_ctn": int,
# # # #     "ctn_pcs": int, "color": str or None, "upc": str or None, "mixed": bool}
# # # #     """
# # # #     if not sublines:
# # # #         return []
# # # #     pack = sublines[0].get("items_per_outer_pack") or 1

# # # #     full_rows = []
# # # #     remainder_pool = []  # [size, qty_left, color, upc]
# # # #     for s in sublines:
# # # #         qty = s["qty"]
# # # #         size = s["size"]
# # # #         full_ctn = qty // pack
# # # #         rem = qty % pack
# # # #         if full_ctn > 0:
# # # #             full_rows.append({
# # # #                 "sizes": {size: pack},
# # # #                 "total_ctn": full_ctn,
# # # #                 "ctn_pcs": pack,
# # # #                 "color": s["color"],
# # # #                 "upc": s["upc"],
# # # #                 "mixed": False,
# # # #             })
# # # #         if rem > 0:
# # # #             remainder_pool.append([size, rem, s["color"], s["upc"]])

# # # #     mixed_rows = []
# # # #     current = {}
# # # #     current_total = 0
# # # #     idx = 0
# # # #     while idx < len(remainder_pool):
# # # #         size, rem, color, upc = remainder_pool[idx]
# # # #         space_left = pack - current_total
# # # #         take = min(rem, space_left)
# # # #         if take > 0:
# # # #             current[size] = current.get(size, 0) + take
# # # #             current_total += take
# # # #             remainder_pool[idx][1] -= take
# # # #         if remainder_pool[idx][1] == 0:
# # # #             idx += 1
# # # #         if current_total == pack:
# # # #             mixed_rows.append({
# # # #                 "sizes": dict(current), "total_ctn": 1, "ctn_pcs": pack,
# # # #                 "color": None, "upc": None, "mixed": True,
# # # #             })
# # # #             current = {}
# # # #             current_total = 0
# # # #     if current:
# # # #         mixed_rows.append({
# # # #             "sizes": dict(current), "total_ctn": 1, "ctn_pcs": current_total,
# # # #             "color": None, "upc": None, "mixed": True,
# # # #         })

# # # #     return full_rows + mixed_rows


# # # # def style_cell(cell, font=None, fill=None, border=None, alignment=None):
# # # #     if font is not None:
# # # #         cell.font = font
# # # #     if fill is not None:
# # # #         cell.fill = fill
# # # #     if border is not None:
# # # #         cell.border = border
# # # #     if alignment is not None:
# # # #         cell.alignment = alignment


# # # # def build_sheet(ws, line, sublines):
# # # #     sizes = [s["size"] for s in sublines]
# # # #     n_sizes = len(sizes)

# # # #     # ---- Title block ----
# # # #     ws.merge_cells("A1:H1")
# # # #     ws["A1"] = "CREATIVE COLLECTIONS LTD. U-1-A"
# # # #     style_cell(ws["A1"], font=title_font, fill=title_fill, alignment=center)
# # # #     ws.row_dimensions[1].height = 22

# # # #     ws.merge_cells("A2:H2")
# # # #     ws["A2"] = "PACKING LIST DETAILS"
# # # #     style_cell(ws["A2"], font=subtitle_font, fill=title_fill, alignment=center)
# # # #     ws.row_dimensions[2].height = 18

# # # #     # ---- Small size-spec table: SIZE / N.W. / N.N.W. / EMPTY CTN ----
# # # #     spec_header_row = 4
# # # #     c = ws.cell(row=spec_header_row, column=1, value="SIZE")
# # # #     style_cell(c, font=header_font, fill=spec_fill, border=box, alignment=center)
# # # #     for sc_idx, size_name in enumerate(sizes):
# # # #         c = ws.cell(row=spec_header_row, column=2 + sc_idx, value=size_name)
# # # #         style_cell(c, font=header_font, fill=spec_fill, border=box, alignment=center)

# # # #     spec_rows = ["N.W.", "N.N.W.", "EMPTY CTN"]
# # # #     for i, label in enumerate(spec_rows):
# # # #         rr = spec_header_row + 1 + i
# # # #         lc = ws.cell(row=rr, column=1, value=label)
# # # #         style_cell(lc, font=bold, fill=spec_fill, border=box, alignment=center)
# # # #         for sc_idx, size_name in enumerate(sizes):
# # # #             c = ws.cell(row=rr, column=2 + sc_idx, value=1)
# # # #             style_cell(c, font=normal, border=box, alignment=center)

# # # #     # ---- Info block (horizontal layout) ----
# # # #     r = spec_header_row + len(spec_rows) + 2

# # # #     # Row 1 labels
# # # #     info_labels_1 = ["BUYER", "LOT NO", "P.O NO", "ORDER QTY", "PACK QTY", "EXCESS/SHORT QTY", "PERCENTAGE"]
# # # #     for col_idx, label in enumerate(info_labels_1, start=1):
# # # #         c = ws.cell(row=r, column=col_idx, value=label)
# # # #         style_cell(c, font=header_font, fill=header_fill, border=box, alignment=center)
# # # #     r += 1

# # # #     # Row 1 values
# # # #     info_vals_1 = [
# # # #         "VANS",
# # # #         line["style"],
# # # #         line["po_line_no"],
# # # #         f'{line["order_qty"]} PCS',
# # # #         f'{line["order_qty"]} PCS',
# # # #         None,  # EXCESS/SHORT – filled later with formula
# # # #         None,  # PERCENTAGE – filled later with formula
# # # #     ]
# # # #     excess_short_row = r
# # # #     percentage_col = 7  # column for PERCENTAGE value
# # # #     excess_col = 6      # column for EXCESS/SHORT value
# # # #     for col_idx, val in enumerate(info_vals_1, start=1):
# # # #         c = ws.cell(row=r, column=col_idx, value=val)
# # # #         style_cell(c, font=normal, border=box, alignment=center)
# # # #     percentage_row = r
# # # #     r += 1

# # # #     # Row 2 labels
# # # #     info_labels_2 = ["CRD", "COUNTRY", "DESCRIPTION"]
# # # #     for col_idx, label in enumerate(info_labels_2, start=1):
# # # #         c = ws.cell(row=r, column=col_idx, value=label)
# # # #         style_cell(c, font=header_font, fill=header_fill, border=box, alignment=center)
# # # #     r += 1

# # # #     # Row 2 values
# # # #     info_vals_2 = [
# # # #         line.get("crd"),
# # # #         line.get("destination_country"),
# # # #         line["description"],
# # # #     ]
# # # #     for col_idx, val in enumerate(info_vals_2, start=1):
# # # #         c = ws.cell(row=r, column=col_idx, value=val)
# # # #         style_cell(c, font=normal, border=box, alignment=center)
# # # #         if col_idx == 3:
# # # #             # DESCRIPTION can be long – merge a few columns
# # # #             ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
# # # #             c.alignment = center
# # # #     r += 1

# # # #     # ---- Main data header ----
# # # #     header_row = r + 2
# # # #     headers = ["CASE LABEL NO.", "CARTON NO.", "COLOR", "UPC Number"] + sizes + \
# # # #               ["CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs.wt.pr ctn", "Net.wt.pr ctn",
# # # #                "total.Grs.wt", "total.net.wt"]
# # # #     for i, h in enumerate(headers, start=1):
# # # #         c = ws.cell(row=header_row, column=i, value=h)
# # # #         style_cell(c, font=header_font, fill=header_fill, border=header_box, alignment=center)
# # # #     n_cols = len(headers)
# # # #     ws.row_dimensions[header_row].height = 30

# # # #     # ---- Data rows (carton rows) ----
# # # #     data_row_start = header_row + 1
# # # #     carton_rows = compute_carton_rows(sublines)
# # # #     for idx, row_data in enumerate(carton_rows):
# # # #         rr = data_row_start + idx
# # # #         row_fill = alt_row_fill if idx % 2 == 1 else None

# # # #         c = ws.cell(row=rr, column=1, value=idx + 1)
# # # #         style_cell(c, font=normal, fill=row_fill, border=box, alignment=center)

# # # #         cB = ws.cell(row=rr, column=2, value=None)
# # # #         style_cell(cB, fill=yellow, border=box, alignment=center)

# # # #         color_val = row_data["color"] if row_data["color"] is not None else "MIXED"
# # # #         color_cell = ws.cell(row=rr, column=3, value=color_val)
# # # #         style_cell(color_cell, font=normal, border=box, alignment=center,
# # # #                    fill=yellow if row_data["mixed"] else row_fill)

# # # #         upc_cell = ws.cell(row=rr, column=4, value=row_data["upc"])
# # # #         style_cell(upc_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # # #         for sc_idx, size_name in enumerate(sizes):
# # # #             col = 5 + sc_idx
# # # #             val = row_data["sizes"].get(size_name)
# # # #             cell = ws.cell(row=rr, column=col, value=val)
# # # #             style_cell(cell, font=normal, fill=row_fill, border=box, alignment=center)

# # # #         # CTN PCS
# # # #         ctn_pcs_col = 5 + n_sizes
# # # #         first_size_letter = get_column_letter(5)
# # # #         last_size_letter = get_column_letter(4 + n_sizes)
# # # #         ctn_cell = ws.cell(
# # # #             row=rr, column=ctn_pcs_col,
# # # #             value=f"=SUM({first_size_letter}{rr}:{last_size_letter}{rr})"
# # # #         )
# # # #         style_cell(ctn_cell, font=bold, fill=row_fill, border=box, alignment=center)

# # # #         total_ctn_col = ctn_pcs_col + 1
# # # #         total_pcs_col = ctn_pcs_col + 2
# # # #         total_ctn_letter = get_column_letter(total_ctn_col)

# # # #         tc_cell = ws.cell(row=rr, column=total_ctn_col, value=row_data["total_ctn"])
# # # #         style_cell(tc_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # # #         ctn_pcs_letter = get_column_letter(ctn_pcs_col)
# # # #         tp_cell = ws.cell(
# # # #             row=rr, column=total_pcs_col,
# # # #             value=f"={total_ctn_letter}{rr}*{ctn_pcs_letter}{rr}"
# # # #         )
# # # #         style_cell(tp_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # # #         # Grs.wt.pr ctn
# # # #         grs_wt_col = total_pcs_col + 1
# # # #         nw_row = spec_header_row + 1
# # # #         empty_row = spec_header_row + 3
# # # #         first_spec_letter = get_column_letter(2)
# # # #         last_spec_letter = get_column_letter(1 + n_sizes)
# # # #         first_main_size_letter = get_column_letter(5)
# # # #         last_main_size_letter = get_column_letter(4 + n_sizes)

# # # #         grs_formula = (
# # # #             f"=SUMPRODUCT({first_main_size_letter}{rr}:{last_main_size_letter}{rr},"
# # # #             f"{first_spec_letter}{nw_row}:{last_spec_letter}{nw_row})"
# # # #             f"+{first_spec_letter}{empty_row}"
# # # #         )
# # # #         grs_cell = ws.cell(row=rr, column=grs_wt_col, value=grs_formula)
# # # #         style_cell(grs_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # # #         # Net.wt.pr ctn
# # # #         net_wt_col = grs_wt_col + 1
# # # #         grs_wt_letter_for_net = get_column_letter(grs_wt_col)
# # # #         net_cell = ws.cell(
# # # #             row=rr, column=net_wt_col,
# # # #             value=f"={grs_wt_letter_for_net}{rr}-{first_spec_letter}{empty_row}"
# # # #         )
# # # #         style_cell(net_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # # #         # total.Grs.wt / total.net.wt
# # # #         total_grs_col = net_wt_col + 1
# # # #         grs_wt_letter = get_column_letter(grs_wt_col)
# # # #         tg_cell = ws.cell(
# # # #             row=rr, column=total_grs_col,
# # # #             value=f"={total_ctn_letter}{rr}*{grs_wt_letter}{rr}"
# # # #         )
# # # #         style_cell(tg_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # # #         total_net_col = total_grs_col + 1
# # # #         net_wt_letter = get_column_letter(net_wt_col)
# # # #         tn_cell = ws.cell(
# # # #             row=rr, column=total_net_col,
# # # #             value=f"={total_ctn_letter}{rr}*{net_wt_letter}{rr}"
# # # #         )
# # # #         style_cell(tn_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # # #     # ---- TOTAL row ----
# # # #     total_row = data_row_start + len(carton_rows)
# # # #     total_label = ws.cell(row=total_row, column=3, value="TOTAL")
# # # #     style_cell(total_label, font=bold, fill=total_fill, alignment=center)

# # # #     total_ctn_letter_for_sum = get_column_letter(total_ctn_col)
# # # #     for sc_idx, size_name in enumerate(sizes):
# # # #         col = 5 + sc_idx
# # # #         col_letter = get_column_letter(col)
# # # #         cell = ws.cell(
# # # #             row=total_row, column=col,
# # # #             value=f"=SUMPRODUCT({col_letter}{data_row_start}:{col_letter}{total_row-1},"
# # # #                   f"{total_ctn_letter_for_sum}{data_row_start}:{total_ctn_letter_for_sum}{total_row-1})"
# # # #         )
# # # #         style_cell(cell, font=bold, fill=total_fill, border=box, alignment=center)

# # # #     total_ctn_col = ctn_pcs_col + 1
# # # #     total_pcs_col = ctn_pcs_col + 2
# # # #     total_ctn_letter = get_column_letter(total_ctn_col)
# # # #     total_pcs_letter = get_column_letter(total_pcs_col)

# # # #     for col, formula in [
# # # #         (total_ctn_col,
# # # #          f'=IF(COUNT({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})=0,"",'
# # # #          f'SUM({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1}))'),
# # # #         (total_pcs_col,
# # # #          f'=IF(COUNT({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1})=0,"",'
# # # #          f'SUM({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1}))'),
# # # #     ]:
# # # #         c = ws.cell(row=total_row, column=col, value=formula)
# # # #         style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# # # #     grs_wt_col = total_pcs_col + 1
# # # #     net_wt_col = total_pcs_col + 2
# # # #     grs_wt_letter = get_column_letter(grs_wt_col)
# # # #     net_wt_letter = get_column_letter(net_wt_col)

# # # #     for col, letter in [(grs_wt_col, grs_wt_letter), (net_wt_col, net_wt_letter)]:
# # # #         c = ws.cell(row=total_row, column=col,
# # # #                     value=f"=SUM({letter}{data_row_start}:{letter}{total_row-1})")
# # # #         style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# # # #     total_grs_col = total_pcs_col + 3
# # # #     total_grs_letter = get_column_letter(total_grs_col)
# # # #     c = ws.cell(row=total_row, column=total_grs_col,
# # # #                 value=f'=IF(COUNT({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1})=0,"",'
# # # #                       f'SUM({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1}))')
# # # #     style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# # # #     total_net_col = total_grs_col + 1
# # # #     total_net_letter = get_column_letter(total_net_col)
# # # #     c = ws.cell(row=total_row, column=total_net_col,
# # # #                 value=f'=IF(COUNT({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1})=0,"",'
# # # #                       f'SUM({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1}))')
# # # #     style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# # # #     # Apply fill/border to empty total-row cells for visual consistency
# # # #     for col in range(1, n_cols + 1):
# # # #         cell = ws.cell(row=total_row, column=col)
# # # #         if cell.fill.fgColor is None or cell.fill.fgColor.rgb == "00000000":
# # # #             cell.fill = total_fill
# # # #         if cell.border.left.style is None:
# # # #             cell.border = box
# # # #         cell.alignment = center

# # # #     # ---- Summary block (SIZE / CUT QTY / ORDER QTY ...) ----
# # # #     sum_header_row = total_row + 3
# # # #     c = ws.cell(row=sum_header_row, column=1, value="SIZE")
# # # #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
# # # #     for sc_idx, size_name in enumerate(sizes):
# # # #         c = ws.cell(row=sum_header_row, column=2 + sc_idx, value=size_name)
# # # #         style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
# # # #     c = ws.cell(row=sum_header_row, column=2 + n_sizes, value="G TOTAL")
# # # #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)

# # # #     gw_col = 3 + n_sizes
# # # #     nw_col = 4 + n_sizes
# # # #     ws.merge_cells(start_row=sum_header_row, start_column=gw_col,
# # # #                     end_row=sum_header_row, end_column=gw_col)
# # # #     c = ws.cell(row=sum_header_row, column=gw_col, value="Grand Total gross weight (kg)")
# # # #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
# # # #     ws.merge_cells(start_row=sum_header_row, start_column=nw_col,
# # # #                     end_row=sum_header_row, end_column=nw_col)
# # # #     c = ws.cell(row=sum_header_row, column=nw_col, value="Grand Total net weight (kg)")
# # # #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)

# # # #     rows_needed = ["CUT QTY", "ORDER QTY", "SHIP QTY", "EXS/SHT QTY", "PERCENTAGE"]
# # # #     for i, label in enumerate(rows_needed):
# # # #         rr = sum_header_row + 1 + i
# # # #         lc = ws.cell(row=rr, column=1, value=label)
# # # #         style_cell(lc, font=bold, border=box, alignment=center)
# # # #         for sc_idx, size_name in enumerate(sizes):
# # # #             col = 2 + sc_idx
# # # #             col_letter = get_column_letter(col)
# # # #             cell = ws.cell(row=rr, column=col)
# # # #             if label == "ORDER QTY":
# # # #                 qty = next(s["qty"] for s in sublines if s["size"] == size_name)
# # # #                 cell.value = qty
# # # #             elif label in ("CUT QTY", "SHIP QTY"):
# # # #                 cell.fill = yellow
# # # #             elif label == "EXS/SHT QTY":
# # # #                 ship_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
# # # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # #                 cell.value = f'=IF({col_letter}{ship_row}="","",{col_letter}{ship_row}-{col_letter}{order_row})'
# # # #             elif label == "PERCENTAGE":
# # # #                 exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # #                 cell.value = f'=IF({col_letter}{exs_row}="","",IFERROR({col_letter}{exs_row}/{col_letter}{order_row},0))'
# # # #                 cell.number_format = "0.00%"
# # # #             style_cell(cell, font=normal, border=box, alignment=center)
# # # #             if label in ("CUT QTY", "SHIP QTY") and cell.fill.fgColor is None:
# # # #                 cell.fill = yellow

# # # #         gt_col = 2 + n_sizes
# # # #         gt_letter = get_column_letter(gt_col)
# # # #         first_letter = get_column_letter(2)
# # # #         last_letter = get_column_letter(1 + n_sizes)
# # # #         if label in ("CUT QTY", "SHIP QTY", "EXS/SHT QTY"):
# # # #             c = ws.cell(row=rr, column=gt_col,
# # # #                         value=f'=IF(COUNT({first_letter}{rr}:{last_letter}{rr})=0,"",SUM({first_letter}{rr}:{last_letter}{rr}))')
# # # #             style_cell(c, font=bold, border=box, alignment=center)
# # # #         elif label == "ORDER QTY":
# # # #             c = ws.cell(row=rr, column=gt_col,
# # # #                         value=f"=SUM({first_letter}{rr}:{last_letter}{rr})")
# # # #             style_cell(c, font=bold, border=box, alignment=center)
# # # #         else:  # PERCENTAGE
# # # #             exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # #             order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # #             c = ws.cell(row=rr, column=gt_col,
# # # #                         value=f'=IF({gt_letter}{exs_row}="","",IFERROR({gt_letter}{exs_row}/{gt_letter}{order_row},0))')
# # # #             c.number_format = "0.00%"
# # # #             style_cell(c, font=bold, border=box, alignment=center)

# # # #     # Link top header EXCESS/SHORT & PERCENTAGE (horizontal info row)
# # # #     exs_gt_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # # #     order_gt_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # # #     gt_col_letter = get_column_letter(2 + n_sizes)

# # # #     c = ws.cell(row=excess_short_row, column=excess_col,
# # # #                 value=f'={gt_col_letter}{exs_gt_row}')
# # # #     style_cell(c, font=normal, border=box, alignment=center)
# # # #     c = ws.cell(row=percentage_row, column=percentage_col,
# # # #                 value=f'=IF({gt_col_letter}{exs_gt_row}="","-",'
# # # #                       f'TEXT(IFERROR({gt_col_letter}{exs_gt_row}/{gt_col_letter}{order_gt_row},0),"0.00%"))')
# # # #     style_cell(c, font=normal, border=box, alignment=center)

# # # #     # ---- Weight / CBM block ----
# # # #     wt_row = sum_header_row + len(rows_needed) + 2
# # # #     labels = ["GROSS WEIGHT", "NET WEIGHT", "CTN MEAS", "CBM"]
# # # #     for i, lbl in enumerate(labels):
# # # #         lc = ws.cell(row=wt_row + i, column=1, value=lbl + " :")
# # # #         style_cell(lc, font=label_font, alignment=center)
# # # #         vc = ws.cell(row=wt_row + i, column=2, value=None)
# # # #         style_cell(vc, fill=yellow, border=box, alignment=center)

# # # #     # Grand Total gross/net weight (kg) - merged
# # # #     first_sum_row = sum_header_row + 1
# # # #     last_sum_row = sum_header_row + len(rows_needed)
# # # #     total_grs_letter = get_column_letter(total_grs_col)
# # # #     total_net_letter = get_column_letter(total_net_col)

# # # #     ws.merge_cells(start_row=first_sum_row, start_column=gw_col,
# # # #                     end_row=last_sum_row, end_column=gw_col)
# # # #     gw_cell = ws.cell(row=first_sum_row, column=gw_col,
# # # #                        value=f'={total_grs_letter}{total_row}')
# # # #     style_cell(gw_cell, font=bold, fill=total_fill, border=box, alignment=center)

# # # #     ws.merge_cells(start_row=first_sum_row, start_column=nw_col,
# # # #                     end_row=last_sum_row, end_column=nw_col)
# # # #     nw_cell = ws.cell(row=first_sum_row, column=nw_col,
# # # #                        value=f'={total_net_letter}{total_row}')
# # # #     style_cell(nw_cell, font=bold, fill=total_fill, border=box, alignment=center)

# # # #     # ---- Column widths ----
# # # #     ws.column_dimensions["A"].width = 18
# # # #     ws.column_dimensions["B"].width = 14
# # # #     ws.column_dimensions["C"].width = 16
# # # #     ws.column_dimensions["D"].width = 16
# # # #     for i in range(n_sizes):
# # # #         ws.column_dimensions[get_column_letter(5 + i)].width = 10
# # # #     for col in range(5 + n_sizes, n_cols + 1):
# # # #         ws.column_dimensions[get_column_letter(col)].width = 13
# # # #     ws.column_dimensions[get_column_letter(gw_col)].width = 24
# # # #     ws.column_dimensions[get_column_letter(nw_col)].width = 24

# # # #     # ---- Legend ----
# # # #     legend_row = wt_row + len(labels) + 2
# # # #     ws.cell(row=legend_row, column=1, value="Legend:").font = bold
# # # #     ws.cell(row=legend_row, column=1).alignment = center
# # # #     legend = ws.cell(row=legend_row + 1, column=1,
# # # #                      value="Yellow cells = data NOT available in the PO PDF (fill in from factory cutting/packing records)")
# # # #     legend.font = Font(name=FONT, size=9, italic=True, color="666666")
# # # #     legend.alignment = left_center


# # # # wb = Workbook()
# # # # wb.remove(wb.active)

# # # # used_names = set()
# # # # for line in data["lines"]:
# # # #     sublines = [s for s in data["sublines"] if s["style"] == line["style"]]
# # # #     if not sublines:
# # # #         continue
# # # #     name = sheet_name_for(line)
# # # #     base_name, i = name, 1
# # # #     while name in used_names:
# # # #         i += 1
# # # #         name = f"{base_name[:28]}_{i}"
# # # #     used_names.add(name)

# # # #     ws = wb.create_sheet(title=name)
# # # #     build_sheet(ws, line, sublines)

# # # # out_path = "packing_list_all_PO_lines.xlsx"
# # # # wb.save(out_path)
# # # # print("Saved:", out_path)
# # # # print("Sheets created:", wb.sheetnames)




# # # """
# # # Build a packing-list-style sheet for EVERY PO line found in
# # # po_extracted.json -- one worksheet per line, in a single workbook.

# # # HOW TO RUN (after extract_po.py has created po_extracted.json):
# # #     python3 build_all_packing_lists.py

# # # Requires:
# # #     pip install openpyxl
# # # """
# # # import json
# # # import re
# # # from openpyxl import Workbook
# # # from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
# # # from openpyxl.utils import get_column_letter

# # # with open("po_extracted.json") as f:
# # #     data = json.load(f)

# # # FONT = "Arial"
# # # bold = Font(name=FONT, bold=True, size=10)
# # # normal = Font(name=FONT, size=10)
# # # title_font = Font(name=FONT, bold=True, size=14, color="1F4E79")
# # # subtitle_font = Font(name=FONT, bold=True, size=11, color="2E75B6")
# # # header_font = Font(name=FONT, bold=True, size=9, color="1F4E79")
# # # label_font = Font(name=FONT, bold=True, size=10, color="333333")
# # # thin = Side(style="thin", color="B4C6E7")
# # # medium = Side(style="medium", color="5B9BD5")
# # # box = Border(left=thin, right=thin, top=thin, bottom=thin)
# # # header_box = Border(left=thin, right=thin, top=medium, bottom=medium)
# # # center = Alignment(horizontal="center", vertical="center", wrap_text=True)
# # # left_center = Alignment(horizontal="left", vertical="center", wrap_text=True)

# # # # Light color palette
# # # header_fill = PatternFill("solid", fgColor="D6E3F0")      # soft blue header
# # # alt_row_fill = PatternFill("solid", fgColor="F2F7FB")     # very light blue alt rows
# # # spec_fill = PatternFill("solid", fgColor="E2EFDA")        # soft green for size-spec
# # # yellow = PatternFill("solid", fgColor="FFF2CC")           # soft yellow for inputs
# # # title_fill = PatternFill("solid", fgColor="D6E3F0")
# # # total_fill = PatternFill("solid", fgColor="DDEBF7")       # light blue for totals
# # # summary_header_fill = PatternFill("solid", fgColor="C6EFCE")  # soft green summary


# # # def sheet_name_for(line):
# # #     raw = f"{line['style']}"
# # #     return re.sub(r'[\[\]:*?/\\]', '-', raw)[:31]


# # # def compute_carton_rows(sublines):
# # #     """
# # #     Splits each subline's qty into full-size cartons (qty // pack) plus a
# # #     shared remainder pool that gets packed into mixed cartons, filled in
# # #     size order until each hits pack capacity; the final leftover carton
# # #     may be under capacity.
# # #     Returns a list of dicts: {"sizes": {size: qty, ...}, "total_ctn": int,
# # #     "ctn_pcs": int, "color": str or None, "upc": str or None, "mixed": bool}
# # #     """
# # #     if not sublines:
# # #         return []
# # #     pack = sublines[0].get("items_per_outer_pack") or 1

# # #     full_rows = []
# # #     remainder_pool = []  # [size, qty_left, color, upc]
# # #     for s in sublines:
# # #         qty = s["qty"]
# # #         size = s["size"]
# # #         full_ctn = qty // pack
# # #         rem = qty % pack
# # #         if full_ctn > 0:
# # #             full_rows.append({
# # #                 "sizes": {size: pack},
# # #                 "total_ctn": full_ctn,
# # #                 "ctn_pcs": pack,
# # #                 "color": s["color"],
# # #                 "upc": s["upc"],
# # #                 "mixed": False,
# # #             })
# # #         if rem > 0:
# # #             remainder_pool.append([size, rem, s["color"], s["upc"]])

# # #     mixed_rows = []
# # #     current = {}
# # #     current_total = 0
# # #     idx = 0
# # #     while idx < len(remainder_pool):
# # #         size, rem, color, upc = remainder_pool[idx]
# # #         space_left = pack - current_total
# # #         take = min(rem, space_left)
# # #         if take > 0:
# # #             current[size] = current.get(size, 0) + take
# # #             current_total += take
# # #             remainder_pool[idx][1] -= take
# # #         if remainder_pool[idx][1] == 0:
# # #             idx += 1
# # #         if current_total == pack:
# # #             mixed_rows.append({
# # #                 "sizes": dict(current), "total_ctn": 1, "ctn_pcs": pack,
# # #                 "color": None, "upc": None, "mixed": True,
# # #             })
# # #             current = {}
# # #             current_total = 0
# # #     if current:
# # #         mixed_rows.append({
# # #             "sizes": dict(current), "total_ctn": 1, "ctn_pcs": current_total,
# # #             "color": None, "upc": None, "mixed": True,
# # #         })

# # #     return full_rows + mixed_rows


# # # def style_cell(cell, font=None, fill=None, border=None, alignment=None):
# # #     if font is not None:
# # #         cell.font = font
# # #     if fill is not None:
# # #         cell.fill = fill
# # #     if border is not None:
# # #         cell.border = border
# # #     if alignment is not None:
# # #         cell.alignment = alignment


# # # def build_sheet(ws, line, sublines):
# # #     sizes = [s["size"] for s in sublines]
# # #     n_sizes = len(sizes)

# # #     # ---- Title block ----
# # #     ws.merge_cells("A1:H1")
# # #     ws["A1"] = "CREATIVE COLLECTIONS LTD. U-1-A"
# # #     style_cell(ws["A1"], font=title_font, fill=title_fill, alignment=center)
# # #     ws.row_dimensions[1].height = 22

# # #     ws.merge_cells("A2:H2")
# # #     ws["A2"] = "PACKING LIST DETAILS"
# # #     style_cell(ws["A2"], font=subtitle_font, fill=title_fill, alignment=center)
# # #     ws.row_dimensions[2].height = 18

# # #     # ---- Small size-spec table: SIZE / N.W. / N.N.W. / EMPTY CTN ----
# # #     spec_header_row = 4
# # #     c = ws.cell(row=spec_header_row, column=1, value="SIZE")
# # #     style_cell(c, font=header_font, fill=spec_fill, border=box, alignment=center)
# # #     for sc_idx, size_name in enumerate(sizes):
# # #         c = ws.cell(row=spec_header_row, column=2 + sc_idx, value=size_name)
# # #         style_cell(c, font=header_font, fill=spec_fill, border=box, alignment=center)

# # #     spec_rows = ["N.W.", "N.N.W.", "EMPTY CTN"]
# # #     for i, label in enumerate(spec_rows):
# # #         rr = spec_header_row + 1 + i
# # #         lc = ws.cell(row=rr, column=1, value=label)
# # #         style_cell(lc, font=bold, fill=spec_fill, border=box, alignment=center)
# # #         for sc_idx, size_name in enumerate(sizes):
# # #             c = ws.cell(row=rr, column=2 + sc_idx, value=1)
# # #             style_cell(c, font=normal, border=box, alignment=center)

# # #     # ---- Info block (horizontal layout) ----
# # #     r = spec_header_row + len(spec_rows) + 2

# # #     # Row 1 labels
# # #     info_labels_1 = ["BUYER", "LOT NO", "P.O NO", "ORDER QTY", "PACK QTY", "EXCESS/SHORT QTY", "PERCENTAGE"]
# # #     for col_idx, label in enumerate(info_labels_1, start=1):
# # #         c = ws.cell(row=r, column=col_idx, value=label)
# # #         style_cell(c, font=header_font, fill=header_fill, border=box, alignment=center)
# # #     r += 1

# # #     # Row 1 values
# # #     info_vals_1 = [
# # #         "VANS",
# # #         line["style"],
# # #         line["po_line_no"],
# # #         f'{line["order_qty"]} PCS',
# # #         f'{line["order_qty"]} PCS',
# # #         None,  # EXCESS/SHORT – filled later with formula
# # #         None,  # PERCENTAGE – filled later with formula
# # #     ]
# # #     excess_short_row = r
# # #     percentage_col = 7  # column for PERCENTAGE value
# # #     excess_col = 6      # column for EXCESS/SHORT value
# # #     for col_idx, val in enumerate(info_vals_1, start=1):
# # #         c = ws.cell(row=r, column=col_idx, value=val)
# # #         style_cell(c, font=normal, border=box, alignment=center)
# # #     percentage_row = r
# # #     r += 1

# # #     # Row 2 labels
# # #     info_labels_2 = ["CRD", "COUNTRY", "DESCRIPTION"]
# # #     for col_idx, label in enumerate(info_labels_2, start=1):
# # #         c = ws.cell(row=r, column=col_idx, value=label)
# # #         style_cell(c, font=header_font, fill=header_fill, border=box, alignment=center)
# # #     r += 1

# # #     # Row 2 values
# # #     info_vals_2 = [
# # #         line.get("crd"),
# # #         line.get("destination_country"),
# # #         line["description"],
# # #     ]
# # #     for col_idx, val in enumerate(info_vals_2, start=1):
# # #         c = ws.cell(row=r, column=col_idx, value=val)
# # #         style_cell(c, font=normal, border=box, alignment=center)
# # #         if col_idx == 3:
# # #             # DESCRIPTION can be long – merge a few columns
# # #             ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
# # #             c.alignment = center
# # #     r += 1

# # #     # ---- Main data header ----
# # #     header_row = r + 2
# # #     headers = ["CASE LABEL NO.", "CARTON NO.", "COLOR", "UPC Number"] + sizes + \
# # #               ["CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs.wt.pr ctn", "Net.wt.pr ctn",
# # #                "total.Grs.wt", "total.net.wt"]
# # #     for i, h in enumerate(headers, start=1):
# # #         c = ws.cell(row=header_row, column=i, value=h)
# # #         style_cell(c, font=header_font, fill=header_fill, border=header_box, alignment=center)
# # #     n_cols = len(headers)
# # #     ws.row_dimensions[header_row].height = 30

# # #     # ---- Data rows (carton rows) ----
# # #     data_row_start = header_row + 1
# # #     carton_rows = compute_carton_rows(sublines)
# # #     for idx, row_data in enumerate(carton_rows):
# # #         rr = data_row_start + idx
# # #         row_fill = alt_row_fill if idx % 2 == 1 else None

# # #         c = ws.cell(row=rr, column=1, value=idx + 1)
# # #         style_cell(c, font=normal, fill=row_fill, border=box, alignment=center)

# # #         cB = ws.cell(row=rr, column=2, value=None)
# # #         style_cell(cB, fill=yellow, border=box, alignment=center)

# # #         color_val = row_data["color"] if row_data["color"] is not None else "MIXED"
# # #         color_cell = ws.cell(row=rr, column=3, value=color_val)
# # #         style_cell(color_cell, font=normal, border=box, alignment=center,
# # #                    fill=yellow if row_data["mixed"] else row_fill)

# # #         upc_cell = ws.cell(row=rr, column=4, value=row_data["upc"])
# # #         style_cell(upc_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # #         for sc_idx, size_name in enumerate(sizes):
# # #             col = 5 + sc_idx
# # #             val = row_data["sizes"].get(size_name)
# # #             cell = ws.cell(row=rr, column=col, value=val)
# # #             style_cell(cell, font=normal, fill=row_fill, border=box, alignment=center)

# # #         # CTN PCS
# # #         ctn_pcs_col = 5 + n_sizes
# # #         first_size_letter = get_column_letter(5)
# # #         last_size_letter = get_column_letter(4 + n_sizes)
# # #         ctn_cell = ws.cell(
# # #             row=rr, column=ctn_pcs_col,
# # #             value=f"=SUM({first_size_letter}{rr}:{last_size_letter}{rr})"
# # #         )
# # #         style_cell(ctn_cell, font=bold, fill=row_fill, border=box, alignment=center)

# # #         total_ctn_col = ctn_pcs_col + 1
# # #         total_pcs_col = ctn_pcs_col + 2
# # #         total_ctn_letter = get_column_letter(total_ctn_col)

# # #         tc_cell = ws.cell(row=rr, column=total_ctn_col, value=row_data["total_ctn"])
# # #         style_cell(tc_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # #         ctn_pcs_letter = get_column_letter(ctn_pcs_col)
# # #         tp_cell = ws.cell(
# # #             row=rr, column=total_pcs_col,
# # #             value=f"={total_ctn_letter}{rr}*{ctn_pcs_letter}{rr}"
# # #         )
# # #         style_cell(tp_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # #         # Grs.wt.pr ctn
# # #         grs_wt_col = total_pcs_col + 1
# # #         nw_row = spec_header_row + 1
# # #         empty_row = spec_header_row + 3
# # #         first_spec_letter = get_column_letter(2)
# # #         last_spec_letter = get_column_letter(1 + n_sizes)
# # #         first_main_size_letter = get_column_letter(5)
# # #         last_main_size_letter = get_column_letter(4 + n_sizes)

# # #         grs_formula = (
# # #             f"=SUMPRODUCT({first_main_size_letter}{rr}:{last_main_size_letter}{rr},"
# # #             f"{first_spec_letter}{nw_row}:{last_spec_letter}{nw_row})"
# # #             f"+{first_spec_letter}{empty_row}"
# # #         )
# # #         grs_cell = ws.cell(row=rr, column=grs_wt_col, value=grs_formula)
# # #         style_cell(grs_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # #         # Net.wt.pr ctn
# # #         net_wt_col = grs_wt_col + 1
# # #         grs_wt_letter_for_net = get_column_letter(grs_wt_col)
# # #         net_cell = ws.cell(
# # #             row=rr, column=net_wt_col,
# # #             value=f"={grs_wt_letter_for_net}{rr}-{first_spec_letter}{empty_row}"
# # #         )
# # #         style_cell(net_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # #         # total.Grs.wt / total.net.wt
# # #         total_grs_col = net_wt_col + 1
# # #         grs_wt_letter = get_column_letter(grs_wt_col)
# # #         tg_cell = ws.cell(
# # #             row=rr, column=total_grs_col,
# # #             value=f"={total_ctn_letter}{rr}*{grs_wt_letter}{rr}"
# # #         )
# # #         style_cell(tg_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # #         total_net_col = total_grs_col + 1
# # #         net_wt_letter = get_column_letter(net_wt_col)
# # #         tn_cell = ws.cell(
# # #             row=rr, column=total_net_col,
# # #             value=f"={total_ctn_letter}{rr}*{net_wt_letter}{rr}"
# # #         )
# # #         style_cell(tn_cell, font=normal, fill=row_fill, border=box, alignment=center)

# # #     # ---- TOTAL row ----
# # #     total_row = data_row_start + len(carton_rows)
# # #     total_label = ws.cell(row=total_row, column=3, value="TOTAL")
# # #     style_cell(total_label, font=bold, fill=total_fill, alignment=center)

# # #     total_ctn_letter_for_sum = get_column_letter(total_ctn_col)
# # #     for sc_idx, size_name in enumerate(sizes):
# # #         col = 5 + sc_idx
# # #         col_letter = get_column_letter(col)
# # #         cell = ws.cell(
# # #             row=total_row, column=col,
# # #             value=f"=SUMPRODUCT({col_letter}{data_row_start}:{col_letter}{total_row-1},"
# # #                   f"{total_ctn_letter_for_sum}{data_row_start}:{total_ctn_letter_for_sum}{total_row-1})"
# # #         )
# # #         style_cell(cell, font=bold, fill=total_fill, border=box, alignment=center)

# # #     total_ctn_col = ctn_pcs_col + 1
# # #     total_pcs_col = ctn_pcs_col + 2
# # #     total_ctn_letter = get_column_letter(total_ctn_col)
# # #     total_pcs_letter = get_column_letter(total_pcs_col)

# # #     for col, formula in [
# # #         (total_ctn_col,
# # #          f'=IF(COUNT({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})=0,"",'
# # #          f'SUM({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1}))'),
# # #         (total_pcs_col,
# # #          f'=IF(COUNT({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1})=0,"",'
# # #          f'SUM({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1}))'),
# # #     ]:
# # #         c = ws.cell(row=total_row, column=col, value=formula)
# # #         style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# # #     grs_wt_col = total_pcs_col + 1
# # #     net_wt_col = total_pcs_col + 2
# # #     grs_wt_letter = get_column_letter(grs_wt_col)
# # #     net_wt_letter = get_column_letter(net_wt_col)

# # #     for col, letter in [(grs_wt_col, grs_wt_letter), (net_wt_col, net_wt_letter)]:
# # #         c = ws.cell(row=total_row, column=col,
# # #                     value=f"=SUM({letter}{data_row_start}:{letter}{total_row-1})")
# # #         style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# # #     total_grs_col = total_pcs_col + 3
# # #     total_grs_letter = get_column_letter(total_grs_col)
# # #     c = ws.cell(row=total_row, column=total_grs_col,
# # #                 value=f'=IF(COUNT({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1})=0,"",'
# # #                       f'SUM({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1}))')
# # #     style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# # #     total_net_col = total_grs_col + 1
# # #     total_net_letter = get_column_letter(total_net_col)
# # #     c = ws.cell(row=total_row, column=total_net_col,
# # #                 value=f'=IF(COUNT({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1})=0,"",'
# # #                       f'SUM({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1}))')
# # #     style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# # #     # Apply fill/border to empty total-row cells for visual consistency
# # #     for col in range(1, n_cols + 1):
# # #         cell = ws.cell(row=total_row, column=col)
# # #         if cell.fill.fgColor is None or cell.fill.fgColor.rgb == "00000000":
# # #             cell.fill = total_fill
# # #         if cell.border.left.style is None:
# # #             cell.border = box
# # #         cell.alignment = center

# # #     # ---- Summary block (SIZE / CUT QTY / ORDER QTY / SHIP QTY ...) ----
# # #     sum_header_row = total_row + 3
# # #     c = ws.cell(row=sum_header_row, column=1, value="SIZE")
# # #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
# # #     for sc_idx, size_name in enumerate(sizes):
# # #         c = ws.cell(row=sum_header_row, column=2 + sc_idx, value=size_name)
# # #         style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
# # #     c = ws.cell(row=sum_header_row, column=2 + n_sizes, value="G TOTAL")
# # #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)

# # #     gw_col = 3 + n_sizes
# # #     nw_col = 4 + n_sizes
# # #     ws.merge_cells(start_row=sum_header_row, start_column=gw_col,
# # #                     end_row=sum_header_row, end_column=gw_col)
# # #     c = ws.cell(row=sum_header_row, column=gw_col, value="Grand Total gross weight (kg)")
# # #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
# # #     ws.merge_cells(start_row=sum_header_row, start_column=nw_col,
# # #                     end_row=sum_header_row, end_column=nw_col)
# # #     c = ws.cell(row=sum_header_row, column=nw_col, value="Grand Total net weight (kg)")
# # #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)

# # #     rows_needed = ["CUT QTY", "ORDER QTY", "SHIP QTY", "EXS/SHT QTY", "PERCENTAGE"]
# # #     for i, label in enumerate(rows_needed):
# # #         rr = sum_header_row + 1 + i
# # #         lc = ws.cell(row=rr, column=1, value=label)
# # #         style_cell(lc, font=bold, border=box, alignment=center)
# # #         for sc_idx, size_name in enumerate(sizes):
# # #             col = 2 + sc_idx
# # #             col_letter = get_column_letter(col)
# # #             cell = ws.cell(row=rr, column=col)
# # #             if label == "ORDER QTY":
# # #                 qty = next(s["qty"] for s in sublines if s["size"] == size_name)
# # #                 cell.value = qty
# # #             elif label == "CUT QTY":
# # #                 cell.fill = yellow
# # #             elif label == "SHIP QTY":
# # #                 # SHIP QTY = packet-wise total from main table TOTAL row
# # #                 # (SUMPRODUCT of size qty × TOTAL CTN already computed there)
# # #                 main_size_col = 5 + sc_idx
# # #                 main_size_letter = get_column_letter(main_size_col)
# # #                 cell.value = f"={main_size_letter}{total_row}"
# # #             elif label == "EXS/SHT QTY":
# # #                 ship_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
# # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # #                 cell.value = f'=IF({col_letter}{ship_row}="","",{col_letter}{ship_row}-{col_letter}{order_row})'
# # #             elif label == "PERCENTAGE":
# # #                 exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # #                 cell.value = f'=IF({col_letter}{exs_row}="","",IFERROR({col_letter}{exs_row}/{col_letter}{order_row},0))'
# # #                 cell.number_format = "0.00%"
# # #             style_cell(cell, font=normal, border=box, alignment=center)
# # #             if label == "CUT QTY":
# # #                 cell.fill = yellow

# # #         gt_col = 2 + n_sizes
# # #         gt_letter = get_column_letter(gt_col)
# # #         first_letter = get_column_letter(2)
# # #         last_letter = get_column_letter(1 + n_sizes)
# # #         if label in ("CUT QTY", "EXS/SHT QTY"):
# # #             c = ws.cell(row=rr, column=gt_col,
# # #                         value=f'=IF(COUNT({first_letter}{rr}:{last_letter}{rr})=0,"",SUM({first_letter}{rr}:{last_letter}{rr}))')
# # #             style_cell(c, font=bold, border=box, alignment=center)
# # #         elif label in ("ORDER QTY", "SHIP QTY"):
# # #             c = ws.cell(row=rr, column=gt_col,
# # #                         value=f"=SUM({first_letter}{rr}:{last_letter}{rr})")
# # #             style_cell(c, font=bold, border=box, alignment=center)
# # #         else:  # PERCENTAGE
# # #             exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # #             order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # #             c = ws.cell(row=rr, column=gt_col,
# # #                         value=f'=IF({gt_letter}{exs_row}="","",IFERROR({gt_letter}{exs_row}/{gt_letter}{order_row},0))')
# # #             c.number_format = "0.00%"
# # #             style_cell(c, font=bold, border=box, alignment=center)

# # #     # Link top header EXCESS/SHORT & PERCENTAGE (horizontal info row)
# # #     exs_gt_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# # #     order_gt_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# # #     gt_col_letter = get_column_letter(2 + n_sizes)

# # #     c = ws.cell(row=excess_short_row, column=excess_col,
# # #                 value=f'={gt_col_letter}{exs_gt_row}')
# # #     style_cell(c, font=normal, border=box, alignment=center)
# # #     c = ws.cell(row=percentage_row, column=percentage_col,
# # #                 value=f'=IF({gt_col_letter}{exs_gt_row}="","-",'
# # #                       f'TEXT(IFERROR({gt_col_letter}{exs_gt_row}/{gt_col_letter}{order_gt_row},0),"0.00%"))')
# # #     style_cell(c, font=normal, border=box, alignment=center)

# # #     # ---- Weight / CBM block ----
# # #     wt_row = sum_header_row + len(rows_needed) + 2
# # #     labels = ["GROSS WEIGHT", "NET WEIGHT", "CTN MEAS", "CBM"]
# # #     for i, lbl in enumerate(labels):
# # #         lc = ws.cell(row=wt_row + i, column=1, value=lbl + " :")
# # #         style_cell(lc, font=label_font, alignment=center)
# # #         vc = ws.cell(row=wt_row + i, column=2, value=None)
# # #         style_cell(vc, fill=yellow, border=box, alignment=center)

# # #     # Grand Total gross/net weight (kg) - merged
# # #     first_sum_row = sum_header_row + 1
# # #     last_sum_row = sum_header_row + len(rows_needed)
# # #     total_grs_letter = get_column_letter(total_grs_col)
# # #     total_net_letter = get_column_letter(total_net_col)

# # #     ws.merge_cells(start_row=first_sum_row, start_column=gw_col,
# # #                     end_row=last_sum_row, end_column=gw_col)
# # #     gw_cell = ws.cell(row=first_sum_row, column=gw_col,
# # #                        value=f'={total_grs_letter}{total_row}')
# # #     style_cell(gw_cell, font=bold, fill=total_fill, border=box, alignment=center)

# # #     ws.merge_cells(start_row=first_sum_row, start_column=nw_col,
# # #                     end_row=last_sum_row, end_column=nw_col)
# # #     nw_cell = ws.cell(row=first_sum_row, column=nw_col,
# # #                        value=f'={total_net_letter}{total_row}')
# # #     style_cell(nw_cell, font=bold, fill=total_fill, border=box, alignment=center)

# # #     # ---- Column widths ----
# # #     ws.column_dimensions["A"].width = 18
# # #     ws.column_dimensions["B"].width = 14
# # #     ws.column_dimensions["C"].width = 16
# # #     ws.column_dimensions["D"].width = 16
# # #     for i in range(n_sizes):
# # #         ws.column_dimensions[get_column_letter(5 + i)].width = 10
# # #     for col in range(5 + n_sizes, n_cols + 1):
# # #         ws.column_dimensions[get_column_letter(col)].width = 13
# # #     ws.column_dimensions[get_column_letter(gw_col)].width = 24
# # #     ws.column_dimensions[get_column_letter(nw_col)].width = 24

# # #     # ---- Legend ----
# # #     legend_row = wt_row + len(labels) + 2
# # #     ws.cell(row=legend_row, column=1, value="Legend:").font = bold
# # #     ws.cell(row=legend_row, column=1).alignment = center
# # #     legend = ws.cell(row=legend_row + 1, column=1,
# # #                      value="Yellow cells = data NOT available in the PO PDF (fill in from factory cutting/packing records)")
# # #     legend.font = Font(name=FONT, size=9, italic=True, color="666666")
# # #     legend.alignment = left_center


# # # wb = Workbook()
# # # wb.remove(wb.active)

# # # used_names = set()
# # # for line in data["lines"]:
# # #     sublines = [s for s in data["sublines"] if s["style"] == line["style"]]
# # #     if not sublines:
# # #         continue
# # #     name = sheet_name_for(line)
# # #     base_name, i = name, 1
# # #     while name in used_names:
# # #         i += 1
# # #         name = f"{base_name[:28]}_{i}"
# # #     used_names.add(name)

# # #     ws = wb.create_sheet(title=name)
# # #     build_sheet(ws, line, sublines)

# # # out_path = "packing_list_all_PO_lines.xlsx"
# # # wb.save(out_path)
# # # print("Saved:", out_path)
# # # print("Sheets created:", wb.sheetnames)
# # """
# # Build a packing-list-style sheet for EVERY PO line found in
# # po_extracted.json -- one worksheet per line, in a single workbook.

# # HOW TO RUN (after extract_po.py has created po_extracted.json):
# #     python3 build_all_packing_lists.py

# # Requires:
# #     pip install openpyxl
# # """
# # import json
# # import re
# # from openpyxl import Workbook
# # from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
# # from openpyxl.utils import get_column_letter

# # with open("po_extracted.json") as f:
# #     data = json.load(f)

# # FONT = "Arial"
# # bold = Font(name=FONT, bold=True, size=10)
# # normal = Font(name=FONT, size=10)
# # title_font = Font(name=FONT, bold=True, size=14, color="1F4E79")
# # subtitle_font = Font(name=FONT, bold=True, size=11, color="2E75B6")
# # header_font = Font(name=FONT, bold=True, size=9, color="1F4E79")
# # label_font = Font(name=FONT, bold=True, size=10, color="333333")
# # thin = Side(style="thin", color="B4C6E7")
# # medium = Side(style="medium", color="5B9BD5")
# # box = Border(left=thin, right=thin, top=thin, bottom=thin)
# # header_box = Border(left=thin, right=thin, top=medium, bottom=medium)
# # center = Alignment(horizontal="center", vertical="center", wrap_text=True)
# # left_center = Alignment(horizontal="left", vertical="center", wrap_text=True)

# # # Light color palette
# # header_fill = PatternFill("solid", fgColor="D6E3F0")
# # alt_row_fill = PatternFill("solid", fgColor="F2F7FB")
# # spec_fill = PatternFill("solid", fgColor="E2EFDA")
# # yellow = PatternFill("solid", fgColor="FFF2CC")
# # title_fill = PatternFill("solid", fgColor="D6E3F0")
# # total_fill = PatternFill("solid", fgColor="DDEBF7")
# # summary_header_fill = PatternFill("solid", fgColor="C6EFCE")


# # def sheet_name_for(line):
# #     raw = f"{line['style']}"
# #     return re.sub(r'[\[\]:*?/\\]', '-', raw)[:31]


# # def compute_carton_rows(sublines):
# #     if not sublines:
# #         return []
# #     pack = sublines[0].get("items_per_outer_pack") or 1

# #     full_rows = []
# #     remainder_pool = []
# #     for s in sublines:
# #         qty = s["qty"]
# #         size = s["size"]
# #         full_ctn = qty // pack
# #         rem = qty % pack
# #         if full_ctn > 0:
# #             full_rows.append({
# #                 "sizes": {size: pack},
# #                 "total_ctn": full_ctn,
# #                 "ctn_pcs": pack,
# #                 "color": s["color"],
# #                 "upc": s["upc"],
# #                 "mixed": False,
# #             })
# #         if rem > 0:
# #             remainder_pool.append([size, rem, s["color"], s["upc"]])

# #     mixed_rows = []
# #     current = {}
# #     current_total = 0
# #     idx = 0
# #     while idx < len(remainder_pool):
# #         size, rem, color, upc = remainder_pool[idx]
# #         space_left = pack - current_total
# #         take = min(rem, space_left)
# #         if take > 0:
# #             current[size] = current.get(size, 0) + take
# #             current_total += take
# #             remainder_pool[idx][1] -= take
# #         if remainder_pool[idx][1] == 0:
# #             idx += 1
# #         if current_total == pack:
# #             mixed_rows.append({
# #                 "sizes": dict(current), "total_ctn": 1, "ctn_pcs": pack,
# #                 "color": None, "upc": None, "mixed": True,
# #             })
# #             current = {}
# #             current_total = 0
# #     if current:
# #         mixed_rows.append({
# #             "sizes": dict(current), "total_ctn": 1, "ctn_pcs": current_total,
# #             "color": None, "upc": None, "mixed": True,
# #         })

# #     return full_rows + mixed_rows


# # def style_cell(cell, font=None, fill=None, border=None, alignment=None):
# #     if font is not None:
# #         cell.font = font
# #     if fill is not None:
# #         cell.fill = fill
# #     if border is not None:
# #         cell.border = border
# #     if alignment is not None:
# #         cell.alignment = alignment


# # def build_sheet(ws, line, sublines):
# #     sizes = [s["size"] for s in sublines]
# #     n_sizes = len(sizes)

# #     # ---- Title block ----
# #     ws.merge_cells("A1:H1")
# #     ws["A1"] = "CREATIVE COLLECTIONS LTD. U-1-A"
# #     style_cell(ws["A1"], font=title_font, fill=title_fill, alignment=center)
# #     ws.row_dimensions[1].height = 22

# #     ws.merge_cells("A2:H2")
# #     ws["A2"] = "PACKING LIST DETAILS"
# #     style_cell(ws["A2"], font=subtitle_font, fill=title_fill, alignment=center)
# #     ws.row_dimensions[2].height = 18

# #     # ---- Size-spec table ----
# #     spec_header_row = 4
# #     c = ws.cell(row=spec_header_row, column=1, value="SIZE")
# #     style_cell(c, font=header_font, fill=spec_fill, border=box, alignment=center)
# #     for sc_idx, size_name in enumerate(sizes):
# #         c = ws.cell(row=spec_header_row, column=2 + sc_idx, value=size_name)
# #         style_cell(c, font=header_font, fill=spec_fill, border=box, alignment=center)

# #     spec_rows = ["N.W.", "N.N.W.", "EMPTY CTN"]
# #     for i, label in enumerate(spec_rows):
# #         rr = spec_header_row + 1 + i
# #         lc = ws.cell(row=rr, column=1, value=label)
# #         style_cell(lc, font=bold, fill=spec_fill, border=box, alignment=center)
# #         for sc_idx, size_name in enumerate(sizes):
# #             c = ws.cell(row=rr, column=2 + sc_idx, value=1)
# #             style_cell(c, font=normal, border=box, alignment=center)

# #     # ---- Info block (horizontal) ----
# #     r = spec_header_row + len(spec_rows) + 2

# #     info_labels_1 = ["BUYER", "LOT NO", "P.O NO", "ORDER QTY", "PACK QTY", "EXCESS/SHORT QTY", "PERCENTAGE"]
# #     for col_idx, label in enumerate(info_labels_1, start=1):
# #         c = ws.cell(row=r, column=col_idx, value=label)
# #         style_cell(c, font=header_font, fill=header_fill, border=box, alignment=center)
# #     r += 1

# #     info_vals_1 = [
# #         "VANS",
# #         line["style"],
# #         line["po_line_no"],
# #         f'{line["order_qty"]} PCS',
# #         f'{line["order_qty"]} PCS',
# #         None,
# #         None,
# #     ]
# #     excess_short_row = r
# #     percentage_col = 7
# #     excess_col = 6
# #     for col_idx, val in enumerate(info_vals_1, start=1):
# #         c = ws.cell(row=r, column=col_idx, value=val)
# #         style_cell(c, font=normal, border=box, alignment=center)
# #     percentage_row = r
# #     r += 1

# #     info_labels_2 = ["CRD", "COUNTRY", "DESCRIPTION"]
# #     for col_idx, label in enumerate(info_labels_2, start=1):
# #         c = ws.cell(row=r, column=col_idx, value=label)
# #         style_cell(c, font=header_font, fill=header_fill, border=box, alignment=center)
# #     r += 1

# #     info_vals_2 = [
# #         line.get("crd"),
# #         line.get("destination_country"),
# #         line["description"],
# #     ]
# #     for col_idx, val in enumerate(info_vals_2, start=1):
# #         c = ws.cell(row=r, column=col_idx, value=val)
# #         style_cell(c, font=normal, border=box, alignment=center)
# #         if col_idx == 3:
# #             ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
# #             c.alignment = center
# #     r += 1

# #     # ---- Main data header ----
# #     header_row = r + 2
# #     headers = ["CASE LABEL NO.", "CARTON NO.", "COLOR", "UPC Number"] + sizes + \
# #               ["CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs.wt.pr ctn", "Net.wt.pr ctn",
# #                "total.Grs.wt", "total.net.wt"]
# #     for i, h in enumerate(headers, start=1):
# #         c = ws.cell(row=header_row, column=i, value=h)
# #         style_cell(c, font=header_font, fill=header_fill, border=header_box, alignment=center)
# #     n_cols = len(headers)
# #     ws.row_dimensions[header_row].height = 30

# #     # ---- Data rows ----
# #     data_row_start = header_row + 1
# #     carton_rows = compute_carton_rows(sublines)
# #     for idx, row_data in enumerate(carton_rows):
# #         rr = data_row_start + idx
# #         row_fill = alt_row_fill if idx % 2 == 1 else None

# #         c = ws.cell(row=rr, column=1, value=idx + 1)
# #         style_cell(c, font=normal, fill=row_fill, border=box, alignment=center)

# #         color_val = row_data["color"] if row_data["color"] is not None else "MIXED"
# #         color_cell = ws.cell(row=rr, column=3, value=color_val)
# #         style_cell(color_cell, font=normal, border=box, alignment=center,
# #                    fill=yellow if row_data["mixed"] else row_fill)

# #         upc_cell = ws.cell(row=rr, column=4, value=row_data["upc"])
# #         style_cell(upc_cell, font=normal, fill=row_fill, border=box, alignment=center)

# #         for sc_idx, size_name in enumerate(sizes):
# #             col = 5 + sc_idx
# #             val = row_data["sizes"].get(size_name)
# #             cell = ws.cell(row=rr, column=col, value=val)
# #             style_cell(cell, font=normal, fill=row_fill, border=box, alignment=center)

# #         ctn_pcs_col = 5 + n_sizes
# #         first_size_letter = get_column_letter(5)
# #         last_size_letter = get_column_letter(4 + n_sizes)
# #         ctn_cell = ws.cell(
# #             row=rr, column=ctn_pcs_col,
# #             value=f"=SUM({first_size_letter}{rr}:{last_size_letter}{rr})"
# #         )
# #         style_cell(ctn_cell, font=bold, fill=row_fill, border=box, alignment=center)

# #         total_ctn_col = ctn_pcs_col + 1
# #         total_pcs_col = ctn_pcs_col + 2
# #         total_ctn_letter = get_column_letter(total_ctn_col)

# #         tc_cell = ws.cell(row=rr, column=total_ctn_col, value=row_data["total_ctn"])
# #         style_cell(tc_cell, font=normal, fill=row_fill, border=box, alignment=center)

# #         # CARTON NO. = cumulative sum of TOTAL CTN
# #         if idx == 0:
# #             carton_formula = f"={total_ctn_letter}{rr}"
# #         else:
# #             carton_formula = f"={get_column_letter(2)}{rr - 1}+{total_ctn_letter}{rr}"
# #         cB = ws.cell(row=rr, column=2, value=carton_formula)
# #         style_cell(cB, font=normal, fill=row_fill, border=box, alignment=center)

# #         ctn_pcs_letter = get_column_letter(ctn_pcs_col)
# #         tp_cell = ws.cell(
# #             row=rr, column=total_pcs_col,
# #             value=f"={total_ctn_letter}{rr}*{ctn_pcs_letter}{rr}"
# #         )
# #         style_cell(tp_cell, font=normal, fill=row_fill, border=box, alignment=center)

# #         grs_wt_col = total_pcs_col + 1
# #         nw_row = spec_header_row + 1
# #         empty_row = spec_header_row + 3
# #         first_spec_letter = get_column_letter(2)
# #         last_spec_letter = get_column_letter(1 + n_sizes)
# #         first_main_size_letter = get_column_letter(5)
# #         last_main_size_letter = get_column_letter(4 + n_sizes)

# #         grs_formula = (
# #             f"=SUMPRODUCT({first_main_size_letter}{rr}:{last_main_size_letter}{rr},"
# #             f"{first_spec_letter}{nw_row}:{last_spec_letter}{nw_row})"
# #             f"+{first_spec_letter}{empty_row}"
# #         )
# #         grs_cell = ws.cell(row=rr, column=grs_wt_col, value=grs_formula)
# #         style_cell(grs_cell, font=normal, fill=row_fill, border=box, alignment=center)

# #         net_wt_col = grs_wt_col + 1
# #         grs_wt_letter_for_net = get_column_letter(grs_wt_col)
# #         net_cell = ws.cell(
# #             row=rr, column=net_wt_col,
# #             value=f"={grs_wt_letter_for_net}{rr}-{first_spec_letter}{empty_row}"
# #         )
# #         style_cell(net_cell, font=normal, fill=row_fill, border=box, alignment=center)

# #         total_grs_col = net_wt_col + 1
# #         grs_wt_letter = get_column_letter(grs_wt_col)
# #         tg_cell = ws.cell(
# #             row=rr, column=total_grs_col,
# #             value=f"={total_ctn_letter}{rr}*{grs_wt_letter}{rr}"
# #         )
# #         style_cell(tg_cell, font=normal, fill=row_fill, border=box, alignment=center)

# #         total_net_col = total_grs_col + 1
# #         net_wt_letter = get_column_letter(net_wt_col)
# #         tn_cell = ws.cell(
# #             row=rr, column=total_net_col,
# #             value=f"={total_ctn_letter}{rr}*{net_wt_letter}{rr}"
# #         )
# #         style_cell(tn_cell, font=normal, fill=row_fill, border=box, alignment=center)

# #     # ---- TOTAL row ----
# #     total_row = data_row_start + len(carton_rows)
# #     total_label = ws.cell(row=total_row, column=3, value="TOTAL")
# #     style_cell(total_label, font=bold, fill=total_fill, alignment=center)

# #     total_ctn_letter_for_sum = get_column_letter(total_ctn_col)
# #     for sc_idx, size_name in enumerate(sizes):
# #         col = 5 + sc_idx
# #         col_letter = get_column_letter(col)
# #         cell = ws.cell(
# #             row=total_row, column=col,
# #             value=f"=SUMPRODUCT({col_letter}{data_row_start}:{col_letter}{total_row-1},"
# #                   f"{total_ctn_letter_for_sum}{data_row_start}:{total_ctn_letter_for_sum}{total_row-1})"
# #         )
# #         style_cell(cell, font=bold, fill=total_fill, border=box, alignment=center)

# #     total_ctn_col = ctn_pcs_col + 1
# #     total_pcs_col = ctn_pcs_col + 2
# #     total_ctn_letter = get_column_letter(total_ctn_col)
# #     total_pcs_letter = get_column_letter(total_pcs_col)

# #     for col, formula in [
# #         (total_ctn_col,
# #          f'=IF(COUNT({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})=0,"",'
# #          f'SUM({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1}))'),
# #         (total_pcs_col,
# #          f'=IF(COUNT({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1})=0,"",'
# #          f'SUM({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1}))'),
# #     ]:
# #         c = ws.cell(row=total_row, column=col, value=formula)
# #         style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# #     grs_wt_col = total_pcs_col + 1
# #     net_wt_col = total_pcs_col + 2
# #     grs_wt_letter = get_column_letter(grs_wt_col)
# #     net_wt_letter = get_column_letter(net_wt_col)

# #     for col, letter in [(grs_wt_col, grs_wt_letter), (net_wt_col, net_wt_letter)]:
# #         c = ws.cell(row=total_row, column=col,
# #                     value=f"=SUM({letter}{data_row_start}:{letter}{total_row-1})")
# #         style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# #     total_grs_col = total_pcs_col + 3
# #     total_grs_letter = get_column_letter(total_grs_col)
# #     c = ws.cell(row=total_row, column=total_grs_col,
# #                 value=f'=IF(COUNT({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1})=0,"",'
# #                       f'SUM({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1}))')
# #     style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# #     total_net_col = total_grs_col + 1
# #     total_net_letter = get_column_letter(total_net_col)
# #     c = ws.cell(row=total_row, column=total_net_col,
# #                 value=f'=IF(COUNT({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1})=0,"",'
# #                       f'SUM({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1}))')
# #     style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

# #     for col in range(1, n_cols + 1):
# #         cell = ws.cell(row=total_row, column=col)
# #         if cell.fill.fgColor is None or cell.fill.fgColor.rgb == "00000000":
# #             cell.fill = total_fill
# #         if cell.border.left.style is None:
# #             cell.border = box
# #         cell.alignment = center

# #     # ---- Summary block ----
# #     sum_header_row = total_row + 3
# #     c = ws.cell(row=sum_header_row, column=1, value="SIZE")
# #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
# #     for sc_idx, size_name in enumerate(sizes):
# #         c = ws.cell(row=sum_header_row, column=2 + sc_idx, value=size_name)
# #         style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
# #     c = ws.cell(row=sum_header_row, column=2 + n_sizes, value="G TOTAL")
# #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)

# #     gw_col = 3 + n_sizes
# #     nw_col = 4 + n_sizes
# #     ws.merge_cells(start_row=sum_header_row, start_column=gw_col,
# #                     end_row=sum_header_row, end_column=gw_col)
# #     c = ws.cell(row=sum_header_row, column=gw_col, value="Grand Total gross weight (kg)")
# #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
# #     ws.merge_cells(start_row=sum_header_row, start_column=nw_col,
# #                     end_row=sum_header_row, end_column=nw_col)
# #     c = ws.cell(row=sum_header_row, column=nw_col, value="Grand Total net weight (kg)")
# #     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)

# #     rows_needed = ["CUT QTY", "ORDER QTY", "SHIP QTY", "EXS/SHT QTY", "PERCENTAGE"]
# #     for i, label in enumerate(rows_needed):
# #         rr = sum_header_row + 1 + i
# #         lc = ws.cell(row=rr, column=1, value=label)
# #         style_cell(lc, font=bold, border=box, alignment=center)
# #         for sc_idx, size_name in enumerate(sizes):
# #             col = 2 + sc_idx
# #             col_letter = get_column_letter(col)
# #             cell = ws.cell(row=rr, column=col)
# #             if label == "ORDER QTY":
# #                 qty = next(s["qty"] for s in sublines if s["size"] == size_name)
# #                 cell.value = qty
# #             elif label == "CUT QTY":
# #                 cell.fill = yellow
# #             elif label == "SHIP QTY":
# #                 # SHIP QTY = packet-wise total from main table TOTAL row
# #                 main_size_col = 5 + sc_idx
# #                 main_size_letter = get_column_letter(main_size_col)
# #                 cell.value = f"={main_size_letter}{total_row}"
# #             elif label == "EXS/SHT QTY":
# #                 ship_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
# #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# #                 cell.value = f'=IF({col_letter}{ship_row}="","",{col_letter}{ship_row}-{col_letter}{order_row})'
# #             elif label == "PERCENTAGE":
# #                 exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# #                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# #                 cell.value = f'=IF({col_letter}{exs_row}="","",IFERROR({col_letter}{exs_row}/{col_letter}{order_row},0))'
# #                 cell.number_format = "0.00%"
# #             style_cell(cell, font=normal, border=box, alignment=center)
# #             if label == "CUT QTY":
# #                 cell.fill = yellow

# #         gt_col = 2 + n_sizes
# #         gt_letter = get_column_letter(gt_col)
# #         first_letter = get_column_letter(2)
# #         last_letter = get_column_letter(1 + n_sizes)
# #         if label in ("CUT QTY", "EXS/SHT QTY"):
# #             c = ws.cell(row=rr, column=gt_col,
# #                         value=f'=IF(COUNT({first_letter}{rr}:{last_letter}{rr})=0,"",SUM({first_letter}{rr}:{last_letter}{rr}))')
# #             style_cell(c, font=bold, border=box, alignment=center)
# #         elif label in ("ORDER QTY", "SHIP QTY"):
# #             c = ws.cell(row=rr, column=gt_col,
# #                         value=f"=SUM({first_letter}{rr}:{last_letter}{rr})")
# #             style_cell(c, font=bold, border=box, alignment=center)
# #         else:
# #             exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# #             order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# #             c = ws.cell(row=rr, column=gt_col,
# #                         value=f'=IF({gt_letter}{exs_row}="","",IFERROR({gt_letter}{exs_row}/{gt_letter}{order_row},0))')
# #             c.number_format = "0.00%"
# #             style_cell(c, font=bold, border=box, alignment=center)

# #     # Link top header EXCESS/SHORT & PERCENTAGE
# #     exs_gt_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
# #     order_gt_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
# #     gt_col_letter = get_column_letter(2 + n_sizes)

# #     c = ws.cell(row=excess_short_row, column=excess_col,
# #                 value=f'={gt_col_letter}{exs_gt_row}')
# #     style_cell(c, font=normal, border=box, alignment=center)
# #     c = ws.cell(row=percentage_row, column=percentage_col,
# #                 value=f'=IF({gt_col_letter}{exs_gt_row}="","-",'
# #                       f'TEXT(IFERROR({gt_col_letter}{exs_gt_row}/{gt_col_letter}{order_gt_row},0),"0.00%"))')
# #     style_cell(c, font=normal, border=box, alignment=center)

# #     # ---- Weight / CBM block ----
# #     wt_row = sum_header_row + len(rows_needed) + 2
# #     labels = ["GROSS WEIGHT", "NET WEIGHT", "CTN MEAS", "CBM"]
# #     for i, lbl in enumerate(labels):
# #         lc = ws.cell(row=wt_row + i, column=1, value=lbl + " :")
# #         style_cell(lc, font=label_font, alignment=center)
# #         vc = ws.cell(row=wt_row + i, column=2, value=None)
# #         style_cell(vc, fill=yellow, border=box, alignment=center)

# #     first_sum_row = sum_header_row + 1
# #     last_sum_row = sum_header_row + len(rows_needed)
# #     total_grs_letter = get_column_letter(total_grs_col)
# #     total_net_letter = get_column_letter(total_net_col)

# #     ws.merge_cells(start_row=first_sum_row, start_column=gw_col,
# #                     end_row=last_sum_row, end_column=gw_col)
# #     gw_cell = ws.cell(row=first_sum_row, column=gw_col,
# #                        value=f'={total_grs_letter}{total_row}')
# #     style_cell(gw_cell, font=bold, fill=total_fill, border=box, alignment=center)

# #     ws.merge_cells(start_row=first_sum_row, start_column=nw_col,
# #                     end_row=last_sum_row, end_column=nw_col)
# #     nw_cell = ws.cell(row=first_sum_row, column=nw_col,
# #                        value=f'={total_net_letter}{total_row}')
# #     style_cell(nw_cell, font=bold, fill=total_fill, border=box, alignment=center)

# #     # ---- Column widths ----
# #     ws.column_dimensions["A"].width = 18
# #     ws.column_dimensions["B"].width = 14
# #     ws.column_dimensions["C"].width = 16
# #     ws.column_dimensions["D"].width = 16
# #     for i in range(n_sizes):
# #         ws.column_dimensions[get_column_letter(5 + i)].width = 10
# #     for col in range(5 + n_sizes, n_cols + 1):
# #         ws.column_dimensions[get_column_letter(col)].width = 13
# #     ws.column_dimensions[get_column_letter(gw_col)].width = 24
# #     ws.column_dimensions[get_column_letter(nw_col)].width = 24

# #     # ---- Legend ----
# #     legend_row = wt_row + len(labels) + 2
# #     ws.cell(row=legend_row, column=1, value="Legend:").font = bold
# #     ws.cell(row=legend_row, column=1).alignment = center
# #     legend = ws.cell(row=legend_row + 1, column=1,
# #                      value="Yellow cells = data NOT available in the PO PDF (fill in from factory cutting/packing records)")
# #     legend.font = Font(name=FONT, size=9, italic=True, color="666666")
# #     legend.alignment = left_center


# # wb = Workbook()
# # wb.remove(wb.active)

# # used_names = set()
# # for line in data["lines"]:
# #     sublines = [s for s in data["sublines"] if s["style"] == line["style"]]
# #     if not sublines:
# #         continue
# #     name = sheet_name_for(line)
# #     base_name, i = name, 1
# #     while name in used_names:
# #         i += 1
# #         name = f"{base_name[:28]}_{i}"
# #     used_names.add(name)

# #     ws = wb.create_sheet(title=name)
# #     build_sheet(ws, line, sublines)

# # out_path = "packing_list_all_PO_lines.xlsx"
# # wb.save(out_path)
# # print("Saved:", out_path)
# # print("Sheets created:", wb.sheetnames)


# """
# Build a packing-list-style sheet for EVERY PO line found in
# po_extracted.json -- one worksheet per line, in a single workbook.

# HOW TO RUN (after extract_po.py has created po_extracted.json):
#     python3 build_all_packing_lists.py

# Requires:
#     pip install openpyxl
# """
# import json
# import re
# from openpyxl import Workbook
# from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
# from openpyxl.utils import get_column_letter

# with open("po_extracted.json") as f:
#     data = json.load(f)

# FONT = "Arial"
# bold = Font(name=FONT, bold=True, size=10)
# normal = Font(name=FONT, size=10)
# title_font = Font(name=FONT, bold=True, size=14, color="1F4E79")
# subtitle_font = Font(name=FONT, bold=True, size=11, color="2E75B6")
# header_font = Font(name=FONT, bold=True, size=9, color="1F4E79")
# label_font = Font(name=FONT, bold=True, size=10, color="333333")
# thin = Side(style="thin", color="B4C6E7")
# medium = Side(style="medium", color="5B9BD5")
# box = Border(left=thin, right=thin, top=thin, bottom=thin)
# header_box = Border(left=thin, right=thin, top=medium, bottom=medium)
# center = Alignment(horizontal="center", vertical="center", wrap_text=True)
# left_center = Alignment(horizontal="left", vertical="center", wrap_text=True)

# header_fill = PatternFill("solid", fgColor="D6E3F0")
# alt_row_fill = PatternFill("solid", fgColor="F2F7FB")
# spec_fill = PatternFill("solid", fgColor="E2EFDA")
# yellow = PatternFill("solid", fgColor="FFF2CC")
# title_fill = PatternFill("solid", fgColor="D6E3F0")
# total_fill = PatternFill("solid", fgColor="DDEBF7")
# summary_header_fill = PatternFill("solid", fgColor="C6EFCE")


# def sheet_name_for(line):
#     raw = f"{line['style']}"
#     return re.sub(r'[\[\]:*?/\\]', '-', raw)[:31]


# def compute_carton_rows(sublines):
#     if not sublines:
#         return []
#     pack = sublines[0].get("items_per_outer_pack") or 1

#     full_rows = []
#     remainder_pool = []
#     for s in sublines:
#         qty = s["qty"]
#         size = s["size"]
#         full_ctn = qty // pack
#         rem = qty % pack
#         if full_ctn > 0:
#             full_rows.append({
#                 "sizes": {size: pack},
#                 "total_ctn": full_ctn,
#                 "ctn_pcs": pack,
#                 "color": s["color"],
#                 "upc": s["upc"],
#                 "mixed": False,
#             })
#         if rem > 0:
#             remainder_pool.append([size, rem, s["color"], s["upc"]])

#     mixed_rows = []
#     current = {}
#     current_total = 0
#     idx = 0
#     while idx < len(remainder_pool):
#         size, rem, color, upc = remainder_pool[idx]
#         space_left = pack - current_total
#         take = min(rem, space_left)
#         if take > 0:
#             current[size] = current.get(size, 0) + take
#             current_total += take
#             remainder_pool[idx][1] -= take
#         if remainder_pool[idx][1] == 0:
#             idx += 1
#         if current_total == pack:
#             mixed_rows.append({
#                 "sizes": dict(current), "total_ctn": 1, "ctn_pcs": pack,
#                 "color": None, "upc": None, "mixed": True,
#             })
#             current = {}
#             current_total = 0
#     if current:
#         mixed_rows.append({
#             "sizes": dict(current), "total_ctn": 1, "ctn_pcs": current_total,
#             "color": None, "upc": None, "mixed": True,
#         })

#     return full_rows + mixed_rows


# def style_cell(cell, font=None, fill=None, border=None, alignment=None):
#     if font is not None:
#         cell.font = font
#     if fill is not None:
#         cell.fill = fill
#     if border is not None:
#         cell.border = border
#     if alignment is not None:
#         cell.alignment = alignment


# def build_sheet(ws, line, sublines):
#     sizes = [s["size"] for s in sublines]
#     n_sizes = len(sizes)

#     ws.merge_cells("A1:H1")
#     ws["A1"] = "CREATIVE COLLECTIONS LTD. U-1-A"
#     style_cell(ws["A1"], font=title_font, fill=title_fill, alignment=center)
#     ws.row_dimensions[1].height = 22

#     ws.merge_cells("A2:H2")
#     ws["A2"] = "PACKING LIST DETAILS"
#     style_cell(ws["A2"], font=subtitle_font, fill=title_fill, alignment=center)
#     ws.row_dimensions[2].height = 18

#     spec_header_row = 4
#     c = ws.cell(row=spec_header_row, column=1, value="SIZE")
#     style_cell(c, font=header_font, fill=spec_fill, border=box, alignment=center)
#     for sc_idx, size_name in enumerate(sizes):
#         c = ws.cell(row=spec_header_row, column=2 + sc_idx, value=size_name)
#         style_cell(c, font=header_font, fill=spec_fill, border=box, alignment=center)

#     spec_rows = ["N.W.", "N.N.W.", "EMPTY CTN"]
#     for i, label in enumerate(spec_rows):
#         rr = spec_header_row + 1 + i
#         lc = ws.cell(row=rr, column=1, value=label)
#         style_cell(lc, font=bold, fill=spec_fill, border=box, alignment=center)
#         for sc_idx, size_name in enumerate(sizes):
#             c = ws.cell(row=rr, column=2 + sc_idx, value=1)
#             style_cell(c, font=normal, border=box, alignment=center)

#     r = spec_header_row + len(spec_rows) + 2

#     info_labels_1 = ["BUYER", "LOT NO", "P.O NO", "ORDER QTY", "PACK QTY", "EXCESS/SHORT QTY", "PERCENTAGE"]
#     for col_idx, label in enumerate(info_labels_1, start=1):
#         c = ws.cell(row=r, column=col_idx, value=label)
#         style_cell(c, font=header_font, fill=header_fill, border=box, alignment=center)
#     r += 1

#     info_vals_1 = [
#         "VANS", line["style"], line["po_line_no"],
#         f'{line["order_qty"]} PCS', f'{line["order_qty"]} PCS', None, None,
#     ]
#     excess_short_row = r
#     percentage_col = 7
#     excess_col = 6
#     for col_idx, val in enumerate(info_vals_1, start=1):
#         c = ws.cell(row=r, column=col_idx, value=val)
#         style_cell(c, font=normal, border=box, alignment=center)
#     percentage_row = r
#     r += 1

#     info_labels_2 = ["CRD", "COUNTRY", "DESCRIPTION"]
#     for col_idx, label in enumerate(info_labels_2, start=1):
#         c = ws.cell(row=r, column=col_idx, value=label)
#         style_cell(c, font=header_font, fill=header_fill, border=box, alignment=center)
#     r += 1

#     info_vals_2 = [line.get("crd"), line.get("destination_country"), line["description"]]
#     for col_idx, val in enumerate(info_vals_2, start=1):
#         c = ws.cell(row=r, column=col_idx, value=val)
#         style_cell(c, font=normal, border=box, alignment=center)
#         if col_idx == 3:
#             ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
#             c.alignment = center
#     r += 1

#     header_row = r + 2
#     headers = ["CASE LABEL NO.", "CARTON NO.", "COLOR", "UPC Number"] + sizes + \
#               ["CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs.wt.pr ctn", "Net.wt.pr ctn",
#                "total.Grs.wt", "total.net.wt"]
#     for i, h in enumerate(headers, start=1):
#         c = ws.cell(row=header_row, column=i, value=h)
#         style_cell(c, font=header_font, fill=header_fill, border=header_box, alignment=center)
#     n_cols = len(headers)
#     ws.row_dimensions[header_row].height = 30

#     data_row_start = header_row + 1
#     carton_rows = compute_carton_rows(sublines)
#     for idx, row_data in enumerate(carton_rows):
#         rr = data_row_start + idx
#         row_fill = alt_row_fill if idx % 2 == 1 else None

#         c = ws.cell(row=rr, column=1, value=idx + 1)
#         style_cell(c, font=normal, fill=row_fill, border=box, alignment=center)

#         color_val = row_data["color"] if row_data["color"] is not None else "MIXED"
#         color_cell = ws.cell(row=rr, column=3, value=color_val)
#         style_cell(color_cell, font=normal, border=box, alignment=center,
#                    fill=yellow if row_data["mixed"] else row_fill)

#         upc_cell = ws.cell(row=rr, column=4, value=row_data["upc"])
#         style_cell(upc_cell, font=normal, fill=row_fill, border=box, alignment=center)

#         for sc_idx, size_name in enumerate(sizes):
#             col = 5 + sc_idx
#             val = row_data["sizes"].get(size_name)
#             cell = ws.cell(row=rr, column=col, value=val)
#             style_cell(cell, font=normal, fill=row_fill, border=box, alignment=center)

#         ctn_pcs_col = 5 + n_sizes
#         first_size_letter = get_column_letter(5)
#         last_size_letter = get_column_letter(4 + n_sizes)
#         ctn_cell = ws.cell(
#             row=rr, column=ctn_pcs_col,
#             value=f"=SUM({first_size_letter}{rr}:{last_size_letter}{rr})"
#         )
#         style_cell(ctn_cell, font=bold, fill=row_fill, border=box, alignment=center)

#         total_ctn_col = ctn_pcs_col + 1
#         total_pcs_col = ctn_pcs_col + 2
#         total_ctn_letter = get_column_letter(total_ctn_col)

#         tc_cell = ws.cell(row=rr, column=total_ctn_col, value=row_data["total_ctn"])
#         style_cell(tc_cell, font=normal, fill=row_fill, border=box, alignment=center)

#         # CARTON NO. = cumulative TOTAL CTN
#         if idx == 0:
#             carton_formula = f"={total_ctn_letter}{rr}"
#         else:
#             carton_formula = f"={get_column_letter(2)}{rr - 1}+{total_ctn_letter}{rr}"
#         cB = ws.cell(row=rr, column=2, value=carton_formula)
#         style_cell(cB, font=normal, fill=row_fill, border=box, alignment=center)

#         ctn_pcs_letter = get_column_letter(ctn_pcs_col)
#         tp_cell = ws.cell(
#             row=rr, column=total_pcs_col,
#             value=f"={total_ctn_letter}{rr}*{ctn_pcs_letter}{rr}"
#         )
#         style_cell(tp_cell, font=normal, fill=row_fill, border=box, alignment=center)

#         grs_wt_col = total_pcs_col + 1
#         nw_row = spec_header_row + 1
#         empty_row = spec_header_row + 3
#         first_spec_letter = get_column_letter(2)
#         last_spec_letter = get_column_letter(1 + n_sizes)
#         first_main_size_letter = get_column_letter(5)
#         last_main_size_letter = get_column_letter(4 + n_sizes)

#         grs_formula = (
#             f"=SUMPRODUCT({first_main_size_letter}{rr}:{last_main_size_letter}{rr},"
#             f"{first_spec_letter}{nw_row}:{last_spec_letter}{nw_row})"
#             f"+{first_spec_letter}{empty_row}"
#         )
#         grs_cell = ws.cell(row=rr, column=grs_wt_col, value=grs_formula)
#         style_cell(grs_cell, font=normal, fill=row_fill, border=box, alignment=center)

#         net_wt_col = grs_wt_col + 1
#         grs_wt_letter_for_net = get_column_letter(grs_wt_col)
#         net_cell = ws.cell(
#             row=rr, column=net_wt_col,
#             value=f"={grs_wt_letter_for_net}{rr}-{first_spec_letter}{empty_row}"
#         )
#         style_cell(net_cell, font=normal, fill=row_fill, border=box, alignment=center)

#         total_grs_col = net_wt_col + 1
#         grs_wt_letter = get_column_letter(grs_wt_col)
#         tg_cell = ws.cell(
#             row=rr, column=total_grs_col,
#             value=f"={total_ctn_letter}{rr}*{grs_wt_letter}{rr}"
#         )
#         style_cell(tg_cell, font=normal, fill=row_fill, border=box, alignment=center)

#         total_net_col = total_grs_col + 1
#         net_wt_letter = get_column_letter(net_wt_col)
#         tn_cell = ws.cell(
#             row=rr, column=total_net_col,
#             value=f"={total_ctn_letter}{rr}*{net_wt_letter}{rr}"
#         )
#         style_cell(tn_cell, font=normal, fill=row_fill, border=box, alignment=center)

#     total_row = data_row_start + len(carton_rows)
#     total_label = ws.cell(row=total_row, column=3, value="TOTAL")
#     style_cell(total_label, font=bold, fill=total_fill, alignment=center)

#     total_ctn_letter_for_sum = get_column_letter(total_ctn_col)
#     for sc_idx, size_name in enumerate(sizes):
#         col = 5 + sc_idx
#         col_letter = get_column_letter(col)
#         cell = ws.cell(
#             row=total_row, column=col,
#             value=f"=SUMPRODUCT({col_letter}{data_row_start}:{col_letter}{total_row-1},"
#                   f"{total_ctn_letter_for_sum}{data_row_start}:{total_ctn_letter_for_sum}{total_row-1})"
#         )
#         style_cell(cell, font=bold, fill=total_fill, border=box, alignment=center)

#     total_ctn_col = ctn_pcs_col + 1
#     total_pcs_col = ctn_pcs_col + 2
#     total_ctn_letter = get_column_letter(total_ctn_col)
#     total_pcs_letter = get_column_letter(total_pcs_col)

#     for col, formula in [
#         (total_ctn_col,
#          f'=IF(COUNT({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})=0,"",'
#          f'SUM({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1}))'),
#         (total_pcs_col,
#          f'=IF(COUNT({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1})=0,"",'
#          f'SUM({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1}))'),
#     ]:
#         c = ws.cell(row=total_row, column=col, value=formula)
#         style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

#     grs_wt_col = total_pcs_col + 1
#     net_wt_col = total_pcs_col + 2
#     grs_wt_letter = get_column_letter(grs_wt_col)
#     net_wt_letter = get_column_letter(net_wt_col)

#     for col, letter in [(grs_wt_col, grs_wt_letter), (net_wt_col, net_wt_letter)]:
#         c = ws.cell(row=total_row, column=col,
#                     value=f"=SUM({letter}{data_row_start}:{letter}{total_row-1})")
#         style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

#     total_grs_col = total_pcs_col + 3
#     total_grs_letter = get_column_letter(total_grs_col)
#     c = ws.cell(row=total_row, column=total_grs_col,
#                 value=f'=IF(COUNT({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1})=0,"",'
#                       f'SUM({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1}))')
#     style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

#     total_net_col = total_grs_col + 1
#     total_net_letter = get_column_letter(total_net_col)
#     c = ws.cell(row=total_row, column=total_net_col,
#                 value=f'=IF(COUNT({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1})=0,"",'
#                       f'SUM({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1}))')
#     style_cell(c, font=bold, fill=total_fill, border=box, alignment=center)

#     for col in range(1, n_cols + 1):
#         cell = ws.cell(row=total_row, column=col)
#         if cell.fill.fgColor is None or cell.fill.fgColor.rgb == "00000000":
#             cell.fill = total_fill
#         if cell.border.left.style is None:
#             cell.border = box
#         cell.alignment = center

#     sum_header_row = total_row + 3
#     c = ws.cell(row=sum_header_row, column=1, value="SIZE")
#     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
#     for sc_idx, size_name in enumerate(sizes):
#         c = ws.cell(row=sum_header_row, column=2 + sc_idx, value=size_name)
#         style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
#     c = ws.cell(row=sum_header_row, column=2 + n_sizes, value="G TOTAL")
#     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)

#     gw_col = 3 + n_sizes
#     nw_col = 4 + n_sizes
#     ws.merge_cells(start_row=sum_header_row, start_column=gw_col,
#                     end_row=sum_header_row, end_column=gw_col)
#     c = ws.cell(row=sum_header_row, column=gw_col, value="Grand Total gross weight (kg)")
#     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)
#     ws.merge_cells(start_row=sum_header_row, start_column=nw_col,
#                     end_row=sum_header_row, end_column=nw_col)
#     c = ws.cell(row=sum_header_row, column=nw_col, value="Grand Total net weight (kg)")
#     style_cell(c, font=header_font, fill=summary_header_fill, border=box, alignment=center)

#     rows_needed = ["CUT QTY", "ORDER QTY", "SHIP QTY", "EXS/SHT QTY", "PERCENTAGE"]
#     for i, label in enumerate(rows_needed):
#         rr = sum_header_row + 1 + i
#         lc = ws.cell(row=rr, column=1, value=label)
#         style_cell(lc, font=bold, border=box, alignment=center)
#         for sc_idx, size_name in enumerate(sizes):
#             col = 2 + sc_idx
#             col_letter = get_column_letter(col)
#             cell = ws.cell(row=rr, column=col)
#             if label == "ORDER QTY":
#                 qty = next(s["qty"] for s in sublines if s["size"] == size_name)
#                 cell.value = qty
#             elif label == "CUT QTY":
#                 cell.fill = yellow
#             elif label == "SHIP QTY":
#                 main_size_col = 5 + sc_idx
#                 main_size_letter = get_column_letter(main_size_col)
#                 cell.value = f"={main_size_letter}{total_row}"
#             elif label == "EXS/SHT QTY":
#                 ship_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
#                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
#                 cell.value = f'=IF({col_letter}{ship_row}="","",{col_letter}{ship_row}-{col_letter}{order_row})'
#             elif label == "PERCENTAGE":
#                 exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
#                 order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
#                 cell.value = f'=IF({col_letter}{exs_row}="","",IFERROR({col_letter}{exs_row}/{col_letter}{order_row},0))'
#                 cell.number_format = "0.00%"
#             style_cell(cell, font=normal, border=box, alignment=center)
#             if label == "CUT QTY":
#                 cell.fill = yellow

#         gt_col = 2 + n_sizes
#         gt_letter = get_column_letter(gt_col)
#         first_letter = get_column_letter(2)
#         last_letter = get_column_letter(1 + n_sizes)
#         if label in ("CUT QTY", "EXS/SHT QTY"):
#             c = ws.cell(row=rr, column=gt_col,
#                         value=f'=IF(COUNT({first_letter}{rr}:{last_letter}{rr})=0,"",SUM({first_letter}{rr}:{last_letter}{rr}))')
#             style_cell(c, font=bold, border=box, alignment=center)
#         elif label in ("ORDER QTY", "SHIP QTY"):
#             c = ws.cell(row=rr, column=gt_col,
#                         value=f"=SUM({first_letter}{rr}:{last_letter}{rr})")
#             style_cell(c, font=bold, border=box, alignment=center)
#         else:
#             exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
#             order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
#             c = ws.cell(row=rr, column=gt_col,
#                         value=f'=IF({gt_letter}{exs_row}="","",IFERROR({gt_letter}{exs_row}/{gt_letter}{order_row},0))')
#             c.number_format = "0.00%"
#             style_cell(c, font=bold, border=box, alignment=center)

#     exs_gt_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
#     order_gt_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
#     gt_col_letter = get_column_letter(2 + n_sizes)

#     c = ws.cell(row=excess_short_row, column=excess_col,
#                 value=f'={gt_col_letter}{exs_gt_row}')
#     style_cell(c, font=normal, border=box, alignment=center)
#     c = ws.cell(row=percentage_row, column=percentage_col,
#                 value=f'=IF({gt_col_letter}{exs_gt_row}="","-",'
#                       f'TEXT(IFERROR({gt_col_letter}{exs_gt_row}/{gt_col_letter}{order_gt_row},0),"0.00%"))')
#     style_cell(c, font=normal, border=box, alignment=center)

#     # ---- Weight / CBM block ----
#     wt_row = sum_header_row + len(rows_needed) + 2
#     total_grs_letter = get_column_letter(total_grs_col)
#     total_net_letter = get_column_letter(total_net_col)
#     total_ctn_letter_final = get_column_letter(total_ctn_col)

#     # GROSS WEIGHT = Grand Total gross weight
#     lc = ws.cell(row=wt_row, column=1, value="GROSS WEIGHT :")
#     style_cell(lc, font=label_font, alignment=center)
#     vc = ws.cell(row=wt_row, column=2, value=f"={total_grs_letter}{total_row}")
#     style_cell(vc, font=bold, border=box, alignment=center)

#     # NET WEIGHT = Grand Total net weight
#     lc = ws.cell(row=wt_row + 1, column=1, value="NET WEIGHT :")
#     style_cell(lc, font=label_font, alignment=center)
#     vc = ws.cell(row=wt_row + 1, column=2, value=f"={total_net_letter}{total_row}")
#     style_cell(vc, font=bold, border=box, alignment=center)

#     # CTN MEAS : user input 1 + user input 2
#     lc = ws.cell(row=wt_row + 2, column=1, value="CTN MEAS :")
#     style_cell(lc, font=label_font, alignment=center)
#     vc1 = ws.cell(row=wt_row + 2, column=2, value=None)
#     style_cell(vc1, fill=yellow, border=box, alignment=center)
#     vc2 = ws.cell(row=wt_row + 3, column=2, value=None)
#     style_cell(vc2, fill=yellow, border=box, alignment=center)

#     # CBM : (yellow numeric input) × TOTAL CTN
#     lc = ws.cell(row=wt_row + 4, column=1, value="CBM :")
#     style_cell(lc, font=label_font, alignment=center)
#     cbm_input = ws.cell(row=wt_row + 4, column=2, value=None)
#     style_cell(cbm_input, fill=yellow, border=box, alignment=center)
#     cbm_result = ws.cell(
#         row=wt_row + 4, column=3,
#         value=f'=IF(B{wt_row + 4}="","",B{wt_row + 4}*{total_ctn_letter_final}{total_row})'
#     )
#     style_cell(cbm_result, font=bold, border=box, alignment=center)

#     first_sum_row = sum_header_row + 1
#     last_sum_row = sum_header_row + len(rows_needed)

#     ws.merge_cells(start_row=first_sum_row, start_column=gw_col,
#                     end_row=last_sum_row, end_column=gw_col)
#     gw_cell = ws.cell(row=first_sum_row, column=gw_col,
#                        value=f'={total_grs_letter}{total_row}')
#     style_cell(gw_cell, font=bold, fill=total_fill, border=box, alignment=center)

#     ws.merge_cells(start_row=first_sum_row, start_column=nw_col,
#                     end_row=last_sum_row, end_column=nw_col)
#     nw_cell = ws.cell(row=first_sum_row, column=nw_col,
#                        value=f'={total_net_letter}{total_row}')
#     style_cell(nw_cell, font=bold, fill=total_fill, border=box, alignment=center)

#     ws.column_dimensions["A"].width = 18
#     ws.column_dimensions["B"].width = 14
#     ws.column_dimensions["C"].width = 16
#     ws.column_dimensions["D"].width = 16
#     for i in range(n_sizes):
#         ws.column_dimensions[get_column_letter(5 + i)].width = 10
#     for col in range(5 + n_sizes, n_cols + 1):
#         ws.column_dimensions[get_column_letter(col)].width = 13
#     ws.column_dimensions[get_column_letter(gw_col)].width = 24
#     ws.column_dimensions[get_column_letter(nw_col)].width = 24

#     legend_row = wt_row + 6
#     ws.cell(row=legend_row, column=1, value="Legend:").font = bold
#     ws.cell(row=legend_row, column=1).alignment = center
#     legend = ws.cell(row=legend_row + 1, column=1,
#                      value="Yellow cells = data NOT available in the PO PDF (fill in from factory cutting/packing records)")
#     legend.font = Font(name=FONT, size=9, italic=True, color="666666")
#     legend.alignment = left_center


# wb = Workbook()
# wb.remove(wb.active)

# used_names = set()
# for line in data["lines"]:
#     sublines = [s for s in data["sublines"] if s["style"] == line["style"]]
#     if not sublines:
#         continue
#     name = sheet_name_for(line)
#     base_name, i = name, 1
#     while name in used_names:
#         i += 1
#         name = f"{base_name[:28]}_{i}"
#     used_names.add(name)

#     ws = wb.create_sheet(title=name)
#     build_sheet(ws, line, sublines)

# out_path = "packing_list_all_PO_lines.xlsx"
# wb.save(out_path)
# print("Saved:", out_path)
# print("Sheets created:", wb.sheetnames)



"""
Build a packing-list-style sheet for EVERY PO line found in
po_extracted.json -- one worksheet per line, in a single workbook.

HOW TO RUN (after extract_po.py has created po_extracted.json):
    python3 build_all_packing_lists.py

Requires:
    pip install openpyxl
"""
import json
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

with open("po_extracted.json") as f:
    data = json.load(f)

FONT = "Arial"
bold = Font(name=FONT, bold=True, size=10)
normal = Font(name=FONT, size=10)
title_font = Font(name=FONT, bold=True, size=13)
thin = Side(style="thin")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
yellow = PatternFill("solid", fgColor="FFFF00")
grey = PatternFill("solid", fgColor="D9D9D9")


def sheet_name_for(line):
    raw = f"{line['style']}"
    return re.sub(r'[\[\]:*?/\\]', '-', raw)[:31]


def compute_carton_rows(sublines):
    """
    Splits each subline's qty into full-size cartons (qty // pack) plus a
    shared remainder pool that gets packed into mixed cartons, filled in
    size order until each hits pack capacity; the final leftover carton
    may be under capacity.
    Returns a list of dicts: {"sizes": {size: qty, ...}, "total_ctn": int,
    "ctn_pcs": int, "color": str or None, "upc": str or None, "mixed": bool}
    """
    if not sublines:
        return []
    pack = sublines[0].get("items_per_outer_pack") or 1

    full_rows = []
    remainder_pool = []  # [size, qty_left, color, upc]
    for s in sublines:
        qty = s["qty"]
        size = s["size"]
        full_ctn = qty // pack
        rem = qty % pack
        if full_ctn > 0:
            full_rows.append({
                "sizes": {size: pack},
                "total_ctn": full_ctn,
                "ctn_pcs": pack,
                "color": s["color"],
                "upc": s["upc"],
                "mixed": False,
            })
        if rem > 0:
            remainder_pool.append([size, rem, s["color"], s["upc"]])

    mixed_rows = []
    current = {}
    current_total = 0
    idx = 0
    while idx < len(remainder_pool):
        size, rem, color, upc = remainder_pool[idx]
        space_left = pack - current_total
        take = min(rem, space_left)
        if take > 0:
            current[size] = current.get(size, 0) + take
            current_total += take
            remainder_pool[idx][1] -= take
        if remainder_pool[idx][1] == 0:
            idx += 1
        if current_total == pack:
            mixed_rows.append({
                "sizes": dict(current), "total_ctn": 1, "ctn_pcs": pack,
                "color": None, "upc": None, "mixed": True,
            })
            current = {}
            current_total = 0
    if current:
        mixed_rows.append({
            "sizes": dict(current), "total_ctn": 1, "ctn_pcs": current_total,
            "color": None, "upc": None, "mixed": True,
        })

    return full_rows + mixed_rows


def build_sheet(ws, line, sublines):
    sizes = [s["size"] for s in sublines]
    n_sizes = len(sizes)

    ws.merge_cells("A1:H1")
    ws["A1"] = "CREATIVE COLLECTIONS LTD. U-1-A"
    ws["A1"].font = title_font
    ws["A1"].alignment = center

    ws.merge_cells("A2:H2")
    ws["A2"] = "PACKING LIST DETAILS"
    ws["A2"].font = Font(name=FONT, bold=True, size=10)
    ws["A2"].alignment = center

    # ---- Small size-spec table: SIZE / N.W. / N.N.W. / EMPTY CTN ----
    spec_header_row = 4
    ws.cell(row=spec_header_row, column=1, value="SIZE").font = bold
    ws.cell(row=spec_header_row, column=1).border = box
    for sc_idx, size_name in enumerate(sizes):
        c = ws.cell(row=spec_header_row, column=2 + sc_idx, value=size_name)
        c.font = bold
        c.alignment = center
        c.border = box
        c.fill = grey

    spec_rows = ["N.W.", "N.N.W.", "EMPTY CTN"]
    for i, label in enumerate(spec_rows):
        rr = spec_header_row + 1 + i
        lc = ws.cell(row=rr, column=1, value=label)
        lc.font = bold
        lc.border = box
        for sc_idx, size_name in enumerate(sizes):
            c = ws.cell(row=rr, column=2 + sc_idx, value=1)
            c.alignment = center
            c.border = box

    r = spec_header_row + len(spec_rows) + 2
    ws.cell(row=r, column=1, value="BUYER").font = bold
    ws.cell(row=r, column=2, value="VANS").font = normal
    r += 1
    ws.cell(row=r, column=1, value="LOT NO").font = bold
    ws.cell(row=r, column=2, value=line["style"]).font = normal
    r += 1
    ws.cell(row=r, column=1, value="P.O NO").font = bold
    ws.cell(row=r, column=2, value=line["po_line_no"]).font = normal
    r += 1
    ws.cell(row=r, column=1, value="ORDER QTY").font = bold
    ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
    ws.cell(row=r, column=3, value="PCS").font = normal
    r += 1
    ws.cell(row=r, column=1, value="PACK QTY").font = bold
    ws.cell(row=r, column=2, value=line["order_qty"]).font = normal
    ws.cell(row=r, column=3, value="PCS").font = normal
    r += 1
    excess_short_row = r  # filled in with a formula once the summary block below is built
    ws.cell(row=r, column=1, value="EXCESS/SHORT QTY").font = bold
    ws.cell(row=r, column=3, value="PCS").font = normal
    r += 1
    percentage_row = r  # filled in with a formula once the summary block below is built
    ws.cell(row=r, column=1, value="PERCENTAGE").font = bold
    r += 1
    ws.cell(row=r, column=1, value="CRD").font = bold
    ws.cell(row=r, column=2, value=line.get("crd")).font = normal
    r += 1
    ws.cell(row=r, column=1, value="COUNTRY").font = bold
    ws.cell(row=r, column=2, value=line.get("destination_country")).font = normal
    r += 1
    ws.cell(row=r, column=1, value="DESCRIPTION").font = bold
    ws.cell(row=r, column=2, value=line["description"]).font = normal

    header_row = r + 2
    headers = ["CASE LABEL NO.", "CARTON NO.", "COLOR", "UPC Number"] + sizes + \
              ["CTN PCS", "TOTAL CTN", "TOTAL PCS", "Grs.wt.pr ctn", "Net.wt.pr ctn",
               "total.Grs.wt", "total.net.wt"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = bold
        c.alignment = center
        c.border = box
        c.fill = grey
    n_cols = len(headers)

    data_row_start = header_row + 1
    carton_rows = compute_carton_rows(sublines)
    for idx, row_data in enumerate(carton_rows):
        rr = data_row_start + idx
        ws.cell(row=rr, column=1, value=idx + 1).border = box
        cB = ws.cell(row=rr, column=2, value=None)
        cB.fill = yellow
        cB.border = box
        color_val = row_data["color"] if row_data["color"] is not None else "MIXED"
        color_cell = ws.cell(row=rr, column=3, value=color_val)
        color_cell.border = box
        if row_data["mixed"]:
            color_cell.fill = yellow
        ws.cell(row=rr, column=4, value=row_data["upc"]).border = box
        for sc_idx, size_name in enumerate(sizes):
            col = 5 + sc_idx
            val = row_data["sizes"].get(size_name)
            cell = ws.cell(row=rr, column=col, value=val)
            cell.border = box
            cell.alignment = center

        # CTN PCS = sum of all size columns in THIS row (handles mixed cartons)
        ctn_pcs_col = 5 + n_sizes
        first_size_letter = get_column_letter(5)
        last_size_letter = get_column_letter(4 + n_sizes)
        ctn_cell = ws.cell(
            row=rr, column=ctn_pcs_col,
            value=f"=SUM({first_size_letter}{rr}:{last_size_letter}{rr})"
        )
        ctn_cell.border = box
        ctn_cell.alignment = center
        ctn_cell.font = bold

        # Remaining columns after CTN PCS: TOTAL CTN, TOTAL PCS, then weights
        total_ctn_col = ctn_pcs_col + 1
        total_pcs_col = ctn_pcs_col + 2
        total_ctn_letter = get_column_letter(total_ctn_col)

        # TOTAL CTN = calculated value (qty // pack, or 1 per mixed carton)
        tc_cell = ws.cell(row=rr, column=total_ctn_col, value=row_data["total_ctn"])
        tc_cell.border = box
        tc_cell.alignment = center

        # TOTAL PCS = TOTAL CTN * CTN PCS
        ctn_pcs_letter = get_column_letter(ctn_pcs_col)
        tp_cell = ws.cell(
            row=rr, column=total_pcs_col,
            value=f"={total_ctn_letter}{rr}*{ctn_pcs_letter}{rr}"
        )
        tp_cell.border = box
        tp_cell.alignment = center

        # Grs.wt.pr ctn = SUMPRODUCT(this row's size qtys, N.W. row) + EMPTY CTN
        # (EMPTY CTN taken from the first size column, since it's normally
        #  the same physical carton regardless of size)
        grs_wt_col = total_pcs_col + 1
        nw_row = spec_header_row + 1       # N.W. row in the spec table
        empty_row = spec_header_row + 3    # EMPTY CTN row in the spec table
        first_spec_letter = get_column_letter(2)
        last_spec_letter = get_column_letter(1 + n_sizes)
        first_main_size_letter = get_column_letter(5)
        last_main_size_letter = get_column_letter(4 + n_sizes)

        grs_formula = (
            f"=SUMPRODUCT({first_main_size_letter}{rr}:{last_main_size_letter}{rr},"
            f"{first_spec_letter}{nw_row}:{last_spec_letter}{nw_row})"
            f"+{first_spec_letter}{empty_row}"
        )
        grs_cell = ws.cell(row=rr, column=grs_wt_col, value=grs_formula)
        grs_cell.border = box
        grs_cell.alignment = center

        # Net.wt.pr ctn = Grs.wt.pr ctn - EMPTY CTN (pulled from spec table,
        # same single value used above - not a visible column here)
        net_wt_col = grs_wt_col + 1
        grs_wt_letter_for_net = get_column_letter(grs_wt_col)
        net_cell = ws.cell(
            row=rr, column=net_wt_col,
            value=f"={grs_wt_letter_for_net}{rr}-{first_spec_letter}{empty_row}"
        )
        net_cell.border = box
        net_cell.alignment = center

        # total.Grs.wt = TOTAL CTN * Grs.wt.pr ctn
        total_grs_col = net_wt_col + 1
        grs_wt_letter = get_column_letter(grs_wt_col)
        tg_cell = ws.cell(
            row=rr, column=total_grs_col,
            value=f"={total_ctn_letter}{rr}*{grs_wt_letter}{rr}"
        )
        tg_cell.border = box
        tg_cell.alignment = center

        # total.net.wt = TOTAL CTN * Net.wt.pr ctn
        total_net_col = total_grs_col + 1
        net_wt_letter = get_column_letter(net_wt_col)
        tn_cell = ws.cell(
            row=rr, column=total_net_col,
            value=f"={total_ctn_letter}{rr}*{net_wt_letter}{rr}"
        )
        tn_cell.border = box
        tn_cell.alignment = center

    total_row = data_row_start + len(carton_rows)
    ws.cell(row=total_row, column=3, value="TOTAL").font = bold
    total_ctn_letter_for_sum = get_column_letter(total_ctn_col)
    for sc_idx, size_name in enumerate(sizes):
        col = 5 + sc_idx
        col_letter = get_column_letter(col)
        # SUMPRODUCT (not SUM): each row's size qty is a PER-CARTON amount,
        # so it must be multiplied by that row's TOTAL CTN before summing
        # (a row with TOTAL CTN=2 represents 2 identical cartons).
        cell = ws.cell(
            row=total_row, column=col,
            value=f"=SUMPRODUCT({col_letter}{data_row_start}:{col_letter}{total_row-1},"
                  f"{total_ctn_letter_for_sum}{data_row_start}:{total_ctn_letter_for_sum}{total_row-1})"
        )
        cell.font = bold
        cell.border = box
    # Grand-total TOTAL CTN = sum of each row's TOTAL CTN
    # Grand-total TOTAL PCS = sum of each row's TOTAL PCS (not the size columns)
    total_ctn_col = ctn_pcs_col + 1
    total_pcs_col = ctn_pcs_col + 2
    total_ctn_letter = get_column_letter(total_ctn_col)
    total_pcs_letter = get_column_letter(total_pcs_col)

    ws.cell(row=total_row, column=total_ctn_col,
            value=f'=IF(COUNT({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})=0,"",'
                  f'SUM({total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1}))').font = bold

    ws.cell(row=total_row, column=total_pcs_col,
            value=f'=IF(COUNT({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1})=0,"",'
                  f'SUM({total_pcs_letter}{data_row_start}:{total_pcs_letter}{total_row-1}))').font = bold

    # Grand-total Grs.wt.pr ctn / Net.wt.pr ctn = sum of each row's value
    # (matches the reference image, which sums these two columns directly)
    grs_wt_col = total_pcs_col + 1
    net_wt_col = total_pcs_col + 2
    grs_wt_letter = get_column_letter(grs_wt_col)
    net_wt_letter = get_column_letter(net_wt_col)

    # SUMPRODUCT (not SUM): these are PER-CARTON weights, so a row with
    # TOTAL CTN=2 must count double toward the grand total.
    ws.cell(row=total_row, column=grs_wt_col,
            value=f"=SUMPRODUCT({grs_wt_letter}{data_row_start}:{grs_wt_letter}{total_row-1},"
                  f"{total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})").font = bold
    ws.cell(row=total_row, column=net_wt_col,
            value=f"=SUMPRODUCT({net_wt_letter}{data_row_start}:{net_wt_letter}{total_row-1},"
                  f"{total_ctn_letter}{data_row_start}:{total_ctn_letter}{total_row-1})").font = bold

    # Grand-total total.Grs.wt = sum of each row's total.Grs.wt
    total_grs_col = total_pcs_col + 3  # skip Grs.wt.pr ctn, Net.wt.pr ctn -> land on total.Grs.wt
    total_grs_letter = get_column_letter(total_grs_col)
    ws.cell(row=total_row, column=total_grs_col,
            value=f'=IF(COUNT({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1})=0,"",'
                  f'SUM({total_grs_letter}{data_row_start}:{total_grs_letter}{total_row-1}))').font = bold

    # Grand-total total.net.wt = sum of each row's total.net.wt
    total_net_col = total_grs_col + 1
    total_net_letter = get_column_letter(total_net_col)
    ws.cell(row=total_row, column=total_net_col,
            value=f'=IF(COUNT({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1})=0,"",'
                  f'SUM({total_net_letter}{data_row_start}:{total_net_letter}{total_row-1}))').font = bold

    sum_header_row = total_row + 3
    ws.cell(row=sum_header_row, column=1, value="SIZE").font = bold
    for sc_idx, size_name in enumerate(sizes):
        ws.cell(row=sum_header_row, column=2 + sc_idx, value=size_name).font = bold
    ws.cell(row=sum_header_row, column=2 + n_sizes, value="G TOTAL").font = bold

    # Two extra header columns: Grand Total gross/net weight (kg)
    gw_col = 3 + n_sizes
    nw_col = 4 + n_sizes
    ws.merge_cells(start_row=sum_header_row, start_column=gw_col,
                    end_row=sum_header_row, end_column=gw_col)
    ws.cell(row=sum_header_row, column=gw_col, value="Grand Total gross weight (kg)").font = bold
    ws.cell(row=sum_header_row, column=gw_col).alignment = center
    ws.merge_cells(start_row=sum_header_row, start_column=nw_col,
                    end_row=sum_header_row, end_column=nw_col)
    ws.cell(row=sum_header_row, column=nw_col, value="Grand Total net weight (kg)").font = bold
    ws.cell(row=sum_header_row, column=nw_col).alignment = center

    rows_needed = ["CUT QTY", "ORDER QTY", "SHIP QTY", "EXS/SHT QTY", "PERCENTAGE"]
    for i, label in enumerate(rows_needed):
        rr = sum_header_row + 1 + i
        ws.cell(row=rr, column=1, value=label).font = bold
        for sc_idx, size_name in enumerate(sizes):
            col = 2 + sc_idx
            col_letter = get_column_letter(col)
            cell = ws.cell(row=rr, column=col)
            if label == "ORDER QTY":
                qty = next(s["qty"] for s in sublines if s["size"] == size_name)
                cell.value = qty
            elif label in ("CUT QTY", "SHIP QTY"):
                cell.fill = yellow
            elif label == "EXS/SHT QTY":
                ship_row = sum_header_row + 1 + rows_needed.index("SHIP QTY")
                order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
                cell.value = f'=IF({col_letter}{ship_row}="","",{col_letter}{ship_row}-{col_letter}{order_row})'
            elif label == "PERCENTAGE":
                exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
                order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
                cell.value = f'=IF({col_letter}{exs_row}="","",IFERROR({col_letter}{exs_row}/{col_letter}{order_row},0))'
                cell.number_format = "0.00%"
            cell.border = box
            cell.alignment = center

        gt_col = 2 + n_sizes
        gt_letter = get_column_letter(gt_col)
        first_letter = get_column_letter(2)
        last_letter = get_column_letter(1 + n_sizes)
        if label in ("CUT QTY", "SHIP QTY", "EXS/SHT QTY"):
            ws.cell(row=rr, column=gt_col,
                    value=f'=IF(COUNT({first_letter}{rr}:{last_letter}{rr})=0,"",SUM({first_letter}{rr}:{last_letter}{rr}))').font = bold
        elif label == "ORDER QTY":
            ws.cell(row=rr, column=gt_col,
                    value=f"=SUM({first_letter}{rr}:{last_letter}{rr})").font = bold
        else:  # PERCENTAGE
            exs_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
            order_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
            ws.cell(row=rr, column=gt_col,
                    value=f'=IF({gt_letter}{exs_row}="","",IFERROR({gt_letter}{exs_row}/{gt_letter}{order_row},0))').number_format = "0.00%"
        ws.cell(row=rr, column=gt_col).border = box

    # Link the top header's EXCESS/SHORT QTY and PERCENTAGE to the
    # summary block's G-Total values (calculated above) - same numbers,
    # no separate manual entry needed.
    exs_gt_row = sum_header_row + 1 + rows_needed.index("EXS/SHT QTY")
    order_gt_row = sum_header_row + 1 + rows_needed.index("ORDER QTY")
    gt_col_letter = get_column_letter(2 + n_sizes)

    ws.cell(row=excess_short_row, column=2,
            value=f'={gt_col_letter}{exs_gt_row}').font = normal
    ws.cell(row=percentage_row, column=2,
            value=f'=IF({gt_col_letter}{exs_gt_row}="","-",'
                  f'TEXT(IFERROR({gt_col_letter}{exs_gt_row}/{gt_col_letter}{order_gt_row},0),"0.00%"))'
            ).font = normal

    wt_row = sum_header_row + len(rows_needed) + 2
    labels = ["GROSS WEIGHT", "NET WEIGHT", "CTN MEAS", "CBM"]
    for i, lbl in enumerate(labels):
        ws.cell(row=wt_row + i, column=1, value=lbl + " :").font = bold
        ws.cell(row=wt_row + i, column=2, value=None).fill = yellow

    # Grand Total gross/net weight (kg) - merged, pulled from the main
    # table's total.Grs.wt / total.net.wt grand totals (computed above)
    first_sum_row = sum_header_row + 1
    last_sum_row = sum_header_row + len(rows_needed)
    total_grs_letter = get_column_letter(total_grs_col)
    total_net_letter = get_column_letter(total_net_col)

    ws.merge_cells(start_row=first_sum_row, start_column=gw_col,
                    end_row=last_sum_row, end_column=gw_col)
    gw_cell = ws.cell(row=first_sum_row, column=gw_col,
                       value=f'={total_grs_letter}{total_row}')
    gw_cell.alignment = center
    gw_cell.font = bold
    gw_cell.border = box

    ws.merge_cells(start_row=first_sum_row, start_column=nw_col,
                    end_row=last_sum_row, end_column=nw_col)
    nw_cell = ws.cell(row=first_sum_row, column=nw_col,
                       value=f'={total_net_letter}{total_row}')
    nw_cell.alignment = center
    nw_cell.font = bold
    nw_cell.border = box

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    for i in range(n_sizes):
        ws.column_dimensions[get_column_letter(5 + i)].width = 10
    for col in range(5 + n_sizes, n_cols + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions[get_column_letter(gw_col)].width = 22
    ws.column_dimensions[get_column_letter(nw_col)].width = 22

    legend_row = wt_row + len(labels) + 2
    ws.cell(row=legend_row, column=1, value="Legend:").font = bold
    ws.cell(row=legend_row + 1, column=1,
            value="Yellow cells = data NOT available in the PO PDF (fill in from factory cutting/packing records)"
            ).font = Font(name=FONT, size=9, italic=True)


wb = Workbook()
wb.remove(wb.active)

used_names = set()
for line in data["lines"]:
    sublines = [s for s in data["sublines"] if s["style"] == line["style"]]
    if not sublines:
        continue
    name = sheet_name_for(line)
    base_name, i = name, 1
    while name in used_names:
        i += 1
        name = f"{base_name[:28]}_{i}"
    used_names.add(name)

    ws = wb.create_sheet(title=name)
    build_sheet(ws, line, sublines)

out_path = "packing_list_all_PO_lines.xlsx"
wb.save(out_path)
print("Saved:", out_path)
print("Sheets created:", wb.sheetnames)