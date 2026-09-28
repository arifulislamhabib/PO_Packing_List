# # import streamlit as st
# # import uuid

# # # ============================================================
# # # PAGE CONFIG
# # # ============================================================

# # st.set_page_config(
# #     page_title="M × N Editable Grid",
# #     page_icon="📊",
# #     layout="wide",
# # )

# # # ============================================================
# # # FIXED DIMENSIONS
# # # ============================================================

# # CELL_WIDTH = 120
# # ACTION_WIDTH = 55

# # HEADER_HEIGHT = 44
# # CELL_HEIGHT = 42

# # GRID_HEIGHT = 600


# # # ============================================================
# # # SESSION STATE
# # # ============================================================

# # if "rows_data" not in st.session_state:
# #     st.session_state.rows_data = [
# #         {
# #             "id": uuid.uuid4().hex,
# #             "values": ["" for _ in range(10)],
# #         }
# #         for _ in range(10)
# #     ]

# # if "num_cols" not in st.session_state:
# #     st.session_state.num_cols = 10


# # # ============================================================
# # # FUNCTIONS
# # # ============================================================

# # def create_row():
# #     return {
# #         "id": uuid.uuid4().hex,
# #         "values": [
# #             "" for _ in range(st.session_state.num_cols)
# #         ],
# #     }


# # def add_row_below(index):
# #     st.session_state.rows_data.insert(
# #         index + 1,
# #         create_row(),
# #     )


# # def delete_row(index):

# #     if len(st.session_state.rows_data) <= 1:
# #         st.warning("At least one row must remain.")
# #         return

# #     st.session_state.rows_data.pop(index)


# # def apply_grid_size(new_rows, new_cols):

# #     old_cols = st.session_state.num_cols

# #     # --------------------------------------------------------
# #     # COLUMN COUNT
# #     # --------------------------------------------------------

# #     if new_cols != old_cols:

# #         for row in st.session_state.rows_data:

# #             old_values = row["values"]

# #             if new_cols > old_cols:

# #                 row["values"] = (
# #                     old_values
# #                     + ["" for _ in range(new_cols - old_cols)]
# #                 )

# #             else:

# #                 row["values"] = old_values[:new_cols]

# #         st.session_state.num_cols = new_cols

# #     # --------------------------------------------------------
# #     # ROW COUNT
# #     # --------------------------------------------------------

# #     current_rows = len(st.session_state.rows_data)

# #     if new_rows > current_rows:

# #         for _ in range(new_rows - current_rows):
# #             st.session_state.rows_data.append(
# #                 create_row()
# #             )

# #     elif new_rows < current_rows:

# #         st.session_state.rows_data = (
# #             st.session_state.rows_data[:new_rows]
# #         )


# # # ============================================================
# # # CSS
# # # ============================================================

# # st.markdown(
# #     f"""
# #     <style>

# #     /* ======================================================
# #        PAGE
# #        ====================================================== */

# #     .block-container {{
# #         padding-top: 1.5rem;
# #         padding-left: 2rem;
# #         padding-right: 2rem;
# #     }}


# #     /* ======================================================
# #        ACTUAL STREAMLIT GRID CONTAINER
# #        ====================================================== */

# #     .st-key-grid_area {{
# #         width: 100% !important;

# #         max-width: 100% !important;

# #         height: {GRID_HEIGHT}px !important;

# #         max-height: {GRID_HEIGHT}px !important;

# #         overflow-x: auto !important;

# #         overflow-y: auto !important;

# #         box-sizing: border-box !important;

# #         border: 1px solid #cbd5e1;

# #         border-radius: 8px;

# #         background: white;
# #     }}


# #     /*
# #        Every row is a horizontal block.

# #        IMPORTANT:
# #        Do not allow the row to shrink.
# #     */

# #     .st-key-grid_area
# #     [data-testid="stHorizontalBlock"] {{

# #         display: flex !important;

# #         flex-wrap: nowrap !important;

# #         gap: 0 !important;

# #         width: max-content !important;

# #         min-width: max-content !important;

# #         max-width: none !important;

# #         flex-shrink: 0 !important;
# #     }}


# #     /* ======================================================
# #        ALL GRID COLUMNS
# #        ====================================================== */

# #     .st-key-grid_area
# #     [data-testid="stColumn"] {{

# #         flex: 0 0 {CELL_WIDTH}px !important;

# #         width: {CELL_WIDTH}px !important;

# #         min-width: {CELL_WIDTH}px !important;

# #         max-width: {CELL_WIDTH}px !important;

# #         box-sizing: border-box !important;

# #         padding-left: 0 !important;

# #         padding-right: 0 !important;

# #         margin: 0 !important;

# #         flex-shrink: 0 !important;
# #     }}


# #     /* ======================================================
# #        FIRST TWO COLUMNS = + / -
# #        ====================================================== */

# #     .st-key-grid_area
# #     [data-testid="stHorizontalBlock"]
# #     [data-testid="stColumn"]:nth-child(1),

# #     .st-key-grid_area
# #     [data-testid="stHorizontalBlock"]
# #     [data-testid="stColumn"]:nth-child(2) {{

# #         flex: 0 0 {ACTION_WIDTH}px !important;

# #         width: {ACTION_WIDTH}px !important;

# #         min-width: {ACTION_WIDTH}px !important;

# #         max-width: {ACTION_WIDTH}px !important;
# #     }}


# #     /* ======================================================
# #        HEADER
# #        ====================================================== */

# #     .grid-header {{

# #         width: {CELL_WIDTH}px !important;

# #         min-width: {CELL_WIDTH}px !important;

# #         max-width: {CELL_WIDTH}px !important;

# #         height: {HEADER_HEIGHT}px !important;

# #         min-height: {HEADER_HEIGHT}px !important;

# #         max-height: {HEADER_HEIGHT}px !important;

# #         box-sizing: border-box !important;

# #         display: flex;

# #         align-items: center;

# #         justify-content: center;

# #         background: #1f2937;

# #         color: white;

# #         border-right: 1px solid #4b5563;

# #         border-bottom: 1px solid #4b5563;

# #         font-size: 14px;

# #         font-weight: 600;

# #         white-space: nowrap;

# #         overflow: hidden;
# #     }}


# #     /* + and - headers */

# #     .grid-header-action {{

# #         width: {ACTION_WIDTH}px !important;

# #         min-width: {ACTION_WIDTH}px !important;

# #         max-width: {ACTION_WIDTH}px !important;
# #     }}


# #     /* ======================================================
# #        TEXT INPUT
# #        ====================================================== */

# #     .st-key-grid_area
# #     [data-testid="stTextInput"] {{

# #         width: {CELL_WIDTH}px !important;

# #         min-width: {CELL_WIDTH}px !important;

# #         max-width: {CELL_WIDTH}px !important;

# #         margin: 0 !important;

# #         padding: 0 !important;
# #     }}


# #     .st-key-grid_area
# #     [data-testid="stTextInput"] input {{

# #         width: {CELL_WIDTH}px !important;

# #         min-width: {CELL_WIDTH}px !important;

# #         max-width: {CELL_WIDTH}px !important;

# #         height: {CELL_HEIGHT}px !important;

# #         min-height: {CELL_HEIGHT}px !important;

# #         max-height: {CELL_HEIGHT}px !important;

# #         box-sizing: border-box !important;

# #         border-radius: 0 !important;

# #         margin: 0 !important;
# #     }}


# #     /* Hide text input labels */

# #     .st-key-grid_area
# #     [data-testid="stTextInput"] label {{

# #         display: none !important;
# #     }}


# #     /* ======================================================
# #        + / - BUTTONS
# #        ====================================================== */

# #     .st-key-grid_area
# #     [data-testid="stButton"] {{

# #         width: {ACTION_WIDTH}px !important;

# #         min-width: {ACTION_WIDTH}px !important;

# #         max-width: {ACTION_WIDTH}px !important;

# #         margin: 0 !important;

# #         padding: 0 !important;
# #     }}


# #     .st-key-grid_area
# #     [data-testid="stButton"] button {{

# #         width: {ACTION_WIDTH}px !important;

# #         min-width: {ACTION_WIDTH}px !important;

# #         max-width: {ACTION_WIDTH}px !important;

# #         height: {CELL_HEIGHT}px !important;

# #         min-height: {CELL_HEIGHT}px !important;

# #         max-height: {CELL_HEIGHT}px !important;

# #         padding: 0 !important;

# #         margin: 0 !important;

# #         border-radius: 0 !important;
# #     }}


# #     /* ======================================================
# #        SCROLLBAR
# #        ====================================================== */

# #     .st-key-grid_area::-webkit-scrollbar {{
# #         width: 12px;
# #         height: 12px;
# #     }}

# #     .st-key-grid_area::-webkit-scrollbar-track {{
# #         background: #f1f5f9;
# #     }}

# #     .st-key-grid_area::-webkit-scrollbar-thumb {{
# #         background: #94a3b8;
# #         border-radius: 6px;
# #     }}

# #     .st-key-grid_area::-webkit-scrollbar-thumb:hover {{
# #         background: #64748b;
# #     }}


# #     /* ======================================================
# #        MOBILE
# #        ====================================================== */

# #     @media (max-width: 768px) {{

# #         .block-container {{
# #             padding-left: 10px;
# #             padding-right: 10px;
# #         }}

# #     }}

# #     </style>
# #     """,
# #     unsafe_allow_html=True,
# # )


# # # ============================================================
# # # TITLE
# # # ============================================================

# # st.title("📊 M × N Editable Grid")

# # st.caption(
# #     "Fixed-size columns • Fixed-size headers • "
# #     "Horizontal & vertical scrolling"
# # )


# # # ============================================================
# # # LEFT / RIGHT LAYOUT
# # # ============================================================

# # left, right = st.columns(
# #     [1, 5],
# #     gap="large",
# # )


# # # ============================================================
# # # LEFT SETTINGS
# # # ============================================================

# # with left:

# #     st.subheader("Grid Settings")

# #     current_rows = len(
# #         st.session_state.rows_data
# #     )

# #     m = st.number_input(
# #         "Rows (M)",
# #         min_value=1,
# #         max_value=500,
# #         value=current_rows,
# #         step=1,
# #     )

# #     n = st.number_input(
# #         "Columns (N)",
# #         min_value=1,
# #         max_value=500,
# #         value=st.session_state.num_cols,
# #         step=1,
# #     )

# #     if st.button(
# #         "Apply Grid Size",
# #         type="primary",
# #         use_container_width=True,
# #     ):

# #         apply_grid_size(
# #             int(m),
# #             int(n),
# #         )

# #         st.rerun()

# #     st.divider()

# #     st.write("### Current Grid")

# #     st.write(
# #         f"Rows: **{len(st.session_state.rows_data)}**"
# #     )

# #     st.write(
# #         f"Columns: **{st.session_state.num_cols}**"
# #     )

# #     st.write(
# #         f"Cell width: **{CELL_WIDTH}px**"
# #     )

# #     st.write(
# #         f"Header height: **{HEADER_HEIGHT}px**"
# #     )


# # # ============================================================
# # # RIGHT GRID
# # # ============================================================

# # with right:

# #     st.subheader("Grid")

# #     # ========================================================
# #     # THIS IS THE REAL SCROLLABLE CONTAINER
# #     # ========================================================

# #     with st.container(
# #         height=GRID_HEIGHT,
# #         border=False,
# #         key="grid_area",
# #     ):

# #         # ====================================================
# #         # HEADER
# #         # ====================================================

# #         header_columns = st.columns(
# #             [ACTION_WIDTH, ACTION_WIDTH]
# #             + [CELL_WIDTH] * st.session_state.num_cols,
# #             gap=0,
# #         )

# #         # + header
# #         with header_columns[0]:

# #             st.markdown(
# #                 """
# #                 <div class="grid-header grid-header-action">
# #                     +
# #                 </div>
# #                 """,
# #                 unsafe_allow_html=True,
# #             )

# #         # - header
# #         with header_columns[1]:

# #             st.markdown(
# #                 """
# #                 <div class="grid-header grid-header-action">
# #                     −
# #                 </div>
# #                 """,
# #                 unsafe_allow_html=True,
# #             )

# #         # Column headers
# #         for c in range(
# #             st.session_state.num_cols
# #         ):

# #             with header_columns[c + 2]:

# #                 st.markdown(
# #                     f"""
# #                     <div class="grid-header">
# #                         Col {c + 1}
# #                     </div>
# #                     """,
# #                     unsafe_allow_html=True,
# #                 )


# #         # ====================================================
# #         # DATA ROWS
# #         # ====================================================

# #         for row_index, row in enumerate(
# #             st.session_state.rows_data
# #         ):

# #             row_id = row["id"]

# #             row_columns = st.columns(
# #                 [ACTION_WIDTH, ACTION_WIDTH]
# #                 + [CELL_WIDTH] * st.session_state.num_cols,
# #                 gap=0,
# #             )

# #             # ------------------------------------------------
# #             # ADD ROW
# #             # ------------------------------------------------

# #             with row_columns[0]:

# #                 if st.button(
# #                     "+",
# #                     key=f"add_{row_id}",
# #                     help="Insert a blank row below",
# #                 ):

# #                     add_row_below(row_index)

# #                     st.rerun()


# #             # ------------------------------------------------
# #             # DELETE ROW
# #             # ------------------------------------------------

# #             with row_columns[1]:

# #                 if st.button(
# #                     "−",
# #                     key=f"delete_{row_id}",
# #                     help="Delete this row",
# #                 ):

# #                     delete_row(row_index)

# #                     st.rerun()


# #             # ------------------------------------------------
# #             # CELLS
# #             # ------------------------------------------------

# #             for c in range(
# #                 st.session_state.num_cols
# #             ):

# #                 with row_columns[c + 2]:

# #                     value = st.text_input(
# #                         f"cell_{row_id}_{c}",
# #                         value=row["values"][c],
# #                         key=f"cell_{row_id}_{c}",
# #                         label_visibility="collapsed",
# #                     )

# #                     st.session_state.rows_data[
# #                         row_index
# #                     ]["values"][c] = value


# # # ============================================================
# # # DATA PREVIEW
# # # ============================================================

# # st.divider()

# # with st.expander("🔍 View Grid Data"):

# #     for i, row in enumerate(
# #         st.session_state.rows_data,
# #         start=1,
# #     ):

# #         st.write(
# #             f"**Row {i}:**",
# #             row["values"],
# #         )



# import streamlit as st
# import uuid


# # ============================================================
# # PAGE CONFIG
# # ============================================================

# st.set_page_config(
#     page_title="M × N Editable Grid",
#     page_icon="📊",
#     layout="wide",
# )


# # ============================================================
# # GRID SETTINGS
# # ============================================================

# CELL_WIDTH = 120
# ACTION_WIDTH = 55
# SUM_WIDTH = 120

# HEADER_HEIGHT = 44
# CELL_HEIGHT = 42

# GRID_HEIGHT = 600


# # ============================================================
# # SESSION STATE
# # ============================================================

# if "rows_data" not in st.session_state:
#     st.session_state.rows_data = [
#         {
#             "id": uuid.uuid4().hex,
#             "values": ["" for _ in range(10)],
#         }
#         for _ in range(10)
#     ]


# if "num_cols" not in st.session_state:
#     st.session_state.num_cols = 10


# # ============================================================
# # FUNCTIONS
# # ============================================================

# def create_row():
#     """
#     Create a new blank row with a unique ID.
#     """

#     return {
#         "id": uuid.uuid4().hex,
#         "values": [
#             "" for _ in range(st.session_state.num_cols)
#         ],
#     }


# def add_row_below(index):
#     """
#     Insert a blank row directly below the selected row.
#     """

#     st.session_state.rows_data.insert(
#         index + 1,
#         create_row(),
#     )


# def delete_row(index):
#     """
#     Delete the selected row.
#     Keep at least one row.
#     """

#     if len(st.session_state.rows_data) <= 1:
#         st.warning("At least one row must remain.")
#         return

#     st.session_state.rows_data.pop(index)


# def apply_grid_size(new_rows, new_cols):
#     """
#     Change number of rows and columns while preserving
#     existing values wherever possible.
#     """

#     old_cols = st.session_state.num_cols

#     # --------------------------------------------------------
#     # CHANGE COLUMN COUNT
#     # --------------------------------------------------------

#     if new_cols != old_cols:

#         for row in st.session_state.rows_data:

#             old_values = row["values"]

#             # Add new blank columns
#             if new_cols > old_cols:

#                 row["values"] = (
#                     old_values
#                     + ["" for _ in range(new_cols - old_cols)]
#                 )

#             # Remove extra columns
#             else:

#                 row["values"] = old_values[:new_cols]

#         st.session_state.num_cols = new_cols

#     # --------------------------------------------------------
#     # CHANGE ROW COUNT
#     # --------------------------------------------------------

#     current_rows = len(st.session_state.rows_data)

#     # Add rows
#     if new_rows > current_rows:

#         for _ in range(new_rows - current_rows):

#             st.session_state.rows_data.append(
#                 create_row()
#             )

#     # Remove rows
#     elif new_rows < current_rows:

#         st.session_state.rows_data = (
#             st.session_state.rows_data[:new_rows]
#         )


# # ============================================================
# # CSS
# # ============================================================

# st.markdown(
#     f"""
#     <style>

#     /* ========================================================
#        MAIN PAGE
#        ======================================================== */

#     .block-container {{
#         padding-top: 1.5rem;
#         padding-left: 2rem;
#         padding-right: 2rem;
#     }}


#     /* ========================================================
#        GRID CONTAINER
#        ======================================================== */

#     .st-key-grid_area {{
#         width: 100% !important;
#         max-width: 100% !important;

#         height: {GRID_HEIGHT}px !important;
#         max-height: {GRID_HEIGHT}px !important;

#         overflow-x: auto !important;
#         overflow-y: auto !important;

#         box-sizing: border-box !important;

#         border: 1px solid #cbd5e1;
#         border-radius: 8px;

#         background: white;
#     }}


#     /* ========================================================
#        EVERY ROW / HEADER
#        FORCE HORIZONTAL LAYOUT
#        ======================================================== */

#     .st-key-grid_area
#     [data-testid="stHorizontalBlock"] {{

#         display: flex !important;

#         flex-wrap: nowrap !important;

#         gap: 0 !important;

#         width: max-content !important;
#         min-width: max-content !important;
#         max-width: none !important;

#         flex-shrink: 0 !important;
#     }}


#     /* ========================================================
#        DEFAULT DATA COLUMN WIDTH
#        ======================================================== */

#     .st-key-grid_area
#     [data-testid="stColumn"] {{

#         flex: 0 0 {CELL_WIDTH}px !important;

#         width: {CELL_WIDTH}px !important;
#         min-width: {CELL_WIDTH}px !important;
#         max-width: {CELL_WIDTH}px !important;

#         box-sizing: border-box !important;

#         padding-left: 0 !important;
#         padding-right: 0 !important;

#         margin: 0 !important;

#         flex-shrink: 0 !important;
#     }}


#     /* ========================================================
#        FIRST TWO COLUMNS = ACTION BUTTONS
#        ======================================================== */

#     .st-key-grid_area
#     [data-testid="stHorizontalBlock"]
#     [data-testid="stColumn"]:nth-child(1),

#     .st-key-grid_area
#     [data-testid="stHorizontalBlock"]
#     [data-testid="stColumn"]:nth-child(2) {{

#         flex: 0 0 {ACTION_WIDTH}px !important;

#         width: {ACTION_WIDTH}px !important;
#         min-width: {ACTION_WIDTH}px !important;
#         max-width: {ACTION_WIDTH}px !important;

#         flex-shrink: 0 !important;
#     }}


#     /* ========================================================
#        LAST COLUMN = SUM
#        ======================================================== */

#     .st-key-grid_area
#     [data-testid="stHorizontalBlock"]
#     [data-testid="stColumn"]:last-child {{

#         flex: 0 0 {SUM_WIDTH}px !important;

#         width: {SUM_WIDTH}px !important;
#         min-width: {SUM_WIDTH}px !important;
#         max-width: {SUM_WIDTH}px !important;

#         flex-shrink: 0 !important;
#     }}


#     /* ========================================================
#        HEADER
#        ======================================================== */

#     .grid-header {{

#         width: {CELL_WIDTH}px !important;
#         min-width: {CELL_WIDTH}px !important;
#         max-width: {CELL_WIDTH}px !important;

#         height: {HEADER_HEIGHT}px !important;
#         min-height: {HEADER_HEIGHT}px !important;
#         max-height: {HEADER_HEIGHT}px !important;

#         box-sizing: border-box !important;

#         display: flex;

#         align-items: center;
#         justify-content: center;

#         background: #1f2937;

#         color: white;

#         border-right: 1px solid #4b5563;
#         border-bottom: 1px solid #4b5563;

#         font-size: 14px;
#         font-weight: 600;

#         white-space: nowrap;

#         overflow: hidden;
#     }}


#     /* ========================================================
#        ACTION HEADERS
#        ======================================================== */

#     .grid-header-action {{

#         width: {ACTION_WIDTH}px !important;
#         min-width: {ACTION_WIDTH}px !important;
#         max-width: {ACTION_WIDTH}px !important;
#     }}


#     /* ========================================================
#        SUM HEADER
#        ======================================================== */

#     .grid-header-sum {{

#         width: {SUM_WIDTH}px !important;
#         min-width: {SUM_WIDTH}px !important;
#         max-width: {SUM_WIDTH}px !important;
#     }}


#     /* ========================================================
#        TEXT INPUT CONTAINER
#        ======================================================== */

#     .st-key-grid_area
#     [data-testid="stTextInput"] {{

#         width: {CELL_WIDTH}px !important;
#         min-width: {CELL_WIDTH}px !important;
#         max-width: {CELL_WIDTH}px !important;

#         margin: 0 !important;
#         padding: 0 !important;
#     }}


#     /* ========================================================
#        TEXT INPUT
#        ======================================================== */

#     .st-key-grid_area
#     [data-testid="stTextInput"] input {{

#         width: {CELL_WIDTH}px !important;
#         min-width: {CELL_WIDTH}px !important;
#         max-width: {CELL_WIDTH}px !important;

#         height: {CELL_HEIGHT}px !important;
#         min-height: {CELL_HEIGHT}px !important;
#         max-height: {CELL_HEIGHT}px !important;

#         box-sizing: border-box !important;

#         border-radius: 0 !important;

#         margin: 0 !important;
#     }}


#     /* ========================================================
#        HIDE TEXT INPUT LABEL
#        ======================================================== */

#     .st-key-grid_area
#     [data-testid="stTextInput"] label {{

#         display: none !important;
#     }}


#     /* ========================================================
#        BUTTON CONTAINER
#        ======================================================== */

#     .st-key-grid_area
#     [data-testid="stButton"] {{

#         width: {ACTION_WIDTH}px !important;
#         min-width: {ACTION_WIDTH}px !important;
#         max-width: {ACTION_WIDTH}px !important;

#         margin: 0 !important;
#         padding: 0 !important;
#     }}


#     /* ========================================================
#        ACTION BUTTON
#        ======================================================== */

#     .st-key-grid_area
#     [data-testid="stButton"] button {{

#         width: {ACTION_WIDTH}px !important;
#         min-width: {ACTION_WIDTH}px !important;
#         max-width: {ACTION_WIDTH}px !important;

#         height: {CELL_HEIGHT}px !important;
#         min-height: {CELL_HEIGHT}px !important;
#         max-height: {CELL_HEIGHT}px !important;

#         padding: 0 !important;

#         margin: 0 !important;

#         border-radius: 0 !important;
#     }}


#     /* ========================================================
#        SUM CELL
#        ======================================================== */

#     .sum-cell {{

#         width: {SUM_WIDTH}px !important;
#         min-width: {SUM_WIDTH}px !important;
#         max-width: {SUM_WIDTH}px !important;

#         height: {CELL_HEIGHT}px !important;
#         min-height: {CELL_HEIGHT}px !important;
#         max-height: {CELL_HEIGHT}px !important;

#         box-sizing: border-box !important;

#         display: flex;

#         align-items: center;
#         justify-content: center;

#         border-right: 1px solid #d1d5db;
#         border-bottom: 1px solid #d1d5db;

#         background: #f8fafc;

#         color: #111827;

#         font-weight: 600;

#         white-space: nowrap;

#         overflow: hidden;
#     }}


#     /* ========================================================
#        SCROLLBAR
#        ======================================================== */

#     .st-key-grid_area::-webkit-scrollbar {{

#         width: 12px;
#         height: 12px;
#     }}


#     .st-key-grid_area::-webkit-scrollbar-track {{

#         background: #f1f5f9;
#     }}


#     .st-key-grid_area::-webkit-scrollbar-thumb {{

#         background: #94a3b8;

#         border-radius: 6px;
#     }}


#     .st-key-grid_area::-webkit-scrollbar-thumb:hover {{

#         background: #64748b;
#     }}


#     /* ========================================================
#        MOBILE
#        ======================================================== */

#     @media (max-width: 768px) {{

#         .block-container {{

#             padding-left: 10px;
#             padding-right: 10px;
#         }}

#     }}

#     </style>
#     """,
#     unsafe_allow_html=True,
# )


# # ============================================================
# # TITLE
# # ============================================================

# st.title("📊 M × N Editable Grid")


# # ============================================================
# # GRID SETTINGS
# # ============================================================

# left, right = st.columns([1, 3])


# with left:

#     st.subheader("Grid Settings")

#     new_rows = st.number_input(
#         "Number of Rows",
#         min_value=1,
#         max_value=500,
#         value=len(st.session_state.rows_data),
#         step=1,
#     )

#     new_cols = st.number_input(
#         "Number of Columns",
#         min_value=1,
#         max_value=200,
#         value=st.session_state.num_cols,
#         step=1,
#     )

#     if st.button(
#         "Apply Grid Size",
#         use_container_width=True,
#     ):

#         apply_grid_size(
#             int(new_rows),
#             int(new_cols),
#         )

#         st.rerun()


# # ============================================================
# # GRID
# # ============================================================

# with right:

#     st.subheader("Grid")

#     # --------------------------------------------------------
#     # REAL STREAMLIT SCROLL CONTAINER
#     # --------------------------------------------------------

#     with st.container(
#         height=GRID_HEIGHT,
#         border=False,
#         key="grid_area",
#     ):

#         # ====================================================
#         # HEADER
#         # ====================================================

#         header_columns = st.columns(
#             [ACTION_WIDTH, ACTION_WIDTH]
#             + [CELL_WIDTH] * st.session_state.num_cols
#             + [SUM_WIDTH],
#             gap=0,
#         )


#         # ----------------------------------------------------
#         # PLUS HEADER
#         # ----------------------------------------------------

#         with header_columns[0]:

#             st.markdown(
#                 """
#                 <div class="grid-header grid-header-action">
#                     +
#                 </div>
#                 """,
#                 unsafe_allow_html=True,
#             )


#         # ----------------------------------------------------
#         # MINUS HEADER
#         # ----------------------------------------------------

#         with header_columns[1]:

#             st.markdown(
#                 """
#                 <div class="grid-header grid-header-action">
#                     −
#                 </div>
#                 """,
#                 unsafe_allow_html=True,
#             )


#         # ----------------------------------------------------
#         # DATA COLUMN HEADERS
#         # ----------------------------------------------------

#         for c in range(st.session_state.num_cols):

#             with header_columns[c + 2]:

#                 st.markdown(
#                     f"""
#                     <div class="grid-header">
#                         Col {c + 1}
#                     </div>
#                     """,
#                     unsafe_allow_html=True,
#                 )


#         # ----------------------------------------------------
#         # SUM HEADER
#         # ----------------------------------------------------

#         with header_columns[-1]:

#             st.markdown(
#                 """
#                 <div class="grid-header grid-header-sum">
#                     Sum
#                 </div>
#                 """,
#                 unsafe_allow_html=True,
#             )


#         # ====================================================
#         # DATA ROWS
#         # ====================================================

#         for row_index, row in enumerate(
#             st.session_state.rows_data
#         ):

#             row_id = row["id"]


#             # ------------------------------------------------
#             # CREATE ROW COLUMNS
#             # ------------------------------------------------

#             row_columns = st.columns(
#                 [ACTION_WIDTH, ACTION_WIDTH]
#                 + [CELL_WIDTH] * st.session_state.num_cols
#                 + [SUM_WIDTH],
#                 gap=0,
#             )


#             # ------------------------------------------------
#             # ADD ROW BUTTON
#             # ------------------------------------------------

#             with row_columns[0]:

#                 if st.button(
#                     "+",
#                     key=f"add_{row_id}",
#                     help="Insert a blank row below",
#                 ):

#                     add_row_below(row_index)

#                     st.rerun()


#             # ------------------------------------------------
#             # DELETE ROW BUTTON
#             # ------------------------------------------------

#             with row_columns[1]:

#                 if st.button(
#                     "−",
#                     key=f"delete_{row_id}",
#                     help="Delete this row",
#                 ):

#                     delete_row(row_index)

#                     st.rerun()


#             # ------------------------------------------------
#             # EDITABLE DATA CELLS
#             # ------------------------------------------------

#             for c in range(
#                 st.session_state.num_cols
#             ):

#                 with row_columns[c + 2]:

#                     value = st.text_input(
#                         f"cell_{row_id}_{c}",

#                         value=row["values"][c],

#                         key=f"cell_{row_id}_{c}",

#                         label_visibility="collapsed",
#                     )

#                     st.session_state.rows_data[
#                         row_index
#                     ]["values"][c] = value


#             # ------------------------------------------------
#             # SUM CELL
#             # ------------------------------------------------

#             with row_columns[-1]:

#                 st.markdown(
#                     """
#                     <div class="sum-cell">
                        
#                     </div>
#                     """,
#                     unsafe_allow_html=True,
#                 )



import streamlit as st
import uuid


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="M × N Editable Grid",
    page_icon="📊",
    layout="wide",
)


# ============================================================
# GRID SETTINGS
# ============================================================

CELL_WIDTH = 120
ACTION_WIDTH = 55
SUM_WIDTH = 120

HEADER_HEIGHT = 44
CELL_HEIGHT = 42
TOTAL_HEIGHT = 44

GRID_HEIGHT = 600


# ============================================================
# SESSION STATE
# ============================================================

if "rows_data" not in st.session_state:
    st.session_state.rows_data = [
        {
            "id": uuid.uuid4().hex,
            "values": ["" for _ in range(10)],
        }
        for _ in range(10)
    ]


if "num_cols" not in st.session_state:
    st.session_state.num_cols = 10


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def create_row():
    """
    Create a new blank row with a unique ID.
    """

    return {
        "id": uuid.uuid4().hex,
        "values": [
            "" for _ in range(st.session_state.num_cols)
        ],
    }


def number_value(value):
    """
    Convert a cell value into a number.

    Examples:
        "100"       -> 100
        "100.50"    -> 100.50
        "-25"       -> -25
        ""          -> 0
        "ABC"       -> 0
    """

    if value is None:
        return 0.0

    text = str(value).strip()

    if text == "":
        return 0.0

    # Remove commas so values such as:
    # 1,000
    # 25,500.50
    # can also be calculated.
    text = text.replace(",", "")

    try:
        return float(text)

    except (ValueError, TypeError):
        return 0.0


def format_number(value):
    """
    Format calculated numbers nicely.

    100.0   -> 100
    100.50  -> 100.5
    100.25  -> 100.25
    """

    if value == int(value):
        return f"{int(value):,}"

    return f"{value:,.2f}".rstrip("0").rstrip(".")


def calculate_row_sum(row):
    """
    Calculate the total of all numeric cells in one row.
    """

    total = 0.0

    for value in row["values"]:
        total += number_value(value)

    return total


def calculate_column_total(column_index):
    """
    Calculate total for one column across all rows.
    """

    total = 0.0

    for row in st.session_state.rows_data:

        if column_index < len(row["values"]):

            total += number_value(
                row["values"][column_index]
            )

    return total


def calculate_grand_total():
    """
    Calculate total of every numeric cell in the grid.
    """

    total = 0.0

    for row in st.session_state.rows_data:

        for value in row["values"]:

            total += number_value(value)

    return total


# ============================================================
# ROW OPERATIONS
# ============================================================

def add_row_below(index):
    """
    Insert a blank row directly below selected row.
    """

    st.session_state.rows_data.insert(
        index + 1,
        create_row(),
    )


def delete_row(index):
    """
    Delete selected row.
    At least one row must remain.
    """

    if len(st.session_state.rows_data) <= 1:

        st.warning(
            "At least one row must remain."
        )

        return

    st.session_state.rows_data.pop(index)


# ============================================================
# GRID SIZE
# ============================================================

def apply_grid_size(new_rows, new_cols):

    old_cols = st.session_state.num_cols

    # --------------------------------------------------------
    # CHANGE COLUMN COUNT
    # --------------------------------------------------------

    if new_cols != old_cols:

        for row in st.session_state.rows_data:

            old_values = row["values"]

            # Add columns
            if new_cols > old_cols:

                row["values"] = (
                    old_values
                    + [
                        ""
                        for _ in range(
                            new_cols - old_cols
                        )
                    ]
                )

            # Remove columns
            else:

                row["values"] = (
                    old_values[:new_cols]
                )

        st.session_state.num_cols = new_cols

    # --------------------------------------------------------
    # CHANGE ROW COUNT
    # --------------------------------------------------------

    current_rows = len(
        st.session_state.rows_data
    )

    # Add rows
    if new_rows > current_rows:

        for _ in range(
            new_rows - current_rows
        ):

            st.session_state.rows_data.append(
                create_row()
            )

    # Remove rows
    elif new_rows < current_rows:

        st.session_state.rows_data = (
            st.session_state.rows_data[
                :new_rows
            ]
        )


# ============================================================
# CSS
# ============================================================

st.markdown(
    f"""
    <style>

    /* ========================================================
       MAIN PAGE
       ======================================================== */

    .block-container {{
        padding-top: 1.5rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }}


    /* ========================================================
       GRID SCROLL CONTAINER
       ======================================================== */

    .st-key-grid_area {{

        width: 100% !important;
        max-width: 100% !important;

        height: {GRID_HEIGHT}px !important;
        max-height: {GRID_HEIGHT}px !important;

        overflow-x: auto !important;
        overflow-y: auto !important;

        box-sizing: border-box !important;

        border: 1px solid #cbd5e1;

        border-radius: 8px;

        background: white;
    }}


    /* ========================================================
       EVERY HORIZONTAL ROW
       ======================================================== */

    .st-key-grid_area
    [data-testid="stHorizontalBlock"] {{

        display: flex !important;

        flex-wrap: nowrap !important;

        gap: 0 !important;

        width: max-content !important;

        min-width: max-content !important;

        max-width: none !important;

        flex-shrink: 0 !important;
    }}


    /* ========================================================
       DEFAULT DATA COLUMN
       ======================================================== */

    .st-key-grid_area
    [data-testid="stColumn"] {{

        flex: 0 0 {CELL_WIDTH}px !important;

        width: {CELL_WIDTH}px !important;

        min-width: {CELL_WIDTH}px !important;

        max-width: {CELL_WIDTH}px !important;

        box-sizing: border-box !important;

        padding-left: 0 !important;

        padding-right: 0 !important;

        margin: 0 !important;

        flex-shrink: 0 !important;
    }}


    /* ========================================================
       FIRST TWO ACTION COLUMNS
       ======================================================== */

    .st-key-grid_area
    [data-testid="stHorizontalBlock"]
    [data-testid="stColumn"]:nth-child(1),

    .st-key-grid_area
    [data-testid="stHorizontalBlock"]
    [data-testid="stColumn"]:nth-child(2) {{

        flex: 0 0 {ACTION_WIDTH}px !important;

        width: {ACTION_WIDTH}px !important;

        min-width: {ACTION_WIDTH}px !important;

        max-width: {ACTION_WIDTH}px !important;

        flex-shrink: 0 !important;
    }}


    /* ========================================================
       LAST COLUMN = SUM
       ======================================================== */

    .st-key-grid_area
    [data-testid="stHorizontalBlock"]
    [data-testid="stColumn"]:last-child {{

        flex: 0 0 {SUM_WIDTH}px !important;

        width: {SUM_WIDTH}px !important;

        min-width: {SUM_WIDTH}px !important;

        max-width: {SUM_WIDTH}px !important;

        flex-shrink: 0 !important;
    }}


    /* ========================================================
       NORMAL HEADER
       ======================================================== */

    .grid-header {{

        width: {CELL_WIDTH}px !important;

        min-width: {CELL_WIDTH}px !important;

        max-width: {CELL_WIDTH}px !important;

        height: {HEADER_HEIGHT}px !important;

        min-height: {HEADER_HEIGHT}px !important;

        max-height: {HEADER_HEIGHT}px !important;

        box-sizing: border-box !important;

        display: flex;

        align-items: center;

        justify-content: center;

        background: #1f2937;

        color: white;

        border-right: 1px solid #4b5563;

        border-bottom: 1px solid #4b5563;

        font-size: 14px;

        font-weight: 600;

        white-space: nowrap;

        overflow: hidden;
    }}


    /* ========================================================
       ACTION HEADER
       ======================================================== */

    .grid-header-action {{

        width: {ACTION_WIDTH}px !important;

        min-width: {ACTION_WIDTH}px !important;

        max-width: {ACTION_WIDTH}px !important;
    }}


    /* ========================================================
       SUM HEADER
       ======================================================== */

    .grid-header-sum {{

        width: {SUM_WIDTH}px !important;

        min-width: {SUM_WIDTH}px !important;

        max-width: {SUM_WIDTH}px !important;
    }}


    /* ========================================================
       TEXT INPUT CONTAINER
       ======================================================== */

    .st-key-grid_area
    [data-testid="stTextInput"] {{

        width: {CELL_WIDTH}px !important;

        min-width: {CELL_WIDTH}px !important;

        max-width: {CELL_WIDTH}px !important;

        margin: 0 !important;

        padding: 0 !important;
    }}


    /* ========================================================
       TEXT INPUT
       ======================================================== */

    .st-key-grid_area
    [data-testid="stTextInput"] input {{

        width: {CELL_WIDTH}px !important;

        min-width: {CELL_WIDTH}px !important;

        max-width: {CELL_WIDTH}px !important;

        height: {CELL_HEIGHT}px !important;

        min-height: {CELL_HEIGHT}px !important;

        max-height: {CELL_HEIGHT}px !important;

        box-sizing: border-box !important;

        border-radius: 0 !important;

        margin: 0 !important;
    }}


    /* ========================================================
       HIDE INPUT LABEL
       ======================================================== */

    .st-key-grid_area
    [data-testid="stTextInput"] label {{

        display: none !important;
    }}


    /* ========================================================
       BUTTON CONTAINER
       ======================================================== */

    .st-key-grid_area
    [data-testid="stButton"] {{

        width: {ACTION_WIDTH}px !important;

        min-width: {ACTION_WIDTH}px !important;

        max-width: {ACTION_WIDTH}px !important;

        margin: 0 !important;

        padding: 0 !important;
    }}


    /* ========================================================
       ACTION BUTTON
       ======================================================== */

    .st-key-grid_area
    [data-testid="stButton"] button {{

        width: {ACTION_WIDTH}px !important;

        min-width: {ACTION_WIDTH}px !important;

        max-width: {ACTION_WIDTH}px !important;

        height: {CELL_HEIGHT}px !important;

        min-height: {CELL_HEIGHT}px !important;

        max-height: {CELL_HEIGHT}px !important;

        padding: 0 !important;

        margin: 0 !important;

        border-radius: 0 !important;
    }}


    /* ========================================================
       ROW SUM CELL
       ======================================================== */

    .sum-cell {{

        width: {SUM_WIDTH}px !important;

        min-width: {SUM_WIDTH}px !important;

        max-width: {SUM_WIDTH}px !important;

        height: {CELL_HEIGHT}px !important;

        min-height: {CELL_HEIGHT}px !important;

        max-height: {CELL_HEIGHT}px !important;

        box-sizing: border-box !important;

        display: flex;

        align-items: center;

        justify-content: center;

        border-right: 1px solid #d1d5db;

        border-bottom: 1px solid #d1d5db;

        background: #f8fafc;

        color: #111827;

        font-weight: 600;

        white-space: nowrap;

        overflow: hidden;
    }}


    /* ========================================================
       TOTAL ROW
       ======================================================== */

    .total-cell {{

        width: {CELL_WIDTH}px !important;

        min-width: {CELL_WIDTH}px !important;

        max-width: {CELL_WIDTH}px !important;

        height: {TOTAL_HEIGHT}px !important;

        min-height: {TOTAL_HEIGHT}px !important;

        max-height: {TOTAL_HEIGHT}px !important;

        box-sizing: border-box !important;

        display: flex;

        align-items: center;

        justify-content: center;

        background: #e5e7eb;

        color: #111827;

        border-right: 1px solid #9ca3af;

        border-bottom: 2px solid #6b7280;

        font-weight: 700;

        white-space: nowrap;

        overflow: hidden;
    }}


    /* ========================================================
       TOTAL ACTION CELLS
       ======================================================== */

    .total-action-cell {{

        width: {ACTION_WIDTH}px !important;

        min-width: {ACTION_WIDTH}px !important;

        max-width: {ACTION_WIDTH}px !important;

        height: {TOTAL_HEIGHT}px !important;

        min-height: {TOTAL_HEIGHT}px !important;

        max-height: {TOTAL_HEIGHT}px !important;

        box-sizing: border-box !important;

        display: flex;

        align-items: center;

        justify-content: center;

        background: #e5e7eb;

        color: #111827;

        border-right: 1px solid #9ca3af;

        border-bottom: 2px solid #6b7280;

        font-weight: 700;
    }}


    /* ========================================================
       TOTAL SUM CELL
       ======================================================== */

    .total-sum-cell {{

        width: {SUM_WIDTH}px !important;

        min-width: {SUM_WIDTH}px !important;

        max-width: {SUM_WIDTH}px !important;

        height: {TOTAL_HEIGHT}px !important;

        min-height: {TOTAL_HEIGHT}px !important;

        max-height: {TOTAL_HEIGHT}px !important;

        box-sizing: border-box !important;

        display: flex;

        align-items: center;

        justify-content: center;

        background: #d1d5db;

        color: #111827;

        border-right: 1px solid #9ca3af;

        border-bottom: 2px solid #6b7280;

        font-weight: 700;

        white-space: nowrap;

        overflow: hidden;
    }}


    /* ========================================================
       SCROLLBAR
       ======================================================== */

    .st-key-grid_area::-webkit-scrollbar {{

        width: 12px;

        height: 12px;
    }}


    .st-key-grid_area::-webkit-scrollbar-track {{

        background: #f1f5f9;
    }}


    .st-key-grid_area::-webkit-scrollbar-thumb {{

        background: #94a3b8;

        border-radius: 6px;
    }}


    .st-key-grid_area::-webkit-scrollbar-thumb:hover {{

        background: #64748b;
    }}


    /* ========================================================
       MOBILE
       ======================================================== */

    @media (max-width: 768px) {{

        .block-container {{

            padding-left: 10px;

            padding-right: 10px;
        }}
    }}

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PAGE TITLE
# ============================================================

st.title("📊 M × N Editable Grid")


# ============================================================
# SETTINGS + GRID
# ============================================================

left, right = st.columns([1, 3])


# ============================================================
# LEFT: SETTINGS
# ============================================================

with left:

    st.subheader("Grid Settings")

    new_rows = st.number_input(
        "Number of Rows",
        min_value=1,
        max_value=500,
        value=len(
            st.session_state.rows_data
        ),
        step=1,
    )

    new_cols = st.number_input(
        "Number of Columns",
        min_value=1,
        max_value=200,
        value=st.session_state.num_cols,
        step=1,
    )

    if st.button(
        "Apply Grid Size",
        use_container_width=True,
    ):

        apply_grid_size(
            int(new_rows),
            int(new_cols),
        )

        st.rerun()


# ============================================================
# RIGHT: GRID
# ============================================================

with right:

    st.subheader("Grid")

    # ========================================================
    # SCROLLABLE GRID CONTAINER
    # ========================================================

    with st.container(
        height=GRID_HEIGHT,
        border=False,
        key="grid_area",
    ):

        # ====================================================
        # HEADER
        # ====================================================

        header_columns = st.columns(
            [ACTION_WIDTH, ACTION_WIDTH]
            + [
                CELL_WIDTH
                for _ in range(
                    st.session_state.num_cols
                )
            ]
            + [SUM_WIDTH],
            gap=0,
        )


        # ----------------------------------------------------
        # PLUS HEADER
        # ----------------------------------------------------

        with header_columns[0]:

            st.markdown(
                """
                <div class="grid-header grid-header-action">
                    +
                </div>
                """,
                unsafe_allow_html=True,
            )


        # ----------------------------------------------------
        # MINUS HEADER
        # ----------------------------------------------------

        with header_columns[1]:

            st.markdown(
                """
                <div class="grid-header grid-header-action">
                    −
                </div>
                """,
                unsafe_allow_html=True,
            )


        # ----------------------------------------------------
        # DATA HEADERS
        # ----------------------------------------------------

        for c in range(
            st.session_state.num_cols
        ):

            with header_columns[c + 2]:

                st.markdown(
                    f"""
                    <div class="grid-header">
                        Col {c + 1}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


        # ----------------------------------------------------
        # SUM HEADER
        # ----------------------------------------------------

        with header_columns[-1]:

            st.markdown(
                """
                <div class="grid-header grid-header-sum">
                    Sum
                </div>
                """,
                unsafe_allow_html=True,
            )


        # ====================================================
        # NORMAL DATA ROWS
        # ====================================================

        for row_index, row in enumerate(
            st.session_state.rows_data
        ):

            row_id = row["id"]


            # ------------------------------------------------
            # CREATE ROW
            # ------------------------------------------------

            row_columns = st.columns(
                [ACTION_WIDTH, ACTION_WIDTH]
                + [
                    CELL_WIDTH
                    for _ in range(
                        st.session_state.num_cols
                    )
                ]
                + [SUM_WIDTH],
                gap=0,
            )


            # ------------------------------------------------
            # ADD ROW
            # ------------------------------------------------

            with row_columns[0]:

                if st.button(
                    "+",
                    key=f"add_{row_id}",
                    help="Insert a blank row below",
                ):

                    add_row_below(
                        row_index
                    )

                    st.rerun()


            # ------------------------------------------------
            # DELETE ROW
            # ------------------------------------------------

            with row_columns[1]:

                if st.button(
                    "−",
                    key=f"delete_{row_id}",
                    help="Delete this row",
                ):

                    delete_row(
                        row_index
                    )

                    st.rerun()


            # ------------------------------------------------
            # EDITABLE CELLS
            # ------------------------------------------------

            for c in range(
                st.session_state.num_cols
            ):

                with row_columns[c + 2]:

                    value = st.text_input(
                        f"cell_{row_id}_{c}",

                        value=row["values"][c],

                        key=f"cell_{row_id}_{c}",

                        label_visibility="collapsed",
                    )

                    st.session_state.rows_data[
                        row_index
                    ]["values"][c] = value


            # ------------------------------------------------
            # AUTOMATIC ROW SUM
            # ------------------------------------------------

            row_sum = calculate_row_sum(
                row
            )

            with row_columns[-1]:

                st.markdown(
                    f"""
                    <div class="sum-cell">
                        {format_number(row_sum)}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


        # ====================================================
        # FIXED TOTAL ROW
        # ====================================================

        total_columns = st.columns(
            [ACTION_WIDTH, ACTION_WIDTH]
            + [
                CELL_WIDTH
                for _ in range(
                    st.session_state.num_cols
                )
            ]
            + [SUM_WIDTH],
            gap=0,
        )


        # ----------------------------------------------------
        # TOTAL LABEL
        # ----------------------------------------------------

        with total_columns[0]:

            st.markdown(
                """
                <div class="total-action-cell">
                    
                </div>
                """,
                unsafe_allow_html=True,
            )


        # ----------------------------------------------------
        # TOTAL LABEL
        # ----------------------------------------------------

        with total_columns[1]:

            st.markdown(
                """
                <div class="total-action-cell">
                    TOTAL
                </div>
                """,
                unsafe_allow_html=True,
            )


        # ====================================================
        # COLUMN TOTALS
        # ====================================================

        for c in range(
            st.session_state.num_cols
        ):

            column_total = (
                calculate_column_total(c)
            )

            with total_columns[c + 2]:

                st.markdown(
                    f"""
                    <div class="total-cell">
                        {format_number(column_total)}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


        # ====================================================
        # GRAND TOTAL
        # ====================================================

        grand_total = calculate_grand_total()

        with total_columns[-1]:

            st.markdown(
                f"""
                <div class="total-sum-cell">
                    {format_number(grand_total)}
                </div>
                """,
                unsafe_allow_html=True,
            )