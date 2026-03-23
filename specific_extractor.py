import pdfplumber
import json

TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "intersection_tolerance": 5,
    "join_tolerance": 5,
    "snap_tolerance": 3,
    "edge_min_length": 20,
    "min_words_vertical": 3,
    "min_words_horizontal": 1,
}

def is_semantic_table_candidate(table_data):
    if not table_data: return False
    rows = len(table_data)
    cols = max((len(r) for r in table_data), default=0)
    if rows < 2 or cols < 2: return False
    non_empty_cells = sum(1 for row in table_data for c in row if c and str(c).strip())
    dense_rows = sum(1 for row in table_data if sum(1 for c in row if c and str(c).strip()) >= 2)
    if dense_rows < 2: return False
    occupancy = non_empty_cells / max(1, rows * cols)
    if occupancy < 0.2: return False
    return True

def generate_specific_json(filename, json_path):
    pages_data = []
    with pdfplumber.open(filename) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            page_obj = {"page_number": page_idx, "elements": []}
            detected_tables = page.find_tables(table_settings=TABLE_SETTINGS)
            for table in detected_tables:
                table_data = table.extract()
                if not table_data or not is_semantic_table_candidate(table_data): continue
                
                rows = len(table_data)
                cols = max(len(row) for row in table_data) if table_data else 0
                grid = [[None for _ in range(cols)] for _ in range(rows)]
                occupied = set()

                for row_idx, row in enumerate(table_data):
                    col_idx = 0
                    for cell_idx, cell in enumerate(row):
                        while (row_idx, col_idx) in occupied:
                            col_idx += 1
                        if col_idx >= cols: break
                        if cell is None: continue

                        rowspan = 1
                        colspan = 1

                        if cell_idx < len(row) - 1 and row[cell_idx + 1] is None:
                            next_col = col_idx + 1
                            while next_col < cols and (row_idx, next_col) not in occupied and (
                                    next_col >= len(row) or row[next_col] is None):
                                colspan += 1; next_col += 1

                        if row_idx < len(table_data) - 1 and len(table_data[row_idx + 1]) > col_idx and \
                                table_data[row_idx + 1][col_idx] is None:
                            next_row = row_idx + 1
                            while next_row < rows and (next_row, col_idx) not in occupied and (
                                    col_idx >= len(table_data[next_row]) or table_data[next_row][col_idx] is None):
                                rowspan += 1; next_row += 1

                        for r in range(row_idx, row_idx + rowspan):
                            for c in range(col_idx, col_idx + colspan):
                                if r < rows and c < cols: occupied.add((r, c))

                        text = str(cell or "").replace("\n", " ").strip()
                        grid[row_idx][col_idx] = {
                            'text': text,
                            'rowspan': str(rowspan) if rowspan > 1 else None,
                            'colspan': str(colspan) if colspan > 1 else None
                        }
                        col_idx += colspan
                        
                json_rows = []
                for r_idx in range(rows):
                    json_row = []
                    for c_idx in range(cols):
                        if (r_idx, c_idx) in occupied and grid[r_idx][c_idx] is None: continue
                        cell = grid[r_idx][c_idx]
                        if cell: json_row.append(cell)
                        else: json_row.append({'text': '', 'rowspan': None, 'colspan': None})
                    json_rows.append(json_row)

                page_obj["elements"].append({"type": "table", "rows": json_rows})
            pages_data.append(page_obj)

    json_data = {"document": {"pages": pages_data}}
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, indent=4)
