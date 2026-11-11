from pathlib import Path
from openpyxl import Workbook, load_workbook


class ExcelManager:
    def __init__(self, file_path):
        self.file_path = Path(file_path)
        self.file_path.parent.mkdir(parents=True, exist_ok=True)

    def ensure_file(self, headers):
        if not self.file_path.exists():
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.append(headers)
            wb.save(self.file_path)

    def _next_available_row(self, ws, headers):
        total_cols = len(headers)

        for row in range(1, ws.max_row + 1):
            empty = True
            for col in range(1, total_cols + 1):
                if ws.cell(row=row, column=col).value is not None:
                    empty = False
                    break
            if empty:
                return row

        return ws.max_row + 1

    def append_row(self, headers, values):
        self.ensure_file(headers)

        wb = load_workbook(self.file_path)
        ws = wb.active

        if ws.max_row == 1 and ws["A1"].value is None:
            ws.append(headers)

        row_index = self._next_available_row(ws, headers)

        for col_index, header in enumerate(headers, start=1):
            ws.cell(row=row_index, column=col_index, value=values.get(header, ""))

        wb.save(self.file_path)