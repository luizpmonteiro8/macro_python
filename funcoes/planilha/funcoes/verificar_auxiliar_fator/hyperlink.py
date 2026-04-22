from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.hyperlink import Hyperlink


def _add_hyperlink(sheet, row, col, planilha, linha_ref):
    """Adiciona hyperlink à célula."""
    if linha_ref > 0:
        cell = sheet.cell(row=row, column=col)
        if not isinstance(cell, MergedCell) and not cell.hyperlink:
            cell.hyperlink = Hyperlink(
                ref=cell.coordinate, location=f"'{planilha}'!A{linha_ref}"
            )
