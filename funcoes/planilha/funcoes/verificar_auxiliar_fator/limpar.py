def _limpar_codigo(codigo):
    """Limpa o código, removendo caracteres especiais."""
    return (
        codigo.replace("\u200b", "").replace("\ufeff", "").strip().split()[0]
        if codigo.replace("\u200b", "").replace("\ufeff", "").strip().split()
        else ""
    )


def _codigo_valido(sheet, row, col_desc_idx):
    """Verifica se a linha contém um código válido."""
    from openpyxl.cell.cell import MergedCell

    val = sheet.cell(row=row, column=col_desc_idx).value
    if not val:
        return False
    s = str(val).strip()
    if not s:
        return False
    for mr in sheet.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col_desc_idx <= mr.max_col:
            if mr.max_col - mr.min_col > 2:
                return False
    return s[0].isdigit() or (s[0].isalpha() and any(c.isdigit() for c in s))
