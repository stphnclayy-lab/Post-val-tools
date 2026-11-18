"""Menulis file Hasil (.xlsx).

Setiap pasangan kolom menjadi tiga kolom: [kolom input | kolom module | check].
Kolom check berisi formula =IF(EXACT(...),"Sama","Beda"), persis seperti contoh di sheet Hasil.
"""
from __future__ import annotations

from io import BytesIO
from typing import Any, Sequence

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


def _tulis_nilai(ws, baris: int, kolom: int, nilai: Any) -> None:
    if nilai is None:
        return
    sel = ws.cell(row=baris, column=kolom)
    sel.value = nilai
    if isinstance(nilai, str) and nilai.startswith("="):
        sel.data_type = "s"  # teks yang kebetulan diawali '=', jangan dibaca sebagai formula


def buat_xlsx_hasil(
    pasangan: Sequence[tuple[str, str]],
    baris: Sequence[tuple[list | None, list | None]],
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Hasil"

    header: list[str] = []
    for kf, km in pasangan:
        header += [kf, km, "check"]
    ws.append(header)
    for kolom in range(1, len(header) + 1):
        sel = ws.cell(row=1, column=kolom)
        sel.font = Font(bold=True)
        sel.alignment = Alignment(wrap_text=True, vertical="center")

    for r, (bi, be) in enumerate(baris, start=2):
        for j in range(len(pasangan)):
            kol_a, kol_b, kol_c = 3 * j + 1, 3 * j + 2, 3 * j + 3
            _tulis_nilai(ws, r, kol_a, bi[j] if bi is not None else None)
            _tulis_nilai(ws, r, kol_b, be[j] if be is not None else None)
            a = get_column_letter(kol_a)
            b = get_column_letter(kol_b)
            ws.cell(row=r, column=kol_c, value=f'=IF(EXACT({a}{r},{b}{r}),"Sama","Beda")')

    akhir = max(len(baris) + 1, 2)
    sorot_beda = CellIsRule(
        operator="equal",
        formula=['"Beda"'],
        fill=PatternFill("solid", fgColor="F8CBAD"),
        font=Font(bold=True, color="9C0006"),
    )
    for j in range(len(pasangan)):
        kol_a, kol_b, kol_c = 3 * j + 1, 3 * j + 2, 3 * j + 3
        huruf_c = get_column_letter(kol_c)
        ws.conditional_formatting.add(f"{huruf_c}2:{huruf_c}{akhir}", sorot_beda)
        ws.column_dimensions[get_column_letter(kol_a)].width = 24
        ws.column_dimensions[get_column_letter(kol_b)].width = 24
        ws.column_dimensions[huruf_c].width = 9

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{akhir}"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
