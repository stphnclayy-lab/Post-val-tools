"""Membaca file tabel (xlsx atau csv). Header ada di baris 1, data dibaca dari sheet pertama."""
from __future__ import annotations

import csv
import io
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def kosong(nilai: Any) -> bool:
    """True jika sel dianggap kosong (None atau teks spasi saja)."""
    return nilai is None or (isinstance(nilai, str) and nilai.strip() == "")


@dataclass
class DataExcel:
    nama_file: str
    header: list[str]
    baris: list[list[Any]] = field(default_factory=list)

    def posisi(self, nama_kolom: str) -> int | None:
        """Nomor urut kolom (0-based) untuk nama header tertentu, atau None."""
        return self.header.index(nama_kolom) if nama_kolom in self.header else None


def _rapikan(baris_header, semua_baris) -> tuple[list[str], list[list[Any]]]:
    header = ["" if kosong(v) else str(v).strip() for v in baris_header]
    while header and header[-1] == "":
        header.pop()  # buang kolom kosong di ujung
    lebar = len(header)
    hasil: list[list[Any]] = []
    for nilai_baris in semua_baris:
        isi = list(nilai_baris[:lebar]) + [None] * max(0, lebar - len(nilai_baris))
        if all(kosong(v) for v in isi):
            continue  # lewati baris kosong
        hasil.append(isi)
    return header, hasil


def baca_excel(sumber: bytes | str | Path, nama_file: str) -> DataExcel:
    """Baca sheet pertama dari file xlsx. `sumber` bisa bytes (hasil upload) atau path."""
    if isinstance(sumber, bytes):
        sumber = BytesIO(sumber)
    wb = load_workbook(sumber, read_only=True, data_only=True)
    try:
        semua = wb.worksheets[0].iter_rows(values_only=True)
        baris_header = next(semua, ())
        header, baris = _rapikan(baris_header, semua)
    finally:
        wb.close()
    return DataExcel(nama_file=nama_file, header=header, baris=baris)


def _decode(data: bytes) -> str:
    """CSV dari Excel bisa UTF-8 (kadang dengan BOM) atau Windows-1252. Coba keduanya."""
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("latin-1")


def _pemisah_csv(teks: str) -> str:
    """Pemisah dicari dari baris header: yang paling banyak muncul. Excel Indonesia sering pakai ';'."""
    baris_pertama = next((b for b in teks.splitlines() if b.strip()), "")
    kandidat = [",", ";", "\t", "|"]
    jumlah = {k: baris_pertama.count(k) for k in kandidat}
    terbanyak = max(kandidat, key=lambda k: jumlah[k])
    return terbanyak if jumlah[terbanyak] > 0 else ","


def baca_csv(sumber: bytes | str | Path, nama_file: str) -> DataExcel:
    """Baca file CSV. Semua nilai dibaca sebagai teks, seperti isi file aslinya."""
    data = Path(sumber).read_bytes() if isinstance(sumber, (str, Path)) else sumber
    teks = _decode(data)
    pemisah = _pemisah_csv(teks)
    baris = list(csv.reader(io.StringIO(teks, newline=""), delimiter=pemisah))
    if not baris:
        return DataExcel(nama_file=nama_file, header=[], baris=[])
    baris = [[None if v == "" else v for v in r] for r in baris]
    header, data_baris = _rapikan(baris[0], baris[1:])
    return DataExcel(nama_file=nama_file, header=header, baris=data_baris)


def baca_tabel(sumber: bytes | str | Path, nama_file: str) -> DataExcel:
    """Pilih pembaca sesuai ekstensi file: .csv atau .xlsx."""
    if nama_file.lower().endswith(".csv"):
        return baca_csv(sumber, nama_file)
    return baca_excel(sumber, nama_file)
