"""Membuat file contoh (data dummy) untuk dicoba.

Jalankan:  python buat_contoh.py
Hasilnya ada di folder contoh/:
  - mapping_contoh.xlsx, 1.A.1 Aset tetap.xlsx, INP.ANR.AT.xlsx  (satu pasangan)
  - batch/                                                        (tiga pasangan, export berupa CSV)
"""
from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook

HEADER_MAPPING = ["nama file", "nama module", "column file", "column module"]
HEADER_INPUT = [
    "file_path",
    "period_koreksi_dms",
    "company_code_dms",
    "uraian_koreksi_dms",
    "kategori_dms",
    "perode_31_januari_2025",
    "total_dms",
]
HEADER_EXPORT = ["file_path", "time", "kode perusahaan", "keterangan", "kelompok", "saldo awal tahun"]

MAPPING = [
    ["1.A.1 Aset tetap", "INP.ANR.AT", "file_path", None],
    ["1.A.1 Aset tetap", "INP.ANR.AT", "period_koreksi_dms", "time"],
    ["1.A.1 Aset tetap", "INP.ANR.AT", "company_code_dms", "kode perusahaan"],
    ["1.A.1 Aset tetap", "INP.ANR.AT", "uraian_koreksi_dms", "keterangan"],
    ["1.A.1 Aset tetap", "INP.ANR.AT", "kategori_dms", "kelompok"],
    ["1.A.1 Aset tetap", "INP.ANR.AT", "perode_31_januari_2025", "saldo awal tahun"],
    ["catatan: baris ini diabaikan oleh aplikasi"],
]

# Data input dan export dibuat sengaja berbeda di beberapa baris, supaya hasil "Beda" terlihat.
BARIS_INPUT = [
    ["file contoh A", "2024-01", "KDA", "mobil", "kendaraan roda 4", 150300000, 150300000],
    ["file contoh B", "2024-01", "KDB", "mobil", "kendaraan roda 4", 126700000, 126700000],
    ["file contoh C", "2025-03", "KDC", "mobil", "kendaraan roda 4", 141000000, 141000000],
    ["file contoh D", "2025-06", 1000, "motor", "kendaraan roda 2", 47000000, 47000000],
    ["file contoh E", "2025-06", 2000, "motor", "kendaraan roda 2", 29500000, 29500000],
]
BARIS_EXPORT = [
    ["file contoh A", "2024-01", "KDA", "mobil", "kendaraan roda 4", 150300000],
    ["file contoh B", "2024-01", "KDX", "helikopter", "kendaraan roda 4", 99999],
    ["file contoh C", "2025-03", "KDC", "mobil", "kendaraan roda 4", 141000000],
    ["file contoh D", "2025-11", 1000, "motor", "kendaraan roda 2", 11111],
    ["file contoh E", "2025-06", 2000, "motor", "kendaraan roda 2", 29500000],
]


def _tulis_xlsx(path: Path, nama_sheet: str, header: list, baris: list) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = nama_sheet
    ws.append(header)
    for b in baris:
        ws.append(b)
    wb.save(path)


def _tulis_csv(path: Path, header: list, baris: list, pemisah: str = ";", encoding: str = "utf-8") -> None:
    with open(path, "w", newline="", encoding=encoding) as f:
        w = csv.writer(f, delimiter=pemisah)
        w.writerow(header)
        for b in baris:
            w.writerow(["" if v is None else v for v in b])


def main(folder: Path | None = None) -> Path:
    folder = Path(folder) if folder else Path(__file__).resolve().parent / "contoh"
    folder.mkdir(parents=True, exist_ok=True)

    # Satu pasangan
    _tulis_xlsx(folder / "mapping_contoh.xlsx", "Mapping", HEADER_MAPPING, MAPPING)
    _tulis_xlsx(folder / "1.A.1 Aset tetap.xlsx", "Input Awal", HEADER_INPUT, BARIS_INPUT)
    _tulis_xlsx(folder / "INP.ANR.AT.xlsx", "Export Module", HEADER_EXPORT, BARIS_EXPORT)

    # Banyak pasangan: input xlsx, export csv (pemisah titik koma, dan satu file cp1252)
    batch = folder / "batch"
    batch.mkdir(exist_ok=True)
    perusahaan = [
        ("PT Alfa", BARIS_INPUT, BARIS_EXPORT, ";", "utf-8"),
        ("PT Beta", BARIS_INPUT, [r[:6] for r in BARIS_INPUT], ";", "utf-8"),  # sengaja sama: hasilnya semua Sama
        ("PT Gamma", BARIS_INPUT, BARIS_EXPORT, ";", "cp1252"),
    ]
    for nama, baris_in, baris_ex, pemisah, enc in perusahaan:
        _tulis_xlsx(batch / f"1.A.1 Aset tetap - {nama}.xlsx", "Input Awal", HEADER_INPUT, baris_in)
        _tulis_csv(batch / f"1.A.1 Aset tetap - {nama}.csv", HEADER_EXPORT, baris_ex, pemisah, enc)
    return folder


if __name__ == "__main__":
    lokasi = main()
    print(f"File contoh dibuat di: {lokasi}")
