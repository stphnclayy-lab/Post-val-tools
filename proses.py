"""Orkestrasi satu pasangan (file input + file export) menjadi satu file Hasil."""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from typing import Sequence

from .baca_excel import DataExcel
from .mapping import EntriMapping
from .pencocokan import hitung_beda, pasangkan
from .tulis_hasil import buat_xlsx_hasil


@dataclass
class HasilProses:
    nama_module: str
    nama_file_output: str
    nama_input: str = ""
    error: list[str] = field(default_factory=list)
    peringatan: list[str] = field(default_factory=list)
    kolom_dilewati: list[str] = field(default_factory=list)
    mode: str = "posisi"  # "kunci" atau "posisi"
    kunci: list[str] = field(default_factory=list)
    jumlah_input: int = 0
    jumlah_export: int = 0
    cocok: int = 0
    hanya_input: int = 0
    hanya_export: int = 0
    kunci_ganda: int = 0
    beda_per_kolom: list[dict] = field(default_factory=list)
    xlsx: bytes | None = None
    lokasi_simpan: str | None = None


def _header_dobel(header: Sequence[str], nama: Sequence[str]) -> list[str]:
    hitung = Counter(h for h in header if h)
    return sorted({n for n in nama if hitung[n] > 1})


def proses(
    entri: EntriMapping,
    kunci: Sequence[str],
    data_input: DataExcel,
    data_export: DataExcel,
    nama_output: str | None = None,
) -> HasilProses:
    h = HasilProses(
        nama_module=entri.nama_module,
        nama_file_output=nama_output or f"Hasil_{entri.nama_module}.xlsx",
        nama_input=data_input.nama_file,
        kunci=list(kunci),
    )

    aktif = entri.pasangan_aktif
    if not aktif:
        h.error.append("Mapping module ini tidak punya pasangan 'column module'.")
        return h

    # 1. Cek kolom yang dibutuhkan ada di kedua file
    pos_in: list[int | None] = []
    pos_ex: list[int | None] = []
    for p in aktif:
        i = data_input.posisi(p.kolom_file)
        j = data_export.posisi(p.kolom_module)
        if i is None:
            h.error.append(f"Kolom input '{p.kolom_file}' ada di mapping, tapi tidak ada di file input.")
        if j is None:
            h.error.append(f"Kolom module '{p.kolom_module}' ada di mapping, tapi tidak ada di file export.")
        pos_in.append(i)
        pos_ex.append(j)

    for nama in _header_dobel(data_input.header, [p.kolom_file for p in aktif]):
        h.error.append(f"Header '{nama}' muncul lebih dari sekali di file input.")
    for nama in _header_dobel(data_export.header, [p.kolom_module for p in aktif]):
        h.error.append(f"Header '{nama}' muncul lebih dari sekali di file export.")

    # 2. Cek kunci masih ada di mapping aktif
    kolom_file_aktif = [p.kolom_file for p in aktif]
    for k in kunci:
        if k not in kolom_file_aktif:
            h.error.append(f"Kunci '{k}' tidak ada di mapping aktif. Atur ulang di menu 2.")

    if h.error:
        return h

    # 3. Info kolom yang tidak ikut dibandingkan
    mapped_input = {p.kolom_file for p in entri.pasangan}
    mapped_export = {p.kolom_module for p in entri.pasangan if p.kolom_module}
    for p in entri.pasangan:
        if not p.kolom_module:
            h.kolom_dilewati.append(f"input '{p.kolom_file}' (tanpa pasangan di module)")
    for nama in data_input.header:
        if nama and nama not in mapped_input:
            h.kolom_dilewati.append(f"input '{nama}' (tidak ada di mapping)")
    for nama in data_export.header:
        if nama and nama not in mapped_export:
            h.kolom_dilewati.append(f"export '{nama}' (tidak ada di mapping)")

    # 4. Kunci pencocokan (urut sesuai daftar kunci yang dipilih)
    indeks_kunci = [j for k in kunci for j, p in enumerate(aktif) if p.kolom_file == k]
    if indeks_kunci:
        h.mode = "kunci"
    else:
        h.mode = "posisi"
        h.peringatan.append(
            "Belum ada kunci untuk module ini, baris dipasangkan per posisi. Atur kunci di menu 2."
        )

    # 5. Susun data dan pasangkan
    baris_in = [[row[i] for i in pos_in] for row in data_input.baris]
    baris_ex = [[row[j] for j in pos_ex] for row in data_export.baris]
    res = pasangkan(baris_in, baris_ex, indeks_kunci)

    pasangan_nama = [(p.kolom_file, p.kolom_module) for p in aktif]
    beda = hitung_beda(res.baris, len(aktif))

    h.jumlah_input = len(baris_in)
    h.jumlah_export = len(baris_ex)
    h.cocok = res.cocok
    h.hanya_input = res.hanya_input
    h.hanya_export = res.hanya_export
    h.kunci_ganda = res.kunci_ganda
    if res.kunci_ganda:
        h.peringatan.append(
            f"{res.kunci_ganda} nilai kunci muncul lebih dari sekali. "
            "Pasangannya diurutkan: kemunculan ke-1 dengan ke-1, dan seterusnya."
        )
    h.beda_per_kolom = [
        {"Kolom input": kf, "Kolom module": km, "Jumlah Beda": n}
        for (kf, km), n in zip(pasangan_nama, beda)
    ]
    h.xlsx = buat_xlsx_hasil(pasangan_nama, res.baris)
    return h
