"""Memproses banyak pasangan sekaligus. Tiap pasangan menghasilkan satu file Hasil terpisah."""
from __future__ import annotations

import io
import re
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from .baca_excel import baca_tabel
from .mapping import EntriMapping
from .proses import HasilProses, proses


@dataclass
class ItemBatch:
    nama_input: str
    data_input: bytes
    entri: EntriMapping
    nama_export: str
    data_export: bytes


def _nama_unik(nama: str, terpakai: set[str]) -> str:
    """Hasil_X.xlsx, lalu Hasil_X_2.xlsx, dst. jika nama sudah dipakai di folder yang sama."""
    dasar, ekstensi = nama.rsplit(".", 1)
    kandidat = nama
    n = 2
    while kandidat.lower() in terpakai:
        kandidat = f"{dasar}_{n}.{ekstensi}"
        n += 1
    terpakai.add(kandidat.lower())
    return kandidat


def jalankan_batch(
    items: list[ItemBatch], kunci_map: dict[str, list[str]], folder: Path | None = None
) -> list[HasilProses]:
    """Jika folder diisi, hasil juga disimpan sebagai file di sana. Jika None, hanya disimpan di memori."""
    if folder is not None:
        folder.mkdir(parents=True, exist_ok=True)
    terpakai: set[str] = set()
    hasil: list[HasilProses] = []
    for it in items:
        din = baca_tabel(it.data_input, it.nama_input)
        dex = baca_tabel(it.data_export, it.nama_export)
        nama_out = _nama_unik(f"Hasil_{Path(it.nama_input).stem}.xlsx", terpakai)
        h = proses(it.entri, kunci_map.get(it.entri.nama_module, []), din, dex, nama_output=nama_out)
        if folder is not None and h.xlsx is not None and not h.error:
            lokasi = folder / nama_out
            lokasi.write_bytes(h.xlsx)
            h.lokasi_simpan = str(lokasi)
        hasil.append(h)
    return hasil


def buat_zip(hasil: list[HasilProses]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for h in hasil:
            if h.xlsx is not None and not h.error:
                z.writestr(h.nama_file_output, h.xlsx)
    return buf.getvalue()


def _alnum(teks: str) -> str:
    """Hanya huruf dan angka, huruf kecil. Jadi 'INP.PPE.030 Aset' dan 'INP_PPE_030_Aset' dianggap sama."""
    return re.sub(r"[\W_]+", "", teks.lower())


def tebak_entri(stem: str, nama_entri: list[str]) -> str | None:
    """Cocokkan nama file input dengan 'nama file' di mapping.
    Nama yang sama (abaikan tanda baca) diutamakan. Kalau tidak ada, pakai nama mapping
    yang paling panjang dan terkandung di nama file input."""
    s = _alnum(stem)
    if not s:
        return None
    for n in nama_entri:
        if _alnum(n) == s:
            return n
    kandidat = [n for n in nama_entri if _alnum(n) and _alnum(n) in s]
    return max(kandidat, key=lambda n: len(_alnum(n))) if kandidat else None


def tebak_export(stems: list[str], modules: list[str | None], nama_export: list[str]) -> list[str | None]:
    """Tebak file export untuk tiap file input. Urutannya:
    1) nama file export sama dengan nama file input,
    2) nama file export sama dengan nama module di mapping,
    3) nama module terkandung di nama file export, dan hanya satu file yang cocok.
    Selain itu dibiarkan kosong supaya dipilih manual di tabel."""
    per_stem = {_alnum(Path(n).stem): n for n in nama_export}
    hasil: list[str | None] = [None] * len(stems)
    terpakai: set[str] = set()

    for i, s in enumerate(stems):
        n = per_stem.get(_alnum(s))
        if n and n not in terpakai:
            hasil[i] = n
            terpakai.add(n)

    for i, m in enumerate(modules):
        if hasil[i] or not m:
            continue
        mk = _alnum(m)
        tersedia = [n for n in nama_export if n not in terpakai]
        sama = [n for n in tersedia if _alnum(Path(n).stem) == mk]
        kandidat = sama if len(sama) == 1 else [n for n in tersedia if mk and mk in _alnum(Path(n).stem)]
        if len(kandidat) == 1:
            hasil[i] = kandidat[0]
            terpakai.add(kandidat[0])
    return hasil


def export_dobel(items: list[ItemBatch]) -> list[str]:
    """Nama file export yang dipasangkan ke lebih dari satu file input."""
    hitung = Counter(it.nama_export for it in items)
    return sorted(n for n, c in hitung.items() if c > 1)
