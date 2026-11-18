"""Pencocokan baris input dengan baris export, serta peniruan perbandingan EXACT Excel.

Setiap baris direpresentasikan sebagai list nilai, urutannya mengikuti daftar pasangan kolom.
Hasil pasangan berupa tuple (baris_input | None, baris_export | None).
"""
from __future__ import annotations

from collections import Counter, deque
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Sequence

EPOCH = datetime(1899, 12, 30)


def teks_excel(nilai: Any) -> str:
    """Meniru cara EXACT di Excel mengubah sebuah nilai menjadi teks sebelum dibandingkan.

    Catatan: angka 1000 dan teks "1000" dianggap sama oleh EXACT, karena keduanya jadi "1000".
    Tanggal dibandingkan sebagai nomor serialnya (misalnya 45658), bukan sebagai tampilan tanggal.
    """
    if nilai is None:
        return ""
    if isinstance(nilai, bool):
        return "TRUE" if nilai else "FALSE"
    if isinstance(nilai, datetime):
        selisih = nilai - EPOCH
        nilai = selisih.days + selisih.seconds / 86400 + selisih.microseconds / 86_400_000_000
    elif isinstance(nilai, date):
        nilai = (nilai - EPOCH.date()).days
    if isinstance(nilai, int):
        return str(nilai)
    if isinstance(nilai, float):
        if nilai.is_integer() and abs(nilai) < 1e15:
            return str(int(nilai))
        return format(nilai, ".15g")
    return str(nilai)


def sama_exact(a: Any, b: Any) -> bool:
    return teks_excel(a) == teks_excel(b)


@dataclass
class HasilPasangan:
    baris: list[tuple[list | None, list | None]]
    cocok: int
    hanya_input: int
    hanya_export: int
    kunci_ganda: int


def _bungkus(baris: list, kunci_ganda: int) -> HasilPasangan:
    return HasilPasangan(
        baris=baris,
        cocok=sum(1 for bi, be in baris if bi is not None and be is not None),
        hanya_input=sum(1 for _, be in baris if be is None),
        hanya_export=sum(1 for bi, _ in baris if bi is None),
        kunci_ganda=kunci_ganda,
    )


def pasangkan(
    baris_input: Sequence[list],
    baris_export: Sequence[list],
    indeks_kunci: Sequence[int],
) -> HasilPasangan:
    """Pasangkan baris input dengan baris export.

    - indeks_kunci kosong  : pasangkan per posisi (baris ke-1 dengan ke-1, dst).
    - indeks_kunci terisi  : pasangkan berdasarkan nilai kunci. Jika kunci muncul berkali-kali,
                             kemunculan ke-n di input dipasangkan dengan kemunculan ke-n di export.
    Baris export yang tidak punya pasangan ditaruh di akhir, dengan sisi input kosong.
    """
    if not indeks_kunci:
        n = max(len(baris_input), len(baris_export))
        baris = [
            (
                baris_input[i] if i < len(baris_input) else None,
                baris_export[i] if i < len(baris_export) else None,
            )
            for i in range(n)
        ]
        return _bungkus(baris, kunci_ganda=0)

    kunci_export = [tuple(teks_excel(be[j]) for j in indeks_kunci) for be in baris_export]
    antrean: dict[tuple, deque[int]] = {}
    for idx, k in enumerate(kunci_export):
        antrean.setdefault(k, deque()).append(idx)

    jumlah_export = Counter(kunci_export)
    jumlah_input: Counter = Counter()
    dipakai = [False] * len(baris_export)
    baris: list = []

    for bi in baris_input:
        k = tuple(teks_excel(bi[j]) for j in indeks_kunci)
        jumlah_input[k] += 1
        q = antrean.get(k)
        if q:
            idx = q.popleft()
            dipakai[idx] = True
            baris.append((bi, baris_export[idx]))
        else:
            baris.append((bi, None))

    for idx, be in enumerate(baris_export):
        if not dipakai[idx]:
            baris.append((None, be))

    kunci_ganda = sum(
        1
        for k in set(jumlah_input) | set(jumlah_export)
        if jumlah_input[k] > 1 or jumlah_export[k] > 1
    )
    return _bungkus(baris, kunci_ganda)


def hitung_beda(baris: Sequence[tuple[list | None, list | None]], jumlah_kolom: int) -> list[int]:
    """Perkiraan jumlah 'Beda' per pasangan kolom (dipakai untuk ringkasan di aplikasi)."""
    jumlah = [0] * jumlah_kolom
    for bi, be in baris:
        for j in range(jumlah_kolom):
            a = bi[j] if bi is not None else None
            b = be[j] if be is not None else None
            if not sama_exact(a, b):
                jumlah[j] += 1
    return jumlah
