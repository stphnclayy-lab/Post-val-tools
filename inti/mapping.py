"""Membaca file mapping. Formatnya sama dengan sheet 'Mapping':

    baris 1 : nama file | nama module | column file | column module
    baris 2+: pasangan kolom (column module boleh kosong = tidak dibandingkan)

Baris yang kolom 'column file'-nya kosong (misalnya catatan) diabaikan.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook

from .baca_excel import kosong

HEADER_WAJIB = ("nama file", "nama module", "column file", "column module")


@dataclass
class Pasangan:
    kolom_file: str
    kolom_module: str | None  # None = tidak dibandingkan


@dataclass
class EntriMapping:
    nama_file: str
    nama_module: str
    pasangan: list[Pasangan] = field(default_factory=list)

    @property
    def pasangan_aktif(self) -> list[Pasangan]:
        return [p for p in self.pasangan if p.kolom_module]


@dataclass
class HasilBacaMapping:
    entri: list[EntriMapping] = field(default_factory=list)
    error: list[str] = field(default_factory=list)
    peringatan: list[str] = field(default_factory=list)


def normalisasi_nama_file(nama: str) -> str:
    """Buang spasi di tepi dan akhiran .xlsx, supaya 'X.xlsx' dan 'X' dianggap sama."""
    nama = nama.strip()
    if nama.lower().endswith(".xlsx"):
        nama = nama[: -len(".xlsx")]
    return nama.strip()


def baca_mapping_excel(sumber: bytes | str | Path) -> HasilBacaMapping:
    hasil = HasilBacaMapping()
    if isinstance(sumber, bytes):
        sumber = BytesIO(sumber)

    wb = load_workbook(sumber, read_only=True, data_only=True)
    try:
        semua = wb.worksheets[0].iter_rows(values_only=True)
        header = ["" if kosong(v) else str(v).strip().lower() for v in next(semua, ())]

        hilang = [h for h in HEADER_WAJIB if h not in header]
        if hilang:
            hasil.error.append("Header berikut tidak ditemukan di baris 1: " + ", ".join(hilang))
            return hasil
        idx = {h: header.index(h) for h in HEADER_WAJIB}

        def ambil(baris, nama: str) -> str | None:
            i = idx[nama]
            v = baris[i] if i < len(baris) else None
            return None if kosong(v) else str(v).strip()

        urutan: dict[str, EntriMapping] = {}
        for nomor, baris in enumerate(semua, start=2):
            nf = ambil(baris, "nama file")
            nm = ambil(baris, "nama module")
            kf = ambil(baris, "column file")
            km = ambil(baris, "column module")

            if kf is None and nm is None:
                continue  # baris kosong atau catatan
            if kf is None:
                hasil.peringatan.append(f"Baris {nomor} dilewati: 'column file' kosong.")
                continue
            if nf is None or nm is None:
                hasil.error.append(f"Baris {nomor}: 'nama file' atau 'nama module' kosong.")
                continue

            nf = normalisasi_nama_file(nf)
            entri = urutan.get(nf)
            if entri is None:
                entri = urutan[nf] = EntriMapping(nama_file=nf, nama_module=nm)
            elif entri.nama_module != nm:
                hasil.error.append(
                    f"Baris {nomor}: '{nf}' dipetakan ke dua module berbeda "
                    f"('{entri.nama_module}' dan '{nm}')."
                )
                continue

            if any(p.kolom_file == kf for p in entri.pasangan):
                hasil.error.append(f"Baris {nomor}: kolom '{kf}' dobel untuk '{nf}'.")
                continue
            entri.pasangan.append(Pasangan(kolom_file=kf, kolom_module=km))
    finally:
        wb.close()

    hasil.entri = list(urutan.values())
    for e in hasil.entri:
        if not e.pasangan_aktif:
            hasil.peringatan.append(
                f"'{e.nama_file}' tidak punya pasangan 'column module', jadi tidak akan dibandingkan."
            )
    return hasil


def gabung_mapping(daftar: list[tuple[str, HasilBacaMapping]]) -> HasilBacaMapping:
    """Gabungkan beberapa file mapping jadi satu. Pasangan yang sama persis cukup dicatat sekali."""
    hasil = HasilBacaMapping()
    peta: dict[str, EntriMapping] = {}
    for nama, h in daftar:
        hasil.error.extend(f"{nama}: {pesan}" for pesan in h.error)
        hasil.peringatan.extend(f"{nama}: {pesan}" for pesan in h.peringatan)
        for e in h.entri:
            tujuan = peta.get(e.nama_file)
            if tujuan is None:
                peta[e.nama_file] = EntriMapping(
                    nama_file=e.nama_file,
                    nama_module=e.nama_module,
                    pasangan=[Pasangan(p.kolom_file, p.kolom_module) for p in e.pasangan],
                )
                continue
            if tujuan.nama_module != e.nama_module:
                hasil.error.append(
                    f"{nama}: '{e.nama_file}' dipetakan ke module berbeda "
                    f"('{tujuan.nama_module}' dan '{e.nama_module}')."
                )
                continue
            for p in e.pasangan:
                ada = next((x for x in tujuan.pasangan if x.kolom_file == p.kolom_file), None)
                if ada is None:
                    tujuan.pasangan.append(Pasangan(p.kolom_file, p.kolom_module))
                elif ada.kolom_module != p.kolom_module:
                    hasil.error.append(
                        f"{nama}: kolom '{p.kolom_file}' di '{e.nama_file}' punya dua pasangan berbeda."
                    )
    hasil.entri = list(peta.values())
    return hasil
