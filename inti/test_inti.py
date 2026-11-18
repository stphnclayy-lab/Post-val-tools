"""Tes untuk logika inti. Jalankan dari folder project:  python -m pytest -q"""
from __future__ import annotations

import io
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import buat_contoh  # noqa: E402
from inti.baca_excel import baca_csv, baca_excel, baca_tabel  # noqa: E402
from inti.batch import ItemBatch, buat_zip, export_dobel, jalankan_batch, tebak_entri, tebak_export  # noqa: E402
from inti.mapping import baca_mapping_excel, gabung_mapping  # noqa: E402
from inti.pencocokan import hitung_beda, pasangkan, sama_exact, teks_excel  # noqa: E402
from inti.proses import proses  # noqa: E402


def _xlsx_bytes(header, baris, teks_saja: bool = False) -> bytes:
    """Buat file xlsx di memori. teks_saja=True menyimpan semua nilai string sebagai teks, bukan formula."""
    wb = Workbook()
    ws = wb.active
    ws.append(header)
    for b in baris:
        ws.append(b)
    if teks_saja:
        for row in ws.iter_rows():
            for sel in row:
                if isinstance(sel.value, str):
                    sel.data_type = "s"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- perbandingan EXACT
def test_exact_meniru_excel():
    assert sama_exact(1000, "1000")  # EXACT mengubah angka jadi teks dulu
    assert sama_exact(None, "")
    assert sama_exact(True, "TRUE")
    assert not sama_exact("abc", "ABC")  # case-sensitive
    assert not sama_exact("a ", "a")  # spasi ikut dihitung
    assert teks_excel(datetime(2025, 1, 1)) == "45658"  # tanggal = nomor serial
    assert teks_excel(1.0) == "1"
    assert teks_excel(0.1 + 0.2) == "0.3"


# ---------------------------------------------------------------- pencocokan baris
def test_pasangkan_per_posisi():
    bi = [["a", 1], ["b", 2], ["c", 3]]
    be = [["a", 1], ["x", 2]]
    res = pasangkan(bi, be, [])
    assert res.cocok == 2
    assert res.hanya_input == 1
    assert res.hanya_export == 0


def test_pasangkan_kunci_dengan_duplikat():
    bi = [["K1", 10], ["K2", 20], ["K1", 11]]
    be = [["K2", 20], ["K1", 10], ["K3", 30], ["K1", 99]]
    res = pasangkan(bi, be, [0])
    assert res.baris[0] == (["K1", 10], ["K1", 10])
    assert res.baris[1] == (["K2", 20], ["K2", 20])
    assert res.baris[2] == (["K1", 11], ["K1", 99])
    assert res.baris[3] == (None, ["K3", 30])
    assert res.cocok == 3
    assert res.hanya_input == 0
    assert res.hanya_export == 1
    assert res.kunci_ganda == 1


def test_hitung_beda_sederhana():
    baris = [([1, "a"], [1, "a"]), ([2, "b"], [3, "b"]), (None, [4, "c"])]
    # kolom 0: baris 2 beda, baris 3 beda (kosong vs 4) -> 2
    # kolom 1: baris 3 beda (kosong vs "c") -> 1
    assert hitung_beda(baris, 2) == [2, 1]


# ---------------------------------------------------------------- membaca mapping
def test_mapping_abaikan_catatan_dan_tanpa_pasangan():
    data = _xlsx_bytes(
        ["nama file", "nama module", "column file", "column module"],
        [
            ["1.A.1 Aset tetap.xlsx", "INP.ANR.AT", "file_path", None],
            ["1.A.1 Aset tetap", "INP.ANR.AT", "period_koreksi_dms", "time"],
            [None, None, None, None],
            ["catatan panjang di kolom A", None, None, None],
        ],
    )
    hasil = baca_mapping_excel(data)
    assert not hasil.error
    assert len(hasil.entri) == 1
    e = hasil.entri[0]
    assert e.nama_file == "1.A.1 Aset tetap"
    assert [p.kolom_module for p in e.pasangan] == [None, "time"]
    assert len(e.pasangan_aktif) == 1


def test_mapping_header_kurang_ditolak():
    assert baca_mapping_excel(_xlsx_bytes(["nama file", "nama module"], [["x", "y"]])).error


def test_gabung_beberapa_file_mapping():
    h1 = baca_mapping_excel(
        _xlsx_bytes(
            ["nama file", "nama module", "column file", "column module"],
            [["A", "M1", "k", "kode"], ["A", "M1", "x", "xx"]],
        )
    )
    h2 = baca_mapping_excel(
        _xlsx_bytes(
            ["nama file", "nama module", "column file", "column module"],
            [["A", "M1", "k", "kode"], ["B", "M2", "y", "yy"]],  # baris 'k' sama persis: tidak dobel
        )
    )
    gab = gabung_mapping([("f1.xlsx", h1), ("f2.xlsx", h2)])
    assert not gab.error
    peta = {e.nama_file: e for e in gab.entri}
    assert [p.kolom_file for p in peta["A"].pasangan] == ["k", "x"]
    assert peta["B"].nama_module == "M2"


def test_gabung_mapping_module_bentrok_ditolak():
    h1 = baca_mapping_excel(_xlsx_bytes(["nama file", "nama module", "column file", "column module"], [["A", "M1", "k", "k"]]))
    h2 = baca_mapping_excel(_xlsx_bytes(["nama file", "nama module", "column file", "column module"], [["A", "M2", "k", "k"]]))
    assert gabung_mapping([("f1", h1), ("f2", h2)]).error


# ---------------------------------------------------------------- membaca CSV
def test_baca_csv_titik_koma():
    data = "nama;kota\nBudi;Bandung\nSiti;Jakarta Selatan\n".encode("utf-8")
    d = baca_csv(data, "x.csv")
    assert d.header == ["nama", "kota"]
    assert d.baris == [["Budi", "Bandung"], ["Siti", "Jakarta Selatan"]]


def test_baca_csv_koma_dengan_bom_dan_kosong_jadi_none():
    data = "﻿kode,nilai\nA,1000\nB,\n".encode("utf-8")
    d = baca_csv(data, "y.csv")
    assert d.header == ["kode", "nilai"]
    assert d.baris == [["A", "1000"], ["B", None]]


def test_baca_csv_cp1252():
    data = "kode;keterangan\nA;Gedung Ö\n".encode("cp1252")
    d = baca_csv(data, "z.csv")
    assert d.baris[0][1] == "Gedung Ö"


def test_baca_tabel_pilih_sesuai_ekstensi():
    csv_data = "a;b\n1;2\n".encode("utf-8")
    assert baca_tabel(csv_data, "f.csv").header == ["a", "b"]
    xlsx_data = _xlsx_bytes(["a", "b"], [[1, 2]])
    assert baca_tabel(xlsx_data, "f.xlsx").baris == [[1, 2]]


def test_proses_csv_hasil_sama_dengan_xlsx(tmp_path):
    folder = buat_contoh.main(tmp_path)
    entri = baca_mapping_excel(folder / "mapping_contoh.xlsx").entri[0]
    din = baca_excel(folder / "1.A.1 Aset tetap.xlsx", "in.xlsx")
    dex_xlsx = baca_excel(folder / "INP.ANR.AT.xlsx", "ex.xlsx")
    # Versi CSV dari export yang sama
    dex_csv = baca_csv((folder / "batch" / "1.A.1 Aset tetap - PT Alfa.csv").read_bytes(), "ex.csv")

    h1 = proses(entri, ["company_code_dms"], din, dex_xlsx)
    h2 = proses(entri, ["company_code_dms"], din, dex_csv)
    assert [x["Jumlah Beda"] for x in h1.beda_per_kolom] == [x["Jumlah Beda"] for x in h2.beda_per_kolom]
    assert h1.cocok == h2.cocok == 4


# ---------------------------------------------------------------- proses end-to-end
def test_proses_contoh_dengan_kunci(tmp_path):
    folder = buat_contoh.main(tmp_path)
    entri = baca_mapping_excel(folder / "mapping_contoh.xlsx").entri[0]
    din = baca_excel(folder / "1.A.1 Aset tetap.xlsx", "1.A.1 Aset tetap.xlsx")
    dex = baca_excel(folder / "INP.ANR.AT.xlsx", "INP.ANR.AT.xlsx")

    h = proses(entri, ["company_code_dms"], din, dex)

    assert not h.error
    assert h.mode == "kunci"
    assert h.nama_input == "1.A.1 Aset tetap.xlsx"
    assert h.cocok == 4  # KDA, KDC, 1000, 2000
    assert h.hanya_input == 1  # KDB
    assert h.hanya_export == 1  # KDX
    assert [x["Jumlah Beda"] for x in h.beda_per_kolom] == [3, 2, 2, 2, 3]
    assert "total_dms" in " ".join(h.kolom_dilewati)
    assert "file_path" in " ".join(h.kolom_dilewati)

    ws = load_workbook(io.BytesIO(h.xlsx))["Hasil"]
    assert ws["C2"].value == '=IF(EXACT(A2,B2),"Sama","Beda")'
    assert ws["O2"].value == '=IF(EXACT(M2,N2),"Sama","Beda")'
    assert ws.max_column == 15  # 5 pasangan x 3 kolom


def test_proses_tanpa_kunci_pakai_posisi(tmp_path):
    folder = buat_contoh.main(tmp_path)
    entri = baca_mapping_excel(folder / "mapping_contoh.xlsx").entri[0]
    din = baca_excel(folder / "1.A.1 Aset tetap.xlsx", "x.xlsx")
    dex = baca_excel(folder / "INP.ANR.AT.xlsx", "y.xlsx")

    h = proses(entri, [], din, dex)

    assert not h.error
    assert h.mode == "posisi"
    assert h.cocok == 5
    assert h.peringatan


def test_proses_kolom_hilang_di_file_ditolak(tmp_path):
    folder = buat_contoh.main(tmp_path)
    entri = baca_mapping_excel(folder / "mapping_contoh.xlsx").entri[0]
    rusak = tmp_path / "rusak.xlsx"
    rusak.write_bytes(_xlsx_bytes(["file_path", "time"], [["a", "2024-01"]]))
    din = baca_excel(folder / "1.A.1 Aset tetap.xlsx", "x.xlsx")
    dex = baca_excel(rusak, "rusak.xlsx")

    h = proses(entri, [], din, dex)

    assert h.error
    assert h.xlsx is None


def test_teks_berawalan_tanda_sama_dengan_tidak_jadi_formula():
    entri = baca_mapping_excel(
        _xlsx_bytes(["nama file", "nama module", "column file", "column module"], [["n", "M1", "kode", "kode"]])
    ).entri[0]
    din = baca_excel(_xlsx_bytes(["kode"], [["=SUM(1,2)"]], teks_saja=True), "n.xlsx")
    dex = baca_excel(_xlsx_bytes(["kode"], [["=SUM(1,2)"]], teks_saja=True), "M1.xlsx")

    h = proses(entri, [], din, dex)
    ws = load_workbook(io.BytesIO(h.xlsx))["Hasil"]
    assert ws["A2"].value == "=SUM(1,2)"
    assert ws["A2"].data_type == "s"
    assert ws["C2"].value == '=IF(EXACT(A2,B2),"Sama","Beda")'


# ---------------------------------------------------------------- batch
def test_batch_banyak_file_nama_unik_dan_zip(tmp_path):
    folder = buat_contoh.main(tmp_path)
    entri = baca_mapping_excel(folder / "mapping_contoh.xlsx").entri[0]
    b = folder / "batch"
    items = []
    for nama in ["PT Alfa", "PT Beta", "PT Gamma"]:
        items.append(
            ItemBatch(
                nama_input=f"1.A.1 Aset tetap - {nama}.xlsx",
                data_input=(b / f"1.A.1 Aset tetap - {nama}.xlsx").read_bytes(),
                entri=entri,
                nama_export=f"1.A.1 Aset tetap - {nama}.csv",
                data_export=(b / f"1.A.1 Aset tetap - {nama}.csv").read_bytes(),
            )
        )
    # Satu input sengaja dobel nama dasarnya, supaya nama output dibuat unik
    items.append(items[0])

    out = tmp_path / "hasil_uji"
    hasil = jalankan_batch(items, {"INP.ANR.AT": ["company_code_dms"]}, out)

    assert all(not h.error for h in hasil)
    nama_file = sorted(p.name for p in out.iterdir())
    assert nama_file == sorted(
        [
            "Hasil_1.A.1 Aset tetap - PT Alfa.xlsx",
            "Hasil_1.A.1 Aset tetap - PT Alfa_2.xlsx",
            "Hasil_1.A.1 Aset tetap - PT Beta.xlsx",
            "Hasil_1.A.1 Aset tetap - PT Gamma.xlsx",
        ]
    )
    # PT Beta sengaja sama persis: semua kolom Sama
    beta = next(h for h in hasil if "Beta" in h.nama_input)
    assert sum(x["Jumlah Beda"] for x in beta.beda_per_kolom) == 0

    import zipfile

    with zipfile.ZipFile(io.BytesIO(buat_zip(hasil))) as z:
        assert len(z.namelist()) == 4


def test_tebak_entri_dan_export():
    nama_entri = ["1.A.1 Aset tetap", "1.A.2 Lainnya"]
    assert tebak_entri("1.A.1 Aset tetap", nama_entri) == "1.A.1 Aset tetap"
    assert tebak_entri("1.A.1 Aset tetap - PT Alfa", nama_entri) == "1.A.1 Aset tetap"
    assert tebak_entri("tidak cocok", nama_entri) is None

    stems = ["X - A", "X - B"]
    exports = ["X - A.csv", "X - B.csv", "INP.ANR.AT.xlsx"]
    assert tebak_export(stems, [None, None], exports) == ["X - A.csv", "X - B.csv"]

    # Tanpa nama sama: hanya satu export yang mengandung nama module -> dipakai
    assert tebak_export(["satu"], ["INP.ANR.AT"], ["INP.ANR.AT.xlsx"]) == ["INP.ANR.AT.xlsx"]
    # Dua export mengandung module -> ambigu, dibiarkan kosong
    assert tebak_export(["satu"], ["INP.ANR.AT"], ["INP.ANR.AT 1.xlsx", "INP.ANR.AT 2.xlsx"]) == [None]


def test_pasangan_lewat_mapping_nama_tidak_sama():
    nama_mapping = ["01.A.1_Rincian_Aset_Tetap_per_Fungsi_DM(1)"]
    # Windows menambah (2) saat file diunduh ulang, nama tetap dikenali
    assert tebak_entri("01.A.1_Rincian_Aset_Tetap_per_Fungsi_DM(1) (2)", nama_mapping) == nama_mapping[0]

    modul = "INP.PPE.030 Aset Tetap Fungsi"
    exports = ["INP_PPE_030_Aset_Tetap_Fungsi.xlsx", "lainnya.csv"]
    assert tebak_export(["x"], [modul], exports) == ["INP_PPE_030_Aset_Tetap_Fungsi.xlsx"]

    # Nama module yang persis sama diutamakan daripada yang hanya mengandung
    assert tebak_export(["x"], ["INP.ANR.AT"], ["INP.ANR.AT Lain.xlsx", "INP.ANR.AT.xlsx"]) == ["INP.ANR.AT.xlsx"]


def test_export_dobel_terdeteksi():
    e = baca_mapping_excel(_xlsx_bytes(["nama file", "nama module", "column file", "column module"], [["A", "M", "k", "k"]])).entri[0]
    items = [
        ItemBatch("a.xlsx", b"", e, "ex.csv", b""),
        ItemBatch("b.xlsx", b"", e, "ex.csv", b""),
    ]
    assert export_dobel(items) == ["ex.csv"]


def test_aplikasi_bisa_dijalankan():
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(str(ROOT / "app.py")).run(timeout=60)
    assert not at.exception
