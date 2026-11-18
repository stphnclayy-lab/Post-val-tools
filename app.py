"""Validasi Export Modul: aplikasi Streamlit.

Lokal   : streamlit run app.py
Online  : deploy dari GitHub ke Streamlit Community Cloud

Mapping dan pengaturan kunci disimpan di sesi browser masing-masing, tidak di server.
Hasil hanya bisa diunduh (satu-satu atau sekaligus dalam zip).
"""
from __future__ import annotations

import hashlib
from collections import Counter
from pathlib import Path

import pandas as pd
import streamlit as st

from inti.batch import ItemBatch, buat_zip, export_dobel, jalankan_batch, tebak_entri, tebak_export
from inti.mapping import EntriMapping, baca_mapping_excel, gabung_mapping
from inti.proses import HasilProses

st.set_page_config(page_title="Validasi Export Modul", layout="wide")

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
TIPE_TABEL = ["xlsx", "csv"]
PILIH = "(pilih)"


def ambil_mapping() -> tuple[list[EntriMapping], str] | None:
    """Mapping aktif di sesi ini: (daftar entri, nama sumber file)."""
    return st.session_state.get("mapping")


def ambil_kunci() -> dict[str, list[str]]:
    """Pengaturan kunci per module di sesi ini."""
    return st.session_state.get("kunci", {})


def tabel_mapping(entri: list[EntriMapping]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Nama file": e.nama_file,
                "Nama module": e.nama_module,
                "Column file": p.kolom_file,
                "Column module": p.kolom_module or "(tidak dibandingkan)",
            }
            for e in entri
            for p in e.pasangan
        ]
    )


# ---------------------------------------------------------------- Menu 1: mapping
def halaman_mapping() -> None:
    st.title("1. Mapping kolom")
    st.write(
        "Upload satu atau banyak file mapping (format seperti sheet **Mapping**: baris 1 berisi header "
        "`nama file`, `nama module`, `column file`, `column module`). Semua file digabung jadi satu mapping. "
        "Mapping berlaku selama halaman ini dibuka. Kalau halaman di-refresh, upload ulang."
    )

    files = st.file_uploader(
        "File mapping (.xlsx, boleh banyak)", type=["xlsx"], accept_multiple_files=True, key="up_mapping"
    )
    if files:
        gab = gabung_mapping([(f.name, baca_mapping_excel(f.getvalue())) for f in files])
        for pesan in gab.error:
            st.error(pesan)
        for pesan in gab.peringatan:
            st.warning(pesan)
        if gab.entri and not gab.error:
            st.subheader(f"Pratinjau: {len(gab.entri)} nama file dari {len(files)} file mapping")
            st.dataframe(tabel_mapping(gab.entri), hide_index=True)
            if st.button("Pakai mapping ini", type="primary"):
                st.session_state["mapping"] = (gab.entri, ", ".join(f.name for f in files))
                st.success("Mapping dipakai. Lanjut ke menu 2 untuk mengatur kunci.")

    st.subheader("Mapping aktif")
    aktif = ambil_mapping()
    if aktif is None:
        st.info("Belum ada mapping. Upload file lalu klik **Pakai mapping ini**.")
    else:
        entri, sumber = aktif
        st.caption(f"Sumber: {sumber}")
        st.dataframe(tabel_mapping(entri), hide_index=True)


# ---------------------------------------------------------------- Menu 2: kunci
def halaman_kunci() -> None:
    st.title("2. Pengaturan kunci")
    st.write(
        "Pilih kolom yang dipakai sebagai **kunci** untuk mencocokkan baris input dengan baris "
        "export di tiap module. Kalau kunci dikosongkan, baris dipasangkan per posisi "
        "(baris 1 dengan baris 1, dan seterusnya)."
    )

    aktif = ambil_mapping()
    if aktif is None:
        st.warning("Pakai mapping dulu di menu 1.")
        return
    entri_list, _ = aktif
    kunci_tersimpan = ambil_kunci()

    kolom_per_module: dict[str, dict[str, str]] = {}  # module -> {kolom_file: label}
    for e in entri_list:
        peta = kolom_per_module.setdefault(e.nama_module, {})
        for p in e.pasangan_aktif:
            peta.setdefault(p.kolom_file, f"{p.kolom_file}  ↔  {p.kolom_module}")

    pilihan: dict[str, list[str]] = {}
    for module, peta in kolom_per_module.items():
        st.subheader(module)
        opsi = list(peta)
        default = [k for k in kunci_tersimpan.get(module, []) if k in peta]
        pilihan[module] = st.multiselect(
            "Kolom kunci",
            opsi,
            default=default,
            format_func=lambda k, peta=peta: peta[k],
            key=f"kunci_{module}",
        )

    if st.button("Simpan pengaturan kunci", type="primary"):
        st.session_state["kunci"] = pilihan
        st.success("Pengaturan kunci tersimpan untuk sesi ini.")


# ---------------------------------------------------------------- Menu 3: proses
def halaman_proses() -> None:
    st.title("3. Proses validasi")

    aktif = ambil_mapping()
    if aktif is None:
        st.warning("Belum ada mapping. Upload dulu di menu 1.")
        return
    entri_list, _ = aktif
    nama_entri = [e.nama_file for e in entri_list]
    peta_entri = {e.nama_file: e for e in entri_list}

    kiri, kanan = st.columns(2)
    with kiri:
        up_input = st.file_uploader(
            "File input awal (xlsx atau csv, boleh banyak)",
            type=TIPE_TABEL,
            accept_multiple_files=True,
            key="up_input",
        )
    with kanan:
        up_export = st.file_uploader(
            "File export module (xlsx atau csv, boleh banyak)",
            type=TIPE_TABEL,
            accept_multiple_files=True,
            key="up_export",
        )

    if not up_input or not up_export:
        st.info("Upload file input awal dan file export module untuk memulai.")
        return

    nama_in = [f.name for f in up_input]
    nama_ex = [f.name for f in up_export]
    if len(set(nama_in)) != len(nama_in) or len(set(nama_ex)) != len(nama_ex):
        st.error("Ada dua file dengan nama yang sama dalam satu kelompok upload. Ganti nama salah satunya.")
        return

    # Tebakan pasangan (bisa diubah di tabel)
    stems = [Path(n).stem for n in nama_in]
    tebak_mapping = [tebak_entri(s, nama_entri) for s in stems]
    modul_tebak = [peta_entri[t].nama_module if t else None for t in tebak_mapping]
    tebak_ex = tebak_export(stems, modul_tebak, nama_ex)

    df_awal = pd.DataFrame(
        {
            "File input": nama_in,
            "Nama file di mapping": [t or PILIH for t in tebak_mapping],
            "File export": [t or PILIH for t in tebak_ex],
        }
    )
    kunci_tabel = "pasangan_" + hashlib.md5("|".join(nama_in + nama_ex).encode()).hexdigest()[:10]

    st.subheader("Pasangan file")
    st.caption(
        "Tebakan otomatis: file input dicocokkan dengan 'nama file' di mapping, lalu file export "
        "dicari lewat nama module. Kalau ada yang salah, ubah langsung di tabel. "
        "File input yang belum punya pasangan akan dilewati."
    )
    edit = st.data_editor(
        df_awal,
        key=kunci_tabel,
        hide_index=True,
        disabled=["File input"],
        column_config={
            "Nama file di mapping": st.column_config.SelectboxColumn(options=[PILIH] + nama_entri, required=True),
            "File export": st.column_config.SelectboxColumn(options=[PILIH] + nama_ex, required=True),
        },
    )

    peta_ex = dict(zip(nama_ex, up_export))
    items: list[ItemBatch] = []
    dilewati: list[str] = []
    for i, (_, baris) in enumerate(edit.iterrows()):
        m, ex = baris["Nama file di mapping"], baris["File export"]
        if m == PILIH or ex == PILIH:
            dilewati.append(nama_in[i])
            continue
        items.append(
            ItemBatch(
                nama_input=nama_in[i],
                data_input=up_input[i].getvalue(),
                entri=peta_entri[m],
                nama_export=ex,
                data_export=peta_ex[ex].getvalue(),
            )
        )

    dobel = export_dobel(items)
    if dobel:
        st.error("File export ini dipasangkan ke lebih dari satu file input: " + ", ".join(dobel))
        return

    teks = f"**{len(items)}** pasangan siap diproses."
    if dilewati:
        teks += f" {len(dilewati)} file input belum punya pasangan dan akan dilewati."
    st.write(teks)

    if st.button("Proses semua", type="primary", disabled=not items):
        with st.spinner(f"Memproses {len(items)} pasangan..."):
            hasil = jalankan_batch(items, ambil_kunci())
        st.session_state["hasil_batch"] = {"hasil": hasil, "zip": buat_zip(hasil)}

    if st.session_state.get("hasil_batch"):
        tampilkan_hasil_batch(st.session_state["hasil_batch"])


def tampilkan_hasil_batch(data: dict) -> None:
    hasil: list[HasilProses] = data["hasil"]
    berhasil = [h for h in hasil if not h.error]

    st.header("Hasil")
    k = st.columns(3)
    k[0].metric("Diproses", len(hasil))
    k[1].metric("Berhasil", len(berhasil))
    k[2].metric("Gagal", len(hasil) - len(berhasil))
    st.caption("Hasil hanya tersedia untuk diunduh. Simpan zip-nya ke komputer kamu.")
    st.download_button(
        "Download semua hasil (.zip)",
        data=data["zip"],
        file_name="hasil_validasi.zip",
        mime="application/zip",
        key="dl_zip",
    )

    ringkasan = pd.DataFrame(
        [
            {
                "File input": h.nama_input,
                "Module": h.nama_module,
                "Status": (f"Gagal: {h.error[0]}" if h.error else ("OK, ada peringatan" if h.peringatan else "OK")),
                "Cocok": h.cocok,
                "Hanya di input": h.hanya_input,
                "Hanya di export": h.hanya_export,
                "Total Beda": sum(x["Jumlah Beda"] for x in h.beda_per_kolom),
                "File hasil": "-" if h.error else h.nama_file_output,
            }
            for h in hasil
        ]
    )
    st.dataframe(ringkasan, hide_index=True)

    if hasil:
        pilih = st.selectbox("Lihat detail file", [h.nama_input for h in hasil], key="detail_pilih")
        tampilkan_detail(next(h for h in hasil if h.nama_input == pilih))


def tampilkan_detail(h: HasilProses) -> None:
    for pesan in h.error:
        st.error(pesan)
    if h.error:
        return
    for pesan in h.peringatan:
        st.warning(pesan)

    label_cocok = "Cocok (per kunci)" if h.mode == "kunci" else "Cocok (per posisi)"
    if h.mode == "kunci":
        st.caption("Kunci: " + ", ".join(h.kunci))
    k = st.columns(5)
    k[0].metric("Baris input", h.jumlah_input)
    k[1].metric("Baris export", h.jumlah_export)
    k[2].metric(label_cocok, h.cocok)
    k[3].metric("Hanya di input", h.hanya_input)
    k[4].metric("Hanya di export", h.hanya_export)

    st.write("Jumlah **Beda** per kolom (perkiraan aplikasi):")
    st.dataframe(pd.DataFrame(h.beda_per_kolom), hide_index=True)
    st.caption("Angka ini hanya perkiraan. Sumber kebenarannya adalah kolom `check` di file Hasil (formula EXACT).")
    if h.kolom_dilewati:
        with st.expander("Kolom yang tidak dibandingkan"):
            for item in h.kolom_dilewati:
                st.write(f"- {item}")
    st.download_button(
        f"Download {h.nama_file_output}",
        data=h.xlsx,
        file_name=h.nama_file_output,
        mime=MIME_XLSX,
        key=f"dl_{h.nama_input}",
    )


# ---------------------------------------------------------------- navigasi
menu = st.sidebar.radio("Menu", ["1. Mapping", "2. Pengaturan kunci", "3. Proses validasi"])
if menu.startswith("1"):
    halaman_mapping()
elif menu.startswith("2"):
    halaman_kunci()
else:
    halaman_proses()
