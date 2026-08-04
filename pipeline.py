from __future__ import annotations

from typing import Any, List

import pandas as pd


def compute_method_order(total_ayat: int, metode: str):
    """Return the shifted row order used for the current method."""
    if total_ayat % 2 == 0:
        center_pos = total_ayat // 2 - 1
    else:
        center_pos = total_ayat // 2

    if metode == "T111":
        target_f_idx = center_pos
    elif metode == "T222":
        target_f_idx = 0
    elif metode == "T333":
        if total_ayat % 2 == 0:
            target_f_idx = total_ayat - 2
        else:
            target_f_idx = total_ayat - 1
    else:
        target_f_idx = center_pos

    start_idx = (target_f_idx - center_pos) % total_ayat
    new_order = [(start_idx + i) % total_ayat for i in range(total_ayat)]
    return center_pos, target_f_idx, start_idx, new_order


def build_staging_met(staging_data: pd.DataFrame | None, metode: str):
    """Build the data ordered according to the selected method."""
    if staging_data is None:
        raise ValueError("staging_data tidak boleh None sebelum menetapkan metode.")

    total_ayat = len(staging_data)
    center_pos, _, _, new_order = compute_method_order(total_ayat, metode)
    staging_data_met = staging_data.iloc[new_order].copy()
    staging_data_met.reset_index(drop=True, inplace=True)

    ayat_awal = str(staging_data_met.loc[0, "surat_ayat"])
    ayat_akhir = str(staging_data_met.loc[total_ayat - 1, "surat_ayat"])
    return staging_data_met, ayat_awal, ayat_akhir, center_pos


def build_process_1_and_2(staging_data_met: pd.DataFrame, metode: str):
    """Build Process 1 and 2 tables and the midpoint label."""
    total_ayat = len(staging_data_met)
    p1_dict = {}

    if total_ayat % 2 == 0:
        mid_left = total_ayat // 2 - 1
        mid_right = total_ayat // 2
        titik_tengah = f"{str(staging_data_met.loc[mid_left, 'surat_ayat'])}+{str(staging_data_met.loc[mid_right, 'surat_ayat'])}"
        left_idx = mid_left - 1
        right_idx = mid_right + 1
    else:
        mid = total_ayat // 2
        titik_tengah = str(staging_data_met.loc[mid, "surat_ayat"])
        left_idx = mid - 1
        right_idx = mid + 1

    p1_dict["F"] = [titik_tengah]

    met1_widths = [
        ("A", "a", 5), ("B", "b", 7), ("C1", "c1", 7), ("C2", "c2", 7),
        ("D1", "d1", 7), ("D2", "d2", 7), ("D3", "d3", 9),
        ("E", "e", 5), ("G", "g", 1)
    ]

    for left_lbl, right_lbl, width in met1_widths:
        left_chunk = staging_data_met.loc[left_idx - width + 1 : left_idx, "surat_ayat"].tolist()
        right_chunk = staging_data_met.loc[right_idx : right_idx + width - 1, "surat_ayat"].tolist()

        p1_dict[left_lbl] = left_chunk
        p1_dict[right_lbl] = right_chunk

        left_idx -= width
        right_idx += width

    order_p1 = [
        "G", "E", "D3", "D2", "D1", "C2", "C1", "B", "A", "F", "a", "b",
        "c1", "c2", "d1", "d2", "d3", "e", "g"
    ]
    df_p1 = pd.DataFrame(dict([(k, pd.Series(p1_dict.get(k, []))) for k in order_p1])).T
    df_p1.reset_index(inplace=True)
    df_p1.columns = ["Label"] + [f"Ayat {i+1}" for i in range(df_p1.shape[1] - 1)]

    order_p2 = ["F", "A", "a", "B", "b", "C2", "C1", "c1", "c2", "D3", "D2", "D1", "d1", "d2", "d3"]
    df_p2 = pd.DataFrame(dict([(k, pd.Series(p1_dict.get(k, []))) for k in order_p2])).T
    df_p2.reset_index(inplace=True)
    df_p2.columns = ["Label"] + [f"Ayat {i+1}" for i in range(df_p2.shape[1] - 1)]

    return p1_dict, df_p1, df_p2, titik_tengah


def build_process_3(df_p2: pd.DataFrame):
    """Flatten Process 2 into Process 3 output with labels and index markers."""
    flat_data = []
    p3_id_counter = 1

    for _, row in df_p2.iterrows():
        label = row["Label"]
        for col in df_p2.columns[1:]:
            ayat = row[col]
            if pd.notna(ayat):
                flat_data.append([p3_id_counter, label, ayat])
                p3_id_counter += 1

    df_p3 = pd.DataFrame(flat_data, columns=["p3_id", "Label", "Ayat"])

    try:
        at_h = str(df_p3.loc[38, "Ayat"])
        at_ah = str(df_p3.loc[49, "Ayat"])
    except (KeyError, IndexError):
        at_h = "Index Tidak Ditemukan"
        at_ah = "Index Tidak Ditemukan"

    return df_p3, at_h, at_ah


def build_init_table(processor, p1_dict: dict, titik_tengah: str, ayat_awal: str, ayat_akhir: str, at_h: str, at_ah: str):
    """Build the Init/Data summary table used for the export."""

    def get_padded_list(label: str, length: int = 5):
        lst = p1_dict.get(label, [])
        return lst + [None] * (length - len(lst))

    tabel_jb_data = [
        ["Variabel", "Ayat1", "Ayat2", "Ayat3", "Ayat4", "Ayat5"],
        [f"Ayat Awal {processor.metode_aktif}", ayat_awal, None, None, None, None],
        [f"Ayat Akhir {processor.metode_aktif}", ayat_akhir, None, None, None, None],
        [f"Ayat Tengah {processor.metode_aktif}", titik_tengah, None, None, None, None],
        [f"Ayat F {processor.metode_aktif}"] + get_padded_list("F"),
        [f"Ayat A {processor.metode_aktif}"] + get_padded_list("A"),
        [f"Ayat a {processor.metode_aktif}"] + get_padded_list("a"),
        [f"Ayat Dhzo {processor.metode_aktif}", at_ah, None, None, None, None],
        [f"Ayat E {processor.metode_aktif}"] + get_padded_list("E"),
        [f"Ayat G {processor.metode_aktif}"] + get_padded_list("G"),
        [f"Ayat e {processor.metode_aktif}"] + get_padded_list("e"),
        [f"Ayat g {processor.metode_aktif}"] + get_padded_list("g"),
        [f"Ayat AT-H {processor.metode_aktif}", at_h, None, None, None, None],
        [f"Ayat AT-Juz {processor.metode_aktif}", titik_tengah, None, None, None, None],
        [f"Ayat AT-AH {processor.metode_aktif}", at_ah, None, None, None, None],
        ["Total Ayat", f"{processor.total_ayat} ayat", None, None, None, None],
    ]

    kolom = [f"Tabel JB {processor.metode_aktif} Juz {processor.NILAI_FILTER}", "Ayat1", "Ayat2", "Ayat3", "Ayat4", "Ayat5"]
    return pd.DataFrame(tabel_jb_data[1:], columns=kolom)


def build_final_table(processor):
    """Build the final Asmaul Husna table used in the exported workbook."""
    latin_chars = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'I', 'II', 'III', 'Iv', 'V', 'VI', 'VII']
    arab_chars = ['ا', 'ب', 'ت', 'ث', 'ج', 'ح', 'خ', 'د', 'ذ', 'ر', 'ز', 'س', 'ش', 'ص', 'ض', 'ط', 'ظ', 'ع', 'غ', 'ف', 'ق', 'ك', 'ل', 'م', 'ن', 'و', 'ھ', 'لا', 'ء', 'ي', 'ال', 'ة', 'هؤ']

    baris_final: List[List[Any]] = []
    baris_final.append([processor.concept_label, None, None, None, None, None, None, None, None, None])
    baris_final.append([None, processor.metode_aktif, None, None, None, None, None, None, f'Juz {processor.NILAI_FILTER}', None])
    baris_final.append(['Latin & Hari', None, None, 'Ayat Khusus & Titik Kesadaran', None, None, None, 'Arab & Dasar', None, None])
    baris_final.append(['Latin', 'No AH', 'Ayat #3', 'No AH', 'Ayat #4', 'No AH', 'Ayat #2', 'No AH', 'Ayat #1', 'Arab'])

    for i in range(1, 34):
        idx_1 = 3 * i - 2
        idx_2 = 3 * i - 1
        idx_3 = 3 * i
        idx_4 = 101 - idx_3

        ayat1 = processor.df_p3.loc[idx_1 - 1, 'Ayat'] if idx_1 <= len(processor.df_p3) else None
        ayat2 = processor.df_p3.loc[idx_2 - 1, 'Ayat'] if idx_2 <= len(processor.df_p3) else None
        ayat3 = processor.df_p3.loc[idx_3 - 1, 'Ayat'] if idx_3 <= len(processor.df_p3) else None
        ayat4 = processor.df_p3.loc[idx_4 - 1, 'Ayat'] if idx_4 <= len(processor.df_p3) else None

        huruf_latin = latin_chars[i - 1]
        huruf_arab = arab_chars[i - 1]

        baris = [huruf_latin, idx_3, ayat3, idx_4, ayat4, idx_2, ayat2, idx_1, ayat1, huruf_arab]
        baris_final.append(baris)

    kolom = ['TABEL ASMAUL HUSNA', None, None, None, None, None, None, None, None, None]
    return pd.DataFrame(baris_final, columns=kolom)
