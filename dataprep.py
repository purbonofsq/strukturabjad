from __future__ import annotations

from typing import Any

import pandas as pd


def load_raw_excel(file_path: str) -> pd.DataFrame | None:
    """Load raw Excel data and return a DataFrame or None if it fails."""
    try:
        data = pd.read_excel(file_path)
        return data
    except Exception as exc:  # pragma: no cover - diagnostics path
        print(f"[ERROR] Error memuat data: {exc}")
        return None


def prepare_staging_data(raw_data: pd.DataFrame, filter_col: str, filter_value: Any) -> pd.DataFrame | None:
    """Filter the raw data and prepare the staging structure shared by the pipeline."""
    if raw_data is None:
        print("[ERROR] raw_data is None, cannot filter.")
        return None

    try:
        if filter_col not in raw_data.columns:
            available = list(raw_data.columns)
            print(f"[ERROR] Kolom '{filter_col}' tidak ditemukan di data. Kolom yang tersedia: {available}")
            return None

        staging_data = raw_data[raw_data[filter_col] == filter_value].copy()
        if len(staging_data) == 0:
            unique_values = raw_data[filter_col].unique()
            print(
                f"[ERROR] Tidak ada data untuk {filter_col} = {filter_value}. "
                f"Nilai yang tersedia: {sorted(unique_values)}"
            )
            return None

        staging_data.reset_index(drop=True, inplace=True)
        staging_data["surat_ayat"] = staging_data["aq_surat_id"].astype(str) + "." + staging_data["surat_ayat_no"].astype(str)
        return staging_data
    except Exception as exc:  # pragma: no cover - diagnostics path
        print(f"[ERROR] Error saat filtering data: {type(exc).__name__}: {exc}")
        import traceback
        traceback.print_exc()
        return None
