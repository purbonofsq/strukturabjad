from pathlib import Path

from flask import Flask, jsonify

from backend1 import PengolahDataJBAH

app = Flask(__name__)

REPO_ROOT = Path(__file__).resolve().parent
DATA_FILE = REPO_ROOT / "datasource" / "Struktur Quran18.xlsx"


@app.get("/")
def home():
    return jsonify(
        {
            "project": "strukturabjad",
            "status": "ok",
            "message": "Python backend is running",
            "repo_root": str(REPO_ROOT),
        }
    )


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.get("/api/files")
def list_files():
    entries = [
        {"name": path.name, "type": "directory" if path.is_dir() else "file"}
        for path in sorted(REPO_ROOT.iterdir())
    ]
    return jsonify({"files": entries})


@app.get("/api/juz/<int:juz_id>")
def process_juz(juz_id: int):
    if not DATA_FILE.exists():
        return jsonify({"error": f"Data file not found: {DATA_FILE}"}), 404

    processor = PengolahDataJBAH(str(DATA_FILE))
    if processor.raw_data is None:
        return jsonify({"error": "Unable to load raw data file."}), 500

    staging_data = processor.filter_data("aq_juz_id", juz_id)
    if staging_data is None:
        return jsonify({"error": f"No data found for juz {juz_id}."}), 404

    results = {}
    for method in ["T111", "T222", "T333"]:
        processor.set_metode(method)
        processor.proses_1_dan_2()
        processor.proses_3()
        processor.generate_tabel_jb_init()
        processor.generate_tabel_ah_final()

        results[method] = {
            "total_ayat": int(getattr(processor, "total_ayat", 0)),
            "ayat_awal": getattr(processor, "ayat_awal", None),
            "ayat_akhir": getattr(processor, "ayat_akhir", None),
            "titik_tengah": getattr(processor, "titik_tengah", None),
            "at_h": getattr(processor, "at_h", None),
            "at_ah": getattr(processor, "at_ah", None),
            "labels": getattr(processor.df_p3, "to_dict", lambda orient="records": [])("records")[:5],
        }

    return jsonify(
        {
            "juz_id": juz_id,
            "data_file": str(DATA_FILE),
            "results": results,
        }
    )


@app.get("/api/juz/<int:juz_id>/summary")
def process_juz_summary(juz_id: int):
    response = process_juz(juz_id)
    if response.status_code != 200:
        return response

    payload = response.get_json()
    summary = {
        "juz_id": payload["juz_id"],
        "methods": {
            method: {
                "total_ayat": data["total_ayat"],
                "ayat_awal": data["ayat_awal"],
                "ayat_akhir": data["ayat_akhir"],
                "titik_tengah": data["titik_tengah"],
                "at_h": data["at_h"],
                "at_ah": data["at_ah"],
            }
            for method, data in payload["results"].items()
        },
    }
    return jsonify(summary)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
