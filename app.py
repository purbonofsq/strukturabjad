from flask import Flask, jsonify

app = Flask(__name__)


@app.get("/")
def home():
    return {"message": "Struktur Abjad Old backend is running"}


@app.get("/health")
def health():
    return jsonify({"status": "ok"})


@app.get("/api/hello")
def hello():
    return jsonify({"hello": "world"})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
