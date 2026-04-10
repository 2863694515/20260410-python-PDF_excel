import argparse
import os
import tempfile
import threading
import webbrowser
from io import BytesIO
from datetime import datetime

from flask import Flask, jsonify, render_template, request, send_file

from pdf_table_extractor import convert_pdf_to_excel

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PDF 转 Excel 本地网页工具")
    parser.add_argument("--host", default="127.0.0.1", help="服务监听地址")
    parser.add_argument("--port", type=int, default=5000, help="服务端口")
    parser.add_argument("--no-browser", action="store_true", help="启动时不自动打开浏览器")
    return parser.parse_args()


def _safe_name(filename: str) -> str:
    base = os.path.basename(filename)
    name, _ = os.path.splitext(base)
    return "".join(ch for ch in name if ch.isalnum() or ch in ("-", "_", " ")).strip() or "output"


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/convert")
def api_convert():
    if "file" not in request.files:
        return jsonify({"ok": False, "message": "未检测到上传文件"}), 400

    file = request.files["file"]
    if not file or not file.filename:
        return jsonify({"ok": False, "message": "请选择 PDF 文件"}), 400

    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"ok": False, "message": "仅支持 .pdf 文件"}), 400

    safe = _safe_name(file.filename)

    try:
        with tempfile.TemporaryDirectory(prefix="pdf2excel_") as tmpdir:
            in_pdf = os.path.join(tmpdir, f"{safe}.pdf")
            out_xlsx = os.path.join(tmpdir, f"{safe}_final.xlsx")
            file.save(in_pdf)

            convert_pdf_to_excel(in_pdf, out_xlsx)

            # Windows 下直接 send_file(磁盘路径) + 临时目录清理，
            # 可能触发文件占用冲突；先读入内存再返回更稳健。
            with open(out_xlsx, "rb") as f:
                payload = f.read()
            io_file = BytesIO(payload)
            io_file.seek(0)

            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            download_name = f"{safe}_{stamp}.xlsx"
            return send_file(
                io_file,
                as_attachment=True,
                download_name=download_name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as exc:
        return jsonify({"ok": False, "message": f"转换失败: {exc}"}), 500


if __name__ == "__main__":
    args = parse_args()
    url = f"http://{args.host}:{args.port}"

    if not args.no_browser:
        # 延时打开浏览器，避免服务尚未就绪导致打不开。
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()

    app.run(host=args.host, port=args.port, debug=False)
