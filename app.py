import os
import uuid
import json
from flask import Flask, request, send_file, render_template, jsonify
from werkzeug.utils import secure_filename
from converter import convert_mode1, convert_mode2, pdf_to_blocks, blocks_to_docx

app = Flask(__name__)
app.config['TEMPLATES_AUTO_RELOAD'] = True
UPLOAD_FOLDER = "/tmp/pdf2word"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed(filename, exts):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in exts


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/parse", methods=["POST"])
def parse():
    """Upload PDF and return structured blocks for preview."""
    if "pdf_file" not in request.files:
        return jsonify({"error": "未上传文件"}), 400
    f = request.files["pdf_file"]
    if not f.filename or not allowed(f.filename, {"pdf"}):
        return jsonify({"error": "请上传 PDF 文件"}), 400

    job_id = uuid.uuid4().hex
    pdf_path = os.path.join(UPLOAD_FOLDER, f"{job_id}.pdf")
    f.save(pdf_path)

    try:
        blocks = pdf_to_blocks(pdf_path)
        return jsonify({"job_id": job_id, "blocks": blocks})
    except Exception as e:
        import traceback; print(traceback.format_exc(), flush=True)
        return jsonify({"error": str(e)}), 500


@app.route("/convert", methods=["POST"])
def convert():
    """Convert using edited blocks JSON → docx."""
    job_id = request.form.get("job_id", "")
    mode   = request.form.get("mode", "1")

    if not job_id or not all(c in "0123456789abcdef" for c in job_id) or len(job_id) != 32:
        return jsonify({"error": "无效的 job_id"}), 400

    pdf_path = os.path.join(UPLOAD_FOLDER, f"{job_id}.pdf")
    out_path = os.path.join(UPLOAD_FOLDER, f"{job_id}.docx")

    if not os.path.exists(pdf_path):
        return jsonify({"error": "PDF 文件已过期，请重新上传"}), 400

    try:
        blocks_json = request.form.get("blocks")
        if blocks_json:
            # Use edited blocks from preview
            blocks_data = json.loads(blocks_json)
            ref_path = None
            if mode == "2" and "ref_doc" in request.files and request.files["ref_doc"].filename:
                ref_file = request.files["ref_doc"]
                if allowed(ref_file.filename, {"docx"}):
                    ref_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_ref.docx")
                    ref_file.save(ref_path)
            blocks_to_docx(blocks_data, out_path, ref_path)
        else:
            # Fallback: direct conversion
            if mode == "1":
                convert_mode1(pdf_path, out_path)
            else:
                if "ref_doc" not in request.files or not request.files["ref_doc"].filename:
                    return jsonify({"error": "模式2需要上传参考 Word 文档"}), 400
                ref_file = request.files["ref_doc"]
                if not allowed(ref_file.filename, {"docx"}):
                    return jsonify({"error": "参考文件必须是 .docx 格式"}), 400
                ref_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_ref.docx")
                ref_file.save(ref_path)
                convert_mode2(pdf_path, out_path, ref_path)

        orig = secure_filename(request.form.get("original_name", "converted"))
        base = os.path.splitext(orig)[0] or "converted"
        return jsonify({"download_id": job_id, "filename": base + ".docx"})

    except Exception as e:
        import traceback; print(traceback.format_exc(), flush=True)
        return jsonify({"error": f"转换失败: {str(e)}"}), 500


@app.route("/ref-styles", methods=["POST"])
def ref_styles():
    if "ref_doc" not in request.files:
        return jsonify({"error": "no file"}), 400
    f = request.files["ref_doc"]
    ref_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex}_ref.docx")
    f.save(ref_path)
    try:
        from docx import Document
        doc = Document(ref_path)
        styles = {}
        for style_name in ["Normal", "Heading 1", "Heading 2"]:
            try:
                s = doc.styles[style_name]
                font = s.font
                pf = s.paragraph_format
                styles[style_name] = {
                    "font_name": font.name or "",
                    "font_size": font.size.pt if font.size else None,
                    "bold": font.bold,
                    "color": f"#{font.color.rgb}" if font.color and font.color.type and font.color.rgb else "#000000",
                    "space_before": pf.space_before.pt if pf.space_before else 0,
                    "space_after": pf.space_after.pt if pf.space_after else 0,
                    "line_spacing": pf.line_spacing.pt if pf.line_spacing else None,
                }
            except Exception:
                styles[style_name] = {}
        return jsonify({"styles": styles})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/download/<job_id>")
def download(job_id):
    if not all(c in "0123456789abcdef" for c in job_id) or len(job_id) != 32:
        return jsonify({"error": "无效ID"}), 400
    path = os.path.join(UPLOAD_FOLDER, f"{job_id}.docx")
    if not os.path.exists(path):
        return jsonify({"error": "文件不存在或已过期"}), 404
    name = secure_filename(request.args.get("name", "converted.docx")) or "converted.docx"
    return send_file(path, as_attachment=True, download_name=name)


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=8765)
