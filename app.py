"""
app.py
------
The whole product, end to end:

  buyer uploads Excel  -->  pool_coder codes + balances teams
                       -->  build_presentation slots them into the template
                       -->  buyer downloads a finished .pptx

Run locally:
    pip install -r requirements.txt
    python3 app.py
    open http://localhost:5000

Deploy on Render: see README.md in this folder.
"""

import io
import os
import tempfile

from flask import Flask, request, send_file, render_template_string

from pool_coder import run as run_pool_coder
from read_excel import read_entries_from_excel
from build_presentation import build_presentation

app = Flask(__name__)

TEMPLATE_PPTX = os.path.join(os.path.dirname(__file__), "template.pptx")

PAGE = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Debate Pool Coder</title>
  <style>
    body { font-family: Calibri, Arial, sans-serif; background: #F7F8FA; color: #15203B;
           max-width: 640px; margin: 60px auto; padding: 0 20px; }
    h1 { font-family: Cambria, Georgia, serif; color: #10204A; }
    .card { background: white; border: 1px solid #E3E6ED; border-radius: 8px; padding: 24px; }
    label { display: block; font-weight: 600; margin-bottom: 6px; margin-top: 16px; }
    input[type=file], input[type=number] { width: 100%; padding: 8px; border: 1px solid #E3E6ED;
           border-radius: 6px; box-sizing: border-box; }
    button { margin-top: 20px; background: #10204A; color: white; border: none; padding: 12px 20px;
           border-radius: 6px; font-weight: 600; cursor: pointer; width: 100%; }
    .error { background: #FBEFEF; border: 1px solid #E2B6B6; color: #7A2E2E; border-radius: 6px;
           padding: 12px; margin-top: 16px; font-size: 14px; }
    .hint { color: #5C6B8A; font-size: 13px; margin-top: 4px; }
  </style>
</head>
<body>
  <h1>Pool Coder</h1>
  <p class="hint">Upload your school list. Get back a finished presentation with teams coded and pools balanced.</p>
  <div class="card">
    <form action="/build" method="post" enctype="multipart/form-data">
      <label>Excel file</label>
      <input type="file" name="excel_file" accept=".xlsx" required>
      <div class="hint">Needs columns named School, Category, Teams in the first row.</div>

      <label>Number of pools per category</label>
      <input type="number" name="num_pools" value="4" min="2" max="12">

      <button type="submit">Generate presentation</button>
    </form>
    {% if errors %}
      <div class="error">
        <strong>Heads up — {{ errors|length }} row(s) skipped:</strong>
        <ul>{% for e in errors %}<li>{{ e }}</li>{% endfor %}</ul>
      </div>
    {% endif %}
  </div>
</body>
</html>
"""


@app.route("/", methods=["GET"])
def index():
    return render_template_string(PAGE, errors=[])


@app.route("/build", methods=["POST"])
def build():
    file = request.files.get("excel_file")
    num_pools = int(request.form.get("num_pools", 4))

    if not file or file.filename == "":
        return render_template_string(PAGE, errors=["No file was uploaded."])

    entries, errors = read_entries_from_excel(file)
    if not entries:
        return render_template_string(PAGE, errors=errors or ["No usable rows found in that file."])

    result = run_pool_coder(entries, num_pools=num_pools)

    # Convert {"Category": {"pools": {"Pool A": [...]}}} -> {"Category": {"A": [...]}}
    pools_by_category = {}
    for category, data in result.items():
        pools_by_category[category] = {
            pool_name.split()[-1]: teams for pool_name, teams in data["pools"].items()
        }

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        output_path = tmp.name

    build_presentation(TEMPLATE_PPTX, output_path, pools_by_category)

    with open(output_path, "rb") as f:
        buffer = io.BytesIO(f.read())
    os.remove(output_path)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Debate_Pools.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
