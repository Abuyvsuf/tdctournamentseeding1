"""
app.py
------
The whole product, end to end, multi-region:

  organizer uploads ONE registration spreadsheet (with a Region column)
  + a .zip of per-region template .pptx files (filename must contain the
    region name, e.g. "Upper_Eastern_Region_....pptx")
       -->  pool_coder codes + balances teams, per region, per language/category
       -->  build_presentation slots them into each region's own template
       -->  organizer downloads a .zip with one finished .pptx per region

Run locally:
    pip install -r requirements.txt
    python3 app.py
    open http://localhost:5000
"""

import io
import os
import re
import tempfile
import zipfile

from flask import Flask, request, send_file, render_template_string

from pool_coder import run_by_region
from read_registration import read_registration
from build_presentation import build_presentation

app = Flask(__name__)

PAGE = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Debate Pool Coder</title>
  <style>
    body { font-family: Calibri, Arial, sans-serif; background: #F7F8FA; color: #15203B;
           max-width: 680px; margin: 60px auto; padding: 0 20px; }
    h1 { font-family: Cambria, Georgia, serif; color: #10204A; }
    .card { background: white; border: 1px solid #E3E6ED; border-radius: 8px; padding: 24px; }
    label { display: block; font-weight: 600; margin-bottom: 6px; margin-top: 16px; }
    input[type=file], input[type=number] { width: 100%; padding: 8px; border: 1px solid #E3E6ED;
           border-radius: 6px; box-sizing: border-box; }
    button { margin-top: 20px; background: #10204A; color: white; border: none; padding: 12px 20px;
           border-radius: 6px; font-weight: 600; cursor: pointer; width: 100%; }
    .error { background: #FBEFEF; border: 1px solid #E2B6B6; color: #7A2E2E; border-radius: 6px;
           padding: 12px; margin-top: 16px; font-size: 14px; }
    .review { background: #FFF8E8; border: 1px solid #E8D69A; color: #6B5510; border-radius: 6px;
           padding: 12px; margin-top: 16px; font-size: 13px; max-height: 240px; overflow-y: auto; }
    .hint { color: #5C6B8A; font-size: 13px; margin-top: 4px; }
    ul { margin: 6px 0; padding-left: 20px; }
  </style>
</head>
<body>
  <h1>Pool Coder</h1>
  <p class="hint">Upload the registration sheet and a zip of each region's template.
  Get back one finished presentation per region.</p>
  <div class="card">
    <form action="/build" method="post" enctype="multipart/form-data">
      <label>Registration spreadsheet (.xlsx)</label>
      <input type="file" name="excel_file" accept=".xlsx" required>
      <div class="hint">Needs: Name of School, Region, and the four team-count columns.</div>

      <label>Region templates (.zip of .pptx files)</label>
      <input type="file" name="templates_zip" accept=".zip" required>
      <div class="hint">Each filename must contain its region's name, e.g. "Upper_Eastern_Region_....pptx".</div>

      <button type="submit">Generate presentations</button>
    </form>
    {% if errors %}
      <div class="error">
        <strong>Couldn't proceed:</strong>
        <ul>{% for e in errors %}<li>{{ e }}</li>{% endfor %}</ul>
      </div>
    {% endif %}
    {% if needs_review %}
      <div class="review">
        <strong>Heads up — {{ needs_review|length }} item(s) worth double-checking:</strong>
        <ul>{% for r in needs_review %}<li>{{ r }}</li>{% endfor %}</ul>
      </div>
    {% endif %}
  </div>
</body>
</html>
"""


def _normalize(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())


@app.route("/", methods=["GET"])
def index():
    return render_template_string(PAGE, errors=[], needs_review=[])


@app.route("/build", methods=["POST"])
def build():
    excel_file = request.files.get("excel_file")
    templates_zip = request.files.get("templates_zip")

    if not excel_file or excel_file.filename == "":
        return render_template_string(PAGE, errors=["No spreadsheet was uploaded."], needs_review=[])
    if not templates_zip or templates_zip.filename == "":
        return render_template_string(PAGE, errors=["No templates zip was uploaded."], needs_review=[])

    try:
        return _do_build(excel_file, templates_zip)
    except Exception as exc:
        app.logger.exception("Unexpected error in /build")
        return render_template_string(
            PAGE,
            errors=[
                f"Something unexpected went wrong ({type(exc).__name__}: {exc}). "
                "Check the Render logs for the full traceback, or share this message."
            ],
            needs_review=[],
        )


def _do_build(excel_file, templates_zip):
    entries, needs_review, errors = read_registration(excel_file)
    if errors:
        return render_template_string(PAGE, errors=errors, needs_review=needs_review)
    if not entries:
        return render_template_string(
            PAGE, errors=["No usable rows found -- check that Region is filled in for every school."],
            needs_review=needs_review,
        )

    by_region = run_by_region(entries)

    with tempfile.TemporaryDirectory() as tmp_dir:
        try:
            zin = zipfile.ZipFile(templates_zip)
        except zipfile.BadZipFile:
            return render_template_string(
                PAGE,
                errors=[
                    "The 'Region templates' file isn't a valid .zip file. "
                    "Make sure you're uploading the zipped folder, not the .pptx itself -- "
                    "right-click your .pptx file(s) and choose 'Compress'/'Send to "
                    "Compressed (zipped) folder' first, then upload that .zip file here."
                ],
                needs_review=[],
            )
        template_paths = {}  # normalized region-name-fragment -> extracted path
        for name in zin.namelist():
            if not name.lower().endswith(".pptx"):
                continue
            extracted_path = os.path.join(tmp_dir, os.path.basename(name))
            with open(extracted_path, "wb") as f:
                f.write(zin.read(name))
            template_paths[_normalize(name)] = extracted_path

        output_buffer = io.BytesIO()
        build_errors = []
        with zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
            for region, data in by_region.items():
                # find a template filename that contains this region's name
                region_key = _normalize(region)
                match = next((path for key, path in template_paths.items() if region_key in key), None)
                if not match:
                    build_errors.append(
                        f"No template filename matched region '{region}' -- skipped. "
                        f"Make sure a .pptx in the zip has '{region}' in its filename."
                    )
                    continue

                pools_by_lang_cat = {
                    key: {n.split()[-1]: teams for n, teams in info["pools"].items()}
                    for key, info in data.items()
                }

                out_path = os.path.join(tmp_dir, f"{region.replace(' ', '_')}_Pools.pptx")
                warnings = build_presentation(match, out_path, pools_by_lang_cat)
                build_errors.extend(f"{region}: {w}" for w in warnings)

                zout.write(out_path, arcname=f"{region.replace(' ', '_')}_Pools.pptx")

        if build_errors:
            needs_review = needs_review + build_errors

        output_buffer.seek(0)

        if needs_review:
            final_buffer = io.BytesIO()
            with zipfile.ZipFile(final_buffer, "w", zipfile.ZIP_DEFLATED) as zfinal:
                with zipfile.ZipFile(output_buffer) as zsrc:
                    for item in zsrc.namelist():
                        zfinal.writestr(item, zsrc.read(item))
                zfinal.writestr("READ_ME_review_notes.txt", "\n".join(needs_review))
            final_buffer.seek(0)
            output_buffer = final_buffer

        return send_file(
            output_buffer,
            as_attachment=True,
            download_name="Debate_Pools_By_Region.zip",
            mimetype="application/zip",
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
    app.run(host="0.0.0.0", port=port, debug=True)
