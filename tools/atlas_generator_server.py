#!/usr/bin/env python3
"""Local HTML front end for the ATLAS report generator."""

from __future__ import annotations

import argparse
import cgi
import html
import mimetypes
import shutil
import sys
import time
import zipfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

import atlas_report_generator as generator


ROOT = Path(__file__).resolve().parents[1]
RUN_ROOT = ROOT / "outputs" / "atlas_uploads" / "html_runs"
ASSET_ROOT = ROOT / "docs" / "portfolio-operations-dashboard" / "assets"


def safe_name(name: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in "._- " else "_" for ch in name).strip()
    return cleaned or "uploaded_report"


def make_zip(output_dir: Path) -> Path:
    zip_path = output_dir / "atlas_csv_package.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file_path in sorted(output_dir.glob("*.csv")) + sorted(output_dir.glob("*.xlsx")) + sorted(output_dir.glob("*.json")):
            archive.write(file_path, file_path.name)
    return zip_path


def read_manifest(output_dir: Path) -> list[dict[str, str]]:
    manifest = output_dir / "manifest.csv"
    if not manifest.exists():
        return []
    import csv

    with manifest.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def read_csv_rows(output_dir: Path, filename: str) -> list[dict[str, str]]:
    path = output_dir / filename
    if not path.exists():
        return []
    import csv

    with path.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def page_shell(title: str, body: str) -> bytes:
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    :root {{
      color-scheme: light;
      --ink: #102638;
      --muted: #5f7889;
      --line: #c8d8e3;
      --surface: #ffffff;
      --band: #eef4f8;
      --navy: #061d30;
      --blue: #0b3c5d;
      --blue-2: #06263d;
      --sky: #78acc5;
      --sky-soft: #dcecf4;
      --green: #1f7a4d;
      --red: #a43a3a;
      --gold: #9d762a;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font: 14px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at 12% -8%, rgba(120,172,197,.28), transparent 34%),
        linear-gradient(180deg, #f7fbfd 0%, var(--band) 46%, #eaf2f7 100%);
    }}
    header {{
      background:
        linear-gradient(135deg, rgba(6,29,48,.98), rgba(10,59,88,.96)),
        radial-gradient(circle at 90% 8%, rgba(120,172,197,.34), transparent 28%);
      color: white;
      padding: 26px 30px;
      border-bottom: 1px solid rgba(120,172,197,.45);
      box-shadow: 0 14px 36px rgba(6,29,48,.18);
    }}
    .brand {{
      max-width: 1180px;
      margin: 0 auto;
      display: grid;
      grid-template-columns: auto 1fr;
      gap: 18px;
      align-items: center;
    }}
    .brand-icon {{
      width: 72px;
      height: 72px;
      border-radius: 18px;
      box-shadow: 0 14px 30px rgba(0,0,0,.22);
    }}
    header h1 {{ margin: 0; font-size: 28px; letter-spacing: .18em; font-weight: 800; }}
    header p {{ margin: 5px 0 0; color: #d9eaf2; max-width: 920px; letter-spacing: .08em; text-transform: uppercase; font-size: 12px; }}
    .tagline {{ margin-top: 8px; color: var(--sky); font-weight: 700; letter-spacing: .2em; }}
    main {{ max-width: 1180px; margin: 24px auto; padding: 0 20px 40px; }}
    section {{
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 22px;
      margin-bottom: 18px;
      box-shadow: 0 12px 28px rgba(16,38,56,.07);
    }}
    h2 {{ margin: 0 0 12px; font-size: 18px; color: var(--navy); }}
    label {{ display: block; font-weight: 700; margin: 14px 0 6px; color: var(--navy); }}
    input[type="text"], input[type="file"] {{
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px;
      background: white;
      color: var(--ink);
    }}
    .checks {{ display: grid; gap: 12px; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); margin-top: 14px; }}
    .check {{
      display: grid;
      grid-template-columns: 18px 1fr;
      gap: 10px;
      align-items: start;
      color: var(--ink);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      background: #fbfdff;
    }}
    .check strong {{ display: block; color: var(--navy); }}
    .check span {{ display: block; color: var(--muted); font-weight: 500; margin-top: 3px; }}
    button, .button {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      border: 0;
      border-radius: 6px;
      background: linear-gradient(180deg, #0d4a72, #063553);
      color: white;
      padding: 11px 15px;
      font-weight: 700;
      text-decoration: none;
      cursor: pointer;
    }}
    button:hover, .button:hover {{ background: var(--blue-2); }}
    .secondary {{ background: #eef5f9; color: var(--blue); border: 1px solid #c8d8e3; }}
    .secondary:hover {{ background: #e2ebf5; }}
    .quiet {{ background: white; color: var(--blue); border: 1px solid var(--line); }}
    .quiet:hover {{ background: var(--sky-soft); }}
    .hint {{ color: var(--muted); margin: 8px 0 0; }}
    .guidance {{ display: grid; gap: 12px; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); }}
    .guidance-card {{ border-left: 3px solid var(--sky); padding: 8px 12px; background: #f7fbfd; }}
    .guidance-card h3 {{ margin: 0 0 4px; font-size: 14px; color: var(--navy); }}
    .guidance-card p {{ margin: 0; color: var(--muted); }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 12px; }}
    th, td {{ border-bottom: 1px solid var(--line); padding: 9px 8px; text-align: left; vertical-align: top; }}
    th {{ color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }}
    .ok {{ color: var(--green); font-weight: 700; }}
    .review {{ color: #9b6400; font-weight: 700; }}
    .error {{ color: var(--red); font-weight: 700; }}
    .actions {{ display: flex; gap: 10px; flex-wrap: wrap; margin-top: 16px; }}
    @media (max-width: 640px) {{
      header {{ padding: 18px; }}
      .brand {{ grid-template-columns: 1fr; }}
      .brand-icon {{ width: 58px; height: 58px; }}
      header h1 {{ font-size: 21px; }}
      main {{ padding: 0 12px 28px; }}
      section {{ padding: 16px; }}
      table {{ font-size: 12px; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="brand">
      <img class="brand-icon" src="/assets/atlas-icon.png" alt="ATLAS">
      <div>
        <h1>ATLAS</h1>
        <p>Portfolio | Tracker | Reports for RISE</p>
        <div class="tagline">LOVE. SERVE. CARE.</div>
      </div>
    </div>
  </header>
  <main>{body}</main>
</body>
</html>""".encode("utf-8")


def index_page() -> bytes:
    return page_shell(
        "ATLAS Report Generator",
        """
<section>
  <h2>Generate ATLAS Upload Package</h2>
  <form action="/generate" method="post" enctype="multipart/form-data">
    <label for="reports">Reports</label>
    <input id="reports" name="reports" type="file" multiple required
      accept=".xlsx,.xlsm,.csv,.pdf">
    <p class="hint">Supported: Box Score, Market Survey, Rent Roll PDF, Renewal Tracker, Trending Occupancy CSV.</p>

    <label for="report_month">Fallback Report Month</label>
    <input id="report_month" name="report_month" type="text" placeholder="Optional, e.g. 2026-04">
    <p class="hint">Source dates are used first. This only fills missing dates unless you force it.</p>

    <div class="checks">
      <label class="check"><input type="checkbox" name="force_report_month"><span><strong>Force fallback month</strong><span>Use only for files without trustworthy dates, or when intentionally reclassifying every uploaded file into one reporting month.</span></span></label>
      <label class="check"><input type="checkbox" name="include_pii"><span><strong>Include resident PII fields</strong><span>Use only for approved restricted uploads that need resident names, phone numbers, balances, or deposits.</span></span></label>
      <label class="check"><input type="checkbox" name="skip_pdf"><span><strong>Skip PDF rent-roll parsing</strong><span>Use when you only need Excel/CSV data or want a faster monthly package without unit-level rent-roll detail.</span></span></label>
    </div>

    <div class="actions">
      <button type="submit">Generate Upload Package</button>
    </div>
  </form>
</section>
<section>
  <h2>Checkbox Guide</h2>
  <div class="guidance">
    <div class="guidance-card"><h3>Force fallback month</h3><p>Usually leave unchecked. Check it only when source dates are wrong or you are intentionally loading every file into the month typed above.</p></div>
    <div class="guidance-card"><h3>Include resident PII</h3><p>Usually leave unchecked. Check it only for secure resident-level workflows approved to hold names, phones, balances, and deposits.</p></div>
    <div class="guidance-card"><h3>Skip PDF rent roll</h3><p>Check it for faster runs when PDFs are attached but you do not need rent-roll unit detail.</p></div>
  </div>
</section>
<section>
  <h2>How To Load</h2>
  <p class="hint">Use the combined Excel workbook for one reviewed package with tabs. Use <strong>monthly_upload</strong> for the current ATLAS monthly upload. Load detail tabs only when ATLAS or Power BI has matching detail tables.</p>
  <p class="hint">Rows marked <strong>REVIEW</strong> should stay out of production until the exception is resolved.</p>
</section>
""",
    )


def results_page(run_id: str, output_dir: Path) -> bytes:
    manifest = read_manifest(output_dir)
    exceptions = read_csv_rows(output_dir, "exceptions.csv")
    reviews = [row for row in read_csv_rows(output_dir, "source_file_log.csv") if row.get("Status") in {"REVIEW", "ERROR"}]
    rows = []
    for item in manifest:
        table = item.get("Table", "")
        csv_file = item.get("CSV_File", "")
        count = item.get("Rows", "")
        link = f"/runs/{quote(run_id)}/{quote(csv_file)}"
        rows.append(
            f"<tr><td>{html.escape(table)}</td><td><a href=\"{link}\">{html.escape(csv_file)}</a></td><td>{html.escape(count)}</td></tr>"
        )
    exception_rows = []
    for item in exceptions:
        exception_rows.append(
            "<tr>"
            f"<td>{html.escape(item.get('Severity', ''))}</td>"
            f"<td>{html.escape(item.get('Source_Location', ''))}</td>"
            f"<td>{html.escape(item.get('Issue', ''))}</td>"
            f"<td>{html.escape(item.get('Recommended_Action', ''))}</td>"
            "</tr>"
        )
    review_rows = []
    for item in reviews:
        review_rows.append(
            "<tr>"
            f"<td>{html.escape(item.get('Detected_Source_Type', ''))}</td>"
            f"<td>{html.escape(item.get('Source_File_Name', ''))}</td>"
            f"<td>{html.escape(item.get('Status', ''))}</td>"
            f"<td>{html.escape(item.get('Notes', ''))}</td>"
            "</tr>"
        )
    status_line = (
        f"<p class=\"review\">Review items detected: {len(exceptions)} exception(s), {len(reviews)} source file(s) flagged REVIEW/ERROR.</p>"
        if exceptions or reviews
        else "<p class=\"ok\">No package warnings or source review flags were generated.</p>"
    )
    body = f"""
<section>
  <h2>ATLAS Upload Package Ready</h2>
  <p class="ok">Generated run {html.escape(run_id)}.</p>
  {status_line}
  <div class="actions">
    <a class="button" href="/runs/{quote(run_id)}/atlas_combined_upload.xlsx">Download Excel Workbook</a>
    <a class="button secondary" href="/runs/{quote(run_id)}/atlas_csv_package.zip">Download CSV ZIP</a>
    <a class="button secondary" href="/">Run Another Batch</a>
  </div>
</section>
<section>
  <h2>Outputs</h2>
  <table>
    <thead><tr><th>Table</th><th>CSV</th><th>Rows</th></tr></thead>
    <tbody>{''.join(rows)}</tbody>
  </table>
</section>
<section>
  <h2>Exceptions</h2>
  <table>
    <thead><tr><th>Severity</th><th>Location</th><th>Issue</th><th>Recommended Action</th></tr></thead>
    <tbody>{''.join(exception_rows) if exception_rows else '<tr><td colspan="4">No exceptions generated.</td></tr>'}</tbody>
  </table>
</section>
<section>
  <h2>Source Review Flags</h2>
  <table>
    <thead><tr><th>Source Type</th><th>File</th><th>Status</th><th>Notes</th></tr></thead>
    <tbody>{''.join(review_rows) if review_rows else '<tr><td colspan="4">No source files were flagged for review.</td></tr>'}</tbody>
  </table>
</section>
"""
    return page_shell("ATLAS Upload Package Ready", body)


class AtlasGeneratorHandler(BaseHTTPRequestHandler):
    def send_html(self, content: bytes, status: HTTPStatus = HTTPStatus.OK) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_html(index_page())
            return
        if parsed.path.startswith("/assets/"):
            filename = unquote(parsed.path.split("/assets/", 1)[1])
            asset_path = (ASSET_ROOT / filename).resolve()
            if ASSET_ROOT.resolve() not in asset_path.parents or not asset_path.exists() or not asset_path.is_file():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            content_type = mimetypes.guess_type(asset_path.name)[0] or "application/octet-stream"
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(asset_path.stat().st_size))
            self.end_headers()
            with asset_path.open("rb") as handle:
                shutil.copyfileobj(handle, self.wfile)
            return
        if parsed.path.startswith("/runs/"):
            parts = [unquote(part) for part in parsed.path.split("/") if part]
            if len(parts) != 3:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            _, run_id, filename = parts
            run_dir = (RUN_ROOT / run_id).resolve()
            file_path = (run_dir / filename).resolve()
            if RUN_ROOT.resolve() not in file_path.parents or not file_path.exists() or not file_path.is_file():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            if file_path.suffix == ".zip":
                content_type = "application/zip"
            elif file_path.suffix == ".csv":
                content_type = "text/csv; charset=utf-8"
            elif file_path.suffix == ".xlsx":
                content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                content_type = "application/octet-stream"
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f'attachment; filename="{file_path.name}"')
            self.send_header("Content-Length", str(file_path.stat().st_size))
            self.end_headers()
            with file_path.open("rb") as handle:
                shutil.copyfileobj(handle, self.wfile)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:  # noqa: N802
        if urlparse(self.path).path != "/generate":
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            },
        )
        run_id = time.strftime("%Y%m%d_%H%M%S")
        output_dir = RUN_ROOT / run_id
        upload_dir = output_dir / "_source_uploads"
        upload_dir.mkdir(parents=True, exist_ok=True)

        upload_items = form["reports"] if "reports" in form else []
        if not isinstance(upload_items, list):
            upload_items = [upload_items]
        files: list[Path] = []
        for item in upload_items:
            if not getattr(item, "filename", ""):
                continue
            destination = upload_dir / safe_name(item.filename)
            with destination.open("wb") as handle:
                shutil.copyfileobj(item.file, handle)
            files.append(destination)

        if not files:
            self.send_html(page_shell("No Files Uploaded", "<section><h2>No Files Uploaded</h2><p class=\"error\">Choose at least one report file.</p><a class=\"button secondary\" href=\"/\">Back</a></section>"), HTTPStatus.BAD_REQUEST)
            return

        report_month = generator.report_month_from_cli(form.getfirst("report_month", "")) if form.getfirst("report_month", "") else ""
        force_report_month = "force_report_month" in form
        include_pii = "include_pii" in form
        skip_pdf = "skip_pdf" in form

        state = generator.RunState()
        for file_path in files:
            generator.parse_file(file_path, state, report_month, force_report_month, include_pii, skip_pdf)
        generator.write_outputs(output_dir, state)
        make_zip(output_dir)
        self.send_html(results_page(run_id, output_dir))

    def log_message(self, format: str, *args: object) -> None:
        sys.stderr.write("%s - %s\n" % (self.address_string(), format % args))


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the local ATLAS generator HTML front end.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=8765, type=int)
    args = parser.parse_args()
    RUN_ROOT.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((args.host, args.port), AtlasGeneratorHandler)
    print(f"ATLAS Report Generator running at http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
