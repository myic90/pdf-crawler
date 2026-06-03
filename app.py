import os
import uuid
import shutil
import traceback
from flask import Flask, render_template, request, send_file, after_this_request
from urllib.parse import urlparse

from crawler import crawl_site

app = Flask(__name__)

BASE_JOBS_FOLDER = "jobs"
ALLOWED_DOMAINS = ["www.nsw.gov.au", "nsw.gov.au"]


def is_valid_nsw_url(url):
    parsed = urlparse(url)
    return parsed.scheme in ["http", "https"] and parsed.netloc in ALLOWED_DOMAINS


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/", methods=["POST"])
def crawl():
    landing_url = request.form.get("landing_url", "").strip()

    if not landing_url.startswith("http"):
        landing_url = "https://" + landing_url

    if not is_valid_nsw_url(landing_url):
        return "Please enter a valid nsw.gov.au URL.", 400

    job_id = str(uuid.uuid4())
    job_folder = os.path.join(BASE_JOBS_FOLDER, job_id)
    os.makedirs(job_folder, exist_ok=True)

    try:
        zip_path, summary = crawl_site(
            start_url=landing_url,
            job_folder=job_folder
        )

        response = send_file(
            zip_path,
            as_attachment=True,
            download_name="pdf_crawl_results.zip"
        )

        response.headers["X-Pages-Crawled"] = str(summary.get("pages_crawled", 0))
        response.headers["X-Page-Limit-Reached"] = str(summary.get("page_limit_reached", "No"))
        response.headers["X-Unique-PDFs"] = str(summary.get("unique_pdfs_found", 0))
        response.headers["X-PDF-Limit-Reached"] = str(summary.get("pdf_limit_reached", "No"))
        response.headers["X-PDF-Records"] = str(summary.get("pdf_records", 0))
        response.headers["X-Duplicate-References"] = str(summary.get("duplicate_references", 0))
        response.headers["X-Failed-Downloads"] = str(summary.get("failed_downloads", 0))
        response.headers["X-Crawl-Duration-Seconds"] = str(summary.get("crawl_duration_seconds", 0))

        @after_this_request
        def cleanup(response):
            try:
                shutil.rmtree(job_folder, ignore_errors=True)

                if os.path.exists(zip_path):
                    os.remove(zip_path)
            except Exception:
                pass

            return response

        return response

    except Exception as e:
        shutil.rmtree(job_folder, ignore_errors=True)
        print(traceback.format_exc())
        return f"Server error: {e}", 500


if __name__ == "__main__":
    app.run(debug=True)
