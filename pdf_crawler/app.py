import os
import uuid
import shutil
from flask import Flask, render_template, request, send_file, after_this_request
from urllib.parse import urlparse

from crawler import crawl_site

app = Flask(__name__)

BASE_JOBS_FOLDER = "jobs"
ALLOWED_DOMAINS = ["www.nsw.gov.au", "nsw.gov.au"]


def is_valid_nsw_url(url):
    parsed = urlparse(url)
    return parsed.scheme in ["http", "https"] and parsed.netloc in ALLOWED_DOMAINS


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html")

    landing_url = request.form.get("landing_url", "").strip()

    if not landing_url.startswith("http"):
        landing_url = "https://" + landing_url

    if not is_valid_nsw_url(landing_url):
        return render_template(
            "index.html",
            error="Please enter a valid nsw.gov.au URL."
        )

    job_id = str(uuid.uuid4())
    job_folder = os.path.join(BASE_JOBS_FOLDER, job_id)
    os.makedirs(job_folder, exist_ok=True)

    try:
        zip_path = crawl_site(
            start_url=landing_url,
            job_folder=job_folder
        )
    except Exception as e:
        shutil.rmtree(job_folder, ignore_errors=True)
        return render_template(
            "index.html",
            error=f"Something went wrong: {e}"
        )

    @after_this_request
    def cleanup(response):
        try:
            shutil.rmtree(job_folder, ignore_errors=True)
        except Exception:
            pass
        return response

    return send_file(
        zip_path,
        as_attachment=True,
        download_name="pdf_crawl_results.zip"
    )


if __name__ == "__main__":
    app.run(debug=True)