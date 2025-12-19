import os
import sqlite3
import pandas as pd
import re
from flask import Flask, render_template, request, send_file
from weasyprint import HTML
from jinja2 import Template

app = Flask(__name__)

# ---------------- PATHS ----------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "certificates.db")
PDF_DIR = os.path.join(BASE_DIR, "generated", "pdfs")


# ---------------- DATABASE INIT ----------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS certificates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            certificate_number TEXT,
            student_name TEXT,
            pdf_path TEXT
        )
    """)
    conn.commit()
    conn.close()


# ---------------- NUMBER SYSTEM ----------------
def get_next_certificate_number():

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    c.execute("SELECT certificate_number FROM certificates ORDER BY id DESC LIMIT 1")
    row = c.fetchone()
    conn.close()

    if not row:
        return "ACDT-C-25-001"

    last_no = row[0]
    match = re.search(r"(\d+)$", last_no)

    next_no = 1 if not match else int(match.group(1)) + 1

    return f"ACDT-C-25-{next_no:03d}"


# ---------------- ROUTE ----------------
@app.route("/", methods=["GET", "POST"])
def upload():

    if request.method == "POST":

        excel_file = request.files.get("excel")
        custom_content = request.form.get("content", "").strip()
        single_name = request.form.get("student_name", "").strip()
        single_date = request.form.get("single_date", "").strip()
        single_place = request.form.get("single_place", "").strip()


        # ======================================================
        # 1️⃣ BULK MODE
        # ======================================================
        if excel_file and excel_file.filename != "":

            df = pd.read_excel(excel_file)

            df.columns = (
                df.columns.astype(str)
                .str.strip()
                .str.lower()
                .str.replace(r"\s+", "_", regex=True)
            )

            if "issue_date" in df.columns:
                df["issue_date"] = pd.to_datetime(df["issue_date"], dayfirst=True)

            pdf_files = []

            for _, row in df.iterrows():

                cert_no = get_next_certificate_number()

                issue_date = ""
                if "issue_date" in row and not pd.isna(row["issue_date"]):
                    issue_date = row["issue_date"].strftime("%d-%m-%Y")

                template = Template(custom_content)

                rendered_body = template.render(
                    college_name = row.get("college_name", ""),
                    college_location = row.get("college_location", ""),
                    semester = row.get("semester", ""),
                    course_name = row.get("course_name", ""),
                    reg_id = row.get("reg_id", ""),
                    internship_hours = row.get("internship_hours", ""),
                    internship_program = row.get("internship_program", "")
                )


                context = {
                    "student_name": row.get("student_name", ""),
                    "certificate_body": rendered_body,
                    "certificate_number": cert_no,
                    "place": row.get("place", ""),
                    "issue_date": issue_date,
                    "base_url": f"file:///{BASE_DIR.replace(os.sep,'/')}"
                }

                html = render_template("certificate.html", **context)

                pdf_path = os.path.join(PDF_DIR, f"{cert_no}.pdf")
                os.makedirs(PDF_DIR, exist_ok=True)

                HTML(string=html, base_url=BASE_DIR).write_pdf(pdf_path)

                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute(
                    "INSERT INTO certificates (certificate_number, student_name, pdf_path) VALUES (?, ?, ?)",
                    (cert_no, row.get("student_name", ""), pdf_path)
                )
                conn.commit()
                conn.close()

                pdf_files.append(pdf_path)

            return send_file(
                pdf_files[0],
                as_attachment=True,
                download_name=os.path.basename(pdf_files[0])
            )

        # ================================
        # 2️⃣ SINGLE MODE
        # ================================
        if single_name and custom_content:

            from datetime import datetime

            # read values from form
            raw_date = request.form.get("single_date", "").strip()
            single_place = request.form.get("single_place", "").strip()

            # convert date format
            single_date = ""
            if raw_date:
                single_date = datetime.strptime(raw_date, "%Y-%m-%d").strftime("%d-%m-%Y")

            cert_no = get_next_certificate_number()

            template = Template(custom_content)
            rendered_body = template.render()

            context = {
                "student_name": single_name,
                "certificate_body": rendered_body,
                "certificate_number": cert_no,

                # excel placeholders blank:
                "place": "",
                "issue_date": "",

                # single mode placeholders:
                "single_place": single_place,
                "single_issue_date": single_date,

                "base_url": f"file:///{BASE_DIR.replace(os.sep,'/')}"
            }

            html = render_template("certificate.html", **context)

            pdf_path = os.path.join(PDF_DIR, f"{cert_no}.pdf")
            os.makedirs(PDF_DIR, exist_ok=True)

            HTML(string=html, base_url=BASE_DIR).write_pdf(pdf_path)

            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute(
                "INSERT INTO certificates (certificate_number, student_name, pdf_path) VALUES (?, ?, ?)",
                (cert_no, single_name, pdf_path)
            )
            conn.commit()
            conn.close()

            return send_file(
                pdf_path,
                as_attachment=True,
                download_name=f"{cert_no}.pdf"
            )


        return "Error: Upload Excel OR enter Student Name."


    return render_template("upload.html")


# ---------------- MAIN ----------------
if __name__ == "__main__":
    init_db()
    app.run(debug=True)




################old code################
# import os
# import sqlite3
# import pandas as pd
# import zipfile
# from flask import Flask, render_template, request, send_file
# from weasyprint import HTML
# from jinja2 import Template

# app = Flask(__name__)

# # ---------------- PATHS ----------------
# BASE_DIR = os.path.abspath(os.path.dirname(__file__))
# DB_PATH = os.path.join(BASE_DIR, "certificates.db")
# PDF_DIR = os.path.join(BASE_DIR, "generated", "pdfs")
# ZIP_PATH = os.path.join(BASE_DIR, "generated", "certificates.zip")

# # ---------------- DATABASE ----------------
# def init_db():
#     conn = sqlite3.connect(DB_PATH)
#     c = conn.cursor()
#     c.execute("""
#         CREATE TABLE IF NOT EXISTS certificates (
#             id INTEGER PRIMARY KEY AUTOINCREMENT,
#             certificate_number TEXT,
#             student_name TEXT,
#             pdf_path TEXT
#         )
#     """)
#     conn.commit()
#     conn.close()

# # def get_next_certificate_number():
# #     conn = sqlite3.connect(DB_PATH)
# #     c = conn.cursor()
# #     c.execute(
# #         "SELECT certificate_number FROM certificates ORDER BY id DESC LIMIT 1"
# #     )
# #     row = c.fetchone()
# #     conn.close()

# #     if not row:
# #         return "ACDT-C-25-001"

# #     last_no = int(row[0].split("-")[-1])
# #     return f"ACDT-C-25-{last_no + 1:03d}"

# import re

# def get_next_certificate_number():
#     conn = sqlite3.connect(DB_PATH)
#     c = conn.cursor()
#     c.execute(
#         "SELECT certificate_number FROM certificates ORDER BY id DESC LIMIT 1"
#     )
#     row = c.fetchone()
#     conn.close()

#     # If no previous record
#     if not row:
#         return "ACDT-C-25-001"

#     last_cert_no = row[0]

#     # Extract only last numeric block (001, 195, etc)
#     match = re.search(r"(\d+)$", last_cert_no)

#     if not match:
#         # fallback if something unexpected appears
#         next_num = 1
#     else:
#         next_num = int(match.group(1)) + 1

#     return f"ACDT-C-25-{next_num:03d}"

# # ---------------- ROUTE ----------------
# @app.route("/", methods=["GET", "POST"])
# def upload():
#     if request.method == "POST":
#         excel_file = request.files["excel"]
#         custom_content = request.form["content"]

#         # 1️⃣ Read Excel
#         df = pd.read_excel(excel_file)

#         # 2️⃣ Normalize column names
#         df.columns = (
#             df.columns
#             .astype(str)
#             .str.strip()
#             .str.lower()
#             .str.replace(r"\s+", "_", regex=True)
#         )

#         # 3️⃣ Parse date
#         df["issue_date"] = pd.to_datetime(df["issue_date"], dayfirst=True)

#         pdf_files = []

#         for _, row in df.iterrows():

#             cert_no = get_next_certificate_number()
#             issue_date = row["issue_date"].strftime("%d-%m-%Y")

#             # Render certificate body
#             template = Template(custom_content)
#             rendered_body = template.render(
#                 college_name=row.get("college_name", ""),
#                 college_location=row.get("college_location", ""),
#                 semester=row.get("semester", ""),
#                 course_name=row.get("course_name", ""),
#                 reg_id=row.get("reg_id", ""),
#                 internship_hours=row.get("internship_hours", ""),
#                 internship_program=row.get("internship_program", "")
#             )

#             context = {
#                 "student_name": row["student_name"],
#                 "certificate_body": rendered_body,
#                 "certificate_number": cert_no,
#                 "place": row["place"],
#                 "issue_date": issue_date,
#                 "base_url": f"file:///{BASE_DIR.replace(os.sep, '/')}"
#             }

#             html = render_template("certificate.html", **context)

#             pdf_path = os.path.join(PDF_DIR, f"{cert_no}.pdf")
#             os.makedirs(PDF_DIR, exist_ok=True)

#             HTML(string=html, base_url=BASE_DIR).write_pdf(pdf_path)

#             # Save record
#             conn = sqlite3.connect(DB_PATH)
#             c = conn.cursor()
#             c.execute(
#                 "INSERT INTO certificates (certificate_number, student_name, pdf_path) VALUES (?, ?, ?)",
#                 (cert_no, row["student_name"], pdf_path)
#             )
#             conn.commit()
#             conn.close()

#             pdf_files.append(pdf_path)

#         # ---------------- PDF DOWNLOAD ----------------
#         # ✅ Download FIRST PDF immediately
#         return send_file(
#             pdf_files[0],
#             as_attachment=True,
#             download_name=os.path.basename(pdf_files[0])
#         )

#         # ---------------- ZIP (COMMENTED, NOT REMOVED) ----------------
#         """
#         os.makedirs(os.path.dirname(ZIP_PATH), exist_ok=True)
#         with zipfile.ZipFile(ZIP_PATH, "w") as zipf:
#             for pdf in pdf_files:
#                 zipf.write(pdf, os.path.basename(pdf))
#         return send_file(ZIP_PATH, as_attachment=True)
#         """

#     return render_template("upload.html")

# # ---------------- RUN ----------------
# if __name__ == "__main__":
#     init_db()
#     app.run(debug=True)
