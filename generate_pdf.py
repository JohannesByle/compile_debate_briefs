import jinja2
import os
from pylatexenc.latexencode import unicode_to_latex
import pandas as pd
from get_briefs import get_briefs, cache_path, path, get_pdf, build_drive
from googleapiclient import http
from datetime import date

latex_jinja_env = jinja2.Environment(
    block_start_string=r'\BLOCK{',
    block_end_string='}',
    variable_start_string=r'\VAR{',
    variable_end_string='}',
    comment_start_string=r'\#{',
    comment_end_string='}',
    line_statement_prefix='%%',
    line_comment_prefix='%#',
    trim_blocks=True,
    autoescape=False,
    loader=jinja2.FileSystemLoader(os.path.abspath(path))
)


def generate_pdf():
    df, drive = get_briefs()

    # Make data latex safe
    cols = ["Title"] + list(df.columns[6:11])
    for col in cols:
        df[col] = df[col].apply(lambda x: unicode_to_latex(x))

    # Generate category list
    categories = [n for n in pd.unique(df[df.columns[6:12]].values.ravel("K")) if isinstance(n, str) and n != "nan"]
    briefs_by_category = {}
    for category in sorted(categories):
        briefs = df[(df[df.columns[6:12]] == category).any(axis=1)]
        briefs_by_category[category] = [{"index": index, "title": brief["Title"]} for index, brief in briefs.iterrows()]

    # Generate flat list
    briefs_sorted = []
    for index, brief in df.sort_values("Title").iterrows():
        file_path = os.path.join(cache_path, brief["file_id"] + ".pdf").replace("\\", "/")
        briefs_sorted.append({"index": index, "title": brief["Title"], "path": file_path, "id": brief["file_id"]})

    briefs = sorted(briefs_sorted, key=lambda item: int(item["index"]))
    latex_string = latex_jinja_env.get_template("template.tex")
    with open(os.path.join(path, "indexed_briefs.tex"), "w") as f:
        f.write(latex_string.render(briefs_by_category=briefs_by_category, briefs_sorted=briefs_sorted, briefs=briefs))

    os.system(r"pdflatex -interaction=nonstopmode " + os.path.join(path, "indexed_briefs.tex"))
    for file in os.listdir(path):
        if file.startswith("indexed_briefs") and not (file.endswith(".pdf") or file.endswith(".tex")):
            os.remove(file)

    # Get list of all files in Google (Fuck Google) Drive
    files = get_pdf(drive_=drive)
    if files:
        drive.files().delete(fileId=files['id']).execute()

    # Upload new file
    file_metadata = {'name': f'Indexed Briefs ({str(date.today())})',
                     'description': 'GlckOayFQgdIdOqRBOL8',
                     "parents": ['1PSgntCxfM-2YidrIjS8hzfzdzoDGv0ze']}
    media = http.MediaFileUpload('indexed_briefs.pdf', mimetype='application/pdf')
    drive.files().create(body=file_metadata, media_body=media).execute()


if __name__ == "__main__":
    generate_pdf()
