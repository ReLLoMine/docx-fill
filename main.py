import numpy as np
from docx import Document
import pandas as pd
from jinja2 import Environment, BaseLoader
import argparse
import os


def fill(string, **kwargs):
    template = Environment(loader=BaseLoader()).from_string(string)
    return template.render(**kwargs)


def main():
    arg_parser = argparse.ArgumentParser(description="Docx fill app")
    arg_parser.add_argument("docx", type=str, help="docx file path")
    arg_parser.add_argument("table", type=str, help="table path (.xlsx)")
    arg_parser.add_argument("--path", type=str, default="", help="path where to save all files")
    arg_parser.add_argument("--save_failed", type=bool, default=False, help="to export failed rows")

    args = arg_parser.parse_args()

    docx_path = args.docx
    table_path = args.table

    table = pd.read_excel(table_path)
    table_copy = None
    docx_file_idx = 0

    if args.save_failed:
        table_copy = table.copy()
    table_dict = table.to_dict()

    idx_to_drop = []

    if not os.path.exists(args.path):
        os.makedirs(args.path, exist_ok=True)

    for idx in range(len(table)):
        docx_file = Document(docx_path)

        def f(pair):
            if pair[1][idx] is np.nan:
                raise ValueError
            return pair[0], pair[1][idx]

        try:
            kwargs = {a: b for a, b in map(f, table_dict.items())}
        except ValueError:
            continue

        for paragraph in docx_file.paragraphs:
            for run in paragraph.runs:
                run.text = fill(run.text, **kwargs)

        new_doc_path = fill(os.path.join(args.path, docx_path), **kwargs)

        if os.path.exists(new_doc_path):
            sep = new_doc_path.split(".")
            new_doc_path = "".join(sep[:-1]) + f" ({(docx_file_idx := docx_file_idx + 1)})." + sep[-1]

        docx_file.save(new_doc_path)

        idx_to_drop.append(idx)

    if args.save_failed:
        table_copy = table_copy.drop(idx_to_drop)
    if args.save_failed and len(table_copy):
        table_copy.to_excel("failed_rows.xlsx")


if __name__ == '__main__':
    main()
