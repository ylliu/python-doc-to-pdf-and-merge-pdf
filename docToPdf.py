import sys
import os
import json
import subprocess
import time
from pathlib import Path
from tqdm.auto import tqdm
from docx2pdf import convert
word_extensions = [".doc", ".odt", ".rtf", ".docx", ".dotm", ".docm"]
ppt_extensions = [".ppt", ".pptx"]


def windows(paths, keep_active):
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    word = win32com.client.dynamic.Dispatch("Word.Application")
    wdFormatPDF = 17
    if paths["batch"]:
        for ext in word_extensions:
            for docx_filepath in tqdm(sorted(Path(paths["input"]).glob(f"*{ext}"))):
                print(docx_filepath)
                pdf_filepath = Path(paths["output"]) / (
                    str(docx_filepath.stem) + ".pdf"
                )
                print(pdf_filepath)
                convert(docx_filepath, pdf_filepath)
    else:
        pbar = tqdm(total=1)
        input_filepath = Path(paths["input"]).resolve()
        pdf_filepath = Path(paths["output"]).resolve()
        if input_filepath.suffix in word_extensions:
            doc = word.Documents.Open(str(input_filepath))
            doc.SaveAs(str(pdf_filepath), FileFormat=wdFormatPDF)
            doc.Close()
        pbar.update(1)

    if not keep_active:
        word.Quit()


def resolve_paths(input_path, output_path):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve() if output_path else None
    output = {}
    if input_path.is_dir():
        output["batch"] = True
        output["input"] = str(input_path)
        if output_path:
            assert output_path.is_dir()
        else:
            output_path = str(input_path)
        output["output"] = output_path
    else:
        output["batch"] = False
        # assert str(input_path).endswith(".docx")
        output["input"] = str(input_path)
        if output_path and output_path.is_dir():
            output_path = str(output_path / (str(input_path.stem) + ".pdf"))
        elif output_path:
            assert str(output_path).endswith(".pdf")
        else:
            output_path = str(input_path.parent / (str(input_path.stem) + ".pdf"))
        output["output"] = output_path
    return output


def convert(input_path, output_path=None, keep_active=False):
    paths = resolve_paths(input_path, output_path)
    print(paths)
    if sys.platform == "win32":
        return windows(paths, keep_active)
    else:
        raise NotImplementedError(
            "This script is not implemented for linux and macOS as it requires Microsoft Word to be installed"
        )


def main():
    print("Processing...")
    input_path = os.path.abspath("./")
    convert(input_path)
    print("Processed...")


if __name__ == "__main__":
    main()
