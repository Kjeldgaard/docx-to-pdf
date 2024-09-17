#!/usr/bin/env python

import argparse
from pathlib import Path
from spire.doc import *
from spire.doc.common import *


def main(input: Path, dryrun: bool):
    files = []
    if input.is_dir():
        files = [file for file in input.rglob("*.docx")]
    elif input.is_file():
        files.append(input)

    for file in files:
        out_file_name = file.with_suffix(".pdf")
        print(f"Convert {file} to {out_file_name}")
        if dryrun:
            print(f"Dryrun active, no conversion")
            continue
        document = Document()
        document.LoadFromFile(str(file))
        document.SaveToFile(str(out_file_name), FileFormat.PDF)
        document.Close()


if __name__ == "__main__":
    # Parse arguments
    parser = argparse.ArgumentParser(
        description="Script to convert .docx files to .pdf files",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("input", type=Path, help="Input .docx file to convert to .pdf")
    parser.add_argument(
        "--dryrun", "-d", action="store_true", help="Dryrun, no conversion takes place"
    )

    args = parser.parse_args()
    main(**vars(args))
