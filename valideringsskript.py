import os
import re
from pypdf import PdfReader
import sys
from pathlib import Path


from openpyxl import load_workbook


def validate_output(
        input_root,
        output_root,
        log_callback=None
):
    """
    Kontrollerar:

    - alla volymer finns kvar
    - alla PDF:er verkar ha splittrats
    - tomma personnummer i index.xlsx
    """

    report = []

    input_root = Path(input_root)
    output_root = Path(output_root)

    # --------------------------------------------------
    # Volymer
    # --------------------------------------------------

    input_volumes = {
        p.name
        for p in input_root.iterdir()
        if p.is_dir()
    }

    output_volumes = {
        p.name
        for p in output_root.iterdir()
        if p.is_dir()
    }

    missing_volumes = sorted(
        input_volumes - output_volumes
    )

    report.append(
        f"Volymer i input: {len(input_volumes)}"
    )

    report.append(
        f"Volymer i output: {len(output_volumes)}"
    )

    report.append(
        f"Saknade volymer: {len(missing_volumes)}"
    )

    report.append("")

    if missing_volumes:

        report.append(
            "SAKNADE VOLYMER"
        )

        report.extend(missing_volumes)

        report.append("")

    # --------------------------------------------------
    # PDF-split kontroll
    # --------------------------------------------------

    successful_page_checks = 0
    failed_page_checks = []
    successful_pdfs = 0
    failed_pdfs = []

    for pdf in input_root.rglob("*.pdf"):

        rel_dir = pdf.parent.relative_to(
            input_root
        )

        output_dir = (
                output_root
                / rel_dir
        )

        pattern = re.compile(
            re.escape(pdf.stem)
            + r"_\d+\.pdf$",
            re.IGNORECASE
        )

        found = False

        if output_dir.exists():

            for output_pdf in output_dir.glob(
                    "*.pdf"
            ):
                if pattern.match(
                        output_pdf.name
                ):
                    found = True
                    break

        if found:

            successful_pdfs += 1

            try:

                original_pages = len(
                    PdfReader(str(pdf)).pages
                )

                split_pages = 0

                for output_pdf in output_dir.glob("*.pdf"):

                    if pattern.match(output_pdf.name):
                        split_pages += len(
                            PdfReader(
                                str(output_pdf)
                            ).pages
                        )

                if original_pages == split_pages:

                    successful_page_checks += 1

                else:

                    failed_page_checks.append(
                        (
                            str(
                                pdf.relative_to(
                                    input_root
                                )
                            ),
                            original_pages,
                            split_pages
                        )
                    )

            except Exception as e:

                failed_page_checks.append(
                    (
                        str(
                            pdf.relative_to(
                                input_root
                            )
                        ),
                        "ERROR",
                        str(e)
                    )
                )

        else:

            failed_pdfs.append(
                str(
                    pdf.relative_to(
                        input_root
                    )
                )
            )

    report.append(
        f"PDF-filer verifierade: {successful_pdfs}"
    )

    report.append(
        f"PDF-filer med korrekt sidantal: {successful_page_checks}"
    )

    report.append(
        f"PDF-filer utan split: {len(failed_pdfs)}"
    )

    report.append("")

    report.append("")

    report.append(
        "SIDANTALSKONTROLL"
    )

    if failed_page_checks:

        for pdf_name, original_pages, split_pages in failed_page_checks:
            report.append(
                f"{pdf_name} | original={original_pages} | split={split_pages}"
            )

    else:

        report.append(
            "Inga sidavvikelser."
        )

    report.append("")

    # --------------------------------------------------
    # index.xlsx
    # --------------------------------------------------

    index_file = (
            output_root
            / "index.xlsx"
    )

    extracted = 0
    missing = 0
    missing_rows = []

    if index_file.exists():

        wb = load_workbook(
            index_file,
            read_only=True
        )

        ws = wb.active

        headers = [
            str(c.value).strip().lower()
            if c.value
            else ""
            for c in ws[1]
        ]

        try:
            pnr_col = headers.index(
                "personnummer"
            )

        except ValueError:

            report.append(
                "Kunde inte hitta kolumnen personnummer."
            )

            pnr_col = None

        if pnr_col is not None:

            for row_num, row in enumerate(
                    ws.iter_rows(
                        min_row=2,
                        values_only=True
                    ),
                    start=2
            ):

                value = row[pnr_col]

                if (
                        value
                        and str(value).strip()
                ):
                    extracted += 1
                else:
                    missing += 1
                    missing_rows.append(
                        row_num
                    )

        wb.close()

    report.append(
        f"Personnummer extraherade: {extracted}"
    )

    report.append(
        f"Personnummer saknas: {missing}"
    )

    report.append("")

    if missing_rows:

        report.append(
            "RADER UTAN PERSONNUMMER"
        )

        report.extend(
            f"Rad {r}"
            for r in missing_rows
        )

    validation_passed = (
        len(missing_volumes) == 0
        and len(failed_pdfs) == 0
        and len(failed_page_checks) == 0
        and missing == 0
    )

    report.insert(
        0,
        "=" * 60
    )

    report.insert(
        1,
        f"VALIDATION: {'PASSED' if validation_passed else 'FAILED'}"
    )

    report.insert(
        2,
        "=" * 60
    )

    report.insert(
        3,
        ""
    )
### validation file write
    report_file = (
            output_root
            / "output_validation.txt"
    )

    with open(
            report_file,
            "w",
            encoding="utf-8"
    ) as f:

        f.write(
            "\n".join(report)
        )

    if log_callback:
        log_callback(
            f"Validation report created: {report_file}"
        )

    return report_file

if __name__ == "__main__":

    if len(sys.argv) != 3:
        print(
            "Användning: python validation.py <input_mapp> <output_mapp>"
        )
        sys.exit(1)

    input_root = Path(sys.argv[1])
    output_root = Path(sys.argv[2])

    if not input_root.exists():
        print(f"Input-mappen finns inte: {input_root}")
        sys.exit(1)

    if not output_root.exists():
        print(f"Output-mappen finns inte: {output_root}")
        sys.exit(1)

    report_file = validate_output(
        input_root=input_root,
        output_root=output_root
    )

    print(f"Rapport skapad: {report_file}")
