import sys
from docx2pdf import convert


def main():
    if len(sys.argv) != 3:
        print("Usage: convert_docx_to_pdf.py <input.docx> <output.pdf>")
        return 1

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        convert(input_path, output_path)
    except Exception as exc:
        print(f"Conversion failed: {exc}")
        return 2

    return 0


if __name__ == "__main__":
    sys.exit(main())
