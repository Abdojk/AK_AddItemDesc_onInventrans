"""main.py — CLI entry point for the SOW01 tool."""

import argparse
import sys
from pathlib import Path

from extractor import extract_text
from analyser import analyse_sow
from exporter import export_results


def parse_args() -> argparse.Namespace:
    """Parse and return command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "SOW01 — Automated Scope of Work review for D365 Finance & Operations. "
            "Extracts text from a .docx SOW document, sends it to Claude for "
            "assessment, and exports the results."
        ),
        epilog="Example: python main.py --input sow.docx --format both --verbose",
    )
    parser.add_argument(
        "--input", "-i",
        type=Path,
        required=True,
        help="Path to the .docx SOW file to analyse.",
    )
    parser.add_argument(
        "--format", "-f",
        choices=["md", "docx", "both", "none"],
        default="none",
        help="Output format: md, docx, both, or none (default: none — stdout only).",
    )
    parser.add_argument(
        "--output", "-o",
        type=Path,
        default=Path("./output/"),
        help="Output directory (default: ./output/).",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Print extracted text and diagnostics before the API call.",
    )
    return parser.parse_args()


def main() -> None:
    """Orchestrate the SOW01 pipeline: extract → analyse → export."""
    args = parse_args()

    # Validate input file
    if not args.input.exists():
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)
    if args.input.suffix.lower() != ".docx":
        print(f"Error: Expected a .docx file, got '{args.input.suffix}'.", file=sys.stderr)
        sys.exit(1)

    if args.verbose:
        print(f"SOW01 — Processing: {args.input}", file=sys.stderr)
        print(f"Output format: {args.format}", file=sys.stderr)
        print(f"Output directory: {args.output}", file=sys.stderr)

    # Step 1: Extract
    extracted = extract_text(args.input, verbose=args.verbose)

    if args.verbose:
        print("\n--- Extracted Text ---", file=sys.stderr)
        print(extracted, file=sys.stderr)
        print("--- End Extracted Text ---\n", file=sys.stderr)

    # Step 2: Analyse
    assessment = analyse_sow(extracted, verbose=args.verbose)

    # Step 3: Export
    export_results(
        assessment=assessment,
        source_filename=args.input.stem,
        fmt=args.format,
        output_dir=args.output,
        verbose=args.verbose,
    )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        sys.exit(130)
