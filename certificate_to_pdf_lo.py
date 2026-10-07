"""Compatibility entry point; PDF export now uses Microsoft Word."""

from certificate_to_pdf import convert_one, main


if __name__ == "__main__":
    raise SystemExit(main())
