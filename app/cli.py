import argparse
from app.converter import convert_file

def main():
    parser = argparse.ArgumentParser(description="Convert .doc/.docx to Markdown")
    parser.add_argument('input', help='Path to input file (.doc or .docx)')
    parser.add_argument('-o', '--output', help='Output .md file path (default: input with .md)')
    args = parser.parse_args()
    try:
        out = convert_file(args.input, args.output)
        print(f"Converted to {out}")
    except Exception as e:
        print(f"Error: {e}")
        return 1
    return 0

if __name__ == '__main__':
    exit(main())