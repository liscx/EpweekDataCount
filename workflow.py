from xyt_export import main as login_export
from process_data import process
import sys


def main(mode="auto"):
    print("=== Step 1: Login & Export ===")
    login_export()

    print("\n=== Step 2: Data Processing ===")
    result = process(mode)

    print("\n=== All Done ===")
    return result


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "auto"
    main(mode)
