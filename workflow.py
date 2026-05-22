from xyt_export import main as login_export
from process_data import process


def main():
    print("=== Step 1: Login & Export ===")
    login_export()

    print("\n=== Step 2: Data Processing ===")
    result = process()

    print("\n=== All Done ===")
    return result


if __name__ == "__main__":
    main()
