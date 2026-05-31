import os
from datetime import datetime
import pandas as pd
from tkinter import Tk, filedialog


def get_creation_date(path):
    try:
        timestamp = os.path.getctime(path) #reusable func , raw timestamp
        return datetime.fromtimestamp(timestamp) #convert raw timestamp
    except Exception:
        return None


def scan_folder(folder_path):
    data = []

    for root, dirs, files in os.walk(folder_path):

        for folder in dirs:
            folder_full_path = os.path.join(root, folder)

            data.append({
                "Name": folder,
                "Type": "Folder",
                "Full Path": folder_full_path,
                "Creation Date": get_creation_date(folder_full_path)
            })


        for file in files:
            file_full_path = os.path.join(root, file)

            data.append({
                "Name": file,
                "Type": "File",
                "Full Path": file_full_path,
                "Creation Date": get_creation_date(file_full_path)
            })

    return data


def save_to_excel(data, selected_folder):
    df = pd.DataFrame(data)

    df = df.sort_values(by="Creation Date")
    
    df["Creation Date"] = df["Creation Date"].dt.strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    output_file = os.path.join(
        selected_folder,
        "folder_scan_report.xlsx"
    )

    df.to_excel(output_file, index=False)

    print(f"\nExcel file saved successfully!")
    print(f"Location: {output_file}")


def choose_folder():
    root = Tk()
    root.withdraw()  

    folder_selected = filedialog.askdirectory(
        title="Select Folder to Scan"
    )

    return folder_selected


def main():
    print("Select a folder to scan...")

    folder_path = choose_folder()

    if not folder_path:
        print("No folder selected.")
        return

    print("\nScanning folders and files...")
    data = scan_folder(folder_path)

    print(f"Found {len(data)} items.")

    print("Generating Excel report...")
    save_to_excel(data, folder_path)

    print("Done!")


if __name__ == "__main__":
    main()