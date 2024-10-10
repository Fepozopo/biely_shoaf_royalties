import tkinter.filedialog
from datetime import datetime
from artist import update_artist


start_time = datetime.now()


def main():
    while True:
        # Ask user if they want to update the royalties
        answer = input("Do you want to update the royalties? (Y/N): ")

        if answer.lower() == "y":
            artists = ["BACKLAND", "BROUITT", "CAROTHERS", "EDMON", "HEATH", "JENSEN", "KONAR", "METZGER", "NICKEL", "PARISI", "PARKHURST", "RENFRO"]

            # Select royalty and report files
            royalty_file = tkinter.filedialog.askopenfilename(title="Select Royalty File", filetypes=[("Excel", "*.xlsx")])
            report_file = tkinter.filedialog.askopenfilename(title="Select Report File", filetypes=[("Excel", "*.xlsx")])

            # Ensure all files are selected
            if not royalty_file or not report_file:
                print("One or more files were not selected. Please try again.")
                continue
            
            # Update artist
            update_artist(artists, royalty_file, report_file)

        elif answer.lower() == "n":
            break

        else:
            print("Invalid input. Please enter 'Y' or 'N'.")
            continue

    print(f"Updating artist complete. Time elapsed: {datetime.now() - start_time}")


if __name__ == "__main__":
    main()