import tkinter.filedialog
from datetime import datetime
from artist import update_sales, update_royalty, update_dates, rename_royalty_file


start_time = datetime.now()


def main():
    while True:
        # Ask user if they want to update the royalties
        answer = input("Do you want to update the artist royalties? (Y/N): ")

        if answer.lower() == "y":
            # Select artists file, read it, and split it into a list
            print("Select artists file:")
            artists_file = tkinter.filedialog.askopenfilename(title="Select Artists File", filetypes=[("Text", "*.txt")])
            artists_data = open(artists_file, "r", encoding="utf-8")
            artists_lst = artists_data.read().splitlines()
            artists_data.close()

            # Ensure file is selected
            if not artists_file:
                print("File was not selected. Please try again.")
                continue

            # Sort the artists list alphabetically
            artists = sorted(artists_lst)

            # Select royalty and report files
            print("Select royalty and report files:")
            royalty_file = tkinter.filedialog.askopenfilename(title="Select Royalty File", filetypes=[("Excel", "*.xlsx")])
            report_file = tkinter.filedialog.askopenfilename(title="Select Report File", filetypes=[("Excel", "*.xlsx")])

            # Ensure all files are selected
            if not royalty_file or not report_file:
                print("One or more files were not selected. Please try again.")
                continue

            # Rename royalty file
            new_royalty_file = rename_royalty_file(royalty_file)

            # Update royalty and sales
            update_dates(artists, new_royalty_file, report_file)
            update_sales(artists, new_royalty_file, report_file)
            update_royalty(artists, new_royalty_file)

        elif answer.lower() == "n":
            break

        else:
            print("Invalid input. Please enter 'Y' or 'N'.")
            continue

    print(f"Updating artist royalties complete. Time elapsed: {datetime.now() - start_time}")


if __name__ == "__main__":
    main()