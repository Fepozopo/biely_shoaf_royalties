import tkinter.filedialog
from datetime import datetime
from artist import Update_Artist


start_time = datetime.now()


def main():
    while True:
        # Select artist
        answer = input("Do you want to update the royalties? (Y/N): ")

        if answer.lower() == "y":
            # Select royalty and report files
            royalty_file = tkinter.filedialog.askopenfilename(title="Select Royalty File", filetypes=[("Excel", "*.xlsx")])
            report_file = tkinter.filedialog.askopenfilename(title="Select Report File", filetypes=[("Excel", "*.xlsx")])
            
            # Update artist
            print(f"Updating artist on {royalty_file} and {report_file}...")
            Update_Artist(royalty_file, report_file).update_backland()
            Update_Artist(royalty_file, report_file).update_brouitt()
            Update_Artist(royalty_file, report_file).update_carothers()
            Update_Artist(royalty_file, report_file).update_edmon()
            Update_Artist(royalty_file, report_file).update_heath()
            Update_Artist(royalty_file, report_file).update_jensen()
            Update_Artist(royalty_file, report_file).update_konar()
            Update_Artist(royalty_file, report_file).update_metzger()
            Update_Artist(royalty_file, report_file).update_nickel()

        elif answer.lower() == "n":
            break

        else:
            print("Invalid input. Please enter 'Y' or 'N'.")
            continue

    print(f"Updating artist complete. Time elapsed: {datetime.now() - start_time}")


if __name__ == "__main__":
    main()