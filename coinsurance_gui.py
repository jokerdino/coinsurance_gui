import tkinter as tk
import tkinter.filedialog
import tkinter.simpledialog
import pandas as pd

from company_wise_reports import merge_files
from pivot_table import generate_pivot_table
from split_follower_office_code import split_files
from merge_coinsurance_files import merge_multiple_files

df_premium_file = pd.DataFrame()
df_claim_file = pd.DataFrame()
df_claim_data_file = pd.DataFrame()
bool_hub = False


class MyGUI:
    def __init__(self):
        self.root = tk.Tk()

        self.root.title("Coinsurance GUI")
        self.root.geometry("575x600")

        # menu bar
        self.menubar = tk.Menu(self.root)
        self.root.configure(menu=self.menubar)
        self.filemenu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="File", menu=self.filemenu)
        self.filemenu.add_command(label="About", command=show_version)
        self.filemenu.add_command(label="Quit", command=self.root.destroy)

        self.premium_button = tk.Button(
            self.root,
            text="Premium payable",
            font=("Arial", 18),
            command=open_premium_file,
        )

        self.claim_button = tk.Button(
            self.root,
            text="Claim receivable",
            font=("Arial", 18),
            command=open_claim_file,
        )

        self.claim_data_button = tk.Button(
            self.root,
            text="Claims data",
            font=("Arial", 18),
            command=open_multiple_claim_data_files,
        )

        self.premium_button.grid(row=10, column=0, pady=20)
        self.claim_button.grid(row=10, column=1, pady=20)
        self.claim_data_button.grid(row=10, column=2, pady=20)

        global bool_hub
        bool_hub = tk.BooleanVar()
        self.check_Button = tk.Checkbutton(
            self.root,
            variable=bool_hub,
            onvalue=1,
            offvalue=0,
            height=5,
            width=20,
            text="Only OO entries",
        )
        self.check_Button.grid(row=15, column=1, pady=20)

        self.generate_reports_button = tk.Button(
            self.root,
            text="Generate reports",
            font=("Arial", 18),
            command=generate_reports,
        )

        self.new_summary_company_report_button = tk.Button(
            self.root,
            text="Generate new summary",
            font=("Arial", 12),
            command=generate_summary,
        )
        self.generate_reports_button.grid(row=20, column=1, pady=50)


        self.split_follower_office_code_button = tk.Button(
            self.root,
            text="Split follower office code",
            font=("Arial", 12),
            command=split_follower_office_code,
        )



        self.merge_files_button = tk.Button(
            self.root,
            text="Merge files",
            font=("Arial", 12),
            command=merge_multiple_files_button,
        )

        self.new_summary_company_report_button.grid(row=35, column=0, pady=50)
        self.split_follower_office_code_button.grid(row=35, column=1, pady=50)
        self.merge_files_button.grid(row=35, column=2, pady=50)
        self.root.mainloop()

def show_version():
    tk.messagebox.showinfo(
            title="Version", message=(f"Current version: 0.2")
    )


def open_premium_file():
    premium_file = tk.filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])

    global df_premium_file
    df_premium_file = pd.read_csv(
        premium_file, converters={"TXT_FOLLOWER_OFF_CD_CODE": str}
    )
    tk.messagebox.showinfo(
        title="Message", message=(f"Premium file {premium_file} has been selected.")
    )


def open_claim_file():
    claim_file = tk.filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])

    global df_claim_file
    df_claim_file = pd.read_csv(
        claim_file,
        converters={"TXT_LEADER_OFFICE_CODE": str, "TXT_FOLLOWER_OFFICE_CODE": str},
    )

    tk.messagebox.showinfo(
        title="Message", message=(f"Claim file {claim_file} has been selected.")
    )


def open_multiple_claim_data_files():
    claim_data_files = tk.filedialog.askopenfilenames(
        filetypes=[("CSV files", "*.csv")]
    )
    global df_claim_data_file

    list_claim_data_files = []
    for file in claim_data_files:
        df_claim_data = pd.read_csv(file, converters={"Office Code": str})
        list_claim_data_files.append(df_claim_data)
    df_claim_data_file = pd.concat(list_claim_data_files)

    tk.messagebox.showinfo(
        title="Message",
        message=(f"Claim data files {claim_data_files} has been selected."),
    )

def generate_reports():
    path_string_wip = tk.simpledialog.askstring(
        title="Enter folder name", prompt="Enter folder name"
    )

    merge_files(
        df_premium_file, df_claim_file, df_claim_data_file, path_string_wip, bool_hub.get()
    )
    tk.messagebox.showinfo(title="Message", message="Reports have been generated.")


def generate_summary():
    current_company_file = tk.filedialog.askopenfilename(
        filetypes=[("Excel file", "*.xlsx")]
    )
    company_name = current_company_file.split(".", 1)[0]

    generate_pivot_table(current_company_file, company_name, None)
    tk.messagebox.showinfo(
        title="Message", message=(f"Summary has been generated for {company_name}.")
    )

def split_follower_office_code():
    current_company_file = tk.filedialog.askopenfilename(
        filetypes=[("Excel file", "*.xlsx")]
    )
    company_name = current_company_file.split(".", 1)[0]

    follower_office_code = tk.simpledialog.askstring(
        title="Enter follower office code", prompt="Enter follower office code"
    )

    split_files(current_company_file, follower_office_code)

    tk.messagebox.showinfo(
        title="Message", message=(f"Statement has been split for {company_name}: {follower_office_code}.")
    )

def merge_multiple_files_button():

    multiple_files = tk.filedialog.askopenfilenames(
        filetypes=[("Excel files", "*.xlsx")]
    )
    new_file_name = tk.simpledialog.askstring(
        title="Enter new file name", prompt="Enter new file name"
    )

    merge_multiple_files(multiple_files, new_file_name.split(".", 1)[0])

    tk.messagebox.showinfo(
        title="Message", message=(f"{new_file_name}.xlsx has been created.")
    )

if __name__ == "__main__":
    MyGUI()
