import tkinter as tk
import tkinter.filedialog
import tkinter.simpledialog
import pandas as pd

from company_wise_reports import merge_files
from pivot_table import generate_pivot_table

df_premium_file = pd.DataFrame()
df_claim_file = pd.DataFrame()
df_claim_data_file = pd.DataFrame()
bool_hub = False


class MyGUI:
    def __init__(self):
        self.root = tk.Tk()

        self.root.title("Coinsurance GUI")
        self.root.geometry("575x500")

        self.premium_button = tk.Button(
            self.root,
            text="Premium payable",
            font=("Arial", 18),
            command=open_premium_file,
        )
        #       self.premium_button.pack(padx=20, pady=10)

        self.claim_button = tk.Button(
            self.root,
            text="Claim receivable",
            font=("Arial", 18),
            command=open_claim_file,
        )

        #      self.claim_button.pack(padx=20, pady=10)
        self.claim_data_button = tk.Button(
            self.root,
            text="Claims data",
            font=("Arial", 18),
            command=open_multiple_claim_data_files,
        )

        self.premium_button.grid(row=10, column=0, pady=20)
        self.claim_button.grid(row=10, column=1, pady=20)
        self.claim_data_button.grid(row=10, column=2, pady=20)
        #  self.premium_button.pack(padx=20,pady=10)

#        self.check_bool_hub = tk.BooleanVar()
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
           # command=check_button_clicked,
        )
        self.check_Button.grid(row=15, column=1, pady=20)

        self.generate_reports_button = tk.Button(
            self.root,
            text="Generate reports",
            font=("Arial", 18),
            command=generate_reports,
        )
        #        self.generate_reports_button.pack(padx=20, pady=10)

        self.new_summary_company_report_button = tk.Button(
            self.root,
            text="Generate new summary",
            font=("Arial", 12),
            command=generate_summary,
        )
        self.generate_reports_button.grid(row=20, column=1, pady=50)

        self.new_summary_company_report_button.grid(row=40, column=1, pady=50)
        #       self.new_summary_company_report_button.pack(padx=20, pady=10)
        self.root.mainloop()


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


MyGUI()
