import tkinter as tk
from tkinter import filedialog
import time

import xlsxwriter

from metocean_data import MetoceanData

# from scatter import Scatter    # dont need this
from scatter_report import print_scatter_report
from NSS import NSS


def main():
    # Inisialise the tkinter interface
    root = tk.Tk()
    root.iconbitmap("OW_logo.ico")
    root.withdraw()
    # Asks the user to select the config file and stores the full path.
    config_filepath = filedialog.askopenfilename(
        title="Select the metocean configuration file.",
        filetypes=[("Excel Files", "*.xlsx")],
    )
    # Create the MetoceanData oject that will hold all of the data and configuration setup.
    metocean_data = MetoceanData(config_filepath)
    print(metocean_data.data.head())

    # ---------------------------------------------------------------------------------------------
    # -----------------------------------Creating the NSS Table Report-----------------------------
    # ---------------------------------------------------------------------------------------------

    # Method for taking mean or median within bin to be implemented
    if (
        metocean_data.config["nss_report"]
        & metocean_data.config["wind_status"]
        & metocean_data.config["wave_status"]
    ):
        NSS_tables = NSS.NSS(metocean_data)

    # ---------------------------------------------------------------------------------------------
    # ---------------------------------Creating the Scatter Table Report---------------------------
    # ---------------------------------------------------------------------------------------------
    if metocean_data.config["scatter_report"]:
        print_scatter_report(metocean_data)


if __name__ == "__main__":
    main()