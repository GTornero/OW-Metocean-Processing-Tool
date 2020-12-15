import pandas as pd

import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import sys


class MetoceanData:
    def __init__(self, filepath):
        self.filepath = filepath
        # Initialise a config attribute which will be a dictionary containing all of the configuration options for the report.
        self.config = {}
        workbook = load_workbook(filepath, read_only=True)
        self.parse_config(workbook)
        self.parse_data(filepath)

    def parse_config(self, workbook):
        """parse_config: Parses the 'Config' sheet and stores all configuration parameters in a dictionary.

        Args:
            workbook (openpyxl.workbook.workbook.Workbook): excel workbook containing the configuration settings for the metocean processing.
        """

        # Check if the 'Config' sheet exists in the config file.
        if "Config" in workbook.sheetnames:
            config_sheet = workbook["Config"]
        else:
            sys.exit("No 'Config' worksheet found.")

        print("Parsing configuration...", end="")

        self.config["project"] = config_sheet["D5"].value

        if config_sheet["D7"].value == "left" or "right":
            self.config["bin_type"] = config_sheet["D7"].value
        else:
            self.config["bin_type"] = "left"

        self.config["data_threshold"] = config_sheet["D6"].value / 100
        # Parsing Config of wind data
        if config_sheet["F9"].value == "ON":
            self.config["wind_status"] = True
        else:
            self.config["wind_status"] = False

        if self.config["wind_status"]:
            self.config["wind_source"] = config_sheet["D10"].value
            self.config["wind_projection"] = config_sheet["D11"].value
            self.config["wind_easting"] = config_sheet["D12"].value
            self.config["wind_northing"] = config_sheet["D13"].value
            self.config["hub_weibull_a"] = config_sheet["D14"].value
            self.config["hub_weibull_k"] = config_sheet["D15"].value
            self.config["hub_height"] = config_sheet["D16"].value
            self.config["10m"] = config_sheet["D17"].value
            self.config["wind_bin_size"] = config_sheet["D18"].value
            self.config["wind_sectors"] = config_sheet["D19"].value
        # Parsing config of wave data
        if config_sheet["F21"].value == "ON":
            self.config["wave_status"] = True
        else:
            self.config["wave_status"] = False

        if self.config["wave_status"]:
            self.config["wave_source"] = config_sheet["D22"].value
            self.config["wave_projection"] = config_sheet["D23"].value
            self.config["wave_easting"] = config_sheet["D24"].value
            self.config["wave_northing"] = config_sheet["D25"].value
            self.config["wave_spectral"] = config_sheet["D26"].value
            self.config["peak_enhancement"] = config_sheet["D27"].value
            self.config["wave_height_bin_size"] = config_sheet["D28"].value
            self.config["wave_period_bin_size"] = config_sheet["D29"].value
            self.config["wave_sectors"] = config_sheet["D30"].value
        # Parsing config of current data
        if config_sheet["F32"].value == "ON":
            self.config["current_status"] = True
        else:
            self.config["current_status"] = False

        if self.config["current_status"]:
            self.config["current_source"] = config_sheet["D33"].value
            self.config["current_projection"] = config_sheet["D34"].value
            self.config["current_easting"] = config_sheet["D35"].value
            self.config["current_northing"] = config_sheet["D36"].value
            self.config["current_bin_size"] = config_sheet["D37"].value
            self.config["current_sectors"] = config_sheet["D38"].value
            self.config["current_components"] = config_sheet["D39"].value
        # Parsing config of seawater data
        if config_sheet["F41"].value == "ON":
            self.config["water_status"] = True
        else:
            self.config["water_status"] = False
        if self.config["water_status"]:
            self.config["water_source"] = config_sheet["D42"].value
            self.config["water_projection"] = config_sheet["D43"].value
            self.config["water_easting"] = config_sheet["D44"].value
            self.config["water_northing"] = config_sheet["D45"].value

        print("Parsing configuration complete!")

    def parse_data(self, filepath):
        """parse_data Reads the config/input file using pandas to create a dataframe and storing if in an attribute called 'data''.

        Args:
            filepath (string): filepath of the config/input excel sheet. Used by pandas to read the data into a dataframe.
        """
        columns = []
        # Selecting which wind data columns to include
        if self.config["wind_status"]:
            if self.config["10m"]:
                columns.append("C:J")
            else:
                columns.append("C:F")
        # Selecting which wave data columns to include
        if self.config["wave_status"]:
            if self.config["wave_spectral"]:
                if self.config["peak_enhancement"]:
                    columns.append("K:Y")
                else:
                    columns.append("K:N")
                    columns.append("P:S")
                    columns.append("U:X")
            else:
                if self.config["peak_enhancement"]:
                    columns.append("K:O")
                else:
                    columns.append("K:N")
        # Selecting which current data columns to include
        if self.config["current_status"]:
            if self.config["current_components"]:
                columns.append("Z:AH")
            else:
                columns.append("Z:AB")
        # Selecting seawater columns to include
        if self.config["water_status"]:
            columns.append("AI:AK")

        data = pd.read_excel(
            filepath, sheet_name="Data", header=1, usecols=",".join(columns)
        )

        self.data = data


def main():
    root = tk.Tk()
    root.iconbitmap("OW_logo.ico")
    # Asks the user to select the config file and stores the full path.
    filepath = filedialog.askopenfilename(
        title="Select the metocean configuration file."
    )

    metocean_data = MetoceanData(filepath)
    print(metocean_data.data.head())


if __name__ == "__main__":
    main()
