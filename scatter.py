import numpy as np
import pandas as pd

import time
import xlsxwriter
from tqdm import tqdm


class Scatter:
    """Class to represent a scatter table."""

    def __init__(
        self, met_data, variables, keys=[False, False], x_filt=False, y_filt=False
    ):
        """__init__ Initialises the Scatter class.

        Args:
            met_data (MetoceanData): MetoceanData object to extract statistics from.
            variables (list): List of strings. Each string must correspond to a key of the met_data dataframe.
                First variable will be plotted on the horizontal axis, second variable will be plotted on the vertical axis.
            keys (list, optional): List of strings. Each string must correspond to a key of the met_data dataframe. Defaults to [False, False].
            x_filt (int or float, optional): Value by which to filter the first key in keys. Defaults to False.
            y_filt (int or float, optional): Value by which to filter the second key in keys. Defaults to False.
        """
        self.samples = len(met_data.data)  # Total number of samples
        self.x_var = variables[0]  # Key for the horizontal variable
        self.y_var = variables[1]  # Key for the vertical variable
        self.x_key = keys[0]  # Key for the filter variable of the horizontal axis
        self.y_key = keys[1]  # Key for thefilter varaible of the  vertical axis
        self.x_filt = x_filt  # Sector number of the horizontal varaible
        self.y_filt = y_filt  # Sector number of the vertical variable
        self.bin_type = met_data.config["bin_type"]  # Variable bin discretisation logic

        # Check if the user has set sectors for both variables to filter by
        if self.x_key and self.y_key and self.x_filt and self.y_filt:
            # Create a filter so only rows with both variables are filtered per their corresponding sector
            filt = (met_data.data[self.x_key] == self.x_filt) & (
                met_data.data[self.y_key] == self.y_filt
            )
            # Create a reduced dataframe only of the 2 filtered varaibles
            temp_data = met_data.data[filt].loc[:, [self.x_var, self.y_var]]
        elif self.x_key and self.x_filt:
            filt = met_data.data[self.x_key] == self.x_filt
            temp_data = met_data.data[filt].loc[:, [self.x_var, self.y_var]]
        elif self.y_key and self.y_filt:
            filt = met_data.data[self.y_key] == self.y_filt
            temp_data = met_data.data[filt].loc[:, [self.x_var, self.y_var]]
        else:
            temp_data = met_data.data.loc[:, [self.x_var, self.y_var]]

        if self.x_var in ["WnD_sectors", "WnD_10_sectors"]:
            self.x_bins = np.arange(met_data.config["wind_sectors"]) + 1
        elif self.x_var in ["WvD_sectors", "WvD_W_sectors", "WvD_S_sectors"]:
            self.x_bins = np.arange(met_data.config["wave_sectors"]) + 1
        elif self.x_var in ["CD_sectors", "CD_Tid_sectors", "CD_Res_sectors"]:
            self.x_bins = np.arange(met_data.config["current_sectors"]) + 1
        else:
            self.x_bins = met_data.bins[self.x_var.replace("_bins", "")]

        if self.y_var in ["WnD_sectors", "WnD_10_sectors"]:
            self.y_bins = np.arange(met_data.config["wind_sectors"]) + 1
        elif self.y_var in ["WvD_sectors", "WvD_W_sectors", "WvD_S_sectors"]:
            self.y_bins = np.arange(met_data.config["wave_sectors"]) + 1
        elif self.y_var in ["CD_sectors", "CD_Tid_sectors", "CD_Res_sectors"]:
            self.y_bins = np.arange(met_data.config["current_sectors"]) + 1
        else:
            self.y_bins = met_data.bins[self.y_var.replace("_bins", "")]

        # Initialise the empty table and fill with nan
        self.table = np.empty([len(self.y_bins), len(self.x_bins)]) * np.nan

        start = time.perf_counter()
        # Loop through the entire table and calculate the probability
        for row in tqdm(
            range(len(self.y_bins)),
            leave=False,
            desc=f"Calculating table: {self.y_var} Vs. {self.x_var}",
        ):
            for col in tqdm(
                range(len(self.x_bins)),
                leave=False,
                desc=f"Row {row} of {len(self.y_bins)}",
            ):
                prob = (
                    len(
                        temp_data[
                            (temp_data[self.x_var] == self.x_bins[col].round(4))
                            & (temp_data[self.y_var] == self.y_bins[row].round(4))
                        ]
                    )
                    / self.samples
                )
                if prob != 0:
                    self.table[row][col] = prob
        finish = time.perf_counter()
        elapsed = finish - start

        if self.x_key and self.x_filt and self.y_key and self.y_filt:
            print(
                f"Table {self.y_var} Vs. {self.x_var} [{self.x_key} = {self.x_filt}] [{self.y_key} = {self.y_filt}] complete! Time taken: {round(elapsed, 2)} seconds."
            )
        elif self.x_key and self.x_filt:
            print(
                f"Table {self.y_var} Vs. {self.x_var} [{self.x_key} = {self.x_filt}] complete! Time taken: {round(elapsed, 2)} seconds."
            )
        elif self.y_key and self.y_filt:
            print(
                f"Table {self.y_var} Vs. {self.x_var} [{self.y_key} = {self.y_filt}] complete! Time taken: {round(elapsed, 2)} seconds."
            )
        else:
            print(
                f"Table {self.y_var} Vs. {self.x_var} complete! Time taken: {round(elapsed, 2)} seconds."
            )

    def print_table(self, workbook, worksheet, row=0, col=0):
        """print_table Function to print the scatter table into an excel file with all the pretty formatting.

        Args:
            workbook (xlsxwriter.Workbook): xlsxwriter library Workbook class. Excel workbook at which to print the scatter table.
            worksheet (xlsxwriter.Worksheet): xlsxwriter library worksheet class. Excel sheet at which to print the sactter table.
            row (int, optional): Zero-indexed row number in the excel sheet to place the table. Refers to the upper-left. Defaults to 0.
            col (int, optional): Zero-indexed column number in the excel sheet to place the table. Refers to the upper-left. Defaults to 0.. Defaults to 0.
        """

        VAR_TITLES = {
            "WS_bins": "Wind Speed @ Hub Height, [m/s]",
            "WnD_sectors": "Wind Direction @ Hub Height, [degN]",
            "WS_10_bins": "Wind Speed @ 10m MSL, [m/s]",
            "WnD_10_sectors": "Wind Direction @ 10m MSL, [degN]",
            "Hs_bins": "Significant Wave Height (Totalsea), Hm0 [m]",
            "Tp_bins": "Peak Wave Period (Totalsea), Tp [s]",
            "Tz_bins": "Zero-Crossing Period (Totalsea), Tz [s]",
            "WvD_sectors": "Mean Wave Direction (Totalsea), [degN]",
            "Hs_W_bins": "Significant Wave height (Windsea), Hm0 [m]",
            "Tp_W_bins": "Peak Wave Period (Windsea), Tp [s]",
            "Tz_W_bins": "Zero-Crossing Wave Period (Windsea), Tz [s]",
            "WvD_W_sectors": "Mean Wave Direction (Windsea), [degN]",
            "Hs_S_bins": "Significant Wave Height (Swell), Hm0 [m]",
            "Tp_S_bins": "Peak Wave Period (Swell), Tp [s]",
            "Tz_S_bins": "Zero-Crossing Wave Period (Swell), Tz [s]",
            "WvD_S_sectors": "Mean Wave Direction (Swell), [degN]",
            "SV_bins": "Current Surface Speed (Total), [m/s]",
            "DaV_bins": "Current Depth Averaged Speed (Total), [m/s]",
            "CD_sectors": "Mean Current Direction (Total), [DegN, going]",
            "SV_Tid_bins": "Current Surface Speed (Tidal), [m/s]",
            "DaV_Tid_bins": "Current Depth Averaged Speed (Tidal), [m/s]",
            "CD_Tid_sectors": "Mean Current Direction (Tidal), [degN, going]",
            "SV_Res_bins": "Current Surface Speed (Residual), [m/s]",
            "DaV_Res_bins": "Current Depth Averaged Speed (Residual), [m/s]",
            "CD_Res_sectors": "Mean Current Direction (Residual), [degN, going]",
        }
        # Creates lists of the y lower and upper bounds
        # If the y variable is direction sectors
        if "sectors" in self.y_var:
            n_sect = len(self.y_bins)
            sector_width = 360 / n_sect
            y_lower_bound = []
            y_upper_bound = []
            for i in range(n_sect):
                # The first sector has different logic
                if i == 0:
                    y_lower_bound.append(360 - sector_width / 2)
                    y_upper_bound.append(sector_width / 2)
                else:
                    y_lower_bound.append((sector_width / 2) + ((i - 1) * sector_width))
                    y_upper_bound.append((sector_width / 2) + (i * sector_width))
        # For all other non-direction variables
        else:
            # Find step between bins
            step = self.y_bins[1] - self.y_bins[0]
            # List comprehension for the lower and upper limit lists
            y_lower_bound = [x - step / 2 for x in self.y_bins]
            y_upper_bound = [x + step / 2 for x in self.y_bins]

        # Creates lists for the x lower and upper bounds
        # If the x variable is direction sectors
        if "sectors" in self.x_var:
            n_sect = len(self.x_bins)
            sector_width = 360 / n_sect
            x_lower_bound = []
            x_upper_bound = []
            for i in range(n_sect):
                # The first sector has different logic
                if i == 0:
                    x_lower_bound.append(360 - sector_width / 2)
                    x_upper_bound.append(sector_width / 2)
                else:
                    x_lower_bound.append((sector_width / 2) + ((i - 1) * sector_width))
                    x_upper_bound.append((sector_width / 2) + (i * sector_width))
        # For all other non-direction variables
        else:
            # Find step between bins
            step = self.x_bins[1] - self.x_bins[0]
            # List comprehension for the lower and upper limit lists
            x_lower_bound = [x - step / 2 for x in self.x_bins]
            x_upper_bound = [x + step / 2 for x in self.x_bins]

        # Define formats of different parts of the table
        # Table header format
        header_format = workbook.add_format(
            {
                "bold": True,
                "border": 2,
                "font_color": "#FFFFFF",
                "bg_color": "072B31",
                "align": "center",
            }
        )
        # y Header Format
        y_header_format = workbook.add_format(
            {
                "border": 2,
                "bold": True,
                "valign": "vcenter",
                "align": "center",
                "bg_color": "C0C0C0",
                "rotation": 90,
            }
        )
        # X Header Format
        x_header_format = workbook.add_format(
            {"border": 2, "bold": True, "align": "center", "bg_color": "C0C0C0"}
        )
        # Upper and lower bound format
        bounds_format = workbook.add_format(
            {"border": 1, "align": "center", "bg_color": "C0C0C0"}
        )
        # Main data format
        data_format = workbook.add_format({"border": 1, "align": "center"})

        # Create table header text
        # If both x and y filter are applied
        if self.x_filt and self.y_filt:
            header_text = f"{VAR_TITLES[self.x_var]} Vs. {VAR_TITLES[self.y_var]}. {self.x_key} = {self.x_filt}, {self.y_key} = {self.y_filt}."
        # If only ther x filter is applied
        elif self.x_filt:
            header_text = f"{VAR_TITLES[self.x_var]} Vs. {VAR_TITLES[self.y_var]}. {self.x_key} = {self.x_filt}."
        # If only the y filter is applied
        elif self.y_filt:
            header_text = f"{VAR_TITLES[self.x_var]} Vs. {VAR_TITLES[self.y_var]}. {self.y_key} = {self.y_filt}."
        # If no filters are applied
        else:
            header_text = f"{VAR_TITLES[self.x_var]} Vs. {VAR_TITLES[self.y_var]}"

        # Create table header merged range
        worksheet.merge_range(
            row,
            col,
            row,
            col + self.table.shape[1] + 3,
            header_text,
            header_format,
        )

        # Create Y header merged range
        worksheet.merge_range(
            row + 2,
            col,
            row + 4 + self.table.shape[0],
            col,
            f"{VAR_TITLES[self.y_var]}",
            y_header_format,
        )

        # Adjust columns widths to fit contents
        worksheet.set_column(col, col, 3)
        worksheet.set_column(col + 1, col + 2, 10)
        # Create x header merged range
        worksheet.merge_range(
            row + 1,
            col + 1,
            row + 1,
            col + 3 + self.table.shape[1],
            f"{VAR_TITLES[self.x_var]}",
            x_header_format,
        )

        # Fill in that awkward square between x and y headers
        worksheet.write(
            row + 1, col, None, workbook.add_format({"bg_color": "C0C0C0", "border": 2})
        )

        # Fill in the Lower and Upper headers for the bins
        if self.bin_type == "left":
            lower_msg = "Lower (>=)"
            upper_msg = "Upper (<)"
        else:
            lower_msg = "Lower (>)"
            upper_msg = "Upper (<=)"

        worksheet.write_string(
            row + 2,
            col + 1,
            lower_msg,
            workbook.add_format(
                {"bg_color": "C0C0C0", "border": 1, "align": "center", "bold": True}
            ),
        )
        worksheet.write(
            row + 2,
            col + 2,
            None,
            workbook.add_format({"bg_color": "C0C0C0", "border": 1}),
        )
        worksheet.write(
            row + 3,
            col + 1,
            None,
            workbook.add_format({"bg_color": "C0C0C0", "border": 1}),
        )
        worksheet.write_string(
            row + 3,
            col + 2,
            upper_msg,
            workbook.add_format(
                {"bg_color": "C0C0C0", "border": 1, "align": "center", "bold": True}
            ),
        )

        # Prints the contents of the table and formats cells
        for row_num, line in enumerate(self.table):
            for col_num, data in enumerate(line):
                if np.isnan(data):
                    worksheet.write_string(
                        row + 4 + row_num, col + 3 + col_num, "NaN", data_format
                    )
                else:
                    worksheet.write_number(
                        row + 4 + row_num, col + 3 + col_num, data, data_format
                    )
        # Applied conditional formatting to the main table body
        worksheet.conditional_format(
            row + 4,
            col + 3,
            row + 3 + self.table.shape[0],
            col + 2 + self.table.shape[1],
            {"type": "3_color_scale"},
        )

        # Print the upper and lower bounds for the x and y variables
        for i in range(len(x_lower_bound)):
            worksheet.write_number(
                row + 2, col + 3 + i, x_lower_bound[i], bounds_format
            )
            worksheet.write_number(
                row + 3, col + 3 + i, x_upper_bound[i], bounds_format
            )

        for i in range(len(y_lower_bound)):
            worksheet.write_number(
                row + 4 + i, col + 1, y_lower_bound[i], bounds_format
            )
            worksheet.write_number(
                row + 4 + i, col + 2, y_upper_bound[i], bounds_format
            )

        # Print Sum headers at the bottom and right of the main table
        # SUM header cell format
        sum_format = workbook.add_format(
            {
                "bold": True,
                "border": 2,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "C0C0C0",
            }
        )
        # Print bottom SUM header
        worksheet.merge_range(
            row + 4 + self.table.shape[0],
            col + 1,
            row + 4 + self.table.shape[0],
            col + 2,
            "SUM",
            sum_format,
        )
        # Print right SUM header
        worksheet.merge_range(
            row + 2,
            col + 3 + self.table.shape[1],
            row + 3,
            col + 3 + self.table.shape[1],
            "SUM",
            sum_format,
        )

        # Calculate Row and Column sum totals
        col_totals = np.nansum(self.table, axis=0)
        row_totals = np.nansum(self.table, axis=1)

        # Print columns sum totals.
        for i, total in enumerate(col_totals):
            worksheet.write_number(
                row + 4 + self.table.shape[0],
                col + 3 + i,
                total,
                workbook.add_format(
                    {
                        "bold": True,
                        "align": "center",
                        "border": 1,
                        "bottom": 2,
                        "top": 2,
                    }
                ),
            )

        worksheet.conditional_format(
            row + 4 + self.table.shape[0],
            col + 3,
            row + 4 + self.table.shape[0],
            col + 2 + self.table.shape[1],
            {"type": "3_color_scale"},
        )

        # Print row sum totals
        for i, total in enumerate(row_totals):
            worksheet.write_number(
                row + 4 + i,
                col + 3 + self.table.shape[1],
                total,
                workbook.add_format(
                    {
                        "bold": True,
                        "align": "center",
                        "border": 1,
                        "right": 2,
                        "left": 2,
                    }
                ),
            )

        worksheet.conditional_format(
            row + 4,
            col + 3 + self.table.shape[1],
            row + 3 + self.table.shape[0],
            col + 3 + self.table.shape[1],
            {"type": "3_color_scale"},
        )

        # Print full table sum total
        worksheet.write_number(
            row + 4 + self.table.shape[0],
            col + 3 + self.table.shape[1],
            np.nansum(self.table),
            workbook.add_format({"bold": True, "border": 2, "align": "center"}),
        )
