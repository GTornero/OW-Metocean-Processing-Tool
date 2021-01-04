import numpy as np
import time
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

    def print_table(self, ws, row, col):
        # TODO Create a function that can print a scatter table (all data, headers, excel formatting, total columns, etc.)
        # at a specified [row, col] (top left corner) in a specified worksheet object which is passed to the function.
        pass
