import pandas as pd
import numpy as np
import os
import openpyxl
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScale, FormatObject
from openpyxl.styles import NamedStyle, Border, Color, Font, Alignment, PatternFill, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl import utils

class NSS():    
    """ A class to calculate and print NSS tables from an instance of the MetoceanData object."""

    # Initialise the NSS object using an instance of the MetoceanData object
    def __init__(self, metocean_data):
        # Initialise attributes of NSS object by taking informtation from the metocean_data instance
        self.set_up(metocean_data)
        # Select the relevant data from the metocean_data.data attribute
        self.parse_data(metocean_data)
        # Use the selected data to calculate the NSS tables
        self.get_NSS_tables()
        # Print the NSS tables to excel files
        self.print_NSS_tables()

    def set_up(self, metocean_data):
        """set_up: [Initialises the attributes of NSS from information contained in the MetoceanData object]

        Args:
            metoecan_data: [An instance of the MetoceanData object]
        """
        print("Calculating NSS tables...", end="")
        # doesnt check for wind and wave status bc shouldn't be called if they're FALSE
        self.NSectors_wind = metocean_data.config["wind_sectors"]
        self.WS_bins_list = metocean_data.bins["WS"] 
        self.WS_bin_size = metocean_data.config["wind_bin_size"]
        self.WS_HH = metocean_data.config["hub_height"]
        self.NSectors_wave = metocean_data.config["wave_sectors"]
        self.peak_enhancement = metocean_data.config["peak_enhancement"]
        self.derive_peak_enhancement = metocean_data.config["derive_peak_enhancement"]
        self.method = metocean_data.config["method"]
        self.wave_spectral = metocean_data.config["wave_spectral"]
        self.Total_Count = metocean_data.data.shape[0]

        # Create empty data attribute where to store wind and wave data conviniently. 
        # Create empty tables attribute of the right size to populate afterwards
        self.Total_data = []
        self.Total_tables = np.empty((self.NSectors_wind + 1,self.NSectors_wave + 1,self.WS_bins_list.size,4))
        # If Wind and Swell wave data is included in the MetoceanData object, create data and tables attributes for them too 
        if self.wave_spectral: 
            self.Wind_data, self.Swell_data = [],[]
            self.Wind_tables = np.empty((self.NSectors_wind + 1,self.NSectors_wave + 1,self.WS_bins_list.size,4))
            self.Swell_tables = np.empty((1,self.NSectors_wave + 1,self.WS_bins_list.size,4)) # Swell sea not impacted by Wind Direction

    def parse_data(self, metocean_data):
        """parse_data: [Populates the data attributes with wind and wave data taken from the MetoceanData object]

        Args:
            metoecan_data: [An instance of the MetoceanData object]
        """
        # Total_data attribute is always present
        if self.peak_enhancement == False and self.derive_peak_enhancement == False:
            self.Total_data = metocean_data.data.loc[:,("WS_bins","WnD_sectors","WvD_sectors","Hs","Tp")]
            self.Total_data["G"] = np.NAN
        else:
           self.Total_data = metocean_data.data.loc[:,("WS_bins","WnD_sectors","WvD_sectors","Hs","Tp","G")]

        # If Wind and Swell data are to be included, populate their respective attributes 
        # and rename their variables for convinient handling 
        if self.wave_spectral:
            if self.peak_enhancement == False and self.derive_peak_enhancement == False:
                self.Wind_data = metocean_data.data.loc[:,("WS_bins","WnD_sectors","WvD_W_sectors","Hs_W","Tp_W")]
                self.Wind_data["G_W"] = np.NAN
                self.Swell_data = metocean_data.data.loc[:,("WS_bins","WnD_sectors","WvD_S_sectors","Hs_S","Tp_S")]
                self.Swell_data["G_S"] = np.NAN
            else:
                self.Wind_data = metocean_data.data.loc[:,("WS_bins","WnD_sectors","WvD_W_sectors","Hs_W","Tp_W","G_W")]
                self.Swell_data = metocean_data.data.loc[:,("WS_bins","WnD_sectors","WvD_S_sectors","Hs_S","Tp_S","G_S")]
            self.Wind_data.rename(
                columns={"WvD_W_sectors": "WvD_sectors","Hs_W": "Hs","Tp_W":"Tp","G_W":"G"}, inplace=True)
            self.Swell_data.rename(
                columns={"WvD_S_sectors": "WvD_sectors","Hs_S": "Hs","Tp_S":"Tp","G_S":"G"}, inplace=True)

    def get_NSS_tables(self):
        """get_NSS_tables: [Populates the tables attributes]

           Tables are uniform in size and containt 4 dimensions, for:
            1. Wind Direction Sectors
            2. Wave Direction Sectors
            3. Wind Speed bins. Empty wind speed bins are populated with NaNs
            4. Hs, Tp, Peak enhancement factor and Probability of ocurrence

        """       
        # Calculate tables for NSS Total Sea and populate NSS.Total_tables attribute
        for WnSector in range(0,self.NSectors_wind + 1):
            for WvSector in range(0,self.NSectors_wave + 1):
                if WnSector == 0: # OMNIDIRECTIONAL
                    if WvSector == 0: # OMNIDIRECTIONAL
                        self.Total_tables[WnSector][WvSector] = self.calc_table(self.Total_data)
                    else:
                        df_temp =  self.Total_data[self.Total_data.WvD_sectors == WvSector]
                        self.Total_tables[WnSector][WvSector] = self.calc_table(df_temp)
                else:
                    df_temp = self.Total_data[
                        (self.Total_data.WvD_sectors == WvSector) & (self.Total_data.WnD_sectors == WnSector)]
                    self.Total_tables[WnSector][WvSector] = self.calc_table(df_temp)        
                
        # Calculate tables for NSS Wind and Swell Sea
        if self.wave_spectral:
            for WnSector in range(0,self.NSectors_wind + 1):
                for WvSector in range(0,self.NSectors_wave + 1):
                    if WnSector == 0: 
                        if WvSector == 0: 
                            self.Swell_tables[WnSector][WvSector] = self.calc_table(self.Swell_data)
                            self.Wind_tables[WnSector][WvSector] = self.calc_table(self.Wind_data)
                        else:
                            df_temp =  self.Swell_data[self.Swell_data.WvD_sectors == WvSector]
                            self.Swell_tables[WnSector][WvSector] = self.calc_table(df_temp)
                            df_temp =  self.Wind_data[self.Wind_data.WvD_sectors == WvSector]
                            self.Wind_tables[WnSector][WvSector] = self.calc_table(df_temp)
                    else: # SWELL COMPONENT NOT AFFECTED BY WIND
                        if WvSector == 0: # Wind tables contain values for filtered wind but omnidirectional waves
                            df_temp = self.Wind_data[
                                self.Wind_data.WnD_sectors == WnSector]
                            self.Wind_tables[WnSector][WvSector] = self.calc_table(df_temp)
                        else:
                            df_temp = self.Wind_data[
                                (self.Wind_data.WvD_sectors == WvSector) & (self.Wind_data.WnD_sectors == WnSector)]
                            self.Wind_tables[WnSector][WvSector] = self.calc_table(df_temp)

        print("NSS Tables calculated!")

    def calc_table(self, NSS_data):
        """ calc_table: [creates a single NSS table for a specific combination of wind and wave direction sector.
                    Works the same for Total, Wind or Swell waves.]

            Args: 
                NSS_data ([pandas Dataframe]): a dataframe containing wind and wave data for this particular table,
                    already filtered by wind and wave sector

            Returns:
                tab ([numpy array]): numpy array containing the NSS table
        """
        # Create an empty table of the right size
        tab = np.zeros((self.WS_bins_list.size,4))

        # Iterate and populate with calculated values or NaNs if there are no registries for a particular wind speed bin
        for i in range(self.WS_bins_list.size):
            bin_centre = self.WS_bins_list[i]
            df_temp = NSS_data[NSS_data.WS_bins == bin_centre]
            if df_temp.shape[0] == 0:
                tab[i] = [np.NAN, np.NAN, np.NAN, np.NAN]
            else:
                # probability of ocurrence of this particular wind speed bin in this particular wind and wave direction sector combination
                # over the total number of events in the timeseries
                tab[i][3] = df_temp["Hs"].size/self.Total_Count

                # Mean or median depending on user selection
                if self.method == "mean":
                    tab[i][0] = df_temp["Hs"].mean()
                    tab[i][1] = df_temp["Tp"].mean()
                    tab[i][2] = df_temp["G"].mean()
                        
                else:
                    tab[i][0] = df_temp["Hs"].median()
                    tab[i][1] = df_temp["Tp"].median()
                    tab[i][2] = df_temp["G"].median()

        return tab

    def print_NSS_tables(self):
        wb = Workbook()
        create_styles(wb)

        ws_Total = wb.active
        ws_Total.title = "NSS Total sea"
        ws_Total.sheet_properties.tabColor = "072B31"
        ws_Total.sheet_view.showGridLines = False
        
        startCol = 3
        startRow = 18

        table_WSbins = np.stack((
            self.WS_bins_list - self.WS_bin_size/2,
            self.WS_bins_list,
            self.WS_bins_list + self.WS_bin_size), axis=1)
        WS_bin_title = "Hourly Mean Wind Speed at {}mMSL [m/s]".format(self.WS_HH)
        WS_bin_headers = ["Lower","Mean","Upper"]
        NSS_Table_headers = ["Hs [m]","Tp [s]","γ [-]","Prob [%]"]

        for WnSector in range(1, self.NSectors_wind + 1):
            
            print_table(table_WSbins, ws_Total, startRow, startCol, WS_bin_headers,WS_bin_title)
            startCol += table_WSbins.shape[1]
            for WvSector in range(1, self.NSectors_wave + 1):
                print_table(self.Total_tables[WnSector][WvSector], ws_Total, startRow, startCol, NSS_Table_headers,cond_form=True)
                sum_prob = np.nansum(self.Total_tables[WnSector][WvSector][:,3])
                temp_data = np.array([["","","",sum_prob]])
                print_table(temp_data, ws_Total, startRow+self.Total_tables[WnSector][WvSector].shape[0]+1, startCol)
                startCol += temp_data.shape[1]
            
            startCol = 3
            startRow += table_WSbins.shape[0] + 7

        if self.wave_spectral:
            #COPY SHEET FROM TOTAL SEA ?
            ws_Wind = wb.create_sheet("NSS Wind sea", 1)
            ws_Wind.sheet_properties.tabColor = "D9D9D6"
            ws_Swell = wb.create_sheet("NSS Swell sea", 2)
            ws_Swell.sheet_properties.tabColor = "D9D9D6"
        
        wb.save("NSS_test_output.xlsx")

def create_styles(wb):
    NSS_header = NamedStyle(name="NSS_header")
    NSS_header.font = Font(bold=True, color="00FFFFFF")
    NSS_header.fill = PatternFill(fill_type="solid", fgColor="072B31")
    NSS_header.alignment = Alignment(horizontal="center")
    wb.add_named_style(NSS_header)

def print_table(data, ws, startRow, startCol, headers=None, title=None,footer=None,cond_form=False):
    
    if title != None:
        ws.cell(row = startRow-1, column= startCol).value = title
        c = startCol + data.shape[1] - 1
        ws.merge_cells(start_row=startRow-1, start_column=startCol, end_row=startRow-1, end_column=c)

    if headers != None:
        for c in range(len(headers)):
            col = startCol + c
            ws.cell(row = startRow, column = col).value = headers[c]
            ws.cell(row = startRow, column = col).style = "NSS_header"
        startRow += 1
    
    for r in range(data.shape[0]):
        row = startRow + r
        for c in range(data.shape[1]):
            col = startCol + c
            if data.dtype == np.float and np.isnan(data[r][c]):
                ws.cell(row = row, column = col).value = "NaN"
                ws.cell(row = row, column = col).alignment = Alignment(horizontal="center")
            else:
                if data[r][c] != "":
                    ws.cell(row = row, column = col).value = np.float(data[r][c])
                    ws.cell(row = row, column = col).alignment = Alignment(horizontal="center")
                    if c == 3:
                        ws.cell(row = row, column = col).number_format = "0.00%"
                    else:
                        ws.cell(row = row, column = col).number_format = "0.00"
            if c == 0:
                if r == 0:
                    ws.cell(row = row, column = col).border = Border(top=Side(style="thin"), left=Side(style="thin"))
                elif r == data.shape[0] - 1:
                    ws.cell(row = row, column = col).border = Border(bottom=Side(style="thin"), left=Side(style="thin"))
                else:
                    ws.cell(row = row, column = col).border = Border(left=Side(style="thin"))
            elif c == data.shape[1] - 1:
                if r == 0:
                    ws.cell(row = row, column = col).border = Border(top=Side(style="thin"), right=Side(style="thin"))
                elif r == data.shape[0] - 1:
                    ws.cell(row = row, column = col).border = Border(bottom=Side(style="thin"), right=Side(style="thin"))
                else:
                    ws.cell(row = row, column = col).border = Border(right=Side(style="thin"))
            else:
                if r == 0:
                    ws.cell(row = row, column = col).border = Border(top=Side(style="thin"))
                elif r == data.shape[0] - 1:
                    ws.cell(row = row, column = col).border = Border(bottom=Side(style="thin"))
    
    if cond_form == True: 
        for c in range(data.shape[1]): 
            col = startCol + c
            col_letter = utils.cell.get_column_letter(col)
            col_range = "{}{}:{}{}".format(col_letter,startRow,col_letter,startRow+data.shape[0]-1)
            ws.conditional_formatting.add(col_range,
                        ColorScaleRule(start_type="min", start_color="63BE7B",
                                        mid_type="percentile", mid_value=50, mid_color="FFEB84",
                                        end_type="max", end_color="F8696B")
                        )
