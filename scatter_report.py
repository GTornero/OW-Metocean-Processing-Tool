import time
import xlsxwriter

from scatter import Scatter


def print_scatter_report(metocean_data):
    """print_scatter_report Function that takes the metocean_data object and creates all the necessary
    scatter tables and prints them into an excel .xlsx scatter table report.

    Args:
        metocean_data (MetoceanData): A MetoceanData object from the metocean_data module.
    """

    start_time = time.perf_counter()

    with xlsxwriter.Workbook(
        f"{metocean_data.config['project']}_Metocean_Scatter_Tables.xlsx"
    ) as wb:
        tables = []
        # -----------------------------------------------------------------------------------------
        # ------------------------Wind Speed Vs Wind Direction Tables (Omni)-----------------------
        # -----------------------------------------------------------------------------------------
        # If wind data has been input
        if metocean_data.config["wind_status"]:
            # Make Wind Speed Vs Wind Direction Scatter Tables
            table = Scatter(metocean_data, ["WnD_sectors", "WS_bins"])
            ws = wb.add_worksheet("WndSpd-WndDir (@HH)")
            ws.hide_gridlines(2)
            table.print_table(wb, ws, row=1, col=1)
            # If wind @ 10m MSL has been input
            if metocean_data.config["10m"]:
                table = Scatter(metocean_data, ["WnD_10_sectors", "WS_10_bins"])
                ws = wb.add_worksheet("WndSpd-WndDir (@10m)")
                ws.hide_gridlines(2)
                table.print_table(wb, ws, row=1, col=1)
        # -----------------------------------------------------------------------------------------
        # ---------------------------Hs Vs Wave Direction Tables (Omni)----------------------------
        # -----------------------------------------------------------------------------------------
        # If wave data has been input
        if metocean_data.config["wave_status"]:
            table = Scatter(metocean_data, ["WvD_sectors", "Hs_bins"])
            tables.append(table)
            # If spectral wave components have been input
            if metocean_data.config["wave_spectral"]:
                table = Scatter(metocean_data, ["WvD_S_sectors", "Hs_S_bins"])
                tables.append(table)
                table = Scatter(metocean_data, ["WvD_W_sectors", "Hs_W_bins"])
                tables.append(table)
            ws = wb.add_worksheet("Hs-WaveDir")
            ws.hide_gridlines(2)
            for i, table in enumerate(tables):
                table.print_table(
                    wb,
                    ws,
                    row=1,
                    col=(1 + i * (5 + table.table.shape[1])),
                )
            tables.clear()

        # If both wind and wave data have been input
        if metocean_data.config["wind_status"] and metocean_data.config["wave_status"]:
            # -----------------------------------------------------------------------------------------
            # -------------------------Hs vs Wind Direction (@HH) Tables (Omni)------------------------
            # -----------------------------------------------------------------------------------------
            table = Scatter(metocean_data, ["WnD_sectors", "Hs_bins"])
            tables.append(table)
            # If spectral wave components have been input
            if metocean_data.config["wave_spectral"]:
                table = Scatter(metocean_data, ["WnD_sectors", "Hs_S_bins"])
                tables.append(table)
                table = Scatter(metocean_data, ["WnD_sectors", "Hs_W_bins"])
                tables.append(table)
            ws = wb.add_worksheet("Hs-WindDir (@HH)")
            ws.hide_gridlines(2)
            for i, table in enumerate(tables):
                table.print_table(
                    wb,
                    ws,
                    row=1,
                    col=(1 + i * (5 + table.table.shape[1])),
                )
            tables.clear()
            if metocean_data.config["10m"]:
                # -----------------------------------------------------------------------------------------
                # -------------------------Hs vs Wind Direction (@10m) Tables (Omni)-----------------------
                # -----------------------------------------------------------------------------------------
                table = Scatter(metocean_data, ["WnD_10_sectors", "Hs_bins"])
                tables.append(table)
                # If spectral wave components have been input
                if metocean_data.config["wave_spectral"]:
                    table = Scatter(metocean_data, ["WnD_10_sectors", "Hs_S_bins"])
                    tables.append(table)
                    table = Scatter(metocean_data, ["WnD_10_sectors", "Hs_W_bins"])
                    tables.append(table)
                ws = wb.add_worksheet("Hs-WindDir (@10m)")
                ws.hide_gridlines(2)
                for i, table in enumerate(tables):
                    table.print_table(
                        wb,
                        ws,
                        row=1,
                        col=(1 + i * (5 + table.table.shape[1])),
                    )
                tables.clear()
            # -----------------------------------------------------------------------------------------
            # ----------------Wind Speed (@HH) vs Hs (Totalsea) Tables (misalignments)-----------------
            # -----------------------------------------------------------------------------------------
            # Omnidirectional table first
            tables.append([])
            table = Scatter(metocean_data, ["Hs_bins", "WS_bins"])
            tables[0].append(table)
            tables.append([])
            # Omnidirectional wind, directional wave tables
            for wave_sect in range(metocean_data.config["wave_sectors"]):
                table = Scatter(
                    metocean_data,
                    ["Hs_bins", "WS_bins"],
                    keys=["WvD_sectors", False],
                    x_filt=wave_sect + 1,
                )
                tables[1].append(table)
            tables.append([])
            # Omnidirecitonal wave, directional wind tables
            for wind_sect in range(metocean_data.config["wind_sectors"]):
                table = Scatter(
                    metocean_data,
                    ["Hs_bins", "WS_bins"],
                    keys=[False, "WnD_sectors"],
                    y_filt=wind_sect + 1,
                )
                tables[2].append(table)
            # 144 wind-wave misalignment tables
            for wind_sect in range(metocean_data.config["wind_sectors"]):
                tables.append([])
                for wave_sect in range(metocean_data.config["wave_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Hs_bins", "WS_bins"],
                        ["WvD_sectors", "WnD_sectors"],
                        x_filt=wind_sect + 1,
                        y_filt=wave_sect + 1,
                    )
                    tables[3 + wind_sect].append(table)
            ws = wb.add_worksheet("WndSpd (@HH)-Hs (Totalsea)")
            ws.hide_gridlines(2)
            for i, row in enumerate(tables):
                for j, table in enumerate(row):
                    table.print_table(
                        wb,
                        ws,
                        row=(1 + i * (6 + table.table.shape[0])),
                        col=(1 + j * (5 + table.table.shape[1])),
                    )
            tables.clear()
            if metocean_data.config["wave_spectral"]:
                # -----------------------------------------------------------------------------------------
                # -----------------Wind Speed (@HH) vs Hs (Swell) Tables (misalignments)-------------------
                # -----------------------------------------------------------------------------------------
                # Omnidirectional table first
                tables.append([])
                table = Scatter(metocean_data, ["Hs_S_bins", "WS_bins"])
                tables[0].append(table)
                tables.append([])
                # Omnidirectional wind, directional wave tables
                for wave_sect in range(metocean_data.config["wave_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Hs_S_bins", "WS_bins"],
                        keys=["WvD_S_sectors", False],
                        x_filt=wave_sect + 1,
                    )
                    tables[1].append(table)
                tables.append([])
                # Omnidirecitonal wave, directional wind tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Hs_S_bins", "WS_bins"],
                        keys=[False, "WnD_sectors"],
                        y_filt=wind_sect + 1,
                    )
                    tables[2].append(table)
                # 144 wind-wave misalignment tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    tables.append([])
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Hs_S_bins", "WS_bins"],
                            ["WvD_S_sectors", "WnD_sectors"],
                            x_filt=wind_sect + 1,
                            y_filt=wave_sect + 1,
                        )
                        tables[3 + wind_sect].append(table)
                ws = wb.add_worksheet("WndSpd (@HH)-Hs (Swell)")
                ws.hide_gridlines(2)
                for i, row in enumerate(tables):
                    for j, table in enumerate(row):
                        table.print_table(
                            wb,
                            ws,
                            row=(1 + i * (6 + table.table.shape[0])),
                            col=(1 + j * (5 + table.table.shape[1])),
                        )
                tables.clear()
                # -----------------------------------------------------------------------------------------
                # -----------------Wind Speed (@HH) vs Hs (Windsea) Tables (misalignments)-----------------
                # -----------------------------------------------------------------------------------------
                # Omnidirectional table first
                tables.append([])
                table = Scatter(metocean_data, ["Hs_W_bins", "WS_bins"])
                tables[0].append(table)
                tables.append([])
                # Omnidirectional wind, directional wave tables
                for wave_sect in range(metocean_data.config["wave_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Hs_W_bins", "WS_bins"],
                        keys=["WvD_W_sectors", False],
                        x_filt=wave_sect + 1,
                    )
                    tables[1].append(table)
                tables.append([])
                # Omnidirecitonal wave, directional wind tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Hs_W_bins", "WS_bins"],
                        keys=[False, "WnD_sectors"],
                        y_filt=wind_sect + 1,
                    )
                    tables[2].append(table)
                # 144 wind-wave misalignment tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    tables.append([])
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Hs_W_bins", "WS_bins"],
                            ["WvD_W_sectors", "WnD_sectors"],
                            x_filt=wind_sect + 1,
                            y_filt=wave_sect + 1,
                        )
                        tables[3 + wind_sect].append(table)
                ws = wb.add_worksheet("WndSpd (@HH)-Hs (Windsea)")
                ws.hide_gridlines(2)
                for i, row in enumerate(tables):
                    for j, table in enumerate(row):
                        table.print_table(
                            wb,
                            ws,
                            row=(1 + i * (6 + table.table.shape[0])),
                            col=(1 + j * (5 + table.table.shape[1])),
                        )
                tables.clear()
            if metocean_data.config["10m"]:
                # -----------------------------------------------------------------------------------------
                # ----------------Wind Speed (@10m) vs Hs (Totalsea) Tables (misalignments)----------------
                # -----------------------------------------------------------------------------------------
                # Omnidirectional table first
                tables.append([])
                table = Scatter(metocean_data, ["Hs_bins", "WS_10_bins"])
                tables[0].append(table)
                tables.append([])
                # Omnidirectional wind, directional wave tables
                for wave_sect in range(metocean_data.config["wave_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Hs_bins", "WS_10_bins"],
                        keys=["WvD_sectors", False],
                        x_filt=wave_sect + 1,
                    )
                    tables[1].append(table)
                tables.append([])
                # Omnidirecitonal wave, directional wind tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Hs_bins", "WS_10_bins"],
                        keys=[False, "WnD_10_sectors"],
                        y_filt=wind_sect + 1,
                    )
                    tables[2].append(table)
                # 144 wind-wave misalignment tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    tables.append([])
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Hs_bins", "WS_10_bins"],
                            ["WvD_sectors", "WnD_10_sectors"],
                            x_filt=wind_sect + 1,
                            y_filt=wave_sect + 1,
                        )
                        tables[3 + wind_sect].append(table)
                ws = wb.add_worksheet("WndSpd (@10m)-Hs (Totalsea)")
                ws.hide_gridlines(2)
                for i, row in enumerate(tables):
                    for j, table in enumerate(row):
                        table.print_table(
                            wb,
                            ws,
                            row=(1 + i * (6 + table.table.shape[0])),
                            col=(1 + j * (5 + table.table.shape[1])),
                        )
                tables.clear()
                if metocean_data.config["wave_spectral"]:
                    # -----------------------------------------------------------------------------------------
                    # -----------------Wind Speed (@10m) vs Hs (Swell) Tables (misalignments)------------------
                    # -----------------------------------------------------------------------------------------
                    # Omnidirectional table first
                    tables.append([])
                    table = Scatter(metocean_data, ["Hs_S_bins", "WS_10_bins"])
                    tables[0].append(table)
                    tables.append([])
                    # Omnidirectional wind, directional wave tables
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Hs_S_bins", "WS_10_bins"],
                            keys=["WvD_S_sectors", False],
                            x_filt=wave_sect + 1,
                        )
                        tables[1].append(table)
                    tables.append([])
                    # Omnidirecitonal wave, directional wind tables
                    for wind_sect in range(metocean_data.config["wind_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Hs_S_bins", "WS_10_bins"],
                            keys=[False, "WnD_10_sectors"],
                            y_filt=wind_sect + 1,
                        )
                        tables[2].append(table)
                    # 144 wind-wave misalignment tables
                    for wind_sect in range(metocean_data.config["wind_sectors"]):
                        tables.append([])
                        for wave_sect in range(metocean_data.config["wave_sectors"]):
                            table = Scatter(
                                metocean_data,
                                ["Hs_S_bins", "WS_10_bins"],
                                ["WvD_S_sectors", "WnD_10_sectors"],
                                x_filt=wind_sect + 1,
                                y_filt=wave_sect + 1,
                            )
                            tables[3 + wind_sect].append(table)
                    ws = wb.add_worksheet("WndSpd (@10m)-Hs (Swell)")
                    ws.hide_gridlines(2)
                    for i, row in enumerate(tables):
                        for j, table in enumerate(row):
                            table.print_table(
                                wb,
                                ws,
                                row=(1 + i * (6 + table.table.shape[0])),
                                col=(1 + j * (5 + table.table.shape[1])),
                            )
                    tables.clear()
                    # -----------------------------------------------------------------------------------------
                    # -----------------Wind Speed (@10m) vs Hs (Windsea) Tables (misalignments)----------------
                    # -----------------------------------------------------------------------------------------
                    # Omnidirectional table first
                    tables.append([])
                    table = Scatter(metocean_data, ["Hs_W_bins", "WS_10_bins"])
                    tables[0].append(table)
                    tables.append([])
                    # Omnidirectional wind, directional wave tables
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Hs_W_bins", "WS_10_bins"],
                            keys=["WvD_W_sectors", False],
                            x_filt=wave_sect + 1,
                        )
                        tables[1].append(table)
                    tables.append([])
                    # Omnidirecitonal wave, directional wind tables
                    for wind_sect in range(metocean_data.config["wind_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Hs_W_bins", "WS_10_bins"],
                            keys=[False, "WnD_10_sectors"],
                            y_filt=wind_sect + 1,
                        )
                        tables[2].append(table)
                    # 144 wind-wave misalignment tables
                    for wind_sect in range(metocean_data.config["wind_sectors"]):
                        tables.append([])
                        for wave_sect in range(metocean_data.config["wave_sectors"]):
                            table = Scatter(
                                metocean_data,
                                ["Hs_W_bins", "WS_10_bins"],
                                ["WvD_W_sectors", "WnD_10_sectors"],
                                x_filt=wind_sect + 1,
                                y_filt=wave_sect + 1,
                            )
                            tables[3 + wind_sect].append(table)
                    ws = wb.add_worksheet("WndSpd (@10m)-Hs (Windsea)")
                    ws.hide_gridlines(2)
                    for i, row in enumerate(tables):
                        for j, table in enumerate(row):
                            table.print_table(
                                wb,
                                ws,
                                row=(1 + i * (6 + table.table.shape[0])),
                                col=(1 + j * (5 + table.table.shape[1])),
                            )
                    tables.clear()
            # -----------------------------------------------------------------------------------------
            # ---------------------Hs Vs Tp (Totalsea) Tables (@ HH misalignments)---------------------
            # -----------------------------------------------------------------------------------------
            # Omnidirectional table first
            tables.append([])
            table = Scatter(metocean_data, ["Tp_bins", "Hs_bins"])
            tables[0].append(table)
            tables.append([])
            # Omnidirectional wind, directional wave tables
            for wave_sect in range(metocean_data.config["wave_sectors"]):
                table = Scatter(
                    metocean_data,
                    ["Tp_bins", "Hs_bins"],
                    keys=["WvD_sectors", False],
                    x_filt=wave_sect + 1,
                )
                tables[1].append(table)
            tables.append([])
            # Omnidirecitonal wave, directional wind tables
            for wind_sect in range(metocean_data.config["wind_sectors"]):
                table = Scatter(
                    metocean_data,
                    ["Tp_bins", "Hs_bins"],
                    keys=[False, "WnD_sectors"],
                    y_filt=wind_sect + 1,
                )
                tables[2].append(table)
            # 144 wind-wave misalignment tables
            for wind_sect in range(metocean_data.config["wind_sectors"]):
                tables.append([])
                for wave_sect in range(metocean_data.config["wave_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Tp_bins", "Hs_bins"],
                        ["WvD_sectors", "WnD_sectors"],
                        x_filt=wind_sect + 1,
                        y_filt=wave_sect + 1,
                    )
                    tables[3 + wind_sect].append(table)
            ws = wb.add_worksheet("Hs-Tp (Totalsea) (Wind @HH)")
            ws.hide_gridlines(2)
            for i, row in enumerate(tables):
                for j, table in enumerate(row):
                    table.print_table(
                        wb,
                        ws,
                        row=(1 + i * (6 + table.table.shape[0])),
                        col=(1 + j * (5 + table.table.shape[1])),
                    )
            tables.clear()
            if metocean_data.config["wave_spectral"]:
                # -----------------------------------------------------------------------------------------
                # -----------------------Hs Vs Tp (Swell) Tables (@ HH misalignments)----------------------
                # -----------------------------------------------------------------------------------------
                # Omnidirectional table first
                tables.append([])
                table = Scatter(metocean_data, ["Tp_S_bins", "Hs_S_bins"])
                tables[0].append(table)
                tables.append([])
                # Omnidirectional wind, directional wave tables
                for wave_sect in range(metocean_data.config["wave_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Tp_S_bins", "Hs_S_bins"],
                        keys=["WvD_S_sectors", False],
                        x_filt=wave_sect + 1,
                    )
                    tables[1].append(table)
                tables.append([])
                # Omnidirecitonal wave, directional wind tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Tp_S_bins", "Hs_S_bins"],
                        keys=[False, "WnD_sectors"],
                        y_filt=wind_sect + 1,
                    )
                    tables[2].append(table)
                # 144 wind-wave misalignment tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    tables.append([])
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Tp_S_bins", "Hs_S_bins"],
                            ["WvD_S_sectors", "WnD_sectors"],
                            x_filt=wind_sect + 1,
                            y_filt=wave_sect + 1,
                        )
                        tables[3 + wind_sect].append(table)
                ws = wb.add_worksheet("Hs-Tp (Swell) (Wind @HH)")
                ws.hide_gridlines(2)
                for i, row in enumerate(tables):
                    for j, table in enumerate(row):
                        table.print_table(
                            wb,
                            ws,
                            row=(1 + i * (6 + table.table.shape[0])),
                            col=(1 + j * (5 + table.table.shape[1])),
                        )
                tables.clear()
                # -----------------------------------------------------------------------------------------
                # ----------------------Hs Vs Tp (Windsea) Tables (@ HH misalignments)---------------------
                # -----------------------------------------------------------------------------------------
                # Omnidirectional table first
                tables.append([])
                table = Scatter(metocean_data, ["Tp_W_bins", "Hs_W_bins"])
                tables[0].append(table)
                tables.append([])
                # Omnidirectional wind, directional wave tables
                for wave_sect in range(metocean_data.config["wave_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Tp_W_bins", "Hs_W_bins"],
                        keys=["WvD_W_sectors", False],
                        x_filt=wave_sect + 1,
                    )
                    tables[1].append(table)
                tables.append([])
                # Omnidirecitonal wave, directional wind tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Tp_W_bins", "Hs_W_bins"],
                        keys=[False, "WnD_sectors"],
                        y_filt=wind_sect + 1,
                    )
                    tables[2].append(table)
                # 144 wind-wave misalignment tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    tables.append([])
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Tp_W_bins", "Hs_W_bins"],
                            ["WvD_W_sectors", "WnD_sectors"],
                            x_filt=wind_sect + 1,
                            y_filt=wave_sect + 1,
                        )
                        tables[3 + wind_sect].append(table)
                ws = wb.add_worksheet("Hs-Tp (Windsea) (Wind @HH)")
                ws.hide_gridlines(2)
                for i, row in enumerate(tables):
                    for j, table in enumerate(row):
                        table.print_table(
                            wb,
                            ws,
                            row=(1 + i * (6 + table.table.shape[0])),
                            col=(1 + j * (5 + table.table.shape[1])),
                        )
                tables.clear()
            if metocean_data.config["10m"]:
                # -----------------------------------------------------------------------------------------
                # --------------------Hs Vs Tp (Totalsea) Tables (@ 10m misalignments)---------------------
                # -----------------------------------------------------------------------------------------
                # Omnidirectional table first
                tables.append([])
                table = Scatter(metocean_data, ["Tp_bins", "Hs_bins"])
                tables[0].append(table)
                tables.append([])
                # Omnidirectional wind, directional wave tables
                for wave_sect in range(metocean_data.config["wave_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Tp_bins", "Hs_bins"],
                        keys=["WvD_sectors", False],
                        x_filt=wave_sect + 1,
                    )
                    tables[1].append(table)
                tables.append([])
                # Omnidirecitonal wave, directional wind tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    table = Scatter(
                        metocean_data,
                        ["Tp_bins", "Hs_bins"],
                        keys=[False, "WnD_10_sectors"],
                        y_filt=wind_sect + 1,
                    )
                    tables[2].append(table)
                # 144 wind-wave misalignment tables
                for wind_sect in range(metocean_data.config["wind_sectors"]):
                    tables.append([])
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Tp_bins", "Hs_bins"],
                            ["WvD_sectors", "WnD_10_sectors"],
                            x_filt=wind_sect + 1,
                            y_filt=wave_sect + 1,
                        )
                        tables[3 + wind_sect].append(table)
                ws = wb.add_worksheet("Hs-Tp (Totalsea) (Wind @10m)")
                ws.hide_gridlines(2)
                for i, row in enumerate(tables):
                    for j, table in enumerate(row):
                        table.print_table(
                            wb,
                            ws,
                            row=(1 + i * (6 + table.table.shape[0])),
                            col=(1 + j * (5 + table.table.shape[1])),
                        )
                tables.clear()
                if metocean_data.config["wave_spectral"]:
                    # -----------------------------------------------------------------------------------------
                    # ----------------------Hs Vs Tp (Swell) Tables (@ 10m misalignments)----------------------
                    # -----------------------------------------------------------------------------------------
                    # Omnidirectional table first
                    tables.append([])
                    table = Scatter(metocean_data, ["Tp_S_bins", "Hs_S_bins"])
                    tables[0].append(table)
                    tables.append([])
                    # Omnidirectional wind, directional wave tables
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Tp_S_bins", "Hs_S_bins"],
                            keys=["WvD_S_sectors", False],
                            x_filt=wave_sect + 1,
                        )
                        tables[1].append(table)
                    tables.append([])
                    # Omnidirecitonal wave, directional wind tables
                    for wind_sect in range(metocean_data.config["wind_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Tp_S_bins", "Hs_S_bins"],
                            keys=[False, "WnD_10_sectors"],
                            y_filt=wind_sect + 1,
                        )
                        tables[2].append(table)
                    # 144 wind-wave misalignment tables
                    for wind_sect in range(metocean_data.config["wind_sectors"]):
                        tables.append([])
                        for wave_sect in range(metocean_data.config["wave_sectors"]):
                            table = Scatter(
                                metocean_data,
                                ["Tp_S_bins", "Hs_S_bins"],
                                ["WvD_S_sectors", "WnD_10_sectors"],
                                x_filt=wind_sect + 1,
                                y_filt=wave_sect + 1,
                            )
                            tables[3 + wind_sect].append(table)
                    ws = wb.add_worksheet("Hs-Tp (Swell) (Wind @10m)")
                    ws.hide_gridlines(2)
                    for i, row in enumerate(tables):
                        for j, table in enumerate(row):
                            table.print_table(
                                wb,
                                ws,
                                row=(1 + i * (6 + table.table.shape[0])),
                                col=(1 + j * (5 + table.table.shape[1])),
                            )
                    tables.clear()
                    # -----------------------------------------------------------------------------------------
                    # ----------------------Hs Vs Tp (Windsea) Tables (@ 10m misalignments)--------------------
                    # -----------------------------------------------------------------------------------------
                    # Omnidirectional table first
                    tables.append([])
                    table = Scatter(metocean_data, ["Tp_W_bins", "Hs_W_bins"])
                    tables[0].append(table)
                    tables.append([])
                    # Omnidirectional wind, directional wave tables
                    for wave_sect in range(metocean_data.config["wave_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Tp_W_bins", "Hs_W_bins"],
                            keys=["WvD_W_sectors", False],
                            x_filt=wave_sect + 1,
                        )
                        tables[1].append(table)
                    tables.append([])
                    # Omnidirecitonal wave, directional wind tables
                    for wind_sect in range(metocean_data.config["wind_sectors"]):
                        table = Scatter(
                            metocean_data,
                            ["Tp_W_bins", "Hs_W_bins"],
                            keys=[False, "WnD_10_sectors"],
                            y_filt=wind_sect + 1,
                        )
                        tables[2].append(table)
                    # 144 wind-wave misalignment tables
                    for wind_sect in range(metocean_data.config["wind_sectors"]):
                        tables.append([])
                        for wave_sect in range(metocean_data.config["wave_sectors"]):
                            table = Scatter(
                                metocean_data,
                                ["Tp_W_bins", "Hs_W_bins"],
                                ["WvD_W_sectors", "WnD_10_sectors"],
                                x_filt=wind_sect + 1,
                                y_filt=wave_sect + 1,
                            )
                            tables[3 + wind_sect].append(table)
                    ws = wb.add_worksheet("Hs-Tp (Windsea) (Wind @10m)")
                    ws.hide_gridlines(2)
                    for i, row in enumerate(tables):
                        for j, table in enumerate(row):
                            table.print_table(
                                wb,
                                ws,
                                row=(1 + i * (6 + table.table.shape[0])),
                                col=(1 + j * (5 + table.table.shape[1])),
                            )
                    tables.clear()
            # -----------------------------------------------------------------------------------------
            # -------------------Wind Direction (@HH) vs Wave Direction Tables (Omni)------------------
            # -----------------------------------------------------------------------------------------
            table = Scatter(metocean_data, ["WvD_sectors", "WnD_sectors"])
            tables.append(table)
            # If spectral wave components have been input
            if metocean_data.config["wave_spectral"]:
                table = Scatter(metocean_data, ["WvD_S_sectors", "WnD_sectors"])
                tables.append(table)
                table = Scatter(metocean_data, ["WvD_W_sectors", "WnD_sectors"])
                tables.append(table)
            ws = wb.add_worksheet("WindDir-WaveDir (@HH)")
            ws.hide_gridlines(2)
            for i, table in enumerate(tables):
                table.print_table(
                    wb,
                    ws,
                    row=1,
                    col=(1 + i * (5 + table.table.shape[1])),
                )
            tables.clear()
            # -----------------------------------------------------------------------------------------
            # ----------------Wind Direction (@HH) vs Wave Direction Tables (by WndSpd)----------------
            # -----------------------------------------------------------------------------------------
            for i, wind_bin in enumerate(metocean_data.bins["WS"]):
                tables.append([])
                table = Scatter(
                    metocean_data,
                    ["WvD_sectors", "WnD_sectors"],
                    keys=["WS_bins", False],
                    x_filt=wind_bin,
                )
                tables[i].append(table)
                # If Spectral wave components have been input
                if metocean_data.config["wave_spectral"]:
                    table = Scatter(
                        metocean_data,
                        ["WvD_S_sectors", "WnD_sectors"],
                        keys=["WS_bins", False],
                        x_filt=wind_bin,
                    )
                    tables[i].append(table)
                    table = Scatter(
                        metocean_data,
                        ["WvD_W_sectors", "WnD_sectors"],
                        keys=["WS_bins", False],
                        x_filt=wind_bin,
                    )
                    tables[i].append(table)
            ws = wb.add_worksheet("WindDir-WaveDir by WndSpd (@HH)")
            ws.hide_gridlines(2)
            for i, row in enumerate(tables):
                for j, table in enumerate(row):
                    table.print_table(
                        wb,
                        ws,
                        row=(1 + i * (6 + table.table.shape[0])),
                        col=(1 + j * (5 + table.table.shape[1])),
                    )
            tables.clear()
            # -----------------------------------------------------------------------------------------
            # -------------------Wind Direction (@10m) vs Wave Direction Tables (Omni)------------------
            # -----------------------------------------------------------------------------------------
            # If 10m MSL wind data has been input
            if metocean_data.config["10m"]:
                table = Scatter(metocean_data, ["WvD_sectors", "WnD_10_sectors"])
                tables.append(table)
                # If spectral wave components have been input
                if metocean_data.config["wave_spectral"]:
                    table = Scatter(metocean_data, ["WvD_S_sectors", "WnD_10_sectors"])
                    tables.append(table)
                    table = Scatter(metocean_data, ["WvD_W_sectors", "WnD_10_sectors"])
                    tables.append(table)
                ws = wb.add_worksheet("WindDir-WaveDir (@10m)")
                ws.hide_gridlines(2)
                for i, table in enumerate(tables):
                    table.print_table(
                        wb,
                        ws,
                        row=1,
                        col=(1 + i * (5 + table.table.shape[1])),
                    )
                tables.clear()
                # -----------------------------------------------------------------------------------------
                # ----------------Wind Direction (@10m) vs Wave Direction Tables (by WndSpd)---------------
                # -----------------------------------------------------------------------------------------
                for i, wind_bin in enumerate(metocean_data.bins["WS"]):
                    tables.append([])
                    table = Scatter(
                        metocean_data,
                        ["WvD_sectors", "WnD_10_sectors"],
                        keys=["WS_10_bins", False],
                        x_filt=wind_bin,
                    )
                    tables[i].append(table)
                    # If Spectral wave components have been input
                    if metocean_data.config["wave_spectral"]:
                        table = Scatter(
                            metocean_data,
                            ["WvD_S_sectors", "WnD_10_sectors"],
                            keys=["WS_10_bins", False],
                            x_filt=wind_bin,
                        )
                        tables[i].append(table)
                        table = Scatter(
                            metocean_data,
                            ["WvD_W_sectors", "WnD_10_sectors"],
                            keys=["WS_10_bins", False],
                            x_filt=wind_bin,
                        )
                        tables[i].append(table)
                ws = wb.add_worksheet("WindDir-WaveDir by WndSpd(@10m)")
                ws.hide_gridlines(2)
                for i, row in enumerate(tables):
                    for j, table in enumerate(row):
                        table.print_table(
                            wb,
                            ws,
                            row=(1 + i * (6 + table.table.shape[0])),
                            col=(1 + j * (5 + table.table.shape[1])),
                        )
                tables.clear()
        # -----------------------------------------------------------------------------------------
        # ----------------Surface Current Speed Vs Current Direction Tables (Omni)-----------------
        # -----------------------------------------------------------------------------------------
        # If current data has been input
        if metocean_data.config["current_status"]:
            # Omnidirectional surface current speed  table first
            table = Scatter(metocean_data, ["CD_sectors", "SV_bins"])
            tables.append(table)
            # If current data by components has been input
            if metocean_data.config["current_components"]:
                # Create tidal surface current table
                table = Scatter(metocean_data, ["CD_Tid_sectors", "SV_Tid_bins"])
                tables.append(table)
                # Create residual surface current table
                table = Scatter(metocean_data, ["CD_Res_sectors", "SV_Res_bins"])
                tables.append(table)
            ws = wb.add_worksheet("Srfc CurrentSpd-CurrentDir")
            ws.hide_gridlines(2)
            for i, table in enumerate(tables):
                table.print_table(
                    wb,
                    ws,
                    row=1,
                    col=(1 + i * (5 + table.table.shape[1])),
                )
            tables.clear()
            # -----------------------------------------------------------------------------------------
            # ------------Depth Averaged Current Speed Vs Current Direction Tables (Omni)--------------
            # -----------------------------------------------------------------------------------------
            # Omnidirectional depth averaged current speed table first
            table = Scatter(metocean_data, ["CD_sectors", "DaV_bins"])
            tables.append(table)
            # If current data by components has been input
            if metocean_data.config["current_components"]:
                # Create tidal depth averaged current table
                table = Scatter(metocean_data, ["CD_Tid_sectors", "DaV_Tid_bins"])
                tables.append(table)
                # Create residual depth averaged current table
                table = Scatter(metocean_data, ["CD_Res_sectors", "DaV_Res_bins"])
                tables.append(table)
            ws = wb.add_worksheet("DpthAvg CurrentSpd-CurrentDir")
            ws.hide_gridlines(2)
            for i, table in enumerate(tables):
                table.print_table(
                    wb,
                    ws,
                    row=1,
                    col=(1 + i * (5 + table.table.shape[1])),
                )
            tables.clear()

    end_time = time.perf_counter()
    print(f"Report Finished in {round((end_time - start_time)/60, 2)} minutes.")