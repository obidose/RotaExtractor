import pandas as pd
from tkinter import filedialog as fd
from datetime import timedelta
import tkinter as tk
import tkinter.messagebox
# import numpy as np




def open_file():
    """Uses a GUI to select a file then returns the content of this file as a dataframe"""
    filename = fd.askopenfilename(title="Select a source file", filetypes=[("Excel Files", "*.xlsx")])
    dataframe1 = pd.read_excel(filename, skiprows=6, header=None, usecols="B:O")
    return dataframe1


def export_xls(df):
    """Opens a window to select a save directory. Takes a dataframe and saves it to the selected directory as
    output.xlsx """
    df.to_excel(fd.askdirectory(title="Select Save Directory") + "/" + "output.xlsx")
    tk.messagebox.showinfo('Save Complete', 'Output saved!')


def export_csv(df):
    """Opens a window to select a save directory. Takes a dataframe and saves it to the selected directory as
    output.csv """
    df.to_csv(fd.askdirectory(title="Select Save Directory") + "/" + "output.csv")
    tk.messagebox.showinfo('Save Complete', 'Output saved!')


def stack_days_vertically(df):
    """Splits the document on each day and stacks them so that there is a one-day wide version of the rota"""

    df_mon = pd.DataFrame(df.iloc[:, 0:2])
    df_mon.reset_index(drop=True, inplace=True)
    df_mon.columns = range(df_mon.shape[1])

    df_tue = pd.DataFrame(df.iloc[:, 2:4])
    df_tue.reset_index(drop=True, inplace=True)
    df_tue.columns = range(df_tue.shape[1])

    df_wed = pd.DataFrame(df.iloc[:, 4:6])
    df_wed.reset_index(drop=True, inplace=True)
    df_wed.columns = range(df_tue.shape[1])

    df_thu = pd.DataFrame(df.iloc[:, 6:8])
    df_thu.reset_index(drop=True, inplace=True)
    df_thu.columns = range(df_tue.shape[1])

    df_fri = pd.DataFrame(df.iloc[:, 8:10])
    df_fri.reset_index(drop=True, inplace=True)
    df_fri.columns = range(df_tue.shape[1])

    df_sat = pd.DataFrame(df.iloc[:, 10:12])
    df_sat.reset_index(drop=True, inplace=True)
    df_sat.columns = range(df_tue.shape[1])

    df_sun = pd.DataFrame(df.iloc[:, 12:14])
    df_sun.reset_index(drop=True, inplace=True)
    df_sun.columns = range(df_tue.shape[1])

    frames = [df_mon, df_tue, df_wed, df_thu, df_fri, df_sat, df_sun]
    dfo = pd.concat(frames)
    return dfo


def sort_to_row(df):
    """ Uses known template of relative cell positions to sort data into individual rows"""

    df_out = pd.DataFrame(columns=["Date", "Start", "End", "Comment"])

    weekdays = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")

    n = 0
    for i in range(len(df)):
        comment = ""
        date = ""
        start = 0
        end = 0

        if df.iloc[i, 0] in weekdays:
            try:
                float(df.iloc[i + 2, 0])
            except ValueError:
                comment = df.iloc[i + 2, 0]
            else:
                start = df.iloc[i + 2, 0]
                end = df.iloc[i + 2, 1]
            date = df.iloc[i + 1, 0]
            df_out.loc[n, "Date"] = date
            df_out.loc[n, "Start"] = start
            df_out.loc[n, "End"] = end
            df_out.loc[n, "Comment"] = comment
            n += 1

    df_out["Date"] = pd.to_datetime(df_out["Date"], format='%Y-%m-%d %H:%M:%S')
    df_out = df_out.sort_values(by="Date")
    return df_out


def name_shifts(df):
    """ Adds names to shifts following known patterns"""

    for i in range(len(df)):
        if df.loc[i, "Start"] == 17:
            df.loc[i, "Comment"] = "Late"
        if df.loc[i, "Start"] == 8:
            if df.loc[i, "End"] == 17.15:
                df.loc[i, "Comment"] = "Early"
            if df.loc[i, "End"] == 20:
                df.loc[i, "Comment"] = "Long Day"
        if df.loc[i, "Start"] == 22:
            df.loc[i, "Comment"] = "Night"
        if df.loc[i, "Start"] == 13:
            df.loc[i, "Comment"] = "Long Late"

    return df


def format_time(time):
    """Formats times which have been stored as inappropriate integers to a consistent string (e.g. "15.42")
    *Not a very clean method and can likely be improved"""

    n = 2
    time = str(time)
    new_time = "Error"
    for letter in time:
        if letter == ".":
            new_time = "0" * n + time
        n -= 1
    trails = 5 - len(new_time)
    new_time = new_time + "0" * trails
    return new_time


def calculate_shifts(df):
    df["Start Date"] = ""
    df["End Date"] = ""

    # Adjust end dates for shifts ending after midnight
    for i in range(len(df)):
        start_time = float(df.loc[i, "Start"])
        end_time = float(df.loc[i, "End"])
        start_date = df.loc[i, "Date"]
        end_date = "Error"

        if end_time == 24:
            end_time = 0.0

        if end_time >= start_time:
            end_date = start_date

        if end_time < start_time:
            end_date = start_date + timedelta(days=1)

        df.loc[i, "Start Date"] = start_date
        df.loc[i, "Start Date"] = df.loc[i, "Start Date"].strftime('%Y-%m-%d')
        df.loc[i, "End Date"] = end_date
        df.loc[i, "End Date"] = df.loc[i, "End Date"].strftime('%Y-%m-%d')

        # Format start and end times
        df.loc[i, "Start"] = format_time(start_time)
        df.loc[i, "End"] = format_time(end_time)

    # Combine dates and times
    df["Start DT"] = pd.to_datetime(df['Start Date'] + " " + df['Start'], format='%Y-%m-%d %H.%M')
    df["End DT"] = pd.to_datetime(df['End Date'] + " " + df['End'], format='%Y-%m-%d %H.%M')

    # Calculate shift lengths
    df["Shift Length"] = df["End DT"] - df["Start DT"]
    print(df["Shift Length"])

    return df


def show_only_shifts(df):
    df1 = pd.DataFrame()
    for i in range(len(df)):
        if df.loc[i, "Shift Length"] != 0:
            df1 = df1.append(df.loc[i])
    # df2 = df1.drop("Date", "Start", "End", "Start Date", "End Date")
    df2 = df1[["Start DT", "End DT", "Shift Length", "Comment"]]
    return df2


def calendar_format(df):
    df = show_only_shifts(df)
    df["Start Date"] = [d.date() for d in df["Start DT"]]
    df["End Date"] = [d.date() for d in df["End DT"]]
    df["Start Time"] = [d.time() for d in df["Start DT"]]
    df["End Time"] = [d.time() for d in df["End DT"]]
    df["All Day Event"] = "False"

    df = df.rename(columns={"Comment": "Subject"})
    df = df[["Start Date", "Start Time", "End Date", "End Time", "Subject", "All Day Event"]]
    return df


def extract_and_organise():
    """Runs all functions required to extract a rota document in the correct order"""

    df = calculate_shifts(
        name_shifts(
            sort_to_row(
                stack_days_vertically(
                    open_file(
                    )
                )
            )
        )
    )
    return df