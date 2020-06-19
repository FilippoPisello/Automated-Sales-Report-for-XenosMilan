#!/usr/bin/env python

# ------------------------------------------------------------------------------
# 0) DOWNLOAD THE REQUIRED PACKAGES
# ------------------------------------------------------------------------------
from webbot import Browser

import time

import os

import pandas as pd
import numpy as np

from datetime import date

from matplotlib import pyplot as plt
import matplotlib.ticker as ticker

import importlib

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_UNDERLINE

import win32com.client as client

# ------------------------------------------------------------------------------
# 1) SOURCE CSV DOWNLOAD
# ------------------------------------------------------------------------------
# Open the default browser and go to the website
web = Browser()
web.go_to("https://my.bigcartel.com")

# Provide login information and log in
web.type("#USERNAME",
         into="account_subdomain",
         id="account_subdomain")
web.type("#PASSWORD", into="password", id="password")
web.click("Log In")

# Navigate through the website
web.click("Orders")
web.click("Shipped")

# Downloadthereport
web.go_to("https://my.bigcartel.com/orders_exports.csv")
time.sleep(8)

# Close the browser

web.quit()

#NOTE##########################################
# Setting the directory into the resources folder
###############################################
resources_path = os.getcwd() + "/Resources"

os.chdir(resources_path)
# ------------------------------------------------------------------------------
# 2) INDENTIFY SOURCE CSV THROUGH THE PATH IN THE TXT FILE
# ------------------------------------------------------------------------------

with open('Report_path.txt', 'r') as file:
    data = file.read()
download_dir = data.split("\n")[0].replace('"', "").replace("\\", "/")

# ------------------------------------------------------------------------------
# 3) MOVING THE SOURCE CSV TO THE RESOURCES FOLDER
# ------------------------------------------------------------------------------
# Get the download directory and file name
filename = "orders.csv"

# Indentify location of origin and destination inside the directory
old_report_path = download_dir + "/" + filename
new_report_path = resources_path + "/" + filename

# Move the file
os.replace(old_report_path, new_report_path)

# ------------------------------------------------------------------------------
# 4) REPORT ANALYSIS
# ------------------------------------------------------------------------------
# 4.0) Getting the table ready
# Open the table
df = pd.read_csv(new_report_path)

# Matching columns to default names to be used in the code
name_code = "Code"
name_column_date = "Date"
name_column_full_name = "Name surname"
name_status = "Status"
name_payment_status = "Payment status"
name_shipping_status = "Shipping status"
name_shipping_country = "Shipping country"
name_shipping_city = "Shipping city"
name_items = "Items"
name_size = "Size"
name_item_count = "Items count"
name_total_price = "Raw price"
name_paid_price = "Paid price"
name_shipping = "Shipping price"
name_discount = "Discount"

df[name_column_full_name] = df["Buyer first name"] + " " + df["Buyer last name"]

# Estimate gender based on the name of the customer, creating a boolean column
dr_names = pd.read_csv("Names_list.csv")
df["First first name"] = df["Buyer first name"].str.split(pat=" ").str[0]


condition_male = df["First first name"].isin(list(dr_names["Male names"]))
condition_female = df["First first name"].isin(list(dr_names["Female names"]))

df["Male"] = np.where(condition_male, True,
                      np.where(condition_female, False, np.NaN)).astype("bool")

# Convert date into datetime format
df["Date"] = pd.to_datetime(df["Date"],
                            yearfirst=True,
                            utc=True)

# Add a column for a univoque monthcode
df["Month"] = pd.DatetimeIndex(df["Date"]).month
df["Month code"] = (pd.DatetimeIndex(df["Date"]).year.astype(str)
                    + "M"
                    + np.where(df["Month"].astype(str).str.len() == 1,
                               "0" + df["Month"].astype(str),
                               df["Month"].astype(str))
                    )
del(df["Month"])

# Extracting the size from the items columns
df[name_size] = df["Items"].str.extract(pat="(on_name:[A-Z]?[A-Z]?[A-Z])")
df[name_size] = df[name_size].str.slice(start=8)

# Exctracting the item name from the items column
df["Items"] = df["Items"].str.split(pat="|").str[0].str.split(pat=":").str[1]

# Extracting the type of item from the items column
df["Items type"] = df["Items"].str.split(pat=" ").str[-1].str.capitalize()

# Define net earnings for each item
df["Net earnings"] = df["Item total"] - df["Total discount"]

df.rename(
    {"Number": name_code,
        "Date": name_column_date,
        "Status": name_status,
        "Payment status": name_payment_status,
        "Shipping status": name_shipping_status,
        "Shipping country": name_shipping_country,
        "Shipping city": name_shipping_city,
        "Items": name_items,
        "Item count": name_item_count,
        "Item total": name_total_price,
        "Total price": name_paid_price,
        "Total shipping": name_shipping,
        "Total discount": name_discount
     }, axis=1, inplace=True)

# Rule out the columns currently considered useless
df = df[[
    "Code", "Name surname", "Male", "Month code", "Date", "Status",
    "Payment status", "Shipping status", "Shipping city", "Items",
    "Items type", "Size", "Items count", "Raw price", "Paid price",
    "Shipping price", "Discount", "Net earnings"
]]

# ------------------------------------------
# 4.1) Defining the functions to be used


def aggregate_by_date(original_dataframe,
                      time_serie_to_groupby,
                      serie_name="Period"):
    """Perform a groupby on a time serie with premade operations

    Parameters:
    - orginal_dataframe = name of the dataframe on which the aggregation is
    performed (ex: "df")
    - time_serie_to_groupby = name of the serie/column which is used as key
    for the aggregation, is intended to be a time serie (ex: "Month code")
    - serie_name = name to be assigned to the column containing the key of the
    aggregation after the groupby is performed (ex: "column 1")
    """

    aggregate = {"Items count": ["sum", "mean"], "Male": "mean",
                 "Raw price": ["sum", "mean"], "Net earnings": "sum"}

    df1 = original_dataframe.groupby([time_serie_to_groupby],
                                     as_index=False).agg(aggregate)

    df1.columns = df1.columns.droplevel(0)

    df1.columns = [serie_name, "Tot items ordered", "Avg items per order",
                   "Males percentage", "Tot raw price", "Avg raw price",
                   "Tot net earnings"]

    df1["Share of net earnings"] = ((df1["Tot net earnings"]
                                     / df1["Tot net earnings"].sum())).round(2)

    df1["Cumulative net earnings"] = df1["Tot net earnings"].cumsum().round(1)

    df1["Cumulative items sold"] = df1["Tot items ordered"].cumsum()

    df1 = df1.round({"Males percentage": 2,
                     "Tot raw price": 1,
                     "Avg raw price": 1,
                     "Tot net earnings": 1})
    return df1


def aggregate_by_category(original_dataframe,
                          serie_to_groupby,
                          serie_name="Values"):
    """Perform a groupby with premade operations

    Parameters:
    - orginal_dataframe = name of the dataframe on which the aggregation is
    performed (ex: "df")
    - serie_to_groupby = name of the serie/column which is used as key for the
    aggregation (ex: "column 1")
    - serie_name = name to be assigned to the column containing the key of the
    aggregation after the groupby is performed (ex: "column 1")
    """

    aggregate = {"Items count": "sum", "Male": "mean",
                 "Raw price": ["sum", "mean"], "Net earnings": "sum"}

    df1 = original_dataframe.groupby([serie_to_groupby],
                                     as_index=False).agg(aggregate)

    df1.columns = df1.columns.droplevel(0)
    df1.columns = [serie_name, "Tot items count", "Males percentage",
                   "Tot raw price", "Avg raw price", "Tot net earnings"]

    df1["Share net earnings"] = ((df1["Tot net earnings"]
                                  / df1["Tot net earnings"].sum())).round(2)

    df1 = df1.round({"Males percentage": 2,
                     "Tot raw price": 1,
                     "Avg raw price": 1,
                     "Tot net earnings": 1})
    return df1


def plot_vert_time_serie(dataframe, x1_serie, y1_serie, y2_serie, gender_serie,
                         title, x1_title, y1_title, y2_title,
                         label_x1="Month",
                         label_y1_males="Total number of items sold to males",
                         label_y1_females="Total number of items sold to females",
                         label_y2="Total net earnings",
                         width=0.3,
                         color_y1_males=(0, 0, 0),
                         color_y1_females=(0, 0, 0, 0.65),
                         color_y2="#666699",
                         save=False,
                         file_name="Image 1.jpg",
                         show=True):
    """Create a vertical plot of two series same x axis.

    The x axis is meant to represent a time serie. The serie y1 is designed to
    be a stacked bar based on a third serie "gender_serie" providing the gender
    proportion.
    -----------------------------------------------------------------------
    Parameters:
    - dataframe = name of the dataframe (ex: "df")
    - x1_serie, y1_serie, y2_serie = name of the chosen columns (ex: "Type")
    - gender_serie = name of the serie containing the proportion of males as
    a decimal
    - file_name = works only if save is True, name of the image saved in the
    work directory
    """

    # Create image
    fig = plt.figure(figsize=(8 * 1.5, 4.5 * 1.5), dpi=80)

    # Create the subplots and rename the axis
    plot1 = fig.add_subplot()
    plot1.set_title(title, fontsize=20)
    plot1.set_xlabel(x1_title, fontsize=15)
    plot1.set_ylabel(y1_title, fontsize=15)

    plot1b = plot1.twinx()
    plot1b.set_ylabel(y2_title, fontsize=15)

    # Get the series ready to be plotted
    x1_serie = dataframe[x1_serie]
    y1_serie = dataframe[y1_serie]
    y2_serie = dataframe[y2_serie]
    gender_serie = dataframe[gender_serie]

    x_serie = np.arange(len(x1_serie))

    items_males = y1_serie * gender_serie
    items_females = y1_serie * (1 - gender_serie)

    # Plot the series
    plot1.bar(x_serie - width / 2 - 0.01, items_males,
              width=width,
              color=color_y1_males,
              label=label_y1_males)
    plot1.bar(x_serie - width / 2 - 0.01, items_males + items_females,
              width=width,
              color=color_y1_females,
              label=label_y1_females)

    plot1b.bar(x_serie + width / 2 + 0.01, y2_serie,
               width=width,
               color=color_y2,
               label=label_y2)

    # Set axis properties
    number_x_ticks = max(1, len(x1_serie) // 8)
    plot1.set_xticks(np.arange(0, len(x1_serie), number_x_ticks))
    plot1.set_xticklabels(x1_serie)

    plot1.yaxis.set_major_locator(ticker.MaxNLocator(10))
    plot1b.yaxis.set_major_locator(ticker.MaxNLocator(10))

    formatter = ticker.FormatStrFormatter('€%1.0f')
    plot1b.yaxis.set_major_formatter(formatter)

    # Add further details
    fig.legend(loc="lower center", ncol=3, fontsize=9)
    plot1.grid(True, axis="both", ls="--")

    # Save
    if save:
        fig.savefig(file_name)

    # Show
    if not show:
        plt.close()


def plot_cumulative_time_serie(dataframe, x1_serie, y1_serie, y2_serie,
                               title, x1_title, y1_title, y2_title,
                               label_x1="Month",
                               label_y1="Cumulative number of items sold",
                               label_y2="Cumulative net earnings",
                               width=0.3,
                               color_y1=(0, 0, 0),
                               color_y2="#666699",
                               save=False,
                               file_name="Image 1.jpg",
                               show=True):
    """Create a line plot of two series with the same x axis.

    -----------------------------------------------------------------------
    Parameters:
    - dataframe = name of the dataframe (ex: "df")
    - x1_serie, y1_serie, y2_serie = name of the chosen columns (ex: "Type")
    - file_name = works only if save is True, name of the image saved in the
    work directory
    """

    # Create image
    fig = plt.figure(figsize=(8 * 1.5, 4.5 * 1.5), dpi=80)

    # Create the subplots and rename the axis
    plot1 = fig.add_subplot()
    plot1.set_title(title, fontsize=20)
    plot1.set_xlabel(x1_title, fontsize=15)
    plot1.set_ylabel(y1_title, fontsize=15)

    plot1b = plot1.twinx()
    plot1b.set_ylabel(y2_title, fontsize=15)

    # Get the series ready to be plotted
    x1_serie = dataframe[x1_serie]
    y1_serie = dataframe[y1_serie]
    y2_serie = dataframe[y2_serie]

    x_serie = np.arange(len(x1_serie))

    # Plot the series
    plot1.plot(x_serie, y1_serie,
               "--", marker="o",
               color=color_y1,
               label=label_y1)

    plot1b.plot(x_serie, y2_serie,
                "--", marker="o",
                color=color_y2,
                label=label_y2)

    # Set axis properties
    number_x_ticks = max(1, len(x1_serie) // 8)
    plot1.set_xticks(np.arange(0, len(x1_serie), number_x_ticks))
    plot1.set_xticklabels(x1_serie)

    plot1.yaxis.set_major_locator(ticker.MaxNLocator(8))
    plot1b.yaxis.set_major_locator(ticker.MaxNLocator(8))

    formatter = ticker.FormatStrFormatter('€%1.0f')
    plot1b.yaxis.set_major_formatter(formatter)

    # Add further details
    fig.legend(loc="lower center", ncol=3, fontsize=9)
    plot1.grid(True, axis="both", ls="--")

    # Save
    if save:
        fig.savefig(file_name)

    # Show
    if not show:
        plt.close()


def plot_horizontal_bar(dataframe, y1_serie, x1_serie, x2_serie, gender_serie,
                        title, y1_title, x1_title, x2_title,
                        label_x1_males="Total number of items sold to males",
                        label_x1_females="Total number of items sold to females",
                        label_x2="Total net earnings",
                        custom_width=0.4,
                        color_x1_males=(0, 0, 0),
                        color_x1_females=(0, 0, 0, 0.65),
                        color_x2="#666699",
                        save=False,
                        file_name="Image 1.jpg",
                        show=True):
    """Create an horizontal plot of two series same y axis.

    The y axis is meant to be categorical. The serie x1 is designed to be a
    stacked bar based on a third serie "gender_serie" providing the gender
    proportion.
    -----------------------------------------------------------------------
    Parameters:
    - dataframe = name of the dataframe (ex: "df")
    - y1_serie, x1_serie, x2_serie = name of the chosen columns (ex: "Type")
    - file_name = works only if save is True, name of the image saved in the
    work directory
    """

    # Create image
    fig = plt.figure(figsize=(8 * 1.5, 4.5 * 1.5), dpi=80)

    # Create the subplots and rename the axis
    plot1 = fig.add_subplot()
    plot1.set_title(title, fontsize=20)
    plot1.set_ylabel(y1_title, fontsize=15)
    plot1.set_xlabel(x1_title, fontsize=15)

    plot1b = plot1.twiny()
    plot1b.set_xlabel(x2_title, fontsize=15)

    # Get the series ready to be plotted
    y1_serie = dataframe[y1_serie]
    x1_serie = dataframe[x1_serie]
    x2_serie = dataframe[x2_serie]
    gender_serie = dataframe[gender_serie]

    width = custom_width
    y_serie = np.arange(len(y1_serie))
    items_males = x1_serie * gender_serie
    items_females = x1_serie * (1 - gender_serie)

    # Plot the series
    plot1.barh(y_serie - width / 2 - 0.01, items_males,
               height=width,
               color=color_x1_males,
               label=label_x1_males)
    plot1.barh(y_serie - width / 2 - 0.01, items_males + items_females,
               height=width,
               color=color_x1_females,
               label=label_x1_females)
    plot1b.barh(y_serie + width / 2 + 0.01, x2_serie,
                height=width,
                color=color_x2,
                label=label_x2)

    # Set axis properties
    plot1.set_yticks(np.arange(0, len(y1_serie), 1))
    plot1.set_yticklabels(y1_serie)
    plot1.set_xticks(np.arange(0, x1_serie.max() + 1, 1))

    formatter = ticker.FormatStrFormatter('€%1.0f')
    plot1b.xaxis.set_major_formatter(formatter)

    # Add further details
    fig.legend(loc="lower center", ncol=3, fontsize=9)

    # Save
    if save:
        fig.savefig(file_name)
    # Show
    if not show:
        plt.close()


def convert_to_pdf(filepath: str):
    """Save a pdf of a docx file.

    Requires [import win32com.client as client]
    """
    try:
        word = client.DispatchEx("Word.Application")
        target_path = filepath.replace(".docx", r".pdf")
        word_doc = word.Documents.Open(filepath)
        word_doc.SaveAs(target_path, FileFormat=17)
        word_doc.Close()
    except Exception as e:
        raise e
    finally:
        word.Quit()

# ------------------------------------------
# 4.2) Creating report for total values
# Generate descriptive indicators for volume of sales and revenues


values_tot = {}
values_mean = {}
values_min = {}
values_max = {}

columns_of_interest = ["Items count", "Raw price", "Net earnings"]

for column in columns_of_interest:
    values_tot[column] = round(df[column].sum(), 0)
    values_mean[column] = round(df[column].mean(), 0)
    values_min[column] = round(df[column].min(), 0)
    values_max[column] = round(df[column].max(), 0)

df1 = aggregate_by_date(df, "Month code", "Month")

plot_vert_time_serie(df1, "Month", "Tot items ordered",
                     "Tot net earnings",
                     "Males percentage",
                     title="Items sold and revenues over time",
                     x1_title="Months",
                     y1_title="Item counts",
                     y2_title="Net earnings",
                     save=True,
                     file_name="Items sold and revenues_time.png",
                     show=False)

plot_cumulative_time_serie(df1, "Month", "Cumulative items sold",
                           "Cumulative net earnings",
                           "Cumulative sales and earnings", "Month",
                           "Items sold", "Net earnings",
                           save=True,
                           file_name="Cumulative sales and earnings_time.png",
                           show=False)

df2 = aggregate_by_category(df, "Items type", "Items type")

plot_horizontal_bar(dataframe=df2,
                    y1_serie="Items type",
                    x1_serie="Tot items count",
                    x2_serie="Tot net earnings",
                    gender_serie="Males percentage",
                    title="Indicators for item type",
                    y1_title="Item type",
                    x1_title="Item count",
                    x2_title="Net earnings",
                    save=True, file_name="Indicators for item type.png",
                    show=False)

df3 = aggregate_by_category(df, "Items", "Items")

plot_horizontal_bar(dataframe=df3,
                    y1_serie="Items",
                    x1_serie="Tot items count",
                    x2_serie="Tot net earnings",
                    gender_serie="Males percentage",
                    title="Indicators for item",
                    y1_title="Item",
                    x1_title="Item count",
                    x2_title="Net earnings",
                    save=True, file_name="Indicators for item.png",
                    show=False)

# ------------------------------------------------------------------------------
# 5) Document setting and customization
# ------------------------------------------------------------------------------
document = Document("Xenos report template.docx")

# Footer
section = document.sections[0]
footer = section.footer
paragraph = footer.paragraphs[0]
today = date.today().strftime("%d/%m/%Y")
run = paragraph.add_run(f"\tConfidenziale - Xenos Milan\t{today}")
run.italic = True

# Define the various styles in descending hierarchical order
my_styles = document.styles

# Title
title_style = my_styles.add_style("Custom title", WD_STYLE_TYPE.PARAGRAPH)
title_style.base_style = my_styles['Normal']
format_title_style = title_style.paragraph_format
title_style.hidden = False
title_style.quick_style = True
title_style.priority = 2

# Heading 1
heading_1_style = my_styles.add_style("Custom heading 1",
                                      WD_STYLE_TYPE.PARAGRAPH)
heading_1_style.base_style = my_styles['Normal']
format_heading_1_style = heading_1_style.paragraph_format
heading_1_style.hidden = False
heading_1_style.quick_style = True
heading_1_style.priority = 3

# Heading 2
heading_2_style = my_styles.add_style("Custom heading 2",
                                      WD_STYLE_TYPE.PARAGRAPH)
heading_2_style.base_style = my_styles['Normal']
format_heading_2_style = heading_2_style.paragraph_format
heading_2_style.hidden = False
heading_2_style.quick_style = True
heading_2_style.priority = 4

# Heading 3
heading_3_style = my_styles.add_style("Custom heading 3",
                                      WD_STYLE_TYPE.PARAGRAPH)
heading_3_style.base_style = my_styles['Normal']
format_heading_3_style = heading_3_style.paragraph_format
heading_3_style.hidden = False
heading_3_style.quick_style = True
heading_3_style.priority = 5

# Body
body_style = my_styles.add_style("Custom body",
                                 WD_STYLE_TYPE.PARAGRAPH)
body_style.base_style = my_styles['Normal']
format_body_style = body_style.paragraph_format
body_style.hidden = False
body_style.quick_style = True
body_style.priority = 1

# Customization
# Customize title
format_title_style.alignment = WD_ALIGN_PARAGRAPH.CENTER
format_title_style.space_before = Pt(3.0)
format_title_style.space_after = Pt(12.0)
format_title_style.line_spacing_rule = WD_LINE_SPACING.DOUBLE

title_style.font.name = "Arial Black"
title_style.font.bold = True
title_style.font.underline = WD_UNDERLINE.THICK
title_style.font.size = Pt(20)
#title_style.font.color = RGBColor(47, 84, 150)

# Customize heading 1
format_heading_1_style.alignment = WD_ALIGN_PARAGRAPH.LEFT
format_heading_1_style.space_before = Pt(12.0)
format_heading_1_style.space_after = Pt(12.0)
format_heading_1_style.line_spacing_rule = WD_LINE_SPACING.SINGLE
#format_heading_1_style.left_indent = Cm(-1.0)
#format_heading_1_style.right_indent = Cm(-1.0)

heading_1_style.font.name = "Arial"
heading_1_style.font.italic = True
heading_1_style.font.bold = True
heading_1_style.font.underline = True
heading_1_style.font.size = Pt(16)
#heading_1_style.font.color = RGBColor(47, 84, 150)

# Customize heading 2
format_heading_2_style.alignment = WD_ALIGN_PARAGRAPH.LEFT
format_heading_2_style.space_before = Pt(8.0)
format_heading_2_style.space_after = Pt(6.0)
format_heading_2_style.line_spacing_rule = WD_LINE_SPACING.SINGLE
#format_heading_2_style.left_indent = Cm(-1.0)
#format_heading_2_style.right_indent = Cm(-1.0)

heading_2_style.font.name = "Arial"
heading_2_style.font.italic = True
heading_2_style.font.bold = True
heading_2_style.font.underline = True
heading_2_style.font.size = Pt(13)
#heading_2_style.font.color = RGBColor(47, 84, 150)

# Customize heading 3
format_heading_3_style.alignment = WD_ALIGN_PARAGRAPH.LEFT
format_heading_3_style.space_before = Pt(6.0)
format_heading_3_style.space_after = Pt(6.0)
format_heading_3_style.line_spacing_rule = WD_LINE_SPACING.SINGLE
#format_heading_3_style.left_indent = Cm(-1.0)
#format_heading_3_style.right_indent = Cm(-1.0)

heading_3_style.font.name = "Arial"
heading_3_style.font.italic = True
heading_3_style.font.bold = True
heading_3_style.font.size = Pt(11)
#heading_3_style.font.color = RGBColor(47, 84, 150)


# Customize body
format_body_style.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_MED
format_body_style.space_before = Pt(3.0)
format_body_style.space_after = Pt(3.0)
format_body_style.line_spacing_rule = WD_LINE_SPACING.SINGLE
#format_body_style.left_indent = Cm(-1.0)
#format_body_style.right_indent = Cm(-1.0)

body_style.font.name = "Arial"
body_style.font.size = Pt(11)

filename = "Xenos report.docx"
# document.save(filename)

# ------------------------------------------------------------------------------
# 6) Writing in the document
# ------------------------------------------------------------------------------

#filename = "Xenos report.docx"
#document = Document(filename)

title = document.add_paragraph(str.upper("Report Xenos"))
title.style = document.styles["Custom title"]

paragraph2 = document.add_paragraph("Dati storici")
paragraph2.style = document.styles["Custom heading 1"]

text = ("Questa sezione contiene i dati calcolati su tutto il periodo di attività.")
paragraph3 = document.add_paragraph(text)
paragraph3.style = document.styles["Custom body"]

paragraph4 = document.add_paragraph("Overview")
paragraph4.style = document.styles["Custom heading 2"]

text1 = f"Numero totale di capi venduti: {values_tot['Items count']}"
paragraph5 = document.add_paragraph(text1)
paragraph5.style = document.styles["Custom body"]

text2 = f"Numero medio di capi venduti per ordine: {values_mean['Items count']}"
paragraph6 = document.add_paragraph(text2)
paragraph6.style = document.styles["Custom body"]

text3 = f"Numero massimo di capi venduti in un ordine: {values_max['Items count']}"
paragraph7 = document.add_paragraph(text3)
paragraph7.style = document.styles["Custom body"]

text4 = f"Valore complessivo dei capi acquistati: {values_tot['Raw price']}"
paragraph6 = document.add_paragraph(text4)
paragraph6.style = document.styles["Custom body"]

text5 = f"Valore medio dei capi acquistati per ordine: {values_mean['Raw price']}"
paragraph8 = document.add_paragraph(text5)
paragraph8.style = document.styles["Custom body"]

text6 = f"Valore massimo dei capi acquistati in un ordine: {values_mean['Raw price']}"
paragraph9 = document.add_paragraph(text6)
paragraph9.style = document.styles["Custom body"]

text7 = f"Totale ricavi al netto dei costi di vendita: {values_tot['Net earnings']}"
paragraph10 = document.add_paragraph(text7)
paragraph10.style = document.styles["Custom body"]

text8 = f"Media ricavi al netto dei costi di vendita per ordine: {values_mean['Net earnings']}"
paragraph11 = document.add_paragraph(text8)
paragraph11.style = document.styles["Custom body"]

text9 = f"Massimo ricavi al netto dei costi di vendita in un ordine: {values_max['Net earnings']}"
paragraph12 = document.add_paragraph(text9)
paragraph12.style = document.styles["Custom body"]

paragraph13 = document.add_paragraph("Dati nel tempo")
paragraph13.style = document.styles["Custom heading 2"]

text = "Il seguente grafico mostra l'andamento del numero di vendite e dei ricavi al netto dei costi di vendita nel corso del tempo. L'unità di misura utilizzata è il mese."
paragraph14 = document.add_paragraph(text)
paragraph14.style = document.styles["Custom body"]
document.add_picture("Items sold and revenues_time.png", width=Cm(15.0))

text = "Il grafico successivo mostra invece l'andamento del tempo della quantità cumulativa di numeri di capi venduti e ricavi al netto dei costi di vendita. La quantità cumulativa è ottenuta sommando il valore del periodo alla somma di tutti i valori precedenti."
paragraph15 = document.add_paragraph(text)
paragraph15.style = document.styles["Custom body"]
document.add_picture("Cumulative sales and earnings_time.png",
                     width=Cm(15.0))

paragraph16 = document.add_paragraph("Dati per tipologia di capo")
paragraph16.style = document.styles["Custom heading 2"]

text = "Il grafico che segue mostra il numero totale di ordini per tipologia di capo, divisa in due colorazioni a seconda del genere. Insieme a questo, sul secondo asse, sono indicati i ricavi totali al netto dei costi di vendita per tipologia di prodotto."
paragraph17 = document.add_paragraph(text)
paragraph17.style = document.styles["Custom body"]
document.add_picture("Indicators for item type.png",
                     width=Cm(15.0))

paragraph17 = document.add_paragraph("Dati per capo")
paragraph17.style = document.styles["Custom heading 2"]

text = "Il grafico è del tutto simile al precedente, con i singoli capi al posto delle tipologie di capo."
paragraph18 = document.add_paragraph(text)
paragraph18.style = document.styles["Custom body"]
document.add_picture("Indicators for item.png",
                     width=Cm(15.0))

document.save(filename)

# ------------------------------------------------------------------------------
# 7) Transform the document into a pdf
# ------------------------------------------------------------------------------

path = os.getcwd() + "/" + filename
convert_to_pdf(path)

# ------------------------------------------------------------------------------
# 8) Move the document to the main folder
# ------------------------------------------------------------------------------

pdf_filename = filename.replace(".docx", ".pdf")
old_pdf_path = os.getcwd() + "/" + pdf_filename
new_pdf_path = os.getcwd().replace("Resources", pdf_filename)
os.replace(old_pdf_path, new_pdf_path)

print("Process completed")
input("Press any key to exit")
